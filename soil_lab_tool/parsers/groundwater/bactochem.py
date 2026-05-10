"""
parsers/groundwater/bactochem.py
---------------------------------
Parser for בקטוכם (Bactochem) groundwater monitoring reports.

File format (CSV or XLSX):
  - Single header row (row 0) with Hebrew column names
  - Column 'רכיב'          : compound / parameter name (English)
  - Column 'תוצאה'         : result value or 'Not Detected'
  - Column 'תיאור דוגמה'   : sample location / borehole ID
  - Column 'תאריך דיגום'   : sampling datetime
  - Column 'אנליזה'        : analysis code (may or may not be present)

Analysis type mapping is done by compound name (not analysis code):
  GW_VOC         : Benzene, Toluene, Ethyl Benzene, Xylene, MTBE, Naphthalene, TBA
  GW_FIELD_PARAMS: pH, EC, Temperature, DO, Turbidity, Redox, depth params

GW thresholds sourced from soil_vsl_tier1_v7_2024.xlsx  "Groundwater" column:
  Benzene      1   mg/L
  Toluene    600   mg/L
  Ethylbenzene 700 mg/L
  Xylene     500   mg/L
  MTBE       240   mg/L
"""

from __future__ import annotations

import io
import os
import re

import pandas as pd

from parsers.base import BaseParser
from core.lab_value_parser import LabValueParser
from core.cas_lookup import name_to_cas


# ── BTEX extraction helpers ───────────────────────────────────────────────────

# CAS registry number → preferred English display name
_CAS_TO_NAME: dict[str, str] = {
    "71-43-2":   "Benzene",
    "108-88-3":  "Toluene",
    "100-41-4":  "Ethylbenzene",
    "1330-20-7": "Xylene",
    "1634-04-4": "MTBE",
    "91-20-3":   "Naphthalene",
    "75-65-0":   "TBA",
}

# Fallback: recognise compound name from the tail of a CAS line
_BTEX_NAME_RE = re.compile(
    r"\b(Benzene|Toluene|Ethyl\s*Benzene|Ethylbenzene|Xylenes?|MTBE|"
    r"Methyl\s+Tert-Butyl\s+Ether|Naphthalene|TBA|Tert-Butyl\s+Alcohol)\b",
    re.I,
)

# Reversed-RTL sample-header pattern produced by pdfplumber on Hebrew PDFs.
# Logical: "מספר הדוגמה: 1984651"  →  visual: "1984651 :המגודה רפסמ …"
_BC_SAMPLE_HDR_RE = re.compile(r"(\d{5,})\s+:המגודה רפסמ")

# Full CAS-line pattern (mg/L only) — matches lines like:
#   CAS #: 71-43-2  0.001  mg/L  Not Detected  Benzene
#   CAS #: 1634-04-4  0.001  mg/L  3.200  MTBE
_GW_CAS_LINE_RE = re.compile(
    r"CAS\s*#:\s*(?P<cas>[\w.\-]+)"        # CAS number
    r"(?:\s+(?P<loq>[\d.]+))?"             # optional LOQ
    r"\s+mg/L"                             # unit (GW only)
    r"(?:\s+X[≤≥<>≠]\s*[\d.]+)?"          # optional threshold (discarded)
    r"\s+(?P<result>Not\s+Detected|<[\d.]+|[\d.]+(?:[Ee][+\-]?\d+)?)"
    r"(?:\s+\d+/)?"                        # optional sample ref (discarded)
    r"\s*(?P<compound>.*)?$",
    re.IGNORECASE,
)


# ── PDF field-params helpers ──────────────────────────────────────────────────

_ROW_TOL   = 6
_HEBREW_RE = re.compile(r"[א-ת]")

# Raw pdfplumber-visual tokens that identify a field-parameter row.
# pdfplumber extracts Hebrew RTL text in visual order (characters reversed),
# so logical "מומס" arrives as visual "סמומ", "מוליכות" → "תוכילומ", etc.
_BC_FP_TOKENS: frozenset[str] = frozenset({
    "סמומ",      # מומס  (dissolved)     — DO
    "ןצמח",      # חמצן  (oxygen)        — DO
    "תוכילומ",   # מוליכות (conductivity)
    "הרוטרפמט",  # טמפרטורה (temperature)
    "סקודר",     # רדוקס (redox)
    "תוריכע",    # עכירות (turbidity)
    "קמוע",      # עומק  (depth)
    "הביאש",     # שאיבה (pumping)
    "LOWFLOW",
    "pH",
})


def _fix_rtl_bc(text: str) -> str:
    """Reverse visual Hebrew (pdfplumber) back to logical reading order."""
    if not _HEBREW_RE.search(text):
        return text
    tokens = text.split()
    tokens.reverse()
    return " ".join(t[::-1] if _HEBREW_RE.search(t) else t for t in tokens)


def _cluster_rows_bc(words: list[dict]) -> list[list[dict]]:
    """Cluster pdfplumber word dicts into rows by y-position."""
    groups: list[list[dict]] = []
    for w in sorted(words, key=lambda x: (x["top"], x["x0"])):
        for g in groups:
            if abs(w["top"] - g[0]["top"]) <= _ROW_TOL:
                g.append(w)
                break
        else:
            groups.append([w])
    for g in groups:
        g.sort(key=lambda x: x["x0"])
    return groups


def _canonical_bc_fp(name_fixed: str) -> str | None:
    """Map a reconstructed Hebrew name → canonical GW_FIELD_PARAMS compound.
    Returns None if the name is not a recognised field parameter."""
    n = name_fixed.lower()
    if "מוליכות" in n:
        return "מוליכות"
    if "טמפרטורה" in n:
        return "טמפרטורה"
    if "רדוקס" in n:
        return "רדוקס"
    if "עכירות" in n:
        return "עכירות"
    if "חמצן" in n:
        return "חמצן מומס DO"
    if "ph" in n:
        return "הגבה pH"
    if "עומק" in n:
        if "שאיבה" in n:
            return "עומק שאיבה"
        if "פני" in n or "מים" in n:
            return "עומק פני המים"
        if "כללי" in n:
            return "עומק כללי קידוח"
        if "דיגום" in n or "lowflow" in n:
            return "עומק דיגום"
        if "עליון" in n or "מפלס" in n:
            return "מפלס עליון"
        if "קידוח" in n:
            return "עומק קידוח"
        return "עומק"
    if "lowflow" in n:
        return "עומק דיגום LOWFLOW"
    return None


def _parse_bc_fp_row(group: list[dict], sample_id: str, vp) -> dict | None:
    """Parse one word-row into a GW_FIELD_PARAMS record, or return None.

    Row layout (left→right in PDF x-coords, RTL document):
        unit | value | name_tokens (reversed Hebrew)

    Steps:
      1. Row must contain at least one _BC_FP_TOKENS keyword.
      2. Numeric token (^-?[\\d.]+$) = value.
      3. Leftmost token (index 0) = unit; if tokens 0-1 are 'pH units', use both.
      4. All tokens after value = name → reconstruct with _fix_rtl_bc → classify.
    """
    texts = [w["text"] for w in group]

    if not any(t in _BC_FP_TOKENS for t in texts):
        return None

    val_idx = next(
        (i for i, t in enumerate(texts) if re.match(r"^-?[\d.]+$", t)),
        None,
    )
    if val_idx is None:
        return None
    val_str = texts[val_idx]

    if val_idx >= 2 and texts[0].lower() == "ph" and texts[1].lower() == "units":
        unit_str = "pH units"
    elif val_idx >= 1:
        unit_str = texts[0]
    else:
        unit_str = ""

    name_tokens = texts[val_idx + 1:]
    if not name_tokens:
        return None
    name_fixed = _fix_rtl_bc(" ".join(name_tokens))

    compound = _canonical_bc_fp(name_fixed)
    if compound is None:
        return None

    value, flag = vp.parse(val_str)
    return {
        "lab":           "בקטוכם",
        "sample_id":     sample_id,
        "compound":      compound,
        "cas":           "",
        "value":         value,
        "flag":          flag,
        "unit":          unit_str,
        "lod":           None,
        "analysis_type": "GW_FIELD_PARAMS",
    }


# ── English compound names → CAS (same as KTE groundwater)
GW_CAS: dict[str, str] = {
    "benzene":                    "71-43-2",
    "toluene":                    "108-88-3",
    "ethyl benzene":              "100-41-4",
    "ethylbenzene":               "100-41-4",
    "xylene":                     "1330-20-7",
    "xylenes":                    "1330-20-7",
    "mtbe":                       "1634-04-4",
    "methyl tert-butyl ether":    "1634-04-4",
    "naphthalene":                "91-20-3",
    "tba":                        "75-65-0",   # tert-Butyl Alcohol
    "tert-butanol":               "75-65-0",
    "tert-butyl alcohol":         "75-65-0",
}

# Compound name keywords that indicate LOWFLOW (field parameters)
_LOWFLOW_KEYWORDS = (
    "ph",
    "conductivity",
    "temp",
    "dissolved o",
    "turbidity",
    "redox",
    "depth",
    "drilling",
    "sampling depth",
    "upper level",
    "water level",
)

# Compound names that are clearly GW_VOC analytes
_VOC_KEYWORDS = (
    "benzene", "toluene", "xylene", "mtbe", "naphthalene",
    "ethyl", "ethylbenzene", "tba", "tert-butyl",
)


def _resolve_cas(compound: str) -> str:
    key = compound.strip().lower()
    return GW_CAS.get(key) or name_to_cas(compound) or ""


def _classify_compound(name: str) -> str | None:
    """Return 'GW_VOC', 'GW_FIELD_PARAMS', or None (skip row)."""
    low = name.strip().lower()
    if any(k in low for k in _VOC_KEYWORDS):
        return "GW_VOC"
    if any(k in low for k in _LOWFLOW_KEYWORDS):
        return "GW_FIELD_PARAMS"
    return None


class BactochemGroundwaterParser(BaseParser):
    """
    Parses Bactochem groundwater lab reports (CSV, XLSX, or PDF).

    Bactochem files use a **single** Hebrew header row (unlike KTE which has
    two). Compound names are English. Analysis type is inferred from the
    compound name rather than an analysis-code column.

    PDF input → GW_FIELD_PARAMS records (field parameters outside BTEX table).
    CSV/XLSX input → GW_VOC and LOWFLOW records (tabular BTEX data).
    """

    LAB_NAME = "בקטוכם"
    ANALYSIS_TYPES = ["GW_VOC", "GW_FIELD_PARAMS"]

    # Named columns used by Bactochem
    COL_COMPOUND = "רכיב"
    COL_RESULT   = "תוצאה"
    COL_LOCATION = "תיאור דוגמה"
    COL_DATE     = "תאריך דיגום"

    def __init__(self, debug: bool | None = None):
        self._vp    = LabValueParser()
        self._debug = debug if debug is not None else bool(os.environ.get("BACTOCHEM_DEBUG"))

    # ------------------------------------------------------------------
    def parse(self, file_obj: io.BytesIO | str) -> list[dict]:
        # ── Path detection ────────────────────────────────────────────
        if isinstance(file_obj, io.BytesIO):
            file_obj.seek(0)
            header = file_obj.read(4)
            file_obj.seek(0)
            if self._debug:
                print(f"[BC DEBUG] Input: BytesIO  header={header!r}  "
                      f"is_pdf={header == b'%PDF'}")
            if header == b"%PDF":
                if self._debug:
                    print("[BC DEBUG] → PDF path")
                return self._parse_pdf(file_obj)
            if self._debug:
                print("[BC DEBUG] → CSV/XLSX path")
        elif isinstance(file_obj, str):
            if self._debug:
                print(f"[BC DEBUG] Input: str path={file_obj!r}  "
                      f"ends_pdf={file_obj.lower().endswith('.pdf')}")
            if file_obj.lower().endswith(".pdf"):
                if self._debug:
                    print("[BC DEBUG] → PDF path (str)")
                with open(file_obj, "rb") as fh:
                    return self._parse_pdf(io.BytesIO(fh.read()))
            if self._debug:
                print("[BC DEBUG] → CSV/XLSX path (str)")
        else:
            if self._debug:
                print(f"[BC DEBUG] Input: unexpected type {type(file_obj).__name__}")

        # ── Tabular path (CSV / XLSX) ─────────────────────────────────
        df = self._read(file_obj)
        if df is None or df.empty:
            if self._debug:
                print("[BC DEBUG] DataFrame is empty — 0 records")
            return []

        if self._debug:
            print(f"[BC DEBUG] DataFrame shape: {df.shape}")
            print(f"[BC DEBUG] Columns: {list(df.columns)}")
            for col in (self.COL_COMPOUND, self.COL_RESULT,
                        self.COL_LOCATION, self.COL_DATE):
                present = col in df.columns
                print(f"[BC DEBUG]   col {col!r}: {'FOUND' if present else 'MISSING'}")

        records: list[dict] = []
        skipped_unknown = []
        for _, row in df.iterrows():
            compound = str(row.get(self.COL_COMPOUND, "")).strip()
            raw_val  = str(row.get(self.COL_RESULT,   "")).strip()
            loc      = str(row.get(self.COL_LOCATION, "")).strip()
            date_val = str(row.get(self.COL_DATE,     "")).strip()

            if not compound or compound.lower() in ("nan", ""):
                continue

            atype = _classify_compound(compound)
            if atype is None:
                skipped_unknown.append(compound)
                continue

            if raw_val.lower() in ("not detected", "nd", "n.d.", "n/d",
                                   "<dl", "none", ""):
                value, flag = None, "ND"
            else:
                value, flag = self._vp.parse(raw_val)

            cas = _resolve_cas(compound) if atype == "GW_VOC" else ""

            sample_id = (loc if loc and loc.lower() not in ("nan", "")
                         else "Sample")
            date_str = self._short_date(date_val)
            if date_str:
                sample_id = f"{sample_id} ({date_str})"

            records.append({
                "lab":           self.LAB_NAME,
                "sample_id":     sample_id,
                "compound":      compound,
                "cas":           cas,
                "value":         value,
                "flag":          flag,
                "unit":          "mg/L" if atype == "GW_VOC" else "",
                "lod":           None,
                "analysis_type": atype,
            })

        if self._debug:
            print(f"\n[BC DEBUG] Tabular parse complete: {len(records)} records")
            if skipped_unknown:
                print(f"[BC DEBUG] Skipped (unclassified): {skipped_unknown}")
            from collections import Counter
            for atype, cnt in Counter(r["analysis_type"] for r in records).most_common():
                print(f"[BC DEBUG]   {atype}: {cnt}")

        return records

    # ------------------------------------------------------------------
    def _parse_pdf(self, file_obj: io.BytesIO) -> list[dict]:
        """Extract GW_FIELD_PARAMS records from a Bactochem PDF."""
        try:
            import pdfplumber
        except ImportError as exc:
            raise ImportError(
                "pdfplumber is required for PDF parsing: pip install pdfplumber"
            ) from exc

        records: list[dict] = []
        sample_id = "Sample"

        with pdfplumber.open(file_obj) as pdf:
            if self._debug:
                print(f"[BC DEBUG] PDF: {len(pdf.pages)} page(s)  "
                      f"size={pdf.pages[0].width:.0f}×{pdf.pages[0].height:.0f}")

            for page_num, page in enumerate(pdf.pages):
                if page_num == 0:
                    sample_id = self._extract_sample_id_bc(page)
                    if self._debug:
                        print(f"[BC DEBUG] sample_id={sample_id!r}")

                words = page.extract_words(
                    x_tolerance=4, y_tolerance=4,
                    keep_blank_chars=False, use_text_flow=False,
                ) or []

                if self._debug:
                    print(f"\n[BC DEBUG] Page {page_num}: {len(words)} words")
                    print(f"[BC DEBUG] First 20 words:")
                    for w in words[:20]:
                        print(f"  x0={w['x0']:6.1f} top={w['top']:6.1f} "
                              f"text={w['text']!r}")

                if not words:
                    continue

                row_groups = _cluster_rows_bc(words)
                if self._debug:
                    print(f"[BC DEBUG] {len(row_groups)} row groups on page {page_num}")

                page_hits = 0
                for gi, group in enumerate(row_groups):
                    texts = [w["text"] for w in group]
                    rec = _parse_bc_fp_row(group, sample_id, self._vp)
                    if rec is not None:
                        records.append(rec)
                        page_hits += 1
                        if self._debug:
                            print(f"  row {gi:3d} HIT  texts={texts}  "
                                  f"→ compound={rec['compound']!r}  "
                                  f"value={rec['value']}  unit={rec['unit']!r}")
                    elif self._debug and any(t in _BC_FP_TOKENS for t in texts):
                        print(f"  row {gi:3d} TOKEN-MATCH-NO-VALUE  texts={texts}")

                if self._debug:
                    print(f"[BC DEBUG] Page {page_num}: {page_hits} field-param records")

                # ── BTEX table extraction ──────────────────────────────────
                if self._debug:
                    print(f"[BC DEBUG] Page {page_num}: scanning tables for BTEX…")
                btex_recs = self._extract_btex(page, sample_id)
                records.extend(btex_recs)
                if self._debug:
                    print(f"[BC DEBUG] Page {page_num}: {len(btex_recs)} BTEX (GW_VOC) records")

        if self._debug:
            print(f"\n[BC DEBUG] PDF parse complete: {len(records)} total records")
            from collections import Counter
            for atype, cnt in Counter(r["analysis_type"] for r in records).most_common():
                print(f"[BC DEBUG]   {atype}: {cnt}")

        return records

    def _extract_btex(self, page, sample_id: str) -> list[dict]:
        """Extract GW_VOC records from BTEX data lines in a PDF page.

        The BTEX compounds appear as free-form text lines (not in proper PDF
        table cells), so page.extract_text() is used.  Each line matching
        the pattern "CAS #: {cas} {loq} mg/L {result} {compound}" becomes
        one GW_VOC record.
        """
        records: list[dict] = []
        text = page.extract_text() or ""

        for line in text.splitlines():
            m = _GW_CAS_LINE_RE.search(line)
            if not m:
                continue

            cas        = m.group("cas").strip()
            result_raw = m.group("result").strip()
            name_tail  = (m.group("compound") or "").strip()

            compound = _CAS_TO_NAME.get(cas)
            if compound is None:
                nm = _BTEX_NAME_RE.search(name_tail)
                compound = nm.group(1) if nm else (name_tail or None)
            if not compound:
                if self._debug:
                    print(f"  [BTEX] unrecognised CAS {cas!r}: {line!r}")
                continue

            value, flag = self._vp.parse(result_raw)

            if self._debug:
                print(f"  [BTEX] compound={compound!r}  cas={cas!r}  "
                      f"result={result_raw!r}  →  value={value}  flag={flag!r}")

            records.append({
                "lab":           self.LAB_NAME,
                "sample_id":     sample_id,
                "compound":      compound,
                "cas":           cas,
                "value":         value,
                "flag":          flag,
                "unit":          "mg/L",
                "lod":           None,
                "analysis_type": "GW_VOC",
            })

        return records

    def _extract_sample_id_bc(self, page) -> str:
        """Extract a borehole/well ID from the first page of a Bactochem PDF."""
        text = page.extract_text() or ""
        if self._debug:
            print("[BC DEBUG] First 25 lines of page 0 text:")
            for i, ln in enumerate(text.splitlines()[:25]):
                print(f"  {i:2d}: {ln!r}")
        for line in text.splitlines()[:25]:
            m = re.search(r'מספר הדוגמה[:\s]+(\d+)', line)
            if m:
                if self._debug:
                    print(f"[BC DEBUG] sample_id matched 'מספר הדוגמה': {m.group(1)!r}")
                return m.group(1)
            m = re.search(
                r"\b(?:GW|BH|PZ|MW|OBS|MON|BOR|קידוח|באר)[-_]?\d+\b",
                line, re.I,
            )
            if m:
                if self._debug:
                    print(f"[BC DEBUG] sample_id matched well-ID pattern: {m.group(0)!r}")
                return m.group(0).upper()
            if re.search(r"מספר\s*(קידוח|דגימה|באר|פיזומטר)", line):
                m2 = re.search(r"[:]\s*(\S+)\s*$", line) or re.search(r"(\S+)$", line)
                if m2:
                    if self._debug:
                        print(f"[BC DEBUG] sample_id matched Hebrew label: {m2.group(1)!r}")
                    return m2.group(1)
        if self._debug:
            print("[BC DEBUG] sample_id: no pattern matched, using 'Sample'")
        return "Sample"

    # ------------------------------------------------------------------
    def _read(self, file_obj: io.BytesIO | str) -> pd.DataFrame | None:
        """
        Read a Bactochem file.  The file has a single Hebrew header row
        (row 0 = column names; row 1 = first data row).
        """
        try:
            if isinstance(file_obj, str) and file_obj.lower().endswith(".csv"):
                df = pd.read_csv(
                    file_obj, encoding="utf-8-sig", dtype=str,
                    engine="python", usecols=list(range(21)),
                ).fillna("")
            else:
                xl = pd.ExcelFile(file_obj)
                raw = xl.parse(xl.sheet_names[0], header=None,
                               dtype=str).fillna("")
                # Auto-detect: if row 0 is a Hebrew header (not a data row),
                # use it as column names; otherwise fall back to positional access.
                first_cell = str(raw.iloc[0, 0]).strip()
                if not first_cell.replace("-", "").isdigit():
                    df = raw.iloc[1:].reset_index(drop=True)
                    df.columns = [str(c).strip() for c in raw.iloc[0].values]
                else:
                    df = raw
                df = df.fillna("")

            # Normalise column names (strip whitespace)
            df.columns = [str(c).strip() for c in df.columns]
            return df

        except Exception as e:
            raise ValueError(
                f"BactochemGroundwaterParser: cannot read file — {e}"
            ) from e

    @staticmethod
    def _short_date(date_str: str) -> str:
        m = re.search(r"(\d{4})-(\d{2})-(\d{2})", date_str)
        if m:
            return f"{m.group(3)}.{m.group(2)}.{m.group(1)}"
        return ""
