"""
parsers/soil/bactochem.py
--------------------------
Parser for Bactochem laboratory PDF reports.

PDF format (Hebrew RTL, extracted with pdfplumber):
  Sample header : "{sid} :המגודה רפסמ {name} :המגודה רואית"
  Section header: line containing ICP SOIL | SVOC | VOC | TPH-DRO+ORO | Explosives | ICP | PFAS
  Data line     : "[fn] CAS #: {cas} [{loq}] {unit} [X≤threshold] {result} [{n}/] {compound}"
  Unit wrap     : "substance" on its own line is joined to the preceding "mg/kg dry" line.

Analysis type mapping:
  ICP SOIL  + mg/kg → SOIL_METALS
  SVOC      + mg/kg → SOIL_SVOC
  VOC       + mg/kg → SOIL_VOC
  TPH-DRO   + mg/kg → SOIL_TPH
  Explosives        → SOIL_EXPLOSIVES
  ICP       + mg/L  → GW_METALS
  TPH       + mg/L  → GW_TPH
  PFAS      + ng/kg → SOIL_PFAS
  PFAS      + ng/L  → GW_PFAS
"""

from __future__ import annotations

import io
import re

from parsers.base import BaseParser


# ── Patterns ──────────────────────────────────────────────────────────────────

# Sample header (reversed RTL Hebrew): "{sid} :המגודה רפסמ {name} :המגודה רואית"
_SAMPLE_RE = re.compile(
    r"(?P<sid>\S+)\s+:המגודה רפסמ\s+(?P<sname>.+?)\s+:המגודה רואית"
)

# CAS data line — fields in order:
#   [footnote] CAS #: {cas} [{loq}] {unit} [X≤threshold] {result} [{n}/] {compound}
_CAS_RE = re.compile(
    r"CAS\s*#:\s*(?:CAS\s+)?(?P<cas>[\w.+\-]+)"                               # CAS / pseudo-id (some PDFs repeat "CAS" before the number)
    r"(?:\s+(?P<loq>[\d.]+))?"                                                  # optional LOQ
    r"\s+(?P<unit>%|ng/(?:kg|L)|mg/(?:kg(?:\s+dry(?:\s+substance)?)?|L))"      # unit (inc ng/kg, ng/L)
    r"(?:\s+X[≤≥<>≠]\s*[\d.]+(?:\s*/\s*[\d.]+)?)?"                            # optional threshold
    r"\s+(?P<result>Not\s+Detected|<[\d.]+|[\d.]+(?:[Ee][+\-]?\d+)?)"          # result
    r"(?:\s+\d+/)?"                                                              # optional sample ref
    r"\s*(?P<compound>.*)$",                                                     # rest = compound name
    re.IGNORECASE,
)

# Section keywords — checked in priority order (SVOC before VOC, ICP SOIL before ICP)
_SECTION_PATTERNS: list[tuple[re.Pattern, str]] = [
    (re.compile(r"ICP\s+SOIL",                    re.I), "ICP_SOIL"),
    (re.compile(r"Explosives?\s+and\s+propellant", re.I), "EXPLOSIVES"),
    (re.compile(r"TPH[-\s]*DRO",                  re.I), "TPH"),
    (re.compile(r'\bBTEX\b',                        re.I), "VOC"),
    (re.compile(r'\bMTBE\b',                        re.I), "VOC"),
    (re.compile(r"SVOC",                           re.I), "SVOC"),
    (re.compile(r"\bVOC\b",                        re.I), "VOC"),
    (re.compile(r"\bICP\b",                        re.I), "ICP"),
    # PFAS must follow VOC/ICP checks; also catches "Perfluorinated" and "PFAS Analysis"
    (re.compile(r"\bPFAS\b|Perfluorin",            re.I), "PFAS"),
    (re.compile(r"Microbiol|חיידקים|ביולוגי",      re.I), "MICROBIOLOGY"),
]

# PFAS compound names that may appear without a CAS #: prefix
_PFAS_BARE_RE = re.compile(
    r"^(?P<compound>PF(?:OA|OS|HxS|HpA|BA|HxA|NA|DA|BS|PeA|PeS|HpS|TrDA|DoDA|TA|UnDA)"
    r"|FOSA|FHxSA)\b"
    r".*?"
    r"(?P<result>Not\s+Detected|<[\d.]+|[\d.]+(?:[Ee][+\-]?\d+)?)"
    r"\s+(?P<unit>ng/(?:kg|L))",
    re.IGNORECASE,
)

# Microbiology result line — compound + result + CFU/MPN unit (no CAS number)
_MICRO_LINE_RE = re.compile(
    r"(?P<compound>[^\t<>]{3,60}?)"
    r"[:\s]+"
    r"(?P<result><\s*[\d.]+|Not\s+Detected|ND|[\d]+(?:[.,]\d+)?(?:\s*[×xX]\s*10\^?\d+)?)"
    r"\s+"
    r"(?P<unit>CFU/(?:100\s*)?m[Ll]|MPN/(?:100\s*)?m[Ll]|m[Ll]/(?:100\s*)?CFU)",
    re.IGNORECASE | re.UNICODE,
)

# Summary lines with no CAS number — RTL visual order: "{unit} {result} {compound}"
# e.g. "mg/kg 1086.840 Total BTEX"
_SUMMARY_RE = re.compile(
    r"^(?P<unit>mg/(?:kg(?:\s+dry(?:\s+substance)?)?|L)|ng/(?:kg|L)|%)"
    r"\s+(?P<result>[\d.]+(?:[Ee][+\-]?\d+)?)"
    r"\s+(?P<compound>.+)$",
    re.IGNORECASE,
)

# BTEX / VOC bare-ND line — no CAS prefix, format: "{loq} {unit} [X≤{thresh}] NOT DETECTED {compound}"
# e.g. "0.02 mg/kg X≤ 0.28 NOT DETECTED Benzene"
_BTEX_ND_RE = re.compile(
    r"^\s*(?P<loq>[\d.]+)\s+(?P<unit>mg/(?:kg|L))"
    r"(?:\s+X[≤≥<>]\s*[\d.]+)?"
    r"\s+NOT\s+DETECTED\s+(?P<compound>.+)$",
    re.IGNORECASE,
)

# MTBE bare line — no CAS prefix, format: "{loq} {unit} {result} MTBE"
# e.g. "0.025 mg/kg 1.470 MTBE"
_MTBE_RE = re.compile(
    r"^\s*(?P<loq>[\d.]+)\s+(?P<unit>mg/kg)\s+(?P<result>[\d.]+)\s+(?P<compound>MTBE)\s*$",
    re.IGNORECASE,
)
_MTBE_CAS = "1634-04-4"

# Page footer — skip these lines
_PAGE_FOOTER_RE = re.compile(r"Page\s+\d+\s+of\s+\d+", re.I)

# ── CSV patterns ──────────────────────────────────────────────────────────────

# Extract borehole and depth from Bactochem CSV column תיאור דוגמה
# e.g. "קרקע ק1- 0.7" → bh="ק1", depth="0.7"
_CSV_DESC_RE = re.compile(
    r'(?:קרקע\s+)?(?P<bh>[^\d\s\-]+\d+)[-\s]*(?P<depth>[\d.]+)',
    re.UNICODE,
)

# Compounds containing these keywords map to SOIL_TPH
_TPH_KW_RE = re.compile(r'\b(dro|oro|tph)\b', re.IGNORECASE)

# Canonical pseudo-CAS identifiers for TPH fractions (maps to themselves;
# used as a normalisation reference and for threshold lookups in both PDF and CSV paths).
_PSEUDO_CAS_MAP: dict[str, str] = {
    "DRO":     "DRO",
    "ORO":     "ORO",
    "DRO-ORO": "DRO-ORO",
}

# Compound name (lower-case) → pseudo-CAS for the CSV parser
_CSV_COMPOUND_CAS: dict[str, str] = {
    "total dro":     "DRO",
    "total oro":     "ORO",
    "total dro+oro": "DRO-ORO",
    "dro+oro":       "DRO-ORO",
    "dro":           "DRO",
    "oro":           "ORO",
}


# ── Helpers ───────────────────────────────────────────────────────────────────

_HEBREW_RE = re.compile(r"[א-תיִ-פֿ]")


def _fix_rtl(text: str) -> str:
    """
    Correct RTL-reversed text extracted from a Hebrew PDF by pdfplumber.

    pdfplumber returns characters in visual (left-to-right) order.  For Hebrew
    text this means:
      • individual Hebrew words have their characters reversed
      • the word order across the whole string is also reversed

    Algorithm:
      1. If the string contains no Hebrew characters it is purely
         English/numeric — pdfplumber already gives these in the right order,
         so return unchanged.
      2. Otherwise reverse the list of whitespace-separated tokens, then
         reverse the character sequence of every token that contains Hebrew.
         English tokens (element symbols, abbreviations) keep their characters
         intact but move with the overall reversal.

    Examples
    --------
    'ףסכ (Ag)'   →  '(Ag) כסף'     (silver)
    'B1 עקרק'   →  'קרקע B1'       (soil borehole)
    'Benzene'    →  'Benzene'        (pure English — unchanged)
    """
    if not _HEBREW_RE.search(text):
        return text
    tokens = text.split()
    tokens.reverse()
    return " ".join(t[::-1] if _HEBREW_RE.search(t) else t for t in tokens)


def _analysis_type(section: str | None, unit: str) -> str:
    is_soil = "kg" in unit
    if section == "ICP_SOIL":
        return "SOIL_METALS"
    if section == "SVOC":
        return "SOIL_SVOC"
    if section == "VOC":
        return "SOIL_VOC"
    if section == "TPH":
        return "SOIL_TPH" if is_soil else "GW_TPH"
    if section == "EXPLOSIVES":
        return "SOIL_EXPLOSIVES"
    if section == "ICP":
        return "SOIL_METALS" if is_soil else "GW_METALS"
    if section == "PFAS":
        return "SOIL_PFAS" if is_soil else "GW_PFAS"
    # Fallback: use unit as guide when section not detected
    if "ng" in unit:
        return "SOIL_PFAS" if is_soil else "GW_PFAS"
    return "SOIL_SVOC"


def _parse_result(raw: str, loq: float | None) -> tuple[float | None, str]:
    r = raw.strip()
    if r.lower().replace(" ", "") in ("notdetected", "nd", "n.d."):
        return loq, "ND"
    if r.startswith("<"):
        try:
            return float(r[1:]), "<LOQ"
        except ValueError:
            return loq, "<LOQ"
    try:
        return float(r), ""
    except ValueError:
        return None, ""


def _parse_csv_desc(desc: str) -> tuple[str, str]:
    """Extract (borehole, depth_str) from Bactochem CSV column תיאור דוגמה.
    'קרקע ק1- 0.7' → ('ק-1', '0.7')
    """
    m = _CSV_DESC_RE.search(desc)
    if not m:
        return desc.strip(), ""
    bh = re.sub(r'([^\d\s\-])(\d)', r'\1-\2', m.group("bh"))
    return bh, m.group("depth")


def _csv_compound_cas(compound: str) -> str:
    """Return the canonical pseudo-CAS for known TPH fraction compound names."""
    return _CSV_COMPOUND_CAS.get(compound.strip().lower(), "")


def _csv_atype_loq(compound: str) -> tuple[str, float]:
    """Return (analysis_type, LOQ) inferred from the compound name."""
    if _TPH_KW_RE.search(compound):
        return "SOIL_TPH", 10.0
    return "SOIL_VOC", 0.02


def _parse_csv(file_obj: io.BytesIO) -> list[dict]:
    """Parse Bactochem CSV export format.

    Expected columns: מספר דוגמה, רכיב, תוצאה, יחידות מידה, תיאור דוגמה
    """
    try:
        import pandas as pd
    except ImportError:
        return []

    df = None
    for enc in ("utf-8-sig", "utf-8", "cp1255", "latin-1"):
        try:
            file_obj.seek(0)
            df = pd.read_csv(file_obj, dtype=str, encoding=enc).fillna("")
            df.columns = [str(c).strip() for c in df.columns]
            break
        except Exception:
            continue
    if df is None or df.empty:
        return []

    col_compound = "רכיב"
    col_result   = "תוצאה"
    col_unit     = "יחידות מידה"
    col_desc     = "תיאור דוגמה"

    if not all(c in df.columns for c in (col_compound, col_result, col_unit, col_desc)):
        return []

    records: list[dict] = []
    for _, row in df.iterrows():
        compound = str(row.get(col_compound, "")).strip()
        if not compound or compound.lower() in ("nan", ""):
            continue

        result_raw = str(row.get(col_result, "")).strip()
        unit_raw   = str(row.get(col_unit, "")).strip()
        desc       = str(row.get(col_desc, "")).strip()

        borehole, depth_str = _parse_csv_desc(desc)
        sample_id = f"{borehole} {depth_str}" if depth_str else borehole

        atype, loq = _csv_atype_loq(compound)
        value, flag = _parse_result(result_raw, loq)

        if value is None and not flag:
            continue

        unit = unit_raw if unit_raw not in ("nan", "") else "mg/kg"

        records.append({
            "sample_id":     sample_id,
            "compound":      compound,
            "cas":           _csv_compound_cas(compound),
            "value":         value,
            "flag":          flag,
            "unit":          unit,
            "lod":           None,
            "loq":           loq,
            "analysis_type": atype,
        })

    return records


# ── Parser ────────────────────────────────────────────────────────────────────

class BactochemSoilParser(BaseParser):
    """
    Bactochem PDF report parser — handles soil and groundwater samples
    in a single file, routing each CAS record to the correct analysis type.
    """

    LAB_NAME = "בקטוכם"
    ANALYSIS_TYPES = [
        "SOIL_METALS", "SOIL_SVOC", "SOIL_VOC", "SOIL_TPH", "SOIL_EXPLOSIVES",
        "SOIL_PFAS", "GW_METALS", "GW_TPH", "GW_PFAS", "GW_MICROBIOLOGY",
    ]

    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        file_obj.seek(0)
        magic = file_obj.read(4)
        file_obj.seek(0)

        if magic[:4] == b"%PDF":
            try:
                import pdfplumber
            except ImportError as exc:
                raise ImportError(
                    "pdfplumber is required: pip install pdfplumber"
                ) from exc
            lines = _extract_lines(file_obj, pdfplumber)
            return _parse_lines(lines)

        return _parse_csv(file_obj)


# ── Module-level parsing functions (testable without instantiation) ───────────

def _extract_lines(file_obj: io.BytesIO, pdfplumber) -> list[str]:
    """Extract and pre-process lines from the PDF."""
    raw: list[str] = []
    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            raw.extend(text.splitlines())
    # "substance" appears on its own line after "mg/kg dry" lines due to PDF
    # word-wrap; skipping it is safe because the unit regex matches "mg/kg dry"
    # without the trailing word, and we normalize below.
    stripped = [s for line in raw if (s := line.strip())]

    # Some PDFs split "NOT DETECTED" across two lines:
    # line N:   "0.02 mg/kg X≤ 0.28 NOT Benzene"
    # line N+1: "DETECTED"
    # Insert DETECTED right after NOT in the previous line.
    joined: list[str] = []
    for line in stripped:
        if line.strip().upper() == "DETECTED" and joined:
            prev = joined[-1]
            idx = prev.upper().rfind(" NOT ")
            if idx != -1:
                joined[-1] = prev[:idx] + " NOT DETECTED" + prev[idx + 4:]
            else:
                joined[-1] = prev + " DETECTED"
        else:
            joined.append(line)
    return joined


def _parse_lines(lines: list[str]) -> list[dict]:
    records: list[dict] = []
    sample_id   = "Unknown"
    sample_name = "Unknown"
    section: str | None = None

    for line in lines:
        if _PAGE_FOOTER_RE.search(line):
            continue

        # ── Sample header? ──────────────────────────────────────────────────
        m = _SAMPLE_RE.search(line)
        if m:
            sample_id   = m.group("sid").strip()
            sname = _fix_rtl(m.group("sname").strip())
            _bh = re.search(r'(P-\d+)', sname)
            _dp = re.search(r'([\d.]+)', sname)
            sample_name = f"{_bh.group(1)} {_dp.group(1)}" if _bh and _dp else sname
            section     = None
            continue

        # ── Section header? (only on lines that are not CAS data lines) ────
        if "CAS" not in line or "#:" not in line:
            for pat, key in _SECTION_PATTERNS:
                if pat.search(line):
                    section = key
                    break

            # ── Microbiology line (CFU/MPN unit, no CAS number) ────────
            m_micro = _MICRO_LINE_RE.search(line)
            if m_micro:
                compound = _fix_rtl(m_micro.group("compound").strip(" :-"))
                res_raw  = m_micro.group("result").strip()
                raw_unit = m_micro.group("unit").strip()
                unit_lo  = raw_unit.lower()
                if "100" in unit_lo:
                    norm_unit = "CFU/100mL"
                elif "mpn" in unit_lo:
                    norm_unit = "MPN/100mL"
                else:
                    norm_unit = "CFU/mL"
                value, flag = _parse_result(res_raw, None)
                if compound and (value is not None or flag):
                    records.append({
                        "sample_id":     sample_name,
                        "compound":      compound,
                        "cas":           "",
                        "value":         value,
                        "flag":          flag,
                        "unit":          norm_unit,
                        "lod":           None,
                        "loq":           None,
                        "analysis_type": "GW_MICROBIOLOGY",
                    })
                continue

            # ── Bare PFAS line (no CAS #: prefix) ──────────────────────
            m_pfas = _PFAS_BARE_RE.search(line)
            if m_pfas:
                compound = m_pfas.group("compound").strip().upper()
                res_raw  = m_pfas.group("result").strip()
                unit     = m_pfas.group("unit").strip()
                is_soil  = "kg" in unit.lower()
                value, flag = _parse_result(res_raw, None)
                if value is not None or flag:
                    records.append({
                        "sample_id":     sample_name,
                        "compound":      compound,
                        "cas":           "",
                        "value":         value,
                        "flag":          flag,
                        "unit":          unit,
                        "lod":           None,
                        "loq":           None,
                        "analysis_type": "SOIL_PFAS" if is_soil else "GW_PFAS",
                    })
                continue

            # ── MTBE bare line (loq-first, no CAS prefix) ──────────────
            # e.g. "0.025 mg/kg 1.470 MTBE"
            if section in ("VOC", "MBTEX"):
                m_mtbe = _MTBE_RE.match(line)
                if m_mtbe:
                    loq_v  = float(m_mtbe.group("loq"))
                    unit   = m_mtbe.group("unit").strip()
                    res_raw = m_mtbe.group("result").strip()
                    value, flag = _parse_result(res_raw, loq_v)
                    if value is not None or flag:
                        records.append({
                            "sample_id":     sample_name,
                            "compound":      "MTBE",
                            "cas":           _MTBE_CAS,
                            "value":         value,
                            "flag":          flag,
                            "unit":          unit,
                            "lod":           None,
                            "loq":           loq_v,
                            "analysis_type": "SOIL_VOC",
                        })
                    continue

            # ── BTEX/VOC bare NOT DETECTED line (no CAS prefix) ────────
            # e.g. "0.02 mg/kg X≤ 0.28 NOT DETECTED Benzene"
            m_btex_nd = _BTEX_ND_RE.match(line)
            if m_btex_nd:
                compound = m_btex_nd.group("compound").strip()
                loq_v    = float(m_btex_nd.group("loq"))
                unit     = m_btex_nd.group("unit").strip()
                records.append({
                    "sample_id":     sample_name,
                    "compound":      compound,
                    "cas":           "",
                    "value":         loq_v,
                    "flag":          "ND",
                    "unit":          unit,
                    "lod":           None,
                    "loq":           loq_v,
                    "analysis_type": "SOIL_VOC",
                })
                continue

            # ── Summary line (unit-first RTL layout, no CAS number) ────
            # e.g. "mg/kg 1086.840 Total BTEX"
            m_sum = _SUMMARY_RE.match(line)
            if m_sum:
                compound = _fix_rtl(m_sum.group("compound").strip())
                res_raw  = m_sum.group("result").strip()
                unit     = m_sum.group("unit").strip()
                value, flag = _parse_result(res_raw, None)
                if compound and (value is not None or flag):
                    records.append({
                        "sample_id":     sample_name,
                        "compound":      compound,
                        "cas":           "",
                        "value":         value,
                        "flag":          flag,
                        "unit":          unit,
                        "lod":           None,
                        "loq":           None,
                        "analysis_type": _analysis_type(section, unit),
                    })
                continue

            continue

        # ── CAS data line ───────────────────────────────────────────────────
        m2 = _CAS_RE.search(line)
        if not m2:
            continue

        cas      = m2.group("cas").strip()
        loq_raw  = m2.group("loq") or ""
        unit     = m2.group("unit").strip()
        res_raw  = m2.group("result").strip()
        compound = _fix_rtl((m2.group("compound") or "").strip())

        try:
            loq: float | None = float(loq_raw) if loq_raw else None
        except ValueError:
            loq = None

        value, flag = _parse_result(res_raw, loq)
        if value is None and not flag:
            continue

        # Normalize the soil-metals unit ("mg/kg dry" → "mg/kg dry substance")
        if unit == "mg/kg dry":
            unit = "mg/kg dry substance"

        records.append({
            "sample_id":     sample_name,
            "compound":      compound,
            "cas":           cas,
            "value":         value,
            "flag":          flag,
            "unit":          unit,
            "lod":           None,
            "loq":           loq,
            "analysis_type": _analysis_type(section, unit),
        })

    return records
