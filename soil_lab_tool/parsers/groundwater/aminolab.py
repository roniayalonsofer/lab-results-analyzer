"""
parsers/groundwater/aminolab.py
--------------------------------
Parser for Aminolab groundwater PDF reports (Hebrew RTL).

Expected table columns (RTL — rightmost in PDF = test name):
    שם הבדיקה / פרמטר  →  compound name
    יחידות              →  units
    תוצאה               →  result value
    הערות               →  notes / qualifier

Analysis type mapping
---------------------
    pH, מוליכות, conductivity, temperature, DO, ORP  → LOWFLOW
    MTBE, BTEX, benzene, toluene, ethylbenzene,
    xylene, naphthalene, styrene                      → GW_VOC

Debug
-----
Set AMINOLAB_DEBUG=1 in the environment to print raw pdfplumber output
to stdout, or pass debug=True to AminolabGroundwaterParser().
"""

from __future__ import annotations

import io
import os
import re

from parsers.base import BaseParser
from core.lab_value_parser import LabValueParser
from core.cas_lookup import name_to_cas


# ── Compound → (CAS, analysis_type) ─────────────────────────────────────────
_COMPOUND_MAP: dict[str, tuple[str, str]] = {
    # BTEX / MTBE — GW_VOC
    "mtbe":               ("1634-04-4", "GW_VOC"),
    "btex":               ("",          "GW_VOC"),
    "btex + mtbe":        ("",          "GW_VOC"),
    "btex+mtbe":          ("",          "GW_VOC"),
    "benzene":            ("71-43-2",   "GW_VOC"),
    "בנזן":               ("71-43-2",   "GW_VOC"),
    "toluene":            ("108-88-3",  "GW_VOC"),
    "טולואן":             ("108-88-3",  "GW_VOC"),
    "ethyl benzene":      ("100-41-4",  "GW_VOC"),
    "ethylbenzene":       ("100-41-4",  "GW_VOC"),
    "אתיל בנזן":          ("100-41-4",  "GW_VOC"),
    "אתיל-בנזן":          ("100-41-4",  "GW_VOC"),
    "אתילבנזן":           ("100-41-4",  "GW_VOC"),
    "xylene":             ("1330-20-7", "GW_VOC"),
    "xylenes":            ("1330-20-7", "GW_VOC"),
    "קסילן":              ("1330-20-7", "GW_VOC"),
    "קסילנים":            ("1330-20-7", "GW_VOC"),
    "o-xylene":           ("95-47-6",   "GW_VOC"),
    "m-xylene":           ("108-38-3",  "GW_VOC"),
    "p-xylene":           ("106-42-3",  "GW_VOC"),
    "naphthalene":        ("91-20-3",   "GW_VOC"),
    "נפטלן":              ("91-20-3",   "GW_VOC"),
    "styrene":            ("100-42-5",  "GW_VOC"),
    "סטירן":              ("100-42-5",  "GW_VOC"),
    # Field parameters — LOWFLOW (no environmental threshold)
    "ph":                 ("", "LOWFLOW"),
    "חומציות":            ("", "LOWFLOW"),
    "conductivity":       ("", "LOWFLOW"),
    "מוליכות":            ("", "LOWFLOW"),
    "מוליכות חשמלית":     ("", "LOWFLOW"),
    "temperature":        ("", "LOWFLOW"),
    "טמפרטורה":           ("", "LOWFLOW"),
    "temp":               ("", "LOWFLOW"),
    "do":                 ("", "LOWFLOW"),
    "dissolved oxygen":   ("", "LOWFLOW"),
    "חמצן מומס":          ("", "LOWFLOW"),
    "orp":                ("", "LOWFLOW"),
    "turbidity":          ("", "LOWFLOW"),
    "עכירות":             ("", "LOWFLOW"),
    "depth to water":     ("", "LOWFLOW"),
    "depth":              ("", "LOWFLOW"),
    "עומק":               ("", "LOWFLOW"),
    "עומק מי תהום":       ("", "LOWFLOW"),
}

_HEBREW_RE = re.compile(r"[א-ת]")
_ND_RE     = re.compile(
    r"^(?:nd|n\.d\.|not\s+detected|לא\s+זוהה|לא\s+נמצא|<\s*dl|<\s*mdl|bdl)$",
    re.I,
)

# Hebrew header keywords and their visually-reversed forms as extracted by
# pdfplumber from RTL PDFs stored in visual (left-to-right) character order.
# Checked alongside the normal forms so both PDF encoding styles work.
_HEADER_RESULT_KEYWORDS = ("תוצאה", "האוצת", "result")
_HEADER_NAME_KEYWORDS   = ("שם", "בדיקה", "הקידב", "פרמטר", "רמטרפ",
                            "רכיב", "name", "test", "parameter")
_HEADER_UNITS_KEYWORDS  = ("יחידות", "תודיחי", "unit")
_HEADER_NOTES_KEYWORDS  = ("הערות", "תורעה", "note", "remark", "הסבר")


# ── RTL helpers ───────────────────────────────────────────────────────────────

def _fix_rtl(text: str) -> str:
    """Reverse RTL Hebrew text extracted visually by pdfplumber."""
    if not _HEBREW_RE.search(text):
        return text
    tokens = text.split()
    tokens.reverse()
    return " ".join(t[::-1] if _HEBREW_RE.search(t) else t for t in tokens)


# ── Compound classification ───────────────────────────────────────────────────

def _classify(name: str) -> tuple[str, str, str]:
    """Return (display_name, cas, analysis_type) for a raw test-name string."""
    key = name.strip().lower()

    if key in _COMPOUND_MAP:
        cas, atype = _COMPOUND_MAP[key]
        if not cas:
            cas = name_to_cas(key) or ""
        return name, cas, atype

    for k, (cas, atype) in _COMPOUND_MAP.items():
        if k and (k in key or key.startswith(k)):
            if not cas:
                cas = name_to_cas(k) or ""
            return name, cas, atype

    cas   = name_to_cas(key) or ""
    atype = "GW_VOC" if cas else "LOWFLOW"
    return name, cas, atype


# ── Sample ID extraction ──────────────────────────────────────────────────────

def _extract_sample_id(page) -> str:
    """Try to pull a well / sample ID from the first lines of the page."""
    text = page.extract_text() or ""
    lines = [_fix_rtl(l.strip()) for l in text.splitlines() if l.strip()]

    for line in lines[:20]:
        if re.search(r"מספר\s+(דגימה|בור|תעודה|דגם|פיזומטר)|sample\s*(id|no|number)", line, re.I):
            m = re.search(r"[:\-]\s*(\S+)\s*$", line) or re.search(r"(\S+)$", line)
            if m:
                return m.group(1)

    for line in lines[:15]:
        m = re.search(r"\b(?:GW|BH|PZ|MW|OBS|MON|BOR)[-_]?\d+\b", line, re.I)
        if m:
            return m.group(0).upper()

    return "Unknown"


# ── Record builder ────────────────────────────────────────────────────────────

def _make_record(
    name: str,
    raw_units: str,
    raw_result: str,
    notes: str,
    sample_id: str,
    vp: LabValueParser,
) -> dict | None:
    name = name.strip()
    if not name or name.lower() in ("nan", "none", ""):
        return None

    raw_result = (raw_result or "").strip()

    if not raw_result or _ND_RE.match(raw_result):
        value, flag = None, "ND"
    else:
        value, flag = vp.parse(raw_result)

    if notes and re.search(r"<\s*(?:dl|mdl)|not\s+detected|לא\s+(?:זוהה|נמצא)", notes, re.I):
        value, flag = None, "ND"

    compound, cas, atype = _classify(name)

    unit = _fix_rtl(raw_units.strip()) if raw_units else ""
    if not unit:
        unit = "mg/L" if atype == "GW_VOC" else ""

    return {
        "lab":           "Aminolab",
        "sample_id":     sample_id,
        "compound":      compound,
        "cas":           cas,
        "value":         value,
        "flag":          flag,
        "unit":          unit,
        "lod":           None,
        "loq":           None,
        "analysis_type": atype,
    }


# ── Table parser ──────────────────────────────────────────────────────────────

def _col_index(header_row: list[str], *keywords: str) -> int | None:
    """Return the first column index whose header contains any keyword."""
    for j, h in enumerate(header_row):
        hl = h.lower()
        if any(kw.lower() in hl for kw in keywords):
            return j
    return None


def _parse_page_table(page, sample_id: str, vp: LabValueParser,
                      debug: bool = False) -> list[dict]:
    """Extract records using pdfplumber table detection.

    Tries three extraction strategies in order:
      1. Default (line-based) — works for PDFs with visible table borders.
      2. Text-based vertical + horizontal — works for borderless tables.
      3. Text-based vertical + explicit horizontal — aggressive snap.
    """
    strategies = [
        {},   # pdfplumber default
        {"vertical_strategy": "text", "horizontal_strategy": "text",
         "snap_tolerance": 5, "join_tolerance": 5},
        {"vertical_strategy": "text", "horizontal_strategy": "text",
         "snap_tolerance": 10, "join_tolerance": 10,
         "intersection_tolerance": 10},
    ]

    for strategy in strategies:
        tables = page.extract_tables(strategy) if strategy else (page.extract_tables() or [])
        tables = tables or []

        if debug and strategy == {}:
            print(f"\n[AMINOLAB DEBUG] extract_tables() found {len(tables)} table(s)")
            for ti, tbl in enumerate(tables):
                print(f"  Table {ti}: {len(tbl)} rows")
                for ri, row in enumerate(tbl[:5]):
                    print(f"    row {ri}: {row}")

        records = _tables_to_records(tables, sample_id, vp, debug)
        if records:
            return records

    return []


def _tables_to_records(tables: list, sample_id: str, vp: LabValueParser,
                       debug: bool = False) -> list[dict]:
    records = []
    for table in tables:
        if not table or len(table) < 2:
            continue

        header_idx: int | None = None
        col_name = col_units = col_result = col_notes = None

        for i, row in enumerate(table):
            if not row:
                continue
            cells = [str(c or "").strip() for c in row]
            # Apply RTL fix BEFORE checking for Hebrew header keywords so
            # both logical-order and visual-order (reversed) PDFs are handled.
            fixed = [_fix_rtl(c) for c in cells]

            if any(kw in c.lower() for c in fixed for kw in _HEADER_RESULT_KEYWORDS):
                header_idx = i
                if debug:
                    print(f"[AMINOLAB DEBUG] Header row {i}: raw={cells}  fixed={fixed}")
                col_name   = _col_index(fixed, *_HEADER_NAME_KEYWORDS)
                col_units  = _col_index(fixed, *_HEADER_UNITS_KEYWORDS)
                col_result = _col_index(fixed, *_HEADER_RESULT_KEYWORDS)
                col_notes  = _col_index(fixed, *_HEADER_NOTES_KEYWORDS)
                if debug:
                    print(f"  col_name={col_name} col_units={col_units} "
                          f"col_result={col_result} col_notes={col_notes}")
                break

        if header_idx is None or col_result is None:
            continue

        if col_name is None:
            col_name = len(table[header_idx]) - 1  # RTL: test name is rightmost

        for row in table[header_idx + 1:]:
            if not row:
                continue
            cells = [str(c or "").strip() for c in row]

            def _get(idx):
                return cells[idx] if idx is not None and idx < len(cells) else ""

            rec = _make_record(
                _fix_rtl(_get(col_name)),
                _get(col_units),
                _get(col_result),
                _fix_rtl(_get(col_notes)),
                sample_id,
                vp,
            )
            if rec:
                records.append(rec)

    return records


# ── Text fallback parser ──────────────────────────────────────────────────────

# Both logical-order and visual-order (reversed) Hebrew header markers.
_TEXT_TABLE_START_RE = re.compile(
    r"תוצאה|האוצת|result|הבדיקה|הקידבה|פרמטר|רמטרפ",
    re.I,
)


def _parse_page_text(page, sample_id: str, vp: LabValueParser,
                     debug: bool = False) -> list[dict]:
    """Line-by-line fallback when table extraction yields nothing."""
    text = page.extract_text() or ""
    records = []
    in_table = False

    if debug:
        print("\n[AMINOLAB DEBUG] Full page text (raw):")
        for ln, line in enumerate(text.splitlines()):
            print(f"  {ln:3d}: {repr(line)}")

    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue

        # Detect table start using both normal and reversed Hebrew keywords
        if _TEXT_TABLE_START_RE.search(line):
            in_table = True
            continue

        if not in_table:
            continue

        parts = re.split(r"[ \t]{2,}", line)
        if len(parts) < 2:
            continue

        result_idx = None
        for i, p in enumerate(parts):
            if re.match(r"^[<>]?\s*[\d.]", p) or _ND_RE.match(p):
                result_idx = i
                break

        if result_idx is None:
            continue

        raw_result = parts[result_idx]
        raw_units  = parts[result_idx - 1].strip() if result_idx > 0 else ""
        raw_name   = parts[-1] if result_idx < len(parts) - 1 else parts[0]

        rec = _make_record(_fix_rtl(raw_name), raw_units, raw_result, "", sample_id, vp)
        if rec:
            records.append(rec)

    return records


# ── Parser class ──────────────────────────────────────────────────────────────

class AminolabGroundwaterParser(BaseParser):
    """Aminolab groundwater PDF report parser (Hebrew RTL)."""

    LAB_NAME       = "Aminolab"
    ANALYSIS_TYPES = ["GW_VOC", "LOWFLOW"]

    def __init__(self, debug: bool | None = None):
        self._vp    = LabValueParser()
        self._debug = debug if debug is not None else bool(os.environ.get("AMINOLAB_DEBUG"))

    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        try:
            import pdfplumber
        except ImportError as exc:
            raise ImportError("pdfplumber is required: pip install pdfplumber") from exc

        records: list[dict] = []
        sample_id = "Unknown"

        with pdfplumber.open(file_obj) as pdf:
            for page_num, page in enumerate(pdf.pages):
                if self._debug:
                    print(f"\n{'='*60}")
                    print(f"[AMINOLAB DEBUG] PAGE {page_num}")
                    print(f"{'='*60}")
                    raw_text = page.extract_text() or ""
                    print("[AMINOLAB DEBUG] Raw page text:")
                    for ln, line in enumerate(raw_text.splitlines()):
                        print(f"  {ln:3d}: {repr(line)}")

                if page_num == 0:
                    sample_id = _extract_sample_id(page)
                    if self._debug:
                        print(f"[AMINOLAB DEBUG] sample_id = {sample_id!r}")

                page_recs = _parse_page_table(page, sample_id, self._vp, self._debug)
                if self._debug:
                    print(f"[AMINOLAB DEBUG] table records: {len(page_recs)}")

                if not page_recs:
                    page_recs = _parse_page_text(page, sample_id, self._vp, self._debug)
                    if self._debug:
                        print(f"[AMINOLAB DEBUG] text-fallback records: {len(page_recs)}")

                records.extend(page_recs)

        if self._debug:
            print(f"\n[AMINOLAB DEBUG] Total records: {len(records)}")
            for r in records:
                print(f"  {r}")

        return records
