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
"""

from __future__ import annotations

import io
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

    # Exact match
    if key in _COMPOUND_MAP:
        cas, atype = _COMPOUND_MAP[key]
        if not cas:
            cas = name_to_cas(key) or ""
        return name, cas, atype

    # Substring / prefix match (handles "BTEX + MTBE (סה\"כ)" etc.)
    for k, (cas, atype) in _COMPOUND_MAP.items():
        if k and (k in key or key.startswith(k)):
            if not cas:
                cas = name_to_cas(k) or ""
            return name, cas, atype

    # CAS lookup by name
    cas = name_to_cas(key) or ""
    atype = "GW_VOC" if cas else "LOWFLOW"
    return name, cas, atype


# ── Sample ID extraction ──────────────────────────────────────────────────────

def _extract_sample_id(page) -> str:
    """Try to pull a well / sample ID from the first lines of the page."""
    text = page.extract_text() or ""
    lines = [_fix_rtl(l.strip()) for l in text.splitlines() if l.strip()]

    # Look for explicit label patterns
    for line in lines[:20]:
        # "מספר דגימה", "מספר בור", "תעודה", "מספר פיזומטר"
        if re.search(r"מספר\s+(דגימה|בור|תעודה|דגם|פיזומטר)|sample\s*(id|no|number)", line, re.I):
            m = re.search(r"[:\-]\s*(\S+)\s*$", line) or re.search(r"(\S+)$", line)
            if m:
                return m.group(1)

    # Well codes: GW-1, BH-1, PZ-1, MW-1, OBS-1
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

    # Notes can carry qualifiers like "<DL" or "לא נמצא"
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
        if any(kw in hl for kw in keywords):
            return j
    return None


def _parse_page_table(page, sample_id: str, vp: LabValueParser) -> list[dict]:
    """Extract records using pdfplumber table detection."""
    tables = page.extract_tables() or []
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

            # Identify header row by presence of "תוצאה" or "result"
            if any("תוצאה" in c or "result" in c.lower() for c in cells):
                header_idx = i
                fixed = [_fix_rtl(c) for c in cells]
                col_name   = _col_index(fixed, "שם", "בדיקה", "פרמטר", "רכיב", "name", "test", "parameter")
                col_units  = _col_index(fixed, "יחידות", "unit")
                col_result = _col_index(fixed, "תוצאה", "result")
                col_notes  = _col_index(fixed, "הערות", "note", "remark", "הסבר")
                break

        if header_idx is None or col_result is None:
            continue

        if col_name is None:
            # RTL fallback: test name is the rightmost column
            col_name = len(table[header_idx]) - 1

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

def _parse_page_text(page, sample_id: str, vp: LabValueParser) -> list[dict]:
    """Line-by-line fallback when table extraction yields nothing."""
    text = page.extract_text() or ""
    records = []
    in_table = False

    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue

        # Detect table start
        if re.search(r"תוצאה|result|הבדיקה|פרמטר", line, re.I):
            in_table = True
            continue

        if not in_table:
            continue

        # Split on 2+ spaces or tab
        parts = re.split(r"[ \t]{2,}", line)
        if len(parts) < 2:
            continue

        # Find the part that looks like a result (number or ND)
        result_idx = None
        for i, p in enumerate(parts):
            if re.match(r"^[<>]?\s*[\d.]", p) or _ND_RE.match(p):
                result_idx = i
                break

        if result_idx is None:
            continue

        raw_result = parts[result_idx]
        raw_units  = parts[result_idx - 1].strip() if result_idx > 0 else ""
        # In RTL layout the name is usually the last token
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

    def __init__(self):
        self._vp = LabValueParser()

    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        try:
            import pdfplumber
        except ImportError as exc:
            raise ImportError("pdfplumber is required: pip install pdfplumber") from exc

        records: list[dict] = []
        sample_id = "Unknown"

        with pdfplumber.open(file_obj) as pdf:
            for page_num, page in enumerate(pdf.pages):
                if page_num == 0:
                    sample_id = _extract_sample_id(page)

                page_recs = _parse_page_table(page, sample_id, self._vp)
                if not page_recs:
                    page_recs = _parse_page_text(page, sample_id, self._vp)
                records.extend(page_recs)

        return records
