"""
parsers/groundwater/aminolab.py
--------------------------------
Parser for Aminolab groundwater PDF reports (Hebrew RTL).

PDF column layout (RTL — rightmost in PDF = test name):
    שם הבדיקה / פרמטר  →  compound name   (rightmost)
    יחידות מידה         →  units
    תוצאה               →  result value
    הערות               →  notes / qualifier (leftmost)

Three extraction strategies are tried in order:
  1. pdfplumber extract_tables() — works when table borders are present
  2. extract_words() positional layout — works for borderless tables (primary)
  3. extract_text() line-by-line fallback

Analysis type mapping
---------------------
    pH, מוליכות, conductivity, temperature, DO, ORP  → LOWFLOW
    MTBE, BTEX, benzene, toluene, ethylbenzene,
    xylene, naphthalene, styrene                      → GW_VOC

Record format (identical to KTE/Bactochem groundwater parsers)
--------------------------------------------------------------
    lab, sample_id, compound, cas, value (float|None),
    flag ('', 'ND', '<LOD', '<LOQ'), unit, lod (None), analysis_type

Debug
-----
Set AMINOLAB_DEBUG=1 in the environment or pass debug=True to the parser.
"""

from __future__ import annotations

import io
import os
import re

from parsers.base import BaseParser
from core.lab_value_parser import LabValueParser
from core.cas_lookup import name_to_cas


# ── Compound → (CAS, analysis_type) ──────────────────────────────────────────
_COMPOUND_MAP: dict[str, tuple[str, str]] = {
    # BTEX / MTBE / aromatics — GW_VOC
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
_ND_RE = re.compile(
    r"^(?:nd|n\.d\.|not\s+detected|לא\s+זוהה|לא\s+נמצא|<\s*dl|<\s*mdl|bdl)$",
    re.I,
)

# Both logical-order (Unicode) and visual-order (reversed by pdfplumber) forms.
_HEADER_RESULT_KW = ("תוצאה", "האוצת", "result")
_HEADER_NAME_KW   = ("שם", "בדיקה", "הקידב", "פרמטר", "רמטרפ",
                      "רכיב", "name", "test", "parameter")
_HEADER_UNITS_KW  = ("יחידות", "תודיחי", "unit")
_HEADER_NOTES_KW  = ("הערות", "תורעה", "note", "remark", "הסבר")


# ── RTL helpers ───────────────────────────────────────────────────────────────

def _fix_rtl(text: str) -> str:
    """Reverse RTL Hebrew text extracted visually by pdfplumber."""
    if not _HEBREW_RE.search(text):
        return text
    tokens = text.split()
    tokens.reverse()
    return " ".join(t[::-1] if _HEBREW_RE.search(t) else t for t in tokens)


def _cell_contains(cell: str, keywords: tuple) -> bool:
    fixed = _fix_rtl(cell).lower()
    raw   = cell.lower()
    return any(kw.lower() in fixed or kw.lower() in raw for kw in keywords)


# ── Compound classification ───────────────────────────────────────────────────

def _classify(name: str) -> tuple[str, str, str]:
    """Return (display_name, cas, analysis_type)."""
    key = name.strip().lower()
    if key in _COMPOUND_MAP:
        cas, atype = _COMPOUND_MAP[key]
        return name, cas or name_to_cas(key) or "", atype
    for k, (cas, atype) in _COMPOUND_MAP.items():
        if k and (k in key or key.startswith(k)):
            return name, cas or name_to_cas(k) or "", atype
    cas   = name_to_cas(key) or ""
    atype = "GW_VOC" if cas else "LOWFLOW"
    return name, cas, atype


# ── Sample ID extraction ──────────────────────────────────────────────────────

def _extract_sample_id(page) -> str:
    text  = page.extract_text() or ""
    lines = [_fix_rtl(l.strip()) for l in text.splitlines() if l.strip()]
    for line in lines[:20]:
        if re.search(r"מספר\s+(דגימה|בור|תעודה|דגם|פיזומטר)|sample\s*(id|no|number)",
                     line, re.I):
            m = re.search(r"[:\-]\s*(\S+)\s*$", line) or re.search(r"(\S+)$", line)
            if m:
                return m.group(1)
    for line in lines[:15]:
        m = re.search(r"\b(?:GW|BH|PZ|MW|OBS|MON|BOR)[-_]?\d+\b", line, re.I)
        if m:
            return m.group(0).upper()
    return "Unknown"


# ── Record builder (same keys as KTE / Bactochem groundwater parsers) ─────────

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

    if notes and re.search(
            r"<\s*(?:dl|mdl)|not\s+detected|לא\s+(?:זוהה|נמצא)", notes, re.I):
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
        "analysis_type": atype,
    }


# ── Strategy 1: pdfplumber extract_tables() ───────────────────────────────────

def _parse_page_table(page, sample_id: str, vp: LabValueParser,
                      debug: bool = False) -> list[dict]:
    """Try pdfplumber table detection (works when PDF has visible borders)."""
    strategies = [
        None,  # default (line-based)
        {"vertical_strategy": "text", "horizontal_strategy": "text",
         "snap_tolerance": 5, "join_tolerance": 5},
    ]
    for strat in strategies:
        tables = (page.extract_tables(strat) if strat else page.extract_tables()) or []
        if debug and strat is None:
            print(f"\n[AMINOLAB DEBUG] extract_tables(): {len(tables)} table(s)")
            for ti, tbl in enumerate(tables):
                print(f"  Table {ti}: {len(tbl)} rows")
                for ri, row in enumerate(tbl[:4]):
                    print(f"    row {ri}: {row}")
        recs = _tables_to_records(tables, sample_id, vp, debug)
        if recs:
            return recs
    return []


def _tables_to_records(tables, sample_id, vp, debug=False):
    records = []
    for table in tables:
        if not table or len(table) < 2:
            continue
        header_idx = col_name = col_units = col_result = col_notes = None
        for i, row in enumerate(table):
            if not row:
                continue
            cells = [str(c or "").strip() for c in row]
            # Fix RTL before checking — critical for reversed-char PDFs
            fixed = [_fix_rtl(c) for c in cells]
            if any(_cell_contains(c, _HEADER_RESULT_KW) for c in fixed):
                header_idx = i
                col_name   = next((j for j, h in enumerate(fixed)
                                   if _cell_contains(h, _HEADER_NAME_KW)), None)
                col_units  = next((j for j, h in enumerate(fixed)
                                   if _cell_contains(h, _HEADER_UNITS_KW)), None)
                col_result = next((j for j, h in enumerate(fixed)
                                   if _cell_contains(h, _HEADER_RESULT_KW)), None)
                col_notes  = next((j for j, h in enumerate(fixed)
                                   if _cell_contains(h, _HEADER_NOTES_KW)), None)
                if debug:
                    print(f"[AMINOLAB DEBUG] table header row {i}: {fixed}")
                    print(f"  cols: name={col_name} units={col_units} "
                          f"result={col_result} notes={col_notes}")
                break
        if header_idx is None or col_result is None:
            continue
        if col_name is None:
            col_name = len(table[header_idx]) - 1  # RTL: test name is rightmost
        for row in table[header_idx + 1:]:
            if not row:
                continue
            cells = [str(c or "").strip() for c in row]
            get = lambda idx: (cells[idx] if idx is not None and idx < len(cells) else "")
            rec = _make_record(
                _fix_rtl(get(col_name)), get(col_units),
                get(col_result), _fix_rtl(get(col_notes)),
                sample_id, vp,
            )
            if rec:
                records.append(rec)
    return records


# ── Strategy 2: extract_words() positional reconstruction (primary for RTL) ───

_ROW_TOL = 6   # points — words within this vertical distance are on the same row


def _parse_page_words(page, sample_id: str, vp: LabValueParser,
                      debug: bool = False) -> list[dict]:
    """
    Reconstruct the table from word bounding boxes.

    pdfplumber returns one dict per word with keys x0, x1, top, bottom, text.
    We cluster words into rows by top-coordinate, then sort left-to-right.
    For RTL PDFs the column order is (left→right): notes | result | units | name.
    We identify columns by matching header-row words to keyword sets, then
    assign data-row words to columns by x-overlap.
    """
    words = page.extract_words(
        x_tolerance=4, y_tolerance=4,
        keep_blank_chars=False, use_text_flow=False,
    ) or []
    if not words:
        return []

    if debug:
        print(f"\n[AMINOLAB DEBUG] extract_words(): {len(words)} words (first 30):")
        for w in words[:30]:
            print(f"  x0={w['x0']:.1f} top={w['top']:.1f} text={repr(w['text'])}")

    # ── Cluster into rows ─────────────────────────────────────────────────────
    row_groups: list[list[dict]] = []
    for word in sorted(words, key=lambda w: (w['top'], w['x0'])):
        for group in row_groups:
            if abs(word['top'] - group[0]['top']) <= _ROW_TOL:
                group.append(word)
                break
        else:
            row_groups.append([word])
    for g in row_groups:
        g.sort(key=lambda w: w['x0'])

    if debug:
        print(f"[AMINOLAB DEBUG] Word rows ({len(row_groups)}):")
        for ri, g in enumerate(row_groups[:20]):
            print(f"  row {ri} (top≈{g[0]['top']:.0f}): "
                  f"{[w['text'] for w in g]}")

    # ── Find header row ───────────────────────────────────────────────────────
    header_idx = None
    for i, group in enumerate(row_groups):
        for word in group:
            if _cell_contains(word['text'], _HEADER_RESULT_KW):
                header_idx = i
                break
        if header_idx is not None:
            break

    if header_idx is None:
        if debug:
            print("[AMINOLAB DEBUG] word-parser: no header row found")
        return []

    header_words = row_groups[header_idx]
    if debug:
        print(f"[AMINOLAB DEBUG] Header row {header_idx}: "
              f"{[w['text'] for w in header_words]}")

    # ── Build column spans from header words ──────────────────────────────────
    # Each header word centres a column; column boundaries are midpoints between
    # adjacent header words.
    n = len(header_words)
    page_w = float(page.width)
    col_specs: list[tuple[float, float, str]] = []  # (x_start, x_end, col_type)
    for i, hw in enumerate(header_words):
        x_start = (0.0 if i == 0
                   else (header_words[i - 1]['x1'] + hw['x0']) / 2)
        x_end   = (page_w if i == n - 1
                   else (hw['x1'] + header_words[i + 1]['x0']) / 2)
        cname   = _fix_rtl(hw['text'])
        if _cell_contains(cname, _HEADER_NAME_KW):
            ctype = "name"
        elif _cell_contains(cname, _HEADER_UNITS_KW):
            ctype = "units"
        elif _cell_contains(cname, _HEADER_RESULT_KW):
            ctype = "result"
        elif _cell_contains(cname, _HEADER_NOTES_KW):
            ctype = "notes"
        else:
            ctype = "unknown"
        col_specs.append((x_start, x_end, ctype))

    # Fallback: if name column not found, the rightmost column is the test name
    # (RTL: שם הבדיקה is on the right side of the page = highest x values)
    if not any(ct == "name" for _, _, ct in col_specs) and col_specs:
        col_specs[-1] = (col_specs[-1][0], col_specs[-1][1], "name")

    if debug:
        print("[AMINOLAB DEBUG] Column specs:")
        for xs, xe, ct in col_specs:
            print(f"  x {xs:.0f}–{xe:.0f} → {ct}")

    # ── Parse data rows ───────────────────────────────────────────────────────
    records = []
    for group in row_groups[header_idx + 1:]:
        buckets: dict[str, list[str]] = {
            "name": [], "units": [], "result": [], "notes": []}
        for word in group:
            wx = (word['x0'] + word['x1']) / 2
            for xs, xe, ctype in col_specs:
                if xs <= wx <= xe:
                    if ctype in buckets:
                        buckets[ctype].append(word['text'])
                    break

        name   = _fix_rtl(" ".join(buckets["name"]))
        units  = " ".join(buckets["units"])
        result = " ".join(buckets["result"])
        notes  = _fix_rtl(" ".join(buckets["notes"]))

        if debug and (name or result):
            print(f"[AMINOLAB DEBUG] data row: "
                  f"name={repr(name)} result={repr(result)} "
                  f"units={repr(units)} notes={repr(notes)}")

        rec = _make_record(name, units, result, notes, sample_id, vp)
        if rec:
            records.append(rec)

    return records


# ── Strategy 3: extract_text() line-by-line ───────────────────────────────────

_TABLE_START_RE = re.compile(
    r"תוצאה|האוצת|result|הבדיקה|הקידבה|פרמטר|רמטרפ", re.I)


def _parse_page_text(page, sample_id: str, vp: LabValueParser,
                     debug: bool = False) -> list[dict]:
    """Last-resort line-by-line parser for when word/table extraction fails."""
    text = page.extract_text() or ""
    records = []
    in_table = False

    if debug:
        print("\n[AMINOLAB DEBUG] Raw page text (extract_text):")
        for ln, line in enumerate(text.splitlines()):
            print(f"  {ln:3d}: {repr(line)}")

    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        if _TABLE_START_RE.search(line):
            in_table = True
            continue
        if not in_table:
            continue
        parts = re.split(r"[ \t]{2,}", line)
        if len(parts) < 2:
            continue
        result_idx = next(
            (i for i, p in enumerate(parts)
             if re.match(r"^[<>]?\s*[\d.]", p) or _ND_RE.match(p)),
            None,
        )
        if result_idx is None:
            continue
        raw_result = parts[result_idx]
        raw_units  = parts[result_idx - 1].strip() if result_idx > 0 else ""
        raw_name   = parts[-1] if result_idx < len(parts) - 1 else parts[0]
        rec = _make_record(
            _fix_rtl(raw_name), raw_units, raw_result, "", sample_id, vp)
        if rec:
            records.append(rec)

    return records


# ── Parser class ──────────────────────────────────────────────────────────────

class AminolabGroundwaterParser(BaseParser):
    """Aminolab groundwater PDF parser — matches KTE/Bactochem record format."""

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
                    print(f"[AMINOLAB DEBUG] PAGE {page_num}  "
                          f"(size {page.width:.0f}×{page.height:.0f})")
                    print(f"{'='*60}")

                if page_num == 0:
                    sample_id = _extract_sample_id(page)
                    if self._debug:
                        print(f"[AMINOLAB DEBUG] sample_id={sample_id!r}")

                # Strategy 1: table lines
                page_recs = _parse_page_table(page, sample_id, self._vp, self._debug)
                if self._debug:
                    print(f"[AMINOLAB DEBUG] Strategy 1 (tables): {len(page_recs)} records")

                # Strategy 2: word positions (best for borderless RTL PDFs)
                if not page_recs:
                    page_recs = _parse_page_words(page, sample_id, self._vp, self._debug)
                    if self._debug:
                        print(f"[AMINOLAB DEBUG] Strategy 2 (words): {len(page_recs)} records")

                # Strategy 3: text lines
                if not page_recs:
                    page_recs = _parse_page_text(page, sample_id, self._vp, self._debug)
                    if self._debug:
                        print(f"[AMINOLAB DEBUG] Strategy 3 (text): {len(page_recs)} records")

                records.extend(page_recs)

        if self._debug:
            print(f"\n[AMINOLAB DEBUG] Total records: {len(records)}")
            for r in records:
                print(f"  {r['compound']:30s}  {r['value']}  {r['flag']}  "
                      f"{r['unit']}  {r['analysis_type']}")

        return records
