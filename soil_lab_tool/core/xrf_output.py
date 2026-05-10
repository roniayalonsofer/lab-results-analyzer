"""
core/xrf_output.py
-------------------
Simple pivot-table Excel output for XRF soil metals data.
No thresholds, no CAS, no colour coding — just a plain
Sample × Element table matching the original XRF file structure.
"""
from __future__ import annotations

import io

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

_HDR_FILL  = PatternFill("solid", fgColor="2D4A5A")
_HDR_FONT  = Font(bold=True, color="FFFFFF", name="Arial", size=9)
_HDR_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
_DATA_FONT  = Font(name="Arial", size=9)
_DATA_ALIGN = Alignment(horizontal="center", vertical="center")
_LEFT_ALIGN = Alignment(horizontal="left",   vertical="center")


def _fmt(value, flag: str, lod) -> object:
    """Return the display value for one XRF cell (number or string)."""
    if flag == "" and value is not None:
        v = float(value)
        if v == int(v) and abs(v) < 1e9:
            return int(v)
        return round(v, 3)
    if flag == "<LOD" and lod is not None:
        lf = float(lod)
        s  = f"{lf:.4f}".rstrip("0").rstrip(".")
        return f"< {s}"
    return "< LOD"


def build_xrf_simple_excel(
    records: list[dict],
    output_path: str | io.BytesIO,
) -> None:
    """
    Write a plain Sample × Element Excel table.

    Columns: Sample | Location | element_1 | element_2 | …
    Rows   : one row per unique (sample_id, location) pair, in input order.
    Values : numeric for detects; '< LOD' or '< {lod}' for non-detects.
    """
    seen_samples: set = set()
    samples: list[tuple[str, str]] = []
    seen_elems: set = set()
    elements: list[str] = []

    for r in records:
        sid = str(r.get("sample_id", "")).strip()
        loc = str(r.get("location",  "")).strip()
        key = (sid, loc)
        if key not in seen_samples:
            seen_samples.add(key)
            samples.append(key)
        cmp = r.get("compound", "")
        if cmp and cmp not in seen_elems:
            seen_elems.add(cmp)
            elements.append(cmp)

    # Build lookup: (sample_id, location) → compound → display value
    lookup: dict[tuple[str, str], dict[str, object]] = {}
    for r in records:
        sid = str(r.get("sample_id", "")).strip()
        loc = str(r.get("location",  "")).strip()
        lookup.setdefault((sid, loc), {})[r.get("compound", "")] = _fmt(
            r.get("value"), r.get("flag", "ND"), r.get("lod")
        )

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "XRF Results"

    # ── Header row ───────────────────────────────────────────────────
    def _hdr(col: int, text: str) -> None:
        c = ws.cell(1, col, text)
        c.fill      = _HDR_FILL
        c.font      = _HDR_FONT
        c.alignment = _HDR_ALIGN

    _hdr(1, "Sample")
    _hdr(2, "Location")
    for ci, elem in enumerate(elements, 3):
        _hdr(ci, elem)

    # ── Data rows ────────────────────────────────────────────────────
    for ri, (sid, loc) in enumerate(samples, 2):
        c = ws.cell(ri, 1, sid);  c.font = _DATA_FONT; c.alignment = _LEFT_ALIGN
        c = ws.cell(ri, 2, loc);  c.font = _DATA_FONT; c.alignment = _LEFT_ALIGN
        row_vals = lookup.get((sid, loc), {})
        for ci, elem in enumerate(elements, 3):
            c = ws.cell(ri, ci, row_vals.get(elem, ""))
            c.font      = _DATA_FONT
            c.alignment = _DATA_ALIGN

    # ── Column widths ────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 14
    for ci in range(3, 3 + len(elements)):
        ws.column_dimensions[get_column_letter(ci)].width = 9

    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "C2"

    wb.save(output_path)
