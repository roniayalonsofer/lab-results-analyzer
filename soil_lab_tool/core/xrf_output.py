"""
core/xrf_output.py
-------------------
Plain Index × Element Excel table for XRF soil data.
No thresholds, no CAS numbers, no colour coding.
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
_IDX_ALIGN  = Alignment(horizontal="left",   vertical="center")


def _fmt(value, flag: str, lod) -> object:
    """Return the display value for one XRF cell."""
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
    Write a plain Index × Element Excel table.

    Column A : Index (instrument reading number)
    Columns B+: element symbols in order of first appearance (Mo, Zr, Sr, Pb, As …)
    Values    : numeric for detects; '< LOD' or '< {limit}' for non-detects.
    """
    seen_idx: set  = set()
    indices:  list = []
    seen_sym: set  = set()
    symbols:  list = []

    for r in records:
        idx = str(r.get("sample_id", "")).strip()
        sym = r.get("compound", "")
        if idx and idx not in seen_idx:
            seen_idx.add(idx)
            indices.append(idx)
        if sym and sym not in seen_sym:
            seen_sym.add(sym)
            symbols.append(sym)

    # lookup[index][symbol] = display value
    lookup: dict[str, dict[str, object]] = {}
    for r in records:
        idx = str(r.get("sample_id", "")).strip()
        sym = r.get("compound", "")
        lookup.setdefault(idx, {})[sym] = _fmt(
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

    _hdr(1, "Index")
    for ci, sym in enumerate(symbols, 2):
        _hdr(ci, sym)

    # ── Data rows ────────────────────────────────────────────────────
    for ri, idx in enumerate(indices, 2):
        c = ws.cell(ri, 1, idx)
        c.font = _DATA_FONT
        c.alignment = _IDX_ALIGN
        row_vals = lookup.get(idx, {})
        for ci, sym in enumerate(symbols, 2):
            c = ws.cell(ri, ci, row_vals.get(sym, ""))
            c.font      = _DATA_FONT
            c.alignment = _DATA_ALIGN

    # ── Column widths ────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 10
    for ci in range(2, 2 + len(symbols)):
        ws.column_dimensions[get_column_letter(ci)].width = 8

    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "B2"

    wb.save(output_path)
