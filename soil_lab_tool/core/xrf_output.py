"""
core/xrf_output.py
-------------------
Professional Excel report for XRF soil-metals data.

Layout (Index × Elements — transposed from the standard lab report):
  Rows 1-3  : Project header (name, client, date)
  Row  4    : Column headers  → "Index" | "Mo" | "Zr" | "Pb" | "As" | …
  Row  5    : CAS numbers
  Row  6    : Units           (mg/kg for all elements)
  Row  7+   : One row per selected threshold (VSL, Tier-1 …)
  Row  N+   : Data rows       → index value | measurements (highlighted on exceedance)
  Last rows : Legend
"""
from __future__ import annotations

import io
from typing import TYPE_CHECKING

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

if TYPE_CHECKING:
    from core.threshold_manager import ThresholdManager

# ── Palette ───────────────────────────────────────────────────────────────────
_C_DARK    = "2D4A5A"   # header / column-label background
_C_MID     = "4A7A8A"   # subtitle bar
_C_THRESH  = "D5E8EC"   # threshold rows (VSL)
_C_TIER1   = "FFF3E0"   # threshold rows (Tier-1)
_C_META    = "EEF3F5"   # CAS / units rows
_C_ALT     = "F8FAFC"   # alternate data row
_C_WHITE   = "FFFFFF"
_C_YELLOW  = "FFFF00"   # VSL exceedance
_C_ORANGE  = "FFC000"   # Tier-1 exceedance
_C_LEGEND  = "E2E8F0"   # legend strip

# ── Styles ────────────────────────────────────────────────────────────────────
def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)

def _font(bold=False, color="000000", size=9, name="Arial") -> Font:
    return Font(bold=bold, color=color, size=size, name=name)

def _align(h="center", v="center", wrap=False) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

_THIN = Side(style="thin", color="BFCDD4")

def _border(top=False, bottom=False, left=False, right=False) -> Border:
    t = _THIN if top    else None
    b = _THIN if bottom else None
    l = _THIN if left   else None
    r = _THIN if right  else None
    return Border(top=t, bottom=b, left=l, right=r)

_ALL_BORDER = Border(top=_THIN, bottom=_THIN, left=_THIN, right=_THIN)

# Threshold keys that map to VSL (yellow on exceedance); others → Tier-1 (orange)
_VSL_KEYS = frozenset({"VSL_SOIL", "GW", "PFAS_VSL"})


# ── Helpers ───────────────────────────────────────────────────────────────────

def _fmt_value(value, flag: str, lod) -> object:
    """Display value for one data cell: number, '< {lod}', or '< LOD'."""
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


def _fmt_thresh(v) -> object:
    """Format a threshold value for display."""
    if v is None:
        return "—"
    if isinstance(v, float):
        if v == int(v):
            return int(v)
        return round(v, 2)
    return v


def _exceeds(raw_value, flag: str, thresh) -> bool:
    """True when a detected value exceeds a numeric threshold."""
    if flag != "" or raw_value is None or thresh is None:
        return False
    try:
        return float(raw_value) > float(thresh)
    except (TypeError, ValueError):
        return False


def _write(ws, row: int, col: int, value,
           fill=None, font=None, align=None, border=None) -> None:
    c = ws.cell(row, col, value)
    if fill   is not None: c.fill      = fill
    if font   is not None: c.font      = font
    if align  is not None: c.alignment = align
    if border is not None: c.border    = border


# ── Main builder ─────────────────────────────────────────────────────────────

def build_xrf_simple_excel(
    records: list[dict],
    output_path: str | io.BytesIO,
    threshold_manager: "ThresholdManager | None" = None,
    selected_thresholds: list[str] | None = None,
    project_name: str = "",
    client: str = "",
    report_date: str = "",
) -> None:
    """
    Write a professional Index × Element Excel table for XRF records.

    Parameters
    ----------
    records             : parsed XRF records (sample_id=Index, compound=symbol)
    output_path         : file path or BytesIO
    threshold_manager   : loaded ThresholdManager (optional; omit = no thresh rows)
    selected_thresholds : list of threshold keys (e.g. ['VSL_SOIL', 'TIER1_RES_SOIL_VH'])
    project_name        : shown in header
    client              : shown in header
    report_date         : shown in header (DD.MM.YYYY)
    """
    from core.threshold_manager import THRESHOLD_LABELS

    thresh_keys = selected_thresholds or []
    has_thresh  = bool(threshold_manager and thresh_keys)

    # ── Collect ordered unique indices and element symbols ───────────
    seen_idx: set  = set()
    indices:  list = []
    seen_sym: set  = set()
    symbols:  list = []
    cas_map:  dict = {}   # symbol → CAS

    for r in records:
        idx = str(r.get("sample_id", "")).strip()
        sym = r.get("compound", "")
        cas = r.get("cas", "")
        if idx and idx not in seen_idx:
            seen_idx.add(idx)
            indices.append(idx)
        if sym and sym not in seen_sym:
            seen_sym.add(sym)
            symbols.append(sym)
            if cas:
                cas_map[sym] = cas

    # Build lookup: index → symbol → (raw_value, flag, lod, display)
    lookup: dict[str, dict[str, tuple]] = {}
    for r in records:
        idx  = str(r.get("sample_id", "")).strip()
        sym  = r.get("compound", "")
        val  = r.get("value")
        flag = r.get("flag", "ND")
        lod  = r.get("lod")
        lookup.setdefault(idx, {})[sym] = (val, flag, lod, _fmt_value(val, flag, lod))

    # Pre-compute threshold values: thresh_key → symbol → value
    thresh_vals: dict[str, dict[str, object]] = {}
    if has_thresh:
        for key in thresh_keys:
            thresh_vals[key] = {}
            for sym in symbols:
                cas  = cas_map.get(sym, "")
                tv   = threshold_manager.get_threshold_with_name(cas, key, compound_name=sym)
                thresh_vals[key][sym] = tv

    # ── Workbook ─────────────────────────────────────────────────────
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "XRF Results"
    ws.right_to_left = False   # XRF table is LTR (Index on left)

    n_cols = 1 + len(symbols)   # "Index" col + one col per element

    # ── Row geometry ─────────────────────────────────────────────────
    ROW_TITLE  = 1
    ROW_CLIENT = 2
    ROW_BLANK  = 3
    ROW_HDR    = 4
    ROW_CAS    = 5
    ROW_UNIT   = 6
    row_thresh_start = 7
    row_data_start   = row_thresh_start + len(thresh_keys)

    # ── Header section ────────────────────────────────────────────────
    # Row 1: Title
    ws.merge_cells(start_row=ROW_TITLE, start_column=1,
                   end_row=ROW_TITLE,   end_column=n_cols)
    title_text = f"דוח XRF — מתכות בקרקע"
    if project_name:
        title_text = f"דוח XRF — {project_name}"
    _write(ws, ROW_TITLE, 1, title_text,
           fill  = _fill(_C_DARK),
           font  = _font(bold=True, color=_C_WHITE, size=12, name="Arial"),
           align = _align("right", wrap=True))
    ws.row_dimensions[ROW_TITLE].height = 24

    # Row 2: Client / Date sub-bar
    ws.merge_cells(start_row=ROW_CLIENT, start_column=1,
                   end_row=ROW_CLIENT,   end_column=n_cols)
    parts = []
    if client:      parts.append(f"לקוח: {client}")
    if report_date: parts.append(f"תאריך: {report_date}")
    parts.append("יחידות: mg/kg DW")
    _write(ws, ROW_CLIENT, 1, "   |   ".join(parts),
           fill  = _fill(_C_MID),
           font  = _font(bold=False, color=_C_WHITE, size=9),
           align = _align("right"))
    ws.row_dimensions[ROW_CLIENT].height = 16

    # Row 3: blank separator
    ws.row_dimensions[ROW_BLANK].height = 6

    # ── Column-header row ─────────────────────────────────────────────
    _write(ws, ROW_HDR, 1, "Index",
           fill=_fill(_C_DARK), font=_font(bold=True, color=_C_WHITE),
           align=_align("center"), border=_ALL_BORDER)
    for ci, sym in enumerate(symbols, 2):
        _write(ws, ROW_HDR, ci, sym,
               fill=_fill(_C_DARK), font=_font(bold=True, color=_C_WHITE),
               align=_align("center"), border=_ALL_BORDER)
    ws.row_dimensions[ROW_HDR].height = 20

    # ── CAS row ───────────────────────────────────────────────────────
    _write(ws, ROW_CAS, 1, "CAS",
           fill=_fill(_C_META), font=_font(bold=True, size=8),
           align=_align("center"), border=_ALL_BORDER)
    for ci, sym in enumerate(symbols, 2):
        _write(ws, ROW_CAS, ci, cas_map.get(sym, "—"),
               fill=_fill(_C_META), font=_font(size=8, color="64748B"),
               align=_align("center"), border=_ALL_BORDER)
    ws.row_dimensions[ROW_CAS].height = 14

    # ── Units row ─────────────────────────────────────────────────────
    _write(ws, ROW_UNIT, 1, "יחידות",
           fill=_fill(_C_META), font=_font(bold=True, size=8),
           align=_align("center"), border=_ALL_BORDER)
    for ci in range(2, n_cols + 1):
        _write(ws, ROW_UNIT, ci, "mg/kg",
               fill=_fill(_C_META), font=_font(size=8, color="64748B"),
               align=_align("center"), border=_ALL_BORDER)
    ws.row_dimensions[ROW_UNIT].height = 14

    # ── Threshold rows ────────────────────────────────────────────────
    for ti, key in enumerate(thresh_keys):
        row_idx  = row_thresh_start + ti
        label    = THRESHOLD_LABELS.get(key, key)
        row_fill = _fill(_C_THRESH) if key in _VSL_KEYS else _fill(_C_TIER1)

        _write(ws, row_idx, 1, label,
               fill=row_fill, font=_font(bold=True, size=8),
               align=_align("right", wrap=True), border=_ALL_BORDER)

        for ci, sym in enumerate(symbols, 2):
            tv   = thresh_vals.get(key, {}).get(sym)
            disp = _fmt_thresh(tv)
            _write(ws, row_idx, ci, disp,
                   fill=row_fill, font=_font(size=8),
                   align=_align("center"), border=_ALL_BORDER)
        ws.row_dimensions[row_idx].height = 16

    # ── Data rows ─────────────────────────────────────────────────────
    for ri, idx in enumerate(indices):
        row_idx  = row_data_start + ri
        alt_fill = _fill(_C_WHITE) if ri % 2 == 0 else _fill(_C_ALT)

        _write(ws, row_idx, 1, idx,
               fill=alt_fill, font=_font(bold=True, size=9),
               align=_align("left"), border=_ALL_BORDER)

        row_data = lookup.get(idx, {})
        for ci, sym in enumerate(symbols, 2):
            entry = row_data.get(sym)
            if entry is None:
                _write(ws, row_idx, ci, "",
                       fill=alt_fill, border=_ALL_BORDER)
                continue

            raw_val, flag, lod, display = entry

            # Determine highlight colour (Tier-1 > VSL > none)
            cell_fill = alt_fill
            if has_thresh:
                # Check Tier-1 keys first (orange, higher priority)
                for key in thresh_keys:
                    if key not in _VSL_KEYS:
                        if _exceeds(raw_val, flag, thresh_vals.get(key, {}).get(sym)):
                            cell_fill = _fill(_C_ORANGE)
                            break
                # Check VSL keys (yellow)
                if cell_fill is alt_fill:
                    for key in thresh_keys:
                        if key in _VSL_KEYS:
                            if _exceeds(raw_val, flag, thresh_vals.get(key, {}).get(sym)):
                                cell_fill = _fill(_C_YELLOW)
                                break

            is_bold = cell_fill.fgColor.rgb in (_C_ORANGE, _C_YELLOW)
            _write(ws, row_idx, ci, display,
                   fill=cell_fill, font=_font(bold=is_bold, size=9),
                   align=_align("center"), border=_ALL_BORDER)

        ws.row_dimensions[row_idx].height = 16

    # ── Legend ────────────────────────────────────────────────────────
    if has_thresh:
        legend_row = row_data_start + len(indices) + 1
        ws.merge_cells(start_row=legend_row, start_column=1,
                       end_row=legend_row,   end_column=n_cols)
        _write(ws, legend_row, 1,
               "🟡 ערך מעל VSL     🟠 ערך מעל Tier-1",
               fill=_fill(_C_LEGEND), font=_font(size=8),
               align=_align("right"))
        ws.row_dimensions[legend_row].height = 14

    # ── Column widths ─────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 10
    for ci in range(2, n_cols + 1):
        ws.column_dimensions[get_column_letter(ci)].width = 8

    # ── Freeze panes: keep Index col + header rows visible ────────────
    ws.freeze_panes = ws.cell(row_data_start, 2)

    wb.save(output_path)
