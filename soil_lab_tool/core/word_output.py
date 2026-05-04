"""
core/word_output.py
--------------------
Builds a multi-section Hebrew RTL Word lab report using python-docx.

One section per analysis type found in the records.

Layout (compounds as rows, samples as columns):
  Col 0 : תרכובת
  Col 1 : CAS Number
  Col 2…: threshold column(s)
  next  : יחידות
  rest  : one column per sample (header = borehole\\ndepth)

Colour coding (cell background):
  Yellow  (FFFF00) — detected value exceeds a VSL threshold
                     (VSL_SOIL, GW, PFAS_VSL)
  Orange  (FFC000) — detected value exceeds a Tier-1 threshold
                     (any TIER1_*, GAS_*, PFAS_TIER1_* key)
  Gray    (D3D3D3) — non-detect / below-limit but LOD > strictest threshold

Standalone helpers (add a single section to an existing Document):
  add_tph_table(doc, records, tm, selected_thresholds)
  add_metals_table(doc, records, tm, selected_thresholds)
  add_voc_table(doc, records, tm, selected_thresholds)
  add_pfas_table(doc, records, tm, selected_thresholds)

Main class:
  LabReportWord – mirrors the LabReportExcel API
"""

from __future__ import annotations

import os
from collections import defaultdict
from datetime import date

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from core.threshold_manager import (
    ThresholdManager,
    ANALYSIS_THRESHOLDS,
    THRESHOLD_LABELS,
)
from core.excel_output import SHEET_CONFIG, _ordered_unique, _split_sample_depth

# ── Colour constants ──────────────────────────────────────────────────
_YELLOW = "FFFF00"   # VSL exceedance
_ORANGE = "FFC000"   # Tier-1 exceedance
_GRAY   = "D3D3D3"   # LOD > threshold (uncertain non-detect)

# Threshold keys that yield yellow on exceedance; all others yield orange
_VSL_KEYS: frozenset[str] = frozenset({"VSL_SOIL", "GW", "PFAS_VSL"})

# Font names / sizes
_FONT_HE = "David"
_FONT_EN = "Times New Roman"
_SZ_HE   = Pt(9)
_SZ_EN   = Pt(8)
_SZ_SM   = Pt(8)   # data cells


# ── Low-level XML / formatting helpers ────────────────────────────────

def _has_hebrew(s: str) -> bool:
    return any('א' <= c <= 'ת' for c in str(s))


def _pick_font(s: str) -> tuple[str, Pt]:
    """Return (font_name, font_size) based on whether content is Hebrew."""
    s = str(s) if s is not None else ""
    has_he = _has_hebrew(s)
    has_en = any(c.isalpha() and c.isascii() for c in s)
    if has_he or not has_en:
        return _FONT_HE, _SZ_HE
    return _FONT_EN, _SZ_EN


def _set_cell_bg(cell, hex_color: str) -> None:
    """Set table cell background to a solid hex colour."""
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for ex in tcPr.findall(qn("w:shd")):
        tcPr.remove(ex)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)


def _set_rtl_para(para) -> None:
    """Enable RTL text direction on a paragraph."""
    pPr  = para._p.get_or_add_pPr()
    bidi = pPr.find(qn("w:bidi"))
    if bidi is None:
        bidi = OxmlElement("w:bidi")
        pPr.append(bidi)
    bidi.set(qn("w:val"), "1")


def _set_rtl_table(table) -> None:
    """Set bidiVisual on a table so columns display right-to-left."""
    tblPr = table._tbl.get_or_add_tblPr()
    if tblPr.find(qn("w:bidiVisual")) is None:
        tblPr.append(OxmlElement("w:bidiVisual"))


def _set_table_full_width(table) -> None:
    """Stretch table to 100 % of the text area."""
    tblPr = table._tbl.get_or_add_tblPr()
    for ex in tblPr.findall(qn("w:tblW")):
        tblPr.remove(ex)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:type"), "pct")
    tblW.set(qn("w:w"),    "5000")   # 5000 = 100 % in fiftieths of a percent
    tblPr.append(tblW)


def _clear_para_runs(para) -> None:
    p = para._p
    for r in p.findall(qn("w:r")):
        p.remove(r)


def _write_cell(
    cell,
    text: str,
    bold: bool = False,
    italic: bool = False,
    size: Pt | None = None,
    align: WD_ALIGN_PARAGRAPH = WD_ALIGN_PARAGRAPH.CENTER,
    gray_text: bool = False,
) -> None:
    """Write *text* into *cell* with auto font selection and RTL support."""
    para = cell.paragraphs[0]
    _clear_para_runs(para)
    _set_rtl_para(para)
    para.alignment = align
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    if text is None or str(text) == "":
        return

    run = para.add_run(str(text))
    fname, fsize = _pick_font(str(text))
    run.font.name   = fname
    run.font.size   = size or fsize
    run.font.bold   = bold
    run.font.italic = italic
    if gray_text:
        run.font.color.rgb = RGBColor(0x80, 0x80, 0x80)


# ── Number / value formatting ─────────────────────────────────────────

def _fmt_lod(lod: float) -> str:
    if lod == int(lod):
        return str(int(lod))
    return f"{round(lod, 3):.3f}".rstrip("0").rstrip(".")


def _fmt_thresh(val) -> str:
    if val is None:
        return "לא קיים"
    if isinstance(val, (int, float)):
        if val == int(val):
            return str(int(val))
        return f"{round(val, 2):.2f}".rstrip("0").rstrip(".")
    return str(val)


def _fmt_value(value, flag: str, lod) -> str:
    """Format a measured sample value for Word display."""
    if flag == "ND" or (value is None and flag not in ("<LOD", "<LOQ", "<")):
        return _fmt_lod(lod) if isinstance(lod, (int, float)) else "ND"
    if flag == "<LOD":
        return f"<{_fmt_lod(lod)}" if isinstance(lod, (int, float)) else "ND"
    if flag == "<LOQ":
        ref = lod if isinstance(lod, (int, float)) else value
        return _fmt_lod(ref) if isinstance(ref, (int, float)) else "ND"
    if flag == "<":
        return f"<{round(value, 2)}" if isinstance(value, (int, float)) else f"<{value}"
    if isinstance(value, float):
        return f"{round(value, 3):.3f}".rstrip("0").rstrip(".")
    return str(value) if value is not None else "ND"


# ── Colour-coding logic ───────────────────────────────────────────────

def _cell_color(
    value,
    flag: str,
    lod,
    t_vals: dict,
    thresh_keys: list[str],
) -> str | None:
    """
    Return hex background colour for a sample cell, or None.

    Priority: Orange (Tier-1) > Yellow (VSL) > Gray (LOD > strictest threshold).
    """
    is_nondetect = flag in ("ND", "<LOD", "<LOQ", "<")

    if is_nondetect:
        lod_ref: float | None = None
        if isinstance(lod, (int, float)):
            lod_ref = lod
        elif isinstance(value, (int, float)) and flag in ("<", "<LOQ"):
            lod_ref = value
        if lod_ref is not None:
            strictest = min(
                (v for v in t_vals.values() if isinstance(v, (int, float))),
                default=None,
            )
            if strictest is not None and lod_ref > strictest:
                return _GRAY
        return None

    if not isinstance(value, (int, float)):
        return None

    # Tier-1 → orange (checked first for priority)
    for k in thresh_keys:
        if k not in _VSL_KEYS:
            t = t_vals.get(k)
            if isinstance(t, (int, float)) and value > t:
                return _ORANGE

    # VSL → yellow
    for k in thresh_keys:
        if k in _VSL_KEYS:
            t = t_vals.get(k)
            if isinstance(t, (int, float)) and value > t:
                return _YELLOW

    return None


# ── Document-level helpers ────────────────────────────────────────────

def _add_section_heading(doc: Document, title: str) -> None:
    p = doc.add_paragraph()
    _set_rtl_para(p)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(title)
    run.font.name = _FONT_HE
    run.font.size = Pt(12)
    run.font.bold = True


def _add_legend(doc: Document, include_gray: bool = False) -> None:
    """Append a colour legend below a table."""
    items: list[tuple[str, str]] = [
        ("חריגה מערך סף VSL",    _YELLOW),
        ("חריגה מערך סף Tier-1", _ORANGE),
    ]
    if include_gray:
        items.append(("ספג גילוי גבוה מערך הסף", _GRAY))

    for label, hex_color in items:
        p = doc.add_paragraph()
        _set_rtl_para(p)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(f" {label} ")
        run.font.name = _FONT_HE
        run.font.size = _SZ_HE
        run.font.bold = True
        # Background colour applied as character shading (w:rPr > w:shd)
        rPr = run._r.get_or_add_rPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"),   "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"),  hex_color)
        rPr.append(shd)


def _ordered_thresh_keys(atype: str, sel: list[str] | None) -> list[str]:
    valid = set(ANALYSIS_THRESHOLDS.get(atype, []))
    if sel is not None:
        return [k for k in sel if k in valid]
    return ANALYSIS_THRESHOLDS.get(atype, [])


# ── Core table builder ────────────────────────────────────────────────

def _build_analysis_table(
    doc: Document,
    records: list[dict],
    tm: ThresholdManager,
    cfg: dict,
    thresh_keys: list[str],
) -> None:
    """
    Append one analysis-type table (compounds × samples) to *doc*.

    Header columns displayed right-to-left:
      תרכובת | CAS Number | threshold(s) | יחידות | sample_1 … sample_n
    """
    samples   = _ordered_unique(r["sample_id"] for r in records)
    compounds = _ordered_unique(r["compound"]  for r in records)
    unit      = cfg.get("unit", "")

    if not compounds or not samples:
        return

    # Pivot: compound → sample_id → (value, flag, lod)
    pivot:    dict[str, dict] = {}
    cas_map:  dict[str, str]  = {}
    unit_map: dict[str, str]  = {}
    for r in records:
        cmp = r["compound"]
        sid = r["sample_id"]
        if cmp not in pivot:
            pivot[cmp]    = {}
            cas_map[cmp]  = r.get("cas", "")
            unit_map[cmp] = r.get("unit", unit)
        pivot[cmp][sid] = (r.get("value"), r.get("flag", ""), r.get("lod"))

    # Threshold values per compound
    thresh_vals: dict[str, dict] = {
        cmp: {
            k: tm.get_threshold_with_name(cas_map[cmp], k, compound_name=cmp)
            for k in thresh_keys
        }
        for cmp in compounds
    }

    split_map = {sid: _split_sample_depth(sid) for sid in samples}

    n_thresh = len(thresh_keys)
    # Fixed cols: תרכובת(0) | CAS(1) | thresh…(2…n+1) | יחידות(n+2)
    n_fixed  = 2 + n_thresh + 1
    n_cols   = n_fixed + len(samples)

    # ── Create table ──────────────────────────────────────────────────
    table = doc.add_table(rows=1, cols=n_cols)
    table.style = "Table Grid"
    _set_rtl_table(table)
    _set_table_full_width(table)

    # ── Header row ────────────────────────────────────────────────────
    hdr = table.rows[0].cells
    _write_cell(hdr[0], "תרכובת",     bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT)
    _write_cell(hdr[1], "CAS Number",  bold=True)
    for i, k in enumerate(thresh_keys):
        _write_cell(hdr[2 + i], THRESHOLD_LABELS.get(k, k), bold=True)
    _write_cell(hdr[2 + n_thresh], "יחידות", bold=True)
    for j, sid in enumerate(samples):
        bh, dep = split_map[sid]
        hdr_text = f"{bh}\n{dep}" if dep else bh
        _write_cell(hdr[n_fixed + j], hdr_text, bold=True)

    # ── Data rows ─────────────────────────────────────────────────────
    has_gray = False

    for cmp in compounds:
        row_cells = table.add_row().cells
        cas    = cas_map.get(cmp, "")
        t_vals = thresh_vals.get(cmp, {})
        u      = unit_map.get(cmp, unit)

        _write_cell(row_cells[0], cmp, align=WD_ALIGN_PARAGRAPH.RIGHT)
        _write_cell(row_cells[1], cas if cas else "—", size=_SZ_SM)

        for i, k in enumerate(thresh_keys):
            tv   = t_vals.get(k)
            tv_r = round(tv, 2) if isinstance(tv, (int, float)) else None
            text = _fmt_thresh(tv_r)
            _write_cell(row_cells[2 + i], text,
                        italic=   (text == "לא קיים"),
                        gray_text=(text == "לא קיים"),
                        size=_SZ_SM)

        _write_cell(row_cells[2 + n_thresh], u, size=_SZ_SM)

        for j, sid in enumerate(samples):
            v, flag, lod = pivot.get(cmp, {}).get(sid, (None, "ND", None))
            display = _fmt_value(v, flag, lod)
            cell    = row_cells[n_fixed + j]

            color = _cell_color(v, flag, lod, t_vals, thresh_keys)
            if color:
                _set_cell_bg(cell, color)
                if color == _GRAY:
                    has_gray = True

            bold_val = color in (_ORANGE, _YELLOW)
            _write_cell(cell, display, bold=bold_val, size=_SZ_SM)

    doc.add_paragraph()
    _add_legend(doc, include_gray=has_gray)
    doc.add_paragraph()


# ── Standalone section-adder functions ───────────────────────────────

def add_tph_table(
    doc: Document,
    records: list[dict],
    tm: ThresholdManager,
    selected_thresholds: list[str] | None = None,
) -> None:
    """Append a TPH analysis section (heading + table + legend) to *doc*."""
    atype = "SOIL_TPH"
    recs  = [r for r in records if r.get("analysis_type") == atype]
    if not recs:
        return
    cfg         = SHEET_CONFIG.get(atype, {"name": "קרקע TPH", "unit": "mg/kg"})
    thresh_keys = _ordered_thresh_keys(atype, selected_thresholds)
    _add_section_heading(doc, cfg["name"])
    _build_analysis_table(doc, recs, tm, cfg, thresh_keys)


def add_metals_table(
    doc: Document,
    records: list[dict],
    tm: ThresholdManager,
    selected_thresholds: list[str] | None = None,
) -> None:
    """Append a Metals analysis section to *doc*."""
    atype = "SOIL_METALS"
    recs  = [r for r in records if r.get("analysis_type") == atype]
    if not recs:
        return
    cfg         = SHEET_CONFIG.get(atype, {"name": "קרקע מתכות", "unit": "mg/kg DW"})
    thresh_keys = _ordered_thresh_keys(atype, selected_thresholds)
    _add_section_heading(doc, cfg["name"])
    _build_analysis_table(doc, recs, tm, cfg, thresh_keys)


def add_voc_table(
    doc: Document,
    records: list[dict],
    tm: ThresholdManager,
    selected_thresholds: list[str] | None = None,
) -> None:
    """
    Append VOC/BTEX sections to *doc*.

    Covers SOIL_VOC, SOIL_MBTEX, GW_VOC, SOIL_GAS_VOC — one subsection each.
    """
    for atype in ("SOIL_VOC", "SOIL_MBTEX", "GW_VOC", "SOIL_GAS_VOC"):
        recs = [r for r in records if r.get("analysis_type") == atype]
        if not recs:
            continue
        cfg         = SHEET_CONFIG.get(atype, {"name": atype, "unit": ""})
        thresh_keys = _ordered_thresh_keys(atype, selected_thresholds)
        _add_section_heading(doc, cfg["name"])
        _build_analysis_table(doc, recs, tm, cfg, thresh_keys)


def add_pfas_table(
    doc: Document,
    records: list[dict],
    tm: ThresholdManager,
    selected_thresholds: list[str] | None = None,
) -> None:
    """
    Append PFAS sections to *doc*.

    Covers SOIL_PFAS and GW_PFAS — one subsection each.
    """
    for atype in ("SOIL_PFAS", "GW_PFAS"):
        recs = [r for r in records if r.get("analysis_type") == atype]
        if not recs:
            continue
        cfg         = SHEET_CONFIG.get(atype, {"name": atype, "unit": ""})
        thresh_keys = _ordered_thresh_keys(atype, selected_thresholds)
        _add_section_heading(doc, cfg["name"])
        _build_analysis_table(doc, recs, tm, cfg, thresh_keys)


# ── Main class ────────────────────────────────────────────────────────

class LabReportWord:
    """
    Build a multi-section Hebrew RTL Word lab report.

    Parameters mirror LabReportExcel for a consistent API.

    Parameters
    ----------
    records : list[dict]
        Flat list of measurement records, each with keys:
        compound, cas, sample_id, value, flag, unit, lod, analysis_type
    threshold_manager : ThresholdManager
    output_path : str | BytesIO
    project_name : str
    client : str
    report_date : str   (DD.MM.YYYY)
    selected_thresholds : list[str] | None
        Override which threshold keys to show.  None = defaults per analysis type.
    combine_tph_voc : bool
        Merge SOIL_TPH + SOIL_VOC records into one combined table.
    combine_tph_mbtex : bool
        Merge SOIL_TPH + SOIL_MBTEX records into one combined table.
    """

    def __init__(
        self,
        records: list[dict],
        threshold_manager: ThresholdManager,
        output_path: str = "lab_report.docx",
        project_name: str = "",
        client: str = "",
        report_date: str = "",
        selected_thresholds: list[str] | None = None,
        combine_tph_voc: bool = False,
        combine_tph_mbtex: bool = False,
    ):
        self.records           = records
        self.tm                = threshold_manager
        self.out_path          = output_path
        self.project           = project_name
        self.client            = client
        self.rep_date          = report_date or date.today().strftime("%d.%m.%Y")
        self.sel_thresh        = selected_thresholds
        self.combine_tph_voc   = combine_tph_voc
        self.combine_tph_mbtex = combine_tph_mbtex

    # ------------------------------------------------------------------
    def build(self) -> str:
        """Generate the Word document and return output_path."""
        doc = Document()

        # Landscape A4 with narrow margins to accommodate wide tables
        section = doc.sections[0]
        section.orientation   = WD_ORIENT.LANDSCAPE
        section.page_width    = Cm(29.7)
        section.page_height   = Cm(21.0)
        section.left_margin   = Cm(1.5)
        section.right_margin  = Cm(1.5)
        section.top_margin    = Cm(1.5)
        section.bottom_margin = Cm(1.5)

        self._write_header(doc)

        # Group records by analysis_type
        groups: dict[str, list[dict]] = defaultdict(list)
        for r in self.records:
            groups[r.get("analysis_type", "UNKNOWN")].append(r)

        if self.combine_tph_voc and "SOIL_TPH" in groups and "SOIL_VOC" in groups:
            groups["SOIL_TPH_VOC"] = (
                list(groups.pop("SOIL_TPH")) + list(groups.pop("SOIL_VOC"))
            )

        if self.combine_tph_mbtex and "SOIL_TPH" in groups and "SOIL_MBTEX" in groups:
            groups["SOIL_TPH_MBTEX"] = (
                list(groups.pop("SOIL_TPH")) + list(groups.pop("SOIL_MBTEX"))
            )

        _ORDER = [
            "SOIL_TPH", "SOIL_TPH_VOC", "SOIL_TPH_MBTEX",
            "SOIL_METALS",
            "SOIL_VOC", "SOIL_MBTEX",
            "SOIL_PFAS",
            "GW_VOC", "GW_PFAS",
            "SOIL_GAS_VOC",
        ]
        ordered = [k for k in _ORDER if k in groups]
        ordered += [k for k in groups if k not in ordered and k != "LOWFLOW"]

        for atype in ordered:
            recs        = groups[atype]
            cfg         = SHEET_CONFIG.get(atype, {"name": atype, "unit": ""})
            thresh_keys = _ordered_thresh_keys(atype, self.sel_thresh)
            _add_section_heading(doc, cfg["name"])
            _build_analysis_table(doc, recs, self.tm, cfg, thresh_keys)

        if isinstance(self.out_path, (str, os.PathLike)):
            os.makedirs(os.path.dirname(str(self.out_path)) or ".", exist_ok=True)
        doc.save(self.out_path)
        return self.out_path

    # ------------------------------------------------------------------
    def _write_header(self, doc: Document) -> None:
        p = doc.add_paragraph()
        _set_rtl_para(p)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        text = "   |   ".join([
            f"שם פרויקט: {self.project}",
            f"תאריך: {self.rep_date}",
            f"מזמין: {self.client}",
        ])
        run = p.add_run(text)
        run.font.name = _FONT_HE
        run.font.size = Pt(11)
        run.font.bold = True
        doc.add_paragraph()
