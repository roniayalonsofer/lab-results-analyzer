"""
Groundwater Report Updater
==========================
Merges Bactochem PDF lab results into an existing Word monitoring report,
updates the historical data table, regenerates trend charts, and updates
the Mann-Kendall XLS file.

Public API:
    run_update_bytes(word_bytes, lab_pdf_bytes, mk_xls_bytes, field_pdf_bytes=None)
        -> (updated_word_bytes, updated_mk_xls_bytes)
"""

import copy
import re
import tempfile
from datetime import datetime
from io import BytesIO
from pathlib import Path

import numpy as np
import xlrd
import xlwt
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import pdfplumber
from docx import Document
from docx.oxml.ns import qn
from docx.shared import RGBColor
from docx.enum.text import WD_COLOR_INDEX

# ──────────────────────────────────────────────────────────────────────────────
# BACTOCHEM PDF PARSER
# ──────────────────────────────────────────────────────────────────────────────

CAS_TO_HEB = {
    "71-43-2":   "בנזן",
    "100-41-4":  "אתיל בנזן",
    "1634-04-4": "MTBE",
    "108-88-3":  "טולואן",
    "1330-20-7": "קסילן",
}


def parse_aminolab_pdf(pdf_path: str) -> dict:
    """
    Parse a single Aminolab groundwater monitoring certificate PDF.

    NOTE: unlike Bactochem, one Aminolab PDF certificate covers a SINGLE
    well/sample. Use parse_aminolab_pdfs() to merge several certificates
    (one per well) from one sampling round.

    Returns the same shape as parse_bactochem_pdf():
        {
          "sampling_date": "09.06.26",
          "samples": {
            "מת-1": {
              "date": "09.06.26",
              "results": {"בנזן": None, "MTBE": None, "טולואן": 0.035, ...},
              "field": {"pH": 6.6, "EC": 1388.0, "עומק פני המים": 54.25, ...},
              "water_level": 54.25,
              "total_depth": 57.15,
              "floating_layer": False,
            },
          }
        }
    """
    well_re = re.compile(r'^(\d+)\s*-\s*םוהת\s*ימ\s*:\s*המגודה\s*רואת$')
    date_re = re.compile(r'^(\d{2}/\d{2}/\d{4})\s*:םוגידה ךיראת$')
    ph_re   = re.compile(r'^-\s*([\d.]+)\s*-\s*pH הבגה$')
    ec_re   = re.compile(r'^-\s*([\d,]+)\s*[¥µu]S/cm תוכילומ$')
    temp_re = re.compile(r'^-\s*([\d.]+)\s*.C\s*הרוטרפמט$')
    do_re   = re.compile(r'^-\s*([\d.]+)\s*mg/L\s*DO\s*-\s*סמומ\s*ןצמח$')
    turb_re = re.compile(r'^-\s*([\d.]+)\s*NTU\s*תוריכע$')
    redox_re = re.compile(r'^-\s*([\d.]+)\s*mv\s*סקודר$')
    wl_re   = re.compile(r'^-\s*([\d.]+)\s*M\s+םימה ינפ קמוע\s*$')
    td_re   = re.compile(r'^-\s*([\d.]+)\s*M\s+חודיק קמוע\s*$')
    samp_d_re = re.compile(r'^-\s*([\d.]+)\s*M\s*חודיקב\s*םימה\s*ינפמ\s*םוגידה\s*קמוע$')
    floating_re = re.compile(r'הפצ הבכש')
    sampler_re = re.compile(r'(\S+)\s+(\S+)\s+-באלונימא\s+:י"ע\s+םגדנ$')

    # ── Inline BTEX format (e.g. "- 0.01 mg/L Benzene", "( -) 2 mg/L MTBE") ──
    # Used in sites where BTEX results appear directly in the field-measurements
    # table rather than as a separate multi-page VOC scan.
    INLINE_BTEX = {
        "Benzene":      "בנזן",
        "Toluene":      "טולואן",
        "Ethyl benzene": "אתיל בנזן",
        "Xylene":       "קסילן",
        "MTBE":         "MTBE",
    }
    inline_btex_re = re.compile(
        r'^[\s\d()+-]*(<?\d+\.?\d*)\s*mg/L\s+(' + '|'.join(re.escape(k) for k in INLINE_BTEX) + r')\s*$'
    )

    # Lines like: "* Toluene 108-88-3 35 84" / "Xylene's - 40 -" / "Benzene 71 -43 -2 <10 -"
    compound_re = re.compile(
        r'^\*?\s*(?P<name>.+?)\s*\*?\s+(?P<cas>[\d][\d\s\-]*\d|-)\s+'
        r'(?P<result><?\d+(?:\.\d+)?)\s+(?P<qual>-|\d+(?:\.\d+)?)\s*$'
    )

    well_id = None
    sample_date = None
    sampler_name = None
    field = {}
    water_level = None
    total_depth = None
    floating_layer = False
    results = {}

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for raw_line in text.splitlines():
                line = raw_line.strip()

                if well_id is None:
                    m = well_re.match(line)
                    if m:
                        well_id = f"מת-{m.group(1)}"
                        continue

                if sample_date is None:
                    m = date_re.match(line)
                    if m:
                        try:
                            dt = datetime.strptime(m.group(1), "%d/%m/%Y")
                            sample_date = dt.strftime("%d.%m.%y")
                        except ValueError:
                            pass
                        continue

                if sampler_name is None:
                    m = sampler_re.search(line)
                    if m:
                        sampler_name = f"{m.group(2)[::-1]} {m.group(1)[::-1]}"
                        continue

                m = ph_re.match(line)
                if m:
                    field["pH"] = float(m.group(1))
                    continue

                m = ec_re.match(line)
                if m:
                    field["EC"] = float(m.group(1).replace(",", ""))
                    continue

                m = temp_re.match(line)
                if m:
                    field["טמפרטורה"] = float(m.group(1))
                    continue

                m = do_re.match(line)
                if m:
                    field["חמצן מומס"] = float(m.group(1))
                    continue

                m = turb_re.match(line)
                if m:
                    field["עכירות"] = float(m.group(1))
                    continue

                m = redox_re.match(line)
                if m:
                    field["רדוקס"] = float(m.group(1))
                    continue

                m = wl_re.match(line)
                if m:
                    water_level = float(m.group(1))
                    field["עומק פני המים"] = water_level
                    continue

                m = td_re.match(line)
                if m:
                    total_depth = float(m.group(1))
                    field["עומק כללי של הקידוח"] = total_depth
                    continue

                m = samp_d_re.match(line)
                if m:
                    field["עומק דגימה מפני המים"] = float(m.group(1))
                    continue

                if floating_re.search(line):
                    floating_layer = True
                    continue

                # VOC compound rows (BTEX) — only on pages with "Compound"/"CAS No."
                m = compound_re.match(line)
                if m:
                    name = m.group("name").strip()
                    cas = re.sub(r'\s*-\s*', '-', m.group("cas").strip())
                    heb = CAS_TO_HEB.get(cas)
                    if not heb and "xylen" in name.lower():
                        heb = "קסילן"
                    if heb:
                        raw_result = m.group("result")
                        val = None if raw_result.startswith("<") else float(raw_result)
                        if val is not None:
                            val = val / 1000.0  # ppb -> mg/L
                        results[heb] = val

                # Inline BTEX format: "- 0.01 mg/L Benzene" / "( -) 2 mg/L MTBE"
                m = inline_btex_re.match(line)
                if m:
                    raw_result = m.group(1)
                    eng_name = m.group(2)
                    heb = INLINE_BTEX.get(eng_name)
                    if heb and heb not in results:  # don't overwrite scan-page values
                        val = None if raw_result.startswith("<") else float(raw_result)
                        results[heb] = val

    if well_id is None:
        well_id = "מת-?"

    sample = {
        "date": sample_date,
        "sampler_name": sampler_name,
        "results": results,
        "field": field,
        "water_level": water_level,
        "total_depth": total_depth,
        "floating_layer": floating_layer,
    }

    return {
        "sampling_date": sample_date,
        "samples": {well_id: sample},
    }


def parse_aminolab_pdfs(pdf_paths: list) -> dict:
    """
    Parse and merge several Aminolab certificates (one per well) from the
    same sampling round into the combined {"sampling_date", "samples"} shape.
    """
    combined_samples = {}
    sampling_date = None

    for pdf_path in pdf_paths:
        lab = parse_aminolab_pdf(pdf_path)
        if sampling_date is None:
            sampling_date = lab["sampling_date"]
        combined_samples.update(lab["samples"])

    return {"sampling_date": sampling_date, "samples": combined_samples}


def parse_bactochem_pdf(pdf_path: str) -> dict:
    """
    Parse a Bactochem groundwater monitoring PDF.

    Returns:
        {
          "sampling_date": "03.08.25",
          "samples": {
            "מת-1": {
              "date": "03.08.25",
              "results": {"בנזן": 0.014, "MTBE": 0.9, ...},  # None = ND
              "field": {"pH": 7.11, "EC": 6640, "עומק פני המים": 15.92, ...},
              "water_level": 15.92,
              "total_depth": 21.62,
              "floating_layer": False,
            },
          }
        }
    """
    samples = {}
    global_date = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines = text.splitlines()

            # Find all sample-description lines in this page
            # Pattern (RTL-mirrored): "1933818 :המגודה רפסמ 5-תמ הדוהי רוא לונוס-םוהת ימ םוגיד"
            segments = []
            for i, line in enumerate(lines):
                m = re.search(r'(\d+)-תמ', line)
                if m and ("הדוהי" in line or "לונוס" in line or "דלק" in line or "תמ" in line):
                    if "בלנק" not in line and "קיפ" not in line:
                        segments.append((f"מת-{m.group(1)}", i))
            segments.append((None, len(lines)))  # sentinel

            for seg_idx in range(len(segments) - 1):
                well_id, start = segments[seg_idx]
                _, end = segments[seg_idx + 1]
                seg_lines = lines[start:end]

                if well_id in samples:
                    continue  # already parsed (blank sample continuation)

                sample = {
                    "date": None,
                    "results": {},
                    "field": {},
                    "water_level": None,
                    "total_depth": None,
                    "floating_layer": False,
                }

                for line in seg_lines:
                    # Sampling date: "03/08/2025 07:20 :םוגיד דעומ"
                    m = re.search(r'(\d{2}/\d{2}/\d{4})\s+\d{2}:\d{2}', line)
                    if m and sample["date"] is None:
                        try:
                            dt = datetime.strptime(m.group(1), "%d/%m/%Y")
                            sample["date"] = dt.strftime("%d.%m.%y")
                            if global_date is None:
                                global_date = sample["date"]
                        except ValueError:
                            pass

                    # Analyte: "CAS #: 71-43-2 0.001 mg/L 0.720 MTBE"
                    m = re.search(
                        r'CAS #:\s*([\d\-]+)\s+[\d.]+\s+mg/L\s+([\d.]+|Not Detected)',
                        line, re.IGNORECASE
                    )
                    if m:
                        cas, raw_val = m.group(1), m.group(2)
                        heb = CAS_TO_HEB.get(cas)
                        if heb:
                            sample["results"][heb] = (
                                None if raw_val.lower() == "not detected"
                                else float(raw_val)
                            )

                    # Floating layer: "ןכ : הפצ הבכש תוחכונ"
                    if "הפצ הבכש" in line and "ןכ" in line:
                        sample["floating_layer"] = True

                    # Field: water level "M 15.59 ןוילע סלפמ קמוע"
                    m = re.search(r'M\s+([\d.]+)\s+ןוילע סלפמ קמוע', line)
                    if m:
                        v = float(m.group(1))
                        sample["water_level"] = v
                        sample["field"]["עומק פני המים"] = v

                    # EC: "µS/cm 6640 רחאל) תילמשח תוכילומ"
                    m = re.search(r'µS/cm\s+([\d,]+)', line)
                    if m:
                        sample["field"]["EC"] = float(m.group(1).replace(",", ""))

                    # pH: "pH units 7.11"
                    m = re.search(r'pH units\s+([\d.]+)', line)
                    if m:
                        sample["field"]["pH"] = float(m.group(1))

                    # DO: "mg/L 2.1 (תובצייתה רחאל) סמומ ןצמח"
                    m = re.search(r'mg/L\s+([\d.]+)\s+\(תובצייתה רחאל\) סמומ ןצמח', line)
                    if m:
                        sample["field"]["חמצן מומס"] = float(m.group(1))

                    # Redox: "mV -229 (תובצייתה רחאל) סקודר"
                    m = re.search(r'mV\s+(-?[\d.]+)\s+\(תובצייתה רחאל\) סקודר', line)
                    if m:
                        sample["field"]["רדוקס"] = float(m.group(1))

                    # Temp: "C° 27.9 (תובצייתה רחאל) הרוטרפמט"
                    m = re.search(r'C°\s+([\d.]+)\s+\(תובצייתה רחאל\) הרוטרפמט', line)
                    if m:
                        sample["field"]["טמפרטורה"] = float(m.group(1))

                    # Turbidity: "NTU 18 (תובצייתה רחאל) תוריכע"
                    m = re.search(r'NTU\s+([\d.]+)\s+\(תובצייתה רחאל\) תוריכע', line)
                    if m:
                        sample["field"]["עכירות"] = float(m.group(1))

                    # Total depth: "M 21.62 חודיקה לש יללכ קמוע"
                    m = re.search(r'M\s+([\d.]+)\s+חודיקה לש יללכ קמוע', line)
                    if m:
                        v = float(m.group(1))
                        sample["total_depth"] = v
                        sample["field"]["עומק כללי של הקידוח"] = v

                samples[well_id] = sample

    return {"sampling_date": global_date, "samples": samples}


# ──────────────────────────────────────────────────────────────────────────────
# MANN-KENDALL
# ──────────────────────────────────────────────────────────────────────────────

def _xl_date(xl_num):
    if not xl_num:
        return None
    try:
        return datetime.fromordinal(datetime(1899, 12, 30).toordinal() + int(xl_num))
    except Exception:
        return None


def _read_mann_kendall(xls_path: str) -> dict:
    wb = xlrd.open_workbook(xls_path)
    result = {}
    for sh_name in wb.sheet_names():
        if sh_name == "Kendall_Dist":
            continue
        ws = wb.sheet_by_name(sh_name)
        well_row = ws.row_values(13)
        wells = [str(w).strip() for w in well_row[3:] if str(w).strip()]
        events = []
        for r in range(14, ws.nrows):
            row = ws.row_values(r)
            if not row[1] or not row[2]:
                break
            dt = _xl_date(row[2])
            if dt is None:
                break
            values = {
                well: (float(row[3 + wi]) if row[3 + wi] != "" else None)
                for wi, well in enumerate(wells)
            }
            events.append((dt, values))
        result[sh_name] = {"wells": wells, "events": events}
    return result


def _update_mann_kendall(mk_data: dict, new_dt: datetime, new_values: dict) -> dict:
    for heb, sh in {"בנזן": "בנזן", "MTBE": "MTBE"}.items():
        if sh not in mk_data:
            continue
        sheet = mk_data[sh]
        well_vals = new_values.get(heb, {})
        sheet["events"].append((new_dt, {w: well_vals.get(w) for w in sheet["wells"]}))
    return mk_data


def _mk_s(values):
    clean = [v for v in values if v is not None]
    n = len(clean)
    if n < 4:
        return 0, n
    s = sum(
        1 if clean[j] > clean[i] else (-1 if clean[j] < clean[i] else 0)
        for i in range(n - 1)
        for j in range(i + 1, n)
    )
    return s, n


def _mk_conf(s, n):
    if n < 4:
        return 0.5
    var_s = n * (n - 1) * (2 * n + 5) / 18
    if var_s == 0:
        return 0.5
    z = (s - 1) / np.sqrt(var_s) if s > 0 else (s + 1) / np.sqrt(var_s) if s < 0 else 0
    try:
        from scipy.special import erfc
        return float(0.5 * erfc(-z / np.sqrt(2)))
    except ImportError:
        return 0.5


def _mk_trend(s, conf, cov):
    if conf >= 0.95:
        return "Increasing" if s > 0 else "Decreasing"
    if conf >= 0.90:
        return "Probably Increasing" if s > 0 else "Probably Decreasing"
    if s > 0:
        return "No Trend"
    return "Stable" if cov < 1.0 else "No Trend"


def _write_mann_kendall(mk_data: dict, orig_xls: str, out_path: str):
    orig_wb = xlrd.open_workbook(orig_xls, formatting_info=True)
    new_wb = xlwt.Workbook()
    for sh_name in orig_wb.sheet_names():
        orig_ws = orig_wb.sheet_by_name(sh_name)
        new_ws = new_wb.add_sheet(sh_name, cell_overwrite_ok=True)
        for r in range(orig_ws.nrows):
            for c in range(orig_ws.ncols):
                new_ws.write(r, c, orig_ws.cell_value(r, c))
        if sh_name == "Kendall_Dist" or sh_name not in mk_data:
            continue
        sheet = mk_data[sh_name]
        wells, events = sheet["wells"], sheet["events"]
        for ei, (dt, vals) in enumerate(events):
            row = 14 + ei
            new_ws.write(row, 1, ei + 1)
            new_ws.write(row, 2, (dt - datetime(1899, 12, 30)).days)
            for wi, w in enumerate(wells):
                v = vals.get(w)
                new_ws.write(row, 3 + wi, v if v is not None else "")
        # Update statistics rows 54-57
        for wi, w in enumerate(wells):
            col = 3 + wi
            all_v = [e[1].get(w) for e in events]
            numeric = [v for v in all_v if v is not None]
            cov = (np.std(numeric) / np.mean(numeric)
                   if numeric and np.mean(numeric) != 0 else 0)
            s, n = _mk_s(all_v)
            conf = _mk_conf(s, n)
            new_ws.write(54, col, round(cov, 4))
            new_ws.write(55, col, s)
            new_ws.write(56, col, round(conf, 4))
            new_ws.write(57, col, _mk_trend(s, conf, cov))
        new_ws.write(3, 2, (datetime.today() - datetime(1899, 12, 30)).days)
    new_wb.save(out_path)


# ──────────────────────────────────────────────────────────────────────────────
# CHART GENERATION
# ──────────────────────────────────────────────────────────────────────────────

PARAM_INFO = {
    "בנזן":      ("Benzene / בנזן",           "mg/L", 0.005),
    "MTBE":      ("MTBE",                      "mg/L", 0.04),
    "טולואן":    ("Toluene / טולואן",          "mg/L", 0.70),
    "אתיל בנזן": ("Ethyl Benzene / אתיל בנזן", "mg/L", 0.30),
    "קסילן":     ("Xylene / קסילן",            "mg/L", 0.50),
}
_COLORS = [
    "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728",
    "#9467bd", "#8c564b", "#e377c2", "#bcbd22",
]


def _generate_charts(mk_data: dict, out_dir: Path) -> dict:
    out_dir.mkdir(parents=True, exist_ok=True)
    paths = {}
    plt.rcParams.update({"font.family": "DejaVu Sans", "axes.unicode_minus": False})

    for sh_name, sheet in mk_data.items():
        wells, events = sheet["wells"], sheet["events"]
        if not events:
            continue
        label, unit, thresh = PARAM_INFO.get(sh_name, (sh_name, "mg/L", None))
        dates = [e[0] for e in events]

        fig, ax = plt.subplots(figsize=(10, 5))
        plotted = False
        for wi, well in enumerate(wells):
            vals = [e[1].get(well) for e in events]
            pairs = [(d, v) for d, v in zip(dates, vals) if v is not None]
            if not pairs:
                continue
            xs, ys = zip(*pairs)
            ax.plot(xs, ys, marker="o", label=well,
                    color=_COLORS[wi % len(_COLORS)], linewidth=1.5, markersize=5)
            plotted = True

        if not plotted:
            plt.close(fig)
            continue

        if thresh:
            ax.axhline(thresh, color="red", linestyle="--", linewidth=1.2,
                       label=f"Target ({thresh} {unit})")

        ax.xaxis.set_major_formatter(mdates.DateFormatter("%m/%y"))
        ax.xaxis.set_major_locator(mdates.AutoDateLocator())
        fig.autofmt_xdate()
        ax.set_xlabel("תאריך", fontsize=10)
        ax.set_ylabel(unit, fontsize=10)
        ax.set_title(label, fontsize=12, fontweight="bold")
        ax.legend(loc="upper right", fontsize=8, ncol=2)
        ax.grid(True, alpha=0.3)
        ax.text(0.01, 0.01,
                "ריכוזים קטנים מ-0.01 מג/ל מוצגים בתחתית הגרף",
                transform=ax.transAxes, fontsize=7, color="gray", va="bottom")
        plt.tight_layout()
        p = out_dir / f"chart_{sh_name}.png"
        plt.savefig(p, dpi=150, bbox_inches="tight")
        plt.close(fig)
        paths[sh_name] = str(p)

    return paths


# ──────────────────────────────────────────────────────────────────────────────
# WORD DOCUMENT UPDATER
# ──────────────────────────────────────────────────────────────────────────────

THRESH = {
    "בנזן": 0.005, "טולואן": 0.70, "אתיל בנזן": 0.30,
    "קסילן": 0.50, "MTBE": 0.04,
}
COL_PARAMS = ["MTBE", "בנזן", "טולואן", "אתיל בנזן", "קסילן"]


def _set_cell(cell, text, bold=False, color=None, highlight=False):
    """
    Replace cell text. Reuses the cell's existing first run (if any) so the
    original font/size formatting is preserved, instead of creating a brand
    new run with Word's default formatting.
    """
    para = cell.paragraphs[0]
    runs = para.runs
    if runs:
        keep = runs[0]
        keep.text = str(text)
        for extra in runs[1:]:
            extra._element.getparent().remove(extra._element)
    else:
        keep = para.add_run(str(text))

    # Remove any extra paragraphs beyond the first (keep cell single-paragraph)
    for p in cell.paragraphs[1:]:
        p._element.getparent().remove(p._element)

    if bold:
        keep.bold = True
    if color:
        keep.font.color.rgb = color
    if highlight:
        keep.font.highlight_color = WD_COLOR_INDEX.YELLOW


def _insert_row_after(table, after_idx: int, tmpl_idx: int):
    """Insert a deep-copied row after after_idx."""
    tmpl_tr = copy.deepcopy(table.rows[tmpl_idx]._tr)
    for t in tmpl_tr.findall(".//" + qn("w:t")):
        t.text = ""
    table.rows[after_idx]._tr.addnext(tmpl_tr)
    return table.rows[after_idx + 1]


def _last_row_for_well(table, well_id: str) -> int:
    """
    Find the last data row in this well's table. Matches on column 1
    holding a real DD.MM.YY date rather than column 0 (well name), since
    column 0 is often a vertically-merged cell that doesn't reliably read
    back per-row, and some rows may legitimately have a blank well-name
    cell while still holding real historical data.
    """
    date_pat = re.compile(r'^\d{2}\.\d{2}\.\d{2}$')
    last = -1
    for i, row in enumerate(table.rows):
        if len(row.cells) < 2:
            continue
        if date_pat.match(row.cells[1].text.strip()):
            last = i
    return last


def _update_field_table(doc: Document, field_data: dict, wells: list):
    """
    Update the field-findings table (doc.tables[1]).

    Two layouts are supported:
      - Single well: each data row is [param name, unit, value] — the value
        is always the LAST cell, no column matching needed.
      - Multiple wells: row 0 holds well names as column headers, used to
        find which column belongs to which well.
    """
    tbl = doc.tables[1]
    FIELD_ROWS = {
        "pH": 1, "EC": 2, "טמפרטורה": 3, "חמצן מומס": 4,
        "עכירות": 5, "רדוקס": 6, "עומק פני המים": 7,
        "עומק כללי של הקידוח": 8, "עומק דגימה מפני המים": 9,
    }

    if len(wells) == 1:
        well = wells[0]
        well_field = field_data.get(well, {}).get("field", {})
        # Detect rows dynamically by matching parameter name in col 0,
        # to handle docs where there's an extra "תאריך" row at index 1.
        name_to_row = {}
        for ri, row in enumerate(tbl.rows):
            row_name = row.cells[0].text.strip().rstrip("*")
            name_to_row[row_name] = ri

        FIELD_ROWS_ALT = {
            "pH":   "pH",
            "EC":   "EC",
            "טמפרטורה":  "טמפרטורה",
            "חמצן מומס": "חמצן מומס",
            "עכירות":    "עכירות",
            "רדוקס":     "רדוקס",
            "עומק פני המים":           "עומק פני המים",
            "עומק כללי של הקידוח":     "עומק כללי של הקידוח",
            "עומק דגימה מפני המים":    "עומק דגימה מפני המים",
        }

        for field, row_label in FIELD_ROWS_ALT.items():
            val = well_field.get(field)
            if val is None:
                continue
            ri = name_to_row.get(row_label)
            if ri is None:
                # fallback: hardcoded indices for the simple layout
                ri = {"pH": 1, "EC": 2, "טמפרטורה": 3, "חמצן מומס": 4,
                      "עכירות": 5, "רדוקס": 6, "עומק פני המים": 7,
                      "עומק כללי של הקידוח": 8, "עומק דגימה מפני המים": 9}.get(field)
            if ri is None or ri >= len(tbl.rows):
                continue
            txt = str(int(val)) if isinstance(val, float) and val == int(val) else str(val)
            row_cells = tbl.rows[ri].cells
            _set_cell(row_cells[-1], txt, highlight=True)
        return

    header = tbl.rows[0]
    well_col = {
        c.text.strip(): ci
        for ci, c in enumerate(header.cells)
        if c.text.strip() in wells
    }
    for field, ri in FIELD_ROWS.items():
        if ri >= len(tbl.rows):
            continue
        row = tbl.rows[ri]
        for well, ci in well_col.items():
            val = field_data.get(well, {}).get("field", {}).get(field)
            if val is not None:
                txt = str(int(val)) if isinstance(val, float) and val == int(val) else str(val)
                _set_cell(row.cells[ci], txt, highlight=True)


def _update_historical_tables(doc: Document, new_date: str,
                               samples: dict, wells_order: list):
    # Map well → table index
    well_table = {}
    for ti, tbl in enumerate(doc.tables[2:], 2):
        for row in tbl.rows:
            w = row.cells[0].text.strip()
            if w in wells_order:
                well_table[w] = ti

    for well in wells_order:
        if well not in samples:
            continue
        ti = well_table.get(well)
        if ti is None:
            continue
        tbl = doc.tables[ti]
        sample = samples[well]
        results = sample.get("results", {})
        floating = sample.get("floating_layer", False)

        last = _last_row_for_well(tbl, well)
        if last < 0:
            continue

        new_row = _insert_row_after(tbl, last, last)
        cells = new_row.cells
        # NOTE: cells[0] (well name) is often a vertically-merged cell shared
        # across the whole table — do NOT write to it here, or it will wipe
        # out the merge-origin text for every row in the table.
        _set_cell(cells[1], new_date, highlight=True)
        _set_cell(cells[2], "")

        if floating:
            _set_cell(cells[3], "לא נדגם עקב שכבה צפה", highlight=True)
            for ci in range(4, min(8, len(cells))):
                _set_cell(cells[ci], "")
        else:
            for ci, param in enumerate(COL_PARAMS):
                if 3 + ci >= len(cells):
                    break
                val = results.get(param)
                thresh = THRESH.get(param)
                if val is None:
                    txt, exceeds = "<0.001", False
                else:
                    exceeds = thresh is not None and val > thresh
                    txt = str(val)
                color = RGBColor(0xC0, 0x00, 0x00) if exceeds else None
                _set_cell(cells[3 + ci], txt, bold=exceeds, color=color, highlight=True)


def _detect_layout(doc: Document) -> str:
    """
    Detect document layout:
      'B' = has a full historical-field table (table[2] with 10+ cols, pH/EC headers)
      'A' = simple layout (table[2] is the chemistry table, no separate field history)
    """
    if len(doc.tables) >= 4:
        t2 = doc.tables[2]
        if len(t2.columns) >= 9:
            headers = [c.text.strip() for c in t2.rows[1].cells] if len(t2.rows) > 1 else []
            if any(h in ("pH", "EC", "pH units") for h in headers):
                return "B"
    return "A"


def _update_historical_field_table(doc: Document, new_date: str, field: dict, well_id: str):
    """
    Layout B: append a new row to the historical field-data table (doc.tables[2]).
    Columns: well, date, pH, EC, temp, DO, turbidity, redox, water_level, total_depth, samp_depth
    """
    tbl = doc.tables[2]
    COLS = ["pH", "EC", "טמפרטורה", "חמצן מומס", "עכירות", "רדוקס",
            "עומק פני המים", "עומק כללי של הקידוח", "עומק דגימה מפני המים"]

    # Find last data row (has a date in col 1)
    date_pat = re.compile(r'^\d{2}\.\d{2}\.\d{2,4}$')
    last = -1
    for i, row in enumerate(tbl.rows):
        if len(row.cells) >= 2 and date_pat.match(row.cells[1].text.strip()):
            last = i
    if last < 0:
        return

    new_row = _insert_row_after(tbl, last, last)
    cells = new_row.cells
    # col 0: well name (merged — don't touch)
    _set_cell(cells[1], new_date, highlight=True)
    for ci, key in enumerate(COLS, 2):
        val = field.get(key)
        if val is None:
            txt = "-"
        elif isinstance(val, float) and val == int(val):
            txt = str(int(val))
        else:
            txt = str(val)
        _set_cell(cells[ci], txt, highlight=True)


def _update_chem_table_layout_b(doc: Document, new_date: str, water_level, results: dict):
    """
    Layout B: append a new row to the chemistry table (doc.tables[3]).
    Columns: well, date, analysis, water_level, MTBE, בנזן, טולואן, אתיל בנזן, [כסילן]
    """
    tbl = doc.tables[3]

    CHEM_COLS = ["MTBE", "בנזן", "טולואן", "אתיל בנזן", "קסילן", "כסילן"]
    COL_ALIASES = {"כסילן": "קסילן"}  # doc uses כ, results dict uses ק
    CHEM_THRESH = {"MTBE": 0.02, "בנזן": 0.0025, "טולואן": 0.35, "אתיל בנזן": 0.15, "קסילן": 0.5}

    # Find column indices by exact header match, then fallback to contains
    header = [c.text.strip() for c in tbl.rows[0].cells]
    col_idx = {}
    for ci, h in enumerate(header):
        for param in CHEM_COLS:
            if h == param:  # exact match first
                canonical = COL_ALIASES.get(param, param)
                col_idx[canonical] = ci
    water_col = next((ci for ci, h in enumerate(header) if "עומק" in h and "מים" in h), 3)

    date_pat = re.compile(r'^\d{2}\.\d{2}\.\d{2,4}$')
    last = -1
    for i, row in enumerate(tbl.rows):
        if len(row.cells) >= 2 and date_pat.match(row.cells[1].text.strip()):
            last = i
    if last < 0:
        return

    new_row = _insert_row_after(tbl, last, last)
    cells = new_row.cells

    # col 0: well name (merged — don't touch)
    _set_cell(cells[1], new_date, highlight=True)
    _set_cell(cells[2], "", highlight=False)  # analysis type

    wl_txt = str(water_level) if water_level is not None else "-"
    _set_cell(cells[water_col], wl_txt, highlight=True)

    for param, ci in col_idx.items():
        val = results.get(param)
        thresh = CHEM_THRESH.get(param)
        if val is None:
            txt = "<0.001"
            exceeds = False
        else:
            txt = str(val)
            exceeds = thresh is not None and val > thresh
        color = RGBColor(0xC0, 0x00, 0x00) if exceeds else None
        _set_cell(cells[ci], txt, bold=exceeds, color=color, highlight=True)


def _update_mk_layout_b(mk_xls_bytes: bytes, new_dt: datetime, results: dict) -> bytes:
    """
    Layout B Mann-Kendall: add one row per sheet (MTBE, BENZEN/BENZENE).
    Sheet format: col B = event#, col C = xl-date-float, col D = concentration.
    """
    import tempfile as _tmp
    import xlrd, xlwt
    from xlutils.copy import copy as xl_copy

    SHEET_TO_PARAM = {
        "MTBE": "MTBE", "mtbe": "MTBE",
        "BENZEN": "בנזן", "benzen": "בנזן", "BENZENE": "בנזן", "benzene": "בנזן",
        "TOLUENE": "טולואן", "toluene": "טולואן",
        "XYLENE": "קסילן", "xylene": "קסילן",
        "ETHYLBENZENE": "אתיל בנזן",
    }

    with _tmp.NamedTemporaryFile(suffix=".xls", delete=False) as f:
        f.write(mk_xls_bytes)
        tmp_in = f.name

    import xlwt, xlrd
    rb = xlrd.open_workbook(tmp_in, formatting_info=True)
    wb = xl_copy(rb)
    date_mode = rb.datemode

    import datetime as _dt
    xl_date = (_dt.datetime(new_dt.year, new_dt.month, new_dt.day)
               - _dt.datetime(1899, 12, 30)).days

    date_style = xlwt.XFStyle()
    date_style.num_format_str = 'M/D/YY'
    yellow_style = xlwt.XFStyle()
    yellow_style.pattern = xlwt.Pattern()
    yellow_style.pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    yellow_style.pattern.pattern_fore_colour = 13  # yellow

    for si in range(rb.nsheets):
        rs = rb.sheet_by_index(si)
        ws = wb.get_sheet(si)
        sname = rs.name.strip().upper()
        param_heb = SHEET_TO_PARAM.get(sname) or SHEET_TO_PARAM.get(rs.name.strip())
        if param_heb is None:
            continue

        # Find last data row (col B = event#, numeric)
        last_r = 13
        for r in range(14, rs.nrows):
            b = rs.cell(r, 1).value
            if b and isinstance(b, float):
                last_r = r
            elif b == '':
                break

        new_event = int(rs.cell(last_r, 1).value) + 1
        val = results.get(param_heb)
        if val is None:
            val = 0.001

        nr = last_r + 1
        ws.write(nr, 1, new_event, yellow_style)
        ws.write(nr, 2, xl_date, date_style)
        ws.write(nr, 3, val, yellow_style)

    with _tmp.NamedTemporaryFile(suffix=".xls", delete=False) as f:
        out_path = f.name
    wb.save(out_path)
    data = open(out_path, "rb").read()
    import os
    os.unlink(tmp_in)
    os.unlink(out_path)
    return data


HEBREW_MONTHS = ["ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני",
                  "יולי", "אוגוסט", "ספטמבר", "אוקטובר", "נובמבר", "דצמבר"]


def _update_narrative_layout_b(doc: Document, date_str: str, results: dict):
    """
    Layout B: update the date and BTEX concentration values in the
    narrative sentences.
    - Date: appears as a separate run after "בתאריך " (run index 1)
    - BTEX: fixed value-run positions (6,16,26,32,35) hold the previous
      round's numeric values which need to be overwritten with new ones.
      Order in text: MTBE, Benzene, Ethylbenzene, Toluene, Xylene.
    """
    ORDERED = ["MTBE", "בנזן", "אתיל בנזן", "טולואן", "קסילן"]

    for p in doc.paragraphs:
        runs = p.runs
        text = p.text

        # Update "בתאריך [date] נערך" sentences — date is in run immediately
        # after the "בתאריך" run (run 1 in afr layout)
        if "נערך" in text and "בתאריך" in text:
            for i, run in enumerate(runs):
                if run.text.rstrip() == "בתאריך" or run.text.rstrip().endswith("בתאריך"):
                    # The next run holds the old date value
                    if i + 1 < len(runs):
                        runs[i + 1].text = date_str + " "
                        runs[i + 1].font.highlight_color = WD_COLOR_INDEX.YELLOW
                    else:
                        run.text = run.text.rstrip() + " " + date_str + " "
                        run.font.highlight_color = WD_COLOR_INDEX.YELLOW

        # Update the BTEX concentration sentence
        # The value runs are the numeric/non-Hebrew runs between the fixed
        # label phrases. Identify them by: they follow a "של " or ", " run
        # and precede a "מ\"ג" run.
        if "בדיגום" in text and "MTBE" in text and "בריכוזים" in text:
            value_run_indices = []
            for i, run in enumerate(runs):
                t = run.text.strip()
                # A numeric value run: contains only digits, dots, "<"
                if re.match(r'^<?[\d.]+$', t) and i + 1 < len(runs):
                    nxt = runs[i + 1].text
                    if 'מ"ג' in nxt or "מ\"ג" in nxt or 'מ' in nxt:
                        value_run_indices.append(i)

            for slot_idx, param in zip(value_run_indices, ORDERED):
                val = results.get(param)
                txt = str(val) if val is not None else "<0.001"
                runs[slot_idx].text = txt
                runs[slot_idx].font.highlight_color = WD_COLOR_INDEX.YELLOW


def _heb_join_words(items):
    if len(items) == 1:
        return items[0]
    return ", ".join(items[:-1]) + f" ו{items[-1]}"


def _heb_join_values(items):
    if len(items) == 1:
        return items[0]
    return ", ".join(items[:-1]) + f" ו-{items[-1]}"


def _update_title_month_year(doc: Document, new_dt: datetime) -> bool:
    """Update the report title's month + year (e.g. 'פברואר' -> 'יוני 2026')."""
    month_name = HEBREW_MONTHS[new_dt.month - 1]
    year = new_dt.year
    updated = False
    for p in doc.paragraphs:
        if "ניטור מי תהום תקופתי" not in p.text:
            continue
        runs = p.runs
        for i, run in enumerate(runs):
            if run.text.strip() in HEBREW_MONTHS:
                run.text = f" {month_name}"
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                if i + 1 < len(runs):
                    nxt = runs[i + 1]
                    if nxt.text.strip() == "" or re.match(r'^\d{4}$', nxt.text.strip()):
                        nxt.text = f" {year}"
                        nxt.font.highlight_color = WD_COLOR_INDEX.YELLOW
                updated = True
    return updated


def _update_narrative_placeholders(doc: Document, date_str: str, sampler_name: str):
    """
    Replace the date and sampler-name VALUE runs in the CURRENT-ROUND
    sampling sentences only (identified by the anchor phrase 'נערך דיגום').
    This must NOT touch unrelated mentions of dates/'מר X' elsewhere in the
    document (e.g. historical references in the background section).
    """
    for p in doc.paragraphs:
        if "נערך דיגום" not in p.text:
            continue
        runs = p.runs
        for i, run in enumerate(runs):
            if i > 0 and runs[i - 1].text.rstrip().endswith("בתאריך"):
                run.text = date_str
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            if sampler_name and run.text.strip() == "מר" and i + 1 < len(runs):
                runs[i + 1].text = sampler_name
                runs[i + 1].font.highlight_color = WD_COLOR_INDEX.YELLOW


def _update_cover_page_date(doc: Document, new_dt: datetime) -> bool:
    """Update the standalone cover-page '<month> <year>' line."""
    month_name = HEBREW_MONTHS[new_dt.month - 1]
    year = new_dt.year
    pat = re.compile(r'^(' + "|".join(HEBREW_MONTHS) + r')\s+\d{4}$')
    for p in doc.paragraphs:
        if not pat.match(p.text.strip()):
            continue
        runs = p.runs
        if not runs:
            continue
        runs[0].text = f"{month_name} {year}"
        for extra in runs[1:]:
            extra.text = ""
        runs[0].font.highlight_color = WD_COLOR_INDEX.YELLOW
        return True
    return False


def _update_concentration_summary(doc: Document, results: dict):
    """
    Rebuild the BTEX concentration summary sentence
    ('בדיגום שנערך אותרו X, Y בריכוזים של ... בהתאמה') based on which
    compounds were actually detected this round.
    """
    ordered = [("MTBE", results.get("MTBE")), ("בנזן", results.get("בנזן")),
               ("טולואן", results.get("טולואן")), ("אתיל בנזן", results.get("אתיל בנזן")),
               ("קסילן", results.get("קסילן"))]
    detected = [(n, v) for n, v in ordered if v is not None]

    for p in doc.paragraphs:
        if "בדיגום שנערך" not in p.text:
            continue
        runs = p.runs
        start_idx = end_idx = None
        for i, r in enumerate(runs):
            if r.text.strip() == "בדיגום שנערך":
                start_idx = i + 1
            elif start_idx is not None and "בהתאמה" in r.text:
                end_idx = i
                break
        if start_idx is None or end_idx is None:
            continue

        if not detected:
            sentence = "לא אותרו מרכיבי BTEX בריכוזים מעל סף הזיהוי"
        elif len(detected) == 1:
            name, val = detected[0]
            sentence = f'אותר {name} בריכוז של {val} מ"ג/ליטר'
        else:
            names = _heb_join_words([n for n, _ in detected])
            vals = _heb_join_values([f'{v} מ"ג/ליטר' for _, v in detected])
            sentence = f'אותרו {names} בריכוזים של {vals} בהתאמה'

        runs[start_idx].text = sentence
        runs[start_idx].font.highlight_color = WD_COLOR_INDEX.YELLOW
        for j in range(start_idx + 1, end_idx + 1):
            runs[j].text = ""
        return


def _replace_chart_image(doc: Document, shape_idx: int, img_path: str):
    if shape_idx >= len(doc.inline_shapes):
        return
    shape = doc.inline_shapes[shape_idx]
    w_emu, h_emu = shape.width, shape.height
    blip = shape._inline.find(".//" + qn("a:blip"))
    if blip is None:
        return
    rId = blip.get(qn("r:embed"))
    image_part = doc.part.related_parts[rId]
    with open(img_path, "rb") as f:
        image_part._blob = f.read()
    extent = shape._inline.find(qn("wp:extent"))
    if extent is not None:
        extent.set("cx", str(w_emu))
        extent.set("cy", str(h_emu))


# ──────────────────────────────────────────────────────────────────────────────
# PUBLIC API
# ──────────────────────────────────────────────────────────────────────────────

def run_update_bytes(
    word_bytes: bytes,
    lab_pdf_bytes,
    mk_xls_bytes: bytes = None,
    field_pdf_bytes: bytes = None,   # reserved for future OCR use
    lab_type: str = "bactochem",     # "bactochem" or "aminolab"
) -> tuple:
    """
    Process all inputs in memory and return (updated_word_bytes, updated_mk_xls_bytes).
    updated_mk_xls_bytes is None if mk_xls_bytes was not provided.

    lab_pdf_bytes:
      - "bactochem": a single PDF's bytes (one PDF covers all wells).
      - "aminolab":  bytes of ONE PDF, or a list/tuple of bytes — one per
        well/certificate — since each Aminolab certificate covers a single well.
    """
    import os

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)

        # Write inputs to temp files
        word_path = tmp / "report.docx"
        out_word  = tmp / "updated_report.docx"
        out_mk    = tmp / "updated_mk.xls"

        word_path.write_bytes(word_bytes)

        has_mk = mk_xls_bytes is not None
        if has_mk:
            mk_path = tmp / "mk.xls"
            mk_path.write_bytes(mk_xls_bytes)

        # ── Parse lab results ──────────────────────────────────────
        if lab_type == "aminolab":
            pdf_bytes_list = (
                lab_pdf_bytes if isinstance(lab_pdf_bytes, (list, tuple))
                else [lab_pdf_bytes]
            )
            lab_paths = []
            for i, pdf_bytes in enumerate(pdf_bytes_list):
                p = tmp / f"lab_{i}.pdf"
                p.write_bytes(pdf_bytes)
                lab_paths.append(str(p))
            lab = parse_aminolab_pdfs(lab_paths)
            samples, new_date = lab["samples"], lab["sampling_date"]
            if not new_date:
                raise ValueError("לא ניתן לחלץ תאריך דיגום מה-PDF. ודא שהקובץ הוא תעודת אמינולאב תקנית.")
            if not samples:
                raise ValueError("לא נמצאו דגימות ב-PDF. ודא שהקובץ הוא תעודת אמינולאב לדיגום מי תהום.")
        else:
            lab_path = tmp / "lab.pdf"
            lab_path.write_bytes(lab_pdf_bytes)
            lab = parse_bactochem_pdf(str(lab_path))
            samples, new_date = lab["samples"], lab["sampling_date"]
            if not new_date:
                raise ValueError("לא ניתן לחלץ תאריך דיגום מה-PDF. ודא שהקובץ הוא דוח בקטוכם תקני.")
            if not samples:
                raise ValueError("לא נמצאו דגימות ב-PDF. ודא שהקובץ הוא דוח בקטוכם לדיגום מי תהום.")

        new_dt = datetime.strptime(new_date, "%d.%m.%y")
        chart_paths = {}
        out_mk_bytes = None

        if has_mk:
            # ── Update Mann-Kendall ────────────────────────────────────
            mk_data = _read_mann_kendall(str(mk_path))
            mk_new = {}
            for well, s in samples.items():
                for param, val in s.get("results", {}).items():
                    mk_new.setdefault(param, {})[well] = val
            mk_data = _update_mann_kendall(mk_data, new_dt, mk_new)

            # ── Generate charts ────────────────────────────────────────
            chart_paths = _generate_charts(mk_data, tmp / "charts")

            # ── Write updated MK XLS ──────────────────────────────────
            _write_mann_kendall(mk_data, str(mk_path), str(out_mk))
            out_mk_bytes = out_mk.read_bytes()

        # ── Update Word document ──────────────────────────────────
        doc = Document(str(word_path))
        layout = _detect_layout(doc)

        # Detect the well-naming convention actually used in THIS document
        # (different sites use different prefixes, e.g. "מת-1" or "צא - 1"),
        # then remap the parsed sample keys (always "מת-N") onto it by well number.
        doc_wells = set()
        for tbl in doc.tables:
            for row in tbl.rows:
                if not row.cells:
                    continue
                w = row.cells[0].text.strip()
                if w and re.search(r'\d', w) and len(w) < 20:
                    doc_wells.add(w)

        num_to_doc_well = {}
        for w in doc_wells:
            m = re.search(r'(\d+)\s*$', w)
            if m:
                num_to_doc_well.setdefault(m.group(1), w)

        remapped_samples = {}
        unmatched = []
        for well_key, sample in samples.items():
            m = re.search(r'(\d+)$', well_key)
            num = m.group(1) if m else None
            doc_well = num_to_doc_well.get(num)
            if doc_well:
                remapped_samples[doc_well] = sample
            else:
                remapped_samples[well_key] = sample
                unmatched.append(well_key)
        samples = remapped_samples
        wells_order = list(samples.keys())

        if unmatched:
            raise ValueError(
                "לא נמצאה בדוח ה-Word התאמה לבארות הבאות מה-PDF: "
                + ", ".join(unmatched)
                + ". ודא שמספרי הבארות בדוח ה-PDF תואמים לבארות בדוח ה-Word."
            )

        _update_field_table(doc, samples, wells_order)

        if layout == "B":
            # Layout B: historical field table (table 2) + chemistry table (table 3)
            first_sample = next(iter(samples.values()), {})
            _update_historical_field_table(doc, new_date, first_sample.get("field", {}),
                                           wells_order[0] if wells_order else "")
            _update_chem_table_layout_b(doc, new_date,
                                        first_sample.get("water_level"),
                                        first_sample.get("results", {}))
            # MK update for Layout B uses a different sheet structure
            if has_mk:
                out_mk_bytes = _update_mk_layout_b(
                    mk_xls_bytes, new_dt, first_sample.get("results", {})
                )
        else:
            _update_historical_tables(doc, new_date, samples, wells_order)

        # ── Update narrative text (title month/year, date/sampler gaps,
        #    BTEX concentration summary sentence) ────────────────────────
        full_date_str = new_dt.strftime("%d.%m.%Y")
        first_sample = next(iter(samples.values()), {})
        sampler_name = first_sample.get("sampler_name")
        combined_results = {}
        for s in samples.values():
            combined_results.update(s.get("results", {}))

        _update_title_month_year(doc, new_dt)
        _update_cover_page_date(doc, new_dt)
        if layout == "B":
            _update_narrative_layout_b(doc, full_date_str, combined_results)
        else:
            _update_narrative_placeholders(doc, full_date_str, sampler_name)
            _update_concentration_summary(doc, combined_results)

        # Replace chart images (shape 1 = Benzene, shape 2 = MTBE)
        for chart_idx, sh_name in enumerate(["בנזן", "MTBE"]):
            if sh_name in chart_paths:
                _replace_chart_image(doc, 1 + chart_idx, chart_paths[sh_name])

        doc.save(str(out_word))

        return out_word.read_bytes(), out_mk_bytes
