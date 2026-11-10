"""
Groundwater Report Updater
==========================
Merges one sampling round (lab results + field form) into an existing Word
monitoring report: the field tables, the history tables, the per-well
result tables, the narrative, and the Mann-Kendall pages.

Public API:
    run_update_bytes(word_bytes, lab_pdf_bytes, mk_xls_bytes=None, ...)
        -> (updated_word_bytes, mk_paste_rows_xlsx_bytes | None)

The Mann-Kendall .xls is only read, never written (it holds formulas and
charts); see core/gw_mann_kendall.py.
"""

import copy
import re
import tempfile
from datetime import datetime
from io import BytesIO
from pathlib import Path

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

    # Concatenate lines from ALL pages before segment detection. A sample's
    # description line can appear near the bottom of one page while its
    # CAS-result rows spill onto the next page (the results table simply
    # continues) — processing page-by-page would silently drop those rows.
    lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines.extend(text.splitlines())

    # Find all sample-description lines in the document.
    # Anchored on the "מספר הדוגמה" label (RTL-mirrored: "המגודה רפסמ"),
    # which appears on every sample-description line regardless of site.
    # Well-code token appears right after that label in one of two forms:
    #   - hyphenated, no space (e.g. "5-תמ" mirrored == "מת-5")
    #   - space-separated (e.g. "3 פפ" mirrored == "פפ-3" / "פפ 3")
    # An optional apostrophe/geresh (' or ׳) can appear between the dash
    # and the letters (e.g. "1-'תמ" mirrored == "מת'-1") depending on how
    # the lab formats the well code — tolerate it without changing the
    # normalized well_id we emit ("מת-<n>").
    segments = []
    for i, line in enumerate(lines):
        if "רפסמ" not in line or "המגודה" not in line:
            continue
        if "בלנק" in line or "קיפ" in line:
            continue
        m = re.search(r"(\d+)-['\u05f3]?([א-ת]{1,3})(?!\S)", line)
        if not m:
            m = re.search(r"(\d+)\s+['\u05f3]?([א-ת]{1,3})(?:\s|$|-)", line)
        if m:
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

            # Pumping depth: "M 12.87 הביאש קמוע"
            m = re.search(r'M\s+([\d.]+)\s+הביאש קמוע', line)
            if m:
                sample["field"]["עומק שאיבה"] = float(m.group(1))

        samples[well_id] = sample

    return {"sampling_date": global_date, "samples": samples}


# ──────────────────────────────────────────────────────────────────────────────
# WORD DOCUMENT UPDATER
# ──────────────────────────────────────────────────────────────────────────────

THRESH = {
    "בנזן": 0.005, "טולואן": 0.70, "אתיל בנזן": 0.30,
    "קסילן": 0.50, "MTBE": 0.04,
}
COL_PARAMS = ["MTBE", "בנזן", "טולואן", "אתיל בנזן", "קסילן"]

# Different site templates order these columns differently (some put MTBE
# first, others last) and even spell Xylene two different ways ("קסילן" vs
# "כסילן"). Map every spelling variant seen so far to one canonical name —
# the same canonical names used as keys in `results` dicts from the PDF
# parsers and in THRESH above.
_PARAM_CANONICAL = {
    "MTBE": "MTBE",
    "בנזן": "בנזן",
    "טולואן": "טולואן",
    "אתיל בנזן": "אתיל בנזן",
    "קסילן": "קסילן",
    "כסילן": "קסילן",  # alternate spelling for Xylene used by some templates
}


def _detect_chem_table_columns(tbl) -> dict:
    """
    Read a chemistry-history table's own header block to build a
    {column_index: canonical_param_name} map, instead of assuming a fixed
    column order starting at COL_PARAMS[0]. Checks the first few rows (the
    header block, which may span 2-3 rows for name/units/standard) and
    picks whichever row has the most recognizable chemical-name cells.
    Returns {} if no header row with recognizable names is found, so
    callers can fall back to the fixed COL_PARAMS order for older/simpler
    documents.
    """
    best_row, best_hits = None, 0
    for row in tbl.rows[:4]:
        hits = sum(1 for c in row.cells if c.text.strip() in _PARAM_CANONICAL)
        if hits > best_hits:
            best_hits, best_row = hits, row
    if best_row is None or best_hits == 0:
        return {}
    col_map = {}
    for ci, c in enumerate(best_row.cells):
        label = c.text.strip()
        if label in _PARAM_CANONICAL:
            col_map[ci] = _PARAM_CANONICAL[label]
    return col_map


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

    # Explicitly set bold/color state — not just when turning them ON.
    # `keep` is often a run inherited from a deep-copied template row (new
    # historical-table rows are cloned from the previous round's row), so
    # it can already carry bold/red formatting from an earlier exceedance.
    # Only ever setting `bold = True` / a color and never clearing them
    # means that formatting would persist forever on every later row, even
    # once a value drops back under threshold.
    keep.bold = bool(bold)
    if color:
        keep.font.color.rgb = color
    else:
        rPr = keep._element.get_or_add_rPr()
        color_elm = rPr.find(qn("w:color"))
        if color_elm is not None:
            rPr.remove(color_elm)
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
    Find the last data row belonging to `well_id` in this table.

    Some tables hold a single well (column 0 vertically merged across the
    whole table), others hold multiple wells back-to-back (e.g. a shared
    table for both "פפ-2" and "פפ-3", each with its own vertical-merge
    block in column 0). python-docx resolves a vertical-merge continuation
    cell's .text to the merge origin's text, so tracking the most recently
    seen non-empty column-0 value as we scan down correctly attributes each
    row to its well, even across merge boundaries.

    Falls back to "last dated row in the whole table" only if no row's
    column-0 text ever matches well_id (e.g. truly blank/unreadable column),
    to preserve prior behavior for documents where that heuristic fails.
    """
    date_pat = re.compile(r'^\d{2}\.\d{2}\.\d{2}$')
    last_any = -1
    last_for_well = -1
    current_well = None
    for i, row in enumerate(table.rows):
        if len(row.cells) < 2:
            continue
        w = row.cells[0].text.strip()
        if w:
            current_well = w
        if date_pat.match(row.cells[1].text.strip()):
            last_any = i
            if current_well == well_id:
                last_for_well = i
    return last_for_well if last_for_well >= 0 else last_any


_FIELD_PARAM_HINTS = {
    "pH", "EC", "טמפרטורה", "חמצן מומס", "עכירות", "רדוקס",
    "מוליכות חשמלית", "עומק פני המים", "עומק דיגום", "מפלס עליון",
    "עומק שאיבה", "עומק כללי של הקידוח", "עומק כללי קידוח",
    "תאריך", "תאריך דיגום", "עומק דגימה מפני המים",
}


def _find_field_table(doc: Document) -> int:
    """
    Locate the field-findings table by content signature (how many of its
    row-0-cell labels match known field-parameter names), rather than
    assuming it's always doc.tables[1]. Some documents have extra tables
    (signature blocks, timestamp tables) before it, which shifts its real
    index.
    """
    best_ti, best_score = None, 0
    for ti, tbl in enumerate(doc.tables):
        score = 0
        for row in tbl.rows:
            if not row.cells:
                continue
            label = row.cells[0].text.strip().rstrip("*")
            if label in _FIELD_PARAM_HINTS:
                score += 1
        if score > best_score:
            best_score, best_ti = score, ti
    return best_ti if best_score >= 3 else 1  # fallback to legacy default


_FIELD_ROW_ALIASES = {
    "תאריך דיגום": "__date__",
    "תאריך": "__date__",
    "עומק דיגום": "עומק פני המים",
    "מפלס עליון": "עומק פני המים",
    "עומק פני המים": "עומק פני המים",
    "חמצן מומס": "חמצן מומס",
    "מוליכות חשמלית": "EC",
    "EC": "EC",
    "pH": "pH",
    "רדוקס": "רדוקס",
    "עומק שאיבה": "עומק שאיבה",
    "טמפרטורה": "טמפרטורה",
    "עומק כללי קידוח": "עומק כללי של הקידוח",
    "עומק כללי של הקידוח": "עומק כללי של הקידוח",
    "עומק דגימה מפני המים": "עומק דגימה מפני המים",
    "עכירות": "עכירות",
}


def _fmt_field_val(val) -> str:
    if isinstance(val, float) and val == int(val):
        return str(int(val))
    return str(val)


def _update_field_table_dual_column(tbl, field_data: dict, wells: list):
    """
    Some sites' field-findings table shows each well across TWO adjacent
    columns — the two most recent sampling dates side by side (older on
    the left, newer on the right) — instead of one current-value column
    per well. To add a new reading: shift the newer column's values into
    the older column's slot (dropping the oldest reading), then write the
    new reading into the (now-vacated) newer column.
    """
    header_cells = tbl.rows[0].cells
    well_cols: dict = {}
    for ci, c in enumerate(header_cells):
        text = c.text.strip()
        if text in wells:
            well_cols.setdefault(text, []).append(ci)

    for well, cols in well_cols.items():
        if len(cols) < 2:
            continue
        old_col, new_col = sorted(cols)[-2:]
        sample = field_data.get(well, {})
        well_field = sample.get("field", {})
        sample_date = sample.get("date")

        for ri, row in enumerate(tbl.rows):
            if ri == 0:
                continue
            cells = row.cells
            if new_col >= len(cells) or old_col >= len(cells):
                continue
            label = cells[0].text.strip().rstrip("*")
            key = _FIELD_ROW_ALIASES.get(label)

            # Shift the previously-newest reading into the older slot
            shifted_text = cells[new_col].text.strip()
            _set_cell(cells[old_col], shifted_text)

            if key == "__date__":
                if sample_date:
                    _set_cell(cells[new_col], sample_date, highlight=True)
                continue
            if key is None:
                _set_cell(cells[new_col], "-", highlight=True)
                continue  # unrecognized/unparsed row — no source data available
            val = well_field.get(key)
            _set_cell(cells[new_col], _fmt_field_val(val) if val is not None else "-",
                      highlight=True)


def _update_field_table(doc: Document, field_data: dict, wells: list):
    """
    Update the field-findings table.

    Layouts supported:
      - Single well: each data row is [param name, unit, value] — the value
        is always the LAST cell, no column matching needed.
      - Multiple wells, one column each: row 0 holds well names as column
        headers, used to find which column belongs to which well.
      - Multiple wells, two columns each (rolling last-2-dates view): see
        _update_field_table_dual_column.
    """
    ti = _find_field_table(doc)
    tbl = doc.tables[ti]
    FIELD_ROWS = {
        "pH": 1, "EC": 2, "טמפרטורה": 3, "חמצן מומס": 4,
        "עכירות": 5, "רדוקס": 6, "עומק פני המים": 7,
        "עומק כללי של הקידוח": 8, "עומק דגימה מפני המים": 9,
    }

    header_texts = [c.text.strip() for c in tbl.rows[0].cells] if tbl.rows else []
    well_counts: dict = {}
    for h in header_texts:
        if h in wells:
            well_counts[h] = well_counts.get(h, 0) + 1
    if any(cnt >= 2 for cnt in well_counts.values()):
        _update_field_table_dual_column(tbl, field_data, wells)
        return

    if len(wells) == 1:
        well = wells[0]
        well_field = field_data.get(well, {}).get("field", {})
        sample_date = field_data.get(well, {}).get("date")

        # Detect rows dynamically by matching parameter name in col 0,
        # to handle docs where there's an extra "תאריך" row at index 1.
        name_to_row = {}
        for ri, row in enumerate(tbl.rows):
            row_name = row.cells[0].text.strip().rstrip("*")
            name_to_row[row_name] = ri

        # Update the date row if present
        if sample_date and "תאריך" in name_to_row:
            ri = name_to_row["תאריך"]
            _set_cell(tbl.rows[ri].cells[-1], sample_date, highlight=True)

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
            txt = _fmt_field_val(val)
            row_cells = tbl.rows[ri].cells
            _set_cell(row_cells[-1], txt, highlight=True)
        return

    # Column headers may spell the well differently from the history tables
    # ("חק 2" vs "חק-2"), so match them by well number.
    header = tbl.rows[0]
    num_to_well = {_well_num(w): w for w in wells}
    well_col = {}
    for ci, c in enumerate(header.cells):
        w = num_to_well.get(_well_num(c.text)) if ci > 0 else None
        if w and w not in well_col:
            well_col[w] = ci
    # Look up each row by its actual label (column 0), the same way the
    # single-well branch above already does — NOT by a fixed row index.
    # Different documents order these rows differently (e.g. some list
    # "טמפרטורה" before "pH"/"EC", others after), so a hardcoded index
    # silently writes each value into the wrong row on documents whose
    # order doesn't match the assumption.
    name_to_row = {}
    for ri, row in enumerate(tbl.rows):
        row_label = row.cells[0].text.strip().rstrip("*")
        name_to_row[row_label] = ri

    for field in FIELD_ROWS:
        ri = name_to_row.get(field)
        if ri is None or ri >= len(tbl.rows):
            continue
        row = tbl.rows[ri]
        for well, ci in well_col.items():
            val = field_data.get(well, {}).get("field", {}).get(field)
            if val is not None:
                txt = _fmt_like(row.cells[ci].text, val)
                _set_cell(row.cells[ci], txt, highlight=True)


def _extract_target_thresholds(tbl, col_map: dict = None) -> dict:
    """
    Read this table's own site-specific threshold row — either a
    "ריכוזי יעד לסיום ניטור" (monitoring end-target, set per-site by the
    regulator) or a "תקן מי שתייה" (drinking-water standard) row,
    whichever this particular template uses — and return
    {canonical_param: threshold}.

    This is what should actually decide the bold/red exceedance marking.
    A single hardcoded threshold is wrong in general: different sites use
    different standards, and a value can be well under THIS site's own
    threshold while looking like an exceedance under a generic one (or
    vice-versa). Returns {} if no such row is found, so callers can fall
    back to a generic default.
    """
    if col_map is None:
        col_map = _detect_chem_table_columns(tbl)
    if not col_map:
        col_map = {3 + i: p for i, p in enumerate(COL_PARAMS)}

    for row in tbl.rows:
        cells = row.cells
        if len(cells) < 3:
            continue
        label = cells[2].text.strip()
        is_target_row = ("יעד" in label and ("ניטור" in label or "סיום" in label)) or "תקן" in label
        if not is_target_row:
            continue
        out = {}
        for ci, param in col_map.items():
            if ci >= len(cells):
                continue
            txt = cells[ci].text.strip()
            try:
                out[param] = float(txt)
            except ValueError:
                continue
        if out:
            return out
    return {}


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

        # Column order for MTBE/BTEX varies by site template (some put
        # MTBE first, others last) and Xylene is spelled two different
        # ways ("קסילן"/"כסילן") — read the table's own header instead of
        # assuming COL_PARAMS' fixed order. Falls back to that fixed order
        # only if the header can't be recognized.
        col_map = _detect_chem_table_columns(tbl)
        if not col_map:
            col_map = {3 + i: p for i, p in enumerate(COL_PARAMS)}
        table_thresh = _extract_target_thresholds(tbl, col_map)
        chem_cols = sorted(col_map)

        last = _last_row_for_well(tbl, well)
        if last < 0:
            continue

        # If a row for this exact well+date already exists (e.g. the same
        # round's PDF was accidentally included twice in a multi-round
        # upload, or the document was reprocessed with overlapping
        # rounds), overwrite that row instead of inserting a duplicate.
        if tbl.rows[last].cells[1].text.strip() == new_date:
            new_row = tbl.rows[last]
        else:
            new_row = _insert_row_after(tbl, last, last)
        cells = new_row.cells
        # NOTE: cells[0] (well name) is often a vertically-merged cell shared
        # across the whole table — do NOT write to it here, or it will wipe
        # out the merge-origin text for every row in the table.
        _set_cell(cells[1], new_date, highlight=True)
        _set_cell(cells[2], "")

        if floating:
            _set_cell(cells[3], "לא נדגם עקב שכבה צפה", highlight=True)
            for ci in chem_cols:
                if ci != 3 and ci < len(cells):
                    _set_cell(cells[ci], "")
        else:
            for ci in chem_cols:
                if ci >= len(cells):
                    continue
                param = col_map[ci]
                val = results.get(param)
                thresh = table_thresh.get(param, THRESH.get(param))
                if val is None:
                    txt, exceeds = "<0.001", False
                else:
                    exceeds = thresh is not None and val > thresh
                    txt = str(val)
                color = RGBColor(0xC0, 0x00, 0x00) if exceeds else None
                _set_cell(cells[ci], txt, bold=exceeds, color=color, highlight=True)


def _update_historical_field_table_layout_a(doc: Document, new_date: str,
                                             samples: dict, wells_order: list):
    """
    Append one new row per well to the historical field-findings table
    ("טבלה מספר 2 – ממצאי בדיקות השדה בעבר") for simple (Layout A)
    documents, where field history lives in its own dedicated table
    (columns: קידוח, תאריך, pH, EC, טמפרטורה, ...) rather than inside a
    per-well block (that's Layout B, handled by
    _update_historical_field_table). This table was previously never
    updated for Layout A documents at all.

    The table is located by content signature (a "קידוח"/"תאריך" header
    row followed by several recognized field-parameter names), and the
    column→field mapping is read from that header row rather than
    assumed, so it works regardless of column order.
    """
    ti, header_cells = None, None
    for cand_ti, tbl in enumerate(doc.tables):
        for row in tbl.rows:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) < 4 or cells[0] != "קידוח" or cells[1] != "תאריך":
                continue
            hint_hits = sum(1 for c in cells[2:] if c.rstrip("*") in _FIELD_PARAM_HINTS)
            if hint_hits >= 3:
                ti, header_cells = cand_ti, cells
                break
        if ti is not None:
            break
    if ti is None:
        return  # no such table in this document — nothing to do

    tbl = doc.tables[ti]
    col_field = {}
    for ci, label in enumerate(header_cells):
        if ci < 2:
            continue
        key = _FIELD_ROW_ALIASES.get(label.rstrip("*"))
        if key and key != "__date__":
            col_field[ci] = key

    for well in wells_order:
        if well not in samples:
            continue
        sample = samples[well]
        field = sample.get("field", {})
        last = _last_row_for_well(tbl, well)
        if last < 0:
            continue
        if len(tbl.rows[last].cells) > 1 and tbl.rows[last].cells[1].text.strip() == new_date:
            new_row = tbl.rows[last]
        else:
            new_row = _insert_row_after(tbl, last, last)
        cells = new_row.cells
        prev = tbl.rows[last].cells
        if len(cells) > 1:
            _set_cell(cells[1], new_date, highlight=True)
        for ci, key in col_field.items():
            if ci >= len(cells):
                continue
            val = field.get(key)
            # "-" for a reading the round doesn't have, never the old value
            _set_cell(cells[ci], _fmt_like(prev[ci].text if ci < len(prev) else "", val),
                      highlight=True)


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

        # Update the BTEX concentration sentence by rebuilding it from scratch.
        # We replace the entire paragraph content so we don't rely on fragile
        # slot detection. Only detected (non-null) compounds are listed,
        # with no threshold-exceedance parentheses (user adds those manually).
        if "בדיגום" in text and "MTBE" in text and "בריכוזים" in text:
            ORDERED = [("MTBE", "MTBE"), ("בנזן", "בנזן"), ("אתיל בנזן", "אתיל בנזן"),
                       ("טולואן", "טולואן"), ("קסילן", "קסילן")]
            detected = [(name_en, results.get(heb)) for name_en, heb in ORDERED
                        if results.get(heb) is not None]

            if detected:
                if len(detected) == 1:
                    name, val = detected[0]
                    new_sentence = f'בדיגום אותר {name} בריכוז של {val} מ"ג/ליטר.'
                else:
                    names = ", ".join(n for n, _ in detected[:-1]) + f" ו-{detected[-1][0]}"
                    vals = ", ".join(str(v) for _, v in detected[:-1]) + f" ו-{detected[-1][1]}"
                    new_sentence = f'בדיגום אותרו {names} בריכוזים של {vals} מ"ג/ליטר בהתאמה.'
            else:
                new_sentence = 'בדיגום לא אותרו מרכיבי BTEX/MTBE בריכוזים מעל סף הזיהוי.'

            # Clear all runs and write the new sentence into the first run
            if runs:
                runs[0].text = new_sentence
                runs[0].font.highlight_color = WD_COLOR_INDEX.YELLOW
                for extra in runs[1:]:
                    extra.text = ""


def _heb_join_words(items):
    if len(items) == 1:
        return items[0]
    return ", ".join(items[:-1]) + f" ו{items[-1]}"


def _heb_join_values(items):
    if len(items) == 1:
        return items[0]
    return ", ".join(items[:-1]) + f" ו-{items[-1]}"


def _replace_value_after_anchor(paragraph, anchor: str, new_value: str,
                                 value_charset: str = r'[\d\./\-]') -> bool:
    """
    Find `anchor` in the paragraph's concatenated text — which may span
    multiple runs — then replace whatever value immediately follows it
    (a contiguous run of characters matching `value_charset`, e.g. a date)
    with `new_value`, regardless of how many runs the anchor or the old
    value happen to be split across.

    Word frequently fragments text into many small runs after
    track-changes / spellcheck / re-saves, which breaks naive
    "check the very next run" replacement logic in two ways: it can miss
    the value entirely (if the anchor itself is split across runs), or it
    can overwrite only the first fragment of the old value and leave the
    rest sitting right next to the new one. This handles both.

    Returns True if a replacement was made.
    """
    runs = paragraph.runs
    if not runs:
        return False

    chars = []  # (char, run_idx) for every character in the paragraph
    for ridx, r in enumerate(runs):
        for ch in r.text:
            chars.append((ch, ridx))

    full_text = "".join(c for c, _ in chars)
    anchor_pos = full_text.find(anchor)
    if anchor_pos == -1:
        return False

    n = len(chars)
    anchor_end = anchor_pos + len(anchor)
    i = anchor_end
    while i < n and chars[i][0] in " \t":  # skip whitespace right after anchor
        i += 1

    pat = re.compile(value_charset)
    value_end = i
    while value_end < n and pat.match(chars[value_end][0]):
        value_end += 1
    if value_end <= i:
        return False  # nothing value-like follows the anchor

    run_spans = []
    pos = 0
    for r in runs:
        run_spans.append((pos, pos + len(r.text)))
        pos += len(r.text)

    first_touched = None
    for ridx, (s, e) in enumerate(run_spans):
        if e <= anchor_end or s >= value_end:
            continue  # run not part of the value span at all
        text = runs[ridx].text
        local_start = max(anchor_end, s) - s
        local_end = min(value_end, e) - s
        prefix = text[:local_start]
        suffix = text[local_end:]
        if first_touched is None:
            runs[ridx].text = prefix + " " + new_value + suffix
            runs[ridx].font.highlight_color = WD_COLOR_INDEX.YELLOW
            first_touched = ridx
        else:
            runs[ridx].text = prefix + suffix
    return True


def _update_title_month_year(doc: Document, new_dt: datetime) -> bool:
    """Update the report title's month + year wherever it appears in a title paragraph."""
    month_name = HEBREW_MONTHS[new_dt.month - 1]
    year = new_dt.year
    updated = False
    month_pat = re.compile(r'^(' + '|'.join(HEBREW_MONTHS) + r')(\s+\d{4})?$')
    stray_year_pat = re.compile(r'^\s*\d{4}\s*$')
    whitespace_pat = re.compile(r'^\s*$')

    for p in doc.paragraphs:
        if "ניטור מי תהום" not in p.text:
            continue
        runs = p.runs
        for i, run in enumerate(runs):
            m = month_pat.match(run.text.strip())
            if m:
                # Preserve any leading whitespace/punctuation
                leading = run.text[: len(run.text) - len(run.text.lstrip())]
                run.text = f"{leading}{month_name} {year}"
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                # Clear any stray leftover year fragment(s) that may follow —
                # possibly separated from this run by whitespace-only runs,
                # a common artifact of Word's run-splitting. We only clear
                # the whitespace runs too once a stray year is actually
                # found past them, so we don't disturb normal spacing when
                # there's nothing to fix.
                j = i + 1
                lookahead = []
                found_year = False
                while j < len(runs):
                    t = runs[j].text
                    if t.strip() and stray_year_pat.match(t):
                        lookahead.append(j)
                        found_year = True
                        j += 1
                        continue
                    if whitespace_pat.match(t):
                        lookahead.append(j)
                        j += 1
                        continue
                    break
                if found_year:
                    for jj in lookahead:
                        runs[jj].text = ""
                updated = True
    return updated


_CURRENT_ROUND_TRIGGERS = [
    "נערך דיגום",
    "דיגום תקופתי",
    "נערך ניטור מי תהום",
    "נערך ניטור בתחנה",
]


def _update_narrative_placeholders(doc: Document, date_str: str, sampler_name: str):
    """
    Replace the date and sampler-name VALUE runs in the CURRENT-ROUND
    sampling sentences only. This must NOT touch unrelated mentions of
    dates/'מר X' elsewhere in the document (e.g. historical references in
    the background section, which also often start with "בתאריך").

    Different site templates phrase this sentence differently — compare
    "בתאריך X נערך דיגום תקופתי של מי התהום בקידוחים..." vs "בתאריך X נערך
    בתחנה דיגום תקופתי על ידי מר..." vs "בתאריך X נערך ניטור מי תהום
    בקידוח...". Matching on the paragraph both STARTING with "בתאריך" and
    containing one of several known trigger phrases catches all of these
    while still excluding unrelated historical sentences (pipe-integrity
    tests, well installation, well development) that also start with
    "בתאריך" but describe a different kind of event.

    Uses _replace_value_after_anchor() rather than naive adjacent-run
    matching, because Word often splits both the anchor text ("בתאריך")
    and the old date across several small runs — adjacent-run matching
    then either misses the date entirely or leaves stray fragments of the
    old date sitting right next to the newly-inserted one.
    """
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text.startswith("בתאריך"):
            continue
        if not any(trig in text for trig in _CURRENT_ROUND_TRIGGERS):
            continue
        _replace_value_after_anchor(p, "בתאריך", date_str, value_charset=r'[\d\./\-]')

        runs = p.runs
        for i, run in enumerate(runs):
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


def _update_summary_well_bullets(doc: Document, samples: dict):
    """
    Rebuild the per-well BTEX/MTBE bullets in the 'סיכום' section
    (e.g. 'בקידוח מת-2 אותרו בנזן ו-MTBE בריכוזים של ... בהתאמה;') so they
    reflect the latest round's results, per well. Some report templates
    (e.g. multi-well sites) list results per well like this instead of — or
    in addition to — the single combined sentence handled by
    _update_concentration_summary. Only touches paragraphs that already
    start with 'בקידוח <well>' for a well present in `samples`.
    """
    ORDERED = [("MTBE", "MTBE"), ("בנזן", "בנזן"), ("אתיל בנזן", "אתיל בנזן"),
               ("טולואן", "טולואן"), ("קסילן", "קסילן")]

    for well, sample in samples.items():
        results = sample.get("results", {})
        detected = [(name, results.get(heb)) for name, heb in ORDERED if results.get(heb) is not None]

        if not detected:
            sentence = f'בקידוח {well} לא אותרו מרכיבי BTEX/MTBE בריכוזים מעל סף הזיהוי.'
        elif len(detected) == 1:
            name, val = detected[0]
            sentence = f'בקידוח {well} אותר {name} בריכוז של {val} מ"ג לליטר.'
        else:
            names = _heb_join_words([n for n, _ in detected])
            vals = _heb_join_values([f'{v} מ"ג לליטר' for _, v in detected])
            sentence = f'בקידוח {well} אותרו {names} בריכוזים של {vals} בהתאמה.'

        anchor = f"בקידוח {well}"
        for p in doc.paragraphs:
            text = p.text.strip()
            if not text.startswith(anchor):
                continue
            if "אותר" not in text and "לא אותרו" not in text:
                continue
            if "חורג" in text:
                continue  # exceedance bullets: _update_exceedance_summary
            runs = p.runs
            if not runs:
                continue
            # Preserve the paragraph's own closing punctuation style
            # (bullets in this section may end with ';' or '.').
            out_sentence = sentence
            if text.endswith(";"):
                out_sentence = out_sentence.rstrip(".") + ";"
            runs[0].text = out_sentence
            runs[0].font.highlight_color = WD_COLOR_INDEX.YELLOW
            for extra in runs[1:]:
                extra.text = ""
            break


def _update_report_meta(doc: Document, author_name: str = None, submission_date: str = None):
    """
    Fill in the report-metadata table on the cover page: the report
    author's name, and the submission date. Layout observed:
        ['חתימה', 'מחבר הדו"ח']   <- label row
        ['',      'ערן רזניק']    <- value row (name goes here)
        ['חתימה', 'הדו"ח מאשר']
        ['',      'אחיעד ווייס']
        ['',      'תאריך הגשה']   <- label row; the date value goes in
                                      THIS row's own first cell (column 0),
                                      not a new row — unlike the name
                                      fields, there's no signature to leave
                                      room for here.
    Only touches the table if it actually looks like this metadata table
    (has a 'מחבר' label), so it's a no-op on documents structured
    differently. Both arguments are optional — pass None/empty to leave
    that field untouched.
    """
    if not author_name and not submission_date:
        return
    if not doc.tables:
        return

    tbl = doc.tables[0]
    labels = [row.cells[1].text.strip() if len(row.cells) > 1 else "" for row in tbl.rows]
    if not any("מחבר" in lbl for lbl in labels):
        return  # doesn't look like the metadata table — don't touch it

    if author_name:
        for ri, lbl in enumerate(labels):
            if "מחבר" in lbl and ri + 1 < len(tbl.rows):
                _set_cell(tbl.rows[ri + 1].cells[1], author_name, highlight=True)
                break

    if submission_date:
        for ri, lbl in enumerate(labels):
            if "תאריך" in lbl and "הגש" in lbl:
                if len(tbl.rows[ri].cells) > 0:
                    _set_cell(tbl.rows[ri].cells[0], submission_date, highlight=True)
                # Clean up a stray extra value row that an earlier version
                # of this function may have added below (it incorrectly
                # inserted a new row instead of writing into this row's
                # own first cell) — otherwise the date would appear twice.
                if ri + 1 < len(tbl.rows):
                    nxt_cells = tbl.rows[ri + 1].cells
                    nxt0 = nxt_cells[0].text.strip() if len(nxt_cells) > 0 else ""
                    nxt1 = nxt_cells[1].text.strip() if len(nxt_cells) > 1 else ""
                    next_is_label = any(k in nxt1 for k in ("מחבר", "מאשר", "חתימה", "תאריך"))
                    looks_like_stray_date = bool(re.match(r'^[\d/\.\-]+$', nxt1))
                    if not nxt0 and not next_is_label and looks_like_stray_date:
                        _set_cell(nxt_cells[1], "")
                break


def _well_num(name: str):
    m = re.search(r'(\d+)\s*$', str(name).strip())
    return m.group(1) if m else None


def _merge_field_form(samples: dict, form: dict) -> list:
    """
    Put the field form's readings into the lab samples, matching wells by
    their number ("חק-2" in the form, "HK-2" at the lab). The form is the
    field record, so its values win over field values printed by the lab.
    Returns the form wells that had no lab sample.
    """
    by_num = {_well_num(k): k for k in samples}
    unmatched = []
    for well, w in form.get("wells", {}).items():
        key = by_num.get(_well_num(well))
        if key is None:
            unmatched.append(well)
            continue
        s = samples[key]
        s.setdefault("field", {}).update(w.get("field", {}))
        for k in ("water_level", "total_depth", "volume_l", "flow_lpm"):
            if w.get(k) is not None:
                s[k] = w[k]
        if w.get("floating_layer"):
            s["floating_layer"] = True
    return unmatched


def _fmt_conc(val) -> str:
    """Concentration as text without exponent notation: 0.00421, 7.11, 2700."""
    if val is None:
        return ""
    if val == 0:
        return "0"
    if abs(val) >= 1000:
        return f"{val:.0f}"
    if abs(val) < 1:
        # three significant figures, as the lab prints them
        digits = -int(f"{abs(val):e}".split("e")[1]) + 2
        return f"{val:.{max(digits, 0)}f}".rstrip("0").rstrip(".")
    return f"{val:.3g}" if abs(val) < 100 else f"{val:.0f}"


def _fmt_like(old_text: str, val) -> str:
    """
    Field value in the style of the cell it replaces (or the row above):
    same number of decimals, thousands separator if the old text had one.
    """
    if val is None:
        return "-"
    old = (old_text or "").strip()
    m = re.fullmatch(r'-?[\d,]*\.?(\d*)-?', old)
    if not old or not m or not re.search(r'\d', old):
        return _fmt_field_val(val)
    own = f"{val:.3f}".rstrip("0").split(".")
    own_decimals = len(own[1]) if len(own) > 1 else 0
    decimals = max(len(m.group(1)), own_decimals)  # never round away a measured digit
    if "," in old:
        return f"{val:,.{decimals}f}"
    return f"{val:.{decimals}f}"


def _cell_targets(tbl) -> dict:
    """
    The table's own target row ("ריכוז יעד לשיקום", "ריכוזי יעד לסיום
    ניטור", "תקן מי שתייה" — any label with יעד or תקן, in any column).
    """
    col_map = _detect_chem_table_columns(tbl) or {3 + i: p for i, p in enumerate(COL_PARAMS)}
    for row in tbl.rows[:6]:
        cells = row.cells
        labels = " ".join(c.text for ci, c in enumerate(cells) if ci not in col_map)
        if "יעד" not in labels and "תקן" not in labels:
            continue
        out = {}
        for ci, param in col_map.items():
            if ci < len(cells):
                try:
                    out[param] = float(cells[ci].text.strip().replace(",", ""))
                except ValueError:
                    pass
        if out:
            return out
    return {}


def _result_text(results: dict, limits: dict, param: str):
    """(text, value) for one compound: '<0.0002' for a non-detect."""
    if param not in results:
        return "-", None
    val = results[param]
    if val is None:
        lim = limits.get(param)
        return (f"<{_fmt_conc(lim)}" if lim else "<0.001"), None
    return _fmt_conc(val), val


def _is_chem_history_table(tbl) -> bool:
    if len(tbl.columns) < 6:
        return False
    head = " ".join(c.text for row in tbl.rows[:3] for c in row.cells)
    return "תאריך" in head and bool(_detect_chem_table_columns(tbl))


def _update_chem_history_tables(doc: Document, new_date: str, samples: dict, wells_order: list) -> bool:
    """
    One new row per well in the chemistry history table that holds that well
    (one or several tables, e.g. "חלק 1 מתוך 2"). The columns come from the
    table's header, the targets from its own target row, and a water level
    column ("עומק פני המים") is filled when the table has one. Returns False
    when the document has no such table (older layouts are handled by
    _update_historical_tables).
    """
    tables = [t for t in doc.tables if _is_chem_history_table(t)]
    if not tables:
        return False
    for well in wells_order:
        sample = samples.get(well)
        if not sample:
            continue
        tbl = next((t for t in tables
                    if any(r.cells and r.cells[0].text.strip() == well for r in t.rows)), None)
        if tbl is None:
            continue
        col_map = _detect_chem_table_columns(tbl)
        targets = _cell_targets(tbl)
        head = [c.text.strip() for c in tbl.rows[0].cells]
        level_col = next((ci for ci, h in enumerate(head) if "פני המים" in h or "מפלס" in h), None)

        last = _last_row_for_well(tbl, well)
        if last < 0:
            continue
        prev_text = {ci: c.text for ci, c in enumerate(tbl.rows[last].cells)}
        if tbl.rows[last].cells[1].text.strip() == new_date:
            new_row = tbl.rows[last]
        else:
            new_row = _insert_row_after(tbl, last, last)
        cells = new_row.cells
        _set_cell(cells[1], new_date, highlight=True)
        if level_col is not None and level_col < len(cells):
            lvl = sample.get("water_level")
            _set_cell(cells[level_col], _fmt_like(prev_text.get(level_col, ""), lvl), highlight=True)
        results, limits = sample.get("results", {}), sample.get("limits", {})
        for ci, param in col_map.items():
            if ci >= len(cells):
                continue
            if sample.get("floating_layer"):
                _set_cell(cells[ci], "שכבה צפה" if ci == min(col_map) else "", highlight=True)
                continue
            txt, val = _result_text(results, limits, param)
            exceeds = val is not None and param in targets and val > targets[param]
            _set_cell(cells[ci], txt, bold=exceeds, highlight=True)
    return True


def _is_well_card(tbl):
    """
    A small per-well result table ("חק-4 / 02.06.26 / תרכובת | ריכוז יעד |
    ריכוז"): row 0 is the well name in every cell, row 1 the date. Returns
    the well name, or None.
    """
    if len(tbl.rows) < 4 or len(tbl.columns) > 4:
        return None
    r0 = {c.text.strip() for c in tbl.rows[0].cells}
    r1 = {c.text.strip() for c in tbl.rows[1].cells}
    if len(r0) != 1 or len(r1) != 1:
        return None
    if not re.fullmatch(r'\d{2}\.\d{2}\.\d{2,4}', next(iter(r1))):
        return None
    if "תרכובת" not in tbl.rows[2].cells[0].text:
        return None
    return next(iter(r0))


def _update_well_cards(doc: Document, new_date: str, samples: dict) -> int:
    """Date and results of the per-well tables on the site map. Returns how many were updated."""
    by_num = {_well_num(w): w for w in samples}
    done = 0
    for tbl in doc.tables:
        well = _is_well_card(tbl)
        if not well:
            continue
        key = by_num.get(_well_num(well))
        if key is None:
            continue
        sample = samples[key]
        results, limits = sample.get("results", {}), sample.get("limits", {})
        for c in {id(c._tc): c for c in tbl.rows[1].cells}.values():
            _set_cell(c, new_date, highlight=True)
        for row in tbl.rows[3:]:
            cells = row.cells
            param = _PARAM_CANONICAL.get(cells[0].text.strip())
            if not param:
                continue
            txt, val = _result_text(results, limits, param)
            try:
                target = float(cells[1].text.strip())
            except (ValueError, IndexError):
                target = None
            exceeds = val is not None and target is not None and val > target
            _set_cell(cells[-1], txt, bold=exceeds, highlight=True)
        done += 1
    return done


def _replace_in_paragraph(paragraph, pattern, repl) -> bool:
    """
    Regex replace over the paragraph's whole text, even when Word split it
    into many runs. The new text goes into the first run the match touches
    (keeping its formatting, highlighted); the rest of the match is removed.
    """
    runs = paragraph.runs
    if not runs:
        return False
    text = "".join(r.text for r in runs)
    m = re.search(pattern, text)
    if not m:
        return False
    new = m.expand(repl) if isinstance(repl, str) else repl(m)
    start, end = m.span()
    pos = 0
    first = True
    for r in runs:
        rs, rend = pos, pos + len(r.text)
        pos = rend
        if rend <= start or rs >= end:
            continue
        a, b = max(start, rs) - rs, min(end, rend) - rs
        if first:
            r.text = r.text[:a] + new + r.text[b:]
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
            first = False
        else:
            r.text = r.text[:a] + r.text[b:]
    return not first


def _num_txt(v: float) -> str:
    return f"{int(v)}" if v == int(v) else f"{v:g}"


def _update_sampling_paragraphs(doc: Document, samples: dict, sampler_name: str):
    """
    The field-round sentences: the sampler's name ("על ידי מר X, דוגם"),
    the pumped volume ("נעה בין כ- 16 ל-22 ליטר") and the flow ("בספיקה של
    0.4 ליטר בדקה" / "בספיקה שבין 0.40-0.48 ליטר בדקה").
    """
    vols = [s["volume_l"] for s in samples.values() if s.get("volume_l") is not None]
    flows = sorted({round(s["flow_lpm"], 2) for s in samples.values() if s.get("flow_lpm") is not None})
    for p in doc.paragraphs:
        t = p.text
        if sampler_name and "על ידי מר" in t and "דוגם" in t:
            _replace_in_paragraph(p, r"(על ידי מר)\s*[^,]+?(,\s*דוגם)", rf"\1 {sampler_name}\2")
        if vols and "ליטר" in t and "נעה בין" in t:
            lo, hi = round(min(vols)), round(max(vols))
            if lo == hi:
                _replace_in_paragraph(p, r"נעה בין כ-\s*[\d.]+\s*ל-\s*[\d.]+\s*ליטר",
                                      f"עמדה על כ- {_num_txt(lo)} ליטר")
            else:
                _replace_in_paragraph(p, r"(נעה בין כ-)\s*[\d.]+\s*ל-\s*[\d.]+\s*(ליטר)",
                                      rf"\1 {_num_txt(lo)} ל-{_num_txt(hi)} \2")
        if flows and "בספיקה" in t:
            if len(flows) == 1:
                flow_txt = f"של {_num_txt(flows[0])}"
            else:
                flow_txt = f"שבין {flows[0]:.2f}-{flows[-1]:.2f}"
            _replace_in_paragraph(p, r"בספיקה\s+(?:של\s+[\d.]+|שבין\s+[\d.]+\s*-\s*[\d.]+)",
                                  f"בספיקה {flow_txt}")


def _update_exceedance_summary(doc: Document, samples: dict, wells_order: list, targets: dict) -> bool:
    """
    The summary bullets about exceedances of the site targets ("בקידוח חק-7
    אותר בנזן בריכוז 0.743 מ"ג/ל' אשר חורג מערך היעד לשיקום (0.0935
    מ"ג/ליטר)." and "בקידוחים חק-2, חק-3 ו- חק-4 לא אותרו ריכוזי BTEX MTBE
    החורגים מערכי היעד לשיקום;"). Rebuilt from this round against the
    report's own targets: one bullet for the wells without exceedances, one
    per well with them. Returns False when the report has no such bullets.
    """
    from docx.text.paragraph import Paragraph

    paras = [p for p in doc.paragraphs
             if p.text.strip().startswith("בקידוח") and "חורג" in p.text]
    if not paras or not targets:
        return False
    clean, exceed = [], []
    for well in wells_order:
        s = samples.get(well)
        if not s or s.get("floating_layer"):
            continue
        res = s.get("results", {})
        over = [(p, res[p]) for p in COL_PARAMS
                if res.get(p) is not None and p in targets and res[p] > targets[p]]
        (exceed if over else clean).append((well, over))

    sentences = []
    if clean:
        names = [w for w, _ in clean]
        if len(names) == 1:
            sentences.append(f"בקידוח {names[0]} לא אותרו ריכוזי BTEX ו-MTBE החורגים מערכי היעד לשיקום;")
        else:
            sentences.append("בקידוחים " + ", ".join(names[:-1]) + f" ו-{names[-1]}"
                             " לא אותרו ריכוזי BTEX ו-MTBE החורגים מערכי היעד לשיקום;")
    for well, over in exceed:
        parts = [f'{p} בריכוז {_fmt_conc(v)} מ"ג/ל\' אשר חורג מערך היעד לשיקום '
                 f'({_fmt_conc(targets[p])} מ"ג/ליטר)' for p, v in over]
        sentences.append(f"בקידוח {well} אותר " + "; ".join(parts) + ".")

    for p, sentence in zip(paras, sentences):
        runs = p.runs
        runs[0].text = sentence
        runs[0].font.highlight_color = WD_COLOR_INDEX.YELLOW
        for r in runs[1:]:
            r.text = ""
    last = paras[min(len(paras), len(sentences)) - 1]
    for sentence in sentences[len(paras):]:
        new_p = copy.deepcopy(last._p)
        last._p.addnext(new_p)
        np_ = Paragraph(new_p, last._parent)
        np_.runs[0].text = sentence
        np_.runs[0].font.highlight_color = WD_COLOR_INDEX.YELLOW
        for r in np_.runs[1:]:
            r.text = ""
        last = np_
    for p in paras[len(sentences):]:
        p._p.getparent().remove(p._p)
    return True


def _replace_textbox_dates(doc: Document, old_date: str, new_date: str) -> int:
    """Dates written in text boxes on the site map (e.g. '02.06.26' labels)."""
    if not old_date or old_date == new_date:
        return 0
    n = 0
    for txbx in doc.element.body.iter(qn("w:txbxContent")):
        for t in txbx.iter(qn("w:t")):
            if t.text and old_date in t.text:
                t.text = t.text.replace(old_date, new_date)
                n += 1
    return n


def _latest_history_date(doc: Document):
    """The newest date already in the chemistry history (the previous round)."""
    best = None
    for tbl in doc.tables:
        if not _is_chem_history_table(tbl):
            continue
        for row in tbl.rows:
            if len(row.cells) < 2:
                continue
            t = row.cells[1].text.strip().rstrip("*")
            try:
                d = datetime.strptime(t, "%d.%m.%y")
            except ValueError:
                continue
            if best is None or d > best[0]:
                best = (d, t)
    return best[1] if best else None


_MK_REMINDER = ("⚠️ שים לב: יש לעדכן את קובץ Mann-Kendall בתיקייה "
                "ולהדביק את הגרפים המעודכנים לדוח.")


def _remove_mk_reminders(doc: Document):
    for p in list(doc.paragraphs):
        if "יש לעדכן את קובץ Mann-Kendall" in p.text:
            p._p.getparent().remove(p._p)


def _add_mk_reminder(doc: Document):
    """Yellow note before the first chart, once (not again on every round)."""
    if any("יש לעדכן את קובץ Mann-Kendall" in p.text for p in doc.paragraphs):
        return
    chart_para = None
    for p in doc.paragraphs:
        if re.match(r"\s*גרף מספר", p.text):
            chart_para = p
            break
    if chart_para is None:
        for pi, p in enumerate(doc.paragraphs):
            if pi > 40 and p._p.findall('.//' + qn('w:drawing')):
                chart_para = p
                break
    from docx.oxml import OxmlElement
    new_p = OxmlElement('w:p')
    r = OxmlElement('w:r')
    rpr = OxmlElement('w:rPr')
    b = OxmlElement('w:b')
    hl = OxmlElement('w:highlight')
    hl.set(qn('w:val'), 'yellow')
    rpr.append(b)
    rpr.append(hl)
    r.append(rpr)
    t = OxmlElement('w:t')
    t.text = _MK_REMINDER
    r.append(t)
    new_p.append(r)
    if chart_para is not None:
        chart_para._p.addprevious(new_p)
    else:
        doc.element.body.append(new_p)


_CHART_PARAM_WORDS = [
    ("אתיל בנזן", "אתיל בנזן"), ("בנזן", "בנזן"), ("MTBE", "MTBE"),
    ("טולואן", "טולואן"), ("קסילן", "קסילן"), ("כסילן", "קסילן"),
]


def _chart_slots(doc: Document) -> dict:
    """
    {param: inline drawing} for each caption "גרף מספר N: ... <compound>"
    followed (within a few paragraphs) by a paragraph with a picture.
    """
    paras = doc.paragraphs
    slots = {}
    for i, p in enumerate(paras):
        t = p.text.strip()
        if not t.startswith("גרף מספר"):
            continue
        param = next((pp for word, pp in _CHART_PARAM_WORDS if word in t), None)
        if not param or param in slots:
            continue
        for q in paras[i + 1:i + 5]:
            inl = q._p.find('.//' + qn('wp:inline'))
            if inl is not None:
                slots[param] = inl
                break
    return slots


def _put_picture(doc: Document, inline, png_bytes: bytes):
    """Replace the picture in an inline drawing: same width, height to the new aspect."""
    from PIL import Image
    rId, _ = doc.part.get_or_add_image(BytesIO(png_bytes))
    blip = inline.find('.//' + qn('a:blip'))
    blip.set(qn('r:embed'), rId)
    w_px, h_px = Image.open(BytesIO(png_bytes)).size
    ext = inline.find(qn('wp:extent'))
    cx = int(ext.get('cx'))
    cy = int(cx * h_px / w_px)
    ext.set('cy', str(cy))
    for e in inline.iter(qn('a:ext')):
        if e.get('cx') is not None:
            e.set('cx', str(cx))
            e.set('cy', str(cy))


def _update_mann_kendall(doc: Document, mk_xls_bytes: bytes, new_dt: datetime,
                         samples: dict) -> tuple:
    """
    Add the round to the toolkit sheets in memory, draw each sheet the
    report shows, and put the pictures in place of the old ones.
    Returns (paste_rows_xlsx_bytes, [params whose picture was replaced], [notes]).
    """
    from soil_lab_tool.core import gw_mann_kendall as mk

    wb = mk.read_mk_workbook(mk_xls_bytes)
    sheets = mk.report_sheets(wb)
    slots = _chart_slots(doc)
    by_num = {_well_num(w): s for w, s in samples.items()}
    notes, updated, touched = [], [], []
    for param, sh in sheets.items():
        values = {}
        for w in sh.wells:
            s = by_num.get(_well_num(w))
            if s is None or s.get("floating_layer"):
                values[w] = None  # not sampled this round
                continue
            res = s.get("results", {})
            values[w] = mk.mk_value(res.get(param), measured=param in res)
        if all(v is None for v in values.values()):
            continue
        if not mk.add_round(sh, new_dt, values):
            notes.append(f"בגיליון {sh.name} בקובץ Mann-Kendall אין שורה פנויה לסבב חדש.")
            continue
        touched.append(sh)
        if param in slots:
            stats = mk.sheet_stats(sh, wb.kendall_dist)
            _put_picture(doc, slots[param], mk.render_sheet_png(sh, stats, new_dt))
            updated.append(param)
    paste = mk.paste_rows_xlsx(touched, new_dt) if touched else None
    return paste, updated, notes


# ──────────────────────────────────────────────────────────────────────────────
# PUBLIC API
# ──────────────────────────────────────────────────────────────────────────────

def run_update_bytes(
    word_bytes: bytes,
    lab_pdf_bytes,
    mk_xls_bytes: bytes = None,
    field_pdf_bytes: bytes = None,   # kept for old callers; not read
    lab_type: str = "bactochem",     # "bactochem", "aminolab" or "als"
    author_name: str = None,         # optional: fills the cover-page "מחבר הדו"ח" field
    submission_date: str = None,     # optional: fills the cover-page "תאריך הגשה" field
    field_form_bytes: bytes = None,  # optional: Yeda field form (.xlsx)
    report: dict = None,             # optional: filled with what was done, for the app
) -> tuple:
    """
    Apply one sampling round to the report and return
    (updated_word_bytes, mk_paste_rows_xlsx_bytes or None).

    lab_pdf_bytes:
      - "bactochem": a single PDF's bytes (one PDF covers all wells).
      - "aminolab":  bytes of ONE PDF, or a list/tuple of bytes — one per
        well/certificate — since each Aminolab certificate covers a single well.
      - "als":       the ALS certificate, PDF or the XLSX appendix.
    """
    if report is None:
        report = {}
    report.setdefault("notes", [])

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)

        # ── Parse lab results ──────────────────────────────────────
        if lab_type == "als":
            from soil_lab_tool.core.gw_round_inputs import parse_als_round
            lab = parse_als_round(lab_pdf_bytes)
            samples, new_date = lab["samples"], lab["sampling_date"]
        elif lab_type == "aminolab":
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

        # ── Field form ─────────────────────────────────────────────
        sampler_name = next((s.get("sampler_name") for s in samples.values() if s.get("sampler_name")), None)
        if field_form_bytes:
            from soil_lab_tool.core.gw_round_inputs import parse_yeda_field_form
            form = parse_yeda_field_form(field_form_bytes)
            missing = _merge_field_form(samples, form)
            if missing:
                report["notes"].append("בטופס השדה יש בארות שאין להן תוצאות מעבדה: " + ", ".join(missing))
            if form.get("sampler_name"):
                sampler_name = form["sampler_name"]
            if form.get("sampling_date"):
                form_date = form["sampling_date"].strftime("%d.%m.%y")
                if new_date and form_date != new_date:
                    report["notes"].append(
                        f"תאריך הדיגום בטופס השדה ({form_date}) שונה מהתאריך בתעודת המעבדה ({new_date}). "
                        "נלקח התאריך של המעבדה.")
                new_date = new_date or form_date
        if not new_date:
            raise ValueError("לא נמצא תאריך דיגום בתעודת המעבדה.")
        if not samples:
            raise ValueError("לא נמצאו דגימות בתעודת המעבדה.")
        new_dt = datetime.strptime(new_date, "%d.%m.%y")

        # ── Update Word document ──────────────────────────────────
        doc = Document(BytesIO(word_bytes))
        layout = _detect_layout(doc)
        previous_date = _latest_history_date(doc)

        # Detect the well-naming convention actually used in THIS document
        # (different sites use different prefixes, e.g. "מת-1" or "צא - 1"),
        # then remap the parsed sample keys onto it by well number.
        #
        # Some documents have MULTIPLE tables that both mention the same well
        # number under different name spellings — e.g. a one-off "before/after
        # injection" snapshot table using Latin "PP-1", alongside the routine
        # long-running chronological table using Hebrew "פפ-1". Score each
        # name variant by how many chronological date-rows exist in the
        # table(s) it appears in, and prefer the variant from the table with
        # the deepest history — the routine table meant to receive new rows.
        _date_pat = re.compile(r'^\d{2}\.\d{2}\.\d{2}$')
        well_depth = {}
        for tbl in doc.tables:
            date_rows = sum(
                1 for row in tbl.rows
                if len(row.cells) >= 2 and _date_pat.match(row.cells[1].text.strip())
            )
            for row in tbl.rows:
                if not row.cells:
                    continue
                w = row.cells[0].text.strip()
                if w and re.search(r'\d', w) and len(w) < 20:
                    if date_rows > well_depth.get(w, -1):
                        well_depth[w] = date_rows

        num_to_doc_well = {}
        num_to_depth = {}
        for w, depth in well_depth.items():
            num = _well_num(w)
            if num:
                if num not in num_to_doc_well or depth > num_to_depth[num]:
                    num_to_doc_well[num] = w
                    num_to_depth[num] = depth

        remapped_samples = {}
        unmatched = []
        for well_key, sample in samples.items():
            doc_well = num_to_doc_well.get(_well_num(well_key))
            if doc_well:
                remapped_samples[doc_well] = sample
            else:
                remapped_samples[well_key] = sample
                unmatched.append(well_key)
        samples = remapped_samples
        wells_order = sorted(samples, key=lambda w: int(_well_num(w) or 0))

        if unmatched:
            raise ValueError(
                "לא נמצאה בדוח ה-Word התאמה לבארות הבאות מתעודת המעבדה: "
                + ", ".join(unmatched)
                + ". ודא שמספרי הבארות בתעודה תואמים לבארות בדוח ה-Word."
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
        else:
            if not _update_chem_history_tables(doc, new_date, samples, wells_order):
                _update_historical_tables(doc, new_date, samples, wells_order)
            _update_historical_field_table_layout_a(doc, new_date, samples, wells_order)
        cards = _update_well_cards(doc, new_date, samples)
        if cards:
            report["well_cards"] = cards
        _replace_textbox_dates(doc, previous_date, new_date)

        # ── Narrative ─────────────────────────────────────────────
        full_date_str = new_dt.strftime("%d.%m.%Y")
        combined_results = {}
        for s in samples.values():
            combined_results.update(s.get("results", {}))

        _update_title_month_year(doc, new_dt)
        _update_cover_page_date(doc, new_dt)
        if layout == "B":
            _update_narrative_layout_b(doc, full_date_str, combined_results)
        else:
            _update_narrative_placeholders(doc, new_date, sampler_name)
            _update_sampling_paragraphs(doc, samples, sampler_name)
            _update_concentration_summary(doc, combined_results)
            targets = next((_cell_targets(t) for t in doc.tables if _is_chem_history_table(t)), {})
            _update_exceedance_summary(doc, samples, wells_order, targets)
            _update_summary_well_bullets(doc, samples)

        _update_report_meta(doc, author_name=author_name, submission_date=submission_date)

        # ── Mann-Kendall ──────────────────────────────────────────
        paste_bytes = None
        if mk_xls_bytes:
            paste_bytes, charts, mk_notes = _update_mann_kendall(doc, mk_xls_bytes, new_dt, samples)
            report["mk_charts"] = charts
            report["notes"].extend(mk_notes)
            _remove_mk_reminders(doc)
            if not charts:
                report["notes"].append(
                    "לא נמצאו בדוח גרפים עם כיתוב 'גרף מספר' לבנזן/MTBE/טולואן, ולכן הגרפים לא הוחלפו.")
                _add_mk_reminder(doc)
        else:
            _add_mk_reminder(doc)

        report["sampling_date"] = new_date
        report["wells"] = wells_order
        out = BytesIO()
        doc.save(out)
        return out.getvalue(), paste_bytes
