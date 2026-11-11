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
    # chlorinated compounds, for sites with a "תרכובות מוכלרות" table
    "127-18-4":  "PCE",
    "79-01-6":   "TCE",
    "156-59-2":  "cis-1,2-DCE",
    "75-35-4":   "1,1-DCE",
    "75-01-4":   "VC",
    "107-06-2":  "1,2-DCA",
    "106-93-4":  "EDB",
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
    # the lab formats the well code — tolerate it. The well_id keeps the
    # letters the certificate prints ("פפ-3", "מת-5"); the report updater
    # matches wells by their number.
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
            segments.append((f"{m.group(2)[::-1]}-{m.group(1)}", i))
            continue
        # Latin well codes ("EAZ-2", "N-1"), as in the certificates Yeda Water sends
        m = re.search(r"(?<![A-Za-z])([A-Z]{1,4})-?(\d{1,3})(?![\d.])", line)
        if m:
            segments.append((f"{m.group(1)}-{m.group(2)}", i))
    segments.append((None, len(lines)))  # sentinel

    # Sampler, in the header: "... ןאדמח הללאדבע :םגוד" (mirrored "דוגם: עבדאללה חמדאן")
    # A certificate inside Yeda Water's "דוח ביצוע" was sampled by Yeda.
    yeda_org = "י.ד.ע מים וסביבה" if any("ע.ד.י" in ln for ln in lines[:15]) else None
    sampler = None
    for line in lines:
        m = re.search(r'(\S+)\s+(\S+)\s+:םגוד\s*$', line)
        if m:
            sampler = f"{m.group(2)[::-1]} {m.group(1)[::-1]}"
            break

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
            "field_text": {},  # the values as printed ("11.80"), for the report
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

            # Analyte: "CAS #: 71-43-2 0.001 mg/L 0.720 MTBE", sometimes with
            # the drinking-water standard between unit and result ("X≤ 0.005")
            m = re.search(
                r'CAS #:\s*([\d\-]+)\s+([\d.]+)\s+mg/L\s+(?:X\S?\s*[\d.]+\s+)?([\d.]+|Not Detected)',
                line, re.IGNORECASE
            )
            if m:
                cas, loq, raw_val = m.group(1), m.group(2), m.group(3)
                heb = CAS_TO_HEB.get(cas)
                if heb:
                    sample["results"][heb] = (
                        None if raw_val.lower() == "not detected"
                        else float(raw_val)
                    )
                    if raw_val.lower() == "not detected":
                        try:
                            sample.setdefault("limits", {})[heb] = float(loq)
                        except ValueError:
                            pass
                    if raw_val.lower() != "not detected":
                        # as printed ("0.370"), so the report keeps the lab's digits
                        sample.setdefault("result_text", {})[heb] = raw_val

            # Floating layer: "ןכ : הפצ הבכש תוחכונ"
            if "הפצ הבכש" in line and "ןכ" in line:
                sample["floating_layer"] = True

            # Field: water level "M 15.59 ןוילע סלפמ קמוע"
            m = re.search(r'M\s+([\d.]+)\s+ןוילע סלפמ קמוע', line)
            if m:
                v = float(m.group(1))
                sample["water_level"] = v
                sample["field"]["עומק פני המים"] = v
                sample["field_text"]["עומק פני המים"] = m.group(1)

            # EC: "µS/cm 6640 רחאל) תילמשח תוכילומ"
            m = re.search(r'µS/cm\s+([\d,]+)', line)
            if m:
                sample["field"]["EC"] = float(m.group(1).replace(",", ""))
                sample["field_text"]["EC"] = m.group(1)

            # pH: "pH units 7.11"
            m = re.search(r'pH units\s+([\d.]+)', line)
            if m:
                sample["field"]["pH"] = float(m.group(1))
                sample["field_text"]["pH"] = m.group(1)

            # DO: "mg/L 2.1 (תובצייתה רחאל) סמומ ןצמח"
            m = re.search(r'mg/L\s+([\d.]+)\s+\(תובצייתה רחאל\) סמומ ןצמח', line)
            if m:
                sample["field"]["חמצן מומס"] = float(m.group(1))
                sample["field_text"]["חמצן מומס"] = m.group(1)

            # Redox: "mV -229 (תובצייתה רחאל) סקודר"
            m = re.search(r'mV\s+(-?[\d.]+)\s+\(תובצייתה רחאל\) סקודר', line)
            if m:
                sample["field"]["רדוקס"] = float(m.group(1))
                sample["field_text"]["רדוקס"] = m.group(1)

            # Temp: "C° 27.9 (תובצייתה רחאל) הרוטרפמט"
            m = re.search(r'C°\s+([\d.]+)\s+\(תובצייתה רחאל\) הרוטרפמט', line)
            if m:
                sample["field"]["טמפרטורה"] = float(m.group(1))
                sample["field_text"]["טמפרטורה"] = m.group(1)

            # Turbidity: "NTU 18 (תובצייתה רחאל) תוריכע"
            m = re.search(r'NTU\s+([\d.]+)\s+\(תובצייתה רחאל\) תוריכע', line)
            if m:
                sample["field"]["עכירות"] = float(m.group(1))
                sample["field_text"]["עכירות"] = m.group(1)

            # Total depth: "M 21.62 חודיקה לש יללכ קמוע"
            m = re.search(r'M\s+([\d.]+)\s+חודיקה לש יללכ קמוע', line)
            if m:
                v = float(m.group(1))
                sample["total_depth"] = v
                sample["field"]["עומק כללי של הקידוח"] = v
                sample["field_text"]["עומק כללי של הקידוח"] = m.group(1)

            # Pumping depth: "M 12.87 הביאש קמוע"
            m = re.search(r'M\s+([\d.]+)\s+הביאש קמוע', line)
            if m:
                sample["field"]["עומק שאיבה"] = float(m.group(1))

        if sampler:
            sample["sampler_name"] = sampler
            sample["sampler_org"] = yeda_org or "בקטוכם"
        samples[well_id] = sample

    return {"sampling_date": global_date, "samples": samples}


# Bactochem CSV (the certificate's data file, opens in Excel): one row per
# sample × component, the low-flow readings included. English component
# name → field key used by the PDF parser.
_BACTO_CSV_FIELDS = {
    "depth of upper level": "עומק פני המים",
    "dissolved o2": "חמצן מומס",
    "elctr.conductivity": "EC",
    "electr.conductivity": "EC",
    "electrical conductivity": "EC",
    "ph": "pH",
    "redox": "רדוקס",
    "temp": "טמפרטורה",
    "turbidity": "עכירות",
    "total depth of drilling": "עומק כללי של הקידוח",
    "sampling depth": "עומק שאיבה",
}


def parse_bactochem_csv(data: bytes) -> dict:
    """
    Read Bactochem's CSV export into the same shape as parse_bactochem_pdf.
    Columns are found by their header ("רכיב", "תוצאה", "תאריך דיגום",
    "תיאור דוגמה", "CAS_NUM"); the well is the end of the sample description
    ("דיגום מי תהום-דלק פרבר- פפ 3" → "פפ-3"). The CSV has no sampler's full
    name, so the sampler sentence of the report is left as it is.
    """
    import csv as _csv
    for enc in ("utf-8-sig", "cp1255"):
        try:
            text = data.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError("קידוד לא מוכר")
    import io as _io
    rows = list(_csv.reader(_io.StringIO(text)))
    if not rows:
        raise ValueError("קובץ ריק")
    head = [h.strip() for h in rows[0]]

    def col(*names):
        for n in names:
            if n in head:
                return head.index(n)
        return None

    c_comp, c_val = col("רכיב"), col("תוצאה")
    c_date, c_desc, c_cas = col("תאריך דיגום"), col("תיאור דוגמה"), col("CAS_NUM")
    if None in (c_comp, c_val, c_date, c_desc):
        raise ValueError("לא קובץ CSV של בקטוכם")

    samples, global_date = {}, None
    for r in rows[1:]:
        if len(r) <= max(c_comp, c_val, c_date, c_desc):
            continue
        m = re.search(r"([א-ת]{1,3})['׳]?[\s\-]*(\d+)\s*$", r[c_desc].strip())
        if not m:
            continue
        well = f"{m.group(1)}-{m.group(2)}"
        try:
            date = datetime.strptime(r[c_date].strip()[:10], "%d/%m/%Y").strftime("%d.%m.%y")
        except ValueError:
            continue
        global_date = global_date or date
        s = samples.setdefault(well, {
            "date": date, "results": {}, "field": {}, "field_text": {},
            "water_level": None, "total_depth": None, "floating_layer": False,
        })
        comp, raw = r[c_comp].strip(), r[c_val].strip()
        cas = r[c_cas].strip() if c_cas is not None and c_cas < len(r) else ""
        heb = CAS_TO_HEB.get(cas)
        if heb:
            nd = raw.lower() in ("not detected", "nd", "n.d.") or raw.startswith("<")
            try:
                s["results"][heb] = None if nd else float(raw)
            except ValueError:
                continue
            if not nd:
                s.setdefault("result_text", {})[heb] = raw
            continue
        key = next((v for k, v in _BACTO_CSV_FIELDS.items()
                    if comp.lower() == k or comp.lower().startswith(k + " ")), None)
        if not key:
            continue
        try:
            v = float(raw.replace(",", ""))
        except ValueError:
            continue
        s["field"][key] = v
        if key != "עומק שאיבה":
            s["field_text"][key] = raw
        if key == "עומק פני המים":
            s["water_level"] = v
        elif key == "עומק כללי של הקידוח":
            s["total_depth"] = v
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
    # chlorinated compounds: the codes and names report tables use
    "TECE": "PCE", "PCE": "PCE", "Tetrachloroethene": "PCE",
    "TCE": "TCE", "Trichloroethene": "TCE",
    "CDCE": "cis-1,2-DCE", "cis-1,2-DCE": "cis-1,2-DCE", "cis-DCE": "cis-1,2-DCE",
    "DCEY": "1,1-DCE", "1,1-DCE": "1,1-DCE",
    "VYCL": "VC", "VC": "VC", "Vinyl Chloride": "VC", "ויניל כלוריד": "VC",
    "1,2-DCA": "1,2-DCA", "1,2-Dichloroethane": "1,2-DCA",
    "EDB": "EDB", "Ethylene Dibromide": "EDB",
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


def _find_field_table(doc: Document):
    """
    Locate the current-round field-findings table by content signature (how
    many of its column-0 labels are known field-parameter names), rather
    than assuming it's always doc.tables[1]. Tables nested inside a table
    cell count too (some reports put the field table in a one-cell frame).

    A table with one column per well is preferred over one where a well has
    two columns: the latter is either the rolling two-dates layout or a
    one-off comparison (e.g. before/after an injection), and the one-off
    must never be touched when the real field table exists.
    """
    def well_twice(tbl):
        heads = [c.text.strip() for c in tbl.rows[0].cells] if tbl.rows else []
        heads = [h for h in heads if _well_num(h) and re.search(r'[A-Za-zא-ת]', h)]
        return len(heads) != len(set(heads))

    best, best_key = None, None
    for tbl in _all_tables(doc):
        score = 0
        for row in tbl.rows:
            if not row.cells:
                continue
            label = row.cells[0].text.strip().rstrip("*")
            if label in _FIELD_PARAM_HINTS or label in _FIELD_ROW_ALIASES:
                score += 1
        if score < 3:
            continue
        key = (not well_twice(tbl), score)
        if best_key is None or key > best_key:
            best, best_key = tbl, key
    if best is None and len(doc.tables) > 1:
        return doc.tables[1]  # legacy default
    return best


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
    "עומק הקידוח": "עומק כללי של הקידוח",
    "עומק דגימה מפני המים": "עומק דגימה מפני המים",
    "עכירות": "עכירות",
}


def _fmt_raw(text: str) -> str:
    """A value as the lab printed it ("11.80", "0.5"), with a thousands separator ("1,171")."""
    t = str(text).replace(",", "").strip()
    try:
        v = float(t)
    except ValueError:
        return str(text)
    if abs(v) >= 1000:
        decimals = len(t.split(".")[1]) if "." in t else 0
        return f"{v:,.{decimals}f}"
    return t


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
    tbl = _find_field_table(doc)
    if tbl is None:
        return
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

    # A long field table can be split in parts ("חלק 1 מתוך 3"), each a
    # table of its own with the same rows and other wells.
    for part in _field_table_parts(doc, tbl):
        _update_field_columns(part, field_data, wells)


def _field_table_parts(doc: Document, best) -> list:
    """The field table and its other parts: tables with the same kind of rows,
    one column per well (never a one-off table that shows a well twice)."""
    parts = []
    for tbl in _all_tables(doc):
        if not tbl.rows:
            continue
        labels = {row.cells[0].text.strip().rstrip("*") for row in tbl.rows if row.cells}
        if sum(1 for lb in labels if lb in _FIELD_PARAM_HINTS or lb in _FIELD_ROW_ALIASES) < 3:
            continue
        heads = [c.text.strip() for c in tbl.rows[0].cells[1:]]
        heads = [h for h in heads if _well_num(h) and re.search(r"[A-Za-zא-ת]", h)]
        if not heads or len(heads) != len(set(heads)):
            continue
        parts.append(tbl)
    return parts if best in parts else [best]


def _update_field_columns(tbl, field_data: dict, wells: list):
    FIELD_ROWS = {
        "pH": 1, "EC": 2, "טמפרטורה": 3, "חמצן מומס": 4,
        "עכירות": 5, "רדוקס": 6, "עומק פני המים": 7,
        "עומק כללי של הקידוח": 8, "עומק דגימה מפני המים": 9,
    }
    # Column headers may spell the well differently from the history tables
    # ("חק 2" vs "חק-2"), so match them by well number.
    header = tbl.rows[0]
    num_to_well = {_wkey(w): w for w in wells}
    well_col = {}
    for ci, c in enumerate(header.cells):
        w = num_to_well.get(_wkey(c.text)) if ci > 0 else None
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
        name_to_row.setdefault(_FIELD_ROW_ALIASES.get(row_label, row_label), ri)

    for field in FIELD_ROWS:
        ri = name_to_row.get(field)
        if ri is None or ri >= len(tbl.rows):
            continue
        row = tbl.rows[ri]
        for well, ci in well_col.items():
            sample = field_data.get(well, {})
            val = sample.get("field", {}).get(field)
            if val is not None:
                raw = sample.get("field_text", {}).get(field)
                txt = _fmt_raw(raw) if raw else _fmt_like(row.cells[ci].text, val)
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
            # Parameters as column headers; a field table whose row labels
            # start with "pH" (one column per well) is not this layout.
            names = {"pH", "EC", "pH units", "טמפרטורה", "רדוקס", "חמצן מומס", "עכירות", "מוליכות חשמלית"}
            if len({h for h in headers if h in names}) >= 2:
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
        if not text.startswith(("בתאריך", "בתאריכים")):
            continue
        if not any(trig in text for trig in _CURRENT_ROUND_TRIGGERS):
            continue
        if text.startswith("בתאריכים"):
            # last report covered two rounds: "בתאריכים 07.02.2026 ו-11.02.2026 התבצע..."
            _replace_in_paragraph(
                p, r"בתאריכים\s+[\d./]+(?:\s*(?:,|ו\s*-?)\s*[\d./]+)*", f"בתאריך {date_str}")
        else:
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


def _update_exceedance_bullets(doc: Document, samples: dict, targets: dict):
    """
    The bullets after "להלן פירוט ריכוזי המרכיבים ... החורגים מריכוז היעד":
    one per well with an exceedance this round ("בקידוח פפ-3 אותר בנזן
    בריכוז של 0.370 מ"ג/ליטר;"). Existing "בקידוח ..." bullets right after
    the intro are rewritten, extra ones removed, missing ones added.
    """
    from docx.text.paragraph import Paragraph

    paras = doc.paragraphs
    intro = next((i for i, p in enumerate(paras) if "החורגים" in p.text and "להלן" in p.text), None)
    if intro is None:
        return
    bullets = []
    for p in paras[intro + 1:]:
        t = p.text.strip()
        if t.startswith("בקידוח") or (t.startswith("לא אותרו") and "החורג" in t):
            bullets.append(p)
        elif t:
            break
    sentences = []
    for well in sorted(samples, key=lambda w: int(_well_num(w) or 0)):
        s = samples[well]
        if s.get("floating_layer"):
            continue
        res = s.get("results", {})
        over = [(p, res[p]) for p in COL_PARAMS
                if res.get(p) is not None and p in targets and res[p] > targets[p]]
        if not over:
            continue
        if len(over) == 1:
            p, v = over[0]
            sentences.append(f'בקידוח {well} אותר {p} בריכוז של {_result_text(res, {}, p, s.get("result_text"))[0]} מ"ג/ליטר;')
        else:
            names = _heb_join_words([p for p, _ in over])
            vals = _heb_join_values([_result_text(res, {}, p, s.get("result_text"))[0] for p, _ in over])
            sentences.append(f'בקידוח {well} אותרו {names} בריכוזים של {vals} מ"ג/ליטר בהתאמה;')
    if not sentences:
        sentences = ["לא אותרו ריכוזים החורגים מערכי היעד לשיקום;"]

    if not bullets:
        anchor = paras[intro]
        for sentence in reversed(sentences):
            new_p = copy.deepcopy(anchor._p)
            anchor._p.addnext(new_p)
            np_ = Paragraph(new_p, anchor._parent)
            if np_.runs:
                np_.runs[0].text = sentence
                np_.runs[0].font.highlight_color = WD_COLOR_INDEX.YELLOW
                for r in np_.runs[1:]:
                    r.text = ""
        return
    last = bullets[0]
    for i, sentence in enumerate(sentences):
        if i < len(bullets):
            p = bullets[i]
        else:
            new_p = copy.deepcopy(last._p)
            last._p.addnext(new_p)
            p = Paragraph(new_p, last._parent)
        runs = p.runs
        if runs:
            runs[0].text = sentence
            runs[0].font.highlight_color = WD_COLOR_INDEX.YELLOW
            for r in runs[1:]:
                r.text = ""
        last = p
    for p in bullets[len(sentences):]:
        p._p.getparent().remove(p._p)


def _update_summary_well_bullets(doc: Document, samples: dict, targets: dict = None):
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

    # "להלן פירוט ריכוזי המרכיבים ... החורגים מריכוז היעד": the bullets list
    # only the exceedances of the report's own targets.
    if targets and any("החורגים" in p.text and "להלן" in p.text for p in doc.paragraphs):
        _update_exceedance_bullets(doc, samples, targets)
        return

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
    m = re.search(r'(\d+)\s*\*?\s*$', str(name).strip())
    return str(int(m.group(1))) if m else None


# Wells are matched by number ("חק-2" in a form, "HK-2" at the lab), unless
# the report uses one number under two prefixes ("EAZ-1" and "נ-1"): then
# the prefix counts too, Hebrew letters read as Latin ("נ-1" = "N-1").
_HEB_LAT = {"א": "A", "ב": "B", "ג": "G", "ד": "D", "ה": "H", "ו": "V", "ז": "Z",
            "ח": "H", "ט": "T", "י": "Y", "כ": "K", "ך": "K", "ל": "L", "מ": "M",
            "ם": "M", "נ": "N", "ן": "N", "ס": "S", "ע": "A", "פ": "P", "ף": "P",
            "צ": "TZ", "ץ": "TZ", "ק": "K", "ר": "R", "ש": "SH", "ת": "T"}
_WELL_AMBIG: set = set()


def _well_prefix(name: str) -> str:
    m = re.match(r"\s*(.*?)\d+\s*\*?\s*$", str(name))
    letters = re.sub(r"[^A-Za-zא-ת]", "", m.group(1) if m else "")
    return "".join(_HEB_LAT.get(ch, ch) for ch in letters).upper()


def _set_well_ambiguity(names) -> None:
    """Numbers the report uses under more than one prefix."""
    prefixes: dict = {}
    for n in names:
        num = _well_num(n)
        if num:
            prefixes.setdefault(num, set()).add(_well_prefix(n))
    _WELL_AMBIG.clear()
    _WELL_AMBIG.update(num for num, ps in prefixes.items() if len(ps) > 1)


def _wkey(name: str):
    """The key two spellings of one well share."""
    num = _well_num(name)
    if num is None:
        return None
    return f"{_well_prefix(name)}-{num}" if num in _WELL_AMBIG else num


def _merge_field_form(samples: dict, form: dict) -> list:
    """
    Put the field form's readings into the lab samples, matching wells by
    their number ("חק-2" in the form, "HK-2" at the lab). The form is the
    field record, so its values win over field values printed by the lab.
    Returns the form wells that had no lab sample.
    """
    by_num = {_wkey(k): k for k in samples}
    unmatched = []
    for well, w in form.get("wells", {}).items():
        key = by_num.get(_wkey(well))
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


def _fmt_conc(val, min_decimals: int = 0) -> str:
    """
    Concentration as text without exponent notation: 0.00421, 7.11, 2700.
    min_decimals=3 keeps the lab's three decimals for results ("0.370",
    "0.050"), as the certificates and the report tables print them.
    """
    if val is None:
        return ""
    if val == 0:
        return "0"
    if abs(val) >= 1000:
        return f"{val:.0f}"
    if abs(val) < 1:
        # three significant figures, as the lab prints them
        digits = -int(f"{abs(val):e}".split("e")[1]) + 2
        txt = f"{val:.{max(digits, 0)}f}".rstrip("0").rstrip(".")
        decimals = len(txt.split(".")[1]) if "." in txt else 0
        if decimals < min_decimals:
            txt = f"{val:.{min_decimals}f}"
        return txt
    # 1 to 1000: the digits as reported (7.11, 1.443), never more than four decimals
    return f"{round(val, 4):.4f}".rstrip("0").rstrip(".")


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
    if "," in old or abs(val) >= 1000:
        return f"{val:,.{decimals}f}"
    return f"{val:.{decimals}f}"


def _cell_targets(tbl) -> dict:
    """
    The table's own target row ("ריכוז יעד לשיקום", "ריכוזי יעד לסיום
    ניטור", "תקן מי שתייה" — any label with יעד or תקן, in any column).
    """
    header = None
    for row in tbl.rows[:8]:
        h = _history_header(row)
        if h:
            header = h
            continue
        if header is None:
            continue
        params = header["params"]
        labels = " ".join(c.text for a, b, c in _grid_cells(row) if a not in params)
        if "יעד" not in labels and "תקן" not in labels and "סף" not in labels:
            continue
        out = {}
        for param, c in _row_param_cells(row, header).items():
            try:
                out[param] = float(c.text.strip().replace(",", ""))
            except ValueError:
                pass
        if out:
            return out
    # Older layout: the target row's labels sit in a fixed column order
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


def _result_text(results: dict, limits: dict, param: str, printed: dict = None):
    """
    (text, value) for one compound: '<0.0002' for a non-detect. `printed`
    holds the values as the certificate printed them ("0.370"), used when
    given so the report keeps the lab's digits.
    """
    if param not in results:
        return "-", None
    val = results[param]
    if val is None:
        lim = limits.get(param)
        return (f"<{_fmt_conc(lim)}" if lim else "<0.001"), None
    raw = (printed or {}).get(param)
    if raw:
        try:
            if abs(float(raw) - val) < 1e-12:
                return raw, val
        except ValueError:
            pass
    return _fmt_conc(val), val


def _is_chem_history_table(tbl) -> bool:
    if len(tbl.columns) < 6:
        return False
    head = " ".join(c.text for row in tbl.rows[:3] for c in row.cells)
    return "תאריך" in head and bool(_detect_chem_table_columns(tbl))


# ── Grid-aware table reading ────────────────────────────────────────────────
# Word tables often have rows whose cells span different grid columns
# (a header "כסילן" over two columns, a well name over two columns, rows
# that start with w:gridBefore). python-docx's row.cells is a list per row,
# so the same list index can be a different column in two rows. These
# helpers work by grid column instead.

def _grid_cells(row) -> list:
    """[(first_grid_col, last_grid_col, cell)] for the row's real cells."""
    from docx.table import _Cell
    tr = row._tr
    g = 0
    trPr = tr.trPr
    if trPr is not None:
        gb = trPr.find(qn("w:gridBefore"))
        if gb is not None:
            g = int(gb.get(qn("w:val")))
    out = []
    for tc in tr.findall(qn("w:tc")):
        span = tc.grid_span
        out.append((g, g + span - 1, _Cell(tc, row.table)))
        g += span
    return out


def _cell_at(row, col):
    for a, b, c in _grid_cells(row):
        if a <= col <= b:
            return c
    return None


_DATE_CELL = re.compile(r'^\*?(\d{2}[./]\d{2}[./]\d{2})\*?$')


def _history_header(row):
    """
    A chemistry header row ("שם קידוח | תאריך דיגום | בנזן | ... | MTBE"):
    {"params": {grid_col: param}, "date": grid_col, "well": grid_col}, or None.
    """
    params, date_col, well_col = {}, None, None
    for a, b, c in _grid_cells(row):
        t = c.text.strip()
        if t in _PARAM_CANONICAL:
            params[a] = _PARAM_CANONICAL[t]
        elif "תאריך" in t and date_col is None:
            date_col = a
        elif "קידוח" in t and well_col is None:
            well_col = a
    if len(params) < 2 or date_col is None:
        return None
    if well_col is None:
        well_col = max(0, date_col - 1)
    return {"params": params, "date": date_col, "well": well_col, "row": row}


def _row_param_cells(row, header) -> dict:
    """
    {param: cell} of a data row under `header`. The header's compound cells
    and the row's cells don't always span the same grid columns (the header
    "כסילן" over one column and the row's xylene over two), so when the row
    has one cell per compound in that stretch they are paired by order;
    otherwise (a merged "לא נדגם" row) by the column the header starts at.
    """
    params = sorted(header["params"].items())
    lo = params[0][0]
    hi = max(b for a, b, c in _grid_cells(header["row"]) if a == params[-1][0]) if header.get("row") is not None else params[-1][0]
    cells = [(a, b, c) for a, b, c in _grid_cells(row) if b >= lo and a <= hi]
    if len(cells) == len(params):
        return {p: c for (col, p), (a, b, c) in zip(params, cells)}
    out = {}
    for col, p in params:
        c = _cell_at(row, col)
        if c is not None:
            out[p] = c
    return out


def _history_rows(tbl):
    """
    The dated rows of a chemistry history table, each with the header that
    applies to it (a table can hold several parts, each with its own header):
    [(row_index, well, date_text, header)].
    """
    out, header, well = [], None, None
    for ri, row in enumerate(tbl.rows):
        h = _history_header(row)
        if h:
            header = h
            continue
        if header is None:
            continue
        wc = _cell_at(row, header["well"])
        wt = wc.text.strip() if wc is not None else ""
        if wt and re.search(r'\d', wt) and len(wt) < 20:
            well = wt
        dc = _cell_at(row, header["date"])
        m = _DATE_CELL.match(dc.text.strip()) if dc is not None else None
        if m and well:
            out.append((ri, well, m.group(1).replace("/", "."), header))
    return out


def _all_tables(doc: Document):
    """Every table in the body, including tables nested inside table cells."""
    def walk(tables):
        for t in tables:
            yield t
            seen = set()
            for row in t.rows:
                for c in row.cells:
                    if id(c._tc) in seen:
                        continue
                    seen.add(id(c._tc))
                    if c.tables:
                        yield from walk(c.tables)
    yield from walk(doc.tables)


def _clone_row_after(tbl, after_idx: int, tmpl_idx: int):
    """
    Copy row tmpl_idx and put it after row after_idx. Texts are kept (a "*"
    note column stays), except a cell that starts a vertical merge (the well
    name), which becomes a continuation of the merge above.
    """
    from docx.table import _Row
    tr = copy.deepcopy(tbl.rows[tmpl_idx]._tr)
    for tc in tr.findall(qn("w:tc")):
        tcPr = tc.tcPr
        vm = tcPr.find(qn("w:vMerge")) if tcPr is not None else None
        if vm is not None and vm.get(qn("w:val")) == "restart":
            vm.attrib.pop(qn("w:val"), None)
            for t in tc.iter(qn("w:t")):
                t.text = ""
    tbl.rows[after_idx]._tr.addnext(tr)
    return _Row(tr, tbl)


def _well_history(doc: Document) -> dict:
    """{well: [(date_text, {param: cell_text})]} from every chemistry history table."""
    hist = {}
    for tbl in _all_tables(doc):
        for ri, well, date_txt, header in _history_rows(tbl):
            vals = {p: c.text.strip() for p, c in _row_param_cells(tbl.rows[ri], header).items()}
            hist.setdefault(well, []).append((date_txt, vals))
    return hist


def _update_chem_history_tables(doc: Document, new_date: str, samples: dict, wells_order: list) -> bool:
    """
    One new row per well in the chemistry history table that holds that well
    (one or several tables, e.g. "חלק 1 מתוך 2"). The columns come from the
    header above the row (by grid column, so merged cells don't shift the
    values), the targets from the table's own target row, and a water level
    column ("עומק פני המים") is filled when the table has one. Returns False
    when the document has no such table (older layouts are handled by
    _update_historical_tables).
    """
    tables = [t for t in _all_tables(doc) if _is_chem_history_table(t)]
    if not tables:
        return False
    for well in wells_order:
        sample = samples.get(well)
        if not sample:
            continue
        # Read again for every well: a row added for the previous well in
        # the same table moves the rows below it.
        by_num = {}
        for tbl in tables:
            for ri, w, date_txt, header in _history_rows(tbl):
                by_num.setdefault(_wkey(w), []).append((tbl, ri, w, date_txt, header))
        found = by_num.get(_wkey(well))
        if not found:
            continue
        # A well can sit in two kinds of table (BTEX/MTBE, and chlorinated
        # compounds): one new row in the last table of each kind.
        kinds = {}
        for r in found:
            kinds.setdefault(frozenset(r[4]["params"].values()), []).append(r)
        for rows in kinds.values():
            _add_history_row(rows, sample, new_date)
    return True


def _add_history_row(rows, sample, new_date):
    if True:
        tbl = rows[-1][0]
        rows = [r for r in rows if r[0] is tbl]
        _, last, _, last_date, header = rows[-1]
        targets = _cell_targets(tbl)
        level_col = None
        for row in tbl.rows[:6]:
            for a, b, c in _grid_cells(row):
                if "פני המים" in c.text or "מפלס" in c.text:
                    level_col = a
        # A row with a cell per compound to copy (not a merged "לא נדגם" row)
        tmpl = next((ri for _, ri, _, _, h in reversed(rows)
                     if len({id(c._tc) for c in _row_param_cells(tbl.rows[ri], h).values()}) == len(h["params"])),
                    last)
        prev = tbl.rows[tmpl]
        # A round sampled over several days: each well keeps its own date.
        row_date = sample.get("date") if re.fullmatch(r"\d{2}\.\d{2}\.\d{2}", str(sample.get("date") or "")) else new_date
        style = _exceed_style(tbl, rows, targets)
        if last_date == row_date:
            new_row = tbl.rows[last]
        else:
            new_row = _clone_row_after(tbl, last, tmpl)
        prev_date_txt = _cell_at(prev, header["date"]).text.strip() if _cell_at(prev, header["date"]) is not None else ""
        shown_date = row_date.replace(".", "/") if "/" in prev_date_txt else row_date
        date_cell = _cell_at(new_row, header["date"])
        _set_cell(date_cell, shown_date, highlight=True)
        if level_col is not None and _cell_at(new_row, level_col) is not None:
            old = _cell_at(prev, level_col)
            _set_cell(_cell_at(new_row, level_col),
                      _fmt_like(old.text if old is not None else "", sample.get("water_level")),
                      highlight=True)
        results, limits = sample.get("results", {}), sample.get("limits", {})
        first = header["params"][min(header["params"])]
        for param, cell in _row_param_cells(new_row, header).items():
            if sample.get("floating_layer"):
                _set_cell(cell, "שכבה צפה" if param == first else "", highlight=True)
                continue
            txt, val = _result_text(results, limits, param, sample.get("result_text"))
            exceeds = val is not None and param in targets and val > targets[param]
            _set_cell(cell, txt, bold=exceeds and style["bold"], highlight=True)
            if style["fill"]:
                _set_fill(cell, style["fill"] if exceeds else _fill_of(date_cell))
    return True


def _fill_of(cell):
    tcPr = cell._tc.tcPr
    shd = tcPr.find(qn("w:shd")) if tcPr is not None else None
    fill = shd.get(qn("w:fill")) if shd is not None else None
    return fill if fill and fill.lower() not in ("auto", "ffffff") else None


def _set_fill(cell, fill):
    """Cell background: `fill` ("D9D9D9"), or none."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = tcPr.find(qn("w:shd"))
    if fill is None:
        if shd is not None:
            shd.set(qn("w:fill"), "auto")
        return
    if shd is None:
        from docx.oxml import OxmlElement
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        tcPr.append(shd)
    shd.set(qn("w:fill"), fill)


def _exceed_style(tbl, rows, targets) -> dict:
    """
    How this table marks a result over its limit, read from a result that
    is over it: a gray cell ("התוצאות המסומנות באפור") or bold text.
    """
    from collections import Counter
    fills, bolds = Counter(), Counter()
    for ri, _, _, h in _history_rows(tbl):
        for p, c in _row_param_cells(tbl.rows[ri], h).items():
            try:
                v = float(re.sub(r"[^\d.]", "", c.text) or "x")
            except ValueError:
                continue
            if p in targets and v > targets[p] and not re.search(r"[<>]", c.text):
                fills[_fill_of(c)] += 1
                bolds[any(r.bold for par in c.paragraphs for r in par.runs)] += 1
    if not fills:
        return {"fill": None, "bold": True}
    fill = fills.most_common(1)[0][0]
    bold = bolds.most_common(1)[0][0]
    return {"fill": fill, "bold": bold or not fill}


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
    by_num = {_wkey(w): w for w in samples}
    done = 0
    for tbl in doc.tables:
        well = _is_well_card(tbl)
        if not well:
            continue
        key = by_num.get(_wkey(well))
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
            txt, val = _result_text(results, limits, param, sample.get("result_text"))
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


def _update_sampling_paragraphs(doc: Document, samples: dict, sampler_name: str, sampler_org: str = None):
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
            if sampler_org:
                # "דוגם מעבדת י.ד.ע מים וסביבה." → the lab that sent the sampler
                _replace_in_paragraph(p, r"(דוגם\s+מעבדת)\s+.+?(?=\.\s|\.$)", rf"\1 {sampler_org}")
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


_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_LABEL_TAG = "gw-label:"


def _txbx_docpr(txbx):
    """The wp:docPr of the drawing that holds this text box (also for the VML fallback copy)."""
    el = txbx.getparent()
    while el is not None:
        if el.tag in (qn("wp:anchor"), qn("wp:inline")):
            return el.find(qn("wp:docPr"))
        if el.tag == f"{{{_MC_NS}}}AlternateContent":
            return next(el.iter(qn("wp:docPr")), None)
        el = el.getparent()
    return None


def _txbx_set_text(txbx, text: str):
    ts = list(txbx.iter(qn("w:t")))
    if not ts:
        return
    ts[0].text = text
    for t in ts[1:]:
        t.text = ""
    r = ts[0].getparent()
    rPr = r.find(qn("w:rPr"))
    if rPr is None:
        from docx.oxml import OxmlElement
        rPr = OxmlElement("w:rPr")
        r.insert(0, rPr)
    # a number is left-to-right: an inherited <w:rtl/> shows "<0.001" as "0.001>"
    for rtl in rPr.findall(qn("w:rtl")):
        rPr.remove(rtl)
    hl = rPr.find(qn("w:highlight"))
    if hl is None:
        from docx.oxml import OxmlElement
        hl = OxmlElement("w:highlight")
        rPr.append(hl)
    hl.set(qn("w:val"), "yellow")


def _conc_key(text: str):
    t = text.strip().replace(",", "")
    flag = "<" if t.startswith("<") else ""
    try:
        return flag, round(float(t.lstrip("<")), 9)
    except ValueError:
        return None


def _update_map_labels(doc: Document, new_date: str, samples: dict, history: dict, report: dict) -> int:
    """
    The labels on the site map (שרטוט): the round's date, and per well a
    "MTBE/בנזן" pair such as "0.0117/0.00057". A pair is matched to its well
    by the previous round's values in the history table (the map labels show
    the last round). The match is written into the drawing's hidden title,
    so the next round finds the label even when two wells show the same
    values. Returns how many labels were updated.
    """
    from itertools import permutations
    body = doc.element.body
    boxes = list(body.iter(qn("w:txbxContent")))
    new_dt = datetime.strptime(new_date, "%d.%m.%y")

    # Date: the newest date written on the map, when older than this round
    dates = set()
    for b in boxes:
        for m in re.finditer(r'\b(\d{2}\.\d{2}\.\d{2})\b', "".join(t.text or "" for t in b.iter(qn("w:t")))):
            try:
                d = datetime.strptime(m.group(1), "%d.%m.%y")
            except ValueError:
                continue
            if d < new_dt:
                dates.add((d, m.group(1)))
    shown = any("".join(t.text or "" for t in b.iter(qn("w:t"))).strip() == new_date for b in boxes)
    if dates and not shown:
        old = max(dates)[1]
        for b in boxes:
            txt = "".join(t.text or "" for t in b.iter(qn("w:t")))
            if txt.strip() == old:
                _txbx_set_text(b, new_date)

    # Value pairs
    by_num = {_wkey(w): w for w in samples}
    last_row = {w: rows[-1][1] for w, rows in history.items() if rows}
    params = COL_PARAMS
    found, pending = [], []
    for b in boxes:
        txt = "".join(t.text or "" for t in b.iter(qn("w:t"))).strip()
        m = re.fullmatch(r'(<?[\d.,]+)\s*/\s*(<?[\d.,]+)', txt)
        if not m:
            continue
        docpr = _txbx_docpr(b)
        tag = docpr.get("title", "") if docpr is not None else ""
        if tag.startswith(_LABEL_TAG):
            parts = tag[len(_LABEL_TAG):].split("|")
            if len(parts) == 3:
                found.append((b, docpr, parts[0], parts[1], parts[2]))
                continue
        a, c = _conc_key(m.group(1)), _conc_key(m.group(2))
        cands = set()
        for well, vals in last_row.items():
            for p1, p2 in permutations(params, 2):
                if p1 in vals and p2 in vals and _conc_key(vals[p1]) == a and _conc_key(vals[p2]) == c:
                    cands.add((well, p1, p2))
        pending.append((b, docpr, cands))

    # The order of the pair is the same for every label: take the one the
    # unambiguous labels agree on (MTBE/בנזן by default)
    votes = {}
    for _, _, cands in pending:
        wells = {w for w, _, _ in cands}
        if len(wells) == 1:
            for _, p1, p2 in cands:
                votes[(p1, p2)] = votes.get((p1, p2), 0) + 1
    order = max(votes, key=votes.get) if votes else ("MTBE", "בנזן")
    for b, docpr, cands in pending:
        wells = {w for w, p1, p2 in cands if (p1, p2) == order}
        if len(wells) == 1:
            found.append((b, docpr, wells.pop(), order[0], order[1]))

    done, matched = 0, set()
    for b, docpr, well, p1, p2 in found:
        key = by_num.get(_wkey(well))
        if key is None:
            continue
        s = samples[key]
        res, lim = s.get("results", {}), s.get("limits", {})
        pr = s.get("result_text")
        _txbx_set_text(b, f"{_result_text(res, lim, p1, pr)[0]}/{_result_text(res, lim, p2, pr)[0]}")
        if docpr is not None:
            docpr.set("title", f"{_LABEL_TAG}{well}|{p1}|{p2}")
        matched.add(id(b))
        done += 1
    missed = len({id(d) if d is not None else id(b) for b, d, _ in pending if id(b) not in matched})
    if missed:
        report.setdefault("notes", []).append(
            f"בשרטוט נמצאו {missed} תוויות ריכוז שלא ברור לאיזו באר הן שייכות, ולכן לא עודכנו.")
    return done


def _latest_history_date(doc: Document):
    """The newest date already in the chemistry history (the previous round)."""
    best = None
    for tbl in _all_tables(doc):
        if not _is_chem_history_table(tbl):
            continue
        dates = [d for _, _, d, _ in _history_rows(tbl)]
        if not dates:
            dates = [row.cells[1].text.strip().rstrip("*") for row in tbl.rows if len(row.cells) > 1]
        for t in dates:
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
    paras = doc.paragraphs
    intro = _trend_intro_index(paras)
    if chart_para is None and intro is not None and intro + 1 < len(paras):
        chart_para = paras[intro + 1]
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


def _trend_intro_index(paras):
    """The paragraph that introduces the trend charts ("בגרפים הבאים מוצגת מגמת ...")."""
    return next((i for i, p in enumerate(paras)
                 if "מגמת" in p.text and ("בגרפים" in p.text or "בגרף" in p.text)), None)


def _chart_slots(doc: Document, sheet_order: list = None) -> dict:
    """
    {param: inline drawing} for each caption "גרף מספר N: ... <compound>"
    followed (within a few paragraphs) by a paragraph with a picture.

    Reports without captions put the toolkit pages right after "בגרפים
    הבאים מוצגת מגמת ...": those pictures are taken in the order of the
    sheets in the Mann-Kendall file (sheet_order).
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
    if slots or not sheet_order:
        return slots
    intro = _trend_intro_index(paras)
    if intro is None:
        return slots
    pics = []
    for q in paras[intro + 1:]:
        inls = q._p.findall('.//' + qn('wp:inline'))
        if inls:
            pics.extend(inls)
        elif q.text.strip() and "יש לעדכן את קובץ Mann-Kendall" not in q.text:
            break
    for param, inl in zip(sheet_order, pics):
        slots[param] = inl
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


def _update_mann_kendall(doc: Document, mk_xls_bytes: bytes, rounds: list) -> tuple:
    """
    Add the rounds [(datetime, samples), ...] to the toolkit sheets in
    memory, draw each sheet the report shows, and put the pictures in place
    of the old ones.
    Returns (paste_rows_xlsx_bytes, [params whose picture was replaced], [notes]).
    """
    from soil_lab_tool.core import gw_mann_kendall as mk

    wb = mk.read_mk_workbook(mk_xls_bytes)
    sheets = mk.report_sheets(wb)
    slots = _chart_slots(doc, list(sheets))
    rounds = sorted(rounds, key=lambda r: r[0])
    last_dt = rounds[-1][0]
    notes, updated, added = [], [], []
    for param, sh in sheets.items():
        changed = False
        for when, samples in rounds:
            by_num = {_wkey(w): s for w, s in samples.items()}
            values = {}
            for w in sh.wells:
                s = by_num.get(_wkey(w))
                if s is None or s.get("floating_layer"):
                    values[w] = None  # not sampled this round
                    continue
                res = s.get("results", {})
                values[w] = mk.mk_value(res.get(param), measured=param in res)
            if all(v is None for v in values.values()):
                continue
            if not mk.add_round(sh, when, values):
                notes.append(f"בגיליון {sh.name} בקובץ Mann-Kendall אין שורה פנויה לדיגום {when:%d.%m.%y}.")
                continue
            changed = True
            added.append((sh, when))
        if changed and param in slots:
            stats = mk.sheet_stats(sh, wb.kendall_dist)
            _put_picture(doc, slots[param], mk.render_sheet_png(sh, stats, last_dt))
            updated.append(param)
    paste = mk.paste_rows_xlsx(added) if added else None
    return paste, updated, notes


def apply_mann_kendall(word_bytes: bytes, mk_xls_bytes: bytes, rounds: list, report: dict = None) -> tuple:
    """
    Replace the Mann-Kendall pages of an (already updated) report with the
    toolkit pages including all `rounds` [(datetime, samples), ...].
    Returns (word_bytes, paste_rows_xlsx_bytes | None).
    """
    if report is None:
        report = {}
    report.setdefault("notes", [])
    doc = Document(BytesIO(word_bytes))
    paste, charts, notes = _update_mann_kendall(doc, mk_xls_bytes, rounds)
    report["mk_charts"] = charts
    report["notes"].extend(notes)
    if charts:
        _remove_mk_reminders(doc)
    else:
        report["notes"].append(
            "לא נמצאו בדוח גרפים עם כיתוב 'גרף מספר' לבנזן/MTBE/טולואן, ולכן הגרפים לא הוחלפו.")
        _add_mk_reminder(doc)
    out = BytesIO()
    doc.save(out)
    return out.getvalue(), paste


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
    parsed_round: dict = None,       # optional: {"sampling_date", "samples"} already read
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
        if parsed_round is not None:
            samples = copy.deepcopy(parsed_round["samples"])
            new_date = parsed_round["sampling_date"]
        elif lab_type == "als":
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
        sampler_org = next((s.get("sampler_org") for s in samples.values() if s.get("sampler_org")), None)
        if field_form_bytes:
            from soil_lab_tool.core.gw_round_inputs import parse_yeda_field_form
            # one form, or several for a round sampled over several days
            forms = [parse_yeda_field_form(b) for b in (
                field_form_bytes if isinstance(field_form_bytes, (list, tuple)) else [field_form_bytes])]
            form = {"wells": {}, "sampler_name": None, "sampling_date": None}
            for f in forms:
                form["wells"].update(f.get("wells", {}))
                form["sampler_name"] = form["sampler_name"] or f.get("sampler_name")
                d = f.get("sampling_date")
                if d and (form["sampling_date"] is None or d < form["sampling_date"]):
                    form["sampling_date"] = d
            missing = _merge_field_form(samples, form)
            if missing:
                report["notes"].append("בטופס השדה יש בארות שאין להן תוצאות מעבדה: " + ", ".join(missing))
            if form.get("sampler_name"):
                sampler_name = form["sampler_name"]
                sampler_org = "י.ד.ע מים וסביבה"
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
        _date_pat = re.compile(r'^\*?\d{2}[./]\d{2}[./]\d{2}\*?$')
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

        _set_well_ambiguity(well_depth)
        num_to_doc_well = {}
        num_to_depth = {}
        for w, depth in well_depth.items():
            num = _wkey(w)
            if num:
                if num not in num_to_doc_well or depth > num_to_depth[num]:
                    num_to_doc_well[num] = w
                    num_to_depth[num] = depth

        remapped_samples = {}
        unmatched = []
        for well_key, sample in samples.items():
            doc_well = num_to_doc_well.get(_wkey(well_key))
            if doc_well:
                remapped_samples[doc_well] = sample
            else:
                remapped_samples[well_key] = sample
                unmatched.append(well_key)
        samples = remapped_samples
        wells_order = sorted(samples, key=lambda w: (_well_prefix(w), int(_well_num(w) or 0)))

        if unmatched:
            raise ValueError(
                "לא נמצאה בדוח ה-Word התאמה לבארות הבאות מתעודת המעבדה: "
                + ", ".join(unmatched)
                + ". ודא שמספרי הבארות בתעודה תואמים לבארות בדוח ה-Word."
            )

        _update_field_table(doc, samples, wells_order)
        history_before = _well_history(doc)

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
        _update_map_labels(doc, new_date, samples, history_before, report)

        # ── Narrative ─────────────────────────────────────────────
        full_date_str = new_dt.strftime("%d.%m.%Y")
        combined_results = {}
        for s in samples.values():
            combined_results.update(s.get("results", {}))

        _update_title_month_year(doc, new_dt)
        _update_cover_page_date(doc, new_dt)
        targets = {}
        if layout == "B":
            _update_narrative_layout_b(doc, full_date_str, combined_results)
        else:
            _update_narrative_placeholders(doc, new_date, sampler_name)
            _update_sampling_paragraphs(doc, samples, sampler_name, sampler_org)
            _update_concentration_summary(doc, combined_results)
            targets = {}
            for t in _all_tables(doc):
                if _is_chem_history_table(t):
                    for k, v in _cell_targets(t).items():
                        targets.setdefault(k, v)
            if not _update_exceedance_summary(doc, samples, wells_order, targets):
                _update_summary_well_bullets(doc, samples, targets)
            else:
                _update_summary_well_bullets(doc, samples)

        _update_report_meta(doc, author_name=author_name, submission_date=submission_date)

        # ── Mann-Kendall ──────────────────────────────────────────
        paste_bytes = None
        if mk_xls_bytes:
            paste_bytes, charts, mk_notes = _update_mann_kendall(doc, mk_xls_bytes, [(new_dt, samples)])
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
        report["exceedances"] = [
            (w, p, samples[w]["results"][p], targets[p])
            for w in wells_order for p in targets
            if samples[w].get("results", {}).get(p) is not None
            and samples[w]["results"][p] > targets[p]
        ]
        report["has_targets"] = bool(targets)
        out = BytesIO()
        doc.save(out)
        return out.getvalue(), paste_bytes
