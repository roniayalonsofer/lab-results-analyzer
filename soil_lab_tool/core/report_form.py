"""
The Water Authority report form ("טופס דיווח"), an .xlsx every monitoring
site fills after each round: the sampling date in B3, one column per well
(well names in the row that holds them), field rows (water level,
temperature, turbidity, oxygen, pH, redox, EC) and compound rows found by the
authority's code in column A (2009 MTBE, 84 benzene, ...).

Sites fill it differently, so the form's own style is kept:
  - units: one site writes mg/L under a "µg/L" header, another real µg/L.
    The previous report's last results tell which (the form holds the same
    round): `unit_factor`.
  - non-detects: "<0.20" in one form, the limit as a number in another.
  - the date: a real date, or a text listing every sampling day.

The file is changed in its XML, cell by cell, so the form's drop-down lists
(which openpyxl drops) and links stay as they were.
"""
from __future__ import annotations

import io
import re
import zipfile
from datetime import datetime

from lxml import etree

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_N = "{%s}" % _NS

FIELD_CODES = {"20": "טמפרטורה", "21": "עכירות", "22": "חמצן מומס",
               "25": "pH", "26": "רדוקס", "212": "EC"}
CHEM_CODES = {"2009": "MTBE", "84": "בנזן", "226": "אתיל בנזן",
              "539": "טולואן", "590": "קסילן"}


def _col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _read(form_bytes: bytes):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(form_bytes), data_only=True)
    return wb, wb.worksheets[0]


def _well_columns(ws) -> tuple:
    """(row, {col: well}) for the row that lists the wells from column D on."""
    from soil_lab_tool.gw_report_updater import _well_num
    for r in range(1, min(ws.max_row, 20) + 1):
        cols = {}
        for c in range(4, ws.max_column + 1):
            v = ws.cell(r, c).value
            t = str(v).strip() if v is not None else ""
            if t and _well_num(t) and re.search(r"[A-Za-zא-ת]", t):
                cols[c] = t
            elif cols:
                break
        if len(cols) >= 2 or (cols and r > 3):
            return r, cols
    return None, {}


def _rows(ws) -> dict:
    """{key: row} — field/compound keys by code in column A, plus water level."""
    out = {}
    for r in range(1, ws.max_row + 1):
        code = str(ws.cell(r, 1).value or "").strip()
        label = str(ws.cell(r, 2).value or "")
        if code in FIELD_CODES:
            out[FIELD_CODES[code]] = r
        elif code in CHEM_CODES:
            out[CHEM_CODES[code]] = r
        elif "OIL" in label.upper():
            out["_oil"] = r
        elif ("מים" in label or "מפלס" in label) and str(ws.cell(r, 3).value or "").strip() == "m":
            out["עומק פני המים"] = r
    return out


def _num(v):
    try:
        return float(str(v).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def unit_factor(form_bytes: bytes, previous: dict) -> float:
    """
    1 when the form's compounds are in mg/L, 1000 when in µg/L, found by
    comparing the form (still holding the previous round) with the previous
    report's last results {well: {param: mg/L}}. Without overlap, the unit
    cell decides.
    """
    from soil_lab_tool.gw_report_updater import _wkey
    wb, ws = _read(form_bytes)
    _, cols = _well_columns(ws)
    rows = _rows(ws)
    prev = {_wkey(w): v for w, v in previous.items()}
    ratios = []
    for c, well in cols.items():
        for param in CHEM_CODES.values():
            r = rows.get(param)
            if r is None:
                continue
            f, p = _num(ws.cell(r, c).value), prev.get(_wkey(well), {}).get(param)
            if f and p:
                ratios.append(f / p)
    if ratios:
        ratios.sort()
        mid = ratios[len(ratios) // 2]
        return 1000.0 if mid > 30 else 1.0
    # No report to compare with: a value under 0.01 is mg/L (µg/L forms
    # don't go that low); "<0.20"-style or large values mean µg/L.
    vals = [_num(ws.cell(r, c).value) for p, r in rows.items() if p in CHEM_CODES.values() for c in cols]
    vals = [v for v in vals if v]
    if vals:
        return 1.0 if min(vals) < 0.01 else 1000.0
    unit = " ".join(str(ws.cell(r, 3).value or "") for r in rows.values())
    return 1000.0 if re.search(r"[µμu]g", unit) else 1.0


def _fmt(v: float):
    v = round(v, 6)
    return int(v) if v == int(v) and abs(v) >= 1 else v


# ── XML writing ─────────────────────────────────────────────────────────────

def _cell_el(sheet, ref: str):
    m = re.match(r"([A-Z]+)(\d+)$", ref)
    rnum = m.group(2)
    data = sheet.find(_N + "sheetData")
    row = next((r for r in data.findall(_N + "row") if r.get("r") == rnum), None)
    if row is None:
        row = etree.SubElement(data, _N + "row", r=rnum)
    for c in row.findall(_N + "c"):
        if c.get("r") == ref:
            return c

    def colnum(s):
        n = 0
        for ch in s:
            n = n * 26 + ord(ch) - 64
        return n
    new = etree.Element(_N + "c", r=ref)
    for c in row.findall(_N + "c"):
        if colnum(re.match(r"[A-Z]+", c.get("r")).group(0)) > colnum(m.group(1)):
            c.addprevious(new)
            return new
    row.append(new)
    return new


def _set(sheet, ref: str, value):
    c = _cell_el(sheet, ref)
    for ch in list(c):
        c.remove(ch)
    c.attrib.pop("t", None)
    if value is None or value == "":
        return
    if isinstance(value, (int, float)):
        etree.SubElement(c, _N + "v").text = repr(float(value)) if isinstance(value, float) else str(value)
    else:
        c.set("t", "inlineStr")
        is_ = etree.SubElement(c, _N + "is")
        etree.SubElement(is_, _N + "t").text = str(value)


def update_report_form(form_bytes: bytes, samples: dict, dates: list, factor: float = 1.0,
                       report: dict = None, field_forms: list = None) -> bytes:
    """
    Fill the form with one round: `samples` {well: {"field", "results",
    "limits", "water_level"}} (mg/L), `dates` the round's sampling days,
    `field_forms` the round's Yeda work forms (.xlsx bytes), whose readings
    win over the lab's. Wells of the form not sampled in this round are emptied.
    """
    import copy
    from soil_lab_tool.gw_report_updater import _merge_field_form, _set_well_ambiguity, _wkey
    if report is None:
        report = {}
    samples = copy.deepcopy(samples)
    for fb in field_forms or []:
        from soil_lab_tool.core.gw_round_inputs import parse_yeda_field_form
        _set_well_ambiguity(list(samples))
        _merge_field_form(samples, parse_yeda_field_form(fb))
    wb, ws = _read(form_bytes)
    _, cols = _well_columns(ws)
    if not cols:
        raise ValueError("לא נמצאה בטופס הדיווח שורה עם שמות הקידוחים.")
    rows = _rows(ws)
    _set_well_ambiguity(list(cols.values()))
    by_key = {_wkey(w): s for w, s in samples.items()}

    # how this form writes a non-detect
    lt_style = any(str(ws.cell(r, c).value or "").strip().startswith("<")
                   for p, r in rows.items() if p in CHEM_CODES.values() for c in cols)
    ec_unit = str(ws.cell(rows["EC"], 3).value or "") if "EC" in rows else ""
    ec_div = 1000.0 if re.search(r"m[SC]/cm", ec_unit) else 1.0

    z = zipfile.ZipFile(io.BytesIO(form_bytes))
    sheet_name = "xl/worksheets/sheet1.xml"
    sheet = etree.fromstring(z.read(sheet_name))

    # date
    days = sorted({d for d in dates if d})
    b3 = ws["B3"].value
    if days:
        if isinstance(b3, datetime) or isinstance(b3, (int, float)):
            from soil_lab_tool.core.gw_mann_kendall import date_to_excel
            _set(sheet, "B3", date_to_excel(days[0]))
        else:
            _set(sheet, "B3", ", ".join(f"{d:%d/%m/%y}" for d in days))

    filled, emptied = [], []
    for c, well in cols.items():
        col = _col_letter(c)
        s = by_key.get(_wkey(well))
        if s is None:
            for key, r in rows.items():
                if key != "_oil" and ws.cell(r, c).value is not None:
                    _set(sheet, f"{col}{r}", None)
            emptied.append(well)
            continue
        filled.append(well)
        field = s.get("field", {})
        wl = s.get("water_level", field.get("עומק פני המים"))
        if "עומק פני המים" in rows:
            _set(sheet, f"{col}{rows['עומק פני המים']}", wl)
        for param in FIELD_CODES.values():
            if param in rows:
                v = field.get(param)
                if v is not None and param == "EC":
                    v = round(v / ec_div, 4)
                _set(sheet, f"{col}{rows[param]}", v)
        res, lim = s.get("results", {}), s.get("limits", {})
        for param in CHEM_CODES.values():
            if param not in rows:
                continue
            if param not in res:
                _set(sheet, f"{col}{rows[param]}", None)
                continue
            v = res[param]
            if v is None:
                limit = _fmt((lim.get(param) or 0.001) * factor)
                v = f"<{limit:.2f}" if lt_style and factor > 1 else (f"<{limit}" if lt_style else limit)
            else:
                v = _fmt(v * factor)
            _set(sheet, f"{col}{rows[param]}", v)

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in z.infolist():
            data = z.read(item.filename)
            if item.filename == sheet_name:
                data = etree.tostring(sheet, xml_declaration=True, encoding="UTF-8", standalone=True)
            zo.writestr(item, data)
    report["form_filled"] = filled
    report["form_emptied"] = emptied
    return out.getvalue()


def previous_results(word_bytes: bytes) -> dict:
    """{well: {param: mg/L}} from the last history row of each well in a report."""
    from docx import Document
    from soil_lab_tool.gw_report_updater import _well_history
    hist = _well_history(Document(io.BytesIO(word_bytes)))
    out = {}
    for well, rows in hist.items():
        vals = {}
        for _, row in rows:
            for p, t in row.items():
                ok = t and not re.search(r"[<>]", t) and re.search(r"\d", t)
                v = _num(re.sub(r"[^\d.]", "", t)) if ok else None
                if v is not None:
                    vals[p] = v
        out[well] = vals
    return out
