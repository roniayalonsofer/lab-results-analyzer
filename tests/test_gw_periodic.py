"""Periodic groundwater report: field form, Mann-Kendall, and the Word update.

Small synthetic files that mimic the real layouts (Yeda field form, GSI
Mann-Kendall toolkit, the Hof report tables), so no client data is committed.
"""
import io
import os
import sys
from datetime import datetime

import openpyxl
import pytest
import xlwt
from docx import Document
from docx.shared import Inches

ROOT = os.path.join(os.path.dirname(__file__), "..")
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "soil_lab_tool"))

from soil_lab_tool import gw_report_updater as upd  # noqa: E402
from soil_lab_tool.core import gw_mann_kendall as mk  # noqa: E402
from soil_lab_tool.core import gw_round_inputs as gri  # noqa: E402

HEADER = ["Date", "Client Attributions", "Client » Name", "On site Contact » Name",
          "On site Contact » Phone", "Schedule » User::Full Name", "Location",
          "Quote » Location::Address", "Storage", "Well Form » Weather", "Notes",
          "Exam Form » Supplier::Name", "EXAMINATION::Location",
          "EXAMINATION::Well Header » Drill Depth", "EXAMINATION::Well Header » Drill Diameter",
          "EXAMINATION::Well Header » Sea Depth", "EXAMINATION::Well Header » Extraction Depth",
          "EXAMINATION::Well Header » Floating Layer",
          "EXAMINATION::Well Header » Floating Layer Depth",
          "EXAMINATION::Well Header » Water Volume", "EXAMINATION::Well Header » Water Flow",
          "EXAMINATION::Time", "EXAMINATION::Well » Temperature", "EXAMINATION::Well » PH",
          "EXAMINATION::Well » Conductivity", "EXAMINATION::Well » Cloudness",
          "EXAMINATION::Well » Oxygen", "EXAMINATION::Well » Redox"]


def _field_form() -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(HEADER)
    col = {h: i for i, h in enumerate(HEADER)}

    def row(**kw):
        r = [None] * len(HEADER)
        for k, v in kw.items():
            r[col[k]] = v
        ws.append(r)

    row(**{"Date": datetime(2026, 10, 1), "Schedule » User::Full Name": "דוגם בדיקה"})
    for well, depth, water, readings in [
        ("דיגום באר חק 2", "22.14", "18.23", [(27.4, 7.24, 1066, 34.1, 1.08, 12.7),
                                             (28.7, 7.0, 1295, 4.22, 0.38, -12.7)]),
        ("דיגום באר חק 7", "23.80", "18.06", [(28.2, 6.96, 1217, 10.7, 1.2, 51.8)]),
    ]:
        row(**{"EXAMINATION::Location": well,
               "EXAMINATION::Well Header » Drill Depth": depth,
               "EXAMINATION::Well Header » Sea Depth": water,
               "EXAMINATION::Well Header » Extraction Depth": f"{float(water) + 1:.2f}",
               "EXAMINATION::Well Header » Floating Layer": "לא",
               "EXAMINATION::Well Header » Water Volume": "17.4 ליטר",
               "EXAMINATION::Well Header » Water Flow": "440 מייל"})
        ws.append([None] * len(HEADER))
        for t, (temp, ph, ec, ntu, do, orp) in enumerate(readings):
            row(**{"EXAMINATION::Time": f"9:{50 + t * 5}:00",
                   "EXAMINATION::Well » Temperature": temp, "EXAMINATION::Well » PH": ph,
                   "EXAMINATION::Well » Conductivity": ec, "EXAMINATION::Well » Cloudness": ntu,
                   "EXAMINATION::Well » Oxygen": do, "EXAMINATION::Well » Redox": orp})
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_field_form_takes_last_reading_and_depths():
    b = _field_form()
    assert gri.is_yeda_field_form(b)
    form = gri.parse_yeda_field_form(b)
    assert form["sampler_name"] == "דוגם בדיקה"
    assert form["sampling_date"] == datetime(2026, 10, 1).date()
    w2 = form["wells"]["חק-2"]
    assert w2["field"]["pH"] == 7.0 and w2["field"]["רדוקס"] == -12.7  # last reading, not first
    assert w2["field"]["עומק פני המים"] == 18.23
    assert w2["field"]["עומק כללי של הקידוח"] == 22.14
    assert w2["field"]["עומק דגימה מפני המים"] == 1.0  # below the water, owner's choice
    assert w2["volume_l"] == 17.4 and w2["flow_lpm"] == 0.44
    assert set(form["wells"]) == {"חק-2", "חק-7"}


# ── Mann-Kendall ────────────────────────────────────────────────────────────

def _excel_date(y, m, d):
    return (datetime(y, m, d) - datetime(1899, 12, 30)).days


def _mk_xls(junk_row=True) -> bytes:
    """A GSI toolkit sheet: header block, wells row, 40 event slots, statistics."""
    wb = xlwt.Workbook()
    dist = wb.add_sheet("Kendall_Dist")
    for c, n in enumerate(range(4, 11), start=1):
        dist.write(0, c, n)
    # p for |S| = 0..3 and n = 4..10 (values only need to be plausible here)
    for s in range(0, 4):
        dist.write(s + 1, 0, s)
        for c in range(1, 8):
            dist.write(s + 1, c, round(0.5 - 0.1 * s, 3))
    ws = wb.add_sheet("BENZEN NEW")
    ws.write(3, 1, "Evaluation Date:")
    ws.write(3, 2, "02.06.2026")
    ws.write(4, 1, "Facility Name:")
    ws.write(4, 2, "תחנה")
    ws.write(4, 6, "Constituent:")
    ws.write(4, 7, "BENZEN")
    ws.write(5, 6, "Concentration Units:")
    ws.write(5, 7, "mg/L")
    ws.write(7, 2, "Sampling Point ID:")
    ws.write(7, 3, "חק-2")
    ws.write(7, 4, "חק-7")
    data = [((2025, 1, 1), 0.5, 3.0), ((2025, 6, 1), 0.3, 2.0), ((2025, 10, 1), 0.2, 1.0),
            ((2026, 2, 1), 0.1, 0.5), ((2026, 6, 2), 0.01, 0.743)]
    for i in range(40):
        ws.write(14 + i, 1, i + 1)
    for i, (d, a, b) in enumerate(data):
        ws.write(14 + i, 2, _excel_date(*d))
        ws.write(14 + i, 3, a)
        ws.write(14 + i, 4, b)
    if junk_row:
        ws.write(53, 2, _excel_date(2024, 2, 6))  # a leftover date with no values
    ws.write(54, 2, "Coefficient of Variation:")
    ws.write(55, 2, "Mann-Kendall Statistic (S):")
    ws.write(56, 2, "Confidence Factor:")
    ws.write(57, 2, "Concentration Trend:")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_mk_statistics_and_trend_rules():
    assert mk.mk_statistic([5, 4, 3, 2, 1]) == -10
    assert mk.mk_statistic([1, 1, 1, 1]) == 0
    assert mk.trend(-10, 0.99, 1.5) == "Decreasing"
    assert mk.trend(-5, 0.92, 1.5) == "Prob. Decreasing"
    assert mk.trend(-1, 0.6, 0.5) == "Stable"
    assert mk.trend(-1, 0.6, 1.5) == "No Trend"
    assert mk.trend(8, 0.97, 0.5) == "Increasing"
    assert mk.trend(0, None, None) == ""  # fewer than four results
    assert mk.mk_value(None) == 0.01 and mk.mk_value(0.004) == 0.01 and mk.mk_value(0.3) == 0.3
    assert mk.mk_value(0.3, measured=False) == "-"


def test_mk_reads_sheet_skips_leftover_date_and_adds_round():
    wb = mk.read_mk_workbook(_mk_xls())
    sh = mk.report_sheets(wb)["בנזן"]
    assert sh.wells == ["חק-2", "חק-7"]
    assert len(sh.events) == 5  # the leftover 2024 date without values is ignored
    assert sh.next_free_row == 19
    when = datetime(2026, 10, 1)
    assert mk.add_round(sh, when, {"חק-2": 0.01, "חק-7": 0.01})
    assert sh.event_rows[-1] == 19
    assert mk.add_round(sh, when, {"חק-2": 0.01, "חק-7": 0.02})  # same date: replaced
    assert len(sh.events) == 6 and sh.events[-1][1]["חק-7"] == 0.02
    stats = mk.sheet_stats(sh, wb.kendall_dist)
    assert stats["חק-7"]["s"] == mk.mk_statistic([3.0, 2.0, 1.0, 0.5, 0.743, 0.02])
    png = mk.render_sheet_png(sh, stats, when)
    assert png[:8] == b"\x89PNG\r\n\x1a\n"
    rows = openpyxl.load_workbook(io.BytesIO(mk.paste_rows_xlsx([sh], when))).active
    values = [[c for c in r if c is not None] for r in rows.iter_rows(values_only=True)]
    assert ["גיליון: BENZEN NEW", "שורה 20"] in values
    assert [6, "01/10/2026", 0.01, 0.02] in values


def test_bidi_labels_for_charts(monkeypatch):
    # matplotlib with libraqm orders Hebrew itself: the text goes in as is
    monkeypatch.setattr(mk, "_mpl_does_bidi", lambda: True)
    assert mk._bidi("פפ-1") == "פפ-1"
    monkeypatch.setattr(mk, "_mpl_does_bidi", lambda: False)
    assert mk._bidi("חק-2") == "2-קח"
    assert mk._bidi("סונול החוף") == "ףוחה לונוס"
    assert mk._bidi("MTBE") == "MTBE"


# ── Word update ─────────────────────────────────────────────────────────────

def _png() -> bytes:
    from PIL import Image
    buf = io.BytesIO()
    Image.new("RGB", (40, 50), "white").save(buf, format="PNG")
    return buf.getvalue()


def _report() -> bytes:
    doc = Document()
    doc.add_paragraph("דו\"ח ניטור מי תהום תקופתי, יוני 2026- תחנה")
    doc.add_paragraph("בתאריך 02.06.26 נערך בתחנה דיגום תקופתי של מי התהום מקידוחים חק-2 ו- חק-7 "
                      "על ידי מר דוגם קודם, דוגם חברת י.ד.ע מים.")
    doc.add_paragraph("כמות המים שנשאבה מהקידוחים נעה בין כ- 16 ל-22 ליטר לכל קידוח בספיקה של 0.4 ליטר בדקה.")

    t = doc.add_table(rows=5, cols=4)
    for ci, h in enumerate(["בדיקה", "יחידות", "חק 2", "חק 7"]):
        t.rows[0].cells[ci].text = h
    for ri, (lab, unit, a, b) in enumerate([("pH", "-", "6.80", "6.32"),
                                            ("רדוקס", "mv", "113", "10.0"),
                                            ("עומק פני המים", "m", "18.06", "17.89"),
                                            ("עומק דגימה מפני המים*", "m", "1.00", "1.00")], 1):
        for ci, v in enumerate([lab, unit, a, b]):
            t.rows[ri].cells[ci].text = v

    card = doc.add_table(rows=5, cols=3)
    a = card.rows[0].cells[0].merge(card.rows[0].cells[2])
    a.text = "חק-7"
    d = card.rows[1].cells[0].merge(card.rows[1].cells[2])
    d.text = "02.06.26"
    for ci, v in enumerate(["תרכובת", "ריכוז יעד לשיקום [מ\"ג/ל']", "ריכוז [מ\"ג/ל']"]):
        card.rows[2].cells[ci].text = v
    for ci, v in enumerate(["MTBE", "0.74", "0.00665"]):
        card.rows[3].cells[ci].text = v
    for ci, v in enumerate(["בנזן", "0.094", "0.743"]):
        card.rows[4].cells[ci].text = v

    chem = doc.add_table(rows=5, cols=9)
    heads = [["קידוח", "תאריך דיגום", "עומק פני המים (מ')", "", "MTBE", "בנזן", "טולואן", "אתיל בנזן", "כסילן"],
             ["קידוח", "תאריך דיגום", "עומק פני המים (מ')", "יחידות", "מ''ג/ל", "מ''ג/ל", "מ''ג/ל", "מ''ג/ל", "מ''ג/ל"],
             ["קידוח", "תאריך דיגום", "עומק פני המים (מ')", "ריכוז יעד לשיקום^", "0.748", "0.0935", "13.09", "5.61", "9.35"],
             ["חק-7", "05.02.26", "17.74", "", "0.00049", "0.0854", "0.00206", "0.00708", "0.00256"],
             ["חק-7", "02.06.26", "17.89", "", "0.00665", "0.743", "0.0225", "0.0439", "0.0465"]]
    for ri, row in enumerate(heads):
        for ci, v in enumerate(row):
            chem.rows[ri].cells[ci].text = v
    chem2 = doc.add_table(rows=4, cols=9)
    for ri, row in enumerate(heads[:3] + [["חק-2", "02.06.26", "18.06", "", "0.608", "0.00028", "0.00454", "0.00903", "0.00769"]]):
        for ci, v in enumerate(row):
            chem2.rows[ri].cells[ci].text = v

    doc.add_paragraph("גרף מספר 1: שינוי ריכוזי בנזן כנגד הזמן")
    doc.add_paragraph().add_run().add_picture(io.BytesIO(_png()), width=Inches(4))
    doc.add_paragraph("סיכום")
    doc.add_paragraph("בתאריך 02.06.26 נערך בתחנה דיגום תקופתי של מי התהום מקידוחים חק-2 ו- חק-7;")
    doc.add_paragraph("בקידוחים חק-2 לא אותרו ריכוזי BTEX MTBE החורגים מערכי היעד לשיקום;")
    doc.add_paragraph("בקידוח חק-7 אותר בנזן בריכוז 0.743 מ\"ג/ל' אשר חורג מערך היעד לשיקום (0.0935 מ\"ג/ליטר).")
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


@pytest.fixture
def als_round(monkeypatch):
    rnd = {"sampling_date": "01.10.26", "samples": {
        "HK-2": {"results": {"MTBE": 0.115, "בנזן": None, "טולואן": 0.0161, "אתיל בנזן": 0.0639,
                             "קסילן": 0.0617},
                 "limits": {"בנזן": 0.0002}},
        "HK-7": {"results": {"MTBE": None, "בנזן": 0.2, "טולואן": 0.00191, "אתיל בנזן": 0.00089,
                             "קסילן": 0.00123},
                 "limits": {"MTBE": 0.0002}},
    }}
    monkeypatch.setattr(gri, "parse_als_round", lambda b, name="": {
        "sampling_date": rnd["sampling_date"],
        "samples": {k: {"results": dict(v["results"]), "limits": dict(v["limits"])}
                    for k, v in rnd["samples"].items()}})
    return rnd


def _texts(doc):
    return [p.text for p in doc.paragraphs]


def test_word_update_with_field_form_and_mk(als_round):
    rep = {}
    out, paste = upd.run_update_bytes(_report(), b"als", _mk_xls(), lab_type="als",
                                      field_form_bytes=_field_form(), report=rep)
    doc = Document(io.BytesIO(out))
    field, card, chem, chem2 = doc.tables

    # table 1: last field reading, matched to "חק 2" by number, measured digits kept
    assert field.rows[1].cells[2].text == "7.00"
    assert field.rows[2].cells[2].text == "-12.7"
    assert field.rows[3].cells[3].text == "18.06"

    # per-well card: date, "<limit" for a non-detect, bold over the card's own target
    assert card.rows[1].cells[0].text == "01.10.26"
    assert card.rows[3].cells[2].text == "<0.0002"
    assert card.rows[4].cells[2].text == "0.2"
    assert card.rows[4].cells[2].paragraphs[0].runs[0].bold is True

    # chemistry history: a new row with the water level from the form
    last = [c.text for c in chem.rows[-1].cells]
    assert last[1:3] == ["01.10.26", "18.06"]
    assert last[4:] == ["<0.0002", "0.2", "0.00191", "0.00089", "0.00123"]
    assert chem.rows[-1].cells[5].paragraphs[0].runs[0].bold is True   # 0.2 > 0.0935
    assert not chem.rows[-1].cells[6].paragraphs[0].runs[0].bold
    assert [c.text for c in chem2.rows[-1].cells][1:5] == ["01.10.26", "18.23", "", "0.115"]

    texts = _texts(doc)
    assert any("מר דוגם בדיקה, דוגם" in t for t in texts)
    assert any("נעה בין כ- 17 ל-17 ליטר" in t or "עמדה על כ- 17 ליטר" in t for t in texts)
    assert any("בספיקה של 0.44 ליטר" in t for t in texts)
    assert any(t.startswith("בתאריך 01.10.26 נערך") for t in texts)
    assert "בקידוח חק-2 לא אותרו ריכוזי BTEX ו-MTBE החורגים מערכי היעד לשיקום;" in texts
    assert any(t.startswith("בקידוח חק-7 אותר בנזן בריכוז 0.2") for t in texts)

    # MK picture replaced, no reminder note, rows to paste returned
    assert rep["mk_charts"] == ["בנזן"]
    assert not any("Mann-Kendall" in t for t in texts)
    assert paste is not None

    # running the same round again changes nothing
    out2, _ = upd.run_update_bytes(out, b"als", _mk_xls(), lab_type="als",
                                   field_form_bytes=_field_form())
    doc2 = Document(io.BytesIO(out2))
    assert [len(t.rows) for t in doc2.tables] == [len(t.rows) for t in doc.tables]
    assert _texts(doc2) == texts


def test_without_mk_the_reminder_is_added_once(als_round):
    out, paste = upd.run_update_bytes(_report(), b"als", None, lab_type="als")
    out, _ = upd.run_update_bytes(out, b"als", None, lab_type="als")
    texts = _texts(Document(io.BytesIO(out)))
    assert paste is None
    assert sum("יש לעדכן את קובץ Mann-Kendall" in t for t in texts) == 1


def test_fmt_like_never_drops_measured_digits():
    assert upd._fmt_like("113", -12.7) == "-12.7"
    assert upd._fmt_like("6.80", 7.0) == "7.00"
    assert upd._fmt_like("1,382.00", 1295.0) == "1,295.00"
    assert upd._fmt_conc(0.00421) == "0.00421" and upd._fmt_conc(7.11) == "7.11"
    assert upd._fmt_conc(0.0002) == "0.0002" and upd._fmt_conc(2700) == "2700"
    assert upd._fmt_conc(1.443) == "1.443" and upd._fmt_conc(95.0) == "95"


def test_xylene_total_from_its_parts_when_the_certificate_has_no_sum():
    wb = openpyxl.load_workbook(io.BytesIO(_als_xlsx()))
    ws = wb.active
    ws.append(["BTEX"])
    ws.append(["meta- & para-Xylene", "W-VOCGMS01", "µg/L", "0.20", "1240", "<0.20"])
    ws.append(["ortho-Xylene", "W-VOCGMS01", "µg/L", "0.10", "203", "<0.10"])
    buf = io.BytesIO()
    wb.save(buf)
    rnd = gri.parse_als_round(buf.getvalue())
    assert rnd["samples"]["HK-2"]["results"]["קסילן"] == pytest.approx(1.443)
    assert rnd["samples"]["HK-7"]["results"]["קסילן"] is None
    assert rnd["samples"]["HK-7"]["limits"]["קסילן"] == pytest.approx(0.0003)


# ── sorting the round's files ───────────────────────────────────────────────

def _als_xlsx(samples=("HK-2", "HK-7"), date="01/10/2026") -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Client WATER 1"
    ws["A2"] = "CLIENT"
    ws["A4"], ws["B4"] = "Work Order:", "PR00C0000"
    ws["D8"] = "Client Sample ID"
    ws["D9"] = "Laboratory Sample ID"
    ws["D10"] = "Client Sampling Date"
    for i, s in enumerate(samples):
        ws.cell(row=8, column=5 + i, value=s)
        ws.cell(row=9, column=5 + i, value=f"PR00C000000{i + 1}")
        ws.cell(row=10, column=5 + i, value=date)
    ws["A11"] = "RESULTS OF ANALYSIS"
    for ci, h in enumerate(["Parameter", "Method", "Unit", "LOR"], 1):
        ws.cell(row=13, column=ci, value=h)
    rows = [("BTEX",),
            ("Benzene", "W-VOCGMS01", "µg/L", "0.20") + ("<0.20", "4.21")[:len(samples)],
            ("Non-Halogenated Volatile Organic Compounds",),
            ("Methyl tert-Butyl Ether (MTBE)", "W-VOCGMS01", "µg/L", "0.20") + ("115", "<0.20")[:len(samples)]]
    for ri, row in enumerate(rows, 15):
        for ci, v in enumerate(row, 1):
            ws.cell(row=ri, column=ci, value=v)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _empty_xlsx() -> bytes:
    buf = io.BytesIO()
    openpyxl.Workbook().save(buf)
    return buf.getvalue()


def test_round_files_are_sorted_without_choosing_a_lab():
    from soil_lab_tool.core.gw_round_files import sort_round_files
    res = sort_round_files([
        ("נספח.XLSX", _als_xlsx()),
        ("טופס ביצוע עבודה.xlsx", _field_form()),
        ("mann kendall.xls", _mk_xls()),
        ("משהו אחר.xlsx", _empty_xlsx()),
    ])
    assert len(res["rounds"]) == 1
    r = res["rounds"][0]
    assert (r.lab, r.date_str, r.wells) == ("als", "01.10.26", ["HK-2", "HK-7"])
    assert r.samples["HK-2"]["results"]["MTBE"] == pytest.approx(0.115)
    assert r.samples["HK-2"]["results"]["בנזן"] is None
    assert r.field_form[0] == "טופס ביצוע עבודה.xlsx"
    assert res["mk"][0] == "mann kendall.xls"
    assert [n for n, _ in res["ignored"]] == ["משהו אחר.xlsx"]


def test_lone_sample_named_after_the_site_takes_the_form_well():
    from soil_lab_tool.core.gw_round_files import sort_round_files
    form = openpyxl.load_workbook(io.BytesIO(_field_form()))
    ws = form.active
    # keep only the first well (header, round row, חק 2 block)
    ws.delete_rows(7, ws.max_row)
    buf = io.BytesIO()
    form.save(buf)
    res = sort_round_files([("a.xlsx", _als_xlsx(samples=("Hof Station",))),
                            ("form.xlsx", buf.getvalue())])
    r = res["rounds"][0]
    assert r.wells == ["חק-2"]
    assert res["notes"]


def test_word_update_from_a_sorted_round_with_mk_over_two_rounds():
    from soil_lab_tool.core.gw_round_files import sort_round_files
    res = sort_round_files([("n.XLSX", _als_xlsx()), ("f.xlsx", _field_form())])
    r = res["rounds"][0]
    rep = {}
    out, _ = upd.run_update_bytes(_report(), None, None, lab_type=r.lab,
                                  field_form_bytes=r.field_form[1], report=rep,
                                  parsed_round={"sampling_date": r.date_str, "samples": r.samples})
    assert rep["has_targets"] and rep["exceedances"] == []
    earlier = (datetime(2026, 8, 1), {"HK-7": {"results": {"בנזן": 0.5}}})
    out, paste = upd.apply_mann_kendall(out, _mk_xls(), [earlier, (r.date, r.samples)], report=rep)
    assert rep["mk_charts"] == ["בנזן"]
    rows = openpyxl.load_workbook(io.BytesIO(paste)).active
    values = [[c for c in row if c is not None] for row in rows.iter_rows(values_only=True)]
    assert [6, "01/08/2026", "", 0.5] in values or [6, "01/08/2026", 0.5] in values
    assert any(v[:2] == [7, "01/10/2026"] for v in values if len(v) >= 2)


# ── Report with merged cells, a nested field table and a one-off table ─────

def _parbar_like() -> Document:
    """
    Mimics the Parbar report: a one-off before/after table with each well
    twice (must stay as is), the real field table nested in a one-cell
    frame, and history tables whose header and rows span different grid
    columns (header MTBE over two columns, a row's xylene over two, a well
    name over two columns).
    """
    doc = Document()
    doc.add_paragraph("טבלה מספר 4 – ממצאי שדה לפני ההחדרה ואחריה")
    once = doc.add_table(rows=3, cols=6)
    for ci, t in enumerate(["פרמטר", "יחידות", "פפ-1", "פפ-1", "פפ-2", "פפ-2"]):
        once.cell(0, ci).text = t
    for ci, t in enumerate(["תאריך דיגום", "", "14.12.25", "07.01.26", "14.12.25", "07.01.26"]):
        once.cell(1, ci).text = t
    for ci, t in enumerate(["pH", "-", "6.71", "7.33", "6.89", "8.1"]):
        once.cell(2, ci).text = t
    for label in ("חמצן מומס", "רדוקס", "עכירות", "טמפרטורה"):
        cells = once.add_row().cells
        cells[0].text = label
        for ci in range(2, 6):
            cells[ci].text = "1"

    doc.add_paragraph("ממצאי בדיקות השדה מוצגים בטבלה מספר 6:")
    frame = doc.add_table(rows=1, cols=1)
    inner = frame.cell(0, 0).add_table(rows=1, cols=4)
    for ci, t in enumerate(["בדיקה", "יחידות", "פפ-1", "פפ-2"]):
        inner.cell(0, ci).text = t
    for label, unit, a, b in (("pH", "-", "6.71", "6.67"), ("EC", "μS/cm", "223", "773"),
                              ("רדוקס", "mV", "34", "-35"), ("עומק פני המים", "m", "11.78", "11.86"),
                              ("עומק הקידוח", "m", "17.97", "18.01")):
        cells = inner.add_row().cells
        for ci, t in enumerate((label, unit, a, b)):
            cells[ci].text = t

    for well, rows in (("פפ-1", [("14.12.25", ["0.001", "<0.001", "<0.001", "0.001", "<0.001"]),
                                 ("11.02.26", ["0.00057", "0.0051", "0.124", "0.246", "0.0117"])]),
                       ("פפ-2", [("11.02.26", ["0.0785", "0.00631", "0.031", "0.0209", "0.00201"])])):
        doc.add_paragraph(f"טבלה מספר 7 – {well}")
        t = doc.add_table(rows=3, cols=9)
        head = ["שם קידוח", "תאריך דיגום", "", "בנזן", "טולואן", "אתיל בנזן", "כסילן", "MTBE", ""]
        for ci, x in enumerate(head):
            t.cell(0, ci).text = x
        t.cell(0, 7).merge(t.cell(0, 8))             # MTBE over two columns
        t.cell(1, 2).text = "ערך יעד^"
        for ci, x in enumerate(["0.1", "14", "6", "10", "0.8"]):
            t.cell(1, 3 + ci).text = x
        t.cell(1, 7).merge(t.cell(1, 8))
        t.cell(2, 0).text = well
        for i, (date, vals) in enumerate(rows):
            r = t.rows[2] if i == 0 else t.add_row()
            ri = len(t.rows) - 1 if i else 2
            t.cell(ri, 1).text = date
            for ci, x in enumerate(vals[:3]):
                t.cell(ri, 3 + ci).text = x
            t.cell(ri, 6).merge(t.cell(ri, 7))       # xylene over two columns
            t.cell(ri, 6).text = vals[3]
            t.cell(ri, 8).text = vals[4]
    doc.add_paragraph("סיכום")
    doc.add_paragraph("בתאריכים 07.02.2026 ו-11.02.2026 התבצע דיגום תקופתי של מי התהום בקידוחים פפ-1 ו-פפ-2;")
    doc.add_paragraph("להלן פירוט ריכוזי המרכיבים שאותרו באירוע ניטור זה החורגים מריכוז היעד לשיקום:")
    doc.add_paragraph('בקידוח פפ-2 אותר בנזן בריכוז של 0.364 ו-0.112 מ"ג לליטר בתאריכים 07.02.2026 ו-11.02.2026 בהתאמה.')
    doc.add_paragraph("המלצה")
    return doc


def _june_samples():
    return {
        "פפ-1": {"results": {"בנזן": None, "טולואן": None, "אתיל בנזן": None, "קסילן": 0.013, "MTBE": None},
                 "result_text": {"קסילן": "0.013"},
                 "field": {"pH": 7.09, "EC": 601.0, "רדוקס": -75.0, "עומק פני המים": 11.8,
                           "עומק כללי של הקידוח": 17.7},
                 "field_text": {"pH": "7.09", "EC": "601", "רדוקס": "-75", "עומק פני המים": "11.80",
                                "עומק כללי של הקידוח": "17.70"}},
        "פפ-2": {"results": {"בנזן": 0.37, "טולואן": 0.023, "אתיל בנזן": 0.05, "קסילן": 0.079, "MTBE": 0.07},
                 "result_text": {"בנזן": "0.370", "טולואן": "0.023", "אתיל בנזן": "0.050",
                                 "קסילן": "0.079", "MTBE": "0.070"},
                 "field": {"pH": 6.94, "EC": 1171.0, "רדוקס": -187.0, "עומק פני המים": 11.91,
                           "עומק כללי של הקידוח": 17.71},
                 "field_text": {"pH": "6.94", "EC": "1171", "רדוקס": "-187", "עומק פני המים": "11.91",
                                "עומק כללי של הקידוח": "17.71"}},
    }


def test_merged_cells_nested_field_table_and_one_off_table():
    doc = _parbar_like()
    samples = _june_samples()
    wells = ["פפ-1", "פפ-2"]
    once_before = [[c.text for c in r.cells] for r in doc.tables[0].rows]

    upd._update_field_table(doc, samples, wells)
    assert upd._update_chem_history_tables(doc, "10.06.26", samples, wells)
    targets = upd._cell_targets(doc.tables[2])
    assert targets == {"בנזן": 0.1, "טולואן": 14.0, "אתיל בנזן": 6.0, "קסילן": 10.0, "MTBE": 0.8}
    upd._update_narrative_placeholders(doc, "10.06.26", None)
    upd._update_summary_well_bullets(doc, samples, targets)

    # the one-off before/after table is untouched
    assert [[c.text for c in r.cells] for r in doc.tables[0].rows] == once_before
    # the nested field table took the round, as the lab printed it
    inner = doc.tables[1].cell(0, 0).tables[0]
    rows = {r.cells[0].text: [c.text for c in r.cells[2:]] for r in inner.rows}
    assert rows["pH"] == ["7.09", "6.94"]
    assert rows["EC"] == ["601", "1,171"]
    assert rows["עומק פני המים"] == ["11.80", "11.91"]
    assert rows["עומק הקידוח"] == ["17.70", "17.71"]
    # history: each value in its own column despite the merged cells
    hist = upd._well_history(doc)
    assert hist["פפ-1"][-1] == ("10.06.26", {"בנזן": "<0.001", "טולואן": "<0.001", "אתיל בנזן": "<0.001",
                                             "קסילן": "0.013", "MTBE": "<0.001"})
    assert hist["פפ-2"][-1] == ("10.06.26", {"בנזן": "0.370", "טולואן": "0.023", "אתיל בנזן": "0.050",
                                             "קסילן": "0.079", "MTBE": "0.070"})
    assert hist["פפ-1"][-2][1]["MTBE"] == "0.0117"  # the previous row is unchanged
    # summary: the round's date and only the exceedances
    texts = [p.text for p in doc.paragraphs]
    assert "בתאריך 10.06.26 התבצע דיגום תקופתי של מי התהום בקידוחים פפ-1 ו-פפ-2;" in texts
    i = next(i for i, t in enumerate(texts) if t.startswith("להלן"))
    assert texts[i + 1] == 'בקידוח פפ-2 אותר בנזן בריכוז של 0.370 מ"ג/ליטר;'
    assert texts[i + 2] == "המלצה"


def test_mk_pages_found_after_the_trend_sentence_without_captions():
    doc = Document()
    doc.add_paragraph("בגרפים הבאים מוצגת מגמת השתנות הפרמטרים העיקריים כנגד הזמן בקידוחים:")
    doc.add_paragraph().add_run().add_picture(io.BytesIO(_png()), width=Inches(1))
    doc.add_paragraph().add_run().add_picture(io.BytesIO(_png()), width=Inches(1))
    doc.add_paragraph("סיכום והמלצות")
    slots = upd._chart_slots(doc, ["בנזן", "MTBE"])
    assert list(slots) == ["בנזן", "MTBE"]
    assert slots["בנזן"] is not slots["MTBE"]


def test_two_wells_in_one_history_table_each_get_their_row_at_the_end():
    doc = Document()
    t = doc.add_table(rows=1, cols=8)
    for ci, x in enumerate(["קידוח", "תאריך דיגום", "", "MTBE", "בנזן", "טולואן", "אתיל בנזן", "כסילן"]):
        t.cell(0, ci).text = x
    for well, dates in (("חק-2", ["05.02.26", "02.06.26"]), ("חק-3", ["05.02.26", "02.06.26"])):
        for i, d in enumerate(dates):
            cells = t.add_row().cells
            cells[0].text = well if i == 0 else ""
            cells[1].text = d
            for ci in range(3, 8):
                cells[ci].text = "0.001"
    samples = {w: {"results": {"MTBE": 0.002, "בנזן": None, "טולואן": None, "אתיל בנזן": None, "קסילן": None}}
               for w in ("חק-2", "חק-3")}
    upd._update_chem_history_tables(doc, "01.10.26", samples, ["חק-2", "חק-3"])
    assert [r.cells[1].text for r in t.rows[1:]] == [
        "05.02.26", "02.06.26", "01.10.26", "05.02.26", "02.06.26", "01.10.26"]
