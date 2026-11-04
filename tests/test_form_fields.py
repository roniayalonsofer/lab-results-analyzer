"""Regression tests for typing values from scanned field forms, and soil gas sample names.

Run:  py -3 -m pytest tests -q
"""
import io
import os
import sys

import fitz
import openpyxl

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "soil_lab_tool"))
sys.path.insert(0, os.path.dirname(__file__))

from core.form_fields import apply_form_values, custody_page, page_count, page_png  # noqa: E402


def _pdf(*page_texts) -> bytes:
    doc = fitz.open()
    for text in page_texts:
        doc.new_page().insert_text((72, 72), text)
    raw = doc.tobytes()
    doc.close()
    return raw


def test_scan_with_no_text_opens_on_the_second_to_last_page():
    assert custody_page(_pdf("", "", "", "")) == 3
    assert custody_page(_pdf("")) == 1


def test_page_helpers():
    raw = _pdf("one", "two")
    assert page_count(raw) == 2
    assert page_png(raw, 2).startswith(b"\x89PNG")


def test_custody_page_is_found_by_its_title():
    from unittest import mock

    class _Page:
        def __init__(self, text):
            self.text = text

        def get_text(self):
            return self.text

    class _Doc(list):
        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    doc = _Doc([_Page("a"), _Page("טופס שרשרת משמורת"), _Page("b"), _Page("c")])
    with mock.patch("core.form_fields.fitz.open", return_value=doc):
        assert custody_page(b"") == 2
    doc = _Doc([_Page("a"), _Page("b"), _Page("יביטקא עקרק זג םוגידל תרומשמ תרשרש ספוט"), _Page("c")])
    with mock.patch("core.form_fields.fitz.open", return_value=doc):
        assert custody_page(b"") == 3


def test_typed_values_reach_the_soil_gas_sheet(tmp_path):
    from core.excel_output import LabReportExcel
    from core.threshold_manager import ThresholdManager
    th = os.path.join(os.path.dirname(__file__), "..", "soil_lab_tool", "thresholds")
    tm = ThresholdManager(os.path.join(th, "soil_vsl_tier1_v7_2024.xlsx"),
                          vsl_full_path=os.path.join(th, "soil_vsl_v7_full.xlsx"))
    records = [dict(compound=c, cas=cas, value=120.0, flag="", unit="µg/m³", sample_id=s,
                    analysis_type="SOIL_GAS_VOC", canister_num=n, lod=10.0, loq=50.0)
               for c, cas in (("Benzene", "71-43-2"), ("Toluene", "108-88-3"))
               for s, n in (("גק-1", "8610"), ("גק-4", "11922"))]
    apply_form_values(records, {"גק-1": {"בדיקת IPA": "V", "תאריך טופס": "19/10/26", "PID [ppm]": "1.8"},
                                "גק-4": {"בדיקת IPA": "X", "תאריך טופס": "19/10/26", "PID [ppm]": ""}})
    out = str(tmp_path / "gas.xlsx")
    LabReportExcel(records=records, threshold_manager=tm, output_path=out, project_name="", client="",
                   report_date="", selected_thresholds=["GAS_INDOOR_RES"], combine_tph_voc=False,
                   combine_tph_mbtex=False).build()
    ws = openpyxl.load_workbook(out).worksheets[0]
    rows = {str(r[0].value): [c.value for c in r[1:] if c.value is not None] for r in ws.iter_rows(max_row=7)}
    assert rows["בדיקת IPA"] == ["V", "X"]
    assert rows["PID [ppm]"] == ["1.8"]
    assert rows["תאריך טופס"] == ["19/10/26", "19/10/26"]   # only the sampling date row stays manual
    assert rows["תאריך ביצוע הדיגום"] == []


def test_soil_gas_pdf_sample_names_are_not_reversed():
    import pdfplumber

    from parsers.soil_gas.alchem_pdf import AlchemSoilGasPDFParser
    raw = _pdf("page")
    # A test PDF's font can't hold Hebrew, so the table the parser reads is faked
    table = [["Canister Number:", "", "8610", "", ""], ["Analysis Time:", "", "10:52", "", ""],
             ["Analysis Location:", "", "1-קג", "", ""], ["Benzene", "71-43-2", "120", "10", "50"]]
    with pdfplumber.open(io.BytesIO(raw)) as real:
        page_class = type(real.pages[0])
    original = page_class.extract_tables
    page_class.extract_tables = lambda self, *a, **k: [table]
    try:
        records = AlchemSoilGasPDFParser().parse(io.BytesIO(raw))
    finally:
        page_class.extract_tables = original
    assert records[0]["sample_id"] == "גק-1"
    assert records[0]["canister_num"] == "8610"
