"""ALS groundwater reports and the drinking water standard.

Uses small synthetic files that mimic the real layouts, so no client data is committed.
"""
import io
import os
import sys

import openpyxl
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "soil_lab_tool"))

from core.compound_choices import find_ambiguous_compounds  # noqa: E402
from core.drinking_water import DRINKING_WATER, drinking_water_limit  # noqa: E402
from core.excel_output import LabReportExcel  # noqa: E402
from core.threshold_manager import ThresholdManager  # noqa: E402
from parsers import auto_detect_category, auto_detect_lab, lab_mismatch  # noqa: E402
from parsers.groundwater.als_gw import (ALSGroundwaterParser, _best_results, _record,  # noqa: E402
                                        _with_cas)

XLSX = "PR00C0000_0_EXCEL_GENERIC.xlsx"


def _workbook() -> bytes:
    """The real export's layout: sample labels sit in the LOR column."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Client WATER 1"
    ws["A2"] = "CLIENT"
    ws["A4"], ws["B4"] = "Work Order:", "PR00C0000"
    ws["D8"], ws["E8"] = "Client Sample ID", "Well A-1"
    ws["D9"], ws["E9"] = "Laboratory Sample ID", "PR00C0000001"
    ws["D10"], ws["E10"] = "Client Sampling Date", "23/09/2026"
    ws["A11"] = "RESULTS OF ANALYSIS"
    for ci, h in enumerate(["Parameter", "Method", "Unit", "LOR"], 1):
        ws.cell(row=13, column=ci, value=h)
    rows = [
        ("BTEX",),
        ("Benzene", "W-VOCGMS01", "µg/L", "0.20", "0.31"),
        ("Halogenated Volatile Organic Compounds",),
        ("Hexachlorobutadiene", "W-VOCGMS01", "µg/L", "1.0", "<1.0"),
        ("1.2-Dibromoethane (EDB)", "W-VOCGMS01", "µg/L", "0.50", "<0.50"),
        ("Organochlorine Pesticides",),
        ("Alachlor", "W-OCPECD01", "µg/L", "0.010", "<0.010"),
        ("Alachlor", "W-PESLMS02", "µg/L", "0.050", "<0.050"),
        ("Hexachlorobutadiene", "W-OCPECD01", "µg/L", "0.010", "<0.010"),
        ("Nonmetallic Inorganic Parameters",),
        ("Chloride", "W-CL-IC", "mg/L", "1.00", "59.9"),
        ("Acid neutralizing capacity (alkalinity) pH 4.5", "W-ALK-PCT", "mmol/L", "0.150", "3.26"),
        ("Total Metals / Major Cations",),
        ("Aluminium", "W-METAXDG1", "mg/L", "0.010", "0.726"),
        ("Arsenic", "W-METMSDG1", "µg/L", "1.0", "3.1"),
        ("Perfluorinated Compounds",),
        ("Perfluorooctanoic acid (PFOA)", "W-PFCLMS03", "ng/L", "10.0", "<10.0"),
    ]
    for ri, row in enumerate(rows, 15):
        for ci, v in enumerate(row, 1):
            ws.cell(row=ri, column=ci, value=v)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


@pytest.fixture(scope="module")
def tm():
    th = os.path.join(os.path.dirname(__file__), "..", "soil_lab_tool", "thresholds")
    return ThresholdManager(os.path.join(th, "soil_vsl_tier1_v7_2024.xlsx"),
                            pfas_path=os.path.join(th, "pfas_thresholds.xlsx"),
                            vsl_full_path=os.path.join(th, "soil_vsl_v7_full.xlsx"))


@pytest.fixture()
def records():
    return ALSGroundwaterParser().parse(io.BytesIO(_workbook()))


def _one(records, name, atype=None):
    found = [r for r in records if r["compound"] == name
             and (atype is None or r["analysis_type"] == atype)]
    assert len(found) == 1, (name, found)
    return found[0]


def test_real_xlsx_export_is_read(records):
    assert {r["sample_id"] for r in records} == {"Well A-1"}
    assert {r["sampling_date"] for r in records} == {"23.09.26"}
    benzene = _one(records, "Benzene")
    assert (benzene["value"], benzene["unit"], benzene["flag"]) == (0.00031, "mg/L", "")
    assert benzene["analysis_type"] == "GW_VOC"


def test_each_result_goes_to_its_sheet(records):
    assert _one(records, "Chloride")["analysis_type"] == "GW_CHEM"  # "Nonmetallic" isn't metals
    assert _one(records, "Aluminium")["analysis_type"] == "GW_METALS"
    assert _one(records, "Perfluorooctanoic acid (PFOA)")["analysis_type"] == "GW_PFAS"
    assert _one(records, "Alachlor")["analysis_type"] == "GW_SVOC"


def test_a_rows_own_unit_stays_visible_in_the_mixed_chemistry_sheet(records):
    alk = _one(records, "Acid neutralizing capacity (alkalinity) pH 4.5 [mmol/L]")
    assert (alk["value"], alk["unit"]) == (3.26, "mmol/L")


def test_every_compound_of_the_sheet_keeps_a_result_at_the_lowest_limit(records):
    assert _one(records, "Alachlor")["loq"] == pytest.approx(0.00001)
    # The two methods of this one land in different sheets, so both stay
    assert _one(records, "Hexachlorobutadiene", "GW_VOC")["loq"] == pytest.approx(0.001)
    assert _one(records, "Hexachlorobutadiene", "GW_SVOC")["loq"] == pytest.approx(0.00001)


def test_a_detection_beats_a_lower_non_detect():
    detected = _record("1.4-Dioxane", "W-VOCGMS01", "µg/L", 50.0, "60", "A", "", "VOC")
    lower = _record("1,4-Dioxane", "W-8260VOC", "µg/L", 0.2, "<0.200", "A", "", "VOC")
    assert _best_results([lower, detected]) == [detected]


def test_pdf_sample_name_with_spaces():
    ids = ALSGroundwaterParser._split_sample_ids(
        "Site Name GW-1 ---- ----", ["PR00C0000001", "----", "----"])
    assert ids == ["Site Name GW-1", "----", "----"]


def test_pdf_several_samples_without_spaces_keep_their_columns():
    ids = ALSGroundwaterParser._split_sample_ids("PP-2 PP-1 (EB) ----", ["L1", "L2", "----"])
    assert ids == ["PP-2", "PP-1 (EB)", "----"]


def test_pdf_lost_micro_sign_is_still_micrograms():
    r = _record("Benzene", "W-VOCGMS01", "g/L", 0.2, "0.31", "A", "", "BTEX")
    assert (r["value"], r["unit"]) == (0.00031, "mg/L")


def test_detection_of_the_real_xlsx():
    data = _workbook()
    assert auto_detect_lab(XLSX, data) == "als"
    assert auto_detect_category(XLSX, data, lab="als") == "groundwater"
    assert lab_mismatch(XLSX, data, "Alchem") == "als"
    assert lab_mismatch(XLSX, data, "ALS") is None


@pytest.mark.parametrize("name,cas,limit", [
    ("Benzene", "", 0.005),                     # First Addendum, Table C
    ("Chloroform", "67-66-3", 0.08),
    ("1.2-Dichloroethane", "", 0.004),          # the labs' dot spelling
    ("Nitrates", "", 70.0),                     # Table A, as NO3
    ("Mercury", "", 0.001),
    ("Aluminium", "", 0.2),                     # Second Addendum, in mg/L
    ("Sum of xylenes (M1)", "", 0.5),           # a standard for the sum
    ("", "7440-38-2", 0.01),                    # by CAS alone
])
def test_limits_copied_from_the_regulations(name, cas, limit):
    assert drinking_water_limit(cas, name) == pytest.approx(limit)


@pytest.mark.parametrize("name", ["Nitrate as N", "Atrazine-desethyl", "1,4-Dioxane", "Perchlorate"])
def test_compounds_the_regulations_dont_list(name):
    assert drinking_water_limit("", name) is None


def test_regulation_comparison_asks_nothing(records, tm):
    assert find_ambiguous_compounds(records, tm, [DRINKING_WATER]) == []


def test_report_compares_and_colours_against_the_regulations(records, tm, tmp_path):
    out = str(tmp_path / "gw.xlsx")
    LabReportExcel(records=records, threshold_manager=tm, output_path=out, project_name="", client="",
                   report_date="", selected_thresholds=[DRINKING_WATER], combine_tph_voc=False,
                   combine_tph_mbtex=False).build()
    wb = openpyxl.load_workbook(out)

    metals = wb["מי תהום מתכות"]
    header = [c.value for c in metals[2]]
    assert "LOQ [mg/L]" in header and "תקנות מי שתייה 2013 [mg/L]" in header
    rows = {row[0].value: row for row in metals.iter_rows(min_row=3)}
    assert rows["Aluminium"][3].value == pytest.approx(0.2)
    assert rows["Aluminium"][4].fill.fgColor.rgb.endswith("FFC000")  # 0.726 > 0.2
    assert rows["Arsenic"][4].fill.fill_type is None                 # 0.0031 < 0.01

    # A limit too small to round away, and the LOQ above the limit stands out
    voc = wb["מי תהום VOC"]
    edb = next(row for row in voc.iter_rows(min_row=3) if row[0].value == "1.2-Dibromoethane (EDB)")
    assert edb[3].value == pytest.approx(0.00005)
    assert edb[4].font.bold
    assert edb[4].number_format == '"<"0.0###'


@pytest.mark.parametrize("name,cas", [
    ("4:2 Fluorotelomer sulfonic acid (4:2 FTS)", "757124-72-4"),  # was 6:2 FTS's number
    ("Perfluorododecane sulfonic acid (PFDoDS)", "79780-39-5"),    # was PFDS's number
    ("1.2.3.4-Tetrachlorobenzene", "634-66-2"),                    # was a trichlorobenzene's
    ("Titanium", "7440-32-6"),                                     # was titanium tetrachloride's
    ("Perfluoro(2-ethoxyethane)sulfon ic acid (PFEESA)", "113507-82-7"),  # as a wrapped PDF line
    ("Chloride", "16887-00-6"),
    ("Total Cyanide", "57-12-5"),
])
def test_cas_numbers_of_the_als_package(name, cas):
    r = _with_cas(_record(name, "W-X", "mg/L", 1.0, "<1.0", "A", "", "Other"))
    assert r["cas"] == cas


def test_wrapped_pdf_name_gets_the_certificates_spelling():
    r = _with_cas(_record("Perfluoro(2-ethoxyethane)sulfon ic acid (PFEESA)", "W-X", "ng/L", 10.0,
                          "<10.0", "A", "", "Perfluorinated Compounds"))
    assert r["compound"] == "Perfluoro(2-ethoxyethane)sulfonic acid (PFEESA)"


@pytest.mark.parametrize("name", ["Sum of xylenes", "Nitrate as N", "Phosphorus (as P2O5)",
                                  "Isoproturon-desmethyl", "Total Petroleum Hydrocarbons"])
def test_no_cas_is_better_than_a_wrong_one(name):
    assert _with_cas(_record(name, "W-X", "mg/L", 1.0, "<1.0", "A", "", "Other"))["cas"] == ""


def test_lindane_under_the_labs_name():
    assert drinking_water_limit("", "Hexachlorocyclohexane Gamma") == pytest.approx(0.001)


def test_alpha_hch_isnt_lindane():
    assert drinking_water_limit("319-84-6", "Hexachlorocyclohexane Alpha") is None
