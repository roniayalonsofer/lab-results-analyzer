"""Regression tests for soil thresholds shown in the Excel output.

Hebrew/acronym compound names used to be flagged 'uncertain' against the English
threshold table, which printed "לא קיים" and skipped exceedance colouring.
Run:  py -3 -m pytest tests -q
"""
import os
import sys

import openpyxl
import pytest

ROOT = os.path.join(os.path.dirname(__file__), "..")
TOOL = os.path.join(ROOT, "soil_lab_tool")
sys.path.insert(0, TOOL)
sys.path.insert(1, ROOT)

from core.compound_choices import (  # noqa: E402
    NO_THRESHOLD, ResolvedThresholds, apply_compound_choices, find_ambiguous_compounds,
)
from core.excel_output import LabReportExcel  # noqa: E402
from core.threshold_manager import ThresholdManager  # noqa: E402

TH = os.path.join(TOOL, "thresholds")
PINK, YELLOW = "00FFB6C1", "00FFFF00"

# name, CAS, VSL, Tier 1 industrial low sensitivity (threshold table v7, 12/2024)
MBTEX = [
    ("MTBE", "1634-04-4", 0.7988618324158887, 110.4279459383089),
    ("בנזן", "71-43-2", 0.45303504887257845, 1.5126427434913277),
    ("טולואן", "108-88-3", 90.94567574764879, 168.29941402780491),
    ("אתיל-בנזן", "100-41-4", 7.115635767407982, 11.219645542638824),
    ("קסילן", "1330-20-7", 49.70136989317072, 49.70136989317072),
]


@pytest.fixture(scope="module")
def tm():
    return ThresholdManager(
        os.path.join(TH, "soil_vsl_tier1_v7_2024.xlsx"),
        pfas_path=os.path.join(TH, "pfas_thresholds.xlsx"),
        vsl_full_path=os.path.join(TH, "soil_vsl_v7_full.xlsx"),
    )


@pytest.mark.parametrize("name,cas,vsl,tier1", MBTEX)
def test_mbtex_thresholds_are_exact(tm, name, cas, vsl, tier1):
    assert tm.get_threshold_with_confidence(cas, "VSL_SOIL", name) == (
        pytest.approx(vsl), "exact")
    assert tm.get_threshold_with_confidence(cas, "TIER1_IND_SOIL_LOW", name) == (
        pytest.approx(tier1), "exact")


def test_cas_decides_whatever_the_lab_calls_the_compound(tm):
    assert tm.get_threshold_with_confidence(
        "71-43-2", "VSL_SOIL", "Some other spelling")[1] == "exact"
    # A summed m+p result is not the single isomer the threshold is for.
    assert tm.get_threshold_with_confidence(
        "108-38-3", "VSL_SOIL", "Meta- & Para-Xylene")[1] == "uncertain"


def _rec(name, cas):
    return dict(compound=name, cas=cas, value=5.0, flag="", unit="mg/kg",
                sample_id="S1", analysis_type="SOIL_VOC")


def _asked(tm, name, cas):
    return [f["compound"] for f in find_ambiguous_compounds([_rec(name, cas)], tm, ["VSL_SOIL"])]


def test_compound_with_cas_in_table_is_not_asked_about(tm):
    assert _asked(tm, "Some other spelling", "71-43-2") == []


@pytest.mark.parametrize("name", ["MTBE", "בנזן", "Xylenes"])
def test_clear_name_without_cas_is_not_asked_about(tm, name):
    assert _asked(tm, name, "") == []


def test_unclear_name_without_cas_is_asked_about(tm):
    found = find_ambiguous_compounds([_rec("Xylene mix", "")], tm, ["VSL_SOIL"])
    assert [f["compound"] for f in found] == ["Xylene mix"]
    assert "1330-20-7" in [c for _, c in found[0]["candidates"]]


def test_cas_not_in_table_means_no_threshold_and_no_question(tm):
    assert _asked(tm, "Bismuth", "7440-69-9") == []
    assert _asked(tm, "1.3.5-Trichlorobenzene", "108-70-3") == []


def test_known_name_under_a_cas_the_table_lacks_is_asked_about(tm):
    assert _asked(tm, "Benzene", "0000-00-0") == ["Benzene"]


def test_compound_with_nothing_similar_is_not_asked_about(tm):
    assert _asked(tm, "Qqqqq", "") == []


def test_no_uncertainty_footer_is_written(tm, tmp_path):
    # Lookups that used to pile "חוסר וודאות" lines on the shared threshold manager
    find_ambiguous_compounds([_rec("Xylene mix", "")], tm, ["VSL_SOIL"])
    out = str(tmp_path / "report.xlsx")
    records = [dict(_rec("בנזן", "71-43-2"), analysis_type="SOIL_MBTEX")]
    LabReportExcel(
        records=records, threshold_manager=tm, output_path=out, project_name="", client="",
        report_date="", selected_thresholds=["VSL_SOIL"], combine_tph_voc=False,
        combine_tph_mbtex=False,
    ).build()
    ws = openpyxl.load_workbook(out)["קרקע MBTEX"]
    assert not [c.coordinate for row in ws.iter_rows() for c in row
                if "חוסר וודאות" in str(c.value or "")]


@pytest.mark.parametrize("name,cas", [
    ("MTBE", "1634-04-4"),
    ("MTBE*", "1634-04-4"),
    ("M.T.B.E", "1634-04-4"),
    ("MTBE (Methyl tert-butyl ether)", "1634-04-4"),
    ("Methyl-t-butyl ether", "1634-04-4"),
    ("Methyl tertiary butyl ether", "1634-04-4"),
    ("Total Xylenes", "1330-20-7"),
    ("נפתלן", "91-20-3"),
    ("1,2-DCA", "107-06-2"),
    ("1,2,4-TMB", "95-63-6"),
])
def test_clear_names_are_not_ambiguous(tm, name, cas):
    assert tm.get_threshold_with_confidence(cas, "VSL_SOIL", name)[1] == "exact"


@pytest.mark.parametrize("name", ["m+p-Xylene", "m,p-Xylene", "m/p-Xylene", "Meta- & Para-Xylene"])
def test_summed_xylene_needs_an_answer(tm, name):
    rec = dict(compound=name, cas="108-38-3", value=5.0, flag="", unit="mg/kg",
               sample_id="S1", analysis_type="SOIL_VOC")
    found = find_ambiguous_compounds([rec], tm, ["VSL_SOIL"])
    assert [f["compound"] for f in found] == [name]
    cands = [c for _, c in found[0]["candidates"]]
    assert cands[0] == "108-38-3"   # its own CAS row first
    assert "1330-20-7" in cands     # total xylenes is offered


def _build_mp_xylene_sheet(tm, tmp_path, choice):
    name = "Meta- & Para-Xylene"
    records = [dict(compound=name, cas="108-38-3", value=80.3, flag="", unit="mg/kg",
                    sample_id="S1", analysis_type="SOIL_MBTEX")]
    choices = {name: choice}
    out = str(tmp_path / "report.xlsx")
    LabReportExcel(
        records=apply_compound_choices(records, choices),
        threshold_manager=ResolvedThresholds(tm, choices),
        output_path=out, project_name="", client="", report_date="",
        selected_thresholds=["VSL_SOIL"], combine_tph_voc=False, combine_tph_mbtex=False,
    ).build()
    return openpyxl.load_workbook(out)["קרקע MBTEX"]


def test_chosen_compound_gets_its_threshold_with_no_note(tm, tmp_path):
    ws = _build_mp_xylene_sheet(tm, tmp_path, "1330-20-7")
    cells = [c for row in ws.iter_rows() for c in row]
    assert not [c.coordinate for c in cells if "תרכובת דומה" in str(c.value or "")]
    assert not [c.coordinate for c in cells if str(c.value or "").endswith(" *")]
    assert any(c.value == "1330-20-7" for c in cells)
    assert any(c.value == pytest.approx(49.7, abs=0.05) for c in cells
               if isinstance(c.value, float))  # total xylenes VSL, as displayed
    value = next(c for c in cells if c.value == 80.3)
    assert value.fill.start_color.rgb == YELLOW


def test_no_threshold_answer_leaves_the_compound_without_one(tm, tmp_path):
    ws = _build_mp_xylene_sheet(tm, tmp_path, NO_THRESHOLD)
    cells = [c for row in ws.iter_rows() for c in row]
    assert any(c.value == "לא קיים" for c in cells)
    assert not [c.coordinate for c in cells if "תרכובת דומה" in str(c.value or "")]
    assert not next(c for c in cells if c.value == 80.3).fill.fill_type


def test_mbtex_sheet_shows_thresholds_and_colours(tm, tmp_path):
    values = {"MTBE": 2.73, "בנזן": 1.82, "טולואן": 0.17, "אתיל-בנזן": 9.54, "קסילן": 80.3}
    records = [
        dict(compound=n, cas=c, value=values[n], flag="", unit="mg/kg",
             sample_id="ב-1-1.5", analysis_type="SOIL_MBTEX")
        for n, c, _, _ in MBTEX
    ]
    out = str(tmp_path / "report.xlsx")
    LabReportExcel(
        records=records, threshold_manager=tm, output_path=out, project_name="",
        client="", report_date="", selected_thresholds=["VSL_SOIL", "TIER1_IND_SOIL_LOW"],
        combine_tph_voc=False, combine_tph_mbtex=False,
    ).build()

    ws = openpyxl.load_workbook(out)["קרקע MBTEX"]
    assert not [c.coordinate for row in ws.iter_rows() for c in row if c.value == "לא קיים"]
    assert not [c.coordinate for row in ws.iter_rows() for c in row
                if "תרכובת דומה" in str(c.value or "")]

    cells = {c.value: c for row in ws.iter_rows() for c in row if c.value in values.values()}
    fill = lambda v: cells[v].fill.start_color.rgb if cells[v].fill.fill_type else None
    assert fill(2.73) == YELLOW   # MTBE above VSL only
    assert fill(1.82) == PINK     # benzene above Tier 1
    assert fill(0.17) is None     # toluene below both
    assert fill(9.54) == YELLOW   # ethylbenzene above VSL only
    assert fill(80.3) == PINK     # xylene above Tier 1
