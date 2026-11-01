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


def test_different_compound_with_same_cas_stays_uncertain(tm):
    # A lab name our map doesn't tie to this CAS must keep the safety flag.
    assert tm.get_threshold_with_confidence(
        "71-43-2", "VSL_SOIL", "Benzene derivative")[1] == "uncertain"
    # A summed m+p result is not the single isomer the threshold is for.
    assert tm.get_threshold_with_confidence(
        "108-38-3", "VSL_SOIL", "Meta- & Para-Xylene")[1] == "uncertain"


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
