"""EDD water sheets (KTE / ALS Houston) and Aminolab soil helpers, on made-up data."""
import io
import os
import sys

import pandas as pd

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "soil_lab_tool"))


def _xlsx(rows):
    buf = io.BytesIO()
    pd.DataFrame(rows[1:], columns=rows[0]).to_excel(buf, index=False)
    buf.seek(0)
    return buf


EDD_HEAD = ["ClientSampID", "CollectionDate", "TestName", "AnalyteType", "Analyte", "Result",
            "Qualifier", "FinalResult", "Runits", "Rlimit", "CAS", "TestNo", "SampID"]


def test_edd_detect_nondetect_and_surrogate():
    from parsers import get_parser
    f = _xlsx([
        EDD_HEAD,
        ["MW-1", "03/23/2023 12:00:00 AM", "Volatiles", "A", "Benzene", "2.5", "", "2.5", "ug/L", "1.0", "71-43-2", "SW8260", "X-01"],
        ["MW-1", "03/23/2023 12:00:00 AM", "Volatiles", "A", "Toluene", "ND", "U", "<1.0", "ug/L", "1.0", "108-88-3", "SW8260", "X-01"],
        ["MW-1", "03/23/2023 12:00:00 AM", "Volatiles", "S", "Toluene-d8", "98", "", "98", "%REC", "", "", "SW8260", "X-01"],
        ["MW-1", "03/23/2023 12:00:00 AM", "Semivolatiles SIM", "A", "1,4-Dioxane", "0.61", "J", "0.61", "ug/L", "1.0", "123-91-1", "SW8270SIM", "X-01"],
    ])
    rs = get_parser("kte", "groundwater").parse(f)
    v = {r["compound"]: r for r in rs}
    assert sorted(v) == ["1,4-Dioxane", "Benzene", "Toluene"]          # the surrogate is QC
    assert v["Benzene"]["sample_id"] == "MW-1 (23.03.2023)"           # month/day/year
    assert (v["Benzene"]["value"], v["Benzene"]["unit"], v["Benzene"]["analysis_type"]) == (0.0025, "mg/L", "GW_VOC")
    assert (v["Toluene"]["value"], v["Toluene"]["flag"]) == (0.001, "<LOQ")
    assert v["1,4-Dioxane"]["analysis_type"] == "GW_SVOC" and v["1,4-Dioxane"]["qualifier"] == "J"


def test_aminolab_sample_names():
    from parsers.soil.aminolab import _sample_name
    assert _sample_name("קרקע - קר -1 - (0.1) - טופס נטילה : 012601") == "קר-1 (0.1)"
    assert _sample_name("קרקע - מכ (0.3)-1 - טופס נטילה : 012752") == "מכ-1 (0.3)"
    assert _sample_name("קרקע B1 טופס נטילה 013319\n 12/03/2025, שעה: 10:30") == "B1"


def test_cas_check_digit():
    from parsers.soil.aminolab import _cas_ok
    assert _cas_ok("96-18-4") and _cas_ok("100-01-6") and _cas_ok("7440-22-4")
    assert not _cas_ok("95-18-4") and not _cas_ok("100-01-8")    # Aminolab misprints


def test_aminolab_excel_appendix():
    """Columns by header; section rows decide VOC (ppb) / SVOC (ppm); metals by symbol."""
    from parsers import get_parser
    rows = [["x"] * 4, ["הבדיקה", "יחידות מידה", "קרקע S-7 טופס נטילה 1\n01/01/2025", "הערות"],
            ["DRO+ORO", "ppm", "120", ""],
            ["סריקת מתכות מלאה ב- ICP", "", "", "2"],
            ["  עופרת - Pb  ", "mg/Kg", "<5", ""],
            ["VOC", "", "", ""],
            ["* Benzene", "", "<20", ""],
            ["SVOC - Semi-Volatiles", "", "", ""],
            ["4-Nitroaniline", "", "1.5", ""],
            ["הערות לבדיקה:", "", "", ""]]
    buf = io.BytesIO()
    pd.DataFrame(rows).to_excel(buf, index=False, header=False)
    buf.seek(0)
    v = {r["compound"]: r for r in get_parser("aminolab", "soil").parse(buf)}
    assert {r["sample_id"] for r in v.values()} == {"S-7"}
    assert (v["DRO-ORO"]["value"], v["DRO-ORO"]["analysis_type"]) == (120, "SOIL_TPH")
    assert (v["Lead"]["value"], v["Lead"]["flag"], v["Lead"]["cas"]) == (5, "<LOQ", "7439-92-1")
    assert (v["Benzene"]["value"], v["Benzene"]["analysis_type"]) == (0.02, "SOIL_VOC")
    assert (v["4-Nitroaniline"]["cas"], v["4-Nitroaniline"]["analysis_type"]) == ("100-01-6", "SOIL_SVOC")
