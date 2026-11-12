"""
Real lab certificates, one or more per lab and format. Where they are is
in tests/_real_files_local.py (not in git: client folder names); without
it, or without the drive, every test here is skipped. Where X: is not mounted (the
server, another computer) every test here is skipped.

Each case was checked by hand: against the PDF / Excel / CSV of the same
certificate (the two must agree), and spot values against the certificate
or the report that used it. A lab or format the app can't read yet is an
xfail with the reason, so the list of gaps stays visible.
"""
import io
import logging
import os
import sys
import warnings

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "soil_lab_tool"))
logging.disable(logging.WARNING)
warnings.filterwarnings("ignore")

try:   # where the files are: a local file, not in git (client folder names)
    from _real_files_local import *  # noqa: F401,F403
    _HERE = os.path.isdir(X)  # noqa: F405
except ImportError:
    _HERE = False
pytestmark = pytest.mark.skipif(not _HERE, reason="real lab files not reachable here")


def _read(path):
    return open(path, "rb").read()


def _records(path, lab, cat):
    from parsers import get_parser
    return get_parser(lab, cat).parse(io.BytesIO(_read(path)))


# ── Bactochem groundwater: PDF and CSV of each certificate agree ────────────

BACTO_PAIRS = [
    ("2024/מרץ/D230124-0048-01110896-C.pdf", "2024/מרץ/2732024165022.csv", "05.03.24"),
    ("2024/ספטמבר/D130524-0013-01159095-C.pdf", "2024/ספטמבר/D130524-0013.csv", "03.09.24"),
    ("2025/מרץ/D260125-0043-01205441-C.pdf", "2025/מרץ/303202510244.csv", "09.03.25"),       # CSV without CAS
    ("2025/ספטמבר/D310725-0097-01252010-HE-F.pdf", "2025/ספטמבר/3182025144119.csv", "17.08.25"),  # scrambled Hebrew font
    ("2026/D230726-0017-01378240-HE-R.pdf", "2026/D230726-0017-01378240-HE-R.csv", "16.09.26"),  # 2026 layout
]


@pytest.mark.parametrize("pdf, csv, date", BACTO_PAIRS)
def test_bactochem_pdf_and_csv_agree(pdf, csv, date):
    from soil_lab_tool.gw_report_updater import parse_bactochem_csv, parse_bactochem_pdf
    a = parse_bactochem_pdf(f"{MAALIA}/{pdf}")
    b = parse_bactochem_csv(_read(f"{MAALIA}/{csv}"))
    assert a["sampling_date"] == b["sampling_date"] == date
    assert sorted(a["samples"]) == sorted(b["samples"]) == ["P-1", "P-2", "T-1", "T-2"]
    for w in b["samples"]:
        assert a["samples"][w]["results"] == b["samples"][w]["results"], w


def test_bactochem_values_as_in_the_annual_report():
    """P-2 is the well before the equipment blank: the blank's results must not replace its own."""
    from soil_lab_tool.gw_report_updater import parse_bactochem_pdf
    m25 = parse_bactochem_pdf(f"{MAALIA}/2025/מרץ/D260125-0043-01205441-C.pdf")["samples"]
    a25 = parse_bactochem_pdf(f"{MAALIA}/2025/ספטמבר/D310725-0097-01252010-HE-F.pdf")["samples"]
    s24 = parse_bactochem_pdf(f"{MAALIA}/2024/ספטמבר/D130524-0013-01159095-C.pdf")["samples"]
    assert m25["P-2"]["results"]["MTBE"] == 0.029
    assert a25["P-2"]["results"]["MTBE"] == 0.024
    assert s24["T-2"]["results"]["MTBE"] == 1.29
    assert m25["T-2"]["results"]["MTBE"] == 3.0
    assert m25["P-1"]["field"]["pH"] == 7.37


def test_bactochem_2026_layout_sampler_and_field():
    from soil_lab_tool.gw_report_updater import parse_bactochem_pdf
    s = parse_bactochem_pdf(f"{MAALIA}/2026/D230726-0017-01378240-HE-R.pdf")["samples"]
    assert s["P-1"]["sampler_name"] == "עבדאללה חמדאן"
    assert s["P-1"]["field"]["pH"] == 7.37 and s["P-1"]["water_level"] == 1.10
    assert s["P-2"]["water_level"] == 2.10 and s["P-2"]["total_depth"] == 2.82   # bailer


# ── Alchem ──────────────────────────────────────────────────────────────────

def test_alchem_soil_gas_pdf_and_excel_agree():
    base = ALCHEM_GAS
    a = {(r["sample_id"], r["compound"]): r for r in _records(base + "table742117B-592.xlsx", "alchem", "soil_gas")}
    b = {(r["sample_id"], r["compound"]): r for r in _records(base + "table742117B-320.PDF", "alchem", "soil_gas_pdf")}
    assert len(a) == len(b) == 146
    for k, r in a.items():
        assert r["flag"] == b[k]["flag"], k
        if r["value"] is not None:   # the Excel keeps more digits than the PDF prints
            assert abs(r["value"] - b[k]["value"]) <= 0.011 * max(abs(b[k]["value"]), 1), k


def test_alchem_soil():
    rs = _records(ALCHEM_SOIL, "alchem", "soil")
    assert len({r["sample_id"] for r in rs}) == 78
    v = {(r["sample_id"], r["compound"]): r["value"] for r in rs}
    assert v[("K1-4.0m", "1,1-Biphenyl")] == 0.54
    assert {r["analysis_type"] for r in rs} == {"SOIL_VOC", "SOIL_SVOC", "SOIL_TPH", "SOIL_METALS"}


def test_alchem_to15_soil_gas():
    rs = _records(ALCHEM_TO15, "alchem", "soil_gas")
    assert len({r["sample_id"] for r in rs}) == 10 and len(rs) == 730


def test_xrf():
    rs = _records(XRF, "אלכם (xrf)", "soil")
    assert len({r["sample_id"] for r in rs}) == 37


# ── ALS ─────────────────────────────────────────────────────────────────────

def test_als_soil():
    rs = _records(ALS_SOIL, "als", "soil")
    assert len({r["sample_id"] for r in rs}) == 5
    v = {(r["sample_id"], r["compound"]): r["value"] for r in rs}
    assert v[("TM3.1 (4.0)", "Aluminium")] == 1750.0


@pytest.mark.xfail(reason="known: ALS soil files potassium (and maybe other metals) under SOIL_SVOC")
def test_als_soil_metals_on_the_metals_sheet():
    rs = _records(ALS_SOIL, "als", "soil")
    assert all(r["analysis_type"] == "SOIL_METALS" for r in rs if r["compound"] == "Potassium")


# ── Known gaps (not read yet) ───────────────────────────────────────────────

@pytest.mark.xfail(reason="not supported yet: ALS soil gas TO-15")
def test_als_to15_soil_gas():
    rs = _records(ALS_TO15, "als", "soil_gas")
    assert rs


@pytest.mark.xfail(reason="not read yet: KTE groundwater EDD Excel")
def test_kte_groundwater_edd():
    rs = _records(KTE_EDD, "kte", "groundwater")
    assert rs


@pytest.mark.xfail(reason="not read yet: older Machon HaNeft Excel (4231 / 2689)")
def test_machon_haneft_older_excel():
    rs = _records(HANEFT_OLD, "מכון הנפט", "soil")
    assert rs


@pytest.mark.xfail(reason="not read yet in the analysis page: Aminolab PDF certificates")
def test_aminolab_pdf_in_analysis_page():
    rs = _records(AMINOLAB_PDF, "aminolab", "groundwater")
    assert rs
