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


def test_als_soil_elements_on_the_metals_sheet():
    """Potassium, sodium, calcium ... with the metals; ketones and alcohols with the VOCs."""
    rs = _records(ALS_SOIL, "als", "soil")
    t = {r["compound"]: r["analysis_type"] for r in rs}
    for el in ("Potassium", "Sodium", "Calcium", "Magnesium", "Aluminium", "Titanium"):
        assert t[el] == "SOIL_METALS", el
    for v in ("Acetone", "2-Butanone (MEK)", "tert-Butyl alcohol"):
        assert t[v] == "SOIL_VOC", v
    assert "Dry matter @ 105°C" not in t


# ── Known gaps (not read yet) ───────────────────────────────────────────────

def test_als_to15_soil_gas():
    """One sheet per sample; QC sheets skipped; "<0.36" a non-detect at the MDL; "J" kept."""
    rs = _records(ALS_TO15, "als", "soil_gas")
    assert len({r["sample_id"] for r in rs}) == 6 and len(rs) == 450
    v = {(r["sample_id"], r["compound"]): r for r in rs}
    assert (v[("EL 3.2", "Propene")]["value"], v[("EL 3.2", "Propene")]["flag"]) == (16.0, "")
    assert (v[("EL 3.2", "Chloromethane")]["value"], v[("EL 3.2", "Chloromethane")]["flag"]) == (0.36, "<LOD")
    assert v[("EL 3.2", "1,2-Dichloropropane")]["qualifier"] == "J"
    assert ("EL 3.2", "1,2-Dichloro-1,1,2,2-tetrafluoroethane (CFC 114)") in v   # name over two lines
    assert {r["unit"] for r in rs} == {"µg/m³"}


def test_kte_groundwater_edd():
    """An EDD sheet (ALS Houston, sent by KTE): one row per result, µg/L → mg/L."""
    rs = _records(KTE_EDD, "kte", "groundwater")
    v = {r["sample_id"]: r for r in rs}
    assert sorted(v) == ["4 (23.03.2023)", "5 (23.03.2023)", "6 (23.03.2023)"]
    assert (v["4 (23.03.2023)"]["value"], v["4 (23.03.2023)"]["cas"]) == (0.016, "123-91-1")
    assert v["6 (23.03.2023)"]["loq"] == 0.0001 and {r["unit"] for r in rs} == {"mg/L"}


# ── Machon HaNeft, files laid out as the printed certificate (2020-2023) ────

def test_machon_haneft_certificate_tph():
    rs = _records(HANEFT_OLD, "מכון הנפט", "soil")
    assert {(r["sample_id"], r["compound"], r["value"], r["flag"]) for r in rs} == {
        ("EL-2-1", "TPH (C10-C40)", 50.0, "<LOQ"), ("EL-2-3", "TPH (C10-C40)", 50.0, "<LOQ")}


def test_machon_haneft_certificate_voc():
    rs = _records(HANEFT_4231_VOC, "מכון הנפט", "soil")
    v = {r["compound"]: r for r in rs}
    assert len(rs) == 62 and {r["sample_id"] for r in rs} == {"EL-2-3"}
    assert (v["Methylene chloride"]["value"], v["Methylene chloride"]["flag"]) == (0.08, "")
    assert (v["Benzene"]["value"], v["Benzene"]["flag"], v["Benzene"]["loq"]) == (0.003, "<LOD", 0.01)


def test_machon_haneft_certificate_metals():
    rs = _records(HANEFT_2689, "מכון הנפט", "soil")
    v = {(r["sample_id"], r["compound"]): r for r in rs}
    assert len(rs) == 44
    assert v[("SZ-1.1-8 A-1", "Arsenic")]["value"] == 7.7 and v[("PS-1.1-14 A-3", "Iron")]["value"] == 1755
    assert (v[("PS-1.1-14 A-3", "Boron")]["value"], v[("PS-1.1-14 A-3", "Boron")]["flag"]) == (2.0, "<LOQ")


def test_machon_haneft_certificate_2023_job():
    """One job in three files: the same sample name in each; a value name over two rows."""
    voc = _records(HANEFT_2384 + "11.04.23/2384 SVOC.xlsx", "מכון הנפט", "soil")
    met = _records(HANEFT_2384 + "13.04.23/2384 מתכות.xlsx", "מכון הנפט", "soil")
    tph = _records(HANEFT_2384 + "16.04.23/2384 TPH+D+O.xlsx", "מכון הנפט", "soil")
    assert (len(voc), len(met), len(tph)) == (58, 22, 3)
    assert {r["sample_id"] for r in voc + met + tph} == {"TBM-3 R297-T A-1"}
    assert {r["analysis_type"] for r in voc} == {"SOIL_VOC"}      # the file name says SVOC
    assert "Isopropylbenzene (Cumene)" in {r["compound"] for r in voc}
    assert {r["compound"] for r in tph} == {"TPH (C10-C40)", "DRO", "ORO"}


# ── Aminolab soil ───────────────────────────────────────────────────────────

def test_aminolab_soil_pdf_and_excel_agree():
    """The certificate and its Excel appendix: the same 94 results, CAS from the lab's own list."""
    a = {(r["compound"].lower().replace(" ", "")): r
         for r in _records(AMINOLAB_XLSX, "aminolab", "soil")}
    b = {(r["compound"].lower().replace(" ", "")): r
         for r in _records(AMINOLAB_CERT, "aminolab", "soil")}
    assert len(a) == len(b) == 94
    for k, r in a.items():
        assert (r["value"], r["flag"], r["cas"], r["analysis_type"]) ==             (b[k]["value"], b[k]["flag"], b[k]["cas"], b[k]["analysis_type"]), k
    assert b["dro-oro"]["value"] == 1411 and b["calcium"]["value"] == 152065
    assert b["di-(2-ethylhexyl)phthalate"]["value"] == 3.12
    assert (b["benzo(a)pyrene"]["value"], b["benzo(a)pyrene"]["flag"]) == (0.03, "<LOQ")
    assert {r["sample_id"] for r in b.values()} == {"B1"}


AMINOLAB_PDFS = [
    ("AMINOLAB_1",
     ["קר-1 (0.1)", "קר-2 (0.2)"], 302),
    ("AMINOLAB_2",
     ["מכ-1 (0.3)", "קר-1 (0.4)", "קר-2 (0.4)", "קר-3 (0.3)", "קר-4 (0.2)"], 415),   # + a note "8.26" per sample, not a result
    ("AMINOLAB_3",
     ["מג-1 (0.1)", "ממ-1 (0.1)"], 298),
]


@pytest.mark.parametrize("path, samples, n", AMINOLAB_PDFS)
def test_aminolab_soil_pdf(path, samples, n):
    """Every value in the result columns is read (counted on the pages), each under its sample."""
    rs = _records(globals()[path], "aminolab", "soil")
    assert sorted({r["sample_id"] for r in rs}) == sorted(samples)
    assert len(rs) == n
    assert all(r["cas"] for r in rs if r["analysis_type"] in ("SOIL_VOC", "SOIL_SVOC", "SOIL_METALS"))


def test_aminolab_soil_pdf_values():
    rs = _records(AMINOLAB_1, "aminolab", "soil")
    v = {(r["sample_id"], r["compound"]): r for r in rs}
    s = "קר-1 (0.1)"
    assert v[(s, "DRO-ORO")]["value"] == 805 and v[(s, "GRO")]["value"] == 259
    assert v[(s, "Chromium")]["value"] == 3 and v[(s, "Copper")]["value"] == 2   # values printed off their line
    assert (v[(s, "Silver")]["value"], v[(s, "Silver")]["flag"]) == (5, "<LOQ")
    assert v[(s, "1,2,4-Trimethylbenzene")]["value"] == 0.33       # 330 ppb
    assert v[(s, "1,2-Dibromoethane (Ethylene dibromide)")]["cas"] == "106-93-4"


def test_aminolab_custody_form_is_not_a_certificate():
    """A scanned sample-collection form: nothing to read."""
    assert _records(AMINOLAB_PDF, "aminolab", "soil") == []
