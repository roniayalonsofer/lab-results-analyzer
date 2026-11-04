"""Regression tests for reading scanned field forms. The API is faked: no form leaves the machine.

Values follow the layout of a real soil gas custody page.
Run:  py -3 -m pytest tests -q
"""
import inspect
import io
import json
import os
import sys
from types import SimpleNamespace

import openpyxl
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "soil_lab_tool"))
sys.path.insert(0, os.path.dirname(__file__))

from core.field_form_ai import (  # noqa: E402
    MODEL, FormReadError, apply_form_values, field_options, find_api_key, form_rows, match_sample,
    read_form, secrets_problem,
)

READING = {
    "form_type": "גז קרקע",
    "general": [{"field": "טמפרטורה", "value": "24", "certain": True, "page": 7}],
    "samples": [
        {"sample_name": "גק-1", "canister": "8610", "values": [
            {"field": "בדיקת IPA", "value": "V", "certain": True, "page": 7},
            {"field": "PID", "value": "1.8", "certain": True, "page": 7},
        ]},
        {"sample_name": "גק-4", "canister": "11922", "values": [
            {"field": "בדיקת IPA", "value": "", "certain": True, "page": 7},
            {"field": "PID", "value": "1.5", "certain": False, "page": 7},
        ]},
    ],
}


class _Stream:
    def __init__(self, message):
        self.message = message

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def get_final_message(self):
        return self.message


def _client(text="", stop_reason="end_turn"):
    calls = []

    def stream(**kwargs):
        calls.append(kwargs)
        return _Stream(SimpleNamespace(
            stop_reason=stop_reason, content=[SimpleNamespace(type="text", text=text)]))

    client = SimpleNamespace(beta=SimpleNamespace(messages=SimpleNamespace(stream=stream)))
    return client, calls


def test_reading_sends_the_pdf_and_parses_the_answer():
    client, calls = _client(json.dumps(READING))
    assert read_form(client, b"%PDF-fake") == READING
    block = calls[0]["messages"][0]["content"][0]
    assert block["type"] == "document" and block["source"]["media_type"] == "application/pdf"
    assert calls[0]["model"] == MODEL
    assert calls[0]["output_config"]["format"]["type"] == "json_schema"


@pytest.mark.parametrize("stop_reason,text", [("refusal", ""), ("max_tokens", "{"), ("end_turn", "not json")])
def test_a_failed_reading_is_reported_in_hebrew(stop_reason, text):
    client, _ = _client(text, stop_reason)
    with pytest.raises(FormReadError):
        read_form(client, b"%PDF-fake")


def test_the_call_uses_parameters_the_sdk_accepts():
    anthropic = pytest.importorskip("anthropic")
    client = anthropic.Anthropic(api_key="not-a-real-key")
    params = inspect.signature(client.beta.messages.stream).parameters
    for name in ("model", "max_tokens", "betas", "fallbacks", "output_config", "system", "messages"):
        assert name in params


@pytest.mark.parametrize("secrets", [
    {"ANTHROPIC_API_KEY": "abc"},
    {" anthropic_api_key ": " abc "},
    {"general": {"ANTHROPIC_API_KEY": "abc"}},
])
def test_the_key_is_found_where_it_may_have_been_put(secrets):
    assert find_api_key(secrets) == "abc"


def test_a_missing_key_message_names_what_exists_but_no_values():
    assert find_api_key({"OTHER": "secret-value"}) is None
    message = secrets_problem({"OTHER": "secret-value"})
    assert "OTHER" in message and "secret-value" not in message
    assert secrets_problem({}) != message


def test_soil_gas_pdf_sample_names_are_not_reversed():
    from parsers.soil_gas.alchem_pdf import AlchemSoilGasPDFParser
    from test_field_forms import _pdf
    raw = _pdf(grid=[["Canister Number:", "", "8610"], ["Analysis Location:", "", "SG-1"],
                     ["Benzene", "71-43-2", "120", "10", "50"]])
    # The test PDF font can't hold Hebrew, so the parser's table is faked
    import pdfplumber
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


def test_options_list_every_field_once():
    assert field_options(READING) == ["בדיקת IPA", "PID", "טמפרטורה"]


def test_rows_merge_form_wide_values_and_flag_unsure_readings():
    rows = form_rows(READING)
    assert rows[1]["values"] == {"טמפרטורה": "24", "בדיקת IPA": "", "PID": "1.5"}
    assert rows[1]["unsure"] == ["PID"]
    assert rows[0]["unsure"] == []


def test_samples_match_by_canister_before_name():
    records = [dict(sample_id="ST-A", canister_num="8610"), dict(sample_id="גק - 4", canister_num="")]
    assert match_sample({"sample": "גק-9", "canister": "8610"}, records) == "ST-A"
    assert match_sample({"sample": "גק-4", "canister": "1"}, records) == "גק - 4"
    assert match_sample({"sample": "גק-7", "canister": ""}, records) is None


def test_chosen_values_reach_the_soil_gas_sheet(tmp_path):
    from core.excel_output import LabReportExcel
    from core.threshold_manager import ThresholdManager
    th = os.path.join(os.path.dirname(__file__), "..", "soil_lab_tool", "thresholds")
    tm = ThresholdManager(os.path.join(th, "soil_vsl_tier1_v7_2024.xlsx"),
                          vsl_full_path=os.path.join(th, "soil_vsl_v7_full.xlsx"))
    records = [dict(compound=c, cas=cas, value=120.0, flag="", unit="µg/m³", sample_id=s,
                    analysis_type="SOIL_GAS_VOC", canister_num=n, lod=10.0, loq=50.0)
               for c, cas in (("Benzene", "71-43-2"), ("Toluene", "108-88-3"))
               for s, n in (("גק-1", "8610"), ("גק-4", "11922"))]
    apply_form_values(records, {"גק-1": {"בדיקת IPA": "V", "תאריך טופס": "19/10/26"},
                                "גק-4": {"בדיקת IPA": "X", "תאריך טופס": "19/10/26"}})
    out = str(tmp_path / "gas.xlsx")
    LabReportExcel(records=records, threshold_manager=tm, output_path=out, project_name="", client="",
                   report_date="", selected_thresholds=["GAS_INDOOR_RES"], combine_tph_voc=False,
                   combine_tph_mbtex=False).build()
    ws = openpyxl.load_workbook(out).worksheets[0]
    rows = {str(r[0].value): [c.value for c in r[1:] if c.value is not None] for r in ws.iter_rows(max_row=7)}
    assert rows["בדיקת IPA"] == ["V", "X"]
    assert rows["תאריך טופס"] == ["19/10/26", "19/10/26"]   # only the sampling date row stays manual
    assert rows["תאריך ביצוע הדיגום"] == []
