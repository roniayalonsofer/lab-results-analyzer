"""Regression tests for field forms (PID) and for choosing a PDF's parser.

Uses small synthetic files, so no client data is committed.
Run:  py -3 -m pytest tests -q
"""
import io
import os
import sys

import openpyxl
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "soil_lab_tool"))

from core.excel_output import _pid_key  # noqa: E402
from core.field_forms import _find_columns, _read_rows, merge_field_data, parse_field_form  # noqa: E402
from parsers import category_layout  # noqa: E402
from parsers.soil.alchem import AlchemMultiSectionParser  # noqa: E402
from parsers.soil.bactochem import fix_rtl_name  # noqa: E402

HEADER = ["קידוח", "מעומק (מ')", "עד עומק (מ')", "סוג קרקע", "תיאור", "צבע", "לחות", "PID (ppm)"]


def _xlsx(*rows, header=HEADER) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    if header:
        ws.append(header)
    for r in rows:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _row(bh, depth_to, pid):
    return [bh, 0, depth_to, "", "", "", "", pid]


def test_excel_form_with_borehole_only_on_its_first_row():
    data = parse_field_form("form.xlsx", _xlsx(
        _row("ק-1", 1.5, 0.7), _row(None, 3, None), _row(None, 4, "N/A"), _row(None, 5, 12),
        _row("Ķ-31", 2, 0),
    ))
    assert data[_pid_key("ק-1")] == [(1.5, 0.7), (5, 12)]
    assert data[_pid_key("K-31")] == [(2, 0)]


def test_excel_form_columns_are_found_by_header():
    header = ["PID", "הערות", "עד עומק", "קידוח"]
    data = parse_field_form("form.xlsx", _xlsx([3.3, "", 2, "P-1"], header=header))
    assert data[_pid_key("P-1")] == [(2, 3.3)]


def test_excel_form_without_header_keeps_the_known_columns():
    data = parse_field_form("form.xlsx", _xlsx(_row("P-2", 3, 4.5), header=None))
    assert data[_pid_key("P-2")] == [(3, 4.5)]


def test_file_with_no_readings_is_reported():
    with pytest.raises(ValueError):
        parse_field_form("form.xlsx", _xlsx(["a", "b"], header=["x", "y"]))


def test_several_forms_merge_and_the_later_one_wins_at_the_same_depth():
    merged: dict = {}
    merge_field_data(merged, parse_field_form("a.xlsx", _xlsx(_row("P-1", 1, 1.0), _row("P-2", 1, 2.0))))
    merge_field_data(merged, parse_field_form("b.xlsx", _xlsx(_row("P-1", 1, 9.0), _row("P-1", 2, 3.0))))
    assert merged[_pid_key("P-1")] == [(1, 9.0), (2, 3.0)]
    assert merged[_pid_key("P-2")] == [(1, 2.0)]


def test_pdf_table_with_reversed_columns_and_hebrew():
    # A PDF table can arrive right to left, with Hebrew text in visual order
    rows = [["PID\n(ppm)", "קמוע דע", "חודיק"], ["0.7", "1.5", "12-ק"], ["12", "3", ""]]
    start, columns = _find_columns(rows)
    data = _read_rows(rows[start + 1:], columns, fix_text=fix_rtl_name)
    assert data[_pid_key("ק-12")] == [(1.5, 0.7), (3, 12)]


def _pdf(texts=(), grid=None) -> bytes:
    """A minimal one-page PDF: *texts* as lines of text, *grid* as a ruled table."""
    ops = [f"BT /F1 10 Tf 100 {700 - 20 * i} Td ({t}) Tj ET" for i, t in enumerate(texts)]
    if grid:
        x0, y0, w, h = 100, 700, 80, 20
        n_rows, n_cols = len(grid), len(grid[0])
        for r in range(n_rows + 1):
            ops.append(f"{x0} {y0 - r * h} m {x0 + n_cols * w} {y0 - r * h} l S")
        for c in range(n_cols + 1):
            ops.append(f"{x0 + c * w} {y0} m {x0 + c * w} {y0 - n_rows * h} l S")
        for r, row in enumerate(grid):
            for c, cell in enumerate(row):
                ops.append(f"BT /F1 10 Tf {x0 + c * w + 4} {y0 - (r + 1) * h + 6} Td ({cell}) Tj ET")
    stream = "\n".join(ops).encode()
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R "
        b"/Resources << /Font << /F1 5 0 R >> >> >>",
        b"<< /Length %d >>\nstream\n%s\nendstream" % (len(stream), stream),
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]
    out, offsets = b"%PDF-1.4\n", []
    for i, obj in enumerate(objs, 1):
        offsets.append(len(out))
        out += b"%d 0 obj\n%s\nendobj\n" % (i, obj)
    xref = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1)
    out += b"".join(b"%010d 00000 n \n" % o for o in offsets)
    out += b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objs) + 1, xref)
    return out


def test_pdf_form_end_to_end():
    raw = _pdf(grid=[["Borehole", "Depth to", "PID"], ["P-1", "1.5", "0.7"], ["", "3", "12"], ["P-2", "2", "4"]])
    data = parse_field_form("form.PDF", raw)
    assert data[_pid_key("P-1")] == [(1.5, 0.7), (3, 12)]
    assert data[_pid_key("P-2")] == [(2, 4)]


def test_pdf_without_a_pid_table_is_reported():
    with pytest.raises(ValueError):
        parse_field_form("scan.pdf", _pdf(["no table here"]))


def _alchem_soil_gas_pdf() -> bytes:
    return _pdf(["Canister Number: 9330", "EPA TO-15"])


@pytest.mark.parametrize("lab", ["Alchem", "alchem"])
def test_alchem_soil_gas_pdf_gets_the_pdf_parser_with_the_lab_as_displayed(lab):
    assert category_layout("soil_gas", "table.PDF", _alchem_soil_gas_pdf(), lab) == "soil_gas_pdf"


def test_layout_is_used_only_when_the_lab_has_a_parser_for_it():
    assert category_layout("soil_gas", "table.PDF", _alchem_soil_gas_pdf(), "KTE") == "soil_gas"


@pytest.mark.parametrize("visual,logical", [
    ("1_12-קג", "גק-12_1"),
    ("(0.5) 12-ק", "ק-12 (0.5)"),
    ("SG1_12", "SG1_12"),
])
def test_reversed_sample_names_keep_their_digits(visual, logical):
    assert fix_rtl_name(visual) == logical


def _alchem_voc_header():
    return [["Sample Name:", "", "1_12-קג", "SG1_12", "", ""],
            ["Compound Name", "CAS", "Final Conc.", "Final Conc.", "LOD", "LOQ"]]


def test_alchem_voc_continuation_page_without_a_header_is_read():
    parser = AlchemMultiSectionParser()
    first = parser._parse_voc_table(
        _alchem_voc_header() + [["Benzene", "71-43-2", "<MDL", "1.5", "0.01", "0.02"]])
    assert {r["sample_id"] for r in first} == {"גק-12_1", "SG1_12"}

    # The form's details table, on a later page, is not a continuation
    assert parser._parse_voc_table([["תאריך", "20/10/2026", "x", "y", "z", "w"]]) == []

    rest = parser._parse_voc_table([["Toluene", "108-88-3", "2.5", "<MDL", "0.01", "0.02"]])
    values = {(r["sample_id"], r["compound"]): (r["value"], r["flag"]) for r in rest}
    assert values == {("גק-12_1", "Toluene"): (2.5, ""), ("SG1_12", "Toluene"): (0.01, "<LOD")}


def test_alchem_metals_continuation_page_without_a_header_is_read():
    parser = AlchemMultiSectionParser()
    parser._parse_metals_table([["Sample Name:", "1_12-קג", "", ""],
                                ["Name", "Final Conc.", "LOD", "LOQ"],
                                ["Pb - Lead", "5.5", "0.1", "0.2"]])
    rest = parser._parse_metals_table([["Zn - Zinc", "7", "0.1", "0.2"]])
    assert [(r["sample_id"], r["compound"], r["value"]) for r in rest] == [("גק-12_1", "Zn - Zinc", 7.0)]


def test_alchem_tph_table_keeps_total_limits_and_not_detected():
    table = [["Sample Name", "DRO", "ORO", "TPH"],
             ["1_12-קג", "<LOQ", "N.D.", "<LOQ"],
             ["LOD", "7.00", "7.00", "15.00"],
             ["LOQ", "30.00", "20.00", "50.00"]]
    records = AlchemMultiSectionParser()._parse_tph_table(table)
    found = {r["compound"]: (r["sample_id"], r["value"], r["flag"], r["lod"], r["loq"]) for r in records}
    assert found["ORO"] == ("גק-12_1", 7.0, "<LOD", 7.0, 20.0)
    assert found["Total TPH"] == ("גק-12_1", 50.0, "<LOQ", 15.0, 50.0)


def test_threshold_cache_is_keyed_on_the_code_version():
    # Streamlit leaves a parameter starting with "_" out of the cache key, so an
    # update would keep serving a cached object of the old class.
    source = open(os.path.join(os.path.dirname(__file__), "..", "app.py"), encoding="utf-8").read()
    assert "def load_threshold_manager(code_mtime):" in source


@pytest.mark.parametrize("own_pid,form,expected", [
    (393.0, None, 393),                                  # typed in the app
    (0.0, None, 0),                                      # zero is a reading
    (393.0, {_pid_key("SG-7"): [(0.0, 5.0)]}, 393),      # typed wins over a form
    ("", None, None),                                    # nothing: no PID row
])
def test_soil_gas_report_shows_the_sample_pid(tmp_path, own_pid, form, expected):
    from core.excel_output import LabReportExcel
    from core.threshold_manager import ThresholdManager
    th = os.path.join(os.path.dirname(__file__), "..", "soil_lab_tool", "thresholds")
    tm = ThresholdManager(os.path.join(th, "soil_vsl_tier1_v7_2024.xlsx"),
                          vsl_full_path=os.path.join(th, "soil_vsl_v7_full.xlsx"))
    rec = dict(compound="Benzene", cas="71-43-2", value=120.0, flag="", unit="µg/m³", sample_id="SG-7",
               analysis_type="SOIL_GAS_VOC", canister_num="11919", lod=10.0, loq=50.0, pid_reading=own_pid)
    out = str(tmp_path / "gas.xlsx")
    LabReportExcel(records=[rec], threshold_manager=tm, output_path=out, project_name="", client="",
                   report_date="", selected_thresholds=["GAS_INDOOR_RES"], combine_tph_voc=False,
                   combine_tph_mbtex=False, pid_map=form).build()
    ws = openpyxl.load_workbook(out, rich_text=True).worksheets[0]
    rows = {str(r[0].value): [c.value for c in r[1:] if c.value is not None] for r in ws.iter_rows(max_row=5)}
    pid_rows = [v for k, v in rows.items() if "PID" in k]
    if expected is None:
        assert pid_rows == []
    else:
        assert pid_rows == [[expected]]


def test_word_shows_a_result_below_the_limit_as_below_it():
    from core.word_output import _fmt_value
    assert _fmt_value(50.0, "<LOQ", 15.0) == "<50"
    assert _fmt_value(7.0, "<LOD", 7.0) == "<7"
    assert _fmt_value(50.6, "", 15.0) == "50.6"


def test_chosen_category_is_kept_when_the_file_looks_like_another():
    assert category_layout("groundwater", "table.PDF", _alchem_soil_gas_pdf(), "Alchem") == "groundwater"
