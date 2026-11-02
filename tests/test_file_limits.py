"""Regression tests for the LOD / LOQ values shown from a lab file.

Run:  py -3 -m pytest tests -q
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "soil_lab_tool"))

from core.file_limits import compound_limits, field_values  # noqa: E402


def _rec(compound, atype, lod, loq, cas="1-1-1"):
    return dict(compound=compound, cas=cas, value=1.0, flag="", unit="mg/kg",
                sample_id="S1", analysis_type=atype, lod=lod, loq=loq)


def test_fields_fill_when_all_their_records_share_one_value():
    records = [
        _rec("DRO", "SOIL_TPH", 7.0, 30.0), _rec("ORO", "SOIL_TPH", 7.0, 20.0),
        _rec("Total TPH", "SOIL_TPH", 15.0, 50.0),
        _rec("Benzene", "SOIL_VOC", 0.01, 0.02), _rec("Toluene", "SOIL_VOC", 0.01, 0.02),
    ]
    assert field_values(records) == {
        "ov_lod_dro": 7.0, "ov_loq_dro": 30.0, "ov_lod_oro": 7.0, "ov_loq_oro": 20.0,
        "ov_lod_tph": 15.0, "ov_loq_tph": 50.0, "ov_lod_voc": 0.01, "ov_loq_voc": 0.02,
    }


def test_field_stays_empty_when_compounds_have_different_values():
    records = [_rec("Benzene", "SOIL_VOC", 0.01, 0.02), _rec("Toluene", "SOIL_VOC", 0.05, 0.02)]
    assert field_values(records) == {"ov_loq_voc": 0.02}


def test_soil_gas_has_no_shared_field_but_is_listed_per_compound():
    records = [_rec("Benzene", "SOIL_GAS_VOC", 16.2, 81.0, "71-43-2"),
               _rec("Benzene", "SOIL_GAS_VOC", 16.2, 81.0, "71-43-2"),
               _rec("Toluene", "SOIL_GAS_VOC", 19.0, 95.0, "108-88-3")]
    assert field_values(records) == {}
    rows = compound_limits(records)
    assert [(r["תרכובת"], r["LOD"], r["LOQ"]) for r in rows] == [("Benzene", 16.2, 81.0), ("Toluene", 19.0, 95.0)]
