"""
kte.py
------
Parser for KTE laboratory Excel reports.

KTE reports typically use a transposed layout where compounds are columns
and samples are rows, or use a specific Hebrew header row structure.
Adjust the _read_excel method to match your actual KTE report format.
"""

from __future__ import annotations

import io

import pandas as pd

from parsers.base import BaseParser
from core.lab_value_parser import LabValueParser
from core.cas_lookup import name_to_cas


class KTEParser(BaseParser):
    LAB_NAME = "KTE"

    def __init__(self):
        self._vp = LabValueParser()

    # ------------------------------------------------------------------
    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        df = self._read_excel(file_obj)
        records = []

        for _, row in df.iterrows():
            compound  = str(row.get("compound", row.get("תרכובת", ""))).strip()
            cas       = str(row.get("cas", row.get("CAS", ""))).strip()
            if not cas:
                cas = name_to_cas(compound) or ""
            raw_val   = str(row.get("result", row.get("ריכוז", ""))).strip()
            unit      = str(row.get("unit", row.get("יחידה", "µg/m³"))).strip()
            sample_id = str(row.get("sample_id", row.get("מזהה", ""))).strip()

            value, flag = self._vp.parse(raw_val)

            records.append({
                "lab":       self.LAB_NAME,
                "sample_id": sample_id,
                "compound":  compound,
                "cas":       cas,
                "value":     value,
                "flag":      flag,
                "unit":      unit,
            })

        return records

    # ------------------------------------------------------------------
    def _read_excel(self, file_obj: io.BytesIO) -> pd.DataFrame:
        xl = pd.ExcelFile(file_obj)
        sheet = xl.sheet_names[0]
        df = xl.parse(sheet, dtype=str).fillna("")
        # Normalise column names
        df.columns = [str(c).strip().lower() for c in df.columns]
        return df
