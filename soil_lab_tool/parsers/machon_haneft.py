"""
machon_haneft.py
----------------
Parser for מכון הנפט (Machon HaNeft) laboratory Excel reports.

Machon HaNeft reports are in Hebrew and may have merged header cells.
The parser attempts to locate the data table by scanning for a row
that contains 'CAS' or 'תרכובת'.
"""

from __future__ import annotations

import io

import pandas as pd

from parsers.base import BaseParser
from core.lab_value_parser import LabValueParser
from core.cas_lookup import name_to_cas


class MachonHaneftParser(BaseParser):
    LAB_NAME = "מכון הנפט"

    def __init__(self):
        self._vp = LabValueParser()

    # ------------------------------------------------------------------
    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        df = self._read_excel(file_obj)
        records = []

        for _, row in df.iterrows():
            compound  = self._find_col(row, ["תרכובת", "compound", "שם כימי"])
            cas       = self._find_col(row, ["cas", "מספר cas", "cas no"])
            if not cas:
                cas = name_to_cas(compound) or ""
            raw_val   = self._find_col(row, ["ריכוז", "result", "תוצאה", "value"])
            unit      = self._find_col(row, ["יחידה", "unit", "units"]) or "µg/m³"
            sample_id = self._find_col(row, ["מזהה", "sample", "sample id", "id"])

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

        # Read without header first to find the actual header row
        raw = xl.parse(sheet, header=None, dtype=str).fillna("")

        header_row = 0
        for i, row in raw.iterrows():
            row_str = " ".join(str(v) for v in row.values).lower()
            if "cas" in row_str or "תרכובת" in row_str:
                header_row = i
                break

        df = xl.parse(sheet, header=header_row, dtype=str).fillna("")
        df.columns = [str(c).strip() for c in df.columns]
        return df

    @staticmethod
    def _find_col(row, aliases: list[str]) -> str:
        for alias in aliases:
            for col in row.index:
                if str(col).strip().lower() == alias.lower():
                    val = str(row[col]).strip()
                    if val and val.lower() not in ("nan", "none", ""):
                        return val
        return ""
