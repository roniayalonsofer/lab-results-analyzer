"""
parsers/pfas/kte.py
--------------------
Parser for KTE PFAS soil analysis reports in long-format (tidy) layout.

Same 24-column LIMS export. Analysis code: PFAS_SOIL
Units: ng/kg
"""

from __future__ import annotations

import io
import re

import pandas as pd

from parsers.base import BaseParser
from core.lab_value_parser import LabValueParser


PFAS_ANALYSIS_CODES = {"PFAS_SOIL", "PFAS_WATER", "PFAS_GW", "PFAS"}


class KTEPFASParser(BaseParser):
    LAB_NAME = "KTE"
    ANALYSIS_TYPES = ["SOIL_PFAS", "GW_PFAS"]

    C_SAMPLE = 1
    C_ACODE  = 2
    C_CPND   = 4
    C_RESULT = 5
    C_UNIT   = 6
    C_DATE   = 8
    C_LOC    = 13

    def __init__(self):
        self._vp = LabValueParser()

    def parse(self, file_obj: io.BytesIO | str) -> list[dict]:
        df = self._read(file_obj)
        if df is None or df.empty:
            return []

        records: list[dict] = []
        for _, row in df.iterrows():
            acode = str(row.iloc[self.C_ACODE]).strip().upper()
            atype = self._resolve_atype(acode)
            if atype is None:
                continue

            compound = str(row.iloc[self.C_CPND]).strip()
            raw_val  = str(row.iloc[self.C_RESULT]).strip()
            unit     = str(row.iloc[self.C_UNIT]).strip()
            loc      = str(row.iloc[self.C_LOC]).strip()
            date_val = str(row.iloc[self.C_DATE]).strip()

            if not compound or compound.lower() in ("nan", ""):
                continue

            if raw_val.lower() in ("not detected", "nd", "n.d.", "n/d", "<dl", ""):
                value, flag = None, "ND"
            else:
                value, flag = self._vp.parse(raw_val)

            sample_id = loc if loc and loc.lower() not in ("nan", "") else f"Sample-{row.iloc[self.C_SAMPLE]}"
            date_str = self._short_date(date_val)
            if date_str:
                sample_id = f"{sample_id} ({date_str})"

            records.append({
                "lab":           self.LAB_NAME,
                "sample_id":     sample_id,
                "compound":      compound,
                "cas":           "",   # CAS for PFAS often in compound name
                "value":         value,
                "flag":          flag,
                "unit":          unit or ("ng/kg" if atype == "SOIL_PFAS" else "ng/L"),
                "lod":           None,
                "analysis_type": atype,
            })

        return records

    def _read(self, file_obj: io.BytesIO | str) -> pd.DataFrame | None:
        try:
            if isinstance(file_obj, str) and file_obj.lower().endswith(".csv"):
                df = pd.read_csv(
                    file_obj, encoding="utf-8-sig", dtype=str,
                    header=None, engine="python",
                    on_bad_lines="skip", quotechar='"',
                ).fillna("")
                if df.shape[0] > 2:
                    first = str(df.iloc[0, 0]).strip()
                    if not first.replace("-", "").isdigit():
                        df = df.iloc[2:].reset_index(drop=True)
            else:
                xl = pd.ExcelFile(file_obj)
                df = xl.parse(xl.sheet_names[0], header=None, dtype=str).fillna("")
                first_cell = str(df.iloc[0, 0]).strip()
                if not first_cell.replace("-", "").isdigit():
                    df = df.iloc[2:].reset_index(drop=True)
        except Exception as e:
            raise ValueError(f"KTEPFASParser: cannot read file — {e}") from e
        return df

    @staticmethod
    def _resolve_atype(acode: str) -> str | None:
        if "PFAS_SOIL" in acode or "PFAS_W" not in acode and "PFAS" in acode:
            return "SOIL_PFAS"
        if "PFAS_WATER" in acode or "PFAS_GW" in acode:
            return "GW_PFAS"
        return None

    @staticmethod
    def _short_date(date_str: str) -> str:
        m = re.search(r'(\d{4})-(\d{2})-(\d{2})', date_str)
        if m:
            return f"{m.group(3)}.{m.group(2)}.{m.group(1)}"
        return ""
