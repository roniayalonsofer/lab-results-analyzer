"""
parsers/pfas/rj_lee.py
-----------------------
Parser for RJ Lee Group PFAS EDD Excel files.

Wide-format layout:
  Row 0 (header): Sample Name | PFHxA | PFOA | PFHxS | PFOS | ...
  Data rows:      <Lab ID>    | <number or ND> | ...
"""

from __future__ import annotations

import io

import pandas as pd

from parsers.base import BaseParser

_ND_VALUES = frozenset({"nd", "not detected", "n.d.", "n/d", "<dl", "", "nan"})
_DEFAULT_LOQ = 0.2


class RJLeePFASParser(BaseParser):
    LAB_NAME = "RJ Lee"
    ANALYSIS_TYPES = ["SOIL_PFAS"]

    def parse(self, file_obj: io.BytesIO | str) -> list[dict]:
        try:
            xl = pd.ExcelFile(file_obj)
            df = xl.parse(xl.sheet_names[0], header=0, dtype=str).fillna("")
        except Exception as e:
            raise ValueError(f"RJLeePFASParser: cannot read file — {e}") from e

        if df.empty:
            return []

        sample_col = df.columns[0]
        compound_cols = [
            c for c in df.columns[1:]
            if str(c).strip() and str(c).strip().lower() != "nan"
        ]

        records: list[dict] = []
        for _, row in df.iterrows():
            sample_id = str(row[sample_col]).strip()
            if not sample_id or sample_id.lower() in ("nan", ""):
                continue

            for compound in compound_cols:
                raw = str(row[compound]).strip()
                if raw.lower() in _ND_VALUES:
                    value, flag, loq = None, "ND", _DEFAULT_LOQ
                else:
                    try:
                        value = float(raw)
                    except ValueError:
                        value = None
                    flag = ""
                    loq = None

                records.append({
                    "lab":           self.LAB_NAME,
                    "sample_id":     sample_id,
                    "compound":      str(compound).strip(),
                    "cas":           "",
                    "value":         value,
                    "flag":          flag,
                    "unit":          "ng/kg",
                    "lod":           None,
                    "loq":           loq,
                    "analysis_type": "SOIL_PFAS",
                })

        return records
