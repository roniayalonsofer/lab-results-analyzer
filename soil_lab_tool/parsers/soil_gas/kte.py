"""
parsers/soil_gas/kte.py
------------------------
Parser for KTE TO-15 soil gas laboratory Excel reports.

Expected sheet name pattern: "<job_number>-TO-15-<descriptor>" (e.g. "40411-TO-15-1ppbv")

The format mirrors the Alchem soil gas layout:
  Meta rows (above header): Sample Name / Analysis Location / Canister Number / etc.
  Header row: Compound Name | CAS Number | LOD [...] | LOQ [...] | | Final Conc. | ...
  Data rows:  compound | cas | lod_val | loq_val | ... | conc1 | conc2 | ...

N.D. / ND = Not Detected (below LOD).
"""

from __future__ import annotations

import io

import pandas as pd

from parsers.base import BaseParser
from core.lab_value_parser import LabValueParser


class KTESoilGasParser(BaseParser):
    LAB_NAME = "KTE"
    ANALYSIS_TYPES = ["SOIL_GAS_VOC"]

    def __init__(self):
        self._vp = LabValueParser()

    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        xl = pd.ExcelFile(file_obj)

        # Find the TO-15 sheet (any sheet whose name contains TO-15 or ppbv)
        sheet = next(
            (s for s in xl.sheet_names
             if "to-15" in s.lower() or "ppbv" in s.lower()),
            xl.sheet_names[0],
        )
        raw = xl.parse(sheet, header=None, dtype=str).fillna("")

        header_row = self._find_header_row(raw)

        # Extract sample metadata from rows above the header
        sample_ids    = self._extract_meta_row(raw, header_row,
                                               ["sample name", "sample id",
                                                "analysis location", "location"])
        canister_nums = self._extract_meta_row(raw, header_row,
                                               ["canister", "cylinder"])
        analysis_times = self._extract_meta_row(raw, header_row,
                                                ["analysis time", "date"])
        pid_readings  = self._extract_meta_row(raw, header_row, ["pid"])

        headers = [str(v).strip() for v in raw.iloc[header_row].values]

        col_compound = self._find_col(headers, ["compound name", "compound", "chemical", "analyte", "name"])
        col_cas      = self._find_col(headers, ["cas", "cas no", "cas number"])
        col_lod      = self._find_col(headers, ["lod", "mdl", "dl"])
        col_loq      = self._find_col(headers, ["loq", "mql", "rl", "lor"])
        conc_cols    = [i for i, h in enumerate(headers)
                        if "final conc" in h.lower() or "concentration" in h.lower()
                        or ("conc" in h.lower() and h.lower() not in ("", "nan"))]

        if col_compound is None:
            # Fallback: assume col 0 = compound, col 1 = CAS
            col_compound, col_cas = 0, 1

        # If no explicit "Final Conc." columns, treat all columns after the last
        # fixed column (lod/loq) as sample columns
        if not conc_cols:
            last_fixed = max(c for c in [col_compound, col_cas, col_lod, col_loq]
                             if c is not None)
            conc_cols = list(range(last_fixed + 1, len(headers)))

        def _sample_id(col_idx: int, i: int) -> str:
            if sample_ids and i < len(sample_ids):
                return sample_ids[i]
            if canister_nums and i < len(canister_nums):
                return f"Canister-{canister_nums[i]}"
            return f"Sample-{i + 1}"

        records = []
        for _, row in raw.iloc[header_row + 1:].iterrows():
            values = list(row.values)

            compound = str(values[col_compound]).strip() if col_compound < len(values) else ""
            if not compound or compound.lower() in ("", "nan", "compound name", "compound"):
                continue
            if "total voc" in compound.lower() or "total toc" in compound.lower():
                continue

            cas_raw = str(values[col_cas]).strip() if col_cas is not None and col_cas < len(values) else ""
            cas = cas_raw.split()[0] if " " in cas_raw else cas_raw

            lod = self._parse_float(values, col_lod)
            loq = self._parse_float(values, col_loq)

            for i, ci in enumerate(conc_cols):
                raw_val   = str(values[ci]).strip() if ci < len(values) else ""
                sample_id = _sample_id(ci, i)

                if raw_val.upper() in ("N.D.", "ND", "N/D", "NOT DETECTED", ""):
                    value, flag = lod, "ND"
                elif raw_val.upper() in ("<DL", "<MDL", "<LOD", "<MRL"):
                    value, flag = lod, "<LOD"
                elif raw_val.upper() == "<LOQ":
                    value, flag = loq, "<LOQ"
                else:
                    value, flag = self._vp.parse(raw_val)

                records.append({
                    "lab":           self.LAB_NAME,
                    "sample_id":     sample_id,
                    "compound":      compound,
                    "cas":           cas,
                    "value":         value,
                    "flag":          flag,
                    "unit":          "µg/m³",
                    "lod":           lod,
                    "loq":           loq,
                    "analysis_type": "SOIL_GAS_VOC",
                    "canister_num":  canister_nums[i]   if i < len(canister_nums)   else "",
                    "sampling_date": analysis_times[i]  if i < len(analysis_times)  else "",
                    "pid_reading":   pid_readings[i]    if i < len(pid_readings)    else "",
                })

        return records

    # ------------------------------------------------------------------
    def _find_header_row(self, df: pd.DataFrame) -> int:
        for i, row in df.iterrows():
            row_str = " ".join(str(v).lower() for v in row.values)
            if ("compound" in row_str or "analyte" in row_str) and "cas" in row_str:
                return i
        # Fallback: return the row that has the most non-empty cells in the first 10 rows
        best, best_count = 3, 0
        for i in range(min(10, len(df))):
            cnt = sum(1 for v in df.iloc[i].values if str(v).strip() not in ("", "nan"))
            if cnt > best_count:
                best_count, best = cnt, i
        return best

    def _extract_meta_row(self, df: pd.DataFrame, header_row: int,
                          keywords: list[str]) -> list[str]:
        """Return non-empty values after col 3 from the first row above the header
        whose first 3 cells contain any of the keywords."""
        for i in range(header_row):
            row = df.iloc[i]
            row_str = " ".join(str(v) for v in row.values[:4]).lower()
            if any(kw in row_str for kw in keywords):
                vals = [str(v).strip() for v in row.values[3:]]
                return [v for v in vals if v and v.lower() not in ("nan", "")]
        return []

    @staticmethod
    def _find_col(headers: list[str], aliases: list[str]) -> int | None:
        for alias in aliases:
            for i, h in enumerate(headers):
                if alias.lower() in h.lower():
                    return i
        return None

    @staticmethod
    def _parse_float(values: list, col: int | None) -> float | None:
        if col is None or col >= len(values):
            return None
        try:
            return float(str(values[col]).strip())
        except (ValueError, TypeError):
            return None
