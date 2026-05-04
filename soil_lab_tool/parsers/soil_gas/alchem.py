"""
alchem.py
---------
Parser for Alchem / TO-15 soil gas laboratory Excel reports.

Expected layout (multi-sample format):
  Row 0: Canister Number:  | ... | ... | 8396 | 8573 | 8390
  Row 1: Analysis Time:    | ... | ... | ...
  Row 2: Analysis Location:| ... | ... | SG-2 | SG-6 | SG-5
  Row 3: Compound Name | CAS | LOD | LOQ | %UC | Final Conc. | Final Conc. | ...
  Row 4+: data rows

N.D. = Not Detected (below LOD)
"""

from __future__ import annotations

import io

import pandas as pd

from parsers.base import BaseParser
from core.lab_value_parser import LabValueParser


class AlchemSoilGasParser(BaseParser):
    LAB_NAME = "Alchem"
    ANALYSIS_TYPES = ["SOIL_GAS_VOC"]

    def __init__(self):
        self._vp = LabValueParser()

    # ------------------------------------------------------------------
    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        xl = pd.ExcelFile(file_obj)
        sheet = xl.sheet_names[0]
        raw = xl.parse(sheet, header=None, dtype=str).fillna("")

        # --- Find the header row (contains "Compound Name" or "CAS") ---
        header_row = self._find_header_row(raw)

        # --- Extract sample metadata from rows above the header ---
        sample_ids     = self._extract_meta_row(raw, header_row, "Analysis Location")
        canister_nums  = self._extract_meta_row(raw, header_row, "Canister Number")
        analysis_times = self._extract_meta_row(raw, header_row, "Analysis Time")
        pid_readings   = self._extract_meta_row(raw, header_row, "PID")

        # --- Parse header row ---
        headers = [str(v).strip() for v in raw.iloc[header_row].values]

        # Find column indices
        col_compound = self._find_col_idx(headers, ["compound name", "compound", "chemical", "analyte", "name"])
        col_cas      = self._find_col_idx(headers, ["cas", "cas no", "cas number"])
        col_lod      = self._find_col_idx(headers, ["lod", "lod [ug/m^3]", "mdl"])
        col_loq      = self._find_col_idx(headers, ["loq", "loq [ug/m^3]", "mql"])
        # All "Final Conc." columns (one per sample)
        conc_cols    = [i for i, h in enumerate(headers)
                        if "final conc" in h.lower() or "concentration" in h.lower()]

        if col_compound is None or col_cas is None:
            raise ValueError(
                f"❌ לא נמצאו עמודות Compound/CAS ב-row {header_row}. "
                f"כותרות שנמצאו: {headers}"
            )

        # --- Map conc columns → sample IDs ---
        # sample_ids list is aligned to the extra columns after the fixed cols
        def get_sample_id(col_idx: int, i: int) -> str:
            if sample_ids and i < len(sample_ids):
                return sample_ids[i]
            if canister_nums and i < len(canister_nums):
                return f"Canister-{canister_nums[i]}"
            return f"Sample-{i+1}"

        # --- Parse data rows ---
        records = []
        data_rows = raw.iloc[header_row + 1:].reset_index(drop=True)

        for _, row in data_rows.iterrows():
            values = list(row.values)
            compound = str(values[col_compound]).strip() if col_compound < len(values) else ""
            cas      = str(values[col_cas]).strip()      if col_cas      < len(values) else ""

            # Skip empty or summary rows
            if not compound or compound.lower() in ("", "nan", "compound name"):
                continue
            if "total voc" in compound.lower():
                continue  # skip summary rows

            # Handle dual-CAS compounds like "108-38-3 106-42-3"
            # Use first CAS for threshold lookup
            if " " in cas:
                cas = cas.split()[0]

            # LOD / LOQ values
            lod = None
            if col_lod is not None and col_lod < len(values):
                try:
                    lod = float(values[col_lod])
                except (ValueError, TypeError):
                    lod = None

            loq = None
            if col_loq is not None and col_loq < len(values):
                try:
                    loq = float(values[col_loq])
                except (ValueError, TypeError):
                    loq = None

            # One record per sample column
            for i, col_idx in enumerate(conc_cols):
                raw_val   = str(values[col_idx]).strip() if col_idx < len(values) else ""
                sample_id = get_sample_id(col_idx, i)

                # N.D. → plain ND (no < in input)
                if raw_val.upper() in ("N.D.", "ND", "N/D", "NOT DETECTED", ""):
                    value = lod
                    flag  = "ND"
                # <DL / <MDL / <LOD → explicit < prefix → display as <number
                elif raw_val.upper() in ("<DL", "<MDL", "<LOD", "<MRL", "<MDL"):
                    value = lod
                    flag  = "<LOD"
                # <LOQ → use LOQ as value with '<LOQ' flag
                elif raw_val.upper() == "<LOQ":
                    value = loq
                    flag  = "<LOQ"
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
                    "canister_num":  canister_nums[i]  if i < len(canister_nums)  else "",
                    "sampling_date": analysis_times[i] if i < len(analysis_times) else "",
                    "pid_reading":   pid_readings[i]   if i < len(pid_readings)   else "",
                })

        return records

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _find_header_row(self, df: pd.DataFrame) -> int:
        for i, row in df.iterrows():
            row_str = " ".join(str(v).lower() for v in row.values)
            if (("compound" in row_str) or ("name" in row_str)) and "cas" in row_str:
                return i
        return 3  # fallback

    def _extract_meta_row(self, df: pd.DataFrame, header_row: int, keyword: str) -> list[str]:
        """Find a row above header that contains `keyword`, return non-empty values after col 4."""
        for i in range(header_row):
            row = df.iloc[i]
            row_str = " ".join(str(v) for v in row.values[:3]).lower()
            if keyword.lower() in row_str:
                vals = [str(v).strip() for v in row.values[4:]]
                return [v for v in vals if v and v.lower() not in ("nan", "")]
        return []

    @staticmethod
    def _find_col_idx(headers: list[str], aliases: list[str]) -> int | None:
        for alias in aliases:
            for i, h in enumerate(headers):
                if alias.lower() in h.lower():
                    return i
        return None
