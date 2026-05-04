"""
alchem_soil.py
--------------
Parser for Alchem soil analysis Excel reports.

Expected layout:
  Sheet "VOC":
    Row 0: Analysis Location | ... | ק-3 | ק-4 | ק-5 | ק-1 | ק-2 | ...
    Row 1: Compound Name | CAS number | LOD | LOQ | %U.C. | Final Concentration | ...
    Row 2+: data rows

  Sheet "TPH":
    Row 0: Sample Name | DRO [mg/kg] | ORO [mg/kg] | Total TPH [mg/kg]
    Row 1+: data rows

Units: mg/kg (not µg/m³)
Flags: <MDL = below Method Detection Limit, <MRL = below Method Reporting Limit
"""

from __future__ import annotations

import io

import pandas as pd

from parsers.base import BaseParser
from core.lab_value_parser import LabValueParser


class AlchemSoilParser(BaseParser):
    LAB_NAME = "Alchem Soil"

    def __init__(self):
        self._vp = LabValueParser()

    # ------------------------------------------------------------------
    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        xl = pd.ExcelFile(file_obj)
        sheet_names_lower = {s.lower(): s for s in xl.sheet_names}

        records = []

        # --- Parse VOC sheet ---
        if "voc" in sheet_names_lower:
            voc_records = self._parse_voc_sheet(xl, sheet_names_lower["voc"])
            records.extend(voc_records)
        else:
            # Fallback: try first sheet
            voc_records = self._parse_voc_sheet(xl, xl.sheet_names[0])
            records.extend(voc_records)

        # --- Parse TPH sheet ---
        if "tph" in sheet_names_lower:
            tph_records = self._parse_tph_sheet(xl, sheet_names_lower["tph"])
            records.extend(tph_records)

        return records

    # ------------------------------------------------------------------
    def _parse_voc_sheet(self, xl: pd.ExcelFile, sheet_name: str) -> list[dict]:
        raw = xl.parse(sheet_name, header=None, dtype=str).fillna("")

        # Find header row (contains "Compound" and "CAS")
        header_row = self._find_header_row(raw)

        # Extract sample IDs from row above header ("Analysis Location")
        sample_ids = self._extract_sample_ids(raw, header_row)

        # Parse column headers
        headers = [str(v).strip() for v in raw.iloc[header_row].values]

        col_compound = self._find_col_idx(headers, ["compound name", "compound", "analyte"])
        col_cas      = self._find_col_idx(headers, ["cas number", "cas no", "cas"])
        col_lod      = self._find_col_idx(headers, ["lod", "mdl", "method detection"])
        col_loq      = self._find_col_idx(headers, ["loq", "mrl", "method reporting"])

        # Final Concentration columns — one per sample
        conc_cols = [i for i, h in enumerate(headers)
                     if "final conc" in h.lower() or
                        ("concentration" in h.lower() and "final" in h.lower())]

        # Fallback: if no "final conc" found, use columns after col 4
        if not conc_cols:
            fixed_cols = max(
                c for c in [col_compound, col_cas, col_lod, col_loq] if c is not None
            ) + 1
            conc_cols = list(range(fixed_cols, len(headers)))

        if col_compound is None or col_cas is None:
            raise ValueError(
                f"❌ לא נמצאו עמודות Compound/CAS ב-VOC sheet "
                f"(row {header_row}). כותרות: {headers}"
            )

        records = []
        data_rows = raw.iloc[header_row + 1:].reset_index(drop=True)

        for _, row in data_rows.iterrows():
            values = list(row.values)
            compound = str(values[col_compound]).strip() if col_compound < len(values) else ""
            cas      = str(values[col_cas]).strip()      if col_cas      < len(values) else ""

            if not compound or compound.lower() in ("", "nan", "compound name"):
                continue
            if "total" in compound.lower() and "voc" in compound.lower():
                continue

            # Handle dual-CAS (e.g. "108-38-3 106-42-3" for m/p-Xylene)
            if " " in cas:
                cas = cas.split()[0]

            # Global LOD / LOQ for this compound
            lod = self._parse_float(values[col_lod]) if col_lod is not None else None
            loq = self._parse_float(values[col_loq]) if col_loq is not None else None

            for i, col_idx in enumerate(conc_cols):
                raw_val   = str(values[col_idx]).strip() if col_idx < len(values) else ""
                sample_id = sample_ids[i] if i < len(sample_ids) else f"Sample-{i+1}"

                # Map detection flags
                if raw_val.upper() in ("N.D.", "ND", "N/D", "NOT DETECTED", ""):
                    value = lod
                    flag  = "ND"
                elif raw_val.lower() in ("<mdl", "<dl"):
                    value = lod
                    flag  = "ND"
                elif raw_val.lower() in ("<mrl", "<loq", "<rl"):
                    value = loq
                    flag  = "<LOQ"
                else:
                    value, flag = self._vp.parse(raw_val)

                records.append({
                    "lab":       self.LAB_NAME,
                    "sample_id": sample_id,
                    "compound":  compound,
                    "cas":       cas,
                    "value":     value,
                    "flag":      flag,
                    "unit":      "mg/kg",
                    "lod":       lod,
                    "loq":       loq,
                    "analysis_type": "VOC",
                })

        return records

    # ------------------------------------------------------------------
    def _parse_tph_sheet(self, xl: pd.ExcelFile, sheet_name: str) -> list[dict]:
        """Parse TPH sheet: Sample Name | DRO | ORO | Total TPH"""
        raw = xl.parse(sheet_name, header=0, dtype=str).fillna("")

        records = []
        for _, row in raw.iterrows():
            sample_id = str(row.iloc[0]).strip()
            if not sample_id or sample_id.lower() in ("nan", "sample name"):
                continue

            # Each parameter as a separate compound record
            tph_params = []
            for col_name in raw.columns[1:]:
                raw_val = str(row[col_name]).strip()
                if raw_val.lower() in ("nan", ""):
                    continue

                if raw_val.upper() in ("N.D.", "ND", "N/D"):
                    value = None
                    flag  = "ND"
                else:
                    try:
                        value = float(raw_val)
                        flag  = None
                    except ValueError:
                        value, flag = self._vp.parse(raw_val)

                # Extract just the parameter name (e.g. "DRO" from "DRO [mg/kg]")
                param_name = col_name.split("[")[0].strip()

                records.append({
                    "lab":           self.LAB_NAME,
                    "sample_id":     sample_id,
                    "compound":      param_name,
                    "cas":           "",
                    "value":         value,
                    "flag":          flag,
                    "unit":          "mg/kg",
                    "lod":           None,
                    "loq":           None,
                    "analysis_type": "TPH",
                })

        return records

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _find_header_row(self, df: pd.DataFrame) -> int:
        for i, row in df.iterrows():
            row_str = " ".join(str(v).lower() for v in row.values)
            if "compound" in row_str and "cas" in row_str:
                return i
        return 1  # fallback

    def _extract_sample_ids(self, df: pd.DataFrame, header_row: int) -> list[str]:
        """Find sample IDs from the row containing 'Analysis Location'."""
        for i in range(header_row):
            row = df.iloc[i]
            row_str = " ".join(str(v) for v in row.values[:3]).lower()
            if "analysis location" in row_str or "location" in row_str:
                vals = [str(v).strip() for v in row.values]
                return [v for v in vals if v and v.lower() not in ("nan", "", "analysis location")]
        # Fallback: use non-empty values from header_row - 1
        if header_row > 0:
            row = df.iloc[header_row - 1]
            vals = [str(v).strip() for v in row.values]
            return [v for v in vals if v and v.lower() not in ("nan", "")]
        return []

    @staticmethod
    def _find_col_idx(headers: list[str], aliases: list[str]) -> int | None:
        for alias in aliases:
            for i, h in enumerate(headers):
                if alias.lower() in h.lower():
                    return i
        return None

    @staticmethod
    def _parse_float(val) -> float | None:
        try:
            return float(val)
        except (ValueError, TypeError):
            return None
