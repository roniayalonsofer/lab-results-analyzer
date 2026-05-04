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

  Sheet "ICP" (Metals):
    Row 0: Analysis Location: | ... | sample-id1 | sample-id2 | ...
    Row 1: Name | LOD [mg/kg] | LOQ. [mg/kg] | U.C. % | Final Conc. [mg/kg] | ...
    Row 2+: data rows ("Ag - Silver", "Al - Aluminum", etc.)
    Note: No CAS column — CAS resolved from element symbol.

  Sheet "pH" (field parameters):
    Row 0: Sample Name | PH   (header)
    Row 1+: sample rows (plain float pH values, dimensionless)

Units: mg/kg (not µg/m³)
Flags: <MDL = below Method Detection Limit, <MRL = below Method Reporting Limit
"""

from __future__ import annotations

import io
import re

import pandas as pd

from parsers.base import BaseParser
from core.lab_value_parser import LabValueParser


# Element symbol → CAS number (used to resolve ICP compound CAS from "Ag - Silver" format)
_METAL_SYMBOL_CAS: dict[str, str] = {
    "Ag": "7440-22-4",   # Silver
    "Al": "7429-90-5",   # Aluminum
    "As": "7440-38-2",   # Arsenic
    "Ba": "7440-39-3",   # Barium
    "Be": "7440-41-7",   # Beryllium
    "Cd": "7440-43-9",   # Cadmium
    "Co": "7440-48-4",   # Cobalt
    "Cr": "7440-47-3",   # Chromium
    "Cu": "7440-50-8",   # Copper
    "Fe": "7439-89-6",   # Iron
    "Hg": "7439-97-6",   # Mercury
    "Li": "7439-93-2",   # Lithium
    "Mn": "7439-96-5",   # Manganese
    "Mo": "7439-98-7",   # Molybdenum
    "Ni": "7440-02-0",   # Nickel
    "Pb": "7439-92-1",   # Lead
    "Sb": "7440-36-0",   # Antimony
    "Se": "7782-49-2",   # Selenium
    "Tl": "7440-28-0",   # Thallium
    "V":  "7440-62-2",   # Vanadium
    "Zn": "7440-66-6",   # Zinc
}


class AlchemSoilParser(BaseParser):
    LAB_NAME = "Alchem Soil"
    ANALYSIS_TYPES = ["SOIL_VOC", "SOIL_TPH", "SOIL_METALS", "LOWFLOW"]

    def __init__(self):
        self._vp = LabValueParser()

    # ------------------------------------------------------------------
    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        xl = pd.ExcelFile(file_obj)
        sheet_names_lower = {s.lower(): s for s in xl.sheet_names}

        records = []

        # Helper: find a sheet whose name contains any of the given substrings.
        # Uses word-boundary matching: the keyword must NOT be immediately preceded
        # by a letter (prevents "ph" from matching inside "tph").
        # Priority: (1) exact match, (2) word-boundary substring match.
        def _find_sheet(*keywords: str) -> str | None:
            # Pass 1 — exact match
            for kw in keywords:
                if kw in sheet_names_lower:
                    return sheet_names_lower[kw]
            # Pass 2 — keyword appears as a whole token
            # e.g. "ph" matches "ph", "40280-ph" but NOT "tph" or "40280-tph"
            for kw in keywords:
                pat = re.compile(r'(?<![a-z])' + re.escape(kw) + r'(?![a-z])')
                for key, name in sheet_names_lower.items():
                    if pat.search(key):
                        return name
            return None

        # --- Parse VOC sheet ---
        voc_sheet = _find_sheet("voc")
        if voc_sheet:
            records.extend(self._parse_voc_sheet(xl, voc_sheet))
        else:
            # Fallback: try first sheet
            records.extend(self._parse_voc_sheet(xl, xl.sheet_names[0]))

        # --- Parse TPH sheet ---
        tph_sheet = _find_sheet("tph")
        if tph_sheet:
            records.extend(self._parse_tph_sheet(xl, tph_sheet))

        # --- Parse ICP / Metals sheet ---
        icp_sheet = _find_sheet("icp", "metals", "metal")
        if icp_sheet:
            records.extend(self._parse_icp_sheet(xl, icp_sheet))

        # --- Parse pH / field parameters sheet ---
        ph_sheet = _find_sheet("ph", "field", "שדה")
        if ph_sheet:
            records.extend(self._parse_ph_sheet(xl, ph_sheet))

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
                    "analysis_type": "SOIL_VOC",
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

                # Normalize alternate total-TPH column names → canonical "TPH"
                # (dimer_1 uses "Total TPH", dimer_2 uses "TPH")
                if param_name.upper() in ("TOTAL TPH", "TOTAL-TPH", "TPH TOTAL"):
                    param_name = "TPH"

                # CAS: only the combined TPH total maps to the VSL threshold
                # (TPH - DRO + ORO (Tier 1), CAS C10-C40, VSL = 350 mg/kg).
                # DRO and ORO have no individual VSL threshold in the file.
                cas = "C10-C40" if param_name.upper() == "TPH" else ""

                records.append({
                    "lab":           self.LAB_NAME,
                    "sample_id":     sample_id,
                    "compound":      param_name,
                    "cas":           cas,
                    "value":         value,
                    "flag":          flag,
                    "unit":          "mg/kg",
                    "lod":           None,
                    "loq":           None,
                    "analysis_type": "SOIL_TPH",
                })

        return records

    # ------------------------------------------------------------------
    def _parse_icp_sheet(self, xl: pd.ExcelFile, sheet_name: str) -> list[dict]:
        """Parse ICP / Metals sheet → SOIL_METALS records.

        Layout is identical to VOC except there is no CAS column.
        CAS is resolved from the element symbol in the compound name
        string, e.g. "Ag - Silver" → symbol "Ag" → _METAL_SYMBOL_CAS lookup.
        """
        raw = xl.parse(sheet_name, header=None, dtype=str).fillna("")

        header_row = self._find_icp_header_row(raw)
        sample_ids = self._extract_sample_ids(raw, header_row)
        headers    = [str(v).strip() for v in raw.iloc[header_row].values]

        col_compound = self._find_col_idx(headers, ["name"])
        col_lod      = self._find_col_idx(headers, ["lod", "method detection"])
        col_loq      = self._find_col_idx(headers, ["loq", "method reporting"])

        # Final Conc. columns — one per sample
        conc_cols = [i for i, h in enumerate(headers)
                     if "final conc" in h.lower() or
                        ("concentration" in h.lower() and "final" in h.lower())]
        if not conc_cols:
            fixed_end = max(
                c for c in [col_compound, col_lod, col_loq] if c is not None
            ) + 1
            conc_cols = list(range(fixed_end, len(headers)))

        if col_compound is None:
            return []

        records = []
        data_rows = raw.iloc[header_row + 1:].reset_index(drop=True)

        for _, row in data_rows.iterrows():
            values   = list(row.values)
            compound = str(values[col_compound]).strip() if col_compound < len(values) else ""

            if not compound or compound.lower() in ("", "nan"):
                continue

            # CAS from element symbol: "Ag - Silver" → "Ag"
            symbol = compound.split(" - ")[0].strip()
            cas    = _METAL_SYMBOL_CAS.get(symbol, "")

            lod = self._parse_float(values[col_lod]) if col_lod is not None else None
            loq = self._parse_float(values[col_loq]) if col_loq is not None else None

            for i, col_idx in enumerate(conc_cols):
                raw_val   = str(values[col_idx]).strip() if col_idx < len(values) else ""
                sample_id = sample_ids[i] if i < len(sample_ids) else f"Sample-{i+1}"

                if raw_val.upper() in ("N.D.", "ND", "N/D", "NOT DETECTED", ""):
                    value, flag = lod, "ND"
                elif raw_val.lower() in ("<mdl", "<dl"):
                    value, flag = lod, "ND"
                elif raw_val.lower() in ("<mrl", "<loq", "<rl"):
                    value, flag = loq, "<LOQ"
                else:
                    value, flag = self._vp.parse(raw_val)

                records.append({
                    "lab":           self.LAB_NAME,
                    "sample_id":     sample_id,
                    "compound":      compound,   # keep full "Ag - Silver"
                    "cas":           cas,
                    "value":         value,
                    "flag":          flag,
                    "unit":          "mg/kg",
                    "lod":           lod,
                    "loq":           loq,
                    "analysis_type": "SOIL_METALS",
                })

        return records

    # ------------------------------------------------------------------
    def _parse_ph_sheet(self, xl: pd.ExcelFile, sheet_name: str) -> list[dict]:
        """Parse pH / field-parameters sheet → LOWFLOW records (no thresholds).

        Layout: Row 0 = header (Sample Name | PH | ...),
                Rows 1+ = sample rows with plain float values.
        Each parameter column becomes one LOWFLOW record per sample.
        """
        raw = xl.parse(sheet_name, header=0, dtype=str).fillna("")

        records = []
        for _, row in raw.iterrows():
            sample_id = str(row.iloc[0]).strip()
            if not sample_id or sample_id.lower() in ("nan", "sample name", ""):
                continue

            for col_name in raw.columns[1:]:
                raw_val = str(row[col_name]).strip()
                if raw_val.lower() in ("nan", "", "n.d.", "nd"):
                    value, flag = None, "ND"
                else:
                    try:
                        value = float(raw_val)
                        flag  = None
                    except ValueError:
                        value, flag = self._vp.parse(raw_val)

                # Strip units from column name if present (e.g. "pH [units]" → "pH")
                param_name = str(col_name).split("[")[0].strip()
                # Normalize uppercase "PH" → canonical "pH"
                if param_name.upper() == "PH":
                    param_name = "pH"

                records.append({
                    "lab":           self.LAB_NAME,
                    "sample_id":     sample_id,
                    "compound":      param_name,
                    "cas":           "",
                    "value":         value,
                    "flag":          flag,
                    "unit":          "",   # pH is dimensionless
                    "lod":           None,
                    "loq":           None,
                    "analysis_type": "LOWFLOW",
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

    def _find_icp_header_row(self, df: pd.DataFrame) -> int:
        """Find the data-header row in ICP sheet (contains 'lod' and 'name' or 'final')."""
        for i, row in df.iterrows():
            row_str = " ".join(str(v).lower() for v in row.values)
            if "lod" in row_str and ("name" in row_str or "final" in row_str):
                return i
        return 1  # fallback

    def _extract_sample_ids(self, df: pd.DataFrame, header_row: int) -> list[str]:
        """Find sample IDs from the row containing 'Analysis Location' or 'Sample Name'."""
        _SKIP = {"nan", "", "analysis location", "analysis location:",
                 "sample name", "sample name:"}
        for i in range(header_row):
            row = df.iloc[i]
            row_str = " ".join(str(v) for v in row.values[:3]).lower()
            if "analysis location" in row_str or "sample name" in row_str:
                vals = [str(v).strip() for v in row.values]
                return [v for v in vals if v and v.lower() not in _SKIP]
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
