"""
parsers/soil/als.py
--------------------
Parser for ALS laboratory soil reports.

Sheet format ("Client SOIL 1"):
  Row  8 (idx 7 ): sample IDs in columns idx 4+
  Row 13 (idx 12): column headers — Parameter | Method | Unit | LOR | <sample cols>
  Rows 14+       : compound at idx 0, unit at idx 2, LOR at idx 3, values at idx 4+

  Values like "<0.050" = below LOR (flag <LOQ, value = numeric after <).
  "ND" / "N.D." = not detected (flag ND, value = LOR when available).

ALSGrainSizeParser uses the same sheet format but recognises fraction parameters
("Fraction X-Y mm") and tags records as SOIL_GRAIN_SIZE.
"""

from __future__ import annotations

import io

import pandas as pd

from parsers.base import BaseParser
from core.cas_lookup import name_to_cas


# ---------------------------------------------------------------------------
# Shared format reader
# ---------------------------------------------------------------------------

def _parse_als_sheet(xl: pd.ExcelFile, sheet_name: str) -> tuple[list[str], dict[int, str], list[tuple]]:
    """
    Parse an ALS "Client SOIL" sheet.

    Returns
    -------
    (header_values, sample_cols, data_rows)

    header_values : list of str — the header row (row 13, idx 12)
    sample_cols   : {col_idx: sample_id}
    data_rows     : list of (compound, unit, lor_val, {col_idx: raw_str})
    """
    raw = xl.parse(sheet_name, header=None, dtype=str).fillna("")

    # ── Locate header row (row 13 = idx 12; verify by finding "LOR" / "Parameter") ──
    hdr_row_idx = 12
    for ri in range(8, min(20, len(raw))):
        row_vals = [str(v).strip().upper() for v in raw.iloc[ri]]
        if "LOR" in row_vals or "PARAMETER" in row_vals:
            hdr_row_idx = ri
            break

    hdr = [str(v).strip() for v in raw.iloc[hdr_row_idx]]

    # Find key column indices from header
    lor_col       = next((ci for ci, h in enumerate(hdr) if h.upper() == "LOR"),  3)
    unit_col      = next((ci for ci, h in enumerate(hdr) if h.upper() == "UNIT"), lor_col - 1)
    compound_col  = next((ci for ci, h in enumerate(hdr) if h.upper() in ("PARAMETER", "COMPOUND", "ANALYTE")), 0)
    first_smp_col = lor_col + 1

    # ── Sample IDs from row 8 (idx 7) — same column positions as data ──
    sample_row = raw.iloc[7]
    sample_cols: dict[int, str] = {}
    for ci in range(first_smp_col, len(sample_row)):
        sid = str(sample_row.iloc[ci]).strip()
        if sid and sid.lower() not in ("nan", ""):
            sample_cols[ci] = sid

    # If row 8 yielded nothing, fall back to header row sample columns
    if not sample_cols:
        for ci in range(first_smp_col, len(hdr)):
            sid = hdr[ci]
            if sid and sid.lower() not in ("nan", ""):
                sample_cols[ci] = sid

    # ── Data rows ──
    data_rows = []
    for ri in range(hdr_row_idx + 1, len(raw)):
        row = raw.iloc[ri]
        compound = str(row.iloc[compound_col]).strip()
        if not compound or compound.lower() in ("nan", "", "parameter", "compound", "analyte"):
            continue

        unit_val = str(row.iloc[unit_col]).strip() if unit_col < len(row) else "mg/kg"
        if not unit_val or unit_val.lower() == "nan":
            unit_val = "mg/kg"

        loq: float | None = None
        lor_raw = str(row.iloc[lor_col]).strip().lstrip("<") if lor_col < len(row) else ""
        try:
            loq = float(lor_raw)
        except (ValueError, TypeError):
            pass

        sample_vals: dict[int, str] = {}
        for ci in sample_cols:
            if ci < len(row):
                sample_vals[ci] = str(row.iloc[ci]).strip()

        data_rows.append((compound, unit_val, loq, sample_vals))

    return hdr, sample_cols, data_rows


def _parse_value(raw_val: str, loq: float | None) -> tuple[float | None, str | None]:
    """Parse a raw cell string into (value, flag)."""
    v = raw_val.strip()
    if not v or v.lower() == "nan":
        return None, None
    if v.upper() in ("ND", "N.D.", "N/D", "<LOR", "< LOR", "NOT DETECTED"):
        return loq, "ND"
    if v.startswith("<"):
        try:
            num = float(v[1:])
            return num, "<LOQ"
        except ValueError:
            return loq, "<LOQ"
    try:
        return float(v), None
    except ValueError:
        return None, None


# ---------------------------------------------------------------------------
# ALSSoilParser
# ---------------------------------------------------------------------------

class ALSSoilParser(BaseParser):
    """ALS laboratory soil report — sheet 'Client SOIL 1'."""

    LAB_NAME = "ALS"

    _METALS = frozenset({
        "LEAD", "ZINC", "COPPER", "NICKEL", "CADMIUM", "CHROMIUM", "ARSENIC",
        "MERCURY", "BARIUM", "SILVER", "MANGANESE", "IRON", "ALUMINUM",
        "ALUMINIUM", "SELENIUM", "ANTIMONY", "BERYLLIUM", "COBALT",
        "MOLYBDENUM", "THALLIUM", "VANADIUM", "COBALT", "TIN",
    })
    _VOC = frozenset({
        "BENZENE", "TOLUENE", "XYLENE", "ETHYLBENZENE", "STYRENE",
        "NAPHTHALENE", "MTBE", "BTEX",
    })

    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        xl = pd.ExcelFile(file_obj)
        sheet = next((s for s in xl.sheet_names if "Client SOIL" in s), None)
        if sheet is None:
            raise ValueError(f"No 'Client SOIL' sheet found. Sheets: {xl.sheet_names}")

        _, sample_cols, data_rows = _parse_als_sheet(xl, sheet)
        records = []

        for compound, unit, loq, sample_vals in data_rows:
            cas = name_to_cas(compound)
            atype = self._analysis_type(compound)
            for ci, raw_val in sample_vals.items():
                value, flag = _parse_value(raw_val, loq)
                if value is None and flag is None:
                    continue
                records.append({
                    "compound":      compound,
                    "cas":           cas,
                    "value":         value,
                    "flag":          flag or "",
                    "unit":          unit,
                    "sample_id":     sample_cols[ci],
                    "lod":           None,
                    "loq":           loq,
                    "analysis_type": atype,
                })

        return records

    def _analysis_type(self, compound: str) -> str:
        c = compound.upper()
        if any(k in c for k in self._METALS):
            return "SOIL_METALS"
        if any(k in c for k in self._VOC):
            return "SOIL_VOC"
        if any(k in c for k in ("TPH", "PETROLEUM", "DRO", "ORO", "GRO")):
            return "SOIL_TPH"
        if any(k in c for k in ("PFAS", "PFOA", "PFOS", "PFBA")):
            return "SOIL_PFAS"
        return "SOIL_SVOC"


# ---------------------------------------------------------------------------
# ALSGrainSizeParser
# ---------------------------------------------------------------------------

class ALSGrainSizeParser(BaseParser):
    """ALS grain-size (sieve analysis) — sheet 'Client SOIL 1' with Fraction parameters."""

    LAB_NAME = "ALS"

    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        xl = pd.ExcelFile(file_obj)
        sheet = next((s for s in xl.sheet_names if "Client SOIL" in s), None)
        if sheet is None:
            raise ValueError(f"No 'Client SOIL' sheet found. Sheets: {xl.sheet_names}")

        _, sample_cols, data_rows = _parse_als_sheet(xl, sheet)
        records = []

        for compound, unit, loq, sample_vals in data_rows:
            cas = ""
            for ci, raw_val in sample_vals.items():
                value, flag = _parse_value(raw_val, loq)
                if value is None and flag is None:
                    continue
                records.append({
                    "compound":      compound,
                    "cas":           cas,
                    "value":         value,
                    "flag":          flag or "",
                    "unit":          unit or "%",
                    "sample_id":     sample_cols[ci],
                    "lod":           None,
                    "loq":           loq,
                    "analysis_type": "SOIL_GRAIN_SIZE",
                })

        return records
