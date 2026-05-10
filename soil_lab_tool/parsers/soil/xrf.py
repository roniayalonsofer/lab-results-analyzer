"""
parsers/soil/xrf.py
--------------------
Generic parser for XRF (X-Ray Fluorescence) soil metals reports.

Expected file format — wide CSV or XLSX:
  Row 0 (or auto-detected): column headers
    Column 0     : Sample ID  (any name: "Sample", "ID", "Sample ID", …)
    Column 1     : Location   (optional: "Location", "Site", "Borehole", …)
    Columns 2+   : Element symbols (Mo, Zr, Sr, Pb, As, …) with optional unit
                   in header, e.g. "Pb (mg/kg)" or just "Pb"

  Each data row is one sample.  Values may be numeric, "<LOQ", "ND", "<DL", etc.
  Units default to "mg/kg" (XRF soil standard); overridden per column if explicit.
"""

from __future__ import annotations

import io
import re

import pandas as pd

from parsers.base import BaseParser
from core.lab_value_parser import LabValueParser


# ── Element → CAS registry ────────────────────────────────────────────────────

# Symbol (upper-case) → CAS number for elements commonly measured by XRF
_ELEMENT_CAS: dict[str, str] = {
    "AG": "7440-22-4",   # Silver
    "AL": "7429-90-5",   # Aluminium
    "AS": "7440-38-2",   # Arsenic
    "AU": "7440-57-5",   # Gold
    "BA": "7440-39-3",   # Barium
    "BE": "7440-41-7",   # Beryllium
    "BI": "7440-69-9",   # Bismuth
    "CA": "7440-70-2",   # Calcium
    "CD": "7440-43-9",   # Cadmium
    "CE": "7440-45-1",   # Cerium
    "CO": "7440-48-4",   # Cobalt
    "CR": "7440-47-3",   # Chromium
    "CS": "7440-46-2",   # Caesium
    "CU": "7440-50-8",   # Copper
    "FE": "7439-89-6",   # Iron
    "GA": "7440-55-3",   # Gallium
    "GE": "7440-56-4",   # Germanium
    "HG": "7439-97-6",   # Mercury
    "IN": "7440-74-6",   # Indium
    "K":  "7440-09-7",   # Potassium
    "LA": "7439-91-0",   # Lanthanum
    "MN": "7439-96-5",   # Manganese
    "MO": "7439-98-7",   # Molybdenum
    "NB": "7440-03-1",   # Niobium
    "NI": "7440-02-0",   # Nickel
    "P":  "7723-14-0",   # Phosphorus
    "PB": "7439-92-1",   # Lead
    "RB": "7440-17-7",   # Rubidium
    "RE": "7440-15-5",   # Rhenium
    "S":  "7704-34-9",   # Sulphur
    "SB": "7440-36-0",   # Antimony
    "SE": "7782-49-2",   # Selenium
    "SI": "7440-21-3",   # Silicon
    "SN": "7440-31-5",   # Tin
    "SR": "7440-24-6",   # Strontium
    "TE": "13494-80-9",  # Tellurium
    "TH": "7440-29-1",   # Thorium
    "TI": "7440-32-6",   # Titanium
    "TL": "7440-28-0",   # Thallium
    "U":  "7440-61-1",   # Uranium
    "V":  "7440-62-2",   # Vanadium
    "W":  "7440-33-7",   # Tungsten
    "Y":  "7440-65-5",   # Yttrium
    "ZN": "7440-66-6",   # Zinc
    "ZR": "7440-67-7",   # Zirconium
}

# Full English name for each element (used as compound name in records)
_ELEMENT_NAME: dict[str, str] = {
    "AG": "Silver",      "AL": "Aluminium",   "AS": "Arsenic",
    "AU": "Gold",        "BA": "Barium",      "BE": "Beryllium",
    "BI": "Bismuth",     "CA": "Calcium",     "CD": "Cadmium",
    "CE": "Cerium",      "CO": "Cobalt",      "CR": "Chromium",
    "CS": "Caesium",     "CU": "Copper",      "FE": "Iron",
    "GA": "Gallium",     "GE": "Germanium",   "HG": "Mercury",
    "IN": "Indium",      "K":  "Potassium",   "LA": "Lanthanum",
    "MN": "Manganese",   "MO": "Molybdenum",  "NB": "Niobium",
    "NI": "Nickel",      "P":  "Phosphorus",  "PB": "Lead",
    "RB": "Rubidium",    "RE": "Rhenium",     "S":  "Sulphur",
    "SB": "Antimony",    "SE": "Selenium",    "SI": "Silicon",
    "SN": "Tin",         "SR": "Strontium",   "TE": "Tellurium",
    "TH": "Thorium",     "TI": "Titanium",    "TL": "Thallium",
    "U":  "Uranium",     "V":  "Vanadium",    "W":  "Tungsten",
    "Y":  "Yttrium",     "ZN": "Zinc",        "ZR": "Zirconium",
}

# Columns that are never elements (metadata columns)
_SKIP_COLS = frozenset({
    "sample", "sample id", "sample_id", "id", "sampleid",
    "location", "site", "borehole", "well", "point", "station",
    "description", "תיאור", "מזהה", "מספר", "שם", "קידוח",
    "date", "depth", "unit", "units", "comment", "notes", "remarks",
    "latitude", "longitude", "lat", "lon", "northing", "easting",
    "field", "method", "matrix",
})

# Columns that contain the sample location / description
_LOCATION_HINTS = frozenset({
    "location", "site", "borehole", "well", "point", "station",
    "description", "name", "label", "תיאור", "מיקום", "קידוח",
})

# Header row indicators (row contains these → it's the header)
_HEADER_MARKERS = frozenset({"sample", "id", "location", "site", "pb", "zn", "as", "cu", "ni"})

# Unit pattern embedded in a column header, e.g. "Pb (mg/kg)" or "As [ppm]"
_UNIT_RE = re.compile(r"\(([^)]+)\)|\[([^\]]+)\]")


def _parse_header_col(raw: str) -> tuple[str, str]:
    """
    Strip an optional unit from a column header and return (symbol_upper, unit).
    E.g. 'Pb (mg/kg)' → ('PB', 'mg/kg');  'As' → ('AS', 'mg/kg').
    """
    m = _UNIT_RE.search(raw)
    unit = (m.group(1) or m.group(2)).strip() if m else "mg/kg"
    sym = _UNIT_RE.sub("", raw).strip().upper()
    return sym, unit


def _find_header_row(df: pd.DataFrame) -> int:
    """
    Scan the first 10 rows to find the header row.
    Returns the row index whose cells best match known element symbols or
    the skip-column keywords.
    """
    best_row, best_score = 0, 0
    for ri in range(min(10, len(df))):
        row_vals = [str(v).strip().upper() for v in df.iloc[ri]]
        score = sum(
            1 for v in row_vals
            if v in _ELEMENT_CAS or v.lower() in _SKIP_COLS
        )
        if score > best_score:
            best_score, best_row = score, ri
    return best_row


class XRFSoilParser(BaseParser):
    """
    Parse wide-format XRF soil metals reports (CSV or XLSX).

    Output analysis_type: SOIL_METALS (same sheet + thresholds as ICP metals).
    """

    LAB_NAME = "XRF"
    ANALYSIS_TYPES = ["SOIL_METALS"]

    def __init__(self):
        self._vp = LabValueParser()

    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        file_obj.seek(0)
        header4 = file_obj.read(4)
        file_obj.seek(0)

        # Distinguish CSV (text) from XLSX (ZIP magic PK\x03\x04 or XLS D0CF)
        if header4[:2] == b"PK" or header4[:2] == b"\xd0\xcf":
            df_raw = self._read_excel(file_obj)
        else:
            df_raw = self._read_csv(file_obj)

        if df_raw is None or df_raw.empty:
            return []

        return self._parse_wide(df_raw)

    # ------------------------------------------------------------------

    def _read_excel(self, file_obj: io.BytesIO) -> pd.DataFrame | None:
        try:
            xl = pd.ExcelFile(file_obj)
            return xl.parse(xl.sheet_names[0], header=None, dtype=str).fillna("")
        except Exception:
            return None

    def _read_csv(self, file_obj: io.BytesIO) -> pd.DataFrame | None:
        for enc in ("utf-8-sig", "utf-8", "cp1255", "latin-1"):
            try:
                file_obj.seek(0)
                return pd.read_csv(file_obj, header=None, dtype=str,
                                   encoding=enc).fillna("")
            except Exception:
                continue
        return None

    def _parse_wide(self, df_raw: pd.DataFrame) -> list[dict]:
        hdr_idx = _find_header_row(df_raw)

        # Build column map from the detected header row
        header_vals = [str(v).strip() for v in df_raw.iloc[hdr_idx]]

        # Identify sample-ID column (leftmost non-empty, or best keyword match)
        id_col: int | None = None
        loc_col: int | None = None
        element_cols: list[tuple[int, str, str, str]] = []  # (col_idx, symbol, name, unit)

        for ci, raw_hdr in enumerate(header_vals):
            if not raw_hdr or raw_hdr.lower() in ("nan", ""):
                continue
            low = raw_hdr.strip().lower()
            sym, unit = _parse_header_col(raw_hdr)

            if low in ("sample", "sample id", "sample_id", "sampleid", "id",
                       "מזהה", "מספר"):
                id_col = id_col if id_col is not None else ci
                continue
            if low in _LOCATION_HINTS:
                loc_col = ci
                continue
            if low in _SKIP_COLS:
                continue
            if sym in _ELEMENT_CAS:
                name = _ELEMENT_NAME.get(sym, sym.capitalize())
                cas  = _ELEMENT_CAS[sym]
                element_cols.append((ci, sym, name, unit))

        # Fallback: treat column 0 as sample ID if nothing matched
        if id_col is None:
            id_col = 0

        records: list[dict] = []
        for ri in range(hdr_idx + 1, len(df_raw)):
            row = df_raw.iloc[ri]
            sid_raw = str(row.iloc[id_col]).strip()
            if not sid_raw or sid_raw.lower() in ("nan", "", "sample", "id"):
                continue

            sample_id = sid_raw
            if loc_col is not None:
                loc_val = str(row.iloc[loc_col]).strip()
                if loc_val and loc_val.lower() not in ("nan", ""):
                    sample_id = f"{sid_raw} – {loc_val}"

            for ci, sym, compound, unit in element_cols:
                raw_val = str(row.iloc[ci]).strip()
                if not raw_val or raw_val.lower() in ("nan", ""):
                    continue

                value, flag = self._vp.parse(raw_val)
                if value is None and not flag:
                    continue

                records.append({
                    "sample_id":     sample_id,
                    "compound":      compound,
                    "cas":           _ELEMENT_CAS.get(sym, ""),
                    "value":         value,
                    "flag":          flag,
                    "unit":          unit,
                    "lod":           None,
                    "loq":           None,
                    "analysis_type": "SOIL_METALS",
                })

        return records
