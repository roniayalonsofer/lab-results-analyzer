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

  Each data row is one sample.  Values may be numeric, "< LOD", "ND", "<DL", etc.
  Units default to "mg/kg" (XRF soil standard); overridden per column if explicit.
"""

from __future__ import annotations

import io
import re

import pandas as pd

from parsers.base import BaseParser


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
    # XRF instrument metadata columns
    "duration", "time", "reading", "reading #", "reading#", "type",
    "dose", "instrument", "operator", "sequence", "sequence #",
    "application", "user", "group", "alloy", "grade", "pass/fail",
    "pass", "fail", "quality",
})

# Column headers that identify the sample ID — checked before _SKIP_COLS
_SAMPLE_ID_HINTS = frozenset({
    "sample", "sample id", "sample_id", "sampleid", "id",
    "מזהה", "מספר",
    # Olympus / Bruker / ThermoFisher handheld XRF exports
    "sample name", "sample_name", "label", "user label", "user_label",
    "point id", "point_id", "no", "no.", "#", "reading #", "reading#",
    "test #", "test#",
})

# Columns that contain the sample location / description
_LOCATION_HINTS = frozenset({
    "location", "site", "borehole", "well", "point", "station",
    "description", "name", "label", "תיאור", "מיקום", "קידוח",
})

# Unit pattern embedded in a column header, e.g. "Pb (mg/kg)" or "As [ppm]"
_UNIT_RE = re.compile(r"\(([^)]+)\)|\[([^\]]+)\]")

# Non-detect text patterns found in XRF data cells (before numeric < patterns)
_ND_TEXT_RE = re.compile(
    r'^(?:[<\s]*(?:lod|loq|dl|mdl|bdl|bql|nd|n\.d\.?|not\s+detected|below|'
    r'no\s+detect|not\s+detect)|---+|-+|n/a|na)$',
    re.IGNORECASE,
)
# Numeric below-detection: "< 0.001", "<0.5", "< 1.23E-4"
_BELOW_NUM_RE = re.compile(
    r'^[<＜]\s*(\d+\.?\d*(?:[eE][+-]?\d+)?)\s*$'
)


def _parse_xrf_value(raw: str) -> tuple[float | None, str, float | None]:
    """
    Parse one XRF data cell.

    Returns (value, flag, lod) where:
        value : float | None — numeric measurement (None for non-detects)
        flag  : ''           — detected value
                '<LOD'       — below numeric detection limit  (lod = the limit)
                'ND'         — non-detect (no numeric limit given)
                '<'          — explicit less-than with no recognised LOD keyword
        lod   : float | None — detection limit when known
    """
    s = raw.strip()
    if not s or s.lower() in ("nan", ""):
        return None, "ND", None

    # Text-based non-detects: "< LOD", "< DL", "ND", "BDL", "N/A", …
    if _ND_TEXT_RE.match(s):
        return None, "ND", None

    # Numeric below-detection: "< 0.001"  →  lod=0.001, value=None, flag='<LOD'
    m = _BELOW_NUM_RE.match(s)
    if m:
        lod_num = float(m.group(1))
        return None, "<LOD", lod_num

    # Regular positive number
    try:
        return float(s), "", None
    except ValueError:
        pass

    # Fallback: still looks like a less-than but with text after (e.g. "< LOD 0.5")
    if s.startswith("<") or s.startswith("＜"):
        return None, "ND", None

    return None, "ND", None


def _parse_header_col(raw: str) -> tuple[str, str]:
    """
    Strip an optional unit from a column header and return (symbol_upper, unit).
    E.g. 'Pb (mg/kg)' → ('PB', 'mg/kg');  'As' → ('AS', 'mg/kg').
    Only the FIRST word (before any space/underscore) is used as the symbol,
    so 'Pb_ppm', 'Pb ppm', 'Pb total' all resolve to 'PB'.
    """
    m = _UNIT_RE.search(raw)
    unit = (m.group(1) or m.group(2)).strip() if m else "mg/kg"
    # Strip unit notation, then take only the first token (handles "Pb ppm", "Pb_ppm")
    base = _UNIT_RE.sub("", raw).strip()
    sym  = re.split(r'[\s_\-/]', base)[0].upper()
    return sym, unit


def _find_header_row(df: pd.DataFrame) -> int:
    """
    Scan the first 10 rows to find the header row.
    A row scores +1 for each cell that is a known element symbol or a
    recognised metadata keyword.  The row with the highest score is the header.
    Ties are broken by preferring the earlier row.
    """
    best_row, best_score = 0, 0
    for ri in range(min(10, len(df))):
        row_vals = [str(v).strip() for v in df.iloc[ri]]
        score = 0
        for v in row_vals:
            up = v.upper()
            lo = v.lower()
            if up in _ELEMENT_CAS:
                score += 1
            elif lo in _SKIP_COLS or lo in _SAMPLE_ID_HINTS or lo in _LOCATION_HINTS:
                score += 1
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
            # Find the data sheet: prefer one with many element-symbol headers
            best_sheet, best_score = xl.sheet_names[0], -1
            for name in xl.sheet_names:
                try:
                    df = xl.parse(name, header=None, dtype=str, nrows=8).fillna("")
                    for ri in range(len(df)):
                        row = [str(v).strip().upper() for v in df.iloc[ri]]
                        score = sum(1 for v in row if v in _ELEMENT_CAS)
                        if score > best_score:
                            best_score, best_sheet = score, name
                except Exception:
                    continue
            return xl.parse(best_sheet, header=None, dtype=str).fillna("")
        except Exception:
            return None

    def _read_csv(self, file_obj: io.BytesIO) -> pd.DataFrame | None:
        # Try comma and semicolon separators (European CSVs use ;)
        for sep in (",", ";", "\t"):
            for enc in ("utf-8-sig", "utf-8", "cp1255", "latin-1"):
                try:
                    file_obj.seek(0)
                    df = pd.read_csv(file_obj, header=None, dtype=str,
                                     sep=sep, encoding=enc).fillna("")
                    # Accept this df only if it has enough columns to plausibly
                    # be a proper XRF file (at least 3 columns)
                    if df.shape[1] >= 3:
                        return df
                except Exception:
                    continue
        return None

    def _parse_wide(self, df_raw: pd.DataFrame) -> list[dict]:
        hdr_idx = _find_header_row(df_raw)

        # Build column map from the detected header row
        header_vals = [str(v).strip() for v in df_raw.iloc[hdr_idx]]

        id_col: int | None = None
        loc_col: int | None = None
        element_cols: list[tuple[int, str, str, str]] = []  # (col_idx, symbol, name, unit)

        for ci, raw_hdr in enumerate(header_vals):
            if not raw_hdr or raw_hdr.lower() in ("nan", ""):
                continue
            low = raw_hdr.strip().lower()
            sym, unit = _parse_header_col(raw_hdr)

            # Sample ID — check against broader hint set first
            if low in _SAMPLE_ID_HINTS:
                if id_col is None:
                    id_col = ci
                continue
            # Location / site column
            if low in _LOCATION_HINTS and low not in _SAMPLE_ID_HINTS:
                loc_col = ci
                continue
            # Metadata skip
            if low in _SKIP_COLS:
                continue
            # Element column — symbol must exactly match _ELEMENT_CAS
            if sym in _ELEMENT_CAS:
                name = _ELEMENT_NAME.get(sym, sym.capitalize())
                element_cols.append((ci, sym, name, unit))

        # Fallback: treat column 0 as sample ID if nothing matched
        if id_col is None:
            id_col = 0

        # If id_col and loc_col are the same (misconfigured), clear loc_col
        if loc_col == id_col:
            loc_col = None

        _SKIP_VALUES = frozenset({"nan", "", "sample", "sample id", "id", "no", "#"})

        records: list[dict] = []
        for ri in range(hdr_idx + 1, len(df_raw)):
            row = df_raw.iloc[ri]
            sid_raw = str(row.iloc[id_col]).strip()
            if not sid_raw or sid_raw.lower() in _SKIP_VALUES:
                continue

            loc_val = ""
            if loc_col is not None:
                _lv = str(row.iloc[loc_col]).strip()
                if _lv and _lv.lower() not in ("nan", ""):
                    loc_val = _lv

            for ci, sym, compound, unit in element_cols:
                raw_val = str(row.iloc[ci]).strip()
                if not raw_val or raw_val.lower() in ("nan", ""):
                    continue

                value, flag, lod = _parse_xrf_value(raw_val)

                records.append({
                    "sample_id":     sid_raw,   # raw sample number only
                    "location":      loc_val,   # separate location field
                    "compound":      compound,
                    "cas":           _ELEMENT_CAS.get(sym, ""),
                    "value":         value,
                    "flag":          flag,
                    "unit":          unit,
                    "lod":           lod,
                    "loq":           None,
                    "analysis_type": "SOIL_METALS",
                })

        return records
