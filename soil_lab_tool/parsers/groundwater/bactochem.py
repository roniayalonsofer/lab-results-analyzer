"""
parsers/groundwater/bactochem.py
---------------------------------
Parser for בקטוכם (Bactochem) groundwater monitoring reports.

File format (CSV or XLSX):
  - Single header row (row 0) with Hebrew column names
  - Column 'רכיב'          : compound / parameter name (English)
  - Column 'תוצאה'         : result value or 'Not Detected'
  - Column 'תיאור דוגמה'   : sample location / borehole ID
  - Column 'תאריך דיגום'   : sampling datetime
  - Column 'אנליזה'        : analysis code (may or may not be present)

Analysis type mapping is done by compound name (not analysis code):
  GW_VOC  : Benzene, Toluene, Ethyl Benzene, Xylene, MTBE, Naphthalene, TBA
  LOWFLOW : pH, EC, Temperature, DO, Turbidity, Redox, depth params

GW thresholds sourced from soil_vsl_tier1_v7_2024.xlsx  "Groundwater" column:
  Benzene      1   mg/L
  Toluene    600   mg/L
  Ethylbenzene 700 mg/L
  Xylene     500   mg/L
  MTBE       240   mg/L
"""

from __future__ import annotations

import io
import re

import pandas as pd

from parsers.base import BaseParser
from core.lab_value_parser import LabValueParser
from core.cas_lookup import name_to_cas


# English compound names → CAS (same as KTE groundwater)
GW_CAS: dict[str, str] = {
    "benzene":                    "71-43-2",
    "toluene":                    "108-88-3",
    "ethyl benzene":              "100-41-4",
    "ethylbenzene":               "100-41-4",
    "xylene":                     "1330-20-7",
    "xylenes":                    "1330-20-7",
    "mtbe":                       "1634-04-4",
    "methyl tert-butyl ether":    "1634-04-4",
    "naphthalene":                "91-20-3",
    "tba":                        "75-65-0",   # tert-Butyl Alcohol
    "tert-butanol":               "75-65-0",
    "tert-butyl alcohol":         "75-65-0",
}

# Compound name keywords that indicate LOWFLOW (field parameters)
_LOWFLOW_KEYWORDS = (
    "ph",
    "conductivity",
    "temp",
    "dissolved o",
    "turbidity",
    "redox",
    "depth",
    "drilling",
    "sampling depth",
    "upper level",
    "water level",
)

# Compound names that are clearly GW_VOC analytes
_VOC_KEYWORDS = (
    "benzene", "toluene", "xylene", "mtbe", "naphthalene",
    "ethyl", "ethylbenzene", "tba", "tert-butyl",
)


def _resolve_cas(compound: str) -> str:
    key = compound.strip().lower()
    return GW_CAS.get(key) or name_to_cas(compound) or ""


def _classify_compound(name: str) -> str | None:
    """Return 'GW_VOC', 'LOWFLOW', or None (skip row)."""
    low = name.strip().lower()
    if any(k in low for k in _VOC_KEYWORDS):
        return "GW_VOC"
    if any(k in low for k in _LOWFLOW_KEYWORDS):
        return "LOWFLOW"
    return None


class BactochemGroundwaterParser(BaseParser):
    """
    Parses Bactochem groundwater lab reports (CSV or XLSX).

    Bactochem files use a **single** Hebrew header row (unlike KTE which has
    two). Compound names are English. Analysis type is inferred from the
    compound name rather than an analysis-code column.
    """

    LAB_NAME = "בקטוכם"
    ANALYSIS_TYPES = ["GW_VOC", "LOWFLOW"]

    # Named columns used by Bactochem
    COL_COMPOUND = "רכיב"
    COL_RESULT   = "תוצאה"
    COL_LOCATION = "תיאור דוגמה"
    COL_DATE     = "תאריך דיגום"

    def __init__(self):
        self._vp = LabValueParser()

    # ------------------------------------------------------------------
    def parse(self, file_obj: io.BytesIO | str) -> list[dict]:
        df = self._read(file_obj)
        if df is None or df.empty:
            return []

        records: list[dict] = []
        for _, row in df.iterrows():
            compound = str(row.get(self.COL_COMPOUND, "")).strip()
            raw_val  = str(row.get(self.COL_RESULT,   "")).strip()
            loc      = str(row.get(self.COL_LOCATION, "")).strip()
            date_val = str(row.get(self.COL_DATE,     "")).strip()

            if not compound or compound.lower() in ("nan", ""):
                continue

            atype = _classify_compound(compound)
            if atype is None:
                continue

            if raw_val.lower() in ("not detected", "nd", "n.d.", "n/d",
                                   "<dl", "none", ""):
                value, flag = None, "ND"
            else:
                value, flag = self._vp.parse(raw_val)

            cas = _resolve_cas(compound) if atype == "GW_VOC" else ""

            sample_id = (loc if loc and loc.lower() not in ("nan", "")
                         else "Sample")
            date_str = self._short_date(date_val)
            if date_str:
                sample_id = f"{sample_id} ({date_str})"

            records.append({
                "lab":           self.LAB_NAME,
                "sample_id":     sample_id,
                "compound":      compound,
                "cas":           cas,
                "value":         value,
                "flag":          flag,
                "unit":          "mg/L" if atype == "GW_VOC" else "",
                "lod":           None,
                "analysis_type": atype,
            })

        return records

    # ------------------------------------------------------------------
    def _read(self, file_obj: io.BytesIO | str) -> pd.DataFrame | None:
        """
        Read a Bactochem file.  The file has a single Hebrew header row
        (row 0 = column names; row 1 = first data row).
        """
        try:
            if isinstance(file_obj, str) and file_obj.lower().endswith(".csv"):
                df = pd.read_csv(
                    file_obj, encoding="utf-8-sig", dtype=str,
                    engine="python", usecols=list(range(21)),
                ).fillna("")
            else:
                xl = pd.ExcelFile(file_obj)
                raw = xl.parse(xl.sheet_names[0], header=None,
                               dtype=str).fillna("")
                # Auto-detect: if row 0 is a Hebrew header (not a data row),
                # use it as column names; otherwise fall back to positional access.
                first_cell = str(raw.iloc[0, 0]).strip()
                if not first_cell.replace("-", "").isdigit():
                    df = raw.iloc[1:].reset_index(drop=True)
                    df.columns = [str(c).strip() for c in raw.iloc[0].values]
                else:
                    df = raw
                df = df.fillna("")

            # Normalise column names (strip whitespace)
            df.columns = [str(c).strip() for c in df.columns]
            return df

        except Exception as e:
            raise ValueError(
                f"BactochemGroundwaterParser: cannot read file — {e}"
            ) from e

    @staticmethod
    def _short_date(date_str: str) -> str:
        m = re.search(r"(\d{4})-(\d{2})-(\d{2})", date_str)
        if m:
            return f"{m.group(3)}.{m.group(2)}.{m.group(1)}"
        return ""
