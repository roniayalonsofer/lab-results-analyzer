"""
parsers/groundwater/kte.py
---------------------------
Parser for KTE laboratory groundwater reports in long-format (tidy) layout.

Same 24-column LIMS export as soil format (see parsers/soil/kte.py).

Handled analysis codes:
  BTEX_MTBE_DR_WATER → GW_VOC
  LOWFLOW            → LOWFLOW   (field parameters — no threshold comparison)

LOWFLOW parameters (field measurements, not lab analyses):
  Sampling depth, Total depth, Upper water level, Conductivity, Temperature,
  pH, DO, Turbidity, Redox
"""

from __future__ import annotations

import io
import re
import xml.etree.ElementTree as ET

import pandas as pd

from parsers.base import BaseParser
from core.lab_value_parser import LabValueParser
from core.cas_lookup import name_to_cas

# Groundwater compound CAS (KTE uses English names)
GW_CAS: dict[str, str] = {
    "benzene":        "71-43-2",
    "toluene":        "108-88-3",
    "ethyl benzene":  "100-41-4",
    "ethylbenzene":   "100-41-4",
    "xylene":         "1330-20-7",
    "xylenes":        "1330-20-7",
    "mtbe":           "1634-04-4",
    "methyl tert-butyl ether": "1634-04-4",
    "naphthalene":    "91-20-3",
}

GW_ANALYSIS_MAP: dict[str, str] = {
    "BTEX_MTBE_DR_WATER": "GW_VOC",
    "BTEX_MTBE_GW":       "GW_VOC",
    "GW_BTEX":            "GW_VOC",
    "LOWFLOW":            "LOWFLOW",
}


def _resolve_cas(compound: str) -> str:
    key = compound.strip().lower()
    if key in GW_CAS:
        return GW_CAS[key]
    return name_to_cas(compound) or ""


class KTEGroundwaterParser(BaseParser):
    LAB_NAME = "KTE"
    ANALYSIS_TYPES = ["GW_VOC", "LOWFLOW"]

    C_ORDER   = 0
    C_SAMPLE  = 1
    C_ACODE   = 2
    C_CPND    = 4
    C_RESULT  = 5
    C_UNIT    = 6
    C_DATE    = 8
    C_REPORT  = 11
    C_SITE    = 12
    C_LOC     = 13

    def __init__(self):
        self._vp = LabValueParser()

    # ------------------------------------------------------------------
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

            cas = _resolve_cas(compound) if atype == "GW_VOC" else ""

            sample_id = loc if loc and loc.lower() not in ("nan", "") else f"Sample-{row.iloc[self.C_SAMPLE]}"
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
                "unit":          unit or ("mg/L" if atype == "GW_VOC" else ""),
                "lod":           None,
                "analysis_type": atype,
            })

        return records

    # ------------------------------------------------------------------
    def _read(self, file_obj: io.BytesIO | str) -> pd.DataFrame | None:
        try:
            if isinstance(file_obj, str) and file_obj.lower().endswith(".csv"):
                # KTE CSV: data rows have more cols than header → pre-allocate 30 cols
                df = pd.read_csv(
                    file_obj, encoding="utf-8-sig", dtype=str,
                    header=None, engine="python",
                    names=list(range(30)), quotechar='"',
                ).fillna("")
                # Row 0 = Hebrew header, Row 1 = empty → skip both
                if df.shape[0] > 2:
                    first = str(df.iloc[0, 0]).strip()
                    if not first.replace("-", "").isdigit():
                        df = df.iloc[2:].reset_index(drop=True)
            else:
                # Read raw bytes to detect SpreadsheetML (XML-based XLS)
                if hasattr(file_obj, "read"):
                    raw = file_obj.read()
                else:
                    with open(file_obj, "rb") as f:
                        raw = f.read()
                sniff = raw.lstrip()[:200]
                if b"<?xml" in sniff or b"<Workbook" in sniff:
                    return self._read_spreadsheetml(raw)
                # Regular binary Excel
                xl = pd.ExcelFile(io.BytesIO(raw))
                df = xl.parse(xl.sheet_names[0], header=None, dtype=str).fillna("")
                first_cell = str(df.iloc[0, 0]).strip()
                if not first_cell.replace("-", "").isdigit():
                    df = df.iloc[2:].reset_index(drop=True)
        except ValueError:
            raise
        except Exception as e:
            raise ValueError(f"KTEGroundwaterParser: cannot read file — {e}") from e
        return df

    # ------------------------------------------------------------------
    def _read_spreadsheetml(self, raw: bytes) -> pd.DataFrame | None:
        """Parse KTE SpreadsheetML (XML-based .XLS) wide format → long DataFrame."""
        xml_str = re.sub(r'\s+xmlns(:\w+)?="[^"]*"', "", raw.decode("utf-8", errors="replace"))
        xml_str = re.sub(r"<(\w+):", "<", xml_str)
        xml_str = re.sub(r"</(\w+):", "</", xml_str)
        xml_str = re.sub(r"(\s)(\w+):", r"\1", xml_str)
        root = ET.fromstring(xml_str)

        # Find groundwater worksheet
        ws_xml = None
        for w in root.findall(".//Worksheet"):
            name = (w.get("Name") or "").strip().lower()
            if "groundwater" in name or "ground water" in name:
                ws_xml = w
                break
        if ws_xml is None:
            return None

        table = ws_xml.find(".//Table")
        if table is None:
            return None

        # Parse XML rows → list[list[str]]
        rows: list[list[str]] = []
        for row_el in table.findall("Row"):
            cells: list[str] = []
            prev_idx = 0
            for cell_el in row_el.findall("Cell"):
                idx_attr = cell_el.get("Index")
                if idx_attr is not None:
                    idx = int(idx_attr)
                    gap = idx - 1 - prev_idx
                    if gap > 0:
                        cells.extend([""] * gap)
                data_el = cell_el.find("Data")
                val = data_el.text if (data_el is not None and data_el.text) else ""
                cells.append(val.strip())
                prev_idx = len(cells) - 1
            rows.append(cells)

        # Locate sample ID and date header rows
        sample_row = date_row = None
        for r in rows:
            joined = " ".join(r).lower()
            if "client sample id" in joined:
                sample_row = r
            if "client sampling date" in joined:
                date_row = r
        if sample_row is None:
            return None

        well_names = sample_row[4:]
        dates = date_row[4:] if date_row else [""] * len(well_names)

        # BTEX/MTBE keyword patterns → GW_VOC
        BTEX_KEYS = ("benzene", "toluene", "ethylbenzene", "ethyl benzene",
                     "xylene", "naphthalene", "methyl tert-butyl", "mtbe")

        def _convert(val: str, unit: str) -> str:
            """Convert µg/L → mg/L, return as string."""
            factor = 0.001 if unit.strip() == "µg/L" else 1.0
            v = val.strip()
            if not v:
                return ""
            if v.startswith("<"):
                try:
                    return f"<{round(float(v[1:]) * factor, 6)}"
                except ValueError:
                    return v
            try:
                return str(round(float(v) * factor, 6))
            except ValueError:
                return v

        long_rows: list[list[str]] = []
        order = 0
        for r in rows:
            if not r or not r[0].strip():
                continue
            cpnd = r[0].strip()
            cpnd_key = cpnd.lower()
            if not any(k in cpnd_key for k in BTEX_KEYS):
                continue  # skip non-BTEX rows

            unit_raw = r[2].strip() if len(r) > 2 else ""
            out_unit = "mg/L"  # normalised output unit

            for wi, well in enumerate(well_names):
                if not well or well.lower() in ("nan", ""):
                    continue
                val_raw = r[4 + wi].strip() if (4 + wi) < len(r) else ""
                date_val = dates[wi] if wi < len(dates) else ""
                converted = _convert(val_raw, unit_raw)

                row = [""] * 14
                row[self.C_ORDER]  = str(order)
                row[self.C_SAMPLE] = well
                row[self.C_ACODE]  = "BTEX_MTBE_DR_WATER"
                row[self.C_CPND]   = cpnd
                row[self.C_RESULT] = converted
                row[self.C_UNIT]   = out_unit
                row[self.C_DATE]   = date_val
                row[self.C_LOC]    = well
                long_rows.append(row)
                order += 1

        if not long_rows:
            return None
        return pd.DataFrame(long_rows)

    @staticmethod
    def _resolve_atype(acode: str) -> str | None:
        for key, atype in GW_ANALYSIS_MAP.items():
            if key.upper() in acode or acode in key.upper():
                return atype
        return None

    @staticmethod
    def _short_date(date_str: str) -> str:
        m = re.search(r'(\d{4})-(\d{2})-(\d{2})', date_str)
        if m:
            return f"{m.group(3)}.{m.group(2)}.{m.group(1)}"
        return ""
