"""
parsers/soil/kte.py
--------------------
Parser for KTE laboratory soil reports in long-format (tidy) layout.

Supports both .xlsx and .csv exports from KTE's LIMS system.

File structure (24 columns, row 0 = Hebrew header, row 1 = empty, row 2+ = data):
  col 0  – מספר תעודה   (order/certificate ID)
  col 1  – מספר דוגמה   (sample ID — unique per sample)
  col 2  – אנליזה       (analysis code)
  col 3  – תיאור אנליזה (analysis description in Hebrew)
  col 4  – רכיב         (compound / parameter name)
  col 5  – תוצאה        (result value or "Not Detected")
  col 6  – יחידות מידה  (units)
  col 7  – לקוח         (client code)
  col 8  – תאריך דיגום  (sampling datetime)
  col 9  – מוצר         (matrix: SOIL / GROUND_WATER)
  col 10 – סוג דוגמה    (medium: SOIL / WATER)
  col 11 – מספר פרויקט  (report / project number)
  col 12 – אתר דיגום    (site name, Hebrew)
  col 13 – תיאור דוגמה  (sample location / borehole ID, e.g. ק-1)
  col 14 – הערות         (notes / depth)
  col 15 – מספר רכיב    (parameter order index)

Handled analysis codes → analysis_type:
  BTEX_MTBE_SOIL_WS_WT  → SOIL_VOC
  TPH_DRO_ORO           → SOIL_TPH
  METALS / ICP / ICP_OES → SOIL_METALS
  SVOC / PAH             → SOIL_VOC
  (groundwater codes are handled by groundwater/kte.py)
"""

from __future__ import annotations

import io
import logging
import re
import xml.etree.ElementTree as ET

import pandas as pd

from parsers.base import BaseParser
from core.lab_value_parser import LabValueParser
from core.cas_lookup import name_to_cas


# Analysis code → analysis_type string
SOIL_ANALYSIS_MAP: dict[str, str] = {
    # VOC / BTEX
    "BTEX_MTBE_SOIL_WS_WT": "SOIL_VOC",
    "BTEX_MTBE_SOIL":       "SOIL_VOC",
    "VOC_SOIL":             "SOIL_VOC",
    # TPH
    "TPH_DRO_ORO":          "SOIL_TPH",
    "TPH":                  "SOIL_TPH",
    # Metals
    "METALS":               "SOIL_METALS",
    "ICP":                  "SOIL_METALS",
    "METALS_SOIL":          "SOIL_METALS",
    "ICP_SOIL":             "SOIL_METALS",
    "ICP_OES":              "SOIL_METALS",
    "ICP_OES_SOIL":         "SOIL_METALS",
    # SVOCs / PAHs — VSL + Tier-1 soil thresholds apply (same as SOIL_VOC)
    "SVOC":                 "SOIL_VOC",
    "SVOC_SOIL":            "SOIL_VOC",
    "PAH":                  "SOIL_VOC",
    "PAH_SOIL":             "SOIL_VOC",
}

_log = logging.getLogger(__name__)

# Compound name → CAS (KTE uses English names)
KTE_CAS: dict[str, str] = {
    # VOC / BTEX
    "benzene":        "71-43-2",
    "toluene":        "108-88-3",
    "ethyl benzene":  "100-41-4",
    "ethylbenzene":   "100-41-4",
    "xylene":         "1330-20-7",
    "xylenes":        "1330-20-7",
    "mtbe":           "1634-04-4",
    "methyl tert-butyl ether": "1634-04-4",
    "naphthalene":    "91-20-3",
    "1,2,3-trimethylbenzene": "526-73-8",
    "1,2,4-trimethylbenzene": "95-63-6",
    "1,3,5-trimethylbenzene": "108-67-8",
    "styrene":        "100-42-5",
    # TPH fractions
    "total dro":      "C10-C40",
    "total oro":      "C10-C40",
    "total dro+oro":  "C10-C40",
    "dro":            "C10-C40",
    "oro":            "C10-C40",
    # Metals (English names from KTE LIMS)
    "lead":           "7439-92-1",   # Pb
    "zinc":           "7440-66-6",   # Zn
    "copper":         "7440-50-8",   # Cu
    "arsenic":        "7440-38-2",   # As
    "cadmium":        "7440-43-9",   # Cd
    "chromium":       "7440-47-3",   # Cr (total)
    "nickel":         "7440-02-0",   # Ni
    "mercury":        "7439-97-6",   # Hg
    "iron":           "7439-89-6",   # Fe
    "manganese":      "7439-96-5",   # Mn
    "barium":         "7440-39-3",   # Ba
    "vanadium":       "7440-62-2",   # V
    "cobalt":         "7440-48-4",   # Co
    "selenium":       "7782-49-2",   # Se
    "antimony":       "7440-36-0",   # Sb
    "silver":         "7440-22-4",   # Ag
    "aluminium":      "7429-90-5",   # Al
    "aluminum":       "7429-90-5",   # Al (US spelling)
    "boron":          "7440-42-8",   # B
    "molybdenum":     "7439-98-7",   # Mo
    "thallium":       "7440-28-0",   # Tl
    # Common PAHs (SVOC)
    "acenaphthylene":          "208-96-8",
    "acenaphthene":            "83-32-9",
    "fluorene":                "86-73-7",
    "phenanthrene":            "85-01-8",
    "anthracene":              "120-12-7",
    "fluoranthene":            "206-44-0",
    "pyrene":                  "129-00-0",
    "benzo[a]anthracene":      "56-55-3",
    "chrysene":                "218-01-9",
    "benzo[b]fluoranthene":    "205-99-2",
    "benzo[k]fluoranthene":    "207-08-9",
    "benzo[a]pyrene":          "50-32-8",
    "indeno[1,2,3-cd]pyrene":  "193-39-5",
    "dibenz[a,h]anthracene":   "53-70-3",
    "benzo[ghi]perylene":      "191-24-2",
}


def _resolve_cas(compound: str) -> str:
    key = compound.strip().lower()
    if key in KTE_CAS:
        return KTE_CAS[key]
    return name_to_cas(compound) or ""


class KTESoilParser(BaseParser):
    LAB_NAME = "KTE"
    ANALYSIS_TYPES = ["SOIL_VOC", "SOIL_TPH", "SOIL_METALS"]

    # Column indices (0-based)
    C_ORDER   = 0
    C_SAMPLE  = 1
    C_ACODE   = 2
    C_CPND    = 4
    C_RESULT  = 5
    C_UNIT    = 6
    C_DATE    = 8
    C_REPORT  = 11
    C_SITE    = 12
    C_LOC     = 13   # borehole / location ID  e.g.  ק-1, ק-3
    _REQUIRED_NCOLS = C_LOC + 1

    def __init__(self):
        self._vp = LabValueParser()

    # ------------------------------------------------------------------
    def parse(self, file_obj: io.BytesIO | str) -> list[dict]:
        df = self._read(file_obj)
        if df is None or df.empty:
            return []

        # EXCEL_GENERIC wide-format detection (pivot layout: parameters as rows, samples as columns)
        if self._is_generic_wide(df):
            return self._parse_generic_wide(df)

        # Ensure positional indexing is always safe (some exports have missing trailing columns).
        if df.shape[1] < self._REQUIRED_NCOLS:
            df = df.reindex(columns=range(self._REQUIRED_NCOLS), fill_value="")

        records: list[dict] = []
        for _, row in df.iterrows():
            # Skip rows that are too short / empty (defensive for malformed rows)
            if len(row) < self._REQUIRED_NCOLS:
                continue

            acode    = str(row.iloc[self.C_ACODE]).strip().upper()
            atype    = self._resolve_atype(acode)
            if atype is None:
                _log.warning("KTESoilParser: unknown analysis code %r — row skipped", acode)
                continue

            compound = str(row.iloc[self.C_CPND]).strip()
            raw_val  = str(row.iloc[self.C_RESULT]).strip()
            unit     = str(row.iloc[self.C_UNIT]).strip()
            loc      = str(row.iloc[self.C_LOC]).strip()
            date_val = str(row.iloc[self.C_DATE]).strip()

            if not compound or compound.lower() in ("nan", ""):
                continue

            # Parse value
            if raw_val.lower() in ("not detected", "nd", "n.d.", "n/d", "<dl", ""):
                value, flag = None, "ND"
            else:
                value, flag = self._vp.parse(raw_val)

            cas = _resolve_cas(compound)

            # Build sample_id: location + date (date as day only to keep short)
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
                "unit":          unit if unit else ("mg/kg" if atype == "SOIL_VOC" else "mg/kg"),
                "lod":           None,
                "analysis_type": atype,
            })

        return records

    # ------------------------------------------------------------------
    @staticmethod
    def _is_generic_wide(df: pd.DataFrame) -> bool:
        """
        Detect KTE EXCEL_GENERIC wide/pivot format.
        _read() strips the first 2 rows (CLIENT + blank), so after stripping
        row 0 is 'Work Order:'.  We also accept the un-stripped case.
        """
        try:
            cell00 = str(df.iloc[0, 0]).strip()
            if cell00.upper() == "CLIENT" or "Work Order:" in cell00:
                return True
        except (IndexError, TypeError):
            pass
        return False

    def _parse_generic_wide(self, df: pd.DataFrame) -> list[dict]:
        """
        Parse KTE EXCEL_GENERIC wide-format (pivot layout).

        Structure:
          row 6 col 4+  — Client Sample IDs  (N-5 (0.6), S-5 (0.6), …)
          row 8 col 4+  — Sampling dates     (DD/MM/YYYY)
          row 13, 23, … — Section headers    (BTEX / Total Petroleum Hydrocarbons / …)
          row 14+        — Data rows: [Compound, Method, Unit, LOR, val1, val2, …]
        """
        SECTION_MAP = [
            ("btex",              "SOIL_VOC"),
            ("non-halogenated",   "SOIL_VOC"),
            ("volatile organic",  "SOIL_VOC"),
            ("total petroleum",   "SOIL_TPH"),
            ("tph",               "SOIL_TPH"),
        ]
        # Summary / physical rows to skip
        SKIP_RE = re.compile(
            r"^(sum of|dry matter|n-dekan|physical param)", re.IGNORECASE
        )

        # ── Step 1: locate sample IDs and dates ─────────────────────
        sample_ids: list[str] = []
        sample_dates: list[str] = []

        for _, row in df.iterrows():
            vals = [str(v).strip() for v in row.values]
            if "Client Sample ID" in vals:
                idx = vals.index("Client Sample ID")
                sample_ids = [v for v in vals[idx + 1:] if v]
            if "Client Sampling Date" in vals:
                idx = vals.index("Client Sampling Date")
                sample_dates = vals[idx + 1: idx + 1 + len(sample_ids)]

        if not sample_ids:
            _log.warning("KTESoilParser (generic): no sample IDs found")
            return []

        n = len(sample_ids)

        # ── Step 2: iterate data rows ───────────────────────────────
        current_atype: str | None = None
        records: list[dict] = []

        for _, row in df.iterrows():
            vals = [str(v).strip() for v in row.values]
            compound = vals[0] if vals else ""
            method   = vals[1] if len(vals) > 1 else ""
            unit_raw = vals[2] if len(vals) > 2 else ""
            lor_str  = vals[3] if len(vals) > 3 else ""
            data_vals = vals[4: 4 + n]

            # Section header?
            clow = compound.lower()
            for key, atype in SECTION_MAP:
                if key in clow:
                    current_atype = atype
                    break

            # Skip if not a proper data row
            if not method or method in ("Method", "Unit", ""):
                continue
            if SKIP_RE.match(compound) or not compound:
                continue
            if current_atype is None:
                continue

            # Unit normalization: µg/kg → mg/kg (/1000)
            unit_lower = unit_raw.lower().replace("\xb5", "µ").replace("\u03bc", "µ")
            if "µg" in unit_lower or (unit_raw.startswith("?") and "g/kg" in unit_raw):
                scale = 1000.0
                unit  = "mg/kg"
            else:
                scale = 1.0
                unit  = unit_raw.replace(" DW", "").replace(" dw", "").strip() or "mg/kg"

            # LOR
            try:
                lor = float(lor_str) / scale if lor_str else None
            except ValueError:
                lor = None

            cas = _resolve_cas(compound)

            for sid, sdate, raw_val in zip(sample_ids, sample_dates, data_vals):
                # Keep "Name (depth)" format so _split_sample_depth in excel_output
                # can extract the depth for the depth column (e.g. "N-5 (0.6)" → N-5, 0.6).
                # Adding a date suffix breaks the depth regex, so we omit it here.
                sample_id = sid

                # Value
                if raw_val.lower() in ("not detected", "nd", "n.d.", "n/d", "<dl", "<lor", ""):
                    value, flag, lod = None, "ND", lor
                else:
                    v, f = self._vp.parse(raw_val)
                    if f in ("<", "<LOD", "<LOQ") or (v is None and raw_val.startswith("<")):
                        m = re.match(r"<\s*([0-9.]+)", raw_val)
                        lod_val = float(m.group(1)) / scale if m else lor
                        value, flag, lod = None, "<LOD", lod_val
                    else:
                        value = (v / scale) if v is not None else None
                        flag, lod = f or "", lor

                records.append({
                    "lab":           self.LAB_NAME,
                    "sample_id":     sample_id,
                    "compound":      compound,
                    "cas":           cas,
                    "value":         value,
                    "flag":          flag,
                    "unit":          unit,
                    "lod":           lod,
                    "analysis_type": current_atype,
                })

        return records

    @staticmethod
    def _generic_date(date_str: str) -> str:
        """Convert DD/MM/YYYY to DD.MM.YYYY."""
        m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", date_str)
        if m:
            return f"{int(m.group(1)):02d}.{int(m.group(2)):02d}.{m.group(3)}"
        return ""

    # ------------------------------------------------------------------
    def _read(self, file_obj: io.BytesIO | str) -> pd.DataFrame | None:
        try:
            if isinstance(file_obj, str) and file_obj.lower().endswith(".csv"):
                df = pd.read_csv(
                    file_obj, encoding="utf-8-sig", dtype=str,
                    header=None, engine="python",
                    on_bad_lines="skip", quotechar='"',
                ).fillna("")
                if df.shape[0] > 2:
                    first = str(df.iloc[0, 0]).strip()
                    if not first.replace("-", "").isdigit():
                        df = df.iloc[2:].reset_index(drop=True)
            else:
                df = None

                raw_head: bytes = b""
                if isinstance(file_obj, io.BytesIO):
                    raw_head = file_obj.getbuffer()[:256].tobytes()
                elif isinstance(file_obj, str):
                    try:
                        with open(file_obj, "rb") as f:
                            raw_head = f.read(256)
                    except OSError:
                        raw_head = b""

                # KTE sometimes exports "EXCEL_GENERIC.XLS" as SpreadsheetML (Excel 2003 XML),
                # not a real BIFF8 .xls file.
                if raw_head.lstrip().startswith(b"<?xml") and b"urn:schemas-microsoft-com:office:spreadsheet" in raw_head:
                    if isinstance(file_obj, io.BytesIO):
                        raw = file_obj.getvalue()
                    else:
                        with open(file_obj, "rb") as f:
                            raw = f.read()
                    df = self._read_spreadsheetml(raw).fillna("")
                else:
                    engine = None
                    if raw_head.startswith(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"):
                        engine = "xlrd"      # legacy XLS (requires xlrd)
                    elif raw_head.startswith(b"PK"):
                        engine = "openpyxl"  # XLSX

                    xl = pd.ExcelFile(file_obj, engine=engine)
                    df = xl.parse(xl.sheet_names[0], header=None, dtype=str).fillna("")

                if df.shape[0] > 2:
                    first_cell = str(df.iloc[0, 0]).strip()
                    if not first_cell.replace("-", "").isdigit():
                        df = df.iloc[2:].reset_index(drop=True)
        except Exception as e:
            raise ValueError(f"KTESoilParser: cannot read file — {e}") from e
        return df

    @staticmethod
    def _read_spreadsheetml(raw: bytes) -> pd.DataFrame:
        """
        Read Excel 2003 XML / SpreadsheetML exported as .XLS.
        Extracts the first worksheet table into a DataFrame.
        """
        ns = {
            "ss": "urn:schemas-microsoft-com:office:spreadsheet",
        }
        root = ET.fromstring(raw)
        table = root.find(".//ss:Worksheet/ss:Table", ns)
        if table is None:
            return pd.DataFrame()

        rows_out: list[list[str]] = []
        for row in table.findall("ss:Row", ns):
            out_row: list[str] = []
            col_pos = 1  # SpreadsheetML column positions are 1-based for ss:Index
            for cell in row.findall("ss:Cell", ns):
                idx_attr = cell.get(f"{{{ns['ss']}}}Index")
                if idx_attr:
                    try:
                        idx = int(idx_attr)
                        while col_pos < idx:
                            out_row.append("")
                            col_pos += 1
                    except ValueError:
                        pass

                data = cell.find("ss:Data", ns)
                text = ""
                if data is not None and data.text is not None:
                    text = str(data.text)
                out_row.append(text)
                col_pos += 1
            rows_out.append(out_row)

        if not rows_out:
            return pd.DataFrame()

        max_cols = max(len(r) for r in rows_out)
        norm = [r + [""] * (max_cols - len(r)) for r in rows_out]
        return pd.DataFrame(norm)

    @staticmethod
    def _resolve_atype(acode: str) -> str | None:
        for key, atype in SOIL_ANALYSIS_MAP.items():
            if key.upper() in acode or acode in key.upper():
                return atype
        return None

    @staticmethod
    def _short_date(date_str: str) -> str:
        """Extract DD.MM.YYYY from datetime string like '2026-01-19 00:00:00'."""
        m = re.search(r'(\d{4})-(\d{2})-(\d{2})', date_str)
        if m:
            return f"{m.group(3)}.{m.group(2)}.{m.group(1)}"
        return ""
