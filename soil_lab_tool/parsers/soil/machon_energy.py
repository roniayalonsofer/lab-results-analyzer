"""
parsers/soil/machon_energy.py
------------------------------
Parser for המכון הישראלי לאנרגיה ולסביבה Excel reports.

Expected format:
  Sheets named 'SVOC' or 'VOC'.
  Rows 0–18 (0-indexed): metadata — project name, date, etc.
  Row 19 (0-indexed): header row:
    Cas.No. | <blank> | Compound | יחידות | גבול גילוי | גבול כימות | תוצאה
  Row 20+ (0-indexed): data rows:
    מספר | CAS | compound_name | יחידות | detection_limit | LOQ | result

  Sample ID: row where col 0 contains 'פרויקט:' — value after the colon.
  Date:      row where col 0 contains 'תאריך לקיחת המדגם'.
  Result:    'ND' (not detected) or a numeric value.

PDF input raises a clean ValueError — upload Excel (.xlsx) instead.
"""

from __future__ import annotations

import io
import os

import pandas as pd

from parsers.base import BaseParser
from core.lab_value_parser import LabValueParser

_HEADER_ROW = 19   # 0-indexed
_DATA_START  = 20  # 0-indexed
_META_SCAN   = 35  # scan at most this many rows for metadata

_PDF_MAGIC = b"%PDF"


class MachonEnergyParser(BaseParser):
    LAB_NAME       = "המכון הישראלי לאנרגיה ולסביבה"
    ANALYSIS_TYPES = ["SOIL_GAS_VOC", "SOIL_VOC"]

    def __init__(self):
        self._vp = LabValueParser()

    # ------------------------------------------------------------------
    def parse(self, file_obj: io.BytesIO, filename: str = "", **_kw) -> list[dict]:
        # Reject PDFs early with a clean message
        file_obj.seek(0)
        magic = file_obj.read(4)
        file_obj.seek(0)
        if magic == _PDF_MAGIC:
            raise ValueError(
                "PDF format not supported for this lab — please upload an Excel file (.xlsx)"
            )

        xl   = pd.ExcelFile(file_obj)
        smap = {s.strip().upper(): s for s in xl.sheet_names}

        target_sheets: list[tuple[str, str]] = []
        for key in ("VOC", "SVOC"):
            if key in smap:
                target_sheets.append((key, smap[key]))

        if not target_sheets:
            target_sheets = [("VOC", xl.sheet_names[0])]

        records: list[dict] = []
        for sheet_key, sheet_name in target_sheets:
            records.extend(self._parse_sheet(xl, sheet_name, sheet_key, filename))
        return records

    # ------------------------------------------------------------------
    def _parse_sheet(self, xl: pd.ExcelFile, sheet_name: str,
                     sheet_key: str, filename: str) -> list[dict]:
        raw = xl.parse(sheet_name, header=None, dtype=str).fillna("")

        sample_id     = self._extract_sample_id(raw, filename)
        sampling_date = self._extract_date(raw)
        header_row_idx = self._find_header_row(raw)

        # Column positions — fixed layout; refined from header text if possible
        col_cas      = 1
        col_compound = 2
        col_unit     = 3
        col_lod      = 4
        col_loq      = 5
        col_result   = 6

        if header_row_idx < len(raw):
            hdrs = [str(v).strip() for v in raw.iloc[header_row_idx].values]
            for ci, h in enumerate(hdrs):
                hl = h.lower()
                if "cas" in hl and col_cas == 1:
                    col_cas = ci
                elif ("compound" in hl or "שם" in hl) and col_compound == 2:
                    col_compound = ci
                elif ("יחידות" in hl or "unit" in hl) and col_unit == 3:
                    col_unit = ci
                elif ("גבול גילוי" in hl or "detection" in hl) and col_lod == 4:
                    col_lod = ci
                elif ("גבול כימות" in hl or "loq" in hl) and col_loq == 5:
                    col_loq = ci
                elif ("תוצאה" in hl or "result" in hl) and col_result == 6:
                    col_result = ci

        default_atype = "SOIL_VOC" if sheet_key == "SVOC" else "SOIL_GAS_VOC"

        records: list[dict] = []
        data_start = max(_DATA_START, header_row_idx + 1)

        for i in range(data_start, len(raw)):
            vals = list(raw.iloc[i].values)
            if not vals:
                continue

            # Col 0 must be a row serial number (integer stored as int or float)
            if not _is_integer_str(str(vals[0]).strip()):
                continue

            compound = str(vals[col_compound]).strip() if col_compound < len(vals) else ""
            if not compound or compound.lower() in ("", "nan"):
                continue

            cas = str(vals[col_cas]).strip() if col_cas < len(vals) else ""
            if cas.lower() in ("", "nan"):
                cas = ""

            unit = str(vals[col_unit]).strip() if col_unit < len(vals) else ""
            if unit.lower() in ("", "nan"):
                unit = ""

            analysis_type = _infer_analysis_type(unit, default_atype)

            lod = _parse_float(vals, col_lod)
            loq = _parse_float(vals, col_loq)

            raw_val = str(vals[col_result]).strip() if col_result < len(vals) else ""

            if raw_val.upper() in ("ND", "N.D.", "N/D", "NOT DETECTED", "", "NAN"):
                value, flag = lod, "ND"
            else:
                value, flag = self._vp.parse(raw_val)

            records.append({
                "lab":           self.LAB_NAME,
                "sample_id":     sample_id,
                "compound":      compound,
                "cas":           cas,
                "value":         value,
                "flag":          flag,
                "unit":          unit,
                "lod":           lod,
                "loq":           loq,
                "analysis_type": analysis_type,
                "sampling_date": sampling_date,
            })

        return records

    # ------------------------------------------------------------------
    def _extract_sample_id(self, df: pd.DataFrame, filename: str) -> str:
        project = ""
        for i in range(min(_META_SCAN, len(df))):
            row_flat = " ".join(str(v).strip() for v in df.iloc[i].values)
            if "פרויקט" in row_flat:
                # Find the cell that contains the keyword and extract the value
                for ci in range(len(df.iloc[i])):
                    cell = str(df.iloc[i, ci]).strip()
                    if "פרויקט" in cell:
                        if ":" in cell:
                            project = cell.split(":", 1)[1].strip()
                        if not project or project.lower() == "nan":
                            project = _first_nonempty(df.iloc[i], start=ci + 1)
                        break
                break

        fn_base = os.path.splitext(os.path.basename(filename))[0] if filename else ""
        if project and fn_base:
            return f"{project} — {fn_base}"
        return project or fn_base or "Sample"

    def _extract_date(self, df: pd.DataFrame) -> str:
        for i in range(min(_META_SCAN, len(df))):
            row_flat = " ".join(str(v).strip() for v in df.iloc[i].values)
            if "תאריך לקיחת המדגם" in row_flat:
                for ci in range(len(df.iloc[i])):
                    cell = str(df.iloc[i, ci]).strip()
                    if "תאריך לקיחת המדגם" in cell:
                        if ":" in cell:
                            part = cell.split(":", 1)[1].strip()
                            if part and part.lower() not in ("", "nan"):
                                return part
                        return _first_nonempty(df.iloc[i], start=ci + 1)
        return ""

    def _find_header_row(self, df: pd.DataFrame) -> int:
        # Try fixed position first
        if _HEADER_ROW < len(df):
            row_str = " ".join(str(v).lower() for v in df.iloc[_HEADER_ROW].values)
            if "compound" in row_str or "cas" in row_str or "גבול" in row_str:
                return _HEADER_ROW

        # Scan up to row 30 as fallback
        for i in range(min(30, len(df))):
            row_str = " ".join(str(v).lower() for v in df.iloc[i].values)
            if ("compound" in row_str or "גבול" in row_str) and "cas" in row_str:
                return i

        return _HEADER_ROW


# ------------------------------------------------------------------
# Module-level helpers
# ------------------------------------------------------------------

def _is_integer_str(s: str) -> bool:
    """Return True for strings that represent a whole number (int or float like '1.0')."""
    try:
        f = float(s)
        return f == int(f) and f >= 0
    except (ValueError, TypeError):
        return False


def _first_nonempty(row, start: int = 1) -> str:
    for ci in range(start, len(row)):
        val = str(row.iloc[ci]).strip()
        if val and val.lower() not in ("", "nan"):
            return val
    return ""


def _parse_float(vals: list, col: int) -> float | None:
    if col >= len(vals):
        return None
    try:
        return float(str(vals[col]).strip().replace(",", ""))
    except (ValueError, TypeError):
        return None


def _infer_analysis_type(unit: str, default: str) -> str:
    u = unit.lower()
    if "kg" in u:
        return "SOIL_VOC"
    if "/l" in u or "mg/l" in u:
        return "GW_VOC"
    return default


# ------------------------------------------------------------------
# Detection helpers (used by parsers/__init__.py auto-detection)
# ------------------------------------------------------------------

def is_machon_energy_excel(file_bytes: bytes) -> bool:
    """Return True if the Excel file is a מכון האנרגיה report.

    Checks for a sheet named exactly 'VOC' or 'SVOC' (case-insensitive)
    AND any cell in the first 35 rows containing 'פרויקט' or
    'תאריך לקיחת המדגם' (scans all columns, not just column 0).
    """
    try:
        import io as _io
        xl   = pd.ExcelFile(_io.BytesIO(file_bytes))
        smap = {s.strip().upper(): s for s in xl.sheet_names}

        target = next((smap[k] for k in ("VOC", "SVOC") if k in smap), None)
        if target is None:
            return False

        df = xl.parse(target, header=None, dtype=str, nrows=_META_SCAN).fillna("")
        for i in range(len(df)):
            row_flat = " ".join(str(v).strip() for v in df.iloc[i].values)
            if "פרויקט" in row_flat or "תאריך לקיחת המדגם" in row_flat:
                return True
    except Exception:
        pass
    return False


def is_machon_energy_pdf(file_bytes: bytes) -> bool:
    """Return True if the PDF text identifies it as a מכון האנרגיה report."""
    try:
        import io as _io
        import pdfplumber
        with pdfplumber.open(_io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages[:3]:
                text = page.extract_text() or ""
                if "המכון הישראלי לאנרגיה" in text or "אנרגיה ולסביבה" in text:
                    return True
    except Exception:
        pass
    return False
