"""
machon_haneft.py
----------------
Parser for מכון הנפט (Machon HaNeft) laboratory Excel reports.

Wide-format layout: compounds as rows, samples as columns.
  Fixed columns (0-based): 0=row#, 1=CAS, 2=compound, 3=units, 4=LOD, 5=LOQ, 6+=samples.
  Sample-names row: first row where any of cols 6-8 (0-based) matches S\\d+ or K\\d+.
  Header row: the row immediately after the sample-names row.
  Data rows: col 0 (1-indexed col 1) is an integer row number,
             col 1 (1-indexed col 2) is a CAS number string containing '-'.
  Analysis type is derived from the sheet name (SVOC → SOIL_SVOC, VOC → SOIL_VOC,
  TPH → SOIL_TPH).
"""

from __future__ import annotations

import io
import re

from parsers.base import BaseParser
from core.lab_value_parser import LabValueParser


_SAMPLE_ID_RE = re.compile(r'^[SK]\d+')


def _analysis_type_from_sheet(sheet_name: str) -> str:
    n = sheet_name.strip().upper()
    if "SVOC" in n:
        return "SOIL_SVOC"
    if "VOC" in n:
        return "SOIL_VOC"
    if "TPH" in n:
        return "SOIL_TPH"
    return "SOIL_SVOC"


class MachonHaneftParser(BaseParser):
    LAB_NAME = "מכון הנפט"

    def __init__(self):
        self._vp = LabValueParser()

    # ------------------------------------------------------------------
    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        import openpyxl

        wb = openpyxl.load_workbook(file_obj, read_only=True, data_only=True)
        records: list[dict] = []

        for ws in wb.worksheets:
            analysis_type = _analysis_type_from_sheet(ws.title or "")

            # Read all rows, keeping raw cell values (None stays None)
            all_rows: list[list] = [list(row) for row in ws.iter_rows(values_only=True)]

            # Find sample-names row: first row where any of cols 6-8 (0-based)
            # contains a value matching S\d+ or K\d+
            sample_row_idx = None
            for i, row in enumerate(all_rows):
                for ci in range(6, min(9, len(row))):
                    v = row[ci]
                    if v is not None and _SAMPLE_ID_RE.match(str(v).strip()):
                        sample_row_idx = i
                        break
                if sample_row_idx is not None:
                    break

            if sample_row_idx is None:
                continue

            header_row_idx = sample_row_idx + 1
            if header_row_idx >= len(all_rows):
                continue

            sample_row = all_rows[sample_row_idx]

            # Collect sample columns: start at index 6, stop at the first None cell
            sample_col_indices: list[int] = []
            sample_ids: list[str] = []
            for ci in range(6, len(sample_row)):
                v = sample_row[ci]
                if v is None:
                    break
                sid = str(v).strip()
                if not sid or sid.lower() in ("none", "nan", ""):
                    break
                sample_col_indices.append(ci)
                sample_ids.append(sid)

            if not sample_col_indices:
                continue

            # Parse data rows (skip the header row itself)
            for row in all_rows[header_row_idx + 1:]:
                if len(row) < 2:
                    continue

                # col 0 (1-indexed col 1): integer row number
                # col 1 (1-indexed col 2): CAS number
                col0 = row[0]
                col1 = row[1]

                if col0 is None or col1 is None:
                    continue

                try:
                    int(float(str(col0)))
                except (ValueError, TypeError):
                    continue

                cas = str(col1).strip()
                if "-" not in cas:
                    continue

                compound = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
                unit = str(row[3]).strip() if len(row) > 3 and row[3] is not None else ""
                if not unit or unit.lower() in ("none", "nan", ""):
                    unit = "mg/kg"

                lod = None
                if len(row) > 4 and row[4] is not None:
                    try:
                        lod = float(str(row[4]).replace(",", ""))
                    except (ValueError, TypeError):
                        pass

                loq = None
                if len(row) > 5 and row[5] is not None:
                    try:
                        loq = float(str(row[5]).replace(",", ""))
                    except (ValueError, TypeError):
                        pass

                for ci, sid in zip(sample_col_indices, sample_ids):
                    if ci >= len(row):
                        continue
                    cell_val = row[ci]
                    if cell_val is None:
                        continue
                    raw = str(cell_val).strip()
                    if not raw or raw.lower() in ("none", "nan", ""):
                        continue

                    if raw.upper() == "ND":
                        value, flag = None, "ND"
                    elif raw.startswith("<"):
                        value, flag = self._vp.parse(raw)
                    else:
                        try:
                            value = float(str(cell_val))
                            flag = ""
                        except (ValueError, TypeError):
                            value, flag = self._vp.parse(raw)

                    records.append({
                        "lab":           self.LAB_NAME,
                        "sample_id":     sid,
                        "compound":      compound,
                        "cas":           cas,
                        "value":         value,
                        "flag":          flag,
                        "unit":          unit,
                        "lod":           lod,
                        "loq":           loq,
                        "analysis_type": analysis_type,
                    })

        return records
