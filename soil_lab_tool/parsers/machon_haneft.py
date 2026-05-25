"""
machon_haneft.py
----------------
Parser for מכון הנפט (Machon HaNeft) laboratory Excel reports.

Wide-format layout: compounds as rows, samples as columns.
  - Sample-names row: first row where column index 6+ holds a value
    matching S\\d+ / K\\d+ or the borehole-depth form S1-11.0.
  - Header row: the row immediately after the sample-names row.
    Fixed columns (0-based): 0=row#, 1=CAS, 2=compound, 3=units, 4=LOD, 5=LOQ.
  - Data rows: col 0 is numeric, col 1 contains a hyphen (CAS pattern).
"""

from __future__ import annotations

import io
import re

from parsers.base import BaseParser
from core.lab_value_parser import LabValueParser


_SAMPLE_ID_RE = re.compile(r'^[SK]\d+')


class MachonHaneftParser(BaseParser):
    LAB_NAME = "מכון הנפט"

    def __init__(self):
        self._vp = LabValueParser()

    # ------------------------------------------------------------------
    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        import openpyxl

        wb = openpyxl.load_workbook(file_obj, read_only=True, data_only=True)
        ws = wb.active

        all_rows: list[list[str]] = []
        for row in ws.iter_rows(values_only=True):
            all_rows.append([
                str(v).strip() if v is not None else ""
                for v in row
            ])

        # Find sample-names row: first row where col index 6+ matches a sample ID
        sample_row_idx = None
        for i, row in enumerate(all_rows):
            if any(_SAMPLE_ID_RE.match(v) for v in row[6:] if v and v.lower() != "none"):
                sample_row_idx = i
                break

        if sample_row_idx is None:
            return []

        header_row_idx = sample_row_idx + 1
        if header_row_idx >= len(all_rows):
            return []

        sample_row = all_rows[sample_row_idx]

        # Collect sample column indices and IDs
        sample_col_indices: list[int] = []
        sample_ids: list[str] = []
        for ci in range(6, len(sample_row)):
            v = sample_row[ci]
            if v and v.lower() not in ("none", "nan", "") and _SAMPLE_ID_RE.match(v):
                sample_col_indices.append(ci)
                sample_ids.append(v)

        if not sample_col_indices:
            return []

        records: list[dict] = []
        for row in all_rows[header_row_idx + 1:]:
            if len(row) < 2:
                continue

            col0 = row[0]
            col1 = row[1]

            # Validate data row: col 0 numeric, col 1 contains a hyphen (CAS)
            try:
                float(col0)
            except (ValueError, TypeError):
                continue

            if not col1 or "-" not in col1:
                continue

            compound = row[2] if len(row) > 2 else ""
            unit = row[3] if len(row) > 3 else ""
            if not unit or unit.lower() in ("none", "nan", ""):
                unit = "mg/kg"

            lod = None
            if len(row) > 4 and row[4]:
                try:
                    lod = float(row[4].replace(",", ""))
                except (ValueError, TypeError):
                    pass

            loq = None
            if len(row) > 5 and row[5]:
                try:
                    loq = float(row[5].replace(",", ""))
                except (ValueError, TypeError):
                    pass

            for ci, sid in zip(sample_col_indices, sample_ids):
                if ci >= len(row):
                    continue
                raw_val = row[ci]
                if not raw_val or raw_val.lower() in ("none", "nan", ""):
                    continue

                if raw_val.lower() in ("nd", "n.d.", "not detected", "<dl"):
                    value, flag = None, "ND"
                elif raw_val.startswith("<"):
                    value, flag = self._vp.parse(raw_val)
                else:
                    try:
                        value = float(raw_val.replace(",", ""))
                        flag = ""
                    except ValueError:
                        value, flag = self._vp.parse(raw_val)

                records.append({
                    "lab":       self.LAB_NAME,
                    "sample_id": sid,
                    "compound":  compound,
                    "cas":       col1,
                    "value":     value,
                    "flag":      flag,
                    "unit":      unit,
                    "lod":       lod,
                    "loq":       loq,
                })

        return records
