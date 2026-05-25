from __future__ import annotations
import io
import re
import openpyxl
from parsers.base import BaseParser

class MachonHaneftParser(BaseParser):
    LAB_NAME = "מכון הנפט"

    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        wb = openpyxl.load_workbook(file_obj, data_only=True)
        records = []
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            all_rows = list(ws.iter_rows(values_only=True))
            if sheet_name.upper() in ("SVOC", "VOC", "TPH"):
                analysis_type = f"SOIL_{sheet_name.upper()}"
            else:
                analysis_type = "SOIL_SVOC"
            sample_row_idx = None
            for i, row in enumerate(all_rows):
                for cell in row:
                    if cell and isinstance(cell, str) and re.match(r'[SK]\d+', str(cell)):
                        sample_row_idx = i
                        break
                if sample_row_idx is not None:
                    break
            if sample_row_idx is None:
                continue
            sample_row = all_rows[sample_row_idx]
            sample_cols = {}
            for j, cell in enumerate(sample_row):
                if cell and isinstance(cell, str) and re.match(r'[SK]\d+', str(cell)):
                    sample_cols[j] = str(cell).strip()
            for row in all_rows[sample_row_idx + 2:]:
                if row[0] is None:
                    continue
                try:
                    int(row[0])
                except (TypeError, ValueError):
                    continue
                compound = str(row[2]).strip() if row[2] else ""
                cas = str(row[1]).strip() if row[1] else ""
                unit = str(row[3]).strip() if row[3] else "mg/Kg"
                lod = row[4] if row[4] is not None else None
                loq = row[5] if row[5] is not None else None
                for col_idx, sample_id in sample_cols.items():
                    raw = row[col_idx]
                    if str(raw).strip().upper() == "ND":
                        value, flag = None, "ND"
                    else:
                        try:
                            value, flag = float(raw), ""
                        except (TypeError, ValueError):
                            value, flag = None, ""
                    records.append({
                        "lab": self.LAB_NAME,
                        "sample_id": sample_id,
                        "compound": compound,
                        "cas": cas,
                        "value": value,
                        "flag": flag,
                        "unit": unit,
                        "lod": lod,
                        "loq": loq,
                        "analysis_type": analysis_type,
                    })
        return records
