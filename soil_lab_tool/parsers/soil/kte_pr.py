"""
parsers/soil/kte_pr.py
-----------------------
Parser for KTE / Dr. Katz laboratory reports in EXCEL_GENERIC (SpreadsheetML XML) format.

These files have a .XLS extension but are actually XML in SpreadsheetML format
(ISO-8859-1 encoded).  xlrd cannot read them; they must be parsed as XML.

Sheet used: 'Client SOIL - 1'

Typical layout:
  Row 0:  Work Order: <id>
  Row 1:  Client: <name>
  Row 2:  Project: <id>
  Row 3-5: Metadata
  Row 6-7: Method headers
  Row 8:  "Client Sample ID" |  | S85 (0.5) | S85 (1.0) | ...
  Row 9+: compound rows: symbol | CAS | LOR | ... | value_per_sample...

Analysis type: SOIL_METALS (typically contains SVOCs / metals / PAHs)
"""

from __future__ import annotations

import io
import re
import xml.etree.ElementTree as ET

from parsers.base import BaseParser
from core.lab_value_parser import LabValueParser


# SpreadsheetML namespace
_SS_NS = "urn:schemas-microsoft-com:office:spreadsheet"


def _strip_ns(root):
    """Return a new tree with all namespace prefixes stripped for easier searching."""
    xml_str = ET.tostring(root, encoding="unicode")
    xml_str = re.sub(r'\s+xmlns(?::\w+)?="[^"]*"', '', xml_str)
    xml_str = re.sub(r'<(\w+):', r'<', xml_str)
    xml_str = re.sub(r'</(\w+):', r'</', xml_str)
    xml_str = re.sub(r'(\s)(\w+):', r'\1', xml_str)
    return ET.fromstring(xml_str)


class KTEPRParser(BaseParser):
    LAB_NAME = "KTE"
    ANALYSIS_TYPES = ["SOIL_METALS"]

    TARGET_SHEET = "client soil - 1"

    def __init__(self):
        self._vp = LabValueParser()

    # ------------------------------------------------------------------
    def parse(self, file_obj: io.BytesIO | str) -> list[dict]:
        try:
            if isinstance(file_obj, str):
                with open(file_obj, "rb") as f:
                    raw_bytes = f.read()
            else:
                raw_bytes = file_obj.read()

            # Try UTF-8 first, fall back to ISO-8859-1
            for enc in ("utf-8", "iso-8859-1", "windows-1252"):
                try:
                    content = raw_bytes.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
            else:
                content = raw_bytes.decode("iso-8859-1", errors="replace")

            root = ET.fromstring(content)
        except ET.ParseError:
            return []

        root = _strip_ns(root)

        # Find the target worksheet
        target_ws = None
        for ws in root.findall(".//Worksheet"):
            name = ws.get("Name", "")
            if name.strip().lower() == self.TARGET_SHEET:
                target_ws = ws
                break
        if target_ws is None and root.findall(".//Worksheet"):
            target_ws = root.findall(".//Worksheet")[0]
        if target_ws is None:
            return []

        # Extract rows → list of lists
        table = target_ws.find(".//Table")
        if table is None:
            return []

        rows: list[list[str]] = []
        for row_el in table.findall("Row"):
            cells: list[str] = []
            prev_index = 0
            for cell_el in row_el.findall("Cell"):
                # Handle ss:Index for sparse rows
                idx_attr = cell_el.get("Index")
                if idx_attr is not None:
                    gap = int(idx_attr) - 1 - prev_index
                    cells.extend([""] * gap)
                data_el = cell_el.find("Data")
                val = data_el.text if (data_el is not None and data_el.text) else ""
                cells.append(str(val).strip())
                prev_index = len(cells) - 1
            rows.append(cells)

        return self._parse_rows(rows)

    # ------------------------------------------------------------------
    def _parse_rows(self, rows: list[list[str]]) -> list[dict]:
        # Find the "Client Sample ID" row → sample IDs
        sample_row_idx = None
        for i, row in enumerate(rows):
            row_flat = " ".join(row).lower()
            if "client sample id" in row_flat or "sample id" in row_flat:
                sample_row_idx = i
                break

        if sample_row_idx is None:
            return []

        sample_row = rows[sample_row_idx]

        # Find fixed columns: compound name, CAS, LOR
        # Scan first few rows for a row that contains "CAS"
        header_row_idx = None
        for i in range(sample_row_idx, min(sample_row_idx + 5, len(rows))):
            row_flat = " ".join(rows[i]).lower()
            if "cas" in row_flat:
                header_row_idx = i
                break

        if header_row_idx is None:
            header_row_idx = sample_row_idx + 1

        header_row = rows[header_row_idx]

        def find_col(keywords):
            for k in keywords:
                for i, h in enumerate(header_row):
                    if k.lower() in h.lower():
                        return i
            return None

        col_cmp = find_col(["analyte", "compound", "parameter", "name"]) or 0
        col_cas = find_col(["cas"])
        col_lor = find_col(["lor", "lod", "mdl", "limit"])

        # Sample columns: find from sample_row (non-empty values after fixed cols)
        fixed_max = max(c for c in [col_cmp, col_cas, col_lor] if c is not None) + 1
        sample_col_ids: list[tuple[int, str]] = []
        for c in range(fixed_max, len(sample_row)):
            sid = sample_row[c].strip()
            if sid and sid.lower() not in ("nan", "", "client sample id"):
                sample_col_ids.append((c, sid))

        if not sample_col_ids:
            return []

        records: list[dict] = []
        for row in rows[header_row_idx + 1:]:
            if not row:
                continue
            compound = row[col_cmp].strip() if col_cmp < len(row) else ""
            if not compound or compound.lower() in ("nan", ""):
                continue
            # Skip QC rows (MB, DUP, LCS)
            if any(kw in compound.lower() for kw in ("blank", "duplicate", "spike", "mb")):
                continue

            cas = row[col_cas].strip() if (col_cas and col_cas < len(row)) else ""
            if cas.lower() in ("nan", ""):
                cas = ""

            lor = None
            if col_lor and col_lor < len(row):
                try:
                    lor = float(row[col_lor].replace(",", ""))
                except (ValueError, TypeError):
                    pass

            for col_idx, sid in sample_col_ids:
                raw_val = row[col_idx].strip() if col_idx < len(row) else ""
                if raw_val.lower() in ("nan", ""):
                    continue

                if raw_val.startswith("<") or raw_val.lower() in ("nd", "not detected",
                                                                    "bdl", "blq"):
                    value = lor
                    flag = "ND"
                else:
                    try:
                        value = float(raw_val.replace(",", ""))
                        flag = ""
                    except ValueError:
                        value, flag = self._vp.parse(raw_val)

                records.append({
                    "lab":           self.LAB_NAME,
                    "sample_id":     sid,
                    "compound":      compound,
                    "cas":           cas,
                    "value":         value,
                    "flag":          flag,
                    "unit":          "mg/kg DW",
                    "lod":           lor,
                    "analysis_type": "SOIL_METALS",
                })

        return records
