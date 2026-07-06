"""
alchem_pdf.py
-------------
Parser for Alchem / TO-15 soil gas laboratory PDF reports (the "נספח לדוח
אנליזה" appendix format). This is the PDF twin of soil_gas/alchem.py, which
only handles the Excel version of the same report.

A single PDF can contain several "canister groups" (one Excel-style table
per group of samples that shared an analysis run), each introduced by its
own "Canister Number:" / "Analysis Time:" / "Analysis Location:" header
rows, e.g.:

  Canister Number:   | | 8569 | 8382 | 11703 |
  Analysis Time:     | | 19:18| 19:52| 20:28  |
  Analysis Location: | | SG-1 | SG-2 | SG-3   |
  Name | CAS | Final Conc. | Final Conc. | Final Conc. | LOD | LOQ
  <compound rows...>

  Canister Number:   | | 8611 |
  Analysis Time:     | | 21:03|
  Analysis Location: | | SG-4 |
  Name | CAS | Final Conc. | LOD | LOQ
  <compound rows...>

Continuation pages repeat only the compound rows (no header), so the
current group's sample list stays in effect until a new
"Canister Number:" row appears.
"""

from __future__ import annotations

import io
import re

from parsers.base import BaseParser
from core.lab_value_parser import LabValueParser

_CAS_RE = _re = re.compile(r'^\d{1,7}-\d{2}-\d$')


class AlchemSoilGasPDFParser(BaseParser):
    LAB_NAME = "Alchem"
    ANALYSIS_TYPES = ["SOIL_GAS_VOC"]

    def __init__(self):
        self._vp = LabValueParser()

    # ------------------------------------------------------------------
    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        import pdfplumber

        records: list[dict] = []
        sample_ids: list[str] = []
        canister_nums: list[str] = []
        analysis_times: list[str] = []

        with pdfplumber.open(file_obj) as pdf:
            for page in pdf.pages:
                for table in page.extract_tables():
                    for row in table:
                        cells = [self._clean(c) for c in row]
                        cells = [c for c in cells if c != ""]   # drop merged-cell padding
                        if not cells:
                            continue
                        first = cells[0]

                        if first.startswith("Canister Number"):
                            canister_nums = cells[1:]
                            continue
                        if first.startswith("Analysis Time"):
                            analysis_times = cells[1:]
                            continue
                        if first.startswith("Analysis Location"):
                            sample_ids = cells[1:]
                            continue
                        if first in ("Name", "Compound Name", "Analyte", "CAS"):
                            continue  # column-header row (or its wrapped continuation)
                        if "total voc" in first.lower():
                            continue  # summary row, not a compound
                        if len(cells) < 2:
                            continue

                        cas = cells[1]
                        cas = re.sub(r'-\s+', '-', cas)   # rejoin "10061-01-\n5" → "10061-01-5"
                        if " " in cas:
                            cas = cas.split()[0]   # dual-CAS rows, e.g. m/p-Xylene
                        if not _CAS_RE.match(cas):
                            continue  # not a real compound data row

                        rest = cells[2:]
                        if len(rest) < 3:
                            continue  # need >=1 conc column + LOD + LOQ
                        lod_str, loq_str = rest[-2], rest[-1]
                        conc_strs = rest[:-2]

                        lod = self._to_float(lod_str)
                        loq = self._to_float(loq_str)
                        compound = first

                        for i, raw_val in enumerate(conc_strs):
                            sid = sample_ids[i] if i < len(sample_ids) else f"Sample-{i + 1}"
                            up = raw_val.upper()
                            if up in ("N.D.", "ND", "N/D", "NOT DETECTED", ""):
                                value, flag = lod, "<LOD"
                            elif up in ("<DL", "<MDL", "<LOD", "<MRL"):
                                value, flag = lod, "<LOD"
                            elif up == "<LOQ":
                                value, flag = loq, "<LOQ"
                            else:
                                value, flag = self._vp.parse(raw_val)

                            records.append({
                                "lab":           self.LAB_NAME,
                                "sample_id":     sid,
                                "compound":      compound,
                                "cas":           cas,
                                "value":         value,
                                "flag":          flag,
                                "unit":          "µg/m³",
                                "lod":           lod,
                                "loq":           loq,
                                "analysis_type": "SOIL_GAS_VOC",
                                "canister_num":  canister_nums[i]  if i < len(canister_nums)  else "",
                                "sampling_date": analysis_times[i] if i < len(analysis_times) else "",
                                "pid_reading":   "",
                            })

        return records

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    @staticmethod
    def _clean(cell) -> str:
        if cell is None:
            return ""
        return " ".join(str(cell).replace("\u200f", "").split())

    @staticmethod
    def _to_float(s: str):
        try:
            return float(s)
        except (ValueError, TypeError):
            return None
