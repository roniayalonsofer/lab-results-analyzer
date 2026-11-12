"""
parsers/soil_gas/als_to15.py
-----------------------------
ALS Environmental (Simi Valley) soil gas, EPA TO-15, Excel "RESULTS OF ANALYSIS".

Layout (one sheet per sample, "Sample", "Sample (2)", ...):
  - "Client Sample ID:" in column A, the sample name three columns to the right.
  - Every page of the sheet repeats the header block, then a table:
        CAS # | | | Compound | Result | (ND) | MRL | MDL | Result | | MRL | MDL | Data Qualifier
                                µg/m³          µg/m³  µg/m³   ppbV ...
    A detected value sits in the first "Result" column; a non-detect is
    written in the column after it as "<MDL" ("<0.36").
  - "J" in the qualifier column: detected between the MDL and the MRL (the
    value is kept as reported).
  - Sheets "MBlank", "Surrogates", "LCS" are laboratory quality control: skipped.
Results are kept in µg/m³ (the unit the soil gas thresholds use).
"""
from __future__ import annotations

import io
import re

import pandas as pd

from parsers.base import BaseParser

_QC_SHEETS = ("mblank", "method blank", "surrogate", "lcs", "duplicate")
_CAS = re.compile(r"^\d{2,7}-\d{2}-\d$")


def _num(v):
    try:
        return float(str(v).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


class ALSSoilGasTO15Parser(BaseParser):
    LAB_NAME = "ALS"
    ANALYSIS_TYPES = ["SOIL_GAS_VOC"]

    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        sheets = pd.read_excel(file_obj, sheet_name=None, header=None, dtype=str)
        records = []
        for name, df in sheets.items():
            if any(q in name.lower() for q in _QC_SHEETS):
                continue
            df = df.fillna("")
            rows = df.values.tolist()
            sample = None
            for r in rows:
                if str(r[0]).strip().lower().startswith("client sample id"):
                    sample = next((str(x).strip() for x in r[1:] if str(x).strip()), None)
                    break
            if not sample or "blank" in sample.lower():
                continue
            cols = None
            seen = set()
            for r in rows:
                cells = [str(x).strip() for x in r]
                low = [c.lower() for c in cells]
                if "compound" in low or ("result" in low and "mrl" in low):
                    # the header (its parts may be split over two rows on later pages)
                    res = [i for i, c in enumerate(low) if c == "result"]
                    mrl = [i for i, c in enumerate(low) if c == "mrl"]
                    mdl = [i for i, c in enumerate(low) if c == "mdl"]
                    if res and mrl and mdl:
                        cols = {"res": res[0], "mrl": mrl[0], "mdl": mdl[0],
                                "qual": next((i for i, c in enumerate(low) if "qualifier" in c or c == "data"), None)}
                    continue
                if cols is None:
                    continue
                cas = next((c for c in cells[:3] if _CAS.match(c)), None)
                if not cas:
                    continue
                compound = next((c for c in cells[2:cols["res"]] if c and not _CAS.match(c)), "")
                compound = re.sub(r"-\s*\n\s*", "-", compound)
                compound = re.sub(r"\s+", " ", compound).strip()
                if not compound or (sample, cas) in seen:
                    continue
                seen.add((sample, cas))
                raw = cells[cols["res"]]
                nd = cells[cols["res"] + 1] if cols["res"] + 1 < len(cells) else ""
                mrl, mdl = _num(cells[cols["mrl"]]), _num(cells[cols["mdl"]])
                if nd.startswith("<"):
                    value, flag = _num(nd[1:]), "<LOD"
                elif _num(raw) is not None:
                    value, flag = _num(raw), ""
                else:
                    continue
                rec = {
                    "lab": self.LAB_NAME,
                    "sample_id": sample,
                    "compound": compound,
                    "cas": cas,
                    "value": value,
                    "flag": flag,
                    "unit": "µg/m³",
                    "lod": mdl,
                    "loq": mrl,
                    "analysis_type": "SOIL_GAS_VOC",
                }
                q = cells[cols["qual"]] if cols.get("qual") is not None and cols["qual"] < len(cells) else ""
                if q:
                    rec["qualifier"] = q
                records.append(rec)
        return records
