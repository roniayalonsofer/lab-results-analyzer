"""
parsers/groundwater/edd.py
--------------------------
Water results as an EDD (electronic data deliverable) sheet: one row per
result, columns named in the first row, e.g. the ALS Houston file KTE sends
for 1,4-dioxane ("HS23031751 KTE27823H EDD.xlsx"):

    ClientSampID | CollectionDate | TestName | AnalyteType | Analyte | Result |
    Qualifier | FinalResult | Runits | Rlimit | CAS | TestNo | SampID | ...

- Columns are found by their header name, never by position.
- AnalyteType "A" is a result; "S" (surrogate) and other QC types are skipped.
- A non-detect: "ND" / empty result, a "U" qualifier, or "<x" → the limit with
  flag "<LOQ" (Rlimit is the reporting limit).
- µg/L becomes mg/L, the groundwater sheets' common unit (as the ALS reader).
"""
from __future__ import annotations

import re

import pandas as pd

_SAMPLE = ("clientsampid", "client sample id", "clientsampleid", "sample name")
_ANALYTE = ("analyte", "parameter", "compound")


def _key(h) -> str:
    return re.sub(r"[\s_]+", "", str(h).strip().lower())


def edd_header(df: pd.DataFrame) -> dict | None:
    """{field: column index} when the first rows hold an EDD header, else None."""
    for ri in range(min(5, len(df))):
        hdr = [_key(h) for h in df.iloc[ri].tolist()]
        def col(*names):
            return next((i for i, h in enumerate(hdr) if h in {_key(n) for n in names}), None)
        c = {
            "sample": col(*_SAMPLE),
            "analyte": col(*_ANALYTE),
            "result": col("result"),
            "final": col("finalresult"),
            "qual": col("qualifier", "qual"),
            "unit": col("runits", "units", "unit"),
            "limit": col("rlimit", "reportinglimit", "rl", "mrl", "loq"),
            "mdl": col("mdl", "dl", "lod"),
            "cas": col("cas", "casno", "casnumber"),
            "date": col("collectiondate", "sampledate", "datecollected"),
            "test": col("testname", "method", "testno"),
            "atype": col("analytetype"),
            "row": ri,
        }
        if c["sample"] is not None and c["analyte"] is not None and c["result"] is not None:
            return c
    return None


def _num(v):
    try:
        return float(str(v).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def _analysis_type(test: str, unit: str) -> str:
    t = test.lower()
    if unit.lower() == "ng/l" or "pfas" in t or "perfluor" in t or "537" in t or "1633" in t:
        return "GW_PFAS"
    if "metal" in t or "6010" in t or "6020" in t or "200.8" in t:
        return "GW_METALS"
    if "semivol" in t or "8270" in t or "pah" in t:
        return "GW_SVOC"
    return "GW_VOC"


def _date(v: str) -> str:
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", v)          # US: month/day/year
    if m:
        return f"{int(m.group(2)):02d}.{int(m.group(1)):02d}.{m.group(3)}"
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", v)
    if m:
        return f"{m.group(3)}.{m.group(2)}.{m.group(1)}"
    return ""


def parse_edd(df: pd.DataFrame, lab: str) -> list[dict]:
    c = edd_header(df)
    if c is None:
        return []
    rows = df.fillna("").astype(str).values.tolist()[c["row"] + 1:]

    def get(r, k):
        i = c[k]
        return r[i].strip() if i is not None and i < len(r) else ""

    records = []
    for r in rows:
        sample, compound = get(r, "sample"), get(r, "analyte")
        if not sample or not compound:
            continue
        at = get(r, "atype").upper()
        if at and at != "A":
            continue  # surrogate / QC
        raw = get(r, "result") or get(r, "final")
        qual = get(r, "qual")
        unit = get(r, "unit") or "µg/L"
        limit, mdl = _num(get(r, "limit")), _num(get(r, "mdl"))
        if raw.startswith("<"):
            value, flag = _num(raw[1:]) or limit, "<LOQ"
        elif raw.upper() in ("", "ND", "N.D.", "U") or "U" in qual.upper().split():
            value, flag = limit, "<LOQ"
        elif _num(raw) is not None:
            value, flag = _num(raw), ""
        else:
            continue
        atype = _analysis_type(get(r, "test"), unit)
        if unit.lower() in ("ug/l", "µg/l", "μg/l"):
            f = 1 / 1000.0
            value = round(value * f, 9) if value is not None else None
            limit = round(limit * f, 9) if limit is not None else None
            mdl = round(mdl * f, 9) if mdl is not None else None
            unit = "mg/L"
        d = _date(get(r, "date"))
        rec = {
            "lab": lab,
            "sample_id": f"{sample} ({d})" if d else sample,
            "compound": compound,
            "cas": get(r, "cas"),
            "value": value,
            "flag": flag,
            "unit": unit,
            "lod": mdl,
            "loq": limit,
            "analysis_type": atype,
        }
        if qual and qual.upper() != "U":
            rec["qualifier"] = qual
        records.append(rec)
    return records
