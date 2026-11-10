"""
Inputs of one groundwater sampling round for the periodic report updater,
for sites sampled by a field contractor with the results from ALS.

Field form ("טופס ביצוע עבודה <n> <date>.xlsx", Yeda Water & Environment):
  one sheet, row 0 is the header ("EXAMINATION::Location",
  "EXAMINATION::Well Header » Sea Depth", "EXAMINATION::Well » PH", ...).
  Row 1 holds the round (date, sampler name). Then per well: one row with
  the well header ("דיגום באר חק 2", depths, floating layer, volume, flow),
  followed by the low-flow readings, one row every five minutes. The report
  takes the LAST reading of each well (the stabilised one), which is what
  table 1 of the Hof report shows.

  Columns are found by their header text, never by position.

Lab results: ALS certificate (PDF or the "נספח" XLSX), read by the app's
ALS groundwater parser. Sample names "HK-2" are matched to the report's
wells by their number in gw_report_updater.
"""
from __future__ import annotations

import io
import re
from datetime import date, datetime

import pandas as pd

# Header text (substring of the column header) → our key
_FORM_COLUMNS = {
    "date": "Date",
    "sampler": "Schedule » User::Full Name",
    "location": "EXAMINATION::Location",
    "drill_depth": "Drill Depth",
    "water_level": "Sea Depth",
    "extraction_depth": "Extraction Depth",
    "floating": "Floating Layer",
    "volume": "Water Volume",
    "flow": "Water Flow",
    "time": "EXAMINATION::Time",
    "temp": "Temperature",
    "ph": "» PH",
    "ec": "Conductivity",
    "turbidity": "Cloudness",
    "oxygen": "Oxygen",
    "redox": "Redox",
}

# Reading key → the field name the report updater uses
_READING_FIELDS = {
    "temp": "טמפרטורה",
    "ph": "pH",
    "ec": "EC",
    "turbidity": "עכירות",
    "oxygen": "חמצן מומס",
    "redox": "רדוקס",
}


def _num(v):
    """'17.4 ליטר' → 17.4, '440 מייל' → 440.0, 22.14 → 22.14, '' → None."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    m = re.search(r"-?\d+(?:\.\d+)?", str(v).replace(",", ""))
    return float(m.group(0)) if m else None


def _find_columns(header: list) -> dict:
    cols = {}
    for key, needle in _FORM_COLUMNS.items():
        for ci, h in enumerate(header):
            h = str(h)
            if key == "floating" and "Depth" in h:
                continue  # "Floating Layer Depth" is not the yes/no column
            if key == "date" and h.strip() != "Date":
                continue
            if needle in h:
                cols[key] = ci
                break
    return cols


def _well_from_location(text: str) -> str | None:
    """'דיגום באר חק 2' → 'חק-2'."""
    m = re.search(r"([א-ת]{1,3})['׳]?\s*-?\s*(\d+)\s*$", str(text).strip())
    if not m:
        return None
    return f"{m.group(1)}-{m.group(2)}"


def _to_date(v) -> date | None:
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%d.%m.%Y", "%d/%m/%y", "%d.%m.%y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    return None


def is_yeda_field_form(file_bytes: bytes) -> bool:
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), header=None, nrows=1)
    except Exception:
        return False
    header = [str(h) for h in df.iloc[0].tolist()]
    return any("EXAMINATION::Location" in h for h in header)


def parse_yeda_field_form(file_bytes: bytes) -> dict:
    """
    Returns
        {"sampling_date": date|None, "sampler_name": str|None,
         "wells": {"חק-2": {"field": {...}, "water_level": 18.23,
                            "total_depth": 22.14, "volume_l": 17.4,
                            "flow_lpm": 0.44, "floating_layer": False}}}
    """
    df = pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=object)
    if df.empty:
        raise ValueError("טופס השדה ריק.")
    header = df.iloc[0].tolist()
    cols = _find_columns(header)
    if "location" not in cols or "time" not in cols:
        raise ValueError("זה לא טופס ביצוע עבודה של ידע מים (חסרות העמודות של שם הבאר והשעה).")

    def cell(row, key):
        ci = cols.get(key)
        if ci is None or ci >= len(row):
            return None
        v = row[ci]
        return None if (isinstance(v, float) and pd.isna(v)) else v

    sampling_date, sampler = None, None
    wells: dict = {}
    current = None
    for ri in range(1, len(df)):
        row = df.iloc[ri].tolist()
        if sampling_date is None and cell(row, "date") is not None:
            sampling_date = _to_date(cell(row, "date"))
        if sampler is None and cell(row, "sampler"):
            sampler = str(cell(row, "sampler")).strip()
        loc = cell(row, "location")
        if loc:
            well = _well_from_location(loc)
            if not well:
                current = None
                continue
            current = well
            water = _num(cell(row, "water_level"))
            extraction = _num(cell(row, "extraction_depth"))
            flow_ml = _num(cell(row, "flow"))
            field = {}
            if water is not None:
                field["עומק פני המים"] = water
            depth = _num(cell(row, "drill_depth"))
            if depth is not None:
                field["עומק כללי של הקידוח"] = depth
            if water is not None and extraction is not None:
                # Sampling depth below the water surface (owner's choice:
                # the report shows 1.0, not the depth from ground level).
                field["עומק דגימה מפני המים"] = round(extraction - water, 2)
            floating = str(cell(row, "floating") or "").strip()
            wells[well] = {
                "field": field,
                "water_level": water,
                "total_depth": depth,
                "volume_l": _num(cell(row, "volume")),
                "flow_lpm": round(flow_ml / 1000, 3) if flow_ml is not None else None,
                "floating_layer": floating.startswith("כן"),
                "_last_reading": None,
            }
            continue
        if current and cell(row, "time") is not None:
            reading = {}
            for key, name in _READING_FIELDS.items():
                v = _num(cell(row, key))
                if v is not None:
                    reading[name] = v
            if reading:
                wells[current]["_last_reading"] = reading

    for w in wells.values():
        w["field"].update(w.pop("_last_reading") or {})
    if not wells:
        raise ValueError("לא נמצאו בארות בטופס השדה.")
    return {"sampling_date": sampling_date, "sampler_name": sampler, "wells": wells}


# ALS compound name → the report's parameter name (mg/L, as in the tables)
_ALS_TO_PARAM = {
    "benzene": "בנזן",
    "toluene": "טולואן",
    "ethylbenzene": "אתיל בנזן",
    "sum of xylenes": "קסילן",
    "methyl tert-butyl ether (mtbe)": "MTBE",
    "mtbe": "MTBE",
    # parts of the xylene total, used when the certificate has no sum row
    "meta- & para-xylene": "_mp_xylene",
    "m+p-xylene": "_mp_xylene",
    "ortho-xylene": "_o_xylene",
    "o-xylene": "_o_xylene",
}

_TO_MG_L = {"mg/l": 1.0, "µg/l": 0.001, "μg/l": 0.001, "ug/l": 0.001, "g/l": 0.001}


def parse_als_round(file_bytes: bytes, filename: str = "") -> dict:
    """
    Read an ALS water certificate (PDF or XLSX) into the updater's round
    format: {"sampling_date": "01.10.26", "samples": {"HK-2": {"results":
    {"בנזן": None|float, ...}, "limits": {"בנזן": 0.0002, ...}}}}.
    A non-detect is None in results, with its reporting limit in limits,
    so the tables can print "<0.0002" as the certificate does.
    """
    import sys
    from pathlib import Path
    root = str(Path(__file__).resolve().parents[1])
    if root not in sys.path:
        sys.path.insert(0, root)
    from parsers.groundwater.als_gw import ALSGroundwaterParser

    records = ALSGroundwaterParser().parse(io.BytesIO(file_bytes))
    samples: dict = {}
    dates = set()
    for r in records:
        param = _ALS_TO_PARAM.get(str(r.get("compound", "")).strip().lower())
        if not param:
            continue
        unit = str(r.get("unit", "")).strip().lower()
        factor = _TO_MG_L.get(unit)
        if factor is None:
            continue
        s = samples.setdefault(r["sample_id"], {"results": {}, "limits": {}})
        if r.get("sampling_date"):
            dates.add(r["sampling_date"])
        flag = r.get("flag") or ""
        val = r.get("value")
        if flag.startswith("<") or flag == "ND":
            s["results"][param] = None
            lim = r.get("loq") or val
            if lim is not None:
                s["limits"][param] = round(float(lim) * factor, 10)
        elif val is not None:
            s["results"][param] = round(float(val) * factor, 10)
    for s in samples.values():
        res, lim = s["results"], s["limits"]
        parts = ("_mp_xylene", "_o_xylene")
        if "קסילן" not in res and all(p in res for p in parts):
            found = [res[p] for p in parts if res[p] is not None]
            if found:
                res["קסילן"] = round(sum(found), 10)  # a non-detect part counts as 0
            else:
                res["קסילן"] = None
                lim["קסילן"] = round(sum(lim.get(p, 0) for p in parts), 10)
        for p in parts:
            res.pop(p, None)
            lim.pop(p, None)
    if not samples:
        raise ValueError("לא נמצאו תוצאות BTEX/MTBE בקובץ של ALS.")
    sampling_date = sorted(dates)[-1] if dates else None
    return {"sampling_date": sampling_date, "samples": samples}
