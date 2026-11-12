"""
parsers/soil/machon_haneft_cert.py
----------------------------------
Machon HaNeft Excel files laid out as the printed certificate ("תעודת בדיקה" /
"תוספת מס' 1 לתעודת בדיקה", 2020-2023), one sheet each:

  metals   שם | סימול | CAS No. | <samples> | גבול גילוי | גבול כימות
  VOC      #  | Cas.No. | Compound | יחידות | <samples> | גבול הגילוי | גבול הכימות
  TPH      התכונה הנבדקת | שיטה | <samples> | גבול כימות   (DRO / ORO / C10-C40)

- The sample name is printed over several header rows above its column
  ("EL-2" / "-3", "TBM-3" / "R297-T" / "A-1"); the parts are joined
  ("EL-2-3", "TBM-3 R297-T A-1"), the same in every certificate of the job.
- The header block starts at the first row with "גבול"; the sample columns
  are between the last label column (CAS / Compound / יחידות / שיטה) and the
  first limit column. Columns are found by these words, never by position.
- "ND" is below the detection limit (value = LOD, "<LOD"); "<1" and the
  reversed "50>" are below the quantitation limit (value = 1 / 50, "<LOQ").
- Results are mg/kg dry weight. "חומר יבש" (dry matter) is not a result.
- A page break repeats the header block; those rows hold no values.
"""
from __future__ import annotations

import re

import pandas as pd

_CAS = re.compile(r"(\d{2,7}-\d{2}-\d)")
_VAL = re.compile(r"^(?:<\s*)?\d+(?:[.,]\d+)?\s*>?$|^N\.?D\.?$", re.I)
_LABELS = ("cas", "compound", "יחידות", "ש י ט ה", "שיטה", "סימול", "התכונה", "המתכת")
_END = ("סוף תעודה", "הערות", "ND –", "שיטת הכנת", "^ ", "חושב על בסיס")


def _num(s):
    try:
        return float(str(s).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def _value(raw: str, lod, loq):
    raw = raw.strip()
    if re.fullmatch(r"N\.?D\.?", raw, re.I):
        return lod, "<LOD"
    m = re.fullmatch(r"<\s*([\d.,]+)|([\d.,]+)\s*>", raw)     # "<1", or "50>" printed right to left
    if m:
        return _num(m.group(1) or m.group(2)), "<LOQ"
    return _num(raw), ""


def _join_name(parts: list[str]) -> str:
    out = ""
    for p in parts:
        p = p.strip()
        if not p:
            continue
        out = out + p if p.startswith("-") else (out + " " + p).strip()
    return out


def _tph_compound(name: str):
    n = name.replace(" ", "")
    if "חומריבש" in n:
        return None
    if re.search(r"DRO", name) and not re.search(r"ORO", name):
        return "DRO", "DRO"
    if re.search(r"ORO", name) and not re.search(r"DRO", name):
        return "ORO", "ORO"
    if "פחממנים" in n or "C10-C40" in n:
        return "TPH (C10-C40)", "C10-C40"
    return None


def parse_certificate_sheet(df: pd.DataFrame, lab: str) -> list[dict]:
    from parsers.soil.aminolab import _METALS
    rows = [[str(c).strip() for c in r] for r in df.fillna("").astype(str).values.tolist()]
    r0 = next((i for i, r in enumerate(rows) if any(c.startswith("גבול") for c in r)), None)
    if r0 is None:
        return []
    limit_cols = [j for j, c in enumerate(rows[r0]) if c.startswith("גבול")]
    first_limit = min(limit_cols)
    # the first row with a value under the limit columns ends the header block
    data0 = next((i for i in range(r0 + 1, len(rows))
                  if any(_VAL.match(rows[i][j]) for j in limit_cols if j < len(rows[i]))), None)
    if data0 is None:
        return []
    block = rows[r0:data0]

    def col_text(j):
        return " ".join(r[j] for r in block if j < len(r) and r[j])

    label_cols = [j for j in range(first_limit) if any(k in col_text(j).lower() for k in _LABELS)]
    if not label_cols:
        return []
    samples = {j: _join_name([r[j] for r in block if j < len(r)])
               for j in range(max(label_cols) + 1, first_limit)}
    samples = {j: s for j, s in samples.items() if s}
    lod_col = next((j for j in limit_cols if "גילוי" in col_text(j)), None)
    loq_col = next((j for j in limit_cols if "כימות" in col_text(j)), None)
    lowered = {j: col_text(j).lower() for j in label_cols}
    cas_col = next((j for j, t in lowered.items() if "cas" in t), None)
    name_col = next((j for j, t in lowered.items() if "compound" in t), None)
    sym_col = next((j for j, t in lowered.items() if "סימול" in t), None)
    unit_col = next((j for j, t in lowered.items() if "יחידות" in t), None)
    sheet_text = " ".join(" ".join(r) for r in rows[:data0])
    if sym_col is not None:
        kind = "SOIL_METALS"
    elif name_col is not None:
        kind = "SOIL_SVOC" if re.search(r"\bSVOC\b|semi", sheet_text, re.I) else "SOIL_VOC"
    else:
        kind = "SOIL_TPH"

    records = []
    for i in range(data0, len(rows)):
        r = rows[i]
        line = " ".join(c for c in r if c)
        if any(line.startswith(k) for k in _END):
            if line.startswith("סוף תעודה"):
                break
            continue
        if not any(j < len(r) and _VAL.match(r[j]) for j in samples):
            continue
        get = lambda j: r[j] if j is not None and j < len(r) else ""
        lod, loq = _num(get(lod_col)), _num(get(loq_col))
        if kind == "SOIL_METALS":
            sym = get(sym_col)
            if sym not in _METALS:
                continue
            compound, cas = _METALS[sym]
            m = _CAS.search(get(cas_col))
            cas = m.group(1) if m else cas
        elif kind in ("SOIL_VOC", "SOIL_SVOC"):
            compound = re.sub(r"\s*\*\s*$", "", get(name_col)).strip()
            nxt = rows[i + 1] if i + 1 < len(rows) else []
            if (nxt and name_col < len(nxt) and nxt[name_col].startswith("(")
                    and not any(_VAL.match(c) for c in nxt)):
                compound += " " + nxt[name_col]          # "Isopropylbenzene" / "(Cumene)"
            compound = re.sub(r"\s*\*\s*", " ", compound).strip()
            m = _CAS.search(get(cas_col))
            cas = m.group(1) if m else ""
            if not compound:
                continue
        else:
            name = " ".join(c for c in r[:min(samples)] if c)
            nxt = rows[i + 1] if i + 1 < len(rows) else []
            if nxt and not any(_VAL.match(c) for c in nxt) and nxt[0].startswith("("):
                name += " " + nxt[0]                     # "1.תכולת פחממנים" / "(C10-C40), מ"ג/ק"ג"
            found = _tph_compound(name)
            if not found:
                continue
            compound, cas = found
        unit = get(unit_col).replace("Kg", "kg") or "mg/kg"
        for j, sample in samples.items():
            raw = get(j)
            if not _VAL.match(raw):
                continue
            value, flag = _value(raw, lod, loq)
            records.append({
                "lab": lab, "sample_id": sample, "compound": compound, "cas": cas,
                "value": value, "flag": flag, "unit": unit,
                "lod": lod, "loq": loq if loq is not None else (value if flag == "<LOQ" else None),
                "analysis_type": kind,
            })
    return records
