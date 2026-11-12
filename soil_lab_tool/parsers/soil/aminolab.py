"""
parsers/soil/aminolab.py
------------------------
Aminolab soil (and sludge / waste) certificates: the PDF "תעודה לתוצאות
המעבדה" and its Excel appendix ("נספח לדוח").

PDF (text, logical order through PyMuPDF; read by word positions):
  - Each sample starts with "תאור הדוגמה: קרקע - קר-1 - (0.1) - טופס נטילה: 012601"
    and runs until the next one (several pages: its tables go on).
  - A Hebrew table "הבדיקה | יחידות מידה | תוצאה | הערות": pH, mineral oil,
    DRO+ORO, GRO, and the ICP metals under "סריקת מתכות מלאה ב- ICP"
    ("Ag - כסף" ...). Some rows are printed a little above or below their
    name line, so each value goes to the nearest name line.
  - English tables "Compound | CAS No. | Results (ppb|ppm) | Library Search
    Quality" under "VOC" (ppb) and "SVOC - Semi-Volatiles" (ppm), going on
    over pages (the header repeats). A leading "*" marks an accredited test.
  - "<20" is below the limit of quantitation (LOQ = 20): value 20, "<LOQ".
    "ND" is below the detection limit.
  Results are on a dry-weight basis (the certificate's note); ppm = mg/kg, ppb
  becomes mg/kg (÷1000).

Excel appendix: one header row "הבדיקה | יחידות מידה | <one column per
sample>" (the sample cell: "קרקע B1 טופס נטילה 013319\n12/03/2025, שעה: 10:30"),
then the same rows as the PDF, with the section rows "SVOC - Semi-Volatiles"
/ "VOC" and no unit on their compounds (the PDF's: VOC ppb, SVOC ppm).

A scanned certificate (no text) gives no records.
"""
from __future__ import annotations

import io
import re

from parsers.base import BaseParser

# ICP symbols → (English name, CAS)
_METALS = {
    "Ag": ("Silver", "7440-22-4"), "Al": ("Aluminium", "7429-90-5"),
    "As": ("Arsenic", "7440-38-2"), "B": ("Boron", "7440-42-8"),
    "Ba": ("Barium", "7440-39-3"), "Be": ("Beryllium", "7440-41-7"),
    "Ca": ("Calcium", "7440-70-2"), "Cd": ("Cadmium", "7440-43-9"),
    "Co": ("Cobalt", "7440-48-4"), "Cr": ("Chromium", "7440-47-3"),
    "Cu": ("Copper", "7440-50-8"), "Fe": ("Iron", "7439-89-6"),
    "Hg": ("Mercury", "7439-97-6"), "K": ("Potassium", "7440-09-7"),
    "Li": ("Lithium", "7439-93-2"), "Mg": ("Magnesium", "7439-95-4"),
    "Mn": ("Manganese", "7439-96-5"), "Mo": ("Molybdenum", "7439-98-7"),
    "Na": ("Sodium", "7440-23-5"), "Ni": ("Nickel", "7440-02-0"),
    "P": ("Phosphorus", "7723-14-0"), "Pb": ("Lead", "7439-92-1"),
    "S": ("Sulfur", "7704-34-9"), "Sb": ("Antimony", "7440-36-0"),
    "Se": ("Selenium", "7782-49-2"), "Sn": ("Tin", "7440-31-5"),
    "Sr": ("Strontium", "7440-24-6"), "Ti": ("Titanium", "7440-32-6"),
    "Tl": ("Thallium", "7440-28-0"), "V": ("Vanadium", "7440-62-2"),
    "Zn": ("Zinc", "7440-66-6"),
}

# Hebrew-table rows that are not metals: name pattern → (compound, cas, type)
_HEB_ROWS = [
    (r"DRO\s*[+-]\s*ORO", ("DRO-ORO", "DRO-ORO", "SOIL_TPH")),
    (r"^GRO$", ("GRO", "GRO", "SOIL_TPH")),
    (r"^DRO$", ("DRO", "DRO", "SOIL_TPH")),
    (r"^ORO$", ("ORO", "ORO", "SOIL_TPH")),
    (r"TPH", ("TPH", "TPH", "SOIL_TPH")),
    (r"שמן מינרלי", ("Mineral oil (שמן מינרלי)", "", "SOIL_TPH")),
]

# CAS numbers as Aminolab prints them in its VOC / SVOC tables (the Excel
# appendix has no CAS column); each passes the CAS check digit.
_AMINOLAB_CAS = {
    "4-nitroaniline": "100-01-6",   # printed "100-01-8", which fails the check digit
    "1,1,1,2 tetrachlorethane": "630-20-6",
    "1,1,1-trichloroethane": "71-55-6",
    "1,1,2,2-tetrachloroethane": "79-34-5",
    "1,1,2-trichloroethane": "79-00-5",
    "1,1-dichloroethane": "75-34-3",
    "1,1-dichloroethylene": "75-35-4",
    "1,1-dichloropropene": "563-58-6",
    "1,2,3-trichlorobenzene": "87-61-6",
    "1,2,3-trichloropropane": "96-18-4",
    "1,2,4-trichlorobenzene": "120-82-1",
    "1,2,4-trimethylbenzene": "95-63-6",
    "1,2-dibromo-3-chloropropane": "96-12-8",
    "1,2-dibromoethane (ethylene": "106-93-4",
    "1,2-dibromoethane (ethylene dibromide)": "106-93-4",
    "1,2-dichlorobenzene": "95-50-1",
    "1,2-dichloroethane": "107-06-2",
    "1,2-dichloropropane": "78-87-5",
    "1,3,5-trimethylbenzene": "108-67-8",
    "1,3-dichlorobenzene": "541-73-1",
    "1,3-dichloropropane": "142-28-9",
    "1,4-dichlorobenzene": "106-46-7",
    "1-chloronaphthalene": "90-13-1",
    "1-methylnaphtalene": "90-12-0",
    "2,2-dichloropropane": "594-20-7",
    "2,4,5-trichlorophenol": "95-95-4",
    "2,4,6-trichlorophenol": "88-06-2",
    "2,4-dimethylphenol": "105-67-9",
    "2,4-dinitrotoluene": "121-14-2",
    "2,6-dinitrotoluene": "606-20-2",
    "2-chloronaphthalene": "91-58-7",
    "2-chlorotoluene": "95-49-8",
    "2-methyl phenol": "95-48-7",
    "2-methylnaphthalene": "91-57-6",
    "2-nitroaniline": "88-74-4",
    "2-nitrophenol": "88-75-5",
    "3,3-dichlorobenzidine": "91-94-1",
    "3-nitroaniline": "99-09-2",
    "4,6-dinitro-2-methylphenol": "534-52-1",
    "4-bromophenyl phenyl ether": "101-55-3",
    "4-chloro-3-methylphenol": "59-50-7",
    "4-chloroaniline": "106-47-8",
    "4-chlorophenyl phenyl ether": "7005-72-3",
    "4-chlorotoluene": "106-43-4",
    "4-methyl phenol": "106-44-5",
    "4-nitrophenol": "100-02-7",
    "acenaphthylene": "208-96-8",
    "acenaphthylene, 1,2-dihydro": "83-32-9",
    "acetophenone": "98-86-2",
    "aniline": "62-53-3",
    "anthracene": "120-12-7",
    "benzene": "71-43-2",
    "benzenemethanol": "100-51-6",
    "benzidine": "92-87-5",
    "benzo (a) anthracene": "56-55-3",
    "benzo (a) pyrene": "50-32-8",
    "benzo (b) fluoranthene": "205-99-2",
    "benzo (g,h,i) perylene": "191-24-2",
    "benzo (k) fluoranthene": "207-08-9",
    "bis (2-chloroethoxy) methane": "111-91-1",
    "bis (2-chloroethyl) ether": "111-44-4",
    "bis (2-chloroisopropyl) ether": "108-60-1",
    "bromobenzene": "108-86-1",
    "bromochloromethan": "74-97-5",
    "bromodichloromethane": "75-27-4",
    "bromoform": "75-25-2",
    "butyl benzyl phthalate": "85-68-7",
    "carbon tetrachloride": "56-23-5",
    "chlorobenzene": "108-90-7",
    "chlorodibromomethane": "124-48-1",
    "chloroethane": "75-00-3",
    "chloroform": "67-66-3",
    "chrysene": "218-01-9",
    "cis-1,2-dichloroethene": "156-59-2",
    "cis-1,3-dichloropropene": "10061-01-5",
    "di-(2-ethylhexyl) phthalate": "117-81-7",
    "di-n-butyl phthalate": "84-74-2",
    "dibenz (a,h) anthracene": "53-70-3",
    "dibenzofuran": "132-64-9",
    "dibromomethane": "74-95-3",
    "dichlorodifluoromethane": "75-71-8",
    "diethyl phthalate": "84-66-2",
    "dimethyl phthalate": "131-11-3",
    "diphenylamine": "122-39-4",
    "ethyl methane sulfonate": "62-50-0",
    "ethylbenzene": "100-41-4",
    "fluoranthene": "206-44-0",
    "fluorene": "86-73-7",
    "hexachloro-1,3-butadiene": "87-68-3",
    "hexachloroethane": "67-72-1",
    "indeno (1,2,3-cd) pyrene": "193-39-5",
    "isophorone": "78-59-1",
    "isopropyl benzene": "98-82-8",
    "methane, chloro": "74-87-3",
    "methyl bromide": "74-83-9",
    "methylene chloride": "75-09-2",
    "mtbe": "1634-04-4",
    "n-butyl benzene": "104-51-8",
    "n-nitrosodi-n-butylamine": "924-16-3",
    "n-nitrosodi-n-propylamine": "621-64-7",
    "n-nitrosodimethylamine": "62-75-9",
    "n-propylbenzene": "103-65-1",
    "naphthalene": "91-20-3",
    "nitrobenzene": "98-95-3",
    "o-xylene": "95-47-6",
    "p-isopropyl toluene": "99-87-6",
    "phenanthrene": "85-01-8",
    "phenol": "108-95-2",
    "pyrene": "129-00-0",
    "pyridine": "110-86-1",
    "sec-butyl benzene": "135-98-8",
    "styrene": "100-42-5",
    "tert-butyl benzene": "98-06-6",
    "tetrachloroethene": "127-18-4",
    "toluene": "108-88-3",
    "trans-1,2-dichloroethene": "156-60-5",
    "trans-1,3-dichloropropene": "10061-02-6",
    "trichloroethene": "79-01-6",
    "trichloromonofluoromethane": "75-69-4",
    "vinyl chloride": "75-01-4",
    "xylenes": "1330-20-7",
}

_CAS_RE = re.compile(r"^\d{2,7}-\d{2}-\d$")
_VAL_RE = re.compile(r"^(?:<\s*)?\d+(?:[.,]\d+)?$|^ND$|^N\.D\.?$", re.I)
_END_KW = ("חתימה", "אושר", "הערות לבדיקה", "גבול", "נבדקה", "נתונות על בסיס",
           "מיהול", "הנתונים", "יש להתייחס", "אין הערות")


def _cas_ok(cas: str) -> bool:
    """The CAS check digit: a misprint ("95-18-4" for 96-18-4) fails it."""
    d = cas.replace("-", "")
    return d.isdigit() and int(d[-1]) == sum(int(x) * (i + 1) for i, x in enumerate(reversed(d[:-1]))) % 10


def _known_cas(name: str) -> str:
    return _AMINOLAB_CAS.get(re.sub(r"\s+", " ", name.lower()).strip(), "")


def _num(s):
    try:
        return float(str(s).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def _value(raw: str, factor: float):
    """'<20' → (20·f, '<LOQ', loq), 'ND' → (None, 'ND', None), '330' → (330·f, '', None)."""
    raw = raw.strip()
    if re.fullmatch(r"N\.?D\.?", raw, re.I):
        return None, "ND", None
    if raw.startswith("<"):
        v = _num(raw[1:])
        v = round(v * factor, 9) if v is not None else None
        return v, "<LOQ", v
    v = _num(raw)
    if v is None:
        return None, None, None
    return round(v * factor, 9), "", None


def _mirror(s: str) -> str:
    """PyMuPDF keeps brackets of RTL lines as drawn: ')a(' → '(a)'."""
    return re.sub(r"\)([^()]*)\(", r"(\1)", s)


def _sample_name(text: str) -> str:
    """'קרקע - קר-1 - (0.1) - טופס נטילה: 012601' → 'קר-1 (0.1)'; 'קרקע B1 טופס נטילה 013319' → 'B1'."""
    t = text.split("\n")[0]
    t = re.sub(r"טופס\s+נטילה\s*:?\s*\d*", "", t)
    t = re.sub(r"^\s*(?:קרקע|בוצה|פסולת|מים|תשטיף|דוגמה|סחף)\b", "", t.strip())
    t = re.sub(r"\s*-\s*\(", " (", t)
    t = re.sub(r"([א-תA-Za-z]+)\s*\(([\d.]+)\)-(\d+)", r"\1-\3 (\2)", t)   # "מכ (0.3)-1"
    t = re.sub(r"(?<=[א-תA-Za-z])\s+-(?=\d)", "-", t)     # "קר -1" → "קר-1"
    t = re.sub(r"\s+", " ", t).strip(" -:")
    return t


def _heb_row(name: str, group: str | None):
    """(compound, cas, analysis_type) for a row of the Hebrew table."""
    m = re.search(r"\b([A-Z][a-z]?)\s*-\s*[א-ת]", name) or re.search(r"[א-ת]\s*-\s*([A-Z][a-z]?)\b", name)
    if m and m.group(1) in _METALS:
        en, cas = _METALS[m.group(1)]
        return en, cas, "SOIL_METALS"
    for pat, out in _HEB_ROWS:
        if re.search(pat, name, re.I):
            return out
    if group == "metals":
        return name, "", "SOIL_METALS"
    return name, "", "SOIL_NUTRIENTS"     # pH, organic matter, ... (the chemistry sheet)


# ── PDF ─────────────────────────────────────────────────────────────────────

def _lines(words, tol=2.5):
    """Words grouped into lines by their vertical centre: [(y, [words])]."""
    out = []
    for w in sorted(words, key=lambda w: (w[1] + w[3]) / 2):
        y = (w[1] + w[3]) / 2
        if out and abs(out[-1][0] - y) <= tol:
            out[-1][1].append(w)
        else:
            out.append([y, [w]])
    return out


def _rtl(ws) -> str:
    return " ".join(w[4] for w in sorted(ws, key=lambda w: -w[0]))


def _ltr(ws) -> str:
    return " ".join(w[4] for w in sorted(ws, key=lambda w: w[0]))


def _assign(rows, others, max_dy=7.0):
    """Each value word goes to the name line nearest to it (rows: [(y, ws)])."""
    for w in others:
        y = (w[1] + w[3]) / 2
        best = min(rows, key=lambda r: abs(r[0] - y), default=None)
        if best is not None and abs(best[0] - y) <= max_dy:
            best[2].append(w)


class AminolabSoilParser(BaseParser):
    LAB_NAME = "Aminolab"
    ANALYSIS_TYPES = ["SOIL_TPH", "SOIL_METALS", "SOIL_VOC", "SOIL_SVOC", "SOIL_NUTRIENTS"]

    def parse(self, file_obj) -> list[dict]:
        data = file_obj.read() if hasattr(file_obj, "read") else open(file_obj, "rb").read()
        if data[:2] == b"PK":
            records = self._parse_xlsx(data)
            if records:
                return records
            # an older soil-nutrients appendix (sample columns "קרקע h-1 (4m)")
            from parsers.groundwater.aminolab import AminolabGroundwaterParser
            return AminolabGroundwaterParser().parse(io.BytesIO(data))
        return self._parse_pdf(data)

    # ------------------------------------------------------------------
    def _parse_pdf(self, data: bytes) -> list[dict]:
        import fitz
        records: list[dict] = []
        sample = None
        mode = None          # "heb" | "eng"
        section = None       # "VOC" | "SVOC"
        factor = 1.0
        group = None
        cols = {}
        eng_res = {}
        with fitz.open(stream=data, filetype="pdf") as doc:
            for page in doc:
                lines = _lines(page.get_text("words"))
                mode = None   # the table header repeats on every page
                table_rows = []   # [y, name words, value words] of the current table
                pending = []      # value-zone words of the current table

                def flush():
                    nonlocal table_rows, pending
                    if table_rows:
                        _assign(table_rows, pending)
                        for y, nws, vws, kind, sec, fac, grp, smp in table_rows:
                            rec = self._row(kind, nws, vws, cols_at[id(nws)], sec, fac, grp, smp)
                            if rec:
                                records.append(rec)
                    table_rows, pending = [], []

                cols_at = {}
                for y, ws in lines:
                    rtl = _rtl(ws)
                    ltr = _ltr(ws)
                    if "תאור" in rtl and "הדוגמה" in rtl:
                        flush()
                        txt = rtl.split(":", 1)[1] if ":" in rtl else rtl
                        sample = _sample_name(_mirror(txt))
                        mode, section, group = None, None, None
                        continue
                    heads = {w[4]: w[0] for w in ws}
                    if "הבדיקה" in heads and ("תוצאה" in heads or "יחידות" in heads):
                        flush()
                        mode = "heb"
                        cols = {"name": heads["הבדיקה"], "unit": heads.get("יחידות", heads.get("מידה")),
                                "result": heads.get("תוצאה"), "notes": heads.get("הערות")}
                        continue
                    if "Results" in heads and "Library" in heads:
                        eng_res = {"result": heads["Results"], "quality": heads["Library"]}
                        continue
                    if "Compound" in heads and "CAS" in heads:
                        flush()
                        mode = "eng"
                        cols = {"cas": heads["CAS"], **eng_res}
                        continue
                    if re.search(r"\(ppb\)", ltr):
                        factor = 0.001
                        if mode != "eng":
                            continue
                    if re.search(r"\(ppm\)", ltr):
                        factor = 1.0
                        if mode != "eng":
                            continue
                    if re.fullmatch(r"\(pp[mb]\)(?: Quality)?|Quality", ltr):
                        continue
                    if re.fullmatch(r"VOC", ltr.strip()):
                        flush()
                        section, factor, mode = "VOC", 0.001, None
                        continue
                    if re.search(r"\bSVOC\b", ltr) and "Semi" in ltr:
                        flush()
                        section, factor, mode = "SVOC", 1.0, None
                        continue
                    if mode is None or sample is None:
                        continue
                    if any(k in rtl for k in _END_KW):
                        flush()
                        mode = None
                        continue
                    if mode == "heb":
                        name_cut = (cols["name"] + (cols["unit"] or cols["name"] - 200)) / 2
                        nws = [w for w in ws if w[0] >= name_cut]
                        vws = [w for w in ws if w[0] < name_cut]
                        if nws:
                            name = _mirror(_rtl(nws))
                            if "סריקת" in name and "מתכות" in name:
                                group = "metals"
                            row = [y, nws, [], "heb", section, factor, group, sample]
                            cols_at[id(nws)] = dict(cols)
                            table_rows.append(row)
                        pending.extend(vws)
                    else:
                        cas_x = cols["cas"] - 12
                        nws = [w for w in ws if w[2] <= cas_x + 4 and w[0] < cas_x]
                        vws = [w for w in ws if w not in nws]
                        if nws:
                            row = [y, nws, [], "eng", section, factor, group, sample]
                            cols_at[id(nws)] = dict(cols)
                            table_rows.append(row)
                        pending.extend(vws)
                flush()
        return records

    # ------------------------------------------------------------------
    def _row(self, kind, nws, vws, cols, section, factor, group, sample):
        if kind == "heb":
            name = _mirror(_rtl(nws)).strip()
            res_x, unit_x = cols.get("result"), cols.get("unit")
            # the value is the word in the result column: nearest to its header
            cand = [w for w in vws if _VAL_RE.match(w[4].replace(" ", ""))
                    and (unit_x is None or w[2] < unit_x - 5)
                    and (cols.get("notes") is None or w[0] > (cols["notes"] + (res_x or 0)) / 2)]
            if not cand:
                return None
            raw = min(cand, key=lambda w: abs(w[0] - (res_x or w[0])))[4]
            units = [w[4] for w in vws if unit_x is not None and abs(w[0] - unit_x) < 40 and re.search(r"[A-Za-z]", w[4])]
            unit = units[0] if units else ""
            fac = 0.001 if unit.lower() in ("ppb", "µg/kg", "ug/kg") else 1.0
            compound, cas, atype = _heb_row(name, group)
            value, flag, loq = _value(raw, fac)
            if flag is None:
                return None
            out_unit = "mg/kg" if atype != "SOIL_NUTRIENTS" or unit.lower() in ("mg/kg", "ppm") else unit
            return self._rec(sample, compound, cas, value, flag, loq, out_unit, atype)
        # English table
        name = re.sub(r"\s*\*\s*", " ", _mirror(_ltr(nws))).strip()
        if not name or name.lower().startswith(("compound", "results")):
            return None
        if re.search(r"[א-ת]", name):
            return None    # a Hebrew note under the table, not a compound
        cas_x, res_x, q_x = cols.get("cas"), cols.get("result"), cols.get("quality")
        lim = ((res_x or 400) + (q_x or 480)) / 2
        cas_ws = [w for w in vws if w[0] < (res_x or 380) - 15]
        res_ws = [w for w in vws if (res_x or 380) - 15 <= w[0] < lim]
        cas = re.sub(r"\s+", "", "".join(w[4] for w in sorted(cas_ws, key=lambda w: w[0])))
        cas = re.sub(r"-+", "-", cas)
        if not _CAS_RE.match(cas) or not _cas_ok(cas):
            cas = _known_cas(name)
        if name.lower().startswith("xylen"):
            cas = cas or "1330-20-7"
            name = "Xylenes"
        raw = next((w[4] for w in res_ws if _VAL_RE.match(w[4])), None)
        if raw is None:
            return None
        value, flag, loq = _value(raw, factor)
        if flag is None:
            return None
        atype = "SOIL_VOC" if section == "VOC" else "SOIL_SVOC"
        return self._rec(sample, name, cas, value, flag, loq, "mg/kg", atype)

    def _rec(self, sample, compound, cas, value, flag, loq, unit, atype):
        return {
            "lab": self.LAB_NAME, "sample_id": sample, "compound": compound, "cas": cas,
            "value": value, "flag": flag, "unit": unit, "lod": None, "loq": loq,
            "analysis_type": atype,
        }

    # ------------------------------------------------------------------
    def _parse_xlsx(self, data: bytes) -> list[dict]:
        import pandas as pd
        records = []
        for _, df in pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, dtype=str).items():
            rows = df.fillna("").astype(str).values.tolist()
            hi = next((i for i, r in enumerate(rows)
                       if r and r[0].strip() == "הבדיקה" and any("יחידות" in c for c in r[1:3])), None)
            if hi is None:
                continue
            samples = {}
            for ci, c in enumerate(rows[hi]):
                if ci >= 2 and c.strip() and "הערות" not in c:
                    if re.search(r"\bh[-_]?\d+\b", c, re.I):
                        return []          # the older nutrients appendix: its own reader
                    samples[ci] = _sample_name(c)
            section, group = None, None
            for r in rows[hi + 1:]:
                name = r[0].strip()
                if not name or name.startswith("הערות"):
                    if name.startswith("הערות"):
                        break
                    continue
                if re.fullmatch(r"VOC.*", name) and "SVOC" not in name:
                    section, group = "VOC", None
                    continue
                if name.startswith("SVOC"):
                    section, group = "SVOC", None
                    continue
                if "סריקת מתכות" in name:
                    group = "metals"
                    continue
                unit = r[1].strip() if len(r) > 1 else ""
                for ci, smp in samples.items():
                    raw = r[ci].strip() if ci < len(r) else ""
                    if not raw:
                        continue
                    if section in ("VOC", "SVOC") and not unit:
                        fac = 0.001 if section == "VOC" else 1.0
                        value, flag, loq = _value(raw, fac)
                        if flag is None:
                            continue
                        cname = re.sub(r"^\*\s*", "", name).strip()
                        cas = ""
                        if cname.lower().startswith("xylen"):
                            cname, cas = "Xylenes", "1330-20-7"
                        records.append(self._rec(smp, cname, cas, value, flag, loq, "mg/kg",
                                                 "SOIL_VOC" if section == "VOC" else "SOIL_SVOC"))
                        continue
                    fac = 0.001 if unit.lower() in ("ppb", "µg/kg", "ug/kg") else 1.0
                    value, flag, loq = _value(raw, fac)
                    if flag is None:
                        continue
                    compound, cas, atype = _heb_row(name, group)
                    out_unit = "mg/kg" if atype != "SOIL_NUTRIENTS" or unit.lower() in ("mg/kg", "ppm") else unit
                    records.append(self._rec(smp, compound, cas, value, flag, loq, out_unit, atype))
        _fill_cas(records)
        return records


def _fill_cas(records):
    """Excel has no CAS column: take it from the shared name lookup."""
    from core.cas_lookup import name_to_cas
    for r in records:
        if not r["cas"] and r["analysis_type"] in ("SOIL_VOC", "SOIL_SVOC"):
            r["cas"] = _known_cas(r["compound"]) or name_to_cas(r["compound"]) or ""
