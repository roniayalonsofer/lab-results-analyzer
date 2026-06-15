"""
parsers/soil/als.py
--------------------
Parser for ALS laboratory soil reports.

Sheet format ("Client SOIL 1"):
  Row  8 (idx 7 ): sample IDs in columns idx 4+
  Row 13 (idx 12): column headers — Parameter | Method | Unit | LOR | <sample cols>
  Rows 14+       : compound at idx 0, unit at idx 2, LOR at idx 3, values at idx 4+

  Values like "<0.050" = below LOR (flag <LOQ, value = numeric after <).
  "ND" / "N.D." / "----" / empty = not detected (flag <LOQ, value = LOR).

ALSGrainSizeParser uses the same sheet format but recognises fraction parameters
("Fraction X-Y mm") and tags records as SOIL_GRAIN_SIZE.
"""

from __future__ import annotations

import io
import re
import xml.etree.ElementTree as _ET

import pandas as pd

from parsers.base import BaseParser
from core.cas_lookup import fuzzy_name_to_cas


# ---------------------------------------------------------------------------
# SpreadsheetML helpers
# ---------------------------------------------------------------------------

_SML_NS = "urn:schemas-microsoft-com:office:spreadsheet"


def _parse_spreadsheetml(data: bytes) -> dict[str, pd.DataFrame]:
    """Parse a SpreadsheetML XML file into {sheet_name: DataFrame}.

    Handles sparse rows (ss:Index), merged cells (ss:MergeAcross), and both
    String and Number cell types.  Each DataFrame has all-string dtype and
    empty strings in place of missing values — identical to what
    pd.ExcelFile.parse(..., header=None, dtype=str).fillna("") produces.
    """
    root = _ET.fromstring(data)
    # Detect the actual namespace URI from the root tag (may differ slightly)
    ns_uri = _SML_NS
    if root.tag.startswith("{"):
        ns_uri = root.tag[1:root.tag.index("}")]
    ns = {"ss": ns_uri}

    sheets: dict[str, pd.DataFrame] = {}

    for ws in root.findall(".//ss:Worksheet", ns):
        name = ws.get(f"{{{ns_uri}}}Name", "")
        table = ws.find("ss:Table", ns)
        if table is None:
            sheets[name] = pd.DataFrame()
            continue

        rows_data: list[list[str]] = []
        for row_el in table.findall("ss:Row", ns):
            cells: list[str] = []
            for cell_el in row_el.findall("ss:Cell", ns):
                # Sparse index — pad to target column
                idx_attr = cell_el.get(f"{{{ns_uri}}}Index")
                if idx_attr:
                    target = int(idx_attr) - 1
                    while len(cells) < target:
                        cells.append("")

                data_el = cell_el.find("ss:Data", ns)
                cells.append(data_el.text or "" if data_el is not None else "")

                # MergeAcross — fill merged columns with empty strings
                merge = cell_el.get(f"{{{ns_uri}}}MergeAcross")
                if merge:
                    cells.extend("" for _ in range(int(merge)))

            rows_data.append(cells)

        if not rows_data:
            sheets[name] = pd.DataFrame()
            continue

        max_cols = max(len(r) for r in rows_data)
        for r in rows_data:
            while len(r) < max_cols:
                r.append("")

        sheets[name] = pd.DataFrame(rows_data, dtype=str).fillna("")

    return sheets


class _FakeExcelFile:
    """Minimal pd.ExcelFile-compatible wrapper for pre-parsed DataFrames."""

    def __init__(self, sheets: dict[str, pd.DataFrame]) -> None:
        self.sheet_names = list(sheets.keys())
        self._sheets = sheets

    def parse(self, sheet_name: str, **_kwargs) -> pd.DataFrame:
        return self._sheets[sheet_name]


# ---------------------------------------------------------------------------
# Shared format reader
# ---------------------------------------------------------------------------

def _parse_als_sheet(xl: pd.ExcelFile, sheet_name: str) -> tuple[list[str], dict[int, str], list[tuple], dict[int, dict]]:
    """
    Parse an ALS "Client SOIL" sheet.

    Returns
    -------
    (header_values, sample_cols, data_rows)

    header_values : list of str — the header row (row 13, idx 12)
    sample_cols   : {col_idx: sample_id}
    data_rows     : list of (compound, unit, lor_val, {col_idx: raw_str})
    """
    raw = xl.parse(sheet_name, header=None, dtype=str).fillna("")

    # ── Locate header row (row 13 = idx 12; verify by finding "LOR" / "Parameter") ──
    hdr_row_idx = 12
    for ri in range(8, min(20, len(raw))):
        row_vals = [str(v).strip().upper() for v in raw.iloc[ri]]
        if "LOR" in row_vals or "PARAMETER" in row_vals:
            hdr_row_idx = ri
            break

    hdr = [str(v).strip() for v in raw.iloc[hdr_row_idx]]

    # Find key column indices from header
    lor_col       = next((ci for ci, h in enumerate(hdr) if h.upper() == "LOR"),  3)
    unit_col      = next((ci for ci, h in enumerate(hdr) if h.upper() == "UNIT"), lor_col - 1)
    compound_col  = next((ci for ci, h in enumerate(hdr) if h.upper() in ("PARAMETER", "COMPOUND", "ANALYTE")), 0)
    first_smp_col = lor_col + 1

    # ── Sample IDs: prefer "Client Sample" row, fall back to row 8 (idx 7), then header ──
    sample_cols: dict[int, str] = {}
    _client_row = None
    for ri in range(0, hdr_row_idx):
        _first_cell = next(
            (str(v).strip() for v in raw.iloc[ri] if str(v).strip() and str(v).strip().lower() != "nan"),
            ""
        )
        if "client sample" in _first_cell.lower():
            _client_row = raw.iloc[ri]
            break
    sample_row = _client_row if _client_row is not None else raw.iloc[7]
    for ci in range(first_smp_col, len(sample_row)):
        sid = str(sample_row.iloc[ci]).strip()
        if sid and sid.lower() not in ("nan", ""):
            sample_cols[ci] = sid

    # If neither row yielded anything, fall back to header row sample columns
    if not sample_cols:
        for ci in range(first_smp_col, len(hdr)):
            sid = hdr[ci]
            if sid and sid.lower() not in ("nan", ""):
                sample_cols[ci] = sid

    # ── Borehole / depth from sample IDs like "BH1-1.5m" ──
    _DEPTH_RE = re.compile(r'^(.+?)-(\d+(?:[.,]\d+)?)m$', re.IGNORECASE)
    sample_meta: dict[int, dict] = {}
    for ci, sid in sample_cols.items():
        m = _DEPTH_RE.match(sid)
        if m:
            borehole = m.group(1)
            depth = float(m.group(2).replace(',', '.'))
            sample_meta[ci] = {"borehole": borehole, "depth_from": depth, "depth_to": depth}
        else:
            sample_meta[ci] = {"borehole": None, "depth_from": None, "depth_to": None}

    # ── Data rows ──
    data_rows = []
    for ri in range(hdr_row_idx + 1, len(raw)):
        row = raw.iloc[ri]
        compound = str(row.iloc[compound_col]).strip()
        if not compound or compound.lower() in ("nan", "", "parameter", "compound", "analyte"):
            continue
        # Skip section header rows: "Alcohols / Esters", "PAHs", etc.
        if re.search(r'[A-Za-z].*/', compound) and not re.search(r'\d', compound):
            continue
        # Skip continuation headers: "PAHs - Continued"
        if compound.endswith("- Continued"):
            continue

        unit_val = str(row.iloc[unit_col]).strip() if unit_col < len(row) else "mg/kg"
        if not unit_val or unit_val.lower() == "nan":
            unit_val = "mg/kg"
        if unit_val == "mg/kg" and "DW" not in unit_val:
            continue

        loq: float | None = None
        lor_raw = str(row.iloc[lor_col]).strip().lstrip("<") if lor_col < len(row) else ""
        try:
            loq = float(lor_raw)
        except (ValueError, TypeError):
            pass

        sample_vals: dict[int, str] = {}
        for ci in sample_cols:
            if ci < len(row):
                sample_vals[ci] = str(row.iloc[ci]).strip()

        data_rows.append((compound, unit_val, loq, sample_vals))

    return hdr, sample_cols, data_rows, sample_meta


def _parse_value(raw_val: str, loq: float | None) -> tuple[float | None, str | None]:
    """Parse a raw cell string into (value, flag)."""
    v = raw_val.strip()
    if not v or v.lower() == "nan" or v == "----":
        return None, "ND"
    if v.upper() in ("ND", "N.D.", "N/D", "<LOR", "< LOR", "NOT DETECTED"):
        return loq, "<LOQ"
    if v.startswith("<"):
        try:
            num = float(v[1:])
            return num, "<LOQ"
        except ValueError:
            return loq, "<LOQ"
    try:
        return float(v), None
    except ValueError:
        return loq, "<LOQ"


# ---------------------------------------------------------------------------
# Module-level CAS lookup (safe to call before ALSSoilPDFParser is defined)
# ---------------------------------------------------------------------------

_ALS_CAS_MAP: dict[str, str] = {
    "1,1`-biphenyl":                        "92-52-4",
    "1-chloronaphthalene":                  "90-13-1",
    "2-chloronaphthalene":                  "91-58-7",
    "2-methylnaphthalene":                  "91-57-6",
    "4-bromophenyl phenyl ether":           "101-55-3",
    "4-chlorophenyl phenyl ether":          "7005-72-3",
    "carbazole":                            "86-74-8",
    "acenaphthylene":                       "208-96-8",
    "phenanthrene":                         "85-01-8",
    "benz(a)anthracene":                    "56-55-3",
    "benzo(b)fluoranthene":                 "205-99-2",
    "benzo(k)fluoranthene":                 "207-08-9",
    "benzo(a)pyrene":                       "50-32-8",
    "indeno(1.2.3.cd)pyrene":               "193-39-5",
    "dibenz(a.h)anthracene":                "53-70-3",
    "benzo(g.h.i)perylene":                 "191-24-2",
    "n-nitrosodi-n-propylamine":            "621-64-7",
    "4-chloroaniline":                      "106-47-8",
    "2-nitrophenol":                        "88-75-5",
    "4-nitrophenol":                        "100-02-7",
    "2,4-dinitrotoluene":                   "121-14-2",
    "2,6-dinitrotoluene":                   "606-20-2",
    "2.6-dinitrotoluene":                   "606-20-2",
    "2,4-dinitrophenol":                    "51-28-5",
    "4,6-dinitro-2-methylphenol":           "534-52-1",
    "2-nitroaniline":                       "88-74-4",
    "3-nitroaniline":                       "99-09-2",
    "4-nitroaniline":                       "100-01-6",
    "2-chlorophenol":                       "95-57-8",
    "2,6-dichlorophenol":                   "87-65-0",
    "2.6-dichlorophenol":                   "87-65-0",
    "2.4@2.5-dichlorophenol":               "120-83-2",
    "2,4,6-trichlorophenol":                "88-06-2",
    "2.4.6-trichlorophenol":                "88-06-2",
    "2,4,5-trichlorophenol":                "95-95-4",
    "2.4.5-trichlorophenol":                "95-95-4",
    "2-methylphenol":                       "95-48-7",
    "3- & 4-methylphenol":                  "108-39-4",
    "2,4-dimethylphenol":                   "105-67-9",
    "4-chloro-3-methylphenol":              "59-50-7",
    "dimethyl phthalate":                   "131-11-3",
    "di-n-butyl phthalate":                 "84-74-2",
    "di-n-octyl phthalate":                 "117-84-0",
    "6-caprolactam":                        "105-60-2",
    "bis(2-chloroisopropyl)ether":          "108-60-1",
    "bis(2-chloroisopropyl)ether (all isomers)": "108-60-1",
    "dibenzofuran":                         "132-64-9",
    # TPH fractions
    "c10 - c28 fraction (dro)":             "DRO",
    "c24 - c40 fraction (oro)":             "ORO",
}


def _als_lookup_cas(compound: str) -> str:
    """Look up CAS by compound name using the ALS-specific map."""
    import re as _re
    key = compound.lower().strip()
    if key in _ALS_CAS_MAP:
        return _ALS_CAS_MAP[key]
    key2 = _re.sub(r'\s*\(all\s+isomers\)\s*$', '', key).strip()
    if key2 in _ALS_CAS_MAP:
        return _ALS_CAS_MAP[key2]
    return ""


# ---------------------------------------------------------------------------
# ALSSoilParser
# ---------------------------------------------------------------------------

class ALSSoilParser(BaseParser):
    """ALS laboratory soil report — sheet 'Client SOIL 1'."""

    LAB_NAME = "ALS"

    _METALS = frozenset({
        "LEAD", "ZINC", "COPPER", "NICKEL", "CADMIUM", "CHROMIUM", "ARSENIC",
        "MERCURY", "BARIUM", "SILVER", "MANGANESE", "IRON", "ALUMINUM",
        "ALUMINIUM", "SELENIUM", "ANTIMONY", "BERYLLIUM", "COBALT",
        "MOLYBDENUM", "THALLIUM", "VANADIUM", "COBALT", "TIN",
    })
    _VOC = frozenset({
        "BENZENE", "TOLUENE", "XYLENE", "ETHYLBENZENE", "STYRENE",
        "NAPHTHALENE", "MTBE", "BTEX",
        # Chlorinated / halogenated solvents and related VOCs
        "DICHLOROETHANE", "DICHLOROETHENE", "DICHLOROPROPENE", "DICHLOROPROPANE",
        "TRICHLOROETHANE", "TRICHLOROETHENE", "TETRACHLOROETHANE", "TETRACHLOROETHENE",
        "TETRACHLOROMETHANE", "TRICHLOROMETHANE", "DICHLOROMETHANE", "CHLOROFORM",
        "CHLOROMETHANE", "CHLOROETHANE", "BROMOMETHANE", "BROMOFORM", "DIBROMOMETHANE",
        "TRIHALOMETHANE", "VINYL CHLORIDE", "CHLORINATED ETHENE",
        "DICHLOROBENZENE", "TRICHLOROBENZENE", "CHLOROTOLUENE",
        "BROMOCHLOROMETHANE", "BROMODICHLOROMETHANE", "DIBROMOCHLOROMETHANE",
        "HEXACHLOROBUTADIENE", "FLUOROMETHANE",
        # Grouped / sum parameters
        "SUM OF",
        # Alkylbenzenes and other common VOCs
        "TRIMETHYLBENZENE", "BUTYLBENZENE", "PROPYLBENZENE",
        "ISOPROPYLBENZENE", "ISOPROPYLTOLUENE", "INDANE", "DIOXANE",
    })
    _WATER_ATYPE_MAP = {
        "SOIL_VOC":    "GW_VOC",
        "SOIL_METALS": "GW_METALS",
        "SOIL_SVOC":   "GW_SVOC",
        "SOIL_PFAS":   "GW_PFAS",
        "SOIL_TPH":    "GW_VOC",
    }

    # Compounds to skip entirely when parsing WATER sheets or water PDFs.
    # These are general chemistry / physical parameters and macro-elements that
    # are not environmental contaminants, plus aggregate sum parameters.
    _SKIP_WATER_COMPOUNDS: frozenset[str] = frozenset({
        # General chemistry / physical parameters
        "total organic carbon", "chloride", "fluoride",
        "nitrates", "nitrate as n", "nitrites", "nitrite as n",
        "sulphate as so4 2-",
        # Macro-elements (not contaminants)
        "bismuth", "boron", "calcium", "lithium", "magnesium",
        "phosphorus", "potassium", "silicon", "sodium", "strontium",
        "sulphur", "tellurium", "titanium", "zirconium",
        # Aggregate sum parameters (not individual contaminants)
        "sum of btex", "sum of tex", "sum of xylenes", "sum of btexs",
        "sum of 3 dichlorobenzenes", "sum of 3 trichlorobenzenes",
        "sum of 4 trihalomethanes", "sum of 5 chlorinated ethenes",
        "sum of 1.2-dichloroethenes", "sum of 16 pah", "sum of 17 pah",
        "sum of 7 pcbs",
    })

    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        # Sniff file type
        head = file_obj.read(512)
        file_obj.seek(0)

        if head[:4] == b"%PDF":
            try:
                import pdfplumber as _pl_sniff, io as _io_sniff
            except ImportError:
                pass
            else:
                file_obj.seek(0)
                _pdf_bytes = file_obj.read()
                with _pl_sniff.open(_io_sniff.BytesIO(_pdf_bytes)) as _p:
                    _p1_text = " ".join((_p.pages[i].extract_text() or "") for i in range(min(3, len(_p.pages))))
                if "Sub-Matrix: WATER" in _p1_text:
                    return self._parse_water_pdf(_io_sniff.BytesIO(_pdf_bytes))
                file_obj.seek(0)
            return self._parse_pdf(file_obj)

        # SpreadsheetML: <?xml … urn:schemas-microsoft-com:office:spreadsheet
        if head.lstrip().startswith(b"<?xml"):
            all_bytes = file_obj.read()
            if b"urn:schemas-microsoft-com:office:spreadsheet" in all_bytes:
                xl: pd.ExcelFile | _FakeExcelFile = _FakeExcelFile(
                    _parse_spreadsheetml(all_bytes)
                )
            else:
                file_obj.seek(0)
                xl = pd.ExcelFile(file_obj)
        else:
            xl = pd.ExcelFile(file_obj)

        # Collect all "Client *" sheets (SOIL, PFAS, VOC …) so PFAS-only
        # sheets named "Client PFAS 1" are not silently dropped.
        target_sheets = [s for s in xl.sheet_names
                         if any(tag in s for tag in ("Client SOIL", "Client PFAS",
                                                     "Client VOC", "Client SVOC",
                                                     "Client Metal", "Client WATER"))]
        if not target_sheets:
            # Fallback: accept any single "SOIL" or "PFAS" sheet
            target_sheets = [s for s in xl.sheet_names
                             if "SOIL" in s.upper() or "PFAS" in s.upper()]
        if not target_sheets:
            raise ValueError(f"No recognisable ALS sheet found. Sheets: {xl.sheet_names}")

        records = []
        for sheet in target_sheets:
            try:
                _, sample_cols, data_rows, sample_meta = _parse_als_sheet(xl, sheet)
            except Exception as _e:
                import warnings
                warnings.warn(f"ALS: skipping sheet {sheet!r}: {_e}")
                continue

            if "WATER" in sheet:
                sample_cols = {ci: sid for ci, sid in sample_cols.items()
                               if "blank" not in sid.lower()}

            for compound, unit, loq, sample_vals in data_rows:
                if "WATER" in sheet and compound.strip().lower() in self._SKIP_WATER_COMPOUNDS:
                    continue
                cas   = _als_lookup_cas(compound) or fuzzy_name_to_cas(compound) or ""
                atype = self._analysis_type(compound)
                if "WATER" in sheet:
                    atype = self._WATER_ATYPE_MAP.get(atype, "GW_VOC")
                # µg/kg DW == ng/g numerically — normalise unit label for PFAS
                norm_unit = "ng/g" if atype == "SOIL_PFAS" else unit
                for ci, raw_val in sample_vals.items():
                    if ci not in sample_cols:
                        continue
                    value, flag = _parse_value(raw_val, loq)
                    if value is None and flag is None:
                        flag = "ND"
                    meta = sample_meta.get(ci, {})
                    records.append({
                        "compound":      compound,
                        "cas":           cas,
                        "value":         value,
                        "flag":          flag or "",
                        "unit":          norm_unit,
                        "sample_id":     sample_cols[ci],
                        "lod":           None,
                        "loq":           loq,
                        "analysis_type": atype,
                        "borehole":      meta.get("borehole"),
                        "depth_from":    meta.get("depth_from"),
                        "depth_to":      meta.get("depth_to"),
                    })

        return records

    def _analysis_type(self, compound: str, cas: str = "", method: str = "") -> str:
        m = method.upper()
        c = compound.upper()
        if "METAXHB" in m or "DRY" in m:
            return "SOIL_METALS"
        if "TPHFID" in m:
            return "SOIL_TPH"
        if "VOCGMS" in m or "VOC" in m:
            _PAH = {"naphthalene", "anthracene", "pyrene", "fluorene", "phenanthrene",
                    "chrysene", "fluoranthene", "acenaphthylene", "acenaphthene",
                    "benzo", "indeno", "dibenz", "perylene"}
            if any(p in compound.lower() for p in _PAH):
                return "SOIL_SVOC"
            _SVOC_NAMES = {"1.2-dibromoethane", "dibromoethane"}
            if any(p in compound.lower() for p in _SVOC_NAMES):
                return "SOIL_SVOC"
            return "SOIL_VOC"
        if any(k in c for k in self._METALS):
            return "SOIL_METALS"
        if any(k in c for k in self._VOC):
            return "SOIL_VOC"
        if "TPH" in c or "PETROLEUM" in c or "DRO" in c or "ORO" in c:
            return "SOIL_TPH"
        if "PFAS" in c or "PFOA" in c or "PFOS" in c:
            return "SOIL_PFAS"
        return "SOIL_SVOC"

    def _parse_water_pdf(self, file_obj) -> list[dict]:
        """Parse ALS water PDF (Sub-Matrix: WATER, e.g. PR2021666_0_COA_Standard_CAI.pdf).

        Line format (whitespace-separated):
          <compound> <method-code> <LOR> <unit> <val1> [± <MU%>] <val2> [± <MU%>] …
        Values: <X → <LOQ, ---- → ND at LOR, numeric → detected.
        """
        import pdfplumber, re as _re

        _METHOD_RE = _re.compile(
            r'^(.+?)\s+(W-[A-Z0-9-]+)\s+([\d.]+)\s+(\S+)\s+(.+)$'
        )
        _BLANK_RE = _re.compile(r'\bblank\b', _re.IGNORECASE)

        records: list[dict] = []
        sample_ids: list[str] = []

        file_obj.seek(0)
        with pdfplumber.open(file_obj) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                lines = text.splitlines()

                # Collect sample IDs from "Client sample ID" line on this page
                for line in lines:
                    if "client sample" in line.lower() and "id" in line.lower():
                        tokens = line.split()
                        for j, tok in enumerate(tokens):
                            if tok.lower().rstrip(":") in ("id", "id:") or (
                                tok.lower().startswith("id") and ":" in tok
                            ):
                                candidates = [
                                    t for t in tokens[j + 1:]
                                    if t
                                    and not _BLANK_RE.search(t)
                                    and not _re.match(r'^\d+$', t)
                                    and not (t.startswith("(") and t.endswith(")"))
                                ]
                                if candidates:
                                    sample_ids = candidates
                                break
                        break

                for line in lines:
                    m = _METHOD_RE.match(line.strip())
                    if not m:
                        continue
                    compound = m.group(1).strip()
                    if compound.lower() in self._SKIP_WATER_COMPOUNDS:
                        continue
                    lor      = float(m.group(3))
                    unit     = m.group(4)
                    rest     = m.group(5)

                    # Strip ± uncertainty tokens; keep values
                    tokens = rest.split()
                    values: list[str] = []
                    i = 0
                    while i < len(tokens):
                        t = tokens[i]
                        if t == "±":
                            i += 2  # skip ± and the percentage
                            continue
                        if t.endswith("%") and values:
                            i += 1
                            continue
                        values.append(t)
                        i += 1

                    cas   = _als_lookup_cas(compound) or fuzzy_name_to_cas(compound) or ""
                    atype = self._WATER_ATYPE_MAP.get(self._analysis_type(compound), "GW_VOC")

                    for j, val_str in enumerate(values):
                        sid = sample_ids[j] if j < len(sample_ids) else f"Sample-{j + 1}"
                        rv  = val_str.strip()
                        if not rv or rv in ("----", "---", "--"):
                            value, flag = lor, "ND"
                        elif rv.startswith("<"):
                            try:
                                value = float(rv[1:])
                            except (ValueError, TypeError):
                                value = lor
                            flag = "<LOQ"
                        else:
                            try:
                                value, flag = float(rv), ""
                            except (ValueError, TypeError):
                                continue
                        records.append({
                            "compound":      compound,
                            "cas":           cas,
                            "value":         value,
                            "flag":          flag,
                            "unit":          unit,
                            "sample_id":     sid,
                            "lod":           None,
                            "loq":           lor,
                            "analysis_type": atype,
                        })
        return records

    def _parse_pdf(self, file_obj) -> list[dict]:
        """Parse ALS Czech Republic PDF soil reports."""
        import pdfplumber, re

        records = []
        file_obj.seek(0)

        _SECTION = re.compile(
            r"^(Pesticides|Physical|Halogenated|Aromatic|Polycyclic|"
            r"Chlorinated|Nitrosoamines|Anilines|Nitroaromatic|Chlorophenols|"
            r"Cresols|Phthalates|Aldehydes|Alcohols|Other|Brief|Analytical|"
            r"When sampling|Key:|right solutions|The company|Location|"
            r"The end|The method|The symbol|CZ_SOP|S-[A-Z])",
            re.I
        )
        _DATA_LINE = re.compile(
            r"^(.+?)\s+S-[A-Z]+\d*\s+([\d.]+)\s+(mg/kg\s*DW|µg/kg\s*DW)\s+(.+)$"
        )

        def _parse_results(raw_str: str, sample_ids: list, compound: str, lor: float):
            """Parse space-separated result values and emit records."""
            parts = raw_str.split()
            vals = []
            for p in parts:
                if p == "----":
                    vals.append(None)
                elif p == "*":
                    vals.append(None)
                elif p.startswith("<"):
                    try:
                        vals.append(("<LOQ", float(p[1:])))
                    except ValueError:
                        vals.append(None)
                else:
                    try:
                        vals.append(("", float(p)))
                    except ValueError:
                        vals.append(None)

            cas = (_als_lookup_cas(compound) or fuzzy_name_to_cas(compound) or "")
            atype = self._analysis_type(compound.upper())
            for i, sid in enumerate(sample_ids):
                if i >= len(vals) or vals[i] is None:
                    continue
                flag, val = vals[i] if isinstance(vals[i], tuple) else ("", vals[i])
                records.append({
                    "lab": "ALS", "sample_id": sid,
                    "compound": compound, "cas": cas,
                    "value": val, "flag": flag,
                    "loq": lor, "lod": None,
                    "unit": "mg/kg DW", "analysis_type": atype,
                })

        with pdfplumber.open(file_obj) as pdf:
            sample_ids: list[str] = []
            for page in pdf.pages:
                table = page.extract_table()
                if not table:
                    continue
                for row in table:
                    if not row:
                        continue
                    cells = [str(c).strip() if c else "" for c in row]
                    first = cells[0]

                    # Sample ID row: None None None P10 P4 P9
                    if first == "" and any(re.match(r'^[A-Z]\d+$', c) for c in cells[1:] if c):
                        new_ids = [c for c in cells if c and not re.match(r'^(PR\d+|[\d\-]+)$', c) and c != ""]
                        if new_ids:
                            sample_ids = new_ids
                        continue

                    # Format 1: multi-line cell — "Category\nCompound method LOR unit v1 v2 v3"
                    if first and all(c == "" for c in cells[1:]):
                        for line in first.split("\n"):
                            line = line.strip()
                            if not line or _SECTION.match(line):
                                continue
                            m = _DATA_LINE.match(line)
                            if m:
                                compound = m.group(1).strip()
                                try:
                                    lor = float(m.group(2))
                                except ValueError:
                                    lor = None
                                _parse_results(m.group(4), sample_ids, compound, lor)
                        continue

                    # Format 2: compound+method | LOR | unit | v1 | v2 | v3
                    if len(cells) >= 5 and cells[1] and cells[2]:
                        compound_method = first
                        # Strip method suffix (S-XXXXX)
                        compound = re.sub(r'\s+S-[A-Z0-9]+$', '', compound_method).strip()
                        if not compound or _SECTION.match(compound):
                            continue
                        try:
                            lor = float(cells[1])
                        except (ValueError, TypeError):
                            continue
                        if "mg/kg" not in cells[2] and "µg" not in cells[2]:
                            continue
                        _parse_results(" ".join(c for c in cells[3:] if c), sample_ids, compound, lor)

        return records


# ---------------------------------------------------------------------------
# ALSGrainSizeParser
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# ALSSoilPDFParser
# ---------------------------------------------------------------------------

class ALSSoilPDFParser(BaseParser):
    """ALS Czech Republic soil report — PDF 'Certificate of Analysis' format.

    Table layout per page:
      • "Client sample ID" row: sample names in columns 4+
      • Data rows: Parameter | Method | LOR | Unit | Result × N samples
      • Category header rows (Method/LOR/Unit cells all empty): section label
      • Values: <N.NN → flag=<LOQ value=N.NN;  ---- or * → skip (not analysed)
    """

    LAB_NAME = "ALS"

    # ALS uses compound names that differ from VSL table — map them here
    _ALS_CAS_MAP: dict[str, str] = {
        "1,1`-biphenyl":                        "92-52-4",
        "1-chloronaphthalene":                  "90-13-1",
        "2-chloronaphthalene":                  "91-58-7",
        "2-methylnaphthalene":                  "91-57-6",
        "4-bromophenyl phenyl ether":           "101-55-3",
        "4-chlorophenyl phenyl ether":          "7005-72-3",
        "carbazole":                            "86-74-8",
        "acenaphthylene":                       "208-96-8",
        "phenanthrene":                         "85-01-8",
        "benz(a)anthracene":                    "56-55-3",
        "benzo(b)fluoranthene":                 "205-99-2",
        "benzo(k)fluoranthene":                 "207-08-9",
        "benzo(a)pyrene":                       "50-32-8",
        "indeno(1.2.3.cd)pyrene":               "193-39-5",
        "dibenz(a.h)anthracene":                "53-70-3",
        "benzo(g.h.i)perylene":                 "191-24-2",
        "n-nitrosodi-n-propylamine":            "621-64-7",
        "4-chloroaniline":                      "106-47-8",
        "2-nitrophenol":                        "88-75-5",
        "4-nitrophenol":                        "100-02-7",
        "2,4-dinitrotoluene":                   "121-14-2",
        "2,6-dinitrotoluene":                   "606-20-2",
        "2.6-dinitrotoluene":                   "606-20-2",
        "2,4-dinitrophenol":                    "51-28-5",
        "4,6-dinitro-2-methylphenol":           "534-52-1",
        "2-nitroaniline":                       "88-74-4",
        "3-nitroaniline":                       "99-09-2",
        "4-nitroaniline":                       "100-01-6",
        "2-chlorophenol":                       "95-57-8",
        "2,6-dichlorophenol":                   "87-65-0",
        "2.6-dichlorophenol":                   "87-65-0",
        "2.4@2.5-dichlorophenol":               "120-83-2",
        "2,4,6-trichlorophenol":                "88-06-2",
        "2.4.6-trichlorophenol":                "88-06-2",
        "2,4,5-trichlorophenol":                "95-95-4",
        "2.4.5-trichlorophenol":                "95-95-4",
        "2-methylphenol":                       "95-48-7",
        "3- & 4-methylphenol":                  "108-39-4",
        "2,4-dimethylphenol":                   "105-67-9",
        "4-chloro-3-methylphenol":              "59-50-7",
        "dimethyl phthalate":                   "131-11-3",
        "di-n-butyl phthalate":                 "84-74-2",
        "di-n-octyl phthalate":                 "117-84-0",
        "6-caprolactam":                        "105-60-2",
        "bis(2-chloroisopropyl)ether":          "108-60-1",
        "bis(2-chloroisopropyl)ether (all isomers)": "108-60-1",
        "dibenzofuran":                         "132-64-9",
        # TPH fractions
        "c10 - c28 fraction (dro)":             "DRO",
        "c24 - c40 fraction (oro)":             "ORO",
    }

    @classmethod
    def _normalize_compound(cls, name: str) -> str:
        """Clean up compound name: strip method suffix, fix newlines."""
        import re
        # Handle case: "Compound (all S-METHOD\nisomers)" → "Compound (all isomers)"
        name = re.sub(r'\s+S-[A-Z0-9]+\s*\n\s*', ' ', name)
        # Remove trailing method suffix
        name = re.sub(r'\s+S-[A-Z0-9]+(\s+.*)?$', '', name).strip()
        # Remove any remaining newlines
        name = name.replace('\n', ' ').strip()
        return name

    @classmethod
    def _lookup_cas(cls, compound: str) -> str:
        """Look up CAS by compound name, trying ALS-specific map first."""
        key = compound.lower().strip()
        # ALS-specific map
        if key in cls._ALS_CAS_MAP:
            return cls._ALS_CAS_MAP[key]
        # Strip parenthetical isomers suffix and retry
        import re
        key2 = re.sub(r'\s*\(all\s+isomers\)\s*$', '', key).strip()
        if key2 in cls._ALS_CAS_MAP:
            return cls._ALS_CAS_MAP[key2]
        return ""

    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        try:
            import pdfplumber
        except ImportError as exc:
            raise ImportError("pdfplumber is required: pip install pdfplumber") from exc

        # Build CAS lookup via ThresholdManager
        try:
            from core.threshold_manager import ThresholdManager
            import os
            _thresh_dir = os.path.join(os.path.dirname(__file__), '..', '..', 'thresholds')
            _tm = ThresholdManager(
                os.path.join(_thresh_dir, 'soil_vsl_tier1_v7_2024.xlsx'),
                vsl_full_path=os.path.join(_thresh_dir, 'soil_vsl_v7_full.xlsx')
            )
            self._cas_lookup = _tm.get_cas_by_name
        except Exception:
            self._cas_lookup = lambda x: fuzzy_name_to_cas(x) or ""

        file_obj.seek(0)
        records: list[dict] = []
        sample_cols: list[str] = []
        current_atype = "SOIL_SVOC"

        with pdfplumber.open(file_obj) as pdf:
            for page in pdf.pages:
                page_recs, sample_cols, current_atype = self._parse_page(
                    page, sample_cols, current_atype
                )
                records.extend(page_recs)

        return records

    # Regex for concatenated single-cell data rows:
    # e.g. "Dinoseb S-SMVGMS03 0.50 mg/kg DW <0.50 <0.50 <0.50"
    # Group 1: compound, 2: LOR, 3: unit (mg/kg DW or %), 4: space-separated results
    _CONCAT_ROW_RE = re.compile(
        r'^(.+?)\s+S-\w+\d*\s+([\d.]+)\s+(mg/kg\s*DW|%)\s+(.+)$'
    )
    # Regex to strip trailing method code from compound names (e.g. "Naphthalene S-SMVGMS03")
    _METHOD_SUFFIX_RE = re.compile(r'\s+S-[A-Z0-9]+$')
    # Sample ID pattern: letter followed by digits (e.g. P10, K3)
    _SAMPLE_ID_RE = re.compile(r'^[A-Z]\d+$')

    def _parse_page(
        self,
        page,
        carry_sample_cols: list[str],
        carry_atype: str,
    ) -> tuple[list[dict], list[str], str]:
        sample_cols   = carry_sample_cols[:]
        current_atype = carry_atype
        records: list[dict] = []

        tables = page.extract_tables()
        if not tables:
            return records, sample_cols, current_atype

        for table in tables:
            for row in table:
                if not row:
                    continue
                cells = [str(c or "").strip() for c in row]
                if not any(cells):
                    continue

                first    = re.sub(r'\s+S-[A-Z0-9]+$', '', cells[0]).strip()
                first_lo = first.lower()

                # PDF merges section header + first data row into one cell with \n
                # Split them and process header first, then continue with data row
                if "\n" in first and all(
                    not cells[i] or cells[i] in ("-", "nan")
                    for i in range(1, min(4, len(cells)))
                ):
                    header_part, data_part = first.split("\n", 1)
                    current_atype = self._section_to_atype(header_part)
                    # Reconstruct cells with just the data part
                    cells[0] = data_part.strip()
                    first = cells[0]
                    first_lo = first.lower()

                # ── Sample ID header row ──────────────────────────────────
                # Format: cells[0] empty, sample names start at cells[3]
                if (not first and
                        len(cells) > 3 and
                        self._SAMPLE_ID_RE.match(cells[3])):
                    new = [c for c in cells[3:] if c and self._SAMPLE_ID_RE.match(c)]
                    if new:
                        sample_cols = new
                    continue
                # Fallback: explicit "Client sample ID" label
                if "client sample id" in first_lo:
                    new = [c for c in cells[4:] if c and c.lower() not in ("nan", "")]
                    if new:
                        sample_cols = new
                    continue

                # ── Standard column header row — skip ─────────────────────
                if first_lo in ("parameter", "analyte", "compound") or \
                   first_lo.startswith("parameter method"):
                    continue

                # ── Single-cell concatenated row ──────────────────────────
                # All content landed in cells[0]; parse with regex.
                if len([c for c in cells if c]) == 1 and first:
                    for line in first.splitlines():
                        line = line.strip()
                        m = self._CONCAT_ROW_RE.match(line)
                        if not m:
                            continue
                        compound = self._normalize_compound(m.group(1))
                        loq_str   = m.group(2)
                        row_unit  = m.group(3).strip()
                        results   = m.group(4).split()
                        loq: float | None = None
                        try:
                            loq = float(loq_str)
                        except ValueError:
                            pass
                        cas = (_als_lookup_cas(compound) or self._cas_lookup(compound) or fuzzy_name_to_cas(compound) or "")
                        for i, raw_val in enumerate(results):
                            if i >= len(sample_cols):
                                break
                            value, flag = self._parse_pdf_value(raw_val, loq)
                            if value is None and flag is None:
                                continue
                            records.append({
                                "compound":      compound,
                                "cas":           cas,
                                "value":         value,
                                "flag":          flag or "",
                                "unit":          row_unit,
                                "sample_id":     sample_cols[i],
                                "lod":           None,
                                "loq":           loq,
                                "analysis_type": current_atype,
                            })
                    continue

                # ── Empty first cell (not sample-ID row) — skip ───────────
                if not first:
                    continue

                # ── Section header: Method/LOR/Unit all empty ─────────────
                if not any(len(cells) > i and cells[i] for i in (1, 2, 3)):
                    current_atype = self._section_to_atype(first)
                    continue

                # ── Normal multi-cell data row ────────────────────────────
                lor_raw  = cells[1] if len(cells) > 1 else ""
                unit     = cells[2] if len(cells) > 2 else "mg/kg"
                if not unit or unit.lower() == "nan":
                    unit = "mg/kg"
                compound = self._normalize_compound(first)

                loq = None
                try:
                    loq = float(lor_raw.lstrip("<"))
                except (ValueError, TypeError):
                    pass

                cas = (_als_lookup_cas(compound) or self._cas_lookup(compound) or fuzzy_name_to_cas(compound) or "")

                for i, raw_val in enumerate(cells[3:]):
                    if i >= len(sample_cols):
                        break
                    sample_id = sample_cols[i]
                    if not sample_id:
                        continue
                    value, flag = self._parse_pdf_value(raw_val, loq)
                    if value is None and flag is None:
                        continue
                    records.append({
                        "compound":      compound,
                        "cas":           cas,
                        "value":         value,
                        "flag":          flag or "",
                        "unit":          unit,
                        "sample_id":     sample_id,
                        "lod":           None,
                        "loq":           loq,
                        "analysis_type": current_atype,
                    })

        return records, sample_cols, current_atype

    @staticmethod
    def _parse_pdf_value(raw: str, loq: float | None) -> tuple[float | None, str | None]:
        v = raw.strip()
        if not v or v in ("----", "*", "-", "n/a", "n.a.", "N/A"):
            return None, None  # not analysed — skip record
        if v.upper() in ("ND", "N.D.", "N/D", "<LOR", "< LOR", "NOT DETECTED"):
            return loq, "<LOQ"
        if v.startswith("<"):
            try:
                return float(v[1:]), "<LOQ"
            except ValueError:
                return loq, "<LOQ"
        try:
            return float(v), None
        except ValueError:
            return loq, "<LOQ"

    _SECTION_MAP = [
        (["extractable metals", "major cations", "physical parameter"], "SOIL_METALS"),
        (["total petroleum", "petroleum hydrocarbon", "c10 -", "c24 -"], "SOIL_TPH"),
        (["btex", "halogenated volatile", "non-halogenated volatile", "volatile organic"], "SOIL_VOC"),
        (["polycyclic", "pah"], "SOIL_SVOC"),
    ]

    def _section_to_atype(self, header: str) -> str:
        lo = header.lower()
        for keywords, atype in self._SECTION_MAP:
            if any(kw in lo for kw in keywords):
                return atype
        return "SOIL_SVOC"


# ---------------------------------------------------------------------------
# ALSGrainSizeParser
# ---------------------------------------------------------------------------

class ALSGrainSizeParser(BaseParser):
    """ALS grain-size (sieve analysis) — sheet 'Client SOIL 1' with Fraction parameters."""

    LAB_NAME = "ALS"

    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        xl = pd.ExcelFile(file_obj)
        sheet = next((s for s in xl.sheet_names if "Client SOIL" in s), None)
        if sheet is None:
            raise ValueError(f"No 'Client SOIL' sheet found. Sheets: {xl.sheet_names}")

        _, sample_cols, data_rows, sample_meta = _parse_als_sheet(xl, sheet)
        records = []

        for compound, unit, loq, sample_vals in data_rows:
            cas = ""
            for ci, raw_val in sample_vals.items():
                value, flag = _parse_value(raw_val, loq)
                if value is None and flag is None:
                    continue
                meta = sample_meta.get(ci, {})
                records.append({
                    "compound":      compound,
                    "cas":           cas,
                    "value":         value,
                    "flag":          flag or "",
                    "unit":          unit or "%",
                    "sample_id":     sample_cols[ci],
                    "lod":           None,
                    "loq":           loq,
                    "analysis_type": "SOIL_GRAIN_SIZE",
                    "borehole":      meta.get("borehole"),
                    "depth_from":    meta.get("depth_from"),
                    "depth_to":      meta.get("depth_to"),
                })

        return records
