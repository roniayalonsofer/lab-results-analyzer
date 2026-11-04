"""
parsers/groundwater/als_gw.py
-------------------------------
Parser for ALS laboratory groundwater (Certificate of Analysis) reports.

Format: SpreadsheetML (.xls/.xlsx exported as Excel 2003 XML), sheet "Client WATER - 1"
  row 6:  "Client Sample ID"     | PP-2 | PP-1 | PP-5 | ...
  row 7:  "Laboratory Sample ID" | PR2607572001 | ...
  row 8:  "Client Sampling Date" | 07/01/2026 | ...
  row 11: "Parameter | Method | Unit | LOR | <sample value columns>"
  row 13+: category header rows (e.g. "BTEX", "Halogenated Volatile Organic
           Compounds") followed by compound rows with values per sample.

  Units are µg/L; converted to mg/L for cross-lab comparability with
  Bactochem groundwater reports (which report BTEX/MTBE in mg/L).

Also read:
  - Real .xlsx export (EXCEL_GENERIC.xlsx): the same layout, with the label cells
    of the sample rows in the LOR column. Sheets "Client WATER 1", "Client WATER 2"...
  - ALS Italia "Analysis Report" PDF (ZP...): one sample per report, named in
    "Sample Description : <project> - <sample>"; rows are
    "<compound> <method> <start date> <LOQ> <result> <MU> <unit> ...".
    Next to another file of the same upload it adds its results to that file's samples.
"""

from __future__ import annotations

import io
import re
import xml.etree.ElementTree as _ET

import pandas as pd

from parsers.base import BaseParser

_SML_NS = "urn:schemas-microsoft-com:office:spreadsheet"


def _parse_spreadsheetml(data: bytes) -> dict[str, pd.DataFrame]:
    """Parse SpreadsheetML XML into {sheet_name: DataFrame}, all-string dtype."""
    root = _ET.fromstring(data)
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
                idx_attr = cell_el.get(f"{{{ns_uri}}}Index")
                if idx_attr:
                    target = int(idx_attr) - 1
                    while len(cells) < target:
                        cells.append("")
                data_el = cell_el.find("ss:Data", ns)
                cells.append(data_el.text or "" if data_el is not None else "")
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


def _parse_value(raw: str, lor: float | None) -> tuple[float | None, str | None]:
    v = (raw or "").strip()
    if not v or v in ("----", "nan", "NaN", "-"):
        return None, "dash"
    if v.startswith("<"):
        try:
            return float(v[1:]), "<LOQ"
        except ValueError:
            return lor, "<LOQ"
    try:
        return float(v), None
    except ValueError:
        return None, "dash"


def _normalize_date(s: str) -> str:
    """'07/01/2026' (DD/MM/YYYY) -> 'dd.mm.yy'."""
    s = s.strip()
    m = re.match(r'(\d{1,2})/(\d{1,2})/(\d{4})', s)
    if m:
        d, mo, y = m.groups()
        return f"{d.zfill(2)}.{mo.zfill(2)}.{y[2:]}"
    return s


# CAS lookup for common GW BTEX/VOC compounds
_GW_CAS_MAP = {
    "benzene": "71-43-2",
    "toluene": "108-88-3",
    "ethylbenzene": "100-41-4",
    "meta- & para-xylene": "1330-20-7",
    "ortho-xylene": "95-47-6",
    "methyl tert-butyl ether (mtbe)": "1634-04-4",
    "naphthalene": "91-20-3",
}

# Field/physical parameters → GW_FIELD_PARAMS
_FIELD_PARAM_NAMES = frozenset({
    "ELECTRICAL CONDUCTIVITY @ 25°C", "PH VALUE", "REDOX POTENTIAL",
    "DISSOLVED OXYGEN", "OXYGEN SATURATION",
})


def _normalize_unit(unit: str) -> str:
    """pdfplumber drops the µ of some ALS certificates, leaving "g/L"."""
    return "µg/L" if unit.strip().lower() in ("g/l", "ug/l", "µg/l", "μg/l") else unit.strip()


def _analysis_type(compound: str, method: str, unit: str, category: str) -> str:
    """Which report sheet a result belongs to, by the certificate's own section."""
    cat = category.lower()
    if compound.upper() in _FIELD_PARAM_NAMES:
        return "GW_FIELD_PARAMS"
    if "perfluor" in cat or unit.lower() == "ng/l":
        return "GW_PFAS"
    # Before metals: the section is called "Nonmetallic Inorganic Parameters"
    if "inorganic" in cat or "agregate" in cat or "aggregate" in cat or cat == "other":
        return "GW_CHEM"
    if "metal" in cat:
        return "GW_METALS"
    if any(k in cat for k in ("pesticide", "explosive", "pharmaceutical")):
        return "GW_SVOC"
    if "polycyclic" in cat and "VOC" not in method.upper():
        return "GW_SVOC"
    return "GW_VOC"


def _record(compound: str, method: str, unit: str, lor: float | None, raw: str,
            sample_id: str, sampling_date: str, category: str) -> dict:
    """One result; µg/L becomes mg/L, the groundwater sheets' common unit."""
    unit = _normalize_unit(unit) or "µg/L"
    value, flag = _parse_value(raw, lor)
    atype = _analysis_type(compound, method, unit, category)
    if unit == "µg/L":
        if value is not None:
            value = round(value / 1000.0, 9)
        if lor is not None:
            lor = round(lor / 1000.0, 9)
        unit = "mg/L"
    return {
        "compound":      compound,
        "cas":           _GW_CAS_MAP.get(compound.lower().strip(), ""),
        "value":         value,
        "flag":          flag or "",
        "unit":          unit,
        "sample_id":     sample_id,
        "lod":           None,
        "loq":           lor,
        "analysis_type": atype,
        "sampling_date": sampling_date,
        "category":      category,
    }


def _compound_key(name: str) -> str:
    """1.4-Dioxane and 1,4-Dioxane are one compound."""
    return re.sub(r"(?<=\d)\.(?=\d)", ",", name.strip().lower())


def _result_rank(r: dict) -> tuple:
    """Lower is better: a detection, then the lower limit, which says more."""
    return (r["flag"] != "", r["loq"] if r["loq"] is not None else float("inf"))


def _best_results(records: list[dict]) -> list[dict]:
    """One result per compound in a sample and sheet, since a sheet has one cell for it.

    The certificate can report a compound by two methods, or two labs can test it.
    """
    best: dict[tuple, dict] = {}
    for r in records:
        key = (r["sample_id"], r["analysis_type"], _compound_key(r["compound"]))
        current = best.get(key)
        if current is None:
            best[key] = r
        elif _result_rank(r) < _result_rank(current):
            # The other report may lack the sampling date
            best[key] = {**r, "sampling_date": r["sampling_date"] or current["sampling_date"]}
    return list(best.values())


def _is_italy_report(data: bytes) -> bool:
    """ALS Italia "Analysis Report" PDF, as opposed to the Czech "Certificate of Analysis"."""
    import pdfplumber
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            first = (pdf.pages[0].extract_text() or "") if pdf.pages else ""
    except Exception:
        return False
    return first.lstrip().startswith("Analysis Report") and "Sample Description" in first


# Lines that end a block of results in the Czech certificate
_PDF_BLOCK_END = ("Sub-Matrix:", "Issue Date", "When sampling date", "Key:", "Brief Method",
                  "The end of", "right solutions", "Analytical Methods")

_PDF_SECTION_RE = re.compile(
    r"(Compounds|Parameters|Pesticides|Hydrocarbons|Cations|Explosives|BTEX|\(PAHs\)|Other)"
    r"( - .*| Continued)?( - FTIR)?$")

_ITALY_ROW_RE = re.compile(
    r"^(?P<compound>.+?)\s+(?P<method>[A-Z]-\S+)\s+\d{2}/\d{2}/\d{4}\s+"
    r"(?P<loq>[\d.]+)\s+(?P<result><?[\d.]+)\s+\S+\s+(?P<unit>\S+/\S+)")


class ALSGroundwaterParser(BaseParser):
    """ALS laboratory groundwater Certificate of Analysis — 'Client WATER' sheet."""

    LAB_NAME = "ALS"
    ANALYSIS_TYPES = ["GW_VOC", "GW_SVOC", "GW_METALS", "GW_PFAS", "GW_CHEM", "GW_FIELD_PARAMS"]

    def _parse_pdf(self, data: bytes) -> list[dict]:
        """
        Parse ALS Czech Republic 'CERTIFICATE OF ANALYSIS' PDF (groundwater).

        Layout (per Sub-Matrix block, may span multiple pages):
          'Sub-Matrix: WATER Client sample ID <id1> <id2> <id3> ...'
          'Laboratory sample ID <lab_id1> <lab_id2> ...'
          'Client sampling date / time <date1> <date2> ...'
          'Parameter Method LOR Unit Result Result Result'
          <category line, e.g. 'BTEX', 'Physical Parameters'>
          '<Compound> <Method> <LOR> <Unit> <result1> <result2> <result3>'
          ...
        Multiple Sub-Matrix blocks may appear (one per group of ~3 samples);
        blank/'----' sample columns are skipped.
        """
        import pdfplumber

        full_text = ""
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for page in pdf.pages:
                full_text += (page.extract_text() or "") + "\n"

        lines = full_text.split("\n")
        records: list[dict] = []
        current_category = ""  # A section continues on the next page

        i = 0
        n = len(lines)
        while i < n:
            line = lines[i].strip()
            if line.startswith("Sub-Matrix:") and "Client sample ID" in line:
                # Extract sample IDs after "Client sample ID"
                after = line.split("Client sample ID", 1)[1].strip()
                lab_ids = next((lines[x].strip().split("Laboratory sample ID", 1)[1].split()
                                for x in range(i + 1, min(i + 4, n))
                                if "Laboratory sample ID" in lines[x]), [])
                # Keep "----" as an explicit placeholder so n_samples matches
                # the actual number of result columns in the data rows.
                sample_ids = self._split_sample_ids(after, lab_ids)

                sampling_dates: list[str] = []
                hdr_idx = None
                j = i + 1
                while j < min(i + 6, n):
                    jl = lines[j].strip()
                    if jl.startswith("Client sampling date"):
                        after_d = jl.split("/ time", 1)[-1].strip() if "/ time" in jl else \
                                  jl.split("date", 1)[-1].strip()
                        sampling_dates = self._extract_dates(after_d)
                    if jl.startswith("Parameter") and "Method" in jl and "LOR" in jl:
                        hdr_idx = j
                        break
                    j += 1

                if hdr_idx is None or not any(sid != "----" for sid in sample_ids):
                    i += 1
                    continue

                n_samples = len(sample_ids)
                k = hdr_idx + 1
                last_row_size = 0
                while k < n:
                    kl = lines[k].strip()
                    if not kl:
                        k += 1
                        continue
                    if kl.startswith(_PDF_BLOCK_END):
                        break

                    parsed = self._parse_pdf_data_line(kl, n_samples)
                    if parsed is None:
                        if kl.startswith("Parameter "):
                            pass
                        # A section header (e.g. "BTEX", "Pesticides - Continued")
                        elif _PDF_SECTION_RE.search(kl) and "----" not in kl:
                            current_category = kl.split(" - Continued")[0].strip()
                        # A long name wraps below its results: "... sulfon" / "ic acid (PFEESA)"
                        elif "----" not in kl:
                            for r in records[-last_row_size:] if last_row_size else []:
                                joint = "" if r["compound"].endswith("-") else " "
                                r["compound"] = f'{r["compound"]}{joint}{kl}'
                        k += 1
                        continue

                    compound, method, lor_val, unit_val, results = parsed
                    last_row_size = 0
                    for si, sid in enumerate(sample_ids):
                        if si >= len(results) or sid == "----":
                            continue
                        records.append(_record(
                            compound, method, unit_val, lor_val, results[si], sid,
                            sampling_dates[si] if si < len(sampling_dates) else "",
                            current_category))
                        last_row_size += 1
                    k += 1
                i = k
                continue
            i += 1

        return records

    @staticmethod
    def _merge_eb_tokens(tokens: list[str]) -> list[str]:
        """Merge '(EB)' style suffix tokens into the preceding sample ID."""
        out: list[str] = []
        for t in tokens:
            if t.startswith("(") and out:
                out[-1] = f"{out[-1]} {t}"
            else:
                out.append(t)
        return out

    @classmethod
    def _split_sample_ids(cls, text: str, lab_ids: list[str]) -> list[str]:
        """Sample names per result column. A name may contain spaces, so the
        Laboratory sample ID row, which never does, gives the number of columns."""
        tokens = cls._merge_eb_tokens(text.split())
        n_cols = len(lab_ids)
        if not n_cols or len(tokens) == n_cols:
            return tokens
        real = [c for c, lab_id in enumerate(lab_ids) if lab_id != "----"]
        out = ["----"] * n_cols
        if len(real) == 1:
            out[real[0]] = " ".join(t for t in tokens if t != "----")
        else:
            # Several names with spaces can't be told apart: the lab's ID is unambiguous
            for c in real:
                out[c] = lab_ids[c]
        return out

    @staticmethod
    def _extract_dates(text: str) -> list[str]:
        """Extract DD-Mon-YYYY dates from a text fragment, normalized to dd.mm.yy."""
        import datetime as _dt
        found = re.findall(r'\d{1,2}-[A-Za-z]{3}-\d{4}', text)
        out = []
        for f in found:
            try:
                d = _dt.datetime.strptime(f, "%d-%b-%Y")
                out.append(d.strftime("%d.%m.%y"))
            except ValueError:
                out.append(f)
        return out

    @staticmethod
    def _parse_pdf_data_line(line: str, n_samples: int):
        """
        Parse a single PDF data row:
          '<Compound text> <Method> <LOR> <Unit> <result1> <result2> ... <resultN>'
        Returns (compound, method, lor_value, unit, [results]) or None if not a data row.
        """
        # Results are the last n_samples whitespace-separated tokens that look like
        # numbers, '<N', or '----'
        tokens = line.split()
        if len(tokens) < 4 + n_samples:
            return None
        result_tokens = tokens[-n_samples:]
        if not all(re.match(r'^(<?\d|----|\*)', t) for t in result_tokens):
            return None

        remaining = tokens[:-n_samples]
        if len(remaining) < 3:
            return None
        unit_val = remaining[-1]
        lor_raw  = remaining[-2].lstrip("<")
        method   = remaining[-3]
        compound = " ".join(remaining[:-3]).strip()
        if not compound:
            return None
        try:
            lor_val = float(lor_raw)
        except ValueError:
            lor_val = None
        return compound, method, lor_val, unit_val, result_tokens

    def _parse_italy_pdf(self, data: bytes) -> list[dict]:
        """ALS Italia "Analysis Report": the sample's name is the part after the project."""
        import pdfplumber

        with pdfplumber.open(io.BytesIO(data)) as pdf:
            lines = [ln.strip() for p in pdf.pages
                     for ln in (p.extract_text() or "").split("\n")]

        sample_id = ""
        for ln in lines:
            m = re.match(r"Sample Description\s*:\s*(.+?)(?:\s+Client:)?$", ln)
            if m:
                sample_id = m.group(1).rsplit(" - ", 1)[-1].strip()
                break

        records: list[dict] = []
        category = ""
        for ln in lines:
            m = _ITALY_ROW_RE.match(ln)
            if m:
                records.append(_record(m["compound"], m["method"], m["unit"], float(m["loq"]),
                                       m["result"], sample_id, "", category))
            elif ln.endswith(("Compounds", "Parameters", "Pesticides", "Metals")):
                category = ln
        return records

    def _parse_sheets(self, sheets: dict[str, pd.DataFrame]) -> list[dict]:
        """Every "Client WATER" sheet: a large job spreads its samples over several."""
        names = ([n for n in sheets if "client water" in n.lower()]
                 or [n for n in sheets if "water" in n.lower()][:1])
        records: list[dict] = []
        for name in names:
            records.extend(self._parse_water_sheet(sheets[name]))
        return records

    def parse(self, file_obj, pdf_bytes: list[bytes] | None = None) -> list[dict]:
        data = file_obj.read() if hasattr(file_obj, "read") else file_obj
        if isinstance(data, str):
            data = data.encode("utf-8")

        other_pdfs = [b for b in (pdf_bytes or []) if b != data]
        italy_reports = [b for b in other_pdfs if _is_italy_report(b)]

        if data[:4] == b"%PDF":
            if _is_italy_report(data):
                # A certificate in the same upload collects the Italian results (below),
                # so the same sample doesn't become two columns.
                if len(italy_reports) < len(other_pdfs):
                    return []
                records = self._parse_italy_pdf(data)
            else:
                records = self._parse_pdf(data)
        elif data[:2] == b"PK":
            sheets = pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, dtype=str)
            records = self._parse_sheets({n: df.fillna("") for n, df in sheets.items()})
        else:
            records = self._parse_sheets(_parse_spreadsheetml(data))

        # A Czech certificate next to the data file is a copy of it, but an Italian
        # report holds results found nowhere else.
        for report in italy_reports:
            records.extend(self._parse_italy_pdf(report))
        records = _best_results(records)
        # The chemistry sheet is in mg/L; its alkalinity rows are in mmol/L
        for r in records:
            if r["analysis_type"] == "GW_CHEM" and r["unit"] != "mg/L":
                r["compound"] = f'{r["compound"]} [{r["unit"]}]'
        return records

    def _parse_water_sheet(self, df: pd.DataFrame) -> list[dict]:
        n_rows = len(df)
        records: list[dict] = []

        # ── Find key rows: Client Sample ID, Laboratory Sample ID, Sampling Date, header ──
        sample_id_row = lab_id_row = date_row = hdr_row_idx = None
        for i in range(min(20, n_rows)):
            first = str(df.iloc[i, 0] if df.shape[1] > 0 else "").strip()
            # These label cells may live in any of the first few columns
            row_vals = [str(v).strip() for v in df.iloc[i]]
            joined = " ".join(row_vals).lower()
            if "client sample id" in joined and sample_id_row is None:
                sample_id_row = i
            if "laboratory sample id" in joined and lab_id_row is None:
                lab_id_row = i
            if "sampling date" in joined and date_row is None:
                date_row = i
            if first.lower() == "parameter" and hdr_row_idx is None:
                hdr_row_idx = i

        if hdr_row_idx is None or sample_id_row is None:
            return []

        hdr = [str(v).strip() for v in df.iloc[hdr_row_idx]]
        method_col = next((ci for ci, h in enumerate(hdr) if h.lower() == "method"), 1)
        unit_col   = next((ci for ci, h in enumerate(hdr) if h.lower() == "unit"), 2)
        lor_col    = next((ci for ci, h in enumerate(hdr) if h.upper() == "LOR"), 3)
        first_smp_col = lor_col + 1

        # Sample IDs + dates (skip the label cell itself)
        sample_row = [str(v).strip() for v in df.iloc[sample_id_row]]
        date_vals  = [str(v).strip() for v in df.iloc[date_row]] if date_row is not None else []

        sample_cols: dict[int, str] = {}
        for ci in range(first_smp_col, len(sample_row)):
            sid = sample_row[ci]
            if sid and "client sample" not in sid.lower():
                sample_cols[ci] = sid

        sampling_dates: dict[int, str] = {}
        for ci in sample_cols:
            if ci < len(date_vals) and date_vals[ci]:
                sampling_dates[ci] = _normalize_date(date_vals[ci])

        # ── Data rows: iterate from header+1 to end, tracking category ──
        current_category = ""
        for ri in range(hdr_row_idx + 1, n_rows):
            row = [str(v).strip() for v in df.iloc[ri]]
            compound = row[0] if row else ""
            if not compound:
                continue

            method_val = row[method_col] if method_col < len(row) else ""
            unit_val   = row[unit_col] if unit_col < len(row) else ""

            # Category header row: has text in col0 but nothing in method/unit
            if compound and not method_val and not unit_val:
                current_category = compound
                continue

            if compound.lower() in ("parameter",):
                continue

            lor_raw = row[lor_col].lstrip("<") if lor_col < len(row) else ""
            try:
                lor_val = float(lor_raw)
            except (ValueError, TypeError):
                lor_val = None

            for ci, sid in sample_cols.items():
                if ci >= len(row):
                    continue
                records.append(_record(compound, method_val, unit_val, lor_val, row[ci], sid,
                                       sampling_dates.get(ci, ""), current_category))

        return records
