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


class ALSGroundwaterParser(BaseParser):
    """ALS laboratory groundwater Certificate of Analysis — 'Client WATER' sheet."""

    LAB_NAME = "ALS"
    ANALYSIS_TYPES = ["GW_VOC", "GW_FIELD_PARAMS"]

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

        i = 0
        n = len(lines)
        while i < n:
            line = lines[i].strip()
            if line.startswith("Sub-Matrix:") and "Client sample ID" in line:
                # Extract sample IDs after "Client sample ID"
                after = line.split("Client sample ID", 1)[1].strip()
                sample_ids = [s for s in after.split() if s and s != "----"]
                # Sample IDs may include things like "PP-1" or "PP-1" + "(EB)" split by space
                sample_ids = self._merge_eb_tokens(sample_ids)

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

                if hdr_idx is None or not sample_ids:
                    i += 1
                    continue

                n_samples = len(sample_ids)
                current_category = ""
                k = hdr_idx + 1
                while k < n:
                    kl = lines[k].strip()
                    if not kl:
                        k += 1
                        continue
                    if kl.startswith("Sub-Matrix:") or kl.startswith("Issue Date") or \
                       kl.startswith("When sampling date") or kl.startswith("Key:") or \
                       kl.startswith("Brief Method") or kl.startswith("The end of"):
                        break

                    parsed = self._parse_pdf_data_line(kl, n_samples)
                    if parsed is None:
                        # Might be a bare category header (e.g. "BTEX", "Physical Parameters")
                        if not re.search(r'\d', kl) and len(kl) < 60 and "----" not in kl:
                            current_category = kl
                        k += 1
                        continue

                    compound, method, lor_val, unit_val, results = parsed
                    cmp_lower = compound.lower().strip()
                    cas = _GW_CAS_MAP.get(cmp_lower, "")
                    atype = "GW_FIELD_PARAMS" if compound.upper() in _FIELD_PARAM_NAMES else "GW_VOC"

                    for si, sid in enumerate(sample_ids):
                        if si >= len(results):
                            continue
                        raw_val = results[si]
                        val, flag = _parse_value(raw_val, lor_val)

                        conv_val, conv_loq, out_unit = val, lor_val, unit_val
                        if "µg/l" in unit_val.lower() or "ug/l" in unit_val.lower():
                            if val is not None:
                                conv_val = round(val / 1000.0, 6)
                            if lor_val is not None:
                                conv_loq = round(lor_val / 1000.0, 6)
                            out_unit = "mg/L"

                        records.append({
                            "compound":      compound,
                            "cas":           cas,
                            "value":         conv_val,
                            "flag":          flag or "",
                            "unit":          out_unit,
                            "sample_id":     sid,
                            "lod":           None,
                            "loq":           conv_loq,
                            "analysis_type": atype,
                            "sampling_date": sampling_dates[si] if si < len(sampling_dates) else "",
                            "category":      current_category,
                        })
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

    def parse(self, file_obj) -> list[dict]:
        data = file_obj.read() if hasattr(file_obj, "read") else file_obj
        if isinstance(data, str):
            data = data.encode("utf-8")

        # PDF Certificate of Analysis — route to dedicated PDF extractor
        if data[:4] == b"%PDF":
            return self._parse_pdf(data)

        sheets = _parse_spreadsheetml(data)
        sheet_name = next(
            (n for n in sheets if "client water" in n.lower()),
            next((n for n in sheets if "water" in n.lower()), None)
        )
        if sheet_name is None:
            return []

        df = sheets[sheet_name]
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

            cmp_lower = compound.lower().strip()
            cas = _GW_CAS_MAP.get(cmp_lower, "")

            atype = "GW_FIELD_PARAMS" if compound.upper() in _FIELD_PARAM_NAMES else "GW_VOC"

            for ci, sid in sample_cols.items():
                if ci >= len(row):
                    continue
                raw_val = row[ci]
                val, flag = _parse_value(raw_val, lor_val)

                conv_val = val
                conv_loq = lor_val
                out_unit = unit_val or "µg/L"
                if "µg/l" in out_unit.lower() or "ug/l" in out_unit.lower():
                    if val is not None:
                        conv_val = round(val / 1000.0, 6)
                    if lor_val is not None:
                        conv_loq = round(lor_val / 1000.0, 6)
                    out_unit = "mg/L"

                records.append({
                    "compound":      compound,
                    "cas":           cas,
                    "value":         conv_val,
                    "flag":          flag or "",
                    "unit":          out_unit,
                    "sample_id":     sid,
                    "lod":           None,
                    "loq":           conv_loq,
                    "analysis_type": atype,
                    "sampling_date": sampling_dates.get(ci, ""),
                    "category":      current_category,
                })

        return records
