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

import pandas as pd

from parsers.base import BaseParser
from core.cas_lookup import name_to_cas


# ---------------------------------------------------------------------------
# Shared format reader
# ---------------------------------------------------------------------------

def _parse_als_sheet(xl: pd.ExcelFile, sheet_name: str) -> tuple[list[str], dict[int, str], list[tuple]]:
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

    # ── Sample IDs from row 8 (idx 7) — same column positions as data ──
    sample_row = raw.iloc[7]
    sample_cols: dict[int, str] = {}
    for ci in range(first_smp_col, len(sample_row)):
        sid = str(sample_row.iloc[ci]).strip()
        if sid and sid.lower() not in ("nan", ""):
            sample_cols[ci] = sid

    # If row 8 yielded nothing, fall back to header row sample columns
    if not sample_cols:
        for ci in range(first_smp_col, len(hdr)):
            sid = hdr[ci]
            if sid and sid.lower() not in ("nan", ""):
                sample_cols[ci] = sid

    # ── Data rows ──
    data_rows = []
    for ri in range(hdr_row_idx + 1, len(raw)):
        row = raw.iloc[ri]
        compound = str(row.iloc[compound_col]).strip()
        if not compound or compound.lower() in ("nan", "", "parameter", "compound", "analyte"):
            continue

        unit_val = str(row.iloc[unit_col]).strip() if unit_col < len(row) else "mg/kg"
        if not unit_val or unit_val.lower() == "nan":
            unit_val = "mg/kg"

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

    return hdr, sample_cols, data_rows


def _parse_value(raw_val: str, loq: float | None) -> tuple[float | None, str | None]:
    """Parse a raw cell string into (value, flag)."""
    v = raw_val.strip()
    if not v or v.lower() == "nan" or v == "----":
        return loq, "<LOQ"
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
    })

    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        xl = pd.ExcelFile(file_obj)

        # Collect all "Client *" sheets (SOIL, PFAS, VOC …) so PFAS-only
        # sheets named "Client PFAS 1" are not silently dropped.
        target_sheets = [s for s in xl.sheet_names
                         if any(tag in s for tag in ("Client SOIL", "Client PFAS",
                                                     "Client VOC", "Client SVOC",
                                                     "Client Metal"))]
        if not target_sheets:
            # Fallback: accept any single "SOIL" or "PFAS" sheet
            target_sheets = [s for s in xl.sheet_names
                             if "SOIL" in s.upper() or "PFAS" in s.upper()]
        if not target_sheets:
            raise ValueError(f"No recognisable ALS sheet found. Sheets: {xl.sheet_names}")

        records = []
        for sheet in target_sheets:
            try:
                _, sample_cols, data_rows = _parse_als_sheet(xl, sheet)
            except Exception:
                continue

            for compound, unit, loq, sample_vals in data_rows:
                cas   = name_to_cas(compound)
                atype = self._analysis_type(compound)
                # µg/kg DW == ng/g numerically — normalise unit label for PFAS
                norm_unit = "ng/g" if atype == "SOIL_PFAS" else unit
                for ci, raw_val in sample_vals.items():
                    value, flag = _parse_value(raw_val, loq)
                    if value is None and flag is None:
                        continue
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
                    })

        return records

    def _analysis_type(self, compound: str) -> str:
        c = compound.upper()
        if any(k in c for k in self._METALS):
            return "SOIL_METALS"
        if any(k in c for k in self._VOC):
            return "SOIL_VOC"
        # PFAS must be checked before TPH: "ORO" (Oil Range Organics keyword)
        # is a substring of "PERFLUORO", so TPH would wrongly claim all PFAS.
        if any(k in c for k in (
            "PFAS", "PFOA", "PFOS", "PFBS", "PFBA", "PFNA", "PFDA", "PFUA",
            "PFHX", "PFPE", "PFDO", "PFDE", "FOSA", "HFPO",
            "PERFLUORO", "FLUOROTELOMER", "SULFONAMIDE",
        )):
            return "SOIL_PFAS"
        if any(k in c for k in ("TPH", "PETROLEUM", "DRO", "ORO", "GRO")):
            return "SOIL_TPH"
        return "SOIL_SVOC"


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

    _SECTION_MAP: list[tuple[tuple[str, ...], str]] = [
        (("polycyclic", "aromatic", "pah"),         "SOIL_SVOC"),
        (("semi-volatile", "svoc", "extractable"),   "SOIL_SVOC"),
        (("volatile", "voc"),                        "SOIL_VOC"),
        (("metal", "heavy metal", "inorganic"),      "SOIL_METALS"),
        (("petroleum", "tph", "dro", "gro", "oro"),  "SOIL_TPH"),
        (("pfas", "perfluoro", "fluorotelomer"),     "SOIL_PFAS"),
    ]

    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        try:
            import pdfplumber
        except ImportError as exc:
            raise ImportError("pdfplumber is required: pip install pdfplumber") from exc

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
                if first_lo in ("parameter", "analyte", "compound"):
                    continue

                # ── Single-cell concatenated row ──────────────────────────
                # All content landed in cells[0]; parse with regex.
                if len([c for c in cells if c]) == 1 and first:
                    for line in first.splitlines():
                        line = line.strip()
                        m = self._CONCAT_ROW_RE.match(line)
                        if not m:
                            continue
                        compound  = re.sub(r'\s+S-[A-Z0-9]+$', '', m.group(1)).strip()
                        loq_str   = m.group(2)
                        row_unit  = m.group(3).strip()
                        results   = m.group(4).split()
                        loq: float | None = None
                        try:
                            loq = float(loq_str)
                        except ValueError:
                            pass
                        cas = name_to_cas(compound) or ""
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
                lor_raw  = cells[2] if len(cells) > 2 else ""
                unit     = cells[3] if len(cells) > 3 else "mg/kg"
                if not unit or unit.lower() == "nan":
                    unit = "mg/kg"
                compound = self._METHOD_SUFFIX_RE.sub('', first).strip()

                loq = None
                try:
                    loq = float(lor_raw.lstrip("<"))
                except (ValueError, TypeError):
                    pass

                cas = name_to_cas(compound) or ""

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

        _, sample_cols, data_rows = _parse_als_sheet(xl, sheet)
        records = []

        for compound, unit, loq, sample_vals in data_rows:
            cas = ""
            for ci, raw_val in sample_vals.items():
                value, flag = _parse_value(raw_val, loq)
                if value is None and flag is None:
                    continue
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
                })

        return records
