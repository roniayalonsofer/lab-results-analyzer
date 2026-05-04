"""
parsers/soil/machon_haneft.py
------------------------------
Parser for מכון הנפט (Machon HaNeft / Israel Petroleum Institute) soil Excel reports.

Supports two sheet types in the same report:
  1. TPH sheet  — DRO / ORO / Total TPH per sample
  2. מתכות sheet — Metals (ICP-OES), one metal per row, sample columns

Report structure (wide format):
  Rows 0–N:  metadata / header text
  Then:      compound rows × sample columns

The parser locates the data block by scanning for keywords:
  - "TPH", "DRO" → TPH sheet header
  - "CAS" or "Cas.No" → metals sheet data start

File naming patterns: 1018*.xlsx, 1494*.xlsx, 1705*.xlsx
"""

from __future__ import annotations

import io
import re

import pandas as pd

from parsers.base import BaseParser
from core.lab_value_parser import LabValueParser


# BTEX/VOC compounds that may appear as column headers in the second TPH-sheet block.
# Keys are lowercase.  Hebrew names added for Machon HaNeft shimshon-format files.
_BTEX_COLS: dict[str, str] = {
    # English names
    "mtbe":             "1634-04-4",
    "benzene":          "71-43-2",
    "toluene":          "108-88-3",
    "ethylbenzene":     "100-41-4",
    "ethyl benzene":    "100-41-4",
    "xylene":           "1330-20-7",
    "xylenes":          "1330-20-7",
    "naphthalene":      "91-20-3",
    "styrene":          "100-42-5",
    # Hebrew names (shimshon / Machon HaNeft format)
    "בנזן":             "71-43-2",
    "טולואן":           "108-88-3",
    "אתיל-בנזן":        "100-41-4",
    "אתיל בנזן":        "100-41-4",
    "קסילן":            "1330-20-7",
    "קסילנים":          "1330-20-7",
}

_CAS_RE2     = re.compile(r'^\d{2,7}-\d{2}-\d$')
# Also match CAS embedded in strings like "CAS No.1634-04-4"
_CAS_EXTRACT = re.compile(r'(\d{2,7}-\d{2}-\d)')

# Map Hebrew metals names → CAS
METALS_HE_CAS: dict[str, str] = {
    "עופרת":    "7439-92-1",   # Pb
    "אבץ":      "7440-66-6",   # Zn
    "נחושת":    "7440-50-8",   # Cu
    "כסף":      "7440-22-4",   # Ag
    "אלומיניום": "7429-90-5",  # Al
    "ארסן":     "7440-38-2",   # As
    "בורון":    "7440-42-8",   # B
    "בריום":    "7440-39-3",   # Ba
    "ברום":     "7726-95-6",   # Be (Note: ברום = Bromine; בריליום = Be)
    "בריליום":  "7440-41-7",   # Be
    "קדמיום":   "7440-43-9",   # Cd
    "קובלט":    "7440-48-4",   # Co
    # NOTE: CAS 7440-47-3 = Chromium (total). The Israeli V7 standard has NO VSL
    # threshold for total Cr — only for Cr(VI) (18540-29-9, VSL=1.33 mg/kg) and
    # Cr(III) insoluble salts (16065-83-1, VSL=109,449 mg/kg).
    # Threshold lookup for Cr will return None until a site-specific speciation is known.
    "כרום":     "7440-47-3",   # Cr total — see note above
    "ברזל":     "7439-89-6",   # Fe
    "מנגן":     "7439-96-5",   # Mn
    "מוליבדן":  "7439-98-7",   # Mo
    "ניקל":     "7440-02-0",   # Ni
    "סלניום":   "7782-49-2",   # Se
    "אנטימון":  "7440-36-0",   # Sb
    "תלאיום":   "7440-28-0",   # Tl
    "ונדיום":   "7440-62-2",   # V
    "כספית":    "7439-97-6",   # Hg
    "בריום":    "7440-39-3",   # Ba
}

SYMBOL_CAS: dict[str, str] = {
    "pb": "7439-92-1",
    "zn": "7440-66-6",
    "cu": "7440-50-8",
    "ag": "7440-22-4",
    "al": "7429-90-5",
    "as": "7440-38-2",
    "b":  "7440-42-8",
    "ba": "7440-39-3",
    "be": "7440-41-7",
    "cd": "7440-43-9",
    "co": "7440-48-4",
    "cr": "7440-47-3",   # total Cr — no VSL in Israeli V7 standard (see METALS_HE_CAS note)
    "fe": "7439-89-6",
    "mn": "7439-96-5",
    "mo": "7439-98-7",
    "ni": "7440-02-0",
    "se": "7782-49-2",
    "sb": "7440-36-0",
    "tl": "7440-28-0",
    "v":  "7440-62-2",
    "hg": "7439-97-6",
}


class MachonHaneftSoilParser(BaseParser):
    LAB_NAME = "מכון הנפט"
    ANALYSIS_TYPES = ["SOIL_TPH", "SOIL_METALS", "SOIL_MBTEX"]

    def __init__(self):
        self._vp = LabValueParser()

    # ------------------------------------------------------------------
    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        xl = pd.ExcelFile(file_obj)
        sheet_map = {s.strip().lower(): s for s in xl.sheet_names}

        records: list[dict] = []

        # --- TPH sheet (also parses BTEX/MTBE second block if present) ---
        for key in ("tph", "tph+btex", "tph & btex"):
            if key in sheet_map:
                records.extend(self._parse_tph(xl, sheet_map[key]))
                records.extend(self._parse_btex_block(xl, sheet_map[key]))
                break
        else:
            # Try first sheet if it looks like TPH
            first = xl.sheet_names[0]
            raw = xl.parse(first, header=None, dtype=str).fillna("")
            flat = " ".join(raw.values.flatten().tolist()).lower()
            if "tph" in flat or "dro" in flat:
                records.extend(self._parse_tph(xl, first))
                records.extend(self._parse_btex_block(xl, first))

        # --- Metals sheet ---
        for key in ("מתכות", "metals", "metal"):
            if key in sheet_map:
                records.extend(self._parse_metals(xl, sheet_map[key]))
                break

        return records

    # ------------------------------------------------------------------
    # TPH sheet
    # ------------------------------------------------------------------
    def _parse_tph(self, xl: pd.ExcelFile, sheet_name: str) -> list[dict]:
        raw = xl.parse(sheet_name, header=None, dtype=str).fillna("")

        # Find the header row: contains "TPH" and "DRO"
        header_row_idx = None
        for i, row in raw.iterrows():
            vals_lower = [str(v).strip().lower() for v in row.values]
            if "tph" in vals_lower and "dro" in vals_lower:
                header_row_idx = i
                break

        if header_row_idx is None:
            return []

        headers = [str(v).strip() for v in raw.iloc[header_row_idx].values]

        # Find column indices for each parameter
        def find_col(keywords):
            for k in keywords:
                for i, h in enumerate(headers):
                    if k.lower() in h.lower():
                        return i
            return None

        col_name  = find_col(["בדיקה", "sample", "שם"])   # sample name column
        col_tph   = find_col(["tph"])
        col_dro   = find_col(["dro"])
        col_oro   = find_col(["oro"])

        if col_tph is None and col_dro is None:
            return []

        # If no explicit sample name column, use first non-empty column
        if col_name is None:
            col_name = 0

        records: list[dict] = []
        # Keywords whose presence in col-A marks non-sample rows
        _TPH_SKIP_LO = {"nan", "", "lod", "loq", "method", "units",
                        "mg/kg", "%", "limit", "שיטות", "שיטה", "יחידות"}
        _TPH_SKIP_STARTS = ("גבול", "ערך סף", "ערך יעד", "שיט",
                            "epa", "target", "threshold")

        # Find first row after header where col-A looks like a real sample name
        data_start = header_row_idx + 1
        for i in range(header_row_idx + 1, min(header_row_idx + 10, len(raw))):
            val    = str(raw.iloc[i, col_name]).strip()
            val_lo = val.lower()
            if (val_lo in _TPH_SKIP_LO or
                    any(val_lo.startswith(p) for p in _TPH_SKIP_STARTS)):
                continue
            data_start = i
            break

        # ── Extract LOD/LOQ values from metadata rows between header and data_start ──
        # Matches "גבול גילוי" (detection limit), "LOD", "LOQ", "גבול כימות" rows.
        # When only some compound columns have a value, the highest found is used
        # as fallback for the others (same instrument / same method implies same limit).
        lod_per_col: dict[int, float] = {}
        for i in range(header_row_idx + 1, data_start):
            val_a = str(raw.iloc[i, col_name]).strip().lower()
            is_lod_row = (
                "גבול גילוי" in val_a or "גבול כימות" in val_a
                or val_a in ("lod", "loq", "detection limit", "quantitation limit")
            )
            if not is_lod_row:
                continue
            found: dict[int, float] = {}
            for col in (col_dro, col_oro, col_tph):
                if col is None or col >= raw.shape[1]:
                    continue
                try:
                    found[col] = float(str(raw.iloc[i, col]).replace(",", ""))
                except (ValueError, TypeError):
                    pass
            if found:
                # Propagate the max found value to any compound columns left blank
                fallback = max(found.values())
                for col in (col_dro, col_oro, col_tph):
                    if col is not None:
                        lod_per_col[col] = found.get(col, fallback)

        # Parse data rows — compound order: DRO | ORO | Total TPH (Total last = rightmost column)
        params = []
        if col_dro is not None:
            params.append(("DRO", "C10-C40", col_dro))
        if col_oro is not None:
            params.append(("ORO", "C10-C40", col_oro))
        if col_tph is not None:
            params.append(("Total TPH", "C10-C40", col_tph))

        for i in range(data_start, len(raw)):
            row = raw.iloc[i]
            sample_name = str(row.iloc[col_name]).strip()
            sn_lo = sample_name.lower()
            if not sample_name or sn_lo == "nan":
                continue
            # Skip inline metadata rows (LOQ/method/threshold rows interspersed in data)
            if sn_lo in _TPH_SKIP_LO:
                continue
            if any(sn_lo.startswith(p) for p in _TPH_SKIP_STARTS):
                continue
            # Stop at section boundaries and footer rows.
            # "בדיקה" alone (exact) signals the start of a new data section header.
            if sn_lo == "בדיקה":
                break
            if any(kw in sn_lo for kw in ("חתימה", "signature", "approved",
                                           "מאשר", "note")):
                break

            for cmp, cas, col in params:
                if col is None or col >= len(row):
                    continue
                raw_val = str(row.iloc[col]).strip()
                if raw_val.lower() in ("nan", ""):
                    continue

                lod = lod_per_col.get(col)
                if raw_val.lower() in ("not detected", "nd", "n.d.", "<dl"):
                    value, flag = None, "ND"
                elif raw_val.startswith("<"):
                    # Below detection/quantitation limit — extract the numeric threshold
                    # as the LOD for this cell and mark as below-LOD (not just ND).
                    try:
                        cell_limit = float(raw_val[1:].replace(",", "").strip())
                        # Use the cell-embedded limit as LOD if not set from metadata row
                        if lod is None:
                            lod = cell_limit
                        value, flag = None, "<LOD"
                    except ValueError:
                        value, flag = self._vp.parse(raw_val)
                else:
                    try:
                        value = float(raw_val.replace(",", ""))
                        flag = ""
                    except ValueError:
                        value, flag = self._vp.parse(raw_val)
                    # Treat numeric value ≤ LOD as not-detected (same as MBTEX logic)
                    if value is not None and lod is not None and value <= lod + 1e-9:
                        value, flag = None, "ND"

                records.append({
                    "lab":           self.LAB_NAME,
                    "sample_id":     sample_name,
                    "compound":      cmp,
                    "cas":           cas,
                    "value":         value,
                    "flag":          flag,
                    "unit":          "mg/kg",
                    "lod":           lod,
                    "analysis_type": "SOIL_TPH",
                })

        return records

    # ------------------------------------------------------------------
    # BTEX / MTBE block (second data block in TPH sheet)
    # ------------------------------------------------------------------
    def _parse_btex_block(self, xl: pd.ExcelFile, sheet_name: str) -> list[dict]:
        """Parse the BTEX/MBTEX data block that follows the TPH block in the
        same sheet.  Compounds appear as column headers; each data row is one
        sample.  Supports both English and Hebrew compound headers (shimshon format).
        Emits records with analysis_type='SOIL_MBTEX'."""
        raw = xl.parse(sheet_name, header=None, dtype=str).fillna("")

        # Scan for a header row (starting from row 20 to skip title/metadata,
        # but well before the old hardcoded row 50 so shimshon_1 at row ~46 is found)
        # that contains ≥ 2 recognised BTEX compound names as cell values.
        btex_header_row = None
        for i in range(20, len(raw)):
            vals_lo = [str(v).strip().lower() for v in raw.iloc[i].values]
            # Must have ≥2 BTEX names — this never matches the TPH header (TPH/DRO/ORO)
            if sum(1 for v in vals_lo if v in _BTEX_COLS) >= 2:
                btex_header_row = i
                break

        if btex_header_row is None:
            return []

        headers_lo   = [str(v).strip().lower() for v in raw.iloc[btex_header_row].values]
        headers_orig = [str(v).strip()         for v in raw.iloc[btex_header_row].values]

        # Build compound column list: [(col_idx, display_name, cas), ...]
        # Use the original (non-lowercased) header value as display name so Hebrew
        # names ("בנזן") and uppercase acronyms ("MTBE") are preserved.
        compound_cols: list[tuple[int, str, str]] = []
        for ci, (h, h_orig) in enumerate(zip(headers_lo, headers_orig)):
            if h in _BTEX_COLS:
                compound_cols.append((ci, h_orig, _BTEX_COLS[h]))

        if not compound_cols:
            return []

        # Scan the next few rows for:
        #   (a) a CAS row — overrides the dict CAS values
        #   (b) LOD row   — "גבול גילוי הבדיקה" or "lod"
        #   (c) LOQ row   — "גבול כימות הבדיקה" or "loq"
        lod_per_col: dict[int, float] = {}
        loq_per_col: dict[int, float] = {}

        _LOOKAHEAD = 10
        for ri in range(btex_header_row + 1,
                        min(btex_header_row + _LOOKAHEAD, len(raw))):
            row_vals = [str(v).strip() for v in raw.iloc[ri].values]
            col_a    = row_vals[0].lower() if row_vals else ""

            # CAS row: cell values that look like CAS numbers (plain or embedded)
            extracted_cas = [_CAS_EXTRACT.search(v) for v in row_vals]
            if any(m for m in extracted_cas):
                compound_cols = [
                    (ci, name,
                     extracted_cas[ci].group(1)
                     if ci < len(extracted_cas) and extracted_cas[ci]
                     else cas)
                    for ci, name, cas in compound_cols
                ]

            # LOD row
            if "גבול גילוי" in col_a or col_a == "lod":
                for ci, _, _ in compound_cols:
                    try:
                        lod_per_col[ci] = float(row_vals[ci].replace(",", ""))
                    except (ValueError, IndexError, AttributeError):
                        pass

            # LOQ row
            if "גבול כימות" in col_a or col_a == "loq":
                for ci, _, _ in compound_cols:
                    try:
                        loq_per_col[ci] = float(row_vals[ci].replace(",", ""))
                    except (ValueError, IndexError, AttributeError):
                        pass

        # Find data_start: first row after the header where col-0 looks like a
        # sample name (not a metadata / units / LOD / CAS row).
        col_name = 0
        _SKIP_PREFIXES = ("גבול", "lod", "loq", "cas", "units", "method",
                          "limit", "mg/kg", "%", "nan", "")
        data_start = btex_header_row + _LOOKAHEAD   # safe default
        for i in range(btex_header_row + 1, min(btex_header_row + _LOOKAHEAD + 2, len(raw))):
            val    = str(raw.iloc[i, col_name]).strip()
            val_lo = val.lower()
            if (val_lo in ("nan", "") or
                    any(val_lo.startswith(p) for p in _SKIP_PREFIXES) or
                    _CAS_RE2.match(val) or
                    _CAS_EXTRACT.search(val)):
                continue
            data_start = i
            break

        records: list[dict] = []
        for i in range(data_start, len(raw)):
            row = raw.iloc[i]
            sample_name = str(row.iloc[col_name]).strip()
            if not sample_name or sample_name.lower() in ("nan", ""):
                continue
            if any(kw in sample_name.lower() for kw in
                   ("חתימה", "signature", "approved", "מאשר", "note", "הערה")):
                break

            for ci, cname, cas in compound_cols:
                if ci >= len(row):
                    continue
                raw_val = str(row.iloc[ci]).strip()
                if not raw_val or raw_val.lower() in ("nan", ""):
                    continue

                if raw_val.lower() in ("not detected", "nd", "n.d.", "<dl"):
                    value, flag = None, "ND"
                elif raw_val.startswith("<"):
                    value, flag = self._vp.parse(raw_val)
                else:
                    try:
                        value = float(raw_val.replace(",", ""))
                        flag = ""
                    except ValueError:
                        value, flag = self._vp.parse(raw_val)
                    # Machon HaNeft shimshon reports write the LOD value in cells
                    # where a compound was not detected ("גבול גילוי הבדיקה" = ND).
                    # Treat any numeric value ≤ LOD as not-detected.
                    lod_v = lod_per_col.get(ci)
                    if value is not None and lod_v is not None and value <= lod_v + 1e-9:
                        value, flag = None, "ND"

                records.append({
                    "lab":           self.LAB_NAME,
                    "sample_id":     sample_name,
                    "compound":      cname,
                    "cas":           cas,
                    "value":         value,
                    "flag":          flag,
                    "unit":          "mg/kg",
                    "lod":           lod_per_col.get(ci),
                    "analysis_type": "SOIL_MBTEX",
                })

        return records

    # ------------------------------------------------------------------
    # Metals sheet
    # ------------------------------------------------------------------
    def _parse_metals(self, xl: pd.ExcelFile, sheet_name: str) -> list[dict]:
        raw = xl.parse(sheet_name, header=None, dtype=str).fillna("")

        # Find the row containing "Cas.No" or "CAS" → compound column header row
        cas_row_idx = None
        for i, row in raw.iterrows():
            row_str = " ".join(str(v) for v in row.values).lower()
            if "cas" in row_str and ("compound" in row_str or "cas.no" in row_str or
                                      "יחידות" in row_str or len(row.values) > 5):
                cas_row_idx = i
                break

        if cas_row_idx is None:
            return []

        # Sample IDs are in the row just ABOVE cas_row_idx (in the data columns)
        # Find which columns have sample data (cols 5+)
        sample_id_row = raw.iloc[cas_row_idx - 1] if cas_row_idx > 0 else None

        headers = [str(v).strip() for v in raw.iloc[cas_row_idx].values]

        def find_col(keys):
            for k in keys:
                for i, h in enumerate(headers):
                    if k.lower() in h.lower():
                        return i
            return None

        col_symbol = find_col(["compound", "symbol", "element", "param"]) or 0
        col_name   = 1   # Usually col 1 is Hebrew name
        col_cas    = find_col(["cas"])
        col_unit   = find_col(["יחידות", "unit"])
        col_lod    = find_col(["lod", "detect", "dl", "גילוי"])
        col_loq    = find_col(["loq", "כימות"])

        # Sample columns: everything after LOQ/LOD columns (typically col 6+)
        fixed_max = max(c for c in [col_symbol, col_cas, col_unit, col_lod, col_loq]
                        if c is not None) + 1
        sample_cols = list(range(fixed_max, len(headers)))

        # Extract sample IDs from the row above headers
        sample_ids: list[str] = []
        if sample_id_row is not None:
            for c in sample_cols:
                val = str(sample_id_row.iloc[c]).strip() if c < len(sample_id_row) else ""
                sample_ids.append(val if val and val.lower() not in ("nan", "") else f"S{c}")
        else:
            sample_ids = [f"S{c}" for c in sample_cols]

        records: list[dict] = []
        # Process all data blocks (there may be multiple batches with different sample sets)
        # Re-read the entire sheet to handle multiple header blocks
        self._parse_metals_block(raw, cas_row_idx, col_symbol, col_cas, col_unit,
                                 col_lod, col_loq, sample_cols, sample_ids, records)

        # Check for additional batches (another block further down)
        for extra_row in range(cas_row_idx + 30, len(raw) - 2):
            row_str = " ".join(str(v) for v in raw.iloc[extra_row].values).lower()
            if "cas" in row_str:
                extra_headers = [str(v).strip() for v in raw.iloc[extra_row].values]
                extra_max = max(c for c in [col_symbol, col_cas, col_unit, col_lod, col_loq]
                                if c is not None) + 1
                extra_sample_cols = [c for c in range(extra_max, len(extra_headers))
                                     if str(extra_headers[c]).strip() not in ("nan", "")]
                extra_sample_id_row = raw.iloc[extra_row - 1]
                extra_sids = [str(extra_sample_id_row.iloc[c]).strip()
                              if c < len(extra_sample_id_row) else f"S{c}"
                              for c in extra_sample_cols]
                self._parse_metals_block(raw, extra_row, col_symbol, col_cas, col_unit,
                                         col_lod, col_loq, extra_sample_cols, extra_sids, records)
                break

        return records

    def _parse_metals_block(self, raw, header_idx, col_symbol, col_cas, col_unit,
                            col_lod, col_loq, sample_cols, sample_ids, records):
        for i in range(header_idx + 1, len(raw)):
            row = raw.iloc[i]
            symbol = str(row.iloc[col_symbol]).strip() if col_symbol < len(row) else ""
            if not symbol or symbol.lower() in ("nan", ""):
                continue
            # Stop at footer
            if any(kw in symbol.lower() for kw in ("חתימה", "signature", "approved",
                                                    "note", "הערה")):
                break

            # Get CAS — prefer column value, fall back to symbol lookup
            cas = ""
            if col_cas is not None and col_cas < len(row):
                cas = str(row.iloc[col_cas]).strip()
                if cas.lower() in ("nan", ""):
                    cas = ""
            if not cas:
                cas = SYMBOL_CAS.get(symbol.strip().lower(), "")

            unit_val = ""
            if col_unit is not None and col_unit < len(row):
                unit_val = str(row.iloc[col_unit]).strip()
                if unit_val.lower() in ("nan", ""):
                    unit_val = "mg/kg"

            lod = None
            if col_lod is not None and col_lod < len(row):
                try:
                    lod = float(str(row.iloc[col_lod]).replace(",", ""))
                except (ValueError, TypeError):
                    pass

            for j, sc in enumerate(sample_cols):
                if sc >= len(row):
                    continue
                raw_val = str(row.iloc[sc]).strip()
                if raw_val.lower() in ("nan", ""):
                    continue

                if raw_val.lower() in ("not detected", "nd", "n.d."):
                    value, flag = None, "ND"
                elif raw_val.startswith("<"):
                    value, flag = self._vp.parse(raw_val)
                else:
                    try:
                        value = float(raw_val.replace(",", ""))
                        flag = ""
                    except ValueError:
                        value, flag = self._vp.parse(raw_val)

                sid = sample_ids[j] if j < len(sample_ids) else f"S{sc}"
                records.append({
                    "lab":           self.LAB_NAME,
                    "sample_id":     sid,
                    "compound":      symbol,
                    "cas":           cas,
                    "value":         value,
                    "flag":          flag,
                    "unit":          unit_val or "mg/kg",
                    "lod":           lod,
                    "analysis_type": "SOIL_METALS",
                })
