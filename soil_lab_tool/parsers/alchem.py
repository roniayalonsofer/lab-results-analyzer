"""
alchem.py
---------
Parser for Alchem / TO-15 soil gas laboratory Excel reports.

Expected layout (multi-sample format):
  Row 0: Canister Number:  | ... | ... | 8396 | 8573 | 8390
  Row 1: Analysis Time:    | ... | ... | ...
  Row 2: Analysis Location:| ... | ... | SG-2 | SG-6 | SG-5
  Row 3: Compound Name | CAS | LOD | LOQ | %UC | Final Conc. | Final Conc. | ...
  Row 4+: data rows

N.D. = Not Detected (below LOD)
"""

from __future__ import annotations

import io

import pandas as pd

from parsers.base import BaseParser
from core.lab_value_parser import LabValueParser


class AlchemParser(BaseParser):
    LAB_NAME = "Alchem"

    def __init__(self):
        self._vp = LabValueParser()

    # ------------------------------------------------------------------
    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        file_obj.seek(0)
        if file_obj.read(4) == b"%PDF":
            return self._parse_alchem_tph_pdf(file_obj)
        file_obj.seek(0)
        xl = pd.ExcelFile(file_obj)
        sheet = xl.sheet_names[0]
        raw = xl.parse(sheet, header=None, dtype=str).fillna("")

        # --- Find the header row (contains "Compound Name" or "CAS") ---
        header_row = self._find_header_row(raw)

        # --- Extract sample metadata from rows above the header ---
        sample_ids    = self._extract_meta_row(raw, header_row, "Analysis Location")
        canister_nums = self._extract_meta_row(raw, header_row, "Canister Number")

        # --- Parse header row ---
        headers = [str(v).strip() for v in raw.iloc[header_row].values]

        # Find column indices
        col_compound = self._find_col_idx(headers, ["compound name", "compound", "chemical", "analyte", "name"])
        col_cas      = self._find_col_idx(headers, ["cas", "cas no", "cas number"])
        col_lod      = self._find_col_idx(headers, ["lod", "lod [ug/m^3]", "mdl"])
        # All "Final Conc." columns (one per sample)
        conc_cols    = [i for i, h in enumerate(headers)
                        if "final conc" in h.lower() or "concentration" in h.lower()]

        if col_compound is None or col_cas is None:
            raise ValueError(
                f"❌ לא נמצאו עמודות Compound/CAS ב-row {header_row}. "
                f"כותרות שנמצאו: {headers}"
            )

        # --- Map conc columns → sample IDs ---
        # sample_ids list is aligned to the extra columns after the fixed cols
        def get_sample_id(col_idx: int, i: int) -> str:
            if sample_ids and i < len(sample_ids):
                return sample_ids[i]
            if canister_nums and i < len(canister_nums):
                return f"Canister-{canister_nums[i]}"
            return f"Sample-{i+1}"

        # --- Parse data rows ---
        records = []
        data_rows = raw.iloc[header_row + 1:].reset_index(drop=True)

        for _, row in data_rows.iterrows():
            values = list(row.values)
            compound = str(values[col_compound]).strip() if col_compound < len(values) else ""
            cas      = str(values[col_cas]).strip()      if col_cas      < len(values) else ""

            # Skip empty or summary rows
            if not compound or compound.lower() in ("", "nan", "compound name"):
                continue
            if "total voc" in compound.lower():
                continue  # skip summary rows

            # Handle dual-CAS compounds like "108-38-3 106-42-3"
            # Use first CAS for threshold lookup
            if " " in cas:
                cas = cas.split()[0]

            # LOD value
            lod = None
            if col_lod is not None and col_lod < len(values):
                try:
                    lod = float(values[col_lod])
                except (ValueError, TypeError):
                    lod = None

            # One record per sample column
            for i, col_idx in enumerate(conc_cols):
                raw_val   = str(values[col_idx]).strip() if col_idx < len(values) else ""
                sample_id = get_sample_id(col_idx, i)

                # N.D. → use LOD as value with '<' flag
                if raw_val.upper() in ("N.D.", "ND", "N/D", "<DL", "NOT DETECTED", ""):
                    value = lod
                    flag  = "ND"
                else:
                    value, flag = self._vp.parse(raw_val)

                records.append({
                    "lab":       self.LAB_NAME,
                    "sample_id": sample_id,
                    "compound":  compound,
                    "cas":       cas,
                    "value":     value,
                    "flag":      flag,
                    "unit":      "µg/m³",
                    "lod":       lod,
                })

        return records

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _find_header_row(self, df: pd.DataFrame) -> int:
        for i, row in df.iterrows():
            row_str = " ".join(str(v).lower() for v in row.values)
            if (("compound" in row_str) or ("name" in row_str)) and "cas" in row_str:
                return i
        return 3  # fallback

    def _extract_meta_row(self, df: pd.DataFrame, header_row: int, keyword: str) -> list[str]:
        """Find a row above header that contains `keyword`, return non-empty values after col 4."""
        for i in range(header_row):
            row = df.iloc[i]
            row_str = " ".join(str(v) for v in row.values[:3]).lower()
            if keyword.lower() in row_str:
                vals = [str(v).strip() for v in row.values[4:]]
                return [v for v in vals if v and v.lower() not in ("nan", "")]
        return []

    @staticmethod
    def _find_col_idx(headers: list[str], aliases: list[str]) -> int | None:
        for alias in aliases:
            for i, h in enumerate(headers):
                if alias.lower() in h.lower():
                    return i
        return None

    def _parse_alchem_tph_pdf(self, file_obj: io.BytesIO, filename: str = "") -> list[dict]:
        import re as _re

        def decode(s: str) -> str:
            return ''.join(chr(ord(c) + 9) if 0x20 <= ord(c) <= 0x76 else c for c in s)

        def is_sample_id(line: str) -> bool:
            # Decoded sample-ID lines start with chr(0x0E) (the 'p' glyph CID)
            # followed by the first digit of the sample number.
            return len(line) >= 2 and ord(line[0]) == 0x0E and line[1].isdigit()

        def sample_label(line: str) -> str:
            # Build "p<num> - <depth>" from the decoded sample-ID line.
            # chr(0x0E) → 'p'; keep digits and hyphens, replace everything else with space.
            s = "p" + "".join(
                c if (c.isdigit() or c == "-") else " "
                for c in line[1:]
            )
            return _re.sub(r"\s+", " ", s).strip()

        import fitz
        file_obj.seek(0)
        doc = fitz.open(stream=file_obj.read(), filetype="pdf")

        # Decode every line up-front so all later checks operate on readable text.
        lines: list[str] = []
        for page in doc:
            for raw in page.get_text().splitlines():
                decoded = decode(raw.strip())
                if decoded.strip():
                    lines.append(decoded)
        doc.close()

        # Read actual LOQ values from the LOQ row when present.
        loq = {"DRO": 30.0, "ORO": 20.0, "TPH": 50.0}
        for i, line in enumerate(lines):
            if line == "LOQ" and i + 3 < len(lines):
                try:
                    loq["DRO"] = float(lines[i + 1].replace(",", ""))
                    loq["ORO"] = float(lines[i + 2].replace(",", ""))
                    loq["TPH"] = float(lines[i + 3].replace(",", ""))
                except Exception:
                    pass
                break

        # State machine: sample-ID line followed by exactly 3 value lines.
        records: list[dict] = []
        i = 0
        while i < len(lines):
            if is_sample_id(lines[i]) and i + 3 < len(lines):
                sample = sample_label(lines[i])
                if sample.startswith("p") and len(sample) > 1 and sample[1:2].isdigit():
                    sample = "ק" + sample[1:]
                for compound, val in zip(["DRO", "ORO", "TPH"], lines[i + 1: i + 4]):
                    if val in ("N.D.", "ND"):
                        value, flag = None, "ND"
                    elif val == "<LOQ":
                        value, flag = loq[compound], "<"
                    else:
                        try:
                            value, flag = float(val.replace(",", "")), ""
                        except Exception:
                            continue
                    records.append({
                        "lab": self.LAB_NAME, "sample_id": sample,
                        "compound": compound, "cas": compound,
                        "value": value, "flag": flag,
                        "unit": "mg/kg", "analysis_type": "SOIL_TPH",
                        "sampling_date": "",
                        "loq": loq[compound],
                    })
                i += 4
            else:
                i += 1
        return records


# ──────────────────────────────────────────────────────────────────────────────
# Alchem TPH PDF parser (CID-font encoded, requires pymupdf / fitz)
# ──────────────────────────────────────────────────────────────────────────────

def decode_alchem(s: str) -> str:
    """Shift each character in the PDF's custom CID encoding back to ASCII.

    The font stores printable ASCII [0x20–0x76] shifted down by 9, so each
    character must be shifted up by 9 to recover the original text.
    Characters outside that range (Hebrew CIDs, control chars) are left as-is.
    """
    return ''.join(chr(ord(c) + 9) if 0x20 <= ord(c) <= 0x76 else c for c in s)


def parse_alchem_tph_pdf(file_bytes: bytes) -> list[dict]:
    """Extract DRO/ORO/TPH records from an Alchem TPH PDF.

    The PDF uses a custom CID font encoding; decode_alchem() recovers the text.
    After decoding, data rows start with 'p' + digit (e.g. 'p1 - 0-5') and
    the last three tokens are the DRO, ORO, TPH values.
    The final two rows contain LOD and LOQ per column.
    """
    import fitz  # pymupdf

    doc = fitz.open(stream=file_bytes, filetype="pdf")
    records: list[dict] = []
    loq_dro, loq_oro, loq_tph = 30.0, 20.0, 50.0

    for page in doc:
        raw = page.get_text()
        lines = [decode_alchem(l.strip()) for l in raw.splitlines() if l.strip()]
        for line in lines:
            parts = line.split()
            if len(parts) >= 4 and parts[0].startswith('p') and parts[0][1:2].isdigit():
                sample = ' '.join(parts[:-3])
                dro_raw, oro_raw, tph_raw = parts[-3], parts[-2], parts[-1]
                for compound, raw_val, loq in [
                    ('DRO', dro_raw, loq_dro),
                    ('ORO', oro_raw, loq_oro),
                    ('TPH', tph_raw, loq_tph),
                ]:
                    if raw_val in ('N.D.', 'ND'):
                        value, flag = loq, '<'
                    elif raw_val == '<LOQ':
                        value, flag = loq, '<'
                    else:
                        try:
                            value, flag = float(raw_val.replace(',', '')), ''
                        except Exception:
                            continue
                    records.append({
                        'lab': 'אלכם', 'sample_id': sample,
                        'compound': compound, 'cas': compound,
                        'value': value, 'flag': flag,
                        'unit': 'mg/kg', 'analysis_type': 'SOIL_TPH',
                    })

    doc.close()
    return records


class AlchemTPHPDFParser(AlchemParser):
    """Alchem TPH PDF parser; inherits _parse_alchem_tph_pdf from AlchemParser."""

    LAB_NAME = "Alchem Soil"
    ANALYSIS_TYPES = ["SOIL_TPH"]

    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        file_obj.seek(0)
        if file_obj.read(4) == b"%PDF":
            return self._parse_alchem_tph_pdf(file_obj)
        file_obj.seek(0)
        return super().parse(file_obj)
