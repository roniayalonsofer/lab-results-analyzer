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
            file_obj.seek(0)
            raw = file_obj.read().decode('latin-1', errors='ignore')
            if '3CFH' in raw:
                file_obj.seek(0)
                return self._parse_alchem_tph_pdf(file_obj)
            else:
                file_obj.seek(0)
                return self._parse_alchem_readable_pdf(file_obj)
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

        file_obj.seek(0)
        raw_bytes = file_obj.read()

        # Encoded (CID-font) format: use fitz
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
        doc = fitz.open(stream=raw_bytes, filetype="pdf")

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

    def _parse_alchem_readable_pdf(self, file_obj: io.BytesIO, filename: str = "") -> list[dict]:
        import pdfplumber
        file_obj.seek(0)
        raw_bytes = file_obj.read()
        records = []
        with pdfplumber.open(io.BytesIO(raw_bytes)) as pdf:
            for page in pdf.pages:
                page_text = (page.extract_text() or "").lower()
                for table in (page.extract_tables() or []):
                    if not table or len(table) < 2:
                        continue
                    h0 = [str(c or "").strip() for c in table[0]]
                    h0_low = [h.lower() for h in h0]
                    hdrs = " ".join(h0_low)
                    if "dro" in hdrs and "oro" in hdrs and "tph" in hdrs:
                        records.extend(self._parse_readable_tph_table(table))
                    elif "sample name" in hdrs or (h0_low and "sample name" in h0_low[0]):
                        atype = "SOIL_SVOC" if "8270" in page_text else "SOIL_VOC"
                        records.extend(self._parse_readable_voc_table(table, atype))
        return records

    def _parse_readable_tph_table(self, table):
        h = [str(c or "").strip().lower() for c in table[0]]
        col_s = 0
        col_d = next((i for i, x in enumerate(h) if "dro" in x), None)
        col_o = next((i for i, x in enumerate(h) if "oro" in x), None)
        col_t = next((i for i, x in enumerate(h) if "tph" in x), None)
        loq = {"DRO": 30.0, "ORO": 20.0, "TPH": 50.0}
        records = []
        for row in table[1:]:
            if not row: continue
            s = str(row[col_s] or "").strip()
            if not s or s.lower() in ("sample name", "loq", "lod", "[mg/kg]", ""): continue
            if not any(c.isdigit() for c in s): continue
            sid, depth = self._normalize_sample(s)
            for cmp, ci, lq in [("DRO", col_d, loq["DRO"]), ("ORO", col_o, loq["ORO"]), ("TPH", col_t, loq["TPH"])]:
                if ci is None or ci >= len(row): continue
                value, flag = self._parse_readable_value(str(row[ci] or "").strip(), lq)
                records.append({"lab": self.LAB_NAME, "sample_id": sid, "depth": depth,
                    "compound": cmp, "cas": cmp, "value": value, "flag": flag,
                    "unit": "mg/kg", "analysis_type": "SOIL_TPH", "sampling_date": ""})
        return records

    def _parse_readable_voc_table(self, table, analysis_type="SOIL_VOC"):
        h0 = [str(c or "").strip() for c in table[0]]
        h0_low = [h.lower() for h in h0]
        col_lod = next((i for i, h in enumerate(h0_low) if h.strip() == "lod"), None)
        col_loq = next((i for i, h in enumerate(h0_low) if h.strip() == "loq"), None)
        end_col = col_lod if col_lod else (len(h0) - 2)
        sample_names = []
        for i in range(2, end_col):
            if h0[i].strip():
                sid, depth = self._normalize_sample(h0[i])
                sample_names.append((i, sid, depth))
        records = []
        for row in table[2:]:
            if not row or not row[0]: continue
            compound = str(row[0] or "").strip().replace("\n", " ")
            if not compound or compound.lower() in ("compound name", "", "[mg/kg]", "lod", "loq", "sample name", "cas"): continue
            cas = str(row[1] or "").strip() if len(row) > 1 else ""
            lod_val = 0.01
            if col_lod and col_lod < len(row):
                try: lod_val = float(str(row[col_lod] or "").strip())
                except: pass
            loq_val = 0.02
            if col_loq and col_loq < len(row):
                try: loq_val = float(str(row[col_loq] or "").strip())
                except: pass
            for ci, sid, depth in sample_names:
                if ci >= len(row): continue
                value, flag = self._parse_readable_value(str(row[ci] or "").strip(), loq=loq_val, lod=lod_val)
                records.append({"lab": self.LAB_NAME, "sample_id": sid, "depth": depth,
                    "compound": compound, "cas": cas, "value": value, "flag": flag,
                    "unit": "mg/kg", "analysis_type": analysis_type, "sampling_date": "",
                    "lod": lod_val, "loq": loq_val})
        return records

    def _normalize_sample(self, raw: str):
        import re as _re

        def _fix_sid(sid: str) -> str:
            """Normalize '9ק' → 'ק-9', '9ק-DUP' → 'ק-9-DUP'."""
            sid = _re.sub(r'^(\d+)(ק)$', r'ק-\1', sid)
            sid = _re.sub(r'^(\d+)(ק)-DUP$', r'ק-\1-DUP', sid)
            return sid

        s = raw.strip()
        # Convert K to ק
        s = s.replace("K-", "ק-").replace("K ", "ק ").replace("k-", "ק-")

        # Format: "3-0 - 11ק" → split by " - "
        if " - " in s:
            parts = s.split(" - ", 1)
            depth_candidate = parts[0].strip().replace("-", ".")
            try:
                float(depth_candidate)
                depth = depth_candidate
                sample_id = "ק-" + parts[1].strip().lstrip("ק").lstrip("-") if parts[1].strip().startswith("ק") else parts[1].strip()
                return _fix_sid(sample_id), depth
            except ValueError:
                pass

        # Format: "K-10-3.0" or "ק-10-3.0"
        parts = s.rsplit("-", 1)
        if len(parts) == 2:
            try:
                float(parts[1].replace("m", ""))
                return _fix_sid(parts[0]), parts[1].replace("m", "")
            except ValueError:
                pass

        # Format: "ק-9 3.0m"
        parts2 = s.rsplit(" ", 1)
        if len(parts2) == 2:
            try:
                float(parts2[1].replace("m", ""))
                return _fix_sid(parts2[0]), parts2[1].replace("m", "")
            except ValueError:
                pass

        return _fix_sid(s), ""

    def _parse_readable_value(self, v, loq=0.02, lod=None):
        if v in ("<MDL", "<mdl"):
            return lod, "<"
        if v in ("<LOQ", "<MRL", "<loq"):
            return loq, "<"
        if v in ("N.D.", "ND", "N.D", "n.d."):
            return None, "ND"
        try:
            return float(v.replace(",", "")), ""
        except:
            return None, "ND"


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
            file_obj.seek(0)
            raw = file_obj.read().decode('latin-1', errors='ignore')
            if '3CFH' in raw:
                file_obj.seek(0)
                return self._parse_alchem_tph_pdf(file_obj)
            else:
                file_obj.seek(0)
                return self._parse_alchem_readable_pdf(file_obj)
        file_obj.seek(0)
        return super().parse(file_obj)
