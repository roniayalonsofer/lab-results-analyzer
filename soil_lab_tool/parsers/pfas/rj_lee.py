"""
parsers/pfas/rj_lee.py
-----------------------
Parser for RJ Lee Group PFAS EDD Excel files.

Wide-format layout:
  Row 0 (header): Sample Name | PFHxA | PFOA | PFHxS | PFOS | ...
  Data rows:      <Lab ID>    | <number or ND> | ...

Accepts an optional PDF (Chain of Custody / sample log) to resolve
7-digit Lab IDs in the sample_id column to human-readable borehole
names such as "K10 - 4.0".
"""

from __future__ import annotations

import io
import re

import pandas as pd

from parsers.base import BaseParser

_ND_VALUES   = frozenset({"nd", "not detected", "n.d.", "n/d", "<dl", "", "nan"})
_DEFAULT_LOQ = 0.2

_PFAS_CAS_MAP: dict[str, str] = {
    "pfhxa": "307-24-4",
    "pfoa":  "335-67-1",
    "pfhxs": "355-46-4",
    "pfos":  "1763-23-1",
}

_LAB_ID_RE   = re.compile(r'^\d{7}$')
_BOREHOLE_RE = re.compile(r'\b([A-Za-z]+\d*\s*-\s*[\d.]+)\b')


def extract_lab_id_map(pdf_bytes_list: list[bytes]) -> dict[str, str]:
    """Return a merged mapping of Lab ID (7-digit str) → Client ID (borehole name).

    Accepts a list of PDF byte strings and merges results from all of them,
    scanning every page via table extraction then raw text lines.
    """
    mapping: dict[str, str] = {}
    try:
        import pdfplumber
    except ImportError:
        return mapping

    for pdf_bytes in pdf_bytes_list:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                for table in (page.extract_tables() or []):
                    for row in (table or []):
                        if row:
                            _scan_row(row, mapping)

                text = page.extract_text() or ""
                for line in text.splitlines():
                    _scan_tokens(line.split(), mapping)

    return mapping


def _scan_row(row: list, mapping: dict[str, str]) -> None:
    cells     = [str(c or "").strip() for c in row]
    lab_ids   = [c for c in cells if _LAB_ID_RE.match(c)]
    boreholes = [m.group(1) for c in cells for m in [_BOREHOLE_RE.search(c)] if m]
    if lab_ids and boreholes:
        for lid in lab_ids:
            mapping.setdefault(lid, boreholes[0])
        return
    for i, cell in enumerate(cells):
        if not _LAB_ID_RE.match(cell):
            continue
        window = cells[max(0, i - 3): i + 4]
        for c in window:
            m = _BOREHOLE_RE.search(c)
            if m:
                mapping.setdefault(cell, m.group(1))
                break


def _scan_tokens(tokens: list[str], mapping: dict[str, str]) -> None:
    for i, tok in enumerate(tokens):
        if not _LAB_ID_RE.match(tok):
            continue
        window = " ".join(tokens[max(0, i - 3): i + 4])
        m = _BOREHOLE_RE.search(window)
        if m:
            mapping.setdefault(tok, m.group(1))


class RJLeePFASParser(BaseParser):
    LAB_NAME       = "RJ Lee"
    ANALYSIS_TYPES = ["SOIL_PFAS"]

    def parse(self, file_obj: io.BytesIO | str,
              pdf_bytes: list[bytes] | None = None) -> list[dict]:
        lab_id_map: dict[str, str] = {}
        if pdf_bytes:
            lab_id_map = extract_lab_id_map(pdf_bytes)

        try:
            xl = pd.ExcelFile(file_obj)
            df = xl.parse(xl.sheet_names[0], header=0, dtype=str).fillna("")
        except Exception as e:
            raise ValueError(f"RJLeePFASParser: cannot read file — {e}") from e

        if df.empty:
            return []

        sample_col    = df.columns[0]
        compound_cols = [
            c for c in df.columns[1:]
            if str(c).strip() and str(c).strip().lower() != "nan"
        ]

        records: list[dict] = []
        for _, row in df.iterrows():
            sample_id = str(row[sample_col]).strip()
            if not sample_id or sample_id.lower() in ("nan", ""):
                continue

            if _LAB_ID_RE.match(sample_id):
                sample_id = lab_id_map.get(sample_id, sample_id)

            for compound in compound_cols:
                raw = str(row[compound]).strip()
                if raw.lower() in _ND_VALUES:
                    value, flag, loq = None, "ND", _DEFAULT_LOQ
                else:
                    try:
                        value = float(raw)
                    except ValueError:
                        value = None
                    flag = ""
                    loq  = None

                compound_str = str(compound).strip()
                cas = _PFAS_CAS_MAP.get(compound_str.lower(), "")

                records.append({
                    "lab":           self.LAB_NAME,
                    "sample_id":     sample_id,
                    "compound":      compound_str,
                    "cas":           cas,
                    "value":         value,
                    "flag":          flag,
                    "unit":          "ng/g",
                    "lod":           None,
                    "loq":           loq,
                    "analysis_type": "SOIL_PFAS",
                })

        return records
