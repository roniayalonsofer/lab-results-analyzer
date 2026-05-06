"""
parsers/soil/bactochem.py
--------------------------
Parser for Bactochem laboratory PDF reports.

PDF format (Hebrew RTL, extracted with pdfplumber):
  Sample header : "{sid} :המגודה רפסמ {name} :המגודה רואית"
  Section header: line containing ICP SOIL | SVOC | VOC | TPH-DRO+ORO | Explosives | ICP
  Data line     : "[fn] CAS #: {cas} [{loq}] {unit} [X≤threshold] {result} [{n}/] {compound}"
  Unit wrap     : "substance" on its own line is joined to the preceding "mg/kg dry" line.

Analysis type mapping:
  ICP SOIL  + mg/kg → SOIL_METALS
  SVOC      + mg/kg → SOIL_SVOC
  VOC       + mg/kg → SOIL_VOC
  TPH-DRO   + mg/kg → SOIL_TPH
  Explosives        → SOIL_EXPLOSIVES
  ICP       + mg/L  → GW_METALS
  TPH       + mg/L  → GW_TPH
"""

from __future__ import annotations

import io
import re

from parsers.base import BaseParser


# ── Patterns ──────────────────────────────────────────────────────────────────

# Sample header (reversed RTL Hebrew): "{sid} :המגודה רפסמ {name} :המגודה רואית"
_SAMPLE_RE = re.compile(
    r"(?P<sid>\S+)\s+:המגודה רפסמ\s+(?P<sname>.+?)\s+:המגודה רואית"
)

# CAS data line — fields in order:
#   [footnote] CAS #: {cas} [{loq}] {unit} [X≤threshold] {result} [{n}/] {compound}
_CAS_RE = re.compile(
    r"CAS\s*#:\s*(?P<cas>[\w.+\-]+)"                          # CAS / pseudo-id
    r"(?:\s+(?P<loq>[\d.]+))?"                                 # optional LOQ
    r"\s+(?P<unit>%|mg/(?:kg(?:\s+dry(?:\s+substance)?)?|L))" # unit
    r"(?:\s+X[≤≥<>≠]\s*[\d.]+(?:\s*/\s*[\d.]+)?)?"           # optional threshold (discarded)
    r"\s+(?P<result>Not\s+Detected|<[\d.]+|[\d.]+)"            # result
    r"(?:\s+\d+/)?"                                            # optional sample ref (discarded)
    r"\s*(?P<compound>.*)$",                                   # rest = compound name
    re.IGNORECASE,
)

# Section keywords — checked in priority order (SVOC before VOC, ICP SOIL before ICP)
_SECTION_PATTERNS: list[tuple[re.Pattern, str]] = [
    (re.compile(r"ICP\s+SOIL",                    re.I), "ICP_SOIL"),
    (re.compile(r"Explosives?\s+and\s+propellant", re.I), "EXPLOSIVES"),
    (re.compile(r"TPH[-\s]*DRO",                  re.I), "TPH"),
    (re.compile(r"SVOC",                           re.I), "SVOC"),
    (re.compile(r"\bVOC\b",                        re.I), "VOC"),
    (re.compile(r"\bICP\b",                        re.I), "ICP"),
]

# Page footer — skip these lines
_PAGE_FOOTER_RE = re.compile(r"Page\s+\d+\s+of\s+\d+", re.I)


# ── Helpers ───────────────────────────────────────────────────────────────────

_HEBREW_RE = re.compile(r"[א-תיִ-פֿ]")


def _fix_rtl(text: str) -> str:
    """
    Correct RTL-reversed text extracted from a Hebrew PDF by pdfplumber.

    pdfplumber returns characters in visual (left-to-right) order.  For Hebrew
    text this means:
      • individual Hebrew words have their characters reversed
      • the word order across the whole string is also reversed

    Algorithm:
      1. If the string contains no Hebrew characters it is purely
         English/numeric — pdfplumber already gives these in the right order,
         so return unchanged.
      2. Otherwise reverse the list of whitespace-separated tokens, then
         reverse the character sequence of every token that contains Hebrew.
         English tokens (element symbols, abbreviations) keep their characters
         intact but move with the overall reversal.

    Examples
    --------
    'ףסכ (Ag)'   →  '(Ag) כסף'     (silver)
    'B1 עקרק'   →  'קרקע B1'       (soil borehole)
    'Benzene'    →  'Benzene'        (pure English — unchanged)
    """
    if not _HEBREW_RE.search(text):
        return text
    tokens = text.split()
    tokens.reverse()
    return " ".join(t[::-1] if _HEBREW_RE.search(t) else t for t in tokens)


def _analysis_type(section: str | None, unit: str) -> str:
    is_soil = "kg" in unit
    if section == "ICP_SOIL":
        return "SOIL_METALS"
    if section == "SVOC":
        return "SOIL_SVOC"
    if section == "VOC":
        return "SOIL_VOC"
    if section == "TPH":
        return "SOIL_TPH" if is_soil else "GW_TPH"
    if section == "EXPLOSIVES":
        return "SOIL_EXPLOSIVES"
    if section == "ICP":
        return "SOIL_METALS" if is_soil else "GW_METALS"
    return "SOIL_SVOC"


def _parse_result(raw: str, loq: float | None) -> tuple[float | None, str]:
    r = raw.strip()
    if r.lower().replace(" ", "") in ("notdetected", "nd", "n.d."):
        return loq, "ND"
    if r.startswith("<"):
        try:
            return float(r[1:]), "<LOQ"
        except ValueError:
            return loq, "<LOQ"
    try:
        return float(r), ""
    except ValueError:
        return None, ""


# ── Parser ────────────────────────────────────────────────────────────────────

class BactochemSoilParser(BaseParser):
    """
    Bactochem PDF report parser — handles soil and groundwater samples
    in a single file, routing each CAS record to the correct analysis type.
    """

    LAB_NAME = "בקטוכם"
    ANALYSIS_TYPES = [
        "SOIL_METALS", "SOIL_SVOC", "SOIL_VOC", "SOIL_TPH", "SOIL_EXPLOSIVES",
        "GW_METALS", "GW_TPH",
    ]

    def parse(self, file_obj: io.BytesIO) -> list[dict]:
        try:
            import pdfplumber
        except ImportError as exc:
            raise ImportError(
                "pdfplumber is required: pip install pdfplumber"
            ) from exc

        lines = _extract_lines(file_obj, pdfplumber)
        return _parse_lines(lines)


# ── Module-level parsing functions (testable without instantiation) ───────────

def _extract_lines(file_obj: io.BytesIO, pdfplumber) -> list[str]:
    """Extract and pre-process lines from the PDF."""
    raw: list[str] = []
    with pdfplumber.open(file_obj) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            raw.extend(text.splitlines())
    # "substance" appears on its own line after "mg/kg dry" lines due to PDF
    # word-wrap; skipping it is safe because the unit regex matches "mg/kg dry"
    # without the trailing word, and we normalize below.
    return [s for line in raw if (s := line.strip())]


def _parse_lines(lines: list[str]) -> list[dict]:
    records: list[dict] = []
    sample_id   = "Unknown"
    sample_name = "Unknown"
    section: str | None = None

    for line in lines:
        if _PAGE_FOOTER_RE.search(line):
            continue

        # ── Sample header? ──────────────────────────────────────────────────
        m = _SAMPLE_RE.search(line)
        if m:
            sample_id   = m.group("sid").strip()
            sample_name = _fix_rtl(m.group("sname").strip())
            section     = None
            continue

        # ── Section header? (only on lines that are not CAS data lines) ────
        if "CAS" not in line or "#:" not in line:
            for pat, key in _SECTION_PATTERNS:
                if pat.search(line):
                    section = key
                    break
            continue

        # ── CAS data line ───────────────────────────────────────────────────
        m2 = _CAS_RE.search(line)
        if not m2:
            continue

        cas      = m2.group("cas").strip()
        loq_raw  = m2.group("loq") or ""
        unit     = m2.group("unit").strip()
        res_raw  = m2.group("result").strip()
        compound = _fix_rtl((m2.group("compound") or "").strip())

        try:
            loq: float | None = float(loq_raw) if loq_raw else None
        except ValueError:
            loq = None

        value, flag = _parse_result(res_raw, loq)
        if value is None and not flag:
            continue

        # Normalize the soil-metals unit ("mg/kg dry" → "mg/kg dry substance")
        if unit == "mg/kg dry":
            unit = "mg/kg dry substance"

        records.append({
            "sample_id":     sample_name,
            "compound":      compound,
            "cas":           cas,
            "value":         value,
            "flag":          flag,
            "unit":          unit,
            "lod":           None,
            "loq":           loq,
            "analysis_type": _analysis_type(section, unit),
        })

    return records
