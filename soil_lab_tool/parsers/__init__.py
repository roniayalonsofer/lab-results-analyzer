"""
parsers/__init__.py
--------------------
Central registry that maps (lab_name, category) → parser class.
"""

import re as _re

from parsers.base import BaseParser

from parsers.soil_gas.alchem    import AlchemSoilGasParser
from parsers.soil_gas.kte       import KTESoilGasParser
from parsers.soil.alchem        import AlchemSoilParser
from parsers.alchem             import AlchemTPHPDFParser
from parsers.soil.kte           import KTESoilParser
from parsers.soil.kte_pr        import KTEPRParser
from parsers.soil.machon_haneft import MachonHaneftSoilParser
from parsers.soil.machon_energy import MachonEnergyParser, is_machon_energy_excel, is_machon_energy_pdf
from parsers.soil.als           import ALSSoilParser, ALSGrainSizeParser
from parsers.soil.bactochem     import BactochemSoilParser
from parsers.soil.xrf           import XRFSoilParser
from parsers.groundwater.kte        import KTEGroundwaterParser
from parsers.groundwater.bactochem  import BactochemGroundwaterParser
from parsers.groundwater.aminolab   import AminolabGroundwaterParser
from parsers.pfas.kte               import KTEPFASParser
from parsers.pfas.rj_lee            import RJLeePFASParser


_REGISTRY: dict[tuple[str, str], type[BaseParser]] = {
    ("alchem",        "soil_gas"):    AlchemSoilGasParser,
    ("kte",           "soil_gas"):    KTESoilGasParser,
    ("מכון האנרגיה",  "soil_gas"):   MachonEnergyParser,
    ("מכון האנרגיה",  "soil"):       MachonEnergyParser,
    ("machon energy", "soil_gas"):   MachonEnergyParser,
    ("machon energy", "soil"):       MachonEnergyParser,
    ("alchem",        "soil"):        AlchemSoilParser,
    ("alchem",        "soil_tph_pdf"): AlchemTPHPDFParser,
    ("kte",           "soil"):        KTESoilParser,
    ("kte",           "groundwater"): KTEGroundwaterParser,
    ("kte",           "pfas"):        KTEPFASParser,
    ("rj lee",        "pfas"):        RJLeePFASParser,
    ("rj lee",        "soil_pfas"):   RJLeePFASParser,
    ("kte",           "pr"):          KTEPRParser,
    ("מכון הנפט",    "soil"):        MachonHaneftSoilParser,
    ("machon haneft", "soil"):        MachonHaneftSoilParser,
    ("machon_haneft", "soil"):        MachonHaneftSoilParser,
    ("בקטוכם",       "groundwater"): BactochemGroundwaterParser,
    ("bactochem",     "groundwater"): BactochemGroundwaterParser,
    ("בקטוכם",       "soil"):        BactochemSoilParser,
    ("bactochem",     "soil"):        BactochemSoilParser,
    ("als",           "soil"):        ALSSoilParser,
    ("als",           "grain_size"):  ALSGrainSizeParser,
    ("aminolab",      "groundwater"): AminolabGroundwaterParser,
    ("אמינולאב",     "groundwater"): AminolabGroundwaterParser,
    ("xrf",           "soil"):        XRFSoilParser,
    ("אלכם",          "soil"):        XRFSoilParser,   # XRF method, lab = אלכם
    ("אלכם (xrf)",    "soil"):        XRFSoilParser,   # dropdown display key
}


# Sheet names Alchem uses: "<job_number>-VOC", "<job_number>-SVOC", etc.
_ALCHEM_SHEET_RE = _re.compile(r'^\d+-(?:VOC|SVOC|TPH|ICP|PH|METALS)$', _re.IGNORECASE)

# Alchem soil-gas variant: sheets are bare job-number integers ("40344", "38276")
_ALCHEM_SG_NUMERIC_RE = _re.compile(r'^\d+$')

# KTE TO-15 soil gas sheets: "<job_number>-TO-15-..." or contain "ppbv"
_KTE_SOIL_GAS_RE = _re.compile(r'TO-15|ppbv', _re.IGNORECASE)


def _is_alchem_excel(sheet_names: list[str]) -> bool:
    return any(_ALCHEM_SHEET_RE.match(s) for s in sheet_names)


def _is_alchem_soil_gas_numeric(sheet_names: list[str], file_bytes: bytes) -> bool:
    """Detect Alchem soil-gas files whose sheets are bare job-number integers.

    Primary check: scan rows 2-6 for a header row containing a compound-name
    column ('Compound Name' or 'Name'), 'CAS', and 'LOD [ug/m^3]'.
    Fallback: when row 1 is empty (layout shifted), look for 'Analysis Time:'
    in rows 2+ alongside LOD headers anywhere in the sheet.
    """
    numeric_sheets = [s for s in sheet_names if _ALCHEM_SG_NUMERIC_RE.match(s)]
    if not numeric_sheets:
        return False
    try:
        import io as _io
        import pandas as _pd
        xl = _pd.ExcelFile(_io.BytesIO(file_bytes))
        df = xl.parse(numeric_sheets[0], header=None, dtype=str, nrows=8).fillna("")
        if len(df) < 4:
            return False

        # Primary: scan rows 2–6 for the characteristic header row.
        # Accepts both "Compound Name" and bare "Name" as compound column.
        for ri in range(2, min(7, len(df))):
            row = [str(v).strip() for v in df.iloc[ri]]
            has_compound = any(v == "Name" or "Compound Name" in v for v in row)
            has_cas      = any(v.upper() == "CAS" for v in row)
            has_lod      = any("LOD" in v and "ug/m" in v for v in row)
            if has_compound and has_cas and has_lod:
                return True

        # Fallback: row 1 is empty → check rows 2+ for "Analysis Time:"
        # combined with LOD headers anywhere in the peeked rows.
        row0 = [str(v).strip() for v in df.iloc[0]]
        if all(v in ("", "nan") for v in row0):
            flat = " ".join(str(v) for v in df.values.flat)
            analysis_time_found = any(
                "Analysis Time" in "".join(str(v) for v in df.iloc[ri])
                for ri in range(1, len(df))
            )
            if analysis_time_found and "LOD" in flat and "ug/m" in flat:
                return True

        return False
    except Exception:
        return False


def _is_kte_soil_gas_excel(sheet_names: list[str]) -> bool:
    return any(_KTE_SOIL_GAS_RE.search(s) for s in sheet_names)


_MACHON_HANEFT_MARKERS = ("EPA 8270", "EPA 3550", "תעודת בדיקה", "גבול גילוי")


def _is_machon_haneft_excel(file_bytes: bytes) -> bool:
    """Return True if any of the first 15 rows in the first sheet contain a
    Machon HaNeft marker string."""
    try:
        import io
        import pandas as pd
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
        first_sheet = xl.sheet_names[0]
        df = xl.parse(first_sheet, header=None, dtype=str, nrows=15).fillna("")
        flat = " ".join(df.values.flatten().tolist())
        return any(m in flat for m in _MACHON_HANEFT_MARKERS)
    except Exception:
        return False


# "אמינולאב" reversed char-by-char — pdfplumber extracts Hebrew RTL text visually,
# so logical order "אמינולאב" arrives as visual order "בלאונימא".
_AMINOLAB_HE_REVERSED = "אמינולאב"[::-1]


def _is_aminolab_pdf(file_bytes: bytes) -> bool:
    """Return True if the PDF content identifies this as an Aminolab report.

    Scans the first 3 pages.  Checks:
      - "aminolab" (Latin, LTR — not reversed by pdfplumber)
      - "אמינולאב"  (Hebrew logical order — modern Unicode PDFs)
      - reversed form (Hebrew visual order — older PDFs where pdfplumber
        extracts characters left-to-right, reversing RTL words)
    """
    try:
        import io as _io, pdfplumber as _plumber
        with _plumber.open(_io.BytesIO(file_bytes)) as _pdf:
            for page in _pdf.pages[:3]:
                t = (page.extract_text() or "").lower()
                if "aminolab" in t:
                    return True
                if "אמינולאב" in t:
                    return True
                if _AMINOLAB_HE_REVERSED in t:
                    return True
    except Exception:
        pass
    return False


def _is_alchem_tph_pdf(file_bytes: bytes) -> bool:
    """Return True if the PDF is an Alchem TPH report (CID-font, needs pymupdf).

    The font uses Identity-H encoding with a +9 CID shift, so readable text
    like "DRO"/"N.D."/"<LOQ" appears as "3CFH"/"E%;%"/"CF;" in the text layer
    extracted by fitz.  All three must be present.
    """
    try:
        import fitz  # pymupdf
        doc = fitz.open(stream=file_bytes, filetype="pdf")
        full_text = "".join(page.get_text() for page in doc)
        doc.close()
        return "3CFH" in full_text and "E%;%" in full_text and "CF;" in full_text
    except Exception:
        pass
    return False


def _xlsx_sheet_names(file_bytes: bytes) -> list[str]:
    """Return sheet names from Excel bytes. Tries pandas first, zipfile XML fallback."""
    import io as _io
    try:
        import pandas as _pd
        return _pd.ExcelFile(_io.BytesIO(file_bytes)).sheet_names
    except Exception:
        pass
    try:
        import zipfile as _zf
        with _zf.ZipFile(_io.BytesIO(file_bytes)) as zf:
            if "xl/workbook.xml" in zf.namelist():
                wb = zf.read("xl/workbook.xml").decode("utf-8", errors="ignore")
                return _re.findall(r'<sheet\b[^>]*\bname="([^"]*)"', wb)
    except Exception:
        pass
    return []


_XRF_ELEMENTS = frozenset({
    "MO", "ZR", "SR", "U", "RB", "TH", "PB", "AU", "AS", "HG",
    "ZN", "CU", "NI", "CO", "FE", "MN", "CR", "V", "TI", "CA",
    "K", "S", "BA", "AG", "CD", "SB", "SE", "SN", "W", "Y", "NB",
})
_XRF_MIN_ELEMENTS = 6  # at least this many element columns to be considered XRF


def _is_xrf_tabular(file_bytes: bytes, is_csv: bool = False) -> bool:
    """Return True if the file looks like a wide-format XRF metals table."""
    try:
        import io as _io, pandas as _pd
        if is_csv:
            for enc in ("utf-8-sig", "utf-8", "latin-1"):
                try:
                    df = _pd.read_csv(_io.BytesIO(file_bytes), header=None,
                                      dtype=str, nrows=8, encoding=enc).fillna("")
                    break
                except Exception:
                    continue
            else:
                return False
        else:
            xl = _pd.ExcelFile(_io.BytesIO(file_bytes))
            df = xl.parse(xl.sheet_names[0], header=None, dtype=str,
                          nrows=8).fillna("")
        # Scan first 8 rows for a row with many element-symbol columns
        for ri in range(len(df)):
            row = [_re.sub(r"\s*[\(\[].*", "", str(v)).strip().upper()
                   for v in df.iloc[ri]]
            hit = sum(1 for v in row if v in _XRF_ELEMENTS)
            if hit >= _XRF_MIN_ELEMENTS:
                return True
    except Exception:
        pass
    return False


def _is_xrf_excel(file_bytes: bytes) -> bool:
    return _is_xrf_tabular(file_bytes, is_csv=False)


def auto_detect_lab(filename: str, file_bytes: bytes | None = None) -> str | None:
    """
    Attempt to identify the lab from filename and/or file content.
    Returns a lab key matching _REGISTRY (e.g. 'alchem', 'kte'), or None if uncertain.
    Content-based checks run before filename-based KTE check so that ALS/etc.
    files are not misidentified even when the filename has no useful hints.
    """
    n = filename.lower()

    try:
        import pdfplumber as _pl, io as _io2
        if file_bytes and filename.lower().endswith('.pdf'):
            with _pl.open(_io2.BytesIO(file_bytes)) as _p:
                _t = _p.pages[0].extract_text() or ''
                if 'םכ-לא' in _t:
                    return "alchem"
    except Exception:
        pass

    # Unambiguous filename hints (checked first — these are specific enough)
    if "xrf" in n:
        return "אלכם"
    if "alchem" in n:
        return "alchem"
    if any(k in n for k in ("aminolab", "אמינולאב")):
        return "aminolab"
    if any(k in n for k in ("בקטוכם", "bactochem")):
        return "בקטוכם"
    if any(k in n for k in ("מכון הנפט", "machon", "haneft", "neft")):
        return "מכון הנפט"
    if any(k in n for k in ("rjlg", "rj lee", "1633")):
        return "rj lee"

    # PDF content-based detection
    import logging
    logging.warning(f"DEBUG auto_detect_lab: filename={filename}, n={n}")
    if file_bytes is not None and n.endswith(".pdf"):
        if _is_aminolab_pdf(file_bytes):
            return "aminolab"
        if is_machon_energy_pdf(file_bytes):
            return "מכון האנרגיה"
        try:
            import pdfplumber, io as _io
            with pdfplumber.open(_io.BytesIO(file_bytes)) as _pdf:
                _text = (_pdf.pages[0].extract_text() or "") + \
                        (_pdf.pages[1].extract_text() if len(_pdf.pages) > 1 else "")
                if ("al-chem.com" in _text.lower() or
                        'םכ-לא' in _text or
                        'al-chem' in _text.lower()):
                    return "alchem"
        except Exception:
            pass
        if _is_alchem_tph_pdf(file_bytes):
            return "alchem"
        try:
            import io as _io, pdfplumber as _plumber
            with _plumber.open(_io.BytesIO(file_bytes)) as _pdf:
                first_text = (_pdf.pages[0].extract_text() or "").lower()
            if "bactochem" in first_text:
                return "בקטוכם"
        except Exception:
            pass

    # Content-based detection (runs BEFORE "kte" filename fallback so that
    # ALS files whose filenames happen to match "kte" patterns are caught here)
    if file_bytes is not None and (n.endswith(".xlsx") or n.endswith(".xls")):
        try:
            import io
            import pandas as pd
            xl = pd.ExcelFile(io.BytesIO(file_bytes))
            if _is_alchem_excel(xl.sheet_names):
                return "alchem"
            if _is_alchem_soil_gas_numeric(xl.sheet_names, file_bytes):
                return "alchem"
            if any("Client SOIL" in s for s in xl.sheet_names):
                return "als"
            if is_machon_energy_excel(file_bytes):
                return "מכון האנרגיה"
            if _is_machon_haneft_excel(file_bytes):
                return "מכון הנפט"
            if _is_kte_soil_gas_excel(xl.sheet_names):
                return "kte"
            if _is_xrf_excel(file_bytes):
                return "אלכם"
        except Exception:
            pass

    # CSV content check for XRF (run before KTE fallback)
    if file_bytes is not None and n.endswith(".csv"):
        if _is_xrf_tabular(file_bytes, is_csv=True):
            return "אלכם"

    # Filename fallback for KTE (after content checks)
    if any(k in n for k in ("kte", "excel_generic")):
        return "kte"

    return None


def get_parser(lab: str, category: str) -> BaseParser:
    key = (lab.strip().lower(), category.strip().lower())
    if key not in _REGISTRY:
        available = [f"({l}, {c})" for l, c in _REGISTRY]
        raise KeyError(
            f"No parser for lab='{lab}', category='{category}'.\n"
            f"Available: {available}"
        )
    return _REGISTRY[key]()


def list_parsers() -> list[dict]:
    return [{"lab": l, "category": c, "class": cls.__name__,
             "analysis_types": getattr(cls, "ANALYSIS_TYPES", [])}
            for (l, c), cls in _REGISTRY.items()]


def auto_detect_category(filename: str, file_bytes: bytes | None = None) -> str:
    """
    Guess analysis category from filename, and optionally peek at file content.
    Content-based Excel checks always run BEFORE filename-based checks so that
    ALS/Alchem/KTE files are not mis-detected by filename patterns (e.g. "pr*").
    """
    n = filename.lower()

    # RJ Lee Method 1633 PFAS files — unambiguous, check before content sniffing
    if "1633" in n:
        return "pfas"

    # ── PDF content-based detection ──────────────────────────────────────────────
    if file_bytes is not None and n.endswith(".pdf"):
        if _is_aminolab_pdf(file_bytes):
            return "groundwater"
        if _is_alchem_tph_pdf(file_bytes):
            return "soil_tph_pdf"
        try:
            import io as _io, pdfplumber as _plumber
            with _plumber.open(_io.BytesIO(file_bytes)) as _pdf:
                first_text = (_pdf.pages[0].extract_text() or "").lower()
            if "bactochem" in first_text:
                # Bactochem produces both soil PDFs (SVOC/ICP/mg/kg sections)
                # and groundwater PDFs (BTEX field params, mg/L).
                # Soil markers are distinctive; absent → treat as groundwater.
                _soil_markers = ("svoc", "icp soil", "tph-dro", "mg/kg")
                if any(m in first_text for m in _soil_markers):
                    return "soil"
                return "groundwater"
        except Exception:
            pass

    # ── Content-based detection for Excel (runs BEFORE any filename logic) ──────
    if file_bytes is not None and (n.endswith(".xlsx") or n.endswith(".xls")):
        sheet_names = _xlsx_sheet_names(file_bytes)

        # ALS: sheet name contains "Client SOIL"
        if any("Client SOIL" in s for s in sheet_names):
            try:
                import io, pandas as pd
                xl    = pd.ExcelFile(io.BytesIO(file_bytes))
                sheet = next(s for s in xl.sheet_names if "Client SOIL" in s)
                raw   = xl.parse(sheet, header=None, dtype=str).fillna("")

                # Locate header row and compound column (mirrors _parse_als_sheet)
                hdr_row_idx  = 12
                compound_col = 0
                for ri in range(8, min(20, len(raw))):
                    row_vals = [str(v).strip().upper() for v in raw.iloc[ri]]
                    if "LOR" in row_vals or "PARAMETER" in row_vals:
                        hdr_row_idx = ri
                        for ci, h in enumerate(row_vals):
                            if h in ("PARAMETER", "COMPOUND", "ANALYTE"):
                                compound_col = ci
                                break
                        break

                # grain_size only if majority of compounds are grain-size parameters
                compounds = [
                    str(raw.iloc[ri, compound_col]).strip()
                    for ri in range(hdr_row_idx + 1, len(raw))
                    if str(raw.iloc[ri, compound_col]).strip()
                    and str(raw.iloc[ri, compound_col]).strip().lower() not in ("nan", "parameter", "compound", "analyte")
                ]
                grain_count = sum(
                    1 for c in compounds
                    if "Fraction" in c or "Physical Parameters" in c
                )
                if compounds and grain_count > len(compounds) / 2:
                    return "grain_size"
            except Exception:
                pass
            return "soil"

        if _is_alchem_soil_gas_numeric(sheet_names, file_bytes):
            return "soil_gas"

        if _is_alchem_excel(sheet_names):
            return "soil"

        if is_machon_energy_excel(file_bytes):
            return "soil_gas"

        if _is_kte_soil_gas_excel(sheet_names):
            return "soil_gas"

    # ── Filename-based checks (run only after content checks) ────────────────────
    if "xrf" in n:
        return "soil"
    if "excel_generic" in n or n.startswith("pr"):
        return "pr"
    if any(k in n for k in ("pfas", "edd", "1633")):
        return "pfas"
    if any(k in n for k in ("soil_gas", "canister", "to-15", "to15", "ppbv")):
        return "soil_gas"
    if any(k in n for k in ("gw", "groundwater", "mei_tehom", "lowflow", "תלפיות")):
        return "groundwater"

    # KTE "EXCEL_GENERIC.XLS" uploads are often SpreadsheetML XML (not real .xls).
    if file_bytes is not None:
        head = file_bytes.lstrip()[:512]
        if head.startswith(b"<?xml") and b"urn:schemas-microsoft-com:office:spreadsheet" in file_bytes[:4096]:
            return "pr"

    # Content peek for CSV and remaining Excel cases (KTE analysis code detection)
    if file_bytes is not None and (n.endswith(".xlsx") or n.endswith(".xls") or n.endswith(".csv")):
        try:
            import io, pandas as pd
            if n.endswith(".csv"):
                df = pd.read_csv(io.BytesIO(file_bytes), encoding="utf-8-sig",
                                 header=None, nrows=6, dtype=str,
                                 names=list(range(30)), engine="python").fillna("")
            else:
                xl = pd.ExcelFile(io.BytesIO(file_bytes))
                df = xl.parse(xl.sheet_names[0], header=None, dtype=str,
                              nrows=6).fillna("")
            peek = " ".join(str(v) for v in df.values.flat).lower()
            if "canister number" in peek:
                return "soil_gas"
            if df.shape[0] >= 3:
                acode = str(df.iloc[2, 2]).strip().upper()
                if "PFAS" in acode:
                    return "pfas"
                if any(k in acode for k in ("WATER", "GW", "LOWFLOW")):
                    return "groundwater"
                if "SOIL" in acode or "TPH" in acode:
                    return "soil"
        except Exception:
            pass

    return "soil"

