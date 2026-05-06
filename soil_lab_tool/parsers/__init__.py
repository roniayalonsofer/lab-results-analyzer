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
from parsers.soil.kte           import KTESoilParser
from parsers.soil.kte_pr        import KTEPRParser
from parsers.soil.machon_haneft import MachonHaneftSoilParser
from parsers.soil.als           import ALSSoilParser, ALSGrainSizeParser
from parsers.groundwater.kte        import KTEGroundwaterParser
from parsers.groundwater.bactochem  import BactochemGroundwaterParser
from parsers.pfas.kte               import KTEPFASParser


_REGISTRY: dict[tuple[str, str], type[BaseParser]] = {
    ("alchem",        "soil_gas"):    AlchemSoilGasParser,
    ("kte",           "soil_gas"):    KTESoilGasParser,
    ("alchem",        "soil"):        AlchemSoilParser,
    ("kte",           "soil"):        KTESoilParser,
    ("kte",           "groundwater"): KTEGroundwaterParser,
    ("kte",           "pfas"):        KTEPFASParser,
    ("kte",           "pr"):          KTEPRParser,
    ("מכון הנפט",    "soil"):        MachonHaneftSoilParser,
    ("machon haneft", "soil"):        MachonHaneftSoilParser,
    ("machon_haneft", "soil"):        MachonHaneftSoilParser,
    ("בקטוכם",       "groundwater"): BactochemGroundwaterParser,
    ("bactochem",     "groundwater"): BactochemGroundwaterParser,
    ("als",           "soil"):        ALSSoilParser,
    ("als",           "grain_size"):  ALSGrainSizeParser,
}


# Sheet names Alchem uses: "<job_number>-VOC", "<job_number>-SVOC", etc.
_ALCHEM_SHEET_RE = _re.compile(r'^\d+-(?:VOC|SVOC|TPH|ICP|PH|METALS)$', _re.IGNORECASE)

# KTE TO-15 soil gas sheets: "<job_number>-TO-15-..." or contain "ppbv"
_KTE_SOIL_GAS_RE = _re.compile(r'TO-15|ppbv', _re.IGNORECASE)


def _is_alchem_excel(sheet_names: list[str]) -> bool:
    return any(_ALCHEM_SHEET_RE.match(s) for s in sheet_names)


def _is_kte_soil_gas_excel(sheet_names: list[str]) -> bool:
    return any(_KTE_SOIL_GAS_RE.search(s) for s in sheet_names)


def auto_detect_lab(filename: str, file_bytes: bytes | None = None) -> str | None:
    """
    Attempt to identify the lab from filename and/or file content.
    Returns a lab key matching _REGISTRY (e.g. 'alchem', 'kte'), or None if uncertain.
    Content-based checks run before filename-based KTE check so that ALS/etc.
    files are not misidentified even when the filename has no useful hints.
    """
    n = filename.lower()

    # Unambiguous filename hints (checked first — these are specific enough)
    if "alchem" in n:
        return "alchem"
    if any(k in n for k in ("בקטוכם", "bactochem")):
        return "בקטוכם"
    if any(k in n for k in ("מכון הנפט", "machon", "haneft", "neft")):
        return "מכון הנפט"

    # Content-based detection (runs BEFORE "kte" filename fallback so that
    # ALS files whose filenames happen to match "kte" patterns are caught here)
    if file_bytes is not None and (n.endswith(".xlsx") or n.endswith(".xls")):
        try:
            import io
            import pandas as pd
            xl = pd.ExcelFile(io.BytesIO(file_bytes))
            if _is_alchem_excel(xl.sheet_names):
                return "alchem"
            if any("Client SOIL" in s for s in xl.sheet_names):
                return "als"
            if _is_kte_soil_gas_excel(xl.sheet_names):
                return "kte"
        except Exception:
            pass

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
    If file_bytes is supplied (for KTE XLSX/CSV), inspects the analysis code in
    the first data row to distinguish soil / groundwater / pfas.
    """
    n = filename.lower()
    if "excel_generic" in n or n.startswith("pr"):
        return "pr"
    if "pfas" in n:
        return "pfas"
    if any(k in n for k in ("soil_gas", "canister", "to-15", "to15", "ppbv")):
        return "soil_gas"
    if any(k in n for k in ("gw", "groundwater", "mei_tehom", "lowflow", "תלפיות")):
        return "groundwater"

    # KTE "EXCEL_GENERIC.XLS" uploads are often SpreadsheetML XML (not real .xls).
    # When the filename is generic (e.g., "upload.xls"), detect via content.
    if file_bytes is not None:
        head = file_bytes.lstrip()[:512]
        if head.startswith(b"<?xml") and b"urn:schemas-microsoft-com:office:spreadsheet" in file_bytes[:4096]:
            return "pr"

    # Peek at file content for format-level detection
    if file_bytes is not None and (n.endswith(".xlsx") or n.endswith(".xls") or n.endswith(".csv")):
        try:
            import io, pandas as pd
            if n.endswith(".csv"):
                df = pd.read_csv(io.BytesIO(file_bytes), encoding="utf-8-sig",
                                 header=None, nrows=6, dtype=str,
                                 names=list(range(30)), engine="python").fillna("")
            else:
                xl = pd.ExcelFile(io.BytesIO(file_bytes))

                # Alchem soil files: sheet names like "40752-VOC", "40752-SVOC", etc.
                if _is_alchem_excel(xl.sheet_names):
                    return "soil"

                # ALS files: "Client SOIL" sheet name
                if any("Client SOIL" in s for s in xl.sheet_names):
                    # Peek to distinguish grain-size from regular soil.
                    # Check first few columns since compound column may be 0 or 1.
                    try:
                        sheet = next(s for s in xl.sheet_names if "Client SOIL" in s)
                        peek = xl.parse(sheet, header=None, dtype=str,
                                        nrows=35).fillna("")
                        for ri in range(10, min(35, len(peek))):
                            for ci in range(min(4, peek.shape[1])):
                                cell = str(peek.iloc[ri, ci]).strip().lower()
                                if "fraction" in cell or "physical parameter" in cell:
                                    return "grain_size"
                    except Exception:
                        pass
                    return "soil"

                # KTE TO-15 soil gas
                if _is_kte_soil_gas_excel(xl.sheet_names):
                    return "soil_gas"

                df = xl.parse(xl.sheet_names[0], header=None, dtype=str,
                              nrows=6).fillna("")

            # Flatten all peeked text for keyword scanning
            peek = " ".join(str(v) for v in df.values.flat).lower()

            # Alchem soil-gas indicator: "Canister Number" row (unique to TO-15 format)
            if "canister number" in peek:
                return "soil_gas"

            # For KTE files: inspect analysis code in row 2, col 2
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
