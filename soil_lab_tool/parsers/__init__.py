"""
parsers/__init__.py
--------------------
Central registry that maps (lab_name, category) → parser class.
"""

from parsers.base import BaseParser

from parsers.soil_gas.alchem    import AlchemSoilGasParser
from parsers.soil.alchem        import AlchemSoilParser
from parsers.soil.kte           import KTESoilParser
from parsers.soil.kte_pr        import KTEPRParser
from parsers.soil.machon_haneft import MachonHaneftSoilParser
from parsers.groundwater.kte        import KTEGroundwaterParser
from parsers.groundwater.bactochem  import BactochemGroundwaterParser
from parsers.pfas.kte               import KTEPFASParser


_REGISTRY: dict[tuple[str, str], type[BaseParser]] = {
    ("alchem",        "soil_gas"):    AlchemSoilGasParser,
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
}


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
    if any(k in n for k in ("soil_gas", "canister", "to-15", "to15")):
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
