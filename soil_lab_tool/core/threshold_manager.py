"""
core/threshold_manager.py
--------------------------
Loads environmental threshold values from Excel files and provides
CAS-based lookup across multiple threshold types.

Supported threshold keys
------------------------
From soil_vsl_tier1_v7_2024.xlsx  (sheet "VSL Tier1 2024"):
  VSL_SOIL          ← "Soil Direct Contact"   (mg/kg)
  TIER1_INDOOR_RES  ← "Indoor Residential"    (µg/m³)
  TIER1_OUTDOOR_RES ← "Outdoor Residential"   (µg/m³)
  TIER1_INDOOR_IND  ← "Indoor Industrial"     (µg/m³)
  TIER1_OUTDOOR_IND ← "Outdoor Industrial"    (µg/m³)
  GW                ← "Groundwater"           (mg/L)

From נספח לטבלת ערכי סף - PFAS.xlsx:
  PFAS_VSL          ← sheet " Soil VSL",       col [mg/kg]
  PFAS_TIER1_RES    ← sheet "Tier 1 Residential RBTL", first [mg/kg] col
  PFAS_TIER1_IND    ← sheet "Tier 1 - Industrial RBTL", first [mg/kg] col

From soil_vsl_v7_full.xlsx  (Tier 1 RBTL sheets) — soil vapor inhalation:
  GAS_INDOOR_RES    ← "Tier 1 Residential RBTL", Indoor Vapor col  (µg/m³)
  GAS_OUTDOOR_RES   ← "Tier 1 Residential RBTL", Outdoor Vapor col (µg/m³)
  GAS_INDOOR_IND    ← "Tier 1 - Industrial RBTL", Indoor Vapor col (µg/m³)
  GAS_OUTDOOR_IND   ← "Tier 1 - Industrial RBTL", Outdoor Vapor col(µg/m³)

From soil_vsl_v7_full.xlsx  (Tier 1 RBTL sheets) — soil direct contact
  with aquifer sensitivity & depth to groundwater (mg/kg):

  Col D (0-based 3): Very High sensitivity, all depths
  Col E (0-based 4): High/Medium sensitivity, 0-6 m depth
  Col G (0-based 6): High/Medium sensitivity, >6 m depth
  Col I (0-based 8): Low sensitivity

  TIER1_RES_SOIL_VH      ← Residential, Very High sensitivity
  TIER1_RES_SOIL_HM_0_6  ← Residential, High/Medium, 0-6 m
  TIER1_RES_SOIL_HM_6    ← Residential, High/Medium, >6 m
  TIER1_RES_SOIL_LOW     ← Residential, Low sensitivity
  TIER1_IND_SOIL_VH      ← Industrial,  Very High sensitivity
  TIER1_IND_SOIL_HM_0_6  ← Industrial,  High/Medium, 0-6 m
  TIER1_IND_SOIL_HM_6    ← Industrial,  High/Medium, >6 m
  TIER1_IND_SOIL_LOW     ← Industrial,  Low sensitivity
"""

from __future__ import annotations

import re
import pandas as pd
from core.cas_lookup import fuzzy_name_to_cas


# ── Main file column mapping ──────────────────────────────────────────
_MAIN_COL_MAP: dict[str, str] = {
    "VSL_SOIL":          "Soil Direct Contact",
    "TIER1_INDOOR_RES":  "Indoor Residential",
    "TIER1_OUTDOOR_RES": "Outdoor Residential",
    "TIER1_INDOOR_IND":  "Indoor Industrial",
    "TIER1_OUTDOOR_IND": "Outdoor Industrial",
    "GW":                "Groundwater",
}

# Display labels for threshold keys (Hebrew)
THRESHOLD_LABELS: dict[str, str] = {
    "VSL_SOIL":          "ערך סף (VSL)",
    "TIER1_INDOOR_RES":  "TIER1 גז פנים-מגורים",
    "TIER1_OUTDOOR_RES": "TIER1 גז חוץ-מגורים",
    "TIER1_INDOOR_IND":  "TIER1 גז פנים-תעשייה",
    "TIER1_OUTDOOR_IND": "TIER1 גז חוץ-תעשייה",
    "GW":                'ערך סף מי"ת',
    "PFAS_VSL":              "PFAS VSL",
    "PFAS_TIER1_RES":        "PFAS TIER1 מגורים",
    "PFAS_TIER1_IND":        "PFAS TIER1 תעשייה",
    "PFAS_TIER1_RES_VERY_HIGH": "PFAS TIER1 מגורים - רגישות גבוהה מאוד",
    "PFAS_TIER1_RES_0_6":      "PFAS TIER1 מגורים - רגישות גבוהה/בינונית, 0-6מ'",
    "PFAS_TIER1_RES_6PLUS":    "PFAS TIER1 מגורים - רגישות גבוהה/בינונית, >6מ'",
    "PFAS_TIER1_RES_NO_GW":    "PFAS TIER1 מגורים - רגישות נמוכה",
    "PFAS_TIER1_IND_VERY_HIGH": "PFAS TIER1 תעשייה - רגישות גבוהה מאוד",
    "PFAS_TIER1_IND_0_6":      "PFAS TIER1 תעשייה - רגישות גבוהה/בינונית, 0-6מ'",
    "PFAS_TIER1_IND_6PLUS":    "PFAS TIER1 תעשייה - רגישות גבוהה/בינונית, >6מ'",
    "PFAS_TIER1_IND_NO_GW":    "PFAS TIER1 תעשייה - רגישות נמוכה",
    # Soil-vapor RBTL keys (full V7 Tier1 RBTL sheets)
    "GAS_INDOOR_RES":  "גז קרקע — הגנה על אוויר פנים מבני (מגורים)",
    "GAS_OUTDOOR_RES": "ערך סף מגורים — חוץ",
    "GAS_INDOOR_IND":  "גז קרקע — הגנה על אוויר פנים מבני (תעשייה)",
    "GAS_OUTDOOR_IND": "ערך סף תעשייה — חוץ",
    "GAS_AMBIENT_RES": "אוויר תוך מבני — Ambient Air (מגורים)",
    "GAS_AMBIENT_IND": "אוויר תוך מבני — Ambient Air (תעשייה)",
    # Soil direct-contact RBTL keys with aquifer sensitivity & depth
    "TIER1_RES_SOIL_VH":     "TIER1 מגורים - רגיש מאוד",
    "TIER1_RES_SOIL_HM_0_6": "TIER1 מגורים - רגיש/בינוני, 0-6מ'",
    "TIER1_RES_SOIL_HM_6":   "TIER1 מגורים - רגיש/בינוני, >6מ'",
    "TIER1_RES_SOIL_LOW":    "TIER1 מגורים - לא רגיש",
    "TIER1_IND_SOIL_VH":     "TIER1 תעשייה - רגיש מאוד",
    "TIER1_IND_SOIL_HM_0_6": "TIER1 תעשייה - רגיש/בינוני, 0-6מ'",
    "TIER1_IND_SOIL_HM_6":   "TIER1 תעשייה - רגיש/בינוני, >6מ'",
    "TIER1_IND_SOIL_LOW":    "TIER1 תעשייה - לא רגיש",
}

# Explicit per-CAS VSL overrides for compounds that use pseudo-CAS identifiers.
# DRO, ORO, DRO-ORO: VSL = 350 mg/kg (Israeli environmental standard).
_CAS_VSL_OVERRIDES: dict[str, float | None] = {
    "DRO": 350.0,
    "ORO": 350.0,
}

_NO_VSL_CAS: frozenset[str] = frozenset({
    "7440-70-2",   # Calcium
    "7440-21-3",   # Silicon
    "7440-23-5",   # Sodium
    "7440-09-7",   # Potassium
    "7439-95-4",   # Magnesium
    "7723-14-0",   # Phosphorus
    "7704-34-9",   # Sulphur
    "7440-42-8",   # Boron
    "7440-24-6",   # Strontium
    "7550-45-0",   # Titanium
    "13494-80-9",  # Tellurium
    "7440-67-7",   # Zirconium
    "7440-69-9",   # Bismuth
    "7439-93-2",   # Lithium
})

# VSL_SOIL thresholds for TPH pseudo-CAS identifiers (mg/kg).
_TPH_STANDARDS: dict[str, float] = {
    "DRO":     350.0,
    "ORO":     350.0,
    "DRO-ORO": 350.0,
}

# Israeli drinking water standards for metals (mg/L).
# Source: Israeli Drinking Water Regulations.
_GW_METALS_STANDARDS: dict[str, float] = {
    "7440-38-2": 0.01,    # Arsenic
    "7440-39-3": 0.7,     # Barium
    "7440-43-9": 0.003,   # Cadmium
    "7440-47-3": 0.05,    # Chromium
    "7440-50-8": 1.0,     # Copper
    "7439-89-6": 0.3,     # Iron
    "7439-92-1": 0.01,    # Lead
    "7439-96-5": 0.5,     # Manganese
    "7439-97-6": 0.001,   # Mercury
    "7440-02-0": 0.02,    # Nickel
    "7782-49-2": 0.01,    # Selenium
    "7440-22-4": 0.1,     # Silver
    "7440-66-6": 3.0,     # Zinc
    "7429-90-5": 0.2,     # Aluminium
    "7440-36-0": 0.005,   # Antimony
    "7440-41-7": 0.004,   # Beryllium
    "7440-48-4": 0.05,    # Cobalt
    "7439-98-7": 0.07,    # Molybdenum
    "7440-28-0": 0.002,   # Thallium
    "7440-62-2": 0.05,    # Vanadium
}

# All Tier-1 soil direct-contact keys (used to declare valid keys per atype)
_SOIL_TIER1_KEYS: list[str] = [
    "TIER1_RES_SOIL_VH",    "TIER1_RES_SOIL_HM_0_6",
    "TIER1_RES_SOIL_HM_6",  "TIER1_RES_SOIL_LOW",
    "TIER1_IND_SOIL_VH",    "TIER1_IND_SOIL_HM_0_6",
    "TIER1_IND_SOIL_HM_6",  "TIER1_IND_SOIL_LOW",
]

# Which threshold keys apply to which analysis type
ANALYSIS_THRESHOLDS: dict[str, list[str]] = {
    # Soil gas: soil-vapor RBTL keys from full V7 Tier1 RBTL sheets
    "SOIL_GAS_VOC": ["GAS_INDOOR_RES",  "GAS_OUTDOOR_RES",
                     "GAS_INDOOR_IND",  "GAS_OUTDOOR_IND",
                     "GAS_AMBIENT_RES", "GAS_AMBIENT_IND"],
    # Soil types: VSL direct-contact + all aquifer-sensitivity variants
    "SOIL_VOC":       ["VSL_SOIL"] + _SOIL_TIER1_KEYS,
    "SOIL_SVOC":      ["VSL_SOIL"] + _SOIL_TIER1_KEYS,
    "SOIL_MBTEX":     ["VSL_SOIL"] + _SOIL_TIER1_KEYS,
    "SOIL_TPH":       ["VSL_SOIL"] + _SOIL_TIER1_KEYS,
    "SOIL_METALS":    ["VSL_SOIL"] + _SOIL_TIER1_KEYS,
    # Combined sheets (same threshold keys as SOIL_VOC/SOIL_TPH)
    "SOIL_TPH_VOC":   ["VSL_SOIL"] + _SOIL_TIER1_KEYS,
    "SOIL_TPH_MBTEX": ["VSL_SOIL"] + _SOIL_TIER1_KEYS,
    "SOIL_PFAS":    ["PFAS_VSL",
                    "PFAS_TIER1_RES", "PFAS_TIER1_IND",
                    "PFAS_TIER1_RES_VERY_HIGH", "PFAS_TIER1_RES_0_6", "PFAS_TIER1_RES_6PLUS", "PFAS_TIER1_RES_NO_GW",
                    "PFAS_TIER1_IND_VERY_HIGH", "PFAS_TIER1_IND_0_6", "PFAS_TIER1_IND_6PLUS", "PFAS_TIER1_IND_NO_GW"],
    "GW_VOC":           ["GW"],
    "GW_SVOC":          ["GW"],
    "GW_METALS":        ["GW"],
    "GW_PFAS":          [],
    "LOWFLOW":          [],
    "SOIL_GRAIN_SIZE":  [],
}


class ThresholdManager:
    """Manages environmental threshold values from one or more Excel files."""

    def __init__(
        self,
        main_path: str,
        pfas_path: str | None = None,
        vsl_full_path: str | None = None,
    ):
        self._main: pd.DataFrame = self._load_main(main_path)
        # Full V7 file: 800+ compounds including metals & TPH; used for VSL_SOIL
        self._vsl_full: pd.DataFrame | None = None
        # Soil-vapor RBTL thresholds (from Tier1 Residential/Industrial sheets in full V7 file)
        self._rbtl: dict[str, pd.DataFrame] = {}
        if vsl_full_path:
            self._vsl_full = self._load_vsl_full(vsl_full_path)
            self._rbtl    = self._load_tier1_rbtl(vsl_full_path)
        self._pfas: dict[str, pd.DataFrame] = {}
        if pfas_path:
            self._pfas = self._load_pfas(pfas_path)

    # ------------------------------------------------------------------
    # Loaders
    # ------------------------------------------------------------------
    @staticmethod
    def _load_main(path: str) -> pd.DataFrame:
        df = pd.read_excel(path, dtype=str).fillna("")
        df.columns = [c.strip() for c in df.columns]
        # Normalise CAS column name
        for col in df.columns:
            if "cas" in col.lower():
                df = df.rename(columns={col: "CAS No."})
                break
        df["CAS No."] = df["CAS No."].str.strip()
        return df

    @staticmethod
    def _load_vsl_full(path: str) -> pd.DataFrame | None:
        """
        Load the full V7 threshold file (ערכי סף מעודכנים - גרסה 7).

        Sheet: " Soil VSL"  |  header row: 2  |  value column: "[mg/kg]"
        Returns a DataFrame with columns normalised to match _MAIN_COL_MAP
        ('Soil Direct Contact', 'CAS No.').
        """
        try:
            df = pd.read_excel(
                path, sheet_name=" Soil VSL", header=2, dtype=str
            ).fillna("")
            df.columns = [c.strip() for c in df.columns]
            # Rename value column to match _MAIN_COL_MAP key
            if "[mg/kg]" in df.columns:
                df = df.rename(columns={"[mg/kg]": "Soil Direct Contact"})
            # Normalise CAS column
            for col in df.columns:
                if col.lower().startswith("cas"):
                    df = df.rename(columns={col: "CAS No."})
                    break
            df["CAS No."] = df["CAS No."].str.strip()
            return df
        except Exception:
            return None

    @staticmethod
    def _load_tier1_rbtl(path: str) -> dict[str, pd.DataFrame]:
        """
        Load Tier 1 RBTL thresholds from the full V7 file.

        Reads two sheets:
          'Tier 1 Residential RBTL'   → prefix 'res'
          'Tier 1 - Industrial RBTL'  → prefix 'ind'

        Each sheet yields six DataFrames (CAS No. | value):
          Soil vapor (µg/m³):
            {prefix}_indoor        — Soil Vapor Indoor inhalation
            {prefix}_outdoor       — Soil Vapor Outdoor inhalation
          Soil direct contact (mg/kg) with aquifer sensitivity / depth:
            {prefix}_soil_vh       — Very High sensitivity (all depths) col D
            {prefix}_soil_hm_0_6   — High/Medium sensitivity, 0-6 m    col E
            {prefix}_soil_hm_6     — High/Medium sensitivity, >6 m     col G
            {prefix}_soil_low      — Low sensitivity                   col I

        Column positions are discovered by scanning the multi-row header
        (first 8 rows) for keyword patterns; fixed fallbacks ensure
        robustness against minor layout changes in future file versions.
        """
        result: dict[str, pd.DataFrame] = {}
        _CAS_RE = re.compile(r"^\d+-\d+-\d+$")
        # Broader pattern: also matches pseudo-CAS like "C10-C40" (starts with letter)
        _CAS_BROAD_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9\-]+$")

        sheet_map = {
            "Tier 1 Residential RBTL":   "res",
            "Tier 1 - Industrial RBTL":  "ind",
        }
        try:
            xl = pd.ExcelFile(path)
        except Exception:
            return result

        for sheet_name, prefix in sheet_map.items():
            if sheet_name not in xl.sheet_names:
                continue
            try:
                raw = xl.parse(sheet_name, header=None, dtype=str).fillna("")

                # ── Scan header rows (0-7) for column positions ───────────────
                sv_indoor_col:    int | None = None
                sv_outdoor_col:   int | None = None
                ambient_air_col:  int | None = None
                soil_vh_col:      int | None = None
                soil_hm_0_6_col:  int | None = None
                soil_hm_6_col:    int | None = None
                soil_low_col:     int | None = None

                for ri in range(min(8, len(raw))):
                    for ci, v in enumerate(raw.iloc[ri].values):
                        vs = str(v).strip().lower()
                        if not vs or ci < 2:
                            continue
                        # Soil vapor columns (contain both "soil vapor" + direction)
                        if "soil vapor" in vs and "indoor" in vs and sv_indoor_col is None:
                            sv_indoor_col = ci
                        elif "soil vapor" in vs and "outdoor" in vs and sv_outdoor_col is None:
                            sv_outdoor_col = ci
                        # Ambient Air column (col 10 in V7; row 2 header "Ambient Air")
                        if "ambient air" in vs and "soil vapor" not in vs and ambient_air_col is None:
                            ambient_air_col = ci
                        # Soil direct contact: sensitivity / depth keywords
                        if "very high" in vs and soil_vh_col is None:
                            soil_vh_col = ci
                        # Match both "0-6" and "0 - 6" (some V7 files use spaces around dash)
                        if ("0-6" in vs or "0 - 6" in vs) and soil_hm_0_6_col is None:
                            soil_hm_0_6_col = ci
                        if ">6" in vs and soil_hm_6_col is None:
                            soil_hm_6_col = ci
                        if "low" in vs and "sensit" in vs and soil_low_col is None:
                            soil_low_col = ci

                # Fallback to known V7 column positions when scanning finds nothing
                ambient_air_col = ambient_air_col if ambient_air_col is not None else 10
                soil_vh_col     = soil_vh_col     if soil_vh_col     is not None else 3
                soil_hm_0_6_col = soil_hm_0_6_col if soil_hm_0_6_col is not None else 4
                soil_hm_6_col   = soil_hm_6_col   if soil_hm_6_col   is not None else 6
                soil_low_col    = soil_low_col    if soil_low_col    is not None else 8

                # ── Find first data row (col 1 matches CAS pattern) ──────────
                data_start: int | None = None
                for ri in range(len(raw)):
                    if _CAS_RE.match(str(raw.iloc[ri, 1]).strip()):
                        data_start = ri
                        break
                if data_start is None:
                    continue

                data = raw.iloc[data_start:].reset_index(drop=True)
                cas_series = data.iloc[:, 1].str.strip()

                # ── Build one DataFrame per column type ───────────────────────
                col_specs = [
                    (sv_indoor_col,   "indoor"),
                    (sv_outdoor_col,   "outdoor"),
                    (ambient_air_col,  "ambient_air"),
                    (soil_vh_col,     "soil_vh"),
                    (soil_hm_0_6_col, "soil_hm_0_6"),
                    (soil_hm_6_col,   "soil_hm_6"),
                    (soil_low_col,    "soil_low"),
                ]
                name_series = data.iloc[:, 0].str.strip()
                for col_idx, key_suffix in col_specs:
                    if col_idx is None or col_idx >= data.shape[1]:
                        continue
                    full_key = f"{prefix}_{key_suffix}"
                    df = pd.DataFrame({
                        "CAS No.": cas_series,
                        "name":    name_series,
                        "value":   data.iloc[:, col_idx],
                    })
                    df = df[df["CAS No."].apply(lambda x: bool(_CAS_BROAD_RE.match(str(x))) if x else False)].copy()
                    result[full_key] = df

            except Exception:
                continue

        return result

    @staticmethod
    def _load_pfas(path: str) -> dict[str, pd.DataFrame]:
        result: dict[str, pd.DataFrame] = {}
        xl = pd.ExcelFile(path)

        for sheet in xl.sheet_names:
            key = sheet.strip().lower()

            if "vsl" in key:
                # Header on row 2 (0-indexed)
                df = xl.parse(sheet, header=2, dtype=str).fillna("")
                df.columns = [c.strip() for c in df.columns]
                result["vsl"] = df
            elif "residential" in key or "res" in key:
                df = xl.parse(sheet, header=5, dtype=str).fillna("")
                df.columns = [c.strip() for c in df.columns]
                result["tier1_res"] = df
            elif "industrial" in key or "ind" in key:
                df = xl.parse(sheet, header=5, dtype=str).fillna("")
                df.columns = [c.strip() for c in df.columns]
                result["tier1_ind"] = df

        return result

    @staticmethod
    def _extract_pfas_tier1_cols(
        df: pd.DataFrame, prefix: str, result: dict
    ) -> None:
        """Extract the 3 depth/aquifer value columns from a Tier1 PFAS sheet.

        Pandas renames duplicate '[mg/kg]' headers as '[mg/kg]', '[mg/kg].1',
        '[mg/kg].2' for the 0-6m, >6m, and B-1/C columns respectively.
        Stores pre-extracted DataFrames (CAS No. | value) into *result*.
        """
        cas_col = next((c for c in df.columns if "cas" in c.lower()), None)
        if cas_col is None:
            return
        mg_cols = [c for c in df.columns if c.startswith("[mg/kg]")]
        suffixes = ["0_6", "6plus", "no_gw"]
        for col_name, suffix in zip(mg_cols, suffixes):
            key = f"tier1_{prefix}_{suffix}"
            result[key] = pd.DataFrame({
                "CAS No.": df[cas_col].str.strip(),
                "value":   df[col_name],
            })

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def get_threshold(self, cas: str, threshold_key: str) -> float | None:
        """
        Return the threshold value for a CAS number and threshold key.

        Parameters
        ----------
        cas : str
            CAS number, e.g. '71-43-2'
        threshold_key : str
            One of the keys defined in _MAIN_COL_MAP or 'PFAS_*'

        Returns
        -------
        float | None
        """
        cas = str(cas).strip() if cas is not None else ""
        if not cas or cas == "None":
            return None

        if threshold_key == "VSL_SOIL" and cas in _NO_VSL_CAS:
            return None

        if threshold_key == "VSL_SOIL" and cas.upper() in _TPH_STANDARDS:
            return _TPH_STANDARDS[cas.upper()]

        if cas.upper() in ("DRO", "ORO", "DRO-ORO") and threshold_key in (
            "TIER1_IND_SOIL_VH", "TIER1_IND_SOIL_HM_0_6",
            "TIER1_IND_SOIL_HM_6", "TIER1_IND_SOIL_LOW",
        ):
            return 1280.0
        if cas.upper() in ("DRO", "ORO", "DRO-ORO") and threshold_key in (
            "TIER1_RES_SOIL_VH", "TIER1_RES_SOIL_HM_0_6",
            "TIER1_RES_SOIL_HM_6", "TIER1_RES_SOIL_LOW",
        ):
            return 350.0

        if threshold_key == "GW" and cas in _GW_METALS_STANDARDS:
            return _GW_METALS_STANDARDS[cas]

        if threshold_key in _MAIN_COL_MAP:
            return self._lookup_main(cas, _MAIN_COL_MAP[threshold_key])
        if threshold_key == "PFAS_VSL":
            return self._lookup_pfas("vsl", cas)
        if threshold_key == "PFAS_TIER1_RES":
            return self._lookup_pfas("tier1_res", cas)
        if threshold_key == "PFAS_TIER1_IND":
            return self._lookup_pfas("tier1_ind", cas)
        _pfas_direct_map = {
            "PFAS_TIER1_RES_VERY_HIGH": "tier1_res_very_high",
            "PFAS_TIER1_RES_0_6":       "tier1_res_0_6",
            "PFAS_TIER1_RES_6PLUS":     "tier1_res_6plus",
            "PFAS_TIER1_RES_NO_GW":     "tier1_res_no_gw",
            "PFAS_TIER1_IND_VERY_HIGH": "tier1_ind_very_high",
            "PFAS_TIER1_IND_0_6":       "tier1_ind_0_6",
            "PFAS_TIER1_IND_6PLUS":     "tier1_ind_6plus",
            "PFAS_TIER1_IND_NO_GW":     "tier1_ind_no_gw",
        }
        if threshold_key in _pfas_direct_map:
            return self._lookup_pfas_direct(_pfas_direct_map[threshold_key], cas)
        # RBTL keys — soil vapor + soil direct-contact (from full V7 Tier1 sheets)
        _rbtl_map = {
            # Soil vapor inhalation — indoor (col 12) and outdoor
            "GAS_INDOOR_RES":         "res_indoor",
            "GAS_OUTDOOR_RES":        "res_outdoor",
            "GAS_INDOOR_IND":         "ind_indoor",
            "GAS_OUTDOOR_IND":        "ind_outdoor",
            # Ambient Air (col 10 in Tier 1 Residential/Industrial RBTL)
            "GAS_AMBIENT_RES":        "res_ambient_air",
            "GAS_AMBIENT_IND":        "ind_ambient_air",
            # Soil direct contact — residential, with aquifer sensitivity & depth
            "TIER1_RES_SOIL_VH":      "res_soil_vh",
            "TIER1_RES_SOIL_HM_0_6":  "res_soil_hm_0_6",
            "TIER1_RES_SOIL_HM_6":    "res_soil_hm_6",
            "TIER1_RES_SOIL_LOW":     "res_soil_low",
            # Soil direct contact — industrial, with aquifer sensitivity & depth
            "TIER1_IND_SOIL_VH":      "ind_soil_vh",
            "TIER1_IND_SOIL_HM_0_6":  "ind_soil_hm_0_6",
            "TIER1_IND_SOIL_HM_6":    "ind_soil_hm_6",
            "TIER1_IND_SOIL_LOW":     "ind_soil_low",
        }
        if threshold_key in _rbtl_map:
            return self._lookup_rbtl(_rbtl_map[threshold_key], cas)
        return None

    # Compound names/CAS values that map to the VSL entry "TPH - DRO + ORO (Tier 1)".
    # "oro" is handled via _CAS_VSL_OVERRIDES and must not alias to C10-C40.
    _TPH_ALIASES: frozenset[str] = frozenset({"tph", "dro", "tph - dro + oro"})

    def get_threshold_with_name(self, cas: str, threshold_key: str,
                                compound_name: str = "") -> float | None:
        """
        Like get_threshold but falls back to compound name lookup when CAS fails.
        The compound name column in the threshold file is 'chimical' (or similar).
        """
        # Explicit pseudo-CAS overrides (DRO=350, ORO=None) take priority.
        cas_upper = str(cas).strip().upper() if cas else ""
        if cas_upper in _CAS_VSL_OVERRIDES:
            return _CAS_VSL_OVERRIDES[cas_upper] if "VSL" in threshold_key else None

        # TPH/DRO/ORO compounds return 350 when matched by name.
        TPH_COMPOUNDS = {"tph", "dro", "oro"}
        if compound_name and compound_name.strip().lower() in TPH_COMPOUNDS:
            return 350.0

        # TPH sub-fractions (DRO, TPH) share a single VSL entry keyed on C10-C40.
        # ORO is intentionally excluded so CAS "ORO" returns None rather than 350.
        name_lo = str(compound_name).strip().lower() if compound_name else ""
        cas_lo  = str(cas).strip().lower() if cas else ""
        if name_lo in self._TPH_ALIASES or cas_lo in self._TPH_ALIASES:
            cas = "C10-C40"

        # Elements without VSL — return None immediately, skip all fallbacks
        if threshold_key == "VSL_SOIL" and str(cas).strip() in _NO_VSL_CAS:
            return None

        val = self.get_threshold(cas, threshold_key)
        # Fuzzy CAS resolve: if direct CAS lookup missed, try matching compound name
        # to a known CAS via cas_lookup, then re-run the threshold lookup.
        if val is None and compound_name:
            _fuzzy_cas = fuzzy_name_to_cas(compound_name)
            if _fuzzy_cas and _fuzzy_cas != cas:
                val = self.get_threshold(_fuzzy_cas, threshold_key)
        if val is not None or not compound_name:
            return val
        # CAS lookup failed — try by compound name
        name = str(compound_name).strip()
        _rbtl_map = {
            "GAS_INDOOR_RES":  "res_indoor",  "GAS_OUTDOOR_RES": "res_outdoor",
            "GAS_INDOOR_IND":  "ind_indoor",  "GAS_OUTDOOR_IND": "ind_outdoor",
            "GAS_AMBIENT_RES": "res_ambient_air", "GAS_AMBIENT_IND": "ind_ambient_air",
            "TIER1_RES_SOIL_VH": "res_soil_vh", "TIER1_RES_SOIL_HM_0_6": "res_soil_hm_0_6",
            "TIER1_RES_SOIL_HM_6": "res_soil_hm_6", "TIER1_RES_SOIL_LOW": "res_soil_low",
            "TIER1_IND_SOIL_VH": "ind_soil_vh", "TIER1_IND_SOIL_HM_0_6": "ind_soil_hm_0_6",
            "TIER1_IND_SOIL_HM_6": "ind_soil_hm_6", "TIER1_IND_SOIL_LOW": "ind_soil_low",
        }
        if threshold_key in _rbtl_map:
            return self._lookup_rbtl_by_name(_rbtl_map[threshold_key], name)
        if threshold_key in _MAIN_COL_MAP:
            if str(cas).strip() in _NO_VSL_CAS:
                return None
            return self._lookup_main_by_name(name, _MAIN_COL_MAP[threshold_key])
        _pfas_sheet_map = {
            "PFAS_VSL":       "vsl",
            "PFAS_TIER1_RES": "tier1_res",
            "PFAS_TIER1_IND": "tier1_ind",
        }
        if threshold_key in _pfas_sheet_map:
            return self._lookup_pfas_by_name(_pfas_sheet_map[threshold_key], name)
        _pfas_direct_map = {
            "PFAS_TIER1_RES_VERY_HIGH": "tier1_res_very_high",
            "PFAS_TIER1_RES_0_6":       "tier1_res_0_6",
            "PFAS_TIER1_RES_6PLUS":     "tier1_res_6plus",
            "PFAS_TIER1_RES_NO_GW":     "tier1_res_no_gw",
            "PFAS_TIER1_IND_VERY_HIGH": "tier1_ind_very_high",
            "PFAS_TIER1_IND_0_6":       "tier1_ind_0_6",
            "PFAS_TIER1_IND_6PLUS":     "tier1_ind_6plus",
            "PFAS_TIER1_IND_NO_GW":     "tier1_ind_no_gw",
        }
        if threshold_key in _pfas_direct_map:
            return self._lookup_pfas_direct_by_name(
                _pfas_direct_map[threshold_key], name
            )
        return None

    def get_cas_by_name(self, compound_name: str) -> str | None:
        """Return the CAS number for a compound looked up by name.

        Searches _vsl_full first (800+ compounds), then _main.
        Normalizations applied before matching:
          - backticks → apostrophes
          - parens → brackets (PAH notation: Benz(a)anthracene → Benz[a]anthracene)
          - dot-as-comma between digits (2.6-Dinitrotoluene → also tries 2,6- variant)
        Falls back to a partial/contains match as a last resort.
        """
        # ── Normalize ────────────────────────────────────────────────────────
        norm = compound_name.replace('`', "'")
        norm = re.sub(r'\(([a-zA-Z0-9,]+)\)', r'[\1]', norm)

        # Build deduplicated candidate list across all input variants
        raw_names = [norm]
        dot_comma = re.sub(r'(\d)\.(\d)', r'\1,\2', norm)
        if dot_comma != norm:
            raw_names.append(dot_comma)
        if compound_name != norm:
            raw_names.append(compound_name)

        seen: set[str] = set()
        candidates: list[str] = []
        for name in raw_names:
            for v in self._name_variants(name):
                key = v.lower()
                if key not in seen:
                    seen.add(key)
                    candidates.append(v)

        # ── Prepare df info ───────────────────────────────────────────────────
        df_info = []
        for df in ([self._vsl_full] if self._vsl_full is not None else []) + [self._main]:
            if df is None:
                continue
            name_col = next((c for c in df.columns if any(k in c.lower() for k in ("name", "compound", "chemical"))), None)
            cas_col  = next((c for c in df.columns if "cas" in c.lower()), None)
            if name_col and cas_col:
                df_info.append((df, cas_col, df[name_col].str.strip().str.lower()))

        # ── Pass 1: exact (case-insensitive) match ────────────────────────────
        for df, cas_col, col_lower in df_info:
            for variant in candidates:
                row = df[col_lower == variant.lower()]
                if not row.empty:
                    val = str(row.iloc[0][cas_col]).strip()
                    if val and val.lower() != "nan":
                        return val

        # ── Pass 2: partial / contains match ─────────────────────────────────
        # VSL name must be no more than 20% longer than input to prevent
        # "Calcium" from matching "Calcium Cyanide"
        for df, cas_col, col_lower in df_info:
            name_col = next((c for c in df.columns if any(k in c.lower() for k in ("name", "compound", "chemical"))), None)
            if not name_col:
                continue
            for variant in candidates:
                if len(variant) < 4:
                    continue
                try:
                    mask = col_lower.str.contains(re.escape(variant.lower()), na=False)
                    rows = df[mask]
                except Exception:
                    continue
                for _, r in rows.iterrows():
                    vsl_name = str(r[name_col]).strip()
                    if len(vsl_name) <= len(variant) * 1.2:
                        val = str(r[cas_col]).strip()
                        if val and val.lower() != "nan":
                            return val

        return None

    def get_thresholds_for_analysis(
        self, cas: str, analysis_type: str
    ) -> dict[str, float | None]:
        """Return all applicable thresholds for a given analysis type."""
        keys = ANALYSIS_THRESHOLDS.get(analysis_type, [])
        return {k: self.get_threshold(cas, k) for k in keys}

    def threshold_label(self, threshold_key: str) -> str:
        """Return the Hebrew display label for a threshold key."""
        return THRESHOLD_LABELS.get(threshold_key, threshold_key)

    def available_keys(self) -> list[str]:
        keys = list(_MAIN_COL_MAP.keys())
        if self._pfas:
            keys += [
                "PFAS_VSL",
                "PFAS_TIER1_RES", "PFAS_TIER1_IND",
                "PFAS_TIER1_RES_VERY_HIGH", "PFAS_TIER1_RES_0_6", "PFAS_TIER1_RES_6PLUS", "PFAS_TIER1_RES_NO_GW",
                "PFAS_TIER1_IND_VERY_HIGH", "PFAS_TIER1_IND_0_6", "PFAS_TIER1_IND_6PLUS", "PFAS_TIER1_IND_NO_GW",
            ]
        return keys

    @property
    def has_full_vsl(self) -> bool:
        return self._vsl_full is not None

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------
    def _lookup_rbtl(self, key: str, cas: str) -> float | None:
        df = self._rbtl.get(key)
        if df is None:
            return None
        row = df[df["CAS No."] == cas]
        if row.empty:
            return None
        return self._to_float(row.iloc[0]["value"])

    def _lookup_rbtl_by_name(self, key: str, name: str) -> float | None:
        """Fallback: find RBTL row by compound name (case-insensitive partial match)."""
        df = self._rbtl.get(key)
        if df is None:
            return None
        name_col = next((c for c in df.columns
                         if any(k in c.lower() for k in ("name", "compound", "chemical", "chimical"))), None)
        if name_col is None:
            return None
        name_lo = name.lower()
        mask = df[name_col].str.strip().str.lower() == name_lo
        row  = df[mask]
        if row.empty:
            # Try partial / contains match
            mask = df[name_col].str.strip().str.lower().apply(lambda x: name_lo in str(x) if x else False)
            row  = df[mask]
        if row.empty:
            return None
        return self._to_float(row.iloc[0]["value"])

    def _lookup_main_by_name(self, name: str, col_name: str) -> float | None:
        """Fallback: find main-table row by compound name (case-insensitive).
        Partial/contains matching is only applied for names longer than 4 characters
        to avoid false positives for short abbreviations (DRO, ORO, TPH, etc.).
        """
        for df in ([self._vsl_full] if self._vsl_full is not None else []) + [self._main]:
            if df is None or col_name not in df.columns:
                continue
            name_col = next((c for c in df.columns
                             if any(k in c.lower() for k in ("name", "compound", "chemical", "chimical"))), None)
            if name_col is None:
                continue
            name_lo = name.lower()
            mask = df[name_col].str.strip().str.lower() == name_lo
            row  = df[mask]
            if row.empty and len(name_lo) > 4:
                # Only partial-match for names longer than 4 chars to avoid abbreviation collisions
                mask = df[name_col].apply(lambda x: name_lo in str(x).strip().lower() if x is not None else False)
                row  = df[mask]
            if not row.empty:
                val = self._to_float(row.iloc[0][col_name])
                if val is not None:
                    return val
        return None

    def _lookup_main(self, cas: str, col_name: str) -> float | None:
        # For VSL_SOIL ("Soil Direct Contact"): try full V7 file first (has metals/TPH)
        if col_name == "Soil Direct Contact" and self._vsl_full is not None:
            val = self._lookup_df(self._vsl_full, cas, col_name)
            if val is not None:
                return val
        # Fall back to compact main file (also has TIER1 gas and GW columns)
        return self._lookup_df(self._main, cas, col_name)

    @staticmethod
    def _lookup_df(df: pd.DataFrame, cas: str, col_name: str) -> float | None:
        if "CAS No." not in df.columns or col_name not in df.columns:
            return None
        row = df[df["CAS No."] == cas]
        if row.empty:
            return None
        return ThresholdManager._to_float(row.iloc[0][col_name])

    # Maps granular Tier1 key → (parent sheet key, [mg/kg] column index 0/1/2)
    _PFAS_DIRECT_KEY_MAP: dict[str, tuple[str, int]] = {
        "tier1_res_very_high": ("tier1_res", 0),
        "tier1_res_0_6":       ("tier1_res", 0),
        "tier1_res_6plus":     ("tier1_res", 1),
        "tier1_res_no_gw":     ("tier1_res", 2),
        "tier1_ind_very_high": ("tier1_ind", 0),
        "tier1_ind_0_6":       ("tier1_ind", 0),
        "tier1_ind_6plus":     ("tier1_ind", 1),
        "tier1_ind_no_gw":     ("tier1_ind", 2),
    }

    @staticmethod
    def _nth_numeric_after_cas(row, df_columns, cas_col, n: int) -> float | None:
        """Return the n-th (0-based) numeric value in columns that come after cas_col.

        This is the same column-discovery strategy used by _lookup_pfas (VSL):
        scan left-to-right after the CAS column and count numeric hits.
        Avoids any dependency on column names (e.g. '[mg/kg]').
        """
        cas_pos = list(df_columns).index(cas_col)
        count = 0
        for col in df_columns[cas_pos + 1:]:
            val = ThresholdManager._to_float(row[col])
            if val is not None:
                if count == n:
                    return val
                count += 1
        return None

    def _lookup_pfas_direct(self, sheet_key: str, cas: str) -> float | None:
        """CAS lookup on a granular PFAS Tier1 column (col_idx 0/1/2).

        Uses [mg/kg]-prefixed column names (pandas renames duplicates as
        [mg/kg], [mg/kg].1, [mg/kg].2) to avoid counting citation-reference
        integer columns as numeric thresholds.
        """
        parent_key, col_idx = self._PFAS_DIRECT_KEY_MAP.get(sheet_key, (sheet_key, 0))
        df = self._pfas.get(parent_key)
        if df is None:
            return None
        cas_col = next((c for c in df.columns if "cas" in c.lower()), None)
        if cas_col is None:
            return None
        row_df = df[df[cas_col].str.strip() == cas]
        if row_df.empty:
            return None
        mg_cols = [c for c in df.columns if str(c).startswith("[mg/kg]")]
        if col_idx >= len(mg_cols):
            return None
        val = self._to_float(row_df.iloc[0][mg_cols[col_idx]])
        return val * 1000 if val is not None else None

    def _lookup_pfas_direct_by_name(self, sheet_key: str, name: str) -> float | None:
        """Name fallback for a granular PFAS Tier1 column — same strategy as _lookup_pfas_by_name."""
        parent_key, col_idx = self._PFAS_DIRECT_KEY_MAP.get(sheet_key, (sheet_key, 0))
        df = self._pfas.get(parent_key)
        if df is None:
            return None
        name_col = next(
            (c for c in df.columns if any(k in c.lower() for k in ("chemical", "name", "compound"))),
            None,
        )
        if name_col is None:
            return None
        cas_col = next((c for c in df.columns if "cas" in c.lower()), None)
        if cas_col is None:
            return None
        name_lo = name.lower()
        mask = df[name_col].str.strip().str.lower() == name_lo
        row_df = df[mask]
        if row_df.empty:
            mask = df[name_col].str.strip().str.lower().apply(lambda x: name_lo in str(x) if x else False)
            row_df = df[mask]
        if row_df.empty:
            return None
        mg_cols = [c for c in df.columns if str(c).startswith("[mg/kg]")]
        if col_idx >= len(mg_cols):
            return None
        val = self._to_float(row_df.iloc[0][mg_cols[col_idx]])
        return val * 1000 if val is not None else None

    def _lookup_pfas(self, sheet_key: str, cas: str) -> float | None:
        df = self._pfas.get(sheet_key)
        if df is None:
            return None
        # Find CAS column
        cas_col = next((c for c in df.columns if "cas" in c.lower()), None)
        if cas_col is None:
            return None
        row = df[df[cas_col].str.strip() == cas]
        if row.empty:
            return None
        # First numeric-looking column after CAS column
        cas_idx = list(df.columns).index(cas_col)
        for col in df.columns[cas_idx + 1:]:
            val = self._to_float(row.iloc[0][col])
            if val is not None:
                # Threshold file is in mg/kg; lab data is in ng/g → multiply by 1000
                return val * 1000
        return None

    def _lookup_pfas_by_name(self, sheet_key: str, name: str) -> float | None:
        df = self._pfas.get(sheet_key)
        if df is None:
            return None
        name_col = next(
            (c for c in df.columns if any(k in c.lower() for k in ("chemical", "name", "compound"))),
            None,
        )
        if name_col is None:
            return None
        name_lo = name.lower()
        mask = df[name_col].str.strip().str.lower() == name_lo
        row = df[mask]
        if row.empty:
            mask = df[name_col].str.strip().str.lower().apply(lambda x: name_lo in str(x) if x else False)
            row = df[mask]
        if row.empty:
            return None
        cas_col = next((c for c in df.columns if "cas" in c.lower()), None)
        if cas_col is None:
            return None
        cas_idx = list(df.columns).index(cas_col)
        for col in df.columns[cas_idx + 1:]:
            val = self._to_float(row.iloc[0][col])
            if val is not None:
                return val * 1000
        return None

    @staticmethod
    def _name_variants(name: str) -> list[str]:
        """Return name variants to try when matching compound names.

        Always includes the original name. Also strips trailing parenthetical
        notes (e.g. "Benzene (total)" → "Benzene") as a second candidate.
        """
        s = name.strip()
        variants = [s]
        # Strip trailing parenthetical notes
        base = re.sub(r'\s*\([^)]*\)\s*$', '', s).strip()
        if base and base != s:
            variants.append(base)
        return variants

    @staticmethod
    def _to_float(val) -> float | None:
        try:
            f = float(str(val).replace(",", "").strip())
            return f if f == f else None   # NaN check
        except (ValueError, TypeError):
            return None
