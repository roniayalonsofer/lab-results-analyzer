# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the app

```
py -3 -m streamlit run app.py
```

The app runs on port 8501. All lab-parsing logic lives under `soil_lab_tool/` which is added to `sys.path` at startup.

## Architecture

### Data flow

```
uploaded file(s)
    → auto_detect_lab() / auto_detect_category()   [parsers/__init__.py]
    → get_parser(lab, category)                     [parsers/__init__.py registry]
    → parser.parse(BytesIO) → list[dict]            [one record per compound×sample]
    → LabReportExcel / LabReportWord                [core/excel_output.py, core/word_output.py]
```

Every record dict must contain: `compound`, `cas`, `value` (float|None), `flag` (''|'ND'|'<LOQ'|'<LOD'), `unit`, `sample_id`, `analysis_type`, and optionally `lod`, `loq`.

### Parser registry (`soil_lab_tool/parsers/__init__.py`)

Maps `(lab_key, category_key) → ParserClass`. Lab keys are lowercase strings (`"kte"`, `"alchem"`, `"aminolab"`, …). Category keys: `"soil"`, `"groundwater"`, `"soil_gas"`, `"pfas"`, `"pr"`, `"grain_size"`.

`auto_detect_lab()` and `auto_detect_category()` try filename hints first, then content-based inspection (sheet names, PDF text, row peeks). Content checks always run **before** the KTE filename fallback.

### Adding a new parser

1. Create `soil_lab_tool/parsers/<category>/<lab>.py` — subclass `BaseParser`, set `LAB_NAME` and `ANALYSIS_TYPES`, implement `parse(file_obj) → list[dict]`.
2. Import and register it in `parsers/__init__.py` (`_REGISTRY` dict).
3. Add detection logic to `auto_detect_lab()` and `auto_detect_category()` in the same file.
4. If the lab name can appear in filenames, add it to the filename-hint block at the top of `auto_detect_lab()`.

### Analysis types

Defined by convention (string keys): `SOIL_VOC`, `SOIL_SVOC`, `SOIL_TPH`, `SOIL_METALS`, `SOIL_PFAS`, `SOIL_GAS_VOC`, `GW_VOC`, `GW_PFAS`, `LOWFLOW`. The Excel output builder (`core/excel_output.py`) maps each to a sheet name and unit label. Unknown types are silently skipped.

### Threshold manager (`core/threshold_manager.py`)

Loaded once at startup via `@st.cache_resource`. Key method: `get_threshold_with_name(cas, threshold_key, compound_name)` — tries CAS lookup first, falls back to compound name. VSL keys contain `"VSL"`; Tier-1 keys are everything else (`TIER1_*`, `GAS_*`, `GW`, `PFAS_TIER1_*`).

### RTL / Hebrew

PDF text from Hebrew PDFs arrives visually reversed (characters and token order). Use the `_fix_rtl()` helper in `parsers/soil/bactochem.py` as a reference for correct reversal logic.

CAS lookups for Hebrew compound names are in `core/cas_lookup.py`. Add new Hebrew ↔ English ↔ CAS mappings there.

### `LabValueParser` (`core/lab_value_parser.py`)

Parses raw strings like `"<0.5"`, `"N.D."`, `"1.2E-3"` into `(float|None, flag)`. Handles `<`, `>`, ND sentinel strings, and Hebrew non-detect phrases.

## Key files

| File | Purpose |
|---|---|
| `app.py` | Streamlit UI — upload, threshold selection, download, preview |
| `parsers/__init__.py` | Registry + auto-detection |
| `core/threshold_manager.py` | VSL / Tier-1 threshold lookups |
| `core/excel_output.py` | Multi-sheet Excel builder with color coding |
| `core/cas_lookup.py` | Hebrew/English name → CAS mapping |
