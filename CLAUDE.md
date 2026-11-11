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

The analysis page has no automatic option: the user must choose the lab (primary and each secondary file) and the category before anything is parsed. `auto_detect_category()` only picks the PDF layout parser inside the chosen category (`soil` → `soil_pdf`). `auto_detect_lab()` stays for the annual report's info caption and the tests.

`auto_detect_lab()` and `auto_detect_category()` try filename hints first, then content-based inspection (sheet names, PDF text, row peeks). Content checks always run **before** the KTE filename fallback.

### Adding a new parser

1. Create `soil_lab_tool/parsers/<category>/<lab>.py` — subclass `BaseParser`, set `LAB_NAME` and `ANALYSIS_TYPES`, implement `parse(file_obj) → list[dict]`.
2. Import and register it in `parsers/__init__.py` (`_REGISTRY` dict).
3. Add detection logic to `auto_detect_lab()` and `auto_detect_category()` in the same file.
4. If the lab name can appear in filenames, add it to the filename-hint block at the top of `auto_detect_lab()`.

Detection runs on the **data files first** (non-PDF); PDFs uploaded next to an xlsx are companion copies and must not decide the lab/parser (`app.py`). Put the most specific signature (e.g. Machon HaNeft certificate header + EPA codes) *before* sheet-name heuristics (a sheet called `TPH` matches Alchem). Never put `""` in a `startswith` prefix tuple — it matches everything. See `.claude/skills/lab-app-dev/SKILL.md` → "Lessons learned".

### Analysis types

Defined by convention (string keys): `SOIL_VOC`, `SOIL_SVOC`, `SOIL_TPH`, `SOIL_METALS`, `SOIL_PFAS`, `SOIL_GAS_VOC`, `GW_VOC`, `GW_PFAS`, `LOWFLOW`. The Excel output builder (`core/excel_output.py`) maps each to a sheet name and unit label. Unknown types are silently skipped.

### Threshold manager (`core/threshold_manager.py`)

Loaded once at startup via `@st.cache_resource`. Key method: `get_threshold_with_name(cas, threshold_key, compound_name)` — tries CAS lookup first, falls back to compound name. VSL keys contain `"VSL"`; Tier-1 keys are everything else (`TIER1_*`, `GAS_*`, `GW`, `PFAS_TIER1_*`).

### Drinking water standard (`core/drinking_water.py`)

The maximum concentrations of the 2013 drinking water regulations, in mg/L like the groundwater sheets, under the threshold key `DRINKING_WATER`. Names match exactly, then the CAS, so it never asks a compound question. An exceedance is orange in Excel. The older `GW` column stays as an option, off by default.

### Ambiguous compounds (`core/compound_choices.py`)

Thresholds are matched by CAS: a CAS with a row in the threshold table is never questioned, whatever the lab calls the compound, and a CAS the table doesn't have means no threshold. A compound with no CAS is asked about in the app (`_compound_picker`) when the table has similar names. Also asked: a known compound name under a CAS the table lacks. A summed result with a CAS ("m+p-Xylene", 108-38-3) takes that CAS's threshold, by the owner's decision. No report is built until every question is answered. The answers go to Excel, Word and the preview together: `apply_compound_choices` sets the chosen CAS on the records and `ResolvedThresholds` wraps the threshold manager. The lab's compound name stays as printed on the certificate. A clear spelling reported without a CAS (e.g. MTBE) belongs in `CHEMICAL_MAP`, so it is never asked about (`name_matches_cas`).

### Deploying (`code_reload.py`)

A push makes Streamlit Cloud re-run the new `app.py` in the same process, so modules under `soil_lab_tool` would stay at their old version. `drop_stale_modules` forgets them when the code changes, and the threshold manager's cache is keyed on the same version. After a push, check the live site by running a full report, not only by loading the page.

### LOD / LOQ fields (`core/file_limits.py`)

The limits section is drawn after the file is read: each field is filled when all its records share one value, and a table lists the LOD and LOQ of every compound. A field still showing the file's value is not a manual entry; only a value the user changed fills what the file lacks.

### Field forms (`core/field_forms.py`)

The PID uploader takes several forms, Excel or PDF, and merges them (a later form replaces the reading at the same depth). Columns are found by their header ("קידוח", "עד עומק", "PID"), with the known positions as a fallback for an Excel file with no header. PDF reading needs a ruled table with text: a scanned form is reported, not read. A `pid_reading` on a record (from a soil gas file) is preferred over the field forms.

### Scanned field forms (`core/form_fields.py`)

Field forms are handwritten scans the app can't read. The owner decided against a paid AI reading service, so nothing leaves the app: for a PDF form with no readable table, `_form_entry` shows the scan (the custody page first), the user chooses fields from a list and types the values per sample. Chosen values reach the soil gas sheet as extra header rows (`form_values`). Other sample types have no place in the report yet, and need a real form of each type first.

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

## Groundwater periodic report updater (`gw_report_updater.py`)

### What gets updated, and from what

Every update is driven by the **most recent sampling round's results only**
(the latest date parsed from the lab PDF(s)). There is no "merge with
history" logic anywhere except the explicitly historical tables (2 and 3),
which each get exactly **one new row appended** per well per round —
everything else in those tables (older rows) is left untouched.

- **Table 1 — current-round field snapshot** (pH/EC/temperature/etc. per
  well): values are **overwritten in place**, always reflecting only the
  latest round. Row lookup is by the row's own label text (column 0), never
  a hardcoded row index — different documents order these rows differently.
- **Table 2 — historical field-findings table** (if present): **one new
  row appended per well** with the latest round's field data. Not every
  report has this table (some sites only have table 1); the updater
  detects it by content signature (a "קידוח"/"תאריך" header row followed by
  recognized field-parameter names) and is a silent no-op if it's absent —
  never treat its absence as an error.
- **Table 3/4 — chemistry (BTEX/MTBE) history**: **one new row appended per
  well**. Exceedance bold/red marking is decided by **that table's own
  "ריכוזי יעד לסיום ניטור" row** (site-specific targets), never a
  hardcoded threshold — different sites have different approved targets.
- **Narrative text** (title month/year, sampling-date paragraphs, the
  per-well and combined BTEX summary sentences): rebuilt from the latest
  round's results/date every time, via `_replace_value_after_anchor()`
  rather than naive adjacent-run text matching — Word frequently splits
  anchor text and values across many runs (track-changes/spellcheck
  artifacts), which breaks anything that assumes a clean 1:1 run mapping.
- **Cover-page metadata table** (author name, submission date): only
  filled in if the caller passes `author_name`/`submission_date` to
  `run_update_bytes()` — both optional, both no-ops if omitted.

### Inputs: labs, field form, Mann-Kendall

- **Labs:** Bactochem PDF or its CSV (`parse_bactochem_csv`, same shape;
  the PDF wins when both come, the CSV has no sampler's name), Aminolab PDFs, or ALS (PDF or the "נספח" XLSX,
  read by `core/gw_round_inputs.parse_als_round`, results in mg/L; a
  non-detect is `None` with its limit in `limits`, printed as `<0.0002`).
- **Field form:** Yeda Water's "טופס ביצוע עבודה" .xlsx
  (`parse_yeda_field_form`). Columns by header; each well takes its LAST
  low-flow reading; "עומק דגימה מפני המים" is extraction depth minus water
  level (1.0), by the owner's choice. Its values win over the lab's.
- **Mann-Kendall** (`core/gw_mann_kendall.py`): the GSI toolkit .xls holds
  thousands of formulas and charts, so it is **only read, never written**
  (owner's decision). The round is added in memory, S/COV/CF/trend are
  computed (they match Excel's cached values; CF comes from the workbook's
  own Kendall_Dist sheet), the toolkit page is drawn as a picture and put in
  place of the picture under the caption "גרף מספר N: ... <compound>", and
  the app hands back the rows to paste into the .xls. Values below 0.01 or
  non-detects are entered as 0.01 (the site's convention). The sheet for a
  compound is the one whose events reach the latest date. Hebrew in the
  pictures: whether matplotlib orders it right to left is tested by drawing
  (`_mpl_does_bidi`), never by version — the server's Linux build of the
  same version drew it backwards.
- **No lab choice in the app:** the periodic tab takes all of a round's
  files at once; `core/gw_round_files.sort_round_files` identifies each file
  by READING it (lab parser must return samples — `auto_detect_lab` misses
  some ALS appendices), groups them into rounds by date, prefers an ALS
  xlsx over its PDF copy, attaches the field form by date, and picks up a
  MK workbook. The app then runs `run_update_bytes(..., parsed_round=...)`
  per round, oldest first, and `apply_mann_kendall` once with all rounds.
- ALS without a "Sum of xylenes" row: xylene = m+p + o.
- The annual report engine (`core/annual_gw_report.py`) only knows "מת-"
  wells and the last round per well; it takes a MK file through
  `apply_mann_kendall`, but needs a real annual sample before more work.
- Per-well result tables on the site map ("חק-4 / date / תרכובת | ריכוז יעד |
  ריכוז") are updated by `_update_well_cards`, bold over the card's own target.
- Map labels in text boxes: the newest date on the map becomes the round's
  date, and "MTBE/בנזן" pairs ("0.0117/0.00057") are matched to their well by
  the previous round's history values, then tagged in the drawing's hidden
  title (`gw-label:<well>|<p1>|<p2>`) so later rounds find them even when two
  wells show the same values (`_update_map_labels`).
- MK pages without a "גרף מספר" caption: the pictures right after "בגרפים
  הבאים מוצגת מגמת ..." are taken in the MK file's sheet order.

### Tables: grid columns, nested tables, one-off tables

- Read chemistry tables by **grid column**, never by `row.cells` index:
  headers and rows span different columns (MTBE header over two columns, a
  row's xylene over two, `w:gridBefore`). `_history_header`,
  `_row_param_cells` and `_history_rows` do this; a table may hold several
  parts, each row uses the header above it.
- Search tables with `_all_tables(doc)`: some field tables sit inside a
  one-cell frame table, which `doc.tables` does not list.
- A one-off table (e.g. before/after an injection) shows a well twice; the
  current field table is the one with one column per well, so the one-off
  is never touched.
- Keep the lab's printed digits: Bactochem gives `result_text` / `field_text`
  ("0.370", "11.80"), used instead of reformatting the float.

### Idempotency / re-running safely

`run_update_bytes()` can be called multiple times in a chain (the
multi-round upload feature in `app.py` does exactly this — one call per
uploaded PDF, sorted chronologically, each call's output feeding the next
call's input). Because of this, every row-insertion path checks whether a
row for that exact well+date already exists before inserting, and reuses
it instead of adding a duplicate. If you add a new "append a row" style
update, follow the same pattern (see `_update_historical_tables`).

### Formatting gotcha

New table rows are created by deep-copying the previous row for that well
(`_insert_row_after`), which means they **inherit that row's formatting**
(bold, color) as a starting point. `_set_cell()` explicitly resets
bold/color on every call rather than only turning them on — if you write
new cell-formatting logic, always set the "off" state explicitly too, or
formatting from an old exceedance will silently persist forever on every
later row.

### Standing rule: never assume a fixed row/column order

This has come up twice already (table 1's field-parameter rows, and the
BTEX/MTBE chemistry table's columns) — different site templates order
these differently, and some even spell a parameter two different ways
(e.g. Xylene as "קסילן" vs "כסילן"). **Always read the table's own
header/label row to build a name→position map at runtime** (see
`name_to_row` in `_update_field_table` and `_detect_chem_table_columns`)
rather than hardcoding an index or column order. The same applies to
threshold rows: read the table's own threshold row (whatever it's
labeled — "ריכוזי יעד לסיום ניטור", "תקן מי שתייה", etc.) rather than a
single hardcoded number, since different sites use different standards.

