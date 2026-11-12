---
name: lab-app-dev
description: Work with Roni on the lab-results-analyzer Streamlit app: add or fix a lab-report parser, adapt to a new lab/report format, redesign the UI professionally, or QA the Excel/Word output.
---

# lab-results-analyzer — development partner

App: https://lab-results-analyzer-alsdxcagccll2wx3yo76e9.streamlit.app/
Repo: https://github.com/roniayalonsofer/lab-results-analyzer (Streamlit Cloud redeploys automatically on every push to the `master` branch).
Owner: Adama (אדמה) — environmental consulting. Users are Hebrew-speaking hydrogeologists / environmental engineers. Always reply to Roni in Hebrew; code, identifiers and commit messages may be English (Hebrew commit messages are also fine — match the repo history).

The repo's own `CLAUDE.md` is the source of truth for architecture. Read it at the start of every session; this skill adds the working method on top of it.

## 0. Session start (every time)

1. Read `CLAUDE.md`. Make sure dependencies are installed (`py -3 -m pip install -r requirements.txt` on Windows).
2. Ask Roni (one short question, only if not already clear) which mode this session is: **A. new/broken lab report**, **B. design**, **C. output QA**, **D. cleanup/refactor**. Then follow that section.
3. Run the app locally (`py -3 -m streamlit run app.py`, port 8501) so Roni can check every change in her browser before it goes live.
4. Pushes go live (Streamlit Cloud redeploys from `master`). Roni gave a standing OK to push after each verified change (see "Working style → Push policy"); outside that, never push without her explicit OK.

## Project map (quick reference)

| Area | Path |
|---|---|
| UI (single 2,000+ line file, CSS inline ~lines 120–530) | `app.py` |
| Parser registry + auto-detection | `soil_lab_tool/parsers/__init__.py` (`_REGISTRY`, `auto_detect_lab`, `auto_detect_category`) |
| Parsers by category | `soil_lab_tool/parsers/{soil,groundwater,soil_gas,pfas}/<lab>.py` |
| Raw value → (float, flag) | `soil_lab_tool/core/lab_value_parser.py` |
| Hebrew/English name → CAS | `soil_lab_tool/core/cas_lookup.py` |
| Thresholds (VSL, Tier-1, GW, PFAS) | `soil_lab_tool/core/threshold_manager.py` + xlsx in `soil_lab_tool/thresholds/` |
| Excel output (sheet per analysis type, `SHEET_CONFIG`) | `soil_lab_tool/core/excel_output.py` |
| Word output | `soil_lab_tool/core/word_output.py` |
| GW periodic report updater (.docx in place) | `soil_lab_tool/gw_report_updater.py` |
| Annual GW report | `soil_lab_tool/core/annual_gw_report.py` |

Labs currently supported: KTE, Alchem/אלכם (incl. XRF, soil-gas PDF, TPH PDF), ALS (soil, soil PDF, GW, grain size), Aminolab/אמינולאב, Bactochem/בקטוכם, מכון הנפט, מכון האנרגיה, RJ Lee (PFAS).

Record contract (every parser output row): `compound`, `cas`, `value` (float|None), `flag` (`''|'ND'|'<LOQ'|'<LOD'`, also `'<'`/`'>'`), `unit`, `sample_id`, `analysis_type` (`SOIL_VOC`, `SOIL_SVOC`, `SOIL_TPH`, `SOIL_METALS`, `SOIL_PFAS`, `SOIL_GAS_VOC`, `GW_VOC`, `GW_PFAS`, `LOWFLOW`), optional `lod`, `loq`. An `analysis_type` missing from `SHEET_CONFIG` is **silently dropped** from the Excel — always check this when adding a type.

## A. New lab / new report format / parser bug

Goal: every lab certificate Roni receives produces a correct table with no manual fixing.

1. **Get a real sample.** Ask Roni for the actual file (PDF/XLSX/XLS) — never write a parser from a description. If several layouts exist from the same lab (different years, different analysis packages), ask for one of each.
2. **Inspect before coding.** For PDFs: `pdfplumber` text + `extract_tables()` on every page, and render page 1 to PNG and look at it. Hebrew PDF text arrives visually reversed — reuse `_fix_rtl()` from `parsers/soil/bactochem.py`; do not write a new reversal. For Excel: list sheet names, dump the first ~40 rows of each sheet with `pandas.read_excel(header=None)`.
3. **Write down the layout** in a short comment block at the top of the parser: where sample IDs are, where compound names are, units row, LOQ/LOD columns, how non-detects are written, multi-page continuation.
4. **Decide: extend or new.** Same lab with a new layout → extend the existing parser with a layout-detection branch; do not fork a near-copy file. New lab → new file `parsers/<category>/<lab>.py` subclassing `BaseParser`.
5. **Standing rules (see CLAUDE.md):**
   - Never assume fixed row/column order. Build a name→index map from the header row at runtime.
   - Accept spelling variants (e.g. קסילן/כסילן, Xylenes/Xylene (total), m+p-Xylene). Add new Hebrew↔English↔CAS mappings to `cas_lookup.py`, not inside the parser.
   - All raw value strings go through `LabValueParser` — `<0.5`, `N.D.`, `לא נמצא`, `1.2E-3`, comma decimals.
   - Keep the reported unit; convert only if the threshold table uses a different unit, and say so in a comment.
6. **Register** in `_REGISTRY` and add detection in `auto_detect_lab()` / `auto_detect_category()`: content checks first (logo text, lab name, sheet names), filename hints last. Confirm the new detection does not steal files from other labs — run detection over every sample in `tests/samples/`.
7. **Regression test (build once, then keep it):** if `tests/` does not exist, create `tests/samples/<lab>/<category>/` and `tests/test_parsers.py` that, for each sample, asserts: detected lab/category, number of samples, number of records, and 3–5 spot values checked against the certificate. Client data: ask Roni before committing samples; if confidential, keep them in `.gitignore`. Run the whole suite after every parser change — fixing one lab must not break another.
8. **Check the output end to end:** build the Excel via `LabReportExcel`, open it, confirm every analysis type has a sheet, exceedances are colored against the selected thresholds, and no compound is missing a CAS/threshold where one exists. Report to Roni in a short Hebrew table: compounds parsed / matched to CAS / matched to threshold, and list the unmatched ones.

## B. Design — make the app look professional

Brand base already in `app.py`: navy `#1e3a4f` (primary), teal `#4a7a8a` / `#2d5a6e` (secondary), slate greys `#64748b #94a3b8 #e2e8f0`, background `#f0f4f8`, font Heebo, RTL, logo `assets/adama_logo.png`. Keep this identity; refine, don't replace.

Before changing anything, load the `frontend-design` skill if available, look at the running app (screenshots of each page: soil, groundwater, soil gas, PFAS, preview dialog) and give Roni a short critique — top 5 issues by impact. Agree on scope before a big redesign.

Principles:
- **Workflow first.** upload → detect lab & category → pick thresholds (land use, sensitivity, depth) → preview → download Excel/Word. Make it a clear numbered flow (`_steps()` exists — use it consistently), show detection plainly ("זוהה: KTE · קרקע · 12 דגימות") with an easy override.
- **Tokens, not scattered hex.** Colors, radii, spacing, font sizes as CSS variables in one `:root` block; replace hardcoded hex with `var(--…)`. Same palette in Excel/Word headers so outputs match the app.
- **Move CSS out of `app.py`** into `assets/style.css`; consider splitting `app.py` into `ui/` modules per page — only with Roni's agreement, as a separate step from visual changes.
- **RTL correctness:** numbers, CAS numbers, units and English compound names inside RTL text need `dir="ltr"` / `unicode-bidi: isolate`.
- **Tables:** exceedance coloring explained by a visible legend (VSL vs Tier-1), ND/<LOQ shown consistently.
- **Streamlit limits:** global selectors like `* { font-family }` break icon fonts (already patched once); `data-testid` selectors change between Streamlit versions, so pin the Streamlit version in `requirements.txt` when relying on them.
- Verify every design change in the running app before showing Roni.

## C. Output QA (Excel / Word)

When Roni reports a wrong output: get the input file and the bad output, reproduce headlessly, find the root cause (parser vs CAS match vs threshold vs formatting), fix it, add the case to `tests/`. For Word outputs follow the house standard (David 13 Hebrew, TNR 11 English, RTL, A4 — see the `hebrew-doc-format-v2` skill if available). For the GW periodic updater obey CLAUDE.md's rules (latest round only, idempotent row append, explicit bold/color reset, thresholds from the table's own target row).

## D. Repo hygiene (offer once, do with approval)

Known clutter: ~80 editor temp files `core/*.py.tmp.NNNN` (the `.gitignore` pattern `*.tmp` doesn't match them — add `*.tmp.*`), `debug_*.py`, `check_thresh*.py`, `tmp_*.py`, `tmp_page1.png`, `from raw to excel (36).py`, `soil_lab_tool_share.zip`, `test_output.xlsx`. Propose moving debug scripts to `scripts/`, deleting temp/zip files, pinning versions in `requirements.txt`. Never delete without listing the files to Roni first.

## Working style with Roni

- Small, reviewable changes; one topic per commit; explain each change in 1–2 Hebrew sentences.
- Always show proof: test output, a screenshot, or the specific cell values that now match the certificate.
- When something in a lab file is ambiguous (which column is LOQ, whether a value is dry-weight), ask Roni — she knows the labs; don't guess.
- Push policy (Roni's standing OK, given 2026-10-28): after each finished, verified change, commit (only the files of that change) and push to `master`, then tell her. This overrides the "ask before push" step for this session type; never push anything unverified, never push `.claude/settings.local.json`, and stop pushing if she says so.
- Debug headlessly first: reproduce with a short Python script that emulates the `app.py` flow (detect lab → detect category → `get_parser` → `parse`), write its output to a UTF-8 file and read that (Hebrew garbles in the PowerShell console). Use the browser only for visual checks.

## Self-improvement loop (do this, don't just read it)

This skill must get better every session. Triggers and actions:

1. **Whenever a bug is fixed or a new lab/layout is handled**, append an entry to "Lessons learned" below *in the same commit* as the fix: symptom → root cause → rule to prevent it. One to three lines, no client data.
2. **Whenever Roni corrects my approach or states a preference**, add it to "Working style" or "Lessons learned" immediately, not at the end.
3. **At the end of each session** (or when Roni says we're done): re-read this file, merge duplicates, delete entries that are no longer true, move general rules into the relevant section (A/B/C/D), and mirror any architecture-level rule into `CLAUDE.md`.
4. **Every fix gets a regression test** in `tests/` using a small synthetic workbook/PDF (no client data in git). Real client samples live only on the X: drive; their absolute folder is recorded below under "Sample locations".

## Lessons learned (newest first)

- **2026-11-12 — Machon HaNeft files from 2020-2023 gave 0 records** (metals, VOC, TPH laid out as the printed certificate). `parsers/soil/machon_haneft_cert.py` reads them when the main reader finds nothing: header block from the first "גבול" row, sample columns between the label columns and the limit columns, the sample name joined from its header rows ("EL-2" + "-3"), "50>" is "<50". Rule: name the analysis type by the sheet's content, not the file name ("2384 SVOC.xlsx" holds VOC).

- **2026-11-12 — Aminolab soil PDFs came out as one sample "Unknown" with junk names; a KTE EDD gave 0 records.** The Aminolab reader was written for groundwater; soil certificates (Hebrew TPH/ICP table, then English VOC ppb / SVOC ppm tables with CAS, several samples per PDF) now have `parsers/soil/aminolab.py` (PyMuPDF words, each value to the nearest name line). An EDD sheet (ClientSampID/Analyte/Result/Rlimit/CAS) is read by `parsers/groundwater/edd.py`. Rules: a file called "Aminolab(…)" may be the custody form, not the certificate, so look at it before calling it a gap; check a printed CAS with the check digit (Aminolab prints 95-18-4 and 100-01-8); verify a PDF reader by counting the value words in the result column of every page.

- **2026-11-12 — Bactochem 2026 PDF layout ("-HE-R").** Sample line "16/09/2026 09:00 P-1 1 הדוקנ- ..." (no "מספר הדוגמה"), result "① 0.001 mg/L 0.002 MTBE" with the CAS on the next line, field-line parentheses flipped, pH written "7.37 )…( pH", and a sample that runs over a page repeats its line (the old reader skipped the rest). Rule: compare a new PDF layout against the CSV of the same certificate, field by field; merge a repeated sample, never skip it.
- **2026-11-11 — Egged Azor (Yeda Water) rounds not read at all.** Latin well names (forms, certificates), wells matched by number only (EAZ-1 = נ-1), dates "*02/01/17", limits in "ערך סף", gray exceedances, split field table, and a field table whose 2nd row starts "pH" taken for layout B. Rule: run a new site's real previous report through every round in its lab folder and diff the text; never decide a layout from one cell.
- **2026-11-11 — MK pictures: Hebrew backwards on the live site only; Bactochem CSV refused.** Both machines run matplotlib with libraqm, but the server lays every line out left to right, so 'פפ-1' came out as 'פפ-1' (plain Hebrew words looked fine). Fix: Hebrew labels go in inside RLE…PDF, and `_bidi_mode` picks that or manual reversal by drawing test words. Rule: test platform behaviour by doing it, test a mixed label (letters + number), and check the live site's output — paste small synthetic files into the live uploader with JS (base64 → File → DataTransfer) when localhost is blocked. The periodic uploader had no `csv` type, so Bactochem's CSV (opens in Excel) was refused; it is now read like the PDF.
- **2026-11-10 — Parbar periodic report (Bactochem): wrong values, a missing well, a one-off table overwritten, map/MK/summary not updated.** Causes: tables read by `row.cells` index while header and rows span different grid columns (MTBE written over xylene, a well name over two columns hid the dates); rows of the second well in a shared table read before the first well's insert; the real field table nested in a one-cell frame, so the before/after-injection table was taken instead; MK pages had no "גרף מספר" caption; "בתאריכים" starts with "בתאריכ", not "בתאריך"; matplotlib 3.11 orders Hebrew itself (libraqm), so pre-reversing broke labels. Rule: read Word tables by grid column and re-read after every insert; search nested tables; reproduce the user's output from the real inputs and diff it against the hand-made report of the same round before changing code.

- **2026-11-10 — Periodic GW report: the Hof site could not be updated at all, and the MK code would have wiped the workbook.** The site is sampled by Yeda Water with ALS results (only Bactochem/Aminolab were read), the field-form uploader was ignored, and the old MK writer rebuilt the .xls with xlwt, dropping 2,550 formulas and 7 charts. Rule: never write an owner's .xls/.xlsm with formulas or charts; read it, compute, and hand back what to paste. Test a report updater by applying round N+1 to the round-N report and diffing against the round-N+1 report made by hand (found 4 hand errors: a history row overwritten, swapped targets, o-xylene as total xylenes, xylene typed into the MK toluene sheet).

- **2026-11-04 — A report with 297 of 515 CAS numbers empty, some wrong.** ALS prints no CAS, and the report builder's name matching filled only what the threshold tables know, sometimes with a similar compound's number (4:2 FTS got 6:2 FTS's). `parsers/groundwater/als_cas.py` holds reviewed numbers. Rule: a name the lab prints gets a checked number or none; audit a new lab's CAS column against PubChem by name.
- **2026-11-04 — ALS water certificate: real .xlsx crashed, PDF gave 0 records.** The parser read only SpreadsheetML, and split sample names on spaces ("Site Name GW-1"). Rule: the Laboratory sample ID row gives the column count. Also: "Nonmetallic" contains "metal" (check inorganics first); pdfplumber drops the µ, leaving "g/L". Compare an xlsx and PDF of one job row by row.
- **2026-11-04 — Drinking water regulations 2013 are a threshold column (`core/drinking_water.py`, mg/L, key `DRINKING_WATER`).** The older `GW` column has values of unclear unit and never coloured, so it is off by default. Never fuzzy-match to this list: "Atrazine-desethyl" isn't atrazine.
- **2026-10-28 — Machon HaNeft xlsx → "no records".** Three independent causes, all fixed:
  (a) `auto_detect_lab` matched the sheet name `TPH` to Alchem before checking the Machon HaNeft certificate → Alchem parser returned 0. Rule: when adding a detector, put the *most specific* signature (certificate header + method codes) before sheet-name heuristics, and run detection over every available sample.
  (b) `app.py` decided lab/category from the *first* uploaded file; a companion PDF (detected as מכון האנרגיה) overrode the xlsx. Rule: PDFs are companions of data files, so detection uses non-PDF files first.
  (c) `_SKIP_PREFIXES` contained `""`; `str.startswith("")` is always True, so every row looked like metadata and the first samples of the BTEX table were silently dropped (4 of 21 samples). Rule: never put an empty string in a prefix tuple; after any parser change check *samples × analysis types* counts, not just total records.
- **Silent data loss is the main risk.** A parser can return a plausible-looking record count while missing samples. Always compare the sample list from the certificate against `{r["sample_id"]}` and check each sample has every expected analysis type.
- **PDF + xlsx pair from the same lab** (e.g. Machon HaNeft 1018 / 6408): the xlsx is the data source, the PDF is a visual copy. The app parses only the non-PDF files when both are uploaded.
- **2026-10-28 — Machon HaNeft TPH limit is LOQ (Roni confirmed).** The TPH limit row (50) is a LOQ; `<50` must be `flag="<LOQ"`, `value=loq`, `loq=50`, `lod=None`. Rule: keep LOD and LOQ rows separate in parsers; use the label printed in the certificate, never merge them.
- **2026-10-28 — Excel showed `<0.015` for a certificate value `<0.03`.** `LabValueParser.parse("<0.03")` returns `(0.015, '<')` (half the limit). Rule: when a certificate prints `<X`, keep X (`value=X, flag="<LOQ", loq=X`); never let the half-limit convention leak into the report. Always check the *generated Excel cells* against the certificate, not only the parsed records (also check `number_format`: `<LOQ` numbers display as `<X` through the format `"<"0.0##`). Other parsers may still use `LabValueParser` for `<X` — check them when they are touched.
- **2026-10-28 — UI regression from a broad CSS selector.** Styling `[data-testid="stFileUploaderDropzone"] button` (with `::after` text) also hit the per-file remove/"add files" buttons that appear *after* upload, so they all read "בחירת קבצים". Rule: target Streamlit widgets by their specific `data-testid` (e.g. `button[data-testid="stBaseButton-secondary"]`), and verify every UI change in *both* states (empty and after uploading files, dialog open). To test uploads in the built-in browser (no upload tool): serve the sample over a tiny CORS http server on 127.0.0.1 and inject it with `fetch → File → DataTransfer → input.files → dispatch change`; resize the viewport to ≥1440 first (the pane is narrow).
- **2026-10-28 — Mobile layout.** The app hides Streamlit's header, so the native sidebar toggle is unreachable on phones; the forced-open sidebar then took ~80% of a 375px screen. Fix (in the `max-width:760px` media queries): stack the sidebar *below* the main content (`order`), make `stApp/stAppViewContainer/stMain` `position:relative` (Streamlit makes them `absolute`), hide `stSidebarResizeHandle` (7px overflow), compact nav. Rule: test every page at 375px and check `document.documentElement.scrollWidth == innerWidth`; list overflowing elements with `getBoundingClientRect().right > clientWidth`.
- **Preview dialog (`_sheet_to_html`)** renders the generated Excel sheet as styled HTML (sticky header/first column, fills from the workbook). Its legend must match `excel_output.py` colors (YELLOW=VSL, L_BLUE=Tier1 residential, PINK=Tier1 industrial, GRAY=uncertain); update both together.
- **Open questions for Roni (verify before changing):** (3) ALS PDF parser (`ALSSoilPDFParser`) returns 0 records on the PR2601739 TPH layout (the XLS version parses fine).

## Sample locations (real client files, never committed)

- Sonol Afridar soil survey 2025: `X:\Shared\Adama\פרויקטים\חברות דלק\סונול\אפרידר\11. סקר קרקע 2025\4. תוצאות מעבדה` — Machon HaNeft 1018 + 6408 (pdf+xlsx), ALS PR2601739 (pdf+xls).
