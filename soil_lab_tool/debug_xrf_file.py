"""
debug_xrf_file.py — diagnostic for XRF file parsing issues.

Usage (from soil_lab_tool/ directory):
    py -3 debug_xrf_file.py path/to/your_xrf_file.xlsx
or
    py -3 debug_xrf_file.py path/to/your_xrf_file.csv

Prints:
  - Raw first 12 rows as pandas sees them
  - Detected header row index and column assignments
  - First 5 parsed records
  - Any suspicious values (time-like floats, mismatched column counts)
"""
import sys, io, re
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, ".")

import pandas as pd

FILE = sys.argv[1] if len(sys.argv) > 1 else None
if not FILE:
    print("Usage: py -3 debug_xrf_file.py <path_to_xrf_file>")
    sys.exit(1)

print(f"\n{'='*60}")
print(f"Diagnosing: {FILE}")
print('='*60)

# ── 1. Read raw ───────────────────────────────────────────────────
with open(FILE, "rb") as f:
    raw = f.read()

magic4 = raw[:4]
is_xlsx = magic4[:2] in (b"PK", b"\xd0\xcf")

print(f"\nFormat: {'XLSX/XLS' if is_xlsx else 'CSV'}")

if is_xlsx:
    try:
        xl = pd.ExcelFile(io.BytesIO(raw))
        print(f"Sheets: {xl.sheet_names}")
        df_raw = xl.parse(xl.sheet_names[0], header=None, dtype=str).fillna("")
    except Exception as e:
        print(f"XLSX read failed: {e}")
        sys.exit(1)
else:
    df_raw = None
    for sep in (",", ";", "\t"):
        for enc in ("utf-8-sig", "utf-8", "cp1255", "latin-1"):
            try:
                df_raw = pd.read_csv(io.BytesIO(raw), header=None, dtype=str,
                                     sep=sep, encoding=enc).fillna("")
                print(f"CSV read OK  sep={sep!r}  enc={enc}  shape={df_raw.shape}")
                break
            except Exception:
                continue
        if df_raw is not None:
            break
    if df_raw is None:
        print("CSV read failed with all separators/encodings")
        sys.exit(1)

print(f"\nDataFrame shape: {df_raw.shape}  ({df_raw.shape[0]} rows × {df_raw.shape[1]} cols)")

# ── 2. Show first 12 rows raw ─────────────────────────────────────
print(f"\n--- First 12 raw rows (truncated to 80 chars per cell) ---")
for ri in range(min(12, len(df_raw))):
    row_vals = [str(v)[:20] for v in df_raw.iloc[ri]]
    print(f"  Row {ri:2d}: {' | '.join(row_vals[:10])}{'...' if len(row_vals) > 10 else ''}")

# ── 3. Detect header row (same logic as parser) ───────────────────
from parsers.soil.xrf import _ELEMENT_CAS, _SKIP_COLS, _LOCATION_HINTS

def _score_row(df, ri):
    row_vals = [str(v).strip().upper() for v in df.iloc[ri]]
    return sum(1 for v in row_vals if v in _ELEMENT_CAS or v.lower() in _SKIP_COLS)

scores = [(ri, _score_row(df_raw, ri)) for ri in range(min(10, len(df_raw)))]
print(f"\n--- Header row detection scores ---")
for ri, sc in scores:
    marker = " ← BEST" if sc == max(s for _, s in scores) else ""
    print(f"  Row {ri}: score={sc}{marker}")

best_row = max(scores, key=lambda x: x[1])[0]
print(f"\nDetected header row: {best_row}")

# ── 4. Show column assignments ────────────────────────────────────
header_vals = [str(v).strip() for v in df_raw.iloc[best_row]]
print(f"\nHeader columns ({len(header_vals)} total):")

from parsers.soil.xrf import _parse_header_col, _ELEMENT_NAME
id_col = None
loc_col = None
element_cols = []
skipped = []
unrecognized = []

for ci, raw_hdr in enumerate(header_vals):
    if not raw_hdr or raw_hdr.lower() in ("nan", ""):
        continue
    low = raw_hdr.strip().lower()
    sym, unit = _parse_header_col(raw_hdr)

    if low in ("sample", "sample id", "sample_id", "sampleid", "id", "מזהה", "מספר",
               "#", "no", "no.", "test #", "reading #", "reading", "point id"):
        role = f"SAMPLE_ID (col {ci})"
        if id_col is None:
            id_col = ci
    elif low in _LOCATION_HINTS:
        role = f"LOCATION (col {ci})"
        loc_col = ci
    elif low in _SKIP_COLS:
        role = f"SKIP"
        skipped.append(f"  col {ci}: {raw_hdr!r}")
    elif sym in _ELEMENT_CAS:
        role = f"ELEMENT {sym} → {_ELEMENT_NAME.get(sym, sym)} (col {ci}, unit={unit})"
        element_cols.append((ci, sym, _ELEMENT_NAME.get(sym, sym), unit))
    else:
        role = f"UNRECOGNIZED"
        unrecognized.append(f"  col {ci}: {raw_hdr!r}  (sym={sym!r})")

    if "ELEMENT" in role or "SAMPLE" in role or "LOCATION" in role:
        print(f"  [{ci:3d}] {raw_hdr!r:<25} → {role}")

if id_col is None:
    print(f"\n  !! No sample ID column found — falling back to col 0")
    id_col = 0

print(f"\nSkipped (metadata) columns ({len(skipped)}):")
for s in skipped[:10]:
    print(s)
print(f"\nUnrecognized columns (not element, not skip): {len(unrecognized)}")
for u in unrecognized[:10]:
    print(u)

print(f"\nElement columns detected: {len(element_cols)}")
print(f"Sample ID column: {id_col}  (header: {header_vals[id_col]!r})")
print(f"Location column:  {loc_col}  (header: {header_vals[loc_col]!r if loc_col is not None else 'n/a'})")

# ── 5. Show first 5 data rows as parser sees them ─────────────────
print(f"\n--- First 5 data rows as parser sees them ---")
data_start = best_row + 1
shown = 0
for ri in range(data_start, min(data_start + 20, len(df_raw))):
    row = df_raw.iloc[ri]
    sid_raw = str(row.iloc[id_col]).strip()
    if not sid_raw or sid_raw.lower() in ("nan", "", "sample", "id"):
        print(f"  Row {ri}: SKIPPED (sid_raw={sid_raw!r})")
        continue
    loc_val = str(row.iloc[loc_col]).strip() if loc_col is not None else ""
    sample_id = f"{sid_raw} – {loc_val}" if loc_val and loc_val.lower() not in ("nan", "") else sid_raw

    print(f"\n  Row {ri}: sample_id={sample_id!r}")
    for ci, sym, compound, unit in element_cols[:8]:
        raw_val = str(row.iloc[ci]).strip()
        print(f"    {sym:4s} (col {ci:3d}): raw={raw_val!r}")

    shown += 1
    if shown >= 5:
        break

# ── 6. Spot-check row 5 (1-indexed data row) ─────────────────────
print(f"\n--- Data row 5 check (original file row {data_start + 4}) ---")
ri5 = data_start + 4
if ri5 < len(df_raw):
    row5 = df_raw.iloc[ri5]
    sid5 = str(row5.iloc[id_col]).strip()
    loc5 = str(row5.iloc[loc_col]).strip() if loc_col is not None else ""
    print(f"  sample_id: {sid5!r}   location: {loc5!r}")
    pb_cols = [(ci, sym) for ci, sym, _, _ in element_cols if sym == "PB"]
    as_cols = [(ci, sym) for ci, sym, _, _ in element_cols if sym == "AS"]
    for ci, sym in pb_cols:
        print(f"  Pb (col {ci}): {str(row5.iloc[ci])!r}")
    for ci, sym in as_cols:
        print(f"  As (col {ci}): {str(row5.iloc[ci])!r}")
else:
    print("  (file has fewer than 5 data rows)")

# ── 7. Count non-detect patterns in Pb column ────────────────────
print(f"\n--- Pb column value distribution ---")
pb_col_idx = next((ci for ci, sym, _, _ in element_cols if sym == "PB"), None)
if pb_col_idx is not None:
    pb_vals = [str(df_raw.iloc[ri].iloc[pb_col_idx]).strip()
               for ri in range(data_start, len(df_raw))
               if str(df_raw.iloc[ri].iloc[id_col]).strip() not in ("", "nan", "sample", "id")]
    nd_patterns = {}
    numeric_count = 0
    for v in pb_vals:
        low = v.lower()
        if any(x in low for x in ("lod", "loq", "dl", "mdl", "nd", "bdl", "n.d")):
            nd_patterns[v] = nd_patterns.get(v, 0) + 1
        elif v.startswith("<"):
            nd_patterns[v] = nd_patterns.get(v, 0) + 1
        elif v == "" or v == "nan":
            nd_patterns["(blank)"] = nd_patterns.get("(blank)", 0) + 1
        else:
            numeric_count += 1
    print(f"  Total Pb values: {len(pb_vals)}")
    print(f"  Numeric (detected): {numeric_count}")
    print(f"  Non-detect patterns:")
    for pat, cnt in sorted(nd_patterns.items(), key=lambda x: -x[1])[:15]:
        print(f"    {cnt:5d}×  {pat!r}")
else:
    print("  No Pb column found in element_cols")

print(f"\n{'='*60}\nDone. If values look wrong, check the 'Unrecognized columns' and 'SAMPLE_ID' sections.\n")
