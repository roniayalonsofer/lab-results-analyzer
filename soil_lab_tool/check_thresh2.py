import sys; sys.path.insert(0, '.')
import pandas as pd

vsl_path = r'thresholds/soil_vsl_v7_full.xlsx'
xl = pd.ExcelFile(vsl_path)

# Print header rows of Industrial RBTL to find column structure
sname = 'Tier 1 - Industrial RBTL'
raw = xl.parse(sname, header=None, dtype=str).fillna('')
print(f"=== {sname} first 8 rows ===")
for i in range(8):
    vals = list(raw.iloc[i].values)[:12]
    print(f"  row {i}: {vals}")

# Print the C10-C40 row with all column values
print()
for i, row in raw.iterrows():
    if 'C10-C40' in str(row.values):
        print(f"  row {i} full: {list(row.values)}")
        break

# Check what the column scanner actually finds
sv_indoor_col = sv_outdoor_col = None
soil_vh_col = soil_hm_0_6_col = soil_hm_6_col = soil_low_col = None
for ri in range(min(8, len(raw))):
    for ci, v in enumerate(raw.iloc[ri].values):
        vs = str(v).strip().lower()
        if not vs or ci < 2:
            continue
        if "soil vapor" in vs and "indoor" in vs and sv_indoor_col is None:
            sv_indoor_col = ci
            print(f"  Found sv_indoor at row={ri} col={ci}: {v!r}")
        elif "soil vapor" in vs and "outdoor" in vs and sv_outdoor_col is None:
            sv_outdoor_col = ci
            print(f"  Found sv_outdoor at row={ri} col={ci}: {v!r}")
        if "very high" in vs and soil_vh_col is None:
            soil_vh_col = ci
            print(f"  Found soil_vh at row={ri} col={ci}: {v!r}")
        if "0-6" in vs and soil_hm_0_6_col is None:
            soil_hm_0_6_col = ci
            print(f"  Found soil_hm_0_6 at row={ri} col={ci}: {v!r}")
        if ">6" in vs and soil_hm_6_col is None:
            soil_hm_6_col = ci
            print(f"  Found soil_hm_6 at row={ri} col={ci}: {v!r}")
        if "low" in vs and "sensit" in vs and soil_low_col is None:
            soil_low_col = ci
            print(f"  Found soil_low at row={ri} col={ci}: {v!r}")

print(f"\n  Final columns: vh={soil_vh_col} hm0_6={soil_hm_0_6_col} hm_6={soil_hm_6_col} low={soil_low_col}")
print(f"  sv_indoor={sv_indoor_col} (must not be None for data to load)")
