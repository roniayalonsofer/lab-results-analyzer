"""
debug_bactochem.py
------------------
Run the BactochemGroundwaterParser against a local file (PDF, CSV, or XLSX)
and print full debug output.

Usage (from the soil_lab_tool directory):
    python debug_bactochem.py "C:\\path\\to\\bactochem_report.pdf"
    python debug_bactochem.py "C:\\path\\to\\bactochem_report.xlsx"
"""
import sys
import os
import io

sys.path.insert(0, os.path.dirname(__file__))

from parsers.groundwater.bactochem import BactochemGroundwaterParser

if len(sys.argv) < 2:
    print("Usage: python debug_bactochem.py <path_to_file>")
    sys.exit(1)

path = sys.argv[1]
if not os.path.exists(path):
    print(f"File not found: {path}")
    sys.exit(1)

print(f"File : {path}")
print(f"Size : {os.path.getsize(path):,} bytes")
print(f"Ext  : {os.path.splitext(path)[1].lower()}")
print("=" * 70)

with open(path, "rb") as f:
    data = f.read()

# Show raw magic bytes so we can confirm the file type
print(f"Magic bytes (first 8): {data[:8]!r}")
print("=" * 70)

parser = BactochemGroundwaterParser(debug=True)
records = parser.parse(io.BytesIO(data))

print("\n" + "=" * 70)
print(f"TOTAL RECORDS: {len(records)}")
print("=" * 70)

if not records:
    print("  (none)")
else:
    fmt = "{:<35} {:<18} {:>12}  {:<6}  {:<10}  {}"
    print(fmt.format("compound", "analysis_type", "value", "flag", "unit", "sample_id"))
    print("-" * 105)
    for r in records:
        val = r.get("value")
        val_str = (f"{val:.4g}" if isinstance(val, float)
                   else (str(val) if val is not None else ""))
        print(fmt.format(
            str(r.get("compound", ""))[:35],
            str(r.get("analysis_type", ""))[:18],
            val_str,
            str(r.get("flag", "")),
            str(r.get("unit", "")),
            str(r.get("sample_id", "")),
        ))

from collections import Counter
print("\nSUMMARY BY analysis_type:")
for atype, cnt in Counter(r.get("analysis_type", "?") for r in records).most_common():
    print(f"  {atype:<22} {cnt} records")
