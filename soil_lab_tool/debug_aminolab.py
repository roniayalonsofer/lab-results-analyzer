"""
debug_aminolab.py
-----------------
Run the AminolabGroundwaterParser against a local PDF and print every record
it produces.  Mirrors the style of test_issues.py.

Usage (from the soil_lab_tool directory):
    python debug_aminolab.py "C:\path\to\aminolab_report.pdf"
"""
import sys
import os
import io

sys.path.insert(0, os.path.dirname(__file__))

from parsers.groundwater.aminolab import AminolabGroundwaterParser

if len(sys.argv) < 2:
    print("Usage: python debug_aminolab.py <path_to_aminolab_pdf>")
    sys.exit(1)

pdf_path = sys.argv[1]
if not os.path.exists(pdf_path):
    print(f"File not found: {pdf_path}")
    sys.exit(1)

print(f"PDF: {pdf_path}  ({os.path.getsize(pdf_path):,} bytes)")
print("=" * 70)

with open(pdf_path, "rb") as f:
    data = f.read()

parser = AminolabGroundwaterParser(debug=True)
records = parser.parse(io.BytesIO(data))

print("\n" + "=" * 70)
print(f"RECORDS PRODUCED: {len(records)}")
print("=" * 70)

if not records:
    print("  (none)")
else:
    # Column widths
    fmt = "{:<35} {:<15} {:>12}  {:<6}  {:<10}  {}"
    print(fmt.format("compound", "analysis_type", "value", "flag", "unit", "sample_id"))
    print("-" * 100)
    for r in records:
        val = r.get("value")
        val_str = f"{val:.4g}" if isinstance(val, float) else (str(val) if val is not None else "")
        print(fmt.format(
            str(r.get("compound", ""))[:35],
            str(r.get("analysis_type", ""))[:15],
            val_str,
            str(r.get("flag", "")),
            str(r.get("unit", "")),
            str(r.get("sample_id", "")),
        ))

# Summary by analysis_type
from collections import Counter
by_type = Counter(r.get("analysis_type", "?") for r in records)
print("\nSUMMARY BY analysis_type:")
for atype, cnt in by_type.most_common():
    print(f"  {atype:<20} {cnt} records")
