"""
Reproduce the Excel build for a large XRF dataset (743 samples x 24 compounds).
Tests both the full XRF parser pipeline and the Excel builder.
Run from soil_lab_tool/ directory.
"""
import sys, io, traceback, os
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, ".")

from core.excel_output import LabReportExcel
from core.threshold_manager import ThresholdManager

THRESH = os.path.join("thresholds", "soil_vsl_tier1_v7_2024.xlsx")
if os.path.exists(THRESH):
    tm = ThresholdManager(THRESH)
    print("ThresholdManager loaded from file")
else:
    tm = None
    print("No threshold file found — using None")

# ── Test 1: synthetic CSV through XRFSoilParser ──────────────────────────────
print("\n=== Test 1: Full pipeline via XRFSoilParser ===")
from parsers.soil.xrf import XRFSoilParser

import random; random.seed(42)
ELEMENTS = ["Mo","Pb","As","Zn","Cu","Ni","Fe","Cr","Ba","Mn",
            "V","Ti","Ca","K","Sr","Rb","U","Th","Zr","Y","Co","Ag","Cd","Sb"]
header = "Sample ID,Location," + ",".join(ELEMENTS)
rows = [header]
for i in range(743):
    loc = f"BH-{(i // 5) + 1:03d}"
    sid = f"S-{i+1:04d}"
    vals = []
    for _ in ELEMENTS:
        r = random.random()
        if r < 0.05:
            vals.append("ND")
        elif r < 0.10:
            vals.append(f"<{round(random.uniform(0.05, 0.5), 3)}")
        else:
            vals.append(str(round(random.uniform(0.1, 500), 2)))
    rows.append(f"{sid},{loc},{','.join(vals)}")

csv_bytes = "\n".join(rows).encode("utf-8")
parser = XRFSoilParser()
try:
    records = parser.parse(io.BytesIO(csv_bytes))
    from collections import Counter
    samples_count = len(set(r["sample_id"] for r in records))
    compounds_count = len(set(r["compound"] for r in records))
    print(f"Parser: {len(records):,} records  {samples_count} samples  {compounds_count} compounds")
    print(f"First sample_id: {records[0]['sample_id']!r}")
    atypes = Counter(r["analysis_type"] for r in records)
    print(f"Analysis types: {dict(atypes)}")
except Exception as e:
    print(f"Parser FAILED: {e}")
    traceback.print_exc()
    records = []

print(f"Records: {len(records):,}")

# Try to build the Excel
buf = io.BytesIO()
try:
    builder = LabReportExcel(
        records             = records,
        threshold_manager   = tm,
        output_path         = buf,
        project_name        = "Test XRF",
        client              = "Test Client",
        report_date         = "10.05.2026",
        selected_thresholds = ["VSL_SOIL"],
    )
    builder.build()
    size_kb = len(buf.getvalue()) / 1024
    print(f"SUCCESS — Excel built: {size_kb:.1f} KB")
except Exception as e:
    print(f"FAILED: {e}")
    traceback.print_exc()
