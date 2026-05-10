"""
Reproduce the Excel build for a large XRF dataset (743 samples x 24 compounds).
Run from soil_lab_tool/ directory.
"""
import sys, io, traceback
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, ".")

from core.excel_output import LabReportExcel
from core.threshold_manager import ThresholdManager
import os

# Minimal ThresholdManager — load real one if available
THRESH = os.path.join("thresholds", "soil_vsl_tier1_v7_2024.xlsx")
if os.path.exists(THRESH):
    tm = ThresholdManager(THRESH)
    print("ThresholdManager loaded from file")
else:
    tm = None
    print("No threshold file found — using None (will skip threshold lookups)")

# Build 743 samples x 24 compounds = 17,832 records
COMPOUNDS = [
    ("Arsenic",     "7440-38-2"),
    ("Barium",      "7440-39-3"),
    ("Cadmium",     "7440-43-9"),
    ("Chromium",    "7440-47-3"),
    ("Cobalt",      "7440-48-4"),
    ("Copper",      "7440-50-8"),
    ("Iron",        "7439-89-6"),
    ("Lead",        "7439-92-1"),
    ("Manganese",   "7439-96-5"),
    ("Mercury",     "7439-97-6"),
    ("Molybdenum",  "7439-98-7"),
    ("Nickel",      "7440-02-0"),
    ("Rubidium",    "7440-17-7"),
    ("Silver",      "7440-22-4"),
    ("Strontium",   "7440-24-6"),
    ("Thorium",     "7440-29-1"),
    ("Tin",         "7440-31-5"),
    ("Titanium",    "7440-32-6"),
    ("Uranium",     "7440-61-1"),
    ("Vanadium",    "7440-62-2"),
    ("Yttrium",     "7440-65-5"),
    ("Zinc",        "7440-66-6"),
    ("Zirconium",   "7440-67-7"),
    ("Calcium",     "7440-70-2"),
]

import random
random.seed(42)

records = []
for i in range(743):
    loc = f"BH-{(i // 5) + 1:03d}"
    sample_id = f"S-{i+1:04d} – {loc}"
    for compound, cas in COMPOUNDS:
        val = round(random.uniform(0.1, 500), 2)
        flag = "" if random.random() > 0.1 else "ND"
        records.append({
            "sample_id":     sample_id,
            "compound":      compound,
            "cas":           cas,
            "value":         None if flag == "ND" else val,
            "flag":          flag,
            "unit":          "mg/kg",
            "lod":           None,
            "loq":           None,
            "analysis_type": "SOIL_METALS",
        })

print(f"Records: {len(records):,}  Samples: 743  Compounds: {len(COMPOUNDS)}")

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
