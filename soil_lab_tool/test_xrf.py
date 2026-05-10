import sys, io
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, ".")
from parsers.soil.xrf import XRFSoilParser

csv_data = (
    "Sample ID,Location,Mo,Pb,As,Zn,Cu,Ni,Fe,Cr,Ba\n"
    "S-01,Well A,2.1,45.3,12.5,89.2,34.1,22.8,32000,45.6,345\n"
    "S-02,Well B,<0.5,23.1,8.2,67.4,28.9,19.3,28000,38.2,280\n"
    "S-03,Well C,3.4,ND,15.8,110.5,41.2,28.1,35000,52.3,420\n"
)
parser = XRFSoilParser()
records = parser.parse(io.BytesIO(csv_data.encode("utf-8")))
print(f"Total records: {len(records)}")
for r in records[:12]:
    sid = r["sample_id"]
    cmp = r["compound"]
    val = r["value"]
    flg = r["flag"]
    unit = r["unit"]
    cas = r["cas"]
    print(f"  {sid:<20} {cmp:<12} {val}  {flg:<6} {unit}  CAS={cas}")

# Test auto_detect_lab
from parsers import auto_detect_lab, auto_detect_category
b = csv_data.encode("utf-8")
lab = auto_detect_lab("xrf_results.csv", b)
cat = auto_detect_category("xrf_results.csv", b)
print(f"\nauto_detect_lab('xrf_results.csv') -> {lab!r}")
print(f"auto_detect_category('xrf_results.csv') -> {cat!r}")

lab2 = auto_detect_lab("metals_report.csv", b)
cat2 = auto_detect_category("metals_report.csv", b)
print(f"auto_detect_lab('metals_report.csv') -> {lab2!r}  (content-based)")
print(f"auto_detect_category('metals_report.csv') -> {cat2!r}")
