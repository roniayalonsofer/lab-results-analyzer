"""
Run from the project root:
  py -3 debug_pfas_pdf.py "path\to\בקטוכם_פיפס.pdf"
"""
import io
import os
import sys

sys.path.insert(0, "soil_lab_tool")
os.environ["BACTOCHEM_DEBUG"] = "1"

pdf_path = sys.argv[1] if len(sys.argv) > 1 else None
if not pdf_path:
    print("Usage: py -3 debug_pfas_pdf.py <path_to_pdf>")
    sys.exit(1)

import pdfplumber

# ── Step 1: raw text extraction ───────────────────────────────────────────────
print("\n" + "=" * 70)
print("STEP 1: Raw text extracted by pdfplumber")
print("=" * 70)
with open(pdf_path, "rb") as fh:
    data = fh.read()

full_text = ""
with pdfplumber.open(io.BytesIO(data)) as pdf:
    print(f"Pages: {len(pdf.pages)}")
    for i, page in enumerate(pdf.pages):
        text = page.extract_text() or ""
        full_text += text + "\n"
        print(f"\n--- Page {i} ({len(text)} chars) ---")
        for j, line in enumerate(text.splitlines()[:40]):
            print(f"  {j:3}: {line!r}")
        if len(text.splitlines()) > 40:
            print(f"  ... ({len(text.splitlines()) - 40} more lines)")

# ── Step 2: sniffer flags ─────────────────────────────────────────────────────
print("\n" + "=" * 70)
print("STEP 2: PDF content sniff flags")
print("=" * 70)
ft_lo = full_text.lower()
has_pfas  = any(kw in ft_lo for kw in ("pfas", "pfoa", "pfos", "pfhxs", "ng/l"))
has_micro = any(kw in ft_lo for kw in ("cfu", "חיידקים", "coliform", "mpl", "mpn"))
has_btex  = any(kw in ft_lo for kw in ("benzene", "toluene", "mtbe"))

from parsers.groundwater.bactochem import _BC_FP_TOKENS
has_fp = any(kw in full_text for kw in _BC_FP_TOKENS)

print(f"  has_pfas  = {has_pfas}")
print(f"  has_micro = {has_micro}")
print(f"  has_btex  = {has_btex}")
print(f"  has_fp    = {has_fp}")

# ── Step 3: CAS line regex matches ───────────────────────────────────────────
print("\n" + "=" * 70)
print("STEP 3: Lines matching _GW_CAS_LINE_RE")
print("=" * 70)
from parsers.groundwater.bactochem import _GW_CAS_LINE_RE
matched = 0
for line in full_text.splitlines():
    m = _GW_CAS_LINE_RE.search(line)
    if m:
        matched += 1
        print(f"  MATCH unit={m.group('unit')!r:8} cas={m.group('cas')!r:15} "
              f"result={m.group('result')!r:20} compound={m.group('compound')!r}")
if matched == 0:
    print("  *** NO CAS lines matched — showing lines containing 'CAS' ***")
    for line in full_text.splitlines():
        if "CAS" in line or "cas" in line.lower():
            print(f"    {line!r}")

# ── Step 4: auto-detection ────────────────────────────────────────────────────
print("\n" + "=" * 70)
print("STEP 4: auto_detect_lab / auto_detect_category")
print("=" * 70)
from parsers import auto_detect_lab, auto_detect_category
fname = os.path.basename(pdf_path)
with open(pdf_path, "rb") as fh:
    bio = io.BytesIO(fh.read())
lab = auto_detect_lab(bio, fname)
bio.seek(0)
cat = auto_detect_category(bio, fname, lab)
print(f"  lab      = {lab!r}")
print(f"  category = {cat!r}")

# ── Step 5: full parse ────────────────────────────────────────────────────────
print("\n" + "=" * 70)
print("STEP 5: Full parser run")
print("=" * 70)
from parsers import get_parser
parser_cls = get_parser(lab, cat)
print(f"  Parser class: {parser_cls}")
if parser_cls:
    p = parser_cls(debug=True)
    with open(pdf_path, "rb") as fh:
        records = p.parse(io.BytesIO(fh.read()))
    print(f"\n  Total records: {len(records)}")
    for r in records[:20]:
        print(f"    {r['analysis_type']:20} | {r['compound']:30} | {r['value']} {r['unit']}")
    if len(records) > 20:
        print(f"  ... ({len(records) - 20} more)")
else:
    print("  *** No parser found for this lab/category combination ***")
