import fitz
import glob

matches = glob.glob('C:/Users/RoniAyalon/AppData/Local/Temp/41111-PDF-*.pdf')
path = matches[0]
doc = fitz.open(path)

# Render page 0 as PNG
page = doc[0]
mat = fitz.Matrix(2, 2)  # 2x zoom
pix = page.get_pixmap(matrix=mat)
out = 'C:/Users/RoniAyalon/lab-results-analyzer/tmp_page1.png'
pix.save(out)
print(f"Saved page 1 image: {out}")

# Also dump all text blocks with ASCII-range chars for pattern analysis
print("\n--- ASCII-range strings in blocks (may be numbers/identifiers) ---")
for i, page in enumerate(doc):
    raw = page.get_text()
    lines = raw.split('\n')
    print(f"\nPage {i+1}:")
    for line in lines:
        printable = ''.join(c if 32 <= ord(c) < 127 else '·' for c in line)
        if printable.strip():
            print(f"  {repr(line[:60])} -> {printable[:60]}")
