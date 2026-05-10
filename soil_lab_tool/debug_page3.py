import sys, io
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, ".")
import pdfplumber

_BC_FP_TOKENS = frozenset({
    "סמומ", "ןצמח", "תוכילומ", "הרוטרפמט", "סקודר", "תוריכע",
    "קמוע", "הביאש", "LOWFLOW", "pH",
})
_ROW_TOL = 6

PDF = r"C:\Users\r5901\Downloads\D021125-0179-01297598-HE-F.pdf"

with open(PDF, "rb") as f:
    data = f.read()

with pdfplumber.open(io.BytesIO(data)) as pdf:
    page = pdf.pages[2]
    words = page.extract_words(x_tolerance=4, y_tolerance=4) or []
    # cluster
    groups = []
    for w in sorted(words, key=lambda x: (x["top"], x["x0"])):
        if not groups or abs(w["top"] - groups[-1][0]["top"]) > _ROW_TOL:
            groups.append([w])
        else:
            groups[-1].append(w)

    print(f"Page 3: {len(groups)} row groups\n")
    for i, g in enumerate(groups):
        texts = [w["text"] for w in g]
        hit = any(t in _BC_FP_TOKENS for t in texts)
        print(f"  row {i:2d} y={g[0]['top']:5.1f}  {'HIT ' if hit else '    '}  {texts}")
