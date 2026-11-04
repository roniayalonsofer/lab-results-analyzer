"""
core/form_fields.py
-------------------
Field forms that are handwritten scans can't be read by the app. The user reads
them from the scan shown next to a table, chooses the fields wanted in the report,
and types the values per sample.
"""

import fitz

# Fields of Adama's soil gas forms: the custody page, then the per-sample purge form.
# Other fields can be added by name in the app.
SOIL_GAS_FIELDS = [
    "בדיקת IPA",
    "PID [ppm]",
    "עומק הגשש [מ']",
    "שעת התחלת דיגום",
    "שעת סיום דיגום",
    "ואקום בקניסטר - התחלה",
    "ואקום בקניסטר - סיום",
    "Shut In - לחץ התחלה",
    "Shut In - לחץ סוף",
    "ספיקת המשאבה [ml/min]",
    "משך שאיבת השטיפה [דקות]",
    "1 או 20 ppbv",
    "PID באוהל לפני IPA",
    "PID באוהל אחרי IPA",
    "IPA בראש הקידוח",
]

# Fields answered by a mark, not a typed value
MARK_FIELDS = {"בדיקת IPA"}


def custody_page(raw_pdf: bytes) -> int:
    """1-based page to show first: the custody page, by its title if the page has
    text (which arrives reversed), else the second to last page."""
    with fitz.open(stream=raw_pdf, filetype="pdf") as doc:
        for number, page in enumerate(doc, 1):
            text = page.get_text()
            if "שרשרת משמורת" in text or "תרומשמ תרשרש" in text:
                return number
        return max(len(doc) - 1, 1)


def page_png(raw_pdf: bytes, page_number: int, dpi: int = 130) -> bytes:
    with fitz.open(stream=raw_pdf, filetype="pdf") as doc:
        return doc[page_number - 1].get_pixmap(dpi=dpi).tobytes("png")


def page_count(raw_pdf: bytes) -> int:
    with fitz.open(stream=raw_pdf, filetype="pdf") as doc:
        return len(doc)


def apply_form_values(records: list[dict], by_sample: dict[str, dict[str, str]]) -> None:
    """Attach the typed form values to the report's records, per sample."""
    for r in records:
        if r["sample_id"] in by_sample:
            r["form_values"] = by_sample[r["sample_id"]]
