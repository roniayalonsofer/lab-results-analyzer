"""
Extras for a monitoring report that are not lab results:

- well photos: the photo tables ("תאריך – 03.09.2024" / "T-1 | T-2" / the
  pictures) get the new photos, matched to a well by the file name
  ("P-1.jpg"); an old photo with no new one is removed and its date cell
  marked yellow. The signatures on the cover (also pictures in a table) are
  never touched: a photo table is one with a row of well names right above
  a row of pictures.

- reports of the year (soil survey, cleanliness verification, ...): the
  report's title, date and its summary ("סיכום") become one paragraph,
  added at the end of the background chapter, in yellow for review.
  Nothing leaves the app (no outside service reads the report).
"""
from __future__ import annotations

import io
import re
from datetime import datetime

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn


# ── Well photos ─────────────────────────────────────────────────────────────

def _cell_text(tc) -> str:
    return "".join(t.text or "" for t in tc.iter(qn("w:t"))).strip()


def _photo_tables(doc):
    """[(table, header_row_index|None, names_row_index, pics_row_index)]."""
    from soil_lab_tool.gw_report_updater import _well_num
    out = []
    for tbl in doc.tables:
        rows = tbl.rows
        for ri in range(1, len(rows)):
            pics = [tc for tc in rows[ri]._tr.findall(qn("w:tc")) if tc.findall(".//" + qn("a:blip"))]
            if not pics:
                continue
            names = [_cell_text(tc) for tc in rows[ri - 1]._tr.findall(qn("w:tc"))]
            if names and all(n and _well_num(n) and re.search(r"[A-Za-zא-ת]", n) for n in names):
                head = ri - 2 if ri >= 2 and "תארי" in " ".join(
                    _cell_text(tc) for tc in rows[ri - 2]._tr.findall(qn("w:tc"))) else None
                out.append((tbl, head, ri - 1, ri))
                break
    return out


def _photo_for(well: str, photos: list):
    """The uploaded photo whose file name names this well ("P-1.jpg", "צילום T-2.jpeg")."""
    from soil_lab_tool.gw_report_updater import _wkey
    key = _wkey(well)
    for name, data in photos:
        stem = re.sub(r"\.[A-Za-z]+$", "", name)
        for m in re.finditer(r"([A-Za-zא-ת]{1,4})\s*[-_ ]?\s*(\d+)", stem):
            if _wkey(f"{m.group(1)}-{m.group(2)}") == key:
                return data
    return None


def _set_text(tc, text: str, yellow: bool = True):
    ts = list(tc.iter(qn("w:t")))
    if not ts:
        return
    ts[0].text = text
    for t in ts[1:]:
        t.text = ""
    if yellow:
        from docx.text.run import Run
        Run(ts[0].getparent(), None).font.highlight_color = WD_COLOR_INDEX.YELLOW


def _fit_picture(doc, inline, data: bytes):
    """The new photo in the old photo's frame: no wider and no taller than the old one."""
    from PIL import Image
    ext = inline.find(qn("wp:extent"))
    box_w, box_h = int(ext.get("cx")), int(ext.get("cy"))
    rId, _ = doc.part.get_or_add_image(io.BytesIO(data))
    inline.find(".//" + qn("a:blip")).set(qn("r:embed"), rId)
    w, h = Image.open(io.BytesIO(data)).size
    scale = min(box_w / w, box_h / h)
    cx, cy = int(w * scale), int(h * scale)
    for e in [ext] + [x for x in inline.iter(qn("a:ext")) if x.get("cx") is not None]:
        e.set("cx", str(cx))
        e.set("cy", str(cy))


def update_well_photos(word_bytes: bytes, photos: list, well_dates: dict, report: dict = None) -> bytes:
    """
    photos: [(file name, image bytes)]; well_dates: {well: datetime} (the
    latest sampling of each well, written in the date row).
    """
    from soil_lab_tool.gw_report_updater import _set_well_ambiguity, _wkey
    if report is None:
        report = {}
    doc = Document(io.BytesIO(word_bytes))
    tables = _photo_tables(doc)
    replaced, removed = [], []
    _set_well_ambiguity([_cell_text(tc) for tbl, _, nr, _ in tables
                         for tc in tbl.rows[nr]._tr.findall(qn("w:tc"))] + list(well_dates))
    dates_by_key = {_wkey(w): d for w, d in well_dates.items()}
    for tbl, head, nr, pr in tables:
        names = [_cell_text(tc) for tc in tbl.rows[nr]._tr.findall(qn("w:tc"))]
        cells = tbl.rows[pr]._tr.findall(qn("w:tc"))
        shown = {}
        for well, tc in zip(names, cells):
            data = _photo_for(well, photos)
            drawings = tc.findall(".//" + qn("w:drawing"))
            if data is not None and drawings:
                inline = drawings[0].find(qn("wp:inline"))
                if inline is None:
                    inline = drawings[0].find(qn("wp:anchor"))
                _fit_picture(doc, inline, data)
                for extra in drawings[1:]:
                    extra.getparent().remove(extra)
                replaced.append(well)
                d = dates_by_key.get(_wkey(well))
                shown[well] = f"{d:%d.%m.%Y}" if d else ""
            else:
                for dr in drawings:
                    dr.getparent().remove(dr)
                removed.append(well)
                shown[well] = ""
        if head is not None:
            hcells = tbl.rows[head]._tr.findall(qn("w:tc"))
            if len(hcells) == len(names):
                for well, tc in zip(names, hcells):
                    _set_text(tc, f"תאריך – {shown[well]}")
            else:  # one merged cell over the photos
                parts = [f"{w} – {shown[w]}" for w in names]
                _set_text(hcells[0], "תאריכים; " + ", ".join(parts))
    report["photos_replaced"] = replaced
    report["photos_removed"] = removed
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ── Reports of the year as a paragraph ──────────────────────────────────────

_MONTHS = ["ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני", "יולי", "אוגוסט",
           "ספטמבר", "אוקטובר", "נובמבר", "דצמבר"]


def _report_lines(name: str, data: bytes) -> list:
    if name.lower().endswith(".docx"):
        d = Document(io.BytesIO(data))
        return [p.text.strip() for p in d.paragraphs if p.text.strip()]
    import fitz
    with fitz.open(stream=data, filetype="pdf") as pdf:
        text = "\n".join(page.get_text() for page in pdf)
    return [ln.strip() for ln in text.splitlines() if ln.strip()]


def report_paragraph(name: str, data: bytes) -> str:
    """'<title>, <month year>: <summary>' from a report's own title and summary section."""
    lines = _report_lines(name, data)
    if not lines:
        raise ValueError("לא נמצא טקסט בקובץ (אולי סרוק).")
    head = lines[:25]
    stem = re.sub(r"\.[A-Za-z]+$", "", name).replace("_", " ").strip()
    if len(re.findall(r"[א-ת]{2,}", stem)) >= 2:
        title = stem   # a file name in words reads right; PDF text can come in reversed order
    else:
        title = max((ln for ln in head if re.search(r'דו"?ח|דוח|סקר|וידוא|סיכום ממצאים|חוות דעת', ln)),
                    key=len, default=stem)
    title = re.sub(r"\s+", " ", title).strip(" -–:")
    when = None
    for ln in head:
        m = re.search(r"(" + "|".join(_MONTHS) + r")\s+(\d{4})", ln)
        if m:
            when = f"{m.group(1)} {m.group(2)}"
            break
        m = re.search(r"\b(\d{1,2})[./](\d{1,2})[./](\d{4})\b", ln)
        if m:
            when = f"{_MONTHS[int(m.group(2)) - 1]} {m.group(3)}"
            break
    # the summary section: after a "סיכום" heading, up to "המלצות" or a new chapter
    def norm(ln):
        return re.sub(r"\s+", " ", ln.replace(":", " ")).strip(" .")
    heads = ("סיכום", "סיכום ומסקנות", "סיכום והמלצות", "מסקנות", "תקציר")
    start = next((i for i, ln in enumerate(lines)
                  if re.sub(r"^\d+\.?\s*", "", norm(ln)) in heads and "...." not in ln), None)
    if start is None:
        start = next((i for i, ln in enumerate(lines) if ln.startswith(("סיכום", "מסקנות", "תקציר"))), None)
    sentences = []
    if start is not None:
        for ln in lines[start + 1:start + 60]:
            n = re.sub(r"^\d+\.?\s*", "", norm(ln))
            if n in ("המלצות", "המלצה", "נספחים", "מקורות") or ln.startswith("-סוף"):
                break
            if re.fullmatch(r"\d+\.", ln.strip()):   # a list number ("3."), not a value ("12")
                if sentences and not sentences[-1].endswith((";", ".")):
                    sentences[-1] += ";"
                continue
            sentences.append(ln.lstrip("-•· ").strip())
    summary = " ".join(sentences)
    summary = re.sub(r"\s+", " ", summary).strip()
    if len(summary) > 900:
        cut = summary[:900]
        summary = cut[:cut.rfind(";") + 1 or cut.rfind(".") + 1 or 900]
    if not summary:
        summary = "(לא נמצא בדו\"ח פרק סיכום — יש להשלים)"
    lead = f"{title}" + (f", {when}" if when else "")
    return f"{lead}: {summary}"


def add_report_paragraphs(word_bytes: bytes, reports: list, report: dict = None) -> bytes:
    """reports: [(file name, bytes)] → one yellow paragraph each, at the end of the background
    chapter (before "מהלך הדיגום ..."), or before the summary when there is no such heading."""
    if report is None:
        report = {}
    doc = Document(io.BytesIO(word_bytes))
    paras = doc.paragraphs
    anchor = next((p for p in paras if p.text.strip().startswith("מהלך הדיגום")), None)
    if anchor is None:
        anchor = next((p for p in paras if p.text.strip().startswith("סיכום")), None)
    template = None
    if anchor is not None:
        i = paras.index(anchor)
        template = next((paras[j] for j in range(i - 1, 0, -1) if len(paras[j].text.strip()) > 40), None)
    added, failed = [], []
    for name, data in reports:
        try:
            text = report_paragraph(name, data)
        except Exception as e:
            failed.append((name, str(e)))
            continue
        import copy
        if template is not None:
            new_p = copy.deepcopy(template._p)
            for r in new_p.findall(qn("w:r"))[1:]:
                new_p.remove(r)
            for el in list(new_p):
                if el.tag in (qn("w:hyperlink"), qn("w:ins"), qn("w:del")):
                    new_p.remove(el)
        else:
            from docx.oxml import OxmlElement
            new_p = OxmlElement("w:p")
        if anchor is not None:
            anchor._p.addprevious(new_p)
        else:
            doc.element.body.append(new_p)
        from docx.text.paragraph import Paragraph
        p = Paragraph(new_p, anchor._parent if anchor is not None else doc._body)
        if p.runs:
            p.runs[0].text = text
        else:
            p.add_run(text)
        p.runs[0].font.highlight_color = WD_COLOR_INDEX.YELLOW
        added.append(name)
    report["reports_added"] = added
    report["reports_failed"] = failed
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()
