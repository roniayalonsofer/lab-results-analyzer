"""
Sort the files of one or more groundwater sampling rounds, so the user can
drop a whole "תוצאות מעבדה" folder without choosing a lab.

Every file is identified by READING it (the lab's own parser must return
samples), not by its name: an ALS appendix is not always recognised by
auto_detect_lab, and field-contractor PDFs look like nothing at all.

    sort_round_files([(name, bytes), ...]) -> {
        "rounds":  [Round, ...]   oldest first,
        "mk":      (name, bytes) | None,
        "ignored": [(name, reason), ...],   files that could not be used
        "side":    [(name, reason), ...],   known papers the update doesn't need
        "notes":   [str, ...],
    }
"""
from __future__ import annotations

import io
import os
import tempfile
from dataclasses import dataclass, field
from datetime import datetime

_ROUND_DAYS = 7  # samples this close together belong to one round

LAB_NAMES = {"als": "ALS", "bactochem": "בקטוכם", "aminolab": "אמינולאב"}


@dataclass
class Round:
    date: datetime
    lab: str                       # "als" | "bactochem" | "aminolab"
    lab_files: list = field(default_factory=list)      # [(name, bytes)] used for the results
    samples: dict = field(default_factory=dict)        # {sample: {"results", "limits", ...}}
    field_form: tuple | None = None                    # (name, bytes), the first form
    companion_files: list = field(default_factory=list)  # names of copies not read (PDF of an xlsx)
    end_date: datetime | None = None                   # last day of a round sampled over days
    field_forms: list = field(default_factory=list)    # [(name, bytes)], one per sampling day

    @property
    def date_str(self) -> str:
        return self.date.strftime("%d.%m.%y")

    @property
    def span_str(self) -> str:
        """'03.09.25' or '03-04.09.25' for a round sampled over two days."""
        end = self.end_date or self.date
        if end.date() == self.date.date():
            return self.date_str
        if (end.month, end.year) == (self.date.month, self.date.year):
            return f"{self.date:%d}-{end:%d.%m.%y}"
        return f"{self.date_str}-{end:%d.%m.%y}"

    @property
    def form_bytes(self):
        """The field forms' bytes for run_update_bytes, or None."""
        return [b for _, b in self.field_forms] or None

    @property
    def lab_input(self):
        """What run_update_bytes expects for this lab."""
        if self.lab == "aminolab":
            return [b for _, b in self.lab_files]
        return self.lab_files[0][1]

    @property
    def wells(self) -> list:
        import re
        def num(s):
            m = re.search(r"(\d+)\s*$", s)
            return int(m.group(1)) if m else 0
        return sorted(self.samples, key=num)


def _parse_date(s):
    for fmt in ("%d.%m.%y", "%d.%m.%Y", "%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(str(s).strip(), fmt)
        except (ValueError, TypeError):
            continue
    return None


def _pdf_text(data: bytes, pages: int = 3) -> str:
    # PyMuPDF first: a scanned page takes pdfplumber seconds, PyMuPDF a few ms.
    try:
        import fitz
        with fitz.open(stream=data, filetype="pdf") as d:
            return "\n".join(d[i].get_text() for i in range(min(pages, len(d))))
    except Exception:
        pass
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            return "\n".join((p.extract_text() or "") for p in pdf.pages[:pages])
    except Exception:
        return ""


def _with_tmp_pdf(data: bytes, fn):
    fd, path = tempfile.mkstemp(suffix=".pdf")
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(data)
        return fn(path)
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


def _is_report_form(data: bytes) -> bool:
    """The Water Authority report form: compound codes 2009 (MTBE) and 84 (benzene) in column A."""
    try:
        import openpyxl
        ws = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True).worksheets[0]
        codes = {str(r[0]).strip() for r in ws.iter_rows(min_col=1, max_col=1, max_row=80, values_only=True)
                 if r and r[0] is not None}
        return {"2009", "84"} <= codes
    except Exception:
        return False


def identify(name: str, data: bytes) -> dict:
    """
    {"kind": "als"|"bactochem"|"aminolab"|"field_form"|"mk"|"ignored",
     "date": datetime|None, "samples": {...}, "reason": str}
    """
    from soil_lab_tool.core import gw_round_inputs as gri

    low = name.lower()
    if low.endswith(".csv"):
        from soil_lab_tool.gw_report_updater import parse_bactochem_csv
        try:
            lab = parse_bactochem_csv(data)
        except Exception:
            lab = {}
        if lab.get("samples"):
            return {"kind": "bactochem", "date": _parse_date(lab.get("sampling_date")),
                    "samples": lab["samples"]}
        return {"kind": "ignored", "reason": "קובץ CSV שלא זוהה כתוצאות בקטוכם"}

    if low.endswith((".xlsx", ".xls", ".xlsm")):
        if low.endswith((".xlsx", ".xlsm")) and _is_report_form(data):
            return {"kind": "report_form"}
        if low.endswith(".xlsx") and gri.is_yeda_field_form(data):
            form = gri.parse_yeda_field_form(data)
            d = form.get("sampling_date")
            return {"kind": "field_form", "date": datetime(d.year, d.month, d.day) if d else None,
                    "form": form}
        if low.endswith(".xls"):
            try:
                from soil_lab_tool.core.gw_mann_kendall import read_mk_workbook
                wb = read_mk_workbook(data)
                dates = [d for sh in wb.sheets for d, _ in sh.events]
                return {"kind": "mk", "last": max(dates) if dates else None}
            except Exception:
                pass
        try:
            rnd = gri.parse_als_round(data, name)
            return {"kind": "als", "date": _parse_date(rnd["sampling_date"]), "samples": rnd["samples"]}
        except Exception:
            pass
        return {"kind": "ignored", "reason": "לא זוהה כתעודת מעבדה, טופס שדה או קובץ Mann-Kendall"}

    if low.endswith(".pdf"):
        text = _pdf_text(data)
        if "ALS" in text and ("CERTIFICATE OF ANALYSIS" in text or "Work Order" in text):
            try:
                rnd = gri.parse_als_round(data, name)
                return {"kind": "als", "date": _parse_date(rnd["sampling_date"]),
                        "samples": rnd["samples"]}
            except Exception:
                pass
        # Yeda Water's field sheets ("טפסי עבודה", heavy scans): known from
        # their first pages, before any lab reader goes through every page.
        is_yeda = "ע.ד.י" in text or "yeda-water" in text.lower()
        if is_yeda and "CAS" not in text and "CERTIFICATE" not in text.upper():
            return {"kind": "side",
                    "reason": "טפסי שדה ב-PDF (לנספח). הנתונים נקראים מטופס ביצוע העבודה ב-Excel"}
        from soil_lab_tool.gw_report_updater import parse_aminolab_pdf, parse_bactochem_pdf
        for kind, fn in (("bactochem", parse_bactochem_pdf), ("aminolab", parse_aminolab_pdf)):
            try:
                lab = _with_tmp_pdf(data, fn)
            except Exception:
                continue
            samples = lab.get("samples")
            if kind == "aminolab" and not samples and lab.get("well_id"):
                samples = {lab["well_id"]: lab}
            date = _parse_date(lab.get("sampling_date") or lab.get("date"))
            if samples and date:
                return {"kind": kind, "date": date, "samples": samples}
        if "ע.ד.י" in text or "yeda-water" in text.lower():
            return {"kind": "side",
                    "reason": "טפסי שדה ב-PDF (לנספח). הנתונים נקראים מטופס ביצוע העבודה ב-Excel"}
        if not text.strip():
            return {"kind": "side",
                    "reason": "מסמך סרוק (למשל טופס נטילה מהשטח), לנספח. לא צריך אותו לעדכון"}
        if "bactochem" in text.lower():
            return {"kind": "side",
                    "reason": "טופס של בקטוכם בלי תוצאות (שרשרת משמורת / קבלת דוגמאות), לנספח. "
                              "לא צריך אותו לעדכון"}
        return {"kind": "ignored", "reason": "לא זוהה כתעודת מעבדה"}

    return {"kind": "ignored", "reason": "סוג קובץ לא נתמך"}


def sort_round_files(files: list) -> dict:
    rounds: dict = {}
    forms, ignored, notes, side = [], [], [], []
    mk, mk_last = None, None
    report_form = None
    for name, data in files:
        try:
            info = identify(name, data)
        except Exception as e:  # a broken file must not stop the others
            info = {"kind": "ignored", "reason": f"שגיאה בקריאה: {e}"}
        kind = info["kind"]
        if kind == "ignored":
            ignored.append((name, info["reason"]))
        elif kind == "side":
            side.append((name, info["reason"]))
        elif kind == "mk":
            # Several toolkit files (old per-well templates next to the real
            # one): use the one whose data reaches the latest date.
            last = info.get("last")
            if mk is None or (last and (mk_last is None or last > mk_last)):
                if mk:
                    side.append((mk[0], "קובץ Mann-Kendall ישן יותר מקובץ אחר שהועלה"))
                mk, mk_last = (name, data), last
            else:
                side.append((name, "קובץ Mann-Kendall ישן יותר מקובץ אחר שהועלה"))
        elif kind == "report_form":
            if report_form:
                side.append((name, "כבר הועלה טופס דיווח אחר"))
            else:
                report_form = (name, data)
        elif kind == "field_form":
            forms.append((name, data, info.get("date")))
        else:
            if info.get("date") is None:
                ignored.append((name, "לא נמצא תאריך דיגום בתעודה"))
                continue
            key = (kind, info["date"].date())
            rnd = rounds.get(key)
            if rnd is None:
                rnd = rounds[key] = Round(date=info["date"], lab=kind)
            is_pdf = name.lower().endswith(".pdf")
            if kind == "als" and rnd.lab_files:
                # The same certificate as PDF and as Excel: read the Excel.
                prev_is_pdf = rnd.lab_files[0][0].lower().endswith(".pdf")
                if prev_is_pdf and not is_pdf:
                    rnd.companion_files.append(rnd.lab_files[0][0])
                    rnd.lab_files = [(name, data)]
                    rnd.samples = info["samples"]
                else:
                    rnd.companion_files.append(name)
                continue
            if kind == "bactochem":
                # The certificate as PDF and as CSV: the PDF wins (it has the
                # sampler's name); the CSV is read only for wells the PDF lacks.
                is_csv = name.lower().endswith(".csv")
                if is_csv:
                    if rnd.samples and set(info["samples"]) <= set(rnd.samples):
                        rnd.companion_files.append(name)
                        continue
                    rnd.lab_files.append((name, data))
                    for k, v in info["samples"].items():
                        rnd.samples.setdefault(k, v)
                    continue
                rnd.lab_files.append((name, data))
                rnd.samples.update(info["samples"])
                for fn, fb in list(rnd.lab_files):
                    if fn.lower().endswith(".csv"):
                        try:
                            from soil_lab_tool.gw_report_updater import parse_bactochem_csv
                            wells = set(parse_bactochem_csv(fb)["samples"])
                        except Exception:
                            wells = set()
                        pdf_wells = set(info["samples"])
                        if wells <= pdf_wells:
                            rnd.lab_files.remove((fn, fb))
                            rnd.companion_files.append(fn)
                continue
            rnd.lab_files.append((name, data))
            for k, v in info["samples"].items():
                rnd.samples.setdefault(k, v)

    ordered = sorted(rounds.values(), key=lambda r: r.date)
    # One round sampled over a few days (03.09 and 04.09) is one round: each
    # sample keeps its own date for the history tables.
    merged = []
    for r in ordered:
        for v in r.samples.values():
            v.setdefault("date", r.date_str)
        last = merged[-1] if merged else None
        if last and last.lab == r.lab and (r.date - (last.end_date or last.date)).days <= _ROUND_DAYS:
            last.lab_files += r.lab_files
            last.companion_files += r.companion_files
            for k, v in r.samples.items():
                last.samples.setdefault(k, v)
            last.end_date = r.date
        else:
            merged.append(r)
    ordered = merged
    form_days: dict = {}
    for name, data, d in forms:
        match = [r for r in ordered
                 if d and r.date.date() <= d.date() <= (r.end_date or r.date).date()]
        if not match and len(ordered) == 1 and d is None:
            match = ordered
        if match:
            day = d.date() if d else None
            if day in form_days.setdefault(id(match[0]), set()):
                ignored.append((name, "כבר הועלה טופס שדה לאותו דיגום"))
            else:
                form_days[id(match[0])].add(day)
                match[0].field_forms.append((name, data))
                match[0].field_form = match[0].field_form or (name, data)
        else:
            ds = d.strftime("%d.%m.%y") if d else "ללא תאריך"
            ignored.append((name, f"טופס שדה מ-{ds}, ואין תעודת מעבדה מאותו תאריך"))

    for r in ordered:
        if r.lab == "bactochem" and r.lab_files and all(
                n.lower().endswith(".csv") for n, _ in r.lab_files):
            notes.append(f"בדיגום {r.date_str} התוצאות נקראו מקובץ ה-CSV של בקטוכם. אין בו את שם "
                         "הדוגם, ולכן משפט הדוגם בדוח לא עודכן. אפשר להעלות גם את תעודת ה-PDF.")

    # A lab may name a lone sample after the site ("Hof Station"); when the
    # round's field form has exactly one well, that sample is that well.
    import re
    from soil_lab_tool.core.gw_round_inputs import parse_yeda_field_form
    for r in ordered:
        nameless = [k for k in r.samples if not re.search(r"\d+\s*$", k)]
        if len(r.samples) == 1 and nameless and r.field_form:
            wells = list(parse_yeda_field_form(r.field_form[1])["wells"])
            if len(wells) == 1:
                r.samples = {wells[0]: r.samples[nameless[0]]}
                notes.append(f"בדיגום {r.date_str} הדגימה '{nameless[0]}' שויכה לבאר {wells[0]} "
                             "לפי טופס השדה.")
    return {"rounds": ordered, "mk": mk, "ignored": ignored, "notes": notes, "side": side,
            "report_form": report_form}
