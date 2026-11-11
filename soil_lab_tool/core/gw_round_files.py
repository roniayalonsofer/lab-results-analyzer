"""
Sort the files of one or more groundwater sampling rounds, so the user can
drop a whole "תוצאות מעבדה" folder without choosing a lab.

Every file is identified by READING it (the lab's own parser must return
samples), not by its name: an ALS appendix is not always recognised by
auto_detect_lab, and field-contractor PDFs look like nothing at all.

    sort_round_files([(name, bytes), ...]) -> {
        "rounds":  [Round, ...]   oldest first,
        "mk":      (name, bytes) | None,
        "ignored": [(name, reason), ...],
        "notes":   [str, ...],
    }
"""
from __future__ import annotations

import io
import os
import tempfile
from dataclasses import dataclass, field
from datetime import datetime

LAB_NAMES = {"als": "ALS", "bactochem": "בקטוכם", "aminolab": "אמינולאב"}


@dataclass
class Round:
    date: datetime
    lab: str                       # "als" | "bactochem" | "aminolab"
    lab_files: list = field(default_factory=list)      # [(name, bytes)] used for the results
    samples: dict = field(default_factory=dict)        # {sample: {"results", "limits", ...}}
    field_form: tuple | None = None                    # (name, bytes)
    companion_files: list = field(default_factory=list)  # names of copies not read (PDF of an xlsx)

    @property
    def date_str(self) -> str:
        return self.date.strftime("%d.%m.%y")

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
        if low.endswith(".xlsx") and gri.is_yeda_field_form(data):
            form = gri.parse_yeda_field_form(data)
            d = form.get("sampling_date")
            return {"kind": "field_form", "date": datetime(d.year, d.month, d.day) if d else None,
                    "form": form}
        if low.endswith(".xls"):
            try:
                from soil_lab_tool.core.gw_mann_kendall import read_mk_workbook
                read_mk_workbook(data)
                return {"kind": "mk"}
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
            return {"kind": "ignored",
                    "reason": "טפסי שדה ב-PDF (לנספח). הנתונים נקראים מטופס ביצוע העבודה ב-Excel"}
        return {"kind": "ignored", "reason": "לא זוהה כתעודת מעבדה"}

    return {"kind": "ignored", "reason": "סוג קובץ לא נתמך"}


def sort_round_files(files: list) -> dict:
    rounds: dict = {}
    forms, ignored, notes = [], [], []
    mk = None
    for name, data in files:
        try:
            info = identify(name, data)
        except Exception as e:  # a broken file must not stop the others
            info = {"kind": "ignored", "reason": f"שגיאה בקריאה: {e}"}
        kind = info["kind"]
        if kind == "ignored":
            ignored.append((name, info["reason"]))
        elif kind == "mk":
            if mk:
                ignored.append((name, "כבר הועלה קובץ Mann-Kendall אחר"))
            else:
                mk = (name, data)
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
    for name, data, d in forms:
        match = [r for r in ordered if d and r.date.date() == d.date()]
        if not match and len(ordered) == 1 and d is None:
            match = ordered
        if match:
            if match[0].field_form:
                ignored.append((name, "כבר הועלה טופס שדה לאותו דיגום"))
            else:
                match[0].field_form = (name, data)
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
    return {"rounds": ordered, "mk": mk, "ignored": ignored, "notes": notes}
