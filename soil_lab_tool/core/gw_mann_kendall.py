"""
GSI Mann-Kendall Toolkit (.xls) for the groundwater periodic report.

The owner's MK workbook holds 2,500+ formulas and embedded charts, and no
Python library can write an old .xls without losing them. So the file is
only READ here (owner's decision, 2026-11-10):

  * read each constituent sheet (wells from the "Sampling Point ID:" row,
    events from the "Sampling Event / Date" block — never fixed positions),
  * add the new round in memory,
  * compute COV, S, the confidence factor and the trend the way the toolkit
    does (the confidence comes from the workbook's own Kendall_Dist sheet),
  * draw the toolkit page (header, data table, statistics, log chart) as a
    picture for the Word report, and
  * give the rows to paste into the .xls by hand.

Site convention seen in the workbook: a result below 0.01 mg/L, or a
non-detect, is entered as 0.01 ("ריכוזים קטנים מ-0.01 מוצגים בתחתית
הגרף"). "-" means not measured.
"""
from __future__ import annotations

import io
import math
from dataclasses import dataclass, field
from datetime import datetime, timedelta

import xlrd

MK_FLOOR = 0.01
_EXCEL_EPOCH = datetime(1899, 12, 30)

# Constituent cell text → report parameter name
_CONSTITUENT_PARAM = {
    "MTBE": "MTBE",
    "BENZEN": "בנזן",
    "BENZENE": "בנזן",
    "TOLUENE": "טולואן",
    "ETHYLBENZENE": "אתיל בנזן",
    "XYLENE": "קסילן",
    "XYLENES": "קסילן",
}

PARAM_DISPLAY = {
    "MTBE": "MTBE",
    "בנזן": "BENZEN",
    "טולואן": "TOLUENE",
    "אתיל בנזן": "ETHYLBENZENE",
    "קסילן": "XYLENES",
}


def excel_to_date(n: float) -> datetime:
    return _EXCEL_EPOCH + timedelta(days=int(n))


def date_to_excel(d: datetime) -> int:
    return (datetime(d.year, d.month, d.day) - _EXCEL_EPOCH).days


@dataclass
class MKSheet:
    name: str
    constituent: str            # as written in the sheet ("BENZEN")
    param: str | None           # report parameter name ("בנזן")
    facility: str
    conducted_by: str
    units: str
    wells: list                 # ["חק-2", ...]
    well_cols: list             # column index of each well
    events: list = field(default_factory=list)  # [(datetime, {well: float|None|"-"})]
    event_rows: list = field(default_factory=list)  # sheet row (0-based) of each event
    first_event_row: int = 0
    next_free_row: int | None = None  # first empty row after the last event
    event_slots: int = 0        # rows reserved for events before the statistics
    cached_stats: dict = field(default_factory=dict)  # {well: {"cov","s","cf","trend"}}


def _txt(v) -> str:
    return str(v).strip() if v is not None else ""


def _find_label(ws, label: str):
    for r in range(ws.nrows):
        for c in range(ws.ncols):
            if _txt(ws.cell_value(r, c)).rstrip(":").lower() == label.lower():
                return r, c
    return None


def _value_right_of(ws, r, c):
    for cc in range(c + 1, ws.ncols):
        v = ws.cell_value(r, cc)
        if _txt(v):
            return v
    return ""


def _read_sheet(ws) -> MKSheet | None:
    pos = _find_label(ws, "Constituent")
    sp = _find_label(ws, "Sampling Point ID")
    if not pos or not sp:
        return None
    constituent = _txt(_value_right_of(ws, *pos))
    fac = _find_label(ws, "Facility Name")
    by = _find_label(ws, "Conducted By")
    un = _find_label(ws, "Concentration Units")

    r_sp, c_sp = sp
    wells, cols = [], []
    for c in range(c_sp + 1, ws.ncols):
        w = _txt(ws.cell_value(r_sp, c))
        if not w:
            break
        wells.append(w)
        cols.append(c)
    if not wells:
        return None

    cov = _find_label(ws, "Coefficient of Variation")
    stats_row = cov[0] if cov else ws.nrows

    # Event block: column of "Sampling Event" numbers, date one column right.
    ev_col = cols[0] - 2
    first, slots = None, 0
    for r in range(r_sp + 1, stats_row):
        v = ws.cell_value(r, ev_col)
        if isinstance(v, float) and v == int(v) and v >= 1:
            if first is None:
                first = r
            slots += 1
    if first is None:
        return None

    sheet = MKSheet(
        name=ws.name, constituent=constituent,
        param=_CONSTITUENT_PARAM.get(constituent.upper().replace(" ", "")),
        facility=_txt(_value_right_of(ws, *fac)) if fac else "",
        conducted_by=_txt(_value_right_of(ws, *by)) if by else "",
        units=_txt(_value_right_of(ws, *un)) if un else "mg/L",
        wells=wells, well_cols=cols, first_event_row=first, event_slots=slots,
    )
    for r in range(first, first + slots):
        d = ws.cell_value(r, ev_col + 1)
        if not isinstance(d, float) or d <= 0:
            if sheet.next_free_row is None:
                sheet.next_free_row = r
            continue
        vals = {}
        for w, c in zip(wells, cols):
            v = ws.cell_value(r, c)
            if isinstance(v, float):
                vals[w] = v
            elif _txt(v) == "-":
                vals[w] = "-"
            else:
                vals[w] = None
        if all(v is None for v in vals.values()):
            continue  # a leftover date with no results; the toolkit ignores it too
        sheet.events.append((excel_to_date(d), vals))
        sheet.event_rows.append(r)
        sheet.next_free_row = None

    if cov:
        labels = {"cov": "Coefficient of Variation", "s": "Mann-Kendall Statistic (S)",
                  "cf": "Confidence Factor", "trend": "Concentration Trend"}
        rows = {}
        for k, lab in labels.items():
            hit = _find_label(ws, lab)
            if hit:
                rows[k] = hit[0]
        for w, c in zip(wells, cols):
            sheet.cached_stats[w] = {k: ws.cell_value(r, c) for k, r in rows.items()}
    return sheet


@dataclass
class MKWorkbook:
    sheets: list
    kendall_dist: dict  # {(abs_s, n): probability}


def read_mk_workbook(xls_bytes: bytes) -> MKWorkbook:
    wb = xlrd.open_workbook(file_contents=xls_bytes)
    sheets, dist = [], {}
    for ws in wb.sheets():
        if ws.name.strip().lower() == "kendall_dist":
            ns = ws.row_values(0)
            for r in range(1, ws.nrows):
                s = ws.cell_value(r, 0)
                if not isinstance(s, float):
                    continue
                for c in range(1, ws.ncols):
                    n, p = ns[c], ws.cell_value(r, c)
                    if isinstance(n, float) and isinstance(p, float):
                        dist[(int(s), int(n))] = p
            continue
        sh = _read_sheet(ws)
        if sh:
            sheets.append(sh)
    if not sheets:
        raise ValueError("לא נמצאו בקובץ גיליונות של GSI Mann-Kendall (חסרה השורה Sampling Point ID).")
    return MKWorkbook(sheets=sheets, kendall_dist=dist)


def report_sheets(wb: MKWorkbook) -> dict:
    """
    The sheet used for each parameter in the report: the workbook keeps old
    periods too ("MTBE (2010-2019)"), so take the sheet whose last event is
    the most recent; on a tie, the one with more events.
    """
    best: dict = {}
    for sh in wb.sheets:
        if not sh.param or not sh.events:
            continue
        key = (max(d for d, _ in sh.events), len(sh.events))
        if sh.param not in best or key > best[sh.param][0]:
            best[sh.param] = (key, sh)
    return {p: v[1] for p, v in best.items()}


def mk_value(result, measured: bool = True):
    """Lab result (mg/L, None = non-detect) → the value entered in the toolkit."""
    if not measured:
        return "-"
    if result is None:
        return MK_FLOOR
    return result if result >= MK_FLOOR else MK_FLOOR


def add_round(sheet: MKSheet, when: datetime, values: dict) -> bool:
    """
    Add (or replace, when the date is already there) the round's values.
    `values` maps the sheet's well names to the toolkit value. Returns False
    when the sheet has no free event row left.
    """
    for i, (d, _) in enumerate(sheet.events):
        if d.date() == when.date():
            sheet.events[i] = (when, values)
            return True
    if sheet.next_free_row is None:
        return False
    row = sheet.next_free_row
    sheet.events.append((when, values))
    sheet.event_rows.append(row)
    nxt = row + 1
    sheet.next_free_row = nxt if nxt < sheet.first_event_row + sheet.event_slots else None
    return True


def _series(sheet: MKSheet, well: str) -> list:
    return [v[well] for _, v in sheet.events
            if isinstance(v.get(well), float) and not math.isnan(v[well])]


def mk_statistic(values: list) -> int:
    s = 0
    for i in range(len(values) - 1):
        for j in range(i + 1, len(values)):
            d = values[j] - values[i]
            s += (d > 0) - (d < 0)
    return s


def coefficient_of_variation(values: list) -> float | None:
    n = len(values)
    if n < 2:
        return None
    mean = sum(values) / n
    if mean == 0:
        return None
    sd = math.sqrt(sum((v - mean) ** 2 for v in values) / (n - 1))  # Excel STDEV
    return sd / mean


def confidence(s: int, n: int, dist: dict) -> float | None:
    """Confidence that the trend is real, 1 - p, as the toolkit reports it."""
    if n < 4:
        return None
    p = dist.get((abs(s), n))
    if p is None and n <= 40 and dist:
        # beyond the table's last S for this n: p is essentially zero
        known = [k for k in dist if k[1] == n]
        if known and abs(s) > max(k[0] for k in known):
            p = 0.0
    if p is None:
        var = n * (n - 1) * (2 * n + 5) / 18
        z = (abs(s) - 1) / math.sqrt(var) if s else 0.0
        p = 0.5 * math.erfc(z / math.sqrt(2))
    return 1 - p


def trend(s: int, cf: float | None, cov: float | None) -> str:
    """GSI toolkit decision rules."""
    if cf is None:
        return ""
    if s > 0:
        if cf > 0.95:
            return "Increasing"
        if cf >= 0.90:
            return "Prob. Increasing"
        return "No Trend"
    if s < 0 and cf > 0.95:
        return "Decreasing"
    if s < 0 and cf >= 0.90:
        return "Prob. Decreasing"
    if cov is not None and cov < 1:
        return "Stable"
    return "No Trend"


def sheet_stats(sheet: MKSheet, dist: dict) -> dict:
    out = {}
    for w in sheet.wells:
        vals = _series(sheet, w)
        n = len(vals)
        s = mk_statistic(vals)
        cov = coefficient_of_variation(vals)
        cf = confidence(s, n, dist)
        out[w] = {"n": n, "s": s, "cov": cov, "cf": cf, "trend": trend(s, cf, cov)}
    return out


def format_cf(cf: float | None) -> str:
    if cf is None:
        return ""
    if cf >= 0.999:
        return ">99.9%"
    return f"{cf * 100:.1f}%"


# ── rows to paste into the .xls ─────────────────────────────────────────────

def paste_rows_xlsx(added: list, when: datetime = None) -> bytes:
    """
    A small workbook listing the rows to type into the owner's .xls: for
    each (sheet, round date) the sheet's row number, event number, date and
    one value per well in the sheet's own column order. `added` may also be
    a list of sheets, with `when` the round date (one round).
    """
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    if added and isinstance(added[0], MKSheet):
        added = [(sh, when) for sh in added]
    wb = Workbook()
    ws = wb.active
    ws.title = "שורות להדבקה"
    ws.sheet_view.rightToLeft = True
    bold = Font(bold=True)
    fill = PatternFill("solid", fgColor="FFF2CC")
    dates = sorted({d for _, d in added})
    ws.append(["שורות חדשות לקובץ Mann-Kendall — דיגום " + ", ".join(f"{d:%d.%m.%Y}" for d in dates)])
    ws["A1"].font = Font(bold=True, size=12)
    ws.append(["בכל גיליון: להקליד את הערכים בשורה שמספרה מופיע כאן, ולשמור. "
               "הנוסחאות והגרפים בקובץ יתעדכנו לבד."])
    ws.append([])
    for sh, d in added:
        idx = next((i for i, (ed, _) in enumerate(sh.events) if ed.date() == d.date()), None)
        if idx is None:
            continue
        ev_date, vals = sh.events[idx]
        row0 = sh.event_rows[idx]
        event_no = row0 - sh.first_event_row + 1
        ws.append([f"גיליון: {sh.name}", f"שורה {row0 + 1}"])
        ws.cell(ws.max_row, 1).font = bold
        ws.append(["Sampling Event", "Sampling Date"] + sh.wells)
        for c in range(1, len(sh.wells) + 3):
            ws.cell(ws.max_row, c).font = bold
        row = [event_no, ev_date.strftime("%d/%m/%Y")] + [
            vals.get(w) if vals.get(w) is not None else "" for w in sh.wells]
        ws.append(row)
        for c in range(1, len(row) + 1):
            ws.cell(ws.max_row, c).fill = fill
            ws.cell(ws.max_row, c).alignment = Alignment(horizontal="center")
        ws.append([])
    ws.column_dimensions["A"].width = 26
    for col in "BCDEFGHIJ":
        ws.column_dimensions[col].width = 14
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── the toolkit page as a picture ───────────────────────────────────────────

_NAVY = "#1f3864"
_HEAD = "#808080"
_WELL_COLORS = ["#000080", "#c0504d", "#9bbb59", "#8064a2", "#4bacc6", "#f79646"]
_MARKERS = ["D", "s", "^", "x", "*", "o"]


_BIDI_MODE = None   # "rle": matplotlib orders it, inside an RTL embedding; "manual"


def _visual_rtl(text: str) -> str:
    """'חק-2' → '2-קח', 'סונול החוף' → 'ףוחה לונוס' (for a renderer with no bidi)."""
    import re
    parts = re.split(r"([A-Za-z0-9][A-Za-z0-9.,/]*)", text)
    out = [p if re.match(r"[A-Za-z0-9]", p or " ") else p[::-1] for p in parts]
    return "".join(reversed(out))


def _with_mode(text: str, mode: str) -> str:
    if mode == "rle":
        return "‫" + text + "‬"   # RIGHT-TO-LEFT EMBEDDING … POP
    return _visual_rtl(text)


def _hebrew_on_the_right(drawn: str) -> bool:
    """Draw *drawn* and tell whether its ink leans right (where Hebrew goes)."""
    import numpy as np
    import matplotlib
    matplotlib.use("Agg")
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    fig = Figure(figsize=(4, 1), dpi=100)
    canvas = FigureCanvasAgg(fig)
    fig.text(0.5, 0.5, drawn, fontsize=30, ha="center", va="center", family="DejaVu Sans")
    canvas.draw()
    cols = (255 - np.asarray(canvas.buffer_rgba())[:, :, 0].astype(float)).sum(axis=0)
    xs = np.nonzero(cols > 0)[0]
    if len(xs) == 0:
        return False
    return (cols * np.arange(len(cols))).sum() / cols.sum() > (xs[0] + xs[-1]) / 2


def _bidi_mode() -> str:
    """
    How to hand Hebrew to matplotlib on THIS machine, found by drawing: 3.11
    lays out bidi text itself (libraqm), but the server's build takes every
    line as left to right, so 'פפ-1' came out as 'פפ-1' instead of '1-פפ'.
    Two test words, alef (wide) then yods, and alefs then '-1': written
    right, the alefs end up on the right.
    """
    global _BIDI_MODE
    if _BIDI_MODE is None:
        _BIDI_MODE = "manual"
        try:
            for mode in ("rle", "manual"):
                if all(_hebrew_on_the_right(_with_mode(w, mode)) for w in ("אייייי", "אאא-1")):
                    _BIDI_MODE = mode
                    break
        except Exception:
            pass
    return _BIDI_MODE


def _mpl_does_bidi() -> bool:
    return _bidi_mode() == "rle"


def _bidi(text: str) -> str:
    """A Hebrew label, ready for matplotlib to draw as a right-to-left line."""
    import re
    if not re.search(r"[֐-׿]", text):
        return text
    return _with_mode(text, "rle" if _mpl_does_bidi() else "manual")


def _fmt_val(v) -> str:
    if v is None:
        return ""
    if isinstance(v, str):
        return v
    if v >= 100:
        return f"{v:g}"
    return f"{v:.4g}"


def render_sheet_png(sheet: MKSheet, stats: dict, evaluation_date: datetime) -> bytes:
    """Draw the toolkit page (as the report shows it) and return PNG bytes."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.dates as mdates
    import matplotlib.pyplot as plt
    from matplotlib.patches import Rectangle

    plt.rcParams.update({"font.family": "DejaVu Sans", "axes.unicode_minus": False})
    wells = sheet.wells
    n_rows = len(sheet.events) + 1  # the events and one empty row, as the report shows

    W, H = 7.0, 9.6
    fig = plt.figure(figsize=(W, H), dpi=160)
    tab = fig.add_axes([0, 0.30, 1, 0.70])
    tab.set_xlim(0, 100)
    tab.set_ylim(0, 100)
    tab.axis("off")

    def box(x, y, w, h, fc="white", ec="#4f6228", lw=0.6):
        tab.add_patch(Rectangle((x, y), w, h, facecolor=fc, edgecolor=ec, linewidth=lw))

    def text(x, y, s, size=6.5, color="black", weight="normal", ha="center"):
        tab.text(x, y, s, fontsize=size, color=color, fontweight=weight, ha=ha, va="center")

    # title band
    box(0, 92, 100, 8, fc="#365f91", ec="#365f91")
    text(50, 97.4, "GSI MANN-KENDALL TOOLKIT", size=12, color="white", weight="bold")
    text(50, 94.2, "for Constituent Trend Analysis", size=9, color="white", weight="bold")

    # header fields
    y0 = 89.6
    text(14, y0, "Evaluation Date:", weight="bold", ha="right")
    text(15, y0, evaluation_date.strftime("%d.%m.%Y"), color=_NAVY, ha="left")
    text(14, y0 - 2, "Facility Name:", weight="bold", ha="right")
    text(15, y0 - 2, _bidi(sheet.facility), color=_NAVY, ha="left")
    text(14, y0 - 4, "Conducted By:", weight="bold", ha="right")
    text(15, y0 - 4, _bidi(sheet.conducted_by), color=_NAVY, ha="left")
    text(76, y0 - 2, "Constituent:", weight="bold", ha="right")
    text(77, y0 - 2, sheet.constituent, color=_NAVY, ha="left")
    text(76, y0 - 4, "Concentration Units:", weight="bold", ha="right")
    text(77, y0 - 4, sheet.units, color=_NAVY, ha="left")

    # well header
    x_ev, x_dt, x0 = 0.5, 9.5, 19.0
    col_w = min(13.0, (100 - x0 - 1) / max(len(wells), 1))
    y = 82.0
    text(x0 - 1, y, "Sampling Point ID:", weight="bold", ha="right")
    for i, w in enumerate(wells):
        box(x0 + i * col_w, y - 0.9, col_w, 1.8)
        text(x0 + i * col_w + col_w / 2, y, _bidi(w), color=_NAVY, weight="bold")

    # data table
    row_h = min(1.75, 66.0 / (n_rows + 3))
    y = 80.0
    box(x_ev, y - 2 * row_h, x_dt - x_ev, 2 * row_h, fc=_HEAD)
    box(x_dt, y - 2 * row_h, x0 - x_dt, 2 * row_h, fc=_HEAD)
    text((x_ev + x_dt) / 2, y - row_h, "Sampling\nEvent", size=5.5, color="white", weight="bold")
    text((x_dt + x0) / 2, y - row_h, "Sampling\nDate", size=5.5, color="white", weight="bold")
    box(x0, y - 2 * row_h, col_w * len(wells), 2 * row_h, fc=_HEAD)
    text(x0 + col_w * len(wells) / 2, y - row_h,
         f"{sheet.constituent.upper()} CONCENTRATION ({sheet.units})", color="white", weight="bold")
    y -= 2 * row_h
    for i in range(n_rows):
        y -= row_h
        box(x_ev, y, x_dt - x_ev, row_h)
        box(x_dt, y, x0 - x_dt, row_h)
        text((x_ev + x_dt) / 2, y + row_h / 2, str(i + 1), size=5.5)
        ev = sheet.events[i] if i < len(sheet.events) else None
        if ev:
            text((x_dt + x0) / 2, y + row_h / 2, f"{ev[0].day}-{ev[0]:%b-%y}",
                 size=5.5, color=_NAVY)
        for j, w in enumerate(wells):
            box(x0 + j * col_w, y, col_w, row_h)
            if ev:
                text(x0 + j * col_w + col_w / 2, y + row_h / 2, _fmt_val(ev[1].get(w)),
                     size=5.5, color=_NAVY)

    # statistics
    labels = [("Coefficient of Variation:", lambda st: "" if st["cov"] is None else f"{st['cov']:.2f}"),
              ("Mann-Kendall Statistic (S):", lambda st: f"{st['s']:d}"),
              ("Confidence Factor:", lambda st: format_cf(st["cf"])),
              ("Concentration Trend:", lambda st: st["trend"])]
    for k, (lab, fn) in enumerate(labels):
        y -= row_h * (1.25 if k == 3 else 1.0)
        h = row_h * (1.25 if k == 3 else 1.0)
        fc = "#365f91" if k == 3 else "#7f7f7f"
        box(x_ev, y, x0 - x_ev, h, fc=fc)
        text(x0 - 1, y + h / 2, lab, size=5.5, color="white", weight="bold", ha="right")
        for j, w in enumerate(wells):
            box(x0 + j * col_w, y, col_w, h, fc=fc)
            text(x0 + j * col_w + col_w / 2, y + h / 2, fn(stats[w]), size=5.8,
                 color="white", weight="bold")

    # chart
    ax = fig.add_axes([0.11, 0.035, 0.66, 0.235])
    for j, w in enumerate(wells):
        pts = [(d, v[w]) for d, v in sheet.events if isinstance(v.get(w), float)]
        if not pts:
            continue
        xs, ys = zip(*pts)
        ax.plot(xs, ys, marker=_MARKERS[j % len(_MARKERS)], markersize=3.5, linewidth=1.4,
                color=_WELL_COLORS[j % len(_WELL_COLORS)], label=_bidi(w))
    ax.set_yscale("log")
    ax.set_ylim(MK_FLOOR, 1000)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%m/%y"))
    ax.grid(True, which="major", color="#c0c0c0", linewidth=0.5)
    ax.set_xlabel("Sampling Date", fontsize=8, fontweight="bold")
    ax.set_ylabel(f"Concentration ({sheet.units})", fontsize=8, fontweight="bold")
    ax.tick_params(labelsize=6)
    ax.legend(loc="center left", bbox_to_anchor=(1.03, 0.5), fontsize=6.5, frameon=True)

    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=160)
    plt.close(fig)
    return buf.getvalue()
