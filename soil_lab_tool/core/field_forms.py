"""
core/field_forms.py
-------------------
Reads PID readings from field forms (Excel or PDF) into the map the reports use:
{borehole key: [(depth_to, pid), ...]}.

Layout of the survey summary form ("סיכום סקר"): a header row with "קידוח",
"עד עומק (מ')" and "PID (ppm)"; the borehole name appears only on its first row.
Columns are found by their header, since a PDF table can arrive in reverse order.
"""

import io

import pandas as pd

from core.excel_output import _pid_key

# Column positions of the form, for an Excel file with no header row
_NO_HEADER_COLUMNS = (0, 2, 7)


def _clean(cell) -> str:
    return " ".join(str(cell).split()) if cell is not None else ""


def _is_pid(cell: str) -> bool:
    return "PID" in cell.upper()


def _is_borehole(cell: str) -> bool:
    low = cell.lower()
    return "קידוח" in cell or "חודיק" in cell or "borehole" in low or "well" in low


def _is_depth_to(cell: str) -> bool:
    low = cell.lower()
    return (("עד" in cell and "עומק" in cell) or ("דע" in cell and "קמוע" in cell)
            or "depth to" in low or "to depth" in low)


def _find_columns(rows: list[list]) -> tuple[int, tuple[int, int, int]] | None:
    """Return (header row index, (borehole, depth to, PID columns)), if there is a header."""
    for ri, row in enumerate(rows):
        cells = [_clean(c) for c in row]
        pid = next((i for i, c in enumerate(cells) if _is_pid(c)), None)
        borehole = next((i for i, c in enumerate(cells) if _is_borehole(c)), None)
        depth_to = next((i for i, c in enumerate(cells) if _is_depth_to(c)), None)
        if pid is not None and borehole is not None and depth_to is not None:
            return ri, (borehole, depth_to, pid)
    return None


def _to_float(cell) -> float | None:
    try:
        val = float(_clean(cell))
    except ValueError:
        return None
    return None if pd.isna(val) else val


def _read_rows(rows: list[list], columns: tuple[int, int, int], fix_text=None) -> dict:
    bh_col, depth_col, pid_col = columns
    pid_data: dict = {}
    borehole = ""
    for row in rows:
        if len(row) <= max(bh_col, depth_col):
            continue
        name = _clean(row[bh_col])
        if name and name.lower() not in ("nan", "none"):
            # "Ķ" is a typing artifact some forms use for the plain "K"
            borehole = (fix_text(name) if fix_text else name).replace("Ķ", "K").replace("ķ", "k")
        depth_to = _to_float(row[depth_col])
        pid = _to_float(row[pid_col]) if pid_col < len(row) else None
        if borehole and depth_to is not None and pid is not None:
            pid_data.setdefault(_pid_key(borehole), []).append((depth_to, pid))
    return pid_data


def _read_excel(raw: bytes) -> dict:
    xl = pd.ExcelFile(io.BytesIO(raw))
    sheets = [xl.parse(s, header=None, dtype=object).values.tolist() for s in xl.sheet_names]
    pid_data: dict = {}
    for rows in sheets:
        found = _find_columns(rows)
        if found:
            merge_field_data(pid_data, _read_rows(rows[found[0] + 1:], found[1]))
    if not pid_data and sheets:
        pid_data = _read_rows(sheets[0], _NO_HEADER_COLUMNS)
    return pid_data


def _read_pdf(raw: bytes) -> dict:
    import pdfplumber

    from parsers.soil.bactochem import fix_rtl_name

    pid_data: dict = {}
    columns = None
    with pdfplumber.open(io.BytesIO(raw)) as pdf:
        if not any((page.extract_text() or "").strip() for page in pdf.pages):
            raise ValueError("הקובץ סרוק כתמונה, ואי אפשר לקרוא ממנו קריאות PID. "
                             "הדוח ייבנה בלי PID. כדי לכלול אותן, העלו טופס ב-Excel.")
        for page in pdf.pages:
            for rows in page.extract_tables():
                found = _find_columns(rows)
                if found:
                    columns = found[1]
                    rows = rows[found[0] + 1:]
                if columns:  # a table without a header continues the previous page's
                    merge_field_data(pid_data, _read_rows(rows, columns, fix_text=fix_rtl_name))
    return pid_data


def parse_field_form(filename: str, raw: bytes) -> dict:
    """Read one field form. Raises ValueError when it holds no PID readings."""
    if filename.lower().endswith(".pdf"):
        pid_data = _read_pdf(raw)
        if not pid_data:
            raise ValueError("לא נמצאה ב-PDF טבלה עם עמודות קידוח, עד עומק ו-PID. ייתכן שהקובץ סרוק.")
    else:
        pid_data = _read_excel(raw)
        if not pid_data:
            raise ValueError("לא נמצאו בקובץ קריאות PID.")
    return pid_data


def merge_field_data(target: dict, new: dict) -> dict:
    """Add *new* to *target*. A later form replaces the reading at the same depth."""
    for key, entries in new.items():
        by_depth = dict(target.get(key, []))
        by_depth.update(entries)
        target[key] = sorted(by_depth.items())
    return target
