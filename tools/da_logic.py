"""
DA (Daily Allowance) engine - logic only, no UI.
================================================
Originally DA_Tool.pyw.  Reads a DA data sheet (.xlsx / .xls / .ods / .csv) and builds ONE
Excel workbook containing

  * one "[Name] DA [Month][Year]" bill sheet per person (same order as the data sheet)
  * a final "[Unit] DA Summary [Month][Year]" sheet that adds everything up

Expected columns in the data sheet (header names, any order of rows):
    Code | Name | Designation | Department | Areas ... (any number of columns) | DA
Everything between "Department" and "DA" is treated as a visit area.

Used by views/da_maker.py (web).  Command line (optional):
    python -m tools.da_logic data.xlsx --unit-code 1234 --unit-name Muridke --month 9 --year 2026

Requires: openpyxl, pandas, odfpy (.ods), xlrd (.xls)
"""
import argparse
import calendar
import io
import os
import random
import re
import sys
from dataclasses import dataclass
from datetime import date, timedelta

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins

# --------------------------------------------------------------------------
# Settings you may want to change
# --------------------------------------------------------------------------
DEFAULT_EXCLUDE_SUNDAYS = True     # starting state of the "Exclude Sundays" tick box in the window
DEFAULT_EXCLUDE_FRIDAYS = False    # starting state of the "Exclude Fridays" tick box in the window

# --------------------------------------------------------------------------
# Styles
# --------------------------------------------------------------------------
FONT = "Arial"
thin = Side(style="thin")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
RIGHT = Alignment(horizontal="right", vertical="center")
PANEL_FILL = PatternFill("solid", fgColor="DDEBF7")
ACCOUNTING = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)'


def f(size=10, bold=False, italic=False, color="000000"):
    return Font(name=FONT, size=size, bold=bold, italic=italic, color=color)


def box_range(ws, rng):
    for row in ws[rng]:
        for c in row:
            c.border = BOX


class ToolError(Exception):
    """A problem the user can fix (bad file, missing column, ...)."""


# --------------------------------------------------------------------------
# Input handling
# --------------------------------------------------------------------------
def parse_month(text):
    text = str(text).strip()
    if text.isdigit() and 1 <= int(text) <= 12:
        return int(text)
    for i in range(1, 13):
        if text.lower() in (calendar.month_name[i].lower(), calendar.month_abbr[i].lower()):
            return i
    raise ValueError(f"'{text}' is not a valid month (use a name like September / Sep, or 1-12).")


def ask(prompt, value=None):
    if value not in (None, ""):
        return str(value).strip()
    ans = ""
    while not ans:
        ans = input(prompt).strip()
    return ans


def load_data(source, filename=None):
    """`source` = file path, or raw bytes / file-like object (then pass `filename`
    so the format can be recognised from its extension)."""
    if isinstance(source, (bytes, bytearray)):
        source = io.BytesIO(source)
    ext = os.path.splitext(filename or (source if isinstance(source, str) else ""))[1].lower()
    if ext not in (".csv", ".ods", ".xlsx", ".xls", ".xlsm"):
        raise ToolError("Unsupported file type. Please use .xlsx, .xls, .ods or .csv.")
    try:
        if ext == ".csv":
            df = pd.read_csv(source, dtype=str, header=0)
        elif ext == ".ods":
            df = pd.read_excel(source, dtype=str, header=0, engine="odf")
        else:
            df = pd.read_excel(source, dtype=str, header=0)
    except ToolError:
        raise
    except Exception as e:                       # corrupt / wrong-format file
        raise ToolError(f"The data sheet could not be opened ({type(e).__name__}). "
                        "Check that it is a valid spreadsheet.") from e

    cols = [str(c).strip().lower() for c in df.columns]

    def col(name):
        if name not in cols:
            raise ToolError(f"Column '{name}' not found in the data sheet. Found: {list(df.columns)}")
        return cols.index(name)

    ic, in_, idg, idp, ida = col("code"), col("name"), col("designation"), col("department"), col("da")
    if ida <= idp:
        raise ToolError("The 'DA' column must come after the 'Department' column.")
    area_idx = list(range(idp + 1, ida))

    people = []
    for _, r in df.iterrows():
        name = str(r.iloc[in_]).strip() if pd.notna(r.iloc[in_]) else ""
        if not name:
            continue
        areas = [str(r.iloc[i]).strip() for i in area_idx if pd.notna(r.iloc[i]) and str(r.iloc[i]).strip()]
        try:
            rate = float(str(r.iloc[ida]).replace(",", "").strip())
        except ValueError:
            raise ToolError(f"DA rate for '{name}' is not a number: {r.iloc[ida]!r}")
        if rate.is_integer():
            rate = int(rate)
        people.append(dict(
            code=str(r.iloc[ic]).strip() if pd.notna(r.iloc[ic]) else "",
            name=name,
            designation=str(r.iloc[idg]).strip() if pd.notna(r.iloc[idg]) else "",
            department=str(r.iloc[idp]).strip() if pd.notna(r.iloc[idp]) else "",
            areas=areas,
            rate=rate,
        ))
    if not people:
        raise ToolError("No rows with a Name were found in the data sheet.")
    return people


# --------------------------------------------------------------------------
# Sheet names (Excel: max 31 characters, no  [ ] : * ? / \ )
# --------------------------------------------------------------------------
def clean(s):
    return re.sub(r"[\[\]:*?/\\]", "", s).strip()


def make_sheet_names(people, month, year):
    """Returns one unique sheet name per person; uses 'MonthYear', falling back to 'MonYear'
    for ALL sheets if any name would not fit in 31 characters."""
    for mtxt in (calendar.month_name[month], calendar.month_abbr[month]):
        tail = f" DA {mtxt}{year}"
        if all(len(clean(p["name"])) + len(tail) <= 31 for p in people):
            break
    names, used = [], set()
    for p in people:
        base = clean(p["name"])[: 31 - len(tail)].rstrip()
        cand, n = base + tail, 2
        while cand.lower() in used:            # two people with the same name
            suffix = f" ({n})"
            cand = base[: 31 - len(tail) - len(suffix)].rstrip() + suffix + tail
            n += 1
        used.add(cand.lower())
        names.append(cand)
    return names


def q(name):
    """Quote a sheet name for use inside a formula."""
    return "'" + name.replace("'", "''") + "'"



# --------------------------------------------------------------------------
# Random areas and leave parsing
# --------------------------------------------------------------------------
def random_areas(areas, working, rng):
    """One random area per calendar day.  `working` is a list of booleans (True = a day that will
    be claimed).  Claimed days draw from a shuffled 'bag', so the order is random, every area is
    used about equally often, and the same area never appears on two consecutive VISIT days.
    Days that are currently excluded also get a random area, so the date still has one if you
    later untick Sundays/Fridays or delete a leave."""
    if not areas:
        return []
    out, bag, last = [], [], None
    for is_working in working:
        if not is_working:
            out.append(rng.choice(areas))
            continue
        if not bag:
            bag = areas[:]
            rng.shuffle(bag)
            if len(bag) > 1 and bag[-1] == last:                # avoid an immediate repeat
                bag[0], bag[-1] = bag[-1], bag[0]
        last = bag.pop()
        out.append(last)
    return out


def _one_date(txt, month, year):
    """'5' -> 5th of the chosen month;  '5-9-2026' / '5/9/2026' -> that exact date."""
    parts = [x for x in re.split(r"[-/.]", txt.strip()) if x != ""]
    try:
        if len(parts) == 1:
            return date(year, month, int(parts[0]))
        if len(parts) == 3:
            d, m, y = (int(x) for x in parts)
            return date(y, m, d)
    except ValueError:
        pass
    raise ToolError(f"'{txt.strip()}' is not a valid date.")


def parse_leaves(text, month, year):
    """Turns text such as  '3, 5-7, 14'  or  '3-9-2026 to 5-9-2026, 14-9-2026'
    into a list of (from_date, to_date) tuples; consecutive days are merged into ranges."""
    days = set()
    for tok in re.split(r"[,;]", text or ""):
        tok = tok.strip()
        if not tok:
            continue
        m = re.fullmatch(r"(\d{1,2})\s*-\s*(\d{1,2})", tok)           # 5-7  -> days 5..7
        if m:
            a_, b_ = _one_date(m.group(1), month, year), _one_date(m.group(2), month, year)
        elif re.search(r"\bto\b", tok, re.I):                          # 3-9-2026 to 5-9-2026
            l, r_ = re.split(r"\bto\b", tok, flags=re.I)
            a_, b_ = _one_date(l, month, year), _one_date(r_, month, year)
        else:
            a_ = b_ = _one_date(tok, month, year)
        if b_ < a_:
            raise ToolError(f"Leave range '{tok}' ends before it starts.")
        d = a_
        while d <= b_:
            days.add(d)
            d += timedelta(days=1)
    for d in days:
        if (d.year, d.month) != (year, month):
            raise ToolError(f"Leave date {d:%d-%m-%Y} is outside {calendar.month_name[month]} {year}.")
    ranges = []
    for d in sorted(days):
        if ranges and d == ranges[-1][1] + timedelta(days=1):
            ranges[-1][1] = d
        else:
            ranges.append([d, d])
    return [(a, b) for a, b in ranges]


# --------------------------------------------------------------------------
# One DA sheet  (printable bill only)
# --------------------------------------------------------------------------
ROW_H, TOTAL_H, GAP_H = 18, 26, 14          # row heights (points) on the DA sheet
PRINT_SCALE = 100                           # print size in percent - same for every DA sheet
# Height of the longest possible bill (31 claimed days) = 3*20 + 24 + 18 + 18*31 + 26 + 3*14 + 18 = 746 pt,
# which fits one A4 page (about 784 pt usable) at 100%, so no shrinking is needed.


def build_da_sheet(wb, title, p, month, year, n_area_rows=None, opts=None, rng=None):
    ws = wb.create_sheet(title)
    opts = opts or {}
    rng = rng or random.Random()
    ndays = calendar.monthrange(year, month)[1]
    mname = calendar.month_name[month]
    ex_sun = opts.get("sun", DEFAULT_EXCLUDE_SUNDAYS)
    ex_fri = opts.get("fri", DEFAULT_EXCLUDE_FRIDAYS)

    # days that are left out of the bill
    leave_days = set()
    for d1, d2 in opts.get("leaves", []):
        d = d1
        while d <= d2:
            leave_days.add(d)
            d += timedelta(days=1)
    all_days = [date(year, month, i) for i in range(1, ndays + 1)]
    working = [not ((d.weekday() == 6 and ex_sun) or (d.weekday() == 4 and ex_fri) or d in leave_days)
               for d in all_days]
    areas = random_areas(p["areas"], working, rng)
    claimed = [(d, areas[i] if areas else None) for i, (d, w) in enumerate(zip(all_days, working)) if w]

    R_HDR, R_TITLE, R_FIRST = 4, 5, 6
    nrows = max(len(claimed), 1)          # exactly one row per claimed day - no empty rows
    R_LAST = R_FIRST + nrows - 1
    R_TOTAL = R_LAST + 1
    R_SIGN = R_TOTAL + 4

    for c, w in zip("ABCD", (15, 28, 28, 15)):
        ws.column_dimensions[c].width = w

    # Name / Designation / Department
    for r, (lab, val) in enumerate([("Name:", f"{p['code']} - {p['name']}".strip(" -")),
                                    ("Designation:", p["designation"]),
                                    ("Department:", p["department"])], start=1):
        ws.cell(r, 1, lab).font = f(10, True)
        ws.cell(r, 1).alignment = LEFT
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        ws.cell(r, 2, val).font = f(11, True)
        ws.cell(r, 2).alignment = LEFT
        ws.cell(r, 2).border = Border(bottom=thin)
        ws.row_dimensions[r].height = 20

    # table header
    ws.cell(R_HDR, 1, "Date")
    ws.cell(R_HDR, 2, "Particulars")
    ws.merge_cells(start_row=R_HDR, start_column=2, end_row=R_HDR, end_column=3)
    ws.cell(R_HDR, 4, "Amount")
    for c in (1, 2, 4):
        ws.cell(R_HDR, c).font = f(10, True)
        ws.cell(R_HDR, c).alignment = CENTER
    ws.row_dimensions[R_HDR].height = 24
    ws.row_dimensions[R_TITLE].height = ROW_H

    ws.merge_cells(start_row=R_TITLE, start_column=2, end_row=R_TITLE, end_column=3)
    ws.cell(R_TITLE, 2, f"DA Bill FMO {mname}-{year}.").font = f(10, True, True)
    ws.cell(R_TITLE, 2).alignment = LEFT

    # one line per claimed day
    for k in range(nrows):
        r = R_FIRST + k
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        if k < len(claimed):
            d, area = claimed[k]
            ws.cell(r, 1, d)
            ws.cell(r, 2, f"Market visit to {area}" if area else "Market visit")
            ws.cell(r, 4, p["rate"])
        ws.cell(r, 1).number_format = "d-m-yyyy"
        ws.cell(r, 1).alignment = CENTER
        ws.cell(r, 2).alignment = LEFT
        ws.cell(r, 4).number_format = ACCOUNTING
        for c in (1, 2, 4):
            ws.cell(r, c).font = f(10)
        ws.row_dimensions[r].height = ROW_H

    ws.merge_cells(start_row=R_TOTAL, start_column=1, end_row=R_TOTAL, end_column=3)
    ws.cell(R_TOTAL, 1, "Total").font = f(11, True)
    ws.cell(R_TOTAL, 1).alignment = RIGHT
    ws.cell(R_TOTAL, 4, f"=SUM(D{R_FIRST}:D{R_LAST})").font = f(11, True)
    ws.cell(R_TOTAL, 4).number_format = ACCOUNTING
    ws.cell(R_TOTAL, 4).alignment = RIGHT
    ws.row_dimensions[R_TOTAL].height = TOTAL_H
    box_range(ws, f"A{R_HDR}:D{R_TOTAL}")
    for r in range(R_TOTAL + 1, R_SIGN):
        ws.row_dimensions[r].height = GAP_H
    ws.row_dimensions[R_SIGN].height = ROW_H

    # signatures
    for col, lab in zip((1, 2, 3, 4), ("Claimed By", "Incharge", "Manager", "Approval")):
        ws.cell(R_SIGN, col).border = Border(top=thin)
        ws.cell(R_SIGN, col, lab).font = f(10, True)
        ws.cell(R_SIGN, col).alignment = Alignment(horizontal="center", vertical="top")

    ws.print_area = f"A1:D{R_SIGN}"
    # Every DA sheet prints at 100% on A4 (same size for all bills).
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = 9
    ws.page_setup.scale = PRINT_SCALE
    ws.sheet_properties.pageSetUpPr.fitToPage = False
    ws.page_margins = PageMargins(left=0.6, right=0.6, top=0.4, bottom=0.4, header=0.2, footer=0.2)
    ws.print_options.horizontalCentered = True
    ws.sheet_view.showGridLines = False
    return ws, f"D{R_TOTAL}", len(claimed)


# --------------------------------------------------------------------------
# Summary sheet
# --------------------------------------------------------------------------
def build_summary(wb, title, unit_code, unit_name, month, year, people, sheet_names, total_cells):
    ws = wb.create_sheet(title)
    for c, w in zip("ABCDEFG", (6.5, 13, 28, 14, 13, 13, 13)):
        ws.column_dimensions[c].width = w
    ws.column_dimensions["H"].width = 3
    ws.column_dimensions["I"].width = 36

    ws.merge_cells("A1:G1")
    ws["A1"] = "Summary of Daily Allowance & Conveyance Bills"
    ws["A1"].font = f(12, True, True)
    ws["A1"].alignment = CENTER
    ws.merge_cells("A2:D2")
    ws["A2"] = f"Unit: {unit_code}-{unit_name}"
    ws["A2"].font = f(11, True)
    ws["A2"].alignment = LEFT
    ws.merge_cells("E2:G2")
    ws["E2"] = f"Month: {calendar.month_name[month]}-{year}"
    ws["E2"].font = f(11, True)
    ws["E2"].alignment = LEFT
    ws.row_dimensions[1].height = ws.row_dimensions[2].height = 25

    for rng in ("A3:A4", "B3:B4", "C3:C4", "D3:D4", "E3:F3", "G3:G4"):
        ws.merge_cells(rng)
    for cell, t in (("A3", "Sr.#"), ("B3", "E. Code"), ("C3", "Claimed By"), ("D3", "Designation"),
                    ("E3", "Expense"), ("G3", "Total\nAmount"), ("E4", "D.A"), ("F4", "T.A")):
        ws[cell] = t
        ws[cell].font = f(10, True)
        ws[cell].alignment = CENTER
    ws.row_dimensions[3].height = 25
    ws.row_dimensions[4].height = 15

    first = 5
    for i, (p, sn, tc) in enumerate(zip(people, sheet_names, total_cells)):
        r = first + i
        ws.cell(r, 1, i + 1)
        ws.cell(r, 2, p["code"]).number_format = "@"
        ws.cell(r, 3, p["name"])
        ws.cell(r, 4, p["designation"])
        ws.cell(r, 5, f"={q(sn)}!{tc}")                 # linked to that person's DA sheet
        ws.cell(r, 6, 0)                                  # T.A  (input)
        ws.cell(r, 7, f"=SUM(E{r}:F{r})")
        for c in range(1, 8):
            ws.cell(r, c).font = f(10)
            ws.cell(r, c).alignment = CENTER if c in (1, 2, 4) else (LEFT if c == 3 else RIGHT)
        ws.cell(r, 6).font = f(10, color="0000FF")
        ws.cell(r, 7).font = f(10, True)
        for c in (5, 6, 7):
            ws.cell(r, c).number_format = ACCOUNTING
        ws.row_dimensions[r].height = 28
    last = first + len(people) - 1

    rt, rg = last + 1, last + 2
    ws.merge_cells(f"A{rt}:D{rt}")
    ws.merge_cells(f"A{rg}:D{rg}")
    ws.merge_cells(f"E{rg}:G{rg}")
    ws[f"A{rt}"] = "Total Amount"
    ws[f"A{rg}"] = "Grand Total Amount"
    for c in "EFG":
        ws[f"{c}{rt}"] = f"=SUM({c}{first}:{c}{last})"
        ws[f"{c}{rt}"].number_format = ACCOUNTING
        ws[f"{c}{rt}"].font = f(11, True)
        ws[f"{c}{rt}"].alignment = RIGHT
    ws[f"E{rg}"] = f"=E{rt}+F{rt}"
    ws[f"E{rg}"].number_format = ACCOUNTING
    ws[f"E{rg}"].font = f(11, True)
    ws[f"E{rg}"].alignment = RIGHT
    for r in (rt, rg):
        ws[f"A{r}"].font = f(11, True)
        ws[f"A{r}"].alignment = CENTER
        ws.row_dimensions[r].height = 36
    box_range(ws, f"A1:G{rg}")
    box_range(ws, f"A3:G{rg}")

    # signature lines
    r = rg + 5
    for lab in ("Claimed By", "Unit Manager", "S & D Manager"):
        ws.merge_cells(f"B{r}:C{r}")
        ws[f"B{r}"] = lab
        ws[f"B{r}"].font = f(11, True, True)
        ws[f"B{r}"].alignment = Alignment(horizontal="center", vertical="top")
        ws[f"B{r}"].border = Border(top=thin)
        ws[f"C{r}"].border = Border(top=thin)
        r += 5

    ws["I3"] = "Blue figures (T.A) are inputs - type the T.A amount per person."
    ws["I4"] = "D.A is linked to each person's DA sheet."
    for c in ("I3", "I4"):
        ws[c].font = f(8, italic=True, color="808080")
        ws[c].alignment = LEFT

    ws.print_area = f"A1:G{r - 3}"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.sheet_view.showGridLines = False
    return ws


# --------------------------------------------------------------------------
@dataclass
class PersonSummary:
    """One row of the on-screen result table (the Excel file itself uses formulas)."""
    name: str
    sheet: str
    days: int
    rate: float
    total: float


@dataclass
class DAResult:
    data: bytes                 # the .xlsx file
    filename: str
    summary_sheet: str
    people: list


def parse_settings(people, month, year, settings=None):
    """Validate per-person settings -> [{'sun','fri','leaves': [(d1, d2), ...]}].
    settings = optional list (same order as the data sheet) of dicts:
               {'sun': bool, 'fri': bool, 'leaves': 'text such as 3, 5-7, 14'}"""
    settings = settings or [{} for _ in people]
    parsed = []
    for p, st in zip(people, settings):
        try:
            lv = parse_leaves(st.get('leaves', ''), month, year)
        except ToolError as e:
            raise ToolError(f"{p['name']}: {e}")
        parsed.append(dict(sun=st.get('sun', DEFAULT_EXCLUDE_SUNDAYS),
                           fri=st.get('fri', DEFAULT_EXCLUDE_FRIDAYS), leaves=lv))
    return parsed


def build_workbook(people, unit_code, unit_name, month, year, settings=None):
    """Build the workbook in memory. Returns (workbook, [PersonSummary], summary_sheet_title)."""
    sheet_names = make_sheet_names(people, month, year)
    parsed = parse_settings(people, month, year, settings)

    rng = random.Random()
    wb = Workbook()
    wb.remove(wb.active)
    total_cells, summaries = [], []
    for p, sn, o in zip(people, sheet_names, parsed):    # same order as the data sheet
        _, tc, n_days = build_da_sheet(wb, sn, p, month, year, None, o, rng)
        total_cells.append(tc)
        summaries.append(PersonSummary(p["name"], sn, n_days, p["rate"], n_days * p["rate"]))

    mtxt = sheet_names[0].rsplit(" DA ", 1)[1][: -4]    # same month text the DA sheets ended up with
    sum_title = clean(f"{unit_code} DA Summary {mtxt}{year}")[:31]
    build_summary(wb, sum_title, unit_code, unit_name, month, year, people, sheet_names, total_cells)
    wb.calculation.fullCalcOnLoad = True
    return wb, summaries, sum_title


def generate_bytes(people, unit_code, unit_name, month, year, settings=None):
    """Web entry point: people (from load_data) -> DAResult holding the .xlsx bytes."""
    wb, summaries, sum_title = build_workbook(people, unit_code, unit_name, month, year, settings)
    buf = io.BytesIO()
    wb.save(buf)
    return DAResult(buf.getvalue(), default_output_name(unit_code, unit_name, month, year), sum_title, summaries)


def generate(path, unit_code, unit_name, month, year, out=None, settings=None):
    """Original file-based entry point: builds the workbook, saves it, returns its path."""
    if not os.path.isfile(path):
        raise ToolError(f"File not found:\n{path}")
    people = load_data(path)
    wb, _, _ = build_workbook(people, unit_code, unit_name, month, year, settings)
    out = out or default_output_name(unit_code, unit_name, month, year)
    wb.save(out)
    return os.path.abspath(out)


def default_output_name(unit_code, unit_name, month, year):
    """'[unit code] [unit name] DA [month] [year].xlsx', e.g. '1234 Muridke DA September 2026.xlsx'."""
    raw = f"{unit_code} {unit_name} DA {calendar.month_name[month]} {year}"
    safe = re.sub(r'[\[\]:*?/\\"<>|]', "", raw)                 # characters not allowed in file names
    return re.sub(r"\s+", " ", safe).strip() + ".xlsx"




TEMPLATE_CSV = (
    "Code,Name,Designation,Department,Area 1,Area 2,Area 3,DA\n"
    "1001,Example Person,Sales Officer,Sales,Area A,Area B,Area C,500\n"
)


# --------------------------------------------------------------------------
# Command line (optional) - only used when arguments are given
# --------------------------------------------------------------------------
def run_cli(a):
    path = a.data_file.strip('"')
    month = parse_month(a.month)
    if not (str(a.year).isdigit() and 1900 <= int(a.year) <= 2200):
        raise ToolError("Please give a 4-digit --year.")
    out = generate(path, a.unit_code, a.unit_name, month, int(a.year), a.output)
    print(f"Done. DA sheets + summary written to:\n  {out}")


def main():
    ap = argparse.ArgumentParser(description="Build DA sheets + DA summary (command-line mode; the web app is the main UI).")
    ap.add_argument("data_file", nargs="?", help="DA data sheet (.xlsx/.xls/.ods/.csv)")
    ap.add_argument("--unit-code")
    ap.add_argument("--unit-name")
    ap.add_argument("--month", help="name (September / Sep) or number (9)")
    ap.add_argument("--year")
    ap.add_argument("--output", help="output .xlsx path")
    a = ap.parse_args()

    if a.data_file is None:
        ap.print_help()                          # the window now lives in the web app
        return
    missing = [n for n in ("unit_code", "unit_name", "month", "year") if not getattr(a, n)]
    if missing:
        sys.exit("Command-line mode needs: --unit-code --unit-name --month --year\n"
                 "(or run without arguments to use the window).")
    try:
        run_cli(a)
    except (ToolError, ValueError) as e:
        sys.exit(f"ERROR: {e}")


if __name__ == "__main__":
    main()
