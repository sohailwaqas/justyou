"""
processing.py  -  Turns sales rows into report data (cases x rate) and checks the data.
No Excel formatting happens here.
"""
import calendar
import os
import re
from dataclasses import dataclass, field

from tools.claims import config
from tools.claims import readers
from tools.claims import report_writer

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


@dataclass
class ReportData:
    code: str
    name: str
    rate: float
    sheet_name: str
    unit: str = ""
    route: str = ""
    area_incharge: str = ""
    rows: list = field(default_factory=list)     # [(pack label, cases)]
    issues: list = field(default_factory=list)   # [(issue type, details)]


def date_label(year, month):
    """'Aug-26'"""
    return f"{MONTHS[month - 1]}-{year % 100:02d}"


def _ordinal(n):
    if 10 <= n % 100 <= 13:
        return f"{n}th"
    suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


def description_text(rate, year, month):
    """'Freight Charges @ Rs: 1.5 1st-Aug.-26 To 31st-Aug.-26'"""
    last = calendar.monthrange(year, month)[1]
    mon, yy = MONTHS[month - 1] + ".", f"{year % 100:02d}"
    return f"Freight Charges @ Rs: {rate:g} {_ordinal(1)}-{mon}-{yy} To {_ordinal(last)}-{mon}-{yy}"


def make_sheet_name(dist_name, year, month, used):
    """Build '<Distributor> Freight Charges <Mon-yy>' within Excel's 31-character limit.
    Returns (sheet_name, note or None)."""
    dl = date_label(year, month)
    clean = " ".join(re.sub(r"[\[\]:*?\\]", " ", dist_name.replace("/", "-")).split())   # "/" -> "-" (Excel forbids it)
    note = None
    name = f"{clean} {config.FULL_WORD} {dl}"
    if len(name) > config.MAX_SHEET_NAME:
        name = f"{clean} {config.SHORT_WORD} {dl}"
        note = f"Sheet name too long for Excel (31 characters max): '{config.FULL_WORD}' shortened to '{config.SHORT_WORD}'."
        if len(name) > config.MAX_SHEET_NAME:
            room = config.MAX_SHEET_NAME - len(f" {config.SHORT_WORD} {dl}")
            name = f"{clean[:room].rstrip()} {config.SHORT_WORD} {dl}"
            note = (f"Sheet name too long for Excel (31 characters max): '{config.FULL_WORD}' shortened "
                    f"to '{config.SHORT_WORD}' and the distributor name trimmed.")
    base, n = name, 2
    while name.lower() in used:                      # make sure names are unique
        suffix = f" ({n})"
        name = base[:config.MAX_SHEET_NAME - len(suffix)] + suffix
        n += 1
    used.add(name.lower())
    return name, note


def build_report_data(dist, rate, year, month, used_names, pack_columns, unit="", route="", area_incharge=""):
    """Cases per pack for one distributor, plus a list of things the user should check."""
    sheet_name, name_note = make_sheet_name(dist.name, year, month, used_names)
    rd = ReportData(dist.code, dist.name, rate, sheet_name, unit, route, area_incharge)
    if name_note:
        rd.issues.append(("Sheet name shortened", f"Sheet is named '{sheet_name}'. {name_note}"))

    listed_cols = set()
    for label, sources in config.PACKS:
        cases = sum(dist.packs.get(c, 0.0) for c in sources)
        listed_cols.update(sources)
        rd.rows.append((label, cases))
        if abs(cases - round(cases)) > 1e-9:
            rd.issues.append(("Decimal cases",
                              f"{label}: {cases:g} (shown rounded on the sheet; the exact value is used in Amount)"))

    for col in dist.blank_cols:
        if col in listed_cols:
            rd.issues.append(("Blank cell", f"'{col}' was empty in the sales file and is counted as 0."))
    for col, text in dist.bad_cells:
        if col in listed_cols:
            rd.issues.append(("Not a number", f"'{col}' contains '{text}'; counted as 0. Please check."))

    for col in pack_columns:                          # packs that are not on the report layout
        if col not in listed_cols and dist.packs.get(col, 0) != 0:
            rd.issues.append(("Not on sheet", f"{col}: {dist.packs[col]:g} cases are NOT included "
                                              "(this pack is not in the report layout)."))
    if sum(c for _, c in rd.rows) == 0:
        rd.issues.append(("Zero sales", "No cases in any listed pack; the sheet shows zeros."))
    return rd


def unique_path(folder, filename):
    """Never overwrite an existing file: add (2), (3) ... if needed."""
    base, ext = os.path.splitext(filename)
    path, n = os.path.join(folder, filename), 2
    while os.path.exists(path):
        path = os.path.join(folder, f"{base} ({n}){ext}")
        n += 1
    return path


def generate_report(selections, pack_columns, year, month, unit, output_folder=None, source_name="", load_notes=()):
    """selections = [(Distributor, rate, route, area_incharge), ...]; unit is the same for all.
    Returns (saved file path, list of summary lines)."""
    if not selections:
        raise readers.FriendlyError("No distributor was ticked.")
    output_folder = output_folder or config.OUTPUT_FOLDER
    os.makedirs(output_folder, exist_ok=True)

    used, reports = set(), []
    for dist, rate, route, incharge in selections:
        reports.append(build_report_data(dist, rate, year, month, used, pack_columns, unit, route, incharge))

    path = unique_path(output_folder, f"Freight Charges {date_label(year, month)}.xlsx")
    try:
        report_writer.write_workbook(reports, year, month, path, source_name, load_notes)
    except PermissionError:
        raise readers.FriendlyError(f"Could not save '{path}'. Close it in Excel and try again.")

    lines = [f"Saved: {path}", f"Sheets created: {len(reports)}"]
    lines += [f"[File note] {n}" for n in load_notes]
    for r in reports:
        for kind, text in r.issues:
            lines.append(f"[{kind}] {r.name} ({r.code}): {text}")
    return path, lines
