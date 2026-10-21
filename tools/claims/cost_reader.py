"""
cost_reader.py  -  Reads the two files needed for the "Distribution Wise Cost" sheet (v0.6):
   * the Discount Summary workbook made earlier by this tool  (all the claim amounts, columns D to K, per distributor)
   * the sales file  (the Raw Cases of every distributor)
Nothing is changed in the original files.  Every problem is reported as a note; nothing is dropped silently.
"""
import datetime
import os
from dataclasses import dataclass, field

from tools.claims import config
from tools.claims import ds_reader
from tools.claims import readers
from tools.claims.readers import FriendlyError

AMOUNT_COLS = "DEFGHIJK"          # the eight amount columns of the Discount Summary sheet


@dataclass
class SummaryRow:
    code: str
    name: str
    amounts: list = field(default_factory=list)    # 8 values (columns D..K); None = empty cell


@dataclass
class RawEntry:
    name: str
    raw: float = None            # value of the Raw Cases column (None = empty / not a number)
    pack_sum: float = 0.0        # sum of the pack columns (cross-check)
    problem: str = ""            # "" / "blank" / "text:<what was there>"


def _is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


# ----------------------------------------------------------------------------- Discount Summary workbook
def read_summary_book(path):
    """Discount Summary .xlsx -> (rows [SummaryRow], notes [(type, details)], (month, year) or None)."""
    wb, wv = ds_reader._open_book(path, "Discount Summary")
    if config.DS_SHEET_NAME not in wb.sheetnames:
        raise FriendlyError(f"The Discount Summary file has no sheet named '{config.DS_SHEET_NAME}'.\n"
                            "Please choose the file made by the 'Discount Summary...' button.")
    ws, wsv = wb[config.DS_SHEET_NAME], wv[config.DS_SHEET_NAME]
    if str(ws["A1"].value).strip() != config.DS_TITLE or str(ws["B2"].value).strip() != "Code":
        raise FriendlyError("This does not look like a Discount Summary file made by this tool "
                            f"(title '{config.DS_TITLE}' / heading 'Code' not found).")
    rows, notes = [], []
    for r in range(3, ws.max_row + 1):
        if str(ws[f"A{r}"].value).strip() == "Total":
            break
        if readers._is_blank(ws[f"B{r}"].value):
            continue
        code = readers._clean_code(ws[f"B{r}"].value)
        name = " ".join(str(ws[f"C{r}"].value or "").split())
        amounts = []
        for col in AMOUNT_COLS:
            v = ws[f"{col}{r}"].value
            if isinstance(v, str) and v.startswith("="):          # a formula typed by hand: needs Excel's stored result
                cached = wsv[f"{col}{r}"].value
                if _is_num(cached):
                    v = cached
                else:
                    notes.append(("WARNING Formula not calculated",
                                  f"{name} ({code}), cell {col}{r} of the Discount Summary holds a formula with no stored "
                                  "result, so it is treated as EMPTY. Open the file in Excel, save it, and run again."))
                    v = None
            if readers._is_blank(v):
                amounts.append(None)
            elif _is_num(v):
                amounts.append(float(v))
            else:
                notes.append(("WARNING Not a number",
                              f"{name} ({code}), cell {col}{r} of the Discount Summary contains '{v}'; treated as EMPTY. "
                              "Please check."))
                amounts.append(None)
        rows.append(SummaryRow(code, name, amounts))
    if not rows:
        raise FriendlyError("No distributor rows were found in the Discount Summary file.")
    d = ws["N3"].value
    month = (d.month, d.year) if isinstance(d, (datetime.date, datetime.datetime)) else None
    return rows, notes, month


# ----------------------------------------------------------------------------- sales file (Raw Cases)
def read_raw_cases(path):
    """Sales file -> ({code: RawEntry}, notes [(type, details)]).  Raw Cases = the 'Grand Total Raw Cases Vol.' column."""
    if not path or not os.path.isfile(path):
        raise FriendlyError(f"Sales file not found:\n{path}")
    ext = os.path.splitext(path)[1].lower()
    if ext not in (".xls", ".xlsx", ".csv"):
        raise FriendlyError(f"File type '{ext}' is not supported. Use .xls, .xlsx or .csv.")
    raw, sheet = readers._load_raw_table(path, ext)
    if raw.empty:
        raise FriendlyError("The sales file (sheet) is empty.")

    head = None
    for i in range(min(30, len(raw))):
        if config.COL_CODE in [str(x).strip() for x in raw.iloc[i].tolist()]:
            head = i
            break
    if head is None:
        raise FriendlyError(f"Could not find the heading '{config.COL_CODE}' in the first 30 rows of sheet '{sheet}'.\n"
                            "Fix: check the file, or change COL_CODE in config.py.")
    headings = [str(x).strip() for x in raw.iloc[head].tolist()]
    body = raw.iloc[head + 1:].reset_index(drop=True)

    def pos(name):
        return headings.index(name) if name in headings else None

    p_code, p_name, p_addr = pos(config.COL_CODE), pos(config.COL_NAME), pos(config.COL_ADDRESS)
    p_raw, p_end = pos(config.CO_COL_RAW_CASES), pos(config.COL_PACKS_END)
    if p_name is None:
        raise FriendlyError(f"Heading '{config.COL_NAME}' not found. Check COL_NAME in config.py.")
    if p_raw is None:
        raise FriendlyError(f"Heading '{config.CO_COL_RAW_CASES}' (the Raw Cases column) not found in the sales file.\n"
                            "Fix: check the file, or change CO_COL_RAW_CASES in config.py.")
    packs = range((p_addr if p_addr is not None else p_name) + 1, p_end) if p_end is not None else []

    out, notes = {}, []
    for _, row in body.iterrows():
        code_raw, name_raw = row.iloc[p_code], row.iloc[p_name]
        if readers._is_blank(code_raw) or readers._is_blank(name_raw):
            continue                                              # subtotal rows
        code = readers._clean_code(code_raw)
        entry = RawEntry(" ".join(str(name_raw).split()))
        v = row.iloc[p_raw]
        if readers._is_blank(v):
            entry.problem = "blank"
        else:
            try:
                entry.raw = float(v)
            except (TypeError, ValueError):
                entry.problem = f"text:{v}"
        for p in packs:
            pv = row.iloc[p]
            if not readers._is_blank(pv):
                try:
                    entry.pack_sum += float(pv)
                except (TypeError, ValueError):
                    pass
        if code in out:
            notes.append(("WARNING Code repeated in sales file",
                          f"Code {code} appears more than once in the sales file ('{out[code].name}' and '{entry.name}'). "
                          "The FIRST row is used. Please check."))
            continue
        out[code] = entry
    if not out:
        raise FriendlyError("No distributor rows were found in the sales file.")
    return out, notes
