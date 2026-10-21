"""
ws_reader.py  -  Reads the WHOLESALE DATA file (.xls / .xlsx / .csv) for the Wholesale Claim report.
Nothing is changed in the original file.  (The freight tool uses readers.py; this file does not touch it.)
"""
import os
import re
from dataclasses import dataclass

import pandas as pd

from tools.claims import config
from tools.claims import readers
from tools.claims.readers import FriendlyError

OPTIONAL_COLS = {"promo_cases"}          # these headings may be missing from the file


@dataclass
class WsDistributor:
    code: str
    name: str
    unit: str            # taken from the data (the user can change it in the window)
    sdm: str             # SDM/O/S name taken from the data (the user can change it)
    total_cases: float


@dataclass
class WsData:
    distributors: list   # [WsDistributor]
    rows: object         # pandas table, one line per transaction line (see _build_rows)
    packs: list          # pack names found in the data
    brands: list         # brand names found in the data
    sheet_used: str
    notes: list          # notes about the whole file


def _norm(v):
    """Text with extra spaces removed ('' for empty cells)."""
    return "" if readers._is_blank(v) else " ".join(str(v).split())


def _load_raw(path, ext):
    """Return (raw table, sheet name, index of the heading row)."""
    heading = config.WS_COLS["dist_code"]

    def find_header(raw):
        for i in range(min(30, len(raw))):
            if heading in [str(x).strip() for x in raw.iloc[i].tolist()]:
                return i
        return None

    try:
        if ext == ".csv":
            raw = None
            for enc in ("utf-8-sig", "cp1252"):
                try:
                    raw = pd.read_csv(path, header=None, encoding=enc)
                    break
                except UnicodeDecodeError:
                    continue
            if raw is None:
                raise FriendlyError("Could not read the CSV file (unknown text encoding).")
            idx = find_header(raw)
            if idx is None:
                raise FriendlyError(f"Could not find the heading '{heading}' in the first 30 rows of the CSV file.\n"
                                    "Fix: check the file, or change WS_COLS in config.py.")
            return raw, "CSV file", idx
        book = pd.ExcelFile(path)
        for name in book.sheet_names:                    # first sheet that has the heading row
            raw = book.parse(name, header=None)
            idx = find_header(raw)
            if idx is not None:
                return raw, name, idx
        raise FriendlyError(f"No sheet with the heading '{heading}' was found in the first 30 rows.\n"
                            f"Sheets in the file: {', '.join(book.sheet_names)}\n"
                            "Fix: check you chose the wholesale data file, or change WS_COLS in config.py.")
    except FriendlyError:
        raise
    except ImportError:
        raise FriendlyError("A reading library is missing (needed for .xls files).\n"
                            "Fix: run INSTALL_REQUIREMENTS.bat (or: pip install xlrd).")
    except PermissionError:
        raise FriendlyError("The wholesale data file is locked. Close it in Excel and try again.")
    except Exception as e:
        raise FriendlyError("Could not open the wholesale data file. It may be damaged or not a real "
                            f"Excel/CSV file.\nDetails: {e}")


_ISO_LIKE = re.compile(r"^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}")


def _is_text_or_serial(v):
    """True for cells that pandas' month-first guessing gets wrong: text dates and Excel serial numbers."""
    if isinstance(v, bool) or readers._is_blank(v):
        return False
    return isinstance(v, (str, int, float))


def _parse_text_or_serial(v):
    """PHASE 4 FIX (the only logic change in this module).
    Text dates are read DAY-first (05-10-2026 = 5 Oct 2026, as everywhere else in this tool; the original read it as
    10 May, so the sale fell outside the month and the claim came out empty). Year-first text (2026-10-05) is
    unchanged. Excel serial numbers (e.g. 46300) are converted to the real date instead of 1970."""
    if isinstance(v, (int, float)):
        return pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(v)) if 20000 <= v <= 80000 else pd.NaT
    t = str(v).strip()
    if _ISO_LIKE.match(t):
        return pd.to_datetime(t, errors="coerce")
    return pd.to_datetime(t, errors="coerce", dayfirst=True)


def _most_common(series):
    s = [x for x in series if x]
    return pd.Series(s).value_counts().index[0] if s else ""


def read_wholesale_file(path):
    """Read the wholesale data file and return WsData. Raises FriendlyError on problems."""
    if not path or not os.path.isfile(path):
        raise FriendlyError(f"Wholesale data file not found:\n{path}")
    ext = os.path.splitext(path)[1].lower()
    if ext not in (".xls", ".xlsx", ".csv"):
        raise FriendlyError(f"File type '{ext}' is not supported. Use .xls, .xlsx or .csv.")

    raw, sheet, hidx = _load_raw(path, ext)
    headings = [str(x).strip() for x in raw.iloc[hidx].tolist()]

    wanted = {k: h for k, h in config.WS_COLS.items()}
    missing = [h for k, h in wanted.items() if k not in OPTIONAL_COLS and h not in headings]
    if missing:
        raise FriendlyError("These headings were not found in the wholesale data file:\n  " + "\n  ".join(missing) +
                            "\nFix: check the file, or correct WS_COLS in config.py.")
    dupes = [h for k, h in wanted.items() if headings.count(h) > 1]
    if dupes:
        raise FriendlyError("These headings appear more than once in the file, so the tool cannot tell which "
                            "column to use:\n  " + "\n  ".join(dupes))

    body = raw.iloc[hidx + 1:].copy()
    body.columns = headings
    body["_excel_row"] = body.index + 1                     # row number as seen in Excel
    notes = []

    c = {k: body[h] for k, h in wanted.items() if h in headings}
    has_code = ~c["dist_code"].map(readers._is_blank)       # rows without a distributor code = totals / empty rows
    has_outlet = ~c["outlet_code"].map(readers._is_blank)
    for r in body.loc[has_code & ~has_outlet, "_excel_row"]:
        notes.append(f"Excel row {r} has a distributor but no Outlet_Code; the row was skipped.")
    keep = has_code & has_outlet
    if not keep.any():
        raise FriendlyError("No transaction rows were found in the wholesale data file.")
    body = body[keep]
    c = {k: v[keep] for k, v in c.items()}

    df = pd.DataFrame({
        "excel_row": body["_excel_row"].astype(int),
        "dist_code": c["dist_code"].map(readers._clean_code),
        "dist_name": c["dist_name"].map(_norm),
        "unit": c["unit"].map(_norm),
        "sdm": c["sdm"].map(_norm),
        "outlet_code": c["outlet_code"].map(readers._clean_code),
        "outlet_name": c["outlet_name"].map(_norm),
        "outlet_address": c["outlet_address"].map(_norm),
        "invoice": c["invoice"].map(lambda v: "" if readers._is_blank(v) else readers._clean_code(v)),
        "pack": c["pack"].map(_norm),
        "brand": c["brand"].map(_norm),
    })
    df["pack_key"] = df["pack"].str.upper()
    df["brand_key"] = df["brand"].str.upper()

    # --- cases (blank / text cells are counted as 0 and flagged) ---
    def numbers(series):
        blank = series.map(readers._is_blank)
        num = pd.to_numeric(series, errors="coerce")
        flag = pd.Series("", index=series.index)
        flag[blank] = "blank"
        bad = ~blank & num.isna()
        flag[bad] = "text:" + series[bad].astype(str)
        return num.fillna(0.0).astype(float), flag

    df["cases"], df["cases_flag"] = numbers(c["cases"])
    if "promo_cases" in c:
        df["promo_cases"], _ = numbers(c["promo_cases"])
    else:
        df["promo_cases"] = 0.0

    # --- dates (booking = ORDER_DATE, delivery = DELIVERY_DATE) ---
    def dates(series):
        blank = series.map(readers._is_blank)
        try:
            d = pd.to_datetime(series, errors="coerce", format="mixed")
        except (TypeError, ValueError):
            d = pd.to_datetime(series, errors="coerce")
        fix = series.map(_is_text_or_serial).astype(bool)          # phase 4 fix: day-first text dates, serial numbers
        if fix.any():
            d = d.copy()
            d[fix] = pd.to_datetime(series[fix].map(_parse_text_or_serial), errors="coerce")
        flag = pd.Series("", index=series.index)
        flag[blank] = "blank"
        flag[~blank & d.isna()] = "text"
        return d, flag

    df["order_date"], df["order_date_flag"] = dates(c["order_date"])
    df["delivery_date"], df["delivery_date_flag"] = dates(c["delivery_date"])
    df = df.reset_index(drop=True)

    # --- distributor list ---
    dists = []
    for code, g in df.groupby("dist_code", sort=False):
        name = _most_common(g["dist_name"]) or code
        unit, sdm = _most_common(g["unit"]), _most_common(g["sdm"])
        if g["unit"].nunique() > 1 or g["sdm"].nunique() > 1:
            notes.append(f"Distributor {name} ({code}) has more than one Unit or SDM name in the file; "
                         f"the most common one ('{unit}' / '{sdm}') is pre-filled. Please check.")
        dists.append(WsDistributor(code, name, unit, sdm, float(g["cases"].sum())))
    dists.sort(key=lambda d: d.name.lower())

    order = list(config.WS_BRAND_LABELS)
    brands = sorted({b for b in df["brand"] if b},
                    key=lambda b: (order.index(b.upper()) if b.upper() in order else len(order), b))
    packs = sorted({p for p in df["pack"] if p})
    return WsData(dists, df, packs, brands, sheet, notes)
