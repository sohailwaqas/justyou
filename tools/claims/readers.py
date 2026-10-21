"""
readers.py  -  Reads the sales file (.xls / .xlsx / .csv) and returns the list of distributors.
Nothing is changed in the original file.
"""
import os
from dataclasses import dataclass, field

import pandas as pd

from tools.claims import config


class FriendlyError(Exception):
    """An error whose message is written for the user (shown in a message box)."""


@dataclass
class Distributor:
    code: str
    name: str
    packs: dict = field(default_factory=dict)       # sales-file column heading -> cases
    blank_cols: list = field(default_factory=list)  # columns that were empty
    bad_cells: list = field(default_factory=list)   # (column, text) that were not numbers


@dataclass
class SalesData:
    distributors: list
    pack_columns: list
    sheet_used: str
    notes: list


def _is_blank(v):
    return v is None or (isinstance(v, float) and pd.isna(v)) or (isinstance(v, str) and v.strip() == "")


def _clean_code(v):
    """101518.0 -> '101518'  (codes are kept as text, never changed otherwise)."""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def _load_raw_table(path, ext):
    """Return the sheet as a raw table (no header assumptions) and the sheet name."""
    try:
        if ext == ".csv":
            for enc in ("utf-8-sig", "cp1252"):
                try:
                    return pd.read_csv(path, header=None, encoding=enc), "CSV file"
                except UnicodeDecodeError:
                    continue
            raise FriendlyError("Could not read the CSV file (unknown text encoding).")
        book = pd.ExcelFile(path)
        names = book.sheet_names
        match = [n for n in names if config.SALES_SHEET_KEYWORD.lower() in n.lower()]
        if match:
            sheet = match[0]
        elif len(names) == 1:
            sheet = names[0]
        else:
            raise FriendlyError(
                f"No sheet with '{config.SALES_SHEET_KEYWORD}' in its name was found.\n"
                f"Sheets in the file: {', '.join(names)}\n"
                "Fix: rename the sheet, or change SALES_SHEET_KEYWORD in config.py.")
        return book.parse(sheet, header=None), sheet
    except FriendlyError:
        raise
    except ImportError:
        raise FriendlyError("A reading library is missing (needed for .xls files).\n"
                            "Fix: run INSTALL_REQUIREMENTS.bat (or: pip install xlrd).")
    except PermissionError:
        raise FriendlyError("The sales file is locked. Close it in Excel and try again.")
    except Exception as e:
        raise FriendlyError(f"Could not open the sales file. It may be damaged or not a real "
                            f"Excel/CSV file.\nDetails: {e}")


def read_sales_file(path):
    """Read the sales file and return a SalesData object. Raises FriendlyError on problems."""
    if not path or not os.path.isfile(path):
        raise FriendlyError(f"Sales file not found:\n{path}")
    ext = os.path.splitext(path)[1].lower()
    if ext not in (".xls", ".xlsx", ".csv"):
        raise FriendlyError(f"File type '{ext}' is not supported. Use .xls, .xlsx or .csv.")

    raw, sheet = _load_raw_table(path, ext)
    if raw.empty:
        raise FriendlyError("The sales file (sheet) is empty.")

    # --- find the heading row (the row that contains the 'Distr. SAP Code' heading) ---
    header_idx = None
    for i in range(min(30, len(raw))):
        if config.COL_CODE in [str(x).strip() for x in raw.iloc[i].tolist()]:
            header_idx = i
            break
    if header_idx is None:
        raise FriendlyError(f"Could not find the heading '{config.COL_CODE}' in the first 30 rows "
                            f"of sheet '{sheet}'.\nFix: check the file, or change COL_CODE in config.py.")

    headings = [str(x).strip() for x in raw.iloc[header_idx].tolist()]
    body = raw.iloc[header_idx + 1:].reset_index(drop=True)

    def pos(name):
        return headings.index(name) if name in headings else None

    p_code, p_name = pos(config.COL_CODE), pos(config.COL_NAME)
    if p_name is None:
        raise FriendlyError(f"Heading '{config.COL_NAME}' not found. Check COL_NAME in config.py.")
    p_addr = pos(config.COL_ADDRESS)
    p_end = pos(config.COL_PACKS_END)
    if p_end is None:
        raise FriendlyError(f"Heading '{config.COL_PACKS_END}' not found. The tool uses it to know "
                            "where the pack columns end. Check COL_PACKS_END in config.py.")
    start = (p_addr if p_addr is not None else p_name) + 1
    pack_positions = list(range(start, p_end))
    pack_columns = [headings[p] for p in pack_positions]

    # --- pack columns named in config.PACKS that are NOT in the file are counted as 0 (and reported) ---
    wanted = [c for _, srcs in config.PACKS for c in srcs]
    missing = [c for c in wanted if c not in pack_columns]
    if len(missing) == len(wanted):
        raise FriendlyError("None of the pack columns named in config.py (PACKS) were found in the sales file.\n"
                            "Is this the right file? Otherwise correct the names in PACKS (config.py).")

    # --- read distributor rows (rows with no code/name are subtotal rows and are skipped) ---
    distributors, notes, seen = [], [], {}
    for c in missing:
        notes.append(f"Pack column '{c}' is not in the sales file (removed or renamed?). Its cases are counted as 0 on every sheet.")
    for _, row in body.iterrows():
        code_raw, name_raw = row.iloc[p_code], row.iloc[p_name]
        if _is_blank(code_raw) or _is_blank(name_raw):
            continue
        d = Distributor(code=_clean_code(code_raw), name=str(name_raw).strip())
        for p, col in zip(pack_positions, pack_columns):
            v = row.iloc[p]
            if _is_blank(v):
                d.blank_cols.append(col)
                d.packs[col] = 0.0
                continue
            try:
                d.packs[col] = float(v)
            except (TypeError, ValueError):
                d.bad_cells.append((col, str(v)))
                d.packs[col] = 0.0
        if d.code in seen:
            notes.append(f"Code {d.code} appears more than once in the sales file "
                         f"({seen[d.code]} and {d.name}). Check which row is correct.")
        seen[d.code] = d.name
        distributors.append(d)

    if not distributors:
        raise FriendlyError("No distributor rows were found in the sales file.")
    return SalesData(distributors, pack_columns, sheet, notes)
