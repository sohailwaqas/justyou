"""
wcd_reader.py  -  Reads the two sources of the "Monthly Wholesale Claims Detail" report (v0.7):
   * the Wholesale Claim workbook made earlier by this tool  -> raw cases per pack for every distributor sheet
   * the HTH Claims and Card Discount Claims PDFs            -> cases per package for every distributor page
Primary Lifting comes from the sales file (readers.py is used as it is, nothing is changed there).
Nothing is changed in the original files.  Every problem is reported as a note; nothing is dropped silently.
"""
import os
import re
from dataclasses import dataclass, field

from tools.claims import config
from tools.claims import ds_reader
from tools.claims import readers
from tools.claims.readers import FriendlyError

_NUM = ds_reader._NUM
_LINE = re.compile(rf"^(.+?)\s+({_NUM})\s+({_NUM})$")          # "<package> <cases> <amount>"
_TOTAL = re.compile(rf"^Total:\s+({_NUM})\s+({_NUM})$")


def pack_key(text):
    """Pack / package names are compared without capitals, spaces and brackets: '1000 ML (PET)' = '1000ML PET'."""
    return re.sub(r"[^A-Z0-9]", "", str(text).upper())


def _is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


# ----------------------------------------------------------------------------- Wholesale Claim workbook
@dataclass
class WsSheet:
    code: str
    name: str
    sheet: str
    month: tuple = None                           # (month, year) found on the sheet
    packs: dict = field(default_factory=dict)     # pack_key -> [pack label as written on the sheet, raw cases]


def read_claim_book(path):
    """Wholesale Claim .xlsx -> ({code: WsSheet}, notes [(type, details)]).
    Cases of a pack = sum of the 'Raw Cases' (column H) of every outlet row with that pack (computed here, so it
    does not matter whether Excel has already calculated the file)."""
    wb, _ = ds_reader._open_book(path, "Wholesale Claim")
    out, notes = {}, []
    for ws in wb.worksheets:
        if ws.title.lower() == "check":
            continue
        if str(ws["B2"].value).strip() != config.WS_REPORT_TITLE:
            notes.append(("WARNING Sheet not recognised", f"Wholesale file, sheet '{ws.title}': not a Wholesale Claim "
                                                          "sheet (title not found); skipped."))
            continue
        head = str(ws["C4"].value or "")
        m = ds_reader._CODE.search(head)
        if not m:
            notes.append(("WARNING Sheet not read", f"Wholesale file, sheet '{ws.title}': no '(code) name' in C4; skipped."))
            continue
        sub = None
        for r in range(6, ws.max_row + 1):
            if str(ws[f"B{r}"].value).strip() == "Sub Total":
                sub = r
                break
        if sub is None:
            notes.append(("WARNING Sheet not read", f"Wholesale file, sheet '{ws.title}': 'Sub Total' row not found; skipped."))
            continue
        item = WsSheet(m.group(1), head[m.end():].strip(), ws.title)
        label = None
        for r in range(6, sub):
            v = ws[f"D{r}"].value                    # the pack label is written once per merged block: carry it down
            if not readers._is_blank(v):
                label = " ".join(str(v).split())
            cases = ws[f"H{r}"].value
            if label and _is_num(cases):
                slot = item.packs.setdefault(pack_key(label), [label, 0.0])
                slot[1] += float(cases)
        mm = re.search(r"01-([A-Za-z]{3})-(\d{4})", str(ws["F4"].value or ""))
        if mm and mm.group(1).title() in ds_reader.MON3:
            item.month = (ds_reader.MON3.index(mm.group(1).title()) + 1, int(mm.group(2)))
        if item.code in out:
            notes.append(("WARNING Wholesale sheet repeated",
                          f"Two Wholesale Claim sheets have the code {item.code} ('{out[item.code].sheet}' and "
                          f"'{ws.title}'). The first is used. Please check."))
            continue
        out[item.code] = item
    if not out:
        raise FriendlyError("No Wholesale Claim sheets were found in the Wholesale Claim file.")
    return out, notes


# ----------------------------------------------------------------------------- HTH / Card Discount PDFs
@dataclass
class PackPage:
    kind: str                                      # "HTH" or "CARD"
    code: str
    name: str
    file: str
    page: int
    month: tuple = None
    packs: dict = field(default_factory=dict)      # pack_key -> [package name as printed, cases]
    lines_cases: float = 0.0                       # sum of the package lines (cross-check)
    total_cases: float = None                      # printed "Total:" cases


def read_pdf_packs(paths):
    """Every PDF -> (pages [PackPage], notes [(type, details)], files [(file name, text)]).
    Raises FriendlyError only when no HTH / Card Discount page could be read at all."""
    try:
        import pdfplumber
    except ImportError:
        raise FriendlyError("The PDF reading library is missing.\nFix: run INSTALL_REQUIREMENTS.bat "
                            "(or: pip install pdfplumber).")
    pages, notes, files = [], [], []
    for path in paths:
        fname = os.path.basename(path)
        if not os.path.isfile(path) or not path.lower().endswith(".pdf"):
            notes.append(("WARNING File skipped", f"'{fname}' is not a PDF file, so it was skipped."))
            files.append((fname, "skipped (not a PDF)"))
            continue
        found = 0
        try:
            with pdfplumber.open(path) as pdf:
                for pno, page in enumerate(pdf.pages, 1):
                    text = (page.extract_text() or "").strip()
                    if not text:
                        continue                                      # blank page (the HTH files have one): fine
                    lines = [l.strip() for l in text.split("\n") if l.strip()]
                    title = next((i for i, l in enumerate(lines) if l.startswith("Distributor Credit Note")), None)
                    if title is None:
                        if not re.match(rf"^Grand Total\s+{_NUM}\s+{_NUM}$", text):
                            notes.append(("WARNING Page not recognised",
                                          f"{fname}, page {pno}: this page is not an HTH / Card Discount page "
                                          "and was NOT read. Please check it."))
                        continue
                    if config.DS_PDF_CARD_KEY in lines[title]:
                        kind = "CARD"
                    elif config.DS_PDF_HTH_KEY in lines[title]:
                        kind = "HTH"
                    else:
                        notes.append(("WARNING Page not recognised",
                                      f"{fname}, page {pno}: the title '{lines[title]}' is neither HTH nor Card "
                                      "Discount. The page was NOT read."))
                        continue
                    code, name = ds_reader._page_name(lines, title)
                    head = next((i for i, l in enumerate(lines) if l.startswith("Package")), None)
                    tot = next((i for i, l in enumerate(lines) if l.startswith("Total:")), None)
                    if not code or head is None or tot is None or tot < head:
                        notes.append(("WARNING Page not read",
                                      f"{fname}, page {pno} ({kind}): could not find the distributor code or the "
                                      "package table. The page was NOT read. Please check it."))
                        continue
                    pg = PackPage(kind, code, name, fname, pno)
                    mm = re.search(r"Month:\s*([A-Za-z]{3})-(\d{4})", text)
                    if mm and mm.group(1).title() in ds_reader.MON3:
                        pg.month = (ds_reader.MON3.index(mm.group(1).title()) + 1, int(mm.group(2)))
                    tm = _TOTAL.match(lines[tot])
                    if tm:
                        pg.total_cases = ds_reader._num(tm.group(1))
                    for l in lines[head + 1:tot]:
                        lm = _LINE.match(l)
                        if not lm:
                            notes.append(("WARNING Line not read",
                                          f"{fname}, page {pno} ({kind}, {name}): the line '{l}' could not be read "
                                          "and was NOT counted. Please check it."))
                            continue
                        cases = ds_reader._num(lm.group(2))
                        slot = pg.packs.setdefault(pack_key(lm.group(1)), [" ".join(lm.group(1).split()), 0.0])
                        slot[1] += cases
                        pg.lines_cases += cases
                    pages.append(pg)
                    found += 1
        except Exception as e:
            notes.append(("WARNING File not read", f"'{fname}' could not be opened as a PDF ({e}). It was skipped."))
            files.append((fname, "could not be read"))
            continue
        if found == 0:
            notes.append(("WARNING No data found",
                          f"'{fname}' has no readable HTH / Card Discount pages (a scanned picture instead of "
                          "text, or another kind of report?)."))
        kinds = sorted({p.kind for p in pages if p.file == fname})
        files.append((fname, f"{found} distributor page(s)" + (f" ({'/'.join(kinds)})" if kinds else "")))
    if not pages:
        raise FriendlyError("No HTH or Card Discount data could be read from the PDF file(s).\n"
                            + "\n".join(f"- {t}: {d}" for t, d in notes))
    return pages, notes, files
