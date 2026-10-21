"""Claims Generator - service layer (no Streamlit import).

The five report engines (freight / wholesale / discount summary / cost / WS claims detail) are the original
desktop modules and still work with file paths. This module is the only bridge to the web app:
uploaded bytes are written into a private temporary folder, the ORIGINAL functions are called on those paths,
and the finished workbook is returned as bytes. The temporary folder is deleted before returning, so nothing is
stored on the server.

Public functions (all return an `Output` unless stated otherwise):
  load_sales(data, name)                      -> readers.SalesData
  load_wholesale(data, name)                  -> ws_reader.WsData
  describe_pdf(name, data)                    -> (label, {(month, year), ...})
  make_freight(sales, selections, year, month, unit, source_name)
  make_wholesale(wsdata, selections, conds, year, month, source_name)
  make_discount_summary(pdfs, freight, wholesale, year, month, unit)
  make_cost(summary, sales, sampling_cases, sampling_price, year, month, unit)
  make_wcd(selected, sales, ws_book, pdfs, year, month, unit, region, sales_name)
Files are passed as (file name, bytes) tuples; the file NAME matters for PDFs and for the Check sheets.
"""
from __future__ import annotations

import gc
import tempfile
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional

import re

from tools.claims import (config, cost_processing, ds_processing, ds_reader, processing, readers, wcd_processing,
                          ws_processing, ws_reader)
from tools.claims.readers import FriendlyError

__all__ = ["FriendlyError", "Output", "GeneratedFile", "load_sales", "load_wholesale", "describe_pdf",
           "ws_pack_choices", "ws_brand_choices", "brand_label",
           "make_freight", "make_wholesale", "make_discount_summary", "make_cost", "make_wcd"]

File = tuple  # (file name, bytes)


@dataclass
class Output:
    """A finished report: workbook bytes + the summary lines the desktop tool showed in its pop-up."""
    name: str
    data: bytes
    lines: list


@dataclass
class GeneratedFile:
    """A report made earlier in this session that a later report can pick up (replaces the 'output' folder)."""
    name: str
    data: bytes
    year: int
    month: int
    unit: str = ""


# ----------------------------------------------------------------------------- temp folder helpers
@contextmanager
def _tmp():
    """Private temporary folder, always removed (also on Windows, where open handles block deletion)."""
    folder = tempfile.TemporaryDirectory(prefix="handy_claims_", ignore_cleanup_errors=True)
    try:
        yield Path(folder.name)
    finally:
        gc.collect()
        folder.cleanup()


def _stage(root: Path, group: str, files: Iterable[File]) -> list:
    """Write (name, bytes) files into root/group/<n>/<name>; own folder each, so equal names cannot clash."""
    paths = []
    for i, (name, data) in enumerate(files):
        folder = root / group / str(i)
        folder.mkdir(parents=True, exist_ok=True)
        p = folder / (Path(name).name or f"file{i}")
        p.write_bytes(data)
        paths.append(str(p))
    return paths


def _finish(path: str, lines: list) -> Output:
    """Read the saved workbook back and hide the temporary path in the summary lines."""
    p = Path(path)
    lines = [f"Saved: {p.name}" if str(l).startswith("Saved: ") else l for l in lines]
    return Output(p.name, p.read_bytes(), lines)


# ----------------------------------------------------------------------------- readers
def load_sales(data: bytes, name: str):
    """Sales file (.xls / .xlsx / .csv) -> readers.SalesData (everything is held in memory)."""
    with _tmp() as t:
        (p,) = _stage(t, "in", [(name, data)])
        return readers.read_sales_file(p)


def load_wholesale(data: bytes, name: str):
    """Wholesale data file (.xls / .xlsx / .csv) -> ws_reader.WsData."""
    with _tmp() as t:
        (p,) = _stage(t, "in", [(name, data)])
        return ws_reader.read_wholesale_file(p)


def describe_pdf(name: str, data: bytes):
    """Quick look at one PDF so the page can say what it holds. Returns (label, set of (month, year))."""
    with _tmp() as t:
        (p,) = _stage(t, "in", [(name, data)])
        try:
            res = ds_reader.read_pdfs([p])
        except FriendlyError:
            return "NO HTH / CARD DISCOUNT DATA FOUND (will be reported)", set()
        return res.files[0][1], {e.month for e in res.entries if e.month}


# ----------------------------------------------------------------------------- Wholesale condition lists (from ws_app.py)
def _key(text) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(text).upper())


def _merge(found, extra) -> list:
    """Names found in the data + names always offered (config). A name in both is shown once, as written in the data."""
    out, keys = list(found), {_key(x) for x in found}
    for e in extra:
        if _key(e) not in keys:
            out.append(e)
            keys.add(_key(e))
    return out


def _pack_order(p: str):
    m = re.match(r"\s*(\d+)", p)
    return (int(m.group(1)) if m else 10 ** 9, p.upper())


def ws_pack_choices(wsdata) -> list:
    """Packs offered in a condition: packs found in the data + config.WS_EXTRA_PACKS, ordered by size."""
    return sorted(_merge(wsdata.packs, config.WS_EXTRA_PACKS), key=_pack_order)


def ws_brand_choices(wsdata) -> list:
    """Brands offered in a condition: brands found in the data + every brand of config.WS_BRAND_LABELS."""
    order = list(config.WS_BRAND_LABELS)
    return sorted(_merge(wsdata.brands, order), key=lambda b: (order.index(b.upper()) if b.upper() in order else len(order), b))


def brand_label(b: str) -> str:
    """How a brand is shown in the list, e.g. 'Pepsi   [PEPSI]'."""
    return config.WS_BRAND_LABELS.get(b.upper(), b.title()) + f"   [{b}]"


# ----------------------------------------------------------------------------- reports
def make_freight(sales, selections, year: int, month: int, unit: str, source_name: str = "") -> Output:
    """selections = [(Distributor, rate, route, area_incharge)]  ->  'Freight Charges <Mon-yy>.xlsx'"""
    with _tmp() as t:
        path, lines = processing.generate_report(selections, sales.pack_columns, year, month, unit,
                                                 str(t), source_name, sales.notes)
        return _finish(path, lines)


def make_wholesale(wsdata, selections, conds, year: int, month: int, source_name: str = "") -> Output:
    """selections = [(WsDistributor, unit, sdm)]; conds = [ws_processing.Condition]  ->  'Wholesale Claim <Mon-yy>.xlsx'"""
    with _tmp() as t:
        path, lines = ws_processing.generate_wholesale_report(selections, wsdata, conds, year, month, str(t), source_name)
        return _finish(path, lines)


def make_discount_summary(pdfs: list, freight: Optional[File], wholesale: Optional[File],
                          year: int, month: int, unit: str) -> Output:
    """pdfs = [(name, bytes)]; freight / wholesale = (name, bytes) or None (that column stays EMPTY)."""
    with _tmp() as t:
        pdf_paths = _stage(t, "pdf", pdfs)
        fr = _stage(t, "freight", [freight])[0] if freight else None
        wh = _stage(t, "wholesale", [wholesale])[0] if wholesale else None
        out = t / "out"
        out.mkdir()
        path, lines = ds_processing.generate_discount_summary(pdf_paths, fr, wh, year, month, unit, str(out))
        return _finish(path, lines)


def make_cost(summary: File, sales: File, sampling_cases: float, sampling_price: float,
              year: int, month: int, unit: str) -> Output:
    """summary = the Discount Summary workbook (name, bytes); sales = the sales file (name, bytes)."""
    with _tmp() as t:
        sm = _stage(t, "summary", [summary])[0]
        sa = _stage(t, "sales", [sales])[0]
        out = t / "out"
        out.mkdir()
        path, lines = cost_processing.generate_cost_sheet(sm, sa, sampling_cases, sampling_price, year, month, unit, str(out))
        return _finish(path, lines)


def make_wcd(selected, sales, ws_book: Optional[File], pdfs: list, year: int, month: int,
             unit: str, region: str, sales_name: str = "") -> Output:
    """selected = [Distributor]; sales = readers.SalesData; ws_book = Wholesale Claim workbook or None; pdfs may be []."""
    with _tmp() as t:
        ws = _stage(t, "wholesale", [ws_book])[0] if ws_book else None
        pdf_paths = _stage(t, "pdf", pdfs)
        out = t / "out"
        out.mkdir()
        path, lines = wcd_processing.generate_wcd_report(selected, sales, ws, pdf_paths, year, month, unit, region,
                                                         str(out), sales_name)
        return _finish(path, lines)
