"""Small non-UI helpers for the Toll Calculator page (lifted from the original
Streamlit app.py so the page file contains display code only)."""
from __future__ import annotations

import os
import tempfile
from typing import Callable, Optional

import pandas as pd

from tools import toll_core as core


def read_trips(pdf_bytes: bytes, progress: Optional[Callable[[str], None]] = None):
    """OCR the PDF (slow part) and return the list of trips."""
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as fh:
        fh.write(pdf_bytes)
        path = fh.name
    try:
        return core.canonicalise_stations(core.read_pdf(path, progress=progress))
    finally:
        os.unlink(path)


def to_date(v):
    """Editor cells arrive as Timestamp / date / None / NaT."""
    if v is None or pd.isna(v):
        return None
    return v.date() if hasattr(v, "date") else v
