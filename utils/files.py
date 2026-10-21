"""Safe handling of uploaded files."""
from __future__ import annotations

import hashlib
from typing import Optional

MAX_PDF_MB = 25
MAX_SHEET_MB = 10


def check_size(uploaded, max_mb: int) -> Optional[str]:
    """Return a friendly error message if the upload is too large, else None."""
    size_mb = uploaded.size / (1024 * 1024)
    if size_mb > max_mb:
        return (f"'{uploaded.name}' is {size_mb:.1f} MB. "
                f"The limit for this tool is {max_mb} MB.")
    return None


def file_id(data: bytes) -> str:
    """Stable content hash used to detect a new / changed upload."""
    return hashlib.md5(data).hexdigest()
