"""Single source of truth for every merged tool.
Adding a tool in a later phase = add one Tool(...) entry here + its view file;
navigation (app.py) and the Home card grid pick it up automatically."""
from __future__ import annotations

from dataclasses import dataclass

TAGLINE = "Your everyday office tools, in one place."


@dataclass(frozen=True)
class Tool:
    slug: str            # URL path, e.g. /da-maker
    title: str
    summary: str         # one-line card description
    page_file: str       # view file, relative to app.py
    icon: str            # key in utils/icons.py
    nav_icon: str        # Material icon for the sidebar (":material/...:")
    accepts: str         # chip: inputs
    produces: str        # chip: outputs


TOOLS: tuple[Tool, ...] = (
    Tool(
        slug="da-maker",
        title="DA Maker",
        summary="Create Daily Allowance bills for sales staff and a final summary sheet, "
                "all in one Excel workbook.",
        page_file="views/da_maker.py",
        icon="calendar-check",
        nav_icon=":material/event_available:",
        accepts="XLSX · XLS · ODS · CSV",
        produces="Excel workbook",
    ),
    Tool(
        slug="toll-calculator",
        title="Toll Calculator",
        summary="Read a toll-transaction PDF (scanned files too) and calculate the toll "
                "for any period.",
        page_file="views/toll_calculator.py",
        icon="route",
        nav_icon=":material/route:",
        accepts="PDF (OCR)",
        produces="Total · CSV audit",
    ),
    Tool(
        slug="nka-filler",
        title="NKA Filler",
        summary="Take the NKA sales data and fill the NKA file with each product's cases, "
                "ready to download.",
        page_file="views/nka_filler.py",
        icon="table",
        nav_icon=":material/table_chart:",
        accepts="Sale + NKA XLSX",
        produces="Filled NKA workbook",
    ),
    Tool(
        slug="merger",
        title="Merger",
        summary="Combine per-distribution Excel files into one clean workbook: filters removed, "
                "links broken, only each distribution's own rows kept.",
        page_file="views/merger.py",
        icon="merge",
        nav_icon=":material/merge:",
        accepts="Distribution XLSX files",
        produces="Merged Excel workbook",
    ),
    Tool(
        slug="claims-generator",
        title="Claims Generator",
        summary="Generate the monthly claims along with related summaries: freight, wholesale, "
                "discount summary, distribution cost and wholesale claims detail.",
        page_file="views/claims_generator.py",
        icon="receipt",
        nav_icon=":material/receipt_long:",
        accepts="Sales · Wholesale data · PDFs",
        produces="5 Excel reports",
    ),
)
