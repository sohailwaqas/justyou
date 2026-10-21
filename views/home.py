"""Home / dashboard: hero, tool card grid, quick stats."""
import streamlit as st

from utils import ui
from utils.registry import TAGLINE, TOOLS

ui.hero("Welcome to handy", f"{TAGLINE} Pick a tool below to get started - "
        "each one keeps its full feature set in a clean, consistent workspace.",
        eyebrow="Dashboard")

ui.section_header("Your tools", "Every tool follows the same flow: upload, set options, run, download.")

COLS = 2
for start in range(0, len(TOOLS), COLS):
    cols = st.columns(COLS, gap="medium")
    for col, tool in zip(cols, TOOLS[start:start + COLS]):
        with col:
            with ui.card(f"tool-{tool.slug}"):
                ui.tool_card_body(tool.icon, tool.title, tool.summary, [tool.accepts, tool.produces])
                if st.button("Open tool  →", key=f"home_open_{tool.slug}", type="primary",
                             width="stretch"):
                    st.switch_page(tool.page_file)

ui.section_header("At a glance")
c1, c2, c3 = st.columns(3, gap="medium")
with c1:
    ui.metric_card("Tools available", str(len(TOOLS)), "More are added in every phase")
with c2:
    ui.metric_card("Reads", "PDF · Excel · ODS · CSV", "Scanned PDFs supported via OCR")
with c3:
    ui.metric_card("Exports", "Excel · CSV", "Download straight from each tool")
