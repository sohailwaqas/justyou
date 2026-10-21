"""handy - entry point.  Run:  streamlit run app.py

Responsibilities: page config, global CSS, sidebar branding + navigation.
Tools are declared in utils/registry.py; pages live in views/."""
from pathlib import Path

import streamlit as st
from PIL import Image

from utils import ui
from utils.registry import TOOLS

ASSETS = Path(__file__).parent / "assets"


@st.cache_data(show_spinner=False)
def _read_text(path: str, mtime: float) -> str:
    return Path(path).read_text(encoding="utf-8")


def inject_extra_css(name: str) -> None:
    """Load an extra stylesheet from assets/ (after style.css). Silently skipped if the file is missing."""
    p = ASSETS / name
    if p.exists():
        st.markdown(f"<style>{_read_text(str(p), p.stat().st_mtime)}</style>", unsafe_allow_html=True)


st.set_page_config(page_title="handy", page_icon=Image.open(ASSETS / "favicon.png"), layout="wide",
                   initial_sidebar_state="expanded")
ui.inject_css()
ui.install_download_note()          # subtle 'verify after download' reminder under every download button
inject_extra_css("premium.css")       # subtle accents for Claims Generator (sidebar item, Home card, page header)
st.logo(str(ASSETS / "logo.png"), icon_image=str(ASSETS / "icon.png"), size="large")

pages = {
    "Overview": [st.Page("views/home.py", title="Home", icon=":material/home:", default=True)],
    "Tools": [st.Page(t.page_file, title=t.title, icon=t.nav_icon, url_path=t.slug) for t in TOOLS],
}
navigation = st.navigation(pages)
ui.sidebar_footer(version="1.0", phase=4)
navigation.run()
