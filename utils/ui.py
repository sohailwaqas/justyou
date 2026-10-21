"""Reusable UI helpers for handy. Pages import this module as `ui`.
All visual styling lives in assets/style.css; helpers only emit markup/classes."""
from __future__ import annotations

import base64
import functools
import html
import logging
from pathlib import Path
from typing import Optional

import streamlit as st

from utils.icons import icon_svg

ROOT = Path(__file__).resolve().parent.parent
log = logging.getLogger("handy")


def _html(markup: str) -> str:
    """Collapse to one line so Markdown never mistakes indented HTML for code."""
    return "".join(line.strip() for line in markup.strip().splitlines())


@st.cache_data(show_spinner=False)
def _read_css(path: str, mtime: float) -> str:
    return Path(path).read_text(encoding="utf-8")


def inject_css() -> None:
    """Load assets/style.css (call once per run, from app.py)."""
    p = ROOT / "assets" / "style.css"
    st.markdown(f"<style>{_read_css(str(p), p.stat().st_mtime)}</style>", unsafe_allow_html=True)


def page_header(title: str, description: str = "", icon: Optional[str] = None,
                eyebrow: Optional[str] = None) -> None:
    """Title block at the top of every tool page."""
    ic = f'<span class="hy-page-icon">{icon_svg(icon, 28)}</span>' if icon else ""
    eb = f'<div class="hy-eyebrow">{html.escape(eyebrow)}</div>' if eyebrow else ""
    st.markdown(_html(f"""<div class="hy-page-header">{ic}<div>{eb}
        <h1>{html.escape(title)}</h1><p>{html.escape(description)}</p></div></div>"""),
                unsafe_allow_html=True)


def section_header(title: str, subtitle: str = "", step: Optional[int] = None) -> None:
    """Numbered section title (step badge in navy with an orange ring)."""
    badge = f'<span class="hy-step">{step}</span>' if step is not None else ""
    sub = f"<p>{html.escape(subtitle)}</p>" if subtitle else ""
    st.markdown(_html(f"""<div class="hy-section">{badge}<div>
        <h2>{html.escape(title)}</h2>{sub}</div></div>"""), unsafe_allow_html=True)


def card(key: str):
    """White rounded surface. Use as `with ui.card("da-upload"):`.
    `key` must be unique per page; it becomes the CSS hook `st-key-hy-card-<key>`."""
    return st.container(key=f"hy-card-{key}")


def metric_card(label: str, value: str, caption: str = "") -> None:
    """Key-figure card with an orange accent edge."""
    cap = f'<div class="c">{html.escape(caption)}</div>' if caption else ""
    st.markdown(_html(f"""<div class="hy-metric"><div class="l">{html.escape(label)}</div>
        <div class="v">{html.escape(str(value))}</div>{cap}</div>"""), unsafe_allow_html=True)


def empty_state(title: str, message: str, icon: str = "upload") -> None:
    """Friendly placeholder shown before the user has provided input."""
    st.markdown(_html(f"""<div class="hy-empty"><div class="i">{icon_svg(icon, 36)}</div>
        <div class="t">{html.escape(title)}</div><div class="m">{html.escape(message)}</div></div>"""),
                unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def _data_uri(path: str) -> str:
    """Embed a local PNG in markup (cached, so it is read from disk only once)."""
    return "data:image/png;base64," + base64.b64encode(Path(path).read_bytes()).decode()


def hero(title: str, subtitle: str, eyebrow: str = "", show_logo: bool = True) -> None:
    """Large navy banner used on the Home page (full logo on the right, hidden on phones)."""
    eb = f'<div class="hy-eyebrow">{html.escape(eyebrow)}</div>' if eyebrow else ""
    logo = (f'<img class="hy-hero-logo" alt="handy logo" '
            f'src="{_data_uri(str(ROOT / "assets" / "logo_full.png"))}">') if show_logo else ""
    st.markdown(_html(f"""<div class="hy-hero"><div class="hy-hero-text">{eb}
        <h1>{html.escape(title)}</h1><p>{html.escape(subtitle)}</p></div>{logo}</div>"""),
                unsafe_allow_html=True)


def tool_card_body(icon: str, title: str, summary: str, chips: list[str]) -> None:
    """Inner markup of a Home tool card (put inside `ui.card("tool-<slug>")`)."""
    chip_html = "".join(f'<span class="hy-chip">{html.escape(c)}</span>' for c in chips)
    st.markdown(_html(f"""<div class="hy-tool-icon">{icon_svg(icon, 26)}</div>
        <div class="hy-tool-title">{html.escape(title)}</div>
        <p class="hy-tool-desc">{html.escape(summary)}</p><div>{chip_html}</div>"""),
                unsafe_allow_html=True)


def sidebar_footer(version: str = "1.0", phase: int = 1) -> None:
    """Small caption under the sidebar navigation."""
    st.sidebar.markdown(_html(f"""<div class="hy-sidebar-foot"><b>handy</b> v{version}<br>
        Built in phases · Phase {phase}</div>
        <div class="hy-trial">Trial version</div>"""), unsafe_allow_html=True)


def show_error(message: str, exc: Optional[BaseException] = None) -> None:
    """Friendly error for end users. The traceback goes to the server log only."""
    if exc is not None:
        log.exception("%s", message, exc_info=exc)
    st.error(message, icon=":material/error:")


def show_success(message: str) -> None:
    st.success(message, icon=":material/check_circle:")


DOWNLOAD_NOTE = "handy can make mistakes, verify after download"


def install_download_note() -> None:
    """Show a very subtle reminder under every `st.download_button` of every tool (present and future).
    Called once from app.py; safe to call on every rerun. Downloads whose label mentions 'template'
    (a blank layout, not a generated result) are left without the note."""
    if getattr(st.download_button, "_handy_note", False):
        return
    original = st.download_button

    @functools.wraps(original)
    def download_button(label, *args, **kwargs):
        clicked = original(label, *args, **kwargs)
        if "template" not in str(label).lower():
            st.markdown(f'<div style="margin:-4px 0 6px;font-size:11.5px;line-height:1.3;'
                        f'color:var(--hy-muted,#486080)">{html.escape(DOWNLOAD_NOTE)}</div>',
                        unsafe_allow_html=True)
        return clicked

    download_button._handy_note = True
    st.download_button = download_button
