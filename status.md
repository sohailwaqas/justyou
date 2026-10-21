# handy: Project Status

## 1. Snapshot
- Phase completed: **4** | 2026-10-20 | Tools merged so far: **5**
- Python: 3.11 recommended (tested 3.12.3) | Streamlit 1.66.0 | Run: `streamlit run app.py` | Deployment: ready for Streamlit Community Cloud, not yet deployed

## 2. Tool Registry
| # | Tool | Page file | Logic module | Purpose | Inputs | Outputs | Phase | Parity |
|---|---|---|---|---|---|---|---|---|
| 1 | DA Maker | views/da_maker.py | tools/da_logic.py | DA bills per salesperson + summary sheet | data sheet .xlsx/.xls/.ods/.csv; unit code/name; month/year; per-person Sun/Fri/leaves; common leaves | .xlsx workbook (bytes) + on-screen totals table | 1 | verified (tests) |
| 2 | Toll Calculator | views/toll_calculator.py | tools/toll_core.py (verbatim) + tools/toll_service.py | OCR toll PDF -> filter -> tick sets -> total | PDF; date range; exclude Sundays; leave periods; ticked sets | total Rs.; audit CSV | 1 | verified (tests) |
| 3 | NKA Filler | views/nka_filler.py | tools/nka_filler.py (original logic verbatim + 2 wrappers) | Place Sale-file cases into key-account column(s) of the NKA workbook | Sale .xlsx/.xlsm; NKA .xlsx/.xlsm; month; year; key accounts | filled workbook (bytes) + metrics, placed/unmatched tables | 2 | verified (CLI + AppTest; logic diff = 0 lines) |
| 4 | Merger | views/merger.py | tools/merger.py (original logic + `summary_out` kwarg + `delete_rows_except()`, + merge_bytes(), CLI) | Merge per-distribution Excel files into one workbook | 1+ .xlsx/.xlsm (file name = Dist. Code), order; unit; nature; month; year; optional sheet/Dist. Code header/header row | `<Unit> <Nature> <Month Year>.xlsx` + metrics, summary table, log | 3 | verified (2 logic edits, see section 9) |
| 5 | Claims Generator | views/claims_generator.py (5 tabs) | tools/claims/ package (16 original modules, import lines only changed) + tools/claims/service.py (bytes wrappers) | Monthly claims + summaries: Freight Charges, Wholesale Claim, Discount Summary, Distribution Cost, WS Claims Detail | Sales file .xls/.xlsx/.csv; Wholesale data .xls/.xlsx/.csv; HTH + Card Discount PDFs (text); month; year; unit; region; conditions; freight rates/route/incharge; sampling cases/rate | 5 .xlsx workbooks (bytes, download) + warnings/notes (also on each workbook's Check sheet) | 4 | verified: web output vs original modules cell-by-cell identical (5 workbooks, synthetic data); AppTest of full chain |

## 3. Navigation Map
`st.navigation` dict, built in app.py from utils/registry.py.
- **Overview**: Home (`views/home.py`, default, url `/`, icon :material/home:)
- **Tools**: DA Maker (`/da-maker`, :material/event_available:), Toll Calculator (`/toll-calculator`, :material/route:), NKA Filler (`/nka-filler`, :material/table_chart:), Merger (`/merger`, :material/merge:), Claims Generator (`/claims-generator`, :material/receipt_long:)
- Tagline: "Your everyday office tools, in one place." (`utils/registry.TAGLINE`). Home grid = 2 columns, 5 cards (last row has one).

## 4. File Tree
```
app.py                       MODIFIED (phase 4 + post)  sidebar_footer(phase=4); calls `ui.install_download_note()`; loads assets/premium.css via `inject_extra_css()`; nav is built from registry
views/__init__.py
views/home.py                hero + tool card grid (from registry) + 3 metric cards (unchanged; verified with 5 tools)
views/da_maker.py            DA Maker UI
views/toll_calculator.py     Toll Calculator UI
views/nka_filler.py          NKA Filler UI
views/merger.py              Merger UI
views/claims_generator.py    NEW (phase 4)  Claims Generator UI: shared inputs + 5 tabs (header icon `receipt`; header wrapped in container key `claims-header`, eyebrow 'Premium tool')
tools/__init__.py
tools/da_logic.py, toll_core.py (UNMODIFIED), toll_service.py, nka_filler.py, merger.py
tools/claims/__init__.py     NEW (empty)
tools/claims/service.py      NEW  bytes in / bytes out wrappers (the ONLY module the page imports logic from, plus ws_processing.Condition)
tools/claims/config.py       NEW  original config.py, byte-identical (packs, headings, WS_/DS_/CO_/WCD_ settings; desktop folder paths unused)
tools/claims/readers.py, processing.py, report_writer.py                 NEW  Freight Charges (original app.py's engine)
tools/claims/ws_reader.py, ws_processing.py, ws_writer.py                NEW  Wholesale Claim
tools/claims/ds_reader.py, ds_processing.py, ds_writer.py                NEW  Discount Summary
tools/claims/cost_reader.py, cost_processing.py, cost_writer.py          NEW  Distribution Cost
tools/claims/wcd_reader.py, wcd_processing.py, wcd_writer.py             NEW  WS Claims Detail
   (the 15 modules above config.py differ from the originals ONLY in import lines: `import x` -> `from tools.claims import x`)
utils/__init__.py
utils/ui.py                  MODIFIED (post-phase-4)  + `install_download_note()`, `DOWNLOAD_NOTE`
utils/icons.py               MODIFIED (phase 4)  + `receipt`
utils/registry.py            MODIFIED (phase 4)  + Claims Generator entry (drives nav + Home cards)
utils/files.py               unchanged
assets/style.css, logo.png, icon.png, favicon.png, logo_full.png        unchanged
assets/premium.css           NEW (post-phase-4)  subtle premium accents for Claims Generator ONLY: sidebar item, Home card (`.st-key-hy-card-tool-claims-generator`), page header (`.st-key-claims-header`); loaded by app.py after style.css
run_handy.bat, packages.txt, .gitignore, .streamlit/config.toml, .streamlit/secrets.toml.example   unchanged
requirements.txt             MODIFIED (phase 4)  + pdfplumber==0.11.9
README.md                    MODIFIED (phase 4)  tool table (5 tools), pdfplumber note, package layout note
status.md                    this file
```
Retired (desktop-only, not shipped): DA_Tool.pyw, toll_calculator.py (tkinter), old app.py, Instructions.txt; claims desktop windows app.py / ws_app.py / ds_app.py / cost_app.py / wcd_app.py, INSTALL_REQUIREMENTS.bat, RUN_TOOL.bat, README_HOW_TO_RUN.txt.

## 5. Design System Reference
**Tokens (CSS vars in style.css `:root`)**: `--hy-bg #F4F7FC`, `--hy-navy #0F2E57`, `--hy-orange #FF9F43`, `--hy-white #FFF`, `--hy-muted #486080`, `--hy-border rgba(15,46,87,.12)`, `--hy-border-strong .28`, `--hy-tint .06`, `--hy-ring rgba(255,159,67,.40)`.
Spacing `--s1..--s6` = 8/16/24/32/40/48px. Radius `--r-md 12px`, `--r-lg 16px`. Shadows `--shadow`, `--shadow-hover`. Font: Inter (Google Fonts @import), fallback system-ui.
Rules: orange only as fill/border/icon/underline; text on orange = navy; primary btn = orange fill + navy text; secondary = white + navy outline; focus = orange ring; min button height 44px.

**utils/ui.py** (import as `from utils import ui`)
- `inject_css() -> None` : app.py only
- `page_header(title, description="", icon=None, eyebrow=None)` : `ui.page_header("X","desc",icon="route",eyebrow="Tool")`
- `section_header(title, subtitle="", step=None)` : numbered badge `ui.section_header("Upload",step=1)`
- `card(key) -> container` : `with ui.card("da-upload"):`; key unique per page; CSS hook `.st-key-hy-card-<key>`; keys starting `tool-` get hover lift
- `metric_card(label, value, caption="")`
- `empty_state(title, message, icon="upload")`
- `hero(title, subtitle, eyebrow="", show_logo=True)` : Home only
- `tool_card_body(icon, title, summary, chips: list[str])` : Home only
- `sidebar_footer(version="1.0", phase=2)` : app.py only; **bump phase each phase** (now 4)
- `show_error(message, exc=None)` : friendly st.error; traceback to logger only (`ui.log` is the logger)
- `show_success(message)`
- `install_download_note() -> None` : app.py only; wraps `st.download_button` ONCE per process so EVERY download in every tool (also future ones) gets the subtle note "handy can make mistakes, verify after download" (11.5px, muted navy); labels containing "template" are skipped. Pages need no change: just keep using `st.download_button(...)` (the module-level function, not `container.download_button`).
**utils/icons.py**: `icon_svg(name, size=24) -> str`; names: calendar-check, route, layers, table, merge, receipt (Claims Generator), upload; unknown names fall back to `layers` (add new paths to `_PATHS`).
**utils/files.py**: `MAX_PDF_MB=25`, `MAX_SHEET_MB=10`, `check_size(uploaded, max_mb) -> str|None`, `file_id(data: bytes) -> str`.
**utils/registry.py**: `Tool(slug,title,summary,page_file,icon,nav_icon,accepts,produces)`; `TOOLS`; `TAGLINE`.
**CSS classes pages may use**: `hy-page-header`, `hy-section`, `hy-step`, `hy-metric`, `hy-empty`, `hy-chip`, `hy-tool-*`, `hy-hero`, `hy-sidebar-foot`, `hy-eyebrow`.
Native widgets (st.metric, tabs, expander, inputs, uploader, alerts, buttons, data_editor) are styled globally; no per-page CSS needed.

**tools/claims/service.py** (no streamlit import; files are `(name, bytes)` tuples; every `make_*` returns `Output(name, data, lines)`)
- `Output(name, data, lines)`, `GeneratedFile(name, data, year, month, unit="")`, re-exported `FriendlyError`
- `load_sales(data, name) -> readers.SalesData` (`.distributors[Distributor(code,name,packs,blank_cols,bad_cells)]`, `.pack_columns`, `.sheet_used`, `.notes`)
- `load_wholesale(data, name) -> ws_reader.WsData` (`.distributors[WsDistributor(code,name,unit,sdm,total_cases)]`, `.rows` DataFrame, `.packs`, `.brands`, `.sheet_used`)
- `describe_pdf(name, data) -> (label, {(month, year)})`
- `ws_pack_choices(wsdata)`, `ws_brand_choices(wsdata)`, `brand_label(b)`  (from ws_app.py)
- `make_freight(sales, selections[(Distributor, rate, route, incharge)], year, month, unit, source_name="")`
- `make_wholesale(wsdata, selections[(WsDistributor, unit, sdm)], conds[ws_processing.Condition(pack, brands tuple, day_from, day_to, rate)], year, month, source_name="")`
- `make_discount_summary(pdfs[(name, bytes)], freight|None, wholesale|None, year, month, unit)`
- `make_cost(summary, sales, sampling_cases, sampling_price, year, month, unit)`
- `make_wcd(selected[Distributor], sales: SalesData, ws_book|None, pdfs[], year, month, unit, region, sales_name="")`
Mechanism: bytes -> private temp folder (`tempfile.TemporaryDirectory(ignore_cleanup_errors=True)`, own sub-folder per file so equal names cannot clash, original file name kept) -> ORIGINAL path-based function -> read the workbook back -> folder deleted. "Saved: <temp path>" lines are rewritten to "Saved: <file name>".

## 6. Conventions
- **Page template**: `page_header` -> numbered `section_header` + `ui.card` blocks (upload -> options -> action button -> results -> download). Before input: `ui.empty_state` then `st.stop()` (Claims Generator instead shows a hint inside each tab, because the tabs need different files).
- **Errors**: logic raises own exception (DA: `da_logic.ToolError`; Merger: `MergeError`; Claims: `readers.FriendlyError`) -> `ui.show_error(str(e))`; unexpected `except Exception as e: ui.show_error("friendly text", e)`. Validation problems the user can fix = `st.warning(..., icon=":material/warning:")`.
- **Caching/state**: heavy work done once per upload, keyed by `file_id(data)` in session_state (no cache decorators). Generated bytes stored in session_state so download buttons don't recompute.
- **Uploads**: always `check_size` first; type via `file_uploader(type=...)`; server cap 25 MB.
- **Naming**: view = `views/<snake>.py`, logic = `tools/<snake>.py` or package `tools/<name>/` with one `service.py` entry point (no `streamlit` import), widget keys `<tool>_<name>`, card keys `<tool>-<part>`.
- **Widgets**: use `width="stretch"` (not deprecated `use_container_width`); button `icon=":material/...:"`.
- **Desktop "Continue anyway?" dialogs** -> a tick box that must be ticked before the button works.
- Downloads: use `st.download_button(...)` at module level so the global 'verify after download' note appears automatically.
- Switch pages with `st.switch_page("views/<file>.py")`.

## 7. State and Secrets
- DA (`da_`): `da_file_id`, `da_people`, `da_settings`, `da_editor_ver`, `da_all_sun`, `da_all_fri`, `da_common_leaves`, `da_unit_code`, `da_unit_name`, `da_month`, `da_year`, `da_result`, `da_upload`.
- Toll (`toll_`): `toll_file_id`, `toll_trips`, `toll_unticked`, `toll_reset`, `toll_result`, `toll_no_sun`, `toll_upload`; dynamic `toll_range_<fid>`, `toll_leave_<fid>`, `toll_sets_<hash>`.
- Merger (`merger_`): `merger_up_ver`, `merger_order`, `merger_unit`, `merger_nature`, `merger_month`, `merger_year`, `merger_sheet`, `merger_dist_col`, `merger_hdr_row`, `merger_run`, `merger_clear`, `merger_download`, `merger_result`; dynamic `merger_up_<fid>`, `merger_dn_<fid>`.
- NKA (`nka_`): `nka_sale_upload`, `nka_nka_upload`, `nka_month`, `nka_year`, `nka_info`, `nka_info_key`, `nka_run`, `nka_result`, `nka_download`; dynamic `nka_accounts_<nka_fid>_<tab>`.
- Claims (`claims_`):
  - shared: `claims_month` (3-letter name or None), `claims_year`, `claims_unit`; uploaders `claims_sales_upload`, `claims_wsdata_upload`, `claims_pdf_upload` (multi)
  - parsed caches: `claims_sales_rec`, `claims_wsdata_rec` = (file_id, object, error text); `claims_pdf_info` = {file_id: (label, months)}
  - chain (replaces the desktop `output` folder): `claims_gen_freight`, `claims_gen_wholesale`, `claims_gen_summary` = `GeneratedFile`; used by a later tab only when month/year match (summary also unit, case-insensitive); an uploaded file in the tab wins
  - results: `claims_res_freight|wholesale|summary|cost|wcd` = (signature, `Output`); downloads `claims_dl_<tab>`
  - Freight: `claims_fr_editor_<sales_fid>`, `claims_fr_run`
  - Wholesale: `claims_ws_fid`, `claims_ws_ids` (list of condition ids, never reused), `claims_ws_next`, `claims_ws_add`, `claims_ws_overlap_ok`, `claims_ws_editor_<ws_fid>`, `claims_ws_run`; per condition `claims_ws_pack_/brand_/from_/to_/rate_/rm_<cid>`
  - Summary: `claims_ds_fr_up`, `claims_ds_wh_up`, `claims_ds_missing_ok`, `claims_ds_run`
  - Cost: `claims_co_sm_up`, `claims_co_cases`, `claims_co_price`, `claims_co_run`
  - WS Claims Detail: `claims_wcd_region`, `claims_wcd_ws_up`, `claims_wcd_ws_ok`, `claims_wcd_pdf_ok`, `claims_wcd_editor_<sales_fid>`, `claims_wcd_run`
- Secrets: **none required**. `.streamlit/secrets.toml.example` has commented format only.

## 8. Dependencies
requirements.txt (pinned; `pip check` clean with the full set, tested together on Python 3.12.3):
```
streamlit==1.66.0
pandas==3.0.2
numpy==2.4.4
openpyxl==3.1.5
xlrd==2.0.2
odfpy==1.4.1
pillow==12.1.1
pymupdf==1.28.2
rapidocr-onnxruntime==1.4.4
onnxruntime==1.24.4
pdfplumber==0.11.9
```
packages.txt: `libgl1`, `libglib2.0-0` (OpenCV runtime, pulled in by rapidocr). pdfplumber needs no system packages (pypdfium2 wheel).
Note: pdfplumber 0.11.10 requires Pillow 13 and conflicts with `pillow==12.1.1`; 0.11.9 is the newest that fits. Re-check when Pillow is bumped.

## 9. Decisions and Assumptions
- Earlier phases: logic extracted from desktop code with minimal edits; toll_core.py untouched; NKA logic byte-identical; Merger has 2 logged edits (`summary_out` kwarg, one-pass `delete_rows_except`); orange never used as text; logo sits on a white card; icon set is inline SVG; Material icons only in nav/buttons.
- Claims Generator: ONE page "Claims Generator", five tabs in workflow order (Freight, Wholesale, Discount Summary, Cost, WS Claims Detail); shared Month / Year / Unit and three shared uploads (sales, wholesale data, HTH + Card PDFs) entered once at the top (the desktop windows copied them from the main window). Month starts empty (as the original), Year defaults to the current year.
- The 15 report modules were copied from the desktop tool with ONLY import lines changed (verified by diff); config.py byte-identical. No calculation, rule, message or Excel layout was touched. Original `app.py/processing.py/readers.py` names kept inside the `tools/claims/` package (no clash with handy's own files).
- Bridge = `service.py`: uploaded bytes -> temp folder -> original path-based functions -> bytes. Chosen over rewriting the readers so parity stays provable. Nothing is stored on the server.
- The `output` folder hand-over between reports became session_state (`claims_gen_*`) + optional upload override in the tab, with the same month (and, for the Cost sheet, unit) matching rule as the desktop file-name search.
- IMPORTANT workflow change: the Cost sheet reads the Discount Summary as saved by hand in Excel on the desktop. On the web the session copy has the hand-filled columns EMPTY; the Cost tab says so and offers an upload for the user's saved copy.
- tkinterdnd2 drag-and-drop -> multi-file `st.file_uploader` (so `tkinterdnd2` is NOT in requirements); folder drop is not available. Remove selected / Clear all -> the X on each uploaded file. `os.startfile` open file / folder -> `st.download_button`. "(2)" unique file names not needed (nothing is saved on the server).
- Desktop yes/no dialogs (overlapping conditions, missing Wholesale / Freight / PDF files) -> "Continue anyway" tick boxes. Desktop pop-up with notes -> "Warnings" (expanded when any line contains WARNING) and "All notes" expanders; the same notes remain on each workbook's Check sheet.
- Wholesale conditions: one row per condition (Pack select, Brand multiselect where none or all = "All", From/To day, Rate); packs = data + `WS_EXTRA_PACKS`, brands = data + `WS_BRAND_LABELS` (as v0.7.1). Condition ids are never reused so stale widget state cannot leak into a new data file.
- Freight / WS Claims Detail / Wholesale distributor lists are `st.data_editor` tables (Tick, Code, Distributor, plus editable columns); rows map to distributors by position (sorting in the UI is display-only).
- Premium accent (Claims Generator only, by request; an earlier heavy 'hero / stepper / banner' version was REVERTED): very subtle and minimalist, CSS only, in assets/premium.css. Sidebar item: 2px orange left edge + faint tint + tiny orange dot. Home card: 2px orange top hairline, small outlined 'PREMIUM' tag (muted navy text), orange ring on the icon. Page header: eyebrow 'Premium tool', short orange underline, orange ring on the icon. Orange only as hairline / dot / ring, never text. Other tools untouched.
- Optional enhancements added (labelled): per-file "what was found" table for PDFs, PDF month hint, "inputs changed" notice per tab, per-tab download buttons.

- Post-deploy: sidebar collapse `<<` button forced white + always visible (style.css); subtle fixed 'Trial version' note (`.hy-trial`, emitted by `ui.sidebar_footer`). Collapse-button rules use `data-testid` selectors (version-sensitive).
- Round 3 changes: DA Maker per-person table is now ordinary widgets (keys `da_sun_{i}`, `da_fri_{i}`, `da_leave_{i}`; old data_editor + `da_settings`/`da_editor_ver` removed) because the grid cell editor floated away on scroll; output name = `da.default_output_name(unit_code, unit_name, month, year)` -> '1234 Muridke DA September 2026.xlsx'. Toll step 2 = Month + Year + one 'Leave / excluded dates' text box (keys `toll_month_{fid}`, `toll_year_{fid}`, `toll_leaves_{fid}`; parsed with `da.parse_leaves`). style.css ends with a 'Theme fallback' block forcing palette on checkbox/radio/tags/inputs when config.toml is missing; `.hy-th` = table header label.

## 10. Known Issues / Tech Debt / Optional Enhancements
- Visual styling relies on Streamlit `data-testid` selectors (version-sensitive); CSS could not be eyeballed in a browser during build (phases 1-4): report any visual glitch. New this phase: the Wholesale condition row (6 columns, small number inputs, "✕" button), the 5-tab strip, the 5-card Home grid.
- assets/premium.css: the sidebar rule targets `a[data-testid="stSidebarNavLink"][href$="claims-generator"]` (with a `stSidebarNav` fallback) and the card/header rules use the container-key classes `st-key-...`; the data-testid part is version-sensitive (re-check after upgrading Streamlit). NOT yet seen in a browser: report whether the sidebar dot / edge shows, and whether the PREMIUM tag overlaps the card icon on narrow screens.
- st.dataframe / st.data_editor are canvas-rendered: theme comes from config.toml, header colours cannot be set via CSS.
- Claims Generator tested on SYNTHETIC files only (sales xlsx/csv, wholesale xlsx/csv, text PDFs written to match the parser's expectations). Verify with a real month: especially the HTH / Card Discount PDF layouts and a real `.xls` sales file (xlrd path not exercised here).
- Chain files live only for the browser session; closing the tab loses them (download each report, re-upload in the override boxes later).
- Wholesale data / sales files are limited to `MAX_SHEET_MB` (10 MB) each and PDFs to 25 MB each; raise `utils/files.py` and `maxUploadSize` together if real files are bigger.
- Optional: Select all / none for the tick tables; one "Make everything" button; remember last unit/region; a tests/ folder (tests were run ad hoc: engine parity script + AppTest with a fake uploader).
- OCR memory/time on Community Cloud for large PDFs; consider page-count guard. DA: random areas differ on every Generate. NKA / Merger tested on synthetic files only.

## 11. Next Phase Instructions
Upload: this status.md + the new tool file(s). Re-upload `utils/registry.py` and `views/home.py` only if you changed them; re-upload `utils/ui.py`, `assets/style.css`, `requirements.txt`, `app.py` ONLY if you changed them yourself or this file says they changed (phase 4 changed `app.py`, `utils/icons.py`, `utils/registry.py`, `requirements.txt`, `README.md`; the next phase needs `utils/icons.py` and `utils/registry.py` again if a new icon / entry is added). For each new tool say: purpose, inputs, outputs, any API keys.
Report: any visual issue (Claims Generator tab strip, condition rows, 5-card Home grid); whether the five Claims reports match the desktop output on a real month (cell values, Check sheet notes); whether the HTH / Card Discount PDFs were read (the "What was found" table); and anything missing from the Parity table (section 2 plus the Phase 4 reply).
