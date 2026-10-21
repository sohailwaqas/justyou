# handy

Your everyday office tools, in one place. A multi-page Streamlit app.

| Tool | What it does |
|---|---|
| **DA Maker** | Builds Daily Allowance bills for sales staff + a final summary sheet in one Excel workbook |
| **Toll Calculator** | Reads a toll-transaction PDF (OCR for scans) and totals the toll for any period |
| **NKA Filler** | Places each product's sale cases into the key-account column(s) of the NKA workbook |
| **Merger** | Merges per-distribution Excel files into one clean workbook |
| **Claims Generator** | Monthly claims and summaries: Freight Charges, Wholesale Claim, Discount Summary, Distribution Cost, Wholesale Claims Detail (5 Excel reports) |

## Run locally (Windows)
**Easiest:** install Python 3.11 from python.org (tick **"Add python.exe to PATH"**), unzip the project, then double-click **`run_handy.bat`**.
The first run creates a private environment and installs everything (a few minutes); later runs start in seconds.
Your browser opens at <http://localhost:8501>. Close the black window to stop the app.

**Manual way** (Command Prompt inside the project folder):
```bash
py -3.11 -m venv .venv
.venv\Scripts\activate            # macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```
Use Python 3.11 or 3.12 (3.13 is not supported by `rapidocr-onnxruntime`; 3.10 or older is too old for the pinned pandas/numpy).
The Claims Generator reads HTH / Card Discount PDFs with `pdfplumber` (text PDFs only; no OCR, no extra system packages).
If the OCR engine fails to start on Windows, install the *Microsoft Visual C++ Redistributable 2015-2022 (x64)* and restart.

## Deploy on Streamlit Community Cloud
**1. Push to GitHub** (from inside the project folder; create an empty repo on github.com first, no README/licence):
```bash
git init
git add .
git commit -m "handy: first deploy"
git branch -M main
git remote add origin https://github.com/<your-username>/handy.git
git push -u origin main
```
Check on github.com that `app.py`, `requirements.txt`, `packages.txt`, `assets/`, `tools/`, `utils/`, `views/` and `.streamlit/config.toml` are in the repo root (not inside an extra sub-folder), and that no `.venv` folder was uploaded.

**Important - the hidden `.streamlit` folder.** Dragging files into the GitHub web page often skips folders that start with a dot. If `.streamlit/config.toml` is missing from the repo, the app loses its theme (red checkboxes, grey inputs, grey tables). Check that it exists on github.com; if not, use **Add file -> Create new file**, type `.streamlit/config.toml` as the name and paste the file contents. Pushing with the git commands above always includes it.

**2. Create the app:** go to share.streamlit.io -> **Create app** -> pick the repo, branch `main`, main file path `app.py`.

**3. Advanced settings** -> Python version **3.12** (3.11 also works; do not use 3.13, which `rapidocr-onnxruntime` does not support).

**4. Secrets:** none needed at present. If a tool ever needs one, paste it into the **Secrets** box using the names in `.streamlit/secrets.toml.example`.

**5. Deploy.** The first build takes several minutes. `packages.txt` installs the system libraries OpenCV needs (`libgl1`, `libglib2.0-0`). Later `git push` commits redeploy automatically.

Note: a private repo works too, but Community Cloud must be given access to it. Uploaded files are processed in memory/temp folders only and are not stored.

### Common errors
| Symptom | Fix |
|---|---|
| `ImportError: libGL.so.1` | `packages.txt` missing/not committed - it must sit next to `requirements.txt` |
| Build fails on `onnxruntime` / `rapidocr` | Set Python to 3.11 in Advanced settings, then reboot the app |
| App restarts on big PDFs | Memory limit hit - split the PDF or reduce pages |
| `.xls` file will not open | `xlrd` must be in requirements (it is) - re-save the file as `.xlsx` if it is corrupt |

## Project layout
See `status.md` (section 4) for the annotated file tree and conventions.
Add a tool: create `tools/<name>.py` (logic), `views/<name>.py` (UI), then one `Tool(...)` entry in `utils/registry.py`.
Multi-module tools live in a package (`tools/claims/`) with one `service.py` as the only entry point used by the page.
