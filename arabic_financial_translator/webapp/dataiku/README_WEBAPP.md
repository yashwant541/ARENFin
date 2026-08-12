# Dataiku Web App — Arabic → English Financial Converter

A modern, multi-file web app that wraps the `afx` engine: upload many Arabic
submissions or GL workbooks at once, and get one clean **English left-to-right**
`.xlsx` back per file — download them individually or as a single `.zip`.

Each file is converted **independently**: if one fails (e.g. a legacy `.xls` with
no `xlrd` on the server), its row shows the error and the rest still succeed.

---

## What's in `webapp/dataiku/`

| File | Where it goes in Dataiku |
|---|---|
| `backend.py` | Web app → **Python** (backend) tab |
| `body.html` | Web app → **HTML** tab |
| `style.css` | Web app → **CSS** tab |
| `app.js` | Web app → **JavaScript** tab |
| `standalone_app.py` | **Not used in DSS** — runs the app locally for testing |

The frontend talks to the backend through `getWebAppBackendUrl()`, which DSS
injects automatically — so `app.js` is identical in DSS and locally.

---

## Option A — try it locally first (no Dataiku needed)

From the project root (the folder that contains `afx/`):

```bash
pip install flask openpyxl pandas        # rapidfuzz/xlrd optional
python webapp/dataiku/standalone_app.py
# open http://127.0.0.1:5000
```

Drag several `.xlsx` / `.csv` files in, click **Convert all**, download each
result or the zip. (The engine finds its glossary automatically; override with
`AFX_DICT=/path/to/financial_dictionary.csv` if needed.)

---

## Option B — deploy as a Dataiku standard web app

1. **Make `afx` importable in the project.** Easiest: in DSS go to
   **</> (Code) → Libraries → Python**, and add the `afx/` folder (the whole
   package, including `financial_dictionary.csv`) under `python/`. That puts it on
   the project's Python path so `import afx` works from the web app backend.
   *Alternative:* leave `afx/` on the DSS server's filesystem and set the
   environment variable `AFX_HOME` to the folder that contains it (and optionally
   `AFX_DICT` to the glossary file).

2. **Create the web app.** Project → **</> → Webapps → + New Webapp → Standard**
   (HTML/CSS/JS + Python backend). Give it a name.

3. **Paste the four tabs:**
   - **Python** ← `backend.py`
   - **HTML** ← `body.html`
   - **CSS** ← `style.css`
   - **JavaScript** ← `app.js`

4. **Enable the backend.** In the web app **Settings**, make sure *"the backend is
   enabled"* is on (a Python backend is required), then **Save**.

5. **Start / view.** Click **Start backend**, then the eye/▶ to preview. Upload
   files and convert.

> The app writes each converted workbook to a temp job folder on the DSS server
> and streams it back on download. Nothing is written to a dataset or managed
> folder, so no extra DSS objects are needed. (If you *want* the outputs persisted
> to a managed folder, that's a small addition to `_convert_one` in `backend.py` —
> ask and I'll wire `dataiku.Folder(...).upload_stream(...)` in.)

---

## Requirements on the DSS code environment

The web app backend must run in a code environment that has:

- `flask` (DSS standard web apps already provide it)
- `openpyxl`, `pandas`
- optional: `xlrd` (only to read legacy `.xls`), `rapidfuzz` (not required — the
  matcher falls back to `difflib`)

Add these under the code env's **Packages to install** if they aren't present.

---

## How the UI maps to the engine

- **Drag-drop / browse** → multi-select, de-duplicated by name+size.
- **Convert all** → one `POST /api/convert` with every file; the backend converts
  each and returns per-file results.
- Each row shows: sheet count, an **RTL flipped** badge when columns were mirrored,
  and a **cells to review** badge (fuzzy/untranslated) — plus up to five Arabic
  terms it couldn't match, as a hint for what to add to `financial_dictionary.csv`.
- **Download** (per row) → `GET /api/download/<job>/<fid>`.
- **Download all (.zip)** → `GET /api/download_all/<job>`.
- The header chip shows the loaded glossary size from `GET /api/health`.

---

## Endpoints (for reference / integration)

| Method | Path | Purpose |
|---|---|---|
| GET | `/api/health` | glossary stats + accepted extensions |
| POST | `/api/convert` | multipart `files` (repeatable) → per-file results + `job_id` |
| GET | `/api/download/<job_id>/<fid>` | one converted `.xlsx` |
| GET | `/api/download_all/<job_id>` | all outputs of a job as `.zip` |

### Downloads in Dataiku

The front end downloads files by **`fetch()`-ing them as a blob** and saving
client-side, rather than pointing a link straight at the backend URL. This is
deliberate: in DSS, navigating a browser tab directly to a backend endpoint makes
the DSS proxy return an HTML wrapper, so the browser saves something like
`f0.htm` instead of your workbook. Fetching as a blob returns the real bytes and
lets the app set the correct filename (`<name>_EN.xlsx`, or `converted_<job>.zip`).
If you extend the app, keep downloads on the `fetch → blob` path for the same
reason.
