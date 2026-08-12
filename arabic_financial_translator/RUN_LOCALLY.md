# RUN LOCALLY — Arabic → English Financial Converter

Two ways to use it: a **web app** (upload multiple files in a browser, download each
result) or the **command line** (point at a file or folder). Both use the same
engine and the same editable dictionary.

---

## 0. Prerequisites

- **Python 3.9+** (check with `python --version` or `python3 --version`)
- The files in this package (you're reading `RUN_LOCALLY.md` at its root)

Open a terminal **in this folder** (the one that contains `afx/`, `convert.py`,
`webapp/`).

---

## 1. One-time setup (create an isolated environment + install deps)

**macOS / Linux**
```bash
cd path/to/arabic_financial_translator
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

**Windows (PowerShell)**
```powershell
cd path\to\arabic_financial_translator
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

> `openpyxl`, `pandas`, and `flask` install from `requirements.txt`. Legacy `.xls`
> inputs also need `xlrd` (`pip install xlrd`); `.xlsx`/`.xlsm`/`.csv` do not.

---

## 2A. Run the WEB APP (recommended — multi-file upload UI)

```bash
python webapp/dataiku/standalone_app.py
```

You'll see `Serving on http://127.0.0.1:5000`. Open that URL in your browser, then:

1. **Drag & drop** (or browse) one or more Arabic files — `.xlsx`, `.xlsm`, `.xls`,
   `.csv`. Select as many as you want.
2. Click **Convert all**.
3. Each file gets its own row with status and badges (sheets, RTL-flipped, cells to
   review). Click **⤓ Download** on any row, or **⤓ Download all (.zip)**.

Stop the server with **Ctrl-C**.

Try it immediately with the included Arabic samples in `samples/`
(`balance_sheet_ar.xlsx`, `income_statement_ar.xlsx`, `cash_flow_ar.csv`).

*(This is the same app you paste into Dataiku — see `webapp/dataiku/README_WEBAPP.md`
for the DSS deployment steps.)*

---

## 2B. Run the COMMAND LINE (batch a folder, no browser)

**One file:**
```bash
python convert.py samples/balance_sheet_ar.xlsx --out ./english_output
```

**A whole folder** (every `.xlsx/.xlsm/.xls/.csv` inside):
```bash
python convert.py --in ./my_arabic_files --out ./english_output
```

**Point at the editable CSV dictionary explicitly (optional):**
```bash
python convert.py --in ./my_arabic_files --dict afx/financial_dictionary.csv --out ./english_output
```

Each `FILE.ext` becomes `FILE_EN.xlsx` in the output folder, left-to-right and in
English, with a **Translation Log** tab listing anything matched fuzzily or left
untranslated.

---

## 3. Editing the dictionary (add your own terms)

Open **`afx/financial_dictionary.csv`** in Excel (it's UTF-8 with ~1,500 rows,
one Arabic variant per row: `english, category, arabic`).

- Add a term → add a row (same `english` groups variants together).
- Remove a term → delete its row(s).
- Save as **CSV UTF-8** so the Arabic is preserved.

The tool reads the CSV directly, so your edits take effect on the next run. Anything
the Translation Log flagged as `untranslated` is your worklist.

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| `python: command not found` | Use `python3` instead of `python`. |
| `ModuleNotFoundError: flask` | Activate the venv and `pip install -r requirements.txt`. |
| A `.xls` file errors on convert | `pip install xlrd`, or re-save it as `.xlsx`. |
| Port 5000 busy | Edit the last line of `standalone_app.py` (`port=5000` → e.g. `5050`). |
| Arabic shows as `?????` after editing the CSV in Excel | Re-save via **File → Save As → CSV UTF-8**. |
| Web page can't reach backend | Make sure you started `standalone_app.py` and are on `http://127.0.0.1:5000`. |

---

## What's in this package

```
arabic_financial_translator/
├── RUN_LOCALLY.md            ← you are here
├── README.md                 engine overview, architecture, dictionary docs
├── requirements.txt
├── convert.py                command-line entry point (file or folder)
├── make_samples.py           regenerates the Arabic test files in samples/
├── afx/                      the engine (importable package)
│   ├── arabic.py             Arabic normalization (no external deps)
│   ├── dictionary_source.py  base glossary (source of truth)
│   ├── dictionary_extra.py   large supplemental glossary
│   ├── build_dictionary.py   compiles source -> JSON + CSV
│   ├── dictionary_io.py       CSV <-> JSON loading
│   ├── financial_dictionary.csv   ← edit this (Excel-friendly)
│   ├── financial_dictionary.json  ← generated
│   ├── translator.py         matching engine (exact / article / fuzzy)
│   └── converter.py          RTL→LTR Excel/CSV conversion + formatting + log
├── webapp/dataiku/           the web app (local + Dataiku)
│   ├── standalone_app.py     run locally: python webapp/dataiku/standalone_app.py
│   ├── backend.py            Flask backend (also the DSS "Python" tab)
│   ├── body.html style.css app.js   the DSS front-end tabs
│   └── README_WEBAPP.md      Dataiku deployment guide
├── samples/                  Arabic RTL example inputs to try
└── example_english_output/   what the converted results look like
```
