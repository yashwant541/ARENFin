# AFX — Arabic → English Financial Spreadsheet Translator

Point it at an Arabic financial workbook (`.xlsx` / `.xlsm` / `.xls` / `.csv`) that
reads **right-to-left**, and it returns a clean, **left-to-right English** workbook
with the same structure — labels first, then note, then period columns in natural
order — numbers preserved exactly.

Built as two stages, as specified:

1. **Advanced translation mechanism** — a normalized Arabic→English financial
   glossary plus a matching engine (exact → article-stripped → fuzzy) that tolerates
   the spelling variance real statements contain.
2. **RTL→LTR table conversion** — detects reading direction, translates every cell,
   mirrors the columns when needed, and writes a formatted English financial workbook
   with an audit log.

---

## Quick start

```bash
# one file
python convert.py samples/balance_sheet_ar.xlsx --out ./english_output

# a whole folder (every .xlsx/.xlsm/.xls/.csv inside)
python convert.py --in ./incoming --out ./english_output

# make the fuzzy matcher stricter (fewer, higher-confidence matches)
python convert.py file.csv --threshold 0.85
```

Each `FILE.ext` becomes `FILE_EN.xlsx` in the output folder. Every workbook gets a
**Translation Log** tab listing any cell that was matched fuzzily or left
untranslated, so a human can audit and extend the glossary.

Regenerate sample inputs any time with `python make_samples.py`.

---

## How it works

### Stage 1 — the translation mechanism

**`afx/arabic.py`** — dependency-free Arabic handling. `normalize()` folds the
spelling variants that break naive matching into one canonical lookup key:

| Variation | Example | Normalized |
|---|---|---|
| Alef / hamza forms | `أ إ آ ٱ` → `ا` | `الاصول` |
| Teh marbuta | `ة` → `ه` | `الضريبه` |
| Diacritics / tatweel | `الضَّريبة` | `الضريبه` |
| Arabic-Indic digits | `٢٠٢٥` | `2025` |

`normalize_no_al()` additionally drops the definite article `ال`, so a header written
with or without it still resolves to one entry.

**`afx/dictionary_source.py`** + **`afx/dictionary_extra.py`** — the human-maintained
glossary, the single source of truth (the base set plus a large supplemental set,
merged at build time). `english_canonical -> [arabic variant, …]` across **20
categories**: income statement, balance-sheet assets / liabilities / equity, cash
flow, comprehensive income, Islamic (sharia-compliant) finance, insurance / takaful,
tax & Zakat, audit & governance, derivatives & treasury, credit risk & capital,
ratios, segments & geography, currencies & units, month/period names, general
accounting, headers, statement titles, common terms. **~660 entries / ~1,500 Arabic
variants**, modelled on the line items real Arab-market statements use — add a line
in either file, rerun the build, and every document benefits.

**`afx/build_dictionary.py`** — compiles the glossary into `financial_dictionary.json`
(the flat form the engine loads) and reports any normalized-key collisions.

**`afx/translator.py`** — the matching cascade, first hit wins, each carrying a
confidence and method so uncertain cells can be flagged:

1. `exact` — normalized text equals a glossary key (confidence 1.0)
2. `no_article` — matches after stripping `ال` (0.95)
3. `fuzzy` — closest key above threshold, via a token Dice coefficient with `difflib`
   as tie-breaker; **no `rapidfuzz`/Levenshtein install required** (0.72–…)
4. `passthrough` — numbers, dates, codes, already-Latin text → kept, digits ASCII-ified
5. `untranslated` — Arabic with no match → returned unchanged, flagged for review

### Stage 2 — the RTL→LTR converter

**`afx/converter.py`** — per sheet:

1. Reads the grid (`.csv` sniffs delimiter + BOM; `.xlsx/.xlsm` via `openpyxl`;
   `.xls` via `pandas`/`xlrd` if present).
2. **Decides whether to mirror columns from the *physical* position of the Arabic
   labels**, not merely the `rightToLeft` view flag — because an Arabic sheet is often
   *stored* logically (label in column A) and only *displayed* RTL. If the labels
   physically sit in the rightmost columns, the grid is in mirrored/visual order and
   gets flipped so the label lands in column A. The view flag is only a tie-breaker.
3. Translates every cell through Stage 1.
4. Writes a clean English `.xlsx`: Arial, frozen header, `#,##0;(#,##0);-` number
   format, LTR sheet view, tab named from the translated statement title. Fuzzy cells
   are shaded amber, untranslated cells orange, and both are itemized in the log.

Numbers are never "translated" — only Arabic-Indic digits inside text are normalized.
Nothing is summarized or dropped: the output has the same logical shape as the input,
just mirrored and in English.

---

## Layout

```
afx/
  arabic.py               Arabic normalization (no deps)
  dictionary_source.py    the glossary — edit this
  build_dictionary.py     glossary -> financial_dictionary.json
  financial_dictionary.json
  translator.py           matching engine (exact / no_article / fuzzy / passthrough)
  converter.py            RTL->LTR Excel/CSV conversion + formatting + audit log
convert.py                CLI: file or folder in, English workbooks out
make_samples.py           generates Arabic RTL test workbooks
samples/                  example Arabic inputs
```

## Editing the glossary

You can edit the glossary as a **CSV in Excel** — no JSON required. The tool loads
either format; just point `--dict` at whichever you keep:

```bash
python convert.py --in ./incoming --dict afx/financial_dictionary.csv --out ./english_output
```

`afx/financial_dictionary.csv` has one row per Arabic variant:

| english | category | arabic |
|---|---|---|
| Total assets | assets | إجمالي الأصول |
| Total assets | assets | مجموع الأصول |
| Net interest income | income_statement | صافي إيرادات الفوائد |

- **Add a term** → add a row (same `english` groups variants together).
- **Remove a term** → delete its row(s).
- `category` is free text for grouping only — leave it blank if you don't care.
- Blank rows and rows starting with `#` are ignored, so you can annotate.
- Save as **CSV UTF-8** so Arabic is preserved.

Round-trip back to JSON if you want it:

```python
from afx.dictionary_io import csv_to_json
csv_to_json("afx/financial_dictionary.csv", "afx/financial_dictionary.json")
```

Alternatively, edit the Python source of truth (`afx/dictionary_source.py`) and run
`python -m afx.build_dictionary`, which regenerates **both** the JSON and the CSV.
Anything the Translation Log marks `untranslated` or low-confidence `fuzzy` is your
worklist for which variants to add.

---

## Extending the glossary (source-of-truth route)

Add to the relevant category in `afx/dictionary_source.py`:

```python
"Net interest income": ["صافي إيرادات الفوائد", "صافي دخل الفوائد", "<new variant>"],
```

then `python -m afx.build_dictionary`. Anything the Translation Log marks
`untranslated` or low-confidence `fuzzy` is your worklist for which variants to add.

## Notes & limits

- **Legacy `.xls`** needs `xlrd` (`pip install xlrd`) or a one-time re-save to `.xlsx`;
  the converter says so plainly rather than failing silently.
- The output is static (no formulas), so no recalculation step is required.
- Percentages and ratio strings (`14.1%`, `52 / 68`) are preserved as text, not
  coerced to numbers, to avoid changing their meaning.
