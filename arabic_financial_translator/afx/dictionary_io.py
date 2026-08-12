"""
dictionary_io.py — read/write the glossary as CSV so it can be edited in Excel.

CSV shape (one row per Arabic variant — the spreadsheet-friendliest form):

    english,category,arabic
    Total assets,assets,إجمالي الأصول
    Total assets,assets,مجموع الأصول
    Net interest income,income_statement,صافي إيرادات الفوائد
    ...

To add a term: add a row. To remove one: delete the row. `category` is free text
used only for grouping/QA — leave it blank if you don't care. Blank rows and rows
beginning with '#' are ignored, so you can annotate the file.

Both the CSV and the JSON produce the *same* in-memory entry list, so the engine
and converter work with either.
"""
from __future__ import annotations
import csv
import json
from collections import OrderedDict
from pathlib import Path

CSV_HEADER = ["english", "category", "arabic"]


def load_entries(path: str | Path) -> tuple[list[dict], dict]:
    """Return (entries, meta). Accepts .csv or .json, decided by extension."""
    path = Path(path)
    if path.suffix.lower() == ".csv":
        return _load_csv(path)
    data = json.loads(path.read_text(encoding="utf-8"))
    return data["entries"], data.get("_meta", {})


def _load_csv(path: Path) -> tuple[list[dict], dict]:
    grouped: "OrderedDict[str, dict]" = OrderedDict()
    with open(path, encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        header = next(reader, None)
        cols = [h.strip().lower() for h in (header or [])]
        # tolerate any column order / extra columns
        idx = {name: cols.index(name) for name in CSV_HEADER if name in cols}
        if "english" not in idx or "arabic" not in idx:
            raise ValueError(
                f"{path.name}: CSV must have at least 'english' and 'arabic' columns; "
                f"found {cols}"
            )
        for row in reader:
            if not row or (row[0].strip().startswith("#")):
                continue
            eng = row[idx["english"]].strip() if idx["english"] < len(row) else ""
            ar = row[idx["arabic"]].strip() if idx["arabic"] < len(row) else ""
            cat = row[idx["category"]].strip() if "category" in idx and idx["category"] < len(row) else ""
            if not eng or not ar:
                continue
            key = eng
            if key not in grouped:
                grouped[key] = {"english": eng, "category": cat, "arabic": []}
            if ar not in grouped[key]["arabic"]:
                grouped[key]["arabic"].append(ar)
            if cat and not grouped[key]["category"]:
                grouped[key]["category"] = cat

    entries = list(grouped.values())
    meta = {
        "entry_count": len(entries),
        "arabic_variant_count": sum(len(e["arabic"]) for e in entries),
        "categories": sorted({e["category"] for e in entries if e["category"]}),
        "source": path.name,
    }
    return entries, meta


def entries_to_csv(entries: list[dict], out_path: str | Path) -> Path:
    """Write entries as one-row-per-variant CSV (utf-8-sig so Excel shows Arabic)."""
    out_path = Path(out_path)
    with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(CSV_HEADER)
        for e in entries:
            for ar in e["arabic"]:
                w.writerow([e["english"], e.get("category", ""), ar])
    return out_path


def csv_to_json(csv_path: str | Path, json_path: str | Path) -> dict:
    """Round-trip: rebuild the JSON the engine can also load, from an edited CSV."""
    entries, meta = _load_csv(Path(csv_path))
    meta["note"] = "Arabic->English financial glossary. Keys matched on normalized Arabic."
    payload = {"_meta": meta, "entries": entries}
    Path(json_path).write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return {"entries": len(entries), "variants": meta["arabic_variant_count"]}
