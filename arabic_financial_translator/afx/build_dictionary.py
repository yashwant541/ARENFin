"""
build_dictionary.py — compile dictionary_source.GLOSSARY into the flat JSON
the FinancialTranslator loads. Run whenever the glossary changes.

    python -m afx.build_dictionary            # writes afx/financial_dictionary.json
"""
from __future__ import annotations
import json
from pathlib import Path

from .dictionary_source import GLOSSARY
from .arabic import normalize

OUT = Path(__file__).with_name("financial_dictionary.json")


def _merged_glossary() -> dict:
    """Union the base GLOSSARY with the supplemental GLOSSARY_EXTRA.

    Combined per category; within a category, variants for the same english
    term are unioned (order preserved, duplicates dropped). Extra categories
    that don't exist in the base are added wholesale.
    """
    try:
        from .dictionary_extra import GLOSSARY_EXTRA
    except ImportError:
        GLOSSARY_EXTRA = {}

    merged: dict[str, dict[str, list]] = {}
    for glossary in (GLOSSARY, GLOSSARY_EXTRA):
        for category, mapping in glossary.items():
            cat = merged.setdefault(category, {})
            for english, variants in mapping.items():
                bucket = cat.setdefault(english, [])
                for v in variants:
                    if v not in bucket:
                        bucket.append(v)
    return merged


def build() -> dict:
    glossary = _merged_glossary()
    entries = []
    seen_keys: dict[str, str] = {}   # normalized arabic -> english (collision check)
    collisions = []

    for category, mapping in glossary.items():
        for english, arabic_variants in mapping.items():
            variants = []
            for ar in arabic_variants:
                key = normalize(ar)
                if not key:
                    continue
                variants.append(ar)
                if key in seen_keys and seen_keys[key] != english:
                    collisions.append((ar, seen_keys[key], english))
                seen_keys.setdefault(key, english)
            entries.append({
                "english": english,
                "category": category,
                "arabic": variants,
            })

    payload = {
        "_meta": {
            "entry_count": len(entries),
            "arabic_variant_count": sum(len(e["arabic"]) for e in entries),
            "categories": sorted(glossary.keys()),
            "note": "Arabic->English financial glossary. Keys matched on normalized Arabic.",
        },
        "entries": entries,
    }
    OUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    # also emit the editable CSV (one row per Arabic variant)
    from .dictionary_io import entries_to_csv
    csv_path = OUT.with_suffix(".csv")
    entries_to_csv(entries, csv_path)

    return {"entries": len(entries),
            "variants": payload["_meta"]["arabic_variant_count"],
            "collisions": collisions,
            "path": str(OUT),
            "csv_path": str(csv_path)}


if __name__ == "__main__":
    result = build()
    print(f"Wrote {result['entries']} entries "
          f"({result['variants']} Arabic variants)")
    print(f"  JSON -> {result['path']}")
    print(f"  CSV  -> {result['csv_path']}")
    if result["collisions"]:
        print(f"\n{len(result['collisions'])} normalized-key collisions "
              "(same Arabic -> different English), first 10:")
        for ar, a, b in result["collisions"][:10]:
            print(f"  {ar!r}: {a!r} vs {b!r}")
