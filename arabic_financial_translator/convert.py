#!/usr/bin/env python3
"""
convert.py — command-line front door.

    # one file
    python convert.py samples/balance_sheet_ar.xlsx

    # a whole folder (all .xlsx/.xlsm/.xls/.csv inside)
    python convert.py --in ./incoming --out ./english

    # tune fuzzy sensitivity
    python convert.py file.csv --threshold 0.8

Output: for each input FILE.ext -> OUT/FILE_EN.xlsx, LTR, English, with a
"Translation Log" tab flagging anything matched fuzzily or left untranslated.
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

from afx import FinancialTranslator, ExcelConverter

DICT = Path(__file__).parent / "afx" / "financial_dictionary.json"
EXTS = {".xlsx", ".xlsm", ".xls", ".csv", ".xltx"}


def gather(target: Path) -> list[Path]:
    if target.is_file():
        return [target]
    if target.is_dir():
        return sorted(p for p in target.iterdir()
                      if p.suffix.lower() in EXTS and not p.name.startswith("~$"))
    raise SystemExit(f"Path not found: {target}")


def main(argv=None):
    ap = argparse.ArgumentParser(description="Translate Arabic RTL financial spreadsheets to English LTR.")
    ap.add_argument("path", nargs="?", help="input file or folder")
    ap.add_argument("--in", dest="in_", help="input folder (alternative to positional)")
    ap.add_argument("--out", default="./english_output", help="output folder (default ./english_output)")
    ap.add_argument("--dict", default=str(DICT), help="path to financial_dictionary.json")
    ap.add_argument("--threshold", type=float, default=0.72, help="fuzzy match threshold 0-1 (default 0.72)")
    args = ap.parse_args(argv)

    target = Path(args.path or args.in_ or ".")
    if not Path(args.dict).exists():
        raise SystemExit(f"Dictionary not found: {args.dict}. Run: python -m afx.build_dictionary")

    tr = FinancialTranslator(args.dict, fuzzy_threshold=args.threshold)
    conv = ExcelConverter(tr)
    print(f"Loaded glossary: {tr.stats()}")

    files = gather(target)
    if not files:
        raise SystemExit(f"No .xlsx/.xlsm/.xls/.csv files found in {target}")

    out_dir = Path(args.out)
    print(f"Converting {len(files)} file(s) -> {out_dir}/\n")
    total_review = 0
    for f in files:
        try:
            rep = conv.convert_file(f, out_dir)
        except Exception as exc:
            print(f"  ✗ {f.name}: {exc}")
            continue
        review = sum(s.fuzzy + s.untranslated for s in rep.sheets)
        total_review += review
        rtl = sum(1 for s in rep.sheets if s.was_rtl)
        print(f"  ✓ {f.name} -> {Path(rep.output_path).name}  "
              f"[{len(rep.sheets)} sheet(s), {rtl} RTL-flipped, "
              f"{review} cell(s) to review]")

    print(f"\nDone. {total_review} cell(s) flagged in the Translation Log tab(s).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
