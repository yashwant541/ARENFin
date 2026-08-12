"""
translator.py — turn one Arabic cell into its English equivalent.

Matching cascade (first hit wins), each carrying a confidence and a method tag
so the converter can flag anything uncertain for review:

  1. exact        normalized text == a dictionary key
  2. no_article   same, after stripping the definite article 'ال'
  3. fuzzy        closest key above a similarity threshold (dependency-free)
  4. passthrough  numbers, dates, codes, and anything already Latin -> kept as-is
  5. untranslated Arabic with no match -> returned unchanged, flagged low-conf

The fuzzy step uses a token Dice coefficient plus difflib as a tie-breaker, so
no rapidfuzz/Levenshtein install is required.
"""
from __future__ import annotations
import json
import re
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path

from .arabic import normalize, normalize_no_al, is_arabic, to_ascii_digits

# A cell that is essentially a number / code / date -> never "translate".
_NUMERIC = re.compile(r"^[\s()%+\-.,/:0-9]+$")
_PLAIN_INT = re.compile(r"^-?\d{1,3}(?:,\d{3})*$|^-?\d+$")
_PLAIN_FLOAT = re.compile(r"^-?\d{1,3}(?:,\d{3})*\.\d+$|^-?\d+\.\d+$")
_PARENS_NUM = re.compile(r"^\((\d[\d,]*(?:\.\d+)?)\)$")   # (1,234) -> -1234


def _coerce_number(text: str):
    """Turn a numeric string into int/float; leave %, ratios, codes as text."""
    t = text.strip()
    if _PLAIN_INT.match(t):
        return int(t.replace(",", ""))
    if _PLAIN_FLOAT.match(t):
        return float(t.replace(",", ""))
    m = _PARENS_NUM.match(t)
    if m:
        v = m.group(1).replace(",", "")
        return -float(v) if "." in v else -int(v)
    return text  # percentages, dashes, ranges, ids -> keep as displayed


@dataclass
class Match:
    source: str            # original cell text
    english: str           # best English rendering (may equal source if kept)
    method: str            # exact | no_article | fuzzy | passthrough | untranslated
    confidence: float      # 0.0 - 1.0
    category: str = ""     # dictionary category, when known

    @property
    def needs_review(self) -> bool:
        return self.confidence < 0.80 and self.method in ("fuzzy", "untranslated")


def _tokens(norm: str) -> set[str]:
    return set(t for t in norm.split(" ") if t)


def _dice(a: set[str], b: set[str]) -> float:
    if not a or not b:
        return 0.0
    return 2 * len(a & b) / (len(a) + len(b))


class FinancialTranslator:
    """Loads financial_dictionary.json and resolves Arabic labels."""

    def __init__(self, dictionary_path: str | Path, fuzzy_threshold: float = 0.72):
        from .dictionary_io import load_entries
        entries, self.meta = load_entries(dictionary_path)  # accepts .csv or .json

        self.fuzzy_threshold = fuzzy_threshold
        self._exact: dict[str, tuple[str, str]] = {}      # norm -> (english, cat)
        self._no_al: dict[str, tuple[str, str]] = {}      # norm_no_al -> (english, cat)
        self._index: list[tuple[set[str], str, str, str]] = []  # (tokens, norm, eng, cat)

        for e in entries:
            eng, cat = e["english"], e.get("category", "")
            for ar in e["arabic"]:
                n = normalize(ar)
                if n and n not in self._exact:
                    self._exact[n] = (eng, cat)
                na = normalize_no_al(ar)
                if na and na not in self._no_al:
                    self._no_al[na] = (eng, cat)
                self._index.append((_tokens(n), n, eng, cat))

    # -- public ------------------------------------------------------------
    def translate_cell(self, value) -> Match:
        if value is None or (isinstance(value, str) and not value.strip()):
            return Match("", value if value is not None else "", "passthrough", 1.0)

        # Real numbers/dates from the sheet come through as int/float — keep them.
        if isinstance(value, (int, float)):
            return Match(str(value), value, "passthrough", 1.0)

        text = str(value).strip()
        ascii_text = to_ascii_digits(text)   # ٢٠٢٥ -> 2025 before any checks

        # Pure numeric / punctuation / percentage cells: coerce numbers, keep the rest.
        if _NUMERIC.match(ascii_text):
            return Match(text, _coerce_number(ascii_text), "passthrough", 1.0)

        # Non-Arabic, non-alpha shape (codes, symbols): keep with ASCII digits.
        if not is_arabic(ascii_text) and not any(c.isalpha() for c in ascii_text):
            return Match(text, ascii_text, "passthrough", 1.0)

        # Already-Latin label (e.g. an English header mixed in): leave it.
        if not is_arabic(text):
            return Match(text, text, "passthrough", 1.0)

        norm = normalize(text)

        # 1. exact
        if norm in self._exact:
            eng, cat = self._exact[norm]
            return Match(text, eng, "exact", 1.0, cat)

        # 2. article-stripped
        na = normalize_no_al(text)
        if na in self._no_al:
            eng, cat = self._no_al[na]
            return Match(text, eng, "no_article", 0.95, cat)

        # 3. fuzzy — token Dice, difflib as tie-break
        src_tokens = _tokens(norm)
        best = (0.0, None, None)
        for tok, cand_norm, eng, cat in self._index:
            score = _dice(src_tokens, tok)
            if score >= 0.5:  # only spend difflib on plausible candidates
                score = 0.7 * score + 0.3 * SequenceMatcher(None, norm, cand_norm).ratio()
            if score > best[0]:
                best = (score, eng, cat)
        if best[1] and best[0] >= self.fuzzy_threshold:
            return Match(text, best[1], "fuzzy", round(best[0], 3), best[2] or "")

        # 4. give up: return the Arabic unchanged, flagged for review
        return Match(text, text, "untranslated", 0.0)

    def translate(self, value) -> str:
        return self.translate_cell(value).english

    def stats(self) -> dict:
        return {
            "entries": self.meta.get("entry_count"),
            "arabic_variants": len(self._exact),
            "categories": self.meta.get("categories"),
        }
