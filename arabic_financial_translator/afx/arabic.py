"""
arabic.py — Arabic text handling with zero external dependencies.

Everything here is pure-Python so the translator runs in any environment
(no pyarabic / camel-tools needed). Two jobs:

  1. normalize()  -> a canonical form used as the dictionary lookup key.
                     Collapses the spelling variants that make Arabic term
                     matching hard (alef/hamza forms, teh marbuta, tatweel,
                     diacritics, Arabic-Indic digits, whitespace).

  2. is_arabic() / rtl helpers -> detect Arabic content and RTL sheets so the
                     converter knows a workbook needs flipping.
"""
from __future__ import annotations
import re
import unicodedata

# --- Unicode ranges -------------------------------------------------------
# Arabic diacritics (tashkeel), superscript alef, Quranic marks.
_TASHKEEL = re.compile(
    "[\u0610-\u061A\u064B-\u065F\u0670\u06D6-\u06DC\u06DF-\u06E8\u06EA-\u06ED]"
)
_TATWEEL = re.compile("\u0640")           # kashida elongation
_ARABIC_LETTER = re.compile(r"[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF]")
_WS = re.compile(r"\s+")

# Arabic-Indic and Eastern Arabic-Indic digits -> ASCII
_DIGIT_MAP = {ord(c): str(i % 10)
              for i, c in enumerate("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹")}

# Letter folding applied only for the *lookup key*, never for display text.
_LETTER_FOLD = {
    "أ": "ا", "إ": "ا", "آ": "ا", "ٱ": "ا", "ٲ": "ا", "ٳ": "ا",  # alef forms
    "ة": "ه",                                                    # teh marbuta
    "ى": "ي", "ئ": "ي",                                          # alef maksura / yeh hamza
    "ؤ": "و",                                                    # waw hamza
    "ﻹ": "لا", "ﻷ": "لا", "ﻵ": "لا", "ﻻ": "لا",                  # lam-alef ligatures
    "ك": "ك", "ي": "ي",
}
_FOLD_TABLE = {ord(k): v for k, v in _LETTER_FOLD.items()}


def strip_diacritics(text: str) -> str:
    """Remove tashkeel and tatweel; keep the base letters intact."""
    text = _TASHKEEL.sub("", text)
    text = _TATWEEL.sub("", text)
    return text


def to_ascii_digits(text: str) -> str:
    """٢٠٢٥ -> 2025, keeping the surrounding text untouched."""
    return text.translate(_DIGIT_MAP)


def is_arabic(text: str) -> bool:
    """True if the string contains at least one Arabic letter."""
    return bool(text) and bool(_ARABIC_LETTER.search(str(text)))


def arabic_ratio(text: str) -> float:
    """Share of letters (ignoring spaces/digits/punct) that are Arabic."""
    letters = [c for c in str(text) if c.isalpha()]
    if not letters:
        return 0.0
    ar = sum(1 for c in letters if _ARABIC_LETTER.match(c))
    return ar / len(letters)


def normalize(text) -> str:
    """
    Canonical lookup key.

    "الأصول غير المتداولة" and "الاصول غير المتداوله" collapse to the same
    key so a single dictionary entry matches both spellings.
    """
    if text is None:
        return ""
    s = unicodedata.normalize("NFKC", str(text))
    s = to_ascii_digits(s)
    s = strip_diacritics(s)
    s = s.translate(_FOLD_TABLE)
    # drop the definite article "ال" prefix variance? No — keep it, but also
    # expose a stripped variant via normalize_no_al() for fallback matching.
    s = _WS.sub(" ", s).strip().lower()
    return s


_AL = re.compile(r"(?:^|(?<=\s))ال")


def normalize_no_al(text) -> str:
    """normalize() but with the definite article 'ال' removed everywhere.

    Used only as a *fallback* key: "المطلوبات" -> "مطلوبات" so a header written
    with or without the article still resolves to one entry.
    """
    return _WS.sub(" ", _AL.sub("", normalize(text))).strip()
