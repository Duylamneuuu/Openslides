"""Shared constants for the OpenSlides builder and checker."""

from __future__ import annotations

import re

EMU_PER_INCH = 914400
SLIDE_W_EMU = 12_192_000  # 13.333... in, the 16:9 widescreen width
SLIDE_H_EMU = 6_858_000  # 7.5 in
SLIDE_W_IN = SLIDE_W_EMU / EMU_PER_INCH
SLIDE_H_IN = SLIDE_H_EMU / EMU_PER_INCH

# Precomposed Vietnamese letters (Latin-1 Vietnamese plus the 1EA0–1EF9 block).
VIETNAMESE_RE = re.compile(
    "["
    "\u00c0-\u00c3\u00c8-\u00ca\u00cc\u00cd\u00d2-\u00d5\u00d9\u00da\u00dd"
    "\u00e0-\u00e3\u00e8-\u00ea\u00ec\u00ed\u00f2-\u00f5\u00f9\u00fa\u00fd"
    "\u0102\u0103\u0110\u0111\u0128\u0129\u0168\u0169\u01a0\u01a1\u01af\u01b0"
    "\u1ea0-\u1ef9"
    "]"
)

# Faces that do not carry Vietnamese diacritics. Matched on the family name.
NON_VIETNAMESE_FONTS = (
    "anton",
    "impact",
    "bebas neue",
    "bebas",
    "arial black",
)

URL_RE = re.compile(r"^https?://\S+$", re.IGNORECASE)


def has_vietnamese(text: str) -> bool:
    return bool(VIETNAMESE_RE.search(text or ""))


def non_vietnamese_font(name: str | None) -> bool:
    if not name:
        return False
    lowered = name.strip().lower()
    return any(lowered == item or lowered.startswith(item + " ") for item in NON_VIETNAMESE_FONTS)
