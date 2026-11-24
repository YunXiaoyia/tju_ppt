#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tju_theme.py — single source of truth for the TJU deck: palette, fonts, sizes, page grid, text-width budget.

Every other module imports constants from here; never hard-code a color, font, or grid coordinate elsewhere.
All geometry is in inches on the 13.333 × 7.50 in (16:9) TJU template.
"""

import math
import re

from pptx.dml.color import RGBColor

# --------------------------------------------------------------------------- #
# Palette (the only colors allowed anywhere in the deck and in self-drawn figures)
# --------------------------------------------------------------------------- #
HEX_BLUE, HEX_BLUE_LIGHT, HEX_RED, HEX_BLACK, HEX_WHITE = "00337F", "EAF1FB", "C00000", "000000", "FFFFFF"
ALLOWED_HEX = {HEX_BLUE, HEX_BLUE_LIGHT, HEX_RED, HEX_BLACK, HEX_WHITE}

C_BLUE = RGBColor.from_string(HEX_BLUE)              # primary: headers, core boxes, borders, page numbers
C_BLUE_LIGHT = RGBColor.from_string(HEX_BLUE_LIGHT)  # secondary fill: card bodies, zebra rows, equation boxes
C_RED = RGBColor.from_string(HEX_RED)                # emphasis: constraints, failure boundaries, takeaway line
C_BLACK = RGBColor.from_string(HEX_BLACK)            # body text
C_WHITE = RGBColor.from_string(HEX_WHITE)            # canvas, inverted text

# Tone → (header fill, body fill, border).  "blue" is the default; use "red" for at most one focal item per row.
TONES = {
    "blue": (C_BLUE, C_BLUE_LIGHT, C_BLUE),
    "red": (C_RED, C_WHITE, C_RED),
}

# --------------------------------------------------------------------------- #
# Fonts and sizes
# --------------------------------------------------------------------------- #
FONT_LATIN = "Times New Roman"
FONT_CJK = "微软雅黑"
FONT_SLOTS = (("a:latin", FONT_LATIN), ("a:ea", FONT_CJK), ("a:cs", FONT_LATIN), ("a:sym", FONT_LATIN))

MIN_PT = 18          # every visible run on a slide (raster figure internals excepted)
PT_BANNER = 28       # white text in the dark-blue top banner
PT_SUBTITLE = 20     # blue claim sentence under the banner
PT_HEADING = 19      # column / card / panel headings
PT_BODY = 18
PT_TAKEAWAY = 19     # bottom red takeaway line — fixed; shorten the sentence instead of shrinking it

# --------------------------------------------------------------------------- #
# Page grid
# --------------------------------------------------------------------------- #
SLIDE_W, SLIDE_H = 13.333, 7.50
GRID_L, GRID_R = 0.45, 12.88          # left/right edge of every content element (= subtitle divider rule)
GRID_W = GRID_R - GRID_L              # 12.43
GUTTER = 0.23                         # horizontal gap between side-by-side regions
VGAP = 0.12                           # vertical gap between stacked zones

HEADER_RULE_Y = 1.56                  # blue divider under the subtitle (drawn by the content skeleton)
BODY_TOP = 1.78                       # first usable y when the page has no lead lines
BODY_BOTTOM = 7.00                    # last usable y when the page has no takeaway (page number sits at 7.05)
LEAD_TOP = 1.66                       # first lead line
LEAD_LINE_H = 0.38                    # one 18 pt lead line incl. spacing
TAKEAWAY_TOP, TAKEAWAY_H = 6.42, 0.46
PANEL_HEADER_H = 0.40                 # header bar of a figure panel / card

ASPECT_TOL = 0.05                     # max relative aspect-ratio mismatch between an image and its slot


class LayoutError(ValueError):
    """Raised at build time when content cannot fit its slot under the contract (fix content, not the rule)."""


# --------------------------------------------------------------------------- #
# Text-width budget (bold 微软雅黑 / Times New Roman approximation, conservative)
# --------------------------------------------------------------------------- #
_MARKUP = re.compile(r"[_^]\{([^{}]*)\}")
_WIDE_RANGES = ((0x2E80, 0x9FFF), (0xF900, 0xFAFF), (0xFF00, 0xFFEF), (0x3000, 0x303F),
                (0x2460, 0x24FF), (0x25A0, 0x25FF))


def is_wide(ch):
    o = ord(ch)
    return any(a <= o <= b for a, b in _WIDE_RANGES)


def char_em(ch):
    if is_wide(ch):
        return 1.0
    if ch == " ":
        return 0.25
    if ch.isupper():
        return 0.70
    if ch.isdigit() or ch.islower():
        return 0.52
    return 0.38


def text_width_in(text, pt=PT_BODY):
    """Estimated single-line width (inches) of `text` at `pt`. _{}/^{} markup counts as body-size glyphs."""
    plain = _MARKUP.sub(r"\1", text)
    em = sum(char_em(c) for c in plain)
    # CJK↔Latin auto-spacing adds ~0.25 em at each script boundary
    prev = None
    for c in plain:
        if not c.strip():
            continue
        w = is_wide(c)
        if prev is not None and w != prev:
            em += 0.25
        prev = w
    return em * pt / 72.0


def lines_needed(text, width_in, pt=PT_BODY, pad_in=0.10):
    usable = max(0.05, width_in - pad_in)
    return max(1, math.ceil(text_width_in(text, pt) / usable))


def fits_one_line(text, width_in, pt=PT_BODY, pad_in=0.10):
    return text_width_in(text, pt) <= width_in - pad_in


def cjk_chars_per_line(width_in, pt=PT_BODY, pad_in=0.10):
    """How many CJK characters fit on one line — use it to write copy to length before building."""
    return int((width_in - pad_in) * 72.0 / pt)
