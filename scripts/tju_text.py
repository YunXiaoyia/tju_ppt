#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tju_text.py — run-level typography: four font slots, bold, >= 18 pt, inline _{sub}/^{sup}.

Paragraph spec accepted by every text-writing helper in this skill:
    "plain string"                                   → default color / size
    ("text", color, pt)  or  ("text", color, pt, bold)
    [("label：", C_BLUE, 18), ("body text", C_BLACK, 18)]   → multi-color segments in one paragraph
Inside any text, `x_{k}^{ref}` becomes a real OpenXML subscript/superscript (size stays >= 18 pt).
"""

import re

from lxml import etree
from pptx.oxml.ns import qn
from pptx.util import Pt

from tju_theme import C_BLACK, FONT_LATIN, FONT_SLOTS, MIN_PT, PT_BODY

SUB_SUP_PATTERN = re.compile(r"(\_\{[^{}]+\}|\^\{[^{}]+\})")


def ensure_font_slots(rPr):
    for tag, typeface in FONT_SLOTS:
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", typeface)


def set_run_font(run, size_pt=PT_BODY, bold=True, color=C_BLACK, italic=False, baseline=None, min_pt=MIN_PT):
    """Four font slots (CJK 微软雅黑, Latin Times New Roman), bold, size >= min_pt, optional baseline shift."""
    run.font.size = Pt(max(min_pt, size_pt))
    is_bold = True if bold is None else bold
    run.font.bold = is_bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    run.font.name = FONT_LATIN
    rPr = run._r.get_or_add_rPr()
    rPr.set("b", "1" if is_bold else "0")
    if italic:
        rPr.set("i", "1")
    elif "i" in rPr.attrib:
        del rPr.attrib["i"]
    if baseline is not None:
        rPr.set("baseline", str(baseline))
    elif "baseline" in rPr.attrib:
        del rPr.attrib["baseline"]
    ensure_font_slots(rPr)


def append_rich_text(paragraph, text, size_pt=PT_BODY, color=C_BLACK, bold=True, min_pt=MIN_PT):
    """Append `text` to a paragraph, converting _{sub}/^{sup} markup into real baseline-shifted runs."""
    for tok in SUB_SUP_PATTERN.split(text):
        if not tok:
            continue
        r = paragraph.add_run()
        if tok.startswith("_{") and tok.endswith("}"):
            r.text = tok[2:-1]
            set_run_font(r, size_pt=size_pt, bold=bold, color=color, baseline=-25000, min_pt=min_pt)
        elif tok.startswith("^{") and tok.endswith("}"):
            r.text = tok[2:-1]
            set_run_font(r, size_pt=size_pt, bold=bold, color=color, baseline=30000, min_pt=min_pt)
        else:
            r.text = tok
            set_run_font(r, size_pt=size_pt, bold=bold, color=color, min_pt=min_pt)


def normalize_paragraph(item, default_pt=PT_BODY, default_color=C_BLACK):
    """Paragraph spec → list of (text, color, pt, bold) segments."""
    if isinstance(item, str):
        return [(item, default_color, default_pt, True)]
    if isinstance(item, tuple):
        txt, col, pt = item[0], item[1], item[2] if len(item) > 2 else default_pt
        bold = item[3] if len(item) > 3 else True
        return [(txt, col, pt, bold)]
    if isinstance(item, list):
        if len(item) == 1 and isinstance(item[0], list):
            item = item[0]
        segs = []
        for seg in item:
            segs.extend(normalize_paragraph(seg, default_pt, default_color))
        return segs
    return [(str(item), default_color, default_pt, True)]


def paragraph_text(item):
    return "".join(s[0] for s in normalize_paragraph(item))


def write_paragraph(paragraph, item, default_pt=PT_BODY, default_color=C_BLACK, min_pt=MIN_PT):
    for txt, col, pt, bold in normalize_paragraph(item, default_pt, default_color):
        append_rich_text(paragraph, txt, size_pt=max(min_pt, pt), color=col, bold=bold, min_pt=min_pt)


def clear_runs(paragraph):
    for child in list(paragraph._p):
        if not child.tag.endswith("pPr"):
            paragraph._p.remove(child)


def update_paragraph_text(paragraph, new_text, size_pt=PT_BODY, bold=True, color=None, align=None):
    """Replace an existing paragraph's text (keeps paragraph properties)."""
    if align is not None:
        paragraph.alignment = align
    clear_runs(paragraph)
    append_rich_text(paragraph, new_text, size_pt=max(MIN_PT, size_pt), color=color or C_BLACK, bold=bold)


def update_shape_paragraphs(shape, lines, size_pt=PT_BODY, bold=True, color=None):
    tf = shape.text_frame
    for idx, line in enumerate(lines):
        if idx < len(tf.paragraphs):
            update_paragraph_text(tf.paragraphs[idx], line, size_pt=size_pt, bold=bold, color=color)
