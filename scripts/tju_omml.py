#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tju_omml.py — native, editable PowerPoint equations (<a14:m> OMML).

    eq = m_sub(m_r("T"), m_r("move", italic=False)) + m_r(" + ", italic=False) + ...
    add_omml_equation(slide, x, y, w, h, eq, jc="center", fill_color=C_BLUE_LIGHT, line_color=C_BLUE)

Rules baked in: every run is >= 18 pt, bold, four font slots; hats use U+0302 (never ASCII '^').
LibreOffice/StarMath caveats are in references/pptx_pitfalls.md.
"""

import html

from lxml import etree
from pptx.enum.text import MSO_ANCHOR
from pptx.util import Inches, Pt

from tju_theme import FONT_CJK, FONT_LATIN, HEX_BLACK


def m_r(text, italic=True, sz=1800, color=HEX_BLACK):
    """One math run.  italic=False for operators, digits, function names, and descriptive subscripts."""
    mrpr = "" if italic else '<m:rPr><m:sty m:val="p"/></m:rPr>'
    return (
        f"<m:r>{mrpr}"
        f'<a:rPr sz="{max(1800, sz)}" b="1" i="{1 if italic else 0}">'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'<a:latin typeface="{FONT_LATIN}"/><a:ea typeface="{FONT_CJK}"/>'
        f'<a:cs typeface="{FONT_LATIN}"/><a:sym typeface="{FONT_LATIN}"/>'
        f"</a:rPr><m:t>{html.escape(text)}</m:t></m:r>"
    )


def m_sub(base_xml, sub_xml):
    return f"<m:sSub><m:e>{base_xml}</m:e><m:sub>{sub_xml}</m:sub></m:sSub>"


def m_sup(base_xml, sup_xml):
    return f"<m:sSup><m:e>{base_xml}</m:e><m:sup>{sup_xml}</m:sup></m:sSup>"


def m_subsup(base_xml, sub_xml, sup_xml):
    return f"<m:sSubSup><m:e>{base_xml}</m:e><m:sub>{sub_xml}</m:sub><m:sup>{sup_xml}</m:sup></m:sSubSup>"


def m_frac(num_xml, den_xml):
    return f"<m:f><m:num>{num_xml}</m:num><m:den>{den_xml}</m:den></m:f>"


def m_sqrt(inner_xml):
    return f'<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>{inner_xml}</m:e></m:rad>'


def m_acc(inner_xml, ch="\u0302"):
    """Accent (default combining circumflex U+0302 — ASCII '^' sinks onto capitals in PowerPoint)."""
    return f'<m:acc><m:accPr><m:chr m:val="{html.escape(ch)}"/></m:accPr><m:e>{inner_xml}</m:e></m:acc>'


def add_omml_equation(slide, left, top, width, height, omath_children_xml, jc="left",
                      fill_color=None, line_color=None, line_width=1.25, pad_lr=None):
    """Insert an editable equation in one text box (fill/border on the same shape — no stacked background).
    pad_lr: horizontal inner margin; LibreOffice scales OMML to the text-area width, so a wider margin keeps
    sub/superscripts off the border there (PowerPoint keeps the 18 pt size either way)."""
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    if fill_color is not None:
        tb.fill.solid()
        tb.fill.fore_color.rgb = fill_color
    if line_color is not None:
        tb.line.color.rgb = line_color
        tb.line.width = Pt(line_width)
    tf = tb.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    pad = pad_lr if pad_lr is not None else (0.06 if line_color else 0.02)
    tf.margin_left = tf.margin_right = Inches(pad)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    omml = (
        '<a14:m xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">'
        '<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<m:oMathParaPr><m:jc m:val="{jc}"/></m:oMathParaPr>'
        f"<m:oMath>{omath_children_xml}</m:oMath></m:oMathPara></a14:m>"
    )
    tf.paragraphs[0]._p.append(etree.fromstring(omml))
    return tb
