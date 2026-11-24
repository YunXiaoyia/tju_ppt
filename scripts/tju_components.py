#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tju_components.py — slide components.  All colors / sizes come from tju_theme.

Region-based (preferred; geometry from tju_layout, fit checked at build time → LayoutError):
    add_figure_panel(slide, region, header, image, tone, fit)   header bar + image flush underneath
    figure_slot(region) / image_card_slot(region, n)            (w, h) to draw a figure at 1:1 size
    add_card(slide, region, header, body_lines, tone)           joined header + body card (card_height())
    add_image_card(slide, region, header, image, caption, tone) header + image + caption (challenge cards)
    add_table(slide, region, col_ratios, rows)                  4-color table, row heights from wrapped text

Coordinate-based primitives (kept for in-place edits of existing decks):
    add_text_block, add_solid_box, add_joined_card, add_dashed_box, add_clean_tju_table,
    add_image_fit, add_arrow, update_table_cells
"""

from lxml import etree
from PIL import Image
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from tju_layout import as_box
from tju_text import normalize_paragraph, paragraph_text, update_paragraph_text, write_paragraph
from tju_theme import (
    ASPECT_TOL, C_BLACK, C_BLUE, C_BLUE_LIGHT, C_WHITE, MIN_PT, PANEL_HEADER_H, PT_BODY, PT_HEADING,
    TONES, LayoutError, lines_needed, text_width_in,
)

LINE_H_FACTOR = 1.25   # rendered line height ≈ 1.25 × font size for bold 微软雅黑


# --------------------------------------------------------------------------- #
# Fit checks
# --------------------------------------------------------------------------- #
def text_capacity_check(lines, width_in, height_in, pt=PT_BODY, margin_lr=0.10, margin_tb=0.05,
                        line_spacing=1.10, label="box", strict=True):
    """Raise LayoutError (or return False) if `lines` cannot fit the box at `pt` without shrinking."""
    if isinstance(lines, (str, tuple)):
        lines = [lines]
    need_lines = 0
    for item in lines:
        segs = normalize_paragraph(item, pt)
        p_pt = max(s[2] for s in segs) if segs else pt
        need_lines += lines_needed(paragraph_text(item), width_in - 2 * margin_lr, p_pt, pad_in=0.04)
    need_h = need_lines * pt * LINE_H_FACTOR * line_spacing / 72.0 + 2 * margin_tb
    ok = need_h <= height_in + 0.02
    if not ok and strict:
        raise LayoutError(
            f"[tju-ppt] {label}: needs {need_lines} line(s) ≈ {need_h:.2f} in but the box is "
            f"{width_in:.2f}×{height_in:.2f} in at {pt} pt. Shorten the text or give it a larger region "
            f"(never drop below {MIN_PT} pt)."
        )
    return ok


def _one_line_check(text, width_in, pt, label):
    if text_width_in(text, pt) > width_in - 0.12:
        raise LayoutError(f"[tju-ppt] {label} {text!r} does not fit one line in {width_in:.2f} in at {pt} pt; shorten it.")


# --------------------------------------------------------------------------- #
# Primitives
# --------------------------------------------------------------------------- #
def no_theme_effects(sp):
    """Zero the <p:style> effect reference so autoshapes never inherit the theme's shadow."""
    style = sp._element.find(qn("p:style"))
    if style is not None:
        eff = style.find(qn("a:effectRef"))
        if eff is not None:
            eff.set("idx", "0")
    return sp


def _frame(tf, margin_lr, margin_tb, anchor=MSO_ANCHOR.MIDDLE):
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(margin_lr)
    tf.margin_top = tf.margin_bottom = Inches(margin_tb)
    tf.vertical_anchor = anchor


def add_text_block(slide, left, top, width, height, lines, default_pt=PT_BODY, default_color=C_BLACK,
                   align=PP_ALIGN.LEFT, line_spacing=1.10, anchor=MSO_ANCHOR.TOP):
    """Borderless multi-paragraph text box (paragraph specs per tju_text)."""
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    _frame(tb.text_frame, 0.04, 0.02, anchor)
    for i, item in enumerate(lines):
        p = tb.text_frame.paragraphs[0] if i == 0 else tb.text_frame.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        p.space_after = Pt(2)
        write_paragraph(p, item, default_pt, default_color)
    return tb


def add_solid_box(slide, left, top, width, height, text_lines, fill_color=None, line_color=C_BLUE,
                  text_color=C_BLACK, size_pt=PT_BODY, align=PP_ALIGN.CENTER, line_width=1.5,
                  margin_lr=0.07, margin_tb=0.04, line_spacing=1.08):
    """Rectangle with explicit fill/border (never theme fallback) and rich text."""
    sp = no_theme_effects(slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width),
                                                 Inches(height)))
    if fill_color is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill_color
    if line_color is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line_color
        sp.line.width = Pt(line_width)
    _frame(sp.text_frame, margin_lr, margin_tb)
    if isinstance(text_lines, (str, tuple)):
        text_lines = [text_lines]
    for i, item in enumerate(text_lines):
        p = sp.text_frame.paragraphs[0] if i == 0 else sp.text_frame.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        write_paragraph(p, item, size_pt, text_color)
    return sp


def add_joined_card(slide, left, top, width, header_h, body_h, header_text, body_lines,
                    header_bg=C_BLUE, body_bg=C_BLUE_LIGHT, border_color=C_BLUE,
                    header_pt=PT_HEADING, body_pt=PT_BODY, align=PP_ALIGN.LEFT):
    """Header bar touching a body box (oy = 0, no declared overlap)."""
    hdr = add_solid_box(slide, left, top, width, header_h, [header_text], fill_color=header_bg,
                        line_color=border_color, text_color=C_WHITE, size_pt=header_pt,
                        align=PP_ALIGN.LEFT, margin_lr=0.10, margin_tb=0.03)
    body = add_solid_box(slide, left, top + header_h, width, body_h, body_lines, fill_color=body_bg,
                         line_color=border_color, text_color=C_BLACK, size_pt=body_pt, align=align,
                         margin_lr=0.10, margin_tb=0.05)
    return hdr, body


def add_dashed_box(slide, left, top, width, height, line_color=C_BLACK, weight_pt=1.5):
    """Unfilled dotted column container.  Only for columns of loose text + small figures; never around
    a panel or card that already has its own header bar (one container level)."""
    sp = no_theme_effects(slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width),
                                                 Inches(height)))
    sp.fill.background()
    sp.line.color.rgb = line_color
    sp.line.width = Pt(weight_pt)
    ln = sp.line._get_or_add_ln()
    etree.SubElement(ln, qn("a:prstDash")).set("val", "sysDot")
    return sp


def add_arrow(slide, left, top, width, height, direction="right", color=C_BLUE):
    shape_type = {"right": MSO_SHAPE.RIGHT_ARROW, "down": MSO_SHAPE.DOWN_ARROW,
                  "up": MSO_SHAPE.UP_ARROW, "left": MSO_SHAPE.LEFT_ARROW}[direction]
    sp = no_theme_effects(slide.shapes.add_shape(shape_type, Inches(left), Inches(top), Inches(width), Inches(height)))
    sp.fill.solid()
    sp.fill.fore_color.rgb = color
    sp.line.fill.background()
    return sp


def image_aspect(path):
    with Image.open(path) as im:
        return im.size[0] / im.size[1]


def add_image_fit(slide, img_path, left, top, box_w, box_h, border_color=C_BLUE, border_pt=0.75):
    """Letterbox an image inside a box (keeps ratio, centers).  Do NOT use under a header bar —
    use add_figure_panel so the image is flush with the header."""
    ratio = image_aspect(img_path)
    if box_w / box_h > ratio:
        dh, dw = box_h, box_h * ratio
    else:
        dw, dh = box_w, box_w / ratio
    pic = slide.shapes.add_picture(img_path, Inches(left + (box_w - dw) / 2), Inches(top + (box_h - dh) / 2),
                                   width=Inches(dw), height=Inches(dh))
    if border_color is not None:
        pic.line.color.rgb = border_color
        pic.line.width = Pt(border_pt)
    return pic


def check_aspect(img_path, region, fit="exact", label="image"):
    """Raise LayoutError when fit='exact' and the image aspect differs from the slot by > ASPECT_TOL."""
    x, y, w, h = as_box(region)
    img_ar, slot_ar = image_aspect(img_path), w / h
    mismatch = abs(img_ar / slot_ar - 1.0)
    if fit == "exact" and mismatch > ASPECT_TOL:
        raise LayoutError(
            f"[tju-ppt] {label}: image aspect {img_ar:.2f} ≠ slot {slot_ar:.2f} ({mismatch:.0%} off). "
            f"Redraw it at {w:.2f}×{h:.2f} in, or pass fit='crop' for photos."
        )
    return img_ar, slot_ar, mismatch


def add_picture_exact(slide, img_path, region, fit="exact", border_color=None, border_pt=1.0, label="image"):
    """Fill `region` exactly.  fit="exact": image must already have the slot's aspect ratio (±ASPECT_TOL)
    — draw figures at figure_slot() size.  fit="crop": center-crop (photos, GIFs) via picture crop."""
    x, y, w, h = as_box(region)
    img_ar, slot_ar, mismatch = check_aspect(img_path, region, fit, label)
    pic = slide.shapes.add_picture(img_path, Inches(x), Inches(y), width=Inches(w), height=Inches(h))
    if fit == "crop" and mismatch > 1e-3:
        if img_ar > slot_ar:                       # too wide → crop left/right
            c = (1 - slot_ar / img_ar) / 2
            pic.crop_left = pic.crop_right = c
        else:                                      # too tall → crop top/bottom
            c = (1 - img_ar / slot_ar) / 2
            pic.crop_top = pic.crop_bottom = c
    if border_color is not None:
        pic.line.color.rgb = border_color
        pic.line.width = Pt(border_pt)
    return pic


# --------------------------------------------------------------------------- #
# Region-based components
# --------------------------------------------------------------------------- #
def figure_slot(region, header_h=PANEL_HEADER_H):
    """(width, height) in inches of the image area of add_figure_panel(region) — draw figures at this size."""
    _, _, w, h = as_box(region)
    return w, h - header_h


def add_figure_panel(slide, region, header, image, tone="blue", fit="exact", header_h=PANEL_HEADER_H,
                     header_pt=PT_BODY):
    """Header bar + image filling the rest of `region`, both exactly region-wide (flush edges)."""
    x, y, w, h = as_box(region)
    hdr_fill, _, border = TONES[tone]
    _one_line_check(header, w, header_pt, "panel header")
    check_aspect(image, (x, y + header_h, w, h - header_h), fit, f"panel {header!r}")
    hdr = add_solid_box(slide, x, y, w, header_h, [header], fill_color=hdr_fill, line_color=hdr_fill,
                        text_color=C_WHITE, size_pt=header_pt, margin_lr=0.06, margin_tb=0.02)
    pic = add_picture_exact(slide, image, (x, y + header_h, w, h - header_h), fit=fit,
                            border_color=border, border_pt=1.0, label=f"panel {header!r}")
    return hdr, pic


def card_height(width, body_lines, line_spacing=1.15, header_h=PANEL_HEADER_H, pad=0.30):
    """Height a card needs for its text (use with region.take_top(card_height(...)) to avoid hollow cards)."""
    if isinstance(body_lines, (str, tuple)):
        body_lines = [body_lines]
    n = sum(lines_needed(paragraph_text(it), width - 0.24, PT_BODY, pad_in=0.04) for it in body_lines)
    return header_h + n * PT_BODY * LINE_H_FACTOR * line_spacing / 72.0 + pad


def add_card(slide, region, header, body_lines, tone="blue", header_h=PANEL_HEADER_H,
             align=PP_ALIGN.LEFT, line_spacing=1.15, header_pt=PT_BODY, header_align=PP_ALIGN.CENTER,
             body_anchor=MSO_ANCHOR.MIDDLE):
    """Joined card filling `region`; raises LayoutError if the body text cannot fit at 18 pt.
    Size the region with card_height() when the text is short, so the card is not mostly empty, or pass
    body_anchor=MSO_ANCHOR.TOP for a list that should start under the header."""
    x, y, w, h = as_box(region)
    hdr_fill, body_fill, border = TONES[tone]
    _one_line_check(header, w, header_pt, "card header")
    text_capacity_check(body_lines, w, h - header_h, margin_lr=0.12, margin_tb=0.08,
                        line_spacing=line_spacing, label=f"card {header!r}")
    hdr = add_solid_box(slide, x, y, w, header_h, [header], fill_color=hdr_fill, line_color=border,
                        text_color=C_WHITE, size_pt=header_pt, align=header_align, margin_lr=0.12, margin_tb=0.02)
    body = add_solid_box(slide, x, y + header_h, w, h - header_h, body_lines, fill_color=body_fill,
                         line_color=border, text_color=C_BLACK, align=align, line_width=1.2,
                         margin_lr=0.12, margin_tb=0.08 if body_anchor == MSO_ANCHOR.MIDDLE else 0.18,
                         line_spacing=line_spacing)
    body.text_frame.vertical_anchor = body_anchor
    return hdr, body


def _caption_h(n_lines):
    return 0.16 + n_lines * PT_BODY * LINE_H_FACTOR * 1.15 / 72.0


def image_card_slot(region, n_caption_lines=2, header_h=PANEL_HEADER_H):
    """(width, height) of the image area of add_image_card(region) with n caption lines."""
    _, _, w, h = as_box(region)
    return w, h - header_h - _caption_h(n_caption_lines)


def add_image_card(slide, region, header, image, caption_lines, tone="blue", caption_h=None,
                   fit="exact", header_h=PANEL_HEADER_H):
    """Header + image + short caption stacked in one column (e.g. one of three challenge cards)."""
    x, y, w, h = as_box(region)
    if isinstance(caption_lines, str):
        caption_lines = caption_lines.split("\n")
    caption_h = caption_h or _caption_h(len(caption_lines))
    hdr_fill, body_fill, border = TONES[tone]
    _one_line_check(header, w, PT_HEADING, "card header")
    text_capacity_check(caption_lines, w, caption_h, margin_lr=0.08, margin_tb=0.06, line_spacing=1.15,
                        label=f"caption of {header!r}")
    check_aspect(image, (x, y + header_h, w, h - header_h - caption_h), fit, f"image of {header!r}")
    hdr = add_solid_box(slide, x, y, w, header_h, [header], fill_color=hdr_fill, line_color=hdr_fill,
                        text_color=C_WHITE, size_pt=PT_HEADING, margin_lr=0.06, margin_tb=0.02)
    img_h = h - header_h - caption_h
    pic = add_picture_exact(slide, image, (x, y + header_h, w, img_h), fit=fit, border_color=border,
                            label=f"image of {header!r}")
    cap = add_solid_box(slide, x, y + header_h + img_h, w, caption_h, caption_lines, fill_color=body_fill,
                        line_color=border, line_width=1.2, margin_lr=0.08, margin_tb=0.06, line_spacing=1.15)
    return hdr, pic, cap


# --------------------------------------------------------------------------- #
# Tables
# --------------------------------------------------------------------------- #
def _set_cell_borders(cell, hex_color="00337F", width_pt=0.75):
    tcPr = cell._tc.get_or_add_tcPr()
    w_emu = str(int(round(width_pt * 12700)))
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        old = tcPr.find(qn(tag))
        if old is not None:
            tcPr.remove(old)
        ln = etree.SubElement(tcPr, qn(tag), w=w_emu, cap="flat", cmpd="sng", algn="ctr")
        sf = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr"), val=hex_color)


def _build_table(slide, left, top, width, col_widths, row_heights, rows_data, col_aligns, header_pt,
                 body_pt, badge_col0):
    n_rows, n_cols = len(rows_data), len(rows_data[0])
    shape = slide.shapes.add_table(n_rows, n_cols, Inches(left), Inches(top), Inches(width),
                                   Inches(sum(row_heights)))
    tbl = shape.table
    tblPr = tbl._tbl.tblPr
    for child in list(tblPr):
        if child.tag.endswith("tableStyleId"):       # drop theme table style → no schemeClr leaks
            tblPr.remove(child)
    for ci, cw in enumerate(col_widths):
        tbl.columns[ci].width = Inches(cw)
    for ri, row in enumerate(rows_data):
        tbl.rows[ri].height = Inches(row_heights[ri])
        is_hdr = ri == 0
        for ci, val in enumerate(row):
            is_badge = badge_col0 and not is_hdr and ci == 0
            bg = C_BLUE if (is_hdr or is_badge) else (C_BLUE_LIGHT if ri % 2 == 1 else C_WHITE)
            fg = C_WHITE if (is_hdr or is_badge) else C_BLACK
            cell = tbl.cell(ri, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            _set_cell_borders(cell, "FFFFFF" if is_hdr else "00337F")
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = cell.margin_right = Inches(0.08)
            cell.margin_top = cell.margin_bottom = Inches(0.04)
            tf = cell.text_frame
            tf.word_wrap = True
            while len(tf.paragraphs) > 1:
                tf._txBody.remove(tf.paragraphs[-1]._p)
            p = tf.paragraphs[0]
            if is_hdr or is_badge:
                p.alignment = PP_ALIGN.CENTER
            elif col_aligns and ci < len(col_aligns):
                p.alignment = col_aligns[ci]
            else:
                p.alignment = PP_ALIGN.CENTER if ci == 0 else PP_ALIGN.LEFT
            for child in list(p._p):
                if not child.tag.endswith("pPr"):
                    p._p.remove(child)
            write_paragraph(p, val if isinstance(val, list) else str(val),
                            header_pt if is_hdr else body_pt, fg)
    return shape


def add_table(slide, region, col_ratios, rows_data, col_aligns=None, header_pt=PT_HEADING,
              body_pt=PT_BODY, badge_col0=False, fill=False):
    """4-color table placed at region.x/y/w; row heights follow the wrapped text (fill=True spreads the
    spare height over the body rows).  Raises LayoutError when the rows need more height than the region."""
    x, y, w, h = as_box(region)
    tot = float(sum(col_ratios))
    col_widths = [w * r / tot for r in col_ratios]
    heights = []
    for ri, row in enumerate(rows_data):
        pt = header_pt if ri == 0 else body_pt
        n = max(lines_needed(paragraph_text(v) if isinstance(v, list) else str(v), cw - 0.16, pt, pad_in=0.04)
                for v, cw in zip(row, col_widths))
        heights.append(max(0.40, n * pt * LINE_H_FACTOR / 72.0 + 0.12))
    if sum(heights) > h + 0.02:
        raise LayoutError(f"[tju-ppt] table needs {sum(heights):.2f} in but region is {h:.2f} in high; "
                          f"cut rows/words or widen columns.")
    if fill and len(heights) > 1:
        extra = (h - sum(heights)) / (len(heights) - 1)
        heights = [heights[0]] + [v + extra for v in heights[1:]]
    return _build_table(slide, x, y, w, col_widths, heights, rows_data, col_aligns, header_pt, body_pt, badge_col0)


def add_clean_tju_table(slide, left, top, width, height, col_widths, rows_data, col_aligns=None,
                        header_pt=PT_HEADING, body_pt=PT_BODY, badge_col0=False):
    """Legacy coordinate form with equal row heights (prefer add_table)."""
    row_h = height / len(rows_data)
    return _build_table(slide, left, top, width, col_widths, [row_h] * len(rows_data), rows_data,
                        col_aligns, header_pt, body_pt, badge_col0)


def update_table_cells(table, cell_data, header_pt=PT_HEADING, body_pt=PT_BODY):
    """Rewrite an existing table's text (header row white, body black, >= 18 pt)."""
    for r_idx, row_vals in enumerate(cell_data):
        is_header = r_idx == 0
        pt = max(MIN_PT, header_pt if is_header else body_pt)
        color = C_WHITE if is_header else C_BLACK
        for c_idx, val in enumerate(row_vals):
            cell = table.cell(r_idx, c_idx)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            lines = [val] if isinstance(val, str) else val
            while len(tf.paragraphs) < len(lines):
                tf.add_paragraph()
            while len(tf.paragraphs) > len(lines):
                tf._txBody.remove(tf.paragraphs[-1]._p)
            for p_idx, line in enumerate(lines):
                update_paragraph_text(tf.paragraphs[p_idx], line, size_pt=pt, color=color)
