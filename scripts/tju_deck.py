#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tju_deck.py — the TJU template, slide skeletons, and the Deck / ContentPage builder.

New deck (preferred):
    deck = Deck()                                         # loads references/tju_template.pptx, removes samples
    deck.cover(["题目第一行", "题目第二行"], ["汇报人：…  导师：…", "2026 年 11 月"])
    deck.toc(["01  研究背景与现状", "02  研究内容与方案"], active=0)
    p = deck.content("研究内容与方案", "研究内容一：……（声明式结论句）",
                     lead=[("视觉场景理解：", "……"), ("干扰识别：", "……")],
                     takeaway="……一句话结论……")
    left, right = p.body.split(2)                        # Regions flush with the grid
    deck.save("out.pptx")                                # enforce 18 pt → staged save → audit report

Editing an existing deck in place: new_*_slide / setup_tju_content_slide / rearrange_slides / clear_slide_body.
"""

import os
import shutil
import tempfile

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from tju_components import add_dashed_box, add_text_block, no_theme_effects
from tju_layout import Region
from tju_text import append_rich_text, ensure_font_slots, normalize_paragraph, set_run_font, update_paragraph_text
from tju_theme import (
    BODY_BOTTOM, BODY_TOP, C_BLACK, C_BLUE, C_RED, C_WHITE, GRID_L, GRID_W, HEADER_RULE_Y, LEAD_LINE_H,
    LEAD_TOP, MIN_PT, PT_BANNER, PT_BODY, PT_SUBTITLE, PT_TAKEAWAY, TAKEAWAY_H, TAKEAWAY_TOP, VGAP,
    LayoutError, text_width_in,
)

SKILL_DIR = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
TEMPLATE_PATH = os.path.join(SKILL_DIR, "references", "tju_template.pptx")

_LAYOUT_NAMES = {"cover": ("标题幻灯片",), "thanks": ("标题幻灯片",), "content": ("1_自定义版式",),
                 "toc": ("1_仅标题",), "blank": ("空白",)}
_LAYOUT_FALLBACK = {"cover": 0, "thanks": 0, "content": 1, "toc": 2, "blank": 3}
MAX_LEAD_LINES = 3


# --------------------------------------------------------------------------- #
# Low-level slide operations
# --------------------------------------------------------------------------- #
def _remove_alternate_content(slide):
    """Drop mc:AlternateContent wrappers PowerPoint leaves around re-saved OMML (ghost shapes)."""
    tree = slide.shapes._spTree
    for child in list(tree):
        if child.tag.endswith("AlternateContent"):
            tree.remove(child)


def _in_mc_fallback(el):
    cur = el.getparent()
    while cur is not None:
        if cur.tag.endswith("}Fallback"):
            return True
        cur = cur.getparent()
    return False


def _clear_notes(slide):
    if slide.has_notes_slide:
        slide.notes_slide.notes_text_frame.text = ""


def _keep_only(slide, keep_elements):
    """Remove every shape except those whose XML element is in keep_elements (compare elements, not wrappers)."""
    for sh in list(slide.shapes):
        if sh._element not in keep_elements:
            sh._element.getparent().remove(sh._element)


def _write_lines(shape, lines, pt, color, align):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    while len(tf.paragraphs) > len(lines):
        tf._txBody.remove(tf.paragraphs[-1]._p)
    for i, line in enumerate(lines):
        p = tf.paragraphs[i] if i < len(tf.paragraphs) else tf.add_paragraph()
        update_paragraph_text(p, line, size_pt=pt, color=color, align=align)


def _page_number(slide, page_num, shape=None):
    if shape is None:
        shape = slide.shapes.add_textbox(Inches(12.35), Inches(7.05), Inches(0.75), Inches(0.36))
    else:
        shape.left, shape.top, shape.width, shape.height = Inches(12.35), Inches(7.05), Inches(0.75), Inches(0.36)
    shape.name = f"Slide Number {page_num}"
    tf = shape.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    while len(tf.paragraphs) > 1:
        tf._txBody.remove(tf.paragraphs[-1]._p)
    update_paragraph_text(tf.paragraphs[0], str(page_num), size_pt=PT_BODY, color=C_BLUE, align=PP_ALIGN.RIGHT)
    return shape


def clear_slide_body(slide):
    """Keep banner (shapes[0]) and subtitle (shapes[1]); remove everything else."""
    _remove_alternate_content(slide)
    _keep_only(slide, {sh._element for sh in list(slide.shapes)[:2]})


def setup_tju_content_slide(slide, banner_title, subtitle_text, page_num=None, subtitle_pt=PT_SUBTITLE,
                            banner_pt=PT_BANNER):
    """Content skeleton: banner title, subtitle with left bar, divider rule at y=1.56, page number."""
    _clear_notes(slide)
    _remove_alternate_content(slide)
    text_shapes = [sh for sh in slide.shapes if sh.has_text_frame]
    banner = (next((s for s in text_shapes if s.top.inches < 0.6 and s.left.inches > 0.6), None)
              or next((s for s in text_shapes if s.top.inches < 0.6), None)
              or (text_shapes[0] if text_shapes else None))
    if banner is None:
        _keep_only(slide, set())
        banner = slide.shapes.add_textbox(Inches(0.92), Inches(0.11), Inches(11.50), Inches(0.59))
    else:
        _keep_only(slide, {banner._element})
    banner.left, banner.top, banner.width, banner.height = Inches(0.92), Inches(0.11), Inches(11.50), Inches(0.59)
    _write_lines(banner, [banner_title], max(MIN_PT, banner_pt), C_WHITE, PP_ALIGN.LEFT)

    bar = no_theme_effects(slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(GRID_L), Inches(1.04), Inches(0.08), Inches(0.36)))
    bar.fill.solid(); bar.fill.fore_color.rgb = C_BLUE; bar.line.fill.background()
    bar.name = "Subtitle Bar"
    sub = slide.shapes.add_textbox(Inches(0.60), Inches(0.94), Inches(12.28), Inches(0.54))
    sub.name = "Subtitle"
    stf = sub.text_frame
    stf.word_wrap = True
    stf.vertical_anchor = MSO_ANCHOR.MIDDLE
    stf.paragraphs[0].alignment = PP_ALIGN.LEFT
    append_rich_text(stf.paragraphs[0], subtitle_text, size_pt=max(MIN_PT, subtitle_pt), color=C_BLUE)
    rule = no_theme_effects(slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(GRID_L), Inches(HEADER_RULE_Y), Inches(GRID_W), Inches(0.02)))
    rule.fill.solid(); rule.fill.fore_color.rgb = C_BLUE; rule.line.fill.background()
    rule.name = "Header Rule"
    if page_num is not None:
        _page_number(slide, page_num)
    return banner, sub


def format_cover_slide(slide, title_lines, meta_lines, title_pt=36, meta_pt=22, text_color=C_WHITE):
    _clear_notes(slide)
    _remove_alternate_content(slide)
    title_lines = title_lines.split("\n") if isinstance(title_lines, str) else list(title_lines)
    meta_lines = meta_lines.split("\n") if isinstance(meta_lines, str) else list(meta_lines)
    texts = sorted([sh for sh in slide.shapes if sh.has_text_frame], key=lambda s: s.top)
    title_sh = next((s for s in texts if s.top.inches < 3.0), None)
    meta_sh = next((s for s in texts if s is not title_sh and 3.0 <= s.top.inches < 6.5), None)
    title_sh = title_sh or slide.shapes.add_textbox(Inches(0.91), Inches(1.23), Inches(11.50), Inches(2.20))
    meta_sh = meta_sh or slide.shapes.add_textbox(Inches(1.71), Inches(3.94), Inches(9.90), Inches(1.40))
    _keep_only(slide, {title_sh._element, meta_sh._element})
    _write_lines(title_sh, title_lines, title_pt, text_color, PP_ALIGN.CENTER)
    _write_lines(meta_sh, meta_lines, meta_pt, text_color, PP_ALIGN.CENTER)
    return title_sh, meta_sh


def format_toc_slide(slide, toc_items, active_idx=0, page_num=None):
    """TOC / section divider: active chapter red with ▸, others blue."""
    _clear_notes(slide)
    _remove_alternate_content(slide)
    toc_sh = pg_sh = None
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        if "Placeholder" in sh.name or sh.top.inches > 6.5:
            pg_sh = sh
        elif toc_sh is None:
            toc_sh = sh
    toc_sh = toc_sh or slide.shapes.add_textbox(Inches(4.80), Inches(2.22), Inches(6.80), Inches(4.60))
    toc_sh.left, toc_sh.top, toc_sh.width, toc_sh.height = Inches(4.80), Inches(2.22), Inches(6.80), Inches(4.60)
    keep = {toc_sh._element}
    if pg_sh is not None and page_num is not None:
        keep.add(pg_sh._element)
    _keep_only(slide, keep)
    tf = toc_sh.text_frame
    tf.word_wrap = True
    while len(tf.paragraphs) < len(toc_items):
        tf.add_paragraph()
    while len(tf.paragraphs) > len(toc_items):
        tf._txBody.remove(tf.paragraphs[-1]._p)
    for idx, title in enumerate(toc_items):
        p = tf.paragraphs[idx]
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(18)
        for child in list(p._p):
            if not child.tag.endswith("pPr"):
                p._p.remove(child)
        act = idx == active_idx
        pt = 26                                   # one size for every row so the titles share a left edge
        marker = p.add_run()                      # same marker on every row (white when inactive)
        marker.text = "▸ "
        set_run_font(marker, size_pt=pt, color=C_RED if act else C_WHITE)
        r = p.add_run()
        r.text = title
        set_run_font(r, size_pt=pt, color=C_RED if act else C_BLUE)
    if page_num is not None:
        _page_number(slide, page_num, pg_sh if pg_sh is not None and pg_sh._element in keep else None)
    return toc_sh


def format_thanks_slide(slide, lines=("恳请各位老师批评指正", "谢谢！"), size_pt=40, text_color=C_WHITE):
    _clear_notes(slide)
    _remove_alternate_content(slide)
    lines = lines.split("\n") if isinstance(lines, str) else list(lines)
    main = next((sh for sh in slide.shapes if sh.has_text_frame), None)
    main = main or slide.shapes.add_textbox(Inches(1.71), Inches(2.20), Inches(9.90), Inches(2.60))
    _keep_only(slide, {main._element})
    _write_lines(main, lines, size_pt, text_color, PP_ALIGN.CENTER)
    return main


def rearrange_slides(prs, src_indices_0based):
    """Reorder / prune slides by 0-based index list; releases relationships of removed slides."""
    lst = prs.slides._sldIdLst
    orig = list(lst)
    keep = set(src_indices_0based)
    if len(keep) != len(src_indices_0based):
        raise ValueError("src_indices_0based must not contain duplicates")
    for i, el in enumerate(orig):
        if i not in keep:
            rid = el.get(qn("r:id"))
            lst.remove(el)
            if rid:
                prs.part.drop_rel(rid)
    for el in list(lst):
        lst.remove(el)
    for i in src_indices_0based:
        lst.append(orig[i])


def clear_all_slides(prs):
    rearrange_slides(prs, [])


def get_tju_layout(prs, kind="content"):
    """Find the TJU layout by semantic kind (works on the 1-master template and legacy 2-master decks)."""
    if kind not in _LAYOUT_NAMES:
        raise ValueError(f"unknown layout kind {kind!r}; use {list(_LAYOUT_NAMES)}")
    for master in prs.slide_masters:
        for lay in master.slide_layouts:
            if lay.name in _LAYOUT_NAMES[kind]:
                return lay
    idx = _LAYOUT_FALLBACK[kind]
    if idx < len(prs.slide_layouts):
        return prs.slide_layouts[idx]
    raise LookupError(f"no TJU layout for kind={kind!r}")


def new_cover_slide(prs, title_lines, meta_lines, title_pt=36, meta_pt=22, text_color=C_WHITE):
    s = prs.slides.add_slide(get_tju_layout(prs, "cover"))
    format_cover_slide(s, title_lines, meta_lines, title_pt, meta_pt, text_color)
    return s


def new_toc_slide(prs, toc_items, active_idx=0, page_num=None):
    s = prs.slides.add_slide(get_tju_layout(prs, "toc"))
    format_toc_slide(s, toc_items, active_idx, page_num)
    return s


def new_content_slide(prs, banner_title, subtitle_text, page_num=None, subtitle_pt=PT_SUBTITLE, banner_pt=PT_BANNER):
    s = prs.slides.add_slide(get_tju_layout(prs, "content"))
    setup_tju_content_slide(s, banner_title, subtitle_text, page_num, subtitle_pt, banner_pt)
    return s


def new_thanks_slide(prs, lines=("恳请各位老师批评指正", "谢谢！"), size_pt=40, text_color=C_WHITE):
    s = prs.slides.add_slide(get_tju_layout(prs, "thanks"))
    format_thanks_slide(s, lines, size_pt, text_color)
    return s


def enforce_min_18pt_all_slides(prs):
    """Final guard: every text run, table cell, and OMML run >= 18 pt, bold, four font slots; clear notes."""
    from pptx.util import Pt as _Pt
    for slide in prs.slides:
        _clear_notes(slide)
        for sh in slide.shapes:
            frames = []
            if sh.has_text_frame:
                frames.append(sh.text_frame)
            if getattr(sh, "has_table", False) and sh.has_table:
                frames += [c.text_frame for row in sh.table.rows for c in row.cells]
            for tf in frames:
                for p in tf.paragraphs:
                    for r in p.runs:
                        if r.font.size is None or r.font.size.pt < MIN_PT:
                            r.font.size = _Pt(MIN_PT)
                        rPr = r._r.get_or_add_rPr()
                        rPr.set("b", "1")
                        ensure_font_slots(rPr)
        for el in slide._element.iter():
            if not (isinstance(el.tag, str) and el.tag.endswith("}r") and "math" in el.tag) or _in_mc_fallback(el):
                continue
            arpr = el.find(qn("a:rPr"))
            if arpr is None:
                arpr = etree.Element(qn("a:rPr"))
                mt = el.find("{http://schemas.openxmlformats.org/officeDocument/2006/math}t")
                el.insert(list(el).index(mt), arpr) if mt is not None else el.append(arpr)
            if arpr.get("sz") is None or int(arpr.get("sz")) < 1800:
                arpr.set("sz", "1800")
            arpr.set("b", "1")
            ensure_font_slots(arpr)


def save_staged_pptx(prs, target_path, staged_path=None, fallback_suffix="_staged"):
    """Save to a temp file then copy, so a deck open in PowerPoint never aborts the build."""
    target = os.path.abspath(target_path)
    os.makedirs(os.path.dirname(target), exist_ok=True)
    auto = staged_path is None
    if auto:
        fd, staged_path = tempfile.mkstemp(prefix="tju_ppt_", suffix=".pptx")
        os.close(fd)
    prs.save(staged_path)
    try:
        shutil.copyfile(staged_path, target)
        return target
    except PermissionError:
        stem, ext = os.path.splitext(target)
        alt = f"{stem}{fallback_suffix}{ext}"
        shutil.copyfile(staged_path, alt)
        return alt
    finally:
        if auto and os.path.exists(staged_path):
            try:
                os.remove(staged_path)
            except OSError:
                pass


# --------------------------------------------------------------------------- #
# Builder
# --------------------------------------------------------------------------- #
class ContentPage:
    """A content slide with its zones resolved: optional lead lines on top, optional takeaway at the bottom,
    and `body` — the Region left for figures, cards, and tables."""

    def __init__(self, slide, page_num, lead=None, takeaway=None):
        self.slide = slide
        self.page_num = page_num
        self.lead_lines = list(lead or [])
        self.takeaway_text = takeaway
        if len(self.lead_lines) > MAX_LEAD_LINES:
            raise LayoutError(f"[tju-ppt] page {page_num}: {len(self.lead_lines)} lead lines > {MAX_LEAD_LINES}; "
                              f"move detail into the figure or the speech.")
        if self.lead_lines:
            self._draw_lead()
        if takeaway:
            self._draw_takeaway()

    @property
    def body(self):
        top = LEAD_TOP + len(self.lead_lines) * LEAD_LINE_H + VGAP if self.lead_lines else BODY_TOP
        bottom = TAKEAWAY_TOP - VGAP if self.takeaway_text else BODY_BOTTOM
        return Region(GRID_L, top, GRID_W, bottom - top)

    def columns(self, n=2, ratios=None, dashed=False):
        """Split the body into columns.  dashed=True draws dotted containers and returns their inner regions —
        use it only for loose text + small figures, never around panels or cards (they have their own frame)."""
        cols = self.body.split(n, ratios)
        if not dashed:
            return cols
        for c in cols:
            add_dashed_box(self.slide, *c.box)
        return [c.inset(0.15, 0.10) for c in cols]

    # ---------------- zones ----------------
    def _draw_lead(self):
        paras = []
        for item in self.lead_lines:
            if isinstance(item, tuple) and len(item) in (2, 3) and all(isinstance(t, str) for t in item[:2]) \
                    and (len(item) == 2 or item[2] in ("blue", "red")):
                label_color = C_RED if len(item) == 3 and item[2] == "red" else C_BLUE
                para = [("▸ " + item[0], label_color, PT_BODY), (item[1], C_BLACK, PT_BODY)]
            else:
                para = item
            text = "".join(s[0] for s in normalize_paragraph(para))
            if text_width_in(text, PT_BODY) > GRID_W - 0.15:
                raise LayoutError(f"[tju-ppt] page {self.page_num}: lead line does not fit one line "
                                  f"({text_width_in(text):.2f} in > {GRID_W - 0.15:.2f} in): {text!r}")
            paras.append(para)
        tb = add_text_block(self.slide, GRID_L, LEAD_TOP, GRID_W, len(paras) * LEAD_LINE_H, paras, line_spacing=1.12)
        for p in tb.text_frame.paragraphs:
            p.space_after = Pt(0)
        tb.name = "Lead"

    def _draw_takeaway(self):
        t = self.takeaway_text
        if text_width_in(t, PT_TAKEAWAY) > GRID_W - 0.20:
            raise LayoutError(f"[tju-ppt] page {self.page_num}: takeaway does not fit one line at {PT_TAKEAWAY} pt "
                              f"({text_width_in(t, PT_TAKEAWAY):.2f} in); shorten the sentence, keep the size.")
        tb = add_text_block(self.slide, GRID_L, TAKEAWAY_TOP, GRID_W, TAKEAWAY_H, [(t, C_RED, PT_TAKEAWAY)],
                            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        tb.name = "Takeaway"


class Deck:
    """New deck on the TJU template.  Page numbers are assigned automatically (cover = 1)."""

    def __init__(self, template=TEMPLATE_PATH):
        self.prs = Presentation(template)
        clear_all_slides(self.prs)
        self.skip_audit = set()          # cover / toc / thanks pages excluded from content audit

    @property
    def next_num(self):
        return len(self.prs.slides) + 1

    def cover(self, title_lines, meta_lines, **kw):
        self.skip_audit.add(self.next_num)
        return new_cover_slide(self.prs, title_lines, meta_lines, **kw)

    def toc(self, items, active=0):
        n = self.next_num
        self.skip_audit.add(n)
        return new_toc_slide(self.prs, items, active, page_num=n)

    def content(self, banner, subtitle, lead=None, takeaway=None):
        n = self.next_num
        slide = new_content_slide(self.prs, banner, subtitle, page_num=n)
        return ContentPage(slide, n, lead=lead, takeaway=takeaway)

    def thanks(self, lines=("恳请各位老师批评指正", "谢谢！")):
        self.skip_audit.add(self.next_num)
        return new_thanks_slide(self.prs, lines)

    def save(self, path, audit=True, content=True, known=()):
        """enforce 18 pt → staged save → structural audit (+ content warnings).  Returns (path, issues).
        known: acronyms / product names already explained (e.g. ("PID", "PX4")), suppressing first-use warnings."""
        enforce_min_18pt_all_slides(self.prs)
        out = save_staged_pptx(self.prs, path)
        issues = []
        if audit:
            from tju_audit import run_audit
            issues = run_audit(out, skip=self.skip_audit, content=content, known=known, verbose=True)
        return out, issues
