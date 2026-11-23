#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TJU-PPT 可复用学术汇报排版辅助库 (ppt_builder_helpers.py)

核心职责（将机械性约束自动化）：
1. 字体四槽位强绑定：中文 微软雅黑 (a:ea)，英文/数字/符号 Times New Roman (a:latin, a:cs, a:sym)，粗体，>=18pt。
2. 严格四色规范：#00337F（天大深蓝）、#EAF1FB（浅蓝填充）、#C00000（强调红）、#000000（纯黑）、#FFFFFF（反白）。
3. 行内上下标真实排版：解析 _{sub} 与 ^{sup} 生成 OpenXML baseline (-25000 / 30000)，保持 sz>=1800 合规，避免裸露代码下划线。
4. 原生可编辑 OMML 公式 (<a14:m>)：重音帽默认使用组合符 U+0302 防止压字，支持单形状浅蓝高亮公式框。
5. 标准骨架与结构化组件：内容页骨架（含清除 mc:AlternateContent 残留）、封面/目录/致谢页、连体卡片、四色表格、安全重排页与暂存写回。
"""

import html
import os
import re
import shutil
import tempfile
from PIL import Image
from lxml import etree
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# 严格限定的四色规范（含浅蓝次级填充）
C_BLUE = RGBColor(0x00, 0x33, 0x7F)        # 主色深蓝（标题、核心框、主边框、页码）
C_BLUE_LIGHT = RGBColor(0xEA, 0xF1, 0xFB)  # 浅蓝填充（次级卡片、公式高亮底、表格交替行）
C_RED = RGBColor(0xC0, 0x00, 0x00)         # 强调红（核心判据、失效边界、关键创新）
C_BLACK = RGBColor(0x00, 0x00, 0x00)       # 正文纯黑
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)       # 反白文字与纯白底
ALLOWED_HEX = {"00337F", "EAF1FB", "C00000", "000000", "FFFFFF"}

SUB_SUP_PATTERN = re.compile(r"(\_\{[^{}]+\}|\^\{[^{}]+\})")

_FONT_SLOTS = (
    ("a:latin", "Times New Roman"),
    ("a:ea", "微软雅黑"),
    ("a:cs", "Times New Roman"),
    ("a:sym", "Times New Roman"),
)


def _ensure_four_font_slots(rPr):
    for tag, typeface in _FONT_SLOTS:
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", typeface)


def _remove_alternate_content(slide):
    """清除 spTree 中残留的 mc:AlternateContent（PowerPoint 二次保存 OMML 后产生的包装节点）。"""
    spTree = slide.shapes._spTree
    for child in list(spTree):
        if child.tag.endswith("AlternateContent"):
            spTree.remove(child)


def _in_mc_fallback(el):
    cur = el.getparent()
    while cur is not None:
        if cur.tag.endswith("}Fallback"):
            return True
        cur = cur.getparent()
    return False


def set_run_font(run, size_pt=18, bold=True, color=C_BLACK, italic=False, baseline=None):
    """设置 run 四槽位字体（中文微软雅黑 + 英文 Times New Roman）、字号 >=18pt 与可选上下标 baseline"""
    final_pt = max(18, size_pt)
    is_bold = True if bold is None else bold
    run.font.size = Pt(final_pt)
    run.font.bold = is_bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    run.font.name = "Times New Roman"
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
    _ensure_four_font_slots(rPr)


def append_rich_text(paragraph, text, size_pt=18, color=C_BLACK, bold=True):
    """解析 _{sub} 与 ^{sup} 标记并写入真实 OpenXML 上下标 run；普通标识符（如 ERR_CODE）原样保留"""
    for tok in SUB_SUP_PATTERN.split(text):
        if not tok:
            continue
        r = paragraph.add_run()
        if tok.startswith("_{") and tok.endswith("}"):
            r.text = tok[2:-1]
            set_run_font(r, size_pt=size_pt, bold=bold, color=color, baseline=-25000)
        elif tok.startswith("^{") and tok.endswith("}"):
            r.text = tok[2:-1]
            set_run_font(r, size_pt=size_pt, bold=bold, color=color, baseline=30000)
        else:
            r.text = tok
            set_run_font(r, size_pt=size_pt, bold=bold, color=color, baseline=None)


def update_paragraph_text(paragraph, new_text, size_pt=18, bold=True, color=None, align=None):
    """替换已有段落文本，支持 _{sub}/^{sup}，强制 >=18pt 且中英分设微软雅黑与 Times New Roman"""
    if align is not None:
        paragraph.alignment = align
    orig_pt = paragraph.runs[0].font.size.pt if (len(paragraph.runs) > 0 and paragraph.runs[0].font.size) else size_pt
    final_pt = max(18, size_pt if size_pt else orig_pt)
    for child in list(paragraph._p):
        if not child.tag.endswith("pPr"):
            paragraph._p.remove(child)
    append_rich_text(paragraph, new_text, size_pt=final_pt, color=color or C_BLACK, bold=bold)


def update_shape_paragraphs(shape, lines, size_pt=18, bold=True, color=None):
    tf = shape.text_frame
    for idx, line in enumerate(lines):
        if idx < len(tf.paragraphs):
            update_paragraph_text(tf.paragraphs[idx], line, size_pt=size_pt, bold=bold, color=color)


def update_table_cells(table, cell_data, header_pt=19, body_pt=18):
    """更新表格内容，确保字号全部 >= 18pt，表头蓝底白字，正文支持行内上下标"""
    for r_idx, row_vals in enumerate(cell_data):
        is_header = (r_idx == 0)
        pt = max(18, header_pt if is_header else body_pt)
        txt_color = C_WHITE if is_header else C_BLACK
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
                update_paragraph_text(tf.paragraphs[p_idx], line, size_pt=pt, bold=True, color=txt_color)


def clear_slide_body(slide):
    """保留顶部标题栏(shapes[0])与副标题(shapes[1])，按底层 XML 元素引用清除其余旧形状及残留 AlternateContent"""
    _remove_alternate_content(slide)
    keep_els = {sh._element for sh in list(slide.shapes)[:2]}
    for sp in list(slide.shapes):
        if sp._element not in keep_els:
            sp._element.getparent().remove(sp._element)


def setup_tju_content_slide(slide, banner_title, subtitle_text, page_num=None, subtitle_pt=20, banner_pt=28):
    """构建标准天大学术内容页骨架：顶部蓝条标题 + 副标题左侧竖条锚点 + 顶部分隔细线 + 右下角页码。
    自动清理 mc:AlternateContent 残留，并在空白版式（0 形状）上自动创建标题文本框。"""
    if slide.has_notes_slide:
        slide.notes_slide.notes_text_frame.text = ""
    _remove_alternate_content(slide)

    banner = None
    for sh in slide.shapes:
        if sh.has_text_frame and sh.top.inches < 0.6 and sh.left.inches > 0.6:
            banner = sh
            break
    if banner is None:
        for sh in slide.shapes:
            if sh.has_text_frame and sh.top.inches < 0.6:
                banner = sh
                break
    if banner is None:
        for sh in slide.shapes:
            if sh.has_text_frame:
                banner = sh
                break
    if banner is None:
        for sh in list(slide.shapes):
            sh._element.getparent().remove(sh._element)
        banner = slide.shapes.add_textbox(Inches(0.92), Inches(0.11), Inches(11.50), Inches(0.59))
    else:
        banner_el = banner._element
        for sh in list(slide.shapes):
            if sh._element is not banner_el:
                sh._element.getparent().remove(sh._element)

    banner.left, banner.top = Inches(0.92), Inches(0.11)
    banner.width, banner.height = Inches(11.50), Inches(0.59)
    tf = banner.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.word_wrap = True
    while len(tf.paragraphs) > 1:
        tf._txBody.remove(tf.paragraphs[-1]._p)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    for child in list(p._p):
        if not child.tag.endswith("pPr"):
            p._p.remove(child)
    r = p.add_run()
    r.text = banner_title
    set_run_font(r, size_pt=max(18, banner_pt), bold=True, color=C_WHITE)

    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.45), Inches(1.04), Inches(0.08), Inches(0.36))
    bar.fill.solid()
    bar.fill.fore_color.rgb = C_BLUE
    bar.line.fill.background()

    sub = slide.shapes.add_textbox(Inches(0.60), Inches(0.94), Inches(12.28), Inches(0.54))
    stf = sub.text_frame
    stf.word_wrap = True
    stf.vertical_anchor = MSO_ANCHOR.MIDDLE
    sp = stf.paragraphs[0]
    sp.alignment = PP_ALIGN.LEFT
    append_rich_text(sp, subtitle_text, size_pt=max(18, subtitle_pt), color=C_BLUE, bold=True)

    rule = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.45), Inches(1.56), Inches(12.43), Inches(0.02))
    rule.fill.solid()
    rule.fill.fore_color.rgb = C_BLUE
    rule.line.fill.background()

    if page_num is not None:
        pg = slide.shapes.add_textbox(Inches(12.35), Inches(7.05), Inches(0.75), Inches(0.36))
        pg.name = f"Slide Number {page_num}"
        ptf = pg.text_frame
        ptf.word_wrap = False
        ptf.vertical_anchor = MSO_ANCHOR.MIDDLE
        pp = ptf.paragraphs[0]
        pp.alignment = PP_ALIGN.RIGHT
        pr = pp.add_run()
        pr.text = str(page_num)
        set_run_font(pr, size_pt=18, bold=True, color=C_BLUE)

    return banner, sub


def format_cover_slide(slide, title_lines, meta_lines, title_pt=36, meta_pt=22, text_color=C_WHITE):
    """格式化天大标准封面页（标题幻灯片版式，深蓝背景配纯白粗体字）。"""
    if slide.has_notes_slide:
        slide.notes_slide.notes_text_frame.text = ""
    _remove_alternate_content(slide)
    if isinstance(title_lines, str):
        title_lines = title_lines.split("\n")
    if isinstance(meta_lines, str):
        meta_lines = meta_lines.split("\n")

    text_shapes = sorted([sh for sh in slide.shapes if sh.has_text_frame], key=lambda s: s.top)
    title_sh = next((s for s in text_shapes if s.top.inches < 3.0), None)
    meta_sh = next((s for s in text_shapes if s is not title_sh and 3.0 <= s.top.inches < 6.5), None)
    if title_sh is None:
        title_sh = slide.shapes.add_textbox(Inches(0.91), Inches(1.23), Inches(11.50), Inches(2.20))
    if meta_sh is None:
        meta_sh = slide.shapes.add_textbox(Inches(1.71), Inches(3.94), Inches(9.90), Inches(1.40))

    keep = {title_sh._element, meta_sh._element}
    for sh in list(slide.shapes):
        if sh._element not in keep:
            sh._element.getparent().remove(sh._element)

    for sh, lines, pt in ((title_sh, title_lines, title_pt), (meta_sh, meta_lines, meta_pt)):
        tf = sh.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        while len(tf.paragraphs) > len(lines):
            tf._txBody.remove(tf.paragraphs[-1]._p)
        for idx, line in enumerate(lines):
            p = tf.paragraphs[idx] if idx < len(tf.paragraphs) else tf.add_paragraph()
            update_paragraph_text(p, line, size_pt=max(18, pt), bold=True, color=text_color, align=PP_ALIGN.CENTER)
    return title_sh, meta_sh


def format_toc_slide(slide, toc_items, active_idx=0, page_num=None):
    """格式化天大标准目录/章节过渡页（1_仅标题版式）：当前章节红字高亮，其余深蓝，消除 a:fld 重复页码。"""
    if slide.has_notes_slide:
        slide.notes_slide.notes_text_frame.text = ""
    _remove_alternate_content(slide)

    toc_sh, pg_sh = None, None
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        if "Placeholder" in sh.name or sh.top.inches > 6.5:
            pg_sh = sh
        elif toc_sh is None:
            toc_sh = sh

    if toc_sh is None:
        toc_sh = slide.shapes.add_textbox(Inches(4.80), Inches(2.22), Inches(6.80), Inches(4.60))
    else:
        toc_sh.left, toc_sh.top, toc_sh.width, toc_sh.height = Inches(4.80), Inches(2.22), Inches(6.80), Inches(4.60)

    keep = {toc_sh._element}
    if pg_sh is not None and page_num is not None:
        keep.add(pg_sh._element)
    for sh in list(slide.shapes):
        if sh._element not in keep:
            sh._element.getparent().remove(sh._element)

    tf = toc_sh.text_frame
    tf.word_wrap = True
    n = len(toc_items)
    while len(tf.paragraphs) < n:
        tf.add_paragraph()
    while len(tf.paragraphs) > n:
        tf._txBody.remove(tf.paragraphs[-1]._p)
    for idx, title in enumerate(toc_items):
        p = tf.paragraphs[idx]
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(18)
        for child in list(p._p):
            if not child.tag.endswith("pPr"):
                p._p.remove(child)
        is_act = (idx == active_idx)
        prefix = "▸ " if is_act else "   "
        r = p.add_run()
        r.text = prefix + title
        set_run_font(r, size_pt=28 if is_act else 25, bold=True, color=C_RED if is_act else C_BLUE)

    if page_num is not None:
        if pg_sh is None:
            pg_sh = slide.shapes.add_textbox(Inches(12.35), Inches(7.05), Inches(0.75), Inches(0.36))
        else:
            pg_sh.left, pg_sh.top, pg_sh.width, pg_sh.height = Inches(12.35), Inches(7.05), Inches(0.75), Inches(0.36)
        pg_sh.name = f"Slide Number {page_num}"
        ptf = pg_sh.text_frame
        ptf.word_wrap = False
        ptf.vertical_anchor = MSO_ANCHOR.MIDDLE
        while len(ptf.paragraphs) > 1:
            ptf._txBody.remove(ptf.paragraphs[-1]._p)
        pp = ptf.paragraphs[0]
        pp.alignment = PP_ALIGN.RIGHT
        for child in list(pp._p):
            if not child.tag.endswith("pPr"):
                pp._p.remove(child)
        pr = pp.add_run()
        pr.text = str(page_num)
        set_run_font(pr, size_pt=18, bold=True, color=C_BLUE)

    return toc_sh


def format_thanks_slide(slide, lines=("恳请各位老师批评指正", "谢谢！"), size_pt=40, text_color=C_WHITE):
    """格式化天大标准致谢页（标题幻灯片版式）。"""
    if slide.has_notes_slide:
        slide.notes_slide.notes_text_frame.text = ""
    _remove_alternate_content(slide)
    if isinstance(lines, str):
        lines = lines.split("\n")

    main_sh = next((sh for sh in slide.shapes if sh.has_text_frame), None)
    if main_sh is None:
        main_sh = slide.shapes.add_textbox(Inches(1.71), Inches(2.20), Inches(9.90), Inches(2.60))
    keep = {main_sh._element}
    for sh in list(slide.shapes):
        if sh._element not in keep:
            sh._element.getparent().remove(sh._element)

    tf = main_sh.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    while len(tf.paragraphs) > len(lines):
        tf._txBody.remove(tf.paragraphs[-1]._p)
    for idx, line in enumerate(lines):
        p = tf.paragraphs[idx] if idx < len(tf.paragraphs) else tf.add_paragraph()
        update_paragraph_text(p, line, size_pt=max(18, size_pt), bold=True, color=text_color, align=PP_ALIGN.CENTER)
    return main_sh


def rearrange_slides(prs, src_indices_0based):
    """按 0-based 索引列表重排并裁剪模板页，同时安全释放被删除页的 rId 关系。"""
    sldIdLst = prs.slides._sldIdLst
    orig_elems = list(sldIdLst)
    keep_set = set(src_indices_0based)
    if len(keep_set) != len(src_indices_0based):
        raise ValueError("src_indices_0based must not contain duplicate indices")

    for i, el in enumerate(orig_elems):
        if i not in keep_set:
            rId = el.get(qn("r:id"))
            sldIdLst.remove(el)
            if rId:
                prs.part.drop_rel(rId)

    ordered = [orig_elems[i] for i in src_indices_0based]
    for el in list(sldIdLst):
        sldIdLst.remove(el)
    for el in ordered:
        sldIdLst.append(el)


def clear_all_slides(prs):
    """清空模板中的示例页（保留母版与版式），便于随后用 new_*_slide() 按需新增任意页数。"""
    rearrange_slides(prs, [])


_TJU_LAYOUT_NAMES = {
    "cover": ("标题幻灯片",),
    "thanks": ("标题幻灯片",),
    "content": ("1_自定义版式",),
    "toc": ("1_仅标题",),
    "blank": ("空白",),
}
_TJU_LAYOUT_FALLBACK_IDX = {"cover": 0, "thanks": 0, "content": 1, "toc": 2, "blank": 3}


def get_tju_layout(prs, kind="content"):
    """按语义类型 ('cover'/'content'/'toc'/'blank'/'thanks') 查找天大母版版式。
    同时兼容新版单母版 tju_template.pptx 与旧版双母版课题组模板。"""
    if kind not in _TJU_LAYOUT_NAMES:
        raise ValueError(f"unknown layout kind {kind!r}; use {list(_TJU_LAYOUT_NAMES)}")
    names = _TJU_LAYOUT_NAMES[kind]
    for master in prs.slide_masters:
        for lay in master.slide_layouts:
            if lay.name in names:
                return lay
    idx = _TJU_LAYOUT_FALLBACK_IDX[kind]
    if idx < len(prs.slide_layouts):
        return prs.slide_layouts[idx]
    raise LookupError(f"Could not find TJU layout for kind={kind!r}")


def new_cover_slide(prs, title_lines, meta_lines, title_pt=36, meta_pt=22, text_color=C_WHITE):
    """基于天大母版新增并格式化封面页。"""
    slide = prs.slides.add_slide(get_tju_layout(prs, "cover"))
    format_cover_slide(slide, title_lines, meta_lines, title_pt=title_pt, meta_pt=meta_pt, text_color=text_color)
    return slide


def new_toc_slide(prs, toc_items, active_idx=0, page_num=None):
    """基于天大母版新增并格式化目录/章节过渡页。"""
    slide = prs.slides.add_slide(get_tju_layout(prs, "toc"))
    format_toc_slide(slide, toc_items, active_idx=active_idx, page_num=page_num)
    return slide


def new_content_slide(prs, banner_title, subtitle_text, page_num=None, subtitle_pt=20, banner_pt=28):
    """基于天大母版新增并格式化标准内容页（顶部蓝条 + 副标题竖条/细线 + 右下角页码）。"""
    slide = prs.slides.add_slide(get_tju_layout(prs, "content"))
    setup_tju_content_slide(
        slide, banner_title, subtitle_text, page_num=page_num,
        subtitle_pt=subtitle_pt, banner_pt=banner_pt,
    )
    return slide


def new_thanks_slide(prs, lines=("恳请各位老师批评指正", "谢谢！"), size_pt=40, text_color=C_WHITE):
    """基于天大母版新增并格式化致谢页。"""
    slide = prs.slides.add_slide(get_tju_layout(prs, "thanks"))
    format_thanks_slide(slide, lines=lines, size_pt=size_pt, text_color=text_color)
    return slide


def save_staged_pptx(prs, target_path, staged_path=None, fallback_suffix="_staged"):
    """先写入暂存文件再复制到目标路径，避免目标 PPTX 正被 PowerPoint 打开时直接报错中断。"""
    target_abs = os.path.abspath(target_path)
    os.makedirs(os.path.dirname(target_abs), exist_ok=True)
    auto_tmp = staged_path is None
    if auto_tmp:
        fd, staged_path = tempfile.mkstemp(prefix="tju_ppt_", suffix=".pptx")
        os.close(fd)
    else:
        os.makedirs(os.path.dirname(os.path.abspath(staged_path)), exist_ok=True)
    prs.save(staged_path)
    try:
        shutil.copyfile(staged_path, target_abs)
        return target_abs
    except PermissionError:
        stem, ext = os.path.splitext(target_abs)
        alt_path = f"{stem}{fallback_suffix}{ext}"
        shutil.copyfile(staged_path, alt_path)
        return alt_path
    finally:
        if auto_tmp and os.path.exists(staged_path):
            try:
                os.remove(staged_path)
            except OSError:
                pass


def add_dashed_box(slide, left, top, width, height, line_color=C_BLACK, weight_pt=1.5):
    """添加双栏虚线外框容器 (sysDot)"""
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    sp.fill.background()
    sp.line.color.rgb = line_color
    sp.line.width = Pt(weight_pt)
    ln = sp.line._get_or_add_ln()
    prstDash = etree.SubElement(ln, qn("a:prstDash"))
    prstDash.set("val", "sysDot")
    return sp


def add_solid_box(
    slide, left, top, width, height, text_lines,
    fill_color=None, line_color=C_BLUE, text_color=C_BLACK,
    size_pt=18, align=PP_ALIGN.CENTER, line_width=1.5,
    margin_lr=0.07, margin_tb=0.04, line_spacing=1.08,
):
    """添加流程图或结构化卡片框（支持单段多色分段列表与 _{sub}/^{sup} 行内上下标）"""
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
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
    tf = sp.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(margin_lr)
    tf.margin_right = Inches(margin_lr)
    tf.margin_top = Inches(margin_tb)
    tf.margin_bottom = Inches(margin_tb)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    if isinstance(text_lines, str):
        text_lines = [text_lines]
    for i, item in enumerate(text_lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        if isinstance(item, list):
            for seg in item:
                append_rich_text(p, seg[0], size_pt=max(18, seg[2]), color=seg[1], bold=True)
        elif isinstance(item, tuple):
            txt, col, pt = item[0], item[1], item[2]
            append_rich_text(p, txt, size_pt=max(18, pt), color=col, bold=True)
        else:
            append_rich_text(p, item, size_pt=max(18, size_pt), color=text_color, bold=True)
    return sp


def add_joined_card(
    slide, left, top, width, header_h, body_h,
    header_text, body_lines,
    header_bg=C_BLUE, body_bg=C_BLUE_LIGHT, border_color=C_BLUE,
    header_pt=19, body_pt=18, align=PP_ALIGN.LEFT,
):
    """添加两段连体卡片（深色标题条与正文框无缝相接，零声明重叠）。"""
    hdr = add_solid_box(
        slide, left, top, width, header_h, [header_text],
        fill_color=header_bg, line_color=border_color, text_color=C_WHITE,
        size_pt=header_pt, align=PP_ALIGN.LEFT, margin_lr=0.10, margin_tb=0.03,
    )
    body = add_solid_box(
        slide, left, top + header_h, width, body_h, body_lines,
        fill_color=body_bg, line_color=border_color, text_color=C_BLACK,
        size_pt=body_pt, align=align, margin_lr=0.10, margin_tb=0.05,
    )
    return hdr, body


def _set_cell_borders(cell, hex_color="00337F", width_pt=0.75):
    """显式写入四色边框，避免 PowerPoint 默认表格样式引入主题色 schemeClr。"""
    tcPr = cell._tc.get_or_add_tcPr()
    w_emu = str(int(round(width_pt * 12700)))
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        old = tcPr.find(qn(tag))
        if old is not None:
            tcPr.remove(old)
        ln = etree.SubElement(tcPr, qn(tag), w=w_emu, cap="flat", cmpd="sng", algn="ctr")
        sf = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr"), val=hex_color)


def add_clean_tju_table(
    slide, left, top, width, height, col_widths, rows_data,
    col_aligns=None, header_pt=19, body_pt=18, badge_col0=False,
):
    """构建严格四色合规表格（蓝底白字表头 + 浅蓝/纯白交替行；可选首列深蓝徽标）。"""
    n_rows = len(rows_data)
    n_cols = len(rows_data[0])
    shape = slide.shapes.add_table(n_rows, n_cols, Inches(left), Inches(top), Inches(width), Inches(height))
    tbl = shape.table
    # 关闭默认表格主题样式引用，防止引入非四色 schemeClr
    tblPr = tbl._tbl.tblPr
    for child in list(tblPr):
        if child.tag.endswith("tableStyleId"):
            tblPr.remove(child)
    for ci, cw in enumerate(col_widths):
        tbl.columns[ci].width = Inches(cw)

    row_h = height / n_rows
    for ri, row in enumerate(rows_data):
        tbl.rows[ri].height = Inches(row_h)
        is_hdr = (ri == 0)
        for ci, val in enumerate(row):
            is_badge = badge_col0 and (not is_hdr) and (ci == 0)
            bg = C_BLUE if (is_hdr or is_badge) else (C_BLUE_LIGHT if ri % 2 == 1 else C_WHITE)
            fg = C_WHITE if (is_hdr or is_badge) else C_BLACK
            pt = max(18, header_pt if is_hdr else body_pt)
            cell = tbl.cell(ri, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            _set_cell_borders(cell, hex_color="FFFFFF" if is_hdr else "00337F", width_pt=0.75)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Inches(0.08)
            cell.margin_right = Inches(0.08)
            cell.margin_top = Inches(0.04)
            cell.margin_bottom = Inches(0.04)
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
            if isinstance(val, list):
                for seg in val:
                    seg_txt, seg_col = seg[0], seg[1]
                    seg_pt = max(18, seg[2]) if len(seg) >= 3 else pt
                    append_rich_text(p, seg_txt, size_pt=seg_pt, color=seg_col, bold=True)
            else:
                append_rich_text(p, str(val), size_pt=pt, color=fg, bold=True)
    return shape


def add_arrow(slide, left, top, width, height, direction="right", color=C_BLUE):
    """添加实心方向箭头"""
    shape_type = {
        "right": MSO_SHAPE.RIGHT_ARROW,
        "down": MSO_SHAPE.DOWN_ARROW,
        "up": MSO_SHAPE.UP_ARROW,
        "left": MSO_SHAPE.LEFT_ARROW,
    }[direction]
    sp = slide.shapes.add_shape(shape_type, Inches(left), Inches(top), Inches(width), Inches(height))
    sp.fill.solid()
    sp.fill.fore_color.rgb = color
    sp.line.fill.background()
    return sp


def add_text_block(
    slide, left, top, width, height, lines,
    default_pt=18, default_color=C_BLACK, align=PP_ALIGN.LEFT, line_spacing=1.10,
):
    """添加多行富文本框，支持多色分段列表与 _{sub}/^{sup} 行内上下标"""
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.04)
    tf.margin_right = Inches(0.04)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    for i, item in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        if isinstance(item, list):
            if len(item) == 1 and isinstance(item[0], list):
                item = item[0]
            p.space_after = Pt(2)
            for seg in item:
                append_rich_text(p, seg[0], size_pt=max(18, seg[2]), color=seg[1], bold=True)
            continue
        if isinstance(item, tuple):
            if len(item) == 4:
                txt, col, pt, bld = item
            else:
                txt, col, pt = item[0], item[1], item[2]
                bld = True
        else:
            txt, col, pt, bld = item, default_color, default_pt, True
        if col == C_BLUE and pt >= 19:
            if i > 0:
                p.space_before = Pt(4)
            p.space_after = Pt(3)
        else:
            p.space_after = Pt(2)
        append_rich_text(p, txt, size_pt=max(18, pt), color=col, bold=bld)
    return tb


def add_image_fit(slide, img_path, left, top, box_w, box_h, border_color=C_BLUE, border_pt=0.75):
    """等比缩放居中嵌入图片，并直接在 Picture 形状上绘制细边框（零背景形状重叠）"""
    with Image.open(img_path) as im:
        iw, ih = im.size
    ratio = iw / ih
    if box_w / box_h > ratio:
        draw_h = box_h
        draw_w = draw_h * ratio
    else:
        draw_w = box_w
        draw_h = draw_w / ratio
    x = left + (box_w - draw_w) / 2.0
    y = top + (box_h - draw_h) / 2.0
    pic = slide.shapes.add_picture(img_path, Inches(x), Inches(y), width=Inches(draw_w), height=Inches(draw_h))
    if border_color is not None:
        pic.line.color.rgb = border_color
        pic.line.width = Pt(border_pt)
    return pic


def add_omml_equation(
    slide, left, top, width, height, omath_children_xml,
    jc="left", fill_color=None, line_color=None, line_width=1.25,
):
    """插入 PowerPoint 原生可编辑数学公式 (<a14:m>)，支持单形状高亮背景与边框"""
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
    pad_lr = 0.06 if line_color else 0.02
    tf.margin_left = Inches(pad_lr)
    tf.margin_right = Inches(pad_lr)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    p = tf.paragraphs[0]
    omml_str = (
        '<a14:m xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">'
        '<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<m:oMathParaPr><m:jc m:val="{jc}"/></m:oMathParaPr>'
        f"<m:oMath>{omath_children_xml}</m:oMath>"
        "</m:oMathPara>"
        "</a14:m>"
    )
    p._p.append(etree.fromstring(omml_str))
    return tb


# OMML 辅助构造函数
def m_r(text, italic=True, sz=1800, color="000000"):
    esc_text = html.escape(text)
    mrpr = "" if italic else '<m:rPr><m:sty m:val="p"/></m:rPr>'
    i_attr = ' i="1"' if italic else ' i="0"'
    return (
        f"<m:r>{mrpr}"
        f'<a:rPr sz="{max(1800, sz)}" b="1"{i_attr}>'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'<a:latin typeface="Times New Roman"/>'
        f'<a:ea typeface="微软雅黑"/>'
        f'<a:cs typeface="Times New Roman"/>'
        f'<a:sym typeface="Times New Roman"/>'
        f"</a:rPr><m:t>{esc_text}</m:t></m:r>"
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
    """公式重音符（默认组合扬抑符 U+0302，严禁用 ASCII '^' 以免下沉压住大写字母）"""
    return f'<m:acc><m:accPr><m:chr m:val="{html.escape(ch)}"/></m:accPr><m:e>{inner_xml}</m:e></m:acc>'


def enforce_min_18pt_all_slides(prs):
    """全量保存前合规守门：确保所有文本、表格单元格与 OMML 公式 >=18pt、粗体、四槽位字体，并完整保留上下标 baseline"""
    for slide in prs.slides:
        if slide.has_notes_slide:
            slide.notes_slide.notes_text_frame.text = ""
        for sh in slide.shapes:
            if sh.has_text_frame:
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        if r.font.size is None or r.font.size.pt < 18:
                            r.font.size = Pt(18)
                        rPr = r._r.get_or_add_rPr()
                        rPr.set("b", "1")
                        _ensure_four_font_slots(rPr)
            if sh.has_table:
                for row in sh.table.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            for r in p.runs:
                                rPr = r._r.get_or_add_rPr()
                                cur_bl = rPr.get("baseline")
                                cur_it = rPr.get("i") == "1"
                                cur_pt = r.font.size.pt if r.font.size is not None else 18
                                cur_col = r.font.color.rgb if (r.font.color and r.font.color.type == 1) else C_BLACK
                                set_run_font(
                                    r,
                                    size_pt=max(18, cur_pt),
                                    bold=True,
                                    color=cur_col,
                                    italic=cur_it,
                                    baseline=cur_bl,
                                )
        # 覆盖全页 OMML <m:r>（含被 PowerPoint 包入 mc:AlternateContent/Choice 的公式，跳过 Fallback）
        for el in slide._element.iter():
            if not (el.tag.endswith("}r") and "math" in el.tag):
                continue
            if _in_mc_fallback(el):
                continue
            arpr = el.find(qn("a:rPr"))
            if arpr is None:
                mt = el.find("{http://schemas.openxmlformats.org/officeDocument/2006/math}t")
                arpr = etree.Element(qn("a:rPr"))
                if mt is not None:
                    el.insert(list(el).index(mt), arpr)
                else:
                    el.append(arpr)
            sz_val = arpr.get("sz")
            if sz_val is None or int(sz_val) < 1800:
                arpr.set("sz", "1800")
            arpr.set("b", "1")
            _ensure_four_font_slots(arpr)
