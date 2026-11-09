#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成全面优化版开题汇报PPT：
《基于多模态感知的移动机器人自主规划与协同控制研究》
严格遵循用户要求：
1. 字体：中文 微软雅黑 (a:ea)，英文/数字 Times New Roman (a:latin, a:cs)
2. 字号：所有文字字号 >= 18pt（图片内文字除外）
3. 颜色：仅使用 黑(#000000)、白(#FFFFFF)、红(#C00000 强调与结论)、蓝(#00337F/#1F4E79 主色调)
4. 公式：全部采用 PowerPoint 原生可编辑插入公式 (<a14:m><m:oMathPara>...)
5. 图片：嵌入从本地15篇论文PDF提取的真实架构/实验图及移动机器人实验平台图
6. 排版：参照用户提供的天津大学参考图，采用左右双栏虚线框、蓝底白字核心流程框、A/B/C递进框、公式符号注释及前后页承上启下逻辑链
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

# 严格限定的四色规范
C_BLUE = RGBColor(0x00, 0x33, 0x7F)       # 主色深蓝（标题、核心框、主边框）
C_BLUE_LIGHT = RGBColor(0xEA, 0xF1, 0xFB) # 浅蓝填充（次级流程框/表头交替）
C_RED = RGBColor(0xC0, 0x00, 0x00)        # 强调红（核心结论、关键判据、底部承上启下）
C_BLACK = RGBColor(0x00, 0x00, 0x00)      # 正文纯黑
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)      # 反白文字


def set_run_font(run, size_pt=18, bold=True, color=C_BLACK, italic=False):
    """设置run字体：中文微软雅黑，英文Times New Roman，字号>=18pt"""
    if size_pt < 18:
        size_pt = 18
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    run.font.name = "Times New Roman"
    rPr = run._r.get_or_add_rPr()
    rPr.set("b", "1" if bold else "0")
    if italic:
        rPr.set("i", "1")
    for tag, typeface in [("a:latin", "Times New Roman"), ("a:ea", "微软雅黑"), ("a:cs", "Times New Roman"), ("a:sym", "Times New Roman")]:
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", typeface)


def update_paragraph_text(paragraph, new_text, size_pt=18, bold=True, color=None, align=None):
    """替换已有段落文本，强制 >=18pt 且中英分设微软雅黑与Times New Roman"""
    if align is not None:
        paragraph.alignment = align
    if len(paragraph.runs) == 0:
        r = paragraph.add_run()
        r.text = new_text
        set_run_font(r, size_pt=size_pt, bold=bold, color=color or C_BLACK)
        return

    first_run = paragraph.runs[0]
    # 读取原有字号，若小于18pt则提升至18pt
    orig_pt = first_run.font.size.pt if first_run.font.size else size_pt
    final_pt = max(18, size_pt if size_pt else orig_pt)
    first_run.text = new_text
    for r in list(paragraph.runs)[1:]:
        paragraph._p.remove(r._r)
    set_run_font(first_run, size_pt=final_pt, bold=bold, color=color)


def update_shape_paragraphs(shape, lines, size_pt=18, bold=True, color=None):
    tf = shape.text_frame
    for idx, line in enumerate(lines):
        if idx < len(tf.paragraphs):
            update_paragraph_text(tf.paragraphs[idx], line, size_pt=size_pt, bold=bold, color=color)


def update_table_cells(table, cell_data, header_pt=19, body_pt=18):
    """更新表格内容，确保字号全部 >= 18pt，表头蓝底白字，正文黑字"""
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
    """保留顶部标题栏(shapes[0])与副标题(shapes[1])，清除其余旧形状以构建双栏充实版式"""
    for sp in list(slide.shapes)[2:]:
        sp._element.getparent().remove(sp._element)


def add_dashed_box(slide, left, top, width, height, line_color=C_BLACK, weight_pt=1.5):
    """添加参考图风格的虚线外框容器"""
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    sp.fill.background()
    sp.line.color.rgb = line_color
    sp.line.width = Pt(weight_pt)
    ln = sp.line._get_or_add_ln()
    prstDash = etree.SubElement(ln, qn("a:prstDash"))
    prstDash.set("val", "sysDot")
    return sp


def add_solid_box(slide, left, top, width, height, text_lines, fill_color=None, line_color=C_BLUE, text_color=C_BLACK, size_pt=18, align=PP_ALIGN.CENTER, line_width=1.75):
    """添加流程图或步骤框（支持蓝底白字、白底蓝框、白底红框）"""
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
    tf.margin_left = Inches(0.08)
    tf.margin_right = Inches(0.08)
    tf.margin_top = Inches(0.05)
    tf.margin_bottom = Inches(0.05)
    sp.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    if isinstance(text_lines, str):
        text_lines = [text_lines]
    for i, item in enumerate(text_lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if isinstance(item, tuple):
            txt, col, pt = item
        else:
            txt, col, pt = item, text_color, size_pt
        r = p.add_run()
        r.text = txt
        set_run_font(r, size_pt=max(18, pt), bold=True, color=col)
    return sp


def add_arrow(slide, left, top, width, height, direction="right", color=C_BLUE):
    """添加蓝色实心箭头（参考图样式）"""
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


def add_text_block(slide, left, top, width, height, lines, default_pt=18, default_color=C_BLACK, align=PP_ALIGN.LEFT, line_spacing=1.15):
    """添加多行富文本框，支持每行传入 (text, color, size_pt, bold)，并在标题与正文间留出间距"""
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.04)
    tf.margin_right = Inches(0.04)
    tf.margin_top = Inches(0.03)
    tf.margin_bottom = Inches(0.03)
    for i, item in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        if isinstance(item, tuple):
            if len(item) == 4:
                txt, col, pt, bld = item
            else:
                txt, col, pt = item
                bld = True
        else:
            txt, col, pt, bld = item, default_color, default_pt, True
        # 在蓝色小标题与正文之间留出呼吸间距
        if col == C_BLUE and pt >= 19:
            if i > 0:
                p.space_before = Pt(4)
            p.space_after = Pt(4)
        else:
            p.space_after = Pt(2)
        r = p.add_run()
        r.text = txt
        set_run_font(r, size_pt=max(18, pt), bold=bld, color=col)
    return tb


def add_omml_equation(slide, left, top, width, height, omath_children_xml):
    """插入 PowerPoint 原生可编辑数学公式 (<a14:m><m:oMathPara><m:oMath>...)"""
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame
    tf.margin_left = Inches(0.02)
    tf.margin_right = Inches(0.02)
    tf.margin_top = Inches(0.01)
    tf.margin_bottom = Inches(0.01)
    p = tf.paragraphs[0]
    omml_str = f"""<a14:m xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">
      <m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr>
        <m:oMath>
          {omath_children_xml}
        </m:oMath>
      </m:oMathPara>
    </a14:m>"""
    p._p.append(etree.fromstring(omml_str))
    return tb


# OMML 辅助构造函数
import html

def m_r(text, italic=True, sz=1800):
    i_attr = ' i="1"' if italic else ' i="0"'
    esc_text = html.escape(text)
    return f'<m:r><a:rPr sz="{sz}"{i_attr}><a:latin typeface="Times New Roman"/><a:ea typeface="微软雅黑"/></a:rPr><m:t>{esc_text}</m:t></m:r>'

def m_sub(base_xml, sub_xml):
    return f'<m:sSub><m:e>{base_xml}</m:e><m:sub>{sub_xml}</m:sub></m:sSub>'

def m_sup(base_xml, sup_xml):
    return f'<m:sSup><m:e>{base_xml}</m:e><m:sup>{sup_xml}</m:sup></m:sSup>'

def m_subsup(base_xml, sub_xml, sup_xml):
    return f'<m:sSubSup><m:e>{base_xml}</m:e><m:sub>{sub_xml}</m:sub><m:sup>{sup_xml}</m:sup></m:sSubSup>'

def m_frac(num_xml, den_xml):
    return f'<m:f><m:num>{num_xml}</m:num><m:den>{den_xml}</m:den></m:f>'

def m_sqrt(inner_xml):
    return f'<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>{inner_xml}</m:e></m:rad>'


print("Helper module loaded successfully.")
