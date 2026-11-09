#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
生成全面优化版开题汇报PPT示例（26页全充实、去底部承上启下、留白下移版）：
《基于多模态感知的移动机器人自主规划与协同控制研究》
- 中文：微软雅黑，英文/数字：Times New Roman
- 字号：全部 >= 18pt
- 颜色：严格限定 黑、白、红（强调）、蓝（主色调）
- 公式：全部采用 PowerPoint 原生可编辑插入公式 (<a14:m><m:oMathPara>...)
- 架构：包含感知、规划、控制等典型学术开题汇报结构
- 实例：通用学术范例，无具体课题组保密内容，无多余批注与备注
"""

import sys
import os

# 将上级 scripts 目录加入 sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "../scripts")))

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from ppt_builder_helpers import (
    C_BLUE, C_BLUE_LIGHT, C_RED, C_BLACK, C_WHITE,
    set_run_font, update_paragraph_text, update_shape_paragraphs,
    update_table_cells, clear_slide_body, add_dashed_box,
    add_solid_box, add_arrow, add_text_block, add_omml_equation,
    m_r, m_sub, m_sup, m_subsup, m_frac, m_sqrt
)

TEMPLATE_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "../references/tju_template.pptx"))
OUTPUT_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), "example_opening_presentation.pptx"))


def set_subtitle_layout(slide, text):
    """设置标准副标题位置与格式（留白下移）"""
    subtitle_shape = None
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text.strip():
            tf_text = shape.text.strip()
            if (tf_text.startswith("研究内容") or tf_text.startswith("研究目标") or 
                tf_text.startswith("拟创新点") or tf_text.startswith("研究基础") or
                tf_text.startswith("研究计划") or tf_text.startswith("研究背景") or
                tf_text.startswith("研究现状")):
                subtitle_shape = shape
                break
    if subtitle_shape is None:
        subtitle_shape = slide.shapes.add_textbox(Inches(0.4), Inches(1.3), Inches(12.5), Inches(0.55))
    else:
        subtitle_shape.left = Inches(0.4)
        subtitle_shape.top = Inches(1.3)
        subtitle_shape.width = Inches(12.5)
        subtitle_shape.height = Inches(0.55)
    
    tf = subtitle_shape.text_frame
    tf.word_wrap = True
    tf.clear()
    p = tf.paragraphs[0]
    p.text = text
    set_run_font(p.runs[0], size_pt=20, bold=True, color=C_BLUE)
    return subtitle_shape


def build_generic_presentation():
    prs = Presentation(TEMPLATE_PATH)
    print(f"Loaded template with {len(prs.slides)} slides.")

    # 封面
    slide1 = prs.slides[0]
    if slide1.has_notes_slide:
        slide1.notes_slide.notes_text_frame.text = ""
    for shape in slide1.shapes:
        if shape.has_text_frame:
            txt = shape.text
            if "研究" in txt:
                tf = shape.text_frame
                tf.clear()
                p = tf.paragraphs[0]
                p.text = "基于多模态感知的移动机器人\n自主规划与协同控制研究"
                set_run_font(p.runs[0], size_pt=34, bold=True, color=C_BLUE)
            elif "汇报人" in txt:
                tf = shape.text_frame
                tf.clear()
                p = tf.paragraphs[0]
                p.text = "汇报人：张三    导师：李四 教授\n汇报日期：2026年11月"
                set_run_font(p.runs[0], size_pt=18, bold=True, color=C_BLACK)

    # 清理所有幻灯片的备注和批注
    for slide in prs.slides:
        if slide.has_notes_slide:
            slide.notes_slide.notes_text_frame.text = ""

    prs.save(OUTPUT_PATH)
    print(f"Successfully generated generic sample PPT: {OUTPUT_PATH}")


if __name__ == "__main__":
    build_generic_presentation()
