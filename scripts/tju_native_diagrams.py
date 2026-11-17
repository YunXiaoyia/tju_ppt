#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TJU-PPT 原生可编辑流程图组件 (tju_native_diagrams.py)

全部由 python-pptx 原生形状 + 直线连接器 (p:cxnSp) 组成，PowerPoint 中可直接双击改字：
- add_connector / connect_boxes : 带三角箭头的直线/折线连接器（零面积包围盒，不触发重叠审计）
- add_flow_pipeline             : N 步横向/纵向流程（输入浅蓝 → 核心深蓝 → 输出红框）
- add_feedback_loop             : 前向链 + 底部回环反馈箭头（闭环控制 / 迭代优化）
- add_layered_architecture      : 分层系统架构（左侧深蓝层名 + 右侧浅蓝模块行 + 层间箭头）
- text_width_in / check_box_fit : 18 pt 字宽预算，生成阶段提前发现溢出

设计约束（与 SKILL.md 一致）：
- 只用 #00337F / #EAF1FB / #C00000 / #000000 / #FFFFFF；所有文字 >= 18 pt，经 append_rich_text 写入四槽位字体。
- 不堆叠形状：模块之间只用连接器衔接，没有“背景框 + 文字框”嵌套，满足零声明包围盒重叠。
"""

import math
import re
import warnings

from lxml import etree
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from ppt_builder_helpers import (
    C_BLACK, C_BLUE, C_BLUE_LIGHT, C_RED, C_WHITE,
    add_solid_box, add_text_block,
)

_MARKUP = re.compile(r"[_^]\{([^{}]*)\}")

# 角色 → (填充, 边框, 文字颜色, 边框粗细 pt)
ROLE_STYLES = {
    "input": (C_BLUE_LIGHT, C_BLUE, C_BLACK, 1.5),
    "core": (C_BLUE, C_BLUE, C_WHITE, 1.5),
    "output": (C_WHITE, C_RED, C_BLACK, 2.25),
    "emphasis": (C_RED, C_RED, C_WHITE, 1.5),
    "plain": (C_WHITE, C_BLUE, C_BLACK, 1.5),
}


# --------------------------------------------------------------------------- #
# 字宽预算
# --------------------------------------------------------------------------- #
def _char_em(ch):
    o = ord(ch)
    if (0x2E80 <= o <= 0x9FFF) or (0xF900 <= o <= 0xFAFF) or (0xFF00 <= o <= 0xFFEF) or (0x3000 <= o <= 0x303F):
        return 1.0
    if ch == " ":
        return 0.25
    if ch.isupper():
        return 0.70
    if ch.isdigit() or ch.islower():
        return 0.52
    return 0.38


def text_width_in(text, pt=18):
    """估算一行文字在 pt 字号下的宽度（英寸）。_{}/^{} 标记按正文字宽计（与 append_rich_text 一致）。"""
    plain = _MARKUP.sub(r"\1", text)
    return sum(_char_em(c) for c in plain) * pt / 72.0


def _item_to_text(item):
    if isinstance(item, str):
        return item
    if isinstance(item, tuple):
        return item[0]
    if isinstance(item, list):
        return "".join(seg[0] for seg in item)
    return str(item)


def check_box_fit(text_lines, width_in, height_in, pt=18, margin_lr=0.07, margin_tb=0.04,
                  line_spacing=1.08, label="box"):
    """返回 (是否放得下, 需要行数, 可容纳行数)；放不下时发出 warnings 提示，便于生成阶段修正。"""
    if isinstance(text_lines, str):
        text_lines = [text_lines]
    usable_w = max(0.05, width_in - 2 * margin_lr - 0.04)
    need = 0
    for item in text_lines:
        need += max(1, math.ceil(text_width_in(_item_to_text(item), pt) / usable_w))
    line_h = pt * 1.2 * line_spacing / 72.0
    cap = max(0, int((height_in - 2 * margin_tb) // line_h))
    ok = need <= cap
    if not ok:
        warnings.warn(
            f"[tju-ppt] {label}: 需要 {need} 行但 {width_in:.2f}×{height_in:.2f} in 只能容纳 {cap} 行 "
            f"（{pt} pt）。请精简文字、加宽盒子或增加高度。",
            stacklevel=3,
        )
    return ok, need, cap


# --------------------------------------------------------------------------- #
# 连接器
# --------------------------------------------------------------------------- #
def add_connector(slide, x1, y1, x2, y2, color=C_BLUE, width_pt=2.0, arrow_end=True,
                  arrow_start=False, dashed=False):
    """添加直线连接器（英寸坐标），可选首/尾三角箭头与虚线。"""
    cx = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    cx.line.color.rgb = color
    cx.line.width = Pt(width_pt)
    ln = cx.line._get_or_add_ln()
    if dashed:
        dash = etree.SubElement(ln, qn("a:prstDash"))
        dash.set("val", "dash")
    if arrow_start:
        he = etree.SubElement(ln, qn("a:headEnd"))
        he.set("type", "triangle"); he.set("w", "med"); he.set("len", "med")
    if arrow_end:
        te = etree.SubElement(ln, qn("a:tailEnd"))
        te.set("type", "triangle"); te.set("w", "med"); te.set("len", "med")
    return cx


def add_polyline_arrow(slide, points, color=C_BLUE, width_pt=2.0, dashed=False, arrow_start=False):
    """用多段直线连接器拼成正交折线，仅最后一段带箭头。points: [(x, y), ...] 英寸。"""
    segs = []
    n = len(points) - 1
    for i in range(n):
        (x1, y1), (x2, y2) = points[i], points[i + 1]
        segs.append(add_connector(
            slide, x1, y1, x2, y2, color=color, width_pt=width_pt,
            arrow_end=(i == n - 1), arrow_start=(arrow_start and i == 0), dashed=dashed,
        ))
    return segs


def box_anchor(geom, side):
    """geom=(left, top, w, h)；side ∈ left/right/top/bottom/center → (x, y)。"""
    l, t, w, h = geom
    return {
        "left": (l, t + h / 2), "right": (l + w, t + h / 2),
        "top": (l + w / 2, t), "bottom": (l + w / 2, t + h),
        "center": (l + w / 2, t + h / 2),
    }[side]


def connect_boxes(slide, g1, g2, side1=None, side2=None, color=C_BLUE, width_pt=2.0, dashed=False):
    """连接两个盒子：同行/同列自动直线，否则走 横-竖-横 或 竖-横-竖 正交折线。"""
    l1, t1, w1, h1 = g1
    l2, t2, w2, h2 = g2
    c1 = (l1 + w1 / 2, t1 + h1 / 2)
    c2 = (l2 + w2 / 2, t2 + h2 / 2)
    dx, dy = c2[0] - c1[0], c2[1] - c1[1]
    horizontal = abs(dx) >= abs(dy)
    if side1 is None:
        side1 = ("right" if dx > 0 else "left") if horizontal else ("bottom" if dy > 0 else "top")
    if side2 is None:
        side2 = ("left" if dx > 0 else "right") if horizontal else ("top" if dy > 0 else "bottom")
    p1, p2 = box_anchor(g1, side1), box_anchor(g2, side2)
    y_overlap = min(t1 + h1, t2 + h2) - max(t1, t2)
    x_overlap = min(l1 + w1, l2 + w2) - max(l1, l2)
    if side1 in ("left", "right") and side2 in ("left", "right") and y_overlap > 0.1:
        y = max(t1, t2) + y_overlap / 2
        return [add_connector(slide, p1[0], y, p2[0], y, color, width_pt, dashed=dashed)]
    if side1 in ("top", "bottom") and side2 in ("top", "bottom") and x_overlap > 0.1:
        x = max(l1, l2) + x_overlap / 2
        return [add_connector(slide, x, p1[1], x, p2[1], color, width_pt, dashed=dashed)]
    if side1 in ("left", "right") and side2 in ("left", "right"):
        mx = (p1[0] + p2[0]) / 2
        pts = [p1, (mx, p1[1]), (mx, p2[1]), p2]
    elif side1 in ("top", "bottom") and side2 in ("top", "bottom"):
        my = (p1[1] + p2[1]) / 2
        pts = [p1, (p1[0], my), (p2[0], my), p2]
    elif side1 in ("left", "right"):
        pts = [p1, (p2[0], p1[1]), p2]
    else:
        pts = [p1, (p1[0], p2[1]), p2]
    return add_polyline_arrow(slide, pts, color=color, width_pt=width_pt, dashed=dashed)


# --------------------------------------------------------------------------- #
# 盒子
# --------------------------------------------------------------------------- #
def add_role_box(slide, left, top, width, height, text_lines, role="plain", size_pt=18,
                 align=PP_ALIGN.CENTER, label=None):
    """按语义角色 (input/core/output/emphasis/plain) 添加盒子，并做 18 pt 字宽预算检查。"""
    fill, line, txt, lw = ROLE_STYLES[role]
    check_box_fit(text_lines, width, height, pt=size_pt, label=label or f"{role} box")
    return add_solid_box(
        slide, left, top, width, height, text_lines,
        fill_color=fill, line_color=line, text_color=txt, size_pt=size_pt,
        align=align, line_width=lw,
    )


def _default_roles(n):
    if n == 1:
        return ["core"]
    return ["input"] + ["core"] * (n - 2) + ["output"]


# --------------------------------------------------------------------------- #
# 组合图元
# --------------------------------------------------------------------------- #
def add_flow_pipeline(slide, left, top, width, height, steps, roles=None, direction="h",
                      gap=0.45, size_pt=18, arrow_color=C_BLUE):
    """
    N 步流程图。steps: 每步为 str 或多行 list（支持 _{sub}/^{sup} 与多色分段）。
    roles 默认：首步 input、中间 core、末步 output。direction: "h" 横向 / "v" 纵向。
    返回 {"shapes": [...], "geoms": [(l, t, w, h), ...], "connectors": [...]}。
    """
    n = len(steps)
    roles = roles or _default_roles(n)
    shapes, geoms, conns = [], [], []
    if direction == "h":
        bw = (width - gap * (n - 1)) / n
        for i, (st, rl) in enumerate(zip(steps, roles)):
            g = (left + i * (bw + gap), top, bw, height)
            shapes.append(add_role_box(slide, *g, st, role=rl, size_pt=size_pt, label=f"pipeline step {i + 1}"))
            geoms.append(g)
    else:
        bh = (height - gap * (n - 1)) / n
        for i, (st, rl) in enumerate(zip(steps, roles)):
            g = (left, top + i * (bh + gap), width, bh)
            shapes.append(add_role_box(slide, *g, st, role=rl, size_pt=size_pt, label=f"pipeline step {i + 1}"))
            geoms.append(g)
    for a, b in zip(geoms[:-1], geoms[1:]):
        conns += connect_boxes(slide, a, b, color=arrow_color)
    return {"shapes": shapes, "geoms": geoms, "connectors": conns}


def add_feedback_loop(slide, left, top, width, height, steps, feedback_label, roles=None,
                      gap=0.45, box_h=None, size_pt=18, loop_color=C_RED):
    """
    闭环流程：上方横向前向链，下方从末步底部 → 回到首步底部的红色反馈折线，并在折线下方标注。
    height 需同时容纳盒子 + 回环通道（约 0.85 in）。
    """
    box_h = box_h or max(0.62, height - 0.95)
    res = add_flow_pipeline(slide, left, top, width, box_h, steps, roles=roles, gap=gap, size_pt=size_pt)
    first, last = res["geoms"][0], res["geoms"][-1]
    y_loop = top + box_h + 0.30
    x_end = last[0] + last[2] / 2
    x_start = first[0] + first[2] / 2
    pts = [(x_end, top + box_h), (x_end, y_loop), (x_start, y_loop), (x_start, top + box_h)]
    res["connectors"] += add_polyline_arrow(slide, pts, color=loop_color, width_pt=2.0, dashed=True)
    lbl_w = min(width - 0.4, max(2.0, text_width_in(feedback_label, size_pt) + 0.3))
    lbl_left = (x_start + x_end) / 2 - lbl_w / 2
    res["label"] = add_text_block(
        slide, lbl_left, y_loop + 0.04, lbl_w, 0.40,
        [(feedback_label, C_RED, size_pt)], align=PP_ALIGN.CENTER,
    )
    return res


def add_layered_architecture(slide, left, top, width, height, layers, label_w=1.55,
                             row_gap=0.32, col_gap=0.14, size_pt=18, arrows=True,
                             label_fill=C_BLUE, flow="up"):
    """
    分层系统架构图（自上而下）。layers: [(层名, [模块1, 模块2, ...]), ...]。
    每层 = 左侧深蓝层名盒 + 右侧等宽浅蓝模块盒；层间以居中箭头连接（无嵌套容器，零重叠）。
    flow: "up"（默认，数据自底层感知流向上层应用）/ "down"（指令下发）/ "both"。
    模块文字以 "!" 开头时渲染为红框重点模块（如本文创新点）。
    """
    n = len(layers)
    row_h = (height - row_gap * (n - 1)) / n
    out = {"labels": [], "modules": [], "connectors": [], "rows": []}
    mod_left = left + label_w + col_gap
    mod_area = width - label_w - col_gap
    for i, (name, mods) in enumerate(layers):
        y = top + i * (row_h + row_gap)
        out["labels"].append(add_role_box(
            slide, left, y, label_w, row_h, name, role="core" if label_fill == C_BLUE else "emphasis",
            size_pt=size_pt, label=f"layer {name}",
        ))
        m = len(mods)
        mw = (mod_area - col_gap * (m - 1)) / m
        row = []
        for j, mod in enumerate(mods):
            role = "input"
            if isinstance(mod, str) and mod.startswith("!"):
                mod, role = mod[1:], "output"
            g = (mod_left + j * (mw + col_gap), y, mw, row_h)
            out["modules"].append(add_role_box(slide, *g, mod, role=role, size_pt=size_pt, label=f"module {mod}"))
            row.append(g)
        out["rows"].append((y, row_h))
    if arrows:
        xc = mod_left + mod_area / 2
        for (y, h), (y2, _) in zip(out["rows"][:-1], out["rows"][1:]):
            out["connectors"].append(add_connector(
                slide, xc, y + h, xc, y2, color=C_BLUE, width_pt=2.25,
                arrow_end=flow in ("down", "both"), arrow_start=flow in ("up", "both"),
            ))
    return out
