#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TJU-PPT 原生可编辑框图组件 (tju_native_diagrams.py)

凡是框图、流程图、系统架构图、闭环图、模块映射图、时间轴演进图与底部汇总条，一律优先用本模块构建
PPT 原生形状 + 扁平直线连接器 (p:cxnSp)，在 PowerPoint/WPS 中可直接双击改字与微调：
- add_native_panel         : 标题栏 + 下方内区 Region（add_figure_panel 的原生可编辑对等组件）
- add_diagram_group        : 无填充虚线分组框 + 编号徽章 + 标题 + 可选状态胶囊，返回内区 Region（零重叠审计）
- add_role_box / add_pill  : 语义角色圆角/方角盒（input/core/output/emphasis/plain）与重点胶囊
- add_diagram_label        : 箭头旁标注 / 轴标签（可选白底遮罩压线不穿字）
- add_junction_dot         : 分支/汇合节点圆点
- add_connector / add_polyline_arrow / connect_boxes : 扁平无阴影三角箭头连接器
- add_flow_pipeline / add_feedback_loop / add_layered_architecture / add_timeline : 组合图元

字号与配色契约：
- 只用 #00337F / #EAF1FB / #C00000 / #000000 / #FFFFFF；四槽位字体（微软雅黑 + Times New Roman）。
- 本模块创建的形状自动带 `Diagram:` 名称前缀：框内文字默认 16–18 pt，紧凑图元最小 14 pt
  （`MIN_DIAGRAM_PT`，整页架构图建议 ≥ 16 pt），不突破正文 18 pt 下限约束。
"""

import math
import warnings

from lxml import etree
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from tju_components import _one_line_check, add_solid_box, add_text_block, no_theme_effects  # noqa: F401
from tju_layout import Region, as_box  # noqa: F401
from tju_text import write_paragraph  # noqa: F401
from tju_theme import (  # noqa: F401
    C_BLACK, C_BLUE, C_BLUE_LIGHT, C_RED, C_WHITE, DIAGRAM_PREFIX, MIN_DIAGRAM_PT, PANEL_HEADER_H,
    PT_BODY, TONES, text_width_in,
)

# 角色 → (填充, 边框, 文字颜色, 边框粗细 pt)
ROLE_STYLES = {
    "input": (C_BLUE_LIGHT, C_BLUE, C_BLACK, 1.5),
    "core": (C_BLUE, C_BLUE, C_WHITE, 1.5),
    "output": (C_WHITE, C_RED, C_BLACK, 2.0),
    "emphasis": (C_RED, C_RED, C_WHITE, 1.5),
    "plain": (C_WHITE, C_BLUE, C_BLACK, 1.5),
}


def _dname(label):
    if label and label.startswith(DIAGRAM_PREFIX):
        return label
    return f"{DIAGRAM_PREFIX} {label}" if label else f"{DIAGRAM_PREFIX} Node"


def _item_to_text(item):
    if isinstance(item, str):
        return item
    if isinstance(item, tuple):
        return item[0]
    if isinstance(item, list):
        return "".join(_item_to_text(seg) for seg in item)
    return str(item)


def check_box_fit(text_lines, width_in, height_in, pt=18, margin_lr=0.06, margin_tb=0.03,
                  line_spacing=1.08, label="box"):
    """返回 (是否放得下, 需要行数, 可容纳行数)；放不下时发出 warnings 提示，便于生成阶段修正。"""
    if isinstance(text_lines, (str, tuple)):
        text_lines = [text_lines]
    usable_w = max(0.05, width_in - 2 * margin_lr - 0.04)
    need = 0
    for item in text_lines:
        need += max(1, math.ceil(text_width_in(_item_to_text(item), pt) / usable_w))
    line_h = pt * 1.2 * line_spacing / 72.0
    cap = max(0, int((height_in - 2 * margin_tb + 0.02) // line_h))
    ok = need <= cap
    if not ok:
        warnings.warn(
            f"[tju-ppt] {label}: 需要 {need} 行但 {width_in:.2f}×{height_in:.2f} in 只能容纳 {cap} 行 "
            f"（{pt} pt）。请精简文字、加宽盒子或增加高度。",
            stacklevel=3,
        )
    return ok, need, cap


# --------------------------------------------------------------------------- #
# 连接器（扁平无阴影）
# --------------------------------------------------------------------------- #
def add_connector(slide, x1, y1, x2, y2, color=C_BLUE, width_pt=2.0, arrow_end=True,
                  arrow_start=False, dashed=False):
    """添加扁平直线连接器（英寸坐标，无默认阴影），可选首/尾三角箭头与虚线。"""
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
    return no_theme_effects(cx)


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
    l, t, w, h = as_box(geom)
    return {
        "left": (l, t + h / 2), "right": (l + w, t + h / 2),
        "top": (l + w / 2, t), "bottom": (l + w / 2, t + h),
        "center": (l + w / 2, t + h / 2),
    }[side]


def connect_boxes(slide, g1, g2, side1=None, side2=None, color=C_BLUE, width_pt=2.0, dashed=False):
    """连接两个盒子/Region：同行/同列自动直线，否则走 横-竖-横 或 竖-横-竖 正交折线。"""
    l1, t1, w1, h1 = as_box(g1)
    l2, t2, w2, h2 = as_box(g2)
    c1 = (l1 + w1 / 2, t1 + h1 / 2)
    c2 = (l2 + w2 / 2, t2 + h2 / 2)
    dx, dy = c2[0] - c1[0], c2[1] - c1[1]
    horizontal = abs(dx) >= abs(dy)
    if side1 is None:
        side1 = ("right" if dx > 0 else "left") if horizontal else ("bottom" if dy > 0 else "top")
    if side2 is None:
        side2 = ("left" if dx > 0 else "right") if horizontal else ("top" if dy > 0 else "bottom")
    p1, p2 = box_anchor((l1, t1, w1, h1), side1), box_anchor((l2, t2, w2, h2), side2)
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
# 基础图元（盒子 / 胶囊 / 标签 / 节点圆点 / 面板 / 分组框）
# --------------------------------------------------------------------------- #
def add_role_box(slide, left, top, width, height, text_lines, role="plain", size_pt=18,
                 align=PP_ALIGN.CENTER, label=None, rounded=True, dashed=False,
                 margin_lr=0.06, margin_tb=0.03, line_spacing=1.08, corner=0.10):
    """按语义角色 (input/core/output/emphasis/plain) 添加原生框图盒子，默认圆角、≥ 14 pt。"""
    pt = max(MIN_DIAGRAM_PT, size_pt)
    fill, line, txt, lw = ROLE_STYLES[role]
    check_box_fit(text_lines, width, height, pt=pt, margin_lr=margin_lr, margin_tb=margin_tb,
                  line_spacing=line_spacing, label=label or f"{role} box")
    return add_solid_box(
        slide, left, top, width, height, text_lines,
        fill_color=fill, line_color=line, text_color=txt, size_pt=pt,
        align=align, line_width=lw, margin_lr=margin_lr, margin_tb=margin_tb,
        line_spacing=line_spacing, rounded=rounded, dashed=dashed,
        min_pt=MIN_DIAGRAM_PT, name=_dname(label or role), corner=corner,
    )


def add_pill(slide, left, top, width, height, text, color=C_BLUE, pt=16, text_color=C_WHITE, label="Pill"):
    """圆角状态/模型胶囊（如 VLM 引擎、GNN、Orin），用于点亮架构图。"""
    pt = max(MIN_DIAGRAM_PT, pt)
    return add_solid_box(
        slide, left, top, width, height, [text], fill_color=color, line_color=color,
        text_color=text_color, size_pt=pt, align=PP_ALIGN.CENTER, line_width=1.0,
        margin_lr=0.03, margin_tb=0.0, rounded=True, corner=0.38,
        min_pt=MIN_DIAGRAM_PT, name=_dname(label),
    )


def add_diagram_label(slide, left, top, width, height, lines, pt=16, color=C_BLUE,
                      align=PP_ALIGN.CENTER, fill=None, line_spacing=1.05, label="Label"):
    """箭头标注 / 坐标轴刻度 / 分支说明；fill=C_WHITE 时可干净遮挡下方连接线。"""
    pt = max(MIN_DIAGRAM_PT, pt)
    if isinstance(lines, (str, tuple)):
        lines = [lines]
    if fill is not None:
        return add_solid_box(
            slide, left, top, width, height, lines, fill_color=fill, line_color=None,
            text_color=color, size_pt=pt, align=align, margin_lr=0.02, margin_tb=0.0,
            line_spacing=line_spacing, rounded=False, min_pt=MIN_DIAGRAM_PT, name=_dname(label),
        )
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tb.name = _dname(label)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = Inches(0.02)
    tf.margin_top = tf.margin_bottom = Inches(0.0)
    for i, item in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        write_paragraph(p, item, pt, color, min_pt=MIN_DIAGRAM_PT)
    return tb


def add_junction_dot(slide, cx, cy, r=0.06, fill=C_BLUE, border=C_BLUE, border_pt=1.5):
    """连接线分支/汇合处的圆点节点（免阴影，不触发文字检查）。"""
    sp = no_theme_effects(slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(cx - r), Inches(cy - r), Inches(2 * r), Inches(2 * r),
    ))
    sp.name = _dname("Dot")
    sp.fill.solid()
    sp.fill.fore_color.rgb = fill
    if border is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = border
        sp.line.width = Pt(border_pt)
    return sp


def add_native_panel(slide, region, header, tone="blue", header_h=PANEL_HEADER_H,
                     header_pt=PT_BODY, framed=False, pad_x=0.0, pad_top=0.10, pad_bottom=0.0):
    """
    add_figure_panel 的原生可编辑对等组件：在 `region` 顶部绘制齐边标题栏，
    返回下方可供直接绘制原生框图的内区 Region（framed=True 时加无填充虚线外框，零重叠审计）。
    """
    x, y, w, h = as_box(region)
    hdr_fill, _, border = TONES[tone]
    _one_line_check(header, w, header_pt, "native panel header")
    add_solid_box(slide, x, y, w, header_h, [header], fill_color=hdr_fill, line_color=hdr_fill,
                  text_color=C_WHITE, size_pt=header_pt, margin_lr=0.06, margin_tb=0.02)
    body_r = Region(x, y + header_h, w, h - header_h)
    if framed:
        add_solid_box(slide, *body_r.box, None, fill_color=None, line_color=border,
                      line_width=1.3, dashed=True, rounded=False, name=_dname(f"Frame {header}"))
        pad_x = max(pad_x, 0.10)
        pad_top = max(pad_top, 0.10)
        pad_bottom = max(pad_bottom, 0.10)
    return Region(body_r.x + pad_x, body_r.y + pad_top,
                  body_r.w - 2 * pad_x, body_r.h - pad_top - pad_bottom)


def add_diagram_group(slide, left, top, width, height, title=None, badge=None, pill=None,
                      color=C_BLUE, pill_color=C_RED, title_pt=17, pill_w=1.45, corner=0.08):
    """
    无填充虚线圆角分组框（绕过声明包围盒重叠、支持多层模块清晰编组），可选左上编号徽章、标题与右上胶囊；
    返回去掉顶部标题行后的内区 Region，便于在框内直接摆放子卡片。
    """
    add_solid_box(slide, left, top, width, height, None, fill_color=None, line_color=color,
                  line_width=1.6, dashed=True, rounded=True, corner=corner, name=_dname(f"Group {title or ''}"))
    top_used = 0.08
    if title or badge or pill:
        hx = left + 0.10
        hy = top + 0.08
        hh = 0.36
        if badge:
            add_solid_box(slide, hx, hy, 0.38, hh, [badge], fill_color=color, line_color=color,
                          text_color=C_WHITE, size_pt=16, margin_lr=0.0, margin_tb=0.0,
                          rounded=True, corner=0.18, min_pt=MIN_DIAGRAM_PT, name=_dname(f"Badge {badge}"))
            hx += 0.44
        right_end = left + width - 0.10
        if pill:
            px = right_end - pill_w
            add_pill(slide, px, hy, pill_w, hh, pill, color=pill_color, pt=16, label=f"Pill {pill}")
            right_end = px - 0.08
        if title:
            add_diagram_label(slide, hx, hy - 0.02, max(0.5, right_end - hx), hh + 0.04,
                              title, pt=max(MIN_DIAGRAM_PT, title_pt), color=color,
                              align=PP_ALIGN.LEFT, label=f"Title {title}")
        top_used = 0.52
    return Region(left + 0.11, top + top_used, width - 0.22, height - top_used - 0.08)


def _default_roles(n):
    if n == 1:
        return ["core"]
    return ["input"] + ["core"] * (n - 2) + ["output"]


# --------------------------------------------------------------------------- #
# 组合图元
# --------------------------------------------------------------------------- #
def add_flow_pipeline(slide, left, top, width, height, steps, roles=None, direction="h",
                      gap=0.45, size_pt=18, arrow_color=C_BLUE, rounded=True):
    """
    N 步原生流程图。steps: 每步为 str 或多行 list（支持 _{sub}/^{sup} 与多色分段）。
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
            shapes.append(add_role_box(slide, *g, st, role=rl, size_pt=size_pt,
                                       rounded=rounded, label=f"pipeline step {i + 1}"))
            geoms.append(g)
    else:
        bh = (height - gap * (n - 1)) / n
        for i, (st, rl) in enumerate(zip(steps, roles)):
            g = (left, top + i * (bh + gap), width, bh)
            shapes.append(add_role_box(slide, *g, st, role=rl, size_pt=size_pt,
                                       rounded=rounded, label=f"pipeline step {i + 1}"))
            geoms.append(g)
    for a, b in zip(geoms[:-1], geoms[1:]):
        conns += connect_boxes(slide, a, b, color=arrow_color)
    return {"shapes": shapes, "geoms": geoms, "connectors": conns}


def add_feedback_loop(slide, left, top, width, height, steps, feedback_label, roles=None,
                      gap=0.45, box_h=None, size_pt=18, loop_color=C_RED):
    """
    闭环流程：上方横向前向链，下方从末步底部 → 回到首步底部的虚线反馈折线，并在折线下方标注。
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
    res["label"] = add_diagram_label(
        slide, lbl_left, y_loop + 0.04, lbl_w, 0.40, feedback_label,
        pt=size_pt, color=loop_color, align=PP_ALIGN.CENTER, label="feedback label",
    )
    return res


def add_layered_architecture(slide, left, top, width, height, layers, label_w=1.55,
                             row_gap=0.32, col_gap=0.14, size_pt=18, arrows=True,
                             label_fill=C_BLUE, flow="up"):
    """
    分层系统架构图（自上而下）。layers: [(层名, [模块1, 模块2, ...]), ...]。
    每层 = 左侧深蓝层名盒 + 右侧等宽浅蓝模块盒；层间以居中箭头连接（无嵌套容器，零重叠）。
    flow: "up"（数据自底层感知流向上层应用）/ "down"（指令下发）/ "both"。
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


def add_timeline(slide, region, items, card_pt=16, axis_label="时间"):
    """
    原生可编辑时间轴演进图（阶梯式左下→右上排布）。
    items: [(year_str, title_str, desc_str, role), ...]（role ∈ plain/input/core/output/emphasis）。
    """
    x, y, w, h = as_box(region)
    n = len(items)
    axis_y = y + h - 0.42
    add_connector(slide, x + 0.10, axis_y, x + w - 0.10, axis_y, color=C_BLUE, width_pt=2.5, arrow_end=True)
    if axis_label:
        add_diagram_label(slide, x + w - 0.94, axis_y - 0.34, 0.84, 0.28, axis_label,
                          pt=max(MIN_DIAGRAM_PT, card_pt - 2), color=C_BLUE, align=PP_ALIGN.RIGHT, label="axis")
    gap = 0.28
    cw = (w - 0.20 - gap * (n - 1)) / n
    ch = min(1.30, max(0.86, (axis_y - y) * 0.42))
    top_span = max(0.0, axis_y - y - ch - 0.42)
    geoms = []
    for i, (yr, title, desc, role) in enumerate(items):
        cx = x + 0.10 + i * (cw + gap)
        cy = (axis_y - ch - 0.28) - (top_span * i / max(1, n - 1))
        mx = cx + cw / 2.0
        col = C_RED if role in ("output", "emphasis") else C_BLUE
        add_connector(slide, mx, cy + ch, mx, axis_y, color=col, width_pt=1.5, arrow_end=False)
        add_junction_dot(slide, mx, axis_y, r=0.065, fill=C_WHITE, border=col, border_pt=2.0)
        yw = min(cw, 1.40)
        add_diagram_label(slide, mx - yw / 2.0, axis_y + 0.08, yw, 0.30, yr,
                          pt=max(MIN_DIAGRAM_PT, card_pt), color=col, align=PP_ALIGN.CENTER, label=f"year {yr}")
        lines = [title, desc] if desc else [title]
        add_role_box(slide, cx, cy, cw, ch, lines, role=role, size_pt=card_pt, label=f"timeline {yr}")
        geoms.append((cx, cy, cw, ch))
    for a, b in zip(geoms[:-1], geoms[1:]):
        connect_boxes(slide, a, b, side1="right", side2="left", color=C_BLUE, width_pt=1.8)
    return geoms
