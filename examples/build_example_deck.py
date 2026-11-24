#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Example opening-report deck built only with the v2 API and the page recipes in references/page_recipes.md.

Topic (generic): 面向动态负载工况的移动机器人自适应协同控制研究
Run:  python examples/build_example_deck.py [out_dir]     (default: examples/out/)
Outputs: example_opening.pptx, figures/*.png|svg, rendered slide PNGs; prints the audit.

Every figure is drawn at the exact size of the slot it goes into (figure_slot / image_card_slot), so
panel images are flush with their header bars and 18 pt in the figure stays 18 pt on the slide.
"""

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "scripts"))

import numpy as np  # noqa: E402

from tjuppt import (  # noqa: E402
    C_BLACK, C_BLUE, C_BLUE_LIGHT, C_RED, C_WHITE, Deck, add_card, add_figure_panel, add_image_card,
    add_omml_equation, add_picture_exact, add_solid_box, add_table, card_height, figure_slot, image_card_slot,
    m_r, m_sub,
)
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from tju_native_diagrams import add_role_box, connect_boxes  # noqa: E402
from tju_plot_style import apply_tju_style, new_figure, save_tju_figure, tju_line_styles  # noqa: E402
from tju_render import render  # noqa: E402
from tju_svg_diagram import TjuDiagram  # noqa: E402

OUT = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "out"))
FIG = os.path.join(OUT, "figures")
os.makedirs(FIG, exist_ok=True)


def save(d, name):
    return d.save(os.path.join(FIG, f"{name}.png"), svg_path=os.path.join(FIG, f"{name}.svg"))


# --------------------------------------------------------------------------- #
# Figures — each takes the slot size (w, h) in inches
# --------------------------------------------------------------------------- #
def fig_scene(w, h):
    d = TjuDiagram(w, h, base_pt=18)
    bw, gap = (w - 0.4 - 2 * 0.3) / 3, 0.3
    xs = [0.2 + i * (bw + gap) for i in range(3)]
    d.node("load", xs[0], 0.25, bw, 0.9, "载荷突变", role="input")
    d.node("path", xs[1], 0.25, bw, 0.9, "窄巷曲率", role="input")
    d.node("fric", xs[2], 0.25, bw, 0.9, "地面摩擦", role="output")
    d.node("robot", w * 0.18, h - 1.25, w * 0.64, 0.95, "高举升移动机器人", role="core")
    for n in ("load", "path", "fric"):
        d.edge(n, "robot", src_side="bottom", dst_side="top", route="straight" if n == "path" else "vhv")
    return save(d, "scene")


def fig_challenge(w, h, kind):
    d = TjuDiagram(w, h, base_pt=18)
    top_txt, mid_txt, bot_txt = {1: ("载荷 0→500 kg", "固定模型", "参数漂移"),
                                 2: ("曲率 × 速度", "线性化误差", "横向误差放大"),
                                 3: ("控制周期 10 ms", "在线优化", "求解超时")}[kind]
    bh, gap = 0.70, (h - 0.4 - 3 * 0.70) / 2
    for i, (nid, txt, role) in enumerate([("a", top_txt, "input"), ("m", mid_txt, "plain"), ("b", bot_txt, "output")]):
        d.node(nid, 0.25, 0.2 + i * (bh + gap), w - 0.5, bh, txt, role=role)
    d.edge("a", "m", src_side="bottom", dst_side="top")
    d.edge("m", "b", src_side="bottom", dst_side="top", color="red")
    return save(d, f"challenge{kind}")


def fig_timeline(w, h):
    d = TjuDiagram(w, h, base_pt=18)
    y_axis = h - 0.55
    d.arrow([(0.3, y_axis), (w - 0.25, y_axis)], color="black")
    eras = [("2000", "PID / 增益调度", "plain"), ("2010", "模型预测控制", "input"),
            ("2018", "强化学习控制", "input"), ("2023", "数据-模型融合自适应", "output")]
    seg = (w - 0.8) / len(eras)
    for i, (year, name, role) in enumerate(eras):
        x = 0.45 + i * seg
        d.text(x + 0.05, y_axis + 0.36, year)
        top = y_axis - 0.95 - i * (y_axis - 1.5) / 3
        d.node(f"e{i}", x, top, seg - 0.35, 0.80, name, role=role)
        if i:
            d.edge(f"e{i - 1}", f"e{i}", route="hvh")
    return save(d, "timeline")


def fig_loop(w, h):
    d = TjuDiagram(w, h, base_pt=18)
    bw = (w - 0.3 - 2 * 0.45) / 3
    xs = [0.15 + i * (bw + 0.45) for i in range(3)]
    y = h * 0.15
    d.node("obs", xs[0], y, bw, 1.0, "扰动观测\n研究内容一", role="input")
    d.node("mpc", xs[1], y, bw, 1.0, "自适应 MPC\n研究内容二", role="core")
    d.node("dep", xs[2], y, bw, 1.0, "嵌入式部署\n研究内容三", role="output")
    d.node("plant", xs[1], h * 0.62, bw, 0.85, "机器人本体", role="plain")
    d.edge("obs", "mpc")
    d.edge("mpc", "dep")
    d.edge("dep", "plant", src_side="bottom", dst_side="right", route="vh")
    d.edge("plant", "obs", src_side="left", dst_side="bottom", route="hv", color="red", dashed=True)
    return save(d, "loop")


def fig_observer(w, h):
    d = TjuDiagram(w, h, base_pt=18)
    d.node("u", 0.2, 0.35, 1.5, 0.8, "控制量 u_{k}", role="input")
    d.node("y", 0.2, h - 1.15, 1.5, 0.8, "测量 y_{k}", role="input")
    d.node("eso", w / 2 - 0.9, h / 2 - 0.55, 1.8, 1.1, "扩张状态\n观测器", role="core")
    d.node("d", w - 1.85, h / 2 - 0.45, 1.65, 0.9, "扰动估计 d_{k}", role="output")
    d.edge("u", "eso", route="hv", dst_side="top")
    d.edge("y", "eso", route="hv", dst_side="bottom")
    d.edge("eso", "d")
    return save(d, "observer")


def fig_error_plot(w, h):
    apply_tju_style(base_pt=18, tick_pt=16, legend_pt=16)
    fig, ax = new_figure(w, h)
    t = np.arange(0, 31, 2)
    rng = np.random.default_rng(7)
    for (name, base, amp), st in zip([("本文方法（预期）", 0.05, 0.20), ("固定参数 MPC", 0.11, 0.22)],
                                     tju_line_styles(markersize=5)):
        ax.plot(t, base + amp * np.exp(-t / 8) + rng.normal(0, 0.003, len(t)), label=name, **st)
    ax.set_xlabel("时间 / s")
    ax.set_ylabel("横向误差 / m")
    ax.set_ylim(0, 0.42)
    ax.legend(loc="upper right")
    return save_tju_figure(fig, os.path.join(FIG, "error.png"))


def fig_mpc(w, h):
    d = TjuDiagram(w, h, base_pt=18)
    bh, gap = 0.72, (h - 0.4 - 3 * 0.72) / 2
    for i, (nid, txt, role) in enumerate([("ref", "参考轨迹与约束集", "input"), ("opt", "带扰动补偿的滚动优化", "core"),
                                          ("out", "首步控制量下发", "output")]):
        d.node(nid, 0.3, 0.2 + i * (bh + gap), w - 0.6, bh, txt, role=role)
    d.edge("ref", "opt")
    d.edge("opt", "out")
    return save(d, "mpc")


# --------------------------------------------------------------------------- #
# Deck
# --------------------------------------------------------------------------- #
def build():
    deck = Deck()
    deck.cover(["面向动态负载工况的移动机器人", "自适应协同控制研究"],
               ["汇报人：张三    导师：李四 教授", "2026 年 11 月"])
    deck.toc(["01  研究背景与现状", "02  研究内容与方案", "03  创新点与可行性", "04  研究计划"], active=0)

    # 背景：两行宏观/行业需求 → 场景图 + 概念卡 → 红字结论
    p = deck.content("研究背景及意义", "应用场景：高举升移动机器人在动态负载下的窄巷循迹作业",
                     lead=[("国家需求：", "智能仓储装备自主可控，重载机器人需高可靠作业"),
                           ("行业需求：", "窄巷作业载荷频繁变化，现有控制器需人工反复整定")],
                     takeaway="载荷与路面动态变化下的稳定循迹，是高举升机器人可靠作业的前提")
    left, right = p.body.split(2)
    add_figure_panel(p.slide, left, "作业场景与扰动来源", fig_scene(*figure_slot(left)))
    concepts = [
        [("• 动态负载：", C_BLUE), ("载荷与重心随时间变化", C_BLACK)],
        [("• 循迹控制：", C_BLUE), ("沿参考路径运动，抑制横向误差", C_BLACK)],
        [("• 自适应：", C_RED), ("控制参数随工况在线调整", C_BLACK)],
    ]
    card, rest = right.take_top(card_height(right.w, concepts, line_spacing=1.3), gap=0.30)
    add_card(p.slide, card, "核心概念", concepts, line_spacing=1.3)
    chain = rest.split(3, gutter=0.45)
    for i, (bx, txt) in enumerate(zip(chain, ["载荷变化", "参数漂移", "误差增大"])):
        add_role_box(p.slide, *bx.box, txt, role=["input", "core", "output"][i])
        if i:
            connect_boxes(p.slide, chain[i - 1].box, bx.box)

    # 挑战：三列图文卡 → 红字结论
    p = deck.content("研究背景及意义", "关键挑战：参数漂移、强耦合与实时性约束同时存在",
                     takeaway="亟需兼顾自适应性与实时性的移动机器人循迹控制方法")
    cols = p.body.split(3)
    specs = [("1. 载荷突变致参数漂移", 1, "blue", ["固定模型失配", "跟踪误差增大"]),
             ("2. 曲率与速度强耦合", 2, "blue", ["窄巷转弯处", "横向误差放大"]),
             ("3. 车载算力受限", 3, "red", ["优化求解耗时", "难满足 10 ms 周期"])]
    for col, (hdr, k, tone, cap) in zip(cols, specs):
        add_image_card(p.slide, col, hdr, fig_challenge(*image_card_slot(col, len(cap)), k), cap, tone=tone,
                       fit="exact")

    # 现状（一）：方法演进时间轴，先交代传统方法局限
    p = deck.content("国内外研究现状", "方法演进：从固定增益、模型预测控制（MPC）到数据-模型融合自适应控制",
                     takeaway="传统方法难以兼顾动态工况适应性与实时性，学习型方法缺乏安全保证")
    add_picture_exact(p.slide, fig_timeline(*p.body.box[2:]), p.body)

    # 总体框架：一行研究目标 → 左挑战-内容映射（原生可编辑）+ 右闭环框图
    p = deck.content("研究内容与方案", "研究目标与总体框架：三项研究内容对应三类挑战",
                     lead=[("研究目标：", "构建动态负载下可在线自适应、满足车载实时性的循迹控制方法")],
                     takeaway="观测、优化、部署三项内容逐一回应三类挑战，构成闭环")
    left, right = p.body.split(2, [5, 7])
    rows = left.vsplit(3, gap=0.30)
    pairs = [("参数漂移", "研究内容一\n扰动观测"), ("强耦合", "研究内容二\n自适应 MPC"), ("算力受限", "研究内容三\n嵌入式部署")]
    for r, (chal, work) in zip(rows, pairs):
        a, b = r.split(2, [2, 3], gutter=0.55)
        add_role_box(p.slide, *a.box, chal, role="output")
        add_role_box(p.slide, *b.box, work.split("\n"), role="core")
        connect_boxes(p.slide, a.box, b.box)
    add_figure_panel(p.slide, right, "总体技术框架（闭环）", fig_loop(*figure_slot(right)))

    # 研究内容一：三条子内容 → 双图框 → 红字结论
    p = deck.content("研究内容与方案", "研究内容一：基于扩张状态观测器的动态负载扰动估计",
                     lead=[("扰动建模：", "将载荷变化与路面摩擦统一建模为集总扰动"),
                           ("观测器设计：", "设计扩张状态观测器在线估计集总扰动 d_{k}"),
                           ("收敛分析：", "给出观测误差有界的条件与带宽选取依据", "red")],
                     takeaway="在线估计集总扰动，为自适应控制提供实时补偿量")
    left, right = p.body.split(2)
    add_figure_panel(p.slide, left, "扰动观测器结构", fig_observer(*figure_slot(left)))
    add_figure_panel(p.slide, right, "预期估计效果（示意）", fig_error_plot(*figure_slot(right)), tone="red")

    # 研究内容二：左图框 + 右（标题条 + 公式 + 要点）
    p = deck.content("研究内容与方案", "研究内容二：融合扰动补偿的自适应模型预测控制",
                     lead=[("预测模型：", "将扰动估计并入预测方程，修正模型失配"),
                           ("代价设计：", "跟踪误差与控制增量加权，权重随工况自适应"),
                           ("约束处理：", "速度、加速度与转角约束保证执行安全", "red")],
                     takeaway="扰动补偿与权重自适应协同，降低动态负载下的跟踪误差")
    left, right = p.body.split(2)
    add_figure_panel(p.slide, left, "自适应 MPC 求解流程", fig_mpc(*figure_slot(left)))
    hdr, rest = right.take_top(0.40)
    eq, notes = rest.take_top(1.20)
    add_solid_box(p.slide, *hdr.box, ["滚动优化代价函数"], fill_color=C_BLUE, line_color=C_BLUE, text_color=C_WHITE)
    cost = (m_r("J", color="00337F") + m_r(" = ", italic=False) + m_r("Σ", italic=False)
            + m_sub(m_r("‖e‖", italic=False), m_r("Q", color="00337F")) + m_r(" + ", italic=False)
            + m_sub(m_r("‖Δu‖", italic=False), m_r("R", color="00337F")))
    add_omml_equation(p.slide, *eq.box, cost, jc="center", fill_color=C_BLUE_LIGHT, line_color=C_BLUE, pad_lr=1.0)
    add_solid_box(p.slide, *notes.box, [
        [("• e：", C_BLUE), ("横向与航向跟踪误差", C_BLACK)],
        [("• Q、R：", C_BLUE), ("随载荷在线调整的权重矩阵", C_BLACK)],
        [("• d_{k}：", C_RED), ("研究内容一给出的扰动估计", C_BLACK)],
    ], line_color=C_BLUE, align=PP_ALIGN.LEFT, margin_lr=0.15, line_spacing=1.5)

    # 创新点：两张卡（问题 → 区别 → 方法 → 预期）
    p = deck.content("拟创新点", "围绕扰动估计与自适应优化提出两点创新",
                     takeaway="以扰动估计驱动权重自适应，兼顾动态适应性与车载实时性")
    strip, cards = p.body.take_bottom(0.80, gap=0.30)
    a, b = cards.vsplit(2, gap=0.25)
    add_card(p.slide, a, "创新点一：观测器驱动的权重自适应机制", [
        "针对载荷突变致权重失配，区别于离线整定，提出扰动估计驱动的在线调权，有望降低误差。"],
        line_spacing=1.3)
    add_card(p.slide, b, "创新点二：面向车载算力的显式求解加速", [
        "针对在线优化难满足 10 ms 周期，区别于通用求解器，提出分段显式求解与热启动加速。"],
        tone="red", line_spacing=1.3)
    boxes = strip.split(3, gutter=0.55)
    for i, (bx, txt) in enumerate(zip(boxes, ["扰动估计", "调权 + 加速", "误差↓ 周期满足"])):
        add_role_box(p.slide, *bx.box, txt, role=["input", "core", "output"][i])
        if i:
            connect_boxes(p.slide, boxes[i - 1].box, bx.box)

    # 可行性：左研究条件 / 右研究基础
    p = deck.content("研究基础与可行性", "研究条件与研究基础：具备实验平台与前期控制算法积累",
                     takeaway="软硬件条件与前期积累可支撑课题按期完成")
    left, right = p.body.split(2)
    flow, top = left.take_bottom(1.20, gap=0.30)
    add_card(p.slide, top, "研究条件（软硬件）", [
        [("• 实验平台：", C_BLUE), ("高举升移动机器人 1 台", C_BLACK)],
        [("• 仿真环境：", C_BLUE), ("Gazebo 动力学仿真", C_BLACK)],
        [("• 算力：", C_BLUE), ("工作站 + 车载嵌入式控制器", C_BLACK)],
    ], line_spacing=2.0)
    steps = flow.split(3, gutter=0.45)
    for i, (s, txt) in enumerate(zip(steps, ["仿真验证", "半实物测试", "实车验证"])):
        add_role_box(p.slide, *s.box, txt, role=["input", "core", "output"][i])
        if i:
            connect_boxes(p.slide, steps[i - 1].box, s.box)
    found = [("1. 整车动力学模型", "已完成建模与参数辨识", "blue"),
             ("2. MPC 循迹基线", "已实现固定参数 MPC 并实车运行", "blue"),
             ("3. 多工况数据", "已采集空载 / 半载 / 满载数据", "red")]
    for r, (hdr_txt, body, tone) in zip(right.vsplit(3, gap=0.20), found):
        add_card(p.slide, r, hdr_txt, [body], tone=tone)

    # 研究计划：表格（行高按文字自动估算）
    p = deck.content("研究计划", "研究计划：12 个月四阶段推进",
                     takeaway="按阶段交付观测器、控制器与实车验证结果")
    tbl, outcome = p.body.take_top(2.75, gap=0.25)
    add_table(p.slide, tbl, [2, 5, 3], fill=True, rows_data=[
        ["时间", "研究任务", "阶段成果"],
        ["2026.12–2027.02", "扰动建模与观测器设计", "观测器仿真验证"],
        ["2027.03–2027.06", "自适应 MPC 设计与求解加速", "控制器仿真对比"],
        ["2027.07–2027.09", "实车实验与参数优化", "实验数据与分析"],
        ["2027.10–2027.12", "总结与论文撰写", "学位论文初稿"],
    ])
    add_card(p.slide, outcome, "预期成果", [
        [("• 指标：", C_BLUE), ("满载工况横向误差较固定参数 MPC 降低，控制周期 ≤ 10 ms", C_BLACK)],
        [("• 成果：", C_BLUE), ("学位论文 1 篇，学术论文 1–2 篇", C_BLACK)],
    ])
    deck.thanks()

    path, issues = deck.save(os.path.join(OUT, "example_opening.pptx"))
    pngs = render(path, os.path.join(OUT, "render"))
    print(f"rendered {len(pngs)} slides → {os.path.join(OUT, 'render')}")
    return path, issues


if __name__ == "__main__":
    _, iss = build()
    sys.exit(1 if iss else 0)
