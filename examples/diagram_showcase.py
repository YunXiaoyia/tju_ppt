#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Figure-route showcase (see references/figures.md): four content pages on the TJU template.
  P1 A1/A2 native flow + feedback loop      (tju_native_diagrams, editable)
  P2 A3 native layered architecture         (tju_native_diagrams, editable)
  P3 C  roadmap drawn at slot size          (TjuDiagram → SVG/PNG) + explanation card
  P4 D  plots drawn at slot size            (tju_plot_style) + evaluation card
Run:  python examples/diagram_showcase.py [out_dir]     (default examples/out/showcase/)
"""

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "scripts"))

import numpy as np  # noqa: E402

from pptx.enum.text import MSO_ANCHOR  # noqa: E402
from tjuppt import (  # noqa: E402
    C_BLACK, C_BLUE, C_RED, Deck, add_card, add_figure_panel, add_text_block, card_height, figure_slot,
)
from tju_audit import audit_image_palette  # noqa: E402
from tju_native_diagrams import add_feedback_loop, add_flow_pipeline, add_layered_architecture  # noqa: E402
from tju_plot_style import TJU_BAR_STYLES, apply_tju_style, new_figure, save_tju_figure, tju_line_styles  # noqa: E402
from tju_svg_diagram import TjuDiagram  # noqa: E402

OUT = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "out", "showcase"))
os.makedirs(OUT, exist_ok=True)


def roadmap(w, h, path):
    """Two research contents merging into the controller; coordinates scale with the slot."""
    d = TjuDiagram(w, h)
    gw = w - 0.50                                   # groups leave a right lane for the merge connector
    gh = (h - 1.45) / 2 - 0.15
    bw = (gw - 0.30 - 2 * 0.35) / 3
    xs = [0.20 + i * (bw + 0.35) for i in range(3)]
    for gi, (label, items) in enumerate([("研究内容一：扰动预测", [("data", "多工况数据", "input"), ("obs", "高阶观测器", "core"),
                                                             ("pred", "扰动预测 d_{k+N}", "core")]),
                                         ("研究内容二：工况约束", [("ana", "响应分析", "input"), ("bnd", "运行边界", "core"),
                                                             ("cons", "约束集 U_{safe}", "core")])]):
        gy = 0.12 + gi * (gh + 0.25)
        d.group(0.05, gy, gw, gh, label)
        for x, (nid, txt, role) in zip(xs, items):
            d.node(nid, x, gy + 0.42, bw, gh - 0.55, txt, role=role)
        d.edge(items[0][0], items[1][0])
        d.edge(items[1][0], items[2][0])
    d.node("rl", 0.25, h - 1.0, 1.9, 0.85, "强化学习调参", role="emphasis")
    d.node("mpc", w - 2.85, h - 1.0, 2.40, 0.85, "工况自适应 MPC", role="output")
    d.edge("pred", "mpc", src_side="right", dst_side="right",
           via=[(w - 0.22, 0.12 + 0.42 + (gh - 0.55) / 2), (w - 0.22, h - 0.575)])
    d.edge("cons", "mpc", label="约束", label_offset=(0.45, 0.05))
    d.edge("rl", "mpc", label="权重 Q_{k}")
    return d.save(path, svg_path=path.replace(".png", ".svg"))


def curve(w, h, path):
    apply_tju_style()
    rng = np.random.default_rng(0)
    x = np.arange(0, 41, 2)
    fig, ax = new_figure(w, h)
    for (name, base), st in zip([("本文方法", 0.06), ("固定权重 MPC", 0.11), ("PID", 0.17), ("LQR", 0.14)],
                                tju_line_styles(markersize=5)):
        ax.plot(x, base + 0.25 * np.exp(-x / 9) + rng.normal(0, 0.004, len(x)), label=name, **st)
    ax.set_xlabel("时间 / s")
    ax.set_ylabel("横向误差 / m")
    ax.set_ylim(0, 0.45)
    ax.legend(ncol=1, loc="center left", bbox_to_anchor=(1.0, 0.5), handlelength=1.6)
    return save_tju_figure(fig, path)


def bars(w, h, path):
    apply_tju_style()
    fig, ax = new_figure(w, h)
    vals = np.array([[0.05, 0.07, 0.08], [0.08, 0.10, 0.13], [0.07, 0.09, 0.12], [0.10, 0.14, 0.19]])
    for i, (m, st) in enumerate(zip(["本文", "w/o 预测", "w/o RL", "固定 MPC"], TJU_BAR_STYLES)):
        ax.bar(np.arange(3) + (i - 1.5) * 0.2, vals[i], 0.2, label=m, **st)
    ax.set_xticks(np.arange(3))
    ax.set_xticklabels(["空载", "半载", "满载"])
    ax.set_ylabel("RMSE / m")
    ax.set_ylim(0, 0.22)
    ax.legend(ncol=1, loc="center left", bbox_to_anchor=(1.0, 0.5), handlelength=1.2)
    return save_tju_figure(fig, path)


def main():
    deck = Deck()
    p = deck.content("研究内容与方案", "总体思路：感知、预测、优化、执行构成闭环链路",
                     takeaway="前向链给出控制量，闭环迭代用实车误差修正模型")
    hdr, rest = p.body.take_top(0.45, gap=0.10)
    add_text_block(p.slide, *hdr.box, [("① 前向控制链", C_BLUE, 19)])
    flow, rest = rest.take_top(0.95, gap=0.40)
    add_flow_pipeline(p.slide, *flow.box, ["状态感知 x_{k}", "扰动预测 d_{k}", "滚动优化 u_{k}^{*}", "执行与约束校核"])
    hdr, loop = rest.take_top(0.45, gap=0.10)
    add_text_block(p.slide, *hdr.box, [("② 闭环迭代", C_BLUE, 19)])
    add_feedback_loop(p.slide, *loop.box, ["仿真采样", "策略更新", "实车验证"], "误差反馈：修正奖励函数与约束边界")

    p = deck.content("研究内容与方案", "系统架构：四层解耦，红框为本文新增模块",
                     takeaway="新增模块集中在决策层与控制层，感知与应用层复用既有成果")
    add_layered_architecture(p.slide, *p.body.box, [
        ("应用层", ["仓储搬运", "高位举升", "狭窄通道循迹"]),
        ("决策层", ["!工况辨识", "!强化学习调参", "任务调度"]),
        ("控制层", ["!扰动预测 MPC", "约束处理", "执行器接口"]),
        ("感知层", ["激光雷达", "IMU / 轮速计", "载荷传感器"]),
    ])

    p = deck.content("研究内容与方案", "技术路线：两条研究内容汇入自适应 MPC",
                     takeaway="核心判据：任何工况下约束不被违反")
    left, right = p.body.split(2, [7, 5])
    road = roadmap(*figure_slot(left), os.path.join(OUT, "roadmap.png"))
    add_figure_panel(p.slide, left, "技术路线图", road)
    add_card(p.slide, right, "路线说明", [
        [("• 内容一：", C_BLUE), ("数据训练观测器", C_BLACK)],
        [("• 内容二：", C_BLUE), ("响应分析得约束集", C_BLACK)],
        [("• 汇合：", C_RED), ("强化学习调节 Q_{k}", C_BLACK)],
        [("• 输出：", C_BLUE), ("首步控制量 u_{k}", C_BLACK)],
        [("• 校核：", C_BLUE), ("约束集在线检查", C_BLACK)],
    ], line_spacing=2.0)

    p = deck.content("研究内容与方案", "预期验证：跟踪误差收敛与消融对比（示意数据）",
                     takeaway="相同底层与工况下对比基线，并逐项消融验证各模块贡献")
    left, right = p.body.split(2, [5, 7])
    design = [
        [("• 指标：", C_BLUE), ("横向误差、RMSE、违约次数", C_BLACK)],
        [("• 工况：", C_BLUE), ("空载 / 半载 / 满载", C_BLACK)],
        [("• 消融：", C_BLUE), ("去掉预测模块或 RL 调参", C_BLACK)],
    ]
    flow, rest = left.take_bottom(0.95)
    hdr, card = rest.take_bottom(0.45, gap=0.30)
    add_card(p.slide, card.with_height(card.h - 0.10), "评价设计", design, line_spacing=2.0)
    add_text_block(p.slide, *hdr.box, [("验证流程", C_BLUE, 19)])
    add_flow_pipeline(p.slide, *flow.box, ["仿真", "训练", "实车"], gap=0.40)
    up, down = right.vsplit(2)
    c = curve(*figure_slot(up), os.path.join(OUT, "curve.png"))
    b = bars(*figure_slot(down), os.path.join(OUT, "ablation.png"))
    add_figure_panel(p.slide, up, "跟踪误差收敛", c)
    add_figure_panel(p.slide, down, "消融对比", b, tone="red")

    deck.save(os.path.join(OUT, "diagram_showcase.pptx"), content=False)
    for f in (road, c, b):
        r = audit_image_palette(f)
        print(f"[image] {os.path.basename(f)} ok={r['ok']} off={r['off_ratio']:.4f}")


if __name__ == "__main__":
    main()
