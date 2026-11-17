#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
科研流程图 / 架构图 / 数据图 四类绘图路线演示（generate_diagram_showcase.py）

在课题组模板上生成 4 页内容页，分别演示 SKILL.md §5 的四条绘图路线：
  P1 原生可编辑：N 步流程 + 闭环反馈            (tju_native_diagrams)
  P2 原生可编辑：分层系统架构                     (tju_native_diagrams)
  P3 复杂技术路线图：结构化描述 → 校验 → SVG/PNG  (tju_svg_diagram)
  P4 答辩投影级实验数据图：曲线 + 消融柱状图      (tju_plot_style)
随后执行 OpenXML 审计与图片调色板审计 (tju_audit)。

运行：python examples/generate_diagram_showcase.py  [输出目录默认 examples/showcase_out/]
"""

import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, "../scripts")))

import numpy as np
from pptx import Presentation
from pptx.enum.text import PP_ALIGN

from ppt_builder_helpers import (
    C_BLACK, C_BLUE, C_RED, add_dashed_box, add_image_fit, add_text_block,
    enforce_min_18pt_all_slides, setup_tju_content_slide,
)
from tju_native_diagrams import add_feedback_loop, add_flow_pipeline, add_layered_architecture
from tju_svg_diagram import TjuDiagram
from tju_plot_style import TJU_BAR_STYLES, apply_tju_style, new_figure, save_tju_figure, tju_line_styles
from tju_audit import audit_image_palette, audit_pptx

TEMPLATE = os.path.abspath(os.path.join(HERE, "../references/tju_template.pptx"))
OUT_DIR = os.path.join(HERE, "showcase_out")
OUT_PPTX = os.path.join(OUT_DIR, "diagram_showcase.pptx")
KEEP = [3, 12, 15, 18]  # 模板中带顶部蓝色标题栏的内容页（1-based）

# 双栏虚线框标准几何（与 SKILL.md §2 一致）
COL_TOP, COL_H = 1.78, 5.22
L_LEFT, L_W = 0.45, 6.10
R_LEFT, R_W = 6.78, 6.10


def keep_slides(prs, keep):
    sld_ids = prs.slides._sldIdLst
    for idx, sid in reversed(list(enumerate(list(sld_ids), start=1))):
        if idx not in keep:
            prs.part.drop_rel(sid.rId)
            sld_ids.remove(sid)


# --------------------------------------------------------------------------- #
def build_svg_roadmap(path):
    """P3：技术路线图（左栏 6.10 × 5.22 in 放置框，按 1:1 尺寸作图）"""
    d = TjuDiagram(5.90, 5.00)
    d.group(0.05, 0.05, 5.80, 1.50, "研究内容一：扰动预测")
    d.group(0.05, 1.80, 5.80, 1.50, "研究内容二：工况约束")
    d.node("data", 0.25, 0.60, 1.55, 0.75, "多工况数据", role="input")
    d.node("obs", 2.20, 0.60, 1.55, 0.75, "高阶观测器", role="core")
    d.node("pred", 4.10, 0.60, 1.55, 0.75, "扰动预测 d_{k+N}", role="core")
    d.node("ana", 0.25, 2.35, 1.55, 0.75, "响应分析", role="input")
    d.node("bnd", 2.20, 2.35, 1.55, 0.75, "运行边界", role="core")
    d.node("cons", 4.10, 2.35, 1.55, 0.75, "约束集 U_{safe}", role="core")
    d.node("rl", 0.25, 3.90, 1.80, 0.85, "强化学习调参", role="emphasis")
    d.node("mpc", 3.40, 3.90, 2.25, 0.85, "工况自适应 MPC", role="output")
    d.edge("data", "obs")
    d.edge("obs", "pred")
    d.edge("ana", "bnd")
    d.edge("bnd", "cons")
    d.edge("pred", "mpc", src_side="right", dst_side="right", via=[(5.78, 0.975), (5.78, 4.325)])
    d.edge("cons", "mpc", label="约束")
    d.edge("rl", "mpc", label="权重 Q_{k}")
    d.save(path, svg_path=path.replace(".png", ".svg"))
    return path


def build_plots(curve_png, bar_png):
    """P4：右栏上下两张数据图，各 6.0 × 2.45 in"""
    apply_tju_style()
    rng = np.random.default_rng(0)
    x = np.arange(0, 41, 2)
    fig, ax = new_figure(6.0, 2.45)
    for (name, base), st in zip([("本文方法", 0.06), ("固定权重 MPC", 0.11), ("PID", 0.17), ("LQR", 0.14)],
                                tju_line_styles(markersize=5)):
        y = base + 0.25 * np.exp(-x / 9) + rng.normal(0, 0.004, len(x))
        ax.plot(x, y, label=name, **st)
    ax.set_xlabel("时间 / s")
    ax.set_ylabel("横向误差 / m")
    ax.set_ylim(0, 0.62)
    ax.legend(ncol=2, loc="upper right")
    save_tju_figure(fig, curve_png)

    fig, ax = new_figure(6.0, 2.45)
    groups = ["空载", "半载", "满载"]
    methods = ["本文", "w/o 预测", "w/o RL", "固定 MPC"]
    vals = np.array([[0.05, 0.07, 0.08], [0.08, 0.10, 0.13], [0.07, 0.09, 0.12], [0.10, 0.14, 0.19]])
    w = 0.2
    for i, (m, st) in enumerate(zip(methods, TJU_BAR_STYLES)):
        ax.bar(np.arange(3) + (i - 1.5) * w, vals[i], w, label=m, **st)
    ax.set_xticks(np.arange(3))
    ax.set_xticklabels(groups)
    ax.set_ylabel("RMSE / m")
    ax.legend(ncol=4, loc="upper left", fontsize=14, handlelength=1.2, columnspacing=0.8)
    ax.set_ylim(0, 0.26)
    save_tju_figure(fig, bar_png)


# --------------------------------------------------------------------------- #
def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    prs = Presentation(TEMPLATE)
    keep_slides(prs, KEEP)
    s1, s2, s3, s4 = prs.slides

    # ---------------- P1 ----------------
    setup_tju_content_slide(s1, "研究内容与方案", "总体思路：感知—预测—优化—执行的闭环链路", page_num=1)
    add_dashed_box(s1, L_LEFT, COL_TOP, 12.43, COL_H)
    add_text_block(s1, 0.70, 1.95, 11.9, 0.42, [("① 前向控制链（add_flow_pipeline）", C_BLUE, 20)])
    add_flow_pipeline(s1, 0.75, 2.45, 11.85, 1.05,
                      ["状态感知 x_{k}", "扰动预测 d̂_{k}", "滚动优化 u_{k}^{*}", "执行与约束校核"])
    add_text_block(s1, 0.70, 3.85, 11.9, 0.42, [("② 闭环迭代（add_feedback_loop）", C_BLUE, 20)])
    add_feedback_loop(s1, 1.60, 4.35, 10.20, 2.05,
                      ["仿真采样", "策略更新", "实车验证"], "误差反馈：修正奖励函数与约束边界")

    # ---------------- P2 ----------------
    setup_tju_content_slide(s2, "研究内容与方案", "系统架构：四层解耦，红框为本文新增模块", page_num=2)
    add_dashed_box(s2, L_LEFT, COL_TOP, 12.43, COL_H)
    add_layered_architecture(s2, 0.75, 2.00, 11.85, 4.75, [
        ("应用层", ["仓储搬运", "高位举升", "狭窄通道循迹"]),
        ("决策层", ["!工况辨识", "!强化学习调参", "任务调度"]),
        ("控制层", ["!扰动预测 MPC", "约束处理", "执行器接口"]),
        ("感知层", ["激光雷达", "IMU / 轮速计", "载荷传感器"]),
    ])

    # ---------------- P3 ----------------
    roadmap = build_svg_roadmap(os.path.join(OUT_DIR, "roadmap.png"))
    setup_tju_content_slide(s3, "研究内容与方案", "技术路线：两条研究内容汇入自适应 MPC", page_num=3)
    add_dashed_box(s3, L_LEFT, COL_TOP, L_W, COL_H)
    add_dashed_box(s3, R_LEFT, COL_TOP, R_W, COL_H)
    add_image_fit(s3, roadmap, L_LEFT + 0.10, COL_TOP + 0.11, L_W - 0.20, COL_H - 0.22, border_color=None)
    add_text_block(s3, R_LEFT + 0.18, COL_TOP + 0.18, R_W - 0.36, COL_H - 0.36, [
        ("路线说明", C_BLUE, 20),
        [("• 研究内容一：", C_BLUE, 18), ("由多工况数据训练高阶观测器，输出未来 N 步扰动预测。", C_BLACK, 18)],
        [("• 研究内容二：", C_BLUE, 18), ("由响应分析得到运行边界，构造安全约束集。", C_BLACK, 18)],
        [("• 汇合：", C_BLUE, 18), ("强化学习在线调节权重 Q_{k}，三路信息共同驱动 MPC。", C_BLACK, 18)],
        ("核心判据：任何工况下约束不被违反", C_RED, 18),
    ])

    # ---------------- P4 ----------------
    curve, bar = os.path.join(OUT_DIR, "curve.png"), os.path.join(OUT_DIR, "ablation.png")
    build_plots(curve, bar)
    setup_tju_content_slide(s4, "研究内容与方案", "预期验证：跟踪误差收敛与消融对比（示意数据）", page_num=4)
    add_dashed_box(s4, L_LEFT, COL_TOP, 5.20, COL_H)
    add_dashed_box(s4, 5.88, COL_TOP, 7.00, COL_H)
    add_text_block(s4, L_LEFT + 0.18, COL_TOP + 0.18, 4.84, 3.00, [
        ("评价设计", C_BLUE, 20),
        [("• 指标：", C_BLUE, 18), ("横向误差、RMSE、违约次数。", C_BLACK, 18)],
        [("• 工况：", C_BLUE, 18), ("空载 / 半载 / 满载三档载荷。", C_BLACK, 18)],
        [("• 消融：", C_BLUE, 18), ("分别去掉预测模块与 RL 调参。", C_BLACK, 18)],
        ("曲线仅用三色，第 4 条起靠线型区分", C_RED, 18),
    ])
    add_text_block(s4, L_LEFT + 0.18, 5.05, 4.84, 0.42, [("验证流程（横向 pipeline）", C_BLUE, 20)])
    add_flow_pipeline(s4, L_LEFT + 0.22, 5.62, 4.76, 1.05, ["仿真", "训练", "实车"], gap=0.40)
    add_image_fit(s4, curve, 5.98, COL_TOP + 0.10, 6.80, 2.45, border_color=None)
    add_image_fit(s4, bar, 5.98, COL_TOP + 2.67, 6.80, 2.45, border_color=None)

    enforce_min_18pt_all_slides(prs)
    staged = os.path.join(OUT_DIR, "staged.pptx")
    prs.save(staged)
    try:
        shutil.copy2(staged, OUT_PPTX)
    except PermissionError:
        print(f"[warn] {OUT_PPTX} 被占用，结果保留在 {staged}")
    print("saved", OUT_PPTX)

    issues = audit_pptx(OUT_PPTX)
    print(f"[audit pptx] {len(issues)} issue(s)")
    for i in issues:
        print("  -", i)
    for p in (roadmap, curve, bar):
        r = audit_image_palette(p)
        print(f"[audit image] {os.path.basename(p)} ok={r['ok']} off={r['off_ratio']:.4f} {r['top_offenders']}")


if __name__ == "__main__":
    main()
