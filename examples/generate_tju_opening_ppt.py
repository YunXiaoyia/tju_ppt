#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TJU-PPT 标准硕士开题汇报示例生成器（10 页全链路合规范例）
题目：《面向动态负载工况的移动机器人自适应协同控制研究》（通用学术示例）

完整示范 SKILL.md 全部规范与可复用原语：
- 模板安全裁剪与重排：rearrange_slides（基于 references/tju_template.pptx 原生母版）
- 封面 / 目录 / 内容骨架 / 致谢：format_cover_slide / format_toc_slide / setup_tju_content_slide / format_thanks_slide
- 结构化组件：add_joined_card / add_clean_tju_table / add_omml_equation（含 U+0302 重音帽）
- 自包含绘图：add_flow_pipeline / add_feedback_loop / TjuDiagram / tju_plot_style
- 全量守门与审计：enforce_min_18pt_all_slides / save_staged_pptx / audit_pptx / audit_content / audit_image_palette

运行：python examples/generate_tju_opening_ppt.py
"""

import os
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, "../scripts")))

import numpy as np
from pptx import Presentation
from pptx.enum.text import PP_ALIGN

from ppt_builder_helpers import (
    C_BLACK, C_BLUE, C_BLUE_LIGHT, C_RED, C_WHITE,
    add_clean_tju_table, add_dashed_box, add_image_fit, add_joined_card,
    add_omml_equation, add_solid_box, add_text_block,
    enforce_min_18pt_all_slides, format_cover_slide, format_thanks_slide,
    format_toc_slide, m_acc, m_r, m_sub, m_sup, rearrange_slides,
    save_staged_pptx, setup_tju_content_slide,
)
from tju_native_diagrams import add_feedback_loop, add_flow_pipeline
from tju_svg_diagram import TjuDiagram
from tju_plot_style import apply_tju_style, new_figure, save_tju_figure, tju_line_styles
from tju_audit import audit_content, audit_image_palette, audit_pptx

TEMPLATE_PATH = os.path.abspath(os.path.join(HERE, "../references/tju_template.pptx"))
OUTPUT_PATH = os.path.abspath(os.path.join(HERE, "example_opening_presentation.pptx"))

TOC_ITEMS = [
    "01  研究背景与现状",
    "02  研究内容与验证方案",
    "03  进度计划与预期成果",
]


def build_scene_diagram(png_path):
    """P3 左栏示意图：按放置框 5.82 × 4.46 in 1:1 尺寸绘制"""
    d = TjuDiagram(5.82, 4.46, base_pt=16)
    d.group(0.04, 0.04, 5.74, 1.85, "作业环境与动态扰动源")
    d.node("load", 0.22, 0.68, 1.56, 0.96, "载荷变化\n质量 / 高度", role="input")
    d.node("path", 2.12, 0.68, 1.56, 0.96, "窄巷曲率\n约束边界", role="input")
    d.node("dist", 4.02, 0.68, 1.56, 0.96, "时变扰动\n摩擦 / 惯量", role="output")

    d.group(0.04, 2.18, 5.74, 2.22, "控制闭环目标")
    d.node("est", 0.22, 2.90, 1.56, 1.15, "扰动观测器\n在线估计", role="core")
    d.node("ctrl", 2.12, 2.90, 1.56, 1.15, "模型预测控制\n滚动求解", role="core")
    d.node("act", 4.02, 2.90, 1.56, 1.15, "底盘与举升\n平稳执行", role="plain")

    d.edge("load", "est")
    d.edge("path", "ctrl")
    d.edge("dist", "ctrl", src_side="bottom", dst_side="top", route="vhv", color="red")
    d.edge("est", "ctrl")
    d.edge("ctrl", "act")
    d.save(png_path)
    return png_path


def build_expected_plot(png_path):
    """P8 右栏示意验证曲线：按放置框 5.82 × 3.00 in 1:1 尺寸绘制"""
    apply_tju_style(base_pt=16, tick_pt=14, legend_pt=14)
    rng = np.random.default_rng(7)
    t = np.arange(0, 31, 2)
    fig, ax = new_figure(5.82, 3.00)
    series = [
        ("本文自适应方案", 0.045, 0.22),
        ("固定权重基线", 0.105, 0.24),
        ("常规反馈基线", 0.160, 0.26),
    ]
    for (name, base, amp), st in zip(series, tju_line_styles(markersize=5)):
        y = base + amp * np.exp(-t / 8.0) + rng.normal(0, 0.003, len(t))
        ax.plot(t, y, label=name, **st)
    ax.set_xlabel("运行时间 / s")
    ax.set_ylabel("横向跟踪误差 / m")
    ax.set_ylim(0, 0.48)
    ax.legend(ncol=1, loc="upper right")
    save_tju_figure(fig, png_path)
    return png_path


def build_example_presentation():
    prs = Presentation(TEMPLATE_PATH)
    # 从 26 页模板精准选取 10 页（0-based）：
    # 0:封面 | 1:目录01 | 2,3,4:内容页 | 10:目录02 | 11,12,13:内容页 | 25:致谢
    rearrange_slides(prs, [0, 1, 2, 3, 4, 10, 11, 12, 13, 25])
    assert len(prs.slides) == 10

    with tempfile.TemporaryDirectory() as tmpdir:
        scene_png = build_scene_diagram(os.path.join(tmpdir, "scene.png"))
        plot_png = build_expected_plot(os.path.join(tmpdir, "curve.png"))

        # ------------------------------------------------------------------ #
        # P1 封面
        # ------------------------------------------------------------------ #
        format_cover_slide(
            prs.slides[0],
            title_lines=["面向动态负载工况的移动机器人", "自适应协同控制研究"],
            meta_lines=["汇报人：张三    导师：李四 教授", "汇报日期：2026 年 11 月"],
        )

        # ------------------------------------------------------------------ #
        # P2 目录（高亮 01）
        # ------------------------------------------------------------------ #
        format_toc_slide(prs.slides[1], TOC_ITEMS, active_idx=0, page_num=2)

        # ------------------------------------------------------------------ #
        # P3 应用场景与核心概念界定
        # ------------------------------------------------------------------ #
        s3 = prs.slides[2]
        setup_tju_content_slide(
            s3, "研究背景及意义",
            "应用场景：高举升移动机器人在动态负载下的窄巷循迹控制",
            page_num=3,
        )
        add_dashed_box(s3, 0.45, 1.78, 6.12, 5.22)
        add_dashed_box(s3, 6.77, 1.78, 6.12, 5.22)

        add_text_block(s3, 0.60, 1.88, 5.82, 0.40, [("1. 动态负载作业闭环示意", C_BLUE, 19)])
        add_image_fit(s3, scene_png, 0.60, 2.36, 5.82, 4.46, border_color=None)

        add_text_block(s3, 6.92, 1.88, 5.82, 0.40, [("2. 核心概念界定", C_BLUE, 19)])
        add_joined_card(
            s3, 6.92, 2.38, 5.82, 0.44, 0.96,
            "概念 ①：时变集总扰动",
            ["载荷质量与举升高度变化引发的惯量摄动及地面摩擦不确定性。"],
            header_bg=C_BLUE, body_bg=C_BLUE_LIGHT, border_color=C_BLUE,
        )
        add_joined_card(
            s3, 6.92, 3.94, 5.82, 0.44, 0.96,
            "概念 ②：模型预测控制（MPC）",
            ["在有限预测时域内滚动求解带约束优化问题以输出控制量。"],
            header_bg=C_BLUE, body_bg=C_WHITE, border_color=C_BLUE,
        )
        add_joined_card(
            s3, 6.92, 5.50, 5.82, 0.44, 0.96,
            "概念 ③：工况相关安全包线",
            ["不同载荷与举升高度下防倾覆所允许的速度与角速度上界。"],
            header_bg=C_RED, body_bg=C_BLUE_LIGHT, border_color=C_RED,
        )

        # ------------------------------------------------------------------ #
        # P4 难点分析与核心科学问题
        # ------------------------------------------------------------------ #
        s4 = prs.slides[3]
        setup_tju_content_slide(
            s4, "研究背景及意义",
            "问题凝练：固定参数控制器难以兼顾动态负载下的精度与安全",
            page_num=4,
        )
        add_solid_box(
            s4, 0.45, 1.78, 6.12, 0.48, ["现有方法局限：固定权重 + 静态边界"],
            fill_color=C_RED, line_color=C_RED, text_color=C_WHITE, size_pt=19,
        )
        add_flow_pipeline(
            s4, 0.55, 2.42, 5.92, 0.92,
            ["负载突变", "固定权重失配", "误差放大"],
            roles=["input", "plain", "output"], gap=0.38,
        )
        add_solid_box(
            s4, 0.55, 3.56, 5.92, 0.84,
            [[("• 滞后问题：", C_BLUE, 18), ("常规观测器难以预测快速时变扰动。", C_BLACK, 18)]],
            fill_color=C_WHITE, line_color=C_BLUE, align=PP_ALIGN.LEFT,
        )
        add_solid_box(
            s4, 0.55, 4.54, 5.92, 0.84,
            [[("• 效率矛盾：", C_BLUE, 18), ("按最坏工况保守限速会降低效率。", C_BLACK, 18)]],
            fill_color=C_WHITE, line_color=C_BLUE, align=PP_ALIGN.LEFT,
        )

        add_solid_box(
            s4, 6.77, 1.78, 6.12, 0.48, ["本文思路：扰动预测 + 自适应 MPC"],
            fill_color=C_BLUE, line_color=C_BLUE, text_color=C_WHITE, size_pt=19,
        )
        add_joined_card(
            s4, 6.77, 2.42, 6.12, 0.44, 0.92,
            "高阶观测与多步补偿",
            ["在线估计扰动及其变化率，并在预测时域内实施前馈补偿。"],
            header_bg=C_BLUE, body_bg=C_BLUE_LIGHT, border_color=C_BLUE,
        )
        add_joined_card(
            s4, 6.77, 3.96, 6.12, 0.44, 0.96,
            "工况自适应权重与动态包线",
            ["结合强化学习（RL）在线调节目标权重，约束随载荷更新。"],
            header_bg=C_RED, body_bg=C_WHITE, border_color=C_RED,
        )

        add_joined_card(
            s4, 0.45, 5.62, 12.43, 0.44, 0.90,
            "本课题核心科学问题",
            ["在载荷质量与举升高度大范围变化条件下，如何协同扰动多步预测、工况安全约束与在线权重调节？"],
            header_bg=C_BLUE, body_bg=C_BLUE_LIGHT, border_color=C_BLUE,
        )

        # ------------------------------------------------------------------ #
        # P5 国内外研究现状与局限
        # ------------------------------------------------------------------ #
        s5 = prs.slides[4]
        setup_tju_content_slide(
            s5, "研究现状",
            "相关研究：扰动补偿与参数自适应分别开展，缺乏动态约束协同",
            page_num=5,
        )
        add_dashed_box(s5, 0.45, 1.78, 6.12, 5.22)
        add_dashed_box(s5, 6.77, 1.78, 6.12, 5.22)

        add_text_block(s5, 0.60, 1.88, 5.82, 0.40, [("1. 扰动观测与前馈补偿", C_BLUE, 19)])
        add_flow_pipeline(
            s5, 0.60, 2.42, 5.82, 0.88,
            ["跟踪误差", "扩张状态观测", "单步补偿"],
            roles=["input", "core", "plain"], gap=0.32,
        )
        add_joined_card(
            s5, 0.60, 3.58, 5.82, 0.44, 1.24,
            "代表思路与主要局限",
            [
                [("• 思路：", C_BLUE, 18), ("以扩张观测器估计外部集总扰动。", C_BLACK, 18)],
                [("• 局限：", C_RED, 18), ("假设扰动常值，大曲率下滞后明显。", C_BLACK, 18)],
            ],
            header_bg=C_BLUE, body_bg=C_WHITE, border_color=C_BLUE,
        )
        add_solid_box(
            s5, 0.60, 5.56, 5.82, 1.20,
            [[("借鉴：", C_BLUE, 18), ("扩展至高阶导数预测，覆盖完整预测时域。", C_BLACK, 18)]],
            fill_color=C_BLUE_LIGHT, line_color=C_BLUE, align=PP_ALIGN.LEFT,
        )

        add_text_block(s5, 6.92, 1.88, 5.82, 0.40, [("2. 学习驱动的 MPC 参数调节", C_BLUE, 19)])
        add_flow_pipeline(
            s5, 6.92, 2.42, 5.82, 0.88,
            ["运行状态", "策略网络", "权重更新"],
            roles=["input", "core", "output"], gap=0.32,
        )
        add_joined_card(
            s5, 6.92, 3.58, 5.82, 0.44, 1.24,
            "代表思路与主要局限",
            [
                [("• 思路：", C_BLUE, 18), ("以 RL 策略随工况动态调节权重。", C_BLACK, 18)],
                [("• 局限：", C_RED, 18), ("多假设恒定载荷，未耦合抗倾覆约束。", C_BLACK, 18)],
            ],
            header_bg=C_BLUE, body_bg=C_WHITE, border_color=C_BLUE,
        )
        add_solid_box(
            s5, 6.92, 5.56, 5.82, 1.20,
            [[("借鉴：", C_RED, 18), ("将载荷与举升高度显式纳入状态与硬约束。", C_BLACK, 18)]],
            fill_color=C_WHITE, line_color=C_RED, align=PP_ALIGN.LEFT,
        )

        # ------------------------------------------------------------------ #
        # P6 章节过渡（高亮 02）
        # ------------------------------------------------------------------ #
        format_toc_slide(prs.slides[5], TOC_ITEMS, active_idx=1, page_num=6)

        # ------------------------------------------------------------------ #
        # P7 研究内容：公式与闭环方案
        # ------------------------------------------------------------------ #
        s7 = prs.slides[6]
        setup_tju_content_slide(
            s7, "研究内容与方案",
            "核心方法：高阶扰动多步预测与工况自适应闭环求解",
            page_num=7,
        )
        add_dashed_box(s7, 0.45, 1.78, 6.12, 5.22)
        add_dashed_box(s7, 6.77, 1.78, 6.12, 5.22)

        add_text_block(s7, 0.60, 1.88, 5.82, 0.40, [("1. 滚动优化目标函数", C_BLUE, 19)])
        eq_xml = (
            m_r("min ", italic=False)
            + m_r("J = ", italic=False)
            + m_r("‖", italic=False)
            + m_sub(m_r("x"), m_r("k"))
            + m_r(" − ", italic=False)
            + m_sub(m_r("x"), m_r("ref", italic=False))
            + m_sup(m_sub(m_r("‖", italic=False), m_r("Q", color="00337F")), m_r("2", italic=False))
            + m_r(" + ", italic=False)
            + m_r("‖", italic=False)
            + m_sub(m_acc(m_r("d"), ch="\u0302"), m_r("k", italic=False))
            + m_sup(m_r("‖", italic=False), m_r("2", italic=False))
        )
        add_omml_equation(
            s7, 0.60, 2.38, 5.82, 0.78, eq_xml,
            jc="center", fill_color=C_BLUE_LIGHT, line_color=C_BLUE,
        )
        add_text_block(
            s7, 0.60, 3.30, 5.82, 3.40,
            [
                [("• 状态项 x_{k}：", C_BLUE, 18), ("车辆位姿与速度跟踪偏差。", C_BLACK, 18)],
                [("• 权重矩阵 Q：", C_BLUE, 18), ("由 RL 随载荷与曲率在线调节。", C_BLACK, 18)],
                [("• 扰动预测 d_{k}：", C_RED, 18), ("高阶观测外推，前馈抵消失配。", C_BLACK, 18)],
                [("• 安全约束：", C_RED, 18), ("控制量受载荷包线实时限幅。", C_BLACK, 18)],
            ],
        )

        add_text_block(s7, 6.92, 1.88, 5.82, 0.40, [("2. 训练与闭环迭代机制", C_BLUE, 19)])
        add_feedback_loop(
            s7, 6.92, 2.50, 5.82, 2.15,
            ["多工况采样", "策略更新", "闭环校核"],
            "误差反馈：更新权重策略与边界参数",
            gap=0.30,
        )
        add_solid_box(
            s7, 6.92, 5.08, 5.82, 1.66,
            [
                [("可行性保障：", C_BLUE, 18), ("底层 MPC 求解器保留硬安全约束。", C_BLACK, 18)],
                [("解耦设计：", C_RED, 18), ("策略仅调权重，不直接输出底层力矩。", C_BLACK, 18)],
            ],
            fill_color=C_BLUE_LIGHT, line_color=C_BLUE, align=PP_ALIGN.LEFT,
        )

        # ------------------------------------------------------------------ #
        # P8 验证方案与预期指标
        # ------------------------------------------------------------------ #
        s8 = prs.slides[7]
        setup_tju_content_slide(
            s8, "研究内容与方案",
            "验证方案：多载荷同条件基线对比与跟踪收敛预期",
            page_num=8,
        )
        add_dashed_box(s8, 0.45, 1.78, 6.12, 5.22)
        add_dashed_box(s8, 6.77, 1.78, 6.12, 5.22)

        add_text_block(s8, 0.60, 1.88, 5.82, 0.40, [("1. 同条件对比基线", C_BLUE, 19)])
        add_clean_tju_table(
            s8, 0.60, 2.38, 5.82, 3.12,
            col_widths=[0.95, 2.05, 2.82],
            rows_data=[
                ["编号", "对比方法", "检验目的"],
                ["B1", "常规反馈基线", "基础跟踪下界"],
                ["B2", "固定权重 MPC", "静态参数失配边界"],
                ["B3", "去扰动预测", "多步补偿独立增益"],
                ["Ours", "本文自适应方案", "精度、效率与安全"],
            ],
            header_pt=18, body_pt=18, badge_col0=True,
        )
        add_solid_box(
            s8, 0.60, 5.70, 5.82, 1.10,
            [[("指标：", C_RED, 18), ("均方根误差（RMSE）、耗时与违约次数。", C_BLACK, 18)]],
            fill_color=C_WHITE, line_color=C_RED, align=PP_ALIGN.LEFT,
        )

        add_text_block(s8, 6.92, 1.88, 5.82, 0.40, [("2. 预期跟踪误差收敛（示意数据）", C_BLUE, 19)])
        add_image_fit(s8, plot_png, 6.92, 2.40, 5.82, 3.10, border_color=None)
        add_solid_box(
            s8, 6.92, 5.70, 5.82, 1.10,
            [[("统计口径：", C_BLUE, 18), ("覆盖空载/半载/满载及未见组合，3 随机种子。", C_BLACK, 18)]],
            fill_color=C_BLUE_LIGHT, line_color=C_BLUE, align=PP_ALIGN.LEFT,
        )

        # ------------------------------------------------------------------ #
        # P9 进度计划与预期成果
        # ------------------------------------------------------------------ #
        s9 = prs.slides[8]
        setup_tju_content_slide(
            s9, "研究计划及预期成果",
            "分阶段推进计划（2026.11–2027.10）与预期成果",
            page_num=9,
        )
        add_clean_tju_table(
            s9, 0.45, 1.80, 12.43, 3.10,
            col_widths=[2.65, 6.18, 3.60],
            rows_data=[
                ["时间阶段", "主要研究任务", "阶段交付物"],
                ["2026.11–2027.01", "搭建动态负载仿真环境与基线控制器", "仿真平台与基线数据"],
                ["2027.02–2027.04", "完成高阶扰动预测与安全包线建模", "观测器与约束模块"],
                ["2027.05–2027.07", "开展自适应权重训练、消融与实车验证", "对比数据与实车记录"],
                ["2027.08–2027.10", "整理实验边界与结论，撰写学位论文", "硕士学位论文初稿"],
            ],
            header_pt=19, body_pt=18,
        )
        add_joined_card(
            s9, 0.45, 5.18, 12.43, 0.44, 1.28,
            "预期成果",
            [
                "• 算法与验证：形成动态负载协同控制模块，相较固定权重基线降低跟踪 RMSE 且满足安全约束；",
                "• 论文与归档：完成多工况仿真与实车对照验证，撰写硕士学位论文 1 篇并归档代码。",
            ],
            header_bg=C_BLUE, body_bg=C_BLUE_LIGHT, border_color=C_BLUE,
        )

        # ------------------------------------------------------------------ #
        # P10 致谢
        # ------------------------------------------------------------------ #
        format_thanks_slide(prs.slides[9])

        # 全量守门 + 暂存写回
        enforce_min_18pt_all_slides(prs)
        out = save_staged_pptx(prs, OUTPUT_PATH)
        print(f"Saved example PPTX: {out}")

        # OpenXML + 内容 + 调色板审计
        skip = {1, 2, 6, 10}
        issues = audit_pptx(out, skip_slides=skip)
        print(f"[audit pptx] {len(issues)} issue(s)")
        for iss in issues:
            print("  -", iss)
        assert len(issues) == 0, f"OpenXML audit failed: {issues}"

        for img in (scene_png, plot_png):
            r = audit_image_palette(img)
            print(f"[audit image] {os.path.basename(img)} ok={r['ok']} off={r['off_ratio']:.4f}")
            assert r["ok"], f"Off-palette image: {r}"

        warns = audit_content(out, skip_slides=skip)
        print(f"[audit content] {len(warns)} warning(s)")
        for w in warns:
            print("  ~", w)


if __name__ == "__main__":
    build_example_presentation()
