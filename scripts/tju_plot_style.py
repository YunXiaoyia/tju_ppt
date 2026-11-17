#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TJU-PPT 答辩投影级 Matplotlib 数据图样式 (tju_plot_style.py)

- apply_tju_style()  : 全局 rcParams（TJU 四色 + Times New Roman / 微软雅黑 + 投影大字号 + 去上右边框）
- new_figure(w, h)   : 按 PPT 放置框尺寸（英寸）建图，保证图中字号 ≈ PPT 实际字号
- tju_line_styles()  : 线图样式循环（颜色仅 3 种，靠线型 + 标记区分第 4 条及以后曲线）
- TJU_BAR_STYLES     : 柱状图样式（填充 + 描边 + 阴影线），超过 3 组仍可区分且不越色
- TJU_CMAP_SEQ / TJU_CMAP_DIV : 热力图连续 / 发散色图（白→浅蓝→深蓝；深蓝→白→红）
- save_tju_figure()  : 逐个 Text 按内容指定中/英字体（兼容 matplotlib < 3.6 无字形回退的问题），
                       检查最小字号，导出 300 dpi PNG（可选同时导出 PDF/SVG）

用法：
    from tju_plot_style import apply_tju_style, new_figure, tju_line_styles, save_tju_figure
    apply_tju_style()
    fig, ax = new_figure(5.6, 3.6)              # 与 PPT 中的放置框同尺寸
    for (x, y, lab), st in zip(series, tju_line_styles()):
        ax.plot(x, y, label=lab, **st)
    ax.set_xlabel("迭代次数"); ax.set_ylabel("RMSE / m")
    save_tju_figure(fig, "fig_rmse.png")
"""

import re
import warnings
from itertools import cycle

import matplotlib
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.colors import LinearSegmentedColormap

BLUE, BLUE_LIGHT, RED, BLACK, WHITE = "#00337F", "#EAF1FB", "#C00000", "#000000", "#FFFFFF"
_CJK = re.compile(r"[\u2e80-\u9fff\uf900-\ufaff\uff00-\uffef\u3000-\u303f]")
_MPL_VER = tuple(int(x) for x in re.findall(r"\d+", matplotlib.__version__)[:2])

TJU_CMAP_SEQ = LinearSegmentedColormap.from_list("tju_seq", [WHITE, BLUE_LIGHT, BLUE])
TJU_CMAP_DIV = LinearSegmentedColormap.from_list("tju_div", [BLUE, WHITE, RED])

TJU_BAR_STYLES = [
    dict(color=BLUE, edgecolor=BLUE, linewidth=1.2),
    dict(color=WHITE, edgecolor=BLUE, hatch="//", linewidth=1.2),
    dict(color=RED, edgecolor=RED, linewidth=1.2),
    dict(color=BLUE_LIGHT, edgecolor=BLUE, linewidth=1.2),
    dict(color=WHITE, edgecolor=BLACK, hatch="xx", linewidth=1.2),
    dict(color=WHITE, edgecolor=RED, hatch="..", linewidth=1.2),
]

_LINE_STYLES = [
    dict(color=BLUE, linestyle="-", marker="o"),
    dict(color=RED, linestyle="--", marker="s"),
    dict(color=BLACK, linestyle="-.", marker="^"),
    dict(color=BLUE, linestyle=":", marker="D"),
    dict(color=RED, linestyle=":", marker="v"),
    dict(color=BLACK, linestyle="--", marker="x"),
]


def _first_available(cands):
    have = {f.name for f in font_manager.fontManager.ttflist}
    for c in cands:
        if c in have:
            return c
    return None


LATIN_FONT = _first_available(["Times New Roman", "Liberation Serif", "DejaVu Serif"]) or "serif"
CJK_FONT = _first_available(["Microsoft YaHei", "Noto Sans CJK SC", "Noto Sans CJK JP", "Source Han Sans SC",
                             "WenQuanYi Micro Hei", "SimHei"])


def tju_line_styles(markersize=7, linewidth=2.4, markevery=None):
    """返回无限循环的线图样式字典，可直接 **st 传给 ax.plot。"""
    for st in cycle(_LINE_STYLES):
        d = dict(st, linewidth=linewidth, markersize=markersize)
        if markevery:
            d["markevery"] = markevery
        yield d


def apply_tju_style(base_pt=18, tick_pt=16, legend_pt=16):
    """
    全局样式。默认按“图在 PPT 中以 1:1 尺寸放置”设计：轴标签 18 pt、刻度/图例 16 pt。
    （SKILL 规定图片内部文字不受 18 pt 约束，但仍建议 >= 14 pt 保证投影可读。）
    """
    if LATIN_FONT != "Times New Roman":
        warnings.warn(f"[tju-ppt] 未找到 Times New Roman，西文暂用 {LATIN_FONT}；Linux 可装 ttf-mscorefonts-installer。")
    if CJK_FONT is None:
        warnings.warn("[tju-ppt] 未找到中文字体，中文将显示为方块；请安装 微软雅黑 或 Noto Sans CJK SC。")
    family = [LATIN_FONT] + ([CJK_FONT] if CJK_FONT else [])
    plt.rcParams.update({
        "font.family": family if _MPL_VER >= (3, 6) else [LATIN_FONT],
        "font.weight": "bold",
        "font.size": base_pt,
        "axes.titlesize": base_pt,
        "axes.titleweight": "bold",
        "axes.labelsize": base_pt,
        "axes.labelweight": "bold",
        "xtick.labelsize": tick_pt,
        "ytick.labelsize": tick_pt,
        "legend.fontsize": legend_pt,
        "legend.frameon": False,
        "axes.edgecolor": BLACK,
        "axes.linewidth": 1.4,
        "axes.labelcolor": BLACK,
        "axes.spines.top": False,
        "axes.spines.right": False,
        "axes.grid": True,
        "grid.color": BLUE_LIGHT,
        "grid.linewidth": 1.2,
        "axes.axisbelow": True,
        "xtick.color": BLACK,
        "ytick.color": BLACK,
        "xtick.major.width": 1.3,
        "ytick.major.width": 1.3,
        "lines.linewidth": 2.4,
        "axes.prop_cycle": matplotlib.cycler(color=[BLUE, RED, BLACK]),
        "figure.facecolor": WHITE,
        "axes.facecolor": WHITE,
        "savefig.facecolor": WHITE,
        "mathtext.fontset": "custom",
        "mathtext.rm": f"{LATIN_FONT}",
        "mathtext.it": f"{LATIN_FONT}:italic:bold",
        "mathtext.bf": f"{LATIN_FONT}:bold",
        "axes.unicode_minus": False,
        "image.cmap": "tju_seq",
    })
    try:
        matplotlib.colormaps.register(TJU_CMAP_SEQ)
        matplotlib.colormaps.register(TJU_CMAP_DIV)
    except Exception:
        try:
            plt.register_cmap(cmap=TJU_CMAP_SEQ)
            plt.register_cmap(cmap=TJU_CMAP_DIV)
        except Exception:
            pass


def new_figure(width_in, height_in, nrows=1, ncols=1, **kw):
    """按 PPT 放置框尺寸建图（constrained_layout 自动避让标签）。"""
    return plt.subplots(nrows, ncols, figsize=(width_in, height_in), constrained_layout=True, **kw)


def _assign_fonts(fig):
    """逐个 Text 按内容指定字体：含中文 → CJK 字体（其拉丁字形也清晰），纯西文 → Times New Roman。"""
    from matplotlib.text import Text
    for t in fig.findobj(Text):
        s = t.get_text()
        if not s:
            continue
        if _CJK.search(s) and CJK_FONT:
            if _MPL_VER >= (3, 6):
                t.set_fontfamily([LATIN_FONT, CJK_FONT])
            else:
                t.set_fontfamily(CJK_FONT)
        else:
            t.set_fontfamily(LATIN_FONT)
        t.set_fontweight("bold")


def save_tju_figure(fig, path, dpi=300, extra_formats=(), min_pt=14):
    """统一导出：字体分配 → 最小字号检查 → PNG(300 dpi)，extra_formats 可加 ("pdf", "svg")。"""
    from matplotlib.text import Text
    _assign_fonts(fig)
    small = sorted({round(t.get_fontsize(), 1) for t in fig.findobj(Text)
                    if t.get_text() and t.get_visible() and t.get_fontsize() < min_pt})
    if small:
        warnings.warn(f"[tju-ppt] 图中存在小于 {min_pt} pt 的文字 {small}，投影可能看不清。")
    fig.savefig(path, dpi=dpi)
    stem = path.rsplit(".", 1)[0]
    for ext in extra_formats:
        fig.savefig(f"{stem}.{ext}")
    return path
