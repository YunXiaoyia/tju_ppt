# `tju-ppt` — 天津大学课题组学术汇报 PPT Skill

面向 Antigravity / Gemini / Claude Code / Codex 的 Agent Skill：在课题组 16:9 模板上生成、审计、修改
**开题汇报、中期考核、学位答辩、组会** PPT 及配套讲稿。智能体的使用说明在 [`SKILL.md`](SKILL.md)，本文件只面向人。

## 设计
- **三层**：叙事（`references/story.md`）→ 页面配方（`references/page_recipes.md`）→ 代码构建（`scripts/tjuppt.py`）。
- **坐标来自网格**：内容统一落在 `0.45–12.88 in`，左右分栏用 `Region.split()`，图框尺寸由 `figure_slot()` 给出，
  图按槽位 1:1 绘制后与标题栏齐边。
- **放不下就报错**：引导句、红字结论、卡片、表格在生成时检查是否放得下（`LayoutError`），不缩字号。
- **审计闭环**：`tju_audit.py`（字体/配色/网格/齐边/红字/重叠 + 内容告警）、`tju_render.py`（渲染 PNG）、
  `tju_script_check.py`（讲稿语速与时长）。

## 目录
```text
SKILL.md                      智能体入口：流程、硬约束、判断规则
scripts/
  tjuppt.py                   一次性导入
  tju_theme.py                配色、字体、字号、页面网格、字宽估算（唯一来源）
  tju_layout.py               Region：split / vsplit / take_top …
  tju_text.py                 四槽位字体、行内上下标
  tju_components.py           图框、卡片、图文卡、表格及坐标式基础组件
  tju_omml.py                 原生可编辑公式
  tju_deck.py                 模板、Deck / ContentPage、页面骨架、保存
  tju_native_diagrams.py      原生可编辑流程图 / 闭环 / 分层架构
  tju_svg_diagram.py          TjuDiagram 示意图引擎（严格校验 → SVG/PNG）
  tju_plot_style.py           Matplotlib 答辩样式
  tju_audit.py  tju_render.py  tju_script_check.py   命令行工具
  ppt_builder_helpers.py      旧脚本兼容层
references/  story.md  page_recipes.md  figures.md  pptx_pitfalls.md  tju_template.pptx  tju_mermaid_theme.json
examples/    build_example_deck.py（12 页开题示例）  diagram_showcase.py（四条绘图路线）
tests/       test_tjuppt.py（契约测试）
```

## 安装与依赖
```bash
git clone https://github.com/YunXiaoyia/tju_ppt.git ~/.gemini/config/skills/tju-ppt
pip install python-pptx pillow numpy lxml cairosvg fonttools matplotlib
```
渲染需要 LibreOffice + poppler（Linux）或 PowerPoint + PyMuPDF（Windows）。

## 快速验证
```bash
python tests/test_tjuppt.py            # 契约测试
python examples/build_example_deck.py  # 生成示例开题 PPT、渲染 PNG 并打印审计结果
```

## 从 v1 迁移
旧脚本 `from ppt_builder_helpers import ...` 仍可运行。新脚本改用 `from tjuppt import *` 与 `Deck`，
用 `page.body.split()` 代替手写坐标，用 `add_figure_panel` 代替“标题条 + add_image_fit”，
用 `deck.content(..., lead=..., takeaway=...)` 代替手写引导句与红字结论。
