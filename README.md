# `tju-ppt` — 天津大学课题组学术汇报 PPT 自动生成与优化 Skill

面向 Claude Code / Gemini / Antigravity / Codex 的智能体技能（Agent Skill），用于基于课题组 16:9 PowerPoint 模板（`.pptx`）自动生成、审计与优化**硕士/博士开题汇报、中期考核与学位答辩 PPT**。

---

## ✨ 核心特性与规范

1. **严格中英分设字体与字号下限（$\ge 18\text{ pt}$）**
   - 通过底层 OpenXML（`a:ea`, `a:latin`, `a:cs`, `a:sym`）强制所有中文使用 **`微软雅黑` 加粗**，所有英文、数字与数学符号使用 **`Times New Roman` 加粗**。
   - 全局字号**不小于 $18\text{ pt}$**（正文/图注/表格/公式 $18\text{ pt}$，栏目标题 $19\text{--}20\text{ pt}$，副标题 $20\text{--}22\text{ pt}$，顶部导航标题 $26\text{--}28\text{ pt}$），内置 `enforce_min_18pt_all_slides()` 覆盖普通文本、表格与 OMML 公式自动兜底。
2. **四色学术规范（主色调深蓝）**
   - 仅使用 **深蓝 `#00337F`**（主色调：标题、核心流程框、表头）、**纯黑 `#000000`**（正文、推导与图注）、**纯白 `#FFFFFF`**（背景与反白字）、**强调红 `#C00000`**（核心物理约束、关键判据、强调框），辅以浅蓝填充 `#EAF1FB` 作次级框与斑马纹表格底色。
3. **PowerPoint 原生可编辑公式（`<a14:m>` OMML）与行内真实上下标**
   - 避免低清公式截图与正文中裸露的 `x_k^ref`、`Q_k` 代码下划线：行间公式直接注入 `<a14:m>` 原生对象（重音帽默认使用组合符 `U+0302` 避免压字）；正文、卡片与表格支持 `_{sub}` / `^{sup}` 自动转为 OpenXML `baseline` 真实上下标（同时保持 `sz >= 1800` 合规）。
   - 内置 LibreOffice Impress (`StarMath`) 与原生 PowerPoint 双环境渲染避坑规则，规避公式渲染异常。
4. **干净单母版模板与可复用组件库（`tju_template.pptx` + `ppt_builder_helpers.py`）**
   - `references/tju_template.pptx` 为精简清理后的**单母版、4 标准版式、4 页示例骨架**（`标题幻灯片 / 1_自定义版式 / 1_仅标题 / 空白`，内置 TJU 四色与中英主题字体，0 审计问题，体积约 1.5 MB）。
   - 内置 `clear_all_slides` + `new_cover_slide` / `new_toc_slide` / `new_content_slide` / `new_thanks_slide` 按需新建任意页数，同时支持存量课件原地编辑（`rearrange_slides`、`setup_tju_content_slide` 自动清理 `mc:AlternateContent` 幽灵残留）与暂存写回（`save_staged_pptx`）。
   - 内置连体双层卡片（`add_joined_card`）、严格四色表格/徽标表（`add_clean_tju_table`）与单形状等比图框（`add_image_fit`），拒绝满屏纯项目符号。
   - 内部双栏虚线框排布参考：`ref_layout_1_mechanism_formula.jpg`、`ref_layout_2_perception_icp.jpg`、`ref_layout_3_abc_ladder.jpg`（仅参考其双栏虚线框内组件层次，顶栏与页码以本规范为准）。
5. **版面留白与双平台渲染自检闭环**
   - 去除底部冗余“承上启下”文字，副标题与虚线框之间留出清晰呼吸白边，小标题与正文自动插入段前/段后间距。
   - 支持 **Windows 原生 PowerPoint COM + PyMuPDF**（含暂存文件防锁机制）与 **Linux Headless LibreOffice + pdftoppm** 导出高清 PNG 逐页核验，确保零重叠、零孤立标点换行。
6. **科研流程图 / 技术路线图 / 数据图（自包含，无需安装其他 Skill）**
   - **原生可编辑**（`tju_native_diagrams.py`）：N 步流程、闭环反馈、分层系统架构、正交连接器，导师可在 PowerPoint 中直接改字。
   - **复杂示意图引擎**（`tju_svg_diagram.py` · `TjuDiagram`）：Python API / JSON spec / 自动分层布局 → 严格校验（溢出、重叠、穿框、出界）→ SVG + 2× PNG，中西文逐段分设字体。
   - **答辩投影级数据图**（`tju_plot_style.py`）：TJU 四色 + 线型/阴影线区分多组、大字号、按 PPT 放置尺寸作图。
   - **自动审计**（`tju_audit.py`）：OpenXML 全量审计（普通文本 + OMML 公式、严格四色 + 主题色回退、裸数学下标、包围盒重叠）+ 自绘图片调色板审计 + `--content` 内容告警。
   - 路由规则与自检清单见 `references/diagram_routing.md`，完整演示见 `examples/generate_diagram_showcase.py` 与 `examples/generate_tju_opening_ppt.py`。
7. **叙事与内容规范（面向跨方向评委）**
   - 固定叙事弧：背景/场景 → 特殊性与难点 → 抽象科学问题 → 相关研究与局限 → 我的工作（先总图）→ 验证方案 → 创新/可行性/计划；副标题连读即完整论证线。
   - 术语先定义后使用（降门槛不降格调，禁止口水话/俗语比喻）、单页一个结论、按讲稿时长控制字数（40–60 s 页 ≈ 150–200 字）、每页配视觉元素、论断与数字可溯源且推理成立。
   - `tju_audit.py --content` 自动提示：文字密度、纯文字页（原生流程图放行）、多行溢出估计、绝对化表述、口语化表述、缩写首次出现；详见 `references/narrative_principles.md`。

---

## 📂 仓库目录结构

```text
tju_ppt/
├── SKILL.md                                 # Agent Skill 核心指令与避坑规范
├── README.md                                # 仓库说明与快速上手指南
├── scripts/
│   ├── ppt_builder_helpers.py               # python-pptx + OpenXML 骨架、组件、原生公式与合规守门
│   ├── tju_native_diagrams.py               # 原生可编辑流程图 / 闭环 / 分层架构 / 连接器
│   ├── tju_svg_diagram.py                   # TjuDiagram 科研示意图引擎（校验 → SVG/PNG）
│   ├── tju_plot_style.py                    # Matplotlib 答辩投影样式
│   └── tju_audit.py                         # OpenXML 审计 + 内容告警 + 图片调色板审计
├── examples/
│   ├── generate_tju_opening_ppt.py          # 10 页开题汇报全链路合规范例脚本（输出 example_opening_presentation.pptx）
│   ├── example_opening_presentation.pptx    # 0 审计问题的通用开题汇报示例成品
│   └── generate_diagram_showcase.py         # 四类绘图路线演示 + 自动审计
└── references/
    ├── tju_template.pptx                    # 天津大学 16:9 纯净单母版模板（4 版式 + 4 页合规骨架，0 审计问题）
    ├── diagram_routing.md                   # 绘图路由、已知限制与视觉自检清单
    ├── narrative_principles.md              # 叙事弧、语体、术语、密度、论断规范与开题完整性清单
    ├── tju_mermaid_theme.json               # 可选：Mermaid TJU 主题
    └── reference_layouts/                   # 三张内部双栏虚线框视觉层次参考图
        ├── ref_layout_1_mechanism_formula.jpg
        ├── ref_layout_2_perception_icp.jpg
        └── ref_layout_3_abc_ladder.jpg
```

---

## 🚀 安装方法

将本仓库克隆到您的 Agent Skills 目录即可自动生效：

```bash
# Antigravity / Gemini CLI
git clone https://github.com/YunXiaoyia/tju_ppt.git ~/.gemini/config/skills/tju-ppt

# Claude Code / Codex / 通用 Agent Skills 目录
git clone https://github.com/YunXiaoyia/tju_ppt.git ~/.agents/skills/tju-ppt
```

依赖：`pip install python-pptx pillow numpy lxml`；绘制复杂示意图另需 `cairosvg`，数据图需 `matplotlib`（建议 >= 3.6，以便同一标签内中西文分设字体）。

---

## 💡 推荐调用提示词（Prompt 模板）

在安装了 `tju-ppt` 技能后，您可以直接向 Agent 发送如下指令：

> 请使用 `tju-ppt` 技能，基于工作区内的开题报告草稿与论文库，结合课题组模板生成一份开题汇报 PPT：
> 1. 叙事与语体：面向跨方向评审老师，按“背景与概念界定 → 难点与科学问题 → 相关研究与局限 → 本文方法 → 验证方案 → 进度与预期成果”组织，保持规范学术语体；
> 2. 字体与配色：中文微软雅黑加粗、英文与数字 Times New Roman 加粗，字号不小于 18pt；仅用黑、白、红（强调）、蓝（主色调，含浅蓝辅色）；
> 3. 图片与公式：优先从本地论文 PDF 提取原图或用内置绘图路线生成示意图与数据图；公式使用原生 OMML（`<a14:m>`）并附符号说明；
> 4. 审核与渲染：运行 `tju_audit.py --content` 确保 OpenXML 0 问题，并渲染 PNG 逐页检查确保零重叠、零文字溢出。
