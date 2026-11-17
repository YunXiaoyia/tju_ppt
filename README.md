# `tju-ppt` — 天津大学课题组学术汇报 PPT 自动生成与优化 Skill

面向 Claude Code / Gemini / Antigravity / Codex 的智能体技能（Agent Skill），用于基于课题组 16:9 PowerPoint 模板（`.pptx`）自动生成或优化**硕士/博士开题汇报、中期考核与学位答辩 PPT**。

---

## ✨ 核心特性与排版铁律

1. **严格中英分设字体与字号下限（$\ge 18\text{ pt}$）**
   - 通过底层 OpenXML（`a:ea`, `a:latin`, `a:cs`, `a:sym`）强制所有中文使用 **`微软雅黑` 加粗**，所有英文、数字与数学符号使用 **`Times New Roman` 加粗**。
   - 全局字号**不小于 $18\text{ pt}$**（正文/图注/表格/公式注释 $18\text{ pt}$，栏目标题 $19\text{--}20\text{ pt}$，副标题 $21\text{--}22\text{ pt}$，顶部导航标题 $26\text{--}28\text{ pt}$），内置 `enforce_min_18pt_all_slides()` 自动扫描兜底。
2. **四色学术规范（主色调深蓝）**
   - 仅使用 **深蓝 `#00337F`**（主色调：标题、核心流程框、表头）、**纯黑 `#000000`**（正文、推导与图注）、**纯白 `#FFFFFF`**（背景与反白字）、**强调红 `#C00000`**（核心物理约束、关键判据、C 阶段高亮框），辅以浅蓝填充 `#EAF1FB` 作次级框与斑马纹表格底色。
3. **PowerPoint 原生可编辑公式（`<a14:m>` OMML）与行内真实上下标**
   - 杜绝低清公式截图与正文中裸露的 `M_t^tmp`、`L_pair` 代码下划线：行间公式直接注入 `<a14:m>` 原生对象（重音帽默认使用标准组合符 `U+0302` 避免压字）；正文、卡片与表头支持 `_{sub}` / `^{sup}` 自动转为 OpenXML `baseline` 真实上下标（同时保持 `sz >= 1800` 合规）。
   - 内置 LibreOffice Impress (`StarMath`) 与原生 PowerPoint 双环境渲染避坑规则，彻底消除 `¿` 公式报错。
4. **高审美学术版式与反“纯文字墙”规范**
   - **标准内容页骨架（`setup_tju_content_slide`）**：副标题左侧深蓝竖条锚点 + 顶部分隔细线 + 右下角统一页码，方便答辩评委记录页号。
   - **对标课题组经典双栏版式与结构化组件**：内置三框横向流程图、连体双层卡片、`2×4` 场景/指标交替网格、深蓝徽标消融表与单形状等比图框（`add_image_fit`），拒绝满屏纯项目符号。
   - 经典参考图：`ref_layout_1_mechanism_formula.jpg`、`ref_layout_2_perception_icp.jpg`、`ref_layout_3_abc_ladder.jpg`。
5. **版面留白与双平台渲染自检闭环**
   - 去除底部冗余“承上启下”文字，副标题与虚线框之间留出清晰呼吸白边，小标题与正文自动插入段前/段后间距。
   - 支持 **Windows 原生 PowerPoint COM + PyMuPDF**（含暂存文件防锁机制）与 **Linux Headless LibreOffice + pdftoppm** 导出高清 PNG 逐页核验，确保零重叠、零孤立标点换行。

6. **科研流程图 / 技术路线图 / 数据图（自包含，无需安装其他 Skill）**
   - **原生可编辑**（`tju_native_diagrams.py`）：N 步流程、闭环反馈、分层系统架构、正交连接器，导师可在 PowerPoint 中直接改字。
   - **复杂示意图引擎**（`tju_svg_diagram.py` · `TjuDiagram`）：Python API / JSON spec / 自动分层布局 → 严格校验（溢出、重叠、连线穿框）→ SVG + 2× PNG，中西文逐段分设字体。
   - **答辩投影级数据图**（`tju_plot_style.py`）：TJU 四色 + 线型/阴影线区分多组、大字号、按 PPT 放置尺寸作图。
   - **自动审计**（`tju_audit.py`）：OpenXML 字号/字体槽位/越色/重叠检查 + 自绘图片调色板检查。
   - 路由规则与自检清单见 `references/diagram_routing.md`，完整演示见 `examples/generate_diagram_showcase.py`。

7. **叙事与内容规范（面向跨方向评委）**
   - 固定叙事弧：背景/场景 → 特殊性与难点 → 抽象科学问题 → 别人的工作与局限 → 我的工作（先总图）→ 验证方案 → 创新/可行性/计划；副标题连读即故事线。
   - 术语先定义后使用（降门槛不降格调，禁止口水话/比喻）、单页一个结论、按讲稿时长控制字数（40–60 s 页 ≈ 150–200 字）、每页配图、论断与数字可溯源且推理成立。
   - `tju_audit.py --content` 自动提示：文字密度、无图内容页、文字溢出估计、绝对化表述、口语化表述、缩写首次出现；详见 `references/narrative_principles.md`。

---

## 📂 仓库目录结构

```text
tju_ppt/
├── SKILL.md                                 # Agent Skill 核心指令与避坑规范
├── README.md                                # 仓库说明与快速上手指南
├── scripts/
│   ├── ppt_builder_helpers.py               # python-pptx + OpenXML 排版与原生公式辅助库
│   ├── tju_native_diagrams.py               # 原生可编辑流程图 / 闭环 / 分层架构 / 连接器
│   ├── tju_svg_diagram.py                   # TjuDiagram 科研示意图引擎（校验 → SVG/PNG）
│   ├── tju_plot_style.py                    # Matplotlib 答辩投影样式
│   └── tju_audit.py                         # OpenXML 审计 + 图片调色板审计
├── examples/
│   ├── generate_tju_opening_ppt.py          # 开题汇报 PPT 生成示例脚本
│   └── generate_diagram_showcase.py         # 四类绘图路线演示 + 自动审计
└── references/
    ├── tju_template.pptx                    # 天津大学课题组 16:9 标准控制模板
    ├── diagram_routing.md                   # 绘图路由、已知限制与视觉自检清单
    ├── narrative_principles.md              # 叙事弧、术语、密度、论断规范与开题完整性清单
    ├── tju_mermaid_theme.json               # 可选：Mermaid TJU 主题
    └── reference_layouts/                   # 三张标准双栏排版参考图
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
> 1. 字体方面：中文使用微软雅黑加粗，英文与数字使用 Times New Roman 加粗；字号不小于 18pt（图片内文字除外）；
> 2. 颜色方面：仅使用黑、白、红（强调部分）、蓝（主色调，用于目录与标题）四种颜色；
> 3. 图片与公式：优先从本地论文 PDF 提取高清架构图与实验图；公式使用 PowerPoint 原生插入公式（`<a14:m>`）并附符号说明；
> 4. 整体排版：参照参考图的双栏虚线框版式，去掉底部承上启下，注意副标题与虚线框、小标题与正文之间留出适当间距，生成后渲染 PNG 逐页检查确保零重叠。
