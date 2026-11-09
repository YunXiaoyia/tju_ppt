# `tju-ppt` — 天津大学课题组学术汇报 PPT 自动生成与优化 Skill

面向 Claude Code / Gemini / Antigravity / Codex 的智能体技能（Agent Skill），用于基于课题组 16:9 PowerPoint 模板（`.pptx`）自动生成或优化**硕士/博士开题汇报、中期考核与学位答辩 PPT**。

---

## ✨ 核心特性与排版铁律

1. **严格中英分设字体与字号下限（$\ge 18\text{ pt}$）**
   - 通过底层 OpenXML（`a:ea`, `a:latin`, `a:cs`, `a:sym`）强制所有中文使用 **`微软雅黑` 加粗**，所有英文、数字与数学符号使用 **`Times New Roman` 加粗**。
   - 全局字号**不小于 $18\text{ pt}$**（正文/图注/表格/公式注释 $18\text{ pt}$，栏目标题 $19\text{--}20\text{ pt}$，副标题 $21\text{--}22\text{ pt}$，顶部导航标题 $26\text{--}28\text{ pt}$），内置 `enforce_min_18pt_all_slides()` 自动扫描兜底。
2. **四色学术规范（主色调深蓝）**
   - 仅使用 **深蓝 `#00337F`**（主色调：标题、核心流程框、表头）、**纯黑 `#000000`**（正文、推导与图注）、**纯白 `#FFFFFF`**（背景与反白字）、**强调红 `#C00000`**（核心物理约束、关键判据、C 阶段高亮框），辅以浅蓝填充 `#EAF1FB` 作次级框与斑马纹表格底色。
3. **PowerPoint 原生可编辑插入公式（`<a14:m>` OMML）**
   - 杜绝低清公式截图，直接注入 `<a14:m><m:oMathPara><m:oMath>...</m:oMath></m:oMathPara></a14:m>` 原生 PowerPoint 公式对象，并紧跟 $18\text{ pt}$ 符号物理意义注释。
   - 内置 LibreOffice Impress (`StarMath`) 渲染避坑规则，彻底消除 `¿` 公式解析报错。
4. **对标课题组三种经典双栏虚线框版式**
   - **版式 1（机制建模 + 原生公式 + 框图 + 论文实证图）**：`references/reference_layouts/ref_layout_1_mechanism_formula.jpg`
   - **版式 2（分层感知管线 + 三步流程 + 最小二乘配准公式）**：`references/reference_layouts/ref_layout_2_perception_icp.jpg`
   - **版式 3（A → B → C 阶梯递进消融框 + 损失/奖励数学建模）**：`references/reference_layouts/ref_layout_3_abc_ladder.jpg`
5. **版面留白与渲染自检闭环**
   - 去除底部冗余“承上启下”文字，副标题固定于 `top = 0.95 in`，双栏虚线框/表格整体下移至 `top = 1.78–1.90 in`（留出 $0.35\text{--}0.40\text{ in}$ 清晰留白），小标题与正文之间自动插入段前/段后呼吸间距（`Pt(4)`）。
   - 强制通过 `libreoffice --headless --convert-to pdf` + `pdftoppm -png` 逐页渲染 PNG 视觉检查，确保零文字重叠、零越界。

---

## 📂 仓库目录结构

```text
tju_ppt/
├── SKILL.md                                 # Agent Skill 核心指令与避坑规范
├── README.md                                # 仓库说明与快速上手指南
├── scripts/
│   └── ppt_builder_helpers.py               # python-pptx + OpenXML 排版与原生公式辅助库
├── examples/
│   └── generate_tju_opening_ppt.py          # 26 页硕士开题汇报 PPT 完整生成示例脚本
└── references/
    ├── tju_template.pptx                    # 天津大学课题组 16:9 标准控制模板
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

---

## 💡 推荐调用提示词（Prompt 模板）

在安装了 `tju-ppt` 技能后，您可以直接向 Agent 发送如下指令：

> 请使用 `tju-ppt` 技能，基于工作区内的开题报告草稿与论文库，结合课题组模板生成一份开题汇报 PPT：
> 1. 字体方面：中文使用微软雅黑加粗，英文与数字使用 Times New Roman 加粗；字号不小于 18pt（图片内文字除外）；
> 2. 颜色方面：仅使用黑、白、红（强调部分）、蓝（主色调，用于目录与标题）四种颜色；
> 3. 图片与公式：优先从本地论文 PDF 提取高清架构图与实验图；公式使用 PowerPoint 原生插入公式（`<a14:m>`）并附符号说明；
> 4. 整体排版：参照参考图的双栏虚线框版式，去掉底部承上启下，注意副标题与虚线框、小标题与正文之间留出适当间距，生成后渲染 PNG 逐页检查确保零重叠。
