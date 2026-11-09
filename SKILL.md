---
name: tju-ppt
description: Use when creating, editing, or optimizing Tianjin University (TJU / 天津大学) or academic group opening report (开题汇报), mid-term, or thesis defense PPTX presentations with python-pptx, especially when enforcing strict font rules (微软雅黑 + Times New Roman >= 18pt), four-color palette, native editable OMML equations, real paper figures, and zero-overlap dual-column dashed-box layouts.
---

# TJU Academic Presentation (天津大学课题组学术汇报 PPT) Skill

## Overview

Build publication- and defense-grade academic PowerPoint (`.pptx`) presentations based on the Tianjin University (天津大学) 16:9 group template using `python-pptx` and OpenXML. Every slide enforces strict typography (`微软雅黑` + `Times New Roman`, $\ge 18\text{ pt}$), a four-color palette, native editable PowerPoint math equations (`<a14:m>`), authentic figures cropped from research papers, and zero-overlap dual-column dashed-box (`sysDot`) layouts verified via headless LibreOffice PDF/PNG rendering.

## When to Use

- Creating or modifying Master's/PhD opening reports (开题汇报), mid-term checks (中期汇报), or thesis defenses (学位答辩) using a `.pptx` template.
- Converting markdown research reports or literature surveys into dense, visually structured academic slides with boxes, arrows, formulas, and paper figures.
- Fixing typography, small fonts ($<18\text{ pt}$), uneditable screenshot formulas, or overlapping text/tables in `python-pptx` generated slides.

---

## 1. Non-Negotiable Typography & Color Contract

### 1.1 Dual-Font Rule (中文微软雅黑 + 英文/数字 Times New Roman)
Setting `run.font.name` in `python-pptx` only sets the Latin font (`a:latin`) and leaves East Asian characters (`a:ea`) in the theme default. Always set all four OpenXML font slots explicitly on every run:

| OpenXML Tag | Script Role | Required Typeface | Weight |
| :--- | :--- | :--- | :--- |
| `a:ea` | Chinese / East Asian characters | `微软雅黑` | Bold (`b="1"`) |
| `a:latin` | English letters & standard numbers | `Times New Roman` | Bold (`b="1"`) |
| `a:cs` | Complex script / numbers | `Times New Roman` | Bold (`b="1"`) |
| `a:sym` | Mathematical & special symbols | `Times New Roman` | Bold (`b="1"`) |

### 1.2 Minimum Font Size Rule ($\ge 18\text{ pt}$)
- **All text across every slide** (body bullets, flowchart boxes, equation symbol definitions, figure captions, table cells, and bibliography entries) **MUST be $\ge 18\text{ pt}$**. Only text embedded inside raster images is exempt.
- Recommended hierarchy:
  - Top navigation banner (`shapes[0]`): `26–28 pt` (White, Bold)
  - Slide subtitle (`shapes[1]`): `20–22 pt` (Primary Blue, Bold)
  - Column / section headings: `19–20 pt` (Primary Blue, Bold)
  - Body text, flow boxes, table cells, symbol definitions, references: `18 pt` (Bold)
- Always run `enforce_min_18pt_all_slides(prs)` before saving to guarantee zero violations.

### 1.3 Strict Four-Color Palette (主色调蓝色)
Never introduce arbitrary colors (no green, orange, purple, or grey body text). Use only:
1. **Primary Blue (`#00337F`, `RGBColor(0x00, 0x33, 0x7F)`)**: Top banner, slide subtitles, section headings, primary highlight boxes, arrows, and table headers.
   - *Light Blue tint (`#EAF1FB`, `RGBColor(0xEA, 0xF1, 0xFB)`)*: Secondary step box fills and alternating table rows.
2. **Pure Black (`#000000`, `RGBColor(0x00, 0x00, 0x00)`)**: Main body explanations, equation variable definitions, figure captions, and table body text.
3. **Emphasis Red (`#C00000`, `RGBColor(0xC0, 0x00, 0x00)`)**: Core physical constraints, failure thresholds, quantitative highlights, and final stage (`C`) borders.
4. **Pure White (`#FFFFFF`, `RGBColor(0xFF, 0xFF, 0xFF)`)**: Slide background and inverted text inside `#00337F` boxes/headers.

---

## 2. Slide Geometry, Vertical Offset & Spacing Rules

Standard 16:9 widescreen dimensions: **`13.333 in × 7.50 in`**.

1. **No Bottom "承上启下" Clutter**:
   - Do **NOT** place red transition sentences (`承上启下：...`) at the bottom of content slides. They compress vertical space and collide with tables or dashed boxes. Express page-to-page logical flow through the slide subtitle (`shapes[1]`) instead.
2. **Clear Gap Between Subtitle and Body (`0.35–0.40 in`)**:
   - Top banner (`shapes[0]`): `top = 0.11 in, height = 0.59 in`.
   - Slide subtitle (`shapes[1]`): Fix at `top = Inches(0.95), height = Inches(0.55)` with `vertical_anchor = MSO_ANCHOR.MIDDLE` (text occupies `y ≈ 1.05..1.38 in`).
   - Main body (dual-column dashed boxes or tables): Start at **`top = Inches(1.78)` to `Inches(1.90)`** (`height = Inches(5.22)`, spanning `1.78..7.00 in`). Never start dashed boxes at `1.50 in`, which touches the subtitle.
3. **Breathing Room Between Section Headings and Body Text**:
   - Inside dashed boxes (`top = 1.78 in`), start the first text block at `top = 1.90 in` (`0.12 in` top padding).
   - For every blue section heading (`color == C_BLUE` and `size_pt >= 19`), set `p.space_before = Pt(4)` (when `i > 0`) and `p.space_after = Pt(4)` so headings have a visible gap above their body paragraphs.
   - On single-column template slides, maintain at least `0.52–0.55 in` between a heading shape's `top` and its body paragraph shape's `top`.

---

## 3. Three Signature TJU Reference Layouts

### Layout 1: Mechanism + Native Formulas + Architecture Box + Paper Figures (对标参考图1)
- **Left Dashed Box (`left=0.45, top=1.78, width=6.65, height=5.22`)**:
  1. Blue heading (`19pt`): Mathematical / constitutive / control model title (`y=1.90`).
  2. Two native OMML equations (`y=2.34` and `y=2.80`, `height=0.42`).
  3. Compact 18pt symbol definitions + blue subheading (`y=3.28..4.45`).
  4. 3-box horizontal flow (`y=4.60..5.45`): `[White/Blue Input Box] → [Dark Blue Core Model Box] → [White/Blue Output Box]`, with an upward arrow `↑` from a full-width `[Red-Bordered Constraint/Safety Box]` (`y=5.78, height=0.62`) + 1-line red takeaway (`y=6.46`).
- **Right Dashed Box (`left=7.25, top=1.78, width=5.60, height=5.22`)**:
  1. Blue heading (`19pt`, `y=1.90`).
  2. One or two real figures cropped from cited papers (`y=2.40, height=2.30`).
  3. Figure citations (`图 X`) + blue/red engineering advantage summary (`y=4.80..6.85`).

### Layout 2: Layered Perception Pipeline + 3-Step Flow + ICP/Fitting Equations (对标参考图2)
- **Left Dashed Box (`left=0.45, top=1.78, width=5.65, height=5.22`)**:
  1. Blue heading + 2-line background challenge (`y=1.90..2.90`).
  2. Cropped pipeline figure from literature (`y=2.96, width=5.15`).
  3. Figure caption + blue subheading on baseline error + red conclusion (`y=4.46..6.80`).
- **Right Dashed Box (`left=6.25, top=1.78, width=6.60, height=5.22`)**:
  1. Blue heading (`y=1.90`).
  2. 3 horizontal step boxes connected by right arrows (`y=2.40, height=0.82`).
  3. 1-line measurement summary + 1-line blue equation intro (`y=3.32..4.08`).
  4. Two native OMML equations (`y=4.18` and `y=4.68`, `height=0.42`).
  5. 18pt symbol definitions + red grounding summary (`y=5.20..6.85`).

### Layout 3: Progressive `A → B → C` Ladder + Loss/Reward Modeling (对标参考图3)
- **Left Dashed Box (`left=0.45, top=1.78, width=6.45, height=5.22`)**:
  1. Blue heading (`y=1.90`).
  2. **Box A**: Dark blue header bar (`y=2.38, h=0.42`, white 18pt text `"A  基础..."`) + white body box (`y=2.80, h=0.46`, black 18pt text).
  3. Down arrow `↓` (`y=3.30`) → **Box B** (`y=3.58, h=0.86`, indented `left=0.82`, blue border: `"B  在 A 基础上引入..."` + 1-line detail).
  4. Down arrow `↓` (`y=4.48`) → **Box C** (`y=4.76, h=0.86`, indented `left=1.02`, red border: `"C  在 B 基础上开展..."` + 1-line detail).
  5. Bottom 2-line evaluation/deployment summary (`y=5.74`).
- **Right Dashed Box (`left=7.05, top=1.78, width=5.80, height=5.22`)**:
  1. Blue heading (`y=1.90`) + 2 native OMML equations (`y=2.34` and `y=2.80`).
  2. 18pt variable definitions (`y=3.26`) + paper figure (`y=4.06`) + 1-line figure citation (`y=6.44`).

---

## 4. Native PowerPoint Equations (`<a14:m>`) & StarMath Pitfalls

Always insert native editable PowerPoint equations using `add_omml_equation()` (`<a14:m><m:oMathPara><m:oMath>...</m:oMath></m:oMathPara></a14:m>`).

### Critical Pitfalls When Rendering via LibreOffice Impress (StarMath)
LibreOffice converts OMML `<m:oMath>` through its internal StarMath parser when exporting PDF/PNG. Violating any of the 4 rules below produces a giant red upside-down question mark (`¿`) or `XMLSyntaxError`:

| Pitfall | Cause in OMML / StarMath | Mandatory Fix |
| :--- | :--- | :--- |
| **`<` or `>` inside `m_r()`** | Raw `<` or `>` breaks `lxml.etree.fromstring` | Always run `html.escape(text)` inside `m_r(text)` |
| **Unmatched `(` or `[` across `m_sup()`** | Writing `m_r(" = (") + ... + m_sup(m_r(")"), m_r("T"))` puts `)` inside `{ ")" }^{T}`, leaving `(` unmatched (`¿`) | Keep `(` and `)` at the **exact same nesting level** (e.g., `m_r(" = (0, 0, 1, 1, 1, 1)")`) |
| **Raw `[` `]` or `\|` inside `<m:t>`** | StarMath parses raw `\|` as boolean `or` and `[` `]` as special grouping syntax | Use parentheses `( ... )` for vectors, `‖` (`\u2016`) for norms, and `e_θ` instead of `\|Δθ\|` |
| **Subscripts `equiv` or `succ`** | StarMath parses `equiv` (`≡`) and `succ` (`≻`) as binary operators missing operands (`¿`) | Use abbreviated subscripts `L_eqv` (instead of `L_equiv`) and `R_goal` (instead of `R_succ`) |

---

## 5. Character Width Budget & Visual Verification Loop

### 5.1 Single-Line Character Budget at 18pt
At `18pt`, each Chinese character is `~0.25 in` wide (and `~0.265 in` when Linux LibreOffice substitutes `微软雅黑` with `Noto Sans CJK`).
- In a `5.50 in` column, **at most 20–21 Chinese characters fit on a single line**.
- In a `6.35 in` column, **at most 23–24 Chinese characters fit on a single line**.
- In a `1.65–1.85 in` flowchart box, **at most 6–7 Chinese characters fit on a single line**.
- **Rule**: Keep single-line captions, variable definitions, and table cells within these character budgets so they never wrap onto an unexpected extra line and collide with equations or boxes below.

### 5.2 Mandatory Visual Verification Loop
Never deliver a `.pptx` without rendering and visually inspecting the slide PNGs:

```bash
python3 generate_ppt.py
libreoffice --headless --convert-to pdf output.pptx --outdir ./scratch/
pdftoppm -png -r 110 ./scratch/output.pdf ./scratch/slide_check
```

Inspect every rendered `slide_check-*.png` to verify:
1. Zero text-on-text or text-on-box overlap;
2. Zero `¿` equation syntax errors;
3. Clean `0.35–0.40 in` white space between the top blue subtitle and the dashed boxes/tables;
4. Visible breathing room between blue section headings and black body text.

---

## Supporting Files in This Skill

- `scripts/ppt_builder_helpers.py`: Reusable helper library providing `set_run_font`, `update_paragraph_text`, `update_shape_paragraphs`, `update_table_cells`, `clear_slide_body`, `add_dashed_box`, `add_solid_box`, `add_arrow`, `add_text_block`, `add_omml_equation`, `m_r`, `m_sub`, `m_sup`, `m_subsup`, `m_frac`, `m_sqrt`.
- `examples/generate_tju_opening_ppt.py`: Complete reference script generating a 26-slide Tianjin University opening report PPTX.
- `references/reference_layouts/`: The 3 canonical Tianjin University slide layout reference images (`ref_layout_1_mechanism_formula.jpg`, `ref_layout_2_perception_icp.jpg`, `ref_layout_3_abc_ladder.jpg`) and base template `tju_template.pptx`.
