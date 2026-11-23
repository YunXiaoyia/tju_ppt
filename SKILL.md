---
name: tju-ppt
description: Use when creating, editing, auditing, or revising Tianjin University (TJU / 天津大学) or academic group opening report (开题汇报), mid-term check (中期汇报), or thesis defense (学位答辩) PowerPoint (.pptx) decks, drawing defense-ready flowcharts, system architectures, or plots, or fixing decks criticized as too text-heavy, jargon-heavy, colloquial, badly ordered, or hard for cross-field committee members to follow.
---

# TJU Academic Presentation (天津大学课题组学术汇报 PPT) Skill

Build defense-grade academic PowerPoint (`.pptx`) decks on the Tianjin University 16:9 template (`13.333 in × 7.50 in`) using `python-pptx` and `scripts/ppt_builder_helpers.py`.

## When to Use & Non-Triggers
- **Use when**: Building or polishing Master's/PhD opening defenses (开题汇报), mid-term checks (中期汇报), or thesis defenses (学位答辩) requiring TJU group visual conventions (`>=18 pt`, 4-color palette, native OMML math, dual-column dashed containers).
- **Also use when**: Reviewing or revising a deck after advisor/committee feedback such as "听不懂 / 专业词太多 / 口水话太多 / 字太多 / 每页太满 / 背景没讲清 / 逻辑顺序不对", or auditing whether a deck is ready to present.
- **Do NOT use when**: Creating general commercial pitch decks, Beamer LaTeX slides (`paper-slides`), conference posters (`paper-poster`), or SVG-first slides (`ppt-master`).

---

## 0. Story & Audience Contract (decide content BEFORE layout)

Typography compliance does not make a deck understandable. Full rules + checklist: `references/narrative_principles.md`.

1. **Audience = smart non-expert.** Assume half the committee works in an adjacent field and has never seen your scene. Slides carry the argument; the written report carries the details — never paste report paragraphs, symbol sets, or full ablation lists.
   - **Accessible ≠ childish.** They are professors outside your sub-field, not laypeople: lower the entry barrier by *defining* correct terms once, never by replacing them with slang, folk metaphors, or marketing buzzwords (e.g., 草稿台 / 算总账 / 痛点 / 搞定 / 最快). Every printed sentence should be acceptable in the written report. Use neutral labels: 问题 / 方法 / 局限 / 借鉴 / 判据 / 预期收益.
2. **Fixed arc:** Background/scene → Specificity & difficulty (why existing rules/methods break, why the new mechanism works) → Abstracted research question → Related work & limits → My work (overview first) → Validation plan → Innovation / feasibility / dated plan. Mid-term and defense variants are in `references/narrative_principles.md`.
3. **Subtitle-only test:** the slide subtitles read in order must tell the story as concise declarative claims, not generic topic labels or quiz questions.
4. **Define before use:** every domain term, acronym, or internal ID gets a picture or one-line definition before it is used in an argument; ≤ 3 new terms per slide.
5. **One slide, one message, speakable in its slot:** ≈ 150–200 visible characters for a 40–60 s slide (tables included); if the per-slide speech script overruns its time budget, cut or split.
6. **Picture over prose:** each content slide carries a visual (real photo/GIF, cited paper figure, native diagram, self-drawn schematic, or plot); ≤ 2–3 core formulas in the whole talk, each followed by a one-line explanation.
7. **Related work = 1–2 works per direction** with original figure + idea / limit / borrow; each limit maps to one of your contributions and vice versa.
8. **Claim hygiene:** every number is traceable with its test condition and valid inference; avoid unproven absolutes (彻底 / 杜绝 / 100 % / 零误差); expected outcomes are measurable (metric + baseline + direction).
9. **Completeness (opening):** research question, gaps, method, **validation plan (baselines, scenarios, metrics, statistics)**, innovations, feasibility + risks, concrete-date timeline, measurable outcomes.
10. **Consistency:** title, terms, numbers, and page numbers match across slides, speech script, and written report; renumber everything after inserting/deleting slides.

---

## 1. Non-Negotiable Typography, Inline Sub/Sup & Color Contract

### 1.1 Dual-Font 4-Slot Rule & Minimum Size ($\ge 18\text{ pt}$)
`run.font.name` only sets `a:latin`. Always set all four OpenXML slots (`a:ea="微软雅黑"`, `a:latin="Times New Roman"`, `a:cs="Times New Roman"`, `a:sym="Times New Roman"`, Bold `b="1"`) via `set_run_font()` and run `enforce_min_18pt_all_slides(prs)` before saving:
- Top banner (`shapes[0]`): `26–28 pt` White (`#FFFFFF`) (`setup_tju_content_slide(..., banner_pt=28)`)
- Slide subtitle: `20–22 pt` Primary Blue (`#00337F`) (`setup_tju_content_slide(..., subtitle_pt=20)`)
- Column / card headings: `19–20 pt` Primary Blue (`#00337F`)
- All body text, flow boxes, tables, OMML equation runs, slide numbers, and references: **`18 pt` minimum** (only raster image internals are exempt).

### 1.2 Zero Raw Code Math in Prose & Tables
- **Never** leave raw ASCII/LaTeX math strings (`x_k^ref`, `d_{k+1}^obs`, `Q_k`, `L_reg`, `x^2`) in body text, headings, or table cells.
- Pass `_{sub}` and `^{sup}` markup (e.g., `"x_{k}^{ref}"`, `"d_{k+1}^{obs}"`, `"Q_{k}"`) to `append_rich_text`, `add_text_block`, `add_solid_box`, or `add_clean_tju_table`. It writes OpenXML `baseline="-25000"` / `"30000"` while keeping `sz="1800"` (`>=18 pt` compliant) and leaves multi-letter code identifiers (`CODE_ID`) intact.

### 1.3 Strict Four-Color Palette
Allow **only** these five hex values across all shapes, borders, fills, tables, and runs (never leave unstyled `add_shape()` theme `schemeClr` fills/lines):
1. **Primary Blue (`#00337F`)** + **Light Blue tint (`#EAF1FB`)** for secondary cards, highlighted equation boxes, and alternating table rows.
2. **Pure Black (`#000000`)** for body explanations and variable definitions.
3. **Emphasis Red (`#C00000`)** for physical constraints, failure boundaries, and core takeaways (use sparingly — at most 1–2 focal elements per column).
4. **Pure White (`#FFFFFF`)** for slide canvas, alternating cards, and inverted text on `#00337F` / `#C00000` headers.

---

## 2. Template Map, Visual Hierarchy & Reusable Primitives

### 2.1 Template & Reference Layout Boundaries
- **`references/tju_template.pptx`** (26-slide source deck): contains two Slide Masters; `Master 1` holds the actual TJU layouts:
  - Cover / Thanks (`标题幻灯片`): 0-based slide indices `0`, `25` → style with `format_cover_slide()` / `format_thanks_slide()` (white bold text on dark-blue background).
  - TOC / Section transitions (`1_仅标题`): 0-based indices `1`, `10`, `18`, `21` → style with `format_toc_slide(slide, toc_items, active_idx, page_num)`.
  - Content slides (`1_自定义版式`): 0-based indices `2–9`, `11–17`, `19–20`, `22–24` → rebuild every selected slide via `setup_tju_content_slide()` so legacy template colors/fonts are completely cleared.
  - Prune and reorder safely with `rearrange_slides(prs, src_indices_0based)` (automatically drops unused `rId` relations).
- **`references/reference_layouts/ref_layout_1..3.jpg`**: from an older light-banner template (white top banner, bottom red summary strip, no page numbers). Use them **only as reference for interior dual-column dashed-box hierarchy** (3-box flows, joined cards, equation+annotation pairing); follow §1–§2 for the dark-blue banner, bottom-right slide number, and omission of bottom transition strips.

### 2.2 Standard Skeleton & Structured Components (`scripts/ppt_builder_helpers.py`)
1. **Standard Content Header & Slide Number (`setup_tju_content_slide`)**:
   - Anchors every subtitle (`top=0.94 in`) with a left `#00337F` vertical bar (`0.08 × 0.36 in`) and horizontal divider (`y=1.56 in`), leaving `0.22 in` clean space above `y=1.78 in` dashed boxes (`height=5.22 in`, ending at `7.00 in`), plus bottom-right `#00337F` `18 pt` slide number (`top=7.05 in`).
   - Automatically strips residual `mc:AlternateContent` ghost shapes (created when PowerPoint re-saves `<a14:m>` equations) and creates a banner textbox safely if invoked on an empty layout.
   - **Never** add bottom `"承上启下：..."` text; express page-to-page logic in the subtitle.
   - **Wrapper Identity Pitfall**: when clearing shapes manually in `python-pptx`, never compare `sh is banner` (wrappers regenerate); always compare `sh._element is banner_el`.
2. **Break Text Monotony with Structured Components**:
   - Never fill both columns with 6+ plain `•` paragraphs or misaligned `｜` text pipes.
   - **3-Box Horizontal Flow**: `add_flow_pipeline(...)` (`[Light-Blue Input] → [Dark-Blue Core] → [Red-Bordered Output]`).
   - **Two-Part Joined Cards**: `add_joined_card(...)` — dark `#00337F` / `#C00000` header bar seamlessly touching (`oy=0`) a `#EAF1FB` / `#FFFFFF` body box.
   - **4-Color Tables & Badge Tables**: `add_clean_tju_table(..., badge_col0=True)` — blue header, alternating `#EAF1FB`/`#FFFFFF` rows, explicit `#00337F` cell borders (no default theme `schemeClr`), optional `#00337F` white-text ID column (`B1..B4` / `M1..M4`).
   - **Colored Lead-Ins**: multi-segment runs `[("• 关键机制：", C_BLUE, 18), ("具体说明...", C_BLACK, 18)]` and `[1]` tags in `#00337F` on reference slides.
3. **Single-Shape Framing (Zero Declared-BBox Overlap)**:
   - Frame figures directly via `pic.line` in `add_image_fit(..., border_color=C_BLUE)` and highlight equations directly via `fill_color=C_BLUE_LIGHT, line_color=C_BLUE` on `add_omml_equation()` — never stack a background rectangle behind a picture or textbox.
4. **Grounded Academic Defense Tone**:
   - Individual thesis defenses present system module interfaces objectively (never multi-person team task allocations).
   - Avoid meta-disclaimers or bureaucratic imperatives; state assumptions, constraints, and experimental controls directly.

---

## 3. Native OMML Equations (`<a14:m>`) & Rendering Rules

1. **Accent Hat Rule (`m_acc`)**: Always use Unicode combining circumflex `U+0302` (`ch="\u0302"`), **never** ASCII `"^"` (`U+005E`), which sinks onto capital letters (`S`, `T`, `D`, `O`) in PowerPoint.
2. **Upright Operators**: Render `log`, `exp`, `min`, `max`, `∑`, parentheses, and descriptive subscripts (`ref`, `obs`, `avg`) with `italic=False`.
3. **LibreOffice StarMath Compatibility**: Escape `<`/`>` via `html.escape`, keep `(` and `)` at the same nesting level, use `‖` (`\u2016`) instead of raw `|` or `[` `]` inside `<m:t>`, and avoid reserved StarMath operator names (`equiv`, `succ`) as bare subscripts.

---

## 4. Character Width Budget & Mandatory Visual Verification

### 4.1 18 pt Single-Line Budget & Orphan Punctuation Check
- `5.7–5.8 in` column: **21–22 CJK chars/line**; `2.9 in` grid card: **10–11 CJK chars/line**; `1.75 in` flow box: **6–7 CJK chars/line**.
- Inspect rendered PNGs to ensure zero orphan punctuation (`，`, `；`, `。`) wrapping alone onto a new line and zero 2-line overflow inside single-line `0.40–0.58 in` bars.

### 4.2 End-to-End Audit & Render Loop
1. **Enforce & Save via Staging File**: Run `enforce_min_18pt_all_slides(prs)` then `save_staged_pptx(prs, target_path)` so open PowerPoint windows on Windows never block builds.
2. **OpenXML & Palette Audit**: Run `python <skill_dir>/scripts/tju_audit.py deck.pptx --skip 1,2 --images <self-drawn PNGs>` (must report `0 issue(s)`). It checks:
   - 0 notes, 0 `"承上启下"`, 0 unparsed `_{`/`^{` or raw ASCII math (`x_k`, `x^2`);
   - all `<a:r>` and OMML `<m:r>` runs (outside `mc:Fallback`): `sz >= 1800`, `b="1"`, 4 font slots;
   - strict palette: explicit `srgbClr` in `ALLOWED_HEX`, no explicit `schemeClr`, no unstyled shapes falling back to theme `fillRef`/`lnRef`;
   - 0 declared BBox overlaps (`ox > 0.03` and `oy > 0.03`; connectors and unfilled dashed containers exempt);
   - ≤ 0.5 % off-palette pixels in self-drawn figures.
   - **Content warnings** (`--content`, add cover/TOC/references/thanks pages to `--skip`): text density > 200 visible units, content slide without visual (accepts pictures, charts, and native diagrams/arrows), estimated multi-line text overflow, unproven absolute claims, colloquial wording, and first appearance of non-trivial acronyms on content slides. Warnings are judgment calls: fix or justify; `--strict-content` turns them into build failures.
3. **Render & Visually Inspect Every Modified Slide PNG**:
   - **Windows (PowerPoint COM + PyMuPDF — exact native fidelity)**: Export staged `.pptx` via `win32com.client.DispatchEx("PowerPoint.Application")` (`WithWindow=False`, `SaveAs(..., 32)` to PDF, followed by `deck.Close(); app.Quit(); gc.collect()`), render `144 dpi` PNGs via `pymupdf`, and inspect every modified slide PNG.
   - **Linux**: `libreoffice --headless --convert-to pdf staged.pptx` + `pdftoppm -png -r 144`.

---

## 5. Research Diagrams & Data Figures (Self-Contained, No External Skills)

Full routing table, limits, and QA checklist: `references/diagram_routing.md`. Working demos: `examples/generate_diagram_showcase.py` and `examples/generate_tju_opening_ppt.py`.

### 5.1 Route Selection (prefer editable native shapes)
| Need | Use |
|---|---|
| ≤ 6-step linear flow | `add_flow_pipeline(slide, l, t, w, h, steps, direction="h"/"v")` — roles default input → core → output |
| Closed loop / iterative training | `add_feedback_loop(slide, l, t, w, h, steps, "反馈说明")` — red dashed return path + label |
| Layered system architecture | `add_layered_architecture(slide, l, t, w, h, [(层名, [模块...]), ...], flow="up")` — prefix a module with `!` for red-bordered new module |
| Arbitrary box links | `connect_boxes(slide, g1, g2)` / `add_connector(...)` — straight or orthogonal, triangle arrowheads |
| Multi-branch roadmap, cross-group links, algorithm block diagram | `TjuDiagram` (`scripts/tju_svg_diagram.py`) → validated SVG + 2× PNG → `add_image_fit(..., border_color=None)` |
| Curves / bars / ablations / heatmaps | `scripts/tju_plot_style.py`: `apply_tju_style()`, `new_figure(w, h)`, `tju_line_styles()`, `TJU_BAR_STYLES`, `TJU_CMAP_SEQ/DIV`, `save_tju_figure()` |
| Sequence / state diagrams (only if `mmdc` exists) | `mmdc -i x.mmd -o x.png -s 2 -c references/tju_mermaid_theme.json`; otherwise fall back to native shapes or `TjuDiagram` |

Native-shape modules live in `scripts/tju_native_diagrams.py`; they run `check_box_fit()` (18 pt width budget) and warn before overflow reaches the render.

### 5.2 Figure Rules
1. **Draw at final placement size**: pass the PPT placement box size (inches) to `TjuDiagram(w, h)` / `new_figure(w, h)` so 18 pt in the figure stays 18 pt on the slide. Never draw large and shrink.
2. **Palette inside figures**: node roles `input`/`core`/`output`/`emphasis`/`plain`/`note`; data plots use only blue/red/black lines and distinguish the 4th+ series by line style + marker; bars use hatch patterns.
3. **Fonts**: the SVG engine assigns Times New Roman to Latin runs and 微软雅黑 (or a local CJK sans fallback) to CJK runs separately, because cairosvg uses only the first `font-family` and has no glyph fallback.
4. **No combining accents** (`x̂`, `ẋ`) in figure text — they render as tofu. Use `_{}`/`^{}` there and put hatted symbols in native OMML on the slide.
5. **Strict validation**: `TjuDiagram.save()` raises on text overflow, node overlap, edges crossing unrelated nodes, label collisions, or out-of-canvas elements (nodes, groups, edge waypoints, free arrows, texts). Fix the layout; do not pass `strict=False` for final decks.
6. **Keep sources**: store generator scripts and `.svg` files beside the PNGs (`figures/`) because routes C/D are not editable inside PowerPoint.

## Supporting Files
- `scripts/ppt_builder_helpers.py`: OpenXML typography, `_{sub}`/`^{sup}` rich text, OMML equations, framed images, cover/TOC/content/thanks skeletons, joined cards, 4-color tables, slide pruning/reordering, staging save, and `enforce_min_18pt_all_slides`.
- `scripts/tju_native_diagrams.py`: Editable native flow, feedback-loop, and layered-architecture diagrams, connectors, and 18 pt box-fit budget.
- `scripts/tju_svg_diagram.py`: `TjuDiagram` research-schematic engine (Python API, JSON spec, or `layered()` auto layout; validates, then renders SVG/PNG).
- `scripts/tju_plot_style.py`: Matplotlib style for projected defense slides using the TJU palette.
- `scripts/tju_audit.py`: Executable OpenXML audit (`a:r` + OMML `m:r`, strict palette & theme-fallback check, raw math, overlaps), content-warning audit (`--content`), and self-drawn image palette audit.
- `examples/generate_tju_opening_ppt.py`: Complete 10-slide compliant opening-defense generator example (outputs `examples/example_opening_presentation.pptx`).
- `examples/generate_diagram_showcase.py`: Four-slide demo of every diagram route plus automated audit.
- `references/tju_template.pptx`: Tianjin University 16:9 PowerPoint template (`Master 1` holds cover/TOC/content layouts).
- `references/reference_layouts/`: Three interior dual-column dashed-box visual hierarchy references (`ref_layout_1..3.jpg`).
- `references/diagram_routing.md`, `references/tju_mermaid_theme.json`: Diagram routing, known limits, QA checklist, and optional Mermaid theme.
- `references/narrative_principles.md`: Audience, story arc, academic register, define-before-use, density budget, related-work compression, claim hygiene, and opening-report completeness checklist (§0 in full).
