---
name: tju-ppt
description: Use when creating, editing, or optimizing Tianjin University (TJU / 天津大学) or academic group opening report (开题汇报), mid-term, or thesis defense PPTX presentations with python-pptx, especially when enforcing strict font rules (微软雅黑 + Times New Roman >= 18pt), four-color palette, native editable OMML equations, real paper figures, and zero-overlap dual-column dashed-box layouts.
---

# TJU Academic Presentation (天津大学课题组学术汇报 PPT) Skill

Build defense-grade academic PowerPoint (`.pptx`) decks on the Tianjin University 16:9 template (`13.333 in × 7.50 in`) using `python-pptx` and `scripts/ppt_builder_helpers.py`.

## When to Use & Non-Triggers
- **Use when**: Building or polishing Master's/PhD opening defenses (开题汇报), mid-term checks (中期汇报), or thesis defenses (学位答辩) requiring TJU group visual conventions (`>=18 pt`, 4-color palette, native OMML math, dual-column dashed containers).
- **Do NOT use when**: Creating general commercial pitch decks, Beamer LaTeX slides (`paper-slides`), conference posters (`paper-poster`), or SVG-first slides (`ppt-master`).

---

## 1. Non-Negotiable Typography, Inline Sub/Sup & Color Contract

### 1.1 Dual-Font 4-Slot Rule & Minimum Size ($\ge 18\text{ pt}$)
`run.font.name` only sets `a:latin`. Always set all four OpenXML slots (`a:ea="微软雅黑"`, `a:latin="Times New Roman"`, `a:cs="Times New Roman"`, `a:sym="Times New Roman"`, Bold `b="1"`) via `set_run_font()` and run `enforce_min_18pt_all_slides(prs)` before saving:
- Top banner (`shapes[0]`): `26–28 pt` White (`#FFFFFF`)
- Slide subtitle: `20–22 pt` Primary Blue (`#00337F`)
- Column / card headings: `19–20 pt` Primary Blue (`#00337F`)
- All body text, flow boxes, tables, equation runs, slide numbers, and references: **`18 pt` minimum** (only raster image internals are exempt).

### 1.2 Zero Raw Code Math in Prose & Tables
- **Never** leave raw ASCII/LaTeX math strings (`M_t^tmp`, `S_{t+1}^obs`, `F_table^truck`, `L_pair`, `ϕ_split`) in body text, headings, or table cells.
- Pass `_{sub}` and `^{sup}` markup (e.g., `"M_{t}^{tmp}"`, `"S_{t+1}^{obs}"`, `"L_{pair}"`) to `append_rich_text`, `add_text_block`, or `add_solid_box`. It sets OpenXML `baseline="-25000"` / `"30000"` while keeping `sz="1800"` (`>=18 pt` compliant) and leaves code identifiers (`ERR_DST_OCCUPIED`) intact.

### 1.3 Strict Four-Color Palette
Allow **only** these five hex values across all shapes, borders, fills, and runs:
1. **Primary Blue (`#00337F`)** + **Light Blue tint (`#EAF1FB`)** for secondary cards, highlighted equation boxes, and alternating table rows.
2. **Pure Black (`#000000`)** for body explanations and variable definitions.
3. **Emphasis Red (`#C00000`)** for physical constraints, failure boundaries, and core takeaways (use sparingly — at most 1–2 focal elements per column).
4. **Pure White (`#FFFFFF`)** for slide canvas, alternating cards, and inverted text on `#00337F` / `#C00000` headers.

---

## 2. Academic Visual Hierarchy & Anti-"Wall-of-Text" Aesthetics

1. **Standard Content Header & Slide Number (`setup_tju_content_slide`)**:
   - Anchor every content subtitle (`top=0.94 in`) with a left `#00337F` vertical bar (`0.08 × 0.36 in`) and a crisp `#00337F` horizontal divider (`y=1.56 in`), leaving `0.22 in` clean white space above `y=1.78 in` dashed boxes (`height=5.22 in`, ending at `7.00 in`).
   - Always render a bottom-right `#00337F` `18 pt` slide number (`top=7.05 in`, `Slide Number <n>`) on content slides so defense committee members can cite page numbers during Q&A.
   - **Never** add bottom `"承上启下：..."` text; express page-to-page logic in the subtitle.
   - **Wrapper Identity Pitfall**: When clearing shapes in `python-pptx`, never compare `sh is banner` (wrappers regenerate per iteration); always compare `sh._element is banner_el`.
2. **Break Text Monotony with Structured Visual Components**:
   - Never fill both columns with 6+ plain `•` paragraphs or misaligned `｜` text pipes.
   - Convert comparisons, state machines, recovery examples, scenario lists, and ablations into structured visual elements:
     - **3-Box Horizontal Flow**: `[Light-Blue Input] → [Dark-Blue Core] → [Red-Bordered Output/Constraint]` (`ref_layout_1` / `ref_layout_2`).
     - **Two-Part Joined Cards**: Dark `#00337F` / `#C00000` header bar (`h=0.46 in`) seamlessly touching (`oy=0`) a `#EAF1FB` / `#FFFFFF` body box (`ref_layout_3` / Innovation cards).
     - **2×4 Card Grids & Badge Tables**: Use `2×4` alternating `#EAF1FB`/`#FFFFFF` cards for 8 scenarios/metrics and compact badge tables (`col0` `#00337F` white ID `A1..A9`) for ablations.
     - **Colored Lead-Ins**: Use multi-segment runs `[("• 关键机制：", C_BLUE, 18), ("具体工程说明...", C_BLACK, 18)]` and highlight `[R1]` tags in `#00337F` on reference slides.
3. **Single-Shape Framing (Zero Declared-BBox Overlap)**:
   - Frame paper/platform figures directly via `pic.line` in `add_image_fit(..., border_color=C_BLUE)` and highlight key formulas directly via `fill_color=C_BLUE_LIGHT, line_color=C_BLUE` on `add_omml_equation()` — never stack a separate background rectangle behind a picture or textbox.
4. **Grounded Academic Defense Tone**:
   - Individual thesis defenses must present system module interfaces objectively (never list multi-person team task allocations).
   - Avoid AI-style meta-disclaimers or bureaucratic scolding (`"口径声明"`, `"严禁预编"`, `"强制冻结...零重做"`, `"独立证伪"`). State assumptions, constraints, and experimental controls directly.

---

## 3. Native OMML Equations (`<a14:m>`) & Rendering Rules

1. **Accent Hat Rule (`m_acc`)**: Always use Unicode combining circumflex `U+0302` (`ch="\u0302"`), **never** ASCII `"^"` (`U+005E`), which sinks onto capital letters (`S`, `T`, `D`, `O`) in PowerPoint.
2. **Upright Operators**: Render `log`, `exp`, `min`, `max`, `∑`, parentheses, and descriptive subscripts (`obs`, `tmp`, `pair`, `avg`) with `italic=False`.
3. **LibreOffice StarMath Compatibility**: Escape `<`/`>` via `html.escape`, keep `(` and `)` at the same nesting level, use `‖` (`\u2016`) instead of raw `|` or `[` `]` inside `<m:t>`, and avoid reserved StarMath operator names (`equiv`, `succ`) as bare subscripts.

---

## 4. Character Width Budget & Mandatory Visual Verification

### 4.1 18 pt Single-Line Budget & Orphan Punctuation Check
- `5.7–5.8 in` column: **21–22 CJK chars/line**; `2.9 in` grid card: **10–11 CJK chars/line**; `1.75 in` flow box: **6–7 CJK chars/line**.
- Inspect rendered PNGs to ensure zero orphan punctuation (`，`, `；`, `。`) wrapping alone onto the start of a new line and zero 2-line overflow inside single-line `0.40–0.58 in` bars.

### 4.2 End-to-End Audit & Render Loop
1. **Save via Staging File**: Save to a scratch `staged.pptx` first, then `shutil.copy2` to the target `.pptx` inside `try...except PermissionError` so open user PowerPoint windows never block builds.
2. **OpenXML Audit**: Verify 0 notes, 0 `"承上启下"`, 0 unparsed `_{`/`^{`, all runs `sz >= 1800` + `b="1"` + 4 font slots + `ALLOWED_HEX`, and 0 declared BBox overlaps (`ox > 0.03` and `oy > 0.03`).
3. **Render & Visually Inspect Every Modified Slide PNG**:
   - **Windows (PowerPoint COM + PyMuPDF — exact native fidelity)**: Export `staged.pptx` with `win32com.client.DispatchEx("PowerPoint.Application")` (`WithWindow=False`, `SaveAs(..., 32)` to PDF, followed by `deck.Close(); app.Quit(); gc.collect()`), render `144 dpi` PNGs via `pymupdf`, and inspect every modified slide PNG.
   - **Linux**: `libreoffice --headless --convert-to pdf staged.pptx` + `pdftoppm -png -r 144`.

## Supporting Files
- `scripts/ppt_builder_helpers.py`: Reusable OpenXML typography, `_{sub}`/`^{sup}` rich text, OMML equations, framed images, header/footer builder, and `enforce_min_18pt_all_slides`.
- `examples/generate_tju_opening_ppt.py`: Complete multi-slide opening defense generator example.
- `references/reference_layouts/`: Canonical TJU reference layouts (`ref_layout_1..3.jpg`) and `tju_template.pptx`.
