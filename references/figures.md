# Figures

Four self-contained routes, no external skill needed.  Required packages: `python-pptx Pillow numpy lxml`;
route C adds `cairosvg` (and `fontTools` for glyph checks), route D adds `matplotlib` (≥ 3.6 for mixed
CJK/Latin labels).

## 1. Route selection — Native Editable First

**Default rule:** If a diagram can be expressed with boxes, rounded pills, grouping frames, dots, labels, and
straight/orthogonal arrows, **always build it natively in PowerPoint with `tju_native_diagrams`** (all exported
from `tjuppt`).  Never paste a static PNG/SVG image for an ordinary flowchart, system architecture, closed loop,
challenge-to-content map, timeline, or bottom summary strip — advisors expect to click and edit diagram text
directly in PowerPoint.

| Need | Route | Tool | Editable in PowerPoint |
|---|---|---|---|
| Titled native diagram panel (replaces `add_figure_panel` for block diagrams) | **A0 (default)** | `add_native_panel(slide, region, "标题", framed=True)` | **yes** |
| Linear / vertical flow (≤ 6 steps) | **A1 (default)** | `add_flow_pipeline(slide, *r.box, steps, direction="h"/"v")` | **yes** |
| Closed loop / feedback iteration | **A2 (default)** | `add_feedback_loop(slide, *r.box, steps, "反馈说明")` | **yes** |
| Layered system architecture | **A3 (default)** | `add_layered_architecture(slide, *r.box, [(层, [模块…]), …])` | **yes** |
| Grouped multi-module architecture / arbitrary topology | **A4 (default)** | `add_diagram_group` + `add_role_box` + `add_pill` + `add_junction_dot` + `add_diagram_label` + `connect_boxes` / `add_polyline_arrow` | **yes** |
| Method-evolution timeline | **A5 (default)** | `add_timeline(slide, region, [(year, title, desc, role), …])` or diagonal `add_connector` + `add_role_box` | **yes** |
| Real photos, GIFs, cited paper figures | — | `add_figure_panel(..., fit="crop")` / `add_picture_exact` | media |
| Data curves, bars, ablations, heatmaps | D | `tju_plot_style.py` (Matplotlib) | data plot |
| Complex non-box geometry, curved schematics, dense custom vector art that native shapes cannot express | C *(fallback only)* | `TjuDiagram` (`tju_svg_diagram.py`) → SVG + PNG | no (keep script + `.svg`) |
| Sequence / state diagram, only if `mmdc` exists and A cannot express it | M *(fallback only)* | `mmdc -c references/tju_mermaid_theme.json` | no |

**Native diagram typography & framing contract:**
- Shapes built by `tju_native_diagrams` are tagged with `Diagram:` in their shape name and strip default theme
  drop shadows automatically (`no_theme_effects`).
- Keep text inside native diagram boxes at **16–18 pt** wherever possible (`≥ 16 pt` for full-slide system
  architectures); compact milestone/timeline cards may go down to **`MIN_DIAGRAM_PT = 14 pt`**.  Never shrink
  regular slide body text (`MIN_PT = 18 pt`).
- Use `add_diagram_group` (unfilled dashed rounded frame) for visual grouping — never place a solid filled
  background rectangle behind boxes (which triggers overlap audit errors).
- Keep palette restrained (≤ 3 colors: deep blue `#00337F`, ice-blue tint `#EAF1FB`, restrained red `#C00000`
  only for focal triggers/constraints — never shout a secondary module in bright red).

## 2. Draw at slot size (the alignment rule)

```python
left, right = page.body.split(2)
w, h = figure_slot(left)                       # image area under the panel header
d = TjuDiagram(w, h, base_pt=18)               # canvas in inches = slot size
...
add_figure_panel(page.slide, left, "标题", d.save("fig/a.png", svg_path="fig/a.svg"))
```
- Self-drawn figures must have the slot's aspect ratio (±5 %); `add_figure_panel` raises otherwise.
  Photos/GIFs use `fit="crop"` (center crop via picture crop, the file is untouched).
- Never draw large and shrink, never letterbox under a header bar.
- Text sizes inside figures: node/box labels 18 pt (`base_pt=18`); axis ticks and legends ≥ 16 pt
  (`apply_tju_style()` defaults).  `save_tju_figure` warns below 14 pt.

## 3. Palette and fonts inside figures
- Node roles: `input` (light-blue fill, blue border) · `core` (blue fill, white text) · `output` (white,
  red border) · `emphasis` (red fill) · `plain` · `note`.  Lines: blue / red / black only.
- Plots: three line colors; the 4th+ series differ by line style + marker (`tju_line_styles()`); bars use
  `TJU_BAR_STYLES` hatching.  `tju_audit.py --images fig.png` checks the palette (anti-aliasing tolerated).
- Fonts: Latin Times New Roman, CJK 微软雅黑 (Linux falls back to Noto Sans CJK SC).  The SVG engine assigns
  fonts per run and routes any symbol missing from Times (e.g. `‖`) to the CJK font; `save()` fails if
  neither font has a glyph.  `python scripts/tju_svg_diagram.py` prints the resolved fonts.
- Sub/superscripts: `x_{k}^{ref}` works in node text.  Combining accents (`x̂`) do not render in cairosvg —
  put hatted symbols in a native OMML equation on the slide.

## 4. TjuDiagram essentials
- `node(id, x, y, w, h, text, role, dashed=False, check=True)`; `group(x, y, w, h, label)`;
  `edge(src, dst, label, src_side, dst_side, route="auto|straight|hv|vh|hvh|vhv", color, dashed)`;
  `arrow(points)` free polyline; `text(x, y, content, anchor)`; `row/column` helpers; `layered(...)` auto
  layout; `from_spec(json)`.
- `save(png, svg_path=...)` validates first (strict): text overflow, node overlap, edge through a node,
  edge or arrow crossing a group label, label collisions, out-of-canvas elements, missing glyphs,
  combining accents.  Fix the layout; never pass `strict=False` for a final deck.
- Group labels sit inside the group's top-left corner; edges entering a group from the top must avoid that
  strip (route via the side or leave ≈ 0.35 in above the first node row).
- Node text wraps by real glyph widths and balances line lengths to avoid single-character last lines.

## 5. Visual checks after rendering (`tju_render.py`)
- Arrows follow the data/control flow, touch box edges, never cross text.
- Legends do not hide curves (`ax.set_ylim` headroom or `ncol`).
- Figure text looks the same size as slide body text.
- At most 1–2 red elements per column.
