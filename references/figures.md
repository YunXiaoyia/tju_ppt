# Figures

Four self-contained routes, no external skill needed.  Required packages: `python-pptx Pillow numpy lxml`;
route C adds `cairosvg` (and `fontTools` for glyph checks), route D adds `matplotlib` (≥ 3.6 for mixed
CJK/Latin labels).

## 1. Route selection

| Need | Route | Tool | Editable in PowerPoint |
|---|---|---|---|
| ≤ 6-step linear flow | A1 | `add_flow_pipeline(slide, *region.box, steps)` | yes |
| Closed loop / iterative training | A2 | `add_feedback_loop(slide, *region.box, steps, "反馈说明")` | yes |
| Layered architecture | A3 | `add_layered_architecture(slide, *region.box, [(层, [模块…]), …])`, `!` prefix = new module | yes |
| Mapping (挑战 → 研究内容), arbitrary links | A4 | `add_role_box` + `connect_boxes` | yes |
| Multi-branch roadmap, grouped modules, algorithm block diagram, timeline | C | `TjuDiagram` (`tju_svg_diagram.py`) → SVG + PNG | no (keep script + `.svg`) |
| Curves, bars, ablations, heatmaps | D | `tju_plot_style.py` | no (keep script + data) |
| Sequence / state diagram, only if `mmdc` exists | M | `mmdc -c references/tju_mermaid_theme.json` | no |
| Real photos, GIFs, cited paper figures | — | `add_figure_panel(..., fit="crop")` / `add_picture_exact` | — |

Prefer A when it can express the figure (the advisor can edit text directly).  Switch to C when lines must
detour, groups are needed, or ≥ 3 edges enter one node.

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
