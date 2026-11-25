# Page Recipes

Default page shapes for TJU opening reports (开题), reusable for 中期 / 答辩.  They are defaults: if the
advisor prescribes a different structure for a page, follow the advisor and keep the grid and contract.
Working code for every recipe: `examples/build_example_deck.py`.

## The page grid

```
y 0.11 ┌ banner (white 28 pt on the template's dark-blue bar) ───────────────────────────┐
y 0.94 │▌subtitle: one declarative claim, 20 pt blue                                       │
y 1.56 ├─ divider rule  x = 0.45 … 12.88  (GRID_L … GRID_R) ────────────────────────────────┤
y 1.66 │ ▸ lead line 1   (optional, ≤ 3 lines, each exactly one line)                      │
       │ ▸ lead line 2                                                                      │
       │ body = page.body  → split(n) / vsplit(n) / take_top(h) …   gutter 0.23, gap 0.12   │
y 6.42 │                 takeaway: one red 19 pt sentence, centered (optional)              │
y 7.05 └──────────────────────────────────────────────────────────────────── page number ─┘
```

- `deck.content(banner, subtitle, lead=[...], takeaway="...")` draws the zones; `page.body` is what is left.
  Without lead the body starts at 1.78; without takeaway it ends at 7.00.
- Lead item: `("标签：", "一句话")`, or `("标签：", "一句话", "red")` for the one item to emphasise.
- Takeaway = the page's conclusion in one line. It is not a "承上启下 / 下一页将…" transition (banned).
- Split, never type coordinates: `left, right = page.body.split(2)`; weighted `split(2, [5, 7])`;
  rows `vsplit(3)`; strips `take_top(0.40)` / `take_bottom(0.80)` (return *(taken, rest)*).
- One frame level: a panel or card already has its border; never wrap it in a dashed box.
  `page.columns(2, dashed=True)` is only for columns of loose text and small figures.
- Copy length: `cjk_chars_per_line(width)` tells how many 18 pt characters fit one line — write to it.
- Hollow boxes read as "太空" just as overfull ones read as "太满": size short text with `card_height()`
  and spend the remaining height on a figure or flow.

## Components

| Need | Call | Notes |
|---|---|---|
| Figure under a title bar | `add_figure_panel(slide, region, "标题", png, tone, fit)` | draw at `figure_slot(region)`; `fit="crop"` for photos/GIFs |
| Picture filling a region | `add_picture_exact(slide, png, region, fit)` | same aspect rule |
| Header + text card | `add_card(slide, region, "标题", lines, tone)` | raises if text cannot fit at 18 pt |
| Header + image + caption | `add_image_card(slide, region, "标题", png, caption, tone)` | image size `image_card_slot(region, n_lines)` |
| Table | `add_table(slide, region, col_ratios, rows)` | row heights from wrapped text; raises if too tall |
| Equation | `add_omml_equation(slide, *region.box, xml, jc="center", fill_color=C_BLUE_LIGHT, line_color=C_BLUE)` | |
| Editable flow / mapping | `add_role_box`, `connect_boxes`, `add_flow_pipeline` (tju_native_diagrams) | |

`tone="blue"` by default; `tone="red"` for at most one focal item per row.

## Opening-report recipes

**背景** — `lead` = 2 lines (国家/战略需求, 行业需求与价值) · body `split(2)`: left real photo/GIF of the
scene as a figure panel (`fit="crop"`), right a concept card (≤ 3 one-line definitions) or a scene schematic
· `takeaway` = why this scene matters.

**关键挑战** — no lead · body `split(3)` → three `add_image_card`: short header (≤ 1 line), picture of the
failure, 2-line caption (existing assumption → what breaks) · `takeaway` = what kind of method is needed.
Aim ≤ 120 body characters.

**现状（一）方法演进** — a time × method-family figure filling `page.body` (`add_picture_exact`): classic
rule/optimisation methods first, then learning, then the newest family that your work belongs to; mark
where each family fails in your setting · `takeaway` = the gap.  This page answers "为什么不用传统方法".

**现状（二）代表工作** — body `split(2)` then each `vsplit(2)` → four cards, each = original paper figure
(cited `[n]`) + two one-line items `思路 / 局限`.

**研究目标与总体框架** — `lead` = 1 line `研究目标：…` · body `split(2, [5, 7])`: left an explicit
`挑战 → 研究内容` mapping (native boxes + connectors, editable), right a compact closed-loop/system figure
panel marking reused vs new modules.

**研究内容 N (one page each)** — `lead` = 3 lines, one per sub-content (label + one plain sentence saying
what is done; mark the key one red; no reuse disclaimers — scope is stated on the overview) · body `split(2)` → two figure panels (method schematic + mechanism/flow or expected-result plot),
or a panel + a card/equation column · `takeaway` = what this content delivers.  Subtitle begins with
`研究内容一：`.  Never split one content over (上)/(下) pages.

**拟创新点** — 2 cards, header `创新点一：<短名>`, body 1–2 plain sentences: what is new and what it enables
(optionally labelled one-line items 问题/方法/效果; see story.md S2b) · size the cards with `card_height()`
(stacked full-width or `split(2)` side by side) and give the rest of the body to a native
`问题 → 方法 → 预期` strip or a figure — never stretch a 1–2-line card over half the page · `takeaway` optional.

**研究基础与可行性** — body `split(2)`: left `研究条件（软硬件）` (platform photo panel, or a card sized with
`card_height()` / `body_anchor=MSO_ANCHOR.TOP` + a staged-verification flow), right `研究基础（已有工作）`
(`vsplit(3)` small cards, one line each) · `takeaway` optional.  No parameter/constraint tables on this page.

**研究计划与预期成果** — `add_table` (时间 / 任务 / 成果) + a short measurable-outcome card if space allows.

**目录 / 章节过渡** — `deck.toc(items, active=k)`; **封面 / 致谢** — `deck.cover(...)`, `deck.thanks()`.

## 中期 / 答辩 variants
- 中期: after 总体框架, one page per content showing `计划 → 进展 → 结果 → 偏差与调整`; keep the same names.
- 答辩: results pages use `split(2)` with a plot panel + a conclusion card; the takeaway states the measured
  gain with its condition.

## Editing an existing deck
Back up the `.pptx` first.  Prefer regenerating from the deck's generator script; for in-place edits use
`setup_tju_content_slide` / `clear_slide_body` and rebuild the body with Regions computed from the grid
(`tju_layout.full_width(y, h).split(...)`), then `enforce_min_18pt_all_slides` + `save_staged_pptx` + audit.
