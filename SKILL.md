---
name: tju-ppt
description: Use when creating, editing, auditing, or revising Tianjin University (TJU / 天津大学) group academic PowerPoint (.pptx) decks — 开题汇报, 中期汇报, 学位答辩, 组会 — or their speech scripts (讲稿); when drawing defense figures (flowcharts, architectures, timelines, plots); or when a deck is criticized as 字太多 / 太满 / 图和标题栏对不齐 / 听不懂 / 口语化 / 逻辑顺序不对 / 超时.
---

# TJU Academic Deck

## Overview
A defense deck is an argument for professors outside your sub-field, built on the TJU 16:9 template.
Three layers, in this order:
1. **Story** — what each page claims and how long it is spoken (`references/story.md`).
2. **Page recipe** — the default shape of each page type (`references/page_recipes.md`).
3. **Build** — `scripts/tjuppt.py`: geometry comes from a grid and every component checks fit at build
   time, so most layout rules cannot be broken silently.  Audit, render, and script timing close the loop.

Not for: Beamer (`paper-slides`), posters (`paper-poster`), SVG-first decks (`ppt-master`), pitch decks.

## Workflow
1. **Story.** Write all subtitles in order as declarative claims (they must read as the whole argument)
   and give every page a time slot.  Check S3/S11 of `story.md`.
2. **Recipe + copy.** Pick the recipe for each page; write lead lines, captions, and the takeaway to length
   (`cjk_chars_per_line(width)`).
3. **Build** a generator script (keep it next to the deck):
   ```python
   import sys; sys.path.insert(0, "<skill_dir>/scripts")
   from tjuppt import *
   deck = Deck()
   p = deck.content("研究内容与方案", "研究内容一：……",
                    lead=[("标签：", "一句话"), ("标签：", "一句话", "red")], takeaway="本页结论一句话")
   left, right = p.body.split(2)
   add_figure_panel(p.slide, left, "图题", draw_my_figure(*figure_slot(left)))   # figure drawn at slot size
   add_card(p.slide, right, "要点", [[("• 判据：", C_BLUE), ("……", C_BLACK)]])
   deck.save("deck.pptx")          # enforce 18 pt → staged save → audit (cover/TOC/thanks auto-skipped)
   ```
   Figures: `references/figures.md`.  Low-level traps: `references/pptx_pitfalls.md`.
4. **Audit:** 0 issues.  Fix every warning or state why it stays; pass already-explained acronyms and product
   names via `deck.save(..., known=("PID", "PX4"))` / `--known PID,PX4`.
5. **Render** `python scripts/tju_render.py deck.pptx out/` and look at every changed slide PNG.
6. **Speech script** with a time label per page heading (`### P4｜关键挑战（约 60 秒）`), then
   `python scripts/tju_script_check.py script.md --budget-s <slot>` → no flags.
7. **Editing an existing deck:** back up the `.pptx` first; change the generator script and rebuild rather
   than patching shapes by hand.

## Contract (enforced by code)
| Rule | Enforced by |
|---|---|
| Every run ≥ 18 pt, bold, 微软雅黑 (CJK) + Times New Roman (Latin), 4 font slots | `set_run_font`, `Deck.save`, audit |
| Colors: `#00337F` blue, `#C00000` red, black, white + `#EAF1FB` light-blue tint | `tju_theme`, audit (deck + `--images`) |
| Content edges on the grid `0.45 … 12.88 in`, one gutter | `Region.split`, audit |
| Image under a header bar is flush with it (same left/width, touching) | `add_figure_panel` / `add_image_card` aspect check (`fit="crop"` only for photos), audit |
| ≤ 3 lead lines, each exactly one line | `ContentPage` → `LayoutError` |
| Takeaway: one line, 19 pt, red, centered, one per page | `ContentPage` → `LayoutError`, audit |
| Text fits its box at 18 pt | `add_card` / `add_image_card` / `add_table` → `LayoutError` |
| No "承上启下" transition text; speaker notes are empty; no raw `x_k` math | audit |
| Editable equations, hats via U+0302 | `tju_omml` |
| Figures: slot size, palette, no tofu glyphs, no overlaps/crossings | `TjuDiagram.save` strict, audit `--images` |
| Speech ≈ 215 字/分 with ≥ 10 % slack | `tju_script_check.py` |

## Judgment rules (you check)
- Define every term/acronym before use (`全称（ABBR）`); ≤ 3 new terms per page.
- One message per page; body text far below what is spoken (audit warns > 160 units).
- Every content page has a visual; ≤ 2–3 core formulas in the whole talk.
- Thesis vocabulary: 研究内容一/二/三, 任务中断, formal device names; no slang or absolutes.
- Plain, not defensive (`story.md` S2b): state reuse/scope once, no revision notes like `（去 X）`, no hedge
  stacks or straw-man contrasts, no lab-internal codes or coined labels; one name per mechanism.
- One frame level: never put a panel or card inside a dashed box.
- Innovations follow substance (master's: usually 2), not one per research content; write them as
  `针对……问题，区别于已有的……，提出了……，有望……` in plain words.
- Research-content pages: three sub-items starting with `研究…`, no citations.
- References fit on one page (the only page allowed below 18 pt).
- Feasibility = 研究条件 + 研究基础; no parameter tables there.
- Same names and numbers across overview, content pages, innovations, script, and report.

## When a LayoutError fires
Shorten or restructure the content.  Never lower the font size, never switch to coordinate primitives
(`add_text_block`, `add_image_fit`) to get around the check — that reintroduces exactly the overflow and
misalignment the advisor complained about.

| Temptation | Do instead |
|---|---|
| "Make the takeaway 18 pt so it fits" | Cut the sentence to one line at 19 pt |
| "`add_image_fit` is close enough under the header" | Redraw at `figure_slot()` size, or `fit="crop"` for photos |
| "A dashed box around the cards looks tidy" | Drop it; the cards already have frames |
| "Three contents → three innovations" | Write the 1–2 that are genuinely new |
| "Add `（去 VLA）` / `复用既有…` so nobody thinks we overclaim" | State scope once on the overview; change notes go in the speech |
| "Compress the phrase so it fits one line" | Rewrite in plain words or give it more room; shorthand nobody can expand is not shorter |
| "Split 研究内容二 into (上)/(下)" | Condense to one page; detail goes to the speech |
| "Skip pages 1,2 in the audit" for a partial deck | `--skip` lists only cover/TOC/divider/reference/thanks pages; `Deck.save` computes it |

## Files
- `scripts/tjuppt.py` — single import.  Modules: `tju_theme` (palette, grid, width budget) · `tju_layout`
  (`Region`) · `tju_text` · `tju_components` · `tju_omml` · `tju_deck` (`Deck`, `ContentPage`, template ops) ·
  `tju_native_diagrams` · `tju_svg_diagram` · `tju_plot_style`.
- CLIs: `tju_audit.py`, `tju_render.py`, `tju_script_check.py` (each has `--help`).
- `scripts/ppt_builder_helpers.py` — compatibility shim for pre-v2 deck scripts.
- `references/` — `story.md`, `page_recipes.md`, `figures.md`, `pptx_pitfalls.md`, `tju_template.pptx`,
  `tju_mermaid_theme.json`.
- `examples/build_example_deck.py` — every recipe in a 12-page opening deck (audit 0 issues);
  `examples/diagram_showcase.py` — all figure routes.  `tests/test_tjuppt.py` — contract tests.
