# PPTX Pitfalls (python-pptx, PowerPoint, LibreOffice)

Low-level traps the helpers already handle.  Read when editing XML directly or debugging a rendering
difference.

## Text & fonts
- `run.font.name` sets only `a:latin`; CJK falls back to the theme font.  `set_run_font` writes all four
  slots (`latin/ea/cs/sym`) — use it (or `enforce_min_18pt_all_slides`) for any run you create by hand.
- LibreOffice wraps CJK slightly wider than PowerPoint.  A line that fits in the LibreOffice render fits in
  PowerPoint; not necessarily the reverse.  The width budget (`tju_theme.text_width_in`) is conservative.
- Bold 微软雅黑 line height ≈ 1.3 × size; budget boxes with that, not 1.0 ×.
- LibreOffice adds CJK↔Latin auto-spacing, so renders show `调整  MPC  权重` / `（ MPC ）` even when the source
  has one or no space.  It is a renderer artifact; do not "fix" it by deleting spaces in the source.
- Ambiguous-width punctuation (em dash `—`, `…`, `·`) is drawn with the Latin font and looks thin next to
  bold 微软雅黑.  In CJK sentences prefer `、` or `→`, or full-width forms.
- Text inside OMML is set in the math font (thin, not 微软雅黑).  Keep Chinese notes out of equations; put them
  in a normal text line next to the equation.

## Shapes
- Autoshapes carry a `<p:style>` whose `effectRef` can add the theme shadow; the helpers zero it
  (`no_theme_effects`).  Do the same for any shape you add by hand.
- `add_shape()` without explicit fill/line inherits the theme (`fillRef/lnRef` → accent colors).  Always set
  `fill.solid()+color` or `fill.background()`, and `line.color` or `line.fill.background()`.
- Do not stack a background rectangle behind a picture or text box; put the border on the same shape
  (`pic.line`, text-box fill).  The audit counts declared-bbox overlaps.
- python-pptx shape wrappers are regenerated on each access: compare `sh._element is el`, never `sh is obj`.
- Tables: remove `tableStyleId` (theme colors leak) and write cell borders explicitly; rows grow to fit text
  regardless of the declared height — `add_table` sizes rows from the wrapped text.

## Equations (OMML)
- Wrap in `<a14:m>`; PowerPoint re-saves it inside `mc:AlternateContent` with a picture fallback.  Clear the
  wrapper when rebuilding a slide (`setup_tju_content_slide` / `clear_slide_body` do) or ghost shapes remain.
- python-pptx does not list shapes wrapped in `mc:AlternateContent`: after a deck is saved by PowerPoint its
  equations disappear from `slide.shapes` (they are still in the XML).  Geometry checks then skip them, and a
  rebuild from the generator script overwrites any manual edits made in PowerPoint — port those edits into the
  script first.
- Hats: combining U+0302 via `m_acc`; ASCII `^` sinks onto capital letters.
- LibreOffice converts OMML to StarMath: escape `<`/`>`, keep parentheses balanced inside one run, use `‖`
  rather than `|` or `[ ]` inside `<m:t>`, avoid StarMath keywords (`equiv`, `succ`) as bare subscripts.

## Pictures
- GIFs keep animating in PowerPoint; LibreOffice renders the first frame.  Crop via `pic.crop_*`
  (`fit="crop"`) instead of re-encoding the GIF.
- Embedding very large photos bloats the deck; downscale to ≤ 2× the placed size before inserting.

## Saving & rendering
- Save through `save_staged_pptx` (temp file → copy): a deck open in PowerPoint on Windows otherwise aborts
  the build; on `PermissionError` it writes `<name>_staged.pptx`.
- Windows exact render: PowerPoint COM `SaveAs(pdf, 32)` then PyMuPDF; always `Close()`, `Quit()`,
  `gc.collect()` or PowerPoint stays resident.  `tju_render.py` does this.
- Linux render: `libreoffice --headless --convert-to pdf` + `pdftoppm -r 144`.
