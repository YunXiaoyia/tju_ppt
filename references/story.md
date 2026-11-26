# Story & Content Rules

Applies to 开题 / 中期 / 答辩 / 组会.  Layout rules make a slide legal; these make it understood.
Layout recipes for each page type are in `page_recipes.md`.

## S1. Audience: smart non-experts
- At least half the committee works in an adjacent field and has never seen your scene.
  Test: could a professor from the next lab follow slide *N* using only slides 1…*N*−1 and the talk?
- Slides carry the **argument**; the written report carries the details. Never paste report paragraphs,
  symbol tables, or full ablation matrices.

## S2. Academic register, low entry barrier
- Lower the barrier by **defining** the correct term once (picture or one line), then reuse it. Never
  replace terms with slang or metaphors.
- Printed text is short and plain: a phrase a professor from the next lab understands at first read. It need
  not be a full report sentence, but it must not be telegraphic shorthand only you can expand
  (`静态顶吸无残垛` → `仅静态吸取，无残垛`). Spoken script may be warmer, still academic.
- Name things the way a thesis does: `研究内容一/二/三` (not 工作一/二/三), `任务中断` (not 卡壳),
  `难以求解 / 方法失效` (not 搞不定), the device's formal name (not the workshop nickname).
- Use neutral labels: 问题 / 方法 / 局限 / 借鉴 / 判据 / 优势 / 预期收益. No exclamation marks, no
  `（最快）`-style superlatives, no emoji. `tju_audit.py --content` lists the common offenders with replacements.
- Quantify instead of dramatising: `厘米级偏移`, not `歪几厘米`.

## S2b. Plain, not defensive
Write what the work does, directly.  Text written to pre-empt objections reads as AI-generated and hides the idea.
- **State scope once.** Say which modules are reused (perception, low-level control) on the overview page; do
  not repeat `复用既有… / 直接调用既有…` on every later page.
- **No revision history on slides.** Notes such as `（去 VLA/ACT）`, `吸收…`, `不再…` tell the advisor what changed;
  the committee never saw the old version.  Put change notes in the speech or the reply to the advisor.
- **No hedge stacks or pre-emptive rebuttals.** One calibrated verb is enough (`降低`, `提高`); avoid stacking
  `有望…并…`, straw-man contrasts, and conditions stuffed into the first clause (`同一底层下相较…`).  The
  innovation sentence form in S9 is the one place where `区别于 / 有望` is expected.
- **Research-content sub-items start with `研究`** (`研究……方法 / 研究……机制`), name the research question, and
  carry no citations; engineering labels such as `输出形式 / 接口` are not research contents.
- **Everyday technical words over packaged jargon.** Prefer the plain name of the mechanism
  (`状态编码`, `按需调用`, `每包耗时`) over coined or lab-internal labels (`前缀编码`, `双态唤醒`, `节拍`,
  group project names, configuration codes such as `(11-10-9)×12`).  If a code or project name is necessary,
  define it once on the slide where it first appears.
- **One name per mechanism.** Pick one term and keep it on every page, figure, and the script.
- Implementation detail (bit widths, engine names, loss tricks) stays in the speech or report unless it is the
  contribution itself.

## S3. Story arc
| Step | Audience question | Pages (recipes in page_recipes.md) |
|---|---|---|
| 1 Background | What is this world? Why does it matter? | 背景 |
| 2 Difficulty | Why is it hard *here*? Why do existing methods break? | 关键挑战 |
| 3 Related work | Hasn't someone solved it? | 现状-演进, 现状-代表工作 |
| 4 My work | What exactly will you do, how do parts fit? | 目标与总体框架, one page per 研究内容 |
| 5 Proof & plan | What is new, can you finish, by when? | 创新点, 研究基础与可行性, 计划与预期成果 |

- 中期: replace 3–4 with *plan vs progress → results so far → deviations and fixes*.
- 答辩: lead with results; background ≤ 15 % of time.
- **Subtitle-only test:** read the subtitles in order; they must tell the whole story as declarative claims
  (`X 在 Y 条件下失效`), not topic labels (`相关研究`) or quiz questions.

## S4. Define before use
- Every domain term, acronym, internal code name gets a picture or a one-line definition before it is used
  in an argument; ≤ 3 new terms per slide. Write acronyms as `全称（ABBR）` at first use — the audit
  recognises that form.  Product/model names (PX4, Jetson Orin NX, RTX 4090) have no expansion: state their role
  (飞控 / 机载计算机 / 训练机) and pass them with `known=` / `--known`.
- Replace internal IDs (rule numbers, loss symbols, branch codes) with what they do.

## S5. One slide, one message, speakable in its slot
- Spoken Chinese ≈ 200–230 字/分. Plan at 215 and leave ≥ 10 % slack: a 10-min talk ≈ 1,900–2,000 spoken 字.
- On-slide text stays well under what you say on that slide (rule of thumb ≤ 60 %). The audit warns above
  160 body units (banner/subtitle excluded); a 40 s slide should be far below that.
- If the script for a slide overruns its slot, the slide is too dense: cut or split. Check with
  `tju_script_check.py script.md --budget-s <slot>`.

## S6. Picture over prose
- Every content slide carries a visual: real photo / GIF of the scene, cited original paper figure, native
  diagram, self-drawn schematic, plot, or a table where a table is the natural form.
- ≤ 2–3 core formulas in the whole talk, each followed by a variable/meaning line.
- Replace bullet walls with structure: flow, contrast pair, cards, small table.

## S7. Related work: compress and trace
- First answer the cross-field question "why not the classic methods?" with a method-evolution page
  (time × method family), then 3–4 representative recent works, each with its figure and two short lines:
  `思路 / 局限（在我的场景下）` (optionally `借鉴`). ≤ 6 works on slides; the survey lives in the report.
- Two-way traceability: every limitation listed is answered by one of your research contents, every content
  answers a stated limitation or difficulty.

## S8. Your work: overview, then one page per content
- Overview page: research goal in one sentence → explicit `挑战 → 研究内容` mapping → compact system/closed-loop
  figure marking reused vs new modules (individual thesis: interfaces, not team task lists).
- Then exactly one page per research content, same order and names as the overview. A content that does not
  fit one page is not yet condensed enough.
- Do not repeat the same mechanism (e.g. the same training trick) across two contents; a deployment content
  needs a real lever (quantisation, pruning, call-frequency control), not only "trigger / sleep".

## S9. Innovation, feasibility, plan
- Innovations: summarise `问题 → 方法` and must not introduce unmentioned topics. Count follows substance —
  for a master's thesis usually **2** (1–3), and they need not map 1:1 onto the research contents. Write each
  in the committee's standard form `针对……问题，区别于已有的……，提出了……，有望……` (advisor-required), filled with
  plain words: a concrete problem, one real prior approach (not a straw man), the method by its plain name, and one
  measurable expected effect; ≤ 3 lines.
- Feasibility = **研究条件** (hardware, software, data, compute) + **研究基础** (prior algorithms and engineering
  work of the group and yourself). Parameter tables and constraint lists do not belong here.
- Plan: dated phases consistent with the degree schedule + measurable outcomes (`指标 + 基线 + 方向`).

## S10. Claim hygiene
- Every number traces to a citation or your own measurement with its test condition.
- Check the inference, not just the number: if A fails because of property P, your fix must address P.
- No unproven absolutes (`彻底 / 杜绝 / 100 % / 零误差`); prefer `降低 / 逼近 / 满足`.

## S11. Opening-report completeness (full deck; for an excerpt check only the items it contains)
- [ ] Research question in one sentence
- [ ] Related work with limitations mapped two-way to contents
- [ ] Overview + one page per content
- [ ] Validation plan: baselines under identical conditions, scenarios, metrics, statistics
- [ ] Innovations traceable to difficulties
- [ ] Feasibility (conditions + foundations), main risks and fallbacks
- [ ] Timeline with concrete dates; measurable expected outcomes
- [ ] References for every cited figure/number

## S12. Consistency across artifacts
Title, terms, numbers, page numbers match across slides, speech script, and written report. After inserting
or deleting a slide, renumber the script and all page references.
