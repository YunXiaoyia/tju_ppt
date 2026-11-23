# Narrative & Content Principles for Academic Talks

Applies to opening reports (开题), mid-term checks (中期), thesis defenses (答辩), and group-meeting talks.
Typography rules (SKILL.md §1–§4) make a slide *legal*; these rules make it *understood*.

---

## P1. Design for the smart non-expert
- Assume at least half the committee works in an adjacent field and has **never seen your scene**.
- Test: could a professor from a neighbouring lab follow slide *N* using only slides 1…*N*−1 and the talk?
- A written report (开题报告/论文) carries the details; the slides carry the **argument**. Never paste
  report paragraphs, symbol sets, or full ablation matrices onto slides.

## P1b. Accessible ≠ childish: keep an academic register
The audience is **outside your sub-field, but they are professors**. Lower the *entry barrier*, never
the *register*.
- **Simplify by defining, not by paraphrasing.** Use the correct technical term and give a one-line
  definition once (`概念界定` card), then reuse the term consistently. Do not replace terms with folk metaphors.
- **Avoid on slides** (fine to explain conversationally while speaking, never to print): slang, folk
  metaphors, and marketing buzzwords (e.g., 草稿台 / 算总账 / 痛点 / 搞定 / 写不完 / 抓不准 / 认不出 /
  为什么难), superlatives in parentheses (`（最快）` / `（最灵活）` / `（最稳妥）`), exclamation marks, emoji.
- **Use neutral academic labels:** `问题` / `方法` / `局限` / `借鉴` / `作用` / `判据` / `优势` / `预期收益` / `决策原则`;
  write subtitles as concise declarative findings rather than quiz questions aimed at beginners.
- **Quantify instead of dramatizing:** "厘米级偏移" rather than "歪几厘米"; "状态空间组合爆炸" rather than
  "组合上千种、规则写不完".
- **Speech script** may be warmer and more explanatory than the slide, but stays in spoken academic
  Chinese (no "大坑 / 偷懒 / 瞎猜").
- **Written-report test:** would this printed sentence be acceptable in the written 开题报告 (minus citations)?
  If not, rewrite. `tju_audit.py --content` flags common colloquialisms.

## P2. Fixed story arc: Background → Difficulty → Problem → Others → Mine
| Step | Question the audience has | Typical slides |
|---|---|---|
| 1. Background / scene | "What does this world look like? What are the key terms?" | real photo/video or scene schematic + 2–3 concise concept-definition cards |
| 2. Specificity & difficulty | "Why is this hard *here*? Why do existing methods break? Why could the new mechanism work?" | concrete failure case, existing-method vs proposed-approach contrast |
| 3. Abstract problem | "So what exactly is the research question?" | one boxed sentence derived directly from step 2 |
| 4. Related work & limits | "Hasn't someone solved it?" | 1–2 representative works per direction: idea → limit → what I borrow |
| 5. My work | "What do you do, and how do the parts fit?" | system overview figure first, then one slide per research topic |
| 6. Validation | "How will you prove it?" | baselines under identical conditions, test scenarios, metrics, statistics |
| 7. Innovation, feasibility, plan | "What is new, can you finish it, by when?" | `问题 → 方法` innovation cards; existing platform; dated timeline |

- **Mid-term:** replace steps 4–5 with *plan vs progress → results so far → deviations and fixes*.
- **Defense:** lead with main results; keep background to ≤ 15 % of the time.
- **Subtitle-only test:** read only the slide subtitles in order. They must tell the whole story as
  declarative claim sentences ("X fails under Y"), not generic topic labels ("Related work").

## P3. Define before use (jargon budget)
- Every domain term, acronym, and internal code name is **defined with a figure or a one-line sentence
  before** it appears in an argument. Keep a glossary while drafting; ≤ 3 new terms per slide.
- Replace internal IDs (rule numbers, loss-term symbols, ablation branch codes) with what they *do*,
  unless the committee specifically needs to cite them.
- English acronyms: spell out once at first use (e.g., `模型预测控制（MPC）`), then use consistently.

## P4. Show why it is hard before showing how you solve it
- Use one concrete, visual failure case from the real scene (photo/GIF/diagram of the failure) rather than
  abstract adjectives ("复杂、多变").
- Contrast explicitly: *existing assumption → what actually happens → consequence*.
- Explain **why** the proposed mechanism works, not merely that it is "more intelligent" or "data-driven".
- The research question must be the logical residue of this contrast, not an unrelated idea dropped in.

## P5. One slide, one message, speakable in its time slot
- Budget words by speaking time: spoken Chinese ≈ 200–230 字/min; on-slide text should be well under
  what you say. Rule of thumb for a 40–60 s content slide: **≤ 150–200 visible units** including
  tables (`tju_audit.py --content` warns above 200).
- If the speech script for a slide overflows its time budget, the slide is too dense: cut or split.
- Use concise, structured declarative column headings (`1. 状态表示：…` / `2. 触发判据：…`) that guide the
  audience from problem to mechanism without slipping into conversational quiz questions (see P1b).

## P6. Picture over prose
- Every content slide carries at least one visual occupying ≈ 35–45 % of the area:
  real photos/video GIFs for the scene, **original paper figures** (cited) for related work,
  native diagrams or self-drawn schematics (SKILL.md §5) for your method, plots for results.
- Keep ≤ 2–3 core formulas in the whole talk; each is followed by a concise variable/meaning breakdown.
- Replace bullet walls with structure: flow (`input → core → output`), contrast pairs, joined cards, small tables.

## P7. Related work: compress and trace
- Group by research direction; 1–2 representative works per direction, each with its figure and three
  concise lines: **idea (`思路`) / limit in *my* setting (`局限`) / what I borrow (`借鉴`)**. Avoid acronym walls
  (≤ 6 works on slides; the full literature survey lives in the written report).
- Every limitation listed must be answered by one of your contributions, and every contribution must
  answer a stated limitation or difficulty (two-way traceability).

## P8. Progressive disclosure for your own work
- First an overview figure placing each research topic in the full system and marking **reused vs new**
  modules (individual thesis: interfaces, not team task lists).
- Then one slide per research topic, in the same order and with the same names as the overview.
- Innovation points summarize the **`问题 → 方法`** mapping; they must not introduce unmentioned topics.

## P9. Claim hygiene (what committees attack first)
- Every number on a slide traces to a citation or your own measurement; state the test condition
  (hardware, dataset, setting).
- Check the inference, not just the number: if you argue "method A fails due to property P", your fix
  must actually address P (e.g., distilling to the *same parameter scale* does not fix edge latency by
  itself; the real lever is call frequency, output length, or quantization).
- Avoid unproven absolutes (`彻底` / `杜绝` / `100 %` / `零误差` / `绝对安全`) unless mathematically proven;
  prefer calibrated terms (`逼近` / `降低` / `满足`).
- Expected outcomes are measurable (`metric + baseline + direction`), never vague ("显著提升性能").

## P10. Opening-report completeness checklist
- [ ] Research question stated in one sentence
- [ ] Related work with limitations mapped two-way to your contributions
- [ ] System overview + per-topic method slides
- [ ] **Validation plan**: baselines under identical conditions, test scenarios, metrics, statistical protocol
- [ ] Innovation points (2–3, each traceable to a stated difficulty/limit)
- [ ] Feasibility: existing platform/data, staged verification, main risks and fallbacks
- [ ] Timeline with **concrete dates** consistent with the degree schedule
- [ ] Expected outcomes: deliverables + measurable targets (+ papers/patents if required)
- [ ] References for every cited figure/number; all verified

## P11. Consistency across artifacts
- Title, terminology, numbers (model size, dates, metrics), and page numbers must match across slides,
  speech script (讲稿), and the written report. After inserting or deleting a slide, renumber the
  script and all page references.
- Keep a per-slide speech script with realistic word-count-backed time budgets; total time should leave ≥ 10 % slack.
