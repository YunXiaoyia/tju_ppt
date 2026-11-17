# Narrative & Content Principles for Academic Talks

Applies to opening reports (开题), mid-term checks (中期), thesis defenses (答辩), and group-meeting talks.
Typography rules (SKILL.md §1–§4) make a slide *legal*; these rules make it *understood*.
Source: advisor feedback on a rejected opening deck ("听众是小白、专业词太多、每页太满、字太多、
先讲背景和难题再讲别人和自己") plus recurring committee questions in pre-defense reviews.

---

## P1. Design for the smart non-expert
- Assume at least half the committee works in an adjacent field and has **never seen your scene**.
- Test: could a teacher from a neighbouring lab follow slide *N* using only slides 1…*N*−1 and the talk?
- A written report (开题报告/论文) carries the details; the slides carry the **argument**. Never paste
  report paragraphs, symbol sets, or full ablation lists onto slides.

## P1b. Accessible ≠ childish: keep an academic register
The audience is **outside your sub-field, but they are professors**. Lower the *entry barrier*, never
the *register*. Second-round advisor feedback on an over-simplified deck: "太多口水话、太小白了，
是给老师汇报，不是给小学生".
- **Simplify by defining, not by paraphrasing.** Use the correct technical term and give a one-line
  definition once (概念界定 card), then reuse the term. Do not replace terms with folk metaphors.
- **Banned on slides** (fine to *say* sparingly, never to print): slang and metaphors such as
  草稿台 / 算总账 / 看图 / 手眼微调 / 抓不准 / 认不出 / 写不完 / 顺手 / 松绑 / 上车 / 挡死 / 杜绝盲抓 /
  “科普”, rhetorical labels such as 为什么难 / 连锁反应 / 痛点 / 突破, superlatives in parentheses
  (最快 / 最灵活 / 最稳妥), exclamation marks, emoji.
- **Use neutral academic labels:** 问题 / 方法 / 局限 / 借鉴 / 作用 / 判据 / 优势 / 预期收益 / 决策原则;
  subtitles as declarative findings ("固定规则难以适应动态垛况"), not as quiz questions aimed at children.
  Question-style *column* headings are allowed only when phrased technically ("重规划触发：预测与观测偏差判据").
- **Quantify instead of dramatizing:** "厘米级偏移" not "歪几厘米"; "状态空间组合爆炸" not
  "组合上千种、规则写不完".
- **Speech script** may be warmer and more explanatory than the slide, but stays in spoken academic
  Chinese (no "大坑 / 偷懒 / 瞎猜").
- Check: would this sentence be acceptable in the written 开题报告 (minus citations)? If not, rewrite.
  `tju_audit.py --content` flags common colloquialisms.

## P2. Fixed story arc: Background → Difficulty → Problem → Others → Mine
| Step | Question the audience has | Typical slides |
|---|---|---|
| 1. Background / scene | "What does this world look like? What are the words?" | real photo/video + 2–3 plain-language term cards |
| 2. Specificity & difficulty | "Why is this hard *here*? Why do current rules/methods break? Why could the new tool work?" | concrete failure example, old-way vs new-way contrast |
| 3. Abstract problem | "So what exactly is the research question?" | one boxed sentence, derived from step 2 |
| 4. Others' work & limits | "Hasn't someone solved it?" | 1–2 representative works per direction: idea → limit → what I borrow |
| 5. My work | "What do you do, and how do the parts fit?" | overview figure first, then one slide per work |
| 6. Validation | "How will you prove it?" | baselines, scenarios, metrics, statistics |
| 7. Innovation, feasibility, plan | "What is new, can you finish it, by when?" | pain → breakthrough cards; platform; dated timeline |

- Mid-term: replace 4–5 with *plan vs progress → results so far → deviations and fixes*.
- Defense: lead with results; keep background to ≤ 15 % of the time.
- **Subtitle-only test**: read only the slide subtitles in order. They must tell the whole story as
  claim sentences ("X fails because Y"), not topic labels ("Related work").

## P3. Define before use (jargon budget)
- Every domain term, acronym, and internal code name is **defined with a picture or a plain sentence
  before** it appears in an argument. Keep a glossary while drafting; ≤ 3 new terms per slide.
- Replace internal names (rule IDs, symbol names such as Φ_frag, branch codes A1–A9) with what they
  *do*, unless the committee needs to cite them.
- English acronyms: spell out once at first use (VLM = 视觉语言模型), then use consistently.

## P4. Show why it is hard before showing how you solve it
- Use one concrete, visual failure case from the real scene (photo/GIF of the problem) rather than
  abstract adjectives ("复杂、多变").
- Contrast explicitly: *old assumption → what actually happens → consequence*.
- Explain **why** the proposed tool can work (mechanism), not only that it is "more intelligent".
- The research question must be the logical residue of this contrast, not a new idea dropped in.

## P5. One slide, one message, speakable in its time slot
- Budget words by speaking time: spoken Chinese ≈ 200–230 字/min; on-slide text should be well under
  what you say. Rule of thumb for a 40–60 s content slide: **≤ 150–200 visible characters** including
  tables (`tju_audit.py --content` warns above 200).
- If the speech script for a slide overflows its time budget, the slide is too dense: cut or split.
- Prefer question-style column headings ("怎么看清现场？""什么时候重排？") — they mirror the
  audience's own questions and enforce shallow-to-deep order.

## P6. Picture over prose
- Every content slide carries at least one visual occupying ≈ 40 % of the area:
  real photos/video GIFs for the scene, **original paper figures** (cited) for related work,
  self-drawn schematics (§5 diagram routes) for your method, plots for results.
- Keep ≤ 2–3 core formulas in the whole talk; each is followed by a one-line plain reading
  ("判据含义：…").
- Replace bullet walls with structure: flow (input → core → output), contrast pairs, cards, small tables.

## P7. Related work: compress and trace
- Group by direction; 1–2 representative works per direction, each with its own figure and three
  lines: **idea / limit in *my* setting / what I borrow**. No acronym walls (≤ 6 works on slides;
  the rest live in the report).
- Every limitation you list must be answered by one of your contributions, and every contribution
  must answer some limitation or difficulty (two-way traceability).

## P8. Progressive disclosure for your own work
- First an overview figure that places each work in the full system and marks **reused vs new**
  modules (individual thesis: interfaces, not team task lists).
- Then one slide per work, in the same order and with the same names as the overview.
- Innovation points repeat the *pain → breakthrough* mapping; they must not introduce new content.

## P9. Claim hygiene (what committees attack first)
- Every number on a slide traces to a citation or your own measurement; state the condition
  (hardware, setting).
- Check the inference, not just the number: if you argue "model A is too slow on device D", the fix
  must not share A's property (e.g. a distilled model of the *same size* does not fix latency by itself;
  the real lever is call frequency, output length, or quantization).
- Avoid absolutes (最优 / 彻底 / 杜绝 / 100 % / 首次) unless proven; prefer "逼近 / 降低 / 满足".
- Expected outcomes are measurable (metric + baseline + direction), never vague ("显著提升性能").

## P10. Opening-report completeness checklist
- [ ] Research question stated in one sentence
- [ ] Related work with limitations mapped to your contributions
- [ ] Method overview + per-work slides
- [ ] **Validation plan**: baselines under identical conditions, test scenarios, metrics, statistics
- [ ] Innovation points (2–3, each traceable)
- [ ] Feasibility: existing platform/data, staged verification, main risks and fallbacks
- [ ] Timeline with **concrete dates** consistent with the degree schedule
- [ ] Expected outcomes: deliverables + measurable targets (+ papers/patents if required)
- [ ] References for every cited figure/number; all verified

## P11. Consistency across artifacts
- Title, terms, numbers (model size, months, metrics), and page numbers must match across slides,
  speech script (讲稿), and the written report. After inserting or deleting a slide, renumber the
  script and all page references.
- Keep a per-slide speech script with time budgets; total time should leave ≥ 10 % slack.
