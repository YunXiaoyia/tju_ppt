#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tju_script_check.py — check a speech script (讲稿) against its time budget.

    python tju_script_check.py script.md [--rate 215] [--max-rate 230] [--budget-s 600]

Sections are Markdown headings that carry a time label, e.g.
    ### P4｜关键挑战（约 60 秒）      ### P9 研究内容一 (55 s)      ## 第 3 页 · 1 分钟
Spoken characters = CJK characters + Latin words/numbers in the section body (headings, block quotes
markers, code spans, and lines starting with '>' cues like （指图） are counted as spoken text except
parenthesised stage directions).

Reports per section: characters, labelled seconds, implied 字/分; flags sections above --max-rate and
the total if it leaves < 10 % slack against --budget-s (or the sum of labels when no budget is given).
Exit code 1 when anything is flagged.
"""

import argparse
import re
import sys

_HEAD = re.compile(r"^(#{1,6})\s+(.*)$")
_TIME = re.compile(r"(\d+(?:\.\d+)?)\s*(秒|s\b|sec|分钟|min)", re.I)
_STAGE = re.compile(r"[（(][^）)]{0,30}[）)]")          # （指左图） stage directions are not spoken
_WIDE = re.compile(r"[\u4E00-\u9FFF]")


def spoken_units(text):
    text = _STAGE.sub("", text)
    text = re.sub(r"`[^`]*`", "", text)
    return len(_WIDE.findall(text)) + len(re.findall(r"[A-Za-z]+|\d+(?:\.\d+)?", text))


def parse_sections(md):
    sections, cur = [], None
    for line in md.splitlines():
        m = _HEAD.match(line)
        if m and len(m.group(1)) == 1:          # H1 = document title, never a timed section
            cur = None
            continue
        if m:
            t = _TIME.search(m.group(2))
            if t:
                secs = float(t.group(1)) * (60 if t.group(2).lower() in ("分钟", "min") else 1)
                cur = {"title": m.group(2).strip(), "seconds": secs, "body": []}
                sections.append(cur)
            else:
                cur = None if len(m.group(1)) <= 2 else cur   # a new top-level section without time ends counting
            continue
        if cur is not None:
            cur["body"].append(line)
    for s in sections:
        s["units"] = spoken_units("\n".join(s["body"]))
    return sections


def main():
    ap = argparse.ArgumentParser(description="speech script timing check")
    ap.add_argument("script")
    ap.add_argument("--rate", type=float, default=215.0, help="planning speech rate, 字/分 (default 215)")
    ap.add_argument("--max-rate", type=float, default=230.0, help="flag sections faster than this")
    ap.add_argument("--budget-s", type=float, default=None, help="total time slot in seconds")
    a = ap.parse_args()
    secs = parse_sections(open(a.script, encoding="utf-8").read())
    if not secs:
        print("[script] no timed headings found (expected e.g. '### P3｜背景（约 45 秒）')")
        sys.exit(1)
    flagged = False
    print(f"{'section':<40} {'字':>5} {'s':>5} {'字/分':>6}")
    for s in secs:
        rate = s["units"] / s["seconds"] * 60 if s["seconds"] else float("inf")
        mark = "  ← too fast" if rate > a.max_rate else ""
        flagged |= bool(mark)
        print(f"{s['title'][:38]:<40} {s['units']:>5} {s['seconds']:>5.0f} {rate:>6.0f}{mark}")
    tot_units = sum(s["units"] for s in secs)
    labelled = sum(s["seconds"] for s in secs)
    need = tot_units / a.rate * 60
    budget = a.budget_s or labelled
    print(f"\ntotal: {tot_units} 字, labelled {labelled:.0f} s, needs ≈ {need:.0f} s at {a.rate:.0f} 字/分, "
          f"slot {budget:.0f} s")
    if need > budget * 0.9:
        flagged = True
        target = int(budget * 0.9 / 60 * a.rate)
        print(f"FLAG: leaves < 10 % slack — cut to ≈ {target} 字 or ask for more time")
    sys.exit(1 if flagged else 0)


if __name__ == "__main__":
    main()
