#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TJU-PPT 自动审计 (tju_audit.py)

1. audit_image_palette(png)：检查自绘图（SVG 引擎 / Matplotlib 输出）是否只用 TJU 四色。
   抗锯齿会产生两色之间的过渡色，因此判定标准为“像素到任意两种允许色连线段的 RGB 距离”，
   而不是到 5 个色点的距离；照片 / 论文截图不适用。
2. audit_pptx(pptx)：OpenXML 级审计（SKILL.md §4.2 第 2 步的可执行版本）
   - 文字 run：sz >= 1800、b="1"、四槽位字体 (latin/ea/cs/sym)
   - 颜色：所有 a:srgbClr 只能取 ALLOWED_HEX
   - 残留标记：未解析的 _{ / ^{ 、"承上启下"、备注页文字
   - 包围盒重叠：ox > 0.03 in 且 oy > 0.03 in；直线连接器与无填充虚线容器不计入

命令行：
    python tju_audit.py deck.pptx [--skip 1,2] [--images fig1.png fig2.png]
"""

import argparse
import itertools
import re
import sys

import numpy as np
from PIL import Image

ALLOWED_HEX = {"00337F", "EAF1FB", "C00000", "000000", "FFFFFF"}
_RGB = np.array([[int(h[i:i + 2], 16) for i in (0, 2, 4)] for h in sorted(ALLOWED_HEX)], dtype=float)
EMU = 914400.0


# --------------------------------------------------------------------------- #
# 图片调色板
# --------------------------------------------------------------------------- #
def audit_image_palette(path, tol=28.0, max_off_ratio=0.005, sample=600):
    """
    返回 dict(ok, off_ratio, top_offenders)。tol 为 RGB 欧氏距离容差；
    max_off_ratio 为允许的越色像素比例（默认 0.5%，容纳字体 hinting 杂色）。
    """
    im = Image.open(path).convert("RGB")
    im.thumbnail((sample, sample))
    px = np.asarray(im, dtype=float).reshape(-1, 3)
    best = np.full(len(px), np.inf)
    for a, b in itertools.combinations(range(len(_RGB)), 2):
        A, B = _RGB[a], _RGB[b]
        AB = B - A
        t = np.clip(((px - A) @ AB) / (AB @ AB), 0, 1)
        d = np.linalg.norm(px - (A + t[:, None] * AB), axis=1)
        best = np.minimum(best, d)
    off = best > tol
    ratio = float(off.mean())
    offenders = []
    if off.any():
        cols, counts = np.unique((px[off] // 8 * 8).astype(int), axis=0, return_counts=True)
        for c, n in sorted(zip(map(tuple, cols), counts), key=lambda z: -z[1])[:5]:
            offenders.append(("#%02X%02X%02X" % c, int(n)))
    return {"path": path, "ok": ratio <= max_off_ratio, "off_ratio": ratio, "top_offenders": offenders}


# --------------------------------------------------------------------------- #
# PPTX
# --------------------------------------------------------------------------- #
def _bbox(sh):
    return (sh.left / EMU, sh.top / EMU, sh.width / EMU, sh.height / EMU)


def _is_exempt(sh):
    tag = sh._element.tag
    if tag.endswith("}cxnSp"):
        return True
    xml = sh._element.xml
    if "prstDash" in xml and "<a:noFill/>" in xml.split("<a:ln")[0]:
        return True  # 无填充虚线容器
    return False


def audit_pptx(path, skip_slides=(), overlap_tol=0.03, check_overlap=True):
    from pptx import Presentation
    from pptx.oxml.ns import qn

    prs = Presentation(path)
    issues = []
    for idx, slide in enumerate(prs.slides, start=1):
        if idx in skip_slides:
            continue
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
            issues.append(f"slide {idx}: notes not empty")
        xml = slide._element.xml
        for hexv in set(re.findall(r'srgbClr val="([0-9A-Fa-f]{6})"', xml)):
            if hexv.upper() not in ALLOWED_HEX:
                issues.append(f"slide {idx}: color #{hexv} outside palette")
        if "承上启下" in xml:
            issues.append(f"slide {idx}: contains '承上启下'")
        for r in slide._element.iter(qn("a:r")):
            t = r.find(qn("a:t"))
            text = (t.text or "") if t is not None else ""
            if not text.strip():
                continue
            if "_{" in text or "^{" in text:
                issues.append(f"slide {idx}: unparsed sub/sup markup in {text[:30]!r}")
            rpr = r.find(qn("a:rPr"))
            if rpr is None:
                issues.append(f"slide {idx}: run without rPr {text[:20]!r}")
                continue
            sz = rpr.get("sz")
            if sz is None or int(sz) < 1800:
                issues.append(f"slide {idx}: font size {sz} < 1800 in {text[:20]!r}")
            if rpr.get("b") != "1":
                issues.append(f"slide {idx}: not bold {text[:20]!r}")
            slots = {tag: rpr.find(qn(f"a:{tag}")) for tag in ("latin", "ea", "cs", "sym")}
            missing = [k for k, v in slots.items() if v is None]
            if missing:
                issues.append(f"slide {idx}: missing font slots {missing} in {text[:20]!r}")
        if check_overlap:
            shapes = [s for s in slide.shapes if not _is_exempt(s)]
            for a, b in itertools.combinations(shapes, 2):
                ax, ay, aw, ah = _bbox(a)
                bx, by, bw, bh = _bbox(b)
                ox = min(ax + aw, bx + bw) - max(ax, bx)
                oy = min(ay + ah, by + bh) - max(ay, by)
                if ox > overlap_tol and oy > overlap_tol:
                    issues.append(f"slide {idx}: bbox overlap {a.name!r} × {b.name!r} ({ox:.2f}×{oy:.2f} in)")
    return list(dict.fromkeys(issues))


# --------------------------------------------------------------------------- #
# 内容级告警（narrative_principles.md 的可自动化部分；均为“需人工判断”的 warning）
# --------------------------------------------------------------------------- #
# 口水话 / 过度小白化表述（narrative_principles.md P1b）：只告警，按语境人工判断
COLLOQUIAL_WORDS = ("科普", "草稿台", "算总账", "总账", "看图", "手眼微调", "抓不准", "认不出", "写不完", "顺手",
                    "松绑", "上车", "挡死", "盲抓", "大坑", "偷懒", "瞎", "搞定", "痛点", "为什么难", "连锁反应",
                    "（最快）", "（最灵活）", "（最稳妥）", "！", "!", "白话", "小白")
ABSOLUTE_WORDS = ("全局最优", "彻底", "杜绝", "100%", "100 %", "绝对", "完美", "首次", "零误差")
_ACRONYM = re.compile(r"(?<![A-Za-z])([A-Z][A-Za-z0-9]*[A-Z][A-Za-z0-9\-]*)(?![A-Za-z])")
_WIDE = re.compile(r"[\u2E80-\u9FFF\uFF00-\uFFEF\u3000-\u303F]")


def _slide_text(slide):
    parts = []
    for sh in slide.shapes:
        if sh.has_text_frame:
            parts.append(sh.text_frame.text)
        if getattr(sh, "has_table", False) and sh.has_table:
            for row in sh.table.rows:
                for c in row.cells:
                    parts.append(c.text)
    return "\n".join(parts)


def _visible_units(text):
    """CJK 字符 + 拉丁词/数字串 的个数，近似“观众需要读的量”。"""
    return len(_WIDE.findall(text)) + len(re.findall(r"[A-Za-z0-9]+", text))


def _est_text_height(sh):
    """保守估算自动换行后的文字高度（英寸）：CJK=1em、ASCII≈0.5em、中西文交界 +0.25em 自动间距、
    行高≈1.3×字号（微软雅黑 ascent+descent 较大）。LibreOffice/PowerPoint 的实际换行只会更宽，宁可误报。"""
    from pptx.oxml.ns import qn
    tf = sh.text_frame
    bp = tf._txBody.find(qn("a:bodyPr"))
    def ins(k, d):
        v = bp.get(k) if bp is not None else None
        return int(v) / EMU if v is not None else d
    inner_w = sh.width / EMU - ins("lIns", 0.1) - ins("rIns", 0.1)
    if inner_w <= 0.2:
        return 0.0, 0.0
    total = ins("tIns", 0.05) + ins("bIns", 0.05)
    line_counts = []
    for p in tf.paragraphs:
        width_pt, max_sz, prev_wide = 0.0, 0.0, None
        for r in p.runs:
            sz = (r.font.size.pt if r.font.size is not None else 18.0)
            max_sz = max(max_sz, sz)
            for ch in r.text:
                wide = bool(_WIDE.match(ch))
                width_pt += sz if wide else 0.5 * sz
                if prev_wide is not None and ch.strip() and wide != prev_wide:
                    width_pt += 0.25 * sz  # 中西文自动间距
                if ch.strip():
                    prev_wide = wide
        if max_sz == 0:
            max_sz = 18.0
        lines = max(1, int(-(-width_pt / 72.0 // inner_w)))  # ceil
        line_counts.append(lines)
        ls = p.line_spacing if isinstance(p.line_spacing, float) else 1.0
        total += lines * max_sz * 1.3 * ls / 72.0
        for sp in (p.space_before, p.space_after):
            if sp is not None:
                total += sp.pt / 72.0
    if len(line_counts) == 1 and line_counts[0] == 1:
        return 0.0, sh.height / EMU  # 单行文字纵向略超框不影响观感，只关心“换行后压到下方元素”
    return total, sh.height / EMU


def audit_content(path, skip_slides=(), max_units=200):
    """返回 warning 列表：文字密度、无图内容页、估计溢出、绝对化表述、缩写首次出现。"""
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    prs = Presentation(path)
    warns, seen_acr = [], set()
    for idx, slide in enumerate(prs.slides, start=1):
        text = _slide_text(slide)
        new_acr = sorted({a for a in _ACRONYM.findall(text) if a not in seen_acr})
        seen_acr.update(new_acr)
        if idx in skip_slides:
            continue
        units = _visible_units(text)
        if units > max_units:
            warns.append(f"slide {idx}: dense text ({units} visible units > {max_units}); cut, split, or move detail to the report")
        has_pic = any(sh.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.CHART) or
                      (sh.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER and "pic" in sh._element.xml[:400])
                      for sh in slide.shapes)
        if not has_pic:
            warns.append(f"slide {idx}: no picture/chart on content slide (photo, paper figure, schematic, or plot)")
        for sh in slide.shapes:
            if not sh.has_text_frame or not sh.text_frame.text.strip():
                continue
            est, box = _est_text_height(sh)
            if box > 0 and est > box * 1.05 + 0.05:
                warns.append(f"slide {idx}: possible text overflow in {sh.name!r} (est {est:.2f} in > box {box:.2f} in): {sh.text_frame.text[:24]!r}")
        hits = [w for w in COLLOQUIAL_WORDS if w in text]
        if hits:
            warns.append(f"slide {idx}: colloquial wording {hits} — define the proper term instead (P1b)")
        for w in ABSOLUTE_WORDS:
            if w in text:
                warns.append(f"slide {idx}: absolute claim '{w}' — prove it or soften")
        if new_acr:
            warns.append(f"slide {idx}: first use of {', '.join(new_acr)} — define before use")
    return warns


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx", nargs="?")
    ap.add_argument("--skip", default="", help="跳过的页码，如 1,2（封面/目录）")
    ap.add_argument("--images", nargs="*", default=[])
    ap.add_argument("--no-overlap", action="store_true")
    ap.add_argument("--content", action="store_true", help="内容级告警：密度/无图/溢出估计/绝对化/缩写首现")
    ap.add_argument("--max-units", type=int, default=200, help="单页可见字数上限（CJK 字 + 拉丁词）")
    ap.add_argument("--strict-content", action="store_true", help="内容告警计入失败退出码")
    a = ap.parse_args()
    bad = 0
    if a.pptx:
        skip = {int(x) for x in a.skip.split(",") if x.strip()}
        iss = audit_pptx(a.pptx, skip_slides=skip, check_overlap=not a.no_overlap)
        print(f"[pptx] {a.pptx}: {len(iss)} issue(s)")
        for i in iss:
            print("  -", i)
        bad += len(iss)
        if a.content or a.strict_content:
            ws = audit_content(a.pptx, skip_slides=skip, max_units=a.max_units)
            print(f"[content] {len(ws)} warning(s)")
            for w in ws:
                print("  ~", w)
            if a.strict_content:
                bad += len(ws)
    for p in a.images:
        r = audit_image_palette(p)
        print(f"[image] {p}: {'OK' if r['ok'] else 'OFF-PALETTE'} off_ratio={r['off_ratio']:.4f} {r['top_offenders']}")
        bad += 0 if r["ok"] else 1
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
