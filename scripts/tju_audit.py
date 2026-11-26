#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tju_audit.py — executable contract for TJU decks.

    python tju_audit.py deck.pptx [--skip 1,2] [--content] [--known PID,PX4] [--images fig*.png] [--strict-content]

--skip lists non-content pages (cover / TOC / section dividers / references / thanks); Deck.save() fills it in.

ISSUES (must be 0 before presenting) — audit_pptx():
  typography   every <a:r> and OMML <m:r>: sz >= 1800, bold, four font slots; no raw `x_k` / unparsed `_{`
  palette      explicit srgbClr only from the 5 TJU colors; no schemeClr; no shape falling back to theme fill/line
  hygiene      empty notes; no "承上启下" transition text
  overlap      no declared bbox overlap (> 0.03 in both axes); connectors and unfilled dashed containers exempt
  grid         body shapes stay inside GRID_L..GRID_R
  flush        an image directly under a header bar has the bar's exact left/width and touches it
  takeaway     the bottom red takeaway line fits on one line
WARNINGS (judgment calls; fix or justify) — audit_content():
  density, slide without visual (picture / chart / native diagram / table), estimated text overflow, table rows growing past their declared height,
  panel/card nested inside a dashed container, inconsistent takeaway size, >1 takeaway on a slide,
  colloquial / non-standard terms (with suggested replacement), unproven absolutes, first use of acronyms,
  revision notes like （去 X）, scope disclaimer on ≥ 3 slides, citations on 研究内容 pages.
IMAGES — audit_image_palette(): self-drawn figures use only TJU colors (anti-aliasing tolerated).
"""

import argparse
import itertools
import os
import re
import sys

import numpy as np
from PIL import Image

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tju_theme import (  # noqa: E402
    ALLOWED_HEX, GRID_L, GRID_R, HEX_BLUE, HEX_RED, PT_BODY, lines_needed, text_width_in,
)

EMU = 914400.0
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
_P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
_RGB = np.array([[int(h[i:i + 2], 16) for i in (0, 2, 4)] for h in sorted(ALLOWED_HEX)], dtype=float)
TOL = 0.03
BODY_Y = 1.58            # shapes starting below the header rule are body shapes

_RAW_MATH = re.compile(
    r"(?<![A-Za-z0-9_/\\.])([A-Za-z\u0370-\u03FF]_[A-Za-z0-9]+|[A-Za-z0-9\u0370-\u03FF]\^[A-Za-z0-9+\-*]+)(?![A-Za-z0-9_])"
)


# --------------------------------------------------------------------------- #
# Image palette
# --------------------------------------------------------------------------- #
def audit_image_palette(path, tol=28.0, max_off_ratio=0.005, sample=600):
    """Distance of each pixel to the nearest segment between two allowed colors (anti-aliasing safe)."""
    im = Image.open(path).convert("RGB")
    im.thumbnail((sample, sample))
    px = np.asarray(im, dtype=float).reshape(-1, 3)
    best = np.full(len(px), np.inf)
    for a, b in itertools.combinations(range(len(_RGB)), 2):
        A, AB = _RGB[a], _RGB[b] - _RGB[a]
        t = np.clip(((px - A) @ AB) / (AB @ AB), 0, 1)
        best = np.minimum(best, np.linalg.norm(px - (A + t[:, None] * AB), axis=1))
    off = best > tol
    offenders = []
    if off.any():
        cols, counts = np.unique((px[off] // 8 * 8).astype(int), axis=0, return_counts=True)
        for c, n in sorted(zip(map(tuple, cols), counts), key=lambda z: -z[1])[:5]:
            offenders.append(("#%02X%02X%02X" % c, int(n)))
    ratio = float(off.mean())
    return {"path": path, "ok": ratio <= max_off_ratio, "off_ratio": ratio, "top_offenders": offenders}


# --------------------------------------------------------------------------- #
# Shape helpers
# --------------------------------------------------------------------------- #
def _bbox(sh):
    return (sh.left / EMU, sh.top / EMU, sh.width / EMU, sh.height / EMU)


def _ancestor_tag(el, suffix):
    cur = el.getparent()
    while cur is not None:
        if isinstance(cur.tag, str) and cur.tag.endswith(suffix):
            return True
        cur = cur.getparent()
    return False


def _is_dashed_container(sh):
    xml = sh._element.xml
    return "prstDash" in xml and "<a:noFill/>" in xml.split("<a:ln")[0]


def _is_exempt_overlap(sh):
    return sh._element.tag.endswith("}cxnSp") or _is_dashed_container(sh)


def _solid_fill_hex(sh):
    spPr = sh._element.find(f"{_P}spPr")
    if spPr is None:
        return None
    sf = spPr.find(f"{_A}solidFill")
    if sf is None:
        return None
    c = sf.find(f"{_A}srgbClr")
    return (c.get("val") or "").upper() if c is not None else None


def _is_header_bar(sh):
    """Filled blue/red rectangle no taller than 0.6 in (panel / card header)."""
    if not sh._element.tag.endswith("}sp") or "prst=\"rect\"" not in sh._element.xml[:1500]:
        return False
    return _solid_fill_hex(sh) in (HEX_BLUE, HEX_RED) and sh.height / EMU <= 0.6 and sh.top / EMU >= BODY_Y


def _is_picture(sh):
    return sh._element.tag.endswith("}pic")


def _is_header_zone(sh):
    x, y, w, h = _bbox(sh)
    return y < BODY_Y or y >= 7.0


def _run_colors(sh):
    cols = set()
    for r in sh._element.iter(f"{_A}r"):
        t = r.find(f"{_A}t")
        if t is None or not (t.text or "").strip():
            continue
        c = r.find(f"{_A}rPr/{_A}solidFill/{_A}srgbClr")
        cols.add((c.get("val") or "").upper() if c is not None else None)
    return cols


def _max_run_pt(sh):
    best = 0.0
    for p in sh.text_frame.paragraphs:
        for r in p.runs:
            if r.font.size is not None:
                best = max(best, r.font.size.pt)
    return best or float(PT_BODY)


def _is_takeaway(sh):
    if not sh.has_text_frame or not sh.text_frame.text.strip():
        return False
    if sh.name == "Takeaway":
        return True
    x, y, w, h = _bbox(sh)
    paras = [p for p in sh.text_frame.paragraphs if p.text.strip()]
    return (y >= 6.15 and len(paras) == 1 and _run_colors(sh) == {HEX_RED}
            and paras[0].alignment is not None and int(paras[0].alignment) == 2)   # PP_ALIGN.CENTER


def _check_rpr(issues, idx, rpr, text, kind):
    if rpr is None:
        issues.append(f"slide {idx}: {kind} without rPr {text[:20]!r}")
        return
    sz = rpr.get("sz")
    if sz is None or int(sz) < 1800:
        issues.append(f"slide {idx}: {kind} size {sz} < 1800 in {text[:20]!r}")
    if rpr.get("b") != "1":
        issues.append(f"slide {idx}: {kind} not bold {text[:20]!r}")
    missing = [t for t in ("latin", "ea", "cs", "sym") if rpr.find(f"{_A}{t}") is None]
    if missing:
        issues.append(f"slide {idx}: {kind} missing font slots {missing} in {text[:20]!r}")


def _check_theme_fallback(issues, idx, sh):
    el = sh._element
    if not el.tag.endswith("}sp"):
        return
    nv = el.find(f"{_P}nvSpPr")
    if nv is not None:
        c = nv.find(f"{_P}cNvSpPr")
        nvpr = nv.find(f"{_P}nvPr")
        if (c is not None and c.get("txBox") == "1") or (nvpr is not None and nvpr.find(f"{_P}ph") is not None):
            return
    style, spPr = el.find(f"{_P}style"), el.find(f"{_P}spPr")
    if style is None or spPr is None:
        return
    fills = {f"{_A}{t}" for t in ("noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill")}
    fref, lref = style.find(f"{_A}fillRef"), style.find(f"{_A}lnRef")
    if not any(c.tag in fills for c in spPr) and fref is not None and fref.get("idx", "0") != "0":
        issues.append(f"slide {idx}: shape {sh.name!r} falls back to theme fill (set an explicit fill)")
    ln = spPr.find(f"{_A}ln")
    if not (ln is not None and any(c.tag in fills for c in ln)) and lref is not None and lref.get("idx", "0") != "0":
        issues.append(f"slide {idx}: shape {sh.name!r} falls back to theme line (set an explicit line)")


# --------------------------------------------------------------------------- #
# Structural audit
# --------------------------------------------------------------------------- #
def _flush_issues(idx, shapes):
    out = []
    headers = [s for s in shapes if _is_header_bar(s)]
    pics = [s for s in shapes if _is_picture(s)]
    for p in pics:
        px, py, pw, ph = _bbox(p)
        best, best_gap = None, None
        for h in headers:
            hx, hy, hw, hh = _bbox(h)
            gap = py - (hy + hh)
            ov = min(px + pw, hx + hw) - max(px, hx)
            if -TOL <= gap <= 0.8 and ov >= 0.5 * min(pw, hw):
                if best is None or gap < best_gap:
                    best, best_gap = h, gap
        if best is None:
            continue
        hx, hy, hw, hh = _bbox(best)
        between = [s for s in shapes if s is not best and s is not p and not _is_dashed_container(s)
                   and hy + hh - TOL <= _bbox(s)[1] < py - TOL
                   and min(px + pw, _bbox(s)[0] + _bbox(s)[2]) - max(px, _bbox(s)[0]) > 0.1]
        if between:
            continue
        edges_ok = abs(px - hx) <= TOL and abs(pw - hw) <= TOL
        centered = abs((px + pw / 2) - (hx + hw / 2)) <= 0.05      # letterboxed / pillarboxed under the bar
        near_full = pw >= 0.8 * hw                                   # meant to span the bar but misses it
        hdr = best.text_frame.text[:16]
        if not edges_ok and (centered or near_full):
            out.append(f"slide {idx}: image {p.name!r} not flush with header {hdr!r} (Δleft {px - hx:+.2f}, "
                       f"Δwidth {pw - hw:+.2f} in) — use add_figure_panel and draw the figure at figure_slot() size")
        elif edges_ok and best_gap > 0.10:
            out.append(f"slide {idx}: image {p.name!r} floats {best_gap:.2f} in below header {hdr!r} — "
                       f"likely letterboxed; redraw at the slot aspect ratio")
    return out


def audit_pptx(path, skip_slides=(), overlap_tol=TOL, check_overlap=True):
    from pptx import Presentation

    prs = Presentation(path)
    issues = []
    for idx, slide in enumerate(prs.slides, start=1):
        if idx in skip_slides:
            continue
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text.strip():
            issues.append(f"slide {idx}: notes not empty")
        root = slide._element
        for clr in root.iter(f"{_A}srgbClr"):
            if not _ancestor_tag(clr, "}Fallback"):
                hexv = (clr.get("val") or "").upper()
                if hexv and hexv not in ALLOWED_HEX:
                    issues.append(f"slide {idx}: color #{hexv} outside palette")
        for sclr in root.iter(f"{_A}schemeClr"):
            if not (_ancestor_tag(sclr, "}Fallback") or _ancestor_tag(sclr, "}style")):
                issues.append(f"slide {idx}: theme color schemeClr({sclr.get('val')}) instead of explicit TJU RGB")
        shapes = list(slide.shapes)
        for sh in shapes:
            _check_theme_fallback(issues, idx, sh)

        if "承上启下" in root.xml:
            issues.append(f"slide {idx}: contains '承上启下' transition text (put page logic in the subtitle)")
        for r in root.iter(f"{_A}r"):
            if _ancestor_tag(r, "}Fallback"):
                continue
            t = r.find(f"{_A}t")
            text = (t.text or "") if t is not None else ""
            if not text.strip():
                continue
            if "_{" in text or "^{" in text:
                issues.append(f"slide {idx}: unparsed sub/sup markup in {text[:30]!r}")
            elif "http" not in text and "/" not in text:
                m = _RAW_MATH.search(text)
                if m:
                    issues.append(f"slide {idx}: raw ASCII math {m.group(1)!r} in {text[:30]!r} (use _{{}}/^{{}})")
            _check_rpr(issues, idx, r.find(f"{_A}rPr"), text, "run")
        for mr in root.iter(f"{_M}r"):
            if _ancestor_tag(mr, "}Fallback"):
                continue
            mt = mr.find(f"{_M}t")
            text = (mt.text or "") if mt is not None else ""
            if text.strip():
                _check_rpr(issues, idx, mr.find(f"{_A}rPr"), text, "OMML")

        if check_overlap:
            body = [s for s in shapes if not _is_exempt_overlap(s)]
            for a, b in itertools.combinations(body, 2):
                ax, ay, aw, ah = _bbox(a)
                bx, by, bw, bh = _bbox(b)
                ox = min(ax + aw, bx + bw) - max(ax, bx)
                oy = min(ay + ah, by + bh) - max(ay, by)
                if ox > overlap_tol and oy > overlap_tol:
                    issues.append(f"slide {idx}: bbox overlap {a.name!r} × {b.name!r} ({ox:.2f}×{oy:.2f} in)")

        for sh in shapes:
            if _is_header_zone(sh) or sh._element.tag.endswith("}cxnSp"):
                continue
            x, y, w, h = _bbox(sh)
            if x < GRID_L - TOL or x + w > GRID_R + TOL:
                issues.append(f"slide {idx}: {sh.name!r} outside content grid ({x:.2f}..{x + w:.2f} in; "
                              f"grid {GRID_L}..{GRID_R})")
        issues += _flush_issues(idx, shapes)
        for sh in shapes:
            if _is_takeaway(sh):
                txt = sh.text_frame.text.strip()
                if text_width_in(txt, _max_run_pt(sh)) > sh.width / EMU - 0.10:
                    issues.append(f"slide {idx}: takeaway wraps to 2 lines {txt[:20]!r}… — shorten it (keep the size)")
    return list(dict.fromkeys(issues))


# --------------------------------------------------------------------------- #
# Content warnings
# --------------------------------------------------------------------------- #
# term → suggested replacement (empty = just drop / rephrase).  Avoid words with legitimate technical senses.
TERM_HINTS = {
    "工作一": "研究内容一", "工作二": "研究内容二", "工作三": "研究内容三",
    "卡壳": "任务中断", "搞不定": "难以求解 / 方法失效", "干活": "作业",
    "小白": "", "大白话": "", "口水话": "", "大坑": "", "偷懒": "", "瞎猜": "", "搞定": "完成",
    "痛点": "问题 / 难点", "顺手": "", "松绑": "放宽约束", "写不完": "状态空间组合爆炸", "认不出": "识别失败",
    "抓不准": "定位误差大", "草稿台": "", "算总账": "", "为什么难": "难点",
    "（最快）": "", "（最灵活）": "", "（最稳妥）": "", "！": "", "!": "",
}
# defensive / revision-history patterns (warnings; see story.md S2b)
_REVISION_NOTE = re.compile(r"[（(]\s*去\s*[A-Za-z\u4e00-\u9fff/]+[）)]")
SCOPE_WORDS = ("复用既有", "调用既有", "既有控制", "既有感知", "沿用既有")
ABSOLUTE_WORDS = ("彻底", "杜绝", "100%", "100 %", "完美", "零误差", "绝对可靠", "绝对安全", "绝对最优", "永不")
_ACRONYM = re.compile(r"(?<![A-Za-z0-9\-])([A-Z][A-Za-z0-9]*(?:-[A-Z][A-Za-z0-9]*)*)(?![A-Za-z0-9\-])")
_COMMON_ACRONYMS = {
    "IEEE", "ACM", "CVPR", "ICCV", "ECCV", "ICRA", "IROS", "RSS", "ICLR", "ICML", "AAAI", "IJCAI", "IJRR",
    "T-ASE", "T-RO", "RA-L", "TPAMI", "CoRL", "NeurIPS", "TOG", "RGB", "RGB-D", "2D", "3D", "4D", "6-DoF",
    "PDF", "PNG", "SVG", "JPG", "GIF", "PPT", "PPTX", "GPU", "CPU", "RAM", "VRAM", "ROM", "API", "JSON",
    "XML", "HTML", "URL", "DOI", "ROS", "ROS2", "TJU", "COM", "OS", "IO",
}
_WIDE = re.compile(r"[\u2E80-\u9FFF\uFF00-\uFFEF\u3000-\u303F]")


def _slide_text(slide, body_only=False):
    parts = []
    for sh in slide.shapes:
        if body_only and _is_header_zone(sh):
            continue
        if sh.has_text_frame:
            parts.append(sh.text_frame.text)
        if getattr(sh, "has_table", False) and sh.has_table:
            parts += [c.text for row in sh.table.rows for c in row.cells]
    return "\n".join(parts)


def _visible_units(text):
    return len(_WIDE.findall(text)) + len(re.findall(r"[A-Za-z0-9]+", text))


def _has_visual(slide):
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    for sh in slide.shapes:
        if sh.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.CHART) or sh._element.tag.endswith("}cxnSp"):
            return True
        if getattr(sh, "has_table", False) and sh.has_table:
            return True
        if sh.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER and "pic" in sh._element.xml[:400]:
            return True
        if sh.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and re.search(r'prst="(right|left|up|down)Arrow"', sh._element.xml[:800]):
            return True
    return False


def _est_text_height(sh):
    """Conservative wrapped-text height (in): CJK 1 em, Latin ≈ 0.5 em, line ≈ 1.3 × size."""
    from pptx.oxml.ns import qn
    tf = sh.text_frame
    bp = tf._txBody.find(qn("a:bodyPr"))

    def ins(k, d):
        v = bp.get(k) if bp is not None else None
        return int(v) / EMU if v is not None else d

    inner_w = sh.width / EMU - ins("lIns", 0.1) - ins("rIns", 0.1)
    if inner_w <= 0.2:
        return 0.0, 0.0
    total, counts = ins("tIns", 0.05) + ins("bIns", 0.05), []
    for p in tf.paragraphs:
        txt = "".join(r.text for r in p.runs)
        sz = max([r.font.size.pt for r in p.runs if r.font.size is not None] or [18.0])
        n = lines_needed(txt, inner_w, sz, pad_in=0.0) if txt else 1
        counts.append(n)
        ls = p.line_spacing if isinstance(p.line_spacing, float) else 1.0
        total += n * sz * 1.3 * ls / 72.0
        for sp in (p.space_before, p.space_after):
            if sp is not None:
                total += sp.pt / 72.0
    if counts == [1]:
        return 0.0, sh.height / EMU
    return total, sh.height / EMU


def _table_growth(sh):
    tbl = sh.table
    widths = [c.width / EMU for c in tbl.columns]
    need = 0.0
    for row in tbl.rows:
        rn = 1
        for c, w in zip(row.cells, widths):
            sz = 18.0
            for p in c.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.size is not None:
                        sz = max(sz, r.font.size.pt)
            rn = max(rn, sum(lines_needed(p.text, w - 0.16, sz, pad_in=0.04) for p in c.text_frame.paragraphs))
        need += max(row.height / EMU, rn * 18 * 1.25 / 72.0 + 0.12)
    return need - sh.height / EMU


def _defined_in_place(text):
    """Acronyms introduced with a definition, e.g. 强化学习（RL） / Model Predictive Control (MPC)."""
    return set(re.findall(r"[\u4e00-\u9fffA-Za-z\s\-]{2,}[（(]\s*([A-Z][A-Za-z0-9\-]+)\s*[）)]", text))


def _extract_acronyms(text):
    out = set()
    for tok in _ACRONYM.findall(text):
        if tok in _COMMON_ACRONYMS or sum(1 for c in tok if c.isupper()) < 2 or re.fullmatch(r"[A-Z]\d+", tok):
            continue
        out.add(tok)
    return out


def audit_content(path, skip_slides=(), max_units=160, known=()):
    from pptx import Presentation

    prs = Presentation(path)
    warns, seen_acr, takeaway_sizes, scope_pages = [], set(known), {}, []
    for idx, slide in enumerate(prs.slides, start=1):
        if idx in skip_slides:
            continue
        text = _slide_text(slide)
        shapes = list(slide.shapes)
        units = _visible_units(_slide_text(slide, body_only=True))
        if units > max_units:
            warns.append(f"slide {idx}: dense body text ({units} visible units > {max_units}, banner/subtitle excluded); "
                         f"cut, split, or move it to the speech")
        if not _has_visual(slide):
            warns.append(f"slide {idx}: no visual (photo, paper figure, native diagram, schematic, or plot)")
        for sh in shapes:
            if sh.has_text_frame and sh.text_frame.text.strip():
                est, box = _est_text_height(sh)
                if box > 0 and est > box * 1.05 + 0.05:
                    warns.append(f"slide {idx}: possible text overflow in {sh.name!r} (est {est:.2f} > box {box:.2f} in): "
                                 f"{sh.text_frame.text[:24]!r}")
            if getattr(sh, "has_table", False) and sh.has_table:
                g = _table_growth(sh)
                if g > 0.10:
                    warns.append(f"slide {idx}: table rows will grow ≈ {g:.2f} in past the declared height — "
                                 f"check it does not cover content below (or use add_table)")
        for d in (s for s in shapes if _is_dashed_container(s)):
            dx, dy, dw, dh = _bbox(d)
            for s in shapes:
                if s is d or not _is_header_bar(s):
                    continue
                sx, sy, sw, sh_ = _bbox(s)
                if sx >= dx - 0.02 and sy >= dy - 0.02 and sx + sw <= dx + dw + 0.02 and sy + sh_ <= dy + dh + 0.02:
                    warns.append(f"slide {idx}: header-bar panel/card inside a dashed container — keep one frame level "
                                 f"(drop the dashed box)")
                    break
        tks = [s for s in shapes if _is_takeaway(s)]
        if len(tks) > 1:
            warns.append(f"slide {idx}: {len(tks)} takeaway lines; keep one")
        for s in tks:
            takeaway_sizes.setdefault(_max_run_pt(s), []).append(idx)
        for term, repl in TERM_HINTS.items():
            if term in text:
                warns.append(f"slide {idx}: non-academic term {term!r}" + (f" → use {repl!r}" if repl else " — rephrase"))
        m = _REVISION_NOTE.search(text)
        if m:
            warns.append(f"slide {idx}: revision note {m.group(0)!r} on the slide — the committee never saw the old "
                         f"version; move it to the speech")
        if re.search(r"研究内容[一二三四五]", text[:200]) and re.search(r"\[\d+\]", text):
            warns.append(f"slide {idx}: citation on a research-content page — keep citations on related-work pages")
        if any(w in text for w in SCOPE_WORDS):
            scope_pages.append(idx)
        for w in ABSOLUTE_WORDS:
            if w in text:
                warns.append(f"slide {idx}: absolute claim {w!r} — prove it or soften")
        seen_acr.update(_defined_in_place(text))
        new_acr = sorted(a for a in _extract_acronyms(text) if a not in seen_acr)
        seen_acr.update(new_acr)
        if new_acr:
            warns.append(f"slide {idx}: first use of {', '.join(new_acr)} — define before use")
    if len(scope_pages) >= 3:
        warns.append(f"deck: reuse/scope disclaimer repeated on slides {scope_pages} — state it once on the overview")
    if len(takeaway_sizes) > 1:
        detail = "; ".join(f"{pt:g} pt on {pages}" for pt, pages in sorted(takeaway_sizes.items()))
        warns.append(f"deck: takeaway lines use different sizes ({detail}) — use one size everywhere")
    return warns


# --------------------------------------------------------------------------- #
# Entry points
# --------------------------------------------------------------------------- #
def run_audit(path, skip=(), content=True, images=(), max_units=160, known=(), verbose=True, check_overlap=True):
    """Structural issues (returned) + content warnings and image checks (printed)."""
    issues = audit_pptx(path, skip_slides=set(skip), check_overlap=check_overlap)
    warns = audit_content(path, skip_slides=set(skip), max_units=max_units, known=known) if content else []
    img_bad = []
    for p in images:
        r = audit_image_palette(p)
        if not r["ok"]:
            img_bad.append(f"image {p}: off-palette {r['off_ratio']:.4f} {r['top_offenders']}")
    if verbose:
        print(f"[tju-audit] {path}: {len(issues)} issue(s), {len(warns)} warning(s)")
        for i in issues:
            print("  -", i)
        for w in warns:
            print("  ~", w)
        for b in img_bad:
            print("  -", b)
    return issues + img_bad


def main():
    ap = argparse.ArgumentParser(description="TJU deck audit")
    ap.add_argument("pptx", nargs="?")
    ap.add_argument("--skip", default="", help="pages to skip, e.g. 1,2 (cover/TOC/references/thanks)")
    ap.add_argument("--images", nargs="*", default=[])
    ap.add_argument("--no-overlap", action="store_true")
    ap.add_argument("--content", action="store_true", help="also print content warnings")
    ap.add_argument("--max-units", type=int, default=160, help="body text budget per slide (banner/subtitle excluded)")
    ap.add_argument("--known", default="", help="acronyms already defined/explained, e.g. PID,PX4")
    ap.add_argument("--strict-content", action="store_true", help="content warnings count as failures")
    a = ap.parse_args()
    skip = {int(x) for x in a.skip.split(",") if x.strip()}
    known = {x.strip() for x in a.known.split(",") if x.strip()}
    bad = 0
    if a.pptx:
        issues = run_audit(a.pptx, skip, content=a.content or a.strict_content, images=a.images,
                           max_units=a.max_units, known=known, check_overlap=not a.no_overlap)
        bad += len(issues)
        if a.strict_content:
            bad += len(audit_content(a.pptx, skip, a.max_units, known))
    else:
        for p in a.images:
            r = audit_image_palette(p)
            print(f"[image] {p}: {'OK' if r['ok'] else 'OFF-PALETTE'} off_ratio={r['off_ratio']:.4f} {r['top_offenders']}")
            bad += 0 if r["ok"] else 1
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
