#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TJU-PPT 复杂科研示意图引擎 (tju_svg_diagram.py)

用于原生形状难以表达的技术路线图 / 多分支算法架构 / 系统框图：
结构化描述（Python API 或 JSON spec） → 校验（溢出 / 重叠 / 连线穿框 / 出界） → SVG → 2x PNG。

关键约定：
1. **按最终贴图尺寸作图**：画布单位为英寸，与 PPT 中放置框尺寸一致，因此图中 18 pt 文字
   贴进 PPT 后视觉上仍是 18 pt，与正文字号统一。
2. **中英双字体逐段指定**：cairosvg 只取 font-family 列表的第一个字体且不做字形回退，
   因此本引擎把每行文字按字符拆成 CJK 段（微软雅黑 / 本机等价黑体）与西文段（Times New Roman）。
3. **只用 TJU 四色**：节点角色 input / core / output / emphasis / plain / note，连线 blue / red / black。
4. **硬校验**：save() 默认 strict=True，存在溢出、节点重叠、连线穿过无关节点时直接抛错，
   逼迫修正布局而不是把有瑕疵的图塞进 PPT。

依赖：cairosvg（渲染）、Pillow（按真实字体文件测量字宽，可选；缺失时退回启发式估计）。
"""

import html
import json
import math
import os
import re
import shutil
import subprocess
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from tju_theme import HEX_BLACK, HEX_BLUE, HEX_BLUE_LIGHT, HEX_RED, HEX_WHITE  # noqa: E402

BLUE, BLUE_LIGHT, RED, BLACK, WHITE = (f"#{h}" for h in (HEX_BLUE, HEX_BLUE_LIGHT, HEX_RED, HEX_BLACK, HEX_WHITE))
LINE_COLORS = {"blue": BLUE, "red": RED, "black": BLACK}

# 角色 → (填充, 边框, 文字, 边框宽 px@96ppi)
ROLES = {
    "input": (BLUE_LIGHT, BLUE, BLACK, 2.0),
    "core": (BLUE, BLUE, WHITE, 2.0),
    "output": (WHITE, RED, BLACK, 3.0),
    "emphasis": (RED, RED, WHITE, 2.0),
    "plain": (WHITE, BLUE, BLACK, 2.0),
    "note": (WHITE, None, BLACK, 0.0),
}

PPI = 96.0
_MARKUP = re.compile(r"(_\{[^{}]*\}|\^\{[^{}]*\})")
_COMBINING = re.compile(r"[\u0300-\u036f\u20d0-\u20ff]")


# --------------------------------------------------------------------------- #
# 字体解析与测量
# --------------------------------------------------------------------------- #
def _is_cjk(ch):
    """CJK 及 Times New Roman 缺字的全角符号（①②③、■□▲、㈠ 等）统一交给中文字体渲染。"""
    o = ord(ch)
    return ((0x2E80 <= o <= 0x9FFF) or (0xF900 <= o <= 0xFAFF) or (0xFF00 <= o <= 0xFFEF)
            or (0x3000 <= o <= 0x303F) or (0x2460 <= o <= 0x24FF) or (0x2500 <= o <= 0x25FF)
            or (0x2600 <= o <= 0x27BF) or (0x3200 <= o <= 0x33FF) or (0xFE30 <= o <= 0xFE4F))


def _fc_match(pattern, fmt):
    if not shutil.which("fc-match"):
        return None
    try:
        return subprocess.run(["fc-match", "-f", fmt, pattern], capture_output=True, text=True, timeout=5).stdout.strip()
    except Exception:
        return None


def _resolve_fonts():
    """返回 (latin_family, latin_file, cjk_family, cjk_file)。优先 Times New Roman + 微软雅黑。"""
    latin_family, latin_file = "Times New Roman", None
    cjk_family, cjk_file = "Microsoft YaHei", None
    if sys.platform.startswith("win"):
        fdir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
        for f in ("timesbd.ttf", "times.ttf"):
            if os.path.exists(os.path.join(fdir, f)):
                latin_file = os.path.join(fdir, f); break
        for f in ("msyhbd.ttc", "msyh.ttc"):
            if os.path.exists(os.path.join(fdir, f)):
                cjk_file = os.path.join(fdir, f); break
        return latin_family, latin_file, cjk_family, cjk_file
    fam = _fc_match("Times New Roman:bold", "%{family[0]}")
    if fam and "times" in fam.lower():
        latin_file = _fc_match("Times New Roman:bold", "%{file}")
    for cand in ("Microsoft YaHei", "Noto Sans CJK SC", "Source Han Sans SC", "WenQuanYi Micro Hei"):
        fam = _fc_match(f"{cand}:bold", "%{family[0]}") or ""
        allfam = _fc_match(f"{cand}:bold", "%{family}") or ""
        if cand.lower() in allfam.lower() or cand.lower() == fam.lower():
            cjk_family = cand
            cjk_file = _fc_match(f"{cand}:bold", "%{file}")
            break
    return latin_family, latin_file, cjk_family, cjk_file


LATIN_FAMILY, LATIN_FILE, CJK_FAMILY, CJK_FILE = _resolve_fonts()
_FONT_CACHE = {}


def _pil_font(path, px):
    if not path:
        return None
    key = (path, int(round(px * 4)))
    if key not in _FONT_CACHE:
        try:
            from PIL import ImageFont
            _FONT_CACHE[key] = ImageFont.truetype(path, max(1, int(round(px))))
        except Exception:
            _FONT_CACHE[key] = None
    return _FONT_CACHE[key]


def _measure_run(text, px, cjk):
    font = _pil_font(CJK_FILE if cjk else LATIN_FILE, px)
    if font is not None:
        try:
            return font.getlength(text)
        except AttributeError:
            return font.getsize(text)[0]
    em = 0
    for ch in text:
        em += 1.0 if cjk else (0.25 if ch == " " else 0.70 if ch.isupper() else 0.52 if ch.isalnum() else 0.38)
    return em * px


_CMAPS = None


def _cmap_of(path):
    """Unicode code points covered by a font file (first face of a .ttc); None if unknown."""
    if not path:
        return None
    try:
        from fontTools.ttLib import TTCollection, TTFont
        font = TTCollection(path).fonts[0] if path.lower().endswith(".ttc") else TTFont(path, lazy=True)
        return set(font.getBestCmap() or {})
    except Exception:
        return None


def _cmaps():
    global _CMAPS
    if _CMAPS is None:
        _CMAPS = (_cmap_of(LATIN_FILE), _cmap_of(CJK_FILE))
    return _CMAPS


def _route_cjk(ch):
    """True → render with the CJK font.  Symbols missing from Times New Roman (e.g. ‖ U+2016) are routed to
    the CJK font when it has the glyph, so they never become tofu boxes."""
    if _is_cjk(ch):
        return True
    if ord(ch) < 0x80:
        return False
    lat, cjk = _cmaps()
    return lat is not None and ord(ch) not in lat and cjk is not None and ord(ch) in cjk


def missing_glyphs(text):
    """Characters that neither resolved font can draw (would render as tofu)."""
    lat, cjk = _cmaps()
    out = []
    for ch in _MARKUP.sub(lambda m: m.group(0)[2:-1], str(text)):
        if not ch.strip():
            continue
        table = cjk if _route_cjk(ch) else lat
        if table is not None and ord(ch) not in table:
            out.append(ch)
    return sorted(set(out))


def _split_script(text):
    """按 CJK / 西文切分成 [(text, is_cjk), ...]（按字形可用性路由字体）"""
    runs = []
    for ch in text:
        c = _route_cjk(ch)
        if runs and runs[-1][1] == c:
            runs[-1][0] += ch
        else:
            runs.append([ch, c])
    return [(t, c) for t, c in runs]


def _tokens(line):
    """把一行拆成 [(text, kind)]，kind ∈ normal/sub/sup"""
    out = []
    for tok in _MARKUP.split(line):
        if not tok:
            continue
        if tok.startswith("_{"):
            out.append((tok[2:-1], "sub"))
        elif tok.startswith("^{"):
            out.append((tok[2:-1], "sup"))
        else:
            out.append((tok, "normal"))
    return out


SUBSUP_SCALE = 0.72


def line_width_px(line, px):
    w = 0.0
    for txt, kind in _tokens(line):
        size = px * (SUBSUP_SCALE if kind != "normal" else 1.0)
        for seg, cjk in _split_script(txt):
            w += _measure_run(seg, size, cjk)
    return w


def wrap_line(line, px, max_w):
    """
    按真实字宽自动折行：CJK 逐字可断，西文按空格断；不拆开 _{}/^{} 标记。
    折成多行时自动均衡行宽，避免“账”“偏。”这类单字孤行。
    """
    if line_width_px(line, px) <= max_w:
        return [line]
    lines = _greedy_wrap(line, px, max_w)
    n = len(lines)
    if n > 1:
        w = max_w
        while w > 10:
            trial = _greedy_wrap(line, px, w * 0.97)
            if len(trial) != n:
                break
            w *= 0.97
            lines = trial
    return lines


def _greedy_wrap(line, px, max_w):
    units = []
    for tok in _MARKUP.split(line):
        if not tok:
            continue
        if _MARKUP.fullmatch(tok):
            if units and not units[-1].isspace():
                units[-1] += tok  # 上下标与其主体字符不可拆行
            else:
                units.append(tok)
        else:
            for piece in re.findall(r"[A-Za-z0-9.\-+/%()=:,]+|\s+|.", tok):
                units.append(piece)
    lines, cur = [], ""
    for u in units:
        trial = cur + u
        if cur and line_width_px(trial, px) > max_w:
            lines.append(cur.rstrip())
            cur = u.lstrip()
        else:
            cur = trial
    if cur.strip():
        lines.append(cur.rstrip())
    # 避免行首孤立标点：把行首的中文标点挪回上一行
    for i in range(1, len(lines)):
        while lines[i] and lines[i][0] in "，。；：、！？）」』,.;:!?)":
            lines[i - 1] += lines[i][0]
            lines[i] = lines[i][1:]
    return [l for l in lines if l]


# --------------------------------------------------------------------------- #
# 几何工具
# --------------------------------------------------------------------------- #
def _seg_hits_rect(p, q, r, pad=1.0):
    """正交/任意线段与矩形 (x, y, w, h) 内部是否相交（Liang–Barsky）。"""
    x0, y0, x1, y1 = r[0] + pad, r[1] + pad, r[0] + r[2] - pad, r[1] + r[3] - pad
    dx, dy = q[0] - p[0], q[1] - p[1]
    t0, t1 = 0.0, 1.0
    for pp, qq in ((-dx, p[0] - x0), (dx, x1 - p[0]), (-dy, p[1] - y0), (dy, y1 - p[1])):
        if pp == 0:
            if qq < 0:
                return False
        else:
            t = qq / pp
            if pp < 0:
                t0 = max(t0, t)
            else:
                t1 = min(t1, t)
            if t0 > t1:
                return False
    return True


def _rect_overlap(a, b, tol=0.5):
    return (min(a[0] + a[2], b[0] + b[2]) - max(a[0], b[0]) > tol and
            min(a[1] + a[3], b[1] + b[3]) - max(a[1], b[1]) > tol)


# --------------------------------------------------------------------------- #
# 主类
# --------------------------------------------------------------------------- #
class TjuDiagram:
    """
    用法：
        d = TjuDiagram(5.6, 3.2)                       # 与 PPT 放置框同尺寸（英寸）
        d.node("in", 0.1, 0.2, 1.5, 0.7, "点云输入", role="input")
        d.node("core", 2.0, 0.2, 1.6, 0.7, "位姿估计 T_{k}", role="core")
        d.edge("in", "core", label="ICP")
        d.group(0.05, 0.05, 3.7, 1.1, "感知模块")
        d.save("fig.png", svg_path="fig.svg")         # strict 校验通过才写出
    """

    def __init__(self, width_in, height_in, base_pt=18, line_spacing=1.18):
        self.W, self.H = width_in * PPI, height_in * PPI
        self.width_in, self.height_in = width_in, height_in
        self.base_px = base_pt * PPI / 72.0
        self.line_spacing = line_spacing
        self.nodes, self.groups, self.edges, self.texts, self.arrows = {}, [], [], [], []
        self._order = []

    # ---------------- 元素 ----------------
    def node(self, nid, x, y, w, h, text, role="plain", pt=None, align="center", wrap=True, dashed=False,
             check=True):
        """check=False：示意性图元（如机械臂示意、偏移轮廓）允许与其他节点重叠，不参与重叠/穿框校验。"""
        if nid in self.nodes:
            raise ValueError(f"duplicate node id {nid!r}")
        if role not in ROLES:
            raise ValueError(f"unknown role {role!r}; use {list(ROLES)}")
        self.nodes[nid] = dict(id=nid, x=x * PPI, y=y * PPI, w=w * PPI, h=h * PPI, text=text,
                               role=role, px=(pt * PPI / 72.0) if pt else self.base_px,
                               align=align, wrap=wrap, dashed=dashed, check=check)
        self._order.append(nid)
        return nid

    def group(self, x, y, w, h, label=None, color="blue", dashed=True, pt=None):
        """虚线分组域（画在节点之后层级的下方）；label 写在左上角内侧。"""
        self.groups.append(dict(x=x * PPI, y=y * PPI, w=w * PPI, h=h * PPI, label=label,
                                color=LINE_COLORS[color], dashed=dashed,
                                px=(pt * PPI / 72.0) if pt else self.base_px))

    def text(self, x, y, content, pt=None, color="black", anchor="start"):
        """自由文字（图注、坐标说明）。x, y 为基线位置（英寸）。"""
        self.texts.append(dict(x=x * PPI, y=y * PPI, content=content,
                               px=(pt * PPI / 72.0) if pt else self.base_px,
                               color=LINE_COLORS.get(color, color), anchor=anchor))

    def arrow(self, points, color="blue", dashed=False, width=2.4, both=False):
        """自由折线箭头（不绑定节点），points: [(x, y), ...] 英寸。用于分区之间的示意箭头。"""
        self.arrows.append(dict(pts=[(a * PPI, b * PPI) for a, b in points], color=LINE_COLORS[color],
                                dashed=dashed, width=width, both=both))

    def edge(self, src, dst, label=None, src_side=None, dst_side=None, route="auto", via=None,
             color="blue", dashed=False, both=False, width=2.4, label_pos=0.5, label_offset=None):
        """
        连线。route: auto / straight / hv / vh / hvh / vhv；via: 额外拐点 [(x, y) 英寸]。
        label_pos: 标签放在最长线段上的相对位置；label_offset: (dx, dy) 英寸微调。
        """
        self.edges.append(dict(src=src, dst=dst, label=label, src_side=src_side, dst_side=dst_side,
                               route=route, via=[(a * PPI, b * PPI) for a, b in (via or [])],
                               color=LINE_COLORS[color], dashed=dashed, both=both, width=width,
                               label_pos=label_pos,
                               label_offset=((label_offset[0] * PPI, label_offset[1] * PPI) if label_offset else (0, 0))))

    # ---------------- 布局助手 ----------------
    def row(self, items, x, y, w, h, gap=0.35, roles=None, chain=True, **edge_kw):
        """横向等宽排一行节点并（可选）首尾相连。items: [(id, text), ...]"""
        n = len(items)
        bw = (w - gap * (n - 1)) / n
        roles = roles or ["plain"] * n
        for i, ((nid, txt), rl) in enumerate(zip(items, roles)):
            self.node(nid, x + i * (bw + gap), y, bw, h, txt, role=rl)
        if chain:
            for (a, _), (b, _) in zip(items[:-1], items[1:]):
                self.edge(a, b, **edge_kw)
        return [i[0] for i in items]

    def column(self, items, x, y, w, h, gap=0.30, roles=None, chain=True, **edge_kw):
        n = len(items)
        bh = (h - gap * (n - 1)) / n
        roles = roles or ["plain"] * n
        for i, ((nid, txt), rl) in enumerate(zip(items, roles)):
            self.node(nid, x, y + i * (bh + gap), w, bh, txt, role=rl)
        if chain:
            for (a, _), (b, _) in zip(items[:-1], items[1:]):
                self.edge(a, b, **edge_kw)
        return [i[0] for i in items]

    @classmethod
    def layered(cls, width_in, height_in, ranks, direction="LR", margin=0.12, gap_main=0.55,
                gap_cross=0.22, node_size=None, edges=(), **kw):
        """
        archify 式自动分层布局。ranks: [[(id, text, role), ...], ...]，每个 rank 一列(LR)或一行(TB)。
        edges: [(src, dst) 或 dict(src=..., dst=..., label=...)]。节点在各自 rank 内居中均分。
        """
        d = cls(width_in, height_in, **kw)
        R = len(ranks)
        main = (width_in if direction == "LR" else height_in) - 2 * margin
        cross = (height_in if direction == "LR" else width_in) - 2 * margin
        rank_len = (main - gap_main * (R - 1)) / R
        for ri, rank in enumerate(ranks):
            k = len(rank)
            if node_size:
                cell = node_size[1] if direction == "LR" else node_size[0]
            else:
                cell = min(1.1 if direction == "LR" else 2.2, (cross - gap_cross * (k - 1)) / k)
            used = k * cell + (k - 1) * gap_cross
            start = margin + (cross - used) / 2
            for ci, item in enumerate(rank):
                nid, txt = item[0], item[1]
                role = item[2] if len(item) > 2 else "plain"
                m0 = margin + ri * (rank_len + gap_main)
                c0 = start + ci * (cell + gap_cross)
                if direction == "LR":
                    d.node(nid, m0, c0, rank_len, cell, txt, role=role)
                else:
                    d.node(nid, c0, m0, cell, rank_len, txt, role=role)
        for e in edges:
            if isinstance(e, dict):
                e = dict(e); d.edge(e.pop("src"), e.pop("dst"), **e)
            else:
                d.edge(e[0], e[1])
        return d

    @classmethod
    def from_spec(cls, spec):
        """
        JSON IR → 图。spec 字段：
          width, height (in), base_pt?, line_spacing?,
          nodes: [{id, x, y, w, h, text, role?, pt?, align?, wrap?, dashed?, check?}],
          groups?: [{x, y, w, h, label?, color?, dashed?, pt?}],
          edges?: [{from, to, label?, route?, via?, color?, dashed?, both?, src_side?, dst_side?, label_pos?, label_offset?}],
          arrows?: [{points: [[x, y], ...], color?, dashed?, width?, both?}],
          texts?: [{x, y, content, pt?, color?, anchor?}]
        也可用 {"layout": "layered", "direction": "LR", "ranks": [...], "edges": [...]} 自动布局。
        """
        if isinstance(spec, str):
            with open(spec, "r", encoding="utf-8") as f:
                spec = json.load(f)
        if spec.get("layout") == "layered":
            ranks = [[tuple(n) if isinstance(n, list) else (n["id"], n["text"], n.get("role", "plain")) for n in r]
                     for r in spec["ranks"]]
            edges = [dict(src=e["from"], dst=e["to"], **{k: v for k, v in e.items() if k not in ("from", "to")})
                     for e in spec.get("edges", [])]
            d = cls.layered(spec["width"], spec["height"], ranks, direction=spec.get("direction", "LR"),
                            edges=edges, base_pt=spec.get("base_pt", 18),
                            line_spacing=spec.get("line_spacing", 1.18))
        else:
            d = cls(spec["width"], spec["height"], base_pt=spec.get("base_pt", 18),
                    line_spacing=spec.get("line_spacing", 1.18))
            for n in spec["nodes"]:
                d.node(
                    n["id"], n["x"], n["y"], n["w"], n["h"], n["text"],
                    role=n.get("role", "plain"), pt=n.get("pt"),
                    align=n.get("align", "center"), wrap=n.get("wrap", True),
                    dashed=n.get("dashed", False), check=n.get("check", True),
                )
            for e in spec.get("edges", []):
                kw = {k: v for k, v in e.items() if k not in ("from", "to")}
                d.edge(e["from"], e["to"], **kw)
        for g in spec.get("groups", []):
            d.group(g["x"], g["y"], g["w"], g["h"], g.get("label"), g.get("color", "blue"),
                    g.get("dashed", True), g.get("pt"))
        for a in spec.get("arrows", []):
            d.arrow(a["points"], color=a.get("color", "blue"), dashed=a.get("dashed", False),
                    width=a.get("width", 2.4), both=a.get("both", False))
        for t in spec.get("texts", []):
            d.text(t["x"], t["y"], t["content"], t.get("pt"), t.get("color", "black"), t.get("anchor", "start"))
        return d

    # ---------------- 路由 ----------------
    @staticmethod
    def _anchor(n, side):
        x, y, w, h = n["x"], n["y"], n["w"], n["h"]
        return {"left": (x, y + h / 2), "right": (x + w, y + h / 2),
                "top": (x + w / 2, y), "bottom": (x + w / 2, y + h)}[side]

    def _route(self, e):
        a, b = self.nodes[e["src"]], self.nodes[e["dst"]]
        ac = (a["x"] + a["w"] / 2, a["y"] + a["h"] / 2)
        bc = (b["x"] + b["w"] / 2, b["y"] + b["h"] / 2)
        dx, dy = bc[0] - ac[0], bc[1] - ac[1]
        y_ov = min(a["y"] + a["h"], b["y"] + b["h"]) - max(a["y"], b["y"])
        x_ov = min(a["x"] + a["w"], b["x"] + b["w"]) - max(a["x"], b["x"])
        route = e["route"]
        ss, ds = e["src_side"], e["dst_side"]
        if e["via"]:
            first, last = e["via"][0], e["via"][-1]
            ss = ss or self._side_towards(a, first)
            ds = ds or self._side_towards(b, last)
            p, q = self._anchor(a, ss), self._anchor(b, ds)
            pts = [p]
            # 让首末段保持正交
            if ss in ("left", "right"):
                pts.append((first[0], p[1]))
            else:
                pts.append((p[0], first[1]))
            pts += e["via"]
            if ds in ("left", "right"):
                pts.append((last[0], q[1]))
            else:
                pts.append((q[0], last[1]))
            pts.append(q)
            return self._dedupe(pts)
        if route == "auto":
            if ss is None and ds is None and y_ov > 6 and abs(dx) > 1:
                y = max(a["y"], b["y"]) + y_ov / 2
                return [(a["x"] + a["w"] if dx > 0 else a["x"], y), (b["x"] if dx > 0 else b["x"] + b["w"], y)]
            if ss is None and ds is None and x_ov > 6 and abs(dy) > 1:
                x = max(a["x"], b["x"]) + x_ov / 2
                return [(x, a["y"] + a["h"] if dy > 0 else a["y"]), (x, b["y"] if dy > 0 else b["y"] + b["h"])]
            route = "hvh" if abs(dx) >= abs(dy) else "vhv"
        if route in ("straight", "hvh"):
            horiz = route == "hvh" or abs(dx) >= abs(dy)
            ss = ss or (("right" if dx > 0 else "left") if horiz else ("bottom" if dy > 0 else "top"))
            ds = ds or (("left" if dx > 0 else "right") if horiz else ("top" if dy > 0 else "bottom"))
        elif route == "vhv":
            ss = ss or ("bottom" if dy > 0 else "top")
            ds = ds or ("top" if dy > 0 else "bottom")
        elif route == "hv":
            ss = ss or ("right" if dx > 0 else "left")
            ds = ds or ("top" if dy > 0 else "bottom")
        elif route == "vh":
            ss = ss or ("bottom" if dy > 0 else "top")
            ds = ds or ("left" if dx > 0 else "right")
        p, q = self._anchor(a, ss), self._anchor(b, ds)
        if route == "straight":
            return [p, q]
        h1, h2 = ss in ("left", "right"), ds in ("left", "right")
        if h1 and h2:
            mx = (p[0] + q[0]) / 2
            return self._dedupe([p, (mx, p[1]), (mx, q[1]), q])
        if not h1 and not h2:
            my = (p[1] + q[1]) / 2
            return self._dedupe([p, (p[0], my), (q[0], my), q])
        if h1:
            return self._dedupe([p, (q[0], p[1]), q])
        return self._dedupe([p, (p[0], q[1]), q])

    @staticmethod
    def _side_towards(n, pt):
        cx, cy = n["x"] + n["w"] / 2, n["y"] + n["h"] / 2
        dx, dy = pt[0] - cx, pt[1] - cy
        if abs(dx) / max(n["w"], 1) >= abs(dy) / max(n["h"], 1):
            return "right" if dx > 0 else "left"
        return "bottom" if dy > 0 else "top"

    @staticmethod
    def _dedupe(pts):
        out = [pts[0]]
        for p in pts[1:]:
            if abs(p[0] - out[-1][0]) > 0.01 or abs(p[1] - out[-1][1]) > 0.01:
                out.append(p)
        return out

    # ---------------- 文字排版 ----------------
    def _layout_text(self, n):
        pad_x = 0.07 * PPI
        max_w = n["w"] - 2 * pad_x
        lines = []
        for raw in str(n["text"]).split("\n"):
            lines += wrap_line(raw, n["px"], max_w) if n["wrap"] else [raw]
        return lines

    def _label_box(self, e, pts):
        if not e["label"]:
            return None
        px = self.base_px
        segs = list(zip(pts[:-1], pts[1:]))
        p, q = max(segs, key=lambda s: math.hypot(s[1][0] - s[0][0], s[1][1] - s[0][1]))
        t = e["label_pos"]
        mx, my = p[0] + (q[0] - p[0]) * t, p[1] + (q[1] - p[1]) * t
        tw = line_width_px(e["label"], px) + 8
        th = px * 1.25
        if abs(q[1] - p[1]) < 0.5:  # 水平段：标签放线上方
            bx, by = mx - tw / 2, my - th - 3
        else:                         # 竖直段：标签放线右侧
            bx, by = mx + 5, my - th / 2
        bx += e["label_offset"][0]
        by += e["label_offset"][1]
        return (bx, by, tw, th)

    # ---------------- 校验 ----------------
    def validate(self):
        issues = []
        all_text = [(f"node {k!r}", v["text"]) for k, v in self.nodes.items()]
        all_text += [(f"edge label", e["label"]) for e in self.edges if e["label"]]
        all_text += [(f"group label", g["label"]) for g in self.groups if g["label"]]
        all_text += [(f"text", t["content"]) for t in self.texts]
        for where, txt in all_text:
            miss = missing_glyphs(txt)
            if miss:
                issues.append(f"{where}: no glyph for {''.join(miss)!r} in {LATIN_FAMILY} or {CJK_FAMILY} "
                              f"(renders as tofu) — rephrase, or put the symbol in native OMML on the slide")
            if _COMBINING.search(str(txt)):
                issues.append(f"{where}: combining accent in {txt!r} — cairosvg cannot compose it (renders as tofu). "
                              f"Put hatted symbols in native OMML on the slide, or rephrase (e.g. '预测扰动 d_{{k+N}}').")
        for nid in self._order:
            n = self.nodes[nid]
            lines = self._layout_text(n)
            need_h = len(lines) * n["px"] * self.line_spacing + 0.06 * PPI
            if need_h > n["h"] + 0.5:
                issues.append(f"text overflow: node {nid!r} needs {need_h / PPI:.2f} in height "
                              f"({len(lines)} lines) but has {n['h'] / PPI:.2f} in")
            for ln in lines:
                if line_width_px(ln, n["px"]) > n["w"] - 0.08 * PPI + 0.5:
                    issues.append(f"text overflow: node {nid!r} line {ln!r} wider than box")
            if n["x"] < -0.5 or n["y"] < -0.5 or n["x"] + n["w"] > self.W + 0.5 or n["y"] + n["h"] > self.H + 0.5:
                issues.append(f"out of canvas: node {nid!r}")
        ids = self._order
        for i in range(len(ids)):
            for j in range(i + 1, len(ids)):
                a, b = self.nodes[ids[i]], self.nodes[ids[j]]
                if not (a["check"] and b["check"]):
                    continue
                if _rect_overlap((a["x"], a["y"], a["w"], a["h"]), (b["x"], b["y"], b["w"], b["h"])):
                    issues.append(f"node overlap: {ids[i]!r} × {ids[j]!r}")
        label_boxes = []
        for e in self.edges:
            for k in ("src", "dst"):
                if e[k] not in self.nodes:
                    issues.append(f"edge references unknown node {e[k]!r}")
            if e["src"] not in self.nodes or e["dst"] not in self.nodes:
                continue
            pts = self._route(e)
            for pt_xy in pts:
                if pt_xy[0] < -0.5 or pt_xy[1] < -0.5 or pt_xy[0] > self.W + 0.5 or pt_xy[1] > self.H + 0.5:
                    issues.append(f"out of canvas: edge {e['src']}->{e['dst']} waypoint")
                    break
            for p, q in zip(pts[:-1], pts[1:]):
                for nid, n in self.nodes.items():
                    if nid in (e["src"], e["dst"]) or not n["check"]:
                        continue
                    if _seg_hits_rect(p, q, (n["x"], n["y"], n["w"], n["h"])):
                        issues.append(f"edge {e['src']}->{e['dst']} passes through node {nid!r}")
            lb = self._label_box(e, pts)
            if lb:
                for nid, n in self.nodes.items():
                    if _rect_overlap(lb, (n["x"], n["y"], n["w"], n["h"]), tol=1.0):
                        issues.append(f"edge label {e['label']!r} overlaps node {nid!r}")
                for other in label_boxes:
                    if _rect_overlap(lb, other, tol=1.0):
                        issues.append(f"edge label {e['label']!r} overlaps another label")
                if lb[0] < 0 or lb[1] < 0 or lb[0] + lb[2] > self.W or lb[1] + lb[3] > self.H:
                    issues.append(f"edge label {e['label']!r} out of canvas")
                label_boxes.append(lb)
        for a in self.arrows:
            if len(a["pts"]) < 2:
                issues.append("free arrow needs at least 2 points")
                continue
            for pt_xy in a["pts"]:
                if pt_xy[0] < -0.5 or pt_xy[1] < -0.5 or pt_xy[0] > self.W + 0.5 or pt_xy[1] > self.H + 0.5:
                    issues.append("out of canvas: free arrow")
                    break
            for p, q in zip(a["pts"][:-1], a["pts"][1:]):
                for nid, n in self.nodes.items():
                    if n["check"] and _seg_hits_rect(p, q, (n["x"], n["y"], n["w"], n["h"])):
                        issues.append(f"free arrow passes through node {nid!r}")
        for t in self.texts:
            lines = str(t["content"]).split("\n")
            lh = t["px"] * self.line_spacing
            for li, ln in enumerate(lines):
                tw = line_width_px(ln, t["px"])
                x0 = t["x"] - (tw / 2 if t["anchor"] == "middle" else tw if t["anchor"] == "end" else 0)
                y_line = t["y"] + li * lh
                if x0 < -0.5 or x0 + tw > self.W + 0.5 or y_line - t["px"] < -0.5 or y_line > self.H + 0.5:
                    issues.append(f"text {t['content']!r} out of canvas")
                tb = (x0, y_line - t["px"] * 0.9, tw, t["px"] * 1.1)
                for nid, n in self.nodes.items():
                    if _rect_overlap(tb, (n["x"], n["y"], n["w"], n["h"]), tol=1.0):
                        issues.append(f"text {t['content']!r} overlaps node {nid!r}")
        group_labels = []
        for g in self.groups:
            if g["label"]:
                lw = line_width_px(g["label"], g["px"]) + 0.16 * PPI
                group_labels.append((g["label"], (g["x"] + 0.06 * PPI, g["y"] + 0.03 * PPI, lw, g["px"] * 1.2)))
        for glabel, gb in group_labels:
            for e in self.edges:
                if e["src"] not in self.nodes or e["dst"] not in self.nodes:
                    continue
                pts = self._route(e)
                if any(_seg_hits_rect(p, q, gb, pad=0.0) for p, q in zip(pts[:-1], pts[1:])):
                    issues.append(f"edge {e['src']}->{e['dst']} crosses group label {glabel!r}")
            for a in self.arrows:
                if any(_seg_hits_rect(p, q, gb, pad=0.0) for p, q in zip(a["pts"][:-1], a["pts"][1:])):
                    issues.append(f"free arrow crosses group label {glabel!r}")
        for g in self.groups:
            if g["x"] < -0.5 or g["y"] < -0.5 or g["x"] + g["w"] > self.W + 0.5 or g["y"] + g["h"] > self.H + 0.5:
                issues.append(f"out of canvas: group {g['label']!r}")
            if g["label"]:
                lw = line_width_px(g["label"], g["px"]) + 0.16 * PPI
                if lw > g["w"]:
                    issues.append(f"group label {g['label']!r} wider than group")
                lb = (g["x"] + 0.06 * PPI, g["y"] + 0.03 * PPI, lw, g["px"] * 1.2)
                for nid, n in self.nodes.items():
                    if _rect_overlap(lb, (n["x"], n["y"], n["w"], n["h"]), tol=1.0):
                        issues.append(f"group label {g['label']!r} overlaps node {nid!r}")
        return list(dict.fromkeys(issues))

    # ---------------- SVG 输出 ----------------
    def _text_svg(self, x, y, line, px, color, anchor="middle"):
        """输出一行 <text>：CJK 段 / 西文段分别指定字体，_{}/^{} 用 dy 上下移位。"""
        parts = []
        for txt, kind in _tokens(line):
            size = px * (SUBSUP_SCALE if kind != "normal" else 1.0)
            shift = {"normal": 0.0, "sub": 0.28 * px, "sup": -0.38 * px}[kind]
            for i, (seg, cjk) in enumerate(_split_script(txt)):
                fam = CJK_FAMILY if cjk else LATIN_FAMILY
                parts.append((html.escape(seg), fam, size, shift))
        out, cur_shift = [], 0.0
        for seg, fam, size, shift in parts:
            dy = shift - cur_shift
            cur_shift = shift
            dy_attr = f' dy="{dy:.2f}"' if abs(dy) > 0.01 else ""
            out.append(f'<tspan font-family="{fam}" font-size="{size:.2f}"{dy_attr}>{seg}</tspan>')
        # 由于 anchor=middle 依赖渲染器测宽，这里统一用 start + 预先计算的偏移，结果更稳定
        w = line_width_px(line, px)
        if anchor == "middle":
            x0 = x - w / 2
        elif anchor == "end":
            x0 = x - w
        else:
            x0 = x
        return (f'<text x="{x0:.2f}" y="{y:.2f}" fill="{color}" font-weight="bold" '
                f'xml:space="preserve">{"".join(out)}</text>')

    @staticmethod
    def _arrowhead(p, q, color, size=11.0):
        ang = math.atan2(q[1] - p[1], q[0] - p[0])
        a1, a2 = ang + math.radians(152), ang - math.radians(152)
        pts = [q, (q[0] + size * math.cos(a1), q[1] + size * math.sin(a1)),
               (q[0] + size * math.cos(a2), q[1] + size * math.sin(a2))]
        return (f'<polygon points="{" ".join(f"{x:.2f},{y:.2f}" for x, y in pts)}" '
                f'fill="{color}" stroke="none"/>')

    def to_svg(self):
        o = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{self.W:.1f}" height="{self.H:.1f}" '
             f'viewBox="0 0 {self.W:.1f} {self.H:.1f}">',
             f'<rect x="0" y="0" width="{self.W:.1f}" height="{self.H:.1f}" fill="{WHITE}"/>']
        for g in self.groups:
            dash = ' stroke-dasharray="6,4"' if g["dashed"] else ""
            o.append(f'<rect x="{g["x"]:.2f}" y="{g["y"]:.2f}" width="{g["w"]:.2f}" height="{g["h"]:.2f}" '
                     f'fill="none" stroke="{g["color"]}" stroke-width="2"{dash}/>')
            if g["label"]:
                o.append(self._text_svg(g["x"] + 0.10 * PPI, g["y"] + 0.03 * PPI + g["px"] * 0.95,
                                        g["label"], g["px"], g["color"], anchor="start"))
        for e in self.edges:
            pts = self._route(e)
            # 线段终点回缩，让箭头尖端正好落在边框上
            body = list(pts)
            p, q = body[-2], body[-1]
            L = math.hypot(q[0] - p[0], q[1] - p[1]) or 1
            body[-1] = (q[0] - (q[0] - p[0]) / L * 8, q[1] - (q[1] - p[1]) / L * 8)
            if e["both"]:
                p0, p1 = body[0], body[1]
                L0 = math.hypot(p1[0] - p0[0], p1[1] - p0[1]) or 1
                body[0] = (p0[0] + (p1[0] - p0[0]) / L0 * 8, p0[1] + (p1[1] - p0[1]) / L0 * 8)
            dash = ' stroke-dasharray="8,5"' if e["dashed"] else ""
            o.append(f'<polyline points="{" ".join(f"{x:.2f},{y:.2f}" for x, y in body)}" fill="none" '
                     f'stroke="{e["color"]}" stroke-width="{e["width"]}" stroke-linejoin="miter"{dash}/>')
            o.append(self._arrowhead(pts[-2], pts[-1], e["color"]))
            if e["both"]:
                o.append(self._arrowhead(pts[1], pts[0], e["color"]))
            lb = self._label_box(e, pts)
            if lb:
                o.append(f'<rect x="{lb[0]:.2f}" y="{lb[1]:.2f}" width="{lb[2]:.2f}" height="{lb[3]:.2f}" fill="{WHITE}"/>')
                o.append(self._text_svg(lb[0] + lb[2] / 2, lb[1] + lb[3] * 0.80, e["label"], self.base_px, e["color"]))
        for a in self.arrows:
            pts = a["pts"]
            if len(pts) < 2:
                continue
            body = list(pts)
            p, q = body[-2], body[-1]
            L = math.hypot(q[0] - p[0], q[1] - p[1]) or 1
            body[-1] = (q[0] - (q[0] - p[0]) / L * 8, q[1] - (q[1] - p[1]) / L * 8)
            if a["both"]:
                p0, p1 = body[0], body[1]
                L0 = math.hypot(p1[0] - p0[0], p1[1] - p0[1]) or 1
                body[0] = (p0[0] + (p1[0] - p0[0]) / L0 * 8, p0[1] + (p1[1] - p0[1]) / L0 * 8)
            dash = ' stroke-dasharray="8,5"' if a["dashed"] else ""
            o.append(f'<polyline points="{" ".join(f"{x:.2f},{y:.2f}" for x, y in body)}" fill="none" '
                     f'stroke="{a["color"]}" stroke-width="{a["width"]}"{dash}/>')
            o.append(self._arrowhead(pts[-2], pts[-1], a["color"]))
            if a["both"]:
                o.append(self._arrowhead(pts[1], pts[0], a["color"]))
        for nid in self._order:
            n = self.nodes[nid]
            fill, stroke, tcol, sw = ROLES[n["role"]]
            st = f' stroke="{stroke}" stroke-width="{sw}"' if stroke else ' stroke="none"'
            if n.get("dashed") and stroke:
                st += ' stroke-dasharray="6,4"'
            o.append(f'<rect x="{n["x"]:.2f}" y="{n["y"]:.2f}" width="{n["w"]:.2f}" height="{n["h"]:.2f}" '
                     f'rx="4" fill="{fill}"{st}/>')
            lines = self._layout_text(n)
            lh = n["px"] * self.line_spacing
            total = lh * len(lines)
            y0 = n["y"] + (n["h"] - total) / 2 + n["px"] * 0.92
            for i, ln in enumerate(lines):
                if n["align"] == "left":
                    o.append(self._text_svg(n["x"] + 0.08 * PPI, y0 + i * lh, ln, n["px"], tcol, anchor="start"))
                else:
                    o.append(self._text_svg(n["x"] + n["w"] / 2, y0 + i * lh, ln, n["px"], tcol))
        for t in self.texts:
            lines = str(t["content"]).split("\n")
            lh = t["px"] * self.line_spacing
            for li, ln in enumerate(lines):
                o.append(self._text_svg(t["x"], t["y"] + li * lh, ln, t["px"], t["color"], anchor=t["anchor"]))
        o.append("</svg>")
        return "\n".join(o)

    def save(self, png_path, scale=2.0, svg_path=None, strict=True):
        """校验 → 写 SVG（可选）→ 渲染 PNG。strict=True 时存在问题直接抛 ValueError。"""
        issues = self.validate()
        if issues:
            msg = "TjuDiagram validation failed:\n  - " + "\n  - ".join(issues)
            if strict:
                raise ValueError(msg)
            print("[tju-ppt][warn] " + msg)
        svg = self.to_svg()
        if svg_path:
            os.makedirs(os.path.dirname(os.path.abspath(svg_path)), exist_ok=True)
            with open(svg_path, "w", encoding="utf-8") as f:
                f.write(svg)
        import cairosvg
        os.makedirs(os.path.dirname(os.path.abspath(png_path)), exist_ok=True)
        cairosvg.svg2png(bytestring=svg.encode("utf-8"), write_to=png_path, scale=scale)
        return png_path


def font_report():
    """打印当前解析到的字体，排查“西文不是 Times / 中文成方块”问题。"""
    return {"latin_family": LATIN_FAMILY, "latin_file": LATIN_FILE,
            "cjk_family": CJK_FAMILY, "cjk_file": CJK_FILE}


if __name__ == "__main__":
    print(json.dumps(font_report(), ensure_ascii=False, indent=2))
