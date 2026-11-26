#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Contract tests for the tju-ppt skill.  Run: python tests/test_tjuppt.py"""

import os
import subprocess
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(HERE, "..", "scripts")
sys.path.insert(0, SCRIPTS)

from PIL import Image  # noqa: E402
from pptx.enum.text import PP_ALIGN  # noqa: E402

from tjuppt import (  # noqa: E402
    C_BLUE, C_RED, C_WHITE, GRID_L, GRID_R, GUTTER, Deck, LayoutError, Region, add_card, add_dashed_box,
    add_figure_panel, add_image_fit, add_solid_box, add_table, add_text_block, figure_slot,
)
from tju_audit import audit_content, audit_pptx  # noqa: E402

TMP = tempfile.mkdtemp(prefix="tjuppt_test_")


def png(w_in, h_in, name):
    p = os.path.join(TMP, name)
    Image.new("RGB", (int(w_in * 100), int(h_in * 100)), (0, 51, 127)).save(p)
    return p


def save(deck, name):
    return deck.save(os.path.join(TMP, name), audit=False)[0]


class GridTests(unittest.TestCase):
    def test_split_is_flush_and_gutters_equal(self):
        cols = Region(GRID_L, 2.0, GRID_R - GRID_L, 3.0).split(3, [2, 3, 4])
        self.assertAlmostEqual(cols[0].x, GRID_L)
        self.assertAlmostEqual(cols[-1].right, GRID_R)
        gaps = [b.x - a.right for a, b in zip(cols, cols[1:])]
        for g in gaps:
            self.assertAlmostEqual(g, GUTTER)

    def test_body_zones(self):
        d = Deck()
        p = d.content("t", "s", lead=[("a：", "b"), ("c：", "d")], takeaway="结论")
        self.assertAlmostEqual(p.body.x, GRID_L)
        self.assertLess(p.body.bottom, 6.42)
        self.assertGreater(p.body.y, 2.4)


class BuildTimeChecks(unittest.TestCase):
    def test_long_takeaway_raises(self):
        with self.assertRaises(LayoutError):
            Deck().content("t", "s", takeaway="很长的结论" * 12)

    def test_lead_limits(self):
        with self.assertRaises(LayoutError):
            Deck().content("t", "s", lead=[("a：", "b")] * 4)
        with self.assertRaises(LayoutError):
            Deck().content("t", "s", lead=[("标签：", "非常长的引导句" * 10)])

    def test_figure_aspect_enforced(self):
        d = Deck()
        p = d.content("t", "s")
        left, right = p.body.split(2)
        with self.assertRaises(LayoutError):
            add_figure_panel(p.slide, left, "图", png(4, 4, "square.png"))
        w, h = figure_slot(left)
        add_figure_panel(p.slide, left, "图", png(w, h, "exact.png"))
        add_figure_panel(p.slide, right, "照片", png(4, 4, "photo.png"), fit="crop")
        self.assertEqual(audit_pptx(save(d, "fig.pptx")), [])

    def test_card_and_table_capacity(self):
        d = Deck()
        p = d.content("t", "s")
        with self.assertRaises(LayoutError):
            add_card(p.slide, p.body.with_height(1.0), "卡片", ["一行文字"] * 6)
        with self.assertRaises(LayoutError):
            add_table(p.slide, p.body.with_height(1.0), [1, 2], [["a", "b"]] + [["x", "y"]] * 6)


class AuditTests(unittest.TestCase):
    def test_letterboxed_image_under_header_is_flagged(self):
        d = Deck()
        p = d.content("t", "s")
        add_solid_box(p.slide, GRID_L, 2.0, 6.0, 0.4, ["标题"], fill_color=C_BLUE, line_color=C_BLUE, text_color=C_WHITE)
        add_image_fit(p.slide, png(4, 4, "sq2.png"), GRID_L, 2.4, 6.0, 3.0)
        issues = audit_pptx(save(d, "letterbox.pptx"))
        self.assertTrue(any("not flush" in i for i in issues), issues)

    def test_outside_grid_is_flagged(self):
        d = Deck()
        p = d.content("t", "s")
        add_text_block(p.slide, 0.2, 2.0, 3.0, 0.5, ["出界"])
        self.assertTrue(any("outside content grid" in i for i in audit_pptx(save(d, "grid.pptx"))))

    def test_takeaway_wrapping_is_flagged(self):
        d = Deck()
        p = d.content("t", "s")
        add_text_block(p.slide, GRID_L, 6.42, 4.0, 0.46, [("这是一条很长很长很长很长的红字结论句子", C_RED, 19)],
                       align=PP_ALIGN.CENTER)
        self.assertTrue(any("takeaway wraps" in i for i in audit_pptx(save(d, "tk.pptx"))))

    def test_content_warnings(self):
        d = Deck()
        p = d.content("t", "工作一：强化学习（RL）方法", takeaway="结论一")
        col = p.body.split(2)[0]
        add_dashed_box(p.slide, *col.box)
        add_card(p.slide, col.inset(0.2), "卡片", ["RL 策略"])
        p2 = d.content("t", "s")
        add_text_block(p2.slide, GRID_L, 6.42, 12.43, 0.46, [("另一条结论", C_RED, 18)], align=PP_ALIGN.CENTER)
        w = audit_content(save(d, "warn.pptx"))
        self.assertTrue(any("dashed container" in x for x in w), w)
        self.assertTrue(any("'工作一'" in x for x in w), w)
        self.assertTrue(any("different sizes" in x for x in w), w)
        self.assertFalse(any("first use of RL" in x for x in w), w)   # defined in place


    def test_defensive_wording_warnings(self):
        d = Deck()
        for i in range(3):
            p = d.content("t", f"研究内容{i}：复用既有感知结果")
        add_text_block(p.slide, GRID_L, 2.0, 12.0, 0.5, ["输出动作，直接调用既有控制（去 VLA/ACT）"])
        p2 = d.content("t", "研究内容二：调度方法")
        add_text_block(p2.slide, GRID_L, 3.0, 12.0, 0.5, ["研究状态编码方法：参考场景图 [5]"])
        w = audit_content(save(d, "defensive.pptx"))
        self.assertTrue(any("revision note" in x for x in w), w)
        self.assertTrue(any("citation on a research-content page" in x for x in w), w)
        self.assertTrue(any("disclaimer repeated" in x for x in w), w)


class FigureTests(unittest.TestCase):
    def test_edge_crossing_group_label(self):
        from tju_svg_diagram import TjuDiagram
        d = TjuDiagram(5, 3)
        d.group(0.1, 0.1, 4.8, 2.8, "分组标签很长的说明")
        d.node("a", 0.3, 0.9, 1.2, 0.6, "A")
        d.node("b", 3.3, 2.1, 1.2, 0.6, "B")
        d.arrow([(0.2, 0.25), (4.0, 0.25)])
        self.assertTrue(any("group label" in i for i in d.validate()))

    def test_symbol_routing(self):
        import tju_svg_diagram as t
        lat, cjk = t._cmaps()
        if lat is None or cjk is None:
            self.skipTest("font files not resolvable")
        if 0x2016 not in lat and 0x2016 in cjk:
            self.assertTrue(t._route_cjk("‖"))
        self.assertEqual(t.missing_glyphs("误差 ‖e‖"), [])
        self.assertTrue(t.missing_glyphs("\ue000"))


class ScriptCheckTests(unittest.TestCase):
    def run_check(self, body, budget):
        p = os.path.join(TMP, "s.md")
        open(p, "w", encoding="utf-8").write(body)
        return subprocess.run([sys.executable, os.path.join(SCRIPTS, "tju_script_check.py"), p, "--budget-s", str(budget)],
                              capture_output=True, text=True).returncode

    def test_rates(self):
        ok = "# 讲稿\n\n### P1 背景（约 60 秒）\n" + "字" * 190 + "\n"
        fast = "# 讲稿\n\n### P1 背景（约 30 秒）\n" + "字" * 190 + "\n"
        self.assertEqual(self.run_check(ok, 60), 0)
        self.assertEqual(self.run_check(fast, 30), 1)


class CompatTests(unittest.TestCase):
    def test_legacy_names(self):
        import ppt_builder_helpers as h
        for name in ("C_BLUE", "clear_all_slides", "new_content_slide", "add_solid_box", "add_joined_card",
                     "add_clean_tju_table", "add_text_block", "add_image_fit", "add_omml_equation", "m_r", "m_sub",
                     "enforce_min_18pt_all_slides", "save_staged_pptx", "add_dashed_box", "setup_tju_content_slide"):
            self.assertTrue(hasattr(h, name), name)


if __name__ == "__main__":
    unittest.main(verbosity=2)
