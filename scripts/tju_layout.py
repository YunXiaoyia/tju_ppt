#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tju_layout.py — grid geometry.  Compute regions; never type raw coordinates in a deck script.

    page.body                    → Region of the free content area (after lead lines / before takeaway)
    page.body.split(2)           → two columns separated by GUTTER, flush with GRID_L / GRID_R
    page.body.split(2, [5, 7])   → weighted columns
    r.take_top(0.40)             → (top strip, remainder)
    r.vsplit(3)                  → stacked rows separated by VGAP

Because every region is derived from the grid, left/right edges and gutters stay identical across pages.
"""

from dataclasses import dataclass

from tju_theme import GRID_L, GRID_W, GUTTER, VGAP


@dataclass(frozen=True)
class Region:
    x: float
    y: float
    w: float
    h: float

    # ---------------- geometry ----------------
    @property
    def right(self):
        return self.x + self.w

    @property
    def bottom(self):
        return self.y + self.h

    @property
    def box(self):
        return (self.x, self.y, self.w, self.h)

    @property
    def aspect(self):
        return self.w / self.h

    def split(self, n=2, ratios=None, gutter=GUTTER):
        """Split horizontally into n columns (optionally weighted) with a fixed gutter."""
        ratios = list(ratios) if ratios else [1.0] * n
        if len(ratios) != n:
            n = len(ratios)
        usable = self.w - gutter * (n - 1)
        tot = float(sum(ratios))
        out, x = [], self.x
        for i, r in enumerate(ratios):
            w = usable * r / tot
            if i == n - 1:                       # absorb rounding so the last edge is exactly flush
                w = self.right - x
            out.append(Region(x, self.y, w, self.h))
            x += w + gutter
        return out

    def vsplit(self, n=2, ratios=None, gap=VGAP):
        """Split vertically into n rows (optionally weighted) with a fixed gap."""
        ratios = list(ratios) if ratios else [1.0] * n
        n = len(ratios)
        usable = self.h - gap * (n - 1)
        tot = float(sum(ratios))
        out, y = [], self.y
        for i, r in enumerate(ratios):
            h = usable * r / tot
            if i == n - 1:
                h = self.bottom - y
            out.append(Region(self.x, y, self.w, h))
            y += h + gap
        return out

    def take_top(self, h, gap=0.0):
        """(top strip of height h, remainder below it after `gap`)."""
        return Region(self.x, self.y, self.w, h), Region(self.x, self.y + h + gap, self.w, self.h - h - gap)

    def take_bottom(self, h, gap=0.0):
        return Region(self.x, self.bottom - h, self.w, h), Region(self.x, self.y, self.w, self.h - h - gap)

    def take_left(self, w, gap=GUTTER):
        return Region(self.x, self.y, w, self.h), Region(self.x + w + gap, self.y, self.w - w - gap, self.h)

    def inset(self, dx=0.10, dy=None):
        dy = dx if dy is None else dy
        return Region(self.x + dx, self.y + dy, self.w - 2 * dx, self.h - 2 * dy)

    def with_height(self, h):
        return Region(self.x, self.y, self.w, h)


def as_box(region):
    """Accept a Region or an (x, y, w, h) tuple."""
    if isinstance(region, Region):
        return region.box
    x, y, w, h = region
    return float(x), float(y), float(w), float(h)


def full_width(y, h):
    return Region(GRID_L, y, GRID_W, h)
