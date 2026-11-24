#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Compatibility shim for deck scripts written before the v2 refactor.  New scripts: `from tjuppt import *`."""

from tjuppt import *  # noqa: F401,F403
from tju_deck import _in_mc_fallback, _remove_alternate_content  # noqa: F401
from tju_text import ensure_font_slots as _ensure_four_font_slots  # noqa: F401
from tju_text import SUB_SUP_PATTERN  # noqa: F401
