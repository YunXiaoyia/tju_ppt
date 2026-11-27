#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tjuppt.py — one import for deck scripts.

    import sys; sys.path.insert(0, "<skill_dir>/scripts")
    from tjuppt import *

Modules: tju_theme (palette, grid, width budget) · tju_layout (Region) · tju_text (runs) ·
tju_components (panels, cards, tables) · tju_omml (equations) · tju_deck (template, Deck, ContentPage) ·
tju_native_diagrams (editable flows) · tju_svg_diagram (TjuDiagram) · tju_plot_style (matplotlib) ·
tju_audit / tju_render / tju_script_check (CLIs).
"""

from tju_theme import *  # noqa: F401,F403
from tju_layout import Region, as_box, full_width  # noqa: F401
from tju_text import (  # noqa: F401
    append_rich_text, set_run_font, update_paragraph_text, update_shape_paragraphs, write_paragraph,
)
from tju_components import (  # noqa: F401
    add_arrow, add_card, card_height, check_aspect, add_clean_tju_table, add_dashed_box, add_figure_panel, add_image_card, add_image_fit,
    add_joined_card, add_picture_exact, image_card_slot, add_solid_box, add_table, add_text_block, figure_slot,
    text_capacity_check, update_table_cells,
)
from tju_omml import add_omml_equation, m_acc, m_frac, m_r, m_sqrt, m_sub, m_subsup, m_sup  # noqa: F401
from tju_deck import (  # noqa: F401
    TEMPLATE_PATH, ContentPage, Deck, clear_all_slides, clear_slide_body, enforce_min_18pt_all_slides,
    format_cover_slide, format_thanks_slide, format_toc_slide, get_tju_layout, new_content_slide,
    new_cover_slide, new_thanks_slide, new_toc_slide, rearrange_slides, save_staged_pptx, setup_tju_content_slide,
)
from tju_native_diagrams import (  # noqa: F401
    add_connector, add_diagram_group, add_diagram_label, add_feedback_loop, add_flow_pipeline,
    add_junction_dot, add_layered_architecture, add_native_panel, add_pill, add_polyline_arrow,
    add_role_box, add_timeline, connect_boxes,
)
