"""
pptx_engine.py - Professional Local PowerPoint Generation Engine
Pure Python (python-pptx), zero external APIs, 16:9 widescreen, pixel-perfect layouts.
"""

import os
from typing import List, Dict, Any, Optional
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION

# ---------------------------------------------------------
# THEMES & COLOR PALETTES
# ---------------------------------------------------------
THEMES = {
    "corporate": {
        "name": "Corporate Modern",
        "bg_dark": RGBColor(15, 23, 42),       # Slate 900
        "bg_light": RGBColor(248, 250, 252),   # Slate 50
        "card_bg": RGBColor(255, 255, 255),    # Pure White
        "card_border": RGBColor(226, 232, 240),# Slate 200
        "primary": RGBColor(2, 132, 199),      # Sky 600
        "secondary": RGBColor(14, 165, 233),   # Sky 500
        "accent": RGBColor(56, 189, 248),      # Light Blue
        "text_dark": RGBColor(15, 23, 42),     # Dark slate text
        "text_muted": RGBColor(100, 116, 139), # Slate 500
        "text_light": RGBColor(255, 255, 255), # White text
        "text_light_muted": RGBColor(203, 213, 225),
        "badge_bg": RGBColor(224, 242, 254),   # Sky 100
        "badge_text": RGBColor(3, 105, 161),   # Sky 700
        "chart_colors": [
            RGBColor(2, 132, 199),
            RGBColor(14, 165, 233),
            RGBColor(56, 189, 248),
            RGBColor(186, 230, 253)
        ]
    },
    "dark_tech": {
        "name": "Dark Tech / AI",
        "bg_dark": RGBColor(11, 15, 25),       # Deep Obsidian
        "bg_light": RGBColor(17, 24, 39),      # Gray 900
        "card_bg": RGBColor(24, 33, 47),       # Dark Slate Card
        "card_border": RGBColor(45, 55, 72),   # Border
        "primary": RGBColor(16, 185, 129),     # Emerald 500
        "secondary": RGBColor(52, 211, 153),   # Emerald 400
        "accent": RGBColor(110, 231, 183),     # Emerald 300
        "text_dark": RGBColor(255, 255, 255),  # White on dark
        "text_muted": RGBColor(148, 163, 184), # Slate 400
        "text_light": RGBColor(255, 255, 255),
        "text_light_muted": RGBColor(156, 163, 175),
        "badge_bg": RGBColor(6, 78, 59),       # Emerald 900
        "badge_text": RGBColor(110, 231, 183), # Emerald 300
        "chart_colors": [
            RGBColor(16, 185, 129),
            RGBColor(59, 130, 246),
            RGBColor(168, 85, 247),
            RGBColor(236, 72, 153)
        ]
    },
    "vibrant": {
        "name": "Vibrant Pitch",
        "bg_dark": RGBColor(24, 17, 36),       # Midnight Violet
        "bg_light": RGBColor(250, 245, 255),   # Purple 50
        "card_bg": RGBColor(255, 255, 255),    # White Card
        "card_border": RGBColor(233, 213, 255),# Purple 200
        "primary": RGBColor(99, 102, 241),     # Indigo 500
        "secondary": RGBColor(244, 63, 94),    # Rose 500
        "accent": RGBColor(168, 85, 247),      # Purple 500
        "text_dark": RGBColor(30, 27, 75),     # Deep Indigo
        "text_muted": RGBColor(107, 114, 128),
        "text_light": RGBColor(255, 255, 255),
        "text_light_muted": RGBColor(229, 231, 235),
        "badge_bg": RGBColor(238, 242, 255),
        "badge_text": RGBColor(67, 56, 202),
        "chart_colors": [
            RGBColor(99, 102, 241),
            RGBColor(244, 63, 94),
            RGBColor(168, 85, 247),
            RGBColor(251, 146, 60)
        ]
    }
}

class DeckBuilder:
    """High-level builder for 16:9 widescreen PowerPoint presentations."""

    def __init__(self, theme: str = "corporate", title_font: str = "Calibri", body_font: str = "Calibri"):
        self.prs = Presentation()
        # Enforce 16:9 Widescreen (13.333 x 7.5 inches)
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.blank_layout = self.prs.slide_layouts[6] # Blank slide layout
        self.theme = THEMES.get(theme, THEMES["corporate"])
        self.theme_name = theme
        self.title_font = title_font
        self.body_font = body_font

    def _set_slide_background(self, slide, color: RGBColor):
        """Sets a solid color background across the entire slide."""
        bg = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, 0, 0, self.prs.slide_width, self.prs.slide_height
        )
        bg.fill.solid()
        bg.fill.fore_color.rgb = color
        bg.line.fill.background() # No border
        return bg

    def _add_header(self, slide, category: str, title: str, subtitle: Optional[str] = None, is_dark: bool = False):
        """Standardized, high-contrast header for content slides."""
        tx_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.733), Inches(1.3))
        tf = tx_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        # Category Tracker / Overline
        if category:
            p_cat = tf.paragraphs[0]
            p_cat.text = category.upper()
            p_cat.font.name = self.title_font
            p_cat.font.size = Pt(10)
            p_cat.font.bold = True
            p_cat.font.color.rgb = self.theme["primary"]
            p_cat.space_after = Pt(4)
            p_title = tf.add_paragraph()
        else:
            p_title = tf.paragraphs[0]

        # Action Title
        p_title.text = title
        p_title.font.name = self.title_font
        p_title.font.size = Pt(24)
        p_title.font.bold = True
        p_title.font.color.rgb = self.theme["text_light"] if is_dark else self.theme["text_dark"]

        # Subtitle / Context
        if subtitle:
            p_sub = tf.add_paragraph()
            p_sub.text = subtitle
            p_sub.font.name = self.body_font
            p_sub.font.size = Pt(12)
            p_sub.font.color.rgb = self.theme["text_light_muted"] if is_dark else self.theme["text_muted"]
            p_sub.space_before = Pt(4)

    def _add_notes(self, slide, notes_text: str):
        """Appends plain-text speaker notes to the slide."""
        if notes_text:
            notes_slide = slide.notes_slide
            tf = notes_slide.notes_text_frame
            tf.text = notes_text

    # -----------------------------------------------------
    # SLIDE ARCHETYPES
    # -----------------------------------------------------

    def add_title_slide(self, title: str, subtitle: str, author: str = "", date: str = "", notes: str = ""):
        """Cover / Title slide with deep contrast background."""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._set_slide_background(slide, self.theme["bg_dark"])

        # Subtle decorative accent line
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(2.2), Inches(1.2), Inches(0.08))
        line.fill.solid()
        line.fill.fore_color.rgb = self.theme["primary"]
        line.line.fill.background()

        # Title & Subtitle Box
        box = slide.shapes.add_textbox(Inches(0.8), Inches(2.5), Inches(11.733), Inches(3.0))
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.name = self.title_font
        p1.font.size = Pt(42)
        p1.font.bold = True
        p1.font.color.rgb = self.theme["text_light"]
        p1.space_after = Pt(16)

        p2 = tf.add_paragraph()
        p2.text = subtitle
        p2.font.name = self.body_font
        p2.font.size = Pt(18)
        p2.font.color.rgb = self.theme["text_light_muted"]

        # Presenter Metadata Box
        if author or date:
            meta_box = slide.shapes.add_textbox(Inches(0.8), Inches(6.0), Inches(11.733), Inches(0.8))
            mtf = meta_box.text_frame
            mp = mtf.paragraphs[0]
            meta_str = " • ".join([item for item in [author, date] if item])
            mp.text = meta_str
            mp.font.name = self.body_font
            mp.font.size = Pt(11)
            mp.font.color.rgb = self.theme["text_light_muted"]

        self._add_notes(slide, notes)
        return slide

    def add_statement_slide(self, category: str, statement: str, context: str = "", notes: str = ""):
        """Big statement / mindset shift slide."""
        slide = self.prs.slides.add_slide(self.blank_layout)
        is_dark = (self.theme_name == "dark_tech")
        self._set_slide_background(slide, self.theme["bg_dark"] if is_dark else self.theme["bg_light"])

        box = slide.shapes.add_textbox(Inches(1.2), Inches(2.0), Inches(10.933), Inches(3.5))
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        if category:
            p_cat = tf.paragraphs[0]
            p_cat.text = category.upper()
            p_cat.font.name = self.title_font
            p_cat.font.size = Pt(12)
            p_cat.font.bold = True
            p_cat.font.color.rgb = self.theme["primary"]
            p_cat.space_after = Pt(14)
            p_stmt = tf.add_paragraph()
        else:
            p_stmt = tf.paragraphs[0]

        p_stmt.text = f'"{statement}"'
        p_stmt.font.name = self.title_font
        p_stmt.font.size = Pt(36)
        p_stmt.font.bold = True
        p_stmt.font.color.rgb = self.theme["text_light"] if is_dark else self.theme["text_dark"]
        p_stmt.space_after = Pt(20)

        if context:
            p_ctx = tf.add_paragraph()
            p_ctx.text = context
            p_ctx.font.name = self.body_font
            p_ctx.font.size = Pt(16)
            p_ctx.font.color.rgb = self.theme["text_light_muted"] if is_dark else self.theme["text_muted"]

        self._add_notes(slide, notes)
        return slide

    def add_metrics_slide(self, category: str, title: str, subtitle: str, metrics: List[Dict[str, str]], notes: str = ""):
        """KPI metric callout slide with 2 to 4 metric cards."""
        slide = self.prs.slides.add_slide(self.blank_layout)
        is_dark = (self.theme_name == "dark_tech")
        self._set_slide_background(slide, self.theme["bg_dark"] if is_dark else self.theme["bg_light"])
        self._add_header(slide, category, title, subtitle, is_dark=is_dark)

        count = len(metrics)
        col_gap = 0.3
        total_width = 11.733
        card_width = (total_width - (count - 1) * col_gap) / count
        card_top = 2.2
        card_height = 4.4

        for i, m in enumerate(metrics):
            card_left = 0.8 + i * (card_width + col_gap)
            # Card background
            card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(card_left), Inches(card_top), Inches(card_width), Inches(card_height)
            )
            card.fill.solid()
            card.fill.fore_color.rgb = self.theme["card_bg"]
            card.line.color.rgb = self.theme["card_border"]
            card.line.width = Pt(1)

            # Card Content
            tx = slide.shapes.add_textbox(
                Inches(card_left + 0.3), Inches(card_top + 0.4), Inches(card_width - 0.6), Inches(card_height - 0.8)
            )
            tf = tx.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

            # Huge Number
            p_val = tf.paragraphs[0]
            p_val.text = m.get("value", "")
            p_val.font.name = self.title_font
            p_val.font.size = Pt(40)
            p_val.font.bold = True
            p_val.font.color.rgb = self.theme["primary"]
            p_val.space_after = Pt(8)

            # Trend / Badge
            if "trend" in m and m["trend"]:
                p_trend = tf.add_paragraph()
                p_trend.text = f"▲ {m['trend']}"
                p_trend.font.name = self.body_font
                p_trend.font.size = Pt(11)
                p_trend.font.bold = True
                p_trend.font.color.rgb = self.theme["secondary"]
                p_trend.space_after = Pt(12)

            # Label / Description
            p_lbl = tf.add_paragraph()
            p_lbl.text = m.get("label", "")
            p_lbl.font.name = self.title_font
            p_lbl.font.size = Pt(14)
            p_lbl.font.bold = True
            p_lbl.font.color.rgb = self.theme["text_light"] if is_dark else self.theme["text_dark"]
            p_lbl.space_after = Pt(6)

            if "detail" in m and m["detail"]:
                p_det = tf.add_paragraph()
                p_det.text = m["detail"]
                p_det.font.name = self.body_font
                p_det.font.size = Pt(11)
                p_det.font.color.rgb = self.theme["text_light_muted"] if is_dark else self.theme["text_muted"]

        self._add_notes(slide, notes)
        return slide

    def add_cards_slide(self, category: str, title: str, subtitle: str, cards: List[Dict[str, Any]], notes: str = ""):
        """Multi-column feature/pillar card slide (2, 3, or 4 columns)."""
        slide = self.prs.slides.add_slide(self.blank_layout)
        is_dark = (self.theme_name == "dark_tech")
        self._set_slide_background(slide, self.theme["bg_dark"] if is_dark else self.theme["bg_light"])
        self._add_header(slide, category, title, subtitle, is_dark=is_dark)

        count = len(cards)
        col_gap = 0.3
        total_width = 11.733
        card_width = (total_width - (count - 1) * col_gap) / count
        card_top = 2.2
        card_height = 4.4

        for i, c in enumerate(cards):
            card_left = 0.8 + i * (card_width + col_gap)

            # Container
            card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(card_left), Inches(card_top), Inches(card_width), Inches(card_height)
            )
            card.fill.solid()
            card.fill.fore_color.rgb = self.theme["card_bg"]
            card.line.color.rgb = self.theme["card_border"]
            card.line.width = Pt(1)

            # Text
            tx = slide.shapes.add_textbox(
                Inches(card_left + 0.3), Inches(card_top + 0.35), Inches(card_width - 0.6), Inches(card_height - 0.7)
            )
            tf = tx.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

            # Step / Tag
            if "tag" in c and c["tag"]:
                p_tag = tf.paragraphs[0]
                p_tag.text = c["tag"].upper()
                p_tag.font.name = self.title_font
                p_tag.font.size = Pt(10)
                p_tag.font.bold = True
                p_tag.font.color.rgb = self.theme["primary"]
                p_tag.space_after = Pt(6)
                p_hdr = tf.add_paragraph()
            else:
                p_hdr = tf.paragraphs[0]

            # Card Header
            p_hdr.text = c.get("title", "")
            p_hdr.font.name = self.title_font
            p_hdr.font.size = Pt(16)
            p_hdr.font.bold = True
            p_hdr.font.color.rgb = self.theme["text_light"] if is_dark else self.theme["text_dark"]
            p_hdr.space_after = Pt(12)

            # Points / Bullets
            points = c.get("points", [])
            for pt in points:
                p_pt = tf.add_paragraph()
                p_pt.text = f"•  {pt}"
                p_pt.font.name = self.body_font
                p_pt.font.size = Pt(11)
                p_pt.font.color.rgb = self.theme["text_light_muted"] if is_dark else self.theme["text_muted"]
                p_pt.space_after = Pt(6)

        self._add_notes(slide, notes)
        return slide

    def add_comparison_slide(self, category: str, title: str, subtitle: str,
                             left_title: str, left_points: List[str],
                             right_title: str, right_points: List[str], notes: str = ""):
        """Before vs. After or Traditional vs. Modern comparison slide."""
        slide = self.prs.slides.add_slide(self.blank_layout)
        is_dark = (self.theme_name == "dark_tech")
        self._set_slide_background(slide, self.theme["bg_dark"] if is_dark else self.theme["bg_light"])
        self._add_header(slide, category, title, subtitle, is_dark=is_dark)

        card_width = 5.6
        gap = 0.533
        card_top = 2.2
        card_height = 4.4

        # Left Column (Traditional / Problem)
        left_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.8), Inches(card_top), Inches(card_width), Inches(card_height)
        )
        left_box.fill.solid()
        left_box.fill.fore_color.rgb = self.theme["card_bg"]
        left_box.line.color.rgb = self.theme["card_border"]
        left_box.line.width = Pt(1)

        ltx = slide.shapes.add_textbox(
            Inches(1.1), Inches(card_top + 0.4), Inches(card_width - 0.6), Inches(card_height - 0.8)
        )
        ltf = ltx.text_frame
        ltf.word_wrap = True
        ltf.margin_left = ltf.margin_top = ltf.margin_right = ltf.margin_bottom = 0

        lp_head = ltf.paragraphs[0]
        lp_head.text = left_title
        lp_head.font.name = self.title_font
        lp_head.font.size = Pt(18)
        lp_head.font.bold = True
        lp_head.font.color.rgb = RGBColor(239, 68, 68) if not is_dark else RGBColor(248, 113, 113) # Subtle Red/Alert
        lp_head.space_after = Pt(14)

        for pt in left_points:
            p = ltf.add_paragraph()
            p.text = f"✕  {pt}"
            p.font.name = self.body_font
            p.font.size = Pt(12)
            p.font.color.rgb = self.theme["text_light_muted"] if is_dark else self.theme["text_muted"]
            p.space_after = Pt(8)

        # Right Column (Our Approach / Modern)
        right_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.8 + card_width + gap), Inches(card_top), Inches(card_width), Inches(card_height)
        )
        right_box.fill.solid()
        right_box.fill.fore_color.rgb = self.theme["card_bg"]
        right_box.line.color.rgb = self.theme["primary"] # Highlight border
        right_box.line.width = Pt(2)

        rtx = slide.shapes.add_textbox(
            Inches(0.8 + card_width + gap + 0.3), Inches(card_top + 0.4), Inches(card_width - 0.6), Inches(card_height - 0.8)
        )
        rtf = rtx.text_frame
        rtf.word_wrap = True
        rtf.margin_left = rtf.margin_top = rtf.margin_right = rtf.margin_bottom = 0

        rp_head = rtf.paragraphs[0]
        rp_head.text = right_title
        rp_head.font.name = self.title_font
        rp_head.font.size = Pt(18)
        rp_head.font.bold = True
        rp_head.font.color.rgb = self.theme["primary"]
        rp_head.space_after = Pt(14)

        for pt in right_points:
            p = rtf.add_paragraph()
            p.text = f"✓  {pt}"
            p.font.name = self.body_font
            p.font.size = Pt(12)
            p.font.color.rgb = self.theme["text_light"] if is_dark else self.theme["text_dark"]
            p.space_after = Pt(8)

        self._add_notes(slide, notes)
        return slide

    def add_timeline_slide(self, category: str, title: str, subtitle: str, steps: List[Dict[str, str]], notes: str = ""):
        """Horizontal milestone / roadmap progression slide."""
        slide = self.prs.slides.add_slide(self.blank_layout)
        is_dark = (self.theme_name == "dark_tech")
        self._set_slide_background(slide, self.theme["bg_dark"] if is_dark else self.theme["bg_light"])
        self._add_header(slide, category, title, subtitle, is_dark=is_dark)

        count = len(steps)
        step_width = 11.733 / count
        card_top = 2.4
        card_height = 4.0

        for i, s in enumerate(steps):
            x = 0.8 + i * step_width
            # Stage Badge
            badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(card_top), Inches(1.2), Inches(0.4))
            badge.fill.solid()
            badge.fill.fore_color.rgb = self.theme["badge_bg"]
            badge.line.fill.background()
            btf = badge.text_frame
            btf.paragraphs[0].text = s.get("phase", f"Phase {i+1}").upper()
            btf.paragraphs[0].font.size = Pt(10)
            btf.paragraphs[0].font.bold = True
            btf.paragraphs[0].font.color.rgb = self.theme["badge_text"]

            # Content
            tx = slide.shapes.add_textbox(Inches(x), Inches(card_top + 0.6), Inches(step_width - 0.3), Inches(3.2))
            tf = tx.text_frame
            tf.word_wrap = True
            tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

            p_title = tf.paragraphs[0]
            p_title.text = s.get("title", "")
            p_title.font.name = self.title_font
            p_title.font.size = Pt(15)
            p_title.font.bold = True
            p_title.font.color.rgb = self.theme["text_light"] if is_dark else self.theme["text_dark"]
            p_title.space_after = Pt(8)

            p_desc = tf.add_paragraph()
            p_desc.text = s.get("description", "")
            p_desc.font.name = self.body_font
            p_desc.font.size = Pt(11)
            p_desc.font.color.rgb = self.theme["text_light_muted"] if is_dark else self.theme["text_muted"]

        self._add_notes(slide, notes)
        return slide

    def add_chart_slide(self, category: str, title: str, subtitle: str,
                        chart_type: str, categories: List[str], series_data: Dict[str, List[float]],
                        takeaways: List[str], notes: str = ""):
        """Native PowerPoint chart slide with key takeaway card."""
        slide = self.prs.slides.add_slide(self.blank_layout)
        is_dark = (self.theme_name == "dark_tech")
        self._set_slide_background(slide, self.theme["bg_dark"] if is_dark else self.theme["bg_light"])
        self._add_header(slide, category, title, subtitle, is_dark=is_dark)

        # Build Chart Data
        cdata = CategoryChartData()
        cdata.categories = categories
        for sname, values in series_data.items():
            cdata.add_series(sname, values)

        # Chart type mapping
        ctype_map = {
            "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "bar": XL_CHART_TYPE.BAR_CLUSTERED,
            "line": XL_CHART_TYPE.LINE,
        }
        ctype = ctype_map.get(chart_type.lower(), XL_CHART_TYPE.COLUMN_CLUSTERED)

        # Insert Native Chart on Left (7.0 inches wide)
        cx, cy, cw, ch = Inches(0.8), Inches(2.2), Inches(7.5), Inches(4.5)
        chart_shape = slide.shapes.add_chart(ctype, cx, cy, cw, ch, cdata)
        chart = chart_shape.chart
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.TOP
        chart.legend.include_in_layout = False

        # Takeaways Card on Right (4.0 inches wide)
        tx_left = 8.6
        tcard = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(tx_left), Inches(2.2), Inches(3.933), Inches(4.5)
        )
        tcard.fill.solid()
        tcard.fill.fore_color.rgb = self.theme["card_bg"]
        tcard.line.color.rgb = self.theme["card_border"]

        ttx = slide.shapes.add_textbox(Inches(tx_left + 0.3), Inches(2.5), Inches(3.333), Inches(3.8))
        ttf = ttx.text_frame
        ttf.word_wrap = True
        ttf.margin_left = ttf.margin_top = ttf.margin_right = ttf.margin_bottom = 0

        tp_head = ttf.paragraphs[0]
        tp_head.text = "KEY TAKEAWAYS"
        tp_head.font.name = self.title_font
        tp_head.font.size = Pt(11)
        tp_head.font.bold = True
        tp_head.font.color.rgb = self.theme["primary"]
        tp_head.space_after = Pt(12)

        for pt in takeaways:
            p = ttf.add_paragraph()
            p.text = f"•  {pt}"
            p.font.name = self.body_font
            p.font.size = Pt(12)
            p.font.color.rgb = self.theme["text_light"] if is_dark else self.theme["text_dark"]
            p.space_after = Pt(10)

        self._add_notes(slide, notes)
        return slide

    def add_closing_slide(self, title: str, summary_points: List[str], cta: str = "", contact: str = "", notes: str = ""):
        """High-impact closing / call-to-action slide."""
        slide = self.prs.slides.add_slide(self.blank_layout)
        self._set_slide_background(slide, self.theme["bg_dark"])

        box = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(11.733), Inches(4.5))
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.name = self.title_font
        p1.font.size = Pt(36)
        p1.font.bold = True
        p1.font.color.rgb = self.theme["text_light"]
        p1.space_after = Pt(20)

        for pt in summary_points:
            p = tf.add_paragraph()
            p.text = f"✓  {pt}"
            p.font.name = self.body_font
            p.font.size = Pt(15)
            p.font.color.rgb = self.theme["text_light_muted"]
            p.space_after = Pt(10)

        if cta or contact:
            cta_box = slide.shapes.add_textbox(Inches(0.8), Inches(5.8), Inches(11.733), Inches(1.0))
            ctf = cta_box.text_frame
            cp = ctf.paragraphs[0]
            if cta:
                cp.text = cta
                cp.font.name = self.title_font
                cp.font.size = Pt(16)
                cp.font.bold = True
                cp.font.color.rgb = self.theme["primary"]
                cp.space_after = Pt(4)
            if contact:
                cp2 = ctf.add_paragraph()
                cp2.text = contact
                cp2.font.name = self.body_font
                cp2.font.size = Pt(12)
                cp2.font.color.rgb = self.theme["text_light_muted"]

        self._add_notes(slide, notes)
        return slide

    def save(self, filepath: str):
        """Saves the presentation to disk."""
        os.makedirs(os.path.dirname(os.path.abspath(filepath)), exist_ok=True)
        self.prs.save(filepath)
        print(f"[OK] Presentation successfully generated at: {filepath}")
        return filepath
