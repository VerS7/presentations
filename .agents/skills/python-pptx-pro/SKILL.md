---
name: python-pptx-pro
description: >-
  Generates production-grade, 16:9 widescreen PowerPoint (.pptx) presentations locally using Python and python-pptx. Zero external APIs, zero paid subscriptions. Includes modular slide archetypes (title, KPI metrics, pillar cards, before/after comparisons, milestone roadmaps, native charts, closing CTA) and curated professional color palettes. Trigger whenever asked to generate, build, or code a PowerPoint presentation locally.
---

# Python PPTX Pro: Local Native PowerPoint Generator

A zero-external-API engine for generating pixel-perfect PowerPoint (`.pptx`) decks locally using Python.

- **Status:** Installed and operational locally (`python-pptx`, `pillow`, `matplotlib`).
- **Core Library:** [scripts/pptx_engine.py](./scripts/pptx_engine.py)
- **Aspect Ratio:** 16:9 Widescreen standard (`13.333" × 7.5"`).

---

## 1. Quick Start: Generating a Presentation in 10 Lines

Any agent can write and execute a Python script to create a presentation instantly:

```python
import sys
sys.path.append(r".agents/skills/python-pptx-pro/scripts")
from pptx_engine import DeckBuilder

# Initialize with a curated theme: 'corporate', 'dark_tech', or 'vibrant'
deck = DeckBuilder(theme="corporate")

# 1. Title Slide
deck.add_title_slide(
    title="Next-Generation Autonomous Systems",
    subtitle="Scaling Enterprise Performance with Local Agent Workflows",
    author="Albert • Palych Point",
    date="October 2026"
)

# 2. KPI Metrics Slide
deck.add_metrics_slide(
    category="Performance Impact",
    title="Core Engineering Benchmarks Improved by 3.8x",
    subtitle="Comparing Q3 results against legacy development pipelines",
    metrics=[
        {"value": "3.8x", "trend": "+280%", "label": "Release Velocity", "detail": "Average sprint delivery time decreased from 14 days to 3.7 days."},
        {"value": "99.98%", "trend": "0 incidents", "label": "System Reliability", "detail": "Zero critical production outages during automated deployment."},
        {"value": "$420K", "trend": "-64% cost", "label": "Infrastructure Savings", "detail": "Annual cloud computation and API subscription savings."}
    ]
)

# 3. Save to PPTX
deck.save("output/enterprise_presentation.pptx")
```

---

## 2. Available Slide Archetypes in `pptx_engine.py`

| Method | Best For | Visual Layout |
|---|---|---|
| `add_title_slide(...)` | Deck cover & opener | Dark high-contrast background, bold accent rule, presenter metadata |
| `add_statement_slide(...)` | Core thesis, pivot, or big quote | Full slide focus, 36pt headline, supporting context |
| `add_metrics_slide(...)` | Traction, benchmarks, ROI | 2 to 4 rounded card containers, 40pt bold stats, trend badges |
| `add_cards_slide(...)` | Features, pillars, key findings | 2, 3, or 4 column container cards with headers and bullet lists |
| `add_comparison_slide(...)` | Before vs. After, Us vs. Them | 2 distinct columns (Alert border on left, Primary border on right) |
| `add_timeline_slide(...)` | Roadmaps, phases, project stages | Horizontal milestone steppers with stage badges and descriptions |
| `add_chart_slide(...)` | Data visualization | Native PowerPoint chart (column/bar/line) + takeaway callout card |
| `add_closing_slide(...)` | Executive summary & CTA | Recap bullets, contact metadata, bold call-to-action |

---

## 3. Curated Color Palettes

Switch the palette when instantiating `DeckBuilder(theme=...)`:

1. **`corporate` (Corporate Modern - Default):**
   - Deep Slate `#0F172A`, Sky Blue `#0284C7`, Crisp Slate `#64748B`, Slate-50 background.
   - Ideal for B2B decks, internal strategy, corporate boards, and consulting.
2. **`dark_tech` (Dark Tech / AI):**
   - Obsidian `#0B0F19`, Neon Emerald `#10B981`, Slate `#94A3B8`, Pure White.
   - Ideal for AI, cybersecurity, developer tools, live keynote stages, and tech pitches.
3. **`vibrant` (Vibrant Pitch):**
   - Midnight Violet `#181124`, Indigo `#6366F1`, Sunset Rose `#F43F5E`.
   - Ideal for venture capital pitch decks, creative agencies, and modern SaaS.

---

## 4. Low-Level `python-pptx` Rules & Best Practices

If writing custom low-level `python-pptx` code outside the engine:

1. **Always Set 16:9 Widescreen:**
   ```python
   prs.slide_width = Inches(13.333)
   prs.slide_height = Inches(7.5)
   ```
2. **Always Enable Word Wrap & Reset Margins:**
   ```python
   tf = shape.text_frame
   tf.word_wrap = True
   tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
   ```
3. **Spacing Bullet Points:**
   - Never use `line_spacing` with large numbers (causes unnatural line gaps within the same bullet).
   - Use paragraph spacing: `p.space_after = Pt(8)`.
4. **Speaker Notes:**
   - Always attach speaker notes to the notes slide:
   ```python
   slide.notes_slide.notes_text_frame.text = "Key message: ..."
   ```
5. **Native Charts vs. Images:**
   - Prefer native `slide.shapes.add_chart()` so that numbers and series remain fully editable in PowerPoint.
