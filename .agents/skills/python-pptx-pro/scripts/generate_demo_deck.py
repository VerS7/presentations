"""
generate_demo_deck.py
Builds a complete, 8-slide high-impact presentation demonstrating all archetypes.
Runs completely offline with zero external APIs.
"""

import os
import sys

# Ensure python finds pptx_engine in the same directory
curr_dir = os.path.dirname(os.path.abspath(__file__))
if curr_dir not in sys.path:
    sys.path.insert(0, curr_dir)

from pptx_engine import DeckBuilder

def build_presentation(output_file: str = "demo_presentation.pptx", theme: str = "corporate"):
    deck = DeckBuilder(theme=theme)

    # Slide 1: Cover / Title
    deck.add_title_slide(
        title="AI-Driven Presentation Framework",
        subtitle="Autonomous, High-Impact PowerPoint Generation Without External APIs",
        author="Palych Point • Engineering Team",
        date="October 2026",
        notes="Welcome everyone. Today we are showcasing how local agent skills enable rapid, high-quality presentation generation without relying on any external paid APIs."
    )

    # Slide 2: Core Statement / Shift
    deck.add_statement_slide(
        category="Strategic Paradigm",
        statement="Slides should deliver insights in 3 seconds, not require deciphering paragraphs.",
        context="Modern decision-makers scan rather than read. Visual hierarchy, action headlines, and modular cards replace wall-of-text bullets.",
        notes="Emphasize that the biggest issue with AI-generated presentations has historically been cluttered text. Our framework solves this by enforcing strict card layouts and font hierarchies."
    )

    # Slide 3: KPI Metrics / Traction
    deck.add_metrics_slide(
        category="Performance Benchmarks",
        title="Quantifiable Efficiency Gains Across the Pipeline",
        subtitle="Comparing automated skill generation against manual slide crafting",
        metrics=[
            {
                "value": "10x",
                "trend": "+900%",
                "label": "Creation Velocity",
                "detail": "Complete 10-slide decks generated from briefs in under 30 seconds."
            },
            {
                "value": "100%",
                "trend": "Offline Safe",
                "label": "Data Privacy",
                "detail": "Runs entirely in the local environment with zero third-party telemetry."
            },
            {
                "value": "16:9",
                "trend": "HD Standard",
                "label": "Widescreen Native",
                "detail": "Pixel-perfect alignment optimized for modern enterprise displays."
            }
        ],
        notes="Walk through the three metrics. Point out that the 100% offline aspect is crucial for sensitive corporate and internal financial presentations."
    )

    # Slide 4: 3-Pillar Card Grid
    deck.add_cards_slide(
        category="Architecture",
        title="Three Pillars of Modern Deck Construction",
        subtitle="Every slide must fulfill a concrete communicative function",
        cards=[
            {
                "tag": "Pillar 01",
                "title": "Story Spine",
                "points": [
                    "SCQA framework anchors context",
                    "Action titles state the thesis up front",
                    "Ending defined before the middle"
                ]
            },
            {
                "tag": "Pillar 02",
                "title": "Design Discipline",
                "points": [
                    "Pre-tested 4-color palettes",
                    "WCAG compliant contrast ratios",
                    "Container cards with subtle borders"
                ]
            },
            {
                "tag": "Pillar 03",
                "title": "Native OOXML",
                "points": [
                    "Fully editable in Microsoft PowerPoint",
                    "Native chart objects (no static images)",
                    "Clean vector shapes and typography"
                ]
            }
        ],
        notes="Highlight that each card acts as a standalone unit of thought. The audience can scan the three tags immediately."
    )

    # Slide 5: Before vs After Comparison
    deck.add_comparison_slide(
        category="Evolution",
        title="Legacy Slide Pitfalls vs. Modern Deck Standards",
        subtitle="Moving from unstructured bullets to scannable executive modules",
        left_title="Traditional Flaws (Legacy)",
        left_points=[
            "Wall of text paragraphs that require reading while speaking",
            "Passive topic titles like 'Financial Overview' or 'Architecture'",
            "Uncalibrated colors causing low contrast on conference projectors",
            "Distorted 4:3 aspect ratios stretched on modern widescreen monitors"
        ],
        right_title="Modern Standard (Our Skills)",
        right_points=[
            "Scannable modular cards with 3-second comprehension",
            "Action titles making bold, quantifiable assertions",
            "High-contrast palettes rigorously tested for readability",
            "Native 16:9 widescreen layout with standardized margins"
        ],
        notes="Contrast the left and right sides. Remind the room of painful meetings where someone read 10 bullet points off a slide."
    )

    # Slide 6: Native Chart & Data Takeaways
    deck.add_chart_slide(
        category="Execution Metrics",
        title="Adoption Growth Across Project Cycles",
        subtitle="Quarterly expansion of automated presentation workflows",
        chart_type="column",
        categories=["Q1", "Q2", "Q3", "Q4"],
        series_data={
            "Manual Presentations": [85.0, 60.0, 35.0, 15.0],
            "AI-Skill Powered": [15.0, 40.0, 65.0, 85.0]
        },
        takeaways=[
            "Skill-based workflows overtook manual drafting in Q3.",
            "Slide formatting revisions dropped by 78%.",
            "Cross-team presentation consistency reached 94%."
        ],
        notes="Explain that this chart is a native PowerPoint chart element, meaning anyone opening this file can edit the numbers directly in Excel/PowerPoint."
    )

    # Slide 7: Roadmap / Timeline
    deck.add_timeline_slide(
        category="Roadmap",
        title="Implementation Milestones for Enterprise Rollout",
        subtitle="Four phases to transition from manual to automated presentations",
        steps=[
            {
                "phase": "Phase 1",
                "title": "Skills Setup",
                "description": "Mount .agents/skills in the workspace and configure local python-pptx."
            },
            {
                "phase": "Phase 2",
                "title": "Template Sync",
                "description": "Define corporate color schemes, typography, and standard layouts."
            },
            {
                "phase": "Phase 3",
                "title": "Agent Integration",
                "description": "Empower AI assistants to draft and critique slides autonomously."
            },
            {
                "phase": "Phase 4",
                "title": "Full Scale",
                "description": "Generate executive decks, client pitches, and sprints in seconds."
            }
        ],
        notes="Walk through the 4 phases. Phase 1 is already complete today."
    )

    # Slide 8: Summary & Call to Action
    deck.add_closing_slide(
        title="Ready to Build Better Presentations",
        summary_points=[
            "All skills are registered and permanently available in this project.",
            "Zero external API dependencies: 100% private and locally executable.",
            "Seamless integration between narrative planning and technical rendering."
        ],
        cta="Generate your next deck using python-pptx-pro or marp-pptx.",
        contact="Project: palych-point • Automated Workspace Skill Suite",
        notes="Close with a strong call to action and invite questions."
    )

    deck.save(output_file)

if __name__ == "__main__":
    out = "demo_presentation.pptx"
    if len(sys.argv) > 1:
        out = sys.argv[1]
    build_presentation(out)
