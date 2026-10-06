---
name: pptx-qa-inspector
description: >-
  Audits, validates, and quality-checks PowerPoint (.pptx) presentations before delivery. Inspects aspect ratios (16:9), slide density, action headline compliance, shape counts, speaker notes coverage, and font size floors. Trigger after generating or editing any PowerPoint deck to perform a pre-flight quality gate.
---

# PPTX QA Inspector: Pre-Flight Presentation Audit

A quality gate skill to verify that generated PowerPoint files look executive-ready, maintain consistent visual rhythm, and contain no structural defects.

- **Audit Script:** [scripts/audit_presentation.py](./scripts/audit_presentation.py)
- **Zero API Dependencies:** Uses pure Python and `python-pptx`.

---

## 1. Quick Audit Command

Run the audit script against any generated presentation:

```powershell
python .agents/skills/pptx-qa-inspector/scripts/audit_presentation.py my_deck.pptx
```

The script outputs:
- **Aspect Ratio Validation:** Verifies 16:9 widescreen (`13.333" × 7.5"`).
- **Density per Slide:** Identifies slides with over 90 words that risk becoming walls of text.
- **Speaker Notes Coverage:** Checks what % of slides have complete presenter talk tracks.
- **Empty Slide Detection:** Catches accidental blank slides or orphaned layouts.

---

## 2. The 6-Point Quality Assurance Rubric

Before presenting or delivering a deck to the user, verify these 6 checks:

| # | Quality Check | Standard | How to Verify |
|---|---|---|---|
| 1 | **Aspect Ratio** | 16:9 Widescreen (never 4:3 legacy) | Audit script check or PowerPoint slide size setting |
| 2 | **Action Headlines** | Every slide has an assertion title | Check titles: no passive single-word labels like "Finance" |
| 3 | **Density Floor** | Max 60–80 words per card slide | No paragraphs longer than 3 lines; bold lead-in on bullets |
| 4 | **Contrast Ratio** | WCAG 4.5:1 floor (7:1 for dark rooms) | White text on dark navy/obsidian, or dark slate on paper |
| 5 | **Font Discipline** | Max 2 font families, max 3 font sizes | Title: 24–40pt, Body: 11–14pt, Footnotes: 9–10pt |
| 6 | **Speaker Notes** | Presenter tracks attached | Every slide has a key message, opening hook, and verbal cues |
