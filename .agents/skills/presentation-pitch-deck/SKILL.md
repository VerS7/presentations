---
name: presentation-pitch-deck
description: >-
  Specialized skill for creating investor, startup, and business pitch decks. Enforces proven venture capital frameworks (Sequoia Capital, Y Combinator, Guy Kawasaki), market sizing (TAM/SAM/SOM), unit economics, competitive moats, and ask structure. Trigger whenever asked for an investor pitch, startup deck, seed/Series A presentation, or executive business proposal.
---

# Presentation Pitch Deck: Investor & Business Proposals

Guides the creation of pitch decks designed to secure investment, executive budget, or strategic partnership.

- **Primary Frameworks:** Sequoia Capital Business Plan Outline, Y Combinator Seed Deck Formula, Guy Kawasaki 10/20/30.
- **Key Differentiation:** Distinct rules for **Sent Decks** (read asynchronously by investors on laptops) vs. **Presented Decks** (spoken on a stage or Zoom).

---

## 1. Sent Deck vs. Presented Deck

Investors spend an average of **2 minutes and 40 seconds** reviewing a pitch deck before deciding whether to take a meeting.

| Property | Sent Deck (Read Async) | Presented Deck (Live Pitch) |
|---|---|---|
| **Primary Reader** | Associate or Partner skimming at a desk | Audience in a room or Zoom call |
| **Slide Density** | Standalone: each card has complete context | Minimal: 1 focal point, presenter speaks the context |
| **Reading Time** | 20–40 seconds per slide | 3–5 seconds glance per slide |
| **Word Count** | 40–80 words per slide | 10–25 words per slide |
| **Strategy** | Build the sent deck first; trim it down for live delivery | Cut copy into speaker notes |

---

## 2. The Canonical 10-Slide Pitch Framework

Follow this sequence strictly. In pitch decks, deviating from standard structure makes investors suspect you are hiding a weakness (e.g. putting the Problem on slide 6 or omitting Unit Economics).

```text
Slide 01: Company Purpose (One declarative sentence)
Slide 02: Problem / Pain (Urgent, painful, costly problem)
Slide 03: Solution & Secret Sauce (How you solve it 10x better)
Slide 04: Why Now (The macroeconomic or tech catalyst)
Slide 05: Market Sizing (TAM, SAM, SOM calculations)
Slide 06: Product Tour / Architecture (Proof of execution)
Slide 07: Traction & Unit Economics (Key numbers and growth rate)
Slide 08: Business Model & Pricing (How you make money)
Slide 09: Competition & Defensible Moat (Why you win long-term)
Slide 10: Team & The Ask (Capital requested, runway, milestones)
```

*(See detailed slide-by-slide guides in [references/pitch-framework.md](references/pitch-framework.md) and [references/financials-and-metrics.md](references/financials-and-metrics.md))*

---

## 3. Critical Investor Heuristics

1. **Lead with Numbers on Traction:**
   - If revenue, user count, or retention is impressive, pull Traction earlier (Slide 4 or 5). Strong numbers earn immediate credibility.
2. **"Why Now" is the Most Scrutinized Slide:**
   - Why couldn't this exist 3 years ago? (e.g. LLM reasoning breakthroughs, new regulations, commoditization of GPU cloud).
   - A great idea without a "Why Now" catalyst reads as a project someone already failed at.
3. **The 2x2 Matrix Warning:**
   - Avoid generic 2x2 competitive matrices where your company sits comfortably in the top-right quadrant with two arbitrary axes ("Easy to use" vs "Powerful").
   - Instead, use a **Feature & Dimension Grid** showing 4–5 specific structural capabilities that incumbents cannot easily copy without rebuilding their stack.
4. **The Ask Must Be Milestone-Driven:**
   - Never say: "Raising \$2M for general hiring and marketing."
   - Say: "Raising \$2.5M to reach \$150K MRR and expand enterprise pilots over an 18-month runway."

---

## 4. Implementation with `python-pptx-pro`

Use the `vibrant` or `dark_tech` theme for pitch decks:
- Use `add_metrics_slide()` for Slide 7 (Traction & Unit Economics).
- Use `add_comparison_slide()` for Slide 9 (Competition & Moat).
- Use `add_timeline_slide()` for Slide 10 (Post-funding milestones).
