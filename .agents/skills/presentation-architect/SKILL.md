---
name: presentation-architect
description: >-
  Master skill for structuring high-impact presentations and slide decks. Guides story spine, narrative arc (SCQA, Pyramid Principle), slide-by-slide sequence, action headlines, copy density, visual hierarchy, and speaker notes. Trigger whenever asked to create a presentation, plan slides, turn a document into a deck, or refine slide storytelling.
---

# Presentation Architect: Storytelling, Structure & Slide Copy

Builds memorable, professional slide decks with a narrative spine underneath. Focuses on clarity, audience empathy, high contrast, and disciplined slide density before any design code or file generation begins.

- **Use when:** Planning a new presentation, transforming a document/brief into slides, structuring an executive pitch, defining slide pacing, or drafting speaker notes.
- **Pairs with:** 
  - `python-pptx-pro` for generating native PowerPoint `.pptx` files locally.
  - `presentation-pitch-deck` for investor-specific deck frameworks (Sequoia/YC).
  - `marp-pptx` for generating decks via Markdown.
  - `pptx-qa-inspector` for pre-flight quality checks.

---

## The 7-Step Architecture Workflow

Follow this sequence to ensure the presentation achieves its business objective:

```text
Presentation Progress:
- [ ] Step 1: Context & Constraints (Audience, setting, core takeaway)
- [ ] Step 2: Story Spine & Climax (references/story-structure.md)
- [ ] Step 3: Slide Sequence & Pacing (references/outline-structure.md)
- [ ] Step 4: Write Slide Content & Action Headlines (references/writing-slides.md)
- [ ] Step 5: Visual Layout & Density Mapping
- [ ] Step 6: Speaker Notes & Delivery Prompts (references/speaker-notes.md)
- [ ] Step 7: Technical Generation Handoff (`python-pptx-pro` or `marp-pptx`)
```

---

## Step 1: Context & Constraints

Before writing a single slide, establish:

1. **Audience Persona:**
   - *Executive / C-Suite:* Wants conclusions first (Pyramid Principle), financial impact, clear recommendations.
   - *Technical / Engineering:* Wants architecture, trade-offs, mechanics, benchmark data.
   - *External / Client:* Needs problem validation, trust signals, clear ROI, frictionless next steps.
2. **Setting & Delivery Mode:**
   - *Live Keynote / Presentation:* Low text density (10–20 words per slide), large typography, visual anchor. Speaker carries the detail.
   - *Async / Sent Deck ("Slidedoc"):* Higher text density (30–60 words per slide), self-explanatory cards, clear hierarchy. Must be understandable without a presenter.
3. **The Core Takeaway (The "Rule of One"):**
   - If the audience remembers only **one sentence** tomorrow morning, what must that sentence be?
4. **Three Supporting Pillars:**
   - Every slide in the deck must support one of these three pillars. Anything else is noise.

---

## Step 2: Story Spine (Pixar / Kenn Adams Framework)

A deck is not a list of facts; it is a movement from a problem to a resolved future.

Draft this 6-line spine before organizing slides:

```text
1. Once upon a time, [The established status quo].
2. Every day, [How the world operated, accepted inefficiencies].
3. One day, [The trigger event, market shift, or new challenge].
4. Because of that, [The initial consequence and mounting tension].
5. Because of that, [The critical bottleneck that traditional methods cannot solve].
6. Until finally, [Our solution / breakthrough that creates a new reality].
```

> **Key Rule: Write the ending first.** Define the final decision, recommendation, or call-to-action slide before writing the middle.

*(See full guide in [references/story-structure.md](references/story-structure.md))*

---

## Step 3: Slide Sequence & Pacing

A standard executive/business presentation follows this 4-act structure:

| Act | Purpose | Typical Slide Count | Archetypes Used |
|---|---|---|---|
| **Act I: The Hook & Context** | Establish relevance and current tension | 1–3 slides | Title, Problem Framing, The Burning Platform |
| **Act II: The Shift & Vision** | Introduce the core insight or strategy | 2–4 slides | Core Statement, Solution Architecture, Value Proposition |
| **Act III: Proof & Evidence** | Prove feasibility, data, and mechanics | 3–6 slides | Metric/KPI Cards, Comparison Matrix, Case Study, Roadmap |
| **Act IV: The Ask & Close** | Commit to action and clear next steps | 1–2 slides | Recommendation, Next Milestones, Q&A / Contacts |

*(See archetype catalog in [references/outline-structure.md](references/outline-structure.md))*

---

## Step 4: Write Slide Copy (Action Headlines)

### The Action Title Rule
Never write passive topic labels. A slide title must be an **assertion** that delivers the takeaway even if the reader reads nothing else.

| ❌ Passive Topic Label (Avoid) | ✅ Action Headline (Use) |
|---|---|
| Market Overview | AI adoption in enterprise grew 140% year-over-year |
| Financial Results | Q3 revenue hit \$12M, driven by 85% net retention |
| Architecture | Decoupled microservices reduced deployment time from days to minutes |
| Our Competition | We offer 3x faster setup at 50% lower infrastructure cost |

### Body Copy Rules
- **Bold Lead-in:** Start bullet points or card paragraphs with 2–3 bold words followed by concise explanation:
  - `**Zero-Trust Security:** Built-in hardware-level encryption with no third-party telemetry.`
- **Rule of Three:** Group metrics, benefits, or features into 3 items (or 4 max).
- **Chunking:** Break long paragraphs into scannable cards or metric blocks.

*(See full patterns in [references/writing-slides.md](references/writing-slides.md))*

---

## Step 5: Visual Mapping & Layout Archetypes

Match each slide's content to a tested PowerPoint layout component:

1. **Big Statement / Hero:** Full-bleed color, 40–54pt type, single sentence for maximum emotional weight.
2. **KPI / Metric Callout:** 2–4 large number callouts (44pt+ bold) with small descriptive labels and trend tags (`+42% MoM`).
3. **Card Grid (2, 3, or 4 columns):** Clean container boxes with subtle borders, title, and 2–3 bullet points.
4. **Comparison / Matrix (Before vs. After or Us vs. Others):** 2 distinct columns or table comparing key dimensions.
5. **Horizontal Stepper / Timeline:** 3–5 milestone stages connected sequentially from left to right.
6. **Data / Chart Slide:** Clean native chart (Bar, Line, or Donut) with a bold takeaway headline and callout bullet.

---

## Step 6: Speaker Notes

For presentations delivered live, generate structured speaker notes for each slide:
- **Key Message:** The 1 sentence the speaker must communicate.
- **Opening Line:** Natural verbal hook.
- **Talking Points:** 2–3 conversational bullet points with anecdotes or deep background.
- **Transition Bridge:** Seamless verbal link to the next slide.

*(See template in [references/speaker-notes.md](references/speaker-notes.md))*

---

## Step 7: Handoff to Generator

Once the outline and slide copy are approved:
- Use **`python-pptx-pro`** to generate editable native PowerPoint (`.pptx`) with Python.
- Or use **`marp-pptx`** to render directly from Markdown.
- Run **`pptx-qa-inspector`** to verify font sizes, margins, and contrast ratios.
