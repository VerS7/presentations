# Writing Slides: Headlines, Microcopy & Visual Hierarchy

Slide text is not prose; it is microcopy designed to be absorbed in 3 seconds.

---

## 1. Headline Formulas

A good headline does the heavy lifting. It allows an executive to read only the slide titles and understand the entire presentation.

### Formula A: Assertion + Quantified Metric
- ❌ "Q3 Sales Performance"
- ✅ "Q3 Enterprise Revenue Grew 42% Through Expansion Deals"

### Formula B: Problem + Impact
- ❌ "Legacy Infrastructure Challenges"
- ✅ "Manual Releases Cause 14 Hours of Downtime Every Month"

### Formula C: Contrast / Paradigm Shift
- ❌ "Architecture Approach"
- ✅ "Decoupled Event Pipelines Replace Fragile Monolithic Syncs"

### Formula D: Provocative Question (Used for tension slides)
- ❌ "Next Steps"
- ✅ "What Happens If We Delay Migration By Another Quarter?"

---

## 2. Microcopy Patterns

### The Bold Lead-in Pattern
Never write plain unformatted bullets. The first 2–4 words should summarize the point in bold, followed by a clarifying explanation:

```text
- **Zero-Latency Ingestion:** Processes 50,000 events/sec with sub-5ms p99 latency.
- **Enterprise-Grade RBAC:** Granular role-based controls with automatic SSO audit logs.
- **Self-Healing Failover:** Recovers downed workers in under 800ms without packet loss.
```

### The Rule of Three (and Max Four)
- Humans can hold 3–4 items in working memory.
- If you have 7 items, categorize them into 2 groups or select the top 3 and put the rest in an appendix.

### Number Formatting
- Make numbers huge and self-explanatory.
- Format with units (`$14.2M`, `99.98%`, `3.4x`, `12ms`).
- Always pair a large number with a 1-sentence descriptor underneath explaining what the number means.

---

## 3. Density Rules by Delivery Mode

| Delivery Mode | Words per Slide | Font Size Floor | Structure |
|---|---|---|---|
| **Live Keynote / Stage** | 10–25 words | 28pt body, 44pt title | Massive focal visual, punchy phrase |
| **Meeting / Board Presentation** | 30–60 words | 18pt body, 32pt title | Cards, metric boxes, 3 bullets max |
| **Async / Sent "Slidedoc"** | 60–120 words | 14pt body, 24pt title | Modular card grids, clear subtitles, footnotes |
