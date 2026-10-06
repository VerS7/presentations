---
name: marp-pptx
description: >-
  Converts plain Markdown files directly into native PowerPoint (.pptx) presentations using Marp CLI. Zero external APIs, 100% offline local generation via npx @marp-team/marp-cli. Supports 16:9 widescreen, custom themes (gaia, uncover, default), scoped CSS styles, speaker notes, and multi-column layouts. Trigger whenever asked to create slides from Markdown or convert markdown to PPTX.
---

# Marp PPTX: Markdown-to-PowerPoint Engine

A fast, developer-friendly way to generate native `.pptx` presentations directly from Markdown using Marp CLI.

- **Status:** Runs locally via `npx @marp-team/marp-cli` without needing any global npm installation or cloud API keys.
- **Aspect Ratio:** 16:9 Widescreen by default.

---

## 1. Quick Command: Markdown to PPTX

To compile any markdown presentation file into PowerPoint:

```powershell
npx @marp-team/marp-cli slides.md --pptx -o output.pptx
```

---

## 2. Standard Marp Markdown Header

Always start the Markdown file with these directives:

```markdown
---
marp: true
theme: gaia
_class: lead
paginate: true
backgroundColor: #0f172a
color: #f8fafc
style: |
  section {
    font-family: 'Segoe UI', system-ui, sans-serif;
    padding: 40px 60px;
  }
  h1 {
    color: #38bdf8;
  }
  h2 {
    color: #f8fafc;
  }
  .card-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 20px;
    margin-top: 30px;
  }
  .card {
    background: #1e293b;
    border: 1px solid #334155;
    border-radius: 8px;
    padding: 20px;
    font-size: 0.75em;
  }
---
```

---

## 3. Formatting Rules

1. **Slide Separators:** Use `---` on its own line between slides.
2. **Slide Classes:**
   - `<!-- _class: lead -->` centers content vertically and horizontally (ideal for Cover and Closing slides).
   - `<!-- _class: invert -->` inverts theme colors (switches between light and dark).
3. **Speaker Notes:**
   - Write notes as HTML comments at the end of the slide:
   ```markdown
   <!--
   Speaker Notes:
   - Emphasize the 42% growth in Q3.
   - Mention the pilot with enterprise beta testers.
   -->
   ```
4. **Multi-Column Layouts (Split Slides):**
   ```markdown
   <div style="display: flex; gap: 40px; margin-top: 30px;">
     <div style="flex: 1; background: #1e293b; padding: 24px; border-radius: 8px;">
       <h3>Current Friction</h3>
       <p>Manual pipelines take 14 days per release cycle.</p>
     </div>
     <div style="flex: 1; background: #0284c7; padding: 24px; border-radius: 8px; color: white;">
       <h3>Our Solution</h3>
       <p>Automated agent workflows deploy in 4 minutes.</p>
     </div>
   </div>
   ```

---

## 4. Built-in Templates in this Skill

- Inspect the working starter file: [examples/sample_deck.md](./examples/sample_deck.md).
