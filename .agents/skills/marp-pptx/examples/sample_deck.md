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
    font-size: 2.2em;
  }
  h2 {
    color: #ffffff;
    font-size: 1.5em;
  }
  .grid-3 {
    display: flex;
    gap: 20px;
    margin-top: 30px;
  }
  .card {
    flex: 1;
    background: #1e293b;
    border: 1px solid #334155;
    border-radius: 12px;
    padding: 24px;
    font-size: 0.7em;
  }
  .stat-val {
    font-size: 2.2em;
    font-weight: bold;
    color: #38bdf8;
    margin-bottom: 8px;
  }
---

# Autonomous Deck Systems
### Transforming Markdown to PowerPoint Locally
**Palych Point • Engineering**

<!--
Speaker Notes:
Welcome to the deck. This was built directly from Markdown and exported to native PPTX.
-->

---

<!-- _class: default -->

## Core Engineering Performance

<div class="grid-3">
  <div class="card">
    <div class="stat-val">10x</div>
    <strong>Delivery Velocity</strong>
    <p>Slide generation time dropped from hours to seconds.</p>
  </div>
  <div class="card">
    <div class="stat-val">100%</div>
    <strong>Local Privacy</strong>
    <p>Zero external API calls. Everything runs on the local CPU.</p>
  </div>
  <div class="card">
    <div class="stat-val">16:9</div>
    <strong>Native Resolution</strong>
    <p>Rendered cleanly for widescreen conference room monitors.</p>
  </div>
</div>

<!--
Speaker Notes:
Walk through the 3 core metrics. Note that privacy and speed are our top architectural requirements.
-->

---

<!-- _class: lead -->

# Ready to Build With Marp
### Run: `npx @marp-team/marp-cli sample_deck.md --pptx`
