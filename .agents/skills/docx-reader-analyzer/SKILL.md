---
name: docx-reader-analyzer
description: >-
  Extracts text, headings, tables, outline structure, and metadata from Word documents (.docx, .dotx). Converts DOCX to clean Markdown for fast AI synthesis or presentation drafting. 100% offline via python-docx, zero external APIs. Trigger whenever asked to read, analyze, parse, summarize, or extract content from a Word or .docx file.
---

# DOCX Reader & Analyzer: Word Document Intelligence

A local skill for reading, extracting, analyzing, and transforming Microsoft Word (`.docx`, `.dotx`) documents.

- **Status:** Operational locally (`python-docx`, `lxml`).
- **Core Script:** [scripts/read_docx.py](./scripts/read_docx.py)
- **Zero External APIs:** Runs 100% offline on the local CPU with zero data transmission.

---

## 1. Quick CLI Commands

### A. Quick Structural Analysis & Outline
Inspects word count, author, metadata, and heading hierarchy:
```powershell
python .agents/skills/docx-reader-analyzer/scripts/read_docx.py document.docx
```

### B. Convert Document to Clean Markdown
Converts headings, paragraphs, lists, and tables into structured Markdown (ideal for LLM prompts or feeding into presentations):
```powershell
python .agents/skills/docx-reader-analyzer/scripts/read_docx.py document.docx --markdown > document.md
```

### C. Extract Only Tables as Markdown
Pulls out all embedded Word tables and formats them into clean GFM tables:
```powershell
python .agents/skills/docx-reader-analyzer/scripts/read_docx.py document.docx --tables
```

---

## 2. Using in Python Code

Import directly into Python scripts:

```python
import sys
sys.path.append(r".agents/skills/docx-reader-analyzer/scripts")
from read_docx import docx_to_markdown, get_docx_outline, get_docx_metadata
from docx import Document

# 1. Convert to Markdown
markdown_text = docx_to_markdown("quarterly_report.docx")

# 2. Extract Document Structure
doc = Document("quarterly_report.docx")
outline = get_docx_outline(doc)
for item in outline:
    print(f"Level {item['level']}: {item['text']}")
```

---

## 3. Workflow: DOCX Report to PowerPoint Deck

When the user gives you a Word document (memo, business plan, technical specification, or report) and wants a presentation:

1. **Step 1 (Parse):** Run `docx_to_markdown` to get the structured content and outline.
2. **Step 2 (Distill):** Use **`presentation-architect`** to extract the core problem, 3 key takeaways, and action headlines.
3. **Step 3 (Generate):** Feed the distilled sections into **`python-pptx-pro`** to generate the `.pptx` deck.
4. **Step 4 (QA):** Verify with **`pptx-qa-inspector`**.
