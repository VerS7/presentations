---
name: pdf-reader-analyzer
description: >-
  Extracts text, layout-preserved paragraphs, tables, and metadata from PDF files. Converts complex PDFs (reports, whitepapers, financial filings) into structured Markdown. 100% offline via pypdf and pdfplumber, zero external APIs. Trigger whenever asked to read, analyze, parse, extract tables, or summarize a PDF document.
---

# PDF Reader & Analyzer: Document Extraction & Intelligence

A zero-external-API skill for reading, extracting, analyzing, and transforming PDF documents into clean, structured Markdown and tabular data.

- **Status:** Installed and operational locally (`pypdf`, `pdfplumber`, `pdfminer.six`).
- **Core Script:** [scripts/read_pdf.py](./scripts/read_pdf.py)
- **Zero External APIs:** Runs 100% offline on the local CPU with zero data transmission.

---

## 1. Quick CLI Commands

### A. Quick Structural Analysis
Inspects title, author, page count, dimensions, word count estimate, and table presence:
```powershell
python .agents/skills/pdf-reader-analyzer/scripts/read_pdf.py document.pdf
```

### B. Convert Entire PDF to Structured Markdown
Extracts all pages with headings, paragraphs, and embedded tables into clean Markdown:
```powershell
python .agents/skills/pdf-reader-analyzer/scripts/read_pdf.py document.pdf --markdown > document.md
```

### C. Extract Only Tables as Markdown
Finds all financial/data tables across pages and formats them into GitHub Flavored Markdown tables:
```powershell
python .agents/skills/pdf-reader-analyzer/scripts/read_pdf.py document.pdf --tables
```

### D. Read Specific Page Range
Extracts text from targeted pages (e.g., pages 1 through 3):
```powershell
python .agents/skills/pdf-reader-analyzer/scripts/read_pdf.py document.pdf --pages 1-3
```

---

## 2. Using in Python Code

Import directly into Python scripts:

```python
import sys
sys.path.append(r".agents/skills/pdf-reader-analyzer/scripts")
from read_pdf import pdf_to_markdown, extract_pdf_tables, get_pdf_metadata

# 1. Extract metadata
meta = get_pdf_metadata("annual_report.pdf")
print(f"Total Pages: {meta['page_count']}, Dimensions: {meta['page_width_in']}x{meta['page_height_in']}")

# 2. Extract tables for data analysis
tables_md = extract_pdf_tables("annual_report.pdf")

# 3. Convert entire document to Markdown for LLM synthesis
markdown_content = pdf_to_markdown("annual_report.pdf")
```

---

## 3. Workflow: PDF Whitepaper to PowerPoint Presentation

When the user provides a PDF document (whitepaper, quarterly report, pitch deck, or academic paper) and wants a presentation:

1. **Step 1 (Parse):** Run `pdf_to_markdown` or `extract_pdf_tables` to extract raw structured text and data tables.
2. **Step 2 (Distill):** Use **`presentation-architect`** to extract the core thesis, the problem, key metrics, and roadmap.
3. **Step 3 (Generate):** Invoke **`python-pptx-pro`** to generate the native `.pptx` deck.
4. **Step 4 (QA):** Run **`pptx-qa-inspector`** to audit the final presentation.
