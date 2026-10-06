# PDF Extraction Strategies & Best Practices

PDF (Portable Document Format) is designed for visual fidelity on printers and screens, not semantic hierarchy. This reference outlines strategies for accurate text and table extraction.

---

## 1. Extraction Strategies: `pypdf` vs `pdfplumber`

| Task | Recommended Library | Rationale |
|---|---|---|
| **Fast text dump / Bookmarks** | `pypdf` | Minimal CPU overhead, extracts raw string streams. |
| **Complex layout preservation** | `pdfplumber` | Preserves character bounding boxes and vertical/horizontal alignment. |
| **Table extraction** | `pdfplumber` | Employs visual line detection (`explicit_horizontal_lines`) to build structured grids. |
| **Page splitting / Merging** | `pypdf` | Native page stream manipulation without re-rendering. |

---

## 2. Handling Complex PDF Layouts

- **Multi-Column Text:** In two-column documents (like academic papers or annual reports), raw string extraction can read left-to-right across columns, jumbling sentences. `pdfplumber.extract_text(layout=True)` groups text by x/y coordinates to respect column gutters.
- **Header / Footer Stripping:** When parsing long documents, ignore text with `top < 50` or `bottom > 750` pt to remove repeating page numbers, running headers, and legal disclaimers.
- **Embedded Tables:** Always use `extract_tables()` before extracting text, so table numbers don't get flattened into unreadable paragraphs.
