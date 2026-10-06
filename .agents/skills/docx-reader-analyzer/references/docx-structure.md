# DOCX Architecture & Parsing Reference

A `.docx` file is a ZIP container containing XML files conforming to the Office Open XML (OOXML) standard.

---

## 1. Internal Package Components

```text
document.docx (ZIP archive)
├── [Content_Types].xml
├── _rels/
├── docProps/
│   ├── app.xml           # Total words, paragraphs, lines, template
│   └── core.xml          # Author, created date, title, revision number
└── word/
    ├── document.xml      # The body of the document: paragraphs, tables, runs
    ├── styles.xml        # Hierarchy of heading and character styles
    ├── numbering.xml     # Bulleted and numbered list definitions
    ├── header1.xml       # Headers
    ├── footer1.xml       # Footers
    └── media/            # Embedded PNG, JPG images
```

---

## 2. Key OOXML Elements in `word/document.xml`

- `<w:p>`: Paragraph container.
- `<w:pPr>`: Paragraph properties (style name, alignment, indentation).
- `<w:r>`: Run container (chunk of text with uniform formatting).
- `<w:rPr>`: Run properties (bold `<w:b/>`, italic `<w:i/>`, font color, font size).
- `<w:t>`: Actual text content node.
- `<w:tbl>`: Table container.
- `<w:tr>`: Table row.
- `<w:tc>`: Table cell.
