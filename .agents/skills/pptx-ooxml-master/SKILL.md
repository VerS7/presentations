---
name: pptx-ooxml-master
description: >-
  Advanced skill for direct PowerPoint (.pptx) manipulation, OOXML architecture, template duplication, and pptxgenjs engine rules. Use when unpacking/editing ppt/slides/slideN.xml directly, duplicating slide layouts, re-zipping presentation archives, or troubleshooting PPTX file corruption and schema defects.
---

# PPTX OOXML Master: Architecture, Templates & Low-Level Manipulation

A PowerPoint `.pptx` file is a ZIP archive adhering to the Open Packaging Conventions (OPC) containing XML files and media parts.

---

## 1. Directory Structure of a `.pptx` File

```text
presentation.pptx (ZIP Archive)
├── [Content_Types].xml               # MIME types for all parts in the package
├── _rels/
│   └── .rels                         # Package root relationships
├── docProps/
│   ├── app.xml                       # Application metadata (slide count, app name)
│   └── core.xml                      # Dublin Core metadata (author, title, modified)
└── ppt/
    ├── presentation.xml              # Master presentation manifest (<p:sldIdLst>)
    ├── _rels/
    │   └── presentation.xml.rels     # Maps rIds to slides, notesMaster, themes
    ├── slideLayouts/                 # Slide layout masters (slideLayout1.xml, ...)
    ├── slideMasters/                 # Master slide themes
    ├── slides/                       # The actual slide contents
    │   ├── slide1.xml                # Slide 1 text, shapes, coordinates (EMU)
    │   ├── slide2.xml
    │   └── _rels/
    │       ├── slide1.xml.rels       # Slide 1 relationships (layout, media, charts)
    │       └── slide2.xml.rels
    └── media/                        # Embedded PNGs, JPGs, SVGs
```

---

## 2. Direct Editing Workflow (Unpack → Edit → Repack)

To modify an existing presentation or template without PowerPoint:

```powershell
# 1. Unpack the .pptx archive
python -c "import zipfile; zipfile.ZipFile('deck.pptx').extractall('unpacked')"

# 2. Duplicate a slide layout or slide (using local helper)
python .agents/skills/pptx-ooxml-master/scripts/manage_slide.py duplicate unpacked/ slide2.xml --after slide2.xml

# 3. Edit slide content in unpacked/ppt/slides/slideN.xml
# Update <a:t> text nodes, colors (<a:srgbClr val="..."/>), or shape positions

# 4. Repack cleanly from INSIDE the unpacked directory
python -c "import shutil, zipfile, os; z = zipfile.ZipFile('out.pptx', 'w', zipfile.ZIP_DEFLATED); [z.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), 'unpacked')) for root, dirs, files in os.walk('unpacked') for file in files]; z.close()"
```

> **Critical OOXML Rules:**
> - When re-zipping, zip from **inside** the directory so `[Content_Types].xml` is at the root of the ZIP, not nested inside an `unpacked/` subfolder.
> - Never reorder `<p:sldIdLst>` children without updating `presentation.xml.rels`.
> - Always update `[Content_Types].xml` whenever adding a new `slideN.xml`.

---

## 3. `pptxgenjs` Engine Rules & Footguns

If generating slides using Node.js and `pptxgenjs`:

1. **Hex Colors: NEVER include `#` or 8 digits:**
   - ❌ `color: "#FF0000"` or `color: "00000020"` (CORRUPTS THE FILE)
   - ✅ `color: "FF0000"` (6-digit hex string without `#`)
2. **Set Layout Before Adding Slides:**
   - Default canvas in pptxgenjs is `LAYOUT_16x9` = **10" × 5.625"**.
   - For standard 16:9 HD widescreen, set `pres.layout = "LAYOUT_WIDE"` (13.3" × 7.5").
3. **Shadow Offsets Must Be ≥ 0:**
   - A negative `offset` corrupts the XML schema. For upward shadows, use `angle: 270` with positive offset.
4. **Lists and Bullets:**
   - Pass `bullet: true` on each text item. Never insert literal `•` characters (causes duplicate bullets).
5. **Chart Faults:**
   - On stacked bar/column charts, `dataLabelPosition` must only be `'ctr'`, `'inEnd'`, or `'inBase'`. Using `'outEnd'` corrupts the file.
   - Combo charts with secondary axes must explicitly declare both `valAxes` and `catAxes` with 2 entries each.
