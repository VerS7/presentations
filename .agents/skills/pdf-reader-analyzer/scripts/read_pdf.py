"""
read_pdf.py - Professional Local PDF Reader & Analyzer
Extracts text, tables, metadata, and converts PDF documents to Markdown.
Runs 100% offline using pypdf and pdfplumber. Zero external APIs.
"""

import os
import sys
import argparse
from typing import List, Dict, Any, Optional
from pypdf import PdfReader
import pdfplumber

def get_pdf_metadata(pdf_path: str) -> Dict[str, Any]:
    """Extracts PDF document metadata and security status."""
    reader = PdfReader(pdf_path)
    meta = reader.metadata or {}
    total_pages = len(reader.pages)
    first_page = reader.pages[0] if total_pages > 0 else None
    width_pt = float(first_page.mediabox.width) if first_page else 0.0
    height_pt = float(first_page.mediabox.height) if first_page else 0.0

    return {
        "title": meta.get("/Title", "(No Title)"),
        "author": meta.get("/Author", "(No Author)"),
        "producer": meta.get("/Producer", "(Unknown)"),
        "creator": meta.get("/Creator", "(Unknown)"),
        "creation_date": str(meta.get("/CreationDate", "")),
        "page_count": total_pages,
        "is_encrypted": reader.is_encrypted,
        "page_width_in": round(width_pt / 72.0, 2),
        "page_height_in": round(height_pt / 72.0, 2)
    }

def extract_pdf_text(pdf_path: str, page_range: Optional[List[int]] = None) -> str:
    """Extracts text page-by-page with clear slide/page delimiter markers."""
    text_blocks = []
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        pages_to_read = page_range if page_range else range(total)

        for p_idx in pages_to_read:
            if 0 <= p_idx < total:
                page = pdf.pages[p_idx]
                p_text = page.extract_text(layout=True) or ""
                text_blocks.append(f"<!-- Page {p_idx + 1} of {total} -->\n{p_text.strip()}\n")

    return "\n".join(text_blocks)

def extract_pdf_tables(pdf_path: str) -> str:
    """Extracts all detectable tables from the PDF and converts them to GFM Markdown tables."""
    md_output = []
    with pdfplumber.open(pdf_path) as pdf:
        for p_idx, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables()
            if not tables:
                continue

            for t_idx, tbl in enumerate(tables, 1):
                clean_rows = []
                for row in tbl:
                    if not row or not any(row):
                        continue
                    clean_row = [str(cell).strip().replace("\n", " ") if cell is not None else "" for cell in row]
                    clean_rows.append(clean_row)

                if not clean_rows:
                    continue

                col_count = len(clean_rows[0])
                md_output.append(f"### Page {p_idx} - Table {t_idx}\n")

                # Header
                header = clean_rows[0]
                md_output.append("| " + " | ".join(header) + " |")
                md_output.append("| " + " | ".join(["---"] * col_count) + " |")

                # Data Rows
                for r in clean_rows[1:]:
                    # Ensure row has same column count
                    padded_r = r + [""] * (col_count - len(r))
                    md_output.append("| " + " | ".join(padded_r[:col_count]) + " |")

                md_output.append("\n")

    return "\n".join(md_output)

def pdf_to_markdown(pdf_path: str) -> str:
    """Exports structured text and tables to Markdown format."""
    md_parts = []
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        for p_idx, page in enumerate(pdf.pages, 1):
            md_parts.append(f"\n## Page {p_idx} / {total}\n")

            # Check for tables on this page
            tables = page.extract_tables()
            if tables:
                for t_idx, tbl in enumerate(tables, 1):
                    clean_rows = []
                    for row in tbl:
                        if not row or not any(row):
                            continue
                        clean_row = [str(cell).strip().replace("\n", " ") if cell is not None else "" for cell in row]
                        clean_rows.append(clean_row)

                    if clean_rows:
                        col_count = len(clean_rows[0])
                        md_parts.append(f"\n**Table {t_idx}**\n")
                        md_parts.append("| " + " | ".join(clean_rows[0]) + " |")
                        md_parts.append("| " + " | ".join(["---"] * col_count) + " |")
                        for r in clean_rows[1:]:
                            padded_r = r + [""] * (col_count - len(r))
                            md_parts.append("| " + " | ".join(padded_r[:col_count]) + " |")
                        md_parts.append("\n")

            # Extract regular text
            text = page.extract_text() or ""
            if text:
                md_parts.append(text.strip() + "\n")

    return "\n".join(md_parts)

def analyze_pdf(pdf_path: str):
    """Generates an executive analysis report of the PDF."""
    if not os.path.exists(pdf_path):
        print(f"[ERROR] File not found: {pdf_path}")
        return

    meta = get_pdf_metadata(pdf_path)

    total_words = 0
    table_count = 0
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            total_words += len(t.split())
            tables = page.extract_tables()
            if tables:
                table_count += len(tables)

    print("=" * 60)
    print(f"PDF ANALYSIS REPORT: {os.path.basename(pdf_path)}")
    print("=" * 60)
    print(f"Title:        {meta['title']}")
    print(f"Author:       {meta['author']}")
    print(f"Page Count:   {meta['page_count']} pages")
    print(f"Dimensions:   {meta['page_width_in']}\" x {meta['page_height_in']}\"")
    print(f"Word Count:   {total_words} words (~{total_words // 250 + 1} min read)")
    print(f"Tables Found: {table_count}")
    print(f"Encrypted:    {'Yes [PASSWORD REQUIRED]' if meta['is_encrypted'] else 'No [OPEN]'}")
    print("=" * 60)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Read and analyze PDF documents")
    parser.add_argument("file", help="Path to .pdf file")
    parser.add_argument("--analyze", action="store_true", help="Print summary metrics and page counts")
    parser.add_argument("--markdown", action="store_true", help="Export full PDF to Markdown with tables")
    parser.add_argument("--tables", action="store_true", help="Extract only tables as Markdown")
    parser.add_argument("--pages", help="Specific pages (e.g. 1-3 or 5)")
    args = parser.parse_args()

    if not os.path.exists(args.file):
        print(f"[ERROR] File not found: {args.file}")
        sys.exit(1)

    if args.tables:
        print(extract_pdf_tables(args.file))
    elif args.markdown:
        print(pdf_to_markdown(args.file))
    elif args.pages:
        page_indices = []
        if "-" in args.pages:
            start, end = map(int, args.pages.split("-"))
            page_indices = list(range(start - 1, end))
        else:
            page_indices = [int(args.pages) - 1]
        print(extract_pdf_text(args.file, page_indices))
    else:
        analyze_pdf(args.file)
