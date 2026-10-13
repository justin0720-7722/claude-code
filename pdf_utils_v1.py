"""
pdf_utils_v1.py — Portable PDF read/write utility for any project.

Usage:
    from pdf_utils_v1 import read_text, read_tables, merge, split, create, watermark, protect

Requirements:
    pip install pypdf pdfplumber reportlab
Optional (OCR):
    pip install pytesseract pdf2image
"""

from __future__ import annotations
from pathlib import Path
from typing import Any


# ---------------------------------------------------------------------------
# READ
# ---------------------------------------------------------------------------

def read_text(path: str | Path, *, layout: bool = False) -> str:
    """Extract all text from a PDF. Set layout=True to preserve spacing."""
    import pdfplumber
    with pdfplumber.open(str(path)) as pdf:
        pages = [p.extract_text(layout=layout) or "" for p in pdf.pages]
    return "\n\n".join(pages)


def read_page_text(path: str | Path, page: int = 0) -> str:
    """Extract text from a single page (0-indexed)."""
    import pdfplumber
    with pdfplumber.open(str(path)) as pdf:
        return pdf.pages[page].extract_text() or ""


def read_tables(path: str | Path) -> list[list[list[Any]]]:
    """Return all tables from every page as a list of tables (each a list of rows)."""
    import pdfplumber
    result = []
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            result.extend(page.extract_tables())
    return result


def read_metadata(path: str | Path) -> dict[str, str | None]:
    """Return PDF metadata dict (title, author, subject, creator, producer)."""
    from pypdf import PdfReader
    m = PdfReader(str(path)).metadata
    return {
        "title": m.title,
        "author": m.author,
        "subject": m.subject,
        "creator": m.creator,
        "producer": m.producer,
    }


def read_ocr(path: str | Path) -> str:
    """OCR a scanned PDF (requires pytesseract + pdf2image + poppler)."""
    import pytesseract
    from pdf2image import convert_from_path
    images = convert_from_path(str(path))
    parts = [f"Page {i+1}:\n{pytesseract.image_to_string(img)}" for i, img in enumerate(images)]
    return "\n\n".join(parts)


# ---------------------------------------------------------------------------
# WRITE / MANIPULATE
# ---------------------------------------------------------------------------

def merge(inputs: list[str | Path], output: str | Path) -> Path:
    """Merge multiple PDFs into one output file."""
    from pypdf import PdfWriter, PdfReader
    writer = PdfWriter()
    for f in inputs:
        for page in PdfReader(str(f)).pages:
            writer.add_page(page)
    out = Path(output)
    with out.open("wb") as fh:
        writer.write(fh)
    return out


def split(path: str | Path, output_dir: str | Path | None = None) -> list[Path]:
    """Split a PDF into one file per page. Returns list of created paths."""
    from pypdf import PdfReader, PdfWriter
    src = Path(path)
    dest = Path(output_dir) if output_dir else src.parent
    dest.mkdir(parents=True, exist_ok=True)
    reader = PdfReader(str(src))
    out_paths = []
    for i, page in enumerate(reader.pages):
        writer = PdfWriter()
        writer.add_page(page)
        out = dest / f"{src.stem}_page_{i+1}.pdf"
        with out.open("wb") as fh:
            writer.write(fh)
        out_paths.append(out)
    return out_paths


def rotate(path: str | Path, degrees: int, output: str | Path, pages: list[int] | None = None) -> Path:
    """Rotate pages by degrees (90/180/270). pages=None rotates all."""
    from pypdf import PdfReader, PdfWriter
    reader = PdfReader(str(path))
    writer = PdfWriter()
    for i, page in enumerate(reader.pages):
        if pages is None or i in pages:
            page.rotate(degrees)
        writer.add_page(page)
    out = Path(output)
    with out.open("wb") as fh:
        writer.write(fh)
    return out


def watermark(path: str | Path, watermark_pdf: str | Path, output: str | Path) -> Path:
    """Overlay a watermark PDF onto every page."""
    from pypdf import PdfReader, PdfWriter
    wm_page = PdfReader(str(watermark_pdf)).pages[0]
    reader = PdfReader(str(path))
    writer = PdfWriter()
    for page in reader.pages:
        page.merge_page(wm_page)
        writer.add_page(page)
    out = Path(output)
    with out.open("wb") as fh:
        writer.write(fh)
    return out


def protect(path: str | Path, user_pw: str, owner_pw: str, output: str | Path) -> Path:
    """Password-protect a PDF."""
    from pypdf import PdfReader, PdfWriter
    reader = PdfReader(str(path))
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    writer.encrypt(user_pw, owner_pw)
    out = Path(output)
    with out.open("wb") as fh:
        writer.write(fh)
    return out


# ---------------------------------------------------------------------------
# CREATE
# ---------------------------------------------------------------------------

def create(output: str | Path, sections: list[dict[str, str]]) -> Path:
    """
    Create a simple multi-section PDF.

    sections: list of dicts with keys:
        "heading"  (optional) — section title
        "body"     (required) — paragraph text
        "new_page" (optional, bool) — force page break before this section
    """
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet

    out = Path(output)
    doc = SimpleDocTemplate(str(out), pagesize=letter)
    styles = getSampleStyleSheet()
    story: list[Any] = []

    for sec in sections:
        if sec.get("new_page") and story:
            story.append(PageBreak())
        if sec.get("heading"):
            story.append(Paragraph(sec["heading"], styles["Heading1"]))
            story.append(Spacer(1, 6))
        story.append(Paragraph(sec["body"], styles["Normal"]))
        story.append(Spacer(1, 12))

    doc.build(story)
    return out


# ---------------------------------------------------------------------------
# CLI (optional quick use)
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import argparse, json, sys

    p = argparse.ArgumentParser(description="pdf_utils_v1 CLI")
    sub = p.add_subparsers(dest="cmd")

    t = sub.add_parser("text", help="Extract text")
    t.add_argument("pdf"); t.add_argument("--layout", action="store_true")

    m = sub.add_parser("merge", help="Merge PDFs")
    m.add_argument("inputs", nargs="+"); m.add_argument("-o", "--output", required=True)

    s = sub.add_parser("split", help="Split PDF into pages")
    s.add_argument("pdf"); s.add_argument("-d", "--dir")

    me = sub.add_parser("meta", help="Show metadata as JSON")
    me.add_argument("pdf")

    args = p.parse_args()

    if args.cmd == "text":
        print(read_text(args.pdf, layout=args.layout))
    elif args.cmd == "merge":
        print(merge(args.inputs, args.output))
    elif args.cmd == "split":
        for f in split(args.pdf, args.dir):
            print(f)
    elif args.cmd == "meta":
        print(json.dumps(read_metadata(args.pdf), indent=2, default=str))
    else:
        p.print_help(); sys.exit(1)
