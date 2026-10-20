"""
pdf_tools.py — locate, highlight, bookmark and box regions in PDFs (PyMuPDF).

Usage:
    python pdf_tools.py <cmd> <pdf> [options]

Commands: text, find, blocks, highlight, bookmark, rect, inspect, render.
Pages are 1-indexed; coordinates are PDF points, origin top-left.

Requirements:
    pip install pymupdf
"""

from __future__ import annotations
import argparse
import sys
from pathlib import Path

try:
    import pymupdf as fitz  # PyMuPDF >= 1.24
except ImportError:
    import fitz  # older PyMuPDF

COLORS = {
    "yellow": (1, 1, 0),
    "green": (0, 1, 0),
    "blue": (0.5, 0.8, 1),
    "pink": (1, 0.6, 0.8),
    "red": (1, 0, 0),
}


def _page(doc: fitz.Document, n: int) -> fitz.Page:
    if not 1 <= n <= doc.page_count:
        sys.exit(f"page {n} out of range (1..{doc.page_count})")
    return doc[n - 1]


def _rect(s: str) -> fitz.Rect:
    parts = [float(v) for v in s.split(",")]
    if len(parts) != 4:
        sys.exit(f"rect needs x0,y0,x1,y1, got {s!r}")
    return fitz.Rect(*parts)


def _out(src: str, out: str) -> Path:
    o = Path(out)
    if o.resolve() == Path(src).resolve():
        sys.exit("refusing to overwrite the source PDF; pick a new --out")
    if o.exists():
        sys.exit(f"{o} already exists; pick a new --out")
    return o


def _save(doc: fitz.Document, src: str, out: str) -> None:
    o = _out(src, out)
    doc.save(str(o), garbage=3, deflate=True)
    print(o)


def _fmt(r: fitz.Rect) -> str:
    return f"{r.x0:.1f},{r.y0:.1f},{r.x1:.1f},{r.y1:.1f}"


# ---------------------------------------------------------------------------
# READ
# ---------------------------------------------------------------------------

def cmd_text(a):
    doc = fitz.open(a.pdf)
    pages = [int(p) for p in a.pages.split(",")] if a.pages else range(1, doc.page_count + 1)
    for n in pages:
        print(f"===== page {n} =====")
        print(_page(doc, n).get_text())


def cmd_find(a):
    doc = fitz.open(a.pdf)
    page = _page(doc, a.page)
    print(f"page {a.page} size: {page.rect.width:.0f} x {page.rect.height:.0f}")
    for t in a.text:
        hits = page.search_for(t)
        print(f"{t!r}: {len(hits)} hit(s)")
        for r in hits:
            print(f"  {_fmt(r)}")


def cmd_blocks(a):
    doc = fitz.open(a.pdf)
    page = _page(doc, a.page)
    print(f"page {a.page} size: {page.rect.width:.0f} x {page.rect.height:.0f}")
    for x0, y0, x1, y1, text, no, kind in page.get_text("blocks"):
        label = "img" if kind else " ".join(text.split())[:90]
        print(f"[{no}] {x0:.1f},{y0:.1f},{x1:.1f},{y1:.1f}  {label}")


def cmd_inspect(a):
    doc = fitz.open(a.pdf)
    print(f"pages: {doc.page_count}")
    toc = doc.get_toc(simple=False)
    print(f"bookmarks: {len(toc)}")
    for lvl, title, pno, dest in toc:
        y = dest.get("to").y if isinstance(dest, dict) and dest.get("to") else None
        ys = f" y={y:.0f}" if y is not None else ""
        print(f"  {'  ' * (lvl - 1)}{title!r} -> p{pno}{ys}")
    for page in doc:
        annots = [(an.type[1], _fmt(an.rect)) for an in page.annots()]
        if annots:
            print(f"page {page.number + 1}:")
            for kind, r in annots:
                print(f"  {kind} {r}")


def cmd_render(a):
    doc = fitz.open(a.pdf)
    page = _page(doc, a.page)
    clip = _rect(a.clip) if a.clip else None
    pix = page.get_pixmap(dpi=a.dpi, clip=clip, annots=True)
    pix.save(a.out)
    print(a.out)


# ---------------------------------------------------------------------------
# WRITE
# ---------------------------------------------------------------------------

def cmd_highlight(a):
    doc = fitz.open(a.pdf)
    page = _page(doc, a.page)
    color = COLORS.get(a.color)
    if color is None:
        sys.exit(f"unknown color {a.color!r}; choose from {', '.join(COLORS)}")
    plan = []
    for t in a.text:
        quads = page.search_for(t, quads=True)
        if not quads:
            sys.exit(f"no match for {t!r} on page {a.page}; nothing written")
        plan.append((t, quads if a.all_hits else quads[:1]))
    for t, quads in plan:
        an = page.add_highlight_annot(quads)
        an.set_colors(stroke=color)
        an.update()
        print(f"highlighted {t!r} ({len(quads)} hit(s))")
    _save(doc, a.pdf, a.out)


def cmd_bookmark(a):
    doc = fitz.open(a.pdf)
    _page(doc, a.page)
    toc = doc.get_toc(simple=False)
    dest = {"kind": fitz.LINK_GOTO, "page": a.page - 1, "to": fitz.Point(a.x, a.y), "zoom": 0}
    toc.append([1, a.title, a.page, dest])
    doc.set_toc(toc)
    _save(doc, a.pdf, a.out)


def cmd_rect(a):
    doc = fitz.open(a.pdf)
    page = _page(doc, a.page)
    an = page.add_rect_annot(_rect(a.rect))
    an.set_colors(stroke=COLORS["red"])
    an.set_border(width=a.width)
    an.update()
    _save(doc, a.pdf, a.out)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    p = argparse.ArgumentParser(description="PyMuPDF PDF tools")
    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("text", help="dump text"); s.add_argument("pdf")
    s.add_argument("--pages", help="comma-separated, 1-indexed"); s.set_defaults(fn=cmd_text)

    s = sub.add_parser("find", help="search rects for text"); s.add_argument("pdf")
    s.add_argument("--page", type=int, required=True)
    s.add_argument("--text", action="append", required=True); s.set_defaults(fn=cmd_find)

    s = sub.add_parser("blocks", help="text block coordinates"); s.add_argument("pdf")
    s.add_argument("--page", type=int, required=True); s.set_defaults(fn=cmd_blocks)

    s = sub.add_parser("highlight", help="highlight text"); s.add_argument("pdf")
    s.add_argument("--page", type=int, required=True)
    s.add_argument("--text", action="append", required=True)
    s.add_argument("--out", required=True)
    s.add_argument("--color", default="yellow")
    s.add_argument("--all-hits", action="store_true"); s.set_defaults(fn=cmd_highlight)

    s = sub.add_parser("bookmark", help="append a top-level bookmark"); s.add_argument("pdf")
    s.add_argument("--page", type=int, required=True)
    s.add_argument("--title", required=True)
    s.add_argument("--x", type=float, default=0)
    s.add_argument("--y", type=float, default=0)
    s.add_argument("--out", required=True); s.set_defaults(fn=cmd_bookmark)

    s = sub.add_parser("rect", help="red rectangle annotation"); s.add_argument("pdf")
    s.add_argument("--page", type=int, required=True)
    s.add_argument("--rect", required=True, help="x0,y0,x1,y1")
    s.add_argument("--width", type=float, default=2)
    s.add_argument("--out", required=True); s.set_defaults(fn=cmd_rect)

    s = sub.add_parser("inspect", help="list bookmarks and annotations"); s.add_argument("pdf")
    s.set_defaults(fn=cmd_inspect)

    s = sub.add_parser("render", help="render page to PNG"); s.add_argument("pdf")
    s.add_argument("--page", type=int, required=True)
    s.add_argument("--out", required=True)
    s.add_argument("--clip", help="x0,y0,x1,y1")
    s.add_argument("--dpi", type=int, default=110); s.set_defaults(fn=cmd_render)

    a = p.parse_args()
    a.fn(a)


if __name__ == "__main__":
    main()
