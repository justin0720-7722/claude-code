"""PyMuPDF helpers for reading and annotating PDFs.

Run with the project env, isolated:
  C:\\Users\\pingw\\anaconda3\\envs\\reflex311\\python.exe -I pdf_tools.py <cmd> ...

Never overwrites the input: write commands take an explicit output path.
"""
import argparse
import sys

import pymupdf


def parse_rect(s):
    return pymupdf.Rect(*[float(v) for v in s.split(",")])


def cmd_text(a):
    d = pymupdf.open(a.pdf)
    pages = range(len(d)) if a.pages is None else [int(p) - 1 for p in a.pages.split(",")]
    for i in pages:
        print(f"=== page {i + 1}")
        print(d[i].get_text())


def cmd_find(a):
    p = pymupdf.open(a.pdf)[a.page - 1]
    print("page rect:", p.rect)
    for t in a.text:
        print(t, "->", p.search_for(t))


def cmd_blocks(a):
    p = pymupdf.open(a.pdf)[a.page - 1]
    for b in p.get_text("blocks"):
        print(tuple(round(v) for v in b[:4]), b[4][:40].replace("\n", " "))


def cmd_highlight(a):
    """Highlight text. Pass each wrapped line as its own --text fragment."""
    d = pymupdf.open(a.pdf)
    p = d[a.page - 1]
    rects = []
    for t in a.text:
        r = p.search_for(t)
        print(t, len(r))
        if not r:
            sys.exit(f"not found: {t!r}")
        # one hit is the normal case; a phrase split across words returns
        # several rects that all belong to the highlight, so keep them all.
        rects += r if a.all_hits else r[:1]
    color = {"yellow": (1, 1, 0), "green": (0, 1, 0), "red": (1, 0, 0)}[a.color]
    ann = p.add_highlight_annot(rects)
    ann.set_colors(stroke=color)
    ann.update()
    d.save(a.out)
    print("saved", a.out)


def cmd_bookmark(a):
    d = pymupdf.open(a.pdf)
    toc = d.get_toc()
    dest = {"kind": pymupdf.LINK_GOTO, "to": pymupdf.Point(a.x, a.y), "zoom": 0}
    toc.append([1, a.title, a.page, dest])
    d.set_toc(toc)
    d.save(a.out)
    print("saved", a.out)


def cmd_rect(a):
    d = pymupdf.open(a.pdf)
    ann = d[a.page - 1].add_rect_annot(parse_rect(a.rect))
    ann.set_colors(stroke=(1, 0, 0))
    ann.set_border(width=a.width)
    ann.update()
    d.save(a.out)
    print("saved", a.out)


def cmd_inspect(a):
    d = pymupdf.open(a.pdf)
    print("pages:", len(d), "toc:", d.get_toc())
    for p in d:
        print("page", p.number + 1, [x.type[1] for x in p.annots()])


def cmd_render(a):
    p = pymupdf.open(a.pdf)[a.page - 1]
    clip = parse_rect(a.clip) if a.clip else None
    p.get_pixmap(clip=clip, dpi=a.dpi).save(a.out)
    print("saved", a.out)


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)

    def add(name, fn, *args):
        sp = sub.add_parser(name)
        sp.add_argument("pdf")
        for flags, kw in args:
            sp.add_argument(*flags, **kw)
        sp.set_defaults(fn=fn)

    page = (["--page"], dict(type=int, default=1))
    out = (["--out"], dict(required=True))
    add("text", cmd_text, (["--pages"], dict(help="comma list, 1-based")))
    add("find", cmd_find, page, (["--text"], dict(action="append", required=True)))
    add("blocks", cmd_blocks, page)
    add("highlight", cmd_highlight, page, out,
        (["--text"], dict(action="append", required=True)),
        (["--color"], dict(default="yellow", choices=["yellow", "green", "red"])),
        (["--all-hits"], dict(action="store_true")))
    add("bookmark", cmd_bookmark, page, out,
        (["--title"], dict(required=True)),
        (["--x"], dict(type=float, default=0)),
        (["--y"], dict(type=float, default=0)))
    add("rect", cmd_rect, page, out,
        (["--rect"], dict(required=True, help="x0,y0,x1,y1 in PDF points")),
        (["--width"], dict(type=float, default=2)))
    add("inspect", cmd_inspect)
    add("render", cmd_render, page, out,
        (["--clip"], dict(help="x0,y0,x1,y1")), (["--dpi"], dict(type=int, default=110)))
    a = ap.parse_args()
    a.fn(a)


main()
