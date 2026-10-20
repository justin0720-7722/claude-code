---
name: pdf-read-write-v1
description: Read, locate, highlight, bookmark and box regions in PDFs with PyMuPDF in the reflex311 conda env. Use when asked to find a passage (e.g. a paper's research goal), highlight it, add a bookmark, draw a rectangle around a table/figure, or visually verify PDF annotations.
---

# pdf-read-write-v1

Tool: `scripts/pdf_tools.py` (PyMuPDF). Python env (from `.claude/settings.local.json`):

```
C:\Users\pingw\anaconda3\envs\reflex311\python.exe -I .claude\skills\pdf-read-write-v1\scripts\pdf_tools.py <cmd> <pdf> [options]
```

Plain `python` is the Windows Store stub and does not work; always use the full path.
On macOS/Linux (or if that env is missing), use any Python with `pip install pymupdf`:
`python3 -I .claude/skills/pdf-read-write-v1/scripts/pdf_tools.py <cmd> <pdf> [options]`.

## Commands

| cmd | purpose |
|---|---|
| `text <pdf> [--pages 1,2]` | dump text, to find the passage |
| `find <pdf> --page N --text "frag" ...` | search rects for text fragments; also prints page size |
| `blocks <pdf> --page N` | text block coordinates, for locating tables/captions |
| `highlight <pdf> --page N --text "line1" --text "line2" --out new.pdf [--color yellow] [--all-hits]` | highlight text |
| `bookmark <pdf> --page N --title "kw" --x X --y Y --out new.pdf` | append a top-level bookmark |
| `rect <pdf> --page N --rect x0,y0,x1,y1 --out new.pdf [--width 2]` | red rectangle annotation |
| `inspect <pdf>` | list bookmarks and annotation types per page |
| `render <pdf> --page N --out x.png [--clip x0,y0,x1,y1] [--dpi 110]` | PNG for visual check |

## Workflow

1. `text` to read; pick the passage (e.g. the sentence stating the research goal).
2. `highlight`: pass each wrapped line as its own `--text`, since text wraps across lines. Use `--all-hits` for a fragment that splits across words. Fragments that match nothing abort the run.
3. `bookmark`: put `--y` about 20pt above the highlight so the jump lands just before it; keep the title to a few keywords.
4. For a table/figure: `find`/`blocks` to get the caption and title coordinates, then `rect` with a ~2-4pt margin around the whole object (not the caption).
5. Chain outputs (`a.pdf` -> `b.pdf` -> `c.pdf`) so earlier edits are kept.
6. Verify: `inspect`, then `render` and look at the PNG (use `--clip` for a zoomed check). Check the box does not cut the object or cover neighbouring text.

## Rules

- Never overwrite the source PDF or an existing output; use a new `--out` name (e.g. `_goal_highlighted`, `_goal_table`).
- Coordinates are PDF points, origin top-left; page 1 of the Hyde paper is 594 x 759.
- Put throwaway outputs (PNGs) in the scratchpad, not the project folder.

## Worked example (Hyde 2008, `Hyde_2008_gender_similarities_math.pdf`)

- Goal sentence on p1: "Since this study ... merit a new analysis of gender and math performance." highlighted yellow, bookmark "Research goal: new analysis of gender & math".
- Table "Percentage of children scoring above indicated percentile and ratios": rect `165,349,427,467` on p1.
