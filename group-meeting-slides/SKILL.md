---
name: group-meeting-slides
description: Turn an English research paper PDF into a Chinese group-meeting (组会) presentation .pptx that keeps the paper's own figures and tables as original images. Use when the user supplies a paper PDF and wants a 组会/文献汇报/paper-reading slide deck built from it.
---

# 英文文献 → 中文组会汇报 PPT

Input is one English paper PDF. Output is a `.pptx` for a Chinese lab group
meeting, plus speaker notes. **Every figure and table in the deck must be the
paper's own pixels**, cropped from the PDF by `scripts/extract.py`. Never
redraw a chart, never invent numbers, never leave a placeholder rectangle.

## Why the split

The scripts own everything mechanical and reproducible: PDF geometry, figure
cropping, typography, and aspect-ratio-safe image placement. You own the part
that needs judgement: reading the paper, deciding the argument, and writing
Chinese slide copy. Do not try to do layout by hand in python-pptx, and do not
try to summarise the paper with a script.

## Workflow

Run scripts with the Python executable reported by `load_workspace_dependencies`
(it has PyMuPDF, python-pptx and Pillow). Script paths are relative to this
skill directory. Keep all outputs in the task workspace, in one directory per
paper, e.g. `<workspace>/<paper-slug>/`.

### 1. Extract

```bash
"<python>" scripts/extract.py <paper.pdf> --out <workdir> --dpi 300
```

This writes:

| File | What it is |
|---|---|
| `paper.md` | page-by-page text with `[[FIGURE-3]]` / `[[TABLE-2]]` markers inline |
| `figures.json` | manifest: key, kind, page, bbox, caption, path, pixel size |
| `figures/fig-N.png` | original figure pixels |
| `tables/tab-N.png` | original table pixels |
| `pages/page-N.png` | full page renders (fallback) |

### 2. Inspect every extracted image before writing any slide

Open **each** file in `figures/` and `tables/` with the image-reading tool. This
is not optional: crops are geometry heuristics and they do fail.

**First, the semantic check — every time.** Read the `caption` for that image in
`figures.json`, then look at the picture:

- Does the table/figure number match ("Table 2" really is Table 2)?
- Do the row labels and headline numbers in the image match what the caption
  describes?
- Are two crops byte-identical (`Get-FileHash`/`md5sum`)? That means two
  captions claimed the same object, and at least one is wrong.

This check catches the failures geometry cannot see. A crop can be perfectly
framed and still be the wrong table.

Then the geometric checks:

- **Chopped panel** — a multi-panel figure where only one panel was captured.
- **Caption bleed** — the paper's own "Figure 4: ..." line inside the crop.
- **Missing heading** — a sub-panel title sliced in half at the top.
- **Missing axis furniture** — a chart with no axis titles, no tick labels, or
  no legend, because those are text and not part of the drawing geometry.
- **Empty or near-empty crop** — a white strip, or a few points of a line.

When a crop is wrong, do not accept it. In practice **manual re-crop is the
normal repair path** — the flag tweaks rarely fix geometry bugs:

1. Re-run with a different `--min-area` — occasionally recovers a dropped small
   panel. `--dpi` only changes sharpness, never framing.
2. Re-crop from the PDF yourself. This always works and is the reliable route:

```bash
"<python>" - <<'PY'
import pymupdf
doc = pymupdf.open("<paper.pdf>")
page = doc[<page>]          # `page` in figures.json is 0-BASED; pass it directly
rect = pymupdf.Rect(x0, y0, x1, y1)   # PDF points; open pages/page-N.png to pick them
page.get_pixmap(matrix=pymupdf.Matrix(300/72, 300/72), clip=rect).save("<workdir>/figures/fig-3.png")
PY
```

Useful geometry facts when hand-picking a box:

- `figures.json` records the bbox actually clipped, so you can compare it with
  the image you got and see in which direction it is wrong.
- Chart axis titles and tick labels are **text**, sitting a few points outside
  the plot frame; include roughly 30–40pt beyond the frame.
- A journal caption is set to nearly the full width of the art; in-figure panel
  labels are not. That distinction decides where the bottom edge belongs.

Quote the caption in the manifest to confirm you cropped the right thing. If a
figure is simply unusable after all of this, drop that slide rather than
shipping a bad image — a missing slide is obvious, a wrong one is not.

### 3. Read the paper properly

Read `paper.md` end to end. You need, at minimum: the problem, the gap in prior
work, the core idea, the mechanism, the headline numbers (from the tables, not
from memory), the ablations, and the limitations. Read the actual caption text
in `figures.json` so your Chinese explanation matches what the figure plots.

### 4. Write the deck JSON

Follow `references/deck-schema.md` for the exact schema and
`references/storyline.md` for the Chinese group-meeting narrative. Write it to
`<workdir>/deck.json`; image paths are relative to the JSON file.

### 5. Build

```bash
"<python>" scripts/build_pptx.py <workdir>/deck.json --out "<workdir>/<课题>_组会汇报.pptx"
```

### 6. Verify before delivering

Reopen and check the numbers, then render and look at it:

```bash
"<python>" - <<'PY'
from pptx import Presentation
p = Presentation("<deck.pptx>")
print(len(p.slides), "slides")
for i, s in enumerate(p.slides, 1):
    pics = [sh for sh in s.shapes if sh.shape_type == 13]
    texts = [sh.text_frame.text.split("\n")[0] for sh in s.shapes if sh.has_text_frame]
    print(i, f"pics={len(pics)}", texts[:3])
PY
```

Then render with the `office-render` skill and actually look at the slides.
Check: no text overlapping the rule under the title, no caption colliding with
a takeaway, images not distorted, no slide with a lonely half-empty column.
Fix the JSON (or, if it is a layout bug, `build_pptx.py`) and rebuild.

Finally write the speaker notes file (`讲稿.md`) and deliver both with
`present`.

## Hard rules

- **Never fabricate a number.** Every figure quoted in slide copy must be
  traceable to the paper text or to a table you cropped.
- **Never re-create a chart** with python-pptx from estimated values. If the
  user explicitly asks for a redrawn chart, say what it will cost in fidelity
  and get agreement first.
- **Keep the aspect ratio.** `build_pptx.py` enforces this; do not bypass it
  with raw `add_picture` calls.
- **Cite the figure** under every image: keep the original English caption at
  small size so the audience can find it in the paper.
- Slide copy is Chinese; paper-specific technical terms stay in English on
  first use, e.g. 缩放点积注意力（scaled dot-product attention）.

## Reference files

- `references/deck-schema.md` — layouts, fields, and a worked example.
- `references/storyline.md` — the 12–18 slide group-meeting structure and the
  Chinese writing conventions.
