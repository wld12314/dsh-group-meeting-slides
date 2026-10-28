#!/usr/bin/env python3
"""Extract structure, text and ORIGINAL figures/tables from an English paper PDF.

The whole point of this script is that it never draws, redraws or invents a
chart.  Every image it emits is the publisher's own pixels, cropped from the
page at the figure's real bounding box, so numbers in the final slide deck are
exactly the numbers the authors plotted.

Usage:
    python extract.py paper.pdf --out out_dir [--dpi 300] [--min-area 0.0025]

Output (all inside --out):
    paper.md          per-page text with [[FIGURE-N]] / [[TABLE-N]] markers
    figures.json      manifest: key, page (0-BASED), bbox as rendered, caption, path
    figures/fig-3.png original pixels of Figure 3
    tables/tab-2.png  original pixels of Table 2
    pages/page-3.png  full-page render (fallback for any bad crop)

Note the file names use fig-N / tab-N while the manifest keys are figure-N /
table-N.  The `bbox` recorded in figures.json is the box actually clipped, so it
can be compared against the image to debug a bad crop.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

import pymupdf

# A caption starts a figure/table block.  Papers use all of these variants.
CAPTION_RE = re.compile(
    r"^\s*(?:(?P<num>\d{1,2})\s*[|｜]\s*)?"          # "3 |" side-label style
    r"(?P<kind>Figure|Fig\.?|FIGURE|Table|TABLE)\s*"
    r"(?P<label>[0-9]{1,2}|[A-Z][0-9]{0,2}|[IVX]{1,4})"
    r"\s*(?P<sep>[:.．|—–])\s*(?P<text>.*)$"
)

# A caption must be punctuated after its label.  Without this, a body sentence
# such as "Table 2 summarizes our results" is mistaken for a caption and the
# paragraph gets cropped as if it were the table.

SECTION_RE = re.compile(
    r"^\s*(?P<num>\d{1,2}(?:\.\d{1,2}){0,2})?\.?\s*"
    r"(?P<title>Abstract|Introduction|Related Work|Background|Method(?:s|ology)?|"
    r"Approach|Model|Experiments?|Results?|Evaluation|Discussion|"
    r"Conclusion(?:s)?|Limitations?|Acknowledge?ments?|References)\b",
    re.IGNORECASE,
)


def norm_kind(kind: str) -> str:
    return "table" if kind.lower().startswith("tab") else "figure"


def clean(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def text_blocks(page) -> list[tuple]:
    """Text blocks only: (x0, y0, x1, y1, text, block_no, block_type)."""
    return [b for b in page.get_text("blocks") if b[6] == 0]


def collect_captions(doc) -> list[dict]:
    """Find every caption and the page region that holds it."""
    out = []
    for pno in range(len(doc)):
        page = doc[pno]
        for block in text_blocks(page):
            text = clean(block[4])
            if not text or len(text) > 1200:
                continue
            m = CAPTION_RE.match(text)
            if not m:
                continue
            body = m.group("text")
            # A caption carries its own text; "Table 3" alone is a cross
            # reference, not a caption.
            if len(body) < 8:
                continue
            out.append({
                "page": pno,
                "kind": norm_kind(m.group("kind")),
                "label": m.group("label"),
                "caption": text,
                "bbox": list(block[:4]),
            })
    return out


def detect_columns(page) -> list[tuple[float, float]]:
    """Split the page into text columns from the words' x positions.

    A two-column paper puts a wide empty gutter between the columns; tables and
    their captions live inside one column, and a table box that crosses the
    gutter is almost certainly two tables welded together.
    """
    words = page.get_text("words")
    if not words:
        return [(page.rect.x0, page.rect.x1)]
    xs = sorted(w[0] for w in words)
    mid = (page.rect.x0 + page.rect.x1) / 2
    left = [w for w in words if w[0] < mid]
    right = [w for w in words if w[0] >= mid]
    if not left or not right:
        return [(page.rect.x0, page.rect.x1)]
    left_edge = min(w[0] for w in left)
    right_edge = max(w[2] for w in words if w[2] < mid) if any(w[2] < mid for w in words) else mid
    gutter = min(w[0] for w in right) - right_edge
    if gutter < 8:
        return [(page.rect.x0, page.rect.x1)]
    return [(left_edge, right_edge), (min(w[0] for w in right), page.rect.x1)]


def column_of(box: list[float], columns: list[tuple[float, float]]) -> int:
    """Index of the column holding the box's centre."""
    cx = (box[0] + box[2]) / 2
    for i, (x0, x1) in enumerate(columns):
        if x0 - 20 <= cx <= x1 + 20:
            return i
    return 0 if cx < (columns[0][1] + columns[-1][0]) / 2 else len(columns) - 1


def clip_to_column(rect: list[float], col: tuple[float, float]) -> list[float]:
    return [max(rect[0], col[0] - 6), rect[1], min(rect[2], col[1] + 6), rect[3]]


def assign_tables(page, caps):
    """Pair each table caption on a page with a distinct table grid.

    Distance alone is ambiguous: on a page that stacks two tables with their
    captions underneath, both captions are the same distance from the same
    grid.  So pair per column and in document order - within a column the
    sequence is always table, caption, table, caption - and require the result
    to stay inside that column.
    """
    columns = detect_columns(page)

    try:
        found = page.find_tables()
    except Exception:
        found = None
    raw_grids = [list(t.bbox) for t in (getattr(found, "tables", []) or [])] if found else []

    # find_tables() happily returns one box spanning both columns when two
    # side-by-side tables share rows.  Split any grid that straddles the gutter
    # before pairing, otherwise one box serves two different tables.
    grids = []
    for g in raw_grids:
        straddles = any(col[1] - 2 < (g[0] + g[2]) / 2 < col[1] + 2 for col in columns[:-1])
        inside = None
        for col in columns:
            if g[0] >= col[0] - 8 and g[2] <= col[1] + 8:
                inside = col
                break
        if inside is not None:
            grids.append(g)
            continue
        if straddles or len(columns) > 1:
            for col in columns:
                piece = clip_to_column(g, col)
                if piece[2] - piece[0] > 40:
                    grids.append(piece)
            continue
        grids.append(g)

    # Deduplicate near-identical boxes produced by the split.
    unique = []
    for g in grids:
        if not any(abs(g[0] - u[0]) < 3 and abs(g[1] - u[1]) < 3
                   and abs(g[2] - u[2]) < 3 and abs(g[3] - u[3]) < 3 for u in unique):
            unique.append(g)
    grids = unique

    result: dict[int, list[float]] = {}

    # Group captions and grids by column.
    caps_by_col: dict[int, list[int]] = {}
    for ci, cap in enumerate(caps):
        caps_by_col.setdefault(column_of(cap["bbox"], columns), []).append(ci)
    grids_by_col: dict[int, list[int]] = {}
    for gi, g in enumerate(grids):
        grids_by_col.setdefault(column_of(g, columns), []).append(gi)

    for col_i, cap_ids in caps_by_col.items():
        col = columns[col_i] if col_i < len(columns) else (page.rect.x0, page.rect.x1)
        cap_ids.sort(key=lambda i: caps[i]["bbox"][1])
        grid_ids = sorted(grids_by_col.get(col_i, []), key=lambda i: grids[i][1])

        def gap_for(ci, gi):
            cx0, cy0, cx1, cy1 = caps[ci]["bbox"]
            r = grids[gi]
            if r[3] <= cy0 + 4:
                return cy0 - r[3]
            if r[1] >= cy1 - 4:
                return r[1] - cy1
            return None

        # Pass 1: index-wise pairing in document order.  Within a column a
        # paper interleaves [grid][caption] (or [caption][grid]) regularly, so
        # the k-th caption belongs to the k-th grid.  Distance alone cannot
        # decide this: two stacked tables can be equally close to both
        # captions, and a greedy nearest-match then swaps them.
        if grid_ids and len(cap_ids) == len(grid_ids):
            trial = list(zip(cap_ids, grid_ids))
            gaps = [gap_for(ci, gi) for ci, gi in trial]
            if all(g is not None and g <= 45 for g in gaps):
                for ci, gi in trial:
                    result[ci] = clip_to_column(grids[gi], col)
                continue

        # Pass 2: unequal counts - fall back to nearest match, never reusing a
        # grid, which at least cannot duplicate an image.
        pairs = []
        for ci in cap_ids:
            for gi in grid_ids:
                g = gap_for(ci, gi)
                if g is None or g > 45:
                    continue
                pairs.append((g, ci, gi))
        pairs.sort()
        used_t, used_c = set(), set()
        for gap, ci, gi in pairs:
            if ci in used_c or gi in used_t:
                continue
            used_c.add(ci)
            used_t.add(gi)
            result[ci] = clip_to_column(grids[gi], col)

        # Pass 3: still-unmatched captions take the leftover grids by order.
        leftover_caps = [ci for ci in cap_ids if ci not in result]
        leftover_grids = [gi for gi in grid_ids if gi not in used_t]
        if len(leftover_caps) == len(leftover_grids):
            for ci, gi in zip(leftover_caps, leftover_grids):
                result[ci] = clip_to_column(grids[gi], col)

    # Captions that found no grid fall back to rules / text heuristics, clipped
    # to their own column so a ruled fallback cannot span the gutter.
    for ci, cap in enumerate(caps):
        if ci in result:
            continue
        col = columns[column_of(cap["bbox"], columns)] if columns else (page.rect.x0, page.rect.x1)
        rect = table_rect_from_rules(page, cap, col) or fallback_table_rect(page, cap)
        if rect is None:
            rect = table_region_from_text(page, cap)
        if rect is not None:
            result[ci] = clip_to_column(list(rect), col)
    return result


def find_table_rect(page, cap):
    """Single-caption lookup, kept for interactive debugging."""
    return assign_tables(page, [cap]).get(0)


def table_rect_from_rules(page, cap, col=None):
    """Rebuild a ruled table's box from its horizontal rules.

    Rules are grouped into runs that stay within the given column, so two
    stacked tables on one page are not welded together.
    """
    cx0, cy0, cx1, cy1 = cap["bbox"]
    rules = []
    try:
        for d in page.get_drawings():
            r = d["rect"]
            if r.height > 2.5 or r.width < 60:
                continue
            if col is not None and (r.x1 < col[0] - 6 or r.x0 > col[1] + 6):
                continue
            rules.append((r.y0, r.x0, r.x1))
    except Exception:
        return None
    if len(rules) < 3:
        return None

    rules.sort()
    groups, current = [], [rules[0]]
    for item in rules[1:]:
        if item[0] - current[-1][0] <= 60:
            current.append(item)
        else:
            groups.append(current)
            current = [item]
    groups.append(current)

    best = None
    for g in groups:
        if len(g) < 3:
            continue
        top, bottom = g[0][0], g[-1][0]
        left = min(r[1] for r in g)
        right = max(r[2] for r in g)
        if bottom <= cy0 + 6:
            gap = cy0 - bottom
        elif top >= cy1 - 6:
            gap = top - cy1
        else:
            continue
        if gap > 30:
            continue
        if best is None or gap < best[0]:
            best = (gap, [left, top, right, bottom])
    return best[1] if best else None


def fallback_table_rect(page, cap):
    blocks = text_blocks(page)
    cx0, cy0, cx1, cy1 = cap["bbox"]
    best = None
    for b in blocks:
        bx0, by0, bx1, by1 = b[:4]
        if abs(bx0 - cx0) < 1 and abs(by0 - cy0) < 1:
            continue
        if by0 >= cy1 - 2:
            gap = by0 - cy1
        elif by1 <= cy0 + 2:
            gap = cy0 - by1
        else:
            continue
        if gap > 40:
            continue
        txt = b[4]
        # Reject prose: a table row has short cells and little punctuation.
        words = txt.split()
        if len(words) > 40 or txt.count(".") > 6:
            continue
        area = (bx1 - bx0) * (by1 - by0)
        score = gap - area / 1e5
        if best is None or score < best[0]:
            best = (score, [bx0, by0, bx1, by1])
    return best[1] if best else None


def caption_key(c: dict) -> str:
    return f"{c['kind']}-{c['label']}"


def find_visual_near(page, cap: dict, min_area_ratio: float):
    """Collect every visual cluster that belongs to this caption.

    Multi-panel figures place several independent images side by side on the
    same baseline; picking only the largest one silently drops a panel, so the
    candidates are grouped by vertical band and the band above the caption is
    returned as one bbox.
    """
    page_area = abs(page.rect)
    cx0, cy0, cx1, cy1 = cap["bbox"]
    cap_top, cap_bottom = cy0, cy1

    clusters = []
    for info in page.get_image_info():
        bbox = info["bbox"]
        clusters.append((list(bbox), "image"))
    try:
        drawings = page.get_drawings()
    except Exception:
        drawings = []
    for d in drawings:
        r = d["rect"]
        clusters.append(([r.x0, r.y0, r.x1, r.y1], "vector"))

    # Keep clusters that are big enough, sit clear of the caption, and fall
    # inside the caption's horizontal band.
    above, below = [], []
    for bbox, kind in clusters:
        x0, y0, x1, y1 = bbox
        w, h = x1 - x0, y1 - y0
        if w < 30 or h < 25:
            continue
        if (w * h) / page_area < min_area_ratio / 2:
            continue
        if x1 < cx0 - 40 or x0 > cx1 + 40:
            continue
        if y1 <= cap_top + 8:
            above.append([x0, y0, x1, y1])
        elif y0 >= cap_bottom - 8:
            below.append([x0, y0, x1, y1])

    def band_boxes(items, limit):
        """Group items whose vertical spans overlap; return the closest band."""
        if not items:
            return None
        items.sort(key=lambda r: r[1])
        bands = []
        for r in items:
            placed = False
            for band in bands:
                # Same visual row: vertical spans must overlap.
                if min(band[3], r[3]) - max(band[1], r[1]) > -30:
                    band[0] = min(band[0], r[0]); band[1] = min(band[1], r[1])
                    band[2] = max(band[2], r[2]); band[3] = max(band[3], r[3])
                    placed = True
                    break
            if not placed:
                bands.append(list(r))
        bands.sort(key=lambda b: abs((b[1] if b[1] > cap_top else cap_top) - limit))
        return bands

    above_bands = band_boxes(above, cap_top)
    if above_bands:
        chosen = max(above_bands, key=lambda b: (b[2] - b[0]) * (b[3] - b[1]))
        # A wide band that also holds a much smaller neighbour on the same row
        # is the multi-panel case: keep the whole row.
        row = [b for b in above_bands
               if min(chosen[3], b[3]) - max(chosen[1], b[1]) > -30]
        box = list(chosen)
        for b in row:
            box[0] = min(box[0], b[0]); box[1] = min(box[1], b[1])
            box[2] = max(box[2], b[2]); box[3] = max(box[3], b[3])
        return {"bbox": box, "kind": "image", "side": "above",
                "gap": max(0.0, cap_top - box[3]), "score": 1.0}
    if below:
        box = below[0]
        for b in below[1:]:
            box = [min(box[0], b[0]), min(box[1], b[1]),
                   max(box[2], b[2]), max(box[3], b[3])]
        return {"bbox": box, "kind": "image", "side": "below",
                "gap": max(0.0, box[1] - cap_bottom), "score": 1.0}
    return None


def merge_cluster(page, seed: list[float], min_area_ratio: float) -> list[float]:
    """Grow a seed bbox to swallow neighbouring sub-panels and their labels."""
    try:
        drawings = page.get_drawings()
    except Exception:
        drawings = []
    rects = [[d["rect"].x0, d["rect"].y0, d["rect"].x1, d["rect"].y1] for d in drawings]
    for info in page.get_image_info():
        rects.append(list(info["bbox"]))

    x0, y0, x1, y1 = seed
    changed = True
    while changed:
        changed = False
        for rx0, ry0, rx1, ry1 in rects:
            w, h = rx1 - rx0, ry1 - ry0
            if (w * h) / abs(page.rect) < min_area_ratio / 4:
                continue
            # Horizontal overlap required; vertical gap tolerance ~ 55pt.
            h_overlap = min(x1, rx1) - max(x0, rx0)
            if h_overlap <= 0:
                continue
            if ry1 < y0 - 55 or ry0 > y1 + 55:
                continue
            nx0, ny0 = min(x0, rx0), min(y0, ry0)
            nx1, ny1 = max(x1, rx1), max(y1, ry1)
            if [nx0, ny0, nx1, ny1] != [x0, y0, x1, y1]:
                x0, y0, x1, y1 = nx0, ny0, nx1, ny1
                changed = True
    return [x0, y0, x1, y1]


def pad(bbox, page_rect, pady=6.0, padx=4.0):
    return pymupdf.Rect(
        max(page_rect.x0, bbox[0] - padx),
        max(page_rect.y0, bbox[1] - pady),
        min(page_rect.x1, bbox[2] + padx),
        min(page_rect.y1, bbox[3] + pady),
    )


def render_region(page, rect, zoom: float):
    pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), clip=rect, alpha=False)
    return pix


def include_panel_titles(page, box: list[float], cap: dict) -> list[float]:
    """Grow the crop upward over short headings that sit just above the art.

    Sub-panel labels such as "Multi-Head Attention" are separate text blocks
    above each panel; excluding them leaves chopped title fragments.
    """
    x0, y0, x1, y1 = box
    top, left, right = y0, x0, x1
    for b in text_blocks(page):
        bx0, by0, bx1, by1 = b[:4]
        text = b[4].strip()
        if len(text) > 90 or by1 > y0 + 6:
            continue
        # A heading may start slightly outside the art box (centred over a
        # narrower panel), so allow a generous horizontal margin.
        if bx1 < x0 - 120 or bx0 > x1 + 120:
            continue
        if y0 - by1 > 40:
            continue
        if by0 < cap["bbox"][1] - 2:
            top = min(top, by0)
            left = min(left, bx0)
            right = max(right, bx1)
    if top < y0 or left < x0 or right > x1:
        return [left - 2, top - 2, right + 2, y1]
    return box


def include_axis_furniture(page, box: list[float], cap: dict) -> list[float]:
    """Extend the box over axis titles and tick labels around a chart.

    Axis text is drawn as text, never as a drawing, so a band built only from
    graphics stops at the plot frame and the chart ships with no axis labels.
    Grow the box outward while nearby text looks like chart furniture
    (short tokens, numbers, no sentence punctuation) and stays outside the
    caption.
    """
    x0, y0, x1, y1 = box
    cap_top = cap["bbox"][1]
    margin = 42.0

    changed = True
    rounds = 0
    while changed and rounds < 4:
        changed = False
        rounds += 1
        for b in text_blocks(page):
            bx0, by0, bx1, by1 = b[:4]
            text = b[4].strip()
            if not text:
                continue
            # Chart furniture is short; prose is not.
            if len(text) > 60 or text.count(".") > 2:
                continue
            # Never grow into the caption or past it.
            if by1 > cap_top - 2:
                continue
            # Must touch the current box (within the margin) to be furniture.
            near_x = bx1 >= x0 - margin and bx0 <= x1 + margin
            near_y = by1 >= y0 - margin and by0 <= y1 + margin
            if not (near_x and near_y):
                continue
            # And must actually be adjacent, not just somewhere on the page.
            outside = (by1 <= y0 + 12 or by0 >= y1 - 12
                       or bx1 <= x0 + 12 or bx0 >= x1 - 12)
            if not outside:
                continue
            nx0, ny0 = min(x0, bx0), min(y0, by0)
            nx1, ny1 = max(x1, bx1), max(y1, by1)
            if [nx0, ny0, nx1, ny1] != [x0, y0, x1, y1]:
                x0, y0, x1, y1 = nx0, ny0, nx1, ny1
                changed = True
    return [x0, y0, x1, y1]


def crop_baked_caption(page, bbox: list[float], cap: dict) -> list[float]:
    """Trim the paper's own caption out of a crop when it sits inside the art.

    Attention/heat-map figures often embed their legend text inside the image
    area; keeping it makes the picture tiny once it is scaled to the slide.
    A caption can be split over several text blocks, so cut at the topmost one.
    """
    x0, y0, x1, y1 = bbox
    cx0, cy0, cx1, cy1 = cap["bbox"]
    overlaps_x = cx0 >= x0 - 20 and cx1 <= x1 + 20
    near_bottom = cy0 > y0 + (y1 - y0) * 0.55
    if not (overlaps_x and near_bottom):
        return bbox

    cut = cy0
    # A caption can be split over several blocks.  Only prose-like blocks count:
    # rotated axis words inside an attention figure are single tokens, and
    # treating them as caption text would cut away a whole sub-panel.
    art_w = x1 - x0
    for b in text_blocks(page):
        bx0, by0, bx1, by1 = b[:4]
        text = b[4].strip()
        # A real journal caption runs essentially the full measure of the art.
        # A wide in-figure panel label ("Pre-training  Fine-Tuning") does not,
        # and cutting on it decapitates the figure.
        spans_art = (bx1 - bx0) >= 0.75 * art_w
        is_prose = len(text) >= 50
        if not (spans_art or is_prose):
            continue
        if bx0 >= x0 - 20 and bx1 <= x1 + 20 and by0 > y0 + (y1 - y0) * 0.5:
            if by0 < cut:
                cut = by0
    new_y1 = cut - 4
    if new_y1 - y0 > 80:
        return [x0, y0, x1, new_y1]
    return bbox


def cluster_text_bars(page, seed: list[float]) -> list[float]:
    """Merge thin, side-by-side text tokens into one tight bounding box.

    Vector figures drawn from hundreds of tiny glyph boxes (attention maps with
    rotated words) otherwise leave a ragged border, because each token is its
    own drawing primitive rather than a single large rect.
    """
    try:
        drawings = page.get_drawings()
    except Exception:
        return seed
    x0, y0, x1, y1 = seed
    pad_x, pad_y = 12.0, 12.0
    xs0, ys0, xs1, ys1 = [], [], [], []
    for d in drawings:
        r = d["rect"]
        if r.x0 < x0 - pad_x or r.x1 > x1 + pad_x:
            continue
        if r.y0 < y0 - pad_y or r.y1 > y1 + pad_y:
            continue
        xs0.append(r.x0); ys0.append(r.y0)
        xs1.append(r.x1); ys1.append(r.y1)
    if not xs0:
        return seed
    return [min(xs0), min(ys0), max(xs1), max(ys1)]


def build_markdown(doc, captions, visuals) -> str:
    by_page: dict[int, list[dict]] = {}
    for v in visuals:
        by_page.setdefault(v["page"], []).append(v)

    lines = ["# Extracted paper", ""]
    lines.append(f"- pages: {len(doc)}")
    meta = doc.metadata or {}
    if meta.get("title"):
        lines.append(f"- pdf-title: {clean(meta['title'])}")
    lines.append("")

    lines.append("## Captions found")
    for c in captions:
        key = caption_key(c)
        hit = "no visual" if not any(v["key"] == key for v in visuals) else "image ok"
        lines.append(f"- p{c['page']+1} **{key}** ({hit}): {c['caption'][:300]}")
    lines.append("")

    for pno in range(len(doc)):
        page = doc[pno]
        text = page.get_text("text")
        lines.append(f"## Page {pno + 1}")
        for v in by_page.get(pno, []):
            lines.append(f"> [{v['key'].upper()}] extracted to {v['path']} "
                         f"({v['width']}x{v['height']}px)")
        # Mark caption positions inline so the reader can pair text and figure.
        for c in [c for c in captions if c["page"] == pno]:
            text = text.replace(c["caption"][:60], f"[[{caption_key(c).upper()}]] " + c["caption"][:60], 1)
        lines.append(text.strip())
        lines.append("")
    return "\n".join(lines)


def extract(pdf_path: Path, out_dir: Path, dpi: int, min_area: float) -> dict:
    doc = pymupdf.open(pdf_path)
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / "figures").mkdir(exist_ok=True)
    (out_dir / "tables").mkdir(exist_ok=True)
    (out_dir / "pages").mkdir(exist_ok=True)

    zoom = dpi / 72.0
    captions = collect_captions(doc)

    # Table captions must be paired per page, otherwise two stacked tables can
    # both claim the same grid.
    table_boxes: dict[int, list[float]] = {}
    by_page_caps: dict[int, list[tuple[int, dict]]] = {}
    for idx, c in enumerate(captions):
        if c["kind"] == "table":
            by_page_caps.setdefault(c["page"], []).append((idx, c))
    for pno, items in by_page_caps.items():
        assigned = assign_tables(doc[pno], [c for _, c in items])
        for local_i, (idx, _) in enumerate(items):
            if local_i in assigned:
                table_boxes[idx] = assigned[local_i]

    visuals: list[dict] = []
    seen: set[str] = set()

    for idx, cap in enumerate(captions):
        key = caption_key(cap)
        page = doc[cap["page"]]

        if cap["kind"] == "table":
            bbox = table_boxes.get(idx)
            if bbox is None:
                continue
        else:
            seed = find_visual_near(page, cap, min_area)
            if seed is None:
                continue
            bbox = merge_cluster(page, seed["bbox"], min_area)
            bbox = cluster_text_bars(page, bbox)
            bbox = include_panel_titles(page, bbox, cap)
            bbox = include_axis_furniture(page, bbox, cap)
            bbox = crop_baked_caption(page, bbox, cap)

        if key in seen:
            key = f"{key}-p{cap['page'] + 1}"
        seen.add(key)

        rect = pad(bbox, page.rect)
        # Padding is applied after the art bbox is known, so re-assert the
        # caption cut here or the pad would poke back into the caption text.
        cap_top = cap["bbox"][1]
        if rect.y0 < cap_top and rect.y1 > cap_top - 2:
            rect.y1 = min(rect.y1, cap_top - 5)
        if rect.width < 20 or rect.height < 20:
            continue
        pix = render_region(page, rect, zoom)
        sub = "tables" if cap["kind"] == "table" else "figures"
        name = f"{cap['kind'][:3]}-{cap['label']}.png"
        rel = f"{sub}/{name}"
        pix.save(out_dir / rel)
        visuals.append({
            "key": key,
            "kind": cap["kind"],
            "label": cap["label"],
            "page": cap["page"],
            "caption": cap["caption"],
            "bbox": [round(v, 2) for v in bbox],
            "path": rel,
            "width": pix.width,
            "height": pix.height,
            "aspect": round(pix.width / pix.height, 4) if pix.height else 0,
        })

    # Full page renders are the documented fallback when a bbox misses.
    for pno in range(len(doc)):
        pix = doc[pno].get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
        pix.save(out_dir / "pages" / f"page-{pno + 1}.png")

    manifest = {
        "source": str(pdf_path),
        "pages": len(doc),
        "dpi": dpi,
        "metadata": {k: v for k, v in (doc.metadata or {}).items() if v},
        "figures": visuals,
        "caption_count": len(captions),
    }
    (out_dir / "figures.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    (out_dir / "paper.md").write_text(
        build_markdown(doc, captions, visuals), encoding="utf-8")
    doc.close()
    return manifest


def table_region_from_text(page, cap):
    """Fallback: a table body is a dense block of text next to the caption.

    Kept for pages where neither find_tables() nor the ruling lines resolve a
    grid.  Returns a Rect or None.
    """
    blocks = text_blocks(page)
    blocks.sort(key=lambda b: b[1])
    cx0, cy0, cx1, cy1 = cap["bbox"]
    best = None
    for b in blocks:
        bx0, by0, bx1, by1 = b[:4]
        if abs(by0 - cy1) < 2 and abs(by1 - cy0) < 2:
            continue
        if by0 >= cy1 - 4:
            gap = by0 - cy1
        elif by1 <= cy0 + 4:
            gap = cy0 - by1
        else:
            continue
        if gap > 90:
            continue
        area = (bx1 - bx0) * (by1 - by0)
        if best is None or gap < best[0]:
            best = (gap, [bx0, by0, bx1, by1])
    if best is None:
        return None
    return pymupdf.Rect(best[1])


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("pdf")
    ap.add_argument("--out", required=True)
    ap.add_argument("--dpi", type=int, default=300,
                    help="rasterisation dpi for crops (300 keeps axis labels crisp)")
    ap.add_argument("--min-area", type=float, default=0.0025,
                    help="ignore visual clusters smaller than this fraction of a page")
    args = ap.parse_args(argv)

    pdf = Path(args.pdf)
    if not pdf.exists():
        print(f"error: {pdf} not found", file=sys.stderr)
        return 2
    m = extract(pdf, Path(args.out), args.dpi, args.min_area)
    print(f"pages={m['pages']} captions={m['caption_count']} figures={len(m['figures'])}")
    for f in m["figures"]:
        print(f"  {f['key']:<12} p{f['page']+1:<3} {f['width']}x{f['height']}  {f['path']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
