#!/usr/bin/env python3
"""Build a Chinese group-meeting deck (.pptx) from a deck JSON specification.

The specification is written by the agent after reading the paper; this script
owns only the mechanical part - layout, typography, image placement and the
guarantee that every figure keeps the aspect ratio of the original pixels.

Usage:
    python build_pptx.py deck.json [--out talk.pptx] [--theme academic]

Deck JSON (see references/deck-schema.md for the full contract):
{
  "title": "...", "subtitle": "...", "meta": {"venue": "...", "date": "..."},
  "slides": [
    {"layout": "title",  "title": "...", "subtitle": "...", "meta": "..."},
    {"layout": "bullets","title": "...", "bullets": ["...", {"text": "...", "level": 1}]},
    {"layout": "figure", "title": "...", "image": "fig-1.png",
     "caption": "...", "takeaway": "...", "bullets": ["..."]},
    {"layout": "two-col","title": "...", "left": {"bullets": [...]}, "right": {"image": "..."}},
    {"layout": "section","title": "研究背景"},
    {"layout": "summary","title": "...", "bullets": [...]}
  ]
}
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN = Inches(0.62)
# Title block: title 0.50-1.22, subtitle 1.26-1.56, rule 1.62, content below.
TITLE_TOP = Inches(0.50)
SUBTITLE_TOP = Inches(1.26)
RULE_TOP = Inches(1.62)
CONTENT_TOP = Inches(1.82)
CONTENT_BOTTOM = Inches(6.86)

FONT_CJK = "微软雅黑"
FONT_LATIN = "Calibri"
FONT_MONO = "Consolas"

THEMES = {
    "academic": {
        "bg": RGBColor(0xFF, 0xFF, 0xFF),
        "ink": RGBColor(0x1A, 0x1A, 0x1A),
        "muted": RGBColor(0x5F, 0x6B, 0x7A),
        "accent": RGBColor(0x1F, 0x4E, 0x79),
        "accent_soft": RGBColor(0xE8, 0xEF, 0xF7),
        "rule": RGBColor(0xC8, 0xD3, 0xE0),
    },
    "dark": {
        "bg": RGBColor(0x14, 0x18, 0x1F),
        "ink": RGBColor(0xF2, 0xF4, 0xF7),
        "muted": RGBColor(0x9A, 0xA6, 0xB5),
        "accent": RGBColor(0x62, 0xB0, 0xF5),
        "accent_soft": RGBColor(0x22, 0x2A, 0x35),
        "rule": RGBColor(0x33, 0x3D, 0x4A),
    },
}


class DeckError(RuntimeError):
    pass


def set_font(run, size, bold=False, color=None, mono=False):
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.name = FONT_MONO if mono else FONT_LATIN
    if color is not None:
        f.color.rgb = color
    # East Asian font must be set explicitly or PowerPoint substitutes it.
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}" + tag.split(":")[1])
        if el is None:
            from lxml import etree
            el = etree.SubElement(rPr, "{http://schemas.openxmlformats.org/drawingml/2006/main}" + tag.split(":")[1])
        el.set("typeface", FONT_MONO if mono else FONT_CJK)


def add_text(slide, left, top, width, height, text, size, *, bold=False,
             color=None, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             line_spacing=1.25, theme=None):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = line_spacing
    run = p.add_run()
    run.text = text
    set_font(run, size, bold=bold, color=color if color is not None else (theme or {}).get("ink"))
    return box


def add_bullets(slide, left, top, width, height, bullets, theme, *, size=18,
                bullet_color=None, space_after=10):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0

    first = True
    for item in bullets:
        if isinstance(item, str):
            text, level = item, 0
        else:
            text, level = item.get("text", ""), int(item.get("level", 0))
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.level = min(level, 4)
        p.space_after = Pt(space_after if level == 0 else space_after * 0.5)
        p.line_spacing = 1.3
        marker = "· " if level == 0 else "    " * level + "– "
        r0 = p.add_run()
        r0.text = marker
        set_font(r0, size - level * 2, bold=False,
                 color=bullet_color or theme["accent"])
        r1 = p.add_run()
        r1.text = text
        set_font(r1, size - level * 2, color=theme["ink"])
    return box


def add_rule(slide, left, top, width, theme, height=Pt(1.2)):
    from pptx.enum.shapes import MSO_SHAPE
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, int(height))
    shp.fill.solid()
    shp.fill.fore_color.rgb = theme["rule"]
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def add_accent_bar(slide, left, top, theme, height=Inches(0.42), width=Pt(5)):
    from pptx.enum.shapes import MSO_SHAPE
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, int(width), int(height))
    shp.fill.solid()
    shp.fill.fore_color.rgb = theme["accent"]
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def add_title(slide, title, theme, *, subtitle=None):
    add_accent_bar(slide, MARGIN, TITLE_TOP + Inches(0.06), theme,
                   height=Inches(0.62))
    add_text(slide, MARGIN + Inches(0.20), TITLE_TOP,
             SLIDE_W - 2 * MARGIN - Inches(0.20), Inches(0.72),
             title, 27, bold=True, color=theme["accent"], anchor=MSO_ANCHOR.MIDDLE)
    if subtitle:
        add_text(slide, MARGIN + Inches(0.20), SUBTITLE_TOP,
                 SLIDE_W - 2 * MARGIN - Inches(0.20), Inches(0.30),
                 subtitle, 13, color=theme["muted"])
    add_rule(slide, MARGIN, RULE_TOP, SLIDE_W - 2 * MARGIN, theme)


def add_page_number(slide, index, total, theme):
    add_text(slide, SLIDE_W - MARGIN - Inches(1.2), SLIDE_H - Inches(0.46),
             Inches(1.2), Inches(0.28), f"{index} / {total}", 10,
             color=theme["muted"], align=PP_ALIGN.RIGHT)


def fit_image(slide, image_path, left, top, max_w, max_h):
    """Place an image inside the box without ever distorting it."""
    with Image.open(image_path) as im:
        w_px, h_px = im.size
    if not w_px or not h_px:
        raise DeckError(f"unreadable image: {image_path}")
    aspect = w_px / h_px
    box_aspect = max_w / max_h
    if aspect >= box_aspect:
        width = max_w
        height = Emu(int(max_w / aspect))
    else:
        height = max_h
        width = Emu(int(max_h * aspect))
    x = Emu(int(left + (max_w - width) / 2))
    y = Emu(int(top + (max_h - height) / 2))
    return slide.shapes.add_picture(str(image_path), x, y, width=width, height=height)


def blank(prs, theme):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = theme["bg"]
    return slide


# ---------------------------------------------------------------- layouts

def layout_title(prs, spec, theme, ctx):
    slide = blank(prs, theme)
    add_text(slide, MARGIN, Inches(2.15), SLIDE_W - 2 * MARGIN, Inches(1.7),
             spec.get("title", ""), 34, bold=True, color=theme["ink"],
             line_spacing=1.2)
    if spec.get("subtitle"):
        add_text(slide, MARGIN, Inches(3.85), SLIDE_W - 2 * MARGIN, Inches(0.9),
                 spec["subtitle"], 17, color=theme["muted"], line_spacing=1.35)
    add_rule(slide, MARGIN, Inches(4.95), Inches(2.4), theme, height=Pt(3))
    meta = spec.get("meta") or ctx.get("meta") or ""
    if meta:
        add_text(slide, MARGIN, Inches(5.25), SLIDE_W - 2 * MARGIN, Inches(1.0),
                 meta, 14, color=theme["muted"], line_spacing=1.4)
    return slide


def layout_section(prs, spec, theme, ctx):
    slide = blank(prs, theme)
    add_text(slide, MARGIN, Inches(2.85), SLIDE_W - 2 * MARGIN, Inches(1.0),
             spec.get("title", ""), 32, bold=True, color=theme["accent"])
    if spec.get("subtitle"):
        add_text(slide, MARGIN, Inches(3.90), SLIDE_W - 2 * MARGIN, Inches(0.6),
                 spec["subtitle"], 16, color=theme["muted"])
    add_rule(slide, MARGIN, Inches(4.66), Inches(2.0), theme, height=Pt(2))
    return slide


def estimate_bullet_height(bullets, width_emu, size_pt, space_after=10):
    """Rough text height so sparse slides can be centred and dense ones shrunk."""
    # CJK glyphs are ~1.0em wide, latin ~0.55em; assume a mixed 0.85em average.
    em_width = size_pt * 0.85 * 12700  # EMU per character
    per_line = max(1, int(width_emu / em_width))
    total = 0.0
    for item in bullets:
        text = item if isinstance(item, str) else item.get("text", "")
        level = 0 if isinstance(item, str) else int(item.get("level", 0))
        lines = max(1, -(-len(text) // per_line))
        line_h = (size_pt - level * 2) * 1.3 * 12700
        total += lines * line_h + (space_after if level == 0 else space_after * 0.5) * 12700
    return total


def add_bullets_auto(slide, left, top, width, height, bullets, theme,
                     *, max_size=20, min_size=14, space_after=10):
    """Pick a font size that fills the frame without overflowing, then centre."""
    size = max_size
    while size > min_size:
        est = estimate_bullet_height(bullets, width, size, space_after)
        if est <= height * 0.94:
            break
        size -= 1
    est = estimate_bullet_height(bullets, width, size, space_after)
    if est < height * 0.72:
        top += int((height - est) / 2)
        height = int(est * 1.08)
    return add_bullets(slide, left, top, width, height, bullets, theme,
                       size=size, space_after=space_after)


def layout_bullets(prs, spec, theme, ctx):
    slide = blank(prs, theme)
    add_title(slide, spec.get("title", ""), theme, subtitle=spec.get("subtitle"))
    bullets = spec.get("bullets") or []
    if not bullets:
        raise DeckError(f"bullets slide {spec.get('title')!r} has no bullets")
    add_bullets_auto(slide, MARGIN, CONTENT_TOP, SLIDE_W - 2 * MARGIN,
                     CONTENT_BOTTOM - CONTENT_TOP, bullets, theme,
                     max_size=int(spec.get("font_size", 21)))
    return slide


def layout_figure(prs, spec, theme, ctx):
    slide = blank(prs, theme)
    add_title(slide, spec.get("title", ""), theme, subtitle=spec.get("subtitle"))

    img = ctx["resolve_image"](spec.get("image"))
    side_bullets = spec.get("bullets") or []
    takeaway = spec.get("takeaway")
    caption = spec.get("caption")

    with Image.open(img) as im:
        aspect = im.size[0] / im.size[1]

    full_w = SLIDE_W - 2 * MARGIN
    full_h = int(CONTENT_BOTTOM - CONTENT_TOP)

    # A tall figure gains nothing from a side column: it would shrink to a
    # ribbon.  Put the bullets underneath instead and let the art use the width.
    tall = aspect < 1.05
    if side_bullets and tall:
        band = min(Inches(1.75), Inches(0.34) * len(side_bullets) + Inches(0.30))
        media_h = int(full_h - band - Inches(0.12))
        fit_image(slide, img, MARGIN, CONTENT_TOP, full_w, media_h)
        add_bullets(slide, MARGIN, CONTENT_TOP + media_h + Inches(0.14),
                    full_w, int(band), side_bullets, theme, size=14,
                    space_after=5)
        if caption:
            add_text(slide, MARGIN, SLIDE_H - Inches(0.40), full_w, Inches(0.24),
                     caption, 9, color=theme["muted"])
        return slide

    # A figure with no side matter that is taller than the content area should
    # use the full height and stay centred horizontally, not be shrunk to fit
    # the width.  Lower the takeaway beside it so the art gets the full height.
    if not side_bullets and aspect < 1.0:
        # A tall figure leaves a wide empty gutter; give that gutter a card so
        # the whitespace reads as deliberate rather than as a missing block.
        media_h = int((SLIDE_W - 2 * MARGIN) * 0.62 / aspect)
        media_h = min(int(full_h), media_h)
        media_w = int(media_h * aspect)
        fit_image(slide, img, MARGIN, CONTENT_TOP, media_w, media_h)

        text_left = MARGIN + media_w + Inches(0.50)
        text_w = int(SLIDE_W - MARGIN - text_left)
        if text_w < Inches(2.6):
            text_left = MARGIN + media_w + Inches(0.30)
            text_w = int(SLIDE_W - MARGIN - text_left)

        card_top = CONTENT_TOP + Inches(0.30)
        card_h = Inches(3.15) if takeaway else Inches(1.6)
        from pptx.enum.shapes import MSO_SHAPE
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, text_left,
                                      card_top, text_w, int(card_h))
        card.adjustments[0] = 0.06
        card.fill.solid()
        card.fill.fore_color.rgb = theme["accent_soft"]
        card.line.color.rgb = theme["rule"]
        card.line.width = Pt(0.75)
        card.shadow.inherit = False

        inner = text_left + Inches(0.34)
        inner_w = text_w - Inches(0.68)
        add_text(slide, inner, card_top + Inches(0.26), inner_w, Inches(0.30),
                 "核 心 结 论", 12, bold=True, color=theme["accent"])
        if takeaway:
            add_text(slide, inner, card_top + Inches(0.66), inner_w,
                     int(card_h) - Inches(0.9), takeaway, 18, bold=True,
                     color=theme["ink"], line_spacing=1.45)
        if caption:
            add_text(slide, text_left, card_top + int(card_h) + Inches(0.26),
                     text_w, Inches(1.0), caption, 9, color=theme["muted"],
                     line_spacing=1.18)
        return slide

    if side_bullets:
        media_w = int(full_w * 0.60)
        fit_image(slide, img, MARGIN, CONTENT_TOP, media_w, full_h)

        text_left = MARGIN + media_w + Inches(0.45)
        text_w = SLIDE_W - MARGIN - text_left
        add_text(slide, text_left, CONTENT_TOP, text_w, Inches(0.30),
                 "要点", 13, bold=True, color=theme["accent"])
        # The caption owns a reserved strip at the bottom of the column; the
        # bullets get whatever is left and shrink to fit it, so a five-line
        # list can no longer run into the caption.
        cap_h = Inches(1.0) if caption else 0
        bullets_top = CONTENT_TOP + Inches(0.40)
        bullets_h = int(CONTENT_BOTTOM - bullets_top - cap_h)
        add_bullets_auto(slide, text_left, bullets_top, text_w, bullets_h,
                         side_bullets, theme, max_size=15, min_size=11,
                         space_after=8)
        if caption:
            add_text(slide, text_left, CONTENT_BOTTOM - Inches(0.95), text_w,
                     Inches(0.95), caption, 9, color=theme["muted"],
                     line_spacing=1.18)
    else:
        band = 0
        if takeaway:
            band += 0.56
        if caption:
            band += 0.26
        media_h = int(full_h - Inches(band + (0.14 if band else 0)))
        fit_image(slide, img, MARGIN, CONTENT_TOP, full_w, media_h)

        y = CONTENT_TOP + media_h + Inches(0.14)
        if takeaway:
            add_accent_bar(slide, MARGIN, y + Inches(0.03), theme,
                           height=Inches(0.46), width=Pt(4))
            add_text(slide, MARGIN + Inches(0.16), y,
                     full_w - Inches(1.3), Inches(0.48),
                     takeaway, 15, bold=True, color=theme["ink"],
                     anchor=MSO_ANCHOR.MIDDLE)
            y += Inches(0.56)
        if caption:
            add_text(slide, MARGIN, y, full_w - Inches(1.3),
                     Inches(0.26), caption, 9, color=theme["muted"],
                     line_spacing=1.18)
    return slide


def layout_two_col(prs, spec, theme, ctx):
    slide = blank(prs, theme)
    add_title(slide, spec.get("title", ""), theme, subtitle=spec.get("subtitle"))

    gap = Inches(0.5)
    col_w = int((SLIDE_W - 2 * MARGIN - gap) / 2)
    height = int(CONTENT_BOTTOM - CONTENT_TOP)

    left, right = spec.get("left") or {}, spec.get("right") or {}
    for col, (left_pos, spec_col) in enumerate(
            ((MARGIN, left), (MARGIN + col_w + gap, right))):
        if spec_col.get("image"):
            img = ctx["resolve_image"](spec_col["image"])
            cap = spec_col.get("caption")
            media_h = int(height - (Inches(0.44) if cap else 0))
            fit_image(slide, img, left_pos, CONTENT_TOP, col_w, media_h)
            if cap:
                add_text(slide, left_pos, CONTENT_TOP + media_h + Inches(0.10),
                         col_w, Inches(0.34), cap, 9.5, color=theme["muted"])
        elif spec_col.get("bullets"):
            if spec_col.get("heading"):
                add_text(slide, left_pos, CONTENT_TOP, col_w, Inches(0.32),
                         spec_col["heading"], 14, bold=True, color=theme["accent"])
            add_bullets(slide, left_pos, CONTENT_TOP + Inches(0.44), col_w,
                        height, spec_col["bullets"], theme, size=15)
        elif spec_col.get("text"):
            add_text(slide, left_pos, CONTENT_TOP, col_w, height,
                     spec_col["text"], 16, color=theme["ink"], line_spacing=1.4)
    return slide


def layout_summary(prs, spec, theme, ctx):
    slide = blank(prs, theme)
    add_title(slide, spec.get("title", "小结"), theme, subtitle=spec.get("subtitle"))
    bullets = spec.get("bullets") or []
    add_bullets_auto(slide, MARGIN, CONTENT_TOP, SLIDE_W - 2 * MARGIN,
                     CONTENT_BOTTOM - CONTENT_TOP, bullets, theme,
                     max_size=int(spec.get("font_size", 21)))
    return slide


LAYOUTS = {
    "title": layout_title,
    "section": layout_section,
    "bullets": layout_bullets,
    "figure": layout_figure,
    "two-col": layout_two_col,
    "summary": layout_summary,
}


def build(deck: dict, base_dir: Path, out_path: Path, theme_name: str) -> Path:
    theme = THEMES.get(theme_name)
    if theme is None:
        raise DeckError(f"unknown theme {theme_name!r}; choose from {sorted(THEMES)}")

    image_root = (base_dir / deck["image_dir"]).resolve() if deck.get("image_dir") else base_dir

    def resolve_image(name):
        if not name:
            raise DeckError("slide references an image but has no 'image' value")
        p = (image_root / name)
        if not p.exists():
            p = base_dir / name
        if not p.exists():
            raise DeckError(f"image not found: {name} (looked in {image_root})")
        return p

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    slides = deck.get("slides") or []
    if not slides:
        raise DeckError("deck has no slides")

    ctx = {"resolve_image": resolve_image, "meta": deck.get("meta_line", "")}
    for i, spec in enumerate(slides, start=1):
        layout = spec.get("layout", "bullets")
        fn = LAYOUTS.get(layout)
        if fn is None:
            raise DeckError(f"slide {i}: unknown layout {layout!r}; "
                            f"choose from {sorted(LAYOUTS)}")
        slide = fn(prs, spec, theme, ctx)
        if layout not in ("title", "section"):
            add_page_number(slide, i, len(slides), theme)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out_path))
    return out_path


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("deck")
    ap.add_argument("--out", default=None)
    ap.add_argument("--theme", default="academic", choices=sorted(THEMES))
    args = ap.parse_args(argv)

    deck_path = Path(args.deck).resolve()
    deck = json.loads(deck_path.read_text(encoding="utf-8"))
    out = Path(args.out) if args.out else deck_path.with_suffix(".pptx")
    try:
        result = build(deck, deck_path.parent, out, args.theme)
    except DeckError as e:
        print(f"error: {e}", file=sys.stderr)
        return 1
    print(f"wrote {result} ({len(deck.get('slides', []))} slides)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
