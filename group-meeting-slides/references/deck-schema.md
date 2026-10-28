# Deck JSON schema

`build_pptx.py` reads one JSON file. All fields are optional unless marked
**required**. Unknown keys are ignored, so a typo in a key name fails silently —
prefer copying the examples below.

```jsonc
{
  "title": "Attention Is All You Need",      // fallback title / record
  "subtitle": "...",
  "meta_line": "...",                        // default meta for layout "title"
  "image_dir": ".",                          // image paths resolve against this
  "slides": [ { "layout": "...", ... } ]     // required, non-empty
}
```

## Layouts

### `title` — cover

```json
{"layout": "title", "title": "论文标题（英文原题）",
 "subtitle": "一句话中文概括", "meta": "作者 · 会议年份\n汇报人：XXX · 组会"}
```

### `section` — divider between chapters

```json
{"layout": "section", "title": "实验结果", "subtitle": "BLEU 与训练开销"}
```

### `bullets` — text-only slide

```json
{"layout": "bullets", "title": "研究背景", "subtitle": "为什么要放弃 RNN",
 "bullets": [
   "一级要点，整句陈述，句末不加句号",
   {"text": "二级要点，用于展开或限定", "level": 1}
 ],
 "font_size": 21}
```

`font_size` is a ceiling: the builder shrinks text to fit and centres short
lists. Keep to 3–5 top-level bullets; a slide is not a paragraph.

### `figure` — the workhorse

```json
{"layout": "figure", "title": "主要结果：BLEU 与训练成本",
 "image": "tables/tab-2.png",
 "takeaway": "一句话结论，必须含具体数字或机理",
 "bullets": ["可选：右侧分栏要点"],
 "caption": "Table 2: ...（原文图注，照抄）"}
```

The builder chooses the arrangement from the image's aspect ratio and which
fields you supply. `aspect = width / height`, so `< 1.0` means portrait:

| Condition | Result |
|---|---|
| has `bullets`, aspect < 1.05 | image on top, bullets underneath |
| has `bullets`, aspect ≥ 1.05 | image left 60%, bullets right |
| no `bullets`, aspect < 1.0 | image left, takeaway in a card on the right |
| no `bullets`, aspect ≥ 1.0 | image on top, takeaway line + caption below |

So: give `bullets` for a wide figure you have a lot to say about, and give only
`takeaway` for a tall one.

**Near-square figures (aspect 0.9–1.15) look bad in `figure`.** They render
marooned in the middle of the slide with two dead gutters. Use `two-col`
instead: image in `left`, bullets in `right`.

**Side bullets must stay short.** In the bullets-right arrangement the column
fits about 4 top-level lines at 15pt; longer copy is auto-shrunk down to 11pt and
then stops, and the caption strip at the bottom is reserved. If you need more
than 4 top-level bullets, split them across two slides or move them under a
tall figure.

### `two-col` — compare two things

```json
{"layout": "two-col", "title": "base 与 big 模型对比",
 "left":  {"heading": "base", "bullets": ["..."]},
 "right": {"image": "figures/fig-3.png", "caption": "..."}}
```

Each side accepts `image` + `caption`, or `bullets` + optional `heading`, or
`text`.

### `summary` — discussion / wrap-up

```json
{"layout": "summary", "title": "讨论", "subtitle": "供组会讨论的问题",
 "bullets": ["问题一", "问题二"]}
```

## Common mistakes

- `"image"` must include the subdirectory: `"figures/fig-1.png"`, not `"fig-1.png"`.
- **File names and manifest keys differ.** `extract.py` writes `fig-1.png` /
  `tab-1.png`, but `figures.json` keys are `figure-1` / `table-1`. Use the file
  name in `"image"`, the key when looking things up in the manifest.
- `bullets` entries are strings or `{"text": ..., "level": 0|1}` objects. A bare
  object without `text` renders empty.
- `caption` is the **original English caption**. Do not translate it into this
  field; Chinese explanation belongs in `takeaway` or `bullets`.
- `takeaway` must add information. Restating the title wastes the line.
