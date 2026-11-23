---
name: iag-slides
description: Builds PowerPoint decks (.pptx) in the IAG Transform corporate format using the official template, brand colours, Arial typography, logos, icons and proven slide patterns. Use when the user asks for slides, a deck, a presentation, an executive update, "hazme unas slides / una presentación / un PowerPoint", "slides con el formato de la empresa / IAG / Transform", or to turn notes, data or a document into branded slides.
---

# IAG Transform slides

Produce `.pptx` decks that follow the IAG Transform template exactly. Build them with python-pptx and the helper library in `scripts/iag_deck.py`, which opens the official template (`assets/template.pptx`, 58 layouts, no slides) and provides ready-made components matching the reference slides.

Read these before building (progressive disclosure, load only what the task needs):
- `references/design-system.md` — palette, typography, sizes, spacing, do/don't. Read always.
- `references/patterns.md` — the catalogue of slide patterns with recipes. Read always; this is where the structure of each slide comes from.
- `references/assets.md` — logos, icons, photos available.
- `references/layouts.md` — all template layouts, only when a pattern needs a non-standard layout (dividers, image title slides, columns, statement slides).
- `assets/examples/patterns-example.pptx` (+ `previews/pattern-slide-NN.jpg`) — one slide per reference pattern with invented placeholder data. Open the JPGs with Read to see the target look; `scripts/build_examples.py` regenerates the deck. Never copy its placeholder text or figures into a real deck.

## Workflow

1. **Clarify the content, not the format.** Identify purpose, audience, number of slides, the data and the message. If the user gave a document, notes or numbers, read them fully first. Ask at most one short question if something blocks the work (for example no data for a chart); otherwise decide and state assumptions at the end.
2. **Plan the storyline.** One message per slide. Typical deck: cover → content slides (executive summary first for 5+ slides) → next steps → closing slide (cover and closing are mandatory; no dividers unless asked). Write each slide as: title `Topic | Subtopic` and an orange subtitle that states the conclusion ("so what") in one sentence.
3. **Pick a pattern per slide** from `references/patterns.md`. Favour the reference patterns whenever the content fits their structure (KPIs + chart + insights, A-vs-B comparison, opportunity with cards and next steps, numbered Q&A grid, scenario comparison, dark photo slide with lever cards, chevron timeline). Vary patterns across the deck. If the content does not fit any of them (long narrative, a single big chart, a dense table, a process with many steps), choose a different layout from `references/layouts.md` or compose from the components, keep the same visual language, and mention the choice to the user in one line.
4. **Pick the Python to run it** (the user must not have to install anything). Use the bundled portable Python when it exists for this OS, otherwise the system Python:
   - Windows: `"<this skill's folder>/runtime/win/python/python.exe" script.py` — it already includes python-pptx, Pillow and lxml. Always quote the path (it contains spaces).
   - Any other case (macOS/Linux, or the `runtime/` folder is missing): use `python3` (or `py` on Windows). If `import pptx, PIL, lxml` fails, run `python3 -m pip install python-pptx Pillow lxml` and tell the user you did.
   Use that same interpreter for `build_examples.py`, the text QA and any other script below.

   **Write the build script** (python, kept outside the output folder), import the library, build slides, save the `.pptx`:
   ```python
   import sys; sys.path.insert(0, "<this skill's folder>/scripts")
   from iag_deck import *
   deck = Deck()
   deck.title_slide("Title", "Subtitle")
   s = deck.content_slide("Topic | Subtopic", "Key message in orange")
   ...components...
   footnote(s, "Source: ...")
   deck.end_slide()
   deck.save("deck.pptx")
   ```
   `scripts/build_examples.py` is a working example that uses every component: read it for call signatures, run it to see the output. Use only python-pptx; never edit the template layouts or master.
5. **QA (required).**
   - Text: the same Python (`-c`) or any extractor to confirm every figure and label is right and no placeholder text remains.
   - Visual: convert to PDF with LibreOffice (`soffice --headless --convert-to pdf`; if bare `soffice` hangs, use the wrapper script available in the environment), rasterise with `pdftoppm -jpeg -r 80`, and look at every slide. Fix text overflow, overlaps, text touching card edges, low contrast, uneven gaps, and charts whose labels collide. Re-render after fixes.
6. **Deliver** the `.pptx` and one or two sentences: what was built, any assumption, and the pattern choices that departed from the references. Offer a concrete next step (for example adding speaker notes).

## Hard rules

- **Information slides have a plain white background (house rule).** Every slide that carries content (text, KPIs, charts, tables, cards) is white: no photos, no dark or coloured full-slide backgrounds, no overlays. Dark appears only as cards, bands and headers on the white slide. Information slides are always white, whatever the pattern looks like elsewhere.
- **Every deck has a cover and a closing slide, always.** The cover (`title_slide`, photo layout) is the first slide and the closing slide (`end_slide`) the last, whatever the length of the deck (even 1–3 content slides). They keep the template look. `Deck.save()` refuses to save without a cover and adds the closing slide if you forgot it.
- **No section dividers and no PowerPoint sections.** Do not create `divider_slide` slides or section headers between slides unless the user explicitly asks for them (e.g. "add an Appendix divider"). Standard deck: cover → information slides (white) → closing slide. `Deck` already strips the template's own leftover PowerPoint sections; never add new ones. Single-statement slides (`statement_slide`) keep their template look and are used only when the user asks for one.

- Colours only from the palette in `design-system.md`; orange is the single accent and is used to highlight, not to fill big areas (except title/divider layouts).
- Arial only (theme fonts). Titles 32 pt bold via the layout placeholder; never restyle placeholders by hand.
- **All slide content is in British English (en-GB), whatever language the user writes in.** Titles, subtitles, labels, chart titles and axes, table text, footnotes and speaker notes are written in English with UK spelling (organisation, programme, optimise, colour, centre, analyse, licence/license rules, dd/mm/yyyy dates, £/€ as given). If the user's notes, data or request are in Spanish or another language, translate the content into British English; keep proper names, product names and quoted terms as they are. Reply to the user in the language they use; only the deck is in English. `Deck.save()` tags all text as en-GB.
- Sentence case everywhere, no full-uppercase paragraphs (short labels and KPI captions in caps are the exception).
- Body text never below 9 pt on cards, 8 pt for source lines; keep slides scannable: short statements, numbers large, one idea per card.
- Every content slide has a source/notes line via `footnote()` when it shows data. Include units and period in chart titles.
- Charts are native PowerPoint charts (editable) created with the helpers; never paste chart images.
- Icons: white glyph on an orange tile (`icon_tile`). Logos: choose the variant that contrasts with the background.
- Do not invent data. If a figure is missing, say so in the reply and leave a clearly marked gap, never a placeholder that looks real.
- **Logos are mandatory when a company is named.** When the content or the user's request mentions a supplier, partner or group company (e.g. "include Boeing and Lufthansa Technik", "compare the OpCos"), show that company's logo, loose (no frame, as in the reference slides), with `logo_chip(slide, "Boeing", x, y, w, h)`, `logo_row(slide, [names], x, y, w)` or, for a supplier panel, `logo_group(slide, "Integrator", [names], x, y, w, dimmed=(non-participants))` (label in italic caps, non-participants faded with a "Not bidding" caption); under chart bars/columns place one logo per category; names, abbreviations and aliases resolve automatically (`find_logo`). Never redraw or recolour a logo. `list_logos()` shows what is bundled. If `logo_chip` reports no logo for a name, it shows the name as text: say so in the reply and tell the user they can add `assets/brand/logos/suppliers/<name>.png` (transparent PNG) to the plugin; do not fetch or invent logos.
- The classification footer ("INTERNAL USE ONLY │ NOT FOR ONWARD DISTRIBUTION") is added to title/divider/end slides by the library; keep it.
