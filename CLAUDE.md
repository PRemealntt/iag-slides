# iag-slides plugin — project notes for Claude Code

Plugin with one skill (`skills/iag-slides`) that builds PowerPoint decks in the IAG Transform corporate format with python-pptx.

## Layout
- `skills/iag-slides/SKILL.md` — workflow and hard rules (read it before changing behaviour).
- `skills/iag-slides/scripts/iag_deck.py` — component library (Deck, kpi_band, panel, charts, tables, cards, steps, logos...).
- `skills/iag-slides/scripts/build_examples.py` — builds `assets/examples/patterns-example.pptx` (one slide per pattern, INVENTED data).
- `skills/iag-slides/references/` — design-system, patterns, assets, layouts.
- `skills/iag-slides/assets/` — `template.pptx` (official, no slides), `brand/logos` (+ `suppliers/`), `brand/photos`, `fa-icons`, `examples/previews`.

## House rules (do not break)
- Information slides: plain white background. Cover, dividers and closing slide keep the template look.
- Palette/typography only from `references/design-system.md` (Arial; Asphalt, Orange, tints).
- Show the logo of every supplier / group company that is named; logos are loose (no frame) by default; faded + "Not bidding" when not participating. Missing logo → text + tell the user.
- No real company data in this repo: examples, docs and previews use invented figures and generic names. Do not add the original source decks.
- Never edit or rename template layouts.

## After any change to `iag_deck.py`, patterns or components
1. `python skills/iag-slides/scripts/build_examples.py skills/iag-slides/assets/examples/patterns-example.pptx`
2. Render to PDF/JPG (LibreOffice + `pdftoppm -jpeg -r 80`) and look at every slide; refresh `assets/examples/previews/pattern-slide-NN.jpg`.
3. Update `references/patterns.md` / `assets.md` and `SKILL.md` if behaviour or names changed.
4. Bump `version` in `.claude-plugin/plugin.json`.

## Adding a supplier logo
Transparent PNG, ≤ 500 px high, cropped, saved as `skills/iag-slides/assets/brand/logos/suppliers/<kebab-name>.png` (add an alias in `_ALIASES` only for abbreviations).

## Requirements
Python 3 with `python-pptx`, `Pillow`, `lxml`; LibreOffice and `pdftoppm` for rendering checks.
