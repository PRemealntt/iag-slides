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

## Portable Python (Windows)
`skills/iag-slides/runtime/win/python/` is the official Python *embeddable* package (same version as the dev machine) with the three libraries installed in `Lib/site-packages`, so Windows users install nothing. It is committed to the repo (~50 MB). `python<ver>._pth` enables `Lib\site-packages` and `import site`.
To rebuild or upgrade it: download `python-<ver>-embed-amd64.zip` from python.org, extract to that folder, restore the `._pth` lines (`python<ver>.zip`, `.`, `Lib\site-packages`, `import site`), then `py -m pip install --target runtime/win/python/Lib/site-packages --no-compile python-pptx Pillow lxml` using the same Python version. Test: `runtime/win/python/python.exe skills/iag-slides/scripts/build_examples.py <tmp>.pptx`. 
## Portable Python (macOS Apple Silicon)
`skills/iag-slides/runtime/macos-arm64/python/` is python-build-standalone (`cpython-<ver>+<tag>-aarch64-apple-darwin-install_only`, from github.com/astral-sh/python-build-standalone) with tkinter/tcl/tk, idlelib, include, tests and pip removed, plus the Mac wheels installed with `pip install --target <...>/lib/python3.15/site-packages --platform macosx_11_0_arm64 --python-version 3.15 --implementation cp --abi cp315 --only-binary=:all: python-pptx Pillow lxml`. It was assembled on Windows (symlinks in `bin/` cannot be created there, so SKILL.md calls `bin/python3.15` directly) and the executable bit is set in git with `git update-index --chmod=+x`. Intel Macs are not covered (fall back to system Python).
