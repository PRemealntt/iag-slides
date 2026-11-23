# Brand assets bundled with the skill

All paths are relative to the skill folder. Use the helpers in `scripts/iag_deck.py` (`logo()`, `icon_tile()`, `photo_background()`) instead of hand-placing files.

## Logos (`assets/brand/logos/`, PNG 400 px high + SVG)
Pick by background: `_on-light` / `_color` for light backgrounds, `_white` / `_on-dark` for dark or photo backgrounds. Never recolour a logo.

- IAG Transform: `iag-transform_on-light`, `iag-transform_on-dark`, `iag-transform_white`, `iag-transform_black`; icon only: `iag-transform-icon_asphalt`, `_orange`, `_white`, `_black`
- British Airways: `british-airways_color`, `british-airways_white`, `british-airways-icon_color`, `british-airways-icon_white`
- Iberia: `iberia_color`, `iberia_white`, `iberia-icon_color`, `iberia-icon_white`
- Iberia Express: `iberia-express_color`, `iberia-express_white`
- Aer Lingus: `aer-lingus_color`, `aer-lingus_white`, `aer-lingus-icon_color`, `aer-lingus-icon_white`
- Vueling: `vueling_color`, `vueling_white`, `vueling-icon_color`, `vueling-icon_white`
- Level: `level_color`, `level_white`, `level-icon_color`, `level-icon_white`
- IAG Loyalty: `iag-loyalty_color`, `iag-loyalty_white`; IAG Cargo: `iag-cargo_color`, `iag-cargo_white`
- IAG: `iag_dark`, `iag_white`, `iag-international_dark`, `iag-international_white`

### Supplier / partner logos (`assets/brand/logos/suppliers/`, transparent PNG)
`boeing`, `collins-aerospace`, `honeywell`, `lufthansa-technik`, `moog`, `ajw`, `turkish-technic`, `air-france-industries-klm-em`. One version each (colour), meant to be placed loose on white or light panels (no frame; `framed=True` only if asked). Place them with `logo_chip()` / `logo_row()` / `logo_group()`; `find_logo(name)` resolves names and abbreviations ("LHT", "Collins", "AF-KLM E&M", "BA", "Iberia Express"…). Group companies resolve to the `_color` variant on white (`variant="white"` on dark).
**Adding a supplier:** save a cropped transparent PNG (≈ 500 px high max) as `suppliers/<kebab-case-name>.png`; any file there is found by its name without further changes (add an alias in `_ALIASES` in `iag_deck.py` only for abbreviations).

The IAG Transform logo and the small orange arrow icon (top right) already come from the slide layouts. Only add logos for OpCos, suppliers or partners.

## Icons (`assets/fa-icons/`, white PNG 256 px, 88 icons)
White glyphs meant to sit on an orange tile (`icon_tile`). Pass the file name without extension, e.g. `"scale-balanced"`. List them with `ls assets/fa-icons`. Families: finance (coins, sterling-sign, piggy-bank, receipt, wallet), analytics (chart-line, chart-pie, chart-column, arrow-trend-up/down, gauge-high), operations (plane, plane-up, gear, wrench, truck, warehouse, box), people (users, people-group, handshake, headset), process (route, diagram-project, sitemap, rotate, filter, list-check), status (circle-check, triangle-exclamation, flag, bullseye, lightbulb, rocket) and more.
If an icon is missing, render a Font Awesome 6 glyph in white to a 256 px PNG (react-icons `fa6` + sharp) and use its path with `icon_tile` after copying it into `assets/fa-icons`.

## Photos (`assets/brand/photos/`)
Brand photography taken from the template and reference decks (also `photo-sky-clouds`, `photo-wing-banner`, `photo-engine-2`): sunset aircraft (`photo-11`), mountains (`photo-12`), window views (`photo-13`, `photo-14`, `photo-129`), sunset aircraft 2 (`photo-15`, `photo-130`), engines (`photo-119`, `photo-142`), aircraft on stand (`photo-115`, `photo-141`), OpCo tails (`photo-143`), handshake (`photo-152`), cabin (`photo-118`), landing gear (`photo-120`), laptop (`photo-134`).
**Not used on information slides** (they are white). Cover, divider and closing slides already bring their own photography through the layouts; use `photo_background` only if the user explicitly asks. Never stretch or crop to distort.

## Template
`assets/template.pptx` is the official IAG Transform template with all 58 layouts and no slides (guidance slides removed). Theme fonts: Arial Bold headings, Arial body.
