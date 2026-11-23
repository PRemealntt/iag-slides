# IAG Transform design system (from the official template)

## Canvas and grid
- 16:9, 13.333 × 7.5 in. Safe margins: left/right 0.39 in (content spans x 0.39 → 12.94, width 12.55), top 0.19 in, bottom content limit y ≈ 6.7; footnote at y 6.86; slide number at x 12.2, y 6.95.
- 4-column grid: use 3, 4, 6 or 8 equal columns with 0.2 in gutters. Gaps between blocks 0.15–0.25 in; keep gutters constant within a slide.
- Title placeholder: x 0.39, y 0.19, 32 pt Arial Bold, asphalt. Orange subtitle: x 0.39, y 0.88, ~20 pt bold orange (provided by the layout). Content starts at y ≈ 1.70.

## Palette (theme)
| Role | Name | Hex |
|---|---|---|
| Primary dark / text | Asphalt | `233038` |
| Secondary dark (chevrons, 2nd cards) | Asphalt −1 | `415057` |
| Captions, secondary numbers | Asphalt −2 | `819095` |
| Lines, borders | Asphalt −3 | `C1D1D4` |
| Accent (highlights, icons, key numbers) | Orange | `F15D22` |
| Orange tints | Orange −2 / −3 | `F9A477` / `FDD9C4` |
| Light primary | Light blue | `E0F2F3` |
| Panels on white slides | Panel | `F2F7F8` |
| Black / white | | `1D1D1B` / `FFFFFF` |
Constants exist in `iag_deck.py` (`ASPHALT`, `ORANGE`, `PANEL`, ...). Do not use other colours. For series in charts use orange (focus) and asphalt (rest), then asphalt −2 and orange −2.

## Typography (Arial)
| Element | Size | Weight |
|---|---|---|
| Slide title | 32 | Bold (placeholder) |
| Orange subtitle / key message | 20 | Bold, orange |
| Panel header (white on asphalt) | 13 | Bold |
| Card title | 11–12.5 | Bold |
| Big number | 18–28 | Bold, orange or white on dark |
| KPI value / KPI caption | 20 / 9 caps | Bold |
| Body in cards | 10–12 | Regular, **bold** for the key phrase |
| Chart labels | 10–11 | Bold for values |
| Source / footnote | 8 | Regular, grey |
Highlight inside body text with bold, or bold orange for the one number or phrase that matters (`{{text}}` markup). Left-aligned, ragged right; centre only big numbers in white pills.

## Shapes
- Rounded rectangles; corner radius ratio 0.03–0.06 for panels/cards, 0.12–0.18 for headers, chips and icon tiles. No shadows, no gradients, no outlines on dark cards; white cards get a 0.75 pt `C1D1D4` line.
- Dark card on white slide: fill asphalt, white text. Light panel: `F2F7F8` with an asphalt header strip. Highlight column/cell: orange (header) or `FDD9C4` (cells).
- Information slides: plain white background; dark appears only in cards, bands and headers. Cover, dividers and closing slide keep the template's own look.

## Do
- One message per slide; the orange subtitle is the answer, the title is the topic.
- Number cards 01, 02… when order or count matters.
- Make the one important figure the biggest thing on its card.
- Keep similar elements the same size and aligned; consistent gaps.

## Don't
- No accent lines under titles, no decorative bars or stripes, no clip-art, no WordArt, no new fonts, no stretched logos or photos.
- No text-only slides unless it is a statement slide; add a chart, KPI, icon or card structure.
- No more than ~45 words in a card body; split instead of shrinking text.
