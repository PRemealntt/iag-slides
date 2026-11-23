# Slide patterns

Reference patterns, shown with invented placeholder data in `assets/examples/patterns-example.pptx` (JPG previews in `assets/examples/previews/pattern-slide-NN.jpg`, slide N = pattern P(N); slide 1 is the cover, slides 2–10 are P2–P10, 11 = P12 agenda, 12 = P13, 13 = P14, 14 = P15, 15 = P16, 16 = P17, 17 = P11 logo usage, 18 = divider, 19 = closing slide). They are the preferred structures: use one whenever the content fits it, and adapt counts (3 vs 4 cards, 4 vs 5 steps) rather than forcing content into them. Coordinates are inches; all helpers are in `scripts/iag_deck.py`; `scripts/build_examples.py` shows each recipe running.

## How to choose
| Content looks like… | Pattern |
|---|---|
| A headline number set + a chart + 3 explanations | P2 KPI + chart + insights |
| Two options compared on several criteria | P3 Comparison A vs B |
| An opportunity/finding with chart, callouts, ambition and next steps | P4 Opportunity dashboard |
| 4–6 questions, risks or FAQs | P5 Numbered Q&A grid |
| 3 scenarios/options, each with a mini table | P6 Scenario columns |
| Plan with phases, decision dates and inputs needed | P7 Next steps timeline |
| Drivers + alternatives / levers with big numbers (impact slide) | P8 Dark photo lever cards |
| 5–6 items each with icon and short description | P9 Numbered grid on photo |
| Market/news items, 6 short insights | P10 Insight tiles on photo |
| Meeting agenda (3–6 items) | P12 Agenda |
| Health of a data set / process split into old vs new, with RAG status | P13 Dashboard with RAG pills |
| The same status shown for several systems (progress, pending, mix) | P14 Tracker slide (repeatable) |
| Causes vs mitigations / problems vs actions | P15 Drivers vs changes |
| Usage / adoption metrics by team with totals | P16 Adoption scorecard |
| Programme status: 3 workstreams with proof points + what comes next | P17 Programme summary |
| Anything else | Compose with components, keep the visual language, or use a layout from `layouts.md` |

## P1 Title slide — `deck.title_slide(title, subtitle)`
Window-photo layout `Title Slide [Image 3]` (standard cover; other cover layouts in `layouts.md`). Title short (≤ 6 words), subtitle = audience + context ("Steering committee", "Topic | year approach"). Other title layouts are in `layouts.md`.

## P2 KPI band + chart + insights
Title `Topic | Subtopic`, orange subtitle with the key number.
1. `kpi_band(s, 0.39, 1.70, 12.55, 0.85, title, period, [(LABEL, value, suffix), …])` — 3–4 KPIs.
2. Left (w 7.75): `panel(..., header)` + a native chart (`stacked_pct_chart`, `bar_chart`, `line_chart`) + a one-line bold takeaway under it.
3. Right (w 4.59): `panel(..., "Why X is different")` + `icon_text_rows` with 3 insights (title + 1–2 lines).
4. No footnote or source line on the slide (units and period go in the chart title and KPI band).

## P3 Comparison A vs B
`comparison_table(s, 0.39, 1.75, 12.55, row_labels, ["A", "B"], cells, icons=[...])` — 4–6 criteria; the recommended option is the last column (orange header, light-orange cells, bold text). Close with a light-blue strip (`box(..., LIGHT_BLUE)`) stating the principle or recommendation. Keep cells ≤ 14 words.

## P4 Opportunity dashboard
Left: `panel` + `bar_chart` (highlight the key bar) + two `chip`s (share-of-total facts). Middle: two dark `stat_card`s with one big number each. Right: white card with the ambition/target in 28 pt orange (`€4m → €6m`) and one line of how. Bottom: `steps_band` with 4–5 dated steps, last one highlighted.

## P5 Numbered Q&A grid
Title `Appendix | Topic`; 3 × 2 grid of `numbered_card(..., "01", question, answer, icon=…, dark=True)`; card size ≈ 4.1 × 2.4, gaps 0.17. Answer ≤ 40 words, start with "Yes."/"No."/"Open point:" when it is a yes/no question.

## P6 Scenario columns
Three columns A / B / C: big orange letter + scenario name + one-sentence description + a small native table (header asphalt, body `F2F7F8`, highlight cells orange/`FDD9C4`, 9–10 pt), the recommended one outlined in orange. Under them four numbered lever cards (`numbered_card`, 4 across, h ≈ 1.45).

## P7 Next steps timeline
Top: chevron row (`steps_band` or the chevron `MSO_SHAPE.PENTAGON` shapes with dates in caps, colours asphalt −1, last/goal in orange) with bullet detail boxes under each. Right: a white card with decision gates (committee names with dates in 20 pt orange). Bottom dark band "Inputs to agree" with 3 icon items.

## P8 Lever cards
`deck.content_slide(title, subtitle)` — white background, never the photo. Two sections each introduced by `section_label` (caps + orange rule). Top: three asphalt cards, each with `icon_tile` + title, a white pill with the big number (24 pt bold, centred) and one explanatory line with bold orange key words. Bottom: four `numbered_card`s. Use for impact/“why it matters” slides.

## P9 Numbered grid
White slide; left tile in solid orange with a big "#5" and label; surrounding numbered items 01–05 with icon tile at the right edge of each asphalt card.

## P10 Insight tiles
White slide; title + orange sub-title + 2 × 3 grid of `F2F7F8` tiles (asphalt text), each with a bold headline, 2 short bullets and an icon tile at the right. Use for market insights or news summaries.

## P12 Agenda — `deck.agenda_slide([...], draft=False)`
One white card per item with a big orange number (`01`…) and a 20 pt bold label. 3–6 items; more than 6 → group them.

## P13 Dashboard with RAG pills
Standard `Topic | Subtopic` title (one line) with the orange subtitle; the panels start at y 1.45. Two big white panels (`box` + title + `status_pill(..., "green"/"amber"/"red")`). Left = what is new/healthy; right = what needs work. Building blocks: `kpi_tile` (dark tiles), `split_bar` (2-segment bar: orange = focus, asphalt = rest), `numbered_steps` (dark circles), `simple_table` with `orange_cols` for "today" values, small white callout cards with a big orange `[xx]`, and a dark closing band with a light-orange pill for what is deprioritised. Use `[xx]` for figures still to be confirmed (never invent them). Do not exceed ~6 blocks per panel; if there are more, split into two slides.

## P14 Tracker slide — `deck.tracker_slide(title, subtitle, (total, resolved, progress, pending), [(name, total, pending, resolved, mix), ...], definition)`
Dark metric strip (progress in orange) + two panels with mini stats and a status-mix bar with legend (`Fixed` asphalt, `Not required` grey, `Pending` orange, `Error` red). Repeat the same slide for each system (same order, same colours, one-line subtitle that says what is tracked until when). For 3+ panels use `tracker_panel` at w ≈ 4.0 and drop the legend rows to 3.

## P15 Drivers vs changes
`banner()` with the one-sentence conclusion, then two `numbered_list_panel`s: left in orange (problems/drivers), right in asphalt (actions/positive changes; use asphalt circles, not green). 3–5 rows each, title + one line.

## P16 Adoption scorecard — `content_slide(..., draft=True)` optional
Row of 5 KPI cards (value 24 pt, caps caption, small "x of y") and two side-by-side tables (`simple_table` with `total_rows`, `sub_rows` for the "For information | not included in total" row, `orange_cols` for the metric to watch) each with a team title and two `status_pill`s (neutral + orange). Definitions and source in 8.5 pt below. If a table has more than 9 rows, show the top rows and move the rest to an appendix slide.

## P17 Programme summary
Three `column_card`s (caps label = system, title, big orange figure + caption, 3 bullets on a light panel) and a `next_band` with what comes next and an optional pill for status/ownership. Keep bullets ≤ 12 words.

## P11 Logos of suppliers / group companies
Whenever a supplier or group company is named, show its logo loose (no frame), as in the reference slides: `logo_group(s, "Integrator", [...], x, y, w, cols=3, dimmed=("AJW",))` for a supplier panel with role labels (INTEGRATOR / OEM) and faded non-participants; `logo_chip(s, name, x, y, w, h)` under each bar or column of a chart, or in a card header; `logo_row` for a strip. When two suppliers are combined, stack both logos with a small "+" between them. Typical uses: supplier comparison (one logo per bar/column), market overview, proposal per supplier. See `assets.md` for the bundled list and how to add more.

## P18 Dividers, statements and closing
- These keep the template look (not white). `deck.end_slide()` — "Together we transform aviation" closing slide — is mandatory and always last, as the cover (`title_slide`) is mandatory and always first.
- `deck.divider_slide("Appendix", "2")` and `deck.statement_slide("One sentence")` exist but are used **only when the user explicitly asks** for a divider or a single-statement slide. By default a deck has no dividers or sections between slides.

## Building blocks not wrapped by a helper
Compose them from `box`, `add_text`, `icon_tile`, `hline/vline`:
- **Timeline/Gantt strip**: a thin asphalt line, circular dots (`MSO_SHAPE.OVAL`) at each milestone, date in 9 pt above, label below; a highlighted milestone has an orange dot and an orange-outlined callout.
- **Tables**: python-pptx `add_table`; header row fill asphalt with white 10 pt bold, body 10 pt on `F2F7F8`, no cell borders, highlight cells with orange `FDD9C4`, number columns right-aligned.
- **Supplier/OpCo logo rows**: `logo()` at h 0.35–0.5 in a white rounded chip.

## If a requested structure doesn't suit these patterns
Say so in one line and propose the closest alternative, for example: "A 12-row table doesn't work in the card layout; I used a full-width table slide (Title and Content layout) with the same colours."
