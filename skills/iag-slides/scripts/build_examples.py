"""Builds assets/examples/patterns-example.pptx: one slide per reference pattern (P1-P11) with INVENTED data.
Run:  python build_examples.py [out.pptx]
Read it for call signatures; copy a block and replace the content. None of the figures or names are real."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from iag_deck import *
from pptx.enum.shapes import MSO_SHAPE

d = Deck()

# P1 ---- cover
d.title_slide("Project title", "Subtitle | audience and context")

# P2 ---- KPI band + chart + insights
s = d.content_slide("Topic | Subtopic", "Key message: the conclusion of the slide in one sentence")
kpi_band(s, 0.39, 1.70, 12.55, 0.85, "Overview", "Jan – Dec 20XX",
         [("Total", "€10.0m", None), ("Category 1", "€4.0m", "40%"), ("Category 2", "€6.0m", "60%"), ("Share of Unit A", "70%", None)])
panel(s, 0.39, 2.75, 7.75, 3.95, "Mix by unit, Jan – Dec 20XX (% of total)")
stacked_pct_chart(s, 0.54, 3.25, 7.45, 2.9, ["Unit A", "Unit B", "Unit C"],
                  {"Category 1": [0.30, 0.60, 0.80], "Category 2": [0.70, 0.40, 0.20]})
add_text(s, 0.59, 6.2, 7.35, 0.4, "Takeaway in one bold line under the chart", size=10, bold=True)
panel(s, 8.35, 2.75, 4.59, 3.95, "Why it matters")
icon_text_rows(s, 8.55, 3.35, 4.2, [
    ("scale-balanced", "First insight title", "One or two lines explaining the insight."),
    ("chart-line", "Second insight title", "One or two lines explaining the insight."),
    ("layer-group", "Third insight title", "One or two lines explaining the insight.")])
footnote(s, "Source: name of the report or system, period. Notes on assumptions and conversions.")

# P3 ---- comparison A vs B
s = d.content_slide("Comparison | Criteria", "Option B is the better fit on most criteria")
comparison_table(s, 0.39, 1.75, 12.55, ["Criterion one", "Criterion two", "Criterion three", "Criterion four"],
                 ["Option A", "Option B (recommended)"],
                 [["Short statement about option A", "**Key phrase** about option B"],
                  ["Short statement about option A", "Short statement about option B"],
                  ["Short statement about option A", "Short statement about option B"],
                  ["Short statement about option A", "Short statement about option B"]],
                 icons=["shuffle", "money-bill-transfer", "magnifying-glass-chart", "clock"], row_h=0.78)
box(s, 0.39, 6.1, 12.55, 0.55, fill=LIGHT_BLUE, radius=0.12)
add_text(s, 0.55, 6.1, 12.2, 0.55, "**Principle:** the recommendation or rule that follows from the comparison.", size=11, anchor="m")
footnote(s, "Source: name of the document, date.")

# P4 ---- opportunity dashboard
s = d.content_slide("Opportunity | Next year", "The opportunity is concentrated in a few items")
panel(s, 0.39, 1.70, 5.0, 3.30, "Value by category, 20XX (€m)")
bar_chart(s, 0.49, 2.15, 4.8, 1.95, ["Category 1", "Category 2", "Category 3", "Category 4", "Category 5"],
          [7.0, 2.0, 1.5, 0.8, 0.4], number_format='"€"0.0"m"')
chip(s, 0.54, 4.15, 2.27, 0.72, "70%", "of the value sits in 10 items")
chip(s, 2.96, 4.15, 2.27, 0.72, "5%", "of the value sits in 100 items", big_color=ASPHALT_2)
stat_card(s, 5.59, 1.70, 3.65, 1.57, "plane-up", "First callout", "{{€2.0m}} short explanation of the callout")
stat_card(s, 5.59, 3.42, 3.65, 1.57, "handshake", "Second callout", "{{€1.0m}} short explanation of the callout")
box(s, 9.44, 1.70, 3.5, 3.3, fill=WHITE, radius=0.04, line=ASPHALT_3)
icon_tile(s, 9.64, 1.9, 0.5, "bullseye")
add_text(s, 10.29, 1.9, 2.5, 0.5, "NEXT YEAR AMBITION", size=11, bold=True, color=ORANGE, anchor="m")
add_text(s, 9.64, 2.55, 3.1, 2.4, [("Target name", {"size": 12, "bold": True}),
         ("€4m → €6m", {"size": 28, "bold": True, "color": ORANGE}), ("by December 20XX", {"size": 12}),
         ("How the target will be reached, one sentence.", {"size": 10})], space_after=4)
steps_band(s, 0.39, 5.2, 12.55, 1.45, "Next steps", [
    ("Month 1 – Month 2", "Step one", "Short detail"),
    ("Month 2", "Step two", "Short detail"),
    ("Month 3", "Step three", "Short detail"),
    ("Month 4", "Step four", "Short detail"),
    ("Year end", "Goal", "Short detail")])
footnote(s, "Source: name of the report, period.")

# P5 ---- numbered Q&A grid
s = d.content_slide("Appendix | Questions and answers", "Answers to the questions raised by the stakeholders")
qa = [("percent", "Question one?", "Answer: start with Yes / No / Open point and give the reason in two or three lines."),
      ("chart-line", "Question two?", "Answer: start with Yes / No / Open point and give the reason in two or three lines."),
      ("scale-balanced", "Question three?", "Answer: start with Yes / No / Open point and give the reason in two or three lines."),
      ("list-check", "Question four?", "Answer: start with Yes / No / Open point and give the reason in two or three lines."),
      ("users-gear", "Question five?", "Answer: start with Yes / No / Open point and give the reason in two or three lines."),
      ("file-signature", "Question six?", "Answer: start with Yes / No / Open point and give the reason in two or three lines.")]
for i, (ic, q, a) in enumerate(qa):
    numbered_card(s, 0.39 + (i % 3) * 4.2, 1.75 + (i // 3) * 2.5, 4.15, 2.35, f"0{i+1}", q, a, icon=ic)
footnote(s, "Source: name of the document, date.")

# P6 ---- scenario columns with mini tables
s = d.content_slide("Scenarios | Options", "Option C balances cost and risk and is proposed for joint assessment")
rows = [["", "Part 1", "Part 2", "Part 3"], ["Item A", "X", "X", "X"], ["Item B", "X", "Y", "X"], ["Item C", "X", "Y", "Y"]]
for i, (letter, name, desc, hl) in enumerate([("A", "Option A", "Description of option A in one sentence.", ()),
                                              ("B", "Option B", "Description of option B in one sentence.", {(1, 2), (2, 2), (3, 2)}),
                                              ("C", "Option C (proposed)", "Description of option C in one sentence.", {(2, 2), (3, 2), (3, 3)})]):
    x = 0.39 + i * 4.2
    box(s, x, 1.70, 4.15, 2.75, fill=WHITE, radius=0.04, line=ORANGE if letter == "C" else ASPHALT_3, line_w=1.5 if letter == "C" else 0.75)
    add_text(s, x + 0.15, 1.78, 0.5, 0.5, letter, size=28, bold=True, color=ORANGE, anchor="m")
    add_text(s, x + 0.7, 1.78, 3.3, 0.5, name, size=13, bold=True, anchor="m")
    add_text(s, x + 0.15, 2.3, 3.85, 0.45, desc, size=10)
    simple_table(s, x + 0.15, 2.8, 3.85, rows, col_widths=[1.4, 1, 1, 1], row_h=0.36, size=9.5, highlight=hl)
add_text(s, 0.39, 4.6, 6, 0.3, "LEVERS", size=10, bold=True, color=ORANGE)
for i, (n, t, b) in enumerate([("01", "Lever one", "Short explanation of the lever"), ("02", "Lever two", "Short explanation of the lever"),
                               ("03", "Lever three", "Short explanation of the lever"), ("04", "Lever four", "Short explanation of the lever")]):
    numbered_card(s, 0.39 + i * 3.17, 4.95, 3.07, 1.7, n, t, b)
footnote(s, "Source: name of the analysis. Illustrative data.")

# P7 ---- next steps timeline + decision gates + inputs
s = d.content_slide("Next steps | Plan", "Four phases to a joint recommendation")
phases = [("Month 1", "Phase one", ["Detail one", "Detail two"]), ("Month 1", "Phase two", ["Detail one", "Detail two"]),
          ("Months 2–3", "Phase three", ["Detail one", "Detail two"]), ("Goal", "Phase four", ["Detail one", "Detail two"])]
for i, (date, name, bullets) in enumerate(phases):
    x = 0.39 + i * 2.62
    last = i == len(phases) - 1
    box(s, x, 1.70, 2.55, 0.75, fill=ORANGE if last else ASPHALT_1, radius=0, shape=MSO_SHAPE.PENTAGON)
    add_text(s, x + 0.15, 1.70, 2.0, 0.75, [(date.upper(), {"size": 8.5, "bold": True, "color": WHITE}),
                                           (name, {"size": 11, "bold": True, "color": WHITE})], anchor="m")
    box(s, x, 2.55, 2.45, 1.45, fill=PANEL, radius=0.05)
    add_text(s, x + 0.12, 2.65, 2.2, 1.3, ["• " + b for b in bullets], size=10, space_after=3)
box(s, 10.95, 1.70, 1.99, 2.3, fill=WHITE, radius=0.05, line=ASPHALT_3)
add_text(s, 11.07, 1.8, 1.8, 2.1, [("DECISION GATES", {"size": 9, "bold": True, "color": ORANGE}),
                                   ("Committee A", {"size": 10, "bold": True}), ("Month 4", {"size": 18, "bold": True, "color": ORANGE}),
                                   ("Committee B", {"size": 10, "bold": True}), ("Month 5", {"size": 18, "bold": True, "color": ORANGE})], space_after=3)
box(s, 0.39, 4.3, 12.55, 2.35, fill=ASPHALT, radius=0.06)
add_text(s, 0.6, 4.4, 6, 0.3, "INPUTS TO AGREE", size=10, bold=True, color=ORANGE)
for i, (ic, t, b) in enumerate([("box", "Input one", "What needs to be agreed and with whom."), ("filter", "Input two", "What needs to be agreed and with whom."),
                                ("gauge-high", "Input three", "What needs to be agreed and with whom.")]):
    x = 0.6 + i * 4.1
    icon_tile(s, x, 4.85, 0.5, ic)
    add_text(s, x + 0.65, 4.8, 3.2, 1.7, [(t, {"size": 12, "bold": True, "color": WHITE}), (b, {"size": 10.5, "color": WHITE})], space_after=3)
footnote(s, "Illustrative timeline.")

# P8 ---- lever cards (white background)
s = d.content_slide("Drivers | Alternatives", "Cost drivers and alternatives to reduce the impact")
section_label(s, 0.39, 1.72, 12.55, "Drivers")
for i, (ic, t, big, txt) in enumerate([("arrow-trend-up", "Driver one", "+50%", "Explanation with **key words** in bold"),
                                       ("clock", "Driver two", "+100%", "Explanation with **key words** in bold"),
                                       ("filter", "Driver three", "2", "Explanation with **key words** in bold")]):
    x = 0.39 + i * 4.23
    box(s, x, 1.85, 4.09, 2.0, fill=ASPHALT, radius=0.06)
    icon_tile(s, x + 0.12, 1.95, 0.42, ic)
    add_text(s, x + 0.65, 1.95, 3.3, 0.42, t, size=14, bold=True, color=WHITE, anchor="m")
    box(s, x + 0.2, 2.5, 3.7, 0.65, fill=WHITE, radius=0.15)
    add_text(s, x + 0.2, 2.5, 3.7, 0.65, big, size=24, bold=True, align="c", anchor="m")
    add_text(s, x + 0.2, 3.25, 3.7, 0.55, txt, size=12, color=WHITE, align="c")
section_label(s, 0.39, 4.1, 12.55, "Alternatives")
for i, (ic, t, txt) in enumerate([("gear", "Alternative one", "Description with **key words** in bold"),
                                  ("diagram-project", "Alternative two", "Description with **key words** in bold"),
                                  ("magnifying-glass", "Alternative three", "Description with **key words** in bold"),
                                  ("handshake", "Alternative four", "Description with **key words** in bold")]):
    numbered_card(s, 0.39 + i * 3.17, 4.25, 3.07, 2.3, f"0{i+1}", t, txt, icon=ic)
footnote(s, "Illustrative data.")

# P9 ---- numbered grid with big orange tile
s = d.content_slide("Alternatives | Five levers", "Five levers we are exploring")
box(s, 0.39, 1.75, 3.6, 4.85, fill=ORANGE, radius=0.04)
add_text(s, 0.6, 2.3, 3.2, 1.2, "#5", size=54, bold=True, color=WHITE)
add_text(s, 0.6, 3.6, 3.2, 1.0, "levers to reduce total cost", size=18, bold=True, color=WHITE)
items = [("gear", "Lever one", "Short description of the lever."), ("diagram-project", "Lever two", "Short description of the lever."),
         ("magnifying-glass", "Lever three", "Short description of the lever."), ("handshake", "Lever four", "Short description of the lever."),
         ("lightbulb", "Lever five", "Short description of the lever.")]
for i, (ic, t, b) in enumerate(items):
    x = 4.2 + (i % 2) * 4.38
    y = 1.75 + (i // 2) * 1.65
    box(s, x, y, 4.3, 1.55, fill=ASPHALT, radius=0.06)
    add_text(s, x + 0.2, y + 0.12, 0.6, 0.4, f"0{i+1}", size=16, bold=True, color=ORANGE)
    add_text(s, x + 0.2, y + 0.55, 3.0, 0.95, [(t, {"size": 12, "bold": True, "color": WHITE}), (b, {"size": 10.5, "color": WHITE})], space_after=2)
    icon_tile(s, x + 3.6, y + 0.12, 0.5, ic)
footnote(s, "Illustrative data.")

# P10 ---- insight tiles
s = d.content_slide("Insights | Market", "Key headwinds and opportunities to drive value", )
tiles = [("globe", "Insight headline one", "First short bullet with a fact.\nSecond short bullet with a fact."),
         ("bolt", "Insight headline two", "First short bullet with a fact.\nSecond short bullet with a fact."),
         ("building", "Insight headline three", "First short bullet with a fact.\nSecond short bullet with a fact."),
         ("truck", "Insight headline four", "First short bullet with a fact.\nSecond short bullet with a fact."),
         ("leaf", "Insight headline five", "First short bullet with a fact.\nSecond short bullet with a fact."),
         ("users", "Insight headline six", "First short bullet with a fact.\nSecond short bullet with a fact.")]
for i, (ic, h, b) in enumerate(tiles):
    x = 0.39 + (i % 2) * 6.33
    y = 1.75 + (i // 2) * 1.65
    box(s, x, y, 6.22, 1.55, fill=PANEL, radius=0.05)
    add_text(s, x + 0.2, y + 0.12, 4.9, 0.4, h, size=13, bold=True)
    add_text(s, x + 0.2, y + 0.55, 4.9, 0.95, ["• " + t for t in b.split("\n")], size=10.5, space_after=3)
    icon_tile(s, x + 5.45, y + 0.12, 0.55, ic)
footnote(s, "Source: name of the report, date.")

# P11 ---- logo usage: loose logos, grouped by role, dimmed when not participating
s = d.content_slide("Suppliers | Logo usage", "Any supplier named by the user gets its own logo, loose and grouped by role")
panel(s, 0.39, 1.70, 5.6, 4.95, "Suppliers")
y = logo_group(s, "Integrator", ["Boeing", "Lufthansa Technik", "AJW", "Turkish Technic", "Air France Industries KLM E&M"],
               0.6, 2.25, 5.2, cols=3, cell_h=0.6, dimmed=("AJW", "Turkish Technic", "Air France Industries KLM E&M"))
hline(s, 0.6, y + 0.02, 5.2, ASPHALT_3)
logo_group(s, "OEM", ["Collins Aerospace", "Honeywell", "Moog"], 0.6, y + 0.15, 5.2, cols=3, cell_h=0.6, dimmed=("Honeywell", "Moog"))
panel(s, 6.2, 1.70, 6.74, 4.95, "Logo under each bar / column")
bar_chart(s, 6.35, 2.2, 6.4, 2.6, ["Option 1", "Option 2", "Option 3", "Option 4"], [480, 670, 890, 860], number_format='#,##0',
          horizontal=False, highlight=(2,))
for i, nm in enumerate(["Boeing", "Lufthansa Technik", "Collins Aerospace", "Honeywell"]):
    logo_chip(s, nm, [7.25, 8.78, 10.3, 11.83][i] - 0.7, 4.95, 1.4, 0.5)
add_text(s, 6.4, 5.65, 6.3, 0.9, ["**Rule:** show the logo of every supplier or group company mentioned, loose (no frame), at a consistent height.",
                                  "Not participating = faded with a small caption. Missing logo = name in text + tell the user."], size=10.5, space_after=4)
footnote(s, "Reference slide: logos used for layout purposes only; figures are invented.")

# P12 ---- divider + closing
d.divider_slide("Appendix", "2")
d.end_slide()

d.save(sys.argv[1] if len(sys.argv) > 1 else "patterns-example.pptx")
