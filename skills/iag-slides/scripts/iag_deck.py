"""iag_deck.py - helper library to build IAG Transform branded decks with python-pptx.

Usage (from a script placed anywhere):
    import sys; sys.path.insert(0, "<skill dir>/scripts")
    from iag_deck import *
    deck = Deck()                                   # opens assets/template.pptx (no slides, 58 layouts)
    s = deck.title_slide("Title", "Subtitle")
    s = deck.content_slide("Topic | Subtopic", "Orange key message")
    kpi_band(s, 0.39, 1.70, 12.55, 0.85, "Header", "sub", [("LABEL", "€1.0m", "28%"), ...])
    ...
    footnote(s, "Source: ...")
    deck.end_slide()
    deck.save("out.pptx")

All coordinates are in inches on a 13.333 x 7.5 in slide. Colours are hex strings without '#'.
Text markup accepted by every text helper:  **bold**   {{orange bold}}   \n = new paragraph.
"""
import copy
import os
import re

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.normpath(os.path.join(HERE, "..", "assets"))
TEMPLATE = os.path.join(ASSETS, "template.pptx")
LOGOS = os.path.join(ASSETS, "brand", "logos")
SUPPLIER_LOGOS = os.path.join(LOGOS, "suppliers")
PHOTOS = os.path.join(ASSETS, "brand", "photos")
ICONS = os.path.join(ASSETS, "fa-icons")

# ---- brand palette (theme values of the IAG Transform template) -------------------------
ASPHALT = "233038"      # primary dark
ASPHALT_1 = "415057"    # lighter dark (step chevrons, secondary cards)
ASPHALT_2 = "819095"    # mid grey (captions, secondary figures)
ASPHALT_3 = "C1D1D3"    # light grey (lines, borders)
ORANGE = "F15D22"       # Transform orange (accent / highlight)
ORANGE_2 = "F9A477"
ORANGE_3 = "FDD9C4"     # light orange (soft highlight cells)
LIGHT_BLUE = "E0F2F3"   # primary light
PANEL = "F2F7F8"        # very light panel background used in the model slides
WHITE = "FFFFFF"
BLACK = "1D1D1B"
CLASSIFICATION = "INTERNAL USE ONLY │ NOT FOR ONWARD DISTRIBUTION"


def rgb(hexstr):
    return RGBColor.from_string(hexstr.upper())


# ------------------------------------------------------------------------------------------
# Rich text
# ------------------------------------------------------------------------------------------
_TOKEN = re.compile(r"(\*\*.+?\*\*|\{\{.+?\}\})", re.S)


def _runs(text):
    """Split markup into (text, bold, orange) tuples."""
    out = []
    for part in _TOKEN.split(text):
        if not part:
            continue
        if part.startswith("**"):
            out.append((part[2:-2], True, False))
        elif part.startswith("{{"):
            out.append((part[2:-2], True, True))
        else:
            out.append((part, None, False))
    return out


def _fill_paragraph(p, text, size, color, bold, italic=False, accent=ORANGE):
    for t, b, orange in _runs(text):
        r = p.add_run()
        r.text = t
        r.font.size = Pt(size)
        r.font.bold = bold if b is None else b
        r.font.italic = italic
        r.font.color.rgb = rgb(accent if orange else color)


def add_text(slide, x, y, w, h, text, size=12, color=ASPHALT, bold=False, align="l",
             anchor="t", italic=False, space_after=0, line_spacing=None, margin=0.0,
             name=None, accent=ORANGE):
    """Text box. `text` may be a str (use \\n for paragraphs) or a list of str / (str, dict) items.
    A dict may override size, color, bold, align, space_after for that paragraph."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if name:
        tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    for side in ("left", "right", "top", "bottom"):
        setattr(tf, f"margin_{side}", Inches(margin))
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    items = text.split("\n") if isinstance(text, str) else list(text)
    for i, item in enumerate(items):
        opts = {}
        if isinstance(item, tuple):
            item, opts = item
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[opts.get("align", align)]
        p.space_after = Pt(opts.get("space_after", space_after))
        p.space_before = Pt(0)
        if line_spacing:
            p.line_spacing = line_spacing
        _fill_paragraph(p, item, opts.get("size", size), opts.get("color", color),
                        opts.get("bold", bold), opts.get("italic", italic), accent)
    return tb


# ------------------------------------------------------------------------------------------
# Shapes
# ------------------------------------------------------------------------------------------
def set_alpha(shape, transparency_pct):
    """Make a solid fill translucent (0 = opaque, 100 = invisible)."""
    clr = shape.fill._xPr.find(qn("a:solidFill"))[0]
    for old in clr.findall(qn("a:alpha")):
        clr.remove(old)
    a = etree.SubElement(clr, qn("a:alpha"))
    a.set("val", str(int((100 - transparency_pct) * 1000)))


def box(slide, x, y, w, h, fill=PANEL, radius=0.04, line=None, line_w=0.75, transparency=0,
        shape=None, name=None):
    """Rounded rectangle (radius = fraction of the short side; model slides use 0.03-0.18)."""
    kind = shape or (MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE)
    sp = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    if radius and kind == MSO_SHAPE.ROUNDED_RECTANGLE:
        sp.adjustments[0] = radius
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = rgb(fill)
        if transparency:
            set_alpha(sp, transparency)
    if line:
        sp.line.color.rgb = rgb(line)
        sp.line.width = Pt(line_w)
    else:
        sp.line.fill.background()
    sp.shadow.inherit = False
    if name:
        sp.name = name
    return sp


def vline(slide, x, y, h, color=ASPHALT_2, width=0.75):
    ln = slide.shapes.add_connector(1, Inches(x), Inches(y), Inches(x), Inches(y + h))
    ln.line.color.rgb = rgb(color)
    ln.line.width = Pt(width)
    return ln


def hline(slide, x, y, w, color=ORANGE, width=0.75):
    ln = slide.shapes.add_connector(1, Inches(x), Inches(y), Inches(x + w), Inches(y))
    ln.line.color.rgb = rgb(color)
    ln.line.width = Pt(width)
    return ln


def icon_path(name):
    p = os.path.join(ICONS, name if name.endswith(".png") else name + ".png")
    if not os.path.exists(p):
        raise FileNotFoundError(f"icon '{name}' not in assets/fa-icons (see references/assets.md)")
    return p


def icon_tile(slide, x, y, size=0.5, icon="bullseye", fill=ORANGE, radius=0.15):
    """Orange rounded tile with a white icon centred in it."""
    box(slide, x, y, size, size, fill=fill, radius=radius, name="Icon tile")
    pad = size * 0.22
    slide.shapes.add_picture(icon_path(icon), Inches(x + pad), Inches(y + pad),
                             Inches(size - 2 * pad), Inches(size - 2 * pad))


def logo(slide, name, x, y, h=None, w=None):
    """Place a logo from assets/brand/logos (see references/assets.md for names)."""
    p = os.path.join(LOGOS, name if name.endswith(".png") else name + ".png")
    kw = {}
    if h:
        kw["height"] = Inches(h)
    if w:
        kw["width"] = Inches(w)
    return slide.shapes.add_picture(p, Inches(x), Inches(y), **kw)


# ---- logo lookup by company name ----------------------------------------------------------
_ALIASES = {
    # suppliers / partners (assets/brand/logos/suppliers/<file>.png)
    "boeing": "boeing", "collins": "collins-aerospace", "collinsaerospace": "collins-aerospace",
    "rtx": "collins-aerospace", "honeywell": "honeywell", "lufthansatechnik": "lufthansa-technik",
    "lufthansa": "lufthansa-technik", "lht": "lufthansa-technik", "moog": "moog",
    "moogaircraftgroup": "moog", "ajw": "ajw", "ajwgroup": "ajw", "turkishtechnic": "turkish-technic",
    "turkish": "turkish-technic", "airfranceindustries": "air-france-industries-klm-em",
    "afiklm": "air-france-industries-klm-em", "afklm": "air-france-industries-klm-em",
    "afiklmem": "air-france-industries-klm-em", "afklmem": "air-france-industries-klm-em", "airfrancklm": "air-france-industries-klm-em",
    "airfranceklm": "air-france-industries-klm-em", "klm": "air-france-industries-klm-em",
    "airfranceklmem": "air-france-industries-klm-em", "klmem": "air-france-industries-klm-em",
    # IAG group (assets/brand/logos/<file>_<variant>.png)
    "ba": "british-airways", "britishairways": "british-airways", "iberia": "iberia",
    "ib": "iberia", "iberiaexpress": "iberia-express", "aerlingus": "aer-lingus", "el": "aer-lingus",
    "vueling": "vueling", "vy": "vueling", "level": "level", "iagloyalty": "iag-loyalty",
    "loyalty": "iag-loyalty", "iagcargo": "iag-cargo", "cargo": "iag-cargo", "iag": "iag",
    "iagtransform": "iag-transform", "transform": "iag-transform",
}


def _norm(name):
    return re.sub(r"[^a-z0-9]", "", name.lower())


def find_logo(name, variant="color"):
    """Return the path of the logo for a company name, or None.
    Accepts names, abbreviations and file stems ("Boeing", "LHT", "Collins Aerospace", "BA", "Iberia Express").
    variant: "color" for light backgrounds, "white" for dark ones (group logos only; suppliers have one version)."""
    key = _ALIASES.get(_norm(name))
    candidates = []
    if key:
        candidates += [os.path.join(SUPPLIER_LOGOS, key + ".png"),
                       os.path.join(LOGOS, f"{key}_{variant}.png"),
                       os.path.join(LOGOS, f"{key}.png")]
    # any file dropped into suppliers/ is found by its own name, no alias needed
    n = _norm(name)
    if os.path.isdir(SUPPLIER_LOGOS):
        for fn in os.listdir(SUPPLIER_LOGOS):
            if fn.lower().endswith(".png") and _norm(fn[:-4]) in (n, n.replace("group", "")):
                candidates.append(os.path.join(SUPPLIER_LOGOS, fn))
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def list_logos():
    """Names of the bundled supplier logos and group logos."""
    sup = sorted(f[:-4] for f in os.listdir(SUPPLIER_LOGOS) if f.endswith(".png")) if os.path.isdir(SUPPLIER_LOGOS) else []
    grp = sorted({f.rsplit("_", 1)[0] for f in os.listdir(LOGOS) if f.endswith(".png")})
    return {"suppliers": sup, "group": grp}


def _dimmed_stream(path, opacity=0.35):
    import io
    from PIL import Image
    im = Image.open(path).convert("RGBA")
    im.putalpha(im.getchannel("A").point(lambda v: int(v * opacity)))
    buf = io.BytesIO()
    im.save(buf, "PNG")
    buf.seek(0)
    return buf


def logo_chip(slide, name, x, y, w=1.7, h=0.6, framed=False, dimmed=False, note=None,
              variant="color", pad=0.05, fill=WHITE, line=ASPHALT_3):
    """Company logo fitted (aspect ratio kept) inside the cell x,y,w,h.
    Default = loose logo, as in the reference slides (no frame). framed=True adds a white rounded chip with a thin border.
    dimmed=True fades it (e.g. a supplier that is not participating) and `note` prints a small grey italic caption below
    ("Not bidding"). If no logo is bundled for `name`, the name is written as text and a warning is printed: tell the user and
    suggest dropping the PNG into assets/brand/logos/suppliers/. Returns True if a logo was used."""
    from PIL import Image
    if framed:
        box(slide, x, y, w, h, fill=fill, radius=0.12, line=line, name=f"Logo chip {name}")
    path = find_logo(name, variant)
    if not path:
        print(f"WARNING: no logo bundled for '{name}'; used text. Add assets/brand/logos/suppliers/<name>.png")
        add_text(slide, x + pad, y, w - 2 * pad, h, name, size=11, bold=True, align="c", anchor="m",
                 color=ASPHALT_2 if dimmed else ASPHALT)
        return False
    iw, ih = Image.open(path).size
    note_h = 0.16 if note else 0.0
    aw, ah = w - 2 * pad, h - 2 * pad - note_h
    scale = min(aw / iw, ah / ih)
    pw, ph = iw * scale, ih * scale
    src = _dimmed_stream(path) if dimmed else path
    slide.shapes.add_picture(src, Inches(x + (w - pw) / 2), Inches(y + pad + (ah - ph) / 2), Inches(pw), Inches(ph))
    if note:
        add_text(slide, x, y + h - note_h - 0.02, w, note_h, note, size=8, italic=True, color=ASPHALT_2, align="c")
    return True


def logo_row(slide, names, x, y, w, h=0.6, gap=0.15, **kw):
    """Row of logos in equal cells. Returns list of booleans (logo found). Accepts the logo_chip options."""
    n = len(names)
    cw = (w - gap * (n - 1)) / n
    return [logo_chip(slide, nm, x + i * (cw + gap), y, cw, h, **kw) for i, nm in enumerate(names)]


def logo_group(slide, label, names, x, y, w, cols=3, cell_h=0.6, gap=0.12, dimmed=(), note="Not bidding"):
    """Labelled group of logos, as in the supplier panel of the reference slides: a small italic caps label
    (e.g. INTEGRATOR, OEM, PARTNERS) over a grid of loose logos. `dimmed` = names of companies to fade with a caption
    (suppliers not participating). Returns the y just below the group."""
    add_text(slide, x, y, w, 0.2, label.upper(), size=9, bold=True, italic=True, color=ASPHALT)
    gy = y + 0.28
    cw = (w - gap * (cols - 1)) / cols
    for i, nm in enumerate(names):
        r, c = divmod(i, cols)
        d = nm in dimmed
        logo_chip(slide, nm, x + c * (cw + gap), gy + r * (cell_h + gap), cw, cell_h, dimmed=d, note=note if d else None)
    rows = (len(names) + cols - 1) // cols
    return gy + rows * (cell_h + gap)


def photo_background(slide, photo, overlay=ASPHALT, overlay_transparency=22):
    """OPT-IN ONLY. Full-bleed photo + dark overlay. Information slides must stay white; use this only if the user explicitly asks."""
    p = photo if os.path.isabs(photo) else os.path.join(PHOTOS, photo)
    pic = slide.shapes.add_picture(p, 0, 0, Inches(13.333), Inches(7.5))
    ov = box(slide, 0, 0, 13.333, 7.5, fill=overlay, radius=0, transparency=overlay_transparency,
             name="Photo overlay")
    tree = slide.shapes._spTree
    for el in (ov._element, pic._element):          # photo first, overlay above it, both behind text
        tree.remove(el)
    tree.insert(2, pic._element)
    tree.insert(3, ov._element)
    return pic


# ------------------------------------------------------------------------------------------
# Footer pieces
# ------------------------------------------------------------------------------------------
def _slide_number_field(slide, color=ASPHALT_2, x=12.2, y=6.95, w=0.74, h=0.3):
    tb = add_text(slide, x, y, w, h, "", size=9, color=color, align="r", name="Slide number")
    p = tb.text_frame.paragraphs[0]
    for r in list(p.runs):
        p._p.remove(r._r)
    fld = etree.SubElement(p._p, qn("a:fld"))
    fld.set("id", "{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}")
    fld.set("type", "slidenum")
    rpr = etree.SubElement(fld, qn("a:rPr"))
    rpr.set("lang", "en-GB")
    rpr.set("sz", "900")
    sf = etree.SubElement(rpr, qn("a:solidFill"))
    c = etree.SubElement(sf, qn("a:srgbClr"))
    c.set("val", color)
    t = etree.SubElement(fld, qn("a:t"))
    t.text = "‹#›"
    return tb


def footnote(slide, text, dark=False):
    """Source / notes line (8 pt grey) + slide number. Content slides only."""
    col = "B5C2C5" if dark else ASPHALT_2
    add_text(slide, 1.5, 6.86, 9.9, 0.4, text, size=8, color=col, name="Footnote", anchor="t")
    _slide_number_field(slide, color="FFFFFF" if dark else ASPHALT_2)


# ------------------------------------------------------------------------------------------
# Components taken from the reference slides
# ------------------------------------------------------------------------------------------
def kpi_band(slide, x, y, w, h, title, subtitle, kpis, fill=ASPHALT):
    """Dark band: left title/subtitle, then N KPI cells separated by thin lines.
    kpis = [(LABEL, value, small_suffix_or_None), ...]  (label in orange caps)."""
    box(slide, x, y, w, h, fill=fill, radius=0.12, name="KPI band")
    left_w = 2.2
    add_text(slide, x + 0.2, y, left_w, h, [(title, {"size": 14, "bold": True, "color": WHITE}),
                                            (subtitle, {"size": 10, "color": "DCE5E7"})],
             anchor="m", name="KPI title")
    cell_w = (w - left_w - 0.25) / len(kpis)
    for i, (label, value, suffix) in enumerate(kpis):
        cx = x + left_w + 0.25 + i * cell_w
        vline(slide, cx - 0.1, y + h * 0.22, h * 0.56)
        paras = [(label.upper(), {"size": 9, "bold": True, "color": ORANGE}),
                 (value + (f"  {{{{{suffix}}}}}" if False else ""), {"size": 20, "bold": True, "color": WHITE})]
        tb = add_text(slide, cx + 0.02, y + 0.1, cell_w - 0.2, h - 0.2, paras, anchor="m", name="KPI")
        if suffix:
            r = tb.text_frame.paragraphs[1].add_run()
            r.text = "  " + suffix
            r.font.size = Pt(11)
            r.font.bold = False
            r.font.color.rgb = rgb("DCE5E7")


def panel(slide, x, y, w, h, header=None, fill=PANEL, radius=0.03, header_h=0.40):
    """Light panel with optional dark header strip (13 pt white bold). Returns inner top-y."""
    box(slide, x, y, w, h, fill=fill, radius=radius, name="Panel")
    if header:
        box(slide, x, y, w, header_h, fill=ASPHALT, radius=0.12, name="Panel header")
        add_text(slide, x + 0.12, y, w - 0.24, header_h, header, size=13, bold=True, color=WHITE,
                 anchor="m", name="Panel header text")
        return y + header_h + 0.05
    return y


def icon_text_rows(slide, x, y, w, rows, row_h=1.08, text_size=11, title_size=12, color=ASPHALT,
                   tile=0.5):
    """Vertical list: orange icon tile + bold title + body.  rows = [(icon, title, body), ...]"""
    for i, (icon, title, body) in enumerate(rows):
        ry = y + i * row_h
        icon_tile(slide, x, ry + 0.05, tile, icon)
        add_text(slide, x + tile + 0.15, ry, w - tile - 0.15, row_h - 0.05,
                 [(title, {"size": title_size, "bold": True, "color": color}),
                  (body, {"size": text_size, "color": color})], name="Row text")


def stat_card(slide, x, y, w, h, icon, title, big_text, dark=True, title_size=12.5, text_size=11,
              big_size=18):
    """Card with icon tile + title and a big-number sentence. `big_text` markup: {{€3.0m}} text."""
    fg = WHITE if dark else ASPHALT
    box(slide, x, y, w, h, fill=ASPHALT if dark else WHITE, radius=0.06,
        line=None if dark else ASPHALT_3, name="Stat card")
    icon_tile(slide, x + 0.18, y + 0.18, 0.46, icon)
    add_text(slide, x + 0.78, y + 0.18, w - 0.95, 0.46, title, size=title_size, bold=True, color=fg,
             anchor="m")
    add_text(slide, x + 0.18, y + 0.78, w - 0.36, h - 0.9, big_text, size=text_size, color=fg,
             name="Stat text")


def chip(slide, x, y, w, h, big, text, big_color=ORANGE, big_w=0.95):
    """White rounded chip: big % on the left, caption on the right."""
    box(slide, x, y, w, h, fill=WHITE, radius=0.12, line=ASPHALT_3, name="Chip")
    add_text(slide, x + 0.12, y, big_w, h, big, size=20, bold=True, color=big_color, anchor="m")
    add_text(slide, x + 0.12 + big_w, y, w - big_w - 0.2, h, text, size=9.5, color=ASPHALT, anchor="m")


def steps_band(slide, x, y, w, h, title, steps, highlight_last=True):
    """Dark band with chevron steps. steps = [(DATE, title, body), ...] (4-5 steps)."""
    box(slide, x, y, w, h, fill=ASPHALT, radius=0.08, name="Steps band")
    add_text(slide, x + 0.2, y + 0.1, 3, 0.3, title.upper(), size=10, bold=True, color=ORANGE)
    n = len(steps)
    gap = 0.12
    sw = (w - 0.4 - gap * (n - 1)) / n
    for i, (date, head, body) in enumerate(steps):
        sx = x + 0.2 + i * (sw + gap)
        last = highlight_last and i == n - 1
        box(slide, sx, y + h - 0.88 - 0.12, sw, 0.88, fill=ORANGE if last else ASPHALT_1,
            radius=0, shape=MSO_SHAPE.PENTAGON, name="Step")
        add_text(slide, sx + 0.15, y + h - 1.0, sw - 0.5, 0.88,
                 [(date.upper(), {"size": 8.5, "bold": True, "color": WHITE}),
                  (head, {"size": 11, "bold": True, "color": WHITE}),
                  (body, {"size": 8.5, "color": WHITE})], anchor="m", name="Step text")


def numbered_card(slide, x, y, w, h, number, title, body, icon=None, dark=True):
    """Q&A / lever card: icon + title, 01-style number top right, body below."""
    fg = WHITE if dark else ASPHALT
    box(slide, x, y, w, h, fill=ASPHALT if dark else WHITE, radius=0.06,
        line=None if dark else ASPHALT_3, name="Numbered card")
    tx = x + 0.18
    if icon:
        icon_tile(slide, x + 0.15, y + 0.15, 0.42, icon)
        tx = x + 0.72
    add_text(slide, tx, y + 0.12, w - (tx - x) - 0.75, 0.5, title, size=11.5, bold=True, color=fg, anchor="m")
    add_text(slide, x + w - 0.7, y + 0.12, 0.55, 0.5, number, size=16, bold=True, color=ORANGE,
             align="r", anchor="m")
    add_text(slide, x + 0.18, y + 0.75, w - 0.36, h - 0.85, body, size=10.5, color=fg)


def section_label(slide, x, y, w, text, color=ASPHALT):
    """Small caps label followed by a thin orange rule (used on dark photo slides)."""
    tb = add_text(slide, x, y - 0.1, 3.6, 0.22, text.upper(), size=10, bold=True, color=color, anchor="m")
    hline(slide, x + 0.1 + len(text) * 0.085, y, w - len(text) * 0.085 - 0.1, ORANGE, 0.75)
    return tb


def comparison_table(slide, x, y, w, row_labels, col_heads, cells, label_w=2.4, row_h=0.62,
                     head_h=0.42, icons=None, col_fills=(ASPHALT_1, ORANGE), text_size=11):
    """Row-label column (dark tiles, optional icon) + N comparison columns with coloured headers.
    cells[r][c] = text (markup ok). Last column is highlighted (light-orange cells)."""
    ncols = len(col_heads)
    cw = (w - label_w - 0.1 * ncols) / ncols
    for c, head in enumerate(col_heads):
        cx = x + label_w + 0.1 + c * (cw + 0.1)
        fill = col_fills[min(c, len(col_fills) - 1)] if len(col_fills) < ncols else col_fills[c]
        box(slide, cx, y, cw, head_h, fill=fill, radius=0.12)
        add_text(slide, cx + 0.15, y, cw - 0.3, head_h, head, size=13, bold=True, color=WHITE, anchor="m")
    for r, label in enumerate(row_labels):
        ry = y + head_h + 0.1 + r * (row_h + 0.08)
        box(slide, x, ry, label_w, row_h, fill=ASPHALT, radius=0.1)
        lx = x + 0.12
        if icons:
            icon_tile(slide, x + 0.1, ry + (row_h - 0.4) / 2, 0.4, icons[r])
            lx = x + 0.6
        add_text(slide, lx, ry, label_w - (lx - x) - 0.08, row_h, label, size=11, bold=True,
                 color=WHITE, anchor="m")
        for c in range(ncols):
            cx = x + label_w + 0.1 + c * (cw + 0.1)
            last = c == ncols - 1
            box(slide, cx, ry, cw, row_h, fill=ORANGE_3 if last else PANEL, radius=0.1)
            add_text(slide, cx + 0.15, ry, cw - 0.3, row_h, cells[r][c], size=text_size,
                     color=ASPHALT, anchor="m", bold=last)


def simple_table(slide, x, y, w, data, col_widths=None, row_h=0.32, size=10, highlight=(),
                 first_col_bold=True, align_numbers=True):
    """Compact brand table. data[0] is the header row. highlight = {(row, col), ...} (0-based incl. header)
    shown in orange tint. Header asphalt/white, body light panel, no borders."""
    from pptx.oxml.ns import qn as _qn
    rows, cols = len(data), len(data[0])
    gf = slide.shapes.add_table(rows, cols, Inches(x), Inches(y), Inches(w), Inches(row_h * rows))
    tbl = gf.table
    tblPr = tbl._tbl.tblPr
    for attr in ("firstRow", "bandRow"):
        tblPr.set(attr, "0")
    style = tblPr.find(_qn("a:tableStyleId"))
    if style is not None:
        style.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"      # 'No Style, No Grid'
    if col_widths:
        tot = sum(col_widths)
        for i, cw in enumerate(col_widths):
            tbl.columns[i].width = Inches(w * cw / tot)
    for r in range(rows):
        tbl.rows[r].height = Inches(row_h)
        for c in range(cols):
            cell = tbl.cell(r, c)
            cell.margin_left = cell.margin_right = Inches(0.08)
            cell.margin_top = cell.margin_bottom = Inches(0.02)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if r == 0:
                cell.fill.fore_color.rgb = rgb(ASPHALT)
            elif (r, c) in highlight:
                cell.fill.fore_color.rgb = rgb(ORANGE_3)
            else:
                cell.fill.fore_color.rgb = rgb(PANEL)
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            if c > 0 and align_numbers:
                p.alignment = PP_ALIGN.CENTER
            col = WHITE if r == 0 else ASPHALT
            _fill_paragraph(p, str(data[r][c]), size, col, r == 0 or (c == 0 and first_col_bold) or (r, c) in highlight)
    return gf


# ------------------------------------------------------------------------------------------
# Native charts (editable in PowerPoint) styled like the reference slides
# ------------------------------------------------------------------------------------------
def _style_chart(chart, legend=False, size=11):
    chart.font.size = Pt(size)
    chart.font.name = "Arial"
    chart.font.color.rgb = rgb(ASPHALT)
    chart.has_title = False
    chart.has_legend = legend
    if legend:
        chart.legend.position = XL_LEGEND_POSITION.TOP
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(size)


def stacked_pct_chart(slide, x, y, w, h, categories, series, colors=(ORANGE, ASPHALT),
                      number_format='0%', label_colors=(WHITE, WHITE)):
    """100%-stacked columns (e.g. settlement mix). series = {name: [values 0-1]}."""
    cd = CategoryChartData()
    cd.categories = categories
    for n, v in series.items():
        cd.add_series(n, v)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED_100, Inches(x), Inches(y), Inches(w),
                                Inches(h), cd)
    ch = gf.chart
    _style_chart(ch, legend=True)
    ch.plots[0].gap_width = 80
    ch.plots[0].overlap = 100
    ch.value_axis.visible = False
    ch.value_axis.has_major_gridlines = False
    ch.category_axis.format.line.color.rgb = rgb(ASPHALT_2)
    ch.category_axis.tick_labels.font.size = Pt(11)
    for i, s in enumerate(ch.series):
        s.format.fill.solid()
        s.format.fill.fore_color.rgb = rgb(colors[i % len(colors)])
        s.data_labels.show_value = True
        s.data_labels.number_format = number_format
        s.data_labels.number_format_is_linked = False
        s.data_labels.position = XL_LABEL_POSITION.CENTER
        s.data_labels.font.bold = True
        s.data_labels.font.size = Pt(11)
        s.data_labels.font.color.rgb = rgb(label_colors[i % len(label_colors)])
    return gf


def bar_chart(slide, x, y, w, h, categories, values, highlight=(0,), number_format='"€"0.0"m"',
              horizontal=True, base_color=ASPHALT, highlight_color=ORANGE, size=10.5, name="Series"):
    """Single-series bar/column chart with highlighted bar(s) in orange and value labels."""
    cd = CategoryChartData()
    cd.categories = categories
    cd.add_series(name, values)
    kind = XL_CHART_TYPE.BAR_CLUSTERED if horizontal else XL_CHART_TYPE.COLUMN_CLUSTERED
    gf = slide.shapes.add_chart(kind, Inches(x), Inches(y), Inches(w), Inches(h), cd)
    ch = gf.chart
    _style_chart(ch, size=size)
    plot = ch.plots[0]
    plot.gap_width = 45
    plot.vary_by_categories = False
    ch.value_axis.visible = False
    ch.value_axis.has_major_gridlines = False
    ch.category_axis.format.line.color.rgb = rgb(ASPHALT_2)
    ch.category_axis.tick_labels.font.size = Pt(size)
    if horizontal:
        ch.category_axis.reverse_order = True
    s = ch.series[0]
    s.format.fill.solid()
    s.format.fill.fore_color.rgb = rgb(base_color)
    for i in range(len(categories)):
        pt = s.points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = rgb(highlight_color if i in highlight else base_color)
    s.data_labels.show_value = True
    s.data_labels.number_format = number_format
    s.data_labels.number_format_is_linked = False
    s.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    s.data_labels.font.bold = True
    s.data_labels.font.size = Pt(size)
    return gf


def line_chart(slide, x, y, w, h, categories, series, colors=(ORANGE, ASPHALT, ASPHALT_2),
               number_format="General", legend=True):
    cd = CategoryChartData()
    cd.categories = categories
    for n, v in series.items():
        cd.add_series(n, v)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(x), Inches(y), Inches(w),
                                Inches(h), cd)
    ch = gf.chart
    _style_chart(ch, legend=legend and len(series) > 1)
    ch.value_axis.has_major_gridlines = True
    ch.value_axis.major_gridlines.format.line.color.rgb = rgb(ASPHALT_3)
    ch.value_axis.format.line.fill.background()
    ch.value_axis.tick_labels.number_format = number_format
    ch.value_axis.tick_labels.number_format_is_linked = False
    ch.category_axis.format.line.color.rgb = rgb(ASPHALT_2)
    for i, s in enumerate(ch.series):
        c = rgb(colors[i % len(colors)])
        s.format.line.color.rgb = c
        s.format.line.width = Pt(2.5)
        s.smooth = False
        s.marker.format.fill.solid()
        s.marker.format.fill.fore_color.rgb = c
        s.marker.format.line.color.rgb = c
    return gf


# ------------------------------------------------------------------------------------------
# Deck
# ------------------------------------------------------------------------------------------
class Deck:
    """Wrapper over the IAG Transform template (python-pptx Presentation)."""

    def __init__(self, template=TEMPLATE):
        self.prs = Presentation(template)
        self._layouts = {l.name: l for l in self.prs.slide_layouts}
        self._has_cover = False
        self._has_end = False
        self._strip_sections()

    def _strip_sections(self):
        """The template carries PowerPoint sections of its own ('Guides and Guidelines...',
        'About the template'...) that point to deleted slides. Remove them so the deck has no
        section headers between slides."""
        pres = self.prs.part._element
        for ext in list(pres.iter(qn("p:ext"))):
            if ext.find(".//{http://schemas.microsoft.com/office/powerpoint/2010/main}sectionLst") is not None:
                parent = ext.getparent()
                parent.remove(ext)
                if len(parent) == 0:
                    parent.getparent().remove(parent)

    # -- layouts -------------------------------------------------------------------------
    def layout(self, name):
        if name not in self._layouts:
            raise KeyError(f"layout '{name}' not found. Available: {sorted(self._layouts)}")
        return self._layouts[name]

    def _new(self, layout_name):
        return self.prs.slides.add_slide(self.layout(layout_name))

    def _clean_placeholders(self, slide, keep=("title", "subtitle", "body")):
        """Remove empty placeholders so no 'Click to add text' prompts remain."""
        for ph in list(slide.placeholders):
            if ph.has_text_frame and not ph.text_frame.text.strip():
                ph._element.getparent().remove(ph._element)

    def _footer_placeholders(self, slide):
        """Clone footer + slide-number placeholders from the layout (title / end / divider slides)."""
        for ph in slide.slide_layout.placeholders:
            t = ph.placeholder_format.type
            if "FOOTER" in str(t) or "SLIDE_NUMBER" in str(t):
                el = copy.deepcopy(ph._element)
                slide.shapes._spTree.append(el)
                if "FOOTER" in str(t):
                    for tx in el.iter(qn("a:t")):
                        tx.text = CLASSIFICATION
                    if not list(el.iter(qn("a:t"))):
                        tb = el.find(qn("p:txBody"))
                        p = tb.find(qn("a:p"))
                        r = etree.SubElement(p, qn("a:r"))
                        etree.SubElement(r, qn("a:rPr")).set("lang", "en-GB")
                        etree.SubElement(r, qn("a:t")).text = CLASSIFICATION

    # -- slide builders ------------------------------------------------------------------
    def title_slide(self, title, subtitle="", layout="Title Slide [Image 3]"):
        s = self._new(layout)
        for ph in s.placeholders:
            idx = ph.placeholder_format.idx
            if idx == 0:
                ph.text_frame.text = title
                ph.text_frame.vertical_anchor = MSO_ANCHOR.BOTTOM   # title sits on the subtitle, as in the models
            elif idx == 1:
                ph.text_frame.text = subtitle
        self._footer_placeholders(s)
        self._clean_placeholders(s)
        self._has_cover = True
        return s

    def content_slide(self, title, subtitle="", layout="Title and Content (Subtitle)"):
        """Title (32 pt bold) + orange subtitle line. Body area is free for components."""
        s = self._new(layout)
        for ph in list(s.placeholders):
            idx = ph.placeholder_format.idx
            if idx == 0:
                ph.text_frame.text = title
            elif ph.placeholder_format.type is not None and "BODY" in str(ph.placeholder_format.type) and subtitle:
                ph.text_frame.text = subtitle
                ph.left, ph.top, ph.width, ph.height = Inches(0.39), Inches(0.88), Inches(12.55), Inches(0.7)
            else:
                ph._element.getparent().remove(ph._element)
        return s

    def divider_slide(self, title, number="", layout="Divider Slide [Colour]"):
        s = self._new(layout)
        for ph in s.placeholders:
            idx = ph.placeholder_format.idx
            if idx == 0:
                ph.text_frame.text = title
            elif idx == 13:
                ph.text_frame.text = number
        self._footer_placeholders(s)
        self._clean_placeholders(s)
        return s

    def statement_slide(self, text, layout="Statement [Dark]"):
        s = self._new(layout)
        for ph in s.placeholders:
            if ph.placeholder_format.idx == 0:
                ph.text_frame.text = text
        self._footer_placeholders(s)
        self._clean_placeholders(s)
        return s

    def end_slide(self, layout="End slide"):
        """Closing slide (photo, 'Together we transform aviation'). Standard last slide of a deck."""
        s = self._new(layout)
        self._footer_placeholders(s)
        self._clean_placeholders(s)
        self._has_end = True
        return s

    def save(self, path):
        """Every deck must have a cover (first) and a closing slide (last)."""
        if not self._has_cover:
            raise ValueError("Deck has no cover: call deck.title_slide(...) first, before any other slide.")
        if not self._has_end:
            self.end_slide()   # closing slide has no content to fill in: add it automatically
        self.prs.save(path)
        return path
