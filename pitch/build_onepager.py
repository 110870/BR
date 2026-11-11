"""Builds the 'explore goes global' pitch one-pager (single 16:9 slide).

Run: python3 pitch/build_onepager.py  ->  pitch/explore_goes_global_onepager.pptx
"""
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

RED = RGBColor(0xE4, 0x00, 0x28)
NAVY = RGBColor(0x09, 0x28, 0x47)
BLUE_GRAY = RGBColor(0x6D, 0x7A, 0x93)
LIGHT = RGBColor(0xF3, 0xF4, 0xF7)
LINE = RGBColor(0xD9, 0xDC, 0xE3)
DARK = RGBColor(0x1A, 0x1A, 0x1A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = "Verdana"

OUT = Path(__file__).with_name("explore_goes_global_onepager.pptx")

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
s = prs.slides.add_slide(prs.slide_layouts[6])


def box(x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None):
    sh = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(0.75)
    if radius is not None:
        sh.adjustments[0] = radius
    sh.shadow.inherit = False
    return sh


def text(x, y, w, h, runs, anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.LEFT, shape=None):
    """runs: list of paragraphs; each paragraph = list of (text, size, bold, color)."""
    if shape is None:
        shape = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.08)
    tf.margin_top = tf.margin_bottom = Inches(0.04)
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(3)
        for t, size, bold, color in para:
            r = p.add_run()
            r.text = t
            r.font.name = FONT
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.color.rgb = color
    return shape


# ---------------------------------------------------------------- header
box(0, 0, 0.18, 7.5, fill=RED)
text(0.45, 0.28, 8.2, 0.7, [[("explore ", 30, True, DARK), ("goes global", 30, True, RED)]])
text(0.45, 0.98, 8.6, 0.5, [[
    ("Booking verkauft Bausteine. ", 15, True, DARK),
    ("Wir verkaufen die ganze Reise – weltweit.", 15, False, DARK),
]])

thesis = box(9.05, 0.3, 3.95, 1.25, fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
text(0, 0, 0, 0, [
    [("UNSERE THESE", 9, True, RGBColor(0xFC, 0x5F, 0x76))],
    [("Mehrtägige, verknüpfte Rundreisen – betreut und mit voller Veranstalter­haftung – "
      "für Kunden aus jedem Quellmarkt.", 10.5, False, WHITE)],
], anchor=MSO_ANCHOR.MIDDLE, shape=thesis)

# ---------------------------------------------------------------- left: comparison
LX, TOP = 0.45, 1.85
text(LX, TOP, 5.9, 0.35, [[("Der Unterschied", 13, True, DARK)]])

rows = [
    ("Produkt", "Einzelbaustein: Hotel, Mietwagen, Ticket", "Multiday-Itinerary über Länder & Regionen"),
    ("Verknüpfung", "Kunde puzzelt selbst", "Flug, Hotel, Transfer, Guide, Event – connected"),
    ("Service", "Self-Service bis zum Check-in", "Betreut vor & während der Reise"),
    ("Haftung", "Vermittler – Risiko liegt beim Kunden", "Veranstalter: ein Vertrag, ein Verantwortlicher"),
]
col_w = [1.35, 2.2, 2.35]
hy = TOP + 0.45
hdr = box(LX + col_w[0], hy, col_w[1], 0.42, fill=LIGHT)
text(0, 0, 0, 0, [[("Booking & Co.", 10.5, True, BLUE_GRAY)]], anchor=MSO_ANCHOR.MIDDLE,
     align=PP_ALIGN.CENTER, shape=hdr)
hdr2 = box(LX + col_w[0] + col_w[1], hy, col_w[2], 0.42, fill=RED)
text(0, 0, 0, 0, [[("explore", 10.5, True, WHITE)]], anchor=MSO_ANCHOR.MIDDLE,
     align=PP_ALIGN.CENTER, shape=hdr2)

ry, rh = hy + 0.42, 0.58
for label, them, us in rows:
    box(LX, ry + rh, sum(col_w), 0.0001, line=LINE)
    text(LX, ry, col_w[0], rh, [[(label, 9.5, True, DARK)]], anchor=MSO_ANCHOR.MIDDLE)
    text(LX + col_w[0], ry, col_w[1], rh, [[(them, 9, False, BLUE_GRAY)]], anchor=MSO_ANCHOR.MIDDLE)
    text(LX + col_w[0] + col_w[1], ry, col_w[2], rh, [[(us, 9, True, DARK)]], anchor=MSO_ANCHOR.MIDDLE)
    ry += rh

# ---------------------------------------------------------------- right: source markets
RX = 6.75
text(RX, TOP, 6.3, 0.35, [[("Ein Europa-Produkt. Jeder Quellmarkt.", 13, True, DARK)]])

cases = [
    ("USA", "Mallorca", "Der klassische Badeurlaub – für US-Kunden aus einer Hand gebucht und betreut."),
    ("USA", "Griechenland-Rundreise", "Athen, Peloponnes, Inseln: Fähren, Hotels und Guides verknüpft."),
    ("China", "Jede Europa-Rundreise", "Jede gewünschte Route darstellbar und verkaufbar – in der Sprache des Kunden."),
    ("Australien", "Europa-Tour", "Mehrländer-Rundreise inkl. Oktoberfest – Tickets und Zelt-Reservierung inklusive."),
]
cw, ch, gap = 3.05, 1.36, 0.17
for i, (src, dst, desc) in enumerate(cases):
    cx = RX + (i % 2) * (cw + gap)
    cy = TOP + 0.45 + (i // 2) * (ch + gap)
    box(cx, cy, cw, ch, fill=WHITE, line=LINE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
    box(cx, cy + 0.18, 0.07, ch - 0.36, fill=RED)
    text(cx + 0.15, cy + 0.1, cw - 0.25, 0.35, [[
        (src, 11, True, RED), ("  →  ", 11, True, BLUE_GRAY), (dst, 11, True, DARK),
    ]])
    text(cx + 0.15, cy + 0.5, cw - 0.25, ch - 0.6, [[(desc, 9, False, DARK)]])

# ---------------------------------------------------------------- bottom: moat pillars
BY = 5.3
text(LX, BY, 12.5, 0.35, [[("Warum das nur wir können", 13, True, DARK)]])
pillars = [
    ("1", "Itinerary-Kompetenz", "Mehrtägige Routen über Länder hinweg planen, bündeln und als ein Produkt verkaufen."),
    ("2", "Service on Trip", "Betreuung vor und während der Reise – wenn etwas schiefgeht, lösen wir es."),
    ("3", "Haftung als Veranstalter", "Wir stehen für die Gesamtleistung ein. Das gibt Kunden aus Übersee Sicherheit."),
]
pw = (12.45 - 2 * 0.2) / 3
for i, (n, head, body) in enumerate(pillars):
    px = LX + i * (pw + 0.2)
    box(px, BY + 0.42, pw, 1.0, fill=LIGHT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
    num = box(px + 0.15, BY + 0.62, 0.55, 0.55, fill=RED, shape=MSO_SHAPE.OVAL)
    text(0, 0, 0, 0, [[(n, 14, True, WHITE)]], anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER, shape=num)
    text(px + 0.82, BY + 0.5, pw - 0.95, 0.9, [
        [(head, 10.5, True, DARK)],
        [(body, 8.5, False, DARK)],
    ], anchor=MSO_ANCHOR.MIDDLE)

# ---------------------------------------------------------------- footer claim
text(LX, 6.85, 12.5, 0.4, [[
    ("Booking verkauft Nächte. ", 11, True, BLUE_GRAY),
    ("explore verkauft Reisen – von jedem Markt der Welt nach Europa.", 11, True, RED),
]], anchor=MSO_ANCHOR.MIDDLE)

prs.save(OUT)
print(f"saved {OUT}")
