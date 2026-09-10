"""
DigitalProducts_TOM_Proposal_v2.0.pptx — final client deck.
17 slides · client-facing · brand-clean · KB-grounded · zero internal labels.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pathlib import Path

OUT = Path("/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento")

# ── Brand palette ──
NAVY        = RGBColor(0x21, 0x3C, 0x69)
NAVY_DARK   = RGBColor(0x11, 0x23, 0x3E)
BLUE        = RGBColor(0x67, 0x85, 0xC1)
BLUE_LIGHT  = RGBColor(0xDC, 0xE6, 0xF5)
ACCENT      = RGBColor(0xE6, 0xB6, 0x00)
GREEN       = RGBColor(0x4C, 0x8C, 0x3F)
GREEN_LIGHT = RGBColor(0xE8, 0xF1, 0xE3)
TEAL        = RGBColor(0x2B, 0x6B, 0x77)
TEAL_LIGHT  = RGBColor(0xDC, 0xEC, 0xEF)
AMBER       = RGBColor(0xE8, 0x8B, 0x2F)
AMBER_LIGHT = RGBColor(0xFD, 0xEC, 0xD7)
INDIGO      = RGBColor(0x2C, 0x5A, 0x88)
INDIGO_LIGHT= RGBColor(0xDD, 0xE6, 0xF0)
RED         = RGBColor(0xB6, 0x3B, 0x1F)
RED_LIGHT   = RGBColor(0xF8, 0xDF, 0xD8)
PURPLE      = RGBColor(0x7A, 0x66, 0xB8)
PURPLE_LIGHT= RGBColor(0xE5, 0xDF, 0xF0)

WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
DARK        = RGBColor(0x2C, 0x3E, 0x50)
GREY        = RGBColor(0x6E, 0x78, 0x87)
LIGHT       = RGBColor(0xF5, 0xF5, 0xF7)
VACANT      = RGBColor(0xC8, 0x4A, 0x35)

FONT = "Arial"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW = prs.slide_width
SH = prs.slide_height


def blank():
    return prs.slides.add_slide(prs.slide_layouts[6])


def rect(slide, x, y, w, h, fill, line=None, line_w=0):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line:
        s.line.color.rgb = line
        s.line.width = Pt(line_w)
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    return s


def round_rect(slide, x, y, w, h, fill, line=None, line_w=0, radius=0.04):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    s.adjustments[0] = radius
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line:
        s.line.color.rgb = line
        s.line.width = Pt(line_w)
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    return s


def text_box(slide, x, y, w, h, text, size=14, bold=False, color=DARK,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = anchor
    tf.text = text
    p = tf.paragraphs[0]
    p.alignment = align
    for run in p.runs:
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color
        run.font.name = FONT
    return tb


def add_lines(slide, x, y, w, h, lines, size=12, color=DARK, bold=False,
              line_spacing=1.15):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.05)
    tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.04)
    tf.margin_bottom = Inches(0.04)
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.line_spacing = line_spacing
        for run in p.runs:
            run.font.size = Pt(size)
            run.font.bold = bold
            run.font.color.rgb = color
            run.font.name = FONT
    return tb


def header_bar(slide, title, color=NAVY_DARK):
    """Top action title bar — NO yellow accent (per client request)."""
    rect(slide, 0, 0, SW, Inches(0.85), color)
    text_box(slide, Inches(0.4), Inches(0.10), SW - Inches(0.8),
             Inches(0.65), title, size=20, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)


def footer_bar(slide):
    y = SH - Inches(0.32)
    text_box(slide, Inches(0.4), y, SW - Inches(0.8), Inches(0.3),
             "Avangrid · Digital Products · Target Operating Model",
             size=9, color=GREY, italic=True)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 1 — COVER
# ════════════════════════════════════════════════════════════════════════
s = blank()
rect(s, 0, 0, SW, SH, NAVY_DARK)
rect(s, 0, Inches(2.4), Inches(2.5), Inches(0.06), ACCENT)
text_box(s, Inches(0.7), Inches(2.05), Inches(8), Inches(0.4),
         "AVANGRID  ·  DIGITAL PRODUCTS",
         size=14, bold=True, color=ACCENT)
text_box(s, Inches(0.7), Inches(2.65), Inches(12), Inches(1.5),
         "Target Operating Model",
         size=56, bold=True, color=WHITE)
text_box(s, Inches(0.7), Inches(3.95), Inches(12), Inches(0.6),
         "Wave 1 — Demand Management Design & Pilot",
         size=24, color=BLUE)
text_box(s, Inches(0.7), Inches(5.0), Inches(12), Inches(1.2),
         "A product-centric operating model for IT Business Applications and "
         "Customer Services Applications under VP IT Business Solutions Delivery.",
         size=14, color=WHITE)
rect(s, 0, SH - Inches(0.6), SW, Inches(0.6), BLUE)
text_box(s, Inches(0.7), SH - Inches(0.55), Inches(10), Inches(0.5),
         "April 2026",
         size=12, bold=True, color=WHITE,
         anchor=MSO_ANCHOR.MIDDLE)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 2 — Executive Summary
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Executive Summary")

gt_y = Inches(1.1)
round_rect(s, Inches(0.5), gt_y, SW - Inches(1), Inches(1.1),
           BLUE_LIGHT, line=BLUE, line_w=2, radius=0.06)
text_box(s, Inches(0.7), gt_y + Inches(0.1), SW - Inches(1.4), Inches(0.35),
         "OUR PROPOSAL",
         size=11, bold=True, color=BLUE)
text_box(s, Inches(0.7), gt_y + Inches(0.45), SW - Inches(1.4), Inches(0.7),
         "Adopt a product-centric operating model with eight Value Stream Teams "
         "aligned to customer outcomes — converting today's fragmented intake into "
         "a scalable, governed framework that supports projected growth.",
         size=15, color=DARK, italic=True)

col_y = Inches(2.5)
col_h = Inches(3.7)
col_w = (SW - Inches(1.2)) / 3
gap = Inches(0.15)

cols = [
    ("WHY NOW", NAVY,
     ["Projected 30–50% growth in Digital Products demand",
      "Multiple uncoordinated intake channels create lost visibility",
      "No shared definition of value across business units",
      "Manual reporting disconnected from operational reality"]),
    ("WHAT CHANGES", BLUE,
     ["Single intake channel per Value Stream (Jira)",
      "Two-gate governance: weekly triage + monthly portfolio",
      "Class-of-Service framework with capacity caps",
      "Unified backlog system with one source of truth"]),
    ("HOW WE GET THERE", GREEN,
     ["Five-step Wave 1 over fourteen weeks",
      "Pilot one Value Stream end-to-end",
      "Refine and validate · scale to remaining teams",
      "Continuous improvement through feedback loops"]),
]
for i, (label, color, bullets) in enumerate(cols):
    x = Inches(0.5) + i * (col_w + gap)
    round_rect(s, x, col_y, col_w, col_h, WHITE, line=color, line_w=2, radius=0.04)
    rect(s, x, col_y, col_w, Inches(0.55), color)
    text_box(s, x, col_y + Inches(0.05), col_w, Inches(0.45),
             label, size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_lines(s, x + Inches(0.2), col_y + Inches(0.7),
              col_w - Inches(0.4), col_h - Inches(0.85),
              ["•  " + b for b in bullets], size=11.5, color=DARK,
              line_spacing=1.3)

footer_bar(s)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 3 — Current State Analysis (4 columns, matches user's image)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Current State Analysis · Digital Products spans two director branches plus Product Management")

# Sub-title
text_box(s, Inches(0.5), Inches(1.0), SW - Inches(1), Inches(0.4),
         "Digital Products spans two director branches (~50 FTE) across nine sub-teams "
         "and a separate Product Management group.",
         size=12, italic=True, color=BLUE, bold=True)

# Vitor at top
vp_y = Inches(1.6)
round_rect(s, Inches(0.5), vp_y, SW - Inches(1), Inches(0.7), NAVY_DARK, radius=0.05)
text_box(s, Inches(0.5), vp_y + Inches(0.05), SW - Inches(1), Inches(0.6),
         "Vitor Cruz — VP IT Business Solutions Delivery",
         size=15, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 4 columns
col_y = Inches(2.5)
col_h = Inches(4.3)
ncols = 4
gap = Inches(0.12)
col_w = (SW - Inches(1.0) - gap * (ncols - 1)) / ncols

# Column 1 — Carlos
x1 = Inches(0.5)
round_rect(s, x1, col_y, col_w, Inches(0.65), NAVY, radius=0.04)
text_box(s, x1 + Inches(0.1), col_y + Inches(0.05), col_w - Inches(0.2), Inches(0.55),
         "Carlos Garbin — Director\nIT Business Applications  (~26 FTE)",
         size=11, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

carlos_subs = [
    ("Bruno Moreto", "Electric & Gas Ops", "7 FTE", "G&E Champion: Andrew McGibbon", True),
    ("George Porto", "GIS & OSG", "6 FTE", "GIS Champion: George Porto", True),
    ("Sarath Garikapati", "Corporate Applications", "8 FTE", "Corporate Champion: Amit Patel", True),
    ("Vanessa Colombo", "Global SAP — Iberdrola\nFI, PO, Budget (Spain)", "5 FTE",
     "Global SAP Champion: Michael Dubois", True),
]
for i, (name, team, fte, champ, has_champ) in enumerate(carlos_subs):
    y = col_y + Inches(0.75) + i * Inches(0.85)
    round_rect(s, x1, y, col_w, Inches(0.78), LIGHT, line=GREY, line_w=0.5, radius=0.05)
    if has_champ:
        rect(s, x1, y, Inches(0.06), Inches(0.78), GREEN)
    text_box(s, x1 + Inches(0.15), y + Inches(0.05),
             col_w - Inches(0.7), Inches(0.3),
             name, size=10, bold=True, color=DARK)
    text_box(s, x1 + Inches(0.15), y + Inches(0.28),
             col_w - Inches(0.7), Inches(0.35),
             team, size=9, color=DARK)
    rect(s, x1 + col_w - Inches(0.55), y + Inches(0.08),
         Inches(0.5), Inches(0.32), BLUE)
    text_box(s, x1 + col_w - Inches(0.55), y + Inches(0.08),
             Inches(0.5), Inches(0.32), fte,
             size=9, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x1 + Inches(0.15), y + Inches(0.55),
             col_w - Inches(0.2), Inches(0.22),
             champ, size=8, italic=True, color=GREY)

# Column 2 — Kumar
x2 = Inches(0.5) + col_w + gap
round_rect(s, x2, col_y, col_w, Inches(0.65), NAVY, radius=0.04)
text_box(s, x2 + Inches(0.1), col_y + Inches(0.05),
         col_w - Inches(0.2), Inches(0.55),
         "Kumar Jaladanki — Director\nCustomer Services Apps  (~24 FTE)",
         size=11, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

kumar_subs = [
    ("Ivelisse Velazquez", "FICA, Billing & Collection", "8 FTE",
     "SAP Champion: Stephanie", True),
    ("TBD (vacant)", "CRM & Contact Center Ops", "4 FTE",
     "", False),
    ("Vincent Lam", "Billing, Device & Metering", "8 FTE",
     "Smart Metering Champion: Iram Chowdhury", True),
    ("Open (vacant)", "SAP & CS Architecture", "2 FTE",
     "", False),
    ("Nagi Reddy", "SAP Basis", "2 FTE", "", False),
]
for i, (name, team, fte, champ, has_champ) in enumerate(kumar_subs):
    y = col_y + Inches(0.75) + i * Inches(0.66)
    is_vacant = "vacant" in name.lower()
    bg = RED_LIGHT if is_vacant else LIGHT
    round_rect(s, x2, y, col_w, Inches(0.6), bg, line=GREY, line_w=0.5, radius=0.05)
    if has_champ:
        rect(s, x2, y, Inches(0.06), Inches(0.6), GREEN)
    name_color = VACANT if is_vacant else DARK
    text_box(s, x2 + Inches(0.15), y + Inches(0.03),
             col_w - Inches(0.7), Inches(0.25),
             name, size=10, bold=True, color=name_color)
    text_box(s, x2 + Inches(0.15), y + Inches(0.24),
             col_w - Inches(0.7), Inches(0.22),
             team, size=8.5, color=DARK)
    rect(s, x2 + col_w - Inches(0.55), y + Inches(0.05),
         Inches(0.5), Inches(0.28), ACCENT)
    text_box(s, x2 + col_w - Inches(0.55), y + Inches(0.05),
             Inches(0.5), Inches(0.28), fte,
             size=9, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    if champ:
        text_box(s, x2 + Inches(0.15), y + Inches(0.4),
                 col_w - Inches(0.2), Inches(0.2),
                 champ, size=7.5, italic=True, color=GREY)

# Column 3 — Heather
x3 = Inches(0.5) + 2*(col_w + gap)
round_rect(s, x3, col_y, col_w, Inches(0.65), NAVY, radius=0.04)
text_box(s, x3 + Inches(0.1), col_y + Inches(0.05),
         col_w - Inches(0.2), Inches(0.55),
         "Heather LaJoie — Director\nProduct Management  (~2 FTE)",
         size=11, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

heath_y = col_y + Inches(0.75)
round_rect(s, x3, heath_y, col_w, Inches(0.85), RED_LIGHT, line=GREY,
           line_w=0.5, radius=0.05)
text_box(s, x3 + Inches(0.15), heath_y + Inches(0.05),
         col_w - Inches(0.7), Inches(0.3),
         "Vacant", size=10, bold=True, color=VACANT)
text_box(s, x3 + Inches(0.15), heath_y + Inches(0.28),
         col_w - Inches(0.7), Inches(0.35),
         "Product Management", size=9, color=DARK)
rect(s, x3 + col_w - Inches(0.55), heath_y + Inches(0.08),
     Inches(0.5), Inches(0.32), PURPLE)
text_box(s, x3 + col_w - Inches(0.55), heath_y + Inches(0.08),
         Inches(0.5), Inches(0.32), "2 FTE",
         size=9, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Description text below
desc_y = heath_y + Inches(1.0)
desc_h = col_h - Inches(2.5)
round_rect(s, x3, desc_y, col_w, desc_h, BLUE_LIGHT, line=BLUE, line_w=1, radius=0.04)
text_box(s, x3 + Inches(0.15), desc_y + Inches(0.1),
         col_w - Inches(0.3), desc_h - Inches(0.2),
         "This team acts as the strategic coordination layer across delivery and "
         "execution teams, ensuring alignment between business priorities and "
         "technical implementation.\n\n"
         "It is responsible for shaping and supporting the prioritization of work, "
         "providing direction throughout execution. The team enables consistent "
         "focus on value creation, ensuring efforts are aligned with business "
         "outcomes through effective orchestration of dependent teams.",
         size=9, color=DARK, italic=True)

# Column 4 — IT Program Delivery
x4 = Inches(0.5) + 3*(col_w + gap)
round_rect(s, x4, col_y, col_w, Inches(0.65), NAVY, radius=0.04)
text_box(s, x4 + Inches(0.1), col_y + Inches(0.05),
         col_w - Inches(0.2), Inches(0.55),
         "IT Program Delivery\n(Vacant)",
         size=11, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Champion legend at bottom
leg_y = col_y + col_h + Inches(0.1)
rect(s, Inches(0.5), leg_y, Inches(0.25), Inches(0.25), GREEN)
text_box(s, Inches(0.85), leg_y, Inches(4), Inches(0.25),
         "Area champion identified (6)",
         size=10, bold=True, italic=True, color=DARK)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 4 — TOM Overview with applications
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Target Operating Model · eight Value Stream Teams aligned to customer outcomes")

streams_apps = [
    ("Meter-to-Cash · Service Lifecycle",
     "Customer signs up · activates · changes · disconnects",
     "SAP Metering · MV90", ACCENT),
    ("Meter-to-Cash · Revenue Operations",
     "Accurate billing · easy payment · fair recovery",
     "SAP FICA · ECC · StreamServe", ACCENT),
    ("Customer Care & Contact Center",
     "Issue resolution · self-service · voice of customer",
     "SAP CRM · Medallia · Blue Prism", ACCENT),
    ("Field Operations",
     "Field worker / inspector / lineman productivity",
     "Lemur-SF · FWM · ARCOS · Field Force", BLUE),
    ("Engineering & Asset Management",
     "Engineer / planner designs with reliable data",
     "ArcFM · ProjectWise · Bentley · CYME", BLUE),
    ("Network Operations & Control",
     "Grid operator prevents and minimizes outages",
     "Echelon · Cimplicity · SLAM · Switching Orders", BLUE),
    ("Corporate & Regulatory",
     "HR / employee / regulatory user self-service",
     "OpenText · Rate Case Advisor · Utilisphere", BLUE),
    ("Iberdrola Corporate Services",
     "Spain finance / procurement / budget",
     "SAP FI · PO · Budget · S/4HANA", BLUE),
]

grid_y = Inches(1.05)
grid_h = Inches(5.0)
ncols = 4
nrows = 2
gap = Inches(0.12)
col_w = (SW - Inches(1.0) - gap * (ncols - 1)) / ncols
row_h = (grid_h - gap) / nrows

for i, (name, outcome, apps, color) in enumerate(streams_apps):
    row = i // ncols
    col = i % ncols
    x = Inches(0.5) + col * (col_w + gap)
    y = grid_y + row * (row_h + gap)
    round_rect(s, x, y, col_w, row_h, WHITE, line=color, line_w=2, radius=0.04)
    rect(s, x, y, col_w, Inches(0.7), color)
    text_box(s, x + Inches(0.1), y + Inches(0.08),
             col_w - Inches(0.2), Inches(0.55),
             name, size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.15), y + Inches(0.85),
             col_w - Inches(0.3), Inches(0.7),
             outcome, size=10.5, color=DARK, italic=True,
             align=PP_ALIGN.CENTER)
    # Apps section
    rect(s, x + Inches(0.15), y + row_h - Inches(0.95),
         col_w - Inches(0.3), Inches(0.05), color)
    text_box(s, x + Inches(0.15), y + row_h - Inches(0.85),
             col_w - Inches(0.3), Inches(0.25),
             "EXAMPLE APPLICATIONS",
             size=8, bold=True, color=color)
    text_box(s, x + Inches(0.15), y + row_h - Inches(0.6),
             col_w - Inches(0.3), Inches(0.5),
             apps, size=10, color=DARK,
             align=PP_ALIGN.CENTER)

# Legend
leg_y = Inches(6.3)
rect(s, Inches(0.5), leg_y, Inches(0.3), Inches(0.3), ACCENT)
text_box(s, Inches(0.9), leg_y, Inches(4.5), Inches(0.3),
         "Customer Services Applications", size=11, color=DARK,
         anchor=MSO_ANCHOR.MIDDLE)
rect(s, Inches(5.5), leg_y, Inches(0.3), Inches(0.3), BLUE)
text_box(s, Inches(5.9), leg_y, Inches(4.5), Inches(0.3),
         "IT Business Applications", size=11, color=DARK,
         anchor=MSO_ANCHOR.MIDDLE)

footer_bar(s)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 5 — Stage 01 Demand Management overview
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Stage 01 — Demand Management · five framework dimensions")

cat_y = Inches(1.3)
cat_h = Inches(5.1)
ncats = 5
gap = Inches(0.15)
cat_w = (SW - Inches(1.0) - gap * (ncats - 1)) / ncats

categories = [
    ("CAPABILITIES", "14 items", GREEN, GREEN_LIGHT,
     ["What the organization must be able to do",
      "",
      "Intake · classification · scoring · triage · routing · visibility · de-duplication"]),
    ("ROLES", "4 items", TEAL, TEAL_LIGHT,
     ["Who owns and operates the framework",
      "",
      "Requestors · Product/Demand Managers · Architects · Triage Committee"]),
    ("TOOLS", "5 items", AMBER, AMBER_LIGHT,
     ["Systems and platforms used",
      "",
      "Capture · tracking · communication · documentation · reporting"]),
    ("GOVERNANCE", "6 items", INDIGO, INDIGO_LIGHT,
     ["Decisions, forums, controls",
      "",
      "Forums · decision practices · escalation · guidelines · alignment · oversight"]),
    ("PERFORMANCE", "7 items", RED, RED_LIGHT,
     ["How success is measured",
      "",
      "Volume · quality · lead time · channel mix · urgency · duplication · visibility"]),
]
for i, (name, count, color, light, lines) in enumerate(categories):
    x = Inches(0.5) + i * (cat_w + gap)
    round_rect(s, x, cat_y, cat_w, cat_h, WHITE, line=color, line_w=2, radius=0.04)
    rect(s, x, cat_y, cat_w, Inches(0.7), color)
    text_box(s, x, cat_y + Inches(0.05), cat_w, Inches(0.35),
             name, size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x, cat_y + Inches(0.4), cat_w, Inches(0.3),
             count, size=10.5, italic=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_lines(s, x + Inches(0.2), cat_y + Inches(0.9),
              cat_w - Inches(0.4), cat_h - Inches(1.0),
              lines, size=11, color=DARK, line_spacing=1.4)

footer_bar(s)


# ════════════════════════════════════════════════════════════════════════
# Helper: build a category slide with N items in a grid
# ════════════════════════════════════════════════════════════════════════
def category_slide(category_name, count_label, color, items, ncols, action_title):
    s = blank()
    header_bar(s, action_title)
    badge_y = Inches(1.0)
    round_rect(s, Inches(0.5), badge_y, Inches(3.5), Inches(0.5),
               color, radius=0.2)
    text_box(s, Inches(0.5), badge_y, Inches(3.5), Inches(0.5),
             category_name + "  ·  " + count_label,
             size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    grid_y = Inches(1.7)
    grid_h = Inches(5.1)
    nrows = (len(items) + ncols - 1) // ncols
    gap_x = Inches(0.1)
    gap_y = Inches(0.1)
    col_w = (SW - Inches(1.0) - gap_x * (ncols - 1)) / ncols
    row_h = (grid_h - gap_y * (nrows - 1)) / nrows

    for i, (title, body) in enumerate(items):
        row = i // ncols
        col = i % ncols
        x = Inches(0.5) + col * (col_w + gap_x)
        y = grid_y + row * (row_h + gap_y)
        round_rect(s, x, y, col_w, row_h, WHITE, line=color,
                   line_w=1.5, radius=0.05)
        rect(s, x, y, col_w, Inches(0.45), color)
        text_box(s, x + Inches(0.1), y + Inches(0.03),
                 col_w - Inches(0.2), Inches(0.4),
                 title, size=10.5, bold=True, color=WHITE,
                 anchor=MSO_ANCHOR.MIDDLE)
        text_box(s, x + Inches(0.15), y + Inches(0.55),
                 col_w - Inches(0.3), row_h - Inches(0.65),
                 body, size=10, color=DARK, anchor=MSO_ANCHOR.TOP)
    footer_bar(s)
    return s


# ════════════════════════════════════════════════════════════════════════
# SLIDE 6 — CAPABILITIES (14 items) — UPDATED Capability #1 with Jira+IT Now
# ════════════════════════════════════════════════════════════════════════
capabilities = [
    ("Demand intake channels",
     "Jira intake forms — separate for Large, Medium, Small demand. "
     "IT Now reserved for incidents only. Email · Teams · Word · OneNote decommissioned."),
    ("Demand registration & logging",
     "Auto-created backlog item from intake form. Immutable audit fields: requester, "
     "business unit, team, date. Every state change logged."),
    ("Channel consolidation",
     "One adaptive intake form per team. Target: 90% of demand via single channel by month 6. "
     "No alternative formal channels accepted."),
    ("Standard intake formats",
     "Canonical template with mandatory Definition of Ready fields: title, business value, "
     "acceptance criteria, class of service, due date, sponsor."),
    ("Demand classification",
     "Three independent axes: type, size, urgency. Closed taxonomy. "
     "Classification done at intake gate, not self-declared."),
    ("Run · Change · Regulatory · Innovation",
     "Four Classes of Service with capacity caps. Replaces ad-hoc 'urgent override' culture "
     "with explicit policy."),
    ("Business value definition",
     "Weighted scoring formula combining business value, user experience, safety, and effort. "
     "Sponsor signs off the score at intake."),
    ("Acceptance criteria",
     "Mandatory in Definition of Ready. Format: Given/When/Then or measurable outcome. "
     "Reviewed before scoring. Immutable post-prioritization."),
    ("Technical feasibility input",
     "Lightweight check at intake gate by rotating architect. Time-boxed (2 hours max). "
     "Output: Feasible / Conditional / Not feasible."),
    ("Triage & validation",
     "Weekly 30-minute triage committee. Outcomes: Accept · Send-back · Reject. "
     "Five-business-day SLA."),
    ("Demand routing & assignment",
     "Routing by Value Stream, not by team or person. Auto-routing based on requester selection. "
     "Cross-stream items go to Architecture review."),
    ("Urgent or exception requests",
     "Expedite Class of Service capped at 20% of throughput. "
     "Real-time incidents bypass intake (Service Desk SLA)."),
    ("Demand visibility & tracking",
     "Single dashboard, public read-only. Views: inflow funnel · backlog by team · "
     "class-of-service mix · aging · top priorities."),
    ("Duplicate & overlap identification",
     "Similarity search at submit (top-3 matches surfaced). Weekly de-dup query. "
     "Overlaps escalated to Portfolio Committee."),
]
category_slide("CAPABILITIES", "14 items", GREEN,
               capabilities, ncols=5,
               action_title="Capabilities · what the organization must be able to do")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 7 — ROLES (4)
# ════════════════════════════════════════════════════════════════════════
roles = [
    ("Business requestors & stakeholders",
     "All originating business units route through a single intake form per Value Stream. "
     "Includes Customer Service Products organization, Networks Business sponsors, "
     "Iberdrola corporate, and BRS UIL/CMP/NY operating companies."),
    ("Product or Demand Managers",
     "One Product Manager per Value Stream Team — accountable for outcome, scoring, and "
     "backlog priority. One Demand Manager (cross-team) audits scoring and intake quality."),
    ("Architects & technical reviewers",
     "Architecture Council with rotating members from existing technical leads. "
     "Time-boxed advisory input at the intake gate. Provides feasibility view, not a veto."),
    ("Triage / intake committee",
     "Weekly committee per portfolio. Composition: Demand Manager (chair) plus Product "
     "Managers and Architect. Decisions logged with reason. Cross-portfolio escalations "
     "go to monthly Portfolio Committee."),
]
category_slide("ROLES", "4 items", TEAL,
               roles, ncols=2,
               action_title="Roles · who owns and operates the framework")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 8 — TOOLS (5)
# ════════════════════════════════════════════════════════════════════════
tools = [
    ("Demand capture",
     "Jira intake form per Value Stream Team — Large, Medium, Small demand. "
     "Mandatory fields enforce Definition of Ready. IT Now reserved for incidents only."),
    ("Tracking & backlog",
     "Single backlog system as system of record. Item hierarchy: Initiative → Epic → Story. "
     "Team Kanban with five phases and dual track (Discovery + Delivery)."),
    ("Communication channels",
     "Triage outcomes auto-published to requesters. Per-requester portal showing live status. "
     "Cross-team coordination through dedicated forums."),
    ("Documentation",
     "Single source of truth for Definition of Ready, Class of Service taxonomy, "
     "per-team product brief. Sponsor approval logged within the intake form itself."),
    ("Reporting & visibility",
     "Demand dashboard with five views: inflow funnel, backlog by team, class-of-service mix, "
     "aging, top priorities. Auto-refresh under one hour. Director-level rollups."),
]
category_slide("TOOLS", "5 items", AMBER,
               tools, ncols=3,
               action_title="Tools · systems and platforms used")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 9 — GOVERNANCE (6)
# ════════════════════════════════════════════════════════════════════════
governance = [
    ("Demand review forums",
     "Two-gate structure. Gate 1: weekly triage (30 minutes). "
     "Gate 2: monthly Portfolio Committee (90 minutes). "
     "Quarterly cross-portfolio Big Room Planning."),
    ("Decision-making practices",
     "Four explicit outcomes: Accept · Send-back · Reject · Redirect. "
     "Every decision audit-logged with reason. No silent overrides."),
    ("Escalation mechanisms",
     "Cross-team conflicts resolved at Customer Journey Forum. "
     "Cross-portfolio conflicts resolved at VP Portfolio Sync. "
     "Every Expedite reviewed at the next Portfolio Committee."),
    ("Definition of intake guidelines",
     "Single source of truth for the Definition of Ready template, Class of Service taxonomy, "
     "and capacity caps. Reviewed every six months by the Demand Council."),
    ("Alignment with business stakeholders",
     "Quarterly Big Room Planning. Customer Journey Forum per major customer journey. "
     "Sponsor sign-off mandatory at intake (business value approval)."),
    ("Oversight of inflow & backlog health",
     "Demand Manager owns the demand dashboard. Aging beyond threshold triggers automatic "
     "escalation. Quarterly demand-vs-delivery conversion review."),
]
category_slide("GOVERNANCE", "6 items", INDIGO,
               governance, ncols=3,
               action_title="Governance · decisions, forums, and controls")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 10 — PERFORMANCE (7)
# ════════════════════════════════════════════════════════════════════════
performance = [
    ("Demand volume & inflow trends",
     "Submissions per week, month, quarter. Breakdown by Value Stream, Class of Service, "
     "OPCO, and source business area."),
    ("Demand quality",
     "Definition-of-Ready pass-rate at first submission (target 70% or higher). "
     "Re-submission rate per requester. Acceptance criteria quality score."),
    ("Triage lead time",
     "Median time from submission to Gate 1 decision (target 5 business days or less). "
     "90th percentile to surface outliers. Breakdown per Class of Service."),
    ("% via formal channels",
     "Target 90% via single channel by month 6. "
     "Decommissioned-channel residuals tracked monthly. Per-business-unit adoption visible."),
    ("% urgent or bypassed",
     "Expedite as percent of total throughput (cap 20%). Regulatory subset tracked. "
     "Auto-downgrade rate of unauthorized 'urgent' claims."),
    ("Duplication & overlap",
     "Duplicate-link rate (% submissions matched to existing). "
     "Overlap escalations resolved at Portfolio Committee."),
    ("Visibility & traceability",
     "Percent of requesters who can find own submission status (target 100%). "
     "Average time to status update. Live per-business-unit portal."),
]
category_slide("PERFORMANCE", "7 items", RED,
               performance, ncols=4,
               action_title="Performance · how success is measured")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 11 — Roles & Responsibilities (with names)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Roles & Responsibilities · current people and proposed roles")

# Top: existing leadership row
ex_y = Inches(1.05)
ex_h = Inches(2.2)
text_box(s, Inches(0.5), ex_y - Inches(0.1), Inches(8), Inches(0.3),
         "EXISTING LEADERSHIP",
         size=12, bold=True, color=NAVY)

leadership = [
    ("Vitor Cruz", "VP IT Business Solutions", "Strategy · sponsor · cross-portfolio sync", NAVY),
    ("Carlos Garbin", "Director · IT Business Apps", "IT Apps Portfolio Board owner", BLUE),
    ("Kumar Jaladanki", "Director · Customer Services Apps", "CS Apps Portfolio Board owner", ACCENT),
    ("Heather LaJoie", "Director · Product Management", "Cross-base prioritization layer", PURPLE),
    ("Travis · PMO Lead", "Large Projects PDD", "Bridge to legacy PDD process", GREEN),
]
ldr_w = (SW - Inches(1)) / 5
for i, (name, title, role, color) in enumerate(leadership):
    x = Inches(0.5) + i * ldr_w
    round_rect(s, x + Inches(0.05), ex_y, ldr_w - Inches(0.1), ex_h,
               WHITE, line=color, line_w=2, radius=0.04)
    rect(s, x + Inches(0.05), ex_y, ldr_w - Inches(0.1), Inches(0.5), color)
    text_box(s, x + Inches(0.1), ex_y + Inches(0.05),
             ldr_w - Inches(0.2), Inches(0.4),
             name, size=11, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.15), ex_y + Inches(0.6),
             ldr_w - Inches(0.3), Inches(0.5),
             title, size=10, bold=True, color=DARK,
             align=PP_ALIGN.CENTER)
    text_box(s, x + Inches(0.15), ex_y + Inches(1.15),
             ldr_w - Inches(0.3), ex_h - Inches(1.25),
             role, size=9.5, color=GREY, italic=True,
             align=PP_ALIGN.CENTER)

# Middle: New roles needed
new_y = Inches(3.55)
new_h = Inches(2.2)
text_box(s, Inches(0.5), new_y - Inches(0.1), Inches(8), Inches(0.3),
         "NEW ROLES TO ESTABLISH",
         size=12, bold=True, color=ACCENT)

new_roles = [
    ("Product Manager × 8",
     "One per Value Stream Team",
     "Outcome owner · scoring · backlog priority",
     ACCENT),
    ("Demand Manager",
     "Cross-team",
     "Audits scoring · intake quality · dashboard",
     ACCENT),
    ("Portfolio Manager",
     "Cross-portfolio",
     "VP Dashboard · capacity allocation · monthly sync",
     ACCENT),
    ("UX Designer",
     "Cross-team",
     "Customer journey research · usability testing",
     ACCENT),
    ("Knowledge Manager",
     "Cross-team",
     "Documentation · onboarding · lessons learned",
     ACCENT),
]
nr_w = (SW - Inches(1)) / 5
for i, (name, title, role, color) in enumerate(new_roles):
    x = Inches(0.5) + i * nr_w
    round_rect(s, x + Inches(0.05), new_y, nr_w - Inches(0.1), new_h,
               WHITE, line=color, line_w=2, radius=0.04)
    rect(s, x + Inches(0.05), new_y, nr_w - Inches(0.1), Inches(0.5), color)
    text_box(s, x + Inches(0.1), new_y + Inches(0.05),
             nr_w - Inches(0.2), Inches(0.4),
             name, size=11, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.15), new_y + Inches(0.6),
             nr_w - Inches(0.3), Inches(0.5),
             title, size=10, bold=True, color=DARK,
             align=PP_ALIGN.CENTER)
    text_box(s, x + Inches(0.15), new_y + Inches(1.15),
             nr_w - Inches(0.3), new_h - Inches(1.25),
             role, size=9.5, color=GREY, italic=True,
             align=PP_ALIGN.CENTER)

# Bottom: existing champions / SMEs
ch_y = Inches(6.0)
ch_h = Inches(0.95)
text_box(s, Inches(0.5), ch_y - Inches(0.1), Inches(10), Inches(0.3),
         "AREA CHAMPIONS — candidates to re-skill into Product Manager roles",
         size=12, bold=True, color=GREEN)

champs = [
    "Iram Chowdhury — Smart Metering",
    "Andrew McGibbon — G&E",
    "George Porto — GIS",
    "Michael Dubois — Global SAP",
    "Stephanie — SAP / CS",
    "Amit Patel — Corporate",
]
ch_w = (SW - Inches(1)) / 6
for i, name in enumerate(champs):
    x = Inches(0.5) + i * ch_w
    round_rect(s, x + Inches(0.05), ch_y, ch_w - Inches(0.1), ch_h,
               GREEN_LIGHT, line=GREEN, line_w=1, radius=0.05)
    text_box(s, x + Inches(0.1), ch_y + Inches(0.05),
             ch_w - Inches(0.2), ch_h - Inches(0.1),
             name, size=10, bold=True, color=DARK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

footer_bar(s)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 12 — Forums (5 forums visual)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Forums · five cross-team coordination venues with explicit cadences")

forums = [
    ("REGIONAL FORUM",
     "Quarterly", BLUE,
     "Per OPCO discussion",
     ["BRS UIL — UI · SCG · CNG · BGC (Connecticut)",
      "BRS CMP — Central Maine Power",
      "BRS NY — NYSEG · RGE",
      "Members: Value Stream PMs + business sponsors per OPCO"]),
    ("FUNCTIONAL FORUM",
     "Bi-weekly", TEAL,
     "Per role / function",
     ["PM Council (cross-stream Product Managers)",
      "Tech Lead Council (cross-stream architecture review)",
      "Project Manager pool review (large-project coordination)",
      "Members: same role across teams"]),
    ("TECHNOLOGICAL FORUM",
     "Bi-weekly", AMBER,
     "Per technology stack",
     ["SAP Forum — Local + Iberdrola",
      "Non-SAP Forum — GIS · Mobility · Document Mgmt",
      "Platform Forum — DevOps · Security · Test automation",
      "Members: technical leads + platform owners"]),
    ("CUSTOMER JOURNEY FORUM",
     "Weekly during active initiative", PURPLE,
     "End-to-end customer experience",
     ["M2C Journey — meter → bill → payment",
      "Self-Service Journey — portal · CRM · Contact Center",
      "Field Worker Journey — mobility · doc viewer · storm",
      "Members: PMs + UX + business sponsors"]),
    ("PRODUCT FORUM",
     "Quarterly (Big Room Planning)", GREEN,
     "Cross-stream product strategy",
     ["BRP cadence with all Value Stream Teams",
      "Cross-stream roadmap sync",
      "Trade-off + sequencing decisions at portfolio level",
      "Members: Directors + PMs + Portfolio Manager"]),
]

f_y = Inches(1.1)
f_h = Inches(5.4)
ncols = 5
gap = Inches(0.12)
f_w = (SW - Inches(1.0) - gap * (ncols - 1)) / ncols

for i, (name, cad, color, sub, items) in enumerate(forums):
    x = Inches(0.5) + i * (f_w + gap)
    round_rect(s, x, f_y, f_w, f_h, WHITE, line=color, line_w=2, radius=0.04)
    rect(s, x, f_y, f_w, Inches(0.7), color)
    text_box(s, x, f_y + Inches(0.05), f_w, Inches(0.35),
             name, size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x, f_y + Inches(0.4), f_w, Inches(0.3),
             cad, size=9.5, italic=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Sub
    text_box(s, x + Inches(0.1), f_y + Inches(0.85),
             f_w - Inches(0.2), Inches(0.4),
             sub, size=10.5, bold=True, color=color,
             align=PP_ALIGN.CENTER)
    # Items
    add_lines(s, x + Inches(0.18), f_y + Inches(1.4),
              f_w - Inches(0.36), f_h - Inches(1.5),
              ["•  " + it for it in items], size=9.5, color=DARK,
              line_spacing=1.4)

footer_bar(s)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 13 — Portfolio Architecture (3 layers · FIXED design)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Portfolio architecture · three layers from VP rollup to team execution")

# Layer 0
l0_y = Inches(1.1)
round_rect(s, Inches(2.5), l0_y, Inches(8.3), Inches(0.85),
           BLUE_LIGHT, line=BLUE, line_w=2, radius=0.04)
text_box(s, Inches(2.7), l0_y + Inches(0.05), Inches(8), Inches(0.4),
         "VP ROLL-UP DASHBOARD",
         size=14, bold=True, color=BLUE)
text_box(s, Inches(2.7), l0_y + Inches(0.4), Inches(8), Inches(0.4),
         "Aggregated view across both portfolios · monthly Portfolio Sync",
         size=10, color=DARK, italic=True)

# Down arrow
shp = s.shapes.add_connector(1, Inches(6.65), l0_y + Inches(0.85),
                              Inches(6.65), Inches(2.15))
shp.line.color.rgb = NAVY
shp.line.width = Pt(2)

# Layer 1 — 2 Portfolio Boards
l1_y = Inches(2.15)
pb_w = Inches(6)
pb_h = Inches(1.2)

round_rect(s, Inches(0.5), l1_y, pb_w, pb_h, WHITE, line=ACCENT, line_w=2, radius=0.04)
rect(s, Inches(0.5), l1_y, pb_w, Inches(0.45), ACCENT)
text_box(s, Inches(0.5), l1_y + Inches(0.05), pb_w, Inches(0.35),
         "PORTFOLIO BOARD  ·  Customer Services Applications",
         size=12, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text_box(s, Inches(0.7), l1_y + Inches(0.5), pb_w - Inches(0.4), Inches(0.65),
         "3 Value Streams · cards = Initiatives  ·  Owner: Kumar  ·  monthly Portfolio Committee",
         size=10, color=DARK)

round_rect(s, Inches(6.85), l1_y, pb_w, pb_h, WHITE, line=BLUE, line_w=2, radius=0.04)
rect(s, Inches(6.85), l1_y, pb_w, Inches(0.45), BLUE)
text_box(s, Inches(6.85), l1_y + Inches(0.05), pb_w, Inches(0.35),
         "PORTFOLIO BOARD  ·  IT Business Applications",
         size=12, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text_box(s, Inches(7.05), l1_y + Inches(0.5), pb_w - Inches(0.4), Inches(0.65),
         "5 Value Streams · cards = Initiatives  ·  Owner: Carlos  ·  monthly Portfolio Committee",
         size=10, color=DARK)

# Down arrow
shp = s.shapes.add_connector(1, Inches(6.65), l1_y + pb_h,
                              Inches(6.65), Inches(3.85))
shp.line.color.rgb = NAVY
shp.line.width = Pt(2)

# Layer 2 — 8 team boards · FIXED design with proper spacing
l2_y = Inches(4.1)
text_box(s, Inches(0.5), l2_y - Inches(0.3), SW - Inches(1), Inches(0.3),
         "EIGHT VALUE STREAM TEAM BOARDS  ·  cards = Epics → Stories  ·  five phases × dual track",
         size=13, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER)

teams = [
    ("M2C-SL", "Service Lifecycle", ACCENT),
    ("M2C-RO", "Revenue Operations", ACCENT),
    ("CCC", "Customer Care", ACCENT),
    ("FOPS", "Field Operations", BLUE),
    ("E&A", "Engineering & Asset", BLUE),
    ("NOC", "Network Operations", BLUE),
    ("C&R", "Corporate & Regulatory", BLUE),
    ("IBD", "Iberdrola Corp", BLUE),
]
team_w = (SW - Inches(1.0)) / 8
for i, (short, full, color) in enumerate(teams):
    x = Inches(0.5) + i * team_w
    round_rect(s, x + Inches(0.05), l2_y, team_w - Inches(0.1),
               Inches(2.4), WHITE, line=color, line_w=1.5, radius=0.04)
    rect(s, x + Inches(0.05), l2_y, team_w - Inches(0.1), Inches(0.45), color)
    text_box(s, x + Inches(0.05), l2_y + Inches(0.05),
             team_w - Inches(0.1), Inches(0.4),
             short, size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.1), l2_y + Inches(0.55),
             team_w - Inches(0.2), Inches(0.6),
             full, size=10, bold=True, color=DARK,
             align=PP_ALIGN.CENTER)
    text_box(s, x + Inches(0.1), l2_y + Inches(1.25),
             team_w - Inches(0.2), Inches(1.1),
             "5 phases\n×\n2 tracks\nDiscovery + Delivery",
             size=9.5, color=GREY, italic=True,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

footer_bar(s)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 14 — Portfolio Board mockup (Initiative flow)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Portfolio Board · cards are Initiatives flowing through stages")

text_box(s, Inches(0.5), Inches(1.0), SW - Inches(1), Inches(0.4),
         "Example  ·  IT Business Applications Portfolio Board",
         size=13, italic=True, color=GREY, align=PP_ALIGN.CENTER)

stage_y = Inches(1.6)
stage_h = Inches(4.4)
stages = ["Funnel", "Reviewing", "Analyzing", "Backlog", "Implementing", "Done"]
ncols = len(stages)
gap = Inches(0.06)
col_w = (SW - Inches(1.0) - gap * (ncols - 1)) / ncols

for i, stg in enumerate(stages):
    x = Inches(0.5) + i * (col_w + gap)
    rect(s, x, stage_y, col_w, Inches(0.5), NAVY_DARK)
    text_box(s, x, stage_y + Inches(0.05), col_w, Inches(0.4),
             stg, size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, x, stage_y + Inches(0.5), col_w, stage_h - Inches(0.5),
         LIGHT, line=GREY, line_w=0.5)

def init_card(slide, x, y, w, h, title, owner, cos_color):
    round_rect(slide, x, y, w, h, WHITE, line=GREY, line_w=0.5, radius=0.06)
    rect(slide, x, y, Inches(0.08), h, cos_color)
    text_box(slide, x + Inches(0.15), y + Inches(0.05),
             w - Inches(0.2), Inches(0.4),
             title, size=10, bold=True, color=DARK)
    text_box(slide, x + Inches(0.15), y + Inches(0.4),
             w - Inches(0.2), Inches(0.3),
             owner, size=9, color=GREY, italic=True)

# Cards
init_card(s, Inches(0.5) + Inches(0.05), stage_y + Inches(0.6),
          col_w - Inches(0.1), Inches(0.7), "Storm program v2", "Field Operations", AMBER)
init_card(s, Inches(0.5) + Inches(0.05), stage_y + Inches(1.4),
          col_w - Inches(0.1), Inches(0.7), "Mobility upgrades", "Field Operations", GREEN)
init_card(s, Inches(0.5) + col_w + gap + Inches(0.05), stage_y + Inches(0.6),
          col_w - Inches(0.1), Inches(0.7), "Drawing platform", "Engineering & Asset", GREEN)
init_card(s, Inches(0.5) + 2*(col_w + gap) + Inches(0.05), stage_y + Inches(0.6),
          col_w - Inches(0.1), Inches(0.7), "Network deployment", "Engineering & Asset", GREEN)
init_card(s, Inches(0.5) + 2*(col_w + gap) + Inches(0.05), stage_y + Inches(1.4),
          col_w - Inches(0.1), Inches(0.7), "OMS modernization", "Network Ops", RED)
init_card(s, Inches(0.5) + 3*(col_w + gap) + Inches(0.05), stage_y + Inches(0.6),
          col_w - Inches(0.1), Inches(0.7), "Doc viewer scale", "Field Operations", GREEN)
init_card(s, Inches(0.5) + 4*(col_w + gap) + Inches(0.05), stage_y + Inches(0.6),
          col_w - Inches(0.1), Inches(0.7), "Mobile field migration", "Field Operations", GREEN)
init_card(s, Inches(0.5) + 4*(col_w + gap) + Inches(0.05), stage_y + Inches(1.4),
          col_w - Inches(0.1), Inches(0.7), "ERP transition", "Iberdrola Corp", BLUE)
init_card(s, Inches(0.5) + 4*(col_w + gap) + Inches(0.05), stage_y + Inches(2.2),
          col_w - Inches(0.1), Inches(0.7), "Asset registry refresh", "Engineering & Asset", GREEN)
init_card(s, Inches(0.5) + 5*(col_w + gap) + Inches(0.05), stage_y + Inches(0.6),
          col_w - Inches(0.1), Inches(0.7), "Storm program Q1", "Field Operations", GREEN)

# Approval gate
gate_x = Inches(0.5) + 3*(col_w + gap) - Inches(0.05)
rect(s, gate_x, stage_y - Inches(0.05), Inches(0.05), stage_h + Inches(0.1), RED)
text_box(s, gate_x - Inches(1.5), stage_y - Inches(0.05) + Inches(0.1),
         Inches(3), Inches(0.3),
         "▲ APPROVAL GATE", size=10, bold=True, color=RED, align=PP_ALIGN.CENTER)

# Legend
leg_y = Inches(6.3)
text_box(s, Inches(0.5), leg_y, Inches(2), Inches(0.3),
         "CLASS OF SERVICE", size=10, bold=True, color=GREY)
items = [(GREEN, "Standard"), (BLUE, "Fixed-Date"), (RED, "Expedite"), (AMBER, "Intangible")]
for i, (color, label) in enumerate(items):
    x = Inches(2.5) + i * Inches(2.3)
    rect(s, x, leg_y + Inches(0.05), Inches(0.2), Inches(0.2), color)
    text_box(s, x + Inches(0.3), leg_y, Inches(2), Inches(0.3),
             label, size=10, color=DARK)

footer_bar(s)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 15 — Team Kanban with Dual Track (simplified · Epic level only)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Team Kanban · five phases by two tracks · cards are Epics with linked Stories")

text_box(s, Inches(0.5), Inches(1.0), SW - Inches(1), Inches(0.4),
         "Example  ·  Field Operations Team Board  ·  inspired by P&T Lean Operating Model",
         size=12, italic=True, color=GREY, align=PP_ALIGN.CENTER)

ph_y = Inches(1.6)
ph_h = Inches(4.4)
phases_lbl = ["Phase 0\nDISCOVERY", "Phase 1\nMATERIALIZATION",
              "Phase 2\nCREATION", "Phase 3\nONGOING", "Phase 4\nMAINTENANCE"]
nph = len(phases_lbl)
gap = Inches(0.06)
left_panel = Inches(1.4)
ph_w = (SW - Inches(1.0) - left_panel - gap * (nph - 1)) / nph

for i, p in enumerate(phases_lbl):
    x = Inches(0.5) + left_panel + i * (ph_w + gap)
    rect(s, x, ph_y, ph_w, Inches(0.65), NAVY_DARK)
    lines = p.split("\n")
    text_box(s, x, ph_y + Inches(0.05), ph_w, Inches(0.3),
             lines[0], size=10, italic=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x, ph_y + Inches(0.3), ph_w, Inches(0.3),
             lines[1], size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, x, ph_y + Inches(0.65), ph_w, ph_h - Inches(0.65),
         LIGHT, line=GREY, line_w=0.5)

# Track labels
mid_y = ph_y + Inches(0.65) + (ph_h - Inches(0.65)) / 2
rect(s, Inches(0.5), ph_y + Inches(0.65), left_panel, mid_y - ph_y - Inches(0.65), BLUE)
text_box(s, Inches(0.5), ph_y + Inches(0.65), left_panel, mid_y - ph_y - Inches(0.65),
         "DISCOVERY\nTRACK\n\nalways-on\nresearch ·\nvalidation",
         size=10, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

rect(s, Inches(0.5), mid_y, left_panel, ph_y + ph_h - mid_y, AMBER)
text_box(s, Inches(0.5), mid_y, left_panel, ph_y + ph_h - mid_y,
         "DELIVERY\nTRACK\n\nsprint cadence\nplanning ·\nbuild · validate",
         size=10, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

shp = s.shapes.add_connector(1, Inches(0.5) + left_panel, mid_y,
                              SW - Inches(0.5), mid_y)
shp.line.color.rgb = DARK
shp.line.width = Pt(2)

def epic_card(slide, x, y, w, h, title, owner, story_count, color):
    round_rect(slide, x, y, w, h, WHITE, line=color, line_w=1.5, radius=0.06)
    rect(slide, x, y, w, Inches(0.25), color)
    text_box(slide, x + Inches(0.08), y + Inches(0.02),
             w - Inches(0.5), Inches(0.22),
             "EPIC", size=8, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    # Story badge on right
    rect(slide, x + w - Inches(0.45), y + Inches(0.02), Inches(0.4), Inches(0.2), WHITE)
    text_box(slide, x + w - Inches(0.45), y + Inches(0.02),
             Inches(0.4), Inches(0.2), str(story_count) + " S",
             size=8, bold=True, color=color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(slide, x + Inches(0.1), y + Inches(0.3),
             w - Inches(0.2), h - Inches(0.5),
             title, size=10, bold=True, color=DARK)
    text_box(slide, x + Inches(0.1), y + h - Inches(0.25),
             w - Inches(0.2), Inches(0.2),
             owner, size=8, color=GREY, italic=True)

def x_at(idx):
    return Inches(0.5) + left_panel + idx * (ph_w + gap)

# Discovery row Epics
epic_card(s, x_at(0) + Inches(0.05), ph_y + Inches(0.85),
          ph_w - Inches(0.1), Inches(0.95),
          "Migrate DLI inspections", "exploring offline reqs", 6, BLUE)
epic_card(s, x_at(1) + Inches(0.05), ph_y + Inches(0.85),
          ph_w - Inches(0.1), Inches(0.95),
          "Migrate TLI inspections", "sizing with vendor", 8, BLUE)

# Delivery row Epics
epic_card(s, x_at(2) + Inches(0.05), mid_y + Inches(0.15),
          ph_w - Inches(0.1), Inches(0.95),
          "Migrate DGI/TGI", "team assigned · sprint plan", 12, BLUE)
epic_card(s, x_at(3) + Inches(0.05), mid_y + Inches(0.15),
          ph_w - Inches(0.1), Inches(0.95),
          "Offline DLI capability", "in dev · Sprint 5", 5, BLUE)
epic_card(s, x_at(3) + Inches(0.05), mid_y + Inches(1.25),
          ph_w - Inches(0.1), Inches(0.85),
          "Mobile TLI UI", "in dev · Sprint 5", 4, BLUE)
epic_card(s, x_at(4) + Inches(0.05), mid_y + Inches(0.15),
          ph_w - Inches(0.1), Inches(0.95),
          "Decommission FWM", "hypercare · 30d", 3, BLUE)

# Legend
leg_y = Inches(6.3)
items = [
    ("EPIC", BLUE, "1-3 months · Initiative breakdown"),
    ("Story badge", GREY, "Stories tracked in Sprint sub-view"),
    ("Card flow", DARK, "items move left → right through phases"),
]
for i, (label, color, desc) in enumerate(items):
    x = Inches(0.5) + i * Inches(4.3)
    rect(s, x, leg_y + Inches(0.05), Inches(0.2), Inches(0.2), color)
    text_box(s, x + Inches(0.3), leg_y, Inches(4), Inches(0.15),
             label, size=10, bold=True, color=DARK)
    text_box(s, x + Inches(0.3), leg_y + Inches(0.18), Inches(4), Inches(0.15),
             desc, size=9, color=GREY, italic=True)

footer_bar(s)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 16 — Class of Service framework
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Class of Service framework · capacity caps replace ad-hoc 'urgent' overrides")

cos_y = Inches(1.2)
cos_h = Inches(5.0)
cos_w = (SW - Inches(1.2)) / 4
gap = Inches(0.15)

cos_list = [
    ("STANDARD", GREEN, "≈50% baseline",
     "Normal demand flowing through the standard intake gate.",
     "Routine enhancements, planned change, tech debt addressed in cycle."),
    ("FIXED-DATE", BLUE, "≈20%",
     "Demand with a hard regulatory or contractual deadline.",
     "Prioritized by date, not by score. Date breach triggers escalation."),
    ("EXPEDITE", RED, "≤20%  ·  CAPPED",
     "Critical regulatory or safety items requiring fast-track.",
     "Bypasses Gate 1, logged for review at next Portfolio Committee. Cap is enforced."),
    ("INTANGIBLE", ACCENT, "≈10% reserved",
     "Innovation, research, user-experience exploration.",
     "Protects the strategic-partner positioning beyond pure delivery."),
]

for i, (name, color, cap, what, when) in enumerate(cos_list):
    x = Inches(0.5) + i * (cos_w + gap)
    round_rect(s, x, cos_y, cos_w, cos_h, WHITE, line=color, line_w=2, radius=0.04)
    rect(s, x, cos_y, cos_w, Inches(0.7), color)
    text_box(s, x, cos_y + Inches(0.05), cos_w, Inches(0.35),
             name, size=15, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x, cos_y + Inches(0.42), cos_w, Inches(0.3),
             cap, size=11, italic=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.2), cos_y + Inches(0.85),
             cos_w - Inches(0.4), Inches(1.0),
             what, size=11, color=DARK)
    rect(s, x + Inches(0.2), cos_y + Inches(2.1),
         cos_w - Inches(0.4), Inches(0.05), color)
    text_box(s, x + Inches(0.2), cos_y + Inches(2.25),
             cos_w - Inches(0.4), Inches(0.3),
             "WHEN IT APPLIES", size=9, bold=True, color=color)
    text_box(s, x + Inches(0.2), cos_y + Inches(2.55),
             cos_w - Inches(0.4), cos_h - Inches(2.7),
             when, size=11, color=DARK)

cb_y = Inches(6.5)
round_rect(s, Inches(0.5), cb_y, SW - Inches(1), Inches(0.55),
           NAVY_DARK, radius=0.1)
text_box(s, Inches(0.7), cb_y + Inches(0.05), SW - Inches(1.4), Inches(0.45),
         "Capacity caps are enforced monthly at the Portfolio Committee · "
         "Expedite breach triggers a governance escalation, not a quiet override.",
         size=12, bold=True, italic=True, color=WHITE,
         anchor=MSO_ANCHOR.MIDDLE)

footer_bar(s)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 17 — Roadmap (Wave 1 · 5 steps over 14 weeks · plus Wave 2 / 3)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Roadmap · Wave 1 in five steps over fourteen weeks · scale in subsequent waves")

# Wave 1 — 5 steps
w1_y = Inches(1.1)
w1_h = Inches(2.7)
text_box(s, Inches(0.5), w1_y - Inches(0.05), Inches(8), Inches(0.3),
         "WAVE 1  ·  Demand Management Design & Pilot  ·  14 weeks",
         size=13, bold=True, color=NAVY)

steps = [
    ("Step 1", "Weeks 1-2", "ENABLEMENT", BLUE,
     "Where are we starting from?",
     "Stakeholder interviews · prior materials review · portfolio assessment · baseline report"),
    ("Step 2", "Weeks 3-4", "DEFINE LIFECYCLE", BLUE,
     "How should we work?",
     "End-to-end process map · Value Framework v1.0 · Roles & RACI Matrix · Demand Typology"),
    ("Step 3", "Weeks 5-6", "DESIGN TOOL ECOSYSTEM", GREEN,
     "What supports the process?",
     "Jira Configuration v1.0 · Intake Form Design · Portfolio Concept Blueprint"),
    ("Step 4", "Weeks 7-13", "PILOT EXECUTION", AMBER,
     "Does it work in reality?",
     "Execute one Value Stream end-to-end · pilot team training · real-time feedback log"),
    ("Step 5", "Week 14", "LEARN & ADJUST", RED,
     "What needs refinement?",
     "Analyze pilot data · refine Jira config v2.0 · adjust Value Scoring · finalize Playbook"),
]
nstp = len(steps)
gap = Inches(0.08)
stp_w = (SW - Inches(1.0) - gap * (nstp - 1)) / nstp

for i, (step, when, what, color, q, desc) in enumerate(steps):
    x = Inches(0.5) + i * (stp_w + gap)
    round_rect(s, x, w1_y + Inches(0.3), stp_w, w1_h - Inches(0.3),
               WHITE, line=color, line_w=2, radius=0.05)
    rect(s, x, w1_y + Inches(0.3), stp_w, Inches(0.4), color)
    text_box(s, x, w1_y + Inches(0.32), stp_w, Inches(0.35),
             step, size=11, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x, w1_y + Inches(0.75), stp_w, Inches(0.3),
             when, size=10, italic=True, color=color,
             align=PP_ALIGN.CENTER)
    text_box(s, x + Inches(0.1), w1_y + Inches(1.05),
             stp_w - Inches(0.2), Inches(0.4),
             what, size=11, bold=True, color=DARK,
             align=PP_ALIGN.CENTER)
    text_box(s, x + Inches(0.1), w1_y + Inches(1.5),
             stp_w - Inches(0.2), Inches(0.35),
             q, size=9, italic=True, color=GREY,
             align=PP_ALIGN.CENTER)
    text_box(s, x + Inches(0.15), w1_y + Inches(1.85),
             stp_w - Inches(0.3), Inches(1.0),
             desc, size=9, color=DARK,
             align=PP_ALIGN.CENTER)

    # Arrow to next
    if i < nstp - 1:
        text_box(s, x + stp_w - Inches(0.05), w1_y + Inches(0.3) + (w1_h - Inches(0.3))/2 - Inches(0.15),
                 Inches(0.15), Inches(0.3), "→",
                 size=18, bold=True, color=GREY,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Wave 2 + Wave 3
w23_y = Inches(4.2)
w23_h = Inches(1.2)

round_rect(s, Inches(0.5), w23_y, Inches(6.3), w23_h,
           BLUE_LIGHT, line=BLUE, line_w=2, radius=0.04)
rect(s, Inches(0.5), w23_y, Inches(6.3), Inches(0.4), BLUE)
text_box(s, Inches(0.5), w23_y + Inches(0.03),
         Inches(6.3), Inches(0.35),
         "WAVE 2  ·  Scale to remaining Value Stream Teams",
         size=12, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text_box(s, Inches(0.7), w23_y + Inches(0.5),
         Inches(6.0), w23_h - Inches(0.55),
         "Apply validated playbook to next 3-4 teams · refined intake forms · "
         "trained Product Managers · ~12 weeks · Q3 2026",
         size=10.5, color=DARK, italic=True)

round_rect(s, Inches(7.0), w23_y, Inches(5.8), w23_h,
           GREEN_LIGHT, line=GREEN, line_w=2, radius=0.04)
rect(s, Inches(7.0), w23_y, Inches(5.8), Inches(0.4), GREEN)
text_box(s, Inches(7.0), w23_y + Inches(0.03),
         Inches(5.8), Inches(0.35),
         "WAVE 3  ·  Steady-state operation",
         size=12, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text_box(s, Inches(7.2), w23_y + Inches(0.5),
         Inches(5.5), w23_h - Inches(0.55),
         "All 8 Value Stream Teams operating model · continuous improvement · "
         "Q4 2026 onwards · quarterly Big Room Planning",
         size=10.5, color=DARK, italic=True)

# Pilot recommendation
pr_y = Inches(5.6)
round_rect(s, Inches(0.5), pr_y, SW - Inches(1), Inches(1.4),
           BLUE_LIGHT, line=NAVY, line_w=2, radius=0.04)
text_box(s, Inches(0.7), pr_y + Inches(0.1),
         SW - Inches(1.4), Inches(0.35),
         "PILOT TEAM RECOMMENDATION",
         size=12, bold=True, color=NAVY)
text_box(s, Inches(0.7), pr_y + Inches(0.4),
         SW - Inches(1.4), Inches(0.45),
         "Meter-to-Cash · Service Lifecycle (Smart Metering)",
         size=18, bold=True, color=NAVY_DARK)
text_box(s, Inches(0.7), pr_y + Inches(0.85),
         SW - Inches(1.4), Inches(0.5),
         "Closest existing analog to a Product Manager already in place — minimal change-management lift. "
         "Vendor delivery model already stable. Clear, scoped user base.",
         size=10, color=DARK, italic=True)

footer_bar(s)


# ════════════════════════════════════════════════════════════════════════
# Save
# ════════════════════════════════════════════════════════════════════════
# ════════════════════════════════════════════════════════════════════════
# SLIDE — Guiding Principles (connect to AS-IS deliverable vrevisada)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Guiding Principles · the four pillars that anchor the product-centric model")

principles = [
    ("01", "PRODUCT-ALIGNED UNIQUE INTAKE LAYER", BLUE,
     ["Single front door per product or value stream",
      "Standard intake criteria with mandatory data inputs",
      "Defined ownership for demand triage and routing",
      "Full visibility of demand within product backlogs"]),
    ("02", "VALUE-BASED PRIORITIZATION", ACCENT,
     ["Common scoring model with four dimensions: value, urgency, risk, effort",
      "Transparent prioritization across competing demands",
      "Product-level backlog ownership · each team owns its priorities",
      "Trade-offs escalated at portfolio governance level"]),
    ("03", "PRODUCT PORTFOLIO MANAGEMENT & PLANNING", GREEN,
     ["Integrated portfolio view spanning all products",
      "Capacity planning aligned to product priorities, not project timelines",
      "End-to-end visibility from demand through delivery",
      "Coordinated planning across capacity and dependencies"]),
    ("04", "PRODUCT GOVERNANCE & ROLES", PURPLE,
     ["Defined roles · accountability · decision rights per layer",
      "Forum cadences for cross-stream coordination",
      "Two-gate governance: weekly triage + monthly portfolio",
      "Sponsor sign-off mandatory at intake"]),
]

p_y = Inches(1.1)
p_h = Inches(5.6)
ncols = 4
gap = Inches(0.15)
p_w = (SW - Inches(1.0) - gap * (ncols - 1)) / ncols

for i, (num, name, color, bullets) in enumerate(principles):
    x = Inches(0.5) + i * (p_w + gap)
    round_rect(s, x, p_y, p_w, p_h, WHITE, line=color, line_w=2, radius=0.04)
    rect(s, x, p_y, p_w, Inches(0.95), color)
    text_box(s, x, p_y + Inches(0.1), p_w, Inches(0.4),
             num, size=24, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.1), p_y + Inches(0.5), p_w - Inches(0.2), Inches(0.45),
             name, size=11, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_lines(s, x + Inches(0.2), p_y + Inches(1.15),
              p_w - Inches(0.4), p_h - Inches(1.3),
              ["•  " + b for b in bullets], size=10.5, color=DARK,
              line_spacing=1.4)

footer_bar(s)


# ════════════════════════════════════════════════════════════════════════
# SLIDE — Demand Management Practices (Ana's five blocks)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Demand Management · five core practices that reshape the intake conversation")

practices = [
    ("SINGLE FRONT DOOR FOR DEMAND",
     "All demand types in one place",
     ["Incidents",
      "Enhancements",
      "Regulatory items",
      "Innovation initiatives",
      "Technical debt"], BLUE),
    ("DEMAND SHAPING",
     "Structured problem-impact-value conversations",
     ["Don't accept 'solutions' directly",
      "Work with requesters to surface the underlying problem",
      "Quantify impact before discussing solution",
      "Apply value lens before scoping work"], ACCENT),
    ("CLASSIFICATION BY PRODUCT TYPE",
     "Two complementary axes",
     ["Run · Grow · Transform   (investment intent)",
      "Core SAP   vs   Extensions / satellites   (architecture impact)",
      "Both axes captured at intake form",
      "Drives capacity allocation and routing"], GREEN),
    ("LIGHTWEIGHT BUT FORMAL DISCOVERY",
     "Required artifacts before scoring",
     ["Problem statement (one paragraph)",
      "Users / personas affected",
      "Hypothesis of value",
      "Hypothesis of risk"], PURPLE),
    ("VALUE & PRIORITIZATION FRAMEWORK",
     "Multi-factor score, not single judgment",
     ["Business value (revenue · cost · strategic fit)",
      "Risk (operational · regulatory · technical)",
      "Regulatory urgency (deadline-driven items)",
      "Cost of Delay (value lost per period of delay)"], RED),
]

pr_y = Inches(1.1)
pr_h = Inches(5.6)
ncols = 5
gap = Inches(0.12)
pr_w = (SW - Inches(1.0) - gap * (ncols - 1)) / ncols

for i, (name, sub, items, color) in enumerate(practices):
    x = Inches(0.5) + i * (pr_w + gap)
    round_rect(s, x, pr_y, pr_w, pr_h, WHITE, line=color, line_w=2, radius=0.04)
    rect(s, x, pr_y, pr_w, Inches(1.0), color)
    text_box(s, x + Inches(0.1), pr_y + Inches(0.05),
             pr_w - Inches(0.2), Inches(0.55),
             name, size=11, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.15), pr_y + Inches(0.65),
             pr_w - Inches(0.3), Inches(0.35),
             sub, size=9.5, italic=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_lines(s, x + Inches(0.2), pr_y + Inches(1.2),
              pr_w - Inches(0.4), pr_h - Inches(1.35),
              ["•  " + it for it in items], size=10, color=DARK,
              line_spacing=1.4)

footer_bar(s)


# ════════════════════════════════════════════════════════════════════════
# SLIDE — Step policies and Roles & Responsibilities per board level
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Step-level policies · what happens at each stage and who is responsible")

# Two big sections side by side: Portfolio Board · Team Kanban
sec_y = Inches(1.0)
sec_h = Inches(6.0)
sec_w = (SW - Inches(1.1)) / 2

# Section 1 — Portfolio Board
x_left = Inches(0.5)
round_rect(s, x_left, sec_y, sec_w, sec_h, WHITE, line=ACCENT, line_w=2, radius=0.04)
rect(s, x_left, sec_y, sec_w, Inches(0.55), ACCENT)
text_box(s, x_left, sec_y + Inches(0.05), sec_w, Inches(0.45),
         "PORTFOLIO BOARD  ·  cards = INITIATIVES",
         size=13, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

portfolio_steps = [
    ("Funnel",
     "Idea submitted with problem statement",
     "Owner: Sponsor",
     "Reviewer: Demand Manager"),
    ("Reviewing",
     "Business case refined · classification applied",
     "Owner: Sponsor + Demand Manager",
     "SLA: 2 weeks"),
    ("Analyzing",
     "Sized · scored · feasibility checked",
     "Owner: Demand Manager + Architect",
     "Output: Pain Score, T-shirt effort, Class of Service"),
    ("Approval Gate",
     "Go / No-Go decision",
     "Decider: Director (Carlos / Kumar)",
     "Cadence: monthly Portfolio Committee"),
    ("Backlog",
     "Approved · awaiting capacity",
     "Owner: Portfolio Manager",
     "Sequencing by Class of Service + score"),
    ("Implementing",
     "Assigned to Value Stream Team",
     "Owner: Team PM",
     "Tracked at Layer 0 VP Dashboard"),
    ("Done",
     "Value realized · post-implementation review",
     "Owner: Sponsor + PM",
     "Outcome metrics captured"),
]
step_y = sec_y + Inches(0.7)
step_h = (sec_h - Inches(0.85)) / len(portfolio_steps)
for i, (step, what, who, extra) in enumerate(portfolio_steps):
    y = step_y + i * step_h
    rect(s, x_left + Inches(0.2), y, Inches(0.08), step_h - Inches(0.05), ACCENT)
    text_box(s, x_left + Inches(0.4), y + Inches(0.02),
             sec_w - Inches(0.5), Inches(0.25),
             step, size=11, bold=True, color=ACCENT)
    text_box(s, x_left + Inches(0.4), y + Inches(0.22),
             sec_w - Inches(0.5), Inches(0.22),
             what, size=10, color=DARK)
    text_box(s, x_left + Inches(0.4), y + Inches(0.42),
             sec_w - Inches(0.5), Inches(0.22),
             who, size=9, italic=True, color=GREY)
    text_box(s, x_left + Inches(0.4), y + Inches(0.6),
             sec_w - Inches(0.5), Inches(0.22),
             extra, size=9, italic=True, color=GREY)

# Section 2 — Team Kanban
x_right = Inches(0.5) + sec_w + Inches(0.1)
round_rect(s, x_right, sec_y, sec_w, sec_h, WHITE, line=BLUE, line_w=2, radius=0.04)
rect(s, x_right, sec_y, sec_w, Inches(0.55), BLUE)
text_box(s, x_right, sec_y + Inches(0.05), sec_w, Inches(0.45),
         "TEAM KANBAN  ·  cards = EPICS · STORIES",
         size=13, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

team_phases = [
    ("Phase 0 — Discovery",
     "Problem framing · users · hypothesis",
     "Owner: Product Manager",
     "Always-on; supports Materialization"),
    ("Phase 1 — Materialization",
     "Definition · solutioning · sizing",
     "Owner: Product Manager + Tech Lead",
     "Output: Roadmap, Budget, Sizing"),
    ("Phase 2 — Creation",
     "Planning · team / vendor assignment",
     "Owner: Product Manager",
     "Output: Outcome definition, KPIs, Strategic alignment"),
    ("Phase 3 — Ongoing (Sprints)",
     "Build · test · deploy iteratively",
     "Owner: Squad (PM + Tech Lead + BA + dev/QA)",
     "Sprint cadence · daily / review / retro"),
    ("Phase 4 — Maintenance",
     "Hypercare · stabilization · knowledge",
     "Owner: Operations + Tech Lead",
     "Handoff to BAU within 30 days post-launch"),
    ("Discovery ↔ Delivery handoff",
     "Items move only when DoR is met",
     "Gatekeeper: Product Manager",
     "Discovery feeds Delivery continuously"),
]
step_y = sec_y + Inches(0.7)
step_h = (sec_h - Inches(0.85)) / len(team_phases)
for i, (step, what, who, extra) in enumerate(team_phases):
    y = step_y + i * step_h
    rect(s, x_right + Inches(0.2), y, Inches(0.08), step_h - Inches(0.05), BLUE)
    text_box(s, x_right + Inches(0.4), y + Inches(0.02),
             sec_w - Inches(0.5), Inches(0.25),
             step, size=11, bold=True, color=BLUE)
    text_box(s, x_right + Inches(0.4), y + Inches(0.25),
             sec_w - Inches(0.5), Inches(0.25),
             what, size=10, color=DARK)
    text_box(s, x_right + Inches(0.4), y + Inches(0.5),
             sec_w - Inches(0.5), Inches(0.25),
             who, size=9, italic=True, color=GREY)
    text_box(s, x_right + Inches(0.4), y + Inches(0.7),
             sec_w - Inches(0.5), Inches(0.22),
             extra, size=9, italic=True, color=GREY)

footer_bar(s)


# ════════════════════════════════════════════════════════════════════════
out_file = OUT / "DigitalProducts_TOM_Proposal_v2.0.pptx"
prs.save(str(out_file))
print(f"✓ Saved: {out_file}")
print(f"  Slides: {len(prs.slides)}")
