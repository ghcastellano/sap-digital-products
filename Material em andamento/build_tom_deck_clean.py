"""
DigitalProducts_TOM_Proposal_v1.0.pptx — CLIENT-FACING version.
No internal references (no KB, no source citations, no internal labels).
Focus: Stage 01 Demand Management complete + TOM context.
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

# Section colors (match Stage 01 Demand Management board)
GREEN       = RGBColor(0x4C, 0x8C, 0x3F)        # CAPABILITIES
GREEN_LIGHT = RGBColor(0xE8, 0xF1, 0xE3)
TEAL        = RGBColor(0x2B, 0x6B, 0x77)        # ROLES
TEAL_LIGHT  = RGBColor(0xDC, 0xEC, 0xEF)
AMBER       = RGBColor(0xE8, 0x8B, 0x2F)        # TOOLS
AMBER_LIGHT = RGBColor(0xFD, 0xEC, 0xD7)
INDIGO      = RGBColor(0x2C, 0x5A, 0x88)        # GOVERNANCE
INDIGO_LIGHT= RGBColor(0xDD, 0xE6, 0xF0)
RED         = RGBColor(0xB6, 0x3B, 0x1F)        # PERFORMANCE
RED_LIGHT   = RGBColor(0xF8, 0xDF, 0xD8)

WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
DARK        = RGBColor(0x2C, 0x3E, 0x50)
GREY        = RGBColor(0x6E, 0x78, 0x87)
LIGHT       = RGBColor(0xF5, 0xF5, 0xF7)

FONT = "Arial"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW = prs.slide_width
SH = prs.slide_height


# ── Helpers ──────────────────────────────────────────────────────────────
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
    rect(slide, 0, 0, SW, Inches(0.85), color)
    text_box(slide, Inches(0.4), Inches(0.10), SW - Inches(0.8),
             Inches(0.65), title, size=20, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    rect(slide, SW - Inches(0.18), 0, Inches(0.18), Inches(0.85), ACCENT)


def footer_bar(slide, text=""):
    """Discreet footer — only project + page indicator, no sources."""
    y = SH - Inches(0.35)
    text_box(slide, Inches(0.4), y, SW - Inches(0.8), Inches(0.3),
             "Avangrid · Digital Products · Target Operating Model",
             size=9, color=GREY, italic=True)


def section_card(slide, x, y, w, h, header_text, body_text,
                 header_color, body_color):
    """Standard section card: colored header band + light body."""
    round_rect(slide, x, y, w, h, body_color, line=header_color,
               line_w=1.5, radius=0.08)
    rect(slide, x + Inches(0.02), y + Inches(0.02),
         w - Inches(0.04), Inches(0.45), header_color)
    text_box(slide, x + Inches(0.1), y + Inches(0.03),
             w - Inches(0.2), Inches(0.42),
             header_text, size=11, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(slide, x + Inches(0.15), y + Inches(0.55),
             w - Inches(0.3), h - Inches(0.65),
             body_text, size=10.5, color=DARK, anchor=MSO_ANCHOR.TOP)


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
         "Demand Management — Stage 01",
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

# Governing thought
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

# 3 supporting columns
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
     ["Single intake channel per Value Stream",
      "Two-gate governance (weekly triage + monthly portfolio)",
      "Class-of-Service framework with capacity caps",
      "Unified backlog system with one source of truth"]),
    ("HOW WE GET THERE", GREEN,
     ["Pilot with one Value Stream over 14 weeks",
      "Refine and validate the governance model",
      "Scale to remaining Value Streams in waves",
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
# SLIDE 3 — Digital Products Organization
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Digital Products organization · approximately 50 people across two branches")

# VP box
vp_y = Inches(1.2)
round_rect(s, Inches(4.2), vp_y, Inches(5), Inches(0.8),
           NAVY_DARK, radius=0.2)
text_box(s, Inches(4.2), vp_y + Inches(0.05), Inches(5), Inches(0.7),
         "VP  ·  IT Business Solutions Delivery",
         size=15, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Two branches
branch_y = Inches(2.3)
branch_h = Inches(4.3)
b_w = Inches(6.2)
left_x = Inches(0.4)
right_x = Inches(6.8)

# Connectors
conn_y = vp_y + Inches(0.8)
shp = s.shapes.add_connector(1, Inches(6.7), conn_y, Inches(3.5), branch_y)
shp.line.color.rgb = BLUE
shp.line.width = Pt(2)
shp = s.shapes.add_connector(1, Inches(6.7), conn_y, Inches(9.9), branch_y)
shp.line.color.rgb = BLUE
shp.line.width = Pt(2)

# Carlos branch (IT Business Apps)
round_rect(s, left_x, branch_y, b_w, branch_h, WHITE, line=BLUE, line_w=2, radius=0.04)
rect(s, left_x, branch_y, b_w, Inches(0.6), BLUE)
text_box(s, left_x, branch_y + Inches(0.08), b_w, Inches(0.5),
         "IT BUSINESS APPLICATIONS",
         size=14, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text_box(s, left_x, branch_y + Inches(0.7), b_w, Inches(0.3),
         "approximately 26 people  ·  4 sub-teams",
         size=11, italic=True, color=GREY,
         align=PP_ALIGN.CENTER)

carlos = [
    "Electric & Gas Operations  —  ~7 people",
    "GIS & Operations Support  —  ~6 people",
    "Corporate Applications  —  ~8 people",
    "Global SAP (Iberdrola)  —  ~5 people",
]
for i, t in enumerate(carlos):
    y = branch_y + Inches(1.15) + i * Inches(0.7)
    round_rect(s, left_x + Inches(0.3), y, b_w - Inches(0.6), Inches(0.55),
               LIGHT, line=GREY, line_w=0.5, radius=0.1)
    text_box(s, left_x + Inches(0.4), y, b_w - Inches(0.7), Inches(0.55),
             t, size=12, color=DARK, anchor=MSO_ANCHOR.MIDDLE)

# Kumar branch (Customer Services Apps)
round_rect(s, right_x, branch_y, b_w, branch_h, WHITE, line=ACCENT, line_w=2, radius=0.04)
rect(s, right_x, branch_y, b_w, Inches(0.6), ACCENT)
text_box(s, right_x, branch_y + Inches(0.08), b_w, Inches(0.5),
         "CUSTOMER SERVICES APPLICATIONS",
         size=14, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text_box(s, right_x, branch_y + Inches(0.7), b_w, Inches(0.3),
         "approximately 24 people  ·  5 sub-teams",
         size=11, italic=True, color=GREY,
         align=PP_ALIGN.CENTER)

kumar = [
    "Billing, FICA & Collections  —  ~8 people",
    "Device & Metering  —  ~8 people",
    "CRM & Contact Center  —  ~4 people",
    "SAP & CS Architecture  —  ~2 people",
    "SAP Basis  —  ~2 people",
]
for i, t in enumerate(kumar):
    y = branch_y + Inches(1.15) + i * Inches(0.55)
    round_rect(s, right_x + Inches(0.3), y, b_w - Inches(0.6), Inches(0.45),
               LIGHT, line=GREY, line_w=0.5, radius=0.1)
    text_box(s, right_x + Inches(0.4), y, b_w - Inches(0.7), Inches(0.45),
             t, size=12, color=DARK, anchor=MSO_ANCHOR.MIDDLE)

footer_bar(s)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 4 — Target Operating Model overview
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Target Operating Model · eight Value Stream Teams aligned to customer outcomes")

streams = [
    ("Meter-to-Cash",
     "Service Lifecycle",
     "Customer signs up · activates · changes · disconnects",
     ACCENT),
    ("Meter-to-Cash",
     "Revenue Operations",
     "Accurate billing · easy payment · fair recovery",
     ACCENT),
    ("Customer Care",
     "& Contact Center",
     "Issue resolution · self-service · voice of customer",
     ACCENT),
    ("Field Operations",
     "",
     "Field worker / inspector / lineman productivity",
     BLUE),
    ("Engineering &",
     "Asset Management",
     "Engineer / planner designs with reliable data",
     BLUE),
    ("Network Operations",
     "& Control",
     "Grid operator prevents and minimizes outages",
     BLUE),
    ("Corporate &",
     "Regulatory Services",
     "HR / employee / regulatory user self-service",
     BLUE),
    ("Iberdrola Corporate",
     "Services",
     "Spain finance / procurement / budget",
     BLUE),
]

# 2 rows × 4 columns grid
grid_y = Inches(1.15)
grid_h = Inches(4.2)
ncols = 4
nrows = 2
gap = Inches(0.15)
col_w = (SW - Inches(1.0) - gap * (ncols - 1)) / ncols
row_h = (grid_h - gap) / nrows

for i, (n1, n2, outcome, color) in enumerate(streams):
    row = i // ncols
    col = i % ncols
    x = Inches(0.5) + col * (col_w + gap)
    y = grid_y + row * (row_h + gap)
    round_rect(s, x, y, col_w, row_h, WHITE, line=color, line_w=2, radius=0.04)
    rect(s, x, y, col_w, Inches(0.95), color)
    text_box(s, x, y + Inches(0.1), col_w, Inches(0.4),
             n1, size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x, y + Inches(0.45), col_w, Inches(0.4),
             n2, size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.15), y + Inches(1.05),
             col_w - Inches(0.3), row_h - Inches(1.15),
             outcome, size=11, color=DARK, italic=True,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Legend
leg_y = Inches(5.6)
rect(s, Inches(0.5), leg_y, Inches(0.3), Inches(0.3), ACCENT)
text_box(s, Inches(0.9), leg_y, Inches(4.5), Inches(0.3),
         "Customer Services Applications", size=11, color=DARK,
         anchor=MSO_ANCHOR.MIDDLE)
rect(s, Inches(5.5), leg_y, Inches(0.3), Inches(0.3), BLUE)
text_box(s, Inches(5.9), leg_y, Inches(4.5), Inches(0.3),
         "IT Business Applications", size=11, color=DARK,
         anchor=MSO_ANCHOR.MIDDLE)

# Bottom note
text_box(s, Inches(0.5), Inches(6.2), SW - Inches(1), Inches(0.5),
         "Each team operates a single backlog with end-to-end ownership of its outcome. "
         "Cross-team coordination happens through dedicated forums, not shared backlogs.",
         size=12, italic=True, color=GREY, align=PP_ALIGN.CENTER)

footer_bar(s)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 5 — Stage 01 Demand Management overview
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Stage 01 — Demand Management · five framework dimensions")

# 5 category cards in a row
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
def category_slide(category_name, count_label, color, light_color,
                   items, ncols, action_title):
    """items = list of (title, body) tuples."""
    s = blank()
    header_bar(s, action_title)

    # Section badge
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
# SLIDE 6 — CAPABILITIES (14 items in 5×3 grid)
# ════════════════════════════════════════════════════════════════════════
capabilities = [
    ("Demand intake channels",
     "Single front door per Value Stream Team. Email, Teams, ad-hoc messaging are decommissioned as demand entry. Incident-only path remains."),
    ("Demand registration & logging",
     "Auto-created backlog item from intake form. Immutable audit fields: requester, business unit, team, date. Every state change logged."),
    ("Channel consolidation",
     "One adaptive intake form per team. Target: 90% of demand via single channel by month 6. No alternative formal channels accepted."),
    ("Standard intake formats",
     "Canonical template with mandatory Definition of Ready fields: title, business value, acceptance criteria, class of service, due date, sponsor."),
    ("Demand classification",
     "Three independent axes: type, size, urgency. Closed taxonomy. Classification done at intake gate, not self-declared."),
    ("Run · Change · Regulatory · Innovation",
     "Four Classes of Service with capacity caps. Replaces ad-hoc 'urgent override' culture with explicit policy."),
    ("Business value definition",
     "Weighted scoring formula combining business value, user experience, safety, and effort. Sponsor signs off the score at intake."),
    ("Acceptance criteria",
     "Mandatory in Definition of Ready. Format: Given/When/Then or measurable outcome. Reviewed before scoring. Immutable post-prioritization."),
    ("Technical feasibility input",
     "Lightweight check at intake gate by rotating architect. Time-boxed (2 hours max). Output: Feasible / Conditional / Not feasible."),
    ("Triage & validation",
     "Weekly 30-minute triage committee. Outcomes: Accept · Send-back · Reject. Five-business-day SLA."),
    ("Demand routing & assignment",
     "Routing by Value Stream, not by team or person. Auto-routing based on requester selection. Cross-stream items go to Architecture review."),
    ("Urgent or exception requests",
     "Expedite Class of Service capped at 20% of throughput. Real-time incidents bypass intake (Service Desk SLA)."),
    ("Demand visibility & tracking",
     "Single dashboard, public read-only. Views: inflow funnel · backlog by team · class-of-service mix · aging · top priorities."),
    ("Duplicate & overlap identification",
     "Similarity search at submit (top-3 matches surfaced). Weekly de-dup query. Overlaps escalated to Portfolio Committee."),
]
category_slide("CAPABILITIES", "14 items", GREEN, GREEN_LIGHT,
               capabilities, ncols=5,
               action_title="Capabilities · what the organization must be able to do")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 7 — ROLES (4 items)
# ════════════════════════════════════════════════════════════════════════
roles = [
    ("Business requestors & stakeholders",
     "All originating business units route through a single intake form per Value Stream. "
     "Includes Customer Services Products organization, Networks Business sponsors, Iberdrola corporate, and BRS UIL/CMP/NY."),
    ("Product or Demand Managers",
     "One Product Manager per Value Stream Team — accountable for outcome, scoring, and backlog priority. "
     "One Demand Manager (cross-team) audits scoring and intake quality."),
    ("Architects & technical reviewers",
     "Architecture Council with rotating members from existing technical leads. "
     "Time-boxed advisory input at the intake gate. Provides feasibility view, not a veto."),
    ("Triage / intake committee",
     "Weekly committee per portfolio. Composition: Demand Manager (chair) plus Product Managers and Architect. "
     "Decisions logged with reason. Cross-portfolio escalations go to monthly Portfolio Committee."),
]
category_slide("ROLES", "4 items", TEAL, TEAL_LIGHT,
               roles, ncols=2,
               action_title="Roles · who owns and operates the framework")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 8 — TOOLS (5 items)
# ════════════════════════════════════════════════════════════════════════
tools = [
    ("Demand capture",
     "Single intake form per Value Stream Team with adaptive logic. "
     "Mandatory fields enforce Definition of Ready. Separate dedicated channel for real-time grid operations."),
    ("Tracking & backlog",
     "Single backlog system as system of record. Item hierarchy: Initiative → Epic → Feature → Story. "
     "Team Kanban with five phases and dual track (Discovery + Delivery)."),
    ("Communication channels",
     "Triage outcomes auto-published to requesters. Per-requester portal showing live status. "
     "Cross-team coordination through dedicated forums."),
    ("Documentation",
     "Single source of truth for Definition of Ready, Class of Service taxonomy, per-team product brief. "
     "Sponsor approval logged within the intake form itself."),
    ("Reporting & visibility",
     "Demand dashboard with five views: inflow funnel, backlog by team, class-of-service mix, aging, top priorities. "
     "Auto-refresh under one hour. Director-level rollups."),
]
category_slide("TOOLS", "5 items", AMBER, AMBER_LIGHT,
               tools, ncols=3,
               action_title="Tools · systems and platforms used")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 9 — GOVERNANCE (6 items)
# ════════════════════════════════════════════════════════════════════════
governance = [
    ("Demand review forums",
     "Two-gate structure. Gate 1: weekly triage (30 minutes). "
     "Gate 2: monthly Portfolio Committee (90 minutes). Quarterly cross-portfolio Big Room Planning."),
    ("Decision-making practices",
     "Four explicit outcomes: Accept · Send-back · Reject · Redirect. "
     "Every decision audit-logged with reason. No silent overrides."),
    ("Escalation mechanisms",
     "Cross-team conflicts resolved at Customer Journey Forum. "
     "Cross-portfolio conflicts resolved at VP Portfolio Sync. "
     "Every Expedite reviewed at the next Portfolio Committee."),
    ("Definition of intake guidelines",
     "Single source of truth for the Definition of Ready template, Class of Service taxonomy, and capacity caps. "
     "Reviewed every six months by the Demand Council."),
    ("Alignment with business stakeholders",
     "Quarterly Big Room Planning. Customer Journey Forum per major customer journey. "
     "Sponsor sign-off mandatory at intake (business value approval)."),
    ("Oversight of inflow & backlog health",
     "Demand Manager owns the demand dashboard. Aging beyond threshold triggers automatic escalation. "
     "Quarterly demand-vs-delivery conversion review."),
]
category_slide("GOVERNANCE", "6 items", INDIGO, INDIGO_LIGHT,
               governance, ncols=3,
               action_title="Governance · decisions, forums, and controls")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 10 — PERFORMANCE (7 items)
# ════════════════════════════════════════════════════════════════════════
performance = [
    ("Demand volume & inflow trends",
     "Submissions per week, month, quarter. Breakdown by Value Stream, Class of Service, OPCO, and source business area."),
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
category_slide("PERFORMANCE", "7 items", RED, RED_LIGHT,
               performance, ncols=4,
               action_title="Performance · how success is measured")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 11 — Portfolio Architecture (3 layers)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Portfolio architecture · three layers from VP rollup to team execution")

# Layer 0
l0_y = Inches(1.2)
round_rect(s, Inches(2.5), l0_y, Inches(8.3), Inches(1.0),
           BLUE_LIGHT, line=BLUE, line_w=2, radius=0.04)
text_box(s, Inches(2.7), l0_y + Inches(0.1), Inches(8), Inches(0.4),
         "VP ROLL-UP DASHBOARD",
         size=15, bold=True, color=BLUE)
text_box(s, Inches(2.7), l0_y + Inches(0.45), Inches(8), Inches(0.5),
         "Aggregated view across both portfolios · monthly Portfolio Sync · cross-portfolio dependencies",
         size=11, color=DARK, italic=True)

# Down arrow
shp = s.shapes.add_connector(1, Inches(6.65), l0_y + Inches(1.0),
                              Inches(6.65), Inches(2.5))
shp.line.color.rgb = NAVY
shp.line.width = Pt(2)

# Layer 1
l1_y = Inches(2.5)
# Two portfolio boards
pb_w = Inches(6)
pb_h = Inches(1.4)

round_rect(s, Inches(0.5), l1_y, pb_w, pb_h, WHITE, line=ACCENT, line_w=2, radius=0.04)
rect(s, Inches(0.5), l1_y, pb_w, Inches(0.5), ACCENT)
text_box(s, Inches(0.5), l1_y + Inches(0.05), pb_w, Inches(0.4),
         "PORTFOLIO BOARD  ·  Customer Services Applications",
         size=12, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text_box(s, Inches(0.7), l1_y + Inches(0.6), pb_w - Inches(0.4), Inches(0.7),
         "3 Value Streams · cards = Initiatives\nMonthly Portfolio Committee",
         size=11, color=DARK)

round_rect(s, Inches(6.85), l1_y, pb_w, pb_h, WHITE, line=BLUE, line_w=2, radius=0.04)
rect(s, Inches(6.85), l1_y, pb_w, Inches(0.5), BLUE)
text_box(s, Inches(6.85), l1_y + Inches(0.05), pb_w, Inches(0.4),
         "PORTFOLIO BOARD  ·  IT Business Applications",
         size=12, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text_box(s, Inches(7.05), l1_y + Inches(0.6), pb_w - Inches(0.4), Inches(0.7),
         "5 Value Streams · cards = Initiatives\nMonthly Portfolio Committee",
         size=11, color=DARK)

# Down arrow
shp = s.shapes.add_connector(1, Inches(6.65), l1_y + pb_h,
                              Inches(6.65), Inches(4.4))
shp.line.color.rgb = NAVY
shp.line.width = Pt(2)

# Layer 2 — 8 crew kanbans
l2_y = Inches(4.4)
text_box(s, Inches(0.5), l2_y - Inches(0.4), SW - Inches(1), Inches(0.3),
         "EIGHT VALUE STREAM TEAM BOARDS  ·  cards = Epics → Features → Stories",
         size=13, bold=True, color=NAVY,
         align=PP_ALIGN.CENTER)

teams = ["Meter-to-Cash\nService Lifecycle", "Meter-to-Cash\nRevenue Operations",
         "Customer Care\n& Contact Center", "Field Operations",
         "Engineering\n& Asset Mgmt", "Network Operations\n& Control",
         "Corporate &\nRegulatory", "Iberdrola\nCorporate Services"]
team_w = (SW - Inches(1)) / 8
for i, t in enumerate(teams):
    x = Inches(0.5) + i * team_w
    color = ACCENT if i < 3 else BLUE
    round_rect(s, x + Inches(0.05), l2_y, team_w - Inches(0.1),
               Inches(2.0), WHITE, line=color, line_w=1.5, radius=0.04)
    rect(s, x + Inches(0.05), l2_y, team_w - Inches(0.1), Inches(0.5), color)
    lines = t.split("\n")
    text_box(s, x + Inches(0.05), l2_y + Inches(0.03),
             team_w - Inches(0.1), Inches(0.5),
             lines[0], size=9, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    if len(lines) > 1:
        text_box(s, x + Inches(0.05), l2_y + Inches(0.25),
                 team_w - Inches(0.1), Inches(0.3),
                 lines[1], size=9, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER)
    text_box(s, x + Inches(0.1), l2_y + Inches(0.7),
             team_w - Inches(0.2), Inches(1.2),
             "5 phases\n× 2 tracks\n(Discovery + Delivery)",
             size=10, color=DARK, italic=True,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

footer_bar(s)


# ════════════════════════════════════════════════════════════════════════
# SLIDE — Visual mockup · Portfolio Board (initiatives flowing)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Portfolio Board · cards are Initiatives flowing through stages")

# Sub-title
text_box(s, Inches(0.5), Inches(1.0), SW - Inches(1), Inches(0.4),
         "Example  ·  IT Business Applications Portfolio Board",
         size=13, italic=True, color=GREY, align=PP_ALIGN.CENTER)

# Stages
stage_y = Inches(1.6)
stage_h = Inches(5.0)
stages = ["Funnel", "Reviewing", "Analyzing", "Backlog", "Implementing", "Done"]
ncols = len(stages)
gap = Inches(0.06)
col_w = (SW - Inches(1.0) - gap * (ncols - 1)) / ncols

# Stage headers
for i, stg in enumerate(stages):
    x = Inches(0.5) + i * (col_w + gap)
    rect(s, x, stage_y, col_w, Inches(0.5), NAVY_DARK)
    text_box(s, x, stage_y + Inches(0.05), col_w, Inches(0.4),
             stg, size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, x, stage_y + Inches(0.5), col_w, stage_h - Inches(0.5),
         LIGHT, line=GREY, line_w=0.5)

# Mock cards in each column
def init_card(slide, x, y, w, h, title, owner, cos_color):
    round_rect(slide, x, y, w, h, WHITE, line=GREY, line_w=0.5, radius=0.06)
    rect(slide, x, y, Inches(0.08), h, cos_color)
    text_box(slide, x + Inches(0.15), y + Inches(0.05),
             w - Inches(0.2), Inches(0.4),
             title, size=10, bold=True, color=DARK)
    text_box(slide, x + Inches(0.15), y + Inches(0.4),
             w - Inches(0.2), Inches(0.3),
             owner, size=9, color=GREY, italic=True)

# Funnel
init_card(s, Inches(0.5) + Inches(0.05), stage_y + Inches(0.6),
          col_w - Inches(0.1), Inches(0.7),
          "Storm program v2", "Field Operations", AMBER)
init_card(s, Inches(0.5) + Inches(0.05), stage_y + Inches(1.4),
          col_w - Inches(0.1), Inches(0.7),
          "Mobility upgrades", "Field Operations", GREEN)

# Reviewing
init_card(s, Inches(0.5) + col_w + gap + Inches(0.05), stage_y + Inches(0.6),
          col_w - Inches(0.1), Inches(0.7),
          "Drawing platform", "Engineering & Asset", GREEN)

# Analyzing
init_card(s, Inches(0.5) + 2*(col_w + gap) + Inches(0.05), stage_y + Inches(0.6),
          col_w - Inches(0.1), Inches(0.7),
          "Network deployment", "Engineering & Asset", GREEN)
init_card(s, Inches(0.5) + 2*(col_w + gap) + Inches(0.05), stage_y + Inches(1.4),
          col_w - Inches(0.1), Inches(0.7),
          "OMS modernization", "Network Ops", RED)

# Backlog
init_card(s, Inches(0.5) + 3*(col_w + gap) + Inches(0.05), stage_y + Inches(0.6),
          col_w - Inches(0.1), Inches(0.7),
          "Doc viewer scale", "Field Operations", GREEN)

# Implementing — multiple
init_card(s, Inches(0.5) + 4*(col_w + gap) + Inches(0.05), stage_y + Inches(0.6),
          col_w - Inches(0.1), Inches(0.7),
          "Mobile field migration", "Field Operations", GREEN)
init_card(s, Inches(0.5) + 4*(col_w + gap) + Inches(0.05), stage_y + Inches(1.4),
          col_w - Inches(0.1), Inches(0.7),
          "ERP transition", "Iberdrola Corp", BLUE)
init_card(s, Inches(0.5) + 4*(col_w + gap) + Inches(0.05), stage_y + Inches(2.2),
          col_w - Inches(0.1), Inches(0.7),
          "Asset registry refresh", "Engineering & Asset", GREEN)

# Done
init_card(s, Inches(0.5) + 5*(col_w + gap) + Inches(0.05), stage_y + Inches(0.6),
          col_w - Inches(0.1), Inches(0.7),
          "Storm program Q1", "Field Operations", GREEN)

# Shark Tank gate marker
gate_x = Inches(0.5) + 3*(col_w + gap) - Inches(0.05)
gate_y = stage_y - Inches(0.05)
rect(s, gate_x, gate_y, Inches(0.05), stage_h + Inches(0.1), RED)
text_box(s, gate_x - Inches(1.5), gate_y + Inches(0.1),
         Inches(3), Inches(0.3),
         "▲ APPROVAL GATE", size=10, bold=True, color=RED,
         align=PP_ALIGN.CENTER)

# Legend
leg_y = Inches(6.7)
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
# SLIDE — Visual mockup · Crew Kanban (5 phases × 2 tracks)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Team Kanban · five phases by two tracks · cards are Epics, Features, Stories")

# Sub-title
text_box(s, Inches(0.5), Inches(1.0), SW - Inches(1), Inches(0.4),
         "Example  ·  Field Operations Team Kanban",
         size=13, italic=True, color=GREY, align=PP_ALIGN.CENTER)

# Phases
ph_y = Inches(1.6)
ph_h = Inches(5.0)
phases_lbl = ["Phase 0\nDISCOVERY", "Phase 1\nMATERIALIZATION",
              "Phase 2\nCREATION", "Phase 3\nONGOING", "Phase 4\nMAINTENANCE"]
nph = len(phases_lbl)
gap = Inches(0.06)
left_panel = Inches(1.4)  # for track labels
ph_w = (SW - Inches(1.0) - left_panel - gap * (nph - 1)) / nph

# Phase headers
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

# Track labels (left side)
mid_y = ph_y + Inches(0.65) + (ph_h - Inches(0.65)) / 2
# Discovery (top)
rect(s, Inches(0.5), ph_y + Inches(0.65), left_panel, mid_y - ph_y - Inches(0.65), BLUE)
text_box(s, Inches(0.5), ph_y + Inches(0.65), left_panel, mid_y - ph_y - Inches(0.65),
         "DISCOVERY\nTRACK", size=12, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Delivery (bottom)
rect(s, Inches(0.5), mid_y, left_panel, ph_y + ph_h - mid_y, AMBER)
text_box(s, Inches(0.5), mid_y, left_panel, ph_y + ph_h - mid_y,
         "DELIVERY\nTRACK", size=12, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Track separator line
shp = s.shapes.add_connector(1, Inches(0.5) + left_panel, mid_y,
                              SW - Inches(0.5), mid_y)
shp.line.color.rgb = DARK
shp.line.width = Pt(2)

# Mock cards
def epic_card(slide, x, y, w, h, type_label, title, owner, color):
    round_rect(slide, x, y, w, h, WHITE, line=color, line_w=1.2, radius=0.08)
    rect(slide, x, y, w, Inches(0.22), color)
    text_box(slide, x + Inches(0.08), y + Inches(0.01),
             w - Inches(0.16), Inches(0.2),
             type_label, size=8, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    text_box(slide, x + Inches(0.1), y + Inches(0.27),
             w - Inches(0.2), Inches(0.45),
             title, size=10, bold=True, color=DARK)
    text_box(slide, x + Inches(0.1), y + h - Inches(0.25),
             w - Inches(0.2), Inches(0.2),
             owner, size=8, color=GREY, italic=True)

# Discovery track cards
def x_at(idx):
    return Inches(0.5) + left_panel + idx * (ph_w + gap)

# Phase 0 (Discovery)
epic_card(s, x_at(0) + Inches(0.05), ph_y + Inches(0.8),
          ph_w - Inches(0.1), Inches(0.85),
          "EPIC", "Migrate DLI inspections", "exploring offline reqs", BLUE)

# Phase 1 (Materialization)
epic_card(s, x_at(1) + Inches(0.05), ph_y + Inches(0.8),
          ph_w - Inches(0.1), Inches(0.85),
          "EPIC", "Migrate TLI inspections", "sizing with vendor", BLUE)

# Phase 2 (Creation, delivery)
epic_card(s, x_at(2) + Inches(0.05), mid_y + Inches(0.15),
          ph_w - Inches(0.1), Inches(0.85),
          "EPIC", "Migrate DGI/TGI", "team assigned · sprint plan", BLUE)

# Phase 3 (Ongoing — Features)
epic_card(s, x_at(3) + Inches(0.05), mid_y + Inches(0.15),
          ph_w - Inches(0.1), Inches(0.7),
          "FEATURE", "Offline DLI capability", "in dev · Sprint 5", AMBER)
epic_card(s, x_at(3) + Inches(0.05), mid_y + Inches(0.95),
          ph_w - Inches(0.1), Inches(0.7),
          "FEATURE", "Mobile TLI UI", "in dev · Sprint 5", AMBER)
epic_card(s, x_at(3) + Inches(0.05), mid_y + Inches(1.75),
          ph_w - Inches(0.1), Inches(0.5),
          "STORY", "Sync logic — QA", "Sprint 4", GREEN)

# Phase 4 (Maintenance)
epic_card(s, x_at(4) + Inches(0.05), mid_y + Inches(0.15),
          ph_w - Inches(0.1), Inches(0.7),
          "EPIC", "Decommission FWM", "hypercare · 30d", BLUE)

# Legend at bottom
leg_y = Inches(6.7)
items = [(BLUE, "EPIC  ·  multi-sprint"),
         (AMBER, "FEATURE  ·  2-6 weeks"),
         (GREEN, "STORY  ·  1 sprint")]
for i, (color, label) in enumerate(items):
    x = Inches(0.5) + i * Inches(4.2)
    rect(s, x, leg_y + Inches(0.05), Inches(0.2), Inches(0.2), color)
    text_box(s, x + Inches(0.3), leg_y, Inches(4), Inches(0.3),
             label, size=11, color=DARK)

footer_bar(s)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 12 — Class of Service framework
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

# Bottom callout
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
# SLIDE 13 — Roadmap
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Roadmap · 14-week pilot validates the model before scaling")

phases_y = Inches(1.4)
phases_h = Inches(2.3)
phases = [
    ("Weeks 1–2", "REVIEW & BASELINE", BLUE,
     "Stakeholder interviews · channel-volume baseline · vendor model decision"),
    ("Weeks 3–4", "TEAM DESIGN", BLUE,
     "Charter the pilot team · Product Manager onboarded · core team staffed"),
    ("Weeks 5–6", "SETUP", GREEN,
     "Portfolio Board + Team Kanban configured · DoR template · Class of Service"),
    ("Weeks 7–13", "PILOT EXECUTION", AMBER,
     "Team runs the new model · weekly triage · monthly portfolio review"),
    ("Week 14", "RETROSPECTIVE & SCALE", RED,
     "Validate KPIs · refine playbook · approval to scale to next team"),
]

ph_w = (SW - Inches(1.0)) / 5
gap = Inches(0.1)
for i, (when, what, color, desc) in enumerate(phases):
    x = Inches(0.5) + i * (ph_w)
    round_rect(s, x + gap/2, phases_y, ph_w - gap, phases_h,
               WHITE, line=color, line_w=2, radius=0.05)
    rect(s, x + gap/2, phases_y, ph_w - gap, Inches(0.5), color)
    text_box(s, x + gap/2, phases_y + Inches(0.05),
             ph_w - gap, Inches(0.4),
             when, size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.1), phases_y + Inches(0.6),
             ph_w - Inches(0.2), Inches(0.5),
             what, size=12, bold=True, color=DARK,
             align=PP_ALIGN.CENTER)
    text_box(s, x + Inches(0.15), phases_y + Inches(1.15),
             ph_w - Inches(0.3), phases_h - Inches(1.25),
             desc, size=10.5, color=GREY, italic=True,
             align=PP_ALIGN.CENTER)

    # Arrow to next
    if i < len(phases) - 1:
        text_box(s, x + ph_w - Inches(0.1), phases_y + phases_h/2 - Inches(0.15),
                 Inches(0.2), Inches(0.3), "→",
                 size=18, bold=True, color=GREY,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Bottom: pilot recommendation
pr_y = Inches(4.2)
round_rect(s, Inches(0.5), pr_y, SW - Inches(1), Inches(2.5),
           BLUE_LIGHT, line=BLUE, line_w=2, radius=0.04)
text_box(s, Inches(0.7), pr_y + Inches(0.15),
         SW - Inches(1.4), Inches(0.4),
         "PILOT TEAM RECOMMENDATION",
         size=13, bold=True, color=BLUE)
text_box(s, Inches(0.7), pr_y + Inches(0.55),
         SW - Inches(1.4), Inches(0.6),
         "Meter-to-Cash · Service Lifecycle (Smart Metering)",
         size=24, bold=True, color=NAVY_DARK)
add_lines(s, Inches(0.7), pr_y + Inches(1.35),
          SW - Inches(1.4), Inches(1.0),
          ["Closest existing analog to a Product Manager already in place — minimal change-management lift.",
           "Vendor delivery model already stable — Avangrid acts as product owner today.",
           "Clear, scoped user base (field services, meter installers, customer)  ·  contained risk."],
          size=12, color=DARK, line_spacing=1.4)

footer_bar(s)


# ════════════════════════════════════════════════════════════════════════
# Save
# ════════════════════════════════════════════════════════════════════════
out_file = OUT / "DigitalProducts_TOM_Proposal_v1.0.pptx"
prs.save(str(out_file))
print(f"✓ Saved: {out_file}")
print(f"  Slides: {len(prs.slides)}")
