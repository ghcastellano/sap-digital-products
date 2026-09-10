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

# ── Brand palette ── (aligned to 2604 AS-IS final deliverable tokens)
NAVY        = RGBColor(0x0F, 0x1C, 0x50)  # primary header bar / titles
NAVY_DARK   = RGBColor(0x09, 0x12, 0x35)
INDIGO      = RGBColor(0x1F, 0x3A, 0x68)  # secondary navy
INDIGO_LIGHT= RGBColor(0xDD, 0xE6, 0xF0)
BLUE        = RGBColor(0x2C, 0x57, 0x8B)  # mid blue boxes
BLUE_LIGHT  = RGBColor(0x67, 0x85, 0xC1)  # periwinkle accent
BLUE_PALE   = RGBColor(0xDC, 0xE6, 0xF5)
CYAN        = RGBColor(0x0B, 0x76, 0xA0)  # accent cyan
ACCENT      = RGBColor(0xF9, 0xA8, 0x25)  # gold
GREEN       = RGBColor(0x2E, 0x7D, 0x32)
GREEN_LIGHT = RGBColor(0xE8, 0xF1, 0xE3)
TEAL        = RGBColor(0x2B, 0x6B, 0x77)
TEAL_LIGHT  = RGBColor(0xDC, 0xEC, 0xEF)
AMBER       = RGBColor(0xE8, 0x8B, 0x2F)
AMBER_LIGHT = RGBColor(0xFD, 0xEC, 0xD7)
RED         = RGBColor(0xB6, 0x3B, 0x1F)
RED_LIGHT   = RGBColor(0xF8, 0xDF, 0xD8)
PURPLE      = RGBColor(0x7A, 0x66, 0xB8)
PURPLE_LIGHT= RGBColor(0xE5, 0xDF, 0xF0)

WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
DARK        = RGBColor(0x2C, 0x3E, 0x50)  # body text
GREY        = RGBColor(0x95, 0xA5, 0xA6)  # muted text
GREY_DARK   = RGBColor(0x6E, 0x78, 0x87)
LIGHT       = RGBColor(0xF5, 0xF5, 0xF5)  # background panels
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


def header_bar(slide, title, color=BLUE):
    """Top header bar — ALL CAPS section + ALL CAPS subtitle in white (mirrors 2604 final deliverable)."""
    rect(slide, 0, 0, SW, Inches(0.95), color)
    # Split on " · " to separate section name (kicker) from subtitle
    if " · " in title:
        kicker, sub = title.split(" · ", 1)
    else:
        kicker, sub = title, ""
    # Section name (top, larger)
    text_box(slide, Inches(0.4), Inches(0.08), SW - Inches(0.8),
             Inches(0.42), kicker.upper(),
             size=20, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    # Subtitle (bottom, smaller, ALL CAPS, WHITE — not gold)
    if sub:
        text_box(slide, Inches(0.4), Inches(0.52), SW - Inches(0.8),
                 Inches(0.35), sub.upper(),
                 size=11, bold=True, color=WHITE,
                 anchor=MSO_ANCHOR.MIDDLE)
    # NTT DATA brand mark top-right of header
    text_box(slide, SW - Inches(2.3), Inches(0.20),
             Inches(2.0), Inches(0.55),
             "NTT  DATA", size=12, bold=True, color=WHITE,
             align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


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
# Helper · Table of Contents (re-used as section divider with active highlight)
# ════════════════════════════════════════════════════════════════════════
toc_entries = [
    ("01", "Context",
     "Why Avangrid is consolidating Customer Services and IT Business Apps into Digital Products."),
    ("02", "Executive Summary",
     "The proposal in one page — scope, structural issues, governing principles, decisions ahead."),
    ("03", "Current State Analysis",
     "Vitor Cruz's three director branches, ~50 FTE, sub-team mapping and existing champions."),
    ("04", "Structural Issues & Guiding Principles",
     "Seven structural issues from the AS-IS, four guiding principles that resolve them."),
    ("05", "Stage 1 · Demand Management Framework",
     "Five dimensions — Capabilities · Roles · Tools · Governance · Performance — including branch specifics."),
    ("06", "Operating Model · Crews & Bases",
     "unFIX primer · Strongly Aligned Base · Supporting Crews · Sibling Bases · 8 Value Stream deep-dives · Vendor model."),
    ("07", "Demand Flow & Portfolio Operation",
     "Forums · Portfolio + Team Boards · Class of Service · Wave 1 Roadmap · TO-BE Demand Flow."),
    ("08", "Open Decisions & Next Steps",
     "Validation needed before steady-state — vendor model · PM staffing · NOC owner."),
]


def toc_slide(active_idx=None):
    """Render Table of Contents. If active_idx is given, highlight that row."""
    s = blank()
    text_box(s, Inches(0.6), Inches(0.45), Inches(7.5), Inches(0.7),
             "Table of contents", size=34, bold=True, color=CYAN)
    arrow_x = Inches(4.7); arrow_y = Inches(0.55)
    arrow_circle = s.shapes.add_shape(MSO_SHAPE.OVAL,
                                      arrow_x, arrow_y, Inches(0.55), Inches(0.55))
    arrow_circle.fill.solid(); arrow_circle.fill.fore_color.rgb = CYAN
    arrow_circle.line.fill.background()
    text_box(s, arrow_x, arrow_y, Inches(0.55), Inches(0.55),
             "▶", size=20, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, SW - Inches(2.5), Inches(0.45), Inches(2.0), Inches(0.5),
             "NTT  DATA", size=14, bold=True, color=NAVY,
             align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

    toc_y = Inches(1.5)
    toc_h = Inches(5.4)
    col_w = (SW - Inches(1.4)) / 2
    col_gap = Inches(0.3)
    row_h = toc_h / 4

    for i, (num, title, desc) in enumerate(toc_entries):
        col = i // 4
        row = i % 4
        x = Inches(0.7) + col * (col_w + col_gap)
        y = toc_y + row * row_h
        # Active highlight — light-blue rounded panel behind the row
        if active_idx == i:
            round_rect(s, x - Inches(0.1), y - Inches(0.05),
                       col_w + Inches(0.05), row_h - Inches(0.1),
                       BLUE_PALE, line=BLUE_LIGHT, line_w=1, radius=0.05)
        num_color = NAVY if active_idx is None or active_idx == i else GREY
        title_color = NAVY if active_idx is None or active_idx == i else GREY
        desc_color = DARK if active_idx is None or active_idx == i else GREY
        # Big number
        text_box(s, x, y, Inches(0.95), Inches(0.85),
                 num, size=44, bold=True, color=num_color,
                 anchor=MSO_ANCHOR.TOP)
        # Title
        text_box(s, x + Inches(1.0), y + Inches(0.05),
                 col_w - Inches(1.0), Inches(0.4),
                 title, size=18, bold=True, color=title_color)
        # Description
        text_box(s, x + Inches(1.0), y + Inches(0.5),
                 col_w - Inches(1.0), Inches(0.7),
                 desc, size=10.5, color=desc_color)
    return s


# Main TOC (no highlight)
toc_slide()



# ─── Section divider · 02 Executive Summary ───
toc_slide(active_idx=1)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 3 — Executive Summary (comprehensive · pre-content overview)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Executive Summary")

# Top — governing thought
gt_y = Inches(1.2)
round_rect(s, Inches(0.5), gt_y, SW - Inches(1), Inches(1.1),
           NAVY, line=NAVY, line_w=0, radius=0.05)
text_box(s, Inches(0.7), gt_y + Inches(0.1),
         SW - Inches(1.4), Inches(0.3),
         "GOVERNING THOUGHT",
         size=10, bold=True, color=ACCENT)
text_box(s, Inches(0.7), gt_y + Inches(0.4),
         SW - Inches(1.4), Inches(0.7),
         "Convert Avangrid's fragmented intake into a product-centric Target Operating Model — "
         "eight Value Stream Crews, one front door, one prioritization framework, two governance gates.",
         size=15, italic=True, color=WHITE,
         anchor=MSO_ANCHOR.MIDDLE)

# Three-column body — comprehensive proposal
col_y = Inches(2.6)
col_h = Inches(4.4)
col_w = (SW - Inches(1.2)) / 3
gap = Inches(0.15)

cols = [
    ("WHY NOW", BLUE,
     ["Demand fragmented across 6+ disconnected channels",
      "No shared definition of value across initiatives",
      "Portfolio visibility limited to Smart Metering (PPM Now)",
      "PJMs absorb demand reactively — no Product Manager role exists",
      "Forecast 30–50 % growth in Digital Products demand"]),
    ("WHAT WE PROPOSE", ACCENT,
     ["8 Value Stream Crews aligned to customer outcomes",
      "1 Jira request type per VS · mandatory DoR fields",
      "2-gate governance — weekly triage + monthly portfolio",
      "Class of Service framework — Expedite capped ≤ 20 %",
      "Two-level work model — Initiative at portfolio · Story at squad"]),
    ("HOW WE GET THERE", GREEN,
     ["Wave 1 — Demand Mgmt Design & Pilot (14 weeks)",
      "Pilot one Value Stream end to end · learn · adjust",
      "Five-step lifecycle — Enablement → Define → Design → Pilot → Learn",
      "Validated playbook + refined Jira config v2.0 at the end of Wave 1",
      "Continuous feedback log captured during pilot execution"]),
]
for i, (label, color, bullets) in enumerate(cols):
    x = Inches(0.5) + i * (col_w + gap)
    round_rect(s, x, col_y, col_w, col_h, WHITE, line=color, line_w=2, radius=0.04)
    rect(s, x, col_y, col_w, Inches(0.55), color)
    text_box(s, x, col_y + Inches(0.05), col_w, Inches(0.45),
             label, size=13, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_lines(s, x + Inches(0.2), col_y + Inches(0.7),
              col_w - Inches(0.4), col_h - Inches(0.85),
              ["•  " + b for b in bullets], size=11, color=DARK,
              line_spacing=1.35)

footer_bar(s)



# ─── Section divider · 03 Current State Analysis ───
toc_slide(active_idx=2)


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

# Header height for branch columns
hdr_h = Inches(0.95)

# Column 1 — Carlos
x1 = Inches(0.5)
round_rect(s, x1, col_y, col_w, hdr_h, NAVY, radius=0.04)
text_box(s, x1 + Inches(0.1), col_y + Inches(0.08), col_w - Inches(0.2), Inches(0.35),
         "Carlos Garbin — Director", size=11, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text_box(s, x1 + Inches(0.1), col_y + Inches(0.5), col_w - Inches(0.2), Inches(0.4),
         "IT Business Applications  ·  ~26 FTE",
         size=10, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

carlos_subs = [
    ("Bruno Moreto", "Electric & Gas Ops", "7 FTE", "G&E Champion: Andrew McGibbon", True),
    ("George Porto", "GIS & OSG", "6 FTE", "GIS Champion: George Porto", True),
    ("Sarath Garikapati", "Corporate Applications", "8 FTE", "Corporate Champion: Amit Patel", True),
    ("Vanessa Colombo", "Global SAP · Iberdrola FI · PO · Budget", "5 FTE",
     "Global SAP Champion: Michael Dubois", True),
]
for i, (name, team, fte, champ, has_champ) in enumerate(carlos_subs):
    y = col_y + hdr_h + Inches(0.1) + i * Inches(0.78)
    round_rect(s, x1, y, col_w, Inches(0.72), LIGHT, line=GREY, line_w=0.5, radius=0.05)
    if has_champ:
        rect(s, x1, y, Inches(0.06), Inches(0.72), GREEN)
    text_box(s, x1 + Inches(0.15), y + Inches(0.04),
             col_w - Inches(0.7), Inches(0.24),
             name, size=10, bold=True, color=DARK)
    text_box(s, x1 + Inches(0.15), y + Inches(0.26),
             col_w - Inches(0.7), Inches(0.24),
             team, size=8.5, color=DARK)
    rect(s, x1 + col_w - Inches(0.55), y + Inches(0.06),
         Inches(0.5), Inches(0.28), BLUE)
    text_box(s, x1 + col_w - Inches(0.55), y + Inches(0.06),
             Inches(0.5), Inches(0.28), fte,
             size=9, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x1 + Inches(0.15), y + Inches(0.5),
             col_w - Inches(0.2), Inches(0.22),
             champ, size=7.5, italic=True, color=GREY)

# Column 2 — Kumar
x2 = Inches(0.5) + col_w + gap
round_rect(s, x2, col_y, col_w, hdr_h, NAVY, radius=0.04)
text_box(s, x2 + Inches(0.1), col_y + Inches(0.08),
         col_w - Inches(0.2), Inches(0.35),
         "Kumar Jaladanki — Director", size=11, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text_box(s, x2 + Inches(0.1), col_y + Inches(0.5),
         col_w - Inches(0.2), Inches(0.4),
         "Customer Services Apps  ·  ~24 FTE",
         size=10, color=WHITE,
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
    y = col_y + hdr_h + Inches(0.1) + i * Inches(0.62)
    is_vacant = "vacant" in name.lower()
    bg = RED_LIGHT if is_vacant else LIGHT
    round_rect(s, x2, y, col_w, Inches(0.56), bg, line=GREY, line_w=0.5, radius=0.05)
    if has_champ:
        rect(s, x2, y, Inches(0.06), Inches(0.56), GREEN)
    name_color = VACANT if is_vacant else DARK
    text_box(s, x2 + Inches(0.15), y + Inches(0.03),
             col_w - Inches(0.7), Inches(0.22),
             name, size=10, bold=True, color=name_color)
    text_box(s, x2 + Inches(0.15), y + Inches(0.22),
             col_w - Inches(0.7), Inches(0.2),
             team, size=8.5, color=DARK)
    rect(s, x2 + col_w - Inches(0.55), y + Inches(0.05),
         Inches(0.5), Inches(0.26), ACCENT)
    text_box(s, x2 + col_w - Inches(0.55), y + Inches(0.05),
             Inches(0.5), Inches(0.26), fte,
             size=9, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    if champ:
        text_box(s, x2 + Inches(0.15), y + Inches(0.38),
                 col_w - Inches(0.2), Inches(0.18),
                 champ, size=7.5, italic=True, color=GREY)

# Column 3 — Heather
x3 = Inches(0.5) + 2*(col_w + gap)
round_rect(s, x3, col_y, col_w, hdr_h, NAVY, radius=0.04)
text_box(s, x3 + Inches(0.1), col_y + Inches(0.08),
         col_w - Inches(0.2), Inches(0.35),
         "Heather LaJoie — Director", size=11, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text_box(s, x3 + Inches(0.1), col_y + Inches(0.5),
         col_w - Inches(0.2), Inches(0.4),
         "Product Management  ·  ~2 FTE",
         size=10, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

heath_y = col_y + hdr_h + Inches(0.1)
round_rect(s, x3, heath_y, col_w, Inches(0.7), RED_LIGHT, line=GREY,
           line_w=0.5, radius=0.05)
text_box(s, x3 + Inches(0.15), heath_y + Inches(0.05),
         col_w - Inches(0.7), Inches(0.3),
         "Vacant", size=10, bold=True, color=VACANT)
text_box(s, x3 + Inches(0.15), heath_y + Inches(0.3),
         col_w - Inches(0.7), Inches(0.35),
         "Product Management", size=9, color=DARK)
rect(s, x3 + col_w - Inches(0.55), heath_y + Inches(0.08),
     Inches(0.5), Inches(0.28), PURPLE)
text_box(s, x3 + col_w - Inches(0.55), heath_y + Inches(0.08),
         Inches(0.5), Inches(0.28), "2 FTE",
         size=9, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Description text below
desc_y = heath_y + Inches(0.85)
desc_h = col_h - Inches(2.0)
round_rect(s, x3, desc_y, col_w, desc_h, BLUE_LIGHT, line=BLUE, line_w=1, radius=0.04)
text_box(s, x3 + Inches(0.15), desc_y + Inches(0.1),
         col_w - Inches(0.3), desc_h - Inches(0.2),
         "Strategic coordination layer across delivery and execution teams. "
         "Ensures alignment between business priorities and technical "
         "implementation.\n\n"
         "Shapes and supports prioritization of work · provides direction "
         "throughout execution · enables consistent focus on value creation.",
         size=9, color=DARK, italic=True)

# Column 4 — IT Program Delivery
x4 = Inches(0.5) + 3*(col_w + gap)
round_rect(s, x4, col_y, col_w, hdr_h, NAVY, radius=0.04)
text_box(s, x4 + Inches(0.1), col_y + Inches(0.08),
         col_w - Inches(0.2), Inches(0.35),
         "IT Program Delivery", size=11, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text_box(s, x4 + Inches(0.1), col_y + Inches(0.5),
         col_w - Inches(0.2), Inches(0.4),
         "(Vacant)",
         size=10, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
# Vacant card body
itd_y = col_y + hdr_h + Inches(0.1)
round_rect(s, x4, itd_y, col_w, Inches(0.7), RED_LIGHT, line=GREY,
           line_w=0.5, radius=0.05)
text_box(s, x4 + Inches(0.15), itd_y + Inches(0.05),
         col_w - Inches(0.3), Inches(0.6),
         "Vacant\n\nProgram Delivery role to be staffed",
         size=10, bold=True, color=VACANT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Champion legend at bottom
leg_y = col_y + col_h + Inches(0.1)
rect(s, Inches(0.5), leg_y, Inches(0.25), Inches(0.25), GREEN)
text_box(s, Inches(0.85), leg_y, Inches(4), Inches(0.25),
         "Area champion identified (6)",
         size=10, bold=True, italic=True, color=DARK)



# ─── Section divider · 04 Structural Issues & Guiding Principles ───
toc_slide(active_idx=3)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 4a — Seven Structural Issues (mirrors final AS-IS deliverable)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Seven structural issues surface — no single-area fix resolves them; the TOM must standardize end-to-end")

issues = [
    ("01", "No Product Entry Point",
     "Demand enters through multiple disconnected channels, with no single front door aligned to products or value streams.",
     "(6+ disconnected channels — Jira · IT Now · RITM · SharePoint · Email · Word. No single front door.)",
     "Product model requires a unified intake layer structured around products, not tools or channels."),
    ("02", "No Value-Based Decisioning",
     "Prioritization is inconsistent and driven by local criteria, with no common definition of value across initiatives.",
     "(Each area uses different criteria — PDD · impact × users · criticality · operational urgency · common sense.)",
     "Product model requires a standardized, value-based prioritization framework tied to business outcomes."),
    ("03", "No End-to-End Product View",
     "Portfolio management capabilities are fragmented, limiting visibility across demand, capacity, and delivery.",
     "(No consistent and standardized use of portfolio tool. Only Smart Metering uses PPM Now.)",
     "Product model requires integrated portfolio tooling to manage products holistically across their lifecycle."),
    ("04", "No Ownership of Outcomes",
     "No standard process to manage delivery. Ownership is unclear across the lifecycle, with teams and vendors defining timelines and follow-up independently.",
     "(Each team estimates, sets timelines, and follows up on delivery on its own.)",
     "Product model requires clear ownership, performance management, and outcome-based governance."),
    ("05", "No Product-Level Steering",
     "Existing forums are not structured to drive portfolio-level or product-based decisions.",
     "(Different recurring forums exist — but no forum covers portfolio-level decisions.)",
     "Product model requires governance aligned to product portfolios, with clear accountability and decision rights."),
    ("06", "No Common Language of Value",
     "Lack of standardized definitions of impact (client, business value) prevents consistent evaluation and comparison of work.",
     "(Each demand request is worked depending on perception of value and demand size.)",
     "Product model requires a shared taxonomy and classification model as a foundation for prioritization and planning."),
    ("07", "No Product Ownership Model",
     "Work is organized around projects and applications, with limited alignment to value streams or business outcomes.",
     "(Demand is organized by application silo, not by value stream. IT requests don't trace to business OKRs.)",
     "Product model requires organizing around products, with clear ownership and alignment to business OKRs."),
]

iss_y = Inches(1.05)
row_h = (SH - iss_y - Inches(0.55)) / len(issues)
num_w = Inches(0.7)
left_w = Inches(6.6)
arrow_w = Inches(0.3)
right_w = SW - Inches(1.0) - num_w - left_w - arrow_w

for i, (num, title, desc, evidence, response) in enumerate(issues):
    y = iss_y + i * row_h
    # Number badge (left)
    round_rect(s, Inches(0.5), y + Inches(0.04), num_w, row_h - Inches(0.1),
               BLUE_LIGHT, radius=0.05)
    text_box(s, Inches(0.5), y + Inches(0.04), num_w, row_h - Inches(0.1),
             num, size=18, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Issue title + description
    title_x = Inches(0.5) + num_w + Inches(0.1)
    text_box(s, title_x, y + Inches(0.02), left_w - Inches(0.1), Inches(0.28),
             title, size=12, bold=True, color=BLUE)
    text_box(s, title_x, y + Inches(0.30), left_w - Inches(0.1), row_h - Inches(0.55),
             desc, size=8.5, color=DARK)
    text_box(s, title_x, y + row_h - Inches(0.27), left_w - Inches(0.1), Inches(0.25),
             evidence, size=7.5, italic=True, color=GREY)
    # Arrow
    arrow_x = title_x + left_w
    text_box(s, arrow_x, y + Inches(0.04), arrow_w, row_h - Inches(0.1),
             "▶", size=14, bold=True, color=NAVY,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Response
    resp_x = arrow_x + arrow_w
    text_box(s, resp_x, y + Inches(0.05), right_w, row_h - Inches(0.1),
             response, size=10, color=DARK,
             anchor=MSO_ANCHOR.MIDDLE)
    # Separator line under row (except last)
    if i < len(issues) - 1:
        rect(s, title_x, y + row_h - Inches(0.01),
             SW - title_x - Inches(0.5), Inches(0.005), GREY)

footer_bar(s)



# ════════════════════════════════════════════════════════════════════════
# SLIDE 4b — Guiding Principles (connect to AS-IS deliverable vrevisada)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Four Guiding Principles · each principle answers specific structural issues from the AS-IS")

principles = [
    ("01", "DEMAND INTAKE & STRUCTURING", BLUE,
     "Resolves Issue 01 — No Product Entry Point",
     ["All demand captured through unified intake layer aligned to products / value streams",
      "Single entry point per product / value stream",
      "Standard intake criteria and mandatory fields",
      "Consistent demand categorization (type, size, complexity, value driver)",
      "Clear ownership for triage and backlog assignment"]),
    ("02", "VALUE-BASED PRIORITIZATION", ACCENT,
     "Resolves Issues 02 + 06 — No Value-Based Decisioning · No Common Language of Value",
     ["Demand evaluated using a standardized framework tied to business outcomes",
      "Common scoring model: value · urgency · risk · effort",
      "Transparent prioritization across competing demands",
      "Product-level backlog ownership",
      "Trade-offs made at portfolio / product governance level"]),
    ("03", "PRODUCT PORTFOLIO MGMT & PLANNING", GREEN,
     "Resolves Issue 03 — No End-to-End Product View",
     ["End-to-end visibility enables coordinated planning across demand · capacity · delivery",
      "Integrated portfolio view across all products",
      "Capacity planning aligned to product priorities",
      "Roadmaps managed at product level (not project level)",
      "Continuous tracking of delivery and outcomes"]),
    ("04", "PRODUCT GOVERNANCE & OWNERSHIP", PURPLE,
     "Resolves Issues 04 + 05 + 07 — Ownership · Steering · Product Model",
     ["Clear accountability and governance align decisions to product outcomes",
      "Defined Squad roles — Product Manager (NEW) · Tech Lead · Business Analyst",
      "Product-based governance forums with decision rights",
      "Vendor management aligned to product performance",
      "KPIs tied to product outcomes and business value"]),
]

p_y = Inches(1.1)
p_h = Inches(5.6)
ncols = 4
gap = Inches(0.15)
p_w = (SW - Inches(1.0) - gap * (ncols - 1)) / ncols

for i, (num, name, color, resolves, bullets) in enumerate(principles):
    x = Inches(0.5) + i * (p_w + gap)
    round_rect(s, x, p_y, p_w, p_h, WHITE, line=color, line_w=2, radius=0.04)
    rect(s, x, p_y, p_w, Inches(0.95), color)
    text_box(s, x, p_y + Inches(0.1), p_w, Inches(0.4),
             num, size=24, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.1), p_y + Inches(0.5), p_w - Inches(0.2), Inches(0.45),
             name, size=11, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # "Resolves" tag bar
    rect(s, x, p_y + Inches(0.95), p_w, Inches(0.42), LIGHT)
    text_box(s, x + Inches(0.15), p_y + Inches(0.97), p_w - Inches(0.3), Inches(0.4),
             resolves, size=8.5, bold=True, italic=True, color=color,
             anchor=MSO_ANCHOR.MIDDLE)
    add_lines(s, x + Inches(0.2), p_y + Inches(1.5),
              p_w - Inches(0.4), p_h - Inches(1.65),
              ["•  " + b for b in bullets], size=10.5, color=DARK,
              line_spacing=1.4)

footer_bar(s)



# ─── Section divider · 05 Stage 1 · Demand Management Framework ───
toc_slide(active_idx=4)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 5 — Demand Management Practices (five core blocks)
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



# ─── Section divider · 06 Operating Model · Crews & Bases ───
toc_slide(active_idx=5)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 5a — unFIX framework primer
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "unFIX framework · a shared language for the product-centric operating model")

# Lead
text_box(s, Inches(0.5), Inches(1.05), SW - Inches(1), Inches(0.45),
         "Before describing what we propose, here is the language we will use. "
         "unFIX is a flexible pattern library — we adopt only the building blocks Avangrid needs.",
         size=11, italic=True, color=BLUE, bold=True)

# Three building blocks
bb_y = Inches(1.7)
bb_h = Inches(3.0)
bb_w = (SW - Inches(1.2)) / 3
bb_gap = Inches(0.15)

blocks = [
    ("CREW", BLUE,
     "A stable, multi-functional team",
     ["3–9 people · single product or capability ownership",
      "Permanent — not a project team that disbands",
      "7 Crew archetypes — we use Value Stream · Capability · Experience · Platform · Partnership · Facilitation · Governance"]),
    ("BASE", GREEN,
     "A collection of Crews + Forums sharing a strategic alignment",
     ["4 Base types — Strongly Aligned · Loosely Aligned · Coupled · Segregated",
      "Avangrid Digital Products will be a Strongly Aligned Base",
      "Sibling bases (Digital CoE · P&S/P&T) operate Loosely Aligned in parallel"]),
    ("FORUM", AMBER,
     "A recurring coordination venue · no decision rights · signals only",
     ["10 Forum archetypes — we use Regional · Functional · Technological · Customer Journey · Product",
      "Surfaces dependencies, alignment, and reuse opportunities",
      "Decisions stay with Crews — Forums make them visible"]),
]

for i, (name, color, tag, items) in enumerate(blocks):
    x = Inches(0.5) + i * (bb_w + bb_gap)
    round_rect(s, x, bb_y, bb_w, bb_h, WHITE, line=color, line_w=2, radius=0.05)
    rect(s, x, bb_y, bb_w, Inches(0.55), color)
    text_box(s, x, bb_y + Inches(0.05), bb_w, Inches(0.45),
             name, size=18, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.2), bb_y + Inches(0.65),
             bb_w - Inches(0.4), Inches(0.6),
             tag, size=11, italic=True, color=color, bold=True)
    add_lines(s, x + Inches(0.2), bb_y + Inches(1.3),
              bb_w - Inches(0.4), bb_h - Inches(1.4),
              ["•  " + it for it in items], size=10, color=DARK,
              line_spacing=1.35)

# Bottom takeaway
tk_y = Inches(4.95)
round_rect(s, Inches(0.5), tk_y, SW - Inches(1), Inches(1.95),
           BLUE_PALE, line=NAVY, line_w=1.5, radius=0.05)
text_box(s, Inches(0.7), tk_y + Inches(0.1),
         SW - Inches(1.4), Inches(0.3),
         "WHY UNFIX FOR AVANGRID",
         size=10, bold=True, color=NAVY)
add_lines(s, Inches(0.7), tk_y + Inches(0.45),
          SW - Inches(1.4), tk_y + Inches(1.5),
          ["•  Maps cleanly onto Avangrid's reality — value streams, sub-teams, vendors and forums already exist; we are giving them a consistent name and a clear interface.",
           "•  Modular — we adopt what fits (Strongly Aligned Base + Value Stream Crews + 5 Forums) and ignore what doesn't (e.g. mob squads, two-pizza rule).",
           "•  Compatible with the existing P&S/P&T NOM v1.5 implementation and Digital CoE patterns — no conflict between siblings.",
           "•  Source — unfix.com — open framework by Jurgen Appelo."],
          size=10, color=DARK, line_spacing=1.35)

footer_bar(s)



# ════════════════════════════════════════════════════════════════════════
# SLIDE 5b — Strongly Aligned Base · master visual
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Strongly Aligned Base · the Digital Products Base hosts the full operating model")

# Outer base container
base_x = Inches(0.4); base_y = Inches(1.1); base_w = SW - Inches(3.4); base_h = Inches(5.7)
round_rect(s, base_x, base_y, base_w, base_h, BLUE_PALE, line=NAVY, line_w=2.5, radius=0.03)
text_box(s, base_x + Inches(0.2), base_y + Inches(0.08),
         base_w - Inches(0.4), Inches(0.32),
         "STRONGLY ALIGNED BASE  ·  Avangrid Digital Products",
         size=12, bold=True, color=NAVY)

# 8 Value Stream Crews (top row inside Base)
vs_y = base_y + Inches(0.5)
vs_h = Inches(1.6)
vs_names = [("M2C-SL", BLUE), ("M2C-RO", BLUE), ("CCC", BLUE),
            ("FOPS", GREEN), ("E&A", GREEN), ("NOC", GREEN),
            ("C&R", GREEN), ("IBD", GREEN)]
vs_w = (base_w - Inches(0.4) - Inches(0.06) * 7) / 8
text_box(s, base_x + Inches(0.2), vs_y - Inches(0.1),
         base_w - Inches(0.4), Inches(0.25),
         "▸ EIGHT VALUE STREAM CREWS", size=9, bold=True, color=NAVY)
for i, (n, c) in enumerate(vs_names):
    vx = base_x + Inches(0.2) + i * (vs_w + Inches(0.06))
    round_rect(s, vx, vs_y + Inches(0.18), vs_w, vs_h - Inches(0.18),
               WHITE, line=c, line_w=1.5, radius=0.05)
    rect(s, vx, vs_y + Inches(0.18), vs_w, Inches(0.35), c)
    text_box(s, vx, vs_y + Inches(0.18), vs_w, Inches(0.35),
             n, size=10, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, vx, vs_y + Inches(0.55),
             vs_w, vs_h - Inches(0.55),
             "VS\nCrew\n(Squad)", size=8.5, color=DARK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 4 Supporting Crews + 1 Facilitation Crew (middle row)
sup_y = vs_y + vs_h + Inches(0.2)
sup_h = Inches(1.4)
sup_names = [
    ("CAPABILITY", "Architecture\nCouncil", PURPLE),
    ("EXPERIENCE", "UX  ·  Customer\nJourney", PURPLE),
    ("PLATFORM", "DevOps · SAP Basis\nApp Sec · TOSCA\nData Ops · DocMgmt", PURPLE),
    ("PARTNERSHIP", "Vendor mgmt\n(centralized)", PURPLE),
    ("FACILITATION", "Process · Coaches\nKnowledge · PJM pool", AMBER),
]
text_box(s, base_x + Inches(0.2), sup_y - Inches(0.1),
         base_w - Inches(0.4), Inches(0.25),
         "▸ FOUR SUPPORTING CREWS  +  ONE FACILITATION CREW",
         size=9, bold=True, color=NAVY)
sup_w = (base_w - Inches(0.4) - Inches(0.08) * 4) / 5
for i, (n, body, c) in enumerate(sup_names):
    sx = base_x + Inches(0.2) + i * (sup_w + Inches(0.08))
    round_rect(s, sx, sup_y + Inches(0.18), sup_w, sup_h - Inches(0.18),
               WHITE, line=c, line_w=1.5, radius=0.05)
    rect(s, sx, sup_y + Inches(0.18), sup_w, Inches(0.32), c)
    text_box(s, sx, sup_y + Inches(0.18), sup_w, Inches(0.32),
             n, size=9, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, sx + Inches(0.05), sup_y + Inches(0.55),
             sup_w - Inches(0.1), sup_h - Inches(0.6),
             body, size=8.5, color=DARK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 5 Forums (bottom row, slim)
fr_y = sup_y + sup_h + Inches(0.2)
fr_h = Inches(0.85)
forums = ["Regional", "Functional", "Technological", "Customer Journey", "Product (BRP)"]
text_box(s, base_x + Inches(0.2), fr_y - Inches(0.1),
         base_w - Inches(0.4), Inches(0.25),
         "▸ FIVE FORUMS  ·  signals  ·  no decision rights",
         size=9, bold=True, color=NAVY)
fr_w = (base_w - Inches(0.4) - Inches(0.08) * 4) / 5
for i, f in enumerate(forums):
    fx = base_x + Inches(0.2) + i * (fr_w + Inches(0.08))
    round_rect(s, fx, fr_y + Inches(0.15), fr_w, fr_h - Inches(0.15),
               AMBER_LIGHT, line=AMBER, line_w=1, radius=0.3)
    text_box(s, fx, fr_y + Inches(0.15), fr_w, fr_h - Inches(0.15),
             f, size=10, bold=True, color=AMBER,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Right-side panel — sibling + segregated bases
sb_x = base_x + base_w + Inches(0.15)
sb_w = SW - sb_x - Inches(0.4)

# Loosely aligned sibling bases (top)
sb1_h = Inches(2.6)
round_rect(s, sb_x, base_y, sb_w, sb1_h, WHITE, line=GREY_DARK, line_w=2, radius=0.05)
rect(s, sb_x, base_y, sb_w, Inches(0.4), GREY_DARK)
text_box(s, sb_x, base_y + Inches(0.04), sb_w, Inches(0.35),
         "LOOSELY ALIGNED SIBLINGS",
         size=10, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text_box(s, sb_x + Inches(0.15), base_y + Inches(0.5),
         sb_w - Inches(0.3), Inches(0.4),
         "DIGITAL COE", size=10, bold=True, color=NAVY)
text_box(s, sb_x + Inches(0.15), base_y + Inches(0.78),
         sb_w - Inches(0.3), Inches(0.6),
         "CX customer-facing · OPCO-aligned · 5 product teams",
         size=8.5, italic=True, color=DARK)
text_box(s, sb_x + Inches(0.15), base_y + Inches(1.4),
         sb_w - Inches(0.3), Inches(0.4),
         "P&S / P&T  (NETWORKS IT)", size=10, bold=True, color=NAVY)
text_box(s, sb_x + Inches(0.15), base_y + Inches(1.7),
         sb_w - Inches(0.3), Inches(0.85),
         "Networks IT operating model · NOM v1.5 implementing · parallel cadence",
         size=8.5, italic=True, color=DARK)

# Segregated SW Delivery (bottom)
sb2_y = base_y + sb1_h + Inches(0.2)
sb2_h = base_y + base_h - sb2_y
round_rect(s, sb_x, sb2_y, sb_w, sb2_h, WHITE, line=RED, line_w=2, radius=0.05)
rect(s, sb_x, sb2_y, sb_w, Inches(0.4), RED)
text_box(s, sb_x, sb2_y + Inches(0.04), sb_w, Inches(0.35),
         "SEGREGATED BASE",
         size=10, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text_box(s, sb_x + Inches(0.15), sb2_y + Inches(0.5),
         sb_w - Inches(0.3), Inches(0.35),
         "SW DELIVERY  ·  vendors", size=10, bold=True, color=NAVY)
add_lines(s, sb_x + Inches(0.15), sb2_y + Inches(0.85),
          sb_w - Inches(0.3), sb2_h - Inches(0.95),
          ["•  Accenture",
           "•  SWF",
           "•  HCL",
           "•  Infosys",
           "•  Smart Metering vendor",
           "•  SAP partners",
           "•  Itron · Esri · GE / Siemens"],
          size=9, color=DARK, line_spacing=1.3)

footer_bar(s)



# ════════════════════════════════════════════════════════════════════════
# SLIDE 5c — Supporting Crews · 4 + 1 Facilitation
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Supporting Crews · five Crews enable the eight Value Streams without slowing them down")

text_box(s, Inches(0.5), Inches(1.05), SW - Inches(1), Inches(0.4),
         "Each Value Stream Crew owns its product end to end. Five supporting Crews provide the "
         "shared capabilities every Value Stream needs — without recreating the same work eight times.",
         size=11, italic=True, color=BLUE, bold=True)

sup_specs = [
    ("CAPABILITY", PURPLE, "Architecture & technical standards",
     ["Architecture Council — rotating membership across Tech Leads",
      "Solution Architect (SAP) and Integration Architect roles owned here",
      "Time-boxed feasibility reviews at the intake gate (≤ 2 h)",
      "Sets non-functional standards · reuse patterns · technical roadmaps",
      "Output: Feasible · Conditional · Not feasible — advisory, no veto"]),
    ("EXPERIENCE", PURPLE, "User experience and customer journey",
     ["UX research and design embedded across Value Streams",
      "Owns end-to-end Customer Journey artifacts (M2C · CCC · Field)",
      "Aligns digital channels and self-service experiences",
      "Partners with Heather LaJoie (Customer Operations) for outcome KPIs"]),
    ("PLATFORM", PURPLE, "Shared technical foundation",
     ["DevOps · CI/CD pipelines · environment management",
      "App Security · TOSCA test automation · SAP Basis",
      "Data Ops — Data Governance Lead · MDM · BW / SAC analytics ownership",
      "Document Management (5 cross-stream apps from IT Networks Assessment)"]),
    ("PARTNERSHIP", PURPLE, "Vendor relationship management — centralized",
     ["Today: vendor management is distributed across PJMs and sub-team leads",
      "Centralization is a real change — single accountability for vendor health",
      "Owns contracts · SLAs · scorecards · joint roadmaps",
      "Coordinates across Accenture · SWF · HCL · Infosys · OEMs"]),
    ("FACILITATION", AMBER, "Process improvement and cross-Crew coaching",
     ["Process Improvement · Coaches · Knowledge Manager",
      "Sarath Garikapati's PJM pool (5–6 PJMs) — shared resource across Crews",
      "Owns playbooks · onboarding · operating model evolution",
      "Note — we use Kanban with WIP limits, not Scrum (no Scrum Masters)"]),
]

sc_y = Inches(1.6)
sc_h = Inches(5.3)
sc_w = (SW - Inches(1.0) - Inches(0.12) * 4) / 5
for i, (name, color, tag, items) in enumerate(sup_specs):
    x = Inches(0.5) + i * (sc_w + Inches(0.12))
    round_rect(s, x, sc_y, sc_w, sc_h, WHITE, line=color, line_w=2, radius=0.05)
    rect(s, x, sc_y, sc_w, Inches(0.5), color)
    text_box(s, x, sc_y + Inches(0.05), sc_w, Inches(0.4),
             name, size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.15), sc_y + Inches(0.6),
             sc_w - Inches(0.3), Inches(0.6),
             tag, size=10, italic=True, color=color, bold=True)
    add_lines(s, x + Inches(0.2), sc_y + Inches(1.25),
              sc_w - Inches(0.4), sc_h - Inches(1.4),
              ["•  " + it for it in items], size=9, color=DARK,
              line_spacing=1.3)

footer_bar(s)



# ════════════════════════════════════════════════════════════════════════
# SLIDE 5d — Sibling + Segregated Bases · how Digital Products fits in the broader org
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Bases around us · two Loosely Aligned siblings and one Segregated SW Delivery base")

text_box(s, Inches(0.5), Inches(1.05), SW - Inches(1), Inches(0.4),
         "Digital Products is not standalone. Two sibling Bases run product-centric models in parallel, "
         "and the actual software delivery happens in a Segregated Base operated by external vendors.",
         size=11, italic=True, color=BLUE, bold=True)

# 3 columns: Sibling 1 · Sibling 2 · Segregated
b_y = Inches(1.65)
b_h = Inches(5.0)
b_w = (SW - Inches(1.2)) / 3
b_gap = Inches(0.15)

bases = [
    ("DIGITAL COE", "Loosely Aligned sibling", GREY_DARK,
     ["Customer-experience customer-facing applications",
      "OPCO-aligned (UIL · CMP · NY)",
      "5 product teams already operating in product mode",
      "Owns the customer digital channels and CX layer"],
     "INTERFACE WITH DIGITAL PRODUCTS",
     ["CCC · M2C-RO trade-off discussions go to Customer Journey Forum",
      "CX backlog and CRM backlog reconciled monthly",
      "Shared standards via Capability Crew"]),
    ("P&S / P&T  ·  NETWORKS IT", "Loosely Aligned sibling", GREY_DARK,
     ["Networks IT operating model implementing NOM v1.5",
      "Field Operations · Engineering · System Operations focus",
      "Parallel cadence — independent governance",
      "Heavy regulatory and reliability orientation"],
     "INTERFACE WITH DIGITAL PRODUCTS",
     ["Field Ops · E&A · NOC overlap → Technological Forum",
      "Storm response coordination via Regional Forum",
      "Shared platform services (Document Mgmt · GIS data)"]),
    ("SW DELIVERY  ·  vendors", "Segregated Base", RED,
     ["Vendors operate outside the Avangrid Base",
      "Avangrid = product owner · vendor = delivery muscle",
      "Members — Accenture · SWF · HCL · Infosys · Smart Metering vendor · SAP partners · Itron · Esri · GE / Siemens",
      "Delivery happens via stable Crews (Option A) or pool (Option B)"],
     "INTERFACE WITH DIGITAL PRODUCTS",
     ["Partnership Crew owns vendor contracts and SLAs",
      "Each Value Stream Crew has a vendor-extension team",
      "Joint roadmaps reviewed quarterly at Big Room Planning"]),
]

for i, (name, tag, color, body, intf_lbl, intf_items) in enumerate(bases):
    x = Inches(0.5) + i * (b_w + b_gap)
    round_rect(s, x, b_y, b_w, b_h, WHITE, line=color, line_w=2, radius=0.05)
    rect(s, x, b_y, b_w, Inches(0.55), color)
    text_box(s, x, b_y + Inches(0.05), b_w, Inches(0.45),
             name, size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.15), b_y + Inches(0.65),
             b_w - Inches(0.3), Inches(0.3),
             tag, size=10, italic=True, color=color, bold=True)
    add_lines(s, x + Inches(0.2), b_y + Inches(1.05),
              b_w - Inches(0.4), Inches(2.0),
              ["•  " + it for it in body], size=9.5, color=DARK,
              line_spacing=1.3)
    # Interface mini-block
    intf_y = b_y + Inches(3.1)
    rect(s, x + Inches(0.2), intf_y, b_w - Inches(0.4), Inches(0.005), color)
    text_box(s, x + Inches(0.2), intf_y + Inches(0.05),
             b_w - Inches(0.4), Inches(0.3),
             intf_lbl, size=9, bold=True, color=color)
    add_lines(s, x + Inches(0.2), intf_y + Inches(0.4),
              b_w - Inches(0.4), b_h - Inches(3.55),
              ["•  " + it for it in intf_items], size=8.5,
              color=DARK, line_spacing=1.3)

footer_bar(s)



# ════════════════════════════════════════════════════════════════════════
# SLIDE 6 — TOM Overview with applications
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Target Operating Model · eight Value Stream Teams aligned to customer outcomes")

streams_apps = [
    ("Meter-to-Cash · Service Lifecycle",
     "Customer signs up · activates · changes · disconnects",
     "SAP Metering · MV90", ACCENT,
     "DOMAIN · Customer Service / M2C"),
    ("Meter-to-Cash · Revenue Operations",
     "Accurate billing · easy payment · fair recovery",
     "SAP FICA · ECC · StreamServe", ACCENT,
     "DOMAIN · Customer Service · Finance"),
    ("Customer Care & Contact Center",
     "Issue resolution · self-service · voice of customer",
     "SAP CRM · Medallia · Blue Prism", ACCENT,
     "DOMAIN · Customer Service"),
    ("Field Operations",
     "Field worker / inspector / lineman productivity",
     "Lemur-SF · FWM · ARCOS · Field Force", BLUE,
     "DOMAIN · Field / Network Ops"),
    ("Engineering & Asset Management",
     "Engineer / planner designs with reliable data",
     "ArcFM · ProjectWise · Bentley · CYME", BLUE,
     "DOMAIN · Asset Management"),
    ("Network Operations & Control",
     "Grid operator prevents and minimizes outages",
     "Echelon · Cimplicity · SLAM · Switching Orders", BLUE,
     "DOMAIN · Network Operations"),
    ("Corporate & Regulatory",
     "HR / employee / regulatory user self-service",
     "OpenText · Rate Case Advisor · Utilisphere", BLUE,
     "DOMAIN · Finance · HR · Procurement · Regulatory"),
    ("Iberdrola Corporate Services",
     "Spain finance / procurement / budget",
     "SAP FI · PO · Budget · S/4HANA", BLUE,
     "DOMAIN · Finance · Procurement (Iberdrola)"),
]

grid_y = Inches(1.05)
grid_h = Inches(5.0)
ncols = 4
nrows = 2
gap = Inches(0.12)
col_w = (SW - Inches(1.0) - gap * (ncols - 1)) / ncols
row_h = (grid_h - gap) / nrows

for i, (name, outcome, apps, color, domain) in enumerate(streams_apps):
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
    # Domain badge — small italic line below title
    text_box(s, x + Inches(0.15), y + Inches(0.72),
             col_w - Inches(0.3), Inches(0.22),
             domain, size=8, bold=True, italic=True, color=color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.15), y + Inches(0.95),
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
# SLIDE 7-pre — TOM lifecycle · four stages · Stage 01 is the focus of this proposal
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Target Operating Model lifecycle · four stages · Wave 1 focuses on Stage 01")

text_box(s, Inches(0.5), Inches(1.05), SW - Inches(1), Inches(0.45),
         "A complete product-centric operating model spans four stages, end to end. "
         "This proposal designs Stage 01 in detail and pilots it. Stages 02–04 are acknowledged here "
         "but are out of Wave 1 scope — they will be designed once Stage 01 is validated.",
         size=11, italic=True, color=BLUE, bold=True)

stages_def = [
    ("01", "DEMAND MANAGEMENT",
     "How demand enters, is shaped, classified and prioritized.",
     ["Single intake per Value Stream",
      "Definition of Ready · Class of Service",
      "Two-gate governance (weekly + monthly)",
      "Pain Score prioritization"],
     "WAVE 1 — IN SCOPE", GREEN, True),
    ("02", "PLANNING",
     "How approved demand becomes a sequenced roadmap with capacity commitments.",
     ["Quarterly planning cadence",
      "Capacity allocation per Value Stream",
      "Big Room Planning · cross-VS dependency mapping",
      "Roadmap publication and OKR linking"],
     "OUT OF WAVE 1 SCOPE", GREY_DARK, False),
    ("03", "DELIVERY",
     "How squads + vendors execute the work end to end.",
     ["Discovery + Delivery dual track",
      "Initiative → Story default · Epic when > 3 months",
      "Kanban with WIP limits",
      "Vendor-extension execution model"],
     "OUT OF WAVE 1 SCOPE", GREY_DARK, False),
    ("04", "OPERATIONS  &  RUN",
     "How delivered products are operated, supported, evolved and retired.",
     ["Hypercare and stabilization",
      "Production support + incident management",
      "Continuous improvement loop",
      "Decommission and lifecycle exit"],
     "OUT OF WAVE 1 SCOPE", GREY_DARK, False),
]

st_y = Inches(1.7)
st_h = Inches(4.7)
st_w = (SW - Inches(1.0) - Inches(0.15) * 3) / 4
for i, (num, name, tag, items, status, color, in_scope) in enumerate(stages_def):
    x = Inches(0.5) + i * (st_w + Inches(0.15))
    border_w = 3 if in_scope else 1.5
    round_rect(s, x, st_y, st_w, st_h, WHITE, line=color, line_w=border_w, radius=0.05)
    rect(s, x, st_y, st_w, Inches(0.55), color)
    text_box(s, x + Inches(0.15), st_y + Inches(0.05),
             Inches(0.7), Inches(0.45),
             num, size=18, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.85), st_y + Inches(0.05),
             st_w - Inches(1.0), Inches(0.45),
             name, size=11, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    # Tag/description
    text_box(s, x + Inches(0.2), st_y + Inches(0.7),
             st_w - Inches(0.4), Inches(0.7),
             tag, size=10, italic=True, color=color, bold=True)
    add_lines(s, x + Inches(0.25), st_y + Inches(1.5),
              st_w - Inches(0.5), st_h - Inches(2.4),
              ["•  " + it for it in items], size=9.5,
              color=DARK if in_scope else GREY, line_spacing=1.35)
    # Status pill at the bottom
    pill_y = st_y + st_h - Inches(0.7)
    round_rect(s, x + Inches(0.25), pill_y,
               st_w - Inches(0.5), Inches(0.4),
               color if in_scope else LIGHT,
               line=color, line_w=1, radius=0.3)
    text_box(s, x + Inches(0.25), pill_y,
             st_w - Inches(0.5), Inches(0.4),
             status, size=9, bold=True,
             color=WHITE if in_scope else color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Bottom takeaway
tk_y = Inches(6.55)
round_rect(s, Inches(0.5), tk_y, SW - Inches(1), Inches(0.4),
           BLUE_PALE, line=NAVY, line_w=1, radius=0.05)
text_box(s, Inches(0.7), tk_y, SW - Inches(1.4), Inches(0.4),
         "WHY START WITH STAGE 01 — fixing intake first protects every downstream stage. "
         "Without a single front door and a value-based prioritization, planning and delivery inherit the same fragmentation.",
         size=9.5, italic=True, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)

footer_bar(s)



# ════════════════════════════════════════════════════════════════════════
# SLIDE 7 — Stage 01 Demand Management overview
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
# SLIDES 7a–7e — Stage 1 visual deep-dives (one per dimension)
# Premium, visual layouts — replace text walls with diagrams + callouts.
# ════════════════════════════════════════════════════════════════════════
def stage1_kicker(slide, color, block_name, tagline, resolves):
    rect(slide, 0, Inches(1.0), SW, Inches(0.55), color)
    text_box(slide, Inches(0.5), Inches(1.05),
             Inches(3.5), Inches(0.45),
             f"STAGE 01 · {block_name}",
             size=12, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    text_box(slide, Inches(4.0), Inches(1.05),
             Inches(6.0), Inches(0.45),
             tagline, size=11, italic=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    text_box(slide, SW - Inches(4.0), Inches(1.05),
             Inches(3.5), Inches(0.45),
             resolves, size=9, bold=True, color=ACCENT,
             align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)


def callout_row(slide, y, items, color):
    n = len(items)
    gap = Inches(0.15)
    w = (SW - Inches(1.0) - gap * (n - 1)) / n
    h = Inches(1.1)
    for i, (kicker, body) in enumerate(items):
        x = Inches(0.5) + i * (w + gap)
        round_rect(slide, x, y, w, h, LIGHT, line=color, line_w=1, radius=0.04)
        rect(slide, x, y, Inches(0.08), h, color)
        text_box(slide, x + Inches(0.2), y + Inches(0.08),
                 w - Inches(0.3), Inches(0.3),
                 kicker, size=10, bold=True, color=color)
        text_box(slide, x + Inches(0.2), y + Inches(0.38),
                 w - Inches(0.3), h - Inches(0.45),
                 body, size=9, color=DARK)


# ─── 7a · CAPABILITIES (visual funnel) ─────────────────────────────────
s = blank()
header_bar(s, "Capabilities · 14 capabilities organize into a four-stage demand funnel")
stage1_kicker(s, GREEN, "CAPABILITIES",
              "What the organization must be able to do — end to end.",
              "Resolves Issues 01 + 02 + 06")

# 4-stage funnel
fn_y = Inches(1.85)
fn_h = Inches(2.4)
stages = [
    ("INTAKE", BLUE,
     ["Demand intake channels (Jira request type)",
      "Standard intake formats (DoR)",
      "Channel consolidation",
      "Technical feasibility input"]),
    ("CLASSIFY", ACCENT,
     ["Demand classification",
      "Run · Change · Regulatory · Innovation",
      "Business value definition (Pain Score)",
      "Acceptance criteria"]),
    ("DECIDE", AMBER,
     ["Triage & validation (weekly)",
      "Demand routing & assignment",
      "Urgent / exception requests (≤20%)"]),
    ("VISUALIZE", GREEN,
     ["Demand registration & logging",
      "Demand visibility & tracking",
      "Duplicate & overlap identification"]),
]
fn_gap = Inches(0.15)
fn_w = (SW - Inches(1.0) - fn_gap * (len(stages) - 1)) / len(stages)
for i, (name, color, items) in enumerate(stages):
    x = Inches(0.5) + i * (fn_w + fn_gap)
    round_rect(s, x, fn_y, fn_w, fn_h, WHITE, line=color, line_w=2, radius=0.05)
    rect(s, x, fn_y, fn_w, Inches(0.5), color)
    text_box(s, x, fn_y + Inches(0.05), fn_w, Inches(0.4),
             name, size=14, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_lines(s, x + Inches(0.2), fn_y + Inches(0.6),
              fn_w - Inches(0.4), fn_h - Inches(0.7),
              ["•  " + it for it in items], size=9.5, color=DARK,
              line_spacing=1.35)
    if i < len(stages) - 1:
        text_box(s, x + fn_w - Inches(0.15), fn_y + fn_h/2 - Inches(0.18),
                 Inches(0.3), Inches(0.4), "▶",
                 size=18, bold=True, color=color,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Bottom callouts
callout_row(s, Inches(4.5), [
    ("ONE FRONT DOOR",
     "Jira request type per Value Stream replaces Email · Teams · Word · OneNote ad-hoc requests."),
    ("DEFINITION OF READY",
     "Mandatory fields: title · BV · acceptance criteria · CoS · sponsor."),
    ("PAIN SCORE",
     "H(x) = (BV·w₁ + UX·w₂ + Safety·w₃) / Effort — sponsor signs off at intake."),
    ("EXPEDITE CAP",
     "Bypass priority capped ≤ 20 % of throughput. No silent overrides."),
], GREEN)

# Closing takeaway band
tk_y = Inches(5.85)
round_rect(s, Inches(0.5), tk_y, SW - Inches(1), Inches(1.05),
           BLUE_PALE, line=NAVY, line_w=1.5, radius=0.05)
text_box(s, Inches(0.7), tk_y + Inches(0.1),
         SW - Inches(1.4), Inches(0.3),
         "TAKEAWAY", size=10, bold=True, color=NAVY)
text_box(s, Inches(0.7), tk_y + Inches(0.4),
         SW - Inches(1.4), Inches(0.6),
         "The 14 capabilities are not a checklist — they are sequenced. Demand only moves stage-to-stage "
         "after passing the previous gate, which is what converts inflow into traceable, governed work.",
         size=11, italic=True, color=DARK)
footer_bar(s)


# ─── 7b · ROLES (4 roles + accountabilities) ───────────────────────────
s = blank()
header_bar(s, "Roles · four accountabilities own the demand flow end to end")
stage1_kicker(s, TEAL, "ROLES",
              "Who owns and operates the framework.",
              "Resolves Issues 04 + 07")

roles_visual = [
    ("BUSINESS REQUESTOR", BLUE,
     "Originating business unit",
     ["Submits demand via Jira request type (one per VS)",
      "Sponsors the request and signs off business value",
      "Receives status updates through requester portal"],
     "Decides WHAT business outcome is needed"),
    ("PRODUCT MANAGER  [NEW]", ACCENT,
     "One per Value Stream Team · IT-business bridge",
     ["Owns backlog priority for the Value Stream",
      "Owns Pain Score and outcome KPIs",
      "Bridges IT to business — Process Owners, Demand Manager scope folded in",
      "Chairs the Portfolio review for the VS"],
     "Decides WHEN and IN WHAT ORDER work is done"),
    ("ARCHITECT REVIEWER", GREEN,
     "Architecture Council (rotating)",
     ["Provides feasibility view at intake gate (≤2 h)",
      "Verdict: Feasible · Conditional · Not feasible",
      "Advisory — does not veto"],
     "Decides HOW it should be built"),
    ("TRIAGE / DEMAND MGR", PURPLE,
     "Cross-VS chair",
     ["Chairs weekly Gate 1 triage (30 min)",
      "Audits scoring and intake quality",
      "Escalates cross-VS conflicts to Gate 2"],
     "Decides IF the demand is ready to enter the funnel"),
]
ry = Inches(1.85)
rh = Inches(3.6)
rw = (SW - Inches(1.0) - Inches(0.15) * 3) / 4
for i, (name, color, scope, duties, accountability) in enumerate(roles_visual):
    x = Inches(0.5) + i * (rw + Inches(0.15))
    round_rect(s, x, ry, rw, rh, WHITE, line=color, line_w=2, radius=0.05)
    rect(s, x, ry, rw, Inches(0.55), color)
    text_box(s, x + Inches(0.1), ry + Inches(0.05),
             rw - Inches(0.2), Inches(0.45),
             name, size=11, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.15), ry + Inches(0.65),
             rw - Inches(0.3), Inches(0.32),
             scope, size=9.5, italic=True, color=color)
    add_lines(s, x + Inches(0.2), ry + Inches(1.05),
              rw - Inches(0.4), Inches(2.0),
              ["•  " + d for d in duties], size=9, color=DARK, line_spacing=1.35)
    # Accountability footer
    rect(s, x, ry + rh - Inches(0.5), rw, Inches(0.5), LIGHT)
    text_box(s, x + Inches(0.15), ry + rh - Inches(0.45),
             rw - Inches(0.3), Inches(0.4),
             accountability, size=8.5, bold=True, italic=True, color=color,
             anchor=MSO_ANCHOR.MIDDLE)

# Bottom takeaway
tk_y = Inches(5.65)
round_rect(s, Inches(0.5), tk_y, SW - Inches(1), Inches(1.25),
           AMBER_LIGHT, line=AMBER, line_w=1.5, radius=0.05)
text_box(s, Inches(0.7), tk_y + Inches(0.1),
         SW - Inches(1.4), Inches(0.3),
         "WHY THE PRODUCT MANAGER ROLE IS NEW", size=10, bold=True, color=AMBER)
text_box(s, Inches(0.7), tk_y + Inches(0.42),
         SW - Inches(1.4), Inches(0.8),
         "Today Avangrid has Project Managers (PJMs), Sr. Managers and Champions — not Product "
         "Managers. The PM role owns outcomes (not just delivery) and is the single accountable name "
         "for each Value Stream's portfolio. This role is created from scratch as part of the TOM.",
         size=10, color=DARK, italic=True)
footer_bar(s)


# ─── 7c · TOOLS (technology stack) ─────────────────────────────────────
s = blank()
header_bar(s, "Tools · five layers consolidate today's fragmented toolset into one stack")
stage1_kicker(s, AMBER, "TOOLS",
              "The systems and platforms used end to end.",
              "Resolves Issue 03")

# Stack diagram on the left, comparison on the right
stack_x = Inches(0.5)
stack_w = Inches(6.5)
stack_y = Inches(1.85)
stack_layers = [
    ("01  CAPTURE",     "One Jira request type per Value Stream", BLUE),
    ("02  TRACK",       "Jira backlog · Initiative → Story · Run vs Build separated tracks per squad", BLUE),
    ("03  COMMUNICATE", "Requester portal · auto-status · forum routing", ACCENT),
    ("04  DOCUMENT",    "Confluence (DoR · CoS · per-VS product brief)", GREEN),
    ("05  REPORT",      "Demand dashboard · 5 views · refresh < 1 h", PURPLE),
]
slh = Inches(0.85)
sl_gap = Inches(0.08)
for i, (lbl, body, color) in enumerate(stack_layers):
    y = stack_y + i * (slh + sl_gap)
    round_rect(s, stack_x, y, stack_w, slh, WHITE, line=color, line_w=2, radius=0.04)
    rect(s, stack_x, y, Inches(1.7), slh, color)
    text_box(s, stack_x, y, Inches(1.7), slh,
             lbl, size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, stack_x + Inches(1.85), y + Inches(0.05),
             stack_w - Inches(1.95), slh - Inches(0.1),
             body, size=10.5, color=DARK,
             anchor=MSO_ANCHOR.MIDDLE)

# Right side — before/after comparison
comp_x = Inches(7.3)
comp_w = SW - comp_x - Inches(0.5)
# Before
round_rect(s, comp_x, stack_y, comp_w, Inches(2.4),
           RED_LIGHT, line=RED, line_w=1.5, radius=0.04)
text_box(s, comp_x + Inches(0.15), stack_y + Inches(0.1),
         comp_w - Inches(0.3), Inches(0.3),
         "BEFORE  ·  fragmented", size=11, bold=True, color=RED)
add_lines(s, comp_x + Inches(0.2), stack_y + Inches(0.45),
          comp_w - Inches(0.4), Inches(1.9),
          ["•  6+ disconnected channels (Jira · IT Now · RITM · SharePoint · Email · Word)",
           "•  No standard intake form",
           "•  No common backlog · per-team spreadsheets",
           "•  Status communicated by ad-hoc replies",
           "•  Reporting manual, monthly, after the fact"],
          size=9.5, color=DARK, line_spacing=1.35)

# After
after_y = stack_y + Inches(2.55)
round_rect(s, comp_x, after_y, comp_w, Inches(2.4),
           GREEN_LIGHT, line=GREEN, line_w=1.5, radius=0.04)
text_box(s, comp_x + Inches(0.15), after_y + Inches(0.1),
         comp_w - Inches(0.3), Inches(0.3),
         "AFTER  ·  one stack", size=11, bold=True, color=GREEN)
add_lines(s, comp_x + Inches(0.2), after_y + Inches(0.45),
          comp_w - Inches(0.4), Inches(1.9),
          ["•  One Jira request type per Value Stream",
           "•  Mandatory DoR fields enforced at submit",
           "•  Single backlog system of record",
           "•  Live requester portal · auto-status",
           "•  Demand dashboard refreshed under 1 hour"],
          size=9.5, color=DARK, line_spacing=1.35)
footer_bar(s)


# ─── 7c-bridge · Portfolio Architecture (3 layers) ─────────────────────
# Show the structural artifact BEFORE governance maps gates onto it.
# ────────────────────────────────────────────────────────────────────────
s = blank()
header_bar(s, "Portfolio architecture · two work levels — Initiatives at portfolio · Stories at squad")

l0_y = Inches(1.1)
round_rect(s, Inches(2.5), l0_y, Inches(8.3), Inches(0.85),
           BLUE_LIGHT, line=BLUE, line_w=2, radius=0.04)
text_box(s, Inches(2.7), l0_y + Inches(0.05), Inches(8), Inches(0.4),
         "VP ROLL-UP DASHBOARD",
         size=14, bold=True, color=BLUE)
text_box(s, Inches(2.7), l0_y + Inches(0.4), Inches(8), Inches(0.4),
         "Aggregated view across both portfolios · monthly Portfolio Sync",
         size=10, color=DARK, italic=True)

rect(s, Inches(6.63), l0_y + Inches(0.85), Inches(0.04), Inches(2.15) - (l0_y + Inches(0.85)), NAVY)

l1_y = Inches(2.15)
pb_w = Inches(6); pb_h = Inches(1.2)

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

rect(s, Inches(6.63), l1_y + pb_h, Inches(0.04), Inches(3.85) - (l1_y + pb_h), NAVY)

l2_y = Inches(4.1)
text_box(s, Inches(0.5), l2_y - Inches(0.3), SW - Inches(1), Inches(0.3),
         "EIGHT TEAM KANBANS  ·  one per Squad  ·  cards = STORIES  ·  five phases × dual track",
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

note_y = Inches(6.7)
round_rect(s, Inches(0.5), note_y, SW - Inches(1), Inches(0.55),
           AMBER_LIGHT, line=AMBER, line_w=1, radius=0.04)
text_box(s, Inches(0.7), note_y + Inches(0.05),
         Inches(2.0), Inches(0.45),
         "EPIC IS OPTIONAL",
         size=10, bold=True, color=AMBER, anchor=MSO_ANCHOR.MIDDLE)
text_box(s, Inches(2.6), note_y + Inches(0.05),
         SW - Inches(3.1), Inches(0.45),
         "Use Epic only when an Initiative is too large for one quarter (> 3 months) — "
         "to slice it into deliverable chunks. Default flow is Initiative → Story.",
         size=9.5, color=DARK, anchor=MSO_ANCHOR.MIDDLE)

footer_bar(s)



# ─── 7d · GOVERNANCE (two-gate flow) ───────────────────────────────────
s = blank()
header_bar(s, "Governance · two gates absorb the demand · one cadence drives the portfolio")
stage1_kicker(s, INDIGO, "GOVERNANCE",
              "Decisions, forums, and controls that keep the funnel honest.",
              "Resolves Issues 04 + 05")

# Gate 1 / Gate 2 / Big Room
gv_y = Inches(1.85)
gv_h = Inches(2.6)
gates = [
    ("GATE 1 · TRIAGE", "Weekly · 30 min", BLUE,
     "Demand Manager (chair) · PMs · Architect",
     ["Accept", "Send-back", "Reject", "Redirect"],
     "Filters and routes demand. Verdict logged with reason. SLA: 5 business days."),
    ("GATE 2 · PORTFOLIO", "Monthly · 90 min", ACCENT,
     "VP · Directors · Product Managers",
     ["Prioritize", "Defer", "Cancel", "Re-scope"],
     "Trade-offs across Value Streams. Capacity vs CoS commitments. Owns the roadmap."),
    ("BIG ROOM PLANNING", "Quarterly · ½ day", GREEN,
     "All VSs + business sponsors + Iberdrola",
     ["Confirm WHY", "Align HOW", "Spot DEPS", "Lock OKRs"],
     "Cross-portfolio alignment. Surfaces dependencies early and locks quarterly OKRs."),
]
gv_gap = Inches(0.2)
gv_w = (SW - Inches(1.0) - gv_gap * 2) / 3
for i, (name, cadence, color, who, outcomes, body) in enumerate(gates):
    x = Inches(0.5) + i * (gv_w + gv_gap)
    round_rect(s, x, gv_y, gv_w, gv_h, WHITE, line=color, line_w=2, radius=0.05)
    rect(s, x, gv_y, gv_w, Inches(0.55), color)
    text_box(s, x + Inches(0.1), gv_y + Inches(0.05),
             gv_w - Inches(0.2), Inches(0.45),
             name, size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.15), gv_y + Inches(0.65),
             gv_w - Inches(0.3), Inches(0.3),
             cadence, size=10, italic=True, color=color)
    text_box(s, x + Inches(0.15), gv_y + Inches(0.95),
             gv_w - Inches(0.3), Inches(0.3),
             who, size=9, color=DARK)
    # Outcome chips
    chip_y = gv_y + Inches(1.35)
    chip_w = (gv_w - Inches(0.3) - Inches(0.1) * 3) / 4
    for j, oc in enumerate(outcomes):
        cx = x + Inches(0.15) + j * (chip_w + Inches(0.08))
        round_rect(s, cx, chip_y, chip_w, Inches(0.3), color, radius=0.4)
        text_box(s, cx, chip_y, chip_w, Inches(0.3),
                 oc, size=8.5, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.15), gv_y + Inches(1.8),
             gv_w - Inches(0.3), Inches(0.7),
             body, size=9, italic=True, color=DARK)
    if i < 2:
        text_box(s, x + gv_w - Inches(0.05), gv_y + gv_h/2 - Inches(0.2),
                 gv_gap, Inches(0.4), "▶",
                 size=18, bold=True, color=GREY,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Bottom: escalation lane
esc_y = Inches(4.7)
round_rect(s, Inches(0.5), esc_y, SW - Inches(1), Inches(0.7),
           LIGHT, line=NAVY, line_w=1.5, radius=0.05)
text_box(s, Inches(0.7), esc_y + Inches(0.05),
         Inches(3), Inches(0.3),
         "ESCALATION LANE", size=10, bold=True, color=NAVY)
text_box(s, Inches(0.7), esc_y + Inches(0.32),
         SW - Inches(1.4), Inches(0.4),
         "Cross-VS conflicts → Customer Journey Forum  ·  Cross-portfolio conflicts → VP Portfolio Sync  "
         "·  Every Expedite reviewed at the next Gate 2",
         size=9.5, color=DARK)

# Takeaway
tk_y = Inches(5.6)
round_rect(s, Inches(0.5), tk_y, SW - Inches(1), Inches(1.25),
           BLUE_PALE, line=NAVY, line_w=1.5, radius=0.05)
text_box(s, Inches(0.7), tk_y + Inches(0.1),
         SW - Inches(1.4), Inches(0.3),
         "TAKEAWAY", size=10, bold=True, color=NAVY)
text_box(s, Inches(0.7), tk_y + Inches(0.4),
         SW - Inches(1.4), Inches(0.85),
         "Two cadences are sufficient — weekly clears noise, monthly steers value, quarterly aligns the "
         "portfolio. Replaces today's mix of ad-hoc forums where 'regulatory always wins' bypassed the queue.",
         size=10.5, italic=True, color=DARK)
footer_bar(s)


# ─── 7e · PERFORMANCE (KPI dashboard mockup) ───────────────────────────
s = blank()
header_bar(s, "Performance · seven KPIs make demand health observable in one dashboard")
stage1_kicker(s, RED, "PERFORMANCE",
              "How success is measured across the demand funnel.",
              "Resolves Issues 02 + 06")

kpis = [
    ("DEMAND VOLUME",
     "Submissions / week", "by VS · CoS · OPCO", BLUE,
     "Target", "trend visible"),
    ("DoR PASS RATE",
     "% accepted at first submit", "Target ≥ 70 %", GREEN,
     "Target", "≥ 70 %"),
    ("TRIAGE LEAD TIME",
     "Submit → Gate 1 verdict", "Target ≤ 5 business days", AMBER,
     "Target", "≤ 5 d (P50)"),
    ("CHANNEL ADOPTION",
     "% via single front door", "Target 90 % by month 6", PURPLE,
     "Target", "≥ 90 %"),
    ("EXPEDITE SHARE",
     "% of throughput", "Cap 20 %", RED,
     "Cap", "≤ 20 %"),
    ("DUPLICATION RATE",
     "% submissions matched", "Trend: down", TEAL,
     "Trend", "↓"),
    ("STATUS VISIBILITY",
     "% with live status", "Target 100 %", INDIGO,
     "Target", "100 %"),
]
# 7 tiles in 4 + 3 layout
kp_y = Inches(1.85)
kp_h = Inches(2.0)
top_n = 4
top_gap = Inches(0.15)
top_w = (SW - Inches(1.0) - top_gap * (top_n - 1)) / top_n
for i in range(top_n):
    name, sub, target_text, color, lbl, val = kpis[i]
    x = Inches(0.5) + i * (top_w + top_gap)
    round_rect(s, x, kp_y, top_w, kp_h, WHITE, line=color, line_w=2, radius=0.05)
    rect(s, x, kp_y, top_w, Inches(0.45), color)
    text_box(s, x + Inches(0.1), kp_y + Inches(0.04),
             top_w - Inches(0.2), Inches(0.4),
             name, size=10, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.15), kp_y + Inches(0.55),
             top_w - Inches(0.3), Inches(0.4),
             sub, size=10, italic=True, color=DARK)
    # Big metric
    text_box(s, x + Inches(0.15), kp_y + Inches(1.0),
             top_w - Inches(0.3), Inches(0.5),
             val, size=22, bold=True, color=color,
             anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.15), kp_y + Inches(1.5),
             top_w - Inches(0.3), Inches(0.4),
             target_text, size=8.5, italic=True, color=GREY,
             anchor=MSO_ANCHOR.MIDDLE)

# Bottom row · 3 tiles
bot_y = Inches(4.05)
bot_n = 3
bot_w = (SW - Inches(1.0) - top_gap * (bot_n - 1)) / bot_n
for i in range(bot_n):
    idx = top_n + i
    name, sub, target_text, color, lbl, val = kpis[idx]
    x = Inches(0.5) + i * (bot_w + top_gap)
    round_rect(s, x, bot_y, bot_w, kp_h, WHITE, line=color, line_w=2, radius=0.05)
    rect(s, x, bot_y, bot_w, Inches(0.45), color)
    text_box(s, x + Inches(0.1), bot_y + Inches(0.04),
             bot_w - Inches(0.2), Inches(0.4),
             name, size=10, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.15), bot_y + Inches(0.55),
             bot_w - Inches(0.3), Inches(0.4),
             sub, size=10, italic=True, color=DARK)
    text_box(s, x + Inches(0.15), bot_y + Inches(1.0),
             bot_w - Inches(0.3), Inches(0.5),
             val, size=22, bold=True, color=color,
             anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.15), bot_y + Inches(1.5),
             bot_w - Inches(0.3), Inches(0.4),
             target_text, size=8.5, italic=True, color=GREY,
             anchor=MSO_ANCHOR.MIDDLE)

# Takeaway
tk_y = Inches(6.25)
round_rect(s, Inches(0.5), tk_y, SW - Inches(1), Inches(0.65),
           BLUE_PALE, line=NAVY, line_w=1.5, radius=0.05)
text_box(s, Inches(0.7), tk_y + Inches(0.04),
         SW - Inches(1.4), Inches(0.3),
         "TAKEAWAY", size=10, bold=True, color=NAVY)
text_box(s, Inches(0.7), tk_y + Inches(0.27),
         SW - Inches(1.4), Inches(0.4),
         "These seven KPIs are the contract — if any one drifts, the demand funnel is leaking. "
         "Live dashboard means leakage is visible the next day, not the next quarter.",
         size=10, italic=True, color=DARK)
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
# SLIDE 8 — CAPABILITIES (14 items) — UPDATED Capability #1 with Jira+IT Now
# ════════════════════════════════════════════════════════════════════════
capabilities = [
    ("Demand intake channels",
     "One Jira request type per Value Stream · size captured as a field (S · M · L · XL). "
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
# SLIDE 9 — ROLES (4)
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
# SLIDE 10 — TOOLS (5)
# ════════════════════════════════════════════════════════════════════════
tools = [
    ("Demand capture",
     "One Jira request type per Value Stream Team — size captured as a field. "
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
# SLIDE 11 — GOVERNANCE (6)
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
# SLIDE 12 — PERFORMANCE (7)
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
# SLIDE 13 — Roles & Responsibilities (with names)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Roles & Responsibilities · current people and proposed roles")

# Top: existing leadership row
ex_y = Inches(1.3)
ex_h = Inches(1.95)
text_box(s, Inches(0.5), Inches(0.95), Inches(8), Inches(0.3),
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
new_y = Inches(3.7)
new_h = Inches(1.95)
text_box(s, Inches(0.5), Inches(3.35), Inches(8), Inches(0.3),
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
ch_y = Inches(6.15)
ch_h = Inches(0.95)
text_box(s, Inches(0.5), Inches(5.8), Inches(10), Inches(0.3),
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
# SLIDE 12b — Value Stream = Squad · two-level work model (Initiative + Story)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Value Stream = Squad · two-level work model with Epic as exception")

# Top — VS=Squad explanation
top_y = Inches(1.1)
top_h = Inches(1.4)
round_rect(s, Inches(0.5), top_y, SW - Inches(1), top_h,
           BLUE_PALE, line=BLUE, line_w=2, radius=0.05)
text_box(s, Inches(0.7), top_y + Inches(0.1),
         SW - Inches(1.4), Inches(0.32),
         "VALUE STREAM = SQUAD",
         size=12, bold=True, color=NAVY)
text_box(s, Inches(0.7), top_y + Inches(0.42),
         SW - Inches(1.4), top_h - Inches(0.5),
         "Each Value Stream Crew is a permanent, multi-functional Squad — single-product ownership, "
         "stable composition, ≈3–5 Avangrid FTE (Product Manager · Tech Lead · BA · SMEs) plus a vendor "
         "extension that does the actual delivery. Eight Value Streams = eight Squads. The Squad owns "
         "its product backlog and operates one Team Kanban.",
         size=11, color=DARK, italic=True)

# Bottom — 2-level work model with Epic exception
m_y = Inches(2.7)
m_h = Inches(4.3)

# Left panel — default flow (Initiative → Story)
lp_w = Inches(6.2)
round_rect(s, Inches(0.5), m_y, lp_w, m_h, WHITE, line=GREEN, line_w=2, radius=0.05)
rect(s, Inches(0.5), m_y, lp_w, Inches(0.55), GREEN)
text_box(s, Inches(0.5), m_y + Inches(0.05), lp_w, Inches(0.45),
         "DEFAULT FLOW  ·  Initiative → Story",
         size=13, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Initiative box
ib_y = m_y + Inches(0.8)
round_rect(s, Inches(0.8), ib_y, lp_w - Inches(0.6), Inches(1.0),
           NAVY, radius=0.05)
text_box(s, Inches(0.95), ib_y + Inches(0.08),
         lp_w - Inches(0.9), Inches(0.3),
         "INITIATIVE", size=12, bold=True, color=ACCENT)
text_box(s, Inches(0.95), ib_y + Inches(0.38),
         lp_w - Inches(0.9), Inches(0.55),
         "Outcome at portfolio level · sized for ≤ 3 months · sponsor-approved · "
         "lives on the Portfolio Board",
         size=10, color=WHITE, italic=True)

# Down arrow
text_box(s, Inches(3.2), ib_y + Inches(1.05),
         Inches(0.6), Inches(0.4),
         "▼", size=20, bold=True, color=GREEN,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Stories row
sr_y = ib_y + Inches(1.5)
nstories = 4
sw = (lp_w - Inches(0.7) - Inches(0.1) * (nstories - 1)) / nstories
for i in range(nstories):
    sx = Inches(0.8) + i * (sw + Inches(0.1))
    round_rect(s, sx, sr_y, sw, Inches(0.7), GREEN_LIGHT,
               line=GREEN, line_w=1.5, radius=0.05)
    text_box(s, sx, sr_y, sw, Inches(0.35),
             "STORY", size=9, bold=True, color=GREEN,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, sx, sr_y + Inches(0.32), sw, Inches(0.35),
             f"#{i+1}", size=10, bold=True, color=DARK,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

text_box(s, Inches(0.8), sr_y + Inches(0.85),
         lp_w - Inches(0.6), Inches(0.5),
         "Stories live on the Team Kanban (the Squad's board) — pulled by capacity, "
         "small enough to finish in days/weeks.",
         size=9.5, italic=True, color=DARK,
         align=PP_ALIGN.CENTER)

# Right panel — exception (Initiative > 3 months → Epic)
rp_x = Inches(0.5) + lp_w + Inches(0.15)
rp_w = SW - Inches(0.5) - rp_x
round_rect(s, rp_x, m_y, rp_w, m_h, WHITE, line=AMBER, line_w=2, radius=0.05)
rect(s, rp_x, m_y, rp_w, Inches(0.55), AMBER)
text_box(s, rp_x, m_y + Inches(0.05), rp_w, Inches(0.45),
         "EXCEPTION  ·  Initiative > 3 months → use Epics",
         size=13, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Initiative
ib2_y = m_y + Inches(0.75)
round_rect(s, rp_x + Inches(0.2), ib2_y, rp_w - Inches(0.4), Inches(0.55),
           NAVY, radius=0.05)
text_box(s, rp_x + Inches(0.2), ib2_y, rp_w - Inches(0.4), Inches(0.55),
         "INITIATIVE  (large · multi-quarter)",
         size=10.5, bold=True, color=ACCENT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Down arrow
text_box(s, rp_x + rp_w/2 - Inches(0.2),
         ib2_y + Inches(0.6),
         Inches(0.4), Inches(0.3),
         "▼", size=16, bold=True, color=AMBER,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Epics
ep_y = ib2_y + Inches(0.95)
nepics = 3
ew = (rp_w - Inches(0.4) - Inches(0.1) * (nepics - 1)) / nepics
for i in range(nepics):
    ex = rp_x + Inches(0.2) + i * (ew + Inches(0.1))
    round_rect(s, ex, ep_y, ew, Inches(0.5), AMBER_LIGHT,
               line=AMBER, line_w=1.5, radius=0.05)
    text_box(s, ex, ep_y, ew, Inches(0.5),
             f"EPIC #{i+1}", size=9.5, bold=True, color=AMBER,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Down arrow
text_box(s, rp_x + rp_w/2 - Inches(0.2),
         ep_y + Inches(0.55),
         Inches(0.4), Inches(0.3),
         "▼", size=14, bold=True, color=GREEN,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Stories below epics
st2_y = ep_y + Inches(0.95)
nst2 = 6
sw2 = (rp_w - Inches(0.4) - Inches(0.06) * (nst2 - 1)) / nst2
for i in range(nst2):
    sx = rp_x + Inches(0.2) + i * (sw2 + Inches(0.06))
    round_rect(s, sx, st2_y, sw2, Inches(0.4), GREEN_LIGHT,
               line=GREEN, line_w=1, radius=0.05)
    text_box(s, sx, st2_y, sw2, Inches(0.4),
             "S", size=10, bold=True, color=GREEN,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

text_box(s, rp_x + Inches(0.2), st2_y + Inches(0.55),
         rp_w - Inches(0.4), Inches(0.7),
         "Epics slice an oversized Initiative into deliverable chunks. "
         "Stories under each Epic continue to live on the Team Kanban.",
         size=9, italic=True, color=DARK,
         align=PP_ALIGN.CENTER)

footer_bar(s)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 12c — Stage 1 · Branch specifics (CS Apps vs IT Business Apps)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Stage 1 specifics · how the framework lands differently in CS Apps vs IT Business Apps")

# Two badges at top
badge_y = Inches(1.05)
badge_h = Inches(0.45)
badge_w = (SW - Inches(1.0) - Inches(0.2)) / 2

# CS Apps badge (green)
round_rect(s, Inches(0.5), badge_y, badge_w, badge_h, GREEN, radius=0.2)
text_box(s, Inches(0.5), badge_y, badge_w, badge_h,
         "CUSTOMER SERVICES APPS  ·  Kumar Jaladanki  ·  3 Value Streams",
         size=12, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# IT Business Apps badge (blue)
round_rect(s, Inches(0.5) + badge_w + Inches(0.2), badge_y,
           badge_w, badge_h, BLUE, radius=0.2)
text_box(s, Inches(0.5) + badge_w + Inches(0.2), badge_y,
         badge_w, badge_h,
         "IT BUSINESS APPS  ·  Carlos Garbin  ·  5 Value Streams",
         size=12, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Comparison rows — 5 dimensions
dims = [
    ("INTAKE / DEMAND SOURCES",
     "Customer Operations · OPCO call centers · Marketing · Customer Service GenAI · Heather LaJoie's Product Management.",
     "G&E Operations · Engineering · Storm Ops · System Operations control room · Iberdrola corporate · regulatory affairs."),
    ("APPLICATIONS",
     "SAP CRM · IS-U FICA · Smart Metering / Itron · Medallia · IVR · Contact Center.",
     "GIS (ArcGIS · Smallworld) · OMS · SCADA · EMS · Maintenance (16 apps) · Global SAP."),
    ("VENDOR EXTENSION",
     "Lemur-Salesforce (CRM) · Itron / Smart Metering · Accenture (FICA) · IVR partner · Medallia partner.",
     "HCL (GIS) · Esri partners · GE / Siemens / ABB (SCADA / EMS / OMS) · SWF · Infosys · Iberdrola IT."),
    ("CLASS-OF-SERVICE MIX  [estimate]",
     "Standard ~55 % · Expedite ~20 % (customer impact) · Fixed-date ~15 % (rate cases) · Intangible ~10 %.",
     "Standard ~45 % · Fixed-date ~30 % (regulatory · storm · capital) · Expedite ~15 % · Intangible ~10 %."),
    ("RUN / GROW / TRANSFORM MIX  [estimate]",
     "Run ~50 % · Grow ~35 % · Transform ~15 % — customer-experience and digital channel weight Grow.",
     "Run ~60 % · Grow ~25 % · Transform ~15 % — operational reliability and asset health weight Run."),
]

dim_y = Inches(1.65)
dim_h = (SH - dim_y - Inches(0.5)) / len(dims)
for i, (dim_name, cs_text, it_text) in enumerate(dims):
    y = dim_y + i * dim_h
    # Row separator
    if i > 0:
        rect(s, Inches(0.5), y, SW - Inches(1), Inches(0.005), GREY)
    # Dimension label (left column)
    text_box(s, Inches(0.5), y + Inches(0.05),
             Inches(2.4), dim_h - Inches(0.1),
             dim_name, size=10, bold=True, color=NAVY,
             anchor=MSO_ANCHOR.MIDDLE)
    # CS Apps cell (left)
    cs_x = Inches(3.0)
    cs_w = (SW - Inches(0.5) - cs_x - Inches(0.1)) / 2
    rect(s, cs_x - Inches(0.05), y + Inches(0.08),
         Inches(0.05), dim_h - Inches(0.16), GREEN)
    text_box(s, cs_x + Inches(0.05), y + Inches(0.05),
             cs_w - Inches(0.1), dim_h - Inches(0.1),
             cs_text, size=9, color=DARK,
             anchor=MSO_ANCHOR.MIDDLE)
    # IT Apps cell (right)
    it_x = cs_x + cs_w + Inches(0.1)
    rect(s, it_x - Inches(0.05), y + Inches(0.08),
         Inches(0.05), dim_h - Inches(0.16), BLUE)
    text_box(s, it_x + Inches(0.05), y + Inches(0.05),
             cs_w - Inches(0.1), dim_h - Inches(0.1),
             it_text, size=9, color=DARK,
             anchor=MSO_ANCHOR.MIDDLE)

footer_bar(s)



# ════════════════════════════════════════════════════════════════════════
# SLIDES 12d / 12e — Stage 1 INTAKE detail per branch (CS Apps · IT Business Apps)
# Sources · channels · classification · prioritization — branch-specific
# ════════════════════════════════════════════════════════════════════════
def branch_intake_slide(branch_label, branch_color, branch_owner, vs_count_label,
                        sources, channels_lines, classification_lines, prioritization_lines):
    s = blank()
    header_bar(s, f"Stage 1 intake detail · {branch_label}",
               color=branch_color)

    # Branch badge top — large pill
    bb_y = Inches(1.05)
    round_rect(s, Inches(0.5), bb_y, SW - Inches(1), Inches(0.5),
               branch_color, radius=0.2)
    text_box(s, Inches(0.5), bb_y, SW - Inches(1), Inches(0.5),
             f"{branch_label.upper()}  ·  {branch_owner}  ·  {vs_count_label}",
             size=12, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Four panels (2 × 2)
    p_y = Inches(1.7)
    p_h = (SH - p_y - Inches(0.5)) / 2
    p_w = (SW - Inches(1.15)) / 2
    p_gap = Inches(0.15)

    panels = [
        ("DEMAND SOURCES — who sends the work", branch_color, sources),
        ("CHANNELS BY TYPE — how it arrives", BLUE,
         channels_lines),
        ("CLASSIFICATION — branch-specific rules", AMBER, classification_lines),
        ("PRIORITIZATION — Pain Score weighting", PURPLE, prioritization_lines),
    ]

    for i, (title, color, items) in enumerate(panels):
        col = i % 2
        row = i // 2
        x = Inches(0.5) + col * (p_w + p_gap)
        y = p_y + row * (p_h + Inches(0.15))
        round_rect(s, x, y, p_w, p_h, WHITE, line=color, line_w=2, radius=0.05)
        rect(s, x, y, p_w, Inches(0.45), color)
        text_box(s, x + Inches(0.15), y + Inches(0.04),
                 p_w - Inches(0.3), Inches(0.4),
                 title, size=11, bold=True, color=WHITE,
                 anchor=MSO_ANCHOR.MIDDLE)
        add_lines(s, x + Inches(0.25), y + Inches(0.55),
                  p_w - Inches(0.5), p_h - Inches(0.7),
                  ["•  " + it for it in items], size=9.5,
                  color=DARK, line_spacing=1.35)

    footer_bar(s)
    return s


# ─── 12d · CS Apps intake detail ──────────────────────────────────────
branch_intake_slide(
    "Customer Services Apps",
    GREEN,
    "Kumar Jaladanki",
    "3 Value Streams — M2C-SL · M2C-RO · CCC",
    sources=[
        "Customer Operations (Heather LaJoie · Product Management)",
        "OPCO call centers (UIL · CMP · NY) — escalations and trends",
        "Marketing & Communications — campaign and digital channel demand",
        "Customer Service GenAI initiative — pipeline of new use cases",
        "Smart Metering program — CT · ME · NY rollout drivers",
        "Regulatory affairs — rate cases · tariff changes · tax updates"],
    channels_lines=[
        "ONE Jira request type per Value Stream (M2C-SL · M2C-RO · CCC)",
        "Mandatory fields: VS · type · size · CoS · OPCO · sponsor · Pain Score",
        "IT Now — restricted to CCC / Smart Metering INCIDENTS only",
        "Email · Teams · Word — DECOMMISSIONED for new demand",
        "Heather LaJoie's Product Management — feeds CCC backlog directly",
        "[GAP] confirm whether Marketing demand routes via M2C-RO or CCC"],
    classification_lines=[
        "Run / Grow / Transform — labeled at intake (default ~50 / 35 / 15 %)",
        "Class of Service — Standard · Expedite (customer-impact) · Fixed-date · Intangible",
        "Customer impact tier P1–P4 — customer-facing severity drives Expedite",
        "CSAT-sensitive items flagged separately (Medallia signal)",
        "Regulatory tariff items always Fixed-date with rate-case calendar lock",
        "Smart Metering CT / ME mandates — pre-classified as Fixed-date"],
    prioritization_lines=[
        "Pain Score H(x) = (BV·w₁ + UX·w₂ + Safety·w₃) / Effort",
        "CS weighting — UX heavy: w₁=0.35 · w₂=0.40 · w₃=0.25  [proposal]",
        "Sponsor sign-off mandatory at intake (Heather / Marketing / Ops)",
        "Backlog ranked monthly at Gate 2 by the VS Product Manager",
        "Cross-VS conflicts (e.g. CRM ↔ Smart Metering) resolved at Customer Journey Forum",
        "Capacity reserved 10–15 % per quarter for Intangible (tech debt)"]
)


# ─── 12e · IT Business Apps intake detail ─────────────────────────────
branch_intake_slide(
    "IT Business Apps",
    BLUE,
    "Carlos Garbin",
    "5 Value Streams — Field Ops · E&A · NOC · C&R · IBD",
    sources=[
        "Gas & Electric Operations (Bruno Moreto · Andrew McGibbon Champion)",
        "Engineering · Asset planners · capital project sponsors",
        "Storm Operations & Emergency Management",
        "System Operations control room (24/7) — control engineers",
        "Iberdrola corporate (Bilbao) — Vanessa Colombo intake",
        "Corporate functions — Finance · HR · Procurement · Legal · Regulatory",
        "Sarath Garikapati's PJM pool — today distributed, will centralize"],
    channels_lines=[
        "ONE Jira request type per Value Stream (Field Ops · E&A · NOC · C&R · IBD)",
        "Mandatory fields: VS · type · size · CoS · OPCO · sponsor · Pain Score",
        "IT Now — restricted to OT / SCADA / OMS INCIDENTS only",
        "PPM Now (Smart Metering) → migrated into VS-level Jira backlog",
        "Email · SharePoint · Word — DECOMMISSIONED for new demand",
        "Iberdrola Global SAP release calendar — cross-routed via IBD VS"],
    classification_lines=[
        "Run / Grow / Transform — labeled at intake (default ~60 / 25 / 15 %)",
        "Class of Service — Standard · Fixed-date (regulatory · storm) · Expedite · Intangible",
        "Safety-sensitive items (NOC · Field Ops) auto-routed Expedite",
        "Storm response triggers Fixed-date with reliability deadline lock",
        "Capital plan items synced with Avangrid annual capital cycle (rate base)",
        "Regulatory submissions (NY PSC · CT PURA · ME PUC) pre-tagged Fixed-date"],
    prioritization_lines=[
        "Pain Score H(x) = (BV·w₁ + UX·w₂ + Safety·w₃) / Effort",
        "IT weighting — Safety heavy: w₁=0.40 · w₂=0.20 · w₃=0.40  [proposal]",
        "Sponsor sign-off mandatory at intake (G&E · Engineering · Iberdrola)",
        "Backlog ranked monthly at Gate 2 by the VS Product Manager",
        "Cross-VS conflicts (e.g. NOC ↔ Field Ops storm) resolved at VP Portfolio Sync",
        "Reliability-impacting items override capacity caps (logged exception)"]
)



# ════════════════════════════════════════════════════════════════════════
# SLIDES 13a–13h — Value Stream deep-dives (one per VS · 8 slides)
# Each slide: outcome · supporting apps · demand sources · team & vendors
# ════════════════════════════════════════════════════════════════════════
def vs_deep_dive(vs_id, vs_name, branch, header_color, outcome,
                 apps, demand_sources, core_team, vendors, mix_label):
    s = blank()
    header_bar(s, f"Value Stream · {vs_id} {vs_name} — {outcome[0]}")

    # Top breadcrumb / branch tag
    bc_y = Inches(1.0)
    round_rect(s, Inches(0.5), bc_y, Inches(3.4), Inches(0.4),
               header_color, radius=0.2)
    text_box(s, Inches(0.5), bc_y, Inches(3.4), Inches(0.4),
             branch, size=10, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Mix label on the right
    mx_x = SW - Inches(4.4)
    round_rect(s, mx_x, bc_y, Inches(3.9), Inches(0.4),
               LIGHT, line=header_color, line_w=1, radius=0.2)
    text_box(s, mx_x, bc_y, Inches(3.9), Inches(0.4),
             "Run · Grow · Transform mix  ·  " + mix_label,
             size=9.5, bold=True, color=header_color,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Outcome banner
    out_y = Inches(1.5)
    round_rect(s, Inches(0.5), out_y, SW - Inches(1), Inches(0.65),
               BLUE_LIGHT, line=header_color, line_w=1.5, radius=0.05)
    text_box(s, Inches(0.7), out_y + Inches(0.06),
             SW - Inches(1.4), Inches(0.25),
             "OUTCOME", size=9, bold=True, color=NAVY)
    text_box(s, Inches(0.7), out_y + Inches(0.27),
             SW - Inches(1.4), Inches(0.4),
             outcome[1], size=11, italic=True, color=DARK)

    # 4 quadrants
    q_y = Inches(2.3)
    q_h = Inches(4.3)
    q_w = (SW - Inches(1.15)) / 2
    q_gap = Inches(0.15)

    panels = [
        ("APPLICATIONS SUPPORTED", header_color,
         apps, "(KB-confirmed unless flagged [GAP])"),
        ("DEMAND SOURCES — who sends the work", BLUE,
         demand_sources, "Routed via the Jira request type for this Value Stream"),
        ("AVANGRID CORE TEAM", GREEN,
         core_team, "≈ 3–5 FTE per Value Stream"),
        ("VENDOR EXTENSION (delivery)", AMBER,
         vendors, "Avangrid = product owner · vendor = delivery muscle"),
    ]
    for i, (title, color, items, note) in enumerate(panels):
        col = i % 2
        row = i // 2
        x = Inches(0.5) + col * (q_w + q_gap)
        y = q_y + row * (q_h / 2 + Inches(0.1))
        h = q_h / 2
        round_rect(s, x, y, q_w, h, WHITE, line=color, line_w=1.5, radius=0.05)
        rect(s, x, y, q_w, Inches(0.4), color)
        text_box(s, x + Inches(0.15), y + Inches(0.03),
                 q_w - Inches(0.3), Inches(0.34),
                 title, size=10.5, bold=True, color=WHITE,
                 anchor=MSO_ANCHOR.MIDDLE)
        add_lines(s, x + Inches(0.2), y + Inches(0.5),
                  q_w - Inches(0.4), h - Inches(0.85),
                  ["•  " + it for it in items], size=9.5,
                  color=DARK, line_spacing=1.3)
        text_box(s, x + Inches(0.15), y + h - Inches(0.28),
                 q_w - Inches(0.3), Inches(0.25),
                 note, size=8, italic=True, color=GREY)

    footer_bar(s)
    return s


vs_specs = [
    ("M2C-SL", "Meter-to-Cash · Service Lifecycle",
     "CS Apps · Kumar Jaladanki · Vincent Lam (sub-team lead)",
     BLUE,
     ("Smart-meter to invoice with zero data loss",
      "Capture every meter read accurately · feed billing on time · resolve "
      "device-to-data exceptions before they reach the customer."),
     ["Itron OpenWay (AMI head-end)",
      "Meter Data Management (MDM)",
      "SAP IS-U Device Management",
      "Smart Metering platform",
      "Field meter read interfaces"],
     ["Smart Metering program (CT · ME · NY rollout)",
      "Field Operations exceptions and escalations",
      "Regulatory drivers (CT PURA · ME PUC · NY PSC mandates)",
      "AMI vendor release schedule"],
     ["Product Manager (NEW)",
      "Tech Lead — Vincent Lam / Iram Chowdhury (Champion)",
      "Business Analyst",
      "SMEs from Smart Metering program"],
     ["Smart Metering vendor (stable team — Option A model)",
      "Itron managed services",
      "SAP partner for IS-U Device"],
     "Run 60% · Grow 30% · Transform 10% [estimate]"),

    ("M2C-RO", "Meter-to-Cash · Revenue Operations",
     "CS Apps · Kumar Jaladanki · Ivelisse Velazquez (sub-team lead)",
     BLUE,
     ("Bill once, bill right, collect on time",
      "Standardize billing, FICA and collection processes across OPCOs · respond "
      "to rate cases without ad-hoc workarounds."),
     ["SAP IS-U FICA",
      "SAP CRM Billing",
      "Collection systems",
      "Rate-case calculation tooling",
      "Tax engine integrations"],
     ["Finance & Billing Operations (UIL · CMP · NY)",
      "Regulatory (rate-case filings, tariff changes)",
      "Tax updates and federal/state mandates",
      "Customer escalations from CCC"],
     ["Product Manager (NEW)",
      "Tech Lead",
      "Business Analyst — billing & FICA SME",
      "Stephanie (SAP Champion · cross-link)"],
     ["Accenture (FICA / IS-U delivery)",
      "SAP partners",
      "[GAP] confirm primary FICA vendor with Vitor"],
     "Run 50% · Grow 35% · Transform 15% [estimate]"),

    ("CCC", "Customer Care & Contact Center",
     "CS Apps · Kumar Jaladanki · CRM/CCC sub-team (vacancy)",
     BLUE,
     ("Resolve customer interactions on first contact",
      "Unify CRM, Contact Center and digital self-service so every customer "
      "interaction lands in a single record with a clear next step."),
     ["SAP CRM",
      "Contact Center platform (CCC)",
      "Medallia (CSAT and journey signal)",
      "IVR Solution",
      "Digital self-service / chatbot"],
     ["Customer Operations (Heather LaJoie's Product Mgmt)",
      "Marketing & Communications (campaign demand)",
      "OPCO call centers (UIL · CMP · NY)",
      "Customer Service GenAI initiative pipeline"],
     ["Product Manager (NEW)",
      "Tech Lead — vacancy (CRM & Contact Center Ops)",
      "Business Analyst — CX SME",
      "[GAP] confirm CCC sub-team lead post-vacancy"],
     ["Lemur-Salesforce (CRM delivery)",
      "IVR vendor",
      "Medallia partner",
      "Digital channel agency"],
     "Run 45% · Grow 35% · Transform 20% [estimate]"),

    ("FO", "Field Operations",
     "IT Apps · Carlos Garbin · George Porto (sub-team lead)",
     GREEN,
     ("Field crews dispatched, tracked, and closed in one workflow",
      "GIS, workforce, outage and storm response capabilities operate as one "
      "integrated field operations stack."),
     ["ESRI ArcGIS (network model)",
      "Smallworld GIS",
      "OSG — Outage / Storm Management",
      "Mobile workforce apps",
      "Crew dispatch tooling"],
     ["G&E Operations (Bruno Moreto's team)",
      "Field workforce and dispatch supervisors",
      "Storm response & emergency management",
      "Iberdrola G&E global program"],
     ["Product Manager (NEW)",
      "Tech Lead — George Porto (GIS Champion)",
      "Business Analyst",
      "Field SMEs (rotating)"],
     ["HCL (GIS support and maintenance)",
      "Esri partners",
      "OSG vendor"],
     "Run 65% · Grow 25% · Transform 10% [estimate]"),

    ("EAM", "Engineering & Asset Management",
     "IT Apps · Carlos Garbin · cross sub-teams",
     GREEN,
     ("Asset data trusted enough to drive capital and maintenance decisions",
      "Reliable engineering and maintenance data feeds the rate base, capital "
      "plan, and reliability commitments."),
     ["Maintenance management apps (16 apps · largest cluster)",
      "Asset Management systems",
      "Engineering design tools",
      "Reliability analytics"],
     ["G&E Engineering",
      "Asset planners and capital project sponsors",
      "Maintenance and reliability engineers",
      "Capital plan governance"],
     ["Product Manager (NEW)",
      "Tech Lead",
      "Business Analyst",
      "Engineering & Maintenance SMEs"],
     ["SWF (delivery partner)",
      "OEMs (asset-specific)",
      "[GAP] consolidate vendor map for the 16-app maintenance cluster"],
     "Run 55% · Grow 30% · Transform 15% [estimate]"),

    ("NOC", "Network Operations & Control  [NEW STREAM]",
     "IT Apps · Carlos Garbin · surfaced from IT Networks Application Assessment",
     GREEN,
     ("Real-time grid visibility and control with no blind spots",
      "SCADA, EMS, OMS and emerging DERMS/ADMS capabilities operate as one "
      "control-room stack — proposed as a NEW Value Stream after the 46-app "
      "assessment surfaced this gap."),
     ["SCADA",
      "EMS — Energy Management System",
      "OMS — Outage Management System",
      "DERMS / ADMS (emerging)",
      "(5 apps total — IT Networks Assessment Slide 10/23)"],
     ["System Operations control room (24/7)",
      "Control engineers and reliability operators",
      "Storm Operations and Emergency Mgmt",
      "Reliability and grid-modernization initiatives"],
     ["Product Manager (NEW)",
      "Tech Lead — control-room platform SME",
      "Business Analyst",
      "[GAP] confirm sub-team owner with Carlos and Travis"],
     ["GE (SCADA / EMS partner)",
      "Siemens / ABB (OEM)",
      "[GAP] confirm primary OMS / DERMS vendor"],
     "Run 70% · Grow 20% · Transform 10% [estimate · NEW stream]"),

    ("C&R", "Corporate & Regulatory",
     "IT Apps · Carlos Garbin · Sarath Garikapati (sub-team lead)",
     GREEN,
     ("Corporate and regulatory demand routed once, fulfilled predictably",
      "Finance, HR, Procurement and Regulatory IT requests share one product "
      "team — no more 5-6 PJMs absorbing demand reactively."),
     ["SAP Finance",
      "SAP HR",
      "SAP Procurement",
      "Regulatory reporting tooling",
      "Document management cross-cuts"],
     ["Corporate functions (Finance · HR · Procurement · Legal)",
      "Regulatory affairs across OPCOs",
      "Sarath's 5-6 PJM intake (today distributed)",
      "Iberdrola corporate compliance asks"],
     ["Product Manager (NEW)",
      "Tech Lead — Amit Patel (Corporate Champion)",
      "Business Analyst",
      "PJM pool from Sarath (shared resource)"],
     ["Accenture",
      "Infosys",
      "SAP partners (Finance / HR / Procurement)"],
     "Run 50% · Grow 35% · Transform 15% [estimate]"),

    ("IBD", "Iberdrola Corporate Services",
     "Cross-branch · Vanessa Colombo (Global SAP / FI / PO / Budget)",
     GREEN,
     ("Avangrid in lock-step with Iberdrola's global SAP releases",
      "Global SAP, FI and budget integrations stay aligned to Iberdrola's "
      "release calendar without absorbing local demand variability."),
     ["Global SAP (Iberdrola template)",
      "Iberdrola FI integration",
      "PO / Budget integration",
      "Global compliance reporting"],
     ["Iberdrola corporate (Bilbao)",
      "Vanessa Colombo's existing intake",
      "Global SAP release pipeline",
      "Group-wide finance and budget cycles"],
     ["Product Manager (NEW)",
      "Tech Lead — Michael Dubois (Global SAP Champion)",
      "Business Analyst",
      "Vanessa Colombo (sub-team lead · cross-link)"],
     ["Global SAP partners (Iberdrola-managed)",
      "Iberdrola IT (intra-group dependency)"],
     "Run 70% · Grow 25% · Transform 5% [estimate · stable cycle]"),
]

for spec in vs_specs:
    vs_deep_dive(*spec)



# ════════════════════════════════════════════════════════════════════════
# SLIDE 13i — Vendor delivery model · Option A vs Option B
# Storyline: VSs need delivery muscle from the Segregated SW Delivery Base.
# Two patterns exist today — we recommend a hybrid.
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Vendor delivery model · two patterns observed today, one recommended hybrid")

text_box(s, Inches(0.5), Inches(1.05), SW - Inches(1), Inches(0.45),
         "Avangrid already operates two vendor-engagement patterns. The new operating model needs an "
         "explicit choice per Value Stream — the right pattern depends on whether work is continuous or episodic.",
         size=11, italic=True, color=BLUE, bold=True)

# Two side-by-side panels for A and B
opt_y = Inches(1.7)
opt_h = Inches(4.4)
opt_w = (SW - Inches(1.15)) / 2

# Option A
ax = Inches(0.5)
round_rect(s, ax, opt_y, opt_w, opt_h, WHITE, line=GREEN, line_w=2, radius=0.05)
rect(s, ax, opt_y, opt_w, Inches(0.55), GREEN)
text_box(s, ax, opt_y + Inches(0.05), opt_w, Inches(0.45),
         "OPTION A  ·  Stable vendor team per Crew",
         size=13, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text_box(s, ax + Inches(0.2), opt_y + Inches(0.7),
         opt_w - Inches(0.4), Inches(0.4),
         "Today's example  ·  Smart Metering",
         size=10, italic=True, color=GREEN, bold=True)
add_lines(s, ax + Inches(0.2), opt_y + Inches(1.05),
          opt_w - Inches(0.4), Inches(2.4),
          ["•  Vendor team is permanently assigned to one Value Stream Crew",
           "•  Avangrid is the product owner — vendor is the delivery muscle",
           "•  Continuity of context, domain knowledge, and tooling",
           "•  Faster ramp on new work — no re-explanation required",
           "•  Joint roadmap reviewed quarterly with the vendor",
           "•  Predictable cost envelope per Crew"],
          size=10, color=DARK, line_spacing=1.35)
# Use case
uc_y = opt_y + opt_h - Inches(1.2)
rect(s, ax + Inches(0.2), uc_y, opt_w - Inches(0.4), Inches(0.005), GREEN)
text_box(s, ax + Inches(0.2), uc_y + Inches(0.05),
         opt_w - Inches(0.4), Inches(0.3),
         "BEST FIT", size=10, bold=True, color=GREEN)
text_box(s, ax + Inches(0.2), uc_y + Inches(0.32),
         opt_w - Inches(0.4), Inches(0.85),
         "Continuous, domain-heavy work — M2C-SL · Field Operations · NOC · Engineering & Asset · Iberdrola Corp.",
         size=10, color=DARK, italic=True)

# Option B
bx = ax + opt_w + Inches(0.15)
round_rect(s, bx, opt_y, opt_w, opt_h, WHITE, line=AMBER, line_w=2, radius=0.05)
rect(s, bx, opt_y, opt_w, Inches(0.55), AMBER)
text_box(s, bx, opt_y + Inches(0.05), opt_w, Inches(0.45),
         "OPTION B  ·  Vendor pool consumed ad-hoc",
         size=13, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text_box(s, bx + Inches(0.2), opt_y + Inches(0.7),
         opt_w - Inches(0.4), Inches(0.4),
         "Today's example  ·  Sarath's Corporate work",
         size=10, italic=True, color=AMBER, bold=True)
add_lines(s, bx + Inches(0.2), opt_y + Inches(1.05),
          opt_w - Inches(0.4), Inches(2.4),
          ["•  Vendor capacity drawn from a shared pool on demand",
           "•  Resources rotate across initiatives — no permanent assignment",
           "•  Maximum flexibility for episodic / cyclical work",
           "•  Lower fixed cost — but higher onboarding overhead each engagement",
           "•  No durable domain knowledge inside the vendor team",
           "•  Risk — re-explanation cost, slower start, weaker accountability"],
          size=10, color=DARK, line_spacing=1.35)
# Use case
uc_y = opt_y + opt_h - Inches(1.2)
rect(s, bx + Inches(0.2), uc_y, opt_w - Inches(0.4), Inches(0.005), AMBER)
text_box(s, bx + Inches(0.2), uc_y + Inches(0.05),
         opt_w - Inches(0.4), Inches(0.3),
         "BEST FIT", size=10, bold=True, color=AMBER)
text_box(s, bx + Inches(0.2), uc_y + Inches(0.32),
         opt_w - Inches(0.4), Inches(0.85),
         "Episodic, project-shaped work — Corporate & Regulatory rate cases · ad-hoc compliance bursts.",
         size=10, color=DARK, italic=True)

# Bottom recommendation band
rec_y = Inches(6.25)
round_rect(s, Inches(0.5), rec_y, SW - Inches(1), Inches(0.7),
           BLUE_PALE, line=NAVY, line_w=1.5, radius=0.05)
text_box(s, Inches(0.7), rec_y + Inches(0.05),
         Inches(3.0), Inches(0.55),
         "RECOMMENDATION", size=11, bold=True, color=NAVY,
         anchor=MSO_ANCHOR.MIDDLE)
text_box(s, Inches(3.5), rec_y + Inches(0.05),
         SW - Inches(4.0), Inches(0.6),
         "Hybrid — Option A as default for the seven continuous Value Streams; Option B reserved for "
         "Corporate & Regulatory's episodic load. Decision per Crew confirmed at Wave 1 Step 2.",
         size=10, italic=True, color=DARK, anchor=MSO_ANCHOR.MIDDLE)

footer_bar(s)



# ─── Section divider · 07 Demand Flow & Portfolio Operation ───
toc_slide(active_idx=6)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 14 — Forums (5 forums visual)
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
     ["ERP Steering — SAP strategic direction · Enterprise Architect view",
      "SAP Forum — Local + Iberdrola",
      "Non-SAP Forum — GIS · Mobility · Document Mgmt",
      "Platform Forum — DevOps · Security · Test automation",
      "Members: technical leads + platform owners + Enterprise Architect"]),
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
# SLIDE 15 — Portfolio Architecture (3 layers) — moved earlier (before Governance)
# This is now placed in the Stage 1 sequence so audience sees the structure
# before the Governance gates are mapped onto it.
# ════════════════════════════════════════════════════════════════════════
# (slide 15 content is now rendered earlier — no slide created here to avoid duplication)



# ════════════════════════════════════════════════════════════════════════
# SLIDE 16 — Step policies and Roles & Responsibilities per board level
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
         "TEAM KANBAN  ·  cards = STORIES  (Epic only when Initiative > 3 months)",
         size=12, bold=True, color=WHITE,
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
    ("Phase 3 — Ongoing (Kanban + WIP limits)",
     "Build · test · deploy iteratively",
     "Owner: Squad (PM + Tech Lead + BA + dev/QA)",
     "Kanban flow · daily / replenishment / service review · WIP limits per stage"),
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
# SLIDE 17 — Portfolio Board mockup (Initiative flow)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Portfolio Board · cards are Initiatives · governance gates and BRP shown in context")

text_box(s, Inches(0.5), Inches(1.0), SW - Inches(1), Inches(0.4),
         "Example  ·  IT Business Applications Portfolio Board  ·  Governance gates marked in red",
         size=12, italic=True, color=GREY, align=PP_ALIGN.CENTER)

# BRP banner above the board (quarterly cadence)
brp_y = Inches(1.4)
round_rect(s, Inches(0.5), brp_y, SW - Inches(1), Inches(0.32),
           GREEN_LIGHT, line=GREEN, line_w=1, radius=0.04)
text_box(s, Inches(0.7), brp_y + Inches(0.04),
         SW - Inches(1.4), Inches(0.28),
         "▲  BIG ROOM PLANNING — quarterly · ½ day · realigns Backlog reorder + Implementing commitments across both Portfolio Boards",
         size=9, bold=True, italic=True, color=GREEN,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

stage_y = Inches(1.85)
stage_h = Inches(4.15)
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

# Gate 1 · weekly Triage — between Funnel and Reviewing
g1_x = Inches(0.5) + 1*(col_w + gap) - Inches(0.05)
rect(s, g1_x, stage_y - Inches(0.05), Inches(0.05), stage_h + Inches(0.1), RED)
text_box(s, g1_x - Inches(1.4), stage_y - Inches(0.4),
         Inches(2.8), Inches(0.32),
         "▲ GATE 1 · TRIAGE  ·  weekly", size=9.5, bold=True, color=RED,
         align=PP_ALIGN.CENTER)

# Gate 2 · monthly Portfolio — between Analyzing and Backlog
g2_x = Inches(0.5) + 3*(col_w + gap) - Inches(0.05)
rect(s, g2_x, stage_y - Inches(0.05), Inches(0.05), stage_h + Inches(0.1), RED)
text_box(s, g2_x - Inches(1.4), stage_y - Inches(0.4),
         Inches(2.8), Inches(0.32),
         "▲ GATE 2 · PORTFOLIO  ·  monthly", size=9.5, bold=True, color=RED,
         align=PP_ALIGN.CENTER)

# S/M bypass note (right of board, near top)
bypass_y = Inches(6.0)
round_rect(s, Inches(0.5), bypass_y, SW - Inches(1), Inches(0.32),
           AMBER_LIGHT, line=AMBER, line_w=1, radius=0.04)
text_box(s, Inches(0.7), bypass_y + Inches(0.04),
         SW - Inches(1.4), Inches(0.28),
         "S / M demands BYPASS this board — they enter the Squad's Team Kanban directly. "
         "Only L / XL Initiatives flow through the Portfolio Board (Gate 2 approval required).",
         size=9, bold=True, italic=True, color=AMBER,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Legend
leg_y = Inches(6.45)
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
# SLIDE 18 — Team Kanban with Dual Track (simplified · Epic level only)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Team Kanban · five phases by two tracks · cards are Stories pulled by capacity")

text_box(s, Inches(0.5), Inches(1.0), SW - Inches(1), Inches(0.4),
         "Example  ·  Field Operations Squad  ·  five phases × dual track with WIP limits  ·  "
         "Epic appears only when an Initiative is too large for one quarter",
         size=11, italic=True, color=GREY, align=PP_ALIGN.CENTER)

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
         "DISCOVERY\nTRACK\n\nalways-on\nresearch +\nvalidation",
         size=10, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

rect(s, Inches(0.5), mid_y, left_panel, ph_y + ph_h - mid_y, AMBER)
text_box(s, Inches(0.5), mid_y, left_panel, ph_y + ph_h - mid_y,
         "DELIVERY\nTRACK\n\nKanban flow\nWIP limits\npull · validate",
         size=10, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Horizontal track separator (replaced connector with rect)
rect(s, Inches(0.5) + left_panel, mid_y - Inches(0.02),
     SW - Inches(0.5) - (Inches(0.5) + left_panel), Inches(0.04), DARK)

def epic_card(slide, x, y, w, h, title, owner, story_count, color):
    """Renders a Story card by default; if story_count > 0, renders an Epic with Story-count badge."""
    label = "STORY" if story_count == 0 else "EPIC"
    round_rect(slide, x, y, w, h, WHITE, line=color, line_w=1.5, radius=0.06)
    rect(slide, x, y, w, Inches(0.25), color)
    text_box(slide, x + Inches(0.08), y + Inches(0.02),
             w - Inches(0.5), Inches(0.22),
             label, size=8, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    if story_count > 0:
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

# Discovery row — Stories (story_count=0)
epic_card(s, x_at(0) + Inches(0.05), ph_y + Inches(0.85),
          ph_w - Inches(0.1), Inches(0.95),
          "DLI inspection — offline reqs", "exploring", 0, BLUE)
epic_card(s, x_at(1) + Inches(0.05), ph_y + Inches(0.85),
          ph_w - Inches(0.1), Inches(0.95),
          "TLI sizing with vendor", "sizing", 0, BLUE)

# Delivery row — mostly Stories, one Epic example for an Initiative > 3 months
epic_card(s, x_at(2) + Inches(0.05), mid_y + Inches(0.15),
          ph_w - Inches(0.1), Inches(0.95),
          "Storm Program v2  (Initiative > 3 mo)", "sliced into Epics", 12, AMBER)
epic_card(s, x_at(3) + Inches(0.05), mid_y + Inches(0.15),
          ph_w - Inches(0.1), Inches(0.95),
          "Offline DLI capability", "in dev · Active", 0, BLUE)
epic_card(s, x_at(3) + Inches(0.05), mid_y + Inches(1.25),
          ph_w - Inches(0.1), Inches(0.85),
          "Mobile TLI UI tweak", "in dev · Active", 0, BLUE)
epic_card(s, x_at(4) + Inches(0.05), mid_y + Inches(0.15),
          ph_w - Inches(0.1), Inches(0.95),
          "Decommission FWM", "hypercare · 30d", 0, BLUE)

# Legend
leg_y = Inches(6.3)
items = [
    ("STORY", BLUE, "Default unit of work · pulled by capacity"),
    ("EPIC", AMBER, "Exception · Initiative > 3 months sliced into Epics (badge shows linked Stories)"),
    ("Card flow", DARK, "Stories move left → right through phases"),
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
# SLIDE 18b — Demand sizing logic (T-shirt sizes · routing · SLA · DoR per size)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Demand sizing · T-shirt size determines the path · routing · SLA · Definition of Ready depth")

text_box(s, Inches(0.5), Inches(1.05), SW - Inches(1), Inches(0.45),
         "Size is a field on the Jira card — not a separate channel. "
         "It tells the system how much governance the demand needs: small items skip the Portfolio Gate; "
         "large items go through full Discovery before sizing is locked.",
         size=11, italic=True, color=BLUE, bold=True)

sizes = [
    ("S", GREEN,
     "Small",
     "≤ $10k  ·  1–3 months  ·  < 5 resources",
     ["Direct to Squad's Team Kanban",
      "Skips Portfolio Gate",
      "Lightweight DoR — title · BV · acceptance criteria",
      "Triage SLA: 2 business days",
      "Sponsor: sub-team lead"]),
    ("M", BLUE,
     "Medium",
     "≤ $500k  ·  3–6 months  ·  ≤ 10 resources",
     ["Direct to Squad — flagged for awareness at Portfolio",
      "Lightweight Discovery (1–2 weeks)",
      "Standard DoR — adds CoS · OPCO · Pain Score",
      "Triage SLA: 5 business days",
      "Sponsor: business unit lead"]),
    ("L", AMBER,
     "Large",
     "> Medium thresholds  ·  > 6 months  ·  cross-team",
     ["Mandatory Portfolio Gate (monthly Gate 2)",
      "Full Discovery — outcome · users · solution shape",
      "Full DoR — adds budget · vendor model · architecture review",
      "Triage SLA: 5 business days · Gate decision next monthly",
      "Sponsor: Director (Carlos / Kumar / Heather)"]),
    ("XL", RED,
     "Extra-Large  [to validate]",
     "Multi-quarter  ·  cross-Value-Stream  ·  multi-vendor",
     ["Mandatory Portfolio Gate + VP Portfolio Sync",
      "Initiative sliced into Epics (Initiative > 3 months rule)",
      "Full DoR + business case + capital plan alignment",
      "Triage SLA: 10 business days · sponsor sign-off required",
      "Sponsor: Vitor Cruz · Iberdrola for cross-OPCO XL"]),
]

sz_y = Inches(1.7)
sz_h = Inches(4.3)
sz_w = (SW - Inches(1.0) - Inches(0.15) * 3) / 4

for i, (lbl, color, name, threshold, items) in enumerate(sizes):
    x = Inches(0.5) + i * (sz_w + Inches(0.15))
    round_rect(s, x, sz_y, sz_w, sz_h, WHITE, line=color, line_w=2, radius=0.05)
    rect(s, x, sz_y, sz_w, Inches(0.7), color)
    text_box(s, x + Inches(0.15), sz_y + Inches(0.05),
             Inches(0.7), Inches(0.6),
             lbl, size=28, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.95), sz_y + Inches(0.08),
             sz_w - Inches(1.05), Inches(0.55),
             name, size=12, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    # Threshold panel
    rect(s, x, sz_y + Inches(0.7), sz_w, Inches(0.5), LIGHT)
    text_box(s, x + Inches(0.15), sz_y + Inches(0.7),
             sz_w - Inches(0.3), Inches(0.5),
             threshold, size=9.5, bold=True, italic=True, color=color,
             anchor=MSO_ANCHOR.MIDDLE)
    add_lines(s, x + Inches(0.2), sz_y + Inches(1.35),
              sz_w - Inches(0.4), sz_h - Inches(1.5),
              ["•  " + it for it in items], size=9.5, color=DARK,
              line_spacing=1.4)

# Bottom — routing decision arrow
rt_y = Inches(6.15)
round_rect(s, Inches(0.5), rt_y, SW - Inches(1), Inches(0.75),
           BLUE_PALE, line=NAVY, line_w=1.5, radius=0.05)
text_box(s, Inches(0.7), rt_y + Inches(0.06),
         Inches(2.5), Inches(0.3),
         "ROUTING RULE",
         size=10, bold=True, color=NAVY)
text_box(s, Inches(0.7), rt_y + Inches(0.32),
         SW - Inches(1.4), Inches(0.45),
         "S / M  →  enter Squad's Team Kanban directly  ·  no Portfolio Gate.   "
         "L / XL  →  enter Portfolio Board · pass Gate 2 monthly · then descend to the Squad.",
         size=10, color=DARK, italic=True)

footer_bar(s)



# ════════════════════════════════════════════════════════════════════════
# SLIDE 19 — Class of Service framework
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
# SLIDE 20 — Roadmap (Wave 1 · 5 steps over 14 weeks · plus Wave 2 / 3)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Roadmap · Wave 1 — Demand Management Design & Pilot · five steps over fourteen weeks")

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

footer_bar(s)


# ════════════════════════════════════════════════════════════════════════
# Save
# ════════════════════════════════════════════════════════════════════════
# SLIDE 21 — TO-BE Demand Flow (one Jira front door · 8 VS routing)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "TO-BE Demand Flow · one Jira front door routes eleven source areas into eight Value Stream Teams")

text_box(s, Inches(0.5), Inches(1.0), SW - Inches(1), Inches(0.4),
         "Single intake form per Value Stream — every requester opens a Jira card with mandatory fields. "
         "The form auto-routes to the right Value Stream Team and applies a Class of Service.",
         size=11, italic=True, color=BLUE, bold=True)

# Layout: 4 columns, left to right flow
flow_y = Inches(1.55)
flow_h = Inches(5.4)
fcols = 4
fgap = Inches(0.12)
fcol_w = (SW - Inches(1.0) - fgap * (fcols - 1)) / fcols

# Column headers
fheaders = [
    ("01 · SOURCES", "Who sends demand (11 areas)"),
    ("02 · INTAKE FORM (JIRA)", "Mandatory fields capture intent"),
    ("03 · VALUE STREAM TEAM", "Auto-routed to one of eight"),
    ("04 · CLASS OF SERVICE", "Sets cadence and capacity rule"),
]
for i, (head, sub) in enumerate(fheaders):
    x = Inches(0.5) + i * (fcol_w + fgap)
    rect(s, x, flow_y, fcol_w, Inches(0.55), NAVY_DARK)
    text_box(s, x, flow_y + Inches(0.05), fcol_w, Inches(0.28),
             head, size=11, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x, flow_y + Inches(0.32), fcol_w, Inches(0.22),
             sub, size=8.5, italic=True, color=ACCENT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

fbody_y = flow_y + Inches(0.65)
fbody_h = flow_h - Inches(0.65)

# Column 1 — 11 sources
x1 = Inches(0.5)
sources = [
    ("BRS (UIL · CMP · NY)", BLUE),
    ("CX & Digital CoE", BLUE),
    ("Data & Analytics", BLUE),
    ("Infrastructure", BLUE),
    ("IA (Industrial Auto.)", BLUE),
    ("OT (Operational Tech.)", BLUE),
    ("Procurement", BLUE),
    ("Regulatory", BLUE),
    ("Corporate Areas", GREY_DARK),
    ("P&T (GIS · Global SAP)", GREY_DARK),
    ("Gas & Electric Ops", GREY_DARK),
]
src_h = (fbody_h - Inches(0.05) * (len(sources) - 1)) / len(sources)
for i, (src, col) in enumerate(sources):
    y = fbody_y + i * (src_h + Inches(0.05))
    round_rect(s, x1, y, fcol_w, src_h, col, radius=0.2)
    text_box(s, x1, y, fcol_w, src_h, src,
             size=9.5, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Column 2 — Single Jira request type (one big card listing fields)
x2 = Inches(0.5) + (fcol_w + fgap)
round_rect(s, x2, fbody_y, fcol_w, fbody_h, WHITE, line=ACCENT, line_w=2, radius=0.05)
rect(s, x2, fbody_y, fcol_w, Inches(0.45), ACCENT)
text_box(s, x2, fbody_y + Inches(0.05), fcol_w, Inches(0.4),
         "JIRA REQUEST FORM",
         size=12, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

form_fields = [
    ("Value Stream destination", "(routes the card)"),
    ("Demand type", "Run · Grow · Transform"),
    ("T-shirt size", "S · M · L · XL"),
    ("Class of Service", "Standard · Fixed-date · Expedite · Intangible"),
    ("OPCO", "UIL · CMP · NY · Iberdrola · Cross"),
    ("Business value statement", "Who benefits and how"),
    ("Sponsor", "(approves at intake)"),
    ("Pain Score inputs", "BV · UX · Safety · Effort"),
]
ff_y = fbody_y + Inches(0.55)
note_h = Inches(0.55)
ff_total_h = fbody_h - Inches(0.6) - note_h
ff_h = ff_total_h / len(form_fields)
for i, (lbl, hint) in enumerate(form_fields):
    y = ff_y + i * ff_h
    text_box(s, x2 + Inches(0.15), y,
             fcol_w - Inches(0.3), Inches(0.2),
             "▸ " + lbl, size=9.5, bold=True, color=NAVY,
             anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x2 + Inches(0.4), y + Inches(0.2),
             fcol_w - Inches(0.55), ff_h - Inches(0.22),
             hint, size=7.5, italic=True, color=GREY,
             anchor=MSO_ANCHOR.TOP)

# Note about IT Now (incidents only) — separated panel inside col 2 footer
note_y = fbody_y + fbody_h - note_h + Inches(0.05)
round_rect(s, x2 + Inches(0.1), note_y,
           fcol_w - Inches(0.2), note_h - Inches(0.15),
           RED_LIGHT, line=RED, line_w=1, radius=0.04)
text_box(s, x2 + Inches(0.2), note_y,
         fcol_w - Inches(0.4), note_h - Inches(0.15),
         "IT Now is reserved for incidents only — not for new demand.",
         size=8.5, italic=True, bold=True, color=RED,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Column 3 — 8 Value Stream Teams
x3 = Inches(0.5) + 2 * (fcol_w + fgap)
vs_teams = [
    ("M2C — Service Lifecycle", "Smart Metering / FICA", BLUE),
    ("M2C — Revenue Operations", "Billing / Collection", BLUE),
    ("Customer Care & CCC", "CRM / Contact Center", BLUE),
    ("Field Operations", "GIS / Workforce / OSG", GREEN),
    ("Engineering & Asset Mgmt", "Maintenance / Engineering", GREEN),
    ("Network Operations & Control", "SCADA · EMS · OMS", GREEN),
    ("Corporate & Regulatory", "FI · HR · Procurement · Reg.", GREEN),
    ("Iberdrola Corp Services", "Global SAP · Corp. Apps", GREEN),
]
vs_h = (fbody_h - Inches(0.06) * (len(vs_teams) - 1)) / len(vs_teams)
for i, (name, scope, col) in enumerate(vs_teams):
    y = fbody_y + i * (vs_h + Inches(0.06))
    round_rect(s, x3, y, fcol_w, vs_h, WHITE, line=col, line_w=1.5, radius=0.05)
    rect(s, x3, y, Inches(0.06), vs_h, col)
    text_box(s, x3 + Inches(0.15), y + Inches(0.04),
             fcol_w - Inches(0.2), Inches(0.24),
             name, size=10, bold=True, color=col)
    text_box(s, x3 + Inches(0.15), y + Inches(0.27),
             fcol_w - Inches(0.2), vs_h - Inches(0.3),
             scope, size=8.5, italic=True, color=DARK)

# Column 4 — Class of Service
x4 = Inches(0.5) + 3 * (fcol_w + fgap)
cos_types = [
    ("STANDARD", BLUE,
     "Default flow — pulled by capacity, prioritized by Pain Score.",
     "Bulk of demand · no special policy."),
    ("FIXED-DATE", AMBER,
     "External deadline (regulatory · contractual). Locked sequence.",
     "Slot reserved at portfolio gate."),
    ("EXPEDITE — capped ≤20% throughput", RED,
     "Bypass priority for genuine urgency. Sponsor approval logged.",
     "Replaces 'regulatory always wins'."),
    ("INTANGIBLE", PURPLE,
     "Tech debt · platform health · spike work.",
     "Reserved 10–15% capacity per quarter."),
]
cos_total_h = fbody_h
cos_h = (cos_total_h - Inches(0.08) * (len(cos_types) - 1)) / len(cos_types)
for i, (lbl, col, desc, note) in enumerate(cos_types):
    y = fbody_y + i * (cos_h + Inches(0.08))
    round_rect(s, x4, y, fcol_w, cos_h, WHITE, line=col, line_w=1.5, radius=0.05)
    rect(s, x4, y, fcol_w, Inches(0.36), col)
    text_box(s, x4 + Inches(0.1), y + Inches(0.03),
             fcol_w - Inches(0.2), Inches(0.32),
             lbl, size=10, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x4 + Inches(0.15), y + Inches(0.42),
             fcol_w - Inches(0.3), cos_h - Inches(0.5),
             desc, size=9, color=DARK)
    text_box(s, x4 + Inches(0.15), y + cos_h - Inches(0.32),
             fcol_w - Inches(0.3), Inches(0.28),
             note, size=8, italic=True, color=col)

# (No bottom legend — keeps the slide clean)


# (Slide 21 ends here — original AS-IS recipient/channel/legend code below has been replaced.)
footer_bar(s)


# ════════════════════════════════════════════════════════════════════════
# ════════════════════════════════════════════════════════════════════════
# SLIDE 22 — Future State Guiding Principles (3 issue areas + 7 gaps)
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Future State Guiding Principles · three problem areas converge into seven structural gaps")

# Top sub-title
text_box(s, Inches(0.5), Inches(1.0), SW - Inches(1), Inches(0.4),
         "Our assessment surfaces three connected problem areas — operating constraints, "
         "delivery pain points, and structural misalignment — that map to seven structural gaps "
         "in the current model.",
         size=11, italic=True, color=BLUE, bold=True)

# 3 problem areas — top row
pa_y = Inches(1.55)
pa_h = Inches(2.95)
pa_w = (SW - Inches(1.2)) / 3
gap = Inches(0.15)

problem_areas = [
    ("OPERATING CONSTRAINTS",
     "Structural and operational limitations constraining the model's "
     "ability to scale and deliver value consistently.",
     ["People — capacity & accountability are structurally constrained",
      "Process — lack of standardization drives inconsistency",
      "Tools — fragmented ecosystem limits visibility and control",
      "Governance — decision-making is fragmented"], BLUE),
    ("DELIVERY EXPERIENCE & PAIN POINTS",
     "Friction points impacting execution and the business or user "
     "experience.",
     ["Fragmented demand and intake drive reactive, non-strategic delivery",
      "Lack of clear product ownership weakens accountability",
      "Limited portfolio visibility constrains end-to-end planning",
      "Uncoordinated capacity and overlapping initiatives reduce impact"], ACCENT),
    ("STRUCTURAL MISALIGNMENT",
     "Misalignment between organizational structure, governance model, "
     "and the way products are managed and delivered.",
     ["No Product Entry Point",
      "No Value-Based Decision-Making",
      "No End-to-End Product View",
      "No Ownership of Outcomes",
      "No Product-Level Steering",
      "No Common Language of Value",
      "No Product Ownership Model"], RED),
]

for i, (name, sub, items, color) in enumerate(problem_areas):
    x = Inches(0.5) + i * (pa_w + gap)
    round_rect(s, x, pa_y, pa_w, pa_h, WHITE, line=color, line_w=2, radius=0.04)
    rect(s, x, pa_y, pa_w, Inches(0.45), color)
    text_box(s, x + Inches(0.1), pa_y + Inches(0.05),
             pa_w - Inches(0.2), Inches(0.4),
             name, size=11, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.15), pa_y + Inches(0.55),
             pa_w - Inches(0.3), Inches(0.55),
             sub, size=9.5, italic=True, color=color)
    add_lines(s, x + Inches(0.2), pa_y + Inches(1.1),
              pa_w - Inches(0.4), pa_h - Inches(1.2),
              ["•  " + it for it in items], size=9, color=DARK,
              line_spacing=1.3)

# Arrow / convergence marker
conv_y = pa_y + pa_h + Inches(0.12)
text_box(s, Inches(0.5), conv_y, SW - Inches(1), Inches(0.3),
         "▼  These problem areas converge into four Guiding Principles for the To-Be model  ▼",
         size=11, bold=True, italic=True, color=NAVY,
         align=PP_ALIGN.CENTER)

# 4 principles addressing the gaps — bottom row
pr_y = conv_y + Inches(0.4)
pr_h = Inches(1.5)
pr_w = (SW - Inches(1.3)) / 4
gap = Inches(0.1)

principles_short = [
    ("01", "DEMAND INTAKE & STRUCTURING",
     "Unified intake layer · single entry point · standard criteria",
     BLUE),
    ("02", "VALUE-BASED PRIORITIZATION",
     "Common scoring · transparent · product-level backlog ownership",
     ACCENT),
    ("03", "PRODUCT PORTFOLIO MGMT & PLANNING",
     "Integrated view · capacity at product level · roadmaps",
     GREEN),
    ("04", "PRODUCT GOVERNANCE & OWNERSHIP",
     "Roles defined · product forums · KPIs tied to outcomes",
     PURPLE),
]

for i, (num, name, desc, color) in enumerate(principles_short):
    x = Inches(0.5) + i * (pr_w + gap)
    round_rect(s, x, pr_y, pr_w, pr_h, color, radius=0.05)
    text_box(s, x + Inches(0.15), pr_y + Inches(0.06),
             Inches(0.7), Inches(0.45),
             num, size=22, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.85), pr_y + Inches(0.06),
             pr_w - Inches(1.0), Inches(0.45),
             name, size=10.5, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.15), pr_y + Inches(0.55),
             pr_w - Inches(0.3), pr_h - Inches(0.65),
             desc, size=9.5, italic=True, color=WHITE)

footer_bar(s)


# ─── Section divider · 08 Open Decisions & Next Steps ───
toc_slide(active_idx=7)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 23 — Open decisions & risks
# Storyline: this proposal is grounded but three decisions remain open and gate steady-state operation.
# ════════════════════════════════════════════════════════════════════════
s = blank()
header_bar(s, "Open decisions & risks · what needs validation before steady-state operation")

text_box(s, Inches(0.5), Inches(1.05), SW - Inches(1), Inches(0.4),
         "Three structural decisions remain open. Each gates a specific part of steady-state operation — "
         "we surface them now so they can be closed during Wave 1 with the right stakeholders.",
         size=11, italic=True, color=BLUE, bold=True)

decisions = [
    ("01", "VENDOR DELIVERY MODEL  ·  Option A vs Option B per Crew",
     "Lock the per-Crew choice between stable team (Option A) and shared pool (Option B). "
     "Recommendation: hybrid — Option A for seven continuous VSs · Option B for Corporate & Regulatory.",
     "OWNERS · Vitor Cruz · Carlos · Kumar · Partnership Crew",
     "WHEN · Wave 1 Step 2",
     "RISK · vendor churn · re-explanation cost",
     GREEN),
    ("02", "PRODUCT MANAGER STAFFING  ·  eight new roles",
     "The Product Manager role is NEW for all 8 Value Stream Crews — no Avangrid PM exists today "
     "(only PJMs · Sr. Managers · Champions). Decisions needed: job description · grade · internal vs external · sequencing.",
     "OWNERS · Vitor · HR · Carlos · Kumar · Heather",
     "WHEN · JD by Step 3 · staffing by Step 5",
     "RISK · framework without single accountability",
     BLUE),
    ("03", "NETWORK OPERATIONS & CONTROL  ·  sub-team owner for the new stream",
     "NOC is a NEW Value Stream surfaced from the IT Networks Application Assessment (5 apps · "
     "SCADA · EMS · OMS). Today there is no single sub-team owner identified inside Carlos's branch.",
     "OWNERS · Carlos · Travis · Vihn · Capability Crew",
     "WHEN · Owner ID by Step 1",
     "RISK · stream cannot be staffed",
     AMBER),
    ("04", "BUSINESS RELATIONSHIP ROLES  ·  per Ana's feedback",
     "Process Owners (business) · IT Business Partners per domain (Finance · Supply Chain · Asset Mgmt) · "
     "Demand Manager · ERP Steering Committee. None exist today as named roles — TOM must identify owners.",
     "OWNERS · Vitor · Carlos · Kumar · Heather · Iberdrola IT",
     "WHEN · role mapping by Step 2 · staffing TBD",
     "RISK · IT remains technical-only · no business voice",
     PURPLE),
    ("05", "RUN / BUILD SEPARATION  ·  per Value Stream",
     "L1 / L2 / L3 Application Support · Service Delivery Manager · Incident / Problem Manager — "
     "today undefined. Run team must be named separately from Build team inside each Squad.",
     "OWNERS · Carlos · Kumar · Platform Crew · Service Mgmt",
     "WHEN · Run track design by Step 3",
     "RISK · 'everyone does everything' persists",
     RED),
    ("06", "DATA & ANALYTICS CAPABILITY  ·  internal vs external",
     "Data Governance Lead · MDM · BW / SAC · Data Engineers — currently external. "
     "TOM places Data Ops inside Platform Crew but staffing model (internal · external · hybrid) is open.",
     "OWNERS · Vitor · Platform Crew lead · Data CoE if exists",
     "WHEN · Data Ops scope by Step 3",
     "RISK · no holistic data layer · vendor lock-in",
     INDIGO),
]

dec_y = Inches(1.6)
dec_h = (SH - dec_y - Inches(0.5)) / len(decisions)
for i, (num, title, body, owners, when, risk, color) in enumerate(decisions):
    y = dec_y + i * dec_h
    round_rect(s, Inches(0.5), y, SW - Inches(1), dec_h - Inches(0.1),
               WHITE, line=color, line_w=2, radius=0.05)
    # Number badge
    rect(s, Inches(0.5), y, Inches(0.7), dec_h - Inches(0.1), color)
    text_box(s, Inches(0.5), y, Inches(0.7), dec_h - Inches(0.1),
             num, size=22, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Body
    bx = Inches(1.35)
    text_box(s, bx, y + Inches(0.06),
             SW - bx - Inches(0.6), Inches(0.32),
             title, size=12, bold=True, color=color)
    text_box(s, bx, y + Inches(0.36),
             SW - bx - Inches(0.6), Inches(0.55),
             body, size=10, color=DARK, italic=True)
    # Footer chips: owners · when · risk
    cy = y + dec_h - Inches(0.6)
    text_box(s, bx, cy,
             Inches(4.3), Inches(0.5),
             owners, size=8.5, bold=True, color=NAVY,
             anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, bx + Inches(4.4), cy,
             Inches(3.5), Inches(0.5),
             when, size=8.5, bold=True, color=GREEN,
             anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, bx + Inches(8.0), cy,
             SW - bx - Inches(8.5), Inches(0.5),
             risk, size=8.5, italic=True, color=RED,
             anchor=MSO_ANCHOR.MIDDLE)

footer_bar(s)


# ════════════════════════════════════════════════════════════════════════
out_file = OUT / "DigitalProducts_TOM_Proposal_v4.0.pptx"
prs.save(str(out_file))
print(f"✓ Saved: {out_file}")
print(f"  Slides: {len(prs.slides)}")
