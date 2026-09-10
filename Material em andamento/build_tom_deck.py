"""
DigitalProducts_TOM_Proposal_v1.0.pptx
NTT Data consulting deck — Avangrid Digital Products TOM
Visual-first · KB-grounded · 20 slides
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pathlib import Path

OUT = Path("/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento")
VIS = OUT / "Miro_Visuals"

# ── NTT brand tokens ──
TETSU       = RGBColor(0x0F, 0x1C, 0x50)
HUMAN       = RGBColor(0x67, 0x85, 0xC1)
HUMAN_LIGHT = RGBColor(0xDC, 0xE6, 0xF5)
ACCENT      = RGBColor(0xE6, 0xB6, 0x00)
RISK        = RGBColor(0xBC, 0x43, 0x28)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
DARK        = RGBColor(0x2C, 0x3E, 0x50)
GREY        = RGBColor(0x6E, 0x78, 0x87)
LIGHT       = RGBColor(0xF5, 0xF5, 0xF5)
SUCCESS     = RGBColor(0x4C, 0x8C, 0x3F)
TEAL        = RGBColor(0x2B, 0x6B, 0x77)

FONT_BOLD = "Helvetica Neue"
FONT_REG  = "Helvetica Neue"

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
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False,
             font=FONT_REG):
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
        run.font.name = font
    return tb


def add_lines(slide, x, y, w, h, lines, size=12, color=DARK, bold=False,
              line_spacing=1.15, font=FONT_REG):
    """Add a multi-line text box where each line is its own paragraph."""
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
            run.font.name = font
    return tb


def action_title(slide, text, color=TETSU):
    """Top action title bar."""
    rect(slide, 0, 0, SW, Inches(0.85), color)
    text_box(slide, Inches(0.4), Inches(0.10), SW - Inches(0.8),
             Inches(0.65), text, size=20, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE, font=FONT_BOLD)
    # Slide number marker (small accent on right)
    rect(slide, SW - Inches(0.18), 0, Inches(0.18), Inches(0.85), ACCENT)


def takeaway(slide, text):
    """Bottom takeaway bar."""
    bar_y = SH - Inches(0.65)
    rect(slide, 0, bar_y, SW, Inches(0.65), TETSU)
    text_box(slide, Inches(0.4), bar_y + Inches(0.08), SW - Inches(0.8),
             Inches(0.5), "Takeaway · " + text, size=13, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE, italic=True, font=FONT_BOLD)


def footer(slide, sources=""):
    """Slide footer with sources + slide number area."""
    y = SH - Inches(0.85)
    text_box(slide, Inches(0.4), y, Inches(10), Inches(0.2),
             "Source: " + sources if sources else "", size=8,
             color=GREY, italic=True)
    text_box(slide, SW - Inches(2.5), y, Inches(2), Inches(0.2),
             "NTT Data · Avangrid · Apr 2026 · v1.0", size=8,
             color=GREY, italic=True, align=PP_ALIGN.RIGHT)


def add_image_safe(slide, path, x, y, w, h):
    if Path(path).exists():
        slide.shapes.add_picture(str(path), x, y, width=w, height=h)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 1 — COVER
# ════════════════════════════════════════════════════════════════════════
s = blank()
rect(s, 0, 0, SW, SH, TETSU)

# Accent bar
rect(s, 0, Inches(2.5), Inches(2.5), Inches(0.06), ACCENT)

# Project label
text_box(s, Inches(0.7), Inches(2.1), Inches(8), Inches(0.4),
         "AVANGRID · DIGITAL PRODUCTS",
         size=14, bold=True, color=ACCENT, font=FONT_BOLD)

# Main title
text_box(s, Inches(0.7), Inches(2.7), Inches(12), Inches(1.4),
         "Target Operating Model",
         size=52, bold=True, color=WHITE, font=FONT_BOLD)

# Subtitle
text_box(s, Inches(0.7), Inches(3.9), Inches(12), Inches(0.7),
         "unFIX-based proposal · 8 Value Stream Crews · 3-layer portfolio architecture",
         size=22, color=HUMAN, font=FONT_REG)

# Description block
text_box(s, Inches(0.7), Inches(5.2), Inches(11), Inches(1),
         "A product-centric operating model for Vitor Cruz's organization (~50 FTE across\n"
         "Carlos Garbin's IT Business Apps and Kumar Jaladanki's Customer Services Apps).",
         size=14, color=WHITE, font=FONT_REG)

# Bottom bar
rect(s, 0, SH - Inches(0.6), SW, Inches(0.6), HUMAN)
text_box(s, Inches(0.7), SH - Inches(0.55), Inches(10), Inches(0.5),
         "NTT DATA · D&A USA · Apr 2026 · v1.0",
         size=12, bold=True, color=WHITE, font=FONT_BOLD,
         anchor=MSO_ANCHOR.MIDDLE)


# ════════════════════════════════════════════════════════════════════════
# SLIDE 2 — EXECUTIVE SUMMARY
# ════════════════════════════════════════════════════════════════════════
s = blank()
action_title(s,
             "Eight product-aligned crews on a Strongly Aligned Base convert "
             "fragmented intake into a scalable operating model")

# Governing thought box
gt_y = Inches(1.1)
round_rect(s, Inches(0.5), gt_y, SW - Inches(1), Inches(1.0),
           HUMAN_LIGHT, line=HUMAN, line_w=2)
text_box(s, Inches(0.7), gt_y + Inches(0.1), SW - Inches(1.4), Inches(0.3),
         "GOVERNING THOUGHT", size=11, bold=True, color=HUMAN, font=FONT_BOLD)
text_box(s, Inches(0.7), gt_y + Inches(0.4), SW - Inches(1.4), Inches(0.6),
         "Adopt the unFIX Strongly Aligned Base — eight Value Stream Crews aligned to "
         "customer outcomes — to scale Digital Products through projected 30–50% growth "
         "without operational chaos.",
         size=14, color=DARK, italic=True)

# 4 supporting columns
col_y = Inches(2.4)
col_h = Inches(3.6)
col_w = (SW - Inches(1.2)) / 4
gap = Inches(0.15)

cols = [
    ("01", "TODAY", TETSU,
     ["50 FTE · 9 sub-teams · 30-50% growth",
      "7+ fragmented intake channels",
      "Zero Product Managers (only PJMs)",
      "\"Regulatory always wins\" by default",
      "Weakest of 5 NTT frameworks at Avangrid"]),
    ("02", "FRAMEWORK", HUMAN,
     ["unFIX — Apr 2026",
      "7 Crew types · 10 Forum types · 4 Base types",
      "Strongly Aligned Base = Digital Products",
      "Sibling Bases = Digital CoE · P&S/P&T",
      "Vendor pool = SW Delivery (segregated)"]),
    ("03", "DESIGN", SUCCESS,
     ["1 Governance Crew (existing directors)",
      "8 Value Stream Crews (CS x3 · IT x5)",
      "4 Supporting Crews + 1 Facilitation",
      "5 Forums for cross-crew coordination",
      "Avangrid core + Vendor extension model"]),
    ("04", "GOVERNANCE", ACCENT,
     ["3-layer portfolio · 11 boards/views",
      "VP Dashboard + 2 Portfolio + 8 Crew",
      "Initiatives → Epics → Features → Stories",
      "Pain Score H(x) = (BV·w₁+UX·w₂+Safety·w₃)/E",
      "Class of Service: Expedite cap ≤20%"]),
]
for i, (num, label, color, bullets) in enumerate(cols):
    x = Inches(0.5) + i * (col_w + gap)
    # Number badge
    round_rect(s, x, col_y, Inches(0.5), Inches(0.5), color, radius=0.15)
    text_box(s, x, col_y + Inches(0.02), Inches(0.5), Inches(0.5),
             num, size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font=FONT_BOLD)
    # Label
    text_box(s, x + Inches(0.6), col_y + Inches(0.05), col_w - Inches(0.6),
             Inches(0.4), label, size=14, bold=True, color=color, font=FONT_BOLD)
    # Body card
    body_y = col_y + Inches(0.65)
    round_rect(s, x, body_y, col_w, col_h - Inches(0.65),
               LIGHT, line=GREY, line_w=0.5)
    add_lines(s, x + Inches(0.1), body_y + Inches(0.1), col_w - Inches(0.2),
              col_h - Inches(0.85),
              ["• " + b for b in bullets], size=11, color=DARK)

takeaway(s, "Eight crews replace tech-aligned silos · vendor-stable delivery scales · Product Manager gap is the load-bearing decision.")
footer(s, "KB · Avangrid Digital Products v2.0 · 2604 AS-IS revisada · Prioritization Criteria & KPIs per Engagement")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 3 — Avangrid context: 3 parallel IT bases
# ════════════════════════════════════════════════════════════════════════
s = blank()
action_title(s,
             "Avangrid runs three parallel IT bases — Digital Products is the "
             "only one without a defined operating model")

# 3 cards
card_y = Inches(1.3)
card_h = Inches(5.0)
card_w = (SW - Inches(1.2)) / 3
gap = Inches(0.15)

bases = [
    ("DIGITAL CoE", "5 product teams (per OPCO)", HUMAN,
     "Most mature · OPCO-aligned",
     ["CNG · SCG · UI · CMP · NYSEG",
      "Strategy/Discovery + Delivery axes",
      "Heather LaJoie = first prioritization layer",
      "Stack: Jira · Big Room Planning · OKRs",
      "Customer-facing apps (NOT SAP)"]),
    ("P&S / P&T NETWORKS IT", "NOM v1.5 in implementation", SUCCESS,
     "Reference operating model · BRP cadence",
     ["Lean Operating Model v1.5 (Apr 2026)",
      "Wave 1 entregue · Wave 2 in proposal",
      "Pain Score 6-fator weighted",
      "2-gate governance (Shark Tank + Portfolio)",
      "Utility infra (grid · field · T&D)"]),
    ("DIGITAL PRODUCTS", "No defined operating model", RISK,
     "Carlos + Kumar · ~50 FTE",
     ["IT Business Apps (Carlos · 26 FTE)",
      "Customer Services Apps (Kumar · 24 FTE)",
      "Smart Metering = only PM exists today",
      "30-50% growth without TOM = chaos",
      "This proposal · Wave 1 design · 14-week"]),
]

for i, (name, sub, color, badge, bullets) in enumerate(bases):
    x = Inches(0.5) + i * (card_w + gap)
    # Card outline
    round_rect(s, x, card_y, card_w, card_h, WHITE, line=color, line_w=2)
    # Header band
    rect(s, x, card_y, card_w, Inches(0.6), color)
    text_box(s, x + Inches(0.15), card_y + Inches(0.05), card_w - Inches(0.3),
             Inches(0.5), name, size=15, bold=True, color=WHITE,
             anchor=MSO_ANCHOR.MIDDLE, font=FONT_BOLD)
    # Subtitle
    text_box(s, x + Inches(0.15), card_y + Inches(0.7), card_w - Inches(0.3),
             Inches(0.3), sub, size=12, italic=True, color=color)
    # Status badge
    round_rect(s, x + Inches(0.15), card_y + Inches(1.05),
               card_w - Inches(0.3), Inches(0.45), color, radius=0.2)
    text_box(s, x + Inches(0.15), card_y + Inches(1.05),
             card_w - Inches(0.3), Inches(0.45), badge,
             size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font=FONT_BOLD)
    # Body
    add_lines(s, x + Inches(0.15), card_y + Inches(1.7),
              card_w - Inches(0.3), card_h - Inches(2.0),
              ["• " + b for b in bullets], size=11, color=DARK)

takeaway(s, "Digital CoE and P&S/P&T have moved · Digital Products needs the same operating-model shift.")
footer(s, "Avangrid_Hierarchy_Overview · NTTDATA-Digital Strategy NextGen Product Ops · nWOW v1.5")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 4 — 5 of 5 NTT engagements moved · Digital Products weakest
# ════════════════════════════════════════════════════════════════════════
s = blank()
action_title(s,
             "Five of five NTT engagements at Avangrid moved to product-centric — "
             "Digital Products is the weakest of the five frameworks")

# Comparison table
tbl_y = Inches(1.2)
headers = ["Engagement", "Pain Score / Formula", "Gating", "KPI ecosystem", "Maturity"]
rows = [
    ["Digital CoE (CX)",   "H(x) = (BV+UX)/E weighted", "VP Sign-off", "3-layer ecosystem", "✓ Most formal"],
    ["P&S / P&T",          "Pain Score 6-fator weighted", "Shark Tank + Portfolio", "Multi-tier", "✓ Most weighted"],
    ["Renewables DTO",     "DoR checklist + Classes of Service", "Lean Kanban gating", "Flow metrics", "✓ Most disciplined"],
    ["Infrastructure",     "Stakeholder leader-driven", "Pilot → expansion", "Pilot effectiveness", "✓ Pragmatic"],
    ["SAP Digital Products", "No formal scoring", "Only PDD (large items)", "Fragmented per sub-team", "✗ Weakest"],
]

col_widths = [Inches(2.6), Inches(2.8), Inches(2.4), Inches(2.4), Inches(2.0)]
row_h = Inches(0.55)

# Header row
x_cur = Inches(0.4)
for i, h in enumerate(headers):
    rect(s, x_cur, tbl_y, col_widths[i], Inches(0.5), TETSU)
    text_box(s, x_cur, tbl_y, col_widths[i], Inches(0.5), h,
             size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font=FONT_BOLD)
    x_cur += col_widths[i]

# Data rows
for r, row in enumerate(rows):
    y = tbl_y + Inches(0.5) + r * row_h
    is_last = (r == len(rows) - 1)
    bg = LIGHT if r % 2 == 0 else WHITE
    if is_last:
        bg = RGBColor(0xFF, 0xE5, 0xE0)  # light red for the weakest
    x_cur = Inches(0.4)
    for i, cell in enumerate(row):
        rect(s, x_cur, y, col_widths[i], row_h, bg, line=GREY, line_w=0.25)
        color = RISK if (is_last and i == 4) else (SUCCESS if "✓" in cell else DARK)
        bold = is_last
        text_box(s, x_cur + Inches(0.1), y, col_widths[i] - Inches(0.2), row_h,
                 cell, size=10, bold=bold, color=color, anchor=MSO_ANCHOR.MIDDLE)
        x_cur += col_widths[i]

# Insight callout
ins_y = tbl_y + Inches(0.5) + 5 * row_h + Inches(0.4)
round_rect(s, Inches(0.4), ins_y, SW - Inches(0.8), Inches(1.0),
           HUMAN_LIGHT, line=HUMAN, line_w=1.5)
text_box(s, Inches(0.6), ins_y + Inches(0.1), SW - Inches(1.2), Inches(0.4),
         "WHAT TO BORROW (cross-engagement synthesis)", size=11, bold=True, color=HUMAN, font=FONT_BOLD)
text_box(s, Inches(0.6), ins_y + Inches(0.4), SW - Inches(1.2), Inches(0.6),
         "Pain Score formula: H(x) = (BV·w₁ + UX·w₂ + Safety·w₃) / E — Digital CoE × P&S/P&T hybrid · "
         "Classes of Service from Renewables DTO · Two-gate governance from P&S/P&T (Shark Tank + Portfolio Committee).",
         size=12, color=DARK)

takeaway(s, "We do not start from zero · we adopt the formula and gating that 4 sibling engagements already validated.")
footer(s, "Prioritization Criteria and KPIs per Engagement · NTT Data · Apr 2026")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 5 — Digital Products scope
# ════════════════════════════════════════════════════════════════════════
s = blank()
action_title(s,
             "Digital Products covers Carlos's 26-FTE IT Business Apps and "
             "Kumar's 24-FTE Customer Services Apps under VP Vitor Cruz")

# VP box at top
vp_y = Inches(1.2)
round_rect(s, Inches(4), vp_y, Inches(5.3), Inches(0.7), TETSU, radius=0.25)
text_box(s, Inches(4), vp_y, Inches(5.3), Inches(0.7),
         "VITOR CRUZ · VP IT Business Solutions Delivery",
         size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font=FONT_BOLD)

# Two branches side-by-side
branch_y = Inches(2.2)
branch_h = Inches(4.2)
b_w = Inches(6.2)
b_x_left = Inches(0.4)
b_x_right = Inches(6.7)

# Connector lines
line_y = vp_y + Inches(0.7)
shp = s.shapes.add_connector(1, Inches(6.6), line_y, Inches(3.5), branch_y)
shp.line.color.rgb = HUMAN
shp.line.width = Pt(2)
shp = s.shapes.add_connector(1, Inches(6.6), line_y, Inches(9.8), branch_y)
shp.line.color.rgb = HUMAN
shp.line.width = Pt(2)

# Carlos branch (left)
round_rect(s, b_x_left, branch_y, b_w, branch_h, WHITE, line=HUMAN, line_w=2)
rect(s, b_x_left, branch_y, b_w, Inches(0.5), HUMAN)
text_box(s, b_x_left + Inches(0.15), branch_y + Inches(0.05), b_w, Inches(0.4),
         "IT BUSINESS APPS · Carlos Garbin · ~26 FTE",
         size=13, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, font=FONT_BOLD)
carlos_subs = [
    ("E&G Operations", "Bruno Moreto + Andrew McGibbon", "~7 FTE", "70-90 apps"),
    ("GIS & OSG", "George Porto", "~6 FTE", "ArcFM · GIS · HANA · OSG"),
    ("Corporate Apps", "Sarath Garikapati + Amit", "~8 FTE", "5-6 PJMs + 2 Tech Leads"),
    ("Global SAP (Iberdrola)", "Vanessa Colombo + Michael", "~5 FTE", "FI · PO · Budget"),
]
for i, (n, lead, fte, scope) in enumerate(carlos_subs):
    y = branch_y + Inches(0.65) + i * Inches(0.85)
    round_rect(s, b_x_left + Inches(0.2), y, b_w - Inches(0.4), Inches(0.75),
               LIGHT, line=GREY, line_w=0.5)
    text_box(s, b_x_left + Inches(0.3), y + Inches(0.05),
             b_w - Inches(0.6), Inches(0.3), n,
             size=12, bold=True, color=DARK, font=FONT_BOLD)
    text_box(s, b_x_left + Inches(0.3), y + Inches(0.32),
             b_w - Inches(0.6), Inches(0.2), lead, size=10, color=GREY)
    text_box(s, b_x_left + Inches(0.3), y + Inches(0.5),
             b_w - Inches(0.6), Inches(0.2), f"{fte}  ·  {scope}",
             size=9, color=GREY, italic=True)

# Kumar branch (right)
round_rect(s, b_x_right, branch_y, b_w, branch_h, WHITE, line=ACCENT, line_w=2)
rect(s, b_x_right, branch_y, b_w, Inches(0.5), ACCENT)
text_box(s, b_x_right + Inches(0.15), branch_y + Inches(0.05),
         b_w, Inches(0.4),
         "CUSTOMER SERVICES APPS · Kumar Jaladanki · ~24 FTE",
         size=13, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE, font=FONT_BOLD)
kumar_subs = [
    ("FICA · Billing · Collections", "Ivelisse Velazquez", "~8 FTE", "FICA · ECC · StreamServe"),
    ("Device & Metering", "Vincent Lam + Iram Chowdhury (PJM)", "~8 FTE", "SAP Metering · vendor stack"),
    ("CRM & Contact Center", "VACANT Sr.M.", "~4 FTE", "SAP CRM · Medallia · Blue Prism"),
    ("SAP & CS Architecture", "VACANT Sr.M.", "~2 FTE", "Architecture · cross-stack"),
    ("SAP Basis", "Nagi Reddy", "~2 FTE", "Platform infrastructure"),
]
for i, (n, lead, fte, scope) in enumerate(kumar_subs):
    y = branch_y + Inches(0.6) + i * Inches(0.69)
    is_vacant = "VACANT" in lead
    bg_color = RGBColor(0xFF, 0xE5, 0xE0) if is_vacant else LIGHT
    round_rect(s, b_x_right + Inches(0.2), y, b_w - Inches(0.4), Inches(0.6),
               bg_color, line=GREY, line_w=0.5)
    text_box(s, b_x_right + Inches(0.3), y + Inches(0.02),
             b_w - Inches(0.6), Inches(0.25), n,
             size=11, bold=True, color=DARK, font=FONT_BOLD)
    lead_color = RISK if is_vacant else GREY
    text_box(s, b_x_right + Inches(0.3), y + Inches(0.24),
             b_w - Inches(0.6), Inches(0.2), lead, size=9.5,
             color=lead_color, bold=is_vacant)
    text_box(s, b_x_right + Inches(0.3), y + Inches(0.4),
             b_w - Inches(0.6), Inches(0.2), f"{fte}  ·  {scope}",
             size=9, color=GREY, italic=True)

takeaway(s, "Total ~50 FTE Avangrid · 9 sub-teams · 2 critical Sr. Manager vacancies in Kumar branch.")
footer(s, "Avangrid_Hierarchy_Overview §2.1 + §2.2 · 2604 AS-IS revisada Slide 39 + 41")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 6 — unFIX framework primer
# ════════════════════════════════════════════════════════════════════════
s = blank()
action_title(s, "unFIX framework offers seven Crew types, ten Forum types, and four Base types as the design vocabulary")

# 3 columns: Bases · Crews · Forums
col_y = Inches(1.2)
col_h = Inches(5.2)
col_w = (SW - Inches(1.2)) / 3
gap = Inches(0.15)

bases_list = [
    ("Fully Integrated", "single value stream · monolithic"),
    ("Strongly Aligned", "multiple value streams · cohesive product · dependencies"),
    ("Loosely Aligned", "multiple value streams · occasional coordination"),
    ("Fully Segregated", "independent / competing products"),
]
crews_list = [
    ("Governance", TETSU, "management team of the Base"),
    ("Value Stream", ACCENT, "end-to-end value · 3-7 people"),
    ("Capability", RISK, "rare expertise to other Crews"),
    ("Experience", RGBColor(0xE5, 0x8F, 0xB5), "customer journey across products"),
    ("Platform", SUCCESS, "shared services to other Crews"),
    ("Partnership", RGBColor(0xB6, 0xD5, 0x4E), "vendors / freelancers / partners"),
    ("Facilitation", RGBColor(0x7A, 0x66, 0xB8), "supports other Crews · no own value stream"),
]
forums_list = [
    "Business Model",
    "Channel",
    "Customer Journey",
    "Functional",
    "Market",
    "Organizational",
    "Product",
    "Regional",
    "Seasonal",
    "Technological",
]

# Bases column
x = Inches(0.5)
round_rect(s, x, col_y, col_w, col_h, WHITE, line=HUMAN, line_w=1.5)
rect(s, x, col_y, col_w, Inches(0.55), HUMAN)
text_box(s, x, col_y, col_w, Inches(0.55), "4 BASE TYPES",
         size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font=FONT_BOLD)
for i, (n, desc) in enumerate(bases_list):
    by = col_y + Inches(0.7) + i * Inches(1.05)
    is_ours = (n == "Strongly Aligned")
    fill = HUMAN_LIGHT if is_ours else LIGHT
    round_rect(s, x + Inches(0.15), by, col_w - Inches(0.3), Inches(0.95),
               fill, line=HUMAN if is_ours else GREY, line_w=2 if is_ours else 0.5)
    text_box(s, x + Inches(0.25), by + Inches(0.1), col_w - Inches(0.5), Inches(0.3),
             n + (" ← Digital Products" if is_ours else ""),
             size=12, bold=True, color=HUMAN if is_ours else DARK, font=FONT_BOLD)
    text_box(s, x + Inches(0.25), by + Inches(0.4), col_w - Inches(0.5), Inches(0.5),
             desc, size=10, color=GREY, italic=True)

# Crews column
x = Inches(0.5) + col_w + gap
round_rect(s, x, col_y, col_w, col_h, WHITE, line=ACCENT, line_w=1.5)
rect(s, x, col_y, col_w, Inches(0.55), ACCENT)
text_box(s, x, col_y, col_w, Inches(0.55), "7 CREW TYPES",
         size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font=FONT_BOLD)
for i, (n, color, desc) in enumerate(crews_list):
    by = col_y + Inches(0.65) + i * Inches(0.65)
    rect(s, x + Inches(0.15), by + Inches(0.1), Inches(0.18), Inches(0.45), color)
    text_box(s, x + Inches(0.4), by + Inches(0.1),
             col_w - Inches(0.55), Inches(0.22), n,
             size=11, bold=True, color=DARK, font=FONT_BOLD)
    text_box(s, x + Inches(0.4), by + Inches(0.32),
             col_w - Inches(0.55), Inches(0.22), desc,
             size=9, color=GREY, italic=True)

# Forums column
x = Inches(0.5) + 2 * (col_w + gap)
round_rect(s, x, col_y, col_w, col_h, WHITE, line=SUCCESS, line_w=1.5)
rect(s, x, col_y, col_w, Inches(0.55), SUCCESS)
text_box(s, x, col_y, col_w, Inches(0.55), "10 FORUM TYPES",
         size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font=FONT_BOLD)
for i, n in enumerate(forums_list):
    by = col_y + Inches(0.7) + i * Inches(0.43)
    is_used = n in ("Customer Journey", "Functional", "Organizational",
                    "Product", "Regional", "Technological")
    fill = HUMAN_LIGHT if is_used else LIGHT
    round_rect(s, x + Inches(0.15), by, col_w - Inches(0.3), Inches(0.36),
               fill, line=SUCCESS if is_used else GREY, line_w=1 if is_used else 0.3)
    text_box(s, x + Inches(0.3), by + Inches(0.04),
             col_w - Inches(0.5), Inches(0.28), n + (" ✓" if is_used else ""),
             size=11, bold=is_used,
             color=DARK if is_used else GREY, font=FONT_BOLD if is_used else FONT_REG)

takeaway(s, "Strongly Aligned Base + 7 Crews + 6 of 10 Forums applicable to Avangrid Digital Products.")
footer(s, "unfix.com — Crew Types · Base Types · Forum Types (Apr 2026 fetch)")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 7 — Strongly Aligned Base · master visual (embed PNG)
# ════════════════════════════════════════════════════════════════════════
s = blank()
action_title(s, "Digital Products becomes a Strongly Aligned Base with one Governance Crew, eight Value Stream Crews, four Supporting Crews, and one Facilitation Crew")

# Embed the existing diagram
img_path = VIS / "DigitalProducts_Unfix_Squads.png"
if img_path.exists():
    add_image_safe(s, img_path,
                   Inches(0.4), Inches(1.05),
                   Inches(12.5), Inches(5.6))

footer(s, "Synthesis · unFIX framework + Avangrid org · Apr 2026")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 8 — Governance Crew detail
# ════════════════════════════════════════════════════════════════════════
s = blank()
action_title(s, "The Governance Crew anchors strategy with the six existing directors — no new structural hires at the top")

# 6 director cards in 2x3
gx = Inches(0.4)
gy = Inches(1.2)
gw = (SW - Inches(1.0)) / 3
gh = Inches(2.2)
gap_x = Inches(0.1)
gap_y = Inches(0.15)

directors = [
    ("Vitor Cruz", "VP IT Business Solutions Delivery",
     "Sponsor of the Base · vision · cross-base alignment", TETSU),
    ("Carlos Garbin", "Director · IT Business Apps",
     "Owns IT Apps Portfolio Board · 5 Value Stream Crews", HUMAN),
    ("Kumar Jaladanki", "Director · Customer Services Apps",
     "Owns CS Apps Portfolio Board · 3 Value Stream Crews", HUMAN),
    ("Heather LaJoie", "Director · Product Management (business)",
     "2 FTE · Cross-base bridge · first prioritization layer", ACCENT),
    ("Travis · PMO Lead", "Large Projects PDD owner",
     "Bridge to Avangrid PDD process · NEW interview needed", SUCCESS),
    ("Portfolio Manager", "NEW role · cross-portfolio governance",
     "Owns Layer 0 VP Dashboard · monthly Portfolio Sync", RISK),
]

for i, (name, title, role, color) in enumerate(directors):
    row = i // 3
    col = i % 3
    x = gx + col * (gw + gap_x)
    y = gy + row * (gh + gap_y)
    round_rect(s, x, y, gw, gh, WHITE, line=color, line_w=2)
    # Top accent
    rect(s, x, y, gw, Inches(0.4), color)
    text_box(s, x + Inches(0.15), y + Inches(0.05), gw - Inches(0.3), Inches(0.3),
             name, size=13, bold=True, color=WHITE, font=FONT_BOLD,
             anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.15), y + Inches(0.5), gw - Inches(0.3), Inches(0.4),
             title, size=11, bold=True, color=DARK, font=FONT_BOLD)
    text_box(s, x + Inches(0.15), y + Inches(0.95), gw - Inches(0.3),
             gh - Inches(1.05), role, size=10, color=GREY, italic=True)

takeaway(s, "Five of six already in place · only Portfolio Manager is a NEW seat to fill.")
footer(s, "Avangrid_Hierarchy_Overview · Org_Chart_Digital_Products")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 9 — 8 Value Stream Crews overview
# ════════════════════════════════════════════════════════════════════════
s = blank()
action_title(s, "Eight Value Stream Crews align to customer outcomes — Network Operations & Control is the stream surfaced by the IT Networks 46-app Assessment")

tbl_y = Inches(1.1)
headers = ["#", "Value Stream Crew", "Branch", "Outcome (user)", "FTE Avangrid core"]
rows = [
    ["1", "M2C — Service Lifecycle", "CS · Kumar", "Customer signs up · activates · changes · disconnects", "3-4"],
    ["2", "M2C — Revenue Operations", "CS · Kumar", "Accurate billing · easy payment · fair recovery", "4-6"],
    ["3", "Customer Care & CCC", "CS · Kumar", "Issue resolution · self-service · feedback", "2-4"],
    ["4", "Field Operations", "IT · Carlos", "Field worker / inspector / lineman productivity", "3-4"],
    ["5", "Engineering & Asset Mgmt", "IT · Carlos", "Engineer / planner designs with reliable data", "4-5"],
    ["6", "Network Operations & Control", "IT · Carlos", "Grid operator prevents outages · NEW STREAM", "2-3"],
    ["7", "Corporate & Regulatory", "IT · Carlos", "HR / employee / non-NERC compliance self-service", "2-3"],
    ["8", "Iberdrola Corp Services", "IT · Carlos", "Iberdrola Spain finance / procurement / budget", "3-4"],
]
col_widths = [Inches(0.5), Inches(3.2), Inches(1.6), Inches(5.4), Inches(1.8)]
row_h = Inches(0.45)

# Header
x_cur = Inches(0.4)
for i, h in enumerate(headers):
    rect(s, x_cur, tbl_y, col_widths[i], Inches(0.45), TETSU)
    text_box(s, x_cur, tbl_y, col_widths[i], Inches(0.45), h,
             size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font=FONT_BOLD)
    x_cur += col_widths[i]

for r, row in enumerate(rows):
    y = tbl_y + Inches(0.45) + r * row_h
    bg = LIGHT if r % 2 == 0 else WHITE
    if "NEW" in row[3]:
        bg = RGBColor(0xFF, 0xF3, 0xCD)
    x_cur = Inches(0.4)
    branch_color = ACCENT if "CS" in row[2] else HUMAN
    for i, cell in enumerate(row):
        rect(s, x_cur, y, col_widths[i], row_h, bg, line=GREY, line_w=0.25)
        if i == 0:
            text_box(s, x_cur, y, col_widths[i], row_h, cell,
                     size=12, bold=True, color=branch_color, align=PP_ALIGN.CENTER,
                     anchor=MSO_ANCHOR.MIDDLE, font=FONT_BOLD)
        elif i == 2:
            text_box(s, x_cur, y, col_widths[i], row_h, cell,
                     size=10, bold=True, color=branch_color, align=PP_ALIGN.CENTER,
                     anchor=MSO_ANCHOR.MIDDLE, font=FONT_BOLD)
        else:
            color = TETSU if "NEW" in cell else DARK
            text_box(s, x_cur + Inches(0.1), y, col_widths[i] - Inches(0.2),
                     row_h, cell, size=10, color=color, anchor=MSO_ANCHOR.MIDDLE,
                     bold="NEW" in cell)
        x_cur += col_widths[i]

# Footnote about new stream
ft_y = tbl_y + Inches(0.45) + 8 * row_h + Inches(0.2)
round_rect(s, Inches(0.4), ft_y, SW - Inches(0.8), Inches(0.65),
           HUMAN_LIGHT, line=HUMAN, line_w=1)
text_box(s, Inches(0.6), ft_y + Inches(0.1), SW - Inches(1.2), Inches(0.5),
         "Network Operations & Control surfaced from IT Networks Application Assessment "
         "— 46 apps grouped into 7 functional domains; SCADA/EMS/OMS (5 apps) had no owner stream in our initial design.",
         size=11, color=DARK, italic=True)

takeaway(s, "Eight crews · ~25-32 FTE Avangrid core · vendors deliver the bulk of execution.")
footer(s, "Application Assessment IT Networks Slide 23 · 2604 AS-IS revisada · Avangrid_Hierarchy_Overview")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 10 — CS Apps streams detail
# ════════════════════════════════════════════════════════════════════════
s = blank()
action_title(s, "Customer Services Apps splits into three crews because M2C's 16-FTE size exceeds the unFIX 3-7 person rule")

cs_streams = [
    ("M2C — SERVICE LIFECYCLE",
     "Customer onboarding · activation · changes · disconnection",
     "Vincent Lam (Sr.M.) + Iram Chowdhury (PJM today)",
     "SAP Metering · MV90 · SAP CRM (front-office)",
     "Smart Metering vendor (100%) · Accenture · SWF",
     ACCENT),
    ("M2C — REVENUE OPERATIONS",
     "Accurate billing · easy payment · fair collections · revenue recovery",
     "Ivelisse Velazquez (Sr.M.)",
     "SAP FICA · SAP ECC · StreamServe · SOVOS · Precisely",
     "Accenture · SWF · SOVOS partner",
     ACCENT),
    ("CUSTOMER CARE & CCC",
     "Issue resolution · self-service · voice of customer",
     "VACANT Sr.M. CRM/CCC + Heather LaJoie (business)",
     "SAP CRM · Medallia · Blue Prism · Nuance · IVR",
     "Accenture · HCL/Nuance · Blue Prism partner",
     ACCENT),
]
y_cur = Inches(1.15)
card_h = Inches(1.85)
gap = Inches(0.1)

for name, outcome, lead, apps, vendor, color in cs_streams:
    round_rect(s, Inches(0.4), y_cur, SW - Inches(0.8), card_h,
               WHITE, line=color, line_w=2)
    # Header band
    rect(s, Inches(0.4), y_cur, Inches(3.5), card_h, color)
    text_box(s, Inches(0.55), y_cur + Inches(0.1), Inches(3.3), Inches(0.4),
             name, size=14, bold=True, color=WHITE, font=FONT_BOLD)
    text_box(s, Inches(0.55), y_cur + Inches(0.55), Inches(3.3), Inches(0.5),
             outcome, size=11, italic=True, color=WHITE)
    text_box(s, Inches(0.55), y_cur + Inches(1.15), Inches(3.3), Inches(0.4),
             "PM: NEW", size=11, bold=True, color=WHITE, font=FONT_BOLD)
    text_box(s, Inches(0.55), y_cur + Inches(1.45), Inches(3.3), Inches(0.4),
             "Branch: CS Apps · Kumar", size=10, italic=True, color=WHITE)

    # Body
    body_x = Inches(4.0)
    body_w = SW - Inches(4.5)
    text_box(s, body_x, y_cur + Inches(0.1), body_w, Inches(0.3),
             "TODAY · Lead", size=9, bold=True, color=GREY, font=FONT_BOLD)
    text_box(s, body_x, y_cur + Inches(0.35), body_w, Inches(0.3),
             lead, size=11, color=DARK)
    text_box(s, body_x, y_cur + Inches(0.7), body_w, Inches(0.3),
             "CURRENT APPS (KB)", size=9, bold=True, color=GREY, font=FONT_BOLD)
    text_box(s, body_x, y_cur + Inches(0.95), body_w, Inches(0.3),
             apps, size=11, color=DARK)
    text_box(s, body_x, y_cur + Inches(1.3), body_w, Inches(0.3),
             "VENDOR EXTENSION", size=9, bold=True, color=GREY, font=FONT_BOLD)
    text_box(s, body_x, y_cur + Inches(1.55), body_w, Inches(0.3),
             vendor, size=11, color=DARK)
    y_cur += card_h + gap

takeaway(s, "M2C is split into 2 crews (Service Lifecycle · Revenue Operations) coordinated via Customer Journey Forum.")
footer(s, "Slide 41 vrevisada · APR 2026 Status Reports · Initiative Cards")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 11 — IT Apps streams detail (5 streams)
# ════════════════════════════════════════════════════════════════════════
s = blank()
action_title(s, "IT Business Apps splits into five crews mapped to the seven functional domains of the Application Assessment")

it_streams = [
    ("FIELD OPERATIONS",
     "Field worker / inspector / lineman productivity",
     "Bruno Moreto (Sr.M.) + Andrew McGibbon (Champion)",
     "Lemur-Salesforce · DLI/TLI/DGI/TGI · 70-90 field apps · Storm toolkit",
     "Storm & Workforce (4) + Maintenance subset"),
    ("ENGINEERING & ASSET MANAGEMENT",
     "Engineer / planner designs with reliable data",
     "George Porto (Sr.M. + Champion · dual role)",
     "ArcFM · GIS web · ProjectWise · Microstation · SAP HANA · OSG · IRTH",
     "Engineering (5) + Planning (7) + GIS+Asset"),
    ("NETWORK OPERATIONS & CONTROL · NEW",
     "Grid operator prevents outages · 24/7 real-time",
     "TBD · likely under Bruno's E&G today",
     "SCADA · EMS · OMS · Echelon 2.0 · real-time monitoring",
     "System Operations & Control (5)"),
    ("CORPORATE & REGULATORY",
     "HR / employee / non-NERC compliance self-service",
     "Sarath Garikapati + Amit",
     "Corporate apps · Network business · OpenText · SharePoint→O365 · Rate Case Advisor",
     "Other (4) + non-Networks corporate"),
    ("IBERDROLA CORP SERVICES",
     "Iberdrola Spain finance / procurement / budget",
     "Vanessa Colombo + Michael Dubois",
     "SAP FI · SAP PO · Budget · S/4HANA · Supply Chain Digit.",
     "Iberdrola corp governance overlay"),
]

y_cur = Inches(1.05)
card_h = Inches(1.13)
gap = Inches(0.05)

for i, (name, outcome, lead, apps, domain) in enumerate(it_streams):
    is_new = "NEW" in name
    color = ACCENT if is_new else HUMAN
    round_rect(s, Inches(0.4), y_cur, SW - Inches(0.8), card_h,
               WHITE, line=color, line_w=2)
    rect(s, Inches(0.4), y_cur, Inches(3.6), card_h, color)
    text_box(s, Inches(0.5), y_cur + Inches(0.05), Inches(3.4), Inches(0.3),
             name, size=12, bold=True, color=WHITE, font=FONT_BOLD)
    text_box(s, Inches(0.5), y_cur + Inches(0.35), Inches(3.4), Inches(0.4),
             outcome, size=10, italic=True, color=WHITE)
    text_box(s, Inches(0.5), y_cur + Inches(0.78), Inches(3.4), Inches(0.3),
             domain, size=9, color=WHITE, italic=True)

    body_x = Inches(4.1)
    body_w = SW - Inches(4.5)
    text_box(s, body_x, y_cur + Inches(0.05), body_w, Inches(0.25),
             "TODAY · " + lead, size=10, bold=True, color=DARK, font=FONT_BOLD)
    text_box(s, body_x, y_cur + Inches(0.32), body_w, Inches(0.7),
             "Apps: " + apps, size=10, color=DARK)
    y_cur += card_h + gap

takeaway(s, "Five crews map cleanly to the 46-app assessment · NOC is the new stream that surfaced as a gap.")
footer(s, "Application Assessment IT Networks Slide 23 · APR 2026 Status Reports · Echelon Kick Off")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 12 — Vendor reality + Option A vs B
# ════════════════════════════════════════════════════════════════════════
s = blank()
action_title(s, "Each crew is Avangrid core (3-5 FTE) plus Vendor extension — choosing stable vs ad-hoc vendor model is the load-bearing decision")

# Top: visual model
top_y = Inches(1.1)
# Avangrid core box
round_rect(s, Inches(0.6), top_y, Inches(5.5), Inches(2.5),
           HUMAN_LIGHT, line=HUMAN, line_w=2)
text_box(s, Inches(0.7), top_y + Inches(0.1), Inches(5.3), Inches(0.4),
         "AVANGRID CORE  ·  3-5 FTE per crew", size=14, bold=True,
         color=HUMAN, font=FONT_BOLD)
add_lines(s, Inches(0.8), top_y + Inches(0.55), Inches(5.3), Inches(1.9),
          ["• Product Manager  (NEW role — does not exist today)",
           "• Tech Lead / Solution Architect",
           "• Business Analyst (BA)",
           "• 1-2 SMEs / Champions (Iram-style)",
           "Owns: outcome · scoring · backlog priority",
           "Does NOT do most of the build."], size=11, color=DARK)

# Plus sign
text_box(s, Inches(6.1), top_y + Inches(0.9), Inches(0.6), Inches(0.7),
         "+", size=72, bold=True, color=ACCENT, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, font=FONT_BOLD)

# Vendor extension box
round_rect(s, Inches(6.8), top_y, Inches(6.0), Inches(2.5),
           RGBColor(0xFF, 0xF3, 0xCD), line=ACCENT, line_w=2)
text_box(s, Inches(6.9), top_y + Inches(0.1), Inches(5.8), Inches(0.4),
         "VENDOR EXTENSION  ·  variable, peak 6-12 per crew", size=14,
         bold=True, color=ACCENT, font=FONT_BOLD)
add_lines(s, Inches(7.0), top_y + Inches(0.55), Inches(5.7), Inches(1.9),
          ["• Accenture (Global SAP · Corporate · CRM)",
           "• Software Factory (cross-stream build capacity)",
           "• Smart Metering vendor (100% delivered today)",
           "• HCL · Nuance · Infosys · OEM partners",
           "Owns: actual code · QA · operations.",
           "Avangrid = product owner only."], size=11, color=DARK)

# Bottom: Option A vs Option B
bot_y = Inches(3.85)
round_rect(s, Inches(0.6), bot_y, Inches(6.05), Inches(2.6),
           WHITE, line=SUCCESS, line_w=2)
rect(s, Inches(0.6), bot_y, Inches(6.05), Inches(0.5), SUCCESS)
text_box(s, Inches(0.7), bot_y + Inches(0.05), Inches(5.85), Inches(0.4),
         "OPTION A  ·  Stable vendor team per crew",
         size=13, bold=True, color=WHITE, font=FONT_BOLD,
         anchor=MSO_ANCHOR.MIDDLE)
add_lines(s, Inches(0.8), bot_y + Inches(0.6), Inches(5.85), Inches(2.0),
          ["Vendor X always operates with Crew Y.",
           "Vendor named as crew member.",
           "Smart Metering today = this model.",
           "+ Knowledge accumulates · accountability shared.",
           "− Harder to swap vendors · cost may rise 10-20%.",
           "− Requires SOW rewrites for stability."],
          size=11, color=DARK)

round_rect(s, Inches(6.85), bot_y, Inches(6.05), Inches(2.6),
           WHITE, line=RISK, line_w=2)
rect(s, Inches(6.85), bot_y, Inches(6.05), Inches(0.5), RISK)
text_box(s, Inches(6.95), bot_y + Inches(0.05), Inches(5.85), Inches(0.4),
         "OPTION B  ·  Vendor pool consumed ad-hoc",
         size=13, bold=True, color=WHITE, font=FONT_BOLD,
         anchor=MSO_ANCHOR.MIDDLE)
add_lines(s, Inches(7.05), bot_y + Inches(0.6), Inches(5.85), Inches(2.0),
          ["Vendor capacity = pool consumed by demand.",
           "Sarath's Corporate today = this model.",
           "+ Flexible · contract-friendly · cheaper.",
           "− Vendors don't \"own\" outcome.",
           "− Knowledge churns out · not product-centric.",
           "− Reproduces today's ad-hoc pattern."],
          size=11, color=DARK)

takeaway(s, "Decision: Option A or B materially shapes contracts, culture, and Product Manager mandate.")
footer(s, "Slide 24 vrevisada (Smart Metering vendor model) · Slide 39 (Carlos branch interactions)")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 13 — Supporting Crews + Facilitation
# ════════════════════════════════════════════════════════════════════════
s = blank()
action_title(s, "Supporting Crews absorb architects, the UX function, and Sarath's 5-6 PJMs that today have no home")

crews = [
    ("CAPABILITY", RISK,
     "Architecture Council",
     ["TBD Sr.M. SAP & CS Architecture (vacant)",
      "Corporate Tech Leads (2)",
      "Cross-stream architects",
      "Reviews scope · feasibility · dependencies"]),
    ("EXPERIENCE", RGBColor(0xE5, 0x8F, 0xB5),
     "UX & Customer Journey",
     ["UX Designer (NEW role · NOM v1.5)",
      "Heather LaJoie (cross-base bridge)",
      "CX Researcher (from Digital CoE)",
      "Owns end-to-end customer journey"]),
    ("PLATFORM", SUCCESS,
     "Shared services",
     ["SWF & Tools — Raquel · Donna",
      "DevOps + App Sec — Piotr",
      "SAP Basis — Nagi Reddy",
      "TOSCA · Gold Client · 8760 EDI · Doc Mgmt (5 apps)"]),
    ("PARTNERSHIP", RGBColor(0xB6, 0xD5, 0x4E),
     "Vendor management",
     ["Accenture · SWF · Smart Metering vendor",
      "HCL · Nuance · Infosys · SAP partners",
      "Iberdrola corporate liaison",
      "Centralizing this is a real change today"]),
    ("FACILITATION", RGBColor(0x7A, 0x66, 0xB8),
     "Coaches · process · pool",
     ["Process Improvement (NEW · NOM v1.5)",
      "Agile Coaches (from CoE Agility Chapter)",
      "Knowledge Manager (NEW)",
      "Sarath's 5-6 PJMs as shared pool"]),
]

card_y = Inches(1.2)
card_h = Inches(5.3)
card_w = (SW - Inches(1.4)) / 5
gap = Inches(0.15)

for i, (name, color, sub, lines) in enumerate(crews):
    x = Inches(0.5) + i * (card_w + gap)
    round_rect(s, x, card_y, card_w, card_h, WHITE, line=color, line_w=2)
    rect(s, x, card_y, card_w, Inches(0.65), color)
    text_box(s, x, card_y + Inches(0.05), card_w, Inches(0.3), name,
             size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             font=FONT_BOLD, anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x, card_y + Inches(0.35), card_w, Inches(0.3), sub,
             size=10, italic=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)
    add_lines(s, x + Inches(0.15), card_y + Inches(0.85),
              card_w - Inches(0.3), card_h - Inches(1.0),
              ["• " + ln for ln in lines], size=10, color=DARK)

takeaway(s, "Supporting Crews add ~10-15 FTE · most via repurposing · only UX Designer is a new external hire.")
footer(s, "unfix.com Crew Types · Avangrid_Hierarchy_Overview · NOM v1.5")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 14 — 5 Forums + 3-layer Portfolio architecture
# ════════════════════════════════════════════════════════════════════════
s = blank()
action_title(s, "Five Forums and three Portfolio Kanban layers convert today's fragmented governance into a single cadence")

# Left: 5 Forums
left_x = Inches(0.4)
left_w = Inches(5.8)
y0 = Inches(1.15)

text_box(s, left_x, y0, left_w, Inches(0.4), "5 FORUMS · cross-crew coordination",
         size=14, bold=True, color=TETSU, font=FONT_BOLD)

forums = [
    ("Regional Forum", "BRS UIL · CMP · NY (OPCOs)"),
    ("Functional Forum", "PM Council · TL Council · PJM pool"),
    ("Technological Forum", "SAP · Non-SAP · Platform stacks"),
    ("Customer Journey Forum", "M2C journey · self-service · field worker"),
    ("Product Forum", "Quarterly BRP cadence · cross-stream sync"),
]
for i, (n, desc) in enumerate(forums):
    fy = y0 + Inches(0.5) + i * Inches(1.0)
    round_rect(s, left_x, fy, left_w, Inches(0.85), WHITE, line=SUCCESS, line_w=1.5)
    text_box(s, left_x + Inches(0.15), fy + Inches(0.1),
             left_w - Inches(0.3), Inches(0.3), n,
             size=12, bold=True, color=DARK, font=FONT_BOLD)
    text_box(s, left_x + Inches(0.15), fy + Inches(0.42),
             left_w - Inches(0.3), Inches(0.4), desc,
             size=10, italic=True, color=GREY)

# Right: 3 Portfolio layers
right_x = Inches(6.4)
right_w = Inches(6.5)

text_box(s, right_x, y0, right_w, Inches(0.4),
         "3-LAYER PORTFOLIO · 11 boards/views",
         size=14, bold=True, color=TETSU, font=FONT_BOLD)

# Layer 0
ly0 = y0 + Inches(0.5)
round_rect(s, right_x, ly0, right_w, Inches(0.85), HUMAN_LIGHT, line=HUMAN, line_w=2)
text_box(s, right_x + Inches(0.15), ly0 + Inches(0.05),
         right_w - Inches(0.3), Inches(0.3),
         "L0  ·  VP Roll-up Dashboard  (1 view)",
         size=12, bold=True, color=HUMAN, font=FONT_BOLD)
text_box(s, right_x + Inches(0.15), ly0 + Inches(0.35),
         right_w - Inches(0.3), Inches(0.5),
         "Owner: Vitor + Portfolio Manager · monthly · aggregator",
         size=10, italic=True, color=DARK)

# Layer 1
ly1 = ly0 + Inches(1.05)
round_rect(s, right_x, ly1, right_w, Inches(1.4), HUMAN_LIGHT, line=HUMAN, line_w=2)
text_box(s, right_x + Inches(0.15), ly1 + Inches(0.05),
         right_w - Inches(0.3), Inches(0.3),
         "L1  ·  2 Portfolio Boards (real)  ·  cards = INITIATIVES",
         size=12, bold=True, color=HUMAN, font=FONT_BOLD)
text_box(s, right_x + Inches(0.15), ly1 + Inches(0.4),
         right_w - Inches(0.3), Inches(0.95),
         "• CS Apps  (Kumar)  ·  3 streams — M2C-SL · M2C-RO · CCC\n"
         "• IT Apps  (Carlos)  ·  5 streams — Field Ops · E&A · NOC · C&R · Iberdrola*\n"
         "Stages: Funnel → Reviewing → Analyzing → Backlog → Implementing → Done\n"
         "Shark Tank gate · monthly Portfolio Committee per board",
         size=10, color=DARK)

# Layer 2
ly2 = ly1 + Inches(1.6)
round_rect(s, right_x, ly2, right_w, Inches(2.0), HUMAN_LIGHT, line=HUMAN, line_w=2)
text_box(s, right_x + Inches(0.15), ly2 + Inches(0.05),
         right_w - Inches(0.3), Inches(0.3),
         "L2  ·  8 Crew Kanbans (real)  ·  cards = EPICS · FEATURES · STORIES",
         size=12, bold=True, color=HUMAN, font=FONT_BOLD)
text_box(s, right_x + Inches(0.15), ly2 + Inches(0.4),
         right_w - Inches(0.3), Inches(1.5),
         "1 board per Value Stream Crew · Owner: each crew's PM (NEW)\n"
         "Layout: 5 phases × 2 tracks  (P&S Compass model)\n"
         "  Phase 0 Discovery → 1 Materialization → 2 Creation → 3 Ongoing → 4 Maintenance\n"
         "  Discovery track (always-on) + Delivery track (sprint execution)\n"
         "Sub-views: Kanban · Sprint board · Roadmap · Backlog (Jira-native)",
         size=10, color=DARK)

takeaway(s, "11 boards/views · single cadence per crew replaces today's scattered governance.")
footer(s, "P&S Compass model (nWOW v1.5 Slide 44) · CX CoE Jira playbook · BRP Q1/Q2 2026")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 15 — Demand Management capabilities (14)
# ════════════════════════════════════════════════════════════════════════
s = blank()
action_title(s, "Demand Management adopts single front door per crew, weighted Pain Score, and 20%-capped Expedite Class of Service")

# 14 capabilities in a 3-column grid
caps = [
    ("Demand intake channels",
     "1 front door per crew · decommission Email/Teams/Word"),
    ("Demand registration & logging",
     "Auto Jira issue · immutable fields · audit trail"),
    ("Channel consolidation",
     "≥90% via single channel by month 6 · zero alternative formal channels"),
    ("Standard intake formats",
     "Confluence DoR template · mandatory fields block submission"),
    ("Demand classification",
     "Type · Size T-shirt · Urgency = COS (separate axis)"),
    ("Run/Change/Reg/Innovation",
     "4 Classes of Service · capped: Reg ≤20% · Innovation ≥10%"),
    ("Business value definition",
     "H(x) = (BV·w₁ + UX·w₂ + Safety·w₃) / E · Sponsor signs off"),
    ("Acceptance criteria",
     "Mandatory in DoR · Given/When/Then · immutable post-prio"),
    ("Technical feasibility",
     "Architect rotation · ≤2h time-box · advisory not veto"),
    ("Triage & validation",
     "Gate 1 weekly 30-min · DM + 2 PMs + 1 Architect · SLA ≤5d"),
    ("Demand routing",
     "Routing key = crew · auto-routed by form · ≤5% mis-routing"),
    ("Urgent / exception",
     "Expedite ≤20% cap · breach = escalation · P0 to Service Desk"),
    ("Demand visibility",
     "Single dashboard public · per-requester portal · auto-refresh ≤1h"),
    ("Duplicate / overlap",
     "Form similarity search · weekly dedup · overlaps to Gate 2"),
]

cap_y = Inches(1.1)
cap_h = Inches(5.3)
ncols = 5
nrows = 3
col_w = (SW - Inches(1.0)) / ncols
row_h = cap_h / nrows
gap_x = Inches(0.05)
gap_y = Inches(0.05)

for i, (title, body) in enumerate(caps):
    row = i // ncols
    col = i % ncols
    x = Inches(0.5) + col * (col_w + gap_x)
    y = cap_y + row * (row_h + gap_y)
    round_rect(s, x, y, col_w - gap_x, row_h - gap_y,
               WHITE, line=SUCCESS, line_w=1.5)
    rect(s, x, y, col_w - gap_x, Inches(0.4), SUCCESS)
    text_box(s, x + Inches(0.1), y + Inches(0.03),
             col_w - gap_x - Inches(0.2), Inches(0.34),
             title, size=10, bold=True, color=WHITE, font=FONT_BOLD,
             anchor=MSO_ANCHOR.MIDDLE)
    text_box(s, x + Inches(0.1), y + Inches(0.45),
             col_w - gap_x - Inches(0.2), row_h - Inches(0.55),
             body, size=9, color=DARK)

takeaway(s, "14 capabilities × 8 crews · single front-door principle anchored in Slide 36/59 vrevisada (Ana-validated).")
footer(s, "2604 AS-IS revisada Slide 36/59 · NOM v1.5 · Renewables DTO Class-of-Service · CX CoE H(x) formula")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 16 — Demand Management — Roles + Tools + Governance + Performance
# ════════════════════════════════════════════════════════════════════════
s = blank()
action_title(s, "Roles, Tools, Governance, and Performance complete the Demand Management framework")

# 4 columns
col_y = Inches(1.1)
col_h = Inches(5.4)
col_w = (SW - Inches(1.0)) / 4
gap = Inches(0.1)

categories = [
    ("ROLES (4)", TEAL, [
        ("Business requestors / stakeholders",
         "CSP · Networks Business · Iberdrola corp · BRS"),
        ("Product or demand managers",
         "8 PMs NEW · 1 Demand Mgr NEW · Heather bridge"),
        ("Architects or technical reviewers",
         "Capability Crew · Architecture Council · 1-of-3 rotation"),
        ("Triage / intake committee",
         "Gate 1 weekly · DM + 2 PMs + 1 Architect"),
    ]),
    ("TOOLS (5)", ACCENT, [
        ("Demand capture",
         "ServiceNow form per crew (or Jira Intake direct)"),
        ("Tracking & backlog",
         "Jira · Initiative→Epic→Feature→Story · Crew Kanban"),
        ("Communication",
         "Triage outcomes auto-published · per-requester portal"),
        ("Documentation",
         "Confluence DoR · COS taxonomy · product brief"),
        ("Reporting / visibility",
         "Demand Dashboard public · refresh ≤1h"),
    ]),
    ("GOVERNANCE (6)", HUMAN, [
        ("Demand review forums",
         "Gate 1 weekly · Gate 2 monthly · Layer 0 monthly · BRP quarterly"),
        ("Decision practices",
         "Accept · Send-back · Reject · Redirect · all logged"),
        ("Escalation",
         "Cross-crew → Forum · Cross-portfolio → VP Sync"),
        ("Intake guidelines",
         "DoR + COS + caps · Confluence single source · 6-month review"),
        ("Business alignment",
         "BRP quarterly · Customer Journey Forum · Heather bridge"),
        ("Backlog oversight",
         "Demand Mgr owns dashboard · aging > X = escalation"),
    ]),
    ("PERFORMANCE (7)", RISK, [
        ("Volume / inflow trends",
         "Per crew · per COS · per OPCO · per source area"),
        ("Demand quality",
         "DoR pass-rate ≥70% · re-submission rate · AC quality"),
        ("Triage lead time",
         "Median ≤5 business days · 90th percentile · per COS"),
        ("% formal channels",
         "Target ≥90% by month 6 · today: unmeasured (KB gap)"),
        ("% urgent / bypassed",
         "Expedite ≤20% cap · auto-downgrade rate"),
        ("Duplication / overlap",
         "Dedup rate · overlap escalations to Gate 2"),
        ("Visibility & traceability",
         "% requesters who find own status · target 100%"),
    ]),
]

for i, (name, color, items) in enumerate(categories):
    x = Inches(0.5) + i * (col_w + gap)
    round_rect(s, x, col_y, col_w - gap, col_h, WHITE, line=color, line_w=2)
    rect(s, x, col_y, col_w - gap, Inches(0.5), color)
    text_box(s, x, col_y + Inches(0.05), col_w - gap, Inches(0.4), name,
             size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font=FONT_BOLD)
    item_y = col_y + Inches(0.6)
    item_h = (col_h - Inches(0.6)) / len(items)
    for j, (t, body) in enumerate(items):
        ix = x + Inches(0.1)
        iy = item_y + j * item_h
        text_box(s, ix, iy + Inches(0.03), col_w - gap - Inches(0.2),
                 Inches(0.25), t, size=10, bold=True, color=DARK, font=FONT_BOLD)
        text_box(s, ix, iy + Inches(0.28), col_w - gap - Inches(0.2),
                 item_h - Inches(0.3), body, size=9, color=GREY, italic=True)

takeaway(s, "36 items defined · KB-grounded where citable · ready for workshop validation.")
footer(s, "NOM v1.5 · Renewables DTO · BRP P&S · KB synthesis · Apr 2026")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 17 — Product Manager gap (8 NEW seats)
# ════════════════════════════════════════════════════════════════════════
s = blank()
action_title(s, "Eight Product Manager seats are NEW — today's \"PM\" is everywhere a Project Manager, not a Product Manager")

# Top: PJM vs PM distinction
top_y = Inches(1.1)
round_rect(s, Inches(0.4), top_y, Inches(6.3), Inches(2.4),
           HUMAN_LIGHT, line=HUMAN, line_w=2)
text_box(s, Inches(0.55), top_y + Inches(0.1), Inches(6.0), Inches(0.4),
         "TODAY · Project Manager (PJM)", size=14, bold=True,
         color=HUMAN, font=FONT_BOLD)
add_lines(s, Inches(0.6), top_y + Inches(0.55), Inches(6.0), Inches(1.85),
          ["Owns time-bounded scope · delivery focus",
           "Examples (KB):",
           "  · Iram Chowdhury — Smart Metering Champion / PJM",
           "  · 5-6 PJMs under Sarath in Corporate Apps",
           "  · Travis · PMO Lead",
           "  · Vihn Nguyen · Director IT Program Delivery",
           "Manages: Gantt · scope · timeline · vendor SOW"], size=11, color=DARK)

round_rect(s, Inches(7.0), top_y, Inches(5.9), Inches(2.4),
           RGBColor(0xFF, 0xF3, 0xCD), line=ACCENT, line_w=2)
text_box(s, Inches(7.15), top_y + Inches(0.1), Inches(5.6), Inches(0.4),
         "TO-BE · Product Manager (NEW)", size=14, bold=True,
         color=ACCENT, font=FONT_BOLD)
add_lines(s, Inches(7.2), top_y + Inches(0.55), Inches(5.6), Inches(1.85),
          ["Owns continuous product evolution · outcome focus",
           "Today: ZERO Product Managers in CS Apps or IT Apps",
           "Heather LaJoie has Director of Product Mgmt title",
           "  but on the BUSINESS side (2 FTE only)",
           "Manages: backlog priority · OKRs · roadmap · users",
           "Decides what gets built next · says NO often."], size=11, color=DARK)

# Bottom: 8 PM seats
bot_y = Inches(3.7)
text_box(s, Inches(0.4), bot_y, SW - Inches(0.8), Inches(0.4),
         "EIGHT PRODUCT MANAGER SEATS · candidates · gap to fill",
         size=14, bold=True, color=TETSU, font=FONT_BOLD)

pm_seats = [
    ("M2C — Service Lifecycle", "Iram (re-skill from PJM)", "REASONABLE"),
    ("M2C — Revenue Operations", "Hire OR Ivelisse re-skill", "TBD"),
    ("Customer Care & CCC", "Vacant Sr.M. CRM/CCC hire", "DEPENDS ON HIRE"),
    ("Field Operations", "Andrew (re-skill from Champion)", "REASONABLE"),
    ("Engineering & Asset Mgmt", "George (dual today; needs split)", "RISK · BURDEN"),
    ("Network Operations & Control", "TBD · likely external hire", "EXTERNAL HIRE"),
    ("Corporate & Regulatory", "Sarath (re-skill from Sr.M.)", "REASONABLE"),
    ("Iberdrola Corp Services", "Vanessa (re-skill from Sr.M.)", "REASONABLE"),
]

pm_w = (SW - Inches(1.0)) / 4
pm_h = Inches(0.7)
for i, (crew, candidate, status) in enumerate(pm_seats):
    row = i // 4
    col = i % 4
    x = Inches(0.5) + col * pm_w
    y = bot_y + Inches(0.5) + row * (pm_h + Inches(0.1))
    color = SUCCESS if status == "REASONABLE" else (
        ACCENT if "HIRE" in status or "RISK" in status else GREY)
    round_rect(s, x, y, pm_w - Inches(0.1), pm_h, WHITE, line=color, line_w=1.5)
    text_box(s, x + Inches(0.1), y + Inches(0.03),
             pm_w - Inches(0.3), Inches(0.25), crew,
             size=10, bold=True, color=DARK, font=FONT_BOLD)
    text_box(s, x + Inches(0.1), y + Inches(0.28),
             pm_w - Inches(0.3), Inches(0.2), candidate,
             size=9, color=GREY, italic=True)
    rect(s, x + Inches(0.1), y + pm_h - Inches(0.18),
         pm_w - Inches(0.3), Inches(0.16), color)
    text_box(s, x + Inches(0.1), y + pm_h - Inches(0.18),
             pm_w - Inches(0.3), Inches(0.16), status,
             size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font=FONT_BOLD)

takeaway(s, "Five re-skills + 1 external hire + 2 dependent on Sr. Manager vacancies = realistic 12-month staffing path.")
footer(s, "Slide 24 vrevisada (Iram=PJM) · Avangrid_Hierarchy_Overview · NOM v1.5 role taxonomy")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 18 — Open decisions & missing interviews
# ════════════════════════════════════════════════════════════════════════
s = blank()
action_title(s, "Four open decisions and five missing interviews unblock the design before the pilot")

# Left: open decisions
left_x = Inches(0.4)
left_w = Inches(6.3)
ld_y = Inches(1.15)
text_box(s, left_x, ld_y, left_w, Inches(0.4),
         "FOUR OPEN DECISIONS", size=14, bold=True, color=RISK, font=FONT_BOLD)

decisions = [
    ("Vendor delivery model",
     "Option A (stable team per crew · Smart Metering today) vs Option B (vendor pool ad-hoc · Sarath today). Affects contracts + culture."),
    ("Sr. Manager vacancies (Kumar)",
     "CRM/CCC + SAP & CS Architecture · fill with PM-profile or technical-manager? Determines if Customer Care VS Crew is viable."),
    ("Sarath's 5-6 PJMs",
     "Shared pool vs distributed 1 per Crew. Decides if Corporate is its own crew or an architecture-overflow lane."),
    ("Intake channel",
     "ServiceNow form vs Jira Intake direct vs hybrid. Pending channel volume + Travis/Vihn interviews."),
]

for i, (title, body) in enumerate(decisions):
    dy = ld_y + Inches(0.5) + i * Inches(1.15)
    round_rect(s, left_x, dy, left_w, Inches(1.05), WHITE, line=RISK, line_w=1.5)
    rect(s, left_x, dy, Inches(0.45), Inches(1.05), RISK)
    text_box(s, left_x + Inches(0.1), dy + Inches(0.4), Inches(0.3), Inches(0.3),
             str(i+1), size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font=FONT_BOLD)
    text_box(s, left_x + Inches(0.6), dy + Inches(0.1),
             left_w - Inches(0.7), Inches(0.3), title,
             size=12, bold=True, color=DARK, font=FONT_BOLD)
    text_box(s, left_x + Inches(0.6), dy + Inches(0.4),
             left_w - Inches(0.7), Inches(0.65), body,
             size=10, color=GREY, italic=True)

# Right: missing interviews
right_x = Inches(7.0)
right_w = Inches(5.9)
text_box(s, right_x, ld_y, right_w, Inches(0.4),
         "FIVE MISSING INTERVIEWS", size=14, bold=True, color=ACCENT, font=FONT_BOLD)

interviews = [
    ("Vanessa Colombo", "Sr.M. Global SAP — Iberdrola governance overlay"),
    ("George Porto · expanded", "Sr.M. GIS & OSG — Engineering/Planning users"),
    ("Sarath Garikapati · expanded", "Sr.M. Corporate — PJM pool destination"),
    ("Travis · PMO Lead", "PDD owner — intake channel decision"),
    ("Vihn Nguyen", "Director IT Program Delivery — runs PDD direct"),
]

for i, (name, why) in enumerate(interviews):
    iy = ld_y + Inches(0.5) + i * Inches(0.92)
    round_rect(s, right_x, iy, right_w, Inches(0.82),
               WHITE, line=ACCENT, line_w=1.5)
    rect(s, right_x, iy, Inches(0.45), Inches(0.82), ACCENT)
    text_box(s, right_x + Inches(0.1), iy + Inches(0.3),
             Inches(0.3), Inches(0.3), str(i+1), size=18, bold=True,
             color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font=FONT_BOLD)
    text_box(s, right_x + Inches(0.6), iy + Inches(0.1),
             right_w - Inches(0.7), Inches(0.25), name,
             size=12, bold=True, color=DARK, font=FONT_BOLD)
    text_box(s, right_x + Inches(0.6), iy + Inches(0.35),
             right_w - Inches(0.7), Inches(0.45), why,
             size=10, color=GREY, italic=True)

takeaway(s, "All four decisions are tractable — five interviews can be scheduled in 2 weeks.")
footer(s, "Avangrid_Hierarchy_Overview §7 · Apr 2026 KB gap log")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 19 — Next steps · 14-week pilot
# ════════════════════════════════════════════════════════════════════════
s = blank()
action_title(s, "A 14-week pilot with a single Value Stream Crew validates the design before scaling to all eight")

# Timeline
tl_y = Inches(1.4)
tl_h = Inches(1.2)
phases = [
    ("Weeks 1-2", "Review & Baseline", HUMAN,
     "Interviews · channel volume · vendor model decision"),
    ("Weeks 3-4", "Crew Design", HUMAN,
     "Charter pilot crew · PM hire/re-skill · Avangrid core staffed"),
    ("Weeks 5-6", "Setup", SUCCESS,
     "Portfolio Board + Crew Kanban configured · DoR template · COS"),
    ("Weeks 7-13", "Pilot Execution", ACCENT,
     "Crew operates new model · weekly Gate 1 · monthly Gate 2"),
    ("Week 14", "Retro & Scale", RISK,
     "Validate KPIs · refine playbook · approval to scale to crew #2"),
]
ph_w = (SW - Inches(1.0)) / 5
for i, (when, what, color, desc) in enumerate(phases):
    x = Inches(0.5) + i * ph_w
    round_rect(s, x, tl_y, ph_w - Inches(0.1), tl_h, WHITE, line=color, line_w=2)
    rect(s, x, tl_y, ph_w - Inches(0.1), Inches(0.4), color)
    text_box(s, x + Inches(0.1), tl_y + Inches(0.05),
             ph_w - Inches(0.3), Inches(0.3), when,
             size=11, bold=True, color=WHITE, font=FONT_BOLD)
    text_box(s, x + Inches(0.1), tl_y + Inches(0.45),
             ph_w - Inches(0.3), Inches(0.3), what,
             size=12, bold=True, color=DARK, font=FONT_BOLD)
    text_box(s, x + Inches(0.1), tl_y + Inches(0.75),
             ph_w - Inches(0.3), tl_h - Inches(0.85), desc,
             size=9, color=GREY, italic=True)
    # Arrow between phases
    if i < len(phases) - 1:
        arrow_x = x + ph_w - Inches(0.1)
        text_box(s, arrow_x, tl_y + tl_h/2 - Inches(0.15),
                 Inches(0.1), Inches(0.3), "→",
                 size=18, color=GREY, anchor=MSO_ANCHOR.MIDDLE,
                 align=PP_ALIGN.CENTER, font=FONT_BOLD)

# Pilot crew recommendation
pc_y = Inches(3.1)
round_rect(s, Inches(0.4), pc_y, SW - Inches(0.8), Inches(2.4),
           HUMAN_LIGHT, line=HUMAN, line_w=2)
text_box(s, Inches(0.6), pc_y + Inches(0.1), SW - Inches(1.2), Inches(0.4),
         "PILOT CREW RECOMMENDATION", size=14, bold=True, color=HUMAN, font=FONT_BOLD)
text_box(s, Inches(0.6), pc_y + Inches(0.5),
         SW - Inches(1.2), Inches(0.7),
         "Smart Metering / M2C — Service Lifecycle",
         size=24, bold=True, color=TETSU, font=FONT_BOLD)
add_lines(s, Inches(0.6), pc_y + Inches(1.3), SW - Inches(1.2), Inches(1.0),
          ["Why: Iram is the closest existing PM-equivalent · vendor model already stable (100% Smart Metering vendor)",
           "      · Avangrid acts as product owner today · KB-confirmed (Slide 24 vrevisada)",
           "Risk: contained · single vendor · clear users (field services + meter installers + customer)",
           "Next: validate with Vincent Lam + Iram + Heather + sponsor in week 1-2 · charter the crew formally."],
          size=11, color=DARK)

takeaway(s, "Smart Metering pilot validates Option A vendor model + Iram re-skill path · scales to remaining 7 crews.")
footer(s, "SAP CoE NWoW v2.0 (14-week Wave 1 methodology) · Slide 24 vrevisada · NOM v1.5")


# ════════════════════════════════════════════════════════════════════════
# SLIDE 20 — Appendix · KB sources
# ════════════════════════════════════════════════════════════════════════
s = blank()
action_title(s, "Appendix · Knowledge Base sources and references")

cols = [
    ("AVANGRID DOCUMENTS", TETSU, [
        "2604_AS-IS deliverable vrevisada.pptx (Ana-validated)",
        "Avangrid_Hierarchy_Overview.md",
        "Org_Chart_Digital_Products.md",
        "SAP CoE New Ways of Working v2.0 (Slide 6, 7, 11, 12)",
        "SAP IT-CS Support Model.pptx",
        "Slide 41 vrevisada · Customer Services Apps context",
        "Slide 39 vrevisada · IT Business Apps context",
    ]),
    ("REFERENCE ENGAGEMENTS", HUMAN, [
        "P&S NOM v1.5 (04.13.2026)",
        "Compass Portfolio Management Model · nWOW v1.5 Slide 44",
        "BRP Q1/Q2 2026 P&S Networks IT",
        "Project Management Playbook Slide 79",
        "Renewables DTO Final · Slide 19/27 · Lean Kanban",
        "Digital CoE Operating Model Playbook v2.5",
        "[33] NextGen 2.0 Jira Management User Guide CX",
    ]),
    ("APPLICATION ASSESSMENT", SUCCESS, [
        "Application Assessment IT Networks · 46 apps",
        "Portfolio_Landscape_Analysis_Avangrid_F1.0",
        "Initiative_Cards_Avangrid_V0.1 · Slide 22, 28, 53",
        "Networks_Technology_Roadmap_V0.1",
        "Echelon - Kick Off (SCADA workflow context)",
        "APR 2026 IT Avangrid Project Status Reports",
        "Prioritization Criteria & KPIs per Engagement.docx",
    ]),
    ("FRAMEWORK", ACCENT, [
        "unfix.com — Crew Types · Base Types · Forum Types",
        "Strongly Aligned Base definition",
        "Value Stream Crew (3-7 person rule)",
        "Capability · Experience · Platform · Partnership · Facilitation Crew",
        "Pyramid Principle (Barbara Minto)",
        "Team Topologies (Skelton/Pais)",
        "SVPG / Marty Cagan product team org",
    ]),
]

col_y = Inches(1.1)
col_h = Inches(5.4)
col_w = (SW - Inches(1.0)) / 4
gap = Inches(0.1)

for i, (name, color, items) in enumerate(cols):
    x = Inches(0.5) + i * (col_w + gap)
    round_rect(s, x, col_y, col_w - gap, col_h, WHITE, line=color, line_w=2)
    rect(s, x, col_y, col_w - gap, Inches(0.5), color)
    text_box(s, x, col_y + Inches(0.05), col_w - gap, Inches(0.4), name,
             size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE, font=FONT_BOLD)
    add_lines(s, x + Inches(0.15), col_y + Inches(0.6),
              col_w - gap - Inches(0.3), col_h - Inches(0.7),
              ["• " + it for it in items], size=9.5, color=DARK)

text_box(s, Inches(0.4), SH - Inches(0.85), SW - Inches(0.8), Inches(0.5),
         "KB stats · 22,507 chunks · 634 sources · projects: SAP Digital Products · P&S/P&T · Digital CoE · Renewables DTO · Infrastructure",
         size=10, italic=True, color=GREY, align=PP_ALIGN.CENTER)


# ════════════════════════════════════════════════════════════════════════
# Save
# ════════════════════════════════════════════════════════════════════════
out_file = OUT / "DigitalProducts_TOM_Proposal_v1.0.pptx"
prs.save(str(out_file))
print(f"✓ Saved: {out_file}")
print(f"  Slides: {len(prs.slides)}")
