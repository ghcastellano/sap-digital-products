"""
Integrated Baseline & Demand Map — Avangrid SAP Digital Products
NTT DATA engagement | Current State Diagnosis + TOM Foundations

Unified deliverable consolidating:
- Baseline Report (diagnostic findings, MECE pain points, structural issues, TOM principles)
- Demand Map (per-team drill-downs across demand lifecycle)

Grounded in Knowledge Base:
- Interview_Application_Champions.md (5 champions — GIS, Corporate, G&E, Global SAP, Smart Metering)
- Interview_Business-side lead  —  Customer Services_Kumar_IT_CS.md (IT Customer Service)
- Interview_Carlos_Applications.md (Director of Applications)
- SAP CoE New Ways of Working v2.0.pptx (Slides 3, 7, 12, 21)
- Org_Chart_Digital_Products.md + Validated Miro org chart
- Excel AS-IS canvas per team

Design tokens: NTT DATA brand (tetsu_kon_blue, human_blue, devils_butterfly accent)
Typography: Helvetica Neue (Calibri fallback for rendering compatibility)
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

# ========== NTT BRAND TOKENS ==========
NAVY = RGBColor(0x0F, 0x1C, 0x50)        # tetsu_kon_blue — primary
BLUE = RGBColor(0x67, 0x85, 0xC1)        # human_blue — secondary
ACCENT = RGBColor(0xBC, 0x43, 0x28)      # devils_butterfly — risk/negative (diagnostic)
HIGHLIGHT = RGBColor(0xE6, 0xB6, 0x00)   # plague_brown — caution/partial

WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY = RGBColor(0xF5, 0xF5, 0xF5)
MID_GRAY = RGBColor(0x95, 0xA5, 0xA6)
CHARCOAL = RGBColor(0x2C, 0x3E, 0x50)
BLACK = RGBColor(0x00, 0x00, 0x00)

# Semantic coding for maturity
GREEN_OK = RGBColor(0x2E, 0x7D, 0x32)
YELLOW_GAP = RGBColor(0xF9, 0xA8, 0x25)
RED_CRITICAL = RGBColor(0xBC, 0x43, 0x28)

FONT = "Helvetica Neue"

# Canvas: 16:9 widescreen
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# ========== HELPERS ==========

def set_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_text(slide, left, top, width, height, text, *,
             size=11, bold=False, color=CHARCOAL, font=FONT,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    tf.word_wrap = True
    tf.vertical_anchor = anchor

    lines = text.split("\n") if isinstance(text, str) else text
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        # Kill default paragraph spacing to prevent vertical overflow
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        p.line_spacing = 1.1
        run = p.add_run()
        run.text = line
        run.font.name = font
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color
    return tb


def add_rect(slide, left, top, width, height, fill_color, *,
             line_color=None, line_weight=None, corner=False):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if corner else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_color
    if line_color is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_color
        if line_weight:
            shp.line.width = line_weight
    shp.shadow.inherit = False
    return shp


def add_action_title(slide, title_text):
    """Top-band action title — states a conclusion."""
    add_text(slide, Inches(0.5), Inches(0.35), Inches(12.33), Inches(0.9),
             title_text, size=20, bold=True, color=NAVY, font=FONT)
    # Subtle divider
    div = add_rect(slide, Inches(0.5), Inches(1.25), Inches(12.33), Inches(0.015),
                   LIGHT_GRAY)


def add_takeaway(slide, text, color=NAVY):
    """Bottom-third kicker box — the 'so what'."""
    box = add_rect(slide, Inches(0.5), Inches(6.55), Inches(12.33), Inches(0.55), color)
    add_text(slide, Inches(0.75), Inches(6.60), Inches(11.83), Inches(0.45),
             text, size=11, bold=False, color=WHITE, font=FONT,
             anchor=MSO_ANCHOR.MIDDLE)


def add_footer(slide, citation_text=None, page_label=None):
    """Page label only — sources removed per v1.3."""
    if page_label:
        add_text(slide, Inches(10.0), Inches(7.18), Inches(2.83), Inches(0.25),
                 page_label, size=9, color=MID_GRAY, font=FONT, align=PP_ALIGN.RIGHT)


def add_page_label(slide, label, page_num):
    """Deprecated — page label now lives in the footer."""
    pass


# ========== NAV STRIP  (parent/roadmap reference on every slide) ==========

NAV_SECTIONS = [
    ("01", "EXECUTIVE SUMMARY", "Context · Findings · Methodology"),
    ("02", "ORGANIZATION", "Org · Applications · SAP / Customer Services"),
    ("03", "DEMAND FLOW", "Master Map · Tooling · Governance · 6 drill-downs"),
    ("04", "CAPACITY & PORTFOLIO", "Interview data (to be populated)"),
    ("05", "PAIN POINTS", "MECE pains · 4 bottlenecks · Cross-branch"),
    ("06", "INSIGHTS & TOM", "7 structural issues · Risks · Principles"),
]


def add_nav_strip(slide, current_idx):
    """Thin top roadmap bar — 6 pills, current section highlighted.

    Anchored above the action title. current_idx in [0..5], or None to skip.
    """
    if current_idx is None:
        return
    strip_y = 0.08
    strip_h = 0.22
    pill_gap = 0.04
    total_w = 12.33
    n = len(NAV_SECTIONS)
    pill_w = (total_w - (n - 1) * pill_gap) / n
    for i, (num, name, _sub) in enumerate(NAV_SECTIONS):
        x = 0.5 + i * (pill_w + pill_gap)
        is_current = (i == current_idx)
        bg = NAVY if is_current else LIGHT_GRAY
        fg = WHITE if is_current else MID_GRAY
        add_rect(slide, Inches(x), Inches(strip_y), Inches(pill_w), Inches(strip_h), bg,
                 corner=True)
        add_text(slide, Inches(x), Inches(strip_y), Inches(pill_w), Inches(strip_h),
                 f"{num} · {name}",
                 size=7, bold=is_current, color=fg, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def _add_nav_to_last(prs, idx):
    """Apply nav strip to the most recently added slide."""
    add_nav_strip(prs.slides[-1], idx)


def slide_roadmap(prs):
    """Parent slide — full deck roadmap ('O Todo'): shows all consulting phases,
    where the reader is, and how each section feeds into the next.
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Deck roadmap — six sections grounded in Wave 1 Step 1: diagnose the current state, surface structural issues, ground the TOM")

    # Horizontal strip of 6 section cards
    card_w = 2.02
    card_h = 3.45
    gap_x = 0.06
    y0 = 1.55

    sections_detail = [
        {
            "num": "01",
            "name": "EXECUTIVE SUMMARY",
            "slides": "03 – 05",
            "what": "Why we are here. Executive findings & implications. Methodology (Wave 1 Step 1 sources).",
            "feeds": "Sets the lens for every finding that follows.",
        },
        {
            "num": "02",
            "name": "ORGANIZATION",
            "slides": "06 – 10",
            "what": "Org landscape (~50 FTE · 9 teams · 2 vacancies). Applications branch. SAP / CS + CMP · UIL · NY · CoE · Basis. 7+ OPCOs demanding SAP.",
            "feeds": "Names the players — drill-downs & pain points reference these teams.",
        },
        {
            "num": "03",
            "name": "DEMAND FLOW",
            "slides": "11 – 19",
            "what": "AS-IS Master Map. Tooling landscape. Governance forums. Six team deep-dives across 5 lifecycle stages.",
            "feeds": "Ground truth — every bottleneck and insight traces back to these cells.",
        },
        {
            "num": "04",
            "name": "CAPACITY & PORTFOLIO",
            "slides": "20 – 21",
            "what": "Quantitative baseline from the five champion interviews: capacity (headcount · mix · utilization) and portfolio (projects · vendor mix · cost). Populated in Wave 1 discovery.",
            "feeds": "Turns pain points into measurable gaps — backbone of the TOM business case.",
        },
        {
            "num": "05",
            "name": "PAIN POINTS",
            "slides": "22 – 25",
            "what": "MECE pain cluster (People · Process · Tools · Governance). Four bottlenecks. Three cross-branch systemic issues.",
            "feeds": "Converts evidence into named risks — prioritized for TOM response.",
        },
        {
            "num": "06",
            "name": "INSIGHTS & TOM",
            "slides": "26 – 28",
            "what": "Seven structural issues paired with TOM responses. Five compounding risks. Seven TOM design principles grounded in prior NTT work.",
            "feeds": "Exit-ramp into Step 2 (Define Lifecycle) of Wave 1.",
        },
    ]

    for i, s in enumerate(sections_detail):
        x = 0.5 + i * (card_w + gap_x)
        add_rect(slide, Inches(x), Inches(y0), Inches(card_w), Inches(card_h), LIGHT_GRAY)
        add_rect(slide, Inches(x), Inches(y0), Inches(card_w), Inches(0.50), NAVY)
        add_text(slide, Inches(x + 0.15), Inches(y0 + 0.08), Inches(card_w - 0.3), Inches(0.18),
                 f"SECTION {s['num']}", size=7, bold=True, color=BLUE, font=FONT)
        add_text(slide, Inches(x + 0.15), Inches(y0 + 0.25), Inches(card_w - 0.3), Inches(0.25),
                 s["name"], size=10, bold=True, color=WHITE, font=FONT)
        add_text(slide, Inches(x + 0.15), Inches(y0 + 0.60), Inches(card_w - 0.3), Inches(0.22),
                 "Slides " + s["slides"], size=8, italic=True, color=ACCENT, font=FONT)
        add_text(slide, Inches(x + 0.15), Inches(y0 + 0.95), Inches(card_w - 0.3), Inches(0.22),
                 "WHAT YOU'LL SEE", size=7, bold=True, color=MID_GRAY, font=FONT)
        add_text(slide, Inches(x + 0.15), Inches(y0 + 1.20), Inches(card_w - 0.3), Inches(1.55),
                 s["what"], size=8, color=CHARCOAL, font=FONT)
        add_text(slide, Inches(x + 0.15), Inches(y0 + 2.80), Inches(card_w - 0.3), Inches(0.22),
                 "HOW IT FEEDS", size=7, bold=True, color=ACCENT, font=FONT)
        add_text(slide, Inches(x + 0.15), Inches(y0 + 3.05), Inches(card_w - 0.3), Inches(0.40),
                 s["feeds"], size=8, italic=True, color=BLUE, font=FONT)
        if i < len(sections_detail) - 1:
            arrow_x = x + card_w - 0.02
            arrow = slide.shapes.add_shape(
                MSO_SHAPE.RIGHT_ARROW,
                Inches(arrow_x), Inches(y0 + card_h / 2 - 0.10),
                Inches(0.10), Inches(0.14))
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = ACCENT
            arrow.line.fill.background()
            arrow.shadow.inherit = False

    orient_y = 5.20
    add_rect(slide, Inches(0.5), Inches(orient_y), Inches(12.33), Inches(1.15), LIGHT_GRAY)
    add_rect(slide, Inches(0.5), Inches(orient_y), Inches(0.08), Inches(1.15), ACCENT)
    add_text(slide, Inches(0.70), Inches(orient_y + 0.08), Inches(12), Inches(0.28),
             "HOW TO READ THIS DECK", size=10, bold=True, color=NAVY, font=FONT)
    add_text(slide, Inches(0.70), Inches(orient_y + 0.38), Inches(12), Inches(0.70),
             "Every slide carries a thin roadmap strip at the top — six pills, current section highlighted in navy. "
             "Read each slide bottom-up: action title states the conclusion, content is the evidence, takeaway is the \"so what.\" "
             "Cross-references use format [Slide XX].",
             size=10, color=CHARCOAL, font=FONT)

    add_takeaway(slide,
        "You are here → Section 01 (Executive Summary). Next: the organization that generates the demand we are mapping.",
        color=NAVY)
    add_footer(slide, page_label="Deck Roadmap  |  02 / 28")


# ========== SLIDES ==========

def slide_context(prs):
    """Project context — why we are here, scope, working model."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Avangrid is consolidating SAP CoE and Applications into a single Digital Products organization — demand fragmentation is the structural risk the TOM must resolve")

    # Left column: Why we are here
    add_rect(slide, Inches(0.5), Inches(1.50), Inches(0.08), Inches(4.6), NAVY)
    add_text(slide, Inches(0.75), Inches(1.50), Inches(5.8), Inches(0.32),
             "WHY WE ARE HERE", size=11, bold=True, color=NAVY, font=FONT)
    add_text(slide, Inches(0.75), Inches(1.90), Inches(5.8), Inches(4.3),
             "▸  Digital Products (formerly SAP CoE) is being re-founded under VP VP IT Business Solutions Delivery, merging Applications (Director IT Business Applications) and Customer Services Apps (Director Customer Services Apps) into one ~50-FTE organization.\n\n"
             "▸  Demand volume has grown 30–50% across SAP and Applications in the last 24 months — with no unified intake, classification, or prioritization model in place.\n\n"
             "▸  Six champion interviews and two director-level sessions (Apps Director, Business-side lead  —  Customer Services & CS Director) surfaced six disconnected intake channels and five different prioritization logics.\n\n"
             "▸  NTT DATA was engaged to diagnose the current state, design the Target Operating Model (TOM), and pilot it on one BRS project before scaling.",
             size=11, color=CHARCOAL, font=FONT)

    # Right column: Scope & working model
    add_rect(slide, Inches(7.0), Inches(1.50), Inches(0.08), Inches(4.6), ACCENT)
    add_text(slide, Inches(7.25), Inches(1.50), Inches(5.5), Inches(0.32),
             "SCOPE  &  WORKING MODEL", size=11, bold=True, color=NAVY, font=FONT)
    add_text(slide, Inches(7.25), Inches(1.90), Inches(5.5), Inches(4.3),
             "▸  Two director branches in scope: IT Business Applications (4 teams, ~26 FTE) and Customer Services Apps (5 teams, ~24 FTE, 2 Sr. Manager vacancies).\n\n"
             "▸  Five end-to-end stages assessed per team: Intake → Classification → Prioritization → Portfolio → Delivery.\n\n"
             "▸  Baseline grounded in interview evidence (6 champions + 2 directors), not secondary documents — every finding is traceable to a named source.\n\n"
             "▸  Wave 1 delivers: Baseline Report, Demand Map, Portfolio Landscape (SAP), Lifecycle v1.0, Jira Config v1.0, and a live BRS pilot.",
             size=11, color=CHARCOAL, font=FONT)

    add_takeaway(slide,
        "This deck is the Demand Map — the AS-IS current-state evidence base that feeds the TOM design in Wave 1, Step 2",
        color=NAVY)
    add_footer(slide, page_label="Project Context  |  03 / 28")


def slide_executive_summary(prs):
    """Executive summary — the governing thought and three supporting findings."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Digital Products runs nine sub-teams on six disconnected intake channels — with no shared prioritization logic and two Sr. Manager vacancies in Customer Services")

    # Governing thought banner
    add_rect(slide, Inches(0.5), Inches(1.55), Inches(12.33), Inches(0.85), NAVY)
    add_text(slide, Inches(0.8), Inches(1.65), Inches(11.7), Inches(0.68),
             "GOVERNING THOUGHT  —  Fragmentation is structural, not local. Every team has a working answer for large, PDD-funded demand; no team has one for small and medium demand at scale. The TOM must standardize the end-to-end lifecycle across both director branches.",
             size=12, bold=True, color=WHITE, font=FONT, italic=True, anchor=MSO_ANCHOR.MIDDLE)

    # Three pillar findings
    pillars = [
        ("01", "INTAKE  &  CLASSIFICATION",
         "6+ channels active (Jira, IT Now, RITM, SharePoint, Email, Word). G&E is fully email-driven; Corporate still uses Word docs for requirements.",
         "No single front door. Volumes and duplicates are invisible."),
        ("02", "PRIORITIZATION  &  PORTFOLIO",
         "Each team applies a different logic — PDD, impact × users, criticality, \"intrinsic common sense.\" Only Smart Metering uses a portfolio tool (PPM Now).",
         "Capacity cannot be reconciled across the branch. Strategic work loses to incident volume."),
        ("03", "ORGANIZATION  &  DELIVERY",
         "Two Sr. Manager vacancies in Customer Services (CRM & Contact Center Ops, SAP & CS Architecture). Delivery leans heavily on Accenture + SWF + vendor muscle.",
         "Leadership gap + vendor concentration create structural risk for scaling."),
    ]
    col_w = 4.10
    y = 2.65
    row_h = 3.50
    for i, (num, title, finding, impact) in enumerate(pillars):
        x = 0.5 + i * (col_w + 0.10)
        add_rect(slide, Inches(x), Inches(y), Inches(col_w), Inches(row_h), LIGHT_GRAY)
        add_rect(slide, Inches(x), Inches(y), Inches(col_w), Inches(0.06), ACCENT)
        add_text(slide, Inches(x + 0.25), Inches(y + 0.18), Inches(0.8), Inches(0.5),
                 num, size=22, bold=True, color=ACCENT, font=FONT)
        add_text(slide, Inches(x + 0.95), Inches(y + 0.25), Inches(col_w - 1.1), Inches(0.5),
                 title, size=11, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(x + 0.25), Inches(y + 0.90), Inches(col_w - 0.5), Inches(0.3),
                 "FINDING", size=8, bold=True, color=MID_GRAY, font=FONT)
        add_text(slide, Inches(x + 0.25), Inches(y + 1.15), Inches(col_w - 0.5), Inches(1.20),
                 finding, size=10, color=CHARCOAL, font=FONT)
        add_text(slide, Inches(x + 0.25), Inches(y + 2.40), Inches(col_w - 0.5), Inches(0.3),
                 "SO WHAT", size=8, bold=True, color=ACCENT, font=FONT)
        add_text(slide, Inches(x + 0.25), Inches(y + 2.65), Inches(col_w - 0.5), Inches(0.85),
                 impact, size=10, italic=True, color=BLUE, font=FONT)

    add_takeaway(slide,
        "The rest of this deck walks the evidence per team — then consolidates bottlenecks and structural insights into the TOM scope",
        color=NAVY)
    add_footer(slide, page_label="Executive Summary  |  04 / 28")


def slide_methodology(prs):
    """Methodology — Wave 1 five-step approach with focus on Step 1."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    add_action_title(slide,
        "This Demand Map is the output of Step 1 (Enablement) — grounding the baseline through champion interviews and portfolio mapping before lifecycle design")

    # Top strip: 5-step horizontal timeline (Step 1 highlighted)
    steps = [
        ("01", "Enablement", "Where are we starting from?", True),
        ("02", "Define Lifecycle", "How should we work?", False),
        ("03", "Design Tool Ecosystem", "What supports the process?", False),
        ("04", "Pilot Execution (BRS)", "Does it work in reality?", False),
        ("05", "Learn & Adjust", "What needs refinement?", False),
    ]
    step_w = 2.45
    step_y = 1.42
    for i, (num, title, question, active) in enumerate(steps):
        x = 0.5 + i * (step_w + 0.08)
        box_color = NAVY if active else LIGHT_GRAY
        text_color = WHITE if active else CHARCOAL
        sub_color = BLUE if active else MID_GRAY
        add_rect(slide, Inches(x), Inches(step_y), Inches(step_w), Inches(1.25), box_color)
        add_text(slide, Inches(x + 0.15), Inches(step_y + 0.10), Inches(step_w - 0.3), Inches(0.30),
                 f"Step {num}", size=10, bold=True, color=sub_color, font=FONT)
        add_text(slide, Inches(x + 0.15), Inches(step_y + 0.40), Inches(step_w - 0.3), Inches(0.38),
                 title, size=13, bold=True, color=text_color, font=FONT)
        add_text(slide, Inches(x + 0.15), Inches(step_y + 0.80), Inches(step_w - 0.3), Inches(0.40),
                 question, size=9, italic=True, color=sub_color, font=FONT)
        if i < len(steps) - 1:
            arrow = slide.shapes.add_shape(
                MSO_SHAPE.RIGHT_ARROW,
                Inches(x + step_w - 0.03), Inches(step_y + 0.55),
                Inches(0.15), Inches(0.15))
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = ACCENT
            arrow.line.fill.background()
            arrow.shadow.inherit = False

    # Step 1 zoom — left: what we did, right: deliverables
    zoom_y = 3.10
    add_rect(slide, Inches(0.5), Inches(zoom_y), Inches(12.33), Inches(0.38), ACCENT)
    add_text(slide, Inches(0.5), Inches(zoom_y), Inches(12.33), Inches(0.38),
             "STEP 1  —  ENABLEMENT  ·  ZOOM IN",
             size=11, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Left: three activities
    activities = [
        ("Prior Materials Review",
         "Reviewed 40+ prior artifacts: SAP CoE decks, IT CS Operating Model, P&S Playbook, DTO NextGen, Org Chart, Competency Framework."),
        ("Champion Stakeholder Interviews",
         "Six 1:1 champion interviews (GIS, Corporate, G&E, Global SAP, Smart Metering) + two director sessions (Apps Director; Business-side lead  —  Customer Services & CS Director). Captured intake channels, tools, roles, pain points and aspirations per team."),
        ("Portfolio Assessment",
         "Mapped demand typology per team across the 5 lifecycle stages. Reconciled interview evidence with the RFI workbook and the validated Miro board."),
    ]
    y = 3.62
    add_text(slide, Inches(0.5), Inches(y), Inches(7.5), Inches(0.30),
             "WHAT WE DID", size=10, bold=True, color=NAVY, font=FONT)
    y += 0.32
    for title, desc in activities:
        add_rect(slide, Inches(0.5), Inches(y), Inches(0.08), Inches(0.78), NAVY)
        add_text(slide, Inches(0.70), Inches(y), Inches(7.1), Inches(0.25),
                 title, size=11, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(0.70), Inches(y + 0.26), Inches(7.1), Inches(0.55),
                 desc, size=10, color=CHARCOAL, font=FONT)
        y += 0.85

    # Right: deliverables
    add_text(slide, Inches(8.5), Inches(3.62), Inches(4.5), Inches(0.30),
             "DELIVERABLES OF STEP 1", size=10, bold=True, color=NAVY, font=FONT)
    add_rect(slide, Inches(8.5), Inches(3.95), Inches(4.33), Inches(2.25), LIGHT_GRAY)
    add_rect(slide, Inches(8.5), Inches(3.95), Inches(0.08), Inches(2.25), ACCENT)
    deliverables = [
        "▪  Baseline Report",
        "▪  Demand Map  (this deck)",
        "▪  Project Portfolio Landscape — SAP",
        "▪  Champion Interview Log (6 interviews)",
        "▪  Cross-Project Comparison (P&S / DTO / Digital CoE)",
    ]
    add_text(slide, Inches(8.75), Inches(4.10), Inches(4.0), Inches(2.0),
             "\n".join(deliverables), size=11, color=CHARCOAL, font=FONT)

    add_takeaway(slide,
        "The Demand Map is the evidence backbone — Step 2 (Define Lifecycle) consumes these findings to draft the Lifecycle v1.0, Demand Typology and Roles & RACI",
        color=NAVY)
    add_footer(slide, page_label="Methodology  |  05 / 28")


def slide_apps_branch_overview(prs):
    """Applications branch overview — nature of work, apps, priorities, interactions."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Applications runs Avangrid's non-SAP operational backbone — four teams supporting Electric & Gas, GIS, Corporate and Iberdrola Global SAP")

    # Branch band
    add_rect(slide, Inches(0.5), Inches(1.32), Inches(12.33), Inches(0.38), BLUE)
    add_text(slide, Inches(0.5), Inches(1.32), Inches(12.33), Inches(0.38),
             "BRANCH 1  —  IT BUSINESS APPLICATIONS  ·  Director: Director IT Business Applications  ·  ~26 FTE  ·  4 teams",
             size=11, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 4 cards: Nature of Work | Applications | Priorities | Interactions
    cards = [
        ("NATURE OF WORK",
         "Operational continuity + regulated change",
         [
             "▪  Keep the lights on — break/fix is the dominant pattern (GIS, G&E, Corporate).",
             "▪  Large demand planned yearly via PDD; small/medium demand handled ad-hoc.",
             "▪  Regulatory and employee-facing systems (Corporate) run under continuous compliance pressure.",
             "▪  Version upgrades (Projectwise, Microstation) gate-kept by Third-Party Risk assessment."
         ]),
        ("APPLICATIONS IN SCOPE",
         "Four technology domains under Apps Director",
         [
             "▪  GIS & OSG  —  SAP HANA, ARCFM, GIS web apps, data replication, SWF platform (GIS Sr. Manager / Champion, ~6 FTE).",
             "▪  Electric & Gas Ops  —  70–90 apps, android/tablets, doc viewer, 811 gas lines (Sr. Manager + G&E Tech Lead, ~7 FTE).",
             "▪  Corporate Applications  —  5-6 PMs, 2 tech leads, corporate + network apps, SWF tooling (Corporate Apps Manager + Corporate Champion, ~8 FTE).",
             "▪  Global SAP (Iberdrola)  —  FI, PO, Budget — Spain-side SAP module support (Global SAP Manager + Global SAP Champion, ~5 FTE)."
         ]),
        ("PRIORITIES TODAY",
         "Break/fix first, regulatory second, strategic last",
         [
             "▪  GIS: keep the lights on; team & SWF resource management; future technology roadmap.",
             "▪  Corporate: regulatory and employee-facing systems — always priority; continuous new requirements.",
             "▪  G&E: ProjectWise + Microstation upgrades; third-party risk interface; business case approvals (slow).",
             "▪  Global SAP: incident volume dominates — strategic projects slip; 1-year hiring lag amplifies load."
         ]),
        ("INTERACTIONS  &  WAY OF WORKING",
         "Inside-IT + vendor + business triangulation",
         [
             "▪  Inside Applications: GIS ↔ G&E (shared SAP/GIS data); Corporate ↔ all teams (shared PMs).",
             "▪  With SAP branch: Global SAP ↔ CS Director's SAP (Billing, FICA); P&T is the intermediary with Business.",
             "▪  With vendors: Accenture (Global SAP + Corporate), SWF / Software Factory (all teams), business vendors (G&E).",
             "▪  Tools mix: Jira (misc), IT Now (incidents), PDD (large), SharePoint + Teams + OneNote + Email (PM)."
         ]),
    ]
    col_w = 6.15
    row_h = 2.35
    gap_x = 0.13
    gap_y = 0.08
    y0 = 1.85
    for idx, (title, subtitle, bullets) in enumerate(cards):
        col = idx % 2
        row = idx // 2
        x = 0.5 + col * (col_w + gap_x)
        y = y0 + row * (row_h + gap_y)
        add_rect(slide, Inches(x), Inches(y), Inches(col_w), Inches(row_h), LIGHT_GRAY)
        add_rect(slide, Inches(x), Inches(y), Inches(col_w), Inches(0.04), BLUE)
        add_text(slide, Inches(x + 0.22), Inches(y + 0.12), Inches(col_w - 0.4), Inches(0.28),
                 title, size=11, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(x + 0.22), Inches(y + 0.40), Inches(col_w - 0.4), Inches(0.28),
                 subtitle, size=9, italic=True, color=BLUE, font=FONT)
        add_text(slide, Inches(x + 0.22), Inches(y + 0.72), Inches(col_w - 0.4), Inches(row_h - 0.82),
                 "\n".join(bullets), size=10, color=CHARCOAL, font=FONT)

    add_takeaway(slide,
        "Applications carries the structural legacy of Avangrid's IT stack — fragmentation here maps directly to the 4 intake bottlenecks surfaced later in this deck",
        color=NAVY)
    add_footer(slide, page_label="Applications Branch Overview  |  07 / 28")


def slide_sap_branch_overview(prs):
    """Customer Services Apps (SAP branch) overview — CMP / UIL / Digital CoE / Infra."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Customer Services Apps serves 7 OPCOs across three BRS groups — with SAP (Billing, CRM, FICA, Metering) as the operational core and two Sr. Manager vacancies")

    # Branch band
    add_rect(slide, Inches(0.5), Inches(1.32), Inches(12.33), Inches(0.38), NAVY)
    add_text(slide, Inches(0.5), Inches(1.32), Inches(12.33), Inches(0.38),
             "BRANCH 2  —  CUSTOMER SERVICES APPS  ·  Director: Director Customer Services Apps  ·  ~24 FTE  ·  5 teams (2 Sr. Manager vacancies)",
             size=11, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 4 cards: Nature of Work | BRS / OPCO Landscape | Teams | Way of Working
    cards = [
        ("NATURE OF WORK",
         "Customer-facing SAP + metering + contact center",
         [
             "▪  SAP (Avangrid US) — Billing, FICA, CRM, Metering — distinct from Global SAP (Iberdrola Spain).",
             "▪  Demand split across BRS (business), PMO (Travis — large projects) and IT-funded (CS Director's teams).",
             "▪  Incidents + small/medium demand dominate volume; large projects run via annual PDD.",
             "▪  Software Factory (SWF) executes; Accenture supports incidents + projects."
         ]),
        ("BRS  /  OPCO LANDSCAPE",
         "3 BRS groups representing 7 operating companies",
         [
             "▪  BRS UIL  —  United Illuminating, SCG, CNG, BGC (Connecticut cluster).",
             "▪  BRS CMP  —  Central Maine Power (Maine operations).",
             "▪  BRS NY  —  NYSEG + RGE (New York cluster).",
             "▪  Plus Digital CoE (customer apps / websites) and Call Center (contact interactions — CS Director).",
             "▪  Each OPCO competes for the same IT capacity — multi-OPCO prioritization is a daily pressure."
         ]),
        ("TEAMS UNDER KUMAR",
         "5 sub-teams, heterogeneous maturity",
         [
             "▪  Sr. Manager  —  FICA / Billing  —  FICA, Billing & Collection (~8 FTE).",
             "▪  Sr. Manager  —  Device & Metering  —  Device & Metering (Smart Metering champion / PM sits here, ~8 FTE).",
             "▪  Manager  —  SAP Basis  —  SAP Basis (platform / infrastructure, ~2 FTE).",
             "▪  TBD (vacant)  —  Sr. Manager CRM & Contact Center Ops (~4 FTE) — leadership gap.",
             "▪  Open (vacant)  —  Sr. Manager SAP & CS Architecture (~2 FTE) — leadership gap."
         ]),
        ("WAY OF WORKING  &  INTERACTIONS",
         "Fragmented intake + vendor-heavy delivery",
         [
             "▪  Small/medium demand via Jira; large via PPM Now (only Smart Metering consistently).",
             "▪  Incidents: IT Now; Minor requests: RITM; Major projects: SharePoint + Teams.",
             "▪  Interactions: CS Director ↔ Travis (PMO for large projects); ↔ BRS leads (business demand); ↔ Global SAP Manager (Global SAP).",
             "▪  Vendor concentration: Smart Metering is 100% vendor-delivered (requirements + code); other teams rely on Accenture + SWF."
         ]),
    ]
    col_w = 6.15
    row_h = 2.35
    gap_x = 0.13
    gap_y = 0.08
    y0 = 1.85
    for idx, (title, subtitle, bullets) in enumerate(cards):
        col = idx % 2
        row = idx // 2
        x = 0.5 + col * (col_w + gap_x)
        y = y0 + row * (row_h + gap_y)
        add_rect(slide, Inches(x), Inches(y), Inches(col_w), Inches(row_h), LIGHT_GRAY)
        add_rect(slide, Inches(x), Inches(y), Inches(col_w), Inches(0.04), NAVY)
        add_text(slide, Inches(x + 0.22), Inches(y + 0.12), Inches(col_w - 0.4), Inches(0.28),
                 title, size=11, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(x + 0.22), Inches(y + 0.40), Inches(col_w - 0.4), Inches(0.28),
                 subtitle, size=9, italic=True, color=BLUE, font=FONT)
        add_text(slide, Inches(x + 0.22), Inches(y + 0.72), Inches(col_w - 0.4), Inches(row_h - 0.82),
                 "\n".join(bullets), size=10, color=CHARCOAL, font=FONT)

    add_takeaway(slide,
        "Customer Services Apps carries both business-facing complexity (3 BRS × 7 OPCOs) and structural gaps (2 Sr. Manager vacancies, vendor-heavy delivery) — the TOM must address both",
        color=NAVY)
    add_footer(slide, page_label="Customer Services Branch Overview  |  08 / 28")


def slide_sap_landscape_detail(prs):
    """SAP landscape detail — CMP, UIL, BRS NY, Digital CoE, SAP Basis."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    add_action_title(slide,
        "The SAP branch serves five distinct business landscapes — three regional BRS groups plus Digital CoE and SAP Basis — each with its own demand profile and governance cadence")

    # Branch band
    add_rect(slide, Inches(0.5), Inches(1.32), Inches(12.33), Inches(0.32), NAVY)
    add_text(slide, Inches(0.5), Inches(1.32), Inches(12.33), Inches(0.32),
             "SAP  /  CUSTOMER SERVICES APPS  —  LANDSCAPE BREAKDOWN",
             size=10, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 5 cards — 3 BRS groups, Digital CoE, SAP Basis
    # Layout: 3 on top row, 2 on bottom row (centered)
    landscapes = [
        {
            "code": "BRS UIL",
            "title": "Connecticut cluster",
            "accent": BLUE,
            "scope": "UI, SCG, CNG, BGC  —  4 OPCOs",
            "bullets": [
                "▪  United Illuminating (UI) — CT electric distribution.",
                "▪  Southern Connecticut Gas (SCG).",
                "▪  Connecticut Natural Gas (CNG).",
                "▪  Berkshire Gas Company (BGC) — Western MA gas.",
                "▪  Customer services SAP footprint: Billing, FICA, CRM.",
                "▪  VP Customer Service → BRS UIL lead.",
            ],
        },
        {
            "code": "BRS CMP",
            "title": "Maine",
            "accent": BLUE,
            "scope": "Central Maine Power  —  1 OPCO",
            "bullets": [
                "▪  Central Maine Power (CMP) — ME electric distribution.",
                "▪  Dedicated BRS — distinct regulatory cadence vs. UIL and NY.",
                "▪  Customer services SAP footprint: Billing, FICA, CRM.",
                "▪  VP Customer Service → BRS CMP lead.",
                "▪  Annual PDD demand process + small/medium via Jira.",
            ],
        },
        {
            "code": "BRS NY",
            "title": "New York cluster",
            "accent": BLUE,
            "scope": "NYSEG, RGE  —  2 OPCOs",
            "bullets": [
                "▪  NYSEG — upstate NY electric & gas.",
                "▪  RGE — Rochester Gas & Electric.",
                "▪  Customer services SAP footprint: Billing, FICA, CRM, Metering (Smart Metering heaviest here).",
                "▪  VP Customer Service NY → BRS NY lead.",
                "▪  Shared AMI / metering program drives device & metering volume.",
            ],
        },
        {
            "code": "DIGITAL CoE",
            "title": "Customer digital experience",
            "accent": ACCENT,
            "scope": "Customer-facing apps, websites, mobile  —  CS Director directly",
            "bullets": [
                "▪  CS Director owns Digital CoE directly (customer interaction layer).",
                "▪  Develops solutions for the customer: portal, mobile apps, websites.",
                "▪  Interfaces with Contact Center + Call Center (customer interaction).",
                "▪  Runs separately from BRS-driven back-office demand — closer to CX journeys than to SAP modules.",
                "▪  Connects to P&T Process & Technology (field operators) and Renewables CoE patterns already mapped.",
            ],
        },
        {
            "code": "SAP BASIS",
            "title": "Platform / horizontal layer",
            "accent": MID_GRAY,
            "scope": "Manager  —  SAP Basis (~2 FTE)",
            "bullets": [
                "▪  Platform + infrastructure services supporting all SAP modules (FICA, Billing, CRM, Metering).",
                "▪  Shared service across BRS UIL / CMP / NY — horizontal layer.",
                "▪  On-prem = IT owned (platform side).",
                "▪  Version upgrades, environment moves, third-party risk interface.",
                "▪  Level-1 / small regulatory demands (Jira as intake + managing, forms standard) surface here before routing.",
            ],
        },
    ]

    # Top row — 3 BRS cards
    card_w = 4.10
    card_h = 2.55
    gap_x = 0.10
    y_top = 1.80
    for i in range(3):
        x = 0.5 + i * (card_w + gap_x)
        _sap_landscape_card(slide, x, y_top, card_w, card_h, landscapes[i])

    # Bottom row — Digital CoE + SAP Basis (centered)
    y_bot = y_top + card_h + 0.12
    # Use a slightly wider width for the two bottom cards
    card_w_bot = 6.20
    gap_bot = 0.13
    # Center the two cards horizontally
    total_bot_w = 2 * card_w_bot + gap_bot
    x_start_bot = (13.333 - total_bot_w) / 2
    for j, k in enumerate([3, 4]):
        x = x_start_bot + j * (card_w_bot + gap_bot)
        _sap_landscape_card(slide, x, y_bot, card_w_bot, card_h - 0.20, landscapes[k])

    add_takeaway(slide,
        "Any TOM recommendation on Customer Services demand must respect three regional cadences (UIL, CMP, NY), one customer-facing track (Digital CoE), and one horizontal platform (SAP Basis) — single model, five consumption patterns",
        color=NAVY)
    add_footer(slide, page_label="SAP Landscape Detail  |  09 / 28")


def slide_sap_opco_demand(prs):
    """SAP serves 7+ business OPCOs through heterogeneous channels."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    add_action_title(slide,
        "SAP serves 7+ business OPCOs through heterogeneous channels — BRS is the backbone, but every other OPCO has its own entry point")

    # Table geometry
    col_widths = [2.10, 2.70, 3.20, 1.95, 2.38]
    col_x = [0.5]
    for w_ in col_widths[:-1]:
        col_x.append(col_x[-1] + w_)
    headers = ["OPCO / AREA", "PRIMARY CHANNEL", "DEMAND PATTERN", "CADENCE", "WHO PRIORITIZES"]

    header_y = 1.40
    row_h = 0.58
    # Header row
    add_rect(slide, Inches(0.5), Inches(header_y), Inches(12.33), Inches(0.40), NAVY)
    for i, h in enumerate(headers):
        add_text(slide, Inches(col_x[i] + 0.12), Inches(header_y), Inches(col_widths[i]), Inches(0.40),
                 h, size=9, bold=True, color=WHITE, font=FONT, anchor=MSO_ANCHOR.MIDDLE)

    rows = [
        ("BRS — CMP / NY",      BLUE,         "BSS log → Jira (intake, all level 2)",                  "Low-to-critical criteria set by Business",                     "Continuous",       "Business shares priority"),
        ("BRS — UIL",           BLUE,         "ITNow",                                                 "Day-to-day issues specific to UIL",                            "Continuous",       "Business"),
        ("Digital CoE",         NAVY,         "Digital PMs (SAP functionals from CS Director's team)",       "30% BRS · 70% Internal IT · Big Room Planning visibility",     "Quarterly BRP",    "Heather (first layer)"),
        ("Data & Analytics",    ACCENT,       "Direct to team (Abdulah · Meenal · Nelly)",             "\"I need it yesterday\" · regulatory-heavy · no formal specs", "Bi-weekly",        "No formal gate"),
        ("Infrastructure",      GREEN_OK,     "Direct to department",                                  "Major maintenance · upgrades",                                 "Weekly  (Sundays)","Infrastructure Leaders"),
        ("Regulatory",          RED_CRITICAL, "Managers direct · Travis PMs",                          "URGENCY #1 — overrides everything",                            "Ad-hoc",           "Auto-priority by label"),
        ("Procurement",         MID_GRAY,     "Procurement liaison",                                   "License renewal · new vendor · ERP fixes",                     "Monthly",          "Procurement + IT"),
        ("OT  (Op. Technology)",YELLOW_GAP,   "Own model · sometimes Jira tickets",                    "Metering-specific · per-OPCO variance",                        "Annual planning",  "No standard governance"),
    ]

    y = header_y + 0.44
    for i, (opco, color, channel, pattern, cadence, prioritizer) in enumerate(rows):
        bg = LIGHT_GRAY if i % 2 == 0 else WHITE
        add_rect(slide, Inches(0.5), Inches(y), Inches(12.33), Inches(row_h), bg)
        # Left color stripe
        add_rect(slide, Inches(0.5), Inches(y), Inches(0.08), Inches(row_h), color)
        # OPCO name
        add_text(slide, Inches(col_x[0] + 0.18), Inches(y), Inches(col_widths[0] - 0.18), Inches(row_h),
                 opco, size=10, bold=True, color=color, font=FONT, anchor=MSO_ANCHOR.MIDDLE)
        # Channel
        add_text(slide, Inches(col_x[1] + 0.12), Inches(y), Inches(col_widths[1] - 0.12), Inches(row_h),
                 channel, size=9, color=CHARCOAL, font=FONT, anchor=MSO_ANCHOR.MIDDLE)
        # Pattern
        pattern_color = RED_CRITICAL if "URGENCY" in pattern else CHARCOAL
        pattern_bold = True if "URGENCY" in pattern else False
        add_text(slide, Inches(col_x[2] + 0.12), Inches(y), Inches(col_widths[2] - 0.12), Inches(row_h),
                 pattern, size=9, color=pattern_color, bold=pattern_bold, font=FONT, anchor=MSO_ANCHOR.MIDDLE, italic=True)
        # Cadence
        add_text(slide, Inches(col_x[3] + 0.12), Inches(y), Inches(col_widths[3] - 0.12), Inches(row_h),
                 cadence, size=9, color=BLUE, font=FONT, anchor=MSO_ANCHOR.MIDDLE)
        # Prioritizer
        priority_color = RED_CRITICAL if "No" in prioritizer or "Auto" in prioritizer else CHARCOAL
        add_text(slide, Inches(col_x[4] + 0.12), Inches(y), Inches(col_widths[4] - 0.12), Inches(row_h),
                 prioritizer, size=9, color=priority_color, font=FONT, anchor=MSO_ANCHOR.MIDDLE)
        y += row_h

    add_takeaway(slide,
        "Eight distinct intake paths feed SAP — BRS is the only formalized channel; Regulatory auto-prioritizes; OT and Data & Analytics run outside the standard process",
        color=ACCENT)
    add_footer(slide, page_label="SAP Demand from OPCOs  |  10 / 28")


def _sap_landscape_card(slide, x, y, w, h, data):
    """Render one SAP landscape card."""
    add_rect(slide, Inches(x), Inches(y), Inches(w), Inches(h), LIGHT_GRAY)
    add_rect(slide, Inches(x), Inches(y), Inches(w), Inches(0.06), data["accent"])
    # Code badge
    add_rect(slide, Inches(x + 0.20), Inches(y + 0.18), Inches(1.40), Inches(0.32),
             data["accent"], corner=True)
    add_text(slide, Inches(x + 0.20), Inches(y + 0.18), Inches(1.40), Inches(0.32),
             data["code"], size=9, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Title
    add_text(slide, Inches(x + 1.70), Inches(y + 0.18), Inches(w - 1.85), Inches(0.32),
             data["title"], size=12, bold=True, color=NAVY, font=FONT,
             anchor=MSO_ANCHOR.MIDDLE)
    # Scope
    add_text(slide, Inches(x + 0.22), Inches(y + 0.58), Inches(w - 0.4), Inches(0.28),
             data["scope"], size=9, italic=True, color=BLUE, font=FONT)
    # Bullets
    add_text(slide, Inches(x + 0.22), Inches(y + 0.92), Inches(w - 0.4), Inches(h - 1.00),
             "\n".join(data["bullets"]), size=9, color=CHARCOAL, font=FONT)


def slide_1_cover(prs):
    """Cover slide — executive-grade, minimal, confident."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    set_bg(slide, NAVY)

    # Left vertical accent bar — subtle brand signal
    add_rect(slide, Inches(0.75), Inches(0.75), Inches(0.04), Inches(6), BLUE)

    # Client + engagement band at top
    add_text(slide, Inches(1.1), Inches(0.85), Inches(11.5), Inches(0.3),
             "AVANGRID  —  SAP DIGITAL PRODUCTS TRANSFORMATION",
             size=10, bold=True, color=BLUE, font=FONT)
    add_text(slide, Inches(1.1), Inches(1.2), Inches(11.5), Inches(0.3),
             "Baseline & Diagnostic Phase  •  April 2026",
             size=10, color=MID_GRAY, font=FONT)

    # Accent pip
    add_rect(slide, Inches(1.1), Inches(3.1), Inches(0.5), Inches(0.05), ACCENT)

    # Main title — dominant
    add_text(slide, Inches(1.1), Inches(3.15), Inches(11.5), Inches(1.3),
             "Baseline & Demand Map", size=52, bold=True, color=WHITE, font=FONT)

    # Subtitle — the governing thought
    add_text(slide, Inches(1.1), Inches(4.55), Inches(11.5), Inches(0.5),
             "Current State Diagnosis of Digital Products Operating Model",
             size=18, color=BLUE, font=FONT)

    # Governing thought — one-liner that frames the deck
    add_text(slide, Inches(1.1), Inches(5.20), Inches(11.5), Inches(1.4),
             "Digital Products operates as two disconnected branches with seven structural issues no single-area fix can\n"
             "resolve. This integrated deck combines demand-flow evidence (per-team drill-downs) with diagnostic findings\n"
             "(pain clusters, bottlenecks, risks) — and grounds the Target Operating Model principles that follow.",
             size=12, color=WHITE, font=FONT, italic=True)

    # Bottom metadata strip
    add_rect(slide, Inches(1.1), Inches(6.7), Inches(0.3), Inches(0.03), BLUE)
    add_text(slide, Inches(1.1), Inches(6.82), Inches(6), Inches(0.3),
             "NTT DATA  —  Consulting", size=10, bold=True, color=WHITE, font=FONT)

    add_text(slide, Inches(10.5), Inches(6.82), Inches(2.5), Inches(0.3),
             "28 slides", size=9, color=MID_GRAY, font=FONT,
             align=PP_ALIGN.RIGHT)
    add_text(slide, Inches(10.5), Inches(7.12), Inches(2.5), Inches(0.25),
             "Version 1.0  •  Integrated", size=8, color=MID_GRAY, font=FONT,
             align=PP_ALIGN.RIGHT, italic=True)


def slide_org_context(prs):
    """Organization context — the real structure behind the map."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)

    add_action_title(slide,
        "Digital Products spans two director branches (~50 FTE) across nine sub-teams — with two Sr. Manager positions currently vacant")

    # Top row: VP
    add_rect(slide, Inches(5.5), Inches(1.45), Inches(2.3), Inches(0.55), NAVY)
    add_text(slide, Inches(5.5), Inches(1.45), Inches(2.3), Inches(0.55),
             "VP — IT Business Solutions Delivery",
             size=9, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Two director boxes
    add_rect(slide, Inches(1.5), Inches(2.3), Inches(5.3), Inches(0.55), BLUE)
    add_text(slide, Inches(1.5), Inches(2.3), Inches(5.3), Inches(0.55),
             "Director IT Business Applications  (~26 FTE)",
             size=10, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    add_rect(slide, Inches(6.9), Inches(2.3), Inches(5.3), Inches(0.55), NAVY)
    add_text(slide, Inches(6.9), Inches(2.3), Inches(5.3), Inches(0.55),
             "Director Customer Services Apps  (~24 FTE)",
             size=10, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Applications branch — 4 sub-teams (FTE from validated org chart)
    apps_teams = [
        ("Sr. Manager  —  Electric & Gas Ops", "Electric & Gas Ops", "7 FTE", True, "G&E Champion (Tech Lead, 70-90 apps)"),
        ("GIS Sr. Manager / Champion", "GIS & OSG", "6 FTE", True, "GIS Sr. Manager (Champion)"),
        ("Sr. Manager  —  Corporate Applications", "Corporate Applications", "8 FTE", True, "Corporate Champion (Tech Lead)"),
        ("Manager IT Apps Global SAP", "Global SAP — Iberdrola\nFI, PO, Budget (Spain)", "5 FTE", True, "Global SAP Champion (Tech Lead)"),
    ]
    y = 3.05
    for name, scope, fte, interviewed, champion in apps_teams:
        color = GREEN_OK if interviewed else RED_CRITICAL
        add_rect(slide, Inches(1.5), Inches(y), Inches(0.12), Inches(0.62), color)
        add_rect(slide, Inches(1.65), Inches(y), Inches(5.15), Inches(0.62), LIGHT_GRAY)
        add_text(slide, Inches(1.75), Inches(y + 0.05), Inches(2.7), Inches(0.30),
                 name, size=11, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(1.75), Inches(y + 0.33), Inches(2.7), Inches(0.28),
                 scope, size=10, color=CHARCOAL, font=FONT)
        # FTE pill — bigger
        add_rect(slide, Inches(4.10), Inches(y + 0.14), Inches(0.65), Inches(0.34), BLUE,
                 corner=True)
        add_text(slide, Inches(4.10), Inches(y + 0.14), Inches(0.65), Inches(0.34),
                 fte, size=10, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        champ_label = champion
        add_text(slide, Inches(4.85), Inches(y + 0.02), Inches(1.95), Inches(0.62),
                 champ_label, size=9, color=NAVY, font=FONT, italic=True,
                 anchor=MSO_ANCHOR.MIDDLE)
        y += 0.68

    # Customer Services branch — 5 sub-teams
    # status: "interviewed" | "operational" | "org_vacancy"
    cs_teams = [
        ("Sr. Manager  —  FICA / Billing", "FICA, Billing & Collection", "8 FTE", "operational", "Customer billing operations"),
        ("TBD (vacant)", "CRM & Contact Center Ops", "4 FTE", "org_vacancy", "Org vacancy — leadership gap"),
        ("Sr. Manager  —  Device & Metering", "Device & Metering (Avangrid)", "8 FTE", "interviewed", "Smart Metering Champion (PM)"),
        ("Open (vacant)", "SAP & CS Architecture", "2 FTE", "org_vacancy", "Org vacancy — leadership gap"),
        ("Manager  —  SAP Basis", "SAP Basis", "2 FTE", "operational", "Platform / infrastructure"),
    ]
    status_colors = {
        "interviewed": GREEN_OK,
        "operational": BLUE,
        "org_vacancy": RED_CRITICAL,
    }
    y = 3.05
    for name, scope, fte, status, label in cs_teams:
        color = status_colors[status]
        add_rect(slide, Inches(6.9), Inches(y), Inches(0.12), Inches(0.50), color)
        add_rect(slide, Inches(7.03), Inches(y), Inches(5.15), Inches(0.50), LIGHT_GRAY)
        add_text(slide, Inches(7.13), Inches(y + 0.03), Inches(2.4), Inches(0.22),
                 name, size=10, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(7.13), Inches(y + 0.25), Inches(2.4), Inches(0.22),
                 scope, size=8, color=CHARCOAL, font=FONT)
        # FTE pill
        add_rect(slide, Inches(9.50), Inches(y + 0.10), Inches(0.60), Inches(0.30), BLUE,
                 corner=True)
        add_text(slide, Inches(9.50), Inches(y + 0.10), Inches(0.60), Inches(0.30),
                 fte, size=9, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        prefix = {"interviewed": "", "operational": "Scope: ", "org_vacancy": ""}[status]
        add_text(slide, Inches(10.20), Inches(y + 0.02), Inches(1.95), Inches(0.52),
                 prefix + label, size=8, color=color, font=FONT, italic=True,
                 anchor=MSO_ANCHOR.MIDDLE)
        y += 0.56

    # Legend — 3 status categories with more breathing room
    legend_y = 6.15
    add_rect(slide, Inches(0.5), Inches(legend_y), Inches(0.25), Inches(0.25), GREEN_OK)
    add_text(slide, Inches(0.85), Inches(legend_y - 0.02), Inches(3.5), Inches(0.30),
             "Area champion identified (5)", size=10, bold=True, color=CHARCOAL, font=FONT)
    add_rect(slide, Inches(4.50), Inches(legend_y), Inches(0.25), Inches(0.25), BLUE)
    add_text(slide, Inches(4.85), Inches(legend_y - 0.02), Inches(3.8), Inches(0.30),
             "Operational scope documented (2)", size=10, bold=True, color=CHARCOAL, font=FONT)
    add_rect(slide, Inches(8.80), Inches(legend_y), Inches(0.25), Inches(0.25), RED_CRITICAL)
    add_text(slide, Inches(9.15), Inches(legend_y - 0.02), Inches(4), Inches(0.30),
             "Org vacancy — leadership gap (2)", size=10, bold=True, color=CHARCOAL, font=FONT)


    add_takeaway(slide,
        "Two Sr. Manager vacancies leave Customer Services running at 60% leadership coverage — a structural risk the TOM must address before scaling",
        color=NAVY)
    add_footer(slide, page_label="Org Context  |  06 / 28")


def slide_2_how_to_read(prs):
    """How to read the map — legend + stages."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)

    add_action_title(slide,
        "This map traces every demand from source to execution across two director branches and their sub-teams")
    add_page_label(slide, "How to Read", "02 / 10")

    # Left column: the 5 stages
    add_text(slide, Inches(0.5), Inches(1.45), Inches(5.5), Inches(0.35),
             "THE 5 END-TO-END STAGES", size=10, bold=True, color=NAVY, font=FONT)

    stages = [
        ("01", "INTAKE", "Channels & Sources — how demand enters"),
        ("02", "CLASSIFICATION", "Type, size, urgency — how demand is sorted"),
        ("03", "PRIORITIZATION", "Scoring & criteria — how decisions are made"),
        ("04", "PORTFOLIO", "Governance & decision — how it's owned"),
        ("05", "DELIVERY", "Execution & flow — how it's built"),
    ]
    y = 2.0
    for num, name, desc in stages:
        # Number badge
        add_rect(slide, Inches(0.5), Inches(y), Inches(0.55), Inches(0.55), NAVY)
        add_text(slide, Inches(0.5), Inches(y), Inches(0.55), Inches(0.55),
                 num, size=14, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, Inches(1.2), Inches(y), Inches(4.5), Inches(0.3),
                 name, size=11, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(1.2), Inches(y + 0.3), Inches(4.5), Inches(0.3),
                 desc, size=9, color=CHARCOAL, font=FONT)
        y += 0.75

    # Right column: legend — maturity coding
    add_text(slide, Inches(6.5), Inches(1.55), Inches(6.5), Inches(0.35),
             "READING THE CELLS", size=10, bold=True, color=NAVY, font=FONT)

    legend_items = [
        (GREEN_OK, "Process works — formal channel + clear ownership"),
        (YELLOW_GAP, "Partial gap — informal or inconsistent"),
        (RED_CRITICAL, "Critical gap — no formal process or tool"),
        (MID_GRAY, "Not observed — data not captured in interviews"),
    ]
    y = 2.0
    for color, desc in legend_items:
        add_rect(slide, Inches(6.5), Inches(y), Inches(0.35), Inches(0.35), color)
        add_text(slide, Inches(7.0), Inches(y), Inches(6), Inches(0.35),
                 desc, size=10, color=CHARCOAL, font=FONT, anchor=MSO_ANCHOR.MIDDLE)
        y += 0.55

    # Additional legend: areas
    add_text(slide, Inches(6.5), Inches(4.4), Inches(6.5), Inches(0.3),
             "ORGANIZATIONAL UNITS COVERED", size=10, bold=True, color=NAVY, font=FONT)
    # Two-branch compact legend
    add_text(slide, Inches(6.5), Inches(4.75), Inches(6.5), Inches(0.25),
             "Applications branch — Director Director IT Business Applications",
             size=9, bold=True, color=BLUE, font=FONT)
    apps = "GIS (GIS Champion)  •  Corporate (Corporate Champion)  •  G&E (G&E Champion/E&G Ops Manager)  •  Global SAP (Global SAP Champion/Global SAP Manager)"
    add_text(slide, Inches(6.5), Inches(5.00), Inches(6.5), Inches(0.4),
             apps, size=9, color=CHARCOAL, font=FONT)
    add_text(slide, Inches(6.5), Inches(5.50), Inches(6.5), Inches(0.25),
             "Customer Services branch — Director Director Customer Services Apps",
             size=9, bold=True, color=BLUE, font=FONT)
    cs = "Smart Metering (Smart Metering Champion/Device & Metering Manager)  •  Branch overview from CS Director + Business-side lead  —  Customer Services"
    add_text(slide, Inches(6.5), Inches(5.75), Inches(6.5), Inches(0.4),
             cs, size=9, color=CHARCOAL, font=FONT)

    add_takeaway(slide,
        "The map is organized as 5 area swimlanes × 5 process stages — every cell is a claim grounded in interview evidence")
    add_footer(slide,
        "Source: Champions Workshop (6 interviewees, Apr 2026); IT CS Interview (Business-side lead  —  Customer Services & CS Director); Applications Interview (Apps Director)",
        "How to Read  |  06 / 19")


def _draw_stage_header(slide, stages_y, stage_w, stage_h):
    """Headers for the 5 process stages — used on master map."""
    stages = [
        ("01 INTAKE", "Channels & Sources"),
        ("02 CLASSIFICATION", "Type, Size, Urgency"),
        ("03 PRIORITIZATION", "Scoring & Criteria"),
        ("04 PORTFOLIO", "Governance & Decision"),
        ("05 DELIVERY", "Execution & Flow"),
    ]
    for i, (num_name, sub) in enumerate(stages):
        x = Inches(2.2 + i * stage_w)
        add_rect(slide, x, Inches(stages_y), Inches(stage_w - 0.05), Inches(stage_h),
                 NAVY)
        add_text(slide, x, Inches(stages_y + 0.03), Inches(stage_w - 0.05), Inches(0.27),
                 num_name, size=9, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER)
        add_text(slide, x, Inches(stages_y + 0.28), Inches(stage_w - 0.05), Inches(0.24),
                 sub, size=7, color=WHITE, font=FONT, align=PP_ALIGN.CENTER)
        # Flow arrow between stages (not after the last)
        if i < len(stages) - 1:
            arrow_x = 2.2 + (i + 1) * stage_w - 0.06
            arrow = slide.shapes.add_shape(
                MSO_SHAPE.RIGHT_ARROW,
                Inches(arrow_x), Inches(stages_y + stage_h / 2 - 0.06),
                Inches(0.15), Inches(0.12)
            )
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = ACCENT
            arrow.line.fill.background()
            arrow.shadow.inherit = False


def slide_3_master_map(prs):
    """HERO — 5 areas × 5 stages master map."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)

    add_action_title(slide,
        "Across both director branches, Digital Products runs 6+ disconnected intake channels — no two teams share the same demand flow")

    # Grid setup — 6 rows, balanced breathing room above the matrix
    matrix_top = 1.55
    area_col_w = 1.7
    stage_w = 2.13
    header_h = 0.52
    row_h = 0.60

    # Applications branch banner — colored bar + label
    add_rect(slide, Inches(0.5), Inches(matrix_top - 0.32), Inches(1.65), Inches(0.28), BLUE)
    add_text(slide, Inches(0.5), Inches(matrix_top - 0.32), Inches(1.65), Inches(0.28),
             "BRANCH 1 — APPLICATIONS", size=8, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Inches(2.2), Inches(matrix_top - 0.30), Inches(6), Inches(0.24),
             "Director IT Business Applications, Director IT Business Applications  •  4 teams  •  ~26 FTE",
             size=8, italic=True, color=BLUE, font=FONT, anchor=MSO_ANCHOR.MIDDLE)
    # Stage headers
    _draw_stage_header(slide, matrix_top, stage_w, header_h)

    # Rows grouped by director branch (APPLICATIONS first, then CUSTOMER SERVICES / SAP)
    # Each entry: (area, intake, classification, prioritization, portfolio, delivery, colors, branch)
    areas_data = [
        # ─── APPLICATIONS BRANCH (Director IT Business Applications) ─────────────────────
        ("GIS",
         "Jira + Email + Mtgs + IT-Now",
         "Small vs. PDD (yearly)",
         "Business drives; P&T face",
         "PDD approval",
         "Internal team + SWF",
         [YELLOW_GAP, YELLOW_GAP, YELLOW_GAP, GREEN_OK, GREEN_OK],
         "APPS"),
        ("Corporate",
         "Jira + Confluence + Word",
         "PDD driver; overlap",
         "PDD driver",
         "PDD approval",
         "Internal + Vendor",
         [YELLOW_GAP, RED_CRITICAL, YELLOW_GAP, GREEN_OK, YELLOW_GAP],
         "APPS"),
        ("G&E",
         "Email only (biz/vendor)",
         "Depends on demand nature",
         "Operational = urgent",
         "Vendor sets timeline",
         "Vendor",
         [RED_CRITICAL, RED_CRITICAL, RED_CRITICAL, RED_CRITICAL, RED_CRITICAL],
         "APPS"),
        ("Global SAP\n(Iberdrola)",
         "IT Now + RITM + Jira + SP",
         "By type (4 channels)",
         "Incidents first (P&T)",
         "PDD / SharePoint",
         "Accenture + SWF",
         [GREEN_OK, GREEN_OK, RED_CRITICAL, YELLOW_GAP, YELLOW_GAP],
         "APPS"),
        # ─── CUSTOMER SERVICES / SAP BRANCH (Director Customer Services Apps) ─────
        ("Smart Metering\n(Device & Metering Manager)",
         "Email (IT Now)",
         "IT Now template",
         "Criticality + deadline",
         "PPM Now + weekly mtg",
         "Vendor (100%)",
         [YELLOW_GAP, GREEN_OK, YELLOW_GAP, GREEN_OK, RED_CRITICAL],
         "CS"),
        ("Customer Services\n(CS Director — 5 teams)",
         "Director-level routing only",
         "Travis PMO / BRS / IT-funded split",
         "\"Intrinsic common sense\"; forums",
         "No branch tool (Metering = PPM Now only)",
         "SWF + Accenture + Vendors",
         [RED_CRITICAL, YELLOW_GAP, RED_CRITICAL, RED_CRITICAL, YELLOW_GAP],
         "CS"),
    ]

    y = matrix_top + header_h + 0.05
    prev_branch = None
    for area, intake, classif, prior, portf, deliv, colors, branch in areas_data:
        # Branch divider + label when we transition APPS → CS
        if prev_branch == "APPS" and branch == "CS":
            # Accent divider line
            add_rect(slide, Inches(0.5), Inches(y), Inches(12.33), Inches(0.05), ACCENT)
            y += 0.08
            # Branch banner
            add_rect(slide, Inches(0.5), Inches(y), Inches(1.65), Inches(0.28), NAVY)
            add_text(slide, Inches(0.5), Inches(y), Inches(1.65), Inches(0.28),
                     "BRANCH 2 — CUSTOMER SERVICES", size=7, bold=True, color=WHITE, font=FONT,
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            add_text(slide, Inches(2.2), Inches(y + 0.02), Inches(6), Inches(0.24),
                     "Director Customer Services Apps, Director Customer Services Apps  •  5 teams  •  ~24 FTE",
                     size=8, italic=True, color=NAVY, font=FONT, anchor=MSO_ANCHOR.MIDDLE)
            y += 0.32

        # Row: dark navy for CS branch label, human blue for APPS
        label_color = NAVY if branch == "CS" else BLUE
        add_rect(slide, Inches(0.5), Inches(y), Inches(1.65), Inches(row_h - 0.08), label_color)
        add_text(slide, Inches(0.5), Inches(y), Inches(1.65), Inches(row_h - 0.08),
                 area, size=10, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # 5 cells
        cells = [intake, classif, prior, portf, deliv]
        for i, (text, color) in enumerate(zip(cells, colors)):
            x = Inches(2.2 + i * stage_w)
            add_rect(slide, x, Inches(y), Inches(stage_w - 0.05), Inches(row_h - 0.08),
                     LIGHT_GRAY)
            # Maturity dot
            add_rect(slide, Inches(2.2 + i * stage_w + 0.08), Inches(y + 0.08),
                     Inches(0.15), Inches(0.15), color)
            add_text(slide, Inches(2.2 + i * stage_w + 0.30), Inches(y + 0.05),
                     Inches(stage_w - 0.35), Inches(row_h - 0.15),
                     text, size=10, color=CHARCOAL, font=FONT,
                     anchor=MSO_ANCHOR.MIDDLE)
        y += row_h
        prev_branch = branch

    # Maturity legend at bottom of matrix
    leg_y = 6.20
    add_text(slide, Inches(0.5), Inches(leg_y), Inches(2.5), Inches(0.25),
             "MATURITY  /  READING THE CELLS:", size=9, bold=True, color=NAVY, font=FONT)
    # Green
    add_rect(slide, Inches(3.1), Inches(leg_y + 0.03), Inches(0.18), Inches(0.18), GREEN_OK)
    add_text(slide, Inches(3.33), Inches(leg_y), Inches(2.8), Inches(0.25),
             "Process works — formal channel", size=9, color=CHARCOAL, font=FONT)
    # Yellow
    add_rect(slide, Inches(6.3), Inches(leg_y + 0.03), Inches(0.18), Inches(0.18), YELLOW_GAP)
    add_text(slide, Inches(6.53), Inches(leg_y), Inches(2.8), Inches(0.25),
             "Partial gap — informal / inconsistent", size=9, color=CHARCOAL, font=FONT)
    # Red
    add_rect(slide, Inches(9.65), Inches(leg_y + 0.03), Inches(0.18), Inches(0.18), RED_CRITICAL)
    add_text(slide, Inches(9.88), Inches(leg_y), Inches(2.9), Inches(0.25),
             "Critical gap — no formal process", size=9, color=CHARCOAL, font=FONT)

    add_takeaway(slide,
        "Zero standardization across units: 6 different intake models, 6 prioritization logics, only 1 area uses a portfolio tool — fragmentation is structural, not local")
    add_footer(slide, page_label="Master Demand Map  |  04 / 13")


def slide_4_gis(prs):
    """Drill-down: GIS."""
    _drill_down(
        prs,
        page_num="14 / 28",
        area="GIS",
        champion="GIS & OSG — team led by GIS Sr. Manager / Champion (Sr. Manager)",
        action_title="GIS balances break/fix urgency against yearly PDD planning — with no formal channel for small demand",
        requesters="Electric & Gas Field Operations · T&D Engineering · Asset Management · Planning & Design · Compliance (DIMP/TIMP) · OSG (Operational Support Group)",
        grid={
            "CHANNELS": [
                "Jira (misc. work) + Email / Meetings (small demand) + IT-Now (daily break/fix)",
                "Small: email / Jira / IT-Now routed by type; Large: yearly PDD process",
                "Business drives prioritization; P&T is the face for small demand",
                "PDD-approved portfolio; no dedicated portfolio tool",
                "Internal team for break/fix; SWF brought in for resource impact",
            ],
            "ROLES": [
                "Business users open tickets; GIS Champion's team receives",
                "GIS Champion + team classify; PDD committee for large",
                "\"Impact to most users\" rule (GIS Champion); PDD approvers for large",
                "Informal PM — GIS Champion manages via OneNote / Teams / spreadsheets",
                "GIS team + Software Factory (shared server infra: SAP HANA, ARCFM)",
            ],
            "TOOLS": [
                "Jira · Email · Meetings · IT-Now (RITM)",
                "No single tool — Jira tags, PDD workbook",
                "Spreadsheet-based; no scoring tool",
                "OneNote + Teams + Spreadsheets (informal) + PDD (large)",
                "Jira (internal) + SWF process (vendor capacity)",
            ],
            "GOVERNANCE": [
                "No SLA on small demand; no volume tracking",
                "No classification rule below PDD threshold",
                "\"Highest impact / most users\" — heuristic, not scored",
                "PDD governance works for large; no below-threshold governance",
                "SWF engagement tracked ad-hoc; no KPIs per release",
            ],
        },
        evidence=(
            "\"I prioritize things that have the highest impact to the most users. "
            "In most cases these are break/fix requests\" — GIS Sr. Manager / Champion"
        ),
        takeaway=(
            "GIS has a working large-project process (PDD) but relies on email + meetings "
            "for small demand — capacity risk as volume grows"
        ),
        maturity=["Partial", "Partial", "Partial", "OK", "OK"],
        analysis=[
            "GIS treats \"small demand\" and \"break/fix\" as interchangeable — volume grows, prioritization logic doesn't scale",
            "PDD works for large projects but leaves 60-70% of actual work (break/fix + small) ungoverned",
            "SWF is the elasticity valve — predictable capacity relief, but vendor-dependent and untracked",
        ],
    )


def slide_5_corporate(prs):
    _drill_down(
        prs,
        page_num="15 / 28",
        area="Corporate",
        champion="Corporate Applications — led by Sr. Manager / Corporate Apps  ·  Champion: Corporate Tech Lead (5-6 PMs + 2 Tech Leads)",
        action_title="Corporate runs demand through Jira, Confluence and Word — standardization gap despite the highest regulatory pressure",
        requesters="Finance · HR / People · Procurement · Legal & Compliance · Training & Corporate Functions · Regulatory affairs · Enterprise Shared Services",
        grid={
            "CHANNELS": [
                "PM via Jira + Confluence; Word docs for business requirements",
                "PDD drives for large; no explicit model for small/medium",
                "PDD is the driver; regulatory is treated as standing priority",
                "PDD-approved portfolio; overlap visible in Confluence",
                "Internal + Vendor split; regulatory-driven release cadence",
            ],
            "ROLES": [
                "Business analysts send requirements; PMs own Jira intake",
                "5-6 PMs + 2 Tech Leads classify (central PM pool sits here)",
                "PMs + business leadership — regulatory always lifted",
                "PMs track Jira; overlap visible but not resolved upstream",
                "Internal team + SWF + vendors for specific modules",
            ],
            "TOOLS": [
                "Jira · Confluence · Word docs",
                "Jira labels + PDD workbook",
                "No scoring tool — regulatory tag = priority",
                "Confluence + PDD; no consolidated portfolio tool",
                "Jira + Confluence; SWF for vendor capacity",
            ],
            "GOVERNANCE": [
                "No consolidated intake funnel; Word breaks digital trail",
                "No deduplication between overlapping PDDs",
                "Regulatory urgency is exception-driven, not scored",
                "PDD approvals govern; below-threshold demand untracked",
                "One-year plan + recurring training; no release KPIs",
            ],
        },
        evidence=(
            "\"Regulatory and Employee — not only one year plan and training, "
            "coming with new requirements that are always a priority\" — Corporate Champion (Tech Lead)"
        ),
        takeaway=(
            "Corporate absorbs regulatory volatility through PDD — but Word-doc intake and "
            "overlap between projects create duplication and rework"
        ),
        maturity=["Partial", "Gap", "Partial", "OK", "Partial"],
        analysis=[
            "Word-document requirements betray a pre-agile posture — downstream Jira adoption is superficial, not structural",
            "\"Overlapping projects\" signals classification breakdown: multiple PDDs chase similar outcomes because intake doesn't deduplicate",
            "Regulatory pressure masks the lack of formal scoring — everything becomes urgent by exception",
        ],
    )


def slide_6_ge(prs):
    _drill_down(
        prs,
        page_num="16 / 28",
        area="G&E / Electric & Gas Operations",
        champion="G&E — sub-scope within Electric & Gas Operations Products (led by Sr. Manager / E&G Ops)  ·  Champion: G&E Tech Lead (70-90 apps)",
        action_title="G&E is the only area with NO formal digital intake — every request arrives via email from business or vendor",
        requesters="Electric & Gas Operations (T&D) · Field Services · Outage Management · Third-party vendors (direct) · Cyber & Legal (compliance gates) · Regulatory affairs",
        grid={
            "CHANNELS": [
                "EMAILS only — from business AND from vendor; no ticketing",
                "Ad-hoc by demand nature; depends on upgrade type",
                "G&E Champion's internal call — operational = urgent",
                "Vendor sets its own timeline; upgrades need budget/time eval",
                "Vendor-led delivery; competing priorities without triage",
            ],
            "ROLES": [
                "Business + vendor send email → G&E Champion receives",
                "G&E Champion alone classifies — no second reviewer",
                "G&E Champion prioritizes; cyber / legal / third-party weighted",
                "Vendor owns timeline; G&E Champion tracks manually",
                "Vendor + G&E Champion; no internal release role",
            ],
            "TOOLS": [
                "Email (Outlook) — no Jira, no IT-Now, no RITM",
                "Email threads as the classification evidence",
                "No scoring tool — judgment call",
                "No portfolio tool in use",
                "No delivery tooling — vendor runs its own stack",
            ],
            "GOVERNANCE": [
                "Zero digital trail; volumes unknown; single-person dependency",
                "No written rule; no audit trail",
                "Key-person risk — G&E Champion = only router",
                "Third-Party Risk gates upgrades; interface is slow",
                "Budget / time tension on every upgrade; no KPIs",
            ],
        },
        evidence=(
            "\"I receive requests/input from the business or from the vendor\" — "
            "G&E Champion (Tech Lead, 70-90 apps) (via email)"
        ),
        takeaway=(
            "G&E is the canary of intake fragmentation — zero digital trail, every decision "
            "is individual and unreproducible; must be the first area addressed in the TOM"
        ),
        maturity=["Gap", "Gap", "Gap", "Gap", "Gap"],
        analysis=[
            "Vendor-led model creates structural dependency: Avangrid cannot estimate or schedule its own work — budget/timeline are exogenous",
            "Third-party risk assessments add hidden lead time — upgrades routinely stall at this interface, invisible in any tracking system",
            "Single-person routing (G&E Champion) is a key-person risk — zero redundancy, zero documentation, zero auditability",
        ],
    )


def slide_7_global_sap(prs):
    _drill_down(
        prs,
        page_num="17 / 28",
        area="Global SAP — Iberdrola (Spain)",
        champion="Global SAP — Iberdrola environment (Spain) — led by Manager IT Apps Global SAP, covering FI, PO, Budget control  ·  Champion: Global SAP Tech Lead (35 yrs in business, now IT)",
        action_title="Global SAP — Iberdrola has the most mature channel taxonomy — but is trapped by incident overload",
        requesters="Iberdrola Corporate Finance · Controlling & Budget · Procurement / PO · Accounts Payable · Treasury · Accenture (support partner) · Avangrid Finance (US consolidation)",
        grid={
            "CHANNELS": [
                "IT Now (Incidents) · RITM (Minor) · Jira (Small Demands) · SharePoint + Teams (Projects)",
                "Incidents→IT NOW · Small→JIRA · Minor→RITM · Major→SharePoint",
                "#1 Incidents always first; #2 Small + Major; P&T prioritizes projects",
                "PDD for major projects; SharePoint is project home",
                "Accenture + SWF (Software Factory) run delivery",
            ],
            "ROLES": [
                "Business / Accenture open tickets; Global SAP Manager's 5-FTE team receives",
                "Classification is tool-driven — channel = class",
                "Global SAP Champion + leadership call incidents first; P&T helps with projects",
                "Global SAP Champion owns weekly incident review (Thursday, with Accenture)",
                "Accenture = incidents; SWF = projects; P&T = intermediary",
            ],
            "TOOLS": [
                "IT Now · IT Now RITM · Jira · SharePoint · Teams",
                "IT Now + Jira + SharePoint (by channel)",
                "No scoring tool — incidents-first rule",
                "SharePoint (projects) + PDD (approval) + Thursday incident meeting",
                "Accenture's delivery stack + SWF pipeline",
            ],
            "GOVERNANCE": [
                "Strongest channel taxonomy in the org — but tuned for incidents",
                "Channel taxonomy = de-facto classification",
                "Incident volume caps strategic capacity (\"entire team on incidents\")",
                "Weekly Accenture incident review governs incidents; projects via PDD",
                "Knowledge transfer loss at every turnover — 1-year hiring lag",
            ],
        },
        evidence=(
            "\"Balance all the demands: Incidents, Small Demands, Major projects (PDD). "
            "Limited resources — 1 year to hire someone\" — Global SAP Champion (Tech Lead)"
        ),
        takeaway=(
            "Global SAP's channel taxonomy is the closest to a target state — but the incident "
            "flood caps strategic capacity; reducing incident load is the leverage point"
        ),
        maturity=["OK", "OK", "Gap", "Partial", "Partial"],
        analysis=[
            "Channel taxonomy is the strongest in the organization — but it was designed around incidents, not strategic demand",
            "Incidents-first rule means strategic work enters a de-facto waiting queue — hiring lag (1 year) keeps the queue deep",
            "Accenture + SWF dependency is the hidden scaling mechanism — cost and knowledge concentrate outside Avangrid",
        ],
    )


def slide_8_smart_metering(prs):
    _drill_down(
        prs,
        page_num="18 / 28",
        area="Smart Metering — Avangrid (US)",
        champion="Smart Metering — sub-scope within Device & Metering (led by Sr. Manager / Device & Metering, Customer Services branch)  ·  Champion: Smart Metering PM (works with Corporate Apps Manager + G&E team)",
        action_title="Smart Metering is the only area fully vendor-delivered with PPM Now + Confluence — but relies on \"intrinsic common sense\" for prioritization",
        requesters="Customer Operations · Field Services (meter installers) · Billing & Collections · Grid Ops · BRS UIL / CMP / NY (regional business/regulatory groups) · Vendors (primary delivery partner)",
        grid={
            "CHANNELS": [
                "Small demands via email (IT Now); normal flow — not seasonal",
                "Incidents → IT Now template; requirements usually clear",
                "Criticality + deadline + application priority",
                "PPM Now (large) + Jira + Confluence (PM-to-PM)",
                "100% vendor-delivered portfolio — vendors = consultants + code",
            ],
            "ROLES": [
                "Business opens IT Now / email; Smart Metering Champion (PM) receives",
                "Smart Metering Champion + team classify by template; vendor tech lead joins",
                "Smart Metering Champion prioritizes on deadline + criticality — \"intrinsic common sense\"",
                "Smart Metering Champion + manager (weekly); vendor PMs help PMs prioritize",
                "Vendors do requirements AND code; Avangrid = product owner",
            ],
            "TOOLS": [
                "Email · IT Now (template) · occasional RITM",
                "IT Now template = primary classification",
                "No scoring tool — judgment + weekly cadence",
                "PPM Now + Confluence + Jira — most mature stack in scope",
                "Vendor stack for build; Confluence for portfolio documentation",
            ],
            "GOVERNANCE": [
                "Volumes OK (not seasonal); \"usually no urgent requests\"",
                "Template-based → consistent, but no cross-team view",
                "\"Intrinsic common sense\" = tribal knowledge risk",
                "Weekly meeting governs; PPM Now = only portfolio instance in CS branch",
                "If vendor exits, execution knowledge is lost — 100% dependency",
            ],
        },
        evidence=(
            "\"All portfolio in vendor. Vendors are the consultants for the work — getting "
            "requirements on tech side, steps to follow and do the code also — all custom "
            "applications\" — Smart Metering Champion (PM)"
        ),
        takeaway=(
            "Smart Metering is the most tool-mature area — but full vendor dependency and "
            "intuitive prioritization create knowledge and governance risk"
        ),
        maturity=["Partial", "OK", "Partial", "OK", "Gap"],
        analysis=[
            "Only area using PPM Now + Confluence — the stack is right, but it exists as one-team exception rather than branch standard",
            "100% vendor delivery (requirements + code) means Avangrid owns no execution knowledge — if the vendor exits, everything stops",
            "\"Intrinsic common sense\" for prioritization signals tribal knowledge — works at current scale, fails on leadership transition",
        ],
    )


def _drill_down(prs, *, page_num, area, champion, action_title, stages=None, evidence, takeaway, footer=None,
                maturity=None, analysis=None, grid=None, requesters=None):
    """Template for per-area drill-down slides — 4 categories × 5 stages grid.

    grid: dict with 4 keys (ordered): "CHANNELS", "ROLES", "TOOLS", "GOVERNANCE"
          each mapped to a list of 5 strings (one per lifecycle stage).
          When grid is None, falls back to legacy `stages` dict (kept for safety).
    maturity: list of 5 maturity labels (OK / Partial / Gap / N/A) — one per stage.
    analysis: list of 2-3 analytical bullets.
    requesters: string — the types of business / stakeholders that send demand to this team.
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)

    add_action_title(slide, action_title)

    # Champion strip
    add_text(slide, Inches(0.5), Inches(1.28), Inches(12.33), Inches(0.26),
             champion, size=10, color=BLUE, font=FONT, italic=True)

    # Business requesters banner — describes which business units send demand to this team
    if requesters:
        req_y = 1.56
        req_h = 0.30
        add_rect(slide, Inches(0.5), Inches(req_y), Inches(12.33), Inches(req_h), NAVY)
        add_rect(slide, Inches(0.5), Inches(req_y), Inches(2.0), Inches(req_h), ACCENT)
        add_text(slide, Inches(0.5), Inches(req_y), Inches(2.0), Inches(req_h),
                 "BUSINESS REQUESTERS", size=9, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, Inches(2.6), Inches(req_y), Inches(10.15), Inches(req_h),
                 requesters, size=9, color=WHITE, font=FONT,
                 anchor=MSO_ANCHOR.MIDDLE)
        header_top = 1.98
    else:
        header_top = 1.72

    maturity_colors = {
        "OK": GREEN_OK,
        "Partial": YELLOW_GAP,
        "Gap": RED_CRITICAL,
        "N/A": MID_GRAY,
    }

    # Layout geometry
    label_w = 0.65
    col_x0 = 0.5 + label_w + 0.05
    col_w = 2.37
    col_gap = 0.06

    header_h = 0.38
    pill_h = 0.20
    pill_gap = 0.06

    # --- Stage headers (01 INTAKE, 02 CLASSIFICATION, ...) with maturity pills ---
    stage_names = ["01 INTAKE", "02 CLASSIFICATION", "03 PRIORITIZATION", "04 PORTFOLIO", "05 DELIVERY"]
    for i, sname in enumerate(stage_names):
        x = col_x0 + i * (col_w + col_gap)
        add_rect(slide, Inches(x), Inches(header_top), Inches(col_w), Inches(header_h), NAVY)
        add_text(slide, Inches(x), Inches(header_top), Inches(col_w), Inches(header_h),
                 sname, size=10, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Maturity pill just below header
        if maturity and i < len(maturity):
            pill_label = maturity[i]
            pill_color = maturity_colors.get(pill_label, MID_GRAY)
            pill_y = header_top + header_h + pill_gap
            add_rect(slide, Inches(x + (col_w - 1.2) / 2), Inches(pill_y),
                     Inches(1.2), Inches(pill_h), pill_color, corner=True)
            add_text(slide, Inches(x + (col_w - 1.2) / 2), Inches(pill_y),
                     Inches(1.2), Inches(pill_h),
                     pill_label.upper(), size=9, bold=True, color=WHITE, font=FONT,
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # --- Grid: 4 category rows × 5 stage columns ---
    grid_top = header_top + header_h + pill_h + pill_gap + 0.06
    row_h = 0.54
    categories = [
        ("CHANNELS  ·  HOW IT ARRIVES", GREEN_OK),
        ("ROLES", BLUE),
        ("TOOLS", NAVY),
        ("GOVERNANCE  &  METRICS", ACCENT),
    ]
    grid_keys = ["CHANNELS", "ROLES", "TOOLS", "GOVERNANCE"]

    # Fallback — if no grid supplied, synthesize from stages (legacy safety)
    if grid is None and stages is not None:
        # Put everything under CHANNELS; leave others blank
        stage_items = list(stages.values())
        grid = {
            "CHANNELS": [" · ".join(s) for s in stage_items],
            "ROLES": ["—"] * 5,
            "TOOLS": ["—"] * 5,
            "GOVERNANCE": ["—"] * 5,
        }

    for r, (label, accent) in enumerate(categories):
        y = grid_top + r * (row_h + 0.04)
        # Left label block
        add_rect(slide, Inches(0.5), Inches(y), Inches(label_w), Inches(row_h), accent)
        add_text(slide, Inches(0.5), Inches(y), Inches(label_w), Inches(row_h),
                 label, size=7, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # 5 cells
        items = grid.get(grid_keys[r], ["—"] * 5)
        for i in range(5):
            x = col_x0 + i * (col_w + col_gap)
            add_rect(slide, Inches(x), Inches(y), Inches(col_w), Inches(row_h), LIGHT_GRAY)
            add_text(slide, Inches(x + 0.10), Inches(y + 0.05),
                     Inches(col_w - 0.2), Inches(row_h - 0.1),
                     items[i] if i < len(items) else "—",
                     size=9, color=CHARCOAL, font=FONT, anchor=MSO_ANCHOR.TOP)

    # --- KEY PATTERNS + SUPPORTING QUOTE block ---
    grid_bottom = grid_top + 4 * (row_h + 0.04)
    an_top = grid_bottom + 0.08
    an_h = 6.50 - an_top  # stop BEFORE the takeaway banner (no overflow)
    add_rect(slide, Inches(0.5), Inches(an_top), Inches(12.33), Inches(an_h), LIGHT_GRAY)
    add_rect(slide, Inches(0.5), Inches(an_top), Inches(0.08), Inches(an_h), NAVY)
    add_text(slide, Inches(0.70), Inches(an_top + 0.05), Inches(6), Inches(0.22),
             "KEY PATTERNS  —  WHAT THE FINDINGS REVEAL",
             size=9, bold=True, color=NAVY, font=FONT)
    add_text(slide, Inches(7.0), Inches(an_top + 0.05), Inches(5.65), Inches(0.22),
             "SUPPORTING QUOTE", size=9, bold=True, color=ACCENT, font=FONT)
    add_text(slide, Inches(7.0), Inches(an_top + 0.32), Inches(5.65), Inches(an_h - 0.40),
             evidence, size=9, color=CHARCOAL, font=FONT, italic=True)
    if analysis:
        analysis_text = "\n".join(["▸  " + a for a in analysis])
        add_text(slide, Inches(0.70), Inches(an_top + 0.32), Inches(6.0), Inches(an_h - 0.40),
                 analysis_text, size=9, color=CHARCOAL, font=FONT)

    add_takeaway(slide, takeaway, color=ACCENT)
    add_footer(slide, page_label=f"Drill-down: {area}  |  {page_num}")


def slide_10_customer_services(prs):
    """Drill-down: Customer Services Apps branch (Director Customer Services Apps)."""
    _drill_down(
        prs,
        page_num="19 / 28",
        area="Customer Services Apps (CS Director's Branch)",
        champion="Customer Services Apps — Director Customer Services Apps (Director) — 5 sub-teams: FICA/Billing, CRM & Contact Center (vacant), Device & Metering, SAP & CS Architecture (vacant), SAP Basis",
        action_title="Customer Services runs 5 sub-teams with heterogeneous maturity — only Smart Metering exhibits a mature portfolio practice; the rest route demand director-level",
        requesters="Customer Billing & Collections · Contact Center · CRM operations · Credit & Risk · BRS UIL (CT) · BRS CMP (ME) · BRS NY (New York) · Travis PMO · 7 OPCOs (via BRS)",
        grid={
            "CHANNELS": [
                "Email + IT Now (Metering); Jira (small-med); no shared branch channel",
                "Incidents→IT Now template; Large→Travis PMO; Small/Med→Jira",
                "Criticality + deadline (Metering); director-level calls elsewhere",
                "PPM Now (Metering only); no branch-wide tool",
                "SWF + Accenture + Vendors (100% in Metering)",
            ],
            "ROLES": [
                "Business / BRS / OPCOs open varied channels; 5 Sr. Managers (2 vacant) receive",
                "By sub-team (Metering/FICA/CRM/Basis); Travis PMO for large",
                "CS Director + sub-team managers; \"intrinsic common sense\" (Smart Metering Champion)",
                "Travis PMO (large) + sub-team managers (small); no portfolio lead",
                "SWF runs projects; PM + SAP Architect roles in place (large)",
            ],
            "TOOLS": [
                "IT Now · Email · Jira · no cross-channel funnel",
                "IT Now template + Jira labels + PPM (large)",
                "Forum criteria (strategic / ops / regulatory / size) — applied by sub-team",
                "PPM Now (Metering) · SharePoint · Confluence",
                "SWF delivery pipeline · vendor stacks (per sub-team)",
            ],
            "GOVERNANCE": [
                "2 Sr. Manager vacancies (CRM & Contact Center, SAP & CS Architecture)",
                "Triple routing: Travis PMO / BRS / IT-funded — blurs accountability",
                "No formal director-level model; regulatory overrides size",
                "7 OPCOs compete for the same IT capacity — multi-OPCO contention",
                "Heavy vendor concentration; single-team mature practice (Metering)",
            ],
        },
        evidence=(
            "\"Who is in charge of what — 3 people can run the payment jobs for example. "
            "Resource allocation — multiple OPCOs competing by capacity\" — CS Director + Business-side lead  —  Customer Services"
        ),
        takeaway=(
            "Smart Metering's model (PPM Now + weekly PM meeting + 100% vendor delivery) is the "
            "only mature practice in CS Director's branch — it is also the blueprint and the warning sign"
        ),
        maturity=["Gap", "Partial", "Gap", "Gap", "Partial"],
        analysis=[
            "Three parallel ownership tracks without a routing rule — Travis PMO for large projects, BRS for business-funded work, IT-funded for CS Director's teams — the same request can land in any of them, blurring accountability and prioritization",
            "Smart Metering is the mature outlier; other four teams operate without equivalent tooling — branch has no portfolio consolidation",
            "PM + SAP Architect roles are in place for large projects (CS Director confirmed) — the gap is not in roles, it's in branch-wide portfolio visibility",
        ],
    )


def _bottleneck_card(slide, x_in, y_in, card_w, card_h, b):
    """Single cleaner bottleneck card — minimal palette, no vertical overlap."""
    # Card background
    add_rect(slide, Inches(x_in), Inches(y_in), Inches(card_w), Inches(card_h), LIGHT_GRAY)
    # Top accent stripe
    add_rect(slide, Inches(x_in), Inches(y_in), Inches(card_w), Inches(0.06), ACCENT)
    # Number (big)
    add_text(slide, Inches(x_in + 0.28), Inches(y_in + 0.12), Inches(1.0), Inches(0.48),
             b["num"], size=24, bold=True, color=ACCENT, font=FONT)
    # Stage tag (top-right of number)
    add_text(slide, Inches(x_in + 1.25), Inches(y_in + 0.18), Inches(card_w - 1.45), Inches(0.22),
             b["stage"], size=8, bold=True, color=MID_GRAY, font=FONT)
    # Title
    add_text(slide, Inches(x_in + 1.25), Inches(y_in + 0.41), Inches(card_w - 1.45), Inches(0.30),
             b["title"], size=13, bold=True, color=NAVY, font=FONT)
    # Description
    add_text(slide, Inches(x_in + 0.28), Inches(y_in + 0.78), Inches(card_w - 0.55), Inches(0.60),
             b["desc"], size=10, color=CHARCOAL, font=FONT)
    # IMPACT label + text — bounded, won't overlap with evidence below
    add_text(slide, Inches(x_in + 0.28), Inches(y_in + 1.42), Inches(card_w - 0.55), Inches(0.18),
             "IMPACT", size=8, bold=True, color=MID_GRAY, font=FONT)
    add_text(slide, Inches(x_in + 0.28), Inches(y_in + 1.60), Inches(card_w - 0.55), Inches(0.40),
             b["impact"], size=9, color=CHARCOAL, font=FONT)
    # Evidence pinned near bottom — separated from IMPACT by visible gap
    ev_y = y_in + card_h - 0.30
    add_text(slide, Inches(x_in + 0.28), Inches(ev_y), Inches(card_w - 0.55), Inches(0.25),
             "▸  " + b["evidence"], size=8, color=MID_GRAY, font=FONT, italic=True)


def _bottlenecks_slide(prs, *, branch, director, page_num, section_label, action_title, takeaway, cards):
    """Generic bottlenecks layout — 2×2 cards, single branch."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    add_action_title(slide, action_title)

    # Branch band
    band_color = BLUE if branch == "APPLICATIONS" else NAVY
    add_rect(slide, Inches(0.5), Inches(1.32), Inches(12.33), Inches(0.30), band_color)
    add_text(slide, Inches(0.5), Inches(1.32), Inches(12.33), Inches(0.30),
             f"{branch} BRANCH  —  {director}",
             size=11, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 2×2 grid — sized to sit comfortably between title band and takeaway
    card_w = 6.27
    card_h = 2.35
    gap_x = 0.13
    gap_y = 0.10
    y_row1 = 1.70
    y_row2 = y_row1 + card_h + gap_y

    for i, b in enumerate(cards):
        col = i % 2
        row = i // 2
        x_in = 0.5 + col * (card_w + gap_x)
        y_in = y_row1 if row == 0 else y_row2
        _bottleneck_card(slide, x_in, y_in, card_w, card_h, b)

    add_takeaway(slide, takeaway, color=ACCENT)
    add_footer(slide, page_label=f"{section_label}  |  {page_num}")


def slide_10_apps_bottlenecks(prs):
    cards = [
        {
            "num": "01",
            "stage": "E&G Ops team  •  STAGE: Intake",
            "title": "G&E Email-Only Intake",
            "desc": "Every request to G&E Champion arrives via email from business or vendor. No ticket, no queue, no SLA — each decision is individual and unreproducible.",
            "impact": "Unknown volume. No capacity forecast possible. Single-person dependency on G&E Champion.",
            "evidence": "\"I receive requests/input from the business or from the vendor\" — G&E Champion (Tech Lead, 70-90 apps)",
        },
        {
            "num": "02",
            "stage": "Global SAP Manager's team  •  STAGE: Prioritization",
            "title": "Global SAP Incident Flood",
            "desc": "Incidents outrank all other demand for Iberdrola SAP. Entire team can be consumed by incident work, blocking strategic projects.",
            "impact": "Capacity trapped in BAU. Major projects (PDD) slip. Knowledge transfer lost at each turnover event.",
            "evidence": "\"Sometimes entire team is focusing on incidents\" — Global SAP Champion (Tech Lead)",
        },
        {
            "num": "03",
            "stage": "GIS Sr. Manager / Champion's team  •  STAGES: Intake + Portfolio",
            "title": "GIS Informal PM Stack",
            "desc": "Small demand routes through email, meetings and IT-Now. Project management relies on OneNote + Teams + spreadsheets — no formal portfolio tool below PDD.",
            "impact": "Small + break/fix demand (~60-70% of actual work) is untracked. Capacity consumed by unmeasured flow.",
            "evidence": "\"My team uses Jira for misc. work... OneNote, Teams and Spreadsheets to manage\" — GIS Sr. Manager / Champion",
        },
        {
            "num": "04",
            "stage": "Corporate Apps team  •  STAGE: Classification",
            "title": "Corporate Project Overlap",
            "desc": "Word-document requirements + Jira + Confluence fragment the intake. PDD drives classification but doesn't deduplicate — overlapping projects chase similar outcomes.",
            "impact": "Rework, duplicated effort, and unreconciled regulatory pressure treated as chronic urgency.",
            "evidence": "\"PDD is the driver. Lots of overlapping projects\" — Corporate Champion (Tech Lead)",
        },
    ]
    _bottlenecks_slide(
        prs,
        branch="APPLICATIONS",
        director="Director IT Business Applications",
        page_num="23 / 28",
        section_label="Applications Bottlenecks",
        action_title="Four Applications bottlenecks concentrate demand risk — covering intake, prioritization, portfolio and classification",
        takeaway="These four bottlenecks span all four Applications teams — each is a symptom of the same underlying fragmentation the TOM must address",
        cards=cards,
    )


def slide_11_cs_bottlenecks(prs):
    cards = [
        {
            "num": "05",
            "stage": "CS Director's branch  •  STAGES: Classification + Prioritization",
            "title": "Customer Services Leadership Gap",
            "desc": "Two Sr. Manager positions vacant (CRM & Contact Center Ops, SAP & CS Architecture). Prioritization defaults to director level across Billing, CRM, Metering and Basis.",
            "impact": "No shared prioritization logic across 5 sub-teams. Branch-level demand view is a structural blind spot.",
            "evidence": "\"Who is in charge of what — 3 people can run the payment jobs\" — CS Director + Business-side lead  —  Customer Services",
        },
        {
            "num": "06",
            "stage": "CS Director's branch  •  STAGES: Portfolio + Delivery",
            "title": "No Branch-Wide Portfolio Tool",
            "desc": "PPM Now exists only in Smart Metering. Travis PMO handles large projects without a consolidated view of Billing / CRM / Metering / Basis portfolios.",
            "impact": "Sub-team portfolios are disconnected. Capacity + priorities cannot be reconciled across the branch.",
            "evidence": "\"No tool to support portfolio management... No tools or process to manage delivery of vendors\"",
        },
        {
            "num": "07",
            "stage": "Device & Metering team  •  STAGE: Delivery",
            "title": "Smart Metering Vendor Concentration",
            "desc": "100% of the Smart Metering portfolio is vendor-delivered. Vendors do requirements AND code. All custom applications. \"Intrinsic common sense\" is the prioritization rule.",
            "impact": "If the vendor exits, execution knowledge is lost. Avangrid has no native build capability in this sub-team.",
            "evidence": "\"All portfolio in vendor. Vendors are consultants for the work\" — Smart Metering Champion (PM)",
        },
        {
            "num": "08",
            "stage": "CS Director's branch  •  STAGE: Intake + Prioritization",
            "title": "Multi-OPCO Capacity Competition",
            "desc": "7 business OPCOs compete for the same Customer Services IT capacity. Demand splits across Travis PMO / BRS / IT-funded — triple routing for the same request.",
            "impact": "Ad-hoc capacity allocation. Ownership ambiguity creates delivery risk and prolongs leadership vacancies' impact.",
            "evidence": "\"Resource allocation — multiple OPCOs competing by capacity\" — CS Director + Business-side lead  —  Customer Services",
        },
    ]
    _bottlenecks_slide(
        prs,
        branch="CUSTOMER SERVICES",
        director="Director Customer Services Apps",
        page_num="24 / 28",
        section_label="Customer Services Bottlenecks",
        action_title="Four Customer Services bottlenecks — leadership gaps, portfolio blindness, vendor concentration and multi-OPCO contention",
        takeaway="CS Director's branch carries both structural (leadership) and operational (tooling + vendor) bottlenecks — the TOM must address all four simultaneously",
        cards=cards,
    )


def slide_10_bottlenecks(prs):
    """Volume & bottlenecks overlay."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)

    add_action_title(slide,
        "Four bottlenecks concentrate demand risk — two in Applications (G&E, Global SAP) and two in Customer Services (leadership + portfolio tooling)")

    bottlenecks = [
        {
            "num": "01",
            "title": "G&E Email-Only Intake",
            "stage": "BRANCH: Applications  /  STAGE: Intake",
            "desc": "Zero digital trail. Every request arrives via email from business or vendor. No triage, no queue, no SLA.",
            "impact": "Unknown volume. No capacity forecast possible. Single-person dependency.",
            "evidence": "\"I receive requests/input from the business or from the vendor\" — G&E Champion (Tech Lead, 70-90 apps)",
            "color": RED_CRITICAL,
        },
        {
            "num": "02",
            "title": "Global SAP Incident Flood",
            "stage": "BRANCH: Applications  /  STAGE: Prioritization",
            "desc": "Incidents outrank all other demand. Entire team can be consumed by incident work, blocking strategic projects.",
            "impact": "Capacity trapped in BAU. Major projects (PDD) slip. Turnover compounds loss.",
            "evidence": "\"Sometimes entire team is focusing on incidents\" — Global SAP Champion (Tech Lead)",
            "color": RED_CRITICAL,
        },
        {
            "num": "03",
            "title": "Customer Services Leadership Gap",
            "stage": "BRANCH: Customer Services  /  STAGE: Classification + Prioritization",
            "desc": "Two Sr. Manager positions vacant (CRM & Contact Center Ops, SAP & CS Architecture). Prioritization happens director-level only across Billing, CRM, Metering and Basis — without shared criteria.",
            "impact": "No consolidated branch-level view of demand. Travis PMO absorbs large projects without area-level coordination. Leadership vacancies compound hiring pressure already observed in Global SAP.",
            "evidence": "\"Who is in charge of what — 3 people can run the payment jobs\" — CS Director + Business-side lead  —  Customer Services",
            "color": RED_CRITICAL,
        },
        {
            "num": "04",
            "title": "Portfolio & Vendor Tooling Gap",
            "stage": "BOTH BRANCHES  /  STAGES: Portfolio + Delivery",
            "desc": "No tool supports portfolio management (5 of 6 units). No process manages vendor delivery. PPM Now used only by Smart Metering.",
            "impact": "Vendor sets timelines and estimates. Budget/schedule control degraded. ROI untraceable.",
            "evidence": "\"No tool to support portfolio management. No tools or process to manage delivery process of vendors\"",
            "color": RED_CRITICAL,
        },
    ]

    # 2×2 grid with BRANCH BAND above each row (for visual grouping)
    band_h = 0.26
    card_h = 2.28
    card_w = 6.27
    gap_x = 0.13
    gap_y_between_rows = 0.08
    y_band1 = 1.32
    y_row1 = y_band1 + band_h + 0.04
    y_band2 = y_row1 + card_h + gap_y_between_rows
    y_row2 = y_band2 + band_h + 0.04
    card_top = y_row1

    # BRANCH BAND 1 — Applications
    add_rect(slide, Inches(0.5), Inches(y_band1), Inches(12.77), Inches(band_h), BLUE)
    add_text(slide, Inches(0.5), Inches(y_band1), Inches(12.77), Inches(band_h),
             "BRANCH 1  —  APPLICATIONS  (Director IT Business Applications)",
             size=10, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # BRANCH BAND 2 — Customer Services + cross-branch
    add_rect(slide, Inches(0.5), Inches(y_band2), Inches(12.77), Inches(band_h), NAVY)
    add_text(slide, Inches(0.5), Inches(y_band2), Inches(12.77), Inches(band_h),
             "BRANCH 2  —  CUSTOMER SERVICES  (Director Customer Services Apps)   +   CROSS-BRANCH (both directors)",
             size=10, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    for i, b in enumerate(bottlenecks):
        col = i % 2
        row = i // 2
        x_in = 0.5 + col * (card_w + gap_x)
        y_in = y_row1 if row == 0 else y_row2
        # Card frame
        add_rect(slide, Inches(x_in), Inches(y_in), Inches(card_w), Inches(card_h),
                 LIGHT_GRAY)
        # Left edge branch color (BLUE=Apps, NAVY=CS, ACCENT=Cross)
        branch_edge_color = BLUE if row == 0 else (ACCENT if i == 3 else NAVY)
        add_rect(slide, Inches(x_in), Inches(y_in), Inches(0.1), Inches(card_h),
                 branch_edge_color)
        # Colored top band
        add_rect(slide, Inches(x_in), Inches(y_in), Inches(card_w), Inches(0.06),
                 b["color"])
        # Big number + stage label
        add_text(slide, Inches(x_in + 0.28), Inches(y_in + 0.12), Inches(0.9), Inches(0.45),
                 b["num"], size=24, bold=True, color=b["color"], font=FONT)
        add_text(slide, Inches(x_in + 1.25), Inches(y_in + 0.20), Inches(card_w - 1.45), Inches(0.22),
                 b["stage"], size=8, bold=True, color=MID_GRAY, font=FONT)
        # Title
        add_text(slide, Inches(x_in + 1.25), Inches(y_in + 0.43), Inches(card_w - 1.45), Inches(0.30),
                 b["title"], size=13, bold=True, color=NAVY, font=FONT)
        # Description
        add_text(slide, Inches(x_in + 0.28), Inches(y_in + 0.80), Inches(card_w - 0.55), Inches(0.65),
                 b["desc"], size=10, color=CHARCOAL, font=FONT)
        # Impact block
        add_text(slide, Inches(x_in + 0.28), Inches(y_in + 1.42), Inches(card_w - 0.55), Inches(0.20),
                 "IMPACT", size=8, bold=True, color=b["color"], font=FONT)
        add_text(slide, Inches(x_in + 0.28), Inches(y_in + 1.60), Inches(card_w - 0.55), Inches(0.38),
                 b["impact"], size=9, color=CHARCOAL, font=FONT)
        # Evidence — dark stripe at bottom for contained visual boundary
        ev_h = 0.28
        add_rect(slide, Inches(x_in), Inches(y_in + card_h - ev_h), Inches(card_w), Inches(ev_h), NAVY)
        add_text(slide, Inches(x_in + 0.28), Inches(y_in + card_h - ev_h),
                 Inches(card_w - 0.55), Inches(ev_h),
                 b["evidence"], size=9, color=WHITE, font=FONT, italic=True,
                 anchor=MSO_ANCHOR.MIDDLE)

    add_takeaway(slide,
        "These four bottlenecks cut across both branches — two in Applications and two in Customer Services; no single-area fix resolves them",
        color=ACCENT)
    add_footer(slide,
        "Source: KB synthesis across 3 interview files; bottlenecks validated against governance forum data (RFI, Apr 2026)",
        "Volume & Bottlenecks  |  — / —")


def slide_data_gaps(prs):
    """Data Gaps — what we can't measure yet, Phase 1.5 scope."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)

    add_action_title(slide,
        "Six measurement dependencies remain open — they shape the TOM business case and must be addressed alongside the operating model design")

    gaps = [
        ("01", "Volumetrics",
         "Tickets / mo per channel  •  PDDs / year  •  Email volume  •  RITM count  •  SharePoint project count",
         "How big is the fire?"),
        ("02", "Flow Metrics",
         "Lead Time (request → delivery)  •  Cycle Time  •  Throughput  •  Work In Progress per team",
         "Where does work stall?"),
        ("03", "Rejection & Rework",
         "Demands rejected at triage  •  Duplicates across channels  •  Projects reopened or descoped",
         "What's pure waste?"),
        ("04", "Taxonomy & Thresholds",
         "Objective definition of Small / Medium / Large (hours, $, story points)  •  Urgency criteria",
         "How do we decide consistently?"),
        ("05", "Capacity & Cost",
         "Internal FTE × Vendor × Accenture × SWF breakdown  •  Budget by category  •  Overhead from fragmentation",
         "What's the financial impact?"),
        ("06", "Value Stream Connection",
         "Link IT demand to business OKRs  •  Map applications to customer journeys  •  Project-to-Product lens",
         "Are we building the right things?"),
    ]

    # 3 columns × 2 rows grid
    col_w = 4.10
    row_h = 2.20
    x0 = 0.5
    y0 = 1.55

    for idx, (num, title, detail, question) in enumerate(gaps):
        col = idx % 3
        row = idx // 3
        x = x0 + col * (col_w + 0.10)
        y = y0 + row * (row_h + 0.15)

        # Card background
        add_rect(slide, Inches(x), Inches(y), Inches(col_w), Inches(row_h), LIGHT_GRAY)
        # Top accent
        add_rect(slide, Inches(x), Inches(y), Inches(col_w), Inches(0.06), ACCENT)
        # Number + title row
        add_text(slide, Inches(x + 0.25), Inches(y + 0.18), Inches(0.6), Inches(0.5),
                 num, size=22, bold=True, color=ACCENT, font=FONT)
        add_text(slide, Inches(x + 0.95), Inches(y + 0.25), Inches(col_w - 1.1), Inches(0.4),
                 title, size=13, bold=True, color=NAVY, font=FONT)
        # Detail bullets
        add_text(slide, Inches(x + 0.25), Inches(y + 0.85), Inches(col_w - 0.5), Inches(0.95),
                 detail, size=9, color=CHARCOAL, font=FONT)
        # Guiding question
        add_text(slide, Inches(x + 0.25), Inches(y + row_h - 0.45), Inches(col_w - 0.5), Inches(0.30),
                 "KEY QUESTION", size=7, bold=True, color=MID_GRAY, font=FONT)
        add_text(slide, Inches(x + 0.25), Inches(y + row_h - 0.25), Inches(col_w - 0.5), Inches(0.28),
                 question, size=10, italic=True, color=BLUE, font=FONT)

    add_takeaway(slide,
        "Without these measurements, the TOM business case stays qualitative — instrumentation should run in parallel with the operating model design",
        color=NAVY)
    add_footer(slide,
        "Framing: open measurement items — flagged transparently rather than fabricated",
        "Measurement Dependencies  |  12 / 13")


def slide_11_insights(prs):
    """Structural insights + handoff."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)

    add_action_title(slide,
        "Seven structural issues surface — no single-area fix resolves them; the TOM must standardize end-to-end")

    insights = [
        ("Intake Fragmentation",
         "6+ disconnected channels (Jira, IT Now, RITM, SharePoint, Email, Word). No single front door.",
         "TOM: unified intake layer with tool-of-record per demand type."),
        ("No Shared Prioritization Model",
         "Each area uses different criteria — PDD, impact×users, criticality, operational urgency, common sense.",
         "TOM: value-based scoring framework (ties to existing \"Value Framework V1.0\" ambition)."),
        ("Portfolio Tooling Vacuum",
         "5 of 6 units have no portfolio tool. Only Smart Metering uses PPM Now.",
         "TOM: scale PPM Now (or equivalent) across all Digital Products areas + IT CS."),
        ("Vendor Governance Gap",
         "No process to manage vendor delivery. Vendors estimate, set timelines, and own Smart Metering end-to-end.",
         "TOM: vendor performance framework + NTT-led estimation sanity check."),
        ("Governance Forum Coverage",
         "5 recurring forums exist — but Manager Check-in participants are still TBC; no forum covers portfolio-level decisions.",
         "TOM: add a portfolio steering cadence and validate forum scope."),
        ("Taxonomy & Thresholds Undefined",
         "\"Small\" vs \"Medium\" vs \"Large\" is subjective; each team decides. No hours / $ / complexity threshold exists.",
         "TOM: define objective classification thresholds — prerequisite for prioritization scoring."),
        ("Project vs Product Lens Missing",
         "Demand is organized by application silo, not by Value Stream. IT requests don't trace to business OKRs.",
         "TOM: introduce Value Stream mapping — move from Project to Product operating model."),
    ]

    y = 1.55
    row_h = 0.70
    for i, (title, gap, tom) in enumerate(insights):
        num = f"0{i+1}"
        # Number
        add_rect(slide, Inches(0.5), Inches(y), Inches(0.55), Inches(row_h - 0.08), ACCENT)
        add_text(slide, Inches(0.5), Inches(y), Inches(0.55), Inches(row_h - 0.08),
                 num, size=16, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Title
        add_text(slide, Inches(1.2), Inches(y), Inches(6.8), Inches(0.28),
                 title, size=13, bold=True, color=NAVY, font=FONT)
        # Gap description — bumped 8→11
        add_text(slide, Inches(1.2), Inches(y + 0.30), Inches(6.8), Inches(0.35),
                 gap, size=11, color=CHARCOAL, font=FONT)
        # TOM response — bumped 8→11
        add_text(slide, Inches(8.2), Inches(y + 0.05), Inches(4.6), Inches(0.58),
                 "→  " + tom, size=11, color=BLUE, font=FONT, italic=True,
                 anchor=MSO_ANCHOR.MIDDLE)
        y += row_h

    add_takeaway(slide,
        "Next step: consolidate these 7 issues into the Baseline Report and carry them forward to the Target Operating Model design",
        color=NAVY)
    add_footer(slide,
        "Source: KB synthesis — 3 interview MDs + Org_Chart_Digital_Products.md + Excel data + Gemini review pass (Apr 2026)",
        "Structural Insights  |  26 / 28")


# ========== BASELINE-ONLY SLIDES (ported from build_baseline_report.py) ==========

def bl_master_map(prs):
    """AS-IS Master Demand Map — matrix of 6 teams × 5 stages, with Source + Types columns."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Across both director branches, Digital Products runs 6+ disconnected intake channels — no two teams share the same demand flow")

    # Geometry — more columns now: area label · SOURCE · TYPES · 5 stages
    matrix_top = 1.55
    label_w = 1.25
    src_w = 1.80
    types_w = 1.70
    gap = 0.05
    stage_x0 = 0.5 + label_w + gap + src_w + gap + types_w + gap
    stage_w = (12.33 - (stage_x0 - 0.5)) / 5
    header_h = 0.55
    row_h = 0.66

    # APPS label above
    add_text(slide, Inches(0.5), Inches(matrix_top - 0.25), Inches(3), Inches(0.2),
             "APPLICATIONS — Director IT Business Applications", size=8, bold=True, color=BLUE, font=FONT)

    # Column headers — SOURCE + TYPES + 5 stages
    add_rect(slide, Inches(0.5 + label_w + gap), Inches(matrix_top),
             Inches(src_w), Inches(header_h), ACCENT)
    add_text(slide, Inches(0.5 + label_w + gap), Inches(matrix_top + 0.03),
             Inches(src_w), Inches(0.27),
             "SOURCE", size=9, bold=True, color=WHITE, font=FONT, align=PP_ALIGN.CENTER)
    add_text(slide, Inches(0.5 + label_w + gap), Inches(matrix_top + 0.28),
             Inches(src_w), Inches(0.24),
             "Who sends the demand", size=7, color=WHITE, font=FONT, align=PP_ALIGN.CENTER)

    add_rect(slide, Inches(0.5 + label_w + gap + src_w + gap), Inches(matrix_top),
             Inches(types_w), Inches(header_h), ACCENT)
    add_text(slide, Inches(0.5 + label_w + gap + src_w + gap), Inches(matrix_top + 0.03),
             Inches(types_w), Inches(0.27),
             "TYPES", size=9, bold=True, color=WHITE, font=FONT, align=PP_ALIGN.CENTER)
    add_text(slide, Inches(0.5 + label_w + gap + src_w + gap), Inches(matrix_top + 0.28),
             Inches(types_w), Inches(0.24),
             "What kind of demand", size=7, color=WHITE, font=FONT, align=PP_ALIGN.CENTER)

    stages = [
        ("01 INTAKE", "Channels"),
        ("02 CLASSIFICATION", "Type, Size, Urgency"),
        ("03 PRIORITIZATION", "Scoring & Criteria"),
        ("04 PORTFOLIO", "Governance"),
        ("05 DELIVERY", "Execution"),
    ]
    for i, (num_name, sub) in enumerate(stages):
        x = Inches(stage_x0 + i * (stage_w + gap))
        add_rect(slide, x, Inches(matrix_top), Inches(stage_w), Inches(header_h), NAVY)
        add_text(slide, x, Inches(matrix_top + 0.03), Inches(stage_w), Inches(0.27),
                 num_name, size=8, bold=True, color=WHITE, font=FONT, align=PP_ALIGN.CENTER)
        add_text(slide, x, Inches(matrix_top + 0.28), Inches(stage_w), Inches(0.24),
                 sub, size=7, color=WHITE, font=FONT, align=PP_ALIGN.CENTER)

    # Area data — concise source + types for table readability
    areas_data = [
        {
            "area": "GIS",
            "source": "E&G Ops · T&D · Asset Mgmt · Planning · OSG · Compliance",
            "types":  "Break/fix · Small · Large (PDD)",
            "cells":  ["Jira + Email + Mtgs + IT-Now", "Small vs. PDD", "Business drives; P&T face", "PDD approval", "Internal + SWF"],
            "colors": [YELLOW_GAP, YELLOW_GAP, YELLOW_GAP, GREEN_OK, GREEN_OK],
            "branch": "APPS",
        },
        {
            "area": "Corporate",
            "source": "Finance · HR · Procurement · Legal · Training · Regulatory",
            "types":  "Regulatory · PDDs · Small/Medium",
            "cells":  ["Jira + Confluence + Word", "PDD driver; overlap", "PDD driver", "PDD approval", "Internal + Vendor"],
            "colors": [YELLOW_GAP, RED_CRITICAL, YELLOW_GAP, GREEN_OK, YELLOW_GAP],
            "branch": "APPS",
        },
        {
            "area": "G&E",
            "source": "T&D Ops · Field Services · Vendors · Cyber & Legal",
            "types":  "Version upgrades · Ad-hoc",
            "cells":  ["Email only (biz/vendor)", "Depends on nature", "Operational = urgent", "Vendor sets timeline", "Vendor"],
            "colors": [RED_CRITICAL, RED_CRITICAL, RED_CRITICAL, RED_CRITICAL, RED_CRITICAL],
            "branch": "APPS",
        },
        {
            "area": "Global SAP\n(Iberdrola)",
            "source": "Iberdrola Finance · Controlling · Procurement · Accenture",
            "types":  "Incidents · Small · Minor (RITM) · Major",
            "cells":  ["IT Now + RITM + Jira + SP", "By type (4 channels)", "Incidents first (P&T)", "PDD / SharePoint", "Accenture + SWF"],
            "colors": [GREEN_OK, GREEN_OK, RED_CRITICAL, YELLOW_GAP, YELLOW_GAP],
            "branch": "APPS",
        },
        {
            "area": "Smart Metering\n(Device & Metering)",
            "source": "Customer Ops · Field Services · Billing · BRS · Vendors",
            "types":  "Incidents · Large · Small",
            "cells":  ["Email (IT Now)", "IT Now template", "Criticality + deadline", "PPM Now + weekly mtg", "Vendor (100%)"],
            "colors": [YELLOW_GAP, GREEN_OK, YELLOW_GAP, GREEN_OK, RED_CRITICAL],
            "branch": "CS",
        },
        {
            "area": "Customer Services\n(4 teams TBC)",
            "source": "7 OPCOs via BRS · Travis PMO · Business",
            "types":  "Regulatory · Large · Small/Medium · IT-funded",
            "cells":  ["Director-level only", "Travis / BRS / IT-funded", "\"Intrinsic common sense\"", "No branch tool", "SWF + Accenture + Vendors"],
            "colors": [RED_CRITICAL, YELLOW_GAP, RED_CRITICAL, RED_CRITICAL, YELLOW_GAP],
            "branch": "CS",
        },
    ]

    y = matrix_top + header_h + 0.05
    prev_branch = None
    for rec in areas_data:
        if prev_branch == "APPS" and rec["branch"] == "CS":
            add_rect(slide, Inches(0.5), Inches(y), Inches(12.33), Inches(0.04), ACCENT)
            add_text(slide, Inches(0.5), Inches(y + 0.06), Inches(7), Inches(0.2),
                     "CUSTOMER SERVICES / SAP — Director Customer Services Apps",
                     size=8, bold=True, color=NAVY, font=FONT)
            y += 0.28

        label_color = NAVY if rec["branch"] == "CS" else BLUE

        # Area label
        add_rect(slide, Inches(0.5), Inches(y), Inches(label_w), Inches(row_h - 0.06), label_color)
        add_text(slide, Inches(0.5), Inches(y), Inches(label_w), Inches(row_h - 0.06),
                 rec["area"], size=9, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # SOURCE cell
        src_x = 0.5 + label_w + gap
        add_rect(slide, Inches(src_x), Inches(y), Inches(src_w), Inches(row_h - 0.06), LIGHT_GRAY)
        add_text(slide, Inches(src_x + 0.08), Inches(y + 0.04),
                 Inches(src_w - 0.16), Inches(row_h - 0.14),
                 rec["source"], size=8, color=CHARCOAL, font=FONT, italic=True,
                 anchor=MSO_ANCHOR.MIDDLE)

        # TYPES cell
        tp_x = src_x + src_w + gap
        add_rect(slide, Inches(tp_x), Inches(y), Inches(types_w), Inches(row_h - 0.06), LIGHT_GRAY)
        add_text(slide, Inches(tp_x + 0.08), Inches(y + 0.04),
                 Inches(types_w - 0.16), Inches(row_h - 0.14),
                 rec["types"], size=8, color=NAVY, font=FONT, bold=True,
                 anchor=MSO_ANCHOR.MIDDLE)

        # 5 stage cells (with maturity dot)
        for i, (text, color) in enumerate(zip(rec["cells"], rec["colors"])):
            x = stage_x0 + i * (stage_w + gap)
            add_rect(slide, Inches(x), Inches(y), Inches(stage_w), Inches(row_h - 0.06), LIGHT_GRAY)
            add_rect(slide, Inches(x + 0.06), Inches(y + 0.08),
                     Inches(0.14), Inches(0.14), color)
            add_text(slide, Inches(x + 0.25), Inches(y + 0.05),
                     Inches(stage_w - 0.30), Inches(row_h - 0.16),
                     text, size=8, color=CHARCOAL, font=FONT, anchor=MSO_ANCHOR.MIDDLE)
        y += row_h
        prev_branch = rec["branch"]

    add_takeaway(slide,
        "Zero standardization across units: 6 different sources, 6 intake models, 6 prioritization logics, only 1 area uses a portfolio tool — fragmentation is structural, not local",
        color=ACCENT)
    add_footer(slide, page_label="AS-IS Master Demand Map  |  11 / 28")


def bl_tooling(prs):
    """Tooling landscape across both branches."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Eight platforms circulate between the two branches with near-zero overlap — the toolchain is fragmented by team, not by purpose")

    add_text(slide, Inches(0.5), Inches(1.50), Inches(6.1), Inches(0.35),
             "APPLICATIONS BRANCH — Director IT Business Applications", size=11, bold=True, color=BLUE, font=FONT)

    apps_tooling = [
        ("Jira", "Small demand + improvements", ["GIS", "Corporate", "Global SAP"], "Primary"),
        ("IT Now", "Incidents + break/fix", ["GIS", "Global SAP"], "Primary"),
        ("IT Now RITM", "Minor requests", ["Global SAP"], "Specific"),
        ("SharePoint + Teams", "Major project work", ["Global SAP"], "Specific"),
        ("Confluence", "Documentation / requirements", ["Corporate"], "Specific"),
        ("Word", "Business requirements", ["Corporate"], "Legacy"),
        ("Email + Meetings", "G&E intake (only channel)", ["G&E"], "GAP"),
        ("OneNote / Spreadsheets", "Informal PM", ["GIS"], "Workaround"),
    ]
    y = 1.90
    for tool, purpose, teams, category in apps_tooling:
        cat_color = {"Primary": BLUE, "Specific": MID_GRAY, "Legacy": YELLOW_GAP, "GAP": RED_CRITICAL, "Workaround": YELLOW_GAP}[category]
        add_rect(slide, Inches(0.5), Inches(y), Inches(0.08), Inches(0.48), cat_color)
        add_text(slide, Inches(0.68), Inches(y), Inches(2.0), Inches(0.28),
                 tool, size=10, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(0.68), Inches(y + 0.25), Inches(2.0), Inches(0.25),
                 purpose, size=8, color=CHARCOAL, font=FONT)
        add_text(slide, Inches(2.75), Inches(y), Inches(2.6), Inches(0.5),
                 ", ".join(teams), size=8, color=BLUE, font=FONT, italic=True,
                 anchor=MSO_ANCHOR.MIDDLE)
        add_rect(slide, Inches(5.50), Inches(y + 0.10), Inches(1.00), Inches(0.30),
                 cat_color, corner=True)
        add_text(slide, Inches(5.50), Inches(y + 0.10), Inches(1.00), Inches(0.30),
                 category.upper(), size=7, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        y += 0.53

    add_text(slide, Inches(6.85), Inches(1.50), Inches(6.0), Inches(0.35),
             "CUSTOMER SERVICES / SAP — Director Customer Services Apps", size=11, bold=True, color=NAVY, font=FONT)

    cs_tooling = [
        ("IT Now", "Incidents (template) — Metering", ["Smart Metering"], "Primary"),
        ("Email", "Small demand (Smart Metering)", ["Smart Metering"], "GAP"),
        ("Jira", "Used by PMs (Smart Metering)", ["Smart Metering"], "Specific"),
        ("Confluence", "Portfolio documentation (Metering only)", ["Smart Metering"], "Specific"),
        ("PPM Now", "Large projects — ONLY in Smart Metering", ["Smart Metering"], "Primary"),
        ("Travis PMO (tool TBC)", "Large project management", ["Branch-wide"], "Specific"),
        ("SWF + Accenture", "Build / delivery", ["Branch-wide"], "Specific"),
        ("4 other teams — tooling not yet captured", "FICA/Billing, CRM, SAP Arch, SAP Basis", ["TBC"], "GAP"),
    ]
    y = 1.90
    for tool, purpose, teams, category in cs_tooling:
        cat_color = {"Primary": BLUE, "Specific": MID_GRAY, "Legacy": YELLOW_GAP, "GAP": RED_CRITICAL, "Workaround": YELLOW_GAP}[category]
        add_rect(slide, Inches(6.85), Inches(y), Inches(0.08), Inches(0.48), cat_color)
        add_text(slide, Inches(7.03), Inches(y), Inches(2.0), Inches(0.28),
                 tool, size=10, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(7.03), Inches(y + 0.25), Inches(2.0), Inches(0.25),
                 purpose, size=8, color=CHARCOAL, font=FONT)
        add_text(slide, Inches(9.10), Inches(y), Inches(2.6), Inches(0.5),
                 ", ".join(teams), size=8, color=BLUE, font=FONT, italic=True,
                 anchor=MSO_ANCHOR.MIDDLE)
        add_rect(slide, Inches(11.80), Inches(y + 0.10), Inches(1.00), Inches(0.30),
                 cat_color, corner=True)
        add_text(slide, Inches(11.80), Inches(y + 0.10), Inches(1.00), Inches(0.30),
                 category.upper(), size=7, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        y += 0.53

    add_takeaway(slide,
        "PPM Now exists in only 1 of 9 sub-teams — Jira is the only tool with meaningful cross-team adoption (and it's fragmented across instances)",
        color=ACCENT)
    add_footer(slide, page_label="Tooling Landscape  |  12 / 28")


def bl_governance_forums(prs):
    """Governance forums — five recurring cadences across the portfolio."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Five recurring forums exist — but none covers portfolio-level decisions, and Manager Check-in participants are still TBC")

    col_widths = [3.0, 1.2, 4.0, 4.13]
    x_positions = [0.5]
    for w in col_widths[:-1]:
        x_positions.append(x_positions[-1] + w)

    headers = ["FORUM", "CADENCE", "PARTICIPANTS", "DECISION SCOPE"]
    y_header = 1.55
    add_rect(slide, Inches(0.5), Inches(y_header), Inches(12.33), Inches(0.45), NAVY)
    for i, h in enumerate(headers):
        add_text(slide, Inches(x_positions[i] + 0.1), Inches(y_header), Inches(col_widths[i]), Inches(0.45),
                 h, size=10, bold=True, color=WHITE, font=FONT, anchor=MSO_ANCHOR.MIDDLE)

    forums = [
        ("Critical Integrations", "Weekly", "Business Apps, P&S, Digital Products", "Integration decisions across business apps"),
        ("Daily BSS BI/BW Stand Up", "Weekly", "Data & Analytics Teams, Digital Products", "Daily BI/BW operations alignment"),
        ("Data, Analytics & Solutions Check-In", "Weekly", "Data & Analytics Teams, Digital Products", "Data & analytics demand review"),
        ("CX & Digital CMP Update", "Monthly", "CX Digital, Operations, CMP President, VP CS", "CX-level strategic updates"),
        ("Manager Check-in", "Weekly", "[To be confirmed]", "[To be confirmed]"),
    ]

    y = y_header + 0.5
    for i, (forum, cadence, participants, scope) in enumerate(forums):
        bg = LIGHT_GRAY if i % 2 == 0 else WHITE
        add_rect(slide, Inches(0.5), Inches(y), Inches(12.33), Inches(0.7), bg)
        add_text(slide, Inches(x_positions[0] + 0.1), Inches(y), Inches(col_widths[0]), Inches(0.7),
                 forum, size=10, bold=True, color=NAVY, font=FONT, anchor=MSO_ANCHOR.MIDDLE)
        cadence_color = YELLOW_GAP if "confirmed" in cadence.lower() else CHARCOAL
        add_text(slide, Inches(x_positions[1] + 0.1), Inches(y), Inches(col_widths[1]), Inches(0.7),
                 cadence, size=10, color=cadence_color, font=FONT, anchor=MSO_ANCHOR.MIDDLE)
        p_color = RED_CRITICAL if "confirmed" in participants.lower() else CHARCOAL
        add_text(slide, Inches(x_positions[2] + 0.1), Inches(y), Inches(col_widths[2]), Inches(0.7),
                 participants, size=9, color=p_color, font=FONT, anchor=MSO_ANCHOR.MIDDLE,
                 italic=("confirmed" in participants.lower()))
        s_color = RED_CRITICAL if "confirmed" in scope.lower() else CHARCOAL
        add_text(slide, Inches(x_positions[3] + 0.1), Inches(y), Inches(col_widths[3]), Inches(0.7),
                 scope, size=9, color=s_color, font=FONT, anchor=MSO_ANCHOR.MIDDLE,
                 italic=("confirmed" in scope.lower()))
        y += 0.72

    add_rect(slide, Inches(0.5), Inches(y + 0.1), Inches(12.33), Inches(0.78), LIGHT_GRAY)
    add_rect(slide, Inches(0.5), Inches(y + 0.1), Inches(0.08), Inches(0.78), ACCENT)
    add_text(slide, Inches(0.68), Inches(y + 0.15), Inches(12.0), Inches(0.25),
             "WHAT'S MISSING", size=9, bold=True, color=ACCENT, font=FONT)
    add_text(slide, Inches(0.68), Inches(y + 0.40), Inches(12.0), Inches(0.45),
             "No portfolio steering forum (decisions on trade-offs across branches, budget, capacity)  •  "
             "No vendor performance forum  •  Manager Check-in participants and scope pending definition",
             size=10, color=CHARCOAL, font=FONT)

    add_takeaway(slide,
        "The governance layer is operationally busy (4 weekly + 1 monthly) but strategically incomplete — a portfolio steering cadence must be introduced in the TOM",
        color=NAVY)
    add_footer(slide, page_label="Governance Forums  |  13 / 28")


def bl_pain_points_mece(prs):
    """MECE pain points across People · Process · Tools · Governance."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Pain points cluster MECE across People, Process, Tools and Governance — each category produces its own bottleneck")

    categories = [
        {
            "label": "PEOPLE",
            "color": NAVY,
            "items": [
                "Turnover — \"1 year to hire someone\" (Global SAP)",
                "Unclear ownership — \"3 people can run payment jobs\" (CS Director/Business-side lead  —  Customer Services)",
                "Capacity competition across 7 business OPCOs (IT CS)",
                "Heavy SWF + Accenture + Vendor dependency in delivery",
                "2 vacant Sr. Manager positions in CS Director's branch",
            ],
            "source": "Source: All 4 interview MDs"
        },
        {
            "label": "PROCESS",
            "color": NAVY,
            "items": [
                "Demand intake not standardized across 6+ channels (Apps Director)",
                "Prioritization is ad-hoc; 4 informal criteria per forum",
                "PDD is the only formal process — only for large projects",
                "Vendor sets estimates, timelines, scope (especially Metering)",
                "PMs and Tech Leads \"need to work in the same way\" (Apps Director)",
            ],
            "source": "Source: Interview_Carlos_Applications.md + Excel prioritization criteria"
        },
        {
            "label": "TOOLS",
            "color": ACCENT,
            "items": [
                "6+ disconnected intake channels (Jira, IT Now, RITM, SharePoint, Email, Word)",
                "No portfolio management tool in 5 of 6 units (only Metering uses PPM Now)",
                "No tool / process to manage vendor delivery",
                "Jira fragmentation across instances",
                "Informal PM stack (OneNote, Teams, Spreadsheets) for GIS",
            ],
            "source": "Source: Interview_Application_Champions.md — tool landscape matrix"
        },
        {
            "label": "GOVERNANCE",
            "color": ACCENT,
            "items": [
                "5 forums exist — none covers portfolio-level decisions",
                "Manager Check-in participants + scope still TBC",
                "Regulatory demand overrides size in every forum (no formal rule)",
                "P&T is the only cross-team intermediary for Business ↔ IT",
                "No enterprise-level reconciliation between Global SAP (Iberdrola) and the other SAP footprints",
            ],
            "source": "Source: Governance forums table + Interview_Carlos_Applications.md"
        },
    ]

    card_w = 6.1
    card_h = 2.45
    gap_x = 0.13
    gap_y = 0.15
    x0 = 0.5
    y0 = 1.55

    for i, c in enumerate(categories):
        col = i % 2
        row = i // 2
        x = x0 + col * (card_w + gap_x)
        y = y0 + row * (card_h + gap_y)

        add_rect(slide, Inches(x), Inches(y), Inches(card_w), Inches(card_h), LIGHT_GRAY)
        add_rect(slide, Inches(x), Inches(y), Inches(card_w), Inches(0.06), c["color"])
        add_text(slide, Inches(x + 0.25), Inches(y + 0.15), Inches(card_w - 0.5), Inches(0.35),
                 c["label"], size=13, bold=True, color=c["color"], font=FONT)
        items_text = "\n".join(["▪  " + it for it in c["items"]])
        add_text(slide, Inches(x + 0.25), Inches(y + 0.55), Inches(card_w - 0.5), Inches(1.55),
                 items_text, size=10, color=CHARCOAL, font=FONT)
        add_text(slide, Inches(x + 0.25), Inches(y + card_h - 0.30), Inches(card_w - 0.5), Inches(0.25),
                 c["source"], size=7, color=MID_GRAY, font=FONT, italic=True)

    add_takeaway(slide,
        "Each MECE category produces a distinct bottleneck — addressing only one dimension in the TOM will not resolve the system; all four must move together",
        color=NAVY)
    add_footer(slide, page_label="Pain Points — MECE  |  22 / 28")


def bl_cross_branch(prs):
    """Three cross-branch systemic issues that no local fix resolves."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Three systemic issues cut across both branches — no local fix resolves them; only the TOM can")

    issues = [
        {
            "title": "Prioritization has no shared model",
            "detail": (
                "Four informal criteria (strategic, operational urgency, regulatory, size) are applied per forum, per team, "
                "without a scoring rubric. Regulatory always wins regardless of size. No mechanism balances competing demands "
                "across the two branches."
            ),
            "applications": "PDD approval / impact × users / incidents first / criticality — all different",
            "customer_services": "\"Intrinsic common sense\" (Metering) + forum-based (director-level)",
        },
        {
            "title": "Taxonomy is undefined",
            "detail": (
                "What counts as \"small\", \"medium\", or \"large\" demand is defined per team. No hours / cost / complexity "
                "threshold exists. Without objective classification, prioritization logic cannot be automated, and forums "
                "debate the class before debating the priority."
            ),
            "applications": "PDD covers \"large\" but \"small vs medium\" is informal",
            "customer_services": "Travis PMO owns \"large\" by convention, not definition",
        },
        {
            "title": "Vendor governance is absent",
            "detail": (
                "No process governs vendor delivery in either branch. Vendors estimate scope, timelines, and cost. "
                "Smart Metering is 100% vendor-delivered — vendors do requirements AND code. In G&E, the vendor sets "
                "its own timeline on version upgrades."
            ),
            "applications": "Vendors co-build; no performance framework",
            "customer_services": "\"All portfolio in vendor\" (Smart Metering) — full outsourcing risk",
        },
    ]

    row_h = 1.60
    y = 1.55
    for i, issue in enumerate(issues):
        add_rect(slide, Inches(0.5), Inches(y), Inches(12.33), Inches(row_h - 0.08), LIGHT_GRAY)
        add_rect(slide, Inches(0.5), Inches(y), Inches(0.08), Inches(row_h - 0.08), ACCENT)
        add_rect(slide, Inches(0.65), Inches(y + 0.15), Inches(0.6), Inches(0.6), ACCENT)
        add_text(slide, Inches(0.65), Inches(y + 0.15), Inches(0.6), Inches(0.6),
                 f"0{i+1}", size=16, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, Inches(1.40), Inches(y + 0.10), Inches(11.5), Inches(0.35),
                 issue["title"], size=13, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(1.40), Inches(y + 0.45), Inches(11.5), Inches(0.55),
                 issue["detail"], size=9, color=CHARCOAL, font=FONT)
        add_text(slide, Inches(1.40), Inches(y + row_h - 0.45), Inches(3), Inches(0.22),
                 "APPLICATIONS:", size=8, bold=True, color=BLUE, font=FONT)
        add_text(slide, Inches(3.10), Inches(y + row_h - 0.45), Inches(3.5), Inches(0.35),
                 issue["applications"], size=8, color=CHARCOAL, font=FONT, italic=True)
        add_text(slide, Inches(6.95), Inches(y + row_h - 0.45), Inches(3), Inches(0.22),
                 "CUSTOMER SERVICES:", size=8, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(9.05), Inches(y + row_h - 0.45), Inches(3.7), Inches(0.35),
                 issue["customer_services"], size=8, color=CHARCOAL, font=FONT, italic=True)
        y += row_h

    add_takeaway(slide,
        "These three are enterprise-level issues — the TOM owns them; no team-level fix will solve prioritization, taxonomy, or vendor governance on its own",
        color=NAVY)
    add_footer(slide, page_label="Cross-Branch Systemic Issues  |  25 / 28")


def bl_risks(prs):
    """Five compounding risks if TOM is not pursued this quarter."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Five compounding risks if the current operating model is carried into 30–50% growth unchanged")

    risks = [
        {
            "num": "R1",
            "title": "Unknown demand volume meets scaling load",
            "detail": "30–50% growth applied to unmeasured channels = capacity shortfall cannot be forecast. Hiring decisions remain reactive.",
            "severity": "CRITICAL",
        },
        {
            "num": "R2",
            "title": "Vendor control of timelines and estimates",
            "detail": "Smart Metering is 100% vendor-delivered. G&E upgrades follow vendor schedules. Without governance, budget overruns stay invisible until delivery.",
            "severity": "CRITICAL",
        },
        {
            "num": "R3",
            "title": "ROI traceability breaks",
            "detail": "Without portfolio tooling or value-stream lens, IT spend cannot be traced back to business outcomes. Executive defensibility degrades.",
            "severity": "HIGH",
        },
        {
            "num": "R4",
            "title": "Leadership vacancies prolonged",
            "detail": "CRM Sr. Manager and SAP & CS Architecture positions remain open. Decisions default to director level across ~24 FTE — not sustainable.",
            "severity": "HIGH",
        },
        {
            "num": "R5",
            "title": "Global vs SAP drift",
            "detail": "With no enterprise-level reconciliation, the two SAP footprints evolve independently. Integration cost grows silently each quarter.",
            "severity": "MEDIUM",
        },
    ]

    severity_color = {"CRITICAL": RED_CRITICAL, "HIGH": YELLOW_GAP, "MEDIUM": MID_GRAY}

    y = 1.55
    row_h = 0.95
    for r in risks:
        add_rect(slide, Inches(0.5), Inches(y), Inches(12.33), Inches(row_h - 0.08), LIGHT_GRAY)
        sev_color = severity_color[r["severity"]]
        add_rect(slide, Inches(0.5), Inches(y), Inches(0.08), Inches(row_h - 0.08), sev_color)
        add_rect(slide, Inches(0.65), Inches(y + 0.15), Inches(0.7), Inches(0.55), sev_color)
        add_text(slide, Inches(0.65), Inches(y + 0.15), Inches(0.7), Inches(0.55),
                 r["num"], size=14, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, Inches(1.5), Inches(y + 0.10), Inches(9.5), Inches(0.3),
                 r["title"], size=12, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(1.5), Inches(y + 0.42), Inches(9.5), Inches(0.45),
                 r["detail"], size=9, color=CHARCOAL, font=FONT)
        add_rect(slide, Inches(11.20), Inches(y + 0.30), Inches(1.55), Inches(0.30), sev_color, corner=True)
        add_text(slide, Inches(11.20), Inches(y + 0.30), Inches(1.55), Inches(0.30),
                 r["severity"], size=8, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        y += row_h

    add_takeaway(slide,
        "These risks compound — the longer the TOM is deferred, the harder the reset. The Baseline recommends beginning TOM design this quarter",
        color=ACCENT)
    add_footer(slide, page_label="Risks if Not Addressed  |  27 / 28")


def bl_tom_principles(prs):
    """Seven TOM design principles + two baseline-specific additions."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Seven principles guide the Target Operating Model — each answers one of the 7 structural issues and is grounded in prior NTT work with Avangrid")

    principles = [
        ("01", "Single Front Door", "One structured intake process for all demand types, replacing the 6+ disconnected channels."),
        ("02", "Value-Based Prioritization", "Standardized scoring framework across both branches — objective criteria, not ad-hoc forums."),
        ("03", "Capacity-Aware Assignment", "Assignment paired with real-time capacity visibility — not manual handoffs with opaque availability."),
        ("04", "Standardized Execution (Jira)", "Jira as the execution platform across Digital Products — consistent tracking, management, delivery."),
        ("05", "Integrated Toolchain", "Intake + Jira + PPM Now — three core platforms interoperable across Applications and Customer Services."),
        ("06", "Automated Reporting (SSO)", "Real-time dashboards from a single source of truth — replacing manual data gathering."),
        ("07", "Defined RACI per Lifecycle Phase", "Unambiguous ownership and handoff protocols at every stage — from intake to deployment."),
    ]

    col_w = 6.15
    item_h = 1.18
    x0_left = 0.5
    x0_right = 6.78
    y0 = 1.55

    for i, (num, title, desc) in enumerate(principles):
        col = 0 if i < 4 else 1
        row = i if i < 4 else i - 4
        x = x0_left if col == 0 else x0_right
        y = y0 + row * item_h

        add_rect(slide, Inches(x), Inches(y), Inches(col_w), Inches(item_h - 0.10), LIGHT_GRAY)
        add_rect(slide, Inches(x), Inches(y), Inches(col_w), Inches(0.05), BLUE)
        add_text(slide, Inches(x + 0.25), Inches(y + 0.15), Inches(0.7), Inches(0.4),
                 num, size=20, bold=True, color=BLUE, font=FONT)
        add_text(slide, Inches(x + 1.0), Inches(y + 0.18), Inches(col_w - 1.25), Inches(0.35),
                 title, size=12, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(x + 1.0), Inches(y + 0.55), Inches(col_w - 1.25), Inches(0.55),
                 desc, size=9, color=CHARCOAL, font=FONT)

    callout_y = 6.10
    add_rect(slide, Inches(6.78), Inches(callout_y), Inches(6.15), Inches(0.42), ACCENT)
    add_text(slide, Inches(6.95), Inches(callout_y), Inches(6.0), Inches(0.42),
             "+ From this Baseline: BRANCH-UNIFIED (Apps + CS) and PRODUCT-ORIENTED (Value Streams, not silos)",
             size=10, bold=True, color=WHITE, font=FONT, anchor=MSO_ANCHOR.MIDDLE)

    add_takeaway(slide,
        "These principles are not new — they extend the Transformation Vision already defined in SAP CoE New Ways of Working v2.0 with two baseline-specific additions",
        color=NAVY)
    add_footer(slide, page_label="TOM Design Principles  |  28 / 28")


# ========== PLACEHOLDER SLIDES (Capacity & Portfolio Data from interviews) ==========

def _champion_placeholder_grid(slide, *, y0, metrics, accent):
    """5-area × N-metrics placeholder grid — each cell is a dotted drop zone."""
    areas = [
        ("GIS",            "GIS Sr. Manager / Champion",      GREEN_OK),
        ("Corporate",      "Corporate Champion (Tech Lead)",        ACCENT),
        ("G&E",            "G&E Champion (Tech Lead, 70-90 apps)",   RED_CRITICAL),
        ("Global SAP",     "Global SAP Champion (Tech Lead)",    BLUE),
        ("Smart Metering", "Smart Metering Champion (PM)",    YELLOW_GAP),
    ]
    # Header row with metric names
    label_col_w = 2.2
    metric_w = (12.33 - label_col_w - 0.10) / len(metrics)

    # Metric headers
    for i, m in enumerate(metrics):
        x = 0.5 + label_col_w + 0.10 + i * metric_w
        add_rect(slide, Inches(x), Inches(y0), Inches(metric_w - 0.08), Inches(0.40), NAVY)
        add_text(slide, Inches(x), Inches(y0), Inches(metric_w - 0.08), Inches(0.40),
                 m, size=10, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 5 area rows
    row_h = 0.62
    for r, (area, champion, color) in enumerate(areas):
        y = y0 + 0.44 + r * (row_h + 0.05)
        # Left label
        add_rect(slide, Inches(0.5), Inches(y), Inches(label_col_w), Inches(row_h), color)
        add_text(slide, Inches(0.62), Inches(y + 0.06), Inches(label_col_w - 0.2), Inches(0.28),
                 area, size=11, bold=True, color=WHITE, font=FONT)
        add_text(slide, Inches(0.62), Inches(y + 0.34), Inches(label_col_w - 0.2), Inches(0.24),
                 champion, size=8, color=WHITE, font=FONT, italic=True)
        # Empty cells for data
        for i, m in enumerate(metrics):
            x = 0.5 + label_col_w + 0.10 + i * metric_w
            add_rect(slide, Inches(x), Inches(y), Inches(metric_w - 0.08), Inches(row_h), LIGHT_GRAY,
                     line_color=accent, line_weight=Pt(0.75))
            add_text(slide, Inches(x), Inches(y + 0.10), Inches(metric_w - 0.08), Inches(0.25),
                     "[to populate]", size=8, color=MID_GRAY, font=FONT, italic=True,
                     align=PP_ALIGN.CENTER)


def slide_capacity_data_placeholder(prs):
    """Capacity Baseline — placeholder for data captured from champion interviews."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Capacity Baseline — to be populated from the five champion interviews (headcount · mix · utilization per area)")

    # Status banner
    add_rect(slide, Inches(0.5), Inches(1.35), Inches(12.33), Inches(0.40), ACCENT)
    add_text(slide, Inches(0.7), Inches(1.35), Inches(12), Inches(0.40),
             "STATUS  —  Pending Wave 1 discovery sessions  ·  target close: May 2026  ·  Source: NTT interviews + Excel baseline template",
             size=10, bold=True, color=WHITE, font=FONT, anchor=MSO_ANCHOR.MIDDLE)

    _champion_placeholder_grid(slide, y0=1.95,
        metrics=["FTE  (internal)", "SWF / Accenture / Vendor", "% BAU vs. STRATEGIC", "UTILIZATION %", "KEY-PERSON RISKS"],
        accent=BLUE)

    # Legend + how-to
    add_rect(slide, Inches(0.5), Inches(5.70), Inches(12.33), Inches(0.80), LIGHT_GRAY)
    add_rect(slide, Inches(0.5), Inches(5.70), Inches(0.08), Inches(0.80), BLUE)
    add_text(slide, Inches(0.70), Inches(5.78), Inches(12), Inches(0.24),
             "HOW THIS GETS POPULATED", size=10, bold=True, color=NAVY, font=FONT)
    add_text(slide, Inches(0.70), Inches(6.02), Inches(12), Inches(0.48),
             "Each champion fills their row during the dynamic (Applications Team Discovery Canvas  —  Block 1 Capacity Landscape). "
             "Numbers reconciled against the Excel baseline. Empty cells flagged HYPOTHESIS until verified.",
             size=10, color=CHARCOAL, font=FONT)

    add_takeaway(slide,
        "Once populated, this becomes the quantitative baseline for capacity planning in the TOM — and the reference for Block 2 Resource Mix analysis",
        color=NAVY)
    add_footer(slide, page_label="Capacity Data  |  20 / 28")


def slide_portfolio_data_placeholder(prs):
    """Portfolio Inventory — placeholder for data captured from champion interviews."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Portfolio Inventory — to be populated with projects in flight, vendor mix and cost distribution per area")

    # Status banner
    add_rect(slide, Inches(0.5), Inches(1.35), Inches(12.33), Inches(0.40), ACCENT)
    add_text(slide, Inches(0.7), Inches(1.35), Inches(12), Inches(0.40),
             "STATUS  —  Pending Wave 1 discovery sessions  ·  target close: May 2026  ·  Source: PDD workbook + PPM Now (where available) + champion interviews",
             size=10, bold=True, color=WHITE, font=FONT, anchor=MSO_ANCHOR.MIDDLE)

    _champion_placeholder_grid(slide, y0=1.95,
        metrics=["PROJECTS  IN FLIGHT", "PDD COUNT  (current yr)", "VENDOR MIX %", "COST / QUARTER", "PORTFOLIO TOOL?"],
        accent=ACCENT)

    # Legend + how-to
    add_rect(slide, Inches(0.5), Inches(5.70), Inches(12.33), Inches(0.80), LIGHT_GRAY)
    add_rect(slide, Inches(0.5), Inches(5.70), Inches(0.08), Inches(0.80), ACCENT)
    add_text(slide, Inches(0.70), Inches(5.78), Inches(12), Inches(0.24),
             "HOW THIS GETS POPULATED", size=10, bold=True, color=NAVY, font=FONT)
    add_text(slide, Inches(0.70), Inches(6.02), Inches(12), Inches(0.48),
             "Each champion fills their row during the dynamic (Applications Team Discovery Canvas  —  Block 2 Resource Mix & Cost). "
             "Cross-validated against PDD workbook + PPM Now. Smart Metering is the benchmark row (100% populated today).",
             size=10, color=CHARCOAL, font=FONT)

    add_takeaway(slide,
        "Portfolio data closes the ROI-traceability gap — without it the TOM business case stays qualitative",
        color=NAVY)
    add_footer(slide, page_label="Portfolio Data  |  21 / 28")


# ========== MAIN ==========

def build():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    # ── Cover + Deck Roadmap (parent slide) ─────────────────────────
    slide_1_cover(prs)                                # 01 — no nav strip (cover stands alone)
    slide_roadmap(prs)                                # 02 — parent / roadmap slide

    # ── Section 01 · EXECUTIVE SUMMARY (03–05) ──────────────────────
    slide_context(prs);                 _add_nav_to_last(prs, 0)   # 03
    slide_executive_summary(prs);       _add_nav_to_last(prs, 0)   # 04
    slide_methodology(prs);             _add_nav_to_last(prs, 0)   # 05

    # ── Section 02 · ORGANIZATION (06–10) ───────────────────────────
    slide_org_context(prs);             _add_nav_to_last(prs, 1)   # 06
    slide_apps_branch_overview(prs);    _add_nav_to_last(prs, 1)   # 07
    slide_sap_branch_overview(prs);     _add_nav_to_last(prs, 1)   # 08
    slide_sap_landscape_detail(prs);    _add_nav_to_last(prs, 1)   # 09
    slide_sap_opco_demand(prs);         _add_nav_to_last(prs, 1)   # 10  NEW

    # ── Section 03 · DEMAND FLOW (11–19) ────────────────────────────
    bl_master_map(prs);                 _add_nav_to_last(prs, 2)   # 11
    bl_tooling(prs);                    _add_nav_to_last(prs, 2)   # 12
    bl_governance_forums(prs);          _add_nav_to_last(prs, 2)   # 13
    slide_4_gis(prs);                   _add_nav_to_last(prs, 2)   # 14
    slide_5_corporate(prs);             _add_nav_to_last(prs, 2)   # 15
    slide_6_ge(prs);                    _add_nav_to_last(prs, 2)   # 16
    slide_7_global_sap(prs);            _add_nav_to_last(prs, 2)   # 17
    slide_8_smart_metering(prs);        _add_nav_to_last(prs, 2)   # 18
    slide_10_customer_services(prs);    _add_nav_to_last(prs, 2)   # 19

    # ── Section 04 · CAPACITY & PORTFOLIO (20–21) — data from interviews ──
    slide_capacity_data_placeholder(prs);   _add_nav_to_last(prs, 3)   # 20
    slide_portfolio_data_placeholder(prs);  _add_nav_to_last(prs, 3)   # 21

    # ── Section 05 · PAIN POINTS (22–25) ────────────────────────────
    bl_pain_points_mece(prs);           _add_nav_to_last(prs, 4)   # 22
    slide_10_apps_bottlenecks(prs);     _add_nav_to_last(prs, 4)   # 23
    slide_11_cs_bottlenecks(prs);       _add_nav_to_last(prs, 4)   # 24
    bl_cross_branch(prs);               _add_nav_to_last(prs, 4)   # 25

    # ── Section 06 · INSIGHTS & TOM (26–28) ─────────────────────────
    slide_11_insights(prs);             _add_nav_to_last(prs, 5)   # 26
    bl_risks(prs);                      _add_nav_to_last(prs, 5)   # 27
    bl_tom_principles(prs);             _add_nav_to_last(prs, 5)   # 28

    out = "/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Deliverables/Integrated_Deck/Integrated_Baseline_DemandMap_v1.pptx"
    prs.save(out)
    print(f"Saved: {out}")
    print(f"Slides: {len(prs.slides)}")


if __name__ == "__main__":
    build()
