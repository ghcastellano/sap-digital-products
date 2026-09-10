"""
Demand Map — Avangrid SAP Digital Products
NTT DATA engagement | Current State Diagnosis

Grounded in Knowledge Base:
- Interview_Application_Champions.md (6 champions — GIS, Corporate, G&E, Global SAP, Smart Metering)
- Interview_Stephanie_Kumar_IT_CS.md (IT Customer Service)
- Interview_Carlos_Applications.md (Director of Applications)
- SAP CoE New Ways of Working v2.0.pptx, Slide 3 (30-50% growth, fragmentation context)

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
    ("01", "FRAMING", "Context · Exec Summary · Methodology"),
    ("02", "ORGANIZATION", "Org · Applications · SAP / Customer Services"),
    ("03", "AS-IS DRILL-DOWNS", "6 team deep-dives across 5 stages"),
    ("04", "BOTTLENECKS", "Apps · Customer Services"),
    ("05", "INSIGHTS", "Structural issues + next steps"),
]


def add_nav_strip(slide, current_idx):
    """Thin top roadmap bar — 5 pills, current section highlighted.

    Anchored above the action title. current_idx in [0..4], or None to skip.
    """
    if current_idx is None:
        return
    strip_y = 0.08
    strip_h = 0.22
    pill_gap = 0.05
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
                 f"{num}  ·  {name}",
                 size=8, bold=is_current, color=fg, font=FONT,
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
        "Deck roadmap — five sections, grounded in Wave 1 Step 1 (Enablement): evidence first, interpretation second, recommendations last")

    # Horizontal strip of 5 section cards
    card_w = 2.45
    card_h = 3.45
    gap_x = 0.08
    y0 = 1.55

    # Section details for the big roadmap cards
    sections_detail = [
        {
            "num": "01",
            "name": "FRAMING",
            "slides": "Slides 03 – 06",
            "what": "Why we are here. Executive summary. Methodology (Wave 1 Step 1).",
            "feeds": "Sets the lens for every finding that follows.",
        },
        {
            "num": "02",
            "name": "ORGANIZATION",
            "slides": "Slides 07 – 10",
            "what": "Org context (nine sub-teams, two vacancies). Applications branch. SAP / Customer Services branch + CMP · UIL · NY · Digital CoE · Infrastructure detail.",
            "feeds": "Names the players — drill-downs reference these teams one-by-one.",
        },
        {
            "num": "03",
            "name": "AS-IS DRILL-DOWNS",
            "slides": "Slides 11 – 16",
            "what": "Six deep-dives (GIS · Corporate · G&E · Global SAP · Smart Metering · Customer Services Apps) across 5 lifecycle stages × 4 lenses (Channels · Roles · Tools · Governance).",
            "feeds": "Ground truth — every bottleneck and insight traces back to these cells.",
        },
        {
            "num": "04",
            "name": "BOTTLENECKS",
            "slides": "Slides 17 – 18",
            "what": "Eight bottlenecks — four per branch (Applications + Customer Services). Each card names stage, team, impact and evidence.",
            "feeds": "Converts evidence into named risks — prioritized for TOM response.",
        },
        {
            "num": "05",
            "name": "INSIGHTS",
            "slides": "Slide 19",
            "what": "Seven structural issues across both branches — paired with the TOM response that resolves each.",
            "feeds": "Exit-ramp into Step 2 (Define Lifecycle) of Wave 1.",
        },
    ]

    for i, s in enumerate(sections_detail):
        x = 0.5 + i * (card_w + gap_x)
        # Card background
        add_rect(slide, Inches(x), Inches(y0), Inches(card_w), Inches(card_h), LIGHT_GRAY)
        # Top navy band
        add_rect(slide, Inches(x), Inches(y0), Inches(card_w), Inches(0.50), NAVY)
        add_text(slide, Inches(x + 0.18), Inches(y0 + 0.08), Inches(card_w - 0.3), Inches(0.18),
                 f"SECTION {s['num']}", size=8, bold=True, color=BLUE, font=FONT)
        add_text(slide, Inches(x + 0.18), Inches(y0 + 0.25), Inches(card_w - 0.3), Inches(0.25),
                 s["name"], size=12, bold=True, color=WHITE, font=FONT)
        # Slides range
        add_text(slide, Inches(x + 0.18), Inches(y0 + 0.60), Inches(card_w - 0.3), Inches(0.22),
                 s["slides"], size=9, italic=True, color=ACCENT, font=FONT)
        # What
        add_text(slide, Inches(x + 0.18), Inches(y0 + 0.95), Inches(card_w - 0.3), Inches(0.22),
                 "WHAT YOU'LL SEE", size=8, bold=True, color=MID_GRAY, font=FONT)
        add_text(slide, Inches(x + 0.18), Inches(y0 + 1.20), Inches(card_w - 0.3), Inches(1.55),
                 s["what"], size=9, color=CHARCOAL, font=FONT)
        # Feeds
        add_text(slide, Inches(x + 0.18), Inches(y0 + 2.80), Inches(card_w - 0.3), Inches(0.22),
                 "HOW IT FEEDS", size=8, bold=True, color=ACCENT, font=FONT)
        add_text(slide, Inches(x + 0.18), Inches(y0 + 3.05), Inches(card_w - 0.3), Inches(0.40),
                 s["feeds"], size=9, italic=True, color=BLUE, font=FONT)
        # Arrow between cards
        if i < len(sections_detail) - 1:
            arrow_x = x + card_w - 0.02
            arrow = slide.shapes.add_shape(
                MSO_SHAPE.RIGHT_ARROW,
                Inches(arrow_x), Inches(y0 + card_h / 2 - 0.10),
                Inches(0.12), Inches(0.16))
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = ACCENT
            arrow.line.fill.background()
            arrow.shadow.inherit = False

    # Orientation strip below — how to read the nav bar on every slide
    orient_y = 5.20
    add_rect(slide, Inches(0.5), Inches(orient_y), Inches(12.33), Inches(1.15), LIGHT_GRAY)
    add_rect(slide, Inches(0.5), Inches(orient_y), Inches(0.08), Inches(1.15), ACCENT)
    add_text(slide, Inches(0.70), Inches(orient_y + 0.08), Inches(12), Inches(0.28),
             "HOW TO READ THIS DECK", size=10, bold=True, color=NAVY, font=FONT)
    add_text(slide, Inches(0.70), Inches(orient_y + 0.38), Inches(12), Inches(0.70),
             "Every slide carries a thin roadmap strip at the top — five pills, current section highlighted in navy. "
             "Read each slide bottom-up: action title states the conclusion, content is the evidence, takeaway is the \"so what.\" "
             "Cross-references use format [Slide XX].",
             size=10, color=CHARCOAL, font=FONT)

    add_takeaway(slide,
        "You are here → Section 01 (Framing). Next: the organization that generates the demand we are mapping.",
        color=NAVY)
    add_footer(slide, page_label="Deck Roadmap  |  02 / 19")


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
             "▸  Digital Products (formerly SAP CoE) is being re-founded under VP Vitor Cruz, merging Applications (Carlos Garbin) and Customer Services Apps (Kumar Jaladanki) into one ~50-FTE organization.\n\n"
             "▸  Demand volume has grown 30–50% across SAP and Applications in the last 24 months — with no unified intake, classification, or prioritization model in place.\n\n"
             "▸  Six champion interviews and two director-level sessions (Carlos, Stephanie & Kumar) surfaced six disconnected intake channels and five different prioritization logics.\n\n"
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
    add_footer(slide, page_label="Project Context  |  03 / 19")


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
    add_footer(slide, page_label="Executive Summary  |  04 / 19")


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
         "Six 1:1 champion interviews (GIS, Corporate, G&E, Global SAP, Smart Metering) + two director sessions (Carlos; Stephanie & Kumar). Captured intake channels, tools, roles, pain points and aspirations per team."),
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
    add_footer(slide, page_label="Methodology  |  05 / 19")


def slide_apps_branch_overview(prs):
    """Applications branch overview — nature of work, apps, priorities, interactions."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Applications runs Avangrid's non-SAP operational backbone — four teams supporting Electric & Gas, GIS, Corporate and Iberdrola Global SAP")

    # Branch band
    add_rect(slide, Inches(0.5), Inches(1.32), Inches(12.33), Inches(0.38), BLUE)
    add_text(slide, Inches(0.5), Inches(1.32), Inches(12.33), Inches(0.38),
             "BRANCH 1  —  IT BUSINESS APPLICATIONS  ·  Director: Carlos Garbin  ·  ~26 FTE  ·  4 teams",
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
         "Four technology domains under Carlos",
         [
             "▪  GIS & OSG  —  SAP HANA, ARCFM, GIS web apps, data replication, SWF platform (George Porto, ~6 FTE).",
             "▪  Electric & Gas Ops  —  70–90 apps, android/tablets, doc viewer, 811 gas lines (Bruno Moreto + Andrew McGibbon, ~7 FTE).",
             "▪  Corporate Applications  —  5-6 PMs, 2 tech leads, corporate + network apps, SWF tooling (Sarath + Amit, ~8 FTE).",
             "▪  Global SAP (Iberdrola)  —  FI, PO, Budget — Spain-side SAP module support (Vanessa + Michael, ~5 FTE)."
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
             "▪  With SAP branch: Global SAP ↔ Kumar's Local SAP (Billing, FICA); P&T is the intermediary with Business.",
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
    add_footer(slide, page_label="Applications Branch Overview  |  08 / 19")


def slide_sap_branch_overview(prs):
    """Customer Services Apps (SAP branch) overview — CMP / UIL / Digital CoE / Infra."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Customer Services Apps serves 7 OPCOs across three BRS groups — with Local SAP (Billing, CRM, FICA, Metering) as the operational core and two Sr. Manager vacancies")

    # Branch band
    add_rect(slide, Inches(0.5), Inches(1.32), Inches(12.33), Inches(0.38), NAVY)
    add_text(slide, Inches(0.5), Inches(1.32), Inches(12.33), Inches(0.38),
             "BRANCH 2  —  CUSTOMER SERVICES APPS  ·  Director: Kumar Jaladanki  ·  ~24 FTE  ·  5 teams (2 Sr. Manager vacancies)",
             size=11, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 4 cards: Nature of Work | BRS / OPCO Landscape | Teams | Way of Working
    cards = [
        ("NATURE OF WORK",
         "Customer-facing SAP + metering + contact center",
         [
             "▪  Local SAP (Avangrid US) — Billing, FICA, CRM, Metering — distinct from Global SAP (Iberdrola Spain).",
             "▪  Demand split across BRS (business), PMO (Travis — large projects) and IT-funded (Kumar's teams).",
             "▪  Incidents + small/medium demand dominate volume; large projects run via annual PDD.",
             "▪  Software Factory (SWF) executes; Accenture supports incidents + projects."
         ]),
        ("BRS  /  OPCO LANDSCAPE",
         "3 BRS groups representing 7 operating companies",
         [
             "▪  BRS UIL  —  United Illuminating, SCG, CNG, BGC (Connecticut cluster).",
             "▪  BRS CMP  —  Central Maine Power (Maine operations).",
             "▪  BRS NY  —  NYSEG + RGE (New York cluster).",
             "▪  Plus Digital CoE (customer apps / websites) and Call Center (contact interactions — Kumar).",
             "▪  Each OPCO competes for the same IT capacity — multi-OPCO prioritization is a daily pressure."
         ]),
        ("TEAMS UNDER KUMAR",
         "5 sub-teams, heterogeneous maturity",
         [
             "▪  Ivelisse Velazquez  —  FICA, Billing & Collection (~8 FTE).",
             "▪  Vincent Lam  —  Device & Metering (Iram Chowdhury / Smart Metering sits here, ~8 FTE).",
             "▪  Nagi Reddy  —  SAP Basis (platform / infrastructure, ~2 FTE).",
             "▪  TBD (vacant)  —  Sr. Manager CRM & Contact Center Ops (~4 FTE) — leadership gap.",
             "▪  Open (vacant)  —  Sr. Manager SAP & CS Architecture (~2 FTE) — leadership gap."
         ]),
        ("WAY OF WORKING  &  INTERACTIONS",
         "Fragmented intake + vendor-heavy delivery",
         [
             "▪  Small/medium demand via Jira; large via PPM Now (only Smart Metering consistently).",
             "▪  Incidents: IT Now; Minor requests: RITM; Major projects: SharePoint + Teams.",
             "▪  Interactions: Kumar ↔ Travis (PMO for large projects); ↔ BRS leads (business demand); ↔ Vanessa (Global SAP).",
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
    add_footer(slide, page_label="Customer Services Branch Overview  |  09 / 19")


def slide_sap_landscape_detail(prs):
    """SAP landscape detail — CMP, UIL, BRS NY, Digital CoE, Infrastructure."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    add_action_title(slide,
        "The SAP branch serves five distinct business landscapes — three regional BRS groups plus Digital CoE and Infrastructure — each with its own demand profile and governance cadence")

    # Branch band
    add_rect(slide, Inches(0.5), Inches(1.32), Inches(12.33), Inches(0.32), NAVY)
    add_text(slide, Inches(0.5), Inches(1.32), Inches(12.33), Inches(0.32),
             "SAP  /  CUSTOMER SERVICES APPS  —  LANDSCAPE BREAKDOWN",
             size=10, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 5 cards — 3 BRS groups, Digital CoE, Infrastructure
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
            "scope": "Customer-facing apps, websites, mobile  —  Kumar directly",
            "bullets": [
                "▪  Kumar owns Digital CoE directly (customer interaction layer).",
                "▪  Develops solutions for the customer: portal, mobile apps, websites.",
                "▪  Interfaces with Contact Center + Call Center (customer interaction).",
                "▪  Runs separately from BRS-driven back-office demand — closer to CX journeys than to SAP modules.",
                "▪  Connects to P&T Process & Technology (field operators) and Renewables CoE patterns already mapped.",
            ],
        },
        {
            "code": "INFRASTRUCTURE",
            "title": "SAP Basis / platform",
            "accent": MID_GRAY,
            "scope": "Nagi Reddy  —  SAP Basis (~2 FTE)",
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

    # Bottom row — Digital CoE + Infrastructure (centered)
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
        "Any TOM recommendation on Customer Services demand must respect three regional cadences (UIL, CMP, NY), one customer-facing track (Digital CoE), and one horizontal platform (Infrastructure) — single model, five consumption patterns",
        color=NAVY)
    add_footer(slide, page_label="SAP Landscape Detail  |  10 / 19")


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
    add_text(slide, Inches(1.1), Inches(3.25), Inches(11.5), Inches(1.3),
             "Demand Map", size=60, bold=True, color=WHITE, font=FONT)

    # Subtitle — the governing thought
    add_text(slide, Inches(1.1), Inches(4.65), Inches(11.5), Inches(0.5),
             "AS-IS Current State of Digital Products Demand Flow",
             size=20, color=BLUE, font=FONT)

    # Governing thought — one-liner that frames the deck
    add_text(slide, Inches(1.1), Inches(5.3), Inches(11.5), Inches(1.3),
             "Across Applications (Carlos) and Customer Services (Kumar), demand enters through\n"
             "6+ disconnected channels — no two teams share the same flow. This map exposes the gaps\n"
             "and defines what the Target Operating Model must standardize end-to-end.",
             size=12, color=WHITE, font=FONT, italic=True)

    # Bottom metadata strip
    add_rect(slide, Inches(1.1), Inches(6.7), Inches(0.3), Inches(0.03), BLUE)
    add_text(slide, Inches(1.1), Inches(6.82), Inches(6), Inches(0.3),
             "NTT DATA  —  Consulting", size=10, bold=True, color=WHITE, font=FONT)

    add_text(slide, Inches(10.5), Inches(6.82), Inches(2.5), Inches(0.3),
             "19 slides", size=9, color=MID_GRAY, font=FONT,
             align=PP_ALIGN.RIGHT)
    add_text(slide, Inches(10.5), Inches(7.12), Inches(2.5), Inches(0.25),
             "Version 1.0", size=8, color=MID_GRAY, font=FONT,
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
             "Vitor Cruz\nVP — IT Business Solutions Delivery",
             size=9, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Two director boxes
    add_rect(slide, Inches(1.5), Inches(2.3), Inches(5.3), Inches(0.55), BLUE)
    add_text(slide, Inches(1.5), Inches(2.3), Inches(5.3), Inches(0.55),
             "Carlos Garbin  —  Director IT Business Applications  (~26 FTE)",
             size=10, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    add_rect(slide, Inches(6.9), Inches(2.3), Inches(5.3), Inches(0.55), NAVY)
    add_text(slide, Inches(6.9), Inches(2.3), Inches(5.3), Inches(0.55),
             "Kumar Jaladanki  —  Director Customer Services Apps  (~24 FTE)",
             size=10, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Applications branch — 4 sub-teams (FTE from validated org chart)
    apps_teams = [
        ("Bruno Moreto", "Electric & Gas Ops", "7 FTE", True, "G&E Champion: Andrew McGibbon"),
        ("George Porto", "GIS & OSG", "6 FTE", True, "GIS Champion: George Porto"),
        ("Sarath Garikapati", "Corporate Applications", "8 FTE", True, "Corporate Champion: Amit Patel"),
        ("Vanessa Colombo", "Global SAP — Iberdrola\nFI, PO, Budget (Spain)", "5 FTE", True, "Global SAP Champion: Michael Dubois"),
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
        ("Ivelisse Velazquez", "FICA, Billing & Collection", "8 FTE", "operational", "Customer billing operations"),
        ("TBD (vacant)", "CRM & Contact Center Ops", "4 FTE", "org_vacancy", "Org vacancy — leadership gap"),
        ("Vincent Lam", "Device & Metering (Avangrid)", "8 FTE", "interviewed", "Smart Metering Champion: Iram Chowdhury"),
        ("Open (vacant)", "SAP & CS Architecture", "2 FTE", "org_vacancy", "Org vacancy — leadership gap"),
        ("Nagi Reddy", "SAP Basis", "2 FTE", "operational", "Platform / infrastructure"),
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
    add_footer(slide, page_label="Org Context  |  07 / 19")


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
             "Applications branch — Director Carlos Garbin",
             size=9, bold=True, color=BLUE, font=FONT)
    apps = "GIS (George)  •  Corporate (Amit)  •  G&E (Andrew/Bruno)  •  Global SAP (Michael/Vanessa)"
    add_text(slide, Inches(6.5), Inches(5.00), Inches(6.5), Inches(0.4),
             apps, size=9, color=CHARCOAL, font=FONT)
    add_text(slide, Inches(6.5), Inches(5.50), Inches(6.5), Inches(0.25),
             "Customer Services branch — Director Kumar Jaladanki",
             size=9, bold=True, color=BLUE, font=FONT)
    cs = "Smart Metering (Iram/Vincent)  •  Branch overview from Kumar + Stephanie"
    add_text(slide, Inches(6.5), Inches(5.75), Inches(6.5), Inches(0.4),
             cs, size=9, color=CHARCOAL, font=FONT)

    add_takeaway(slide,
        "The map is organized as 5 area swimlanes × 5 process stages — every cell is a claim grounded in interview evidence")
    add_footer(slide,
        "Source: Champions Workshop (6 interviewees, Apr 2026); IT CS Interview (Stephanie & Kumar); Applications Interview (Carlos)",
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
             "Carlos Garbin, Director IT Business Applications  •  4 teams  •  ~26 FTE",
             size=8, italic=True, color=BLUE, font=FONT, anchor=MSO_ANCHOR.MIDDLE)
    # Stage headers
    _draw_stage_header(slide, matrix_top, stage_w, header_h)

    # Rows grouped by director branch (APPLICATIONS first, then CUSTOMER SERVICES / LOCAL SAP)
    # Each entry: (area, intake, classification, prioritization, portfolio, delivery, colors, branch)
    areas_data = [
        # ─── APPLICATIONS BRANCH (Carlos Garbin) ─────────────────────
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
        # ─── CUSTOMER SERVICES / LOCAL SAP BRANCH (Kumar Jaladanki) ─────
        ("Smart Metering\n(Vincent)",
         "Email (IT Now)",
         "IT Now template",
         "Criticality + deadline",
         "PPM Now + weekly mtg",
         "Vendor (100%)",
         [YELLOW_GAP, GREEN_OK, YELLOW_GAP, GREEN_OK, RED_CRITICAL],
         "CS"),
        ("Customer Services\n(Kumar — 5 teams)",
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
                     "Kumar Jaladanki, Director Customer Services Apps  •  5 teams  •  ~24 FTE",
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
        page_num="11 / 19",
        area="GIS",
        champion="GIS & OSG — team led by George Porto (Sr. Manager)",
        action_title="GIS balances break/fix urgency against yearly PDD planning — with no formal channel for small demand",
        grid={
            "CHANNELS": [
                "Jira (misc. work) + Email / Meetings (small demand) + IT-Now (daily break/fix)",
                "Small: email / Jira / IT-Now routed by type; Large: yearly PDD process",
                "Business drives prioritization; P&T is the face for small demand",
                "PDD-approved portfolio; no dedicated portfolio tool",
                "Internal team for break/fix; SWF brought in for resource impact",
            ],
            "ROLES": [
                "Business users open tickets; George's team receives",
                "George + team classify; PDD committee for large",
                "\"Impact to most users\" rule (George); PDD approvers for large",
                "Informal PM — George manages via OneNote / Teams / spreadsheets",
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
            "In most cases these are break/fix requests\" — George Porto"
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
        page_num="12 / 19",
        area="Corporate",
        champion="Corporate Applications — team led by Sarath Garikapati (Sr. Manager)  ·  Champion: Amit Patel (Tech Lead, 5-6 PMs + 2 Tech Leads)",
        action_title="Corporate runs demand through Jira, Confluence and Word — standardization gap despite the highest regulatory pressure",
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
            "coming with new requirements that are always a priority\" — Amitkumar Patel"
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
        page_num="13 / 19",
        area="G&E / Electric & Gas Operations",
        champion="G&E — sub-scope within Electric & Gas Operations Products (team led by Bruno Moreto, Sr. Manager)  ·  Champion: Andrew McGibbon (70-90 apps supported)",
        action_title="G&E is the only area with NO formal digital intake — every request arrives via email from business or vendor",
        grid={
            "CHANNELS": [
                "EMAILS only — from business AND from vendor; no ticketing",
                "Ad-hoc by demand nature; depends on upgrade type",
                "Andrew's internal call — operational = urgent",
                "Vendor sets its own timeline; upgrades need budget/time eval",
                "Vendor-led delivery; competing priorities without triage",
            ],
            "ROLES": [
                "Business + vendor send email → Andrew receives",
                "Andrew alone classifies — no second reviewer",
                "Andrew prioritizes; cyber / legal / third-party weighted",
                "Vendor owns timeline; Andrew tracks manually",
                "Vendor + Andrew; no internal release role",
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
                "Key-person risk — Andrew = only router",
                "Third-Party Risk gates upgrades; interface is slow",
                "Budget / time tension on every upgrade; no KPIs",
            ],
        },
        evidence=(
            "\"I receive requests/input from the business or from the vendor\" — "
            "Andrew McGibbon (via email)"
        ),
        takeaway=(
            "G&E is the canary of intake fragmentation — zero digital trail, every decision "
            "is individual and unreproducible; must be the first area addressed in the TOM"
        ),
        maturity=["Gap", "Gap", "Gap", "Gap", "Gap"],
        analysis=[
            "Vendor-led model creates structural dependency: Avangrid cannot estimate or schedule its own work — budget/timeline are exogenous",
            "Third-party risk assessments add hidden lead time — upgrades routinely stall at this interface, invisible in any tracking system",
            "Single-person routing (Andrew) is a key-person risk — zero redundancy, zero documentation, zero auditability",
        ],
    )


def slide_7_global_sap(prs):
    _drill_down(
        prs,
        page_num="14 / 19",
        area="Global SAP — Iberdrola (Spain)",
        champion="Global SAP — Iberdrola environment (Spain) — team led by Vanessa Colombo (Manager IT Apps Global SAP) covering FI, PO, Budget control  ·  Champion: Michael Dubois (35 yrs in business, now IT)",
        action_title="Global SAP — Iberdrola has the most mature channel taxonomy — but is trapped by incident overload",
        grid={
            "CHANNELS": [
                "IT Now (Incidents) · RITM (Minor) · Jira (Small Demands) · SharePoint + Teams (Projects)",
                "Incidents→IT NOW · Small→JIRA · Minor→RITM · Major→SharePoint",
                "#1 Incidents always first; #2 Small + Major; P&T prioritizes projects",
                "PDD for major projects; SharePoint is project home",
                "Accenture + SWF (Software Factory) run delivery",
            ],
            "ROLES": [
                "Business / Accenture open tickets; Vanessa's 5-FTE team receives",
                "Classification is tool-driven — channel = class",
                "Michael + leadership call incidents first; P&T helps with projects",
                "Michael owns weekly incident review (Thursday, with Accenture)",
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
            "Limited resources — 1 year to hire someone\" — Michael Dubois"
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
        page_num="15 / 19",
        area="Smart Metering — Avangrid (US)",
        champion="Smart Metering — sub-scope within Device & Metering (team led by Vincent Lam, Sr. Manager — Customer Services branch)  ·  Champion: Iram Chowdhury (PM, works with Sarath; also Andrew's group)",
        action_title="Smart Metering is the only area fully vendor-delivered with PPM Now + Confluence — but relies on \"intrinsic common sense\" for prioritization",
        grid={
            "CHANNELS": [
                "Small demands via email (IT Now); normal flow — not seasonal",
                "Incidents → IT Now template; requirements usually clear",
                "Criticality + deadline + application priority",
                "PPM Now (large) + Jira + Confluence (PM-to-PM)",
                "100% vendor-delivered portfolio — vendors = consultants + code",
            ],
            "ROLES": [
                "Business opens IT Now / email; Iram (PM) receives",
                "Iram + team classify by template; vendor tech lead joins",
                "Iram prioritizes on deadline + criticality — \"intrinsic common sense\"",
                "Iram + manager (weekly); vendor PMs help PMs prioritize",
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
            "applications\" — Iram Chowdhury"
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
                maturity=None, analysis=None, grid=None):
    """Template for per-area drill-down slides — 4 categories × 5 stages grid.

    grid: dict with 4 keys (ordered): "CHANNELS", "ROLES", "TOOLS", "GOVERNANCE"
          each mapped to a list of 5 strings (one per lifecycle stage).
          When grid is None, falls back to legacy `stages` dict (kept for safety).
    maturity: list of 5 maturity labels (OK / Partial / Gap / N/A) — one per stage.
    analysis: list of 2-3 analytical bullets.
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)

    add_action_title(slide, action_title)

    # Champion strip
    add_text(slide, Inches(0.5), Inches(1.28), Inches(12.33), Inches(0.32),
             champion, size=11, color=BLUE, font=FONT, italic=True)

    maturity_colors = {
        "OK": GREEN_OK,
        "Partial": YELLOW_GAP,
        "Gap": RED_CRITICAL,
        "N/A": MID_GRAY,
    }

    # Layout geometry
    # Left label strip (rotated/category label) width + 5 stage columns
    label_w = 0.65
    col_x0 = 0.5 + label_w + 0.05  # x where stage columns start
    col_w = 2.37
    col_gap = 0.06

    header_top = 1.72
    header_h = 0.42
    pill_h = 0.22
    pill_gap = 0.08

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
    grid_top = header_top + header_h + pill_h + pill_gap + 0.12
    row_h = 0.75
    categories = [
        ("CHANNELS  ·  WHO SENDS", GREEN_OK),
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
    an_top = grid_bottom + 0.10
    an_h = 6.55 - an_top  # stop before takeaway
    if an_h < 0.9:
        an_h = 0.9
    add_rect(slide, Inches(0.5), Inches(an_top), Inches(12.33), Inches(an_h), LIGHT_GRAY)
    add_rect(slide, Inches(0.5), Inches(an_top), Inches(0.08), Inches(an_h), NAVY)
    add_text(slide, Inches(0.70), Inches(an_top + 0.05), Inches(6), Inches(0.22),
             "KEY PATTERNS  —  WHAT THE FINDINGS REVEAL",
             size=9, bold=True, color=NAVY, font=FONT)
    add_text(slide, Inches(7.0), Inches(an_top + 0.05), Inches(5.65), Inches(0.22),
             "SUPPORTING QUOTE", size=9, bold=True, color=ACCENT, font=FONT)
    add_text(slide, Inches(7.0), Inches(an_top + 0.32), Inches(5.65), Inches(an_h - 0.40),
             evidence, size=10, color=CHARCOAL, font=FONT, italic=True)
    if analysis:
        analysis_text = "\n".join(["▸  " + a for a in analysis])
        add_text(slide, Inches(0.70), Inches(an_top + 0.32), Inches(6.0), Inches(an_h - 0.40),
                 analysis_text, size=10, color=CHARCOAL, font=FONT)

    add_takeaway(slide, takeaway, color=ACCENT)
    add_footer(slide, page_label=f"Drill-down: {area}  |  {page_num}")


def slide_10_customer_services(prs):
    """Drill-down: Customer Services Apps branch (Kumar Jaladanki)."""
    _drill_down(
        prs,
        page_num="16 / 19",
        area="Customer Services Apps (Kumar's Branch)",
        champion="Customer Services Apps — Kumar Jaladanki (Director) — 5 sub-teams: FICA/Billing (Ivelisse), CRM & Contact Center (vacant), Device & Metering (Vincent), SAP & CS Architecture (vacant), SAP Basis (Nagi)",
        action_title="Customer Services runs 5 sub-teams with heterogeneous maturity — only Smart Metering exhibits a mature portfolio practice; the rest route demand director-level",
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
                "Kumar + sub-team managers; \"intrinsic common sense\" (Iram)",
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
            "Resource allocation — multiple OPCOs competing by capacity\" — Kumar + Stephanie"
        ),
        takeaway=(
            "Smart Metering's model (PPM Now + weekly PM meeting + 100% vendor delivery) is the "
            "only mature practice in Kumar's branch — it is also the blueprint and the warning sign"
        ),
        maturity=["Gap", "Partial", "Gap", "Gap", "Partial"],
        analysis=[
            "Three parallel ownership tracks without a routing rule — Travis PMO for large projects, BRS for business-funded work, IT-funded for Kumar's teams — the same request can land in any of them, blurring accountability and prioritization",
            "Smart Metering is the mature outlier; other four teams operate without equivalent tooling — branch has no portfolio consolidation",
            "PM + SAP Architect roles are in place for large projects (Kumar confirmed) — the gap is not in roles, it's in branch-wide portfolio visibility",
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
            "stage": "Bruno Moreto's team  •  STAGE: Intake",
            "title": "G&E Email-Only Intake",
            "desc": "Every request to Andrew arrives via email from business or vendor. No ticket, no queue, no SLA — each decision is individual and unreproducible.",
            "impact": "Unknown volume. No capacity forecast possible. Single-person dependency on Andrew.",
            "evidence": "\"I receive requests/input from the business or from the vendor\" — Andrew McGibbon",
        },
        {
            "num": "02",
            "stage": "Vanessa Colombo's team  •  STAGE: Prioritization",
            "title": "Global SAP Incident Flood",
            "desc": "Incidents outrank all other demand for Iberdrola SAP. Entire team can be consumed by incident work, blocking strategic projects.",
            "impact": "Capacity trapped in BAU. Major projects (PDD) slip. Knowledge transfer lost at each turnover event.",
            "evidence": "\"Sometimes entire team is focusing on incidents\" — Michael Dubois",
        },
        {
            "num": "03",
            "stage": "George Porto's team  •  STAGES: Intake + Portfolio",
            "title": "GIS Informal PM Stack",
            "desc": "Small demand routes through email, meetings and IT-Now. Project management relies on OneNote + Teams + spreadsheets — no formal portfolio tool below PDD.",
            "impact": "Small + break/fix demand (~60-70% of actual work) is untracked. Capacity consumed by unmeasured flow.",
            "evidence": "\"My team uses Jira for misc. work... OneNote, Teams and Spreadsheets to manage\" — George Porto",
        },
        {
            "num": "04",
            "stage": "Sarath Garikapati's team  •  STAGE: Classification",
            "title": "Corporate Project Overlap",
            "desc": "Word-document requirements + Jira + Confluence fragment the intake. PDD drives classification but doesn't deduplicate — overlapping projects chase similar outcomes.",
            "impact": "Rework, duplicated effort, and unreconciled regulatory pressure treated as chronic urgency.",
            "evidence": "\"PDD is the driver. Lots of overlapping projects\" — Amit Patel",
        },
    ]
    _bottlenecks_slide(
        prs,
        branch="APPLICATIONS",
        director="Carlos Garbin",
        page_num="17 / 19",
        section_label="Applications Bottlenecks",
        action_title="Four Applications bottlenecks concentrate demand risk — covering intake, prioritization, portfolio and classification",
        takeaway="These four bottlenecks span all four Applications teams — each is a symptom of the same underlying fragmentation the TOM must address",
        cards=cards,
    )


def slide_11_cs_bottlenecks(prs):
    cards = [
        {
            "num": "05",
            "stage": "Kumar's branch  •  STAGES: Classification + Prioritization",
            "title": "Customer Services Leadership Gap",
            "desc": "Two Sr. Manager positions vacant (CRM & Contact Center Ops, SAP & CS Architecture). Prioritization defaults to director level across Billing, CRM, Metering and Basis.",
            "impact": "No shared prioritization logic across 5 sub-teams. Branch-level demand view is a structural blind spot.",
            "evidence": "\"Who is in charge of what — 3 people can run the payment jobs\" — Kumar + Stephanie",
        },
        {
            "num": "06",
            "stage": "Kumar's branch  •  STAGES: Portfolio + Delivery",
            "title": "No Branch-Wide Portfolio Tool",
            "desc": "PPM Now exists only in Smart Metering. Travis PMO handles large projects without a consolidated view of Billing / CRM / Metering / Basis portfolios.",
            "impact": "Sub-team portfolios are disconnected. Capacity + priorities cannot be reconciled across the branch.",
            "evidence": "\"No tool to support portfolio management... No tools or process to manage delivery of vendors\"",
        },
        {
            "num": "07",
            "stage": "Vincent Lam's team  •  STAGE: Delivery",
            "title": "Smart Metering Vendor Concentration",
            "desc": "100% of the Smart Metering portfolio is vendor-delivered. Vendors do requirements AND code. All custom applications. \"Intrinsic common sense\" is the prioritization rule.",
            "impact": "If the vendor exits, execution knowledge is lost. Avangrid has no native build capability in this sub-team.",
            "evidence": "\"All portfolio in vendor. Vendors are consultants for the work\" — Iram Chowdhury",
        },
        {
            "num": "08",
            "stage": "Kumar's branch  •  STAGE: Intake + Prioritization",
            "title": "Multi-OPCO Capacity Competition",
            "desc": "7 business OPCOs compete for the same Customer Services IT capacity. Demand splits across Travis PMO / BRS / IT-funded — triple routing for the same request.",
            "impact": "Ad-hoc capacity allocation. Ownership ambiguity creates delivery risk and prolongs leadership vacancies' impact.",
            "evidence": "\"Resource allocation — multiple OPCOs competing by capacity\" — Kumar + Stephanie",
        },
    ]
    _bottlenecks_slide(
        prs,
        branch="CUSTOMER SERVICES",
        director="Kumar Jaladanki",
        page_num="18 / 19",
        section_label="Customer Services Bottlenecks",
        action_title="Four Customer Services bottlenecks — leadership gaps, portfolio blindness, vendor concentration and multi-OPCO contention",
        takeaway="Kumar's branch carries both structural (leadership) and operational (tooling + vendor) bottlenecks — the TOM must address all four simultaneously",
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
            "evidence": "\"I receive requests/input from the business or from the vendor\" — Andrew McGibbon",
            "color": RED_CRITICAL,
        },
        {
            "num": "02",
            "title": "Global SAP Incident Flood",
            "stage": "BRANCH: Applications  /  STAGE: Prioritization",
            "desc": "Incidents outrank all other demand. Entire team can be consumed by incident work, blocking strategic projects.",
            "impact": "Capacity trapped in BAU. Major projects (PDD) slip. Turnover compounds loss.",
            "evidence": "\"Sometimes entire team is focusing on incidents\" — Michael Dubois",
            "color": RED_CRITICAL,
        },
        {
            "num": "03",
            "title": "Customer Services Leadership Gap",
            "stage": "BRANCH: Customer Services  /  STAGE: Classification + Prioritization",
            "desc": "Two Sr. Manager positions vacant (CRM & Contact Center Ops, SAP & CS Architecture). Prioritization happens director-level only across Billing, CRM, Metering and Basis — without shared criteria.",
            "impact": "No consolidated branch-level view of demand. Travis PMO absorbs large projects without area-level coordination. Leadership vacancies compound hiring pressure already observed in Global SAP.",
            "evidence": "\"Who is in charge of what — 3 people can run the payment jobs\" — Kumar + Stephanie",
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
             "BRANCH 1  —  APPLICATIONS  (Carlos Garbin)",
             size=10, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # BRANCH BAND 2 — Customer Services + cross-branch
    add_rect(slide, Inches(0.5), Inches(y_band2), Inches(12.77), Inches(band_h), NAVY)
    add_text(slide, Inches(0.5), Inches(y_band2), Inches(12.77), Inches(band_h),
             "BRANCH 2  —  CUSTOMER SERVICES  (Kumar Jaladanki)   +   CROSS-BRANCH (both directors)",
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
        "Structural Insights  |  19 / 19")


# ========== MAIN ==========

def build():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    # ── Cover + Deck Roadmap (parent slide) ─────────────────────────
    slide_1_cover(prs)                                # 01 — no nav strip (cover stands alone)
    slide_roadmap(prs)                                # 02 — parent / roadmap slide (shows the whole journey)

    # ── Section 01 · FRAMING ────────────────────────────────────────
    slide_context(prs);             _add_nav_to_last(prs, 0)     # 03
    slide_executive_summary(prs);   _add_nav_to_last(prs, 0)     # 04
    slide_methodology(prs);         _add_nav_to_last(prs, 0)     # 05
    slide_2_how_to_read(prs);       _add_nav_to_last(prs, 0)     # 06

    # ── Section 02 · ORGANIZATION ───────────────────────────────────
    slide_org_context(prs);             _add_nav_to_last(prs, 1) # 07
    slide_apps_branch_overview(prs);    _add_nav_to_last(prs, 1) # 08
    slide_sap_branch_overview(prs);     _add_nav_to_last(prs, 1) # 09
    slide_sap_landscape_detail(prs);    _add_nav_to_last(prs, 1) # 10

    # ── Section 03 · AS-IS DRILL-DOWNS ──────────────────────────────
    slide_4_gis(prs);                   _add_nav_to_last(prs, 2) # 11
    slide_5_corporate(prs);             _add_nav_to_last(prs, 2) # 12
    slide_6_ge(prs);                    _add_nav_to_last(prs, 2) # 13
    slide_7_global_sap(prs);            _add_nav_to_last(prs, 2) # 14
    slide_8_smart_metering(prs);        _add_nav_to_last(prs, 2) # 15
    slide_10_customer_services(prs);    _add_nav_to_last(prs, 2) # 16

    # ── Section 04 · BOTTLENECKS ────────────────────────────────────
    slide_10_apps_bottlenecks(prs);     _add_nav_to_last(prs, 3) # 17
    slide_11_cs_bottlenecks(prs);       _add_nav_to_last(prs, 3) # 18

    # ── Section 05 · INSIGHTS ───────────────────────────────────────
    slide_11_insights(prs);             _add_nav_to_last(prs, 4) # 19

    out = "/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Deliverables/Demand_Map/Demand_Map_v1.pptx"
    prs.save(out)
    print(f"Saved: {out}")
    print(f"Slides: {len(prs.slides)}")


if __name__ == "__main__":
    build()
