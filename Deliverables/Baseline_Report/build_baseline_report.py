"""
Baseline Report — Avangrid SAP Digital Products
NTT DATA engagement | Current State Diagnosis + TOM Foundations

Grounded in Knowledge Base:
- Interview_Application_Champions.md (GIS, Corporate/Amit, G&E, Global SAP, Smart Metering)
- Interview_Stephanie_Kumar_IT_CS.md
- Interview_Carlos_Applications.md
- Org_Chart_Digital_Products.md
- SAP CoE Digital products data.xlsx
- SAP CoE New Ways of Working v2.0.pptx (Slides 3, 7, 12, 21)

Reuses visual language + components from Demand Map v1.2
Output: 18 slides
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ========== NTT BRAND TOKENS ==========
NAVY = RGBColor(0x0F, 0x1C, 0x50)
BLUE = RGBColor(0x67, 0x85, 0xC1)
ACCENT = RGBColor(0xBC, 0x43, 0x28)
HIGHLIGHT = RGBColor(0xE6, 0xB6, 0x00)

WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY = RGBColor(0xF5, 0xF5, 0xF5)
MID_GRAY = RGBColor(0x95, 0xA5, 0xA6)
CHARCOAL = RGBColor(0x2C, 0x3E, 0x50)
BLACK = RGBColor(0x00, 0x00, 0x00)

GREEN_OK = RGBColor(0x2E, 0x7D, 0x32)
YELLOW_GAP = RGBColor(0xF9, 0xA8, 0x25)
RED_CRITICAL = RGBColor(0xBC, 0x43, 0x28)

FONT = "Helvetica Neue"

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
    add_text(slide, Inches(0.5), Inches(0.35), Inches(12.33), Inches(0.9),
             title_text, size=20, bold=True, color=NAVY, font=FONT)
    add_rect(slide, Inches(0.5), Inches(1.25), Inches(12.33), Inches(0.015),
             LIGHT_GRAY)


def add_takeaway(slide, text, color=NAVY):
    add_rect(slide, Inches(0.5), Inches(6.55), Inches(12.33), Inches(0.55), color)
    add_text(slide, Inches(0.75), Inches(6.60), Inches(11.83), Inches(0.45),
             text, size=11, bold=False, color=WHITE, font=FONT,
             anchor=MSO_ANCHOR.MIDDLE)


def add_footer(slide, citation_text=None, page_label=None):
    """Page label only — sources omitted for cleaner executive layout."""
    if page_label:
        add_text(slide, Inches(10.0), Inches(7.18), Inches(2.83), Inches(0.25),
                 page_label, size=9, color=MID_GRAY, font=FONT, align=PP_ALIGN.RIGHT)


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


# ========== SLIDES ==========

def slide_1_cover(prs):
    slide = blank(prs)
    set_bg(slide, NAVY)

    add_rect(slide, Inches(0.75), Inches(0.75), Inches(0.04), Inches(6), BLUE)

    add_text(slide, Inches(1.1), Inches(0.85), Inches(11.5), Inches(0.3),
             "AVANGRID  —  SAP DIGITAL PRODUCTS TRANSFORMATION",
             size=10, bold=True, color=BLUE, font=FONT)
    add_text(slide, Inches(1.1), Inches(1.2), Inches(11.5), Inches(0.3),
             "Baseline Phase — Current State Diagnosis  •  April 2026",
             size=10, color=MID_GRAY, font=FONT)

    add_rect(slide, Inches(1.1), Inches(3.1), Inches(0.5), Inches(0.05), ACCENT)
    add_text(slide, Inches(1.1), Inches(3.25), Inches(11.5), Inches(1.3),
             "Baseline Report", size=60, bold=True, color=WHITE, font=FONT)

    add_text(slide, Inches(1.1), Inches(4.65), Inches(11.5), Inches(0.5),
             "Current State Diagnosis of Digital Products Operating Model",
             size=20, color=BLUE, font=FONT)

    add_text(slide, Inches(1.1), Inches(5.3), Inches(11.5), Inches(1.3),
             "Digital Products operates as two disconnected branches with seven structural issues\n"
             "no single-area fix can resolve — the Target Operating Model must standardize intake,\n"
             "prioritization, tooling, governance, taxonomy, vendor management and value-stream lens.",
             size=12, color=WHITE, font=FONT, italic=True)

    add_rect(slide, Inches(1.1), Inches(6.7), Inches(0.3), Inches(0.03), BLUE)
    add_text(slide, Inches(1.1), Inches(6.82), Inches(6), Inches(0.3),
             "NTT DATA  —  Consulting", size=10, bold=True, color=WHITE, font=FONT)

    add_text(slide, Inches(10.5), Inches(6.82), Inches(2.5), Inches(0.3),
             "15 slides", size=9, color=MID_GRAY, font=FONT,
             align=PP_ALIGN.RIGHT)
    add_text(slide, Inches(10.5), Inches(7.12), Inches(2.5), Inches(0.25),
             "Version 1.0", size=8, color=MID_GRAY, font=FONT,
             align=PP_ALIGN.RIGHT, italic=True)


def slide_2_executive_summary(prs):
    slide = blank(prs)
    set_bg(slide, WHITE)
    add_action_title(slide,
        "The Digital Products operating model is structurally fragmented — three findings set the mandate, three implications define the TOM")

    # Left column: 3 findings
    add_text(slide, Inches(0.5), Inches(1.50), Inches(6.1), Inches(0.35),
             "WHAT WE FOUND", size=11, bold=True, color=ACCENT, font=FONT)

    findings = [
        ("01",
         "Two branches, nine sub-teams, ~50 FTE — zero standardization",
         "Demand flows through 6+ disconnected channels. Applications (Carlos) and Customer Services (Kumar) run different intake, tools, and logic. No two teams share the same flow."),
        ("02",
         "Prioritization is ad-hoc across both branches",
         "No shared scoring model. Four informal criteria (strategic, operational urgency, regulatory, size) applied per forum per team. Global SAP drowns in incidents; G&E runs on email; Smart Metering relies on \"intrinsic common sense\"."),
        ("03",
         "Portfolio and vendor governance are largely absent",
         "Only 1 of 6 units uses a portfolio tool (Smart Metering / PPM Now). No process governs vendor delivery — vendors estimate, set timelines, and own full scope in Smart Metering."),
    ]
    y = 1.90
    for num, title, desc in findings:
        add_rect(slide, Inches(0.5), Inches(y), Inches(0.5), Inches(1.35), ACCENT)
        add_text(slide, Inches(0.5), Inches(y + 0.15), Inches(0.5), Inches(0.4),
                 num, size=16, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER)
        add_text(slide, Inches(1.10), Inches(y), Inches(5.5), Inches(0.35),
                 title, size=11, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(1.10), Inches(y + 0.35), Inches(5.5), Inches(1.0),
                 desc, size=9, color=CHARCOAL, font=FONT)
        y += 1.45

    # Right column: 3 implications
    add_text(slide, Inches(6.85), Inches(1.50), Inches(6.0), Inches(0.35),
             "WHAT IT MEANS FOR THE TOM", size=11, bold=True, color=NAVY, font=FONT)

    implications = [
        ("A",
         "Unified intake layer + value-based prioritization framework",
         "A single structured front door replacing the 6+ channels. A scoring model that formalizes the four criteria and applies objectively across both branches."),
        ("B",
         "Portfolio + vendor tooling scaled across Applications and Customer Services",
         "Extend PPM Now (or equivalent) beyond Smart Metering. Introduce vendor performance framework with independent estimation sanity checks."),
        ("C",
         "Taxonomy + value-stream lens — move from Project to Product operating model",
         "Objective small/medium/large thresholds. Demand organized around customer value streams, not application silos — traceable from business OKRs."),
    ]
    y = 1.90
    for num, title, desc in implications:
        add_rect(slide, Inches(6.85), Inches(y), Inches(0.5), Inches(1.35), NAVY)
        add_text(slide, Inches(6.85), Inches(y + 0.15), Inches(0.5), Inches(0.4),
                 num, size=16, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER)
        add_text(slide, Inches(7.45), Inches(y), Inches(5.4), Inches(0.35),
                 title, size=11, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(7.45), Inches(y + 0.35), Inches(5.4), Inches(1.0),
                 desc, size=9, color=CHARCOAL, font=FONT)
        y += 1.45

    add_takeaway(slide,
        "The question is no longer whether to transform — it is how fast the TOM can replace fragmentation with an end-to-end, branch-unified operating model",
        color=NAVY)
    add_footer(slide,
        "Source: KB synthesis across 4 interview MDs + Org_Chart + Excel + SAP CoE New Ways of Working v2.0",
        "Executive Summary  |  02 / 15")


def slide_3_context(prs):
    slide = blank(prs)
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Growth, regulatory pressure and two newly-appointed directors create the window — transformation cannot wait")

    # Three pillars
    pillars = [
        {
            "icon": "GROWTH",
            "title": "30–50% Projected Growth",
            "body": "Digital Products faces 30–50% growth pressure across its portfolio. The current operating model — siloed Jira instances, ad-hoc email intake, fragmented tooling — cannot scale with this volume. Either the TOM ships, or fragmentation compounds.",
            "source": "Source: SAP CoE New Ways of Working v2.0 (Slide 3)",
        },
        {
            "icon": "MANDATE",
            "title": "New Director Structure",
            "body": "CIO Edu de Frutos and VP Vitor Cruz have split Digital Products into two directorates — Carlos Garbin (Applications, ~26 FTE) and Kumar Jaladanki (Customer Services, ~24 FTE). The mandate to standardize is explicit: \"Attention here — the main opportunity is here to improve and standardize\" (Carlos's box, working session).",
            "source": "Source: Validated Miro org chart + Interview_Carlos_Applications.md",
        },
        {
            "icon": "PRESSURE",
            "title": "Regulatory + Capacity Tension",
            "body": "Regulatory demand overrides size in every prioritization forum. Global SAP teams report \"entire team focused on incidents\" and \"1 year to hire someone\". G&E cites third-party risk + cyber legal as top priorities. The organization is running its current model at maximum load.",
            "source": "Source: Interview_Application_Champions.md (Michael Dubois, Andrew McGibbon)",
        },
    ]

    card_w = 4.1
    card_h = 4.5
    gap = 0.1
    x0 = 0.5
    y0 = 1.55

    for i, p in enumerate(pillars):
        x = x0 + i * (card_w + gap)
        add_rect(slide, Inches(x), Inches(y0), Inches(card_w), Inches(card_h), LIGHT_GRAY)
        add_rect(slide, Inches(x), Inches(y0), Inches(card_w), Inches(0.08), ACCENT)
        add_text(slide, Inches(x + 0.25), Inches(y0 + 0.25), Inches(card_w - 0.5), Inches(0.3),
                 p["icon"], size=9, bold=True, color=ACCENT, font=FONT)
        add_text(slide, Inches(x + 0.25), Inches(y0 + 0.55), Inches(card_w - 0.5), Inches(0.6),
                 p["title"], size=14, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(x + 0.25), Inches(y0 + 1.25), Inches(card_w - 0.5), Inches(2.8),
                 p["body"], size=10, color=CHARCOAL, font=FONT)
        add_text(slide, Inches(x + 0.25), Inches(y0 + card_h - 0.4), Inches(card_w - 0.5), Inches(0.3),
                 p["source"], size=7, color=MID_GRAY, font=FONT, italic=True)

    add_takeaway(slide,
        "The confluence of growth pressure, leadership consolidation, and capacity saturation makes this the structural reset moment — not an incremental improvement cycle",
        color=NAVY)
    add_footer(slide,
        "Source: KB — SAP CoE New Ways of Working v2.0 (Slide 3); Interview files; Validated Miro org chart",
        "Context & Why Now  |  03 / 15")


def slide_4_methodology(prs):
    slide = blank(prs)
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Findings are grounded in multiple stakeholder inputs, a validated org chart and the Avangrid Excel baseline — every claim traces to a source")

    # Left: inputs
    add_text(slide, Inches(0.5), Inches(1.50), Inches(5.8), Inches(0.35),
             "INPUTS GROUNDING THIS REPORT", size=10, bold=True, color=NAVY, font=FONT)

    inputs = [
        ("Champion inputs by area", "5 champions representing GIS (George Porto), Corporate (Amit Patel), G&E (Andrew McGibbon), Global SAP (Michael Dubois) and Smart Metering (Iram Chowdhury) — scope and pain points per area."),
        ("Director-level inputs", "Carlos Garbin (Applications) — expectations and structural opportunity. Kumar Jaladanki + Stephanie (Customer Services + business-side) — Travis PMO / BRS / IT-funded split, governance forums."),
        ("Validated Miro org chart", "Signed-off organizational structure — 2 directors, 9 sub-teams, ~50 FTE. Single source of truth for Digital Products team composition and reporting lines."),
        ("Avangrid Excel baseline", "AS-IS Demand × Portfolio canvas per team + TO-BE template. Captures prioritization criteria (strategic / operational urgency / regulatory / size)."),
    ]
    y = 1.85
    for title, desc in inputs:
        add_rect(slide, Inches(0.5), Inches(y), Inches(0.08), Inches(0.9), BLUE)
        add_text(slide, Inches(0.70), Inches(y), Inches(5.7), Inches(0.3),
                 title, size=10, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(0.70), Inches(y + 0.25), Inches(5.7), Inches(0.7),
                 desc, size=9, color=CHARCOAL, font=FONT)
        y += 1.00

    # Right: MECE approach
    add_text(slide, Inches(6.85), Inches(1.50), Inches(6.0), Inches(0.35),
             "MECE DISCIPLINE", size=10, bold=True, color=NAVY, font=FONT)

    principles = [
        ("No claim without source", "Every finding traces to a specific source document, KB section, or validated artifact. Unmeasured metrics flagged as open rather than fabricated."),
        ("MECE decomposition", "The 4-bottleneck map is Mutually Exclusive. The 7 structural issues are Collectively Exhaustive across demand lifecycle dimensions."),
        ("Separation of facts and interpretation", "Quotes are verbatim. Action titles state conclusions. Evidence blocks sit inside drill-down slides — reader can re-litigate any claim."),
        ("Coverage transparency", "5 of 9 sub-teams have a designated area champion. 2 sub-teams pending input + 2 org vacancies flagged on the org chart."),
    ]
    y = 1.85
    for title, desc in principles:
        add_rect(slide, Inches(6.85), Inches(y), Inches(0.08), Inches(0.85), NAVY)
        add_text(slide, Inches(7.05), Inches(y), Inches(5.8), Inches(0.3),
                 title, size=10, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(7.05), Inches(y + 0.25), Inches(5.8), Inches(0.6),
                 desc, size=9, color=CHARCOAL, font=FONT)
        y += 1.00

    add_takeaway(slide,
        "The Baseline is a source-grounded synthesis, not a rollup of opinions — findings are testable and traceable to specific artifacts",
        color=NAVY)
    add_footer(slide,
        "KB path: /Materiais de referencia/Interview Findings/ + Org_Chart_Digital_Products.md + Excel baseline",
        "Methodology  |  04 / 15")


def slide_5_org_landscape(prs):
    slide = blank(prs)
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Digital Products spans two director branches (~50 FTE) across nine sub-teams — with two Sr. Manager positions currently vacant")

    # VP at top
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

    add_rect(slide, Inches(6.9), Inches(2.3), Inches(5.3), Inches(0.55), BLUE)
    add_text(slide, Inches(6.9), Inches(2.3), Inches(5.3), Inches(0.55),
             "Kumar Jaladanki  —  Director Customer Services Apps  (~24 FTE)",
             size=10, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Applications branch
    apps_teams = [
        ("Bruno Moreto", "Electric & Gas Ops", "7 FTE", "interviewed", "G&E Champion: Andrew McGibbon"),
        ("George Porto", "GIS & OSG", "6 FTE", "interviewed", "GIS Champion: George Porto"),
        ("Sarath Garikapati", "Corporate Applications", "8 FTE", "interviewed", "Corporate Champion: Amit Patel"),
        ("Vanessa Colombo", "Global SAP — Iberdrola (Spain)", "5 FTE", "interviewed", "Global SAP Champion: Michael Dubois"),
    ]
    status_colors = {
        "interviewed": GREEN_OK,
        "operational": BLUE,
        "org_vacancy": RED_CRITICAL,
    }
    y = 3.05
    for name, scope, fte, status, champ in apps_teams:
        color = status_colors[status]
        add_rect(slide, Inches(1.5), Inches(y), Inches(0.12), Inches(0.55), color)
        add_rect(slide, Inches(1.65), Inches(y), Inches(5.15), Inches(0.55), LIGHT_GRAY)
        add_text(slide, Inches(1.75), Inches(y + 0.02), Inches(2.9), Inches(0.3),
                 name, size=9, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(1.75), Inches(y + 0.28), Inches(2.9), Inches(0.28),
                 scope, size=8, color=CHARCOAL, font=FONT)
        add_rect(slide, Inches(4.20), Inches(y + 0.12), Inches(0.6), Inches(0.30), BLUE,
                 corner=True)
        add_text(slide, Inches(4.20), Inches(y + 0.12), Inches(0.6), Inches(0.30),
                 fte, size=8, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, Inches(4.95), Inches(y + 0.02), Inches(1.8), Inches(0.55),
                 champ, size=8, color=NAVY, font=FONT, italic=True,
                 anchor=MSO_ANCHOR.MIDDLE)
        y += 0.62

    # Customer Services branch
    cs_teams = [
        ("Ivelisse Velazquez", "FICA, Billing & Collection", "8 FTE", "operational", "Customer billing operations"),
        ("TBD (vacant)", "CRM & Contact Center Ops", "4 FTE", "org_vacancy", "Org vacancy"),
        ("Vincent Lam", "Device & Metering (Avangrid)", "8 FTE", "interviewed", "Smart Metering Champion: Iram Chowdhury"),
        ("Open (vacant)", "SAP & CS Architecture", "2 FTE", "org_vacancy", "Org vacancy"),
        ("Nagi Reddy", "SAP Basis", "2 FTE", "operational", "Platform / infrastructure"),
    ]
    y = 3.05
    for name, scope, fte, status, label in cs_teams:
        color = status_colors[status]
        add_rect(slide, Inches(6.9), Inches(y), Inches(0.12), Inches(0.5), color)
        add_rect(slide, Inches(7.03), Inches(y), Inches(5.15), Inches(0.5), LIGHT_GRAY)
        add_text(slide, Inches(7.13), Inches(y + 0.02), Inches(2.5), Inches(0.25),
                 name, size=9, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(7.13), Inches(y + 0.25), Inches(2.5), Inches(0.25),
                 scope, size=7, color=CHARCOAL, font=FONT)
        add_rect(slide, Inches(9.55), Inches(y + 0.10), Inches(0.55), Inches(0.28), BLUE,
                 corner=True)
        add_text(slide, Inches(9.55), Inches(y + 0.10), Inches(0.55), Inches(0.28),
                 fte, size=7, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        prefix = "Scope: " if status == "operational" else ""
        add_text(slide, Inches(10.20), Inches(y + 0.02), Inches(1.95), Inches(0.5),
                 prefix + label, size=7, color=color, font=FONT, italic=True,
                 anchor=MSO_ANCHOR.MIDDLE)
        y += 0.56

    # Legend — 3 status categories with breathing room
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
        "Two Sr. Manager vacancies (CRM & Contact Center Ops, SAP & CS Architecture) leave Customer Services running at 60% leadership coverage — a structural risk the TOM must address",
        color=NAVY)
    add_footer(slide,
        "Source: Validated Miro org chart (Apr 2026); Interview_Carlos_Applications.md; Org_Chart_Digital_Products.md",
        "Org Landscape  |  05 / 15")


def slide_6_master_map(prs):
    slide = blank(prs)
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Across both director branches, Digital Products runs 6+ disconnected intake channels — no two teams share the same demand flow")

    # Applications label
    matrix_top = 1.55
    area_col_w = 1.7
    stage_w = 2.13
    header_h = 0.55
    row_h = 0.68

    add_text(slide, Inches(0.5), Inches(matrix_top - 0.25), Inches(2.5), Inches(0.2),
             "APPLICATIONS — Carlos Garbin", size=8, bold=True, color=BLUE, font=FONT)

    # Stage headers
    stages = [
        ("01 INTAKE", "Channels & Sources"),
        ("02 CLASSIFICATION", "Type, Size, Urgency"),
        ("03 PRIORITIZATION", "Scoring & Criteria"),
        ("04 PORTFOLIO", "Governance & Decision"),
        ("05 DELIVERY", "Execution & Flow"),
    ]
    for i, (num_name, sub) in enumerate(stages):
        x = Inches(2.2 + i * stage_w)
        add_rect(slide, x, Inches(matrix_top), Inches(stage_w - 0.05), Inches(header_h), NAVY)
        add_text(slide, x, Inches(matrix_top + 0.03), Inches(stage_w - 0.05), Inches(0.27),
                 num_name, size=9, bold=True, color=WHITE, font=FONT, align=PP_ALIGN.CENTER)
        add_text(slide, x, Inches(matrix_top + 0.28), Inches(stage_w - 0.05), Inches(0.24),
                 sub, size=7, color=WHITE, font=FONT, align=PP_ALIGN.CENTER)
        if i < len(stages) - 1:
            arrow_x = 2.2 + (i + 1) * stage_w - 0.06
            arrow = slide.shapes.add_shape(
                MSO_SHAPE.RIGHT_ARROW,
                Inches(arrow_x), Inches(matrix_top + header_h / 2 - 0.06),
                Inches(0.15), Inches(0.12))
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = ACCENT
            arrow.line.fill.background()
            arrow.shadow.inherit = False

    areas_data = [
        ("GIS", "Jira + Email + Mtgs + IT-Now", "Small vs. PDD (yearly)", "Business drives; P&T face", "PDD approval", "Internal team + SWF",
         [YELLOW_GAP, YELLOW_GAP, YELLOW_GAP, GREEN_OK, GREEN_OK], "APPS"),
        ("Corporate", "Jira + Confluence + Word", "PDD driver; overlap", "PDD driver", "PDD approval", "Internal + Vendor",
         [YELLOW_GAP, RED_CRITICAL, YELLOW_GAP, GREEN_OK, YELLOW_GAP], "APPS"),
        ("G&E", "Email only (biz/vendor)", "Depends on demand nature", "Operational = urgent", "Vendor sets timeline", "Vendor",
         [RED_CRITICAL, RED_CRITICAL, RED_CRITICAL, RED_CRITICAL, RED_CRITICAL], "APPS"),
        ("Global SAP\n(Iberdrola)", "IT Now + RITM + Jira + SP", "By type (4 channels)", "Incidents first (P&T)", "PDD / SharePoint", "Accenture + SWF",
         [GREEN_OK, GREEN_OK, RED_CRITICAL, YELLOW_GAP, YELLOW_GAP], "APPS"),
        ("Smart Metering\n(Vincent)", "Email (IT Now)", "IT Now template", "Criticality + deadline", "PPM Now + weekly mtg", "Vendor (100%)",
         [YELLOW_GAP, GREEN_OK, YELLOW_GAP, GREEN_OK, RED_CRITICAL], "CS"),
        ("Customer Services\n(Kumar, 4 teams TBC)", "Director-level only; 4 teams TBC", "Travis PMO / BRS / IT-funded split", "\"Intrinsic common sense\"; forums", "No branch tool (Metering = PPM Now only)", "SWF + Accenture + Vendors",
         [RED_CRITICAL, YELLOW_GAP, RED_CRITICAL, RED_CRITICAL, YELLOW_GAP], "CS"),
    ]

    y = matrix_top + header_h + 0.05
    prev_branch = None
    for area, intake, classif, prior, portf, deliv, colors, branch in areas_data:
        if prev_branch == "APPS" and branch == "CS":
            add_rect(slide, Inches(0.5), Inches(y), Inches(12.33), Inches(0.04), ACCENT)
            add_text(slide, Inches(0.5), Inches(y + 0.06), Inches(6.5), Inches(0.2),
                     "CUSTOMER SERVICES / LOCAL SAP — Kumar Jaladanki",
                     size=8, bold=True, color=NAVY, font=FONT)
            y += 0.30

        label_color = NAVY if branch == "CS" else BLUE
        add_rect(slide, Inches(0.5), Inches(y), Inches(1.65), Inches(row_h - 0.08), label_color)
        add_text(slide, Inches(0.5), Inches(y), Inches(1.65), Inches(row_h - 0.08),
                 area, size=10, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        cells = [intake, classif, prior, portf, deliv]
        for i, (text, color) in enumerate(zip(cells, colors)):
            x = Inches(2.2 + i * stage_w)
            add_rect(slide, x, Inches(y), Inches(stage_w - 0.05), Inches(row_h - 0.08), LIGHT_GRAY)
            add_rect(slide, Inches(2.2 + i * stage_w + 0.08), Inches(y + 0.08),
                     Inches(0.15), Inches(0.15), color)
            add_text(slide, Inches(2.2 + i * stage_w + 0.30), Inches(y + 0.05),
                     Inches(stage_w - 0.35), Inches(row_h - 0.15),
                     text, size=9, color=CHARCOAL, font=FONT, anchor=MSO_ANCHOR.MIDDLE)
        y += row_h
        prev_branch = branch

    add_takeaway(slide,
        "Zero standardization across units: 6 different intake models, 6 prioritization logics, only 1 area uses a portfolio tool — fragmentation is structural, not local",
        color=ACCENT)
    add_footer(slide,
        "Source: Champions workshop (6 interviews, Apr 2026) + IT CS interview (Kumar & Stephanie); KB: Interview_Application_Champions.md, Interview_Stephanie_Kumar_IT_CS.md",
        "AS-IS Master Demand Map  |  06 / 15")


def slide_7_tooling(prs):
    slide = blank(prs)
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Eight platforms circulate between the two branches with near-zero overlap — the toolchain is fragmented by team, not by purpose")

    # Left: Applications tooling
    add_text(slide, Inches(0.5), Inches(1.50), Inches(6.1), Inches(0.35),
             "APPLICATIONS BRANCH — Carlos Garbin", size=11, bold=True, color=BLUE, font=FONT)

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

    # Right: Customer Services tooling
    add_text(slide, Inches(6.85), Inches(1.50), Inches(6.0), Inches(0.35),
             "CUSTOMER SERVICES / LOCAL SAP — Kumar Jaladanki", size=11, bold=True, color=NAVY, font=FONT)

    cs_tooling = [
        ("IT Now", "Incidents (template) — Metering", ["Smart Metering"], "Primary"),
        ("Email", "Small demand (Smart Metering)", ["Smart Metering"], "GAP"),
        ("Jira", "Used by PMs (Smart Metering)", ["Smart Metering"], "Specific"),
        ("Confluence", "Portfolio documentation (Metering only)", ["Smart Metering"], "Specific"),
        ("PPM Now", "Large projects — ONLY in Smart Metering", ["Smart Metering"], "Primary"),
        ("Travis PMO (tool TBC)", "Large project management", ["Branch-wide"], "Specific"),
        ("SWF + Accenture", "Build / delivery", ["Branch-wide"], "Specific"),
        ("4 other teams — tooling not yet captured", "Ivelisse, CRM, SAP Arch, SAP Basis", ["TBC"], "GAP"),
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
    add_footer(slide,
        "Source: KB synthesis — Interview_Application_Champions.md + Excel AS-IS canvas tooling column (Apr 2026)",
        "Tooling Landscape  |  07 / 15")


def slide_8_governance_forums(prs):
    slide = blank(prs)
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Five recurring forums exist — but none covers portfolio-level decisions, and Manager Check-in participants are still TBC")

    # Table header
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

    # Coverage gap call-outs
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
    add_footer(slide,
        "Source: Governance forums table (RFI from manager) + validation session + Interview_Application_Champions.md governance section",
        "Governance Forums  |  08 / 15")


def slide_9_pain_points_mece(prs):
    slide = blank(prs)
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Pain points cluster MECE across People, Process, Tools and Governance — each category produces its own bottleneck")

    # 2x2 grid
    categories = [
        {
            "label": "PEOPLE",
            "color": NAVY,
            "items": [
                "Turnover — \"1 year to hire someone\" (Global SAP)",
                "Unclear ownership — \"3 people can run payment jobs\" (Kumar/Stephanie)",
                "Capacity competition across 7 business OPCOs (IT CS)",
                "Heavy SWF + Accenture + Vendor dependency in delivery",
                "2 vacant Sr. Manager positions in Kumar's branch",
            ],
            "source": "Source: All 4 interview MDs"
        },
        {
            "label": "PROCESS",
            "color": NAVY,
            "items": [
                "Demand intake not standardized across 6+ channels (Carlos)",
                "Prioritization is ad-hoc; 4 informal criteria per forum",
                "PDD is the only formal process — only for large projects",
                "Vendor sets estimates, timelines, scope (especially Metering)",
                "PMs and Tech Leads \"need to work in the same way\" (Carlos)",
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
        # Category label
        add_text(slide, Inches(x + 0.25), Inches(y + 0.15), Inches(card_w - 0.5), Inches(0.35),
                 c["label"], size=13, bold=True, color=c["color"], font=FONT)
        # Items
        items_text = "\n".join(["▪  " + it for it in c["items"]])
        add_text(slide, Inches(x + 0.25), Inches(y + 0.55), Inches(card_w - 0.5), Inches(1.55),
                 items_text, size=10, color=CHARCOAL, font=FONT)
        # Source
        add_text(slide, Inches(x + 0.25), Inches(y + card_h - 0.30), Inches(card_w - 0.5), Inches(0.25),
                 c["source"], size=7, color=MID_GRAY, font=FONT, italic=True)

    add_takeaway(slide,
        "Each MECE category produces a distinct bottleneck — addressing only one dimension in the TOM will not resolve the system; all four must move together",
        color=NAVY)
    add_footer(slide,
        "Source: Synthesis across 4 interview MDs + Excel canvas — MECE grouping applied by NTT",
        "Pain Points — MECE  |  09 / 15")


def slide_10_apps_bottlenecks(prs):
    slide = blank(prs)
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Two Applications-branch bottlenecks dominate: G&E has no digital intake, and Global SAP is trapped by incident overload")

    cards = [
        {
            "num": "01",
            "stage": "BRANCH: Applications / Bruno Moreto  •  STAGE: Intake",
            "title": "G&E Email-Only Intake",
            "desc": "Zero digital trail. Every request arrives via email from business or vendor. No triage, no queue, no SLA. Each decision is individual and unreproducible. Andrew McGibbon is effectively the single point of routing.",
            "impact": "Unknown volume. No capacity forecast possible. Single-person dependency. Third-party risk + cyber/legal priorities compete without triage.",
            "evidence": "\"I receive requests/input from the business or from the vendor\" — Andrew McGibbon",
        },
        {
            "num": "02",
            "stage": "BRANCH: Applications / Vanessa Colombo  •  STAGE: Prioritization",
            "title": "Global SAP Incident Flood",
            "desc": "Incidents (IT Now) outrank all other demand. Small demands (Jira) and major projects (PDD + SharePoint) wait. P&T intermediates but cannot unblock the fundamental capacity shortage — turnover compounds the loss.",
            "impact": "Capacity trapped in BAU. Major projects (PDD) slip. Knowledge transfer breaks at every exit — \"1 year to hire someone\".",
            "evidence": "\"Sometimes entire team is focusing on incidents\" — Michael Dubois",
        },
    ]

    card_w = 6.1
    card_h = 4.85
    x0 = 0.5
    y0 = 1.55
    gap = 0.13

    for i, c in enumerate(cards):
        x = x0 + i * (card_w + gap)
        add_rect(slide, Inches(x), Inches(y0), Inches(card_w), Inches(card_h), LIGHT_GRAY)
        add_rect(slide, Inches(x), Inches(y0), Inches(card_w), Inches(0.08), RED_CRITICAL)
        # Number
        add_text(slide, Inches(x + 0.25), Inches(y0 + 0.2), Inches(1.5), Inches(0.9),
                 c["num"], size=42, bold=True, color=RED_CRITICAL, font=FONT)
        # Stage
        add_text(slide, Inches(x + 0.25), Inches(y0 + 1.15), Inches(card_w - 0.5), Inches(0.25),
                 c["stage"], size=9, bold=True, color=MID_GRAY, font=FONT)
        # Title
        add_text(slide, Inches(x + 0.25), Inches(y0 + 1.45), Inches(card_w - 0.5), Inches(0.45),
                 c["title"], size=18, bold=True, color=NAVY, font=FONT)
        # Description
        add_text(slide, Inches(x + 0.25), Inches(y0 + 2.05), Inches(card_w - 0.5), Inches(1.5),
                 c["desc"], size=10, color=CHARCOAL, font=FONT)
        # Impact
        add_text(slide, Inches(x + 0.25), Inches(y0 + 3.65), Inches(card_w - 0.5), Inches(0.25),
                 "IMPACT", size=9, bold=True, color=RED_CRITICAL, font=FONT)
        add_text(slide, Inches(x + 0.25), Inches(y0 + 3.90), Inches(card_w - 0.5), Inches(0.70),
                 c["impact"], size=10, color=CHARCOAL, font=FONT)
        # Evidence
        add_rect(slide, Inches(x + 0.25), Inches(y0 + card_h - 0.55), Inches(card_w - 0.5), Inches(0.42),
                 NAVY)
        add_text(slide, Inches(x + 0.40), Inches(y0 + card_h - 0.50), Inches(card_w - 0.75), Inches(0.30),
                 c["evidence"], size=9, color=WHITE, font=FONT, italic=True, anchor=MSO_ANCHOR.MIDDLE)

    add_takeaway(slide,
        "These two Applications bottlenecks bracket the demand funnel — one breaks the entry, the other chokes the middle; both must close in the TOM",
        color=ACCENT)
    add_footer(slide,
        "Source: Interview_Application_Champions.md — G&E + Global SAP sections",
        "Applications Bottlenecks  |  10 / 15")


def slide_11_cs_bottlenecks(prs):
    slide = blank(prs)
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Two Customer Services bottlenecks: leadership gaps in Kumar's branch + zero portfolio visibility outside Smart Metering")

    cards = [
        {
            "num": "03",
            "stage": "BRANCH: Customer Services / Kumar  •  STAGES: Classification + Prioritization",
            "title": "Customer Services Leadership Gap",
            "desc": "Two vacant Sr. Manager positions (CRM & Contact Center Ops, SAP & CS Architecture). Prioritization happens director-level only across Billing, CRM, Metering and Basis — \"intrinsic common sense\" (Metering) is the most explicit criterion on record. Travis PMO handles large projects independently.",
            "impact": "No shared prioritization across Billing / CRM / Metering / Basis. Consolidated branch-level view of demand is a structural blind spot for the TOM.",
            "evidence": "\"Who is in charge of what — 3 people can run the payment jobs\" — Kumar & Stephanie",
        },
        {
            "num": "04",
            "stage": "BRANCH: Both (emphasis on Customer Services)  •  STAGES: Portfolio + Delivery",
            "title": "Portfolio & Vendor Tooling Gap",
            "desc": "No tool supports portfolio management in 5 of 6 units — only Smart Metering uses PPM Now. No process governs vendor delivery. In Smart Metering, \"all portfolio in vendor\" (100% outsourced) — vendors estimate, do requirements and code.",
            "impact": "Vendor sets timelines and estimates. Budget/schedule control degraded. ROI from IT investment is untraceable back to business outcomes.",
            "evidence": "\"No tool to support portfolio management. No tools or process to manage delivery process of vendors\"",
        },
    ]

    card_w = 6.1
    card_h = 4.85
    x0 = 0.5
    y0 = 1.55
    gap = 0.13

    for i, c in enumerate(cards):
        x = x0 + i * (card_w + gap)
        add_rect(slide, Inches(x), Inches(y0), Inches(card_w), Inches(card_h), LIGHT_GRAY)
        add_rect(slide, Inches(x), Inches(y0), Inches(card_w), Inches(0.08), RED_CRITICAL)
        add_text(slide, Inches(x + 0.25), Inches(y0 + 0.2), Inches(1.5), Inches(0.9),
                 c["num"], size=42, bold=True, color=RED_CRITICAL, font=FONT)
        add_text(slide, Inches(x + 0.25), Inches(y0 + 1.15), Inches(card_w - 0.5), Inches(0.25),
                 c["stage"], size=9, bold=True, color=MID_GRAY, font=FONT)
        add_text(slide, Inches(x + 0.25), Inches(y0 + 1.45), Inches(card_w - 0.5), Inches(0.45),
                 c["title"], size=17, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(x + 0.25), Inches(y0 + 2.05), Inches(card_w - 0.5), Inches(1.5),
                 c["desc"], size=10, color=CHARCOAL, font=FONT)
        add_text(slide, Inches(x + 0.25), Inches(y0 + 3.65), Inches(card_w - 0.5), Inches(0.25),
                 "IMPACT", size=9, bold=True, color=RED_CRITICAL, font=FONT)
        add_text(slide, Inches(x + 0.25), Inches(y0 + 3.90), Inches(card_w - 0.5), Inches(0.70),
                 c["impact"], size=10, color=CHARCOAL, font=FONT)
        add_rect(slide, Inches(x + 0.25), Inches(y0 + card_h - 0.55), Inches(card_w - 0.5), Inches(0.42),
                 NAVY)
        add_text(slide, Inches(x + 0.40), Inches(y0 + card_h - 0.50), Inches(card_w - 0.75), Inches(0.30),
                 c["evidence"], size=9, color=WHITE, font=FONT, italic=True, anchor=MSO_ANCHOR.MIDDLE)

    add_takeaway(slide,
        "Customer Services has both structural (leadership) and operational (tooling) gaps — resolving one without the other leaves the branch unable to scale",
        color=ACCENT)
    add_footer(slide,
        "Source: Interview_Stephanie_Kumar_IT_CS.md + Interview_Application_Champions.md (Smart Metering) + Excel portfolio gaps",
        "Customer Services Bottlenecks  |  11 / 15")


def slide_12_cross_branch(prs):
    slide = blank(prs)
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
        # Card
        add_rect(slide, Inches(0.5), Inches(y), Inches(12.33), Inches(row_h - 0.08), LIGHT_GRAY)
        # Left accent
        add_rect(slide, Inches(0.5), Inches(y), Inches(0.08), Inches(row_h - 0.08), ACCENT)
        # Number
        add_rect(slide, Inches(0.65), Inches(y + 0.15), Inches(0.6), Inches(0.6), ACCENT)
        add_text(slide, Inches(0.65), Inches(y + 0.15), Inches(0.6), Inches(0.6),
                 f"0{i+1}", size=16, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Title
        add_text(slide, Inches(1.40), Inches(y + 0.10), Inches(11.5), Inches(0.35),
                 issue["title"], size=13, bold=True, color=NAVY, font=FONT)
        # Detail
        add_text(slide, Inches(1.40), Inches(y + 0.45), Inches(11.5), Inches(0.55),
                 issue["detail"], size=9, color=CHARCOAL, font=FONT)
        # Branch-specific row
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
    add_footer(slide,
        "Source: KB synthesis across all 4 interview MDs + Excel + Governance forums table",
        "Cross-Branch Systemic Issues  |  12 / 15")


def slide_13_global_vs_local_sap(prs):
    slide = blank(prs)
    set_bg(slide, WHITE)
    add_action_title(slide,
        "\"Global SAP versus Local SAP — fundamental tension to resolve\" — Carlos Garbin's strategic framing, now mapped to the org structure")

    # Two side-by-side comparison cards
    # Global SAP
    x = 0.5
    y = 1.55
    w = 6.15
    h = 4.85
    add_rect(slide, Inches(x), Inches(y), Inches(w), Inches(h), LIGHT_GRAY)
    add_rect(slide, Inches(x), Inches(y), Inches(w), Inches(0.55), BLUE)
    add_text(slide, Inches(x), Inches(y), Inches(w), Inches(0.55),
             "GLOBAL SAP  —  Iberdrola (Spain)", size=14, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    global_items = [
        ("BRANCH", "Applications — Carlos Garbin"),
        ("MANAGER", "Vanessa Colombo (Manager IT Apps Global SAP)"),
        ("TEAM", "5 FTE (Michael Dubois + 4 others)"),
        ("DOMAIN", "Corporate finance functions — FI, PO, Budget control"),
        ("ENVIRONMENT", "Iberdrola / Spain SAP instance"),
        ("TOOLING", "IT Now + RITM + Jira + SharePoint + Teams (mature)"),
        ("PRIORITIZATION", "Incidents first, then small + major projects (P&T)"),
        ("BOTTLENECK", "Incident flood — \"entire team focused on incidents\""),
        ("COVERAGE", "✅ Champion interviewed (Michael Dubois)"),
    ]
    yi = y + 0.75
    for label, val in global_items:
        add_text(slide, Inches(x + 0.25), Inches(yi), Inches(2.3), Inches(0.35),
                 label, size=8, bold=True, color=BLUE, font=FONT)
        add_text(slide, Inches(x + 2.40), Inches(yi), Inches(w - 2.65), Inches(0.35),
                 val, size=9, color=CHARCOAL, font=FONT)
        yi += 0.43

    # Local SAP
    x2 = 7.18
    add_rect(slide, Inches(x2), Inches(y), Inches(w), Inches(h), LIGHT_GRAY)
    add_rect(slide, Inches(x2), Inches(y), Inches(w), Inches(0.55), NAVY)
    add_text(slide, Inches(x2), Inches(y), Inches(w), Inches(0.55),
             "LOCAL SAP  —  Avangrid (US)", size=14, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    local_items = [
        ("BRANCH", "Customer Services — Kumar Jaladanki"),
        ("MANAGERS", "Vincent Lam (Device & Metering), Ivelisse (FICA/Billing), TBD (CRM), Open (SAP Arch), Nagi Reddy (Basis)"),
        ("TEAM", "~24 FTE across 5 sub-teams"),
        ("DOMAIN", "Customer services — Billing, CRM, FICA, Metering"),
        ("ENVIRONMENT", "Avangrid / US SAP instance"),
        ("TOOLING", "Fragmented: IT Now, Email, Jira, Confluence, PPM Now (Metering only)"),
        ("PRIORITIZATION", "Director-level + \"intrinsic common sense\" (Metering)"),
        ("BOTTLENECK", "Leadership gaps + 4 teams uninterviewed"),
        ("COVERAGE", "⚠ Partial — only Smart Metering (Iram Chowdhury)"),
    ]
    yi = y + 0.75
    for label, val in local_items:
        add_text(slide, Inches(x2 + 0.25), Inches(yi), Inches(2.3), Inches(0.35),
                 label, size=8, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(x2 + 2.40), Inches(yi), Inches(w - 2.65), Inches(0.35),
                 val, size=9, color=CHARCOAL, font=FONT)
        yi += 0.43

    add_takeaway(slide,
        "Global and Local SAP answer different business questions on different geographies — the TOM must preserve their specialization while introducing shared intake, governance and tooling",
        color=ACCENT)
    add_footer(slide,
        "Source: Interview_Carlos_Applications.md (Global vs Local framing) + Org_Chart_Digital_Products.md + Interview_Application_Champions.md",
        "Global vs Local SAP  |  13 / 15")


def slide_14_structural_issues_tom(prs):
    slide = blank(prs)
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
         "TOM: vendor performance framework + independent estimation sanity check."),
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
        add_rect(slide, Inches(0.5), Inches(y), Inches(0.55), Inches(row_h - 0.08), ACCENT)
        add_text(slide, Inches(0.5), Inches(y), Inches(0.55), Inches(row_h - 0.08),
                 num, size=16, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, Inches(1.2), Inches(y), Inches(6.8), Inches(0.28),
                 title, size=13, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(1.2), Inches(y + 0.30), Inches(6.8), Inches(0.35),
                 gap, size=11, color=CHARCOAL, font=FONT)
        add_text(slide, Inches(8.2), Inches(y + 0.05), Inches(4.6), Inches(0.58),
                 "→  " + tom, size=11, color=BLUE, font=FONT, italic=True,
                 anchor=MSO_ANCHOR.MIDDLE)
        y += row_h

    add_takeaway(slide,
        "These 7 issues define the TOM scope — each has a dedicated workstream in the design phase that follows this Baseline",
        color=NAVY)
    add_footer(slide,
        "Source: KB synthesis — 4 interview MDs + Org_Chart + Excel (reused from Demand Map Structural Insights)",
        "TOM Responses  |  13 / 15")


def slide_15_measurement_dependencies(prs):
    slide = blank(prs)
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Six measurement dependencies remain open — they shape the TOM business case and must be addressed alongside the design")

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

    col_w = 4.10
    row_h = 2.20
    x0 = 0.5
    y0 = 1.55

    for idx, (num, title, detail, question) in enumerate(gaps):
        col = idx % 3
        row = idx // 3
        x = x0 + col * (col_w + 0.10)
        y = y0 + row * (row_h + 0.15)

        add_rect(slide, Inches(x), Inches(y), Inches(col_w), Inches(row_h), LIGHT_GRAY)
        add_rect(slide, Inches(x), Inches(y), Inches(col_w), Inches(0.06), ACCENT)
        add_text(slide, Inches(x + 0.25), Inches(y + 0.18), Inches(0.6), Inches(0.5),
                 num, size=22, bold=True, color=ACCENT, font=FONT)
        add_text(slide, Inches(x + 0.95), Inches(y + 0.25), Inches(col_w - 1.1), Inches(0.4),
                 title, size=13, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(x + 0.25), Inches(y + 0.85), Inches(col_w - 0.5), Inches(0.95),
                 detail, size=9, color=CHARCOAL, font=FONT)
        add_text(slide, Inches(x + 0.25), Inches(y + row_h - 0.45), Inches(col_w - 0.5), Inches(0.30),
                 "KEY QUESTION", size=7, bold=True, color=MID_GRAY, font=FONT)
        add_text(slide, Inches(x + 0.25), Inches(y + row_h - 0.25), Inches(col_w - 0.5), Inches(0.28),
                 question, size=10, italic=True, color=BLUE, font=FONT)

    add_takeaway(slide,
        "Without these measurements, the TOM business case stays qualitative — instrumentation should run in parallel with the operating model design",
        color=NAVY)
    add_footer(slide,
        "Framing: open measurement items — flagged transparently rather than fabricated",
        "Measurement Dependencies  |  15 / 15")


def slide_16_risks(prs):
    slide = blank(prs)
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
            "title": "Global vs Local SAP drift",
            "detail": "With no enterprise-level reconciliation, the two SAP footprints evolve independently. Integration cost grows silently each quarter.",
            "severity": "MEDIUM",
        },
    ]

    severity_color = {"CRITICAL": RED_CRITICAL, "HIGH": YELLOW_GAP, "MEDIUM": MID_GRAY}

    y = 1.55
    row_h = 0.95
    for r in risks:
        add_rect(slide, Inches(0.5), Inches(y), Inches(12.33), Inches(row_h - 0.08), LIGHT_GRAY)
        # ID badge
        sev_color = severity_color[r["severity"]]
        add_rect(slide, Inches(0.5), Inches(y), Inches(0.08), Inches(row_h - 0.08), sev_color)
        add_rect(slide, Inches(0.65), Inches(y + 0.15), Inches(0.7), Inches(0.55), sev_color)
        add_text(slide, Inches(0.65), Inches(y + 0.15), Inches(0.7), Inches(0.55),
                 r["num"], size=14, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Title
        add_text(slide, Inches(1.5), Inches(y + 0.10), Inches(9.5), Inches(0.3),
                 r["title"], size=12, bold=True, color=NAVY, font=FONT)
        # Detail
        add_text(slide, Inches(1.5), Inches(y + 0.42), Inches(9.5), Inches(0.45),
                 r["detail"], size=9, color=CHARCOAL, font=FONT)
        # Severity pill
        add_rect(slide, Inches(11.20), Inches(y + 0.30), Inches(1.55), Inches(0.30), sev_color, corner=True)
        add_text(slide, Inches(11.20), Inches(y + 0.30), Inches(1.55), Inches(0.30),
                 r["severity"], size=8, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        y += row_h

    add_takeaway(slide,
        "These risks compound — the longer the TOM is deferred, the harder the reset. The Baseline recommends beginning TOM design this quarter",
        color=ACCENT)
    add_footer(slide,
        "Source: NTT risk synthesis derived from KB findings — severity calibrated against growth + governance gap context",
        "Risks if Not Addressed  |  14 / 15")


def slide_17_tom_principles(prs):
    slide = blank(prs)
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

    # 7 principles in 2 columns: 4 left + 3 right
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

    # Plus two NTT additions as callout
    callout_y = 6.10
    add_rect(slide, Inches(6.78), Inches(callout_y), Inches(6.15), Inches(0.42), ACCENT)
    add_text(slide, Inches(6.95), Inches(callout_y), Inches(6.0), Inches(0.42),
             "+ From this Baseline: BRANCH-UNIFIED (Apps + CS) and PRODUCT-ORIENTED (Value Streams, not silos)",
             size=10, bold=True, color=WHITE, font=FONT, anchor=MSO_ANCHOR.MIDDLE)

    add_takeaway(slide,
        "These principles are not new — they extend the Transformation Vision already defined in SAP CoE New Ways of Working v2.0 with two baseline-specific additions",
        color=NAVY)
    add_footer(slide,
        "Source: SAP CoE New Ways of Working v2.0 (Slide 7, 7 Transformation Dimensions) + NTT baseline additions",
        "TOM Design Principles  |  15 / 15")


def slide_18_next_steps(prs):
    slide = blank(prs)
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Four concurrent workstreams move from Baseline to TOM design — kickoff recommended this quarter")

    workstreams = [
        {
            "num": "W1",
            "title": "Close Coverage Gaps",
            "owner": "NTT + Client sponsors",
            "timeline": "2 weeks",
            "items": [
                "Interview Ivelisse (FICA/Billing)",
                "Interview Nagi Reddy (SAP Basis)",
                "Fill CRM Sr. Manager vacancy",
                "Fill SAP & CS Architecture vacancy",
            ],
        },
        {
            "num": "W2",
            "title": "Measurement Instrumentation",
            "owner": "NTT Analytics + Avangrid IT Ops",
            "timeline": "4–6 weeks",
            "items": [
                "Volumetrics per channel",
                "Lead / Cycle Time baseline",
                "Taxonomy thresholds",
                "Capacity breakdown (FTE × Vendor)",
            ],
        },
        {
            "num": "W3",
            "title": "TOM Design",
            "owner": "NTT + Carlos + Kumar",
            "timeline": "6–8 weeks (parallel to W2)",
            "items": [
                "Unified intake design",
                "Value-based scoring model",
                "RACI per lifecycle phase",
                "Jira + PPM Now integration architecture",
            ],
        },
        {
            "num": "W4",
            "title": "Governance Upgrade",
            "owner": "VP Vitor Cruz + CIO",
            "timeline": "Immediate",
            "items": [
                "Define Portfolio Steering cadence",
                "Confirm Manager Check-in scope",
                "Vendor performance framework",
                "Quarterly enterprise SAP sync",
            ],
        },
    ]

    card_w = 3.05
    card_h = 4.6
    gap = 0.1
    y0 = 1.55
    x0 = 0.5

    for i, w in enumerate(workstreams):
        x = x0 + i * (card_w + gap)
        add_rect(slide, Inches(x), Inches(y0), Inches(card_w), Inches(card_h), LIGHT_GRAY)
        add_rect(slide, Inches(x), Inches(y0), Inches(card_w), Inches(0.6), NAVY)
        add_text(slide, Inches(x + 0.2), Inches(y0 + 0.05), Inches(card_w - 0.4), Inches(0.25),
                 w["num"], size=9, bold=True, color=BLUE, font=FONT)
        add_text(slide, Inches(x + 0.2), Inches(y0 + 0.25), Inches(card_w - 0.4), Inches(0.35),
                 w["title"], size=12, bold=True, color=WHITE, font=FONT)
        # Owner + timeline
        add_text(slide, Inches(x + 0.2), Inches(y0 + 0.75), Inches(card_w - 0.4), Inches(0.25),
                 "OWNER", size=7, bold=True, color=MID_GRAY, font=FONT)
        add_text(slide, Inches(x + 0.2), Inches(y0 + 0.95), Inches(card_w - 0.4), Inches(0.25),
                 w["owner"], size=9, color=CHARCOAL, font=FONT)
        add_text(slide, Inches(x + 0.2), Inches(y0 + 1.30), Inches(card_w - 0.4), Inches(0.25),
                 "TIMELINE", size=7, bold=True, color=MID_GRAY, font=FONT)
        add_text(slide, Inches(x + 0.2), Inches(y0 + 1.50), Inches(card_w - 0.4), Inches(0.25),
                 w["timeline"], size=9, color=ACCENT, bold=True, font=FONT)
        # Items
        add_text(slide, Inches(x + 0.2), Inches(y0 + 1.85), Inches(card_w - 0.4), Inches(0.25),
                 "DELIVERABLES", size=7, bold=True, color=MID_GRAY, font=FONT)
        items_text = "\n".join(["▪  " + it for it in w["items"]])
        add_text(slide, Inches(x + 0.2), Inches(y0 + 2.10), Inches(card_w - 0.4), Inches(2.4),
                 items_text, size=9, color=CHARCOAL, font=FONT)

    add_takeaway(slide,
        "Demand Map (done) → Baseline Report (this) → TOM Design (next). NTT proposes workstream kickoff within 2 weeks",
        color=NAVY)
    add_footer(slide,
        "Timeline proposed by NTT based on Baseline findings — requires client confirmation on capacity and governance cadence",
        "Next Steps & Handoff  |  16 / 15")


# ========== MAIN ==========

def build():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    slide_1_cover(prs)
    slide_2_executive_summary(prs)
    slide_3_context(prs)
    slide_4_methodology(prs)
    slide_5_org_landscape(prs)
    slide_6_master_map(prs)
    slide_7_tooling(prs)
    slide_8_governance_forums(prs)
    slide_9_pain_points_mece(prs)
    slide_10_apps_bottlenecks(prs)
    slide_11_cs_bottlenecks(prs)
    slide_12_cross_branch(prs)
    # (Removed: slide_13_global_vs_local_sap — assumption unvalidated)
    slide_14_structural_issues_tom(prs)
    slide_16_risks(prs)
    slide_17_tom_principles(prs)
    # (Removed: slide_18_next_steps — user request v1.2)
    # Measurement Dependencies also removed per v1.3 scope trim

    out = "/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Deliverables/Baseline_Report/Baseline_Report_v1.pptx"
    prs.save(out)
    print(f"Saved: {out}")
    print(f"Slides: {len(prs.slides)}")


if __name__ == "__main__":
    build()
