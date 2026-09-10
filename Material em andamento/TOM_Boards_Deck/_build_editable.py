"""
Build EDITABLE PPTX for TOM Boards & Intake deck.
Every text is in a native text box; every shape is a native autoshape.
Preserves information architecture from SVG slides; visual is simpler but fully editable.
"""
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Emu, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ============ COLOR PALETTE ============
NAVY = RGBColor(0x0E, 0x22, 0x40)
NAVY_2 = RGBColor(0x17, 0x34, 0x58)
NAVY_3 = RGBColor(0x1F, 0x4E, 0x79)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREY_LIGHT = RGBColor(0x9F, 0xB3, 0xD1)
GREY_MED = RGBColor(0x6F, 0x87, 0xA8)
GREY_DARK = RGBColor(0x42, 0x52, 0x6E)
ORANGE = RGBColor(0xF0, 0x8A, 0x3E)
ORANGE_DARK = RGBColor(0xE8, 0x70, 0x2A)
# VS-1 CS amber
AMBER = RGBColor(0xE8, 0xA2, 0x4C)
AMBER_LIGHT = RGBColor(0xF4, 0xC9, 0x88)
AMBER_DARK = RGBColor(0x5A, 0x3A, 0x0E)
# VS-2 Field green
GREEN = RGBColor(0x5D, 0xA4, 0x60)
GREEN_LIGHT = RGBColor(0xC2, 0xE5, 0xC2)
GREEN_DARK = RGBColor(0x1F, 0x57, 0x22)
# VS-3 Corporate blue
BLUE = RGBColor(0x4A, 0x90, 0xB8)
BLUE_LIGHT = RGBColor(0xB7, 0xDD, 0xEF)
BLUE_DARK = RGBColor(0x0F, 0x4A, 0x6B)
# Platform purple
PURPLE = RGBColor(0x8B, 0x6F, 0xB8)
PURPLE_LIGHT = RGBColor(0xD4, 0xC4, 0xE8)
PURPLE_DARK = RGBColor(0x5C, 0x2A, 0x91)
# Expedite pink
PINK = RGBColor(0xB8, 0x35, 0x67)
# Atlassian Jira
JIRA_BLUE = RGBColor(0x07, 0x47, 0xA6)
JIRA_LIGHT = RGBColor(0xDE, 0xEB, 0xFF)
RED = RGBColor(0xDE, 0x35, 0x0B)
YELLOW = RGBColor(0xFF, 0xAB, 0x00)
SUCCESS = RGBColor(0x36, 0xB3, 0x7E)


# ============ SETUP ============
DECK_DIR = Path(__file__).parent
OUT = DECK_DIR / "TOM_Boards_Intake_EDITABLE.pptx"
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


# ============ HELPERS ============
def set_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_rect(slide, x, y, w, h, fill_color=None, line_color=None, line_width=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.shadow.inherit = False
    if fill_color is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill_color
    if line_color is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_color
        if line_width is not None:
            shp.line.width = line_width
    return shp


def add_round(slide, x, y, w, h, fill_color=None, line_color=None, line_width=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shp.shadow.inherit = False
    shp.adjustments[0] = 0.10
    if fill_color is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill_color
    if line_color is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_color
        if line_width is not None:
            shp.line.width = line_width
    return shp


def add_text(slide, x, y, w, h, text, font_size=12, bold=False, italic=False,
             color=NAVY, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font="Helvetica Neue"):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tb.text_frame.margin_left = Emu(36000)
    tb.text_frame.margin_right = Emu(36000)
    tb.text_frame.margin_top = Emu(18000)
    tb.text_frame.margin_bottom = Emu(18000)
    tb.text_frame.word_wrap = True
    tb.text_frame.vertical_anchor = anchor
    p = tb.text_frame.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.name = font
    r.font.size = Pt(font_size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.color.rgb = color
    return tb


def add_multiline(slide, x, y, w, h, lines, font_size=11, color=WHITE, align=PP_ALIGN.LEFT, font="Helvetica Neue"):
    """lines: list of (text, bold, italic) tuples"""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tb.text_frame.margin_left = Emu(36000)
    tb.text_frame.margin_right = Emu(36000)
    tb.text_frame.margin_top = Emu(18000)
    tb.text_frame.margin_bottom = Emu(18000)
    tb.text_frame.word_wrap = True
    for i, item in enumerate(lines):
        if isinstance(item, str):
            text, bold, italic = item, False, False
        else:
            text = item[0]
            bold = item[1] if len(item) > 1 else False
            italic = item[2] if len(item) > 2 else False
        if i == 0:
            p = tb.text_frame.paragraphs[0]
        else:
            p = tb.text_frame.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = text
        r.font.name = font
        r.font.size = Pt(font_size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color
    return tb


def slide_header(slide, num, section, title, subtitle=None):
    """Standard slide header — number tag + title."""
    add_text(slide, Inches(0.4), Inches(0.25), Inches(8), Inches(0.3),
             f"{num:02d}  ·  {section}", font_size=10, bold=True, color=ORANGE)
    add_text(slide, Inches(0.4), Inches(0.55), Inches(12.5), Inches(0.5),
             title, font_size=24, bold=True, color=WHITE)
    if subtitle:
        add_text(slide, Inches(0.4), Inches(1.05), Inches(12.5), Inches(0.3),
                 subtitle, font_size=11, color=GREY_LIGHT)


def slide_footer(slide, num, total, source=""):
    add_text(slide, Inches(0.4), Inches(7.2), Inches(11), Inches(0.25),
             source, font_size=8, italic=True, color=GREY_MED)
    add_text(slide, Inches(12.4), Inches(7.2), Inches(0.8), Inches(0.25),
             f"{num:02d} / {total}", font_size=8, color=GREY_MED, align=PP_ALIGN.RIGHT)


# ============ SLIDE 01 — COVER ============
def slide_01_cover():
    s = prs.slides.add_slide(BLANK)
    set_bg(s, NAVY)
    # Color stripes left
    add_rect(s, 0, 0, Inches(0.08), SH, AMBER)
    add_rect(s, Inches(0.08), 0, Inches(0.08), SH, GREEN)
    add_rect(s, Inches(0.16), 0, Inches(0.08), SH, BLUE)
    add_rect(s, Inches(0.24), 0, Inches(0.08), SH, PURPLE)
    # Brand strip
    add_text(s, Inches(0.7), Inches(0.7), Inches(10), Inches(0.3),
             "AVANGRID  ·  DIGITAL PRODUCTS  ·  TOM PROPOSAL",
             font_size=12, bold=True, color=GREY_LIGHT)
    # Title
    add_text(s, Inches(0.7), Inches(2.2), Inches(12), Inches(1),
             "Boards & Intake Model",
             font_size=54, bold=True, color=WHITE)
    add_text(s, Inches(0.7), Inches(3.2), Inches(12), Inches(0.6),
             "Three Value Streams · Seven Squads · Platform Teams",
             font_size=28, bold=True, color=AMBER)
    add_text(s, Inches(0.7), Inches(3.9), Inches(12), Inches(0.4),
             "How Customer Services Apps and IT Business Applications work as one integrated flow system",
             font_size=13, color=GREY_LIGHT)
    # VS legend cards
    legend_y = Inches(5)
    legend_data = [
        ("VS-1", "M2C & Customer Service", "~20 FTE · CS branch · Kumar", AMBER, AMBER_DARK),
        ("VS-2", "Field & Network Operations", "~13 FTE · ITBA · Carlos", GREEN, WHITE),
        ("VS-3", "Corporate & Global SAP", "~13 FTE · ITBA · Carlos", BLUE, WHITE),
        ("PLATFORM", "Cross-cutting enablers", "~8 FTE · serve all VS", PURPLE, WHITE),
    ]
    for i, (tag, title, sub, fill, txt) in enumerate(legend_data):
        x = Inches(0.7 + i * 3.05)
        add_round(s, x, legend_y, Inches(2.9), Inches(1), fill)
        add_text(s, x + Inches(0.15), legend_y + Inches(0.1), Inches(2.6), Inches(0.3),
                 tag, font_size=14, bold=True, color=txt)
        add_text(s, x + Inches(0.15), legend_y + Inches(0.4), Inches(2.6), Inches(0.3),
                 title, font_size=12, bold=True, color=txt)
        add_text(s, x + Inches(0.15), legend_y + Inches(0.65), Inches(2.6), Inches(0.3),
                 sub, font_size=9, color=txt)
    # Footer
    add_text(s, Inches(0.7), Inches(6.7), Inches(8), Inches(0.3),
             "VITOR CRUZ — VP, IT BUSINESS SOLUTIONS DELIVERY",
             font_size=11, bold=True, color=GREY_LIGHT)
    add_text(s, Inches(0.7), Inches(6.95), Inches(10), Inches(0.3),
             'Building on "Customer Services — Proposed Value Streams" (April 2026) and AS-IS deliverable v.final',
             font_size=9, italic=True, color=GREY_MED)
    add_text(s, Inches(11.3), Inches(6.7), Inches(2), Inches(0.3),
             "v1.0  ·  April 2026", font_size=12, bold=True, color=ORANGE, align=PP_ALIGN.RIGHT)
    add_text(s, Inches(11.3), Inches(6.95), Inches(2), Inches(0.3),
             "NTT DATA Consulting", font_size=9, color=GREY_MED, align=PP_ALIGN.RIGHT)


# ============ SLIDE 02 — Overview ============
def slide_02_overview():
    s = prs.slides.add_slide(BLANK)
    set_bg(s, NAVY)
    slide_header(s, 2, "OVERVIEW", "Three Value Streams · Seven Squads · ~50 FTE",
                 "Stream-aligned model — each VS owns its applications, backlog, and board. Platform teams enable as-a-service.")

    # 3 VS columns
    cols = [
        # (x_in, fill, accent, dark_text, title, subtitle, squads, pm, board, risk)
        (0.3, AMBER_LIGHT, AMBER_DARK, AMBER_DARK,
         "VS-1 · M2C & Customer Service",
         "CS branch · Kumar Jaladanki · ~20 FTE",
         [
             ("Squad M2C-SL — Service Lifecycle", "Vincent Lam · Iram Champion · 8 FTE · SAP Metering · MV90", "100% vendor-delivered"),
             ("Squad M2C-RO — Revenue Operations", "Ivelisse Velazquez · Stephanie · 8 FTE · SAP FICA · ECC · StreamServe", "Accenture occasional"),
             ("Squad CCC — Customer Care & Contact", "VACANT 🔴 · 4 FTE · SAP CRM · Medallia · IVR · Blue Prism", "Lemur · IVR · Medallia · Digital agency"),
         ],
         "Heather LaJoie + Digital CoE Functional SAP",
         "Kumar Jaladanki — CS Portfolio Board",
         "Squad CCC vacant leadership · M2C-SL 100% vendor lock-in"),
        (4.65, GREEN_LIGHT, GREEN_DARK, GREEN_DARK,
         "VS-2 · Field & Network Operations",
         "ITBA branch · Carlos Garbin · ~13 FTE",
         [
             ("Squad G&E — Electric & Gas Ops", "Bruno Moreto · Andrew McGibbon Champion · 7 FTE", "70–90 apps · Android/tablets · doc viewer"),
             ("Squad GIS & OSG", "George Porto · George Porto Champion · 6 FTE", "SAP HANA · ARCFM · ProjectWise · Microstation"),
         ],
         "P&T pool (existing intermediary)",
         "Carlos Garbin — ITBA-Field Portfolio Board",
         "Break-fix dominant · no SLA tracking · no volume tracking"),
        (9.0, BLUE_LIGHT, BLUE_DARK, BLUE_DARK,
         "VS-3 · Corporate & Global SAP",
         "ITBA branch · Carlos Garbin · ~13 FTE",
         [
             ("Squad Corporate Apps", "Sarath Garikapati · Amit Patel · Stephanie SAP · 8 FTE", "Finance · HR · Procurement · Legal · Reg."),
             ("Squad Global SAP", "Vanessa Colombo · Michael Dubois · 5 FTE", "Iberdrola · FI · PO · Budget control"),
         ],
         "P&T pool (existing intermediary)",
         "Carlos Garbin — ITBA-Corporate Portfolio Board",
         "Squad Global SAP: incident overload caps strategic capacity"),
    ]
    for (xi, fill, accent, txt, title, subtitle, squads, pm, board, risk) in cols:
        x = Inches(xi)
        w = Inches(4.05)
        # column container (dark)
        add_round(s, x, Inches(1.5), w, Inches(5.4), NAVY_2, accent, Pt(2))
        # header
        add_round(s, x, Inches(1.5), w, Inches(0.6), accent)
        add_text(s, x + Inches(0.1), Inches(1.55), w - Inches(0.2), Inches(0.3),
                 title, font_size=14, bold=True, color=WHITE)
        add_text(s, x + Inches(0.1), Inches(1.85), w - Inches(0.2), Inches(0.25),
                 subtitle, font_size=10, color=WHITE)
        # squads
        squad_y = 2.25
        for sq_title, sq_lead, sq_extra in squads:
            add_round(s, x + Inches(0.1), Inches(squad_y), w - Inches(0.2), Inches(0.85), fill)
            add_text(s, x + Inches(0.2), Inches(squad_y + 0.05), w - Inches(0.3), Inches(0.3),
                     sq_title, font_size=11, bold=True, color=txt)
            add_text(s, x + Inches(0.2), Inches(squad_y + 0.32), w - Inches(0.3), Inches(0.25),
                     sq_lead, font_size=9, color=NAVY)
            add_text(s, x + Inches(0.2), Inches(squad_y + 0.58), w - Inches(0.3), Inches(0.25),
                     "Vendor: " + sq_extra, font_size=8, italic=True, color=GREY_DARK)
            squad_y += 0.93
        # PM steward
        add_text(s, x + Inches(0.15), Inches(5.2), w - Inches(0.3), Inches(0.25),
                 "PM STEWARD", font_size=9, bold=True, color=GREY_LIGHT)
        add_text(s, x + Inches(0.15), Inches(5.4), w - Inches(0.3), Inches(0.3),
                 pm, font_size=11, bold=True, color=WHITE)
        # Board
        add_text(s, x + Inches(0.15), Inches(5.75), w - Inches(0.3), Inches(0.25),
                 "BOARD OWNER", font_size=9, bold=True, color=GREY_LIGHT)
        add_text(s, x + Inches(0.15), Inches(5.95), w - Inches(0.3), Inches(0.3),
                 board, font_size=10, color=WHITE)
        # Risk
        add_round(s, x + Inches(0.1), Inches(6.35), w - Inches(0.2), Inches(0.5), NAVY, RED, Pt(1))
        add_text(s, x + Inches(0.2), Inches(6.4), w - Inches(0.3), Inches(0.4),
                 "⚠ " + risk, font_size=9, color=WHITE)

    slide_footer(s, 2, 13,
                 "Source: AS-IS deliverable v.final · Avangrid_Hierarchy_Overview.md · Customer Services — Proposed Value Streams (April 2026) · Champion interviews")


# ============ SLIDE 03-05 — VS Detail ============
def slide_vs_detail(num, vs_label, vs_title, vs_subtitle, accent, light, dark,
                    squads, pm_section, planning_section, forums_section, source):
    s = prs.slides.add_slide(BLANK)
    set_bg(s, NAVY)
    # Color top stripe
    add_rect(s, 0, 0, SW, Inches(0.08), accent)
    slide_header(s, num, "VALUE STREAM DETAIL", vs_title, vs_subtitle)

    # Squads row
    n_squads = len(squads)
    sq_w = (12.6 - (n_squads - 1) * 0.2) / n_squads
    for i, (sq_title, sq_lines) in enumerate(squads):
        x = Inches(0.4 + i * (sq_w + 0.2))
        add_round(s, x, Inches(1.5), Inches(sq_w), Inches(2.0), light)
        add_round(s, x, Inches(1.5), Inches(sq_w), Inches(0.4), dark)
        add_text(s, x + Inches(0.15), Inches(1.55), Inches(sq_w - 0.3), Inches(0.3),
                 sq_title, font_size=13, bold=True, color=light)
        # lines
        y = 2.0
        for label, content in sq_lines:
            add_text(s, x + Inches(0.15), Inches(y), Inches(sq_w - 0.3), Inches(0.25),
                     label, font_size=9, bold=True, color=dark)
            add_text(s, x + Inches(0.15), Inches(y + 0.22), Inches(sq_w - 0.3), Inches(0.4),
                     content, font_size=9, color=NAVY)
            y += 0.55

    # Bottom row: How they work | PM & Planning
    # Left: How they work
    add_round(s, Inches(0.4), Inches(3.7), Inches(6.2), Inches(3.4), NAVY_2, accent, Pt(2))
    add_round(s, Inches(0.4), Inches(3.7), Inches(6.2), Inches(0.4), accent)
    add_text(s, Inches(0.55), Inches(3.75), Inches(6), Inches(0.3),
             "HOW THE SQUADS WORK · Rituals & Cadences", font_size=12, bold=True, color=WHITE)
    y = 4.2
    for ritual_title, ritual_desc in forums_section:
        add_round(s, Inches(0.55), Inches(y), Inches(5.9), Inches(0.55), NAVY_3)
        add_text(s, Inches(0.65), Inches(y + 0.05), Inches(2), Inches(0.25),
                 ritual_title, font_size=9, bold=True, color=light)
        add_text(s, Inches(2.6), Inches(y + 0.05), Inches(3.7), Inches(0.45),
                 ritual_desc, font_size=9, color=WHITE)
        y += 0.6

    # Right: PM & Planning
    add_round(s, Inches(6.7), Inches(3.7), Inches(6.2), Inches(3.4), NAVY_2, accent, Pt(2))
    add_round(s, Inches(6.7), Inches(3.7), Inches(6.2), Inches(0.4), accent)
    add_text(s, Inches(6.85), Inches(3.75), Inches(6), Inches(0.3),
             "PRODUCT MANAGEMENT & PLANNING", font_size=12, bold=True, color=WHITE)
    add_round(s, Inches(6.85), Inches(4.2), Inches(5.9), Inches(1.0), PURPLE_DARK)
    add_text(s, Inches(7.0), Inches(4.25), Inches(5.6), Inches(0.3),
             "PM STEWARD POOL", font_size=10, bold=True, color=PURPLE_LIGHT)
    add_text(s, Inches(7.0), Inches(4.5), Inches(5.6), Inches(0.7),
             pm_section, font_size=10, color=WHITE)
    add_round(s, Inches(6.85), Inches(5.3), Inches(5.9), Inches(0.95), NAVY_3)
    add_text(s, Inches(7.0), Inches(5.35), Inches(5.6), Inches(0.3),
             "PLANNING APPROACH", font_size=10, bold=True, color=light)
    add_text(s, Inches(7.0), Inches(5.6), Inches(5.6), Inches(0.6),
             planning_section, font_size=9, color=WHITE)
    add_round(s, Inches(6.85), Inches(6.35), Inches(5.9), Inches(0.7), NAVY, ORANGE, Pt(1))
    add_text(s, Inches(7.0), Inches(6.4), Inches(5.6), Inches(0.3),
             "FORUMS REQUIRED (cross-VS · governance)", font_size=9, bold=True, color=ORANGE)
    add_text(s, Inches(7.0), Inches(6.65), Inches(5.6), Inches(0.4),
             "Critical Integrations weekly · Portfolio Committee monthly · BRP quarterly",
             font_size=9, color=WHITE)

    slide_footer(s, num, 13, source)


def slide_03_vs1():
    slide_vs_detail(
        3, "VS-1", "VS-1 · M2C & Customer Service",
        "CS branch · Kumar Jaladanki · ~20 FTE · Three sub-VS formally proposed (Apr 2026)",
        AMBER, AMBER_LIGHT, AMBER_DARK,
        squads=[
            ("SQUAD M2C-SL · Service Lifecycle", [
                ("Lead · Champion", "Vincent Lam (Sr Mgr Device & Metering) · Iram Chowdhury (Smart Metering)"),
                ("Capacity", "8 FTE + 100% vendor-delivered"),
                ("Apps", "SAP Metering · MV90"),
                ("Outcome", "Sign up · activate · changes · disconnects"),
            ]),
            ("SQUAD M2C-RO · Revenue Operations", [
                ("Lead · Champion", "Ivelisse Velazquez (Sr Mgr FICA Billing) · Stephanie (SAP)"),
                ("Capacity", "8 FTE + Accenture occasional"),
                ("Apps", "SAP FICA · ECC · StreamServe"),
                ("Outcome", '"Bill once, bill right, collect on time"'),
            ]),
            ("SQUAD CCC · Customer Care & Contact", [
                ("Lead · Champion", "VACANT 🔴 (Sr Mgr CRM/CCC) · TBD post-hire"),
                ("Capacity", "4 FTE + 3 vendors"),
                ("Apps", "SAP CRM · Medallia · IVR · Blue Prism · chatbot"),
                ("Outcome", "Resolve interactions on first contact"),
            ]),
        ],
        pm_section="Heather LaJoie + Digital CoE Functional SAP PMs · 'first prioritization layer' for Digital CoE; bridge to CS Apps. Tuesday weekly with Vitor.",
        planning_section="Continuous prioritization at sprint/BRP. Pain Score H(x)=(BV·0.35+UX·0.40+Safety·0.25)/Effort (UX-heavy). CoS mix 70/15/5–10/10–15. Run/Grow/Transform default 50/35/15%.",
        forums_section=[
            ("DAILY", "Daily Flow Review · 15 min · per squad · Squad Lead + Tech Lead + vendor PM"),
            ("WEEKLY", "Replenishment · PM steward pulls from Portfolio Board to Squad TODO"),
            ("BI-WEEKLY", "Service Delivery Review (Flow KPIs) · Retrospective / Kaizen"),
            ("MONTHLY", "Portfolio Committee (Gate 2) cross-VS · Customer Journey Forum (CS only)"),
            ("QUARTERLY", "Big Room Planning (BRP) · all 3 VS + P&T PMs + CoE + Travis (PMO)"),
        ],
        source="Source: CS — Proposed Value Streams.pptx (Slides 1–6) · Prioritization Criteria doc · AS-IS Slide 21 · Avangrid_Hierarchy_Overview.md")


def slide_04_vs2():
    slide_vs_detail(
        4, "VS-2", "VS-2 · Field & Network Operations",
        "ITBA branch · Carlos Garbin · ~13 FTE · Two squads supporting field crews and grid operations",
        GREEN, GREEN_LIGHT, GREEN_DARK,
        squads=[
            ("SQUAD G&E · Electric & Gas Operations", [
                ("Lead · Champion", "Bruno Moreto (Sr Mgr E&G Ops Products) · Andrew McGibbon"),
                ("Capacity", "7 FTE — Dan Gauthier (Mgr) · Brijesh Pandey · Jeferson Lima · Maynor Rojas · Vinay Math · vacant"),
                ("Apps", "70–90 apps · Android/tablets · doc viewer · field apps · outage tooling"),
                ("Vendor", "Occasional vendor for ProjectWise/Microstation · Third-Party Risk gate"),
            ]),
            ("SQUAD GIS & OSG", [
                ("Lead · Champion", "George Porto (Sr Mgr GIS & OSG) · George Porto (Champion)"),
                ("Capacity", "6 FTE — Anatoliy · Annette · Gang Li · Roy Yeboah · Steven Birney · Tariq"),
                ("Apps", "SAP HANA · ARCFM · GIS web apps · ProjectWise · Microstation · SWF platform"),
                ("Vendor", "SWF for project work · vendor for major upgrades"),
            ]),
        ],
        pm_section="P&T pool (existing intermediary). Verbatim from interview: 'P&T is the face — helping prioritize the work' (George Porto). Same pool serves VS-3.",
        planning_section="Mixed: continuous (break-fix · Standard CoS) + annual PDD (medium/large GIS projects). CoS mix target 60/20/5/15. Run/Grow/Transform 70/25/5 (break-fix dominant).",
        forums_section=[
            ("DAILY", "Daily Flow Review · 15 min · per squad · break-fix focus"),
            ("WEEKLY", "Replenishment · break-fix prioritized first (Standard FIFO)"),
            ("BI-WEEKLY", "Service Delivery Review · Field Ops stakeholder participates"),
            ("MONTHLY", "Field Ops Sync (G&E ↔ GIS ↔ Smart Metering bridge) · Portfolio Committee"),
            ("QUARTERLY", "BRP aligned with P&S Networks BRP cadence · annual PDD Governance"),
        ],
        source="Source: AS-IS Slides 21–22, 26 · Champion interviews (George Porto, Andrew McGibbon) · Avangrid_Hierarchy_Overview.md")


def slide_05_vs3():
    slide_vs_detail(
        5, "VS-3", "VS-3 · Corporate & Global SAP",
        "ITBA branch · Carlos Garbin · ~13 FTE · Two squads · Local Corporate SAP + Iberdrola Global SAP",
        BLUE, BLUE_LIGHT, BLUE_DARK,
        squads=[
            ("SQUAD CORPORATE APPS", [
                ("Lead · Champion", "Sarath Garikapati (Sr Mgr) · Amit Patel · Stephanie (SAP)"),
                ("Capacity", "8 FTE — Cyntia · Ekta · Gary Wiest · Greeshma · Greg Howard · Vivian · Ramkumar"),
                ("Domains", "Finance · HR · Procurement · Legal · Regulatory · Corporate SAP"),
                ("Vendor", "SWF (software factory) · SAP corporate vendor · PDD-driven"),
            ]),
            ("SQUAD GLOBAL SAP", [
                ("Lead · Champion", "Vanessa Colombo (Mgr Global SAP) · Michael Dubois"),
                ("Capacity", "5 FTE — K. Langlois · James Pellegrino · K. Vaughn · Timoth Lugano"),
                ("Scope", "Iberdrola corporate (Spain) · FI · PO · Budget control"),
                ("Vendor", "Accenture (significant) · SWF pipeline · Thursday Accenture Review (existing)"),
            ]),
        ],
        pm_section="P&T pool (existing intermediary) + vendor PMs. 'P&T serves as intermediary — Global SAP manages Jira' (Michael Dubois). Accenture vendor PMs in replenishment.",
        planning_section="Regulatory items always Fixed-Date CoS. PDD-driven for Corporate large projects. CoS mix 50/25/10/15 (regulatory-heavy). Goal: shift Global SAP from 80 Run to more Grow/Transform via Expedite policy.",
        forums_section=[
            ("DAILY", "Daily Flow Review · Squad Lead + Tech Lead + Accenture lead (Global SAP)"),
            ("WEEKLY", "Replenishment + Thursday Accenture Review (KEPT existing)"),
            ("BI-WEEKLY", "Service Delivery Review · Compliance/Regulatory stakeholder participates"),
            ("MONTHLY", "Iberdrola Sync (Spain stakeholders) · Portfolio Committee"),
            ("QUARTERLY · ANNUAL", "BRP + PDD Governance for Corporate large projects"),
        ],
        source="Source: AS-IS Slide 30 · Global SAP & Corporate interviews · Avangrid_Hierarchy_Overview.md")


# ============ SLIDE 06 — Board Architecture (2 levels) ============
def slide_06_architecture():
    s = prs.slides.add_slide(BLANK)
    set_bg(s, NAVY)
    slide_header(s, 6, "BOARD ARCHITECTURE", "Two Board Levels + Native Jira Hierarchy",
                 "Simplified from 3 to 2 boards · Initiatives & Epics on Portfolio · Stories & Tasks on Squad · Enterprise view is read-only dashboard")

    # LEFT: Jira Hierarchy
    add_round(s, Inches(0.4), Inches(1.5), Inches(4.5), Inches(5.5), NAVY_2, ORANGE, Pt(2))
    add_round(s, Inches(0.4), Inches(1.5), Inches(4.5), Inches(0.4), ORANGE)
    add_text(s, Inches(0.55), Inches(1.55), Inches(4), Inches(0.3),
             "JIRA HIERARCHY (native)", font_size=13, bold=True, color=NAVY)

    levels = [
        ("INITIATIVE", "Cross-VS · multi-quarter · strategic outcome", "Vive em: Enterprise Dashboard (read-only) — Vitor",
         "Example: 'Customer self-service adoption +30% by Q4'", PURPLE_DARK),
        ("EPIC", "Single VS · spans 1–3 sprints · 1 outcome", "Vive em: Portfolio Board (per VS) — Kumar / Carlos",
         "Example: 'Rebuild disconnect flow to reduce errors by 50%'", PURPLE),
        ("STORY · TASK · BUG", "Single squad · fits in 1 sprint · independently testable", "Vive em: Squad Board — Squad Lead",
         "Example: 'Add validation on disconnect API endpoint'", BLUE_DARK),
        ("SUB-TASK", "Decomposition within a Story", "Owner: developer · not visible at portfolio level",
         "", GREEN_DARK),
    ]
    y = 2.0
    for label, l1, l2, l3, color in levels:
        add_round(s, Inches(0.6), Inches(y), Inches(4.1), Inches(0.8), color)
        add_text(s, Inches(0.7), Inches(y + 0.05), Inches(3.9), Inches(0.3),
                 label, font_size=12, bold=True, color=WHITE)
        add_text(s, Inches(0.7), Inches(y + 0.30), Inches(3.9), Inches(0.25),
                 l1, font_size=9, bold=True, color=WHITE)
        add_text(s, Inches(0.7), Inches(y + 0.5), Inches(3.9), Inches(0.25),
                 l2, font_size=9, color=WHITE)
        if l3:
            add_text(s, Inches(0.7), Inches(y + 0.65), Inches(3.9), Inches(0.2),
                     l3, font_size=8, italic=True, color=WHITE)
        y += 0.95

    # Standalone cards note
    add_round(s, Inches(0.6), Inches(5.85), Inches(4.1), Inches(1.05), NAVY, YELLOW, Pt(1))
    add_text(s, Inches(0.7), Inches(5.9), Inches(3.9), Inches(0.3),
             "STANDALONE CARDS (no Epic parent)", font_size=10, bold=True, color=YELLOW)
    add_multiline(s, Inches(0.7), Inches(6.15), Inches(3.9), Inches(0.7), [
        "• Production incidents (live outage)",
        "• Expedite items (P1/P2 customer impact)",
        "• Small ad-hoc fixes < 2 days effort",
        "• Regulatory items deadline < 30 days",
    ], font_size=8, color=WHITE)

    # RIGHT: Two Board Levels
    add_round(s, Inches(5.1), Inches(1.5), Inches(7.85), Inches(5.5), NAVY_2, GREY_LIGHT, Pt(2))
    add_round(s, Inches(5.1), Inches(1.5), Inches(7.85), Inches(0.4), GREY_LIGHT)
    add_text(s, Inches(5.25), Inches(1.55), Inches(7.5), Inches(0.3),
             "TWO BOARD LEVELS", font_size=13, bold=True, color=NAVY)

    # Level 0: Enterprise Dashboard
    add_round(s, Inches(5.3), Inches(2.0), Inches(7.45), Inches(0.85), NAVY, PURPLE, Pt(1))
    add_text(s, Inches(5.45), Inches(2.05), Inches(7.2), Inches(0.3),
             "LEVEL 0  ·  ENTERPRISE DASHBOARD  ·  read-only · NOT a board", font_size=9, bold=True, color=PURPLE_LIGHT)
    add_text(s, Inches(5.45), Inches(2.3), Inches(7.2), Inches(0.3),
             "Cross-VS roll-up via Jira filters & advanced roadmaps · Owner: Vitor Cruz",
             font_size=10, bold=True, color=WHITE)
    add_text(s, Inches(5.45), Inches(2.55), Inches(7.2), Inches(0.3),
             "Initiative roadmap · cross-VS dependency map · capacity heat-map · Pain Score top-20",
             font_size=8, italic=True, color=GREY_LIGHT)

    # Level 1: Portfolio Boards
    add_round(s, Inches(5.3), Inches(3.0), Inches(7.45), Inches(1.7), PURPLE_DARK)
    add_text(s, Inches(5.45), Inches(3.05), Inches(7.2), Inches(0.3),
             "LEVEL 1  ·  PORTFOLIO BOARDS  ·  3 boards · Epics only", font_size=9, bold=True, color=PURPLE_LIGHT)
    boards = [
        ("CS Portfolio", "Kumar · Heather + CoE", "→ M2C-SL · M2C-RO · CCC", AMBER, AMBER_DARK),
        ("ITBA-Field", "Carlos · P&T pool", "→ G&E · GIS", GREEN, WHITE),
        ("ITBA-Corporate", "Carlos · P&T pool", "→ Corporate · Global SAP", BLUE, WHITE),
    ]
    for i, (title, owner, routes, fill, txt) in enumerate(boards):
        x = Inches(5.45 + i * 2.45)
        add_round(s, x, Inches(3.4), Inches(2.35), Inches(1.2), fill)
        add_text(s, x + Inches(0.1), Inches(3.45), Inches(2.15), Inches(0.3),
                 title, font_size=10, bold=True, color=txt)
        add_text(s, x + Inches(0.1), Inches(3.7), Inches(2.15), Inches(0.4),
                 owner, font_size=8, color=txt)
        add_text(s, x + Inches(0.1), Inches(4.15), Inches(2.15), Inches(0.4),
                 routes, font_size=8, italic=True, color=txt)

    # Level 2: Squad Boards
    add_round(s, Inches(5.3), Inches(4.85), Inches(7.45), Inches(2.0), NAVY_3)
    add_text(s, Inches(5.45), Inches(4.9), Inches(7.2), Inches(0.3),
             "LEVEL 2  ·  SQUAD BOARDS  ·  7 boards · Stories + Tasks + standalone cards",
             font_size=9, bold=True, color=BLUE_LIGHT)
    add_text(s, Inches(5.45), Inches(5.15), Inches(7.2), Inches(0.3),
             "Each squad has its own board with Class of Service swimlanes",
             font_size=10, color=WHITE)
    add_text(s, Inches(5.45), Inches(5.4), Inches(7.2), Inches(0.3),
             "Vendors visible as card label (Avangrid FTE · Lemur · Accenture · Medallia · IVR · SWF)",
             font_size=8, color=GREY_LIGHT)
    # 7 mini squad boxes
    squads = [("M2C-SL", AMBER), ("M2C-RO", AMBER), ("CCC", AMBER),
              ("G&E", GREEN), ("GIS", GREEN), ("Corporate", BLUE), ("Global SAP", BLUE)]
    for i, (name, color) in enumerate(squads):
        x = Inches(5.45 + (i % 4) * 1.85)
        y2 = Inches(5.7 + (i // 4) * 0.5)
        add_round(s, x, y2, Inches(1.75), Inches(0.4), color)
        add_text(s, x + Inches(0.05), y2 + Inches(0.05), Inches(1.65), Inches(0.3),
                 name, font_size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    # column layout note
    add_text(s, Inches(5.45), Inches(6.65), Inches(7.2), Inches(0.3),
             "All 7 squad boards: TO DO → IN PROGRESS (WIP-limited) → REVIEW / UAT → DONE",
             font_size=9, bold=True, color=ORANGE)

    slide_footer(s, 6, 13, "Source: Atlassian Jira hierarchy native model · Team Topologies 'Stream-aligned' pattern · Renewables DTO Slide 33")


# ============ SLIDE 07 — Intake step-by-step ============
def slide_07_intake_steps():
    s = prs.slides.add_slide(BLANK)
    set_bg(s, NAVY)
    slide_header(s, 7, "INTAKE PROCESS", "Step-by-Step Intake · Size-Conditional Gates",
                 "8 steps from request to squad board · paths differ by Size + Class of Service · Regulatory and Expedite have shortcuts")

    # 8 steps in 2 rows of 4
    steps = [
        ("1", "REQUEST CREATION", "Requestor", "Jira (single front door)", "Status: NEW"),
        ("2", "AUTO-ROUTING", "Jira automation", "VS field triggers route", "Status: ROUTED"),
        ("3", "TRIAGE", "PM steward + Champion", "Daily/weekly · DoR + CoS + Pain Score", "Status: TRIAGED · BOUNCED"),
        ("4", "BACKLOG ENTRY", "PM steward", "Ranked by Pain Score", "Branch by Size"),
        ("5", "GATE 1 · SHARK TANK", "Shark Tank Committee", "Bi-weekly · Build/Transform Large only", "Status: GATE1_PASSED"),
        ("6", "GATE 2 · PORTFOLIO COMMITTEE", "Vitor + Carlos + Kumar + PMs", "Monthly · investment + capacity slot", "Status: READY_FOR_EXEC"),
        ("7", "PULL TO SQUAD", "Squad Lead + PM steward", "Weekly Replenishment", "Status: SQUAD_TODO"),
        ("8", "EXECUTION", "Squad (FTE + vendor)", "TODO → WIP → REVIEW → DONE", "Daily Flow Review"),
    ]
    for i, (num, title, actor, desc, status) in enumerate(steps):
        col = i % 4
        row = i // 4
        x = Inches(0.4 + col * 3.2)
        y = Inches(1.5 + row * 1.4)
        # circle with number
        add_round(s, x, y, Inches(3.0), Inches(1.25), NAVY_2, ORANGE, Pt(1.5))
        add_round(s, x + Inches(0.1), y + Inches(0.1), Inches(0.45), Inches(0.45), ORANGE)
        add_text(s, x + Inches(0.1), y + Inches(0.13), Inches(0.45), Inches(0.4),
                 num, font_size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(s, x + Inches(0.65), y + Inches(0.1), Inches(2.3), Inches(0.3),
                 title, font_size=10, bold=True, color=ORANGE)
        add_text(s, x + Inches(0.15), y + Inches(0.55), Inches(2.8), Inches(0.25),
                 "Actor: " + actor, font_size=8, color=WHITE)
        add_text(s, x + Inches(0.15), y + Inches(0.78), Inches(2.8), Inches(0.25),
                 desc, font_size=8, color=GREY_LIGHT)
        add_text(s, x + Inches(0.15), y + Inches(1.0), Inches(2.8), Inches(0.25),
                 status, font_size=8, bold=True, color=YELLOW)

    # Bottom: paths summary
    add_round(s, Inches(0.4), Inches(4.5), Inches(12.55), Inches(2.5), NAVY_2, ORANGE, Pt(2))
    add_round(s, Inches(0.4), Inches(4.5), Inches(12.55), Inches(0.4), ORANGE)
    add_text(s, Inches(0.55), Inches(4.55), Inches(12), Inches(0.3),
             "PATHS BY SIZE & CLASS OF SERVICE — what to skip, what to require",
             font_size=12, bold=True, color=NAVY)

    paths = [
        ("PATH A · EXPEDITE", "P1/P2 customer impact · prod outage · Reg < 30 days",
         "1 → 2 → 3 → SKIP Gate 1 → SKIP Gate 2 → 7 → 8", "Capacity reserve 5–10% per quarter · WIP limit 1", PINK),
        ("PATH B · REGULATORY", "Regulatory flag = TRUE · CoS = Fixed-Date",
         "1 → 2 → 3 → SKIP Gate 1 → 6 → 7 → 8 (no Go/No-Go)", "Sponsor must be Compliance / Legal", YELLOW),
        ("PATH C · SMALL", "Size = Small · CoS = Standard · effort < 5 days",
         "1 → 2 → 3 → 4 → SKIP Gate 1 → SKIP Gate 2 if < 2d → 7 → 8", "Quick-fix · duty engineer · production errors", GREEN),
        ("PATH D · BUILD/TRANSFORM", "Size = Medium/Large · Work Type = Build/Transform",
         "1 → 2 → 3 → 4 → 5 (Gate 1 Shark Tank) → 6 (Gate 2) → 7 → 8", "Full path · all artifacts required · Epic created", PURPLE),
    ]
    for i, (title, condition, flow, note, color) in enumerate(paths):
        col = i % 2
        row = i // 2
        x = Inches(0.55 + col * 6.1)
        y = Inches(5.0 + row * 0.9)
        add_round(s, x, y, Inches(6.0), Inches(0.85), NAVY, color, Pt(1.5))
        add_text(s, x + Inches(0.1), y + Inches(0.05), Inches(5.8), Inches(0.25),
                 title, font_size=10, bold=True, color=color)
        add_text(s, x + Inches(0.1), y + Inches(0.25), Inches(5.8), Inches(0.2),
                 condition, font_size=8, italic=True, color=GREY_LIGHT)
        add_text(s, x + Inches(0.1), y + Inches(0.43), Inches(5.8), Inches(0.2),
                 flow, font_size=8, color=WHITE)
        add_text(s, x + Inches(0.1), y + Inches(0.62), Inches(5.8), Inches(0.2),
                 note, font_size=7, italic=True, color=GREY_LIGHT)

    slide_footer(s, 7, 13, "Source: Renewables DTO Slide 33 (Submit Jira Request Ticket) · Slide 28 (End-to-End Kanban Workflow) · CS Proposed VS Slide 3")


# ============ SLIDE 08 — Jira Workflow ============
def slide_08_jira():
    s = prs.slides.add_slide(BLANK)
    set_bg(s, NAVY)
    slide_header(s, 8, "JIRA WORKFLOW", "Jira Statuses · Transitions · Automations",
                 "Single workflow for all 7 squads · auto-routing by Value Stream · auto-CoS by Customer Impact Tier · auto-Pain Score")

    # Status flow
    add_round(s, Inches(0.4), Inches(1.5), Inches(12.55), Inches(1.4), NAVY_2, BLUE, Pt(2))
    add_round(s, Inches(0.4), Inches(1.5), Inches(12.55), Inches(0.4), BLUE)
    add_text(s, Inches(0.55), Inches(1.55), Inches(12), Inches(0.3),
             "STATUS FLOW · 9 statuses · valid transitions", font_size=12, bold=True, color=WHITE)
    statuses = [("NEW", GREY_LIGHT, NAVY), ("ROUTED", YELLOW, NAVY), ("TRIAGED", ORANGE_DARK, WHITE),
                ("BACKLOG", JIRA_BLUE, WHITE), ("GATE1_PASSED", PURPLE, WHITE),
                ("READY_FOR_EXEC", SUCCESS, WHITE), ("IN_PROGRESS", GREEN, WHITE),
                ("REVIEW", GREEN_DARK, WHITE), ("DONE", NAVY, SUCCESS)]
    for i, (name, fill, txt) in enumerate(statuses):
        x = Inches(0.55 + i * 1.38)
        add_round(s, x, Inches(2.05), Inches(1.3), Inches(0.45), fill)
        add_text(s, x, Inches(2.05), Inches(1.3), Inches(0.45),
                 name, font_size=8, bold=True, color=txt, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # arrow
        if i < len(statuses) - 1:
            arrow = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, x + Inches(1.30), Inches(2.18), Inches(0.08), Inches(0.18))
            arrow.fill.solid()
            arrow.fill.fore_color.rgb = BLUE
            arrow.line.fill.background()

    # side transitions
    add_text(s, Inches(0.55), Inches(2.55), Inches(12), Inches(0.3),
             "SIDE TRANSITIONS", font_size=9, bold=True, color=YELLOW)
    add_multiline(s, Inches(0.55), Inches(2.75), Inches(12), Inches(0.4), [
        "BOUNCED (TRIAGED → if DoR fails)  ·  REJECTED (BACKLOG/GATE1 → if Shark Tank No-Go)  ·  BLOCKED (IN_PROGRESS → dependency)  ·  EXPEDITE_LANE (TRIAGED → IN_PROGRESS, skip backlog/gates)",
    ], font_size=9, color=WHITE)

    # Automations
    add_round(s, Inches(0.4), Inches(3.1), Inches(8.0), Inches(3.9), NAVY_2, GREEN, Pt(2))
    add_round(s, Inches(0.4), Inches(3.1), Inches(8.0), Inches(0.4), GREEN)
    add_text(s, Inches(0.55), Inches(3.15), Inches(7.7), Inches(0.3),
             "AUTOMATIONS · 'if-this-then-that' rules", font_size=12, bold=True, color=WHITE)
    automations = [
        ("A1 · AUTO-ROUTING", GREEN, "Trigger: Value Stream field selected", "Action: ROUTED · assign to portfolio board · @notify Champion"),
        ("A2 · AUTO-EXPEDITE", PINK, "Trigger: Customer Impact Tier P1/P2", "Action: Force CoS = Expedite · skip backlog · WIP-1"),
        ("A3 · AUTO-FIXED-DATE (Regulatory)", YELLOW, "Trigger: Regulatory flag = TRUE", "Action: CoS = Fixed-Date · Sponsor = Compliance · skip Gate 1"),
        ("A4 · AUTO-PAIN-SCORE", PURPLE, "Trigger: Status → TRIAGED", "Action: H(x)=(BV·w₁+UX·w₂+Safety·w₃)/Effort · sort backlog"),
        ("A5 · AUTO-GATE-SKIP", SUCCESS, "Trigger: Size + Work Type at TRIAGED", "Action: Skip Gate 1 (Build only) · Skip Gate 2 if Size=Small + < 2d"),
        ("A6 · AGING ALERT", GREY_LIGHT, "Trigger: Card aging > CoS threshold", "Action: Slack/email Squad Lead + PM steward"),
    ]
    y = 3.55
    for title, color, trigger, action in automations:
        add_round(s, Inches(0.55), Inches(y), Inches(7.7), Inches(0.5), NAVY, color, Pt(1))
        add_text(s, Inches(0.65), Inches(y + 0.03), Inches(2.5), Inches(0.25),
                 title, font_size=9, bold=True, color=color)
        add_text(s, Inches(3.2), Inches(y + 0.03), Inches(5), Inches(0.25),
                 trigger, font_size=8, color=WHITE)
        add_text(s, Inches(3.2), Inches(y + 0.25), Inches(5), Inches(0.25),
                 action, font_size=8, italic=True, color=GREY_LIGHT)
        y += 0.55

    # Right panel: Required customization
    add_round(s, Inches(8.6), Inches(3.1), Inches(4.35), Inches(3.9), NAVY_2, AMBER, Pt(2))
    add_round(s, Inches(8.6), Inches(3.1), Inches(4.35), Inches(0.4), AMBER)
    add_text(s, Inches(8.75), Inches(3.15), Inches(4.05), Inches(0.3),
             "REQUIRED JIRA CUSTOMIZATION", font_size=12, bold=True, color=AMBER_DARK)
    add_text(s, Inches(8.75), Inches(3.6), Inches(4.05), Inches(0.3),
             "CUSTOM FIELDS (new)", font_size=9, bold=True, color=AMBER_LIGHT)
    fields = [
        "• Value Stream / Domain (single-select)",
        "• Class of Service (4 options)",
        "• Customer Impact Tier (P1–P4 · CS only)",
        "• Work Type (BAU / Build / Upgrade)",
        "• Run/Grow/Transform",
        "• OPCO (UIL / CMP / NY)",
        "• Sponsor (user picker)",
        "• Regulatory flag (checkbox + sub-fields)",
        "• Pain Score components (BV/UX/Safety/Effort)",
        "• Pain Score (computed · read-only)",
        "• Champion (user · auto-by-VS)",
        "• Vendor (multi-select)",
    ]
    add_multiline(s, Inches(8.75), Inches(3.85), Inches(4.05), Inches(2.5),
                  fields, font_size=8, color=WHITE)
    add_text(s, Inches(8.75), Inches(6.4), Inches(4.05), Inches(0.3),
             "PROJECT & SCHEMES", font_size=9, bold=True, color=AMBER_LIGHT)
    add_multiline(s, Inches(8.75), Inches(6.6), Inches(4.05), Inches(0.5), [
        "• Single Jira project: DGTL-PROD",
        "• 1 Issue Type: Demand (Initiative · Epic · Story · Sub-task · Bug)",
        "• 3 portfolio boards · 7 squad boards (JQL filters)",
    ], font_size=8, color=WHITE)

    slide_footer(s, 8, 13, "Source: Renewables DTO Slide 33 · CS Proposed Value Streams Slide 3 · Pain Score formula from Prioritization Criteria doc")


# ============ SLIDE 09 — Field Detail ============
def slide_09_fields():
    s = prs.slides.add_slide(BLANK)
    set_bg(s, NAVY)
    slide_header(s, 9, "INTAKE FIELD DETAIL", "Mandatory Fields · Validation Rules · Options",
                 "Single intake template · all fields validated at Triage · Regulatory flag activates Path B (skip Gate 1)")

    # Table
    table_x = Inches(0.4)
    table_y = Inches(1.5)
    table_w = Inches(12.55)
    headers = [("FIELD", 2.0), ("REQUIRED?", 1.0), ("TYPE / OPTIONS", 4.0), ("VALIDATION RULE", 3.5), ("DRIVES", 2.05)]
    # header row
    add_rect(s, table_x, table_y, table_w, Inches(0.4), JIRA_BLUE)
    cx = 0.4
    for label, w_in in headers:
        add_text(s, Inches(cx), table_y + Inches(0.07), Inches(w_in), Inches(0.3),
                 label, font_size=10, bold=True, color=WHITE)
        cx += w_in

    rows = [
        ("Title", "YES", "Free text · max 120 chars", "Outcome-based — verb + measurable result", "Card display label"),
        ("Description", "YES", "Rich text · min 100 chars", "Describes the problem, not the solution", "Triage decision input"),
        ("Business Value", "YES", "Rich text + numeric BV (1–5)", "Quantify outcome — revenue · cost · NPS · SLA · safety", "Pain Score input"),
        ("Acceptance Criteria", "YES", "Rich text · checklist", "Measurable, verifiable conditions for done", "DoD reference"),
        ("Value Stream / Domain ⭐", "YES", "Single-select · 7 options", "Auto-routes ticket to portfolio board (A1)", "Auto-Routing"),
        ("Type", "YES", "Demand · Incident · Change", "Incidents bypass Gates · go directly to squad", "Path branching"),
        ("Size", "YES", "Small · Medium · Large", "Small + < 2d effort skips Gate 2 (Path C)", "Gate skip rules"),
        ("Class of Service ⭐", "YES", "Standard · Fixed-Date · Expedite · Intangible", "Auto-Expedite if Customer Impact Tier ∈ {P1, P2}", "Squad swimlane"),
        ("Work Type", "YES", "BAU · Build · Upgrade", "BAU bypasses Gate 1 · Build/Transform require Gate 1", "Gate 1 trigger"),
        ("Run / Grow / Transform", "YES", "Run · Grow · Transform", "Default mix per VS (~50/35/15%)", "Strategic mix"),
        ("Customer Impact Tier ⭐", "CS only", "P1 / P2 / P3 / P4", "P1/P2 → automatic Class of Service = Expedite", "Auto-Expedite"),
        ("Regulatory flag ⭐", "Conditional", "Checkbox + Regulator + Deadline", "Auto-Fixed-Date · Sponsor must be Compliance/Legal", "Auto-Fixed-Date"),
        ("OPCO", "YES", "UIL · CMP · NY", "Determines local stakeholder routing", "Reporting filter"),
        ("Sponsor", "YES", "User picker", "Business owner accountable for outcome", "Gate 1 attendance"),
        ("Requested Due Date", "Cond.", "Date picker (YYYY-MM-DD)", "Mandatory if CoS = Fixed-Date · < 30 days → Auto-Expedite", "SLA / aging"),
        ("Pain Score (computed)", "READ-ONLY", "H(x)=(BV·w₁+UX·w₂+Safety·w₃)/Effort", "Auto-computed at TRIAGED status", "Backlog ranking"),
    ]
    row_y = table_y + Inches(0.4)
    for i, row in enumerate(rows):
        bg = NAVY_2 if i % 2 == 0 else NAVY
        add_rect(s, table_x, row_y, table_w, Inches(0.32), bg)
        cx = 0.4
        for j, val in enumerate(row):
            color = WHITE
            bold = j == 0
            if j == 1 and val == "YES":
                color = RED
                bold = True
            elif j == 1 and val in ["READ-ONLY"]:
                color = GREY_LIGHT
                bold = True
            elif j == 1:
                color = YELLOW
                bold = True
            add_text(s, Inches(cx + 0.05), row_y + Inches(0.03), Inches(headers[j][1] - 0.05), Inches(0.28),
                     val, font_size=8, bold=bold, color=color)
            cx += headers[j][1]
        row_y += Inches(0.32)

    slide_footer(s, 9, 13, "Source: CS Proposed Value Streams Slide 3 · Renewables DTO Slide 33 · Pain Score from Prioritization Criteria doc · Pain Point Evidence Matrix Slide 41")


# ============ SLIDE 10 — Portfolio Board ============
def slide_10_portfolio():
    s = prs.slides.add_slide(BLANK)
    set_bg(s, NAVY)
    slide_header(s, 10, "PORTFOLIO BOARD LAYOUT",
                 "Portfolio Board · Epics only · Example: CS Portfolio Board",
                 "Same column structure on all 3 portfolio boards · cards are Initiatives or Epics (not Stories) · Pain Score drives ranking")

    # Board
    add_rect(s, Inches(0.4), Inches(1.5), Inches(12.55), Inches(0.4), AMBER)
    add_text(s, Inches(0.55), Inches(1.55), Inches(12), Inches(0.3),
             "CS Portfolio Board · VS-1 · Owner: Kumar Jaladanki · PM Steward: Heather LaJoie + Digital CoE",
             font_size=11, bold=True, color=AMBER_DARK)
    # Columns
    cols = [("INTAKE", JIRA_BLUE), ("TRIAGE", ORANGE_DARK), ("BACKLOG · ranked Pain Score", PURPLE),
            ("READY FOR EXEC", SUCCESS), ("IN FLIGHT (squads)", GREEN_DARK)]
    col_w = 12.55 / 5
    for i, (label, color) in enumerate(cols):
        x = Inches(0.4 + i * col_w)
        add_rect(s, x, Inches(1.9), Inches(col_w), Inches(0.3), color)
        add_text(s, x, Inches(1.93), Inches(col_w), Inches(0.25),
                 label, font_size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        # column body
        add_rect(s, x, Inches(2.2), Inches(col_w), Inches(4.5), WHITE, GREY_LIGHT, Pt(0.5))

    # Sample cards
    sample_cards = [
        # (column index, title, subtitle, cos_color, badges)
        (0, "EPIC · DGTL-PROD-1247", "Disconnect flow rebuild — reduce errors 50%", "Sub-VS: M2C-SL · Size: L · Type: Build"),
        (0, "EPIC · DGTL-PROD-1248", "Medallia survey upgrade Q3", "Sub-VS: CCC · Size: M · Vendor: Medallia"),
        (1, "EPIC · DGTL-PROD-1241  🅡", "Rate-case tariff banner update", "M2C-RO · Fixed-Date · deadline 2026-08-15"),
        (1, "EPIC · DGTL-PROD-1242", "CRM lookup latency fix", "CCC · Size: Small · Standard"),
        (2, "#1 · DGTL-PROD-1208  🅔", "P1 outage: billing engine down", "CoS: Expedite · P1 · Pain 4.9"),
        (2, "#2 · DGTL-PROD-1192  🅡", "Tax rate update Q3 2026", "M2C-RO · Fixed-Date · Pain 4.5"),
        (2, "#3 · DGTL-PROD-1185", "Self-service portal redesign", "CCC · Large · Pre-Gate-1 · Pain 4.1"),
        (3, "EPIC · DGTL-PROD-1175  🅡", "SOX compliance Q2 release", "M2C-RO · Slot: Squad M2C-RO"),
        (3, "EPIC · DGTL-PROD-1170", "Smart meter rollout NY phase 2", "M2C-SL · vendor delivery"),
        (4, "EPIC · DGTL-PROD-1158", "Bill-once initiative", "Squad M2C-RO · Sprint 14/26"),
        (4, "EPIC · DGTL-PROD-1152", "IVR enhancement v2", "Squad CCC · vendor IVR lead"),
    ]
    cards_per_col = {0: 0, 1: 0, 2: 0, 3: 0, 4: 0}
    for col_idx, key, title, meta in sample_cards:
        x = Inches(0.45 + col_idx * col_w)
        y = Inches(2.3 + cards_per_col[col_idx] * 0.95)
        cards_per_col[col_idx] += 1
        # card
        cw = Inches(col_w - 0.1)
        add_rect(s, x, y, cw, Inches(0.85), WHITE, GREY_LIGHT, Pt(0.5))
        # left stripe — color by column type
        stripe = [JIRA_BLUE, ORANGE_DARK, PURPLE, SUCCESS, GREEN_DARK][col_idx]
        add_rect(s, x, y, Inches(0.05), Inches(0.85), stripe)
        add_text(s, x + Inches(0.1), y + Inches(0.02), cw - Inches(0.1), Inches(0.25),
                 key, font_size=7, bold=True, color=stripe)
        add_text(s, x + Inches(0.1), y + Inches(0.22), cw - Inches(0.1), Inches(0.35),
                 title, font_size=9, bold=True, color=NAVY)
        add_text(s, x + Inches(0.1), y + Inches(0.55), cw - Inches(0.1), Inches(0.3),
                 meta, font_size=7, color=GREY_DARK)

    # Bottom note
    add_round(s, Inches(0.4), Inches(6.85), Inches(12.55), Inches(0.6), NAVY_2, AMBER, Pt(1))
    add_text(s, Inches(0.55), Inches(6.9), Inches(12), Inches(0.5), [
    ][0] if False else "SAME LAYOUT applies to ITBA-Field and ITBA-Corporate Portfolio Boards · cards are Epics or Initiatives only · Stories live on Squad Boards · Forums: Replenishment weekly · Customer Journey Forum monthly (CS only) · Iberdrola Sync (Global SAP) · Portfolio Committee monthly · BRP quarterly",
             font_size=9, color=WHITE)

    slide_footer(s, 10, 13, "Source: Renewables DTO Slide 33 · Customer Services Proposed Value Streams Slide 3 · Pain Score formula")


# ============ SLIDE 11 — Squad Board ============
def slide_11_squad():
    s = prs.slides.add_slide(BLANK)
    set_bg(s, NAVY)
    slide_header(s, 11, "SQUAD BOARD LAYOUT",
                 "Squad Board · Stories & Tasks · Class of Service Swimlanes · Vendor Visibility",
                 "Example: Squad CCC · vendors and FTE on the same board · duty rotation absorbs Expedite + Standard small items")

    # Header
    add_rect(s, Inches(0.4), Inches(1.5), Inches(12.55), Inches(0.4), AMBER_DARK)
    add_text(s, Inches(0.55), Inches(1.55), Inches(12), Inches(0.3),
             "Squad CCC · Customer Care & Contact Center · 4 FTE + 3 vendor PMs · Lead: VACANT 🔴 · PM: Heather LaJoie",
             font_size=11, bold=True, color=AMBER_LIGHT)

    # Column headers
    headers = [("CoS · WIP", NAVY, 1.2), ("TO DO", NAVY_3, 2.85), ("IN PROGRESS", JIRA_BLUE, 2.85),
               ("REVIEW / UAT", PURPLE, 2.85), ("DONE", GREEN_DARK, 2.8)]
    cx = 0.4
    for label, color, w_in in headers:
        add_rect(s, Inches(cx), Inches(1.9), Inches(w_in), Inches(0.3), color)
        add_text(s, Inches(cx), Inches(1.93), Inches(w_in), Inches(0.25),
                 label, font_size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        cx += w_in

    # Swimlanes
    swimlanes = [
        ("EXPEDITE", "P1/P2 · prod\nWIP 1 · 4h alert", PINK, 1.0),
        ("FIXED-DATE", "Regulatory\nWIP 2 · 50% timeline", YELLOW, 1.0),
        ("STANDARD", "FIFO · Pain Score\nWIP 4 · 14 days", NAVY_3, 1.4),
        ("INTANGIBLE", "Tech debt\nWIP 1 · 90 days", GREEN_LIGHT, 0.85),
    ]
    sy = 2.2
    for label, sub, color, h in swimlanes:
        add_rect(s, Inches(0.4), Inches(sy), Inches(1.2), Inches(h), color)
        add_text(s, Inches(0.45), Inches(sy + 0.1), Inches(1.1), Inches(0.3),
                 label, font_size=10, bold=True, color=NAVY if color in [YELLOW, GREEN_LIGHT] else WHITE,
                 align=PP_ALIGN.CENTER)
        add_text(s, Inches(0.45), Inches(sy + 0.4), Inches(1.1), Inches(h - 0.4),
                 sub, font_size=8, color=NAVY if color in [YELLOW, GREEN_LIGHT] else WHITE,
                 align=PP_ALIGN.CENTER)
        # column dividers in this row
        cx2 = 1.6
        for w_in in [2.85, 2.85, 2.85, 2.8]:
            add_rect(s, Inches(cx2), Inches(sy), Inches(w_in), Inches(h), WHITE, GREY_LIGHT, Pt(0.5))
            cx2 += w_in
        sy += h

    # Sample cards (a few key ones)
    cards = [
        # (swim_y_start, col_x, title, meta, vendor_chip_color, vendor_label)
        (2.3, 1.7, "CCC-1340 · STORY", "P1: IVR returning 503\nSLA: 2h", PINK, "IVR"),
        (2.3, 4.55, "CCC-1338 · STORY", "P2: chat router latency\nAging: 3h", NAVY, "FTE"),
        (3.3, 1.7, "CCC-1320 · STORY 🅡", "Tariff banner: Q3 update\nDue: 2026-08-15", NAVY, "FTE"),
        (3.3, 4.55, "CCC-1325 · STORY", "Medallia upgrade Q3\nVendor lead", GREEN, "Medallia"),
        (4.3, 1.7, "CCC-1300 · STORY", "Disconnect flow ph.1\nEpic: 1247", NAVY, "FTE"),
        (4.3, 4.55, "CCC-1295 · STORY", "CRM dedup batch", PURPLE, "Lemur"),
        (5.7, 1.7, "CCC-1290 · TASK", "CRM tech debt cleanup\n10% reserve", NAVY, "FTE"),
        (5.7, 11.7, "CCC-1255 ✓", "Chatbot router v2 release", NAVY, "FTE"),
    ]
    for y, cx, title, meta, chip_color, chip_label in cards:
        add_rect(s, Inches(cx), Inches(y + 0.05), Inches(2.7), Inches(0.85), WHITE, GREY_LIGHT, Pt(0.5))
        add_text(s, Inches(cx + 0.1), Inches(y + 0.1), Inches(2.5), Inches(0.25),
                 title, font_size=8, bold=True, color=NAVY)
        add_text(s, Inches(cx + 0.1), Inches(y + 0.3), Inches(2.5), Inches(0.4),
                 meta, font_size=7, color=GREY_DARK)
        add_round(s, Inches(cx + 0.1), Inches(y + 0.7), Inches(0.7), Inches(0.18), chip_color)
        add_text(s, Inches(cx + 0.1), Inches(y + 0.7), Inches(0.7), Inches(0.18),
                 chip_label, font_size=7, bold=True, color=WHITE if chip_color != GREEN_LIGHT else NAVY,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Footer note
    add_round(s, Inches(0.4), Inches(6.5), Inches(12.55), Inches(0.95), NAVY_2, GREEN, Pt(1.5))
    add_round(s, Inches(0.4), Inches(6.5), Inches(12.55), Inches(0.3), GREEN)
    add_text(s, Inches(0.55), Inches(6.53), Inches(12), Inches(0.25),
             "DUTY ROTATION + VENDOR LEGEND", font_size=10, bold=True, color=WHITE)
    add_multiline(s, Inches(0.55), Inches(6.85), Inches(12), Inches(0.6), [
        ("Duty engineer (rotates weekly): handles Expedite swimlane · Standard small items · production errors · 1 person from 4 FTE on duty per week",),
        ("Rest of squad: protected for Build/Project work · Standard medium/large items · Fixed-Date deliverables",),
        ("Vendor visibility: color-coded chip on each card — FTE (navy) · Lemur (purple) · Medallia (green) · IVR (pink) · Accenture (yellow) · SWF (grey)",),
    ], font_size=8, color=WHITE)

    slide_footer(s, 11, 13, "Source: Class of Service framework from Renewables DTO · Pain Score from Prioritization Criteria doc · Vendor list from CS Proposed VS Slide 6")


# ============ SLIDE 12 — CoS in Practice ============
def slide_12_cos():
    s = prs.slides.add_slide(BLANK)
    set_bg(s, NAVY)
    slide_header(s, 12, "CLASS OF SERVICE IN PRACTICE",
                 "Four Service Classes · Five Paths · One Backlog",
                 "How Standard · Fixed-Date · Expedite · Intangible flow through the same boards with different policies and capacity reserves")

    cs_data = [
        # (header, header_fill, header_text, body_lines)
        ("STANDARD", BLUE_DARK, WHITE,
         "Default flow · FIFO · Pain Score ranked",
         [("WHEN TO USE", "Bulk of demand · normal scheduling · no external deadline · no customer-facing severity"),
          ("CAPACITY", "~60–70% of squad capacity (VS-1 70%, VS-2 60%, VS-3 50%)"),
          ("PATH", "Full intake (1→2→3→4) → Gate 2 → 7 → 8\nSkip Gate 1 if Size=Small or BAU\nSkip Gate 2 if Size=Small + < 2d"),
          ("WIP / AGING", "WIP 3–4 per squad · 14 days aging alert"),
          ("EXAMPLES", "Disconnect flow rebuild · CRM dedup · ARCFM enhancement · HR config")]),
        ("FIXED-DATE", YELLOW, NAVY,
         "Deadline-locked · scheduled backwards",
         [("WHEN TO USE", "External deadline (regulator · contract · vendor lock) · late delivery has compliance/financial penalty"),
          ("CAPACITY", "~15–25% (highest in VS-3 Corporate · regulatory-heavy 25%)"),
          ("PATH", "1→2→3→4 → Skip Gate 1 → Gate 2 → 7 → 8\nAuto if Regulatory flag = TRUE\nSponsor must be Compliance/Legal"),
          ("WIP / AGING", "WIP 2 per squad · 50% remaining timeline alert"),
          ("EXAMPLES", "PUC tariff banner · SOX audit log · Iberdrola Q-close · CT/ME/NY rollouts")]),
        ("EXPEDITE", PINK, WHITE,
         "Jump-the-queue · 1 in flight · capacity reserved",
         [("WHEN TO USE", "P1/P2 customer impact · production outage · Regulatory deadline imminent (< 30 days)"),
          ("CAPACITY", "~5–10% capacity reserve · blocked-out per quarter (flows to Standard if empty)"),
          ("PATH", "1→2→3 → Skip Gates → 7 → 8\nAuto-trigger from Customer Impact Tier P1/P2\nDoR relaxed (best-guess acceptable)"),
          ("WIP / AGING", "WIP 1 per squad · 4 hours alert · auto-escalate to Squad Lead + PM"),
          ("EXAMPLES", "Billing engine down · IVR 503 · Iberdrola month-close blocker · self-service portal outage")]),
        ("INTANGIBLE", GREEN_LIGHT, NAVY,
         "Tech debt · refactor · enablers",
         [("WHEN TO USE", "No direct user-visible value · long-term health · refactor · platform enablers · documentation"),
          ("CAPACITY", "~10–15% capacity reserved per quarter (verbatim CS Slide 3)"),
          ("PATH", "1→2→3→4 → Skip Gate 1 → Gate 2 → 7 → 8\nSame as Standard for gates"),
          ("WIP / AGING", "WIP 1 per squad · 90 days alert (low-priority · not escalated)"),
          ("EXAMPLES", "CRM tech debt · SAP Basis kernel patches · pipeline modernization · doc refresh")]),
    ]
    col_w = (13.0 - 0.4 - 0.4 - 3 * 0.1) / 4  # ~ 3.075
    for i, (label, hfill, htxt, subtitle, body) in enumerate(cs_data):
        x = Inches(0.4 + i * (col_w + 0.1))
        add_round(s, x, Inches(1.5), Inches(col_w), Inches(5.5), NAVY_2, hfill, Pt(2))
        add_round(s, x, Inches(1.5), Inches(col_w), Inches(0.6), hfill)
        add_text(s, x + Inches(0.1), Inches(1.55), Inches(col_w - 0.2), Inches(0.3),
                 label, font_size=14, bold=True, color=htxt)
        add_text(s, x + Inches(0.1), Inches(1.8), Inches(col_w - 0.2), Inches(0.3),
                 subtitle, font_size=9, color=htxt)
        y = 2.2
        for section_label, section_text in body:
            add_text(s, x + Inches(0.1), Inches(y), Inches(col_w - 0.2), Inches(0.25),
                     section_label, font_size=8, bold=True, color=GREY_LIGHT)
            add_text(s, x + Inches(0.1), Inches(y + 0.22), Inches(col_w - 0.2), Inches(0.7),
                     section_text, font_size=8, color=WHITE)
            y += 0.95

    slide_footer(s, 12, 13, "Source: Renewables DTO + CS Proposed Value Streams Slide 3 + Pain Point Evidence Matrix · 'If everything is urgent, nothing is urgent' — Digital CoE Slide 105")


# ============ SLIDE 13 — Lifecycle End-to-End ============
def slide_13_lifecycle():
    s = prs.slides.add_slide(BLANK)
    # Light theme to match reference
    set_bg(s, WHITE)

    # Header band (NTT DATA blue)
    NTT = RGBColor(0x3B, 0x6E, 0x8F)
    add_rect(s, 0, 0, SW, Inches(0.7), NTT)
    add_text(s, Inches(0.4), Inches(0.1), Inches(8), Inches(0.4),
             "LIFECYCLE MODEL", font_size=22, bold=True, color=WHITE)
    add_text(s, Inches(0.4), Inches(0.42), Inches(8), Inches(0.3),
             "DEMAND MANAGEMENT — END-TO-END", font_size=12, bold=True, color=WHITE)
    add_text(s, Inches(11.5), Inches(0.15), Inches(1.8), Inches(0.4),
             "NTT DATA", font_size=14, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)

    # ROLES PANEL (left)
    add_rect(s, Inches(0.2), Inches(1.0), Inches(2.0), Inches(0.4), NTT)
    add_text(s, Inches(0.2), Inches(1.05), Inches(2.0), Inches(0.3),
             "Roles", font_size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_rect(s, Inches(0.2), Inches(1.4), Inches(2.0), Inches(4.4), WHITE, NTT, Pt(1))
    roles = [
        ("Business Requester", "Sponsor · OPCO"),
        ("Application Champion", "First-line triage"),
        ("Product Manager", "Heather + CoE / P&T"),
        ("Squad / Tech Lead", "Pulls work · executes"),
        ("IT Program Delivery", "Travis (PMO Lead)"),
    ]
    ry = 1.5
    for title, sub in roles:
        add_round(s, Inches(0.5), Inches(ry), Inches(0.4), Inches(0.4), AMBER_LIGHT)
        add_text(s, Inches(0.95), Inches(ry), Inches(1.15), Inches(0.25),
                 title, font_size=9, bold=True, color=NAVY)
        add_text(s, Inches(0.95), Inches(ry + 0.2), Inches(1.15), Inches(0.25),
                 sub, font_size=7, color=GREY_DARK)
        ry += 0.85

    # PROCESS BOX (center+right)
    add_rect(s, Inches(2.4), Inches(1.0), Inches(10.7), Inches(0.4), GREEN)
    add_text(s, Inches(2.4), Inches(1.05), Inches(10.7), Inches(0.3),
             "Process", font_size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_rect(s, Inches(2.4), Inches(1.4), Inches(10.7), Inches(4.4), WHITE, GREEN, Pt(1.5))

    # Funnel (left part)
    add_text(s, Inches(2.6), Inches(1.5), Inches(2.0), Inches(0.25),
             "Intake of demand from requestors", font_size=8, bold=True, color=NAVY)
    # Funnel as triangle (using ISOCELES_TRIANGLE flipped)
    funnel = s.shapes.add_shape(MSO_SHAPE.TRAPEZOID, Inches(2.7), Inches(1.85), Inches(1.6), Inches(1.0))
    funnel.rotation = 180
    funnel.fill.solid()
    funnel.fill.fore_color.rgb = NTT
    funnel.line.fill.background()
    add_text(s, Inches(2.7), Inches(1.85), Inches(1.6), Inches(1.0),
             "JIRA", font_size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, Inches(3.25), Inches(2.85), Inches(0.5), Inches(0.4), NTT)
    add_text(s, Inches(2.6), Inches(3.4), Inches(2.0), Inches(0.25),
             "Submits via Jira request", font_size=8, italic=True, color=GREY_DARK, align=PP_ALIGN.CENTER)
    add_text(s, Inches(2.6), Inches(3.6), Inches(2.0), Inches(0.25),
             "single front door", font_size=8, italic=True, color=GREY_DARK, align=PP_ALIGN.CENTER)

    # Auto-redirects arrow
    arrow = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(4.4), Inches(2.0), Inches(0.7), Inches(0.5))
    arrow.fill.solid()
    arrow.fill.fore_color.rgb = GREEN
    arrow.line.fill.background()
    add_text(s, Inches(4.4), Inches(2.55), Inches(0.7), Inches(0.2),
             "auto", font_size=7, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    add_text(s, Inches(4.4), Inches(2.7), Inches(0.7), Inches(0.2),
             "redirects", font_size=7, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

    # 3 Portfolio Boards
    add_text(s, Inches(5.2), Inches(1.5), Inches(7.8), Inches(0.25),
             "3 Portfolio Kanban Boards", font_size=8, bold=True, color=NAVY)
    portfolios = [
        ("CUSTOMER SERVICES PORTFOLIO KANBAN", "Owner: Kumar · PM: Heather + CoE", "→ M2C-SL · M2C-RO · CCC", AMBER, AMBER_DARK),
        ("ITBA-FIELD PORTFOLIO KANBAN", "Owner: Carlos · PM: P&T pool", "→ G&E · GIS", GREEN, WHITE),
        ("ITBA-CORPORATE PORTFOLIO KANBAN", "Owner: Carlos · P&T pool", "→ Corporate · Global SAP", BLUE, WHITE),
    ]
    py = 1.8
    for title, owner, routes, fill, txt in portfolios:
        add_rect(s, Inches(5.2), Inches(py), Inches(3.8), Inches(0.25), fill)
        add_text(s, Inches(5.2), Inches(py + 0.02), Inches(3.8), Inches(0.2),
                 title, font_size=8, bold=True, color=txt, align=PP_ALIGN.CENTER)
        add_rect(s, Inches(5.2), Inches(py + 0.25), Inches(3.8), Inches(1.0), WHITE, fill, Pt(0.5))
        add_text(s, Inches(5.3), Inches(py + 0.3), Inches(3.6), Inches(0.2),
                 owner, font_size=7, color=GREY_DARK)
        # mini kanban columns
        cols_n = ["Funnel", "Reviewing", "Analyzing", "Backlog", "Implem.", "Done"]
        ccw = 3.6 / len(cols_n)
        for ci, cn in enumerate(cols_n):
            cx = 5.3 + ci * ccw
            add_rect(s, Inches(cx), Inches(py + 0.5), Inches(ccw), Inches(0.7), GREY_LIGHT)
            add_text(s, Inches(cx), Inches(py + 0.5), Inches(ccw), Inches(0.18),
                     cn, font_size=6, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        add_text(s, Inches(5.3), Inches(py + 1.05), Inches(3.6), Inches(0.18),
                 routes, font_size=7, italic=True, color=GREY_DARK)
        py += 1.4

    # 7 Squad Boards (right side)
    add_text(s, Inches(9.2), Inches(1.5), Inches(3.8), Inches(0.25),
             "7 Squad Kanban Boards", font_size=8, bold=True, color=NAVY)
    squads = [
        ("SERVICE LIFECYCLE TEAM KANBAN", "M2C-SL · 8 FTE", AMBER, AMBER_DARK),
        ("REVENUE OPERATIONS TEAM KANBAN", "M2C-RO · 8 FTE", AMBER, AMBER_DARK),
        ("CUSTOMER CARE & CONTACT KANBAN", "CCC · 4 FTE", AMBER, AMBER_DARK),
        ("FIELD OPERATIONS TEAM KANBAN", "G&E · 7 FTE", GREEN, WHITE),
        ("GIS & OSG TEAM KANBAN", "GIS · 6 FTE", GREEN, WHITE),
        ("CORPORATE & REGULATORY TEAM KANBAN", "Corporate · 8 FTE", BLUE, WHITE),
        ("IBERDROLA GLOBAL SAP TEAM KANBAN", "Global SAP · 5 FTE", BLUE, WHITE),
    ]
    sy = 1.8
    for title, sub, fill, txt in squads:
        add_rect(s, Inches(9.2), Inches(sy), Inches(3.8), Inches(0.4), WHITE, fill, Pt(0.5))
        add_rect(s, Inches(9.2), Inches(sy), Inches(3.8), Inches(0.18), fill)
        add_text(s, Inches(9.2), Inches(sy + 0.01), Inches(3.0), Inches(0.18),
                 title, font_size=6, bold=True, color=txt)
        add_text(s, Inches(12.1), Inches(sy + 0.01), Inches(0.85), Inches(0.18),
                 sub, font_size=6, bold=True, color=txt, align=PP_ALIGN.RIGHT)
        # mini columns
        ccw = 3.8 / 4
        for ci in range(4):
            cx = 9.2 + ci * ccw
            add_rect(s, Inches(cx), Inches(sy + 0.18), Inches(ccw), Inches(0.22), GREY_LIGHT)
        sy += 0.55

    add_text(s, Inches(2.6), Inches(5.45), Inches(10), Inches(0.25),
             "Submits demand via Jira request and it auto-redirects to Customer Services or IT Business Applications", font_size=7, color=GREY_DARK)
    add_text(s, Inches(2.6), Inches(5.6), Inches(10), Inches(0.25),
             "Portfolio Kanban Board (Epics) → Squad Kanban Boards (Stories + Tasks).", font_size=7, color=GREY_DARK)

    # Bottom: Tools | Governance | Performance
    add_rect(s, Inches(0.2), Inches(5.95), Inches(13), Inches(0.4), NTT)
    add_text(s, Inches(0.2), Inches(6.0), Inches(2.5), Inches(0.3),
             "Tools", font_size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, Inches(2.7), Inches(6.0), Inches(7.5), Inches(0.3),
             "Governance", font_size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, Inches(10.2), Inches(6.0), Inches(3.0), Inches(0.3),
             "Performance", font_size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    add_rect(s, Inches(0.2), Inches(6.35), Inches(13), Inches(1.0), RGBColor(0xF2, 0xF6, 0xFA), NTT, Pt(1))

    # Tools
    add_round(s, Inches(0.3), Inches(6.45), Inches(1.1), Inches(0.4), WHITE, JIRA_BLUE, Pt(1.5))
    add_text(s, Inches(0.3), Inches(6.45), Inches(1.1), Inches(0.4),
             "Jira", font_size=12, bold=True, color=JIRA_BLUE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(1.45), Inches(6.45), Inches(1.2), Inches(0.4),
             "single intake (primary)", font_size=7, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
    add_round(s, Inches(0.3), Inches(6.9), Inches(1.1), Inches(0.4), WHITE, SUCCESS, Pt(1.5))
    add_text(s, Inches(0.3), Inches(6.9), Inches(1.1), Inches(0.4),
             "ServiceNow", font_size=10, bold=True, color=SUCCESS, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(1.45), Inches(6.9), Inches(1.2), Inches(0.4),
             "incidents only", font_size=7, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)

    # Governance
    add_round(s, Inches(2.8), Inches(6.45), Inches(1.7), Inches(0.3), GREEN)
    add_text(s, Inches(2.8), Inches(6.46), Inches(1.7), Inches(0.28),
             "GATE 1 · TRIAGE", font_size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_round(s, Inches(2.8), Inches(6.78), Inches(1.7), Inches(0.3), GREEN)
    add_text(s, Inches(2.8), Inches(6.79), Inches(1.7), Inches(0.28),
             "GATE 2 · PORTFOLIO", font_size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_round(s, Inches(2.8), Inches(7.1), Inches(1.7), Inches(0.25), NTT)
    add_text(s, Inches(2.8), Inches(7.11), Inches(1.7), Inches(0.23),
             "BRP · QUARTERLY", font_size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # CoS pills + outcomes
    cos = [("Standard", BLUE_DARK, WHITE), ("Fixed-Date", YELLOW, NAVY),
           ("Expedite", PINK, WHITE), ("Intangible", GREEN_LIGHT, NAVY)]
    cx = 4.7
    for label, fill, txt in cos:
        add_round(s, Inches(cx), Inches(6.55), Inches(1.0), Inches(0.25), fill)
        add_text(s, Inches(cx), Inches(6.56), Inches(1.0), Inches(0.23),
                 label, font_size=8, bold=True, color=txt, align=PP_ALIGN.CENTER)
        cx += 1.05
    outcomes = [("Accept", GREEN, WHITE), ("Send-back", YELLOW, NAVY),
                ("Reject", RED, WHITE), ("Redirect", PURPLE, WHITE)]
    cx = 4.7
    for label, fill, txt in outcomes:
        add_round(s, Inches(cx), Inches(6.85), Inches(1.0), Inches(0.25), fill)
        add_text(s, Inches(cx), Inches(6.86), Inches(1.0), Inches(0.23),
                 label, font_size=8, bold=True, color=txt, align=PP_ALIGN.CENTER)
        cx += 1.05
    add_text(s, Inches(4.7), Inches(7.15), Inches(5.0), Inches(0.25),
             "Forums: Replenishment · Daily Flow · Service Delivery Review · BRP", font_size=7, italic=True, color=GREY_DARK)

    # Performance
    perf = ["Demand volume", "DoR Pass Rate", "Triage lead time", "Channel adoption",
            "Lead Time", "Throughput", "WIP Aging", "Pain Score"]
    for i, label in enumerate(perf):
        col = i % 4
        row = i // 4
        x = Inches(10.0 + col * 0.78)
        y = Inches(6.45 + row * 0.45)
        add_round(s, x, y, Inches(0.75), Inches(0.4), WHITE, ORANGE_DARK, Pt(1))
        add_text(s, x, y, Inches(0.75), Inches(0.4),
                 label, font_size=7, bold=True, color=NAVY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    add_text(s, Inches(0.2), Inches(7.4), Inches(8), Inches(0.15),
             "© 2026 NTT DATA, Inc.", font_size=7, color=GREY_DARK)
    add_text(s, Inches(12.4), Inches(7.4), Inches(0.8), Inches(0.15),
             "13 / 13", font_size=7, color=GREY_DARK, align=PP_ALIGN.RIGHT)


# ============ ASSEMBLE ============
slide_01_cover()
slide_02_overview()
slide_03_vs1()
slide_04_vs2()
slide_05_vs3()
slide_06_architecture()
slide_07_intake_steps()
slide_08_jira()
slide_09_fields()
slide_10_portfolio()
slide_11_squad()
slide_12_cos()
slide_13_lifecycle()

prs.save(str(OUT))
print(f"Saved → {OUT}")
print(f"Slides: {len(prs.slides)}")
