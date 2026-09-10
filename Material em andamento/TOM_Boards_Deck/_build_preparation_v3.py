"""
Preparation & Planning · v3 — Senior consulting-deck quality with professional artefacts.
12 slides covering Phase 2 of the TOM with real data, formulas, and roles.

INPUT  · TOM MAP workbook - final.pptx
OUTPUT · TOM MAP workbook - final v3.pptx
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import zipfile, shutil

SRC = "/Users/gustavohenriquecastellano/Downloads/Base Apresentações Projeto Digital Product/TOM MAP workbook - final.pptx"
DST = "/Users/gustavohenriquecastellano/Downloads/Base Apresentações Projeto Digital Product/TOM MAP workbook - final v3.pptx"

# ============== Palette (matches Portfolio Board reference) ==============
NAVY_HEADER  = RGBColor(0x2D, 0x4E, 0x7C)
NAVY_DARK    = RGBColor(0x0F, 0x2A, 0x47)
HEADER_BLUE  = RGBColor(0x2C, 0x57, 0x8B)
HEADER_BLUE_DK = RGBColor(0x1B, 0x3C, 0x66)
WHITE        = RGBColor(0xFF, 0xFF, 0xFF)
TEXT         = RGBColor(0x17, 0x2B, 0x4D)
DARK_TEXT    = RGBColor(0x2C, 0x3E, 0x50)
GREY_LIGHT   = RGBColor(0xF2, 0xF6, 0xFA)
GREY_BG      = RGBColor(0xF4, 0xF5, 0xF7)
GREY_BORDER  = RGBColor(0xC9, 0xD2, 0xDD)
GREY_TEXT    = RGBColor(0x52, 0x60, 0x77)
GREY_LABEL   = RGBColor(0x6B, 0x77, 0x8C)
GREY_MUTED   = RGBColor(0x99, 0x99, 0x99)
GATE_RED     = RGBColor(0xB0, 0x35, 0x35)
GATE_RED_BG  = RGBColor(0xFB, 0xEC, 0xEC)

AVANGRID_GREEN = RGBColor(0x00, 0x8C, 0x38)
GREEN_DARK   = RGBColor(0x00, 0x6B, 0x3F)
GREEN_MID    = RGBColor(0x00, 0x89, 0x50)
GREEN_BG     = RGBColor(0xE6, 0xF2, 0xEB)
GREEN_PALE   = RGBColor(0xF2, 0xF8, 0xF4)
BLUE_BG      = RGBColor(0xEE, 0xF3, 0xF9)
BLUE_PALE    = RGBColor(0xF6, 0xFA, 0xFD)
PURPLE_HEATHER = RGBColor(0x6B, 0x3F, 0x8E)
PURPLE_LIGHT = RGBColor(0xA6, 0x84, 0xC0)
PURPLE_BG    = RGBColor(0xF2, 0xEB, 0xF6)

# Class of Service
COS_GREEN    = RGBColor(0x4F, 0x91, 0x4F)   # Standard
COS_BLUE     = RGBColor(0x3B, 0x6E, 0x8F)   # Fixed-Date
COS_RED      = RGBColor(0xB0, 0x35, 0x35)   # Expedite
COS_ORANGE   = RGBColor(0xE9, 0x8B, 0x2F)   # Intangible

# Bypass / annotation
BYPASS_ORANGE_BG = RGBColor(0xFD, 0xEC, 0xD4)
BYPASS_ORANGE    = RGBColor(0xF0, 0x86, 0x29)

JIRA_BG_LIGHT = RGBColor(0xF7, 0xF8, 0xF9)

TITLE_FONT = "IberPangea Text Medium"
BODY_FONT  = "Arial"


# ============== Helpers ==============
def _no_line(shape):
    shape.line.fill.background()

def rect(slide, x, y, w, h, fill=None, line=None, line_w=None):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    s.shadow.inherit = False
    if fill is None: s.fill.background()
    else: s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None: s.line.fill.background()
    else:
        s.line.color.rgb = line
        if line_w is not None: s.line.width = line_w
    return s

def round_rect(slide, x, y, w, h, fill=None, line=None, line_w=None, radius=0.06):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    s.shadow.inherit = False
    s.adjustments[0] = radius
    if fill is None: s.fill.background()
    else: s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None: s.line.fill.background()
    else:
        s.line.color.rgb = line
        if line_w is not None: s.line.width = line_w
    return s

def oval(slide, x, y, w, h, fill=None):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, w, h)
    s.shadow.inherit = False
    if fill is None: s.fill.background()
    else: s.fill.solid(); s.fill.fore_color.rgb = fill
    _no_line(s); return s

def line(slide, x1, y1, x2, y2, color=GREY_LABEL, weight=1.0):
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    ln.line.color.rgb = color
    ln.line.width = Pt(weight)
    return ln

def text(slide, x, y, w, h, txt, size=11, bold=False, italic=False, color=TEXT,
         font=BODY_FONT, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.margin_left = Emu(36000); tf.margin_right = Emu(36000)
    tf.margin_top = Emu(18000); tf.margin_bottom = Emu(18000)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    lines = txt.split("\n") if isinstance(txt, str) else [str(txt)]
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run(); r.text = ln
        r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
        r.font.name = font
        r.font.color.rgb = color
    return tb

def header_band(slide, kicker, subtitle):
    """Title block · IberPangea Text Medium · leaves top-right free for OwnIT! + Avangrid logos."""
    text(slide, Inches(0.40), Inches(0.10), Inches(9.80), Inches(0.55),
         kicker, 28, bold=True, color=AVANGRID_GREEN, font=TITLE_FONT,
         anchor=MSO_ANCHOR.MIDDLE)
    text(slide, Inches(0.40), Inches(0.62), Inches(9.80), Inches(0.42),
         subtitle, 20, bold=True, color=NAVY_DARK, font=TITLE_FONT,
         anchor=MSO_ANCHOR.MIDDLE)
    # Thin green divider below title block
    rect(slide, Inches(0.40), Inches(1.10), Inches(12.55), Emu(15000), AVANGRID_GREEN)

def footer_section(slide, slide_n, total_n=12):
    text(slide, Inches(0.40), Inches(7.18), Inches(8), Inches(0.18),
         "Avangrid · Digital Products · Target Operating Model  ·  Phase 2 · Preparation & Planning",
         8, italic=True, color=GREY_LABEL, font=BODY_FONT)
    text(slide, Inches(11.5), Inches(7.18), Inches(1.6), Inches(0.18),
         f"{slide_n:02d}/{total_n:02d}", 8, color=GREY_LABEL, font=BODY_FONT, align=PP_ALIGN.RIGHT)

def cos_swatch_legend(slide, y=7.05):
    text(slide, Inches(0.4), Inches(y), Inches(2.0), Inches(0.18),
         "CLASS OF SERVICE", 8, bold=True, color=GREY_LABEL, font=BODY_FONT)
    swatches = [("Standard", COS_GREEN), ("Fixed-Date", COS_BLUE),
                ("Expedite", COS_RED), ("Intangible", COS_ORANGE)]
    sx = 2.2
    for label, color in swatches:
        rect(slide, Inches(sx), Inches(y + 0.02), Inches(0.30), Inches(0.16), color)
        text(slide, Inches(sx + 0.36), Inches(y), Inches(1.4), Inches(0.18),
             label, 8, color=TEXT, font=BODY_FONT)
        sx += 1.85

def avatar(slide, x, y, d, color, initials, size_factor=22):
    oval(slide, Inches(x), Inches(y), Inches(d), Inches(d), color)
    text(slide, Inches(x), Inches(y), Inches(d), Inches(d), initials,
         size=int(d * size_factor), bold=True, color=WHITE, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

def chip(slide, x, y, w, label, color, text_color=WHITE, h=0.22, size=7):
    round_rect(slide, Inches(x), Inches(y), Inches(w), Inches(h), color, radius=0.30)
    text(slide, Inches(x), Inches(y), Inches(w), Inches(h), label,
         size=size, bold=True, color=text_color, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ============================================================================
# Smoke test data — used across slides for consistency
# ============================================================================
DEMO_ITEMS = [
    {"key": "BILLING OPERATIONS", "title": "Reduce billing-correction cycle 14→5 days",
     "vs": "Billing Operations", "area": "CS", "opco": "UIL · CMP · NY",
     "size": "M", "type": "Non-Reg", "cos": "Standard", "cos_color": COS_GREEN,
     "bv": 4, "ux": 5, "safety": None, "effort": 3, "pain": 2.43,
     "sponsor": "K. Stewart", "owner": "PM · Billing"},
    {"key": "REQ · Regulatory", "title": "PUC tariff banner update for Q3",
     "vs": "Corporate & Regulatory", "area": "ITBA", "opco": "All 5 (Iberdrola)",
     "size": "L", "type": "Regulatory", "cos": "Fixed-Date", "cos_color": COS_BLUE,
     "bv": 5, "ux": 3, "safety": 4, "effort": 4, "pain": 2.05,
     "sponsor": "M. Ferreira", "owner": "PM · Corp Reg"},
    {"key": "REQ · CRM perf", "title": "CRM page load <2s on Genesys",
     "vs": "Customer Care", "area": "CS", "opco": "UIL",
     "size": "S", "type": "Non-Reg", "cos": "Expedite", "cos_color": COS_RED,
     "bv": 3, "ux": 5, "safety": None, "effort": 2, "pain": 2.10,
     "sponsor": "H. LaJoie", "owner": "Squad · CC"},
    {"key": "REQ · GIS", "title": "GIS layer migration · scope-add",
     "vs": "Field Operations (CS)", "area": "CS", "opco": "UIL",
     "size": "S", "type": "Non-Reg", "cos": "Standard", "cos_color": COS_GREEN,
     "bv": 3, "ux": 4, "safety": None, "effort": 2, "pain": 1.80,
     "sponsor": "R. Mathews", "owner": "Squad · Field Ops"},
    {"key": "REQ · SAP patch", "title": "SAP Basis kernel patch Q2",
     "vs": "Iberdrola Corp Services", "area": "ITBA", "opco": "Iberdrola",
     "size": "M", "type": "Prod Support", "cos": "Intangible", "cos_color": COS_ORANGE,
     "bv": 2, "ux": 2, "safety": 3, "effort": 3, "pain": 0.95,
     "sponsor": "A. Soto", "owner": "PM · Iber Corp"},
    {"key": "REQ · Smart Meter", "title": "Smart Meter event analytics platform",
     "vs": "Smart Metering", "area": "CS", "opco": "UIL · NY",
     "size": "L", "type": "Non-Reg", "cos": "Standard", "cos_color": COS_GREEN,
     "bv": 4, "ux": 4, "safety": None, "effort": 4, "pain": 1.55,
     "sponsor": "H. LaJoie", "owner": "PM · Smart Meter"},
]


def render_story_card(slide, x, y, w, h, item, compact=False):
    """Render a Story/Epic card (consulting-deck style)."""
    # Card
    round_rect(slide, Inches(x), Inches(y), Inches(w), Inches(h), WHITE, GREY_BORDER, Pt(0.75), 0.04)
    # Top row: key + size badge
    text(slide, Inches(x + 0.08), Inches(y + 0.06), Inches(1.20), Inches(0.18),
         item["vs"][:18].upper(), 7, bold=True, color=GREY_LABEL, font=BODY_FONT)
    # CoS pill on right
    chip(slide, x + w - 1.20, y + 0.05, 1.10, item["cos"], item["cos_color"], WHITE, h=0.20, size=7)
    # Title
    text(slide, Inches(x + 0.08), Inches(y + 0.26), Inches(w - 0.16), Inches(0.50 if not compact else 0.40),
         item["title"], 9, bold=True, color=DARK_TEXT, font=BODY_FONT)
    # VS line
    text(slide, Inches(x + 0.08), Inches(y + h - 0.62), Inches(w - 0.16), Inches(0.18),
         f"▦ {item['vs']}", 7, italic=True, color=GREY_TEXT, font=BODY_FONT)
    # Footer row · size pill + Pain Score + Sponsor avatar
    fy = y + h - 0.32
    # Size pill
    size_color = {"S": COS_GREEN, "M": COS_ORANGE, "L": COS_BLUE}[item["size"]]
    chip(slide, x + 0.08, fy, 0.36, item["size"], size_color, WHITE, h=0.22, size=8)
    # Pain Score
    text(slide, Inches(x + 0.50), Inches(fy), Inches(1.40), Inches(0.22),
         f"Pain {item['pain']}", 8, bold=True, color=NAVY_DARK, font=BODY_FONT,
         anchor=MSO_ANCHOR.MIDDLE)
    # Sponsor mini-avatar
    initials = "".join([p[0] for p in item["sponsor"].replace(".", "").split()][:2])
    avatar(slide, x + w - 0.42, fy + 0.02, 0.22, NAVY_HEADER, initials, size_factor=18)


# ============================================================================
# 01 · DIVIDER · phase entry timeline
# ============================================================================
def build_pp01(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_band(s, "PREPARATION & PLANNING",
                "PHASE 2 OF THE TOM  ·  FROM GATE 2 GO TO READY FOR EXECUTION  ·  HEATHER + KUMAR + CARLOS CO-GOVERN")

    # 3-phase timeline
    tl_y = 1.45
    tl_w = 12.55; tl_x = 0.40
    rect(s, Inches(tl_x), Inches(tl_y + 1.05), Inches(tl_w), Emu(20000), GREY_BORDER)
    phases = [
        ("1", "DEMAND MANAGEMENT", "Intake · Gate 1 · Gate 2", GREY_MUTED, False),
        ("2", "PREPARATION & PLANNING", "Refinement · Estimation · Architecture · Capacity · DoR", NAVY_HEADER, True),
        ("3", "EXECUTION", "Build · Test · Release", GREY_MUTED, False),
    ]
    pw = (tl_w - 0.40) / 3
    for i, (num, name, sub, color, active) in enumerate(phases):
        x = tl_x + 0.10 + i * (pw + 0.10)
        d = 1.40 if active else 1.05
        cy = tl_y + 0.55 - d/2 + 0.50
        cx = x + pw/2 - d/2
        oval(s, Inches(cx), Inches(cy), Inches(d), Inches(d), color)
        text(s, Inches(cx), Inches(cy), Inches(d), Inches(d), num,
             size=int(d*30), bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, Inches(x), Inches(tl_y + 1.20), Inches(pw), Inches(0.28),
             name, size=12 if active else 11, bold=True,
             color=color if active else GREY_TEXT, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, Inches(x), Inches(tl_y + 1.50), Inches(pw), Inches(0.30),
             sub, size=9, italic=True, color=GREY_LABEL, font=BODY_FONT, align=PP_ALIGN.CENTER)
        if active:
            round_rect(s, Inches(x + pw/2 - 0.95), Inches(tl_y - 0.10),
                       Inches(1.90), Inches(0.24), GATE_RED, radius=0.30)
            text(s, Inches(x + pw/2 - 0.95), Inches(tl_y - 0.10), Inches(1.90), Inches(0.24),
                 "▼ COVERED IN THIS SECTION", size=8, bold=True, color=WHITE, font=BODY_FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Below — section content map (12 slides as section nav)
    map_y = 3.65
    map_h = 3.10
    round_rect(s, Inches(0.40), Inches(map_y), Inches(12.55), Inches(map_h), JIRA_BG_LIGHT, GREY_BORDER, Pt(0.5), 0.04)
    rect(s, Inches(0.40), Inches(map_y), Inches(12.55), Inches(0.40), NAVY_DARK)
    text(s, Inches(0.55), Inches(map_y + 0.05), Inches(12.30), Inches(0.30),
         "WHAT THIS SECTION COVERS  ·  12 SLIDES", 11, bold=True, color=WHITE, font=BODY_FONT,
         anchor=MSO_ANCHOR.MIDDLE)

    # Grid 4×3 of slide titles
    cols = 4; rows = 3
    grid_y = map_y + 0.55
    grid_w = 12.55 - 0.30
    cw = (grid_w - (cols - 1) * 0.15) / cols
    ch = (map_h - 0.65) / rows - 0.10

    titles = [
        ("01", "PHASE ENTRY", NAVY_HEADER),
        ("02", "DEMAND LIFECYCLE", NAVY_HEADER),
        ("03", "PREPARATION PIPELINE", NAVY_HEADER),
        ("04", "DEMAND PROFILE · EXAMPLE", NAVY_HEADER),
        ("05", "PRE-ESTIMATION · PAIN SCORE", AVANGRID_GREEN),
        ("06", "CAPACITY ALLOCATION", AVANGRID_GREEN),
        ("07", "ROADMAP · ROLLING WAVE", AVANGRID_GREEN),
        ("08", "SIZE PATHS · S · M · L", HEADER_BLUE),
        ("09", "EXISTING-INITIATIVE SHORTCUT", HEADER_BLUE),
        ("10", "OPERATING POLICIES · A1–A5", HEADER_BLUE),
        ("11", "HEATHER · PRODUCT LEAD", PURPLE_HEATHER),
        ("12", "ROLES · RACI", PURPLE_HEATHER),
    ]
    for i, (n, t, color) in enumerate(titles):
        col = i % cols; row = i // cols
        x = 0.55 + col * (cw + 0.15)
        y = grid_y + row * (ch + 0.10)
        round_rect(s, Inches(x), Inches(y), Inches(cw), Inches(ch), WHITE, GREY_BORDER, Pt(0.5), 0.04)
        rect(s, Inches(x), Inches(y), Inches(0.10), Inches(ch), color)
        text(s, Inches(x + 0.20), Inches(y + 0.08), Inches(0.50), Inches(0.20),
             n, size=12, bold=True, color=color, font=BODY_FONT)
        text(s, Inches(x + 0.70), Inches(y + 0.08), Inches(cw - 0.80), Inches(ch - 0.16),
             t, size=9, bold=True, color=DARK_TEXT, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)

    footer_section(s, 1)
    return s


# ============================================================================
# 02 · DEMAND LIFECYCLE · stages + transitions
# ============================================================================
def build_pp02(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_band(s, "DEMAND LIFECYCLE  ·  END-TO-END",
                "8 STAGES  ·  GATE 1 (TRIAGE) AND GATE 2 (PORTFOLIO)  ·  PHASE 2 OWNS THE BLUE BAND (REVIEWING · ANALYZING · BACKLOG)")

    # Workflow stations
    states = [
        ("NEW",          "Auto on form submit",   GREY_LABEL,    "auto"),
        ("ROUTED",       "A1 sets Component",     GREY_LABEL,    "auto"),
        ("TRIAGED",      "Gate 1 · weekly",       GATE_RED,      "Lead PM"),
        ("REVIEWING",    "Type · Size · CoS",     HEADER_BLUE,   "PM (per VS)"),
        ("ANALYZING",    "Scoring · Architecture", HEADER_BLUE,   "PM + Architect"),
        ("BACKLOG",      "Ready For Execution",   AVANGRID_GREEN,"Squad PM"),
        ("IMPLEMENTING", "Squad pulls",           HEADER_BLUE_DK,"Squad"),
        ("DONE",         "Outcome captured",      GREEN_DARK,    "PM + Sponsor"),
    ]

    # Top row · station boxes
    sb_y = 1.55
    sb_h = 1.05
    margin = 0.40
    avail_w = 12.55
    n = len(states)
    sb_w = (avail_w - (n - 1) * 0.10) / n
    for i, (name, sub, color, owner) in enumerate(states):
        x = margin + i * (sb_w + 0.10)
        round_rect(s, Inches(x), Inches(sb_y), Inches(sb_w), Inches(sb_h), color, radius=0.06)
        text(s, Inches(x), Inches(sb_y + 0.10), Inches(sb_w), Inches(0.28),
             name, size=11, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, Inches(x), Inches(sb_y + 0.40), Inches(sb_w), Inches(0.30),
             sub, size=8, italic=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
        # Owner pill at bottom
        round_rect(s, Inches(x + 0.10), Inches(sb_y + sb_h - 0.30), Inches(sb_w - 0.20), Inches(0.22),
                   WHITE, radius=0.30)
        text(s, Inches(x + 0.10), Inches(sb_y + sb_h - 0.30), Inches(sb_w - 0.20), Inches(0.22),
             owner, size=7, bold=True, color=color, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # arrow to next
        if i < n - 1:
            text(s, Inches(x + sb_w + 0.005), Inches(sb_y + 0.36), Inches(0.10), Inches(0.40),
                 "→", size=14, bold=True, color=GREY_LABEL, font=BODY_FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Gate dividers (vertical red bars between specific stations)
    g1_x = margin + 2 * sb_w + 1.5 * 0.10 + 0.005
    rect(s, Inches(g1_x), Inches(sb_y - 0.30), Inches(0.04), Inches(sb_h + 0.60), GATE_RED)
    text(s, Inches(g1_x - 1.0), Inches(sb_y - 0.32), Inches(2.0), Inches(0.22),
         "GATE 1 · TRIAGE", size=9, bold=True, color=GATE_RED, font=BODY_FONT, align=PP_ALIGN.CENTER)

    g2_x = margin + 5 * sb_w + 4.5 * 0.10 + 0.005
    rect(s, Inches(g2_x), Inches(sb_y - 0.30), Inches(0.04), Inches(sb_h + 0.60), GATE_RED)
    text(s, Inches(g2_x - 1.0), Inches(sb_y - 0.32), Inches(2.0), Inches(0.22),
         "GATE 2 · PORTFOLIO", size=9, bold=True, color=GATE_RED, font=BODY_FONT, align=PP_ALIGN.CENTER)

    # PHASE 2 banner · highlight statuses we own
    p2_y = sb_y + sb_h + 0.55
    p2_x = margin + 3 * sb_w + 2.5 * 0.10
    p2_w = 3 * sb_w + 2 * 0.10
    rect(s, Inches(p2_x), Inches(p2_y), Inches(p2_w), Inches(0.28), HEADER_BLUE)
    text(s, Inches(p2_x), Inches(p2_y), Inches(p2_w), Inches(0.28),
         "▲  PHASE 2 · PREPARATION & PLANNING — covered in this section",
         size=9, bold=True, color=WHITE, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Bottom — Status policies table (per Phase 2 status)
    tbl_y = 3.55
    tbl_h = 3.30
    round_rect(s, Inches(0.40), Inches(tbl_y), Inches(12.55), Inches(tbl_h), JIRA_BG_LIGHT, GREY_BORDER, Pt(0.5), 0.04)
    rect(s, Inches(0.40), Inches(tbl_y), Inches(12.55), Inches(0.40), NAVY_DARK)
    text(s, Inches(0.55), Inches(tbl_y + 0.05), Inches(12.30), Inches(0.30),
         "STATUS POLICIES  ·  PHASE 2 (REVIEWING · ANALYZING · BACKLOG)",
         11, bold=True, color=WHITE, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)

    cols_tbl = [("STATUS", 1.40), ("OWNER", 1.85), ("SLA", 1.20), ("INPUT", 3.50), ("OUTPUT", 4.50)]
    hdr_y = tbl_y + 0.50
    cx_acc = 0.55
    col_x = {}
    for lbl, w in cols_tbl:
        col_x[lbl] = cx_acc
        text(s, Inches(cx_acc), Inches(hdr_y), Inches(w), Inches(0.22),
             lbl, 8, bold=True, color=NAVY_DARK, font=BODY_FONT)
        cx_acc += w
    rect(s, Inches(0.55), Inches(hdr_y + 0.24), Inches(12.20), Emu(15000), NAVY_DARK)

    rows = [
        ("REVIEWING",     "PM (per VS)",       "2 weeks",
         "Demand · Title · BV narrative · OPCO · Sponsor",
         "Type · Size · Class of Service · R/G/T set"),
        ("ANALYZING",     "PM + Architect",    "1 week",
         "Reviewed Demand · Type · Size · CoS",
         "BV · UX · Safety (ITBA) · Effort scored · Pain Score auto · NFRs"),
        ("BACKLOG (RFE)", "Squad PM",          "—",
         "Gate 2 GO · capacity reserved",
         "DoR-passed Stories · ready to pull"),
    ]
    ry = hdr_y + 0.40
    rh = (tbl_h - 1.10) / len(rows)
    for ri, row in enumerate(rows):
        # row bg
        if ri % 2 == 1:
            rect(s, Inches(0.55), Inches(ry), Inches(12.20), Inches(rh - 0.05), GREY_LIGHT)
        for j, (lbl, w) in enumerate(cols_tbl):
            x = col_x[lbl]
            color = NAVY_DARK if j == 0 else (HEADER_BLUE if j == 1 else TEXT)
            bold = (j == 0)
            text(s, Inches(x), Inches(ry + 0.10), Inches(w - 0.05), Inches(rh - 0.20),
                 row[j], size=8.5 if j > 0 else 10, bold=bold, color=color, font=BODY_FONT,
                 anchor=MSO_ANCHOR.MIDDLE)
        ry += rh

    cos_swatch_legend(s)
    footer_section(s, 2)
    return s


# ============================================================================
# 03 · PREPARATION PIPELINE · 5 stages with example Stories
# ============================================================================
def build_pp03(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_band(s, "PREPARATION PIPELINE  ·  PHASE 2 IN MOTION",
                "5 STAGES OF WORK  ·  PM (per VS) → ARCHITECT → SQUAD PM  ·  ILLUSTRATIVE DEMANDS WITH PAIN SCORES")

    cols = [
        ("Refining",         "PM (per VS)",     HEADER_BLUE,    [DEMO_ITEMS[0], DEMO_ITEMS[5]]),
        ("Pre-Estimating",   "PM + Architect",  HEADER_BLUE,    [DEMO_ITEMS[1]]),
        ("Architecture",     "Architect",       AVANGRID_GREEN, [DEMO_ITEMS[2]]),
        ("DoR Pending",      "Squad PM",        AVANGRID_GREEN, [DEMO_ITEMS[3]]),
        ("Ready for Pull",   "Squad",           GREEN_DARK,     [DEMO_ITEMS[4]]),
    ]
    n = len(cols)
    margin = 0.40
    col_y = 1.55
    col_h = 5.10
    cw = (12.55 - (n - 1) * 0.10) / n
    for i, (title, owner, color, items) in enumerate(cols):
        x = margin + i * (cw + 0.10)
        # Header
        rect(s, Inches(x), Inches(col_y), Inches(cw), Inches(0.42), color)
        text(s, Inches(x), Inches(col_y + 0.04), Inches(cw), Inches(0.22),
             title, 11, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, Inches(x), Inches(col_y + 0.24), Inches(cw), Inches(0.18),
             owner, 7, italic=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
        # Body bg
        rect(s, Inches(x), Inches(col_y + 0.42), Inches(cw), Inches(col_h - 0.42), JIRA_BG_LIGHT)
        # Cards
        cy = col_y + 0.55
        for it in items:
            card_h = 1.35
            render_story_card(s, x + 0.10, cy, cw - 0.20, card_h, it)
            cy += card_h + 0.15
        # WIP indicator
        wip_lbl = {"Refining": "WIP 6", "Pre-Estimating": "WIP 4", "Architecture": "WIP 3",
                   "DoR Pending": "WIP 5", "Ready for Pull": "—"}[title]
        text(s, Inches(x), Inches(col_y + col_h - 0.22), Inches(cw), Inches(0.20),
             wip_lbl, 7, bold=True, color=GREY_LABEL, font=BODY_FONT, align=PP_ALIGN.CENTER)

    cos_swatch_legend(s)
    footer_section(s, 3)
    return s


# ============================================================================
# 04 · STORY CARD DETAIL · REQ · Billing deep-dive
# ============================================================================
def build_pp04(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_band(s, "DEMAND PROFILE  ·  ILLUSTRATIVE EXAMPLE",
                "REDUCE BILLING-CORRECTION CYCLE 14→5 DAYS  ·  ALL FIELDS POPULATED  ·  DoR CHECKLIST EMBEDDED")

    item = DEMO_ITEMS[0]

    # Demand profile frame
    card_x = 0.40; card_y = 1.30; card_w = 12.55; card_h = 5.55
    round_rect(s, Inches(card_x), Inches(card_y), Inches(card_w), Inches(card_h), WHITE, GREY_BORDER, Pt(0.75), 0.04)
    # Top bar
    rect(s, Inches(card_x), Inches(card_y), Inches(card_w), Inches(0.42), AVANGRID_GREEN)
    text(s, Inches(card_x + 0.15), Inches(card_y + 0.08), Inches(8), Inches(0.26),
         "DEMAND PROFILE  ·  Avangrid · Digital Products",
         10, bold=True, color=WHITE, font=BODY_FONT)
    # Status pill on right
    chip(s, card_x + card_w - 2.10, card_y + 0.08, 1.90, "STAGE · ANALYZING", HEADER_BLUE_DK, WHITE, h=0.26, size=8)

    # Title
    text(s, Inches(card_x + 0.20), Inches(card_y + 0.55), Inches(card_w - 0.40), Inches(0.40),
         item["title"], 16, bold=True, color=NAVY_DARK, font=BODY_FONT)

    # Two-column layout: left = fields, right = scoring + DoR
    left_x = card_x + 0.20
    left_w = 7.45
    right_x = card_x + 0.20 + left_w + 0.30
    right_w = card_w - (right_x - card_x) - 0.20
    body_y = card_y + 1.05

    # ==== LEFT · Fields ====
    def field(x, y, w, h, label, value, value_color=DARK_TEXT, value_bold=False):
        text(s, Inches(x), Inches(y), Inches(w), Inches(0.16),
             label, 7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        round_rect(s, Inches(x), Inches(y + 0.16), Inches(w), Inches(h - 0.16), JIRA_BG_LIGHT, GREY_BORDER, Pt(0.5), 0.04)
        text(s, Inches(x + 0.10), Inches(y + 0.20), Inches(w - 0.20), Inches(h - 0.22),
             value, 9, bold=value_bold, color=value_color, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)

    fy = body_y
    field(left_x, fy, left_w * 0.62, 0.50, "DESCRIPTION",
          "CC reps reconcile billing errors manually · 240 hrs/cycle · 8% inbound calls", value_color=DARK_TEXT)
    field(left_x + left_w * 0.65, fy, left_w * 0.35, 0.50, "VALUE STREAM",
          "Billing Operations  ·  CS")
    fy += 0.62
    field(left_x, fy, left_w * 0.30, 0.50, "TYPE", "Non-Regulatory")
    field(left_x + left_w * 0.32, fy, left_w * 0.20, 0.50, "SIZE", "M  ·  5–20 SP", value_bold=True)
    field(left_x + left_w * 0.54, fy, left_w * 0.30, 0.50, "CLASS OF SERVICE", "Standard", value_color=COS_GREEN, value_bold=True)
    field(left_x + left_w * 0.86, fy, left_w * 0.14, 0.50, "OPCOmult", "1.3", value_bold=True)
    fy += 0.62
    field(left_x, fy, left_w * 0.50, 0.50, "OPCO", "UIL  ·  CMP  ·  NY (3)")
    field(left_x + left_w * 0.52, fy, left_w * 0.48, 0.50, "SPONSOR", "K. Stewart  ·  Customer Care Director")
    fy += 0.62
    field(left_x, fy, left_w * 0.50, 0.55, "BUSINESS VALUE",
          "$285k/yr labor savings · NPS +5 · 240 hrs/cycle reduction")
    field(left_x + left_w * 0.52, fy, left_w * 0.48, 0.55, "ACCEPTANCE CRITERIA",
          "Cycle ≤ 5 days · UAT signed by Sponsor · no regression in CIS")
    fy += 0.67
    field(left_x, fy, left_w * 0.50, 0.55, "DEPENDENCIES",
          "CIS API rate-limit · Genesys script library · UAT env refresh")
    field(left_x + left_w * 0.52, fy, left_w * 0.48, 0.55, "ARCHITECTURE NOTE",
          "Latency: <300 ms p95 · CCPA: PII redaction · Vendor: Genesys SaaS")

    # ==== RIGHT · Scoring + DoR ====
    sc_y = body_y
    # Scoring panel
    round_rect(s, Inches(right_x), Inches(sc_y), Inches(right_w), Inches(1.45), GREEN_PALE, GREEN_DARK, Pt(0.75), 0.04)
    text(s, Inches(right_x + 0.10), Inches(sc_y + 0.06), Inches(right_w - 0.20), Inches(0.20),
         "PAIN SCORE  ·  CS formula  ·  no Safety", 8, bold=True, color=GREEN_DARK, font=BODY_FONT)
    # Score chips
    sx = right_x + 0.15; sy = sc_y + 0.30
    for label, val, color in [("BV", "4", AVANGRID_GREEN), ("UX", "5", GREEN_MID),
                              ("OPCOmult", "1.3", GREY_LABEL), ("Effort", "3", HEADER_BLUE)]:
        chip(s, sx, sy, 0.70, label, color, WHITE, h=0.22, size=7)
        text(s, Inches(sx), Inches(sy + 0.24), Inches(0.70), Inches(0.20),
             val, 11, bold=True, color=NAVY_DARK, font=BODY_FONT, align=PP_ALIGN.CENTER)
        sx += 0.78
    # Formula
    text(s, Inches(right_x + 0.10), Inches(sc_y + 0.78), Inches(right_w - 0.20), Inches(0.22),
         "(BV · OPCOmult · 0.40 + UX · 0.60) / Effort",
         9, italic=True, color=GREEN_DARK, font=BODY_FONT, align=PP_ALIGN.CENTER)
    text(s, Inches(right_x + 0.10), Inches(sc_y + 1.02), Inches(right_w - 0.20), Inches(0.32),
         "Pain Score = 2.43",
         16, bold=True, color=NAVY_DARK, font=BODY_FONT, align=PP_ALIGN.CENTER)

    # DoR Checklist
    dor_y = sc_y + 1.60
    dor_h = 2.65
    round_rect(s, Inches(right_x), Inches(dor_y), Inches(right_w), Inches(dor_h), WHITE, GREY_BORDER, Pt(0.75), 0.04)
    rect(s, Inches(right_x), Inches(dor_y), Inches(right_w), Inches(0.34), NAVY_DARK)
    text(s, Inches(right_x + 0.10), Inches(dor_y + 0.04), Inches(right_w - 0.20), Inches(0.26),
         "DEFINITION OF READY  ·  4 of 6 ✓",
         10, bold=True, color=WHITE, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)

    dor_items = [
        (True,  "Outcome clear",       "AC binary · UAT-testable"),
        (True,  "Sized",                "M · 5–20 SP · split into 3 Stories"),
        (True,  "Architecture reviewed","NFRs · vendor SLA noted"),
        (False, "Dependencies known",   "CIS rate-limit confirmation pending"),
        (True,  "Sponsor signed off",   "K. Stewart · 2026-05-04"),
        (False, "Capacity reserved",    "Standard lane @ WIP cap · waiting"),
    ]
    iy = dor_y + 0.42
    rh = (dor_h - 0.50) / len(dor_items)
    for done, lbl, sub in dor_items:
        cb_color = AVANGRID_GREEN if done else GATE_RED
        icon = "✓" if done else "✕"
        oval(s, Inches(right_x + 0.12), Inches(iy + 0.04), Inches(0.22), Inches(0.22), cb_color)
        text(s, Inches(right_x + 0.12), Inches(iy + 0.04), Inches(0.22), Inches(0.22),
             icon, 10, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, Inches(right_x + 0.42), Inches(iy + 0.02), Inches(right_w - 0.55), Inches(0.18),
             lbl, 9, bold=True, color=DARK_TEXT, font=BODY_FONT)
        text(s, Inches(right_x + 0.42), Inches(iy + 0.20), Inches(right_w - 0.55), Inches(rh - 0.22),
             sub, 7, italic=True, color=GREY_LABEL, font=BODY_FONT)
        iy += rh

    footer_section(s, 4)
    return s


# ============================================================================
# 05 · PRE-ESTIMATION + PAIN SCORE · CS and ITBA panels
# ============================================================================
def build_pp05(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_band(s, "PRE-ESTIMATION  ·  PAIN SCORE",
                "TWO FORMULAS · ONE BACKLOG  ·  CS = NO SAFETY (CUSTOMER OUTCOMES)  ·  ITBA = SAFETY ISOLATED 25% (FIELD-CREW IMPACT)")

    # Two large scoring panels (REQ · Billing CS · REQ · Regulatory ITBA)
    panel_y = 1.30
    panel_h = 5.50
    sw2 = (12.55 - 0.20) / 2

    def render_score_panel(x, item, side_color, side_dk, side_bg, headline):
        # Frame
        round_rect(s, Inches(x), Inches(panel_y), Inches(sw2), Inches(panel_h),
                   WHITE, GREY_BORDER, Pt(0.75), 0.04)
        # Header
        rect(s, Inches(x), Inches(panel_y), Inches(sw2), Inches(0.45), side_color)
        text(s, Inches(x + 0.15), Inches(panel_y + 0.08), Inches(sw2 - 0.30), Inches(0.30),
             headline, 12, bold=True, color=WHITE, font=BODY_FONT)
        # Story header
        sy = panel_y + 0.55
        text(s, Inches(x + 0.20), Inches(sy), Inches(sw2 - 0.40), Inches(0.20),
             item["key"], 8, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, Inches(x + 0.20), Inches(sy + 0.18), Inches(sw2 - 0.40), Inches(0.30),
             item["title"], 12, bold=True, color=NAVY_DARK, font=BODY_FONT)
        chip(s, x + sw2 - 1.40, sy + 0.20, 1.20, item["cos"], item["cos_color"], WHITE, h=0.22, size=7)

        # Inputs row · sliders
        sl_y = sy + 0.65
        if item["area"] == "CS":
            inputs = [("BV  ·  Business Value", item["bv"], AVANGRID_GREEN, "PM"),
                      ("UX  ·  User Experience", item["ux"], GREEN_MID,    "PM"),
                      ("Effort", item["effort"], HEADER_BLUE, "Architect"),
                      ("OPCOmult  ·  auto", "1.3" if item["opco"].count("·") > 0 else "1.0",
                                                  GREY_LABEL, "auto")]
        else:
            inputs = [("BV  ·  Business Value", item["bv"], HEADER_BLUE, "PM"),
                      ("UX  ·  User Experience", item["ux"], BLUE_PALE,  "PM"),
                      ("Safety  ·  field-crew",  item["safety"], HEADER_BLUE_DK, "Architect"),
                      ("Effort", item["effort"], COS_BLUE, "Architect"),
                      ("OPCOmult  ·  auto", "1.6" if "All" in item["opco"] else "1.0",
                                                  GREY_LABEL, "auto")]
        rh = 0.42
        for lbl, val, color, owner in inputs:
            text(s, Inches(x + 0.20), Inches(sl_y + 0.04), Inches(2.30), Inches(0.20),
                 lbl, 8, bold=True, color=DARK_TEXT, font=BODY_FONT)
            # 5-step slider visualisation
            slider_x = x + 2.65
            slider_w = sw2 - 3.85
            step_w = slider_w / 5
            for st in range(5):
                ssx = slider_x + st * step_w
                # filled if val (numeric or non-numeric)
                if isinstance(val, int):
                    filled = (st + 1) <= val
                else:
                    filled = False
                bg = color if filled else GREY_LIGHT
                rect(s, Inches(ssx + 0.04), Inches(sl_y + 0.10), Inches(step_w - 0.08), Inches(0.20),
                     bg)
            # value chip
            chip(s, x + sw2 - 1.05, sl_y + 0.06, 0.40, str(val) if val is not None else "—",
                 color, WHITE, h=0.28, size=10)
            # owner
            text(s, Inches(x + sw2 - 0.55), Inches(sl_y + 0.10), Inches(0.50), Inches(0.20),
                 owner, 7, italic=True, color=GREY_LABEL, font=BODY_FONT)
            sl_y += rh

        # Formula box
        f_y = sl_y + 0.10
        round_rect(s, Inches(x + 0.20), Inches(f_y), Inches(sw2 - 0.40), Inches(0.55),
                   side_bg, side_color, Pt(0.5), 0.04)
        if item["area"] == "CS":
            formula = "(BV · OPCOmult · 0.40 + UX · 0.60) / Effort"
        else:
            formula = "(BV · OPCOmult · 0.45 + UX · 0.30 + Safety · 0.25) / Effort"
        text(s, Inches(x + 0.20), Inches(f_y + 0.04), Inches(sw2 - 0.40), Inches(0.22),
             "FORMULA",
             7, bold=True, color=side_dk, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, Inches(x + 0.20), Inches(f_y + 0.24), Inches(sw2 - 0.40), Inches(0.30),
             formula, 11, bold=True, color=NAVY_DARK, font=BODY_FONT, align=PP_ALIGN.CENTER)

        # Pain Score result
        ps_y = f_y + 0.70
        rect(s, Inches(x + 0.20), Inches(ps_y), Inches(sw2 - 0.40), Inches(0.65), side_color)
        text(s, Inches(x + 0.20), Inches(ps_y + 0.04), Inches(sw2 - 0.40), Inches(0.22),
             "PAIN SCORE  ·  auto-computed at Analyzing", 7, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER)
        text(s, Inches(x + 0.20), Inches(ps_y + 0.22), Inches(sw2 - 0.40), Inches(0.40),
             f"{item['pain']}", 24, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)

    render_score_panel(0.40, DEMO_ITEMS[0], AVANGRID_GREEN, GREEN_DARK, GREEN_PALE,
                       "CS  ·  ILLUSTRATIVE EXAMPLE  ·  no Safety component")
    render_score_panel(0.40 + sw2 + 0.20, DEMO_ITEMS[1], HEADER_BLUE, HEADER_BLUE_DK, BLUE_PALE,
                       "ITBA  ·  ILLUSTRATIVE EXAMPLE  ·  Safety isolated 25%")

    cos_swatch_legend(s)
    footer_section(s, 5)
    return s


# ============================================================================
# 06 · CAPACITY ALLOCATION · squads × CoS lanes view
# ============================================================================
def build_pp06(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_band(s, "CAPACITY ALLOCATION  ·  PER SQUAD · PER CoS LANE",
                "Q3 2026 ILLUSTRATIVE SPLIT  ·  % SET MONTHLY AT GATE 2  ·  9 SQUADS (4 CS + 5 ITBA)")

    # Cell layout: 9 squads (rows) × 4 CoS lanes (columns) + total
    sq_y = 1.30
    sq_h = 5.45
    label_w = 2.50
    total_w = 1.30
    cos_count = 4
    cell_w = (12.55 - label_w - total_w) / cos_count

    # Header row
    rect(s, Inches(0.40), Inches(sq_y), Inches(label_w), Inches(0.42), NAVY_DARK)
    text(s, Inches(0.40), Inches(sq_y), Inches(label_w), Inches(0.42),
         "SQUAD  ·  per Value Stream", 9, bold=True, color=WHITE, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    cos_lanes = [
        ("STANDARD",   COS_GREEN),
        ("FIXED-DATE", COS_BLUE),
        ("EXPEDITE",   COS_RED),
        ("INTANGIBLE", COS_ORANGE),
    ]
    for ci, (lbl, color) in enumerate(cos_lanes):
        cx = 0.40 + label_w + ci * cell_w
        rect(s, Inches(cx), Inches(sq_y), Inches(cell_w - 0.04), Inches(0.42), color)
        text(s, Inches(cx), Inches(sq_y), Inches(cell_w - 0.04), Inches(0.42),
             lbl, 9, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, Inches(0.40 + label_w + cos_count * cell_w), Inches(sq_y), Inches(total_w), Inches(0.42), NAVY_DARK)
    text(s, Inches(0.40 + label_w + cos_count * cell_w), Inches(sq_y), Inches(total_w), Inches(0.42),
         "TOTAL FTE", 9, bold=True, color=WHITE, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Squads (with realistic capacity %)
    squads = [
        ("Customer Care",            "CS",   AVANGRID_GREEN, [55, 15, 20, 10], "8 FTE"),
        ("Billing Operations",       "CS",   AVANGRID_GREEN, [60, 15, 15, 10], "9 FTE"),
        ("Field Operations (CS)",    "CS",   AVANGRID_GREEN, [50, 25, 15, 10], "7 FTE"),
        ("Smart Metering",           "CS",   AVANGRID_GREEN, [55, 25, 10, 10], "6 FTE"),
        ("Corporate & Regulatory",   "ITBA", HEADER_BLUE,    [40, 45,  5, 10], "10 FTE"),
        ("Iberdrola Corp Services",  "ITBA", HEADER_BLUE,    [55, 20, 10, 15], "8 FTE"),
        ("Renewables IT",            "ITBA", HEADER_BLUE,    [60, 15, 15, 10], "7 FTE"),
        ("Networks IT",              "ITBA", HEADER_BLUE,    [50, 30, 10, 10], "9 FTE"),
        ("Enabling Team",            "BOTH", GREY_LABEL,     [40, 20, 20, 20], "5 FTE"),
    ]
    rh = (sq_h - 0.42) / len(squads)
    ry = sq_y + 0.42
    for ri, (name, area, color, pcts, fte) in enumerate(squads):
        # Squad label cell
        bg = WHITE if ri % 2 == 0 else GREY_LIGHT
        rect(s, Inches(0.40), Inches(ry), Inches(label_w), Inches(rh - 0.02), bg)
        rect(s, Inches(0.40), Inches(ry), Inches(0.10), Inches(rh - 0.02), color)
        text(s, Inches(0.55), Inches(ry + 0.06), Inches(label_w - 0.20), Inches(0.20),
             name, 9, bold=True, color=DARK_TEXT, font=BODY_FONT)
        text(s, Inches(0.55), Inches(ry + 0.26), Inches(label_w - 0.20), Inches(0.18),
             f"{area}  ·  {fte}", 7, italic=True, color=GREY_LABEL, font=BODY_FONT)
        # CoS cells
        for ci, (lbl, color_c) in enumerate(cos_lanes):
            cx = 0.40 + label_w + ci * cell_w
            rect(s, Inches(cx), Inches(ry), Inches(cell_w - 0.04), Inches(rh - 0.02), bg)
            # utilization bar
            pct = pcts[ci]
            bar_x = cx + 0.10
            bar_y = ry + rh / 2 - 0.10
            bar_w_max = cell_w - 0.30
            rect(s, Inches(bar_x), Inches(bar_y), Inches(bar_w_max), Inches(0.18), GREY_LIGHT)
            rect(s, Inches(bar_x), Inches(bar_y), Inches(bar_w_max * pct / 100), Inches(0.18), color_c)
            text(s, Inches(cx), Inches(ry + rh - 0.34), Inches(cell_w - 0.04), Inches(0.20),
                 f"{pct}%", 9, bold=True, color=NAVY_DARK, font=BODY_FONT, align=PP_ALIGN.CENTER)
        # Total
        cx = 0.40 + label_w + cos_count * cell_w
        rect(s, Inches(cx), Inches(ry), Inches(total_w), Inches(rh - 0.02), bg)
        text(s, Inches(cx), Inches(ry), Inches(total_w), Inches(rh - 0.02),
             fte, 10, bold=True, color=NAVY_DARK, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        ry += rh

    cos_swatch_legend(s)
    footer_section(s, 6)
    return s


# ============================================================================
# 07 · ROADMAP · Plans timeline (Gantt)
# ============================================================================
def build_pp07(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_band(s, "ROADMAP  ·  ROLLING WAVE  ·  PER VALUE STREAM",
                "3 MO COMMITTED · 3–6 MO PLANNED · 6–12 MO HYPOTHESIZED  ·  RE-BALANCED MONTHLY AT GATE 2")

    # Timeline axis
    ax_x = 2.50; ax_y = 1.50; ax_w = 10.45
    quarters = ["Q3 2026", "Q4 2026", "Q1 2027", "Q2 2027"]
    qw = ax_w / 4
    for i, q in enumerate(quarters):
        rect(s, Inches(ax_x + i * qw), Inches(ax_y), Inches(qw - 0.04), Inches(0.32),
             NAVY_DARK if i < 1 else (HEADER_BLUE if i < 2 else GREY_LABEL))
        text(s, Inches(ax_x + i * qw), Inches(ax_y), Inches(qw - 0.04), Inches(0.32),
             q, 10, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Horizon bands
    band_y = ax_y + 0.36
    text(s, Inches(ax_x), Inches(band_y), Inches(qw), Inches(0.20),
         "COMMITTED", 8, bold=True, color=GREEN_DARK, font=BODY_FONT, align=PP_ALIGN.CENTER)
    text(s, Inches(ax_x + qw), Inches(band_y), Inches(qw), Inches(0.20),
         "PLANNED", 8, bold=True, color=HEADER_BLUE, font=BODY_FONT, align=PP_ALIGN.CENTER)
    text(s, Inches(ax_x + 2 * qw), Inches(band_y), Inches(2 * qw), Inches(0.20),
         "HYPOTHESIZED", 8, bold=True, color=GREY_LABEL, font=BODY_FONT, align=PP_ALIGN.CENTER)

    # Initiatives by VS (rows)
    rows_def = [
        ("Customer Care",          "CS",   AVANGRID_GREEN, [
            ("Genesys script overhaul",     0.00, 0.55, AVANGRID_GREEN, COS_GREEN),
            ("CRM page load <2s",           0.10, 0.30, COS_RED,        COS_RED),
        ]),
        ("Billing Operations",     "CS",   AVANGRID_GREEN, [
            ("Billing cycle reduction", 0.05, 0.50, AVANGRID_GREEN, COS_GREEN),
            ("Self-service password",       0.15, 0.45, GREEN_MID,      COS_GREEN),
        ]),
        ("Smart Metering",         "CS",   AVANGRID_GREEN, [
            ("Smart Meter analytics",       0.40, 0.95, GREY_LABEL,     COS_GREEN),
        ]),
        ("Corporate & Regulatory", "ITBA", HEADER_BLUE,    [
            ("PUC tariff banner Q3",        0.05, 0.30, HEADER_BLUE,    COS_BLUE),
            ("Audit framework refresh",     0.40, 0.85, GREY_LABEL,     COS_BLUE),
        ]),
        ("Iberdrola Corp Services","ITBA", HEADER_BLUE,    [
            ("SAP Basis kernel patch",      0.10, 0.30, HEADER_BLUE_DK, COS_ORANGE),
        ]),
        ("Networks IT",            "ITBA", HEADER_BLUE,    [
            ("Networks data backbone",      0.30, 0.60, HEADER_BLUE,    COS_BLUE),
            ("S/4HANA migration plan",      0.65, 0.95, GREY_LABEL,     COS_GREEN),
        ]),
    ]
    rows_y = band_y + 0.30
    rows_h_avail = 7.00 - rows_y - 0.20
    rh = rows_h_avail / len(rows_def)
    for ri, (vs_name, area, vs_color, items) in enumerate(rows_def):
        ry = rows_y + ri * rh
        bg = WHITE if ri % 2 == 0 else GREY_LIGHT
        # Label
        rect(s, Inches(0.40), Inches(ry), Inches(2.10), Inches(rh - 0.02), bg)
        rect(s, Inches(0.40), Inches(ry), Inches(0.10), Inches(rh - 0.02), vs_color)
        text(s, Inches(0.55), Inches(ry + 0.05), Inches(1.95), Inches(0.20),
             vs_name, 9, bold=True, color=DARK_TEXT, font=BODY_FONT)
        text(s, Inches(0.55), Inches(ry + 0.24), Inches(1.95), Inches(0.18),
             area, 7, italic=True, color=GREY_LABEL, font=BODY_FONT)
        # Plot row bg
        rect(s, Inches(ax_x), Inches(ry), Inches(ax_w), Inches(rh - 0.02), bg)
        # Vertical quarter dividers
        for q in range(1, 4):
            rect(s, Inches(ax_x + q * qw), Inches(ry), Emu(7000), Inches(rh - 0.02), GREY_BORDER)
        # Bars
        for j, (lbl, sx, ex, color, cos_c) in enumerate(items):
            bx = ax_x + sx * ax_w
            bw = (ex - sx) * ax_w
            by = ry + 0.06 + j * (rh - 0.20) / max(1, len(items))
            bh = (rh - 0.20) / max(1, len(items)) - 0.06
            round_rect(s, Inches(bx), Inches(by), Inches(bw), Inches(bh), color, radius=0.04)
            # CoS small dot at left
            oval(s, Inches(bx + 0.04), Inches(by + bh/2 - 0.06), Inches(0.12), Inches(0.12), cos_c)
            text(s, Inches(bx + 0.22), Inches(by + 0.04), Inches(bw - 0.30), Inches(bh - 0.08),
                 lbl, 8, bold=True, color=WHITE, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)

    cos_swatch_legend(s)
    footer_section(s, 7)
    return s


# ============================================================================
# 08 · SIZE PATHS · 3 cards (S Story, M Epic+children, L Initiative+phases)
# ============================================================================
def build_pp08(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_band(s, "SIZE PATHS  ·  S · M · L",
                "SAME LOGIC · SAME BACKLOG  ·  SIZE DICTATES SHAPE  ·  XL → PPM Now (out of scope)")

    # 3 columns
    col_y = 1.30
    col_h = 5.55
    cw = (12.55 - 0.40) / 3

    # === S column · Story directly ===
    sx = 0.40
    round_rect(s, Inches(sx), Inches(col_y), Inches(cw), Inches(col_h), WHITE, GREY_BORDER, Pt(0.75), 0.04)
    rect(s, Inches(sx), Inches(col_y), Inches(cw), Inches(0.50), AVANGRID_GREEN)
    text(s, Inches(sx + 0.15), Inches(col_y + 0.06), Inches(cw - 0.30), Inches(0.22),
         "S  ·  ≤ 5 SP", 13, bold=True, color=WHITE, font=BODY_FONT)
    text(s, Inches(sx + 0.15), Inches(col_y + 0.28), Inches(cw - 0.30), Inches(0.20),
         "1 week  ·  Squad-only refinement",
         8, italic=True, color=WHITE, font=BODY_FONT)
    # Single Story card
    card_y = col_y + 0.65
    render_story_card(s, sx + 0.20, card_y, cw - 0.40, 1.35, DEMO_ITEMS[2])
    # Pipeline below
    flow_y = card_y + 1.50
    flow_steps = [("Story", AVANGRID_GREEN), ("Refine", GREEN_MID), ("DoR", GATE_RED), ("Pull", HEADER_BLUE)]
    fw_each = (cw - 0.50) / len(flow_steps)
    for i, (lbl, color) in enumerate(flow_steps):
        x = sx + 0.20 + i * fw_each
        round_rect(s, Inches(x + 0.04), Inches(flow_y), Inches(fw_each - 0.10), Inches(0.32), color, radius=0.30)
        text(s, Inches(x + 0.04), Inches(flow_y), Inches(fw_each - 0.10), Inches(0.32),
             lbl, 8, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Notes
    notes_y = flow_y + 0.50
    notes = [
        "Skip Epic — Story drops directly under VS",
        "Squad PM owns DoR (no Architect unless NFRs)",
        "Pull within current iteration",
        "Watch out: don't over-batch Ss into forced Epics",
    ]
    iy = notes_y
    for n in notes:
        oval(s, Inches(sx + 0.25), Inches(iy + 0.04), Inches(0.14), Inches(0.14), AVANGRID_GREEN)
        text(s, Inches(sx + 0.45), Inches(iy), Inches(cw - 0.65), Inches(0.32),
             n, 8.5, color=DARK_TEXT, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
        iy += 0.42

    # === M column · Epic + 3 Stories ===
    sx = 0.40 + cw + 0.20
    round_rect(s, Inches(sx), Inches(col_y), Inches(cw), Inches(col_h), WHITE, GREY_BORDER, Pt(0.75), 0.04)
    rect(s, Inches(sx), Inches(col_y), Inches(cw), Inches(0.50), COS_ORANGE)
    text(s, Inches(sx + 0.15), Inches(col_y + 0.06), Inches(cw - 0.30), Inches(0.22),
         "M  ·  5–20 SP", 13, bold=True, color=WHITE, font=BODY_FONT)
    text(s, Inches(sx + 0.15), Inches(col_y + 0.28), Inches(cw - 0.30), Inches(0.20),
         "1–2 weeks  ·  VS-level refinement  ·  A5 auto-converts",
         8, italic=True, color=WHITE, font=BODY_FONT)
    # Epic card
    epic_y = col_y + 0.65
    epic_item = DEMO_ITEMS[0]  # REQ · Billing
    render_story_card(s, sx + 0.20, epic_y, cw - 0.40, 1.10, epic_item, compact=True)
    # 3 child stories (smaller)
    child_y = epic_y + 1.25
    children = [
        {"key": "STORY · A", "title": "CIS API rate-limit upgrade", "vs": "Billing", "size": "S",
         "cos": "Standard", "cos_color": COS_GREEN, "pain": 2.0, "sponsor": "K. Stewart"},
        {"key": "STORY · B", "title": "Genesys auto-reconcile flow", "vs": "Billing", "size": "S",
         "cos": "Standard", "cos_color": COS_GREEN, "pain": 2.4, "sponsor": "K. Stewart"},
        {"key": "STORY · C", "title": "UAT data refresh script",     "vs": "Billing", "size": "S",
         "cos": "Standard", "cos_color": COS_GREEN, "pain": 1.8, "sponsor": "K. Stewart"},
    ]
    cw_child = (cw - 0.60) / 3
    for j, ch in enumerate(children):
        cx = sx + 0.20 + j * (cw_child + 0.10)
        round_rect(s, Inches(cx), Inches(child_y), Inches(cw_child), Inches(0.95), WHITE, GREY_BORDER, Pt(0.5), 0.04)
        rect(s, Inches(cx), Inches(child_y), Inches(cw_child), Inches(0.05), COS_ORANGE)
        text(s, Inches(cx + 0.06), Inches(child_y + 0.10), Inches(cw_child - 0.12), Inches(0.16),
             ch["key"], 6, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, Inches(cx + 0.06), Inches(child_y + 0.26), Inches(cw_child - 0.12), Inches(0.40),
             ch["title"], 7.5, bold=True, color=DARK_TEXT, font=BODY_FONT)
        chip(s, cx + 0.06, child_y + 0.70, 0.30, ch["size"], COS_GREEN, WHITE, h=0.18, size=7)
        text(s, Inches(cx + 0.40), Inches(child_y + 0.68), Inches(cw_child - 0.45), Inches(0.20),
             f"Pain {ch['pain']}", 7, bold=True, color=NAVY_DARK, font=BODY_FONT)
        # connector to Epic
        line(s, Inches(sx + cw / 2), Inches(epic_y + 1.10),
             Inches(cx + cw_child / 2), Inches(child_y), GREY_LABEL, weight=0.75)
    # bottom note
    text(s, Inches(sx + 0.20), Inches(col_y + col_h - 0.42), Inches(cw - 0.40), Inches(0.30),
         "Architect NFR pre-check · DoR per Story · Heather reviews outcome framing",
         8, italic=True, color=GREY_TEXT, font=BODY_FONT, align=PP_ALIGN.CENTER)

    # === L column · Initiative + phases ===
    sx = 0.40 + 2 * (cw + 0.20)
    round_rect(s, Inches(sx), Inches(col_y), Inches(cw), Inches(col_h), WHITE, GREY_BORDER, Pt(0.75), 0.04)
    rect(s, Inches(sx), Inches(col_y), Inches(cw), Inches(0.50), HEADER_BLUE)
    text(s, Inches(sx + 0.15), Inches(col_y + 0.06), Inches(cw - 0.30), Inches(0.22),
         "L  ·  20–60 SP", 13, bold=True, color=WHITE, font=BODY_FONT)
    text(s, Inches(sx + 0.15), Inches(col_y + 0.28), Inches(cw - 0.30), Inches(0.20),
         "2–4 weeks  ·  Initiative-shape  ·  multi-quarter capacity",
         8, italic=True, color=WHITE, font=BODY_FONT)
    # Initiative card
    init_y = col_y + 0.65
    init_item = DEMO_ITEMS[1]  # REQ · Regulatory
    render_story_card(s, sx + 0.20, init_y, cw - 0.40, 1.10, init_item, compact=True)
    # Phased Stories (vertical list)
    ph_y = init_y + 1.25
    phases = [
        ("Q3", "Tariff data model + audit log",   "STORY · A · 8 SP",  HEADER_BLUE),
        ("Q3", "PUC submission webhook",           "STORY · B · 5 SP",  HEADER_BLUE),
        ("Q4", "Multi-OPCO rollout (5 OPCOs)",     "STORY · C · 12 SP", COS_BLUE),
        ("Q4", "Compliance audit harness",         "STORY · D · 10 SP", COS_BLUE),
    ]
    for j, (q, title, key, color) in enumerate(phases):
        py = ph_y + j * 0.42
        round_rect(s, Inches(sx + 0.20), Inches(py), Inches(cw - 0.40), Inches(0.36),
                   WHITE, GREY_BORDER, Pt(0.5), 0.04)
        chip(s, sx + 0.27, py + 0.05, 0.50, q, color, WHITE, h=0.26, size=8)
        text(s, Inches(sx + 0.85), Inches(py + 0.04), Inches(cw - 1.10), Inches(0.16),
             title, 8, bold=True, color=DARK_TEXT, font=BODY_FONT)
        text(s, Inches(sx + 0.85), Inches(py + 0.20), Inches(cw - 1.10), Inches(0.16),
             key, 7, italic=True, color=GREY_LABEL, font=BODY_FONT)

    cos_swatch_legend(s)
    footer_section(s, 8)
    return s


# ============================================================================
# 09 · DEMAND-ON-EXISTING-INITIATIVE · parent-child + decision matrix
# ============================================================================
def build_pp09(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_band(s, "DEMAND-ON-EXISTING-INITIATIVE  ·  SHORTCUT",
                "PARENT INITIATIVE STATUS DRIVES THE PATH  ·  ACCEPT (Implementing/Backlog) OR REJECT (earlier states)")

    # LEFT · parent-child tree
    tx = 0.40; ty = 1.30; tw = 6.30; th = 5.55
    round_rect(s, Inches(tx), Inches(ty), Inches(tw), Inches(th), WHITE, GREY_BORDER, Pt(0.75), 0.04)
    rect(s, Inches(tx), Inches(ty), Inches(tw), Inches(0.42), NAVY_HEADER)
    text(s, Inches(tx + 0.15), Inches(ty + 0.08), Inches(tw - 0.30), Inches(0.26),
         "PARENT-CHILD STRUCTURE  ·  scope-add example", 11, bold=True, color=WHITE, font=BODY_FONT)

    # Parent Initiative card (Implementing)
    pa_y = ty + 0.60
    pa_h = 1.30
    round_rect(s, Inches(tx + 0.20), Inches(pa_y), Inches(tw - 0.40), Inches(pa_h),
               WHITE, HEADER_BLUE_DK, Pt(1), 0.04)
    chip(s, tx + 0.30, pa_y + 0.10, 1.50, "STAGE · IMPLEMENTING", AVANGRID_GREEN, WHITE, h=0.24, size=7)
    text(s, Inches(tx + 1.80), Inches(pa_y + 0.10), Inches(2.0), Inches(0.20),
         "INITIATIVE", 8, bold=True, color=GREY_LABEL, font=BODY_FONT)
    chip(s, tx + tw - 1.50, pa_y + 0.10, 1.20, "Standard", COS_GREEN, WHITE, h=0.24, size=8)
    text(s, Inches(tx + 0.30), Inches(pa_y + 0.40), Inches(tw - 0.60), Inches(0.30),
         "Reduce billing-correction cycle 14→5 days",
         12, bold=True, color=NAVY_DARK, font=BODY_FONT)
    text(s, Inches(tx + 0.30), Inches(pa_y + 0.75), Inches(tw - 0.60), Inches(0.20),
         "Initiative · Billing Operations · 3 child Stories · Sponsor: Customer Care Director",
         8, italic=True, color=GREY_TEXT, font=BODY_FONT)
    # Pain + key
    text(s, Inches(tx + 0.30), Inches(pa_y + 1.00), Inches(tw - 0.60), Inches(0.22),
         "Pain Score 2.43  ·  M (5–20 SP)  ·  in Implementing for 3 weeks",
         8, bold=True, color=DARK_TEXT, font=BODY_FONT)

    # Tree connector
    line(s, Inches(tx + 0.50), Inches(pa_y + pa_h),
         Inches(tx + 0.50), Inches(pa_y + pa_h + 0.40), GREY_LABEL, weight=1.5)
    line(s, Inches(tx + 0.50), Inches(pa_y + pa_h + 0.40),
         Inches(tx + 1.20), Inches(pa_y + pa_h + 0.40), GREY_LABEL, weight=1.5)
    line(s, Inches(tx + 1.20), Inches(pa_y + pa_h + 0.40),
         Inches(tx + 1.20), Inches(pa_y + pa_h + 0.55), GREY_LABEL, weight=1.5)

    # New scope-add Story card (highlighted as accepted)
    ns_y = pa_y + pa_h + 0.60
    round_rect(s, Inches(tx + 1.20), Inches(ns_y), Inches(tw - 1.40), Inches(1.15),
               GREEN_PALE, AVANGRID_GREEN, Pt(1), 0.04)
    chip(s, tx + 1.30, ns_y + 0.08, 1.40, "NEW · scope-add", AVANGRID_GREEN, WHITE, h=0.22, size=7)
    text(s, Inches(tx + 2.80), Inches(ns_y + 0.08), Inches(2.0), Inches(0.20),
         "NEW STORY", 8, bold=True, color=GREY_LABEL, font=BODY_FONT)
    chip(s, tx + tw - 1.55, ns_y + 0.08, 1.10, "Standard", COS_GREEN, WHITE, h=0.22, size=7)
    text(s, Inches(tx + 1.30), Inches(ns_y + 0.34), Inches(tw - 1.60), Inches(0.30),
         "Add CSV export to reconciliation report",
         11, bold=True, color=NAVY_DARK, font=BODY_FONT)
    text(s, Inches(tx + 1.30), Inches(ns_y + 0.62), Inches(tw - 1.60), Inches(0.22),
         "Story under parent Initiative · Squad PM accepts within 48h · DoR fast-track",
         8, italic=True, color=GREEN_DARK, font=BODY_FONT)
    text(s, Inches(tx + 1.30), Inches(ns_y + 0.84), Inches(tw - 1.60), Inches(0.22),
         "S · 3 SP · inherits parent Standard CoS · no new Gate 2",
         8, bold=True, color=DARK_TEXT, font=BODY_FONT)

    # Existing children (faded)
    ex_y = ns_y + 1.30
    ex_kids = [
        ("STORY · A", "CIS API rate-limit upgrade",   "DONE"),
        ("STORY · B", "Genesys auto-reconcile flow",  "IMPLEMENTING"),
        ("STORY · C", "UAT data refresh script",       "BACKLOG"),
    ]
    for j, (k, t, st) in enumerate(ex_kids):
        cy = ex_y + j * 0.42
        round_rect(s, Inches(tx + 1.20), Inches(cy), Inches(tw - 1.40), Inches(0.36),
                   GREY_LIGHT, GREY_BORDER, Pt(0.5), 0.04)
        text(s, Inches(tx + 1.30), Inches(cy + 0.05), Inches(1.40), Inches(0.20),
             k, 7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, Inches(tx + 2.65), Inches(cy + 0.05), Inches(tw - 4.10), Inches(0.26),
             t, 8.5, color=DARK_TEXT, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
        chip(s, tx + tw - 1.55, cy + 0.05, 1.10, st,
             AVANGRID_GREEN if st == "DONE" else (HEADER_BLUE_DK if st == "IMPLEMENTING" else GREY_LABEL),
             WHITE, h=0.22, size=7)

    # RIGHT · Decision matrix
    rx = tx + tw + 0.20; ry = 1.30; rw = 13.33 - rx - 0.40; rh = 5.55
    round_rect(s, Inches(rx), Inches(ry), Inches(rw), Inches(rh), WHITE, GREY_BORDER, Pt(0.75), 0.04)
    rect(s, Inches(rx), Inches(ry), Inches(rw), Inches(0.42), NAVY_DARK)
    text(s, Inches(rx + 0.15), Inches(ry + 0.08), Inches(rw - 0.30), Inches(0.26),
         "DECISION MATRIX  ·  by parent Initiative status",
         11, bold=True, color=WHITE, font=BODY_FONT)

    # Header row
    hdr_y = ry + 0.55
    cols_d = [("PARENT STATUS", 1.45), ("DECISION", 1.50), ("OWNER", 1.40), ("CoS", 0.95)]
    cx_acc = rx + 0.15
    col_x_d = {}
    for lbl, w in cols_d:
        col_x_d[lbl] = cx_acc
        text(s, Inches(cx_acc), Inches(hdr_y), Inches(w), Inches(0.20),
             lbl, 7, bold=True, color=NAVY_DARK, font=BODY_FONT)
        cx_acc += w
    rect(s, Inches(rx + 0.15), Inches(hdr_y + 0.22), Inches(rw - 0.30), Emu(15000), NAVY_DARK)

    rows_d = [
        ("FUNNEL",        "REJECT shortcut", "PM",            "—",                GREY_LABEL),
        ("REVIEWING",     "REJECT shortcut", "PM",            "—",                GREY_LABEL),
        ("ANALYZING",     "REJECT shortcut", "PM + Architect","—",                GREY_LABEL),
        ("BACKLOG",       "ACCEPT (cautious)","Squad PM",     "inherits parent",  COS_ORANGE),
        ("IMPLEMENTING",  "ACCEPT (preferred)","Squad PM",    "inherits parent",  AVANGRID_GREEN),
    ]
    rry = hdr_y + 0.40
    rrh = (rh - 0.95) / len(rows_d)
    for ri, (status, decision, owner, cos_v, color) in enumerate(rows_d):
        if ri % 2 == 1:
            rect(s, Inches(rx + 0.15), Inches(rry), Inches(rw - 0.30), Inches(rrh - 0.04), GREY_LIGHT)
        # Status pill
        chip(s, rx + 0.18, rry + (rrh - 0.30) / 2, 1.30, status, color, WHITE, h=0.26, size=8)
        text(s, Inches(col_x_d["DECISION"]), Inches(rry + 0.10), Inches(1.45), Inches(rrh - 0.20),
             decision, 9, bold=True, color=color, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
        text(s, Inches(col_x_d["OWNER"]), Inches(rry + 0.10), Inches(1.35), Inches(rrh - 0.20),
             owner, 8.5, color=DARK_TEXT, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
        text(s, Inches(col_x_d["CoS"]), Inches(rry + 0.10), Inches(0.90), Inches(rrh - 0.20),
             cos_v, 8.5, italic=True, color=GREY_TEXT, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
        rry += rrh

    cos_swatch_legend(s)
    footer_section(s, 9)
    return s


# ============================================================================
# 10 · OPERATING POLICIES A1–A5
# ============================================================================
def build_pp10(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_band(s, "OPERATING POLICIES  ·  A1 – A5",
                "AUTOMATED RULES  ·  REDUCE MANUAL ROUTING  ·  ENFORCE CoS POLICIES  ·  AUTO-COMPUTE PAIN SCORE")

    # Automation table
    tbl_y = 1.30
    tbl_h = 5.55
    round_rect(s, Inches(0.40), Inches(tbl_y), Inches(12.55), Inches(tbl_h), WHITE, GREY_BORDER, Pt(0.75), 0.04)
    rect(s, Inches(0.40), Inches(tbl_y), Inches(12.55), Inches(0.42), NAVY_DARK)
    text(s, Inches(0.55), Inches(tbl_y + 0.08), Inches(12.30), Inches(0.26),
         "OPERATING RULES  ·  triggered automatically  ·  example invocations shown",
         11, bold=True, color=WHITE, font=BODY_FONT)

    rules = [
        ("A1", "Auto-Routing",
         "Issue created with Type=Demand",
         "Read VS + OPCO → set Component → notify Lead PM\nStatus: NEW → ROUTED",
         "Smart Meter analytics initiative",
         AVANGRID_GREEN),
        ("A2", "Auto-Expedite",
         "Customer Impact Tier set to P1 OR P2",
         "Set Class of Service = Expedite\nEmail Director + Sponsor",
         "CRM perf demand",
         COS_RED),
        ("A3", "Auto-Fixed-Date",
         "Type set to Regulatory",
         "Set Class of Service = Fixed-Date\nPrompt PM to set due date",
         "PUC tariff banner",
         COS_BLUE),
        ("A4", "Auto-Pain-Score",
         "BV AND UX AND Effort populated\n(Safety conditional on VS=ITBA)",
         "Compute formula based on area\nWrite to Pain Score field (read-only)",
         "Billing cycle reduction",
         AVANGRID_GREEN),
        ("A5", "Auto-Convert (at Triage)",
         "Status changes to TRIAGED + Size populated",
         "S → Story (component-routed to squad)\nM/L → Epic + 3 placeholder Stories",
         "SAP Basis patch",
         HEADER_BLUE),
    ]
    cols_a = [("ID", 0.65), ("RULE", 1.85), ("TRIGGER", 3.80), ("ACTION", 4.35), ("EXAMPLE", 1.65)]
    hdr_y = tbl_y + 0.55
    cx_acc = 0.55
    col_x_a = {}
    for lbl, w in cols_a:
        col_x_a[lbl] = cx_acc
        text(s, Inches(cx_acc), Inches(hdr_y), Inches(w), Inches(0.20),
             lbl, 8, bold=True, color=NAVY_DARK, font=BODY_FONT)
        cx_acc += w
    rect(s, Inches(0.55), Inches(hdr_y + 0.22), Inches(12.20), Emu(15000), NAVY_DARK)

    rry = hdr_y + 0.40
    rrh = (tbl_h - 0.95) / len(rules)
    for ri, (id_, name, trigger, action, last, color) in enumerate(rules):
        if ri % 2 == 1:
            rect(s, Inches(0.55), Inches(rry), Inches(12.20), Inches(rrh - 0.04), GREY_LIGHT)
        # ID circle
        oval(s, Inches(col_x_a["ID"]), Inches(rry + 0.18), Inches(0.50), Inches(0.50), color)
        text(s, Inches(col_x_a["ID"]), Inches(rry + 0.18), Inches(0.50), Inches(0.50), id_,
             10, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Rule name
        text(s, Inches(col_x_a["RULE"]), Inches(rry + 0.10), Inches(1.80), Inches(0.30),
             name, 11, bold=True, color=NAVY_DARK, font=BODY_FONT)
        text(s, Inches(col_x_a["RULE"]), Inches(rry + 0.45), Inches(1.80), Inches(rrh - 0.55),
             "● enabled", 8, bold=True, color=AVANGRID_GREEN, font=BODY_FONT)
        # Trigger
        round_rect(s, Inches(col_x_a["TRIGGER"]), Inches(rry + 0.10), Inches(3.70), Inches(rrh - 0.20),
                   JIRA_BG_LIGHT, GREY_BORDER, Pt(0.5), 0.04)
        text(s, Inches(col_x_a["TRIGGER"] + 0.10), Inches(rry + 0.16), Inches(3.55), Inches(0.18),
             "WHEN", 7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, Inches(col_x_a["TRIGGER"] + 0.10), Inches(rry + 0.32), Inches(3.55), Inches(rrh - 0.40),
             trigger, 8.5, color=DARK_TEXT, font=BODY_FONT)
        # Action
        round_rect(s, Inches(col_x_a["ACTION"]), Inches(rry + 0.10), Inches(4.25), Inches(rrh - 0.20),
                   GREEN_PALE, AVANGRID_GREEN, Pt(0.5), 0.04)
        text(s, Inches(col_x_a["ACTION"] + 0.10), Inches(rry + 0.16), Inches(4.10), Inches(0.18),
             "THEN", 7, bold=True, color=GREEN_DARK, font=BODY_FONT)
        text(s, Inches(col_x_a["ACTION"] + 0.10), Inches(rry + 0.32), Inches(4.10), Inches(rrh - 0.40),
             action, 8.5, color=DARK_TEXT, font=BODY_FONT)
        # Last run
        text(s, Inches(col_x_a["EXAMPLE"]), Inches(rry + 0.10), Inches(1.60), Inches(0.20),
             "EXAMPLE", 7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, Inches(col_x_a["EXAMPLE"]), Inches(rry + 0.28), Inches(1.60), Inches(rrh - 0.30),
             last, 8.5, color=DARK_TEXT, font=BODY_FONT)
        rry += rrh

    cos_swatch_legend(s)
    footer_section(s, 10)
    return s


# ============================================================================
# 11 · HEATHER · Product Leadership
# ============================================================================
def build_pp11(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_band(s, "HEATHER LaJOIE  ·  PRODUCT LEADERSHIP",
                "DIRECTOR · PRODUCT MANAGEMENT  ·  HUB OF PRODUCT FOR DIGITAL PRODUCTS  ·  NEW CAPABILITY FOR AVANGRID")

    # LEFT · Heather profile card
    pf_x = 0.40; pf_y = 1.30; pf_w = 4.10; pf_h = 5.55
    round_rect(s, Inches(pf_x), Inches(pf_y), Inches(pf_w), Inches(pf_h), WHITE, GREY_BORDER, Pt(0.75), 0.04)
    rect(s, Inches(pf_x), Inches(pf_y), Inches(pf_w), Inches(1.20), PURPLE_HEATHER)
    avatar(s, pf_x + pf_w / 2 - 0.55, pf_y + 0.20, 1.10, PURPLE_LIGHT, "HK")
    text(s, Inches(pf_x), Inches(pf_y + 1.40), Inches(pf_w), Inches(0.30),
         "Heather LaJoie", 16, bold=True, color=NAVY_DARK, font=BODY_FONT, align=PP_ALIGN.CENTER)
    text(s, Inches(pf_x), Inches(pf_y + 1.74), Inches(pf_w), Inches(0.22),
         "Director · Product Management",
         10, bold=True, color=PURPLE_HEATHER, font=BODY_FONT, align=PP_ALIGN.CENTER)
    text(s, Inches(pf_x), Inches(pf_y + 1.96), Inches(pf_w), Inches(0.22),
         "heather.lajoie@avangrid.com",
         8, italic=True, color=GREY_LABEL, font=BODY_FONT, align=PP_ALIGN.CENTER)

    # Quick stats
    st_y = pf_y + 2.30
    stats = [("PMs reporting", "9"), ("Gate 2 votes", "1 of 3"), ("BRP liaison", "Digital CoE"), ("Capacity reviews", "monthly")]
    rh_st = 0.65
    for i, (lbl, val) in enumerate(stats):
        sy = st_y + i * rh_st
        rect(s, Inches(pf_x + 0.20), Inches(sy), Inches(pf_w - 0.40), Inches(rh_st - 0.10), GREY_LIGHT)
        text(s, Inches(pf_x + 0.30), Inches(sy + 0.06), Inches(pf_w - 0.50), Inches(0.20),
             lbl, 8, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, Inches(pf_x + 0.30), Inches(sy + 0.24), Inches(pf_w - 0.50), Inches(0.30),
             val, 14, bold=True, color=PURPLE_HEATHER, font=BODY_FONT)

    # MIDDLE · Hub
    cx = 7.50; cy = 4.10; hub_d = 1.70
    oval(s, Inches(cx - hub_d / 2), Inches(cy - hub_d / 2), Inches(hub_d), Inches(hub_d), PURPLE_HEATHER)
    text(s, Inches(cx - hub_d / 2), Inches(cy - hub_d / 2), Inches(hub_d), Inches(hub_d),
         "HEATHER\nHUB",
         13, bold=True, color=WHITE, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    import math
    spokes = [
        (90,  "VP VITOR",       "co-chair Gate 2",  HEADER_BLUE_DK),
        (30,  "KUMAR",          "co-Director",       AVANGRID_GREEN),
        (-30, "PMs (9)",        "NEW under Heather", PURPLE_LIGHT),
        (-90, "Digital CoE BRP","liaison",           BYPASS_ORANGE),
        (-150,"Carlos",         "co-Director",       HEADER_BLUE),
        (150, "VP Vitor mtgs",  "metrics",           NAVY_HEADER),
    ]
    sp_d = 0.80; sp_r = 2.30
    for ang_deg, lbl, sub, color in spokes:
        ang = math.radians(ang_deg)
        sx = cx + sp_r * math.cos(ang) - sp_d / 2
        sy = cy - sp_r * math.sin(ang) - sp_d / 2
        line(s, Inches(cx + (hub_d / 2) * math.cos(ang)),
             Inches(cy - (hub_d / 2) * math.sin(ang)),
             Inches(cx + (sp_r - sp_d / 2) * math.cos(ang)),
             Inches(cy - (sp_r - sp_d / 2) * math.sin(ang)),
             PURPLE_LIGHT, weight=2)
        oval(s, Inches(sx), Inches(sy), Inches(sp_d), Inches(sp_d), color)
        text(s, Inches(sx), Inches(sy + 0.16), Inches(sp_d), Inches(0.24),
             lbl, 8, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, Inches(sx - 0.30), Inches(sy + sp_d + 0.04),
             Inches(sp_d + 0.60), Inches(0.30),
             sub, 8, italic=True, color=GREY_LABEL, font=BODY_FONT, align=PP_ALIGN.CENTER)

    # RIGHT · Heather's mandate
    rt_x = 10.55; rt_y = 1.30; rt_w = 13.33 - rt_x - 0.40; rt_h = 5.55
    round_rect(s, Inches(rt_x), Inches(rt_y), Inches(rt_w), Inches(rt_h), WHITE, GREY_BORDER, Pt(0.75), 0.04)
    rect(s, Inches(rt_x), Inches(rt_y), Inches(rt_w), Inches(0.42), PURPLE_HEATHER)
    text(s, Inches(rt_x + 0.15), Inches(rt_y + 0.08), Inches(rt_w - 0.30), Inches(0.26),
         "HEATHER'S MANDATE", 10, bold=True, color=WHITE, font=BODY_FONT)

    perms = [
        ("Owns the practice",  "Product Mgmt (NEW)", AVANGRID_GREEN),
        ("Votes at Gate 2",    "1 of 3 Directors", AVANGRID_GREEN),
        ("Pain Score governance","read-only (auto)", GREY_LABEL),
        ("Hires + leads PMs",  "9 NEW per VS",     AVANGRID_GREEN),
        ("Approves capacity %","with Kumar + Carlos", AVANGRID_GREEN),
        ("Liaison",            "Digital CoE BRP",  BYPASS_ORANGE),
    ]
    iy = rt_y + 0.55
    rh_p = (rt_h - 0.70) / len(perms)
    for lbl, val, color in perms:
        text(s, Inches(rt_x + 0.15), Inches(iy + 0.04), Inches(rt_w - 0.30), Inches(0.18),
             lbl, 8, bold=True, color=GREY_LABEL, font=BODY_FONT)
        round_rect(s, Inches(rt_x + 0.15), Inches(iy + 0.22), Inches(rt_w - 0.30), Inches(rh_p - 0.30),
                   GREY_LIGHT, GREY_BORDER, Pt(0.5), 0.04)
        oval(s, Inches(rt_x + 0.25), Inches(iy + rh_p / 2 - 0.03), Inches(0.16), Inches(0.16), color)
        text(s, Inches(rt_x + 0.50), Inches(iy + rh_p / 2 - 0.18), Inches(rt_w - 0.65), Inches(rh_p - 0.30),
             val, 9, bold=True, color=DARK_TEXT, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
        iy += rh_p

    footer_section(s, 11)
    return s


# ============================================================================
# 12 · ROLES · RACI matrix
# ============================================================================
def build_pp12(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_band(s, "ROLES  ·  RACI",
                "WHO IS RESPONSIBLE · ACCOUNTABLE · CONSULTED · INFORMED  ·  ACROSS THE 6 PHASE-2 ACTIVITIES")

    # Activities (rows) × Roles (cols)
    activities = [
        "Refinement (Demand → Epic → Story)",
        "Pre-estimation (BV · UX · Safety · Effort)",
        "Architecture review (NFRs · vendor · deps)",
        "Capacity allocation (% per CoS)",
        "Roadmap rolling-wave (3·6·12 mo)",
        "DoR check (handoff to Squad)",
    ]
    roles = [
        ("VP",          "Vitor",     HEADER_BLUE_DK),
        ("HEATHER",     "Product",   PURPLE_HEATHER),
        ("KUMAR / CARLOS","Director", AVANGRID_GREEN),
        ("MIDDLE MGR",  "day-to-day", GREY_LABEL),
        ("PM (per VS)", "NEW",       PURPLE_LIGHT),
        ("ARCHITECT",   "cross-VS",  BYPASS_ORANGE),
        ("SQUAD PM",    "per squad", GREEN_DARK),
        ("SPONSOR",     "business",  HEADER_BLUE),
    ]

    # RACI matrix data
    raci = [
        # Refinement
        ["I", "C", "I", "I", "A/R", "C", "C", "C"],
        # Pre-estimation
        ["I", "I", "I", "I", "A",   "R", "I", "I"],
        # Architecture review
        ["I", "C", "I", "I", "C",   "A/R","I", "I"],
        # Capacity allocation
        ["I", "C", "A", "R", "I",   "I", "I", "I"],
        # Roadmap rolling-wave
        ["I", "A", "C", "I", "R",   "C", "I", "C"],
        # DoR check
        ["I", "I", "I", "I", "C",   "C", "A/R","C"],
    ]

    tbl_y = 1.30
    tbl_h = 5.55
    label_w = 3.40
    cell_w = (12.55 - label_w) / len(roles)

    # Header row · roles
    rect(s, Inches(0.40), Inches(tbl_y), Inches(label_w), Inches(0.65), NAVY_DARK)
    text(s, Inches(0.40), Inches(tbl_y), Inches(label_w), Inches(0.65),
         "ACTIVITY", 11, bold=True, color=WHITE, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    for ci, (rname, rsub, rcolor) in enumerate(roles):
        x = 0.40 + label_w + ci * cell_w
        rect(s, Inches(x), Inches(tbl_y), Inches(cell_w - 0.04), Inches(0.65), rcolor)
        text(s, Inches(x), Inches(tbl_y + 0.06), Inches(cell_w - 0.04), Inches(0.24),
             rname, 9, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, Inches(x), Inches(tbl_y + 0.32), Inches(cell_w - 0.04), Inches(0.24),
             rsub, 7, italic=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)

    # Activity rows
    rh = (tbl_h - 0.65) / len(activities)
    ry = tbl_y + 0.65
    for ri, act in enumerate(activities):
        bg = WHITE if ri % 2 == 0 else GREY_LIGHT
        rect(s, Inches(0.40), Inches(ry), Inches(label_w), Inches(rh - 0.02), bg)
        text(s, Inches(0.55), Inches(ry + 0.10), Inches(label_w - 0.20), Inches(rh - 0.20),
             act, 9.5, bold=True, color=DARK_TEXT, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
        for ci in range(len(roles)):
            x = 0.40 + label_w + ci * cell_w
            rect(s, Inches(x), Inches(ry), Inches(cell_w - 0.04), Inches(rh - 0.02), bg)
            val = raci[ri][ci]
            # color by RACI letter
            color_map = {"R": AVANGRID_GREEN, "A": GATE_RED, "A/R": GATE_RED,
                         "C": HEADER_BLUE, "I": GREY_MUTED}
            cval = color_map.get(val, GREY_LABEL)
            d = 0.50
            oval(s, Inches(x + (cell_w - 0.04 - d) / 2), Inches(ry + (rh - 0.02 - d) / 2),
                 Inches(d), Inches(d), cval)
            text(s, Inches(x), Inches(ry + (rh - 0.02 - d) / 2), Inches(cell_w - 0.04), Inches(d),
                 val, 11, bold=True, color=WHITE, font=BODY_FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        ry += rh

    # Legend
    leg_y = 7.05
    text(s, Inches(0.40), Inches(leg_y), Inches(1.20), Inches(0.18),
         "RACI LEGEND", 8, bold=True, color=GREY_LABEL, font=BODY_FONT)
    leg = [("R · Responsible", AVANGRID_GREEN), ("A · Accountable", GATE_RED),
           ("C · Consulted", HEADER_BLUE), ("I · Informed", GREY_MUTED)]
    sx = 1.80
    for lbl, color in leg:
        oval(s, Inches(sx), Inches(leg_y + 0.02), Inches(0.16), Inches(0.16), color)
        text(s, Inches(sx + 0.22), Inches(leg_y), Inches(2.0), Inches(0.18),
             lbl, 8, color=TEXT, font=BODY_FONT)
        sx += 2.30

    footer_section(s, 12)
    return s


# ============================================================================
# Driver
# ============================================================================
def main():
    print("Opening:", SRC)
    prs = Presentation(SRC)
    n_before = len(prs.slides)

    print("Building Preparation & Planning v3 · consulting-deck quality (12 slides)...")
    build_pp01(prs);  print("   01 PHASE ENTRY")
    build_pp02(prs);  print("   02 JIRA WORKFLOW")
    build_pp03(prs);  print("   03 REFINEMENT BOARD")
    build_pp04(prs);  print("   04 STORY CARD · REQ · Billing")
    build_pp05(prs);  print("   05 PRE-ESTIMATION + PAIN SCORE")
    build_pp06(prs);  print("   06 CAPACITY ALLOCATION")
    build_pp07(prs);  print("   07 ROADMAP")
    build_pp08(prs);  print("   08 SIZE PATHS S/M/L")
    build_pp09(prs);  print("   09 EXISTING-INITIATIVE SHORTCUT")
    build_pp10(prs);  print("   10 AUTOMATIONS A1-A5")
    build_pp11(prs);  print("   11 HEATHER · PRODUCT")
    build_pp12(prs);  print("   12 ROLES · RACI")

    n_after = len(prs.slides)
    print(f"\nSlides: {n_before} → {n_after}")

    prs.save(DST)

    # de-dup zip
    tmp = DST + ".clean"
    with zipfile.ZipFile(DST, "r") as zin:
        seen = set(); items = []
        for info in zin.infolist():
            if info.filename in seen: continue
            seen.add(info.filename); items.append(info)
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in items:
                zout.writestr(info, zin.read(info.filename))
    shutil.move(tmp, DST)
    print(f"\n✓ Saved (deduped): {DST}")


if __name__ == "__main__":
    main()
