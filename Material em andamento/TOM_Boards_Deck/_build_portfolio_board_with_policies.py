"""
Standalone editable slide: Portfolio Board with step-level policies populated.
Combines visual template from image 1 (kanban columns + gates) with content
from image 2 (per-column policies, owners, SLAs, outputs).
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

OUT = "/Users/gustavohenriquecastellano/Downloads/Base Apresentações Projeto Digital Product/Portfolio Board with Step-Level Policies.pptx"

# Palette matching the reference image style (dark navy header)
NAVY_HEADER = RGBColor(0x2D, 0x4E, 0x7C)
NAVY_DARK = RGBColor(0x0F, 0x2A, 0x47)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
TEXT = RGBColor(0x17, 0x2B, 0x4D)
GREY_LIGHT = RGBColor(0xF2, 0xF6, 0xFA)
GREY_BORDER = RGBColor(0xC9, 0xD2, 0xDD)
GREY_TEXT = RGBColor(0x52, 0x60, 0x77)
GREY_LABEL = RGBColor(0x6B, 0x77, 0x8C)
GATE_RED = RGBColor(0xB0, 0x35, 0x35)
BRP_GREEN_LIGHT = RGBColor(0xE6, 0xF2, 0xEB)
BRP_GREEN = RGBColor(0x00, 0x6B, 0x3F)
BYPASS_ORANGE_LIGHT = RGBColor(0xFD, 0xEC, 0xD4)
BYPASS_ORANGE = RGBColor(0xF0, 0x86, 0x29)
COS_GREEN = RGBColor(0x4F, 0x91, 0x4F)
COS_BLUE = RGBColor(0x3B, 0x6E, 0x8F)
COS_RED = RGBColor(0xB0, 0x35, 0x35)
COS_ORANGE = RGBColor(0xE9, 0x8B, 0x2F)


prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])


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


def text(slide, x, y, w, h, txt, size=11, bold=False, italic=False, color=TEXT,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.margin_left = Emu(36000); tf.margin_right = Emu(36000)
    tf.margin_top = Emu(18000); tf.margin_bottom = Emu(18000)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = txt
    r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = color
    return tb


# ============================================================================
# HEADER (dark navy band)
# ============================================================================
rect(slide, 0, 0, prs.slide_width, Inches(0.95), NAVY_HEADER)
text(slide, Inches(0.4), Inches(0.13), Inches(10), Inches(0.42),
     "PORTFOLIO BOARD", 22, bold=True, color=WHITE)
text(slide, Inches(0.4), Inches(0.55), Inches(11), Inches(0.3),
     "CARDS ARE INITIATIVES  ·  GOVERNANCE GATES AND BRP SHOWN IN CONTEXT",
     11, bold=True, color=WHITE)
text(slide, Inches(11.3), Inches(0.30), Inches(1.7), Inches(0.4),
     "NTT  DATA", 16, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)


# ============================================================================
# Sub-context line
# ============================================================================
text(slide, Inches(0.4), Inches(1.05), Inches(13), Inches(0.3),
     "Example  ·  Customer Services Portfolio Board  ·  Governance gates marked in red",
     11, italic=True, color=GREY_TEXT, align=PP_ALIGN.CENTER)


# ============================================================================
# BRP banner (light green)
# ============================================================================
brp_y = Inches(1.45)
round_rect(slide, Inches(0.4), brp_y, Inches(12.55), Inches(0.34),
           BRP_GREEN_LIGHT, BRP_GREEN, Pt(1), 0.04)
text(slide, Inches(0.4), brp_y, Inches(12.55), Inches(0.34),
     "▲  BIG ROOM PLANNING  —  quarterly  ·  ½ day  ·  realigns Backlog reorder + Implementing commitments across both Portfolio Boards",
     10, bold=True, italic=True, color=BRP_GREEN,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ============================================================================
# Gate markers (above columns 1-2 boundary and 3-4 boundary)
# ============================================================================
# 6 columns: Funnel, Reviewing, Analyzing, Backlog, Implementing, Done
# Total width: 12.55" - 6 columns of ~2.05" each + 5 gaps of 0.05"
col_w_in = (12.55 - 5 * 0.05) / 6  # ~2.075"
col_x_start = 0.4
gate_y = Inches(1.95)

# Gate 1 marker (between Funnel [col 0] and Reviewing [col 1])
gate1_x = col_x_start + col_w_in + 0.025 - 1.0  # centered above the boundary, label to the left
text(slide, Inches(gate1_x), gate_y, Inches(2.0), Inches(0.22),
     "▲  GATE 1 · TRIAGE  ·  weekly", 10, bold=True, color=GATE_RED,
     align=PP_ALIGN.CENTER)

# Gate 2 marker (between Analyzing [col 2] and Backlog [col 3])
gate2_x = col_x_start + 3 * col_w_in + 3 * 0.05 + 0.025 - 1.0
text(slide, Inches(gate2_x), gate_y, Inches(2.0), Inches(0.22),
     "▲  GATE 2 · PORTFOLIO  ·  monthly", 10, bold=True, color=GATE_RED,
     align=PP_ALIGN.CENTER)


# ============================================================================
# 6 KANBAN COLUMNS with policies
# ============================================================================
col_y = Inches(2.30)
col_h = Inches(4.20)
header_h = Inches(0.4)

columns_data = [
    # (title, body_blocks)
    ("Funnel", [
        ("DESCRIPTION", "Idea submitted with problem statement"),
        ("OWNER", "Sponsor"),
        ("REVIEWER", "Demand Manager"),
        ("INPUT", "Demand placeholder · Title · Description · BV"),
    ]),
    ("Reviewing", [
        ("DESCRIPTION", "Business case refined · classification applied"),
        ("OWNER", "Sponsor + Demand Manager"),
        ("SLA", "2 weeks"),
        ("OUTPUT", "Refined business case · CoS guess"),
    ]),
    ("Analyzing", [
        ("DESCRIPTION", "Sized · scored · feasibility checked"),
        ("OWNER", "Demand Manager + Architect"),
        ("OUTPUT", "Pain Score · T-shirt effort · Class of Service"),
        ("ARTIFACT", "Discovery Report · Concept Prototype"),
    ]),
    ("Backlog", [
        ("DESCRIPTION", "Approved · awaiting capacity"),
        ("OWNER", "Portfolio Manager"),
        ("SEQUENCING", "by Class of Service + Pain Score"),
        ("STATUS", "Ready_For_Exec"),
    ]),
    ("Implementing", [
        ("DESCRIPTION", "Assigned to Value Stream Team"),
        ("OWNER", "Team PM"),
        ("TRACKED AT", "Layer 0 VP Dashboard"),
        ("FLOW", "Pulled into Squad Kanban (Phases 0–4)"),
    ]),
    ("Done", [
        ("DESCRIPTION", "Value realized · post-implementation review"),
        ("OWNER", "Sponsor + PM"),
        ("OUTPUT", "Outcome metrics captured"),
        ("FEEDBACK", "Lessons learned → next BRP cycle"),
    ]),
]

for i, (title, blocks) in enumerate(columns_data):
    cx = Inches(col_x_start + i * (col_w_in + 0.05))
    cw = Inches(col_w_in)

    # Header
    rect(slide, cx, col_y, cw, header_h, NAVY_DARK)
    text(slide, cx, col_y, cw, header_h,
         title, 14, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Body
    body_y = col_y + header_h
    body_h = col_h - header_h
    rect(slide, cx, body_y, cw, body_h, GREY_LIGHT, GREY_BORDER, Pt(0.5))

    # Content
    by = body_y + Inches(0.15)
    for label, content in blocks:
        text(slide, cx + Inches(0.12), by, cw - Inches(0.24), Inches(0.18),
             label, 7, bold=True, color=GREY_LABEL)
        text(slide, cx + Inches(0.12), by + Inches(0.16), cw - Inches(0.24), Inches(0.50),
             content, 8.5, color=TEXT)
        by += Inches(0.85)


# ============================================================================
# Vertical RED gate dividers (between cols 0-1 and 2-3)
# ============================================================================
# Gate 1 — between Funnel (0) and Reviewing (1)
gate1_div_x = Inches(col_x_start + col_w_in + 0.005)
rect(slide, gate1_div_x, col_y, Inches(0.04), col_h, GATE_RED)

# Gate 2 — between Analyzing (2) and Backlog (3)
gate2_div_x = Inches(col_x_start + 3 * col_w_in + 3 * 0.05 - 0.04 + 0.005)
rect(slide, gate2_div_x, col_y, Inches(0.04), col_h, GATE_RED)


# ============================================================================
# Bottom: Bypass note + Class of Service legend
# ============================================================================
bypass_y = Inches(6.62)
round_rect(slide, Inches(0.4), bypass_y, Inches(12.55), Inches(0.34),
           BYPASS_ORANGE_LIGHT, BYPASS_ORANGE, Pt(1), 0.04)
text(slide, Inches(0.4), bypass_y, Inches(12.55), Inches(0.34),
     "S / M demands BYPASS this board — they enter the Squad's Team Kanban directly. Only L / XL Initiatives flow through the Portfolio Board (Gate 2 approval required).",
     9, bold=True, italic=True, color=BYPASS_ORANGE,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# Class of Service legend
cos_y = Inches(7.05)
text(slide, Inches(0.4), cos_y, Inches(2.0), Inches(0.22),
     "CLASS OF SERVICE", 9, bold=True, color=GREY_LABEL)
swatches = [("Standard", COS_GREEN), ("Fixed-Date", COS_BLUE),
            ("Expedite", COS_RED), ("Intangible", COS_ORANGE)]
sx = 2.4
for label, color in swatches:
    rect(slide, Inches(sx), cos_y + Inches(0.02), Inches(0.32), Inches(0.18), color)
    text(slide, Inches(sx + 0.4), cos_y, Inches(1.5), Inches(0.22),
         label, 9, color=TEXT)
    sx += 1.9


# ============================================================================
# Footer (OwnIT! + Avangrid placeholders + copyright)
# ============================================================================
round_rect(slide, Inches(11.7), Inches(7.0), Inches(0.65), Inches(0.4),
           BYPASS_ORANGE, radius=0.15)
text(slide, Inches(11.7), Inches(7.0), Inches(0.65), Inches(0.4),
     "OwnIT!", 9, bold=True, color=WHITE,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
round_rect(slide, Inches(12.4), Inches(7.0), Inches(0.65), Inches(0.4),
           BRP_GREEN, radius=0.15)
text(slide, Inches(12.4), Inches(7.0), Inches(0.65), Inches(0.4),
     "Avangrid", 8, bold=True, color=WHITE,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

text(slide, Inches(0.4), Inches(7.32), Inches(8), Inches(0.18),
     "© 2026 NTT DATA, Inc.", 8, color=GREY_LABEL)


prs.save(OUT)
print(f"Saved: {OUT}")
