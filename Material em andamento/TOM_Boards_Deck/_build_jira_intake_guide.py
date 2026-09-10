"""
Editable PPTX Jira Intake Form Guide.
Avangrid styling (green header + OwnIT branding).

Slides:
  1. Form mockup
  2-9. Field detail (one or paired fields per slide, visually rich)
  10. Pain Score origin + concept
  11. Pain Score formula + auto-computation
  12. POST-SUBMISSION FLOW (visual · boards · status transitions · auto-conversion)
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

OUT = "/Users/gustavohenriquecastellano/Downloads/Base Apresentações Projeto Digital Product/Jira Intake Form Guide.pptx"

# Palette
AVA_GREEN = RGBColor(0x00, 0x6B, 0x3F)
AVA_GREEN_LIGHT = RGBColor(0xE6, 0xF2, 0xEB)
AVA_GREEN_DARK = RGBColor(0x00, 0x4D, 0x2E)
AVA_ORANGE = RGBColor(0xF0, 0x86, 0x29)
AVA_ORANGE_LIGHT = RGBColor(0xFD, 0xEC, 0xD4)
NAVY = RGBColor(0x0F, 0x2A, 0x47)
TEXT = RGBColor(0x17, 0x2B, 0x4D)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREY_BG = RGBColor(0xF4, 0xF5, 0xF7)
GREY_SOFT = RGBColor(0xFA, 0xFB, 0xFC)
GREY_BORDER = RGBColor(0xDF, 0xE1, 0xE6)
GREY_PLACEHOLDER = RGBColor(0xA5, 0xAD, 0xBA)
GREY_LABEL = RGBColor(0x6B, 0x77, 0x8C)
GREY_TEXT = RGBColor(0x42, 0x52, 0x6E)
RED = RGBColor(0xDE, 0x35, 0x0B)
RED_LIGHT = RGBColor(0xFF, 0xE5, 0xDD)
ORANGE = RGBColor(0xFF, 0x8B, 0x00)
YELLOW = RGBColor(0xFF, 0xC1, 0x07)
PURPLE = RGBColor(0x5C, 0x2A, 0x91)
PURPLE_LIGHT = RGBColor(0xEE, 0xE5, 0xF8)
PINK = RGBColor(0xB8, 0x35, 0x67)
JIRA_BLUE = RGBColor(0x07, 0x47, 0xA6)
JIRA_BLUE_LIGHT = RGBColor(0xDE, 0xEB, 0xFF)
ITBA_GREY = RGBColor(0x80, 0x80, 0x80)


prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)


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


def text(slide, x, y, w, h, txt, size=11, bold=False, italic=False, color=NAVY,
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


def header_band(slide, slide_section_title, slide_section_subtitle):
    rect(slide, 0, 0, prs.slide_width, Inches(0.95), WHITE)
    text(slide, Inches(0.4), Inches(0.15), Inches(10), Inches(0.45),
         slide_section_title, 22, bold=True, color=AVA_GREEN)
    text(slide, Inches(0.4), Inches(0.55), Inches(10), Inches(0.3),
         slide_section_subtitle, 13, bold=True, color=NAVY)
    round_rect(slide, Inches(11.2), Inches(0.15), Inches(0.85), Inches(0.55),
               AVA_ORANGE, radius=0.15)
    text(slide, Inches(11.2), Inches(0.15), Inches(0.85), Inches(0.55),
         "OwnIT!", 11, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    round_rect(slide, Inches(12.15), Inches(0.15), Inches(0.85), Inches(0.55),
               AVA_GREEN, radius=0.15)
    text(slide, Inches(12.15), Inches(0.15), Inches(0.85), Inches(0.55),
         "Avangrid", 10, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    rect(slide, 0, Inches(0.92), prs.slide_width, Emu(19050), AVA_GREEN)


def field_slide_template(slide, field_name, field_subtitle, required_badge, ftype, accent=AVA_GREEN):
    """Returns (left_x, right_x, content_y, content_w_each, content_h) for body."""
    header_band(slide, f"FIELD GUIDE  ·  {field_name.upper()}",
                field_subtitle)
    # Top strip with badge
    strip_y = Inches(1.15)
    round_rect(slide, Inches(0.4), strip_y, Inches(12.55), Inches(0.55),
               accent, radius=0.04)
    text(slide, Inches(0.55), strip_y + Inches(0.12), Inches(8), Inches(0.32),
         field_name, 16, bold=True, color=WHITE, anchor=MSO_ANCHOR.MIDDLE)
    text(slide, Inches(0.55), strip_y + Inches(0.04), Inches(8), Inches(0.18),
         f"TYPE  ·  {ftype}", 7, bold=True, color=AVA_GREEN_LIGHT)
    # Required badge
    if required_badge == "MANDATORY":
        round_rect(slide, Inches(11.4), strip_y + Inches(0.15), Inches(1.4), Inches(0.26),
                   RED, radius=0.4)
        text(slide, Inches(11.4), strip_y + Inches(0.15), Inches(1.4), Inches(0.26),
             "● MANDATORY", 9, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    elif required_badge == "CONDITIONAL":
        round_rect(slide, Inches(11.4), strip_y + Inches(0.15), Inches(1.4), Inches(0.26),
                   YELLOW, radius=0.4)
        text(slide, Inches(11.4), strip_y + Inches(0.15), Inches(1.4), Inches(0.26),
             "◐ CONDITIONAL", 9, bold=True, color=NAVY,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    elif required_badge == "AUTO":
        round_rect(slide, Inches(11.4), strip_y + Inches(0.15), Inches(1.4), Inches(0.26),
                   PURPLE, radius=0.4)
        text(slide, Inches(11.4), strip_y + Inches(0.15), Inches(1.4), Inches(0.26),
             "○ READ-ONLY", 9, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def bottom_strip(slide, why_text, triggers_text, accent=AVA_GREEN):
    by = Inches(6.85)
    round_rect(slide, Inches(0.4), by, Inches(6.18), Inches(0.45),
               AVA_GREEN_LIGHT, AVA_GREEN, Pt(1), 0.05)
    text(slide, Inches(0.55), by + Inches(0.04), Inches(5.9), Inches(0.18),
         "WHY IT MATTERS", 8, bold=True, color=AVA_GREEN_DARK)
    text(slide, Inches(0.55), by + Inches(0.20), Inches(5.9), Inches(0.22),
         why_text, 9, color=TEXT)

    round_rect(slide, Inches(6.78), by, Inches(6.18), Inches(0.45),
               AVA_ORANGE_LIGHT, AVA_ORANGE, Pt(1), 0.05)
    text(slide, Inches(6.93), by + Inches(0.04), Inches(5.9), Inches(0.18),
         "TRIGGERS / DRIVES", 8, bold=True, color=AVA_ORANGE)
    text(slide, Inches(6.93), by + Inches(0.20), Inches(5.9), Inches(0.22),
         f"▸ {triggers_text}", 9, italic=True, bold=True, color=AVA_ORANGE)


# ============================================================================
# SLIDE 1 — JIRA REQUEST FORM MOCKUP
# ============================================================================
slide1 = prs.slides.add_slide(prs.slide_layouts[6])
header_band(slide1, "TO-BE DEMAND FLOW",
            "ONE JIRA FRONT DOOR ROUTES TO 5 ITBA + 4 CS DVS + 1 ENABLING TEAM")

form_x = Inches(0.4); form_y = Inches(1.05); form_w = Inches(12.5); form_h = Inches(6.2)
round_rect(slide1, form_x, form_y, form_w, form_h, WHITE, GREY_BORDER, Pt(0.5), 0.02)
rect(slide1, form_x, form_y, form_w, Inches(0.32), AVA_GREEN)
text(slide1, form_x + Inches(0.15), form_y + Inches(0.04), Inches(8), Inches(0.24),
     "▦  Jira Software  /  Avangrid · Digital Products  /  Create Issue", 10, bold=True, color=WHITE)
rect(slide1, form_x, form_y + Inches(0.32), form_w, Inches(0.32), GREY_BG)
round_rect(slide1, form_x + Inches(0.15), form_y + Inches(0.36), Inches(1.2), Inches(0.24),
           WHITE, GREY_BORDER, Pt(0.5))
text(slide1, form_x + Inches(0.2), form_y + Inches(0.38), Inches(1.1), Inches(0.2),
     "★ Demand ▾", 9, bold=True, color=AVA_GREEN)
text(slide1, form_x + Inches(1.5), form_y + Inches(0.38), Inches(4), Inches(0.2),
     "DGTL-PROD-NNNN", 9, bold=True, color=GREY_LABEL)

col_l_x = form_x + Inches(0.2); col_l_w = Inches(7.2)
col_r_x = form_x + Inches(7.55); col_r_w = Inches(4.8)
content_y = form_y + Inches(0.78)

text(slide1, col_l_x, content_y, col_l_w, Inches(0.22), "WHAT", 12, bold=True, color=AVA_GREEN)
rect(slide1, col_l_x, content_y + Inches(0.24), col_l_w, Emu(9525), GREY_BORDER)
fy = content_y + Inches(0.34)
text(slide1, col_l_x, fy, col_l_w, Inches(0.18), "Title *  ·  outcome-based, not task-based", 9, bold=True, color=TEXT)
round_rect(slide1, col_l_x, fy + Inches(0.18), col_l_w, Inches(0.26), GREY_SOFT, GREY_BORDER, Pt(0.5))
text(slide1, col_l_x + Inches(0.08), fy + Inches(0.20), col_l_w, Inches(0.22),
     "e.g. \"Reduce billing-correction cycle from 14 to 5 days\"", 9, italic=True, color=GREY_PLACEHOLDER)
fy += Inches(0.52)
text(slide1, col_l_x, fy, col_l_w, Inches(0.18), "Description *  ·  describe the problem, not the solution", 9, bold=True, color=TEXT)
round_rect(slide1, col_l_x, fy + Inches(0.18), col_l_w, Inches(0.5), GREY_SOFT, GREY_BORDER, Pt(0.5))
text(slide1, col_l_x + Inches(0.08), fy + Inches(0.20), col_l_w, Inches(0.46),
     "Context · current pain · who is affected · what blocks today", 9, italic=True, color=GREY_PLACEHOLDER)
fy += Inches(0.78)
text(slide1, col_l_x, fy, Inches(3.5), Inches(0.18), "Business Value *  ·  why it matters", 9, bold=True, color=TEXT)
round_rect(slide1, col_l_x, fy + Inches(0.18), Inches(3.5), Inches(0.32), GREY_SOFT, GREY_BORDER, Pt(0.5))
text(slide1, col_l_x + Inches(3.7), fy, Inches(3.5), Inches(0.18), "Acceptance Criteria *  ·  measurable", 9, bold=True, color=TEXT)
round_rect(slide1, col_l_x + Inches(3.7), fy + Inches(0.18), Inches(3.5), Inches(0.32), GREY_SOFT, GREY_BORDER, Pt(0.5))

fy += Inches(0.65)
text(slide1, col_l_x, fy, col_l_w, Inches(0.22), "ROUTING  ·  selecting Value Stream auto-routes the ticket", 12, bold=True, color=AVA_GREEN)
rect(slide1, col_l_x, fy + Inches(0.24), col_l_w, Emu(9525), GREY_BORDER)
fy += Inches(0.34)
text(slide1, col_l_x, fy, col_l_w, Inches(0.18), "Value Stream destination *", 9, bold=True, color=TEXT)
fy += Inches(0.20)
round_rect(slide1, col_l_x, fy, col_l_w, Inches(0.30), AVA_ORANGE_LIGHT, AVA_ORANGE, Pt(2), 0.05)
text(slide1, col_l_x + Inches(0.1), fy + Inches(0.04), col_l_w, Inches(0.22),
     "▼  Select Value Stream / Team…", 10, bold=True, color=TEXT)
fy += Inches(0.40)
panel_w = (col_l_w - Inches(0.15)) / 2

itba_x = col_l_x
round_rect(slide1, itba_x, fy, panel_w, Inches(2.0), GREY_BG, ITBA_GREY, Pt(1.5), 0.04)
round_rect(slide1, itba_x, fy, panel_w, Inches(0.28), ITBA_GREY, radius=0.04)
text(slide1, itba_x + Inches(0.1), fy + Inches(0.04), panel_w - Inches(0.2), Inches(0.22),
     "ITBA · Carlos Garbin · 5 teams", 10, bold=True, color=WHITE)
itba_options = ["○  Field Operations", "○  Engineering & Asset Mgmt",
                "○  Network Operations & Control", "○  Corporate & Regulatory",
                "○  Iberdrola Corp Services"]
oy = fy + Inches(0.34)
for opt in itba_options:
    text(slide1, itba_x + Inches(0.12), oy, panel_w - Inches(0.2), Inches(0.22),
         opt, 9, color=TEXT)
    oy += Inches(0.28)
text(slide1, itba_x + Inches(0.12), fy + Inches(1.78), panel_w - Inches(0.2), Inches(0.18),
     "→ ITBA Portfolio Board", 8, italic=True, bold=True, color=GREY_LABEL)

cs_x = col_l_x + panel_w + Inches(0.15)
round_rect(slide1, cs_x, fy, panel_w, Inches(2.0), JIRA_BLUE_LIGHT, JIRA_BLUE, Pt(1.5), 0.04)
round_rect(slide1, cs_x, fy, panel_w, Inches(0.28), JIRA_BLUE, radius=0.04)
text(slide1, cs_x + Inches(0.1), fy + Inches(0.04), panel_w - Inches(0.2), Inches(0.22),
     "CS · Kumar Jaladanki · 4 teams + Enabling", 10, bold=True, color=WHITE)
cs_options = [("●", "Meter & Device Operations", JIRA_BLUE, True),
              ("○", "Billing Operations", TEXT, False),
              ("○", "Revenue & Collections", TEXT, False),
              ("○", "Customer Care & Engagement", TEXT, False),
              ("○", "Enabling Team — Foundation Services", PURPLE, False)]
oy = fy + Inches(0.34)
for radio, opt, c, sel in cs_options:
    text(slide1, cs_x + Inches(0.12), oy, panel_w - Inches(0.2), Inches(0.22),
         f"{radio}  {opt}", 9, color=c, bold=sel)
    oy += Inches(0.28)

text(slide1, col_r_x, content_y, col_r_w, Inches(0.22), "CLASSIFICATION", 12, bold=True, color=AVA_GREEN)
rect(slide1, col_r_x, content_y + Inches(0.24), col_r_w, Emu(9525), GREY_BORDER)
fy = content_y + Inches(0.34)
text(slide1, col_r_x, fy, col_r_w, Inches(0.18), "Class of Service *  ·  swimlane assignment", 9, bold=True, color=TEXT)
fy += Inches(0.22)
cos_w = (col_r_w - Inches(0.06)) / 2
for i, lbl in enumerate(["Standard", "Fixed-Date", "Expedite", "Intangible"]):
    cx = col_r_x + (i % 2) * (cos_w + Inches(0.06))
    cy = fy + (i // 2) * Inches(0.24)
    text(slide1, cx + Inches(0.4), cy, cos_w, Inches(0.22), f"○  {lbl}", 9, color=TEXT)
fy += Inches(0.55)

text(slide1, col_r_x, fy, Inches(2.2), Inches(0.18), "Type *", 9, bold=True, color=TEXT)
round_rect(slide1, col_r_x, fy + Inches(0.18), Inches(2.2), Inches(0.24), GREY_SOFT, GREY_BORDER, Pt(0.5))
text(slide1, col_r_x + Inches(0.08), fy + Inches(0.20), Inches(2.1), Inches(0.22), "Demand", 9, bold=True, color=TEXT)
text(slide1, col_r_x + Inches(2.4), fy, Inches(2.4), Inches(0.18), "Size *", 9, bold=True, color=TEXT)
round_rect(slide1, col_r_x + Inches(2.4), fy + Inches(0.18), Inches(2.4), Inches(0.24), GREY_SOFT, GREY_BORDER, Pt(0.5))
text(slide1, col_r_x + Inches(2.48), fy + Inches(0.20), Inches(2.3), Inches(0.22), "▼ S · M · L · XL", 9, color=TEXT)
fy += Inches(0.5)
text(slide1, col_r_x, fy, Inches(2.2), Inches(0.18), "Work Type *", 9, bold=True, color=TEXT)
round_rect(slide1, col_r_x, fy + Inches(0.18), Inches(2.2), Inches(0.24), GREY_SOFT, GREY_BORDER, Pt(0.5))
text(slide1, col_r_x + Inches(0.08), fy + Inches(0.20), Inches(2.1), Inches(0.22), "▼ BAU · Build · Upgrade", 9, color=TEXT)
text(slide1, col_r_x + Inches(2.4), fy, Inches(2.4), Inches(0.18), "Run / Grow / Transform *", 9, bold=True, color=TEXT)
round_rect(slide1, col_r_x + Inches(2.4), fy + Inches(0.18), Inches(2.4), Inches(0.24), GREY_SOFT, GREY_BORDER, Pt(0.5))
text(slide1, col_r_x + Inches(2.48), fy + Inches(0.20), Inches(2.3), Inches(0.22), "▼ Run · Grow · Transform", 9, color=TEXT)
fy += Inches(0.5)

text(slide1, col_r_x, fy, col_r_w, Inches(0.18),
     "Customer Impact Tier  ·  CS only · drives Auto-Expedite", 9, bold=True, color=TEXT)
fy += Inches(0.20)
tier_w = (col_r_w - Inches(0.18)) / 4
for i, (lbl, fill, txc) in enumerate([("P1 Critical", RED, WHITE), ("P2 High", ORANGE, WHITE),
                                      ("P3 Medium", YELLOW, NAVY), ("P4 Low", AVA_GREEN, WHITE)]):
    tx = col_r_x + i * (tier_w + Inches(0.06))
    text(slide1, tx, fy, tier_w, Inches(0.22), lbl, 8, bold=True, color=GREY_LABEL,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
fy += Inches(0.35)

text(slide1, col_r_x, fy, col_r_w, Inches(0.18),
     "Regulatory flag  ·  if true · auto-Fixed-Date + Compliance sponsor", 9, bold=True, color=TEXT)
fy += Inches(0.20)
text(slide1, col_r_x, fy, col_r_w, Inches(0.22), "☐ Yes (regulatory mandate)", 9, color=TEXT)
fy += Inches(0.36)

text(slide1, col_r_x, fy, col_r_w, Inches(0.22), "CONTEXT", 12, bold=True, color=AVA_GREEN)
rect(slide1, col_r_x, fy + Inches(0.24), col_r_w, Emu(9525), GREY_BORDER)
fy += Inches(0.34)
ctx_w = (col_r_w - Inches(0.12)) / 3
for i, (lbl, val) in enumerate([("OPCO *", "▼ UIL · CMP · NY"),
                                ("Sponsor *", "@ User picker"),
                                ("Due Date", "📅 YYYY-MM-DD")]):
    cx = col_r_x + i * (ctx_w + Inches(0.06))
    text(slide1, cx, fy, ctx_w, Inches(0.18), lbl, 9, bold=True, color=TEXT)
    round_rect(slide1, cx, fy + Inches(0.18), ctx_w, Inches(0.24), GREY_SOFT, GREY_BORDER, Pt(0.5))
    text(slide1, cx + Inches(0.06), fy + Inches(0.20), ctx_w, Inches(0.22), val, 8, color=TEXT)
fy += Inches(0.55)

round_rect(slide1, col_r_x, fy, col_r_w, Inches(0.55), AVA_GREEN_LIGHT, AVA_GREEN, Pt(2))
text(slide1, col_r_x, fy + Inches(0.05), col_r_w, Inches(0.18),
     "PAIN SCORE  ·  auto-computed", 8, bold=True, color=AVA_GREEN_DARK, align=PP_ALIGN.CENTER)
text(slide1, col_r_x, fy + Inches(0.22), col_r_w, Inches(0.30),
     "H(x) = (BV·w₁ + UX·w₂ + Safety·w₃) / Effort",
     11, bold=True, color=AVA_GREEN_DARK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

foot_y = form_y + form_h - Inches(0.32)
round_rect(slide1, form_x + Inches(0.2), foot_y, form_w - Inches(0.4), Inches(0.24),
           AVA_ORANGE_LIGHT, AVA_ORANGE, Pt(1), 0.1)
text(slide1, form_x, foot_y, form_w, Inches(0.24),
     "⚠  IT Now is reserved for incidents only — not for new demand",
     9, bold=True, color=AVA_ORANGE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ============================================================================
# SLIDE 2 — TITLE field (full slide detail)
# ============================================================================
slide2 = prs.slides.add_slide(prs.slide_layouts[6])
field_slide_template(slide2, "Title", "Outcome-based summary line that becomes the card label",
                     "MANDATORY", "Free text · max 120 chars · plain text")

# Body — split in 2 columns
body_y = Inches(1.85)
body_h = Inches(4.85)

# Left: Mock-up + how to fill
round_rect(slide2, Inches(0.4), body_y, Inches(6.18), body_h, WHITE, GREY_BORDER, Pt(0.5), 0.03)
text(slide2, Inches(0.55), body_y + Inches(0.1), Inches(6), Inches(0.22),
     "HOW IT LOOKS IN JIRA", 9, bold=True, color=GREY_LABEL)
# Form field mock
text(slide2, Inches(0.55), body_y + Inches(0.4), Inches(6), Inches(0.22),
     "Title *", 10, bold=True, color=TEXT)
round_rect(slide2, Inches(0.55), body_y + Inches(0.62), Inches(5.88), Inches(0.32),
           GREY_SOFT, GREY_BORDER, Pt(0.5))
text(slide2, Inches(0.65), body_y + Inches(0.66), Inches(5.7), Inches(0.24),
     "Type a clear, outcome-based summary…", 9, italic=True, color=GREY_PLACEHOLDER)

text(slide2, Inches(0.55), body_y + Inches(1.15), Inches(6), Inches(0.22),
     "GOOD EXAMPLES (✓)", 10, bold=True, color=AVA_GREEN)
good = [
    "Reduce billing-correction cycle from 14 to 5 days",
    "Add Spanish language to IVR menu for UIL OPCO",
    "Eliminate manual intervention on rate-case Q3 release",
    "Improve CRM page load from 8s to <2s",
]
gy = body_y + Inches(1.4)
for ex in good:
    round_rect(slide2, Inches(0.55), gy, Inches(5.88), Inches(0.30),
               AVA_GREEN_LIGHT, AVA_GREEN, Pt(1), 0.1)
    text(slide2, Inches(0.65), gy + Inches(0.04), Inches(5.7), Inches(0.24),
         "✓  " + ex, 9, color=TEXT)
    gy += Inches(0.36)

text(slide2, Inches(0.55), gy + Inches(0.1), Inches(6), Inches(0.22),
     "BAD EXAMPLES (✗) — get bounced at triage", 10, bold=True, color=RED)
bad = [
    "Fix system  ·  System is broken  ·  Update CRM",
    "Improve billing  ·  Make it faster",
]
gy += Inches(0.4)
for ex in bad:
    round_rect(slide2, Inches(0.55), gy, Inches(5.88), Inches(0.30),
               RED_LIGHT, RED, Pt(1), 0.1)
    text(slide2, Inches(0.65), gy + Inches(0.04), Inches(5.7), Inches(0.24),
         "✗  " + ex, 9, color=TEXT)
    gy += Inches(0.36)

# Right: Detail + rules
round_rect(slide2, Inches(6.78), body_y, Inches(6.18), body_h, WHITE, GREY_BORDER, Pt(0.5), 0.03)
text(slide2, Inches(6.93), body_y + Inches(0.1), Inches(6), Inches(0.22),
     "WHAT IT IS", 9, bold=True, color=AVA_GREEN)
text(slide2, Inches(6.93), body_y + Inches(0.32), Inches(5.9), Inches(0.7),
     "A short outcome statement that captures the user-visible result of the demand. The Title becomes the card's display label across all boards (Portfolio, Squad, Enterprise) and is the FIRST thing executives see when scanning the backlog.",
     10, color=TEXT)

text(slide2, Inches(6.93), body_y + Inches(1.10), Inches(6), Inches(0.22),
     "WRITING RULES", 9, bold=True, color=AVA_GREEN)
rules = [
    ("VERB-FIRST", "Start with action verb (Reduce · Add · Eliminate · Improve · Replace)"),
    ("MEASURABLE", "Include a number when possible (from X to Y, by Z%)"),
    ("OUTCOME, NOT TASK", "Describe what changes for the user, not what code you'll write"),
    ("ONE LINE", "Max 120 characters · if it doesn't fit, scope is too broad"),
    ("NO JARGON", "If a stakeholder outside IT can't read it, rewrite"),
]
ry = body_y + Inches(1.32)
for rule, desc in rules:
    round_rect(slide2, Inches(6.93), ry, Inches(2.0), Inches(0.32), AVA_GREEN, radius=0.05)
    text(slide2, Inches(6.93), ry + Inches(0.04), Inches(2.0), Inches(0.22),
         rule, 9, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(slide2, Inches(9.05), ry + Inches(0.04), Inches(3.85), Inches(0.32),
         desc, 9, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
    ry += Inches(0.40)

text(slide2, Inches(6.93), body_y + Inches(3.55), Inches(6), Inches(0.22),
     "WHAT BREAKS WITHOUT IT", 9, bold=True, color=RED)
text(slide2, Inches(6.93), body_y + Inches(3.77), Inches(5.9), Inches(0.95),
     "Generic titles like \"Fix system\" make every card look the same on the board. Triage cannot tell at a glance which item is critical. Backlog ranking becomes meaningless. Stakeholders ignore the dashboard.",
     10, color=TEXT)

bottom_strip(slide2,
             "Clear, outcome-based summary becomes the card display label across boards. First field execs see when scanning. Wrong title = invisible card.",
             "Card display label · backlog scannability · search indexing")


# ============================================================================
# SLIDE 3 — DESCRIPTION field
# ============================================================================
slide3 = prs.slides.add_slide(prs.slide_layouts[6])
field_slide_template(slide3, "Description",
                     "Full context · what is happening today and what blocks the business",
                     "MANDATORY", "Rich text · min 100 chars · supports lists, links, mentions")

# Left: structured guide
round_rect(slide3, Inches(0.4), body_y, Inches(6.18), body_h, WHITE, GREY_BORDER, Pt(0.5), 0.03)
text(slide3, Inches(0.55), body_y + Inches(0.1), Inches(6), Inches(0.22),
     "STRUCTURE — 4 paragraphs", 9, bold=True, color=AVA_GREEN)

paras = [
    ("CONTEXT", "What is happening today · the as-is state · who does it · how often",
     "Customer Care reps manually reconcile billing errors after each meter reading cycle (every 30 days, ~1,200 errors per cycle across 7 OPCOs)."),
    ("CURRENT PAIN", "What hurts · what risk · what waste",
     "Each error takes 12 min to fix, totaling 240 hrs/cycle of manual labor. Customers wait 14 days for corrected bill — generates 8% of inbound complaints."),
    ("WHO IS AFFECTED", "Users · OPCOs · regulators · field crews",
     "65 Customer Care agents (CR1) · 7 OPCOs (UIL, CMP, NY×3, RG&E×2) · all residential billing customers in error states."),
    ("WHAT BLOCKS TODAY", "Workarounds · manual steps · missing data",
     "No automated reconciliation between MDM and SAP FICA. Excel-based reports run twice daily by hand. Vendor SAP partner queue is 3 weeks deep."),
]
py = body_y + Inches(0.35)
for label, why, example in paras:
    round_rect(slide3, Inches(0.55), py, Inches(5.88), Inches(1.05), AVA_GREEN_LIGHT, AVA_GREEN, Pt(1), 0.05)
    round_rect(slide3, Inches(0.55), py, Inches(1.7), Inches(1.05), AVA_GREEN, radius=0.05)
    text(slide3, Inches(0.55), py, Inches(1.7), Inches(1.05),
         label, 10, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(slide3, Inches(2.35), py + Inches(0.05), Inches(4.0), Inches(0.18),
         why, 8, italic=True, color=AVA_GREEN_DARK)
    text(slide3, Inches(2.35), py + Inches(0.25), Inches(4.0), Inches(0.78),
         example, 9, color=TEXT)
    py += Inches(1.13)

# Right: rules
round_rect(slide3, Inches(6.78), body_y, Inches(6.18), body_h, WHITE, GREY_BORDER, Pt(0.5), 0.03)
text(slide3, Inches(6.93), body_y + Inches(0.1), Inches(6), Inches(0.22),
     "WHAT IT IS", 9, bold=True, color=AVA_GREEN)
text(slide3, Inches(6.93), body_y + Inches(0.32), Inches(5.9), Inches(0.7),
     "Free-text narrative giving the triage team enough business context to understand the problem WITHOUT scheduling a meeting. Description is the difference between a TRIAGED ticket and a BOUNCED one.",
     10, color=TEXT)

text(slide3, Inches(6.93), body_y + Inches(1.10), Inches(6), Inches(0.22),
     "GOLDEN RULE", 9, bold=True, color=AVA_GREEN)
round_rect(slide3, Inches(6.93), body_y + Inches(1.32), Inches(5.9), Inches(0.5),
           AVA_GREEN_LIGHT, AVA_GREEN, Pt(2), 0.05)
text(slide3, Inches(6.93), body_y + Inches(1.32), Inches(5.9), Inches(0.5),
     "Describe the PROBLEM · NOT the solution",
     14, bold=True, color=AVA_GREEN_DARK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

text(slide3, Inches(6.93), body_y + Inches(1.95), Inches(6), Inches(0.22),
     "WHY THIS MATTERS", 9, bold=True, color=AVA_GREEN)
text(slide3, Inches(6.93), body_y + Inches(2.17), Inches(5.9), Inches(0.85),
     "If you write \"add a button to do X\" the team is locked into your solution. If you write \"users can't do X today, costing $Y/month\" the team can engineer the cheapest path. Solutions belong in Acceptance Criteria, not Description.",
     10, color=TEXT)

text(slide3, Inches(6.93), body_y + Inches(3.10), Inches(6), Inches(0.22),
     "GETS BOUNCED IF…", 9, bold=True, color=RED)
bounced = [
    "Less than 100 characters",
    "No mention of who is affected",
    "No quantification (volume · frequency · impact)",
    "Reads as a feature request, not a problem statement",
]
by_loc = body_y + Inches(3.32)
for b in bounced:
    text(slide3, Inches(6.93), by_loc, Inches(5.9), Inches(0.22),
         "✗  " + b, 9, color=RED)
    by_loc += Inches(0.26)

bottom_strip(slide3,
             "Triage cannot validate DoR or compute Pain Score without business context. Description is the difference between TRIAGED and BOUNCED.",
             "Drives Triage decision · DoR validation · Pain Score component")


# ============================================================================
# SLIDE 4 — BUSINESS VALUE field
# ============================================================================
slide4 = prs.slides.add_slide(prs.slide_layouts[6])
field_slide_template(slide4, "Business Value",
                     "Quantified outcome that justifies investment · feeds Pain Score formula",
                     "MANDATORY", "Rich text + numeric BV component (1–5 score)")

# Left: 5-point scoring guide
round_rect(slide4, Inches(0.4), body_y, Inches(6.18), body_h, WHITE, GREY_BORDER, Pt(0.5), 0.03)
text(slide4, Inches(0.55), body_y + Inches(0.1), Inches(6), Inches(0.22),
     "BV COMPONENT — 5-point scoring guide", 9, bold=True, color=AVA_GREEN)
text(slide4, Inches(0.55), body_y + Inches(0.3), Inches(6), Inches(0.22),
     "Used in Pain Score formula · scored at triage with the Champion",
     9, italic=True, color=GREY_LABEL)

bv_levels = [
    ("5", "TRANSFORMATIONAL",
     "Multi-million $ impact · safety · regulatory non-compliance · existential",
     RED, WHITE),
    ("4", "VERY HIGH",
     "Major contribution to OPCO results · customer NPS shift · year-end target",
     ORANGE, WHITE),
    ("3", "HIGH",
     "Moderate gain · cost / time / quality measurable improvement",
     YELLOW, NAVY),
    ("2", "LOW",
     "Limited gain · convenience · efficiency · single team",
     AVA_GREEN_LIGHT, AVA_GREEN_DARK),
    ("1", "MINIMAL",
     "Cosmetic · nice-to-have · no measurable business impact",
     GREY_BG, GREY_LABEL),
]
ly = body_y + Inches(0.6)
for n, lbl, desc, fill, txc in bv_levels:
    round_rect(slide4, Inches(0.55), ly, Inches(5.88), Inches(0.78), fill, AVA_GREEN, Pt(0.5), 0.04)
    round_rect(slide4, Inches(0.55), ly, Inches(0.85), Inches(0.78), AVA_GREEN, radius=0.04)
    text(slide4, Inches(0.55), ly, Inches(0.85), Inches(0.78),
         n, 24, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(slide4, Inches(1.5), ly + Inches(0.06), Inches(4.8), Inches(0.22),
         lbl, 11, bold=True, color=AVA_GREEN_DARK)
    text(slide4, Inches(1.5), ly + Inches(0.30), Inches(4.8), Inches(0.45),
         desc, 9, color=TEXT)
    ly += Inches(0.85)

# Right: how to write
round_rect(slide4, Inches(6.78), body_y, Inches(6.18), body_h, WHITE, GREY_BORDER, Pt(0.5), 0.03)
text(slide4, Inches(6.93), body_y + Inches(0.1), Inches(6), Inches(0.22),
     "TEXT NARRATIVE — quantify the outcome", 9, bold=True, color=AVA_GREEN)

text(slide4, Inches(6.93), body_y + Inches(0.35), Inches(6), Inches(0.22),
     "✓ INCLUDE NUMBERS", 10, bold=True, color=AVA_GREEN)
include = [
    ("$ impact", "$2M/yr cost avoidance · $500k revenue uplift"),
    ("% delta", "NPS +5 points · CSAT 78% → 88%"),
    ("Time saved", "240 hrs/cycle · 5 days off cycle time"),
    ("Volume", "1,200 errors/cycle · 8% of inbound calls"),
    ("Compliance", "PUC mandate · SOX gap · OSHA finding"),
    ("Safety", "Field crew exposure · public hazard · near-miss"),
]
iy = body_y + Inches(0.58)
for k, v in include:
    text(slide4, Inches(6.93), iy, Inches(2.0), Inches(0.22),
         "▸ " + k, 9, bold=True, color=AVA_GREEN_DARK)
    text(slide4, Inches(9.0), iy, Inches(3.85), Inches(0.22),
         v, 9, color=TEXT)
    iy += Inches(0.26)

text(slide4, Inches(6.93), body_y + Inches(2.30), Inches(6), Inches(0.22),
     "✗ AVOID", 10, bold=True, color=RED)
avoid = [
    "Strategic alignment  ·  Industry best practice",
    "Customer satisfaction (without number)",
    "It would be nice to have  ·  Future-proofing",
    "Better experience (without metric)",
]
ay = body_y + Inches(2.55)
for a in avoid:
    text(slide4, Inches(6.93), ay, Inches(5.9), Inches(0.22),
         "✗  " + a, 9, color=RED)
    ay += Inches(0.24)

text(slide4, Inches(6.93), body_y + Inches(3.55), Inches(6), Inches(0.22),
     "EXAMPLE — high BV statement", 9, bold=True, color=AVA_GREEN)
round_rect(slide4, Inches(6.93), body_y + Inches(3.78), Inches(5.9), Inches(0.95),
           AVA_GREEN_LIGHT, AVA_GREEN, Pt(1), 0.05)
text(slide4, Inches(7.05), body_y + Inches(3.85), Inches(5.7), Inches(0.85),
     '"Reducing billing-correction cycle from 14 to 5 days saves 240 hrs/cycle (~$285k/yr labor) and reduces inbound complaints by ~8% (480 calls/cycle). Aligns with 2026 NPS target +5. BV = 4 (VERY HIGH).',
     9, italic=True, color=AVA_GREEN_DARK)

bottom_strip(slide4,
             "Quantified value makes prioritization defensible. Without numbers, Pain Score becomes opinion. Champions can't game ranking.",
             "Pain Score (BV component · weight 0.35 in CS) · Gate 2 review · Portfolio Committee defense")


# ============================================================================
# SLIDE 5 — ACCEPTANCE CRITERIA field
# ============================================================================
slide5 = prs.slides.add_slide(prs.slide_layouts[6])
field_slide_template(slide5, "Acceptance Criteria",
                     "Measurable conditions that define DONE · used at UAT and Gate sign-off",
                     "MANDATORY", "Rich text · checklist format · each item testable")

# Left: structure
round_rect(slide5, Inches(0.4), body_y, Inches(6.18), body_h, WHITE, GREY_BORDER, Pt(0.5), 0.03)
text(slide5, Inches(0.55), body_y + Inches(0.1), Inches(6), Inches(0.22),
     "FORMAT — checklist with measurable items", 9, bold=True, color=AVA_GREEN)

text(slide5, Inches(0.55), body_y + Inches(0.4), Inches(6), Inches(0.22),
     "STRUCTURE PER ITEM", 9, bold=True, color=GREY_LABEL)
struct = [
    ("Action", "What the system / user does", AVA_GREEN),
    ("Threshold", "Number or boolean state", AVA_GREEN),
    ("Window", "Over what time / volume", AVA_GREEN),
    ("Verification", "Who confirms · how", AVA_GREEN),
]
sy = body_y + Inches(0.65)
for k, v, c in struct:
    round_rect(slide5, Inches(0.55), sy, Inches(1.6), Inches(0.32), c, radius=0.05)
    text(slide5, Inches(0.55), sy, Inches(1.6), Inches(0.32),
         k, 10, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(slide5, Inches(2.3), sy + Inches(0.05), Inches(4.0), Inches(0.22),
         v, 9, color=TEXT)
    sy += Inches(0.4)

text(slide5, Inches(0.55), body_y + Inches(2.45), Inches(6), Inches(0.22),
     "EXAMPLE CHECKLIST", 9, bold=True, color=AVA_GREEN)
ac_items = [
    "Billing engine completes daily run before 06:00 EST",
    "Error rate < 0.5% across all 7 OPCOs · 30-day rolling average",
    "P95 cycle time ≤ 5 days · verified over 30-day window",
    "Documentation updated · runbook published in Confluence",
    "UAT signed off by Customer Care lead (named Sponsor)",
]
ay = body_y + Inches(2.7)
for ac in ac_items:
    round_rect(slide5, Inches(0.55), ay, Inches(5.88), Inches(0.34),
               AVA_GREEN_LIGHT, AVA_GREEN, Pt(0.5), 0.04)
    text(slide5, Inches(0.65), ay + Inches(0.05), Inches(0.3), Inches(0.24),
         "✓", 12, bold=True, color=AVA_GREEN, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(slide5, Inches(1.0), ay + Inches(0.05), Inches(5.4), Inches(0.24),
         ac, 9, color=TEXT)
    ay += Inches(0.40)

# Right: rules + bad examples
round_rect(slide5, Inches(6.78), body_y, Inches(6.18), body_h, WHITE, GREY_BORDER, Pt(0.5), 0.03)
text(slide5, Inches(6.93), body_y + Inches(0.1), Inches(6), Inches(0.22),
     "WHAT IT IS", 9, bold=True, color=AVA_GREEN)
text(slide5, Inches(6.93), body_y + Inches(0.32), Inches(5.9), Inches(0.92),
     "The DEFINITION OF DONE for this demand. Each item is a binary check — either it passed UAT or it didn't. Acceptance Criteria is what UAT tests against, what Gate 1 reviews, and what Sponsor signs off.",
     10, color=TEXT)

text(slide5, Inches(6.93), body_y + Inches(1.30), Inches(6), Inches(0.22),
     "RULES", 9, bold=True, color=AVA_GREEN)
rules = [
    ("BINARY", "Each item must be testable as PASS or FAIL — no shades of gray"),
    ("MEASURABLE", "Includes a number or a defined state · no \"better\" or \"faster\""),
    ("VERIFIABLE", "Names who confirms (Sponsor / UAT team / vendor) and how"),
    ("INDEPENDENT", "Each item testable on its own · no chained dependencies"),
    ("3 TO 7 ITEMS", "Less than 3 = scope vague · more than 7 = scope too wide"),
]
ry = body_y + Inches(1.52)
for k, d in rules:
    round_rect(slide5, Inches(6.93), ry, Inches(1.5), Inches(0.32), AVA_GREEN, radius=0.05)
    text(slide5, Inches(6.93), ry, Inches(1.5), Inches(0.32),
         k, 9, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(slide5, Inches(8.55), ry + Inches(0.04), Inches(4.3), Inches(0.32),
         d, 9, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
    ry += Inches(0.40)

text(slide5, Inches(6.93), body_y + Inches(3.62), Inches(6), Inches(0.22),
     "BAD EXAMPLES (✗ — not measurable)", 9, bold=True, color=RED)
bad_ac = [
    "Users are happy with the new flow",
    "Performance is improved",
    "It works as expected",
    "Solution meets business needs",
]
bay = body_y + Inches(3.85)
for b in bad_ac:
    text(slide5, Inches(6.93), bay, Inches(5.9), Inches(0.22),
         "✗  " + b, 9, color=RED)
    bay += Inches(0.22)

bottom_strip(slide5,
             "Without measurable AC, the squad cannot tell if work is done. UAT becomes opinion. Gates can't sign off. Demand stays in REVIEW forever.",
             "Definition of Done · UAT signoff · Gate 1 reference (if Build/Transform)")


# ============================================================================
# SLIDE 6 — VALUE STREAM destination (the routing tree)
# ============================================================================
slide6 = prs.slides.add_slide(prs.slide_layouts[6])
field_slide_template(slide6, "Value Stream destination",
                     "Single most important field on the form · drives auto-routing",
                     "MANDATORY", "Single-select dropdown · 10 options (5 ITBA + 4 CS DVS + 1 Enabling Team)",
                     accent=AVA_ORANGE)

# Big tree diagram
diag_y = Inches(1.95)
diag_h = Inches(4.95)

# Closed dropdown at top
round_rect(slide6, Inches(0.4), diag_y, Inches(12.55), Inches(0.4),
           AVA_ORANGE_LIGHT, AVA_ORANGE, Pt(2), 0.05)
text(slide6, Inches(0.55), diag_y + Inches(0.06), Inches(11), Inches(0.28),
     "▼  Meter & Device Operations  ·  CS · Kumar Jaladanki",
     12, bold=True, color=TEXT)

# Two panels open
panel_y = diag_y + Inches(0.55)
panel_h = Inches(4.3)
panel_w = (Inches(12.55) - Inches(0.3)) / 2

# ITBA panel
round_rect(slide6, Inches(0.4), panel_y, panel_w, panel_h, GREY_BG, ITBA_GREY, Pt(2), 0.04)
round_rect(slide6, Inches(0.4), panel_y, panel_w, Inches(0.45), ITBA_GREY, radius=0.04)
text(slide6, Inches(0.55), panel_y + Inches(0.08), panel_w - Inches(0.3), Inches(0.3),
     "ITBA · IT Business Apps · Carlos Garbin · 5 teams (unchanged)",
     12, bold=True, color=WHITE)
itba_options = [
    ("Field Operations", "GIS · Workforce · OSG", "Distribution / field crew systems"),
    ("Engineering & Asset Mgmt", "Maintenance · Engineering", "Asset lifecycle · capital planning"),
    ("Network Operations & Control", "SCADA · EMS · OMS", "Real-time grid operations"),
    ("Corporate & Regulatory", "FI · HR · Procurement · Reg.", "Corporate functions · compliance"),
    ("Iberdrola Corp Services", "Global SAP · Corp Apps", "Iberdrola Spain corporate"),
]
oy = panel_y + Inches(0.6)
for opt, sub, ctx in itba_options:
    round_rect(slide6, Inches(0.55), oy, panel_w - Inches(0.3), Inches(0.65),
               WHITE, GREY_BORDER, Pt(0.5), 0.04)
    text(slide6, Inches(0.7), oy + Inches(0.06), Inches(0.3), Inches(0.22),
         "○", 14, color=GREY_LABEL)
    text(slide6, Inches(0.95), oy + Inches(0.04), panel_w - Inches(0.7), Inches(0.22),
         opt, 11, bold=True, color=TEXT)
    text(slide6, Inches(0.95), oy + Inches(0.24), panel_w - Inches(0.7), Inches(0.18),
         sub, 8, italic=True, color=GREY_LABEL)
    text(slide6, Inches(0.95), oy + Inches(0.42), panel_w - Inches(0.7), Inches(0.18),
         ctx, 8, color=GREY_TEXT)
    oy += Inches(0.72)
text(slide6, Inches(0.55), panel_y + panel_h - Inches(0.32), panel_w - Inches(0.3), Inches(0.24),
     "→ ITBA Portfolio Board · Carlos owns",
     10, italic=True, bold=True, color=GREY_LABEL,
     align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

# CS panel
cs_x = Inches(0.4) + panel_w + Inches(0.3)
round_rect(slide6, cs_x, panel_y, panel_w, panel_h, JIRA_BLUE_LIGHT, JIRA_BLUE, Pt(2), 0.04)
round_rect(slide6, cs_x, panel_y, panel_w, Inches(0.45), JIRA_BLUE, radius=0.04)
text(slide6, cs_x + Inches(0.15), panel_y + Inches(0.08), panel_w - Inches(0.3), Inches(0.3),
     "CS · Customer Services Apps · Kumar Jaladanki · 4 DVS + Enabling",
     12, bold=True, color=WHITE)
cs_options = [
    ("Meter & Device Operations", "Smart Metering · MDM · IS-U Device · 8760 EDI", "Capture every read · zero data loss", JIRA_BLUE, True),
    ("Billing Operations", "SAP ECC (FI/SD) · Streamserve · OpenText", "Bill on time · every OPCO · every day", JIRA_BLUE, False),
    ("Revenue & Collections", "SAP FICA · SOVOS Taxware · Payment", "Convert invoices into cash · comply", JIRA_BLUE, False),
    ("Customer Care & Engagement", "SAP CRM · Genesys · Nuance IVR · Medallia", "Resolve interactions on first contact", JIRA_BLUE, False),
    ("Enabling Team — Foundation Services", "Basis · Gold Client · Blue Prism · TOSCA · Jira (no PM)", "Shared services for the 4 DVS", PURPLE, False),
]
oy = panel_y + Inches(0.6)
for opt, sub, ctx, c, sel in cs_options:
    fill = AVA_GREEN_LIGHT if sel else WHITE
    line = AVA_GREEN if sel else GREY_BORDER
    round_rect(slide6, cs_x + Inches(0.15), oy, panel_w - Inches(0.3), Inches(0.65),
               fill, line, Pt(2 if sel else 0.5), 0.04)
    text(slide6, cs_x + Inches(0.3), oy + Inches(0.06), Inches(0.3), Inches(0.22),
         "●" if sel else "○", 14, bold=sel, color=AVA_GREEN if sel else GREY_LABEL)
    text(slide6, cs_x + Inches(0.55), oy + Inches(0.04), panel_w - Inches(0.7), Inches(0.22),
         opt, 11, bold=True, color=c)
    text(slide6, cs_x + Inches(0.55), oy + Inches(0.24), panel_w - Inches(0.7), Inches(0.18),
         sub, 8, italic=True, color=GREY_LABEL)
    text(slide6, cs_x + Inches(0.55), oy + Inches(0.42), panel_w - Inches(0.7), Inches(0.18),
         ctx, 8, color=GREY_TEXT)
    oy += Inches(0.72)
text(slide6, cs_x + Inches(0.15), panel_y + panel_h - Inches(0.32), panel_w - Inches(0.3), Inches(0.24),
     "→ CS Portfolio Board · Kumar owns",
     10, italic=True, bold=True, color=JIRA_BLUE,
     align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)

bottom_strip(slide6,
             "Auto-routing only works if Value Stream is set. Without it, ticket sits in NEW status forever. Champion never gets notified. Triage never happens.",
             "Auto-Routing (A1) · @notify Champion · sets squad swimlane · routes to portfolio board",
             accent=AVA_ORANGE)


# ============================================================================
# SLIDE 7 — CLASS OF SERVICE (4 options · detailed)
# ============================================================================
slide7 = prs.slides.add_slide(prs.slide_layouts[6])
field_slide_template(slide7, "Class of Service",
                     "Defines squad swimlane · WIP limits · aging SLA · 4 options",
                     "MANDATORY", "Single-select · 4 options · best-guess at intake · finalized at triage")

# 4 cards in a 2x2 grid
cw = (Inches(12.55) - Inches(0.3)) / 2
ch = (Inches(4.85) - Inches(0.2)) / 2

cos = [
    ("Standard", "Default flow · FIFO · Pain Score ranked",
     "Bulk of demand · normal scheduling · no external deadline · no customer-facing severity",
     "60–70% of squad capacity",
     "WIP 3–4 per squad · 14 days aging alert",
     "Disconnect flow rebuild · CRM data dedup · ARCFM enhancement",
     JIRA_BLUE, WHITE),
    ("Fixed-Date", "Deadline-locked · scheduled backwards from end-date",
     "External deadline (regulator · contract · vendor lock) · late delivery has compliance / financial penalty",
     "15–25% of capacity (highest in VS-3 Corporate)",
     "WIP 2 per squad · 50% remaining timeline alert",
     "PUC tariff banner · SOX audit log · Iberdrola Q-close · CT/ME/NY rollouts",
     YELLOW, NAVY),
    ("Expedite", "Jump-the-queue · 1 in flight · capacity reserved",
     "P1/P2 customer impact · production outage · regulatory deadline imminent (< 30 days)",
     "5–10% capacity reserve · blocked-out per quarter",
     "WIP 1 per squad · 4-hour alert · auto-escalate",
     "Billing engine down · IVR 503 · Iberdrola month-close blocker",
     PINK, WHITE),
    ("Intangible", "Tech debt · refactor · enablers · platform work",
     "No direct user-visible value · long-term health · refactor · platform · documentation",
     "10–15% capacity reserved per quarter",
     "WIP 1 per squad · 90 days aging · low-priority alert",
     "CRM tech debt · SAP Basis kernel patches · pipeline modernization",
     AVA_GREEN, WHITE),
]
positions = [(Inches(0.4), Inches(1.85)), (Inches(0.4) + cw + Inches(0.3), Inches(1.85)),
             (Inches(0.4), Inches(1.85) + ch + Inches(0.2)), (Inches(0.4) + cw + Inches(0.3), Inches(1.85) + ch + Inches(0.2))]

for i, (lbl, desc, when, capacity, wip, examples, fill, txc) in enumerate(cos):
    px, py = positions[i]
    round_rect(slide7, px, py, cw, ch, WHITE, fill, Pt(2.5), 0.04)
    round_rect(slide7, px, py, cw, Inches(0.5), fill, radius=0.04)
    text(slide7, px + Inches(0.15), py + Inches(0.06), cw, Inches(0.22),
         lbl, 14, bold=True, color=txc)
    text(slide7, px + Inches(0.15), py + Inches(0.27), cw, Inches(0.20),
         desc, 8, italic=True, color=txc)

    # Body content
    body_x = px + Inches(0.2)
    body_y_loc = py + Inches(0.6)
    text(slide7, body_x, body_y_loc, cw - Inches(0.3), Inches(0.18),
         "WHEN TO USE", 8, bold=True, color=GREY_LABEL)
    text(slide7, body_x, body_y_loc + Inches(0.16), cw - Inches(0.3), Inches(0.4),
         when, 9, color=TEXT)

    text(slide7, body_x, body_y_loc + Inches(0.6), cw - Inches(0.3), Inches(0.18),
         "CAPACITY ALLOCATION", 8, bold=True, color=GREY_LABEL)
    text(slide7, body_x, body_y_loc + Inches(0.78), cw - Inches(0.3), Inches(0.22),
         capacity, 10, bold=True, color=fill)

    text(slide7, body_x, body_y_loc + Inches(1.05), cw - Inches(0.3), Inches(0.18),
         "WIP / AGING", 8, bold=True, color=GREY_LABEL)
    text(slide7, body_x, body_y_loc + Inches(1.21), cw - Inches(0.3), Inches(0.22),
         wip, 9, color=TEXT)

    text(slide7, body_x, body_y_loc + Inches(1.50), cw - Inches(0.3), Inches(0.18),
         "EXAMPLES", 8, bold=True, color=GREY_LABEL)
    text(slide7, body_x, body_y_loc + Inches(1.66), cw - Inches(0.3), Inches(0.4),
         examples, 8, italic=True, color=TEXT)


# ============================================================================
# SLIDE 8 — TYPE + SIZE
# ============================================================================
slide8 = prs.slides.add_slide(prs.slide_layouts[6])
field_slide_template(slide8, "Type + Size",
                     "Two operational fields that branch the path through gates",
                     "MANDATORY", "Type: Demand · Incident · Change  ·  Size: S · M · L · XL")

# Type (left)
round_rect(slide8, Inches(0.4), body_y, Inches(6.18), body_h, WHITE, AVA_GREEN, Pt(2), 0.03)
round_rect(slide8, Inches(0.4), body_y, Inches(6.18), Inches(0.4), AVA_GREEN, radius=0.03)
text(slide8, Inches(0.55), body_y + Inches(0.08), Inches(6), Inches(0.24),
     "TYPE  ·  Demand · Incident · Change", 13, bold=True, color=WHITE)

types = [
    ("Demand", "New work · features · projects · enhancements",
     "Most submissions. Goes through full path: NEW → ROUTED → TRIAGED → BACKLOG → Gates → Squad.",
     "Standard intake flow", AVA_GREEN, WHITE),
    ("Incident", "Production outage · break-fix · service degradation",
     "Bypasses Gates. Auto-converts to Bug. Routed directly to Squad Expedite swimlane. CoS = Expedite.",
     "Skip BACKLOG · skip Gates · go to Squad Expedite", RED, WHITE),
    ("Change", "Config change · vendor refresh · regulated change",
     "Goes through Triage but with lighter DoR. Still requires Gate 2 if Size ≥ M. Common for Upgrades.",
     "Light DoR · Gate 2 if M+", JIRA_BLUE, WHITE),
]
ty = body_y + Inches(0.55)
for tag, desc, behavior, path, fill, txc in types:
    round_rect(slide8, Inches(0.55), ty, Inches(5.88), Inches(1.35), WHITE, fill, Pt(1.5), 0.04)
    round_rect(slide8, Inches(0.55), ty, Inches(1.4), Inches(1.35), fill, radius=0.04)
    text(slide8, Inches(0.55), ty, Inches(1.4), Inches(0.7),
         tag, 14, bold=True, color=txc, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(slide8, Inches(0.55), ty + Inches(0.7), Inches(1.4), Inches(0.65),
         desc, 8, italic=True, color=txc, align=PP_ALIGN.CENTER)
    # Behavior
    text(slide8, Inches(2.05), ty + Inches(0.1), Inches(4.4), Inches(0.18),
         "BEHAVIOR", 8, bold=True, color=GREY_LABEL)
    text(slide8, Inches(2.05), ty + Inches(0.27), Inches(4.4), Inches(0.55),
         behavior, 9, color=TEXT)
    text(slide8, Inches(2.05), ty + Inches(0.92), Inches(4.4), Inches(0.18),
         "PATH", 8, bold=True, color=GREY_LABEL)
    text(slide8, Inches(2.05), ty + Inches(1.09), Inches(4.4), Inches(0.22),
         path, 9, italic=True, bold=True, color=fill)
    ty += Inches(1.4)

# Size (right)
round_rect(slide8, Inches(6.78), body_y, Inches(6.18), body_h, WHITE, AVA_ORANGE, Pt(2), 0.03)
round_rect(slide8, Inches(6.78), body_y, Inches(6.18), Inches(0.4), AVA_ORANGE, radius=0.03)
text(slide8, Inches(6.93), body_y + Inches(0.08), Inches(6), Inches(0.24),
     "SIZE  ·  drives auto-conversion to Story / Epic / Initiative",
     13, bold=True, color=WHITE)

sizes = [
    ("S", "Small", "< 5 days · < $10k · 1 squad",
     "Auto-converts to STORY. Skip Gate 1. Skip Gate 2 if < 2 days effort. Goes DIRECTLY to Squad Backlog.",
     AVA_GREEN, WHITE),
    ("M", "Medium", "5–20 days · < $500k · 1 squad",
     "Auto-converts to EPIC with placeholder Stories. Skip Gate 1. Gate 2 required. Goes to Portfolio Backlog.",
     JIRA_BLUE, WHITE),
    ("L", "Large", "20–60 days · 1–2 squads",
     "Auto-converts to EPIC (multi-Story). Gate 1 (Shark Tank) required + Gate 2. Portfolio Backlog.",
     ORANGE, WHITE),
    ("XL", "Extra-Large", "> 60 days · cross-team",
     "Auto-converts to INITIATIVE (multi-Epic). Gate 1 + Gate 2 + BRP slot. Enterprise Dashboard.",
     RED, WHITE),
]
sy = body_y + Inches(0.55)
for size, lbl, dim, desc, fill, txc in sizes:
    round_rect(slide8, Inches(6.93), sy, Inches(5.88), Inches(1.0), WHITE, fill, Pt(1.5), 0.04)
    round_rect(slide8, Inches(6.93), sy, Inches(0.85), Inches(1.0), fill, radius=0.04)
    text(slide8, Inches(6.93), sy, Inches(0.85), Inches(1.0),
         size, 22, bold=True, color=txc, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(slide8, Inches(7.92), sy + Inches(0.05), Inches(4.85), Inches(0.22),
         lbl + "  ·  " + dim, 11, bold=True, color=fill)
    text(slide8, Inches(7.92), sy + Inches(0.28), Inches(4.85), Inches(0.7),
         desc, 9, color=TEXT)
    sy += Inches(1.05)


# ============================================================================
# SLIDE 9 — WORK TYPE vs RUN/GROW/TRANSFORM
# ============================================================================
slide9 = prs.slides.add_slide(prs.slide_layouts[6])
field_slide_template(slide9, "Work Type vs Run/Grow/Transform",
                     "Two complementary lenses · BOTH mandatory · same demand has both tags",
                     "MANDATORY", "Two single-select fields · operational lens + strategic lens")

# Two cards side by side
left = (Inches(0.4), AVA_GREEN, "Work Type", "How will this work be executed?",
        [("BAU", "Business as Usual", "Recurring · keep-the-lights-on · bug fixes · incidents · no capability change", AVA_GREEN),
         ("Build", "New development", "Net-new features · new code · integrations · capability extension · roadmap-driven", JIRA_BLUE),
         ("Upgrade", "Version refresh", "SAP kernel patches · vendor refresh · regulatory mandates · forced by external", ORANGE)])

right = (Inches(6.78), PURPLE, "Run / Grow / Transform", "Why does this work matter?",
         [("Run", "Maintain CURRENT capability", "Bug fixes · upgrades · compliance · keep-the-lights-on · default mix ~50%", NAVY),
          ("Grow", "EXTEND existing capability", "New OPCO on FICA · improve UX · reduce friction · extend reach · default ~35%", JIRA_BLUE),
          ("Transform", "NEW · disruptive capability", "GenAI chatbot · self-service portal · net-new business model · default ~15%", PURPLE)])

for col_x, accent, name, question, options in [left, right]:
    round_rect(slide9, col_x, body_y, Inches(6.18), body_h, WHITE, accent, Pt(2.5), 0.03)
    round_rect(slide9, col_x, body_y, Inches(6.18), Inches(0.5), accent, radius=0.03)
    text(slide9, col_x + Inches(0.15), body_y + Inches(0.06), Inches(6), Inches(0.22),
         name, 13, bold=True, color=WHITE)
    text(slide9, col_x + Inches(0.15), body_y + Inches(0.27), Inches(6), Inches(0.20),
         "ANSWERS:  " + question, 9, italic=True, color=WHITE)
    oy = body_y + Inches(0.65)
    for tag, lbl, desc, c in options:
        round_rect(slide9, col_x + Inches(0.15), oy, Inches(5.88), Inches(1.25),
                   WHITE, c, Pt(1.5), 0.04)
        round_rect(slide9, col_x + Inches(0.15), oy, Inches(1.5), Inches(1.25), c, radius=0.04)
        text(slide9, col_x + Inches(0.15), oy + Inches(0.1), Inches(1.5), Inches(0.4),
             tag, 16, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(slide9, col_x + Inches(0.15), oy + Inches(0.65), Inches(1.5), Inches(0.4),
             lbl, 9, italic=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(slide9, col_x + Inches(1.78), oy + Inches(0.18), Inches(4.2), Inches(0.95),
             desc, 9, color=TEXT)
        oy += Inches(1.30)

# Bottom — combined examples
ex_y = body_y + body_h + Inches(0.05)
text(slide9, Inches(0.4), ex_y, Inches(13), Inches(0.18),
     "EVERY DEMAND HAS BOTH TAGS — combined examples", 9, bold=True, color=AVA_GREEN_DARK)


# ============================================================================
# SLIDE 10 — PAIN SCORE · ORIGIN & CONCEPT
# ============================================================================
slide10 = prs.slides.add_slide(prs.slide_layouts[6])
field_slide_template(slide10, "Pain Score · Origin",
                     "Where the concept came from and why it makes sense for Avangrid",
                     "AUTO", "Numeric · auto-computed at TRIAGED status · read-only",
                     accent=AVA_GREEN)

# Left: P&S original (6-factor)
round_rect(slide10, Inches(0.4), body_y, Inches(6.18), body_h, WHITE, AVA_GREEN, Pt(2), 0.03)
text(slide10, Inches(0.55), body_y + Inches(0.1), Inches(6), Inches(0.22),
     "ORIGIN — P&S Networks (utility infra)", 11, bold=True, color=AVA_GREEN_DARK)
text(slide10, Inches(0.55), body_y + Inches(0.32), Inches(6), Inches(0.22),
     "Source · Pain Point Evidence Matrix · Slide 41 (Discovery Playbook)",
     8, italic=True, color=GREY_LABEL)

# Original formula box
round_rect(slide10, Inches(0.55), body_y + Inches(0.6), Inches(5.88), Inches(0.85),
           NAVY, radius=0.05)
text(slide10, Inches(0.55), body_y + Inches(0.6), Inches(5.88), Inches(0.85),
     "(Freq×0.20) + (TimeLost×0.15) + (SLA×0.20) + (Safety×0.20) + (Cost×0.15) + (Frustration×0.10)",
     10, bold=True, color=YELLOW, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

text(slide10, Inches(0.55), body_y + Inches(1.55), Inches(6), Inches(0.22),
     "6 FACTORS · UTILITY-CALIBRATED", 10, bold=True, color=AVA_GREEN)
factors = [
    ("Frequency", "How often the pain happens", "0.20", AVA_GREEN),
    ("Time Lost", "Productivity impact (hrs/occurrence)", "0.15", AVA_GREEN),
    ("SLA", "Compliance risk · breach exposure", "0.20", AVA_GREEN),
    ("Safety", "Field crew · public · regulatory", "0.20", RED),
    ("Cost", "Direct $ impact / occurrence", "0.15", AVA_GREEN),
    ("Frustration", "User experience pain", "0.10", AVA_GREEN),
]
fy = body_y + Inches(1.78)
for f, d, w, c in factors:
    round_rect(slide10, Inches(0.55), fy, Inches(5.88), Inches(0.32),
               WHITE, c, Pt(0.5), 0.04)
    text(slide10, Inches(0.65), fy + Inches(0.06), Inches(1.5), Inches(0.22),
         f, 9, bold=True, color=c)
    text(slide10, Inches(2.2), fy + Inches(0.06), Inches(3.0), Inches(0.22),
         d, 8, color=TEXT)
    text(slide10, Inches(5.5), fy + Inches(0.06), Inches(0.85), Inches(0.22),
         w, 10, bold=True, color=c, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    fy += Inches(0.36)

text(slide10, Inches(0.55), body_y + Inches(4.05), Inches(6), Inches(0.22),
     "VERBATIM (P&S Slide 41)",
     8, bold=True, color=GREY_LABEL)
text(slide10, Inches(0.55), body_y + Inches(4.22), Inches(5.88), Inches(0.6),
     '"Weights reflect utilities realities: Safety & SLA are heavily weighted, then Frequency, Time Lost, and Cost follow."',
     9, italic=True, color=TEXT)

# Right: CS adaptation (4-factor)
round_rect(slide10, Inches(6.78), body_y, Inches(6.18), body_h, WHITE, JIRA_BLUE, Pt(2), 0.03)
text(slide10, Inches(6.93), body_y + Inches(0.1), Inches(6), Inches(0.22),
     "CS ADAPTATION — Customer Services", 11, bold=True, color=JIRA_BLUE)
text(slide10, Inches(6.93), body_y + Inches(0.32), Inches(6), Inches(0.22),
     "Source · CS Proposed Value Streams Slide 3 · 4-factor · UX-heavy",
     8, italic=True, color=GREY_LABEL)

# CS formula box
round_rect(slide10, Inches(6.93), body_y + Inches(0.6), Inches(5.88), Inches(0.85),
           AVA_GREEN_DARK, radius=0.05)
text(slide10, Inches(6.93), body_y + Inches(0.6), Inches(5.88), Inches(0.85),
     "H(x)  =  ( BV·0.35 + UX·0.40 + Safety·0.25 )  /  Effort",
     12, bold=True, color=YELLOW,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

text(slide10, Inches(6.93), body_y + Inches(1.55), Inches(6), Inches(0.22),
     "WHY 4 FACTORS · WHY UX-HEAVY", 10, bold=True, color=JIRA_BLUE)

mappings = [
    ("BV (35%)", "absorbs Frequency + Time Lost + Cost", "Quantified business impact"),
    ("UX (40%)", "absorbs Frustration + Customer Experience", "CS is customer-facing — UX wins"),
    ("Safety (25%)", "kept (utility realities)", "Field crew · public · regulatory"),
    ("Effort", "moves to denominator", "Forces ROI thinking · high-effort low-value sinks"),
]
my = body_y + Inches(1.78)
for m, d, why in mappings:
    round_rect(slide10, Inches(6.93), my, Inches(5.88), Inches(0.55),
               JIRA_BLUE_LIGHT, JIRA_BLUE, Pt(0.5), 0.04)
    text(slide10, Inches(7.05), my + Inches(0.05), Inches(1.5), Inches(0.22),
         m, 10, bold=True, color=JIRA_BLUE)
    text(slide10, Inches(7.05), my + Inches(0.27), Inches(2.8), Inches(0.22),
         d, 8, italic=True, color=TEXT)
    text(slide10, Inches(9.95), my + Inches(0.05), Inches(2.85), Inches(0.45),
         why, 8, color=GREY_TEXT)
    my += Inches(0.62)

text(slide10, Inches(6.93), body_y + Inches(4.30), Inches(6), Inches(0.22),
     "WHY IT MAKES SENSE FOR AVANGRID", 8, bold=True, color=AVA_GREEN)
why_items = [
    "✓  Utility-calibrated · Safety isolated · not buried in BV",
    "✓  UX-heavy reflects CS strategy (customer-first)",
    "✓  Comparable cross-VS · 4 DVS share same formula",
    "✓  Auto-computed removes politics · Champion can't game ranking",
]
wy = body_y + Inches(4.50)
for w in why_items:
    text(slide10, Inches(6.93), wy, Inches(5.9), Inches(0.20),
         w, 8, color=TEXT)
    wy += Inches(0.22)


# ============================================================================
# SLIDE 11 — PAIN SCORE · FORMULA & AUTO-COMPUTATION
# ============================================================================
slide11 = prs.slides.add_slide(prs.slide_layouts[6])
field_slide_template(slide11, "Pain Score · Auto-Computation",
                     "How the formula is computed · when · what fields feed it",
                     "AUTO", "Numeric · auto-computed at TRIAGED · re-computed if BV/UX/Safety/Effort change",
                     accent=AVA_GREEN)

# Big formula at top
round_rect(slide11, Inches(0.4), Inches(1.85), Inches(12.55), Inches(0.85),
           NAVY, radius=0.04)
text(slide11, Inches(0.4), Inches(1.85), Inches(12.55), Inches(0.85),
     "H(x)  =  ( BV·w₁  +  UX·w₂  +  Safety·w₃ )  /  Effort",
     28, bold=True, color=YELLOW,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 4 components below
comp_y = Inches(2.9)
comp_h = Inches(2.4)
cw_l = Inches(2.95)

components = [
    ("BV", "0.35", "Business Value (1–5)", "from BUSINESS VALUE field",
     "Quantified business impact:\n· $ revenue / cost\n· NPS / CSAT delta\n· Volume affected\n· Compliance gap", AVA_GREEN),
    ("UX", "0.40", "User Experience (1–5)", "scored at TRIAGE with Champion",
     "User-perceived pain:\n· Frustration level\n· Workarounds needed\n· Time-to-task\n· Error frequency", JIRA_BLUE),
    ("Safety", "0.25", "Utility-sector risk (1–5)", "scored at TRIAGE",
     "Safety dimensions:\n· Field crew exposure\n· Public hazard\n· Regulatory breach\n· Cybersecurity", RED),
    ("Effort", "÷", "Complexity (1–5)", "sized by SMEs at TRIAGE",
     "Sizing dimensions:\n· Effort hours\n· Apps touched\n· Vendor dependency\n· Risk of changes", PURPLE),
]
for i, (sym, w, lbl, src, body, c) in enumerate(components):
    cx = Inches(0.4) + i * (cw_l + Inches(0.18))
    round_rect(slide11, cx, comp_y, cw_l, comp_h, WHITE, c, Pt(2.5), 0.04)
    round_rect(slide11, cx, comp_y, cw_l, Inches(0.65), c, radius=0.04)
    text(slide11, cx, comp_y + Inches(0.04), cw_l, Inches(0.32),
         sym, 22, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(slide11, cx, comp_y + Inches(0.40), cw_l, Inches(0.22),
         f"weight {w}", 9, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER)
    text(slide11, cx + Inches(0.15), comp_y + Inches(0.78), cw_l - Inches(0.3), Inches(0.22),
         lbl, 10, bold=True, color=c)
    text(slide11, cx + Inches(0.15), comp_y + Inches(0.98), cw_l - Inches(0.3), Inches(0.22),
         src, 8, italic=True, color=GREY_LABEL)
    text(slide11, cx + Inches(0.15), comp_y + Inches(1.28), cw_l - Inches(0.3), Inches(1.05),
         body, 8, color=TEXT)

# Bottom: when computed
ban_y = Inches(5.4)
text(slide11, Inches(0.4), ban_y, Inches(13), Inches(0.22),
     "WHEN AUTO-COMPUTATION FIRES", 10, bold=True, color=AVA_GREEN)
events = [
    ("Initial",
     "Status changes  NEW → TRIAGED",
     "Pain Score computed first time after Champion + PM agree on BV/UX/Safety/Effort scores",
     AVA_GREEN),
    ("Re-compute",
     "Any BV / UX / Safety / Effort score edited",
     "Triggered automatically · backlog auto-resorted · Slack notification to Squad Lead if rank moves > 3",
     JIRA_BLUE),
    ("Override (rare)",
     "Manual override by Director CS Apps",
     "Audited · requires written justification · captured for monthly review at Portfolio Committee",
     ORANGE),
]
ey = ban_y + Inches(0.28)
for tag, when, desc, c in events:
    round_rect(slide11, Inches(0.4), ey, Inches(12.55), Inches(0.45),
               WHITE, c, Pt(1), 0.04)
    round_rect(slide11, Inches(0.4), ey, Inches(1.5), Inches(0.45), c, radius=0.04)
    text(slide11, Inches(0.4), ey, Inches(1.5), Inches(0.45),
         tag, 10, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(slide11, Inches(2.05), ey + Inches(0.04), Inches(3.5), Inches(0.4),
         when, 9, bold=True, color=c, anchor=MSO_ANCHOR.MIDDLE)
    text(slide11, Inches(5.65), ey + Inches(0.04), Inches(7.2), Inches(0.4),
         desc, 9, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
    ey += Inches(0.50)


# ============================================================================
# SLIDE 12 — POST-SUBMISSION FLOW (visual · boards · status transitions)
# ============================================================================
slide12 = prs.slides.add_slide(prs.slide_layouts[6])
header_band(slide12, "AFTER SUBMISSION  ·  WHAT HAPPENS",
            "Auto-routing → Triage → Auto-conversion to the right Jira hierarchy type")

# Step 1: SUBMIT
s1_y = Inches(1.2)
round_rect(slide12, Inches(5.2), s1_y, Inches(2.95), Inches(0.85),
           AVA_GREEN, radius=0.06)
text(slide12, Inches(5.2), s1_y + Inches(0.04), Inches(2.95), Inches(0.18),
     "STEP 1", 8, bold=True, color=AVA_GREEN_LIGHT, align=PP_ALIGN.CENTER)
text(slide12, Inches(5.2), s1_y + Inches(0.20), Inches(2.95), Inches(0.28),
     "REQUESTER SUBMITS FORM", 12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
text(slide12, Inches(5.2), s1_y + Inches(0.45), Inches(2.95), Inches(0.20),
     "creates  DEMAND-NNNN", 9, italic=True, color=AVA_GREEN_LIGHT, align=PP_ALIGN.CENTER)
round_rect(slide12, Inches(6.4), s1_y + Inches(0.62), Inches(0.55), Inches(0.18), WHITE, radius=0.4)
text(slide12, Inches(6.4), s1_y + Inches(0.62), Inches(0.55), Inches(0.18),
     "NEW", 8, bold=True, color=AVA_GREEN_DARK,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Arrow
arrow1 = slide12.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(6.5), Inches(2.13), Inches(0.35), Inches(0.25))
arrow1.fill.solid(); arrow1.fill.fore_color.rgb = AVA_ORANGE
arrow1.line.fill.background()

# Step 2: AUTO-ROUTING
s2_y = Inches(2.45)
round_rect(slide12, Inches(5.2), s2_y, Inches(2.95), Inches(0.85),
           JIRA_BLUE, radius=0.06)
text(slide12, Inches(5.2), s2_y + Inches(0.04), Inches(2.95), Inches(0.18),
     "STEP 2", 8, bold=True, color=JIRA_BLUE_LIGHT, align=PP_ALIGN.CENTER)
text(slide12, Inches(5.2), s2_y + Inches(0.20), Inches(2.95), Inches(0.28),
     "AUTO-ROUTING (A1)", 12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
text(slide12, Inches(5.2), s2_y + Inches(0.45), Inches(2.95), Inches(0.20),
     "reads Value Stream → assigns board",
     9, italic=True, color=JIRA_BLUE_LIGHT, align=PP_ALIGN.CENTER)
round_rect(slide12, Inches(6.4), s2_y + Inches(0.62), Inches(0.55), Inches(0.18), WHITE, radius=0.4)
text(slide12, Inches(6.4), s2_y + Inches(0.62), Inches(0.55), Inches(0.18),
     "ROUTED", 7, bold=True, color=JIRA_BLUE,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

arrow2 = slide12.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Inches(6.5), Inches(3.38), Inches(0.35), Inches(0.25))
arrow2.fill.solid(); arrow2.fill.fore_color.rgb = AVA_ORANGE
arrow2.line.fill.background()

# Step 3: TRIAGE
s3_y = Inches(3.7)
round_rect(slide12, Inches(4.2), s3_y, Inches(4.95), Inches(0.95),
           PURPLE, radius=0.06)
text(slide12, Inches(4.2), s3_y + Inches(0.04), Inches(4.95), Inches(0.18),
     "STEP 3", 8, bold=True, color=PURPLE_LIGHT, align=PP_ALIGN.CENTER)
text(slide12, Inches(4.2), s3_y + Inches(0.20), Inches(4.95), Inches(0.28),
     "TRIAGE  ·  PM steward + Champion", 12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
text(slide12, Inches(4.2), s3_y + Inches(0.46), Inches(4.95), Inches(0.20),
     "validate DoR · finalize CoS · compute Pain Score · A5 reads Size + Type",
     8, italic=True, color=PURPLE_LIGHT, align=PP_ALIGN.CENTER)
round_rect(slide12, Inches(6.0), s3_y + Inches(0.69), Inches(1.35), Inches(0.20), WHITE, radius=0.4)
text(slide12, Inches(6.0), s3_y + Inches(0.69), Inches(1.35), Inches(0.20),
     "TRIAGED  +  AUTO-CONVERT", 7, bold=True, color=PURPLE,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# 4 branches
branch_y = Inches(4.95)
# Diagonal arrows
for x_to in [Inches(2.0), Inches(5.4), Inches(8.8), Inches(12.2)]:
    arrow_ind = slide12.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, x_to, branch_y - Inches(0.2), Inches(0.3), Inches(0.2))
    arrow_ind.fill.solid(); arrow_ind.fill.fore_color.rgb = AVA_ORANGE
    arrow_ind.line.fill.background()

# 4 destination paths — each shows: condition → issue type → mini board
def render_destination(slide, x, y, w, condition, condition_c, issue_type, type_color,
                       board_label, columns, sample_card_col, board_color):
    """Draw one of the 4 destination paths."""
    path_h = Inches(2.3)
    round_rect(slide, x, y, w, path_h, WHITE, board_color, Pt(2), 0.04)

    # Condition badge at top
    round_rect(slide, x, y, w, Inches(0.32), condition_c, radius=0.04)
    text(slide, x, y + Inches(0.04), w, Inches(0.22),
         condition, 9, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Issue type
    text(slide, x + Inches(0.1), y + Inches(0.40), w - Inches(0.2), Inches(0.18),
         "Auto-converts to", 8, italic=True, color=GREY_LABEL, align=PP_ALIGN.CENTER)
    round_rect(slide, x + Inches(0.4), y + Inches(0.60), w - Inches(0.8), Inches(0.30),
               type_color, radius=0.06)
    text(slide, x + Inches(0.4), y + Inches(0.60), w - Inches(0.8), Inches(0.30),
         issue_type, 11, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Goes to label
    text(slide, x + Inches(0.1), y + Inches(0.97), w - Inches(0.2), Inches(0.18),
         "Goes to", 8, italic=True, color=GREY_LABEL, align=PP_ALIGN.CENTER)
    text(slide, x + Inches(0.1), y + Inches(1.13), w - Inches(0.2), Inches(0.20),
         board_label, 9, bold=True, color=board_color, align=PP_ALIGN.CENTER)

    # Mini kanban board
    board_y = y + Inches(1.40)
    board_h = Inches(0.85)
    round_rect(slide, x + Inches(0.1), board_y, w - Inches(0.2), board_h,
               GREY_SOFT, GREY_BORDER, Pt(0.5), 0.05)
    n_cols = len(columns)
    col_w = (w - Inches(0.3)) / n_cols
    for i, col_name in enumerate(columns):
        col_x = x + Inches(0.15) + i * col_w
        # header
        rect(slide, col_x, board_y + Inches(0.04), col_w - Inches(0.02), Inches(0.18), board_color)
        text(slide, col_x, board_y + Inches(0.04), col_w - Inches(0.02), Inches(0.18),
             col_name, 7, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # sample card
        if i == sample_card_col:
            rect(slide, col_x + Inches(0.02), board_y + Inches(0.28), col_w - Inches(0.06), Inches(0.5),
                 board_color, line=board_color)
            text(slide, col_x + Inches(0.02), board_y + Inches(0.30), col_w - Inches(0.06), Inches(0.20),
                 issue_type[:6], 6, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            text(slide, col_x + Inches(0.02), board_y + Inches(0.48), col_w - Inches(0.06), Inches(0.20),
                 "your card", 6, italic=True, color=WHITE, align=PP_ALIGN.CENTER)


# Branch 1: Size = S → Story → Squad Backlog
render_destination(slide12, Inches(0.4), branch_y, Inches(3.1),
                   "Size = S  ·  < 5 days",
                   AVA_GREEN,
                   "STORY (no Epic parent)",
                   AVA_GREEN,
                   "Squad Backlog · DIRECT (skip portfolio)",
                   ["TODO", "WIP", "REVIEW", "DONE"],
                   sample_card_col=0,
                   board_color=AVA_GREEN)

# Branch 2: Size = M/L → Epic → Portfolio Backlog
render_destination(slide12, Inches(3.7), branch_y, Inches(3.1),
                   "Size = M / L",
                   JIRA_BLUE,
                   "EPIC + placeholder Stories",
                   JIRA_BLUE,
                   "Portfolio Backlog · Gate 2 next",
                   ["INTAKE", "TRIAGE", "BACKLOG", "READY", "FLIGHT"],
                   sample_card_col=2,
                   board_color=JIRA_BLUE)

# Branch 3: Size = XL → Initiative → Enterprise Dashboard
render_destination(slide12, Inches(7.0), branch_y, Inches(3.1),
                   "Size = XL  ·  cross-team",
                   ORANGE,
                   "INITIATIVE + placeholder Epics",
                   ORANGE,
                   "Enterprise Dashboard · Gate 1 first",
                   ["IDEA", "GATE 1", "GATE 2", "RUN", "DONE"],
                   sample_card_col=1,
                   board_color=ORANGE)

# Branch 4: Type = Incident → Bug → Squad Expedite
render_destination(slide12, Inches(10.3), branch_y, Inches(3.1),
                   "Type = Incident",
                   PINK,
                   "BUG · CoS = Expedite",
                   PINK,
                   "Squad Expedite swimlane · WIP-1",
                   ["URGENT", "WIP", "REVIEW", "DONE"],
                   sample_card_col=1,
                   board_color=PINK)

# Footer
text(slide12, Inches(0.4), Inches(7.30), Inches(13), Inches(0.18),
     "Requester never picks Story / Epic / Initiative — Demand placeholder converts automatically based on Size + Type at Triage.",
     9, italic=True, color=GREY_LABEL, align=PP_ALIGN.CENTER)


# ============================================================================
# SLIDE 13 — PAIN SCORE FORMULA · DECODED (w₁/w₂/w₃ explained)
# ============================================================================
slide13 = prs.slides.add_slide(prs.slide_layouts[6])
header_band(slide13, "PAIN SCORE  ·  FORMULA DECODED",
            "What every variable means · why w₁ · w₂ · w₃ exist · how they're set")

# Big formula at top with annotations
fy = Inches(1.15)
round_rect(slide13, Inches(0.4), fy, Inches(12.55), Inches(1.05),
           NAVY, radius=0.04)
text(slide13, Inches(0.4), fy + Inches(0.06), Inches(12.55), Inches(0.22),
     "THE FORMULA  ·  computed at TRIAGED status",
     9, bold=True, color=YELLOW, align=PP_ALIGN.CENTER)
text(slide13, Inches(0.4), fy + Inches(0.32), Inches(12.55), Inches(0.5),
     "H(x)  =  ( BV·w₁  +  UX·w₂  +  Safety·w₃ )  /  Effort",
     30, bold=True, color=YELLOW,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
text(slide13, Inches(0.4), fy + Inches(0.85), Inches(12.55), Inches(0.18),
     "result range typically 0.0 – 5.0 · higher = more painful · drives backlog ranking",
     8, italic=True, color=AVA_GREEN_LIGHT, align=PP_ALIGN.CENTER)

# Section 1 — VARIABLES (BV · UX · Safety · Effort)
sec1_y = Inches(2.35)
text(slide13, Inches(0.4), sec1_y, Inches(13), Inches(0.22),
     "1 ·  THE FOUR VARIABLES (each scored 1–5 at Triage)", 11, bold=True, color=AVA_GREEN)

vars_data = [
    ("BV", "Business Value", "1 minimal · 5 transformational",
     "Quantified business impact · $ revenue / cost · NPS / CSAT · volume affected · compliance gap.\n"
     "Numerator — increases Pain Score linearly with weight w₁.",
     "from Business Value field · scored at Triage", AVA_GREEN, WHITE),
    ("UX", "User Experience", "1 cosmetic · 5 blocker",
     "User-perceived pain · frustration · workarounds needed · time-to-task · error frequency.\n"
     "Numerator — weighted highest in CS (customer-facing).",
     "scored by Champion at Triage", JIRA_BLUE, WHITE),
    ("Safety", "Utility-sector risk", "1 no risk · 5 life-critical",
     "Safety dimensions specific to utilities · field crew exposure · public hazard · regulatory breach · cyber.\n"
     "Numerator — never zeroed out · Avangrid is utility.",
     "scored by Champion + Architect at Triage", RED, WHITE),
    ("Effort", "Complexity / cost", "1 trivial · 5 cross-team",
     "Sizing dimensions · effort hours · apps touched · vendor dependency · risk of changes.\n"
     "DENOMINATOR — divides the pain · forces ROI thinking.",
     "sized by SMEs at Triage · t-shirt size", PURPLE, WHITE),
]
vy = sec1_y + Inches(0.28)
v_h = Inches(1.1)
for i, (sym, lbl, scale, body, src, fill, txc) in enumerate(vars_data):
    row = i // 2
    col = i % 2
    vx = Inches(0.4) + col * Inches(6.36)
    vy_loc = vy + row * (v_h + Inches(0.1))
    round_rect(slide13, vx, vy_loc, Inches(6.18), v_h, WHITE, fill, Pt(2), 0.04)
    round_rect(slide13, vx, vy_loc, Inches(0.95), v_h, fill, radius=0.04)
    text(slide13, vx, vy_loc + Inches(0.08), Inches(0.95), Inches(0.4),
         sym, 22, bold=True, color=txc, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(slide13, vx, vy_loc + Inches(0.55), Inches(0.95), Inches(0.4),
         lbl, 8, italic=True, color=txc, align=PP_ALIGN.CENTER)
    text(slide13, vx + Inches(1.1), vy_loc + Inches(0.04), Inches(4.95), Inches(0.22),
         scale, 9, bold=True, color=fill)
    text(slide13, vx + Inches(1.1), vy_loc + Inches(0.24), Inches(4.95), Inches(0.55),
         body, 8, color=TEXT)
    text(slide13, vx + Inches(1.1), vy_loc + Inches(0.85), Inches(4.95), Inches(0.20),
         "▸ " + src, 7, italic=True, color=GREY_LABEL)

# Section 2 — WEIGHTS (w₁ · w₂ · w₃)
sec2_y = Inches(4.85)
text(slide13, Inches(0.4), sec2_y, Inches(13), Inches(0.22),
     "2 ·  THE WEIGHTS  (w₁ · w₂ · w₃) — calibrate strategic emphasis per Value Stream",
     11, bold=True, color=AVA_GREEN)

# Left half: What weights are
round_rect(slide13, Inches(0.4), sec2_y + Inches(0.28), Inches(6.18), Inches(2.15),
           AVA_GREEN_LIGHT, AVA_GREEN, Pt(2), 0.04)
text(slide13, Inches(0.55), sec2_y + Inches(0.34), Inches(6), Inches(0.22),
     "WHAT THEY ARE", 9, bold=True, color=AVA_GREEN_DARK)
text(slide13, Inches(0.55), sec2_y + Inches(0.55), Inches(5.9), Inches(0.45),
     "Weights are decimal numbers between 0 and 1 that calibrate how much each numerator "
     "component contributes to the final score. They normalize so that w₁ + w₂ + w₃ = 1.0.",
     9, color=TEXT)

text(slide13, Inches(0.55), sec2_y + Inches(1.05), Inches(6), Inches(0.22),
     "MEANING", 9, bold=True, color=AVA_GREEN_DARK)
items = [
    ("w₁", "weight of Business Value", "how much BV matters", AVA_GREEN),
    ("w₂", "weight of User Experience", "how much UX matters", JIRA_BLUE),
    ("w₃", "weight of Safety", "how much Safety matters", RED),
]
iy = sec2_y + Inches(1.27)
for sym, what, why, c in items:
    text(slide13, Inches(0.55), iy, Inches(0.45), Inches(0.20),
         sym, 11, bold=True, color=c)
    text(slide13, Inches(1.1), iy, Inches(2.4), Inches(0.20),
         what, 9, bold=True, color=c)
    text(slide13, Inches(3.6), iy, Inches(2.9), Inches(0.20),
         "= " + why, 9, italic=True, color=GREY_TEXT)
    iy += Inches(0.25)

# Right half: WHY they exist + values per VS
round_rect(slide13, Inches(6.78), sec2_y + Inches(0.28), Inches(6.18), Inches(2.15),
           JIRA_BLUE_LIGHT, JIRA_BLUE, Pt(2), 0.04)
text(slide13, Inches(6.93), sec2_y + Inches(0.34), Inches(6), Inches(0.22),
     "WHY THEY EXIST", 9, bold=True, color=JIRA_BLUE)
text(slide13, Inches(6.93), sec2_y + Inches(0.55), Inches(5.9), Inches(0.45),
     "Different Value Streams have different priorities. CS is customer-facing → UX matters most. "
     "ITBA G&E (field) → Safety matters most. Weights let each VS tune the formula without changing it.",
     9, color=TEXT)

text(slide13, Inches(6.93), sec2_y + Inches(1.05), Inches(6), Inches(0.22),
     "VALUES PER VALUE STREAM", 9, bold=True, color=JIRA_BLUE)

vs_weights = [
    ("CS  ·  4 DVS (current)", "0.35", "0.40", "0.25", "UX-heavy", JIRA_BLUE, True),
    ("ITBA Field  (hypothetical)", "0.30", "0.20", "0.50", "Safety-heavy", AVA_GREEN, False),
    ("ITBA Corporate (hypothetical)", "0.40", "0.35", "0.25", "BV-balanced", PURPLE, False),
]
wy = sec2_y + Inches(1.27)
for vs, w1, w2, w3, note, c, current in vs_weights:
    text(slide13, Inches(6.93), wy, Inches(2.5), Inches(0.20),
         ("● " if current else "○ ") + vs, 8, bold=current, color=c)
    text(slide13, Inches(9.5), wy, Inches(0.5), Inches(0.20),
         w1, 9, bold=True, color=c, align=PP_ALIGN.CENTER)
    text(slide13, Inches(10.05), wy, Inches(0.5), Inches(0.20),
         w2, 9, bold=True, color=c, align=PP_ALIGN.CENTER)
    text(slide13, Inches(10.6), wy, Inches(0.5), Inches(0.20),
         w3, 9, bold=True, color=c, align=PP_ALIGN.CENTER)
    text(slide13, Inches(11.15), wy, Inches(1.7), Inches(0.20),
         note, 8, italic=True, color=GREY_TEXT)
    wy += Inches(0.25)

# Worked example — bottom strip
ex_y = Inches(7.05)
round_rect(slide13, Inches(0.4), ex_y, Inches(12.55), Inches(0.42),
           NAVY, radius=0.04)
text(slide13, Inches(0.55), ex_y + Inches(0.04), Inches(8), Inches(0.18),
     "WORKED EXAMPLE", 8, bold=True, color=YELLOW)
text(slide13, Inches(0.55), ex_y + Inches(0.20), Inches(12), Inches(0.20),
     "BV=4 (high) · UX=5 (blocker) · Safety=2 (low) · Effort=3 (medium)  →  H(x) = (4·0.35 + 5·0.40 + 2·0.25) / 3  =  (1.4 + 2.0 + 0.5) / 3  =  3.9 / 3  =  1.30",
     9, color=WHITE)


# ============================================================================
# SLIDE 14 — PMO / TRAVIS INTEGRATION OPTIONS
# ============================================================================
slide14 = prs.slides.add_slide(prs.slide_layouts[6])
header_band(slide14, "PMO  ·  TRAVIS INTEGRATION OPTIONS",
            "Travis (PMO Lead · Large Projects · PDD) is another demand source — how do we handle?")

# Context strip
ctx_y = Inches(1.15)
round_rect(slide14, Inches(0.4), ctx_y, Inches(12.55), Inches(0.7),
           AVA_GREEN_LIGHT, AVA_GREEN, Pt(1.5), 0.04)
text(slide14, Inches(0.55), ctx_y + Inches(0.06), Inches(12), Inches(0.22),
     "TRAVIS  ·  PMO Lead — runs the PDD process for large projects (annual cycle)",
     11, bold=True, color=AVA_GREEN_DARK)
text(slide14, Inches(0.55), ctx_y + Inches(0.30), Inches(12.2), Inches(0.36),
     "Receives heavy demand from Customer Services for major initiatives. PDD = legacy process · annual cycle · "
     "SharePoint-based. Bridge to legacy. Vihn Nguyen (Director IT Program Delivery) co-runs.  ⚠ Travis was "
     "not yet interviewed in baseline — this is a gap.",
     9, italic=True, color=TEXT)

# 3 Option cards
opt_y = Inches(2.0)
opt_h = Inches(4.45)
opt_w = (Inches(12.55) - Inches(0.3)) / 3

options_data = [
    # (label, recommendation, color, fill, content_blocks)
    ("OPTION A", "Direct Jira Intake", "Travis fills the form like everyone else",
     GREY_LABEL,
     [
         ("HOW IT WORKS",
          "Travis (or his PMO team) opens a Demand ticket via the same Jira form. PDD docs "
          "attached to the ticket. Auto-routes based on Value Stream selection. Goes through "
          "Triage like any other request."),
         ("PROS",
          "✓ Single front door · zero exception\n"
          "✓ Everything visible on one board\n"
          "✓ No integration to maintain\n"
          "✓ KPIs (DoR pass rate · channel adoption) include PMO"),
         ("CONS",
          "✗ PDD annual cycle clashes with weekly intake\n"
          "✗ Travis already pre-screens — Triage redundant\n"
          "✗ Forces PDD format into ticket fields\n"
          "✗ Loses PDD artifact richness"),
         ("BEST FOR",
          "Small PMO · low PDD volume · or as transition path"),
     ]),
    ("OPTION B", "Separate PMO Process", "PMO process feeds Jira via integration",
     ORANGE,
     [
         ("HOW IT WORKS",
          "PMO runs PDD process untouched (annual cycle · SharePoint · committee approvals). "
          "At end of PDD cycle, approved projects bulk-create Initiatives in Jira. PMO owns "
          "the SharePoint workflow; Jira receives output only."),
         ("PROS",
          "✓ PDD discipline preserved\n"
          "✓ Travis's existing artifacts kept\n"
          "✓ No disruption to PMO operations\n"
          "✓ Large projects clearly distinct"),
         ("CONS",
          "✗ Two systems · two processes\n"
          "✗ Annual lag on roadmap visibility\n"
          "✗ Duplicate work to map PDD → Jira fields\n"
          "✗ KPIs are split"),
         ("BEST FOR",
          "Heavy PDD volume · resistant PMO culture · transitional"),
     ]),
    ("OPTION C  ★", "Hybrid — PMO-Pre-Screened",
     "Travis opens Initiative directly · skips Triage gate",
     AVA_GREEN,
     [
         ("HOW IT WORKS",
          "Travis (or PMO designate) creates Initiative-type tickets directly in Jira (skip "
          "Demand placeholder) with PDD docs as attachments. Auto-skips Triage (status moves "
          "from NEW → READY_FOR_GATE_1). Goes to Gate 1 (Shark Tank) immediately."),
         ("PROS",
          "✓ Single front door (Jira) preserved\n"
          "✓ PDD pre-screen recognized · no double-work\n"
          "✓ All artifacts on Jira ticket\n"
          "✓ Roadmap visibility cross-VS\n"
          "✓ KPIs unified (PMO + DVS together)"),
         ("CONS",
          "✗ Special permission scheme needed\n"
          "✗ Training Travis's team on Jira"),
         ("BEST FOR",
          "Avangrid Digital Products today · recommended by NTT"),
     ]),
]

for i, (label, title, subtitle, accent, blocks) in enumerate(options_data):
    ox = Inches(0.4) + i * (opt_w + Inches(0.15))
    is_recommended = (i == 2)
    border_w = Pt(3) if is_recommended else Pt(1.5)
    round_rect(slide14, ox, opt_y, opt_w, opt_h, WHITE, accent, border_w, 0.04)
    round_rect(slide14, ox, opt_y, opt_w, Inches(0.85), accent, radius=0.04)
    text(slide14, ox, opt_y + Inches(0.06), opt_w, Inches(0.22),
         label, 10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    text(slide14, ox, opt_y + Inches(0.28), opt_w, Inches(0.32),
         title, 14, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(slide14, ox + Inches(0.1), opt_y + Inches(0.62), opt_w - Inches(0.2), Inches(0.22),
         subtitle, 8, italic=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Blocks
    by = opt_y + Inches(0.95)
    block_h = (opt_h - Inches(1.0)) / len(blocks)
    for label_b, content in blocks:
        text(slide14, ox + Inches(0.1), by + Inches(0.04), opt_w - Inches(0.2), Inches(0.18),
             label_b, 8, bold=True, color=accent)
        text(slide14, ox + Inches(0.1), by + Inches(0.22), opt_w - Inches(0.2), block_h - Inches(0.24),
             content, 8, color=TEXT)
        by += block_h

# Recommended banner on Option C
rec_y = opt_y - Inches(0.15)
rec_x = Inches(0.4) + 2 * (opt_w + Inches(0.15)) + opt_w - Inches(1.5)
round_rect(slide14, rec_x, rec_y, Inches(1.5), Inches(0.3), YELLOW, NAVY, Pt(1), 0.4)
text(slide14, rec_x, rec_y, Inches(1.5), Inches(0.3),
     "RECOMMENDED",
     8, bold=True, color=NAVY, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# Bottom — governance for Option C
gov_y = opt_y + opt_h + Inches(0.1)
round_rect(slide14, Inches(0.4), gov_y, Inches(12.55), Inches(0.55),
           NAVY, radius=0.04)
text(slide14, Inches(0.55), gov_y + Inches(0.06), Inches(13), Inches(0.18),
     "GOVERNANCE FOR OPTION C  ·  to validate with Travis", 8, bold=True, color=YELLOW)
text(slide14, Inches(0.55), gov_y + Inches(0.24), Inches(12.5), Inches(0.30),
     "Forums: PDD Governance (annual + checkpoints) + Portfolio Committee (monthly · Travis attends)  ·  "
     "Roles: Travis = PMO Lead · Vihn = Director IT Program Delivery · Vitor sponsors  ·  "
     "Artifacts: PDD docs + Initiative ticket + Gate 1 brief  ·  KPIs: PDD throughput + Initiative aging + cross-VS dependency map",
     8, color=WHITE)


prs.save(OUT)
print(f"Saved: {OUT}")
print(f"Slides: {len(prs.slides)}")
