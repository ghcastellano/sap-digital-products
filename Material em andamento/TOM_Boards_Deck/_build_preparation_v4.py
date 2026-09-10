"""
Preparation & Planning · v4 (27 slides)
========================================

Tier-1 storytelling (action titles · Pyramid Principle · MECE) + operational depth
(CS vs ITBA differentiation · 6 capabilities · framework alignment).

Design system:
  - Layout: prs.slide_layouts[14] = 'Report + Image' (clean top · logos top-right inherited)
  - Title: 28pt IberPangea Text Medium · AVANGRID_GREEN · bold · y=0.13
  - Subtitle: 20pt IberPangea Text Medium · DARK_TEXT · bold · y=0.55
  - Palette: AVANGRID_GREEN (CS) · HEADER_BLUE (ITBA) · DARK_TEXT + neutrals · CoS swatches · zero purple
  - Strip stroke 100% (no border ghosts)
  - Reuses existing visual patterns from intake slides (Portfolio Board · Step-level policies ·
    Team Kanban · Form view · Field Guide)

Aligned with the Demand Management section of the TOM
(terms · roles · formulas · gates · classifications).

INPUT  · TOM MAP workbook - final.pptx
OUTPUT · TOM MAP workbook - final v4.pptx
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import zipfile, shutil

SRC = "/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Materiais de referencia/_clean_79.pptx"
DST = "/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Materiais de referencia/TOM MAP workbook - versao final da final.pptx"

# ============================== Palette ==============================
# Brand 3-hue + neutrals · matches intake slides exactly
AVANGRID_GREEN = RGBColor(0x00, 0x8C, 0x38)   # CS · primary
GREEN_DARK     = RGBColor(0x00, 0x6B, 0x3F)   # CS shade
GREEN_MID      = RGBColor(0x00, 0x89, 0x50)   # CS mid
GREEN_LIGHT    = RGBColor(0x6F, 0xC1, 0x8F)   # CS tint
GREEN_BG       = RGBColor(0xE6, 0xF2, 0xEB)   # CS pale bg
GREEN_PALE     = RGBColor(0xF2, 0xF8, 0xF4)   # CS very pale

HEADER_BLUE    = RGBColor(0x2C, 0x57, 0x8B)   # ITBA · primary
HEADER_BLUE_DK = RGBColor(0x1B, 0x3C, 0x66)   # ITBA shade
BLUE_LIGHT     = RGBColor(0x7E, 0x9C, 0xC4)   # ITBA tint
BLUE_BG        = RGBColor(0xEE, 0xF3, 0xF9)   # ITBA pale
BLUE_PALE      = RGBColor(0xF6, 0xFA, 0xFD)   # ITBA very pale

NAVY_DARK      = RGBColor(0x0F, 0x2A, 0x47)   # navy · gates anchor (matches Portfolio Board slide 41)

# Neutrals
WHITE          = RGBColor(0xFF, 0xFF, 0xFF)
DARK_TEXT      = RGBColor(0x2C, 0x3E, 0x50)
TEXT           = RGBColor(0x17, 0x2B, 0x4D)
GREY_LIGHT     = RGBColor(0xF2, 0xF6, 0xFA)
GREY_BG        = RGBColor(0xF4, 0xF5, 0xF7)
GREY_BORDER    = RGBColor(0xC9, 0xD2, 0xDD)
GREY_LABEL     = RGBColor(0x6B, 0x77, 0x8C)
GREY_MUTED     = RGBColor(0x99, 0x99, 0x99)
GREY_TEXT      = RGBColor(0x52, 0x60, 0x77)

# Class of Service swatches (matches Portfolio Board legend)
COS_GREEN      = RGBColor(0x4F, 0x91, 0x4F)   # Standard
COS_BLUE       = RGBColor(0x3B, 0x6E, 0x8F)   # Fixed-Date
COS_RED        = RGBColor(0xB0, 0x35, 0x35)   # Expedite
COS_ORANGE     = RGBColor(0xE9, 0x8B, 0x2F)   # Intangible

# Gates + accents
GATE_RED       = RGBColor(0xB0, 0x35, 0x35)
GATE_RED_BG    = RGBColor(0xFB, 0xEC, 0xEC)
BYPASS_ORANGE  = RGBColor(0xF0, 0x86, 0x29)
BYPASS_BG      = RGBColor(0xFD, 0xEC, 0xD4)
INSIGHT_AMBER  = RGBColor(0xE0, 0x9F, 0x3E)
INSIGHT_BG     = RGBColor(0xFD, 0xF7, 0xEA)

TITLE_FONT = "IberPangea Text Medium"
BODY_FONT  = "Arial"

# ============================== Core helpers ==============================
def _no_line(shape):
    """Strip stroke entirely · prevents border ghost (per memory: feedback_pptx_border_overflow.md)."""
    shape.line.fill.background()

def rect(slide, x, y, w, h, fill=None):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    s.shadow.inherit = False
    if fill is None: s.fill.background()
    else: s.fill.solid(); s.fill.fore_color.rgb = fill
    _no_line(s)
    return s

def round_rect(slide, x, y, w, h, fill=None, radius=0.06):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    s.shadow.inherit = False
    s.adjustments[0] = radius
    if fill is None: s.fill.background()
    else: s.fill.solid(); s.fill.fore_color.rgb = fill
    _no_line(s)
    return s

def oval(slide, x, y, w, h, fill=None):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, w, h)
    s.shadow.inherit = False
    if fill is None: s.fill.background()
    else: s.fill.solid(); s.fill.fore_color.rgb = fill
    _no_line(s)
    return s

def line(slide, x1, y1, x2, y2, color=GREY_LABEL, weight=1.0):
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    ln.line.color.rgb = color
    ln.line.width = Pt(weight)
    return ln

def text(slide, x, y, w, h, txt, size=10, bold=False, italic=False,
         color=DARK_TEXT, font=BODY_FONT,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
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

# ============================== Slide chrome ==============================
def slide_title(slide, topic, summary, section=None):
    """3-line title:
       Line 1 (24pt green) = LIFECYCLE MODEL – PREPARATION & PLANNING
       Line 2 (10pt grey)  = SECTION · {section}
       Line 3 (13pt dark)  = TOPIC: summary
    """
    # Line 1 · always fixed
    text(slide, Inches(0.40), Inches(0.10), Inches(10.20), Inches(0.36),
         "LIFECYCLE MODEL – PREPARATION & PLANNING",
         size=24, bold=True, color=AVANGRID_GREEN, font=TITLE_FONT)
    # Line 2 · section name
    if section:
        text(slide, Inches(0.40), Inches(0.46), Inches(10.20), Inches(0.20),
             f"SECTION · {section.upper()}",
             size=10, italic=True, bold=True, color=GREY_LABEL, font=BODY_FONT)
    # Line 3 · topic + summary
    composed = f"{topic.upper()}: {summary.upper()}" if topic else summary.upper()
    text(slide, Inches(0.40), Inches(0.68), Inches(10.20), Inches(0.55),
         composed, size=13, bold=True, color=DARK_TEXT, font=TITLE_FONT)

def from_phase1_strip(slide, y=1.40):
    """Strip cinza referenciando o que vem da Phase 1 (intake bridge explicit)."""
    text(slide, Inches(0.40), Inches(y), Inches(12.55), Inches(0.18),
         "FROM PHASE 1 · DEMAND MGMT  ·  Pain Score · CoS · Type · Size · OPCO · Sponsor · R/G/T",
         size=8, italic=True, color=GREY_LABEL, font=BODY_FONT)

def insight_callout(slide, x, y, w, h, headline, supporting):
    """McKinsey-style insight box · amber accent · headline conclusion + supporting evidence."""
    round_rect(slide, Inches(x), Inches(y), Inches(w), Inches(h), INSIGHT_BG, radius=0.04)
    rect(slide, Inches(x), Inches(y), Inches(0.10), Inches(h), INSIGHT_AMBER)
    text(slide, Inches(x + 0.20), Inches(y + 0.06), Inches(w - 0.30), Inches(0.18),
         "INSIGHT", size=7, bold=True, color=INSIGHT_AMBER, font=BODY_FONT)
    text(slide, Inches(x + 0.20), Inches(y + 0.24), Inches(w - 0.30), Inches(0.30),
         headline, size=10, bold=True, color=DARK_TEXT, font=BODY_FONT)
    text(slide, Inches(x + 0.20), Inches(y + 0.55), Inches(w - 0.30), Inches(h - 0.60),
         supporting, size=8, color=GREY_TEXT, font=BODY_FONT)

def cos_legend(slide, y=7.05):
    """Same legend bar from Portfolio Board (slide 41)."""
    text(slide, Inches(0.40), Inches(y), Inches(2.0), Inches(0.18),
         "CLASS OF SERVICE", size=8, bold=True, color=GREY_LABEL, font=BODY_FONT)
    swatches = [("Standard", COS_GREEN), ("Fixed-Date", COS_BLUE),
                ("Expedite", COS_RED), ("Intangible", COS_ORANGE)]
    sx = 2.20
    for label, color in swatches:
        rect(slide, Inches(sx), Inches(y + 0.02), Inches(0.30), Inches(0.16), color)
        text(slide, Inches(sx + 0.36), Inches(y), Inches(1.4), Inches(0.18),
             label, size=8, color=DARK_TEXT, font=BODY_FONT)
        sx += 1.85

def footer_bar(slide, n, total=29):
    text(slide, Inches(11.5), Inches(7.18), Inches(1.6), Inches(0.18),
         f"{n:02d} / {total:02d}", size=8, color=GREY_LABEL, font=BODY_FONT, align=PP_ALIGN.RIGHT)

# ============================== Pattern library ==============================
def avatar(slide, x, y, d, color, initials, size_factor=20):
    """Person circle with initials."""
    oval(slide, Inches(x), Inches(y), Inches(d), Inches(d), color)
    text(slide, Inches(x), Inches(y), Inches(d), Inches(d), initials,
         size=int(d * size_factor), bold=True, color=WHITE, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

def chip(slide, x, y, w, label, color, text_color=WHITE, h=0.22, size=7, italic=False):
    """Pill chip · used for status/tag/lane labels."""
    round_rect(slide, Inches(x), Inches(y), Inches(w), Inches(h), color, radius=0.30)
    text(slide, Inches(x), Inches(y), Inches(w), Inches(h), label,
         size=size, bold=True, italic=italic, color=text_color, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

def board_column(slide, x, y, w, h, header_label, header_color, body_blocks,
                 body_size=9, header_size=12, label_color=GREY_LABEL):
    """
    Portfolio Board / Step-level policies column · matches slide 41+46 visual.
    body_blocks = list of (LABEL, content) tuples.
    """
    header_h = 0.42
    rect(slide, Inches(x), Inches(y), Inches(w), Inches(header_h), header_color)
    text(slide, Inches(x), Inches(y + 0.04), Inches(w), Inches(header_h - 0.08),
         header_label, size=header_size, bold=True, color=WHITE, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    rect(slide, Inches(x), Inches(y + header_h), Inches(w), Inches(h - header_h), GREY_LIGHT)
    by = y + header_h + 0.15
    block_h = (h - header_h - 0.20) / max(1, len(body_blocks))
    for label, content in body_blocks:
        text(slide, Inches(x + 0.12), Inches(by), Inches(w - 0.24), Inches(0.18),
             label, size=7, bold=True, color=label_color, font=BODY_FONT)
        text(slide, Inches(x + 0.12), Inches(by + 0.18), Inches(w - 0.24), Inches(block_h - 0.22),
             content, size=body_size, color=TEXT, font=BODY_FONT)
        by += block_h

def gate_divider_v(slide, x, y, h, label):
    """Vertical red gate divider · matches Portfolio Board gates style."""
    rect(slide, Inches(x), Inches(y - 0.30), Inches(0.04), Inches(h + 0.30), GATE_RED)
    text(slide, Inches(x - 1.10), Inches(y - 0.30), Inches(2.20), Inches(0.22),
         label, size=9, bold=True, color=GATE_RED, font=BODY_FONT, align=PP_ALIGN.CENTER)

def bypass_bar(slide, y, message):
    """Orange bypass-style bar · matches Portfolio Board bottom bar."""
    round_rect(slide, Inches(0.40), Inches(y), Inches(12.55), Inches(0.34), BYPASS_BG, radius=0.04)
    rect(slide, Inches(0.40), Inches(y), Inches(0.10), Inches(0.34), BYPASS_ORANGE)
    text(slide, Inches(0.55), Inches(y), Inches(12.40), Inches(0.34),
         message, size=9, bold=True, italic=True, color=BYPASS_ORANGE, font=BODY_FONT,
         anchor=MSO_ANCHOR.MIDDLE)

# ============================== Driver placeholder ==============================
def _slide(prs):
    """'Report + Image' layout · clean top · keep top-right logos · mask inherited line + color bands."""
    s = prs.slides.add_slide(prs.slide_layouts[14])
    # Remove all placeholder shapes (Click to add prompts)
    for ph in list(s.placeholders):
        sp = ph.element
        sp.getparent().remove(sp)
    # Mask the inherited horizontal line at y=0.89 (Conector recto from layout)
    # and the off-screen color bands bleeding into visible area at left edge
    rect(s, Inches(0.0), Inches(0.0), Inches(0.10), Inches(7.50), WHITE)   # left edge mask
    rect(s, Inches(0.0), Inches(0.85), Inches(13.33), Inches(0.10), WHITE) # horizontal line mask
    return s


# ============================================================================
# SECTION 1 · OPENER  (slides 01-03)
# ============================================================================

# 01 · Bridge from Phase 1 — "list is not shippable work"
def build_pp01(prs):
    s = _slide(prs)
    slide_title(s,
        "From a prioritized list to shippable work",
        "What Phase 1 hands over · what Phase 2 adds · what flows to Execution",
        section="Opener")

    # 3-phase compact timeline
    tl_y = 1.65
    tl_w = 12.55
    tl_x = 0.40
    rect(s, Inches(tl_x), Inches(tl_y + 0.55), Inches(tl_w), Emu(20000), GREY_BORDER)
    phases = [
        ("1", "DEMAND MANAGEMENT",     "intake · Gate 1 · Gate 2", GREY_MUTED, False),
        ("2", "PREPARATION & PLANNING", "this section",            AVANGRID_GREEN, True),
        ("3", "EXECUTION",             "build · test · release",   GREY_MUTED, False),
    ]
    pw = (tl_w - 0.40) / 3
    circle_bottom_y = tl_y + 1.05
    label_y = tl_y + 1.18
    sub_y   = tl_y + 1.42
    for i, (num, name, sub, color, active) in enumerate(phases):
        x = tl_x + 0.10 + i * (pw + 0.10)
        d = 0.85 if active else 0.65
        cy = circle_bottom_y - d
        cx = x + pw/2 - d/2
        oval(s, Inches(cx), Inches(cy), Inches(d), Inches(d), color)
        text(s, Inches(cx), Inches(cy), Inches(d), Inches(d), num,
             size=int(d * 28), bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, Inches(x), Inches(label_y), Inches(pw), Inches(0.22),
             name, size=11, bold=True,
             color=color if active else GREY_LABEL, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, Inches(x), Inches(sub_y), Inches(pw), Inches(0.22),
             sub, size=8, italic=True, color=GREY_LABEL, font=BODY_FONT, align=PP_ALIGN.CENTER)
        if active:
            chip(s, x + pw/2 - 0.95, cy - 0.32, 1.90, "▼ COVERED IN THIS SECTION", GATE_RED, WHITE, h=0.22, size=8)

    # Three big boxes: PHASE 1 OUTPUT → PHASE 2 ADDS → READY FOR EXECUTION
    box_y = 3.40
    box_h = 3.45
    n_boxes = 3
    gap = 0.30
    bw = (12.55 - 2 * gap) / n_boxes

    # Box 1 · PHASE 1 OUTPUT (what we receive)
    bx = 0.40
    round_rect(s, Inches(bx), Inches(box_y), Inches(bw), Inches(box_h), WHITE, radius=0.04)
    rect(s, Inches(bx), Inches(box_y), Inches(bw), Inches(0.55), AVANGRID_GREEN)
    text(s, Inches(bx + 0.15), Inches(box_y + 0.06), Inches(bw - 0.30), Inches(0.22),
         "WHAT FLOWS IN  ·  Phase 1 output", size=11, bold=True, color=WHITE, font=BODY_FONT)
    text(s, Inches(bx + 0.15), Inches(box_y + 0.30), Inches(bw - 0.30), Inches(0.22),
         "Output of Gate 2 · GO decision · classified demand",
         size=8, italic=True, color=GREEN_PALE, font=BODY_FONT)
    in_items = [
        ("Pain Score",            "BV · UX · Effort · OPCOmult — preliminary"),
        ("Class of Service",      "Standard · Fixed-Date · Expedite · Intangible — best-guess"),
        ("Type",                  "Production Support · Regulatory · Non-Regulatory"),
        ("Size",                  "S · M · L  (XL → PPM Now · out of scope)"),
        ("OPCO scope",            "single · regional (2-4) · enterprise (5+) — drives multiplier"),
        ("Sponsor",               "named user · acceptance criteria signed at Funnel"),
        
        ("R/G/T classification",  "Run · Grow · Transform — strategic lens"),
    ]
    iy = box_y + 0.65
    rh = (box_h - 0.75) / len(in_items)
    for lbl, body in in_items:
        oval(s, Inches(bx + 0.20), Inches(iy + 0.06), Inches(0.16), Inches(0.16), AVANGRID_GREEN)
        text(s, Inches(bx + 0.45), Inches(iy + 0.02), Inches(bw - 0.55), Inches(0.20),
             lbl, size=8.5, bold=True, color=DARK_TEXT, font=BODY_FONT)
        text(s, Inches(bx + 0.45), Inches(iy + 0.20), Inches(bw - 0.55), Inches(rh - 0.22),
             body, size=8, color=GREY_TEXT, font=BODY_FONT)
        iy += rh

    # Arrow 1
    text(s, Inches(bx + bw + 0.04), Inches(box_y + box_h/2 - 0.20), Inches(gap - 0.08), Inches(0.40),
         "→", size=22, bold=True, color=GREY_LABEL, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Box 2 · PHASE 2 ADDS (what we do)
    bx2 = bx + bw + gap
    round_rect(s, Inches(bx2), Inches(box_y), Inches(bw), Inches(box_h), WHITE, radius=0.04)
    rect(s, Inches(bx2), Inches(box_y), Inches(bw), Inches(0.55), HEADER_BLUE)
    text(s, Inches(bx2 + 0.15), Inches(box_y + 0.06), Inches(bw - 0.30), Inches(0.22),
         "WHAT WE ADD  ·  Phase 2 work", size=11, bold=True, color=WHITE, font=BODY_FONT)
    text(s, Inches(bx2 + 0.15), Inches(box_y + 0.30), Inches(bw - 0.30), Inches(0.22),
         "6 capabilities · 4 gaps closed · CS and ITBA in parallel",
         size=8, italic=True, color=BLUE_PALE, font=BODY_FONT)
    add_items = [
        ("Refinement",          "S → Story · M → Epic · L → Initiative (Epics if > 3 months)"),
        ("Pre-estimation",      "T-shirt confirmed · Effort 1-5 by Architect · Pain Score finalized"),
        ("Architecture",        "NFRs · vendor model · cross-VS feasibility · Architect signs off"),
        ("Capacity reserve",    "% per CoS lane · per squad · monthly @ Gate 2"),
        ("Roadmap entry",       "rolling wave 3·6·12 mo · committed vs hypothesized"),
        ("DoR check",           "6 criteria · single gate to Execution · loop back if fail"),
    ]
    iy = box_y + 0.65
    rh = (box_h - 0.75) / len(add_items)
    for lbl, body in add_items:
        oval(s, Inches(bx2 + 0.20), Inches(iy + 0.06), Inches(0.16), Inches(0.16), HEADER_BLUE)
        text(s, Inches(bx2 + 0.45), Inches(iy + 0.02), Inches(bw - 0.55), Inches(0.20),
             lbl, size=8.5, bold=True, color=DARK_TEXT, font=BODY_FONT)
        text(s, Inches(bx2 + 0.45), Inches(iy + 0.20), Inches(bw - 0.55), Inches(rh - 0.22),
             body, size=8, color=GREY_TEXT, font=BODY_FONT)
        iy += rh

    # Arrow 2
    text(s, Inches(bx2 + bw + 0.04), Inches(box_y + box_h/2 - 0.20), Inches(gap - 0.08), Inches(0.40),
         "→", size=22, bold=True, color=GREY_LABEL, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Box 3 · READY FOR EXECUTION (output)
    bx3 = bx2 + bw + gap
    round_rect(s, Inches(bx3), Inches(box_y), Inches(bw), Inches(box_h), WHITE, radius=0.04)
    rect(s, Inches(bx3), Inches(box_y), Inches(bw), Inches(0.55), NAVY_DARK)
    text(s, Inches(bx3 + 0.15), Inches(box_y + 0.06), Inches(bw - 0.30), Inches(0.22),
         "WHAT FLOWS OUT  ·  Phase 3 input", size=11, bold=True, color=WHITE, font=BODY_FONT)
    text(s, Inches(bx3 + 0.15), Inches(box_y + 0.30), Inches(bw - 0.30), Inches(0.22),
         "DoR-passed · Squad pulls · Execution begins",
         size=8, italic=True, color=BLUE_PALE, font=BODY_FONT)
    out_items = [
        ("Stories (≤ 5 SP)",      "INVEST-compliant · acceptance criteria binary"),
        ("NFRs documented",       "perf · security · compliance · scale · attached to Story"),
        ("Capacity reserved",     "fits inside CoS lane % · WIP not exceeded"),
        ("Acceptance criteria",   "traceable to Sponsor-approved Initiative or Epic"),
        ("Dependencies known",    "cross-VS · vendor · license — surfaced before pull"),
        ("Roadmap entry",         "committed band · re-balanced monthly"),
    ]
    iy = box_y + 0.65
    rh = (box_h - 0.75) / len(out_items)
    for lbl, body in out_items:
        oval(s, Inches(bx3 + 0.20), Inches(iy + 0.06), Inches(0.16), Inches(0.16), NAVY_DARK)
        text(s, Inches(bx3 + 0.45), Inches(iy + 0.02), Inches(bw - 0.55), Inches(0.20),
             lbl, size=8.5, bold=True, color=DARK_TEXT, font=BODY_FONT)
        text(s, Inches(bx3 + 0.45), Inches(iy + 0.20), Inches(bw - 0.55), Inches(rh - 0.22),
             body, size=8, color=GREY_TEXT, font=BODY_FONT)
        iy += rh

    footer_bar(s, 1)
    return s


# 02 · Governing thought · 4 gaps · preview of 5 categories
def build_pp02(prs):
    s = _slide(prs)
    slide_title(s,
        "Preparation & Planning closes 4 gaps",
        "Detail · Sequence · Capacity · Readiness — the four jobs of this phase",
        section="Opener")

    # Banner statement
    bn_y = 1.55
    round_rect(s, Inches(0.40), Inches(bn_y), Inches(12.55), Inches(0.65), GREEN_BG, radius=0.04)
    rect(s, Inches(0.40), Inches(bn_y), Inches(0.10), Inches(0.65), AVANGRID_GREEN)
    text(s, Inches(0.55), Inches(bn_y + 0.06), Inches(12.40), Inches(0.22),
         "GOVERNING THOUGHT", size=9, bold=True, color=GREEN_DARK, font=BODY_FONT)
    text(s, Inches(0.55), Inches(bn_y + 0.26), Inches(12.40), Inches(0.35),
         "A prioritized list says \"we want this\" — Preparation & Planning answers \"how, by whom, with what capacity, and when ready\".",
         size=12, bold=True, color=DARK_TEXT, font=BODY_FONT)

    # 4 gap columns
    gap_y = 2.10
    gap_h = 2.55
    gw = (12.55 - 3 * 0.20) / 4
    gaps = [
        ("01", "DETAIL",
         "A Demand becomes Initiative · Epic · Story (by size) — outcome-named, sized, acceptance-criteria split.",
         "Capability · Refinement + Estimation",
         AVANGRID_GREEN),
        ("02", "SEQUENCE",
         "Pain Score + Class of Service + capacity available decide what gets pulled when.",
         "Capability · Prioritization + Roadmap",
         GREEN_MID),
        ("03", "CAPACITY",
         "% per CoS lane decided monthly · middle managers hold day-to-day reality.",
         "Capability · Capacity + Forecasting",
         HEADER_BLUE),
        ("04", "READINESS",
         "Definition of Ready — 6 criteria · single gate to Execution.",
         "Capability · DoR check",
         HEADER_BLUE_DK),
    ]
    for i, (num, name, claim, cap, color) in enumerate(gaps):
        x = 0.40 + i * (gw + 0.20)
        round_rect(s, Inches(x), Inches(gap_y), Inches(gw), Inches(gap_h), WHITE, radius=0.04)
        rect(s, Inches(x), Inches(gap_y), Inches(gw), Inches(0.55), color)
        text(s, Inches(x + 0.15), Inches(gap_y + 0.05), Inches(0.50), Inches(0.30),
             num, size=20, bold=True, color=WHITE, font=BODY_FONT)
        text(s, Inches(x + 0.70), Inches(gap_y + 0.10), Inches(gw - 0.85), Inches(0.30),
             name, size=14, bold=True, color=WHITE, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
        text(s, Inches(x + 0.20), Inches(gap_y + 0.70), Inches(gw - 0.40), Inches(1.25),
             claim, size=10, color=DARK_TEXT, font=BODY_FONT)
        # capability tag at bottom
        round_rect(s, Inches(x + 0.20), Inches(gap_y + gap_h - 0.55), Inches(gw - 0.40), Inches(0.40),
                   GREEN_PALE, radius=0.04)
        text(s, Inches(x + 0.30), Inches(gap_y + gap_h - 0.50), Inches(gw - 0.55), Inches(0.30),
             cap, size=8, italic=True, color=GREEN_DARK, font=BODY_FONT,
             anchor=MSO_ANCHOR.MIDDLE)

    # 5 categories preview footer (the framework structure)
    cat_y = 4.95
    text(s, Inches(0.40), Inches(cat_y), Inches(12.55), Inches(0.22),
         "ORGANIZED AS 5 CATEGORIES  ·  the framework that structures the rest of this section",
         size=10, bold=True, color=GREY_LABEL, font=BODY_FONT)
    cats = [
        ("CAPABILITIES", "8 slides · what we do",      AVANGRID_GREEN),
        ("ROLES",        "3 slides · who does it",     HEADER_BLUE),
        ("TOOLS",        "1 slide · what supports us", GREEN_MID),
        ("GOVERNANCE",   "4 slides · decision rights", NAVY_DARK),
        ("PERFORMANCE",  "3 slides · how we measure",  COS_ORANGE),
    ]
    cw = (12.55 - 4 * 0.15) / 5
    for i, (name, sub, color) in enumerate(cats):
        x = 0.40 + i * (cw + 0.15)
        round_rect(s, Inches(x), Inches(cat_y + 0.30), Inches(cw), Inches(1.00), WHITE, radius=0.04)
        rect(s, Inches(x), Inches(cat_y + 0.30), Inches(0.10), Inches(1.00), color)
        text(s, Inches(x + 0.20), Inches(cat_y + 0.40), Inches(cw - 0.30), Inches(0.30),
             name, size=11, bold=True, color=color, font=BODY_FONT)
        text(s, Inches(x + 0.20), Inches(cat_y + 0.72), Inches(cw - 0.30), Inches(0.50),
             sub, size=8.5, color=GREY_TEXT, font=BODY_FONT)

    # Insight callout
    insight_callout(s, 0.40, 6.45, 12.55, 0.55,
        "Phase 2 has clear boundaries — we do not run BRPs (Business Review Process), Big Room Planning, Discovery, or plan XL.",
        "Couple to existing P&T BRP + Digital CoE BRP · NO Big Room Planning of our own · vendor or domain runs Discovery (Phase 0/1) · XL initiatives flow through PPM Now")

    footer_bar(s, 2)
    return s


# 03 · Two boards · the visual key · Portfolio ↔ Team
def build_pp03(prs):
    s = _slide(prs)
    slide_title(s,
        "Two boards · one bridge",
        "Portfolio decides the WHAT · Team decides the WHEN · DoR is the bridge between them",
        section="Opener")

    # LEFT · Portfolio Board mini-mockup (60% width)
    pb_x = 0.40; pb_y = 1.30; pb_w = 7.50; pb_h = 5.00
    round_rect(s, Inches(pb_x), Inches(pb_y), Inches(pb_w), Inches(pb_h), WHITE, radius=0.04)
    rect(s, Inches(pb_x), Inches(pb_y), Inches(pb_w), Inches(0.42), NAVY_DARK)
    text(s, Inches(pb_x + 0.15), Inches(pb_y + 0.06), Inches(pb_w - 0.30), Inches(0.30),
         "PORTFOLIO BOARD  ·  cross-VS  ·  Director-level  ·  L only",
         size=11, bold=True, color=WHITE, font=BODY_FONT)
    # Gate markers on top
    text(s, Inches(pb_x + 1.50), Inches(pb_y + 0.50), Inches(2.0), Inches(0.20),
         "▲ GATE 1 · weekly", size=8, bold=True, color=GATE_RED, font=BODY_FONT, align=PP_ALIGN.CENTER)
    text(s, Inches(pb_x + 4.30), Inches(pb_y + 0.50), Inches(2.0), Inches(0.20),
         "▲ GATE 2 · monthly", size=8, bold=True, color=GATE_RED, font=BODY_FONT, align=PP_ALIGN.CENTER)
    # 6 columns mini
    cols = [
        ("Funnel",       GREEN_PALE),
        ("Reviewing",    GREEN_BG),
        ("Analyzing",    GREEN_BG),
        ("Backlog",      GREEN_PASTEL_OK := GREEN_LIGHT),
        ("Implementing", BLUE_BG),
        ("Done",         GREY_LIGHT),
    ]
    col_y = pb_y + 0.80
    col_h = 3.50
    cw_pb = (pb_w - 0.30 - 5 * 0.04) / 6
    for i, (name, bg) in enumerate(cols):
        cx = pb_x + 0.15 + i * (cw_pb + 0.04)
        # column header
        rect(s, Inches(cx), Inches(col_y), Inches(cw_pb), Inches(0.32), AVANGRID_GREEN)
        text(s, Inches(cx), Inches(col_y), Inches(cw_pb), Inches(0.32), name,
             size=8.5, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # body bg
        rect(s, Inches(cx), Inches(col_y + 0.32), Inches(cw_pb), Inches(col_h - 0.32), bg)
    # Vertical gates between cols (Gate 1 between 0-1, Gate 2 between 2-3)
    g1_x = pb_x + 0.15 + 1 * (cw_pb + 0.04) - 0.02
    g2_x = pb_x + 0.15 + 3 * (cw_pb + 0.04) - 0.02
    rect(s, Inches(g1_x), Inches(col_y), Inches(0.04), Inches(col_h), GATE_RED)
    rect(s, Inches(g2_x), Inches(col_y), Inches(0.04), Inches(col_h), GATE_RED)
    # Sample card in Backlog column (the one ready to bridge)
    cx_bl = pb_x + 0.15 + 3 * (cw_pb + 0.04)
    crd_y = col_y + 0.45
    round_rect(s, Inches(cx_bl + 0.08), Inches(crd_y), Inches(cw_pb - 0.16), Inches(0.85),
               WHITE, radius=0.04)
    text(s, Inches(cx_bl + 0.12), Inches(crd_y + 0.05), Inches(cw_pb - 0.24), Inches(0.18),
         "L · Initiative", size=7, bold=True, color=GREY_LABEL, font=BODY_FONT)
    text(s, Inches(cx_bl + 0.12), Inches(crd_y + 0.22), Inches(cw_pb - 0.24), Inches(0.40),
         "PUC tariff banner Q3", size=7.5, bold=True, color=DARK_TEXT, font=BODY_FONT)
    chip(s, cx_bl + 0.12, crd_y + 0.62, 0.70, "Fixed-Date", COS_BLUE, WHITE, h=0.18, size=6)
    # CS+ITBA marker on column (which area)
    chip(s, cx_bl + cw_pb - 0.55, crd_y + 0.05, 0.40, "ITBA", HEADER_BLUE, WHITE, h=0.16, size=6)

    # Phase 2 ownership band on Portfolio Board
    rect(s, Inches(pb_x + 0.15 + 1 * (cw_pb + 0.04)), Inches(col_y + col_h + 0.10),
         Inches(3 * (cw_pb + 0.04) - 0.04), Inches(0.28), HEADER_BLUE)
    text(s, Inches(pb_x + 0.15 + 1 * (cw_pb + 0.04)), Inches(col_y + col_h + 0.10),
         Inches(3 * (cw_pb + 0.04) - 0.04), Inches(0.28),
         "▲ PHASE 2 owns these 3 columns · Reviewing · Analyzing · Backlog (RFE)",
         size=8, bold=True, color=WHITE, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Bypass note on PB
    text(s, Inches(pb_x + 0.15), Inches(col_y + col_h + 0.45), Inches(pb_w - 0.30), Inches(0.20),
         "Only S demands BYPASS this board (go directly to Team Kanban) · M / L flow through Gate 2 · XL → PPM Now",
         size=8, italic=True, color=BYPASS_ORANGE, font=BODY_FONT)

    # MIDDLE · DoR vertical bridge (widened to fit 6 criteria labels)
    bridge_x = pb_x + pb_w + 0.20
    bridge_w = 1.85
    round_rect(s, Inches(bridge_x), Inches(pb_y + 0.80), Inches(bridge_w), Inches(col_h),
               GREEN_DARK, radius=0.04)
    text(s, Inches(bridge_x), Inches(pb_y + 0.95), Inches(bridge_w), Inches(0.30),
         "DoR", size=14, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
    text(s, Inches(bridge_x), Inches(pb_y + 1.25), Inches(bridge_w), Inches(0.20),
         "GATE  ·  6 criteria", size=8, bold=True, color=GREEN_PASTEL_OK, font=BODY_FONT, align=PP_ALIGN.CENTER)
    # 6 checkmarks + labels
    dor_criteria = [
        "Outcome clear",
        "Sized",
        "Architecture reviewed",
        "Dependencies known",
        "Capacity reserved",
        "Acceptance criteria",
    ]
    cm_y = pb_y + 1.55
    for i, lbl in enumerate(dor_criteria):
        oval(s, Inches(bridge_x + 0.10), Inches(cm_y + i * 0.42),
             Inches(0.20), Inches(0.20), WHITE)
        text(s, Inches(bridge_x + 0.10), Inches(cm_y + i * 0.42),
             Inches(0.20), Inches(0.20), "✓",
             size=10, bold=True, color=GREEN_DARK, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, Inches(bridge_x + 0.35), Inches(cm_y + i * 0.42), Inches(bridge_w - 0.40), Inches(0.20),
             lbl, size=8, bold=True, color=WHITE, font=BODY_FONT,
             anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(bridge_x - 0.05), Inches(cm_y + 6 * 0.42 + 0.05),
         Inches(bridge_w + 0.10), Inches(0.30),
         "Definition of Ready",
         size=8, italic=True, color=GREEN_DARK, font=BODY_FONT,
         align=PP_ALIGN.CENTER)

    # RIGHT · Team Kanban mini-mockup (30%)
    tk_x = bridge_x + bridge_w + 0.20
    tk_w = 13.33 - tk_x - 0.40
    round_rect(s, Inches(tk_x), Inches(pb_y), Inches(tk_w), Inches(pb_h), WHITE, radius=0.04)
    rect(s, Inches(tk_x), Inches(pb_y), Inches(tk_w), Inches(0.42), HEADER_BLUE_DK)
    text(s, Inches(tk_x + 0.15), Inches(pb_y + 0.06), Inches(tk_w - 0.30), Inches(0.30),
         "TEAM KANBAN  ·  Squad-level  ·  S direct · M / L after Gate 2",
         size=11, bold=True, color=WHITE, font=BODY_FONT)
    text(s, Inches(tk_x + 0.15), Inches(pb_y + 0.50), Inches(tk_w - 0.30), Inches(0.20),
         "Example · CS Billing Operations Squad",
         size=8, italic=True, color=GREY_TEXT, font=BODY_FONT)
    # 4 columns
    tcols = [("Backlog", GREEN_PALE), ("In Progress", BLUE_BG), ("Review", GREEN_BG), ("Done", GREY_LIGHT)]
    tw = (tk_w - 0.30 - 3 * 0.04) / 4
    for i, (name, bg) in enumerate(tcols):
        cx = tk_x + 0.15 + i * (tw + 0.04)
        rect(s, Inches(cx), Inches(col_y), Inches(tw), Inches(0.32), HEADER_BLUE)
        text(s, Inches(cx), Inches(col_y), Inches(tw), Inches(0.32), name,
             size=7.5, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        rect(s, Inches(cx), Inches(col_y + 0.32), Inches(tw), Inches(col_h - 0.32), bg)
    # Sample S story arriving in Backlog
    cx_b = tk_x + 0.15
    crd_y2 = col_y + 0.45
    round_rect(s, Inches(cx_b + 0.06), Inches(crd_y2), Inches(tw - 0.12), Inches(0.85),
               WHITE, radius=0.04)
    text(s, Inches(cx_b + 0.10), Inches(crd_y2 + 0.04), Inches(tw - 0.20), Inches(0.18),
         "S · Story", size=6.5, bold=True, color=GREY_LABEL, font=BODY_FONT)
    text(s, Inches(cx_b + 0.10), Inches(crd_y2 + 0.20), Inches(tw - 0.20), Inches(0.42),
         "CRM page load <2s", size=7, bold=True, color=DARK_TEXT, font=BODY_FONT)
    chip(s, cx_b + 0.10, crd_y2 + 0.62, 0.70, "Expedite", COS_RED, WHITE, h=0.18, size=6)
    chip(s, cx_b + tw - 0.45, crd_y2 + 0.04, 0.30, "CS", AVANGRID_GREEN, WHITE, h=0.16, size=6)

    # Direction arrow
    text(s, Inches(bridge_x - 0.05), Inches(pb_y + col_h + 0.30), Inches(bridge_w + 0.10), Inches(0.30),
         "→ pull", size=10, bold=True, color=GREEN_DARK, font=BODY_FONT, align=PP_ALIGN.CENTER)

    # Bottom · how it works
    bot_y = 6.40
    insight_callout(s, 0.40, bot_y, 12.55, 0.55,
        "DoR (Definition of Ready) = 6-criteria checklist · the only gate Stories cross before Execution — Stories are pulled, not pushed.",
        "PM (per Squad) checks DoR + WIP, then pulls when capacity allows · same flow for CS and ITBA · DoR detail covered later in the deck")

    cos_legend(s)
    footer_bar(s, 3)
    return s


GREEN_PASTEL = RGBColor(0xC8, 0xE0, 0xCB)  # forward-compat


# ============================================================================
# SECTION 2 · CAPABILITIES  (slides 04-11)
# ============================================================================

# 04 · Capabilities map · 6 capabilities clustered into 3 muscles
def build_pp04(prs):
    s = _slide(prs)
    slide_title(s,
        "Six capabilities · three muscles",
        "Detail · Sequence · Capacity — every capability lives inside one of these three groups",
        section="Capabilities")

    from_phase1_strip(s)

    # 3 muscle groups · big bands
    m_y = 1.65
    m_h = 5.50
    mw = (12.55 - 2 * 0.20) / 3

    muscles = [
        ("DETAIL  ·  detail muscle",
         "Turn vague demand into shippable stories",
         AVANGRID_GREEN, GREEN_PALE,
         [("Refinement",
           "S → Story · M → Epic · L → Initiative · Epics if > 3 months",
           "outcome-named, not deliverable-named",
           "PM (per Squad)",
           "slide 05"),
          ("Estimation",
           "T-shirt + Effort 1-5 + Pain Score",
           "CS uses 2 components · ITBA uses 3 (Safety isolated)",
           "PM scores BV/UX · Architect scores Effort/Safety",
           "slide 06"),
          ("Grouping",
           "Initiative · Epic · Story · Sub-task",
           "hierarchy mirrors the work breakdown",
           "PM (per Squad)",
           "embedded in 05")]),

        ("SEQUENCE  ·  sequence muscle",
         "Decide what gets built when — across CS and ITBA",
         GREEN_MID, GREEN_BG,
         [("Prioritization",
           "2 overrides + Pain Score + CoS",
           "Type=Regulatory → Fixed-Date · severity event → Expedite",
           "PM (per Squad) proposes within-Squad rank · Director arbitrates cross-VS at Gate 2 · VP breaks ties",
           "slide 07"),
          ("Portfolio backlog",
           "Per-area ranked (CS · ITBA independent) · monthly",
           "Pain Score + CoS + capacity-available",
           "PM (per Squad) feeds rank · Director approves portfolio-level at Gate 2",
           "slide 08"),
          ("Roadmap (rolling wave)",
           "3 mo committed · 3-6 planned · 6-12 hypothesized",
           "commits and hypotheses NEVER mixed visually",
           "PM (per Squad) · re-balanced cross-VS at Gate 2 by Director trio",
           "slide 09")]),

        ("CAPACITY  ·  capacity muscle",
         "Match demand to available squad capacity — and react when reality bites",
         HEADER_BLUE, BLUE_BG,
         [("Capacity planning",
           "% per CoS lane · per squad · per quarter",
           "decided MONTHLY at Gate 2 · day-to-day with middle managers",
           "Director (approves) + Middle managers (day-to-day)",
           "slide 10"),
          ("Forecasting + alignment",
           "Squad capacity (FTE × velocity) FIXED quarterly",
           "demand vs available capacity visible at Gate 2",
           "Middle managers maintain capacity reality · PM (per Squad) consumes within their lane",
           "embedded in 10"),
          ("Work allocation + ad-hoc",
           "Auto-Expedite (P1/P2) · scope-add · emergency",
           "3 channels for unplanned demand",
           "PM (per Squad) accepts within 48h",
           "slide 11")]),
    ]
    for i, (title_band, sub, color_dk, color_bg, items) in enumerate(muscles):
        x = 0.40 + i * (mw + 0.20)
        round_rect(s, Inches(x), Inches(m_y), Inches(mw), Inches(m_h), WHITE, radius=0.04)
        rect(s, Inches(x), Inches(m_y), Inches(mw), Inches(0.70), color_dk)
        text(s, Inches(x + 0.20), Inches(m_y + 0.08), Inches(mw - 0.40), Inches(0.30),
             title_band, size=12, bold=True, color=WHITE, font=BODY_FONT)
        text(s, Inches(x + 0.20), Inches(m_y + 0.40), Inches(mw - 0.40), Inches(0.25),
             sub, size=8, italic=True, color=WHITE, font=BODY_FONT)
        # Capability cards
        cy = m_y + 0.85
        ch = (m_h - 1.00) / len(items)
        for cap_name, what, how, who, ref in items:
            round_rect(s, Inches(x + 0.15), Inches(cy), Inches(mw - 0.30), Inches(ch - 0.10),
                       color_bg, radius=0.04)
            rect(s, Inches(x + 0.15), Inches(cy), Inches(0.06), Inches(ch - 0.10), color_dk)
            text(s, Inches(x + 0.30), Inches(cy + 0.06), Inches(mw - 1.0), Inches(0.22),
                 cap_name, size=10, bold=True, color=color_dk, font=BODY_FONT)
            text(s, Inches(x + mw - 1.10), Inches(cy + 0.06), Inches(0.95), Inches(0.20),
                 ref, size=7, italic=True, color=GREY_LABEL, font=BODY_FONT, align=PP_ALIGN.RIGHT)
            text(s, Inches(x + 0.30), Inches(cy + 0.30), Inches(mw - 0.50), Inches(0.22),
                 "▪ " + what, size=8, color=DARK_TEXT, font=BODY_FONT)
            text(s, Inches(x + 0.30), Inches(cy + 0.52), Inches(mw - 0.50), Inches(0.22),
                 "▪ " + how, size=8, color=GREY_TEXT, font=BODY_FONT)
            text(s, Inches(x + 0.30), Inches(cy + 0.74), Inches(mw - 0.50), Inches(0.22),
                 "OWNER · " + who, size=7, italic=True, bold=True, color=color_dk, font=BODY_FONT)
            cy += ch

    footer_bar(s, 4)
    return s


# 05 · Refinement · Demand becomes Initiative + ~3 Stories · outcome-named
def build_pp05(prs):
    s = _slide(prs)
    slide_title(s,
        "Refinement · outcome-named hierarchy",
        "S → Story · M → Epic · L → Initiative · Initiative > 3 months breaks into Epics",
        section="Capabilities")

    from_phase1_strip(s)

    # Hierarchy tree (left to right · 5 levels)
    tree_y = 1.55
    levels_x = [0.40, 2.40, 4.50, 7.30, 10.50]
    levels_w = [1.85, 1.95, 2.55, 2.75, 2.40]
    levels_def = [
        ("DEMAND",     "intake placeholder",   AVANGRID_GREEN, "from Phase 1"),
        ("INITIATIVE", "L size · default Initiative → Story", HEADER_BLUE_DK, "PM (per VS)"),
        ("EPIC",       "M size · or Initiative > 3 months",     HEADER_BLUE,    "PM (per VS)"),
        ("STORY",      "S size direct · or child of Initiative/Epic", GREEN_MID, "Squad PM owns DoR"),
        ("SUB-TASK",   "tech breakdown",        GREY_LABEL,     "Dev within Story"),
    ]
    base_y = tree_y + 0.95
    box_h = 0.62

    # L0..L2 single boxes centered
    for i in range(3):
        x = levels_x[i]; w = levels_w[i]
        round_rect(s, Inches(x), Inches(base_y), Inches(w), Inches(box_h), levels_def[i][2], radius=0.06)
        text(s, Inches(x), Inches(base_y + 0.04), Inches(w), Inches(0.22),
             levels_def[i][0], size=11, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, Inches(x), Inches(base_y + 0.27), Inches(w), Inches(0.18),
             levels_def[i][1], size=8, italic=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, Inches(x), Inches(base_y + 0.45), Inches(w), Inches(0.18),
             levels_def[i][3], size=7, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)

    # Connectors L0→L1→L2
    for i in range(2):
        line(s, Inches(levels_x[i] + levels_w[i]),
             Inches(base_y + box_h/2),
             Inches(levels_x[i+1]),
             Inches(base_y + box_h/2),
             levels_def[i][2], weight=2.0)

    # L3 · 3 Stories fan-out
    story_ys = [tree_y + 0.20, base_y, tree_y + 1.70]
    sw = levels_w[3]; sx = levels_x[3]
    for j, sy in enumerate(story_ys):
        round_rect(s, Inches(sx), Inches(sy), Inches(sw), Inches(box_h), GREEN_MID, radius=0.06)
        text(s, Inches(sx), Inches(sy + 0.04), Inches(sw), Inches(0.22),
             f"STORY {j+1}", size=11, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, Inches(sx), Inches(sy + 0.27), Inches(sw), Inches(0.18),
             "≤ 5 SP · INVEST", size=8, italic=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, Inches(sx), Inches(sy + 0.45), Inches(sw), Inches(0.18),
             "Squad PM owns DoR", size=7, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
        line(s, Inches(levels_x[2] + levels_w[2]), Inches(base_y + box_h/2),
             Inches(sx), Inches(sy + box_h/2), GREEN_MID, weight=1.5)

    # L4 · 2 sub-tasks for first story (illustrative)
    sub_w = levels_w[4]; sub_x = levels_x[4]; sub_h = 0.36
    for j in range(2):
        sy = story_ys[0] + j * 0.42
        round_rect(s, Inches(sub_x), Inches(sy), Inches(sub_w), Inches(sub_h), GREY_LABEL, radius=0.06)
        text(s, Inches(sub_x), Inches(sy), Inches(sub_w), Inches(sub_h),
             f"SUB-TASK {j+1}", size=8, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        line(s, Inches(sx + sw), Inches(story_ys[0] + box_h/2),
             Inches(sub_x), Inches(sy + sub_h/2), GREY_LABEL, weight=1.0)
    text(s, Inches(sub_x), Inches(story_ys[1] + 0.10), Inches(sub_w), Inches(0.42),
         "(sub-tasks per Story\nas needed)", size=7, italic=True, color=GREY_LABEL,
         font=BODY_FONT, align=PP_ALIGN.CENTER)
    text(s, Inches(sub_x), Inches(story_ys[2] + 0.10), Inches(sub_w), Inches(0.42),
         "(sub-tasks per Story\nas needed)", size=7, italic=True, color=GREY_LABEL,
         font=BODY_FONT, align=PP_ALIGN.CENTER)

    # INVEST band
    inv_y = 4.30
    round_rect(s, Inches(0.40), Inches(inv_y), Inches(12.55), Inches(0.95), GREEN_BG, radius=0.04)
    rect(s, Inches(0.40), Inches(inv_y), Inches(0.10), Inches(0.95), AVANGRID_GREEN)
    text(s, Inches(0.55), Inches(inv_y + 0.08), Inches(12.40), Inches(0.20),
         "INVEST  ·  what every Story must satisfy at the Story level (Squad PM verifies)",
         size=10, bold=True, color=GREEN_DARK, font=BODY_FONT)
    invest = [("I", "Independent"), ("N", "Negotiable"), ("V", "Valuable"),
              ("E", "Estimable"), ("S", "Small (≤5 SP)"), ("T", "Testable")]
    cw = (12.55 - 0.30 - 5 * 0.10) / 6
    for i, (letter, lbl) in enumerate(invest):
        x = 0.55 + i * (cw + 0.10)
        oval(s, Inches(x + cw/2 - 0.18), Inches(inv_y + 0.32), Inches(0.36), Inches(0.36), AVANGRID_GREEN)
        text(s, Inches(x + cw/2 - 0.18), Inches(inv_y + 0.32), Inches(0.36), Inches(0.36),
             letter, size=12, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, Inches(x), Inches(inv_y + 0.72), Inches(cw), Inches(0.20),
             lbl, size=8, bold=True, color=DARK_TEXT, font=BODY_FONT, align=PP_ALIGN.CENTER)

    # CS vs ITBA framing comparison
    bot_y = 5.45
    sw2 = (12.55 - 0.20) / 2
    bot_h = 1.45
    # CS
    round_rect(s, Inches(0.40), Inches(bot_y), Inches(sw2), Inches(bot_h), WHITE, radius=0.04)
    rect(s, Inches(0.40), Inches(bot_y), Inches(sw2), Inches(0.32), AVANGRID_GREEN)
    text(s, Inches(0.55), Inches(bot_y + 0.04), Inches(sw2 - 0.30), Inches(0.24),
         "CS  ·  outcomes named in customer language",
         size=10, bold=True, color=WHITE, font=BODY_FONT)
    text(s, Inches(0.55), Inches(bot_y + 0.40), Inches(sw2 - 0.30), Inches(0.40),
         "✓ ACCEPT: \"reduce time-to-resolve by 30%\"  ·  \"NPS +5 points\"  ·  \"deflect 15% of inbound calls\"",
         size=9, color=DARK_TEXT, font=BODY_FONT)
    text(s, Inches(0.55), Inches(bot_y + 0.78), Inches(sw2 - 0.30), Inches(0.40),
         "✕ REJECT: \"add CSV export feature\"  ·  \"build new dashboard\"  ·  \"upgrade Genesys to v9\"",
         size=9, italic=True, color=GREY_LABEL, font=BODY_FONT)

    # ITBA
    cx2 = 0.40 + sw2 + 0.20
    round_rect(s, Inches(cx2), Inches(bot_y), Inches(sw2), Inches(bot_h), WHITE, radius=0.04)
    rect(s, Inches(cx2), Inches(bot_y), Inches(sw2), Inches(0.32), HEADER_BLUE)
    text(s, Inches(cx2 + 0.15), Inches(bot_y + 0.04), Inches(sw2 - 0.30), Inches(0.24),
         "ITBA  ·  outcomes named in business language",
         size=10, bold=True, color=WHITE, font=BODY_FONT)
    text(s, Inches(cx2 + 0.15), Inches(bot_y + 0.40), Inches(sw2 - 0.30), Inches(0.40),
         "✓ ACCEPT: \"close month-end 2 days faster\"  ·  \"audit pass without findings\"  ·  \"$2M cost avoided\"",
         size=9, color=DARK_TEXT, font=BODY_FONT)
    text(s, Inches(cx2 + 0.15), Inches(bot_y + 0.78), Inches(sw2 - 0.30), Inches(0.40),
         "✕ REJECT: \"upgrade SAP module\"  ·  \"build new portal\"  ·  \"add audit field\"",
         size=9, italic=True, color=GREY_LABEL, font=BODY_FONT)

    footer_bar(s, 5)
    return s


# 06 · Estimation · CS+ITBA Pain Score formulas
def build_pp06(prs):
    s = _slide(prs)
    slide_title(s,
        "Pain Score · effort first, then formula",
        "STEP 1 Architect estimates Effort  ·  STEP 2 BV/UX/Safety scored  ·  STEP 3 Pain Score auto-computes",
        section="Capabilities")

    from_phase1_strip(s)

    # Two big formula panels
    panel_y = 1.30
    panel_h = 5.55
    sw2 = (12.55 - 0.20) / 2

    def render_formula(x, color, color_dk, color_bg, area, formula_str, weights_def, who_def, examples):
        round_rect(s, Inches(x), Inches(panel_y), Inches(sw2), Inches(panel_h), WHITE, radius=0.04)
        rect(s, Inches(x), Inches(panel_y), Inches(sw2), Inches(0.55), color)
        text(s, Inches(x + 0.15), Inches(panel_y + 0.06), Inches(sw2 - 0.30), Inches(0.22),
             f"{area}  ·  Pain Score formula", size=12, bold=True, color=WHITE, font=BODY_FONT)
        text(s, Inches(x + 0.15), Inches(panel_y + 0.30), Inches(sw2 - 0.30), Inches(0.20),
             f"finalized at Analyzing · auto-computed (Automation A4)",
             size=8, italic=True, color=BLUE_PALE if area=="ITBA" else GREEN_PALE, font=BODY_FONT)

        # Big formula box
        f_y = panel_y + 0.70
        round_rect(s, Inches(x + 0.20), Inches(f_y), Inches(sw2 - 0.40), Inches(1.00), color_bg, radius=0.04)
        rect(s, Inches(x + 0.20), Inches(f_y), Inches(0.10), Inches(1.00), color)
        text(s, Inches(x + 0.40), Inches(f_y + 0.10), Inches(sw2 - 0.60), Inches(0.30),
             "FORMULA", size=8, bold=True, color=color_dk, font=BODY_FONT)
        text(s, Inches(x + 0.40), Inches(f_y + 0.32), Inches(sw2 - 0.60), Inches(0.55),
             formula_str, size=14, bold=True, color=DARK_TEXT, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Weights panel
        w_y = f_y + 1.15
        round_rect(s, Inches(x + 0.20), Inches(w_y), Inches(sw2 - 0.40), Inches(0.85), GREY_LIGHT, radius=0.04)
        text(s, Inches(x + 0.30), Inches(w_y + 0.06), Inches(sw2 - 0.60), Inches(0.20),
             "WEIGHTS  ·  utility-calibrated", size=8, bold=True, color=GREY_LABEL, font=BODY_FONT)
        wx = x + 0.30; wy = w_y + 0.30
        for label, val, c in weights_def:
            chip(s, wx, wy, 1.10, f"{label} {val}", c, WHITE, h=0.32, size=9)
            wx += 1.20

        # Owner / who scores what
        o_y = w_y + 1.00
        round_rect(s, Inches(x + 0.20), Inches(o_y), Inches(sw2 - 0.40), Inches(1.10), color_bg, radius=0.04)
        text(s, Inches(x + 0.30), Inches(o_y + 0.08), Inches(sw2 - 0.60), Inches(0.20),
             "WHO SCORES WHAT  ·  1-5 scale defined by Architect",
             size=8, bold=True, color=color_dk, font=BODY_FONT)
        oy = o_y + 0.30
        for owner_label, scoring in who_def:
            text(s, Inches(x + 0.30), Inches(oy), Inches(2.20), Inches(0.20),
                 owner_label, size=8, bold=True, color=color_dk, font=BODY_FONT)
            text(s, Inches(x + 2.55), Inches(oy), Inches(sw2 - 2.85), Inches(0.20),
                 scoring, size=8, color=DARK_TEXT, font=BODY_FONT)
            oy += 0.22

        # Worked example
        ex_y = o_y + 1.25
        round_rect(s, Inches(x + 0.20), Inches(ex_y), Inches(sw2 - 0.40), Inches(0.95), WHITE, radius=0.04)
        rect(s, Inches(x + 0.20), Inches(ex_y), Inches(0.10), Inches(0.95), INSIGHT_AMBER)
        text(s, Inches(x + 0.35), Inches(ex_y + 0.06), Inches(sw2 - 0.55), Inches(0.20),
             "WORKED EXAMPLE", size=8, bold=True, color=INSIGHT_AMBER, font=BODY_FONT)
        text(s, Inches(x + 0.35), Inches(ex_y + 0.26), Inches(sw2 - 0.55), Inches(0.30),
             examples[0], size=9, bold=True, color=DARK_TEXT, font=BODY_FONT)
        text(s, Inches(x + 0.35), Inches(ex_y + 0.55), Inches(sw2 - 0.55), Inches(0.40),
             examples[1], size=10, bold=True, color=color, font=BODY_FONT)

    render_formula(0.40, AVANGRID_GREEN, GREEN_DARK, GREEN_PALE, "CS",
        "(BV · OPCOmult · 0.40 + UX · 0.60) / Effort",
        [("BV", "40%", AVANGRID_GREEN), ("UX", "60%", GREEN_MID)],
        [("PM (per VS)",   "scores BV + UX (1-5)"),
         ("Architect",     "scores Effort (1-5) at Analyzing"),
         ("OPCOmult auto", "1.0 · 1.3 · 1.6 (1 / 2-4 / 5+ OPCOs)")],
        ["BV=4 · UX=5 · OPCOmult=1.3 · Effort=3",
         "Pain Score = (4 · 1.3 · 0.40 + 5 · 0.60) / 3 = 2.43"])

    render_formula(0.40 + sw2 + 0.20, HEADER_BLUE, HEADER_BLUE_DK, BLUE_PALE, "ITBA",
        "(BV · OPCOmult · w₁ + UX · w₂ + Safety · w₃) / Effort",
        [("BV w₁", "0.45", HEADER_BLUE), ("UX w₂", "0.30", BLUE_LIGHT), ("Safety w₃", "0.25", HEADER_BLUE_DK)],
        [("PM (per VS)",   "scores BV + UX (1-5)"),
         ("Architect",     "scores Effort + Safety (1-5) at Analyzing"),
         ("Safety isolated","field-crew impact · NOT buried in BV"),
         ("OPCOmult auto",  "1.0 · 1.3 · 1.6 (1 / 2-4 / 5+ OPCOs)")],
        ["BV=5 · UX=3 · Safety=4 · OPCOmult=1.6 · Effort=4",
         "Pain Score = (5 · 1.6 · 0.45 + 3 · 0.30 + 4 · 0.25) / 4 = 1.38"])

    footer_bar(s, 6)
    return s


# 07 · Prioritization · layered overrides + Pain Score
def build_pp07(prs):
    s = _slide(prs)
    slide_title(s,
        "Prioritization is layered",
        "Two overrides protect critical work · Pain Score sequences the rest · CoS lane caps it",
        section="Capabilities")

    from_phase1_strip(s)

    # Layered diagram — 4 horizontal bands stacked
    layer_x = 0.40
    layer_w = 12.55
    layer_y_start = 1.70
    layer_h = 0.95
    gap = 0.06

    layers = [
        ("OVERRIDE 1  ·  TYPE = REGULATORY",
         "Triggered by: Type field set to Regulatory (at Reviewing)",
         "Action: A3 sets CoS = Fixed-Date · deadline pinned · enters Fixed-Date lane",
         "Why first: regulatory deadlines are absolute · cannot be missed",
         GATE_RED, GATE_RED_BG),
        ("OVERRIDE 2  ·  SEVERITY · EMERGENCY",
         "Triggered by: Director declares severity event (production outage · safety · critical impact)",
         "Action: Manual CoS = Expedite · WIP-1 lane · jumps the queue regardless of Pain Score",
         "Why second: severity events take precedence · audit logged",
         COS_RED, RGBColor(0xFB, 0xEC, 0xEC)),
        ("PAIN SCORE  ·  numeric ranking",
         "Computed: A4 auto-computes when BV/UX/Effort/Safety populated ",
         "CS = (BV·OPCOmult·0.40 + UX·0.60) / Effort  ·  ITBA = (BV·OPCOmult·0.45 + UX·0.30 + Safety·0.25) / Effort",
         "Higher score = higher pain = earlier sequence",
         AVANGRID_GREEN, GREEN_BG),
        ("CAPACITY FLOOR  ·  CoS lane availability",
         "Even with high Pain Score, item must fit inside CoS lane capacity %",
         "If lane at WIP cap → next-rank item picks up · waiting items re-evaluated next cycle",
         "Decided monthly at Gate 2 by Director + middle managers",
         HEADER_BLUE, BLUE_BG),
    ]
    for i, (lbl, trigger, action, why, color, bg) in enumerate(layers):
        ly = layer_y_start + i * (layer_h + gap)
        round_rect(s, Inches(layer_x), Inches(ly), Inches(layer_w), Inches(layer_h), bg, radius=0.04)
        rect(s, Inches(layer_x), Inches(ly), Inches(0.10), Inches(layer_h), color)
        # Layer number circle
        oval(s, Inches(layer_x + 0.20), Inches(ly + 0.18), Inches(0.55), Inches(0.55), color)
        text(s, Inches(layer_x + 0.20), Inches(ly + 0.18), Inches(0.55), Inches(0.55),
             str(i+1), size=18, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Label
        text(s, Inches(layer_x + 0.95), Inches(ly + 0.06), Inches(layer_w - 1.10), Inches(0.25),
             lbl, size=11, bold=True, color=color, font=BODY_FONT)
        # 3 rows
        text(s, Inches(layer_x + 0.95), Inches(ly + 0.30), Inches(layer_w - 1.10), Inches(0.20),
             trigger, size=8.5, color=DARK_TEXT, font=BODY_FONT)
        text(s, Inches(layer_x + 0.95), Inches(ly + 0.50), Inches(layer_w - 1.10), Inches(0.20),
             action, size=8.5, color=DARK_TEXT, font=BODY_FONT)
        text(s, Inches(layer_x + 0.95), Inches(ly + 0.70), Inches(layer_w - 1.10), Inches(0.22),
             why, size=8, italic=True, color=GREY_TEXT, font=BODY_FONT)

    # Insight callout at bottom
    insight_callout(s, 0.40, 5.65, 12.55, 1.25,
        "Two items with the same Pain Score? The override layer + capacity floor decide — never just the number.",
        "Example · Pain 2.4 each: \"Regulatory tariff change (ITBA)\" wins over \"CRM perf optimization (CS)\" because Type=Regulatory triggers Override 1 → Fixed-Date lane (deadline-bound). The CS item competes inside Standard lane on Pain Score alone.")

    footer_bar(s, 7)
    return s


# 08 · Portfolio backlogs (independent per area · CS + ITBA)
def build_pp08(prs):
    s = _slide(prs)
    slide_title(s,
        "Backlog · independent per area",
        "CS and ITBA are independent value streams · each ranked separately · same Pain Score logic, different formula inputs",
        section="Capabilities")

    from_phase1_strip(s)

    # === LEFT · CS Backlog ===
    cs_x = 0.40; bl_y = 1.60; cs_w = 6.20; bl_h = 5.25
    round_rect(s, Inches(cs_x), Inches(bl_y), Inches(cs_w), Inches(bl_h), WHITE, radius=0.04)
    rect(s, Inches(cs_x), Inches(bl_y), Inches(cs_w), Inches(0.45), AVANGRID_GREEN)
    text(s, Inches(cs_x + 0.15), Inches(bl_y + 0.08), Inches(cs_w - 0.30), Inches(0.30),
         "CS BACKLOG  ·  Customer Services  ·  Q3 snapshot",
         size=11, bold=True, color=WHITE, font=BODY_FONT)

    # CS column headers
    hdr_y = bl_y + 0.55
    cs_cols = [("RANK", 0.50), ("DEMAND", 3.05), ("PAIN", 0.65), ("CoS", 1.00), ("SIZE", 0.50)]
    cx = cs_x + 0.10
    cs_col_x = {}
    for label, w in cs_cols:
        cs_col_x[label] = cx
        text(s, Inches(cx), Inches(hdr_y), Inches(w), Inches(0.20),
             label, size=7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        cx += w
    rect(s, Inches(cs_x + 0.10), Inches(hdr_y + 0.22), Inches(cs_w - 0.20), Emu(15000), GREY_BORDER)

    cs_items = [
        ("01", "CRM page load <2s on Genesys",            "2.10", "Expedite", COS_RED,    "S"),
        ("02", "Reduce billing-correction cycle 14→5 d",   "2.43", "Standard", COS_GREEN,  "M"),
        ("03", "Self-service password (multi-OPCO)",        "2.05", "Standard", COS_GREEN,  "S"),
        ("04", "GIS layer migration · scope-add",           "1.80", "Standard", COS_GREEN,  "S"),
        ("05", "Smart Meter analytics platform",            "1.55", "Standard", COS_GREEN,  "L"),
    ]
    iy = hdr_y + 0.35
    rh = (bl_h - 1.10) / len(cs_items)
    for i, (rank, demand, pain, cos, cos_color, size) in enumerate(cs_items):
        if i % 2 == 1:
            rect(s, Inches(cs_x + 0.10), Inches(iy), Inches(cs_w - 0.20), Inches(rh - 0.04), GREY_LIGHT)
        text(s, Inches(cs_col_x["RANK"]), Inches(iy + 0.06), Inches(0.50), Inches(rh - 0.10),
             rank, size=11, bold=True, color=GREEN_DARK, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
        text(s, Inches(cs_col_x["DEMAND"]), Inches(iy + 0.06), Inches(3.00), Inches(rh - 0.10),
             demand, size=8.5, color=DARK_TEXT, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
        text(s, Inches(cs_col_x["PAIN"]), Inches(iy + 0.06), Inches(0.60), Inches(rh - 0.10),
             pain, size=10, bold=True, color=NAVY_DARK, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
        chip(s, cs_col_x["CoS"], iy + (rh - 0.22) / 2, 0.95, cos, cos_color, WHITE, h=0.20, size=7)
        size_color = COS_GREEN if size == "S" else (COS_ORANGE if size == "M" else COS_BLUE)
        chip(s, cs_col_x["SIZE"], iy + (rh - 0.22) / 2, 0.45, size, size_color, WHITE, h=0.20, size=7)
        iy += rh

    # === RIGHT · ITBA Backlog ===
    itba_x = cs_x + cs_w + 0.20
    round_rect(s, Inches(itba_x), Inches(bl_y), Inches(cs_w), Inches(bl_h), WHITE, radius=0.04)
    rect(s, Inches(itba_x), Inches(bl_y), Inches(cs_w), Inches(0.45), HEADER_BLUE)
    text(s, Inches(itba_x + 0.15), Inches(bl_y + 0.08), Inches(cs_w - 0.30), Inches(0.30),
         "ITBA BACKLOG  ·  IT Business Apps  ·  Q3 snapshot",
         size=11, bold=True, color=WHITE, font=BODY_FONT)

    cx = itba_x + 0.10
    itba_col_x = {}
    for label, w in cs_cols:
        itba_col_x[label] = cx
        text(s, Inches(cx), Inches(hdr_y), Inches(w), Inches(0.20),
             label, size=7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        cx += w
    rect(s, Inches(itba_x + 0.10), Inches(hdr_y + 0.22), Inches(cs_w - 0.20), Emu(15000), GREY_BORDER)

    itba_items = [
        ("01", "PUC tariff banner Q3 (regulatory)",        "2.05", "Fixed-Date", COS_BLUE,   "L"),
        ("02", "SAP FI year-end close acceleration",       "2.20", "Standard",   COS_GREEN,  "M"),
        ("03", "Audit framework refresh (regulatory)",      "1.85", "Fixed-Date", COS_BLUE,   "L"),
        ("04", "Networks IT data backbone (cross-VS)",      "1.65", "Standard",   COS_GREEN,  "L"),
        ("05", "SAP Basis kernel patch Q2 (tech debt)",     "0.95", "Intangible", COS_ORANGE, "M"),
    ]
    iy = hdr_y + 0.35
    for i, (rank, demand, pain, cos, cos_color, size) in enumerate(itba_items):
        if i % 2 == 1:
            rect(s, Inches(itba_x + 0.10), Inches(iy), Inches(cs_w - 0.20), Inches(rh - 0.04), GREY_LIGHT)
        text(s, Inches(itba_col_x["RANK"]), Inches(iy + 0.06), Inches(0.50), Inches(rh - 0.10),
             rank, size=11, bold=True, color=HEADER_BLUE_DK, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
        text(s, Inches(itba_col_x["DEMAND"]), Inches(iy + 0.06), Inches(3.00), Inches(rh - 0.10),
             demand, size=8.5, color=DARK_TEXT, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
        text(s, Inches(itba_col_x["PAIN"]), Inches(iy + 0.06), Inches(0.60), Inches(rh - 0.10),
             pain, size=10, bold=True, color=NAVY_DARK, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
        chip(s, itba_col_x["CoS"], iy + (rh - 0.22) / 2, 0.95, cos, cos_color, WHITE, h=0.20, size=7)
        size_color = COS_GREEN if size == "S" else (COS_ORANGE if size == "M" else COS_BLUE)
        chip(s, itba_col_x["SIZE"], iy + (rh - 0.22) / 2, 0.45, size, size_color, WHITE, h=0.20, size=7)
        iy += rh

    # === Bottom · re-ranking notes (compact) ===
    nt_y = bl_y + bl_h + 0.10
    insight_callout(s, 0.40, nt_y, 12.55, 0.80,
        "Each area ranks independently — same Pain Score logic, different formula inputs · cross-area only when dependencies surface.",
        "PM (per VS) keeps within-VS rank fresh weekly · Director (CS: Kumar+Heather · ITBA: Carlos+Heather) arbitrates cross-VS at Gate 2 · cross-area dependencies escalate to VP Vitor")

    cos_legend(s)
    footer_bar(s, 8)
    return s


# 09 · Roadmap rolling wave + planning cycles
def build_pp09(prs):
    s = _slide(prs)
    slide_title(s,
        "Rolling-wave roadmap",
        "3 mo committed · 3-6 planned · 6-12 hypothesized · re-balanced monthly at Gate 2",
        section="Capabilities")

    from_phase1_strip(s)

    # Time axis
    ax_x = 1.80; ax_y = 1.70; ax_w = 11.15
    quarters = ["Q3 2026", "Q4 2026", "Q1 2027", "Q2 2027"]
    qw = ax_w / 4
    for i, q in enumerate(quarters):
        rect(s, Inches(ax_x + i * qw), Inches(ax_y), Inches(qw - 0.04), Inches(0.32),
             NAVY_DARK if i == 0 else (HEADER_BLUE if i == 1 else GREY_LABEL))
        text(s, Inches(ax_x + i * qw), Inches(ax_y), Inches(qw - 0.04), Inches(0.32),
             q, size=10, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Horizon labels
    band_y = ax_y + 0.36
    text(s, Inches(ax_x), Inches(band_y), Inches(qw), Inches(0.20),
         "COMMITTED",   size=8, bold=True, color=GREEN_DARK,  font=BODY_FONT, align=PP_ALIGN.CENTER)
    text(s, Inches(ax_x + qw), Inches(band_y), Inches(qw), Inches(0.20),
         "PLANNED",     size=8, bold=True, color=HEADER_BLUE, font=BODY_FONT, align=PP_ALIGN.CENTER)
    text(s, Inches(ax_x + 2*qw), Inches(band_y), Inches(2*qw), Inches(0.20),
         "HYPOTHESIZED", size=8, bold=True, color=GREY_LABEL, font=BODY_FONT, align=PP_ALIGN.CENTER)

    # Items by VS rows · CS+ITBA mixed
    rows_y = band_y + 0.30
    rows_def = [
        ("Customer Care",          "CS",   AVANGRID_GREEN, [
            # Q3 (0-0.25 = committed) → green-dark · Q4 (0.25-0.50 = planned) → header-blue
            ("Genesys script overhaul",  0.00, 0.40, GREEN_DARK,   COS_GREEN),
            ("CRM page load <2s",        0.05, 0.20, GREEN_DARK,   COS_RED),
        ]),
        ("Billing Operations",     "CS",   AVANGRID_GREEN, [
            ("Billing cycle 14→5 days",  0.05, 0.40, GREEN_DARK,   COS_GREEN),
            ("Self-service password",    0.10, 0.30, GREEN_DARK,   COS_GREEN),
        ]),
        ("Smart Metering",         "CS",   AVANGRID_GREEN, [
            # Starts Q4 → planned (header-blue)
            ("Smart Meter analytics",    0.30, 0.75, HEADER_BLUE,  COS_GREEN),
        ]),
        ("Corporate & Regulatory", "ITBA", HEADER_BLUE,    [
            # PUC Q3 only · committed
            ("PUC tariff banner Q3",     0.02, 0.22, GREEN_DARK,   COS_BLUE),
            # Audit refresh Q4-Q1 · planned start
            ("Audit framework refresh",  0.30, 0.65, HEADER_BLUE,  COS_BLUE),
        ]),
        ("Networks IT",            "ITBA", HEADER_BLUE,    [
            ("Networks data backbone",   0.30, 0.55, HEADER_BLUE,  COS_GREEN),
            # S/4HANA · hypothesized start (Q1+)
            ("S/4HANA migration plan",   0.55, 0.95, GREY_LABEL,   COS_GREEN),
        ]),
        ("Iberdrola Corp Services","ITBA", HEADER_BLUE,    [
            # SAP Basis Q3 → committed
            ("SAP Basis patch Q2",       0.05, 0.20, GREEN_DARK,   COS_ORANGE),
        ]),
    ]
    rows_h_avail = 3.55
    rh = rows_h_avail / len(rows_def)
    for ri, (vs_name, area, vs_color, items) in enumerate(rows_def):
        ry = rows_y + ri * rh
        bg = WHITE if ri % 2 == 0 else GREY_LIGHT
        rect(s, Inches(0.40), Inches(ry), Inches(1.35), Inches(rh - 0.02), bg)
        rect(s, Inches(0.40), Inches(ry), Inches(0.10), Inches(rh - 0.02), vs_color)
        text(s, Inches(0.55), Inches(ry + 0.05), Inches(1.20), Inches(0.20),
             vs_name, size=8.5, bold=True, color=DARK_TEXT, font=BODY_FONT)
        text(s, Inches(0.55), Inches(ry + 0.24), Inches(1.20), Inches(0.18),
             area, size=7, italic=True, color=GREY_LABEL, font=BODY_FONT)
        # Plot bg
        rect(s, Inches(ax_x), Inches(ry), Inches(ax_w), Inches(rh - 0.02), bg)
        # Quarter dividers
        for q in range(1, 4):
            rect(s, Inches(ax_x + q * qw), Inches(ry), Emu(7000), Inches(rh - 0.02), GREY_BORDER)
        # Bars
        for j, (lbl, sx, ex, color, cos_c) in enumerate(items):
            bx = ax_x + sx * ax_w
            bw = (ex - sx) * ax_w
            by = ry + 0.06 + j * (rh - 0.20) / max(1, len(items))
            bh = (rh - 0.20) / max(1, len(items)) - 0.06
            round_rect(s, Inches(bx), Inches(by), Inches(bw), Inches(bh), color, radius=0.04)
            oval(s, Inches(bx + 0.04), Inches(by + bh/2 - 0.06), Inches(0.12), Inches(0.12), cos_c)
            text(s, Inches(bx + 0.22), Inches(by + 0.04), Inches(bw - 0.30), Inches(bh - 0.08),
                 lbl, size=7.5, bold=True, color=WHITE, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)

    # Cadence band at bottom
    cad_y = 6.05
    round_rect(s, Inches(0.40), Inches(cad_y), Inches(12.55), Inches(0.85), GREEN_BG, radius=0.04)
    rect(s, Inches(0.40), Inches(cad_y), Inches(0.10), Inches(0.85), AVANGRID_GREEN)
    text(s, Inches(0.55), Inches(cad_y + 0.06), Inches(12.40), Inches(0.20),
         "PLANNING CYCLES  ·  monthly Gate 2 + ad-hoc + BRP coupling",
         size=10, bold=True, color=GREEN_DARK, font=BODY_FONT)
    cads = [
        ("Monthly · Gate 2",   "90 min · VP Vitor chairs · re-balance roadmap · approve capacity % · CS + ITBA reviewed independently"),
        ("BRP coupling",       "Quarterly priorities digest fed to existing P&T BRP + Digital CoE BRP · we do NOT run our own"),
        ("Ad-hoc · re-rank",   "P1/P2 lands · regulatory deadline shifts · capacity changes"),
    ]
    cw = (12.55 - 0.30 - 2 * 0.10) / 3
    for i, (lbl, body) in enumerate(cads):
        x = 0.55 + i * (cw + 0.10)
        text(s, Inches(x), Inches(cad_y + 0.32), Inches(cw), Inches(0.18),
             lbl, size=8.5, bold=True, color=GREEN_DARK, font=BODY_FONT)
        text(s, Inches(x), Inches(cad_y + 0.50), Inches(cw), Inches(0.30),
             body, size=8, color=DARK_TEXT, font=BODY_FONT)

    cos_legend(s)
    footer_bar(s, 9)
    return s


# 10 · Capacity allocation matrix · 9 squads × 4 CoS lanes
def build_pp10(prs):
    s = _slide(prs)
    slide_title(s,
        "Capacity · squad first, lane second",
        "PRIMARY · squad FTE × velocity (fixed quarterly)  ·  SECONDARY · % per CoS lane (Gate 2 monthly)  ·  best practice from Lean Kanban",
        section="Capabilities")

    from_phase1_strip(s)

    # Header note about middle managers
    nm_y = 1.55
    round_rect(s, Inches(0.40), Inches(nm_y), Inches(12.55), Inches(0.42), BLUE_BG, radius=0.04)
    rect(s, Inches(0.40), Inches(nm_y), Inches(0.10), Inches(0.42), HEADER_BLUE)
    text(s, Inches(0.55), Inches(nm_y + 0.06), Inches(12.40), Inches(0.20),
         "CAPACITY HIERARCHY  ·  squad FTE primary · CoS allocation secondary",
         size=8.5, bold=True, color=HEADER_BLUE, font=BODY_FONT)
    text(s, Inches(0.55), Inches(nm_y + 0.22), Inches(12.40), Inches(0.20),
         "FIXED · squad capacity = FTE × velocity (quarterly)  ·  ALLOCATED · % per CoS lane (Gate 2 monthly · Director + middle managers)  ·  CONSUMED · PM (per VS) pulls within lane",
         size=8, color=DARK_TEXT, font=BODY_FONT)

    # Matrix
    mtx_y = 2.20
    label_w = 2.40
    total_w = 1.10
    cos_count = 4
    cell_w = (12.55 - label_w - total_w) / cos_count

    # Header row
    rect(s, Inches(0.40), Inches(mtx_y), Inches(label_w), Inches(0.42), NAVY_DARK)
    text(s, Inches(0.40), Inches(mtx_y), Inches(label_w), Inches(0.42),
         "SQUAD  ·  per Value Stream", size=9, bold=True, color=WHITE, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    cos_lanes = [("STANDARD", COS_GREEN, "WIP 3-4"),
                 ("FIXED-DATE", COS_BLUE, "WIP 2"),
                 ("EXPEDITE", COS_RED, "WIP 1"),
                 ("INTANGIBLE", COS_ORANGE, "WIP 1")]
    for ci, (lbl, color, wip) in enumerate(cos_lanes):
        cx = 0.40 + label_w + ci * cell_w
        rect(s, Inches(cx), Inches(mtx_y), Inches(cell_w - 0.04), Inches(0.42), color)
        text(s, Inches(cx), Inches(mtx_y + 0.02), Inches(cell_w - 0.04), Inches(0.20),
             lbl, size=9, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, Inches(cx), Inches(mtx_y + 0.22), Inches(cell_w - 0.04), Inches(0.18),
             wip, size=7, italic=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
    rect(s, Inches(0.40 + label_w + cos_count * cell_w), Inches(mtx_y), Inches(total_w), Inches(0.42), NAVY_DARK)
    text(s, Inches(0.40 + label_w + cos_count * cell_w), Inches(mtx_y), Inches(total_w), Inches(0.42),
         "FTE", size=9, bold=True, color=WHITE, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Squad rows
    squads = [
        ("Customer Care",            "CS",   AVANGRID_GREEN, [55, 15, 20, 10], "8"),
        ("Billing Operations",       "CS",   AVANGRID_GREEN, [60, 15, 15, 10], "9"),
        ("Field Operations (CS)",    "CS",   AVANGRID_GREEN, [50, 25, 15, 10], "7"),
        ("Smart Metering",           "CS",   AVANGRID_GREEN, [55, 25, 10, 10], "6"),
        ("Corporate & Regulatory",   "ITBA", HEADER_BLUE,    [40, 45,  5, 10], "10"),
        ("Iberdrola Corp Services",  "ITBA", HEADER_BLUE,    [55, 20, 10, 15], "8"),
        ("Renewables IT",            "ITBA", HEADER_BLUE,    [60, 15, 15, 10], "7"),
        ("Networks IT",              "ITBA", HEADER_BLUE,    [50, 30, 10, 10], "9"),
        ("Enabling Team",            "BOTH", GREY_LABEL,     [40, 20, 20, 20], "5"),
    ]
    rh = 0.42
    ry = mtx_y + 0.42
    for ri, (name, area, color, pcts, fte) in enumerate(squads):
        bg = WHITE if ri % 2 == 0 else GREY_LIGHT
        rect(s, Inches(0.40), Inches(ry), Inches(label_w), Inches(rh - 0.02), bg)
        rect(s, Inches(0.40), Inches(ry), Inches(0.10), Inches(rh - 0.02), color)
        text(s, Inches(0.55), Inches(ry + 0.04), Inches(label_w - 0.20), Inches(0.18),
             name, size=8.5, bold=True, color=DARK_TEXT, font=BODY_FONT)
        text(s, Inches(0.55), Inches(ry + 0.22), Inches(label_w - 0.20), Inches(0.16),
             area, size=7, italic=True, color=GREY_LABEL, font=BODY_FONT)
        # CoS cells
        for ci, (lbl, color_c, wip) in enumerate(cos_lanes):
            cx = 0.40 + label_w + ci * cell_w
            rect(s, Inches(cx), Inches(ry), Inches(cell_w - 0.04), Inches(rh - 0.02), bg)
            pct = pcts[ci]
            bar_x = cx + 0.10
            bar_y = ry + rh / 2 - 0.10
            bar_w_max = cell_w - 0.30
            rect(s, Inches(bar_x), Inches(bar_y), Inches(bar_w_max), Inches(0.16), GREY_LIGHT)
            rect(s, Inches(bar_x), Inches(bar_y), Inches(bar_w_max * pct / 100), Inches(0.16), color_c)
            text(s, Inches(cx), Inches(ry + rh - 0.20), Inches(cell_w - 0.04), Inches(0.16),
                 f"{pct}%", size=8, bold=True, color=NAVY_DARK, font=BODY_FONT, align=PP_ALIGN.CENTER)
        # Total
        cx = 0.40 + label_w + cos_count * cell_w
        rect(s, Inches(cx), Inches(ry), Inches(total_w), Inches(rh - 0.02), bg)
        text(s, Inches(cx), Inches(ry), Inches(total_w), Inches(rh - 0.02),
             fte, size=10, bold=True, color=NAVY_DARK, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        ry += rh

    # Insight callout at bottom
    insight_callout(s, 0.40, 5.95, 12.55, 0.95,
        "Capacity reality lives with middle managers — not with the Director who approves the % at Gate 2.",
        "Director sets policy (% per CoS lane) once a month · Middle managers (Kumar's + Carlos's lieutenants) flex day-to-day within that envelope · PM (per VS) consumes the capacity for refinement and pull · misalignment caught at next Gate 2 with audit trail")

    cos_legend(s)
    footer_bar(s, 10)
    return s


# 11 · Work allocation + unplanned/ad-hoc · 3 channels
def build_pp11(prs):
    s = _slide(prs)
    slide_title(s,
        "Unplanned demand · 2 channels",
        "Scope-add shortcut · severity / emergency lane · Director declares severity at any time",
        section="Capabilities")

    from_phase1_strip(s)

    # 2 channel columns (CIT removed · merged into severity-driven channel)
    ch_y = 1.30
    ch_h = 5.55
    cw = (12.55 - 0.20) / 2

    channels = [
        ("CHANNEL 1  ·  SCOPE-ADD",
         "Demand-on-existing-Initiative shortcut",
         AVANGRID_GREEN, GREEN_BG,
         [("TRIGGER",     "New ask scopes-into an Initiative already in Implementing"),
          ("OWNER",       "PM (per Squad) evaluates within 48h"),
          ("DECISION",    "ACCEPT (preferred when parent Implementing)\nREJECT (when parent earlier — submit standalone)"),
          ("PATH",        "Story added under existing Initiative or Epic · DoR fast-track"),
          ("CoS IMPACT",  "Inherits parent CoS · stays in same lane"),
          ("EXAMPLE",     "CRM project live · add CSV export to reconciliation report"),
          ("BYPASS",      "No new Gate 2 · audit logged at parent Initiative")]),

        ("CHANNEL 2  ·  SEVERITY / EMERGENCY",
         "Production outage · regulatory deadline shift · safety · critical impact",
         INSIGHT_AMBER, INSIGHT_BG,
         [("TRIGGER",     "Production outage OR external regulatory shift OR safety incident OR critical impact declared by Director"),
          ("OWNER",       "Director (relevant area) + Architect on-call · severity declared at any time (no intake flag required)"),
          ("PATH",        "Direct to Squad · skip Funnel · ServiceNow-side handles pure incidents"),
          ("LANE",        "Expedite (severity / outage) OR Fixed-Date (regulatory deadline)"),
          ("CoS IMPACT",  "Pre-empts existing WIP if needed · Director call · WIP-1 lane"),
          ("EXAMPLES",    "Genesys queue routing failure during peak hours · PUC issues new regulatory directive with 30-day deadline"),
          ("AUDIT",       "Director signs off · post-mortem at next Gate 2 · all events logged")]),
    ]
    for i, (title_h, sub, color, bg, blocks) in enumerate(channels):
        x = 0.40 + i * (cw + 0.20)
        round_rect(s, Inches(x), Inches(ch_y), Inches(cw), Inches(ch_h), WHITE, radius=0.04)
        rect(s, Inches(x), Inches(ch_y), Inches(cw), Inches(0.65), color)
        text(s, Inches(x + 0.15), Inches(ch_y + 0.06), Inches(cw - 0.30), Inches(0.25),
             title_h, size=11, bold=True, color=WHITE, font=BODY_FONT)
        text(s, Inches(x + 0.15), Inches(ch_y + 0.32), Inches(cw - 0.30), Inches(0.30),
             sub, size=8, italic=True, color=WHITE, font=BODY_FONT)

        iy = ch_y + 0.80
        bh = (ch_h - 0.95) / len(blocks)
        for lbl, body in blocks:
            text(s, Inches(x + 0.15), Inches(iy), Inches(cw - 0.30), Inches(0.18),
                 lbl, size=7, bold=True, color=color, font=BODY_FONT)
            text(s, Inches(x + 0.15), Inches(iy + 0.18), Inches(cw - 0.30), Inches(bh - 0.22),
                 body, size=8, color=DARK_TEXT, font=BODY_FONT)
            iy += bh

    footer_bar(s, 11)
    return s


# ============================================================================
# SECTION 3 · ROLES  (slides 12-14)
# ============================================================================

# 12 · Three roles do 80% of work · Heather sets Product standard
def build_pp12(prs):
    s = _slide(prs)
    slide_title(s,
        "PM per Squad · Architect · Squad team",
        "ONE PM per Squad covers refinement → DoR → pull · portfolio-level decisions stay with Director trio at Gate 2",
        section="Roles")

    from_phase1_strip(s)

    # Top band · Heather as the bridge across CS and ITBA
    hb_y = 1.55
    round_rect(s, Inches(2.50), Inches(hb_y), Inches(8.30), Inches(0.65), GREEN_BG, radius=0.04)
    rect(s, Inches(2.50), Inches(hb_y), Inches(0.10), Inches(0.65), AVANGRID_GREEN)
    avatar(s, 2.70, hb_y + 0.05, 0.55, GREEN_DARK, "HK")
    text(s, Inches(3.40), Inches(hb_y + 0.06), Inches(7.20), Inches(0.22),
         "HEATHER LaJOIE  ·  Director · Product (NEW capability for Avangrid)",
         size=12, bold=True, color=GREEN_DARK, font=BODY_FONT)
    text(s, Inches(3.40), Inches(hb_y + 0.30), Inches(7.20), Inches(0.30),
         "Spans CS and ITBA · leads + mentors new Product capability · liaison to Digital CoE BRP · co-Director at Gate 2",
         size=8.5, italic=True, color=DARK_TEXT, font=BODY_FONT)

    # 3 role columns
    col_y = 2.10
    col_h = 4.80
    cw = (12.55 - 2 * 0.20) / 3

    roles = [
        ("PM (per Squad)",
         "PMs · per VS · NEW capability across CS + ITBA",
         AVANGRID_GREEN, GREEN_PALE,
         [("OWNS",
           "Refinement · Portfolio backlog rank within VS · Roadmap entries"),
          ("PRODUCES",
           "Outcome-named Stories · INVEST-compliant · acceptance criteria split"),
          ("MEETS",
           "Weekly refinement (PM + Architect + Squad)  ·  Bi-weekly grooming  ·  Monthly Gate 2 prep"),
          ("ACCOUNTABLE FOR",
           "Outcome framing · NOT delivery (Squad PM owns) · NOT capacity (Director owns)"),
          ("ONBOARDING + RAMP",
           "12-week onboarding · mentored by Heather · paired with BRS counterpart"),
          ("WHO IS THIS PERSON",
           "Senior individual contributor · product-mindset · domain literacy")]),

        ("ARCHITECT  ·  cross-VS",
         "1 architect per area (CS + ITBA) · cross-VS feasibility",
         HEADER_BLUE, BLUE_PALE,
         [("OWNS",
           "Effort scoring (1-5) · Safety scoring (ITBA · 1-5) · NFR pre-check"),
          ("PRODUCES",
           "Architecture position note per Initiative · cross-VS dependency map"),
          ("MEETS",
           "Weekly refinement (alongside PM)  ·  Architecture review on demand  ·  Monthly Gate 2"),
          ("ACCOUNTABLE FOR",
           "NFRs (perf · security · compliance · scale) · vendor model · cross-VS deps"),
          ("DEFINES SCALE 1-5",
           "Effort: 1=hours · 2=days · 3=weeks · 4=quarter · 5=multi-quarter"),
          ("WHO IS THIS PERSON",
           "Senior architect · 8+ years · vendor + in-house experience")]),

        ("SQUAD PM",
         "1 per squad · 9 squads (4 CS + 5 ITBA) + 1 Enabling",
         GREEN_DARK, GREEN_BG,
         [("OWNS",
           "DoR check · pull from CoS lane · WIP discipline · daily ops"),
          ("PRODUCES",
           "DoR-passed pull queue · iteration-level commitments · cycle-time data"),
          ("MEETS",
           "Daily standup  ·  Weekly refinement (with PM + Architect)  ·  Bi-weekly retro"),
          ("ACCOUNTABLE FOR",
           "Single gate to Execution · pull discipline · NOT prioritization (PM owns)"),
          ("ESCALATES TO",
           "PM (per VS) on outcome drift · Architect on NFR breach · Director on capacity"),
          ("WHO IS THIS PERSON",
           "Senior individual contributor · product-mindset · Scrum/Kanban background")]),
    ]
    for i, (name, sub, color, bg, blocks) in enumerate(roles):
        x = 0.40 + i * (cw + 0.20)
        round_rect(s, Inches(x), Inches(col_y), Inches(cw), Inches(col_h), WHITE, radius=0.04)
        rect(s, Inches(x), Inches(col_y), Inches(cw), Inches(0.65), color)
        text(s, Inches(x + 0.15), Inches(col_y + 0.06), Inches(cw - 0.30), Inches(0.25),
             name, size=11, bold=True, color=WHITE, font=BODY_FONT)
        text(s, Inches(x + 0.15), Inches(col_y + 0.32), Inches(cw - 0.30), Inches(0.30),
             sub, size=8, italic=True, color=WHITE, font=BODY_FONT)

        iy = col_y + 0.80
        bh = (col_h - 0.95) / len(blocks)
        for lbl, body in blocks:
            text(s, Inches(x + 0.15), Inches(iy), Inches(cw - 0.30), Inches(0.18),
                 lbl, size=7, bold=True, color=color, font=BODY_FONT)
            text(s, Inches(x + 0.15), Inches(iy + 0.18), Inches(cw - 0.30), Inches(bh - 0.22),
                 body, size=8, color=DARK_TEXT, font=BODY_FONT)
            iy += bh

    footer_bar(s, 12)
    return s


# 13 · PM deep-dive · NEW capability · onboarding ramp · BRS counterpart
def build_pp13(prs):
    s = _slide(prs)
    slide_title(s,
        "PM onboarding · NEW capability · 12-week ramp",
        "Pair with BRS counterpart · Heather mentors · standalone by Day 90",
        section="Roles")

    from_phase1_strip(s)

    # Timeline · 4 phases of PM onboarding
    tl_y = 1.65
    tl_h = 2.10
    rect(s, Inches(0.40), Inches(tl_y), Inches(12.55), Inches(0.42), AVANGRID_GREEN)
    text(s, Inches(0.55), Inches(tl_y + 0.08), Inches(12.40), Inches(0.26),
         "PM ONBOARDING  ·  12 weeks · paired with BRS counterpart",
         size=11, bold=True, color=WHITE, font=BODY_FONT)

    phases = [
        ("WEEKS 1-3", "ORIENT",
         "Shadow Heather · meet Sponsors · learn taxonomy · observe Gate 2",
         AVANGRID_GREEN),
        ("WEEKS 4-6", "PAIR",
         "Co-refine with senior PM · score 5 demands with Architect · run first DoR",
         GREEN_MID),
        ("WEEKS 7-9", "OWN",
         "Own VS backlog · run weekly refinement · prepare Gate 2 deck",
         HEADER_BLUE),
        ("WEEKS 10-12", "EVALUATE",
         "Heather reviews outcome quality · BRS feedback · pass = standalone",
         HEADER_BLUE_DK),
    ]
    pw = (12.55 - 0.30 - 3 * 0.08) / 4
    for i, (week, name, body, color) in enumerate(phases):
        x = 0.55 + i * (pw + 0.08)
        round_rect(s, Inches(x), Inches(tl_y + 0.55), Inches(pw), Inches(tl_h - 0.65),
                   WHITE, radius=0.04)
        rect(s, Inches(x), Inches(tl_y + 0.55), Inches(0.10), Inches(tl_h - 0.65), color)
        text(s, Inches(x + 0.20), Inches(tl_y + 0.62), Inches(pw - 0.30), Inches(0.20),
             week, size=8, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, Inches(x + 0.20), Inches(tl_y + 0.82), Inches(pw - 0.30), Inches(0.30),
             name, size=14, bold=True, color=color, font=BODY_FONT)
        text(s, Inches(x + 0.20), Inches(tl_y + 1.15), Inches(pw - 0.30), Inches(tl_h - 1.25),
             body, size=8.5, color=DARK_TEXT, font=BODY_FONT)

    # PM × BRS pairing
    pair_y = 3.65
    pair_h = 1.85
    sw2 = (12.55 - 0.20) / 2

    # PM box
    round_rect(s, Inches(0.40), Inches(pair_y), Inches(sw2), Inches(pair_h), WHITE, radius=0.04)
    rect(s, Inches(0.40), Inches(pair_y), Inches(sw2), Inches(0.42), AVANGRID_GREEN)
    text(s, Inches(0.55), Inches(pair_y + 0.06), Inches(sw2 - 0.30), Inches(0.30),
         "PM (per Squad)  ·  the NEW Product role",
         size=11, bold=True, color=WHITE, font=BODY_FONT)
    pm_items = [
        ("Owns",     "Outcome framing · refinement · Pain Score (BV/UX) · roadmap · DoR · pull · daily ops"),
        ("Cadence",  "Daily within VS  ·  weekly refinement  ·  monthly Gate 2"),
        ("Reports",  "To Heather (Director Product)"),
        ("Skill",    "Product mindset · outcome-driven · domain literacy"),
    ]
    iy = pair_y + 0.55
    rh = (pair_h - 0.65) / len(pm_items)
    for lbl, body in pm_items:
        text(s, Inches(0.55), Inches(iy), Inches(1.20), Inches(0.18),
             lbl.upper(), size=7, bold=True, color=GREEN_DARK, font=BODY_FONT)
        text(s, Inches(1.85), Inches(iy + 0.02), Inches(sw2 - 2.10), Inches(rh - 0.10),
             body, size=8.5, color=DARK_TEXT, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
        iy += rh

    # Connecting arrow ↔
    arr_x = 0.40 + sw2 + 0.04
    text(s, Inches(arr_x), Inches(pair_y + pair_h/2 - 0.20), Inches(0.20), Inches(0.40),
         "↔", size=22, bold=True, color=GREY_LABEL, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # BRS box
    cx2 = 0.40 + sw2 + 0.20
    round_rect(s, Inches(cx2), Inches(pair_y), Inches(sw2), Inches(pair_h), WHITE, radius=0.04)
    rect(s, Inches(cx2), Inches(pair_y), Inches(sw2), Inches(0.42), HEADER_BLUE)
    text(s, Inches(cx2 + 0.15), Inches(pair_y + 0.06), Inches(sw2 - 0.30), Inches(0.30),
         "BRS  ·  Business Relationship Steward",
         size=11, bold=True, color=WHITE, font=BODY_FONT)
    brs_items = [
        ("Owns",     "Process owner counterpart · value stream advocacy · outcome challenge"),
        ("Cadence",  "Weekly with PM  ·  monthly at Gate 2  ·  quarterly with VP Vitor"),
        ("Reports",  "To business unit Director (CS or ITBA side)"),
        ("Skill",    "Business process mastery · stakeholder mgmt · governance fluency"),
    ]
    iy = pair_y + 0.55
    for lbl, body in brs_items:
        text(s, Inches(cx2 + 0.15), Inches(iy), Inches(1.20), Inches(0.18),
             lbl.upper(), size=7, bold=True, color=HEADER_BLUE, font=BODY_FONT)
        text(s, Inches(cx2 + 1.45), Inches(iy + 0.02), Inches(sw2 - 1.70), Inches(rh - 0.10),
             body, size=8.5, color=DARK_TEXT, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
        iy += rh

    # Insight callout
    insight_callout(s, 0.40, 5.75, 12.55, 1.10,
        "PM and BRS are partners — PM owns the WHAT (product outcomes), BRS owns the WHY (business value).",
        "Together they ensure each Initiative ties an outcome to a business case · neither role alone is sufficient · both attend Gate 2 · both sign off on roadmap commits")

    footer_bar(s, 13)
    return s


# 14 · RACI matrix · 8 activities × 8 roles
def build_pp14(prs):
    s = _slide(prs)
    slide_title(s,
        "RACI · who does what",
        "8 activities × 8 roles · PM responsible for refinement · Squad PM accountable for DoR",
        section="Roles")

    from_phase1_strip(s)

    activities = [
        "Refinement (Demand → Initiative · Epic · Story)",
        "Pre-estimation (BV · UX · Effort · Safety)",
        "Pain Score computation",
        "Architecture review (NFRs · vendor · deps)",
        "Capacity allocation (% per CoS lane)",
        "Roadmap rolling-wave maintenance",
        "Defer / Reject decision",
        "DoR check + handoff to Squad",
    ]
    roles = [
        ("VP",          "Vitor",     HEADER_BLUE_DK),
        ("DIR · PROD",  "Heather",   AVANGRID_GREEN),
        ("DIR · CS",    "Kumar",     AVANGRID_GREEN),
        ("DIR · ITBA",  "Carlos",    HEADER_BLUE),
        ("MIDDLE MGR",  "day-to-day",GREY_LABEL),
        ("PM (Squad)",  "NEW",       GREEN_MID),
        ("ARCHITECT",   "cross-VS",  INSIGHT_AMBER),
        ("SQUAD TEAM",  "TL+BA+SME", GREEN_DARK),
    ]

    raci = [
        # Refinement
        ["I", "C", "I", "I", "I", "A/R", "C", "C"],
        # Pre-estimation
        ["I", "I", "I", "I", "I", "A/R", "R", "I"],
        # Pain Score (auto · A4)
        ["I", "I", "I", "I", "I", "C",   "I", "I"],
        # Architecture review
        ["I", "C", "I", "I", "I", "C",   "A/R","I"],
        # Capacity allocation
        ["I", "C", "A", "A", "R", "I",   "I", "I"],
        # Roadmap maintenance
        ["I", "A", "C", "C", "I", "R",   "C", "I"],
        # Defer / Reject
        ["I", "C", "A/R", "A/R","I", "C", "C", "I"],
        # DoR check + handoff
        ["I", "I", "I", "I", "I", "A/R", "C", "C"],
    ]

    tbl_y = 1.65
    tbl_h = 5.55
    label_w = 3.60
    cell_w = (12.55 - label_w) / len(roles)

    # Header row
    rect(s, Inches(0.40), Inches(tbl_y), Inches(label_w), Inches(0.65), NAVY_DARK)
    text(s, Inches(0.40), Inches(tbl_y + 0.10), Inches(label_w), Inches(0.45),
         "ACTIVITY", size=11, bold=True, color=WHITE, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    for ci, (rname, rsub, rcolor) in enumerate(roles):
        x = 0.40 + label_w + ci * cell_w
        rect(s, Inches(x), Inches(tbl_y), Inches(cell_w - 0.02), Inches(0.65), rcolor)
        text(s, Inches(x), Inches(tbl_y + 0.06), Inches(cell_w - 0.02), Inches(0.22),
             rname, size=8.5, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, Inches(x), Inches(tbl_y + 0.30), Inches(cell_w - 0.02), Inches(0.20),
             rsub, size=7, italic=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)

    # Activity rows
    rh = (tbl_h - 0.65) / len(activities)
    ry = tbl_y + 0.65
    color_map = {"R": AVANGRID_GREEN, "A": GATE_RED, "A/R": GATE_RED,
                 "C": HEADER_BLUE, "I": GREY_MUTED}
    for ri, act in enumerate(activities):
        bg = WHITE if ri % 2 == 0 else GREY_LIGHT
        rect(s, Inches(0.40), Inches(ry), Inches(label_w), Inches(rh - 0.02), bg)
        text(s, Inches(0.55), Inches(ry + 0.10), Inches(label_w - 0.20), Inches(rh - 0.20),
             act, size=9, bold=True, color=DARK_TEXT, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
        for ci in range(len(roles)):
            x = 0.40 + label_w + ci * cell_w
            rect(s, Inches(x), Inches(ry), Inches(cell_w - 0.02), Inches(rh - 0.02), bg)
            val = raci[ri][ci]
            cval = color_map.get(val, GREY_LABEL)
            d = 0.42
            oval(s, Inches(x + (cell_w - 0.02 - d) / 2), Inches(ry + (rh - 0.02 - d) / 2),
                 Inches(d), Inches(d), cval)
            text(s, Inches(x), Inches(ry + (rh - 0.02 - d) / 2), Inches(cell_w - 0.02), Inches(d),
                 val, size=10, bold=True, color=WHITE, font=BODY_FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        ry += rh

    # Legend
    leg_y = 7.05
    text(s, Inches(0.40), Inches(leg_y), Inches(1.50), Inches(0.18),
         "RACI", size=8, bold=True, color=GREY_LABEL, font=BODY_FONT)
    leg = [("R · Responsible", AVANGRID_GREEN), ("A · Accountable", GATE_RED),
           ("C · Consulted", HEADER_BLUE), ("I · Informed", GREY_MUTED)]
    sx = 1.80
    for lbl, color in leg:
        oval(s, Inches(sx), Inches(leg_y + 0.02), Inches(0.16), Inches(0.16), color)
        text(s, Inches(sx + 0.22), Inches(leg_y), Inches(2.0), Inches(0.18),
             lbl, size=8, color=DARK_TEXT, font=BODY_FONT)
        sx += 2.40

    footer_bar(s, 14)
    return s


# ============================================================================
# SECTION 4 · TOOLS  (slide 15)
# ============================================================================

# 15 · Four tool families + reporting
def build_pp15(prs):
    s = _slide(prs)
    slide_title(s,
        "Four tool families · live dashboards",
        "Backlog · Roadmap · Estimation · Reporting — Heather + VP Vitor see real-time status",
        section="Tools")

    from_phase1_strip(s)

    # 4 tool family columns
    col_y = 1.65
    col_h = 4.25
    cw = (12.55 - 3 * 0.20) / 4

    families = [
        ("01  ·  BACKLOG + PORTFOLIO",
         "Single source of truth for demand state",
         AVANGRID_GREEN,
         [("WHAT IT DOES",  "Portfolio Board · Backlog ranking · status workflow · audit log"),
          ("CADENCE",        "Real-time"),
          ("AUDIENCE",       "PM (per Squad) · Director · VP Vitor"),
          ("KEY VIEW",       "Cross-VS Backlog ranked by Pain Score + CoS"),
          ("INTEGRATION",    "Feeds Roadmap + Capacity + Reporting tools")]),

        ("02  ·  ROADMAP + PLANNING",
         "Rolling-wave timeline · per VS",
         GREEN_MID,
         [("WHAT IT DOES",  "Timeline view · 3·6·12 mo bands · committed vs hypothesized"),
          ("CADENCE",        "Updated weekly · re-balanced monthly @ Gate 2"),
          ("AUDIENCE",       "Director · VP Vitor · BRS counterparts"),
          ("KEY VIEW",       "Per-VS timeline + cross-VS dependencies overlay"),
          ("INTEGRATION",    "Reads from Backlog · feeds capacity reservations")]),

        ("03  ·  ESTIMATION + CAPACITY",
         "Pain Score · Effort · CoS lane utilization",
         HEADER_BLUE,
         [("WHAT IT DOES",  "Score panel · capacity matrix · WIP tracker · forecasting"),
          ("CADENCE",        "Real-time on score · monthly on capacity %"),
          ("AUDIENCE",       "PM (per Squad) · Architect · Director · Middle managers"),
          ("KEY VIEW",       "Squad × CoS lane utilization · capacity forecast Q+1"),
          ("INTEGRATION",    "Pain Score auto-computed · feeds Reporting")]),

        ("04  ·  REPORTING + DASHBOARDS",
         "Heather + Vitor real-time view",
         HEADER_BLUE_DK,
         [("WHAT IT DOES",  "7 metrics · trend charts · variance commentary · exceptions"),
          ("CADENCE",        "Real-time · monthly review at Gate 2"),
          ("AUDIENCE",       "VP Vitor · Heather · Directors · BRS"),
          ("KEY VIEW",       "Demand × capacity · cycle time · planning predictability"),
          ("INTEGRATION",    "Aggregates from Backlog + Roadmap + Capacity tools")]),
    ]
    for i, (name, sub, color, blocks) in enumerate(families):
        x = 0.40 + i * (cw + 0.20)
        round_rect(s, Inches(x), Inches(col_y), Inches(cw), Inches(col_h), WHITE, radius=0.04)
        rect(s, Inches(x), Inches(col_y), Inches(cw), Inches(0.65), color)
        text(s, Inches(x + 0.15), Inches(col_y + 0.06), Inches(cw - 0.30), Inches(0.25),
             name, size=10, bold=True, color=WHITE, font=BODY_FONT)
        text(s, Inches(x + 0.15), Inches(col_y + 0.32), Inches(cw - 0.30), Inches(0.30),
             sub, size=8, italic=True, color=WHITE, font=BODY_FONT)
        iy = col_y + 0.80
        bh = (col_h - 0.95) / len(blocks)
        for lbl, body in blocks:
            text(s, Inches(x + 0.15), Inches(iy), Inches(cw - 0.30), Inches(0.18),
                 lbl, size=7, bold=True, color=color, font=BODY_FONT)
            text(s, Inches(x + 0.15), Inches(iy + 0.18), Inches(cw - 0.30), Inches(bh - 0.22),
                 body, size=8, color=DARK_TEXT, font=BODY_FONT)
            iy += bh

    # Bottom · the dashboard mockup teaser (Heather + Vitor view)
    db_y = 5.75
    round_rect(s, Inches(0.40), Inches(db_y), Inches(12.55), Inches(1.15), GREEN_BG, radius=0.04)
    rect(s, Inches(0.40), Inches(db_y), Inches(0.10), Inches(1.15), AVANGRID_GREEN)
    text(s, Inches(0.55), Inches(db_y + 0.06), Inches(12.40), Inches(0.20),
         "EXECUTIVE DASHBOARD  ·  Heather + VP Vitor  ·  always live",
         size=10, bold=True, color=GREEN_DARK, font=BODY_FONT)
    # 4 KPI tiles
    kpis = [
        ("BV THROUGHPUT", "$ delivered / quarter", AVANGRID_GREEN),
        ("PAIN VELOCITY", "avg score completed", GREEN_MID),
        ("CYCLE TIME",    "intake → release",     HEADER_BLUE),
        ("CoS ADHERENCE", "% in correct lane",    HEADER_BLUE_DK),
    ]
    kpw = (12.55 - 0.30 - 3 * 0.10) / 4
    for i, (kpi, sub, color) in enumerate(kpis):
        x = 0.55 + i * (kpw + 0.10)
        round_rect(s, Inches(x), Inches(db_y + 0.32), Inches(kpw), Inches(0.75), WHITE, radius=0.04)
        rect(s, Inches(x), Inches(db_y + 0.32), Inches(0.06), Inches(0.75), color)
        text(s, Inches(x + 0.12), Inches(db_y + 0.36), Inches(kpw - 0.20), Inches(0.20),
             kpi, size=8.5, bold=True, color=color, font=BODY_FONT)
        text(s, Inches(x + 0.12), Inches(db_y + 0.55), Inches(kpw - 0.20), Inches(0.20),
             sub, size=8, italic=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, Inches(x + 0.12), Inches(db_y + 0.78), Inches(kpw - 0.20), Inches(0.22),
             "▎▎▎▎▎▎▎▎ ▏▏▏",  size=11, color=color, font=BODY_FONT)

    footer_bar(s, 15)
    return s


# ============================================================================
# SECTION 5 · GOVERNANCE  (slides 16-19)
# ============================================================================

# 16 · Gate 2 forum · the only forum
def build_pp16(prs):
    s = _slide(prs)
    slide_title(s,
        "Gate 2 · the only forum",
        "Monthly · 90 min · VP Vitor chairs · 3 Directors vote · audit-trail every decision",
        section="Governance")

    from_phase1_strip(s)

    # LEFT · meeting layout mockup
    mt_x = 0.40; mt_y = 1.65; mt_w = 7.50; mt_h = 5.20
    round_rect(s, Inches(mt_x), Inches(mt_y), Inches(mt_w), Inches(mt_h), WHITE, radius=0.04)
    rect(s, Inches(mt_x), Inches(mt_y), Inches(mt_w), Inches(0.45), NAVY_DARK)
    text(s, Inches(mt_x + 0.15), Inches(mt_y + 0.08), Inches(mt_w - 0.30), Inches(0.30),
         "GATE 2  ·  Portfolio Decision Forum",
         size=11, bold=True, color=WHITE, font=BODY_FONT)

    # Roundtable visual
    table_x = mt_x + 1.40
    table_y = mt_y + 0.85
    table_w = mt_w - 2.80
    table_h = 1.90
    round_rect(s, Inches(table_x), Inches(table_y), Inches(table_w), Inches(table_h), GREY_LIGHT, radius=0.50)
    text(s, Inches(table_x), Inches(table_y + table_h/2 - 0.15), Inches(table_w), Inches(0.30),
         "Gate 2 · monthly · 90 min",
         size=11, bold=True, color=GREY_LABEL, font=BODY_FONT, align=PP_ALIGN.CENTER)

    # Avatars around the table · roles
    seats = [
        (table_x + table_w/2 - 0.30, table_y - 0.50, "VV", HEADER_BLUE_DK, "VP Vitor", "chairs"),
        (table_x - 0.55, table_y + 0.20, "HK", AVANGRID_GREEN, "Heather", "Director · Product"),
        (table_x - 0.55, table_y + table_h - 0.80, "KJ", AVANGRID_GREEN, "Kumar", "Director · CS"),
        (table_x + table_w + 0.05, table_y + 0.20, "CG", HEADER_BLUE, "Carlos", "Director · ITBA"),
        (table_x + table_w + 0.05, table_y + table_h - 0.80, "MM", GREY_LABEL, "Middle mgrs", "advise"),
        (table_x + table_w/2 - 0.30, table_y + table_h + 0.15, "PA", INSIGHT_AMBER, "PMs + Architect", "support"),
    ]
    for x, y, init, color, name, role in seats:
        avatar(s, x, y, 0.55, color, init, size_factor=18)
        text(s, Inches(x - 0.40), Inches(y + 0.55), Inches(1.30), Inches(0.18),
             name, size=8, bold=True, color=DARK_TEXT, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, Inches(x - 0.40), Inches(y + 0.71), Inches(1.30), Inches(0.18),
             role, size=7, italic=True, color=GREY_LABEL, font=BODY_FONT, align=PP_ALIGN.CENTER)

    # Agenda below
    ag_y = mt_y + 3.10
    text(s, Inches(mt_x + 0.20), Inches(ag_y), Inches(mt_w - 0.30), Inches(0.20),
         "AGENDA  ·  90 min", size=8, bold=True, color=GREY_LABEL, font=BODY_FONT)
    agenda = [
        ("00-15",  "Capacity reality check  ·  middle managers report"),
        ("15-35",  "Backlog re-rank · CS + ITBA reviewed independently · Pain Score + CoS"),
        ("35-55",  "Defer/Reject decisions  ·  audit-logged"),
        ("55-75",  "Roadmap rolling-wave realignment  ·  commits vs hypotheses"),
        ("75-85",  "Cross-area dependencies  ·  blockers · escalations between CS and ITBA"),
        ("85-90",  "BRP outputs digest  ·  feeds quarterly priorities to P&T + CoE BRPs"),
    ]
    iy = ag_y + 0.25
    rh = 0.30
    for time_str, desc in agenda:
        rect(s, Inches(mt_x + 0.20), Inches(iy), Inches(0.95), Inches(rh - 0.04), AVANGRID_GREEN)
        text(s, Inches(mt_x + 0.20), Inches(iy), Inches(0.95), Inches(rh - 0.04),
             time_str, size=8, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, Inches(mt_x + 1.25), Inches(iy + 0.04), Inches(mt_w - 1.50), Inches(rh - 0.10),
             desc, size=8.5, color=DARK_TEXT, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
        iy += rh

    # RIGHT · governance principles
    gp_x = mt_x + mt_w + 0.20
    gp_w = 13.33 - gp_x - 0.40
    gp_h = mt_h
    round_rect(s, Inches(gp_x), Inches(mt_y), Inches(gp_w), Inches(gp_h), WHITE, radius=0.04)
    rect(s, Inches(gp_x), Inches(mt_y), Inches(gp_w), Inches(0.45), AVANGRID_GREEN)
    text(s, Inches(gp_x + 0.15), Inches(mt_y + 0.08), Inches(gp_w - 0.30), Inches(0.30),
         "PRINCIPLES", size=11, bold=True, color=WHITE, font=BODY_FONT)

    principles = [
        ("ONE FORUM", "All Phase-2 decisions land here · no parallel committee"),
        ("3-VOICE",   "Heather + Kumar + Carlos must concur · VP breaks ties"),
        ("AUDIT",     "Every Defer/Reject + re-rank logged with reason · visible to all"),
        ("METRICS",   "Heather + VP Vitor metrics drive performance review"),
        ("BRP COUPLING","Outputs feed quarterly digest to P&T BRP + Digital CoE BRP"),
        ("NO PARALLEL","No quarterly event of our own · couples to existing"),
    ]
    iy = mt_y + 0.55
    rh = (gp_h - 0.65) / len(principles)
    for lbl, body in principles:
        chip(s, gp_x + 0.15, iy + 0.10, 1.30, lbl, AVANGRID_GREEN, WHITE, h=0.24, size=8)
        text(s, Inches(gp_x + 0.15), Inches(iy + 0.40), Inches(gp_w - 0.30), Inches(rh - 0.45),
             body, size=8.5, color=DARK_TEXT, font=BODY_FONT)
        iy += rh

    footer_bar(s, 16)
    return s


# 17 · Prioritization criteria (deeper than slide 07)
def build_pp17(prs):
    s = _slide(prs)
    slide_title(s,
        "Prioritization criteria · 5 layers",
        "Value · Risk · Effort · OPCO scope · Overrides — same logic CS+ITBA · different inputs",
        section="Governance")

    from_phase1_strip(s)

    # 5 columns of criteria
    col_y = 1.65
    col_h = 5.55
    cw = (12.55 - 4 * 0.15) / 5

    crits = [
        ("01", "VALUE",
         "BV (1-5)  +  UX (1-5)",
         "PM (per Squad) scores at Analyzing",
         "ITBA: BV weight 0.45 · CS: BV weight 0.40\nITBA: UX weight 0.30 · CS: UX weight 0.60",
         AVANGRID_GREEN),
        ("02", "RISK",
         "Safety (1-5) · ITBA only",
         "Architect scores at Analyzing",
         "Field-crew impact · isolated from BV\nWeight 0.25 (ITBA) · 0 (CS)",
         HEADER_BLUE),
        ("03", "EFFORT",
         "1-5 scale · Architect-defined",
         "Architect scores at Analyzing",
         "1=hours · 2=days · 3=weeks\n4=quarter · 5=multi-quarter",
         HEADER_BLUE_DK),
        ("04", "OPCO SCOPE",
         "OPCOmult automatic",
         "Auto from OPCO field at Funnel",
         "1.0 (1 OPCO) · 1.3 (2-4 regional)\n1.6 (5+ enterprise)",
         GREEN_MID),
        ("05", "OVERRIDES",
         "Type=Reg → Fixed-Date\nCIT P1/P2 → Expedite",
         "Auto · A2 + A3 fire",
         "Bypass Pain Score ranking\nProtect deadline-bound + critical",
         GATE_RED),
    ]
    for i, (num, name, what, who, body, color) in enumerate(crits):
        x = 0.40 + i * (cw + 0.15)
        round_rect(s, Inches(x), Inches(col_y), Inches(cw), Inches(col_h), WHITE, radius=0.04)
        rect(s, Inches(x), Inches(col_y), Inches(cw), Inches(0.85), color)
        text(s, Inches(x + 0.15), Inches(col_y + 0.06), Inches(0.70), Inches(0.40),
             num, size=24, bold=True, color=WHITE, font=BODY_FONT)
        text(s, Inches(x + 0.85), Inches(col_y + 0.10), Inches(cw - 0.95), Inches(0.30),
             name, size=11, bold=True, color=WHITE, font=BODY_FONT)
        text(s, Inches(x + 0.85), Inches(col_y + 0.42), Inches(cw - 0.95), Inches(0.40),
             what, size=8.5, color=WHITE, font=BODY_FONT)

        # Body
        iy = col_y + 1.00
        text(s, Inches(x + 0.15), Inches(iy), Inches(cw - 0.30), Inches(0.20),
             "OWNER", size=7, bold=True, color=color, font=BODY_FONT)
        text(s, Inches(x + 0.15), Inches(iy + 0.20), Inches(cw - 0.30), Inches(0.40),
             who, size=8, color=DARK_TEXT, font=BODY_FONT)

        text(s, Inches(x + 0.15), Inches(iy + 0.75), Inches(cw - 0.30), Inches(0.20),
             "DETAIL", size=7, bold=True, color=color, font=BODY_FONT)
        text(s, Inches(x + 0.15), Inches(iy + 0.95), Inches(cw - 0.30), Inches(2.00),
             body, size=8, color=DARK_TEXT, font=BODY_FONT)

    # Tie-break rule at bottom
    tb_y = 6.00
    insight_callout(s, 0.40, tb_y, 12.55, 0.85,
        "Tie-break rule: Pain Score equal? Then CoS lane > capacity available > Sponsor seniority — never \"squeaky-wheel\".",
        "Example · two demands at Pain 2.4: regulatory with Fixed-Date wins over Standard CS · if both Standard, the one fitting current capacity wins · if both fit, the one with named Director Sponsor wins · audit logged at Gate 2")

    footer_bar(s, 18)
    return s


# 18 · Demand × Capacity oversight
def build_pp18(prs):
    s = _slide(prs)
    slide_title(s,
        "Demand × Capacity · the balance check",
        "Each area checks independently · pending cards vs available FTE per CoS lane · gap escalates at Gate 2",
        section="Governance")

    from_phase1_strip(s)

    # LEFT · pending demand stack
    pd_x = 0.40; pd_y = 1.65; pd_w = 5.85; pd_h = 5.55
    round_rect(s, Inches(pd_x), Inches(pd_y), Inches(pd_w), Inches(pd_h), WHITE, radius=0.04)
    rect(s, Inches(pd_x), Inches(pd_y), Inches(pd_w), Inches(0.45), AVANGRID_GREEN)
    text(s, Inches(pd_x + 0.15), Inches(pd_y + 0.08), Inches(pd_w - 0.30), Inches(0.30),
         "PENDING DEMAND  ·  in Backlog (RFE)  ·  per area",
         size=11, bold=True, color=WHITE, font=BODY_FONT)

    # Stacked demand bars by CoS
    cos_demand = [
        ("Standard",   COS_GREEN,  18, "items"),
        ("Fixed-Date", COS_BLUE,    7, "items"),
        ("Expedite",   COS_RED,     2, "items"),
        ("Intangible", COS_ORANGE,  4, "items"),
    ]
    iy = pd_y + 0.65
    rh = (pd_h - 0.85) / len(cos_demand)
    max_count = max(c[2] for c in cos_demand)
    for lane, color, count, unit in cos_demand:
        round_rect(s, Inches(pd_x + 0.15), Inches(iy), Inches(pd_w - 0.30), Inches(rh - 0.10),
                   GREY_LIGHT, radius=0.04)
        chip(s, pd_x + 0.20, iy + (rh - 0.30) / 2, 1.30, lane, color, WHITE, h=0.30, size=9)
        # Bar
        bar_x = pd_x + 1.65
        bar_max_w = pd_w - 2.20
        rect(s, Inches(bar_x), Inches(iy + 0.20), Inches(bar_max_w), Inches(rh - 0.50), GREY_LIGHT)
        rect(s, Inches(bar_x), Inches(iy + 0.20),
             Inches(bar_max_w * count / max_count), Inches(rh - 0.50), color)
        text(s, Inches(pd_x + pd_w - 0.85), Inches(iy + 0.10), Inches(0.65), Inches(rh - 0.20),
             str(count), size=14, bold=True, color=NAVY_DARK, font=BODY_FONT,
             anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)
        iy += rh

    # RIGHT · available capacity
    ac_x = pd_x + pd_w + 0.20
    ac_w = 13.33 - ac_x - 0.40
    ac_h = pd_h
    round_rect(s, Inches(ac_x), Inches(pd_y), Inches(ac_w), Inches(ac_h), WHITE, radius=0.04)
    rect(s, Inches(ac_x), Inches(pd_y), Inches(ac_w), Inches(0.45), HEADER_BLUE)
    text(s, Inches(ac_x + 0.15), Inches(pd_y + 0.08), Inches(ac_w - 0.30), Inches(0.30),
         "AVAILABLE CAPACITY  ·  next 30 days  ·  per area",
         size=11, bold=True, color=WHITE, font=BODY_FONT)

    cos_cap = [
        ("Standard",   COS_GREEN,  60, "% × 9 squads = 38 SP/wk"),
        ("Fixed-Date", COS_BLUE,    20, "% × 9 squads = 13 SP/wk"),
        ("Expedite",   COS_RED,     10, "% × 9 squads = 6 SP/wk"),
        ("Intangible", COS_ORANGE,  10, "% × 9 squads = 6 SP/wk"),
    ]
    iy = pd_y + 0.65
    rh = (ac_h - 0.85) / len(cos_cap)
    for lane, color, pct, detail in cos_cap:
        round_rect(s, Inches(ac_x + 0.15), Inches(iy), Inches(ac_w - 0.30), Inches(rh - 0.10),
                   GREY_LIGHT, radius=0.04)
        chip(s, ac_x + 0.20, iy + (rh - 0.30) / 2, 1.30, lane, color, WHITE, h=0.30, size=9)
        bar_x = ac_x + 1.65
        bar_max_w = ac_w - 2.20
        rect(s, Inches(bar_x), Inches(iy + 0.20), Inches(bar_max_w), Inches(rh - 0.50), GREY_LIGHT)
        rect(s, Inches(bar_x), Inches(iy + 0.20),
             Inches(bar_max_w * pct / 100), Inches(rh - 0.50), color)
        text(s, Inches(bar_x + 0.10), Inches(iy + 0.10), Inches(bar_max_w - 0.20), Inches(0.20),
             f"{pct}%", size=8, bold=True, color=NAVY_DARK, font=BODY_FONT)
        text(s, Inches(bar_x), Inches(iy + rh - 0.30), Inches(bar_max_w), Inches(0.20),
             detail, size=7, italic=True, color=GREY_LABEL, font=BODY_FONT)
        iy += rh

    # Bottom · gap analysis insight
    insight_callout(s, 0.40, 6.00, 12.55, 0.85,
        "Each area runs this check separately — illustrative numbers shown represent ONE area at a time.",
        "CS Director (Kumar + Heather) reviews CS gap · ITBA Director (Carlos + Heather) reviews ITBA gap · if pending > 4× weekly capacity, escalation to VP Vitor for capacity reallocation OR Defer/Reject")

    cos_legend(s)
    footer_bar(s, 19)
    return s


# 19 · Defer / Reject decision practices
def build_pp19(prs):
    s = _slide(prs)
    slide_title(s,
        "Defer or Reject · 3 triggers",
        "Capacity full · Strategic mismatch · Sponsor unconfirmed — area Director decides at Gate 2 · audit log",
        section="Governance")

    from_phase1_strip(s)

    # Decision tree visual
    root_x = 5.40; root_y = 1.30; root_w = 2.55; root_h = 0.75
    round_rect(s, Inches(root_x), Inches(root_y), Inches(root_w), Inches(root_h), NAVY_DARK, radius=0.10)
    text(s, Inches(root_x), Inches(root_y + 0.10), Inches(root_w), Inches(0.30),
         "DEMAND ARRIVES AT", size=10, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
    text(s, Inches(root_x), Inches(root_y + 0.40), Inches(root_w), Inches(0.30),
         "GATE 2 CHECK", size=10, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)

    # 3 trigger branches
    trig_y = 2.85
    trig_h = 0.55
    triggers = [
        ("CAPACITY FULL?",      "Lane WIP at cap · no slot for next quarter",  AMBER := COS_ORANGE),
        ("STRATEGIC MISMATCH?", "BRP outputs deprioritize this area",           HEADER_BLUE),
        ("SPONSOR UNCONFIRMED?","Sponsor not named · UAT path unsigned",         GATE_RED),
    ]
    n = len(triggers)
    branch_w = (12.55 - 0.40 - (n - 1) * 0.30) / n
    for i, (q, desc, color) in enumerate(triggers):
        bx = 0.40 + i * (branch_w + 0.30)
        line(s, Inches(root_x + root_w/2), Inches(root_y + root_h),
             Inches(bx + branch_w/2), Inches(trig_y), color, weight=2)
        round_rect(s, Inches(bx), Inches(trig_y), Inches(branch_w), Inches(trig_h), color, radius=0.04)
        text(s, Inches(bx), Inches(trig_y + 0.04), Inches(branch_w), Inches(0.22),
             q, size=11, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, Inches(bx), Inches(trig_y + 0.28), Inches(branch_w), Inches(0.22),
             desc, size=8, italic=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)

    # Outcomes per trigger (DEFER vs REJECT)
    out_y = 3.65
    out_h = 1.50
    outcomes = [
        ("DEFER  ·  to next quarter",
         "Item stays in Backlog · re-ranked next Gate 2 · Sponsor notified",
         "Most common · low cost · keeps optionality",
         AMBER),
        ("DEFER  ·  to next BRP cycle",
         "Item parked · waiting on strategic alignment · BRS counterpart notified",
         "Sponsor + BRS escalate via P&T or Digital CoE BRP",
         HEADER_BLUE),
        ("REJECT  ·  back to Sponsor",
         "Item closed · Sponsor must re-submit when criteria met",
         "Sponsor re-confirms acceptance criteria + UAT path · re-enters Funnel",
         GATE_RED),
    ]
    for i, (decision, body, why, color) in enumerate(outcomes):
        bx = 0.40 + i * (branch_w + 0.30)
        line(s, Inches(bx + branch_w/2), Inches(trig_y + trig_h),
             Inches(bx + branch_w/2), Inches(out_y), color, weight=1.5)
        round_rect(s, Inches(bx), Inches(out_y), Inches(branch_w), Inches(out_h), WHITE, radius=0.04)
        rect(s, Inches(bx), Inches(out_y), Inches(0.10), Inches(out_h), color)
        text(s, Inches(bx + 0.20), Inches(out_y + 0.06), Inches(branch_w - 0.30), Inches(0.25),
             decision, size=10, bold=True, color=color, font=BODY_FONT)
        text(s, Inches(bx + 0.20), Inches(out_y + 0.34), Inches(branch_w - 0.30), Inches(0.50),
             body, size=8.5, color=DARK_TEXT, font=BODY_FONT)
        text(s, Inches(bx + 0.20), Inches(out_y + 0.95), Inches(branch_w - 0.30), Inches(out_h - 0.95),
             why, size=8, italic=True, color=GREY_TEXT, font=BODY_FONT)

    # Audit trail bar
    audit_y = 5.40
    round_rect(s, Inches(0.40), Inches(audit_y), Inches(12.55), Inches(1.10), NAVY_DARK, radius=0.04)
    text(s, Inches(0.55), Inches(audit_y + 0.06), Inches(12.40), Inches(0.22),
         "AUDIT TRAIL  ·  every Defer/Reject logged in the Decision Log",
         size=10, bold=True, color=WHITE, font=BODY_FONT)
    audit_items = [
        "Demand ID + title + Sponsor",
        "Trigger (Capacity / Strategic / Sponsor)",
        "Decision (Defer-Quarter / Defer-BRP / Reject)",
        "Decider (CS Director Kumar/Heather OR ITBA Director Carlos)",
        "Date + Gate 2 session reference",
    ]
    iy = audit_y + 0.30
    text(s, Inches(0.55), Inches(iy), Inches(12.40), Inches(0.65),
         "  ·  ".join(audit_items), size=8.5, color=BLUE_PALE, font=BODY_FONT)
    text(s, Inches(0.55), Inches(iy + 0.40), Inches(12.40), Inches(0.30),
         "Visible to · VP Vitor + Director trio + middle managers + Sponsors of affected items + BRS",
         size=8, italic=True, color=BLUE_PALE, font=BODY_FONT)

    footer_bar(s, 20)
    return s


# ============================================================================
# SECTION 6 · PERFORMANCE  (slides 20-22)
# ============================================================================

# 20 · Performance metrics overview · 7 KPIs
def build_pp20(prs):
    s = _slide(prs)
    slide_title(s,
        "Seven performance metrics · the contract",
        "Same 7 KPIs as Demand Management · applied through Phase 2 · live dashboard for VP Vitor + Heather",
        section="Performance")

    from_phase1_strip(s)

    # 7 KPI tiles in a grid (4 + 3)
    kpis = [
        ("DEMAND VOLUME",      "submissions / week · by VS · CoS · OPCO",  "trend visible", "▲ baseline",   AVANGRID_GREEN),
        ("DoR PASS RATE",      "% accepted at first submit",                "≥ 70 %",        "target ≥ 70 %", GREEN_MID),
        ("PHASE 2 LEAD TIME",  "Reviewing → DoR pass (days · P50)",         "≤ 10 d",        "target ≤ 10 d", HEADER_BLUE),
        ("CHANNEL ADOPTION",   "% via single front door",                   "≥ 90 %",        "target month 6",HEADER_BLUE_DK),
        ("EXPEDITE SHARE",     "% of throughput",                            "≤ 20 %",        "cap 20 %",      COS_RED),
        ("DUPLICATION RATE",   "% submissions matched",                      "↓ trend",       "target down",   COS_BLUE),
        ("STATUS VISIBILITY",  "% with live status",                         "100 %",         "target 100 %",  COS_ORANGE),
    ]
    # 4 in row 1 + 3 in row 2
    row1_y = 1.65
    row_h = 2.65
    cw = (12.55 - 3 * 0.15) / 4

    for i in range(4):
        kpi, sub, val, trend, color = kpis[i]
        x = 0.40 + i * (cw + 0.15)
        round_rect(s, Inches(x), Inches(row1_y), Inches(cw), Inches(row_h), WHITE, radius=0.04)
        rect(s, Inches(x), Inches(row1_y), Inches(0.10), Inches(row_h), color)
        text(s, Inches(x + 0.20), Inches(row1_y + 0.10), Inches(cw - 0.30), Inches(0.30),
             kpi, size=9, bold=True, color=color, font=BODY_FONT)
        text(s, Inches(x + 0.20), Inches(row1_y + 0.40), Inches(cw - 0.30), Inches(0.30),
             sub, size=8, italic=True, color=GREY_LABEL, font=BODY_FONT)
        # Big value
        text(s, Inches(x + 0.20), Inches(row1_y + 0.80), Inches(cw - 0.30), Inches(0.55),
             val, size=22, bold=True, color=NAVY_DARK, font=BODY_FONT)
        # Trend
        text(s, Inches(x + 0.20), Inches(row1_y + 1.40), Inches(cw - 0.30), Inches(0.30),
             trend, size=10, bold=True, color=color, font=BODY_FONT)
        # Mini sparkline (visual fill)
        spark_y = row1_y + 1.85
        bar_w = (cw - 0.40) / 8
        heights = [0.20, 0.30, 0.25, 0.40, 0.35, 0.50, 0.45, 0.55]
        for j, hf in enumerate(heights):
            rect(s, Inches(x + 0.20 + j * (bar_w + 0.04)), Inches(spark_y + 0.55 - hf),
                 Inches(bar_w), Inches(hf), color)

    # Row 2 · 3 wider tiles
    row2_y = row1_y + row_h + 0.20
    cw2 = (12.55 - 2 * 0.15) / 3
    for i in range(3):
        kpi, sub, val, trend, color = kpis[i + 4]
        x = 0.40 + i * (cw2 + 0.15)
        round_rect(s, Inches(x), Inches(row2_y), Inches(cw2), Inches(row_h - 0.20), WHITE, radius=0.04)
        rect(s, Inches(x), Inches(row2_y), Inches(0.10), Inches(row_h - 0.20), color)
        text(s, Inches(x + 0.20), Inches(row2_y + 0.10), Inches(cw2 - 0.30), Inches(0.30),
             kpi, size=10, bold=True, color=color, font=BODY_FONT)
        text(s, Inches(x + 0.20), Inches(row2_y + 0.40), Inches(cw2 - 0.30), Inches(0.30),
             sub, size=8, italic=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, Inches(x + 0.20), Inches(row2_y + 0.80), Inches(cw2 - 0.30), Inches(0.55),
             val, size=24, bold=True, color=NAVY_DARK, font=BODY_FONT)
        text(s, Inches(x + 0.20), Inches(row2_y + 1.40), Inches(cw2 - 0.30), Inches(0.30),
             trend, size=10, bold=True, color=color, font=BODY_FONT)
        spark_y = row2_y + 1.85
        bar_w = (cw2 - 0.40) / 8
        heights = [0.20, 0.30, 0.25, 0.40, 0.35, 0.50, 0.45, 0.55]
        for j, hf in enumerate(heights):
            rect(s, Inches(x + 0.20 + j * (bar_w + 0.04)), Inches(spark_y + 0.50 - hf),
                 Inches(bar_w), Inches(hf), color)

    insight_callout(s, 0.40, 6.40, 12.55, 0.50,
        "These seven KPIs are the contract — if any one drifts, the funnel is leaking.",
        "Live dashboard means leakage is visible the next day, not the next quarter · same KPIs reviewed for Demand Management apply through Phase 2")

    footer_bar(s, 22)
    return s


# 21 · Cycle time + planning predictability charts
def build_pp21(prs):
    s = _slide(prs)
    slide_title(s,
        "Cycle time + predictability",
        "12-month rolling · CS vs ITBA · target overlay · variance commentary monthly",
        section="Performance")

    from_phase1_strip(s)

    # LEFT chart: cycle time trend
    ch1_x = 0.40; ch1_y = 1.65; ch1_w = 6.30; ch1_h = 4.30
    round_rect(s, Inches(ch1_x), Inches(ch1_y), Inches(ch1_w), Inches(ch1_h), WHITE, radius=0.04)
    rect(s, Inches(ch1_x), Inches(ch1_y), Inches(ch1_w), Inches(0.45), AVANGRID_GREEN)
    text(s, Inches(ch1_x + 0.15), Inches(ch1_y + 0.08), Inches(ch1_w - 0.30), Inches(0.30),
         "CYCLE TIME  ·  intake → DoR pass (days · median)",
         size=10, bold=True, color=WHITE, font=BODY_FONT)

    # Chart area
    cax = ch1_x + 0.50; cay = ch1_y + 0.65
    caw = ch1_w - 0.80; cah = ch1_h - 1.10
    rect(s, Inches(cax), Inches(cay + cah), Inches(caw), Emu(15000), GREY_BORDER)  # x-axis
    rect(s, Inches(cax), Inches(cay), Emu(15000), Inches(cah), GREY_BORDER)  # y-axis
    # Months
    months_x = ["Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar","Apr","May","Jun","Jul"]
    mw = caw / (len(months_x) - 1)
    for i, m in enumerate(months_x):
        text(s, Inches(cax + i * mw - 0.15), Inches(cay + cah + 0.06), Inches(0.30), Inches(0.18),
             m, size=7, color=GREY_LABEL, font=BODY_FONT, align=PP_ALIGN.CENTER)
    # Y-axis labels (days)
    y_vals = [10, 15, 20, 25, 30]
    for i, yv in enumerate(y_vals):
        ratio = (yv - 10) / 20
        ypy = cay + cah - cah * ratio
        text(s, Inches(cax - 0.45), Inches(ypy - 0.08), Inches(0.40), Inches(0.18),
             str(yv), size=7, color=GREY_LABEL, font=BODY_FONT, align=PP_ALIGN.RIGHT)
        rect(s, Inches(cax), Inches(ypy), Inches(caw), Emu(5000), GREY_BORDER)

    # CS line (12 months)
    cs_data = [28, 26, 24, 25, 22, 21, 20, 19, 18, 17, 18, 16]
    itba_data = [30, 29, 27, 28, 25, 24, 23, 22, 21, 20, 19, 19]
    target = 18
    # Target line
    ratio = (target - 10) / 20
    ty = cay + cah - cah * ratio
    rect(s, Inches(cax), Inches(ty), Inches(caw), Emu(15000), INSIGHT_AMBER)
    text(s, Inches(cax + caw - 1.0), Inches(ty - 0.18), Inches(1.0), Inches(0.18),
         f"target {target}d", size=7, bold=True, color=INSIGHT_AMBER, font=BODY_FONT, align=PP_ALIGN.RIGHT)

    # Plot CS line
    for i in range(len(cs_data) - 1):
        x1 = cax + i * mw
        y1 = cay + cah - cah * (cs_data[i] - 10) / 20
        x2 = cax + (i + 1) * mw
        y2 = cay + cah - cah * (cs_data[i + 1] - 10) / 20
        line(s, Inches(x1), Inches(y1), Inches(x2), Inches(y2), AVANGRID_GREEN, weight=2.5)
        oval(s, Inches(x1 - 0.05), Inches(y1 - 0.05), Inches(0.10), Inches(0.10), AVANGRID_GREEN)
    # Last point
    oval(s, Inches(cax + (len(cs_data)-1) * mw - 0.05),
         Inches(cay + cah - cah * (cs_data[-1] - 10) / 20 - 0.05),
         Inches(0.10), Inches(0.10), AVANGRID_GREEN)

    # Plot ITBA line
    for i in range(len(itba_data) - 1):
        x1 = cax + i * mw
        y1 = cay + cah - cah * (itba_data[i] - 10) / 20
        x2 = cax + (i + 1) * mw
        y2 = cay + cah - cah * (itba_data[i + 1] - 10) / 20
        line(s, Inches(x1), Inches(y1), Inches(x2), Inches(y2), HEADER_BLUE, weight=2.5)
        oval(s, Inches(x1 - 0.05), Inches(y1 - 0.05), Inches(0.10), Inches(0.10), HEADER_BLUE)
    oval(s, Inches(cax + (len(itba_data)-1) * mw - 0.05),
         Inches(cay + cah - cah * (itba_data[-1] - 10) / 20 - 0.05),
         Inches(0.10), Inches(0.10), HEADER_BLUE)

    # Legend
    text(s, Inches(ch1_x + 0.50), Inches(ch1_y + ch1_h - 0.32), Inches(0.30), Inches(0.20),
         "●", size=14, bold=True, color=AVANGRID_GREEN, font=BODY_FONT)
    text(s, Inches(ch1_x + 0.75), Inches(ch1_y + ch1_h - 0.30), Inches(1.0), Inches(0.20),
         "CS", size=9, bold=True, color=DARK_TEXT, font=BODY_FONT)
    text(s, Inches(ch1_x + 1.20), Inches(ch1_y + ch1_h - 0.32), Inches(0.30), Inches(0.20),
         "●", size=14, bold=True, color=HEADER_BLUE, font=BODY_FONT)
    text(s, Inches(ch1_x + 1.45), Inches(ch1_y + ch1_h - 0.30), Inches(1.0), Inches(0.20),
         "ITBA", size=9, bold=True, color=DARK_TEXT, font=BODY_FONT)
    text(s, Inches(ch1_x + ch1_w - 1.50), Inches(ch1_y + ch1_h - 0.30), Inches(1.40), Inches(0.20),
         "CS 16d  ·  ITBA 19d", size=9, bold=True, color=DARK_TEXT, font=BODY_FONT, align=PP_ALIGN.RIGHT)

    # RIGHT chart: predictability
    ch2_x = ch1_x + ch1_w + 0.20; ch2_y = ch1_y; ch2_w = 13.33 - ch2_x - 0.40; ch2_h = ch1_h
    round_rect(s, Inches(ch2_x), Inches(ch2_y), Inches(ch2_w), Inches(ch2_h), WHITE, radius=0.04)
    rect(s, Inches(ch2_x), Inches(ch2_y), Inches(ch2_w), Inches(0.45), HEADER_BLUE)
    text(s, Inches(ch2_x + 0.15), Inches(ch2_y + 0.08), Inches(ch2_w - 0.30), Inches(0.30),
         "PLANNING PREDICTABILITY  ·  planned vs delivered (commit %)",
         size=10, bold=True, color=WHITE, font=BODY_FONT)

    # 4 quarter bars (Q4 24 → Q3 25 example)
    cax2 = ch2_x + 0.60; cay2 = ch2_y + 0.65
    caw2 = ch2_w - 0.90; cah2 = ch2_h - 1.10
    rect(s, Inches(cax2), Inches(cay2 + cah2), Inches(caw2), Emu(15000), GREY_BORDER)
    quarters_x = ["Q4 25", "Q1 26", "Q2 26", "Q3 26"]
    qbw = caw2 / 4
    bar_w = qbw * 0.6
    cs_pred = [70, 75, 80, 86]
    itba_pred = [65, 72, 78, 84]
    for i, q in enumerate(quarters_x):
        # CS bar
        bx = cax2 + i * qbw + 0.10
        h_cs = cah2 * cs_pred[i] / 100
        rect(s, Inches(bx), Inches(cay2 + cah2 - h_cs), Inches(bar_w / 2 - 0.04), Inches(h_cs), AVANGRID_GREEN)
        text(s, Inches(bx), Inches(cay2 + cah2 - h_cs - 0.22), Inches(bar_w / 2 - 0.04), Inches(0.20),
             f"{cs_pred[i]}%", size=8, bold=True, color=AVANGRID_GREEN, font=BODY_FONT, align=PP_ALIGN.CENTER)
        # ITBA bar
        bx2 = bx + bar_w / 2 + 0.04
        h_it = cah2 * itba_pred[i] / 100
        rect(s, Inches(bx2), Inches(cay2 + cah2 - h_it), Inches(bar_w / 2 - 0.04), Inches(h_it), HEADER_BLUE)
        text(s, Inches(bx2), Inches(cay2 + cah2 - h_it - 0.22), Inches(bar_w / 2 - 0.04), Inches(0.20),
             f"{itba_pred[i]}%", size=8, bold=True, color=HEADER_BLUE, font=BODY_FONT, align=PP_ALIGN.CENTER)
        # Quarter label
        text(s, Inches(cax2 + i * qbw), Inches(cay2 + cah2 + 0.06), Inches(qbw), Inches(0.20),
             q, size=8, color=GREY_LABEL, font=BODY_FONT, align=PP_ALIGN.CENTER)
    # Target line
    tgt_pct = 90
    h_tgt = cah2 * tgt_pct / 100
    rect(s, Inches(cax2), Inches(cay2 + cah2 - h_tgt), Inches(caw2), Emu(15000), INSIGHT_AMBER)
    text(s, Inches(cax2 + caw2 - 0.80), Inches(cay2 + cah2 - h_tgt - 0.20), Inches(0.80), Inches(0.18),
         f"target {tgt_pct}%", size=7, bold=True, color=INSIGHT_AMBER, font=BODY_FONT, align=PP_ALIGN.RIGHT)

    # Bottom · variance commentary
    insight_callout(s, 0.40, 5.75, 12.55, 1.10,
        "Cycle time trending toward target · predictability up 16pp YoY · ITBA 3 days behind CS due to SAP regression cycles.",
        "Heather + VP Vitor review monthly · variance >5pp triggers root-cause analysis at next Gate 2 · CS Q3 hit 86% commit predictability vs ITBA 84% · gap closing as SAP regression bench is formalized")

    footer_bar(s, 23)
    return s


# 22 · Roadmap + outcome reporting · Heather's monthly view to VP
def build_pp22(prs):
    s = _slide(prs)
    slide_title(s,
        "Monthly report to VP Vitor",
        "Roadmap + outcomes + variance + asks · CS and ITBA shown side-by-side · feeds quarterly BRP digest",
        section="Performance")

    from_phase1_strip(s)

    # Dashboard mockup
    dash_x = 0.40; dash_y = 1.65; dash_w = 12.55; dash_h = 5.55
    round_rect(s, Inches(dash_x), Inches(dash_y), Inches(dash_w), Inches(dash_h), WHITE, radius=0.04)
    rect(s, Inches(dash_x), Inches(dash_y), Inches(dash_w), Inches(0.50), NAVY_DARK)
    text(s, Inches(dash_x + 0.20), Inches(dash_y + 0.08), Inches(8), Inches(0.30),
         "DIGITAL PRODUCTS  ·  Phase 2 Monthly Report  ·  Q3 2026",
         size=12, bold=True, color=WHITE, font=BODY_FONT)
    text(s, Inches(dash_x + dash_w - 3.5), Inches(dash_y + 0.13), Inches(3.30), Inches(0.22),
         "Reviewed: Heather LaJoie · VP Vitor Cruz",
         size=8, italic=True, color=BLUE_PALE, font=BODY_FONT, align=PP_ALIGN.RIGHT)

    # Top KPI strip · 4 mini tiles
    kpi_y = dash_y + 0.65
    kpis = [
        ("BV THROUGHPUT", "$ 2.4M", "+18% YoY", AVANGRID_GREEN),
        ("CYCLE TIME",    "18 days", "-4d MoM",  HEADER_BLUE),
        ("PREDICTABILITY","86 %",   "+5pp YoY", COS_GREEN),
        ("PAIN VELOCITY", "2.18",   "→ flat",   COS_ORANGE),
    ]
    kw = (dash_w - 0.40 - 3 * 0.10) / 4
    for i, (lbl, val, tr, color) in enumerate(kpis):
        x = dash_x + 0.20 + i * (kw + 0.10)
        round_rect(s, Inches(x), Inches(kpi_y), Inches(kw), Inches(0.85), GREY_LIGHT, radius=0.04)
        rect(s, Inches(x), Inches(kpi_y), Inches(0.06), Inches(0.85), color)
        text(s, Inches(x + 0.15), Inches(kpi_y + 0.05), Inches(kw - 0.30), Inches(0.18),
             lbl, size=7, bold=True, color=color, font=BODY_FONT)
        text(s, Inches(x + 0.15), Inches(kpi_y + 0.22), Inches(kw - 0.30), Inches(0.30),
             val, size=14, bold=True, color=NAVY_DARK, font=BODY_FONT)
        text(s, Inches(x + 0.15), Inches(kpi_y + 0.55), Inches(kw - 0.30), Inches(0.22),
             tr, size=8, bold=True, color=color, font=BODY_FONT)

    # Roadmap mini-timeline
    rm_y = kpi_y + 1.05
    text(s, Inches(dash_x + 0.20), Inches(rm_y), Inches(dash_w - 0.30), Inches(0.20),
         "ROADMAP  ·  Q3 commits + Q4 plans + 2027 hypotheses",
         size=9, bold=True, color=NAVY_DARK, font=BODY_FONT)
    rm_chart_y = rm_y + 0.25
    rm_chart_h = 1.80
    rm_chart_w = dash_w - 0.40
    # 3 bands
    qw_rm = rm_chart_w / 4
    for i, (lbl, color) in enumerate([("Q3 26", AVANGRID_GREEN), ("Q4 26", HEADER_BLUE),
                                       ("Q1 27", GREY_LABEL), ("Q2 27", GREY_LABEL)]):
        rect(s, Inches(dash_x + 0.20 + i * qw_rm), Inches(rm_chart_y), Inches(qw_rm - 0.04), Inches(0.25), color)
        text(s, Inches(dash_x + 0.20 + i * qw_rm), Inches(rm_chart_y), Inches(qw_rm - 0.04), Inches(0.25),
             lbl, size=8, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Initiatives bars
    rm_items = [
        ("Genesys script overhaul",  0.00, 0.50, AVANGRID_GREEN, "CS"),
        ("PUC tariff banner Q3",      0.05, 0.30, COS_BLUE,        "ITBA"),
        ("Billing cycle 14→5 days",   0.05, 0.50, AVANGRID_GREEN, "CS"),
        ("Audit framework refresh",   0.40, 0.85, GREY_LABEL,     "ITBA"),
        ("S/4HANA migration plan",    0.65, 0.95, GREY_LABEL,     "ITBA"),
    ]
    items_y = rm_chart_y + 0.30
    items_h_each = 0.28
    for j, (lbl, sx, ex, color, area) in enumerate(rm_items):
        bx = dash_x + 0.20 + sx * rm_chart_w
        bw = (ex - sx) * rm_chart_w
        by = items_y + j * (items_h_each + 0.04)
        round_rect(s, Inches(bx), Inches(by), Inches(bw), Inches(items_h_each), color, radius=0.04)
        text(s, Inches(bx + 0.10), Inches(by + 0.04), Inches(bw - 0.20), Inches(items_h_each - 0.08),
             lbl, size=7, bold=True, color=WHITE, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
        # Area chip on side
        chip(s, dash_x + 0.20 + rm_chart_w + 0.10, by + 0.02, 0.50,
             area, AVANGRID_GREEN if area == "CS" else HEADER_BLUE, WHITE, h=0.20, size=7)

    # Bottom · variance commentary + asks
    var_y = rm_chart_y + rm_chart_h + 0.20
    sw2 = (dash_w - 0.40 - 0.20) / 2
    round_rect(s, Inches(dash_x + 0.20), Inches(var_y), Inches(sw2), Inches(0.90), GREEN_BG, radius=0.04)
    rect(s, Inches(dash_x + 0.20), Inches(var_y), Inches(0.10), Inches(0.90), AVANGRID_GREEN)
    text(s, Inches(dash_x + 0.35), Inches(var_y + 0.06), Inches(sw2 - 0.30), Inches(0.20),
         "VARIANCE COMMENTARY", size=8, bold=True, color=GREEN_DARK, font=BODY_FONT)
    text(s, Inches(dash_x + 0.35), Inches(var_y + 0.26), Inches(sw2 - 0.30), Inches(0.65),
         "▪ CS hit 86% predictability — best quarter on record\n▪ ITBA 3 days slower cycle time due to SAP regression bench formalization\n▪ Capacity Q4 fully booked · early flag from middle managers",
         size=8, color=DARK_TEXT, font=BODY_FONT)

    cx2 = dash_x + 0.20 + sw2 + 0.20
    round_rect(s, Inches(cx2), Inches(var_y), Inches(sw2), Inches(0.90), BLUE_BG, radius=0.04)
    rect(s, Inches(cx2), Inches(var_y), Inches(0.10), Inches(0.90), HEADER_BLUE)
    text(s, Inches(cx2 + 0.15), Inches(var_y + 0.06), Inches(sw2 - 0.30), Inches(0.20),
         "ASKS TO VP", size=8, bold=True, color=HEADER_BLUE, font=BODY_FONT)
    text(s, Inches(cx2 + 0.15), Inches(var_y + 0.26), Inches(sw2 - 0.30), Inches(0.65),
         "▪ Approve hire #5 of new PM for Networks IT VS\n▪ Confirm Audit framework deadline (Q1 27 or earlier?)\n▪ Greenlight S/4HANA Discovery vendor RFP",
         size=8, color=DARK_TEXT, font=BODY_FONT)

    footer_bar(s, 24)
    return s


# ============================================================================
# SECTION 7 · SIZES + TRACKS  (slides 23-25)
# ============================================================================

# 23 · Size dictates path · 3 size flows in parallel
def build_pp23(prs):
    s = _slide(prs)
    slide_title(s,
        "Size dictates the path",
        "S goes direct · M auto-splits to 3 Stories · L is multi-quarter · XL out of scope",
        section="Sizes + Tracks")

    from_phase1_strip(s)

    # 3 horizontal swim lanes for S/M/L
    lane_y_start = 1.65
    lane_h = 1.55
    lane_gap = 0.10

    sizes = [
        ("S", "≤ 5 SP · ~1 week",
         AVANGRID_GREEN,
         [("Story", "S → Story directly · no parent"),
          ("Squad refinement", "Squad PM owns"),
          ("DoR check", "6 criteria"),
          ("Pull", "in current iteration")],
         "BYPASSES Portfolio Board · enters Team Kanban directly"),
        ("M", "5–20 SP · ~1–2 weeks",
         GREEN_MID,
         [("A5 auto-Epic", "M → Epic + child Stories"),
          ("VS refinement", "PM (per VS) + Architect"),
          ("Architect NFR pre-check", "1-pass review"),
          ("DoR per Story", "6 criteria each")],
         "GOES THROUGH Portfolio Board · Gate 2 approval · Squad pulls Stories sequentially"),
        ("L", "20–60 SP · ~2–4 weeks",
         HEADER_BLUE,
         [("Initiative shape", "L → Initiative · default Stories · Epics if > 3 months"),
          ("VS refinement", "PM (per VS) + Architect"),
          ("Full architecture review", "FULL NFR pass"),
          ("Multi-quarter capacity", "Director approves at Gate 2"),
          ("DoR phased", "early Stories first")],
         "GOES THROUGH Portfolio Board · full Gate 2 · Squad pulls phased Stories"),
    ]
    for i, (size, sub, color, steps, bottom_note) in enumerate(sizes):
        ly = lane_y_start + i * (lane_h + lane_gap)
        # Lane background
        round_rect(s, Inches(0.40), Inches(ly), Inches(12.55), Inches(lane_h), GREY_LIGHT, radius=0.04)
        rect(s, Inches(0.40), Inches(ly), Inches(0.10), Inches(lane_h), color)
        # Size badge
        oval(s, Inches(0.65), Inches(ly + lane_h/2 - 0.40), Inches(0.80), Inches(0.80), color)
        text(s, Inches(0.65), Inches(ly + lane_h/2 - 0.40), Inches(0.80), Inches(0.80),
             size, size=28, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, Inches(0.55), Inches(ly + lane_h - 0.32), Inches(1.0), Inches(0.20),
             sub, size=7, italic=True, color=GREY_LABEL, font=BODY_FONT, align=PP_ALIGN.CENTER)

        # Step pipeline
        n_steps = len(steps)
        sx_start = 1.65
        sx_end = 11.20
        avail_w = sx_end - sx_start
        step_w = (avail_w - (n_steps - 1) * 0.10) / n_steps
        sy = ly + 0.30
        for j, (sname, sdesc) in enumerate(steps):
            sx = sx_start + j * (step_w + 0.10)
            round_rect(s, Inches(sx), Inches(sy), Inches(step_w), Inches(0.55), color, radius=0.04)
            text(s, Inches(sx), Inches(sy + 0.05), Inches(step_w), Inches(0.22),
                 sname, size=9, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
            text(s, Inches(sx), Inches(sy + 0.28), Inches(step_w), Inches(0.22),
                 sdesc, size=7, italic=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
            if j < n_steps - 1:
                text(s, Inches(sx + step_w + 0.005), Inches(sy + 0.10), Inches(0.10), Inches(0.40),
                     "→", size=12, bold=True, color=GREY_LABEL, font=BODY_FONT,
                     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Bottom note
        text(s, Inches(1.65), Inches(ly + 0.95), Inches(11.10), Inches(0.20),
             bottom_note, size=8, italic=True, color=GREY_TEXT, font=BODY_FONT)

        # Quick example chips · CS + ITBA
        ex_y = ly + 1.20
        examples_cs = {"S": "CRM page-load fix", "M": "Genesys integration tweak", "L": "CRM platform migration"}
        examples_itba = {"S": "SAP master-data fix", "M": "SAP enhancement", "L": "S/4HANA migration"}
        chip(s, 1.65, ex_y, 0.60, "CS:", AVANGRID_GREEN, WHITE, h=0.22, size=7)
        text(s, Inches(2.30), Inches(ex_y), Inches(3.5), Inches(0.22),
             examples_cs[size], size=8, color=DARK_TEXT, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
        chip(s, 6.80, ex_y, 0.70, "ITBA:", HEADER_BLUE, WHITE, h=0.22, size=7)
        text(s, Inches(7.55), Inches(ex_y), Inches(3.5), Inches(0.22),
             examples_itba[size], size=8, color=DARK_TEXT, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)

    # XL note at bottom
    xl_y = lane_y_start + 3 * (lane_h + lane_gap) + 0.15
    bypass_bar(s, xl_y,
        "XL initiatives → PPM Now (out of scope here) · we accept S · M · L only")

    footer_bar(s, 25)
    return s


# 24 · Dual track · Discovery (vendor/domain) ↔ Delivery (us · orange)
def build_pp24(prs):
    s = _slide(prs)
    slide_title(s,
        "Discovery upstream · Delivery is ours",
        "Phase 0/1 run by domain or vendor · Phase 2/3/4 our orange band · DoR is the bridge",
        section="Sizes + Tracks")

    from_phase1_strip(s)

    # Top lane · Discovery (gray)
    top_y = 1.65
    lane_h = 1.85
    round_rect(s, Inches(0.40), Inches(top_y), Inches(12.55), Inches(lane_h), GREY_LIGHT, radius=0.04)
    rect(s, Inches(0.40), Inches(top_y), Inches(0.20), Inches(lane_h), GREY_LABEL)
    text(s, Inches(0.70), Inches(top_y + 0.12), Inches(5.0), Inches(0.30),
         "DISCOVERY  ·  upstream", size=14, bold=True, color=GREY_LABEL, font=BODY_FONT)
    chip(s, 3.20, top_y + 0.18, 1.60, "DOMAIN / VENDOR — NOT US", GREY_MUTED, WHITE, h=0.24, size=7)
    text(s, Inches(0.70), Inches(top_y + 0.45), Inches(11.0), Inches(0.20),
         "Run by domain product team or vendor · validates problem before our intake form",
         size=9, italic=True, color=GREY_LABEL, font=BODY_FONT)

    # 2 phase boxes · with deliverables explicit
    phase_y = top_y + 0.85
    phase_h = 0.85
    phase_w = (12.55 - 0.50 - 0.30) / 2
    round_rect(s, Inches(0.70), Inches(phase_y), Inches(phase_w), Inches(phase_h), WHITE, radius=0.04)
    text(s, Inches(0.85), Inches(phase_y + 0.06), Inches(phase_w - 0.30), Inches(0.20),
         "PHASE 0  ·  PROBLEM SHAPING", size=10, bold=True, color=GREY_LABEL, font=BODY_FONT)
    text(s, Inches(0.85), Inches(phase_y + 0.28), Inches(phase_w - 0.30), Inches(0.55),
         "DELIVERABLES → Problem statement · users / personas · value hypothesis · risk hypothesis  ·  produced by domain / vendor",
         size=8, color=DARK_TEXT, font=BODY_FONT)
    text(s, Inches(0.70 + phase_w + 0.04), Inches(phase_y + 0.20), Inches(0.22), Inches(0.50),
         "→", size=18, bold=True, color=GREY_LABEL, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    p2_x = 0.70 + phase_w + 0.30
    round_rect(s, Inches(p2_x), Inches(phase_y), Inches(phase_w), Inches(phase_h), WHITE, radius=0.04)
    text(s, Inches(p2_x + 0.15), Inches(phase_y + 0.06), Inches(phase_w - 0.30), Inches(0.20),
         "PHASE 1  ·  MATERIALIZATION", size=10, bold=True, color=GREY_LABEL, font=BODY_FONT)
    text(s, Inches(p2_x + 0.15), Inches(phase_y + 0.28), Inches(phase_w - 0.30), Inches(0.55),
         "DELIVERABLES → Solution shape · vendor model · roadmap · NFRs draft · budget  ·  ready to hand off",
         size=8, color=DARK_TEXT, font=BODY_FONT)

    # DoR bridge between lanes
    db_y = top_y + lane_h + 0.10
    round_rect(s, Inches(4.30), Inches(db_y), Inches(4.70), Inches(0.50), GREEN_DARK, radius=0.20)
    text(s, Inches(4.30), Inches(db_y), Inches(4.70), Inches(0.50),
         "DoR  ·  the bridge between worlds",
         size=11, bold=True, color=WHITE, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(6.50), Inches(db_y - 0.32), Inches(0.30), Inches(0.30),
         "↓", size=18, bold=True, color=GREEN_DARK, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Bottom lane · Delivery (orange · us)
    bot_y = db_y + 0.60
    round_rect(s, Inches(0.40), Inches(bot_y), Inches(12.55), Inches(lane_h), BYPASS_BG, radius=0.04)
    rect(s, Inches(0.40), Inches(bot_y), Inches(0.20), Inches(lane_h), BYPASS_ORANGE)
    text(s, Inches(0.70), Inches(bot_y + 0.12), Inches(5.0), Inches(0.30),
         "DELIVERY  ·  Digital Products squads", size=14, bold=True, color=BYPASS_ORANGE, font=BODY_FONT)
    chip(s, 5.40, bot_y + 0.18, 1.60, "OUR ZONE — ALMOST ALWAYS", BYPASS_ORANGE, WHITE, h=0.24, size=7)
    text(s, Inches(0.70), Inches(bot_y + 0.45), Inches(11.0), Inches(0.20),
         "Phases 2–4 · we own pull · iteration · UAT · release · hypercare",
         size=9, italic=True, color=BYPASS_ORANGE, font=BODY_FONT)

    # 3 delivery phases
    deliv = [
        ("PHASE 2  ·  BUILD",     "squad pulls Stories · iterations · DoD per Story"),
        ("PHASE 3  ·  TEST + UAT", "internal QA + Sponsor UAT · acceptance criteria validated"),
        ("PHASE 4  ·  RELEASE",   "hypercare · feedback loop · Measure & Learn"),
    ]
    pw = (12.55 - 0.50 - 2 * 0.30) / 3
    py = bot_y + 0.85
    ph = 0.85
    for i, (lbl, body) in enumerate(deliv):
        px = 0.70 + i * (pw + 0.30)
        round_rect(s, Inches(px), Inches(py), Inches(pw), Inches(ph), WHITE, radius=0.04)
        text(s, Inches(px + 0.15), Inches(py + 0.06), Inches(pw - 0.30), Inches(0.22),
             lbl, size=10, bold=True, color=BYPASS_ORANGE, font=BODY_FONT)
        text(s, Inches(px + 0.15), Inches(py + 0.30), Inches(pw - 0.30), Inches(0.55),
             body, size=8, color=DARK_TEXT, font=BODY_FONT)
        if i < 2:
            text(s, Inches(px + pw + 0.05), Inches(py + 0.20), Inches(0.20), Inches(0.50),
                 "→", size=18, bold=True, color=BYPASS_ORANGE, font=BODY_FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Bottom · CS vs ITBA Discovery world
    cn_y = bot_y + lane_h + 0.10
    cn_h = 0.85
    sw2 = (12.55 - 0.20) / 2
    round_rect(s, Inches(0.40), Inches(cn_y), Inches(sw2), Inches(cn_h), GREEN_BG, radius=0.04)
    rect(s, Inches(0.40), Inches(cn_y), Inches(0.10), Inches(cn_h), AVANGRID_GREEN)
    text(s, Inches(0.55), Inches(cn_y + 0.06), Inches(sw2 - 0.30), Inches(0.22),
         "CS  ·  vendor-led Discovery (SaaS-heavy)",
         size=10, bold=True, color=GREEN_DARK, font=BODY_FONT)
    text(s, Inches(0.55), Inches(cn_y + 0.30), Inches(sw2 - 0.30), Inches(cn_h - 0.35),
         "Genesys · CIS · Smart Meter platforms — vendor runs Discovery · Heather oversees vendor relationships",
         size=8, color=DARK_TEXT, font=BODY_FONT)
    cx2 = 0.40 + sw2 + 0.20
    round_rect(s, Inches(cx2), Inches(cn_y), Inches(sw2), Inches(cn_h), BLUE_BG, radius=0.04)
    rect(s, Inches(cx2), Inches(cn_y), Inches(0.10), Inches(cn_h), HEADER_BLUE)
    text(s, Inches(cx2 + 0.15), Inches(cn_y + 0.06), Inches(sw2 - 0.30), Inches(0.22),
         "ITBA  ·  domain-led Discovery (SAP-heavy)",
         size=10, bold=True, color=HEADER_BLUE, font=BODY_FONT)
    text(s, Inches(cx2 + 0.15), Inches(cn_y + 0.30), Inches(sw2 - 0.30), Inches(cn_h - 0.35),
         "SAP CoE · vendor portals · custom builds (Networks IT) — domain teams run Discovery · Architect leads NFR alignment",
         size=8, color=DARK_TEXT, font=BODY_FONT)

    footer_bar(s, 26)
    return s


# 25 · Existing-Initiative shortcut · decision tree
def build_pp25(prs):
    s = _slide(prs)
    slide_title(s,
        "Existing-Initiative shortcut",
        "Parent and new ask must be in the same area (CS or ITBA) · bypass Gate 2 only when parent Implementing · Squad PM accepts in 48h",
        section="Sizes + Tracks")

    from_phase1_strip(s)

    # Root question
    root_x = 5.40; root_y = 1.30; root_w = 2.55; root_h = 0.80
    round_rect(s, Inches(root_x), Inches(root_y), Inches(root_w), Inches(root_h), NAVY_DARK, radius=0.10)
    text(s, Inches(root_x), Inches(root_y + 0.10), Inches(root_w), Inches(0.30),
         "WHAT IS THE PARENT", size=10, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
    text(s, Inches(root_x), Inches(root_y + 0.40), Inches(root_w), Inches(0.30),
         "INITIATIVE STATUS?", size=10, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)

    # 5 branches
    statuses = [
        ("FUNNEL",        "REJECT shortcut",   GREY_LABEL, False),
        ("REVIEWING",     "REJECT shortcut",   GREY_LABEL, False),
        ("ANALYZING",     "REJECT shortcut",   GREY_LABEL, False),
        ("BACKLOG",       "ACCEPT (cautious)", COS_ORANGE, True),
        ("IMPLEMENTING",  "ACCEPT (preferred)",AVANGRID_GREEN, True),
    ]
    n = len(statuses)
    avail_w = 12.55
    branch_w = (avail_w - (n - 1) * 0.20) / n
    branch_y = 2.50
    for i, (status, decision, color, is_accept) in enumerate(statuses):
        bx = 0.40 + i * (branch_w + 0.20)
        line(s, Inches(root_x + root_w / 2), Inches(root_y + root_h),
             Inches(bx + branch_w / 2), Inches(branch_y), color, weight=1.5)
        round_rect(s, Inches(bx), Inches(branch_y), Inches(branch_w), Inches(0.55), color, radius=0.04)
        text(s, Inches(bx), Inches(branch_y), Inches(branch_w), Inches(0.55),
             status, size=10, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        round_rect(s, Inches(bx), Inches(branch_y + 0.65), Inches(branch_w), Inches(0.40),
                   WHITE, radius=0.30)
        text(s, Inches(bx), Inches(branch_y + 0.65), Inches(branch_w), Inches(0.40),
             decision, size=8, bold=True, color=color, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Detail row
    detail_y = 3.85
    detail_h = 2.60
    owners = ["PM (per VS)", "PM (per VS)", "PM + Architect", "PM (Squad)", "PM (Squad)"]
    paths = [
        "submit standalone\nnormal flow",
        "submit standalone\nattach reference",
        "standalone · let\nPain Score compete",
        "Story under parent\nre-baseline scope",
        "Story under parent\nDoR fast-track",
    ]
    cos_v = ["—", "—", "—", "inherits parent", "inherits parent"]
    timing = ["next intake cycle", "next intake cycle", "next Gate 2", "≤48h Squad PM", "≤48h Squad PM"]
    for i, (status, decision, color, is_accept) in enumerate(statuses):
        bx = 0.40 + i * (branch_w + 0.20)
        round_rect(s, Inches(bx), Inches(detail_y), Inches(branch_w), Inches(detail_h),
                   GREEN_BG if is_accept else GREY_BG, radius=0.04)
        line(s, Inches(bx + branch_w / 2), Inches(branch_y + 1.05),
             Inches(bx + branch_w / 2), Inches(detail_y), color, weight=1.0)
        text(s, Inches(bx + 0.10), Inches(detail_y + 0.06), Inches(branch_w - 0.20), Inches(0.18),
             "OWNER", size=7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, Inches(bx + 0.10), Inches(detail_y + 0.24), Inches(branch_w - 0.20), Inches(0.22),
             owners[i], size=9, bold=True, color=color, font=BODY_FONT)
        text(s, Inches(bx + 0.10), Inches(detail_y + 0.55), Inches(branch_w - 0.20), Inches(0.18),
             "PATH", size=7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, Inches(bx + 0.10), Inches(detail_y + 0.72), Inches(branch_w - 0.20), Inches(0.65),
             paths[i], size=8, color=DARK_TEXT, font=BODY_FONT)
        text(s, Inches(bx + 0.10), Inches(detail_y + 1.45), Inches(branch_w - 0.20), Inches(0.18),
             "CoS IMPACT", size=7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, Inches(bx + 0.10), Inches(detail_y + 1.62), Inches(branch_w - 0.20), Inches(0.30),
             cos_v[i], size=8, color=DARK_TEXT, font=BODY_FONT)
        text(s, Inches(bx + 0.10), Inches(detail_y + 1.95), Inches(branch_w - 0.20), Inches(0.18),
             "TIMING", size=7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, Inches(bx + 0.10), Inches(detail_y + 2.12), Inches(branch_w - 0.20), Inches(0.30),
             timing[i], size=8, color=DARK_TEXT, font=BODY_FONT)

    # Insight callout at bottom
    insight_callout(s, 0.40, 6.55, 12.55, 0.40,
        "Rule of thumb · only Implementing parent gives the shortcut its full value (DoR fast-track).",
        "")

    footer_bar(s, 27)
    return s


# ============================================================================
# SECTION 8 · CLOSURE  (slides 26-27)
# ============================================================================

# 26 · Day 1 changes per role
def build_pp26(prs):
    s = _slide(prs)
    slide_title(s,
        "Day 1 changes per role",
        "Six roles change · all keep their reporting line · BRS pairing goes live · 90-day ramp",
        section="Closure")

    from_phase1_strip(s)

    # Table-style · 6 roles × 3 columns (NEW · PRESERVED · 90-DAY RAMP)
    tbl_y = 1.65
    tbl_h = 5.55
    label_w = 2.30
    cols_w = (12.55 - label_w) / 3

    # Header row
    rect(s, Inches(0.40), Inches(tbl_y), Inches(label_w), Inches(0.50), NAVY_DARK)
    text(s, Inches(0.40), Inches(tbl_y + 0.10), Inches(label_w), Inches(0.30),
         "ROLE", size=10, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
    headers = [("NEW", AVANGRID_GREEN), ("PRESERVED", HEADER_BLUE), ("90-DAY RAMP", COS_ORANGE)]
    for i, (h, color) in enumerate(headers):
        cx = 0.40 + label_w + i * cols_w
        rect(s, Inches(cx), Inches(tbl_y), Inches(cols_w - 0.04), Inches(0.50), color)
        text(s, Inches(cx), Inches(tbl_y + 0.10), Inches(cols_w - 0.04), Inches(0.30),
             h, size=10, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)

    rows = [
        ("Heather LaJoie\nDirector · Product",
         "Lead Product Management practice · mentor new PMs · BRS pairing · liaison Digital CoE BRP",
         "CCC director duties · existing Sponsor relationships · compensation",
         "PMs onboarded by W12 · first quarterly metrics review with Vitor"),
        ("Kumar Jaladanki\nDirector · CS Eng",
         "Co-vote at Gate 2 · capacity policy with middle managers · BRS counterpart",
         "Direct reports · technical authority · vendor relationships",
         "Middle managers trained on CoS lane discipline · 4 PMs paired"),
        ("Carlos Garbin\nDirector · ITBA",
         "Co-vote at Gate 2 · capacity policy with middle managers · BRS counterpart",
         "Direct reports · technical authority · P&T BRP relationship",
         "Middle managers trained on CoS lane discipline · 5 PMs paired"),
        ("Middle managers\nKumar's + Carlos's lieutenants",
         "Hold day-to-day CoS lane reality · daily ops with PM · escalations to Director",
         "Existing operational responsibilities · team structures · processes",
         "First Gate 2 capacity reality check · CoS adherence baseline set"),
        ("PM (per VS)\nNEW capability",
         "Outcome-named refinement · weekly cadence · BRS pairing · Heather mentorship",
         "—  (NEW role)",
         "Onboarded W12 · independently leading VS by Q+1 · 5 demands scored"),
        ("Architect\ncross-VS",
         "Effort + Safety scoring · NFR pre-check · cross-VS feasibility log",
         "Architecture authority · vendor architecture decisions",
         "1-5 scale calibrated · NFR templates published · feasibility log live"),
    ]
    rh = (tbl_h - 0.50) / len(rows)
    ry = tbl_y + 0.50
    for ri, (role, new, pres, ramp) in enumerate(rows):
        bg = WHITE if ri % 2 == 0 else GREY_LIGHT
        # Role cell
        rect(s, Inches(0.40), Inches(ry), Inches(label_w), Inches(rh - 0.02), bg)
        rect(s, Inches(0.40), Inches(ry), Inches(0.10), Inches(rh - 0.02), AVANGRID_GREEN)
        text(s, Inches(0.55), Inches(ry + 0.10), Inches(label_w - 0.20), Inches(rh - 0.20),
             role, size=8.5, bold=True, color=DARK_TEXT, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
        # 3 cells
        for ci, value in enumerate([new, pres, ramp]):
            cx = 0.40 + label_w + ci * cols_w
            rect(s, Inches(cx), Inches(ry), Inches(cols_w - 0.04), Inches(rh - 0.02), bg)
            text(s, Inches(cx + 0.10), Inches(ry + 0.06), Inches(cols_w - 0.18), Inches(rh - 0.10),
                 value, size=8, color=DARK_TEXT, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
        ry += rh

    footer_bar(s, 28)
    return s


# 27 · Closure · Phase 2 hands off to Phase 3 · what's next
def build_pp27(prs):
    s = _slide(prs)
    slide_title(s,
        "Phase 2 closes · Phase 3 next",
        "DoR-passed Stories hand off to Execution · what we covered · what flows out · 3 risks managed",
        section="Closure")

    from_phase1_strip(s)

    # 3-phase timeline marker (Phase 2 complete · Phase 3 next)
    tl_y = 1.65
    tl_w = 12.55
    rect(s, Inches(0.40), Inches(tl_y + 0.55), Inches(tl_w), Emu(20000), GREY_BORDER)
    phases = [
        ("1", "DEMAND MGMT",  "complete",  GREEN_DARK, True),
        ("2", "PREP & PLANNING", "covered now", AVANGRID_GREEN, True),
        ("3", "EXECUTION",    "next section", HEADER_BLUE, False),
    ]
    pw = (tl_w - 0.40) / 3
    circle_bottom_y = tl_y + 0.95
    label_y = tl_y + 1.05
    sub_y   = tl_y + 1.28
    for i, (num, name, sub, color, done) in enumerate(phases):
        x = 0.40 + 0.10 + i * (pw + 0.10)
        d = 0.85
        cx = x + pw/2 - d/2
        cy = circle_bottom_y - d
        oval(s, Inches(cx), Inches(cy), Inches(d), Inches(d), color)
        text(s, Inches(cx), Inches(cy), Inches(d), Inches(d), num,
             size=24, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, Inches(x), Inches(label_y), Inches(pw), Inches(0.22),
             name, size=11, bold=True, color=color, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, Inches(x), Inches(sub_y), Inches(pw), Inches(0.22),
             sub, size=9, italic=True, color=GREY_LABEL, font=BODY_FONT, align=PP_ALIGN.CENTER)
        if done and i == 1:
            chip(s, x + pw/2 - 0.95, cy - 0.32, 1.90, "▼ JUST COVERED", AVANGRID_GREEN, WHITE, h=0.22, size=8)

    # 3 boxes: WHAT WE COVERED · WHAT FLOWS OUT · 3 RISKS
    box_y = 2.65
    box_h = 4.20
    n = 3
    bw = (12.55 - 2 * 0.20) / n

    # Box 1 · what we covered
    bx = 0.40
    round_rect(s, Inches(bx), Inches(box_y), Inches(bw), Inches(box_h), WHITE, radius=0.04)
    rect(s, Inches(bx), Inches(box_y), Inches(bw), Inches(0.55), AVANGRID_GREEN)
    text(s, Inches(bx + 0.15), Inches(box_y + 0.06), Inches(bw - 0.30), Inches(0.25),
         "WHAT WE COVERED  ·  Phase 2",
         size=11, bold=True, color=WHITE, font=BODY_FONT)
    text(s, Inches(bx + 0.15), Inches(box_y + 0.32), Inches(bw - 0.30), Inches(0.20),
         "27 slides · 7 sections",
         size=8, italic=True, color=GREEN_PALE, font=BODY_FONT)
    cov = [
        ("OPENER",       "Bridge from Phase 1 · 4 gaps · two-board link"),
        ("CAPABILITIES", "6 capabilities · 3 muscles (Detail · Sequence · Capacity)"),
        ("ROLES",        "PM · Architect · Squad PM · Heather mentors · BRS pair"),
        ("TOOLS",        "Backlog · Roadmap · Estimation · Reporting"),
        ("GOVERNANCE",   "Gate 2 · prioritization criteria · Defer/Reject"),
        ("PERFORMANCE",  "7 metrics · cycle time · predictability · monthly review"),
        ("SIZE FLOWS",   "S · M · L paths · dual track · scope-add shortcut"),
    ]
    iy = box_y + 0.65
    rh = (box_h - 0.75) / len(cov)
    for lbl, body in cov:
        text(s, Inches(bx + 0.18), Inches(iy + 0.02), Inches(1.30), Inches(0.18),
             lbl, size=7.5, bold=True, color=GREEN_DARK, font=BODY_FONT)
        text(s, Inches(bx + 1.55), Inches(iy + 0.02), Inches(bw - 1.70), Inches(rh - 0.08),
             body, size=8, color=DARK_TEXT, font=BODY_FONT)
        iy += rh

    # Box 2 · what flows to Phase 3
    bx2 = bx + bw + 0.20
    round_rect(s, Inches(bx2), Inches(box_y), Inches(bw), Inches(box_h), WHITE, radius=0.04)
    rect(s, Inches(bx2), Inches(box_y), Inches(bw), Inches(0.55), HEADER_BLUE)
    text(s, Inches(bx2 + 0.15), Inches(box_y + 0.06), Inches(bw - 0.30), Inches(0.25),
         "WHAT FLOWS TO PHASE 3  ·  Execution input",
         size=11, bold=True, color=WHITE, font=BODY_FONT)
    text(s, Inches(bx2 + 0.15), Inches(box_y + 0.32), Inches(bw - 0.30), Inches(0.20),
         "DoR-passed · ready for Squad pull",
         size=8, italic=True, color=BLUE_PALE, font=BODY_FONT)
    flow = [
        ("Initiative · Epic · Story",   "auto-converted at Triage · INVEST at Story level"),
        ("NFRs documented",    "perf · security · compliance · scale"),
        ("Capacity reserved",  "fits inside CoS lane % · WIP not exceeded"),
        ("Acceptance criteria","traceable to approved Initiative or Epic"),
        ("Dependencies known", "cross-VS · vendor · license — surfaced"),
        ("Roadmap entry",      "committed band · re-balanced monthly"),
        ("Decision log",       "audit trail · Defer/Reject reasons"),
    ]
    iy = box_y + 0.65
    rh = (box_h - 0.75) / len(flow)
    for lbl, body in flow:
        text(s, Inches(bx2 + 0.18), Inches(iy + 0.02), Inches(1.55), Inches(0.18),
             lbl, size=7.5, bold=True, color=HEADER_BLUE, font=BODY_FONT)
        text(s, Inches(bx2 + 1.80), Inches(iy + 0.02), Inches(bw - 1.95), Inches(rh - 0.08),
             body, size=8, color=DARK_TEXT, font=BODY_FONT)
        iy += rh

    # Box 3 · 3 risks
    bx3 = bx2 + bw + 0.20
    round_rect(s, Inches(bx3), Inches(box_y), Inches(bw), Inches(box_h), WHITE, radius=0.04)
    rect(s, Inches(bx3), Inches(box_y), Inches(bw), Inches(0.55), GATE_RED)
    text(s, Inches(bx3 + 0.15), Inches(box_y + 0.06), Inches(bw - 0.30), Inches(0.25),
         "3 RISKS UNDER MANAGEMENT",
         size=11, bold=True, color=WHITE, font=BODY_FONT)
    text(s, Inches(bx3 + 0.15), Inches(box_y + 0.32), Inches(bw - 0.30), Inches(0.20),
         "with named owners + mitigations",
         size=8, italic=True, color=BYPASS_BG, font=BODY_FONT)
    risks = [
        ("PM RAMP-UP",
         "NEW Product capability · 12-week ramp · risk of slow uptake",
         "Pair with senior PM · weekly Heather check-in · BRS pairing accelerates"),
        ("CAPACITY OPACITY",
         "Day-to-day capacity reality may diverge from approved %",
         "Middle managers report at every Gate 2 · variance flagged at >10%"),
        ("BRP DEPENDENCY DRIFT",
         "Inputs from P&T BRP + Digital CoE BRP may shift mid-quarter",
         "Heather + Carlos pre-Gate 2 sync · resequence at next Gate 2 · audit trail"),
    ]
    iy = box_y + 0.70
    rh = (box_h - 0.80) / len(risks)
    for lbl, what, mit in risks:
        text(s, Inches(bx3 + 0.18), Inches(iy), Inches(bw - 0.30), Inches(0.20),
             lbl, size=9, bold=True, color=GATE_RED, font=BODY_FONT)
        text(s, Inches(bx3 + 0.18), Inches(iy + 0.20), Inches(bw - 0.30), Inches(0.40),
             what, size=8, color=DARK_TEXT, font=BODY_FONT)
        text(s, Inches(bx3 + 0.18), Inches(iy + 0.55), Inches(bw - 0.30), Inches(rh - 0.65),
             "MITIGATION · " + mit, size=7.5, italic=True, color=GREY_TEXT, font=BODY_FONT)
        iy += rh

    footer_bar(s, 29)
    return s


# ============================================================================
# APPENDIX · BRP coupling deep-dive (slide 28)
# ============================================================================

# 28 · BRP coupling · how BRP outputs feed Gate 2
def build_brp(prs):
    s = _slide(prs)
    slide_title(s,
        "BRP coupling · Business Review Process",
        "BRP = Business Review Process (NOT Big Room Planning) · existing P&T BRP + Digital CoE BRP feed Gate 2 · we do not run our own",
        section="Governance")

    from_phase1_strip(s)

    # Top band · WHY no own BRP + WHAT BRP is
    bn_y = 1.55
    round_rect(s, Inches(0.40), Inches(bn_y), Inches(12.55), Inches(0.65), GREEN_BG, radius=0.04)
    rect(s, Inches(0.40), Inches(bn_y), Inches(0.10), Inches(0.65), AVANGRID_GREEN)
    text(s, Inches(0.55), Inches(bn_y + 0.06), Inches(12.40), Inches(0.22),
         "WHAT IS A BRP",
         size=10, bold=True, color=GREEN_DARK, font=BODY_FONT)
    text(s, Inches(0.55), Inches(bn_y + 0.30), Inches(12.40), Inches(0.32),
         "BRP = Business Review Process · NOT Big Room Planning · cross-business strategic alignment forum that already exists at Avangrid · Digital Products couples to existing forums (we do NOT run Big Room Planning)",
         size=9, color=DARK_TEXT, font=BODY_FONT)

    # 2 BRP panels · P&T (ITBA-anchored · Carlos+Asier) + Digital CoE (CS-anchored · Heather liaison)
    panel_y = 2.40
    panel_h = 3.30
    sw2 = (12.55 - 0.20) / 2

    panels = [
        ("P&T BRP",
         "Carlos + Asier own the relationship · ITBA-anchored",
         HEADER_BLUE,
         [("WHO ATTENDS",   "Carlos (ITBA Director) · Asier · BRS counterpart · Heather observer"),
          ("CADENCE",       "Quarterly · cross-business strategic forum"),
          ("INPUT TO US",   "ITBA priorities · regulatory deadlines · OPCO directives · capacity asks"),
          ("FORMAT",         "BRP outputs digest (1-page) shared post-meeting"),
          ("FEEDS",         "ITBA backlog re-rank at next Gate 2 · capacity reallocation triggers"),
          ("SYNC OWNER",    "Carlos digests outputs · pre-Gate 2 sync with Heather + Kumar")]),

        ("Digital CoE BRP",
         "Heather is the liaison · spans CS + Digital portfolio",
         AVANGRID_GREEN,
         [("WHO ATTENDS",   "Heather (Director Product) · Digital CoE leadership · CCC stakeholders"),
          ("CADENCE",       "Quarterly · enterprise-wide Digital alignment"),
          ("INPUT TO US",   "Customer-experience priorities · enterprise digital roadmap · Discovery handoffs"),
          ("FORMAT",         "BRP outputs digest (1-page) shared post-meeting"),
          ("FEEDS",         "CS backlog re-rank at next Gate 2 · roadmap commit adjustments"),
          ("SYNC OWNER",    "Heather digests outputs · pre-Gate 2 sync with Carlos + Kumar")]),
    ]
    for i, (name, sub, color, items) in enumerate(panels):
        x = 0.40 + i * (sw2 + 0.20)
        round_rect(s, Inches(x), Inches(panel_y), Inches(sw2), Inches(panel_h), WHITE, radius=0.04)
        rect(s, Inches(x), Inches(panel_y), Inches(sw2), Inches(0.55), color)
        text(s, Inches(x + 0.15), Inches(panel_y + 0.06), Inches(sw2 - 0.30), Inches(0.22),
             name, size=12, bold=True, color=WHITE, font=BODY_FONT)
        text(s, Inches(x + 0.15), Inches(panel_y + 0.30), Inches(sw2 - 0.30), Inches(0.22),
             sub, size=8, italic=True, color=WHITE, font=BODY_FONT)
        iy = panel_y + 0.70
        rh = (panel_h - 0.85) / len(items)
        for lbl, body in items:
            text(s, Inches(x + 0.20), Inches(iy + 0.02), Inches(2.10), Inches(0.20),
                 lbl, size=8, bold=True, color=color, font=BODY_FONT)
            text(s, Inches(x + 2.40), Inches(iy + 0.02), Inches(sw2 - 2.55), Inches(rh - 0.05),
                 body, size=8.5, color=DARK_TEXT, font=BODY_FONT)
            iy += rh

    # Bottom · 4-step pipeline showing BRP → Gate 2 input flow
    flow_y = 5.85
    flow_h = 1.05
    round_rect(s, Inches(0.40), Inches(flow_y), Inches(12.55), Inches(flow_h), GREY_LIGHT, radius=0.04)
    text(s, Inches(0.55), Inches(flow_y + 0.06), Inches(12.30), Inches(0.20),
         "INPUT FLOW  ·  BRP → our Gate 2",
         size=10, bold=True, color=NAVY_DARK, font=BODY_FONT)
    steps = [
        ("BRP meets",      "quarterly · cross-business",       HEADER_BLUE),
        ("Outputs digest", "1-page summary · 5 business days", AVANGRID_GREEN),
        ("Pre-Gate 2 sync","Heather+Carlos+Kumar align",       COS_ORANGE),
        ("Gate 2",         "re-rank backlog · capacity adjust", NAVY_DARK),
    ]
    sw_step = (12.55 - 0.30 - 3 * 0.10) / 4
    sy = flow_y + 0.32
    for i, (lbl, body, color) in enumerate(steps):
        sx = 0.55 + i * (sw_step + 0.10)
        round_rect(s, Inches(sx), Inches(sy), Inches(sw_step), Inches(flow_h - 0.42), color, radius=0.04)
        text(s, Inches(sx), Inches(sy + 0.06), Inches(sw_step), Inches(0.22),
             lbl, size=9, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, Inches(sx), Inches(sy + 0.28), Inches(sw_step), Inches(flow_h - 0.75),
             body, size=8, italic=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
        if i < 3:
            text(s, Inches(sx + sw_step + 0.005), Inches(sy + 0.05), Inches(0.12), Inches(0.40),
                 "→", size=14, bold=True, color=GREY_LABEL, font=BODY_FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    footer_bar(s, 17)
    return s


# 21 · DoR Gate · Definition of Ready · 3 levels
def build_dor(prs):
    s = _slide(prs)
    slide_title(s,
        "DoR Gate · Definition of Ready",
        "3 levels of readiness · Initiative · Epic · Story · INVEST + DoR are stacked checks (not the same)",
        section="Governance")

    from_phase1_strip(s)

    bn_y = 1.55
    round_rect(s, Inches(0.40), Inches(bn_y), Inches(12.55), Inches(0.65), GREEN_BG, radius=0.04)
    rect(s, Inches(0.40), Inches(bn_y), Inches(0.10), Inches(0.65), AVANGRID_GREEN)
    text(s, Inches(0.55), Inches(bn_y + 0.06), Inches(12.40), Inches(0.22),
         "INVEST vs DoR  ·  stacked checks · not the same",
         size=10, bold=True, color=GREEN_DARK, font=BODY_FONT)
    text(s, Inches(0.55), Inches(bn_y + 0.30), Inches(12.40), Inches(0.32),
         "INVEST = Story content quality (during refinement, by PM) · DoR = handoff readiness (before Squad pulls, by Squad PM) · a Story passes INVEST first, then DoR",
         size=10, color=DARK_TEXT, font=BODY_FONT)

    col_y = 2.40
    col_h = 4.10
    cw = (12.55 - 2 * 0.20) / 3

    levels = [
        ("INITIATIVE READINESS",
         "validated at Gate 2 (post-Phase 1)",
         HEADER_BLUE_DK,
         [("Pain Score computed",     "BV · UX · Effort · Safety scored at Analyzing"),
          ("CoS assigned",            "Standard / Fixed-Date / Expedite / Intangible"),
          ("Capacity reserved",       "multi-quarter slot in CoS lane"),
          ("Sponsor named",           "Director-level signoff"),
          ("Roadmap entry",           "rolling-wave commit band"),
          ("Architecture position",   "vendor model · cross-VS feasibility cleared")]),
        ("EPIC READINESS",
         "validated at VS-level refinement (before Story split)",
         HEADER_BLUE,
         [("Outcome named",           "NPS · cycle time · cost · audit pass"),
          ("NFRs identified",         "perf · security · compliance · scale"),
          ("Vendor model decided",    "build vs buy vs SaaS"),
          ("Child Stories outlined",  "list of child Stories (not detailed)"),
          ("Cross-VS deps logged",    "dependencies surfaced and tracked"),
          ("Sponsor signed off",      "Epic-level outcome accepted")]),
        ("STORY DoR  ·  6 criteria",
         "validated by PM (per Squad) before pull (handoff to Execution)",
         GREEN_DARK,
         [("Outcome clear",           "acceptance criteria UAT-testable · binary"),
          ("Sized",                   "≤ 5 SP · INVEST-compliant"),
          ("Architecture reviewed",   "NFRs noted · vendor model confirmed"),
          ("Dependencies known",      "cross-VS · vendor · license surfaced"),
          ("Capacity reserved",       "fits inside CoS lane · WIP not exceeded"),
          ("Acceptance criteria",    "traceable to Sponsor-approved Initiative/Epic · Squad PM verifies · no story-by-story signoff")]),
    ]
    for i, (title_h, sub, color, items) in enumerate(levels):
        x = 0.40 + i * (cw + 0.20)
        round_rect(s, Inches(x), Inches(col_y), Inches(cw), Inches(col_h), WHITE, radius=0.04)
        rect(s, Inches(x), Inches(col_y), Inches(cw), Inches(0.65), color)
        text(s, Inches(x + 0.15), Inches(col_y + 0.06), Inches(cw - 0.30), Inches(0.25),
             title_h, size=11, bold=True, color=WHITE, font=BODY_FONT)
        text(s, Inches(x + 0.15), Inches(col_y + 0.32), Inches(cw - 0.30), Inches(0.30),
             sub, size=8, italic=True, color=WHITE, font=BODY_FONT)
        iy = col_y + 0.80
        rh = (col_h - 0.95) / len(items)
        for lbl, body in items:
            oval(s, Inches(x + 0.18), Inches(iy + 0.06), Inches(0.22), Inches(0.22), color)
            text(s, Inches(x + 0.18), Inches(iy + 0.06), Inches(0.22), Inches(0.22), "✓",
                 size=10, bold=True, color=WHITE, font=BODY_FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            text(s, Inches(x + 0.50), Inches(iy + 0.04), Inches(cw - 0.65), Inches(0.20),
                 lbl, size=8.5, bold=True, color=DARK_TEXT, font=BODY_FONT)
            text(s, Inches(x + 0.50), Inches(iy + 0.22), Inches(cw - 0.65), Inches(rh - 0.25),
                 body, size=7.5, color=GREY_TEXT, font=BODY_FONT)
            iy += rh

    insight_callout(s, 0.40, 6.65, 12.55, 0.40,
        "Sponsor signs off at INITIATIVE and EPIC level — NOT story-by-story · Story DoR inherits acceptance criteria from approved parent.",
        "")

    footer_bar(s, 21)
    return s


# ============================================================================
# Driver (test mode · slides 1-3 only for now)
# ============================================================================
def main():
    print("Opening:", SRC)
    prs = Presentation(SRC)
    n_before = len(prs.slides)

    print("Building Preparation & Planning v4 · 27 slides · all sections...")
    build_pp01(prs);  print("   01 · Bridge from Phase 1")
    build_pp02(prs);  print("   02 · Governing thought · 4 gaps")
    build_pp03(prs);  print("   03 · Two boards · Portfolio ↔ Team")
    build_pp04(prs);  print("   04 · Capabilities map · 3 muscles")
    build_pp05(prs);  print("   05 · Refinement · hierarchy + INVEST")
    build_pp06(prs);  print("   06 · Estimation · Pain Score CS+ITBA")
    build_pp07(prs);  print("   07 · Prioritization · 4 layers")
    build_pp08(prs);  print("   08 · Portfolio backlog · re-ranking")
    build_pp09(prs);  print("   09 · Roadmap · rolling wave")
    build_pp10(prs);  print("   10 · Capacity allocation matrix")
    build_pp11(prs);  print("   11 · Unplanned demand · 3 channels")
    build_pp12(prs);  print("   12 · Roles overview · 3 roles + Heather")
    build_pp13(prs);  print("   13 · PM deep-dive + BRS pairing")
    build_pp14(prs);  print("   14 · RACI matrix")
    build_pp15(prs);  print("   15 · Tools · 4 families + dashboard")
    build_pp16(prs);  print("   16 · Gate 2 · the only forum")
    build_brp(prs);   print("   17 · BRP coupling · inputs to Gate 2")
    build_pp17(prs);  print("   18 · Prioritization criteria")
    build_pp18(prs);  print("   19 · Demand × Capacity oversight")
    build_pp19(prs);  print("   20 · Defer / Reject decision practices")
    build_dor(prs);   print("   21 · DoR Gate · Definition of Ready (3 levels)")
    build_pp20(prs);  print("   22 · Performance metrics overview")
    build_pp21(prs);  print("   23 · Cycle time + predictability charts")
    build_pp22(prs);  print("   24 · Heather's monthly report to VP")
    build_pp23(prs);  print("   25 · Size paths · S/M/L")
    build_pp24(prs);  print("   26 · Dual track · Discovery ↔ Delivery")
    build_pp25(prs);  print("   27 · Existing-Initiative shortcut")
    build_pp26(prs);  print("   28 · Day 1 changes per role")
    build_pp27(prs);  print("   29 · Closure + Phase 3 bridge")

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
    print(f"\n✓ Saved: {DST}")


if __name__ == "__main__":
    main()

