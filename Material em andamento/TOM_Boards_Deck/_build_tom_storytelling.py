"""
TOM MAP storytelling adapter.

Changes applied to TOM MAP workbook-1.pptx → saved as workbook-2.pptx:

  A. Slide 32 — fix 3 inconsistencies + rename PM steward → Demand Manager
  B. Slides 21–28 — replace MANDATORY badge with stage/owner badge
  C. NEW slide  FIELD PROGRESSION         (inserted before current slide 21)
  D. NEW slide  AUTOMATION CATALOG (A1–A6)(inserted after current slide 32)
  E. NEW slide  TEAM KANBAN DEEP DIVE     (inserted after current slide 35)
  F. NEW slide  GATES DEEP DIVE           (inserted after current slide 39)
  G. NEW slide  BRP DEEP DIVE             (inserted after Gates Deep Dive)
  H. NEW slide  VALUE & PRIORITIZATION    (inserted before current slide 41)

Visual identity:
  - Dark blue header bar (#2C578B) with white IberPangea title — matches slides 17/39/41
  - Avangrid green accents (#008C38 · #007A32 · #00B176)
  - NO stroke on filled colored boxes (border-overflow fix)
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

SRC = "/Users/gustavohenriquecastellano/Downloads/Base Apresentações Projeto Digital Product/TOM MAP workbook-1.pptx"
DST = "/Users/gustavohenriquecastellano/Downloads/Base Apresentações Projeto Digital Product/TOM MAP workbook-2.pptx"

# ============================================================================
# Palette — matches existing deck (dark navy + Avangrid green)
# ============================================================================
HEADER_BLUE      = RGBColor(0x2C, 0x57, 0x8B)
HEADER_BLUE_DK   = RGBColor(0x1E, 0x3A, 0x5F)
HEADER_BLUE_LT   = RGBColor(0x5B, 0x8A, 0xB5)

AVANGRID_GREEN   = RGBColor(0x00, 0x8C, 0x38)
GREEN_CARD       = RGBColor(0x00, 0x7A, 0x32)
GREEN_FRESH      = RGBColor(0x00, 0xB1, 0x76)
GREEN_DARK       = RGBColor(0x00, 0x6B, 0x3F)
GREEN_MID        = RGBColor(0x00, 0x89, 0x50)
GREEN_LIGHT      = RGBColor(0x00, 0x9F, 0x5D)
GREEN_PASTEL     = RGBColor(0xC8, 0xE0, 0xCB)
GREEN_BG         = RGBColor(0xF2, 0xF8, 0xF4)

WHITE            = RGBColor(0xFF, 0xFF, 0xFF)
DARK_TEXT        = RGBColor(0x2C, 0x3E, 0x50)
GREY_LABEL       = RGBColor(0x6B, 0x77, 0x8C)
GREY_BG          = RGBColor(0xF4, 0xF5, 0xF7)

TITLE_FONT = "IberPangea Text Medium"
BODY_FONT  = "Arial"

# ============================================================================
# Robust shape helpers — NO stroke on filled boxes (border-overflow fix)
# ============================================================================

def _no_line(shape):
    shape.line.fill.background()


def rect(slide, x, y, w, h, fill=None):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    s.shadow.inherit = False
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    _no_line(s)
    return s


def round_rect(slide, x, y, w, h, fill=None, radius=0.04):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    s.shadow.inherit = False
    s.adjustments[0] = radius
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    _no_line(s)
    return s


def text(slide, x, y, w, h, txt, size=10, bold=False, italic=False,
         color=DARK_TEXT, font=BODY_FONT,
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
    r.font.name = font
    r.font.color.rgb = color
    return tb


def header_bar(slide, w_total, line1, line2):
    """Dark blue header bar matching slides 17/39/41."""
    rect(slide, 0, 0, w_total, Inches(0.95), HEADER_BLUE)
    rect(slide, 0, Inches(0.95), w_total, Emu(15000), GREEN_FRESH)  # green divider line
    text(slide, Inches(0.40), Inches(0.10), Inches(11), Inches(0.42),
         line1, size=22, bold=True, color=WHITE, font=BODY_FONT,
         anchor=MSO_ANCHOR.MIDDLE)
    text(slide, Inches(0.40), Inches(0.55), Inches(11), Inches(0.32),
         line2, size=10, bold=True, color=WHITE, font=BODY_FONT)
    text(slide, Inches(11.4), Inches(0.30), Inches(1.7), Inches(0.4),
         "NTT  DATA", size=14, bold=True, color=WHITE,
         font=BODY_FONT, align=PP_ALIGN.RIGHT)


def footer(slide, label):
    text(slide, Inches(0.40), Inches(7.18), Inches(8), Inches(0.18),
         "Avangrid · Digital Products · Target Operating Model",
         size=8, italic=True, color=GREY_LABEL, font=BODY_FONT)
    text(slide, Inches(11.5), Inches(7.18), Inches(1.6), Inches(0.18),
         label, size=8, color=GREY_LABEL, font=BODY_FONT,
         align=PP_ALIGN.RIGHT)


# ============================================================================
# Slide reordering helper
# ============================================================================

def move_slide(prs, old_index, new_index):
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    s = slides[old_index]
    xml_slides.remove(s)
    xml_slides.insert(new_index, s)


# ============================================================================
# C — FIELD PROGRESSION (overview before field deep-dives)
# ============================================================================

def build_field_progression(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_bar(s, prs.slide_width,
               "FIELD PROGRESSION",
               "WHO FILLS WHAT · WHEN · the card grows as it advances through the funnel")

    col_y = Inches(1.20)
    col_h = Inches(5.45)
    col_w = Inches(3.05)
    col_gap = Inches(0.10)

    stages = [
        ("FUNNEL · SUBMIT",
         "Requester (Sponsor)",
         "at intake · Jira form",
         GREEN_DARK,
         [
            ("BUSINESS CONTEXT",
             "Title · Description · Business Value (narrative) · Acceptance Criteria"),
            ("ROUTING",
             "Value Stream · OPCO · Sponsor"),
            ("FLAGS",
             "Regulatory flag (Yes/No) · Customer Impact Tier (P1–P4 if known)"),
         ]),
        ("REVIEWING",
         "Demand Manager + Sponsor",
         "SLA: 2 weeks · post-Gate 1",
         GREEN_MID,
         [
            ("CLASSIFICATION",
             "Type (Demand · Incident · Change)"),
            ("SIZE · INITIAL",
             "T-shirt: S · M · L · XL — best-guess"),
            ("CoS · INITIAL",
             "Standard / Fixed-Date / Expedite / Intangible — best-guess"),
            ("LENSES",
             "Work Type (BAU / Build / Upgrade)  ·  Run / Grow / Transform"),
         ]),
        ("ANALYZING",
         "Demand Manager + Architect",
         "Pain Score auto-fires (A4)",
         GREEN_LIGHT,
         [
            ("SCORES (1–5)",
             "BV · UX · Safety · Effort — set at Analyzing"),
            ("PAIN SCORE · AUTO",
             "H(x) = (BV·w₁ + UX·w₂ + Safety·w₃) / Effort"),
            ("FINAL SIZE / CoS",
             "Locked after Architect review"),
            ("ARTIFACTS",
             "Discovery Report · Concept Prototype · Feasibility check"),
         ]),
        ("GATE 2 · PORTFOLIO",
         "Director (Carlos · Kumar)",
         "Monthly · 90 min",
         GREEN_CARD,
         [
            ("DECISION",
             "Go / No-Go / Defer  ·  judgment on the assembled card"),
            ("CAPACITY",
             "Slot to Value Stream · quarter assignment"),
            ("INPUTS REVIEWED",
             "Pain Score · Discovery Report · Feasibility · capacity"),
            ("OUTPUTS",
             "Backlog rank · status = Ready For Execution"),
         ]),
    ]

    for i, (label, who, when, fill_color, blocks) in enumerate(stages):
        cx = Inches(0.40) + i * (col_w + col_gap)
        round_rect(s, cx, col_y, col_w, col_h, WHITE, radius=0.03)
        round_rect(s, cx, col_y, col_w, Inches(1.05), fill_color, radius=0.03)

        text(s, cx + Inches(0.15), col_y + Inches(0.10), col_w - Inches(0.30), Inches(0.28),
             label, size=12, bold=True, color=WHITE, font=BODY_FONT)
        text(s, cx + Inches(0.15), col_y + Inches(0.40), col_w - Inches(0.30), Inches(0.20),
             "OWNER", size=7, bold=True, color=GREEN_PASTEL, font=BODY_FONT)
        text(s, cx + Inches(0.15), col_y + Inches(0.55), col_w - Inches(0.30), Inches(0.22),
             who, size=10, bold=True, color=WHITE, font=BODY_FONT)
        text(s, cx + Inches(0.15), col_y + Inches(0.78), col_w - Inches(0.30), Inches(0.22),
             when, size=8, italic=True, color=GREEN_PASTEL, font=BODY_FONT)

        by = col_y + Inches(1.20)
        text(s, cx + Inches(0.15), by, col_w - Inches(0.30), Inches(0.22),
             "FIELDS ADDED HERE", size=8, bold=True, color=GREY_LABEL, font=BODY_FONT)
        by += Inches(0.28)
        for label_b, content in blocks:
            block_h = Inches(0.92)
            round_rect(s, cx + Inches(0.15), by, col_w - Inches(0.30), block_h, GREEN_BG, radius=0.05)
            text(s, cx + Inches(0.25), by + Inches(0.08), col_w - Inches(0.50), Inches(0.22),
                 label_b, size=8, bold=True, color=GREEN_DARK, font=BODY_FONT)
            text(s, cx + Inches(0.25), by + Inches(0.30), col_w - Inches(0.50), Inches(0.65),
                 content, size=9, color=DARK_TEXT, font=BODY_FONT)
            by += block_h + Inches(0.05)

    note_y = Inches(6.75)
    round_rect(s, Inches(0.40), note_y, Inches(12.55), Inches(0.36), GREEN_BG, radius=0.04)
    text(s, Inches(0.55), note_y, Inches(12.2), Inches(0.36),
         "WHY this design  ·  Single funnel ends siloed intake → one source of truth for OVS-level KPIs (NPS · CSAT · Reliability · cycle time) · "
         "Architect adds feasibility at Analyzing · Director judges the assembled card at Gate 2.",
         size=9, italic=True, color=DARK_TEXT, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)

    footer(s, "Field Progression")
    return s


# ============================================================================
# D — AUTOMATION CATALOG (A1–A6)
# ============================================================================

def build_automation_catalog(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_bar(s, prs.slide_width,
               "AUTOMATION CATALOG  ·  A1 → A6",
               "Six Jira automations — trigger · action · field affected · owner")

    rows = [
        ("A1", "Auto-Routing",
         "Demand submitted via Jira intake form",
         "Reads Value Stream + OPCO → routes ticket to correct Portfolio Board · @notifies Champion",
         "Value Stream · OPCO · Status (NEW → ROUTED)",
         "Demand Manager"),
        ("A2", "Auto-Expedite",
         "Customer Impact Tier = P1 (or P2 critical)",
         "Sets CoS = Expedite · forces WIP-1 on squad · emails Director + Sponsor",
         "Class of Service · WIP lane",
         "Demand Manager"),
        ("A3", "Auto-Fixed-Date",
         "Regulatory flag = Yes",
         "Sets CoS = Fixed-Date · forces a due date · pre-flags Gate 2 priority",
         "Class of Service · Due Date",
         "Demand Manager"),
        ("A4", "Auto-Pain-Score",
         "BV · UX · Safety · Effort scores filled at Analyzing",
         "Computes  H(x) = (BV·w₁ + UX·w₂ + Safety·w₃) / Effort  ·  writes Pain Score field",
         "Pain Score (read-only · system field)",
         "Architect"),
        ("A5", "Auto-Convert",
         "At Triage (Gate 1) when Demand has Size + Type set",
         "Converts Demand placeholder:  S → Story  ·  M/L → Epic  ·  XL → Initiative  ·  Incident → Bug",
         "Issue Type · Parent Link · Status (TRIAGED)",
         "Demand Manager"),
        ("A6", "Aging Alert",
         "Daily — flags any card stuck past CoS-defined SLA",
         "@notifies owner · adds AGING tag · raises a flow blocker on squad board",
         "Aging tag · Flow Blocker",
         "Squad Lead"),
    ]

    table_y = Inches(1.20)
    row_h = Inches(0.78)
    col_x = [Inches(0.40), Inches(1.20), Inches(2.45), Inches(4.95), Inches(8.25), Inches(11.45)]
    col_w_list = [Inches(0.80), Inches(1.20), Inches(2.45), Inches(3.25), Inches(3.15), Inches(1.50)]
    headers = ["ID", "NAME", "TRIGGER", "ACTION", "FIELDS AFFECTED", "OWNER"]

    rect(s, col_x[0], table_y, Inches(13.05), Inches(0.42), HEADER_BLUE)
    for i, hdr in enumerate(headers):
        text(s, col_x[i] + Inches(0.10), table_y + Inches(0.04),
             col_w_list[i] - Inches(0.20), Inches(0.34),
             hdr, size=9, bold=True, color=WHITE, font=BODY_FONT,
             anchor=MSO_ANCHOR.MIDDLE)

    by = table_y + Inches(0.42)
    for i, (aid, name, trigger, action, fields, owner) in enumerate(rows):
        row_fill = GREEN_BG if i % 2 == 0 else WHITE
        rect(s, col_x[0], by, Inches(13.05), row_h, row_fill)
        round_rect(s, col_x[0] + Inches(0.10), by + Inches(0.18),
                   Inches(0.55), Inches(0.42), GREEN_DARK, radius=0.30)
        text(s, col_x[0] + Inches(0.10), by + Inches(0.18),
             Inches(0.55), Inches(0.42),
             aid, size=11, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, col_x[1] + Inches(0.05), by + Inches(0.10),
             col_w_list[1] - Inches(0.10), row_h - Inches(0.20),
             name, size=10, bold=True, color=GREEN_DARK, font=BODY_FONT,
             anchor=MSO_ANCHOR.MIDDLE)
        text(s, col_x[2] + Inches(0.05), by + Inches(0.05),
             col_w_list[2] - Inches(0.10), row_h - Inches(0.10),
             trigger, size=8.5, color=DARK_TEXT, font=BODY_FONT,
             anchor=MSO_ANCHOR.MIDDLE)
        text(s, col_x[3] + Inches(0.05), by + Inches(0.05),
             col_w_list[3] - Inches(0.10), row_h - Inches(0.10),
             action, size=8.5, color=DARK_TEXT, font=BODY_FONT,
             anchor=MSO_ANCHOR.MIDDLE)
        text(s, col_x[4] + Inches(0.05), by + Inches(0.05),
             col_w_list[4] - Inches(0.10), row_h - Inches(0.10),
             fields, size=8.5, italic=True, color=GREEN_DARK, font=BODY_FONT,
             anchor=MSO_ANCHOR.MIDDLE)
        text(s, col_x[5] + Inches(0.05), by + Inches(0.05),
             col_w_list[5] - Inches(0.10), row_h - Inches(0.10),
             owner, size=9, bold=True, color=DARK_TEXT, font=BODY_FONT,
             anchor=MSO_ANCHOR.MIDDLE)
        by += row_h

    note_y = Inches(6.55)
    round_rect(s, Inches(0.40), note_y, Inches(12.55), Inches(0.55), GREEN_BG, radius=0.04)
    text(s, Inches(0.55), note_y + Inches(0.06), Inches(12.2), Inches(0.18),
         "READ THIS WITH THE AFTER-SUBMISSION SLIDE",
         size=8, bold=True, color=GREEN_DARK, font=BODY_FONT)
    text(s, Inches(0.55), note_y + Inches(0.24), Inches(12.2), Inches(0.30),
         "A1 fires at submit · A2/A3 fire when intake flags trip · A4 + A5 fire at Triage · A6 runs daily across the entire portfolio. "
         "Together they shorten ideation-to-delivery cycle by ~40 % and cut handoff delays.",
         size=9, color=DARK_TEXT, font=BODY_FONT)

    footer(s, "Automation Catalog")
    return s


# ============================================================================
# E — TEAM KANBAN DEEP DIVE
# ============================================================================

def build_team_kanban_deepdive(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_bar(s, prs.slide_width,
               "TEAM KANBAN  ·  PHASE DEEP DIVE",
               "Owner · output · cadence per phase  ·  the Discovery↔Delivery handoff rule")

    phase_y = Inches(1.20)
    phase_h = Inches(4.95)
    phase_w = Inches(2.45)
    phase_gap = Inches(0.07)

    phases = [
        ("Phase 0", "DISCOVERY",
         "Problem framing · users · hypothesis",
         "Always-on · supports Materialization",
         "Product Manager",
         "Validated user problem",
         "continuous · weekly review",
         GREEN_DARK),
        ("Phase 1", "MATERIALIZATION",
         "Definition · solutioning · sizing",
         "Output: Roadmap · Budget · Sizing",
         "Product Manager + Tech Lead",
         "Roadmap · Budget · Sizing",
         "2–4 weeks per item",
         GREEN_MID),
        ("Phase 2", "CREATION",
         "Planning · team / vendor assignment",
         "Output: Outcome · KPIs · Strategic alignment",
         "Product Manager",
         "Outcome definition · KPIs",
         "1 sprint · pre-build kick-off",
         GREEN_LIGHT),
        ("Phase 3", "ONGOING",
         "Build · test · deploy iteratively",
         "Kanban flow · daily / replenishment / service review · WIP per stage",
         "Squad (PM + Tech Lead + BA + dev/QA)",
         "Released increments · cycle-time captured",
         "continuous · daily flow review",
         GREEN_CARD),
        ("Phase 4", "MAINTENANCE",
         "Hypercare · stabilization · knowledge",
         "Handoff to BAU within 30 days post-launch",
         "Operations + Tech Lead",
         "Stable production state",
         "30-day hypercare window",
         GREEN_FRESH),
    ]

    for i, (phase_num, phase_name, desc, sub, owner, output, cadence, fill) in enumerate(phases):
        px = Inches(0.40) + i * (phase_w + phase_gap)
        round_rect(s, px, phase_y, phase_w, phase_h, WHITE, radius=0.03)
        round_rect(s, px, phase_y, phase_w, Inches(0.78), fill, radius=0.03)
        text(s, px, phase_y + Inches(0.06), phase_w, Inches(0.22),
             phase_num, size=9, bold=True, color=GREEN_PASTEL, font=BODY_FONT,
             align=PP_ALIGN.CENTER)
        text(s, px, phase_y + Inches(0.30), phase_w, Inches(0.36),
             phase_name, size=14, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        by = phase_y + Inches(0.92)
        text(s, px + Inches(0.15), by, phase_w - Inches(0.30), Inches(0.40),
             desc, size=10, bold=True, color=DARK_TEXT, font=BODY_FONT)
        text(s, px + Inches(0.15), by + Inches(0.42), phase_w - Inches(0.30), Inches(0.60),
             sub, size=8, italic=True, color=GREEN_DARK, font=BODY_FONT)

        text(s, px + Inches(0.15), by + Inches(1.10), phase_w - Inches(0.30), Inches(0.18),
             "OWNER", size=7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, px + Inches(0.15), by + Inches(1.27), phase_w - Inches(0.30), Inches(0.50),
             owner, size=9, color=DARK_TEXT, font=BODY_FONT)

        text(s, px + Inches(0.15), by + Inches(1.85), phase_w - Inches(0.30), Inches(0.18),
             "OUTPUT", size=7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, px + Inches(0.15), by + Inches(2.02), phase_w - Inches(0.30), Inches(0.50),
             output, size=9, color=DARK_TEXT, font=BODY_FONT)

        text(s, px + Inches(0.15), by + Inches(2.60), phase_w - Inches(0.30), Inches(0.18),
             "CADENCE", size=7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, px + Inches(0.15), by + Inches(2.77), phase_w - Inches(0.30), Inches(0.50),
             cadence, size=9, italic=True, color=GREEN_DARK, font=BODY_FONT)

    hf_y = Inches(6.30)
    round_rect(s, Inches(0.40), hf_y, Inches(12.55), Inches(0.78), GREEN_BG, radius=0.04)
    text(s, Inches(0.55), hf_y + Inches(0.06), Inches(12.2), Inches(0.22),
         "DISCOVERY  ↔  DELIVERY HANDOFF  ·  the gatekeeping rule that protects flow",
         size=10, bold=True, color=GREEN_DARK, font=BODY_FONT)
    text(s, Inches(0.55), hf_y + Inches(0.30), Inches(8.5), Inches(0.42),
         "Items move from Phase 0/1 (Discovery) to Phase 2/3 (Delivery) ONLY when DoR is met. "
         "Discovery feeds Delivery continuously · WIP limits never break.",
         size=9, color=DARK_TEXT, font=BODY_FONT)
    text(s, Inches(9.30), hf_y + Inches(0.30), Inches(3.6), Inches(0.42),
         "GATEKEEPER  ·  Product Manager",
         size=10, bold=True, italic=True, color=GREEN_DARK, font=BODY_FONT,
         align=PP_ALIGN.RIGHT)

    footer(s, "Team Kanban · Deep Dive")
    return s


# ============================================================================
# F — GATES DEEP DIVE
# ============================================================================

def build_gates_deepdive(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_bar(s, prs.slide_width,
               "GATES DEEP DIVE",
               "GATE 1 protects QUALITY at intake · GATE 2 protects INVESTMENT at portfolio")

    gate_y = Inches(1.20)
    gate_h = Inches(5.85)
    gate_w = Inches(6.18)
    gate_gap = Inches(0.18)

    gates = [
        ("GATE 1  ·  TRIAGE",
         "Quality gate · entry validation · DoR check",
         "WEEKLY  ·  30 min",
         GREEN_MID,
         [
            ("PURPOSE",
             "Confirms incoming Demand has enough business context to be worked on. "
             "Filters noise · prevents wasted analysis effort downstream."),
            ("WHO PARTICIPATES",
             "▸ CHAIR · Demand Manager (rotating role)\n"
             "▸ ATTENDEES · Champions (one per VS) · Architect (cross-VS standards)\n"
             "▸ OBSERVERS · PM stewards (Heather + P&T pool) · Sponsor on request"),
            ("WHEN",
             "Every Wednesday 09:00–09:30 ET · standing meeting · cancelled only if zero new Demands"),
            ("HOW · STEP-BY-STEP",
             "1. Demand Manager pulls all NEW + ROUTED tickets from prior 7 days\n"
             "2. For each: validate DoR (Title · Description · BV narrative · AC · VS · OPCO · Sponsor)\n"
             "3. Decision  ·  ACCEPT (→ Reviewing)  ·  BOUNCE (→ Requestor with reason)  ·  MERGE\n"
             "4. Auto-routing recorded in Jira · @notify Champion of accepted items\n"
             "5. KPI captured  ·  DoR pass rate · bounce rate · channel adoption"),
            ("INPUTS  /  OUTPUTS",
             "INPUTS  ·  NEW + ROUTED tickets · DoR checklist · Value Stream policies\n"
             "OUTPUTS  ·  TRIAGED · BOUNCED · merge links · weekly KPI report"),
         ]),
        ("GATE 2  ·  PORTFOLIO COMMITTEE",
         "Investment gate · Go/No-Go · capacity slot",
         "MONTHLY  ·  90 min",
         GREEN_DARK,
         [
            ("PURPOSE",
             "Approves investment · slots capacity · re-ranks portfolio backlog by Pain Score + CoS. "
             "The gate where Initiatives become committed work."),
            ("WHO PARTICIPATES",
             "▸ CHAIR · Vitor Cruz (VP IT Business Solutions Delivery)\n"
             "▸ DECIDERS · Carlos Garbin (ITBA Director) · Kumar Jaladanki (CS Director)\n"
             "▸ ATTENDEES · PM stewards · Architect · Travis (PMO Lead · for L/XL only)\n"
             "▸ OBSERVERS · Sponsors on relevant items only"),
            ("WHEN",
             "First Tuesday of the month · 13:00–14:30 ET · pre-read circulated 3 business days ahead"),
            ("HOW · STEP-BY-STEP",
             "1. Pre-read  ·  all Initiatives in ANALYZING with Pain Score + Discovery Report attached\n"
             "2. Walk-through by Pain Score (highest first) · Architect comments on feasibility\n"
             "3. Decision per Initiative  ·  GO  ·  NO-GO  ·  DEFER\n"
             "4. Backlog re-ranked by CoS + Pain Score · capacity heat-map updated\n"
             "5. Decisions logged in Jira · @notify Sponsors + Squad Leads"),
            ("INPUTS  /  OUTPUTS",
             "INPUTS  ·  TRIAGED Initiatives · Pain Scores · Discovery Reports · capacity per VS · vendor pipeline\n"
             "OUTPUTS  ·  Ready For Execution · capacity slot · Backlog rank · OVS KPI panel updated (NPS · CSAT · Reliability · cycle time)"),
         ]),
    ]

    for i, (label, sub, cadence, fill, blocks) in enumerate(gates):
        gx = Inches(0.40) + i * (gate_w + gate_gap)
        round_rect(s, gx, gate_y, gate_w, gate_h, WHITE, radius=0.03)
        round_rect(s, gx, gate_y, gate_w, Inches(0.85), fill, radius=0.03)
        text(s, gx + Inches(0.20), gate_y + Inches(0.10), gate_w - Inches(0.40), Inches(0.32),
             label, size=16, bold=True, color=WHITE, font=BODY_FONT)
        text(s, gx + Inches(0.20), gate_y + Inches(0.42), gate_w - Inches(0.40), Inches(0.22),
             sub, size=10, italic=True, color=GREEN_PASTEL, font=BODY_FONT)
        text(s, gx + Inches(0.20), gate_y + Inches(0.62), gate_w - Inches(0.40), Inches(0.22),
             "▸ " + cadence, size=10, bold=True, color=WHITE, font=BODY_FONT)

        by = gate_y + Inches(1.00)
        for label_b, content in blocks:
            block_h = Inches(0.96)
            round_rect(s, gx + Inches(0.15), by, gate_w - Inches(0.30), block_h, GREEN_BG, radius=0.04)
            text(s, gx + Inches(0.25), by + Inches(0.06), gate_w - Inches(0.50), Inches(0.20),
                 label_b, size=8, bold=True, color=GREEN_DARK, font=BODY_FONT)
            text(s, gx + Inches(0.25), by + Inches(0.24), gate_w - Inches(0.50), block_h - Inches(0.28),
                 content, size=8, color=DARK_TEXT, font=BODY_FONT)
            by += block_h + Inches(0.04)

    footer(s, "Gates · Deep Dive")
    return s


# ============================================================================
# G — BRP DEEP DIVE
# ============================================================================

def build_brp_deepdive(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_bar(s, prs.slide_width,
               "BIG ROOM PLANNING  ·  DEEP DIVE",
               "Quarterly cross-portfolio alignment · 4 acts · half-day · the cadence that surfaces dependencies")

    # Top metadata strip — WHAT/WHY/WHEN/WHO
    meta_y = Inches(1.15)
    meta_h = Inches(1.20)
    cells = [
        ("WHAT",  "Quarterly half-day session aligning all Value Streams: lock OKRs · "
                  "map cross-VS dependencies · commit capacity per quarter."),
        ("WHY",   "Surfaces dependencies BEFORE they become blockers · aligns portfolio with "
                  "corporate OKRs · stops 'reg always wins' bypass."),
        ("WHEN",  "First week of each quarter · 13:00–17:00 ET · 4 hours · "
                  "in-person preferred · pre-read circulated 5 business days ahead."),
        ("WHO",   "CHAIR  ·  VP Vitor Cruz   ·   DECIDERS  ·  Carlos Garbin · Kumar Jaladanki   ·   "
                  "CHAMPIONS  ·  1 per VS   ·   SPONSORS  ·  business CS + ITBA   ·   "
                  "Iberdrola liaison · Architect · Travis (PMO Lead)"),
    ]
    cell_w = (Inches(12.55) - 3 * Inches(0.10)) / 4
    for i, (label, body) in enumerate(cells):
        cx = Inches(0.40) + i * (cell_w + Inches(0.10))
        round_rect(s, cx, meta_y, cell_w, meta_h, WHITE, radius=0.04)
        round_rect(s, cx, meta_y, cell_w, Inches(0.30), HEADER_BLUE, radius=0.04)
        text(s, cx + Inches(0.10), meta_y, cell_w - Inches(0.20), Inches(0.30),
             label, size=10, bold=True, color=WHITE, font=BODY_FONT,
             anchor=MSO_ANCHOR.MIDDLE)
        text(s, cx + Inches(0.12), meta_y + Inches(0.36), cell_w - Inches(0.24), Inches(0.78),
             body, size=8.5, color=DARK_TEXT, font=BODY_FONT)

    # 4 Acts — main flow
    acts_y = Inches(2.55)
    acts_h = Inches(3.30)
    acts_w = (Inches(12.55) - 3 * Inches(0.10)) / 4

    acts = [
        ("ACT 1", "CONFIRM WHY", "30 min",
         "VP reviews corporate OKRs and quarterly direction.",
         ["Corporate OKRs cascade",
          "Strategic priorities recap",
          "Constraints / non-negotiables"],
         GREEN_DARK),
        ("ACT 2", "ALIGN HOW", "90 min",
         "Each Champion presents top-3 Initiatives + dependencies.",
         ["Top-3 Initiatives per VS",
          "Capacity asks per VS",
          "Risks · assumptions called out"],
         GREEN_MID),
        ("ACT 3", "SPOT DEPS", "60 min",
         "Wall mapping — cross-VS dependencies surfaced and resolved.",
         ["Cross-VS dependency map",
          "Owners agreed per arrow",
          "Sequencing locked per quarter"],
         GREEN_LIGHT),
        ("ACT 4", "LOCK OKRs", "60 min",
         "Capacity commitments + OKRs signed by Director per VS.",
         ["Quarterly OKRs locked",
          "Capacity commitments per VS",
          "Escalation list opened"],
         GREEN_CARD),
    ]

    for i, (act_id, act_name, dur, summary, bullets, fill) in enumerate(acts):
        cx = Inches(0.40) + i * (acts_w + Inches(0.10))
        round_rect(s, cx, acts_y, acts_w, acts_h, WHITE, radius=0.03)
        round_rect(s, cx, acts_y, acts_w, Inches(0.78), fill, radius=0.03)
        text(s, cx, acts_y + Inches(0.06), acts_w, Inches(0.22),
             act_id, size=8, bold=True, color=GREEN_PASTEL, font=BODY_FONT,
             align=PP_ALIGN.CENTER)
        text(s, cx, acts_y + Inches(0.26), acts_w, Inches(0.30),
             act_name, size=13, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, cx, acts_y + Inches(0.52), acts_w, Inches(0.20),
             "▸ " + dur, size=8, italic=True, color=GREEN_PASTEL, font=BODY_FONT,
             align=PP_ALIGN.CENTER)

        by = acts_y + Inches(0.95)
        text(s, cx + Inches(0.12), by, acts_w - Inches(0.24), Inches(0.55),
             summary, size=9, italic=True, color=GREEN_DARK, font=BODY_FONT)
        by += Inches(0.65)
        text(s, cx + Inches(0.12), by, acts_w - Inches(0.24), Inches(0.20),
             "OUTPUTS", size=7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        by += Inches(0.22)
        for b in bullets:
            text(s, cx + Inches(0.12), by, acts_w - Inches(0.24), Inches(0.46),
                 "▪ " + b, size=9, color=DARK_TEXT, font=BODY_FONT)
            by += Inches(0.42)

    # Bottom strip — Inputs / Outputs / Cadence
    bot_y = Inches(6.05)
    bot_h = Inches(1.05)
    round_rect(s, Inches(0.40), bot_y, Inches(12.55), bot_h, GREEN_BG, radius=0.04)
    cell_w_b = (Inches(12.55) - Inches(0.10) * 2 - Inches(0.30)) / 3
    bot_cells = [
        ("INPUTS",
         "Outputs of Gate 2 sessions of the quarter · Pain Score rankings · "
         "Capacity heat-map per VS · Corporate OKR cascade"),
        ("OUTPUTS",
         "Quarterly OKRs locked · OVS KPI deltas reviewed (NPS · CSAT · Reliability · cycle time) · "
         "Cross-VS dependency map · Capacity commitments per VS"),
        ("RHYTHM",
         "Q1 · Q2 · Q3 · Q4 — first week of each quarter · "
         "outputs reviewed at the next month's Gate 2"),
    ]
    for i, (label, body) in enumerate(bot_cells):
        cx = Inches(0.55) + i * (cell_w_b + Inches(0.10))
        text(s, cx, bot_y + Inches(0.10), cell_w_b, Inches(0.22),
             label, size=8, bold=True, color=GREEN_DARK, font=BODY_FONT)
        text(s, cx, bot_y + Inches(0.32), cell_w_b, Inches(0.66),
             body, size=9, color=DARK_TEXT, font=BODY_FONT)

    footer(s, "BRP · Deep Dive")
    return s


# ============================================================================
# H — VALUE & PRIORITIZATION (grounding slide before slide 41)
# ============================================================================

def build_value_prioritization(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_bar(s, prs.slide_width,
               "VALUE  &  PRIORITIZATION",
               "What we mean by VALUE · how we rank at PORTFOLIO · how we pull at TEAM")

    # ROW 1 — formula + CoD callout
    row1_y = Inches(1.15)
    row1_h = Inches(2.20)

    # Left: Pain Score formula card
    formula_w = Inches(7.50)
    round_rect(s, Inches(0.40), row1_y, formula_w, row1_h, WHITE, radius=0.04)
    round_rect(s, Inches(0.40), row1_y, formula_w, Inches(0.50), HEADER_BLUE, radius=0.04)
    text(s, Inches(0.55), row1_y + Inches(0.10), formula_w - Inches(0.30), Inches(0.30),
         "PAIN SCORE  ·  the value-vs-effort lens used at portfolio gate",
         size=11, bold=True, color=WHITE, font=BODY_FONT)
    # Formula
    text(s, Inches(0.55), row1_y + Inches(0.65), formula_w - Inches(0.30), Inches(0.55),
         "H(x)  =  ( BV · w₁  +  UX · w₂  +  Safety · w₃ )  ÷  Effort",
         size=18, bold=True, color=AVANGRID_GREEN, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Variables explained — 4 mini-cells
    var_y = row1_y + Inches(1.30)
    var_w = (formula_w - Inches(0.50)) / 4
    var_data = [
        ("BV", "Business Value 1–5",  "$ · NPS · CSAT · compliance"),
        ("UX", "User Experience 1–5", "customer-first · outcome > output"),
        ("Safety", "Risk reduction 1–5", "operational · regulatory · reliability"),
        ("Effort", "T-shirt → weight", "S=1 · M=2 · L=4 · XL=8"),
    ]
    for i, (label, sub, ex) in enumerate(var_data):
        vx = Inches(0.55) + i * var_w
        text(s, vx, var_y, var_w, Inches(0.18),
             label, size=9, bold=True, color=AVANGRID_GREEN, font=BODY_FONT,
             align=PP_ALIGN.CENTER)
        text(s, vx, var_y + Inches(0.20), var_w, Inches(0.18),
             sub, size=8, color=DARK_TEXT, font=BODY_FONT,
             align=PP_ALIGN.CENTER)
        text(s, vx, var_y + Inches(0.38), var_w, Inches(0.18),
             ex, size=7.5, italic=True, color=GREY_LABEL, font=BODY_FONT,
             align=PP_ALIGN.CENTER)

    # Right: Cost of Delay callout
    cod_x = Inches(8.05)
    cod_w = Inches(5.00)
    round_rect(s, cod_x, row1_y, cod_w, row1_h, WHITE, radius=0.04)
    round_rect(s, cod_x, row1_y, cod_w, Inches(0.50), GREEN_CARD, radius=0.04)
    text(s, cod_x + Inches(0.15), row1_y + Inches(0.10), cod_w - Inches(0.30), Inches(0.30),
         "COST OF DELAY  ·  the urgency lens",
         size=11, bold=True, color=WHITE, font=BODY_FONT)
    text(s, cod_x + Inches(0.15), row1_y + Inches(0.62), cod_w - Inches(0.30), Inches(0.40),
         "$ value lost per unit of time the work is delayed",
         size=10, italic=True, color=GREEN_DARK, font=BODY_FONT)
    text(s, cod_x + Inches(0.15), row1_y + Inches(1.05), cod_w - Inches(0.30), Inches(0.50),
         "Two demands with equal Pain Score → tiebreak by CoD.\n"
         "Used as WSJF input  ·  WSJF = CoD ÷ Job Size",
         size=9, color=DARK_TEXT, font=BODY_FONT)
    text(s, cod_x + Inches(0.15), row1_y + Inches(1.65), cod_w - Inches(0.30), Inches(0.30),
         "EXAMPLE  ·  Regulatory deadline missed → fine $50k/month → CoD = $50k/month",
         size=8.5, italic=True, color=GREEN_DARK, font=BODY_FONT)

    # ROW 2 — Portfolio vs Team prioritization
    row2_y = Inches(3.55)
    row2_h = Inches(3.05)
    side_w = (Inches(12.55) - Inches(0.20)) / 2

    # PORTFOLIO side
    round_rect(s, Inches(0.40), row2_y, side_w, row2_h, WHITE, radius=0.03)
    round_rect(s, Inches(0.40), row2_y, side_w, Inches(0.55), HEADER_BLUE_DK, radius=0.03)
    text(s, Inches(0.55), row2_y + Inches(0.10), side_w - Inches(0.30), Inches(0.32),
         "PORTFOLIO  ·  HOW WE RANK",
         size=13, bold=True, color=WHITE, font=BODY_FONT)
    text(s, Inches(0.55), row2_y + Inches(0.40), side_w - Inches(0.30), Inches(0.20),
         "Director + VP at Gate 2 · monthly · sequences the Backlog",
         size=9, italic=True, color=GREEN_PASTEL, font=BODY_FONT)
    pf_rows = [
        ("WHO",         "Director (Carlos · Kumar) · VP (Vitor) · Architect"),
        ("WHEN",        "Monthly · Gate 2 · 90 min"),
        ("RANKED BY",   "1. CoS (Expedite/Fixed-Date jump) · 2. Pain Score · 3. Cost of Delay tiebreak"),
        ("OUTPUT",      "Sequenced Backlog · capacity slots assigned · status = Ready For Execution"),
        ("RE-RANK",     "Quarterly at BRP · monthly at Gate 2 · ad-hoc only for Expedite"),
    ]
    by = row2_y + Inches(0.75)
    for label, body in pf_rows:
        text(s, Inches(0.55), by, Inches(1.30), Inches(0.30),
             label, size=8, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, Inches(1.85), by, side_w - Inches(1.55), Inches(0.36),
             body, size=10, color=DARK_TEXT, font=BODY_FONT)
        by += Inches(0.42)

    # TEAM side
    team_x = Inches(0.40) + side_w + Inches(0.20)
    round_rect(s, team_x, row2_y, side_w, row2_h, WHITE, radius=0.03)
    round_rect(s, team_x, row2_y, side_w, Inches(0.55), GREEN_DARK, radius=0.03)
    text(s, team_x + Inches(0.15), row2_y + Inches(0.10), side_w - Inches(0.30), Inches(0.32),
         "TEAM  ·  HOW WE PULL",
         size=13, bold=True, color=WHITE, font=BODY_FONT)
    text(s, team_x + Inches(0.15), row2_y + Inches(0.40), side_w - Inches(0.30), Inches(0.20),
         "Squad · daily · pulls from the top of the ranked Backlog",
         size=9, italic=True, color=GREEN_PASTEL, font=BODY_FONT)
    tm_rows = [
        ("WHO",         "Tech Lead + Product Manager + squad members"),
        ("WHEN",        "Daily standup · weekly replenishment · service review"),
        ("PULL RULE",   "Free WIP slot → next item top of Backlog respecting CoS lane"),
        ("WIP LIMITS",  "Standard 3–4 · Fixed-Date 2 · Expedite 1 (jumps queue) · Intangible 1"),
        ("OUTPUT",      "Story moves into Phase 2 · cycle-time clock starts"),
    ]
    by = row2_y + Inches(0.75)
    for label, body in tm_rows:
        text(s, team_x + Inches(0.15), by, Inches(1.30), Inches(0.30),
             label, size=8, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, team_x + Inches(1.45), by, side_w - Inches(1.55), Inches(0.36),
             body, size=10, color=DARK_TEXT, font=BODY_FONT)
        by += Inches(0.42)

    # Bottom takeaway with OVS KPI panel
    take_y = Inches(6.75)
    round_rect(s, Inches(0.40), take_y, Inches(12.55), Inches(0.36), GREEN_BG, radius=0.04)
    text(s, Inches(0.55), take_y, Inches(12.2), Inches(0.36),
         "OVS KPI PANEL  ·  Meter-to-Cash  ·  NPS  ·  CSAT  ·  Customer Reliability  ·  Ideation-to-Delivery cycle  ·  "
         "every Backlog item should map to ≥ 1 of these outcomes.",
         size=9, bold=True, italic=True, color=GREEN_DARK, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)

    footer(s, "Value & Prioritization")
    return s


# ============================================================================
# A — Slide 32 fixes
# ============================================================================

SLIDE32_TEXT_REPLACEMENTS = [
    # (old, new)
    ("PM steward + Champion", "Demand Manager + Champion"),
    ("validate DoR · finalize CoS · compute Pain Score · A5 reads Size + Type",
     "validate DoR · A5 reads Size + Type → auto-convert · @notify Champion"),
    ("Enterprise Dashboard · Gate 1 first", "Enterprise Dashboard · Gate 2 next"),
    # Swimlane labels for XL path
    ("IDEA", "FUNNEL"),
    ("GATE 1", "REVIEW"),
    ("GATE 2", "ANALYZE"),
    ("RUN",   "BACKLOG"),
    # NB: the M/L path also has GATE labels? Check before fix.
]


def fix_slide32(slide):
    """Walk every run on slide 32, apply text replacements, log what changed."""
    changes = 0
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                old_text = r.text
                new_text = old_text
                for old, new in SLIDE32_TEXT_REPLACEMENTS:
                    if old in new_text:
                        new_text = new_text.replace(old, new)
                if new_text != old_text:
                    r.text = new_text
                    changes += 1
    return changes


# ============================================================================
# B — Field-guide badge updates (slides 21–28)
# ============================================================================

BADGE_MAP = {
    21: "● FUNNEL  ·  REQUESTER",
    22: "● FUNNEL  ·  REQUESTER",
    23: "● NARRATIVE: FUNNEL · REQUESTER  /  SCORE: ANALYZING · ARCHITECT",
    24: "● FUNNEL  ·  REQUESTER",
    25: "● FUNNEL  ·  REQUESTER",
    26: "● REVIEWING / ANALYZING  ·  DEMAND MANAGER + ARCHITECT",
    27: "● REVIEWING  ·  DEMAND MANAGER",
    28: "● REVIEWING  ·  DEMAND MANAGER",
}


def update_badge(slide, new_text):
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        if "MANDATORY" not in sh.text_frame.text:
            continue
        # Widen the shape so longer text fits
        sh.left = Inches(8.80)
        sh.width = Inches(4.20)
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if "MANDATORY" in r.text:
                    r.text = new_text
                    if len(new_text) > 45:
                        r.font.size = Pt(7)
                    elif len(new_text) > 28:
                        r.font.size = Pt(8)
                    return True
    return False


# ============================================================================
# MAIN
# ============================================================================

print("Loading TOM MAP workbook-1.pptx ...")
prs = Presentation(SRC)
total_before = len(prs.slides)
print(f"  slides BEFORE: {total_before}")

# --- A. Fix slide 32 ---
print("\n[A] Fixing slide 32 (After Submission) ...")
n32 = fix_slide32(prs.slides[31])  # 0-based: slide 32 = index 31
print(f"   text replacements applied: {n32}")

# --- B. Update field-guide badges ---
print("\n[B] Updating field-guide badges (slides 21–28) ...")
for idx, badge in BADGE_MAP.items():
    ok = update_badge(prs.slides[idx-1], badge)
    print(f"   slide {idx:02d}  →  {badge!r:60s}  {'OK' if ok else 'NOT FOUND'}")

# --- C..H. Build new slides ---
print("\n[C–H] Building 6 new slides ...")
build_field_progression(prs);   print("   01 FIELD PROGRESSION         built")
build_automation_catalog(prs);  print("   02 AUTOMATION CATALOG (A1–A6) built")
build_team_kanban_deepdive(prs);print("   03 TEAM KANBAN DEEP DIVE      built")
build_gates_deepdive(prs);      print("   04 GATES DEEP DIVE            built")
build_brp_deepdive(prs);        print("   05 BRP DEEP DIVE              built")
build_value_prioritization(prs);print("   06 VALUE & PRIORITIZATION     built")

# --- Reorder ---
# After adding 6 slides at the end, their 0-based indices are:
#   total_before     → 01 FIELD PROGRESSION
#   total_before + 1 → 02 AUTOMATION CATALOG
#   total_before + 2 → 03 TEAM KANBAN DEEP
#   total_before + 3 → 04 GATES DEEP
#   total_before + 4 → 05 BRP DEEP
#   total_before + 5 → 06 VALUE & PRIORITIZATION
#
# Target placements (0-based target index after each insert applied in order):
#   01 → between slide 20 and 21 → target 0-based 20  (becomes new slide 21)
#   02 → after slide 32          → after the +1 shift, target 0-based 33
#   03 → after slide 35          → after 2 shifts,    target 0-based 37
#   04 → after slide 39          → after 3 shifts,    target 0-based 42
#   05 → right after 04          →                    target 0-based 43
#   06 → before slide 41         → after 4 shifts,    target 0-based 45

print("\n[I] Reordering ...")
# Source indices remain stable across moves because each move goes
# from a HIGH index (end of deck) to a LOWER index — the remove+insert
# offsets cancel out for positions higher than old_index.
fp_idx  = total_before + 0   # 80
ac_idx  = total_before + 1   # 81
tk_idx  = total_before + 2   # 82
gd_idx  = total_before + 3   # 83
brp_idx = total_before + 4   # 84
vp_idx  = total_before + 5   # 85

# 1. FIELD PROGRESSION → index 20 (becomes slide 21)
move_slide(prs, fp_idx, 20)
print(f"   FIELD PROGRESSION   → slide 21")

# 2. AUTOMATION CATALOG → index 33 (after current slide 32-now-33 = original 32)
move_slide(prs, ac_idx, 33)
print(f"   AUTOMATION CATALOG  → slide 34")

# 3. TEAM KANBAN DEEP → index 37 (after original slide 35, now shifted to 36)
move_slide(prs, tk_idx, 37)
print(f"   TEAM KANBAN DEEP    → slide 38")

# 4. GATES DEEP DIVE → index 42 (after original slide 39, now shifted to 41)
move_slide(prs, gd_idx, 42)
print(f"   GATES DEEP DIVE     → slide 43")

# 5. BRP DEEP DIVE → index 43 (right after Gates Deep Dive)
move_slide(prs, brp_idx, 43)
print(f"   BRP DEEP DIVE       → slide 44")

# 6. VALUE & PRIORITIZATION → index 46 (before PRIORITY BACKLOG = original slide 42, now shifted to 46)
move_slide(prs, vp_idx, 46)
print(f"   VALUE & PRIORITIZATION → slide 47")

# --- Save ---
print(f"\n[J] Saving → {DST}")
prs.save(DST)
print(f"   slides AFTER: {len(prs.slides)}")
print("\nDONE.")
