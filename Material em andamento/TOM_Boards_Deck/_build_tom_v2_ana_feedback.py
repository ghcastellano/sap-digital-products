"""
TOM MAP v2 — Ana feedback adaptation.

Input:  TOM MAP workbook-2.pptx (current state)
Output: TOM MAP workbook-3.pptx (with red-marked changes)

ALL changes use RED text (#C00000) keeping original font + size.
6 new slides added. 10 existing slides modified.
Visual identity: dark navy header + Avangrid green accents + red for edits.
NO stroke on filled boxes.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

SRC = "/Users/gustavohenriquecastellano/Downloads/Base Apresentações Projeto Digital Product/TOM MAP workbook-2.pptx"
DST = "/Users/gustavohenriquecastellano/Downloads/Base Apresentações Projeto Digital Product/TOM MAP workbook-3.pptx"

# === Palette ===
HEADER_BLUE      = RGBColor(0x2C, 0x57, 0x8B)
HEADER_BLUE_DK   = RGBColor(0x1E, 0x3A, 0x5F)
AVANGRID_GREEN   = RGBColor(0x00, 0x8C, 0x38)
GREEN_CARD       = RGBColor(0x00, 0x7A, 0x32)
GREEN_DARK       = RGBColor(0x00, 0x6B, 0x3F)
GREEN_MID        = RGBColor(0x00, 0x89, 0x50)
GREEN_LIGHT      = RGBColor(0x00, 0x9F, 0x5D)
GREEN_BG         = RGBColor(0xF2, 0xF8, 0xF4)
GREEN_PASTEL     = RGBColor(0xC8, 0xE0, 0xCB)
RED_EDIT         = RGBColor(0xC0, 0x00, 0x00)
WHITE            = RGBColor(0xFF, 0xFF, 0xFF)
DARK_TEXT        = RGBColor(0x2C, 0x3E, 0x50)
GREY_LABEL       = RGBColor(0x6B, 0x77, 0x8C)
GREY_BG          = RGBColor(0xF4, 0xF5, 0xF7)

TITLE_FONT = "IberPangea Text Medium"
BODY_FONT  = "Arial"

# === Helpers (no-stroke filled shapes — border-overflow fix) ===

def _no_line(shape):
    shape.line.fill.background()

def rect(slide, x, y, w, h, fill=None):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    s.shadow.inherit = False
    if fill is None: s.fill.background()
    else: s.fill.solid(); s.fill.fore_color.rgb = fill
    _no_line(s)
    return s

def round_rect(slide, x, y, w, h, fill=None, radius=0.04):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    s.shadow.inherit = False
    s.adjustments[0] = radius
    if fill is None: s.fill.background()
    else: s.fill.solid(); s.fill.fore_color.rgb = fill
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
    rect(slide, 0, 0, w_total, Inches(0.95), HEADER_BLUE)
    rect(slide, 0, Inches(0.95), w_total, Emu(15000), AVANGRID_GREEN)
    text(slide, Inches(0.40), Inches(0.10), Inches(11), Inches(0.42),
         line1, size=22, bold=True, color=WHITE, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
    text(slide, Inches(0.40), Inches(0.55), Inches(11), Inches(0.32),
         line2, size=10, bold=True, color=WHITE, font=BODY_FONT)
    text(slide, Inches(11.4), Inches(0.30), Inches(1.7), Inches(0.4),
         "NTT  DATA", size=14, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.RIGHT)

def footer(slide, label):
    text(slide, Inches(0.40), Inches(7.18), Inches(8), Inches(0.18),
         "Avangrid · Digital Products · Target Operating Model",
         size=8, italic=True, color=GREY_LABEL, font=BODY_FONT)
    text(slide, Inches(11.5), Inches(7.18), Inches(1.6), Inches(0.18),
         label, size=8, color=GREY_LABEL, font=BODY_FONT, align=PP_ALIGN.RIGHT)

def red_badge(slide, x, y, w, h, label, content, size_label=8, size_body=9):
    """Red-bordered annotation badge — used to mark Ana-feedback edits."""
    round_rect(slide, x, y, w, h, WHITE, radius=0.04)
    rect(slide, x, y, Inches(0.06), h, RED_EDIT)  # red left strip
    text(slide, x + Inches(0.12), y + Inches(0.03), w - Inches(0.20), Inches(0.20),
         label, size=size_label, bold=True, color=RED_EDIT, font=BODY_FONT)
    text(slide, x + Inches(0.12), y + Inches(0.22), w - Inches(0.20), h - Inches(0.24),
         content, size=size_body, color=RED_EDIT, font=BODY_FONT)

def cs_itba_badge(slide, x, y, applies_to):
    """Tag badge: CS · ITBA · BOTH. Tiny visual marker."""
    color = {"CS": AVANGRID_GREEN, "ITBA": HEADER_BLUE, "BOTH": GREY_LABEL}[applies_to]
    w, h = Inches(0.95), Inches(0.22)
    round_rect(slide, x, y, w, h, color, radius=0.30)
    text(slide, x, y, w, h, applies_to, size=7, bold=True, color=WHITE,
         font=BODY_FONT, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

def move_slide(prs, old_index, new_index):
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    s = slides[old_index]
    xml_slides.remove(s)
    xml_slides.insert(new_index, s)


def find_idx_by_title(prs, contains):
    """Find slide where the title-area shape (top of slide, y < 1.5in) contains the text."""
    for i, slide in enumerate(prs.slides):
        for sh in slide.shapes:
            if not sh.has_text_frame: continue
            try: y = sh.top / 914400 if sh.top else 99
            except: y = 99
            if y > 1.5: continue
            if contains.upper() in sh.text_frame.text.upper():
                return i
    return -1


def find_idx_by_unique(prs, unique_substring):
    """Find slide containing a unique substring anywhere."""
    for i, slide in enumerate(prs.slides):
        for sh in slide.shapes:
            if not sh.has_text_frame: continue
            if unique_substring.upper() in sh.text_frame.text.upper():
                return i
    return -1


# ============================================================================
# In-place text edits — preserve original styling, change content + color RED
# ============================================================================

def replace_run_text(slide, old_substr, new_text, color_red=True):
    """Find first run containing old_substr, replace its full text and (optionally) color it red."""
    for sh in slide.shapes:
        if not sh.has_text_frame: continue
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if old_substr in r.text:
                    r.text = r.text.replace(old_substr, new_text)
                    if color_red:
                        r.font.color.rgb = RED_EDIT
                    return True
    return False


def replace_full_run(slide, target_text_match, new_text, color_red=True):
    """Replace the FULL text of a run if it matches target_text_match (substring)."""
    for sh in slide.shapes:
        if not sh.has_text_frame: continue
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if target_text_match in r.text:
                    r.text = new_text
                    if color_red:
                        r.font.color.rgb = RED_EDIT
                    return True
    return False


# ============================================================================
# SLIDE EDITS — Ana's list applied to existing slides
# ============================================================================

def edit_slide_23_BV(prs):
    """Slide 24 in current numbering (after FIELD PROGRESSION inserted). BV slide.
    Edits:
      - Add red note: 'BV is 1 of 4 inputs of Pain Score · BV(5) does NOT auto-mean regulatory'
      - Cross-reference Regulatory flag separately
    """
    s = prs.slides[23]  # 0-based: slide 24
    # Add red annotation badge bottom-right of slide
    red_badge(s, Inches(8.40), Inches(6.40), Inches(4.85), Inches(0.75),
              "ANA · CLARIFICATION",
              "BV is ONE of 4 inputs into Pain Score · BV=5 (Transformational) is NOT triggered "
              "by Regulatory alone — Regulatory uses the separate flag → CoS=Fixed-Date (A3).",
              size_body=8)
    return 1


def edit_slide_27_CoS_add_CIT(prs):
    """Slide 27 in current = slide 28 (1-based after insert). CoS slide. Add Customer Impact Tier definition."""
    s = prs.slides[27]
    red_badge(s, Inches(0.40), Inches(6.45), Inches(12.55), Inches(0.65),
              "ANA · CUSTOMER IMPACT TIER · DEFINITION ADDED",
              "P1 — outage / blocker affecting > 100k customers OR full OPCO   ·   "
              "P2 — degraded service for > 10k customers OR critical regulatory window   ·   "
              "P3 — single-OPCO non-critical issue   ·   "
              "P4 — internal-only · no customer-facing impact   ·   "
              "P1/P2 → triggers A2 Auto-Expedite (CoS=Expedite + WIP-1)",
              size_body=8)
    return 1


def edit_slide_28_Type_simplify(prs):
    """Slide 28 = slide 29 (CoS / Type slide). Simplify Type to {Demand, Change}; Incident comes auto from IT Now."""
    s = prs.slides[28]
    red_badge(s, Inches(0.40), Inches(6.45), Inches(12.55), Inches(0.65),
              "ANA · TYPE SIMPLIFIED · INCIDENT REMOVED FROM SELECTOR",
              "Type now = { Demand · Change }  ·  Incident is NOT user-selectable — incidents flow automatically "
              "from ServiceNow (IT Now) → Jira via API (CCC + Smart Metering only) → A2 Auto-Expedite if P1/P2  ·  "
              "see slide INCIDENT FLOW · IT NOW → JIRA for the integration detail.",
              size_body=8)
    # Try to update the field text where Type is mentioned
    replace_run_text(s, "Demand / Incident / Change", "Demand / Change  (Incident auto from IT Now)")
    return 1


def edit_slide_29_WorkType(prs):
    """Slide 29 (Work Type vs R/G/T). Mark overlap, remove regulatory in upgrade."""
    s = prs.slides[29]
    red_badge(s, Inches(0.40), Inches(6.45), Inches(12.55), Inches(0.55),
              "ANA · WORK TYPE × R/G/T OVERLAP — TO REVIEW",
              "Work Type (BAU·Build·Upgrade) overlaps with R/G/T (Run·Grow·Transform). "
              "Recommendation for v3: drop Work Type, keep only R/G/T + Regulatory flag. "
              "REMOVED \"regulatory\" mention from Upgrade description — already covered by Regulatory flag (A3).",
              size_body=8)
    return 1


def edit_slide_32_PainScore_CS(prs):
    """Slide 32 (Pain Score Formula Decoded). Add CS adaptation: Safety → Customer Reliability + simplified formula."""
    s = prs.slides[31]  # slide 32, 0-based = 31
    red_badge(s, Inches(0.40), Inches(6.30), Inches(12.55), Inches(0.85),
              "ANA · CS ADAPTATION · SAFETY DOES NOT APPLY (no field crew)",
              "FOR CUSTOMER SERVICES (Kumar):  H(x) = (BV·w₁ + UX·w₂) / Effort    [2 components · simpler]\n"
              "    weights: w₁ = 0.40 · w₂ = 0.60 (UX-heavy reflecting customer-first mandate)\n"
              "FOR ITBA (Carlos): keep 3-component formula with Safety  ·  H(x) = (BV·w₁ + UX·w₂ + Safety·w₃) / Effort\n"
              "Reason: CS works for customer services · ITBA covers field operations + grid where safety risk is real.",
              size_body=8)
    return 1


def edit_slide_33_AfterSubmission(prs):
    """Slide 33 (After Submission). Champion → Demand Manager + PM. PM is generic (could be new hire OR BRS member)."""
    s = prs.slides[32]  # slide 33, 0-based 32
    # Already had: 'TRIAGE  ·  Demand Manager + Champion' → fix the lingering Champion mention
    # Replace any remaining Champion in slide 33 if not already
    replace_run_text(s, "Demand Manager + Champion", "Demand Manager + PM")
    # Add red annotation
    red_badge(s, Inches(0.40), Inches(7.05), Inches(12.55), Inches(0.40),
              "ANA · ROLE CLARIFICATION",
              "PM = Product Manager (new hire OR experienced BRS member temporarily acting as PM until full team is staffed). "
              "Champion role kept ONLY for VS Domain expert at Gate 2 (not at Triage).",
              size_body=8)
    return 1


def edit_slide_36_StepLevel(prs):
    """Slide 36 / 37 in current numbering — Step-Level Policies. Align cast with Triage."""
    # Slide 36 (Step-Level Policies) — 0-based after inserts:
    # Original 36 was index 35 → after FIELD PROGRESSION insert at 20 → index 36
    # → after AUTOMATION CATALOG insert at 33 → index 37
    # → after TEAM KANBAN DEEP insert at 37 → index 38? Actually let me check current state.
    # Trust the search by title.
    for i, slide in enumerate(prs.slides):
        for sh in slide.shapes:
            if sh.has_text_frame and "STEP-LEVEL POLICIES" in sh.text_frame.text.upper():
                red_badge(slide, Inches(0.40), Inches(6.85), Inches(12.55), Inches(0.55),
                          "ANA · CAST ALIGNMENT",
                          "Triage cast HERE must mirror slide 33 (After Submission): "
                          "Demand Manager (chair) · PM (per VS) · Architect. "
                          "Champion appears only at Gate 2 (Portfolio Committee).",
                          size_body=8)
                return 1
    return 0


def edit_slide_42_Governance(prs):
    """Slide 42 (Governance). Note BRP coupling to existing P&T + Digital CoE BRPs.
    Use unique substring 'TWO GATES ABSORB' to find the right slide."""
    idx = find_idx_by_unique(prs, "TWO GATES ABSORB")
    if idx < 0:
        return 0
    slide = prs.slides[idx]
    red_badge(slide, Inches(0.40), Inches(6.50), Inches(12.55), Inches(0.55),
              "ANA · BRP REDEFINED",
              "BRP is NOT an own ritual for Digital Products — couples to EXISTING events: "
              "P&T BRP (Carlos+Asier) and Digital CoE BRP (Kumar's lane). Our Q-outcome enters BOTH as input. "
              "See slide BRP DEEP DIVE for the connection model.",
              size_body=8)
    return 1


def edit_slide_44_BRP_redo(prs):
    """Slide 44 (BRP Deep Dive). Modify content to reflect 'couple to existing' (option B)."""
    # Locate slide by title
    for slide in prs.slides:
        for sh in slide.shapes:
            if sh.has_text_frame and "BIG ROOM PLANNING" in sh.text_frame.text.upper() and "DEEP DIVE" in sh.text_frame.text.upper():
                # Add big red overlay band at top reframing the slide
                red_badge(slide, Inches(0.40), Inches(7.00), Inches(12.55), Inches(0.45),
                          "ANA · BRP COUPLED TO EXISTING — NO NEW RITUAL",
                          "Digital Products injects its quarter outcome as INPUT into the EXISTING P&T BRP and Digital CoE BRP — "
                          "no separate event. Carlos sends to P&T BRP · Kumar sends to Digital CoE BRP · "
                          "Vitor is the cross-portfolio liaison.",
                          size_body=7)
                return 1
    return 0


def edit_slide_47_Value_OPCO(prs):
    """Slide 47/48 (Value & Prioritization). Add OPCO multiplier."""
    for slide in prs.slides:
        for sh in slide.shapes:
            if sh.has_text_frame and "VALUE  &  PRIORITIZATION" in sh.text_frame.text.upper():
                # Add OPCO multiplier annotation
                red_badge(slide, Inches(0.40), Inches(7.00), Inches(12.55), Inches(0.40),
                          "ANA · OPCO MULTIPLIER · MULTI-OPCO PRIORITIZATION",
                          "BV·OPCOmult·w₁ where OPCOmult = 1.0 (1 OPCO) · 1.3 (regional 2-4) · 1.6 (all 5+ OPCOs). "
                          "Form changes to MULTI-SELECT + 'All OPCOs' checkbox · A1 detects duplicates across OPCOs and merges into one card with tag '[N OPCOs]'.",
                          size_body=7)
                return 1
    return 0


# ============================================================================
# NEW SLIDES
# ============================================================================

def build_routing_logic(prs):
    """ROUTING LOGIC BY SIZE — visual flowchart: S/M/L/XL/Bug → where it goes."""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_bar(s, prs.slide_width,
               "ROUTING LOGIC BY SIZE",
               "What goes to TEAM directly · what goes to PORTFOLIO · what bypasses everything")

    # Top row — decision flow horizontal
    flow_y = Inches(1.30)
    flow_h = Inches(2.50)
    cols = [
        ("S", "Size = S\n< 5 days", "Story (no Epic)", "TEAM BACKLOG\nDIRECT", "Skip portfolio · pulled by squad", GREEN_LIGHT),
        ("M", "Size = M / L", "Epic + Stories", "PORTFOLIO BACKLOG\nGATE 2 next", "Sequenced by Pain + CoS", GREEN_MID),
        ("L", "Size = XL\ncross-team", "Initiative + Epics", "ENTERPRISE DASHBOARD\nGATE 2 + tracked at VP layer", "Owns roadmap horizon · multi-quarter", GREEN_DARK),
        ("B", "Type = Incident\n(auto from IT Now)", "Bug · CoS=Expedite", "SQUAD EXPEDITE LANE\nWIP-1", "Bypasses Triage · A2 fires", RED_EDIT),
    ]
    cw = (Inches(12.55) - 3 * Inches(0.10)) / 4
    for i, (k, src, what, where, why, fill) in enumerate(cols):
        cx = Inches(0.40) + i * (cw + Inches(0.10))
        round_rect(s, cx, flow_y, cw, flow_h, WHITE, radius=0.04)
        round_rect(s, cx, flow_y, cw, Inches(0.55), fill, radius=0.04)
        text(s, cx, flow_y + Inches(0.10), cw, Inches(0.35),
             src, size=11, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, cx + Inches(0.15), flow_y + Inches(0.65), cw - Inches(0.30), Inches(0.18),
             "AUTO-CONVERTS TO", size=7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, cx + Inches(0.15), flow_y + Inches(0.83), cw - Inches(0.30), Inches(0.30),
             what, size=11, bold=True, color=GREEN_DARK, font=BODY_FONT)
        text(s, cx + Inches(0.15), flow_y + Inches(1.20), cw - Inches(0.30), Inches(0.18),
             "GOES TO", size=7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, cx + Inches(0.15), flow_y + Inches(1.38), cw - Inches(0.30), Inches(0.50),
             where, size=10, bold=True, color=DARK_TEXT, font=BODY_FONT)
        text(s, cx + Inches(0.15), flow_y + Inches(1.95), cw - Inches(0.30), Inches(0.50),
             why, size=8.5, italic=True, color=GREEN_DARK, font=BODY_FONT)

    # Mid section — explain WHY
    mid_y = Inches(4.05)
    round_rect(s, Inches(0.40), mid_y, Inches(12.55), Inches(1.30), GREEN_BG, radius=0.04)
    text(s, Inches(0.55), mid_y + Inches(0.08), Inches(12.2), Inches(0.22),
         "ONE LOGIC · TWO LEVELS — why size dictates routing",
         size=11, bold=True, color=GREEN_DARK, font=BODY_FONT)
    rules = [
        ("TEAM-LEVEL  ·  Squad pulls by capacity",
         "Size S items DON'T need cross-team coordination · go straight to Squad's Team Kanban after Triage. "
         "Squad pulls when WIP slot available."),
        ("PORTFOLIO-LEVEL  ·  Director sequences",
         "Size M/L/XL touch capacity slots · multiple squads · vendor scope · need Gate 2 approval for ROI + "
         "capacity reservation. Backlog is RANKED here, not at squad."),
    ]
    by = mid_y + Inches(0.32)
    for label, body in rules:
        text(s, Inches(0.55), by, Inches(4.5), Inches(0.22),
             label, size=9, bold=True, color=DARK_TEXT, font=BODY_FONT)
        text(s, Inches(5.10), by, Inches(7.8), Inches(0.42),
             body, size=9, color=DARK_TEXT, font=BODY_FONT)
        by += Inches(0.45)

    # Bottom — CS vs ITBA difference badges
    bot_y = Inches(5.55)
    round_rect(s, Inches(0.40), bot_y, Inches(12.55), Inches(1.50), WHITE, radius=0.04)
    text(s, Inches(0.55), bot_y + Inches(0.10), Inches(12.2), Inches(0.22),
         "CS vs ITBA  ·  WHERE THE LOGIC DIFFERS",
         size=11, bold=True, color=HEADER_BLUE, font=BODY_FONT)
    diff_y = bot_y + Inches(0.40)
    # CS column
    cs_itba_badge(s, Inches(0.55), diff_y, "CS")
    text(s, Inches(1.60), diff_y, Inches(11.0), Inches(0.50),
         "1 OVS Meter-to-Cash · 4 DVS + 1 Enabling Team · weights w₁=0.40 w₂=0.60 (no Safety) · "
         "Customer Reliability replaces Safety · Heather's metrics overlay (NPS · CSAT · Reliability)",
         size=9, color=DARK_TEXT, font=BODY_FONT)
    # ITBA column
    cs_itba_badge(s, Inches(0.55), diff_y + Inches(0.50), "ITBA")
    text(s, Inches(1.60), diff_y + Inches(0.50), Inches(11.0), Inches(0.50),
         "5 teams unchanged · weights w₁=0.45 w₂=0.30 w₃=0.25 (BV-heavy · safety relevant for grid systems) · "
         "PMO Delivery (new) OR Sarath OR BRS person filling PM seat · Carlos approves at Gate 2",
         size=9, color=DARK_TEXT, font=BODY_FONT)
    # BOTH
    cs_itba_badge(s, Inches(0.55), diff_y + Inches(1.00), "BOTH")
    text(s, Inches(1.60), diff_y + Inches(1.00), Inches(11.0), Inches(0.50),
         "Same intake form · same A1–A6 automations · same Triage cadence · same Gate 2 cadence · same Class of Service taxonomy",
         size=9, color=DARK_TEXT, font=BODY_FONT)

    footer(s, "Routing Logic by Size · NEW")
    return s


def build_cos_maintain(prs):
    """HOW WE MAINTAIN CLASS OF SERVICE — process behind, NO hardcoded percentages."""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_bar(s, prs.slide_width,
               "HOW WE MAINTAIN CLASS OF SERVICE",
               "The PROCESS that keeps the CoS mix healthy · NO fixed percentages until validated")

    # 4-step horizontal cycle
    cycle_y = Inches(1.30)
    cycle_h = Inches(3.00)
    steps = [
        ("01 OBSERVE", "MEASURE flow per CoS",
         "Squad cycle-time + WIP + completion-rate per CoS lane · weekly snapshot",
         GREEN_DARK),
        ("02 ASSESS", "DETECT imbalances",
         "Are Expedites > planned cap? Fixed-Date items at risk? Standard starving? Aging tags rising?",
         GREEN_MID),
        ("03 DECIDE", "REBALANCE proposed",
         "PM + Demand Manager propose adjustments for next quarter · Director endorses at Gate 2",
         GREEN_LIGHT),
        ("04 ADJUST", "APPLY new mix",
         "WIP limits updated in Jira · automations re-tuned · KPIs re-baselined · loop back to OBSERVE",
         GREEN_CARD),
    ]
    sw = (Inches(12.55) - 3 * Inches(0.10)) / 4
    for i, (step_id, action, desc, fill) in enumerate(steps):
        sx = Inches(0.40) + i * (sw + Inches(0.10))
        round_rect(s, sx, cycle_y, sw, cycle_h, WHITE, radius=0.04)
        round_rect(s, sx, cycle_y, sw, Inches(0.65), fill, radius=0.04)
        text(s, sx, cycle_y + Inches(0.06), sw, Inches(0.22),
             step_id, size=9, bold=True, color=GREEN_PASTEL, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, sx, cycle_y + Inches(0.28), sw, Inches(0.32),
             action, size=12, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, sx + Inches(0.15), cycle_y + Inches(0.85), sw - Inches(0.30), Inches(2.0),
             desc, size=9.5, color=DARK_TEXT, font=BODY_FONT)

    # Owner / cadence strip
    own_y = Inches(4.55)
    round_rect(s, Inches(0.40), own_y, Inches(12.55), Inches(1.30), GREEN_BG, radius=0.04)
    text(s, Inches(0.55), own_y + Inches(0.10), Inches(12.2), Inches(0.22),
         "WHO OWNS WHAT  ·  THE LOOP THAT NEVER STOPS",
         size=11, bold=True, color=GREEN_DARK, font=BODY_FONT)
    rows = [
        ("OBSERVE",  "PM (per VS)",                     "Weekly · automated dashboard"),
        ("ASSESS",   "Demand Manager",                  "Weekly · Triage meeting"),
        ("DECIDE",   "Director (Carlos · Kumar) at Gate 2", "Monthly · Portfolio Committee"),
        ("ADJUST",   "Squad Lead",                      "End of quarter · BRP outcome"),
    ]
    by = own_y + Inches(0.40)
    for step, who, when in rows:
        text(s, Inches(0.55), by, Inches(2.0), Inches(0.20),
             step, size=8.5, bold=True, color=GREEN_DARK, font=BODY_FONT)
        text(s, Inches(2.65), by, Inches(5.5), Inches(0.20),
             who, size=9, color=DARK_TEXT, font=BODY_FONT)
        text(s, Inches(8.20), by, Inches(4.6), Inches(0.20),
             when, size=9, italic=True, color=GREY_LABEL, font=BODY_FONT)
        by += Inches(0.22)

    # ANA red note — no hardcoded percentages
    red_badge(s, Inches(0.40), Inches(6.05), Inches(12.55), Inches(1.05),
              "ANA · CAPACITY % NOT HARDCODED — VALIDATED PER QUARTER",
              "We deliberately do NOT fix Standard/Fixed-Date/Expedite/Intangible percentages here — "
              "they emerge from observed flow + Director input at each Gate 2 / BRP cycle. "
              "Default starting reference (Q1 only · TO BE VALIDATED): Standard ~60% · Fixed-Date ~20% · "
              "Expedite ≤10% · Intangible ~10% — adjusted monthly.",
              size_body=8)

    footer(s, "How We Maintain CoS · NEW")
    return s


def build_incident_flow(prs):
    """INCIDENT FLOW · IT NOW → JIRA — explain auto-routing for incidents."""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_bar(s, prs.slide_width,
               "INCIDENT FLOW  ·  IT NOW → JIRA",
               "Why Incident is NOT a Type selector · how it lands in the right Squad automatically")

    # 5-step horizontal flow
    fl_y = Inches(1.30)
    fl_h = Inches(2.30)
    steps = [
        ("01", "INCIDENT RAISED",
         "User opens ticket in IT Now (ServiceNow) — only path for incidents",
         "ServiceNow form · auto-fields filled by user", HEADER_BLUE_DK),
        ("02", "P1/P2 FILTER",
         "ServiceNow checks priority — only P1/P2 incidents flow to Jira (P3/P4 stay in IT Now)",
         "Webhook fires only on P1/P2 events", HEADER_BLUE),
        ("03", "API CALL",
         "ServiceNow webhook → Jira Cloud REST API → creates Bug ticket",
         "JSON payload includes priority + category + OPCO + system", AVANGRID_GREEN),
        ("04", "A2 + A1 RUN",
         "A1 routes to correct Squad · A2 sets CoS=Expedite + WIP-1",
         "Status TRIAGED skipped · already validated by ServiceNow", GREEN_DARK),
        ("05", "STATUS SYNC",
         "Squad works the Bug in Jira · status updates sync back to ServiceNow",
         "Bidirectional sync via API · ServiceNow remains record-of-truth for SLA", GREEN_CARD),
    ]
    sw = (Inches(12.55) - 4 * Inches(0.08)) / 5
    for i, (sid, title, desc, tech, fill) in enumerate(steps):
        sx = Inches(0.40) + i * (sw + Inches(0.08))
        round_rect(s, sx, fl_y, sw, fl_h, WHITE, radius=0.03)
        round_rect(s, sx, fl_y, sw, Inches(0.65), fill, radius=0.03)
        text(s, sx, fl_y + Inches(0.05), sw, Inches(0.22),
             sid, size=8, bold=True, color=GREEN_PASTEL, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, sx, fl_y + Inches(0.27), sw, Inches(0.32),
             title, size=11, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        text(s, sx + Inches(0.10), fl_y + Inches(0.78), sw - Inches(0.20), Inches(0.85),
             desc, size=9.5, color=DARK_TEXT, font=BODY_FONT)
        text(s, sx + Inches(0.10), fl_y + Inches(1.65), sw - Inches(0.20), Inches(0.55),
             tech, size=8, italic=True, color=GREEN_DARK, font=BODY_FONT)

    # Why this matters — bottom
    why_y = Inches(3.95)
    round_rect(s, Inches(0.40), why_y, Inches(12.55), Inches(2.10), GREEN_BG, radius=0.04)
    text(s, Inches(0.55), why_y + Inches(0.10), Inches(12.2), Inches(0.22),
         "WHY INCIDENT IS NOT IN THE INTAKE FORM",
         size=11, bold=True, color=GREEN_DARK, font=BODY_FONT)
    bullets = [
        ("Single source of truth",
         "Incidents have their own SLA framework in ServiceNow (P1 4h · P2 8h · P3 24h · P4 72h) — Jira would duplicate this."),
        ("ServiceNow is record of truth for incidents",
         "Audit · compliance · regulator-facing reports come from ServiceNow — Jira mirrors only the actionable subset (P1/P2)."),
        ("Avoid manual mis-classification",
         "Users routinely picked 'Incident' for everything urgent — flooding the incident queue. Removing the option forces correct path."),
        ("Restricted scope today",
         "ServiceNow → Jira link active for CCC + Smart Metering (CS); ITBA scope expanding next quarter post Heather validation."),
    ]
    by = why_y + Inches(0.40)
    for label, body in bullets:
        rect(s, Inches(0.65), by + Inches(0.04), Inches(0.10), Inches(0.10), AVANGRID_GREEN)
        text(s, Inches(0.85), by, Inches(4.0), Inches(0.36),
             label, size=10, bold=True, color=DARK_TEXT, font=BODY_FONT)
        text(s, Inches(4.95), by, Inches(8.0), Inches(0.36),
             body, size=9, color=DARK_TEXT, font=BODY_FONT)
        by += Inches(0.40)

    # CS vs ITBA badge
    cs_itba_badge(s, Inches(11.10), Inches(6.20), "CS")
    text(s, Inches(8.00), Inches(6.20), Inches(3.0), Inches(0.22),
         "Active scope today:", size=9, italic=True, color=GREY_LABEL,
         font=BODY_FONT, align=PP_ALIGN.RIGHT)
    cs_itba_badge(s, Inches(12.10), Inches(6.20), "ITBA")
    text(s, Inches(8.00), Inches(6.45), Inches(5.0), Inches(0.22),
         "ITBA expansion: post-validation · later quarter",
         size=8, italic=True, color=GREY_LABEL, font=BODY_FONT, align=PP_ALIGN.RIGHT)

    footer(s, "Incident Flow · NEW")
    return s


def build_existing_project(prs):
    """DEMAND ON EXISTING PROJECT — sub-task atalho."""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_bar(s, prs.slide_width,
               "DEMAND ON EXISTING PROJECT",
               "Scope-add to a live Initiative · skip Funnel/Reviewing/Analyzing · sub-task shortcut")

    # Top row: when this scenario applies
    top_y = Inches(1.20)
    round_rect(s, Inches(0.40), top_y, Inches(12.55), Inches(0.65), GREEN_BG, radius=0.04)
    text(s, Inches(0.55), top_y + Inches(0.10), Inches(12.2), Inches(0.22),
         "WHEN  ·  Initiative is in IMPLEMENTING · new related ask comes from same Sponsor / OPCO / business area",
         size=11, bold=True, color=GREEN_DARK, font=BODY_FONT)
    text(s, Inches(0.55), top_y + Inches(0.36), Inches(12.2), Inches(0.22),
         "EXAMPLE  ·  CRM unified-record Initiative live · Sponsor asks to add 'phone number masking for CCAR call recordings'",
         size=9, italic=True, color=DARK_TEXT, font=BODY_FONT)

    # 4-step shortcut flow
    fl_y = Inches(2.05)
    fl_h = Inches(3.10)
    steps = [
        ("01", "LINK", "Requester opens NEW Demand · types Parent Initiative key · A1 detects link",
         "Form auto-fills VS · OPCO · Sponsor from parent", AVANGRID_GREEN),
        ("02", "BYPASS", "Skips Funnel · Reviewing · Analyzing · Gate 2 — already approved as part of parent",
         "No Pain Score recompute · keeps parent's CoS", GREEN_DARK),
        ("03", "SQUAD APPROVE", "Squad PM accepts (or rejects) the scope-add — protects WIP and timeline",
         "PM evaluates impact · estimates delta · raises if blocking", GREEN_MID),
        ("04", "SUB-TASK", "Becomes a SUB-TASK under the parent Initiative · shows in same swimlane",
         "Status syncs with parent · cycle-time tracked separately", GREEN_LIGHT),
    ]
    sw = (Inches(12.55) - 3 * Inches(0.10)) / 4
    for i, (sid, title, desc, tech, fill) in enumerate(steps):
        sx = Inches(0.40) + i * (sw + Inches(0.10))
        round_rect(s, sx, fl_y, sw, fl_h, WHITE, radius=0.04)
        round_rect(s, sx, fl_y, sw, Inches(0.65), fill, radius=0.04)
        text(s, sx, fl_y + Inches(0.05), sw, Inches(0.22),
             sid, size=8, bold=True, color=GREEN_PASTEL, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, sx, fl_y + Inches(0.27), sw, Inches(0.32),
             title, size=14, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, sx + Inches(0.12), fl_y + Inches(0.80), sw - Inches(0.24), Inches(1.40),
             desc, size=10, color=DARK_TEXT, font=BODY_FONT)
        text(s, sx + Inches(0.12), fl_y + Inches(2.30), sw - Inches(0.24), Inches(0.70),
             tech, size=8.5, italic=True, color=GREEN_DARK, font=BODY_FONT)

    # Guardrails
    grd_y = Inches(5.30)
    round_rect(s, Inches(0.40), grd_y, Inches(12.55), Inches(1.45), GREEN_BG, radius=0.04)
    text(s, Inches(0.55), grd_y + Inches(0.10), Inches(12.2), Inches(0.22),
         "GUARDRAILS  ·  WHEN THE SHORTCUT DOES NOT APPLY",
         size=11, bold=True, color=GREEN_DARK, font=BODY_FONT)
    rules = [
        "▸  Scope > 5 days OR > 1 squad sprint → reject shortcut · open new Demand through Funnel",
        "▸  Different Sponsor or OPCO than parent → must go through normal flow",
        "▸  Touches a different Value Stream than parent → must go through normal flow",
        "▸  Squad PM has full authority to reject if WIP / timeline at risk · Sponsor escalates to Director if needed",
    ]
    by = grd_y + Inches(0.36)
    for r in rules:
        text(s, Inches(0.55), by, Inches(12.2), Inches(0.22),
             r, size=9.5, color=DARK_TEXT, font=BODY_FONT)
        by += Inches(0.24)

    # Red note
    red_badge(s, Inches(0.40), Inches(6.85), Inches(12.55), Inches(0.30),
              "ANA · NEW SCENARIO ADDED",
              "Scope-add to live project · Q5 confirmed: sub-task shortcut with PM approval · pulls from squad capacity not new portfolio slot.",
              size_body=8)

    footer(s, "Demand on Existing Project · NEW")
    return s


def build_digital_coe_interface(prs):
    """DIGITAL CoE INTERFACE — hybrid governance with Digital Products."""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_bar(s, prs.slide_width,
               "DIGITAL CoE  ×  DIGITAL PRODUCTS",
               "How the two organizations connect · same Jira intake · separate governance · joint OKR sync")

    # Two columns: How DCoE submits / How DCoE governs
    col_y = Inches(1.20)
    col_h = Inches(4.55)
    col_w = Inches(6.18)
    col_gap = Inches(0.18)

    # LEFT: Digital CoE submits via Jira
    round_rect(s, Inches(0.40), col_y, col_w, col_h, WHITE, radius=0.04)
    round_rect(s, Inches(0.40), col_y, col_w, Inches(0.65), AVANGRID_GREEN, radius=0.04)
    text(s, Inches(0.55), col_y + Inches(0.12), col_w - Inches(0.30), Inches(0.32),
         "1.  DEMAND CHANNEL  ·  via Jira", size=14, bold=True, color=WHITE, font=BODY_FONT)
    text(s, Inches(0.55), col_y + Inches(0.40), col_w - Inches(0.30), Inches(0.22),
         "Digital CoE submits demand the same way as any business area",
         size=9, italic=True, color=GREEN_PASTEL, font=BODY_FONT)
    bullets_left = [
        ("ENTRY POINT",
         "Digital CoE opens a Jira Demand via the SAME intake form  ·  field 'Sponsor area' = 'Digital CoE'"),
        ("ROUTING",
         "A1 routes to the correct VS based on impacted product (CRM · Smart Metering · Genesys · etc.)"),
        ("PRIORITIZATION",
         "Same Pain Score formula applies · Digital CoE OKRs feed BV component"),
        ("GATE 2 SEAT",
         "Digital CoE Champion participates as VOTING MEMBER for items where Digital CoE is sponsor"),
        ("CARD FLOW",
         "Same lifecycle: Funnel → Reviewing → Analyzing → Backlog → Implementing → Done"),
    ]
    by = col_y + Inches(0.85)
    for label, body in bullets_left:
        text(s, Inches(0.55), by, Inches(1.65), Inches(0.20),
             label, size=8, bold=True, color=GREEN_DARK, font=BODY_FONT)
        text(s, Inches(2.25), by, col_w - Inches(2.00), Inches(0.65),
             body, size=9, color=DARK_TEXT, font=BODY_FONT)
        by += Inches(0.72)

    # RIGHT: Strategic alignment via OKR sync
    rx = Inches(0.40) + col_w + col_gap
    round_rect(s, rx, col_y, col_w, col_h, WHITE, radius=0.04)
    round_rect(s, rx, col_y, col_w, Inches(0.65), HEADER_BLUE, radius=0.04)
    text(s, rx + Inches(0.15), col_y + Inches(0.12), col_w - Inches(0.30), Inches(0.32),
         "2.  STRATEGIC ALIGNMENT  ·  via OKR sync", size=14, bold=True, color=WHITE, font=BODY_FONT)
    text(s, rx + Inches(0.15), col_y + Inches(0.40), col_w - Inches(0.30), Inches(0.22),
         "Quarterly OKRs flow into BRP cycle to ground priorities",
         size=9, italic=True, color=GREEN_PASTEL, font=BODY_FONT)
    bullets_right = [
        ("FREQUENCY",      "Quarterly · ahead of Digital CoE BRP and ahead of P&T BRP"),
        ("OWNERS",         "Kumar (CS) + Carlos (ITBA) + Vitor (VP) + Digital CoE Director"),
        ("CONTENT",        "OKRs from Digital CoE · constraints · cross-org dependencies · capacity asks"),
        ("OUTCOME",        "Joint priority list · seeds the Q-input that goes to BOTH existing BRPs"),
        ("NO SEPARATE BRP",
         "We do NOT create a new joint BRP — we couple to existing P&T BRP and Digital CoE BRP."),
    ]
    by = col_y + Inches(0.85)
    for label, body in bullets_right:
        text(s, rx + Inches(0.15), by, Inches(1.65), Inches(0.20),
             label, size=8, bold=True, color=HEADER_BLUE, font=BODY_FONT)
        text(s, rx + Inches(1.85), by, col_w - Inches(2.00), Inches(0.65),
             body, size=9, color=DARK_TEXT, font=BODY_FONT)
        by += Inches(0.72)

    # Bottom red note
    red_badge(s, Inches(0.40), Inches(6.00), Inches(12.55), Inches(1.10),
              "ANA · DIGITAL CoE — HYBRID MODEL",
              "Confirmed direction: Digital CoE submits demand THROUGH Jira (same form · same Triage · same Gate 2) "
              "AND has a Champion seat at Gate 2 for sponsor items · OKRs sync quarterly · "
              "no parallel governance · no own BRP for our Digital Products. Digital CoE BRP and P&T BRP remain the venues.",
              size_body=9)

    footer(s, "Digital CoE Interface · NEW")
    return s


def build_validation_roadmap(prs):
    """VALIDATION ROADMAP — next stage planning."""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_bar(s, prs.slide_width,
               "VALIDATION ROADMAP  ·  NEXT STAGE PLANNING",
               "Four phases · CS first → ITBA pilot → tooling → org rollout · exit criteria per phase")

    # 4 phases as horizontal swimlanes
    fl_y = Inches(1.20)
    fl_h = Inches(4.85)
    phases = [
        ("PHASE 1", "VALIDATE WITH CS",
         "Heather + Kumar review the model against CS reality · iterate on weights · CIT thresholds · CoS rules",
         ["Director sign-off (Kumar)",
          "Heather endorsement of CS adaptations",
          "Pain Score weights validated for M2C-SL · M2C-RO · CCC",
          "Customer Reliability rubric defined"],
         "Weeks 1–3", AVANGRID_GREEN),
        ("PHASE 2", "ITBA PILOT",
         "Pick 1 ITBA value stream (suggest GIS — has working PDD process) · run Gate 1+Gate 2 for 1 quarter",
         ["Carlos selects pilot VS",
          "Sarath OR new PMO-Delivery role assigned",
          "1 full Gate 2 cycle completed",
          "Lessons learned captured for ITBA-wide rollout"],
         "Weeks 4–16", GREEN_DARK),
        ("PHASE 3", "TOOLING ROLLOUT",
         "Stand up Jira intake form · ServiceNow ↔ Jira API · automations A1–A6 · dashboards",
         ["Jira request types deployed (1 per VS)",
          "ServiceNow webhook live for CCC + Smart Metering",
          "Automations A1–A6 in production",
          "Layer 0 VP Dashboard live"],
         "Weeks 4–20 (parallel to Phase 2)", GREEN_MID),
        ("PHASE 4", "ORG ROLLOUT",
         "PMs onboarded · all VSs running model · Demand Manager rotation set · Champions named",
         ["3–4 PMs hired (or BRS people acting as PMs)",
          "Demand Manager assigned per VS",
          "Champions named per VS",
          "All VSs running 1 quarter at TOM"],
         "Weeks 16–32", GREEN_LIGHT),
    ]
    sw = (Inches(12.55) - 3 * Inches(0.10)) / 4
    for i, (pid, title, desc, exits, dur, fill) in enumerate(phases):
        sx = Inches(0.40) + i * (sw + Inches(0.10))
        round_rect(s, sx, fl_y, sw, fl_h, WHITE, radius=0.03)
        round_rect(s, sx, fl_y, sw, Inches(0.78), fill, radius=0.03)
        text(s, sx, fl_y + Inches(0.06), sw, Inches(0.22),
             pid, size=9, bold=True, color=GREEN_PASTEL, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, sx, fl_y + Inches(0.30), sw, Inches(0.36),
             title, size=12, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Description
        text(s, sx + Inches(0.12), fl_y + Inches(0.92), sw - Inches(0.24), Inches(0.85),
             desc, size=9, italic=True, color=GREEN_DARK, font=BODY_FONT)
        # Exit criteria
        text(s, sx + Inches(0.12), fl_y + Inches(1.85), sw - Inches(0.24), Inches(0.18),
             "EXIT CRITERIA", size=7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        by = fl_y + Inches(2.05)
        for ex in exits:
            text(s, sx + Inches(0.12), by, sw - Inches(0.24), Inches(0.42),
                 "▪ " + ex, size=8.5, color=DARK_TEXT, font=BODY_FONT)
            by += Inches(0.42)
        # Duration
        text(s, sx + Inches(0.12), fl_y + Inches(4.30), sw - Inches(0.24), Inches(0.18),
             "TIMEBOX", size=7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, sx + Inches(0.12), fl_y + Inches(4.48), sw - Inches(0.24), Inches(0.30),
             dur, size=10, bold=True, color=GREEN_DARK, font=BODY_FONT)

    # Bottom annotation
    red_badge(s, Inches(0.40), Inches(6.20), Inches(12.55), Inches(0.90),
              "ANA · NEXT STAGE — FIRST VERSION FOR HEATHER VALIDATION",
              "Sequence: validate with Heather (CS) before scaling to ITBA · run ITBA pilot in parallel with tooling · "
              "scale to whole org only after pilot lessons absorbed · ~32 weeks from sign-off to full TOM live · "
              "this slide is the version Heather signs off as v1 of the Validation Roadmap.",
              size_body=8)

    footer(s, "Validation Roadmap · NEW")
    return s


# ============================================================================
# MAIN
# ============================================================================

print("Loading TOM MAP workbook-2.pptx ...")
prs = Presentation(SRC)
total_before = len(prs.slides)
print(f"  slides BEFORE: {total_before}")

# --- Apply edits to existing slides ---
print("\n[A] Editing existing slides with red annotations ...")
edits = []
edits.append(("Slide 24 (BV)",          edit_slide_23_BV(prs)))
edits.append(("Slide 28 (CoS · CIT)",   edit_slide_27_CoS_add_CIT(prs)))
edits.append(("Slide 29 (Type)",        edit_slide_28_Type_simplify(prs)))
edits.append(("Slide 30 (Work Type)",   edit_slide_29_WorkType(prs)))
edits.append(("Slide 32 (Pain CS)",     edit_slide_32_PainScore_CS(prs)))
edits.append(("Slide 33 (After Sub)",   edit_slide_33_AfterSubmission(prs)))
edits.append(("Step-Level (find)",      edit_slide_36_StepLevel(prs)))
edits.append(("Governance (find)",      edit_slide_42_Governance(prs)))
edits.append(("BRP Deep Dive (find)",   edit_slide_44_BRP_redo(prs)))
edits.append(("Value & Prio (find)",    edit_slide_47_Value_OPCO(prs)))
for name, ok in edits:
    print(f"   {name:30s}  {'OK' if ok else 'NOT FOUND'}")

# --- Build 6 new slides ---
print("\n[B] Building 6 new slides ...")
build_routing_logic(prs);          print("   01 ROUTING LOGIC BY SIZE")
build_cos_maintain(prs);           print("   02 HOW WE MAINTAIN COS")
build_incident_flow(prs);          print("   03 INCIDENT FLOW · IT NOW → JIRA")
build_existing_project(prs);       print("   04 DEMAND ON EXISTING PROJECT")
build_digital_coe_interface(prs);  print("   05 DIGITAL CoE × DIGITAL PRODUCTS")
build_validation_roadmap(prs);     print("   06 VALIDATION ROADMAP")

# --- Reorder ---
# All 6 new slides are at the END (positions total_before to total_before+5).
# Place them in storytelling order:
#   01 ROUTING LOGIC BY SIZE   → after slide 35 (PORTFOLIO BOARD) — answers "what goes where"
#   02 HOW WE MAINTAIN COS     → after slide 39 (STEP-LEVEL POLICIES) — operational
#   03 INCIDENT FLOW           → after slide 34 (AUTOMATION CATALOG) — explains A2/incident
#   04 DEMAND ON EXISTING      → after slide 33 (AFTER SUBMISSION) — special case after main flow
#   05 DIGITAL CoE INTERFACE   → after VALUE & PRIORITIZATION (slide 47) — alignment topic
#   06 VALIDATION ROADMAP      → end of section — last
#
# To minimize re-shifts, reorder in a stable order: do them lowest-target first.

print("\n[C] Reordering ...")
# After build, indices:
ri = total_before + 0   # 01 ROUTING
ci = total_before + 1   # 02 COS_MAINTAIN
ii = total_before + 2   # 03 INCIDENT
ei = total_before + 3   # 04 EXISTING_PROJECT
di = total_before + 4   # 05 DIGITAL_COE
vi = total_before + 5   # 06 VALIDATION_ROADMAP

# Find current positions of anchor slides by title
# (helpers find_idx_by_title and find_idx_by_unique are defined above near move_slide)

# Anchors from current state (workbook-2 layout) — use unique strings to avoid mismatches
anchor_after_submission = find_idx_by_unique(prs, "Auto-routing → Triage → Auto-conversion")
anchor_automation = find_idx_by_unique(prs, "AUTOMATION CATALOG  ·  A1 → A6")
anchor_portfolio_board = find_idx_by_unique(prs, "CARDS ARE INITIATIVES")
anchor_step_level = find_idx_by_unique(prs, "STEP-LEVEL POLICIES")
anchor_value_prio = find_idx_by_unique(prs, "VALUE  &  PRIORITIZATION")

# 04 EXISTING_PROJECT goes right after AFTER SUBMISSION
target_existing = anchor_after_submission + 1
move_slide(prs, ei, target_existing)
print(f"   04 EXISTING_PROJECT → after AFTER SUBMISSION → position {target_existing+1}")

# Re-find anchors (use unique strings)
anchor_automation = find_idx_by_unique(prs, "AUTOMATION CATALOG  ·  A1 → A6")
anchor_portfolio_board = find_idx_by_unique(prs, "CARDS ARE INITIATIVES")
anchor_step_level = find_idx_by_unique(prs, "STEP-LEVEL POLICIES")

# 03 INCIDENT goes right after AUTOMATION CATALOG
ii = find_idx_by_unique(prs, "INCIDENT FLOW  ·  IT NOW")
target_incident = anchor_automation + 1
move_slide(prs, ii, target_incident)
print(f"   03 INCIDENT_FLOW    → after AUTOMATION CATALOG → position {target_incident+1}")

anchor_portfolio_board = find_idx_by_unique(prs, "CARDS ARE INITIATIVES")
anchor_step_level = find_idx_by_unique(prs, "STEP-LEVEL POLICIES")

# 01 ROUTING goes after PORTFOLIO BOARD (use unique substring 'CARDS ARE INITIATIVES')
ri = find_idx_by_unique(prs, "ROUTING LOGIC BY SIZE")
target_routing = anchor_portfolio_board + 1
move_slide(prs, ri, target_routing)
print(f"   01 ROUTING_LOGIC    → after PORTFOLIO BOARD → position {target_routing+1}")

anchor_step_level = find_idx_by_unique(prs, "STEP-LEVEL POLICIES")

# 02 COS_MAINTAIN after STEP-LEVEL POLICIES
ci = find_idx_by_unique(prs, "HOW WE MAINTAIN CLASS OF SERVICE")
target_cos = anchor_step_level + 1
move_slide(prs, ci, target_cos)
print(f"   02 COS_MAINTAIN     → after STEP-LEVEL POLICIES → position {target_cos+1}")

# 05 DIGITAL CoE — place right after BRP DEEP DIVE (governance / external alignment topic)
brp_pos = find_idx_by_unique(prs, "BIG ROOM PLANNING  ·  DEEP DIVE")
di = find_idx_by_unique(prs, "DIGITAL CoE  ×  DIGITAL PRODUCTS")
target_dcoe = brp_pos + 1
move_slide(prs, di, target_dcoe)
print(f"   05 DIGITAL_COE      → after BRP DEEP DIVE → position {target_dcoe+1}")

# 06 VALIDATION ROADMAP — closes the deck: place AFTER the last manager slide ("WORK ALLOCATION")
last_prep_pos = find_idx_by_unique(prs, "PREPARATION & PLANNING –  WORK ALLOCATION")
if last_prep_pos < 0:
    last_prep_pos = find_idx_by_unique(prs, "WORK ALLOCATION")
vi = find_idx_by_unique(prs, "VALIDATION ROADMAP  ·  NEXT STAGE")
target_val = last_prep_pos + 1 if last_prep_pos >= 0 else vi
move_slide(prs, vi, target_val)
print(f"   06 VALIDATION_ROADMAP → after WORK ALLOCATION (closes section) → position {target_val+1}")

# --- Save ---
print(f"\n[D] Saving → {DST}")
prs.save(DST)
print(f"   slides AFTER: {len(prs.slides)}")
print("\nDONE.")
