"""
Build IT Business Applications - Proposed Value Streams + Value Definition.pptx

Two sections:
  PART A — Previous proposals (orange = pre-existing combined CS+ITBA VS) + initial ITBA VS (blue)
    1.  Cover
    2.  8-card combined view (4 orange + 4 blue)
    3-6. Detail slides for the 4 BLUE ITBA value streams

  PART B — Deeper SAFe exercise for ITBA
    7.  Big Picture (4 OVS swim lanes)
    8.  Cards (4 OVS + 4 DVS + Enabling Team)
    9.  Value Stream Framing
    10. OVSs Multi-Domain View
    11-14. DVS 1-4 detail
    15. Enabling Team

Naming aligned with BLUE cards from prior work:
  • Engineering & Asset Management
  • Network Operations & Control
  • Corporate & Regulatory
  • Iberdrola Corporate Services
"""
from copy import deepcopy
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

SRC = "Material em andamento/Customer Services - Proposed Value Streams + Value Definition.pptx"
DST = "Material em andamento/IT Business Applications - Proposed Value Streams + Value Definition.pptx"

NAVY  = RGBColor(0x0F, 0x1B, 0x4F)
BLUE2 = RGBColor(0x2C, 0x57, 0x8B)
BLUE3 = RGBColor(0x5B, 0x7F, 0xB5)
BLUE4 = RGBColor(0xA5, 0xC2, 0xDA)
ORANGE = RGBColor(0xE6, 0x8B, 0x32)  # for the orange cards
GREEN = RGBColor(0x9C, 0xC4, 0x4A)
TEAL  = RGBColor(0x0B, 0x76, 0xA0)
RED   = RGBColor(0x98, 0x00, 0x03)
ENAB  = RGBColor(0x6C, 0x7A, 0x89)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK  = RGBColor(0x2C, 0x3E, 0x50)
GRAY  = RGBColor(0x95, 0xA5, 0xA6)
LIGHTBG = RGBColor(0xEE, 0xEE, 0xEE)

# Lane signature colors for the 4 ITBA VS
LANE_COLORS = [
    RGBColor(0x16, 0xA0, 0x85),  # Engineering & Asset Mgmt — green
    RGBColor(0xD3, 0x54, 0x00),  # Network Ops & Control — orange
    RGBColor(0x8E, 0x44, 0xAD),  # Corporate & Regulatory — purple
    RGBColor(0x2C, 0x57, 0x8B),  # Iberdrola — blue
]


# ---------------- helpers ----------------

def add_text(slide, x, y, w, h, text, *, size=10, bold=False, italic=False,
             color=DARK, name="Arial", align=PP_ALIGN.LEFT, anchor=None, wrap=True):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = wrap
    tf.margin_left = Inches(0.04); tf.margin_right = Inches(0.04)
    tf.margin_top = Inches(0.02); tf.margin_bottom = Inches(0.02)
    if anchor: tf.vertical_anchor = anchor
    for i, line in enumerate(text.split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run(); r.text = line
        r.font.name = name; r.font.size = Pt(size)
        r.font.bold = bold; r.font.italic = italic
        r.font.color.rgb = color
    return tb


def add_rect(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None):
    sh = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None: sh.fill.background()
    else: sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if line is None: sh.line.fill.background()
    else: sh.line.color.rgb = line; sh.line.width = Pt(0.5)
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        sh.adjustments[0] = radius
    return sh


def copy_bg(prs, src_idx, new_slide):
    src = prs.slides[src_idx]
    src_bg = src._element.find(qn('p:cSld') + '/' + qn('p:bg'))
    new_cSld = new_slide._element.find(qn('p:cSld'))
    if src_bg is not None and new_cSld is not None:
        existing = new_cSld.find(qn('p:bg'))
        if existing is not None: new_cSld.remove(existing)
        new_cSld.insert(0, deepcopy(src_bg))


def replace_textbox(shape, text, *, keep_font=True):
    tf = shape.text_frame
    if keep_font and tf.paragraphs and tf.paragraphs[0].runs:
        f0 = tf.paragraphs[0].runs[0].font
        sz = f0.size; bold = f0.bold; italic = f0.italic
        try: color = f0.color.rgb if f0.color and f0.color.rgb else None
        except Exception: color = None
        name = f0.name
    else:
        sz, bold, italic, color, name = None, None, None, None, "Arial"
    tf.clear()
    for i, line in enumerate(text.split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run(); r.text = line
        if name: r.font.name = name
        if sz: r.font.size = sz
        if bold is not None: r.font.bold = bold
        if italic is not None: r.font.italic = italic
        if color: r.font.color.rgb = color


def edit_by_idx(slide, mapping):
    by_idx = {i: s for i, s in enumerate(slide.shapes)}
    for i, text in mapping.items():
        if i in by_idx and by_idx[i].has_text_frame:
            replace_textbox(by_idx[i], text)


def delete_slide_safe(prs, idx):
    slide_id = list(prs.slides._sldIdLst)[idx]
    rId = slide_id.rId
    prs.part.drop_rel(rId)
    prs.slides._sldIdLst.remove(slide_id)


def clone_slide(prs, source_idx):
    src = prs.slides[source_idx]
    new = prs.slides.add_slide(src.slide_layout)
    for shp in list(new.shapes):
        sp = shp._element
        sp.getparent().remove(sp)
    for el in src.shapes._spTree:
        tag = el.tag.split('}', 1)[-1]
        if tag in ('nvGrpSpPr', 'grpSpPr'):
            continue
        new.shapes._spTree.append(deepcopy(el))
    copy_bg(prs, source_idx, new)
    return new


def move_slide(prs, old_idx, new_idx):
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    xml_slides.remove(slides[old_idx])
    xml_slides.insert(new_idx, slides[old_idx])


# ---------------- 8-card combined view (PART A overview) ----------------

def build_combined_view(slide):
    """ITBA initial proposal — only the 4 blue cards."""
    add_text(slide, 0.40, 0.10, 12.53, 0.38,
             "PROPOSED VALUE STREAMS — IT BUSINESS APPLICATIONS",
             size=20, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, 0.40, 0.58, 12.53, 0.23,
             "Initial value streams identified for ITBA  ·  detailed supporting slides follow",
             size=11, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    blue_cards = [
        ("Engineering & Asset Management",
         "DOMAIN · Asset Management",
         "Engineer / planner designs with reliable data",
         "ArcFM · ProjectWise · Bentley · CYME"),
        ("Network Operations & Control",
         "DOMAIN · Network Operations",
         "Grid operator prevents and minimizes outages",
         "Echelon · Cimplicity · SLAM · Switching Orders"),
        ("Corporate & Regulatory",
         "DOMAIN · Finance · HR · Procurement · Regulatory",
         "HR / employee / regulatory user self-service",
         "OpenText · Rate Case Advisor · Utilisphere"),
        ("Iberdrola Corporate Services",
         "DOMAIN · Finance · Procurement (Iberdrola)",
         "Spain finance / procurement / budget",
         "SAP FI · PO · Budget · S/4HANA"),
    ]

    n = 4
    margin = 0.50
    gap = 0.18
    total_w = 13.33 - 2 * margin
    card_w = (total_w - (n - 1) * gap) / n  # ≈ 2.96

    card_y = 1.80
    card_h = 4.50

    def render_card(x, y, w, h, title, domain, desc, apps):
        body = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                      Inches(x), Inches(y + 0.08),
                                      Inches(w), Inches(h - 0.08))
        body.fill.solid(); body.fill.fore_color.rgb = WHITE
        body.line.color.rgb = NAVY; body.line.width = Pt(0.75)
        body.adjustments[0] = 0.04

        hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                     Inches(x - 0.02), Inches(y),
                                     Inches(w + 0.04), Inches(0.85))
        hdr.fill.solid(); hdr.fill.fore_color.rgb = NAVY
        hdr.line.fill.background(); hdr.adjustments[0] = 0.10

        add_text(slide, x, y + 0.12, w, 0.65, title,
                 size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        inner_y = y + 1.00
        add_text(slide, x + 0.10, inner_y, w - 0.20, 0.22, domain,
                 size=8.5, bold=True, italic=True, color=NAVY, align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.10, inner_y + 0.28, w - 0.20, 1.20, desc,
                 size=11, italic=True, color=DARK, align=PP_ALIGN.CENTER)

        ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                    Inches(x + 0.25), Inches(inner_y + 1.55),
                                    Inches(w - 0.50), Inches(0.03))
        ln.fill.solid(); ln.fill.fore_color.rgb = NAVY
        ln.line.fill.background()

        add_text(slide, x + 0.10, inner_y + 1.65, w - 0.20, 0.22,
                 "EXAMPLE APPLICATIONS", size=9, bold=True, color=DARK, align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.10, inner_y + 1.92, w - 0.20, h - 1.92 - 1.10,
                 apps, size=10, color=DARK, align=PP_ALIGN.CENTER)

    for i, (t, d, ds, a) in enumerate(blue_cards):
        x = margin + i * (card_w + gap)
        render_card(x, card_y, card_w, card_h, t, d, ds, a)

    add_text(slide, 0.40, card_y + card_h + 0.20, 12.53, 0.30,
             "Each value stream is detailed in the next four slides — domain, applications, users, "
             "today's team touch-points, and vendor landscape.",
             size=10, italic=True, color=GRAY, align=PP_ALIGN.LEFT)


# ---------------- OVS detail slide (single-OVS, mirrors CS Big Picture style) ----------------

def build_ovs_detail_slide(slide, *, ovs_name, owner_label, trigger_text, chevrons_with_personas,
                           outcome_text, dvs_name, dvs_color, apps_text, other_ovs_text):
    """Single-OVS detail slide using the same visual language as the CS Big Picture image:
    header, light-grey chevron area, TRIGGER (left), personas above chevrons, OUTCOME (right),
    DVS band below, applications row, Enabling Team band."""
    # Header
    add_text(slide, 0.38, 0.17, 12.53, 0.38,
             f"OVS DETAIL — {ovs_name.upper()}",
             size=18, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, 0.40, 0.58, 12.53, 0.23,
             f"Example operational value stream  ·  {owner_label}",
             size=10.5, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # Light grey chevron area
    add_rect(slide, 0.27, 1.22, 12.66, 1.94, fill=LIGHTBG,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
    add_text(slide, 4.82, 1.28, 6.50, 0.22,
             f"OPERATIONAL VALUE STREAM — {ovs_name}",
             size=10.5, bold=True, color=NAVY, align=PP_ALIGN.LEFT)

    # TRIGGER (left)
    add_text(slide, 0.37, 1.80, 1.30, 0.21, "TRIGGER",
             size=10, bold=True, color=GRAY, align=PP_ALIGN.LEFT)
    add_text(slide, 0.37, 2.10, 1.30, 0.85, trigger_text,
             size=10, italic=True, color=DARK, align=PP_ALIGN.LEFT)

    # OUTCOME (right)
    add_text(slide, 11.60, 1.63, 1.30, 0.21, "OUTCOME",
             size=10, bold=True, color=GRAY, align=PP_ALIGN.LEFT)
    add_text(slide, 11.60, 1.92, 1.30, 1.00, outcome_text,
             size=10, italic=True, color=DARK, align=PP_ALIGN.LEFT)

    # Chevrons + personas
    n = len(chevrons_with_personas)
    chev_x_start = 1.85
    avail_w = 11.50 - chev_x_start  # ~9.65
    chev_overlap = 0.10
    chev_w = (avail_w + (n - 1) * chev_overlap) / n
    chev_h = 0.55
    chev_y = 2.30
    persona_y = 1.75

    for i, (step, persona) in enumerate(chevrons_with_personas):
        cx = chev_x_start + i * (chev_w - chev_overlap)
        # Persona above
        add_text(slide, cx, persona_y, chev_w - chev_overlap, 0.40, persona,
                 size=8, italic=True, color=GRAY, align=PP_ALIGN.CENTER)
        # Chevron
        sh = slide.shapes.add_shape(MSO_SHAPE.PENTAGON,
                                    Inches(cx), Inches(chev_y),
                                    Inches(chev_w), Inches(chev_h))
        sh.fill.solid(); sh.fill.fore_color.rgb = NAVY
        sh.line.color.rgb = NAVY; sh.line.width = Pt(0.5)
        tb = slide.shapes.add_textbox(Inches(cx), Inches(chev_y),
                                      Inches(chev_w), Inches(chev_h))
        tf = tb.text_frame; tf.word_wrap = True
        tf.margin_left = Inches(0.04); tf.margin_right = Inches(0.18)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = step
        r.font.name = "Arial"; r.font.size = Pt(9); r.font.bold = True; r.font.color.rgb = WHITE

    # DVS band area (light blue bg)
    dvs_bg = RGBColor(0xE6, 0xED, 0xF7)
    add_rect(slide, 0.27, 3.25, 12.66, 1.25, fill=dvs_bg,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
    add_text(slide, 4.22, 4.21, 5.26, 0.22,
             "DEVELOPMENT VALUE STREAM — who builds & evolves the solution",
             size=10.5, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

    # 1 DVS band spanning chevron area width
    dvs_x = chev_x_start + 0.05
    dvs_w = avail_w - 0.10
    dvs_y = 3.55
    add_rect(slide, dvs_x, dvs_y, dvs_w, 0.45,
             fill=dvs_color, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.30)
    add_text(slide, dvs_x, dvs_y + 0.10, dvs_w, 0.25,
             f"DVS — {dvs_name}",
             size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Applications row
    add_text(slide, 0.47, 4.91, 1.30, 0.21, "APPLICATIONS",
             size=10, bold=True, color=GRAY, align=PP_ALIGN.LEFT)
    add_rect(slide, dvs_x, 4.68, dvs_w, 0.66, fill=TEAL)
    add_text(slide, dvs_x, 4.68, dvs_w, 0.66, apps_text,
             size=9.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Enabling Team red band
    add_rect(slide, dvs_x, 5.64, dvs_w, 0.45,
             fill=RED, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.10)
    add_text(slide, dvs_x, 5.77, dvs_w, 0.20,
             "Enabling Team  ·  shared services consumed by this DVS  "
             "(SWF · demand intake · test/release coord · third-party risk gating · ALM tooling)",
             size=8.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Footer note: other OVSs in the same domain
    add_text(slide, 0.40, 6.25, 12.53, 0.40,
             f"Other operational value streams in this domain: {other_ovs_text}",
             size=9, italic=True, color=GRAY, align=PP_ALIGN.LEFT)
    add_text(slide, 0.40, 6.65, 12.53, 0.22,
             "Pattern: business operational value stream supported by one ITBA Development Value Stream + shared Enabling Team",
             size=8, italic=True, color=GRAY)


# ---------------- big picture (deeper SAFe ex.) ----------------

def build_big_picture(slide):
    add_text(slide, 0.38, 0.17, 12.53, 0.38,
             "DEEPER EXERCISE — IT BUSINESS APPLICATIONS BIG PICTURE",
             size=18, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, 0.40, 0.58, 12.53, 0.23,
             "Four parallel business value chains (OVS) · one IT Development Value Stream per OVS · one enabling team",
             size=10.5, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # Real business OVS examples (one per domain) — each domain has multiple OVSs
    LANES = [
        ("Engineering & Asset Management",
         "Owner: Engineering / Asset Mgmt",
         "Network expansion /\nnew customer connection",
         ["Plan &\nforecast", "Design &\nspecify", "Build /\nconstruct", "Energize &\npublish"],
         "Asset built\n· GIS updated\n· network ready",
         "Engineering & Asset Mgmt", "George Porto + crossover", "tbd",
         "ArcFM · ProjectWise · Bentley · CYME"),
        ("Network Operations & Control",
         "Owner: Network Operations",
         "Outage detected\n(sensor or customer)",
         ["Detect &\ntriage", "Dispatch &\nswitch", "Field repair &\nrestore", "Close & report\n(regulatory)"],
         "Service restored\n· grid stable\n· regulator informed",
         "Network Operations & Control", "Bruno Moreto + Andrew McGibbon", "tbd",
         "Echelon · Cimplicity · SLAM · Switching Orders"),
        ("Corporate & Regulatory",
         "Owner: Regulatory · Finance · Legal",
         "Rate case scheduled /\nnew tariff need",
         ["Plan strategy\n& gather", "Draft &\nreview", "Submit to\nregulator", "Implement\nrates"],
         "Rates approved\n· revenue protected\n· compliance kept",
         "Corporate & Regulatory", "Sarath Garikapati + Amit", "tbd",
         "OpenText · Rate Case Advisor · Utilisphere"),
        ("Iberdrola Corporate Services",
         "Owner: Iberdrola Spain · Avangrid Finance",
         "Iberdrola business\nunit identifies need",
         ["Requisition\n& approve", "Issue PO\n(SAP)", "Receive goods\n& process invoice", "Pay vendor\n& close"],
         "Goods received\n· vendor paid\n· books balanced",
         "Iberdrola Corporate Services", "Vanessa Colombo + Michael Dubois", "tbd",
         "SAP FI · PO · Budget · S/4HANA"),
    ]

    lane_top = 1.00
    lane_h = 1.40
    lane_gap = 0.10
    trigger_w = 1.55
    outcome_w = 1.55
    chev_area_x = 0.40 + trigger_w + 0.10
    chev_area_w = 13.33 - 0.40 - trigger_w - outcome_w - 0.30

    for i, (ovs, owner, trigger, chevrons, outcome, dvs, lead, fte, apps) in enumerate(LANES):
        y = lane_top + i * (lane_h + lane_gap)
        color = LANE_COLORS[i]

        add_rect(slide, 0.27, y - 0.05, 12.79, lane_h, fill=LIGHTBG,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)

        add_rect(slide, 0.40, y, trigger_w, 1.30, fill=NAVY,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
        add_text(slide, 0.42, y + 0.05, trigger_w - 0.04, 0.20,
                 "BUSINESS OVS", size=7, bold=True, color=BLUE4, align=PP_ALIGN.CENTER)
        add_text(slide, 0.42, y + 0.22, trigger_w - 0.04, 0.40, ovs,
                 size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, 0.42, y + 0.62, trigger_w - 0.04, 0.18, owner,
                 size=7, italic=True, color=BLUE4, align=PP_ALIGN.CENTER)
        add_text(slide, 0.42, y + 0.82, trigger_w - 0.04, 0.18,
                 "TRIGGER", size=7, bold=True, color=BLUE4, align=PP_ALIGN.CENTER)
        add_text(slide, 0.42, y + 1.00, trigger_w - 0.04, 0.30, trigger,
                 size=7.5, italic=True, color=WHITE, align=PP_ALIGN.CENTER)

        n = len(chevrons)
        chev_w = (chev_area_w - (n - 1) * (-0.10)) / n
        chev_h = 0.50
        chev_y = y + 0.10
        for j, step in enumerate(chevrons):
            cx = chev_area_x + j * (chev_w - 0.10)
            sh = slide.shapes.add_shape(MSO_SHAPE.PENTAGON,
                                        Inches(cx), Inches(chev_y),
                                        Inches(chev_w), Inches(chev_h))
            sh.fill.solid(); sh.fill.fore_color.rgb = NAVY
            sh.line.color.rgb = NAVY; sh.line.width = Pt(0.5)
            tb = slide.shapes.add_textbox(Inches(cx), Inches(chev_y),
                                          Inches(chev_w), Inches(chev_h))
            tf = tb.text_frame; tf.word_wrap = True
            tf.margin_left = Inches(0.04); tf.margin_right = Inches(0.18)
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
            r = p.add_run(); r.text = step
            r.font.name = "Arial"; r.font.size = Pt(8); r.font.bold = True; r.font.color.rgb = WHITE

        dvs_y = y + 0.68
        add_rect(slide, chev_area_x, dvs_y, chev_area_w, 0.30,
                 fill=color, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.30)
        add_text(slide, chev_area_x, dvs_y, chev_area_w, 0.30,
                 f"DVS — {dvs}  ·  lead: {lead}",
                 size=8.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        add_text(slide, chev_area_x, dvs_y + 0.32, chev_area_w, 0.40,
                 "APPS  ·  " + apps,
                 size=8, color=DARK, align=PP_ALIGN.CENTER)

        out_x = 13.33 - outcome_w - 0.40
        add_rect(slide, out_x, y, outcome_w, 1.30, fill=GREEN,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
        add_text(slide, out_x + 0.02, y + 0.05, outcome_w - 0.04, 0.20,
                 "OUTCOME", size=7, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        add_text(slide, out_x + 0.02, y + 0.30, outcome_w - 0.04, 0.95, outcome,
                 size=8.5, bold=True, color=DARK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    et_y = lane_top + 4 * (lane_h + lane_gap) + 0.10
    add_rect(slide, 0.40, et_y, 12.53, 0.50, fill=RED,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.20)
    add_text(slide, 0.40, et_y, 12.53, 0.50,
             "Enabling Team  ·  shared services consumed by the 4 DVS  "
             "(SWF · demand intake · test/release coord · third-party risk gating · ALM tooling)",
             size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    add_text(slide, 0.40, et_y + 0.55, 12.53, 0.18,
             "Pattern: standard operational value-stream visualization adapted for multi-OVS context",
             size=8, italic=True, color=GRAY)


# ---------------- cards (deeper SAFe ex.) ----------------

def build_cards_slide(slide):
    add_text(slide, 0.40, 0.10, 12.53, 0.38,
             "DEEPER EXERCISE — VALUE STREAM STRUCTURE",
             size=20, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, 0.40, 0.58, 12.53, 0.23,
             "IT BUSINESS APPLICATIONS  ·  4 OPERATIONAL VS + 4 DEVELOPMENT VS + 1 ENABLING TEAM",
             size=11, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    add_text(slide, 0.50, 1.05, 12.33, 0.22,
             "FOUR BUSINESS OPERATIONAL VALUE STREAMS  ·  ONE ITBA DVS PER OVS  ·  ONE PRODUCT MANAGER PER DVS",
             size=10, bold=True, color=NAVY)

    n = 4
    margin = 0.50
    gap = 0.12
    total_w = 13.33 - 2 * margin
    col_w = (total_w - (n - 1) * gap) / n

    ovs_top_y = 1.40
    ovs_top_h = 0.85
    dvs_y = ovs_top_y + ovs_top_h + 0.10
    dvs_h = 4.00

    cards = [
        ("Engineering & Asset Management",
         "Engineer / planner designs with reliable data",
         "Engineering & Asset Management",
         "George Porto (+ crossover with Bruno's team for ProjectWise/Bentley)",
         "size to validate",
         "DOMAIN · Asset Management",
         "Capture, model and publish reliable engineering and asset data",
         "ArcFM\nProjectWise\nBentley\nCYME",
         LANE_COLORS[0]),
        ("Network Operations & Control",
         "Grid operator prevents and minimizes outages",
         "Network Operations & Control",
         "Bruno Moreto + Andrew McGibbon",
         "size to validate",
         "DOMAIN · Network Operations",
         "Operate the systems that monitor, switch and control the network in real time",
         "Echelon\nCimplicity\nSLAM\nSwitching Orders",
         LANE_COLORS[1]),
        ("Corporate & Regulatory",
         "HR / employee / regulatory user self-service",
         "Corporate & Regulatory",
         "Sarath Garikapati + Amit",
         "size to validate",
         "DOMAIN · Finance · HR · Procurement · Regulatory",
         "Deliver regulatory and employee-facing technology under continuous compliance pressure",
         "OpenText\nRate Case Advisor\nUtilisphere",
         LANE_COLORS[2]),
        ("Iberdrola Corporate Services",
         "Spain finance / procurement / budget",
         "Iberdrola Corporate Services",
         "Vanessa Colombo + Michael Dubois",
         "size to validate",
         "DOMAIN · Finance · Procurement (Iberdrola)",
         "Run Spain-side SAP modules and keep US/Iberdrola corporate processes aligned",
         "SAP FI\nSAP PO\nSAP Budget\nS/4HANA",
         LANE_COLORS[3]),
    ]

    for i, c in enumerate(cards):
        (ovs, ovs_outcome, dvs, lead, fte, domain, desc, apps, color) = c
        x = margin + i * (col_w + gap)

        add_rect(slide, x, ovs_top_y, col_w, ovs_top_h,
                 fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
        add_text(slide, x + 0.05, ovs_top_y + 0.05, col_w - 0.10, 0.20,
                 "BUSINESS OVS", size=7.5, bold=True, color=BLUE4, align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.05, ovs_top_y + 0.22, col_w - 0.10, 0.30, ovs,
                 size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.10, ovs_top_y + 0.55, col_w - 0.20, 0.30, ovs_outcome,
                 size=7.5, italic=True, color=BLUE4, align=PP_ALIGN.CENTER)

        body = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                      Inches(x), Inches(dvs_y + 0.08),
                                      Inches(col_w), Inches(dvs_h - 0.08))
        body.fill.solid(); body.fill.fore_color.rgb = WHITE
        body.line.color.rgb = NAVY; body.line.width = Pt(0.5)
        body.adjustments[0] = 0.05

        hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                     Inches(x - 0.02), Inches(dvs_y),
                                     Inches(col_w + 0.04), Inches(0.78))
        hdr.fill.solid(); hdr.fill.fore_color.rgb = color
        hdr.line.fill.background(); hdr.adjustments[0] = 0.10

        add_text(slide, x, dvs_y + 0.08, col_w, 0.32, "ITBA DVS",
                 size=7.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(slide, x, dvs_y + 0.32, col_w, 0.40, dvs,
                 size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        inner_y = dvs_y + 0.92
        add_text(slide, x + 0.10, inner_y, col_w - 0.20, 0.20,
                 lead, size=8, bold=True, italic=True, color=DARK, align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.10, inner_y + 0.22, col_w - 0.20, 0.20,
                 domain, size=7.5, bold=True, italic=True, color=DARK, align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.10, inner_y + 0.42, col_w - 0.20, 0.55,
                 desc, size=9, italic=True, color=DARK, align=PP_ALIGN.CENTER)

        ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                    Inches(x + 0.20), Inches(inner_y + 1.00),
                                    Inches(col_w - 0.40), Inches(0.02))
        ln.fill.solid(); ln.fill.fore_color.rgb = NAVY
        ln.line.fill.background()

        add_text(slide, x + 0.10, inner_y + 1.08, col_w - 0.20, 0.20,
                 "ALL APPLICATIONS", size=8, bold=True, color=DARK, align=PP_ALIGN.CENTER)

        apps_y = inner_y + 1.30
        add_text(slide, x + 0.10, apps_y, col_w - 0.20, dvs_y + dvs_h - apps_y - 0.10,
                 apps, size=9, color=DARK, align=PP_ALIGN.CENTER)

    et_y = dvs_y + dvs_h + 0.18
    et_h = 0.85
    band = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  Inches(margin), Inches(et_y),
                                  Inches(total_w), Inches(et_h))
    band.fill.solid(); band.fill.fore_color.rgb = ENAB
    band.line.fill.background(); band.adjustments[0] = 0.20

    add_text(slide, margin + 0.20, et_y + 0.08, 3.50, 0.30,
             "ENABLING TEAM", size=11, bold=True, color=WHITE)
    add_text(slide, margin + 0.20, et_y + 0.36, 3.50, 0.45,
             "Shared services · supports the 4 DVS\nNo Product Manager · governed by Director ITBA",
             size=8.5, color=WHITE)
    add_text(slide, margin + 3.80, et_y + 0.10, total_w - 4.00, 0.30,
             "FOUNDATION SERVICES", size=9, bold=True, color=WHITE)
    add_text(slide, margin + 3.80, et_y + 0.32, total_w - 4.00, 0.55,
             "SWF (Software Factory) governance  ·  demand intake (PDD · Jira · IT Now)  ·  "
             "test & release coordination  ·  third-party risk gating  ·  vendor-mgmt cadence",
             size=9, color=WHITE)


# ---------------- content for the 4 BLUE VS detail slides (PART A) ----------------

BLUE_VS_CONTENT = [
    {  # Engineering & Asset Management
        0: "VALUE STREAM",
        1: "ENGINEERING & ASSET MANAGEMENT — RELIABLE ASSET DATA FOR PLANNERS AND ENGINEERS",
        3: "ITBA · Asset Management domain · lead crosses GIS + Engineering teams",
        5: "Initial value-stream definition  ·  domain anchor",
        8: ("Capture, model and publish engineering and asset data so planners and engineers can "
            "design and operate the network with reliable data."),
        11: "EXAMPLE APPLICATIONS",
        12: ("•  ArcFM (GIS / asset management)\n"
             "•  ProjectWise (engineering content / drawings)\n"
             "•  Bentley (CAD / engineering platform family)\n"
             "•  CYME (power systems analysis)"),
        15: "USERS / DEMAND SOURCES",
        16: ("•  Engineers — design new builds / modifications\n"
             "•  Planners — forecasts and network capacity\n"
             "•  Asset management — lifecycle decisions\n"
             "•  Field operations — asset updates from the field"),
        17: "Demand intake currently spread across Jira / PDD / IT Now",
        20: "TODAY'S TEAM TOUCH-POINTS",
        21: ("•  GIS & OSG team (George Porto) owns ArcFM, GIS web apps, SAP HANA\n"
             "•  G&E team (Bruno Moreto) currently owns ProjectWise / Bentley / Microstation\n"
             "•  CYME positioning to validate with team\n"
             "•  Cross-team app ownership — explicit value-stream framing helps"),
        22: "Apps cross today's team boundaries — value-stream view is the unifier",
        25: "VENDOR LANDSCAPE",
        26: ("•  Bentley (ProjectWise / Microstation)\n"
             "•  Esri / ArcFM partner (GIS)\n"
             "•  CYME vendor\n"
             "•  SWF (Software Factory) — execution support"),
        27: "Multiple vendors per domain — coordinated through ITBA leadership",
    },
    {  # Network Operations & Control
        0: "VALUE STREAM",
        1: "NETWORK OPERATIONS & CONTROL — GRID OPERATOR PREVENTS AND MINIMIZES OUTAGES",
        3: "ITBA · Network Operations domain",
        5: "Initial value-stream definition  ·  domain anchor",
        8: ("Operate the systems that monitor, switch and control the electricity and gas network "
            "in real time — preventing outages where possible and minimising their impact when they occur."),
        11: "EXAMPLE APPLICATIONS",
        12: ("•  Echelon (control / metering platform)\n"
             "•  Cimplicity (SCADA / HMI)\n"
             "•  SLAM (System Loading and Analysis Module — to confirm)\n"
             "•  Switching Orders (utility-specific switching workflow)"),
        15: "USERS / DEMAND SOURCES",
        16: ("•  Grid operators — real-time monitoring and switching decisions\n"
             "•  Network engineers — analysis and load planning\n"
             "•  Field crews — switching execution\n"
             "•  Regulators — outage reporting"),
        17: "Real-time / safety-critical  ·  high availability requirements",
        20: "TODAY'S TEAM TOUCH-POINTS",
        21: ("•  G&E Operations team (Bruno Moreto + Andrew McGibbon) is the closest current home\n"
             "•  Some apps may sit in OT (Operational Technology) scope — boundary to confirm\n"
             "•  Coordination with Field Operations VS (cross-stream)"),
        22: "Boundary with OT to confirm — some apps may sit outside ITBA scope",
        25: "VENDOR LANDSCAPE",
        26: ("•  Echelon vendor\n"
             "•  GE (Cimplicity)\n"
             "•  SLAM vendor (to confirm)\n"
             "•  Internal switching workflow tooling"),
        27: "Vendor mix to validate — some grid platforms managed cross-Avangrid",
    },
    {  # Corporate & Regulatory
        0: "VALUE STREAM",
        1: "CORPORATE & REGULATORY — HR / EMPLOYEE / REGULATORY USER SELF-SERVICE",
        3: "ITBA · Corporate Applications · Sarath Garikapati + Amit",
        5: "Initial value-stream definition  ·  domain anchor",
        8: ("Deliver and maintain regulatory and employee-facing technology across Avangrid US — "
            "operating under continuous compliance pressure."),
        11: "EXAMPLE APPLICATIONS",
        12: ("•  OpenText (content / records management)\n"
             "•  Rate Case Advisor (regulatory rate cases)\n"
             "•  Utilisphere (utility platform)\n"
             "•  Plus broader corporate applications portfolio"),
        15: "USERS / DEMAND SOURCES",
        16: ("•  HR / Finance / Legal / Compliance\n"
             "•  Regulators — rate cases · mandated reporting\n"
             "•  Employee population — enablement & self-service\n"
             "•  Continuous new requirements"),
        17: "Compliance-driven  ·  always-priority cadence",
        20: "TODAY'S TEAM TOUCH-POINTS",
        21: ("•  Corporate Applications team (Sarath Garikapati + Amit)\n"
             "•  5–6 PMs + 2 tech leads in current composition\n"
             "•  Largest sub-team in ITBA  ·  ~8 FTE total"),
        22: "PM-heavy team today — value-stream framing helps clarify outcome ownership",
        25: "VENDOR LANDSCAPE",
        26: ("•  OpenText vendor\n"
             "•  Accenture (corporate apps + projects)\n"
             "•  SWF (Software Factory)\n"
             "•  Regulatory tooling vendors"),
        27: "Accenture is a key partner for corporate scope",
    },
    {  # Iberdrola Corporate Services
        0: "VALUE STREAM",
        1: "IBERDROLA CORPORATE SERVICES — SPAIN FINANCE / PROCUREMENT / BUDGET",
        3: "ITBA · Global SAP · Vanessa Colombo + Michael Dubois",
        5: "Initial value-stream definition  ·  domain anchor",
        8: ("Run Spain-side SAP modules and keep US/Iberdrola corporate processes aligned — "
            "balancing incidents, small demands, and major projects under hiring constraints."),
        11: "EXAMPLE APPLICATIONS",
        12: ("•  SAP FI (Finance) — Spain side\n"
             "•  SAP PO (Purchase Orders) — Spain side\n"
             "•  SAP Budget control\n"
             "•  S/4HANA (Iberdrola roadmap)"),
        15: "USERS / DEMAND SOURCES",
        16: ("•  Iberdrola corporate (Spain) — process changes\n"
             "•  Avangrid Finance — US-side integration needs\n"
             "•  Incidents (IT Now)\n"
             "•  Small Demands · Major projects (PDD)"),
        17: "Cross-org coordination  ·  transport gating",
        20: "TODAY'S TEAM TOUCH-POINTS",
        21: ("•  Global SAP team (Vanessa Colombo + Michael Dubois)\n"
             "•  ~5 FTE total\n"
             "•  Hiring lag (~1 year) flagged as top risk\n"
             "•  Knowledge-transfer loss is a stated concern"),
        22: "Smallest sub-team in ITBA — capacity is the chronic constraint",
        25: "VENDOR LANDSCAPE",
        26: ("•  Accenture (Global SAP support)\n"
             "•  SAP partner for Spain modules\n"
             "•  Iberdrola Spain integration teams (peer)\n"
             "•  SWF (Software Factory)"),
        27: "Iberdrola peer teams are an additional execution channel",
    },
]


# ---------------- DVS detail content (PART B — deeper SAFe ex.) ----------------

DVS_CONTENT = [
    {  # DVS 1 — Engineering & Asset Management
        0: "DEVELOPMENT VALUE STREAM (DVS) 1 / 4",
        1: "ENGINEERING & ASSET MANAGEMENT — RELIABLE DATA FOR PLANNERS AND ENGINEERS",
        3: "ITBA · George Porto (lead) + crossover with G&E team for ProjectWise/Bentley",
        5: "Run · Grow · Transform mix · to validate with team",
        8: ("Build and evolve the platform that captures, models and publishes engineering & asset data — "
            "so planners and engineers can design and operate the network with reliable data."),
        11: "APPLICATIONS SUPPORTED",
        12: ("•  ArcFM (GIS / asset management)\n"
             "•  ProjectWise (engineering content / drawings)\n"
             "•  Bentley (CAD / engineering platform family)\n"
             "•  CYME (power systems analysis)\n"
             "•  Plus supporting GIS / data integrations from KB scope"),
        15: "DEMAND SOURCES — who sends the work",
        16: ("•  Engineering — new builds · network modifications\n"
             "•  Planning — forecasts · capacity planning\n"
             "•  Asset Management — lifecycle decisions\n"
             "•  Field operations — asset updates from the field\n"
             "•  Vendor-driven upgrades (ProjectWise · Bentley)"),
        17: "Routed via the Jira request type for this DVS",
        20: "AVANGRID CORE TEAM",
        21: ("•  Product Manager (NEW)\n"
             "•  Tech Lead — George Porto\n"
             "•  Crossover engineers from current G&E team for Bentley/ProjectWise scope\n"
             "•  Squad size · to validate"),
        22: "Cross-team scope — clean PM accountability requires crossing today's GIS/G&E line",
        25: "VENDOR EXTENSION (delivery)",
        26: ("•  Bentley (ProjectWise / Microstation)\n"
             "•  Esri / ArcFM partner\n"
             "•  CYME vendor\n"
             "•  SWF (Software Factory) — execution muscle"),
        27: "Avangrid = product owner · vendor = delivery muscle",
    },
    {  # DVS 2 — Network Operations & Control
        0: "DEVELOPMENT VALUE STREAM (DVS) 2 / 4",
        1: "NETWORK OPERATIONS & CONTROL — KEEP THE GRID UP · PREVENT AND MINIMIZE OUTAGES",
        3: "ITBA · Bruno Moreto + Andrew McGibbon (lead)",
        5: "Run · Grow · Transform mix · to validate with team",
        8: ("Build and evolve the systems that monitor, switch and control the network in real time — "
            "preventing outages and minimising their impact when they happen."),
        11: "APPLICATIONS SUPPORTED",
        12: ("•  Echelon\n"
             "•  Cimplicity (SCADA / HMI)\n"
             "•  SLAM (System Loading and Analysis Module — to confirm)\n"
             "•  Switching Orders\n"
             "•  Possibly broader OT-side platforms — boundary to confirm"),
        15: "DEMAND SOURCES — who sends the work",
        16: ("•  Grid operators — real-time monitoring needs\n"
             "•  Network engineers — analysis & load planning\n"
             "•  Field crews — switching execution\n"
             "•  Regulators — outage reporting cadence"),
        17: "Real-time / safety-critical · high-availability constraints apply",
        20: "AVANGRID CORE TEAM",
        21: ("•  Product Manager (NEW)\n"
             "•  Tech Lead — Bruno Moreto / Andrew McGibbon\n"
             "•  Business Analyst — grid operations SME\n"
             "•  Squad size · to validate"),
        22: "OT/IT boundary to confirm — some apps may sit in OT scope, not ITBA",
        25: "VENDOR EXTENSION (delivery)",
        26: ("•  Echelon vendor\n"
             "•  GE (Cimplicity)\n"
             "•  SLAM vendor (to confirm)\n"
             "•  SWF (Software Factory)"),
        27: "Avangrid = product owner · vendor = delivery muscle",
    },
    {  # DVS 3 — Corporate & Regulatory
        0: "DEVELOPMENT VALUE STREAM (DVS) 3 / 4",
        1: "CORPORATE & REGULATORY — REGULATORY COMPLIANCE · EMPLOYEE ENABLEMENT",
        3: "ITBA · Sarath Garikapati + Amit (lead)",
        5: "Run · Grow · Transform mix · to validate with team",
        8: ("Deliver and maintain regulatory and employee-facing technology across Avangrid US — "
            "under continuous compliance pressure with always-priority cadence."),
        11: "APPLICATIONS SUPPORTED",
        12: ("•  OpenText (content / records management)\n"
             "•  Rate Case Advisor (regulatory rate cases)\n"
             "•  Utilisphere (utility platform)\n"
             "•  Corporate applications portfolio\n"
             "•  Network applications · employee portals · regulatory reporting tools"),
        15: "DEMAND SOURCES — who sends the work",
        16: ("•  HR / Finance / Legal / Compliance — corporate process changes\n"
             "•  Regulators — mandated reporting cadence\n"
             "•  Employee population — enablement & training\n"
             "•  Continuous new requirements (always priority)"),
        17: "Routed via the Jira request type for this DVS",
        20: "AVANGRID CORE TEAM",
        21: ("•  Product Manager (NEW)\n"
             "•  Tech Lead — Sarath Garikapati / Amit\n"
             "•  5–6 PMs + 2 tech leads (current composition)\n"
             "•  Squad size · to validate (~8 FTE today)"),
        22: "Largest sub-team — possible internal sub-streams to validate",
        25: "VENDOR EXTENSION (delivery)",
        26: ("•  OpenText vendor\n"
             "•  Accenture (corporate apps + projects)\n"
             "•  SWF (Software Factory)\n"
             "•  Regulatory tooling vendors"),
        27: "Avangrid = product owner · vendor = delivery muscle",
    },
    {  # DVS 4 — Iberdrola Corporate Services
        0: "DEVELOPMENT VALUE STREAM (DVS) 4 / 4",
        1: "IBERDROLA CORPORATE SERVICES — SPAIN-SIDE STABILITY · IBERDROLA ALIGNMENT",
        3: "ITBA · Vanessa Colombo + Michael Dubois (lead)",
        5: "Run · Grow · Transform mix · to validate with team",
        8: ("Run Spain-side SAP modules and keep US/Iberdrola corporate processes aligned — "
            "balancing incidents, small demands, and major projects under hiring constraints."),
        11: "APPLICATIONS SUPPORTED",
        12: ("•  SAP FI (Finance) — Spain side\n"
             "•  SAP PO (Purchase Orders) — Spain side\n"
             "•  SAP Budget control\n"
             "•  S/4HANA (Iberdrola roadmap)\n"
             "•  Cross-org transport coordination"),
        15: "DEMAND SOURCES — who sends the work",
        16: ("•  Iberdrola corporate (Spain) — process changes\n"
             "•  Avangrid Finance — US-side integration needs\n"
             "•  Incidents (IT Now) — Spain SAP module issues\n"
             "•  Small Demands & Major projects (PDD)"),
        17: "Routed via the Jira request type for this DVS",
        20: "AVANGRID CORE TEAM",
        21: ("•  Product Manager (NEW)\n"
             "•  Tech Lead — Vanessa Colombo / Michael Dubois\n"
             "•  Business Analyst — SAP FI/PO SME\n"
             "•  Squad size · to validate (~5 FTE today)"),
        22: "Hiring lag (~1 year) + knowledge-transfer loss flagged as top risks",
        25: "VENDOR EXTENSION (delivery)",
        26: ("•  Accenture (Global SAP support)\n"
             "•  SAP partner for Spain modules\n"
             "•  Iberdrola Spain integration teams (peer)\n"
             "•  SWF (Software Factory)"),
        27: "Avangrid = product owner · vendor = delivery muscle",
    },
]


# ---------------- main ----------------

def main():
    prs = Presentation(SRC)

    # 1. Drop Value Lens (was slide 15, idx 14)
    delete_slide_safe(prs, 14)

    # 2. Clone slide 13 (CCC VS detail) to get a 4th slot for Iberdrola VS detail
    clone_slide(prs, 13)
    # Now: 14 original-1 = 13 + 1 clone = 14 slides total

    # 3. Edit cover (slide 0)
    cover = prs.slides[0]
    for sh in cover.shapes:
        if sh.has_text_frame and "PROPOSED VALUE STREAMS" in sh.text_frame.text and "CUSTOMER SERVICES" in sh.text_frame.text:
            replace_textbox(sh,
                "PROPOSED VALUE STREAMS – IT BUSINESS APPLICATIONS\n"
                "Prior proposals + initial ITBA value streams  ·  then a deeper exercise (4 OVS + 4 DVS + Enabling Team)\n"
                "April 2026")

    # 4. Rebuild slide 1 (was M2C Big Picture) → ITBA Big Picture (deeper SAFe ex.)
    big = prs.slides[1]
    for sh in list(big.shapes):
        sh._element.getparent().remove(sh._element)
    copy_bg(prs, 1, big)
    build_big_picture(big)

    # 5. Rebuild slide 2 (was Cards P1) → ITBA Cards (deeper SAFe ex.)
    cards = prs.slides[2]
    for sh in list(cards.shapes):
        sh._element.getparent().remove(sh._element)
    copy_bg(prs, 2, cards)
    build_cards_slide(cards)

    # 6. Edit slide 3 (Framing) — now reflects ITBA multi-OVS reality
    edit_by_idx(prs.slides[3], {
        0: "VALUE STREAM FRAMING",
        1: "MULTIPLE OPERATIONAL VS  ·  ONE DVS PER OVS  ·  SHARED ENABLING TEAM",
        3: "Director ITBA · Carlos Garbin  ·  4 DVS Product Managers (NEW)",
        5: "Multi-OVS pattern · 1:1 DVS alignment",
        8: ("ITBA does not have a single customer-facing flow like Meter-to-Cash. Instead it serves "
            "four parallel business value chains, each with its own outcome. The deeper exercise pairs "
            "each OVS with one DVS so accountability is explicit."),
        11: "BUSINESS OPERATIONAL VALUE STREAMS (4)",
        12: ("•  Engineering & Asset Management\n"
             "    – Outcome: reliable asset data · planners trust it\n"
             "•  Network Operations & Control\n"
             "    – Outcome: outages prevented · grid stable\n"
             "•  Corporate & Regulatory\n"
             "    – Outcome: compliance kept · employees enabled\n"
             "•  Iberdrola Corporate Services\n"
             "    – Outcome: Iberdrola alignment · Spain-side stability"),
        15: "ITBA DEVELOPMENT VALUE STREAMS (4)",
        16: ("•  DVS 1 — Engineering & Asset Management\n"
             "•  DVS 2 — Network Operations & Control\n"
             "•  DVS 3 — Corporate & Regulatory\n"
             "•  DVS 4 — Iberdrola Corporate Services\n"
             "+ Enabling Team — Foundation services (SWF · intake · test · risk)"),
        17: "Sizes to validate — current sub-teams may need to recombine across DVS boundaries",
        20: "MAPPING — OVS → DVS → APPS",
        21: ("•  Engineering & Asset Mgmt → Engineering & Asset Mgmt DVS · ArcFM · ProjectWise · Bentley · CYME\n"
             "•  Network Ops & Control    → Network Ops & Control DVS    · Echelon · Cimplicity · SLAM · Switching\n"
             "•  Corporate & Regulatory   → Corporate & Regulatory DVS   · OpenText · Rate Case Advisor · Utilisphere\n"
             "•  Iberdrola Corp Services  → Iberdrola Corp Services DVS  · SAP FI · PO · Budget · S/4HANA\n"
             "•  Enabling Team supports all 4 with: SWF · intake · test/release · third-party risk"),
        22: "1 OVS : 1 DVS — clean accountability · enabling team is shared support",
        25: "WHY THIS WORKS",
        26: ("•  4 dedicated Product Managers (one per DVS)\n"
             "•  Foundation services run as small enabling team — no extra PM\n"
             "•  Director ITBA (Carlos) = existing role · no new layer\n"
             "•  Some apps cross today's team boundaries — explicit value-stream framing is the unifier\n"
             "•  Vendor extension model preserved per DVS"),
        27: "Same headcount base · clearer ownership · explicit business alignment",
    })

    # 7. Slide 4 (was OVS detail M2C) → OVS multi-domain view
    edit_by_idx(prs.slides[4], {
        0: "OPERATIONAL VALUE STREAMS — MULTI-DOMAIN VIEW",
        1: "ITBA SUPPORTS FOUR BUSINESS VALUE CHAINS — EACH OWNED BY A BUSINESS STAKEHOLDER",
        3: "ITBA · Carlos Garbin · 4 OVSs supported (1:1 with DVS)",
        5: "Multi-OVS pattern  ·  business owns each OVS",
        8: ("ITBA serves four distinct operational value streams. Each is owned by the business, "
            "not by IT. ITBA's job is to provide the technology pipeline (DVS) that lets each business "
            "OVS deliver its outcome."),
        11: "OVS 1 — ENGINEERING & ASSET MANAGEMENT",
        12: ("•  Owner: Engineering / Asset Management leadership\n"
             "•  Customer: Engineers · planners · asset managers\n"
             "•  Trigger: Asset created/modified · network re-survey · planning need\n"
             "•  Steps: Design/plan → Survey & model → Publish to network → Consume in planning\n"
             "•  Outcome: Reliable asset data · planners trust it\n"
             "•  Supporting DVS: Engineering & Asset Management  ·  George Porto + crossover"),
        15: "OVS 2 — NETWORK OPERATIONS & CONTROL",
        16: ("•  Owner: Network Operations leadership\n"
             "•  Customer: Grid operators · network engineers · field crews\n"
             "•  Trigger: Grid event / switching need / monitoring alert\n"
             "•  Steps: Detect/monitor → Decide & switch → Execute/restore → Update logs & close\n"
             "•  Outcome: Outages prevented · grid stable\n"
             "•  Supporting DVS: Network Operations & Control  ·  Bruno Moreto + Andrew McGibbon"),
        17: "Owners listed are business stakeholders — ITBA's DVS supports them",
        20: "OVS 3 — CORPORATE & REGULATORY",
        21: ("•  Owner: HR · Finance · Legal · Compliance leadership\n"
             "•  Customer: All Avangrid US employees · regulators · auditors\n"
             "•  Trigger: Regulatory ask · employee enablement need · upgrade cycle\n"
             "•  Steps: Capture requirement → Implement & test → Deploy → Comply/report\n"
             "•  Outcome: Compliance kept · employees enabled\n"
             "•  Supporting DVS: Corporate & Regulatory  ·  Sarath Garikapati + Amit"),
        22: "Each OVS has its own demand pattern, stakeholders and SLA shape",
        25: "OVS 4 — IBERDROLA CORPORATE SERVICES",
        26: ("•  Owner: Iberdrola corporate (Spain) · Avangrid Finance\n"
             "•  Customer: Iberdrola corporate stakeholders · Avangrid Finance integration\n"
             "•  Trigger: Iberdrola process change · cross-org reporting · transport coordination\n"
             "•  Steps: Receive change → Align FI/PO/Budget → Coordinate transports → Validate\n"
             "•  Outcome: Iberdrola alignment · Spain-side stability\n"
             "•  Supporting DVS: Iberdrola Corporate Services  ·  Vanessa Colombo + Michael Dubois"),
        27: "All 4 OVSs supported by the same enabling team for shared services",
    })

    # 8. Slides 5-8 (DVS 1-4 detail)
    edit_by_idx(prs.slides[5], DVS_CONTENT[0])
    edit_by_idx(prs.slides[6], DVS_CONTENT[1])
    edit_by_idx(prs.slides[7], DVS_CONTENT[2])
    edit_by_idx(prs.slides[8], DVS_CONTENT[3])

    # 9. Slide 9 (Enabling Team)
    edit_by_idx(prs.slides[9], {
        0:  "ENABLING TEAM",
        1:  "FOUNDATION SERVICES — SHARED CAPABILITIES THAT KEEP THE 4 ITBA DVS RUNNING",
        3:  "ITBA · governed by Director ITBA (Carlos Garbin) + cross-DVS coordination",
        5:  "No PM · no DVS-level backlog · operates as a small shared-service team",
        8:  ("Provide the technical and process foundation — Software Factory governance, "
             "demand intake, test/release coordination, third-party risk gating — that the 4 ITBA DVS rely on."),
        11: "FOUNDATION SERVICES",
        12: ("•  SWF (Software Factory) governance — used across all 4 DVS\n"
             "•  Demand intake orchestration (PDD · Jira · IT Now)\n"
             "•  Test & release coordination (cross-DVS)\n"
             "•  Third-party risk assessment gating (ProjectWise · Microstation · others)\n"
             "•  Vendor-management cadence (Accenture · SWF · business vendors)\n"
             "•  ALM tooling administration"),
        15: "WHO USES THESE SERVICES",
        16: ("•  All 4 ITBA DVS — for SWF capacity · environments · test cycles\n"
             "•  Engineering & Asset Mgmt DVS especially — third-party risk gating for vendor upgrades\n"
             "•  Iberdrola Corp Services DVS — transport coordination · Iberdrola integration\n"
             "•  Director ITBA — portfolio view across DVS\n"
             "•  Optionally shared with CS Apps enabling team (Digital Products consolidation)"),
        17: "Foundation services support all 4 ITBA DVS — they do not constitute a separate product line",
        20: "TEAM SHAPE",
        21: ("•  Coordination led by Director ITBA (Carlos Garbin)\n"
             "•  SWF governance lead\n"
             "•  Demand intake / portfolio admin\n"
             "•  Test & risk coordinators — embedded in DVS where feasible\n"
             "•  No dedicated Product Manager\n"
             "Squad size · to validate"),
        22: "Why not a DVS: at this headcount, productizing foundation services would create overhead without internal customer roadmap",
        25: "VENDOR EXTENSION",
        26: ("•  SWF (Software Factory) vendor — capacity & tooling\n"
             "•  Accenture — for cross-DVS major projects\n"
             "•  Atlassian Jira admins\n"
             "•  Third-party risk assessment vendors"),
        27: "Avangrid governs the team operationally · vendors execute specific capabilities",
    })

    # 10. Slide 10 (was Cards P2 - 3 VS) → Combined 8-card view
    combined = prs.slides[10]
    for sh in list(combined.shapes):
        sh._element.getparent().remove(sh._element)
    copy_bg(prs, 10, combined)
    build_combined_view(combined)

    # 11. Slides 11-13 (3 VS detail) + slide 14 (cloned) → 4 BLUE ITBA VS detail
    edit_by_idx(prs.slides[11], BLUE_VS_CONTENT[0])
    edit_by_idx(prs.slides[12], BLUE_VS_CONTENT[1])
    edit_by_idx(prs.slides[13], BLUE_VS_CONTENT[2])
    edit_by_idx(prs.slides[14], BLUE_VS_CONTENT[3])

    # 12. Reorder so PART A (combined view + 4 blue VS detail) comes first
    # Current order:
    #  0=cover, 1=Big Pic, 2=Cards, 3=Framing, 4=OVS multi, 5-8=DVS, 9=Enabling,
    #  10=Combined, 11-14=4 blue VS detail
    # Target order:
    #  0=cover, 1=Combined, 2-5=4 blue VS, 6=Big Pic, 7=Cards, 8=Framing, 9=OVS multi, 10-13=DVS, 14=Enabling

    move_slide(prs, 10, 1)   # Combined → 1
    move_slide(prs, 11, 2)   # blue VS 1 → 2
    move_slide(prs, 12, 3)   # blue VS 2 → 3
    move_slide(prs, 13, 4)   # blue VS 3 → 4
    move_slide(prs, 14, 5)   # blue VS 4 → 5

    # 13. Append 4 NEW OVS detail slides (one example OVS per domain)
    OVS_DETAILS = [
        dict(  # Engineering & Asset Mgmt — Network Build
            ovs_name="Network Build (New Connection)",
            owner_label="Owner: Engineering / Asset Management",
            trigger_text="Network expansion need or new customer connection request",
            chevrons_with_personas=[
                ("Plan &\nforecast", "Planner"),
                ("Survey\nsite", "Surveyor + engineer"),
                ("Design &\nspecify", "Engineer"),
                ("Procure\nmaterials", "Procurement"),
                ("Build /\nconstruct", "Field crews"),
                ("Energize", "Field crews + grid ops"),
                ("Document\n& publish", "Engineer"),
            ],
            outcome_text="Asset built · GIS updated · network ready for service",
            dvs_name="Engineering & Asset Management",
            dvs_color=LANE_COLORS[0],
            apps_text="ArcFM (asset)  ·  CYME (analysis)  ·  Bentley / Microstation (design)  ·  ProjectWise (drawings)",
            other_ovs_text="Asset Replacement Cycle  ·  Network Inspection  ·  Capital Planning  ·  Asset Decommissioning",
        ),
        dict(  # Network Operations & Control — Outage Response
            ovs_name="Outage Response",
            owner_label="Owner: Network Operations",
            trigger_text="Outage detected (SCADA sensor or customer call) or grid event",
            chevrons_with_personas=[
                ("Detect\n& alert", "SCADA + customer"),
                ("Triage &\nprioritize", "Grid ops"),
                ("Dispatch\ncrews", "Dispatcher"),
                ("Switch\nnetwork", "Grid operator"),
                ("Field\nrepair", "Field crews"),
                ("Restore\nservice", "Grid ops"),
                ("Close &\nreport", "Operator + regulator"),
            ],
            outcome_text="Service restored · grid stable · regulator informed · root cause captured",
            dvs_name="Network Operations & Control",
            dvs_color=LANE_COLORS[1],
            apps_text="Cimplicity (SCADA)  ·  Echelon (control)  ·  Switching Orders (workflow)  ·  SLAM (load analysis)",
            other_ovs_text="Grid Optimization  ·  Routine Switching  ·  Load Balancing  ·  Storm Restoration",
        ),
        dict(  # Corporate & Regulatory — Rate Case Filing
            ovs_name="Rate Case Filing",
            owner_label="Owner: Regulatory Affairs · Finance · Legal",
            trigger_text="Rate case scheduled by regulator or new tariff strategy needed",
            chevrons_with_personas=[
                ("Plan\nstrategy", "Regulatory + Finance"),
                ("Gather\ndata", "Multi-team"),
                ("Draft\nfiling", "Regulatory + tool"),
                ("Internal\nreview", "Legal + leadership"),
                ("Submit to\nregulator", "Regulatory"),
                ("Respond to\ndiscovery", "Regulatory + Legal"),
                ("Implement\nrates", "Finance + IT"),
            ],
            outcome_text="Rates approved · revenue protected · compliance maintained",
            dvs_name="Corporate & Regulatory",
            dvs_color=LANE_COLORS[2],
            apps_text="Rate Case Advisor (drafting)  ·  OpenText (documents)  ·  Utilisphere (utility data)  ·  internal financial systems",
            other_ovs_text="Source-to-Pay  ·  Hire-to-Retire  ·  Regulatory Reporting Cycle  ·  Employee Self-Service",
        ),
        dict(  # Iberdrola Corporate Services — Procure-to-Pay (Spain)
            ovs_name="Procure-to-Pay (Spain)",
            owner_label="Owner: Iberdrola Spain · Avangrid Finance",
            trigger_text="Iberdrola business unit identifies a need (good or service)",
            chevrons_with_personas=[
                ("Requisition", "Requester"),
                ("Approve\nworkflow", "Approver"),
                ("Issue PO\n(SAP)", "Procurement"),
                ("Receive\ngoods", "Receiving"),
                ("Process\ninvoice", "FI / AP"),
                ("Pay\nvendor", "FI / Treasury"),
                ("Close &\nreconcile", "Finance close"),
            ],
            outcome_text="Goods received · vendor paid · books balanced · Iberdrola alignment",
            dvs_name="Iberdrola Corporate Services",
            dvs_color=LANE_COLORS[3],
            apps_text="SAP PO (purchase)  ·  SAP FI (finance)  ·  SAP Budget (control)  ·  S/4HANA (platform)",
            other_ovs_text="Period Close (Spain)  ·  Plan-to-Budget (Spain)  ·  Financial Reporting (Spain)  ·  Treasury Mgmt",
        ),
    ]

    # Use slide 6 (Big Picture, Em Branco layout) as cloning source for the new slides
    new_ovs_idxs = []
    for ovs_data in OVS_DETAILS:
        new = prs.slides.add_slide(prs.slides[6].slide_layout)
        for shp in list(new.shapes):
            shp._element.getparent().remove(shp._element)
        copy_bg(prs, 6, new)
        build_ovs_detail_slide(new, **ovs_data)
        new_ovs_idxs.append(len(prs.slides) - 1)

    # Insert the 4 new OVS detail slides at positions 10..13
    # (between OVS Multi-View at 9 and DVS detail at 10)
    # Currently new slides are at 15, 16, 17, 18. Move sequentially:
    move_slide(prs, 15, 10)
    move_slide(prs, 16, 11)
    move_slide(prs, 17, 12)
    move_slide(prs, 18, 13)

    prs.save(DST)
    print(f"Saved → {DST}")
    print(f"Total slides: {len(prs.slides)}")
    for i, s in enumerate(prs.slides, 1):
        title = ""
        for sh in s.shapes:
            if sh.has_text_frame and sh.text_frame.text.strip():
                title = sh.text_frame.text.strip().split('\n')[0][:80]
                break
        print(f"  {i:2d}. {title}")


if __name__ == "__main__":
    main()
