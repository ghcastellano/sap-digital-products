"""
Build the IT Business Applications value-streams deck.
Mirrors the CS Apps v4 deck structure, but reframes for ITBA's multi-OVS reality:
  - 4 distinct business OVSs (owned by business)
  - 4 ITBA DVSs (one per OVS, 1:1 alignment)
  - 1 Enabling Team (shared services)

Slides:
  1.  Cover
  2.  Big Picture — 4 business OVS swim lanes + their supporting DVSs
  3.  Cards — 4 OVS+DVS column pairs + enabling team band
  4.  Why this reframe
  5.  Value Stream Framing (multi-OVS pattern)
  6.  Intake Detail
  7.  OVSs — Multi-Domain View
  8-11. DVS 1-4 detail
  12. Enabling Team
"""
import re
from copy import deepcopy
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

# Use CS Apps v4 as the master/background source (same look-and-feel)
SRC = "Material em andamento/Customer Services - Proposed Value Streams - v4.pptx"
DST = "Material em andamento/IT Business Applications - Proposed Value Streams - v1.pptx"

NAVY  = RGBColor(0x0F, 0x1B, 0x4F)
BLUE2 = RGBColor(0x2C, 0x57, 0x8B)
BLUE3 = RGBColor(0x5B, 0x7F, 0xB5)
BLUE4 = RGBColor(0xA5, 0xC2, 0xDA)
GREEN = RGBColor(0x9C, 0xC4, 0x4A)
TEAL  = RGBColor(0x3A, 0x6F, 0x7E)
ENAB  = RGBColor(0x6C, 0x7A, 0x89)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK  = RGBColor(0x2C, 0x3E, 0x50)
GRAY  = RGBColor(0x95, 0xA5, 0xA6)

# Distinct color per DVS (used in big picture & cards)
DVS_COLORS = {
    "E&G Operations Apps":      RGBColor(0xD3, 0x54, 0x00),
    "Spatial & Asset Solutions": RGBColor(0x16, 0xA0, 0x85),
    "Corporate Applications":   RGBColor(0x8E, 0x44, 0xAD),
    "Global SAP Integration":   RGBColor(0x2C, 0x57, 0x8B),
}

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
    else:
        sh.line.color.rgb = line; sh.line.width = Pt(0.5)
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


def replace_textbox(shape, text, *, size=None, bold=None, italic=None, color=None,
                    name=None, align=None, keep_font=True):
    tf = shape.text_frame
    if keep_font and tf.paragraphs and tf.paragraphs[0].runs:
        f0 = tf.paragraphs[0].runs[0].font
        size = size if size is not None else (f0.size.pt if f0.size else None)
        bold = bold if bold is not None else f0.bold
        italic = italic if italic is not None else f0.italic
        try: color = color or (f0.color.rgb if f0.color and f0.color.rgb else None)
        except Exception: pass
        name = name or f0.name
    tf.clear()
    for i, line in enumerate(text.split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if align is not None: p.alignment = align
        r = p.add_run(); r.text = line
        if name: r.font.name = name
        if size is not None: r.font.size = Pt(size)
        if bold is not None: r.font.bold = bold
        if italic is not None: r.font.italic = italic
        if color is not None: r.font.color.rgb = color


def edit_vs_slide_by_idx(slide, mapping):
    """mapping = dict of {shape_idx: text}. Reuses existing formatting."""
    by_idx = {i: s for i, s in enumerate(slide.shapes)}
    for i, text in mapping.items():
        if i in by_idx:
            replace_textbox(by_idx[i], text)


def delete_slide_safe(prs, idx):
    slide_id = list(prs.slides._sldIdLst)[idx]
    rId = slide_id.rId
    prs.part.drop_rel(rId)
    prs.slides._sldIdLst.remove(slide_id)


# ---------------- main build ----------------

def main():
    prs = Presentation(SRC)

    # Step 1: Drop slides we won't use — Value Lens (slide 13), and slide 7 (OVS detail) and slide 2 (big picture) will be rebuilt
    # Strategy: keep slides 1 (cover), 4 (why), 5 (framing), 6 (intake), 8-11 (DVS 1-4), 12 (enabling team).
    # Rebuild: 2 (big pic), 3 (cards), 7 (OVS multi-view).
    # Drop: 13 (value lens).

    # Indices: 0=cover, 1=bigpic, 2=cards, 3=why, 4=framing, 5=intake, 6=ovs detail, 7-10=DVS, 11=enabling, 12=value lens
    delete_slide_safe(prs, 12)  # drop Value Lens (CS-specific Heather metrics)

    # ============== Cover ==============
    cover = prs.slides[0]
    for sh in cover.shapes:
        if sh.has_text_frame:
            t = sh.text_frame.text
            if "PROPOSED VALUE STREAMS" in t and "CUSTOMER SERVICES" in t:
                replace_textbox(sh,
                    "PROPOSED VALUE STREAMS – IT BUSINESS APPLICATIONS\n"
                    "v1 — 4 Operational + 4 Development Value Streams + 1 Enabling Team\n"
                    "April 2026",
                    keep_font=True)

    # ============== Slide 2: Big Picture (rebuild) ==============
    big = prs.slides[1]
    for sh in list(big.shapes):
        sh._element.getparent().remove(sh._element)
    copy_bg(prs, 1, big)
    build_big_picture(big)

    # ============== Slide 3: Cards (rebuild) ==============
    cards = prs.slides[2]
    for sh in list(cards.shapes):
        sh._element.getparent().remove(sh._element)
    copy_bg(prs, 2, cards)
    build_cards_slide(cards)

    # ============== Slide 4: Why ==============
    why = prs.slides[3]
    edit_why_slide(why)

    # ============== Slide 5: Framing ==============
    framing = prs.slides[4]
    edit_framing_slide(framing)

    # ============== Slide 6: Intake ==============
    intake = prs.slides[5]
    edit_intake_slide(intake)

    # ============== Slide 7: OVSs Multi-View (rebuild) ==============
    ovs_multi = prs.slides[6]
    for sh in list(ovs_multi.shapes):
        sh._element.getparent().remove(sh._element)
    copy_bg(prs, 6, ovs_multi)
    build_ovs_multi_slide(ovs_multi)

    # ============== Slides 8-11: DVS 1-4 ==============
    edit_dvs_slide(prs.slides[7], DVS_CONTENT[0])
    edit_dvs_slide(prs.slides[8], DVS_CONTENT[1])
    edit_dvs_slide(prs.slides[9], DVS_CONTENT[2])
    edit_dvs_slide(prs.slides[10], DVS_CONTENT[3])

    # ============== Slide 12: Enabling Team ==============
    edit_enabling_slide(prs.slides[11])

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


# ============== big picture builder ==============

def build_big_picture(slide):
    add_text(slide, 0.40, 0.10, 12.53, 0.38,
             "OPERATIONAL VALUE STREAMS — IT BUSINESS APPLICATIONS BIG PICTURE",
             size=18, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, 0.40, 0.52, 12.53, 0.23,
             "Four business value chains  ·  four IT Development Value Streams that support them  ·  one enabling team",
             size=10, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # 4 swim lanes
    lanes_top = 1.00
    lane_h = 1.20
    lane_gap = 0.10
    for i, (ovs_name, dvs_name, lead, fte, trigger, steps, outcome, apps) in enumerate(LANES):
        y = lanes_top + i * (lane_h + lane_gap)
        # Left: OVS label (business badge)
        add_rect(slide, 0.40, y, 2.05, lane_h, fill=NAVY,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.10)
        add_text(slide, 0.45, y + 0.08, 1.95, 0.28,
                 "BUSINESS OVS", size=7, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(slide, 0.45, y + 0.32, 1.95, 0.50, ovs_name,
                 size=10.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, 0.45, y + lane_h - 0.30, 1.95, 0.25, f"Lead: {lead}",
                 size=7, italic=True, color=BLUE4, align=PP_ALIGN.CENTER)

        # Middle: chevrons (3-4 steps) + outcome label at right
        chev_x_start = 2.55
        chev_w = 1.30
        chev_h = 0.55
        chev_overlap = 0.10
        chev_y = y + 0.18
        for j, step in enumerate(steps):
            cx = chev_x_start + j * (chev_w - chev_overlap)
            sh = slide.shapes.add_shape(MSO_SHAPE.PENTAGON,
                                        Inches(cx), Inches(chev_y),
                                        Inches(chev_w), Inches(chev_h))
            sh.fill.solid(); sh.fill.fore_color.rgb = GREEN
            sh.line.color.rgb = NAVY; sh.line.width = Pt(0.5)
            tb = slide.shapes.add_textbox(Inches(cx), Inches(chev_y),
                                          Inches(chev_w), Inches(chev_h))
            tf = tb.text_frame; tf.word_wrap = True
            tf.margin_left = Inches(0.04); tf.margin_right = Inches(0.16)
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
            r = p.add_run(); r.text = step
            r.font.name = "Arial"; r.font.size = Pt(7.5); r.font.bold = True; r.font.color.rgb = DARK

        # DVS label below the chevrons
        dvs_y = y + 0.85
        n_steps = len(steps)
        chev_total_w = n_steps * chev_w - (n_steps - 1) * chev_overlap
        add_rect(slide, chev_x_start, dvs_y,
                 chev_total_w, 0.30,
                 fill=DVS_COLORS[dvs_name], shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.30)
        add_text(slide, chev_x_start, dvs_y, chev_total_w, 0.30,
                 f"DVS — {dvs_name}  ·  {fte}  ·  apps: {apps}",
                 size=7.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE)

        # Outcome on right
        out_x = chev_x_start + chev_total_w + 0.20
        add_text(slide, out_x, y + 0.05, 13.20 - out_x, 0.20,
                 "OUTCOME", size=7, bold=True, color=GRAY)
        add_text(slide, out_x, y + 0.25, 13.20 - out_x, lane_h - 0.30,
                 outcome, size=8.5, bold=True, italic=True, color=DARK)

    # Enabling team band at bottom
    et_y = lanes_top + 4 * (lane_h + lane_gap) + 0.10
    add_rect(slide, 0.40, et_y, 12.53, 0.55,
             fill=ENAB, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.20)
    add_text(slide, 0.50, et_y + 0.05, 3.50, 0.20,
             "ENABLING TEAM", size=10, bold=True, color=WHITE)
    add_text(slide, 0.50, et_y + 0.25, 12.30, 0.30,
             "Shared services across the 4 DVS  ·  SWF (Software Factory) governance  ·  "
             "demand intake (PDD · Jira · IT Now)  ·  test & release coordination  ·  "
             "third-party risk gating",
             size=8.5, color=WHITE)

    # Legend
    legend_y = et_y + 0.70
    add_text(slide, 0.40, legend_y, 12.53, 0.40,
             "•  Four business value chains, each owned by a business stakeholder, supported by one ITBA DVS\n"
             "•  Coloured bar under each chevron row = the DVS that owns and evolves the solutions for that OVS",
             size=8.5, color=DARK)


LANES = [
    # (OVS name, DVS name, lead, fte, trigger, steps, outcome, apps thumbnail)
    ("Field & Network Operations",
     "E&G Operations Apps",
     "Bruno Moreto",
     "~7 FTE",
     "",
     ["Detect outage /\nwork need", "Dispatch &\nschedule",
      "Execute work\nin field", "Update systems\n& close"],
     "Network reliable\n· compliance reports filed",
     "ProjectWise · Microstation · 70-90 apps"),
    ("Asset & Spatial Mgmt",
     "Spatial & Asset Solutions",
     "George Porto",
     "~6 FTE",
     "",
     ["Asset change /\nnew build", "Survey &\nmodel",
      "Map & publish\nto network", "Consume in\nplanning"],
     "Assets visible\n& accurate network-wide",
     "SAP HANA · ArcFM · GIS web · SWF"),
    ("Corporate Functions",
     "Corporate Applications",
     "Sarath Garikapati",
     "~8 FTE",
     "",
     ["Regulatory ask /\nemployee need", "Capture\nrequirement",
      "Implement &\ntest", "Deploy &\nreport"],
     "Compliance maintained\n+ employees enabled",
     "Corporate apps · network apps · SWF tooling"),
    ("Iberdrola Global Integration",
     "Global SAP Integration",
     "Vanessa Colombo",
     "~5 FTE",
     "",
     ["Iberdrola corp.\nprocess change", "Align FI/PO/\nBudget modules",
      "Coordinate\ntransports", "Validate\nSpain side"],
     "Iberdrola alignment\n· no integration drift",
     "SAP FI · PO · Budget (Spain side)"),
]


# ============== cards slide builder ==============

def build_cards_slide(slide):
    add_text(slide, 0.40, 0.10, 12.53, 0.38,
             "PROPOSED VALUE STREAMS",
             size=20, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, 0.40, 0.58, 12.53, 0.23,
             "IT BUSINESS APPLICATIONS",
             size=11, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # Header row: "FOUR BUSINESS OVS  ·  ONE ITBA DVS PER OVS"
    add_text(slide, 0.50, 1.05, 12.33, 0.22,
             "FOUR BUSINESS OPERATIONAL VALUE STREAMS  ·  ONE ITBA DEVELOPMENT VS PER OVS  ·  "
             "ONE PRODUCT MANAGER PER DVS",
             size=10, bold=True, color=NAVY)

    # 4 columns (one per OVS+DVS pair)
    n = 4
    margin = 0.50
    gap = 0.12
    total_w = 13.33 - 2 * margin
    col_w = (total_w - (n - 1) * gap) / n  # ≈ 3.05

    # OVS top card height + DVS card height
    ovs_top_y = 1.40
    ovs_top_h = 0.85
    dvs_y = ovs_top_y + ovs_top_h + 0.10
    dvs_h = 4.30

    cards = [
        # (ovs_name, ovs_outcome, dvs_name, lead, fte, domain, desc, apps)
        ("Field & Network Operations",
         "Network reliable  ·  outage response  ·  field crew enabled",
         "E&G Operations Apps",
         "Bruno Moreto + Andrew McGibbon",
         "~7 FTE",
         "DOMAIN · Field & Network Ops",
         "Support 70–90 apps that keep field crews dispatched and assets maintained",
         "ProjectWise\nMicrostation\nDoc viewer\nField mobile / tablet apps\nWork management interfaces\n(70–90 apps total)"),
        ("Asset & Spatial Mgmt",
         "Assets accurate  ·  visible network-wide  ·  ready for planners",
         "Spatial & Asset Solutions",
         "George Porto",
         "~6 FTE",
         "DOMAIN · Spatial / Asset",
         "Capture, model and publish spatial & asset data so engineering and planning can use it",
         "SAP HANA\nArcFM\nGIS web applications\nData replication services\nSWF (Software Factory)"),
        ("Corporate Functions",
         "Compliance kept  ·  employees enabled  ·  network apps stable",
         "Corporate Applications",
         "Sarath Garikapati + Amit",
         "~8 FTE",
         "DOMAIN · Corporate / Back-office",
         "Deliver regulatory and employee-facing technology under continuous compliance pressure",
         "Corporate applications\nNetwork applications\nEmployee-facing portals\nRegulatory reporting tools\nSWF tooling"),
        ("Iberdrola Global Integration",
         "Iberdrola alignment  ·  Spain-side stability  ·  cross-org consistency",
         "Global SAP Integration",
         "Vanessa Colombo + Michael Dubois",
         "~5 FTE",
         "DOMAIN · Iberdrola Spain SAP",
         "Run Spain-side SAP modules and keep US/Iberdrola corporate processes aligned",
         "SAP FI (Finance)\nSAP PO (Purchase Orders)\nSAP Budget control\nIberdrola integration interfaces"),
    ]

    for i, c in enumerate(cards):
        (ovs_name, ovs_outcome, dvs_name, lead, fte, domain, desc, apps) = c
        x = margin + i * (col_w + gap)

        # OVS top card (smaller, navy)
        add_rect(slide, x, ovs_top_y, col_w, ovs_top_h,
                 fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
        add_text(slide, x + 0.05, ovs_top_y + 0.05, col_w - 0.10, 0.20,
                 "BUSINESS OVS", size=7.5, bold=True, color=BLUE4, align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.05, ovs_top_y + 0.22, col_w - 0.10, 0.30, ovs_name,
                 size=10.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.10, ovs_top_y + 0.52, col_w - 0.20, 0.30, ovs_outcome,
                 size=7.5, italic=True, color=BLUE4, align=PP_ALIGN.CENTER)

        # DVS card body
        body = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                      Inches(x), Inches(dvs_y + 0.08),
                                      Inches(col_w), Inches(dvs_h - 0.08))
        body.fill.solid(); body.fill.fore_color.rgb = WHITE
        body.line.color.rgb = NAVY; body.line.width = Pt(0.5)
        body.adjustments[0] = 0.05

        hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                     Inches(x - 0.02), Inches(dvs_y),
                                     Inches(col_w + 0.04), Inches(0.78))
        hdr.fill.solid(); hdr.fill.fore_color.rgb = DVS_COLORS[dvs_name]
        hdr.line.fill.background()
        hdr.adjustments[0] = 0.10

        add_text(slide, x, dvs_y + 0.08, col_w, 0.32, "ITBA DVS",
                 size=7.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(slide, x, dvs_y + 0.32, col_w, 0.40, dvs_name,
                 size=10.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE)

        inner_y = dvs_y + 0.92
        add_text(slide, x + 0.10, inner_y, col_w - 0.20, 0.20,
                 f"{lead}  ·  {fte}",
                 size=8, bold=True, italic=True, color=DARK, align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.10, inner_y + 0.22, col_w - 0.20, 0.20,
                 domain, size=7.5, bold=True, italic=True, color=DARK, align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.10, inner_y + 0.42, col_w - 0.20, 0.55,
                 desc, size=9, italic=True, color=DARK, align=PP_ALIGN.CENTER)

        ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                    Inches(x + 0.20), Inches(inner_y + 1.05),
                                    Inches(col_w - 0.40), Inches(0.02))
        ln.fill.solid(); ln.fill.fore_color.rgb = NAVY
        ln.line.fill.background()

        add_text(slide, x + 0.10, inner_y + 1.13, col_w - 0.20, 0.20,
                 "ALL APPLICATIONS", size=8, bold=True, color=DARK, align=PP_ALIGN.CENTER)

        apps_y = inner_y + 1.35
        add_text(slide, x + 0.10, apps_y, col_w - 0.20, dvs_y + dvs_h - apps_y - 0.10,
                 apps, size=8.5, color=DARK, align=PP_ALIGN.CENTER)

    # Enabling team band
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
             "test & release coordination  ·  third-party risk gating  ·  vendor-mgmt cadence "
             "(Accenture · SWF · business vendors)",
             size=9, color=WHITE)


# ============== why slide ==============

def edit_why_slide(slide):
    edit_vs_slide_by_idx(slide, {
        0: "WHY THIS REFRAME",
        1: "26 FTE  ·  4 TEAMS  ·  4 DISTINCT BUSINESS DOMAINS — NAMING THEM EXPLICITLY CLARIFIES OWNERSHIP",
        3: "ITBA · Carlos Garbin · 4 sub-teams (G&E · GIS · Corporate · Global SAP)",
        5: "Source: AS-IS deliverable",
        8: ("ITBA serves four distinct business value chains — each with its own customer, demand pattern "
            "and stakeholders. Treating them as one undifferentiated pool hides ownership and makes "
            "prioritisation arbitrary."),
        11: "THE CONSTRAINT — WHAT WE HAVE",
        12: ("•  ~26 FTE in ITBA across 4 sub-teams\n"
             "•  G&E: 70–90 apps, ~7 FTE (Bruno Moreto + Andrew McGibbon)\n"
             "•  GIS & OSG: SAP HANA + ArcFM, ~6 FTE (George Porto)\n"
             "•  Corporate: 5–6 PMs + 2 tech leads, ~8 FTE (Sarath Garikapati + Amit)\n"
             "•  Global SAP: Spain-side FI/PO/Budget, ~5 FTE (Vanessa Colombo + Michael Dubois)"),
        15: "THE PRINCIPLE — FROM AS-IS TO-BE",
        16: ("•  AS-IS: organised by sub-team, each running its own intake & prioritisation\n"
             "•  Break/fix is the dominant pattern (especially GIS · G&E · Corporate)\n"
             "•  Demand split: large = PDD (yearly), small/medium = ad-hoc, no formal process\n"
             "•  TO-BE: name each business value chain · pair it with one ITBA DVS · share enabling services"),
        17: "Source: AS-IS deliverable, ITBA findings",
        20: "THE LEVER",
        21: ("•  Operational Value Stream (OVS) = the business flow (owned by business stakeholder)\n"
             "•  Development Value Stream (DVS) = the IT pipeline that builds & evolves the solutions\n"
             "•  ITBA already aligns 1:1 to 4 OVSs — this reframe makes it explicit\n"
             "•  Each DVS gets one PM accountable for its OVS's outcomes\n"
             "•  Shared services (SWF · intake · test) consolidated into a single enabling team"),
        22: "Reframe = same headcount · clearer ownership · explicit business alignment",
        25: "WHAT CHANGES vs. AS-IS",
        26: ("•  AS-IS: 4 sub-teams loosely coordinated, demand fragmented across PDD/Jira/IT Now\n"
             "•  TO-BE: 4 OVSs explicitly named (business owners) + 4 DVSs (1:1 alignment)\n"
             "•  Single intake per DVS · SWF governance centralised in enabling team\n"
             "•  PMs at DVS level · Director ITBA orchestrates cross-DVS\n"
             "•  Shared third-party risk gating · no per-team workaround"),
        27: "Net: same FTE base · 4 explicit OVS-DVS pairs · enabling team for shared services",
    })


# ============== framing slide ==============

def edit_framing_slide(slide):
    edit_vs_slide_by_idx(slide, {
        0: "VALUE STREAM FRAMING",
        1: "MULTIPLE OPERATIONAL VS — ONE DVS PER OVS — SHARED ENABLING TEAM",
        3: "Director ITBA · Carlos Garbin  ·  4 DVS Product Managers (NEW)",
        5: "Multi-OVS pattern · 1:1 DVS alignment",
        8: ("ITBA does not have a single customer-facing flow like Meter-to-Cash. Instead it serves "
            "four parallel business value chains, each with its own outcome. The reframe pairs each OVS "
            "with one DVS so accountability is explicit."),
        11: "BUSINESS OPERATIONAL VALUE STREAMS (4)",
        12: ("•  Field & Network Operations\n"
             "    – Outcome: network reliable · outage response · field crews enabled\n"
             "•  Asset & Spatial Mgmt\n"
             "    – Outcome: assets accurate & visible network-wide\n"
             "•  Corporate Functions\n"
             "    – Outcome: compliance + employee enablement\n"
             "•  Iberdrola Global Integration\n"
             "    – Outcome: Iberdrola alignment · Spain-side stability"),
        15: "ITBA DEVELOPMENT VALUE STREAMS (4)",
        16: ("•  DVS 1 — E&G Operations Apps  ·  Bruno Moreto + Andrew McGibbon  ·  ~7 FTE\n"
             "•  DVS 2 — Spatial & Asset Solutions  ·  George Porto  ·  ~6 FTE\n"
             "•  DVS 3 — Corporate Applications  ·  Sarath Garikapati + Amit  ·  ~8 FTE\n"
             "•  DVS 4 — Global SAP Integration  ·  Vanessa Colombo + Michael Dubois  ·  ~5 FTE\n"
             "+ Enabling Team — Foundation services (SWF · intake · test · risk)"),
        17: "FTE = current org snapshot · sizes to validate post-design",
        20: "MAPPING — OVS → DVS → APPS",
        21: ("•  Field & Network Ops    → E&G Operations Apps    · ProjectWise · Microstation · 70–90 apps\n"
             "•  Asset & Spatial Mgmt   → Spatial & Asset Sols   · SAP HANA · ArcFM · GIS web · data repl.\n"
             "•  Corporate Functions    → Corporate Applications · corporate · network · SWF tooling\n"
             "•  Iberdrola Integration  → Global SAP Integration · SAP FI · PO · Budget (Spain side)\n"
             "•  Enabling Team supports all 4 with: SWF · intake · test/release · third-party risk"),
        22: "1 OVS : 1 DVS — clean accountability · enabling team is shared support",
        25: "WHY THIS WORKS WITHIN HEADCOUNT",
        26: ("•  4 dedicated Product Managers (one per DVS) — already implicit in 4 sub-team leads\n"
             "•  Foundation services run as small enabling team — no extra PM\n"
             "•  Each DVS keeps current FTE base (5–8) — feasible squad size\n"
             "•  Director ITBA (Carlos) = existing role · no new layer\n"
             "•  Vendor extension model preserved per DVS (Accenture, SWF, business vendors)"),
        27: "Validates against AS-IS constraint: 4 sub-teams · ~26 FTE · multi-domain scope",
    })


# ============== intake slide ==============

def edit_intake_slide(slide):
    # Slide 6 (intake) has 20 shapes — different layout. Replace text fields by approximate match.
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        t = sh.text_frame.text
        if "INTAKE DETAIL" == t.strip() or t.strip().startswith("INTAKE DETAIL"):
            replace_textbox(sh, "INTAKE DETAIL")
        elif "CUSTOMER SERVICES APPS" in t and "INTAKE" not in t.upper():
            replace_textbox(sh, "IT BUSINESS APPLICATIONS")
        elif "CUSTOMER SERVICES APPS · Kumar Jaladanki · 1 OVS" in t:
            replace_textbox(sh,
                "ITBA · Carlos Garbin · 4 OVSs (Field & Network · Asset & Spatial · Corporate · Iberdrola) "
                "+ 1 enabling team")
        elif "DEMAND SOURCES — who sends the work" in t:
            pass  # keep label
        elif "OPCO call centers" in t:
            replace_textbox(sh,
                "•  G&E business — outage response · network projects · field-app requests\n"
                "•  Engineering & Planning — GIS asset changes · spatial data needs\n"
                "•  HR / Finance / Legal / Compliance — corporate apps · regulatory mandates\n"
                "•  Iberdrola Spain corporate — FI / PO / Budget alignment & process changes\n"
                "•  Cross-cutting: vendor-driven upgrades (ProjectWise · Microstation · SAP)")
        elif "CHANNELS BY TYPE" in t:
            pass  # keep label
        elif "ONE Jira request type per DVS" in t:
            replace_textbox(sh,
                "•  ONE Jira request type per DVS (Field & Network · Asset · Corporate · Iberdrola)\n"
                "•  PDD — large yearly demand (capex / strategic projects)\n"
                "•  IT Now — incidents only (break/fix is the dominant ITBA pattern)\n"
                "•  Email · Teams · Word — DECOMMISSIONED for new demand intake\n"
                "•  Mandatory fields: DVS · type · size · CoS · business area · sponsor")
        elif "CLASSIFICATION" in t and "branch-specific" in t:
            pass  # keep label
        elif "Run / Grow / Transform" in t and "labeled at intake" in t:
            replace_textbox(sh,
                "•  Break/Fix · Grow · Transform — labeled at intake (ITBA today is heavily Break/Fix)\n"
                "•  Class of Service — Standard · Expedite (regulatory) · Fixed-date · Intangible\n"
                "•  Regulatory items always Fixed-date — third-party risk gates required\n"
                "•  Version upgrades (ProjectWise · Microstation) pre-classified as Fixed-date\n"
                "•  Iberdrola coordination items routed via Global SAP DVS only")
        elif "PRIORITIZATION" in t and "Pain Score" in t:
            pass  # keep label
        elif "Pain Score H(x)" in t:
            replace_textbox(sh,
                "•  Pain Score H(x) = (BV·w₁ + Risk·w₂ + Compliance·w₃) / Effort  (proposed)\n"
                "•  ITBA weighting — Compliance heavy: w₁=0.30 · w₂=0.30 · w₃=0.40\n"
                "•  Backlog ranked monthly at Gate 2 by the DVS Product Manager\n"
                "•  Cross-DVS conflicts (e.g. SWF capacity contention) resolved at the ITBA Forum\n"
                "•  Capacity reserved 15–20% per quarter for Intangible (tech debt · upgrades)")


# ============== OVS multi-view slide ==============

def build_ovs_multi_slide(slide):
    add_text(slide, 0.40, 0.10, 12.53, 0.38,
             "OPERATIONAL VALUE STREAMS — MULTI-DOMAIN VIEW",
             size=20, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, 0.40, 0.58, 12.53, 0.23,
             "ITBA SUPPORTS FOUR DISTINCT BUSINESS VALUE CHAINS  ·  EACH WITH ITS OWN CUSTOMER, OUTCOME AND DVS",
             size=11, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # 4 cards in a 2x2 grid
    margin = 0.50
    gap = 0.25
    total_w = 13.33 - 2 * margin
    col_w = (total_w - gap) / 2
    row_h = 2.85
    row_gap = 0.20
    top_y = 1.10

    OVS_DETAILS = [
        ("Field & Network Operations",
         "Business owner: G&E Operations leadership",
         "Customer: Field crews · dispatchers · network engineers",
         "Trigger: Outage detected · maintenance scheduled · network upgrade",
         "Steps:  Detect / plan  →  Dispatch  →  Execute work  →  Update systems & close",
         "Outcome: Network reliable · outage time minimised · regulatory reports filed",
         "Supporting DVS: E&G Operations Apps  ·  Bruno Moreto + Andrew McGibbon  ·  ~7 FTE",
         DVS_COLORS["E&G Operations Apps"]),
        ("Asset & Spatial Mgmt",
         "Business owner: Engineering / Asset Management",
         "Customer: Engineering · planning · regulatory reporting",
         "Trigger: Asset created / modified · network re-survey · planning need",
         "Steps:  Survey  →  Model in GIS  →  Publish to network  →  Consume in planning",
         "Outcome: Assets accurate & visible network-wide · planning data trusted",
         "Supporting DVS: Spatial & Asset Solutions  ·  George Porto  ·  ~6 FTE",
         DVS_COLORS["Spatial & Asset Solutions"]),
        ("Corporate Functions",
         "Business owner: HR · Finance · Legal · Compliance leadership",
         "Customer: All Avangrid US employees · regulators · auditors",
         "Trigger: Regulatory ask · employee enablement need · upgrade cycle",
         "Steps:  Capture requirement  →  Implement & test  →  Deploy  →  Comply / report",
         "Outcome: Continuous compliance · employees enabled · network apps stable",
         "Supporting DVS: Corporate Applications  ·  Sarath Garikapati + Amit  ·  ~8 FTE",
         DVS_COLORS["Corporate Applications"]),
        ("Iberdrola Global Integration",
         "Business owner: Iberdrola corporate (Spain) · Avangrid Finance",
         "Customer: Iberdrola corporate stakeholders · Avangrid Finance integration",
         "Trigger: Iberdrola process change · cross-org reporting · transport coordination",
         "Steps:  Receive change  →  Align FI / PO / Budget  →  Coordinate transports  →  Validate",
         "Outcome: Iberdrola alignment · Spain-side stability · no integration drift",
         "Supporting DVS: Global SAP Integration  ·  Vanessa Colombo + Michael Dubois  ·  ~5 FTE",
         DVS_COLORS["Global SAP Integration"]),
    ]

    for i, c in enumerate(OVS_DETAILS):
        col = i % 2
        row = i // 2
        x = margin + col * (col_w + gap)
        y = top_y + row * (row_h + row_gap)
        ovs_name, owner, customer, trigger, steps, outcome, dvs, color = c

        # Body
        body = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                      Inches(x), Inches(y + 0.05),
                                      Inches(col_w), Inches(row_h - 0.05))
        body.fill.solid(); body.fill.fore_color.rgb = WHITE
        body.line.color.rgb = NAVY; body.line.width = Pt(0.5)
        body.adjustments[0] = 0.04

        # Header
        hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                     Inches(x - 0.02), Inches(y),
                                     Inches(col_w + 0.04), Inches(0.55))
        hdr.fill.solid(); hdr.fill.fore_color.rgb = NAVY
        hdr.line.fill.background(); hdr.adjustments[0] = 0.06

        add_text(slide, x + 0.20, y + 0.08, col_w - 3.00, 0.40,
                 ovs_name, size=12, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
        add_text(slide, x + col_w - 3.00, y + 0.10, 2.80, 0.36,
                 "BUSINESS OVS", size=8, bold=True, color=BLUE4, align=PP_ALIGN.RIGHT)

        # Body content
        cy = y + 0.65
        rows = [
            ("BUSINESS OWNER", owner),
            ("CUSTOMER", customer),
            ("TRIGGER", trigger),
            ("VALUE-ADD STEPS", steps),
            ("OUTCOME", outcome),
        ]
        for label, val in rows:
            add_text(slide, x + 0.15, cy, 1.50, 0.20, label,
                     size=7, bold=True, color=GRAY)
            add_text(slide, x + 1.70, cy, col_w - 1.85, 0.40, val,
                     size=8.5, color=DARK)
            cy += 0.32

        # Bottom DVS strip
        sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                    Inches(x + 0.15), Inches(y + row_h - 0.40),
                                    Inches(col_w - 0.30), Inches(0.32))
        sh.fill.solid(); sh.fill.fore_color.rgb = color
        sh.line.fill.background(); sh.adjustments[0] = 0.30
        add_text(slide, x + 0.15, y + row_h - 0.40, col_w - 0.30, 0.32,
                 dvs, size=8.5, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ============== DVS detail slides ==============

DVS_CONTENT = [
    # DVS 1 — E&G Operations Apps
    {
        0: "DEVELOPMENT VALUE STREAM (DVS) 1 / 4",
        1: "E&G OPERATIONS APPS — KEEP FIELD CREWS DISPATCHED · ASSETS MAINTAINED · NETWORK STABLE",
        3: "ITBA · Bruno Moreto (Sr. Manager Electric & Gas Ops) + Andrew McGibbon",
        5: "Run · Grow · Transform mix · to validate with team (heavy break/fix today)",
        8: ("Operate and evolve the 70–90 apps that support Electric & Gas operations — "
            "from outage response to scheduled maintenance — across Avangrid's service territory."),
        11: "APPLICATIONS SUPPORTED",
        12: ("•  ProjectWise (engineering content / drawing mgmt)\n"
             "•  Microstation (CAD / engineering design)\n"
             "•  Doc viewer (field document access)\n"
             "•  Field mobile / tablet apps (Android / tablet)\n"
             "•  Work management interfaces\n"
             "•  Plus 70–90 supporting apps in scope"),
        15: "DEMAND SOURCES — who sends the work",
        16: ("•  G&E business leadership — outage / maintenance / capex priorities\n"
             "•  Field operations crews — app requests via business intake\n"
             "•  Vendor-driven version upgrades (ProjectWise · Microstation)\n"
             "•  Third-party risk assessments — gating for vendor changes"),
        17: "Routed via the Jira request type for this DVS",
        20: "AVANGRID CORE TEAM",
        21: ("•  Product Manager (NEW · same level as other DVS PMs)\n"
             "•  Tech Lead — Bruno Moreto / Andrew McGibbon (existing)\n"
             "•  Business Analyst — field ops SME\n"
             "•  Squad size · to validate (currently ~7 FTE total)"),
        22: "Heaviest break/fix load — capacity reservation for incidents needs explicit guard-rail",
        25: "VENDOR EXTENSION (delivery)",
        26: ("•  SWF (Software Factory) — execution muscle\n"
             "•  Business vendors — field apps · doc viewer\n"
             "•  ProjectWise vendor (Bentley) — upgrades & support\n"
             "•  Microstation vendor (Bentley) — upgrades & support"),
        27: "Avangrid = product owner · vendor = delivery muscle",
    },
    # DVS 2 — Spatial & Asset Solutions
    {
        0: "DEVELOPMENT VALUE STREAM (DVS) 2 / 4",
        1: "SPATIAL & ASSET SOLUTIONS — ASSETS ACCURATE · VISIBLE NETWORK-WIDE · READY FOR PLANNERS",
        3: "ITBA · George Porto (Sr. Manager GIS & OSG)",
        5: "Run · Grow · Transform mix · to validate with team",
        8: ("Capture, model and publish spatial & asset data so engineering, planning and "
            "regulatory reporting can use trusted information."),
        11: "APPLICATIONS SUPPORTED",
        12: ("•  SAP HANA (data platform underlying GIS workloads)\n"
             "•  ArcFM (GIS / asset management)\n"
             "•  GIS web applications (consumer-facing spatial tools)\n"
             "•  Data replication services\n"
             "•  SWF (Software Factory) integration"),
        15: "DEMAND SOURCES — who sends the work",
        16: ("•  Engineering — new builds · network changes\n"
             "•  Planning — spatial data needs for forecasting\n"
             "•  Regulatory reporting — asset disclosure requirements\n"
             "•  Field operations — asset updates from the field\n"
             "•  GIS server / platform upgrade cycle"),
        17: "Routed via the Jira request type for this DVS",
        20: "AVANGRID CORE TEAM",
        21: ("•  Product Manager (NEW)\n"
             "•  Tech Lead — George Porto (existing)\n"
             "•  Business Analyst — GIS / spatial SME\n"
             "•  Squad size · to validate (currently ~6 FTE total)"),
        22: "Adoption of new tech is a stated team challenge — PM should own adoption metrics",
        25: "VENDOR EXTENSION (delivery)",
        26: ("•  SWF (Software Factory) — execution muscle\n"
             "•  ArcFM / Esri partner — GIS platform support\n"
             "•  SAP HANA partner — data platform\n"
             "•  GIS web app vendors"),
        27: "Avangrid = product owner · vendor = delivery muscle",
    },
    # DVS 3 — Corporate Applications
    {
        0: "DEVELOPMENT VALUE STREAM (DVS) 3 / 4",
        1: "CORPORATE APPLICATIONS — REGULATORY COMPLIANCE · EMPLOYEE ENABLEMENT · NETWORK APPS STABLE",
        3: "ITBA · Sarath Garikapati (Sr. Manager Corporate Apps & Projects) + Amit",
        5: "Run · Grow · Transform mix · to validate with team",
        8: ("Deliver and maintain regulatory and employee-facing technology across Avangrid US, "
            "operating under continuous compliance pressure."),
        11: "APPLICATIONS SUPPORTED",
        12: ("•  Corporate applications portfolio\n"
             "•  Network applications\n"
             "•  Employee-facing portals & tools\n"
             "•  Regulatory reporting tools\n"
             "•  SWF (Software Factory) tooling"),
        15: "DEMAND SOURCES — who sends the work",
        16: ("•  HR / Finance / Legal / Compliance — corporate process changes\n"
             "•  Regulators — mandated reporting cadence\n"
             "•  Employee population — enablement & training\n"
             "•  Continuous new requirements (per Sarath: \"always a priority\")"),
        17: "Routed via the Jira request type for this DVS",
        20: "AVANGRID CORE TEAM",
        21: ("•  Product Manager (NEW)\n"
             "•  Tech Lead — Sarath Garikapati / Amit (existing)\n"
             "•  5–6 PMs + 2 tech leads (current composition)\n"
             "•  Squad size · to validate (currently ~8 FTE total)"),
        22: "Largest sub-team and most PM-heavy — possible internal sub-streams to validate",
        25: "VENDOR EXTENSION (delivery)",
        26: ("•  Accenture (corporate apps + projects)\n"
             "•  SWF (Software Factory) — execution muscle\n"
             "•  Regulatory tooling vendors\n"
             "•  Employee portal vendors"),
        27: "Avangrid = product owner · vendor = delivery muscle",
    },
    # DVS 4 — Global SAP Integration
    {
        0: "DEVELOPMENT VALUE STREAM (DVS) 4 / 4",
        1: "GLOBAL SAP INTEGRATION — SPAIN-SIDE STABILITY · IBERDROLA ALIGNMENT · NO INTEGRATION DRIFT",
        3: "ITBA · Vanessa Colombo (Manager IT Apps Global SAP) + Michael Dubois",
        5: "Run · Grow · Transform mix · to validate with team",
        8: ("Run Spain-side SAP modules and keep US/Iberdrola corporate processes aligned — "
            "balancing incidents, small demands, and major projects under hiring constraints."),
        11: "APPLICATIONS SUPPORTED",
        12: ("•  SAP FI (Finance) — Spain side\n"
             "•  SAP PO (Purchase Orders) — Spain side\n"
             "•  SAP Budget control\n"
             "•  Iberdrola integration interfaces (US ↔ Spain)\n"
             "•  Cross-org transport coordination"),
        15: "DEMAND SOURCES — who sends the work",
        16: ("•  Iberdrola corporate (Spain) — process changes\n"
             "•  Avangrid Finance — US-side integration needs\n"
             "•  Incidents (IT Now) — Spain SAP module issues\n"
             "•  Small Demands — module enhancements\n"
             "•  Major projects (PDD) — Spain SAP roadmap"),
        17: "Routed via the Jira request type for this DVS",
        20: "AVANGRID CORE TEAM",
        21: ("•  Product Manager (NEW)\n"
             "•  Tech Lead — Vanessa Colombo / Michael Dubois (existing)\n"
             "•  Business Analyst — SAP FI/PO SME\n"
             "•  Squad size · to validate (currently ~5 FTE total)"),
        22: "Hiring lag (~1 year) + knowledge-transfer loss flagged as top risks by team",
        25: "VENDOR EXTENSION (delivery)",
        26: ("•  Accenture (Global SAP support)\n"
             "•  SWF (Software Factory) — execution muscle\n"
             "•  SAP partner for Spain modules\n"
             "•  Iberdrola Spain integration teams (peer)"),
        27: "Avangrid = product owner · vendor = delivery muscle",
    },
]


def edit_dvs_slide(slide, content):
    edit_vs_slide_by_idx(slide, content)


# ============== enabling team slide ==============

def edit_enabling_slide(slide):
    edit_vs_slide_by_idx(slide, {
        0:  "ENABLING TEAM",
        1:  "FOUNDATION SERVICES — SHARED CAPABILITIES THAT KEEP THE 4 ITBA DVS RUNNING",
        3:  "ITBA · governed by Director ITBA (Carlos Garbin) + cross-DVS coordination",
        5:  "No PM · no DVS-level backlog · operates as a small shared-service team",
        8:  ("Provide the technical and process foundation — Software Factory governance, "
             "demand intake, test/release coordination, third-party risk gating — that the 4 ITBA DVS "
             "rely on. Not a product."),
        11: "FOUNDATION SERVICES",
        12: ("•  SWF (Software Factory) governance — used across all 4 DVS\n"
             "•  Demand intake orchestration (PDD · Jira · IT Now)\n"
             "•  Test & release coordination (cross-DVS)\n"
             "•  Third-party risk assessment gating (ProjectWise · Microstation · others)\n"
             "•  Vendor-management cadence (Accenture · SWF · business vendors)\n"
             "•  ALM tooling administration"),
        15: "WHO USES THESE SERVICES",
        16: ("•  All 4 ITBA DVS — for SWF capacity · environments · test cycles\n"
             "•  G&E DVS especially — third-party risk gating for vendor upgrades\n"
             "•  Global SAP DVS — transport coordination · Iberdrola integration\n"
             "•  Director ITBA — portfolio view across DVS\n"
             "•  Optionally shared with CS Apps enabling team (Digital Products consolidation)"),
        17: "Foundation services support all 4 ITBA DVS — they do not constitute a separate product line",
        20: "TEAM SHAPE",
        21: ("•  Coordination led by Director ITBA (Carlos Garbin)\n"
             "•  SWF governance lead (existing role within Corporate / cross-DVS)\n"
             "•  Demand intake / portfolio admin\n"
             "•  Test & risk coordinators — embedded in DVS where feasible\n"
             "•  No dedicated Product Manager\n"
             "Squad size · to validate"),
        22: "Why not a DVS: at 26 FTE total, productizing foundation services would create overhead without internal customer roadmap",
        25: "VENDOR EXTENSION",
        26: ("•  SWF (Software Factory) vendor — capacity & tooling\n"
             "•  Accenture — for cross-DVS major projects\n"
             "•  Atlassian Jira admins\n"
             "•  Third-party risk assessment vendors"),
        27: "Avangrid governs the team operationally · vendors execute specific capabilities",
    })


if __name__ == "__main__":
    main()
