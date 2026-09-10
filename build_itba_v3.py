"""
build_itba_v3.py — rebuild Part B with SAFe-faithful approach (5 DVSs grouped by solution family).

Final structure (20 slides):
  Part A — Initial proposal (kept from v2)
    1.  Cover
    2.  Combined view (4 blue VS cards)
    3-6. 4 blue VS detail

  Part B — SAFe-faithful exercise
    7.  Big Picture (5 DVSs by solution family + multiple OVSs per domain)
    8.  Cards (5 DVS cards · boundary flags inline)
    9.  Framing — why solution-family approach
    10. Many-to-many matrix (OVS × DVS)
    11-14. 4 OVS detail (Network Build · Outage · Rate Case · P2P) — show multi-DVS support
    15-19. 5 DVS detail (Engineering & Spatial · Grid Ops · Field Workforce · Enterprise Content & Reg · SAP Platform)
    20. Enabling Team
"""
from copy import deepcopy
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

SRC = "Material em andamento/IT Business Applications - Proposed Value Streams + Value Definition.pptx"
DST = "Material em andamento/IT Business Applications - Proposed Value Streams + Value Definition.pptx"

# Brand
NAVY  = RGBColor(0x0F, 0x1B, 0x4F)
BLUE2 = RGBColor(0x2C, 0x57, 0x8B)
BLUE3 = RGBColor(0x5B, 0x7F, 0xB5)
BLUE4 = RGBColor(0xA5, 0xC2, 0xDA)
GREEN = RGBColor(0x9C, 0xC4, 0x4A)
TEAL  = RGBColor(0x0B, 0x76, 0xA0)
RED   = RGBColor(0x98, 0x00, 0x03)
ENAB  = RGBColor(0x6C, 0x7A, 0x89)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK  = RGBColor(0x2C, 0x3E, 0x50)
GRAY  = RGBColor(0x95, 0xA5, 0xA6)
LIGHTBG = RGBColor(0xEE, 0xEE, 0xEE)
WARN  = RGBColor(0xE6, 0x7E, 0x22)  # boundary flag color

# 5 DVS colors (solution family)
DVS5_COLORS = {
    "Engineering & Spatial":     RGBColor(0x16, 0xA0, 0x85),  # green
    "Grid Operations (OT)":      RGBColor(0xC0, 0x39, 0x2B),  # red — OT
    "Field Workforce":           RGBColor(0xF3, 0x9C, 0x12),  # amber
    "Enterprise Content & Reg":  RGBColor(0x8E, 0x44, 0xAD),  # purple
    "SAP Enterprise Platform":   RGBColor(0x2C, 0x57, 0x8B),  # navy-blue
}

# Domain colors (for OVS pills)
DOMAIN_COLORS = {
    "Engineering & Asset":  RGBColor(0x16, 0xA0, 0x85),
    "Network Operations":   RGBColor(0xD3, 0x54, 0x00),
    "Corporate & Regulatory": RGBColor(0x8E, 0x44, 0xAD),
    "Iberdrola":            RGBColor(0x2C, 0x57, 0x8B),
}

# OVSs per domain
OVS_BY_DOMAIN = {
    "Engineering & Asset": [
        "Network Build", "Asset Replacement", "Capital Planning", "Asset Decommissioning",
    ],
    "Network Operations": [
        "Outage Response", "Grid Optimization", "Routine Switching", "Storm Restoration",
    ],
    "Corporate & Regulatory": [
        "Source-to-Pay", "Hire-to-Retire", "Rate Case Filing", "Regulatory Reporting",
    ],
    "Iberdrola": [
        "Procure-to-Pay (Spain)", "Period Close (Spain)", "Plan-to-Budget (Spain)",
    ],
}

# DVSs by solution family
DVS_LIST = [
    {
        "name": "Engineering & Spatial",
        "tag": "Asset / CAD / GIS",
        "rationale": "Vendor-driven · quarterly upgrades · CAD/GIS skillset · planner customer",
        "apps": "ArcFM · ProjectWise · Microstation · Bentley · CYME · SAP HANA spatial · GIS web apps",
        "boundary": None,
        "lead": "George Porto + crossover (Bruno's team for ProjectWise/Bentley)",
        "color": DVS5_COLORS["Engineering & Spatial"],
    },
    {
        "name": "Grid Operations (OT)",
        "tag": "SCADA / Real-time",
        "rationale": "Real-time · safety-critical · 24/7 · OT skillset (DevSecOps + SRE)",
        "apps": "Echelon · Cimplicity · SLAM · Switching Orders",
        "boundary": "⚠ Traditionally OT, not IT — needs executive decision on org placement",
        "lead": "Bruno Moreto + OT specialists (NEW capability if kept in IT)",
        "color": DVS5_COLORS["Grid Operations (OT)"],
    },
    {
        "name": "Field Workforce",
        "tag": "Mobile / Scheduling",
        "rationale": "Mobile · scheduling · workforce mgmt · field crew customer",
        "apps": "Lemur-SF · FWM · ARCOS · Field Force",
        "boundary": "⚠ Currently shown under CS Apps Field Operations VS — overlap to resolve",
        "lead": "TBD — overlaps Bruno's team and CS Apps field-ops",
        "color": DVS5_COLORS["Field Workforce"],
    },
    {
        "name": "Enterprise Content & Reg",
        "tag": "Workflow / Compliance",
        "rationale": "Transactional · workflow · compliance-driven · corporate user",
        "apps": "OpenText · Rate Case Advisor · Utilisphere · corporate apps · employee portals · regulatory tools",
        "boundary": None,
        "lead": "Sarath Garikapati + Amit",
        "color": DVS5_COLORS["Enterprise Content & Reg"],
    },
    {
        "name": "SAP Enterprise Platform",
        "tag": "SAP modules",
        "rationale": "Same SAP stack · same SAP skillset · same transport cadence",
        "apps": "SAP FI · PO · Budget · S/4HANA (Spain)  +  SAP ECC · CRM · FICA (US — currently CS Apps)",
        "boundary": "⚠ Cross-Avangrid: should US (CS Apps) and Spain (ITBA) SAP be one platform DVS?",
        "lead": "Vanessa Colombo + Michael Dubois (Spain) + Vincent Lam / Naggi Reddy (US — CS Apps)",
        "color": DVS5_COLORS["SAP Enterprise Platform"],
    },
]

# Many-to-many mapping: which DVSs each OVS uses
OVS_DVS_MAP = {
    "Network Build":              ["Engineering & Spatial", "Field Workforce", "Enterprise Content & Reg"],
    "Asset Replacement":          ["Engineering & Spatial", "Field Workforce"],
    "Capital Planning":           ["Engineering & Spatial", "Enterprise Content & Reg"],
    "Asset Decommissioning":      ["Engineering & Spatial", "Field Workforce"],
    "Outage Response":            ["Grid Operations (OT)", "Field Workforce"],
    "Grid Optimization":          ["Grid Operations (OT)", "Engineering & Spatial"],
    "Routine Switching":          ["Grid Operations (OT)"],
    "Storm Restoration":          ["Grid Operations (OT)", "Field Workforce", "Enterprise Content & Reg"],
    "Source-to-Pay":              ["Enterprise Content & Reg", "SAP Enterprise Platform"],
    "Hire-to-Retire":             ["Enterprise Content & Reg"],
    "Rate Case Filing":           ["Enterprise Content & Reg"],
    "Regulatory Reporting":       ["Enterprise Content & Reg", "SAP Enterprise Platform"],
    "Procure-to-Pay (Spain)":     ["SAP Enterprise Platform"],
    "Period Close (Spain)":       ["SAP Enterprise Platform"],
    "Plan-to-Budget (Spain)":     ["SAP Enterprise Platform"],
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


def replace_textbox(shape, text):
    tf = shape.text_frame
    if tf.paragraphs and tf.paragraphs[0].runs:
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


def clear_slide(slide):
    for sh in list(slide.shapes):
        sh._element.getparent().remove(sh._element)


# ---------------- BIG PICTURE v3: 5 DVSs supporting many OVSs ----------------

def build_big_picture_v3(slide):
    add_text(slide, 0.38, 0.17, 12.53, 0.38,
             "ITBA — VALUE STREAM BIG PICTURE (SAFe-FAITHFUL)",
             size=18, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, 0.40, 0.58, 12.53, 0.23,
             "5 ITBA Development Value Streams (grouped by SOLUTION FAMILY)  ·  many business OVSs (many-to-many mapping)",
             size=10.5, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # Top: 4 OVS domain groups (color-coded), each with N OVS pills inside
    top_label_y = 1.05
    add_text(slide, 0.40, top_label_y, 12.53, 0.22,
             "BUSINESS OPERATIONAL VALUE STREAMS (owned by business · grouped by domain)",
             size=10, bold=True, color=NAVY)

    domain_y = 1.30
    domain_h = 1.45
    n_domains = 4
    margin = 0.40
    gap = 0.12
    total_w = 13.33 - 2 * margin
    domain_w = (total_w - (n_domains - 1) * gap) / n_domains

    domains = list(OVS_BY_DOMAIN.keys())
    for i, domain in enumerate(domains):
        x = margin + i * (domain_w + gap)
        color = DOMAIN_COLORS[domain]

        add_rect(slide, x, domain_y, domain_w, domain_h, fill=LIGHTBG,
                 line=color, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
        # Domain label header
        add_rect(slide, x, domain_y, domain_w, 0.30, fill=color,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.10)
        add_text(slide, x, domain_y, domain_w, 0.30, domain,
                 size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # OVS pills inside the domain box
        ovs_list = OVS_BY_DOMAIN[domain]
        pill_y = domain_y + 0.38
        pill_h = 0.22
        pill_gap = 0.05
        for j, ovs in enumerate(ovs_list):
            py = pill_y + j * (pill_h + pill_gap)
            add_rect(slide, x + 0.08, py, domain_w - 0.16, pill_h,
                     fill=WHITE, line=color, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.40)
            add_text(slide, x + 0.08, py, domain_w - 0.16, pill_h, ovs,
                     size=7.5, bold=True, color=DARK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Middle: arrows / connector text
    conn_y = domain_y + domain_h + 0.10
    add_text(slide, 0.40, conn_y, 12.53, 0.22,
             "↕  many-to-many mapping  —  one OVS uses several DVS · one DVS serves several OVS  (see matrix slide)",
             size=9, italic=True, color=GRAY, align=PP_ALIGN.CENTER)

    # Bottom: 5 DVS bands
    dvs_label_y = conn_y + 0.30
    add_text(slide, 0.40, dvs_label_y, 12.53, 0.22,
             "ITBA DEVELOPMENT VALUE STREAMS (5)  ·  grouped by SOLUTION FAMILY (lifecycle · stack · skillset · cadence)",
             size=10, bold=True, color=NAVY)

    dvs_y = dvs_label_y + 0.30
    dvs_h = 1.20
    n_dvs = 5
    dvs_w = (total_w - (n_dvs - 1) * gap) / n_dvs

    for i, dvs in enumerate(DVS_LIST):
        x = margin + i * (dvs_w + gap)
        # Header
        add_rect(slide, x, dvs_y, dvs_w, 0.50, fill=dvs["color"],
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.10)
        add_text(slide, x, dvs_y + 0.05, dvs_w, 0.22, "DVS", size=7.5, bold=True,
                 color=WHITE, align=PP_ALIGN.CENTER)
        add_text(slide, x, dvs_y + 0.22, dvs_w, 0.28, dvs["name"],
                 size=8.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        # Body — apps + flag
        add_rect(slide, x, dvs_y + 0.50, dvs_w, dvs_h - 0.50, fill=WHITE,
                 line=dvs["color"], shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
        add_text(slide, x + 0.05, dvs_y + 0.55, dvs_w - 0.10, 0.22, dvs["tag"],
                 size=7, bold=True, italic=True, color=dvs["color"], align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.05, dvs_y + 0.78, dvs_w - 0.10, 0.40, dvs["apps"],
                 size=6.5, color=DARK, align=PP_ALIGN.CENTER)
        if dvs["boundary"]:
            add_text(slide, x + 0.05, dvs_y + dvs_h - 0.30, dvs_w - 0.10, 0.25,
                     "⚠ boundary issue", size=7, bold=True, color=WARN, align=PP_ALIGN.CENTER)

    # Enabling Team band at bottom
    et_y = dvs_y + dvs_h + 0.20
    add_rect(slide, 0.40, et_y, 12.53, 0.42, fill=RED,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.20)
    add_text(slide, 0.40, et_y, 12.53, 0.42,
             "Enabling Team  ·  shared services consumed by all 5 DVS  "
             "(SWF · demand intake · test/release · third-party risk · ALM tooling)",
             size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Footer
    add_text(slide, 0.40, et_y + 0.50, 12.53, 0.20,
             "Pattern: SAFe-faithful identification — solution-family DVSs · many-to-many to business OVSs · "
             "boundary issues (⚠) require executive decision on org placement",
             size=8, italic=True, color=GRAY)


# ---------------- CARDS v3: 5 DVS cards ----------------

def build_cards_v3(slide):
    add_text(slide, 0.40, 0.10, 12.53, 0.38,
             "PROPOSED VALUE STREAM STRUCTURE — 5 DVSs",
             size=20, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, 0.40, 0.58, 12.53, 0.23,
             "GROUPED BY SOLUTION FAMILY  ·  3 BOUNDARY ISSUES FLAGGED  ·  + ENABLING TEAM",
             size=11, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    n = 5
    margin = 0.40
    gap = 0.10
    total_w = 13.33 - 2 * margin
    card_w = (total_w - (n - 1) * gap) / n  # ≈ 2.43

    card_y = 1.10
    card_h = 5.45

    for i, dvs in enumerate(DVS_LIST):
        x = margin + i * (card_w + gap)

        # Body
        body = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                      Inches(x), Inches(card_y + 0.08),
                                      Inches(card_w), Inches(card_h - 0.08))
        body.fill.solid(); body.fill.fore_color.rgb = WHITE
        body.line.color.rgb = dvs["color"]; body.line.width = Pt(0.75)
        body.adjustments[0] = 0.04

        # Header
        hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                     Inches(x - 0.02), Inches(card_y),
                                     Inches(card_w + 0.04), Inches(0.85))
        hdr.fill.solid(); hdr.fill.fore_color.rgb = dvs["color"]
        hdr.line.fill.background(); hdr.adjustments[0] = 0.10

        add_text(slide, x, card_y + 0.06, card_w, 0.20,
                 f"DVS {i+1}", size=7.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(slide, x, card_y + 0.25, card_w, 0.55, dvs["name"],
                 size=10.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        # Tag
        inner_y = card_y + 1.00
        add_text(slide, x + 0.08, inner_y, card_w - 0.16, 0.22, dvs["tag"],
                 size=8, bold=True, italic=True, color=dvs["color"], align=PP_ALIGN.CENTER)

        # Rationale label + content
        add_text(slide, x + 0.08, inner_y + 0.30, card_w - 0.16, 0.18,
                 "SOLUTION-FAMILY RATIONALE", size=7, bold=True, color=DARK, align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.08, inner_y + 0.50, card_w - 0.16, 0.65, dvs["rationale"],
                 size=8, italic=True, color=DARK, align=PP_ALIGN.CENTER)

        # Divider
        add_rect(slide, x + 0.20, inner_y + 1.20, card_w - 0.40, 0.02,
                 fill=dvs["color"])

        # Apps label + content
        add_text(slide, x + 0.08, inner_y + 1.28, card_w - 0.16, 0.18,
                 "APPLICATIONS", size=7, bold=True, color=DARK, align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.08, inner_y + 1.48, card_w - 0.16, 1.40, dvs["apps"],
                 size=8, color=DARK, align=PP_ALIGN.CENTER)

        # Lead
        add_rect(slide, x + 0.20, inner_y + 2.92, card_w - 0.40, 0.02,
                 fill=dvs["color"])
        add_text(slide, x + 0.08, inner_y + 3.00, card_w - 0.16, 0.18,
                 "LEAD", size=7, bold=True, color=DARK, align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.08, inner_y + 3.20, card_w - 0.16, 0.50, dvs["lead"],
                 size=7.5, italic=True, color=DARK, align=PP_ALIGN.CENTER)

        # Boundary flag (if any)
        if dvs["boundary"]:
            add_rect(slide, x + 0.05, card_y + card_h - 0.85, card_w - 0.10, 0.75,
                     fill=WARN, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.15)
            add_text(slide, x + 0.08, card_y + card_h - 0.80, card_w - 0.16, 0.20,
                     "⚠ BOUNDARY ISSUE", size=7, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            add_text(slide, x + 0.08, card_y + card_h - 0.60, card_w - 0.16, 0.55,
                     dvs["boundary"], size=7, italic=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Enabling Team band
    et_y = card_y + card_h + 0.15
    add_rect(slide, margin, et_y, total_w, 0.50, fill=ENAB,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.20)
    add_text(slide, margin, et_y, total_w, 0.50,
             "ENABLING TEAM  ·  shared services across all 5 DVS  "
             "(SWF · demand intake · test/release · third-party risk · ALM tooling)  ·  no PM",
             size=9.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ---------------- MATRIX (OVS × DVS) ----------------

def build_matrix(slide):
    add_text(slide, 0.40, 0.10, 12.53, 0.38,
             "MANY-TO-MANY MATRIX — OVS × DVS",
             size=20, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, 0.40, 0.58, 12.53, 0.23,
             "WHICH ITBA DVS SUPPORTS WHICH BUSINESS OVS  ·  ● = primary support  ○ = supporting role",
             size=11, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # Layout: rows = OVSs (grouped by domain), cols = 5 DVSs
    table_x = 0.40
    table_y = 1.20
    label_w = 3.00
    col_w = 1.85
    row_h = 0.30
    header_h = 0.85

    dvs_names = [d["name"] for d in DVS_LIST]
    dvs_colors = {d["name"]: d["color"] for d in DVS_LIST}

    # Header row: empty corner + 5 DVS column headers
    add_rect(slide, table_x, table_y, label_w, header_h, fill=NAVY,
             shape=MSO_SHAPE.RECTANGLE)
    add_text(slide, table_x, table_y, label_w, header_h,
             "BUSINESS OVS\n↓  by domain", size=9, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    for i, dvs_name in enumerate(dvs_names):
        cx = table_x + label_w + i * col_w
        add_rect(slide, cx, table_y, col_w, header_h,
                 fill=dvs_colors[dvs_name], shape=MSO_SHAPE.RECTANGLE)
        add_text(slide, cx, table_y + 0.05, col_w, 0.20, f"DVS {i+1}",
                 size=7, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(slide, cx, table_y + 0.25, col_w, 0.55, dvs_name,
                 size=8.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Body rows
    cur_y = table_y + header_h
    row_idx = 0
    for domain, ovs_list in OVS_BY_DOMAIN.items():
        domain_color = DOMAIN_COLORS[domain]
        # Domain label spanning across all rows of this domain
        domain_h = len(ovs_list) * row_h
        add_rect(slide, table_x, cur_y, 0.60, domain_h, fill=domain_color,
                 shape=MSO_SHAPE.RECTANGLE)
        add_text(slide, table_x, cur_y, 0.60, domain_h, domain,
                 size=7, bold=True, color=WHITE,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        for ovs in ovs_list:
            # OVS name cell
            add_rect(slide, table_x + 0.60, cur_y, label_w - 0.60, row_h,
                     fill=WHITE if row_idx % 2 == 0 else LIGHTBG,
                     line=GRAY, shape=MSO_SHAPE.RECTANGLE)
            add_text(slide, table_x + 0.65, cur_y, label_w - 0.65, row_h, ovs,
                     size=8, color=DARK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

            # DVS dot cells
            supports = OVS_DVS_MAP.get(ovs, [])
            for i, dvs_name in enumerate(dvs_names):
                cx = table_x + label_w + i * col_w
                add_rect(slide, cx, cur_y, col_w, row_h,
                         fill=WHITE if row_idx % 2 == 0 else LIGHTBG,
                         line=GRAY, shape=MSO_SHAPE.RECTANGLE)
                if dvs_name in supports:
                    # First in list = primary (●), others = supporting (○)
                    is_primary = (supports[0] == dvs_name)
                    symbol = "●" if is_primary else "○"
                    color = dvs_colors[dvs_name] if is_primary else DARK
                    add_text(slide, cx, cur_y, col_w, row_h, symbol,
                             size=14 if is_primary else 12, bold=True, color=color,
                             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

            cur_y += row_h
            row_idx += 1

    # Footer
    add_text(slide, 0.40, cur_y + 0.20, 12.53, 0.40,
             "Reading the matrix:  Each OVS uses one PRIMARY DVS (●) and may use SUPPORTING DVSs (○).  "
             "The matrix surfaces solution boundaries and reveals which DVSs are most leveraged "
             "(SAP Platform · Engineering & Spatial · Enterprise Content & Reg).",
             size=9, italic=True, color=DARK)


# ---------------- BOUNDARY ISSUES INLINE (handled in Cards v3) ----------------
# The boundary flags are already on the Cards slide. No separate slide needed.


# ---------------- DVS detail rebuild (5 slides) ----------------

DVS_DETAIL = [
    {  # DVS 1 — Engineering & Spatial
        0: "DEVELOPMENT VALUE STREAM (DVS) 1 / 5",
        1: "ENGINEERING & SPATIAL — RELIABLE ASSET DATA · DESIGN · ANALYSIS PLATFORMS",
        3: "ITBA · George Porto (lead) + crossover for ProjectWise/Bentley from G&E team",
        5: "Solution family: CAD / GIS / asset · vendor-driven · quarterly upgrade cycle",
        8: ("Build and evolve the engineering & spatial platform that captures, models and publishes "
            "asset and design data — supporting Network Build, Asset Replacement, Capital Planning and "
            "Grid Optimization OVSs."),
        11: "APPLICATIONS (SOLUTION FAMILY)",
        12: ("•  ArcFM (GIS / asset management)\n"
             "•  ProjectWise (engineering content / drawings)\n"
             "•  Microstation (CAD)\n"
             "•  Bentley (engineering platform family)\n"
             "•  CYME (power systems analysis)\n"
             "•  SAP HANA spatial · GIS web apps"),
        15: "OVSs SUPPORTED",
        16: ("•  Network Build (primary)\n"
             "•  Asset Replacement (primary)\n"
             "•  Capital Planning (primary)\n"
             "•  Asset Decommissioning (primary)\n"
             "•  Grid Optimization (supporting role)"),
        17: "Multi-OVS support — strongest leverage in Engineering & Asset domain",
        20: "TEAM SHAPE",
        21: ("•  Product Manager (NEW)\n"
             "•  Tech Lead — George Porto\n"
             "•  CAD/GIS engineers — including crossover from current G&E team for Bentley/Microstation\n"
             "•  Squad size · to validate"),
        22: "Cross-team scope — clean PM accountability requires crossing today's GIS/G&E line",
        25: "VENDOR LANDSCAPE",
        26: ("•  Bentley (ProjectWise / Microstation / CYME)\n"
             "•  Esri / ArcFM partner\n"
             "•  SAP HANA partner\n"
             "•  SWF (Software Factory) — execution muscle"),
        27: "Avangrid = product owner · vendor = delivery muscle",
    },
    {  # DVS 2 — Grid Operations (OT)
        0: "DEVELOPMENT VALUE STREAM (DVS) 2 / 5  ·  ⚠ BOUNDARY ISSUE",
        1: "GRID OPERATIONS (OT) — REAL-TIME · SAFETY-CRITICAL · 24/7 SCADA & CONTROL",
        3: "ITBA · Bruno Moreto (lead) + OT specialists (NEW capability if kept in IT)",
        5: "Solution family: SCADA / control / real-time · OT skillset · DevSecOps + SRE cadence",
        8: ("Build and evolve the real-time grid control platform — SCADA, switching workflow, load "
            "analysis. Safety-critical. Supports Outage Response, Grid Optimization, Routine Switching, "
            "Storm Restoration OVSs."),
        11: "APPLICATIONS (SOLUTION FAMILY)",
        12: ("•  Cimplicity (SCADA / HMI)\n"
             "•  Echelon (control / metering platform)\n"
             "•  SLAM (System Loading and Analysis Module — to confirm)\n"
             "•  Switching Orders (utility-specific switching workflow)"),
        15: "OVSs SUPPORTED",
        16: ("•  Outage Response (primary)\n"
             "•  Grid Optimization (primary)\n"
             "•  Routine Switching (primary)\n"
             "•  Storm Restoration (primary)"),
        17: "All 4 OVSs in Network Operations domain depend on this DVS",
        20: "⚠ BOUNDARY ISSUE — OT vs IT",
        21: ("•  Grid control / SCADA traditionally sits in OT (Operational Technology), not IT\n"
             "•  Avangrid: where does this team currently live? — to confirm with Bruno\n"
             "•  If kept in ITBA: needs OT skillset (DevSecOps + SRE) — new capability hire\n"
             "•  If moved to OT: ITBA scope shrinks · governance simplified · cross-org coordination needed"),
        22: "Decision required — escalate to executive (Carlos + OT leadership)",
        25: "VENDOR LANDSCAPE",
        26: ("•  GE (Cimplicity)\n"
             "•  Echelon vendor\n"
             "•  SLAM vendor (to confirm)\n"
             "•  Possibly broader OT integrator if moved to OT"),
        27: "Avangrid = product owner · vendor mix depends on org placement",
    },
    {  # DVS 3 — Field Workforce
        0: "DEVELOPMENT VALUE STREAM (DVS) 3 / 5  ·  ⚠ BOUNDARY ISSUE",
        1: "FIELD WORKFORCE — MOBILE · SCHEDULING · WORKFORCE MANAGEMENT FOR FIELD CREWS",
        3: "ITBA · TBD — overlaps Bruno's team and CS Apps field-ops",
        5: "Solution family: workforce mgmt · mobile · scheduling · field crew customer",
        8: ("Build and evolve the field workforce platform — dispatch, mobile apps, scheduling, "
            "workforce management — for line workers, inspectors and dispatchers."),
        11: "APPLICATIONS (SOLUTION FAMILY)",
        12: ("•  Lemur-SF (Salesforce-based field service)\n"
             "•  FWM (Field Work Management)\n"
             "•  ARCOS (resource scheduling / callout)\n"
             "•  Field Force (workforce mobile apps)"),
        15: "OVSs SUPPORTED",
        16: ("•  Outage Response (supporting — field repair)\n"
             "•  Network Build (supporting — construction crews)\n"
             "•  Asset Replacement (supporting — replacement crews)\n"
             "•  Storm Restoration (supporting — surge capacity)\n"
             "•  Asset Decommissioning (supporting)"),
        17: "Cross-cutting DVS · serves multiple OVSs across Engineering and Network Ops domains",
        20: "⚠ BOUNDARY ISSUE — overlap with CS Apps Field Operations VS",
        21: ("•  These apps appeared in CS Apps' \"Field Operations\" orange card (initial proposal)\n"
             "•  In reality the customer is field crews — a single workforce platform makes sense\n"
             "•  Question: should this DVS sit in ITBA or in CS Apps?\n"
             "•  Or: should it become a CROSS-AVANGRID platform (joint CS+ITBA ownership)?\n"
             "•  Decision required — escalate to executive (Kumar + Carlos)"),
        22: "Don't fit-force into one branch — surface the structural overlap",
        25: "VENDOR LANDSCAPE",
        26: ("•  Salesforce (Lemur-SF)\n"
             "•  ARCOS Inc.\n"
             "•  Field Force vendor\n"
             "•  FWM vendor (to confirm)"),
        27: "Vendor mix to be confirmed once ownership is decided",
    },
    {  # DVS 4 — Enterprise Content & Regulatory
        0: "DEVELOPMENT VALUE STREAM (DVS) 4 / 5",
        1: "ENTERPRISE CONTENT & REGULATORY — WORKFLOW · COMPLIANCE · CORPORATE USERS",
        3: "ITBA · Sarath Garikapati + Amit (lead)",
        5: "Solution family: workflow · content · compliance-driven · corporate user customer",
        8: ("Deliver and maintain the workflow, content management and regulatory technology platform — "
            "supporting Hire-to-Retire, Source-to-Pay, Rate Case Filing, Regulatory Reporting OVSs."),
        11: "APPLICATIONS (SOLUTION FAMILY)",
        12: ("•  OpenText (content / records management)\n"
             "•  Rate Case Advisor (regulatory rate cases)\n"
             "•  Utilisphere (utility-specific platform)\n"
             "•  Corporate applications portfolio\n"
             "•  Employee-facing portals · regulatory reporting tools"),
        15: "OVSs SUPPORTED",
        16: ("•  Hire-to-Retire (primary)\n"
             "•  Rate Case Filing (primary)\n"
             "•  Regulatory Reporting (primary)\n"
             "•  Source-to-Pay (primary — workflow side)\n"
             "•  Storm Restoration (supporting — outage notifications)"),
        17: "Heaviest-leveraged DVS for Corporate & Regulatory domain",
        20: "TEAM SHAPE",
        21: ("•  Product Manager (NEW)\n"
             "•  Tech Lead — Sarath Garikapati / Amit\n"
             "•  5–6 PMs + 2 tech leads (current composition)\n"
             "•  Squad size · to validate"),
        22: "Largest sub-team — possible internal sub-streams (HR vs Regulatory) to validate",
        25: "VENDOR LANDSCAPE",
        26: ("•  OpenText vendor\n"
             "•  Accenture (corporate apps + projects)\n"
             "•  Regulatory tooling vendors\n"
             "•  SWF (Software Factory)"),
        27: "Avangrid = product owner · vendor = delivery muscle",
    },
    {  # DVS 5 — SAP Enterprise Platform
        0: "DEVELOPMENT VALUE STREAM (DVS) 5 / 5  ·  ⚠ BOUNDARY ISSUE",
        1: "SAP ENTERPRISE PLATFORM — UNIFIED SAP STACK · US (CS Apps) + SPAIN (ITBA)",
        3: "Vanessa Colombo + Michael Dubois (Spain)  +  Vincent Lam / Naggi Reddy (US — currently CS Apps)",
        5: "Solution family: SAP modules · same SAP skillset · same transport cadence · same architecture",
        8: ("Operate and evolve the SAP enterprise platform — Spain side (FI/PO/Budget/S4HANA) and "
            "US side (ECC/CRM/FICA). Single stack, single skillset — natural candidate for unified DVS."),
        11: "APPLICATIONS (SOLUTION FAMILY)",
        12: ("ITBA scope (Spain side):\n"
             "•  SAP FI (Finance)\n"
             "•  SAP PO (Purchase Orders)\n"
             "•  SAP Budget control\n"
             "•  S/4HANA (Iberdrola roadmap)\n\n"
             "Currently in CS Apps (US side):\n"
             "•  SAP ECC · SAP CRM · SAP FICA · SAP Metering"),
        15: "OVSs SUPPORTED",
        16: ("Spain:\n"
             "•  Procure-to-Pay (Spain) (primary)\n"
             "•  Period Close (Spain) (primary)\n"
             "•  Plan-to-Budget (Spain) (primary)\n\n"
             "US (cross-branch):\n"
             "•  Source-to-Pay (supporting)\n"
             "•  Regulatory Reporting (supporting)\n"
             "•  Customer-Service M2C OVS (full — currently CS Apps)"),
        17: "Highest-leverage DVS if US+Spain consolidated",
        20: "⚠ BOUNDARY ISSUE — cross-Avangrid SAP consolidation",
        21: ("•  Today: US SAP under Vincent Lam (CS Apps) · Spain SAP under Vanessa (ITBA)\n"
             "•  Same stack · same skillset · same transports — natural to unify\n"
             "•  Question: should there be ONE SAP Enterprise Platform DVS spanning Digital Products?\n"
             "•  Pro: scale benefits · skill pool · architecture coherence\n"
             "•  Con: current org boundaries · competing priorities (M2C vs Iberdrola alignment)\n"
             "•  Decision required — escalate to executive (Kumar + Carlos + Director)"),
        22: "Naggi Reddy (Sr. Mgr SAP Basis) is well-positioned as platform anchor",
        25: "VENDOR LANDSCAPE",
        26: ("•  Accenture (Global SAP support · both sides)\n"
             "•  SAP partner for Spain modules\n"
             "•  SAP partner for US modules\n"
             "•  Iberdrola Spain integration teams (peer)\n"
             "•  SWF (Software Factory)"),
        27: "Cross-branch SAP partner consolidation also possible",
    },
]


# ---------------- main ----------------

def main():
    prs = Presentation(SRC)

    # Current slide layout (from v2):
    # 0=cover, 1=combined, 2-5=4 blue VS detail, 6=Big Pic, 7=Cards, 8=Framing,
    # 9=OVS Multi, 10-13=4 OVS detail, 14-17=4 DVS detail, 18=Enabling

    # Step 1: Drop slides 14-17 (4 DVS detail) — will be rebuilt as 5 DVS detail slides
    for _ in range(4):
        delete_slide_safe(prs, 14)
    # Now: 0=cover, 1=combined, 2-5=4 blue VS, 6=Big Pic, 7=Cards, 8=Framing,
    #      9=OVS Multi, 10-13=4 OVS detail, 14=Enabling

    # Step 1.5: Clone the Enabling Team slide (idx 14) 5 times BEFORE any other reorders.
    # Enabling uses the DVS-detail template (~28 shapes). Each clone is appended at end.
    for i in range(5):
        new = clone_slide(prs, 14)
        edit_by_idx(new, DVS_DETAIL[i])
    # Now: 0..14 = original, 15..19 = 5 new DVS detail clones (still need to be moved)

    # Step 2: Clear and rebuild slide 6 (Big Picture)
    big = prs.slides[6]
    clear_slide(big)
    copy_bg(prs, 6, big)
    build_big_picture_v3(big)

    # Step 3: Clear and rebuild slide 7 (Cards)
    cards = prs.slides[7]
    clear_slide(cards)
    copy_bg(prs, 7, cards)
    build_cards_v3(cards)

    # Step 4: Edit slide 8 (Framing) to reflect new approach
    edit_by_idx(prs.slides[8], {
        0: "VALUE STREAM FRAMING — SAFe-FAITHFUL APPROACH",
        1: "DVSs GROUPED BY SOLUTION FAMILY  ·  MANY-TO-MANY WITH OVSs  ·  3 BOUNDARY ISSUES SURFACED",
        3: "Director ITBA · Carlos Garbin  ·  5 DVS Product Managers (NEW)",
        5: "Solution-family pattern · matrix many-to-many",
        8: ("The earlier 4-VS proposal mirrored today's team structure. A SAFe-faithful identification "
            "groups DVSs by SOLUTION FAMILY (lifecycle · stack · skillset · cadence), not by current "
            "team boundaries. The result: 5 DVSs with 3 boundary issues that need executive decisions."),
        11: "WHY 5 DVSs (NOT 4)",
        12: ("•  Engineering & Spatial — CAD/GIS/asset platforms (vendor-driven, quarterly upgrades)\n"
             "•  Grid Operations (OT) — real-time, safety-critical, OT skillset  ⚠\n"
             "•  Field Workforce — mobile, scheduling, field crew customer  ⚠\n"
             "•  Enterprise Content & Reg — workflow, compliance, corporate users\n"
             "•  SAP Enterprise Platform — same SAP stack across US+Spain  ⚠\n"
             "Same solution-family lifecycle = same DVS · regardless of current team home"),
        15: "WHY MANY-TO-MANY (NOT 1:1)",
        16: ("•  ONE business OVS uses several DVSs (e.g., Network Build needs Engineering + Field + Content)\n"
             "•  ONE DVS supports several OVSs (e.g., SAP Platform supports P2P, Period Close, Budget, "
             "Source-to-Pay, M2C billing)\n"
             "•  Forcing 1:1 hides this complexity and creates artificial silos\n"
             "•  Matrix on next slide makes the relationships explicit"),
        17: "Many-to-many is the SAFe pattern · 1:1 is convenient but inaccurate",
        20: "THE 3 BOUNDARY ISSUES",
        21: ("•  ⚠ Grid Operations — OT vs IT?  Traditionally OT (DevSecOps/SRE skillset)\n"
             "•  ⚠ Field Workforce — ITBA or CS Apps?  Apps appeared in BOTH initial proposals (orange Field Operations + blue ITBA)\n"
             "•  ⚠ SAP Platform — US (CS Apps) + Spain (ITBA) consolidation?  Same stack, same skillset\n"
             "Each surfaces a structural choice that benefits from explicit executive decision"),
        22: "Structural insights · merit executive escalation · don't bury under fit-forced 1:1",
        25: "WHAT THIS GIVES LEADERSHIP",
        26: ("•  A solution-family DVS structure that scales (same DVS works at 26 FTE or 100 FTE)\n"
             "•  Honest acknowledgment of cross-team / cross-org realities\n"
             "•  Three explicit decisions to make (OT/IT, Field ownership, SAP consolidation)\n"
             "•  Path to consolidate Digital Products (CS Apps + ITBA) under shared platform DVSs\n"
             "•  Foundation for many-to-many resource allocation (matrix model)"),
        27: "Same FTE base · clearer solution-family alignment · structural decisions surfaced",
    })

    # Step 5: Add NEW Matrix slide — clone slide 8 (Framing) for blank canvas
    matrix_slide = clone_slide(prs, 8)
    clear_slide(matrix_slide)
    copy_bg(prs, 8, matrix_slide)
    build_matrix(matrix_slide)
    # Matrix appended at end. Current order: 0..14 original (Enabling at 14),
    #   15..19 = 5 DVS clones, 20 = Matrix
    # Move Matrix from end to position 9 (right after Framing).
    move_slide(prs, len(prs.slides) - 1, 9)
    # After this move: 0..8 unchanged, 9=Matrix, 10=OVS Multi (was 9),
    #   11-14=OVS detail (was 10-13), 15=Enabling (was 14), 16-20=DVS clones (was 15-19)

    # Step 6: Edit slide 10 (OVS Multi-Domain View) to mention boundaries
    edit_by_idx(prs.slides[10], {
        0: "OPERATIONAL VALUE STREAMS — BUSINESS-OWNED",
        1: "ITBA SUPPORTS MULTIPLE OVSs PER DOMAIN — EXAMPLES BELOW (DETAIL IN NEXT 4 SLIDES)",
        3: "Director ITBA · Carlos Garbin · OVSs are business-owned",
        5: "Multi-OVS · multi-DVS support",
        8: ("Each business domain has multiple operational value streams (OVSs). The next 4 slides "
            "detail one example OVS per domain to illustrate the chevron flow + supporting DVSs."),
        11: "ENGINEERING & ASSET MANAGEMENT — example OVSs",
        12: ("•  Network Build (detail in next slide)\n"
             "•  Asset Replacement\n"
             "•  Capital Planning\n"
             "•  Asset Decommissioning\n"
             "Primary DVS support: Engineering & Spatial · Field Workforce"),
        15: "NETWORK OPERATIONS — example OVSs",
        16: ("•  Outage Response (detail in next slide)\n"
             "•  Grid Optimization\n"
             "•  Routine Switching\n"
             "•  Storm Restoration\n"
             "Primary DVS support: Grid Operations (OT) · Field Workforce"),
        17: "Owners listed are business stakeholders — ITBA's DVSs support them",
        20: "CORPORATE & REGULATORY — example OVSs",
        21: ("•  Rate Case Filing (detail in next slide)\n"
             "•  Source-to-Pay\n"
             "•  Hire-to-Retire\n"
             "•  Regulatory Reporting\n"
             "Primary DVS support: Enterprise Content & Reg · SAP Enterprise Platform"),
        22: "Each OVS has its own demand pattern, stakeholders and SLA shape",
        25: "IBERDROLA — example OVSs",
        26: ("•  Procure-to-Pay (Spain) (detail in next slide)\n"
             "•  Period Close (Spain)\n"
             "•  Plan-to-Budget (Spain)\n"
             "Primary DVS support: SAP Enterprise Platform"),
        27: "All OVSs supported by the same Enabling Team for shared services",
    })

    # Step 7: Slides 11-14 (OVS detail) — keep as-is, they show the chevron flow per OVS
    # No modifications needed.

    # Step 8: 5 DVS detail clones are now at positions 16-20, Enabling at 15.
    # Want them at 15-19 with Enabling at 20. Move Enabling from 15 → 20.
    move_slide(prs, 15, 20)
    # After: 0-14 unchanged, 15-19 = DVS clones (shifted up by 1), 20 = Enabling (moved from 15)

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
