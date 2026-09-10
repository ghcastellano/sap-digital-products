"""
build_itba_v5.py — ITBA deck with: 4 initial value streams + SAFe exercise.

Final structure (20 slides):
  PART 1 — Initial proposal (4 ITBA value streams from initial document):
    1.  Cover
    2.  Initial Cards (umbrella OVS + 4 VS + Enabling band) — comprehensive view
    3-6. 4 VS detail (Engineering & Asset · Network Ops & Control · Corporate & Reg · Iberdrola)

  PART 2 — SAFe exercise (deeper analysis · 5 DVSs by solution family):
    7.  SAFe Cards (5 DVS solution family + Enabling band, with boundary flags)
    8.  SAFe Big Picture (5 DVSs + multiple OVSs + Enabling)
    9.  SAFe Framing
    10. Many-to-many Matrix (OVS × DVS)
    11-14. 4 OVS detail (chevron flow — Network Build · Outage Response · Rate Case · P2P Spain)
    15-19. 5 DVS detail by solution family
    20. Enabling Team
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
BLUE4 = RGBColor(0xA5, 0xC2, 0xDA)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK  = RGBColor(0x2C, 0x3E, 0x50)
GRAY  = RGBColor(0x95, 0xA5, 0xA6)
RED   = RGBColor(0x98, 0x00, 0x03)
ENAB  = RGBColor(0x6C, 0x7A, 0x89)
LIGHTBG = RGBColor(0xEE, 0xEE, 0xEE)
WARN  = RGBColor(0xE6, 0x7E, 0x22)
GREEN = RGBColor(0x9C, 0xC4, 0x4A)
TEAL  = RGBColor(0x0B, 0x76, 0xA0)

# 5 DVS colors (solution family — Part 2)
DVS5 = {
    "Engineering & Spatial":     RGBColor(0x16, 0xA0, 0x85),
    "Grid Operations (OT)":      RGBColor(0xC0, 0x39, 0x2B),
    "Field Workforce":           RGBColor(0xF3, 0x9C, 0x12),
    "Enterprise Content & Reg":  RGBColor(0x8E, 0x44, 0xAD),
    "SAP Enterprise Platform":   RGBColor(0x2C, 0x57, 0x8B),
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
    "Engineering & Asset": ["Network Build", "Asset Replacement", "Capital Planning", "Asset Decommissioning"],
    "Network Operations": ["Outage Response", "Grid Optimization", "Routine Switching", "Storm Restoration"],
    "Corporate & Regulatory": ["Source-to-Pay", "Hire-to-Retire", "Rate Case Filing", "Regulatory Reporting"],
    "Iberdrola": ["Procure-to-Pay (Spain)", "Period Close (Spain)", "Plan-to-Budget (Spain)"],
}

# Many-to-many mapping
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


# ============== helpers ==============

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


# ============== Initial Cards (Part 1) — 4 ITBA VS comprehensive ==============

def build_initial_cards(slide):
    add_text(slide, 0.10, 0.10, 12.53, 0.38,
             "PROPOSED VALUE STREAMS — IT BUSINESS APPLICATIONS",
             size=20, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, 0.40, 0.53, 12.53, 0.23,
             "INITIAL PROPOSAL  ·  4 VALUE STREAMS  ·  + ENABLING TEAM",
             size=11, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # OVS umbrella card
    add_rect(slide, 0.51, 1.37, 12.33, 1.21, fill=WHITE,
             line=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
    add_rect(slide, 0.49, 1.32, 12.37, 0.50, fill=NAVY,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
    add_text(slide, 0.71, 1.40, 9.23, 0.40, "IT Business Applications",
             size=13, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, 9.74, 1.42, 2.90, 0.36, "ITBA UMBRELLA · 4 VS",
             size=9, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)
    add_text(slide, 0.71, 1.94, 11.93, 0.20,
             "DOMAIN · Asset Management · Network Operations · Corporate / Regulatory · Iberdrola",
             size=8.5, bold=True, italic=True, color=DARK)
    add_text(slide, 0.71, 2.16, 11.93, 0.34,
             "Operate and evolve the technology platforms that support Avangrid's enterprise operations "
             "across asset management, network operations, corporate functions and Iberdrola integration",
             size=11, italic=True, color=DARK)

    # Banner
    add_rect(slide, 0.51, 2.76, 12.33, 0.55, fill=NAVY,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.10)
    add_text(slide, 0.51, 2.76, 12.33, 0.55,
             "ITBA  ·  4 VALUE STREAMS",
             size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 4 VS cards
    cards = [
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
    margin = 0.51
    gap = 0.10
    total_w = 13.33 - 2 * margin
    card_w = (total_w - (n - 1) * gap) / n
    card_y = 3.55
    card_h = 2.45

    for i, (title, domain, desc, apps) in enumerate(cards):
        x = margin + i * (card_w + gap)
        add_rect(slide, x, card_y + 0.08, card_w, card_h - 0.08, fill=WHITE,
                 line=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
        add_rect(slide, x - 0.02, card_y, card_w + 0.04, 0.78, fill=NAVY,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.10)
        add_text(slide, x, card_y + 0.05, card_w, 0.22, "DEVELOPMENT VALUE STREAM",
                 size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(slide, x, card_y + 0.27, card_w, 0.50, title,
                 size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        inner_y = card_y + 0.92
        add_text(slide, x + 0.10, inner_y, card_w - 0.20, 0.22, domain,
                 size=8, bold=True, italic=True, color=DARK, align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.10, inner_y + 0.25, card_w - 0.20, 0.50, desc,
                 size=9.5, italic=True, color=DARK, align=PP_ALIGN.CENTER)
        add_rect(slide, x + 0.20, inner_y + 0.85, card_w - 0.40, 0.04, fill=NAVY)
        add_text(slide, x + 0.10, inner_y + 0.93, card_w - 0.20, 0.22,
                 "EXAMPLE APPLICATIONS", size=8, bold=True, color=DARK, align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.10, inner_y + 1.18, card_w - 0.20, 0.30, apps,
                 size=9.5, color=DARK, align=PP_ALIGN.CENTER)

    # Enabling Team band
    et_x = 0.51; et_y = 6.20; et_w = total_w; et_h = 0.90
    add_rect(slide, et_x, et_y, et_w, et_h, fill=RED,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.10)
    add_text(slide, et_x + 0.15, et_y + 0.06, 3.50, 0.22,
             "ENABLING TEAM", size=11, bold=True, color=WHITE)
    add_text(slide, et_x + 0.15, et_y + 0.30, 3.50, 0.55,
             "Shared services · supports the 4 DVS\nNo Product Manager · governed by Director ITBA",
             size=9, color=WHITE)
    add_text(slide, et_x + 3.85, et_y + 0.10, et_w - 3.95, 0.20,
             "FOUNDATION SERVICES", size=9, bold=True, color=WHITE)
    add_text(slide, et_x + 3.85, et_y + 0.32, et_w - 3.95, 0.55,
             "SWF (Software Factory)  ·  Demand Intake (PDD · Jira · IT Now)  ·  "
             "Test & Release Coordination  ·  Third-Party Risk Gating  ·  ALM Tooling",
             size=9, color=WHITE)


# ============== SAFe Cards (Part 2) — 5 DVS by solution family ==============

DVS5_LIST = [
    {
        "name": "Engineering & Spatial",
        "tag": "Asset / CAD / GIS",
        "rationale": "Vendor-driven · quarterly upgrades · CAD/GIS skillset · planner customer",
        "apps": "ArcFM · ProjectWise · Microstation · Bentley · CYME · SAP HANA spatial",
        "boundary": None,
        "lead": "George Porto + crossover (Bruno's team for ProjectWise/Bentley)",
        "color": DVS5["Engineering & Spatial"],
    },
    {
        "name": "Grid Operations (OT)",
        "tag": "SCADA / Real-time",
        "rationale": "Real-time · safety-critical · 24/7 · OT skillset (DevSecOps + SRE)",
        "apps": "Echelon · Cimplicity · SLAM · Switching Orders",
        "boundary": "⚠ Traditionally OT, not IT — needs executive decision",
        "lead": "Bruno Moreto + OT specialists (NEW capability)",
        "color": DVS5["Grid Operations (OT)"],
    },
    {
        "name": "Field Workforce",
        "tag": "Mobile / Scheduling",
        "rationale": "Mobile · scheduling · workforce mgmt · field crew customer",
        "apps": "Lemur-SF · FWM · ARCOS · Field Force",
        "boundary": "⚠ Currently shown under CS Apps Field Operations VS — overlap",
        "lead": "TBD — overlaps Bruno's team and CS Apps field-ops",
        "color": DVS5["Field Workforce"],
    },
    {
        "name": "Enterprise Content & Reg",
        "tag": "Workflow / Compliance",
        "rationale": "Transactional · workflow · compliance-driven · corporate user",
        "apps": "OpenText · Rate Case Advisor · Utilisphere · corporate apps",
        "boundary": None,
        "lead": "Sarath Garikapati + Amit",
        "color": DVS5["Enterprise Content & Reg"],
    },
    {
        "name": "SAP Enterprise Platform",
        "tag": "SAP modules",
        "rationale": "Same SAP stack · same SAP skillset · same transport cadence",
        "apps": "SAP FI · PO · Budget · S/4HANA (Spain)  +  SAP ECC · CRM · FICA (US)",
        "boundary": "⚠ Cross-Avangrid: should US (CS Apps) and Spain (ITBA) SAP be one DVS?",
        "lead": "Vanessa Colombo + Vincent Lam / Naggi Reddy (cross-branch)",
        "color": DVS5["SAP Enterprise Platform"],
    },
]


def build_safe_cards(slide):
    add_text(slide, 0.10, 0.10, 12.53, 0.38,
             "SAFe EXERCISE — 5 DVSs BY SOLUTION FAMILY",
             size=20, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, 0.40, 0.53, 12.53, 0.23,
             "GROUPED BY SOLUTION-FAMILY LIFECYCLE  ·  3 BOUNDARY ISSUES SURFACED  ·  + ENABLING TEAM",
             size=11, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    n = 5
    margin = 0.40
    gap = 0.10
    total_w = 13.33 - 2 * margin
    card_w = (total_w - (n - 1) * gap) / n

    card_y = 1.10
    card_h = 5.45

    for i, dvs in enumerate(DVS5_LIST):
        x = margin + i * (card_w + gap)
        add_rect(slide, x, card_y + 0.08, card_w, card_h - 0.08, fill=WHITE,
                 line=dvs["color"], shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
        add_rect(slide, x - 0.02, card_y, card_w + 0.04, 0.85, fill=dvs["color"],
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.10)
        add_text(slide, x, card_y + 0.06, card_w, 0.20, f"DVS {i+1}",
                 size=7.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(slide, x, card_y + 0.25, card_w, 0.55, dvs["name"],
                 size=10.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        inner_y = card_y + 1.00
        add_text(slide, x + 0.08, inner_y, card_w - 0.16, 0.22, dvs["tag"],
                 size=8, bold=True, italic=True, color=dvs["color"], align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.08, inner_y + 0.30, card_w - 0.16, 0.18,
                 "SOLUTION-FAMILY RATIONALE", size=7, bold=True, color=DARK, align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.08, inner_y + 0.50, card_w - 0.16, 0.65, dvs["rationale"],
                 size=8, italic=True, color=DARK, align=PP_ALIGN.CENTER)
        add_rect(slide, x + 0.20, inner_y + 1.20, card_w - 0.40, 0.02, fill=dvs["color"])
        add_text(slide, x + 0.08, inner_y + 1.28, card_w - 0.16, 0.18,
                 "APPLICATIONS", size=7, bold=True, color=DARK, align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.08, inner_y + 1.48, card_w - 0.16, 1.40, dvs["apps"],
                 size=8, color=DARK, align=PP_ALIGN.CENTER)
        add_rect(slide, x + 0.20, inner_y + 2.92, card_w - 0.40, 0.02, fill=dvs["color"])
        add_text(slide, x + 0.08, inner_y + 3.00, card_w - 0.16, 0.18,
                 "LEAD", size=7, bold=True, color=DARK, align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.08, inner_y + 3.20, card_w - 0.16, 0.50, dvs["lead"],
                 size=7.5, italic=True, color=DARK, align=PP_ALIGN.CENTER)

        if dvs["boundary"]:
            add_rect(slide, x + 0.05, card_y + card_h - 0.85, card_w - 0.10, 0.75,
                     fill=WARN, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.15)
            add_text(slide, x + 0.08, card_y + card_h - 0.80, card_w - 0.16, 0.20,
                     "⚠ BOUNDARY ISSUE", size=7, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            add_text(slide, x + 0.08, card_y + card_h - 0.60, card_w - 0.16, 0.55,
                     dvs["boundary"], size=7, italic=True, color=WHITE, align=PP_ALIGN.CENTER)

    et_y = card_y + card_h + 0.15
    add_rect(slide, margin, et_y, total_w, 0.50, fill=ENAB,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.20)
    add_text(slide, margin, et_y, total_w, 0.50,
             "ENABLING TEAM  ·  shared services across all 5 DVS  ·  no PM  ·  governed by Director ITBA",
             size=9.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ============== SAFe Big Picture (Part 2) ==============

def build_safe_big_picture(slide):
    add_text(slide, 0.38, 0.17, 12.53, 0.38,
             "SAFe EXERCISE — BIG PICTURE",
             size=18, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, 0.40, 0.58, 12.53, 0.23,
             "5 ITBA DVSs (solution family) supporting many business OVSs (many-to-many) · 1 Enabling Team",
             size=10.5, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # Top: 4 OVS domain groups
    add_text(slide, 0.40, 1.05, 12.53, 0.22,
             "BUSINESS OPERATIONAL VALUE STREAMS (owned by business · grouped by domain)",
             size=10, bold=True, color=NAVY)

    domain_y = 1.30
    domain_h = 1.45
    n_domains = 4
    margin = 0.40
    gap = 0.12
    total_w = 13.33 - 2 * margin
    domain_w = (total_w - (n_domains - 1) * gap) / n_domains

    for i, (domain, ovs_list) in enumerate(OVS_BY_DOMAIN.items()):
        x = margin + i * (domain_w + gap)
        color = DOMAIN_COLORS[domain]
        add_rect(slide, x, domain_y, domain_w, domain_h, fill=LIGHTBG,
                 line=color, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
        add_rect(slide, x, domain_y, domain_w, 0.30, fill=color,
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.10)
        add_text(slide, x, domain_y, domain_w, 0.30, domain,
                 size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        pill_y = domain_y + 0.38
        pill_h = 0.22
        for j, ovs in enumerate(ovs_list):
            py = pill_y + j * (pill_h + 0.05)
            add_rect(slide, x + 0.08, py, domain_w - 0.16, pill_h, fill=WHITE,
                     line=color, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.40)
            add_text(slide, x + 0.08, py, domain_w - 0.16, pill_h, ovs,
                     size=7.5, bold=True, color=DARK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Connector
    conn_y = domain_y + domain_h + 0.10
    add_text(slide, 0.40, conn_y, 12.53, 0.22,
             "↕  many-to-many mapping  —  one OVS uses several DVS · one DVS serves several OVS  (matrix in next slide)",
             size=9, italic=True, color=GRAY, align=PP_ALIGN.CENTER)

    # Bottom: 5 DVS bands
    dvs_label_y = conn_y + 0.30
    add_text(slide, 0.40, dvs_label_y, 12.53, 0.22,
             "ITBA DEVELOPMENT VALUE STREAMS (5)  ·  grouped by SOLUTION FAMILY",
             size=10, bold=True, color=NAVY)

    dvs_y = dvs_label_y + 0.30
    dvs_h = 1.20
    n_dvs = 5
    dvs_w = (total_w - (n_dvs - 1) * gap) / n_dvs

    for i, dvs in enumerate(DVS5_LIST):
        x = margin + i * (dvs_w + gap)
        add_rect(slide, x, dvs_y, dvs_w, 0.50, fill=dvs["color"],
                 shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.10)
        add_text(slide, x, dvs_y + 0.05, dvs_w, 0.22, "DVS",
                 size=7.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(slide, x, dvs_y + 0.22, dvs_w, 0.28, dvs["name"],
                 size=8.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_rect(slide, x, dvs_y + 0.50, dvs_w, dvs_h - 0.50, fill=WHITE,
                 line=dvs["color"], shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
        add_text(slide, x + 0.05, dvs_y + 0.55, dvs_w - 0.10, 0.22, dvs["tag"],
                 size=7, bold=True, italic=True, color=dvs["color"], align=PP_ALIGN.CENTER)
        add_text(slide, x + 0.05, dvs_y + 0.78, dvs_w - 0.10, 0.40, dvs["apps"],
                 size=6.5, color=DARK, align=PP_ALIGN.CENTER)
        if dvs["boundary"]:
            add_text(slide, x + 0.05, dvs_y + dvs_h - 0.30, dvs_w - 0.10, 0.25,
                     "⚠ boundary issue", size=7, bold=True, color=WARN, align=PP_ALIGN.CENTER)

    et_y = dvs_y + dvs_h + 0.20
    add_rect(slide, 0.40, et_y, 12.53, 0.42, fill=RED,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.20)
    add_text(slide, 0.40, et_y, 12.53, 0.42,
             "Enabling Team  ·  shared services consumed by all 5 DVS",
             size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ============== Matrix (OVS × DVS) ==============

def build_matrix(slide):
    add_text(slide, 0.40, 0.10, 12.53, 0.38,
             "MANY-TO-MANY MATRIX — OVS × DVS",
             size=20, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, 0.40, 0.58, 12.53, 0.23,
             "WHICH ITBA DVS SUPPORTS WHICH BUSINESS OVS  ·  ● = primary support  ○ = supporting role",
             size=11, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    table_x = 0.40
    table_y = 1.20
    label_w = 3.00
    col_w = 1.85
    row_h = 0.30
    header_h = 0.85

    dvs_names = [d["name"] for d in DVS5_LIST]
    dvs_colors = {d["name"]: d["color"] for d in DVS5_LIST}

    add_rect(slide, table_x, table_y, label_w, header_h, fill=NAVY)
    add_text(slide, table_x, table_y, label_w, header_h,
             "BUSINESS OVS\n↓  by domain", size=9, bold=True, color=WHITE,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    for i, dvs_name in enumerate(dvs_names):
        cx = table_x + label_w + i * col_w
        add_rect(slide, cx, table_y, col_w, header_h, fill=dvs_colors[dvs_name])
        add_text(slide, cx, table_y + 0.05, col_w, 0.20, f"DVS {i+1}",
                 size=7, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(slide, cx, table_y + 0.25, col_w, 0.55, dvs_name,
                 size=8.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    cur_y = table_y + header_h
    row_idx = 0
    for domain, ovs_list in OVS_BY_DOMAIN.items():
        domain_color = DOMAIN_COLORS[domain]
        domain_h = len(ovs_list) * row_h
        add_rect(slide, table_x, cur_y, 0.60, domain_h, fill=domain_color)
        add_text(slide, table_x, cur_y, 0.60, domain_h, domain,
                 size=7, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

        for ovs in ovs_list:
            add_rect(slide, table_x + 0.60, cur_y, label_w - 0.60, row_h,
                     fill=WHITE if row_idx % 2 == 0 else LIGHTBG, line=GRAY)
            add_text(slide, table_x + 0.65, cur_y, label_w - 0.65, row_h, ovs,
                     size=8, color=DARK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

            supports = OVS_DVS_MAP.get(ovs, [])
            for i, dvs_name in enumerate(dvs_names):
                cx = table_x + label_w + i * col_w
                add_rect(slide, cx, cur_y, col_w, row_h,
                         fill=WHITE if row_idx % 2 == 0 else LIGHTBG, line=GRAY)
                if dvs_name in supports:
                    is_primary = (supports[0] == dvs_name)
                    symbol = "●" if is_primary else "○"
                    color = dvs_colors[dvs_name] if is_primary else DARK
                    add_text(slide, cx, cur_y, col_w, row_h, symbol,
                             size=14 if is_primary else 12, bold=True, color=color,
                             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            cur_y += row_h
            row_idx += 1

    add_text(slide, 0.40, cur_y + 0.20, 12.53, 0.40,
             "Reading the matrix:  Each OVS uses one PRIMARY DVS (●) and may use SUPPORTING DVSs (○).",
             size=9, italic=True, color=DARK)


# ============== OVS Detail (chevron flow) ==============

def build_ovs_detail_slide(slide, *, ovs_name, owner_label, trigger_text, chevrons_with_personas,
                           outcome_text, dvs_list, apps_text, other_ovs_text):
    """Single-OVS detail with chevrons + personas + multi-DVS support indicator."""
    add_text(slide, 0.38, 0.17, 12.53, 0.38,
             f"OVS DETAIL — {ovs_name.upper()}",
             size=18, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, 0.40, 0.58, 12.53, 0.23,
             f"Example operational value stream  ·  {owner_label}",
             size=10.5, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    add_rect(slide, 0.27, 1.22, 12.66, 1.94, fill=LIGHTBG,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
    add_text(slide, 4.82, 1.28, 6.50, 0.22,
             f"OPERATIONAL VALUE STREAM — {ovs_name}",
             size=10.5, bold=True, color=NAVY, align=PP_ALIGN.LEFT)

    add_text(slide, 0.37, 1.80, 1.30, 0.21, "TRIGGER",
             size=10, bold=True, color=GRAY, align=PP_ALIGN.LEFT)
    add_text(slide, 0.37, 2.10, 1.30, 0.85, trigger_text,
             size=10, italic=True, color=DARK, align=PP_ALIGN.LEFT)

    add_text(slide, 11.60, 1.63, 1.30, 0.21, "OUTCOME",
             size=10, bold=True, color=GRAY, align=PP_ALIGN.LEFT)
    add_text(slide, 11.60, 1.92, 1.30, 1.00, outcome_text,
             size=10, italic=True, color=DARK, align=PP_ALIGN.LEFT)

    n = len(chevrons_with_personas)
    chev_x_start = 1.85
    avail_w = 11.50 - chev_x_start
    chev_overlap = 0.10
    chev_w = (avail_w + (n - 1) * chev_overlap) / n
    chev_h = 0.55
    chev_y = 2.30
    persona_y = 1.75

    for i, (step, persona) in enumerate(chevrons_with_personas):
        cx = chev_x_start + i * (chev_w - chev_overlap)
        add_text(slide, cx, persona_y, chev_w - chev_overlap, 0.40, persona,
                 size=8, italic=True, color=GRAY, align=PP_ALIGN.CENTER)
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

    # DVS area — show MULTIPLE DVS bands (since OVSs use multiple DVSs)
    dvs_bg = RGBColor(0xE6, 0xED, 0xF7)
    add_rect(slide, 0.27, 3.25, 12.66, 1.25, fill=dvs_bg,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
    add_text(slide, 4.22, 4.21, 5.26, 0.22,
             "DEVELOPMENT VALUE STREAMS — who builds & evolves the solutions",
             size=10.5, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

    # Show DVSs that support this OVS as colored bands
    n_dvs = len(dvs_list)
    band_x = 1.88
    band_total_w = avail_w - 0.10
    band_gap = 0.05
    band_w = (band_total_w - (n_dvs - 1) * band_gap) / n_dvs
    band_y = 3.55
    for i, (dvs_name, dvs_color, role) in enumerate(dvs_list):
        bx = band_x + i * (band_w + band_gap)
        add_rect(slide, bx, band_y, band_w, 0.45,
                 fill=dvs_color, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.20)
        add_text(slide, bx, band_y + 0.08, band_w, 0.30,
                 f"DVS — {dvs_name}\n({role})",
                 size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Apps row
    add_text(slide, 0.47, 4.91, 1.30, 0.21, "APPLICATIONS",
             size=10, bold=True, color=GRAY, align=PP_ALIGN.LEFT)
    add_rect(slide, band_x, 4.68, band_total_w, 0.66, fill=TEAL)
    add_text(slide, band_x, 4.68, band_total_w, 0.66, apps_text,
             size=9.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Enabling Team band
    add_rect(slide, band_x, 5.64, band_total_w, 0.45,
             fill=RED, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.10)
    add_text(slide, band_x, 5.77, band_total_w, 0.20,
             "Enabling Team  ·  shared services consumed by all DVS  "
             "(SWF · demand intake · test/release · third-party risk · ALM)",
             size=8.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    add_text(slide, 0.40, 6.25, 12.53, 0.40,
             f"Other operational value streams in this domain: {other_ovs_text}",
             size=9, italic=True, color=GRAY, align=PP_ALIGN.LEFT)


# ============== content for Part 1 — 4 initial VS detail ==============

INITIAL_VS_DETAIL = [
    {  # VS 1 — Engineering & Asset Management
        0: "DEVELOPMENT VALUE STREAM",
        1: "ENGINEERING & ASSET MANAGEMENT — RELIABLE ASSET DATA FOR PLANNERS AND ENGINEERS",
        3: "ITBA · George Porto (lead) + crossover with G&E team for ProjectWise/Bentley",
        5: "Initial proposal · domain anchor",
        8: ("Operate and evolve the engineering & asset platform — capture, model and publish reliable "
            "data so engineers and planners can design and operate the network."),
        11: "APPLICATIONS",
        12: ("•  ArcFM (GIS / asset management)\n"
             "•  ProjectWise (engineering content / drawings)\n"
             "•  Bentley · Microstation (CAD)\n"
             "•  CYME (power systems analysis)"),
        15: "USERS / DEMAND SOURCES",
        16: ("•  Engineers — design new builds / modifications\n"
             "•  Planners — forecasts and capacity planning\n"
             "•  Asset Management — lifecycle decisions\n"
             "•  Field operations — asset updates"),
        17: "Demand routed via Jira / IT Now / PDD",
        20: "TODAY'S TEAM",
        21: ("•  GIS & OSG team (George Porto) owns ArcFM, GIS web apps, SAP HANA\n"
             "•  G&E team (Bruno Moreto) currently owns ProjectWise / Bentley / Microstation\n"
             "•  Some apps cross today's team boundaries"),
        22: "Cross-team scope — value-stream framing helps clarify accountability",
        25: "VENDOR LANDSCAPE",
        26: ("•  Bentley (ProjectWise / Microstation / CYME)\n"
             "•  Esri / ArcFM partner\n"
             "•  SWF (Software Factory) — execution support"),
        27: "Multiple vendors — coordinated through ITBA leadership",
    },
    {  # VS 2 — Network Operations & Control
        0: "DEVELOPMENT VALUE STREAM",
        1: "NETWORK OPERATIONS & CONTROL — GRID OPERATOR PREVENTS AND MINIMIZES OUTAGES",
        3: "ITBA · Bruno Moreto + Andrew McGibbon (lead)",
        5: "Initial proposal · domain anchor",
        8: ("Operate the systems that monitor, switch and control the network in real time — "
            "preventing outages and minimising their impact."),
        11: "APPLICATIONS",
        12: ("•  Echelon (control / metering platform)\n"
             "•  Cimplicity (SCADA / HMI)\n"
             "•  SLAM (System Loading and Analysis Module — to confirm)\n"
             "•  Switching Orders (utility-specific switching workflow)"),
        15: "USERS / DEMAND SOURCES",
        16: ("•  Grid operators — real-time monitoring and switching\n"
             "•  Network engineers — analysis and load planning\n"
             "•  Field crews — switching execution\n"
             "•  Regulators — outage reporting"),
        17: "Real-time / safety-critical · high availability requirements",
        20: "TODAY'S TEAM",
        21: ("•  G&E Operations team (Bruno Moreto + Andrew McGibbon) is closest current home\n"
             "•  Some apps may sit in OT (Operational Technology) scope — boundary to confirm\n"
             "•  Coordination with Field Operations VS (cross-stream)"),
        22: "OT/IT boundary to confirm — apps may sit outside ITBA scope",
        25: "VENDOR LANDSCAPE",
        26: ("•  Echelon vendor\n"
             "•  GE (Cimplicity)\n"
             "•  SLAM vendor (to confirm)\n"
             "•  Internal switching workflow tooling"),
        27: "Vendor mix to validate — some grid platforms managed cross-Avangrid",
    },
    {  # VS 3 — Corporate & Regulatory
        0: "DEVELOPMENT VALUE STREAM",
        1: "CORPORATE & REGULATORY — HR / EMPLOYEE / REGULATORY USER SELF-SERVICE",
        3: "ITBA · Sarath Garikapati + Amit (lead)",
        5: "Initial proposal · domain anchor",
        8: ("Deliver and maintain regulatory and employee-facing technology across Avangrid US — "
            "operating under continuous compliance pressure."),
        11: "APPLICATIONS",
        12: ("•  OpenText (content / records management)\n"
             "•  Rate Case Advisor (regulatory rate cases)\n"
             "•  Utilisphere (utility platform)\n"
             "•  Plus broader corporate applications portfolio"),
        15: "USERS / DEMAND SOURCES",
        16: ("•  HR / Finance / Legal / Compliance\n"
             "•  Regulators — rate cases · mandated reporting\n"
             "•  Employee population — enablement & self-service\n"
             "•  Continuous new requirements"),
        17: "Compliance-driven · always-priority cadence",
        20: "TODAY'S TEAM",
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
    {  # VS 4 — Iberdrola Corporate Services
        0: "DEVELOPMENT VALUE STREAM",
        1: "IBERDROLA CORPORATE SERVICES — SPAIN FINANCE / PROCUREMENT / BUDGET",
        3: "ITBA · Vanessa Colombo + Michael Dubois (lead)",
        5: "Initial proposal · domain anchor",
        8: ("Run Spain-side SAP modules and keep US/Iberdrola corporate processes aligned — "
            "balancing incidents, small demands, and major projects under hiring constraints."),
        11: "APPLICATIONS",
        12: ("•  SAP FI (Finance) — Spain side\n"
             "•  SAP PO (Purchase Orders) — Spain side\n"
             "•  SAP Budget control\n"
             "•  S/4HANA (Iberdrola roadmap)"),
        15: "USERS / DEMAND SOURCES",
        16: ("•  Iberdrola corporate (Spain) — process changes\n"
             "•  Avangrid Finance — US-side integration needs\n"
             "•  Incidents (IT Now)\n"
             "•  Small Demands · Major projects (PDD)"),
        17: "Cross-org coordination · transport gating",
        20: "TODAY'S TEAM",
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


# ============== content for Part 2 — 5 SAFe DVS detail ==============

SAFE_DVS_DETAIL = [
    {  # DVS 1 — Engineering & Spatial
        0: "DEVELOPMENT VALUE STREAM (DVS) 1 / 5",
        1: "ENGINEERING & SPATIAL — RELIABLE ASSET DATA · DESIGN · ANALYSIS PLATFORMS",
        3: "ITBA · George Porto (lead) + crossover for ProjectWise/Bentley from G&E team",
        5: "Solution family: CAD / GIS / asset · vendor-driven · quarterly upgrade cycle",
        8: ("Build and evolve the engineering & spatial platform — captures, models and publishes "
            "asset and design data — supporting Network Build, Asset Replacement, Capital Planning OVSs."),
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
             "•  CAD/GIS engineers — including crossover from current G&E team\n"
             "•  Squad size · to validate"),
        22: "Cross-team scope — clean PM accountability requires crossing today's GIS/G&E line",
        25: "VENDOR LANDSCAPE",
        26: ("•  Bentley (ProjectWise / Microstation / CYME)\n"
             "•  Esri / ArcFM partner\n"
             "•  SAP HANA partner\n"
             "•  SWF (Software Factory)"),
        27: "Avangrid = product owner · vendor = delivery muscle",
    },
    {  # DVS 2 — Grid Operations (OT)
        0: "DEVELOPMENT VALUE STREAM (DVS) 2 / 5  ·  ⚠ BOUNDARY ISSUE",
        1: "GRID OPERATIONS (OT) — REAL-TIME · SAFETY-CRITICAL · 24/7 SCADA & CONTROL",
        3: "ITBA · Bruno Moreto (lead) + OT specialists (NEW capability if kept in IT)",
        5: "Solution family: SCADA / control / real-time · OT skillset · DevSecOps + SRE cadence",
        8: ("Build and evolve the real-time grid control platform — SCADA, switching workflow, load "
            "analysis. Safety-critical. Supports Outage Response, Grid Optimization, Switching OVSs."),
        11: "APPLICATIONS (SOLUTION FAMILY)",
        12: ("•  Cimplicity (SCADA / HMI)\n"
             "•  Echelon (control / metering platform)\n"
             "•  SLAM (System Loading and Analysis Module)\n"
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
             "•  If moved to OT: ITBA scope shrinks · governance simplified"),
        22: "Decision required — escalate to executive (Carlos + OT leadership)",
        25: "VENDOR LANDSCAPE",
        26: ("•  GE (Cimplicity)\n"
             "•  Echelon vendor\n"
             "•  SLAM vendor\n"
             "•  Possibly broader OT integrator if moved to OT"),
        27: "Vendor mix depends on org placement",
    },
    {  # DVS 3 — Field Workforce
        0: "DEVELOPMENT VALUE STREAM (DVS) 3 / 5  ·  ⚠ BOUNDARY ISSUE",
        1: "FIELD WORKFORCE — MOBILE · SCHEDULING · WORKFORCE MGMT FOR FIELD CREWS",
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
             "•  Asset Replacement (supporting)\n"
             "•  Storm Restoration (supporting — surge capacity)\n"
             "•  Asset Decommissioning (supporting)"),
        17: "Cross-cutting DVS · serves multiple OVSs across domains",
        20: "⚠ BOUNDARY ISSUE — overlap with CS Apps Field Operations VS",
        21: ("•  These apps appeared in CS Apps' \"Field Operations\" orange card (initial proposal)\n"
             "•  In reality the customer is field crews — single workforce platform makes sense\n"
             "•  Question: should this DVS sit in ITBA or in CS Apps?\n"
             "•  Or: should it become a CROSS-AVANGRID platform (joint ownership)?"),
        22: "Don't fit-force into one branch — surface the structural overlap",
        25: "VENDOR LANDSCAPE",
        26: ("•  Salesforce (Lemur-SF)\n"
             "•  ARCOS Inc.\n"
             "•  Field Force vendor\n"
             "•  FWM vendor (to confirm)"),
        27: "Vendor mix to be confirmed once ownership is decided",
    },
    {  # DVS 4 — Enterprise Content & Reg
        0: "DEVELOPMENT VALUE STREAM (DVS) 4 / 5",
        1: "ENTERPRISE CONTENT & REGULATORY — WORKFLOW · COMPLIANCE · CORPORATE USERS",
        3: "ITBA · Sarath Garikapati + Amit (lead)",
        5: "Solution family: workflow · content · compliance-driven · corporate user customer",
        8: ("Deliver and maintain the workflow, content management and regulatory technology platform — "
            "supporting Hire-to-Retire, Source-to-Pay, Rate Case Filing, Regulatory Reporting."),
        11: "APPLICATIONS (SOLUTION FAMILY)",
        12: ("•  OpenText (content / records management)\n"
             "•  Rate Case Advisor (regulatory rate cases)\n"
             "•  Utilisphere (utility-specific platform)\n"
             "•  Corporate applications portfolio\n"
             "•  Employee portals · regulatory reporting tools"),
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
        5: "Solution family: SAP modules · same skillset · same transport cadence · same architecture",
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
             "•  Customer-Service M2C OVS (full — currently CS Apps)"),
        17: "Highest-leverage DVS if US+Spain consolidated",
        20: "⚠ BOUNDARY ISSUE — cross-Avangrid SAP consolidation",
        21: ("•  Today: US SAP under Vincent Lam (CS Apps) · Spain SAP under Vanessa (ITBA)\n"
             "•  Same stack · same skillset · same transports — natural to unify\n"
             "•  Question: should there be ONE SAP Enterprise Platform DVS spanning Digital Products?\n"
             "•  Pro: scale benefits · skill pool · architecture coherence\n"
             "•  Con: current org boundaries · competing priorities"),
        22: "Naggi Reddy (Sr. Mgr SAP Basis) is well-positioned as platform anchor",
        25: "VENDOR LANDSCAPE",
        26: ("•  Accenture (Global SAP support · both sides)\n"
             "•  SAP partner for Spain modules\n"
             "•  SAP partner for US modules\n"
             "•  Iberdrola Spain integration teams (peer)"),
        27: "Cross-branch SAP partner consolidation also possible",
    },
]


# ============== main ==============

def main():
    prs = Presentation(SRC)

    # Drop Value Lens
    delete_slide_safe(prs, 14)
    # Now: 0=Cover, 1=Prop1Cards, 2-4=3 DVS detail, 5=Prop2Cards, 6=Prop2Rationale,
    #      7=Framing, 8=OVSdetail, 9-12=4 DVS detail Prop2, 13=Enabling

    # Step 1: Clone slide 4 (CCC DVS detail) to get a 4th slot for Part 1
    clone_slide(prs, 4)  # appended at idx 14
    move_slide(prs, 14, 5)  # move to position 5
    # Now: 0=Cover, 1=Prop1Cards, 2-5=4 DVS detail, 6=Prop2Cards, 7=Prop2Rationale,
    #      8=Framing, 9=OVSdetail, 10-13=4 DVS detail Prop2, 14=Enabling (15 slides)

    # Step 2: Cover
    cover = prs.slides[0]
    for sh in cover.shapes:
        if sh.has_text_frame and "PROPOSED VALUE STREAMS" in sh.text_frame.text and "CUSTOMER SERVICES" in sh.text_frame.text:
            replace_textbox(sh,
                "PROPOSED VALUE STREAMS – IT BUSINESS APPLICATIONS\n"
                "Initial proposal: 4 Value Streams  ·  SAFe Exercise: 5 DVSs by solution family\n"
                "April 2026")

    # Step 3: Rebuild slide 1 (Prop1 Cards) — Initial Cards (4 VS comprehensive)
    p1_cards = prs.slides[1]
    clear_slide(p1_cards)
    copy_bg(prs, 1, p1_cards)
    build_initial_cards(p1_cards)

    # Step 4: Edit slides 2-5 (4 initial VS detail)
    edit_by_idx(prs.slides[2], INITIAL_VS_DETAIL[0])
    edit_by_idx(prs.slides[3], INITIAL_VS_DETAIL[1])
    edit_by_idx(prs.slides[4], INITIAL_VS_DETAIL[2])
    edit_by_idx(prs.slides[5], INITIAL_VS_DETAIL[3])

    # Step 5: Rebuild slide 6 (Prop2 Cards) — SAFe Cards (5 DVS)
    safe_cards = prs.slides[6]
    clear_slide(safe_cards)
    copy_bg(prs, 6, safe_cards)
    build_safe_cards(safe_cards)

    # Step 6: Rebuild slide 7 (Prop2 Rationale) — SAFe Big Picture
    safe_bp = prs.slides[7]
    clear_slide(safe_bp)
    copy_bg(prs, 7, safe_bp)
    build_safe_big_picture(safe_bp)

    # Step 7: Edit slide 8 (Framing) — SAFe Framing
    edit_by_idx(prs.slides[8], {
        0: "VALUE STREAM FRAMING — SAFe-FAITHFUL APPROACH",
        1: "DVSs GROUPED BY SOLUTION FAMILY  ·  MANY-TO-MANY WITH OVSs  ·  3 BOUNDARY ISSUES SURFACED",
        3: "Director ITBA · Carlos Garbin  ·  5 DVS Product Managers (NEW)",
        5: "Solution-family pattern · matrix many-to-many",
        8: ("The initial 4-VS proposal mirrored today's team structure. A SAFe-faithful identification "
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
             "•  ONE DVS supports several OVSs (e.g., SAP Platform supports P2P, Period Close, Budget, M2C)\n"
             "•  Forcing 1:1 hides this complexity and creates artificial silos\n"
             "•  Matrix on next slide makes the relationships explicit"),
        17: "Many-to-many is the SAFe pattern · 1:1 is convenient but inaccurate",
        20: "THE 3 BOUNDARY ISSUES",
        21: ("•  ⚠ Grid Operations — OT vs IT?  Traditionally OT (DevSecOps/SRE skillset)\n"
             "•  ⚠ Field Workforce — ITBA or CS Apps?  Apps appeared in BOTH initial proposals\n"
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

    # Step 8: Rebuild slide 9 (OVSdetail) — Matrix
    matrix = prs.slides[9]
    clear_slide(matrix)
    copy_bg(prs, 9, matrix)
    build_matrix(matrix)

    # Step 9: Rebuild slides 10-13 — 4 OVS detail (chevron)
    ovs_specs = [
        dict(  # OVS — Network Build
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
            dvs_list=[
                ("Engineering & Spatial", DVS5["Engineering & Spatial"], "primary"),
                ("Field Workforce", DVS5["Field Workforce"], "supporting"),
                ("Enterprise Content & Reg", DVS5["Enterprise Content & Reg"], "supporting"),
            ],
            apps_text="ArcFM (asset)  ·  CYME (analysis)  ·  Bentley / Microstation (design)  ·  ProjectWise (drawings)",
            other_ovs_text="Asset Replacement Cycle  ·  Network Inspection  ·  Capital Planning  ·  Asset Decommissioning",
        ),
        dict(  # OVS — Outage Response
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
            dvs_list=[
                ("Grid Operations (OT)", DVS5["Grid Operations (OT)"], "primary"),
                ("Field Workforce", DVS5["Field Workforce"], "supporting"),
            ],
            apps_text="Cimplicity (SCADA)  ·  Echelon (control)  ·  Switching Orders (workflow)  ·  SLAM (load analysis)",
            other_ovs_text="Grid Optimization  ·  Routine Switching  ·  Load Balancing  ·  Storm Restoration",
        ),
        dict(  # OVS — Rate Case Filing
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
            dvs_list=[
                ("Enterprise Content & Reg", DVS5["Enterprise Content & Reg"], "primary"),
            ],
            apps_text="Rate Case Advisor (drafting)  ·  OpenText (documents)  ·  Utilisphere (utility data)  ·  internal financial systems",
            other_ovs_text="Source-to-Pay  ·  Hire-to-Retire  ·  Regulatory Reporting Cycle  ·  Employee Self-Service",
        ),
        dict(  # OVS — Procure-to-Pay (Spain)
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
            dvs_list=[
                ("SAP Enterprise Platform", DVS5["SAP Enterprise Platform"], "primary"),
            ],
            apps_text="SAP PO (purchase)  ·  SAP FI (finance)  ·  SAP Budget (control)  ·  S/4HANA (platform)",
            other_ovs_text="Period Close (Spain)  ·  Plan-to-Budget (Spain)  ·  Financial Reporting (Spain)",
        ),
    ]

    for i, spec in enumerate(ovs_specs):
        s = prs.slides[10 + i]
        clear_slide(s)
        copy_bg(prs, 10 + i, s)
        build_ovs_detail_slide(s, **spec)

    # Step 10: Clone Enabling Team (slide 14) 5 times for SAFe DVS detail
    # First, ensure Enabling Team currently at idx 14 has the DVS detail layout (~28 shapes)
    for i in range(5):
        new = clone_slide(prs, 14)  # always clone from idx 14 (Enabling)
        edit_by_idx(new, SAFE_DVS_DETAIL[i])
    # Now: 0-14 same, 15-19 = 5 SAFe DVS detail clones
    # Move Enabling (idx 14) to position 19 (after the 5 new DVS detail slides)
    move_slide(prs, 14, 19)
    # Final: 0-13 same, 14-18 = 5 SAFe DVS detail (was at 15-19), 19 = Enabling

    # Step 11: Edit slide 19 (Enabling Team) — final modification
    edit_by_idx(prs.slides[19], {
        0:  "ENABLING TEAM",
        1:  "FOUNDATION SERVICES — SHARED CAPABILITIES THAT KEEP THE ITBA DVS RUNNING",
        3:  "ITBA · governed by Director ITBA (Carlos Garbin)",
        5:  "No PM · no DVS-level backlog · operates as a small shared-service team",
        8:  ("Provide the technical and process foundation — Software Factory governance, demand "
             "intake, test & release coordination, third-party risk gating — that the ITBA DVS rely on."),
        11: "FOUNDATION SERVICES",
        12: ("•  SWF (Software Factory) governance\n"
             "•  Demand intake orchestration (PDD · Jira · IT Now)\n"
             "•  Test & release coordination (cross-DVS)\n"
             "•  Third-party risk assessment gating\n"
             "•  Vendor-management cadence (Accenture · SWF · business vendors)\n"
             "•  ALM tooling administration"),
        15: "WHO USES THESE SERVICES",
        16: ("•  All ITBA DVS — for SWF capacity · environments · test cycles\n"
             "•  Engineering DVS — third-party risk gating for vendor upgrades\n"
             "•  SAP Platform DVS — transport coordination · integration\n"
             "•  Director ITBA — portfolio view across DVS\n"
             "•  Optionally shared with CS Apps enabling team (Digital Products consolidation)"),
        17: "Foundation services support all DVS — they do not constitute a separate product line",
        20: "TEAM SHAPE",
        21: ("•  Coordination led by Director ITBA (Carlos Garbin)\n"
             "•  SWF governance lead\n"
             "•  Demand intake / portfolio admin\n"
             "•  Test & risk coordinators\n"
             "•  No dedicated Product Manager"),
        22: "Why not a DVS: at this headcount, productizing foundation services would create overhead",
        25: "VENDOR EXTENSION",
        26: ("•  SWF (Software Factory) vendor\n"
             "•  Accenture — for cross-DVS major projects\n"
             "•  Atlassian Jira admins\n"
             "•  Third-party risk assessment vendors"),
        27: "Avangrid governs operationally · vendors execute specific capabilities",
    })

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
