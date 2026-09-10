"""
v2 → v3 refactor:
  1. Strip [KB ...] / [INFERENCE] / [estimate] tokens
  2. Replace fabricated R/G/T mixes & FTE estimates with "[to validate]"
  3. Rename DVS to outcome/product-shaped names
  4. Recast slide 11 (Platform) as a proper DVS (own PM, own backlog, customer = other DVS)
  5. Rebuild slide 2 layout (OVS card height + new DVS names)
  6. Add Big Picture slide (M2C operational flow with chevrons + solutions + DVS bands)
  7. Reorder: cover → big picture → OVS+5DVS cards → why → framing → intake → OVS detail → DVS 1..5 → value lens
"""
import re
from copy import deepcopy
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

SRC = "Material em andamento/Customer Services - Proposed Value Streams - v2.pptx"
DST = "Material em andamento/Customer Services - Proposed Value Streams - v3.pptx"

# Brand
NAVY  = RGBColor(0x0F, 0x1B, 0x4F)
BLUE2 = RGBColor(0x2C, 0x57, 0x8B)
BLUE3 = RGBColor(0x5B, 0x7F, 0xB5)
BLUE4 = RGBColor(0xA5, 0xC2, 0xDA)
GREEN = RGBColor(0x9C, 0xC4, 0x4A)  # SAFe-style chevron green
TEAL  = RGBColor(0x3A, 0x6F, 0x7E)  # SAFe-style solution box
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK  = RGBColor(0x2C, 0x3E, 0x50)
GRAY  = RGBColor(0x95, 0xA5, 0xA6)
LIGHT = RGBColor(0xF5, 0xF5, 0xF5)

# DVS rename map (old → new), applied to every text frame
DVS_RENAMES = [
    ("Meter Data & Devices", "Meter & Device Operations"),
    ("METER DATA & DEVICES", "METER & DEVICE OPERATIONS"),
    ("Billing & Invoicing", "Billing Operations"),
    ("BILLING & INVOICING", "BILLING OPERATIONS"),
    ("Collections · FICA · Tax", "Revenue & Collections"),
    ("COLLECTIONS · FICA · TAX", "REVENUE & COLLECTIONS"),
    ("Collections, FICA & Tax", "Revenue & Collections"),
    ("Customer Interaction", "Customer Care & Engagement"),
    ("CUSTOMER INTERACTION", "CUSTOMER CARE & ENGAGEMENT"),
    ("Platform & Enablers", "Foundation Platform"),
    ("PLATFORM & ENABLERS", "FOUNDATION PLATFORM"),
]


# ============== text cleanup helpers ==============

def clean_run_text(text):
    """Strip citations, fabricated qualifiers, fabricated mixes, rename DVS."""
    new = text
    # Strip "[KB ...]" with various spacings
    new = re.sub(r"\s*\[KB[^\]]*\]\s*", " ", new)
    # Strip "[INFERENCE]" / "[INFERENCE ...]"
    new = re.sub(r"\s*\[INFERENCE[^\]]*\]\s*", " ", new)
    # Strip "[estimate]" / "[estimate ...]"
    new = re.sub(r"\s*\[estimate[^\]]*\]\s*", " ", new)
    # Strip "[GAP]" prefix
    new = re.sub(r"\s*\[GAP\]\s*", "", new)
    # Strip "[proposed]"
    new = re.sub(r"\s*\[proposed\]\s*", " ", new)
    # Replace Run/Grow/Transform fabricated mixes with placeholder
    new = re.sub(
        r"Run · Grow · Transform\s*·\s*\d+\s*/\s*\d+\s*/\s*\d+(\s*·\s*shared svc)?",
        "Run · Grow · Transform mix · to validate with team",
        new,
    )
    # Replace fabricated FTE estimates "≈ X–Y FTE" / "≈X-Y FTE"
    new = re.sub(r"≈\s*\d+\s*[–-]\s*\d+\s*FTE\s*\(highest run-load\)", "Squad size · to validate", new)
    new = re.sub(r"≈\s*\d+\s*[–-]\s*\d+\s*FTE\b", "Squad size · to validate", new)
    new = re.sub(r"\(≈\s*\d+\s*[–-]\s*\d+\s*FTE\)", "(size to validate)", new)
    new = re.sub(r"~\s*\d+\s*[–-]\s*\d+\s*FTE per DVS", "Squad size to validate", new)
    # Apply DVS renames
    for old, repl in DVS_RENAMES:
        new = new.replace(old, repl)
    # Collapse multiple spaces
    new = re.sub(r" {2,}", " ", new)
    # Trim trailing punctuation/dashes
    new = new.rstrip()
    return new


def walk_replace(prs):
    n = 0
    for slide in prs.slides:
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    new = clean_run_text(r.text)
                    if new != r.text:
                        r.text = new
                        n += 1
    return n


# ============== shape helpers (re-used) ==============

def add_text(slide, x, y, w, h, text, *, size=10, bold=False, italic=False,
             color=DARK, name="Arial", align=PP_ALIGN.LEFT, anchor=None, wrap=True):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.margin_left = Inches(0.04); tf.margin_right = Inches(0.04)
    tf.margin_top = Inches(0.02); tf.margin_bottom = Inches(0.02)
    if anchor: tf.vertical_anchor = anchor
    for i, line in enumerate(text.split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run(); r.text = line
        r.font.name = name
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color
    return tb


def add_rect(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None):
    sh = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(0.5)
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        sh.adjustments[0] = radius
    return sh


def copy_bg(prs, src_idx, new_slide):
    src = prs.slides[src_idx]
    src_bg = src._element.find(qn('p:cSld') + '/' + qn('p:bg'))
    new_cSld = new_slide._element.find(qn('p:cSld'))
    if src_bg is not None and new_cSld is not None:
        existing = new_cSld.find(qn('p:bg'))
        if existing is not None:
            new_cSld.remove(existing)
        new_cSld.insert(0, deepcopy(src_bg))


def move_slide(prs, old_idx, new_idx):
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    xml_slides.remove(slides[old_idx])
    xml_slides.insert(new_idx, slides[old_idx])


# ============== rebuild slide 2 (1 OVS + 5 DVS cards) ==============

def add_dvs_card_v2(slide, x, y, w, h, title, domain, desc, apps_label, apps,
                    *, font_title=10.5, font_apps=8.0):
    body = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  Inches(x), Inches(y + 0.08), Inches(w), Inches(h - 0.08))
    body.fill.solid(); body.fill.fore_color.rgb = WHITE
    body.line.color.rgb = NAVY; body.line.width = Pt(0.5)
    body.adjustments[0] = 0.05

    hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 Inches(x - 0.02), Inches(y), Inches(w + 0.04), Inches(0.78))
    hdr.fill.solid(); hdr.fill.fore_color.rgb = NAVY
    hdr.line.fill.background()
    hdr.adjustments[0] = 0.10

    tb = slide.shapes.add_textbox(Inches(x), Inches(y + 0.10), Inches(w), Inches(0.62))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = Inches(0.05); tf.margin_right = Inches(0.05)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p1 = tf.paragraphs[0]; p1.alignment = PP_ALIGN.CENTER
    r = p1.add_run(); r.text = title
    r.font.name = "Arial"; r.font.size = Pt(font_title); r.font.bold = True; r.font.color.rgb = WHITE

    inner_y = y + 0.92
    add_text(slide, x + 0.10, inner_y, w - 0.20, 0.20, domain,
             size=7.5, bold=True, italic=True, color=DARK, align=PP_ALIGN.CENTER)
    add_text(slide, x + 0.10, inner_y + 0.22, w - 0.20, 0.55, desc,
             size=9.5, italic=True, color=DARK, align=PP_ALIGN.CENTER)

    ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                Inches(x + 0.20), Inches(inner_y + 0.85),
                                Inches(w - 0.40), Inches(0.02))
    ln.fill.solid(); ln.fill.fore_color.rgb = NAVY
    ln.line.fill.background()

    add_text(slide, x + 0.10, inner_y + 0.93, w - 0.20, 0.20, apps_label,
             size=7.5, bold=True, color=DARK, align=PP_ALIGN.CENTER)

    apps_y = inner_y + 1.15
    add_text(slide, x + 0.10, apps_y, w - 0.20, y + h - apps_y - 0.10,
             apps, size=font_apps, color=DARK, align=PP_ALIGN.CENTER)


def rebuild_slide_2(prs):
    slide = prs.slides[1]
    # remove all card shapes (idx >= 2 — keep header bar)
    for sh in list(slide.shapes)[2:]:
        sp = sh._element
        sp.getparent().remove(sp)

    # OVS card on top (taller now to fit content)
    ovs_x, ovs_y, ovs_w, ovs_h = 0.50, 1.05, 12.33, 1.95

    body = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  Inches(ovs_x), Inches(ovs_y + 0.05),
                                  Inches(ovs_w), Inches(ovs_h - 0.05))
    body.fill.solid(); body.fill.fore_color.rgb = WHITE
    body.line.color.rgb = NAVY; body.line.width = Pt(0.5)
    body.adjustments[0] = 0.04

    hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 Inches(ovs_x - 0.02), Inches(ovs_y),
                                 Inches(ovs_w + 0.04), Inches(0.50))
    hdr.fill.solid(); hdr.fill.fore_color.rgb = NAVY
    hdr.line.fill.background()
    hdr.adjustments[0] = 0.06

    add_text(slide, ovs_x + 0.20, ovs_y + 0.08, ovs_w - 3.10, 0.40,
             "Meter-to-Cash", size=13, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, ovs_x + ovs_w - 3.10, ovs_y + 0.10, 2.90, 0.36,
             "OPERATIONAL VALUE STREAM (1)", size=9, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)

    inner_y = ovs_y + 0.62
    add_text(slide, ovs_x + 0.20, inner_y, ovs_w - 0.40, 0.20,
             "DOMAIN · Customer Service · End-to-End Business Outcome across 7 OPCOs",
             size=8.5, bold=True, italic=True, color=DARK)
    add_text(slide, ovs_x + 0.20, inner_y + 0.22, ovs_w - 0.40, 0.34,
             "Capture every meter read · invoice on time · collect cash · serve every customer "
             "interaction — anchored on SAP (Billing · FICA · CRM · Metering)",
             size=11, italic=True, color=DARK)

    ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                Inches(ovs_x + 0.20), Inches(inner_y + 0.62),
                                Inches(ovs_w - 0.40), Inches(0.02))
    ln.fill.solid(); ln.fill.fore_color.rgb = NAVY
    ln.line.fill.background()

    add_text(slide, ovs_x + 0.20, inner_y + 0.68, ovs_w - 0.40, 0.20,
             "STAGES", size=8.5, bold=True, color=DARK)
    add_text(slide, ovs_x + 0.20, inner_y + 0.90, ovs_w - 0.40, 0.40,
             "METER  →  BILL  →  COLLECT  →  SERVE   ·   FOUNDATION PLATFORM (transversal)",
             size=11, bold=True, color=NAVY)

    # Label row
    add_text(slide, 0.50, 3.13, 12.33, 0.22,
             "DEVELOPMENT VALUE STREAMS (5)  ·  ONE PRODUCT MANAGER PER DVS",
             size=10, bold=True, color=NAVY)

    # 5 DVS cards (renamed)
    dvs_y = 3.40
    dvs_h = 3.55
    n = 5
    margin = 0.50
    gap = 0.10
    total_w = 13.33 - 2 * margin
    card_w = (total_w - (n - 1) * gap) / n

    dvs_cards = [
        ("Meter & Device Operations",
         "DOMAIN · Metering",
         "Capture every read · zero data loss to billing",
         "Smart Metering\nMDM (Meter Data Mgmt)\nSAP IS-U Device\nGroup 8760 (EDI)\nField meter read interfaces"),
        ("Billing Operations",
         "DOMAIN · Billing",
         "Run the bill on time · every day · every OPCO",
         "SAP ECC (FI/SD)\nStreamserve / OpenText\nAddress Validation (Precisely)\nPrint & e-delivery interfaces"),
        ("Revenue & Collections",
         "DOMAIN · Finance",
         "Convert invoices into cash · comply at line level",
         "SAP FICA\nSOVOS Taxware\nPayment processing interfaces\n(banks · lockbox · 3rd-party)\nDunning & collections workflows"),
        ("Customer Care & Engagement",
         "DOMAIN · CX",
         "Unify CRM · contact center · digital · feedback",
         "SAP CRM\nGenesys (Infosys)\nNuance IVR / CTI\nMedallia (formerly VHT)\nPureConnect / PureEngage (legacy)"),
        ("Foundation Platform",
         "DOMAIN · Platform-as-a-Product",
         "Productize platform services for the other 4 DVS",
         "SAP Basis\nSAP Gold Client\nSAP Enable Now (WPB)\nBlue Prism (RPA)\nTricentis TOSCA\nJira"),
    ]

    for i, (title, domain, desc, apps) in enumerate(dvs_cards):
        x = margin + i * (card_w + gap)
        add_dvs_card_v2(slide, x, dvs_y, card_w, dvs_h, title, domain, desc,
                        "ALL APPLICATIONS", apps)


# ============== rebuild slide 11 (Foundation Platform as proper DVS) ==============

def rebuild_platform_slide(prs):
    """Slide 11 — was 'Platform & Enablers (Shared Service)'. Recast as full DVS."""
    slide = prs.slides[10]  # 0-indexed slide 11

    by_idx = {i: s for i, s in enumerate(slide.shapes)}
    # Same shape layout as other DVS slides — re-use indices 0..27.
    def setrun(i, text):
        sh = by_idx[i]
        tf = sh.text_frame
        # capture first run formatting
        if tf.paragraphs and tf.paragraphs[0].runs:
            f0 = tf.paragraphs[0].runs[0].font
            sz = f0.size
            bold = f0.bold
            italic = f0.italic
            try: color = f0.color.rgb
            except: color = None
            name = f0.name
        else:
            sz, bold, italic, color, name = Pt(10), False, False, None, "Arial"
        tf.clear()
        for j, line in enumerate(text.split('\n')):
            p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
            r = p.add_run(); r.text = line
            r.font.name = name
            if sz: r.font.size = sz
            r.font.bold = bold
            r.font.italic = italic
            if color: r.font.color.rgb = color

    setrun(0,  "DEVELOPMENT VALUE STREAM (DVS) 5 / 5")
    setrun(1,  "FOUNDATION PLATFORM — PRODUCTIZE THE PLATFORM SO EVERY OTHER DVS CAN MOVE FAST")
    setrun(3,  "CS Apps · Platform DVS · customer = the other 4 DVS")
    setrun(5,  "Run · Grow · Transform mix · to validate with team")
    setrun(8,  "Provide a productized platform — landscape, data, automation, testing, ALM tooling — "
               "that lets every other DVS deliver value without rebuilding common foundations.")
    setrun(11, "APPLICATIONS SUPPORTED")
    setrun(12, "•  SAP Basis (landscape · transports · security · performance)\n"
               "•  SAP Gold Client (selective non-prod data copy · masking)\n"
               "•  SAP Enable Now (training & in-app guidance)\n"
               "•  Blue Prism (RPA for back-office CS processes)\n"
               "•  Tricentis TOSCA (test automation)\n"
               "•  Jira (demand · backlog · agile work mgmt)")
    setrun(15, "PLATFORM CONSUMERS — who pulls value from this DVS")
    setrun(16, "•  Meter & Device Operations DVS — environments, EDI test data\n"
               "•  Billing Operations DVS — SAP transports, regression coverage\n"
               "•  Revenue & Collections DVS — FICA refresh, payment integration sandboxes\n"
               "•  Customer Care & Engagement DVS — CRM masked data, IVR test suites\n"
               "•  Release & change governance — readiness gates, audit trail")
    setrun(17, "Foundation Platform = product · the other 4 DVS = its customers")
    setrun(20, "AVANGRID CORE TEAM")
    setrun(21, "•  Product Manager — Platform PM (NEW · same level as other DVS PMs)\n"
               "•  Tech Lead — Naggi Reddy (Sr. Manager SAP Basis)\n"
               "•  Architecture — Sr. Manager SAP Architecture (vacancy → fill)\n"
               "•  RPA / Test / ALM Engineers (in-team, not matrixed out)\n"
               "Squad size · to validate")
    setrun(22, "Treated as DVS, not enabling team — so platform decisions have a clear owner and a roadmap")
    setrun(25, "VENDOR EXTENSION (delivery)")
    setrun(26, "•  Tricentis (test platform)\n"
               "•  Blue Prism vendor + Automation CoE\n"
               "•  Atlassian Jira admins\n"
               "•  SAP partner for Basis / Gold Client")
    setrun(27, "Avangrid = product owner · vendor = delivery muscle")


# ============== big picture slide ==============

def chevron(slide, x, y, w, h, text, fill=GREEN, txt=DARK, font_size=10):
    sh = slide.shapes.add_shape(MSO_SHAPE.PENTAGON, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    sh.line.color.rgb = NAVY; sh.line.width = Pt(0.5)
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = Inches(0.05); tf.margin_right = Inches(0.18)
    tf.margin_top = Inches(0.02); tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text
    r.font.name = "Arial"; r.font.size = Pt(font_size); r.font.bold = True; r.font.color.rgb = txt


def solution_box(slide, x, y, w, h, text, fill=TEAL, txt=WHITE, font_size=8.5):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = Inches(0.05); tf.margin_right = Inches(0.05)
    tf.margin_top = Inches(0.04); tf.margin_bottom = Inches(0.04)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text
    r.font.name = "Arial"; r.font.size = Pt(font_size); r.font.bold = True; r.font.color.rgb = txt


def add_big_picture(prs):
    src = prs.slides[5]  # any "Em Branco" slide for layout
    new = prs.slides.add_slide(src.slide_layout)
    for sh in list(new.shapes):
        sh._element.getparent().remove(sh._element)
    copy_bg(prs, 5, new)
    slide = new

    # Header
    add_text(slide, 0.40, 0.10, 12.53, 0.38,
             "OPERATIONAL VALUE STREAM — METER-TO-CASH BIG PICTURE",
             size=20, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, 0.40, 0.58, 12.53, 0.23,
             "End-to-end customer flow  ·  solutions that enable each step  ·  "
             "Development Value Streams that build & evolve the solutions",
             size=11, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # Trigger (left) and Outcome (right) labels
    add_text(slide, 0.40, 1.10, 1.30, 0.40,
             "TRIGGER", size=8, bold=True, color=GRAY, align=PP_ALIGN.CENTER)
    add_text(slide, 0.40, 1.48, 1.30, 0.50,
             "Customer needs energy  /  signs contract",
             size=9.5, bold=True, italic=True, color=DARK, align=PP_ALIGN.CENTER)

    add_text(slide, 11.63, 1.10, 1.30, 0.40,
             "OUTCOME", size=8, bold=True, color=GRAY, align=PP_ALIGN.CENTER)
    add_text(slide, 11.63, 1.48, 1.30, 0.50,
             "Bill paid  ·  customer engaged  ·  regulator compliant",
             size=9.5, bold=True, italic=True, color=DARK, align=PP_ALIGN.CENTER)

    # Chevrons row (8 steps)
    steps = [
        "Sign up\n& activate",
        "Acquire\nmeter reads",
        "Validate &\nprocess data",
        "Calculate\n& issue bill",
        "Deliver\ninvoice",
        "Process\npayment",
        "Handle\ninteractions",
        "Reconcile\n& close",
    ]
    chev_y = 1.40
    chev_h = 0.85
    chev_x_start = 1.85
    chev_w = 1.30
    chev_overlap = 0.10
    for i, label in enumerate(steps):
        cx = chev_x_start + i * (chev_w - chev_overlap)
        chevron(slide, cx, chev_y, chev_w, chev_h, label, fill=GREEN, txt=DARK, font_size=8.5)

    # Personas row (above each chevron)
    personas = [
        "Customer\n+ CSR",
        "Smart meter\n+ field",
        "Meter ops\nteam",
        "Billing\nanalyst",
        "Customer\n(channel)",
        "Customer\n+ bank",
        "CSR + customer\n+ digital",
        "Finance\n+ audit",
    ]
    for i, p in enumerate(personas):
        cx = chev_x_start + i * (chev_w - chev_overlap)
        add_text(slide, cx, 0.95, chev_w - chev_overlap, 0.40, p,
                 size=7.0, italic=True, color=GRAY, align=PP_ALIGN.CENTER)

    # Solution boxes below — grouped (some span multiple chevrons)
    sol_y = 2.65
    sol_h = 0.55
    solutions = [
        # (start_step, end_step, text)
        (0, 0, "SAP CRM\nonboarding"),
        (1, 2, "Smart Metering · MDM ·\nGroup 8760 (EDI) · IS-U Device"),
        (3, 3, "SAP ECC (FI/SD)\nbilling engine"),
        (4, 4, "Streamserve /\nOpenText"),
        (5, 5, "SAP FICA ·\nbank integrations"),
        (6, 6, "SAP CRM · Genesys ·\nNuance IVR · Medallia"),
        (7, 7, "SAP FICA ·\nSOVOS Taxware"),
    ]
    for s, e, label in solutions:
        x_left = chev_x_start + s * (chev_w - chev_overlap) + 0.05
        x_right = chev_x_start + e * (chev_w - chev_overlap) + chev_w - 0.05
        sw = x_right - x_left
        solution_box(slide, x_left, sol_y, sw, sol_h, label, fill=TEAL, txt=WHITE, font_size=8.0)

    # DVS bands at bottom (showing ownership)
    band_y = 3.50
    band_h = 0.45
    add_text(slide, 0.40, band_y - 0.30, 12.53, 0.22,
             "DEVELOPMENT VALUE STREAMS — who builds & evolves each solution",
             size=10, bold=True, color=NAVY)

    # Define DVS coverage by chevron range
    # color per DVS
    DVS_COLORS = {
        "Customer Care & Engagement": RGBColor(0x8E, 0x44, 0xAD),
        "Meter & Device Operations":  RGBColor(0x16, 0xA0, 0x85),
        "Billing Operations":         RGBColor(0xD3, 0x54, 0x00),
        "Revenue & Collections":      RGBColor(0x2C, 0x57, 0x8B),
        "Foundation Platform":        RGBColor(0x7F, 0x8C, 0x8D),
    }
    bands = [
        # (chev_start, chev_end, dvs_name, row)
        (0, 0, "Customer Care & Engagement", 0),
        (1, 2, "Meter & Device Operations", 0),
        (3, 4, "Billing Operations", 0),
        (5, 5, "Revenue & Collections", 0),
        (6, 6, "Customer Care & Engagement", 0),
        (7, 7, "Revenue & Collections", 0),
    ]
    for s, e, name, row in bands:
        x_left = chev_x_start + s * (chev_w - chev_overlap) + 0.05
        x_right = chev_x_start + e * (chev_w - chev_overlap) + chev_w - 0.05
        sw = x_right - x_left
        by = band_y + row * (band_h + 0.08)
        sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                    Inches(x_left), Inches(by), Inches(sw), Inches(band_h))
        sh.fill.solid(); sh.fill.fore_color.rgb = DVS_COLORS[name]
        sh.line.fill.background()
        sh.adjustments[0] = 0.30
        add_text(slide, x_left, by, sw, band_h, name,
                 size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Foundation Platform transversal band (full width)
    fp_y = band_y + band_h + 0.20
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                Inches(chev_x_start + 0.05),
                                Inches(fp_y),
                                Inches((chev_w - chev_overlap) * 8 - 0.10),
                                Inches(band_h))
    sh.fill.solid(); sh.fill.fore_color.rgb = DVS_COLORS["Foundation Platform"]
    sh.line.fill.background()
    sh.adjustments[0] = 0.30
    add_text(slide, chev_x_start + 0.05, fp_y,
             (chev_w - chev_overlap) * 8 - 0.10, band_h,
             "Foundation Platform  ·  productized services consumed by every other DVS  "
             "(test data · environments · automation · ALM)",
             size=8.5, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Legend / explainer
    legend_y = fp_y + band_h + 0.40
    add_text(slide, 0.40, legend_y, 12.53, 0.22,
             "HOW TO READ THIS DIAGRAM",
             size=9.5, bold=True, color=NAVY)
    add_text(slide, 0.40, legend_y + 0.25, 12.53, 1.20,
             "•  Top row — the customer-facing flow (the operational outcome customers experience)\n"
             "•  Green chevrons — the 8 value-adding steps that take the customer from \"need energy\" "
             "to \"bill paid\"\n"
             "•  Teal boxes — the solutions / applications that enable each step\n"
             "•  Coloured strips — the Development Value Streams that own & evolve each solution; "
             "Foundation Platform supports all 4 throughout\n"
             "•  This is one operational value chain (Meter-to-Cash) delivered by 5 product-shaped "
             "Development Value Streams",
             size=9.5, color=DARK)

    # Source attribution
    add_text(slide, 0.40, 7.05, 12.53, 0.22,
             "Pattern: standard operational value-stream visualization "
             "(customer trigger → process steps → enabling solutions → development teams → outcome)",
             size=7.5, italic=True, color=GRAY)

    return new


# ============== main ==============

def main():
    prs = Presentation(SRC)

    # 1. Text cleanup pass (removes citations + fab data + DVS rename)
    n = walk_replace(prs)
    print(f"Text cleanup: {n} runs updated")

    # 2. Rebuild slide 2 layout with new DVS names
    rebuild_slide_2(prs)
    print("Slide 2 rebuilt with new DVS names + fixed OVS card height")

    # 3. Recast slide 11 as Foundation Platform DVS
    rebuild_platform_slide(prs)
    print("Slide 11 (Foundation Platform) recast as proper DVS")

    # 4. Add big picture slide → appended at end (idx 12)
    add_big_picture(prs)
    print("Big picture slide added")

    # 5. Reorder: move new big picture (idx 12) to position 1 (after cover)
    move_slide(prs, len(prs.slides) - 1, 1)
    # New order should be:
    #  0 cover, 1 big picture, 2 OVS+5DVS cards, 3 why, 4 framing, 5 intake,
    #  6 OVS detail, 7..11 DVS 1..5, 12 value lens

    prs.save(DST)
    print(f"\nSaved → {DST}")
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
