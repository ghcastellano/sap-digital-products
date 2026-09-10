"""
Standalone editable slide: TO-BE DEMAND FLOW with 4 DVS + 5 ITBA + 1 Enabling Team.
Reformatted for visual clarity — chevrons for M2C stages, badges for ITBA streams,
all 10 team boxes neatly spaced.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

OUT = "/Users/gustavohenriquecastellano/Downloads/Base Apresentações Projeto Digital Product/TO-BE Demand Flow - SAFe Proposal 2.pptx"

# Palette
NTT_BLUE = RGBColor(0x3B, 0x6E, 0x8F)
NAVY_DARK = RGBColor(0x0F, 0x2A, 0x47)
NAVY = RGBColor(0x17, 0x34, 0x58)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
GREY_BG = RGBColor(0xF7, 0xF9, 0xFC)
GREY_LIGHT = RGBColor(0x9F, 0xB3, 0xD1)
GREY_MED = RGBColor(0x6B, 0x77, 0x8C)
GREY_DARK = RGBColor(0x42, 0x52, 0x6E)
ORANGE = RGBColor(0xF0, 0x8A, 0x3E)
ORANGE_LIGHT = RGBColor(0xFF, 0xF6, 0xE6)
SOURCE_BLUE = RGBColor(0x3B, 0x6E, 0x8F)
SOURCE_GREY = RGBColor(0x80, 0x80, 0x80)
RED = RGBColor(0xDE, 0x35, 0x0B)
YELLOW = RGBColor(0xFF, 0xC1, 0x07)
# DVS amber tones
DVS_AMBER = RGBColor(0xE8, 0xA2, 0x4C)
DVS_AMBER_LIGHT = RGBColor(0xFD, 0xEB, 0xCA)
DVS_AMBER_DARK = RGBColor(0x5A, 0x3A, 0x0E)
# ITBA green tones
ITBA_GREEN = RGBColor(0x5D, 0xA4, 0x60)
ITBA_GREEN_LIGHT = RGBColor(0xE2, 0xF0, 0xE2)
ITBA_GREEN_DARK = RGBColor(0x1F, 0x57, 0x22)
# Enabling Team purple
ENABLE_PURPLE = RGBColor(0x5C, 0x2A, 0x91)
ENABLE_PURPLE_LIGHT = RGBColor(0xEE, 0xE5, 0xF8)


prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])


# ---------- helpers ----------
def rect(slide, x, y, w, h, fill=None, line=None, line_w=None):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    s.shadow.inherit = False
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        if line_w is not None: s.line.width = line_w
    return s


def round_rect(slide, x, y, w, h, fill=None, line=None, line_w=None, radius=0.06):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    s.shadow.inherit = False
    s.adjustments[0] = radius
    if fill is None:
        s.fill.background()
    else:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        if line_w is not None: s.line.width = line_w
    return s


def chevron(slide, x, y, w, h, fill, line=None):
    s = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, x, y, w, h)
    s.shadow.inherit = False
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
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


# ============== HEADER ==============
rect(slide, 0, 0, prs.slide_width, Inches(0.95), NTT_BLUE)
text(slide, Inches(0.4), Inches(0.1), Inches(11), Inches(0.5),
     "TO-BE DEMAND FLOW", 26, bold=True, color=WHITE)
text(slide, Inches(0.4), Inches(0.55), Inches(11), Inches(0.35),
     "ONE JIRA FRONT DOOR ROUTES SOURCE AREAS INTO ITBA STREAMS + 4 CS DVS + 1 ENABLING TEAM",
     11, bold=True, color=WHITE)
text(slide, Inches(11.3), Inches(0.3), Inches(1.7), Inches(0.4),
     "NTT  DATA", 16, bold=True, color=WHITE, align=PP_ALIGN.RIGHT)

# Subtitle line
text(slide, Inches(0.4), Inches(1.0), Inches(13), Inches(0.3),
     "Single intake form per Value Stream — every requester opens a Jira card with mandatory fields. The form auto-routes to the right Value Stream Team and applies a Class of Service.",
     10, italic=True, bold=True, color=NAVY)

# ============== COLUMN GEOMETRY ==============
COL1_X = Inches(0.3); COL1_W = Inches(2.8)
COL2_X = Inches(3.2); COL2_W = Inches(2.8)
COL3_X = Inches(6.1); COL3_W = Inches(3.7)
COL4_X = Inches(9.9); COL4_W = Inches(3.15)
HEAD_Y = Inches(1.45)
HEAD_H = Inches(0.55)

cols = [
    (COL1_X, COL1_W, "01 · SOURCES", "Who sends demand"),
    (COL2_X, COL2_W, "02 · INTAKE FORM (JIRA)", "Mandatory fields capture intent"),
    (COL3_X, COL3_W, "03 · MAIN VALUE STREAM", "End-to-end customer outcome"),
    (COL4_X, COL4_W, "04 · VALUE STREAM TEAM", "4 DVS + 5 ITBA + 1 Enabling"),
]
for x, w, h_t, h_s in cols:
    rect(slide, x, HEAD_Y, w, HEAD_H, NAVY_DARK)
    text(slide, x, HEAD_Y + Inches(0.05), w, Inches(0.28),
         h_t, 12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    text(slide, x, HEAD_Y + Inches(0.32), w, Inches(0.22),
         h_s, 8, italic=True, color=YELLOW, align=PP_ALIGN.CENTER)


BODY_Y = Inches(2.10)
BODY_H = Inches(4.95)


# ============== COLUMN 1 — SOURCES ==============
sources = [
    ("BRS (UIL · CMP · NY)", SOURCE_BLUE),
    ("CX & Digital CoE", SOURCE_BLUE),
    ("Data & Analytics", SOURCE_BLUE),
    ("Infrastructure", SOURCE_BLUE),
    ("IA (Industrial Auto.)", SOURCE_BLUE),
    ("OT (Operational Tech.)", SOURCE_BLUE),
    ("Procurement", SOURCE_BLUE),
    ("Regulatory", SOURCE_BLUE),
    ("Corporate Areas", SOURCE_GREY),
    ("P&T (GIS · Global SAP)", SOURCE_GREY),
    ("Gas & Electric Ops", SOURCE_GREY),
]
n_src = len(sources)
src_gap = Inches(0.04)
src_h = (BODY_H - (n_src - 1) * src_gap) / n_src
sy = BODY_Y
for label, fill in sources:
    rect(slide, COL1_X, sy, COL1_W, src_h, fill)
    text(slide, COL1_X, sy, COL1_W, src_h, label, 10, bold=True, color=WHITE,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    sy += src_h + src_gap


# ============== COLUMN 2 — JIRA INTAKE FORM ==============
round_rect(slide, COL2_X, BODY_Y, COL2_W, BODY_H, WHITE, ORANGE, Pt(2.5))
# Form header strip
round_rect(slide, COL2_X, BODY_Y, COL2_W, Inches(0.4), ORANGE)
text(slide, COL2_X, BODY_Y, COL2_W, Inches(0.4),
     "JIRA REQUEST FORM", 13, bold=True, color=WHITE,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

fields = [
    ("▸ Value Stream destination", "(routes the card)"),
    ("▸ Demand type", "Run · Grow · Transform"),
    ("▸ T-shirt size", "S · M · L · XL"),
    ("▸ Class of Service", "Standard · Fixed-date · Expedite · Intangible"),
    ("▸ OPCO", "UIL · CMP · NY · Iberdrola · Cross"),
    ("▸ Business value statement", "Who benefits and how"),
    ("▸ Sponsor", "(approves at intake)"),
    ("▸ Pain Score inputs", "BV · UX · Safety · Effort"),
]
fy = BODY_Y + Inches(0.55)
field_h = Inches(0.42)
for label, sub in fields:
    text(slide, COL2_X + Inches(0.15), fy, COL2_W - Inches(0.3), Inches(0.22),
         label, 10, bold=True, color=NAVY)
    text(slide, COL2_X + Inches(0.30), fy + Inches(0.20), COL2_W - Inches(0.45), Inches(0.18),
         sub, 8, italic=True, color=GREY_DARK)
    fy += field_h
# Bottom warning
note_y = BODY_Y + BODY_H - Inches(0.55)
round_rect(slide, COL2_X + Inches(0.15), note_y, COL2_W - Inches(0.3), Inches(0.45),
           ORANGE_LIGHT, RED, Pt(1))
text(slide, COL2_X + Inches(0.15), note_y, COL2_W - Inches(0.3), Inches(0.45),
     "IT Now is reserved for incidents only — not for new demand.",
     9, bold=True, color=RED, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


# ============== COLUMN 3 — MAIN VALUE STREAM ==============
round_rect(slide, COL3_X, BODY_Y, COL3_W, BODY_H, WHITE, NAVY_DARK, Pt(2.5))

# --- CS section (top half) ---
cs_y = BODY_Y + Inches(0.1)
cs_h = Inches(2.55)

# CS title bar
round_rect(slide, COL3_X + Inches(0.1), cs_y, COL3_W - Inches(0.2), Inches(0.42), DVS_AMBER)
text(slide, COL3_X + Inches(0.1), cs_y, COL3_W - Inches(0.2), Inches(0.42),
     "CUSTOMER SERVICES · METER-TO-CASH", 13, bold=True, color=WHITE,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# CS subtitle
text(slide, COL3_X + Inches(0.15), cs_y + Inches(0.46), COL3_W - Inches(0.3), Inches(0.22),
     "1 OVS · 4 DVS · 1 Enabling Team · Owner: Kumar Jaladanki",
     9, italic=True, color=DVS_AMBER_DARK, align=PP_ALIGN.CENTER)

# 4 stage chevrons in a vertical flow (since column is narrow)
stages = [
    ("METER", "Capture every read", "→ DVS 1 Meter & Device Ops"),
    ("BILL", "Invoice on time · every OPCO", "→ DVS 2 Billing Operations"),
    ("COLLECT", "Convert invoices to cash", "→ DVS 3 Revenue & Collections"),
    ("SERVE", "Resolve every interaction", "→ DVS 4 Customer Care"),
]
stg_y = cs_y + Inches(0.78)
stg_h = Inches(0.36)
stg_gap = Inches(0.05)
for i, (stage, desc, dvs_link) in enumerate(stages):
    sy_local = stg_y + i * (stg_h + stg_gap)
    # Background tint
    round_rect(slide, COL3_X + Inches(0.15), sy_local, COL3_W - Inches(0.3), stg_h,
               DVS_AMBER_LIGHT, DVS_AMBER, Pt(0.75))
    # Stage label (left)
    text(slide, COL3_X + Inches(0.2), sy_local, Inches(0.85), stg_h,
         stage, 11, bold=True, color=DVS_AMBER_DARK, anchor=MSO_ANCHOR.MIDDLE)
    # Description (middle)
    text(slide, COL3_X + Inches(1.05), sy_local + Inches(0.04), Inches(1.6), stg_h - Inches(0.04),
         desc, 8, color=NAVY, anchor=MSO_ANCHOR.MIDDLE)
    # DVS link (right)
    text(slide, COL3_X + Inches(2.65), sy_local, COL3_W - Inches(2.8), stg_h,
         dvs_link, 8, bold=True, italic=True, color=DVS_AMBER_DARK, align=PP_ALIGN.RIGHT,
         anchor=MSO_ANCHOR.MIDDLE)

# Enabling Team support strip
en_y = cs_y + cs_h - Inches(0.3)
round_rect(slide, COL3_X + Inches(0.15), en_y, COL3_W - Inches(0.3), Inches(0.30),
           ENABLE_PURPLE_LIGHT, ENABLE_PURPLE, Pt(1))
text(slide, COL3_X + Inches(0.15), en_y, COL3_W - Inches(0.3), Inches(0.30),
     "▼ ENABLING TEAM supports all 4 stages (no PM)",
     9, bold=True, color=ENABLE_PURPLE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# --- Divider ---
divider_y = cs_y + cs_h + Inches(0.05)
rect(slide, COL3_X + Inches(0.4), divider_y, COL3_W - Inches(0.8), Emu(9525), GREY_LIGHT)

# --- ITBA section (bottom half) ---
itba_y = divider_y + Inches(0.08)
itba_h = BODY_H - cs_h - Inches(0.25)

# ITBA title bar
round_rect(slide, COL3_X + Inches(0.1), itba_y, COL3_W - Inches(0.2), Inches(0.42), ITBA_GREEN)
text(slide, COL3_X + Inches(0.1), itba_y, COL3_W - Inches(0.2), Inches(0.42),
     "IT BUSINESS APPLICATIONS", 13, bold=True, color=WHITE,
     align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

# ITBA subtitle
text(slide, COL3_X + Inches(0.15), itba_y + Inches(0.46), COL3_W - Inches(0.3), Inches(0.22),
     "5 streams · Director: Carlos Garbin (~26 FTE)",
     9, italic=True, color=ITBA_GREEN_DARK, align=PP_ALIGN.CENTER)

# 5 ITBA streams as a 2-column grid
itba_streams = [
    ("Field Operations", "GIS · Workforce · OSG"),
    ("Engineering", "Maintenance · Asset Mgmt"),
    ("Network Ops & Control", "SCADA · EMS · OMS"),
    ("Corporate & Regulatory", "FI · HR · Procurement"),
    ("Iberdrola Corp", "Global SAP · Corp Apps"),
]
itba_streams_y = itba_y + Inches(0.74)
sw = (COL3_W - Inches(0.45)) / 2  # 2 columns
sh = Inches(0.42)
sgap = Inches(0.06)
for i, (stream, sub) in enumerate(itba_streams):
    col = i % 2
    row = i // 2
    sx = COL3_X + Inches(0.15) + col * (sw + Inches(0.15))
    sy_loc = itba_streams_y + row * (sh + sgap)
    round_rect(slide, sx, sy_loc, sw, sh, ITBA_GREEN_LIGHT, ITBA_GREEN, Pt(0.75))
    text(slide, sx, sy_loc, sw, Inches(0.22),
         stream, 9, bold=True, color=ITBA_GREEN_DARK, align=PP_ALIGN.CENTER)
    text(slide, sx, sy_loc + Inches(0.20), sw, Inches(0.20),
         sub, 7, italic=True, color=GREY_DARK, align=PP_ALIGN.CENTER)


# ============== COLUMN 4 — VALUE STREAM TEAM (10 boxes) ==============
n_boxes = 10
box_gap = Inches(0.04)
box_h = (BODY_H - (n_boxes - 1) * box_gap) / n_boxes

teams = [
    ("DVS 1 — Meter & Device Operations", "Smart Metering · MDM · IS-U Device · 8760 EDI", DVS_AMBER, DVS_AMBER_DARK, False),
    ("DVS 2 — Billing Operations", "SAP ECC (FI/SD) · Streamserve · OpenText", DVS_AMBER, DVS_AMBER_DARK, False),
    ("DVS 3 — Revenue & Collections", "SAP FICA · SOVOS · Payment · Dunning", DVS_AMBER, DVS_AMBER_DARK, True),
    ("DVS 4 — Customer Care & Engagement", "SAP CRM · Genesys · Nuance IVR · Medallia", DVS_AMBER, DVS_AMBER_DARK, False),
    ("Field Operations", "GIS · Workforce · OSG", ITBA_GREEN, ITBA_GREEN_DARK, False),
    ("Engineering & Asset Mgmt", "Maintenance · Engineering", ITBA_GREEN, ITBA_GREEN_DARK, False),
    ("Network Operations & Control", "SCADA · EMS · OMS", ITBA_GREEN, ITBA_GREEN_DARK, False),
    ("Corporate & Regulatory", "FI · HR · Procurement · Reg.", ITBA_GREEN, ITBA_GREEN_DARK, False),
    ("Iberdrola Corp Services", "Global SAP · Corp. Apps", ITBA_GREEN, ITBA_GREEN_DARK, False),
    ("Enabling Team — Foundation Services", "Basis · Gold Client · Blue Prism · TOSCA · Enable Now (no PM)", ENABLE_PURPLE, ENABLE_PURPLE, True),
]

ty = BODY_Y
for label, sub, accent, accent_dark, is_new in teams:
    round_rect(slide, COL4_X, ty, COL4_W, box_h, WHITE, accent, Pt(2))
    # Left accent stripe
    rect(slide, COL4_X, ty, Inches(0.08), box_h, accent)
    # Main text
    text(slide, COL4_X + Inches(0.12), ty + Inches(0.02), COL4_W - Inches(0.55), Inches(0.22),
         label, 9, bold=True, color=accent_dark)
    # Sub text
    text(slide, COL4_X + Inches(0.12), ty + Inches(0.21), COL4_W - Inches(0.2), Inches(0.20),
         sub, 7, italic=True, color=GREY_DARK)
    # NEW sticker
    if is_new:
        nx = COL4_X + COL4_W - Inches(0.45)
        ny = ty + Inches(0.05)
        round_rect(slide, nx, ny, Inches(0.4), Inches(0.18), RED, radius=0.4)
        text(slide, nx, ny, Inches(0.4), Inches(0.18),
             "NEW", 8, bold=True, color=WHITE, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    ty += box_h + box_gap


# ============== FOOTER ==============
text(slide, Inches(0.3), Inches(7.16), Inches(8), Inches(0.18),
     "Avangrid · Digital Products · Target Operating Model",
     9, italic=True, color=GREY_MED)
text(slide, Inches(0.3), Inches(7.32), Inches(13), Inches(0.18),
     "© 2026 NTT DATA, Inc.  ·  CS Proposal 2: 1 OVS Meter-to-Cash + 4 DVS + 1 Enabling Team · ITBA streams unchanged",
     8, color=GREY_MED)


prs.save(OUT)
print(f"Saved: {OUT}")
print(f"Slide count: {len(prs.slides)}")
