"""
Digital Products — Unfix framework visualization (orthodox version).
Uses ONLY the 7 official Crew types and 10 official Forum types.
Reference: https://unfix.com/crew-types · /base-types · /forum-types
"""
from PIL import Image, ImageDraw, ImageFont
from pathlib import Path
import math

OUT = Path("/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Miro_Visuals")

# ── Palette (matched to user's Unfix template card colors) ───────────────
VS         = (245, 193,  58)   # Value Stream — yellow
VS_DARK    = (200, 153,  20)
GOV        = (251, 195, 145)   # Governance — peach
GOV_DARK   = (210, 145,  80)
CAP        = (216,  85,  81)   # Capability — red
CAP_DARK   = (160,  50,  46)
EXP        = (229, 143, 181)   # Experience — pink
EXP_DARK   = (180,  90, 130)
PLAT       = (111, 190, 111)   # Platform — green
PLAT_DARK  = ( 60, 140,  60)
PART       = (182, 213,  78)   # Partnership — lime
PART_DARK  = (140, 175,  40)
FAC        = (122, 102, 184)   # Facilitation — purple
FAC_DARK   = ( 80,  60, 145)

BASE_BG    = (231, 232, 234)
BASE_BR    = (140, 145, 155)
SIBLING_BG = (245, 245, 248)
SIBLING_BR = (180, 188, 200)
GUILD_BG   = (250, 250, 252)
GUILD_BR   = (170, 178, 190)

WHITE      = (255, 255, 255)
DARK       = ( 30,  40,  55)
GREY       = ( 90, 100, 115)
LIGHT_GREY = (140, 150, 165)
PERSON_F   = (200, 210, 220)
PERSON_BR  = (150, 160, 175)

W, H = 4600, 3200


def fnt(size, bold=False, italic=False):
    candidates = [
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
    ]
    for c in candidates:
        try:
            return ImageFont.truetype(c, size)
        except (OSError, IOError):
            continue
    return ImageFont.load_default()


def centered(d, xy, text, font, color):
    d.text(xy, text, fill=color, font=font, anchor="mm")


def left(d, xy, text, font, color):
    d.text(xy, text, fill=color, font=font, anchor="lm")


def person(d, cx, cy, r=22, fill=PERSON_F, br=PERSON_BR):
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=fill, outline=br, width=2)
    head_r = r * 0.38
    d.ellipse([cx - head_r, cy - r * 0.55, cx + head_r, cy - r * 0.55 + 2 * head_r],
              fill=br)
    d.pieslice([cx - r * 0.7, cy - r * 0.05, cx + r * 0.7, cy + r * 1.4],
               start=180, end=360, fill=br)


def home_icon(d, cx, cy, r=42, color=DARK):
    for ang in range(0, 360, 12):
        rad = math.radians(ang)
        x1 = cx + (r + 10) * math.cos(rad)
        y1 = cy + (r + 10) * math.sin(rad)
        x2 = cx + (r + 14) * math.cos(rad)
        y2 = cy + (r + 14) * math.sin(rad)
        d.line([x1, y1, x2, y2], fill=color, width=3)
    hw, hh = r * 1.2, r
    d.polygon([(cx - hw/2, cy - hh/4), (cx + hw/2, cy - hh/4),
               (cx + hw/2, cy + hh/2), (cx - hw/2, cy + hh/2)], fill=color)
    d.polygon([(cx - hw/2 - 6, cy - hh/4),
               (cx, cy - hh * 0.85),
               (cx + hw/2 + 6, cy - hh/4)], fill=color)
    dw, dh = hw * 0.25, hh * 0.45
    d.rectangle([cx - dw/2, cy + hh/2 - dh, cx + dw/2, cy + hh/2], fill=WHITE)


img = Image.new("RGB", (W, H), WHITE)
d = ImageDraw.Draw(img)


# ── Header ───────────────────────────────────────────────────────────────
d.rectangle([0, 0, W, 140], fill=DARK)
d.text((80, 30), "DIGITAL PRODUCTS — unFIX MODEL", font=fnt(60, True), fill=WHITE)
d.text((80, 95),
       "1 Strongly Aligned Base · Governance + 6 Value Stream + Capability + Experience + Platform + Partnership + Facilitation Crews",
       font=fnt(20), fill=(199, 211, 228))


# ── Strongly Aligned Base (Digital Products) ─────────────────────────────
sab_x0, sab_y0 = 160, 260
sab_x1, sab_y1 = 3300, 2700
d.rounded_rectangle([sab_x0, sab_y0, sab_x1, sab_y1], radius=40,
                    fill=BASE_BG, outline=BASE_BR, width=4)
home_icon(d, (sab_x0 + sab_x1) / 2, sab_y0 - 40, r=44)

# ── Governance Crew (peach, top) ─────────────────────────────────────────
gc_x0 = sab_x0 + 80
gc_x1 = sab_x1 - 80
gc_y0 = sab_y0 + 90
gc_y1 = gc_y0 + 230
d.rounded_rectangle([gc_x0, gc_y0, gc_x1, gc_y1], radius=22,
                    fill=GOV, outline=GOV_DARK, width=4)

# Tag
tag_w = 280
d.rounded_rectangle([gc_x1 - tag_w - 30, gc_y0 - 28,
                     gc_x1 - 30, gc_y0 + 30],
                    radius=14, fill=WHITE, outline=GOV_DARK, width=2)
centered(d, ((gc_x1 - tag_w - 30 + gc_x1 - 30) / 2, gc_y0 + 1),
         "Governance Crew", fnt(22, True), DARK)

gov_persons = [("Vitor Cruz", "VP IT Bus. Solutions"),
               ("Carlos Garbin", "Dir. IT Business Apps"),
               ("Kumar Jaladanki", "Dir. Customer Svc Apps"),
               ("Heather LaJoie", "Dir. Product Mgmt"),
               ("Travis", "PMO Lead"),
               ("Portfolio Mgr", "TBD")]
px0 = gc_x0 + 100
px_step = (gc_x1 - gc_x0 - 200) / (len(gov_persons) - 1)
for i, (name, title) in enumerate(gov_persons):
    cx = px0 + i * px_step
    person(d, cx, gc_y0 + 70, r=34)
    centered(d, (cx, gc_y0 + 130), name, fnt(17, True), DARK)
    centered(d, (cx, gc_y0 + 152), title, fnt(13), GREY)
left(d, (gc_x0 + 30, gc_y1 - 30),
     "The management team of the Base — sets vision, strategy, governance constraints.",
     fnt(16, italic=True), GREY)


# ── Value Stream Crews row (yellow) ──────────────────────────────────────
vs_y0 = gc_y1 + 80
vs_y1 = vs_y0 + 480

# 6 streams in a row
vs_streams = [
    ("M2C", "Meter-to-Cash",
     "TBD PM (Iram avail.)",
     ["Vincent Lam · Iram", "Ivelisse Velazquez", "Squad Metering · Billing"]),
    ("CCC", "Customer Care",
     "TBD (vacant Sr.M.)",
     ["TBD CRM/CCC mgr", "Heather (business)", "Squad CRM · CCC"]),
    ("FOPS", "Field Operations",
     "Andrew McGibbon",
     ["Bruno Moreto", "Squad Mobility", "Squad Storm/Workforce"]),
    ("E&A", "Engineering & Asset",
     "George Porto (dual)",
     ["George (Sr.M.+PM)", "Squad ArcFM · GIS", "Squad SAP HANA"]),
    ("C&R", "Corporate & Regulatory",
     "Sarath (decision pending)",
     ["Sarath + Amit", "5–6 PJMs", "2 Tech Leads"]),
    ("IBD", "Iberdrola Corp Svc",
     "Vanessa Colombo",
     ["Vanessa + Michael", "Squad FI", "Squad PO · Budget"]),
]

vs_pad = 80
vs_n = len(vs_streams)
vs_avail_w = sab_x1 - sab_x0 - 2 * vs_pad
vs_gap = 16
vs_w = (vs_avail_w - (vs_n - 1) * vs_gap) // vs_n

# Tag for VS row
tag2_w = 320
d.rounded_rectangle([sab_x0 + vs_pad, vs_y0 - 28,
                     sab_x0 + vs_pad + tag2_w, vs_y0 + 30],
                    radius=14, fill=WHITE, outline=VS_DARK, width=2)
centered(d, ((2 * (sab_x0 + vs_pad) + tag2_w) / 2, vs_y0 + 1),
         "Value Stream Crews × 6", fnt(20, True), DARK)
person(d, sab_x0 + vs_pad + 30, vs_y0 + 1, r=18)

for i, (short, name, pm, members) in enumerate(vs_streams):
    x0 = sab_x0 + vs_pad + i * (vs_w + vs_gap)
    x1 = x0 + vs_w
    d.rounded_rectangle([x0, vs_y0, x1, vs_y1], radius=18,
                        fill=VS, outline=VS_DARK, width=3)
    # Header
    d.rounded_rectangle([x0, vs_y0, x1, vs_y0 + 70], radius=18, fill=VS_DARK)
    d.rectangle([x0, vs_y0 + 50, x1, vs_y0 + 70], fill=VS_DARK)
    centered(d, ((x0 + x1) / 2, vs_y0 + 22), short, fnt(28, True), WHITE)
    centered(d, ((x0 + x1) / 2, vs_y0 + 50), name, fnt(14), WHITE)

    # Persons
    for j in range(3):
        person(d, x0 + 50 + j * 50, vs_y0 + 110, r=18)
    # PM
    centered(d, ((x0 + x1) / 2, vs_y0 + 170), "PM", fnt(14, True), DARK)
    centered(d, ((x0 + x1) / 2, vs_y0 + 192), pm, fnt(12), DARK)
    # Members
    centered(d, ((x0 + x1) / 2, vs_y0 + 230), "Members", fnt(13, True), DARK)
    for j, m in enumerate(members):
        centered(d, ((x0 + x1) / 2, vs_y0 + 256 + j * 22),
                 m, fnt(11), DARK)
    # Bottom note
    centered(d, ((x0 + x1) / 2, vs_y1 - 22),
             "End-to-end value stream", fnt(11, italic=True), DARK)


# ── Supporting Crews row (Capability / Experience / Platform / Partnership) ─
sup_y0 = vs_y1 + 70
sup_y1 = sup_y0 + 380

sup_crews = [
    (CAP, CAP_DARK, "Capability Crew", "Architecture Council",
     ["TBD Sr.M. SAP & CS Architecture", "Corporate Tech Leads (2)",
      "Cross-stream architects", "Reviews scope · feasibility"]),
    (EXP, EXP_DARK, "Experience Crew", "UX & Customer Journey",
     ["UX Designer (NOM v1.5)", "Heather LaJoie (CSP)",
      "CX Researcher (from CoE)", "Owns end-to-end customer journey"]),
    (PLAT, PLAT_DARK, "Platform Crew", "Shared Services",
     ["SWF & Tools — Raquel · Donna", "DevOps + App Sec — Piotr",
      "SAP Basis — Nagi", "TOSCA · Gold Client · 8760 EDI"]),
    (PART, PART_DARK, "Partnership Crew", "Vendor & External Mgmt",
     ["Accenture (Global SAP + Corp)", "Smart Metering vendor (100%)",
      "SWF / Software Factory (cross)", "Iberdrola corporate liaison"]),
]

sup_pad = 80
sup_n = len(sup_crews)
sup_avail_w = sab_x1 - sab_x0 - 2 * sup_pad
sup_gap = 16
sup_w = (sup_avail_w - (sup_n - 1) * sup_gap) // sup_n

for i, (col, dark, name, sub, members) in enumerate(sup_crews):
    x0 = sab_x0 + sup_pad + i * (sup_w + sup_gap)
    x1 = x0 + sup_w
    d.rounded_rectangle([x0, sup_y0, x1, sup_y1], radius=18,
                        fill=col, outline=dark, width=3)
    d.rounded_rectangle([x0, sup_y0, x1, sup_y0 + 70], radius=18, fill=dark)
    d.rectangle([x0, sup_y0 + 50, x1, sup_y0 + 70], fill=dark)
    centered(d, ((x0 + x1) / 2, sup_y0 + 22), name, fnt(22, True), WHITE)
    centered(d, ((x0 + x1) / 2, sup_y0 + 48), sub, fnt(15), WHITE)

    for j in range(3):
        person(d, x0 + 60 + j * 50, sup_y0 + 110, r=18, fill=WHITE, br=dark)
    # Members
    for j, m in enumerate(members):
        centered(d, ((x0 + x1) / 2, sup_y0 + 180 + j * 30),
                 m, fnt(13), WHITE)


# ── Facilitation Crew (cross-cutting horizontal band) ────────────────────
fc_y0 = sup_y1 + 50
fc_y1 = fc_y0 + 130
d.rounded_rectangle([sab_x0 + 80, fc_y0, sab_x1 - 80, fc_y1],
                    radius=18, fill=FAC, outline=FAC_DARK, width=3)
left(d, (sab_x0 + 110, fc_y0 + 25), "Facilitation Crew", fnt(24, True), WHITE)
left(d, (sab_x0 + 110, fc_y0 + 60),
     "Process Improvement (NOM v1.5) · Agile Coaches (from CoE Agility Chapter) · Scrum Masters per squad · Knowledge Manager",
     fnt(15), WHITE)
left(d, (sab_x0 + 110, fc_y0 + 95),
     "No own value stream — supports the 6 Value Stream Crews end-to-end",
     fnt(13, italic=True), (220, 220, 240))
# Persons
for j in range(4):
    person(d, sab_x1 - 250 + j * 50, fc_y0 + 65, r=18, fill=WHITE, br=FAC_DARK)


# ── Turf label (bottom of base) ──────────────────────────────────────────
turf_y = fc_y1 + 40
left(d, (sab_x0 + 110, turf_y),
     "Turf — area cultivated and protected by the same people (squads belong to a Value Stream Crew).",
     fnt(15, italic=True), GREY)

# Base label
centered(d, ((sab_x0 + sab_x1) / 2, sab_y1 + 50),
         "Strongly Aligned Base — Digital Products",
         fnt(30, True), DARK)
centered(d, ((sab_x0 + sab_x1) / 2, sab_y1 + 88),
         "Multiple value streams forming a cohesive product set with dependencies — coordinated by the Governance Crew.",
         fnt(17, italic=True), GREY)


# ── Forums column (right side) ───────────────────────────────────────────
fo_x0 = sab_x1 + 80
fo_x1 = W - 100
fo_y0 = sab_y0 + 90
fo_h = sab_y1 - sab_y0 - 90

# Header for forums area
d.text((fo_x0, fo_y0 - 50), "FORUMS",
       font=fnt(36, True), fill=DARK)
d.text((fo_x0, fo_y0 - 14),
       "Cross-Crew discussion venues — official Unfix Forum types",
       font=fnt(14, italic=True), fill=GREY)

forums = [
    ("Regional Forum",        "Per OPCO discussion",
     ["BRS UIL — UI · SCG · CNG · BGC",
      "BRS CMP — Central Maine Power",
      "BRS NY — NYSEG · RGE"]),
    ("Functional Forum",      "Per role / function",
     ["PM Council (cross-stream)",
      "Tech Lead Council",
      "Project Manager pool (Sarath's)"]),
    ("Technological Forum",   "Per tech stack",
     ["SAP Forum — Local + Iberdrola",
      "Non-SAP Forum — GIS · Mobility · Doc Mgmt",
      "Platform Forum — SWF · DevOps · Sec"]),
    ("Customer Journey Forum","End-to-end UX",
     ["M2C journey — meter→bill→pay",
      "Self-service journey",
      "Field worker journey"]),
    ("Product Forum",         "Per product strategy",
     ["Quarterly BRP cadence",
      "Cross-stream roadmap sync",
      "Trade-offs + sequencing"]),
]

f_n = len(forums)
f_gap = 18
f_h = (fo_h - (f_n - 1) * f_gap) // f_n

for i, (name, sub, items) in enumerate(forums):
    y0 = fo_y0 + i * (f_h + f_gap)
    y1 = y0 + f_h
    d.rounded_rectangle([fo_x0, y0, fo_x1, y1], radius=14,
                        fill=WHITE, outline=DARK, width=2)
    # Tab
    d.rounded_rectangle([fo_x0, y0, fo_x0 + 200, y0 + 40],
                        radius=14, fill=DARK)
    d.rectangle([fo_x0, y0 + 25, fo_x0 + 200, y0 + 40], fill=DARK)
    centered(d, (fo_x0 + 100, y0 + 20), "Forum", fnt(15, True), WHITE)

    left(d, (fo_x0 + 215, y0 + 18), name, fnt(20, True), DARK)
    left(d, (fo_x0 + 20, y0 + 65), sub, fnt(15, italic=True), GREY)
    for j, it in enumerate(items):
        left(d, (fo_x0 + 30, y0 + 100 + j * 28),
             "• " + it, fnt(14), DARK)


# ── Sibling Bases (bottom — Digital CoE + P&S/P&T) ───────────────────────
sib_y0 = sab_y1 + 130
sib_h  = 280

# Digital CoE
coe_x0 = 200
coe_x1 = coe_x0 + 1500
d.rounded_rectangle([coe_x0, sib_y0, coe_x1, sib_y0 + sib_h],
                    radius=30, fill=SIBLING_BG, outline=SIBLING_BR, width=3)
home_icon(d, (coe_x0 + coe_x1) / 2, sib_y0 - 30, r=28)
centered(d, ((coe_x0 + coe_x1) / 2, sib_y0 + 30),
         "Digital CoE — Sibling Base", fnt(24, True), DARK)
centered(d, ((coe_x0 + coe_x1) / 2, sib_y0 + 64),
         "Loosely Aligned · 5 Value Stream Crews aligned by OPCO (CNG · SCG · UI · CMP · NYSEG)",
         fnt(15, italic=True), GREY)
# Some persons
for j in range(8):
    person(d, coe_x0 + 100 + j * 170, sib_y0 + 130, r=22, fill=PERSON_F)
left(d, (coe_x0 + 30, sib_y0 + sib_h - 50),
     "Customer-facing apps · websites · contact center digital · separate from SAP/IT-Apps · Heather as cross-base bridge",
     fnt(15, italic=True), GREY)

# P&S/P&T
ps_x0 = coe_x1 + 80
ps_x1 = W - 200
d.rounded_rectangle([ps_x0, sib_y0, ps_x1, sib_y0 + sib_h],
                    radius=30, fill=SIBLING_BG, outline=SIBLING_BR, width=3)
home_icon(d, (ps_x0 + ps_x1) / 2, sib_y0 - 30, r=28)
centered(d, ((ps_x0 + ps_x1) / 2, sib_y0 + 30),
         "P&S / P&T — Sibling Base", fnt(24, True), DARK)
centered(d, ((ps_x0 + ps_x1) / 2, sib_y0 + 64),
         "Loosely Aligned · NOM v1.5 — Networks IT (utility infrastructure)",
         fnt(15, italic=True), GREY)
for j in range(8):
    person(d, ps_x0 + 100 + j * 175, sib_y0 + 130, r=22, fill=PERSON_F)
left(d, (ps_x0 + 30, sib_y0 + sib_h - 50),
     "Reference Operating Model (Wave 1 in implementation) — playbook Digital Products is adapting from",
     fnt(15, italic=True), GREY)


# ── Footer ───────────────────────────────────────────────────────────────
d.text((80, H - 40),
       "Source: unfix.com — official 7 Crew types + 10 Forum types + 4 Base types · Avangrid Digital Products TOM · Apr 2026",
       font=fnt(18), fill=GREY)


OUT.mkdir(parents=True, exist_ok=True)
out_file = OUT / "DigitalProducts_Unfix_Orthodox.png"
img.save(out_file, "PNG", dpi=(150, 150))

ratio = 1600 / W
preview = img.resize((1600, int(H * ratio)), Image.LANCZOS)
preview.save(out_file.with_name(out_file.stem + "_preview.png"), "PNG")
print(f"✓ Saved: {out_file}")
print(f"  Resolution: {W}×{H}")
