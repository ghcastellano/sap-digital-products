"""
Digital Products — Unfix model in OVERLAP / nested-layer style
Mimics the user's reference image precisely:
- Governance Crew overlaps top of Base
- Forums overlap each other (z-order)
- Crews stack inside Forums
- Vertical Facilitation band cuts across Value Stream Crews
- Segregated Base (SW Delivery) overlaps right edge
- Uses ONLY official Unfix 7 Crew types and 10 Forum types
"""
from PIL import Image, ImageDraw, ImageFont
from pathlib import Path
import math

OUT = Path("/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Miro_Visuals")

# ── Palette (mirrors the user's reference image) ─────────────────────────
BASE_BG    = (231, 232, 234)      # light grey base
BASE_BR    = (170, 175, 185)
INNER_BG   = (252, 252, 253)
FORUM_BR   = ( 90, 145,  85)      # green forum borders
FORUM_BG   = (252, 254, 252)
GOV        = (251, 195, 145)      # peach
GOV_BR     = (210, 145,  80)

# Crews
VS         = (215, 200, 245)      # Value Stream — lavender (matches user's "Program Crew" purple)
VS_BR      = (160, 140, 210)
EXP        = (188, 220, 188)      # Experience Crew — green
EXP_BR     = (110, 160, 110)
CAP        = (188, 220, 188)      # Capability Crew — green (Unfix standard)
CAP_BR     = (110, 160, 110)
PLAT       = (188, 220, 188)      # Platform Crew — green
PLAT_BR    = (110, 160, 110)
PART       = (250, 226, 145)      # Partnership Crew — yellow
PART_BR    = (200, 175,  80)
FAC_BAND   = (190, 195, 230)      # Facilitation vertical band — light purple
FAC_BR     = (130, 140, 200)

WHITE      = (255, 255, 255)
DARK       = ( 30,  40,  55)
GREY       = ( 90, 100, 115)
LIGHT_TXT  = (130, 140, 155)
PERSON_F   = (210, 215, 225)
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


def person(d, cx, cy, r=20, fill=PERSON_F, br=PERSON_BR):
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=fill, outline=br, width=2)
    head_r = r * 0.38
    d.ellipse([cx - head_r, cy - r * 0.55, cx + head_r, cy - r * 0.55 + 2 * head_r],
              fill=br)
    d.pieslice([cx - r * 0.7, cy - r * 0.05, cx + r * 0.7, cy + r * 1.4],
               start=180, end=360, fill=br)


def home_dashed(d, cx, cy, r=44, color=DARK):
    """Home icon with dashed circle around it."""
    for ang in range(0, 360, 12):
        rad = math.radians(ang)
        x1 = cx + (r + 8) * math.cos(rad)
        y1 = cy + (r + 8) * math.sin(rad)
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


def dashed_circle(d, cx, cy, r, color=BASE_BR, width=4):
    """Draw a dashed circle outline."""
    for ang in range(0, 360, 4):
        rad = math.radians(ang)
        rad2 = math.radians(ang + 2)
        x1 = cx + r * math.cos(rad)
        y1 = cy + r * math.sin(rad)
        x2 = cx + r * math.cos(rad2)
        y2 = cy + r * math.sin(rad2)
        d.line([x1, y1, x2, y2], fill=color, width=width)


def forum_tag(d, x, y, label, font_size=22):
    """White pill with green outline + person icon — Unfix forum tag style."""
    tag_w = 360
    tag_h = 56
    d.rounded_rectangle([x, y, x + tag_w, y + tag_h], radius=14,
                        fill=WHITE, outline=FORUM_BR, width=3)
    person(d, x + 30, y + tag_h / 2, r=18)
    d.text((x + 65, y + tag_h / 2), label, font=fnt(font_size, True),
           fill=DARK, anchor="lm")
    return tag_w, tag_h


img = Image.new("RGB", (W, H), WHITE)
d = ImageDraw.Draw(img)


# ── Header ───────────────────────────────────────────────────────────────
d.rectangle([0, 0, W, 130], fill=DARK)
d.text((80, 28), "DIGITAL PRODUCTS — unFIX MODEL (overlap layout)",
       font=fnt(54, True), fill=WHITE)
d.text((80, 88),
       "Strongly Aligned Base · 7 Crew types · 4 Forum types · Sibling Segregated Base for SW Delivery",
       font=fnt(20), fill=(199, 211, 228))


# ── Home icon at top of page ─────────────────────────────────────────────
home_dashed(d, 1700, 250, r=42)


# ════════════════════════════════════════════════════════════════════════
#  LAYER 1 — Strongly Aligned Base (the big grey background)
# ════════════════════════════════════════════════════════════════════════
sab_x0, sab_y0 = 200, 380
sab_x1, sab_y1 = 3450, 2680
d.rounded_rectangle([sab_x0, sab_y0, sab_x1, sab_y1], radius=44,
                    fill=BASE_BG, outline=BASE_BR, width=4)


# ════════════════════════════════════════════════════════════════════════
#  LAYER 2 — Governance Crew (overlaps top edge of Base)
# ════════════════════════════════════════════════════════════════════════
gc_x0 = 700
gc_x1 = 2700
gc_y0 = 320     # starts ABOVE sab_y0 = 380 → overlap by 60px
gc_y1 = 580
d.rounded_rectangle([gc_x0, gc_y0, gc_x1, gc_y1], radius=24,
                    fill=GOV, outline=GOV_BR, width=4)
# Tag pill
tag_x = gc_x1 - 320
tag_y = gc_y0 - 30
d.rounded_rectangle([tag_x, tag_y, tag_x + 280, tag_y + 56],
                    radius=14, fill=WHITE, outline=GOV_BR, width=3)
person(d, tag_x + 30, tag_y + 28, r=18)
d.text((tag_x + 65, tag_y + 28), "Governance Crew",
       font=fnt(20, True), fill=DARK, anchor="lm")

gov_persons = [("Vitor Cruz", "VP IT Bus."),
               ("Carlos Garbin", "Dir. IT Apps"),
               ("Kumar Jaladanki", "Dir. CS Apps"),
               ("Heather LaJoie", "Dir. Product"),
               ("Travis", "PMO"),
               ("Portfolio Mgr", "TBD")]
spacing = (gc_x1 - gc_x0 - 140) / (len(gov_persons) - 1)
for i, (nm, role) in enumerate(gov_persons):
    cx = gc_x0 + 70 + i * spacing
    person(d, cx, gc_y0 + 60, r=30)
    centered(d, (cx, gc_y0 + 110), nm, fnt(15, True), DARK)
    centered(d, (cx, gc_y0 + 132), role, fnt(12), GREY)
left(d, (gc_x0 + 30, gc_y1 - 30),
     "The management team of the Base — sets vision, strategy, governance constraints.",
     fnt(14, italic=True), GREY)


# ════════════════════════════════════════════════════════════════════════
#  LAYER 3 — Forums (overlap each other in Z-order)
# ════════════════════════════════════════════════════════════════════════
forum_y0 = 720          # starts below Governance Crew
forum_y1 = 2440         # ends above Turf strip

# ── 3a: Customer Journey Forum (left) ────────────────────────────────────
cj_x0 = 280
cj_x1 = 980
d.rounded_rectangle([cj_x0, forum_y0, cj_x1, forum_y1], radius=22,
                    fill=FORUM_BG, outline=FORUM_BR, width=4)
forum_tag(d, cj_x0 + 24, forum_y0 - 28, "Customer Journey Forum")


# ── 3b: Product Forum (center, biggest) ──────────────────────────────────
pf_x0 = 1020
pf_x1 = 2380
d.rounded_rectangle([pf_x0, forum_y0, pf_x1, forum_y1], radius=22,
                    fill=FORUM_BG, outline=FORUM_BR, width=4)
forum_tag(d, pf_x0 + 24, forum_y0 - 28, "Product Forum")


# ── 3c: Organizational Forum (small, top-right, overlaps Product Forum) ──
of_x0 = 2280            # overlaps Product Forum (which ends at 2380)
of_x1 = 3000
of_y0 = forum_y0 - 60   # sits higher
of_y1 = forum_y0 + 230
d.rounded_rectangle([of_x0, of_y0, of_x1, of_y1], radius=22,
                    fill=FORUM_BG, outline=FORUM_BR, width=4)
forum_tag(d, of_x0 + 24, of_y0 - 28, "Organizational Forum")
# Partnership Crew nested inside — yellow card
part_x0 = of_x0 + 36
part_x1 = of_x1 - 36
part_y0 = of_y0 + 80
part_y1 = of_y1 - 24
d.rounded_rectangle([part_x0, part_y0, part_x1, part_y1], radius=14,
                    fill=PART, outline=PART_BR, width=3)
left(d, (part_x0 + 20, part_y0 + 24),
     "Partnership Crew", fnt(20, True), DARK)
left(d, (part_x0 + 20, part_y0 + 56),
     "Vendors · Accenture · SWF · Iberdrola liaison",
     fnt(13), DARK)
for j in range(3):
    person(d, part_x1 - 30 - j * 36, part_y0 + 70, r=14)


# ── 3d: Technological Forum (right, overlaps with Product Forum AND Organizational Forum) ─
tf_x0 = 2300            # overlaps Product Forum (2380) and Organizational Forum (3000)
tf_x1 = 3320
tf_y0 = of_y1 - 30      # starts where Organizational Forum ends (overlap)
tf_y1 = forum_y1
d.rounded_rectangle([tf_x0, tf_y0, tf_x1, tf_y1], radius=22,
                    fill=FORUM_BG, outline=FORUM_BR, width=4)
forum_tag(d, tf_x0 + 24, tf_y0 - 28, "Technological Forum")


# ════════════════════════════════════════════════════════════════════════
#  LAYER 4 — Crews inside Customer Journey Forum
# ════════════════════════════════════════════════════════════════════════
# Experience Crew (green box) — top of Customer Journey Forum
exp_x0 = cj_x0 + 36
exp_x1 = cj_x1 - 36
exp_y0 = forum_y0 + 80
exp_y1 = exp_y0 + 460
d.rounded_rectangle([exp_x0, exp_y0, exp_x1, exp_y1], radius=18,
                    fill=EXP, outline=EXP_BR, width=3)
centered(d, ((exp_x0 + exp_x1) / 2, exp_y0 + 30),
         "Experience Crew", fnt(22, True), DARK)
for j in range(4):
    person(d, exp_x0 + 60 + j * 56, exp_y0 + 100, r=22)
left(d, (exp_x0 + 24, exp_y0 + 160),
     "UX & customer journey lens:", fnt(14, True), DARK)
exp_lines = [
    "UX Designer (NOM v1.5 role)",
    "Heather LaJoie (CSP)",
    "CX Researcher (from Digital CoE)",
    "Owns end-to-end customer journey",
    "Reminds streams to think holistically",
]
for j, ln in enumerate(exp_lines):
    left(d, (exp_x0 + 24, exp_y0 + 200 + j * 28),
         "• " + ln, fnt(13), DARK)
left(d, (exp_x0 + 24, exp_y1 - 26),
     "Ensures customer experience is great.",
     fnt(12, italic=True), GREY)

# Capability Crew (green) — bottom of Customer Journey Forum
cap_y0 = exp_y1 + 30
cap_y1 = forum_y1 - 60
d.rounded_rectangle([exp_x0, cap_y0, exp_x1, cap_y1], radius=18,
                    fill=CAP, outline=CAP_BR, width=3)
centered(d, ((exp_x0 + exp_x1) / 2, cap_y0 + 30),
         "Capability Crew", fnt(22, True), DARK)
for j in range(4):
    person(d, exp_x0 + 60 + j * 56, cap_y0 + 100, r=22)
left(d, (exp_x0 + 24, cap_y0 + 160),
     "Architecture Council:", fnt(14, True), DARK)
cap_lines = [
    "TBD Sr.M. SAP & CS Architecture",
    "Corporate Tech Leads (2)",
    "Cross-stream architects",
    "Reviews scope · feasibility",
]
for j, ln in enumerate(cap_lines):
    left(d, (exp_x0 + 24, cap_y0 + 200 + j * 28),
         "• " + ln, fnt(13), DARK)
left(d, (exp_x0 + 24, cap_y1 - 26),
     "Rare expertise made available to other Crews.",
     fnt(12, italic=True), GREY)


# ════════════════════════════════════════════════════════════════════════
#  LAYER 5 — Value Stream Crews (purple bars) inside Product Forum
# ════════════════════════════════════════════════════════════════════════
vs_streams = [
    ("Value Stream Crew · M2C",
     "Vincent Lam · Iram · Ivelisse — Squad Metering · Billing · Collections"),
    ("Value Stream Crew · Customer Care & CCC",
     "TBD Sr.M. (vacant) · Heather (business) — Squad CRM · CCC"),
    ("Value Stream Crew · Field Operations",
     "Bruno Moreto · Andrew McGibbon — Squad Mobility · Doc Viewer · Storm"),
    ("Value Stream Crew · Engineering & Asset Mgmt",
     "George Porto — Squad ArcFM · GIS Web · SAP HANA · OSG"),
    ("Value Stream Crew · Corporate & Regulatory",
     "Sarath · Amit — 5–6 PJMs + 2 Tech Leads"),
    ("Value Stream Crew · Iberdrola Corp Services",
     "Vanessa Colombo · Michael — Squad FI · PO · Budget"),
]

vs_block_y0 = forum_y0 + 80
vs_block_y1 = forum_y1 - 80
vs_avail_h = vs_block_y1 - vs_block_y0
vs_n = len(vs_streams)
vs_gap = 14
vs_h = (vs_avail_h - (vs_n - 1) * vs_gap) // vs_n

vs_x0_card = pf_x0 + 36
vs_x1_card = pf_x1 - 36

for i, (name, members) in enumerate(vs_streams):
    y0 = vs_block_y0 + i * (vs_h + vs_gap)
    y1 = y0 + vs_h
    d.rounded_rectangle([vs_x0_card, y0, vs_x1_card, y1], radius=14,
                        fill=VS, outline=VS_BR, width=3)
    left(d, (vs_x0_card + 24, y0 + vs_h / 2 - 14),
         name, fnt(18, True), DARK)
    left(d, (vs_x0_card + 24, y0 + vs_h / 2 + 14),
         members, fnt(12), GREY)
    # Person icons on the right edge
    for j in range(3):
        person(d, vs_x1_card - 50 - j * 40, y0 + vs_h / 2, r=15)


# ════════════════════════════════════════════════════════════════════════
#  LAYER 6 — Facilitation Crew (vertical band cutting across Value Streams)
#            Drawn AFTER Value Streams so it visually cuts through them.
# ════════════════════════════════════════════════════════════════════════
fac_x0 = pf_x0 + 540
fac_x1 = fac_x0 + 110
fac_y0 = vs_block_y0 - 8
fac_y1 = vs_block_y1 + 8
d.rounded_rectangle([fac_x0, fac_y0, fac_x1, fac_y1], radius=22,
                    fill=FAC_BAND, outline=FAC_BR, width=3)
# Vertical text "Facilitation Crew"
tmp = Image.new("RGBA", (fac_y1 - fac_y0 - 60, fac_x1 - fac_x0), (0, 0, 0, 0))
td = ImageDraw.Draw(tmp)
td.text((30, (fac_x1 - fac_x0) / 2), "Facilitation Crew",
        fill=DARK, font=fnt(22, True), anchor="lm")
tmp = tmp.rotate(90, expand=True)
img.paste(tmp, (fac_x0, fac_y0 + 30), tmp)


# ════════════════════════════════════════════════════════════════════════
#  LAYER 7 — Crews inside Technological Forum (right)
# ════════════════════════════════════════════════════════════════════════
plat_x0 = tf_x0 + 36
plat_x1 = tf_x1 - 36
plat_y0 = tf_y0 + 90
plat_y1 = tf_y1 - 60

d.rounded_rectangle([plat_x0, plat_y0, plat_x1, plat_y1], radius=18,
                    fill=PLAT, outline=PLAT_BR, width=3)
centered(d, ((plat_x0 + plat_x1) / 2, plat_y0 + 32),
         "Platform Crew", fnt(26, True), DARK)
for j in range(5):
    person(d, plat_x0 + 60 + j * 56, plat_y0 + 110, r=22)
left(d, (plat_x0 + 24, plat_y0 + 175),
     "Shared services to all Crews:", fnt(15, True), DARK)
plat_lines = [
    "SWF & Tools — Raquel · Donna",
    "DevOps — Piotr Wielogorski",
    "App Security — Piotr Wielogorski",
    "SAP Basis — Nagi Reddy",
    "Test Automation (TOSCA)",
    "Data Ops (Gold Client · 8760 EDI)",
    "Architecture · APIs · environments",
]
for j, ln in enumerate(plat_lines):
    left(d, (plat_x0 + 24, plat_y0 + 215 + j * 28),
         "• " + ln, fnt(13), DARK)
left(d, (plat_x0 + 24, plat_y1 - 30),
     "No own value stream — speeds up the others.",
     fnt(12, italic=True), GREY)


# ════════════════════════════════════════════════════════════════════════
#  LAYER 8 — Turf strip at bottom of Base
# ════════════════════════════════════════════════════════════════════════
turf_y = forum_y1 + 50
left(d, (sab_x0 + 60, turf_y),
     "The Turf is an area cultivated and protected by the same people.",
     fnt(18, italic=True), GREY)
for i in range(3):
    person(d, sab_x0 + 1900 + i * 80, turf_y + 18, r=22)


# ════════════════════════════════════════════════════════════════════════
#  LAYER 9 — Base label below
# ════════════════════════════════════════════════════════════════════════
centered(d, ((sab_x0 + sab_x1) / 2, sab_y1 + 50),
         "Strongly Aligned Base — Digital Products",
         fnt(30, True), DARK)
centered(d, ((sab_x0 + sab_x1) / 2, sab_y1 + 90),
         "Multiple Value Stream Crews forming a cohesive product set with dependencies between them.",
         fnt(17, italic=True), GREY)


# ════════════════════════════════════════════════════════════════════════
#  LAYER 10 — Segregated Base on the right (overlaps page edge)
# ════════════════════════════════════════════════════════════════════════
seg_cx = 3950
seg_cy = 1500
seg_r  = 480
home_dashed(d, seg_cx + 360, seg_cy - 360, r=34)
dashed_circle(d, seg_cx, seg_cy, seg_r, color=BASE_BR, width=4)

# Vendors inside the segregated base
vendors = [("Accenture", "Global SAP + Corp"),
           ("SWF / Software Factory", "All teams"),
           ("HCL · Nuance", "Customer Service"),
           ("Infosys", "System support"),
           ("Smart Metering vendor", "100% delivery"),
           ("SAP partners", "Iberdrola side")]
ncols = 3
for i, (nm, role) in enumerate(vendors):
    row = i // ncols
    col = i % ncols
    cx = seg_cx - 230 + col * 230
    cy = seg_cy - 100 + row * 160
    person(d, cx, cy, r=32, fill=PERSON_F)
    centered(d, (cx, cy + 56), nm, fnt(14, True), DARK)
    centered(d, (cx, cy + 76), role, fnt(11), GREY)

centered(d, (seg_cx, seg_cy + seg_r - 50),
         "SW Delivery Base", fnt(28, True), DARK)
centered(d, (seg_cx, seg_cy + seg_r - 14),
         "(Fully Segregated Base)", fnt(16, italic=True), GREY)
centered(d, (seg_cx, seg_cy + seg_r + 30),
         "External delivery — independent governance, vendors operate outside Avangrid Base",
         fnt(14, italic=True), GREY)


# ── Connector arrow from Strongly Aligned Base → Segregated Base ─────────
arrow_y = 1500
d.line([sab_x1, arrow_y, seg_cx - seg_r, arrow_y], fill=DARK, width=4)
d.polygon([(seg_cx - seg_r, arrow_y),
           (seg_cx - seg_r - 22, arrow_y - 12),
           (seg_cx - seg_r - 22, arrow_y + 12)],
          fill=DARK)


# ── Footer ───────────────────────────────────────────────────────────────
d.text((80, H - 40),
       "Source: unfix.com — official 7 Crew types · 10 Forum types · 4 Base types · Avangrid Digital Products TOM · Apr 2026",
       font=fnt(18), fill=GREY)


OUT.mkdir(parents=True, exist_ok=True)
out_file = OUT / "DigitalProducts_Unfix_Overlap.png"
img.save(out_file, "PNG", dpi=(150, 150))

ratio = 1600 / W
preview = img.resize((1600, int(H * ratio)), Image.LANCZOS)
preview.save(out_file.with_name(out_file.stem + "_preview.png"), "PNG")
print(f"✓ Saved: {out_file}")
print(f"  Resolution: {W}×{H}")
