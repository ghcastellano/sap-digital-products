"""
Digital Products — Unfix Bases / Forums / Crews layered visualization.
Following the user's reference image style (Strongly Aligned Base + Segregated Base).
"""
from PIL import Image, ImageDraw, ImageFont
from pathlib import Path
import math

OUT = Path("/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Miro_Visuals")

# ── Palette (matched to user's reference) ────────────────────────────────
BASE_BG     = (231, 232, 234)
BASE_BR     = (180, 182, 188)
GOV_BG      = (251, 211, 169)
GOV_BR      = (210, 165, 110)
FORUM_BR    = ( 90, 145,  85)
FORUM_BG    = (250, 252, 250)
PROG_BG     = (213, 200, 240)
PROG_BR     = (160, 140, 210)
KAIZEN_BG   = (250, 226, 145)
KAIZEN_BR   = (200, 175,  80)
CAP_BG      = (180, 220, 175)
CAP_BR      = ( 90, 145,  85)
CHANGE_BG   = (190, 195, 230)
CHANGE_BR   = (130, 140, 200)
EXP_BG      = (188, 220, 188)
EXP_BR      = (110, 160, 110)
PERSON_FILL = (200, 210, 220)
PERSON_BR   = (150, 160, 175)
WHITE       = (255, 255, 255)
DARK        = ( 40,  50,  65)
GREY_TXT    = ( 90, 100, 115)
LIGHT_TXT   = (130, 140, 155)


W, H = 4400, 3200


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


def person_icon(d, cx, cy, r=20, fill=PERSON_FILL, border=PERSON_BR):
    """Draw a circular person silhouette icon."""
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=fill, outline=border, width=2)
    # head
    head_r = r * 0.35
    d.ellipse([cx - head_r, cy - r * 0.55, cx + head_r, cy - r * 0.55 + 2 * head_r],
              fill=border)
    # shoulders (chord)
    d.pieslice([cx - r * 0.7, cy - r * 0.05, cx + r * 0.7, cy + r * 1.4],
               start=180, end=360, fill=border)


def home_icon(d, cx, cy, r=40, color=DARK):
    """Simple home (house) icon used in Unfix to mark Bases."""
    # circle dashed look — solid for simplicity
    for ang in range(0, 360, 12):
        rad = math.radians(ang)
        x1 = cx + (r + 10) * math.cos(rad)
        y1 = cy + (r + 10) * math.sin(rad)
        x2 = cx + (r + 14) * math.cos(rad)
        y2 = cy + (r + 14) * math.sin(rad)
        d.line([x1, y1, x2, y2], fill=color, width=3)
    # house shape
    house_w, house_h = r * 1.2, r * 1.0
    base = [(cx - house_w/2, cy - house_h/4),
            (cx + house_w/2, cy - house_h/4),
            (cx + house_w/2, cy + house_h/2),
            (cx - house_w/2, cy + house_h/2)]
    d.polygon(base, fill=DARK)
    # roof
    d.polygon([(cx - house_w/2 - 6, cy - house_h/4),
               (cx, cy - house_h * 0.85),
               (cx + house_w/2 + 6, cy - house_h/4)],
              fill=DARK)
    # door
    door_w = house_w * 0.25
    door_h = house_h * 0.45
    d.rectangle([cx - door_w/2, cy + house_h/2 - door_h,
                 cx + door_w/2, cy + house_h/2], fill=WHITE)


img = Image.new("RGB", (W, H), WHITE)
d = ImageDraw.Draw(img)


# ── Page header ──────────────────────────────────────────────────────────
d.rectangle([0, 0, W, 140], fill=DARK)
d.text((80, 30), "DIGITAL PRODUCTS — UNFIX BASES", font=fnt(60, True), fill=WHITE)
d.text((80, 95), "Strongly Aligned Base · 3 Forums · 7 Crews · Segregated Base for SW Delivery (vendors)",
       font=fnt(24), fill=(199, 211, 228))


# ── Layout coordinates ───────────────────────────────────────────────────
# Strongly Aligned Base (the big grey container)
sab_x0, sab_y0 = 200, 280
sab_x1, sab_y1 = 3200, 2700
d.rounded_rectangle([sab_x0, sab_y0, sab_x1, sab_y1], radius=40,
                    fill=BASE_BG, outline=BASE_BR, width=4)

# Home icon for the Base
home_icon(d, (sab_x0 + sab_x1) / 2, sab_y0 - 40, r=42)

# ── Governance Crew (top, orange, spanning the base width) ───────────────
gc_pad = 80
gc_x0 = sab_x0 + gc_pad
gc_x1 = sab_x1 - gc_pad
gc_y0 = sab_y0 + 90
gc_y1 = gc_y0 + 240
d.rounded_rectangle([gc_x0, gc_y0, gc_x1, gc_y1], radius=24,
                    fill=GOV_BG, outline=GOV_BR, width=4)
# Banner pill on right of governance
banner_x0 = gc_x1 - 320
banner_y0 = gc_y0 - 26
d.rounded_rectangle([banner_x0, banner_y0, gc_x1 - 50, banner_y0 + 52],
                    radius=14, fill=WHITE, outline=GOV_BR, width=2)
centered(d, ((banner_x0 + gc_x1 - 50) / 2, banner_y0 + 26),
         "Governance Crew", fnt(22, True), DARK)

# Governance crew person icons
gov_persons = ["Vitor Cruz · VP", "Carlos Garbin", "Kumar Jaladanki",
               "Heather LaJoie", "Travis · PMO", "Portfolio Manager"]
px = gc_x0 + 80
for i, label in enumerate(gov_persons):
    person_icon(d, px + i * 280, gc_y0 + 60, r=36)
    label_lines = label.split(" · ")
    centered(d, (px + i * 280, gc_y0 + 110),
             label_lines[0], fnt(18, True), DARK)
    if len(label_lines) > 1:
        centered(d, (px + i * 280, gc_y0 + 134),
                 label_lines[1], fnt(15), GREY_TXT)
# Description
left(d, (gc_x0 + 30, gc_y1 - 36),
     "The Governance Crew is the management team of the Base.",
     fnt(20, italic=True), GREY_TXT)


# ── Three Forums (Discovery · Value Chain · IT Projects) ─────────────────
forum_y0 = gc_y1 + 80
forum_y1 = sab_y1 - 200

# Widths
disc_w  = 580
ic_pad  = 30
vc_w    = 1340
itp_w   = 880

disc_x0 = sab_x0 + 80
disc_x1 = disc_x0 + disc_w
vc_x0   = disc_x1 + ic_pad
vc_x1   = vc_x0 + vc_w
itp_x0  = vc_x1 + ic_pad
itp_x1  = itp_x0 + itp_w


# ── Discovery Forum (left) ───────────────────────────────────────────────
d.rounded_rectangle([disc_x0, forum_y0, disc_x1, forum_y1], radius=22,
                    fill=FORUM_BG, outline=FORUM_BR, width=4)
# Forum tag
d.rounded_rectangle([disc_x0 + 20, forum_y0 - 24, disc_x0 + 280, forum_y0 + 28],
                    radius=12, fill=WHITE, outline=FORUM_BR, width=2)
centered(d, ((disc_x0 + 20 + disc_x0 + 280) / 2, forum_y0 + 2),
         "Discovery Forum", fnt(22, True), DARK)
# Person icon in tag (left)
person_icon(d, disc_x0 + 38, forum_y0 + 2, r=18)

# Discovery Crew at top
dc_y0 = forum_y0 + 70
dc_y1 = dc_y0 + 360
d.rounded_rectangle([disc_x0 + 30, dc_y0, disc_x1 - 30, dc_y1],
                    radius=16, fill=WHITE, outline=DARK, width=2)
centered(d, ((disc_x0 + disc_x1) / 2, dc_y0 + 30),
         "Discovery Crew", fnt(24, True), DARK)
# Persons
for i in range(4):
    person_icon(d, disc_x0 + 110 + i * 110, dc_y0 + 100, r=22)
left(d, (disc_x0 + 50, dc_y0 + 160),
     "PM Council:", fnt(18, True), DARK)
disc_pms = ["Andrew McGibbon · Field Ops",
            "George Porto · Eng & Asset",
            "Sarath / TBD · Corporate",
            "Vanessa · Iberdrola Corp",
            "TBD · M2C  /  TBD · CCC"]
for i, ln in enumerate(disc_pms):
    left(d, (disc_x0 + 60, dc_y0 + 200 + i * 30),
         "• " + ln, fnt(16), DARK)
# Description
left(d, (disc_x0 + 50, dc_y1 - 30),
     "Demand intake & opportunity framing",
     fnt(15, italic=True), GREY_TXT)

# Experience Crew nested inside Discovery Forum (lower)
ec_y0 = dc_y1 + 30
ec_y1 = forum_y1 - 60
d.rounded_rectangle([disc_x0 + 30, ec_y0, disc_x1 - 30, ec_y1],
                    radius=18, fill=EXP_BG, outline=EXP_BR, width=3)
left(d, (disc_x0 + 50, ec_y0 + 20),
     "Experience Crew", fnt(22, True), DARK)
for i in range(3):
    person_icon(d, disc_x0 + 70 + i * 64, ec_y0 + 80, r=20)
left(d, (disc_x0 + 50, ec_y0 + 130),
     "UX Designer · CX Researcher", fnt(15), DARK)
left(d, (disc_x0 + 50, ec_y0 + 156),
     "Heather LaJoie (Customer Service Products)",
     fnt(15), DARK)
left(d, (disc_x0 + 50, ec_y0 + 186),
     "Owns usability + feedback loops",
     fnt(14, italic=True), GREY_TXT)


# ── Value Chain Forum (center) ───────────────────────────────────────────
d.rounded_rectangle([vc_x0, forum_y0, vc_x1, forum_y1], radius=22,
                    fill=FORUM_BG, outline=FORUM_BR, width=4)
# Forum tag
d.rounded_rectangle([vc_x0 + 20, forum_y0 - 24, vc_x0 + 320, forum_y0 + 28],
                    radius=12, fill=WHITE, outline=FORUM_BR, width=2)
centered(d, ((vc_x0 + 20 + vc_x0 + 320) / 2, forum_y0 + 2),
         "Value Chain Forum", fnt(22, True), DARK)
person_icon(d, vc_x0 + 38, forum_y0 + 2, r=18)

# Kaizen Crew (yellow, top)
kz_y0 = forum_y0 + 70
kz_y1 = kz_y0 + 100
d.rounded_rectangle([vc_x0 + 30, kz_y0, vc_x1 - 30, kz_y1],
                    radius=14, fill=KAIZEN_BG, outline=KAIZEN_BR, width=3)
left(d, (vc_x0 + 60, kz_y0 + 30),
     "Kaizen Crew", fnt(22, True), DARK)
left(d, (vc_x0 + 60, kz_y0 + 60),
     "Process Improvement (NOM v1.5) · cross-stream Kaizen",
     fnt(15, italic=True), GREY_TXT)
for i in range(3):
    person_icon(d, vc_x1 - 90 - i * 50, kz_y0 + 50, r=18)

# Program Crews (the 6 value streams)
program_streams = [
    ("Program Crew · M2C",
     "Vincent Lam · Iram · Ivelisse — Squad Metering · Billing · Collections"),
    ("Program Crew · Customer Care & CCC",
     "TBD Sr. Mgr (vacant) · Heather (business) — Squad CRM · Self-Service · CCC"),
    ("Program Crew · Field Operations",
     "Bruno Moreto · Andrew McGibbon — Squad Mobility · Doc Viewer · Storm"),
    ("Program Crew · Engineering & Asset Mgmt",
     "George Porto — Squad ArcFM · GIS Web · SAP HANA · OSG"),
    ("Program Crew · Corporate & Regulatory",
     "Sarath · Amit — 5–6 PJMs + 2 Tech Leads · Corp + Regulatory apps"),
    ("Program Crew · Iberdrola Corp Services",
     "Vanessa Colombo · Michael — Squad FI · PO · Budget"),
]

pc_y0 = kz_y1 + 30
pc_avail = forum_y1 - 90 - pc_y0
pc_h = (pc_avail - 5 * 16) // 6

# Vertical Change Management Crew band (purple, cutting through programs)
cmc_w = 90
cmc_x0 = vc_x0 + 380
cmc_x1 = cmc_x0 + cmc_w
cmc_y0 = pc_y0 - 8
cmc_y1 = pc_y0 + 6 * pc_h + 5 * 16 + 8
d.rounded_rectangle([cmc_x0, cmc_y0, cmc_x1, cmc_y1],
                    radius=22, fill=CHANGE_BG, outline=CHANGE_BR, width=3)

# Vertical text "Change Management Crew"
# Use a temporary image rotated 90deg
tmp = Image.new("RGBA", (cmc_y1 - cmc_y0 - 40, cmc_w), (0, 0, 0, 0))
td = ImageDraw.Draw(tmp)
td.text((20, cmc_w / 2), "Change Management Crew",
        fill=DARK, font=fnt(22, True), anchor="lm")
tmp = tmp.rotate(90, expand=True)
img.paste(tmp, (cmc_x0, cmc_y0 + 20), tmp)

# Draw program crews
for i, (pname, members) in enumerate(program_streams):
    py0 = pc_y0 + i * (pc_h + 16)
    py1 = py0 + pc_h
    d.rounded_rectangle([vc_x0 + 30, py0, vc_x1 - 30, py1],
                        radius=14, fill=PROG_BG, outline=PROG_BR, width=3)
    # Program name (left of CMC band)
    left(d, (vc_x0 + 50, py0 + pc_h / 2 - 12),
         pname, fnt(20, True), DARK)
    left(d, (vc_x0 + 50, py0 + pc_h / 2 + 14),
         members, fnt(13), GREY_TXT)
    # Persons (right of CMC band)
    for j in range(4):
        person_icon(d, cmc_x1 + 60 + j * 56, py0 + pc_h / 2, r=18)

# Description
left(d, (vc_x0 + 50, forum_y1 - 50),
     "A Value Chain Crew offers shared services to the other Crews in the Base.",
     fnt(15, italic=True), GREY_TXT)


# ── IT Projects Forum (right) ────────────────────────────────────────────
d.rounded_rectangle([itp_x0, forum_y0, itp_x1, forum_y1], radius=22,
                    fill=FORUM_BG, outline=FORUM_BR, width=4)
# Forum tag
d.rounded_rectangle([itp_x0 + 20, forum_y0 - 24, itp_x0 + 280, forum_y0 + 28],
                    radius=12, fill=WHITE, outline=FORUM_BR, width=2)
centered(d, ((itp_x0 + 20 + itp_x0 + 280) / 2, forum_y0 + 2),
         "IT Projects Forum", fnt(22, True), DARK)
person_icon(d, itp_x0 + 38, forum_y0 + 2, r=18)

# Capability Crew + Release Crew
cap_y0 = forum_y0 + 70
cap_y1 = cap_y0 + 600
d.rounded_rectangle([itp_x0 + 40, cap_y0, itp_x1 - 40, cap_y1],
                    radius=18, fill=CAP_BG, outline=CAP_BR, width=3)
centered(d, ((itp_x0 + itp_x1) / 2, cap_y0 + 30),
         "Capability Crew", fnt(24, True), DARK)
for i in range(4):
    person_icon(d, itp_x0 + 100 + i * 130, cap_y0 + 100, r=24)
left(d, (itp_x0 + 70, cap_y0 + 170),
     "Architecture Council:", fnt(18, True), DARK)
arch_lines = [
    "TBD Sr. Mgr SAP & CS Architecture (vacant)",
    "Corporate Tech Leads (2)",
    "Cross-stream architects",
    "Reviews scope · feasibility · dependencies",
]
for i, ln in enumerate(arch_lines):
    left(d, (itp_x0 + 80, cap_y0 + 210 + i * 30),
         "• " + ln, fnt(15), DARK)
left(d, (itp_x0 + 70, cap_y1 - 30),
     "A Capability Crew offers unique expertise to other Crews.",
     fnt(13, italic=True), GREY_TXT)

# Release Crew
rel_y0 = cap_y1 + 30
rel_y1 = forum_y1 - 60
d.rounded_rectangle([itp_x0 + 40, rel_y0, itp_x1 - 40, rel_y1],
                    radius=18, fill=CAP_BG, outline=CAP_BR, width=3)
centered(d, ((itp_x0 + itp_x1) / 2, rel_y0 + 30),
         "Release Crew", fnt(24, True), DARK)
for i in range(4):
    person_icon(d, itp_x0 + 100 + i * 130, rel_y0 + 100, r=24)
rel_lines = [
    "Platform Services — SWF · DevOps · App Sec",
    "SAP Basis (Nagi Reddy)",
    "Test Automation (TOSCA) · Data Ops",
    "Owners: Raquel · Donna · Piotr · Nagi",
]
for i, ln in enumerate(rel_lines):
    left(d, (itp_x0 + 80, rel_y0 + 150 + i * 28),
         "• " + ln, fnt(15), DARK)
left(d, (itp_x0 + 70, rel_y1 - 30),
     "Release Crew offers a regular cadence to the other Crews.",
     fnt(13, italic=True), GREY_TXT)


# ── Turf (label inside the Base, bottom) ─────────────────────────────────
turf_y = sab_y1 - 130
left(d, (sab_x0 + 110, turf_y),
     "The Turf is an area cultivated and protected by the same people.",
     fnt(20, italic=True), GREY_TXT)
# 3 person icons
for i in range(3):
    person_icon(d, sab_x0 + 1700 + i * 90, turf_y + 30, r=22)


# Base label below
centered(d, ((sab_x0 + sab_x1) / 2, sab_y1 + 50),
         "Digital Products — Strongly Aligned Base",
         fnt(34, True), DARK)
centered(d, ((sab_x0 + sab_x1) / 2, sab_y1 + 90),
         "A Strongly Aligned Base offers a cohesive set of products with many dependencies between each other.",
         fnt(20, italic=True), GREY_TXT)


# ── Segregated Base (right side — SW Delivery / Vendors) ─────────────────
seg_cx = 3800
seg_cy = 1850
seg_r  = 480
# Dashed circle
for ang in range(0, 360, 4):
    rad = math.radians(ang)
    rad2 = math.radians(ang + 2)
    x1 = seg_cx + seg_r * math.cos(rad)
    y1 = seg_cy + seg_r * math.sin(rad)
    x2 = seg_cx + seg_r * math.cos(rad2)
    y2 = seg_cy + seg_r * math.sin(rad2)
    d.line([x1, y1, x2, y2], fill=BASE_BR, width=4)

home_icon(d, seg_cx + 380, seg_cy - 380, r=36)

# People inside (vendor pool)
vendor_persons = [
    "Accenture", "SWF / Software Factory",
    "HCL · Nuance", "Infosys",
    "Smart Metering vendor", "SAP partners",
]
# Arrange in 2 rows of 3
ncols = 3
for i, name in enumerate(vendor_persons):
    row = i // ncols
    col = i % ncols
    cx = seg_cx - 200 + col * 200
    cy = seg_cy - 60 + row * 120
    person_icon(d, cx, cy, r=32)
    centered(d, (cx, cy + 56), name, fnt(15), DARK)

# Label
centered(d, (seg_cx, seg_cy + seg_r - 30),
         "SW Delivery Base", fnt(30, True), DARK)
centered(d, (seg_cx, seg_cy + seg_r + 4),
         "(Fully Segregated Base)", fnt(18, italic=True), GREY_TXT)

# Description text below segregated base
dsc_x0 = seg_cx - 380
dsc_y0 = seg_cy + seg_r + 50
left_lines = [
    "A Fully Segregated Base offers multiple",
    "independent and potentially competing",
    "products or services.",
]
for i, ln in enumerate(left_lines):
    centered(d, (seg_cx, dsc_y0 + i * 26),
             ln, fnt(16, italic=True), GREY_TXT)


# ── Connecting arrows (Strongly Aligned ↔ Segregated) ────────────────────
# Arrow from SAB to SW Delivery
arrow_y = 1700
d.line([sab_x1 + 10, arrow_y, seg_cx - seg_r - 10, arrow_y],
       fill=DARK, width=3)
# Arrowhead
d.polygon([(seg_cx - seg_r - 10, arrow_y),
           (seg_cx - seg_r - 30, arrow_y - 12),
           (seg_cx - seg_r - 30, arrow_y + 12)],
          fill=DARK)


# ── Footer ───────────────────────────────────────────────────────────────
d.text((80, H - 60),
       "Source: Unfix framework — Bases · Forums · Crews · Turf · Avangrid Digital Products TOM · Apr 2026",
       font=fnt(20), fill=GREY_TXT)


OUT.mkdir(parents=True, exist_ok=True)
out_file = OUT / "DigitalProducts_Unfix_Bases.png"
img.save(out_file, "PNG", dpi=(150, 150))

# preview
ratio = 1600 / W
preview = img.resize((1600, int(H * ratio)), Image.LANCZOS)
preview.save(out_file.with_name(out_file.stem + "_preview.png"), "PNG")
print(f"✓ Saved: {out_file}")
print(f"  Resolution: {W}×{H}")
