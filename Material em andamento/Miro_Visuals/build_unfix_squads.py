"""
Digital Products — Unfix model with PRODUCT SQUADS visible inside each Value Stream Crew.
Squads tagged (KB) = exists today; (NEW) = proposed for to-be model.
"""
from PIL import Image, ImageDraw, ImageFont
from pathlib import Path
import math

OUT = Path("/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Miro_Visuals")

# Palette
BASE_BG    = (231, 232, 234)
BASE_BR    = (170, 175, 185)
FORUM_BR   = ( 90, 145,  85)
FORUM_BG   = (252, 254, 252)
GOV        = (251, 195, 145)
GOV_BR     = (210, 145,  80)
VS         = (215, 200, 245)
VS_BR      = (160, 140, 210)
EXP        = (188, 220, 188)
EXP_BR     = (110, 160, 110)
CAP        = (188, 220, 188)
CAP_BR     = (110, 160, 110)
PLAT       = (188, 220, 188)
PLAT_BR    = (110, 160, 110)
PART       = (250, 226, 145)
PART_BR    = (200, 175,  80)
FAC_BAND   = (190, 195, 230)
FAC_BR     = (130, 140, 200)
SQUAD_BG   = (255, 255, 255)
SQUAD_BR   = (160, 140, 210)
KB_TAG     = ( 76, 140,  63)
NEW_TAG    = (232, 139,  47)
WHITE      = (255, 255, 255)
DARK       = ( 30,  40,  55)
GREY       = ( 90, 100, 115)
PERSON_F   = (210, 215, 225)
PERSON_BR  = (150, 160, 175)


W, H = 5400, 3600


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


def person(d, cx, cy, r=18, fill=PERSON_F, br=PERSON_BR):
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=fill, outline=br, width=2)
    head_r = r * 0.38
    d.ellipse([cx - head_r, cy - r * 0.55, cx + head_r, cy - r * 0.55 + 2 * head_r],
              fill=br)
    d.pieslice([cx - r * 0.7, cy - r * 0.05, cx + r * 0.7, cy + r * 1.4],
               start=180, end=360, fill=br)


def home_dashed(d, cx, cy, r=44, color=DARK):
    for ang in range(0, 360, 12):
        rad = math.radians(ang)
        x1 = cx + (r + 8) * math.cos(rad); y1 = cy + (r + 8) * math.sin(rad)
        x2 = cx + (r + 14) * math.cos(rad); y2 = cy + (r + 14) * math.sin(rad)
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
    for ang in range(0, 360, 4):
        rad = math.radians(ang); rad2 = math.radians(ang + 2)
        x1 = cx + r * math.cos(rad); y1 = cy + r * math.sin(rad)
        x2 = cx + r * math.cos(rad2); y2 = cy + r * math.sin(rad2)
        d.line([x1, y1, x2, y2], fill=color, width=width)


def status_tag(d, x, y, text, kb=True):
    """Small KB / NEW tag pill."""
    color = KB_TAG if kb else NEW_TAG
    txt = "KB" if kb else "NEW"
    pad = 8
    tw = 56
    th = 24
    d.rounded_rectangle([x, y, x + tw, y + th], radius=6, fill=color)
    centered(d, (x + tw / 2, y + th / 2), txt, fnt(13, True), WHITE)


def squad_card(d, x0, y0, x1, y1, name, modules, kb, member_count=3):
    d.rounded_rectangle([x0, y0, x1, y1], radius=12, fill=SQUAD_BG,
                        outline=SQUAD_BR, width=2)
    # Header band
    d.rounded_rectangle([x0, y0, x1, y0 + 38], radius=12, fill=VS_BR)
    d.rectangle([x0, y0 + 26, x1, y0 + 38], fill=VS_BR)
    left(d, (x0 + 12, y0 + 19), name, fnt(15, True), WHITE)
    # Status tag (right of header)
    status_tag(d, x1 - 70, y0 + 7, "tag", kb=kb)
    # Modules
    left(d, (x0 + 12, y0 + 56), "Modules: " + modules, fnt(11), GREY)
    # Persons
    for i in range(member_count):
        person(d, x0 + 22 + i * 32, y1 - 22, r=14)
    # Squad type label
    left(d, (x0 + 12, y1 - 10), "squad", fnt(10, italic=True), GREY)


img = Image.new("RGB", (W, H), WHITE)
d = ImageDraw.Draw(img)


# ── Header ───────────────────────────────────────────────────────────────
d.rectangle([0, 0, W, 130], fill=DARK)
d.text((80, 28), "DIGITAL PRODUCTS — unFIX MODEL with PRODUCT SQUADS",
       font=fnt(48, True), fill=WHITE)
d.text((80, 88),
       "Each Value Stream Crew breaks down into product squads · Tag (KB) = exists today · (NEW) = proposed for to-be",
       font=fnt(20), fill=(199, 211, 228))


# ── Home + Strongly Aligned Base ─────────────────────────────────────────
home_dashed(d, 2200, 250, r=42)

sab_x0, sab_y0 = 200, 380
sab_x1, sab_y1 = 4400, 3050
d.rounded_rectangle([sab_x0, sab_y0, sab_x1, sab_y1], radius=44,
                    fill=BASE_BG, outline=BASE_BR, width=4)


# ── Governance Crew ──────────────────────────────────────────────────────
gc_x0 = 900; gc_x1 = 3400
gc_y0 = 320; gc_y1 = 580
d.rounded_rectangle([gc_x0, gc_y0, gc_x1, gc_y1], radius=24,
                    fill=GOV, outline=GOV_BR, width=4)
tag_x = gc_x1 - 320; tag_y = gc_y0 - 30
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
     "The management team of the Base.",
     fnt(14, italic=True), GREY)


# ── Forums ───────────────────────────────────────────────────────────────
forum_y0 = 720
forum_y1 = 2820

# Customer Journey Forum (left, narrower)
cj_x0 = 280; cj_x1 = 880
d.rounded_rectangle([cj_x0, forum_y0, cj_x1, forum_y1], radius=22,
                    fill=FORUM_BG, outline=FORUM_BR, width=4)
tag_w = 360
d.rounded_rectangle([cj_x0 + 24, forum_y0 - 28, cj_x0 + 24 + tag_w, forum_y0 + 28],
                    radius=14, fill=WHITE, outline=FORUM_BR, width=3)
person(d, cj_x0 + 54, forum_y0, r=18)
d.text((cj_x0 + 89, forum_y0), "Customer Journey Forum",
       font=fnt(20, True), fill=DARK, anchor="lm")

# Product Forum (center, wide)
pf_x0 = 920; pf_x1 = 3450
d.rounded_rectangle([pf_x0, forum_y0, pf_x1, forum_y1], radius=22,
                    fill=FORUM_BG, outline=FORUM_BR, width=4)
d.rounded_rectangle([pf_x0 + 24, forum_y0 - 28, pf_x0 + 24 + tag_w, forum_y0 + 28],
                    radius=14, fill=WHITE, outline=FORUM_BR, width=3)
person(d, pf_x0 + 54, forum_y0, r=18)
d.text((pf_x0 + 89, forum_y0), "Product Forum",
       font=fnt(20, True), fill=DARK, anchor="lm")

# Organizational Forum (small, top-right)
of_x0 = 3350; of_x1 = 4150
of_y0 = forum_y0 - 60; of_y1 = forum_y0 + 250
d.rounded_rectangle([of_x0, of_y0, of_x1, of_y1], radius=22,
                    fill=FORUM_BG, outline=FORUM_BR, width=4)
d.rounded_rectangle([of_x0 + 24, of_y0 - 28, of_x0 + 24 + tag_w, of_y0 + 28],
                    radius=14, fill=WHITE, outline=FORUM_BR, width=3)
person(d, of_x0 + 54, of_y0, r=18)
d.text((of_x0 + 89, of_y0), "Organizational Forum",
       font=fnt(18, True), fill=DARK, anchor="lm")
# Partnership Crew nested
part_x0 = of_x0 + 30; part_x1 = of_x1 - 30
part_y0 = of_y0 + 80; part_y1 = of_y1 - 24
d.rounded_rectangle([part_x0, part_y0, part_x1, part_y1], radius=14,
                    fill=PART, outline=PART_BR, width=3)
left(d, (part_x0 + 16, part_y0 + 24),
     "Partnership Crew", fnt(18, True), DARK)
left(d, (part_x0 + 16, part_y0 + 56),
     "Vendors · Accenture · SWF", fnt(12), DARK)
left(d, (part_x0 + 16, part_y0 + 78),
     "Iberdrola corporate liaison", fnt(12), DARK)
for j in range(3):
    person(d, part_x1 - 30 - j * 36, part_y0 + 100, r=14)

# Technological Forum (right, overlapping)
tf_x0 = 3370; tf_x1 = 4350
tf_y0 = of_y1 - 30; tf_y1 = forum_y1
d.rounded_rectangle([tf_x0, tf_y0, tf_x1, tf_y1], radius=22,
                    fill=FORUM_BG, outline=FORUM_BR, width=4)
d.rounded_rectangle([tf_x0 + 24, tf_y0 - 28, tf_x0 + 24 + tag_w, tf_y0 + 28],
                    radius=14, fill=WHITE, outline=FORUM_BR, width=3)
person(d, tf_x0 + 54, tf_y0, r=18)
d.text((tf_x0 + 89, tf_y0), "Technological Forum",
       font=fnt(18, True), fill=DARK, anchor="lm")


# ── Crews inside Customer Journey Forum ──────────────────────────────────
exp_x0 = cj_x0 + 30; exp_x1 = cj_x1 - 30
exp_y0 = forum_y0 + 80; exp_y1 = exp_y0 + 460
d.rounded_rectangle([exp_x0, exp_y0, exp_x1, exp_y1], radius=18,
                    fill=EXP, outline=EXP_BR, width=3)
centered(d, ((exp_x0 + exp_x1) / 2, exp_y0 + 28),
         "Experience Crew", fnt(20, True), DARK)
for j in range(4):
    person(d, exp_x0 + 50 + j * 50, exp_y0 + 90, r=18)
left(d, (exp_x0 + 18, exp_y0 + 140),
     "UX & customer journey:", fnt(13, True), DARK)
exp_lines = ["UX Designer (NOM v1.5)", "Heather LaJoie (CSP)",
             "CX Researcher (CoE)", "End-to-end journey",
             "Holistic experience"]
for j, ln in enumerate(exp_lines):
    left(d, (exp_x0 + 18, exp_y0 + 170 + j * 26),
         "• " + ln, fnt(12), DARK)

cap_y0 = exp_y1 + 20; cap_y1 = forum_y1 - 60
d.rounded_rectangle([exp_x0, cap_y0, exp_x1, cap_y1], radius=18,
                    fill=CAP, outline=CAP_BR, width=3)
centered(d, ((exp_x0 + exp_x1) / 2, cap_y0 + 28),
         "Capability Crew", fnt(20, True), DARK)
for j in range(4):
    person(d, exp_x0 + 50 + j * 50, cap_y0 + 90, r=18)
left(d, (exp_x0 + 18, cap_y0 + 140),
     "Architecture Council:", fnt(13, True), DARK)
cap_lines = ["TBD Sr.M. Architecture", "Corporate Tech Leads (2)",
             "Cross-stream architects", "Reviews scope · feasibility"]
for j, ln in enumerate(cap_lines):
    left(d, (exp_x0 + 18, cap_y0 + 170 + j * 26),
         "• " + ln, fnt(12), DARK)


# ── Value Stream Crews with PRODUCT SQUADS inside (Product Forum) ────────
vs_streams = [
    {
        "name": "Value Stream Crew · M2C",
        "branch": "CS",
        "pm": "Product Manager NEW · Iram = PJM today",
        "squads": [
            ("Smart Metering", "SAP Metering · vendor", False),
            ("Billing & FICA", "FICA · ECC · StreamServe", False),
            ("Collections & Statements", "FICA Coll. · SOVOS · Precisely", False),
        ],
    },
    {
        "name": "Value Stream Crew · Customer Care & CCC",
        "branch": "CS",
        "pm": "Product Manager NEW · Sr.M. vacant",
        "squads": [
            ("CRM Platform", "SAP CRM (Front)", False),
            ("Contact Center", "IVR · Agent Desktop", False),
            ("Self-Service & Feedback", "Portal · Medallia · Blue Prism", False),
        ],
    },
    {
        "name": "Value Stream Crew · Field Operations",
        "branch": "IT",
        "pm": "Product Manager NEW · Andrew = Champion today",
        "squads": [
            ("Mobile Field Worker", "Android/tablets · 70+ apps", False),
            ("Doc Viewer & Storm", "Doc Viewer · Storm toolkit", False),
        ],
    },
    {
        "name": "Value Stream Crew · Engineering & Asset Mgmt",
        "branch": "IT",
        "pm": "Product Manager NEW · George = Sr.M. today",
        "squads": [
            ("GIS Platform", "ArcFM · GIS web", False),
            ("Network Asset & Eng. Drawing", "OSG · ProjectWise · Microstation", False),
            ("Spatial Data Services", "SAP HANA · replication", False),
        ],
    },
    {
        "name": "Value Stream Crew · Corporate & Regulatory",
        "branch": "IT",
        "pm": "Product Manager NEW · Sarath = Sr.M. + 5-6 PJMs",
        "squads": [
            ("HR & Employee Portal", "Corporate apps", False),
            ("Regulatory & Compliance", "Compliance toolkit", False),
        ],
    },
    {
        "name": "Value Stream Crew · Iberdrola Corp Services",
        "branch": "IT",
        "pm": "Product Manager NEW · Vanessa = Sr.M. today",
        "squads": [
            ("Finance (FI)", "SAP FI", False),
            ("Procurement & Budget", "SAP PO · Budget control", False),
        ],
    },
]

BRANCH_COLOR = {
    "CS": ( 70, 130, 220),    # blue badge for Customer Services Apps
    "IT": (220, 120,  60),    # orange badge for IT Business Apps
}

vs_block_y0 = forum_y0 + 80
vs_block_y1 = forum_y1 - 80
vs_avail_h  = vs_block_y1 - vs_block_y0
vs_n = len(vs_streams)
vs_gap = 14
vs_h = (vs_avail_h - (vs_n - 1) * vs_gap) // vs_n

vs_x0_card = pf_x0 + 30
vs_x1_card = pf_x1 - 30

# Value stream bar layout: header (left) + squads (middle) + persons (right)
header_w = 360
persons_w = 100
gap_inner = 18

for i, s in enumerate(vs_streams):
    y0 = vs_block_y0 + i * (vs_h + vs_gap)
    y1 = y0 + vs_h
    # Bar background
    d.rounded_rectangle([vs_x0_card, y0, vs_x1_card, y1], radius=14,
                        fill=VS, outline=VS_BR, width=3)
    # Branch badge (top-left of bar)
    badge_color = BRANCH_COLOR[s["branch"]]
    badge_text = "CS Apps" if s["branch"] == "CS" else "IT Apps"
    badge_w = 96
    badge_h = 26
    d.rounded_rectangle([vs_x0_card + 12, y0 + 10,
                         vs_x0_card + 12 + badge_w, y0 + 10 + badge_h],
                        radius=6, fill=badge_color)
    centered(d, (vs_x0_card + 12 + badge_w / 2, y0 + 10 + badge_h / 2),
             badge_text, fnt(13, True), WHITE)
    # Header (left)
    left(d, (vs_x0_card + 18, y0 + 56), s["name"], fnt(17, True), DARK)
    left(d, (vs_x0_card + 18, y0 + 84), "PM: " + s["pm"], fnt(12), DARK)

    # Squad cards
    squads = s["squads"]
    avail_w = (vs_x1_card - vs_x0_card) - header_w - persons_w - gap_inner * 2
    sq_n = len(squads)
    sq_gap = 12
    sq_w = (avail_w - (sq_n - 1) * sq_gap) // sq_n
    sq_y0 = y0 + 16
    sq_y1 = y1 - 16
    for j, (sname, smods, skb) in enumerate(squads):
        sx0 = vs_x0_card + header_w + gap_inner + j * (sq_w + sq_gap)
        sx1 = sx0 + sq_w
        squad_card(d, sx0, sq_y0, sx1, sq_y1, sname, smods, skb,
                   member_count=3)

    # Persons (right edge — crew leader)
    for k in range(3):
        person(d, vs_x1_card - 30 - k * 30, y0 + vs_h / 2, r=12)


# ── Facilitation Crew vertical band cutting across Value Streams ─────────
fac_x0 = pf_x0 + 360 + 9   # right after the header column
fac_x1 = fac_x0 + 60
fac_y0 = vs_block_y0 - 8
fac_y1 = vs_block_y1 + 8
d.rounded_rectangle([fac_x0, fac_y0, fac_x1, fac_y1], radius=18,
                    fill=FAC_BAND, outline=FAC_BR, width=3)
tmp = Image.new("RGBA", (fac_y1 - fac_y0 - 60, fac_x1 - fac_x0), (0, 0, 0, 0))
td = ImageDraw.Draw(tmp)
td.text((30, (fac_x1 - fac_x0) / 2), "Facilitation Crew",
        fill=DARK, font=fnt(18, True), anchor="lm")
tmp = tmp.rotate(90, expand=True)
img.paste(tmp, (fac_x0, fac_y0 + 30), tmp)


# ── Platform Crew inside Technological Forum ─────────────────────────────
plat_x0 = tf_x0 + 30; plat_x1 = tf_x1 - 30
plat_y0 = tf_y0 + 90; plat_y1 = tf_y1 - 60
d.rounded_rectangle([plat_x0, plat_y0, plat_x1, plat_y1], radius=18,
                    fill=PLAT, outline=PLAT_BR, width=3)
centered(d, ((plat_x0 + plat_x1) / 2, plat_y0 + 28),
         "Platform Crew", fnt(22, True), DARK)
for j in range(5):
    person(d, plat_x0 + 40 + j * 50, plat_y0 + 90, r=18)
left(d, (plat_x0 + 18, plat_y0 + 140),
     "Shared services to all Crews:", fnt(13, True), DARK)
plat_lines = ["SWF & Tools — Raquel · Donna",
              "DevOps — Piotr",
              "App Security — Piotr",
              "SAP Basis — Nagi Reddy",
              "Test Automation (TOSCA)",
              "Data Ops (Gold Client · 8760)",
              "Architecture · APIs · envs"]
for j, ln in enumerate(plat_lines):
    left(d, (plat_x0 + 18, plat_y0 + 170 + j * 26),
         "• " + ln, fnt(12), DARK)


# ── Turf strip ───────────────────────────────────────────────────────────
turf_y = forum_y1 + 60
left(d, (sab_x0 + 60, turf_y),
     "The Turf is an area cultivated and protected by the same people.",
     fnt(18, italic=True), GREY)
for i in range(3):
    person(d, sab_x0 + 2400 + i * 80, turf_y + 18, r=22)


# ── Base label ───────────────────────────────────────────────────────────
centered(d, ((sab_x0 + sab_x1) / 2, sab_y1 + 50),
         "Strongly Aligned Base — Digital Products",
         fnt(28, True), DARK)
centered(d, ((sab_x0 + sab_x1) / 2, sab_y1 + 86),
         "6 Value Stream Crews · ~14 product squads · Multiple value streams forming a cohesive product set with dependencies.",
         fnt(16, italic=True), GREY)


# ── Segregated Base ──────────────────────────────────────────────────────
seg_cx = 4860; seg_cy = 1500; seg_r = 460
home_dashed(d, seg_cx + 320, seg_cy - 320, r=32)
dashed_circle(d, seg_cx, seg_cy, seg_r, color=BASE_BR, width=4)
vendors = [("Accenture", "Global SAP + Corp"),
           ("SWF / Software Factory", "All teams"),
           ("HCL · Nuance", "Customer Service"),
           ("Infosys", "System support"),
           ("Smart Metering vendor", "100% delivery"),
           ("SAP partners", "Iberdrola side")]
ncols = 3
for i, (nm, role) in enumerate(vendors):
    row = i // ncols; col = i % ncols
    cx = seg_cx - 230 + col * 230
    cy = seg_cy - 100 + row * 160
    person(d, cx, cy, r=30, fill=PERSON_F)
    centered(d, (cx, cy + 54), nm, fnt(13, True), DARK)
    centered(d, (cx, cy + 74), role, fnt(11), GREY)
centered(d, (seg_cx, seg_cy + seg_r - 50),
         "SW Delivery Base", fnt(26, True), DARK)
centered(d, (seg_cx, seg_cy + seg_r - 14),
         "(Fully Segregated Base)", fnt(15, italic=True), GREY)
centered(d, (seg_cx, seg_cy + seg_r + 30),
         "External delivery — independent governance, vendors operate outside Avangrid Base",
         fnt(13, italic=True), GREY)


# ── Legend for KB / NEW tags ─────────────────────────────────────────────
leg_x = 80; leg_y = H - 120
d.text((leg_x, leg_y), "Legend:", font=fnt(18, True), fill=DARK)
status_tag(d, leg_x + 110, leg_y - 4, "tag", kb=False)
d.text((leg_x + 180, leg_y),
       "= product squad PROPOSED for to-be model (no Product Managers exist today; only PJMs / Sr. Managers / Champions)",
       font=fnt(14), fill=DARK)


# ── Footer ───────────────────────────────────────────────────────────────
d.text((80, H - 50),
       "Source: unfix.com · Avangrid Digital Products TOM proposal · Apr 2026 · ALL 14 product squads are NEW · No Product Managers exist today (only Project Managers, Sr. Managers, Champions)",
       font=fnt(15), fill=GREY)


OUT.mkdir(parents=True, exist_ok=True)
out_file = OUT / "DigitalProducts_Unfix_Squads.png"
img.save(out_file, "PNG", dpi=(150, 150))

ratio = 1700 / W
preview = img.resize((1700, int(H * ratio)), Image.LANCZOS)
preview.save(out_file.with_name(out_file.stem + "_preview.png"), "PNG")
print(f"✓ Saved: {out_file}")
print(f"  Resolution: {W}×{H}")
