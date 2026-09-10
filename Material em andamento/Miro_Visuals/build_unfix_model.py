"""
Digital Products operating model — Unfix framework visualization.
Following the user's reference image (Unfix crew-card template + Example 1 layout).
"""
from PIL import Image, ImageDraw, ImageFont
from pathlib import Path

OUT = Path("/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Miro_Visuals")

# ── Unfix color palette (matched to user's template) ──────────────────────
VS         = (245, 193,  58)   # Value Stream — yellow
VS_DARK    = (200, 153,  20)
GOV        = ( 74, 136, 216)   # Governance — blue
GOV_DARK   = ( 35,  87, 165)
PLAT       = (111, 190, 111)   # Platform — green
PLAT_DARK  = ( 60, 140,  60)
CAP        = (216,  85,  81)   # Capability — red
CAP_DARK   = (160,  50,  46)
EXP        = (229, 143, 181)   # Experience — pink
EXP_DARK   = (180,  90, 130)
PART       = (182, 213,  78)   # Partnership — lime
PART_DARK  = (140, 175,  40)
FAC        = (122, 102, 184)   # Facilitation — purple
FAC_DARK   = ( 80,  60, 145)
FORUM_REG  = (180, 220, 180)
FORUM_FUN  = (240, 220, 130)
FORUM_TEC  = (200, 180, 230)

GREY       = ( 75,  85,  99)
DARK_TXT   = ( 30,  40,  55)
LIGHT      = (245, 245, 248)
WHITE      = (255, 255, 255)
BORDER     = (200, 208, 220)
GUILD_BG   = (250, 250, 252)
GUILD_BR   = (170, 178, 190)

# ── Canvas ───────────────────────────────────────────────────────────────
W, H = 4200, 3000


def fnt(size, bold=False):
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


def wrap_text(text, max_chars):
    words = text.split()
    lines, cur = [], ""
    for w_ in words:
        if len((cur + " " + w_).strip()) <= max_chars:
            cur = (cur + " " + w_).strip()
        else:
            if cur:
                lines.append(cur)
            cur = w_
    if cur:
        lines.append(cur)
    return lines


img = Image.new("RGB", (W, H), WHITE)
d = ImageDraw.Draw(img)


# ── Page header ──────────────────────────────────────────────────────────
d.rectangle([0, 0, W, 180], fill=DARK_TXT)
d.text((80, 40), "DIGITAL PRODUCTS — OPERATING MODEL", font=fnt(72, True), fill=WHITE)
d.text((80, 120), "Unfix-style topology · 6 Value Stream Crews · 1 Platform Crew · 4 supporting crews · 3 Forums",
       font=fnt(28), fill=(199, 211, 228))


# ── Governance Crew (top, full width) ────────────────────────────────────
gov_y0 = 220
gov_y1 = gov_y0 + 200
d.rounded_rectangle([80, gov_y0, W - 80, gov_y1], radius=24, fill=GOV)
d.text((110, gov_y0 + 30), "GOVERNANCE CREW", font=fnt(40, True), fill=WHITE)
d.text((110, gov_y0 + 85), "Management & strategic alignment for Digital Products",
       font=fnt(22), fill=(220, 230, 250))

# Members of governance
gov_members = [
    "Vitor Cruz · VP IT Business Solutions",
    "Carlos Garbin · Director IT Business Apps",
    "Kumar Jaladanki · Director Customer Services Apps",
    "Heather LaJoie · Director Product Mgmt (business)",
    "Portfolio Manager · TBD",
    "Travis · PMO Lead",
]
mx = 110
for i, m in enumerate(gov_members):
    col = i % 3
    row = i // 3
    bx = 110 + col * 1330
    by = gov_y0 + 130 + row * 36
    d.text((bx, by), "• " + m, font=fnt(22), fill=WHITE)


# ── Two Guilds (Customer Services + IT Business Apps) ────────────────────
guild_y0 = gov_y1 + 60
guild_h  = 1620
guild_x_split = W // 2
guild_pad = 30

# Customer Services guild
cs_x0 = 80
cs_x1 = guild_x_split - 20
d.rounded_rectangle([cs_x0, guild_y0, cs_x1, guild_y0 + guild_h],
                    radius=22, fill=GUILD_BG, outline=GUILD_BR, width=4)
d.text((cs_x0 + 30, guild_y0 + 25), "CUSTOMER SERVICES GUILD",
       font=fnt(34, True), fill=DARK_TXT)
d.text((cs_x0 + 30, guild_y0 + 75), "Kumar Jaladanki · ~24 FTE · 5 sub-teams",
       font=fnt(22), fill=GREY)

# IT Business Apps guild
it_x0 = guild_x_split + 20
it_x1 = W - 80
d.rounded_rectangle([it_x0, guild_y0, it_x1, guild_y0 + guild_h],
                    radius=22, fill=GUILD_BG, outline=GUILD_BR, width=4)
d.text((it_x0 + 30, guild_y0 + 25), "IT BUSINESS APPS GUILD",
       font=fnt(34, True), fill=DARK_TXT)
d.text((it_x0 + 30, guild_y0 + 75), "Carlos Garbin · ~26 FTE · 4 sub-teams",
       font=fnt(22), fill=GREY)


# ── Value Stream Crews (yellow cards inside the guilds) ──────────────────
def draw_vs_crew(x0, y0, x1, y1, name, pm, members, tl, modules):
    """Draw a value stream crew in Unfix style."""
    d.rounded_rectangle([x0, y0, x1, y1], radius=18, fill=VS,
                        outline=VS_DARK, width=3)
    # Header band
    d.rounded_rectangle([x0, y0, x1, y0 + 70], radius=18, fill=VS_DARK)
    d.rectangle([x0, y0 + 50, x1, y0 + 70], fill=VS_DARK)
    d.text((x0 + 24, y0 + 14), name, font=fnt(28, True), fill=WHITE)

    # Body rows: PM | members | Tech Lead
    body_y = y0 + 90
    # PM block
    pm_w = 180
    d.rounded_rectangle([x0 + 16, body_y, x0 + 16 + pm_w, y1 - 16],
                        radius=10, fill=FAC, outline=FAC_DARK, width=2)
    centered(d, (x0 + 16 + pm_w / 2, body_y + 28), "PM", fnt(20, True), WHITE)
    pm_lines = wrap_text(pm, 14)[:3]
    for i, ln in enumerate(pm_lines):
        centered(d, (x0 + 16 + pm_w / 2, body_y + 70 + i * 22),
                 ln, fnt(15), WHITE)

    # Members block
    members_x0 = x0 + 16 + pm_w + 12
    members_x1 = x1 - 16 - pm_w - 12
    d.rounded_rectangle([members_x0, body_y, members_x1, y1 - 16],
                        radius=10, fill=WHITE, outline=VS_DARK, width=2)
    d.text((members_x0 + 16, body_y + 12), "Members / Sub-team",
           font=fnt(16, True), fill=GREY)
    for i, m in enumerate(members[:5]):
        d.text((members_x0 + 16, body_y + 42 + i * 26),
               "• " + m, font=fnt(17), fill=DARK_TXT)
    # Modules note
    if modules:
        mod_y = body_y + 42 + min(len(members), 5) * 26 + 16
        if mod_y < y1 - 50:
            d.text((members_x0 + 16, mod_y),
                   "Modules: " + ", ".join(modules[:4]),
                   font=fnt(15), fill=GREY)

    # TL block
    tl_w = 180
    d.rounded_rectangle([x1 - 16 - tl_w, body_y, x1 - 16, y1 - 16],
                        radius=10, fill=CAP, outline=CAP_DARK, width=2)
    centered(d, (x1 - 16 - tl_w / 2, body_y + 28), "TECH LEAD", fnt(18, True), WHITE)
    tl_lines = wrap_text(tl, 14)[:3]
    for i, ln in enumerate(tl_lines):
        centered(d, (x1 - 16 - tl_w / 2, body_y + 70 + i * 22),
                 ln, fnt(15), WHITE)


# CS guild has 2 value stream crews
cs_streams = [
    {
        "name": "M2C — METER-TO-CASH",
        "pm": "TBD (Iram available)",
        "members": [
            "Vincent Lam (Sr. Mgr Device & Metering)",
            "Iram Chowdhury (Smart Metering PM)",
            "Ivelisse Velazquez (FICA/Billing)",
            "Squad Metering · Squad Billing",
        ],
        "tl": "Tech Lead TBD",
        "modules": ["SAP Metering", "FICA", "ECC", "StreamServe"],
    },
    {
        "name": "CUSTOMER CARE & CCC",
        "pm": "Vacant Sr. Mgr",
        "members": [
            "TBD Sr. Mgr CRM/CCC (vacant)",
            "Heather LaJoie (business)",
            "Squad CRM · Squad Self-Service",
            "Squad Contact Center",
        ],
        "tl": "Vacant",
        "modules": ["SAP CRM", "Medallia", "Blue Prism"],
    },
]

cs_card_x0 = cs_x0 + 30
cs_card_x1 = cs_x1 - 30
cs_card_y0 = guild_y0 + 130
cs_avail_h = guild_h - 160
cs_card_h  = (cs_avail_h - 30) // 2

for i, s in enumerate(cs_streams):
    y0 = cs_card_y0 + i * (cs_card_h + 30)
    y1 = y0 + cs_card_h
    draw_vs_crew(cs_card_x0, y0, cs_card_x1, y1,
                 s["name"], s["pm"], s["members"], s["tl"], s["modules"])


# IT guild has 4 value stream crews
it_streams = [
    {
        "name": "FIELD OPERATIONS",
        "pm": "Andrew McGibbon",
        "members": [
            "Bruno Moreto (Sr. Mgr E&G Ops)",
            "Squad Mobility · Squad Doc Viewer",
            "Squad Storm & Workforce Mgmt",
        ],
        "tl": "TBD",
        "modules": ["Android/tablets", "ProjectWise", "Microstation"],
    },
    {
        "name": "ENGINEERING & ASSET MGMT",
        "pm": "George Porto (dual role)",
        "members": [
            "George Porto (Sr. Mgr GIS & OSG)",
            "Squad ArcFM · Squad SAP HANA",
            "Squad GIS Web · Squad Network Topology",
        ],
        "tl": "TBD",
        "modules": ["SAP HANA", "ArcFM", "GIS Web", "OSG"],
    },
    {
        "name": "CORPORATE & REGULATORY",
        "pm": "Sarath (decision pending)",
        "members": [
            "Sarath Garikapati + Amit",
            "5-6 Project Managers",
            "2 Tech Leads",
        ],
        "tl": "Internal (2 leads)",
        "modules": ["Corporate apps", "Network apps", "SWF tooling"],
    },
    {
        "name": "IBERDROLA CORPORATE SVC",
        "pm": "Vanessa Colombo",
        "members": [
            "Vanessa Colombo (Sr. Mgr Global SAP)",
            "Michael Dubois",
            "Squad FI · Squad PO · Squad Budget",
        ],
        "tl": "TBD",
        "modules": ["SAP FI", "SAP PO", "Budget"],
    },
]

it_card_x0 = it_x0 + 30
it_card_x1 = it_x1 - 30
it_card_y0 = guild_y0 + 130
it_avail_h = guild_h - 160
it_card_h  = (it_avail_h - 90) // 4

for i, s in enumerate(it_streams):
    y0 = it_card_y0 + i * (it_card_h + 30)
    y1 = y0 + it_card_h
    draw_vs_crew(it_card_x0, y0, it_card_x1, y1,
                 s["name"], s["pm"], s["members"], s["tl"], s["modules"])


# ── Platform Crew (green, full width, below guilds) ───────────────────────
plat_y0 = guild_y0 + guild_h + 50
plat_y1 = plat_y0 + 220
d.rounded_rectangle([80, plat_y0, W - 80, plat_y1], radius=24, fill=PLAT)
d.text((110, plat_y0 + 25), "PLATFORM CREW",
       font=fnt(40, True), fill=WHITE)
d.text((110, plat_y0 + 80), "Shared services consumed by all Value Stream Crews — SLA-driven, no end-user demand of its own",
       font=fnt(20), fill=(220, 240, 220))

plat_caps = [
    ("SWF & Tools", "Raquel Soto · Donna Henry"),
    ("DevOps", "Piotr Wielogorski"),
    ("App Security", "Piotr Wielogorski"),
    ("SAP Basis", "Nagi Reddy"),
    ("Test Automation", "Tricentis TOSCA"),
    ("Data Ops", "Gold Client · 8760 EDI"),
]

cap_w = (W - 160 - 5 * 14) // 6
for i, (cname, owner) in enumerate(plat_caps):
    x0 = 110 + i * (cap_w + 14)
    x1 = x0 + cap_w
    cy0 = plat_y0 + 130
    cy1 = plat_y1 - 20
    d.rounded_rectangle([x0, cy0, x1, cy1], radius=12, fill=PLAT_DARK,
                        outline=PLAT_DARK, width=0)
    centered(d, ((x0 + x1) / 2, cy0 + 22), cname, fnt(20, True), WHITE)
    centered(d, ((x0 + x1) / 2, cy0 + 50), owner, fnt(15), (220, 240, 220))


# ── Supporting Crews row (Capability · Experience · Partnership · Facilitation) ──
sup_y0 = plat_y1 + 50
sup_y1 = sup_y0 + 240

sup_crews = [
    (CAP, CAP_DARK,  "CAPABILITY CREW",  "Architecture Council",
     ["TBD Sr. Mgr SAP & CS Architecture (vacant)",
      "Corporate Tech Leads (2)",
      "Cross-stream architects",
      "Reviews scope, feasibility, dependencies"]),
    (EXP, EXP_DARK,  "EXPERIENCE CREW",  "UX & Customer Insight",
     ["UX Designer (NOM v1.5 role)",
      "Heather LaJoie (Customer Service Products)",
      "CX Researcher (from Digital CoE)",
      "Owns usability testing + feedback loops"]),
    (PART, PART_DARK, "PARTNERSHIP CREW", "Vendor Management",
     ["Accenture (Global SAP + Corporate)",
      "SWF / Software Factory (all teams)",
      "Smart Metering vendor (100% delivery)",
      "Iberdrola corporate liaison"]),
    (FAC, FAC_DARK,  "FACILITATION CREW","Coaches & Process",
     ["Process Improvement role (NOM v1.5)",
      "Agile Coaches (from Digital CoE Agility Chapter)",
      "Scrum Masters per squad",
      "Knowledge Manager (NOM v1.5)"]),
]

sup_w = (W - 160 - 3 * 24) // 4
for i, (col, dark, name, sub, members) in enumerate(sup_crews):
    x0 = 80 + i * (sup_w + 24)
    x1 = x0 + sup_w
    d.rounded_rectangle([x0, sup_y0, x1, sup_y1], radius=20, fill=col,
                        outline=dark, width=3)
    d.rounded_rectangle([x0, sup_y0, x1, sup_y0 + 70], radius=20, fill=dark)
    d.rectangle([x0, sup_y0 + 50, x1, sup_y0 + 70], fill=dark)
    d.text((x0 + 20, sup_y0 + 14), name, font=fnt(22, True), fill=WHITE)
    d.text((x0 + 20, sup_y0 + 80), sub, font=fnt(18, True), fill=WHITE)
    for j, m in enumerate(members):
        d.text((x0 + 20, sup_y0 + 120 + j * 28),
               "• " + m, font=fnt(15), fill=WHITE)


# ── Forums row (bottom) ──────────────────────────────────────────────────
forum_y0 = sup_y1 + 50
forum_y1 = forum_y0 + 200

forums = [
    (FORUM_REG, "REGIONAL FORUM",     "Per OPCO discussion",
     ["UIL — UI · SCG · CNG · BGC (Connecticut)",
      "CMP — Central Maine Power",
      "NY — NYSEG · RGE"]),
    (FORUM_FUN, "FUNCTIONAL FORUM",   "Per role discussion",
     ["PM Forum — cross-stream PMs",
      "Tech Lead Forum — architecture review",
      "Project Manager Forum — Sarath's pool"]),
    (FORUM_TEC, "TECHNOLOGICAL FORUM","Per tech-stack discussion",
     ["SAP Forum — Local SAP + Iberdrola",
      "Non-SAP Forum — GIS · ProjectWise · Mobility",
      "Platform Forum — SWF · DevOps · Security"]),
]

for_w = (W - 160 - 2 * 24) // 3
for i, (col, name, sub, members) in enumerate(forums):
    x0 = 80 + i * (for_w + 24)
    x1 = x0 + for_w
    d.rounded_rectangle([x0, forum_y0, x1, forum_y1], radius=18, fill=col,
                        outline=DARK_TXT, width=2)
    d.text((x0 + 20, forum_y0 + 14), name, font=fnt(24, True), fill=DARK_TXT)
    d.text((x0 + 20, forum_y0 + 56), sub, font=fnt(18), fill=GREY)
    for j, m in enumerate(members):
        d.text((x0 + 20, forum_y0 + 100 + j * 26),
               "• " + m, font=fnt(15), fill=DARK_TXT)


# ── Legend / footer ──────────────────────────────────────────────────────
d.text((80, H - 50),
       "Source: Unfix framework crew/forum vocabulary · Avangrid Digital Products TOM · Apr 2026",
       font=fnt(20), fill=GREY)


OUT.mkdir(parents=True, exist_ok=True)
out_file = OUT / "DigitalProducts_Unfix_TOM.png"
img.save(out_file, "PNG", dpi=(150, 150))

# preview
ratio = 1600 / W
preview = img.resize((1600, int(H * ratio)), Image.LANCZOS)
preview.save(out_file.with_name(out_file.stem + "_preview.png"), "PNG")
print(f"✓ Saved: {out_file}")
print(f"  Resolution: {W}×{H}")
