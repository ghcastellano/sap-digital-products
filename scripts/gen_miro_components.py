"""Miro Component Library — drag-drop stickers for Demand × Portfolio sessions"""
from PIL import Image, ImageDraw, ImageFont
import os, shutil

# ── Colors (match canvas style) ─────────────────────────────────────────
BG      = (2, 17, 30)
CARD    = (4, 22, 44)
WHITE   = (255, 255, 255)
DIM     = (160, 175, 195)
MUTED   = (100, 120, 145)
INNER   = (8, 28, 52)
BORD    = (40, 60, 85)

CYAN    = (80, 180, 220)
PINK    = (200, 80, 180)
PINK_L  = (230, 130, 200)
ORANGE  = (247, 148, 29)
BLUE    = (0, 120, 200)
GREEN   = (120, 190, 32)
LIME    = (160, 210, 40)
RED     = (220, 60, 60)
YELLOW  = (240, 200, 60)
GRAY    = (150, 165, 185)

def font(size, bold=False):
    for p in ["/System/Library/Fonts/Helvetica.ttc",
              "/System/Library/Fonts/SFNSText.ttf",
              "/System/Library/Fonts/Supplemental/Arial.ttf"]:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size, index=1 if bold and p.endswith(".ttc") else 0)
            except Exception:
                return ImageFont.truetype(p, size)
    return ImageFont.load_default()

FLABEL = font(26, True)
FMONO_LG = font(42, True)
FMONO_MD = font(34, True)
FMONO_SM = font(26, True)

W_COMP = 340
H_COMP = 140
ICON_W = 110
R      = 14
BORDER_W = 4


def rrect(d, x, y, w, h, fill=CARD, outline=None, r=R, width=BORDER_W):
    d.rounded_rectangle([x, y, x + w, y + h], r, fill=fill,
                        outline=outline, width=width if outline else 0)


def wrap_label(text, max_w, f, d):
    words = text.split()
    lines = []
    line = ""
    for w in words:
        t = f"{line} {w}".strip()
        if d.textlength(t, font=f) <= max_w:
            line = t
        else:
            if line:
                lines.append(line)
            line = w
    if line:
        lines.append(line)
    return lines[:3]


def pick_mono_font(mono, d, max_w):
    """Choose the biggest font where the monogram fits in the icon panel."""
    for f in (FMONO_LG, FMONO_MD, FMONO_SM):
        if d.textlength(mono, font=f) <= max_w:
            return f
    return FMONO_SM


def draw_icon(img, d, x, y, color, mono):
    """Draw a colored square with a centered white monogram."""
    panel_w = ICON_W - 16
    panel_h = H_COMP - 32
    ix = x + 8
    iy = y + 16
    # Fill with category color
    rrect(d, ix, iy, panel_w, panel_h, fill=color, outline=None, r=R - 4, width=0)
    # Monogram in white, centered
    f = pick_mono_font(mono, d, panel_w - 20)
    tw = d.textlength(mono, font=f)
    th = f.size
    d.text((ix + (panel_w - tw) // 2, iy + (panel_h - th) // 2 - 4),
           mono, fill=WHITE, font=f)


def make_component(label, mono, color, out_path):
    img = Image.new("RGBA", (W_COMP, H_COMP), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    # Outer rounded card
    rrect(d, 2, 2, W_COMP - 4, H_COMP - 4, fill=BG, outline=color, r=R, width=BORDER_W)

    # Icon panel with monogram
    draw_icon(img, d, 0, 0, color, mono)

    # Label (right side)
    text_x = ICON_W + 18
    text_w = W_COMP - text_x - 18
    lines = wrap_label(label, text_w, FLABEL, d)
    line_h = 32
    total_h = len(lines) * line_h
    ty = (H_COMP - total_h) // 2
    for ln in lines:
        d.text((text_x, ty), ln, fill=WHITE, font=FLABEL)
        ty += line_h

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    img.save(out_path, "PNG")


# ═══════════════════════════════════════════════════════════════════════
# CATEGORIES — each item is (label, monogram)
# ═══════════════════════════════════════════════════════════════════════
BASE = "/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Miro Discovery Sessions/Components"

categories = [
    {
        "folder": "01_Channels",
        "title":  "CHANNELS",
        "color":  CYAN,
        "items": [
            ("Email",              "@"),
            ("Teams Meeting",      "TM"),
            ("Teams DM",           "DM"),
            ("Phone",              "PH"),
            ("Jira Ticket",        "JI"),
            ("SharePoint Form",    "SP"),
            ("ServiceNow",         "SN"),
            ("Walk-in",            "WI"),
            ("Direct Message",     "MSG"),
            ("Annual Planning (PDD)", "PDD"),
        ],
    },
    {
        "folder": "02_Source_Areas",
        "title":  "SOURCE AREAS",
        "color":  ORANGE,
        "items": [
            ("BRS",                "BRS"),
            ("OPCO NY",            "NY"),
            ("OPCO CT",            "CT"),
            ("OPCO ME",            "ME"),
            ("Regulatory",         "REG"),
            ("Procurement",        "PRC"),
            ("OT",                 "OT"),
            ("IA",                 "IA"),
            ("Data & Analytics",   "D&A"),
            ("Infrastructure",     "INF"),
            ("CX & Digital CoE",   "CX"),
            ("P&T",                "P&T"),
            ("Corporate",          "CRP"),
            ("GIS",                "GIS"),
        ],
    },
    {
        "folder": "03_Request_Types",
        "title":  "REQUEST TYPES",
        "color":  PINK,
        "items": [
            ("Bug",          "BUG"),
            ("Enhancement",  "ENH"),
            ("New Feature",  "NEW"),
            ("Regulatory",   "REG"),
            ("Tech Debt",    "TD"),
            ("Compliance",   "CMP"),
            ("Small",        "S"),
            ("Medium",       "M"),
            ("Large",        "L"),
            ("Urgent",       "!"),
        ],
    },
    {
        "folder": "04_Classes_of_Service",
        "title":  "CLASSES OF SERVICE",
        "color":  PINK_L,
        "items": [
            ("Expedite",     "EXP"),
            ("Fixed Date",   "FD"),
            ("Standard",     "STD"),
            ("Intangible",   "INT"),
        ],
    },
    {
        "folder": "05_Roles",
        "title":  "ROLES",
        "color":  LIME,
        "items": [
            ("Requester",       "REQ"),
            ("Product Owner",   "PO"),
            ("BA",              "BA"),
            ("Developer",       "DEV"),
            ("Project Manager", "PM"),
            ("Architect",       "ARC"),
            ("Stakeholder",     "SH"),
            ("Sponsor",         "SPN"),
            ("Approver",        "APR"),
            ("Vendor",          "VND"),
        ],
    },
    {
        "folder": "06_Tools",
        "title":  "TOOLS / SYSTEMS",
        "color":  BLUE,
        "items": [
            ("Jira",        "J"),
            ("Confluence",  "CF"),
            ("PPMNow",      "PPM"),
            ("SharePoint",  "SP"),
            ("ServiceNow",  "SN"),
            ("SAP",         "SAP"),
            ("Excel",       "XL"),
            ("Teams",       "T"),
            ("Outlook",     "OL"),
            ("BI / BW",     "BI"),
        ],
    },
    {
        "folder": "07_Decision_Gates",
        "title":  "DECISION GATES",
        "color":  GREEN,
        "items": [
            ("Approved",    "OK"),
            ("Rejected",    "NO"),
            ("On Hold",     "II"),
            ("Escalated",   "UP"),
            ("Needs Info",  "?"),
            ("Deferred",    "DEF"),
        ],
        "per_item_colors": {
            "Approved":   GREEN,
            "Rejected":   RED,
            "On Hold":    YELLOW,
            "Escalated":  ORANGE,
            "Needs Info": CYAN,
            "Deferred":   GRAY,
        },
    },
    {
        "folder": "08_Governance_Forums",
        "title":  "GOVERNANCE FORUMS",
        "color":  ORANGE,
        "items": [
            ("Portfolio Committee", "PC"),
            ("Shark Tank",          "ST"),
            ("Sprint Review",       "SR"),
            ("Daily Standup",       "DS"),
            ("Monthly Review",      "MR"),
            ("Annual Planning",     "AP"),
            ("CAB",                 "CAB"),
            ("Steering",            "SC"),
        ],
    },
    {
        "folder": "09_Flow_Shapes",
        "title":  "FLOW SHAPES",
        "color":  GRAY,
        "items": [
            ("Start",    "S"),
            ("End",      "E"),
            ("Decision", "D"),
            ("Process",  "P"),
            ("Fork",     "FK"),
            ("Merge",    "MG"),
        ],
    },
    {
        "folder": "10_Annotations",
        "title":  "ANNOTATIONS",
        "color":  YELLOW,
        "items": [
            ("AS-IS",      "AS"),
            ("TO-BE",      "TB"),
            ("GAP",        "GAP"),
            ("Pain Point", "PP"),
            ("Idea",       "IDE"),
            ("Question",   "?"),
            ("Blocker",    "BLK"),
            ("Risk",       "RSK"),
        ],
        "per_item_colors": {
            "AS-IS":      RED,
            "TO-BE":      GREEN,
            "GAP":        ORANGE,
            "Pain Point": RED,
            "Idea":       YELLOW,
            "Question":   CYAN,
            "Blocker":    RED,
            "Risk":       ORANGE,
        },
    },
]


# ═══════════════════════════════════════════════════════════════════════
# CLEAN + GENERATE INDIVIDUAL COMPONENTS
# ═══════════════════════════════════════════════════════════════════════
if os.path.exists(BASE):
    shutil.rmtree(BASE)

total = 0
for cat in categories:
    folder = os.path.join(BASE, cat["folder"])
    os.makedirs(folder, exist_ok=True)
    for label, mono in cat["items"]:
        color = cat.get("per_item_colors", {}).get(label, cat["color"])
        safe = (label.replace(" ", "_").replace("/", "-").replace("&", "and")
                .replace("(", "").replace(")", ""))
        out = os.path.join(folder, f"{safe}.png")
        make_component(label, mono, color, out)
        total += 1

print(f"Generated {total} component PNGs")


# ═══════════════════════════════════════════════════════════════════════
# MASTER REFERENCE GRID
# ═══════════════════════════════════════════════════════════════════════
COLS = 5
GAP = 24
SECT_GAP = 60
MARGIN = 60
HDR_H = 80

FT = font(56, True)
FSECT = font(36, True)
FSECTSUB = font(24)

grid_w = MARGIN * 2 + COLS * W_COMP + (COLS - 1) * GAP
total_h = MARGIN + 160
for cat in categories:
    n = len(cat["items"])
    rows = (n + COLS - 1) // COLS
    total_h += HDR_H + rows * (H_COMP + GAP) + SECT_GAP
total_h += MARGIN

ref = Image.new("RGB", (grid_w, total_h), BG)
rd = ImageDraw.Draw(ref)

Y = MARGIN
rd.text((MARGIN, Y), "Miro Component Library", fill=WHITE, font=FT)
Y += 72
rd.text((MARGIN, Y), "Drag-drop stickers for Demand × Portfolio discovery sessions  |  Avangrid × NTT DATA",
        fill=DIM, font=font(28))
Y += 60

for cat in categories:
    color = cat["color"]
    rd.rounded_rectangle([MARGIN, Y, grid_w - MARGIN, Y + HDR_H - 16],
                         14, fill=INNER, outline=color, width=3)
    rd.ellipse([MARGIN + 20, Y + (HDR_H - 16) // 2 - 14,
                MARGIN + 48, Y + (HDR_H - 16) // 2 + 14], fill=color)
    rd.text((MARGIN + 64, Y + 8), cat["title"], fill=WHITE, font=FSECT)
    rd.text((MARGIN + 64, Y + 44), f"{len(cat['items'])} components",
            fill=DIM, font=FSECTSUB)
    Y += HDR_H

    for i, (label, mono) in enumerate(cat["items"]):
        r = i // COLS
        c = i % COLS
        cx = MARGIN + c * (W_COMP + GAP)
        cy = Y + r * (H_COMP + GAP)
        color_i = cat.get("per_item_colors", {}).get(label, cat["color"])
        # Card
        rrect(rd, cx + 2, cy + 2, W_COMP - 4, H_COMP - 4,
              fill=BG, outline=color_i, r=R, width=BORDER_W)
        # Icon panel + monogram
        draw_icon(ref, rd, cx, cy, color_i, mono)
        # Label
        tx = cx + ICON_W + 18
        tw = W_COMP - ICON_W - 36
        lines = wrap_label(label, tw, FLABEL, rd)
        ty = cy + (H_COMP - len(lines) * 32) // 2
        for ln in lines:
            rd.text((tx, ty), ln, fill=WHITE, font=FLABEL)
            ty += 32

    rows = (len(cat["items"]) + COLS - 1) // COLS
    Y += rows * (H_COMP + GAP) + SECT_GAP

ref = ref.crop((0, 0, grid_w, Y))
out_ref = os.path.join(BASE, "Components_Library_Reference.png")
ref.save(out_ref, "PNG", optimize=True)
px = ref.size[0] * ref.size[1]
print(f"Master reference: {ref.size[0]}x{ref.size[1]} — {px/1e6:.1f}MP")
