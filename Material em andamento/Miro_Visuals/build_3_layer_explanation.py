"""
Simple 3-layer explanation of the Portfolio architecture.
Shows ONE Initiative (Lemur-SF migration) flowing through all 3 layers.
Designed to be visually clear, not dense.
"""
from PIL import Image, ImageDraw, ImageFont
from pathlib import Path
import math

OUT = Path("/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Miro_Visuals")

# ── Palette ──
NAVY       = ( 33,  60, 105)
NAVY_DARK  = ( 17,  35,  62)
BLUE       = (103, 133, 193)
BLUE_LIGHT = (220, 230, 245)
TEAL       = ( 60, 160, 170)
ORANGE     = (230, 140,  55)
ORANGE_BG  = (255, 240, 220)
AMBER      = (230, 182,   0)
RED        = (200,  60,  50)
GREEN      = ( 90, 165,  85)
GREEN_BG   = (220, 240, 220)
GREY       = (110, 120, 135)
LIGHT_GREY = (240, 242, 246)
CARD_BG    = (255, 255, 255)
DARK       = ( 30,  40,  55)
WHITE      = (255, 255, 255)
HIGHLIGHT  = (255, 235, 130)
PERSON_F   = (210, 215, 225)
PERSON_BR  = (150, 160, 175)


W, H = 4000, 3000


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


def arrow_down(d, x, y_start, y_end, color=RED, width=8):
    """Vertical arrow with arrowhead at bottom."""
    d.line([x, y_start, x, y_end - 20], fill=color, width=width)
    d.polygon([(x, y_end), (x - 22, y_end - 30), (x + 22, y_end - 30)],
              fill=color)


def card(d, x0, y0, x1, y1, title, subtitle="", highlight=False):
    fill = HIGHLIGHT if highlight else CARD_BG
    d.rounded_rectangle([x0, y0, x1, y1], radius=10, fill=fill,
                        outline=GREY, width=2)
    if highlight:
        # Bold title
        d.text((x0 + 12, y0 + 10), title, fill=DARK, font=fnt(15, True), anchor="lt")
    else:
        d.text((x0 + 12, y0 + 10), title, fill=DARK, font=fnt(13, True), anchor="lt")
    if subtitle:
        d.text((x0 + 12, y0 + 32), subtitle, fill=GREY, font=fnt(11), anchor="lt")


img = Image.new("RGB", (W, H), LIGHT_GREY)
d = ImageDraw.Draw(img)


# ── Header ──
d.rectangle([0, 0, W, 130], fill=NAVY_DARK)
d.text((80, 28), "PORTFOLIO ARCHITECTURE — 3 LAYERS",
       font=fnt(54, True), fill=WHITE)
d.text((80, 88),
       "Como uma Initiative grande flui do VP até a squad · exemplo: migrar inspeções para Lemur-Salesforce",
       font=fnt(20), fill=(199, 211, 228))


# ════════════════════════════════════════════════════════════════════════
# LAYER 0 — VP Roll-up Dashboard (top, narrow band)
# ════════════════════════════════════════════════════════════════════════
L0_y0 = 200
L0_y1 = 480

d.rounded_rectangle([100, L0_y0, W - 100, L0_y1], radius=24,
                    fill=GREEN_BG, outline=GREEN, width=4)

# Layer label (left)
d.rounded_rectangle([130, L0_y0 - 30, 380, L0_y0 + 30],
                    radius=14, fill=GREEN)
centered(d, (255, L0_y0), "Nível 3 — VP", fnt(22, True), WHITE)

# Title
left(d, (160, L0_y0 + 60), "VP ROLL-UP DASHBOARD",
     fnt(28, True), DARK)
left(d, (160, L0_y0 + 100),
     "Vitor Cruz + Portfolio Manager · monthly Portfolio Sync",
     fnt(18), GREY)
left(d, (160, L0_y0 + 130),
     "É uma VIEW agregada — sem cards próprios, só puxa dos 2 Portfolio Boards abaixo",
     fnt(15, italic=True), GREY)

# Mock metrics inside the band
metrics_y = L0_y0 + 175
metric_w = 540
for i, (label, value, sub) in enumerate([
    ("Initiatives ativas", "27", "across both portfolios"),
    ("Capacity utilizada", "82%", "across 8 crews"),
    ("Cross-portfolio deps", "4", "monitored"),
    ("Top initiative (this view)", "→ IT05 Lemur-SF", "highlighted"),
]):
    x0 = 160 + i * (metric_w + 16)
    x1 = x0 + metric_w
    d.rounded_rectangle([x0, metrics_y, x1, metrics_y + 110],
                        radius=12, fill=WHITE, outline=GREEN, width=2)
    centered(d, ((x0 + x1) / 2, metrics_y + 28), label, fnt(15, True), DARK)
    centered(d, ((x0 + x1) / 2, metrics_y + 60), value, fnt(28, True), GREEN)
    centered(d, ((x0 + x1) / 2, metrics_y + 90), sub, fnt(12), GREY)

# Persons on top
person(d, 200, L0_y0 - 60, r=24)
d.text((200, L0_y0 - 22), "Vitor", fill=DARK, font=fnt(13, True), anchor="mm")


# ════════════════════════════════════════════════════════════════════════
# LAYER 1 — 2 Portfolio Boards (middle, two side-by-side)
# ════════════════════════════════════════════════════════════════════════
L1_y0 = 580
L1_y1 = 1480

# Left: IT Apps Portfolio (Carlos)
it_x0 = 100
it_x1 = (W - 60) // 2
# Right: CS Apps Portfolio (Kumar)
cs_x0 = it_x1 + 60
cs_x1 = W - 100

# Layer label
d.rounded_rectangle([130, L1_y0 - 30, 460, L1_y0 + 30],
                    radius=14, fill=BLUE)
centered(d, (295, L1_y0), "Nível 2 — Directors", fnt(22, True), WHITE)

# IT Apps Portfolio Board
d.rounded_rectangle([it_x0, L1_y0, it_x1, L1_y1], radius=18,
                    fill=BLUE_LIGHT, outline=BLUE, width=4)
left(d, (it_x0 + 30, L1_y0 + 30), "IT BUSINESS APPS · PORTFOLIO BOARD",
     fnt(22, True), DARK)
left(d, (it_x0 + 30, L1_y0 + 62),
     "Owner: Carlos Garbin · Cards = INITIATIVES (grandes projetos)",
     fnt(14, italic=True), GREY)
person(d, it_x0 + 50, L1_y0 - 60, r=24)
d.text((it_x0 + 50, L1_y0 - 22), "Carlos", fill=DARK, font=fnt(13, True), anchor="mm")

# Kanban columns inside IT Apps board
columns_it = ["Funnel", "Reviewing", "Analyzing", "Backlog", "Implementing", "Done"]
col_y0 = L1_y0 + 100
col_y1 = L1_y1 - 40
col_w = (it_x1 - it_x0 - 40) // len(columns_it)
for i, col_name in enumerate(columns_it):
    cx0 = it_x0 + 20 + i * col_w
    cx1 = cx0 + col_w - 6
    # Column header
    d.rectangle([cx0, col_y0, cx1, col_y0 + 36], fill=NAVY)
    centered(d, ((cx0 + cx1) / 2, col_y0 + 18), col_name, fnt(13, True), WHITE)
    # Column body
    d.rectangle([cx0, col_y0 + 36, cx1, col_y1], fill=WHITE,
                outline=GREY, width=1)

# Cards in IT Apps portfolio
def col_x(idx, x0, w_total):
    return x0 + 20 + idx * w_total

# Funnel
card(d, col_x(0, it_x0, col_w) + 6, col_y0 + 50,
     col_x(0, it_x0, col_w) + col_w - 12, col_y0 + 110,
     "S/4HANA roadmap", "Initiative · cross-branch")
card(d, col_x(0, it_x0, col_w) + 6, col_y0 + 124,
     col_x(0, it_x0, col_w) + col_w - 12, col_y0 + 184,
     "Storm program v2", "Initiative · Field Ops")

# Reviewing
card(d, col_x(1, it_x0, col_w) + 6, col_y0 + 50,
     col_x(1, it_x0, col_w) + col_w - 12, col_y0 + 110,
     "ProjectWise upgrade", "Initiative · E&A")

# Analyzing
card(d, col_x(2, it_x0, col_w) + 6, col_y0 + 50,
     col_x(2, it_x0, col_w) + col_w - 12, col_y0 + 110,
     "ESRI UN deployment", "Initiative · E&A · OpCo")

# Backlog
card(d, col_x(3, it_x0, col_w) + 6, col_y0 + 50,
     col_x(3, it_x0, col_w) + col_w - 12, col_y0 + 110,
     "Doc viewer scale RGE", "Initiative · Field Ops")

# Implementing — HIGHLIGHTED
card(d, col_x(4, it_x0, col_w) + 6, col_y0 + 50,
     col_x(4, it_x0, col_w) + col_w - 12, col_y0 + 130,
     "Lemur-SF Migration",
     "Initiative IT05 · Field Ops Crew",
     highlight=True)
# Arrow indicator from highlighted card
d.text((col_x(4, it_x0, col_w) + col_w / 2, col_y0 + 150),
       "→ goes down to crew", fill=RED, font=fnt(11, True), anchor="mm")

card(d, col_x(4, it_x0, col_w) + 6, col_y0 + 170,
     col_x(4, it_x0, col_w) + col_w - 12, col_y0 + 230,
     "Mobility Program", "Initiative · Field Ops")

# Done
card(d, col_x(5, it_x0, col_w) + 6, col_y0 + 50,
     col_x(5, it_x0, col_w) + col_w - 12, col_y0 + 110,
     "Q1 storm pilot", "Done · 2026 Q1")


# CS Apps Portfolio Board
d.rounded_rectangle([cs_x0, L1_y0, cs_x1, L1_y1], radius=18,
                    fill=BLUE_LIGHT, outline=BLUE, width=4)
left(d, (cs_x0 + 30, L1_y0 + 30), "CUSTOMER SERVICES APPS · PORTFOLIO BOARD",
     fnt(22, True), DARK)
left(d, (cs_x0 + 30, L1_y0 + 62),
     "Owner: Kumar Jaladanki · Cards = INITIATIVES",
     fnt(14, italic=True), GREY)
person(d, cs_x0 + 50, L1_y0 - 60, r=24)
d.text((cs_x0 + 50, L1_y0 - 22), "Kumar", fill=DARK, font=fnt(13, True), anchor="mm")

# Kanban columns inside CS Apps board (same structure)
col_w2 = (cs_x1 - cs_x0 - 40) // len(columns_it)
for i, col_name in enumerate(columns_it):
    cx0 = cs_x0 + 20 + i * col_w2
    cx1 = cx0 + col_w2 - 6
    d.rectangle([cx0, col_y0, cx1, col_y0 + 36], fill=NAVY)
    centered(d, ((cx0 + cx1) / 2, col_y0 + 18), col_name, fnt(13, True), WHITE)
    d.rectangle([cx0, col_y0 + 36, cx1, col_y1], fill=WHITE,
                outline=GREY, width=1)

# Cards in CS Apps portfolio
card(d, cs_x0 + 26, col_y0 + 50, cs_x0 + 20 + col_w2 - 6, col_y0 + 110,
     "Customer Journey W4", "Initiative · CCC")
card(d, cs_x0 + 26, col_y0 + 124, cs_x0 + 20 + col_w2 - 6, col_y0 + 184,
     "Customer Service GenAI", "Initiative · CCC")

card(d, cs_x0 + 20 + col_w2 + 6, col_y0 + 50,
     cs_x0 + 20 + 2 * col_w2 - 6, col_y0 + 110,
     "EBD Certification Auto", "Initiative · M2C-RO")

card(d, cs_x0 + 20 + 2 * col_w2 + 6, col_y0 + 50,
     cs_x0 + 20 + 3 * col_w2 - 6, col_y0 + 110,
     "MV90 Upgrade UIL", "Initiative · M2C-SL")

card(d, cs_x0 + 20 + 3 * col_w2 + 6, col_y0 + 50,
     cs_x0 + 20 + 4 * col_w2 - 6, col_y0 + 110,
     "S/4HANA M2C", "Initiative · cross")

card(d, cs_x0 + 20 + 4 * col_w2 + 6, col_y0 + 50,
     cs_x0 + 20 + 5 * col_w2 - 6, col_y0 + 130,
     "DG SAP Billing Auto", "Initiative · M2C-RO")

card(d, cs_x0 + 20 + 5 * col_w2 + 6, col_y0 + 50,
     cs_x0 + 20 + 6 * col_w2 - 6, col_y0 + 110,
     "Hardship Comm UIL", "Done · 2026 Q1")


# Arrow from IT Apps Implementing → Crew board below
arrow_x = col_x(4, it_x0, col_w) + col_w / 2 - 6
arrow_down(d, arrow_x, L1_y1 + 10, 1620, color=RED, width=8)
d.text((arrow_x + 30, (L1_y1 + 1620) / 2),
       "Initiative is assigned to a crew", fill=RED, font=fnt(15, True), anchor="lm")


# ════════════════════════════════════════════════════════════════════════
# LAYER 2 — Crew Kanban (Field Operations) — bottom, full width
# ════════════════════════════════════════════════════════════════════════
L2_y0 = 1640
L2_y1 = 2880

d.rounded_rectangle([100, L2_y0, W - 100, L2_y1], radius=18,
                    fill=ORANGE_BG, outline=ORANGE, width=4)

# Layer label
d.rounded_rectangle([130, L2_y0 - 30, 580, L2_y0 + 30],
                    radius=14, fill=ORANGE)
centered(d, (355, L2_y0), "Nível 1 — PM da Crew (8 PMs)",
         fnt(22, True), WHITE)

left(d, (160, L2_y0 + 50), "FIELD OPERATIONS CREW · KANBAN",
     fnt(26, True), DARK)
left(d, (160, L2_y0 + 86),
     "Owner: PM Field Operations · Initiative IT05 → 4 Epics → Features → User Stories",
     fnt(15, italic=True), GREY)
person(d, 200, L2_y0 - 60, r=24)
d.text((200, L2_y0 - 22), "PM Crew", fill=DARK, font=fnt(13, True), anchor="mm")

# 5 phases columns
phases = ["Phase 0\nDISCOVERY", "Phase 1\nMATERIALIZATION",
          "Phase 2\nCREATION", "Phase 3\nONGOING", "Phase 4\nMAINTENANCE"]
ph_y0 = L2_y0 + 130
ph_y1 = L2_y1 - 40
ph_w = (W - 100 - 100 - 40) // len(phases)
for i, ph_name in enumerate(phases):
    px0 = 120 + i * ph_w
    px1 = px0 + ph_w - 6
    # Header
    d.rectangle([px0, ph_y0, px1, ph_y0 + 50], fill=NAVY)
    lines = ph_name.split("\n")
    centered(d, ((px0 + px1) / 2, ph_y0 + 16), lines[0], fnt(13), WHITE)
    centered(d, ((px0 + px1) / 2, ph_y0 + 36), lines[1], fnt(15, True), WHITE)
    # Body
    d.rectangle([px0, ph_y0 + 50, px1, ph_y1], fill=WHITE,
                outline=GREY, width=1)

# Dual track horizontal divider
mid_y = ph_y0 + 50 + (ph_y1 - ph_y0 - 50) // 2
d.line([120, mid_y, W - 100 - 20, mid_y], fill=BLUE, width=3)
# Track labels (left side)
d.rectangle([20, ph_y0 + 50, 120, mid_y], fill=BLUE)
tmp = Image.new("RGBA", (mid_y - ph_y0 - 50 - 20, 100), (0, 0, 0, 0))
td = ImageDraw.Draw(tmp)
td.text((20, 50), "DISCOVERY TRACK",
        fill=WHITE, font=fnt(20, True), anchor="lm")
tmp = tmp.rotate(90, expand=True)
img.paste(tmp, (20, ph_y0 + 60), tmp)

d.rectangle([20, mid_y, 120, ph_y1], fill=AMBER)
tmp = Image.new("RGBA", (ph_y1 - mid_y - 20, 100), (0, 0, 0, 0))
td = ImageDraw.Draw(tmp)
td.text((20, 50), "DELIVERY TRACK",
        fill=WHITE, font=fnt(20, True), anchor="lm")
tmp = tmp.rotate(90, expand=True)
img.paste(tmp, (20, mid_y + 10), tmp)

# Cards inside the crew kanban — show the 4 Epics flowing
def crew_col_x(idx):
    return 120 + idx * ph_w

# Phase 0 (Discovery track top)
card(d, crew_col_x(0) + 10, ph_y0 + 70,
     crew_col_x(0) + ph_w - 16, ph_y0 + 130,
     "Epic A · Migrate DLI", "exploring offline reqs",
     highlight=True)

# Phase 1 (Materialization)
card(d, crew_col_x(1) + 10, ph_y0 + 70,
     crew_col_x(1) + ph_w - 16, ph_y0 + 130,
     "Epic B · Migrate TLI", "sizing with Lemur vendor")

# Phase 2 (Creation, delivery track)
card(d, crew_col_x(2) + 10, mid_y + 20,
     crew_col_x(2) + ph_w - 16, mid_y + 80,
     "Epic C · Migrate DGI", "team assigned · Q2 sprint plan")

# Phase 3 (Ongoing — multiple Features in delivery)
card(d, crew_col_x(3) + 10, mid_y + 20,
     crew_col_x(3) + ph_w - 16, mid_y + 80,
     "Feature: Offline DLI", "in dev · Sprint 5",
     highlight=True)
card(d, crew_col_x(3) + 10, mid_y + 90,
     crew_col_x(3) + ph_w - 16, mid_y + 150,
     "Feature: Mobile TLI UI", "in dev · Sprint 5")
card(d, crew_col_x(3) + 10, mid_y + 160,
     crew_col_x(3) + ph_w - 16, mid_y + 220,
     "Feature: Sync logic", "in QA · Sprint 4")

# User stories indicator
d.text((crew_col_x(3) + ph_w / 2 - 6, mid_y + 250),
       "↓ Sprint board view shows User Stories",
       fill=AMBER, font=fnt(13, True), anchor="mm")

# Phase 4 (Maintenance)
card(d, crew_col_x(4) + 10, mid_y + 20,
     crew_col_x(4) + ph_w - 16, mid_y + 80,
     "Epic D · FWM Decommission", "hypercare · 30d")


# ════════════════════════════════════════════════════════════════════════
# Bottom legend / summary
# ════════════════════════════════════════════════════════════════════════
legend_y = L2_y1 + 30
left(d, (160, legend_y),
     "Total: 1 dashboard + 2 portfolio boards + 8 crew kanbans = 11 boards/views.",
     fnt(20, True), DARK)
left(d, (160, legend_y + 40),
     "Cada nível olha apenas o seu — Vitor olha agregado · Director olha initiatives · PM olha Epics/Features · Squad olha Stories no sprint.",
     fnt(16, italic=True), GREY)


OUT.mkdir(parents=True, exist_ok=True)
out_file = OUT / "Portfolio_3Layers_Explained.png"
img.save(out_file, "PNG", dpi=(150, 150))

ratio = 1700 / W
preview = img.resize((1700, int(H * ratio)), Image.LANCZOS)
preview.save(out_file.with_name(out_file.stem + "_preview.png"), "PNG")
print(f"✓ Saved: {out_file}")
print(f"  Resolution: {W}×{H}")
