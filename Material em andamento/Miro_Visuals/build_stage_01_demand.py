"""
Stage 01 — DEMAND MANAGEMENT
Big single-column board: 1 stage × 5 dimensions, lots of breathing room.
Blank cells, ready for Miro fill-in.
"""
from PIL import Image, ImageDraw, ImageFont
from pathlib import Path

OUT = Path("/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Miro_Visuals")

NAVY      = (44, 90, 136)
NAVY_DARK = (21, 42, 71)
TEAL      = (43, 107, 119)
GREEN     = (76, 140, 63)
AMBER     = (232, 139, 47)
RED       = (182, 59, 31)
LIGHT_BG  = (244, 245, 247)
GREY_TXT  = (75, 85, 99)
WHITE     = (255, 255, 255)
BORDER    = (210, 216, 224)

# Big canvas — tall, single column, lots of space
W, H = 2800, 4200


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


img = Image.new("RGB", (W, H), WHITE)
d = ImageDraw.Draw(img)

# ── Header band ────────────────────────────────────────────────────────────
d.rectangle([0, 0, W, 280], fill=NAVY_DARK)
d.text((100, 60), "STAGE 01", font=fnt(48, True), fill=AMBER)
d.text((100, 130), "DEMAND MANAGEMENT", font=fnt(96, True), fill=WHITE)
d.text((100, 230), "What demand looks like at intake — capabilities, roles, tools, governance, performance",
       font=fnt(34), fill=(199, 211, 228))

# ── Layout ─────────────────────────────────────────────────────────────────
margin_x   = 120
top_y      = 360
label_w    = 520
gap_label  = 70
gap_row    = 70
footer_h   = 120

dims = [
    ("CAPABILITIES", "What the org must be able to do",   GREEN),
    ("ROLES",        "Who owns and operates this",         TEAL),
    ("TOOLS",        "Systems and platforms used",         AMBER),
    ("GOVERNANCE",   "Decisions, forums, controls",        NAVY),
    ("PERFORMANCE",  "How success is measured",            RED),
]

n = len(dims)
available_h = H - top_y - footer_h - (n - 1) * gap_row
cell_h      = available_h // n
cell_x0     = margin_x + label_w + gap_label
cell_x1     = W - margin_x

for r, (dim, sub, color) in enumerate(dims):
    y0 = top_y + r * (cell_h + gap_row)
    # Dimension label (left side)
    d.rounded_rectangle([margin_x, y0, margin_x + label_w, y0 + cell_h],
                        radius=24, fill=color)
    centered(d, (margin_x + label_w / 2, y0 + cell_h / 2 - 30),
             dim, fnt(56, True), WHITE)
    centered(d, (margin_x + label_w / 2, y0 + cell_h / 2 + 38),
             sub, fnt(24), (235, 240, 245))

    # Empty cell (big writable area)
    d.rounded_rectangle([cell_x0, y0, cell_x1, y0 + cell_h],
                        radius=28, fill=LIGHT_BG, outline=BORDER, width=4)

# ── Footer ─────────────────────────────────────────────────────────────────
d.text((margin_x, H - 80),
       "Avangrid · Digital Products · Target Operating Model — Stage 01 Demand Management",
       font=fnt(26), fill=GREY_TXT)
d.text((W - 720, H - 80),
       "Workshop board · 1 of 4",
       font=fnt(26, True), fill=AMBER)

OUT.mkdir(parents=True, exist_ok=True)
out_file = OUT / "01_TOM_DemandManagement_BLANK.png"
img.save(out_file, "PNG", dpi=(150, 150))
print(f"✓ Saved: {out_file}")
print(f"  Resolution: {W}×{H}")
