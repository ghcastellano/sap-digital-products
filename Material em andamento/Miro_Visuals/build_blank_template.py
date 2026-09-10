"""
Builds an EMPTY 4×5 TOM template (big format, ready for Miro fill-in).
No cell content — just the framework skeleton.
"""
from PIL import Image, ImageDraw, ImageFont
from pathlib import Path

OUT = Path("/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Miro_Visuals")

NAVY      = (44, 90, 136)
NAVY_DEEP = (33, 70, 108)
NAVY_DARK = (21, 42, 71)
TEAL      = (43, 107, 119)
GREEN     = (76, 140, 63)
AMBER     = (232, 139, 47)
RED       = (182, 59, 31)
LIGHT_BG  = (244, 245, 247)
GREY_TXT  = (75, 85, 99)
WHITE     = (255, 255, 255)
BORDER    = (216, 221, 227)

# 4K resolution for Miro zoom
W, H = 3840, 2160


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


def draw_centered_text(d, xy, text, font, color, anchor="mm"):
    d.text(xy, text, fill=color, font=font, anchor=anchor)


img = Image.new("RGB", (W, H), WHITE)
d = ImageDraw.Draw(img)

# Header bar
d.rectangle([0, 0, W, 180], fill=NAVY)
d.text((80, 44), "TARGET OPERATING MODEL — DIGITAL PRODUCTS",
       font=fnt(68, True), fill=WHITE)
d.text((80, 124), "Framework template · 4 stages × 5 dimensions",
       font=fnt(32), fill=(199, 211, 228))

stages = [
    ("DEMAND MANAGEMENT", "What demand looks like at intake"),
    ("PREPARATION & PLANNING", "How demand becomes a roadmap"),
    ("EXECUTION", "How work gets built"),
    ("DEPLOYMENT", "How releases land in production"),
]
dims = [
    ("CAPABILITIES", GREEN),
    ("ROLES", TEAL),
    ("TOOLS", AMBER),
    ("GOVERNANCE", NAVY),
    ("PERFORMANCE", RED),
]

# Layout — generous gaps between rows and columns
margin_x = 80
margin_top = 220
n_cols = 4
n_rows = 5
label_col_w = 320
gap_x = 50            # horizontal space between columns
gap_y = 36            # vertical space between rows
header_gap = 50       # gap below the stage headers
label_gap = 40        # gap between dim label and first cell

available_w = W - 2 * margin_x - label_col_w - label_gap - (n_cols - 1) * gap_x
cell_w = available_w // n_cols
available_h = H - margin_top - 140  # leave space for footer
header_h = 150
cell_h = (available_h - header_h - header_gap - (n_rows - 1) * gap_y) // n_rows

# Stage headers
header_y = margin_top
for i, (stage, sub) in enumerate(stages):
    x0 = margin_x + label_col_w + label_gap + i * (cell_w + gap_x)
    d.rounded_rectangle([x0, header_y, x0 + cell_w, header_y + header_h],
                        radius=16, fill=NAVY_DARK)
    draw_centered_text(d, (x0 + cell_w / 2, header_y + 50),
                       f"0{i+1} · {stage}", fnt(30, True), WHITE)
    draw_centered_text(d, (x0 + cell_w / 2, header_y + 105),
                       sub, fnt(22), (199, 211, 228))

# Dimension labels (left col) + empty cells
first_row_y = header_y + header_h + header_gap
for r, (dim, dim_color) in enumerate(dims):
    y0 = first_row_y + r * (cell_h + gap_y)
    # left label pill
    d.rounded_rectangle([margin_x, y0, margin_x + label_col_w, y0 + cell_h],
                        radius=16, fill=dim_color)
    draw_centered_text(d, (margin_x + label_col_w / 2, y0 + cell_h / 2),
                       dim, fnt(40, True), WHITE)

    for c, (stage, _) in enumerate(stages):
        x0 = margin_x + label_col_w + label_gap + c * (cell_w + gap_x)
        d.rounded_rectangle(
            [x0, y0, x0 + cell_w, y0 + cell_h],
            radius=18, fill=LIGHT_BG, outline=BORDER, width=3
        )

# Footer
d.text((margin_x, H - 60), "Avangrid · Digital Products · TOM template",
       font=fnt(22), fill=GREY_TXT)
d.text((W - 800, H - 60), "Fill-in workshop · Miro board",
       font=fnt(22, True), fill=AMBER)

OUT.mkdir(parents=True, exist_ok=True)
out_file = OUT / "00_TOM_BLANK_Template.png"
img.save(out_file, "PNG", dpi=(150, 150))
print(f"✓ Saved: {out_file}")
print(f"  Resolution: {W}×{H} (4K)")
