"""
Builds blank single-stage TOM boards (one PNG per stage).
Stages 01..04, same big format as stage 01, ready for Miro fill-in.
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

W, H = 2800, 4200

STAGES = [
    ("01", "DEMAND MANAGEMENT",
     "What demand looks like at intake — capabilities, roles, tools, governance, performance"),
    ("02", "PREPARATION & PLANNING",
     "How demand becomes a roadmap — sizing, prioritization, sequencing, commitments"),
    ("03", "EXECUTION",
     "How work gets built — squads, sprints, engineering, quality, risk control"),
    ("04", "DEPLOYMENT",
     "How releases land in production — change, cutover, hypercare, value realization"),
]

DIMS = [
    ("CAPABILITIES", "What the org must be able to do",   GREEN),
    ("ROLES",        "Who owns and operates this",         TEAL),
    ("TOOLS",        "Systems and platforms used",         AMBER),
    ("GOVERNANCE",   "Decisions, forums, controls",        NAVY),
    ("PERFORMANCE",  "How success is measured",            RED),
]


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


def build_stage(num, title, sub, idx_total=4):
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)

    # Header
    d.rectangle([0, 0, W, 280], fill=NAVY_DARK)
    d.text((100, 60), f"STAGE {num}", font=fnt(48, True), fill=AMBER)
    d.text((100, 130), title, font=fnt(96, True), fill=WHITE)
    d.text((100, 230), sub, font=fnt(34), fill=(199, 211, 228))

    # Layout
    margin_x  = 120
    top_y     = 360
    label_w   = 520
    gap_label = 70
    gap_row   = 70
    footer_h  = 120

    n = len(DIMS)
    available_h = H - top_y - footer_h - (n - 1) * gap_row
    cell_h      = available_h // n
    cell_x0     = margin_x + label_w + gap_label
    cell_x1     = W - margin_x

    for r, (dim, ds, color) in enumerate(DIMS):
        y0 = top_y + r * (cell_h + gap_row)
        d.rounded_rectangle([margin_x, y0, margin_x + label_w, y0 + cell_h],
                            radius=24, fill=color)
        centered(d, (margin_x + label_w / 2, y0 + cell_h / 2 - 30),
                 dim, fnt(56, True), WHITE)
        centered(d, (margin_x + label_w / 2, y0 + cell_h / 2 + 38),
                 ds, fnt(24), (235, 240, 245))

        d.rounded_rectangle([cell_x0, y0, cell_x1, y0 + cell_h],
                            radius=28, fill=LIGHT_BG, outline=BORDER, width=4)

    # Footer
    d.text((margin_x, H - 80),
           f"Avangrid · Digital Products · Target Operating Model — Stage {num} {title.title()}",
           font=fnt(26), fill=GREY_TXT)
    d.text((W - 720, H - 80),
           f"Workshop board · {int(num)} of {idx_total}",
           font=fnt(26, True), fill=AMBER)

    safe = title.replace(" & ", "And").replace(" ", "")
    out_file = OUT / f"{num}_TOM_{safe}_BLANK.png"
    img.save(out_file, "PNG", dpi=(150, 150))

    # preview
    ratio = 1200 / W
    preview = img.resize((1200, int(H * ratio)), Image.LANCZOS)
    preview.save(out_file.with_name(out_file.stem + "_preview.png"), "PNG")
    return out_file


OUT.mkdir(parents=True, exist_ok=True)
for num, title, sub in STAGES:
    out = build_stage(num, title, sub)
    print(f"✓ {out.name}")
