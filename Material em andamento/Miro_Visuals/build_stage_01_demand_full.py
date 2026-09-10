"""
Stage 01 — DEMAND MANAGEMENT (full topic checklist per dimension)
Comprehensive board: every topic to be defined gets its own titled card + writable area.
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
GREY_TXT  = (75, 85, 99)
HINT_TXT  = (140, 150, 165)
WHITE     = (255, 255, 255)
SLOT_BG   = (255, 255, 255)
SLOT_BORDER = (200, 208, 220)

DIMS = [
    ("CAPABILITIES",  GREEN,  5, [  # cols=5
        "Demand intake channels (formal & informal)",
        "Demand registration and logging",
        "Channel consolidation (single vs multiple entry points)",
        "Standard intake formats / templates",
        "Demand classification (type, size, urgency)",
        "Differentiation: run / change / regulatory / innovation",
        "Business value definition",
        "Availability of acceptance criteria",
        "Technical feasibility input",
        "Triage and initial validation activities",
        "Demand routing and assignment",
        "Handling of urgent or exception requests",
        "Demand visibility and tracking",
        "Identification of duplicate / overlapping demand",
    ]),
    ("ROLES",         TEAL,   4, [
        "Business requestors / stakeholders",
        "Product or demand managers",
        "Architects or technical reviewers",
        "Triage / intake committee",
    ]),
    ("TOOLS",         AMBER,  5, [
        "Demand capture tools",
        "Demand tracking & backlog tools",
        "Communication channels",
        "Documentation tools (request details)",
        "Reporting / visibility tools",
    ]),
    ("GOVERNANCE",    NAVY,   3, [
        "Demand review forums",
        "Decision-making practices (accept / reject / redirect)",
        "Escalation mechanisms (urgent / conflicting requests)",
        "Definition of intake guidelines or policies",
        "Alignment mechanisms with business stakeholders",
        "Oversight of demand inflow and backlog health",
    ]),
    ("PERFORMANCE",   RED,    3, [
        "Demand volume and inflow trends",
        "Demand quality (completeness, clarity)",
        "Triage lead time / time-of-demand",
        "% demand entering through formal channels",
        "% urgent or bypassed requests",
        "Demand duplication / overlap indicators",
        "Visibility and traceability of demand status",
    ]),
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


# ── canvas sizing ──────────────────────────────────────────────────────────
W = 3600
margin_x = 120
header_h = 340          # top page header
dim_header_h = 140      # per-dimension banner
gap_between_dims = 70
gap_card = 26
card_h = 460            # full card height (title + writable)
title_band_h = 110      # colored title strip per card
footer_h = 130

# compute total height
total_card_h = 0
for _, _, cols, items in DIMS:
    rows = (len(items) + cols - 1) // cols
    total_card_h += dim_header_h + rows * card_h + (rows - 1) * gap_card

H = header_h + total_card_h + (len(DIMS) - 1) * gap_between_dims + footer_h

img = Image.new("RGB", (W, H), WHITE)
d = ImageDraw.Draw(img)

# ── page header ────────────────────────────────────────────────────────────
d.rectangle([0, 0, W, header_h], fill=NAVY_DARK)
d.text((margin_x, 70), "STAGE 01", font=fnt(54, True), fill=AMBER)
d.text((margin_x, 145), "DEMAND MANAGEMENT", font=fnt(110, True), fill=WHITE)
d.text((margin_x, 270), "Topics to define for the product-centric TOM — fill each card during the workshop",
       font=fnt(34), fill=(199, 211, 228))

# ── dimensions ─────────────────────────────────────────────────────────────
y = header_h + gap_between_dims
for idx, (dim_name, color, cols, items) in enumerate(DIMS):
    rows = (len(items) + cols - 1) // cols

    # dimension banner
    d.rounded_rectangle([margin_x, y, W - margin_x, y + dim_header_h],
                        radius=24, fill=color)
    centered(d, (W / 2, y + dim_header_h / 2),
             dim_name, fnt(72, True), WHITE)

    cards_y0 = y + dim_header_h + gap_card
    avail_w = W - 2 * margin_x
    card_w = (avail_w - (cols - 1) * gap_card) // cols

    for i, item in enumerate(items):
        r = i // cols
        c = i % cols
        cx0 = margin_x + c * (card_w + gap_card)
        cy0 = cards_y0 + r * (card_h + gap_card)
        cx1 = cx0 + card_w
        cy1 = cy0 + card_h

        # outer card
        d.rounded_rectangle([cx0, cy0, cx1, cy1], radius=22,
                            fill=SLOT_BG, outline=SLOT_BORDER, width=3)
        # title band (top of card)
        d.rounded_rectangle([cx0, cy0, cx1, cy0 + title_band_h], radius=22,
                            fill=color, outline=color, width=0)
        # mask bottom corners of the band
        d.rectangle([cx0, cy0 + title_band_h - 22, cx1, cy0 + title_band_h], fill=color)

        # title text (wrapped)
        max_chars = max(8, (card_w - 30) // 14)
        lines = wrap_text(item, max_chars)
        # fit in title band; use up to 3 lines
        lines = lines[:3]
        line_height = 30
        block_h = len(lines) * line_height
        start_y = cy0 + (title_band_h - block_h) / 2 + line_height / 2
        for li, line in enumerate(lines):
            centered(d, ((cx0 + cx1) / 2, start_y + li * line_height),
                     line, fnt(22, True), WHITE)

    y = cards_y0 + rows * card_h + (rows - 1) * gap_card + gap_between_dims

# ── footer ─────────────────────────────────────────────────────────────────
d.text((margin_x, H - 80),
       "Avangrid · Digital Products · Target Operating Model — Stage 01 Demand Management (full topic list)",
       font=fnt(26), fill=GREY_TXT)
d.text((W - 800, H - 80),
       "Workshop board · 1 of 4 · full",
       font=fnt(26, True), fill=AMBER)

OUT.mkdir(parents=True, exist_ok=True)
out_file = OUT / "01_TOM_DemandManagement_FULL.png"
img.save(out_file, "PNG", dpi=(150, 150))

# preview
ratio = 1400 / W
preview = img.resize((1400, int(H * ratio)), Image.LANCZOS)
preview.save(out_file.with_name(out_file.stem + "_preview.png"), "PNG")
print(f"✓ Saved: {out_file}")
print(f"  Resolution: {W}×{H}")
