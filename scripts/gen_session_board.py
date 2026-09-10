"""Session Board — Demand & Portfolio slides organized by project for Miro"""
from PIL import Image, ImageDraw, ImageFont
import os, glob

def font(size, bold=False):
    for p in ["/System/Library/Fonts/Helvetica.ttc", "/System/Library/Fonts/SFNSText.ttf"]:
        if os.path.exists(p):
            try: return ImageFont.truetype(p, size, index=1 if bold and p.endswith(".ttc") else 0)
            except: return ImageFont.truetype(p, size)
    return ImageFont.load_default()

FT    = font(72, True)
FM    = font(34)
FPROJ = font(48, True)
FSECT = font(36, True)
FCAP  = font(24, True)
FTAG  = font(28, True)
FL    = font(28)

BG      = (2, 17, 30)
WHITE   = (255, 255, 255)
DIM_TXT = (160, 175, 195)
MUTED   = (100, 120, 145)
BORD    = (40, 60, 85)

DEMAND_C   = (0, 150, 200)
PORTF_C    = (120, 190, 32)

PROJ_COLORS = {
    "PS":    ((120, 190, 32),  "P&S / P&T"),
    "DCoE":  ((0, 100, 180),   "Digital CoE"),
    "DTO":   ((247, 148, 29),  "Renewables DTO"),
    "Infra": ((160, 100, 200), "Infrastructure"),
    "SAPDP": ((200, 60, 60),   "SAP Digital Products (AS-IS)"),
}

SLIDE_DIR = "/tmp/slide_renders/board_demand_portfolio"
THUMB_W = 860
THUMB_H = 484
CAP_H = 36
M = 50
COL_GAP = 16
ROW_GAP = 20
SECTION_GAP = 30

# Collect slides per project per theme
projects_order = ["PS", "DCoE", "DTO", "Infra", "SAPDP"]
themes = ["demand", "portfolio"]
theme_labels = {"demand": "SESSION 2 — DEMAND SOURCES", "portfolio": "SESSION 3 — PORTFOLIO REVIEW"}
theme_colors = {"demand": DEMAND_C, "portfolio": PORTF_C}

# Parse files
project_slides = {}  # {(proj, theme): [(order, path, caption), ...]}
for f in sorted(os.listdir(SLIDE_DIR)):
    if not f.endswith(".png"): continue
    parts = f.split("_", 3)
    proj = parts[0]
    theme = parts[1]
    order = int(parts[2])
    # Extract caption from filename
    caption = parts[3].replace(".png","").replace("_"," ").replace("--","—").replace("plusUX","+ UX").replace("eq","=").replace("and","&")
    key = (proj, theme)
    if key not in project_slides:
        project_slides[key] = []
    project_slides[key].append((order, f"{SLIDE_DIR}/{f}", caption))

for k in project_slides:
    project_slides[k].sort()

# ═══════════════════════════════════════════════════════════════
# Layout: For each project, show demand slides then portfolio slides
# in horizontal rows. Each project gets a section.
# ═══════════════════════════════════════════════════════════════

# Calculate how many columns we need (max slides in a single theme for any project)
max_cols = 1
for k, slides in project_slides.items():
    max_cols = max(max_cols, len(slides))

# Use max 5 columns, wrap to next row if more
MAX_COLS = 5

def rows_needed(n):
    return max(1, (n + MAX_COLS - 1) // MAX_COLS)

# Calculate total height
total_h = M + 110  # title
for proj in projects_order:
    has_any = False
    for theme in themes:
        if (proj, theme) in project_slides:
            has_any = True
    if not has_any:
        continue
    total_h += 70  # project header
    for theme in themes:
        slides = project_slides.get((proj, theme), [])
        if not slides: continue
        total_h += 50  # theme label
        nr = rows_needed(len(slides))
        total_h += nr * (THUMB_H + CAP_H + ROW_GAP)
    total_h += SECTION_GAP

total_h += M

TOTAL_W = M * 2 + MAX_COLS * (THUMB_W + COL_GAP) - COL_GAP

img = Image.new("RGB", (TOTAL_W, total_h), BG)
d = ImageDraw.Draw(img)

def rrect(x, y, w, h, fill=(4,22,44), outline=None, r=12):
    d.rounded_rectangle([x,y,x+w,y+h], r, fill=fill, outline=outline, width=2 if outline else 0)

# ── TITLE ────────────────────────────────────────────────────
Y = M
d.text((M, Y), "Discovery Sessions — Reference Slides by Project", fill=WHITE, font=FT)
Y += 82
d.text((M, Y), "Slides from delivered projects mapped to Demand Sources (Session 2) & Portfolio Review (Session 3)", fill=DIM_TXT, font=FM)
Y += 50

# ── PER PROJECT ──────────────────────────────────────────────
for proj in projects_order:
    has_any = any((proj, t) in project_slides for t in themes)
    if not has_any:
        continue

    color, proj_name = PROJ_COLORS[proj]

    # Project header bar
    rrect(M, Y, TOTAL_W - M*2, 60, fill=color, r=12)
    tw = d.textlength(proj_name, font=FPROJ)
    d.text((M + 20, Y + 8), proj_name, fill=WHITE, font=FPROJ)
    Y += 70

    for theme in themes:
        slides = project_slides.get((proj, theme), [])
        if not slides: continue

        tc = theme_colors[theme]
        tl = theme_labels[theme]

        # Theme label
        rrect(M, Y, d.textlength(tl, font=FSECT) + 40, 40, fill=(*tc, 30), outline=tc, r=8)
        d.text((M + 16, Y + 4), tl, fill=tc, font=FSECT)
        Y += 50

        # Slides in grid
        for idx, (order, path, caption) in enumerate(slides):
            col = idx % MAX_COLS
            row = idx // MAX_COLS
            cx = M + col * (THUMB_W + COL_GAP)
            cy = Y + row * (THUMB_H + CAP_H + ROW_GAP)

            # Load and paste slide
            slide_img = Image.open(path)
            slide_img = slide_img.resize((THUMB_W, THUMB_H), Image.LANCZOS)
            rrect(cx-2, cy-2, THUMB_W+4, THUMB_H+4, fill=BG, outline=(*color, 100), r=4)
            img.paste(slide_img, (cx, cy))

            # Caption bar
            d.rectangle([cx, cy+THUMB_H, cx+THUMB_W, cy+THUMB_H+CAP_H], fill=(6, 16, 30))
            d.text((cx+8, cy+THUMB_H+6), caption[:80], fill=DIM_TXT, font=FCAP)

        nr = rows_needed(len(slides))
        Y += nr * (THUMB_H + CAP_H + ROW_GAP)

    Y += SECTION_GAP

# ── LEGEND ───────────────────────────────────────────────────
Y += 10
rrect(M, Y, TOTAL_W - M*2, 50, fill=(4,22,44), outline=BORD, r=10)
lx = M + 20
for label, col in [("Demand Sources", DEMAND_C), ("Portfolio Review", PORTF_C)]:
    d.ellipse([lx, Y+16, lx+16, Y+32], fill=col)
    d.text((lx+24, Y+12), label, fill=DIM_TXT, font=FL)
    lx += d.textlength(label, font=FL) + 60
d.text((lx + 40, Y+12), "KB-verified  |  PPT slide numbers  |  High-res for Miro zoom", fill=MUTED, font=FL)
Y += 50 + M

# ── SAVE ─────────────────────────────────────────────────────
img = img.crop((0, 0, TOTAL_W, Y))
out = "/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Discovery_Sessions_Reference_Slides.png"
img.save(out, "PNG", optimize=True)
px = img.size[0] * img.size[1]
print(f"Saved {img.size[0]}x{img.size[1]} — {px/1e6:.1f}MP → {os.path.basename(out)}")
