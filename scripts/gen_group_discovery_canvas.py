"""Group Discovery Canvas — 6 questions × 5 teams grid for live facilitation"""
from PIL import Image, ImageDraw, ImageFont
import os

# ── Colors (match Interview Canvas style) ───────────────────────────────
BG      = (2, 17, 30)
CARD    = (4, 22, 44)
WHITE   = (255, 255, 255)
DIM     = (160, 175, 195)
MUTED   = (100, 120, 145)
GREEN   = (120, 190, 32)
BLUE    = (0, 82, 155)
ORANGE  = (247, 148, 29)
RED     = (220, 60, 60)
PINK    = (200, 80, 180)
LIME    = (160, 210, 40)
CYAN    = (80, 180, 220)
PURPLE  = (150, 110, 210)
INNER   = (8, 28, 52)
BORD    = (40, 60, 85)
TIMER_BG = (30, 55, 85)


def font(size, bold=False):
    for p in ["/System/Library/Fonts/Helvetica.ttc",
              "/System/Library/Fonts/SFNSText.ttf"]:
        if os.path.exists(p):
            try: return ImageFont.truetype(p, size, index=1 if bold and p.endswith(".ttc") else 0)
            except: return ImageFont.truetype(p, size)
    return ImageFont.load_default()


FT     = font(60, True)
FM     = font(32)
FQNUM  = font(54, True)
FQTEXT = font(36, True)
FQSUB  = font(24)
FTEAM  = font(26, True)
FSMALL = font(22)
FTIMER = font(26, True)


W = 5600
M = 50
GAP = 40
INN = 24
FW = W - M*2


def rrect(d, x, y, w, h, fill=CARD, outline=None, r=16, width=2):
    if w <= 0 or h <= 0:
        return
    d.rounded_rectangle([x, y, x+w, y+h], r, fill=fill,
                        outline=outline, width=width if outline else 0)


def generate_canvas():
    questions = [
        ("01", "What are your top priorities today?",                 None, BLUE),
        ("02", "What is your current Role and Responsibilities?",     None, ORANGE),
        ("03", "Which tools do you use?",                              "Project management  ·  Vendors  ·  Development management", GREEN),
        ("04", "What are your main pain points today?",               None, RED),
        ("05", "How do you prioritize work today?",                   None, PINK),
        ("06", "How do you receive demand requests?",                 None, LIME),
    ]

    teams = [
        ("GIS",            CYAN),
        ("Corporate",      PURPLE),
        ("G&E",            GREEN),
        ("Global SAP",     BLUE),
        ("Smart Metering", ORANGE),
    ]

    # ── Block dimensions ───────────────────────────────────────────────
    Q_HDR_H   = 110
    TEAM_HDR_H = 56
    TEAM_BODY_H = 420
    BLOCK_H = Q_HDR_H + TEAM_HDR_H + TEAM_BODY_H + 30

    N_TEAMS = len(teams)
    COL_GAP = 16
    COL_W = (FW - INN*2 - COL_GAP * (N_TEAMS - 1)) // N_TEAMS

    total_h = M + 160 + 200 + len(questions) * (BLOCK_H + GAP) + 120 + M
    img = Image.new("RGB", (W, total_h), BG)
    d = ImageDraw.Draw(img)

    # ── TITLE ──────────────────────────────────────────────────────────
    Y = M
    d.text((M, Y), "Application Champions — Group Discovery", fill=WHITE, font=FT)
    Y += 72
    d.text((M, Y), "60 min", fill=DIM, font=FM)
    Y += 70

    # ── INTRODUCTIONS HEADER (5 min, no interaction) ───────────────────
    INTRO_H = 130
    rrect(d, M, Y, FW, INTRO_H, fill=INNER, outline=PURPLE, r=18)
    d.rectangle([M, Y, M + 10, Y + INTRO_H], fill=PURPLE)

    rrect(d, M + 28, Y + 32, 72, 66, fill=PURPLE, r=14)
    label = "00"
    nw = d.textlength(label, font=FQNUM)
    d.text((M + 28 + (72 - nw) // 2, Y + 30), label, fill=WHITE, font=FQNUM)

    d.text((M + 120, Y + 28), "Introductions", fill=WHITE, font=FQTEXT)
    d.text((M + 120, Y + 74),
           "5 min  ·  round the table — each person shares Name, Role and Area",
           fill=DIM, font=FQSUB)

    Y += INTRO_H + GAP

    # ── QUESTION BLOCKS ────────────────────────────────────────────────
    for num, qtext, qsub, qcolor in questions:
        # Block container
        rrect(d, M, Y, FW, BLOCK_H, fill=CARD, outline=BORD, r=18)

        # Question header bar
        rrect(d, M, Y, FW, Q_HDR_H, fill=INNER, outline=qcolor, r=18)
        # Accent stripe on the left
        d.rectangle([M, Y, M + 10, Y + Q_HDR_H], fill=qcolor)

        # Number badge
        rrect(d, M + 28, Y + 22, 72, 66, fill=qcolor, r=14)
        nw = d.textlength(num, font=FQNUM)
        d.text((M + 28 + (72 - nw) // 2, Y + 20), num, fill=WHITE, font=FQNUM)

        # Question text
        d.text((M + 120, Y + 22), qtext, fill=WHITE, font=FQTEXT)
        if qsub:
            d.text((M + 120, Y + 68), qsub, fill=DIM, font=FQSUB)

        # Team columns
        cy = Y + Q_HDR_H + 16
        for i, (team_name, team_color) in enumerate(teams):
            cx = M + INN + i * (COL_W + COL_GAP)

            # Team label pill (top)
            rrect(d, cx, cy, COL_W, TEAM_HDR_H, fill=team_color, r=10)
            tw = d.textlength(team_name, font=FTEAM)
            d.text((cx + (COL_W - tw) // 2, cy + 14), team_name, fill=WHITE, font=FTEAM)

            # Empty response area
            body_y = cy + TEAM_HDR_H + 8
            rrect(d, cx, body_y, COL_W, TEAM_BODY_H, fill=INNER, outline=team_color, r=12, width=2)

        Y += BLOCK_H + GAP

    # ── FOOTER / Timing ────────────────────────────────────────────────
    rrect(d, M, Y, FW, 90, fill=TIMER_BG, outline=BLUE, r=14)
    d.text((M + 24, Y + 14), "SESSION TIMING — 60 minutes total",
           fill=WHITE, font=FTEAM)
    d.text((M + 24, Y + 50),
           "Introductions (5 min)  ·  6 questions × ~8 min each  ·  round-robin across the 5 teams  ·  wrap-up (5 min)",
           fill=DIM, font=FSMALL)
    Y += 90 + M

    img_out = img.crop((0, 0, W, Y))
    return img_out


base = "/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Miro Discovery Sessions"
os.makedirs(base, exist_ok=True)
result = generate_canvas()
out = f"{base}/SESSION_1_Group_Discovery_Canvas.png"
result.save(out, "PNG")
px = result.size[0] * result.size[1]
print(f"{result.size[0]}x{result.size[1]} — {px/1e6:.1f}MP → {os.path.basename(out)}")
