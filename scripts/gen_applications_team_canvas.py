"""Applications Team Discovery Canvas — grid with pre-positioned questions inside each cell"""
from PIL import Image, ImageDraw, ImageFont
import os
import math

# ── Colors ─────────────────────────────────────────────────────────────
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
YELLOW  = (240, 200, 40)
INNER      = (8, 28, 52)
BORD       = (40, 60, 85)
SLOT_BG    = (12, 34, 60)      # subtle slot background for sticky drop zones
SIDE_BG    = (14, 20, 40)


def font(size, bold=False):
    for p in ["/System/Library/Fonts/Helvetica.ttc",
              "/System/Library/Fonts/SFNSText.ttf"]:
        if os.path.exists(p):
            try: return ImageFont.truetype(p, size, index=1 if bold and p.endswith(".ttc") else 0)
            except: return ImageFont.truetype(p, size)
    return ImageFont.load_default()

FT    = font(56, True)
FM    = font(28)
FSTGN = font(22, True)
FSTG  = font(30, True)
FSTGS = font(20)
FCELLQ = font(17)           # question text inside each cell
FCELLQN = font(18, True)    # question number
FAREA = font(32, True)
FCHMP = font(22, True)
FBADGE = font(18, True)
FSLOT = font(16)

W     = 7600
M     = 50
GAP_COL = 16
GAP_ROW = 18
INN   = 14
SIDE_W  = 500
HEADER_H = 160
ROW_H = 760          # taller rows so 5 questions + sticky slots fit comfortably


def rrect(d, x, y, w, h, fill=CARD, outline=None, r=14, width=2):
    d.rounded_rectangle([x, y, x+w, y+h], r, fill=fill,
                        outline=outline, width=width if outline else 0)


def wrap_text(d, text, font_, max_w):
    words = text.split()
    lines = []
    line = ""
    for w_ in words:
        test = line + (" " if line else "") + w_
        if d.textlength(test, font=font_) > max_w:
            if line:
                lines.append(line)
            line = w_
        else:
            line = test
    if line:
        lines.append(line)
    return lines


def generate_canvas():
    stages = [
        ("01", "CAPACITY",            "Headcount · Mix · Utilization",     ORANGE, [
            "How many FTEs per sub-team (internal · SWF · Accenture · vendor)?",
            "What % of the team is on BAU/incidents vs. strategic work?",
            "What is the average utilization % over the last 3 months?",
            "How many open roles? Average time-to-fill?",
            "Who are the key-person risks (concentrated knowledge)?",
        ]),
        ("02", "RESOURCE MIX & COST",  "Economics behind staffing choices", PINK, [
            "Cost % across internal · SWF · Accenture · Vendor?",
            "Who decides SWF vs. hire internal vs. vendor?",
            "Do you have burn-rate visibility (monthly · per project)?",
            "Average cost per ticket · PDD · release?",
            "If budget cut 20% tomorrow, what goes first and why?",
        ]),
        ("03", "REPORTING INVENTORY",  "USED · ORNAMENTAL · MISSING",        BLUE, [
            "Which reports / dashboards do you generate or receive?",
            "Who consumes each (Carlos · Vitor · business · auditor)?",
            "Where does the data live (Jira · IT Now · PDD · Excel · SP)?",
            "Which report does your team produce but NO ONE reads?",
            "Data you USE but don't report — or REPORT but don't use?",
        ]),
        ("04", "METRICS THAT MATTER",  "MUST · SHOULD · COULD · WON'T",     LIME, [
            "Which 3 KPIs are most important for your sub-team today?",
            "Do you measure Lead Time · Cycle Time · Throughput · WIP?",
            "Which current metric does NOT reflect reality?",
            "If only 1 metric to the board — which one and why?",
            "Which metric links Capacity (col 01) to Reporting (col 03)?",
        ]),
        ("05", "SYNERGY SCENARIOS",    "Stress-test Capacity × Reporting",   RED, [
            "A — Regulatory spike today: how fast do you see capacity impact?",
            "B — Manager asks 2 FTEs in 48h: which data exists? Which is missing?",
            "C — Cost per sub-team this quarter: how fast? How precise?",
            "Which scenario exposed the largest ideal-vs-reality gap?",
            "Smallest reporting change that solves most of these gaps?",
        ]),
    ]

    areas = [
        ("GIS",            "George Porto",                       GREEN),
        ("Corporate",      "Amit Patel  ·  Sarath Garikapati",   ORANGE),
        ("G&E",            "Andrew McGibbon",                    PINK),
        ("Global SAP",     "Michael Dubois  ·  Vanessa Colombo", BLUE),
        ("Smart Metering", "Iram Chowdhury",                     LIME),
    ]

    n_cols = len(stages)
    n_rows = len(areas)

    FW = W - M*2
    grid_w = FW - SIDE_W - GAP_COL
    col_w = (grid_w - GAP_COL * (n_cols - 1)) // n_cols
    grid_h = n_rows * ROW_H + GAP_ROW * (n_rows - 1)
    total_h = M + 170 + HEADER_H + GAP_ROW + grid_h + M

    img = Image.new("RGB", (W, total_h), BG)
    d = ImageDraw.Draw(img)

    Y = M
    d.text((M, Y), "Applications Team  —  AS-IS Discovery Canvas", fill=WHITE, font=FT)
    Y += 74
    d.text((M, Y),
           "Carlos Garbin's branch  |  ~90 min  |  Capacity & Resource Model  \u00d7  Reporting & Metrics Review  |  Place one sticky per question",
           fill=DIM, font=FM)
    Y += 60

    # ── COMPACT STAGE HEADER ROW ──
    header_y = Y
    # Left sidebar header (just label)
    rrect(d, M, header_y, SIDE_W, HEADER_H, fill=SIDE_BG, outline=BORD, r=14)
    d.rectangle([M, header_y, M+8, header_y+HEADER_H], fill=YELLOW)
    d.text((M+22, header_y+20), "DISCOVERY", fill=YELLOW, font=FSTG)
    d.text((M+22, header_y+60), "STAGES  \u2192", fill=YELLOW, font=FSTG)
    d.text((M+22, header_y+110),
           "Each stage has 5 questions per area below",
           fill=DIM, font=FSTGS)

    # 5 stage columns header
    cx = M + SIDE_W + GAP_COL
    for i, (num, stitle, ssub, color, _) in enumerate(stages):
        col_x = cx + i * (col_w + GAP_COL)
        rrect(d, col_x, header_y, col_w, HEADER_H, fill=INNER, outline=color, r=14)
        d.rectangle([col_x, header_y, col_x+6, header_y+HEADER_H], fill=color)
        d.text((col_x+20, header_y+20), num, fill=WHITE, font=FSTGN)
        d.text((col_x+20, header_y+46), stitle, fill=WHITE, font=FSTG)
        d.text((col_x+20, header_y+98), ssub, fill=DIM, font=FSTGS)

    Y = header_y + HEADER_H + GAP_ROW

    # ── GRID ROWS ──
    grid_top = Y

    for r, (area, champ, color) in enumerate(areas):
        ry = grid_top + r * (ROW_H + GAP_ROW)

        # Left sidebar: area + champion
        rrect(d, M, ry, SIDE_W, ROW_H, fill=SIDE_BG, outline=BORD, r=14)
        d.rectangle([M, ry, M+10, ry+ROW_H], fill=color)

        # AREA badge
        rrect(d, M+20, ry+22, 130, 32, fill=color, r=6)
        d.text((M+32, ry+28), f"AREA  #{r+1}", fill=WHITE, font=FBADGE)

        # Area name
        d.text((M+22, ry+74), area, fill=color, font=FAREA)

        # Champion
        d.text((M+22, ry+140), "Champion", fill=MUTED, font=FCHMP)
        # Wrap long champion line
        ch_lines = wrap_text(d, champ, FCHMP, SIDE_W - 44)
        for ci, cl in enumerate(ch_lines):
            d.text((M+22, ry+170 + ci*30), cl, fill=WHITE, font=FCHMP)

        # 5 cells for this row — each with 5 pre-positioned questions
        cx = M + SIDE_W + GAP_COL
        for i, (_, _, _, stage_color, questions) in enumerate(stages):
            cell_x = cx + i * (col_w + GAP_COL)
            rrect(d, cell_x, ry, col_w, ROW_H, fill=INNER, outline=BORD, r=14)
            d.rectangle([cell_x, ry, cell_x+col_w, ry+5], fill=stage_color)

            # Divide cell vertically into 5 question slots
            qpad = 10
            slot_gap = 8
            available_h = ROW_H - 16 - 4 * slot_gap
            slot_h = available_h // 5

            for qi, q in enumerate(questions[:5]):
                sy = ry + 12 + qi * (slot_h + slot_gap)
                # Slot background (very subtle)
                rrect(d, cell_x + qpad, sy, col_w - 2*qpad, slot_h,
                      fill=SLOT_BG, outline=BORD, r=8, width=1)

                # Question number badge
                nb_x = cell_x + qpad + 8
                nb_y = sy + 8
                rrect(d, nb_x, nb_y, 26, 26, fill=stage_color, r=6)
                nb_label = f"Q{qi+1}"
                nbw = d.textlength(nb_label, font=FCELLQN)
                d.text((nb_x + (26 - nbw)//2, nb_y + 3), nb_label, fill=WHITE, font=FCELLQN)

                # Question text (wrapped)
                q_x = nb_x + 36
                q_max_w = col_w - 2*qpad - 44 - 8
                q_lines = wrap_text(d, q, FCELLQ, q_max_w)
                for li, ln in enumerate(q_lines[:3]):
                    d.text((q_x, sy + 10 + li*20), ln, fill=WHITE, font=FCELLQ)

                # Sticky drop hint on the right of the slot
                hint = "\u2190 place sticky"
                hw = d.textlength(hint, font=FSLOT)
                d.text((cell_x + col_w - qpad - hw - 12, sy + slot_h - 24),
                       hint, fill=MUTED, font=FSLOT)

    return img.crop((0, 0, W, total_h))


base = "/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Miro Discovery Sessions"
os.makedirs(base, exist_ok=True)

result = generate_canvas()

TARGET_PX = 28_000_000
if result.size[0] * result.size[1] > TARGET_PX:
    scale = math.sqrt(TARGET_PX / (result.size[0] * result.size[1]))
    new_size = (int(result.size[0] * scale), int(result.size[1] * scale))
    result = result.resize(new_size, Image.LANCZOS)

out = f"{base}/SESSION_1_Applications_Team_Dynamic_Canvas.png"
result.save(out, "PNG", optimize=True)
px = result.size[0] * result.size[1]
file_mb = os.path.getsize(out) / 1e6
print(f"[Applications_Team_Dynamic] {result.size[0]}x{result.size[1]} \u2014 {px/1e6:.1f}MP \u2014 {file_mb:.1f}MB \u2192 {os.path.basename(out)}")
