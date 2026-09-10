"""SAP Team AS-IS Discovery Canvas — same layout as Applications Team
(Guiding Questions as bullets in header + single team row: Kumar SAP)."""
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
INNER   = (8, 28, 52)
BORD    = (40, 60, 85)
SIDE_BG = (14, 20, 40)


def font(size, bold=False):
    for p in ["/System/Library/Fonts/Helvetica.ttc",
              "/System/Library/Fonts/SFNSText.ttf"]:
        if os.path.exists(p):
            try: return ImageFont.truetype(p, size, index=1 if bold and p.endswith(".ttc") else 0)
            except: return ImageFont.truetype(p, size)
    return ImageFont.load_default()

FT      = font(56, True)
FM      = font(26)
FSTGN   = font(22, True)    # stage number
FSTG    = font(30, True)    # stage title
FSTGS   = font(18)          # stage subtitle
FGQH    = font(16, True)    # "GUIDING QUESTIONS" header pill
FGQ     = font(17)          # guiding question bullet text
FAREA   = font(32, True)
FCHMP   = font(22, True)
FBADGE  = font(18, True)
FSIDE   = font(26, True)
FSIDESUB = font(18)

W       = 7600
M       = 50
GAP_COL = 16
GAP_ROW = 20
INN     = 14
SIDE_W  = 480
HEADER_H = 490        # just enough for stage title + 5 GUIDING QUESTIONS bullets
ROW_H   = 900         # single team row (tall empty drop zone)


def rrect(d, x, y, w, h, fill=CARD, outline=None, r=14, width=2):
    d.rounded_rectangle([x, y, x+w, y+h], r, fill=fill,
                        outline=outline, width=width if outline else 0)


def wrap_text(d, text, font_, max_w):
    words = text.split()
    lines = []
    line = ""
    for w_ in words:
        t = (line + " " + w_).strip() if line else w_
        if d.textlength(t, font=font_) <= max_w:
            line = t
        else:
            if line:
                lines.append(line)
            line = w_
    if line:
        lines.append(line)
    return lines


def generate_canvas():
    stages = [
        ("01", "CAPACITY",            "Headcount \u00b7 Mix \u00b7 Utilization",     ORANGE, [
            "How many FTEs per sub-team (internal \u00b7 SWF \u00b7 Accenture \u00b7 vendor)?",
            "What % of the team is on BAU/incidents vs. strategic work?",
            "What is the average utilization % over the last 3 months?",
            "How many open roles? Average time-to-fill?",
            "Who are the key-person risks (concentrated knowledge)?",
        ]),
        ("02", "RESOURCE MIX & COST", "Economics behind staffing choices", PINK, [
            "Cost % across internal \u00b7 SWF \u00b7 Accenture \u00b7 Vendor?",
            "Who decides SWF vs. hire internal vs. vendor, and on what basis?",
            "Do you have burn-rate visibility (monthly \u00b7 per project)?",
            "Average cost per ticket \u00b7 PDD \u00b7 release?",
            "If budget cut 20% tomorrow, what would go first and why?",
        ]),
        ("03", "REPORTING INVENTORY", "What is USED \u00b7 ORNAMENTAL \u00b7 MISSING", BLUE, [
            "Which reports / dashboards do you generate or receive today?",
            "Who consumes each (Carlos \u00b7 Vitor \u00b7 business \u00b7 auditor) and cadence?",
            "Where does the data live (Jira \u00b7 IT Now \u00b7 PDD \u00b7 Excel \u00b7 SP)?",
            "Which report does your team produce but NO ONE reads?",
            "What data do you USE but don't report \u2014 or REPORT but don't use?",
        ]),
        ("04", "METRICS THAT MATTER", "MUST \u00b7 SHOULD \u00b7 COULD \u00b7 WON'T", LIME, [
            "Which 3 KPIs are most important for your sub-team today?",
            "Do you measure Lead Time \u00b7 Cycle Time \u00b7 Throughput \u00b7 WIP \u00b7 rework?",
            "Which current metric do you know does NOT reflect reality?",
            "If only 1 metric to the board \u2014 which one and why?",
            "Which metric links Capacity (col 01) to Reporting (col 03)?",
        ]),
        ("05", "SYNERGY SCENARIOS",   "Stress-test Capacity \u00d7 Reporting", RED, [
            "Scenario A \u2014 A regulatory spike hits today. How long until you see the capacity impact, and which report alerts you?",
            "Scenario B \u2014 Your manager asks for 2 more FTEs with 48h to justify. What capacity data do you already have, and what would be missing?",
            "Scenario C \u2014 Carlos asks: \"How much did your sub-team cost this quarter?\" How fast can you answer, and with what precision?",
            "Across the three scenarios, which exposed the largest gap between answer you'd WANT to give and the answer you CAN give today?",
            "What is the smallest reporting change that would resolve most of those gaps at once?",
        ]),
    ]

    areas = [
        ("Kumar SAP", BLUE),
    ]

    FW = W - M*2
    grid_w = FW - SIDE_W - GAP_COL
    n_cols = len(stages)
    col_w = (grid_w - GAP_COL * (n_cols - 1)) // n_cols

    total_h = M + 170 + HEADER_H + GAP_ROW + ROW_H + M
    img = Image.new("RGB", (W, total_h), BG)
    d = ImageDraw.Draw(img)

    Y = M
    d.text((M, Y), "SAP Team  —  AS-IS Discovery Canvas", fill=WHITE, font=FT)
    Y += 74
    d.text((M, Y),
           "Customer Services / SAP branch (Kumar Jaladanki)  |  ~90 min  |  Capacity & Resource Model  \u00d7  Reporting & Metrics Review  |  Grid: 5 stages \u00d7 1 team",
           fill=DIM, font=FM)
    Y += 60

    # ── HEADER ROW (with GUIDING QUESTIONS bullets) ──
    header_y = Y
    # Sidebar: GUIDING QUESTIONS label
    rrect(d, M, header_y, SIDE_W, HEADER_H, fill=SIDE_BG, outline=BORD, r=14)
    d.rectangle([M, header_y, M+8, header_y+HEADER_H], fill=YELLOW)
    d.text((M+22, header_y+24), "GUIDING", fill=YELLOW, font=FSIDE)
    d.text((M+22, header_y+58), "QUESTIONS", fill=YELLOW, font=FSIDE)
    sub_lines = wrap_text(d,
        "Use as conversation starters when populating each cell below. Not a checklist \u2014 pick what resonates with your area.",
        FSIDESUB, SIDE_W - 44)
    for i, ln in enumerate(sub_lines):
        d.text((M+22, header_y+110+i*24), ln, fill=DIM, font=FSIDESUB)

    # 5 stage columns
    cx = M + SIDE_W + GAP_COL
    for i, (num, stitle, ssub, color, questions) in enumerate(stages):
        col_x = cx + i * (col_w + GAP_COL)
        rrect(d, col_x, header_y, col_w, HEADER_H, outline=BORD, r=14)
        # Title block (top ~90)
        rrect(d, col_x, header_y, col_w, 90, fill=INNER, outline=color, r=14)
        d.rectangle([col_x, header_y+72, col_x+col_w, header_y+90], fill=INNER)
        d.rectangle([col_x, header_y, col_x+6, header_y+HEADER_H], fill=color)
        d.text((col_x+22, header_y+14), num, fill=WHITE, font=FSTGN)
        d.text((col_x+22, header_y+40), stitle, fill=WHITE, font=FSTG)
        d.text((col_x+22, header_y+80), ssub, fill=DIM, font=FSTGS)

        # GUIDING QUESTIONS pill header
        gq_y = header_y + 108
        gq_w = col_w - 32
        rrect(d, col_x+16, gq_y, gq_w, 28, fill=color, r=6)
        d.text((col_x+28, gq_y+4), "GUIDING QUESTIONS", fill=WHITE, font=FGQH)

        # Bullet questions
        qy = gq_y + 42
        for q in questions:
            # Bullet dot
            d.ellipse([col_x+22, qy+6, col_x+32, qy+16], fill=color)
            lines = wrap_text(d, q, FGQ, gq_w - 28)
            for li, ln in enumerate(lines):
                d.text((col_x+38, qy + li*22), ln, fill=DIM, font=FGQ)
            qy += len(lines) * 22 + 12

    Y = header_y + HEADER_H + GAP_ROW

    # ── SINGLE TEAM ROW (empty drop zones) ──
    for r, (area, color) in enumerate(areas):
        ry = Y + r * (ROW_H + GAP_ROW)

        # Sidebar: area name
        rrect(d, M, ry, SIDE_W, ROW_H, fill=SIDE_BG, outline=BORD, r=14)
        d.rectangle([M, ry, M+10, ry+ROW_H], fill=color)

        # AREA pill
        rrect(d, M+22, ry+22, 120, 32, fill=color, r=6)
        d.text((M+36, ry+28), "AREA  #1", fill=WHITE, font=FBADGE)

        # Area name (large, vertically centered)
        d.text((M+22, ry + ROW_H//2 - 20), area, fill=color, font=FAREA)

        # 5 empty cells
        cx = M + SIDE_W + GAP_COL
        for i, (_, _, _, stage_color, _) in enumerate(stages):
            cell_x = cx + i * (col_w + GAP_COL)
            rrect(d, cell_x, ry, col_w, ROW_H, fill=INNER, outline=BORD, r=14)
            d.rectangle([cell_x, ry, cell_x+col_w, ry+5], fill=stage_color)
            d.text((cell_x+16, ry+14), "DROP STICKIES HERE", fill=MUTED, font=FBADGE)

    return img.crop((0, 0, W, Y + ROW_H + M))


base = "/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Miro Discovery Sessions"
os.makedirs(base, exist_ok=True)

result = generate_canvas()

TARGET_PX = 28_000_000
if result.size[0] * result.size[1] > TARGET_PX:
    scale = math.sqrt(TARGET_PX / (result.size[0] * result.size[1]))
    new_size = (int(result.size[0] * scale), int(result.size[1] * scale))
    result = result.resize(new_size, Image.LANCZOS)

out = f"{base}/SESSION_1_SAP_Team_Dynamic_Canvas.png"
result.save(out, "PNG", optimize=True)
px = result.size[0] * result.size[1]
file_mb = os.path.getsize(out) / 1e6
print(f"[SAP_Team_Dynamic] {result.size[0]}x{result.size[1]} \u2014 {px/1e6:.1f}MP \u2014 {file_mb:.1f}MB \u2192 {os.path.basename(out)}")
