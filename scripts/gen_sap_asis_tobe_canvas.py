"""AS-IS / TO-BE Flow Canvas — Demand × Portfolio stages with guiding questions
and area/team rows for AS-IS mapping."""
from PIL import Image, ImageDraw, ImageFont
import os

# ── Colors (match Interview Canvas) ─────────────────────────────────────
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
INNER   = (8, 28, 52)
BORD    = (40, 60, 85)

ASIS_C  = (220, 60, 60)
TOBE_C  = (120, 190, 32)

def font(size, bold=False):
    for p in ["/System/Library/Fonts/Helvetica.ttc",
              "/System/Library/Fonts/SFNSText.ttf"]:
        if os.path.exists(p):
            try: return ImageFont.truetype(p, size, index=1 if bold and p.endswith(".ttc") else 0)
            except: return ImageFont.truetype(p, size)
    return ImageFont.load_default()

FT     = font(60, True)
FM     = font(30)
FL     = font(28, True)
FSM    = font(22)
FROW   = font(56, True)
FSTAGE = font(30, True)
FSTSUB = font(22)
FNUM   = font(38, True)
FQ     = font(20)       # questions
FQH    = font(20, True) # question header
FTEAM  = font(26, True) # team name label
FTEAMHINT = font(18)

W = 5600
M   = 50
GAP = 24
FW  = W - M*2

def rrect(d, x, y, w, h, fill=CARD, outline=None, r=14, width=2):
    d.rounded_rectangle([x, y, x+w, y+h], r, fill=fill,
                        outline=outline, width=width if outline else 0)


def wrap_lines(text, max_w, f, d):
    words = text.split()
    lines = []
    line = ""
    for w_ in words:
        t = f"{line} {w_}".strip()
        if d.textlength(t, font=f) <= max_w:
            line = t
        else:
            if line:
                lines.append(line)
            line = w_
    if line:
        lines.append(line)
    return lines


def generate():
    img = Image.new("RGB", (W, 6500), BG)
    d = ImageDraw.Draw(img)

    # ── TITLE ───────────────────────────────────────────────────────
    Y = M
    d.text((M, Y), "Demand × Portfolio Flow Canvas  —  SAP  (AS-IS)", fill=WHITE, font=FT)
    Y += 72
    d.text((M, Y), "AS-IS current state  |  SAP team (Kumar Jaladanki)  |  End-to-end stages with guiding questions",
           fill=DIM, font=FM)
    Y += 60

    # ═══════════════════════════════════════════════════════════════════
    # STAGE HEADER ROW — 5 stages
    # ═══════════════════════════════════════════════════════════════════
    stages = [
        ("01", "INTAKE",          "Channels & Sources",    CYAN, [
            "What channels does work arrive through today? Who submits, how often, and what type of work?",
            "Which channels are connected to existing governance meetings or routed directly to developers?",
            "Are there sources you don't control (regulatory mandates, external demands)?",
            "What tools/forms capture demand today? Is there an intake template?",
            "What seasonal or cyclical patterns spike demand (quarter-end, audits, budget cycles)?",
            "What's the average time from request submission to first acknowledgment?",
        ]),
        ("02", "CLASSIFICATION",  "Type, Size, Urgency",   PINK, [
            "How are requests classified today? (Type, priority, size, urgency)",
            "What % of requests arrive with clear business value and acceptance criteria?",
            "How often do 'urgent' requests bypass the normal process?",
            "How many requests are duplicates or overlap existing work — and why does it happen?",
        ]),
        ("03", "PRIORITIZATION",  "Scoring & Criteria",    ORANGE, [
            "What criteria are used to prioritize? How do you balance business value, cost, and risk?",
            "Who is responsible for prioritization decisions?",
            "How are conflicts escalated or resolved when priorities clash?",
            "How do dependencies impact prioritization and delivery timelines?",
        ]),
        ("04", "PORTFOLIO",       "Governance & Decision", BLUE, [
            "What types of initiatives belong in the portfolio (Run / Grow / Transform)?",
            "Is there a clear definition of what qualifies as a portfolio initiative?",
            "Do you have a centralized, up-to-date view of all active initiatives?",
            "What committees, forums, or decision-making bodies exist — and what's missing?",
            "How are rejected or deprioritized requests communicated back to requestors?",
            "How are initiatives funded and budgeted? Is funding fixed or flexible?",
        ]),
        ("05", "DELIVERY",        "Execution & Flow",      GREEN, [
            "How do you track progress across initiatives today?",
            "Where are the handoffs or dependencies that slow delivery the most?",
            "How are risks and issues identified, reported, and mitigated?",
            "How is delivery quality measured and enforced?",
        ]),
    ]

    ROW_LABEL_W = 280
    GRID_X = M + ROW_LABEL_W + 24
    GRID_W = FW - ROW_LABEL_W - 24
    N = len(stages)
    COL_W = (GRID_W - (N-1)*GAP) // N

    # Stage headers
    HDR_H = 140
    for i, (num, name, sub, color, _questions) in enumerate(stages):
        hx = GRID_X + i*(COL_W+GAP)
        rrect(d, hx, Y, COL_W, HDR_H, fill=INNER, outline=color, r=14, width=3)
        d.text((hx+22, Y+14), num, fill=color, font=FNUM)
        if i < N-1:
            ax = hx + COL_W + 2
            ay = Y + HDR_H//2
            d.polygon([(ax, ay-14), (ax+GAP-2, ay), (ax, ay+14)], fill=color)
        d.text((hx+22, Y+62), name, fill=WHITE, font=FSTAGE)
        d.text((hx+22, Y+100), sub, fill=DIM, font=FSTSUB)

    Y += HDR_H + 16

    # ═══════════════════════════════════════════════════════════════════
    # GUIDING QUESTIONS ROW — one card per stage
    # ═══════════════════════════════════════════════════════════════════
    # Empty label block on left
    Q_H = 520
    rrect(d, M, Y, ROW_LABEL_W, Q_H, fill=INNER, outline=BORD, r=16, width=2)
    d.text((M+24, Y+24), "GUIDING", fill=ORANGE, font=FL)
    d.text((M+24, Y+62), "QUESTIONS", fill=ORANGE, font=FL)
    d.line([M+24, Y+108, M+ROW_LABEL_W-24, Y+108], fill=BORD, width=2)
    d.text((M+24, Y+126), "Use as conversation", fill=DIM, font=FSM)
    d.text((M+24, Y+154), "starters — not a", fill=DIM, font=FSM)
    d.text((M+24, Y+182), "checklist", fill=DIM, font=FSM)

    for i, (num, name, sub, color, questions) in enumerate(stages):
        qx = GRID_X + i*(COL_W+GAP)
        rrect(d, qx, Y, COL_W, Q_H, fill=INNER, outline=color, r=12, width=2)
        # Color tag top
        d.rectangle([qx, Y, qx+COL_W, Y+6], fill=color)

        # Header label
        d.text((qx+20, Y+18), "GUIDING QUESTIONS", fill=color, font=FQH)

        # Questions
        qy = Y + 54
        for q in questions:
            lines = wrap_lines(q, COL_W-48, FQ, d)
            # Bullet
            d.ellipse([qx+20, qy+8, qx+28, qy+16], fill=color)
            for li, ln in enumerate(lines):
                d.text((qx+36, qy+li*26), ln, fill=DIM, font=FQ)
            qy += len(lines)*26 + 14

    Y += Q_H + 32

    # ═══════════════════════════════════════════════════════════════════
    # AS-IS ROW — with area/team rows
    # ═══════════════════════════════════════════════════════════════════
    N_TEAMS = 1
    TEAM_ROW_H = 760
    ROW_H = N_TEAMS * TEAM_ROW_H + 20

    # Row label block (left)
    rrect(d, M, Y, ROW_LABEL_W, ROW_H, fill=INNER, outline=ASIS_C, r=16, width=3)
    d.ellipse([M+32, Y+52, M+72, Y+92], fill=ASIS_C)
    d.text((M+32, Y+110), "AS-IS", fill=WHITE, font=FROW)
    d.text((M+32, Y+192), "Current state", fill=DIM, font=FL)
    d.line([M+32, Y+238, M+ROW_LABEL_W-32, Y+238], fill=BORD, width=2)
    d.text((M+32, Y+256), "One row per area /", fill=MUTED, font=FSM)
    d.text((M+32, Y+284), "team. Describe how", fill=MUTED, font=FSM)
    d.text((M+32, Y+312), "each topic works", fill=MUTED, font=FSM)
    d.text((M+32, Y+340), "today per stage.", fill=MUTED, font=FSM)

    # Big workspace
    rrect(d, GRID_X, Y, GRID_W, ROW_H, fill=INNER, outline=BORD, r=16, width=2)

    # Draw team rows
    for r in range(N_TEAMS):
        ry = Y + 10 + r*TEAM_ROW_H
        # Row separator (not on first)
        if r > 0:
            d.line([GRID_X+16, ry, GRID_X+GRID_W-16, ry], fill=BORD, width=2)
        # Team name placeholder at top-left of row — compact pill style
        label_text = "SAP  —  Kumar Jaladanki"
        lw = d.textlength(label_text, font=FTEAMHINT)
        pill_w = int(lw + 32)
        pill_h = 40
        nx = GRID_X + 24
        ny = ry + 20
        rrect(d, nx, ny, pill_w, pill_h, fill=BG, outline=ASIS_C, r=10, width=2)
        d.text((nx+16, ny+10), label_text, fill=ASIS_C, font=FTEAMHINT)

        # Vertical column separators inside workspace aligned with stage headers
        for i in range(1, N):
            sx = GRID_X + i*(COL_W+GAP) - GAP//2
            for yy in range(int(ry+10), int(ry+TEAM_ROW_H-10), 18):
                d.rectangle([sx-1, yy, sx+1, yy+7], fill=(30, 50, 75))

    Y += ROW_H + M

    img_out = img.crop((0, 0, W, Y))
    return img_out


import math as _math
base = "/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Miro Discovery Sessions"
os.makedirs(base, exist_ok=True)

result = generate()
# Miro cap at ~30MP
TARGET_PX = 28_000_000
if result.size[0] * result.size[1] > TARGET_PX:
    scale = _math.sqrt(TARGET_PX / (result.size[0] * result.size[1]))
    result = result.resize((int(result.size[0]*scale), int(result.size[1]*scale)), Image.LANCZOS)
out = f"{base}/SESSION_2_SAP_ASIS_TOBE_Flow_Canvas.png"
result.save(out, "PNG", optimize=True)
px = result.size[0] * result.size[1]
print(f"SAP: {result.size[0]}x{result.size[1]} — {px/1e6:.1f}MP → {os.path.basename(out)}")
