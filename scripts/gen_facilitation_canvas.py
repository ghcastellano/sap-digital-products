"""Demand × Portfolio Discovery Canvas — matches Stakeholder Interview style"""
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
ROW_DARK   = (10, 35, 60)
ROW_LIGHT  = (6, 28, 50)
INNER      = (8, 28, 52)
BORD       = (40, 60, 85)
TIMER_BG   = (30, 55, 85)

def font(size, bold=False):
    for p in ["/System/Library/Fonts/Helvetica.ttc",
              "/System/Library/Fonts/SFNSText.ttf"]:
        if os.path.exists(p):
            try: return ImageFont.truetype(p, size, index=1 if bold and p.endswith(".ttc") else 0)
            except: return ImageFont.truetype(p, size)
    return ImageFont.load_default()

FT  = font(60, True)
FM  = font(32)
FS  = font(44, True)
FSB = font(28)
FL  = font(32, True)
FB  = font(28)
FTH = font(28, True)
FTD = font(26)
FSM = font(24)
FTIMER = font(26, True)
FP  = font(30, True)

W = 5600
M   = 50
GAP = 40
INN = 24
FW  = W - M*2


def rrect(d, x, y, w, h, fill=CARD, outline=None, r=18):
    d.rounded_rectangle([x, y, x+w, y+h], r, fill=fill,
                        outline=outline, width=2 if outline else 0)

def hdr(d, x, y, text, color, sub=None):
    d.ellipse([x, y+6, x+26, y+32], fill=color)
    d.text((x+36, y), text, fill=WHITE, font=FS)
    y += 52
    if sub:
        d.text((x+36, y), sub, fill=DIM, font=FSB)
        y += 36
    return y + 10

def inbox(d, x, y, w, h, label=None):
    rrect(d, x, y, w, h, fill=INNER, outline=BORD, r=12)
    if label:
        d.text((x+16, y+12), label, fill=DIM, font=FSM)

def timer_badge(d, block_x, block_y, block_w, minutes):
    text = f"⏱  {minutes} min"
    tw = d.textlength(text, font=FTIMER)
    bw = int(tw + 30)
    bh = 36
    bx = block_x + block_w - bw - INN
    by = block_y + INN
    rrect(d, bx, by, bw, bh, fill=TIMER_BG, outline=BLUE, r=8)
    d.text((bx+15, by+6), text, fill=WHITE, font=FTIMER)


def generate_canvas():
    img = Image.new("RGB", (W, 8000), BG)
    d = ImageDraw.Draw(img)

    # ═══════════════════════════════════════════════════════════════════
    # TITLE
    # ═══════════════════════════════════════════════════════════════════
    Y = M
    d.text((M, Y), "Demand × Portfolio Discovery Canvas", fill=WHITE, font=FT)
    Y += 70
    d.text((M, Y), "Group Discovery  |  90 min  |  Digital Products Team",
           fill=DIM, font=FM)
    Y += 60

    # ═══════════════════════════════════════════════════════════════════
    # BLOCK 0 — CONTEXT & OPENING (5 min)
    # ═══════════════════════════════════════════════════════════════════
    b0_h = 320
    rrect(d, M, Y, FW, b0_h, outline=BORD)
    cy = hdr(d, M+INN, Y+20, "Context & Opening", BLUE,
             "Set the scene and align on objectives")
    timer_badge(d, M, Y, FW, 5)

    # 2 columns
    hw = (FW - INN*2 - 20) // 2
    # Left: objectives
    rrect(d, M+INN, cy, hw, b0_h - (cy - Y) - 20, fill=INNER, outline=BORD, r=12)
    d.text((M+INN+18, cy+16), "TODAY'S OBJECTIVES", fill=BLUE, font=FL)
    objs = [
        "• Map every channel demand enters through",
        "• Define what qualifies as portfolio-level",
        "• Co-design prioritization criteria",
        "• Draft governance model using existing meetings",
        "• Agree on baseline KPIs to reuse",
    ]
    for i, o in enumerate(objs):
        d.text((M+INN+18, cy+58+i*34), o, fill=DIM, font=FSB)

    # Right: ground rules
    rx = M+INN+hw+20
    rrect(d, rx, cy, hw, b0_h - (cy - Y) - 20, fill=INNER, outline=BORD, r=12)
    d.text((rx+18, cy+16), "GROUND RULES", fill=GREEN, font=FL)
    rules = [
        "• Every voice matters — round robin first",
        "• No wrong answers — we're co-designing",
        "• Tangents → parking lot",
        "• Work on stickies, not in discussion",
        "• Time-boxed — we will move",
    ]
    for i, r_ in enumerate(rules):
        d.text((rx+18, cy+58+i*34), r_, fill=DIM, font=FSB)

    Y += b0_h + GAP

    # ═══════════════════════════════════════════════════════════════════
    # BLOCK 1 — DEMAND FLOW MAPPING (15 min)
    # ═══════════════════════════════════════════════════════════════════
    b1_h = 780
    rrect(d, M, Y, FW, b1_h, outline=BORD)
    cy = hdr(d, M+INN, Y+20, "Demand Flow Mapping", ORANGE,
             "Every channel, every source, every hand-off")
    timer_badge(d, M, Y, FW, 15)

    # Prompt
    d.text((M+INN, cy), "PROMPT",  fill=MUTED, font=FSM)
    cy += 32
    d.text((M+INN, cy),
           '"Write every channel you receive work through — one sticky each."',
           fill=WHITE, font=FP)
    cy += 46

    # Main wall + side column
    wall_w = int((FW - INN*2 - 20) * 0.68)
    side_w = FW - INN*2 - 20 - wall_w

    wall_h = Y + b1_h - cy - 20
    inbox(d, M+INN, cy, wall_w, wall_h, "CHANNEL WALL — drop stickies here, group them as they emerge")

    # Side: what to write on each sticky
    sx = M+INN+wall_w+20
    rrect(d, sx, cy, side_w, wall_h, fill=INNER, outline=BORD, r=12)
    d.text((sx+18, cy+16), "ON EACH STICKY", fill=ORANGE, font=FL)
    d.text((sx+18, cy+54), "Write these 4 answers:", fill=MUTED, font=FSM)
    items = [
        ("👤", "Who submits?"),
        ("⏱", "How often?"),
        ("📦", "Type of work?"),
        ("🔗", "Connects to which meeting?"),
    ]
    for i, (icon, label) in enumerate(items):
        iy = cy + 100 + i*70
        d.text((sx+18, iy), icon, fill=ORANGE, font=FP)
        d.text((sx+70, iy+6), label, fill=WHITE, font=FB)

    # Footer prompt
    fy = cy + 100 + 4*70 + 10
    d.text((sx+18, fy), "GUIDE PROMPTS", fill=MUTED, font=FSM)
    fy += 28
    guides = [
        '"Urgent requests — how do',
        'they bypass the process?"',
        '"Which sources are out of',
        'your control?"',
    ]
    for i, g in enumerate(guides):
        d.text((sx+18, fy+i*28), g, fill=DIM, font=FSM)

    Y += b1_h + GAP

    # ═══════════════════════════════════════════════════════════════════
    # BLOCK 2 — WHAT COUNTS AS PORTFOLIO (20 min)
    # ═══════════════════════════════════════════════════════════════════
    b2_h = 780
    rrect(d, M, Y, FW, b2_h, outline=BORD)
    cy = hdr(d, M+INN, Y+20, "What Counts as Portfolio", PINK,
             "Define what sits at portfolio level vs operational")
    timer_badge(d, M, Y, FW, 20)

    d.text((M+INN, cy), "PROMPT", fill=MUTED, font=FSM)
    cy += 32
    d.text((M+INN, cy),
           '"3 stickies each: what SHOULD be visible at portfolio level vs operational day-to-day?"',
           fill=WHITE, font=FP)
    cy += 50

    # 3 cluster columns
    cw = (FW - INN*2 - 40) // 3
    card_h = Y + b2_h - cy - 20
    clusters = [
        ("RUN",       "Day-to-day operations, support,\nmaintenance, incidents", GREEN),
        ("GROW",      "Enhancements, feature work,\nnew capabilities for existing products", ORANGE),
        ("TRANSFORM", "Strategic initiatives, new products,\nplatform changes, innovation", PINK),
    ]
    for i, (title, sub, col) in enumerate(clusters):
        bx = M + INN + i*(cw+20)
        rrect(d, bx, cy, cw, card_h, fill=INNER, outline=BORD, r=12)
        d.rectangle([bx, cy, bx+cw, cy+6], fill=col)
        d.text((bx+18, cy+22), title, fill=col, font=FL)
        for li, ln in enumerate(sub.split("\n")):
            d.text((bx+18, cy+62+li*28), ln, fill=MUTED, font=FSM)
        # sticky drop zone label
        d.text((bx+18, cy+card_h-44), "DROP STICKIES HERE", fill=DIM, font=FSM)

    Y += b2_h + GAP

    # ═══════════════════════════════════════════════════════════════════
    # BLOCK 3 — CRITERIA + TRIAGE (20 min)
    # ═══════════════════════════════════════════════════════════════════
    b3_h = 860
    rrect(d, M, Y, FW, b3_h, outline=BORD)
    cy = hdr(d, M+INN, Y+20, "Prioritization Criteria + Triage", RED,
             "Weighted criteria become triage gates")
    timer_badge(d, M, Y, FW, 20)

    d.text((M+INN, cy), "PROMPT", fill=MUTED, font=FSM)
    cy += 32
    d.text((M+INN, cy),
           '"Distribute 100 points across the criteria. The weights define the triage gates."',
           fill=WHITE, font=FP)
    cy += 50

    # Left: criteria table
    tbl_w = int((FW - INN*2 - 20) * 0.55)
    gate_w = FW - INN*2 - 20 - tbl_w

    rrect(d, M+INN, cy, tbl_w, b3_h - (cy - Y) - 20, fill=INNER, outline=BORD, r=12)
    d.text((M+INN+18, cy+16), "CRITERIA WEIGHTING", fill=RED, font=FL)

    # Table header
    tx = M+INN+18
    ty = cy + 60
    d.rectangle([tx, ty, tx+tbl_w-36, ty+40], fill=(20, 50, 80))
    d.text((tx+16, ty+8), "CRITERION", fill=ORANGE, font=FTH)
    d.text((tx+tbl_w-220, ty+8), "WEIGHT (pts)", fill=ORANGE, font=FTH)
    ty += 44

    criteria = [
        "Business Value",
        "User Experience",
        "Effort / Cost",
        "Compliance / Risk",
        "Dependencies",
        "Strategic Fit",
    ]
    for i, name in enumerate(criteria):
        bg = ROW_DARK if i % 2 == 0 else ROW_LIGHT
        d.rectangle([tx, ty, tx+tbl_w-36, ty+64], fill=bg)
        d.text((tx+16, ty+18), name, fill=WHITE, font=FTD)
        # Weight box
        wbx = tx+tbl_w-200
        rrect(d, wbx, ty+14, 160, 36, fill=BG, outline=BORD, r=6)
        d.text((wbx+12, ty+20), "____ pts", fill=MUTED, font=FTD)
        ty += 66

    ty += 10
    d.text((tx+16, ty), "TOTAL must = 100", fill=RED, font=FL)

    # Right: triage gates
    gx = M+INN+tbl_w+20
    rrect(d, gx, cy, gate_w, b3_h - (cy - Y) - 20, fill=INNER, outline=BORD, r=12)
    d.text((gx+18, cy+16), "TRIAGE GATES", fill=ORANGE, font=FL)
    d.text((gx+18, cy+54), "Criteria become scoring gates:", fill=MUTED, font=FSM)

    gates = [
        ("GATE 1 — FAST TRACK", "Score > 70", "Auto-approve, goes to backlog", GREEN),
        ("GATE 2 — REVIEW",     "Score 40-70","Review in governance forum", ORANGE),
        ("GATE 3 — REJECT",     "Score < 40", "Decline with structured feedback", RED),
    ]
    gy = cy + 100
    for title, score, desc, col in gates:
        rrect(d, gx+18, gy, gate_w-36, 140, fill=BG, outline=col, r=10)
        d.rectangle([gx+18, gy, gx+18+6, gy+140], fill=col)
        d.text((gx+36, gy+16), title, fill=col, font=FL)
        d.text((gx+36, gy+58), score, fill=WHITE, font=FB)
        d.text((gx+36, gy+96), desc, fill=MUTED, font=FSM)
        gy += 156

    Y += b3_h + GAP

    # ═══════════════════════════════════════════════════════════════════
    # BLOCK 4 — GOVERNANCE 3-TIER (15 min)
    # ═══════════════════════════════════════════════════════════════════
    b4_h = 820
    rrect(d, M, Y, FW, b4_h, outline=BORD)
    cy = hdr(d, M+INN, Y+20, "Governance — 3 Tier Model", LIME,
             "Place existing meetings on the right tier, then identify gaps")
    timer_badge(d, M, Y, FW, 15)

    d.text((M+INN, cy), "PROMPT", fill=MUTED, font=FSM)
    cy += 32
    d.text((M+INN, cy),
           '"Take the meetings from Block 1 and place them on the tier they belong. Add what\'s missing."',
           fill=WHITE, font=FP)
    cy += 50

    tier_h = (Y + b4_h - cy - 30) // 3
    tiers = [
        ("STRATEGIC",   "Quarterly / Monthly",  "Portfolio direction, investment decisions, OKRs",  BLUE),
        ("OPERATIONAL", "Bi-weekly / Weekly",   "Prioritization, triage, cross-team coordination",  ORANGE),
        ("TACTICAL",    "Daily / Weekly",       "Execution, blockers, team-level decisions",         GREEN),
    ]
    for i, (lab, cad, scope, col) in enumerate(tiers):
        ty = cy + i*(tier_h+6)
        rrect(d, M+INN, ty, FW-INN*2, tier_h-6, fill=INNER, outline=col, r=12)
        d.rectangle([M+INN, ty, M+INN+8, ty+tier_h-6], fill=col)
        d.text((M+INN+28, ty+14), lab, fill=col, font=FL)
        d.text((M+INN+28, ty+52), cad, fill=WHITE, font=FSB)
        d.text((M+INN+28, ty+84), scope, fill=MUTED, font=FSM)
        # drop zone on the right
        dzx = M+INN+500
        dzw = FW-INN*2 - 520
        rrect(d, dzx, ty+14, dzw, tier_h-34, fill=BG, outline=BORD, r=8)
        d.text((dzx+14, ty+22), "EXISTING MEETINGS + GAPS — drop stickies", fill=MUTED, font=FSM)

    Y += b4_h + GAP

    # ═══════════════════════════════════════════════════════════════════
    # BLOCK 5 — KPIs + CLOSE (15 min)
    # ═══════════════════════════════════════════════════════════════════
    b5_h = 840
    rrect(d, M, Y, FW, b5_h, outline=BORD)
    cy = hdr(d, M+INN, Y+20, "KPIs + Close", GREEN,
             "Baseline metrics and top decisions for next session")
    timer_badge(d, M, Y, FW, 15)

    # Left: KPI grid
    kpi_w = int((FW - INN*2 - 20) * 0.60)
    dec_w = FW - INN*2 - 20 - kpi_w

    rrect(d, M+INN, cy, kpi_w, b5_h - (cy - Y) - 20, fill=INNER, outline=BORD, r=12)
    d.text((M+INN+18, cy+16), "KPI SELECTION — Keep / Adapt / Drop", fill=GREEN, font=FL)
    d.text((M+INN+18, cy+54), "For each KPI, mark one column:", fill=MUTED, font=FSM)

    # Table header
    tx = M+INN+18
    ty = cy + 100
    tbl_in_w = kpi_w - 36

    cols_pct = [0.50, 0.17, 0.17, 0.16]
    col_labels = ["KPI", "KEEP", "ADAPT", "DROP"]
    col_xs = []
    cxx = tx
    for p in cols_pct:
        col_xs.append(int(cxx))
        cxx += tbl_in_w * p

    d.rectangle([tx, ty, tx+tbl_in_w, ty+42], fill=(20, 50, 80))
    for lab, cxx in zip(col_labels, col_xs):
        col = GREEN if lab == "KEEP" else (ORANGE if lab == "ADAPT" else (RED if lab == "DROP" else ORANGE))
        d.text((cxx+14, ty+8), lab, fill=col, font=FTH)
    ty += 46

    kpis = [
        "Lead Time (idea → delivery)",
        "Cycle Time (in progress → done)",
        "Throughput (items / week)",
        "% On-time delivery",
        "Backlog Aging",
        "Value Realization %",
        "Predictability Index",
    ]
    for i, k in enumerate(kpis):
        bg = ROW_DARK if i % 2 == 0 else ROW_LIGHT
        d.rectangle([tx, ty, tx+tbl_in_w, ty+54], fill=bg)
        d.text((col_xs[0]+14, ty+14), k, fill=WHITE, font=FTD)
        for ci in range(3):
            cbx = col_xs[ci+1] + int(tbl_in_w*cols_pct[ci+1]/2) - 16
            rrect(d, cbx, ty+12, 32, 32, fill=BG, outline=BORD, r=6)
        ty += 56

    # Right: top 3 decisions
    dx = M+INN+kpi_w+20
    rrect(d, dx, cy, dec_w, b5_h - (cy - Y) - 20, fill=INNER, outline=BORD, r=12)
    d.text((dx+18, cy+16), "TOP 3 DECISIONS", fill=RED, font=FL)
    d.text((dx+18, cy+54), "What must be built before next session?", fill=MUTED, font=FSM)

    dy = cy + 100
    for i in range(3):
        rrect(d, dx+18, dy, dec_w-36, 150, fill=BG, outline=BORD, r=10)
        d.rectangle([dx+18, dy, dx+24, dy+150], fill=[GREEN, ORANGE, RED][i])
        d.text((dx+36, dy+14), f"DECISION #{i+1}", fill=[GREEN, ORANGE, RED][i], font=FL)
        d.text((dx+36, dy+56), "_____________________________", fill=MUTED, font=FSM)
        d.text((dx+36, dy+86), "Owner:  __________", fill=MUTED, font=FSM)
        d.text((dx+36, dy+116), "Due date:  __________", fill=MUTED, font=FSM)
        dy += 166

    Y += b5_h + GAP

    # ═══════════════════════════════════════════════════════════════════
    # BLOCK 6 — PARKING LOT
    # ═══════════════════════════════════════════════════════════════════
    b6_h = 320
    rrect(d, M, Y, FW, b6_h, outline=BORD)
    cy = hdr(d, M+INN, Y+20, "Parking Lot", ORANGE,
             "Capture tangents and follow-ups for async resolution")

    inbox(d, M+INN, cy, FW-INN*2, b6_h - (cy - Y) - 20,
          "Questions, ideas, topics to revisit — do not let them derail the flow")

    Y += b6_h + GAP

    # ═══════════════════════════════════════════════════════════════════
    # TIMING SUMMARY
    # ═══════════════════════════════════════════════════════════════════
    rrect(d, M, Y, FW, 90, fill=TIMER_BG, outline=BLUE, r=14)
    d.text((M+24, Y+14), "SESSION TIMING — 90 minutes total", fill=WHITE, font=FL)
    d.text((M+24, Y+54),
           "Context (5)  →  Demand Flow (15)  →  What Counts (20)  →  Criteria + Triage (20)  →  Governance (15)  →  KPIs + Close (15)",
           fill=DIM, font=FSM)
    Y += 90 + M

    # ── Crop + save ────────────────────────────────────────────────
    img_out = img.crop((0, 0, W, Y))
    return img_out


# ═══════════════════════════════════════════════════════════════════════
base = "/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Miro Discovery Sessions"
os.makedirs(base, exist_ok=True)

result = generate_canvas()
out = f"{base}/SESSION_2_Demand_Portfolio_Canvas.png"
result.save(out, "PNG")
px = result.size[0] * result.size[1]
print(f"{result.size[0]}x{result.size[1]} — {px/1e6:.1f}MP → {os.path.basename(out)}")
