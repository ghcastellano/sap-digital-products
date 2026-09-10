"""Generate Stakeholder Interview Canvas — two versions: 1-on-1 and Group"""
from PIL import Image, ImageDraw, ImageFont
import os

# ── Colors ──────────────────────────────────────────────────────────────
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
              "/System/Library/Fonts/SFNSText.ttf",
              "/Library/Fonts/Arial.ttf"]:
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
FTIMER_SM = font(22)

W = 5600
M   = 50
GAP = 40
INN = 24
FW  = W - M*2


def generate_canvas(mode="single"):
    """mode: 'single' or 'group'"""
    is_group = mode == "group"

    img = Image.new("RGB", (W, 7000), BG)
    d = ImageDraw.Draw(img)

    def rrect(x, y, w, h, fill=CARD, outline=None, r=18):
        d.rounded_rectangle([x, y, x+w, y+h], r, fill=fill,
                            outline=outline, width=2 if outline else 0)

    def hdr(x, y, text, color, sub=None):
        d.ellipse([x, y+6, x+26, y+32], fill=color)
        d.text((x+36, y), text, fill=WHITE, font=FS)
        y += 52
        if sub:
            d.text((x+36, y), sub, fill=DIM, font=FSB)
            y += 36
        return y + 10

    def inbox(x, y, w, h, label=None):
        rrect(x, y, w, h, fill=INNER, outline=BORD, r=12)
        if label:
            d.text((x+16, y+12), label, fill=DIM, font=FSM)

    def timer_badge(block_x, block_y, block_w, write_min, share_min):
        """Draw a timer badge inside the block, top-right corner"""
        text = f"Write: {write_min} min  |  Share: {share_min} min"
        tw = d.textlength(text, font=FTIMER)
        bw = int(tw + 30)
        bh = 36
        bx = block_x + block_w - bw - INN
        by = block_y + INN
        rrect(bx, by, bw, bh, fill=TIMER_BG, outline=BLUE, r=8)
        d.text((bx+15, by+6), text, fill=WHITE, font=FTIMER)

    # ═══════════════════════════════════════════════════════════════════
    # TITLE
    # ═══════════════════════════════════════════════════════════════════
    Y = M
    d.text((M, Y), "Stakeholder Interview Canvas", fill=WHITE, font=FT)
    Y += 70
    if is_group:
        meta = "Group Discovery  |  45 min  |  Digital Products Team"
    else:
        meta = "1-on-1 Discovery  |  45 min  |  Director, IT Product Management"
    d.text((M, Y), meta, fill=DIM, font=FM)
    Y += 60

    # ═══════════════════════════════════════════════════════════════════
    # BLOCK 1 — YOUR WORLD  (bigger priority area)
    # ═══════════════════════════════════════════════════════════════════
    b1_h = 650
    rrect(M, Y, FW, b1_h, outline=BORD)
    cy = hdr(M+INN, Y+20, "Your World", BLUE, "Understanding scope, evolution, and priorities")

    if is_group:
        timer_badge(M, Y, FW, 5, 5)

    left_w = FW//2 - 40
    draw_h = Y + b1_h - cy - 20

    # Left: org drawing
    inbox(M+INN, cy, left_w, draw_h,
          "DRAW YOUR ORG LANDSCAPE — who do you report to, what teams report into you?")

    # Right: 2 boxes — Role Today (smaller) + Priorities (bigger)
    rx = M + FW//2 + 20
    rw = FW//2 - 44

    # Role Today — 35% of space
    role_h = int(draw_h * 0.30)
    inbox(rx, cy, rw, role_h, "YOUR ROLE TODAY — and how is it changing?")

    # Priorities — 65% of space
    prio_y = cy + role_h + 14
    prio_h = draw_h - role_h - 14
    inbox(rx, prio_y, rw, prio_h, "TOP PRIORITIES")

    Y += b1_h + GAP

    # ═══════════════════════════════════════════════════════════════════
    # BLOCK 2 — PAIN POINTS & GAPS  (replaced "The One Thing")
    # ═══════════════════════════════════════════════════════════════════
    b2_h = 700
    rrect(M, Y, FW, b2_h, outline=BORD)
    cy = hdr(M+INN, Y+20, "Pain Points & Gaps", RED,
             "What hurts most today — internally, for end clients, and where the gaps are")

    if is_group:
        timer_badge(M, Y, FW, 5, 7)

    # 3 columns
    cw = (FW - INN*2 - 28) // 3
    col_data = [
        ("3 MAIN PAIN POINTS", "What are the biggest operational\npain points you face today?", RED),
        ("END CLIENT PAIN POINTS", "What pain are your end clients/users\nexperiencing because of how we work?", ORANGE),
        ("GAPS", "Where are the biggest gaps between\nwhere you are and where you need to be?", BLUE),
    ]
    for i, (title, sub, col) in enumerate(col_data):
        bx = M + INN + i*(cw+14)
        card_h = Y + b2_h - cy - 20
        rrect(bx, cy, cw, card_h, fill=INNER, outline=BORD, r=12)
        d.rectangle([bx, cy, bx+cw, cy+6], fill=col)
        d.text((bx+18, cy+20), title, fill=col, font=FL)
        for li, ln in enumerate(sub.split("\n")):
            d.text((bx+18, cy+60+li*28), ln, fill=MUTED, font=FSM)

    Y += b2_h + GAP

    # ═══════════════════════════════════════════════════════════════════
    # BLOCK 3 — HOW DECISIONS GET MADE  (full width, no prioritization)
    # ═══════════════════════════════════════════════════════════════════
    # Table rows(5*80=400) + header(46) + hdr(100) + missing(250) + silos(280) + pilot(300) + pad(50)
    b3_h = 100 + 46 + 400 + 50 + 250 + 280 + 300 + 50
    rrect(M, Y, FW, b3_h, outline=BORD)
    cy = hdr(M+INN, Y+20, "How Decisions Get Made", PINK,
             "Governance forums from your RFI — let's validate")

    if is_group:
        timer_badge(M, Y, FW, 5, 5)

    TBL_X = M + INN
    TBL_W = FW - INN*2

    # Column positions — wider since full width now
    pcts = [0.20, 0.08, 0.28, 0.30, 0.14]
    col_xs = []
    cx = TBL_X
    for p in pcts:
        col_xs.append(int(cx))
        cx += TBL_W * p

    # Table header
    d.rectangle([TBL_X, cy, TBL_X+TBL_W, cy+42], fill=(20, 50, 80))
    for txt, cxx in zip(["FORUM","CADENCE","PARTICIPANTS","DECISION SCOPE","STATUS"], col_xs):
        d.text((cxx+10, cy+8), txt, fill=ORANGE, font=FTH)
    cy += 46

    forums = [
        ("Critical Integrations",     "Weekly",  "Business Apps, P&S, Digital Products",       "End-to-end cross-team integration issues",       "Active"),
        ("Daily BSS BI/BW Stand Up",  "Weekly",  "Data & Analytics Teams, Digital Products",    "Alignment on priorities, workloads, blockers",     "Active"),
        ("Data, Analytics &\nSolutions Check-In","Weekly","Data & Analytics Teams, Digital Products","Operations, transition, target End State Model","Active"),
        ("CX & Digital CMP Update",   "Monthly", "CX Digital, Operations, CMP President, VP CS","CMP workload review, 1-3 month look ahead",      "Active"),
        ("Manager Check-in",          "Weekly",  "[To be confirmed]",                           "[To be confirmed]",                                "[TBC]"),
    ]

    RH = 80
    for fi, row in enumerate(forums):
        bg = ROW_DARK if fi % 2 == 0 else ROW_LIGHT
        d.rectangle([TBL_X, cy, TBL_X+TBL_W, cy+RH], fill=bg)
        for ci, (val, cxx) in enumerate(zip(row, col_xs)):
            color = WHITE if ci == 0 else DIM
            if "[To be confirmed]" in val or "[TBC]" in val:
                color = ORANGE
            for li, ln in enumerate(val.split("\n")):
                d.text((cxx+10, cy+14+li*30), ln, fill=color, font=FTD)
        cy += RH

    cy += 16
    d.text((TBL_X+6, cy), "Validate each row. Mark corrections directly.", fill=MUTED, font=FSM)

    # ── MISSING FORUMS ──
    cy += 40
    d.text((TBL_X+6, cy), "MISSING forums not listed above:", fill=RED, font=FL)
    cy += 42
    inbox(TBL_X, cy, TBL_W, 200)

    # ── SILOS & CONNECTIONS NEEDED ──
    cy += 220
    d.text((TBL_X+6, cy), "SILOS & CONNECTIONS NEEDED", fill=ORANGE, font=FL)
    cy += 38
    d.text((TBL_X+6, cy), "Where are the disconnects between teams?", fill=MUTED, font=FSM)
    cy += 34
    inbox(TBL_X, cy, TBL_W, 200)

    # ── PILOT CANDIDATES ──
    cy += 220
    d.text((TBL_X+6, cy), "PILOT CANDIDATES", fill=GREEN, font=FL)
    cy += 38
    d.text((TBL_X+6, cy),
           "Which product/team would be best for a first pilot of the new operating model?",
           fill=MUTED, font=FSM)
    cy += 30
    d.text((TBL_X+6, cy),
           "And how should the pilot work in your opinion?",
           fill=MUTED, font=FSM)
    cy += 34
    inbox(TBL_X, cy, TBL_W, 200)

    Y += b3_h + GAP

    # ═══════════════════════════════════════════════════════════════════
    # BLOCK 4 — LOOKING FORWARD
    # ═══════════════════════════════════════════════════════════════════
    b4_h = 480
    rrect(M, Y, FW, b4_h, outline=BORD)
    cy = hdr(M+INN, Y+20, "Looking Forward", GREEN, "Vision and aspirations for the next 12 months")

    if is_group:
        timer_badge(M, Y, FW, 3, 5)

    cw = (FW - INN*2 - 28) // 3
    cards = [
        ("BIGGEST OPPORTUNITY", "If this project succeeds,\nwhat changes most?", GREEN),
        ("BIGGEST RISK",        "What could derail\nthe transformation?", ORANGE),
        ("ONE DECISION PENDING","What decision needs to happen\nin the next 30 days?", RED),
    ]
    for i, (title, sub, col) in enumerate(cards):
        bx = M + INN + i*(cw+14)
        card_h = Y + b4_h - cy - 20
        rrect(bx, cy, cw, card_h, fill=INNER, outline=BORD, r=12)
        d.rectangle([bx, cy, bx+cw, cy+6], fill=col)
        d.text((bx+18, cy+20), title, fill=col, font=FL)
        for li, ln in enumerate(sub.split("\n")):
            d.text((bx+18, cy+60+li*28), ln, fill=MUTED, font=FSM)

    Y += b4_h + GAP

    # ═══════════════════════════════════════════════════════════════════
    # BLOCK 5 — WRAP-UP
    # ═══════════════════════════════════════════════════════════════════
    b5_h = 300
    rrect(M, Y, FW, b5_h, outline=BORD)
    cy = hdr(M+INN, Y+20, "Wrap-Up", BLUE, "Closing reflections")

    if is_group:
        timer_badge(M, Y, FW, 2, 3)

    hw = (FW - INN*2 - 20) // 2
    for i, (title, sub) in enumerate([
        ("VALUE MEASUREMENT", "How would you know this project succeeded? What metric moves?"),
        ("ANYTHING ELSE?",    "Is there something we should have asked but didn't?"),
    ]):
        bx = M + INN + i*(hw+20)
        d.text((bx+4, cy), title, fill=WHITE, font=FL)
        d.text((bx+4, cy+38), sub, fill=MUTED, font=FSM)
        inbox(bx, cy+70, hw, Y + b5_h - (cy+70) - 20)

    Y += b5_h + M

    # ── Group: add timing summary at bottom ──
    if is_group:
        rrect(M, Y, FW, 80, fill=TIMER_BG, outline=BLUE, r=14)
        d.text((M+24, Y+12), "SESSION TIMING — 45 minutes total", fill=WHITE, font=FL)
        d.text((M+24, Y+46),
               "Your World (10 min)   |   Pain Points & Gaps (12 min)   |   Decisions & Governance (10 min)   |   Looking Forward (8 min)   |   Wrap-Up (5 min)",
               fill=DIM, font=FSM)
        Y += 80 + M

    # ── Crop + save ─────────────────────────────────────────────────
    img_out = img.crop((0, 0, W, Y))
    return img_out


# ═══════════════════════════════════════════════════════════════════════
# Generate both versions
# ═══════════════════════════════════════════════════════════════════════
base = "/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Miro Discovery Sessions"

for mode, fname in [("single", "SESSION_1_Interview_Canvas.png"),
                     ("group",  "SESSION_1_Interview_Canvas_GROUP.png")]:
    result = generate_canvas(mode)
    path = f"{base}/{fname}"
    result.save(path, "PNG")
    px = result.size[0] * result.size[1]
    print(f"[{mode.upper()}] {result.size[0]}x{result.size[1]} — {px/1e6:.1f}MP → {fname}")
