"""Question Reference Canvases — Stakeholder Interview + Demand × Portfolio
Same dark-navy style as Stakeholder Interview Canvas"""
from PIL import Image, ImageDraw, ImageFont
import os

# ── Colors (match interview canvas) ─────────────────────────────────────
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
FL  = font(30, True)
FB  = font(26)
FSM = font(22)
FNUM = font(48, True)
FCAT = font(20, True)
FQ   = font(26)

W = 5600
M = 50
GAP = 36
INN = 24
FW = W - M*2

def rrect(d, x, y, w, h, fill=CARD, outline=None, r=18):
    if w <= 0 or h <= 0: return
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

def wrap(d, text, max_w, f):
    words = text.split()
    lines, line = [], ""
    for w_ in words:
        t = f"{line} {w_}".strip()
        if d.textlength(t, font=f) <= max_w:
            line = t
        else:
            if line: lines.append(line)
            line = w_
    if line: lines.append(line)
    return lines


def question_card(d, x, y, w, num, category, qtext, color):
    """Compute height + draw card. Returns the height used."""
    pad = 22
    line_h = 36
    # Wrap question text
    text_w = w - pad*2 - 80  # leave room for number badge
    lines = wrap(d, qtext, text_w, FQ)
    text_h = len(lines) * line_h
    card_h = max(160, text_h + 100)

    # Card box
    rrect(d, x, y, w, card_h, fill=INNER, outline=BORD, r=14)
    # Top accent
    d.rectangle([x, y, x+w, y+6], fill=color)

    # Number badge
    rrect(d, x+pad, y+pad, 64, 64, fill=color, r=12)
    nstr = f"{num:02d}"
    nw = d.textlength(nstr, font=FNUM)
    d.text((x+pad+(64-nw)//2, y+pad+4), nstr, fill=WHITE, font=FNUM)

    # Category tag
    cat_text = category.upper()
    cw_ = d.textlength(cat_text, font=FCAT)
    rrect(d, x+pad+76, y+pad+4, int(cw_+24), 28, fill=color, r=8)
    d.text((x+pad+88, y+pad+8), cat_text, fill=WHITE, font=FCAT)

    # Question text
    ty = y + pad + 50
    for ln in lines:
        d.text((x+pad+76, ty), ln, fill=WHITE, font=FQ)
        ty += line_h

    return card_h


def section_banner(d, x, y, w, title, subtitle, color, h=120):
    rrect(d, x, y, w, h, fill=INNER, outline=color, r=18)
    # Side stripe
    d.rectangle([x, y, x+10, y+h], fill=color)
    d.text((x+34, y+22), title, fill=WHITE, font=FS)
    d.text((x+34, y+76), subtitle, fill=DIM, font=FSB)
    return h


def render_two_col(d, questions, start_y, col_w, gap_x, gap_y, start_num=1):
    """Render question cards in 2-column layout. Returns final y."""
    cy_l = start_y
    cy_r = start_y
    for i, (cat, q, color) in enumerate(questions, start=start_num):
        # Pre-compute card height by simulating
        text_w = col_w - 22*2 - 80
        lines = wrap(d, q, text_w, FQ)
        card_h = max(160, len(lines)*36 + 100)
        # Place in shorter column
        if cy_l <= cy_r:
            x = M
            y = cy_l
            cy_l += card_h + gap_y
        else:
            x = M + col_w + gap_x
            y = cy_r
            cy_r += card_h + gap_y
        question_card(d, x, y, col_w, i, cat, q, color)
    return max(cy_l, cy_r)


# ═══════════════════════════════════════════════════════════════════════
# CANVAS 1 — STAKEHOLDER INTERVIEW QUESTIONS
# ═══════════════════════════════════════════════════════════════════════
def build_stakeholder_questions():
    img = Image.new("RGB", (W, 7000), BG)
    d = ImageDraw.Draw(img)

    Y = M
    d.text((M, Y), "Stakeholder Interview — Questions Reference", fill=WHITE, font=FT)
    Y += 70
    d.text((M, Y), "17 questions to guide the 1-on-1 / group discovery  |  45 min sessions",
           fill=DIM, font=FM)
    Y += 70

    questions = [
        ("Role & Context",       "What is your role and how do you interact with Digital Products? How is your role changing with the new CIO direction and transformation focus, and how can this project support you?", BLUE),
        ("Strategic Priorities", "What are your top priorities for the next 6 months?", BLUE),
        ("Delivery Process",     "Describe how a typical request goes from idea to delivery today.", CYAN),
        ("Pain Points",          "Where do you experience the most friction or delays?", RED),
        ("Prioritization",       "How do you currently prioritize work? Please answer at two levels: (a) how the department prioritizes across areas, and (b) how you personally prioritize within your area.", ORANGE),
        ("Tooling",              "What tools do you use for tracking work? (Jira, ServiceNow, Excel, email, other?) Feel free to share screenshots or examples.", CYAN),
        ("Visibility",           "How do you get visibility into what's being worked on? Which meetings or governance forums give you this visibility today?", LIME),
        ("Pain Points",          "If you could fix one thing about how Digital Products operates today, what would it be?", RED),
        ("Outcome Measurement",  "What would make you say this project was a success?", GREEN),
        ("Governance",           "How are decisions escalated when there's a conflict between teams or priorities?", PINK),
        ("Visibility",           "What information do you need but currently don't have access to?", LIME),
        ("Performance",          "How do you measure the value delivered by Digital Products to your area?", GREEN),
        ("Alignment",            "How is alignment ensured between Digital Products, BRS, and other consuming areas? Which existing meetings or forums support this alignment?", PURPLE),
        ("Decision Making",      "Where do you see silos that need to be connected as part of the future operating model? Is decision-making centralized or distributed?", PINK),
        ("Pilot & Adoption",     "How do you think a pilot could work for your area, and what considerations or tips would you share to scale it across the rest of the department?", LIME),
        ("Risks",                "What do you think is the biggest risk if we don't change how Digital Products operates?", RED),
        ("Open Input",           "Is there anything you think we should know that we haven't asked?", DIM),
    ]

    col_w = (FW - GAP) // 2
    Y = render_two_col(d, questions, Y, col_w, GAP, GAP, start_num=1)
    Y += M

    return img.crop((0, 0, W, Y))


# ═══════════════════════════════════════════════════════════════════════
# CANVAS 2 — DEMAND × PORTFOLIO QUESTIONS
# ═══════════════════════════════════════════════════════════════════════
def build_demand_portfolio_questions():
    img = Image.new("RGB", (W, 12000), BG)
    d = ImageDraw.Draw(img)

    Y = M
    d.text((M, Y), "Demand × Portfolio — Questions Reference", fill=WHITE, font=FT)
    Y += 70
    d.text((M, Y), "44 questions for the joint discovery session  |  Pair with the Flow Canvas",
           fill=DIM, font=FM)
    Y += 70

    # ── PART A: DEMAND ──────────────────────────────────────────────────
    section_banner(d, M, Y, FW,
                   "Part A — Demand Sources Mapping",
                   "How work requests enter Digital Products and how they connect to existing governance",
                   CYAN)
    Y += 140

    demand_qs = [
        ("Channels",       "List every channel through which work requests currently enter Digital Products (email, Teams, meetings, Jira, ServiceNow, SharePoint, phone, direct message, etc.).", CYAN),
        ("Channels",       "For each channel above: who submits the request, how often, what type of work is typically requested, and which existing governance meeting (if any) is connected to it?", CYAN),
        ("Intake",         "Is there a formal intake form or template? If so, what fields does it capture?", BLUE),
        ("Intake",         "How are requests classified today? (Type, priority, size, urgency)", BLUE),
        ("Intake",         "What happens after a request is submitted? Who triages it, and how long does triage typically take?", BLUE),
        ("Intake",         "How often do 'urgent' requests bypass the normal process? What drives these exceptions?", BLUE),
        ("Quality",        "What percentage of requests have clear business value or acceptance criteria at the time of submission?", PINK),
        ("Sources",        "Are there demand sources that Digital Products does not directly control (regulatory, external mandates, compliance)? If so, how many and of what type?", ORANGE),
        ("Tooling",        "What tools are used today to log and track demand? (Jira, ServiceNow, Excel, SharePoint, other)", CYAN),
        ("Trust",          "What would need to be true for requestors to trust and consistently use a single intake channel?", BLUE),
        ("Performance",    "What is the average time from request submission to someone acknowledging it?", GREEN),
        ("Patterns",       "Are there seasonal or cyclical patterns in demand? (Quarter-end, regulatory deadlines, budget cycles, audits, annual planning). Please map the key dates when possible.", LIME),
        ("Quality",        "How many requests are duplicates or overlap with existing work? What is the root cause and when does it most often happen?", PINK),
        ("Governance",     "What happens to requests that are rejected or deprioritized? Are requestors notified? How?", PURPLE),
        ("Channels",       "Is there demand that comes directly to individual developers, bypassing any formal channel?", CYAN),
        ("Dependencies",   "Are there handoffs or dependencies in the intake process that slow things down or cause requests to get lost?", RED),
        ("Alignment",      "How involved are the requesting areas (BRS, OPCOs) in defining and validating what they need before a request reaches Digital Products?", PURPLE),
        ("Automation",     "Are there manual tasks in the intake process that could be automated?", LIME),
    ]

    col_w = (FW - GAP) // 2
    Y = render_two_col(d, demand_qs, Y, col_w, GAP, GAP, start_num=1)
    Y += GAP

    # ── PART B: PORTFOLIO ───────────────────────────────────────────────
    section_banner(d, M, Y, FW,
                   "Part B — Portfolio Review",
                   "How work is defined, prioritized, governed, tracked and measured at portfolio level",
                   PURPLE)
    Y += 140

    portfolio_qs = [
        ("Definition",     "What types of initiatives are currently considered part of the portfolio (run, grow, transform)? If the concept does not exist formally, what would make sense in your view?", PURPLE),
        ("Definition",     "Is there a clear definition of what qualifies as a portfolio initiative? If not, what criteria should define it going forward?", PURPLE),
        ("Inventory",      "Do you have a centralized and up-to-date view of all active initiatives? Where does it live, and how current is it?", BLUE),
        ("Inventory",      "How often is the portfolio inventory reviewed and updated?", BLUE),
        ("Prioritization", "What criteria are used today to prioritize initiatives? Which criteria do you think should be used going forward?", ORANGE),
        ("Prioritization", "How do you balance business value, cost, and risk in prioritization decisions?", ORANGE),
        ("Prioritization", "Who is responsible for prioritization decisions?", ORANGE),
        ("Governance",     "Is there a formal governance structure for portfolio management today? If not, what would the equivalent look like in your view?", PINK),
        ("Governance",     "How are decisions escalated or resolved when there are conflicts between areas or priorities?", PINK),
        ("Governance",     "What committees, forums, or decision-making bodies are currently in place? Which are missing that you believe should exist?", PINK),
        ("Governance",     "How effective is the current governance in balancing speed and control?", PINK),
        ("Alignment",      "How do you ensure initiatives align with strategic objectives?", PURPLE),
        ("Alignment",      "Are strategic priorities clearly translated into the portfolio? How?", PURPLE),
        ("Funding",        "How are initiatives funded and budgeted?", LIME),
        ("Funding",        "Is funding fixed or flexible based on performance and priorities?", LIME),
        ("Dependencies",   "Are there handoffs or dependencies across teams that slow delivery?", RED),
        ("Dependencies",   "How do these dependencies impact prioritization and delivery timelines?", RED),
        ("Tracking",       "What metrics or KPIs do you track regularly at the portfolio level?", GREEN),
        ("Tracking",       "What tools or dashboards are used for portfolio tracking?", GREEN),
        ("Tracking",       "How do you track progress across initiatives?", GREEN),
        ("Performance",    "How do you measure the success of initiatives?", GREEN),
        ("Performance",    "How are these metrics used in decision-making?", GREEN),
        ("Performance",    "Are benefits tracked after delivery (value realization)? If yes, how?", GREEN),
        ("Quality",        "How do you measure quality in delivery?", CYAN),
        ("Risk",           "How are risks and issues managed at the portfolio level?", RED),
        ("Risk",           "Is there a standardized way to report and mitigate risks?", RED),
    ]

    Y = render_two_col(d, portfolio_qs, Y, col_w, GAP, GAP, start_num=1)
    Y += M

    return img.crop((0, 0, W, Y))


# ═══════════════════════════════════════════════════════════════════════
base = "/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Miro Discovery Sessions"
os.makedirs(base, exist_ok=True)

for name, builder in [
    ("Stakeholder_Interview_Questions", build_stakeholder_questions),
    ("Demand_Portfolio_Questions",      build_demand_portfolio_questions),
]:
    result = builder()
    out = f"{base}/SESSION_QUESTIONS_{name}.png"
    result.save(out, "PNG")
    px = result.size[0] * result.size[1]
    print(f"[{name}] {result.size[0]}x{result.size[1]} — {px/1e6:.1f}MP → {os.path.basename(out)}")
