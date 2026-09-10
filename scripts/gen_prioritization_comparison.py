"""Generate Prioritization Comparison across projects — 3 levels: Team, Product, Portfolio"""
from PIL import Image, ImageDraw, ImageFont
import os

W, H_MAX = 5600, 5000
BG     = (2, 17, 30)
CARD   = (4, 22, 44)
WHITE  = (255, 255, 255)
DIM    = (160, 175, 195)
MUTED  = (100, 120, 145)
GREEN  = (120, 190, 32)
BLUE   = (0, 82, 155)
ORANGE = (247, 148, 29)
RED    = (220, 60, 60)
PINK   = (200, 80, 180)
INNER  = (8, 28, 52)
BORD   = (40, 60, 85)
ROW_A  = (10, 35, 60)
ROW_B  = (6, 28, 50)

def font(size, bold=False):
    for p in ["/System/Library/Fonts/Helvetica.ttc",
              "/System/Library/Fonts/SFNSText.ttf",
              "/Library/Fonts/Arial.ttf"]:
        if os.path.exists(p):
            try: return ImageFont.truetype(p, size, index=1 if bold and p.endswith(".ttc") else 0)
            except: return ImageFont.truetype(p, size)
    return ImageFont.load_default()

FT  = font(56, True)
FM  = font(30)
FS  = font(38, True)
FL  = font(30, True)
FB  = font(26)
FSM = font(22)
FTH = font(26, True)
FTD = font(24)
FTAG = font(20, True)

img = Image.new("RGB", (W, H_MAX), BG)
d = ImageDraw.Draw(img)

M   = 50
GAP = 30
INN = 24
FW  = W - M*2

def rrect(x, y, w, h, fill=CARD, outline=None, r=16):
    d.rounded_rectangle([x, y, x+w, y+h], r, fill=fill, outline=outline, width=2 if outline else 0)

def tag(x, y, text, color):
    tw = d.textlength(text, font=FTAG)
    rrect(x, y, int(tw)+20, 28, fill=(*color, 40), outline=color, r=6)
    d.text((x+10, y+4), text, fill=color, font=FTAG)
    return int(tw) + 28

def source_tag(x, y, text):
    d.text((x, y), text, fill=MUTED, font=FSM)

# ═══════════════════════════════════════════════════════════════════════
# TITLE
# ═══════════════════════════════════════════════════════════════════════
Y = M
d.text((M, Y), "Prioritization Across Projects", fill=WHITE, font=FT)
Y += 66
d.text((M, Y), "What each Avangrid project has defined at Team, Product, and Portfolio levels  |  Data from KB visual chunks", fill=DIM, font=FM)
Y += 50

# ═══════════════════════════════════════════════════════════════════════
# LEVEL HEADERS
# ═══════════════════════════════════════════════════════════════════════
levels = [
    ("TEAM LEVEL", "How teams prioritize their backlog day-to-day", GREEN),
    ("PRODUCT LEVEL", "How product owners decide what to build next", BLUE),
    ("PORTFOLIO LEVEL", "How the org decides which products/initiatives get funded", ORANGE),
]

# Project data per level
projects = {
    "P&S / P&T": {
        "color": GREEN,
        "team": [
            "PBI Formula: PBI(x) = (BV + UX) / E",
            "3-step process: PO initial ordering -> Team evaluation -> Priority assignment",
            "Criteria: Business Value (BV), User Experience (UX), Effort (E)",
            "Highest H(x) value = Priority 1 (built first)",
        ],
        "team_src": "nWOW P&S Lean Operating Model, Slide 77",
        "product": [
            "Product Discovery Playbook with 3-level practices library",
            "Continuous Discovery / Continuous Delivery cycle",
            "Portfolio Committee as decision gateway between discovery and delivery",
        ],
        "product_src": "Product Discovery Playbook, Slides 9, 102",
        "portfolio": [
            "Lean Portfolio Funnel: 5 stages across 2 zones",
            "Zone 1 (Top): Discovery/Ideation/Intake -> Initial Review/Triage",
            "Zone 2 (Bottom): Approval/Funding -> Implementation Planning -> Execution",
            "3 demand sources: Strategic Programs, User Needs, OpCo Improvements",
            "Enterprise Backlog with push-to-pull system transition",
        ],
        "portfolio_src": "nWOW P&S Lean Operating Model, Slides 37, 42, 85",
    },
    "Digital CoE": {
        "color": BLUE,
        "team": [
            "Same PBI Formula: PBI(x) = (BV + UX) / E",
            "COORG methodology: Classify, Order, Organize backlog items",
            "Urgency x Value 4-quadrant matrix for classification",
            "Backlog refinement with hierarchical practice mapping",
        ],
        "team_src": "Digital CoE Playbook, Slides 78, 82-84",
        "product": [
            "Product Portfolio Management Life-Cycle Framework",
            "XP Product Management life-cycle: Discovery -> Validate -> Measure & Learn",
            "Quarterly Planning with Estimation Logic framework",
        ],
        "product_src": "Digital CoE Playbook, Slides 34, 38, 83",
        "portfolio": [
            "Portfolio Kanban with Stage-Gate Governance",
            "Portfolio Sync Meeting (Portfolio Learn Cycle cadence)",
            "Portfolio Management Team with defined roles",
            "Classes of Service: Standard, Expedite, Fixed Date, Intangible",
            "Demand Management with service class distribution analysis",
        ],
        "portfolio_src": "Digital CoE Playbook, Slides 25, 33, 72, 105",
    },
    "Renewables DTO": {
        "color": ORANGE,
        "team": [
            "Kanban-based: WIP limits per stage",
            "Classes of Service for work items: Standard, Expedite, Fixed Date, Intangible",
            "Pull-based system (reduce reactive assignment)",
            "Flow metrics: Cycle Time, Lead Time, Throughput, Aging",
        ],
        "team_src": "NextGen DTO Deliverable, Slides 5, 27, 34",
        "product": [
            "Capacity allocation by work type: 60% Build / 30% BAU / 10% Upgrades",
            "Business Value + Urgency + Class of Service for ranking",
            "Data-driven prioritization using flow metrics and demand patterns",
        ],
        "product_src": "NextGen DTO Deliverable, Slides 5, 21-22",
        "portfolio": [
            "4 Pillars: Demand Intake -> Prioritization -> Governance -> Delivery",
            "Single intake channel (Jira) as source of truth",
            "Unified Intake Process with standard request form",
            "Weekly Replenishment & Triage cadence",
            "Portfolio Visibility flagged as OUT OF SCOPE",
        ],
        "portfolio_src": "NextGen DTO Deliverable, Slides 5-7, 19-20, 32",
    },
    "Infrastructure": {
        "color": PINK,
        "team": [
            "Lean & Agile proposed practices (proposal stage)",
            "No specific team-level prioritization formula documented",
            "Service Financing Model per unit served",
        ],
        "team_src": "Agile Management Model Proposal, Slide 38",
        "product": [
            "Portfolio Roadmap to visually manage initiatives",
            "Technical approach with Lean & Agile practices",
            "No detailed product-level prioritization documented",
        ],
        "product_src": "Agile Management Model Proposal, Slides 38, 5",
        "portfolio": [
            "Investment model per unit served",
            "Proposal stage — not yet implemented",
            "No portfolio governance framework documented",
        ],
        "portfolio_src": "Agile Management Model Proposal, Slides 51-55",
    },
    "IT CS (Digital Products)": {
        "color": RED,
        "team": [
            "Currently being discovered (this project)",
            "Governance forums identified (5 active)",
            "No formal prioritization formula documented yet",
        ],
        "team_src": "IT CS Operating Model V1 + Support Model (just ingested)",
        "product": [
            "Operating model under construction",
            "Support model defines L1/L2/L3 tiers",
            "Product-level prioritization: TO BE DEFINED",
        ],
        "product_src": "IT CS Operating Model V1",
        "portfolio": [
            "No portfolio-level prioritization documented",
            "This is what we will help define in Wave 1",
            "Gap: biggest opportunity for the pilot",
        ],
        "portfolio_src": "IT CS Operating Model V1 + RFI Data",
    },
}

# ═══════════════════════════════════════════════════════════════════════
# RENDER 3 LEVEL SECTIONS
# ═══════════════════════════════════════════════════════════════════════
for level_name, level_desc, level_color in levels:
    # Section header
    rrect(M, Y, FW, 60, fill=(*level_color, 20), outline=level_color, r=12)
    d.text((M+20, Y+14), level_name, fill=level_color, font=FS)
    d.text((M+300, Y+18), level_desc, fill=DIM, font=FB)
    Y += 72

    level_key = level_name.split()[0].lower()  # "team", "product", "portfolio"

    for pi, (proj_name, proj_data) in enumerate(projects.items()):
        items = proj_data[level_key]
        src = proj_data[f"{level_key}_src"]
        pc = proj_data["color"]

        # Calculate row height
        row_h = max(len(items) * 32 + 50, 120)

        bg = ROW_A if pi % 2 == 0 else ROW_B
        rrect(M, Y, FW, row_h, fill=bg, outline=BORD, r=10)

        # Project name tag
        tag(M+16, Y+14, proj_name, pc)

        # Items
        ix = M + 280
        iy = Y + 14
        for item in items:
            # Color bullet
            bullet_color = GREEN if "PBI" in item or "Formula" in item or "Kanban" in item else DIM
            if "NOT" in item.upper() or "NO " in item.upper() or "OUT OF SCOPE" in item.upper():
                bullet_color = RED
            if "TO BE DEFINED" in item.upper() or "being discovered" in item.lower():
                bullet_color = ORANGE

            d.text((ix, iy), "•", fill=bullet_color, font=FTD)
            d.text((ix+18, iy), item, fill=WHITE if bullet_color != RED else (180,120,120), font=FTD)
            iy += 32

        # Source reference
        source_tag(M + FW - 520, Y + row_h - 26, f"Source: {src}")

        Y += row_h + 8

    Y += GAP

# ═══════════════════════════════════════════════════════════════════════
# LEGEND
# ═══════════════════════════════════════════════════════════════════════
Y += 10
rrect(M, Y, FW, 70, fill=CARD, outline=BORD, r=12)
lx = M + 20
d.text((lx, Y+10), "MATURITY LEGEND:", fill=WHITE, font=FL)
lx += 260
for label, col in [("Defined & Implemented", GREEN), ("Under Construction", ORANGE),
                    ("Gap / Not Documented", RED), ("Source: KB visual chunk", MUTED)]:
    d.ellipse([lx, Y+18, lx+14, Y+32], fill=col)
    d.text((lx+20, Y+14), label, fill=DIM, font=FTD)
    lx += d.textlength(label, font=FTD) + 50

Y += 70 + M

# ═══════════════════════════════════════════════════════════════════════
# CROP + SAVE
# ═══════════════════════════════════════════════════════════════════════
img = img.crop((0, 0, W, Y))
out = "/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Prioritization_Comparison_3_Levels.png"
img.save(out, "PNG")
px = img.size[0] * img.size[1]
print(f"Saved {img.size[0]}x{img.size[1]} — {px/1e6:.1f}MP → {out}")
