"""
All 4 stages — full topic checklists per dimension.
One PNG per stage, each card has title + writable area, no sub-title under stage header.
"""
from PIL import Image, ImageDraw, ImageFont
from pathlib import Path

OUT = Path("/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Miro_Visuals")

NAVY_DARK   = (21, 42, 71)
GREEN       = (76, 140, 63)
TEAL        = (43, 107, 119)
AMBER       = (232, 139, 47)
RED         = (182, 59, 31)
NAVY        = (44, 90, 136)
WHITE       = (255, 255, 255)
GREY_TXT    = (75, 85, 99)
SLOT_BG     = (255, 255, 255)
SLOT_BORDER = (200, 208, 220)


# Dimension palette is shared across stages
DIM_COLORS = {
    "CAPABILITIES": GREEN,
    "ROLES":        TEAL,
    "TOOLS":        AMBER,
    "GOVERNANCE":   NAVY,
    "PERFORMANCE":  RED,
}

# Cols per dimension (controls grid layout)
DIM_COLS = {
    "CAPABILITIES": 5,
    "ROLES":        4,
    "TOOLS":        5,
    "GOVERNANCE":   3,
    "PERFORMANCE":  3,
}


# ── Topic lists per stage / dimension ──────────────────────────────────────
STAGES = {
    "01": ("DEMAND MANAGEMENT", {
        "CAPABILITIES": [
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
        ],
        "ROLES": [
            "Business requestors / stakeholders",
            "Product or demand managers",
            "Architects or technical reviewers",
            "Triage / intake committee",
        ],
        "TOOLS": [
            "Demand capture tools",
            "Demand tracking & backlog tools",
            "Communication channels",
            "Documentation tools (request details)",
            "Reporting / visibility tools",
        ],
        "GOVERNANCE": [
            "Demand review forums",
            "Decision-making practices (accept / reject / redirect)",
            "Escalation mechanisms (urgent / conflicting requests)",
            "Definition of intake guidelines or policies",
            "Alignment mechanisms with business stakeholders",
            "Oversight of demand inflow and backlog health",
        ],
        "PERFORMANCE": [
            "Demand volume and inflow trends",
            "Demand quality (completeness, clarity)",
            "Triage lead time / time-of-demand",
            "% demand entering through formal channels",
            "% urgent or bypassed requests",
            "Demand duplication / overlap indicators",
            "Visibility and traceability of demand status",
        ],
    }),

    "02": ("PREPARATION & PLANNING", {
        "CAPABILITIES": [
            "Demand consolidation and portfolio structuring",
            "Demand prioritization mechanisms",
            "Use of value, risk and effort considerations",
            "Portfolio backlog management",
            "Roadmap definition (short / mid / long term)",
            "Planning cycles (monthly, quarterly)",
            "Effort estimation practices",
            "Capacity planning and forecasting",
            "Alignment between demand and available capacity",
            "Work allocation to teams / squads",
            "Handling of unplanned or ad-hoc demand",
            "Grouping and structuring of work (epics, features, initiatives)",
        ],
        "ROLES": [
            "Product managers / product owners",
            "Technical leads / architects",
            "Business stakeholders",
        ],
        "TOOLS": [
            "Backlog and portfolio management tools",
            "Roadmapping and planning tools",
            "Estimation and capacity planning tools",
            "Reporting and visualization tools",
        ],
        "GOVERNANCE": [
            "Portfolio review forums",
            "Prioritization & planning decision-making practices",
            "Alignment mechanisms between business and technology",
            "Definition of prioritization criteria and guidelines",
            "Oversight of demand vs capacity balance",
            "Decision practices for deferring or rejecting demand",
            "Cross-area coordination mechanisms (multiple OPCOs / domains)",
        ],
        "PERFORMANCE": [
            "Demand vs capacity visibility",
            "Prioritization cycle time",
            "% demand planned vs unplanned",
            "Alignment between planned and committed work",
            "Planning predictability",
            "Portfolio balance (run / change / regulatory)",
            "Visibility of roadmap and planning outcomes",
        ],
    }),

    "03": ("EXECUTION", {
        "CAPABILITIES": [
            "Intake of work from prioritized backlog",
            "Work breakdown (epics, stories, tasks)",
            "Sprint / iteration management",
            "Work prioritization at team level",
            "Definition of Ready (or equivalent)",
            "Definition of Done (or completion criteria)",
            "Development and build practices",
            "Testing practices (unit, integration, functional, UAT)",
            "Defect management",
            "Dependency management (intra- and inter-team)",
            "Cross-team coordination",
            "Work progress tracking and visibility",
            "Documentation of functional and technical aspects",
        ],
        "ROLES": [
            "Delivery teams",
            "Technical leads",
            "Scrum masters / delivery managers",
            "Product owners / business representatives",
        ],
        "TOOLS": [
            "Work management tools",
            "Collaboration and communication tools",
            "Documentation tools",
        ],
        "GOVERNANCE": [
            "Delivery review forums (sprint reviews, demos)",
            "Issue and risk escalation practices",
            "Oversight of delivery progress and performance",
            "Definition of delivery standards and guidelines",
            "Coordination mechanisms across teams or domains",
            "Alignment with business during execution",
        ],
        "PERFORMANCE": [
            "Velocity / throughput",
            "Lead time and cycle time",
            "Predictability (planned vs delivered)",
            "Defect rates and quality indicators",
            "Work in progress (WIP) levels",
            "Dependency-related delays",
            "Visibility of execution progress",
        ],
    }),

    "04": ("DEPLOYMENT", {
        "CAPABILITIES": [
            "Release planning and scheduling",
            "Release coordination across teams",
            "Change management practices",
            "Deployment execution (manual or automated)",
            "Environment management (dev, test, prod)",
            "Rollback and recovery capabilities",
            "Validation and sign-off processes",
            "Post-deployment verification",
            "Hypercare and stabilization support",
            "Transition to operations and support teams",
        ],
        "ROLES": [
            "Release managers",
            "Operations / support teams",
            "Delivery teams (as contributors to release)",
            "Business stakeholders (validation / sign-off)",
        ],
        "TOOLS": [
            "Release and deployment tools",
            "Monitoring and alerting tools",
            "Incident and support management tools",
        ],
        "GOVERNANCE": [
            "Release and change approval forums (CAB or equivalent)",
            "Decision-making practices for release readiness",
            "Oversight of deployment risks and impacts",
            "Definition of deployment and change policies",
            "Coordination between delivery and operations",
            "Control mechanisms for production changes",
        ],
        "PERFORMANCE": [
            "Deployment success rate",
            "Incidents and defects post-release",
            "Time to recover from failures",
            "Release frequency",
            "Stability of deployed solutions",
            "Effectiveness of hypercare period",
            "Visibility of release outcomes",
        ],
    }),
}


# ── Layout constants ───────────────────────────────────────────────────────
W = 3600
margin_x = 120
header_h = 280            # top header (no subtitle line)
dim_header_h = 140
gap_between_dims = 70
gap_card = 26
card_h = 460
title_band_h = 110
footer_h = 120


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


def build_stage(num, title, dims_dict):
    # height calc
    total = 0
    for dim_name in DIM_COLORS.keys():
        items = dims_dict.get(dim_name, [])
        cols = DIM_COLS[dim_name]
        rows = (len(items) + cols - 1) // cols
        total += dim_header_h + rows * card_h + (rows - 1) * gap_card
    H = header_h + total + (len(DIM_COLORS) - 1) * gap_between_dims + footer_h + gap_between_dims

    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)

    # page header
    d.rectangle([0, 0, W, header_h], fill=NAVY_DARK)
    d.text((margin_x, 60), f"STAGE {num}", font=fnt(54, True), fill=AMBER)
    d.text((margin_x, 135), title, font=fnt(110, True), fill=WHITE)

    y = header_h + gap_between_dims
    for dim_name in DIM_COLORS.keys():
        items = dims_dict.get(dim_name, [])
        color = DIM_COLORS[dim_name]
        cols = DIM_COLS[dim_name]
        rows = (len(items) + cols - 1) // cols

        # banner
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

            d.rounded_rectangle([cx0, cy0, cx1, cy1], radius=22,
                                fill=SLOT_BG, outline=SLOT_BORDER, width=3)
            d.rounded_rectangle([cx0, cy0, cx1, cy0 + title_band_h], radius=22,
                                fill=color)
            d.rectangle([cx0, cy0 + title_band_h - 22, cx1, cy0 + title_band_h], fill=color)

            max_chars = max(8, (card_w - 30) // 14)
            lines = wrap_text(item, max_chars)[:3]
            line_h = 30
            block_h = len(lines) * line_h
            start_y = cy0 + (title_band_h - block_h) / 2 + line_h / 2
            for li, line in enumerate(lines):
                centered(d, ((cx0 + cx1) / 2, start_y + li * line_h),
                         line, fnt(22, True), WHITE)

        y = cards_y0 + rows * card_h + (rows - 1) * gap_card + gap_between_dims

    # footer
    d.text((margin_x, H - 70),
           f"Avangrid · Digital Products · Target Operating Model — Stage {num} {title.title()}",
           font=fnt(26), fill=GREY_TXT)
    d.text((W - 750, H - 70),
           f"Workshop board · {int(num)} of 4",
           font=fnt(26, True), fill=AMBER)

    safe = title.replace(" & ", "And").replace(" ", "")
    out_file = OUT / f"{num}_TOM_{safe}_FULL.png"
    img.save(out_file, "PNG", dpi=(150, 150))

    # preview
    ratio = 1400 / W
    preview = img.resize((1400, int(H * ratio)), Image.LANCZOS)
    preview.save(out_file.with_name(out_file.stem + "_preview.png"), "PNG")
    return out_file, (W, H)


OUT.mkdir(parents=True, exist_ok=True)
for num in ("01", "02", "03", "04"):
    title, dims = STAGES[num]
    f, size = build_stage(num, title, dims)
    print(f"✓ {f.name}  {size[0]}×{size[1]}")
