"""
Build 5 Miro-ready visuals for the Avangrid Digital Products Target Operating Model.

Output: 5 PNG files (1920×1080) in same folder.
  01_TOM_Big_Picture.png            — 4 stages × 5 dimensions matrix
  02_Demand_Single_Front_Door.png   — funnel + classification (Run/Grow/Transform · Core/Satellite)
  03_Product_vs_Application.png     — M2C as product spanning SAP modules
  04_Dual_Track.png                 — Discovery + Delivery side-by-side
  05_Roles_Stack.png                — Product team layers + Heather PM area
"""
from PIL import Image, ImageDraw, ImageFont
from pathlib import Path

OUT = Path("/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Miro_Visuals")

# ---------- Palette (Avangrid + NTT) ----------
NAVY      = (44, 90, 136)
NAVY_DEEP = (33, 70, 108)
NAVY_DARK = (21, 42, 71)
TEAL      = (43, 107, 119)
GREEN     = (76, 140, 63)
GREEN_DK  = (59, 125, 35)
AMBER     = (232, 139, 47)
AMBER_DK  = (217, 110, 37)
RED       = (182, 59, 31)
LIGHT_BG  = (244, 245, 247)
GREY_BG   = (227, 230, 235)
GREY_TXT  = (75, 85, 99)
DARK_TXT  = (26, 31, 44)
WHITE     = (255, 255, 255)
BORDER    = (216, 221, 227)

W, H = 1920, 1080


def fnt(size, bold=False):
    """Best-effort font loading."""
    candidates = [
        f"/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf",
        f"/System/Library/Fonts/Helvetica.ttc",
        "/Library/Fonts/Arial.ttf",
    ]
    for c in candidates:
        try:
            return ImageFont.truetype(c, size)
        except (OSError, IOError):
            continue
    return ImageFont.load_default()


def text_size(draw, text, font):
    bbox = draw.textbbox((0, 0), text, font=font)
    return bbox[2] - bbox[0], bbox[3] - bbox[1]


def wrap_text(draw, text, font, max_width):
    words = text.split()
    lines, line = [], ""
    for w in words:
        test = (line + " " + w).strip()
        tw, _ = text_size(draw, test, font)
        if tw <= max_width:
            line = test
        else:
            lines.append(line)
            line = w
    if line:
        lines.append(line)
    return lines


def draw_rect(draw, xy, fill=None, outline=None, width=1, radius=0):
    if radius:
        draw.rounded_rectangle(xy, radius=radius, fill=fill, outline=outline, width=width)
    else:
        draw.rectangle(xy, fill=fill, outline=outline, width=width)


def draw_centered_text(draw, xy, text, font, color, anchor="mm"):
    draw.text(xy, text, fill=color, font=font, anchor=anchor)


# =========================================================
# 01 — TOM BIG PICTURE (4 stages × 5 dimensions)
# =========================================================
def visual_01_big_picture():
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)

    # Header
    d.rectangle([0, 0, W, 90], fill=NAVY)
    d.text((40, 22), "TARGET OPERATING MODEL — DIGITAL PRODUCTS", font=fnt(34, True), fill=WHITE)
    d.text((40, 62), "4 stages × 5 dimensions · framework grounded in TOM reference",
           font=fnt(16), fill=(199, 211, 228))

    stages = [
        ("DEMAND MANAGEMENT", "What demand looks like at intake"),
        ("PREPARATION & PLANNING", "How demand becomes a roadmap"),
        ("EXECUTION", "How work gets built"),
        ("DEPLOYMENT", "How releases land in production"),
    ]
    dims = [
        ("CAPABILITIES", GREEN),
        ("ROLES", TEAL),
        ("TOOLS", AMBER),
        ("GOVERNANCE", NAVY),
        ("PERFORMANCE", RED),
    ]

    cell_data = {
        ("DEMAND MANAGEMENT", "CAPABILITIES"): [
            "Single Front Door (5 types)", "Demand shaping (problem-impact-value)",
            "Classification: Run/Grow/Transform", "Core SAP vs Satellites",
            "Lightweight Discovery (problem · personas · hypothesis)"
        ],
        ("DEMAND MANAGEMENT", "ROLES"): [
            "Business Requestors", "Product / Demand Managers", "Architects / Tech Reviewers"
        ],
        ("DEMAND MANAGEMENT", "TOOLS"): [
            "Capture · tracking", "Single backlog", "Communication channels"
        ],
        ("DEMAND MANAGEMENT", "GOVERNANCE"): [
            "Demand review forums", "Accept/Reject/Redirect rules",
            "Escalation mechanisms", "Intake guidelines"
        ],
        ("DEMAND MANAGEMENT", "PERFORMANCE"): [
            "Demand inflow trends", "Demand quality", "Triage lead time",
            "% via formal channels"
        ],

        ("PREPARATION & PLANNING", "CAPABILITIES"): [
            "Product definition (Products ≠ Apps)", "Outcomes (not deliverables)",
            "Rolling-wave roadmap (3/6/12)", "Capacity planning",
            "Value · Risk · CoD prioritization"
        ],
        ("PREPARATION & PLANNING", "ROLES"): [
            "Product Managers / POs", "Tech Leads / Architects", "Business Stakeholders"
        ],
        ("PREPARATION & PLANNING", "TOOLS"): [
            "Backlog · portfolio mgmt", "Roadmap & planning", "Capacity / estimation"
        ],
        ("PREPARATION & PLANNING", "GOVERNANCE"): [
            "Portfolio review forums", "Prioritization criteria",
            "Demand vs capacity oversight", "Cross-OPCO coordination"
        ],
        ("PREPARATION & PLANNING", "PERFORMANCE"): [
            "Demand vs capacity visibility", "% planned vs unplanned",
            "Planning predictability", "Portfolio balance Run/Grow/Trans"
        ],

        ("EXECUTION", "CAPABILITIES"): [
            "Stable product teams", "Backlog by product (not app)",
            "Dual-track (Discovery + Delivery)", "DoR / DoD",
            "Dependency mgmt (SAP + vendors)"
        ],
        ("EXECUTION", "ROLES"): [
            "Product Owners", "Tech Leads", "Delivery teams",
            "Scrum masters / delivery mgrs"
        ],
        ("EXECUTION", "TOOLS"): [
            "Work mgmt (Jira)", "Collaboration", "Documentation"
        ],
        ("EXECUTION", "GOVERNANCE"): [
            "Sprint reviews / demos", "Risk escalation",
            "Cross-team coordination", "Delivery standards"
        ],
        ("EXECUTION", "PERFORMANCE"): [
            "Velocity / throughput", "Lead time / cycle time",
            "Predictability", "Defect rates · WIP"
        ],

        ("DEPLOYMENT", "CAPABILITIES"): [
            "Release scheduling", "Hypercare formal", "Rollback / recovery",
            "Measure & Learn (post-release)", "Transition to Ops"
        ],
        ("DEPLOYMENT", "ROLES"): [
            "Release Managers", "Operations / Support", "Business sign-off"
        ],
        ("DEPLOYMENT", "TOOLS"): [
            "Release & deployment", "Monitoring · alerting", "Incident mgmt"
        ],
        ("DEPLOYMENT", "GOVERNANCE"): [
            "CAB / release approval", "Change control",
            "Deployment risk oversight"
        ],
        ("DEPLOYMENT", "PERFORMANCE"): [
            "Deployment success rate", "MTTR · post-release defects",
            "Release frequency", "Hypercare effectiveness"
        ],
    }

    # Layout
    margin_x = 40
    margin_top = 110
    n_cols = 4
    n_rows = 5
    label_col_w = 140
    available_w = W - 2 * margin_x - label_col_w
    cell_w = available_w // n_cols
    available_h = H - margin_top - 50
    header_h = 70
    cell_h = (available_h - header_h) // n_rows

    # Stage headers
    for i, (stage, sub) in enumerate(stages):
        x0 = margin_x + label_col_w + i * cell_w
        d.rectangle([x0 + 2, margin_top, x0 + cell_w - 2, margin_top + header_h], fill=NAVY_DARK)
        draw_centered_text(d, (x0 + cell_w / 2, margin_top + 22),
                           f"0{i+1} · {stage}", fnt(14, True), WHITE)
        draw_centered_text(d, (x0 + cell_w / 2, margin_top + 50),
                           sub, fnt(11), (199, 211, 228))

    # Dimension labels (left col) + cell content
    for r, (dim, dim_color) in enumerate(dims):
        y0 = margin_top + header_h + r * cell_h
        # left label
        d.rectangle([margin_x, y0 + 2, margin_x + label_col_w - 4, y0 + cell_h - 4], fill=dim_color)
        draw_centered_text(d, (margin_x + label_col_w / 2, y0 + cell_h / 2),
                           dim, fnt(13, True), WHITE)

        for c, (stage, _) in enumerate(stages):
            x0 = margin_x + label_col_w + c * cell_w
            d.rounded_rectangle(
                [x0 + 4, y0 + 4, x0 + cell_w - 4, y0 + cell_h - 4],
                radius=6, fill=LIGHT_BG, outline=BORDER, width=1
            )
            items = cell_data.get((stage, dim), [])
            for idx, item in enumerate(items):
                ty = y0 + 14 + idx * 16
                if ty > y0 + cell_h - 14:
                    break
                d.text((x0 + 12, ty), f"• {item}", fill=DARK_TXT, font=fnt(10))

    # Footer
    d.text((margin_x, H - 30), "Avangrid · Digital Products · TOM v1.0",
           fill=GREY_TXT, font=fnt(11))
    d.text((W - 380, H - 30), "Output of next 2 weeks · ready for review with Heather",
           fill=AMBER, font=fnt(11, True))

    img.save(OUT / "01_TOM_Big_Picture.png", "PNG", dpi=(150, 150))
    print("✓ 01_TOM_Big_Picture.png")


# =========================================================
# 02 — DEMAND · SINGLE FRONT DOOR + CLASSIFICATION
# =========================================================
def visual_02_demand_funnel():
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)

    d.rectangle([0, 0, W, 90], fill=NAVY)
    d.text((40, 22), "01 · DEMAND MANAGEMENT — Single Front Door + Classification",
           font=fnt(32, True), fill=WHITE)
    d.text((40, 62), "Every request enters one funnel · gets shaped · classified · routed",
           font=fnt(16), fill=(199, 211, 228))

    # 5 demand sources (top)
    sources = [
        ("INCIDENTS", "Production failures · break/fix", RED),
        ("IMPROVEMENTS", "Small/medium evolutions", AMBER),
        ("REGULATORY", "Compliance · rate cases · external mandate", NAVY_DEEP),
        ("INNOVATION", "New capabilities · POCs", GREEN),
        ("TECHNICAL DEBT", "Refactor · upgrades · platform health", TEAL),
    ]
    src_y = 130
    src_h = 110
    src_gap = 20
    src_total = W - 80 - src_gap * 4
    src_w = src_total // 5

    for i, (label, sub, color) in enumerate(sources):
        x0 = 40 + i * (src_w + src_gap)
        d.rounded_rectangle([x0, src_y, x0 + src_w, src_y + src_h], radius=12,
                            fill=color, outline=None)
        draw_centered_text(d, (x0 + src_w / 2, src_y + 32), label, fnt(18, True), WHITE)
        # wrap subtitle
        wrapped = wrap_text(d, sub, fnt(12), src_w - 30)
        for li, ln in enumerate(wrapped[:2]):
            draw_centered_text(d, (x0 + src_w / 2, src_y + 65 + li * 17), ln, fnt(12), WHITE)

    # Funnel (top wide → narrow)
    fnl_top_y = src_y + src_h + 30
    fnl_bot_y = fnl_top_y + 120
    fnl_top_left = 100
    fnl_top_right = W - 100
    fnl_bot_left = W // 2 - 220
    fnl_bot_right = W // 2 + 220
    d.polygon([
        (fnl_top_left, fnl_top_y), (fnl_top_right, fnl_top_y),
        (fnl_bot_right, fnl_bot_y), (fnl_bot_left, fnl_bot_y)
    ], fill=NAVY_DEEP)
    draw_centered_text(d, (W / 2, fnl_top_y + 38), "SINGLE FRONT DOOR", fnt(24, True), WHITE)
    draw_centered_text(d, (W / 2, fnl_top_y + 78),
                       "Demand Shaping  ·  problem → impact → value  ·  reject 'solutions'",
                       fnt(14, True), AMBER)

    # Classification box (below funnel)
    cls_y = fnl_bot_y + 30
    cls_h = 280
    cls_left, cls_right = 60, W - 60
    d.rounded_rectangle([cls_left, cls_y, cls_right, cls_y + cls_h], radius=14,
                        fill=LIGHT_BG, outline=BORDER, width=2)
    draw_centered_text(d, ((cls_left + cls_right) / 2, cls_y + 28),
                       "CLASSIFICATION — three orthogonal dimensions",
                       fnt(20, True), NAVY)

    # 3 classification dimensions
    dim_w = (cls_right - cls_left - 80) // 3
    dim_y = cls_y + 70
    dim_h = cls_h - 95

    # Dimension 1: Run / Grow / Transform
    x0 = cls_left + 20
    d.rounded_rectangle([x0, dim_y, x0 + dim_w, dim_y + dim_h], radius=10,
                        fill=WHITE, outline=GREEN, width=2)
    draw_centered_text(d, (x0 + dim_w / 2, dim_y + 25),
                       "RUN / GROW / TRANSFORM", fnt(15, True), GREEN_DK)
    rgt = [
        ("RUN", "Keep the lights on · break/fix · BAU", AMBER),
        ("GROW", "Evolve existing products · enhancements", TEAL),
        ("TRANSFORM", "New capabilities · strategic bets", GREEN_DK),
    ]
    for j, (lab, sub, c) in enumerate(rgt):
        y = dim_y + 60 + j * 60
        d.rounded_rectangle([x0 + 20, y, x0 + dim_w - 20, y + 50], radius=6,
                            fill=c, outline=None)
        d.text((x0 + 35, y + 8), lab, fill=WHITE, font=fnt(14, True))
        d.text((x0 + 35, y + 28), sub, fill=WHITE, font=fnt(11))

    # Dimension 2: Core SAP vs Satellite
    x0 = cls_left + 40 + dim_w
    d.rounded_rectangle([x0, dim_y, x0 + dim_w, dim_y + dim_h], radius=10,
                        fill=WHITE, outline=NAVY, width=2)
    draw_centered_text(d, (x0 + dim_w / 2, dim_y + 25),
                       "CORE SAP  vs  SATELLITE", fnt(15, True), NAVY)
    css = [
        ("CORE SAP", "Billing · FICA · CRM · O2C · M2C", NAVY_DEEP),
        ("SATELLITE", "Add-ons · integrations · third-party", TEAL),
        ("SHARED", "Cross-cutting · platform · security", AMBER_DK),
    ]
    for j, (lab, sub, c) in enumerate(css):
        y = dim_y + 60 + j * 60
        d.rounded_rectangle([x0 + 20, y, x0 + dim_w - 20, y + 50], radius=6,
                            fill=c, outline=None)
        d.text((x0 + 35, y + 8), lab, fill=WHITE, font=fnt(14, True))
        d.text((x0 + 35, y + 28), sub, fill=WHITE, font=fnt(11))

    # Dimension 3: Size / Urgency
    x0 = cls_left + 60 + 2 * dim_w
    d.rounded_rectangle([x0, dim_y, x0 + dim_w, dim_y + dim_h], radius=10,
                        fill=WHITE, outline=AMBER, width=2)
    draw_centered_text(d, (x0 + dim_w / 2, dim_y + 25),
                       "SIZE  ·  URGENCY  ·  CoD", fnt(15, True), AMBER_DK)
    sz = [
        ("SMALL", "<40h · Jira · self-service", GREEN),
        ("MEDIUM", "40-200h · Jira · squad backlog", TEAL),
        ("LARGE", ">200h · PPM Now · BRP cycle", NAVY_DEEP),
    ]
    for j, (lab, sub, c) in enumerate(sz):
        y = dim_y + 60 + j * 60
        d.rounded_rectangle([x0 + 20, y, x0 + dim_w - 20, y + 50], radius=6,
                            fill=c, outline=None)
        d.text((x0 + 35, y + 8), lab, fill=WHITE, font=fnt(14, True))
        d.text((x0 + 35, y + 28), sub, fill=WHITE, font=fnt(11))

    # Footer principle
    pr_y = cls_y + cls_h + 25
    d.rectangle([60, pr_y, W - 60, pr_y + 50], fill=NAVY_DARK)
    draw_centered_text(d, (W / 2, pr_y + 25),
                       "PRIORITIZATION = Business Value × (1 / Cost-of-Delay) × Regulatory Weight",
                       fnt(16, True), WHITE)

    img.save(OUT / "02_Demand_Single_Front_Door.png", "PNG", dpi=(150, 150))
    print("✓ 02_Demand_Single_Front_Door.png")


# =========================================================
# 03 — PRODUCT vs APPLICATION (M2C example)
# =========================================================
def visual_03_product_vs_app():
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)

    d.rectangle([0, 0, W, 90], fill=NAVY)
    d.text((40, 22), "PRODUCT  ≠  APPLICATION — Reframing the unit of accountability",
           font=fnt(30, True), fill=WHITE)
    d.text((40, 62), "Example: Meter-to-Cash (M2C) is ONE product spanning many SAP modules + apps",
           font=fnt(16), fill=(199, 211, 228))

    # LEFT side — "Today: app-centric"
    left_x = 60
    box_w = 820
    today_y = 130

    d.rounded_rectangle([left_x, today_y, left_x + box_w, today_y + 880], radius=14,
                        fill=LIGHT_BG, outline=RED, width=3)
    draw_centered_text(d, (left_x + box_w / 2, today_y + 38),
                       "❌  TODAY — APPLICATION-OWNER MINDSET",
                       fnt(22, True), RED)
    d.text((left_x + 40, today_y + 75), "Each app has a 'champion' (Tech Lead). No one owns the joint customer outcome.", font=fnt(13), fill=GREY_TXT)

    apps = [
        ("SAP SD", "Order entry"),
        ("SAP MM", "Materials"),
        ("SAP FI", "Finance posting"),
        ("SAP IS-U", "Utility billing"),
        ("CRM (Avangrid US)", "Customer interactions"),
        ("StreamServe", "Invoice formatting"),
        ("AMI Head-End", "Meter data acquisition"),
        ("Group 8760 (EDI)", "B2B exchange"),
        ("SOVOS Taxware", "Tax compliance"),
    ]
    cols, rows = 3, 3
    cell_pad = 14
    cw = (box_w - 80 - (cols - 1) * cell_pad) // cols
    ch = 80
    grid_y = today_y + 130
    for i, (app, sub) in enumerate(apps):
        c, r = i % cols, i // cols
        x = left_x + 40 + c * (cw + cell_pad)
        y = grid_y + r * (ch + cell_pad)
        d.rounded_rectangle([x, y, x + cw, y + ch], radius=8, fill=WHITE, outline=BORDER, width=1)
        # App badge
        d.rounded_rectangle([x + 12, y + 12, x + cw - 12, y + 36], radius=4, fill=NAVY_DEEP)
        draw_centered_text(d, (x + cw / 2, y + 24), app, fnt(13, True), WHITE)
        d.text((x + 18, y + 45), sub, fill=GREY_TXT, font=fnt(11))
        d.text((x + 18, y + 60), "Owner: Tech Lead", fill=RED, font=fnt(10, True))

    d.rectangle([left_x + 40, today_y + 770, left_x + box_w - 40, today_y + 830], fill=RED)
    draw_centered_text(d, (left_x + box_w / 2, today_y + 800),
                       "9 owners · 0 product accountability for end-to-end M2C",
                       fnt(15, True), WHITE)

    # RIGHT side — "TO-BE: product-centric"
    right_x = W - 60 - box_w
    d.rounded_rectangle([right_x, today_y, right_x + box_w, today_y + 880], radius=14,
                        fill=LIGHT_BG, outline=GREEN_DK, width=3)
    draw_centered_text(d, (right_x + box_w / 2, today_y + 38),
                       "✅  TO-BE — PRODUCT-CENTRIC",
                       fnt(22, True), GREEN_DK)
    d.text((right_x + 40, today_y + 75), "ONE Product Manager owns the M2C outcome end-to-end. Apps become components.", font=fnt(13), fill=GREY_TXT)

    # Big M2C product box
    prod_y = today_y + 130
    prod_h = 130
    d.rounded_rectangle([right_x + 40, prod_y, right_x + box_w - 40, prod_y + prod_h],
                        radius=12, fill=GREEN_DK, outline=None)
    draw_centered_text(d, (right_x + box_w / 2, prod_y + 30),
                       "PRODUCT: Meter-to-Cash (M2C)", fnt(22, True), WHITE)
    draw_centered_text(d, (right_x + box_w / 2, prod_y + 65),
                       "Outcome: every kWh consumed → metered → billed → collected",
                       fnt(13), WHITE)
    draw_centered_text(d, (right_x + box_w / 2, prod_y + 95),
                       "Owner: PRODUCT MANAGER (under Heather's PM area · TO-BE)",
                       fnt(13, True), AMBER)

    # Sub-products / value-stream stages
    flow_y = prod_y + prod_h + 30
    stages = [
        ("METER", "AMI Head-End · meter data", TEAL),
        ("BILL", "SAP IS-U · StreamServe", NAVY_DEEP),
        ("COLLECT", "SAP FI · SOVOS · Precisely", AMBER_DK),
        ("INTERACT", "CRM · HCL Nuance · Medallia", GREEN_DK),
    ]
    sw = (box_w - 80 - 30) // 4
    sh = 110
    for i, (lab, comps, c) in enumerate(stages):
        x = right_x + 40 + i * (sw + 10)
        d.rounded_rectangle([x, flow_y, x + sw, flow_y + sh], radius=10, fill=c, outline=None)
        draw_centered_text(d, (x + sw / 2, flow_y + 28), lab, fnt(16, True), WHITE)
        wrapped = wrap_text(d, comps, fnt(11), sw - 24)
        for li, ln in enumerate(wrapped[:3]):
            draw_centered_text(d, (x + sw / 2, flow_y + 55 + li * 16), ln, fnt(11), WHITE)

    # Arrows between stages
    arr_y = flow_y + sh / 2
    for i in range(3):
        ax = right_x + 40 + (i + 1) * sw + i * 10 + 5
        d.polygon([(ax - 5, arr_y - 8), (ax + 5, arr_y), (ax - 5, arr_y + 8)], fill=DARK_TXT)

    # Roles under TO-BE
    roles_y = flow_y + sh + 40
    d.text((right_x + 40, roles_y), "Stable product team:", font=fnt(13, True), fill=DARK_TXT)
    role_list = [
        ("Product Manager", "Owns M2C outcome · roadmap · value", AMBER),
        ("Product Owner", "Owns backlog · sprint priorities", TEAL),
        ("Tech Lead", "Owns architecture · solution design", NAVY_DEEP),
        ("Squad", "PO + Devs + QA + UX (stable team)", GREEN_DK),
    ]
    for i, (role, sub, c) in enumerate(role_list):
        y = roles_y + 30 + i * 65
        d.rounded_rectangle([right_x + 40, y, right_x + box_w - 40, y + 55], radius=8,
                            fill=WHITE, outline=BORDER, width=1)
        d.rounded_rectangle([right_x + 40, y, right_x + 200, y + 55], radius=8, fill=c)
        draw_centered_text(d, (right_x + 120, y + 27), role, fnt(13, True), WHITE)
        d.text((right_x + 215, y + 18), sub, fill=DARK_TXT, font=fnt(12))

    # Bottom callout
    cb_y = today_y + 825
    d.rectangle([right_x + 40, cb_y, right_x + box_w - 40, cb_y + 50], fill=GREEN_DK)
    draw_centered_text(d, (right_x + box_w / 2, cb_y + 25),
                       "1 product · 1 owner · 1 backlog · 1 outcome",
                       fnt(15, True), WHITE)

    img.save(OUT / "03_Product_vs_Application.png", "PNG", dpi=(150, 150))
    print("✓ 03_Product_vs_Application.png")


# =========================================================
# 04 — DUAL-TRACK (Discovery + Delivery)
# =========================================================
def visual_04_dual_track():
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)

    d.rectangle([0, 0, W, 90], fill=NAVY)
    d.text((40, 22), "DUAL-TRACK — Discovery + Delivery running in parallel",
           font=fnt(30, True), fill=WHITE)
    d.text((40, 62), "Discovery answers WHAT & WHY · Delivery answers HOW & WHEN · both run continuously",
           font=fnt(16), fill=(199, 211, 228))

    # Two horizontal tracks
    track_h = 280
    margin_x = 60
    label_w = 220

    # DISCOVERY track (top)
    disc_y = 140
    d.rounded_rectangle([margin_x, disc_y, W - margin_x, disc_y + track_h], radius=14,
                        fill=LIGHT_BG, outline=AMBER_DK, width=3)
    d.rounded_rectangle([margin_x, disc_y, margin_x + label_w, disc_y + track_h], radius=14,
                        fill=AMBER_DK)
    draw_centered_text(d, (margin_x + label_w / 2, disc_y + 80), "DISCOVERY", fnt(28, True), WHITE)
    draw_centered_text(d, (margin_x + label_w / 2, disc_y + 120), "What & Why", fnt(16), WHITE)
    draw_centered_text(d, (margin_x + label_w / 2, disc_y + 165), "Continuous", fnt(13), WHITE)
    draw_centered_text(d, (margin_x + label_w / 2, disc_y + 200), "PM + Designer +", fnt(11), WHITE)
    draw_centered_text(d, (margin_x + label_w / 2, disc_y + 218), "Tech Lead", fnt(11, True), WHITE)

    disc_steps = [
        ("PROBLEM", "Frame the real problem — not the proposed solution"),
        ("USERS", "Personas · pain points · context"),
        ("HYPOTHESIS", "Value · feasibility · viability assumptions"),
        ("VALIDATE", "Prototype · interview · usability tests"),
        ("DECIDE", "Go / pivot / kill"),
    ]
    step_w = (W - margin_x - label_w - margin_x - 20 - (len(disc_steps) - 1) * 14) // len(disc_steps)
    for i, (lab, sub) in enumerate(disc_steps):
        x = margin_x + label_w + 20 + i * (step_w + 14)
        y = disc_y + 30
        h = track_h - 60
        d.rounded_rectangle([x, y, x + step_w, y + h], radius=10, fill=WHITE, outline=AMBER, width=2)
        d.rounded_rectangle([x, y, x + step_w, y + 36], radius=10, fill=AMBER)
        draw_centered_text(d, (x + step_w / 2, y + 18), lab, fnt(14, True), WHITE)
        wrapped = wrap_text(d, sub, fnt(12), step_w - 20)
        for li, ln in enumerate(wrapped[:5]):
            d.text((x + 12, y + 60 + li * 22), ln, fill=DARK_TXT, font=fnt(12))
        # arrow
        if i < len(disc_steps) - 1:
            ax = x + step_w + 7
            ay = y + h / 2
            d.polygon([(ax - 5, ay - 7), (ax + 5, ay), (ax - 5, ay + 7)], fill=AMBER_DK)

    # DELIVERY track (bottom)
    del_y = disc_y + track_h + 80
    d.rounded_rectangle([margin_x, del_y, W - margin_x, del_y + track_h], radius=14,
                        fill=LIGHT_BG, outline=NAVY_DEEP, width=3)
    d.rounded_rectangle([margin_x, del_y, margin_x + label_w, del_y + track_h], radius=14,
                        fill=NAVY_DEEP)
    draw_centered_text(d, (margin_x + label_w / 2, del_y + 80), "DELIVERY", fnt(28, True), WHITE)
    draw_centered_text(d, (margin_x + label_w / 2, del_y + 120), "How & When", fnt(16), WHITE)
    draw_centered_text(d, (margin_x + label_w / 2, del_y + 165), "Sprint / Iteration", fnt(13), WHITE)
    draw_centered_text(d, (margin_x + label_w / 2, del_y + 200), "Squad: PO + Devs", fnt(11), WHITE)
    draw_centered_text(d, (margin_x + label_w / 2, del_y + 218), "+ QA + Tech Lead", fnt(11, True), WHITE)

    del_steps = [
        ("REFINE", "DoR · acceptance criteria · estimation"),
        ("BUILD", "Code · unit tests · peer review"),
        ("INTEGRATE", "Integration tests · cross-product checks"),
        ("VALIDATE", "UAT · functional · NFR (perf · security)"),
        ("RELEASE", "DoD · deploy · hypercare"),
    ]
    for i, (lab, sub) in enumerate(del_steps):
        x = margin_x + label_w + 20 + i * (step_w + 14)
        y = del_y + 30
        h = track_h - 60
        d.rounded_rectangle([x, y, x + step_w, y + h], radius=10, fill=WHITE, outline=NAVY, width=2)
        d.rounded_rectangle([x, y, x + step_w, y + 36], radius=10, fill=NAVY)
        draw_centered_text(d, (x + step_w / 2, y + 18), lab, fnt(14, True), WHITE)
        wrapped = wrap_text(d, sub, fnt(12), step_w - 20)
        for li, ln in enumerate(wrapped[:5]):
            d.text((x + 12, y + 60 + li * 22), ln, fill=DARK_TXT, font=fnt(12))
        if i < len(del_steps) - 1:
            ax = x + step_w + 7
            ay = y + h / 2
            d.polygon([(ax - 5, ay - 7), (ax + 5, ay), (ax - 5, ay + 7)], fill=NAVY_DEEP)

    # Connecting note (between tracks)
    note_y = disc_y + track_h + 20
    d.rectangle([margin_x + 230, note_y, W - margin_x - 30, note_y + 36], fill=GREEN_DK)
    draw_centered_text(d, (W / 2, note_y + 18),
                       "↑ ↓  Discovery feeds Delivery the validated WHAT;  Delivery feeds Discovery real-world data  ↑ ↓",
                       fnt(13, True), WHITE)

    img.save(OUT / "04_Dual_Track.png", "PNG", dpi=(150, 150))
    print("✓ 04_Dual_Track.png")


# =========================================================
# 05 — ROLES STACK (Heather PM area + product team layers)
# =========================================================
def visual_05_roles_stack():
    img = Image.new("RGB", (W, H), WHITE)
    d = ImageDraw.Draw(img)

    d.rectangle([0, 0, W, 90], fill=NAVY)
    d.text((40, 22), "ROLES STACK — Product Management Area + Stable Product Teams",
           font=fnt(30, True), fill=WHITE)
    d.text((40, 62), "Heather + 2 hires lead a NEW Product Management layer · stable squads execute",
           font=fnt(16), fill=(199, 211, 228))

    # Top layer — Heather PM area (NEW)
    top_y = 130
    layer_h = 130
    d.rounded_rectangle([60, top_y, W - 60, top_y + layer_h], radius=14, fill=AMBER, outline=AMBER_DK, width=3)
    draw_centered_text(d, (W / 2, top_y + 32),
                       "PRODUCT MANAGEMENT AREA  (NEW · TO-BE · led by Heather + 2 hires)",
                       fnt(22, True), WHITE)
    draw_centered_text(d, (W / 2, top_y + 64),
                       "Owns: portfolio strategy · product definition · cross-product roadmap · value framework · OKRs",
                       fnt(14), WHITE)
    draw_centered_text(d, (W / 2, top_y + 90),
                       "Cadence: monthly portfolio review · quarterly BRP planning · alignment with Vitor + Carlos + Kumar",
                       fnt(13, True), WHITE)
    draw_centered_text(d, (W / 2, top_y + 112),
                       "Reports across both branches (IT Apps + CS Apps)",
                       fnt(12), WHITE)

    # Middle layer — Product teams (per major product)
    mid_y = top_y + layer_h + 30
    mid_h = 380
    d.text((60, mid_y), "STABLE PRODUCT TEAMS (one per product, not per app)",
           font=fnt(16, True), fill=DARK_TXT)

    products = [
        ("Meter-to-Cash", "Customer Services", GREEN_DK),
        ("Field Operations", "GIS · OSG · E&G", TEAL),
        ("Corporate Functions", "Finance · HR · Legal", NAVY_DEEP),
        ("Iberdrola Global SAP", "FI · PO · Budget", AMBER_DK),
    ]
    p_y = mid_y + 35
    p_w = (W - 120 - (len(products) - 1) * 20) // len(products)
    p_h = 320
    for i, (name, scope, color) in enumerate(products):
        x = 60 + i * (p_w + 20)
        d.rounded_rectangle([x, p_y, x + p_w, p_y + p_h], radius=12, fill=LIGHT_BG, outline=color, width=2)
        # Product header
        d.rounded_rectangle([x, p_y, x + p_w, p_y + 70], radius=12, fill=color)
        draw_centered_text(d, (x + p_w / 2, p_y + 24), name, fnt(16, True), WHITE)
        draw_centered_text(d, (x + p_w / 2, p_y + 50), scope, fnt(12), WHITE)

        # Roles in product team
        roles = [
            ("PRODUCT MANAGER", "Owns outcome · value · roadmap"),
            ("PRODUCT OWNER", "Owns backlog · acceptance criteria"),
            ("TECH LEAD", "Owns architecture · NFR · tech debt"),
            ("SQUAD", "Devs · QA · UX · BA"),
            ("VENDOR PARTNER", "Accenture / SWF / specific vendor"),
        ]
        r_y = p_y + 90
        for j, (lab, sub) in enumerate(roles):
            y = r_y + j * 45
            d.rounded_rectangle([x + 12, y, x + p_w - 12, y + 38], radius=6, fill=WHITE, outline=BORDER, width=1)
            d.text((x + 22, y + 6), lab, font=fnt(10, True), fill=color)
            d.text((x + 22, y + 22), sub, fill=GREY_TXT, font=fnt(10))

    # Bottom layer — Shared Services (existing)
    sh_y = mid_y + mid_h + 50
    sh_h = 130
    d.rounded_rectangle([60, sh_y, W - 60, sh_y + sh_h], radius=14, fill=NAVY_DARK, outline=None)
    draw_centered_text(d, (W / 2, sh_y + 30),
                       "SHARED SERVICES  (cross-product platforms · existing)",
                       fnt(20, True), WHITE)
    services = [
        "SAP Basis (Nagi Reddy) — platform foundation",
        "SWF & Tools (Raquel Soto · Donna Henry) — vendor capacity orchestration",
        "Dev Ops & App Security (Piotr Wielogorski) — pipeline + security",
        "PMO (Travis) — large-project coordination",
    ]
    sw = (W - 120 - 30) // 2
    for i, txt in enumerate(services):
        x = 80 + (i % 2) * (sw + 30)
        y = sh_y + 60 + (i // 2) * 30
        d.text((x, y), "•  " + txt, fill=WHITE, font=fnt(12))

    img.save(OUT / "05_Roles_Stack.png", "PNG", dpi=(150, 150))
    print("✓ 05_Roles_Stack.png")


if __name__ == "__main__":
    visual_01_big_picture()
    visual_02_demand_funnel()
    visual_03_product_vs_app()
    visual_04_dual_track()
    visual_05_roles_stack()
    print("\nAll 5 visuals saved to:", OUT)
