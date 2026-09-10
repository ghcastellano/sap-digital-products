"""
Stage 01 — DEMAND MANAGEMENT (titled cells)
Big board with named sub-fields per dimension, ready for Miro fill-in.
"""
from PIL import Image, ImageDraw, ImageFont
from pathlib import Path

OUT = Path("/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Miro_Visuals")

NAVY      = (44, 90, 136)
NAVY_DARK = (21, 42, 71)
TEAL      = (43, 107, 119)
GREEN     = (76, 140, 63)
AMBER     = (232, 139, 47)
RED       = (182, 59, 31)
LIGHT_BG  = (248, 249, 251)
SLOT_BG   = (255, 255, 255)
GREY_TXT  = (75, 85, 99)
TITLE_TXT = (33, 42, 56)
HINT_TXT  = (140, 150, 165)
WHITE     = (255, 255, 255)
BORDER    = (210, 216, 224)
SLOT_BORDER = (190, 200, 215)

# Big canvas
W, H = 3600, 5400


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


# Each dimension has 5 titled sub-fields (proposed by senior-consultant analysis)
DIMS = [
    ("CAPABILITIES",  "What the org must be able to do",  GREEN, [
        ("Single intake channel",        "Per product / value stream — kill parallel channels"),
        ("Definition-of-Ready (DoR)",    "Mandatory fields before scoring"),
        ("Demand categorization",        "Type · Size · Complexity · Value driver"),
        ("Triage & routing",             "Accept · Send-back · Reject — weekly"),
        ("Backlog visibility",           "Real-time, product-aligned, single source"),
    ]),
    ("ROLES",         "Who owns and operates this",       TEAL, [
        ("Demand Manager",               "Gatekeeper of intake quality (audits scoring)"),
        ("Product Manager / Owner",      "Owns scoring + product backlog priority"),
        ("Business Sponsor",             "Provides BV input + escalation authority"),
        ("Triage Committee",             "DM + 2 PMs + 1 Architect — weekly"),
        ("Portfolio Manager",            "Capacity allocation across products — monthly"),
    ]),
    ("TOOLS",         "Systems and platforms used",       AMBER, [
        ("ServiceNow (intake form)",     "Single front door — one form per product"),
        ("Jira (product backlog)",       "System of record post-DoR pass"),
        ("PPMNow (portfolio)",           "Capacity, sequencing, cross-product view"),
        ("Confluence (DoR + criteria)",  "Living standard — auditable"),
        ("Demand dashboard",             "Throughput · COS mix · time-to-triage"),
    ]),
    ("GOVERNANCE",    "Decisions, forums, controls",      NAVY, [
        ("Definition-of-Ready gate",     "DoR met → eligible; not met → send back"),
        ("Classes of Service",           "Standard · Fixed-Date · Expedite · Intangible"),
        ("Gate 1 — Triage (weekly)",     "30-min time-box; outcomes logged in Jira"),
        ("Gate 2 — Portfolio (monthly)", "Heads + Sponsors — sequencing + trade-offs"),
        ("Expedite cap policy",          "Regulatory ≤ 20% of throughput; rest protected"),
    ]),
    ("PERFORMANCE",   "How success is measured",          RED, [
        ("% demand via single channel",  "Target ≥ 90% by month 6 (kill legacy)"),
        ("Time-to-triage (median)",      "From submission → Gate 1 decision; target ≤ 5 days"),
        ("Demand mix by COS",            "Standard / Fixed-Date / Expedite / Intangible — monthly"),
        ("DoR pass-rate at first try",   "Proxy for solicitor maturity; target ≥ 70%"),
        ("Demand-to-delivery conversion","% scored items that ship in ≤ 2 quarters"),
    ]),
]


img = Image.new("RGB", (W, H), WHITE)
d = ImageDraw.Draw(img)

# ── Header ─────────────────────────────────────────────────────────────────
d.rectangle([0, 0, W, 320], fill=NAVY_DARK)
d.text((120, 70), "STAGE 01", font=fnt(54, True), fill=AMBER)
d.text((120, 145), "DEMAND MANAGEMENT", font=fnt(108, True), fill=WHITE)
d.text((120, 260), "Product-centric model — single front door, value-based prioritization, two-gate governance",
       font=fnt(34), fill=(199, 211, 228))

# ── Layout ─────────────────────────────────────────────────────────────────
margin_x   = 120
top_y      = 400
label_w    = 460
gap_label  = 60
gap_row    = 60
footer_h   = 130

n_rows = len(DIMS)
n_slots = 5  # sub-fields per row
gap_slot = 24
slot_title_h = 90  # banner inside each slot for the field title

available_h = H - top_y - footer_h - (n_rows - 1) * gap_row
row_h       = available_h // n_rows
slot_x0     = margin_x + label_w + gap_label
slot_area_w = W - margin_x - slot_x0
slot_w      = (slot_area_w - (n_slots - 1) * gap_slot) // n_slots

for r, (dim, dim_sub, color, slots) in enumerate(DIMS):
    y0 = top_y + r * (row_h + gap_row)

    # Dimension label (left)
    d.rounded_rectangle([margin_x, y0, margin_x + label_w, y0 + row_h],
                        radius=26, fill=color)
    centered(d, (margin_x + label_w / 2, y0 + row_h / 2 - 36),
             dim, fnt(56, True), WHITE)
    centered(d, (margin_x + label_w / 2, y0 + row_h / 2 + 38),
             dim_sub, fnt(24), (235, 240, 245))

    # 5 titled slots
    for s, (slot_title, slot_hint) in enumerate(slots):
        sx0 = slot_x0 + s * (slot_w + gap_slot)
        sx1 = sx0 + slot_w
        sy1 = y0 + row_h
        # outer rounded box (white)
        d.rounded_rectangle([sx0, y0, sx1, sy1], radius=20,
                            fill=SLOT_BG, outline=SLOT_BORDER, width=3)
        # title banner (color tinted top)
        d.rounded_rectangle([sx0, y0, sx1, y0 + slot_title_h], radius=20,
                            fill=color, outline=color, width=0)
        # mask the bottom corners of the banner
        d.rectangle([sx0, y0 + slot_title_h - 20, sx1, y0 + slot_title_h], fill=color)

        # title text (wrapped to 2 lines if needed — naive wrap by char count)
        title_font = fnt(26, True)
        if len(slot_title) > 22:
            # split on space near the middle
            mid = len(slot_title) // 2
            left = slot_title.rfind(" ", 0, mid + 4)
            right = slot_title.find(" ", mid - 4)
            cut = left if left != -1 else right
            if cut == -1:
                line1, line2 = slot_title, ""
            else:
                line1, line2 = slot_title[:cut].strip(), slot_title[cut:].strip()
            centered(d, ((sx0 + sx1) / 2, y0 + slot_title_h / 2 - 14),
                     line1, title_font, WHITE)
            centered(d, ((sx0 + sx1) / 2, y0 + slot_title_h / 2 + 18),
                     line2, title_font, WHITE)
        else:
            centered(d, ((sx0 + sx1) / 2, y0 + slot_title_h / 2),
                     slot_title, title_font, WHITE)

        # hint line below banner
        hint_y = y0 + slot_title_h + 22
        # wrap hint to 2 lines (naive)
        max_chars = max(1, slot_w // 13)
        if len(slot_hint) > max_chars:
            words = slot_hint.split()
            line1, line2 = "", ""
            for w_ in words:
                if len(line1) + len(w_) + 1 <= max_chars:
                    line1 = (line1 + " " + w_).strip()
                else:
                    line2 = (line2 + " " + w_).strip()
            centered(d, ((sx0 + sx1) / 2, hint_y + 14),
                     line1, fnt(20), HINT_TXT)
            centered(d, ((sx0 + sx1) / 2, hint_y + 42),
                     line2, fnt(20), HINT_TXT)
        else:
            centered(d, ((sx0 + sx1) / 2, hint_y + 14),
                     slot_hint, fnt(20), HINT_TXT)

# ── Footer ─────────────────────────────────────────────────────────────────
d.text((margin_x, H - 80),
       "Avangrid · Digital Products · Target Operating Model — Stage 01 Demand Management (titled fields)",
       font=fnt(26), fill=GREY_TXT)
d.text((W - 800, H - 80),
       "Workshop board · 1 of 4 · titled",
       font=fnt(26, True), fill=AMBER)

OUT.mkdir(parents=True, exist_ok=True)
out_file = OUT / "01_TOM_DemandManagement_TITLED.png"
img.save(out_file, "PNG", dpi=(150, 150))

# preview
ratio = 1400 / W
preview = img.resize((1400, int(H * ratio)), Image.LANCZOS)
preview.save(out_file.with_name(out_file.stem + "_preview.png"), "PNG")
print(f"✓ Saved: {out_file}")
print(f"  Resolution: {W}×{H}")
