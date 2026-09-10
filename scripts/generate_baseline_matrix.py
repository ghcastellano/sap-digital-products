"""
Generate Baseline Reference Matrix — Prior Art Review PPTX (V3)
================================================================
Platform as a Product (PaaP) lens for SAP Digital Products.
Simplified slides with references + PaaP adaptation + Risks.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

from sapkb.ppt.design_system import BrandPalette, FONT
from sapkb.ppt.slide_builders import DeckBuilder, SLIDE_W, SLIDE_H

brand = BrandPalette(primary="#002060", secondary="#0070C0", accent="#19A3FC", font_family="Arial")

BASE_PATH = "/Users/gustavohenriquecastellano/Downloads/Base Apresentações Projeto Digital Product"
FILES = {
    "nWOW": "nWOW - New Ways of Working - P&S Lean Operating Model 03.10.2026 v1.2.pptx",
    "PD Playbook": "[00] Product Discovery Playbook - P&S Ways of Working 02.03.26 v.1.3 +FIGMA.pptx",
    "PM Playbook": "[00] Project Management Playbook - P&S Ways of Working 10.28.25 v.1.0-2.pptx",
    "BRP Q2": "BRP_Q2_2026_P&S Networks IT - New Ways of Working - v1.4.4 03.09.2026.pptx",
    "CoE Playbook": "[12] Digital CoE Operating Model - Playbook - 01.05.2026 V.2.5. FINAL pptx.pptx",
    "Agility Playbook": "[17] Agility Playbook - Digital Center of Excellence 06.18.2025 V.1.7.pptx",
    "NextGen Jira": "[33] NextGen 2.0 - Jira Management User Guide for CX Digitall CoE 27.06.2024.pptx",
    "CX JIRA": "[CX] JIRA Agile Management Playbook - Center of Excellence 25.03.2024 V.1.0.pdf",
    "NTTDATA": "NTTDATA-Digital Strategy - NextGen Produc Ops Org Design SEAL-SIX Model v.1.7 06.29.pptx",
    "Avangrid": "Agile Management Model _ Proposal Avangrid Architecture V.3.1 18.01.2024 3.pptx",
    "Project Book": "Project_Book_Avangrid_Template v1.0.pptx",
    "TPO FLOW": "TPO FLOW Project - JIRA Training Guide 05-13-2025.pptx",
}

ORANGE = RGBColor(0xE8, 0x6E, 0x00)
GREEN = RGBColor(0x00, 0xB0, 0x50)
RED = RGBColor(0xC5, 0x30, 0x30)
LIGHT_ORANGE = RGBColor(0xFF, 0xF3, 0xE0)
LIGHT_GREEN = RGBColor(0xE8, 0xF5, 0xE9)
LIGHT_RED = RGBColor(0xFF, 0xEB, 0xEE)


def file_uri(short_name):
    fname = FILES.get(short_name, short_name)
    return f"file:///{Path(BASE_PATH) / fname}"


def add_hyperlink(run, url):
    r = run._r
    hlinkClick = r.makeelement(qn('a:hlinkClick'), {})
    rId = run.part.relate_to(url, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)
    hlinkClick.set(qn('r:id'), rId)
    rPr = r.find(qn('a:rPr'))
    if rPr is None:
        rPr = r.makeelement(qn('a:rPr'), {})
        r.insert(0, rPr)
    rPr.append(hlinkClick)


def add_text(tf, text, size=7, color=None, bold=False, indent=0, space_after=2, is_link=False, link_url=None):
    """Helper to add a paragraph to a text frame."""
    p = tf.add_paragraph() if tf.paragraphs[0].text else tf.paragraphs[0]
    if tf.paragraphs[0].text:
        p = tf.add_paragraph()
    run = p.add_run()
    run.text = ("  " * indent) + text
    run.font.size = Pt(size)
    run.font.color.rgb = color or brand.charcoal
    run.font.bold = bold
    run.font.name = brand.font_family
    if is_link and link_url:
        run.font.color.rgb = brand.secondary
        run.font.underline = True
        try:
            add_hyperlink(run, link_url)
        except Exception:
            pass
    p.space_after = Pt(space_after)
    return p


# ═══════════════════════════════════════════
# DELIVERABLES DATA (simplified)
# ═══════════════════════════════════════════
DELIVERABLES = [
    {
        "name": "E2E Operating Model",
        "coverage": {"P&S / P&T": "●", "Digital CoE": "◕", "Infrastructure": "◔"},
        "existing": [
            {"area": "P&S / P&T", "what": "6-phase Kanban lifecycle (Discover → Close), demand funnel, Big Room Planning, phase gates",
             "refs": [("nWOW", "Slides 1, 39, 47, 49"), ("PM Playbook", "Slides 10-20"), ("BRP Q2", "Slides 1-9")]},
            {"area": "Digital CoE", "what": "E2E workflow with Discovery upstream, Strategic Initiatives Funnel, Lean Business Case, Portfolio Learn Cycle",
             "refs": [("CoE Playbook", "Slides 6, 9, 35-55")]},
            {"area": "Infrastructure", "what": "Phased deployment methodology (Pilot → Expand → Optimize), Team Topologies framework",
             "refs": [("Avangrid", "Slides 8, 10, 22, 29-32")]},
        ],
        "paap_shift": "Replace 'Product Lifecycle' with 'Platform Lifecycle': SAP builds reusable capabilities (APIs, integrations) that consuming teams pull on-demand — not a project-by-project delivery model. The lifecycle becomes: Discover Capability Gap → Engineer Reusable Service → Publish & Document → Measure Adoption.",
    },
    {
        "name": "Value & Prioritization Framework",
        "coverage": {"P&S / P&T": "◕", "Digital CoE": "◑", "Infrastructure": "○"},
        "existing": [
            {"area": "P&S / P&T", "what": "Portfolio Prioritization Calculator, weighted scoring, value-driven funnel, portfolio estimation",
             "refs": [("nWOW", "Slides 35-42"), ("PD Playbook", "Slides 82-103")]},
            {"area": "Digital CoE", "what": "PBI formula (BV+UX)/E, backlog prioritization matrix, HEVA cycle, sprint estimation",
             "refs": [("CoE Playbook", "Slide 84"), ("Agility Playbook", "Slides 50-83")]},
        ],
        "paap_shift": "Eliminate the intake scoring funnel. SAP doesn't score business value — it measures platform health: API adoption rate, integration lead time, self-service ratio. Prioritization becomes: which capability, if standardized, would unblock the most consuming teams? Think 'build the Customer Master API once' vs. 'score each OPCO's request individually'.",
    },
    {
        "name": "Demand Typology & Intake",
        "coverage": {"P&S / P&T": "●", "Digital CoE": "○", "Infrastructure": "○"},
        "existing": [
            {"area": "P&S / P&T", "what": "Demand sources mapped, Portfolio Kanban, classification (Strategic/Operational/Regulatory), triage with S/M/L/XL sizing",
             "refs": [("nWOW", "Slides 35-42"), ("BRP Q2", "Slides 1-9")]},
        ],
        "paap_shift": "Replace the intake funnel with a Self-Service Portal + API Catalog. Standard integrations (read Customer, post Order) require zero intake — teams consume from the catalog. Only net-new capabilities or SAP customizations go through a lightweight request process. Target: 75% of integrations are self-service.",
    },
    {
        "name": "Roles & RACI",
        "coverage": {"P&S / P&T": "●", "Digital CoE": "◔", "Infrastructure": "◔"},
        "existing": [
            {"area": "P&S / P&T", "what": "Full RACI per lifecycle phase, crew roles (PO, SM, Dev), stakeholder mapping, resource planning",
             "refs": [("PM Playbook", "Slides 105-129"), ("nWOW", "Slides 13-40")]},
            {"area": "Digital CoE", "what": "Roles implied through Scrum adoption",
             "refs": [("Agility Playbook", "Slides 18-25")]},
            {"area": "Infrastructure", "what": "Team Topologies: Enabling, Stream-Aligned, Platform, Component team types and interaction modes",
             "refs": [("Avangrid", "Slides 29-32")]},
        ],
        "paap_shift": "SAP CoE becomes a Platform Team (Team Topologies). Key role shift: the SAP 'Product Manager' becomes a 'Developer Advocate' focused on DevEx, documentation, and evangelization — not backlog grooming for external requests. RACI simplifies: SAP owns the platform; consuming teams own their integration.",
    },
    {
        "name": "Tool Ecosystem & Jira Config",
        "coverage": {"P&S / P&T": "◕", "Digital CoE": "●", "Infrastructure": "◑"},
        "existing": [
            {"area": "P&S / P&T", "what": "Jira workflows, board configs, KPI dashboards from Jira data",
             "refs": [("PM Playbook", "Slides 88-104")]},
            {"area": "Digital CoE", "what": "Complete JIRA Agile Playbook (41 pages), board design, portfolio view, burndown, NextGen config",
             "refs": [("CX JIRA", "Full doc (41 pages)"), ("NextGen Jira", "Slides 1-22")]},
            {"area": "Infrastructure", "what": "JIRA training guide, Kanban boards, filters",
             "refs": [("TPO FLOW", "Slides 1-14")]},
        ],
        "paap_shift": "Jira tracks platform capabilities and API health, not project tickets. Add: Internal Developer Portal (API docs, onboarding guides, sandbox environments). The Jira board becomes a 'Platform Backlog' of reusable services to build, not a queue of requests to fulfill.",
    },
    {
        "name": "Portfolio & PPMNow Concept",
        "coverage": {"P&S / P&T": "◕", "Digital CoE": "◕", "Infrastructure": "◑"},
        "existing": [
            {"area": "P&S / P&T", "what": "Portfolio Kanban, demand funnel, Big Room Planning, estimation framework, OKR alignment",
             "refs": [("nWOW", "Slides 35-87")]},
            {"area": "Digital CoE", "what": "Portfolio Learn Cycle (quarterly), strategic roadmap, portfolio KPIs",
             "refs": [("CoE Playbook", "Slides 23-98")]},
            {"area": "Infrastructure", "what": "Program-level portfolio, investment model per unit served",
             "refs": [("Avangrid", "Slides 15-19, 51-52")]},
        ],
        "paap_shift": "Portfolio view becomes a 'Platform Capability Map' — which SAP services exist, their maturity, adoption rate, and SLA compliance. PPMNow shows capability coverage across OPCOs, not project status. The roadmap is driven by 'which capabilities to build next' based on cross-OPCO demand patterns.",
    },
    {
        "name": "Baseline, Demand Map & Portfolio Landscape",
        "coverage": {"P&S / P&T": "◔", "Digital CoE": "○", "Infrastructure": "◔"},
        "existing": [
            {"area": "P&S / P&T", "what": "Baseline methodology (interviews, assessment), demand flow visualization, project plan format",
             "refs": [("nWOW", "Slides 1-12, 35-42"), ("PM Playbook", "Slides 131-139")]},
            {"area": "Infrastructure", "what": "Project Book template for structured project documentation",
             "refs": [("Project Book", "Full template")]},
        ],
        "paap_shift": "Build new — but frame it as a Platform Assessment: Which SAP capabilities are most consumed? Where are the integration pain points? What is the current 'Time to Hello World' for a new team trying to use SAP? Map capabilities, not projects.",
    },
    {
        "name": "Pilot Execution & Report",
        "coverage": {"P&S / P&T": "◔", "Digital CoE": "◔", "Infrastructure": "◑"},
        "existing": [
            {"area": "P&S / P&T", "what": "Feasibility assessment matrix, structured validation with go/no-go gate",
             "refs": [("PD Playbook", "Slide 68")]},
            {"area": "Digital CoE", "what": "Maturity assessment approach, MVP pilot with maturity radar",
             "refs": [("NTTDATA", "Slides 9-10")]},
            {"area": "Infrastructure", "what": "Pilot → Expansion → Optimization methodology, rollout governance",
             "refs": [("Avangrid", "Slides 34, 43")]},
        ],
        "paap_shift": "The BRS pilot validates the platform model, not just a process. Success = BRS team can consume a standardized SAP API without filing a request. Measure: Time to Hello World, self-service completion rate, zero-touch integration count.",
    },
    {
        "name": "Ways of Working Playbook",
        "coverage": {"P&S / P&T": "●", "Digital CoE": "◕", "Infrastructure": "○"},
        "existing": [
            {"area": "P&S / P&T", "what": "Product Discovery Playbook (114 slides), Project Management Playbook (150 slides) — templates, governance, RACI, KPIs",
             "refs": [("PD Playbook", "Full doc (114 slides)"), ("PM Playbook", "Full doc (150 slides)")]},
            {"area": "Digital CoE", "what": "Agility Playbook (111 slides) — Scrum/Kanban reference, sprint ceremonies, backlog practices",
             "refs": [("Agility Playbook", "Full doc (111 slides)")]},
        ],
        "paap_shift": "The playbook becomes a 'Platform Engineering Handbook': how to publish an API, how to document a service contract, how to onboard a new consuming team. Reuse P&S structure, but the content is about platform operations — not project delivery.",
    },
    {
        "name": "KPI & Reporting Framework",
        "coverage": {"P&S / P&T": "●", "Digital CoE": "◕", "Infrastructure": "○"},
        "existing": [
            {"area": "P&S / P&T", "what": "Multi-level metrics ecosystem (Portfolio → Program → Team), health dashboards, scorecards, Lead Time, Throughput, WIP",
             "refs": [("PM Playbook", "Slides 138-150"), ("nWOW", "Slides 83-97")]},
            {"area": "Digital CoE", "what": "KPIs per governance level, portfolio health metrics, sprint velocity",
             "refs": [("CoE Playbook", "Slides 84-107")]},
        ],
        "paap_shift": "Replace project KPIs with platform KPIs: Time to Hello World (45d→2d), Self-Service Ratio (0%→75%), API Reliability (SLA compliance), Integration Lead Time, Catalog Coverage (% of SAP capabilities exposed as self-service). Reuse P&S dashboard structure, replace the metrics.",
    },
]

COV_BALLS = {"●": 4, "◕": 3, "◑": 2, "◔": 1, "○": 0}

# ═══════════════════════════════════════════
# BUILD DECK
# ═══════════════════════════════════════════
deck = DeckBuilder(brand)

# ── 1. Cover ──
deck.add_cover(
    title="Prior Art Review & Platform Strategy",
    subtitle="SAP Digital Products — Wave 1\nFrom Demand Governance to Platform as a Product (PaaP)",
    date="April 2026  |  Step 1: Review & Baseline",
)

# ── 2. Section: The Shift ──
deck.add_section_divider("The Strategic Shift", number="1")

# ═══════════════════════════════════════════
# 3. VISUAL DIAGRAM: Bottleneck vs PaaP (replicating the Gemini visual)
# ═══════════════════════════════════════════
slide = deck._blank()
deck._action_title(
    slide,
    "Applying a traditional Product Lifecycle to a foundational component team creates structural bottlenecks — the modern approach is Platform as a Product (PaaP)",
    "Managing SAP CoE through a PaaP lens shifts focus from intake governance to enablement, self-service, and reducing friction for consuming teams",
)

DARK_BG = RGBColor(0x1E, 0x29, 0x3B)
DARK_CARD = RGBColor(0x2D, 0x3B, 0x50)
RED_BG = RGBColor(0x5C, 0x1A, 0x1A)
GREEN_BORDER = RGBColor(0x10, 0xB9, 0x81)
TEAL_BG = RGBColor(0x0E, 0x3D, 0x3D)
CYAN_BORDER = RGBColor(0x06, 0xB6, 0xD4)
RED_BORDER = RGBColor(0xEF, 0x44, 0x44)
EMERALD = RGBColor(0x10, 0xB9, 0x81)
GRAY_TEXT = RGBColor(0x94, 0xA3, 0xB8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

def _make_box(slide, x, y, w, h, fill, border_color=None, border_width=1.5):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if border_color:
        s.line.color.rgb = border_color
        s.line.width = Pt(border_width)
    else:
        s.line.fill.background()
    return s

def _center_text(shape, text, size=10, color=WHITE, bold=False, sub=None):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = brand.font_family
    p.alignment = PP_ALIGN.CENTER
    if sub:
        p2 = tf.add_paragraph()
        p2.text = sub
        p2.font.size = Pt(size - 2)
        p2.font.color.rgb = GRAY_TEXT
        p2.font.name = brand.font_family
        p2.alignment = PP_ALIGN.CENTER

def _arrow(slide, x, y, text="↓", color=GRAY_TEXT):
    a = slide.shapes.add_textbox(x, y, Inches(0.5), Inches(0.35))
    p = a.text_frame.paragraphs[0]
    p.text = text
    p.font.size = Pt(18)
    p.font.color.rgb = color
    p.font.name = brand.font_family
    p.alignment = PP_ALIGN.CENTER

# ── LEFT PANEL: Proposed Model (Bottleneck) ──
panel_y = Inches(1.45)
lx = Inches(0.8)
lw = Inches(5.5)

# Panel background
_make_box(slide, lx, panel_y, lw, Inches(5.1), DARK_BG, RGBColor(0x33, 0x41, 0x55))

# Title
ttl_left = slide.shapes.add_textbox(lx, panel_y + Inches(0.15), lw, Inches(0.4))
p = ttl_left.text_frame.paragraphs[0]
p.text = "Proposed Model (Bottleneck)"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = RED_BORDER
p.font.name = brand.font_family
p.alignment = PP_ALIGN.CENTER
p.font.italic = True

# Box 1: Digital Products / BRS OPCO
b1 = _make_box(slide, lx + Inches(0.6), panel_y + Inches(0.7), Inches(4.3), Inches(0.45), DARK_CARD)
_center_text(b1, "Digital Products / BRS OPCO", 10, GRAY_TEXT)

# Arrow
_arrow(slide, lx + Inches(2.5), panel_y + Inches(1.2))

# Box 2: Intake Funnel (RED highlight)
b2 = _make_box(slide, lx + Inches(0.3), panel_y + Inches(1.6), Inches(4.9), Inches(0.9), RED_BG, RGBColor(0xEF, 0x44, 0x44))
_center_text(b2, "Intake & Prioritization Funnel", 12, RGBColor(0xFC, 0xA5, 0xA5), True,
             "(ROI Scoring, Jira, PPMNow Governance)")

# Arrow
_arrow(slide, lx + Inches(2.5), panel_y + Inches(2.6))

# Box 3: SAP CoE
b3 = _make_box(slide, lx + Inches(0.6), panel_y + Inches(3.0), Inches(4.3), Inches(0.45), DARK_CARD)
_center_text(b3, "SAP CoE (Slow Execution)", 10, GRAY_TEXT)

# Problems list below
prob = slide.shapes.add_textbox(lx + Inches(0.3), panel_y + Inches(3.65), Inches(4.9), Inches(1.3))
tf = prob.text_frame
tf.word_wrap = True
for j, item in enumerate([
    "• More time in prioritization rituals than enabling",
    "• Each OPCO request goes through full funnel",
    "• SAP team is a ticket taker, not a platform",
    "• Scales linearly with headcount, not capabilities",
]):
    p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
    p.text = item
    p.font.size = Pt(8)
    p.font.color.rgb = RGBColor(0xFC, 0xA5, 0xA5)
    p.font.name = brand.font_family
    p.space_after = Pt(2)

# ── RIGHT PANEL: Platform as a Product (PaaP) ──
rx = Inches(6.8)
rw = Inches(5.5)

# Panel background with green glow border
_make_box(slide, rx, panel_y, rw, Inches(5.1), DARK_BG, GREEN_BORDER)

# "Recommended" badge
badge = _make_box(slide, rx + Inches(3.9), panel_y, Inches(1.5), Inches(0.3), EMERALD)
_center_text(badge, "Recommended", 8, DARK_BG, True)

# Title
ttl_right = slide.shapes.add_textbox(rx, panel_y + Inches(0.15), rw, Inches(0.4))
p = ttl_right.text_frame.paragraphs[0]
p.text = "Platform as a Product (PaaP)"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = EMERALD
p.font.name = brand.font_family
p.alignment = PP_ALIGN.CENTER

# 3 consumer boxes side by side
consumers = ["BRS OPCO Pilot", "Digital Products", "Global PMO"]
cw = Inches(1.55)
for ci, c_label in enumerate(consumers):
    cx = rx + Inches(0.25) + ci * Inches(1.7)
    cb = _make_box(slide, cx, panel_y + Inches(0.7), cw, Inches(0.38), DARK_CARD)
    _center_text(cb, c_label, 8, GRAY_TEXT)

# 3 arrows
for ci in range(3):
    cx = rx + Inches(0.8) + ci * Inches(1.7)
    _arrow(slide, cx, panel_y + Inches(1.15))

# Platform box (GREEN highlight)
pb = _make_box(slide, rx + Inches(0.2), panel_y + Inches(1.55), Inches(5.1), Inches(0.85), TEAL_BG, GREEN_BORDER)
_center_text(pb, "APIs, SDKs & Self-Service Portal", 12, RGBColor(0xA7, 0xF3, 0xD0), True,
             "(Standardized Contracts, Docs, Automation)")

# Up arrow
_arrow(slide, rx + Inches(2.5), panel_y + Inches(2.45), "↑", GRAY_TEXT)

# SAP CoE Platform
sb = _make_box(slide, rx + Inches(0.4), panel_y + Inches(2.85), Inches(4.7), Inches(0.45),
               RGBColor(0x0C, 0x31, 0x4D), CYAN_BORDER)
_center_text(sb, "SAP CoE (Platform Evolution)", 10, RGBColor(0x67, 0xE8, 0xF9))

# Advantages list
adv = slide.shapes.add_textbox(rx + Inches(0.3), panel_y + Inches(3.5), Inches(4.9), Inches(1.3))
tf = adv.text_frame
tf.word_wrap = True
for j, item in enumerate([
    "• 75% of integrations via self-service (no intake)",
    "• Time to Hello World: 45 days → 2 days",
    "• Scales across OPCOs without adding headcount",
    "• SAP team focuses on engineering, not governance",
]):
    p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
    p.text = item
    p.font.size = Pt(8)
    p.font.color.rgb = RGBColor(0xA7, 0xF3, 0xD0)
    p.font.name = brand.font_family
    p.space_after = Pt(2)


# ═══════════════════════════════════════════
# 4. VISUAL: The Enablement Roadmap (Wave 1 Pivot)
# ═══════════════════════════════════════════
slide = deck._blank()
deck._action_title(
    slide,
    "The Enablement Roadmap (Wave 1 Pivot)",
    "Instead of dedicating the 14-week program to configuring heavy PPMNow intake flows and ROI matrices, the engagement should pivot towards encapsulating SAP complexities and establishing the foundation for a scalable, developer-friendly internal platform.",
)

STEP_COLORS = [
    (RGBColor(0xF5, 0x9E, 0x0B), RGBColor(0x78, 0x35, 0x0F)),   # Amber
    (RGBColor(0x06, 0xB6, 0xD4), RGBColor(0x16, 0x47, 0x57)),   # Cyan
    (RGBColor(0x10, 0xB9, 0x81), RGBColor(0x06, 0x4E, 0x3B)),   # Emerald
]

steps_data = [
    ("1", "Map Capabilities, Not Projects",
     'Identify which SAP data sets and transactions are most frequently requested by Digital Products and the BRS pilot. Shift the mandate from "delivering the BRS project" to "engineering the master Customer API that BRS and all future OPCOs will consume."'),
    ("2", "Define Contracts & SLAs (Platform Engineering)",
     "Establish crystal-clear interface contracts. Product Management in this context applies to API reliability, endpoint response times, backward compatibility, and treating the technology stack itself as a product consumed by internal developers."),
    ("3", "Evangelization and DevEx",
     'The SAP CoE Product Manager acts as a "Developer Advocate". Their goal is to build integration playbooks, maintain internal developer portals, and guarantee that the SAP landscape acts as an invisible, high-speed enabler rather than a procedural dependency.'),
]

for i, (num, title, desc) in enumerate(steps_data):
    accent_color, dark_accent = STEP_COLORS[i]
    y = Inches(1.55 + i * Inches(1.55).inches)

    # Card background
    card = _make_box(slide, Inches(0.8), y, Inches(11.5), Inches(1.35), DARK_BG, RGBColor(0x33, 0x41, 0x55))

    # Number circle
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                     Inches(1.1), y + Inches(0.3),
                                     Inches(0.55), Inches(0.55))
    circle.fill.solid()
    circle.fill.fore_color.rgb = dark_accent
    circle.line.color.rgb = accent_color
    circle.line.width = Pt(2)
    cp = circle.text_frame.paragraphs[0]
    cp.text = num
    cp.font.size = Pt(20)
    cp.font.bold = True
    cp.font.color.rgb = accent_color
    cp.font.name = brand.font_family
    cp.alignment = PP_ALIGN.CENTER
    circle.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    # Title
    ttl = slide.shapes.add_textbox(Inches(1.9), y + Inches(0.15), Inches(9.5), Inches(0.4))
    p = ttl.text_frame.paragraphs[0]
    p.text = title
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.font.name = brand.font_family

    # Description
    dtxt = slide.shapes.add_textbox(Inches(1.9), y + Inches(0.55), Inches(9.8), Inches(0.7))
    tf = dtxt.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = desc
    p.font.size = Pt(9)
    p.font.color.rgb = GRAY_TEXT
    p.font.name = brand.font_family


# ── 5. Section: Coverage Matrix ──
deck.add_section_divider("Existing Assets & PaaP Adaptation", number="2")

# ── 6-7. Coverage Matrix (2 pages) ──
for batch_idx in range(2):
    batch = DELIVERABLES[batch_idx * 5:(batch_idx + 1) * 5]
    if not batch:
        continue
    slide = deck._blank()
    deck._action_title(slide, f"What exists across areas and how PaaP reframes each deliverable {'(1/2)' if batch_idx == 0 else '(2/2)'}")

    headers = ["Deliverable", "P&S/P&T", "Digital CoE", "Infra", "PaaP Reframing"]
    col_widths = [Inches(1.6), Inches(0.8), Inches(0.8), Inches(0.7), Inches(7.6)]
    n_rows = len(batch) + 1
    n_cols = len(headers)

    ts = slide.shapes.add_table(n_rows, n_cols, Inches(0.8), Inches(1.3), Inches(11.5), Inches(0.9) * n_rows)
    table = ts.table

    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = brand.primary
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(8)
            p.font.bold = True
            p.font.color.rgb = brand.white
            p.font.name = brand.font_family
            p.alignment = PP_ALIGN.CENTER if i < 4 else PP_ALIGN.LEFT
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    for r_idx, d in enumerate(batch):
        bg = brand.white if r_idx % 2 == 0 else brand.light_gray

        cell = table.cell(r_idx + 1, 0)
        cell.text = d["name"]
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(7)
            p.font.bold = True
            p.font.color.rgb = brand.charcoal
            p.font.name = brand.font_family
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

        for c_idx, area in enumerate(["P&S / P&T", "Digital CoE", "Infrastructure"]):
            cell = table.cell(r_idx + 1, c_idx + 1)
            cell.text = d["coverage"].get(area, "○")
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(12)
                p.font.color.rgb = brand.primary
                p.alignment = PP_ALIGN.CENTER
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE

        cell = table.cell(r_idx + 1, 4)
        cell.text = d["paap_shift"][:180] + ("..." if len(d["paap_shift"]) > 180 else "")
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(7)
            p.font.color.rgb = brand.charcoal
            p.font.name = brand.font_family
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    for i, w in enumerate(col_widths):
        table.columns[i].width = int(w)


# ── 8-17. Deliverable detail slides (simplified) ──
deck.add_section_divider("Deliverable Details", number="3")

for idx, d in enumerate(DELIVERABLES, 1):
    slide = deck._blank()
    # Simpler title with deliverable name
    deck._action_title(slide, f"{idx}. {d['name']}")

    # Left side: Existing assets (compact)
    left_x = Inches(0.8)
    left_w = Inches(5.8)
    y_cursor = Inches(1.3)

    hdr_left = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_x, y_cursor, left_w, Inches(0.3))
    hdr_left.fill.solid()
    hdr_left.fill.fore_color.rgb = brand.primary
    hdr_left.line.fill.background()
    p = hdr_left.text_frame.paragraphs[0]
    cov_str = "  |  ".join(f"{a} {s}" for a, s in d["coverage"].items())
    p.text = f"EXISTING ASSETS    {cov_str}"
    p.font.size = Pt(8)
    p.font.bold = True
    p.font.color.rgb = brand.white
    p.font.name = brand.font_family
    hdr_left.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    content_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left_x, y_cursor + Inches(0.3),
                                          left_w, Inches(5.3))
    content_box.fill.solid()
    content_box.fill.fore_color.rgb = brand.light_gray
    content_box.line.color.rgb = brand.mid_gray
    content_box.line.width = Pt(0.5)

    txt = slide.shapes.add_textbox(left_x + Inches(0.1), y_cursor + Inches(0.4),
                                    left_w - Inches(0.2), Inches(5.0))
    tf = txt.text_frame
    tf.word_wrap = True

    first_para = True
    for ex in d["existing"]:
        if not first_para:
            p = tf.add_paragraph()
            p.text = ""
            p.space_before = Pt(4)

        # Area name
        p = tf.add_paragraph() if not first_para else tf.paragraphs[0]
        first_para = False
        p.text = ex["area"]
        p.font.size = Pt(8)
        p.font.bold = True
        p.font.color.rgb = brand.secondary
        p.font.name = brand.font_family
        p.space_after = Pt(1)

        # What exists
        p = tf.add_paragraph()
        p.text = ex["what"]
        p.font.size = Pt(7)
        p.font.color.rgb = brand.charcoal
        p.font.name = brand.font_family
        p.space_after = Pt(2)

        # References with full filename + link + slides
        for ref_name, ref_slides in ex["refs"]:
            full_name = FILES.get(ref_name, ref_name)
            p = tf.add_paragraph()
            # Arrow prefix
            run_arrow = p.add_run()
            run_arrow.text = "  → "
            run_arrow.font.size = Pt(6.5)
            run_arrow.font.color.rgb = brand.dark_gray
            run_arrow.font.name = brand.font_family
            # Full filename as clickable link
            run_file = p.add_run()
            run_file.text = full_name
            run_file.font.size = Pt(6.5)
            run_file.font.color.rgb = brand.secondary
            run_file.font.name = brand.font_family
            run_file.font.underline = True
            try:
                add_hyperlink(run_file, file_uri(ref_name))
            except Exception:
                pass
            # Slides reference
            run_slides = p.add_run()
            run_slides.text = f"  |  {ref_slides}"
            run_slides.font.size = Pt(6.5)
            run_slides.font.bold = True
            run_slides.font.color.rgb = brand.charcoal
            run_slides.font.name = brand.font_family
            p.space_after = Pt(1)

    # Right side: PaaP Shift
    right_x = Inches(6.8)
    right_w = Inches(5.5)

    hdr_right = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, right_x, y_cursor, right_w, Inches(0.3))
    hdr_right.fill.solid()
    hdr_right.fill.fore_color.rgb = GREEN
    hdr_right.line.fill.background()
    p = hdr_right.text_frame.paragraphs[0]
    p.text = "PLATFORM AS A PRODUCT — HOW THIS CHANGES"
    p.font.size = Pt(8)
    p.font.bold = True
    p.font.color.rgb = brand.white
    p.font.name = brand.font_family
    hdr_right.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    paap_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, right_x, y_cursor + Inches(0.3),
                                       right_w, Inches(5.3))
    paap_box.fill.solid()
    paap_box.fill.fore_color.rgb = LIGHT_GREEN
    paap_box.line.color.rgb = GREEN
    paap_box.line.width = Pt(1)

    ptxt = slide.shapes.add_textbox(right_x + Inches(0.15), y_cursor + Inches(0.45),
                                     right_w - Inches(0.3), Inches(4.8))
    tf = ptxt.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = d["paap_shift"]
    p.font.size = Pt(9)
    p.font.color.rgb = brand.charcoal
    p.font.name = brand.font_family


# ── 18. Section: Risks ──
deck.add_section_divider("Risks & Mitigations", number="4")

# ── 19. Risks slide ──
slide = deck._blank()
deck._action_title(
    slide,
    "The PaaP approach carries real risks — organizational readiness, scope creep toward platform engineering, and the gap between aspiration and SAP's current maturity",
)

risks = [
    ("Organizational Maturity Gap", RED,
     "The PaaP model assumes a level of engineering culture (API-first, DevOps, self-service) that may not exist today in the SAP CoE. If the team lacks API engineering skills, building a platform becomes aspirational rather than practical.",
     "Start with documentation and contracts (low-tech PaaP) before investing in API automation. Validate skills in the baseline assessment."),
    ("Scope Creep into Platform Engineering", ORANGE,
     "There is a risk of the 14-week Wave 1 turning into a platform engineering initiative (building APIs, portals, SDKs) instead of the operating model definition it was scoped for. The $187,600 budget was not sized for engineering work.",
     "Keep Wave 1 focused on designing the PaaP operating model and contracts. Actual API engineering belongs to Wave 2+ with a separate budget."),
    ("Resistance from Governance Stakeholders", ORANGE,
     "Eliminating the intake funnel removes control points that PMO and management currently rely on. There may be pushback from stakeholders who equate governance with value, especially if they are used to approval-based cultures.",
     "Frame PaaP as better governance, not less governance. SLA contracts and adoption metrics provide more transparency than intake funnels."),
    ("Self-Service Adoption Risk", brand.secondary,
     "Building self-service capabilities that nobody uses. Consuming teams may still prefer to 'throw requests over the wall' rather than learn new APIs and documentation — especially under deadline pressure.",
     "Measure adoption from Day 1. If self-service ratio doesn't improve by Week 8 of the pilot, revert to a hybrid model with lightweight intake for complex requests."),
    ("BRS Pilot May Not Validate Platform Model", brand.secondary,
     "A single OPCO pilot (BRS) may not generate enough diversity of use cases to validate platform capabilities. If BRS only needs 2-3 integrations, the platform investment may seem disproportionate.",
     "Design the pilot to test both BRS-specific and cross-OPCO use cases. Include at least one capability that serves multiple areas to validate the 'build once, use many' premise."),
]

col_w = Inches(5.55)
col_gap = Inches(0.4)

for i, (title, color, desc, mitigation) in enumerate(risks):
    col = i % 2
    row = i // 2
    x = Inches(0.8) + col * (col_w + col_gap)
    y = Inches(1.35) + row * Inches(1.8)
    w = col_w
    h = Inches(1.65)

    if i == 4:  # Last one centered
        x = Inches(0.8)
        w = Inches(11.5)
        h = Inches(1.3)

    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    box.fill.solid()
    box.fill.fore_color.rgb = brand.light_gray
    box.line.color.rgb = color
    box.line.width = Pt(1.5)

    # Risk title bar
    tbar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, Inches(0.28))
    tbar.fill.solid()
    tbar.fill.fore_color.rgb = color
    tbar.line.fill.background()
    p = tbar.text_frame.paragraphs[0]
    p.text = f"RISK: {title}"
    p.font.size = Pt(8)
    p.font.bold = True
    p.font.color.rgb = brand.white
    p.font.name = brand.font_family
    tbar.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    txt = slide.shapes.add_textbox(x + Inches(0.1), y + Inches(0.35), w - Inches(0.2), h - Inches(0.4))
    tf = txt.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = desc
    p.font.size = Pt(7)
    p.font.color.rgb = brand.charcoal
    p.font.name = brand.font_family
    p.space_after = Pt(3)

    p = tf.add_paragraph()
    p.text = f"Mitigation: {mitigation}"
    p.font.size = Pt(7)
    p.font.bold = True
    p.font.color.rgb = GREEN
    p.font.name = brand.font_family


# ── 20. Section: Conclusions ──
deck.add_section_divider("Conclusions", number="5")

# ── 21. KPI targets ──
deck.add_kpi_slide(
    title="Platform success is measured by friction reduction, not project delivery — these are the metrics that matter for an enabling team",
    kpis=[
        ("45d → 2d", "Time to Hello World", "First SAP API consumption"),
        ("0% → 75%", "Self-Service Ratio", "Integrations without intake"),
        ("10", "Existing Assets", "Deliverables with prior art"),
        ("823", "Knowledge Chunks", "Indexed & searchable in KB"),
    ],
)

# ── 22. Recommendations ──
deck.add_recommendation_slide(
    title="Three actions to transform SAP CoE from ticket taker to enabling platform while reusing 70% of existing assets",
    recommendations=[
        ("1", "Design the Platform Operating Model",
         "Use P&S/P&T's lifecycle and playbook structure as the foundation — reframe every element through the PaaP lens. Replace intake funnels with API catalogs, replace ROI scoring with capability gap analysis, replace project governance with SLA contracts. This is Wave 1's core deliverable.",
         ),
        ("2", "Pilot with BRS as First Consumer",
         "The BRS pilot validates the platform model: can a consuming team use a standardized SAP capability without filing a request? Measure Time to Hello World and self-service ratio. If it works for BRS, it works for all OPCOs.",
         ),
        ("3", "Build the Capability Map, Not the Governance Layer",
         "Instead of spending Wave 1 on PPMNow intake flows and prioritization matrices, invest in mapping SAP capabilities, defining service contracts, and creating the integration playbook. The governance comes from transparency (adoption metrics, SLA dashboards), not from approval workflows.",
         ),
    ],
)

# ── 23. Closing ──
deck.add_closing(
    title="Prior Art Review & Platform Strategy",
    subtitle="SAP Digital Products — Wave 1  |  Platform as a Product",
)

# ── Save ──
output_path = Path(__file__).parent.parent / "data" / "generated" / "baseline_reference_matrix.pptx"
output_path.parent.mkdir(parents=True, exist_ok=True)
n_slides = deck.save(str(output_path))
print(f"Generated {n_slides} slides → {output_path}")
