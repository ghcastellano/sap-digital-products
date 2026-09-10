"""
Consolidate deliverable slides into readable tables.
Removes slides 6-15, replaces with 3 table slides (larger fonts, more detail).
"""
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

INPUT = Path("/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Digital Products - Prior Materials Review .pptx")
OUTPUT = INPUT  # overwrite

GREEN_ACCENT = RGBColor(0x00, 0xA4, 0x43)
BLUE_PRIMARY = RGBColor(0x00, 0x20, 0x60)
BLUE_LINK = RGBColor(0x00, 0x72, 0xBC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
CHARCOAL = RGBColor(0x2C, 0x3E, 0x50)
DARK_GRAY = RGBColor(0x7F, 0x8C, 0x8D)
MID_GRAY = RGBColor(0x95, 0xA5, 0xA6)
LIGHT_GRAY = RGBColor(0xF5, 0xF5, 0xF5)

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

# ── Deliverables with MORE DETAIL on approach ──
DELIVERABLES = [
    {
        "name": "E2E Operating\nModel",
        "cov": "● | ◕ | ◔",
        "ps_pt": {
            "what": "6-phase Kanban lifecycle (Discover → Solution → Plan → Build → Operate → Close). Demand funnel with portfolio prioritization. Big Room Planning cadence for cross-team alignment. Phase gates with entry/exit criteria.",
            "refs": [("nWOW", "Sl. 1, 39, 47, 49"), ("PM Playbook", "Sl. 10-20"), ("BRP Q2", "Sl. 1-9")]},
        "digital_coe": {
            "what": "E2E workflow with upstream Discovery phase. Strategic Initiatives Funnel with PRD (Light Requirements). Lean Business Case at intake. Feedback loops between delivery and portfolio. Quarterly Portfolio Learn Cycle.",
            "refs": [("CoE Playbook", "Sl. 6, 9, 35-55")]},
        "infra": {
            "what": "Phased deployment methodology: Pilot → Expansion → Optimization. Team Topologies framework defining Enabling, Stream-Aligned, Platform and Component team types with interaction modes.",
            "refs": [("Avangrid", "Sl. 8, 10, 22, 29-32")]},
    },
    {
        "name": "Value &\nPrioritization",
        "cov": "◕ | ◑ | ○",
        "ps_pt": {
            "what": "Portfolio Prioritization Calculator with weighted scoring model. Value-driven funnel filtering at intake. Demand classification by strategic alignment and business impact. Adjustable criteria per OPCO.",
            "refs": [("nWOW", "Sl. 35-42"), ("PD Playbook", "Sl. 82-103")]},
        "digital_coe": {
            "what": "PBI prioritization formula: (Business Value + UX) / Effort = Score. Backlog-level prioritization matrix. HEVA cycle for continuous prioritization. Sprint estimation via Planning Poker and T-Shirt sizing.",
            "refs": [("CoE Playbook", "Sl. 84"), ("Agility Playbook", "Sl. 50-83")]},
        "infra": {"what": "", "refs": []},
    },
    {
        "name": "Demand Typology\n& Intake",
        "cov": "● | ○ | ○",
        "ps_pt": {
            "what": "Demand sources fully mapped with entry channels defined. Portfolio Kanban with demand funnel visualization. Classification: Strategic / Operational / Regulatory. Triage criteria with initial sizing (S/M/L/XL). Single front door concept with standardized intake form.",
            "refs": [("nWOW", "Sl. 35-42"), ("BRP Q2", "Sl. 1-9")]},
        "digital_coe": {"what": "", "refs": []},
        "infra": {"what": "", "refs": []},
    },
    {
        "name": "Roles &\nRACI",
        "cov": "● | ◔ | ◔",
        "ps_pt": {
            "what": "Full RACI matrix per lifecycle phase. Crew roles defined: Product Owner, Scrum Master, Dev Team. Stakeholder mapping with engagement matrix. Resource planning framework with allocation model. Clear handoff points between phases.",
            "refs": [("PM Playbook", "Sl. 105-129"), ("nWOW", "Sl. 13-40")]},
        "digital_coe": {
            "what": "Roles implied through Scrum framework adoption. Sprint ceremonies define responsibilities. Agile Master role defined with hygiene checklist.",
            "refs": [("Agility Playbook", "Sl. 18-25")]},
        "infra": {
            "what": "Team Topologies framework: 4 team types (Enabling, Stream-Aligned, Platform, Component) with 3 interaction modes (Collaboration, X-as-a-Service, Facilitating).",
            "refs": [("Avangrid", "Sl. 29-32")]},
    },
    {
        "name": "Tool Ecosystem\n& Jira Config",
        "cov": "◕ | ● | ◑",
        "ps_pt": {
            "what": "Jira as primary execution tool. Workflow definitions per project type. Board configurations for Kanban and Scrum. KPI dashboards fed from Jira data. Admin roles and permission structure.",
            "refs": [("PM Playbook", "Sl. 88-104")]},
        "digital_coe": {
            "what": "Complete JIRA Agile Management Playbook (41 pages). Board design methodology with portfolio view. Custom fields, workflow states, filters. Burndown charts and dashboard setup. NextGen Jira configuration guide. Lean configuration approach.",
            "refs": [("CX JIRA", "Full doc (41 pages)"), ("NextGen Jira", "Sl. 1-22")]},
        "infra": {
            "what": "JIRA process setup for specific workflows. Kanban board training and configuration. Hands-on training guide format with step-by-step.",
            "refs": [("TPO FLOW", "Sl. 1-14")]},
    },
    {
        "name": "Portfolio &\nPPMNow Concept",
        "cov": "◕ | ◕ | ◑",
        "ps_pt": {
            "what": "Portfolio Kanban as primary governance tool. Demand funnel with prioritization at portfolio level. Big Room Planning for cross-portfolio alignment. Estimation framework for portfolio-level sizing. Quarterly planning with OKR alignment.",
            "refs": [("nWOW", "Sl. 35-87")]},
        "digital_coe": {
            "what": "Portfolio Learn Cycle with quarterly cadence. Strategic roadmap with investment buckets. Portfolio-level KPIs and health metrics. Strategic alignment check at portfolio reviews.",
            "refs": [("CoE Playbook", "Sl. 23-98")]},
        "infra": {
            "what": "Program-level portfolio management. Investment model per unit served. Deployment governance with rollout planning.",
            "refs": [("Avangrid", "Sl. 15-19, 51-52")]},
    },
    {
        "name": "Baseline &\nDemand Map",
        "cov": "◔ | ○ | ◔",
        "ps_pt": {
            "what": "Baseline methodology: interview-based current state assessment with gap analysis. Demand flow visualization format. Project plan format with timeline views.",
            "refs": [("nWOW", "Sl. 1-12, 35-42"), ("PM Playbook", "Sl. 131-139")]},
        "digital_coe": {"what": "", "refs": []},
        "infra": {
            "what": "Project Book template for structured project documentation with phase tracking (Discover → Solution → Plan → Build → Operate).",
            "refs": [("Project Book", "Full template")]},
    },
    {
        "name": "Pilot\nExecution",
        "cov": "◔ | ◔ | ◑",
        "ps_pt": {
            "what": "Feasibility assessment matrix for validation. Structured go/no-go gate with clear criteria. Shark Tank validation format for ideas.",
            "refs": [("PD Playbook", "Sl. 68")]},
        "digital_coe": {
            "what": "Maturity assessment approach with health radar. MVP pilot methodology. Results measurement framework for squad performance.",
            "refs": [("NTTDATA", "Sl. 9-10")]},
        "infra": {
            "what": "Pilot → Expansion → Optimization methodology. Agile management model pilot structure with phased rollout. Expansion strategy post-pilot with governance.",
            "refs": [("Avangrid", "Sl. 34, 43")]},
    },
    {
        "name": "Ways of Working\nPlaybook",
        "cov": "● | ◕ | ○",
        "ps_pt": {
            "what": "Product Discovery Playbook (114 slides): full discovery methodology with Design Thinking, research, ideation, validation, prototyping. Project Management Playbook (150 slides): governance, backlog building, RACI, KPIs, contracts, Kanban management.",
            "refs": [("PD Playbook", "Full doc (114 sl.)"), ("PM Playbook", "Full doc (150 sl.)")]},
        "digital_coe": {
            "what": "Agility Playbook (111 slides): back-to-basics agile guide. Scrum/Kanban consistency matrix. Sprint ceremonies and cadences. Product blueprint canvas. Backlog refinement hierarchy. Agile health radars.",
            "refs": [("Agility Playbook", "Full doc (111 sl.)")]},
        "infra": {"what": "", "refs": []},
    },
    {
        "name": "KPI &\nReporting",
        "cov": "● | ◕ | ○",
        "ps_pt": {
            "what": "Multi-level metrics ecosystem: Portfolio → Program → Team. Health dashboards with project scorecards. Continuous improvement metrics: Lead Time, Throughput, WIP, Cycle Time. Portfolio-level KPI aggregation. Automated dashboards from Jira data.",
            "refs": [("PM Playbook", "Sl. 138-150"), ("nWOW", "Sl. 83-97")]},
        "digital_coe": {
            "what": "KPIs defined per governance level. Portfolio health metrics with quarterly reviews. Sprint velocity and capacity tracking. Team-of-teams capacity measurement.",
            "refs": [("CoE Playbook", "Sl. 84-107")]},
        "infra": {"what": "", "refs": []},
    },
]


def file_uri(short_name):
    fname = FILES.get(short_name, short_name)
    full_path = str(Path(BASE_PATH) / fname)
    # macOS PowerPoint expects file://localhost/path or file:///path (exactly 3 slashes)
    # urllib.parse.quote encodes spaces and special chars properly
    from urllib.parse import quote
    return "file:///" + quote(full_path, safe="/:@")


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


def set_cell_rich(cell, area_data, font_size=Pt(8)):
    """Populate cell with approach text + clickable file references."""
    tf = cell.text_frame
    tf.word_wrap = True
    tf.clear()

    if not area_data["what"]:
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = "—"
        run.font.size = font_size
        run.font.color.rgb = MID_GRAY
        run.font.italic = True
        return

    # Approach text
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = area_data["what"]
    run.font.size = font_size
    run.font.color.rgb = CHARCOAL

    # References — full filename + slides + link
    for ref_key, ref_slides in area_data["refs"]:
        full_name = FILES.get(ref_key, ref_key)

        p = tf.add_paragraph()
        p.space_before = Pt(2)

        run_arrow = p.add_run()
        run_arrow.text = "→ "
        run_arrow.font.size = Pt(7)
        run_arrow.font.color.rgb = DARK_GRAY

        run_file = p.add_run()
        run_file.text = full_name
        run_file.font.size = Pt(7)
        run_file.font.color.rgb = BLUE_LINK
        run_file.font.underline = True
        try:
            add_hyperlink(run_file, file_uri(ref_key))
        except Exception:
            pass

        run_sl = p.add_run()
        run_sl.text = f"  |  {ref_slides}"
        run_sl.font.size = Pt(7)
        run_sl.font.bold = True
        run_sl.font.color.rgb = CHARCOAL


# ═══════════════════════════════════════════
# OPEN, REMOVE OLD SLIDES, ADD TABLE SLIDES
# ═══════════════════════════════════════════
prs = Presentation(str(INPUT))

# Remove slides 6-15 (indices 5-14)
for si in range(14, 4, -1):
    sldId = prs.slides._sldIdLst[si]
    rId = sldId.get(qn('r:id'))
    prs.slides._sldIdLst.remove(sldId)
    prs.part.drop_rel(rId)

print(f"After removal: {len(prs.slides)} slides")

# Split into 3 batches: 4, 3, 3
batches = [
    (DELIVERABLES[0:4], "Existing Assets — Operating Model, Value Framework, Demand Intake, Roles & RACI (1/3)"),
    (DELIVERABLES[4:7], "Existing Assets — Tool Ecosystem, Portfolio & PPMNow, Baseline & Demand Map (2/3)"),
    (DELIVERABLES[7:10], "Existing Assets — Pilot Execution, Playbook, KPI & Reporting (3/3)"),
]

FONT_BODY = Pt(8)
FONT_HDR = Pt(8)

for batch, btitle in batches:
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.15), Inches(11.5), Inches(0.5))
    p = title_box.text_frame.paragraphs[0]
    run = p.add_run()
    run.text = btitle
    run.font.size = Pt(13)
    run.font.color.rgb = GREEN_ACCENT

    # Legend
    legend = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.5), Inches(0.25))
    p = legend.text_frame.paragraphs[0]
    run = p.add_run()
    run.text = "● Complete  ◕ Well-defined  ◑ Partial  ◔ Referenced  ○ None          → = clickable link to source file"
    run.font.size = Pt(7)
    run.font.color.rgb = MID_GRAY

    # Table
    n_rows = len(batch) + 1
    n_cols = 4
    table_top = Inches(0.9)
    table_h = Inches(6.2)

    ts = slide.shapes.add_table(n_rows, n_cols,
                                Inches(0.4), table_top, Inches(12.5), table_h)
    table = ts.table

    col_widths = [Inches(1.4), Inches(4.2), Inches(3.9), Inches(3.0)]
    for i, w in enumerate(col_widths):
        table.columns[i].width = int(w)

    # Header
    headers = ["Deliverable", "P&S / P&T", "Digital CoE", "Infrastructure"]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = BLUE_PRIMARY
        tf = cell.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = h
        run.font.size = FONT_HDR
        run.font.bold = True
        run.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    # Rows
    for r_idx, d in enumerate(batch):
        bg = WHITE if r_idx % 2 == 0 else LIGHT_GRAY

        # Col 0: Name + coverage
        cell = table.cell(r_idx + 1, 0)
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg
        tf = cell.text_frame
        tf.word_wrap = True
        tf.clear()
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = d["name"]
        run.font.size = Pt(9)
        run.font.bold = True
        run.font.color.rgb = CHARCOAL
        p2 = tf.add_paragraph()
        run2 = p2.add_run()
        run2.text = d["cov"]
        run2.font.size = Pt(9)
        run2.font.color.rgb = GREEN_ACCENT
        run2.font.bold = True
        cell.vertical_anchor = MSO_ANCHOR.TOP

        # Cols 1-3: Area data
        for c_idx, area_key in enumerate(["ps_pt", "digital_coe", "infra"]):
            cell = table.cell(r_idx + 1, c_idx + 1)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            set_cell_rich(cell, d[area_key], FONT_BODY)
            cell.vertical_anchor = MSO_ANCHOR.TOP

    # Footer elements (matching existing design)
    footer_bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                        Inches(10.5), Inches(7.0), Inches(1.5), Inches(0.5))
    footer_bg.fill.solid()
    footer_bg.fill.fore_color.rgb = WHITE
    footer_bg.line.fill.background()

    footer_accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                            Inches(12.1), Inches(7.1), Inches(1.0), Inches(0.2))
    footer_accent.fill.solid()
    footer_accent.fill.fore_color.rgb = BLUE_LINK
    footer_accent.line.fill.background()


# Move new slides to position 5 (after the "02 Existing Assets" divider)
sldIdLst = prs.slides._sldIdLst
for _ in range(3):
    last = sldIdLst[-1]
    sldIdLst.remove(last)
    sldIdLst.insert(5, last)

prs.save(str(OUTPUT))
print(f"Saved {len(prs.slides)} slides → {OUTPUT}")
