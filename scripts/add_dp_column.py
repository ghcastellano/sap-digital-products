"""
Add 'Digital Products' reuse column to existing table slides.
Removes slides 4-6, recreates with 5 columns.
"""
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from urllib.parse import quote
import re

INPUT = Path("/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Digital Products - Prior Materials Review .pptx")

GREEN_ACCENT = RGBColor(0x00, 0xA4, 0x43)
BLUE_PRIMARY = RGBColor(0x00, 0x20, 0x60)
BLUE_LINK = RGBColor(0x00, 0x72, 0xBC)
DP_COLOR = RGBColor(0x00, 0x5B, 0x96)  # Deep blue for Digital Products column
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
CHARCOAL = RGBColor(0x2C, 0x3E, 0x50)
DARK_GRAY = RGBColor(0x7F, 0x8C, 0x8D)
MID_GRAY = RGBColor(0x95, 0xA5, 0xA6)
LIGHT_GRAY = RGBColor(0xF5, 0xF5, 0xF5)
LIGHT_BLUE_BG = RGBColor(0xE8, 0xF0, 0xFE)

BASE = "/Users/gustavohenriquecastellano/Downloads/Base Apresentações Projeto Digital Product"
F = {
    "nWOW":     "nWOW - New Ways of Working - P&S Lean Operating Model 03.10.2026 v1.2.pptx",
    "PM":       "[00] Project Management Playbook - P&S Ways of Working 10.28.25 v.1.0-2.pptx",
    "PD":       "[00] Product Discovery Playbook - P&S Ways of Working 02.03.26 v.1.3 +FIGMA.pptx",
    "BRP":      "BRP_Q2_2026_P&S Networks IT - New Ways of Working - v1.4.4 03.09.2026.pptx",
    "CoE":      "[12] Digital CoE Operating Model - Playbook - 01.05.2026 V.2.5. FINAL pptx.pptx",
    "Agility":  "[17] Agility Playbook - Digital Center of Excellence 06.18.2025 V.1.7.pptx",
    "NextGen":  "[33] NextGen 2.0 - Jira Management User Guide for CX Digitall CoE 27.06.2024.pptx",
    "CX JIRA":  "[CX] JIRA Agile Management Playbook - Center of Excellence 25.03.2024 V.1.0.pdf",
    "NTTDATA":  "NTTDATA-Digital Strategy - NextGen Produc Ops Org Design SEAL-SIX Model v.1.7 06.29.pptx",
    "Avangrid": "Agile Management Model _ Proposal Avangrid Architecture V.3.1 18.01.2024 3.pptx",
    "ProjBook": "Project_Book_Avangrid_Template v1.0.pptx",
    "TPO":      "TPO FLOW Project - JIRA Training Guide 05-13-2025.pptx",
}


# References in DELIVERABLES are maintained as visible numbers for authoring convenience.
# Convert them back to actual slide indexes so PowerPoint navigation matches the source file.

HIDDEN = {
    "Avangrid": [23, 37, 49, 53, 54, 55],
    "BRP": [11],
    "NTTDATA": [25],
    "PD": [22],
    "PM": [5, 61, 71, 75, 86, 87, 110, 113, 114, 129, 130],
    "Agility": [35, 37, 86, 105, 106, 107, 108, 109, 110, 111, 112],
    "nWOW": [10, 11, 12, 13, 36, 60],
}


def _visible_to_actual(file_key: str, visible_num: int) -> int:
    hidden = set(HIDDEN.get(file_key, []))
    if not hidden:
        return visible_num
    actual = visible_num
    while True:
        hidden_before = sum(1 for h in hidden if h <= actual)
        candidate = visible_num + hidden_before
        if candidate == actual:
            break
        actual = candidate
    return actual

def _normalize_slide_ref(file_key: str, ref_str: str) -> str:
    def convert_token(token: str) -> str:
        token = token.strip()
        if '-' in token and all(part.strip().isdigit() for part in token.split('-', 1)):
            a, b = [int(x.strip()) for x in token.split('-', 1)]
            return f"{_visible_to_actual(file_key, a)}-{_visible_to_actual(file_key, b)}"
        if token.isdigit():
            return str(_visible_to_actual(file_key, int(token)))
        return token

    m = re.search(r'(Sl\.\s*)(.+)$', ref_str)
    if not m:
        return ref_str
    prefix = m.group(1)
    body = m.group(2)
    parts = [convert_token(x) for x in body.split(',')]
    return prefix + ', '.join(parts)


def _clean_approach_text(text: str) -> str:
    # Remove inline "(sl. ...)" hints from approach text; keep references only in the ref list.
    return re.sub(r'\s*\(sl\.[^)]+\)', '', text, flags=re.IGNORECASE).strip()


def uri(key):
    return "file:///" + quote(str(Path(BASE) / F[key]), safe="/:@")


def add_hyperlink(run, url):
    r = run._r
    hl = r.makeelement(qn('a:hlinkClick'), {})
    rId = run.part.relate_to(url, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)
    hl.set(qn('r:id'), rId)
    rPr = r.find(qn('a:rPr'))
    if rPr is None:
        rPr = r.makeelement(qn('a:rPr'), {})
        r.insert(0, rPr)
    rPr.append(hl)


 # Slide references below use ACTUAL slide indexes from source files.
DELIVERABLES = [
    {
        "name": "E2E Operating\nModel",
        "cov": "● | ◕ | ◔",
        "ps_pt": {
            "what": "6-phase lifecycle with Kanban management model. E2E Management Process defined. Demand funnel with portfolio prioritization. Big Room Planning cadence. Governance forums at 3 levels. Methodology detail per lifecycle phase.",
            "refs": [("nWOW", "Sl. 30-50"), ("PM", "Sl. 8, 69-81"), ("BRP", "Sl. 1-9")]},
        "digital_coe": {
            "what": "E2E Operating Model with governance levels. Product Management Life-Cycle Process. Discovery Upstream workflow. Delivery workflow with sprint model. Governance cadences and learn cycles.",
            "refs": [("CoE", "Sl. 6-9, 35-56, 57-77")]},
        "infra": {
            "what": "Phased rollout methodology: Pilot → Expansion → Optimization. Team Topologies framework: Enabling, Stream-Aligned, Platform, Component types with interaction modes.",
            "refs": [("Avangrid", "Sl. 10, 28-31")]},
        "dp": "REUSE: Adapt P&S 6-phase lifecycle structure but reframe phases for SAP context: Intake → Assessment → Design → Build → Deploy → Support. Reuse CoE governance cadences (learn cycles, sprint cadences). Reuse Avangrid Team Topologies to define SAP's team type. NEW: Define SAP-specific intake entry point (single front door for all areas requesting SAP work). Define handoff protocol between SAP and consuming teams.",
    },
    {
        "name": "Value &\nPrioritization",
        "cov": "◕ | ◑ | ○",
        "ps_pt": {
            "what": "Portfolio Prioritization Criteria v1.0. Business Value Criteria. User Value Criteria. Portfolio & Effort Estimation Logic. Demand Funnel Prioritization.",
            "refs": [("nWOW", "Sl. 37, 54, 63-71")]},
        "digital_coe": {
            "what": "Strategic Initiatives Prioritization Method. PBI formula (BV+UX)/E. Estimation Logic. Planning Poker. Portfolio Kanban with prioritization.",
            "refs": [("CoE", "Sl. 25-27, 81-86")]},
        "infra": {"what": "", "refs": []},
        "dp": "REUSE: Adapt nWOW Prioritization Criteria adding SAP-specific dimensions: requesting area priority, technical complexity, cross-area dependency impact, capacity utilization. Reuse CoE PBI formula structure but replace Business Value with 'Requesting Area Priority Score'. Reuse Estimation Logic from both areas. NEW: Build cross-area demand balancing mechanism — SAP must weigh competing requests from P&S, Digital CoE, and Infra simultaneously.",
    },
    {
        "name": "Demand Typology\n& Intake",
        "cov": "● | ◔ | ○",
        "ps_pt": {
            "what": "Portfolio Framework & Demand Management Funnel. Future State Framework. E2E Management Process. Demand Funnel Prioritization with classification. Methodology detail for intake phases.",
            "refs": [("nWOW", "Sl. 31-41"), ("BRP", "Sl. 1-9, 15")]},
        "digital_coe": {
            "what": "Quarterly Program Plan. Demand sources from executive levels, portfolio management funnel. Lean Business Case template for intake.",
            "refs": [("CoE", "Sl. 11, 30-33")]},
        "infra": {"what": "", "refs": []},
        "dp": "REUSE: Adopt P&S single front door concept and demand funnel visualization format. Reuse demand classification (Strategic/Operational/Regulatory) but add SAP typology: Enhancement, Incident/Break-Fix, New Integration, Config Change. Reuse CoE Lean Business Case template for non-standard requests. NEW: Define which demand types require full intake vs. lightweight request vs. self-service. Tag every request with requesting area + SAP module.",
    },
    {
        "name": "Roles &\nRACI",
        "cov": "● | ◔ | ◔",
        "ps_pt": {
            "what": "RACI Matrix templates and methodology. Project Manager Role definition. Stakeholder Mapping with engagement matrix. Resource Planning framework. Crew roles and work model from nWOW.",
            "refs": [("PM", "Sl. 99-101, 106-119"), ("nWOW", "Sl. 13, 20-22, 26, 29")]},
        "digital_coe": {
            "what": "Organizational structure with governance principles. Team topologies and crew structures. Roles implied through Scrum adoption in Agility Playbook.",
            "refs": [("CoE", "Sl. 12-22"), ("Agility", "Sl. 18-25")]},
        "infra": {
            "what": "Team Topologies framework: 4 team types with 3 interaction modes (Collaboration, X-as-a-Service, Facilitating).",
            "refs": [("Avangrid", "Sl. 28-31")]},
        "dp": "REUSE: Adopt PM Playbook RACI template structure and Stakeholder Mapping format. Reuse nWOW crew roles model adapted for SAP team composition. Reuse Avangrid interaction modes to define how SAP relates to consuming teams (X-as-a-Service for standard work, Collaboration for complex). NEW: Define 'Service Requester' role (from consuming area) and 'SAP Technical Lead' role. SAP = Responsible for delivery; consuming area = Accountable for business outcome.",
    },
    {
        "name": "Tool Ecosystem\n& Jira Config",
        "cov": "◕ | ● | ◑",
        "ps_pt": {
            "what": "JIRA Overview and issue type hierarchy. Kanban workflows in Jira. Board creation and customization. Admin levels and permissions. Program dashboards.",
            "refs": [("PM", "Sl. 82-98")]},
        "digital_coe": {
            "what": "Complete JIRA Agile Management Playbook (41 pages): board design, filters, custom fields, workflow states. NextGen Jira Guide: board design, workflow config. Portfolio view.",
            "refs": [("CX JIRA", "Full doc (41p)"), ("NextGen", "Sl. 4, 19, 21, 23")]},
        "infra": {
            "what": "JIRA Training Guide: process setup, Kanban boards, filters, admin levels. Hands-on training format.",
            "refs": [("TPO", "Sl. 2-3, 10-11")]},
        "dp": "REUSE: Adopt Digital CoE's JIRA Playbook as baseline config guide (most comprehensive: 41p). Reuse PM Playbook board creation and workflow patterns. Reuse TPO training format for SAP team onboarding. NEW: Add custom fields: 'Requesting Area', 'SAP Module' (FI/MM/HR/BW), 'Request Type'. Create filtered views per consuming area for self-serve status. Add SLA tracking with due-date automation.",
    },
    {
        "name": "Portfolio &\nPPMNow Concept",
        "cov": "◕ | ◕ | ◑",
        "ps_pt": {
            "what": "Portfolio Kanban. Program Kanban. Governance model with 3 connecting levels. Big Room Planning. Portfolio estimation logic.",
            "refs": [("nWOW", "Sl. 31-32, 46-47, 52, 63-65, 70-74"), ("PM", "Sl. 15, 73-74")]},
        "digital_coe": {
            "what": "Portfolio Governance. Portfolio Kanban with roadmap. Portfolio Learn Cycle (quarterly). Portfolio Metrics. PPM Practices and Architecture.",
            "refs": [("CoE", "Sl. 23-29, 59-77, 79-80, 100-104")]},
        "infra": {
            "what": "Program-level portfolio management. Investment model per unit served. Deployment governance.",
            "refs": [("Avangrid", "Sl. 9, 36, 47-49")]},
        "dp": "REUSE: Adopt nWOW Portfolio Kanban visualization format — it's the clearest. Reuse CoE Portfolio Learn Cycle cadence for quarterly SAP portfolio reviews. Reuse CoE PPM Practices for portfolio architecture. NEW: SAP portfolio must be segmented by consuming area (not internal teams). Show capacity allocation per area. Roadmap driven by cross-area demand patterns. PPMNow should show which areas consume which SAP capabilities.",
    },
    {
        "name": "Baseline &\nDemand Map",
        "cov": "◔ | ◔ | ◔",
        "ps_pt": {
            "what": "Strategy & Vision section provides baseline methodology. Demand flow visualization in Future State Framework. Demand funnel mapping. Project plan formats and timeline views. Discovery research methods.",
            "refs": [("nWOW", "Sl. 6-9, 32-33, 37"), ("PM", "Sl. 120-125"), ("PD", "Sl. 21, 37, 39")]},
        "digital_coe": {
            "what": "Demand sources mapping. Portfolio management team definition. Current state visualization approach.",
            "refs": [("CoE", "Sl. 32-33, 81, 105")]},
        "infra": {
            "what": "Project Book template with phase tracking (Discover → Solution → Plan → Build → Operate).",
            "refs": [("ProjBook", "Full template")]},
        "dp": "REUSE: Adopt nWOW baseline methodology format for SAP current state diagnosis — same interview + assessment structure. Reuse demand flow visualization format but redesign as hub-and-spoke with SAP at center. Reuse PD research methods for stakeholder interviews. Reuse Infra Project Book template for documentation. NEW: Map SAP-specific demand sources, channels, volumes per area. Assess current 'time to fulfill' per request type.",
    },
    {
        "name": "Pilot\nExecution",
        "cov": "◔ | ◔ | ◑",
        "ps_pt": {
            "what": "Validation Report template with user feedback & key insights. Feasibility assessment matrix. Go/No-Go gate methodology.",
            "refs": [("PD", "Sl. 66-67")]},
        "digital_coe": {
            "what": "GO/No-Go gate process. Sprint Report Canvas for pilot tracking. SEAL-SIX maturity assessment and MVP pilot approach. Agility health assessment.",
            "refs": [("CoE", "Sl. 62, 88-89"), ("NTTDATA", "Sl. 9-10"), ("Agility", "Sl. 24")]},
        "infra": {
            "what": "Pilot → Expansion → Optimization deployment methodology. Phased expansion governance. Project Book for structured pilot documentation.",
            "refs": [("Avangrid", "Sl. 10, 22, 43-44"), ("ProjBook", "Sl. 8, 11")]},
        "dp": "REUSE: Adopt Infra's Pilot → Expansion → Optimization methodology — best fit for SAP rollout across OPCOs. Reuse PD Validation Report format for pilot feedback. Reuse CoE Sprint Report Canvas for sprint-level tracking. Reuse CoE Go/No-Go gate process. NEW: BRS pilot must validate both the process AND the service relationship. Success criteria: consuming area satisfaction + cycle time reduction. Scope: 1 OPCO, 1 project, Small/Medium.",
    },
    {
        "name": "Ways of Working\nPlaybook",
        "cov": "● | ◕ | ○",
        "ps_pt": {
            "what": "Product Discovery Playbook v1.3 (114 slides): Design Thinking, research, ideation, validation, prototyping, backlog building. Project Management Playbook v1.0 (150 slides): governance, backlog, Kanban, RACI, KPIs, Jira, contracts.",
            "refs": [("PD", "Full doc (114 sl.)"), ("PM", "Full doc (150 sl.)")]},
        "digital_coe": {
            "what": "Agility Playbook v1.7 (111 slides): Scrum/Kanban methods. Product Blueprint & requirements. Backlog management. Agile metrics & assessments.",
            "refs": [("Agility", "Full doc (111 sl.)")]},
        "infra": {"what": "", "refs": []},
        "dp": "REUSE: Adopt P&S Playbook modular structure (chapters by lifecycle phase) as template — 264 slides of proven structure. Reuse PM Playbook Kanban management section as-is for SAP execution methodology. Reuse PM Playbook governance forums section. Reuse Agility Playbook Scrum/Kanban reference for team-level practices. NEW: Add SAP-specific chapters: 'Request Assessment & Scoping', 'Cross-Area Dependency Management', 'Service Relationship Protocol'.",
    },
    {
        "name": "KPI &\nReporting",
        "cov": "● | ◕ | ○",
        "ps_pt": {
            "what": "Multi-level Metrics Ecosystem: Portfolio, Program, Product, Project levels. Jira Kanban workflows for metrics. Dashboard Program Health. Dashboard Project Health. Continuous Improvement. Scorecards.",
            "refs": [("nWOW", "Sl. 76-91"), ("PM", "Sl. 87-88, 126-134")]},
        "digital_coe": {
            "what": "Portfolio Metrics definition. Team of Teams Metrics. Team Metrics. Data sources for measurements. Capacity measurement. Agile Teams Target Scorecard.",
            "refs": [("CoE", "Sl. 79-80, 93-98, 106-107")]},
        "infra": {"what": "", "refs": []},
        "dp": "REUSE: Adopt P&S multi-level metrics structure — Portfolio → Program → Team cascade. Reuse dashboard layouts from nWOW and PM Playbook. Reuse CoE scorecard approach. Reuse Jira-based metrics automation pattern. NEW: Add service-oriented KPIs: Cycle Time per request type, SLA compliance rate, Queue depth by consuming area, Dependency wait time. Keep Lead Time, Throughput, WIP (universal). Add 'Consuming Area Satisfaction'.",
    },
]


def set_cell_rich(cell, area_data, font_size=Pt(7)):
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
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = _clean_approach_text(area_data["what"])
    run.font.size = font_size
    run.font.color.rgb = CHARCOAL
    for ref_key, ref_slides in area_data["refs"]:
        full_name = F[ref_key]
        actual_ref = _normalize_slide_ref(ref_key, ref_slides)
        p = tf.add_paragraph()
        p.space_before = Pt(1)
        ra = p.add_run()
        ra.text = "→ "
        ra.font.size = Pt(6)
        ra.font.color.rgb = DARK_GRAY
        rf = p.add_run()
        rf.text = full_name
        rf.font.size = Pt(6)
        rf.font.color.rgb = BLUE_LINK
        rf.font.underline = True
        try:
            add_hyperlink(rf, uri(ref_key))
        except:
            pass
        rs = p.add_run()
        rs.text = f"  |  {actual_ref}"
        rs.font.size = Pt(6)
        rs.font.bold = True
        rs.font.color.rgb = CHARCOAL


def set_dp_cell(cell, text, font_size=Pt(7)):
    """Set Digital Products reuse cell with REUSE/NEW highlighting."""
    tf = cell.text_frame
    tf.word_wrap = True
    tf.clear()

    parts = text.split("NEW: ")
    reuse_part = parts[0]
    new_part = parts[1] if len(parts) > 1 else ""

    # REUSE section
    if reuse_part:
        reuse_text = reuse_part.replace("REUSE: ", "")
        p = tf.paragraphs[0]
        rl = p.add_run()
        rl.text = "REUSE: "
        rl.font.size = font_size
        rl.font.bold = True
        rl.font.color.rgb = GREEN_ACCENT
        rt = p.add_run()
        rt.text = reuse_text.strip()
        rt.font.size = font_size
        rt.font.color.rgb = CHARCOAL

    # NEW section
    if new_part:
        p = tf.add_paragraph()
        p.space_before = Pt(3)
        nl = p.add_run()
        nl.text = "NEW: "
        nl.font.size = font_size
        nl.font.bold = True
        nl.font.color.rgb = DP_COLOR
        nt = p.add_run()
        nt.text = new_part.strip()
        nt.font.size = font_size
        nt.font.color.rgb = CHARCOAL


# ═══════════════════════════════════════════
prs = Presentation(str(INPUT))

# Remove table slides (indices 3, 4, 5)
for si in range(5, 2, -1):
    sldId = prs.slides._sldIdLst[si]
    rId = sldId.get(qn('r:id'))
    prs.slides._sldIdLst.remove(sldId)
    prs.part.drop_rel(rId)

print(f"After removal: {len(prs.slides)} slides")

# 3 table slides: 4, 3, 3
batches = [
    (DELIVERABLES[0:4], "Existing Assets — Operating Model, Value Framework, Demand Intake, RACI (1/3)"),
    (DELIVERABLES[4:7], "Existing Assets — Tool Ecosystem, Portfolio & PPMNow, Baseline & Demand Map (2/3)"),
    (DELIVERABLES[7:10], "Existing Assets — Pilot Execution, Playbook, KPI & Reporting (3/3)"),
]

for batch, btitle in batches:
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.1), Inches(12.3), Inches(0.4))
    p = title_box.text_frame.paragraphs[0]
    run = p.add_run()
    run.text = btitle
    run.font.size = Pt(12)
    run.font.color.rgb = GREEN_ACCENT

    legend = slide.shapes.add_textbox(Inches(0.5), Inches(0.45), Inches(12.3), Inches(0.2))
    p = legend.text_frame.paragraphs[0]
    run = p.add_run()
    run.text = "● Complete  ◕ Well-defined  ◑ Partial  ◔ Referenced  ○ None"
    run.font.size = Pt(6)
    run.font.color.rgb = MID_GRAY

    n_rows = len(batch) + 1
    n_cols = 5
    ts = slide.shapes.add_table(n_rows, n_cols, Inches(0.2), Inches(0.7), Inches(12.9), Inches(6.4))
    table = ts.table

    col_widths = [Inches(1.1), Inches(3.0), Inches(2.8), Inches(2.2), Inches(3.8)]
    for i, w in enumerate(col_widths):
        table.columns[i].width = int(w)

    headers = ["Deliverable", "P&S / P&T", "Digital CoE", "Infrastructure", "Digital Products (Reuse)"]
    hdr_colors = [BLUE_PRIMARY, BLUE_PRIMARY, BLUE_PRIMARY, BLUE_PRIMARY, DP_COLOR]
    for i, (h, hc) in enumerate(zip(headers, hdr_colors)):
        cell = table.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = hc
        tf = cell.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = h
        run.font.size = Pt(7)
        run.font.bold = True
        run.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

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
        run.font.size = Pt(8)
        run.font.bold = True
        run.font.color.rgb = CHARCOAL
        p2 = tf.add_paragraph()
        run2 = p2.add_run()
        run2.text = d["cov"]
        run2.font.size = Pt(8)
        run2.font.color.rgb = GREEN_ACCENT
        run2.font.bold = True
        cell.vertical_anchor = MSO_ANCHOR.TOP

        # Cols 1-3: Area data
        for c_idx, key in enumerate(["ps_pt", "digital_coe", "infra"]):
            cell = table.cell(r_idx + 1, c_idx + 1)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            set_cell_rich(cell, d[key], Pt(7))
            cell.vertical_anchor = MSO_ANCHOR.TOP

        # Col 4: Digital Products reuse
        cell = table.cell(r_idx + 1, 4)
        cell.fill.solid()
        cell.fill.fore_color.rgb = LIGHT_BLUE_BG if r_idx % 2 == 0 else RGBColor(0xD6, 0xE4, 0xF0)
        set_dp_cell(cell, d["dp"], Pt(7))
        cell.vertical_anchor = MSO_ANCHOR.TOP

    # Footer
    fb = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(10.5), Inches(7.0), Inches(1.5), Inches(0.5))
    fb.fill.solid()
    fb.fill.fore_color.rgb = WHITE
    fb.line.fill.background()
    fa = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(12.1), Inches(7.1), Inches(1.0), Inches(0.2))
    fa.fill.solid()
    fa.fill.fore_color.rgb = BLUE_LINK
    fa.line.fill.background()

# Move to position 3
sldIdLst = prs.slides._sldIdLst
for _ in range(3):
    last = sldIdLst[-1]
    sldIdLst.remove(last)
    sldIdLst.insert(3, last)

prs.save(str(INPUT))
print(f"Saved {len(prs.slides)} slides → {INPUT}")
