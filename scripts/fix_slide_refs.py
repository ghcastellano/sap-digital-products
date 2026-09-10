"""
Fix slide references in the existing PPTX — verified from KB.
Removes current table slides (6-8) and recreates with corrected data.
"""
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from urllib.parse import quote

INPUT = Path("/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Digital Products - Prior Materials Review .pptx")

GREEN_ACCENT = RGBColor(0x00, 0xA4, 0x43)
BLUE_PRIMARY = RGBColor(0x00, 0x20, 0x60)
BLUE_LINK = RGBColor(0x00, 0x72, 0xBC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
CHARCOAL = RGBColor(0x2C, 0x3E, 0x50)
DARK_GRAY = RGBColor(0x7F, 0x8C, 0x8D)
MID_GRAY = RGBColor(0x95, 0xA5, 0xA6)
LIGHT_GRAY = RGBColor(0xF5, 0xF5, 0xF5)

BASE = "/Users/gustavohenriquecastellano/Downloads/Base Apresentações Projeto Digital Product"

# Full filenames
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


# ── CORRECTED DATA (verified from KB) ──
DELIVERABLES = [
    {
        "name": "E2E Operating\nModel",
        "cov": "● | ◕ | ◔",
        "ps_pt": {
            "what": "6-phase lifecycle with Kanban management model. E2E Management Process defined. Demand funnel with portfolio prioritization. Big Room Planning cadence (sl. 51-52). Governance forums at 3 levels (sl. 54-55). Methodology detail per phase (sl. 39-50).",
            "refs": [("nWOW", "Sl. 34-55 (Process, Demand, Governance)"), ("PM", "Sl. 9, 72-85 (E2E Process, Kanban Model)"), ("BRP", "Sl. 1-9 (Lifecycle overview)")]},
        "digital_coe": {
            "what": "E2E Operating Model with governance levels (sl. 6-7). Product Management Life-Cycle Process (sl. 8). Discovery Upstream workflow (sl. 10, 35-48). Delivery workflow with sprint model (sl. 49-56). Governance cadences and learn cycles (sl. 57-77).",
            "refs": [("CoE", "Sl. 6-9 (E2E Model), Sl. 35-56 (Discovery+Delivery), Sl. 57-77 (Cadences)")]},
        "infra": {
            "what": "Deployment methodology with phased rollout: Pilot → Expansion → Optimization (sl. 10). Team Topologies framework: Enabling, Stream-Aligned, Platform, Component types with interaction modes (sl. 29-32).",
            "refs": [("Avangrid", "Sl. 10 (Deployment), Sl. 29-32 (Team Topologies)")]},
    },
    {
        "name": "Value &\nPrioritization",
        "cov": "◕ | ◑ | ○",
        "ps_pt": {
            "what": "Portfolio Prioritization Criteria v1.0 (sl. 59). Business Value Criteria (sl. 72). User Value Criteria (sl. 74-76). Portfolio Estimation Logic (sl. 69-71, 77). Effort Estimation Logic (sl. 70). Demand Funnel Prioritization (sl. 42).",
            "refs": [("nWOW", "Sl. 42 (Demand Funnel), Sl. 59 (Prioritization Criteria), Sl. 69-77 (Estimation & Value)")]},
        "digital_coe": {
            "what": "Strategic Initiatives Prioritization Method (sl. 81). Criteria with PBI formula (BV+UX)/E (sl. 84). Estimation Logic (sl. 83). Planning Poker technique (sl. 86). Portfolio Kanban with prioritization (sl. 25-27).",
            "refs": [("CoE", "Sl. 25-27 (Portfolio Kanban), Sl. 81-86 (Prioritization & Estimation)")]},
        "infra": {"what": "", "refs": []},
    },
    {
        "name": "Demand Typology\n& Intake",
        "cov": "● | ◔ | ○",
        "ps_pt": {
            "what": "Portfolio Framework & Demand Management Funnel (sl. 35). Future State Framework (sl. 36-37). E2E Management Process (sl. 38). Demand Funnel Prioritization with classification (sl. 42-43). Methodology detail for intake phases (sl. 44-46).",
            "refs": [("nWOW", "Sl. 35-46 (Demand Framework & Intake)"), ("BRP", "Sl. 1-9, 16 (Demand lifecycle)")]},
        "digital_coe": {
            "what": "Quarterly Program Plan (sl. 11). Demand sources: strategic initiatives from executive levels, portfolio management funnel (sl. 32-33). Lean Business Case template for intake (sl. 30).",
            "refs": [("CoE", "Sl. 11 (Program Plan), Sl. 30-33 (Demand & LBC)")]},
        "infra": {"what": "", "refs": []},
    },
    {
        "name": "Roles &\nRACI",
        "cov": "● | ◔ | ◔",
        "ps_pt": {
            "what": "RACI Matrix templates and methodology (sl. 125-130). Project Manager Role definition (sl. 105-107). Stakeholder Mapping with engagement matrix (sl. 115-118, 120). Resource Planning framework (sl. 121-124). Crew roles and work model defined in nWOW (sl. 24-26, 30, 33).",
            "refs": [("PM", "Sl. 105-107 (PM Role), Sl. 115-120 (Stakeholders), Sl. 121-124 (Resources), Sl. 125-130 (RACI)"), ("nWOW", "Sl. 17 (Roles Intro), Sl. 24-26 (Crew Roles), Sl. 30 (Essential Roles), Sl. 33 (Roles Intro)")]},
        "digital_coe": {
            "what": "Organizational structure with governance principles (sl. 12-14). Team topologies and crew structures (sl. 15-22). Roles implied through Scrum adoption in Agility Playbook.",
            "refs": [("CoE", "Sl. 12-22 (Org Structure & Governance)"), ("Agility", "Sl. 18-25 (Scrum Roles)")]},
        "infra": {
            "what": "Team Topologies framework: 4 team types with 3 interaction modes (Collaboration, X-as-a-Service, Facilitating). Defines how Enabling and Platform teams relate to Stream-Aligned teams.",
            "refs": [("Avangrid", "Sl. 29-32 (Team Topologies)")]},
    },
    {
        "name": "Tool Ecosystem\n& Jira Config",
        "cov": "◕ | ● | ◑",
        "ps_pt": {
            "what": "JIRA Overview and issue type hierarchy (sl. 90-92). Kanban workflows in Jira (sl. 88-89, 95). Board creation and customization (sl. 101-104). Admin levels and permissions (sl. 97-100). Dashboards for programs (sl. 93-94).",
            "refs": [("PM", "Sl. 88-104 (Full Jira Section)")]},
        "digital_coe": {
            "what": "Complete JIRA Agile Management Playbook (41 pages) covering board design, filters, custom fields, workflow states. NextGen Jira Guide: board design, workflow config (sl. 4, 19, 21, 23). Portfolio view in Jira.",
            "refs": [("CX JIRA", "Full document (41 pages)"), ("NextGen", "Sl. 4, 19, 21, 23")]},
        "infra": {
            "what": "JIRA Training Guide: process setup, Kanban boards, filters, admin levels (sl. 2-3, 10-11). Hands-on training format.",
            "refs": [("TPO", "Sl. 2-3, 10-11 (Jira Training)")]},
    },
    {
        "name": "Portfolio &\nPPMNow Concept",
        "cov": "◕ | ◕ | ◑",
        "ps_pt": {
            "what": "Portfolio Kanban (sl. 57, 79). Program Kanban (sl. 80). Governance model with 3 levels connecting (sl. 76-78). Big Room Planning (sl. 51-52). Portfolio estimation logic (sl. 69-71). Demand management funnel (sl. 35-37).",
            "refs": [("nWOW", "Sl. 35-37 (Demand Funnel), Sl. 51-52 (BRP), Sl. 57 (Portfolio Kanban), Sl. 69-71 (Estimation), Sl. 76-80 (Governance & Kanban)"), ("PM", "Sl. 16 (Value Charter), Sl. 77-78 (Portfolio/Program Kanban)")]},
        "digital_coe": {
            "what": "Portfolio Governance (sl. 23). Portfolio Kanban with roadmap (sl. 25-29). Portfolio Learn Cycle quarterly (sl. 59-60, 64, 67, 72-73). Portfolio Metrics (sl. 79-80, 93). PPM Practices and Architecture (sl. 100-104).",
            "refs": [("CoE", "Sl. 23-29 (Portfolio Governance & Kanban), Sl. 59-77 (Learn Cycles), Sl. 79-80 (Metrics), Sl. 100-104 (PPM Practices)")]},
        "infra": {
            "what": "Program-level portfolio management. Investment model per unit served (sl. 50-52). Deployment governance (sl. 38, 46).",
            "refs": [("Avangrid", "Sl. 9 (Strategy), Sl. 38 (Governance), Sl. 50-52 (Investment)")]},
    },
    {
        "name": "Baseline &\nDemand Map",
        "cov": "◔ | ◔ | ◔",
        "ps_pt": {
            "what": "nWOW Strategy & Vision section provides baseline methodology (sl. 6-13). Demand flow visualization in Future State Framework (sl. 36-37). Demand funnel mapping (sl. 42). Project plan formats and timeline views (sl. 131-136). Discovery research methods (sl. 22, 38, 40).",
            "refs": [("nWOW", "Sl. 6-13 (Strategy/OKR Baseline), Sl. 36-37 (Future State), Sl. 42 (Demand Funnel)"), ("PM", "Sl. 131-136 (Project Plan & Timelines)"), ("PD", "Sl. 22, 38, 40 (Discovery Research)")]},
        "digital_coe": {
            "what": "Demand sources mapping (sl. 32). Portfolio management team definition (sl. 33). Current state visualization approach (sl. 81, 105).",
            "refs": [("CoE", "Sl. 32-33 (Demand Sources), Sl. 81, 105 (Assessment)")]},
        "infra": {
            "what": "Project Book template with phase tracking (Discover → Solution → Plan → Build → Operate).",
            "refs": [("ProjBook", "Full template")]},
    },
    {
        "name": "Pilot\nExecution",
        "cov": "◔ | ◔ | ◑",
        "ps_pt": {
            "what": "Validation Report template with user feedback & key insights (sl. 67-68). Feasibility assessment matrix. Go/No-Go gate methodology in Discovery Playbook.",
            "refs": [("PD", "Sl. 67-68 (Validation & Feasibility)")]},
        "digital_coe": {
            "what": "GO/No-Go gate process (sl. 62). Sprint Report Canvas for pilot tracking (sl. 88-89). Maturity assessment with SEAL-SIX model and MVP pilot approach (NTTDATA sl. 9-10). Agility health assessment (Agility sl. 24).",
            "refs": [("CoE", "Sl. 62 (Go/No-Go), Sl. 88-89 (Sprint Report)"), ("NTTDATA", "Sl. 9-10 (Maturity & MVP Pilot)"), ("Agility", "Sl. 24 (Health Assessment)")]},
        "infra": {
            "what": "Deployment methodology: Pilot → Expansion → Optimization (sl. 10, 23). Agile rollout governance with phased expansion (sl. 45-46). Project Book for structured pilot documentation (sl. 8, 11).",
            "refs": [("Avangrid", "Sl. 10, 23 (Deployment), Sl. 45-46 (Expansion)"), ("ProjBook", "Sl. 8, 11 (Pilot Documentation)")]},
    },
    {
        "name": "Ways of Working\nPlaybook",
        "cov": "● | ◕ | ○",
        "ps_pt": {
            "what": "Product Discovery Playbook v1.3 (114 slides): Design Thinking, research, ideation, validation, prototyping, portfolio estimation, backlog building. Project Management Playbook v1.0 (150 slides): governance, backlog, Kanban, RACI, KPIs, Jira, contracts.",
            "refs": [("PD", "Full document (114 slides)"), ("PM", "Full document (150 slides)")]},
        "digital_coe": {
            "what": "Agility Playbook v1.7 (111 slides): Scrum/Kanban methods (sl. 18-31, 106). Product Blueprint & requirements (sl. 33-50). Backlog management (sl. 50-83). Agile metrics & assessments (sl. 84-113).",
            "refs": [("Agility", "Full document (111 slides)")]},
        "infra": {"what": "", "refs": []},
    },
    {
        "name": "KPI &\nReporting",
        "cov": "● | ◕ | ○",
        "ps_pt": {
            "what": "Multi-level Metrics Ecosystem (sl. 82-84): Portfolio, Program, Product, Project levels. Jira Kanban workflows for metrics (sl. 87). Dashboard Program Health (sl. 88). Dashboard Project Health (sl. 92-93). Change Management metrics (sl. 95). Continuous Improvement (sl. 96). PM Playbook: Project Metrics (sl. 137-138). All Levels dashboard (sl. 139). Project Health Dashboard (sl. 142-143). Continuous Improvement (sl. 144). Scorecards (sl. 145).",
            "refs": [("nWOW", "Sl. 82-97 (Full Metrics Ecosystem)"), ("PM", "Sl. 93-94 (Program Dashboard), Sl. 137-145 (Metrics & Scorecards)")]},
        "digital_coe": {
            "what": "Portfolio Metrics definition (sl. 79-80, 93). Team of Teams Metrics (sl. 94). Team Metrics (sl. 95). Data sources for measurements (sl. 96). Capacity measurement (sl. 97-98). Agile Teams Target Scorecard (sl. 106-107).",
            "refs": [("CoE", "Sl. 79-80 (Portfolio Metrics), Sl. 93-98 (Team Metrics & Capacity), Sl. 106-107 (Scorecard)")]},
        "infra": {"what": "", "refs": []},
    },
]


def set_cell_rich(cell, area_data, font_size=Pt(8)):
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
    run.text = area_data["what"]
    run.font.size = font_size
    run.font.color.rgb = CHARCOAL

    for ref_key, ref_slides in area_data["refs"]:
        full_name = F[ref_key]
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
            add_hyperlink(run_file, uri(ref_key))
        except:
            pass

        run_sl = p.add_run()
        run_sl.text = f"  |  {ref_slides}"
        run_sl.font.size = Pt(7)
        run_sl.font.bold = True
        run_sl.font.color.rgb = CHARCOAL


# ═══════════════════════════════════════════
prs = Presentation(str(INPUT))

# Remove current table slides (indices 3, 4, 5)
for si in range(5, 2, -1):
    sldId = prs.slides._sldIdLst[si]
    rId = sldId.get(qn('r:id'))
    prs.slides._sldIdLst.remove(sldId)
    prs.part.drop_rel(rId)

print(f"After removal: {len(prs.slides)} slides")

# Recreate 3 table slides
batches = [
    (DELIVERABLES[0:4], "Existing Assets — Operating Model, Value Framework, Demand Intake, Roles & RACI (1/3)"),
    (DELIVERABLES[4:7], "Existing Assets — Tool Ecosystem, Portfolio & PPMNow, Baseline & Demand Map (2/3)"),
    (DELIVERABLES[7:10], "Existing Assets — Pilot Execution, Playbook, KPI & Reporting (3/3)"),
]

for batch, btitle in batches:
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.15), Inches(11.5), Inches(0.5))
    p = title_box.text_frame.paragraphs[0]
    run = p.add_run()
    run.text = btitle
    run.font.size = Pt(13)
    run.font.color.rgb = GREEN_ACCENT

    legend = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.5), Inches(0.25))
    p = legend.text_frame.paragraphs[0]
    run = p.add_run()
    run.text = "● Complete  ◕ Well-defined  ◑ Partial  ◔ Referenced  ○ None          → = clickable link to source file"
    run.font.size = Pt(7)
    run.font.color.rgb = MID_GRAY

    n_rows = len(batch) + 1
    ts = slide.shapes.add_table(n_rows, 4, Inches(0.4), Inches(0.9), Inches(12.5), Inches(6.2))
    table = ts.table

    col_widths = [Inches(1.4), Inches(4.2), Inches(3.9), Inches(3.0)]
    for i, w in enumerate(col_widths):
        table.columns[i].width = int(w)

    for i, h in enumerate(["Deliverable", "P&S / P&T", "Digital CoE", "Infrastructure"]):
        cell = table.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = BLUE_PRIMARY
        tf = cell.text_frame
        tf.clear()
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = h
        run.font.size = Pt(8)
        run.font.bold = True
        run.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    for r_idx, d in enumerate(batch):
        bg = WHITE if r_idx % 2 == 0 else LIGHT_GRAY
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

        for c_idx, key in enumerate(["ps_pt", "digital_coe", "infra"]):
            cell = table.cell(r_idx + 1, c_idx + 1)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            set_cell_rich(cell, d[key], Pt(8))
            cell.vertical_anchor = MSO_ANCHOR.TOP

    footer_bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(10.5), Inches(7.0), Inches(1.5), Inches(0.5))
    footer_bg.fill.solid()
    footer_bg.fill.fore_color.rgb = WHITE
    footer_bg.line.fill.background()
    footer_accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(12.1), Inches(7.1), Inches(1.0), Inches(0.2))
    footer_accent.fill.solid()
    footer_accent.fill.fore_color.rgb = BLUE_LINK
    footer_accent.line.fill.background()

# Move new slides to position 3 (after the first 3 existing slides)
sldIdLst = prs.slides._sldIdLst
for _ in range(3):
    last = sldIdLst[-1]
    sldIdLst.remove(last)
    sldIdLst.insert(3, last)

prs.save(str(INPUT))
print(f"Saved {len(prs.slides)} slides → {INPUT}")
