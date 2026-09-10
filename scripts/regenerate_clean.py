"""
Regenerate Prior Materials Review PPTX — V5 with KB-validated interpretive content
==================================================================================
All content pulled from ChromaDB KB (823 chunks, 15 documents).
Interpretive analysis: not just "what exists" but what it means, how it works,
and what's concretely reusable for SAP Digital Products.
"""
import sys
from pathlib import Path
from urllib.parse import quote

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

GREEN = RGBColor(0x00, 0xB0, 0x50)
ORANGE = RGBColor(0xE8, 0x6E, 0x00)
RED = RGBColor(0xC5, 0x30, 0x30)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
CHARCOAL = RGBColor(0x2C, 0x3E, 0x50)
MID_GRAY = RGBColor(0x95, 0xA5, 0xA6)
LIGHT_GRAY = RGBColor(0xF5, 0xF5, 0xF5)
DARK_GRAY = RGBColor(0x7F, 0x8C, 0x8D)
BLUE_LINK = RGBColor(0x00, 0x72, 0xBC)

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


def file_uri(short_name):
    fname = FILES.get(short_name, short_name)
    full_path = str(Path(BASE_PATH) / fname)
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


# ═══════════════════════════════════════════════════════════════════════════════
# DELIVERABLES — Interpretive content sourced from KB (823 chunks, validated)
# ═══════════════════════════════════════════════════════════════════════════════
DELIVERABLES = [
    {
        "name": "E2E Operating\nModel",
        "cov": "● | ◕ | ◔",
        "ps_pt": {
            "what": (
                "P&S built a fully structured 4-phase lifecycle (Discovery → Materialization → Creation → Ongoing) "
                "with 15+ sub-steps mapped E2E (Sl. 38). Each phase has defined artifacts: Problem Statement Card, "
                "Solution Proposal, Project Book, Budget Release. The governance uses a 3-tier Flight Level model "
                "(Portfolio → Program → Project) with Kanban at each level (Sl. 76). Decision Points and phase gates "
                "are visually mapped with entry/exit criteria. Big Room Planning (Sl. 51) creates quarterly alignment "
                "cadence. This is the most mature operating model — it already operates in production for P&S teams."
            ),
            "refs": [("nWOW", "Sl. 5, 32, 34, 38"), ("PM Playbook", "Sl. 9, 76, 81, 85, 119"), ("PD Playbook", "Sl. 12")]},
        "digital_coe": {
            "what": (
                "The CoE defined a 3-tier governance model (Portfolio → Program → Delivery, Sl. 7) with an E2E flow "
                "that emphasizes Discovery upstream before any delivery commitment (Sl. 9). The unique contribution is "
                "the Strategic Initiatives Funnel with PRD (Light Requirements) and Lean Business Case as intake gates. "
                "Discovery and Delivery are treated as two separate Kanban boards connected by feedback loops (Sl. 48). "
                "The model tracks Cycle Time per phase and Discovery Success Rate as leading indicators (Sl. 41). "
                "This is conceptually sophisticated but less operationalized than P&S — more framework than playbook."
            ),
            "refs": [("CoE Playbook", "Sl. 4, 7, 9, 10, 36, 41, 47")]},
        "infra": {
            "what": (
                "Infrastructure took a different approach: instead of defining lifecycle phases, they focused on "
                "deployment methodology — Pilot → Expansion → Optimization (Sl. 10). The emphasis is on Team Topologies "
                "(4 team types, 3 interaction modes) as the structural foundation (Sl. 29-32). The rollout strategy is "
                "iterative: 1 pilot team first, learn, then expand to 7+ teams with coaching cycles. Less of an operating "
                "model per se, more of a change management and deployment playbook for introducing agile practices."
            ),
            "refs": [("Avangrid", "Sl. 10, 23, 29, 31, 33, 44")]},
        "dp_reuse": (
            "REUSE P&S's 4-phase lifecycle as structural backbone — it's the only one running in production. "
            "ADOPT CoE's Discovery upstream concept (separate Discovery and Delivery Kanban boards with feedback loops). "
            "ADOPT Infra's Pilot → Expansion → Optimization for the SAP rollout across OPCOs. "
            "The SAP model should combine: P&S structure + CoE discovery rigor + Infra deployment methodology."
        ),
    },
    {
        "name": "Value &\nPrioritization",
        "cov": "◕ | ◑ | ○",
        "ps_pt": {
            "what": (
                "P&S developed a concrete Portfolio Prioritization Calculator (Sl. 59) with a 1-10 weighted scoring "
                "survey covering: Business Value, User Value, Effort, Knowledge/Skill complexity, and Size (volume of features). "
                "The formula is PBI(x) = (BV+UX)/E where higher scores get priority (Sl. 77, 82). They defined explicit "
                "Business Value Criteria tied to Avangrid's strategic axes: Delight Users, Asset Health, Field Productivity, "
                "O&M Reduction, Regulatory, Reliability, Safety, Environmental (Sl. 72). The Effort Matrix uses S/M/L sizing "
                "mapped to feature count ranges (1-5=S, 5-10=M, 10+=L). The Demand Funnel (Sl. 42) starts with strategic "
                "initiatives from executives, flows through Shark Tank and Portfolio Committee for qualification."
            ),
            "refs": [("nWOW", "Sl. 35, 42, 59, 69, 71, 77, 85"), ("PD Playbook", "Sl. 82, 103"), ("BRP Q2", "Sl. 16")]},
        "digital_coe": {
            "what": (
                "The CoE uses an identical base formula PBI(x) = (BV+UX)/E (Sl. 84) but adds the HEVA cycle "
                "(Hypothesis → Experiment → Validation → Apprenticeship, Sl. 57) as a continuous discovery loop that feeds "
                "prioritization. The COORG method (Classify, Order, Organize) structures how backlog items are sorted "
                "after scoring (Sl. 78). Planning Poker maps complexity points (not time) using reference tasks (Sl. 79). "
                "The matrix prioritization uses a 2-axis approach: criteria on horizontal, PBIs on vertical (Sl. 83). "
                "The key difference from P&S: CoE prioritizes at backlog level (features/stories), while P&S prioritizes "
                "at portfolio level (initiatives/projects). Both are needed for SAP."
            ),
            "refs": [("CoE Playbook", "Sl. 83, 84"), ("Agility Playbook", "Sl. 51, 57, 78, 79, 82, 83")]},
        "infra": {"what": "", "refs": []},
        "dp_reuse": (
            "REUSE the PBI formula (BV+UX)/E — both P&S and CoE already use it, so there's organizational familiarity. "
            "ADOPT P&S's portfolio-level scoring survey (Sl. 59) but add SAP-specific criteria: cross-OPCO reusability, "
            "integration complexity, data sensitivity. ADOPT CoE's HEVA cycle for continuous backlog prioritization. "
            "Two-tier: portfolio scoring for initiative selection + COORG for backlog ordering."
        ),
    },
    {
        "name": "Demand Typology\n& Intake",
        "cov": "● | ◔ | ○",
        "ps_pt": {
            "what": (
                "P&S has the most complete demand management system. The Portfolio Framework & Demand Funnel (Sl. 35) "
                "maps how work enters the system from three sources: (1) Strategic initiatives from executives, "
                "(2) Product feedback and user field observations, (3) Regulatory/compliance mandates. The Lean Portfolio "
                "Funnel (Sl. 42) defines flow stages from inception to delivery. The Portfolio Kanban (Sl. 79) visualizes "
                "demand with swimlanes: High Risks & Blockers, High Value, Regulatory, OKRs & Strategic Goals. "
                "Classification uses S/M/L/XL sizing. The governance structure has defined forums per funnel stage — "
                "Shark Tank for initial triage (Sl. 43), Portfolio Committee for investment decisions (Sl. 103)."
            ),
            "refs": [("nWOW", "Sl. 35, 42"), ("PM Playbook", "Sl. 72, 84"), ("PD Playbook", "Sl. 88")]},
        "digital_coe": {
            "what": (
                "The CoE has a lighter intake model: the Agile Portfolio Funnel (Sl. 32) channels demand from executives "
                "into a Shark Tank → Think Tank → Biz Tank sequence. The Portfolio Management Team (Sl. 33) includes "
                "Enterprise Architect, VP Digital, and Enterprise Backlog as decision points. Lean Business Case (LBC) "
                "serves as the intake qualification gate. But there is no formalized demand typology or classification — "
                "the focus is more on governance gates than on categorizing the type of demand coming in."
            ),
            "refs": [("CoE Playbook", "Sl. 9, 10, 11, 32, 33")]},
        "infra": {"what": "", "refs": []},
        "dp_reuse": (
            "ADOPT P&S's demand funnel structure — the most mature. REUSE the 3-source model but add SAP-specific "
            "typology: Enhancement Request, Break-Fix/Incident, New Integration, Configuration Change, Data Request. "
            "REUSE CoE's Lean Business Case as the qualification gate. REUSE P&S's Portfolio Kanban swimlanes but adapt "
            "for SAP: by OPCO, by SAP module, by urgency. The key gap to fill: no area has a self-service catalog model."
        ),
    },
    {
        "name": "Roles &\nRACI",
        "cov": "● | ◔ | ◔",
        "ps_pt": {
            "what": (
                "P&S has a production-ready RACI framework. The PM Playbook defines RACI Matrix templates (Sl. 125-130) "
                "integrated with Confluence + Jira for a single source of truth on task ownership. The Project Manager Role "
                "(Sl. 105) is detailed with specific responsibilities per lifecycle phase. Stakeholder Mapping (Sl. 115-117) "
                "uses an engagement matrix to classify stakeholders by influence and interest. Resource Planning (Sl. 121-123) "
                "follows Atlassian best practices for capacity allocation. The nWOW defines crew roles (PO, SM, Dev) with "
                "process navigation that shows which role leads each phase (Sl. 34). The governance model distinguishes "
                "three Flight Levels with different roles at each (Sl. 76)."
            ),
            "refs": [("PM Playbook", "Sl. 115, 116, 126, 127, 128, 129"), ("nWOW", "Sl. 7")]},
        "digital_coe": {
            "what": (
                "The CoE defines roles implicitly through Scrum adoption. The Agility Playbook details the Scrum Framework "
                "(Sl. 18) and Sprint Events (Sl. 20) which imply PO/SM/Dev responsibilities per ceremony. The Agile Master "
                "Role (Sl. 94) is explicitly defined with a hygiene checklist for sprint health. The CoE Playbook shows "
                "governance levels (Sl. 7) which imply Portfolio Manager, Program Manager, and Team Lead responsibilities — "
                "but no formal RACI matrix exists. Roles are ceremony-based, not lifecycle-phase-based."
            ),
            "refs": [("Agility Playbook", "Sl. 20, 43, 94, 95, 98, 99"), ("NTTDATA", "Sl. 10")]},
        "infra": {
            "what": (
                "Infrastructure defines team structure through Team Topologies (Sl. 29-32): 4 team types (Stream-Aligned, "
                "Enabling, Platform, Complicated Subsystem) with 3 interaction modes (Collaboration, X-as-a-Service, "
                "Facilitating). The contribution here is not roles per se, but how teams relate to each other structurally — "
                "the model determines which teams should collaborate closely vs. which should provide self-service interfaces. "
                "This is directly relevant to how the SAP CoE should position itself (as a Platform Team providing X-as-a-Service)."
            ),
            "refs": [("Avangrid", "Sl. 25, 29, 30, 31, 32, 33")]},
        "dp_reuse": (
            "ADOPT P&S's RACI template structure + Stakeholder Mapping format — the only area with a formalized RACI. "
            "ADOPT CoE's Agile Master checklist for sprint hygiene. CRITICAL REUSE: Infra's Team Topologies directly "
            "informs SAP CoE positioning — SAP should be a Platform Team providing X-as-a-Service to consuming areas. "
            "The RACI simplifies: SAP owns the platform; consuming teams own their integration."
        ),
    },
    {
        "name": "Tool Ecosystem\n& Jira Config",
        "cov": "◕ | ● | ◑",
        "ps_pt": {
            "what": (
                "P&S uses Jira as the primary execution tool with Kanban Workflows defined (Sl. 88). The High Level "
                "Process in Jira (Sl. 96) maps the full flow from intake to delivery. Issue Type Hierarchy is defined "
                "(Sl. 92): Initiative → Epic → Story → Task. Board configurations support both Kanban and Scrum (Sl. 101-104). "
                "Dashboards for project health are Jira-native (Sl. 92, 142). Admin roles have defined permission levels "
                "(Sl. 97-100). The PM Playbook covers board creation step-by-step: filter customization, column mapping, "
                "epic creation workflows. The focus is on project-level execution tooling, not portfolio visualization."
            ),
            "refs": [("PM Playbook", "Sl. 82, 88, 91, 93, 95, 97, 100")]},
        "digital_coe": {
            "what": (
                "The CoE has the most comprehensive Jira documentation: the CX JIRA Playbook (41 pages) covers board design, "
                "filters, custom fields, workflow states, burndown charts end-to-end. The NextGen Jira Guide (Sl. 1-22) adds "
                "portfolio-level Jira configuration: how to set up Portfolio Management views (Sl. 8), connect roadmaps to "
                "backlogs, and use Jira for strategic alignment. Discovery and Delivery boards are configured as separate "
                "Kanban flows with connected statuses (Sl. 47-48). The Lean-Agile Team Governance Model defines how Blueprint, "
                "Backlog Management, and Metrics connect through Jira as the single ecosystem. Portfolio Kanban uses "
                "8 stage-gates with entry/exit criteria verified through Jira workflows (Sl. 26-27)."
            ),
            "refs": [("CX JIRA", "Full doc (41p)"), ("NextGen Jira", "Sl. 4, 6, 7, 8, 9, 21"), ("CoE Playbook", "Sl. 26-27, 47-48")]},
        "infra": {
            "what": (
                "TPO FLOW provides a hands-on JIRA Training Guide (Sl. 1-14): board creation, filter customization (Sl. 12), "
                "Epic creation workflows (Sl. 14). This is a training resource, not a configuration standard — it shows teams "
                "how to use Jira rather than defining how Jira should be configured. Useful as an onboarding template for SAP "
                "team members unfamiliar with Jira."
            ),
            "refs": [("TPO FLOW", "Sl. 1, 6, 10, 11, 12")]},
        "dp_reuse": (
            "ADOPT CoE's JIRA Playbook as the baseline config standard — it's the most comprehensive (41 pages + NextGen). "
            "REUSE CoE's Portfolio Kanban stage-gates (8 stages with entry/exit criteria) for SAP portfolio governance. "
            "REUSE P&S's issue type hierarchy (Initiative → Epic → Story → Task). "
            "REUSE TPO FLOW's training format for onboarding SAP team members. "
            "Add: API catalog and internal developer portal integration — gap across all 3 areas."
        ),
    },
    {
        "name": "Portfolio &\nPPMNow Concept",
        "cov": "◕ | ◕ | ◑",
        "ps_pt": {
            "what": (
                "The Portfolio Kanban (Sl. 79) is the centerpiece: swimlanes for High Risks, High Value, Regulatory, OKRs. "
                "The 3-tier governance (Portfolio → Program → Project) uses Flight Levels from Kanban at scale (Sl. 76). "
                "Big Room Planning (Sl. 51) creates quarterly alignment across all P&S teams — this is the forum where "
                "cross-team dependencies are surfaced and resolved. Portfolio Estimation Logic (Sl. 69-71) provides concrete "
                "sizing formulas. The governance model connects strategic goals to portfolio selection (Sl. 32). "
                "Process Navigation (Sl. 34) visually shows how all governance phases connect E2E."
            ),
            "refs": [("nWOW", "Sl. 32, 54, 78, 79"), ("PM Playbook", "Sl. 69, 76, 77, 78, 80")]},
        "digital_coe": {
            "what": (
                "The Portfolio Learn Cycle (Sl. 73) runs quarterly: 30 min reviewing past performance, 30 min planning future. "
                "The Portfolio Architecture (Sl. 102) defines 4 portfolio cycles per year with Strategy Design Workshops at "
                "each transition. Portfolio Metrics (Sl. 89, 93) are explicitly defined: Initiative Lead Time, ROI per Initiative, "
                "Lean Portfolio Management Radar, % Initiatives per Strategic Theme, Validated Business Models. "
                "PPM Practices (Sl. 100) include OKR, Roadmap, Pull System, Value Pyramid. The Strategic Initiatives "
                "Prioritization uses an Importance × Urgency Matrix (Sl. 80). The distinction: CoE's portfolio management "
                "is more metrics-driven and strategically aligned than P&S's which is more operationally focused."
            ),
            "refs": [("CoE Playbook", "Sl. 59, 60, 67, 72, 73, 80, 93, 100, 102")]},
        "infra": {
            "what": (
                "Infrastructure approaches portfolio at program level: Portfolio Roadmap (Sl. 38) visually manages E2E flow "
                "of projects across value stream teams. The Investment Model (Sl. 51-54) proposes per-unit-served financing — "
                "this is unique and directly relevant for SAP's OPCO-based cost allocation. Deployment governance includes "
                "phased rollout planning with clear milestone tracking."
            ),
            "refs": [("Avangrid", "Sl. 10, 15, 23, 37, 38, 46")]},
        "dp_reuse": (
            "ADOPT P&S's Portfolio Kanban format (clearest visualization) with adapted swimlanes for SAP: by OPCO, "
            "by SAP module, by request type. ADOPT CoE's Portfolio Learn Cycle (quarterly) and Portfolio Architecture "
            "(4 cycles/year) for SAP strategic alignment. REUSE CoE's portfolio metrics (ILt, RpI, VbM) — more sophisticated "
            "than P&S's. REUSE Infra's investment model for OPCO cost allocation. "
            "PPMNow should mirror the Portfolio Kanban with automated metrics from Jira."
        ),
    },
    {
        "name": "Baseline &\nDemand Map",
        "cov": "◔ | ◔ | ◔",
        "ps_pt": {
            "what": (
                "P&S has a structured discovery methodology (PD Playbook Sl. 22) for current-state assessment: contextual "
                "research, field observations, pain point identification, journey mapping, SIPOC diagrams. The Demand Funnel "
                "Prioritization (Sl. 42) provides a visual format for mapping demand flow. PM Playbook includes Project Plan "
                "formats with timeline views (Sl. 131-139). The Upstream-Downstream workflow (Sl. 83-84) shows how discovery "
                "feeds delivery through continuous interaction loops. This gives us a methodology for baseline assessment, "
                "but not a ready-made baseline — it would need to be executed for SAP."
            ),
            "refs": [("nWOW", "Sl. 42"), ("PD Playbook", "Sl. 13, 22, 51, 84"), ("PM Playbook", "Sl. 83, 84")]},
        "digital_coe": {
            "what": (
                "CoE's demand mapping focuses at portfolio level: Sl. 32 maps demand sources from executives through the "
                "funnel. The Portfolio Management Team (Sl. 33) defines who participates in demand triage: ideas flow through "
                "Shark Tank → Think Tank → Biz Tank gates. Portfolio Metrics (Sl. 79, 93) provide a framework for measuring "
                "current state. The quarterly Portfolio Results (Sl. 60) format could serve as a baseline template. "
                "However, there is no specific baseline assessment methodology — this is governance structure, not diagnosis."
            ),
            "refs": [("CoE Playbook", "Sl. 32, 33, 73, 79, 80, 102")]},
        "infra": {
            "what": (
                "The Project Book Template (Sl. 1-2) provides a structured E2E artifact: Discover → Solution → Plan → Build → "
                "Operate. The 'How to Use' guide (Sl. 2) defines: fill each section before moving to next. Phase 0 (Discover) "
                "covers Problem Statement, Inspiration, Ideation, Validation. The Blueprint Design method (Sl. 36) provides "
                "step-by-step project scoping. This template is directly reusable for documenting SAP baseline findings."
            ),
            "refs": [("Project Book", "Sl. 1, 2, 10"), ("Avangrid", "Sl. 16, 36")]},
        "dp_reuse": (
            "BUILD NEW using P&S's discovery methodology (research, interviews, journey maps) applied to SAP landscape. "
            "REUSE Infra's Project Book template to document findings per OPCO. "
            "REUSE CoE's Portfolio Results format for presenting the baseline. "
            "The demand map should be SAP-specific: which SAP capabilities are consumed, by whom, how often, and where "
            "the integration pain points are."
        ),
    },
    {
        "name": "Pilot\nExecution",
        "cov": "◔ | ◔ | ◑",
        "ps_pt": {
            "what": (
                "P&S defines two validation mechanisms: (1) the Validation Report (Sl. 68) with structured format: "
                "Problem/Hypothesis statement, what was tested, user feedback summary, key insights, feasibility assessment "
                "matrix, and GO/No-Go recommendation. (2) The Shark Tank Committee (Sl. 72) serves as a quarterly governance "
                "gate that evaluates viability, potential value, and risk profile — it grants or denies approval to proceed. "
                "Duration: 1-2 hours with Senior Leadership, Portfolio Manager, Product Managers. These are validation "
                "formats, not a pilot execution methodology."
            ),
            "refs": [("PD Playbook", "Sl. 3, 68, 70, 72"), ("nWOW", "Sl. 43")]},
        "digital_coe": {
            "what": (
                "The NTTDATA SEAL-SIX model (Sl. 9-10) provides a maturity assessment with scoring: 15-18 = Ready, "
                "10-14 = Launch pilot with guardrails, Below 10 = Focus on foundations first. The Minimum Viable SEAL Team "
                "(MVST) concept defines how to validate a team model with minimal disruption (Sl. 10). Agility Health Radars "
                "(Sl. 93) are visual radar charts for measuring team health bi-weekly or monthly — covering Value Generation, "
                "Experience, Transition indicators. The Sprint Report Canvas (Sl. 88-89) provides consolidated view of team "
                "baselines with comparison across productive cycles. The GO/No-Go gate process is defined for phase transitions."
            ),
            "refs": [("NTTDATA", "Sl. 5, 9, 10, 23, 24"), ("Agility Playbook", "Sl. 93")]},
        "infra": {
            "what": (
                "Infrastructure has the strongest pilot methodology: a 3-phase deployment approach — Pilot (1 team) → "
                "Expansion (7+ teams) → Optimization (Sl. 10, 34, 44). Each phase has defined activities: Pilot includes "
                "close support, team coaching, flow system impact metrics (Sl. 39). Expansion uses Sapiens Learning Cycles for "
                "Kanban implementation across teams (Sl. 43). Results measurement tracks Flow Metrics as the foundational "
                "definition of a successful Kanban team (Sl. 39). This is the most concrete and replicable pilot methodology — "
                "designed specifically for rolling out to multiple teams, which is exactly the SAP OPCO scenario."
            ),
            "refs": [("Avangrid", "Sl. 10, 21, 23, 34, 44, 45")]},
        "dp_reuse": (
            "ADOPT Infra's Pilot → Expansion → Optimization as the BRS pilot structure — best fit for multi-OPCO rollout. "
            "REUSE P&S's Validation Report format for pilot findings documentation. "
            "REUSE CoE's Agility Health Radars for measuring team health during pilot. "
            "REUSE NTTDATA's SEAL-SIX maturity scoring (15-18/10-14/Below 10) for readiness assessment. "
            "REUSE CoE's Sprint Report Canvas for tracking sprint-over-sprint progress."
        ),
    },
    {
        "name": "Ways of Working\nPlaybook",
        "cov": "● | ◕ | ○",
        "ps_pt": {
            "what": (
                "Two comprehensive playbooks totaling 264 slides. The Product Discovery Playbook (114 sl.) defines a "
                "full Design Thinking process (Sl. 6-7): Empathize → Define → Ideate → Prototype → Test, with concrete "
                "methods per step — contextual research, empathy maps, problem statements, Shark Tank validation, "
                "prototyping with Figma+Maze for user testing (Sl. 77). The playbook includes facilitation guides for "
                "Design Thinking Workshops (Sl. 76, 78). The Project Management Playbook (150 sl.) covers Kanban management, "
                "governance model, resource planning, RACI, contracts, KPIs — essentially a complete PM handbook. "
                "These playbooks are modular (chapters by lifecycle phase) and actively used by P&S teams."
            ),
            "refs": [("PD Playbook", "Full (114 sl.) — Sl. 3, 6, 7, 9, 63, 77"), ("PM Playbook", "Full (150 sl.)")]},
        "digital_coe": {
            "what": (
                "The Agility Playbook (111 sl.) is a back-to-basics agile guide. It covers Scrum Framework fundamentals "
                "(Sl. 18): Sprint Events (Sl. 20) with time-boxed 2-week cycles, Product Backlog Building using the PBB method "
                "(Sl. 51) including HEVA cycle, Persona mapping, Feature writing, Story writing templates. Sprint Report "
                "Canvas (Sl. 88-89) provides standardized reporting. Agility Health Radars (Sl. 93) enable visual team "
                "assessment. The JIRA Playbook (41p) acts as a companion tool guide. The approach is more methodological — "
                "teaching teams how to be agile — while P&S's playbooks are more operational — telling teams what to do."
            ),
            "refs": [("Agility Playbook", "Full (111 sl.) — Sl. 2, 20, 93"), ("CoE Playbook", "Sl. 88, 89"), ("CX JIRA", "Full (41p)")]},
        "infra": {"what": "", "refs": []},
        "dp_reuse": (
            "ADOPT P&S's modular playbook structure (chapters by lifecycle phase) as the template for the SAP WoW Playbook. "
            "REUSE P&S's PM Playbook Kanban management section as-is for SAP execution methodology. "
            "REUSE CoE's Agility Playbook for team-level agile practices (sprints, ceremonies, backlog management). "
            "COMBINE: P&S structure (what to do) + CoE methodology (how to be agile) = SAP WoW Playbook. "
            "Total reference: 375 slides of proven content to adapt, not build from scratch."
        ),
    },
    {
        "name": "KPI &\nReporting",
        "cov": "● | ◕ | ○",
        "ps_pt": {
            "what": (
                "P&S built a multi-level KPI ecosystem across 4 levels: Portfolio, Program, Product, Project (Sl. 83, 138-139). "
                "Each level has defined formulas: Throughput (# items delivered / time), Lead Time (days since WIP start), "
                "Velocity (Story Points), Scope Creep (# changes / planned × 100). Health Dashboards (Sl. 92, 142) use "
                "Jira-native widgets for real-time project health. Scorecards (Sl. 97, 145) aggregate KPIs with 2026 targets: "
                "Jira Usage Maturity, Product Backlog Maturity, Burndown Chart compliance, Team Ceremonies Consistency. "
                "The ecosystem explicitly defines when to use each KPI: initiation, reviews, root cause analysis, portfolio "
                "reviews (Sl. 83). This is the most actionable KPI framework — formulas, dashboards, and targets are defined."
            ),
            "refs": [("PM Playbook", "Sl. 137, 138, 140, 144, 145"), ("nWOW", "Sl. 83, 90, 93, 96, 97")]},
        "digital_coe": {
            "what": (
                "The CoE defines Portfolio Metrics (Sl. 89, 93) with specific KPIs: Initiative Lead Time, Lean Portfolio "
                "Management Radar, # Initiatives per Innovation Horizon, % per Strategic Theme, ROI per Initiative, "
                "Validated Business Models, Cost of Innovation. Team-of-Teams capacity measurement (Sl. 98) follows SAFe-style "
                "Agile Release Train assessment. Strategic Initiatives Prioritization (Sl. 80) uses Importance × Urgency Matrix. "
                "Portfolio Results are measured quarterly (Sl. 60) with operational results monthly. The Agile Teams Target "
                "Scorecard tracks maturity across Value Generation, Experience, and Transition indicators (Sl. 23). "
                "More strategically oriented than P&S — focus on value delivery and innovation metrics, not just execution health."
            ),
            "refs": [("CoE Playbook", "Sl. 26, 27, 73, 91, 93, 107"), ("Agility Playbook", "Sl. 87, 88, 89")]},
        "infra": {"what": "", "refs": []},
        "dp_reuse": (
            "ADOPT P&S's multi-level structure (Portfolio → Program → Team cascade) as the KPI architecture. "
            "REUSE P&S's concrete formulas (Throughput, Lead Time, Velocity, Scope Creep) — they're Jira-automated. "
            "REUSE CoE's strategic metrics (Initiative Lead Time, ROI per Initiative, Innovation Horizon alignment). "
            "REUSE P&S's dashboard layouts (Jira-native). REUSE CoE's scorecard approach for quarterly reviews. "
            "Add SAP-specific: Integration Lead Time, API Adoption Rate, Cross-OPCO Reusability Index."
        ),
    },
]


def set_cell_rich(cell, area_data, font_size=Pt(8)):
    """Populate cell with interpretive text + clickable file references."""
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
        run.font.name = brand.font_family
        return

    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = area_data["what"]
    run.font.size = font_size
    run.font.color.rgb = CHARCOAL
    run.font.name = brand.font_family

    for ref_key, ref_slides in area_data["refs"]:
        full_name = FILES.get(ref_key, ref_key)
        p = tf.add_paragraph()
        p.space_before = Pt(2)

        run_arrow = p.add_run()
        run_arrow.text = "→ "
        run_arrow.font.size = Pt(7)
        run_arrow.font.color.rgb = DARK_GRAY
        run_arrow.font.name = brand.font_family

        run_file = p.add_run()
        run_file.text = full_name
        run_file.font.size = Pt(7)
        run_file.font.color.rgb = BLUE_LINK
        run_file.font.underline = True
        run_file.font.name = brand.font_family
        try:
            add_hyperlink(run_file, file_uri(ref_key))
        except Exception:
            pass

        run_sl = p.add_run()
        run_sl.text = f"  |  {ref_slides}"
        run_sl.font.size = Pt(7)
        run_sl.font.bold = True
        run_sl.font.color.rgb = CHARCOAL
        run_sl.font.name = brand.font_family


# ═══════════════════════════════════════════
# BUILD DECK
# ═══════════════════════════════════════════
deck = DeckBuilder(brand)

# ── 1. Cover ──
deck.add_cover(
    title="Digital Products\nReview of Existing Assets",
    subtitle="SAP Digital Products — Wave 1\nInterpretive analysis of what P&S/P&T, Digital CoE & Infrastructure have already built",
    date="April 2026  |  Step 1: Review & Baseline",
)


# ═══════════════════════════════════════════
# 2-3. PaaP COMPARISON SLIDES (from original design)
# ═══════════════════════════════════════════
DARK_BG = RGBColor(0x1E, 0x29, 0x3B)
DARK_CARD = RGBColor(0x2D, 0x3B, 0x50)
RED_BG = RGBColor(0x5C, 0x1A, 0x1A)
GREEN_BORDER = RGBColor(0x10, 0xB9, 0x81)
TEAL_BG = RGBColor(0x0E, 0x3D, 0x3D)
CYAN_BORDER = RGBColor(0x06, 0xB6, 0xD4)
RED_BORDER = RGBColor(0xEF, 0x44, 0x44)
EMERALD = RGBColor(0x10, 0xB9, 0x81)
GRAY_TEXT = RGBColor(0x94, 0xA3, 0xB8)

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

# ── Slide 2: Bottleneck vs PaaP ──
slide = deck._blank()
deck._action_title(
    slide,
    "Applying a traditional Product Lifecycle to a foundational component team creates structural bottlenecks — the modern approach is Platform as a Product (PaaP)",
    "Managing SAP CoE through a PaaP lens shifts focus from intake governance to enablement, self-service, and reducing friction for consuming teams",
)

panel_y = Inches(1.45)
lx = Inches(0.8)
lw = Inches(5.5)

_make_box(slide, lx, panel_y, lw, Inches(5.1), DARK_BG, RGBColor(0x33, 0x41, 0x55))

ttl_left = slide.shapes.add_textbox(lx, panel_y + Inches(0.15), lw, Inches(0.4))
p = ttl_left.text_frame.paragraphs[0]
p.text = "Proposed Model (Bottleneck)"
p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = RED_BORDER
p.font.name = brand.font_family; p.alignment = PP_ALIGN.CENTER; p.font.italic = True

b1 = _make_box(slide, lx + Inches(0.6), panel_y + Inches(0.7), Inches(4.3), Inches(0.45), DARK_CARD)
_center_text(b1, "Digital Products / BRS OPCO", 10, GRAY_TEXT)
_arrow(slide, lx + Inches(2.5), panel_y + Inches(1.2))

b2 = _make_box(slide, lx + Inches(0.3), panel_y + Inches(1.6), Inches(4.9), Inches(0.9), RED_BG, RED_BORDER)
_center_text(b2, "Intake & Prioritization Funnel", 12, RGBColor(0xFC, 0xA5, 0xA5), True,
             "(ROI Scoring, Jira, PPMNow Governance)")
_arrow(slide, lx + Inches(2.5), panel_y + Inches(2.6))

b3 = _make_box(slide, lx + Inches(0.6), panel_y + Inches(3.0), Inches(4.3), Inches(0.45), DARK_CARD)
_center_text(b3, "SAP CoE (Slow Execution)", 10, GRAY_TEXT)

prob = slide.shapes.add_textbox(lx + Inches(0.3), panel_y + Inches(3.65), Inches(4.9), Inches(1.3))
tf = prob.text_frame; tf.word_wrap = True
for j, item in enumerate([
    "• More time in prioritization rituals than enabling",
    "• Each OPCO request goes through full funnel",
    "• SAP team is a ticket taker, not a platform",
    "• Scales linearly with headcount, not capabilities",
]):
    p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
    p.text = item; p.font.size = Pt(8); p.font.color.rgb = RGBColor(0xFC, 0xA5, 0xA5)
    p.font.name = brand.font_family; p.space_after = Pt(2)

rx = Inches(6.8); rw = Inches(5.5)
_make_box(slide, rx, panel_y, rw, Inches(5.1), DARK_BG, GREEN_BORDER)

badge = _make_box(slide, rx + Inches(3.9), panel_y, Inches(1.5), Inches(0.3), EMERALD)
_center_text(badge, "Recommended", 8, DARK_BG, True)

ttl_right = slide.shapes.add_textbox(rx, panel_y + Inches(0.15), rw, Inches(0.4))
p = ttl_right.text_frame.paragraphs[0]
p.text = "Platform as a Product (PaaP)"
p.font.size = Pt(16); p.font.bold = True; p.font.color.rgb = EMERALD
p.font.name = brand.font_family; p.alignment = PP_ALIGN.CENTER

consumers = ["BRS OPCO Pilot", "Digital Products", "Global PMO"]
cw = Inches(1.55)
for ci, c_label in enumerate(consumers):
    cx = rx + Inches(0.25) + ci * Inches(1.7)
    cb = _make_box(slide, cx, panel_y + Inches(0.7), cw, Inches(0.38), DARK_CARD)
    _center_text(cb, c_label, 8, GRAY_TEXT)
for ci in range(3):
    cx = rx + Inches(0.8) + ci * Inches(1.7)
    _arrow(slide, cx, panel_y + Inches(1.15))

pb = _make_box(slide, rx + Inches(0.2), panel_y + Inches(1.55), Inches(5.1), Inches(0.85), TEAL_BG, GREEN_BORDER)
_center_text(pb, "APIs, SDKs & Self-Service Portal", 12, RGBColor(0xA7, 0xF3, 0xD0), True,
             "(Standardized Contracts, Docs, Automation)")
_arrow(slide, rx + Inches(2.5), panel_y + Inches(2.45), "↑", GRAY_TEXT)

sb = _make_box(slide, rx + Inches(0.4), panel_y + Inches(2.85), Inches(4.7), Inches(0.45),
               RGBColor(0x0C, 0x31, 0x4D), CYAN_BORDER)
_center_text(sb, "SAP CoE (Platform Evolution)", 10, RGBColor(0x67, 0xE8, 0xF9))

adv = slide.shapes.add_textbox(rx + Inches(0.3), panel_y + Inches(3.5), Inches(4.9), Inches(1.3))
tf = adv.text_frame; tf.word_wrap = True
for j, item in enumerate([
    "• 75% of integrations via self-service (no intake)",
    "• Time to Hello World: 45 days → 2 days",
    "• Scales across OPCOs without adding headcount",
    "• SAP team focuses on engineering, not governance",
]):
    p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
    p.text = item; p.font.size = Pt(8); p.font.color.rgb = RGBColor(0xA7, 0xF3, 0xD0)
    p.font.name = brand.font_family; p.space_after = Pt(2)


# ── Slide 3: Enablement Roadmap ──
slide = deck._blank()
deck._action_title(
    slide,
    "SAP team as PaaS",
    "Instead of dedicating the 14-week program to configuring heavy PPMNow intake flows and ROI matrices, the engagement should pivot towards encapsulating SAP complexities and establishing the foundation for a scalable, developer-friendly internal platform.",
)

STEP_COLORS = [
    (RGBColor(0xF5, 0x9E, 0x0B), RGBColor(0x78, 0x35, 0x0F)),
    (RGBColor(0x06, 0xB6, 0xD4), RGBColor(0x16, 0x47, 0x57)),
    (RGBColor(0x10, 0xB9, 0x81), RGBColor(0x06, 0x4E, 0x3B)),
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
    y = Inches(1.55 + i * 1.55)

    card = _make_box(slide, Inches(0.8), y, Inches(11.5), Inches(1.35), DARK_BG, RGBColor(0x33, 0x41, 0x55))

    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.1), y + Inches(0.3), Inches(0.55), Inches(0.55))
    circle.fill.solid(); circle.fill.fore_color.rgb = dark_accent
    circle.line.color.rgb = accent_color; circle.line.width = Pt(2)
    cp = circle.text_frame.paragraphs[0]
    cp.text = num; cp.font.size = Pt(20); cp.font.bold = True
    cp.font.color.rgb = accent_color; cp.font.name = brand.font_family
    cp.alignment = PP_ALIGN.CENTER; circle.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    ttl = slide.shapes.add_textbox(Inches(1.9), y + Inches(0.15), Inches(9.5), Inches(0.4))
    p = ttl.text_frame.paragraphs[0]
    p.text = title; p.font.size = Pt(14); p.font.bold = True
    p.font.color.rgb = WHITE; p.font.name = brand.font_family

    dtxt = slide.shapes.add_textbox(Inches(1.9), y + Inches(0.55), Inches(9.8), Inches(0.7))
    tf = dtxt.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = desc; p.font.size = Pt(9); p.font.color.rgb = GRAY_TEXT; p.font.name = brand.font_family


# ═══════════════════════════════════════════
# 4. SUMMARY SLIDE — What each area actually built (interpretive)
# ═══════════════════════════════════════════
slide = deck._blank()
deck._action_title(
    slide,
    "P&S built execution playbooks, CoE built governance frameworks, Infrastructure built deployment methodology — SAP needs to combine all three into a unified operating model",
    "Each area solved different parts of the same problem — understanding what was actually built reveals what's directly reusable vs. what needs adaptation"
)

AREA_SUMMARIES = [
    {
        "name": "P&S / P&T (Networks IT)",
        "color": brand.primary,
        "bg": RGBColor(0xE8, 0xEF, 0xF7),
        "docs": "nWOW v1.2  |  PD Playbook v1.3  |  PM Playbook v1.0  |  BRP Q2 2026",
        "items": [
            ("Operating Model", "The most mature: 4-phase lifecycle running in production with 15+ sub-steps, phase gates, and defined artifacts per stage. Uses 3-tier Flight Level governance (Portfolio → Program → Project)."),
            ("Prioritization", "Concrete calculator: PBI(x) = (BV+UX)/E with 1-10 scoring survey. Business Value criteria tied to Avangrid strategic axes. Effort Matrix uses S/M/L mapped to feature count ranges."),
            ("Demand & Intake", "Most complete: single front door with 3 demand sources mapped. Portfolio Kanban with 4 swimlanes. Shark Tank + Portfolio Committee forums per funnel stage. Classification: Strategic/Operational/Regulatory."),
            ("Roles & RACI", "Production-ready RACI in Confluence+Jira. PM Role defined per lifecycle phase. Stakeholder engagement matrix. Resource planning with Atlassian-based capacity tracking."),
            ("Jira & KPIs", "Kanban/Scrum boards per project type. Multi-level KPI ecosystem with concrete formulas (Throughput, Lead Time, Velocity). Jira-native dashboards. 2026 scorecard targets defined."),
            ("Playbooks", "264 slides of proven structure: PD Playbook (Design Thinking E2E) + PM Playbook (governance, RACI, KPIs, contracts). Modular chapters by lifecycle phase — actively used."),
            ("Maturity", "HIGH — Running in production. Strongest in execution methodology and KPIs."),
        ],
    },
    {
        "name": "Digital CoE",
        "color": RGBColor(0x00, 0x70, 0xC0),
        "bg": RGBColor(0xE3, 0xF2, 0xFD),
        "docs": "CoE Playbook V.2.5  |  Agility Playbook V.1.7  |  NextGen Jira  |  JIRA Playbook (41p)",
        "items": [
            ("Operating Model", "Conceptually sophisticated: E2E with Discovery upstream + separate Delivery flow connected by feedback loops. 3-tier governance. Strategic Initiatives Funnel with Lean Business Case gate."),
            ("Prioritization", "Same PBI formula as P&S but adds HEVA cycle (continuous discovery loop) and COORG method for backlog ordering. Key difference: prioritizes at backlog level, not portfolio level."),
            ("Portfolio Mgmt", "Most metrics-driven: Portfolio Learn Cycle (quarterly), 4 portfolio cycles/year, Architecture with Strategy Workshops. Concrete KPIs: Initiative Lead Time, ROI per Initiative, Innovation Horizon."),
            ("Jira & Tools", "Most comprehensive Jira docs: 41-page playbook + NextGen guide. Portfolio-level Jira config. 8 stage-gate Portfolio Kanban with entry/exit criteria per stage."),
            ("Agile Practices", "111-slide back-to-basics guide: Scrum/Kanban fundamentals, PBB method, Sprint Report Canvas, Health Radars. More teaching-oriented — 'how to be agile' vs P&S's 'what to do'."),
            ("Maturity", "MEDIUM-HIGH — Strong frameworks but more conceptual than operational. Strongest in governance design and portfolio metrics."),
        ],
    },
    {
        "name": "Infrastructure (Avangrid)",
        "color": RGBColor(0x00, 0x96, 0x88),
        "bg": RGBColor(0xE0, 0xF2, 0xF1),
        "docs": "Agile Mgmt Model Avangrid V.3.1  |  Project Book Template  |  TPO FLOW JIRA Training",
        "items": [
            ("Deployment Model", "Unique approach: Pilot → Expansion → Optimization methodology. Designed for rolling out to multiple teams. 1 pilot team → learn → expand to 7+ teams with coaching cycles."),
            ("Team Structure", "Team Topologies framework: 4 types with 3 interaction modes. Directly informs how SAP CoE should position itself (as Platform Team providing X-as-a-Service)."),
            ("Investment", "Per-unit-served financing model — unique and directly relevant for SAP's OPCO cost allocation. Deployment governance with phased rollout planning."),
            ("Documentation", "Project Book template (Discover → Build → Operate) with structured phase tracking. Blueprint Design method for project scoping."),
            ("Jira", "Training-oriented: step-by-step guide for board creation, filters, Epics. More onboarding resource than configuration standard."),
            ("Maturity", "MEDIUM — Strongest in deployment strategy and team structure. Less defined in lifecycle, prioritization, and KPIs."),
        ],
    },
]

col_w = Inches(3.73)
gap = Inches(0.15)
start_x = Inches(0.8)
start_y = Inches(1.35)
col_h = Inches(5.5)

for i, area in enumerate(AREA_SUMMARIES):
    x = start_x + i * (col_w + gap)

    hdr = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, start_y, col_w, Inches(0.35))
    hdr.fill.solid()
    hdr.fill.fore_color.rgb = area["color"]
    hdr.line.fill.background()
    p = hdr.text_frame.paragraphs[0]
    p.text = area["name"]
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.font.name = brand.font_family
    p.alignment = PP_ALIGN.CENTER
    hdr.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

    docs_box = slide.shapes.add_textbox(x + Inches(0.08), start_y + Inches(0.38), col_w - Inches(0.16), Inches(0.32))
    dp = docs_box.text_frame.paragraphs[0]
    dp.text = area["docs"]
    dp.font.size = Pt(6)
    dp.font.color.rgb = DARK_GRAY
    dp.font.italic = True
    dp.font.name = brand.font_family
    docs_box.text_frame.word_wrap = True

    content_bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, start_y + Inches(0.72), col_w, col_h - Inches(0.72))
    content_bg.fill.solid()
    content_bg.fill.fore_color.rgb = area["bg"]
    content_bg.line.color.rgb = area["color"]
    content_bg.line.width = Pt(0.5)

    txt = slide.shapes.add_textbox(x + Inches(0.1), start_y + Inches(0.78), col_w - Inches(0.2), col_h - Inches(0.85))
    tf = txt.text_frame
    tf.word_wrap = True

    for j, (label, desc) in enumerate(area["items"]):
        if j == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
            p.space_before = Pt(3)

        run_lbl = p.add_run()
        run_lbl.text = f"{label}: "
        run_lbl.font.size = Pt(7)
        run_lbl.font.bold = True
        run_lbl.font.color.rgb = area["color"]
        run_lbl.font.name = brand.font_family

        run_desc = p.add_run()
        run_desc.text = desc
        run_desc.font.size = Pt(7)
        run_desc.font.color.rgb = CHARCOAL
        run_desc.font.name = brand.font_family
        p.space_after = Pt(1)


# ═══════════════════════════════════════════
# 3-5. DELIVERABLE TABLES (3 slides)
# ═══════════════════════════════════════════
batches = [
    (DELIVERABLES[0:4], "Existing Assets — Operating Model, Value Framework, Demand Intake, RACI (1/3)"),
    (DELIVERABLES[4:7], "Existing Assets — Tool Ecosystem, Portfolio & PPMNow, Baseline & Demand Map (2/3)"),
    (DELIVERABLES[7:10], "Existing Assets — Pilot Execution, Playbook, KPI & Reporting (3/3)"),
]

for batch, btitle in batches:
    slide = deck._blank()
    deck._action_title(slide, btitle)

    legend = slide.shapes.add_textbox(Inches(0.8), Inches(0.9), Inches(11.5), Inches(0.2))
    lp = legend.text_frame.paragraphs[0]
    lp.text = "● Complete  ◕ Well-defined  ◑ Partial  ◔ Referenced  ○ None          → = clickable link to source file"
    lp.font.size = Pt(7)
    lp.font.color.rgb = MID_GRAY
    lp.font.name = brand.font_family

    n_rows = len(batch) + 1
    n_cols = 5
    table_top = Inches(1.15)
    table_h = Inches(5.7)

    ts = slide.shapes.add_table(n_rows, n_cols,
                                Inches(0.15), table_top, Inches(13.0), table_h)
    table = ts.table

    col_widths = [Inches(1.25), Inches(3.6), Inches(3.2), Inches(2.2), Inches(2.75)]
    for i, w in enumerate(col_widths):
        table.columns[i].width = int(w)

    headers = ["Deliverable", "P&S / P&T", "Digital CoE", "Infrastructure", "Digital Products (Reuse)"]
    header_colors = [brand.primary, brand.primary, brand.primary, brand.primary, GREEN]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = header_colors[i]
        tf = cell.text_frame
        tf.word_wrap = True
        tf.clear()
        p = tf.paragraphs[0]
        p.text = h
        p.font.size = Pt(8)
        p.font.bold = True
        p.font.color.rgb = WHITE
        p.font.name = brand.font_family
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
        p.text = d["name"]
        p.font.size = Pt(9)
        p.font.bold = True
        p.font.color.rgb = CHARCOAL
        p.font.name = brand.font_family
        p2 = tf.add_paragraph()
        p2.text = d["cov"]
        p2.font.size = Pt(9)
        p2.font.color.rgb = GREEN
        p2.font.bold = True
        p2.font.name = brand.font_family
        cell.vertical_anchor = MSO_ANCHOR.TOP

        for c_idx, area_key in enumerate(["ps_pt", "digital_coe", "infra"]):
            cell = table.cell(r_idx + 1, c_idx + 1)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            set_cell_rich(cell, d[area_key], Pt(8))
            cell.vertical_anchor = MSO_ANCHOR.TOP

        cell = table.cell(r_idx + 1, 4)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0xE8, 0xF5, 0xE9) if r_idx % 2 == 0 else RGBColor(0xC8, 0xE6, 0xC9)
        tf = cell.text_frame
        tf.word_wrap = True
        tf.clear()
        p = tf.paragraphs[0]
        p.text = d.get("dp_reuse", "")
        p.font.size = Pt(7)
        p.font.color.rgb = CHARCOAL
        p.font.name = brand.font_family
        cell.vertical_anchor = MSO_ANCHOR.TOP


# ── 6. Closing ──
deck.add_closing(
    title="Prior Art Review",
    subtitle="SAP Digital Products — Wave 1",
)

# ═══════════════════════════════════════════
# SAVE
# ═══════════════════════════════════════════
output_path = Path("/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Digital Products - Prior Materials Review .pptx")
n_slides = deck.save(str(output_path))
print(f"Generated {n_slides} slides → {output_path}")
