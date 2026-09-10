# P&S / P&T — Planning & Scheduling / Planning & Tracking

#project #mature

> **Most mature project** — 394 KB chunks from 4 deliverables. 4-phase lifecycle in production with 15+ sub-steps, 3-tier Flight Level governance, and the PBI prioritization formula.

## Source Documents
| Document | Slides | Focus |
|----------|--------|-------|
| nWOW Operating Model v1.2 | 97 | Master playbook — lifecycle, governance, portfolio, demand |
| Product Discovery Playbook | 114 | Design thinking E2E, validation, Figma+Maze |
| Project Management Playbook | 150 | Kanban, governance, RACI, contracts, KPIs |
| BRP Q2 2026 | 26 | Quarterly business review, progress report |

---

## Lifecycle — Future State Framework (Compass Model)

![[PS_S37_Future_State_Framework.png]]

**What this shows:** The end-to-end Compass Portfolio Management Model — how ideas flow from inception to delivery:
- **Left:** Demand Management system (refinement of ideas and business cases) with **Shark Tank** committee as first filter
- **Center:** Portfolio Backlog feeding into Programs A–F, managed by Portfolio Manager
- **Right:** Budget Management (CAPEX/OPEX) and Kanban Teams (execution per value stream using Team Topologies)
- **Bottom 5 phases:** Phase 0 Discovery → Phase 1 Materialization → Phase 2 Creation → Phase 3 Ongoing → Phase 4 Maintenance
- **Full horizontal bar:** Opportunity Identification → Ideation → Solutioning → Definition → Planning → Execution → Deliver → Maintenance → Knowledge Management

![[PS_S38_E2E_Management_Process.png]]

**E2E Management Process** — the detailed 5-phase breakdown showing how value is identified and created, with sub-steps and gates at each transition.

---

## Demand Intake — Lean Portfolio Funnel

![[PS_S42_Lean_Portfolio_Funnel.png]]

**What this shows:** The Demand Funnel Prioritization process:
- **Top of funnel (open):** 3 sources — Strategic Programs, User Problems, Business Improvement Opportunities (BIO) per OpCo
- **Shark Tank filter:** Ideas submitted first-come first-served, approved by Portfolio ID Committee (Gate 1)
- **Bottom of funnel (focused):** Approval → Funding → Implementation Planning → Execution & Delivery
- **Right side:** Portfolio Backlog managed by Portfolio Manager
- Enterprise Backlog practice: "Capacity is only attainable if the system moves from push to pull"

![[PS_S35_Portfolio_Demand_Funnel.png]]

**Portfolio Framework & Demand Management Funnel** — combined view showing the full funnel with stages and decision points.

---

## Governance — Forums v2 TO-BE Diagram

![[PS_S55_Governance_Forums_TOBE.png]]

**What this shows:** The definitive governance architecture:
- **Vertical axis:** "Doing the Right Things" (strategic) ↔ "Getting Things Done" (tactical)
- **Horizontal:** "Doing Things Better" (improvement)
- **7 forums mapped:** Research Interviews → Discovery Workshops → Design Thinking → Shark Tank Committee → Portfolio Committee → Big Room Planning → Project Planning
- **Cadences below:** Portfolio Sync, Program Syncs, Program Reviews, Changes Sync, Daily Sync, Iteration Plans & Review, Weekly Syncs, Project Releases
- **Continuous process:** Backlog Replenishment & Refinement across the bottom
- **Information flow:** bidirectional Info ↕ Change

![[PS_S53_Governance_Cadences.png]]

**Governance Cadences** — implementation details for all 8 cadence types: frequency, participants, artifacts, and purpose.

---

## Prioritization — BV Criteria + PBI Formula

![[PS_S72_BV_Criteria_Matrix.png]]

**Business Value Criteria Scoring Matrix:**
- **Business Value (Utilities Context):** contribution to "safe, reliable, compliant, and efficient utility operations"
- **Customer / Field Value:** improvement to day-to-day work for crews, field designers, inspectors, back-office
- **Delivery Effort:** IT complexity — systems, integrations, testing, vendor dependencies
- **Scale:** 1-3 (Very Low) → 6 (Medium) → 8 (Very High) → 9-10 (Critical/Transformational)

![[PS_S77_PBI_Formula.png]]

**PBI Formula:** `H(x) = (BV + UX) / E`
- BV = Business Value, UX = User Experience, E = Effort
- Higher = higher priority. Applied at portfolio AND team backlog levels.

![[PS_S59_Prioritization_Criterias.png]]

**Portfolio Prioritization Criterias v1.0** — the full survey instrument defining how each criterion is measured and weighted.

---

## Portfolio — Kanban Swimlanes (Level 1)

![[PS_S79_Portfolio_Kanban_L1.png]]

**Portfolio Kanban Board — Level 1** (highest tier of 3-tier Flight Level):
- **Columns:** Portfolio Backlog → Discovery → Shark Tank → Portfolio Committee → Sprint → Portfolio Backlog → Ongoing Project → Releases → Done
- **Swimlanes:** Program Backlog, OKR's & Strategic Goals
- **Roles:** Portfolio Manager, Program Manager (avatars at each stage)
- **Forums marked:** Flight Rises & Blockers, Regulatory, Big Room Planning (quarterly)

---

## Roles — Crews Work Model

![[PS_S17_Roles_Introduction.png]]

**Roles Introduction:**
- **P&S:** Portfolio Manager, Program Manager, Knowledge Manager, Process Improvement, UX Designer, QA
- **Insights:** PM oversees multiple programs, Portfolio Manager handles budgets, Knowledge Manager builds org ecosystem
- **IT Networks:** Technical Lead evolves to Solution Architect, preserving critical knowledge

![[PS_S26_Crews_Work_Model.png]]

**Crews Model:** Stable cross-functional teams structured around Capabilities, Release processes, and value streams. O2/Circle model for self-organizing structures.

---

## Metrics & KPIs

![[PS_S83_KPI_Ecosystem.png]]

**KPI Ecosystem — "When to Use":** Maps each KPI to purpose, frequency, and governance level. Categories: Flow Metrics (Lead Time, Cycle Time, Throughput, WIP Aging), Project Health, Portfolio Health.

![[PS_S90_Project_KPIs.png]]

**Project-Level KPIs:** Specific thresholds (green/yellow/red) with aging alert mechanism — triggers when items exceed expected cycle time.

---

## Capacity

![[PS_S63_Capacity_Logic.png]]

**Capacity Logic v1.0:**
- **New Project RFP:** resource and capacity predictability through formal planning
- **Stable Teams:** pull systems where capacity is managed within existing team allocation
- Two theories: resource/capacity predictability vs pull systems

---

## Cross-Project Reuse

| What P&S built | Reusable for [[SAP Digital Products]]? |
|----------------|----------------------------------------|
| Compass Framework (S37) | ✓ Adapt the 5-phase lifecycle |
| Demand Funnel (S42) | ✓ Single front door pattern |
| PBI Formula (S77) | ✓ Ready to use as-is |
| BV Criteria (S72) | ✓ Customize scoring for SAP |
| Governance Forums (S55) | ✓ Map to existing DP meetings |
| Portfolio Kanban (S79) | ✓ Configure in Jira |
| KPI Ecosystem (S83) | ✓ Reuse baseline set |

## Related
- [[Digital CoE]] — shares PBI formula, governance model
- [[Renewables DTO]] — adopted similar Kanban approach
- [[Infrastructure]] — Team Topologies complements Crews model
