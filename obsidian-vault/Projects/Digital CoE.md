# Digital CoE — Digital Center of Excellence

#project #governance

> **Governance framework specialist** — 298 KB chunks. Discovery-Delivery Dual Board + 3-Tier Model + 8-stage Portfolio Kanban.

## Source Documents
| Document | Slides | Focus |
|----------|--------|-------|
| Operating Model Playbook [12] | 108 | Master framework — governance, portfolio, demand, metrics |
| Agility Playbook [17] | ~100 | Agile practices, ceremonies, roles |
| SEAL-SIX Org Design | ~30 | Product organization structure, stable teams |

---

## Governance — 3-Tier Model

![[DCoE_S7_3_Tier_Governance.png]]

**What this shows:** Three-tier digital governance pyramid:
- **Strategic tier:** Portfolio direction, investment decisions, OKRs — quarterly/monthly cadence
- **Operational tier:** Prioritization, triage, cross-team coordination — bi-weekly/weekly
- **Tactical tier:** Execution, blockers, team-level decisions — daily/weekly
- Each tier connects via defined feedback loops and escalation paths
- Maps to specific forums with named cadences

![[DCoE_S65_Feedback_Loops_Cadences.png]]

**Feedback Loops with Cadences:** How information flows between the 3 tiers — bidirectional arrows showing escalation (up) and direction (down), with specific meeting cadences at each connection point.

---

## Lifecycle — E2E Operating Model

![[DCoE_S10_E2E_Discovery_Upstream.png]]

**E2E Operating Model — Discovery Upstream:** Shows how work is prioritized and flows through the upstream discovery process before entering delivery. Strategic Initiative Funnel connected by Feedback Loops with Lean Business Case.

![[DCoE_S11_E2E_ROI_Funnel.png]]

**E2E Operating Model — ROI Lens:** Same flow with ROI analysis overlay — work flow from ideation through delivery with value measurement at each stage. Capacity and Time-to-Market metrics embedded.

---

## Demand Intake — Portfolio Funnel + Classes of Service

![[DCoE_S32_Agile_Portfolio_Funnel.png]]

**Agile Portfolio Funnel:** Enterprise demand flowing from strategic initiatives into the portfolio. Shark Tank / Business Case evaluation → Portfolio Backlog → staged execution. Shows how multiple demand sources converge.

![[DCoE_S105_Classes_of_Service.png]]

**Classes of Service & Demand Management:** Distribution by impact class:
- **Expedite:** Emergency, bypass normal flow, immediate capacity allocation
- **Fixed Date:** Regulatory/compliance deadlines, non-negotiable schedule
- **Standard:** Normal business demand, follows regular cadence
- **Intangible:** Tech debt, enablers, infrastructure — no direct business ask but essential
- Each class has different SLA expectations and WIP policies

---

## Portfolio — Kanban Board + Management Team

![[DCoE_S27_Portfolio_Kanban_Board.png]]

**Portfolio Kanban Board:** 8-stage Kanban configured in Jira with WIP limits per column, Definition of Ready/Done at each gate, and swimlanes by business area.

![[DCoE_S25_Stage_Gate_Governance.png]]

**Stage-Gate Governance:** Detailed policies at each gate — what criteria must be met to move between stages, who approves, and what artifacts are required.

![[DCoE_S33_Portfolio_Mgmt_Team.png]]

**Portfolio Management Team Structure:** Who sits on the portfolio team, their responsibilities, and how they interact with product teams and stakeholders.

---

## Prioritization — PBI Formula (CX CoE Adaptation)

![[DCoE_S84_PBI_Formula.png]]

**CX CoE Prioritization Formula:** Same PBI(x) = (BV+UX)/E formula as [[P&S - P&T]], adapted for CX context:
- Criteria-based methodology with structured scoring table
- Applied to Product Backlog Items at team level
- Links to Stage-Gate governance at portfolio level

---

## Roles — Product Organization

![[DCoE_S13_Product_Owners.png]]

**Product Owners — Team Named:** Role descriptions for named POs within the CX Digital organization. Moderate detail in the Playbook — richer content in the SEAL-SIX Org Design deck (Stable Teams, Product Org Structure).

---

## Metrics — KPIs & Team of Teams

![[DCoE_S24_KPI_Strategic_Drivers.png]]

**CoE KPI as Strategic Drivers:** Maps KPIs to strategic objectives of the NextGen Digital CoE. Shows which metrics drive which business outcomes.

![[DCoE_S94_Team_of_Teams_Metrics.png]]

**Team of Teams Metrics Framework:** Aggregated metrics at the ART (Agile Release Train) level — flow metrics, agility metrics, and business results categories.

---

## Capacity — Team of Teams

![[DCoE_S97_Team_Capacity.png]]

**Team Capacity in Complex Work:** Framework for measuring and managing capacity in agile teams — definitions, measurement approaches, and constraints.

---

## Cross-Project Reuse

| What CoE built | Reusable for [[SAP Digital Products]]? |
|----------------|----------------------------------------|
| 3-Tier Governance (S7) | ✓ Template for forum structure |
| Stage-Gate Kanban (S25) | ✓ Jira board configuration |
| Classes of Service (S105) | ✓ Demand classification model |
| PBI Formula (S84) | ✓ Same as P&S, ready to use |
| KPI Strategic Drivers (S24) | ✓ Map to DP objectives |
| Feedback Loops (S65) | ✓ Cadence structure template |

## Related
- [[P&S - P&T]] — shares PBI formula, governance model
- [[Renewables DTO]] — similar pillar structure, simpler execution
