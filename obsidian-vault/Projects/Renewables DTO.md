# Renewables DTO — Digital Transformation Office

#project #lean

> **Lean Kanban operations** — 53 KB chunks. 4-Pillar Model with Jira as single source of truth. Simplest model to replicate.

## Source Documents
| Document | Slides | Focus |
|----------|--------|-------|
| NextGen DTO Deliverable 1 — Final | 53 | Complete operating model — intake, prioritization, delivery, governance |

---

## Lifecycle — 4 Pillars Model

![[DTO_S5_4_Pillars.png]]

**What this shows:** The complete Lean Kanban 4-Pillar structure:
- **Pillar 1: [[Demand Intake]]** — standardized intake through Jira single channel
- **Pillar 2: [[Prioritization]]** — continuous ordering using BV × Urgency × Capacity
- **Pillar 3: Delivery** — pull-based execution with WIP limits enforced
- **Pillar 4: [[Governance]]** — cadences, DoR/DoD, flow monitoring, continuous improvement
- Events and best practices mapped per pillar

---

## Demand Intake — Jira Single Channel

![[DTO_S19_Jira_Single_Intake.png]]

**Jira Single Intake + DoR Gate:** All demand enters through ONE Jira channel regardless of source:
- Standardized ticket format with mandatory fields
- **Definition of Ready** as gate — request must meet DoR before entering backlog
- Triage process validates classification before acceptance
- Eliminates the "multiple entry points" problem that [[SAP Digital Products]] faces

![[DTO_S20_Demand_Classification.png]]

**Demand Classification — Work Types:** How requests are categorized once they enter:
- Enhancement, Bug Fix, New Feature, Configuration
- Mapped to different process paths and SLA expectations
- Connected to capacity allocation (Build 60% / BAU 25% / Upgrades 15%)

---

## Portfolio — Epic Kanban (Strategy to Execution)

![[DTO_S12_Epic_Kanban.png]]

**Epic Kanban Board — Level 1 (Portfolio):** Connects strategy to execution:
- **Columns:** Backlog → PMO → Project Prep → Ready → Design & Build → Done
- **Swimlanes:** Growth (Product), Power Trading, O&M, Canceled
- Evaluated by DTO Project Committee at Replenishment Meeting
- Ideas flow from strategy down to project-level execution

---

## Prioritization — 6 Core Processes

![[DTO_S21_6_Core_Prioritization.png]]

**6 Core Prioritization Processes:** Structured framework covering:
- Continuous ordering of the backlog
- BV × Urgency × Capacity as scoring dimensions
- Pre-estimation before commitment
- Links to [[Governance]] cadences for decision authority

![[DTO_S37_Continuous_Backlog.png]]

**Continuous Backlog Prioritization:** Pull-based model where teams continuously re-order instead of batch prioritization at fixed intervals.

---

## Governance — Facilitate Cadences

![[DTO_S50_Facilitate_Cadences.png]]

**Facilitate Cadences:** Process and artifacts for each governance ceremony:
- Daily Flow Review, Retrospective/Kaizen
- DTO Project Committee (replenishment)
- Mapped to workflow policies: WIP limits, flow monitoring, continuous improvement

---

## Roles — Per Kanban Stage

![[DTO_S17_Roles_per_Stage.png]]

**Roles per Kanban Stage:** Who owns what at each column of the board:
- PO responsible for backlog and prioritization
- Dev team pulls from Ready column
- QA validates before Done
- Clear handoff points with accountability markers

---

## Metrics — Flow Dashboard

![[DTO_S49_Flow_Metrics.png]]

**Flow Metrics Dashboard:** Operational dashboard tracking:
- **Lead Time:** total time from request to delivery
- **Cycle Time:** time in active work
- **Throughput:** items completed per period
- **WIP Aging:** items stuck longer than expected
- Visual dashboard layout for governance reviews

---

## Capacity — 60/25/15 Allocation

![[DTO_S39_Capacity_60_25_15.png]]

**Capacity Allocation Model:**
- **60% Build** — new features and strategic work
- **25% BAU** — business as usual, maintenance
- **15% Upgrades** — technical improvements, infrastructure
- Simple, enforceable, visible allocation that teams can track weekly

---

## Cross-Project Reuse

| What DTO built | Reusable for [[SAP Digital Products]]? |
|----------------|----------------------------------------|
| Single Jira intake (S19) | ✓ Pattern for single front door |
| Work type classification (S20) | ✓ Demand taxonomy template |
| Epic Kanban (S12) | ✓ Portfolio board config |
| 60/25/15 allocation (S39) | ✓ Simple capacity model |
| Flow Metrics (S49) | ✓ Dashboard baseline |
| 6 Core Prioritization (S21) | ✓ Process template |

## Related
- [[Digital CoE]] — shares feedback loops pattern
- [[P&S - P&T]] — more complex version of similar model
- Simplest model to pilot first for [[SAP Digital Products]]
