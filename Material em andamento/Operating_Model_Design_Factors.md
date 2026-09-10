# Operating Model — Design Factors

> **Propósito:** Framework de **fatores a considerar** no desenho do nosso operating model. **Não é o TOM** — é o conjunto de práticas, capabilities, roles, tools, governance e KPIs que devemos avaliar, priorizar e adaptar à realidade Avangrid Digital Products.
>
> **Fonte:** Target Operating Model.pptx (Materiais de referencia) + práticas SAP-flavored a serem incorporadas
>
> **Estrutura:** 4 stages × 5 dimensões (matriz 4×5)
>
> **Status:** Reference framework — input para o desenho do operating model
>
> **Data:** Abril 2026

---

## Estrutura Geral

| | DEMAND MANAGEMENT | PREPARATION & PLANNING | EXECUTION | DEPLOYMENT |
|---|---|---|---|---|
| **CAPABILITIES** | What demand looks like at intake | How demand becomes a roadmap | How work gets built | How releases land in production |
| **ROLES** | Who participates at intake | Who plans and prioritizes | Who delivers | Who releases & supports |
| **TOOLS** | Capture, track, communicate | Backlog, roadmap, capacity | Work mgmt, collab | Release, monitoring |
| **GOVERNANCE** | Forums, intake rules | Portfolio reviews, prioritization rules | Delivery reviews, coordination | Release approval, change control |
| **PERFORMANCE** | Volume, quality, triage | Capacity vs demand, predictability | Velocity, defects, WIP | Success rate, MTTR, stability |

---

## 01 · DEMAND MANAGEMENT

### Capabilities
- Demand intake channels (formal and informal)
- Demand registration and logging
- Channel consolidation level (single vs multiple entry points)
- Standard intake formats or templates
- Demand classification (type, size, urgency)
- Differentiation between run, change, regulatory, and innovation demand
- Business value definition
- Availability of acceptance criteria
- Technical feasibility input
- Triage and initial validation activities
- Demand routing and assignment
- Handling of urgent or exception requests
- Demand visibility and tracking
- Identification of duplicate or overlapping demand

### Roles
- Business requestors / stakeholders
- Product or demand managers
- Architects or technical reviewers

### Tools
- Demand capture tools
- Demand tracking and backlog tools
- Communication channels
- Documentation tools for capturing request details
- Reporting or visibility tools

### Governance
- Demand review forums
- Decision-making practices for accepting, rejecting, or redirecting demand
- Escalation mechanisms for urgent or conflicting requests
- Definition of intake guidelines or policies
- Alignment mechanisms with business stakeholders
- Oversight of demand inflow and backlog health

### Performance
- Demand volume and inflow trends
- Demand quality (completeness, clarity)
- Triage lead time / % of demand entering through formal channels
- % of urgent or bypassed requests
- Demand duplication or overlap indicators
- Visibility and traceability of demand status

### Extensions (TOM refinements)
- **Single front door** — Incidents · Improvements · Regulatory · Innovation · Technical Debt (5 demand types)
- **Demand shaping** — Structured problem–impact–value conversations; do NOT accept "solutions" directly
- **Classification by product type** — Run / Grow / Transform · Core SAP vs extensions / satellites
- **Lightweight but formal Discovery** — Problem statement · Users/personas · Hypothesis of value & risk
- **Value & prioritization framework** — Business value · Risk · Regulatory urgency · Cost of Delay (CoD)

---

## 02 · PREPARATION & PLANNING

### Capabilities
- Demand consolidation and portfolio structuring
- Demand prioritization mechanisms
- Use of value, risk, and effort considerations
- Portfolio backlog management
- Roadmap definition (short / mid / long term)
- Planning cycles (e.g., monthly, quarterly)
- Effort estimation practices (high-level and detailed)
- Capacity planning and forecasting
- Alignment between demand and available capacity
- Work allocation to teams / squads
- Handling of unplanned or ad-hoc demand
- Grouping and structuring of work (epics, features, initiatives)

### Roles
- Product managers / product owners
- Technical leads / architects
- Business stakeholders

### Tools
- Backlog and portfolio management tools
- Roadmapping and planning tools
- Estimation and capacity planning tools
- Reporting and visualization tools

### Governance
- Portfolio review forums
- Prioritization and planning decision-making practices
- Alignment mechanisms between business and technology
- Definition of prioritization criteria and policies
- Oversight of demand vs capacity balance
- Decision practices for deferring or rejecting demand
- Cross-area coordination mechanisms (if multiple OPCOs / domains)

### Performance
- Demand vs capacity visibility
- Prioritization cycle time
- % of demand planned vs unplanned
- Alignment between planned and committed work
- Planning predictability
- Portfolio balance (e.g., run vs change vs regulatory)
- Visibility of roadmap and planning outcomes

### Extensions (TOM refinements)

**A · Product Definition**
- Products ≠ Applications
- Example: "Order-to-Cash SAP" ≠ "SAP SD" (the product spans modules)
- Product vision & outcome definition (outcomes ≠ deliverables)
- Roadmapping continuous · rolling wave (3 / 6 / 12 months)
- Differentiate commitments vs hypotheses

**B · Product Design & Architecture Alignment** *(critical in SAP to avoid Product ≠ Architecture drift)*
- **Dual track**:
  - Discovery (what & why)
  - Delivery (how & when)
- Business process design — E2E processes (O2C, P2P, R2R, M2C)
- Non-functional requirements — Performance · Security · Compliance · Scalability

---

## 03 · EXECUTION

### Capabilities
- Intake of work from prioritized backlog
- Work breakdown (epics, stories, tasks) · Sprint / iteration management
- Work prioritization at team level
- Definition of Ready (or equivalent readiness practices)
- Definition of Done (or completion criteria)
- Development and build practices
- Testing practices (unit, integration, functional, UAT)
- Defect management
- Dependency management (intra- and inter-team)
- Cross-team coordination
- Work progress tracking and visibility
- Documentation of functional and technical aspects

### Roles
- Delivery teams
- Technical leads
- Scrum masters / delivery managers
- Product owners / business representatives

### Tools
- Work management tools
- Collaboration and communication tools
- Documentation tools

### Governance
- Delivery review forums (e.g., sprint reviews, demos)
- Issue and risk escalation practices
- Oversight of delivery progress and performance
- Definition of delivery standards and guidelines
- Coordination mechanisms across teams or domains
- Alignment with business during execution

### Performance
- Velocity / throughput
- Lead time and cycle time
- Predictability (planned vs delivered)
- Defect rates and quality indicators
- Work in progress (WIP) levels
- Dependency-related delays
- Visibility of execution progress

### Extensions (TOM refinements)
- **Stable product teams** — PO · Tech Lead · delivery team (per product, not per app)
- **Backlog management per product** (not per application)
- **Flow-based delivery** (when) — Scrum · Kanban · Hybrid
- **Dependency management** — Between SAP products · with vendors / integrators

---

## 04 · DEPLOYMENT

### Capabilities
- Release scheduling and planning
- Release coordination across teams
- Change management practices
- Deployment execution (manual or automated) · Environment management (dev, test, prod)
- Rollback and recovery capabilities
- Validation and sign-off processes
- Post-deployment verification
- Hypercare and stabilization support
- Transition to operations and support teams

### Roles
- Release managers
- Operations / support teams
- Delivery teams (as contributors to release)
- Business stakeholders (for validation / sign-off)

### Tools
- Release and deployment tools
- Monitoring and alerting tools
- Incident and support management tools

### Governance
- Release and change approval forums (e.g., CAB or equivalents)
- Decision-making practices for release readiness
- Oversight of deployment risks and impacts
- Definition of deployment and change policies
- Coordination between delivery and operations
- Control mechanisms for production changes

### Performance
- Deployment success rate
- Incidents and defects post-release
- Time to recover from failures
- Release frequency
- Stability of deployed solutions
- Effectiveness of hypercare period
- Visibility of release outcomes

### Extensions (TOM refinements)
- **Hypercare and feedback loop** — Measure impact after release
- **Deployment = Measure & Learn** (cultural principle)

---

## Mapping — Como aplicar à Avangrid Digital Products

| Stage | Avangrid AS-IS context | TOM target |
|---|---|---|
| **Demand** | 6+ canais, sem front door único; routing por funding source (BRS/PMO/IT) | Single front door · 5 demand types · Run/Grow/Transform classification · CoD-based prioritization |
| **Preparation** | PDD para large; sem rolling wave; Product ≠ App não articulado | Product definition (M2C, O2C, P2P, R2R) · Rolling wave 3/6/12 · Dual-track Discovery+Delivery |
| **Execution** | Squads em P&T (mature) · resto fragmentado | Stable product teams · backlog by product · flow-based · dependency mgmt |
| **Deployment** | Hypercare ad-hoc; sem feedback loop estruturado | Hypercare formal · Measure & Learn · post-release KPIs |

---

## Conexões com os 8 Guiding Principles

| Princípio | Stage primário | Capability TOM correspondente |
|---|---|---|
| 01 Single Front Door | Demand | Demand intake channels + Channel consolidation |
| 02 Value-Based Prioritization | Demand + Preparation | Value/risk/CoD criteria · prioritization mechanisms |
| 03 Capacity-Aware Assignment | Preparation | Capacity planning · alignment demand vs capacity |
| 04 Standardized Execution Platforms | Execution | Work management tools (Jira) + Roadmapping (PPM Now) |
| 05 Integrated Toolchain | Cross-stage | Tool layer all stages |
| 06 Automated Reporting | Performance | Reporting & visualization tools across stages |
| 07 Defined RACI per Lifecycle Phase | All stages | Roles row across all 4 stages |
| 08 Hybrid Delivery Model | Execution | Flow-based delivery: Scrum/Kanban/Hybrid |
