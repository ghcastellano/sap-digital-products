# Miro Post-its — Digital Products TOM (Unfix model)

Each section below = one cluster of post-its for the Miro board.
Suggested colors match the diagram (`DigitalProducts_Unfix_Squads.png`).
Format per post-it: **title** (bold) + 1-3 short body lines.

---

## 1) GOVERNANCE CREW (peach)

**GOVERNANCE CREW**
The management team of the Base. Sets vision, strategy, and governance constraints for Digital Products.

**Vitor Cruz**
VP IT Business Solutions Delivery
Sponsor of the Base

**Carlos Garbin**
Director — IT Business Apps
~26 FTE · 4 sub-teams

**Kumar Jaladanki**
Director — Customer Services Apps
~24 FTE · 5 sub-teams · 2 vacancies

**Heather LaJoie**
Director — Product Management (business side)
~2 FTE · cross-base bridge to Digital CoE

**Travis (TBC)**
PMO Lead — Large Projects
Owns PDD intake today

**Portfolio Manager**
TBD — NEW role
Capacity allocation across streams

---

## 2) VALUE STREAM CREWS — Customer Services Apps (purple, CS Apps blue badge)

> **IMPORTANT — Vendor reality:** Each crew has an Avangrid core (product ownership, architecture, SMEs) + Vendor extension (the people who actually build). Vendor model decision pending: **Option A** (stable vendor team per crew, like Smart Metering today) vs **Option B** (vendor pool consumed ad-hoc, like Sarath's Corporate Apps today). KB confirms both patterns exist; TO-BE choice is open.

---

**VALUE STREAM CREW · M2C — Service Lifecycle**
Branch: CS Apps (Kumar) · split from M2C because of size
Outcome: Customer onboarding · activation · account changes · disconnection
**Avangrid core (~3-4 FTE):**
· Product Manager (NEW)
· Iram Chowdhury (today: PJM/Champion, candidate to re-skill as PM)
· Vincent Lam (Sr.M. Device & Metering)
· 1-2 SMEs / BA
**Vendor extension (variable, peak ~6-12):**
· Smart Metering vendor (100% delivery — requirements + code)
· Accenture (cross-stream support)
· SWF / Software Factory (build capacity)
**Current applications (KB-confirmed apps):**
· SAP Metering — meter data and device management *(Slide 41 vrevisada · "Extended Application Portfolio")*
· MV90 — meter data acquisition system (Itron) *(KB: AGR-NET-2025-CS-19-MV90 / MV90 Upgrade UIL initiative implies app exists)*
· Vendor-managed Smart Metering platform — 100% vendor-delivered *(Slide 24 vrevisada)*

**Active initiatives (KB · separate from apps):**
· MV90 Upgrade UIL · AGR-NET-2025-CS-19-MV90-NY
*Source: APR 2026 IT Avangrid Project Status Reports · LARGE DEMAND PROJECT PORTFOLIO LANDSCAPE*

> Gap: KB does not list specific Smart Metering vendor stack components (vendor-internal). Confirm with Iram + Vincent.

---

**VALUE STREAM CREW · M2C — Revenue Operations**
Branch: CS Apps (Kumar) · split from M2C because of size
Outcome: Accurate billing · low-friction payment · fair collections · revenue recovery
**Avangrid core (~4-6 FTE):**
· Product Manager (NEW)
· Ivelisse Velazquez (Sr.M. FICA/Billing)
· 1-2 BA / SME (FICA + Collections)
· 1 Tech Lead (FICA architecture)
**Vendor extension (variable):**
· Accenture (FICA, ECC support)
· SWF (build capacity)
· SOVOS partner (tax)
· Precisely / Pitney Bowes (mailing)
**Current applications (KB-confirmed apps):**
· SAP FICA — billing engine + Financial Contract Accounting *(Slide 41 vrevisada)*
· SAP ECC — core financial / operational back-office (GL · AR · AP · invoicing) *(Slide 41)*
· StreamServe — billing letters · invoice printing · simulations · archives · EBPP *(Slide 41)*
· SOVOS Taxware — tax calculation and compliance *(Slide 41)*
· Precisely / Pitney Bowes — address validation · document processing · mailing services *(Slide 41)*

**Active initiatives (KB · separate from apps):**
· Distributed Generation SAP Billing Automations Program (NY) · EBD Certification Automation (NY)
· LIDR — UIL Low Income Discount Rate (UIL) · Hardship Communications Track (UIL)
· Advanced Termination Notice PA 23-199 · S4 HANA Implementation (cross-stream)
*Source: APR 2026 IT Avangrid Project Status Reports · LARGE DEMAND PROJECT PORTFOLIO LANDSCAPE*

---

**VALUE STREAM CREW · Customer Care & CCC**
Branch: CS Apps (Kumar) · ~4-6 FTE Avangrid (after vacant hire)
Outcome: Issue resolution + self-service + voice of customer
**Avangrid core (~2-4 FTE):**
· Product Manager (NEW · depends on vacant Sr.M. CRM/CCC hire)
· Heather LaJoie (cross-base, business side)
· 1-2 BA / SME
**Vendor extension (variable):**
· Accenture (CRM, contact center)
· HCL / Nuance (voice/IVR)
· Blue Prism partner (RPA)
· Medallia (feedback platform)
**Current applications (KB-confirmed apps only):**
· SAP CRM (Front Office) — manages customer interactions, service requests, accounts
· Medallia VHT — customer feedback platform
· Blue Prism — RPA Platform
· Nuance (HCL · Nuance) — voice / IVR / customer service support partner
*Source: Slide 41 vrevisada — "Extended Application Portfolio · All Applications in Scope"*

**Active initiatives on these apps (KB · separate from apps):**
· New IVR and Voice Solution · Customer Service Generative AI · RPA IPA Customer Service
· Customer Journey Enhancements Wave 3 · Digital Customer 360 Analytics · Mobile Redesign
· High Bill High Usage Redesign · Energy Manager Enhancements · Digital Communications
*Source: APR 2026 IT Avangrid Project Status Reports · LARGE DEMAND PROJECT PORTFOLIO LANDSCAPE — these are PROJECTS on the apps above, not apps themselves*

> Note: "Customer self-service portal" not nomeado app específico em KB — inferência. Confirmar com Heather LaJoie / Customer Service Products business side qual é o portal real (SAP CRM Web Channel? custom portal? ambos?).

---

## 3) VALUE STREAM CREWS — IT Business Apps (purple, IT Apps orange badge)

> Refined against the IT Networks Application Assessment (46 apps · 7 functional domains)
> *Source: `020326_Portfolio_Landscape_Analysis_Avangrid_V0.1.pptx` Slide 23*

---

**VALUE STREAM CREW · Field Operations**
Branch: IT Apps (Carlos)
Outcome: Field worker / inspector / lineman productivity — fewer rework loops, offline capability, faster storm response
Functional domains: Storm & Workforce Mgmt (4 apps) + Field Data Capture & Maintenance Execution (subset of Maintenance domain)
**Avangrid core (~3-4 FTE):**
· Product Manager (NEW)
· Andrew McGibbon (today: Champion E&G)
· Bruno Moreto (Sr.M. E&G Operations)
· 1 BA / SME
**Vendor extension (variable):**
· SWF (build capacity)
· Lemur-Salesforce partner (mobile inspection platform)
· Field mobility business vendors
· Storm response platform partners
**Current applications (from IT Networks 46-app Assessment):**

*Storm & Workforce Mgmt domain (4 apps):*
· ARCOS · Field Force · FWM (Android) · FWM (Windows)

*Maintenance domain — field subset:*
· CIR · DataMart · Device Usage History · ESOSR · Customer Locator

*Legacy / target migration:*
· Lemur-Salesforce (target mobile platform)
· DLI / TLI / DGI / TGI (inspection types within FWM, migrating)

**Active initiatives:**
· IT05 — Migration of DLI/TLI/DGI/TGI to Lemur-Salesforce
· Networks Lifecycle Apps (AGR-NET-2025-OP-02 UIL)
· Networks Mobility Program Digitalization (AGR-NET-2025-OP-01)

> Gap: app-to-stream mapping for some Maintenance items overlaps with Engineering & Asset (e.g., Customer Locator could be Doc Mgmt platform-side). Confirm with Bruno + Andrew at first workshop.

---

**VALUE STREAM CREW · Engineering & Asset Management**
Branch: IT Apps (Carlos) · biggest stream
Outcome: Engineer/planner/asset mgr designs and operates with reliable data
Functional domains: Engineering (5 apps) + Planning (7 apps) + GIS + Asset Mgmt (subset of Maintenance domain)
**Avangrid core (~4-5 FTE):**
· Product Manager (NEW)
· George Porto (today: Sr.M. + Champion GIS, dual role)
· 1-2 Tech Lead / Architect
· 1 BA / SME
**Vendor extension (variable):**
· SWF (build capacity)
· Esri / ArcFM partners (GIS)
· ProjectWise / Microstation OEM partners
· SAP HANA partners
**Current applications (from IT Networks 46-app Assessment):**

*Engineering domain (5 apps):*
· ASPEN OneLiner · Bentley PLS-CADD · Bentley STAAD · CYME · CYMECAP

*Planning domain (7 apps):*
· Bentley AssetWise · Bentley View · DSD · EnFuzion · Kaffa · Synergi Gas · PoleForeman

*Maintenance domain — engineering / asset subset:*
· Bentley ProjectWise · Cathodic (ITS) · LifeCycle Viewer (LAN SCG) · LifeCycle Viewer (Mobile SCG) · Lifecycle Full XM Edition · Looking Gas · SCAL-360 N · Dekaflow · LabVantage

*Plus enterprise platforms used by GIS team:*
· ArcFM · GIS web apps · OSG · SAP HANA · IRTH (Locate Services) · ESRI Utility Network

**Active initiatives:**
· Scheduler Project (AGR-NET-23-IT-02) · ESRI UN Deployment (AGR-NET-23-SG-01)

> Notes: This is the largest stream (~17 apps). Some apps overlap with other streams. Bentley ProjectWise also appears in Document Management context. Confirm boundaries with George Porto in workshop.

---

**VALUE STREAM CREW · Network Operations & Control** (NEW STREAM)
Branch: IT Apps (Carlos)
Outcome: Grid operator / dispatcher acts on accurate real-time data, prevents outages, minimizes downtime
Functional domain: System Operations & Control (5 apps)
Criticality: 24/7 real-time · NERC cyber/regulatory governance
**Avangrid core (~2-3 FTE):**
· Product Manager (NEW)
· TBD owner (likely under Bruno's E&G today — confirm)
· 1 SCADA / OMS SME
· NERC compliance liaison (likely shared)
**Vendor extension (variable):**
· OSI / GE / Siemens (SCADA OEMs · KB confirms Echelon partnership)
· SCADA platform integrators
· OMS / EMS specialists
**Current applications (from IT Networks 46-app Assessment):**

*System Operations & Control domain (5 apps):*
· Cimplicity · Echelon · JUMS · SLAM · Switching Orders

*Echelon scope (KB):* workflow management for automated grid devices — Recloser · SCADA Switch · Bracker. Standardized in 2 OpCos; can absorb ESOSR.

> Note: SCADA / EMS / OMS specific OEM (OSI · GE · Siemens · Schneider) not in KB. Confirm with Bruno and NERC compliance lead.

---

**VALUE STREAM CREW · Corporate & Regulatory**
Branch: IT Apps (Carlos) · scope = non-Networks IT (after PJM pool reposicionado)
Outcome: HR/employee/regulatory user self-serves and operates compliantly (non-NERC)
Functional domain: subset of "Other" (4 apps) + non-Networks corporate apps
**Avangrid core (~2-3 FTE):**
· Product Manager (NEW)
· Sarath Garikapati (Sr.M. Corporate Apps)
· Amit
· 2 Tech Leads (today)
**Vendor extension (variable):**
· Accenture (Corporate engagements)
· SWF (build capacity)
· Corporate-specific partners
**Current applications (from IT Networks 46-app Assessment):**

*"Other" domain (4 apps):*
· ECPT (Emergency Contact) · Standard Tracking System (PPE) · Utilisphere · Fluke (Met Cal)

*Plus enterprise platforms not in IT Networks scope:*
· OpenText Cloud ECM (regulatory ECM)
· Rate Case Advisor (regulatory rate-case toolkit)
· SharePoint (on-prem · migrating to O365)

**Active initiatives:**
· OpenText Cloud ECM and Rate Case Advisor (AGR-CORP-25-REG-01)
· On-prem SharePoint → O365 migration

> Note: HR / employee-portal apps are not in the IT Networks 46-app scope (different portfolio). Confirm with Sarath.

> Note: 5-6 PJMs sob Sarath hoje **vão para Facilitation Crew / shared pool** — não fazem parte do Avangrid core dessa crew.
> NERC cyber/regulatory compliance lives in Network Operations & Control, NOT here.

---

**VALUE STREAM CREW · Iberdrola Corp Services**
Branch: IT Apps (Carlos)
Outcome: Iberdrola Spain finance / procurement / budget operates without friction
Functional domain: not in IT Networks Assessment (separate scope — Iberdrola Spain)
Constraint: Iberdrola corporate governance overlay (Spain)
**Avangrid core (~3-4 FTE):**
· Product Manager (NEW)
· Vanessa Colombo (Sr.M. Global SAP)
· Michael Dubois
· 1-2 BA / SME (FI / PO)
**Vendor extension (variable):**
· Accenture (Iberdrola side — strong presence)
· SAP partners (FI, PO modules)
· Iberdrola Spain corporate vendors
**Current applications (KB-confirmed apps):**
· SAP FI — Finance *(Slide 39 vrevisada · Global SAP scope)*
· SAP PO — Purchase Orders *(Slide 39 vrevisada · Global SAP scope)*
· Budget control modules — within SAP *(Slide 39 vrevisada · Global SAP scope · "FI, PO, Budget")*
· SAP S/4HANA — successor platform (in implementation, replacing ECC) *(Initiative AGR-NET-24-IT-01)*

**Active initiatives (KB · separate from apps):**
· AGR-NET-24-IT-01 — Global S/4HANA (cross-stream — also touches CS Apps M2C-RO)
· AGR-NET-24-IT-04 — Supply Chain Digitization
· Testing Center of Excellence (TCOE) SAP Testing Automation
*Source: APR 2026 IT Avangrid Project Status Reports · Slide 4 S/4HANA Project Dashboard*

> Gap: Iberdrola Spain governance overlay constraints not in KB. Confirm with Vanessa Colombo (not yet interviewed).

---

## 4) SUPPORTING CREWS

### Capability Crew (green)

**CAPABILITY CREW**
Architecture Council
Rare expertise made available to other Crews
Reviews scope · feasibility · cross-stream dependencies

**Architect — TBD Sr.M. SAP & CS Architecture**
Vacant role in Kumar branch
~2 FTE — fill with architecture profile

**Corporate Tech Leads**
2 Tech Leads under Sarath today
Cross-stream architecture support

**Cross-stream Architects**
NEW role(s) — TBD
1-2 FTE

---

### Experience Crew (green)

**EXPERIENCE CREW**
UX & Customer Journey
Ensures customer experience is great across products
Reminds streams to think holistically

**UX Designer**
NEW role (NOM v1.5 introduces it)
1-2 FTE

**Heather LaJoie (cross-base)**
Director Product Mgmt (business side, CSP)
Bridges Digital Products + Digital CoE

**CX Researcher**
NEW or borrowed from Digital CoE
End-user research, usability testing

---

### Platform Crew (green)

**PLATFORM CREW**
Shared services to all Value Stream Crews
SLA-driven · no end-user demand of its own

**SWF & Tools**
Software Factory (vendor capacity mgmt)
Owners: Raquel Soto · Donna Henry
~2-3 FTE

**DevOps**
Pipelines · environments (dev/test/prod) · monitoring
Owner: Piotr Wielogorski
~1-2 FTE

**App Security**
Security reviews · identity · vulnerability mgmt
Owner: Piotr Wielogorski
~1 FTE

**SAP Basis**
SAP infrastructure · transports · upgrades
Owner: Nagi Reddy
~2 FTE

**Test Automation (Tricentis TOSCA)**
QA across SAP systems
Owner: TBD

**Data Ops**
Gold Client (SAP data) · Group 8760 (EDI)
Owner: TBD

**Document Management** (cross-stream platform capability)
*Document Management domain (5 apps from IT Networks Assessment):*
· DataCap · Document Viewer (LAN) · Document Viewer (REMOTE) · Mapping Computers · LifeCycle Viewer (LAN/Mobile)
Serves Engineering & Asset (drawings) · Field Ops (operational records) · Corporate (compliance)
Owner: TBD

---

### Partnership Crew (yellow)

**PARTNERSHIP CREW**
Vendor & external relationship management
Vendor journey · fair treatment · friction reduction

**Accenture liaison**
Global SAP + Corporate engagements
Cross-stream

**SWF / Software Factory liaison**
All teams' build capacity
Cross-stream

**Smart Metering vendor liaison**
100% vendor-delivered stream
M2C-specific

**Iberdrola corporate liaison**
Spain-side governance interface
Iberdrola Corp Services-specific

---

### Facilitation Crew (light purple)

**FACILITATION CREW**
Cross-cutting band — supports all Value Stream Crews
No own value stream

**Process Improvement role**
NEW — from P&S NOM v1.5
Drives Kaizen and adoption of new ways of working

**Agile Coaches**
NEW or borrowed from Digital CoE Agility Chapter
Coaches each Value Stream Crew on agile practices

**Scrum Masters**
NEW · 1 per Value Stream Crew (or shared)
Sprint cadence + impediments

**Knowledge Manager**
NEW — from P&S NOM v1.5
Documentation, lessons learned, onboarding

**Project Manager pool (Sarath's 5-6 PJMs)**
Decision pending: distributed 1 per Crew OR central pool
Today operates as central PJM service

---

## 5) FORUMS (white card, green border)

**REGIONAL FORUM**
Per OPCO discussion — multi-OPCO contention is daily pressure
Members: Value Stream PMs + business sponsors per OPCO

**BRS UIL — UI · SCG · CNG · BGC** (Connecticut cluster)
Operates: M2C · Customer Care · Field Ops streams

**BRS CMP — Central Maine Power**
Operates: M2C · Customer Care · Field Ops streams

**BRS NY — NYSEG · RGE** (New York cluster)
Operates: M2C · Customer Care · Field Ops streams

---

**FUNCTIONAL FORUM**
Per role / function discussion
Members: same role across streams

**PM Council**
Cross-stream Product Manager alignment
Quarterly cadence (BRP)

**Tech Lead Council**
Cross-stream architecture review
Bi-weekly

**Project Manager pool review**
Sarath's 5-6 PJMs + Travis (PMO)
Allocation decisions

---

**TECHNOLOGICAL FORUM**
Per tech-stack discussion

**SAP Forum — Local + Global**
Local SAP (Kumar) · Iberdrola Global SAP (Vanessa)
Quarterly · Basis + Architecture

**Non-SAP Forum**
GIS · ProjectWise · Microstation · Mobility · Doc Mgmt
Quarterly

**Platform Forum**
SWF · DevOps · App Sec · Test automation
Bi-weekly

---

**CUSTOMER JOURNEY FORUM**
End-to-end customer experience

**M2C Journey**
Meter → bill → payment
Members: Smart Metering + Billing + Collections PMs + UX

**Self-Service Journey**
Portal · CRM · Contact Center
Members: Customer Care PM + UX + business

**Field Worker Journey**
Mobility · Doc Viewer · Storm
Members: Field Ops PM + UX + field business sponsors

---

**PRODUCT FORUM**
Per product strategy

**Quarterly BRP cadence**
Cross-stream roadmap sync
Sequencing + trade-offs at portfolio level

**Cross-stream dependency mapping**
Architecture Council + PMs
Surfaces inter-stream blocks

**Trade-off + sequencing decisions**
Portfolio Manager + Vitor + business sponsors

---

## 6) BASE LABELS

**STRONGLY ALIGNED BASE — Digital Products**
Multiple Value Stream Crews forming a cohesive product set with dependencies between them.
**7 Value Stream Crews** (2 CS Apps · 5 IT Apps) · 4 supporting Crews · 1 Governance Crew · 1 Facilitation Crew · 5 Forums

Stream count by branch:
· CS Apps (Kumar): M2C · Customer Care & CCC
· IT Apps (Carlos): Field Operations · Engineering & Asset Mgmt · **Network Operations & Control** · Corporate & Regulatory · Iberdrola Corp Services

**TURF**
Area cultivated and protected by the same people.
Squads belong to a Value Stream Crew end-to-end.

---

## 7) INTAKE, PORTFOLIO ARCHITECTURE & CARD HIERARCHY

> **Final architecture:** 3 layers · 11 boards/views · vocabulary aligned with Unfix (Value Stream Crew) and Jira/P&S (Initiatives → Epics → Features → Stories).
> **Source areas tagged \*** = explicitly in KB; rest is reasonable inference to confirm via interviews.

---

### 7.1 Portfolio architecture — 3 layers (final)

**Layer 0 — VP Roll-up Dashboard** *(Jira VIEW, not a board)*
Owner: Vitor Cruz + Portfolio Manager · monthly Portfolio Sync (60 min)
Shows: aggregated capacity (planned vs actual) · cross-portfolio dependencies · top 10 Initiatives by value/risk · Class-of-Service mix
1 view total

**Layer 1 — 2 Portfolio Boards** *(real boards)*
- **Portfolio Board · Customer Services Apps** — Owner: Kumar · 3 Value Stream Crew lanes (M2C-SL · M2C-RO · CCC)
- **Portfolio Board · IT Business Apps** — Owner: Carlos · 5 Value Stream Crew lanes (Field Ops · E&A · NOC · C&R · Iberdrola swimlane*)
- Cards = **Initiatives only**
- Stages (columns): Funnel → Reviewing → Analyzing → Backlog → Implementing → Done
- Gating: Shark Tank between Analyzing and Backlog *(Pain Score H(x) + business case approved · KB: P&S NOM v1.5)*
- Cadence: monthly Portfolio Committee per board (90 min)
2 boards total

\* Iberdrola = swimlane within IT Apps Portfolio with governance overlay Iberdrola corporate (NOT a separate board)

**Layer 2 — 8 Crew Kanbans** *(real boards)*
1 board per Value Stream Crew · Owner: each crew's PM (NEW role)
- Cards = **Epics → Features → Stories → Subtasks** (linked back to parent Initiative on Layer 1)
- Layout: 5 phases × 2 tracks
  - **Phases (columns):** Phase 0 Discovery → Phase 1 Materialization → Phase 2 Creation → Phase 3 Ongoing → Phase 4 Maintenance *(KB: P&S Compass model, Slide 44 nWOW v1.5)*
  - **Tracks (rows):** Discovery (always-on research/validation) + Delivery (sprint execution)
- Sub-views (Jira-native, same project): Kanban view · Sprint board view · Roadmap view · Backlog view
- Cadences: daily standup · sprint planning/review/retro (2-week) · weekly refinement · weekly Discovery sync
8 boards total

**Total: 1 view + 2 portfolio boards + 8 crew kanbans = 11 boards/views**

---

### 7.2 Card hierarchy (Jira issue types)

```
Initiative          (Layer 1 — Portfolio Board)
   └─ Epic          (Layer 2 — Crew Kanban)
         └─ Feature
               └─ Story
                     └─ Subtask
```

**Initiative** · multi-quarter · 3-12 months · approved at Portfolio level · big project
**Epic** · multi-sprint · 1-3 months · large body of work within an Initiative
**Feature** · 2-6 weeks · deliverable increment within an Epic
**Story** · 1 sprint · single user-facing capability
**Subtask** · hours/day · technical breakdown of a Story

KB-confirmed equivalence with P&S terminology *(nWOW v1.5 Slide 81 + BRP Q2 2026 Slide 5)*:

| Avangrid Digital Products (this proposal) | P&S Networks IT (KB) |
|---|---|
| Value Stream Crew | Program |
| Crew Kanban | Program Kanban |
| Initiative · Epic · Feature · Story | identical |

P&S Q1 2026 KB-confirmed real Initiatives: Global SAP S4Hana · DLI/TLI Implementation · Network Mob. Pain Points · NY Gas Inspections · Gas Facilities Record Archive · Storm Program Pain Points · Supply Chain Digitalization · Doc Viewer · ESRI UN Deployment.

---

### 7.3 Intake channel per Value Stream Crew

**M2C — Service Lifecycle**
Areas: Customer Service Products (Heather business)\* · Customer Ops US\* · BRS UIL/CMP/NY\* · Field Services\* · Smart Metering vendor\*
Channel: ServiceNow form "M2C — Service Lifecycle"

**M2C — Revenue Operations**
Areas: Customer Service Products\* · Billing analysts · Collections · Regulatory (LIDR · EBD · Hardship)\* · Tax/SOVOS
Channel: ServiceNow form M2C with sub-routing

**Customer Care & CCC**
Areas: Customer Service Products (Heather)\* · Contact Center supervisors · Voice of Customer (Medallia) · Digital CoE\*
Channel: ServiceNow form "Customer Care" + Medallia integration

**Field Operations**
Areas: Networks Business (P&S adjacent) · Field Services\* · Storm response · Workforce dispatchers
Channel: ServiceNow form "Field Operations" + Lemur-Salesforce integration

**Engineering & Asset Management**
Areas: Networks Business (P&S)\* · Engineering teams · Planning teams · Asset planners
Channel: ServiceNow form "Engineering & Asset" with sub-routing (Engineering / Planning / GIS / Asset)

**Network Operations & Control**
Areas: Network Operations Center (SCADA operators) · NERC compliance · Storm response coordination
Channel: Dedicated 24/7 channel — separate from standard ServiceNow · integrated to Echelon\*

**Corporate & Regulatory**
Areas: HR · Regulatory (non-NERC) · Internal corporate users · Compliance
Channel: ServiceNow form "Corporate & Regulatory"

**Iberdrola Corp Services**
Areas: Iberdrola Spain corporate finance\* · Procurement (Spain) · Budget owners
Channel: Iberdrola corporate channel (Spain governance overlay)

---

### 7.4 Class of Service (Initiative tagging at Portfolio Board)

| Class | % capacity (cap) | Trigger | Bypass Shark Tank? |
|---|---|---|---|
| Standard | ~50% baseline | Normal demand | No |
| Fixed-Date | ~20% | Regulatory deadline · contractual | No · prioritize by date |
| Expedite | **≤20% (cap)** | Regulatory NERC · emergency · safety | Yes · log mandatory |
| Intangible | ~10% | Innovation · tech debt · UX research | No |

KB-grounded: Renewables DTO uses Class-of-Service taxonomy; Slide 7 v2.0 SAP CoE confirms "Regulatory always wins" as AS-IS pain → cap at ≤20% is the proposed replacement.

---

### 7.5 Open decisions (intake & boards)

1. **ServiceNow vs Jira Intake direct** — pending: channel volume + Travis/Vihn interviews
2. **Iberdrola separate channel?** — depends on Iberdrola governance overlay (Vanessa not interviewed)
3. **Network Operations & Control 24/7 channel** — recommend SEPARATE from standard intake (real-time outage tickets cannot wait for weekly triage)
4. **Cross-crew Initiatives** (e.g. S/4HANA touches CS + Iberdrola) — coordinate via Customer Journey Forum + Layer 0 VP Dashboard, NOT a shared board

---

## 9) STAGE 01 — DEMAND MANAGEMENT · DEFINITIONS (per Miro template categories)

> Each post-it is one item. KB-grounded where citable; **NEW** = proposal.

---

### CAPABILITIES (14 items — green)

**Demand intake channels (formal & informal)**
Principle: one front door per Value Stream Crew (Slide 36/59 vrevisada Ana).
Decommission: Email · MS Teams · Word · OneNote.
Keep IT Now/RITM as incidents only.
Tool decision (ServiceNow vs Jira Intake) pending channel volume.

**Demand registration and logging**
Auto-create Jira issue from intake form (zero re-keying).
Immutable fields: requester, BU, crew, date, COS tentative.
Audit trail every state change. Rejected demand also logged.

**Channel consolidation (single vs multiple entry points)**
One form per crew, adaptive based on crew selection.
Zero alternative formal channels post-cutover.
Adoption KPI: ≥90% via single channel by month 6.

**Standard intake formats / templates**
Confluence-managed canonical template, versioned.
Mandatory DoR fields: title, description, BV, AC, COS, work type, due date, sponsor, OPCO, value stream.
Form blocks submission without mandatory fields.

**Demand classification (type, size, urgency)**
Type taxonomy: Feature · Enhancement · Bug · Tech-debt · Compliance · Research.
Size T-shirt at intake (XS/S/M/L/XL). Urgency = Class of Service (separate axis).
Classification by Demand Manager at Gate 1, not self-declared.

**Differentiation: run / change / regulatory / innovation**
Maps to 4 Classes of Service: Standard · Fixed-Date · Expedite · Intangible.
Caps: Run ≤30% · Change ≥40% · Regulatory ≤20% · Innovation ≥10%.
Replaces today's "Regulatory always wins" with capped Expedite lane.

**Business value definition**
Adopt H(x) = (BV·w₁ + UX·w₂ + Safety·w₃)/E (CoE × P&S/P&T hybrid; baseline w₁=3, w₂=2, w₃=2).
BV 1-5 across 4 dims: revenue/cost · customer impact · strategic fit · risk reduction.
Sponsor signs off BV at intake.

**Availability of acceptance criteria**
AC mandatory in DoR — no AC, item sent back, not scored.
Format: Given/When/Then or measurable outcome.
Reviewed by PM + Tech Reviewer before scoring. Immutable post-prioritization.

**Technical feasibility input**
Lightweight check at Gate 1 by named Architect (1-of-3 rotation).
Output: Feasible / Conditional / Not feasible + T-shirt effort.
Time-boxed ≤2h per item; longer becomes Spike.

**Triage and initial validation activities**
Gate 1 weekly, 30-min time-box.
Demand Manager (chair) + 2 PMs + 1 Architect.
Outcomes: Accept / Send-back / Reject. SLA ≤5 business days. Read-only public log.

**Demand routing and assignment**
Routing key = Value Stream Crew, not team or person.
Each crew has named PM + backup. Cross-crew items → Architecture Council.
Mis-routing rate KPI ≤5%.

**Handling of urgent or exception requests**
Urgency = Class of Service decision, not self-declared.
Expedite (Regulatory) bypasses Gate 1; logged for Gate 2 review.
True P0 incidents → Service Desk SLA (NOT intake). Expedite cap ≤20%; breach = escalation.

**Demand visibility and tracking**
Single Demand Dashboard (PowerBI or Jira-native) — read-only public.
Views: inflow funnel · backlog by crew · COS mix · aging · top-20.
Auto-refresh ≤1h. Per-requester portal.

**Identification of duplicate / overlapping demand**
Form runs similarity search at submit (top-3 matches surfaced).
Weekly dedup query. Duplicates linked-as-duplicate-of (original wins).
Overlaps escalated to Gate 2 for consolidation. Dedup rate as KPI.

---

### ROLES (4 items — teal)

**Business requestors / stakeholders**
Customer Service Products org (Heather LaJoie) for CS streams · Networks Business/P&S for IT Apps · Iberdrola Spain corporate · BRS UIL/CMP/NY for OPCO-specific.
Today: 12+ BUs (Slide 6 v2.0); to-be: single intake form per crew routes to right backlog.

**Product or demand managers**
8 Product Managers — **NEW** for all crews (today only PJMs/Sr.Managers; Iram is PJM, not PM).
1 Demand Manager — **NEW** in Facilitation Crew, audits scoring across crews.
Heather LaJoie bridges business→IT for CS streams.
Re-skill candidates: Iram · Andrew · George · Sarath · Vanessa (current Champions/Sr.Managers).

**Architects or technical reviewers**
Capability Crew (Architecture Council).
Members: TBD Sr.M. SAP & CS Architecture (vacancy) · 2 Corporate Tech Leads · cross-stream architects.
Time-boxed feasibility input at Gate 1 (≤2h). Advisory role, not veto.

**Triage / intake committee**
Gate 1 weekly per Portfolio Board.
Composition: Demand Manager (chair) + 2 PMs + 1 Architect.
30-min time-box. Outcomes: Accept · Send-back · Reject.
Cross-portfolio escalations → Gate 2 (monthly Portfolio Committee).

---

### TOOLS (5 items — amber)

**Demand capture tools**
ServiceNow form per Value Stream Crew (recommended) · adaptive based on crew selection.
Backup: Jira Intake direct (Renewables DTO model) if ServiceNow governance is too slow.
Iberdrola separate channel (Spain governance overlay).
Network Operations & Control: dedicated 24/7 channel.

**Demand tracking & backlog tools**
Jira = system of record post-Gate 1.
1 Crew Kanban per Value Stream Crew (Layer 2). Issue types: Initiative → Epic → Feature → Story → Subtask.
5 phases × 2 tracks (Discovery + Delivery) per P&S Compass model.
PPM Now retained as Smart Metering portfolio view (not consolidated yet).

**Communication channels**
Triage outcomes auto-published to requester via Jira automation.
Per-requester portal showing status of own submissions.
Slack/Teams integration for Class Expedite alerts.
Forums (Customer Journey, Product) drive cross-crew comms.

**Documentation tools (request details)**
Confluence: DoR template · Class of Service taxonomy · per-crew product brief.
Sponsor approval logged in same intake form (no separate channel).
Document Management capability (5 apps · Maintenance domain) → Platform Crew.

**Reporting / visibility tools**
Single Demand Dashboard (PowerBI or Jira-native): inflow · backlog · COS mix · aging · top-20.
VP Roll-up Dashboard (Layer 0): cross-portfolio capacity, top Initiatives.
Per-Director board: CS Apps · IT Apps Initiatives view. Auto-refresh ≤1h.

---

### GOVERNANCE (6 items — navy)

**Demand review forums**
Gate 1 (weekly · 30-min) per Portfolio Board · Triage Committee.
Gate 2 (monthly · 90-min) per Portfolio Board · Portfolio Committee.
Layer 0 monthly Portfolio Sync (Vitor + Portfolio Manager).
Per-crew refinement: weekly inside Crew Kanban.
BRP quarterly: cross-portfolio strategic review (KB: P&S model).

**Decision-making practices (accept / reject / redirect)**
Accept: DoR pass + scored + capacity → assigned to crew.
Reject: explicit reason logged, requester notified.
Send-back: DoR fail (missing AC/BV/etc.) — requester re-submits.
Redirect: wrong crew → auto re-routed by form logic.
All 4 outcomes audit-logged.

**Escalation mechanisms (urgent / conflicting requests)**
Cross-crew conflict → Customer Journey Forum (M2C) or Product Forum.
Cross-portfolio conflict → Layer 0 VP Dashboard + monthly Portfolio Sync.
Capacity escalation: Director → Vitor.
Class Expedite reviewed at next Gate 2 (no quiet overrides).
True P0 incident → Service Desk SLA (separate fast path).

**Definition of intake guidelines or policies**
DoR template = Confluence single source of truth.
Class of Service taxonomy + capacity caps documented.
Caps: Run ≤30% · Change ≥40% · Regulatory ≤20% · Innovation ≥10%.
Reviewed every 6 months by Demand Council.
Grounding: P&S NOM v1.5 + Renewables DTO Class-of-Service.

**Alignment mechanisms with business stakeholders**
BRP quarterly (P&S model · KB-confirmed).
Customer Journey Forum (per journey) for CS streams.
Per-Director monthly review with their sponsors.
Heather LaJoie cross-base bridge for CS journeys.
Sponsor sign-off mandatory at intake (BV approval).

**Oversight of demand inflow and backlog health**
Demand Manager owns Demand Dashboard (in Facilitation Crew).
Aging > threshold = auto escalation to Gate 2.
Backlog health: # ready vs # blocked · velocity vs commitment.
Quarterly demand-vs-delivery conversion review.

---

### PERFORMANCE (7 items — red)

**Demand volume and inflow trends**
Submissions per week/month/quarter · per Value Stream Crew · per Class of Service · per OPCO (UIL/CMP/NY) · per source area (CSP/Networks Business/Iberdrola).

**Demand quality (completeness, clarity)**
DoR pass-rate at first try (target ≥70%).
Re-submission rate per requester (high = coaching).
AC quality score (peer-reviewed). Sponsor sign-off rate.
Workshops fed by lowest-scoring BUs.

**Triage lead time / time-of-demand**
Median submission → Gate 1 decision (target ≤5 business days).
90th percentile (outlier catch).
Per Class of Service breakdown · per crew breakdown.

**% demand entering through formal channels**
Target ≥90% via single channel by month 6.
Today: unknown (channel volume not measured — KB gap).
Email/Teams/Word residuals tracked monthly (decommission progress).

**% urgent or bypassed requests**
Class Expedite as % throughput (cap ≤20%).
Regulatory subset within Expedite tracked.
"Urgent" claimed but auto-downgraded to Standard rate.
Bypass attempts (not via single channel) tracked separately.

**Demand duplication / overlap indicators**
Dedup rate (% submissions linked-as-duplicate).
High + rising = upstream noise (BU coordination needed).
Low + falling = healthy intake.
Overlap escalations to Gate 2 (consolidation decisions).

**Visibility and traceability of demand status**
% requesters who find own submission status (target 100%).
Average time to status update (Demand Manager → requester).
Replaces today's "manual data gathering" pain (Slide 7 v2.0).
Per-BU portal showing all their submissions.

---

## 10) SEGREGATED BASE (vendors — outside main Base)

**SW DELIVERY BASE**
Fully Segregated Base
External delivery — independent governance
Vendors operate outside Avangrid Base

**Accenture**
Global SAP + Corporate engagements

**SWF / Software Factory**
All teams (cross-stream build capacity)

**HCL · Nuance**
Customer Service support

**Infosys**
System support, integration

**Smart Metering vendor**
100% delivery (requirements + code)

**SAP partners**
Iberdrola side support

---

## Color guide for Miro post-its

| Crew / Forum | Suggested Miro post-it color |
|---|---|
| Governance Crew | peach / light orange |
| Value Stream Crews | purple |
| Capability / Experience / Platform | green |
| Partnership Crew | yellow |
| Facilitation Crew | light purple |
| Forums | white (with green border or label) |
| Sibling Bases | grey |
| **CS Apps badge** | blue tag |
| **IT Apps badge** | orange tag |
| **NEW tag** | orange |
| **KB tag (exists today)** | green |
