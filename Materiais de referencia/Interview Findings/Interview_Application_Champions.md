# Application Champions — Group Discovery Session
## Project: SAP Digital Products
## Date: April 2026
## Duration: 60 minutes

> ⚠️ **Correction note (2026-04-17)**: Champion area names updated to reflect validated org structure. The 6 interviewed champions span **5 of 9 Sr. Manager / Manager teams under Carlos + Kumar**. Coverage gap on Kumar's branch: 4 teams not yet interviewed (Ivelisse FICA/Billing, TBD CRM, Open SAP Architecture, Nagi SAP Basis).

---

## Champions (mapped to real org structure)

| Champion | Sr. Manager / Team | Branch |
|----------|-------------------|--------|
| Amitkumar Patel | **Sarath Garikapati** — Corporate Applications | Applications (Carlos) |
| George Porto | **George Porto** — GIS & OSG | Applications (Carlos) |
| Andrew McGibbon | **Bruno Moreto** — Electric & Gas Ops (G&E sub-scope) | Applications (Carlos) |
| Michael Dubois | **Vanessa Colombo** — Global SAP (Iberdrola/Spain) | Applications (Carlos) |
| Iram Chowdhury | **Vincent Lam** — Device & Metering / Smart Metering sub-scope (Avangrid) | Customer Services (Kumar) |

---

## Q01: What are your top priorities today?

### GIS (George Porto)
- Keeping the lights on — Share a GIS server — Managing team's SAP tickets including major projects
- Keeping people and removing barriers — defining future technology roadmap
- Team time constraints, what we should be doing / people side is difficult (adoption of the technology)

### Corporate (Amitkumar Patel)
- Regulatory and Employee — not only one year plan and training
- Coming with new requirements that are always a priority, to an estimate basis

> Note: Sarath Garikapati is the Sr. Manager of Corporate Applications but did **not** participate in the workshop. Amit attended on behalf of the Corporate team.

### G&E (Andrew McGibbon)
- Third party risk group — env move is safe enough
- Cyber — legal (physical security)
- Lots of competing priorities that are not priority but my priority
- Version upgrades, business wants to need to evaluate budget/time

### Global SAP — Iberdrola (Michael Dubois, under Vanessa Colombo)
- Balance all the demands (Incidents, Small Demands, Major projects PDD)
- Limited resources (1 year to hire someone), accelerate turnover, knowledge transfer loss
- Scope: Iberdrola (Spain) corporate functions — FI, PO, Budget control

### Smart Metering — Avangrid (Iram Chowdhury, under Vincent Lam / Device & Metering)
- Ability to manage the priorities, limited resources, turnover
- Training process, learning curves, staff augmentation, retirement windows
- Scope: Avangrid customer services (part of Local SAP footprint: Billing, CRM, FICA, Metering)

---

## Q03: Which tools do you use?

### GIS
- Project Management: I don't manage large projects. But I use One Note, Teams and Spreadsheets to manage.
- My team uses Jira for misc. work efforts and improvements to manage some work with the business.
- We also use IT-Now to manage daily break/fix as well as app requests.

### Corporate
- PM Jira and Confluence, word document for business requirements

### G&E
- I receive requests/input from the business or from the vendor

### Global SAP
- IT Now (Incidents), RITM for minor requests. JIRA for Small Demands.
- Teams and Sharepoint are typically used for project work.
- Accenture and the SWF play a significant role in supporting all of these duties.
- P&T serves as the intermediary with business — global SAP manages Jira.
- Meetings: Every thursday to go over all over the incidents — Accenture

### Smart Metering
- (Tools response to be captured)

---

## Q05: How do you prioritize work today?

### GIS (George Porto)
- For small demand or day to day, I prioritize the things that have the highest impact to the most users. In most cases these are break/fix requests.
- For medium to large projects, this is prioritized based on the PDD demand process of approvals and thus my team time/effort is anticipated and planned farther in advance. We look to bring in SWF where we may see a resource impact.

### Global SAP (Michael Dubois)
- First Priority: Incidents. Second priority: Small Demands and Major Projects

---

## Q06: How do you receive demand requests?

### GIS (George Porto)
- Small demand: Email, meetings, Jira, IT-Now tickets
- Medium to large projects: yearly PDD demand process. More formal project management process.

### Global SAP (Michael Dubois)
- Incidents (IT NOW), Small Demands (JIRA), Minor Requests (IT NOW RITM), Major Projects (SharePoint)

---

## Demand × Portfolio Flow Canvas — AS-IS by Area

### Corporate (Amitkumar Patel)
- INTAKE: PM Jira and Confluence, word document for business requirements. Usually small and medium demands.
- CLASSIFICATION: PDD is the driver. Lots of overlapping projects.
- PRIORITIZATION: PDD is the driver.

### Global SAP (Michael Dubois)
- INTAKE: P&T serves as intermediary with business. Global SAP manages Jira. IT Now (Incidents), RITM for minor requests, JIRA for Small Demands, Teams and SharePoint for project work. Accenture and SWF play significant role.
- CLASSIFICATION: Incidents (IT NOW), Small Demands (JIRA), Minor Requests (IT NOW RITM), Major Projects (SharePoint).
- INSIGHT: Difference between small and big projects (PDD) — maybe not the best solution, maybe just managing the projects in Jira.
- PRIORITIZATION: Biggest part are incidents (priority). Sometimes entire team is focusing on incidents. P&T prioritization for small demand and projects.

### GIS (George Porto)
- INTAKE: Jira, emails, meetings, IT Now (break fix). Business opens Jira tickets for small work.
- CLASSIFICATION: Small demand via Email, meetings, Jira, IT-Now. Medium to large via yearly PDD, formal PM process.
- PRIORITIZATION: Business is driving prioritization. P&T is the face — helping prioritize the work.

### G&E (Andrew McGibbon)
- INTAKE: Receives requests/input from the business or from the vendor. Channel: Emails.
- CLASSIFICATION: Depends on the nature of the demand. Depends on upgrade type. They set their own timeline.
- PRIORITIZATION: Internal prioritization depending on the nature of the demand (operational problems = urgent).

### Smart Metering (Iram Chowdhury)
Canvas has 5 stages — Intake → Classification → Prioritization → Portfolio → Delivery

- INTAKE:
  - Small demands via e-mail (IT Now)
  - Normal flow of intake — not seasonal
  - Projects — Incidents (more volume than projects)
  - Goals: ensure stability, support business, implement changes
- CLASSIFICATION:
  - Usually clear requirements
  - **All portfolio in vendor** (100% outsourced)
  - Usually no urgent requests (just incidents and fixes)
  - Incidents — IT Now (template)
- PRIORITIZATION:
  - Jira and Confluence (used by project managers)
  - Usually large projects — PPM Now
  - Criteria: "What is the impact of not happening — criticality"
  - Prioritized by deadline and priority of the application. "Intrinsic common sense what is to be prioritized"
- PORTFOLIO:
  - Weekly meeting — work in priority (internal with manager) — PMs help them to prioritize
- DELIVERY:
  - Portfolio in Confluence
  - **Vendors are the consultants for the work** — getting requirements on tech side, steps to follow and do the code also — all custom applications

**Key insight — Smart Metering is the ONLY area to explicitly use PPM Now for large projects, and is 100% vendor-delivered (vendors do requirements AND code).**

---

## Cross-Area Patterns Identified

### Tool Landscape
| Area | Incidents | Small Demands | Large Projects | PM/Docs |
|------|-----------|---------------|----------------|---------|
| GIS | IT-Now | Jira + Email | PDD + Spreadsheets | OneNote, Teams |
| Corporate | — | Jira | PDD | Confluence, Word |
| G&E | — | Emails from Business/Vendor | — | — |
| Global SAP | IT Now | Jira | SharePoint + PDD | Teams, SharePoint |
| Smart Metering | IT Now | Email (IT Now) | PPM Now | Confluence (portfolio), Jira |

### Common Pain Points Across All Areas
1. Resource constraints / turnover / hiring difficulty (ALL areas)
2. Balancing BAU (incidents) vs strategic projects (GIS, Global SAP, Smart Metering)
3. No standardized prioritization model — each area uses different criteria
4. Fragmented tooling — no single source of truth across areas
5. PDD as the only formal process, but only for large projects
6. Business involvement in prioritization varies by area

### Actual prioritization variables used in forums (from Excel validation)
> *"No prioritization variables defined but forums to prioritize. Currently in the forums: if its strategic, operational urgency, regulatory (most important doesn't matter the size), size"*

Four informal criteria applied in governance forums today:
1. **Strategic fit**
2. **Operational urgency**
3. **Regulatory** (most important — overrides size)
4. **Size**

These are applied ad-hoc, without a shared scoring model.

---

## How Decisions Get Made — Governance Forums

*Source: Governance forums from RFI — to be validated with stakeholders*

| Forum | Cadence | Participants |
|-------|---------|-------------|
| Critical Integrations | Weekly | Business Apps, P&S, Digital Products |
| Daily BSS BI/BW Stand Up | Weekly | Data & Analytics Teams, Digital Products |
| Data, Analytics & Solutions Check-In | Weekly | Data & Analytics Teams, Digital Products |
| CX & Digital CMP Update | Monthly | CX Digital, Operations, CMP President, VP CS |
| Manager Check-in | Weekly | [To be confirmed] |

### Portfolio Management Gaps
- **No tool to support portfolio management**
- **No tools or process to manage delivery process of vendors**
