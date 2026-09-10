# Stakeholder Interview — Kumar Jaladanki + Stephanie (business counterpart)
## Project: SAP Digital Products
## Date: April 2026

> ⚠️ **Correction note (2026-04-17)**: This file was originally titled "Stephanie (Head IT CS), Kumar (Applications)". Corrected against validated Miro org chart:
> - **Kumar Jaladanki** is **Director Customer Services Apps** (IT-side, reports to VP Vitor Cruz — same level as Carlos Garbin).
> - **Stephanie** is NOT in the IT org chart. She is the **business-side counterpart** from Customer Service business (under the VP Customer Service orgs). She joined this session to provide business context.
> - "7 OPCOs" describes the **business Customer Service organization**, not the IT structure.

## Participants
- **Kumar Jaladanki** — Director Customer Services Apps (IT)
- **Stephanie** — Customer Service business-side (not in IT org chart)

---

## Customer Service Context — Business Organization (from Stephanie)

The business side is structured as **3 groups representing 7 OPCOs** (feeds demand into Kumar's IT organization):
- Call Center (customer interaction)
- Digital COE (Customer) — drives customer-facing apps, websites
- VP Customer Service → BRS UIL (UI, SCG, CNG, BGC)
- VP Customer Service → BRS CMP
- VP Customer Service NY → BRS NY (NYSEG / RGE)

Demand flow into IT:
- Customer Finance → Contact Center + Customer Service Organization
- Customer Service Organization splits into: Annual PDD (large projects) and Small / Medium Demand

---

## Kumar's IT Organization (Customer Services Apps branch)

Kumar's 5 direct-report teams (from validated org chart):
- **Ivelisse Velazquez** — Manager FICA, Billing & Collection *(not yet champion-interviewed)*
- **[TBD]** — Sr. Manager CRM & Contact Center Ops *(vacant; not yet interviewed)*
- **Vincent Lam** — Sr. Manager Device & Metering *(Iram Chowdhury / Smart Metering sits here — partially covered)*
- **[Open]** — Sr. Manager SAP & CS Architecture *(vacant)*
- **Nagi Reddy** — Sr. Manager SAP Basis *(not yet interviewed)*

**Local SAP = Kumar's branch** (Avangrid customer services environment: Billing, CRM, FICA, Metering) — contrast with **Global SAP** under Vanessa Colombo (Iberdrola/Spain: FI, PO, Budget control).

---

## Structural Split from Excel (Travis / PMO / BRS / IT funded)

Project ownership in Customer Services:
- **PMO Group (Travis Structure)** — internal PMs for **large projects** (works with Kumar's team)
- **BRS (business)** — business side
- **IT funded** — Kumar's own teams
- **Jira** for small–medium demand; **PPM tool** for large demand (projects)
- **Software Factory (SWF)** typically manages the project itself

### NTT open questions — raised during workshop
These were **NTT's diagnostic questions** surfaced in the post-it canvas. Some were answered live by Kumar/Stephanie, others remain open.

**Answered:**
- Q: Do you have Product Manager + SAP Architect roles in Demand Management for large projects?
  - A: **"Yes, already operating"** — the to-be structure is in place.

---

## Pain Points & Aspirations (IT CS)

### Top Pain Points
1. Access to documentation — better organization (storage system) — need single source of truth
2. Who is in charge of what — 3 people can run the payment jobs for example (unclear ownership)
3. Capacity management
4. Resource allocation — Multiple OPCOs competing by capacity
5. Information (updates) on e-mails not in Jira — incident mix between BRS and Software Factory

### NTT diagnostic questions — status

These are **NTT's open diagnostic questions** surfaced during the workshop. Status is tracked per question.

**Answered during the session:**
- Q8: Why for large projects don't we have Product Manager + SAP Architect in Demand Management?
  → Already operating in the to-be (Kumar)
- Q9: Who is working as the Project Manager?
  → PMO Group (Travis Structure) for large projects; Kumar's team for small-medium
- Q2: Fixes and configuration mainly done by Vendors?
  → Confirmed — vendor-led for most execution

**Still open (require follow-up):**
- Q1: What is the difference between Functional versus Technical SAP Roles?
- Q3: Is small and medium demand only for BRS, or does it span other areas/teams?
- Q4: Who is the System Integrator?
- Q5: What are the large projects per area/team (2026 roadmap)?
- Q6: How are we tracking the work that is in development by the Factory/Vendor?
- Q7: How do the projects connect to applications and support teams?

### Key Findings
- Project estimation is done by the vendor
- Differing estimations, Budget and financial changes are a recurring problem
- Project estimation done by the vendor creates dependency and lack of control

---

## AS-IS Demand Flow — SAP (Kumar's Area)

### CMP / NY
- SAP → BRS → SharePoint
- BSS log database CMP (if can't solve, move it to Jira)
- Jira (intake) — all — level 2
- BRS Business shares the priority
- Low to critical (BSS log) criteria maybe added by Business

### UIL
- ITNow (intake) — day to day issues specific to UIL
- Functional team works with the squads on requirements
- NEED: Establish a process for demand/prioritization + Release cadence

### Digital COE
- Digital PMs (Functionals SAP from Kumar's team overseeing the work)
- Split: 30% BRS and Internal IT, 70% other
- Digital PMs → Big Room Planning / Squads with visibility of what is coming
- Heather is the first layer for prioritization

### Infrastructure (Data & Analytics)
- Data & Analytics has demand directly from 3 main groups
- MLA lot of demand with "I need a priority" urgency
- Regulatory and net requirements to Automation (aptos, BI weekly)
- Major maintenance: they reach to the team for "upgrade" on Sunday, Weekly

### Regulatory
- Sometimes reactive — directly to development
- Regulatory Mandatory SAP (External Drivers) — Urgency #1

### Procurement
- Cadence with procurement — review format related

### OT
- They own outside Technology, automated solutions, infrastructure of automated machines
- Demand planning once a year: small, medium and large (same process, not specific, recurrent)

### PDD (Large Projects)
- PDD goes to all of the groups (universal process for large projects)

---

## Portfolio Findings
- No prioritization model defined. But forums exist to prioritize
- PPM tool needed for large projects — project budgeting management
- Jira is working well for small demands
- Question: Which tools are used to manage Vendor work?
- Question: How are we planning to manage it from now on?
