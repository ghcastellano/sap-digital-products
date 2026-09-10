# Intake Slides · Reviewer Feedback Reference

Cross-reference for **Preparation & Planning v4 QA**. All comments from
`TOM MAP workbook - NEW.pptx`, sourced from `ppt/comments/modernComment*.xml`.

**24 comments · 2 reviewers · 2026-04-24 → 2026-05-06**

---

## Ana Sancha Medina · 12 comments

### Slide 1 · 2026-04-24 · Demand Mgmt + Prep&Planning framework
- Single front door (incidents · improvements · regulatory · innovation · tech debt)
- Demand shaping (problem-impact-value, not solutions)
- Run / Grow / Transform classification + Core SAP vs satellites
- Light but formal Discovery (problem statement · personas · value+risk hypotheses)
- Value + prioritization framework: BV, risk, regulatory urgency, **Cost of Delay**
- **Preparation & Planning · Product Vision:**
  - Products ≠ Applications (e.g. "Order to Cash SAP" ≠ "SAP SD")
  - Outcomes (not deliverables)
  - Roadmapping rolling wave 3-6-12 mo
  - **Differentiate commits vs hypotheses**
- **Preparation & Planning · Product Design + Architecture:**
  - Dual track: Discovery (what+why) vs Delivery (how+when)
  - Business process design E2E (O2C · P2P · R2R)
  - Non-functional requirements (perf · security · compliance · scalability)
- Execution: Product teams (PO + TL + delivery) · backlog per product · flow-based · dependency mgmt
- Deployment: hypercare + feedback loop · Measure & Learn

### Slide 3 · Light discovery
- Don't explain hypotheses of value or hypotheses of risk
- What are the main deliverables we expect and from whom?

### Slide 4 · Roles
- Who is portfolio manager? Who is Product Manager Director?
- Domain product owner not explained
- Intake lead not explained
- Business process owner not explained
- Performance metrics not defined

### Slide 5 · P&T
- P&T is **NOT** GIS + Global SAP
- P&T = **Gas & Electric Ops**
- Slide 18 vs 34 — align both into one

### Slide 6 · Budget
- Should requester indicate if budget exists?
- Put Budget in forms
- Why is budget yellow? Don't see it in forms

### Slide 11 · IT Now
- IT Now tickets stay separate (Renewables does this · works)
- Incidents from IT Now should NOT mix with Jira tickets
- Clarify how production-support differs

### Slide 16 · Size + Gates
- Slide 27 says M only flags portfolio · here says M goes to portfolio gate
- Unify definition
- **XL → PPM Now (per Heather)**
- Slide talks about L · others talk about XL · align

### Slide 18 · Cast
- Missing new vacancy in "decide"
- **Class of Service responsibility to middle managers** (Kumar's + Carlos's)
- Portfolio must include **middle management of both teams** — not just Heather + new PMs
- Day-to-day is in middle managers of Kumar today

### Slide 23 · Discovery + Roles
- Light discovery artefacts not explained
- **Heather missing in director view / new vacancy**
- View by role + stage should integrate into (Requester / Reviewing+Analyzing / Gate 2)

### Slide 25 · Jira visualization
- Show as Jira visualization
- Role taxonomies don't match
- Mix with slide 37, illuminate by stage that applies

### Slide 26 · Middle managers
- **Incorporate middle managers**

### Slide 28 · OKRs + portfolio committee
- **NO OKRs at Avangrid (zero)**
- Use **metrics seen with Heather + Vitor** (not OKRs)
- In portfolio include **middle managers**
- How is this realized (frequent action) — not just one phrase

### Slide 29 · BRPs + BRS
- NO OKRs
- Include **who and when manages outcome of BRPs**
- **BRS as new governance practice** (process owner / value stream counterpart) vs Product Managers under Heather

---

## Alfredo Soto Cuellar · 12 comments

### Slide 7 (×2) · OPCO equation
- OPCOs included in OPCO field for prioritization based on value components
- Usually 1 OPCO request vs 1 request for all OPCOs · which is more valuable?
- Include OPCO equation — they need control over which OPCOs are impacted
- Requests by OPCO may be duplicated between OPCOs

### Slide 12 · Multi-issues
- Slide 28 has Champions which doesn't make sense
- Board configuration different at portfolio level
- Not clear what goes to team level vs portfolio depending on size — need one view explaining the logic
- Vincent Lam structure to be migrated to CS (after talking to Carlos)
- **Meter & Device Operations layer is on top of CS + ITBA — not correct**

### Slide 13 · Effort + Type
- How does effort estimation integrate at this stage?
- Type identified by Product Manager (regulatory / non-reg / production support) — but PM won't know if regulatory or not
- Regulatory included in requester form already — mix
- Production Support could be misunderstood with incidents

### Slide 15 · Type taxonomy
- Production Support · Regulatory · Non-Regulatory + adding **"change category"** — what does regulated change mean? complex the ecosystem
- We already have regulatory flag
- Work type of slide 31 may mix with Type of demand — simplify
- Description of upgrade has regulatory mention (same as before)
- If Production Support = Incidents in IT Now → mixing 2 channels
- **Production Support could be > M** so could enter both — clarify
- Standardize Gate names per type clearly · which gates pass / which don't
- Sponsor sub-team lead · business unit lead — different taxonomy from rest

### Slide 17 · BV
- Considered in another field in BV definition?
- Transformational (5) = any regulatory request automatically?

### Slide 19 · CoS
- Class of Service: without capacity planning assignment process drawn, hard to relate
- Need the process behind to maintain Class of Service

### Slide 20 · Pain Score CS
- Slide 32 — CS adaptation has Safety (department doesn't work with crew employees · works for customer services)
- Pain formulas very complex for them — simplify more
- Don't include the left-side origin (P&T inspiration) — leave 2 clearer blocks for CS pain and ITBA pain

### Slide 21 · OPCO + CS Safety + Effort SMEs
- (OPCO equation — same as slide 7)
- (Slide 32 CS Safety — same as slide 20)
- Effort SMEs at TRIAGE — **who is responsible** and **what is the definition of 1-5 of complexity**?
- Same as BV, the rest of variables are not defined

### Slide 22 · Triage cast
- Slide 36 has PMs and demand manager chair · but not in slide before when speaking about triage

### Slide 41 · Team Kanban Dual Track
- **Discovery side won't be us — never as is**
- Sometimes through vendors
- We will be almost always the **orange band**
- Not reflected · how do we connect both worlds

---

## Status of feedback in current TOM (final.pptx · v3)

### ✅ Addressed
- NO OKRs · use Heather + Vitor metrics
- Middle managers in portfolio (Kumar's + Carlos's)
- Heather as Product Lead (Director · NEW capability)
- Discovery NOT us · orange = us · vendor handoff
- Pain Score 2 blocks · CS no Safety · ITBA Safety isolated 25%
- Type taxonomy: Production Support · Regulatory · Non-Reg only (no Change)
- XL → PPM Now
- Couple to existing BRPs (P&T BRP + Digital CoE BRP) · no own BRP
- Rolling wave 3-6-12
- Outcomes not deliverables (Heather mentors)
- Differentiate commits vs hypotheses

### ☐ Pending (to address in P&P v4)
- **Effort 1-5 scale defined** · Architect responsible explicit (Alfredo s.21)
- **Domain product owner / intake lead / business process owner** — explain or remove (Ana s.4)
- **BRS as new governance practice** (Ana s.29)
- **CoS responsibility to middle managers** explicit (Ana s.18)
- **Capacity planning assignment process** drawn (Alfredo s.19)
- **Light discovery artefacts + "from who"** (Ana s.3, s.23)
- **Cost of Delay** in prioritization (Ana s.1)
- **Multi-OPCO · OPCO duplicate handling** (Alfredo s.7)
- **Standardize Gate names per Type** (Alfredo s.15)
- **Performance metrics defined** (Ana s.4)
- **Production Support clarify (not = incidents)** (Alfredo s.13, s.15)

---

## QA cross-reference for v4 build

Each P&P slide must:
1. Not contradict any addressed item above
2. Address at least one pending item where applicable
3. Maintain consistency with intake taxonomies (formulas · roles · cadences · tools)
