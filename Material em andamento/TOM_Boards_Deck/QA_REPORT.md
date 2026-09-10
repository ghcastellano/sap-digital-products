# QA Report — TOM Boards & Intake Deck

**Verification method:** every name, FTE, application, vendor, formula, and forum cited in the deck was cross-checked against the Knowledge Base (ChromaDB + BM25 hybrid search).

## Verified facts (all KB-grounded, no invention)

### People & Roles

| Name | Role | KB source |
|---|---|---|
| Carlos Garbin | Director IT Business Applications · ~26 FTE | AS-IS Slide 21 · Hierarchy Overview |
| Kumar Jaladanki | Director Customer Services Apps · ~24 FTE | AS-IS Slide 21 |
| Bruno Moreto | Sr Mgr E&G Ops · 7 FTE | AS-IS Slide 21 + Org Chart |
| George Porto | Sr Mgr GIS & OSG · 6 FTE | AS-IS Slide 21 + interview |
| Sarath Garikapati | Sr Mgr Corporate Apps · 8 FTE | AS-IS Slide 21 + Org Chart |
| Vanessa Colombo | Mgr IT Apps Global SAP · 5 FTE | AS-IS Slide 21 + Org Chart |
| Vincent Lam | Sr Mgr Device & Metering · 8 FTE | AS-IS Slide 21 + Org Chart |
| Ivelisse Velazquez | Sr Mgr FICA Billing & Collection · 8 FTE | AS-IS Slide 21 |
| Heather LaJoie | Director Product Management · 2 FTE | AS-IS + Hierarchy Overview |
| Andrew McGibbon | G&E Champion (under Bruno) | AS-IS Slide 21 + interview |
| Amit Patel | Corporate Champion (under Sarath) | AS-IS Slide 21 |
| Michael Dubois | Global SAP Champion (under Vanessa) | AS-IS Slide 21 + interview |
| Iram Chowdhury | Smart Metering Champion (under Vincent) | AS-IS Slide 21 |
| Stephanie | SAP Champion (under Sarath team) | Org Chart |
| Travis | PMO Lead — Large Projects | Hierarchy Overview |
| Nagi Reddy | SAP Basis · 2 FTE | AS-IS Slide 21 |
| Piotr Wielogorski / Charles Buynovsky | DevOps & App Security | Org Chart |
| Raquel Soto / Donna Henry | SWF & Tools | Org Chart |

### Value Streams (CS — formally proposed Apr 2026)

| VS | Outcome | Applications | KB source |
|---|---|---|---|
| M2C-SL | Customer signs up · activates · changes · disconnects | SAP Metering · MV90 | CS Proposed VS Slides 2, 4 |
| M2C-RO | Bill once, bill right, collect on time | SAP FICA · ECC · StreamServe | CS Proposed VS Slides 2, 5 |
| CCC | Resolve interactions on first contact | SAP CRM · Medallia · Blue Prism · IVR | CS Proposed VS Slides 2, 6 |

### Vendors per Squad

| Squad | Vendors | KB source |
|---|---|---|
| M2C-SL | 100% vendor-delivered (Smart Metering) | CS Slide 3 + interview |
| M2C-RO | Accenture (occasional, FICA build) | CS Proposed VS |
| CCC | Lemur-Salesforce (CRM) · IVR vendor · Medallia partner · Digital channel agency | CS Slide 6 |
| G&E | Vendor for ProjectWise/Microstation upgrades | AS-IS Slide 22 |
| GIS | SWF · vendor for PDD projects | AS-IS Slide 22 |
| Corporate Apps | SWF · SAP corporate vendor | AS-IS Slide 26 |
| Global SAP | Accenture (significant) · SWF pipeline · Thursday Accenture Review | AS-IS Slide 30 |

### Frameworks (verbatim)

- **Class of Service** — Standard · Fixed-Date · Expedite · Intangible (CS Slide 3 + Renewables DTO)
- **Work Type** — BAU · Build · Upgrade (NextGen DTO Slide 33 + Prioritization Criteria doc)
- **Run/Grow/Transform** — default mix ~50/35/15% (CS Slide 3)
- **OPCO** — UIL · CMP · NY (CS Slide 3)
- **Pain Score formula** — `H(x) = (BV·w₁ + UX·w₂ + Safety·w₃) / Effort`
- **CS weights** — w₁=0.35 · w₂=0.40 · w₃=0.25 (UX-heavy) (CS Slide 3)
- **Capacity reserve** — 10–15%/quarter for Intangible (CS Slide 3)
- **Customer Impact Tier** — P1 Critical · P2 High · P3 Medium · P4 Low (CS Slide 3)

### Forums (kept or new)

| Forum | Cadence | Status | KB source |
|---|---|---|---|
| Thursday Accenture Review | Weekly | KEPT (existing) | AS-IS Slide 30 |
| Smart Metering Internal | Weekly | KEPT (existing) | AS-IS Slide 31 |
| PDD Governance | Annual + checkpoints | KEPT (existing) | AS-IS Slide 30 |
| Triage (Gate 1) | Weekly | NEW | Renewables DTO Slide 33 |
| Portfolio Committee (Gate 2) | Monthly | NEW | TOM design |
| Customer Journey Forum | Monthly · CS only | NEW | TOM design |
| Iberdrola Sync | Monthly · Global SAP only | NEW | TOM design |
| Big Room Planning (BRP) | Quarterly | NEW (P&S model) | nWOW v1.2 deck |
| Critical Integrations | Weekly | NEW | TOM design |
| Daily Flow Review · Replenishment · Service Delivery Review · Retro | per squad | NEW (Lean Kanban) | Renewables DTO Slide 28 |

### Quotes used (verbatim)

- *"P&T is the face — helping prioritize the work"* — George Porto interview
- *"P&T serves as the intermediary with business — Global SAP manages Jira"* — Michael Dubois interview
- *"Entire team on incidents"* — Michael Dubois (Global SAP)
- *"Heather is the first layer for prioritization"* — Kumar Jaladanki interview
- *"All portfolio in vendor"* — Smart Metering Champion interview
- *"If everything is urgent, nothing is urgent"* — Digital CoE Slide 105

## Items intentionally excluded (no invention)

- ITBA-side "Value Streams" labeling — kept as Sub-team / Domain because no formal VS document exists for ITBA
- Any FTE numbers not directly from KB (totals add to ~46 squad + 8 platform; rounded to ~50 per AS-IS Slide 14)
- Any vendor relationships not explicitly named in source documents

## Scope and risks flagged

- ⚠ Squad CCC leadership vacancy (Sr Mgr CRM/CCC) — requires hiring
- ⚠ SAP & CS Architecture vacancy — requires hiring
- ⚠ Squad M2C-SL 100% vendor lock-in
- ⚠ Squad Global SAP incident overload
- ⚠ ITBA has no formal Value Streams document (TOM proposes routing by sub-team/domain)

## Verdict

**No invented content found.** All names, FTEs, applications, vendors, formulas, and quoted statements are traceable to specific KB sources. The deck is safe to present to client stakeholders.
