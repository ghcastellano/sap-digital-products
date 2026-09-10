# Admin Handoff · DGTLPROD Jira Provisioning

**Audience:** Avangrid Jira Cloud administrators (`avangridjira.atlassian.net`)
**Goal:** Stand up project `DGTLPROD` (Digital Products) per Avangrid Digital Products TOM designed by NTT DATA.
**Requester:** GUSTAVO CASTELLANO · `gustavo.castellano@avangrid.com` · accountId `712020:64921597-83d9-4a9d-81d4-42c67f900fdf`
**Companion document:** `SERVICENOW_REQUEST.md` (the ticket body that references this spec)
**Apply-ready artifacts:** `payloads/` directory — JSON payloads for every artifact below, ready for the Jira REST API or UI import. See `payloads/README.md` for the recommended order and the matching endpoint.

---

## TL;DR for the admin

Two viable execution paths. Either is acceptable.

**Path A · Clone PPR (fastest · recommended)**
PPR ("Pilot Project - Renewables") already uses the `Demand Intake` issue
type and a workflow nearly identical to the TOM Demand Lifecycle. Cloning
PPR into `DGTLPROD` saves ~80% of the work.

1. Clone project `PPR` → new key `DGTLPROD`, name `Digital Products`
2. Add requester as **Project Administrator**
3. Hand back to NTT DATA — they finish the remaining config (components,
   field mapping, automations, boards) via REST API.

**Path B · Build from scratch**
Use the full spec in section "Full provisioning spec" below. NTT DATA has
provisioning scripts ready (Phases 2–9) and only needs the empty project
+ admin role to run them.

Either way, the **only mandatory step Avangrid IT must perform** is the
project creation + admin grant. Everything else can be done by the
requester via API once they have project-admin on DGTLPROD.

---

## Why a single project

- Pain Score is computed and ranked **across** Value Streams — needs one
  shared backlog.
- 5 cross-VS automations (A1–A5) operate on a single rule scope.
- Reporting (Gate 2 portfolio committee) consumes one queue.

Components handle Value Stream routing without project sprawl.

---

## Full provisioning spec

### 1 · Project shell

| Property | Value |
|---|---|
| Key | `DGTLPROD` |
| Name | `Digital Products` |
| Description | Avangrid Digital Products demand intake → triage → portfolio → squad delivery (NTT DATA TOM) |
| Project lead | GUSTAVO CASTELLANO (rotates to Avangrid lead post-handover) |
| Type | Company-managed (classic) — to match instance convention |
| Template | Kanban (Software) · clone from `PPR` if Path A |
| URL | `https://avangridjira.atlassian.net/jira/software/projects/DGTLPROD` |

### 2 · Components (10 · one per Value Stream)

| # | Component name | Area | Director |
|---|---|---|---|
| 1 | Field Operations | ITBA | Carlos Garbin |
| 2 | Engineering & Asset Mgmt | ITBA | Carlos Garbin |
| 3 | Network Operations & Control | ITBA | Carlos Garbin |
| 4 | Corporate & Regulatory | ITBA | Carlos Garbin |
| 5 | Iberdrola Corp Services | ITBA | Carlos Garbin |
| 6 | Meter & Device Operations | CS | Kumar Jaladanki |
| 7 | Billing Operations | CS | Kumar Jaladanki |
| 8 | Revenue & Collections | CS | Kumar Jaladanki |
| 9 | Customer Care & Engagement | CS | Heather LaJoie |
| 10 | Enabling Team — Foundation | CS | Kumar Jaladanki |

(Component leads can be assigned later — not blocking provisioning.)

### 3 · Issue types (5 · all already exist globally · REUSE)

| Type | Global ID | Use |
|---|---|---|
| Demand Intake | `10350` | Placeholder created at intake form submit · auto-converts at Triage |
| Story | `10006` | S-sized work · Squad-level |
| Epic | `10000` | M/L-sized work · Portfolio-level |
| Bug | `10011` | Internal QA defects (NOT incidents — those stay in ServiceNow) |
| Sub-task | `10010` | Technical breakdown within a Story |

XL Initiatives are intentionally OUT of scope (handled in PPM Now per
Ana feedback).

### 4 · Custom fields

#### 4a · REUSE (10 fields · already exist on the instance · no new field IDs)

| Need | Field name | Field ID |
|---|---|---|
| OPCO | OpCo | `customfield_10800` |
| Business Value (narrative) | Business Value | `customfield_11440` |
| Acceptance Criteria | Acceptance Criteria | `customfield_10795` |
| Customer Impact Tier | Customer Impact | `customfield_11061` |
| Demand Type | Demand Type | `customfield_11682` |
| T-Shirt Size | T-Shirt-Sizing | `customfield_10850` |
| Safety score | Safety | `customfield_11438` |
| Effort score | Effort | `customfield_11282` |
| Discovery report | Discovery | `customfield_11646` |
| Epic Link | Epic Link | `customfield_10014` |

**Action by admin:** add these 10 fields to the DGTLPROD default screen
(or Path A inherits them automatically from PPR).

#### 4b · CREATE (8 net-new fields)

Group: `DP Custom Fields`. Naming prefix `dp_` so they're easy to find.

| Field key (suggested) | Type | Options / Notes |
|---|---|---|
| `dp_value_stream_destination` | Single-select | 10 options matching the components list above |
| `dp_sponsor` | User picker | Mandatory at intake |
| `dp_run_grow_transform` | Single-select | Run · Grow · Transform |
| `dp_pain_score` | Number (read-only) | Auto-computed by automation A4 |
| `dp_ux_score` | Number 1–5 | Scored by PM at Analyzing |
| `dp_opco_multiplier` | Number (read-only) | 1.0 / 1.3 / 1.6 — auto by A4 |
| `dp_concept_prototype` | Attachment | Required for Transform-class items |
| `dp_class_of_service` | Single-select | Standard · Fixed-Date · Expedite · Intangible |

### 5 · Workflow `Demand Lifecycle Workflow`

Statuses (8 main + 3 side states):

```
[NEW] ─A1─▶ [ROUTED] ─Gate 1─▶ [TRIAGED] ──┬──▶ [REVIEWING] ──▶ [ANALYZING] ──Gate 2──┬──▶ [BACKLOG] ──▶ [IMPLEMENTING] ──▶ [DONE]
                                            │                                          ├──▶ [DEFERRED]
                                            └──▶ [BOUNCED]                              └──▶ [REJECTED]
```

| Status | Category | SLA |
|---|---|---|
| NEW | To Do | — |
| ROUTED | To Do | A1 fires immediately on create |
| TRIAGED | In Progress | Weekly Gate 1 (Lead PM + Architect) |
| BOUNCED | To Do | Returned to requester for clarification |
| REVIEWING | In Progress | 2-week PM SLA |
| ANALYZING | In Progress | Architect feasibility + scoring |
| BACKLOG | In Progress | Ready for execution |
| IMPLEMENTING | In Progress | Squad pulls |
| DONE | Done | — |
| DEFERRED | Done | Out of scope this cycle |
| REJECTED | Done | Killed at Gate 2 |

**Special transitions:**
- "Demand on existing Initiative" shortcut: Demand → Story directly,
  skipping Reviewing/Analyzing/Gate 2 (validated by Squad PM accept).

**If cloning PPR (Path A):** PPR already has a near-identical workflow.
Map PPR statuses to these names if they differ; do not invent new ones.

### 6 · Field configurations per status (conditional visibility)

Six configurations, mapped via Field Configuration Scheme:

| Config | Status(es) | Editable | Hidden | Editable by role |
|---|---|---|---|---|
| `FC_FUNNEL` | NEW | Title · Description · Business Value · Acceptance Criteria · VS Destination · OpCo · Sponsor · Customer Impact (optional) | Type · Size · CoS · R/G/T · scores · Pain Score | dp_requester |
| `FC_AUTO_ONLY` | ROUTED · TRIAGED | none (read-only) | All Reviewing/Analyzing fields | (automation only) |
| `FC_REVIEW` | REVIEWING | Type · Size · CoS · R/G/T | Scores · Pain Score | dp_pm_per_vs |
| `FC_ANALYZE` | ANALYZING | BV/UX/Effort scores · Discovery · Concept Prototype | Pain Score (read-only · A4 writes) | dp_pm_per_vs + dp_architect |
| `FC_DELIVERY` | BACKLOG · IMPLEMENTING · DONE | progress comments · status | — | dp_squad_lead |
| `FC_TERMINAL` | DEFERRED · REJECTED · BOUNCED | reason field only | everything else | dp_director_* / dp_vp |

### 7 · Project roles + permission scheme

Eight roles to create on DGTLPROD:

| Role | Purpose |
|---|---|
| `dp_requester` | Anyone in Avangrid · can submit demands |
| `dp_pm_per_vs` | PMs assigned per Value Stream (10 sub-assignments) |
| `dp_architect` | Cross-VS architects · score Safety + Effort |
| `dp_director_cs` | Kumar Jaladanki · Heather LaJoie |
| `dp_director_itba` | Carlos Garbin |
| `dp_vp` | Vitor Cruz · chairs Gate 2 |
| `dp_squad_lead` | Per Value Stream — pulls from Backlog |
| `dp_admin` | NTT DATA delivery + Avangrid product owner (post-handover) |

Permission scheme `DP Permission Scheme` enforces the editable/hidden
matrix above by mapping field-level edit permissions to these roles.

### 8 · Boards (1 Portfolio + 10 Squad)

**Portfolio Overview Board** — kanban
- Filter: project = DGTLPROD AND status in (TRIAGED, REVIEWING, ANALYZING, BACKLOG, IMPLEMENTING, DONE)
- Columns: TRIAGED · REVIEWING · ANALYZING · BACKLOG · IMPLEMENTING · DONE
- Swim lanes: by `dp_class_of_service` (Standard · Fixed-Date · Expedite · Intangible)

**10 Squad Boards** — one per Value Stream
- Filter: project = DGTLPROD AND component = `[VS name]` AND status != NEW AND status != ROUTED
- Columns: BACKLOG · IN PROGRESS (= IMPLEMENTING) · REVIEW · DONE
- Swim lanes: by `dp_class_of_service`
- WIP limits per lane:
  - Expedite — 1
  - Fixed-Date — 2
  - Standard — 3–4
  - Intangible — 1

### 9 · Automations (5 rules)

| ID | Name | Trigger | Action |
|---|---|---|---|
| A1 | Auto-Routing | Issue created with Type=Demand Intake | Read VS Destination + OpCo → set Component → @notify Lead PM of that VS · transition NEW → ROUTED |
| A2 | Auto-Expedite | Customer Impact Tier set to P1 OR P2 | Set CoS = Expedite · email Director + Sponsor |
| A3 | Auto-Fixed-Date | Demand Type set to Regulatory | Set CoS = Fixed-Date · prompt PM to set due date |
| A4 | Auto-Pain-Score | BV + UX + Effort populated (Safety conditional on VS area = ITBA) | Compute and write `dp_pain_score`. **ITBA:** `(BV·OPCOmult·0.45 + UX·0.30 + Safety·0.25) / Effort`. **CS:** `(BV·OPCOmult·0.40 + UX·0.60) / Effort`. OPCO multiplier: 1.0 single OpCo · 1.3 (2–4 OpCos) · 1.6 (all 5+) |
| A5 | Auto-Convert at Triage | Status changes to TRIAGED + Size + Demand Type populated | Size=S → create Story (component-routed). Size=M/L → create Epic + 3 placeholder Stories |

A6 (Aging Alerts) was dropped — covered by Jira's native SLA tracking.

### 10 · Out of scope (explicit)

- No ServiceNow ↔ Jira webhook, integration, or scripted bridge.
  All incidents stay in ServiceNow.
- No Jira Service Management. Jira Software is sufficient.
- No XL Initiative type — handled in PPM Now.

---

## Acceptance / handover checklist

Mark each ✓ as it is completed.

- [ ] Project `DGTLPROD` exists
- [ ] Requester `gustavo.castellano@avangrid.com` is in `Administrators`
      role of DGTLPROD
- [ ] API smoke: `curl -u email:token https://avangridjira.atlassian.net/rest/api/3/project/DGTLPROD` returns 200
- [ ] (Path A only) `Demand Intake` issue type is in DGTLPROD's issue type scheme
- [ ] Workflow `Demand Lifecycle Workflow` (or PPR-cloned equivalent) is associated with DGTLPROD
- [ ] Confirmation email sent to requester so NTT DATA can resume Phases 3–9 (`02_create_project.py` onward)

After hand-back, NTT DATA performs:

1. Components (10 · one per VS) — `POST /rest/api/3/component`
2. Custom-field creation (8 new) + screen mapping for the 10 reused fields
3. Field-configuration scheme (6 configs)
4. Permission scheme + 8 project roles
5. Boards (1 Portfolio + 10 Squad) with swim lanes + WIP limits
6. Automations A1–A5 via Atlassian Automation API
7. Smoke test (5 sample demands across S/M/L · CS/ITBA · CoS variants)

All of (1)–(7) are scripted in `/jira_setup/` and idempotent.

---

## Reference

- Engagement: NTT DATA · Avangrid Digital Products TOM
- Sponsors: Vitor Cruz (VP), Carlos Garbin (Director ITBA), Kumar Jaladanki (Director CS), Heather LaJoie (Director CS · CCC)
- Ticket: see `SERVICENOW_REQUEST.md`
- Reference projects on the same instance: `PPR` (clone source) · `PSPS` · `NWOW` (sandbox) · `NTP0` (sandbox)
