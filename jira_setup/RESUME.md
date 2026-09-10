# RESUME · Avangrid Digital Products Jira Setup

**Last activity (2026-05-06):** NWOW sandbox probed — too restricted to prototype (no Demand Intake type, workflow mismatched, no REUSE custom fields on screens). Pivoted to **handoff path**: ServiceNow ticket + admin spec drafted.

**Active artifacts:**
- `SERVICENOW_REQUEST.md` — body to paste into the SN ticket
- `ADMIN_HANDOFF.md` — full technical spec (attach to ticket)
- `.checkpoints/03_probe_nwow.json` — evidence that NWOW is not a viable sandbox

**Next step:** open the ServiceNow ticket → wait for admin to (a) create project DGTLPROD, (b) grant requester project-admin role → resume scripts from `02_create_project.py` onward (or Phase 3+ if admin used Path A · clone PPR).

---

## What's done

### ✅ Phase 1 · API connectivity validated

- Auth works at `https://avangridjira.atlassian.net`
- User: GUSTAVO CASTELLANO (`gustavo.castellano@avangrid.com`)
- accountId: `712020:64921597-83d9-4a9d-81d4-42c67f900fdf`
- Permissions: BROWSE_PROJECTS ✓ · CREATE_ISSUES ✓ · ADMINISTER_PROJECTS ✓ · ADMINISTER ✗ (NOT site admin)

### ✅ Phase 1b/1c · Probed Avangrid Jira

- 136 projects · 294 custom fields · 32 issue types globally
- Reference projects mapped (PPR, PSPS, NWOW, NTP0)
- Existing fields/types cataloged for reuse vs create-new

---

## What's blocked

### ❌ Phase 2 · Cannot create project DGTLPROD

```
HTTP 403 · "You must have global administrator or create next-gen
projects permission in order to create this project."
```

User has NO admin permission on any of 136 projects · cannot create new.

---

## To resume · 4 paths

### A · Request permission (recommended)

Email Avangrid Jira admin asking for one of:
- "Create next-gen projects" permission for `gustavo.castellano@avangrid.com`
- OR "Site Administrator" temporary access to provision DGTLPROD

Then run `python3 02_create_project.py`.

### B · Have admin create the project

Send admin a copy of `/Users/gustavohenriquecastellano/.claude/plans/elegant-exploring-hippo.md` (the implementation plan). They create DGTLPROD with team-managed kanban template + 10 components per VS. Then we resume from Phase 3.

### C · Clone PPR (Pilot Project · Renewables)

PPR has matching workflow (To Do → In Development → Approved for Prod → Training → Business Review → Ready for Intake Review → Ready for Planning → Waiting for Requester → Done) AND already uses "Demand Intake" issue type. Lead: MARISA ROJAS GONG. Ask admin to clone PPR config for DGTLPROD.

### D · Prototype in NWOW (sandbox)

User has CREATE+EDIT in NWOW (test space). Limited — can't change config, only validate flow logic by creating issues with existing fields.

---

## Files in this directory

```
.env                       # API creds (gitignored)
.gitignore                 # excludes .env + .checkpoints/
config.py                  # constants · 10 VSs · OPCO/TYPE/SIZE options
client.py                  # REST wrapper · Basic auth · checkpoints
01_validate_access.py      # ✅ run · auth confirmed
01b_probe_capabilities.py  # ✅ run · capabilities mapped
01c_inventory.py           # ✅ run · field/type reuse mapped
02_create_project.py       # ❌ 403 · awaiting permission
02b_find_admin_project.py  # ✅ run · user not admin anywhere
RESUME.md                  # this file
.checkpoints/              # idempotency state per phase
└── tests/                 # smoke test placeholder
```

---

## Reuse decisions captured

### REUSE (existing in Avangrid)

| Need | Field ID | Field name |
|---|---|---|
| OPCO | `customfield_10800` | OpCo |
| Business Value | `customfield_11440` | Business Value |
| Acceptance Criteria | `customfield_10795` | Acceptance Criteria |
| Customer Impact | `customfield_11061` | Customer Impact |
| Demand Type | `customfield_11682` | Demand Type |
| T-Shirt Size | `customfield_10850` | T-Shirt-Sizing |
| Safety | `customfield_11438` | Safety |
| Effort | `customfield_11282` | Effort |
| Discovery | `customfield_11646` | Discovery |
| Epic Link | `customfield_10014` | Epic Link |

### CREATE (no existing match)

- Value Stream Destination
- Sponsor (user picker)
- Run / Grow / Transform
- Pain Score (auto-computed)
- UX Score
- OPCO multiplier
- Concept Prototype
- Class of Service

### Issue types — REUSE all (already exist)

- Demand Intake (id=10350) → for Demand placeholder
- Story (10006), Epic (10000), Bug (10011), Sub-task (10010), Initiative (10078)

---

## Implementation plan reference

**Full plan:** `/Users/gustavohenriquecastellano/.claude/plans/elegant-exploring-hippo.md`

**Summary:**
- Project: DGTLPROD (Digital Products) · team-managed kanban
- 10 components · one per VS (5 ITBA + 4 CS DVS + 1 Enabling)
- 5 issue types (Demand Intake · Story · Epic · Bug · Sub-task)
- ~20 custom fields (10 reuse + ~10 create)
- 10-status workflow `Demand Lifecycle Workflow`
- 6 field configurations per status (visibility per role/stage)
- 8 project roles + permission scheme
- 1 Portfolio + 10 Squad boards · CoS swim lanes (Standard · Fixed-Date · Expedite · Intangible)
- 5 automations (A1-A5)
- Smoke test: 5 demands covering S/M/L · CS/ITBA · CoS variants

---

## Memory references

- `/Users/gustavohenriquecastellano/.claude/projects/-Users-gustavohenriquecastellano/memory/project_avangrid_digital_products_tom.md`
- `/Users/gustavohenriquecastellano/.claude/projects/-Users-gustavohenriquecastellano/memory/project_avangrid_jira_setup.md`
