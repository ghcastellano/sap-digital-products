# ServiceNow Request — Paste-Ready Message

> Cole o **Subject** no campo de título do ticket. Cole o **Description** no
> corpo. Anexe `ADMIN_HANDOFF.md` ao ticket. Antes de enviar, preencha o
> placeholder `[NTT DATA engagement lead]` no final.

---

## Subject (paste in title field)

```
Provision new Jira Cloud project DGTLPROD (Digital Products) and grant project-admin role to gustavo.castellano@avangrid.com
```

---

## Category / fields (if SN form asks)

- **Category:** Application Support
- **Subcategory:** Atlassian / Jira Cloud
- **Configuration item:** `avangridjira.atlassian.net`
- **Priority / Urgency:** Medium · target within 5 business days

---

## Description (paste in description / details field)

Hi Avangrid Jira Cloud Admin team,

I am requesting the provisioning of a new Jira Cloud project on `avangridjira.atlassian.net` to support the Avangrid Digital Products Target Operating Model (TOM) designed by NTT DATA.

**Requester:** Gustavo Castellano · `gustavo.castellano@avangrid.com` · Atlassian accountId `712020:64921597-83d9-4a9d-81d4-42c67f900fdf`. I am the NTT DATA consultant on the engagement.

**Business sponsor:** Vitor Cruz (VP, Digital Products). **Stakeholders:** Carlos Garbin (Director, ITBA), Kumar Jaladanki (Director, CS), Heather LaJoie (Director, CS · CCC).

**Background.** Avangrid is consolidating Customer Services Applications and IT Business Applications into a single Digital Products organization. NTT DATA has designed an end-to-end TOM covering demand intake → triage → portfolio governance → squad delivery, across 10 Value Streams and 5 OpCos (UIL, CMP, NY, Iberdrola, Cross-OpCo). The TOM has been approved by the Director group and is documented in the 98-slide workbook "TOM MAP workbook - final.pptx". Operationalizing the TOM in Jira requires a single project with a custom Demand Lifecycle workflow, 10 Value-Stream components, approximately 20 custom fields, 8 project roles, a board hierarchy with Class-of-Service swim lanes, and 5 Atlassian Automation rules. There is currently no equivalent project on the instance, and PPR (Pilot Project - Renewables) is the closest existing configuration.

**What I am requesting — three items:**

**1.** Create a new Jira Cloud project with key `DGTLPROD` and name `Digital Products`. Type: company-managed (classic), to match the rest of the Avangrid instance. The recommended approach is to **clone from project PPR** (Pilot Project - Renewables, lead: Marisa Rojas Gong) — PPR already includes the `Demand Intake` issue type and a workflow nearly identical to the TOM Demand Lifecycle, so cloning is the fastest provisioning path. If cloning is not preferred, an empty company-managed Software (Kanban) project is also acceptable and NTT DATA will configure the rest via API.

**2.** Grant the Project Administrator role on DGTLPROD to my account `gustavo.castellano@avangrid.com`. With this role I can configure components, custom-field-to-screen mappings, boards, swim lanes, WIP limits, and automations programmatically using the Jira REST and Atlassian Automation APIs. NTT DATA already has the provisioning scripts written, tested, and ready to run.

**3.** (Optional, time-saving.) Temporarily grant the global "Create next-gen projects" permission to my account for the duration of the setup. This would allow me to perform the project creation step myself, and the admin team can audit afterward. It can be revoked once DGTLPROD is verified in production.

**Explicitly out of scope** — please note, to avoid ambiguity:

- No integration between Jira and ServiceNow. All incidents continue to be managed 100% in ServiceNow. This request does not involve webhooks, scripted REST calls, MID-server connections, or any data flow between the two systems.
- No Jira Service Management licenses or features required. The existing Jira Software SKU is sufficient.
- No new user accounts. All participants already have Avangrid Atlassian identities.

**Reference projects on the same instance** (for context only — no changes needed there):

- `PPR` — Pilot Project - Renewables — recommended clone source
- `PSPS` — Process & Systems Portfolio Space — portfolio-governance reference
- `NWOW`, `NTP0` — sandboxes already used to validate the design

**Technical specification.** The full provisioning spec (issue types, custom fields, workflow, schemes, boards, automations, smoke-test acceptance) is attached to this ticket as `ADMIN_HANDOFF.md`. **Apply-ready JSON payloads** for every artifact (workflow definition, custom fields with options, components, field configurations, permission scheme, project roles, boards with swim lanes, and the 5 Atlassian Automation rules) are provided in the `payloads/` directory — formatted for the Jira REST API and Atlassian Automation rule-import. The admin team has two options: **Path A** — clone PPR into DGTLPROD, grant me admin, and stop there; NTT DATA applies the payloads via API. **Path B** — admin team applies the payloads end-to-end. **Path A is recommended.**

**Acceptance criteria** (how to know the ticket is done):

- The project DGTLPROD is reachable at `https://avangridjira.atlassian.net/jira/software/projects/DGTLPROD`.
- My account `gustavo.castellano@avangrid.com` appears in the `Administrators` role of DGTLPROD.
- The API call `GET https://avangridjira.atlassian.net/rest/api/3/project/DGTLPROD` returns HTTP 200 when authenticated with my API token.

**Why this is needed now.** The work is blocking Phase 3+ of the Jira rollout. Five directors and approximately 30 PMs and architects are waiting to onboard onto the new Demand Lifecycle. The next steering review is scheduled and the TOM smoke test (5 sample demands across S/M/L sizes, CS/ITBA areas, and all four Class-of-Service variants) cannot run until DGTLPROD exists. Target completion: within 5 business days of ticket open.

**Contact for clarifications.**

- Primary: Gustavo Castellano · `gustavo.castellano@avangrid.com`
- Avangrid sponsor: Vitor Cruz, VP Digital Products
- NTT DATA engagement lead: [NTT DATA engagement lead]

Thank you very much for your help — happy to jump on a quick call if any of the above needs clarification before provisioning.

Best regards,
Gustavo Castellano
NTT DATA · Avangrid Digital Products engagement
