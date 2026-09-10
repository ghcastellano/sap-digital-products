# DGTLPROD · Apply-Ready Payloads

Pasta com os artefatos prontos pra aplicar em `avangridjira.atlassian.net` para
provisionar o projeto **DGTLPROD** (Digital Products) conforme TOM.

Todos os JSONs seguem o formato esperado pelos endpoints REST do Jira Cloud.
Comentários estão em chaves `_comment` / `_hint_*` / `_note` — remova antes
de enviar pra API se sua biblioteca recusar campos extras.

---

## Ordem de aplicação

| # | Arquivo | Endpoint | Pré-requisito |
|---|---|---|---|
| 0 | (project shell) | `POST /rest/api/3/project` | Site admin permission |
| 1 | `custom_fields.json` | `POST /rest/api/3/field` × 8 | — |
| 2 | `field_options.json` | `POST /rest/api/3/field/{id}/context/{ctx}/option` | step 1 done · use returned field IDs |
| 3 | `issue_type_scheme.json` | `POST /rest/api/3/issuetypescheme` + associate | — |
| 4 | `workflow.json` | `POST /rest/api/3/workflows/create` + scheme + associate | — |
| 5 | `field_configurations.json` | `POST /rest/api/3/fieldconfiguration` × 6 + scheme + associate | step 1 done · IDs needed |
| 6 | `project_roles.json` | `POST /rest/api/3/role` × 8 + actor assignments | project DGTLPROD exists |
| 7 | `permission_scheme.json` | `POST /rest/api/3/permissionscheme` + associate | step 6 done |
| 8 | `components.json` | `POST /rest/api/3/component` × 10 | project DGTLPROD exists |
| 9 | `boards.json` | `POST /rest/api/3/filter` × 11 + `POST /rest/agile/1.0/board` × 11 | components done |
| 10 | `automations/A1..A5.json` | UI: Project → Automation → Import rule (one per file) | all fields + statuses exist |

NTT DATA tem scripts (`02_*.py` … `09_*.py`) prontos pra automatizar
steps 1-9. Step 10 (Automation rules) é UI-only na maioria das instâncias —
o admin importa manualmente cada arquivo da pasta `automations/`.

---

## Conteúdo · um por arquivo

### `custom_fields.json` — 8 campos novos
- `DP Value Stream Destination` · single-select
- `DP Sponsor` · user picker
- `DP Run / Grow / Transform` · single-select
- `DP Pain Score` · number (read-only · A4)
- `DP UX Score` · number 1-5
- `DP OPCO Multiplier` · number (read-only · A4)
- `DP Concept Prototype` · text/URL
- `DP Class of Service` · single-select

### `field_options.json` — opções dos selects
- 3 fields novos com options de criação
- 4 fields REUSE com options esperadas (verificar antes de adicionar)

### `issue_type_scheme.json` — 5 tipos
Todos REUSE (já existem globalmente). IDs:
- `10350` Demand Intake (default)
- `10006` Story · `10000` Epic · `10011` Bug · `10010` Sub-task

### `workflow.json` — Demand Lifecycle Workflow
- 11 statuses (8 main · NEW → ROUTED → TRIAGED → REVIEWING → ANALYZING → BACKLOG → IMPLEMENTING → DONE · 3 side · BOUNCED · DEFERRED · REJECTED)
- 13 transitions com hints de conditions/validators

### `field_configurations.json` — 6 configurations
- `FC_FUNNEL` (NEW) · `FC_AUTO_ONLY` (ROUTED · TRIAGED) · `FC_REVIEW` (REVIEWING) · `FC_ANALYZE` (ANALYZING) · `FC_DELIVERY` (BACKLOG · IMPLEMENTING · DONE) · `FC_TERMINAL` (DEFERRED · REJECTED · BOUNCED)
- Required / Hidden por campo · por status

### `project_roles.json` — 8 roles
`dp_requester` · `dp_pm_per_vs` · `dp_architect` · `dp_director_cs` · `dp_director_itba` · `dp_vp` · `dp_squad_lead` · `dp_admin`. Inclui actors iniciais por email (TBD onde lista de PMs ainda não foi consolidada).

### `permission_scheme.json` — DP Permission Scheme
12 permissions (BROWSE, CREATE, EDIT, TRANSITION, ASSIGN, etc.) mapeadas pros 8 roles.

### `components.json` — 10 components
Um por Value Stream (5 ITBA + 4 CS DVS + 1 Enabling Team).

### `boards.json` — 1 Portfolio + 10 Squad
- 11 saved filters (JQL com ORDER BY Pain Score DESC)
- Portfolio: 6 colunas · swim lanes por CoS · WIP {Expedite 1 · Fixed-Date 2 · Standard 4 · Intangible 1}
- 10 Squad boards: 4 colunas · mesma swim-lane template

### `automations/A1..A5.json` — 5 regras
- A1 · Auto-Routing (NEW → ROUTED + componente + notify Lead PM)
- A2 · Auto-Expedite (Customer Impact P1/P2 → CoS=Expedite)
- A3 · Auto-Fixed-Date (Demand Type=Regulatory → CoS=Fixed-Date)
- A4 · Auto-Pain-Score (computa Pain Score · fórmula ITBA vs CS)
- A5 · Auto-Convert (TRIAGED + Size → Story OU Epic+3 Stories)

A6 (Aging Alerts) foi descartado — Jira já tem SLA nativo que cobre.

---

## Caminho mais curto (Path A · clone PPR)

Se o admin clonar `PPR` em `DGTLPROD`, **vários artefatos já vêm prontos**:
- ✅ Issue type `Demand Intake` já está no scheme
- ✅ Workflow PPR é quase idêntico ao Demand Lifecycle (renomear statuses para bater)
- ❌ Components precisam ser criados (PPR tem outros)
- ❌ Custom fields novos precisam ser criados
- ❌ Field configurations específicas precisam ser criadas
- ❌ Boards precisam ser configurados
- ❌ Automation rules precisam ser importadas

Ou seja: clonar PPR poupa ~30% (issue types + workflow shell). O resto (~70%)
ainda precisa dos artefatos desta pasta — mas pode ser executado por NTT DATA
via API uma vez que `gustavo.castellano@avangrid.com` tenha role
`Project Administrator` em DGTLPROD.

---

## Validação (após aplicação)

Smoke test em `tests/smoke_test.py` (a escrever) deve criar 5 demands cobrindo:

| # | VS | OpCo | Size | Type | CIT | Resultado esperado |
|---|---|---|---|---|---|---|
| 1 | Billing Operations | UIL,CMP,NY | M | Non-Regulatory | P3 | A5 → Epic + 3 Stories · CoS=Standard · Backlog after Gate 2 |
| 2 | Corporate & Regulatory | All 5 | L | Regulatory | — | A3 → CoS=Fixed-Date · Backlog with deadline |
| 3 | Customer Care & Engagement | UIL | S | Non-Regulatory | P2 | A2 → CoS=Expedite · WIP-1 · Story direct to squad |
| 4 | Field Operations (linked) | UIL | S | Non-Regulatory | — | Story under existing Epic via shortcut · Squad PM accept |
| 5 | Iberdrola Corp Services | Iberdrola | M | Production Support | P4 | A5 → Epic · CoS=Intangible (tech debt) |

Critério de pass: 5 issues criadas, A1 dispara para todas, A2/A3 disparam nas
#3 e #2, A4 calcula Pain Score correto após scoring, A5 converte na transição
para TRIAGED, cards aparecem no board correto na lane de CoS correta.
