# Avangrid — Estrutura Organizacional & Project Landscape

> Visão consolidada a partir do KB (Org_Chart_Digital_Products · Digital Products Team Org · Interview Carlos · Interview Stephanie/Kumar · Interview Application Champions · BRP Q2 2026 P&S Networks IT)
>
> **Data:** Abril 2026
>
> **Propósito:** Resolver dúvidas sobre P&S × P&T, hierarquia, papel da Heather, e relação Digital CoE × P&S antes do desenho do TO-BE.

---

## 1. Glossário rápido — siglas que confundem

| Sigla | Expansão | Onde aparece | O que faz |
|---|---|---|---|
| **P&S** | **Procurement & Services** *(per "Vision 2029" — slide 8 do nWOW deck)* | Avangrid Networks (lado de **negócio**, não IT) | Business unit que lidera grandes programas de Networks com IT em parceria. Roda Big Room Planning (BRP) trimestral. |
| **P&T** | **Process & Technology** *(per BRP Q2 2026 deck — Avangrid Networks)* | Mesmo grupo / lado IT | É o **time de IT que serve P&S**. Atua como **"intermediário com o business"** e como **"face — helping prioritize the work"** (interviews Michael Dubois, George Porto). |
| **Digital CoE** | Center of Excellence — Customer Experience & Digital | Sob Customer Service Products | Camada estratégica que **define padrões de digital products** transversalmente. Não tem hierarquia direta sobre os times de IT — coordena. |
| **Digital Products** | Programa de Transformação | Iniciativa (não é unidade) | É **o nosso projeto** — fundir SAP CoE + Applications em uma única organização. |
| **BRS** | Business Request / Support | Por OPCO | Função business-side que recebe demanda das OPCOs e leva pra IT. **3 grupos regionais**: BRS UIL (CT) · BRS CMP (ME) · BRS NY (New York). |
| **PDD** | Project Demand & Delivery *(processo)* | Avangrid-wide | Processo formal anual para **large projects**. Universal entre todos os grupos (GIS, Corporate, Global SAP, Smart Metering, Customer Services). |
| **SWF** | Software Factory | Shared service interno (sob Carlos) | Time interno Avangrid que **gerencia capacidade dos vendors** (principalmente Accenture). NÃO é vendor — é o orquestrador. |

> ⚠ **Confusão comum:** o KB tem variantes ambíguas — *"P&S (Procurement & Supply / Planning & Strategy)"* aparece em alguns slides. A leitura **mais autoritativa** (Vision 2029) é **Procurement & Services**. Recomendo confirmar com Vitor.

---

## 2. Hierarquia — IT Side (onde Digital Products atua)

```
                    ┌────────────────────────────────────┐
                    │  CIO — Edu de Frutos               │
                    └──────────────┬─────────────────────┘
                                   │
                    ┌──────────────┴─────────────────────┐
                    │  VP — IT Business Solutions Delivery│
                    │  Vitor Cruz                         │
                    └──────────────┬─────────────────────┘
                                   │
            ┌──────────────────────┼──────────────────────┐
            │                                             │
   ┌────────┴───────────┐                       ┌─────────┴───────────────┐
   │ Director —         │                       │ Director —              │
   │ IT Business Apps   │                       │ Customer Services Apps  │
   │ Carlos Garbin      │                       │ Kumar Jaladanki         │
   │ (4 sub-teams)      │                       │ (5 sub-teams · 2 vacant)│
   └────────────────────┘                       └─────────────────────────┘
```

🔴 **Working-session note (Miro)**: red sticky no box do Carlos — *"Attention here — the main opportunity is here to improve and standardize"*. **A branch do Carlos é o foco primário do TO-BE.**

---

### 2.1 Branch Carlos — IT Business Applications (4 sub-teams)

| Sr. Manager | Área | Champion entrevistado | FTE est. |
|---|---|---|---|
| **Bruno Moreto** | Electric & Gas Operations Products | Andrew McGibbon (G&E) | ~7 |
| **George Porto** | GIS & OSG | George Porto (mesmo) | ~6 |
| **Sarath Garikapati** | Corporate Applications | Sarath (Champion) | ~8 (5–6 PMs + 2 Tech Leads) |
| **Vanessa Colombo** | IT Apps Global SAP (Iberdrola) | Michael Dubois (sob Vanessa) | ~5 |

**Shared services sob Carlos:**
- **SWF & Tools** — Raquel Soto, Donna Henry
- **Dev Ops & App Security** — Piotr Wielogorski

**~75% são Sr. Manager** — confirmado no slide 7 do baseline (Applications card).

---

### 2.2 Branch Kumar — Customer Services Applications (5 sub-teams · 2 vacâncias)

| Sub-team | Sr. Manager | Champion | Status |
|---|---|---|---|
| **FICA, Billing & Collection** | Ivelisse Velazquez | (TBC) | Active |
| **Device & Metering** | **Vincent Lam** | **Iram Chowdhury** (PM, Smart Metering) | Active — único totalmente vendor-delivered |
| **CRM & Contact Center** | *(vacante)* | (TBC) | 🔴 **Vacant** |
| **SAP & CS Architecture** | *(vacante)* | (TBC) | 🔴 **Vacant** |
| **SAP Basis** | Nagi Reddy | — | Active |

**Sub-team Meter & Data Management** (sob Vincent Lam) — Manager: Leeton Salahuddin
Membros notáveis: Iram Chowdhury · Fabricio Silva · Le Phung *(retiring April)* · Susie Peters *(retiring)* · Lori Kirtland · Paulo Muniz *(Graduate)* · Simranjeet Singh · Satadal Gosh · Ray Wilkes · S. Nadendla · David Aliu · Kim Kendrick · S. Idiong · B. Kayam

**Sub-team CRM** — Evan Clements · Owais Majoka · Amy Vintenner · S. Nandam

---

### 2.3 Outras lideranças relevantes citadas no Org Chart

| Nome | Título | Relevância |
|---|---|---|
| **Heather LaJoie** | Sr. Manager — End User *(sob Customer Service Products, NÃO sob IT)* | Ver §4 abaixo |
| **Louise Murdock** | Director, Customer Services Applications *(papel sobreposto a Kumar — possivelmente matriz)* | Confirmar |
| **Luiz Sangiorgi** | Sr. Manager, Governance | Cross-cutting |
| **Vihn Nguyen** | Director, IT Program Delivery | Roda PDD |
| **Travis** *(sobrenome TBC)* | PMO Lead — Large Projects | Recebe demand grande de Customer Services |
| **Eric Krause** | Manager, Software Development Center | Cross |

---

## 3. Hierarquia — Business Side (onde P&S atua)

```
                    ┌──────────────────────────────────┐
                    │  Avangrid Networks (Business)     │
                    │  P&S Business Unit                │
                    │  (Procurement & Services)         │
                    └──────────────┬───────────────────┘
                                   │
                    ┌──────────────┴───────────────────┐
                    │  P&T — Process & Technology      │
                    │  (IT Networks team)              │
                    │  Roda BRP trimestral c/ business │
                    └──────────────────────────────────┘
```

> **P&S é o cliente**, **P&T é o IT que serve**. P&T atua como tradutor/intermediário com business — é o que aparece nas interviews quando dizem *"P&T is the face — helping prioritize the work"* (George Porto, GIS).

⚠ **Esses dois grupos são a referência de "como funciona bem"** — o nWOW deck do P&S (versão 1.2 · 03.10.2026) é o playbook de New Ways of Working que Digital Products quer escalar.

---

## 4. Heather — qual é o papel exato?

**Resposta curta:** Heather **NÃO é PM** dedicada de Customer Services nem de IT Business Apps. Ela é **Sr. Manager — End User** (Customer Service Products, lado business) e atua como **"first layer for prioritization" do Digital CoE** (per interview Stephanie/Kumar).

### Por que essa confusão é compreensível

Existem **três fluxos diferentes** onde a Heather aparece:

1. **Customer Service Products (lado business)** — onde ela está formalmente alocada como Sr. Manager — End User. Foca em experiência do usuário final / contact center.

2. **Digital CoE — first prioritization layer** — quando demanda chega aos **Digital PMs** (que são Functionals SAP do time do Kumar trabalhando em Big Room Planning), **Heather é quem faz o primeiro filtro de priorização** antes da demanda virar squad work. Citação direta:
   > *"Digital COE — Digital PMs (Functionals SAP from Kumar's team overseeing the work). Split: 30% BRS and Internal IT, 70% other. Digital PMs → Big Room Planning / Squads with visibility of what is coming. **Heather is the first layer for prioritization**."*
   > — *Interview Kumar Jaladanki + Stephanie*

3. **Stakeholder do nosso projeto** — Heather é **Executive Sponsor** com Vitor (a Tuesday-meeting fixa do plano TO-BE é **com ela**, justamente porque ela atravessa business + Digital CoE).

### Conclusão

Heather **não atende IT Business Applications** (Carlos branch) diretamente. Ela atua **majoritariamente sobre Customer Services + Digital CoE** — atravessa as duas pontes.

---

## 5. Digital CoE × P&S — como se relacionam?

**Resposta curta:** São **paralelos**, não hierárquicos. Compartilham **PRINCÍPIOS** de New Ways of Working (Big Room Planning, squads, value streams), mas operam sobre populações distintas.

| Eixo | Digital CoE | P&S Networks IT |
|---|---|---|
| **Lado** | Customer Services Apps + experiência digital | Avangrid Networks (utility infra) |
| **Foco** | SAP / Customer-facing (Billing, CRM, Smart Metering) | Operations & Engineering (grid, field, T&D) |
| **Dono** | Customer Service Products (lado business) + Kumar (lado IT) | P&S (business) + P&T (IT) |
| **Maturidade nWOW** | Em construção — Big Room Planning + squads recém-implementados | **Mais maduro** — operating model documented (nWOW v1.2 deck) |
| **Cadência** | Big Room Planning **(quando? — confirmar)** | BRP **trimestral** documentado (Q2 2026 deck) |
| **Heather participa?** | ✅ First prioritization layer | ❌ Não |

### Como se conectam na prática

- Ambos seguem o **mesmo playbook conceitual** (Product Discovery Playbook, OKRs, value streams)
- O **deck do P&S** é a referência que Digital Products quer adotar — 8 das melhores práticas vêm de lá
- Não há **squads compartilhados** entre os dois — populações de business e IT separadas
- **Stephanie + Kumar** atuam como ponte entre Digital CoE e Customer Services Apps

---

## 6. Demand & Project Landscape — quem corre o quê?

### 6.1 Onde nasce demanda

| Fonte | Tipo dominante | Canal | Vai para |
|---|---|---|---|
| Iberdrola Corporate Finance (Espanha) | Incidents (FI/PO/Budget) | IT Now | Vanessa / Global SAP |
| Customer Operations (US) | Billing/CRM tickets | IT Now / Email | Kumar branch |
| Field Services (meter installers) | Operations break/fix | IT Now | Smart Metering |
| BRS UIL · CMP · NY (3 OPCOs) | Mixed (regulatory, business) | Email + Jira + meetings | Various |
| **PDD process** (anual) | Large projects | SharePoint | Travis PMO + IT branch |
| Vendors (Accenture, third-parties) | Upgrades, releases | Email | G&E / Smart Metering / Global SAP |

### 6.2 Forums de governança documentados

| Forum | Cadência | Participantes | Decide |
|---|---|---|---|
| **Critical Integrations** | Weekly | TBD | Cross-team integration risks |
| **Thursday Accenture Review** | Weekly | Vanessa + Accenture leads | Iberdrola incident triage |
| **Smart Metering Internal** | Weekly | Iram + Vincent Lam | Vendor work prioritization |
| **PDD Governance** | Anual + checkpoints | Travis PMO + Directors | Large project approval |
| **Digital CoE BRP** | (cadência TBC) | Heather + Digital PMs | Squad priorities |
| **P&S BRP** | Trimestral | P&S business + P&T IT | Networks portfolio |

### 6.3 Ferramentas em uso (8 plataformas — slide 12 do baseline)

| Ferramenta | Para que | Quem usa |
|---|---|---|
| Jira | Small/medium demand · vendor work | Todos (heterogêneo) |
| IT Now | Incidents (RITM = minor) | Todos |
| SharePoint | Large projects (PDD) | Global SAP, Customer Services |
| PPM Now | Portfolio (large) | **Apenas Smart Metering** |
| Confluence | Documentation, portfolio doc | Smart Metering, Corporate |
| Word docs | Business requirements | Corporate (digital trail break) |
| OneNote + Teams | Informal PM | GIS |
| Email | Default for G&E + small CS demand | G&E + sub-teams CS |

---

## 7. Quem entrevistamos × quem falta

### Já entrevistados (KB tem transcripts)

| Pessoa | Papel | Branch |
|---|---|---|
| Carlos Garbin | Director IT Business Apps | IT Apps |
| Kumar Jaladanki | Director Customer Services Apps | CS Apps |
| Stephanie *(sobrenome TBC)* | Business counterpart Kumar | Business |
| George Porto | Sr. Manager GIS & OSG | IT Apps (Carlos) |
| Andrew McGibbon | E&G Champion (sob Bruno Moreto) | IT Apps (Carlos) |
| Sarath Garikapati | Sr. Manager Corporate Apps | IT Apps (Carlos) |
| Michael Dubois | Global SAP (sob Vanessa Colombo) | IT Apps (Carlos) |
| Iram Chowdhury | Smart Metering PM (sob Vincent Lam) | CS Apps (Kumar) |

### Faltam entrevistar (críticos para fechar gaps)

| Pessoa | Por quê |
|---|---|
| **Heather LaJoie** | First-prioritization layer Digital CoE — não temos transcript |
| **Vincent Lam** | Sr. Manager Device & Metering — só temos Iram |
| **Vanessa Colombo** | Sr. Manager Global SAP — só temos Michael (subordinado) |
| **Bruno Moreto** | Sr. Manager E&G — só temos Andrew (subordinado) |
| **Louise Murdock** | Director CS Apps (sobreposição com Kumar — esclarecer) |
| **Travis** | PMO Lead — recebe muita demand mas sem voz no baseline |
| **Vihn Nguyen** | Director IT Program Delivery — toca PDD direto |
| **Sr. Manager FICA/Billing** (Ivelisse Velazquez) | Único sub-team de CS sem voz |

---

## 8. Riscos / pontos de atenção para o TO-BE

1. **2 vacâncias na branch Kumar** (CRM & Contact Center · SAP & CS Architecture) — qualquer modelo TO-BE precisa **assumir as vagas como preenchidas** antes de ser operável. Confirmar timeline com Vitor.

2. **Sobreposição Kumar × Louise Murdock** — KB tem ambos como "Director Customer Services Apps" em momentos diferentes. Pode ser transição ou matriz — **precisa esclarecer**.

3. **P&S × Digital CoE × Digital Products** — três rótulos diferentes para frentes que **se sobrepõem em playbook mas não em pessoas**. Recomendo no slide deixar isso explícito (não fundir as 3 caixas).

4. **Heather como single point** — se ela é a única "first layer of prioritization" do Digital CoE, isso é um risco operacional. Vale validar se há backup.

5. **PDD é universal mas opaco** — todos mencionam PDD como o gate de large projects, mas nenhuma interview explica o critério de aprovação. Importante para Prioritization TO-BE.

---

## 9. Próxima ação recomendada

Antes de fechar o TO-BE Lifecycle Map (próxima semana), agendar **30 min com Vitor** para validar:

1. ✅ P&S = Procurement & Services? (vs Planning & Strategy)
2. ✅ Heather: papel formal × papel real (qual é o único?)
3. ✅ Louise Murdock × Kumar: transição ou matriz?
4. ✅ Timeline para preenchimento das 2 vacâncias na branch CS
5. ✅ Travis (PMO) participa do TO-BE design ou só consome?

---

*Documento gerado a partir de busca KB — 2026-04-25 · Fontes: Org_Chart_Digital_Products.md · Digital Products Team Org - March 2026 - Share.pptx · Interview_Carlos_Applications.md · Interview_Stephanie_Kumar_IT_CS.md · Interview_Application_Champions.md · BRP_Q2_2026_P&S Networks IT v1.4.4 · nWOW - New Ways of Working - P&S Lean Operating Model v1.2*
