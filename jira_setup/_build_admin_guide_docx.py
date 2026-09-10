"""Build the Word document for the Avangrid Jira admin team.

v2 — aligned with TOM MAP workbook - FINAL FINAL.pptx (May 2026)

Major updates vs v1:
  • 6 issue types (added INITIATIVE for L-sized work)
  • Outcome-named hierarchy: S → Story · M → Epic · L → Initiative
  • Initiative > 3 months breaks into Epics
  • Status "BACKLOG" renamed "READY FOR EXECUTION" (matches slide 107)
  • Effort field gets 1–5 scale (1=hours, 5=multi-quarter)
  • Two channels for unplanned demand: Scope-add + Severity/Emergency
  • INVEST + 6-criteria DoR documented
  • Customer Impact SLAs: P1=4h, P2=8h, P3=24h, P4=72h
  • A5 updated to map Size → correct issue type
  • Tie-break rule documented
  • XL → PPM Now handoff documented
"""
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = Path(__file__).parent / "Avangrid_Jira_DGTLPROD_Admin_Guide.docx"

NAVY  = RGBColor(0x1F, 0x38, 0x64)
ACCENT= RGBColor(0x2E, 0x75, 0xB6)
GREY  = RGBColor(0x59, 0x59, 0x59)


# ---------- helpers ----------------------------------------------------------

def set_cell_bg(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tc_pr.append(shd)


def heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = NAVY
        run.font.name = "Calibri"
    return h


def para(doc, text, bold=False, italic=False, size=11, color=None, align=None):
    p = doc.add_paragraph()
    if align:
        p.alignment = align
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = color
    return p


def bullet(doc, text, level=0):
    p = doc.add_paragraph(text, style="List Bullet")
    p.paragraph_format.left_indent = Cm(0.5 + 0.5 * level)
    for run in p.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(11)


def numbered(doc, text):
    p = doc.add_paragraph(text, style="List Number")
    for run in p.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(11)


def table(doc, headers, rows, widths_cm=None, header_bg="1F3864"):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Light Grid Accent 1"
    hdr = t.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = ""
        p = hdr[i].paragraphs[0]
        run = p.add_run(h)
        run.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        run.font.name = "Calibri"
        run.font.size = Pt(10)
        set_cell_bg(hdr[i], header_bg)
        hdr[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    for row in rows:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            run = p.add_run(str(val))
            run.font.name = "Calibri"
            run.font.size = Pt(10)
    if widths_cm:
        for i, w in enumerate(widths_cm):
            for row in t.rows:
                row.cells[i].width = Cm(w)
    return t


def callout(doc, text, color="DEEBF7"):
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cell = t.rows[0].cells[0]
    set_cell_bg(cell, color)
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(text)
    run.font.name = "Calibri"
    run.font.size = Pt(10)
    run.italic = True


def page_break(doc):
    doc.add_page_break()


def hrule(doc):
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:color'), '1F3864')
    pBdr.append(bottom)
    pPr.append(pBdr)


# ---------- document ---------------------------------------------------------

doc = Document()

style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(11)

# === Cover ===================================================================
para(doc, "AVANGRID DIGITAL PRODUCTS", bold=True, size=14, color=ACCENT,
     align=WD_ALIGN_PARAGRAPH.CENTER)
para(doc, "", size=4)
para(doc, "Jira Cloud Provisioning Guide", bold=True, size=28, color=NAVY,
     align=WD_ALIGN_PARAGRAPH.CENTER)
para(doc, "Project DGTLPROD — Admin Setup", bold=False, size=16, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER)
para(doc, "(aligned with TOM MAP workbook · FINAL FINAL · May 2026)",
     bold=False, italic=True, size=11, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER)
para(doc, "", size=20)
hrule(doc)
para(doc, "Prepared by:  NTT DATA · Avangrid Digital Products engagement",
     size=11, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)
para(doc, "Requester:    NTT consultant",
     size=11, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)
para(doc, "Sponsor:      Vitor Cruz, VP Digital Products",
     size=11, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)
para(doc, "Stakeholders: Carlos Garbin (ITBA) · Kumar Jaladanki (CS) · "
     "Heather LaJoie (CS · CCC)", size=11, color=GREY,
     align=WD_ALIGN_PARAGRAPH.CENTER)
para(doc, "Instance:     avangridjira.atlassian.net",
     size=11, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)
hrule(doc)
para(doc, "", size=14)
callout(doc,
        "Single hand-off package for the Avangrid Jira Cloud admin team. "
        "Section 3 explains how the system works for end users; Section 5 "
        "is the admin checklist (UI navigation only, no code or REST). "
        "Aligned with the canonical TOM MAP workbook FINAL FINAL slide deck "
        "approved by the Director group.")

page_break(doc)

# === 1. Executive summary ====================================================
heading(doc, "1 · Executive summary", 1)
para(doc,
     "Avangrid is consolidating Customer Services Applications (CS) and IT "
     "Business Applications (ITBA) into a single Digital Products "
     "organization. NTT DATA has designed an end-to-end Target Operating "
     "Model (TOM) covering demand intake, triage, portfolio governance "
     "and squad delivery, across 10 Value Streams and 5 OpCos.")
para(doc,
     "To operationalize the TOM, we need a single Jira Cloud project, "
     "DGTLPROD (\"Digital Products\"), with a custom workflow, Value-Stream "
     "components, ~20 custom fields, role-based permissions, kanban boards "
     "with Class-of-Service swim lanes, and 5 automation rules. There is "
     "currently no equivalent project on the instance.")
para(doc,
     "The Avangrid Jira admin team is the only role that can create the "
     "project shell and grant project-administrator access to the "
     "requester. Once that is done, NTT DATA can complete the remaining "
     "configuration, or the admin team may apply the entire specification "
     "in this document end-to-end.")

# === 2. What we are asking ===================================================
heading(doc, "2 · What we are asking the admin team to do", 1)
para(doc, "Three concrete actions:", bold=True)

numbered(doc,
    "Create a new Jira Cloud project. Key: DGTLPROD. Name: Digital "
    "Products. Type: Company-managed (classic), to match the rest of the "
    "Avangrid instance. Recommended approach: clone PPR (Pilot Project · "
    "Renewables) — see Section 4.")
numbered(doc,
    "Grant Project Administrator role on DGTLPROD to the NTT "
    "consultant. With this access, NTT DATA can configure the rest of "
    "the system without further admin involvement.")
numbered(doc,
    "(Optional, time-saving) Temporarily grant the global \"Create "
    "next-gen projects\" permission to the requester for the duration "
    "of setup, then revoke once DGTLPROD is verified.")

para(doc, "", size=4)
para(doc, "Out of scope (explicit):", bold=True)
bullet(doc, "No integration between Jira and ServiceNow. All incidents "
            "continue to be managed 100% in ServiceNow.")
bullet(doc, "No Jira Service Management licenses required. Jira Software "
            "(the existing Avangrid SKU) is sufficient.")
bullet(doc, "No new user accounts. All participants already have Avangrid "
            "Atlassian identities.")
bullet(doc, "No XL initiatives. XL flows through PPM Now (a separate "
            "Avangrid system) — the Demand Lifecycle handles S, M, and L.")

page_break(doc)

# === 3. End-user view ========================================================
heading(doc, "3 · End-user view — how the system will work", 1)
para(doc,
     "Before the technical configuration, this section describes what the "
     "system feels like for the people who use it.")

heading(doc, "3.1 · Roles", 2)
table(doc,
      ["Role", "Who", "What they do"],
      [
        ["Requester",       "Anyone in Avangrid",
         "Submits a new Demand using the intake form. Provides title, "
         "narrative business value, acceptance criteria, target Value "
         "Stream, OpCo(s), and a Sponsor."],
        ["Sponsor",         "Business owner named by the Requester",
         "Validates the demand, accepts at Done. Mandatory user-picker; "
         "Requester ≠ Sponsor."],
        ["Lead PM (per VS)","PM assigned to a Value Stream",
         "Reviews each demand at weekly Triage (Gate 1). At Reviewing, "
         "defines Type, Size (initial guess), Class of Service, "
         "Run/Grow/Transform. At Analyzing, scores BV and UX (1-5)."],
        ["Architect",       "Cross-VS architects",
         "At Analyzing, scores Safety (ITBA only) and Effort (1-5 t-shirt "
         "scale). Provides feasibility judgment."],
        ["Director",        "Carlos Garbin (ITBA), Kumar Jaladanki and "
                            "Heather LaJoie (CS)",
         "Decision-maker at Gate 2 (monthly Portfolio Committee). "
         "Approves, defers or rejects each demand based on auto-computed "
         "Pain Score. Can declare Severity at any time (slide 54)."],
        ["VP",              "Vitor Cruz",
         "Chairs Gate 2."],
        ["Squad PM / Lead", "Per VS",
         "Pulls approved work from Ready For Execution into Implementing, "
         "respecting WIP limits per Class of Service. Owns DoR for each "
         "Story (INVEST + 6-criteria DoR check)."],
      ],
      widths_cm=[3.0, 4.0, 10.0])

heading(doc, "3.2 · The demand journey", 2)
para(doc,
     "A demand moves through nine main stages, with three exit states "
     "(deferred, rejected, bounced back to the requester) plus two "
     "shortcut channels for unplanned work.")

table(doc,
      ["Stage", "Who acts", "What happens"],
      [
        ["NEW",          "Requester",
         "Form submitted. Title, description, business value, acceptance "
         "criteria, VS Destination, OpCo, Sponsor captured."],
        ["ROUTED",       "(automation A1)",
         "Demand auto-routed to the correct VS component. Lead PM "
         "notified."],
        ["TRIAGED",      "Lead PM (Gate 1, weekly · 30 min)",
         "Quality gate: keep going, bounce back, or close out. Confirms "
         "incoming demand has enough info to enter Reviewing."],
        ["REVIEWING",    "Lead PM",
         "PM defines Type (Production Support / Regulatory / "
         "Non-Regulatory), Size (S / M / L), Class of Service, "
         "Run/Grow/Transform. 2-week PM SLA."],
        ["ANALYZING",    "Lead PM + Architect",
         "PM scores Business Value and UX (1–5). Architect scores Safety "
         "(ITBA only) and Effort (1–5 t-shirt scale). Pain Score "
         "auto-computed by automation A4."],
        ["READY FOR EXECUTION", "(decision at Gate 2)",
         "Approved by Director / VP at monthly Portfolio Committee. "
         "Replaces the older \"Backlog\" naming (see slide 107). "
         "Awaiting capacity."],
        ["IMPLEMENTING", "Squad PM",
         "Squad pulls work, respecting WIP limits per Class of Service. "
         "Story-level DoR enforced before pull."],
        ["DONE",         "Squad PM + Sponsor",
         "Delivered, Sponsor accepts, value realized, post-implementation "
         "review."],
        ["DEFERRED",     "(Gate 2 decision)",
         "Out of scope this cycle. Re-evaluated at next Gate 2."],
        ["REJECTED",     "(Gate 2 decision)",
         "Will not be done. Reason captured in comment."],
        ["BOUNCED",      "Lead PM (at Triage)",
         "Returned to Requester for clarification before re-entering."],
      ],
      widths_cm=[3.5, 3.5, 10.0])

heading(doc, "3.3 · Two channels for unplanned demand", 2)
para(doc,
     "Beyond the standard funnel, two shortcut channels allow urgent or "
     "scope-additive work to bypass parts of the flow (TOM slide 54):")
bullet(doc, "Channel 1 · Scope-add shortcut — when a new ask scopes into an "
            "Initiative already in Implementing, Squad PM evaluates within "
            "48 h and may add a Story under the existing parent (DoR "
            "fast-track, inherits parent CoS, no new Gate 2 — audit logged "
            "at the parent Initiative).")
bullet(doc, "Channel 2 · Severity / Emergency — production outage, "
            "regulatory deadline shift, safety incident, or critical "
            "impact declared by a Director. Skips Funnel, goes direct to "
            "Squad. Lane = Expedite (severity / outage) or Fixed-Date "
            "(regulatory deadline). Director signs off; post-mortem at "
            "next Gate 2.")
callout(doc, "Pure incidents are not in scope here — they remain 100% in "
             "ServiceNow.")

heading(doc, "3.4 · Class of Service — the prioritization lens", 2)
para(doc,
     "Every demand carries one of four Classes of Service that drives "
     "swim-lane placement and WIP limits:")
bullet(doc, "Expedite — auto-assigned when Customer Impact = P1 or P2 "
            "(automation A2). Squad WIP limit 1.")
bullet(doc, "Fixed-Date — auto-assigned when Demand Type = Regulatory "
            "(automation A3). Squad WIP limit 2.")
bullet(doc, "Standard — default. Squad WIP limit 4.")
bullet(doc, "Intangible — for tech debt, refactoring, internal "
            "improvements. Squad WIP limit 1.")

heading(doc, "3.5 · Customer Impact Tier — auto-trigger SLAs", 2)
table(doc,
      ["Tier", "When to use", "Auto trigger", "SLA"],
      [
        ["P1 Critical",
         ">100k customers OR full OPCO outage · revenue impact >$100k/day",
         "A2 → CoS = Expedite · WIP-1",
         "4 hours"],
        ["P2 High",
         ">10k customers · degraded service · workarounds available",
         "A2 → CoS = Expedite · WIP-1",
         "8 hours"],
        ["P3 Medium",
         "Single-OPCO non-critical · workarounds in place",
         "No automation · Standard CoS",
         "24 hours"],
        ["P4 Low",
         "Internal-only · no customer impact",
         "No automation · Standard or Intangible CoS",
         "72 hours"],
      ],
      widths_cm=[2.5, 6.5, 5.5, 2.5])
callout(doc, "CIT does NOT enter the Pain Score formula — it is a "
             "categorical flag, not a numeric weight. CIT is CS-only "
             "context (no field crew dimension on ITBA).")

heading(doc, "3.6 · Pain Score — the prioritization formula", 2)
para(doc,
     "Pain Score is automatically computed once Business Value, UX, "
     "Effort (and Safety, for ITBA) are filled in. It ranks demands at "
     "Gate 2.")
bullet(doc, "ITBA — 3 components (BV + UX + Safety · BV-heavy):  "
            "Pain Score = ( BV × OPCO multiplier × 0.45 + UX × 0.30 + "
            "Safety × 0.25 ) ÷ Effort")
bullet(doc, "CS — 2 components (BV + UX · UX-heavy, no Safety):  "
            "Pain Score = ( BV × OPCO multiplier × 0.40 + UX × 0.60 ) "
            "÷ Effort")
para(doc, "OPCO multiplier:", bold=True)
bullet(doc, "1 OpCo  → 1.0")
bullet(doc, "2 to 4  → 1.3 (regional)")
bullet(doc, "All 5+  → 1.6 (enterprise)")
para(doc, "Effort is on a 1–5 t-shirt scale:", bold=True)
bullet(doc, "1 = hours")
bullet(doc, "2 = days")
bullet(doc, "3 = weeks")
bullet(doc, "4 = quarter")
bullet(doc, "5 = multi-quarter (signal to consider XL → PPM Now)")

heading(doc, "3.7 · Tie-break rule (Gate 2)", 2)
para(doc, "When two demands have an identical Pain Score, the tie-break "
          "is deterministic and audit-logged at Gate 2 (TOM slide 61):")
numbered(doc, "Class of Service lane (Expedite > Fixed-Date > Standard > "
              "Intangible).")
numbered(doc, "Capacity available — the one fitting the current capacity "
              "wins.")
numbered(doc, "Sponsor seniority — the one with a named Director Sponsor "
              "wins.")
callout(doc, "Never \"squeaky-wheel\". The rule is logged at Gate 2 for "
             "audit.")

page_break(doc)

# === 4. Faster path ==========================================================
heading(doc, "4 · Faster path · clone PPR", 1)
para(doc,
     "PPR (Pilot Project · Renewables, lead Marisa Rojas Gong) already "
     "uses the Demand Intake issue type and a workflow nearly identical "
     "to the one we need.")

table(doc,
      ["What clone gives us", "What still needs doing after the clone"],
      [
        ["Issue types Demand Intake, Story, Epic, Bug, Sub-task — "
         "associated to the project",
         "Issue type Initiative — must be added (now part of the "
         "outcome-named hierarchy, see Section 5.3)"],
        ["Workflow shell — most statuses match (some renaming may be "
         "needed; \"Backlog\" → \"Ready For Execution\")",
         "Components (10 Value Streams, listed in Section 5.2)"],
        ["Default screens",
         "Custom fields (8 new — Section 5.4) and their options"],
        ["",
         "Field configurations per status (Section 5.6)"],
        ["",
         "Project roles + permission scheme (Section 5.7)"],
        ["",
         "Field-level access matrix (Section 5.8)"],
        ["",
         "Boards (1 Portfolio + 10 Squad) and swim lanes (Section 5.9)"],
        ["",
         "Automation rules A1–A5 + prerequisites (Section 5.10)"],
      ],
      widths_cm=[8.5, 8.5])

para(doc,
     "If the admin team performs only the clone + grant of project-admin "
     "access, NTT DATA will configure everything else.")

page_break(doc)

# === 5. Configuration to provision ==========================================
heading(doc, "5 · Configuration to provision (Jira admin view)", 1)
para(doc, "Each subsection has the same structure:", italic=True, color=GREY)
bullet(doc, "What it is — short description.")
bullet(doc, "Where to do it — exact navigation in the Jira UI.")
bullet(doc, "What to put in — table or list of values.")

heading(doc, "5.1 · Project shell", 2)
para(doc, "What it is.", bold=True)
para(doc, "The DGTLPROD project itself.")
para(doc, "Where to do it.", bold=True)
para(doc, "Jira top bar → Projects → Create project. Or, recommended: "
          "Projects → … (more) on PPR → Copy project.")
para(doc, "What to put in.", bold=True)
table(doc, ["Property", "Value"], [
    ["Project name",   "Digital Products"],
    ["Project key",    "DGTLPROD"],
    ["Project type",   "Company-managed (classic)"],
    ["Template",       "Kanban (Software) — or inherited from PPR clone"],
    ["Project lead",   "NTT consultant (rotates to Avangrid lead "
                        "post-handover)"],
    ["Description",    "Avangrid Digital Products demand intake → triage "
                        "→ portfolio governance → squad delivery (NTT DATA "
                        "TOM v3 · FINAL FINAL)."],
], widths_cm=[5.0, 12.0])

heading(doc, "5.2 · Components — 10 Value Streams", 2)
para(doc, "What it is.", bold=True)
para(doc, "Components route demands to the correct Value Stream within "
          "the single project. Each demand carries exactly one component. "
          "Mapped 1:1 to the \"DP Value Stream Destination\" custom field.")
para(doc, "Where to do it.", bold=True)
para(doc, "Project DGTLPROD → Project settings → Components → Add "
          "component (×10).")
para(doc, "What to put in.", bold=True)
table(doc, ["#", "Component name", "Area", "Director / OVS owner"], [
    ["1",  "Field Operations",             "ITBA", "Carlos Garbin"],
    ["2",  "Engineering & Asset Mgmt",     "ITBA", "Carlos Garbin"],
    ["3",  "Network Operations & Control", "ITBA", "Carlos Garbin"],
    ["4",  "Corporate & Regulatory",       "ITBA", "Carlos Garbin"],
    ["5",  "Iberdrola Corp Services",      "ITBA", "Carlos Garbin"],
    ["6",  "Meter & Device Operations",    "CS",   "Kumar Jaladanki"],
    ["7",  "Billing Operations",           "CS",   "Kumar Jaladanki"],
    ["8",  "Revenue & Collections",        "CS",   "Kumar Jaladanki"],
    ["9",  "Customer Care & Engagement",   "CS",   "Heather LaJoie"],
    ["10", "Enabling Team — Foundation",   "CS",   "Kumar Jaladanki"],
], widths_cm=[1.0, 7.0, 2.5, 5.0])
callout(doc, "Component leads can be assigned later — not blocking "
             "provisioning.")

heading(doc, "5.3 · Issue types — six types (all already exist globally)", 2)
para(doc, "What it is.", bold=True)
para(doc,
     "DGTLPROD uses SIX issue types. All six already exist on the Avangrid "
     "instance — no new types are created. The admin only needs to "
     "associate the correct Issue Type Scheme to DGTLPROD.")
para(doc, "Where to do it.", bold=True)
para(doc, "Site administration → Issues → Issue type schemes → Add "
          "scheme \"DP Issue Type Scheme\" → assign to DGTLPROD.")
para(doc, "What to put in.", bold=True)
table(doc, ["Issue type (existing)", "Use", "Default?"], [
    ["Demand Intake", "Placeholder created at intake form submit; "
                      "auto-converts at Triage", "✓ default"],
    ["Initiative",    "L-sized work · default top-level for portfolio "
                      "items · sized for ≤ 3 months · sponsor-approved",
                      ""],
    ["Epic",          "M-sized work, OR exception when Initiative > 3 "
                      "months breaks into Epics", ""],
    ["Story",         "S-sized work, Squad-level · ≤ 5 SP · INVEST · "
                      "Squad PM owns DoR", ""],
    ["Bug",           "Internal QA defects (NOT incidents — those stay "
                      "in ServiceNow)", ""],
    ["Sub-task",      "Technical breakdown within a Story", ""],
], widths_cm=[3.5, 11.0, 2.0])

para(doc, "Outcome-named hierarchy (TOM slide 48):", bold=True)
table(doc, ["Size at Triage", "Issue type to materialize",
            "Owner of refinement"], [
    ["S",         "Story",       "Squad PM (DoR)"],
    ["M",         "Epic",        "PM per VS"],
    ["L",         "Initiative",  "PM per VS"],
    ["L > 3 months", "Initiative breaks into multiple Epics",
                  "PM per VS + Squad PMs"],
    ["XL",        "Routed to PPM Now (out of scope here)", "—"],
], widths_cm=[4.5, 8.5, 4.0])
callout(doc, "Default flow is INITIATIVE → STORY. The Epic level is "
             "an exception used only when an Initiative exceeds 3 months "
             "(slide 48). Squad PM verifies INVEST on every Story: "
             "Independent · Negotiable · Valuable · Estimable · Small "
             "(≤5 SP) · Testable.")

heading(doc, "5.4 · Custom fields", 2)

para(doc, "5.4.1 · Reuse existing Avangrid fields (no new fields needed)",
     bold=True)
para(doc, "Where to do it.", bold=True)
para(doc, "Project DGTLPROD → Project settings → Screens → add the "
          "fields below to the Default Screen (and to Create / Edit / "
          "View screens as appropriate).")
para(doc, "What to add.", bold=True)
table(doc, ["Need", "Existing field name", "Field ID for verification"], [
    ["OpCo",                     "OpCo",                "customfield_10800"],
    ["Business Value (narrative)","Business Value",      "customfield_11440"],
    ["Acceptance Criteria",      "Acceptance Criteria", "customfield_10795"],
    ["Customer Impact Tier",     "Customer Impact",     "customfield_11061"],
    ["Demand Type",              "Demand Type",         "customfield_11682"],
    ["T-Shirt Size (initial)",   "T-Shirt-Sizing",      "customfield_10850"],
    ["Safety score",             "Safety",              "customfield_11438"],
    ["Effort score",             "Effort",              "customfield_11282"],
    ["Discovery report",         "Discovery",           "customfield_11646"],
    ["Epic Link",                "Epic Link",           "customfield_10014"],
], widths_cm=[5.0, 6.0, 5.5])
callout(doc, "If any of these fields shows multiple variants in the "
             "instance, use the field ID listed above to disambiguate.")

para(doc, "", size=4)
para(doc, "5.4.2 · Nine NEW custom fields to create", bold=True)
para(doc, "Where to do it.", bold=True)
para(doc, "Site administration → Issues → Custom fields → Add custom "
          "field. Group them under the label \"DP Custom Fields\" for "
          "easy admin filtering.")
para(doc, "What to create.", bold=True)
table(doc,
      ["Field name", "Type", "Notes"],
      [
        ["DP Value Stream Destination", "Single-select",
         "10 options — see 5.4.3."],
        ["DP Sponsor",                  "User picker",
         "Mandatory at intake. Sponsor ≠ Requester."],
        ["DP Run / Grow / Transform",   "Single-select",
         "3 options: Run, Grow, Transform."],
        ["DP Pain Score",               "Number",
         "Read-only for users; populated by automation A4."],
        ["DP BV Score",                 "Number 1–5",
         "Scored by PM at Analyzing. Distinct from the narrative "
         "Business Value field — this is the numeric input to the "
         "Pain Score formula."],
        ["DP UX Score",                 "Number 1–5",
         "Scored by PM at Analyzing."],
        ["DP OPCO Multiplier",          "Number",
         "Read-only; auto 1.0 / 1.3 / 1.6 set by automation A4."],
        ["DP Concept Prototype",        "URL or Short text",
         "Required for Transform-class items."],
        ["DP Class of Service",         "Single-select",
         "4 options — see 5.4.3."],
      ],
      widths_cm=[5.0, 3.5, 8.5])

para(doc, "", size=4)
para(doc, "5.4.3 · Options (values) for the select fields", bold=True)
para(doc, "After creating each select field, populate its options:",
     italic=True)
table(doc, ["Field", "Options (in order)"], [
    ["DP Value Stream Destination",
     "Field Operations · Engineering & Asset Mgmt · Network Operations & "
     "Control · Corporate & Regulatory · Iberdrola Corp Services · "
     "Meter & Device Operations · Billing Operations · Revenue & "
     "Collections · Customer Care & Engagement · Enabling Team — "
     "Foundation"],
    ["DP Run / Grow / Transform", "Run · Grow · Transform"],
    ["DP Class of Service",
     "Standard · Fixed-Date · Expedite · Intangible"],
    ["OpCo (existing — verify these options exist; add if missing)",
     "UIL · CMP · NY · Iberdrola · Cross-OPCO"],
    ["Customer Impact (existing — verify)",
     "P1 Critical · P2 High · P3 Medium · P4 Low"],
    ["Demand Type (existing — verify)",
     "Production Support · Regulatory · Non-Regulatory"],
    ["T-Shirt-Sizing (existing — verify)",
     "S · M · L  (XL OUT — XL flows through PPM Now)"],
    ["Effort (existing — verify · numeric 1–5)",
     "1 = hours · 2 = days · 3 = weeks · 4 = quarter · 5 = multi-quarter"],
    ["Safety (existing — verify · numeric 1–5)",
     "1–5 scale · ITBA only · Architect-defined · field-crew impact"],
    ["DP UX Score (new · numeric 1–5)",
     "1–5 scale · PM-defined at Analyzing"],
], widths_cm=[5.5, 11.5])

page_break(doc)

# === 5.5 Workflow ============================================================
heading(doc, "5.5 · Workflow — Demand Lifecycle Workflow", 2)
para(doc, "What it is.", bold=True)
para(doc, "A single workflow governs how every issue type moves from "
          "intake to Done. 11 statuses: 8 main + 3 side states.")
para(doc, "Where to do it.", bold=True)
para(doc, "Site administration → Issues → Workflows → Add workflow → "
          "name \"Demand Lifecycle Workflow\". Add the statuses, add the "
          "transitions, then attach via a Workflow Scheme to DGTLPROD.")

para(doc, "Statuses to add:", bold=True)
table(doc, ["Status name", "Category", "Purpose"], [
    ["NEW",          "To Do",       "Demand just submitted by Requester"],
    ["ROUTED",       "To Do",       "Component set, Lead PM notified — "
                                    "awaiting Gate 1"],
    ["TRIAGED",      "In Progress", "Passed Gate 1 — ready for PM Review"],
    ["BOUNCED",      "To Do",       "Returned to Requester for "
                                    "clarification"],
    ["REVIEWING",    "In Progress", "PM defines Type / Size / CoS / RGT "
                                    "(2-week SLA)"],
    ["ANALYZING",    "In Progress", "PM + Architect score; A4 computes "
                                    "Pain Score"],
    ["READY FOR EXECUTION", "In Progress",
                                    "Approved at Gate 2 — awaiting "
                                    "capacity (replaces older \"Backlog\" "
                                    "naming)"],
    ["IMPLEMENTING", "In Progress", "Squad pulled the work · WIP-limited"],
    ["DONE",         "Done",        "Delivered, Sponsor accepts"],
    ["DEFERRED",     "Done",        "Out of scope this cycle"],
    ["REJECTED",     "Done",        "Killed at Gate 2"],
], widths_cm=[3.8, 2.7, 10.5])

para(doc, "Transitions to add:", bold=True)
table(doc, ["Transition", "From → To", "Notes"], [
    ["Submit Demand",                "(create) → NEW",
     "Initial transition."],
    ["A1 Auto-Route",                "NEW → ROUTED",
     "Triggered by automation A1."],
    ["Triage Pass · Gate 1",         "ROUTED → TRIAGED",
     "Allowed only for Lead PM or Architect."],
    ["Triage Bounce",                "ROUTED / TRIAGED → BOUNCED",
     "Comment is mandatory (reason)."],
    ["Resubmit",                     "BOUNCED → NEW",
     "Requester edits and resubmits."],
    ["Begin PM Review",              "TRIAGED → REVIEWING",
     "Allowed only for Lead PM."],
    ["Hand to Architect",            "REVIEWING → ANALYZING",
     "Validator: Type, Size, CoS, RGT must be populated."],
    ["Approve at Gate 2",            "ANALYZING → READY FOR EXECUTION",
     "Allowed only for Director / VP. Validator: Pain Score must be "
     "populated. Outputs: notify Sponsor + Squad Lead."],
    ["Defer at Gate 2",              "ANALYZING → DEFERRED",
     "Allowed only for Director / VP."],
    ["Reject at Gate 2",             "ANALYZING → REJECTED",
     "Allowed only for Director / VP."],
    ["Pull to Squad",                "READY FOR EXECUTION → IMPLEMENTING",
     "Allowed only for Squad PM. Validator: DoR (6 criteria) must pass."],
    ["Mark Done",                    "IMPLEMENTING → DONE",
     "Allowed only for Squad PM. Validator: DoD criteria + Sponsor "
     "acceptance."],
    ["Reopen",                       "(any status) → IMPLEMENTING",
     "Allowed for Squad PM or PM."],
    ["Severity Declared (Channel 2)", "(any status) → IMPLEMENTING",
     "Director-only · skips Funnel · audit comment + post-mortem at next "
     "Gate 2."],
    ["Scope-add (Channel 1)",        "(create child) → READY FOR EXECUTION",
     "Squad PM creates Story under existing Initiative · DoR fast-track · "
     "inherits parent CoS."],
], widths_cm=[5.0, 4.5, 7.5])

callout(doc, "If cloning PPR (Section 4), most of these statuses already "
             "exist; rename or remap PPR statuses to match the names "
             "above. Specifically: PPR's \"Backlog\" → rename to \"Ready "
             "For Execution\".")

page_break(doc)

heading(doc, "5.6 · Field configurations — what is editable when", 2)
para(doc, "What it is.", bold=True)
para(doc,
     "Six field configurations control which fields are visible / "
     "editable / required at each status. The mapping is grounded in TOM "
     "slide 107 (Form Views by Role & Stage).")
para(doc, "Where to do it.", bold=True)
para(doc, "Site administration → Issues → Field configurations → Add "
          "configuration (×6) → then Field configuration schemes → Add "
          "scheme \"DP Field Configuration Scheme\" → associate with "
          "DGTLPROD.")
table(doc,
      ["Configuration", "Status", "Editable", "Hidden", "Editable by"],
      [
        ["FC_FUNNEL", "NEW",
         "Title, Description, Business Value, Acceptance Criteria, DP "
         "Value Stream Destination, OpCo, DP Sponsor, Customer Impact "
         "(optional)",
         "Demand Type, T-Shirt-Sizing, DP Class of Service, DP RGT, all "
         "scoring fields, DP Pain Score",
         "Requester"],
        ["FC_AUTO_ONLY", "ROUTED, TRIAGED",
         "(none — read-only)", "all editing fields",
         "(automation only)"],
        ["FC_REVIEW", "REVIEWING",
         "Demand Type, T-Shirt-Sizing (initial), DP Class of Service, "
         "DP RGT",
         "Scoring fields, DP Pain Score",
         "Lead PM"],
        ["FC_ANALYZE", "ANALYZING",
         "Business Value, DP UX Score, Effort, Safety (ITBA only), "
         "Discovery, DP Concept Prototype (if Transform)",
         "DP Pain Score (display only)",
         "PM + Architect"],
        ["FC_DELIVERY",
         "READY FOR EXECUTION, IMPLEMENTING, DONE",
         "Summary, Description, Assignee, Epic Link, comments",
         "DP Pain Score, DP Class of Service (display only)",
         "Squad PM"],
        ["FC_TERMINAL", "DEFERRED, REJECTED, BOUNCED",
         "Comment (mandatory — captures reason)",
         "everything else",
         "Director / VP"],
      ],
      widths_cm=[2.5, 2.5, 5.0, 4.0, 2.5])

heading(doc, "5.7 · Project roles + permission scheme", 2)
para(doc, "What it is.", bold=True)
para(doc, "Eight project roles, each with their own permissions.")
para(doc, "Where to do it.", bold=True)
para(doc, "Site administration → System → Project roles → Add role "
          "(×8). Then Issues → Permission schemes → Add scheme \"DP "
          "Permission Scheme\" → grant permissions to each role → "
          "associate with DGTLPROD.")
para(doc, "Roles to create:", bold=True)
table(doc, ["Role", "Initial members"], [
    ["dp_requester",     "Group: jira-software-users (everyone in Avangrid)"],
    ["dp_pm_per_vs",     "PMs per VS — list to be confirmed with "
                         "Heather / Kumar / Carlos"],
    ["dp_architect",     "Cross-VS architects — list to be confirmed"],
    ["dp_director_cs",   "kumar.jaladanki@avangrid.com · "
                         "heather.lajoie@avangrid.com"],
    ["dp_director_itba", "carlos.garbin@avangrid.com"],
    ["dp_vp",            "vitor.cruz@avangrid.com"],
    ["dp_squad_pm",      "Squad PMs per VS — list to be confirmed · owns "
                         "DoR + WIP"],
    ["dp_admin",         "NTT consultant (rotates post-handover)"],
], widths_cm=[5.0, 12.0])

para(doc, "Permission grants (matrix):", bold=True)
table(doc,
      ["Permission", "Granted to roles"],
      [
        ["Browse projects",         "All 8 roles"],
        ["Create issues",           "Requester, PM, Squad PM, Admin"],
        ["Edit issues",             "PM, Architect, Squad PM, Director "
                                    "(both), VP, Admin"],
        ["Transition issues",       "PM, Architect, Squad PM, Director "
                                    "(both), VP, Admin"],
        ["Assign issues",           "PM, Squad PM, Admin"],
        ["Assignable user",         "PM, Architect, Squad PM, Admin"],
        ["Add comments",            "All 8 roles"],
        ["Delete issues",           "Admin only"],
        ["Close / Resolve issues",  "Squad PM, Admin"],
        ["Administer projects",     "Admin"],
        ["Manage sprints",          "Squad PM, Admin"],
        ["Declare Severity (Channel 2)",
                                    "Director (both), VP — gated at "
                                    "transition level"],
      ],
      widths_cm=[5.5, 11.5])

callout(doc,
        "Field-level edit restrictions (e.g. only PMs edit scoring "
        "fields, only Director can transition past Gate 2) are enforced "
        "by combining the field configurations from Section 5.6 with the "
        "workflow transition restrictions from Section 5.5.")

page_break(doc)

# === 5.8 NEW · Field-level access matrix =====================================
heading(doc, "5.8 · Field-level access matrix · field progression", 2)
para(doc, "What it is.", bold=True)
para(doc,
     "The matrix below is the source of truth for who edits what, when, "
     "as a demand flows. The TOM principle (slide 94) is \"the card grows "
     "as it advances\" — each role unlocks its own batch of fields at the "
     "right stage, and prior fields become read-only. The admin enforces "
     "this by combining four Jira mechanisms:")
bullet(doc, "Field Configurations (Section 5.6) — fields appear / are "
            "hidden / are required PER STATUS.")
bullet(doc, "Permission Scheme (Section 5.7) — which ROLE can edit any "
            "field on the project at all.")
bullet(doc, "Workflow Validators (on transitions) — block status change "
            "until the right fields are populated.")
bullet(doc, "Workflow Conditions (on transitions) — restrict the "
            "transition itself to a specific role.")
para(doc, "")
para(doc, "Cell legend:", bold=True)
bullet(doc, "Requester* / PM* / Architect* / Director* / Squad PM* / "
            "auto* — REQUIRED at this status by that role / automation.")
bullet(doc, "Requester / PM / Architect / Squad PM — editable but optional.")
bullet(doc, "RO — read-only at this status (display only).")
bullet(doc, "— — hidden from the screen at this status.")
bullet(doc, "auto Ax — value populated by automation Ax (read-only to "
            "humans).")

para(doc, "")
para(doc, "5.8.1 · The matrix", bold=True)
table(doc,
      ["Field", "NEW (Funnel)", "ROUTED + TRIAGED",
       "REVIEWING", "ANALYZING",
       "READY · IMPL · DONE", "DEFER · REJECT · BOUNCE"],
      [
        ["Title (Summary)",
         "Requester*", "RO", "RO", "RO",
         "Squad PM (edit)", "RO"],
        ["Description",
         "Requester*", "RO", "RO", "RO",
         "Squad PM (edit)", "RO"],
        ["Business Value (narrative)",
         "Requester*", "RO", "RO", "RO", "RO", "RO"],
        ["Acceptance Criteria",
         "Requester*", "RO", "RO", "RO",
         "Squad PM (edit)", "RO"],
        ["DP Value Stream Destination",
         "Requester*", "RO", "RO", "RO", "RO", "RO"],
        ["OpCo (multi-select)",
         "Requester*", "RO", "RO", "RO", "RO", "RO"],
        ["DP Sponsor (user picker)",
         "Requester*", "RO", "RO", "RO", "RO", "RO"],
        ["Customer Impact (P1–P4)",
         "Requester (opt)", "RO", "PM (edit)", "RO", "RO", "RO"],
        ["Discovery (attachment)",
         "Requester (opt)", "RO", "RO", "PM (edit)", "RO", "RO"],
        ["Demand Type (PS/Reg/Non-Reg)",
         "—", "—", "PM*", "RO", "RO", "RO"],
        ["T-Shirt-Sizing (S/M/L)",
         "—", "—", "PM*", "RO", "RO", "RO"],
        ["DP Class of Service",
         "—", "—", "PM* (or auto A2/A3)", "RO", "RO", "RO"],
        ["DP Run / Grow / Transform",
         "—", "—", "PM*", "RO", "RO", "RO"],
        ["DP BV Score (1–5)",
         "—", "—", "—", "PM*", "RO", "RO"],
        ["DP UX Score (1–5)",
         "—", "—", "—", "PM*", "RO", "RO"],
        ["Safety (1–5, ITBA only)",
         "—", "—", "—", "Architect*", "RO", "RO"],
        ["Effort (1–5)",
         "—", "—", "—", "Architect*", "RO", "RO"],
        ["DP Concept Prototype (URL)",
         "—", "—", "—",
         "Architect (req if RGT=Transform)", "RO", "RO"],
        ["DP OPCO Multiplier",
         "—", "—", "—", "auto A4", "RO", "RO"],
        ["DP Pain Score",
         "—", "—", "—", "auto A4", "RO", "RO"],
        ["Component (= VS routing)",
         "—", "auto A1", "RO", "RO", "RO", "RO"],
        ["Assignee",
         "—", "auto A1 (Lead PM)",
         "PM (edit)", "PM (edit)",
         "Squad PM (edit)", "Squad PM (edit)"],
        ["Comment",
         "Requester (edit)", "all (edit)",
         "all (edit)", "all (edit)",
         "all (edit)",
         "Director* (mandatory · captures reason)"],
      ],
      widths_cm=[4.0, 2.2, 2.2, 2.5, 2.5, 2.5, 2.5])

para(doc, "")
para(doc, "5.8.2 · Workflow validators (block transition unless ALL "
          "fields populated)", bold=True)
table(doc, ["Transition", "Required-to-pass fields"], [
    ["Submit Demand → NEW",
     "Title · Description · Business Value (narrative) · Acceptance "
     "Criteria · DP Value Stream Destination · OpCo · DP Sponsor"],
    ["Triage Pass · Gate 1",
     "(no extra field validators — role check only)"],
    ["Begin PM Review",
     "(no extra field validators — role check only)"],
    ["Hand to Architect",
     "Demand Type · T-Shirt-Sizing · DP Class of Service · DP RGT"],
    ["Approve at Gate 2",
     "DP BV Score · DP UX Score · Effort · Safety (only if VS ∈ ITBA) "
     "· DP Pain Score (auto by A4) · DP Concept Prototype (only if RGT "
     "= Transform)"],
    ["Pull to Squad",
     "DoR checkbox / comment confirming all 6 DoR criteria pass "
     "(see Section 6)"],
    ["Mark Done",
     "Sponsor acceptance comment + DoD criteria"],
    ["Triage Bounce / Defer / Reject",
     "Comment field mandatory (reason)"],
], widths_cm=[5.5, 11.5])

para(doc, "")
para(doc, "5.8.3 · Workflow conditions (which role may execute each "
          "transition)", bold=True)
table(doc, ["Transition", "Allowed role(s)"], [
    ["Submit Demand",          "dp_requester"],
    ["A1 Auto-Route",          "(automation only)"],
    ["Triage Pass · Gate 1",   "dp_pm_per_vs OR dp_architect"],
    ["Triage Bounce",          "dp_pm_per_vs OR dp_architect"],
    ["Resubmit",               "dp_requester (the original Reporter)"],
    ["Begin PM Review",        "dp_pm_per_vs"],
    ["Hand to Architect",      "dp_pm_per_vs"],
    ["Approve at Gate 2",      "dp_director_cs OR dp_director_itba OR "
                               "dp_vp"],
    ["Defer at Gate 2",        "dp_director_cs OR dp_director_itba OR "
                               "dp_vp"],
    ["Reject at Gate 2",       "dp_director_cs OR dp_director_itba OR "
                               "dp_vp"],
    ["Pull to Squad",          "dp_squad_pm"],
    ["Mark Done",              "dp_squad_pm"],
    ["Reopen",                 "dp_squad_pm OR dp_pm_per_vs"],
    ["Severity Declared (Channel 2)",
                               "dp_director_cs OR dp_director_itba "
                               "(for their area) OR dp_vp"],
    ["Scope-add (Channel 1, child create)",
                               "dp_squad_pm"],
], widths_cm=[5.5, 11.5])
callout(doc,
        "Net effect: a Requester sees ONLY the 8 funnel fields when "
        "submitting; a PM at Reviewing sees the 4 classification fields "
        "appear; an Architect at Analyzing sees the 4 scoring fields "
        "appear; the Director at Gate 2 sees a fully-populated card with "
        "Pain Score auto-computed and read-only. Everyone else's edits "
        "are blocked by the field-config + permission combination.")

page_break(doc)

heading(doc, "5.9 · Boards and swim lanes", 2)
para(doc, "What it is.", bold=True)
para(doc,
     "Eleven kanban boards — one Portfolio Overview for the Director / "
     "VP cross-VS view, plus 10 Squad boards (one per Value Stream). "
     "All use swim lanes by Class of Service, with WIP limits.")
para(doc, "Where to do it.", bold=True)
para(doc, "Top bar → Boards → Create board → Kanban → from saved filter "
          "(×11). Save the filters first.")
para(doc, "Swim-lane configuration (same template for all 11 boards):",
     bold=True)
table(doc, ["Lane", "Filter (JQL)", "WIP limit"], [
    ["Expedite",   '"DP Class of Service" = Expedite',     "1"],
    ["Fixed-Date", '"DP Class of Service" = "Fixed-Date"', "2"],
    ["Standard",   '"DP Class of Service" = Standard',     "4"],
    ["Intangible", '"DP Class of Service" = Intangible',   "1"],
], widths_cm=[3.0, 9.0, 5.0])

para(doc, "Filters to create (one per board):", bold=True)
table(doc, ["Filter name", "JQL"], [
    ["DP_Filter_Portfolio",
     'project = DGTLPROD AND status in (TRIAGED, REVIEWING, ANALYZING, '
     '"Ready For Execution", IMPLEMENTING, DONE) ORDER BY '
     '"DP Pain Score" DESC'],
    ["DP_Filter_Squad_Field_Operations",
     'project = DGTLPROD AND component = "Field Operations" AND status '
     'not in (NEW, ROUTED) ORDER BY "DP Pain Score" DESC'],
    ["DP_Filter_Squad_Engineering_AssetMgmt",
     'project = DGTLPROD AND component = "Engineering & Asset Mgmt" '
     'AND status not in (NEW, ROUTED) ORDER BY "DP Pain Score" DESC'],
    ["DP_Filter_Squad_Network_Ops",
     'project = DGTLPROD AND component = "Network Operations & Control" '
     'AND status not in (NEW, ROUTED) ORDER BY "DP Pain Score" DESC'],
    ["DP_Filter_Squad_Corporate_Regulatory",
     'project = DGTLPROD AND component = "Corporate & Regulatory" AND '
     'status not in (NEW, ROUTED) ORDER BY "DP Pain Score" DESC'],
    ["DP_Filter_Squad_Iberdrola_Corp",
     'project = DGTLPROD AND component = "Iberdrola Corp Services" AND '
     'status not in (NEW, ROUTED) ORDER BY "DP Pain Score" DESC'],
    ["DP_Filter_Squad_Meter_Device",
     'project = DGTLPROD AND component = "Meter & Device Operations" '
     'AND status not in (NEW, ROUTED) ORDER BY "DP Pain Score" DESC'],
    ["DP_Filter_Squad_Billing_Operations",
     'project = DGTLPROD AND component = "Billing Operations" AND '
     'status not in (NEW, ROUTED) ORDER BY "DP Pain Score" DESC'],
    ["DP_Filter_Squad_Revenue_Collections",
     'project = DGTLPROD AND component = "Revenue & Collections" AND '
     'status not in (NEW, ROUTED) ORDER BY "DP Pain Score" DESC'],
    ["DP_Filter_Squad_Customer_Care",
     'project = DGTLPROD AND component = "Customer Care & Engagement" '
     'AND status not in (NEW, ROUTED) ORDER BY "DP Pain Score" DESC'],
    ["DP_Filter_Squad_Enabling_Foundation",
     'project = DGTLPROD AND component = "Enabling Team — Foundation" '
     'AND status not in (NEW, ROUTED) ORDER BY "DP Pain Score" DESC'],
], widths_cm=[5.5, 11.5])

para(doc, "Board columns:", bold=True)
bullet(doc, "Portfolio Overview: TRIAGED · REVIEWING · ANALYZING · "
            "READY FOR EXECUTION · IMPLEMENTING · DONE.")
bullet(doc, "Each Squad board: READY FOR EXECUTION · IMPLEMENTING (In "
            "Progress) · REVIEW · DONE.")

page_break(doc)

heading(doc, "5.10 · Automation rules — five rules", 2)
para(doc,
     "Five automation rules cover the recurring orchestration. They are "
     "the operational backbone of the demand lifecycle and stay with "
     "Avangrid after handover, so this section is intentionally "
     "exhaustive — the admin team should be able to recreate every "
     "rule from this document alone, with no NTT DATA assistance.")

# ----- 5.10.0 prerequisites -----
para(doc, "", size=4)
para(doc, "5.10.0 · Pre-requisites and admin permissions",
     bold=True, color=NAVY)
para(doc, "Verify all of the following BEFORE creating the five rules:")
table(doc, ["#", "Item", "Why / how to check"], [
    ["1", "Atlassian Automation enabled at project level on DGTLPROD",
     "Project Settings → Automation panel must be visible. Available "
     "from Premium tier upward. If not visible, request enablement "
     "from Site Admin."],
    ["2", "Outbound email functional",
     "A1, A2 and A3 use Send Email actions. Confirm SMTP is configured "
     "at site level (Site Admin → System → Mail → Outgoing mail). "
     "Test by sending a manual notification."],
    ["3", "Monthly automation quota",
     "Each rule run consumes from the org's monthly automation quota "
     "(Standard ≈ 500/mo · Premium ≈ 1,000/mo per project · "
     "Enterprise = unlimited). Estimated load: 5 rules × ~3 executions "
     "× ~100 demands/month ≈ 1,500 runs/month. Confirm the Avangrid "
     "tier supports this; otherwise request an upgrade."],
    ["4", "Smart-value lookup tables enabled",
     "A1 looks up Lead PM by VS · A2 looks up Director by VS. If "
     "lookup tables are not available on the tier, NTT DATA will "
     "replicate via hard-coded if/else conditional branches inside "
     "each rule (functionally equivalent, more maintenance)."],
    ["5", "Cross-issue create allowed",
     "A5 creates child issues (Story / Epic / Initiative). Default Cloud "
     "behaviour: yes. Confirm by checking Site Admin → Global automation "
     "→ Allowed actions does NOT exclude \"Create issue\"."],
    ["6", "Rule audit log retention",
     "Default 90 days. If Avangrid compliance requires longer, increase "
     "in Site Admin → Audit log."],
    ["7", "Permission to create the rules",
     "Once the NTT consultant is in the Project Administrator role of "
     "DGTLPROD, the 5 rules can be created without further admin "
     "involvement. NO global automation permission is needed."],
    ["8", "Rule actor (impersonation user)",
     "By default, automations run as the rule owner. Recommended to set "
     "the rule actor to a dedicated service account (e.g. "
     "automation-bot@avangrid.com) so audit trails show \"Automation\" "
     "rather than a person."],
], widths_cm=[0.7, 5.0, 11.3])

para(doc, "", size=4)
para(doc, "5.10.1 · Lookup tables (Lead PM and Director by Value Stream)",
     bold=True, color=NAVY)
para(doc, "A1 and A2 need to map a Value Stream name to an email "
          "recipient. Two options:")
bullet(doc, "Option A · Smart-value lookup table — preferred. "
            "Project Settings → Automation → Library → Create lookup "
            "table. Two tables required:")
bullet(doc, "Option B · Hard-coded if/else inside the rule — works on "
            "any tier but harder to maintain.", level=1)

para(doc, "Lookup table 1 · LeadPM_by_VS (used by A1):", bold=True)
table(doc, ["Key (VS Destination)", "Value (Lead PM email)"], [
    ["Field Operations",             "lead-pm-fieldops@avangrid.com (TBD)"],
    ["Engineering & Asset Mgmt",     "lead-pm-eam@avangrid.com (TBD)"],
    ["Network Operations & Control", "lead-pm-noc@avangrid.com (TBD)"],
    ["Corporate & Regulatory",       "lead-pm-corpreg@avangrid.com (TBD)"],
    ["Iberdrola Corp Services",      "lead-pm-ics@avangrid.com (TBD)"],
    ["Meter & Device Operations",    "lead-pm-mdo@avangrid.com (TBD)"],
    ["Billing Operations",           "lead-pm-billing@avangrid.com (TBD)"],
    ["Revenue & Collections",        "lead-pm-revcol@avangrid.com (TBD)"],
    ["Customer Care & Engagement",   "lead-pm-cce@avangrid.com (TBD)"],
    ["Enabling Team — Foundation",   "lead-pm-enabling@avangrid.com (TBD)"],
], widths_cm=[6.0, 11.0])

para(doc, "Lookup table 2 · Director_by_VS (used by A2):", bold=True)
table(doc, ["Key (VS Destination)", "Value (Director email)"], [
    ["Field Operations",             "carlos.garbin@avangrid.com"],
    ["Engineering & Asset Mgmt",     "carlos.garbin@avangrid.com"],
    ["Network Operations & Control", "carlos.garbin@avangrid.com"],
    ["Corporate & Regulatory",       "carlos.garbin@avangrid.com"],
    ["Iberdrola Corp Services",      "carlos.garbin@avangrid.com"],
    ["Meter & Device Operations",    "kumar.jaladanki@avangrid.com"],
    ["Billing Operations",           "kumar.jaladanki@avangrid.com"],
    ["Revenue & Collections",        "kumar.jaladanki@avangrid.com"],
    ["Customer Care & Engagement",   "heather.lajoie@avangrid.com"],
    ["Enabling Team — Foundation",   "kumar.jaladanki@avangrid.com"],
], widths_cm=[6.0, 11.0])

para(doc, "", size=4)
para(doc, "5.10.2 · How to create a rule (UI walkthrough common to all "
          "five)", bold=True, color=NAVY)
numbered(doc, "Open Project DGTLPROD → Project Settings → Automation.")
numbered(doc, "Click Create rule.")
numbered(doc, "Pick the trigger (\"When this happens…\") from the table "
              "below.")
numbered(doc, "Click Add component → Condition. Pick the IF block from "
              "the table.")
numbered(doc, "Click Add component → Action. Pick the actions from the "
              "table.")
numbered(doc, "Name the rule exactly as shown (A1 · A2 · A3 · A4 · A5).")
numbered(doc, "(Recommended) Set the rule actor to the automation-bot "
              "account, not your personal user.")
numbered(doc, "Save and Enable the rule.")
numbered(doc, "Test by creating a sample issue that should trigger it · "
              "verify in the rule's audit log (right pane).")

para(doc, "", size=4)
para(doc, "Rules detailed below — A1 to A5.", italic=True, color=GREY)

# ----- A1 -----
para(doc, "A1 · Auto-Routing", bold=True, color=NAVY)
table(doc,
      ["Step", "What to set"],
      [
        ["Trigger",   "Issue created."],
        ["Condition", "Issue type = Demand Intake AND DP Value Stream "
                      "Destination is not empty."],
        ["Action 1",  "Edit issue · set Component = same name as DP "
                      "Value Stream Destination."],
        ["Action 2",  "Transition issue to ROUTED."],
        ["Action 3",  "Send email to the Lead PM of that Value Stream "
                      "(use a smart-value lookup table or hard-code a "
                      "mapping; emails to be confirmed before go-live)."],
      ],
      widths_cm=[3.0, 14.0])

# ----- A2 -----
para(doc, "", size=4)
para(doc, "A2 · Auto-Expedite (Customer Impact P1 / P2)",
     bold=True, color=NAVY)
table(doc,
      ["Step", "What to set"],
      [
        ["Trigger",   "Field value changed → Customer Impact."],
        ["Condition", "Customer Impact is one of {P1 Critical, P2 High}."],
        ["Action 1",  "Edit issue · set DP Class of Service = Expedite."],
        ["Action 2",  "Set SLA = 4 hours (P1) OR 8 hours (P2)."],
        ["Action 3",  "Send email to DP Sponsor and to the VS Director."],
        ["Action 4",  "Add comment audit trail (\"Auto-Expedite triggered "
                      "by A2…\")."],
      ],
      widths_cm=[3.0, 14.0])

# ----- A3 -----
para(doc, "", size=4)
para(doc, "A3 · Auto-Fixed-Date (Regulatory)", bold=True, color=NAVY)
table(doc,
      ["Step", "What to set"],
      [
        ["Trigger",   "Field value changed → Demand Type."],
        ["Condition", "Demand Type = Regulatory."],
        ["Action 1",  "Edit issue · set DP Class of Service = Fixed-Date."],
        ["Action 2",  "Send email to Assignee (PM) requesting Due Date "
                      "within 2 business days."],
        ["Action 3",  "Add comment audit trail."],
      ],
      widths_cm=[3.0, 14.0])

# ----- A4 -----
para(doc, "", size=4)
para(doc, "A4 · Auto-Pain-Score", bold=True, color=NAVY)
table(doc,
      ["Step", "What to set"],
      [
        ["Trigger",   "Field value changed on any of: Business Value, "
                      "DP UX Score, Safety, Effort, OpCo."],
        ["Condition", "Business Value, DP UX Score and Effort all "
                      "populated. If DP Value Stream Destination ∈ ITBA "
                      "group, Safety must also be populated; otherwise "
                      "(CS), Safety not required."],
        ["Action 1",  "Compute DP OPCO Multiplier from the count of "
                      "OpCos selected: 1 OpCo → 1.0; 2–4 → 1.3; 5+ → 1.6."],
        ["Action 2",  "Compute DP Pain Score using the appropriate "
                      "formula. ITBA: (BV·OPCOmult·0.45 + UX·0.30 + "
                      "Safety·0.25)/Effort. CS: (BV·OPCOmult·0.40 + "
                      "UX·0.60)/Effort."],
        ["Action 3",  "Add comment audit trail with the new score and "
                      "multiplier."],
      ],
      widths_cm=[3.0, 14.0])
callout(doc, "DP Pain Score and DP OPCO Multiplier must be writable by "
             "automation but read-only to humans — enforce via field "
             "configuration + permission scheme.")

# ----- A5 -----
para(doc, "", size=4)
para(doc, "A5 · Auto-Convert at Triage (size → outcome-named hierarchy)",
     bold=True, color=NAVY)
table(doc,
      ["Step", "What to set"],
      [
        ["Trigger",   "Status changed to TRIAGED."],
        ["Condition", "Issue type = Demand Intake AND T-Shirt-Sizing not "
                      "empty AND Demand Type not empty."],
        ["Branch S",  "If Size = S → create one Story (no parent), copy "
                      "Components and DP Class of Service, link to "
                      "original Demand."],
        ["Branch M",  "If Size = M → create one Epic (no parent "
                      "Initiative), copy Components and CoS."],
        ["Branch L",  "If Size = L → create one Initiative (default "
                      "top-level for L). If estimated > 3 months at "
                      "refinement, Lead PM later breaks the Initiative "
                      "into Epics manually."],
        ["Action",    "Add comment on original Demand recording the "
                      "conversion + new Initiative / Epic / Story key."],
      ],
      widths_cm=[3.0, 14.0])

callout(doc, "A6 (Aging Alerts) was dropped — Jira's native SLA "
             "tracking covers this.")

page_break(doc)

# === 6. INVEST + DoR =========================================================
heading(doc, "6 · Story-level quality bars (INVEST + DoR)", 1)
para(doc,
     "Squad PM is accountable for two checks before any Story enters "
     "Implementing.")
heading(doc, "6.1 · INVEST (every Story)", 2)
table(doc, ["Letter", "Meaning"], [
    ["I", "Independent — can be progressed without depending on others"],
    ["N", "Negotiable — scope is open to refinement"],
    ["V", "Valuable — outcome is named in customer language"],
    ["E", "Estimable — Squad can put a t-shirt size to it"],
    ["S", "Small — ≤ 5 SP"],
    ["T", "Testable — has clear acceptance criteria"],
], widths_cm=[1.5, 15.0])

heading(doc, "6.2 · Definition of Ready (single gate to Execution)", 2)
para(doc, "Six criteria, all must be true before pulling into "
          "Implementing:")
numbered(doc, "Acceptance criteria written and testable.")
numbered(doc, "Effort estimated and on the same scale as the rest of the "
              "backlog.")
numbered(doc, "Dependencies identified and scheduled.")
numbered(doc, "Sponsor confirmed (still active).")
numbered(doc, "Class of Service consistent with parent.")
numbered(doc, "No open clarification with Requester.")

# === 7. Smoke test ===========================================================
heading(doc, "7 · Smoke test — five sample demands", 1)
para(doc,
     "Once provisioning is complete, NTT DATA will create five sample "
     "demands that exercise every path. The admin team can use this as "
     "the acceptance test for the configuration.")
table(doc,
      ["#", "Demand", "Value Stream", "OpCo", "Size", "Type", "CIT",
       "Expected outcome"],
      [
        ["1", "Reduce billing-correction cycle 14→5 days",
         "Billing Operations", "UIL, CMP, NY", "M", "Non-Regulatory",
         "P3",
         "A5 → Epic · CoS = Standard · Ready For Execution after Gate 2"],
        ["2", "PUC tariff banner update for Q3",
         "Corporate & Regulatory", "All 5", "L", "Regulatory", "—",
         "A5 → Initiative · A3 → CoS = Fixed-Date · deadline tracked"],
        ["3", "CRM page load <2s on Genesys",
         "Customer Care & Engagement", "UIL", "S", "Non-Regulatory",
         "P2",
         "A2 → CoS = Expedite · WIP-1 · 8-h SLA · Story direct to squad"],
        ["4", "GIS layer migration · scope-add to existing CRM Initiative",
         "Field Operations (linked)", "UIL", "S", "Non-Regulatory", "—",
         "Channel 1 scope-add · Story under existing Initiative · Squad "
         "PM accepts within 48h"],
        ["5", "SAP Basis kernel patch Q2",
         "Iberdrola Corp Services", "Iberdrola", "M", "Production "
         "Support", "P4",
         "A5 → Epic · CoS = Intangible (tech debt)"],
      ],
      widths_cm=[0.6, 4.5, 3.5, 1.8, 0.8, 1.8, 0.7, 4.0])
para(doc, "Pass criteria.", bold=True)
bullet(doc, "All 5 issues created · A1 fires for all of them.")
bullet(doc, "A2 fires for #3; A3 fires for #2.")
bullet(doc, "After scoring, A4 computes Pain Score correctly per "
            "area-specific formula.")
bullet(doc, "After transition to TRIAGED, A5 converts each Demand to "
            "the correct issue type per the outcome-named hierarchy "
            "(Story for S, Epic for M, Initiative for L).")
bullet(doc, "Cards appear in the correct Squad board, in the correct "
            "Class-of-Service swim lane, respecting the WIP limits.")

# === 8. Acceptance checklist =================================================
heading(doc, "8 · Acceptance / handoff checklist", 1)
para(doc, "Mark each item as it is completed:")
bullet(doc, "Project DGTLPROD exists at "
            "https://avangridjira.atlassian.net/jira/software/projects/DGTLPROD")
bullet(doc, "The NTT consultant is in the Administrators role of "
            "DGTLPROD.")
bullet(doc, "All 10 components are present.")
bullet(doc, "Issue Type Scheme \"DP Issue Type Scheme\" is associated "
            "with DGTLPROD; all SIX issue types (Demand Intake, "
            "Initiative, Epic, Story, Bug, Sub-task) are available; "
            "Demand Intake is the default.")
bullet(doc, "Workflow \"Demand Lifecycle Workflow\" (or PPR-cloned "
            "equivalent) is associated with DGTLPROD; all 11 statuses "
            "and 15 transitions exist (including the two unplanned-"
            "demand channels).")
bullet(doc, "All 9 new custom fields exist with correct options "
            "(including DP BV Score 1-5 and the 1–5 Effort/Safety "
            "scales).")
bullet(doc, "All 10 reused custom fields are added to the DGTLPROD "
            "screens.")
bullet(doc, "Field Configuration Scheme \"DP Field Configuration "
            "Scheme\" is associated with DGTLPROD; the 6 configurations "
            "match Section 5.6 and the consolidated access matrix in "
            "Section 5.8.")
bullet(doc, "Workflow validators in place for every transition listed "
            "in Section 5.8.2 (block bad transitions when required "
            "fields missing).")
bullet(doc, "Workflow conditions in place for every transition listed "
            "in Section 5.8.3 (only the right role can run each "
            "transition).")
bullet(doc, "All 8 project roles exist and Permission Scheme \"DP "
            "Permission Scheme\" is associated with DGTLPROD.")
bullet(doc, "All 11 boards exist with correct filters, columns, swim "
            "lanes and WIP limits.")
bullet(doc, "Automation prerequisites from Section 5.10.0 verified "
            "(SMTP, quota, lookup tables, cross-issue create allowed).")
bullet(doc, "Lookup tables LeadPM_by_VS and Director_by_VS populated "
            "with confirmed (non-TBD) email addresses.")
bullet(doc, "All 5 automation rules (A1–A5) are enabled and have run "
            "at least once successfully on the smoke test.")
bullet(doc, "Confirmation email sent to requester so NTT DATA can begin "
            "user onboarding.")

# === 9. Contact ==============================================================
heading(doc, "9 · Contact for clarifications", 1)
bullet(doc, "Primary  ·  NTT consultant")
bullet(doc, "Avangrid sponsor  ·  Vitor Cruz, VP Digital Products")
bullet(doc, "Stakeholders  ·  Carlos Garbin (ITBA) · Kumar Jaladanki "
            "(CS) · Heather LaJoie (CS · CCC)")
para(doc, "", size=10)
callout(doc,
        "Happy to jump on a quick call if any of the above needs "
        "clarification before provisioning begins. Thank you for the "
        "support.")

# === Save ====================================================================
doc.save(OUT)
print(f"✅ Wrote {OUT}")
print(f"   {OUT.stat().st_size / 1024:.1f} KB")
