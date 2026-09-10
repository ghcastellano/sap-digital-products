"""Build the formal request document for the Avangrid Jira admin team.

Companion to Avangrid_Jira_DGTLPROD_Admin_Guide.docx (the technical "HOW").
This file is the formal "WHAT" — an itemized request the admin team can
acknowledge, accept and execute against. Self-contained, ~6 pages.
"""
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = Path(__file__).parent / "Avangrid_Jira_DGTLPROD_Request.docx"

NAVY  = RGBColor(0x1F, 0x38, 0x64)
ACCENT= RGBColor(0x2E, 0x75, 0xB6)
GREY  = RGBColor(0x59, 0x59, 0x59)


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


# ============================================================================
doc = Document()
style = doc.styles["Normal"]
style.font.name = "Calibri"
style.font.size = Pt(11)


# === Cover ===================================================================
para(doc, "FORMAL REQUEST", bold=True, size=12, color=ACCENT,
     align=WD_ALIGN_PARAGRAPH.CENTER)
para(doc, "", size=4)
para(doc, "Provisioning of Jira Cloud Project DGTLPROD",
     bold=True, size=24, color=NAVY, align=WD_ALIGN_PARAGRAPH.CENTER)
para(doc, "Avangrid Digital Products — Target Operating Model",
     bold=False, size=14, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)
para(doc, "", size=20)
hrule(doc)

para(doc, "TO  ·  Avangrid Jira Cloud Admin team",
     size=11, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)
para(doc, "FROM  ·  NTT DATA — Avangrid Digital Products engagement",
     size=11, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)
para(doc, "REQUESTER  ·  NTT consultant",
     size=11, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)
para(doc, "SPONSOR  ·  Vitor Cruz, VP Digital Products",
     size=11, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)
para(doc, "STAKEHOLDERS  ·  Carlos Garbin (ITBA)  ·  "
          "Kumar Jaladanki (CS)  ·  Heather LaJoie (CS · CCC)",
     size=11, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)
para(doc, "INSTANCE  ·  avangridjira.atlassian.net",
     size=11, color=GREY, align=WD_ALIGN_PARAGRAPH.CENTER)
hrule(doc)
para(doc, "", size=14)

callout(doc,
        "This document is the formal request asking the Avangrid Jira "
        "Cloud admin team to provision a new project (DGTLPROD) and the "
        "supporting configuration needed to operationalize the Avangrid "
        "Digital Products Target Operating Model designed by NTT DATA. "
        "Technical implementation details are in the companion document "
        "\"Avangrid_Jira_DGTLPROD_Admin_Guide.docx\".")

page_break(doc)


# === 1. Background ==========================================================
heading(doc, "1 · Background", 1)
para(doc,
     "Avangrid is consolidating Customer Services Applications (CS) and "
     "IT Business Applications (ITBA) into a single Digital Products "
     "organization. NTT DATA has designed and the Director group has "
     "approved an end-to-end Target Operating Model (TOM) covering "
     "demand intake, triage, portfolio governance and squad delivery "
     "across 10 Value Streams and 5 OpCos.")
para(doc,
     "Operationalizing this TOM requires a single dedicated Jira Cloud "
     "project on the existing Avangrid instance. There is currently no "
     "equivalent project. The admin team's involvement is the unblocker "
     "for the next phase of the engagement.")


# === 2. The request — three asks ============================================
heading(doc, "2 · The request", 1)
para(doc, "We are formally asking the Avangrid Jira admin team for "
          "three actions:", bold=True)

numbered(doc,
    "Create a new Jira Cloud project · key DGTLPROD · name \"Digital "
    "Products\" · type Company-managed (classic). Recommended approach: "
    "clone PPR (Pilot Project · Renewables, lead Marisa Rojas Gong) — "
    "PPR's existing workflow and use of the Demand Intake issue type "
    "are the closest match to the TOM design and shorten provisioning "
    "by an estimated 30 %.")

numbered(doc,
    "Grant Project Administrator role on DGTLPROD to the NTT "
    "consultant. Once granted, NTT DATA can configure the rest of the "
    "system without further admin involvement, OR the admin team can "
    "perform the full configuration described in Section 3 — both "
    "paths are acceptable.")

numbered(doc,
    "(Optional, time-saving) Temporarily grant the global \"Create "
    "next-gen projects\" permission to the requester for the duration "
    "of setup. Can be revoked once DGTLPROD is verified.")

para(doc, "", size=4)
para(doc, "Out of scope (explicit, to remove ambiguity):", bold=True)
bullet(doc, "No integration between Jira and ServiceNow. All incidents "
            "remain 100 % in ServiceNow. This request creates no "
            "webhook, no scripted bridge, no MID-server connection.")
bullet(doc, "No Jira Service Management licenses required. The existing "
            "Jira Software SKU is sufficient.")
bullet(doc, "No new user accounts. Every participant already has an "
            "Avangrid Atlassian identity.")
bullet(doc, "No XL initiatives. XL flows through PPM Now (separate "
            "Avangrid system). The Demand Lifecycle in DGTLPROD handles "
            "S, M and L only.")


# === 3. Configuration items requested =======================================
heading(doc, "3 · Configuration items to be created", 1)
para(doc,
     "The full DGTLPROD configuration consists of the following nine "
     "items. Each is detailed in the companion Admin Guide; the "
     "summary below is enough to size the work and confirm the scope.")

# --- 3.1 Project shell ---
heading(doc, "3.1 · Project shell", 2)
table(doc, ["Property", "Value"], [
    ["Project name",   "Digital Products"],
    ["Project key",    "DGTLPROD"],
    ["Project type",   "Company-managed (classic)"],
    ["Template",       "Kanban (Software) — or inherited from PPR clone"],
    ["Project lead",   "NTT consultant (rotates to Avangrid lead "
                        "post-handover)"],
    ["URL after creation",
                       "https://avangridjira.atlassian.net/jira/software/"
                       "projects/DGTLPROD"],
], widths_cm=[5.0, 12.0])

# --- 3.2 Components ---
heading(doc, "3.2 · Components — 10, one per Value Stream", 2)
para(doc, "Routes demands to the right Value Stream within the single "
          "project. One demand carries one component.")
table(doc, ["#", "Component", "Area", "Director / OVS owner"], [
    ["1", "Field Operations",             "ITBA", "Carlos Garbin"],
    ["2", "Engineering & Asset Mgmt",     "ITBA", "Carlos Garbin"],
    ["3", "Network Operations & Control", "ITBA", "Carlos Garbin"],
    ["4", "Corporate & Regulatory",       "ITBA", "Carlos Garbin"],
    ["5", "Iberdrola Corp Services",      "ITBA", "Carlos Garbin"],
    ["6", "Meter & Device Operations",    "CS",   "Kumar Jaladanki"],
    ["7", "Billing Operations",           "CS",   "Kumar Jaladanki"],
    ["8", "Revenue & Collections",        "CS",   "Kumar Jaladanki"],
    ["9", "Customer Care & Engagement",   "CS",   "Heather LaJoie"],
    ["10","Enabling Team — Foundation",   "CS",   "Kumar Jaladanki"],
], widths_cm=[1.0, 7.0, 2.5, 5.0])

# --- 3.3 Issue types ---
heading(doc, "3.3 · Issue types — 6 (all already exist globally · "
              "REUSE)", 2)
para(doc, "No new issue types created — only an Issue Type Scheme "
          "associating these to DGTLPROD. Default = Demand Intake.")
table(doc, ["Issue type", "Use", "Maps to size"], [
    ["Demand Intake", "Intake placeholder · auto-converts at Triage", "—"],
    ["Initiative",    "Default top-level for portfolio items · "
                      "≤ 3 months · sponsor-approved", "L"],
    ["Epic",          "Mid-sized work · or exception when Initiative "
                      "exceeds 3 months", "M"],
    ["Story",         "Squad-level · INVEST · ≤ 5 SP", "S"],
    ["Bug",           "Internal QA defects (NOT incidents)", "—"],
    ["Sub-task",      "Technical breakdown within a Story", "—"],
], widths_cm=[3.0, 11.0, 3.0])

# --- 3.4 Custom fields ---
heading(doc, "3.4 · Custom fields — 9 new + 10 reuse", 2)

para(doc, "Reuse (existing Avangrid fields, only need to be added to "
          "DGTLPROD screens):", bold=True)
table(doc, ["Need", "Field name", "Field ID"], [
    ["OpCo",                      "OpCo",               "customfield_10800"],
    ["Business Value (narrative)","Business Value",     "customfield_11440"],
    ["Acceptance Criteria",       "Acceptance Criteria","customfield_10795"],
    ["Customer Impact Tier",      "Customer Impact",    "customfield_11061"],
    ["Demand Type",               "Demand Type",        "customfield_11682"],
    ["T-Shirt Size (initial)",    "T-Shirt-Sizing",     "customfield_10850"],
    ["Safety score (1–5, ITBA)",  "Safety",             "customfield_11438"],
    ["Effort score (1–5)",        "Effort",             "customfield_11282"],
    ["Discovery report",          "Discovery",          "customfield_11646"],
    ["Epic Link",                 "Epic Link",          "customfield_10014"],
], widths_cm=[5.0, 6.0, 5.5])

para(doc, "Create (9 net-new fields, prefix \"DP\"):", bold=True)
table(doc, ["Field name", "Type", "Notes"], [
    ["DP Value Stream Destination", "Single-select",
     "10 options (the 10 components above)."],
    ["DP Sponsor",                  "User picker",
     "Mandatory at intake. Sponsor ≠ Requester."],
    ["DP Run / Grow / Transform",   "Single-select",
     "Run · Grow · Transform."],
    ["DP Pain Score",               "Number (read-only)",
     "Auto-populated by automation A4."],
    ["DP BV Score",                 "Number 1–5",
     "Scored by PM at Analyzing. Numeric input to Pain Score formula. "
     "Distinct from the narrative Business Value."],
    ["DP UX Score",                 "Number 1–5",
     "Scored by PM at Analyzing."],
    ["DP OPCO Multiplier",          "Number (read-only)",
     "Auto 1.0 / 1.3 / 1.6 set by automation A4."],
    ["DP Concept Prototype",        "URL or short text",
     "Required for Transform-class items."],
    ["DP Class of Service",         "Single-select",
     "Standard · Fixed-Date · Expedite · Intangible."],
], widths_cm=[5.0, 3.5, 8.5])

# --- 3.5 Workflow ---
heading(doc, "3.5 · Workflow — \"Demand Lifecycle Workflow\"", 2)
para(doc, "11 statuses (8 main + 3 side states) · 15 transitions "
          "(including 2 unplanned-demand channels: Severity / Emergency "
          "and Scope-add).")
para(doc, "Statuses:", bold=True)
table(doc, ["Status", "Category"], [
    ["NEW",                  "To Do"],
    ["ROUTED",               "To Do"],
    ["TRIAGED",              "In Progress"],
    ["BOUNCED",              "To Do (back to Requester)"],
    ["REVIEWING",            "In Progress"],
    ["ANALYZING",            "In Progress"],
    ["READY FOR EXECUTION",  "In Progress (replaces older \"Backlog\")"],
    ["IMPLEMENTING",         "In Progress"],
    ["DONE",                 "Done"],
    ["DEFERRED",             "Done (deferred)"],
    ["REJECTED",             "Done (rejected)"],
], widths_cm=[6.0, 11.0])

# --- 3.6 Field configurations ---
heading(doc, "3.6 · Field configurations — 6 (status-driven visibility)",
        2)
para(doc, "Six configurations control which fields are visible / "
          "editable / required at each stage. Combined with the "
          "Permission Scheme + workflow validators, they implement the "
          "\"card grows as it advances\" model.")
table(doc, ["Configuration", "Statuses", "Editable by"], [
    ["FC_FUNNEL",     "NEW",                                "Requester"],
    ["FC_AUTO_ONLY",  "ROUTED, TRIAGED",                    "(automation only)"],
    ["FC_REVIEW",     "REVIEWING",                          "Lead PM"],
    ["FC_ANALYZE",    "ANALYZING",                          "PM + Architect"],
    ["FC_DELIVERY",   "READY FOR EXECUTION, IMPLEMENTING, DONE",
                                                            "Squad PM"],
    ["FC_TERMINAL",   "DEFERRED, REJECTED, BOUNCED",        "Director / VP "
                                                            "(reason in "
                                                            "comment)"],
], widths_cm=[3.5, 7.5, 6.0])
callout(doc, "The full per-field × per-status × per-role matrix is in "
             "Section 5.8 of the Admin Guide.")

# --- 3.7 Roles & permissions ---
heading(doc, "3.7 · Project roles + permission scheme — 8 roles", 2)
table(doc, ["Role", "Initial members"], [
    ["dp_requester",     "Group jira-software-users (everyone in Avangrid)"],
    ["dp_pm_per_vs",     "PMs per VS — list to be confirmed"],
    ["dp_architect",     "Cross-VS architects — list to be confirmed"],
    ["dp_director_cs",   "kumar.jaladanki@avangrid.com · "
                         "heather.lajoie@avangrid.com"],
    ["dp_director_itba", "carlos.garbin@avangrid.com"],
    ["dp_vp",            "vitor.cruz@avangrid.com"],
    ["dp_squad_pm",      "Squad PMs per VS — owns DoR · WIP · DoD"],
    ["dp_admin",         "NTT consultant (rotates post-handover)"],
], widths_cm=[5.0, 12.0])

# --- 3.8 Boards ---
heading(doc, "3.8 · Boards — 1 Portfolio + 10 Squad", 2)
para(doc, "All 11 boards use swim lanes by Class of Service with WIP "
          "limits.")
table(doc, ["Lane", "WIP limit"], [
    ["Expedite",   "1"],
    ["Fixed-Date", "2"],
    ["Standard",   "4"],
    ["Intangible", "1"],
], widths_cm=[5.0, 12.0])

# --- 3.9 Automation rules ---
heading(doc, "3.9 · Automation rules — 5", 2)
table(doc, ["ID", "Name", "Trigger", "Effect"], [
    ["A1", "Auto-Routing",
     "Issue created (Demand Intake)",
     "Set Component from VS Destination · transition NEW → ROUTED · "
     "notify Lead PM"],
    ["A2", "Auto-Expedite",
     "Customer Impact = P1 or P2",
     "Set CoS = Expedite · WIP-1 · email Sponsor + Director · "
     "SLA 4h (P1) / 8h (P2)"],
    ["A3", "Auto-Fixed-Date",
     "Demand Type = Regulatory",
     "Set CoS = Fixed-Date · email Assignee for due date"],
    ["A4", "Auto-Pain-Score",
     "Any of BV / UX / Safety / Effort / OpCo changed",
     "Compute OPCO multiplier · compute Pain Score (ITBA or CS formula) "
     "· write read-only fields"],
    ["A5", "Auto-Convert at Triage",
     "Status → TRIAGED",
     "Size S → Story · M → Epic · L → Initiative · all linked to "
     "original Demand"],
], widths_cm=[1.0, 3.5, 4.5, 8.0])
callout(doc, "Implementation details, lookup tables (Lead PM by VS, "
             "Director by VS) and a UI walkthrough for each rule are "
             "in Section 5.10 of the Admin Guide.")

page_break(doc)


# === 4. Permissions and prerequisites ========================================
heading(doc, "4 · Permissions and prerequisites we need from the "
              "admin team", 1)
para(doc, "Beyond creating the project, we need the following items "
          "verified or enabled. None require new licenses; all are "
          "standard Jira Cloud configuration.")
table(doc, ["#", "Item", "Why"], [
    ["1", "Project Administrator role granted to the NTT consultant",
     "Enables NTT DATA to perform the rest of the configuration in "
     "Sections 3.2 to 3.9."],
    ["2", "Atlassian Automation enabled at project level on DGTLPROD",
     "Required to create the 5 automation rules. Available from Premium "
     "tier upward; if not visible, request enablement."],
    ["3", "Outbound email working at site level",
     "A1, A2 and A3 send notifications. Verify SMTP is configured "
     "(Site Admin → System → Mail → Outgoing mail)."],
    ["4", "Monthly automation quota sufficient for ~1,500 runs / month",
     "Based on 5 rules × ~3 executions × ~100 demands / month. Confirm "
     "the Avangrid tier supports this."],
    ["5", "Smart-value lookup tables enabled (preferred)",
     "A1 and A2 use lookup tables for Lead PM / Director by VS. If "
     "unavailable, NTT DATA will use hardcoded conditional branches."],
    ["6", "Cross-issue \"Create issue\" automation action allowed",
     "A5 creates child issues (Story / Epic / Initiative). Default in "
     "Cloud is yes; please confirm not blocked at site level."],
    ["7", "Audit log retention",
     "Default 90 days. If Avangrid compliance requires longer, increase "
     "in Site Admin."],
    ["8", "(Optional) Service account \"automation-bot@avangrid.com\"",
     "Recommended as the actor for all 5 rules so audit trails show "
     "\"Automation\" rather than the rule creator's name."],
], widths_cm=[0.7, 5.5, 10.8])


# === 5. Reference documents =================================================
heading(doc, "5 · Reference documents", 1)
para(doc, "Three companion artifacts are available; we will share each "
          "on request:")
bullet(doc, "Avangrid_Jira_DGTLPROD_Admin_Guide.docx — full technical "
            "guide with click-by-click UI navigation, exact field "
            "options, JQL filters, validators, conditions, automation "
            "configurations and the consolidated field-level access "
            "matrix. ~25 pages.")
bullet(doc, "TOM MAP workbook · FINAL FINAL.pptx — the canonical "
            "design deck approved by the Director group (~110 slides).")
bullet(doc, "Apply-ready JSON payloads under jira_setup/payloads/ — "
            "machine-readable artefacts (workflow, field configs, "
            "permission scheme, automation rule exports) NTT DATA can "
            "deploy via the REST API once project-admin access is "
            "granted.")


# === 6. Acceptance criteria =================================================
heading(doc, "6 · Acceptance criteria — when this request is complete",
        1)
para(doc, "We will consider the request resolved when ALL of the "
          "following are true:")
numbered(doc, "Project DGTLPROD is reachable at "
              "https://avangridjira.atlassian.net/jira/software/projects/"
              "DGTLPROD .")
numbered(doc, "The NTT consultant is in the Administrators role of "
              "DGTLPROD (visible under Project settings → People).")
numbered(doc, "An API call to GET /rest/api/3/project/DGTLPROD returns "
              "HTTP 200 with the requester's API token.")
numbered(doc, "The 8 prerequisites in Section 4 are confirmed (Jira "
              "Automation enabled, SMTP working, quota sufficient, "
              "etc.).")
numbered(doc, "Confirmation email sent to the requester so NTT DATA "
              "can begin Phases 3 onward (configuration of components, "
              "fields, workflow, automations, boards) per the Admin "
              "Guide.")


# === 7. Timeline / urgency ==================================================
heading(doc, "7 · Timeline and urgency", 1)
para(doc, "Priority: Medium.", bold=True)
para(doc, "Target completion: within 5 business days of receiving this "
          "request.")
para(doc, "Why this matters now:", bold=True)
bullet(doc, "5 Directors and ~30 PMs / Architects are scheduled to "
            "onboard onto the new Demand Lifecycle in the coming "
            "weeks.")
bullet(doc, "The next Steering Committee review is scheduled, and the "
            "TOM smoke test (5 sample demands across S / M / L · CS / "
            "ITBA · all four Class-of-Service variants) cannot run "
            "until DGTLPROD exists.")
bullet(doc, "Phase 3+ of the Jira rollout is blocked until the project "
            "shell exists and admin access is granted.")


# === 8. Contacts ============================================================
heading(doc, "8 · Contacts", 1)
para(doc, "For any questions or clarifications:", bold=True)
table(doc, ["Role", "Person", "Email"], [
    ["Primary requester",
     "NTT consultant",
     "—"],
    ["Avangrid sponsor",
     "Vitor Cruz, VP Digital Products",
     "vitor.cruz@avangrid.com"],
    ["Director · ITBA",
     "Carlos Garbin",
     "carlos.garbin@avangrid.com"],
    ["Director · CS",
     "Kumar Jaladanki",
     "kumar.jaladanki@avangrid.com"],
    ["Director · CS · CCC",
     "Heather LaJoie",
     "heather.lajoie@avangrid.com"],
    ["NTT DATA engagement lead",
     "[fill in before sending]",
     "[fill in before sending]"],
], widths_cm=[4.0, 5.5, 7.5])

para(doc, "", size=10)
callout(doc,
        "Happy to schedule a 30-minute working session with the admin "
        "team if any item above needs discussion before provisioning. "
        "Thank you for the support.")

# === Save ===================================================================
doc.save(OUT)
print(f"✅ Wrote {OUT}")
print(f"   {OUT.stat().st_size / 1024:.1f} KB")
