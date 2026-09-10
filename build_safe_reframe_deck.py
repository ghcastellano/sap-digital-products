"""
Build the SAFe-reframed Customer Services Value Streams deck.
Uses the existing 'Customer Services - Proposed Value Streams.pptx' as template
to preserve the master, background, and visual style 1:1.

Structure (new):
  1. Cover            (modified subtitle)
  2. Why reframe      (cloned-template, repurposed)
  3. SAFe framing     (cloned-template, repurposed)
  4. Intake detail    (kept from original, slightly adapted)
  5. OVS — Meter-to-Cash
  6. DVS — Meter Data & Devices
  7. DVS — Billing & Invoicing
  8. DVS — Collections, FICA & Tax
  9. DVS — Customer Interaction
 10. DVS — Platform & Enablers (shared services)
"""
from copy import deepcopy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor

SRC = "Material em andamento/Customer Services - Proposed Value Streams.pptx"
DST = "Material em andamento/Customer Services - Proposed Value Streams - SAFe Reframe v1.pptx"

# Brand palette (extracted from existing slide 4)
NAVY        = RGBColor(0x0F, 0x1B, 0x4F)
BLUE_MID    = RGBColor(0x2C, 0x57, 0x8B)
WHITE       = RGBColor(0xFF, 0xFF, 0xFF)
DARK_TEXT   = RGBColor(0x2C, 0x3E, 0x50)
GRAY        = RGBColor(0x95, 0xA5, 0xA6)
LIGHT_BG    = RGBColor(0xF5, 0xF5, 0xF5)


# ---------- helpers ----------
def clone_slide(prs, source_idx):
    """Duplicate slide at source_idx, append at end. Returns new slide."""
    src = prs.slides[source_idx]
    new = prs.slides.add_slide(src.slide_layout)
    # remove default placeholders
    for shp in list(new.shapes):
        sp = shp._element
        sp.getparent().remove(sp)
    # copy spTree contents
    for el in src.shapes._spTree:
        if el.tag.endswith('}spTree'):
            continue
        # only copy shape-like elements (sp, pic, grpSp, graphicFrame, cxnSp), skip nvGrpSpPr/grpSpPr
        tag = el.tag.split('}', 1)[-1]
        if tag in ('nvGrpSpPr', 'grpSpPr'):
            continue
        new.shapes._spTree.append(deepcopy(el))
    # copy slide background (cSld bgPr/bgRef)
    src_cSld = src._element.find('.//{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}cSld') or src._element[0]
    # simpler: copy bg element if present
    from pptx.oxml.ns import qn
    src_bg = src._element.find(qn('p:cSld') + '/' + qn('p:bg'))
    new_cSld = new._element.find(qn('p:cSld'))
    if src_bg is not None and new_cSld is not None:
        # remove existing bg if any
        existing_bg = new_cSld.find(qn('p:bg'))
        if existing_bg is not None:
            new_cSld.remove(existing_bg)
        new_cSld.insert(0, deepcopy(src_bg))
    return new


def move_slide(prs, old_idx, new_idx):
    """Move slide from old_idx to new_idx."""
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    xml_slides.remove(slides[old_idx])
    xml_slides.insert(new_idx, slides[old_idx])


def delete_slide(prs, idx):
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    rId = slides[idx].rId
    prs.part.drop_rel(rId)
    xml_slides.remove(slides[idx])


def find_textbox_by_text(slide, search):
    for shp in slide.shapes:
        if shp.has_text_frame and search in shp.text_frame.text:
            return shp
    return None


def set_text(shape, text, *, size=None, bold=None, color=None, font_name="Arial"):
    """Replace text in a textbox preserving style of first run."""
    tf = shape.text_frame
    # collect existing first-run formatting as defaults
    if tf.paragraphs and tf.paragraphs[0].runs:
        f0 = tf.paragraphs[0].runs[0].font
        size = size or (f0.size.pt if f0.size else None)
        bold = bold if bold is not None else f0.bold
        try:
            color = color or (f0.color.rgb if f0.color and f0.color.rgb else None)
        except Exception:
            pass
        font_name = f0.name or font_name
    tf.clear()
    # split into paragraphs
    lines = text.split('\n')
    for i, line in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        run = para.add_run()
        run.text = line
        f = run.font
        f.name = font_name
        if size is not None:
            f.size = Pt(size)
        if bold is not None:
            f.bold = bold
        if color is not None:
            f.color.rgb = color


def edit_vs_slide(slide, *, header_top, header_sub, owner_chip, mix_chip,
                  outcome, app_title, apps_text,
                  demand_title, demand_text, demand_footnote,
                  team_title, team_text, team_footnote,
                  vendor_title, vendor_text, vendor_footnote):
    """Edit a cloned VS slide by overwriting each known textbox."""
    # We rely on positional order matching slide 4 of the original.
    txts = [s for s in slide.shapes if s.has_text_frame]
    # Map by approximate position from original slide 4 dump:
    #  0 VALUE STREAM, 1 subtitle,
    #  3 owner chip, 5 mix chip,
    #  7 OUTCOME label, 8 outcome body,
    # 11 app title, 12 app body,
    # 15 demand title, 16 demand body, 17 demand footnote,
    # 20 team title, 21 team body, 22 team footnote,
    # 25 vendor title, 26 vendor body, 27 vendor footnote
    by_idx = {i: s for i, s in enumerate(slide.shapes)}
    set_text(by_idx[0], header_top)
    set_text(by_idx[1], header_sub)
    set_text(by_idx[3], owner_chip)
    set_text(by_idx[5], mix_chip)
    set_text(by_idx[8], outcome)
    set_text(by_idx[11], app_title)
    set_text(by_idx[12], apps_text)
    set_text(by_idx[15], demand_title)
    set_text(by_idx[16], demand_text)
    set_text(by_idx[17], demand_footnote)
    set_text(by_idx[20], team_title)
    set_text(by_idx[21], team_text)
    set_text(by_idx[22], team_footnote)
    set_text(by_idx[25], vendor_title)
    set_text(by_idx[26], vendor_text)
    set_text(by_idx[27], vendor_footnote)


# ---------- main build ----------
def main():
    prs = Presentation(SRC)

    # --- 1. Cover: change subtitle ---
    cover = prs.slides[0]
    for shp in cover.shapes:
        if shp.has_text_frame and 'PROPOSED VALUE STREAMS' in shp.text_frame.text:
            # Replace title text but keep the existing first run formatting
            tf = shp.text_frame
            # Keep paragraph[0] formatting; rewrite runs
            p0 = tf.paragraphs[0]
            # capture first run font
            if p0.runs:
                f0 = p0.runs[0].font
                size = f0.size
                bold = f0.bold
                color = None
                try:
                    color = f0.color.rgb
                except Exception:
                    pass
                name = f0.name
            else:
                size, bold, color, name = Pt(40), True, WHITE, "Arial"
            tf.clear()
            p1 = tf.paragraphs[0]
            r1 = p1.add_run()
            r1.text = "PROPOSED VALUE STREAMS – CUSTOMER SERVICES"
            r1.font.size = size
            r1.font.bold = bold
            r1.font.name = name
            if color: r1.font.color.rgb = color
            p2 = tf.add_paragraph()
            r2 = p2.add_run()
            r2.text = "v2 — SAFe Reframe: 1 Operational + 5 Development Value Streams"
            r2.font.size = Pt(16)
            r2.font.bold = False
            r2.font.name = "Arial"
            if color: r2.font.color.rgb = color
            p3 = tf.add_paragraph()
            r3 = p3.add_run()
            r3.text = "April 2026"
            r3.font.size = Pt(14)
            r3.font.name = "Arial"
            if color: r3.font.color.rgb = color

    # --- 2. Adapt slide 3 (INTAKE DETAIL) — change "3 Value Streams" reference ---
    intake = prs.slides[2]
    for shp in intake.shapes:
        if shp.has_text_frame:
            t = shp.text_frame.text
            if "3 Value Streams" in t and "M2C-SL" in t:
                new_t = t.replace(
                    "3 Value Streams — M2C-SL ·    M2C-RO · CCC",
                    "1 OVS (Meter-to-Cash) + 5 DVS — Meter · Bill · Collect · CX · Platform"
                ).replace(
                    "3 Value Streams — M2C-SL · M2C-RO · CCC",
                    "1 OVS (Meter-to-Cash) + 5 DVS — Meter · Bill · Collect · CX · Platform"
                )
                set_text(shp, new_t)
            if "ONE Jira request type per Value Stream (M2C-SL · M2C-RO · CCC)" in t:
                new_t = t.replace(
                    "ONE Jira request type per Value Stream (M2C-SL · M2C-RO · CCC)",
                    "ONE Jira request type per DVS (Meter · Bill · Collect · CX · Platform)"
                )
                set_text(shp, new_t)
            if "Cross-VS conflicts (e.g. CRM ↔ Smart Metering)" in t:
                new_t = t.replace(
                    "Cross-VS conflicts (e.g. CRM ↔ Smart Metering) resolved at Customer Journey Forum",
                    "Cross-DVS conflicts (e.g. CX ↔ Meter) resolved at the Meter-to-Cash OVS Forum"
                )
                set_text(shp, new_t)

    # --- 3. Reuse the 3 VS template slides + clone 3 more, no deletions ---
    # original indices: 0=cover, 1=domain, 2=intake, 3=VS-SL, 4=VS-RO, 5=VS-CCC
    # Reuse slides 3,4,5 for OVS, DVS1, DVS2.  Clone 3 more for DVS 3, 4, 5.
    dvs_slides_extra = [clone_slide(prs, 3) for _ in range(3)]  # DVS 3, 4, 5

    # Now: edit original slide 3 → OVS Meter-to-Cash
    ovs = prs.slides[3]
    edit_vs_slide(
        ovs,
        header_top="OPERATIONAL VALUE STREAM (OVS)",
        header_sub="METER-TO-CASH — END-TO-END BUSINESS OUTCOME EVERY OPCO CUSTOMER EXPERIENCES",
        owner_chip="CS Apps · Kumar Jaladanki · OVS owned cross-DVS",
        mix_chip="SAFe framing · OVS = how the business delivers value",
        outcome=("End-to-end flow from meter read → invoice → cash collection → customer interaction "
                 "across 7 OPCOs, with regulatory compliance per state PUC (CT · ME · NY)."),
        app_title="STAGES OF THE OVS (M2C)",
        apps_text=("•  METER — devices · data acquisition · MDM · regulatory reads\n"
                   "•  BILL — invoice generation · pricing · formatting · address validity\n"
                   "•  COLLECT — payments · FICA · dunning · tax compliance\n"
                   "•  SERVE — customer interaction · CRM · IVR · feedback · self-service\n"
                   "•  PLATFORM — Basis · data mgmt · RPA · test · ALM (transversal enabler)"),
        demand_title="WHY IT IS THE OPERATIONAL VS",
        demand_text=("•  Customer-facing flow that defines revenue and CSAT\n"
                     "•  SLAs are external (regulatory rate cases · billing windows · ANI containment)\n"
                     "•  Failure in any stage = customer + revenue + regulatory triple hit [KB: AS-IS S.23]\n"
                     "•  Anchored on SAP (Billing · FICA · CRM · Metering) + supporting stack\n"
                     "•  Already labelled as the M2C value chain in the AS-IS deliverable [S.23]"),
        demand_footnote="OVS = the flow customers see; DVS = the engineering pipelines that build & evolve it",
        team_title="OVS GOVERNANCE (cross-DVS)",
        team_text=("•  OVS Owner — Kumar Jaladanki (Director CS Apps)\n"
                   "•  Product Lead pairing — Heather LaJoie (CS Products) for outcome alignment\n"
                   "•  M2C Forum — monthly cross-DVS prioritisation and conflict resolution\n"
                   "•  Architecture coherence — Sr. Manager SAP Architecture (vacancy → fill) [KB S.23]\n"
                   "•  Flow metrics — lead time, throughput, customer-impact P1–P4 (per DVS)"),
        team_footnote="Forum participants: 5 DVS Product Managers + Architecture + CS Products",
        vendor_title="OVS-LEVEL DEPENDENCIES",
        vendor_text=("•  S/4 HANA migration — multi-year program touching all 7 OPCOs [KB AS-IS S.23]\n"
                     "•  PUC compliance per state — CT PURA · ME PUC · NY PSC rate cases\n"
                     "•  Smart-meter rollout mandates — CT · ME · NY\n"
                     "•  Vendor concentration — Smart Metering 100% vendor-delivered [KB S.23]\n"
                     "•  OPCO-specific demand — daily multi-OPCO contention [KB S.23]"),
        vendor_footnote="OVS = business outcome layer · DVS slides 6–10 detail the 5 product clusters",
    )

    # Edit each DVS clone (now at indices 6..10 because we kept original 3,4,5)
    # Wait: original indices 4 (VS-RO) and 5 (VS-CCC) still exist; we'll delete them later.
    # Cloned dvs slides are at len(prs.slides)-5 .. -1, but enumerate via dvs_slides list.

    # DVS 1 — Meter Data & Devices  (reuse original slide 4)
    edit_vs_slide(
        prs.slides[4],
        header_top="DEVELOPMENT VALUE STREAM (DVS) 1 / 5",
        header_sub="METER DATA & DEVICES — CAPTURE EVERY READ ACCURATELY · ZERO DATA LOSS TO BILLING",
        owner_chip="CS Apps · Vincent Lam (Smart Metering · Iram Chowdhury Champion)",
        mix_chip="Run · Grow · Transform · 50 / 30 / 20  [estimate]",
        outcome=("Acquire meter reads from devices, manage device lifecycle, and feed billing on time. "
                 "Resolve device-to-data exceptions before they reach the customer."),
        app_title="APPLICATIONS SUPPORTED",
        apps_text=("•  Smart Metering platform (AMI head-end)\n"
                   "•  Meter Data Management (MDM)\n"
                   "•  SAP IS-U Device Management\n"
                   "•  Group 8760 (EDI) — device events / B2B exchanges\n"
                   "•  Field meter read interfaces"),
        demand_title="DEMAND SOURCES — who sends the work",
        demand_text=("•  Smart Metering program (CT · ME · NY rollout)\n"
                     "•  Field Operations exceptions and escalations\n"
                     "•  Regulatory drivers (CT PURA · ME PUC · NY PSC mandates)\n"
                     "•  AMI vendor release schedule"),
        demand_footnote="Routed via the Jira request type for this DVS",
        team_title="AVANGRID CORE TEAM",
        team_text=("•  Product Manager (NEW)\n"
                   "•  Tech Lead — Vincent Lam / Iram Chowdhury (Champion)\n"
                   "•  Business Analyst\n"
                   "•  SMEs from Smart Metering program\n"
                   "≈ 4–5 FTE"),
        team_footnote="[GAP] Smart Metering = 100% vendor-delivered → owner needs explicit KT clauses [KB S.23]",
        vendor_title="VENDOR EXTENSION (delivery)",
        vendor_text=("•  Smart Metering vendor (stable team — Option A model)\n"
                     "•  Itron managed services\n"
                     "•  SAP partner for IS-U Device\n"
                     "•  EDI integration partner"),
        vendor_footnote="Avangrid = product owner · vendor = delivery muscle",
    )

    # DVS 2 — Billing & Invoicing  (reuse original slide 5)
    edit_vs_slide(
        prs.slides[5],
        header_top="DEVELOPMENT VALUE STREAM (DVS) 2 / 5",
        header_sub="BILLING & INVOICING — RUN THE BILL ON TIME · EVERY DAY · EVERY OPCO",
        owner_chip="CS Apps · Vincent Lam (Sr. Manager Billing, Device & Retail)",
        mix_chip="Run · Grow · Transform · 65 / 25 / 10  [estimate]",
        outcome=("Generate accurate invoices for every customer of the 7 OPCOs on a hard daily cadence. "
                 "Format, validate addresses and deliver via every required channel."),
        app_title="APPLICATIONS SUPPORTED",
        apps_text=("•  SAP ECC (FI/SD modules — back-office billing)\n"
                   "•  Streamserve / OpenText (invoice formatting · outbound notifications)\n"
                   "•  Address Validation (Precisely · Pitney Bowes · CASS/NCOA)\n"
                   "•  Print & e-delivery interfaces"),
        demand_title="DEMAND SOURCES — who sends the work",
        demand_text=("•  BRS — small/medium business demand (price changes · tariff types · dunning rules) [KB S.23]\n"
                     "•  PMO — capex projects (S/4 HANA migration · CRM upgrades) [KB S.23]\n"
                     "•  IT Now — billing-run incidents (production support)\n"
                     "•  Regulatory affairs — PUC rate-case calendar locks"),
        demand_footnote="Routed via the Jira request type for this DVS",
        team_title="AVANGRID CORE TEAM",
        team_text=("•  Product Manager (NEW)\n"
                   "•  Tech Lead — Vincent Lam (or designate)\n"
                   "•  Business Analyst — billing rules SME\n"
                   "•  SMEs from each OPCO billing operations\n"
                   "≈ 5–6 FTE (highest run-load)"),
        team_footnote="[INFERENCE] FTE skewed higher: every billing day is a hard deadline [KB S.23]",
        vendor_title="VENDOR EXTENSION (delivery)",
        vendor_text=("•  Accenture (large projects + incident support)\n"
                     "•  Software Factory (SWF) — execution muscle\n"
                     "•  OpenText / Streamserve partner\n"
                     "•  Precisely (Pitney Bowes) services"),
        vendor_footnote="Avangrid = product owner · vendor = delivery muscle",
    )

    # DVS 3 — Collections, FICA & Tax  (cloned)
    edit_vs_slide(
        dvs_slides_extra[0],
        header_top="DEVELOPMENT VALUE STREAM (DVS) 3 / 5",
        header_sub="COLLECTIONS · FICA · TAX — CONVERT INVOICES INTO CASH · COMPLY AT LINE LEVEL",
        owner_chip="CS Apps · Ivelisse Velazquez (Manager FICA & Collection)",
        mix_chip="Run · Grow · Transform · 60 / 30 / 10  [estimate]",
        outcome=("Process payments, run collections and dunning, apply tax compliance accurately and on time. "
                 "Reduce DSO and protect against audit/regulatory risk."),
        app_title="APPLICATIONS SUPPORTED",
        apps_text=("•  SAP FICA (financial accounting · contract accounting)\n"
                   "•  SOVOS Taxware (indirect tax · sales/use · returns)\n"
                   "•  Payment processing interfaces (banks · lockbox · 3rd-party trading partners)\n"
                   "•  Dunning & collections workflows"),
        demand_title="DEMAND SOURCES — who sends the work",
        demand_text=("•  Tax department — jurisdiction tables · content updates · audit support\n"
                     "•  Finance / Billing Ops — collections KPI shifts · DSO targets\n"
                     "•  Regulatory — tariff content updates per state\n"
                     "•  IT Now — payment processing / tax run incidents"),
        demand_footnote="Routed via the Jira request type for this DVS",
        team_title="AVANGRID CORE TEAM",
        team_text=("•  Product Manager (NEW)\n"
                   "•  Tech Lead — Ivelisse Velazquez (or designate)\n"
                   "•  Business Analyst — FICA / Tax SME\n"
                   "•  Finance/Tax dept liaison (matrixed)\n"
                   "≈ 3–4 FTE"),
        team_footnote="\"The Cash of M2C\" — invoicing → payment → collections → dunning across 7 OPCOs [KB S.23]",
        vendor_title="VENDOR EXTENSION (delivery)",
        vendor_text=("•  SOVOS Taxware (content updates · cloud)\n"
                     "•  Accenture (FICA support · projects)\n"
                     "•  Software Factory (SWF) — custom development\n"
                     "•  Banking integration partners"),
        vendor_footnote="Avangrid = product owner · vendor = delivery muscle",
    )

    # DVS 4 — Customer Interaction (CX)  (cloned)
    edit_vs_slide(
        dvs_slides_extra[1],
        header_top="DEVELOPMENT VALUE STREAM (DVS) 4 / 5",
        header_sub="CUSTOMER INTERACTION — UNIFY CRM · CONTACT CENTER · DIGITAL · FEEDBACK",
        owner_chip="CS Apps · Sr. Manager CRM & Contact Center Ops (vacancy → fill) [KB S.23]",
        mix_chip="Run · Grow · Transform · 45 / 35 / 20  [estimate]",
        outcome=("Every customer interaction lands in a single record with a clear next step. "
                 "Improve IVR containment, reduce AHT, raise CSAT/NPS."),
        app_title="APPLICATIONS SUPPORTED",
        apps_text=("•  SAP CRM (front-office · case · order capture)\n"
                   "•  Genesys (Cloud / Engage — Infosys managed)\n"
                   "•  Nuance IVR / CTI (speech IVR · NLU)\n"
                   "•  Medallia (formerly VHT — callback · feedback · CSAT)\n"
                   "•  PureConnect / PureEngage → Now Infosys (legacy CC) [phased out]"),
        demand_title="DEMAND SOURCES — who sends the work",
        demand_text=("•  Customer Operations (Heather LaJoie · CS Products)\n"
                     "•  Marketing & Communications — campaign / digital channel demand\n"
                     "•  OPCO call centers (UIL · CMP · NY) — escalations and trends\n"
                     "•  Customer Service GenAI initiative pipeline"),
        demand_footnote="Routed via the Jira request type for this DVS",
        team_title="AVANGRID CORE TEAM",
        team_text=("•  Product Manager (NEW)\n"
                   "•  Tech Lead — vacancy (CRM & Contact Center Ops) [KB S.23]\n"
                   "•  Business Analyst — CX SME\n"
                   "•  Solution architect (matrixed)\n"
                   "≈ 4–5 FTE"),
        team_footnote="[GAP] confirm sub-team lead post-vacancy fill",
        vendor_title="VENDOR EXTENSION (delivery)",
        vendor_text=("•  Lemur-Salesforce (CRM delivery)\n"
                     "•  Infosys (Genesys managed services)\n"
                     "•  HCL-Nuance (IVR vendor)\n"
                     "•  Medallia partner\n"
                     "•  Digital channel agency"),
        vendor_footnote="Avangrid = product owner · vendor = delivery muscle",
    )

    # DVS 5 — Platform & Enablers  (cloned)
    edit_vs_slide(
        dvs_slides_extra[2],
        header_top="DEVELOPMENT VALUE STREAM (DVS) 5 / 5  ·  SHARED PLATFORM",
        header_sub="PLATFORM & ENABLERS — KEEP THE FACTORY RUNNING FOR EVERY OTHER DVS",
        owner_chip="CS Apps · Naggi Reddy (SAP Basis) + Sr. Manager SAP Architecture (vacancy)",
        mix_chip="Run · Grow · Transform · 70 / 25 / 5  [estimate · shared svc]",
        outcome=("Provide the technical platform, data management, automation, testing and ALM tooling "
                 "that every other DVS depends on. Without it, no bill runs."),
        app_title="APPLICATIONS SUPPORTED",
        apps_text=("•  SAP Basis (landscape · transports · security · performance)\n"
                   "•  SAP Gold Client (selective non-prod data copy · masking)\n"
                   "•  SAP Enable Now (formerly WPB — training & in-app guidance)\n"
                   "•  Blue Prism (RPA for back-office CS processes)\n"
                   "•  SAP Tricentis TOSCA (test automation)\n"
                   "•  Jira (demand · backlog · agile work mgmt)"),
        demand_title="DEMAND SOURCES — who sends the work",
        demand_text=("•  Other 4 DVS — platform requests · environment refresh · test cycles\n"
                     "•  Release management — release readiness · regression coverage\n"
                     "•  Security / compliance — vault · access · audit trail\n"
                     "•  Process owners — RPA opportunities"),
        demand_footnote="Treated as Shared Service · capacity reserved per quarter for each DVS",
        team_title="AVANGRID CORE TEAM",
        team_text=("•  Platform Lead — Naggi Reddy (SAP Basis)\n"
                   "•  Architecture — Sr. Manager SAP Architecture (vacancy → fill) [KB S.23]\n"
                   "•  RPA / Test / ALM SMEs (matrixed across DVS)\n"
                   "•  No dedicated Product Manager — managed as enabling team\n"
                   "≈ 4–5 FTE"),
        team_footnote="[INFERENCE] Treated as SAFe enabling team, not a product DVS — funded transversally",
        vendor_title="VENDOR EXTENSION (delivery)",
        vendor_text=("•  Tricentis (test platform)\n"
                     "•  Blue Prism vendor + Automation CoE\n"
                     "•  Atlassian Jira admins\n"
                     "•  SAP partner for Basis/Gold Client"),
        vendor_footnote="Cross-cuts every DVS · no DVS can deliver in isolation from this layer",
    )

    # --- Add "Why reframe" + "SAFe framing" by cloning the VS template ---
    # Current order: 0=cover, 1=domain, 2=intake, 3=OVS, 4=DVS1, 5=DVS2, 6=DVS3, 7=DVS4, 8=DVS5
    # Clone slide 3 twice → appended as idx 9 (why), 10 (framing)
    why_slide = clone_slide(prs, 3)
    framing_slide = clone_slide(prs, 3)

    # Edit "Why reframe"
    edit_vs_slide(
        why_slide,
        header_top="WHY THIS REFRAME",
        header_sub="17 CS APPLICATIONS · ~24 FTE · 1 APP ≠ 1 PRODUCT — SAFe SEPARATES OUTCOME FROM DELIVERY",
        owner_chip="CS Apps · Kumar Jaladanki · 5 sub-teams · 2 Sr. Manager vacancies [KB S.23 / S.2]",
        mix_chip="Source: AS-IS deliverable · 2026-04",
        outcome=("Treating each application as its own product would split ~24 FTE across 17 apps "
                 "(~1.4 FTE each). Unsustainable. SAFe gives us a way to keep one customer-facing "
                 "outcome (M2C) while organising delivery around a small number of product clusters."),
        app_title="THE CONSTRAINT — WHAT WE HAVE",
        apps_text=("•  ~24 FTE in CS Apps across 5 sub-teams [KB AS-IS S.2 / S.23]\n"
                   "•  2 Sr. Manager vacancies (CRM & CC Ops · SAP Architecture) [KB S.23]\n"
                   "•  17 distinct CS applications in scope (per landscape table)\n"
                   "•  Heterogeneous maturity across sub-teams [KB S.23]\n"
                   "•  Smart Metering = 100% vendor-delivered (knowledge risk) [KB S.23]"),
        demand_title="THE PRINCIPLE — FROM AS-IS TO-BE",
        demand_text=("•  AS-IS: organised by application silo · not by Value Stream [KB S.42]\n"
                     "•  TO-BE: \"organising around products with clear ownership and alignment to "
                     "business OKRs\" [KB S.42]\n"
                     "•  Product ≠ application — product = a cluster of capability serving an outcome\n"
                     "•  ⇒ map 17 apps → 5 product clusters · plus 1 OVS that wraps them"),
        demand_footnote="Source slides: AS-IS deliverable Slides 2, 23, 42",
        team_title="THE SAFe LEVER",
        team_text=("•  Operational Value Stream (OVS) = the business flow customers experience\n"
                   "•  Development Value Stream (DVS) = the engineering pipeline that builds it\n"
                   "•  This separation lets us keep ONE OVS owner (Kumar) AND tight DVS focus\n"
                   "•  Product Managers belong to DVS · OVS owner orchestrates across DVS\n"
                   "•  Aligns with SAFe Lean Portfolio Mgmt already referenced in NextGen materials"),
        team_footnote="Reframe = same headcount · clearer ownership · less product fragmentation",
        vendor_title="WHAT CHANGES vs. v1 (3 VS proposal)",
        vendor_text=("•  v1: 3 Value Streams (M2C-SL · M2C-RO · CCC) treated as a flat list\n"
                     "•  v2: 1 OVS (Meter-to-Cash) + 5 DVS — explicit SAFe layering\n"
                     "•  Bill and Collect now distinct DVS (different SLAs · skills · vendors)\n"
                     "•  Platform & Enablers explicit as Shared Service (no overhead PM)\n"
                     "•  Customer Interaction consolidates CRM + IVR + Genesys + Medallia"),
        vendor_footnote="Net: same FTE base · 5 product clusters instead of 17 apps · OVS owns the outcome",
    )

    # Edit "SAFe framing overview"
    edit_vs_slide(
        framing_slide,
        header_top="SAFe FRAMING OVERVIEW",
        header_sub="ONE OPERATIONAL VS WRAPS FIVE DEVELOPMENT VS — OUTCOME LAYER + DELIVERY LAYER",
        owner_chip="OVS owner · Kumar Jaladanki  ·  5 DVS Product Managers (NEW)",
        mix_chip="SAFe · OVS-DVS pattern",
        outcome=("Operational Value Stream = how the business delivers value to the customer "
                 "(Meter-to-Cash). Development Value Streams = how IT builds and evolves the digital "
                 "products that enable that flow."),
        app_title="OPERATIONAL VALUE STREAM (1)",
        apps_text=("•  M2C — Meter-to-Cash\n"
                   "    – Stages: METER → BILL → COLLECT → SERVE (+ PLATFORM transversal)\n"
                   "    – Customer-facing · regulated · revenue-bearing\n"
                   "    – Owner: Director CS Apps (Kumar Jaladanki)\n"
                   "    – Forum: monthly cross-DVS prioritisation"),
        demand_title="DEVELOPMENT VALUE STREAMS (5)",
        demand_text=("•  DVS 1 — Meter Data & Devices  (≈4–5 FTE)\n"
                     "•  DVS 2 — Billing & Invoicing  (≈5–6 FTE)\n"
                     "•  DVS 3 — Collections, FICA & Tax  (≈3–4 FTE)\n"
                     "•  DVS 4 — Customer Interaction  (≈4–5 FTE)\n"
                     "•  DVS 5 — Platform & Enablers  (Shared Service, ≈4–5 FTE)"),
        demand_footnote="FTE figures = inference based on ~24 FTE total · to validate with Kumar",
        team_title="MAPPING — 17 APPS → 5 DVS",
        team_text=("•  Meter   → Smart Metering · MDM · IS-U Device · Group 8760 (EDI)\n"
                   "•  Bill    → SAP ECC (FI/SD) · Streamserve/OpenText · Address Validation\n"
                   "•  Collect → SAP FICA · SOVOS Taxware\n"
                   "•  Serve   → SAP CRM · Genesys · Nuance IVR · Medallia · PureConnect (legacy)\n"
                   "•  Platform→ Basis · Gold Client · Enable Now · Blue Prism · TOSCA · Jira"),
        team_footnote="Same 17 apps · grouped by capability · 1 PM per DVS (4 PMs · Platform = enabling team)",
        vendor_title="WHY THIS WORKS WITHIN HEADCOUNT",
        vendor_text=("•  4 dedicated Product Managers (one per outcome-bearing DVS)\n"
                     "•  Platform & Enablers run as enabling team — no extra PM\n"
                     "•  ~4–5 FTE per DVS = real squad size · feasible with current headcount\n"
                     "•  OVS owner (Kumar) = existing role · no new layer\n"
                     "•  Vendor extension model (Option A) preserved per DVS"),
        vendor_footnote="Validates against AS-IS constraint: 5 sub-teams · 24 FTE · 2 vacancies [KB S.23]",
    )

    # Move the two new framing slides to positions 2 and 3 (after cover & domain).
    # Current order: 0=cover, 1=domain, 2=intake, 3=OVS, 4=DVS1..8=DVS5, 9=why, 10=framing
    # Target:        0=cover, 1=domain, 2=why, 3=framing, 4=intake, 5=OVS, 6..10=DVS
    move_slide(prs, 9, 2)
    move_slide(prs, 10, 3)

    prs.save(DST)
    print(f"Saved → {DST}")
    print(f"Total slides: {len(prs.slides)}")
    for i, s in enumerate(prs.slides, 1):
        title = ""
        for shp in s.shapes:
            if shp.has_text_frame and shp.text_frame.text.strip():
                title = shp.text_frame.text.strip().split('\n')[0][:80]
                break
        print(f"  {i:2d}. {title}")


if __name__ == "__main__":
    main()
