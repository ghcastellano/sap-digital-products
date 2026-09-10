"""
v3 → v4: revert Foundation Platform from DVS to Enabling Team.
  1. Text-level swaps: "5 DVS" → "4 DVS + 1 enabling team", "DVS X / 5" → "DVS X / 4"
  2. Rebuild slide 3 (cards) — 4 DVS cards (wider) + 1 enabling-team band below
  3. Rebuild slide 12 (was DVS 5) → ENABLING TEAM detail (no PM, governed by Dir + SAP Arch)
  4. Update big-picture slide bottom band: "Enabling Team — shared services"
"""
import re
from copy import deepcopy
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

SRC = "Material em andamento/Customer Services - Proposed Value Streams - v3.pptx"
DST = "Material em andamento/Customer Services - Proposed Value Streams - v4.pptx"

NAVY  = RGBColor(0x0F, 0x1B, 0x4F)
BLUE2 = RGBColor(0x2C, 0x57, 0x8B)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK  = RGBColor(0x2C, 0x3E, 0x50)
GRAY  = RGBColor(0x95, 0xA5, 0xA6)
ENAB  = RGBColor(0x6C, 0x7A, 0x89)  # neutral grey for enabling team

# Order matters — most-specific first
TEXT_SWAPS = [
    # DVS slide titles
    ("DEVELOPMENT VALUE STREAM (DVS) 1 / 5", "DEVELOPMENT VALUE STREAM (DVS) 1 / 4"),
    ("DEVELOPMENT VALUE STREAM (DVS) 2 / 5", "DEVELOPMENT VALUE STREAM (DVS) 2 / 4"),
    ("DEVELOPMENT VALUE STREAM (DVS) 3 / 5", "DEVELOPMENT VALUE STREAM (DVS) 3 / 4"),
    ("DEVELOPMENT VALUE STREAM (DVS) 4 / 5", "DEVELOPMENT VALUE STREAM (DVS) 4 / 4"),
    # Counts
    ("DEVELOPMENT VALUE STREAMS (5)", "DEVELOPMENT VALUE STREAMS (4)"),
    ("5 Development Value Streams", "4 Development Value Streams"),
    ("5 DVS Product Managers (NEW)", "4 DVS Product Managers (NEW)"),
    ("5 DVS Product Managers", "4 DVS Product Managers"),
    # Subtitles / cover
    ("v2 — Reframe: 1 Operational + 5 Development Value Streams",
     "v3 — Reframe: 1 Operational + 4 Development Value Streams + 1 Enabling Team"),
    # Intake / framing references
    ("1 OVS (Meter-to-Cash) + 5 DVS — Meter · Bill · Collect · CX · Platform",
     "1 OVS (Meter-to-Cash) + 4 DVS — Meter · Bill · Collect · CX  (+ enabling team)"),
    ("ONE Jira request type per DVS (Meter · Bill · Collect · CX · Platform)",
     "ONE Jira request type per DVS (Meter · Bill · Collect · CX)"),
    # OVS STAGES line
    ("METER → BILL → COLLECT → SERVE · FOUNDATION PLATFORM (transversal)",
     "METER → BILL → COLLECT → SERVE   (Enabling Team supports all stages)"),
    ("METER → BILL → COLLECT → SERVE · PLATFORM (transversal enabler)",
     "METER → BILL → COLLECT → SERVE   (Enabling Team supports all stages)"),
    # Stages bullet in OVS detail
    ("• PLATFORM — Basis · data mgmt · RPA · test · ALM (transversal enabler across all stages)",
     "(Enabling Team — Basis · data mgmt · RPA · test · ALM — supports all stages)"),
    (" – Stages: METER → BILL → COLLECT → SERVE (+ PLATFORM transversal)",
     " – Stages: METER → BILL → COLLECT → SERVE  (Enabling Team supports all stages)"),
    # Framing slide DVS bullet list
    ("• DVS 5 — Foundation Platform (Shared Service, Squad size · to validate)",
     "+ Enabling Team — Foundation services (size to validate · no PM)"),
    ("• DVS 5 — Foundation Platform", "+ Enabling Team — Foundation services"),
    ("• Platform→ Basis · Gold Client · Enable Now · Blue Prism · TOSCA · Jira",
     "+ Enabling Team supports all 4 with: Basis · Gold Client · Enable Now · Blue Prism · TOSCA · Jira"),
    ("• v2: 1 OVS (Meter-to-Cash) + 5 DVS — explicit layering",
     "• v3: 1 OVS (Meter-to-Cash) + 4 DVS + 1 enabling team — explicit layering"),
    ("• Platform & Enablers explicit as Shared Service (no overhead PM)",
     "• Foundation services run as a small enabling team — no DVS-level PM, no fake product"),
    ("• 4 dedicated Product Managers (one per outcome-bearing DVS)",
     "• 4 dedicated Product Managers (one per DVS)"),
    ("• Platform & Enablers run as enabling team — no extra PM",
     "• Foundation services = enabling team · no extra PM"),
    ("• ~4–5 FTE per DVS = real squad size · feasible with current headcount",
     "• Each DVS gets a real squad size — feasible within current headcount"),
    ("• Customer Interaction consolidates CRM + IVR + Genesys + Medallia",
     "• Customer Care & Engagement consolidates CRM + IVR + Genesys + Medallia"),
    # Cross-DVS forum naming
    ("Cross-DVS conflicts (e.g. CX ↔ Meter) resolved at the Meter-to-Cash OVS Forum",
     "Cross-DVS conflicts (e.g. CX ↔ Meter) resolved at the Meter-to-Cash OVS Forum"),
]


def walk_replace(prs):
    n = 0
    for slide in prs.slides:
        for sh in slide.shapes:
            if not sh.has_text_frame:
                continue
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    new = r.text
                    for old, repl in TEXT_SWAPS:
                        if old in new:
                            new = new.replace(old, repl)
                    new = re.sub(r" {2,}", " ", new).rstrip()
                    if new != r.text:
                        r.text = new
                        n += 1
    return n


def add_text(slide, x, y, w, h, text, *, size=10, bold=False, italic=False,
             color=DARK, name="Arial", align=PP_ALIGN.LEFT, anchor=None, wrap=True):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.margin_left = Inches(0.04); tf.margin_right = Inches(0.04)
    tf.margin_top = Inches(0.02); tf.margin_bottom = Inches(0.02)
    if anchor: tf.vertical_anchor = anchor
    for i, line in enumerate(text.split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run(); r.text = line
        r.font.name = name; r.font.size = Pt(size)
        r.font.bold = bold; r.font.italic = italic
        r.font.color.rgb = color
    return tb


def add_dvs_card(slide, x, y, w, h, title, domain, desc, apps_label, apps,
                 *, font_title=11, font_apps=8.5, header_color=NAVY):
    body = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  Inches(x), Inches(y + 0.08), Inches(w), Inches(h - 0.08))
    body.fill.solid(); body.fill.fore_color.rgb = WHITE
    body.line.color.rgb = NAVY; body.line.width = Pt(0.5)
    body.adjustments[0] = 0.05

    hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 Inches(x - 0.02), Inches(y), Inches(w + 0.04), Inches(0.78))
    hdr.fill.solid(); hdr.fill.fore_color.rgb = header_color
    hdr.line.fill.background()
    hdr.adjustments[0] = 0.10

    add_text(slide, x, y + 0.10, w, 0.62, title,
             size=font_title, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
             anchor=MSO_ANCHOR.MIDDLE)

    inner_y = y + 0.92
    add_text(slide, x + 0.10, inner_y, w - 0.20, 0.20, domain,
             size=8, bold=True, italic=True, color=DARK, align=PP_ALIGN.CENTER)
    add_text(slide, x + 0.10, inner_y + 0.22, w - 0.20, 0.55, desc,
             size=10, italic=True, color=DARK, align=PP_ALIGN.CENTER)

    ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                Inches(x + 0.20), Inches(inner_y + 0.85),
                                Inches(w - 0.40), Inches(0.02))
    ln.fill.solid(); ln.fill.fore_color.rgb = NAVY
    ln.line.fill.background()

    add_text(slide, x + 0.10, inner_y + 0.93, w - 0.20, 0.20, apps_label,
             size=8, bold=True, color=DARK, align=PP_ALIGN.CENTER)

    apps_y = inner_y + 1.15
    add_text(slide, x + 0.10, apps_y, w - 0.20, y + h - apps_y - 0.10,
             apps, size=font_apps, color=DARK, align=PP_ALIGN.CENTER)


def rebuild_cards_slide(prs):
    """Slide 3 — 1 OVS card + 4 DVS cards + 1 enabling-team strip."""
    slide = prs.slides[2]  # 0-idx slide 3 (cards)

    # remove all shapes past header (idx >= 2)
    for sh in list(slide.shapes)[2:]:
        sp = sh._element
        sp.getparent().remove(sp)

    # OVS card
    ovs_x, ovs_y, ovs_w, ovs_h = 0.50, 1.05, 12.33, 1.85

    body = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  Inches(ovs_x), Inches(ovs_y + 0.05),
                                  Inches(ovs_w), Inches(ovs_h - 0.05))
    body.fill.solid(); body.fill.fore_color.rgb = WHITE
    body.line.color.rgb = NAVY; body.line.width = Pt(0.5)
    body.adjustments[0] = 0.04

    hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 Inches(ovs_x - 0.02), Inches(ovs_y),
                                 Inches(ovs_w + 0.04), Inches(0.50))
    hdr.fill.solid(); hdr.fill.fore_color.rgb = NAVY
    hdr.line.fill.background()
    hdr.adjustments[0] = 0.06

    add_text(slide, ovs_x + 0.20, ovs_y + 0.08, ovs_w - 3.10, 0.40,
             "Meter-to-Cash", size=13, bold=True, color=WHITE)
    add_text(slide, ovs_x + ovs_w - 3.10, ovs_y + 0.10, 2.90, 0.36,
             "OPERATIONAL VALUE STREAM (1)", size=9, bold=True, color=WHITE,
             align=PP_ALIGN.RIGHT)

    inner_y = ovs_y + 0.62
    add_text(slide, ovs_x + 0.20, inner_y, ovs_w - 0.40, 0.20,
             "DOMAIN · Customer Service · End-to-End Business Outcome across 7 OPCOs",
             size=8.5, bold=True, italic=True, color=DARK)
    add_text(slide, ovs_x + 0.20, inner_y + 0.22, ovs_w - 0.40, 0.34,
             "Capture every meter read · invoice on time · collect cash · serve every customer "
             "interaction — anchored on SAP (Billing · FICA · CRM · Metering)",
             size=11, italic=True, color=DARK)

    ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                Inches(ovs_x + 0.20), Inches(inner_y + 0.62),
                                Inches(ovs_w - 0.40), Inches(0.02))
    ln.fill.solid(); ln.fill.fore_color.rgb = NAVY
    ln.line.fill.background()
    add_text(slide, ovs_x + 0.20, inner_y + 0.68, ovs_w - 0.40, 0.20,
             "STAGES", size=8.5, bold=True, color=DARK)
    add_text(slide, ovs_x + 0.20, inner_y + 0.90, ovs_w - 0.40, 0.30,
             "METER  →  BILL  →  COLLECT  →  SERVE   (Enabling Team supports all stages)",
             size=11, bold=True, color=NAVY)

    # DVS row label
    add_text(slide, 0.50, 3.05, 12.33, 0.22,
             "DEVELOPMENT VALUE STREAMS (4)  ·  ONE PRODUCT MANAGER PER DVS",
             size=10, bold=True, color=NAVY)

    # 4 DVS cards (wider now)
    dvs_y = 3.32
    dvs_h = 3.20
    n = 4
    margin = 0.50
    gap = 0.12
    total_w = 13.33 - 2 * margin
    card_w = (total_w - (n - 1) * gap) / n  # ≈ 3.05

    dvs_cards = [
        ("Meter & Device Operations",
         "DOMAIN · Metering",
         "Capture every read · zero data loss to billing",
         "Smart Metering\nMDM (Meter Data Mgmt)\nSAP IS-U Device\nGroup 8760 (EDI)\nField meter read interfaces"),
        ("Billing Operations",
         "DOMAIN · Billing",
         "Run the bill on time · every day · every OPCO",
         "SAP ECC (FI/SD)\nStreamserve / OpenText\nAddress Validation (Precisely)\nPrint & e-delivery interfaces"),
        ("Revenue & Collections",
         "DOMAIN · Finance",
         "Convert invoices into cash · comply at line level",
         "SAP FICA\nSOVOS Taxware\nPayment processing interfaces (banks · lockbox · 3rd-party)\nDunning & collections workflows"),
        ("Customer Care & Engagement",
         "DOMAIN · CX",
         "Unify CRM · contact center · digital · feedback",
         "SAP CRM\nGenesys (Infosys)\nNuance IVR / CTI\nMedallia (formerly VHT)\nPureConnect / PureEngage (legacy)"),
    ]

    for i, (title, domain, desc, apps) in enumerate(dvs_cards):
        x = margin + i * (card_w + gap)
        add_dvs_card(slide, x, dvs_y, card_w, dvs_h,
                     title, domain, desc, "ALL APPLICATIONS", apps,
                     font_title=11, font_apps=8.5)

    # Enabling team band
    et_y = dvs_y + dvs_h + 0.20
    et_h = 0.85
    et_x = margin
    et_w = total_w
    band = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  Inches(et_x), Inches(et_y), Inches(et_w), Inches(et_h))
    band.fill.solid(); band.fill.fore_color.rgb = ENAB
    band.line.fill.background()
    band.adjustments[0] = 0.20

    # Two-column text inside band
    add_text(slide, et_x + 0.20, et_y + 0.08, 3.20, 0.30,
             "ENABLING TEAM",
             size=11, bold=True, color=WHITE)
    add_text(slide, et_x + 0.20, et_y + 0.36, 3.20, 0.40,
             "Shared services · supports the 4 DVS\nNo Product Manager · governed by Director CS Apps + Sr. Manager SAP Architecture",
             size=8.5, color=WHITE)
    add_text(slide, et_x + 3.50, et_y + 0.10, et_w - 3.70, 0.30,
             "FOUNDATION SERVICES",
             size=9, bold=True, color=WHITE)
    add_text(slide, et_x + 3.50, et_y + 0.32, et_w - 3.70, 0.50,
             "SAP Basis  ·  SAP Gold Client  ·  SAP Enable Now  ·  Blue Prism (RPA)  ·  "
             "Tricentis TOSCA  ·  Jira",
             size=9.5, color=WHITE)


def rebuild_enabling_team_slide(prs):
    """Slide 12 (was DVS 5/5) → Enabling Team detail."""
    slide = prs.slides[11]  # 0-idx
    by_idx = {i: s for i, s in enumerate(slide.shapes)}

    def setrun(i, text):
        sh = by_idx[i]
        tf = sh.text_frame
        if tf.paragraphs and tf.paragraphs[0].runs:
            f0 = tf.paragraphs[0].runs[0].font
            sz = f0.size; bold = f0.bold; italic = f0.italic
            try: color = f0.color.rgb
            except: color = None
            name = f0.name
        else:
            sz, bold, italic, color, name = Pt(10), False, False, None, "Arial"
        tf.clear()
        for j, line in enumerate(text.split('\n')):
            p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
            r = p.add_run(); r.text = line
            r.font.name = name
            if sz: r.font.size = sz
            r.font.bold = bold; r.font.italic = italic
            if color: r.font.color.rgb = color

    # Recolor header bars to neutral grey to signal "not a DVS"
    # Header bar shapes are at idx 0 area? Actually header is via slide bg.
    # The colored chips (idx 4 = owner chip, idx 10 = section header rectangles 0F1B4F)
    # We'll leave colors alone for now — relabel only.

    setrun(0,  "ENABLING TEAM")
    setrun(1,  "FOUNDATION SERVICES — SHARED CAPABILITIES THAT KEEP THE 4 DVS RUNNING")
    setrun(3,  "CS Apps · governed by Director CS Apps + Sr. Manager SAP Architecture (vacancy)")
    setrun(5,  "No PM · no DVS-level backlog · operates as a small shared-service team")
    setrun(8,  "Provide the technical foundation — landscape, data management, automation, testing and "
               "tooling — that the 4 DVS rely on. Not a product. Owned operationally, not as a product line.")
    setrun(11, "FOUNDATION SERVICES")
    setrun(12, "•  SAP Basis (landscape · transports · security · performance)\n"
               "•  SAP Gold Client (selective non-prod data copy · masking)\n"
               "•  SAP Enable Now (training & in-app guidance)\n"
               "•  Blue Prism (RPA — bots that run inside the 4 DVS)\n"
               "•  Tricentis TOSCA (test automation — skill embedded in DVS where possible)\n"
               "•  Jira (demand · backlog · agile work mgmt — admin)")
    setrun(15, "WHO USES THESE SERVICES")
    setrun(16, "•  All 4 DVS — environments, data refresh, transports, regression cycles\n"
               "•  Release & change governance — readiness gates, audit trail\n"
               "•  Security / compliance — vault · access · audit\n"
               "•  Process owners — RPA opportunities (bots embedded in DVS workflows)")
    setrun(17, "Foundation services support all 4 DVS — they do not constitute a separate product line")
    setrun(20, "TEAM SHAPE")
    setrun(21, "•  Tech Lead — Naggi Reddy (Sr. Manager SAP Basis)\n"
               "•  Architecture — Sr. Manager SAP Architecture (vacancy → fill)\n"
               "•  RPA / Test / ALM Engineers — embedded in DVS where feasible, central where not\n"
               "•  No dedicated Product Manager — coordinated by Director CS Apps\n"
               "Squad size · to validate")
    setrun(22, "Why not a DVS: at this headcount, productizing the platform would create overhead without a real internal customer roadmap")
    setrun(25, "VENDOR EXTENSION")
    setrun(26, "•  Tricentis (test platform)\n"
               "•  Blue Prism vendor + Automation CoE\n"
               "•  Atlassian Jira admins\n"
               "•  SAP partner for Basis / Gold Client")
    setrun(27, "Avangrid governs the team operationally · vendors execute specific capabilities")


def update_big_picture(prs):
    """Slide 2 (big picture) — change bottom transversal band from 'Foundation Platform' to 'Enabling Team'."""
    slide = prs.slides[1]  # 0-idx slide 2 (big picture)
    for sh in slide.shapes:
        if sh.has_text_frame:
            t = sh.text_frame.text
            if "Foundation Platform" in t and "productized services" in t:
                # rewrite this textbox
                tf = sh.text_frame
                if tf.paragraphs and tf.paragraphs[0].runs:
                    f0 = tf.paragraphs[0].runs[0].font
                    sz = f0.size; bold = f0.bold
                    try: color = f0.color.rgb
                    except: color = None
                    name = f0.name
                else:
                    sz, bold, color, name = Pt(8.5), True, WHITE, "Arial"
                tf.clear()
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                r = p.add_run()
                r.text = ("Enabling Team  ·  shared services consumed by the 4 DVS  "
                          "(Basis · data mgmt · RPA · test · ALM tooling)")
                r.font.name = name
                if sz: r.font.size = sz
                r.font.bold = bold
                if color: r.font.color.rgb = color
            elif "Coloured strips — the Development Value Streams" in t:
                tf = sh.text_frame
                # We want to update this paragraph specifically, but text_frame can have many paragraphs.
                # Simpler: rewrite the whole shape with corrected wording.
                # capture first run formatting
                runs = [r for p in tf.paragraphs for r in p.runs]
                if runs:
                    f0 = runs[0].font
                    sz = f0.size; color = None
                    try: color = f0.color.rgb
                    except: pass
                    name = f0.name
                else:
                    sz, color, name = Pt(9.5), DARK, "Arial"
                tf.clear()
                lines = [
                    "•  Top row — the customer-facing flow (the operational outcome customers experience)",
                    "•  Green chevrons — the 8 value-adding steps that take the customer from \"need energy\" to \"bill paid\"",
                    "•  Teal boxes — the solutions / applications that enable each step",
                    "•  Coloured strips — the 4 Development Value Streams that own & evolve each solution",
                    "•  Grey band — the small Enabling Team that supports all 4 DVS with shared technical services",
                ]
                for j, line in enumerate(lines):
                    p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
                    p.alignment = PP_ALIGN.LEFT
                    r = p.add_run(); r.text = line
                    r.font.name = name
                    if sz: r.font.size = sz
                    if color: r.font.color.rgb = color


def main():
    prs = Presentation(SRC)

    n = walk_replace(prs)
    print(f"Text swaps: {n} runs updated")

    rebuild_cards_slide(prs)
    print("Cards slide rebuilt — 4 DVS + 1 enabling team")

    rebuild_enabling_team_slide(prs)
    print("Slide 12 recast → Enabling Team")

    update_big_picture(prs)
    print("Big picture updated → Enabling Team band")

    # Update slide 12 title in deck listing — "DEVELOPMENT VALUE STREAM (DVS) 5 / 5"
    # was already overwritten in setrun(0).
    # However the v3 may have some leftover header references; walk_replace will catch them.

    prs.save(DST)
    print(f"\nSaved → {DST}")
    print(f"Total slides: {len(prs.slides)}")
    for i, s in enumerate(prs.slides, 1):
        title = ""
        for sh in s.shapes:
            if sh.has_text_frame and sh.text_frame.text.strip():
                title = sh.text_frame.text.strip().split('\n')[0][:80]
                break
        print(f"  {i:2d}. {title}")


if __name__ == "__main__":
    main()
