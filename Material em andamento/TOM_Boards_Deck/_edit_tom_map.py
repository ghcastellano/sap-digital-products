"""
Edit TOM MAP workbook.pptx to add CS Proposal 2 (4 DVS + 1 Enabling Team)
WHILE KEEPING ITBA teams visible. Focus on slide 18 (TO-BE DEMAND FLOW).
Changes are highlighted with bold red text and yellow stickers.
Output: same folder, new filename.
"""
import shutil
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

SRC = "/Users/gustavohenriquecastellano/Downloads/Base Apresentações Projeto Digital Product/TOM MAP workbook.pptx"
OUT = "/Users/gustavohenriquecastellano/Downloads/Base Apresentações Projeto Digital Product/TOM MAP workbook - SAFe Proposal 2 CS.pptx"

shutil.copy(SRC, OUT)
prs = Presentation(OUT)

RED = RGBColor(0xDE, 0x35, 0x0B)
YELLOW = RGBColor(0xFF, 0xC1, 0x07)
NAVY = RGBColor(0x0E, 0x22, 0x40)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
AMBER = RGBColor(0xE8, 0xA2, 0x4C)
AMBER_LIGHT = RGBColor(0xF4, 0xC9, 0x88)
PURPLE = RGBColor(0x5C, 0x2A, 0x91)
PURPLE_LIGHT = RGBColor(0xD4, 0xC4, 0xE8)

CHANGES_LOG = []


def mark_run_changed(run):
    run.font.bold = True
    run.font.color.rgb = RED


def replace_in_runs(slide, old_str, new_str, slide_num=None):
    found = False
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if old_str in run.text:
                    run.text = run.text.replace(old_str, new_str)
                    mark_run_changed(run)
                    found = True
                    if slide_num is not None:
                        CHANGES_LOG.append((slide_num, f'"{old_str[:40]}" → "{new_str[:40]}"'))
    return found


def annotate_slide(slide, note, x=Inches(0.2), y=Inches(7.0)):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(13), Inches(0.4))
    shp.fill.solid()
    shp.fill.fore_color.rgb = YELLOW
    shp.line.color.rgb = RED
    shp.line.width = Pt(1.5)
    tf = shp.text_frame
    tf.margin_left = Emu(36000)
    tf.margin_top = Emu(18000)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = "🔴 SAFe PROPOSAL 2 EDIT — " + note
    r.font.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = NAVY


def add_team_box(slide, x, y, w, h, label, accent_fill, label_color=WHITE):
    """Add a value-stream team box matching the existing visual style."""
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = WHITE
    shp.line.color.rgb = accent_fill
    shp.line.width = Pt(2)
    # Top stripe
    stripe = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, Inches(0.18))
    stripe.fill.solid()
    stripe.fill.fore_color.rgb = accent_fill
    stripe.line.fill.background()
    # Text
    tb = slide.shapes.add_textbox(x + Inches(0.05), y + Inches(0.02), w - Inches(0.1), Inches(0.16))
    p = tb.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = label
    r.font.bold = True
    r.font.size = Pt(7)
    r.font.color.rgb = label_color
    return shp


# ============ SLIDE 16 — LIFECYCLE MODEL ============
slide16 = prs.slides[15]
replace_in_runs(slide16, "1 PM for each Value streams (3)",
                "1 PM for each Value stream (3) — ITBA", slide_num=16)
replace_in_runs(slide16, "1 PM for each Value streams (5)",
                "1 PM for each DVS (4) + Enabling Team (no PM) — CS Proposal 2", slide_num=16)
replace_in_runs(slide16, "SERVICE LIFECYCLE TEAM KANBAN BOARD",
                "DVS 1 — METER & DEVICE OPERATIONS KANBAN BOARD", slide_num=16)
replace_in_runs(slide16, "REVENUE OPERATIONS TEAM KANBAN BOARD",
                "DVS 2 — BILLING OPERATIONS KANBAN BOARD", slide_num=16)
replace_in_runs(slide16, "CUSTOMER CAR & CONTACT CENTER TEAM KANBAN BOARD",
                "DVS 4 — CUSTOMER CARE & ENGAGEMENT KANBAN BOARD", slide_num=16)
annotate_slide(slide16,
               "CS side restructured to 4 DVS + 1 Enabling Team (Proposal 2). ITBA side kept as-is. Need to add: DVS 3 + Enabling Team boards.")

# ============ SLIDE 18 — TO-BE DEMAND FLOW ============
slide18 = prs.slides[17]
# Update header to reflect both ITBA and CS structures
replace_in_runs(slide18, "ONE JIRA FRONT DOOR ROUTES ELEVEN SOURCE AREAS INTO EIGHT VALUE STREAM TEAMS",
                "ONE JIRA FRONT DOOR ROUTES SOURCE AREAS INTO ITBA STREAMS + 4 CS DVS + 1 ENABLING TEAM",
                slide_num=18)
# CS-side renames (ITBA stays unchanged)
replace_in_runs(slide18, "Service Lifecycle",
                "DVS 1 — Meter & Device Operations", slide_num=18)
replace_in_runs(slide18, "Smart Metering / FICA",
                "Smart Metering · MDM · SAP IS-U Device · Group 8760 EDI", slide_num=18)
replace_in_runs(slide18, "Revenue Operations",
                "DVS 2 — Billing Operations", slide_num=18)
replace_in_runs(slide18, "Billing / Collection",
                "SAP ECC (FI/SD) · Streamserve/OpenText · Address Validation", slide_num=18)
replace_in_runs(slide18, "Customer Care & CCC",
                "DVS 4 — Customer Care & Engagement", slide_num=18)
replace_in_runs(slide18, "CRM / Contact Center",
                "SAP CRM · Genesys · Nuance IVR · Medallia · PureConnect", slide_num=18)
# ITBA streams kept untouched (no changes)

# ADD 2 new boxes: DVS 3 (Revenue & Collections) + Enabling Team
# Position: bottom-right area of the slide, near the Value Stream Team column
# Slide is 13.333" x 7.5"
# Add DVS 3 box
dvs3_y = Inches(5.55)
add_team_box(slide18, Inches(10.0), dvs3_y, Inches(3.2), Inches(0.65), "", AMBER)
# Title text
tb = slide18.shapes.add_textbox(Inches(10.05), dvs3_y + Inches(0.20), Inches(3.1), Inches(0.25))
p = tb.text_frame.paragraphs[0]
r = p.add_run()
r.text = "DVS 3 — Revenue & Collections"
r.font.bold = True
r.font.size = Pt(10)
r.font.color.rgb = AMBER_LIGHT
# Subtitle
tb2 = slide18.shapes.add_textbox(Inches(10.05), dvs3_y + Inches(0.42), Inches(3.1), Inches(0.22))
p = tb2.text_frame.paragraphs[0]
r = p.add_run()
r.text = "SAP FICA · SOVOS Taxware · Payment processing · Dunning"
r.font.size = Pt(7)
r.font.color.rgb = NAVY

# Add Enabling Team box (purple to distinguish)
en_y = Inches(6.25)
add_team_box(slide18, Inches(10.0), en_y, Inches(3.2), Inches(0.65), "", PURPLE)
tb3 = slide18.shapes.add_textbox(Inches(10.05), en_y + Inches(0.20), Inches(3.1), Inches(0.25))
p = tb3.text_frame.paragraphs[0]
r = p.add_run()
r.text = "ENABLING TEAM — Foundation Services"
r.font.bold = True
r.font.size = Pt(10)
r.font.color.rgb = PURPLE_LIGHT
tb4 = slide18.shapes.add_textbox(Inches(10.05), en_y + Inches(0.42), Inches(3.1), Inches(0.22))
p = tb4.text_frame.paragraphs[0]
r = p.add_run()
r.text = "SAP Basis · Gold Client · Blue Prism · TOSCA · Enable Now · Jira (no PM)"
r.font.size = Pt(7)
r.font.color.rgb = NAVY

# "NEW" sticker on DVS 3 + Enabling Team
new_sticker = slide18.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(13.05), Inches(5.55), Inches(0.6), Inches(0.3))
new_sticker.fill.solid()
new_sticker.fill.fore_color.rgb = RED
new_sticker.line.fill.background()
tb_new = slide18.shapes.add_textbox(Inches(13.05), Inches(5.55), Inches(0.6), Inches(0.3))
p = tb_new.text_frame.paragraphs[0]
p.alignment = PP_ALIGN.CENTER
r = p.add_run()
r.text = "NEW"
r.font.bold = True
r.font.size = Pt(8)
r.font.color.rgb = WHITE

# Connector annotation
ann_y = Inches(6.95)
annotate_slide(slide18,
               "ADDED to CS side: DVS 3 (Revenue & Collections) + Enabling Team (purple). ITBA streams kept unchanged. CS now has 4 DVS + 1 Enabling Team.")

# ============ SLIDE 22 — TEAM KANBAN example ============
slide22 = prs.slides[21]
replace_in_runs(slide22, "Field Operations Squad",
                "DVS 4 Customer Care & Engagement Squad (example)", slide_num=22)
replace_in_runs(slide22, "Storm Program v2",
                "Customer Care GenAI v2", slide_num=22)
replace_in_runs(slide22, "DLI inspection — offline reqs",
                "CRM unified record — offline reqs", slide_num=22)
replace_in_runs(slide22, "TLI sizing with vendor",
                "IVR containment uplift sizing (Genesys)", slide_num=22)
replace_in_runs(slide22, "Offline DLI capability",
                "Offline CRM capability", slide_num=22)
replace_in_runs(slide22, "Mobile TLI UI tweak",
                "Medallia survey UI tweak", slide_num=22)
replace_in_runs(slide22, "Decommission FWM",
                "Decommission PureConnect (legacy CC)", slide_num=22)
annotate_slide(slide22, "Example squad: Customer Care & Engagement (DVS 4 · CS) — same kanban template applies to all DVS + ITBA squads")

# ============ SLIDE 23 — VALUE STREAM = SQUAD ============
slide23 = prs.slides[22]
annotate_slide(slide23,
               "CS scope · 4 DVS = 4 permanent Squads + 1 Enabling Team (no PM, governed by Director CS + Sr. Manager SAP Architecture).")

# ============ SLIDE 34 — DIGITAL PRODUCTS PORTFOLIO KANBAN ============
slide34 = prs.slides[33]
replace_in_runs(slide34, "1 PM for each Value streams (3)",
                "1 PM for each Value stream (3) — ITBA", slide_num=34)
replace_in_runs(slide34, "1 PM for each Value streams (5)",
                "1 PM for each DVS (4) + Enabling Team (no PM) — CS Proposal 2", slide_num=34)
replace_in_runs(slide34, "SERVICE LIFECYCLE TEAM KANBAN BOARD",
                "DVS 1 — METER & DEVICE OPERATIONS", slide_num=34)
replace_in_runs(slide34, "REVENUE OPERATIONS TEAM KANBAN BOARD",
                "DVS 2 — BILLING OPERATIONS", slide_num=34)
replace_in_runs(slide34, "CUSTOMER CAR & CONTACT CENTER TEAM KANBAN BOARD",
                "DVS 4 — CUSTOMER CARE & ENGAGEMENT", slide_num=34)
annotate_slide(slide34,
               "CS Proposal 2: 4 DVS + 1 Enabling Team. ITBA streams kept. Need to add: DVS 3 (Revenue & Collections) + Enabling Team boards.")

# ============ SLIDE 35 — ROLES & RESPONSIBILITIES ============
slide35 = prs.slides[34]
replace_in_runs(slide35, "Product Manager × 8",
                "Product Manager × 7 (3 ITBA + 4 CS DVS) + 1 Enabling Team (no PM)", slide_num=35)
replace_in_runs(slide35, "One per Value Stream Team",
                "One per VS / DVS · CS now has 4 DVS + 1 Enabling Team (no PM)", slide_num=35)
replace_in_runs(slide35, "Iram Chowdhury — Smart Metering",
                "Iram Chowdhury — DVS 1 (Meter & Device Operations) candidate PM", slide_num=35)
replace_in_runs(slide35, "Stephanie — SAP / CS",
                "Stephanie — DVS 3 (Revenue & Collections) candidate PM", slide_num=35)
annotate_slide(slide35,
               "CS Proposal 2: 4 PMs (one per DVS) + Enabling Team has no PM (governed by Director CS Apps + Sr. Manager SAP Architecture). ITBA champions kept.")

# ============ ADD: New slide at position 2 — CHANGE SUMMARY ============
new_slide_layout = prs.slide_layouts[6]
new_slide = prs.slides.add_slide(new_slide_layout)
xml_slides = prs.slides._sldIdLst
slides_list = list(xml_slides)
xml_slides.remove(slides_list[-1])
xml_slides.insert(1, slides_list[-1])

cs_w = prs.slide_width
cs_h = prs.slide_height

bg = new_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, cs_w, cs_h)
bg.fill.solid()
bg.fill.fore_color.rgb = RGBColor(0xFF, 0xFB, 0xE5)
bg.line.fill.background()

hdr = new_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, cs_w, Inches(0.7))
hdr.fill.solid()
hdr.fill.fore_color.rgb = RED
hdr.line.fill.background()
title = new_slide.shapes.add_textbox(Inches(0.4), Inches(0.15), Inches(12), Inches(0.4))
p = title.text_frame.paragraphs[0]
r = p.add_run()
r.text = "CHANGE SUMMARY · SAFe PROPOSAL 2 (CS) — ITBA KEPT"
r.font.bold = True
r.font.size = Pt(22)
r.font.color.rgb = WHITE

sub = new_slide.shapes.add_textbox(Inches(0.4), Inches(0.85), Inches(12.5), Inches(0.4))
p = sub.text_frame.paragraphs[0]
r = p.add_run()
r.text = "CS side restructured: 1 OVS (Meter-to-Cash) + 4 DVS + 1 Enabling Team · Owner: Kumar Jaladanki · ITBA streams unchanged"
r.font.size = Pt(13)
r.font.color.rgb = NAVY
r.font.italic = True

body = new_slide.shapes.add_textbox(Inches(0.4), Inches(1.4), Inches(12.5), Inches(5.5))
tf = body.text_frame
tf.word_wrap = True
sections = [
    ("CS NEW STRUCTURE (Proposal 2 from CS - Proposed Value Streams + Value Definition.pptx)", True, 14, NAVY, [
        "1 OPERATIONAL VALUE STREAM (M2C) — METER → BILL → COLLECT → SERVE · Owner: Kumar Jaladanki",
        "DVS 1 — Meter & Device Operations · Smart Metering · MDM · SAP IS-U Device · Group 8760 EDI",
        "DVS 2 — Billing Operations · SAP ECC (FI/SD) · Streamserve/OpenText · Address Validation",
        "DVS 3 — Revenue & Collections · SAP FICA · SOVOS Taxware · Payment processing · Dunning   ⭐ NEW",
        "DVS 4 — Customer Care & Engagement · SAP CRM · Genesys · Nuance IVR · Medallia · PureConnect",
        "ENABLING TEAM — Foundation Services · NO PM (governed by Director CS + Sr. Mgr SAP Architecture)   ⭐ NEW",
        "    Foundation: SAP Basis · Gold Client · Enable Now · Blue Prism (RPA) · Tricentis TOSCA · Jira admin",
    ]),
    ("ITBA SIDE — UNCHANGED (Carlos Garbin · 3 streams)", True, 14, NAVY, [
        "Field Operations · Network Operations & Control · Engineering & Asset Mgmt · Corporate & Regulatory · Iberdrola Corporate Services",
        "Sources kept: P&T (GIS · Global SAP) · Gas & Electric Ops · IA · OT · Procurement · Corporate Areas",
    ]),
    ("WHAT CHANGED IN EACH SLIDE", True, 14, NAVY, [
        "Slide 16 + 34 (LIFECYCLE / DIGITAL PRODUCTS PORTFOLIO) — CS streams renamed to DVS 1/2/4; PM count updated to '4 DVS + Enabling (no PM)' for Kumar; ITBA boards left intact",
        "Slide 18 (TO-BE DEMAND FLOW) — Header updated; CS streams renamed to DVS 1/2/4; ⭐ TWO NEW BOXES added (DVS 3 + Enabling Team) in bottom-right area; ITBA streams kept",
        "Slide 22 (TEAM KANBAN example) — Example squad relabelled as DVS 4 Customer Care & Engagement (CS) example",
        "Slide 23 (VS = SQUAD) — Annotated to clarify 4 DVS + 1 Enabling Team for CS",
        "Slide 35 (ROLES) — PM count: 8 → 7 (3 ITBA + 4 CS DVS) + 1 Enabling Team no-PM",
    ]),
    ("HIGHLIGHT LEGEND", True, 12, NAVY, [
        "Bold red text = changed content in slides 16, 18, 22, 23, 34, 35",
        "Yellow stickers (bottom of each slide) = annotation describing what was changed",
        "Two new shape boxes on Slide 18 (DVS 3 + Enabling Team) marked with red 'NEW' sticker",
    ]),
]
for header, bold, size, color, items in sections:
    p = tf.add_paragraph() if tf.paragraphs[0].runs else tf.paragraphs[0]
    r = p.add_run()
    r.text = header
    r.font.bold = True
    r.font.size = Pt(size)
    r.font.color.rgb = color
    for item in items:
        p2 = tf.add_paragraph()
        r2 = p2.add_run()
        r2.text = "• " + item
        r2.font.size = Pt(11)
        r2.font.color.rgb = NAVY

prs.save(OUT)
print(f"Saved: {OUT}")
print(f"Changes recorded: {len(CHANGES_LOG)}")
