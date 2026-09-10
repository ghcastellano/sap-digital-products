"""Reference slides organized by phase — print board for canvas sessions"""
from PIL import Image, ImageDraw, ImageFont
import os

BG=(2,17,30); CARD=(4,22,44); WHITE=(255,255,255); DIM=(160,175,195)
MUTED=(100,120,145); INNER=(8,28,52); BORD=(40,60,85)
CYAN=(80,180,220); PINK=(200,80,180); ORANGE=(247,148,29)
BLUE=(0,100,180); GREEN=(120,190,32); LIME=(160,210,40)

PROJ_C = {
    "PS":    ((120,190,32), "P&S / P&T"),
    "DCoE":  ((0,100,180),  "Digital CoE"),
    "DTO":   ((247,148,29), "Renewables DTO"),
    "SAPDP": ((200,60,60),  "SAP Digital Products (AS-IS)"),
    "Infra_PB": ((160,100,200), "Infrastructure"),
}

def font(size,bold=False):
    for p in ["/System/Library/Fonts/Helvetica.ttc","/System/Library/Fonts/SFNSText.ttf"]:
        if os.path.exists(p):
            try: return ImageFont.truetype(p,size,index=1 if bold and p.endswith(".ttc") else 0)
            except: return ImageFont.truetype(p,size)

FT=font(60,True); FM=font(30); FPHASE=font(52,True); FPSUB=font(26)
FCAP=font(24,True); FPROJ=font(22,True); FTAG=font(20)
FNUM=font(72,True)

SLIDES="/tmp/slide_renders/board_demand_portfolio"

# ═══════════════════════════════════════════════════════════════
# Phase definitions with slides
# ═══════════════════════════════════════════════════════════════
phases = [
    ("01", "INTAKE", "Channels & Sources", CYAN, [
        ("PS",    "PS_S42_Lean_Funnel.png",             "S42 — Lean Portfolio Funnel"),
        ("PS",    "PS_S35_Portfolio_Demand_Funnel.png", "S35 — Portfolio & Demand Funnel"),
        ("DTO",   "DTO_S19_Single_Intake.png",          "S19 — Jira Single Intake + DoR"),
        ("DTO",   "DTO_S6_Findings_Demand_Intake.png",  "S6 — Findings AS-IS Intake"),
        ("DCoE",  "DCoE_S32_Agile_Portfolio_Funnel.png","S32 — Agile Portfolio Funnel"),
        ("SAPDP", "SAPDP_S6_Demand_AS_IS.png",          "S6 — AS-IS Baseline"),
    ]),
    ("02", "CLASSIFICATION", "Type, Size, Urgency", PINK, [
        ("DCoE", "DCoE_S105_Classes_of_Service.png",     "S105 — Classes of Service + SLAs"),
        ("DTO",  "DTO_S20_Work_Types.png",               "S20 — Demand Classification"),
        ("DTO",  "DTO_S35_Triage_Intake_to_Backlog.png", "S35 — Triage → Backlog"),
    ]),
    ("03", "PRIORITIZATION", "Scoring & Criteria", ORANGE, [
        ("PS",    "PS_S72_BV_Matrix.png",               "S72 — BV Scoring Matrix"),
        ("PS",    "PS_S77_PBI_Formula.png",             "S77 — PBI Formula (BV+UX)/E"),
        ("PS",    "PS_S59_Prioritization_Criterias.png","S59 — Criterias v1.0"),
        ("DCoE",  "DCoE_S84_PBI_Formula.png",           "S84 — PBI Formula CX CoE"),
        ("DTO",   "DTO_S21_Core_Prioritization.png",    "S21 — 6 Core Prioritization Processes"),
        ("SAPDP", "SAPDP_S8_Value_Criteria.png",        "S8 — Value Criteria & ROI"),
    ]),
    ("04", "PORTFOLIO", "Governance & Decision", BLUE, [
        ("PS",   "PS_S55_Governance_Forums.png",         "S55 — Governance Forums TO-BE"),
        ("PS",   "PS_S43_Shark_Tank.png",                "S43 — Shark Tank Gate 1"),
        ("PS",   "PS_S45_Portfolio_Committee.png",       "S45 — Portfolio Committee Gate 2"),
        ("PS",   "PS_S53_Governance_Cadences.png",       "S53 — Governance Cadences"),
        ("DCoE", "DCoE_S7_3_Tier_Governance.png",        "S7 — 3-Tier Governance"),
        ("DCoE", "DCoE_S28_Portfolio_Roadmap.png",       "S28 — Portfolio Roadmap"),
        ("DCoE", "DCoE_S29_Portfolio_Initiatives.png",   "S29 — Portfolio Initiatives"),
        ("DCoE", "DCoE_S25_Stage_Gate_Governance.png",   "S25 — Stage-Gate Governance"),
        ("DCoE", "DCoE_S26_Stage_Gate_Policies.png",     "S26 — Stage-Gate Policies"),
        ("DCoE", "DCoE_S65_Feedback_Loops_Cadences.png", "S65 — Feedback Loops"),
        ("DTO",  "DTO_S50_Facilitate_Cadences.png",      "S50 — Facilitate Cadences"),
    ]),
    ("05", "DELIVERY", "Execution & Flow", GREEN, [
        ("PS",       "PS_S37_Compass.png",              "S37 — Compass E2E"),
        ("DCoE",     "DCoE_S10_E2E_Discovery_Upstream.png", "S10 — E2E Discovery Upstream"),
        ("DCoE",     "DCoE_S11_E2E_ROI_Funnel.png",     "S11 — E2E ROI Funnel"),
        ("DTO",      "DTO_S12_Epic_Kanban.png",         "S12 — Epic Kanban Strategy→Execution"),
        ("Infra_PB", "Infra_PB_S12_Budget_Resourcing.png","PB S12 — Budget & Resourcing"),
    ]),
]

# Layout
W = 5600
M = 50
THUMB_W = 860
THUMB_H = 484
CAP_H = 90
COL_GAP = 20
ROW_GAP = 24
MAX_COLS = 6
HDR_H = 140

def rrect(d,x,y,w,h,fill,outline=None,r=14,width=2):
    d.rounded_rectangle([x,y,x+w,y+h],r,fill=fill,outline=outline,width=width if outline else 0)

def rows_for(n): return max(1,(n+MAX_COLS-1)//MAX_COLS)

# Calculate total height
total_h = M + 140  # title
for _,_,_,_,slides in phases:
    total_h += HDR_H + 20
    r = rows_for(len(slides))
    total_h += r * (THUMB_H + CAP_H + ROW_GAP)
total_h += M + 40

img = Image.new("RGB",(W,total_h),BG)
d = ImageDraw.Draw(img)

# Title
Y = M
d.text((M,Y),"Reference Slides by Phase",fill=WHITE,font=FT)
Y += 72
d.text((M,Y),"What can be reused from other projects for each Flow Canvas stage",fill=DIM,font=FM)
Y += 60

# Phases
for num,name,sub,color,slides in phases:
    # Phase header
    rrect(d,M,Y,W-M*2,HDR_H,fill=INNER,outline=color,r=16,width=3)
    # Number badge
    d.text((M+28,Y+16),num,fill=color,font=FNUM)
    # Name + subtitle
    d.text((M+150,Y+24),name,fill=WHITE,font=FPHASE)
    d.text((M+150,Y+90),sub,fill=DIM,font=FPSUB)
    # Count of slides
    count_text = f"{len(slides)} reference slides"
    cw = d.textlength(count_text,font=FPSUB)
    d.text((W-M-cw-28,Y+94),count_text,fill=color,font=FPSUB)

    Y += HDR_H + 20

    # Slides grid
    for idx,(proj,fname,caption) in enumerate(slides):
        col = idx % MAX_COLS
        row = idx // MAX_COLS
        cx = M + col * (THUMB_W + COL_GAP)
        cy = Y + row * (THUMB_H + CAP_H + ROW_GAP)

        proj_color,proj_name = PROJ_C[proj]

        # Slide thumbnail
        path = f"{SLIDES}/{fname}"
        if os.path.exists(path):
            t = Image.open(path).resize((THUMB_W,THUMB_H),Image.LANCZOS)
            rrect(d,cx-3,cy-3,THUMB_W+6,THUMB_H+6,fill=BG,outline=proj_color,r=6,width=3)
            img.paste(t,(cx,cy))
        else:
            rrect(d,cx,cy,THUMB_W,THUMB_H,fill=INNER,outline=BORD,r=6,width=2)
            d.text((cx+20,cy+THUMB_H//2),f"missing: {fname}",fill=MUTED,font=FTAG)

        # Caption bar
        d.rectangle([cx,cy+THUMB_H,cx+THUMB_W,cy+THUMB_H+CAP_H],fill=(6,16,30))
        # Project tag
        tag_w = int(d.textlength(proj_name,font=FPROJ)) + 20
        rrect(d,cx+10,cy+THUMB_H+10,tag_w,30,fill=proj_color,r=6)
        d.text((cx+20,cy+THUMB_H+14),proj_name,fill=WHITE,font=FPROJ)
        # Caption text
        d.text((cx+10,cy+THUMB_H+50),caption,fill=DIM,font=FCAP)

    Y += rows_for(len(slides)) * (THUMB_H + CAP_H + ROW_GAP)

img_out = img.crop((0,0,W,Y+M))
out = "/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Miro Discovery Sessions/SESSION_2_Reference_Slides_By_Phase.png"
os.makedirs(os.path.dirname(out), exist_ok=True)
img_out.save(out,"PNG",optimize=True)
px = img_out.size[0]*img_out.size[1]
print(f"{img_out.size[0]}x{img_out.size[1]} — {px/1e6:.1f}MP → {os.path.basename(out)}")
