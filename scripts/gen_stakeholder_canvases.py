"""Stakeholder Interview Canvases — Vitor and Stephanie"""
from PIL import Image, ImageDraw, ImageFont
import os

BG=(2,17,30); CARD=(4,22,44); WHITE=(255,255,255); DIM=(160,175,195)
MUTED=(100,120,145); GREEN=(120,190,32); BLUE=(0,82,155); ORANGE=(247,148,29)
RED=(220,60,60); PINK=(200,80,180); LIME=(160,210,40); CYAN=(80,180,220)
INNER=(8,28,52); BORD=(40,60,85)

def font(size,bold=False):
    for p in ["/System/Library/Fonts/Helvetica.ttc","/System/Library/Fonts/SFNSText.ttf"]:
        if os.path.exists(p):
            try: return ImageFont.truetype(p,size,index=1 if bold and p.endswith(".ttc") else 0)
            except: return ImageFont.truetype(p,size)

FT=font(60,True); FM=font(32); FS=font(44,True); FSB=font(28)
FL=font(30,True); FB=font(26); FSM=font(22); FQ=font(24)

W=4200; M=40; GAP=32; INN=20; FW=W-M*2
SLIDE_MAX_H = 720  # cap slide height for all blocks

def rrect(d,x,y,w,h,fill=CARD,outline=None,r=18):
    if w<=0 or h<=0: return
    d.rounded_rectangle([x,y,x+w,y+h],r,fill=fill,outline=outline,width=2 if outline else 0)

def hdr(d,x,y,text,color,sub=None):
    d.ellipse([x,y+6,x+26,y+32],fill=color)
    d.text((x+36,y),text,fill=WHITE,font=FS)
    y+=52
    if sub:
        d.text((x+36,y),sub,fill=DIM,font=FSB)
        y+=36
    return y+10

def paste_slide(img,d,path,x,y,w,h,color=BLUE):
    if os.path.exists(path):
        t=Image.open(path).resize((w,h),Image.LANCZOS)
        rrect(d,x-3,y-3,w+6,h+6,fill=BG,outline=color,r=6)
        img.paste(t,(x,y))
    else:
        rrect(d,x,y,w,h,fill=INNER,outline=BORD,r=8)
        d.text((x+20,y+h//2-10),f"[missing]",fill=MUTED,font=FSM)

# ═══════════════════════════════════════════════════════════════════════
# STEPHANIE CANVAS
# ═══════════════════════════════════════════════════════════════════════
def build_stephanie():
    slides="/tmp/slide_renders/stephanie/slides"
    img=Image.new("RGB",(W,12000),BG)
    d=ImageDraw.Draw(img)

    SH_FULL = SLIDE_MAX_H
    SW_FULL = int(SH_FULL * 16/9)

    Y=M
    d.text((M,Y),"Stephanie — Stakeholder Interview",fill=WHITE,font=FT)
    Y+=70
    d.text((M,Y),"1-on-1 Discovery  |  30 min",fill=DIM,font=FM)
    Y+=60

    # Two empty blocks — 15 min each. No placeholders, just open workspace.
    BLOCK_H = 2200

    # ── BLOCK 1: Stephanie shares her knowledge (15 min) ──
    rrect(d,M,Y,FW,BLOCK_H,fill=INNER,outline=BLUE,r=18)
    # Header strip
    rrect(d,M,Y,FW,90,fill=BLUE,r=18)
    d.rectangle([M,Y+70,M+FW,Y+90],fill=BLUE)
    d.ellipse([M+24,Y+32,M+50,Y+58],fill=WHITE)
    d.text((M+66,Y+24),"Walkthrough — Stephanie",fill=WHITE,font=FS)
    # Timer on right
    timer_text="⏱  15 min"
    tw=d.textlength(timer_text,font=FL)
    d.text((M+FW-tw-30,Y+28),timer_text,fill=WHITE,font=FL)
    Y+=BLOCK_H+GAP

    # ── BLOCK 2: Open canvas (15 min) ──
    rrect(d,M,Y,FW,BLOCK_H,fill=INNER,outline=GREEN,r=18)
    rrect(d,M,Y,FW,90,fill=GREEN,r=18)
    d.rectangle([M,Y+70,M+FW,Y+90],fill=GREEN)
    d.ellipse([M+24,Y+32,M+50,Y+58],fill=WHITE)
    d.text((M+66,Y+24),"Open Canvas",fill=WHITE,font=FS)
    timer_text="⏱  15 min"
    tw=d.textlength(timer_text,font=FL)
    d.text((M+FW-tw-30,Y+28),timer_text,fill=WHITE,font=FL)
    Y+=BLOCK_H+M

    return img.crop((0,0,W,Y))


# ═══════════════════════════════════════════════════════════════════════
# VITOR CANVAS
# ═══════════════════════════════════════════════════════════════════════
def build_vitor():
    slides="/tmp/slide_renders/stakeholder/slides"
    img=Image.new("RGB",(W,14000),BG)
    d=ImageDraw.Draw(img)

    Y=M
    d.text((M,Y),"Vitor — Strategic Discovery Canvas",fill=WHITE,font=FT)
    Y+=70
    d.text((M,Y),"Stakeholder Conversation  |  30 min  |  Guiding Drivers × Existing Assets × Projects–Demand Link",
           fill=DIM,font=FM)
    Y+=60

    # ── BLOCK 1: Guiding Drivers (Success Criteria) ──
    b_h=720
    rrect(d,M,Y,FW,b_h,outline=BORD)
    cy=hdr(d,M+INN,Y+20,"Guiding Drivers — Success Criteria",ORANGE,
           "The 3–5 drivers every Digital Products decision must pass through")

    card_h = Y+b_h-cy-40
    bx=M+INN
    cw=FW-INN*2
    rrect(d,bx,cy,cw,card_h,fill=INNER,outline=BORD,r=12)
    d.rectangle([bx,cy,bx+cw,cy+6],fill=ORANGE)
    d.text((bx+18,cy+20),"DRIVERS",fill=ORANGE,font=FL)
    d.text((bx+18,cy+62),"What 3–5 drivers must every decision pass through?",
           fill=MUTED,font=FSM)
    Y+=b_h+GAP

    # ── BLOCK 2: Review of Existing Assets (2x2 grid) ──
    sw = (FW-INN*3)//2
    sh = int(sw * 9/16)
    b_h = 100 + 2*(sh+100) + 20
    rrect(d,M,Y,FW,b_h,outline=BORD)
    cy=hdr(d,M+INN,Y+20,"Review of Existing Assets",BLUE,
           "What the 4 other projects already built — the reuse menu")

    assets=[
        ("ASSET_PS_Flight_Level_Kanban.png","P&S / P&T — Flight Level Kanban Governance",GREEN),
        ("ASSET_DCoE_Dual_Board.png","Digital CoE — Discovery-Delivery Dual Board",BLUE),
        ("ASSET_Infra_Team_Topologies.png","Infrastructure — Team Topologies + Phased Deployment",PINK),
        ("ASSET_DTO_Lean_Kanban_4_Pillar.png","Renewables DTO — Lean Kanban 4-Pillar Model",ORANGE),
    ]
    for i,(fname,label,col) in enumerate(assets):
        r,c=i//2, i%2
        sx=M+INN+c*(sw+INN)
        sy=cy+r*(sh+100)
        rrect(d,sx,sy,sw,44,fill=col,r=8)
        d.text((sx+16,sy+8),label,fill=WHITE,font=FL)
        paste_slide(img,d,f"{slides}/{fname}",sx,sy+54,sw,sh,col)
    Y+=b_h+GAP

    # ── BLOCK 3: Link Projects/Apps ↔ Demands ──
    slide_h = 560
    slide_w = int(slide_h * 16/9)
    REFLECT_H = 280
    row_h = 48 + slide_h + 20
    b_h = 100 + 60 + 3*(row_h+24) + REFLECT_H + 40
    rrect(d,M,Y,FW,b_h,outline=BORD)
    cy=hdr(d,M+INN,Y+20,"Link between Projects / Apps and Demands",ORANGE,
           "Walk through the current state — and reflect on how the link is being built")

    rows_data=[
        {
            "title":"2026 Portfolio — Business & IT Funded Projects",
            "subtitle":"The list of projects in flight for the year",
            "color":LIME,
        },
        {
            "title":"Demand Entry and Typologies",
            "subtitle":"Channels · Classifications · Typologies",
            "color":PINK,
        },
        {
            "title":"Applications Support",
            "subtitle":"Digital Products × Key Applications",
            "color":ORANGE,
        },
    ]

    cy_row = cy
    for row in rows_data:
        # Slide title bar (full width)
        rrect(d,M+INN,cy_row,FW-INN*2,40,fill=row["color"],r=8)
        d.text((M+INN+16,cy_row+6),row["title"],fill=WHITE,font=FL)
        # Subtitle on the right of the title bar
        sub=row["subtitle"]
        sw_=d.textlength(sub,font=FSM)
        d.text((M+FW-INN-sw_-20,cy_row+10),sub,fill=WHITE,font=FSM)
        # Empty slide drop area
        sy_slide = cy_row+48
        rrect(d,M+INN,sy_slide,FW-INN*2,slide_h,fill=INNER,outline=row["color"],r=10)
        cy_row += row_h + 24

    # Reflection block
    rrect(d,M+INN,cy_row,FW-INN*2,REFLECT_H,fill=INNER,outline=ORANGE,r=14)
    d.text((M+INN+24,cy_row+18),"REFLECTION SPACE",fill=ORANGE,font=FL)
    d.text((M+INN+24,cy_row+58),
           "How is the link between Apps, Processes and Projects being built today?",
           fill=WHITE,font=FB)

    Y+=b_h+GAP

    # ── BLOCK 4: Alignment on BRS Champion ──
    b_h=500
    rrect(d,M,Y,FW,b_h,outline=BORD)
    cy=hdr(d,M+INN,Y+20,"Alignment on BRS Champion",BLUE,
           "Define the role, expectations and selection of the BRS Champion")
    rrect(d,M+INN,cy,FW-INN*2,Y+b_h-cy-20,fill=INNER,outline=BORD,r=12)
    Y+=b_h+GAP

    # ── BLOCK 5: Product Centric Operations for Digital Products ──
    b_h=560
    rrect(d,M,Y,FW,b_h,outline=BORD)
    cy=hdr(d,M+INN,Y+20,"Product Centric Operations for Digital Products",GREEN,
           "Vision on which role will work as the PO for SAP teams — "
           "since the demand areas own the E2E product")
    rrect(d,M+INN,cy,FW-INN*2,Y+b_h-cy-20,fill=INNER,outline=BORD,r=12)
    Y+=b_h+M

    return img.crop((0,0,W,Y))


# ═══════════════════════════════════════════════════════════════════════
base="/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Miro Discovery Sessions"
os.makedirs(base,exist_ok=True)

def build_drivers_reference():
    """Reference card — value driver categories for Vitor's guiding drivers block"""
    categories=[
        ("Business Impact",GREEN,[
            "Revenue contribution / growth",
            "Cost reduction / cost avoidance",
            "Margin improvement",
            "Time-to-market reduction",
        ]),
        ("Customer Impact",CYAN,[
            "Customer satisfaction (CSAT / NPS)",
            "Customer retention / churn",
            "Adoption & usage rates",
            "Customer effort (ease of interaction)",
        ]),
        ("User Experience & Usability",PINK,[
            "Ease of use / learnability",
            "Task completion rate",
            "Error rates",
            "Accessibility compliance",
        ]),
        ("Operational Efficiency",ORANGE,[
            "Process cycle time",
            "Throughput / productivity",
            "Automation level",
            "Manual effort reduction",
        ]),
        ("Product Effectiveness",BLUE,[
            "Feature usage vs. availability",
            "Alignment with user needs",
            "Success rate of key journeys",
            "Outcome achievement (did it solve the problem?)",
        ]),
        ("Quality & Reliability",RED,[
            "Defect rates / incidents",
            "System uptime / availability",
            "Performance (latency, response time)",
            "Rework / bug fixing effort",
        ]),
        ("Time & Delivery Performance",LIME,[
            "Lead time / cycle time",
            "On-time delivery rate",
            "Release frequency",
            "Predictability of delivery",
        ]),
        ("Strategic Alignment",(150,110,210),[
            "Alignment with business priorities",
            "Contribution to OKRs",
            "Portfolio fit / prioritization score",
            "Innovation contribution",
        ]),
        ("Financial Efficiency",(240,200,60),[
            "Cost per feature / product",
            "ROI / payback period",
            "Budget adherence",
            "Resource utilization",
        ]),
        ("Scalability & Sustainability",(80,200,180),[
            "Ability to scale (users, volume)",
            "Technical debt level",
            "Maintainability",
            "Vendor dependency / flexibility",
        ]),
    ]

    COLS = 2
    ROWS = 5
    CARD_W = (FW - (COLS-1)*30) // COLS
    CARD_H = 520
    card_gap = 30

    total_h = M + 160 + ROWS*CARD_H + (ROWS-1)*card_gap + M
    img=Image.new("RGB",(W,total_h),BG)
    d=ImageDraw.Draw(img)

    Y=M
    d.text((M,Y),"Driver Examples — Reference",fill=WHITE,font=FT)
    Y+=70
    d.text((M,Y),"Use these as conversation starters when exploring success drivers with Vitor",
           fill=DIM,font=FM)
    Y+=60

    for i,(title,color,items) in enumerate(categories):
        r = i // COLS
        c = i % COLS
        bx = M + c*(CARD_W+30)
        by = Y + r*(CARD_H+card_gap)

        rrect(d,bx,by,CARD_W,CARD_H,fill=INNER,outline=color,r=16)
        # Top accent bar
        d.rectangle([bx,by,bx+CARD_W,by+8],fill=color)
        # Number badge
        rrect(d,bx+20,by+26,56,56,fill=color,r=12)
        num=f"{i+1:02d}"
        tw=d.textlength(num,font=FL)
        d.text((bx+20+(56-tw)//2,by+36),num,fill=WHITE,font=FL)
        # Title
        d.text((bx+94,by+32),title,fill=WHITE,font=FS)

        # Items
        iy = by + 130
        for item in items:
            # Bullet
            d.ellipse([bx+30,iy+10,bx+44,iy+24],fill=color)
            # Item text (wrap if needed)
            lines=[]
            words=item.split()
            line=""
            max_w=CARD_W-90
            for w in words:
                t=f"{line} {w}".strip()
                if d.textlength(t,font=FB)<=max_w:
                    line=t
                else:
                    lines.append(line)
                    line=w
            if line: lines.append(line)
            for li,ln in enumerate(lines):
                d.text((bx+54,iy+li*36),ln,fill=DIM,font=FB)
            iy += max(60, len(lines)*36+20)

    return img.crop((0,0,W,Y+ROWS*CARD_H+(ROWS-1)*card_gap+M))


def build_carlos():
    img=Image.new("RGB",(W,9000),BG)
    d=ImageDraw.Draw(img)

    Y=M
    d.text((M,Y),"Carlos — Strategic Discovery Canvas",fill=WHITE,font=FT)
    Y+=70
    d.text((M,Y),"Stakeholder Conversation  |  30 min",fill=DIM,font=FM)
    Y+=60

    questions = [
        ("01",
         "Expectations & Success",
         "What are your expectations for this project? What would make it a clear success for your area?",
         BLUE, 1200),
        ("02",
         "Where to Focus",
         "Where should we put our energy first? Which areas or teams need the most support right now?",
         GREEN, 1200),
        ("03",
         "Critical Processes",
         "Which end-to-end processes are most critical to improve or redesign? What is broken today?",
         ORANGE, 1200),
        ("04",
         "Demand & Projects Distribution",
         "How are projects and demands distributed across Corporate, GIS, G&E and other areas? Where are the gaps in visibility?",
         PINK, 1200),
        ("05",
         "P&T Priorities",
         "Within P&T specifically — what is the priority order? What are you looking for from the transformation?",
         RED, 1200),
    ]

    for num, title, subtitle, color, body_h in questions:
        b_h = 180 + body_h + 20
        rrect(d,M,Y,FW,b_h,outline=BORD)

        # Header
        rrect(d,M,Y,FW,130,fill=INNER,outline=color,r=18)
        d.rectangle([M,Y,M+10,Y+130],fill=color)

        # Number badge
        FQNUM = font(54,True)
        rrect(d,M+28,Y+28,72,72,fill=color,r=14)
        nw=d.textlength(num,font=FQNUM)
        d.text((M+28+(72-nw)//2,Y+28),num,fill=WHITE,font=FQNUM)

        # Title + subtitle
        d.text((M+120,Y+26),title,fill=WHITE,font=FS)
        d.text((M+120,Y+78),subtitle,fill=DIM,font=FSB)

        # Large empty workspace
        cy = Y + 150
        rrect(d,M+INN,cy,FW-INN*2,body_h,fill=INNER,outline=color,r=14)
        d.rectangle([M+INN,cy,M+INN+8,cy+body_h],fill=color)

        Y += b_h + GAP

    return img.crop((0,0,W,Y))


def build_kumar():
    img=Image.new("RGB",(W,5200),BG)
    d=ImageDraw.Draw(img)

    Y=M
    d.text((M,Y),"Kumar — Stakeholder Interview",fill=WHITE,font=FT)
    Y+=70
    d.text((M,Y),"1-on-1 Discovery  |  30 min",fill=DIM,font=FM)
    Y+=60

    BLOCK_H = 2200

    rrect(d,M,Y,FW,BLOCK_H,fill=INNER,outline=BLUE,r=18)
    rrect(d,M,Y,FW,90,fill=BLUE,r=18)
    d.rectangle([M,Y+70,M+FW,Y+90],fill=BLUE)
    d.ellipse([M+24,Y+32,M+50,Y+58],fill=WHITE)
    d.text((M+66,Y+24),"Walkthrough — Kumar",fill=WHITE,font=FS)
    timer_text="15 min"
    tw=d.textlength(timer_text,font=FL)
    d.text((M+FW-tw-30,Y+28),timer_text,fill=WHITE,font=FL)
    Y+=BLOCK_H+GAP

    rrect(d,M,Y,FW,BLOCK_H,fill=INNER,outline=GREEN,r=18)
    rrect(d,M,Y,FW,90,fill=GREEN,r=18)
    d.rectangle([M,Y+70,M+FW,Y+90],fill=GREEN)
    d.ellipse([M+24,Y+32,M+50,Y+58],fill=WHITE)
    d.text((M+66,Y+24),"Open Canvas",fill=WHITE,font=FS)
    timer_text="15 min"
    tw=d.textlength(timer_text,font=FL)
    d.text((M+FW-tw-30,Y+28),timer_text,fill=WHITE,font=FL)
    Y+=BLOCK_H+M

    return img.crop((0,0,W,Y))


def build_applications_team():
    """Applications Team Discovery Dynamic — 6 timeboxed blocks, ~90 min.
    Same visual layout as Carlos canvas: header (130px, color stripe, number badge, title, subtitle)
    + large workspace. Each workspace lists INDIVIDUAL questions (same format) and SYNERGY questions
    (cross-topic) as guides — participants write on top.
    """
    img=Image.new("RGB",(W,12000),BG)
    d=ImageDraw.Draw(img)

    Y=M
    d.text((M,Y),"Applications Team — Discovery Dynamic Canvas",fill=WHITE,font=FT)
    Y+=70
    d.text((M,Y),"Workshop with Carlos Garbin's branch  |  ~90 min  |  Capacity & Resource Model  ×  Reporting & Metrics Review",
           fill=DIM,font=FM)
    Y+=60

    blocks = [
        ("01", "Capacity Landscape",
         "Dynamic: Mural Mapping  —  place post-its on a 3×3 grid (BAU vs. Strategic × Internal/SWF/Vendor)",
         "15 min", GREEN,
         [
            "How many FTEs does your sub-team have today (internal · SWF · Accenture · vendor)?",
            "What % of your team is on BAU / incidents vs. strategic work?",
            "What is your average utilization % over the last 3 months?",
            "How many open roles, and what is the average time-to-fill?",
            "How long is knowledge transfer when someone leaves (onboarding + ramp-up)?",
            "Who are the key-person risks (concentrated knowledge)?",
         ],
         [
            "Bottlenecks that are NOT people — licenses, access, vendor availability?",
            "When regulatory demand arrives, what gets reprioritized and who decides?",
            "Is there idle capacity today that the current reporting does not show?",
            "If you doubled the team, which lifecycle stage would improve first?",
         ]),
        ("02", "Resource Mix & Cost",
         "Dynamic: Dot Voting + Cost Ranking  —  3 dots on the top mix decisions that blocked or accelerated delivery",
         "15 min", ORANGE,
         [
            "Cost distribution % across internal · SWF · Accenture · Vendor in your sub-team?",
            "Who decides when to use SWF vs. hire internal vs. vendor?",
            "How are effort estimates made — who estimates, on what basis?",
            "Do you have burn-rate visibility (monthly · per project)?",
            "What is the average cost per ticket · per PDD · per release?",
            "Which mix decisions would you undo in the last 12 months?",
         ],
         [
            "Does the mix decision depend on criticality, tech, duration, or regulatory?",
            "How do you prove ROI on a capacity expansion to Carlos / Vitor today?",
            "Does the vendor control the timeline — and what does the reporting show?",
            "If budget were cut 20% tomorrow, what would go first and why?",
         ]),
        ("03", "Reporting Inventory",
         "Dynamic: Kanban Audit  —  each report becomes a card in 3 columns: USED · ORNAMENTAL · MISSING",
         "15 min", BLUE,
         [
            "Which reports / dashboards do you generate or receive today?",
            "Who consumes each (Carlos · Vitor · business · auditor) and at what cadence?",
            "Where does the data live — Jira · IT Now · PDD workbook · Excel · SharePoint?",
            "How much time (h/month) do you spend consolidating reports manually?",
            "Do you trust the numbers you report? If not, why?",
            "Which report does your team produce but NO ONE reads?",
         ],
         [
            "Which numbers in today's reporting reflect the real capacity from Block 1?",
            "Is there any report tracking vendor performance (Block 2)? How granular?",
            "If the TOM brought SSO / automated dashboards, which report would disappear first?",
            "Which data do you REPORT but don't USE — and which do you USE but don't report?",
         ]),
        ("04", "Metrics That Matter",
         "Dynamic: MoSCoW Prioritization  —  classify candidate metrics (MUST · SHOULD · COULD · WON'T); land 3–5 MUSTs",
         "15 min", PINK,
         [
            "Which 3 KPIs are most important for your sub-team today?",
            "Which KPIs do you actually report to Carlos · Vitor?",
            "Which metric shows the health of your team (velocity · backlog · tickets)?",
            "Do you measure Lead Time · Cycle Time · Throughput · WIP · rework? Which?",
            "How do you define \"a successful quarter\" for your sub-team?",
            "Which current metric do you know does NOT reflect reality?",
         ],
         [
            "If you could only report 1 metric to the board, which one — and why?",
            "Which metric would link Capacity (Block 1) with Reporting (Block 3) today?",
            "Which metric do you wish existed but instrumentation prevents?",
            "Metrics the team USES ≠ metrics leadership ASKS for — where is the gap?",
         ]),
        ("05", "Synergy Scenarios",
         "Dynamic: Scenario Walk-through  —  group narrates step-by-step: what they'd do, what data they need, what they wouldn't have",
         "15 min", RED,
         [
            "S1 — Regulatory spike hits today: how long until you see the capacity impact? Which report alerts you?",
            "S1 — Who gets reprioritized — and does the reporting tell Carlos about that trade-off?",
            "S2 — Vanessa asks for 2 more FTEs; Vitor wants justification in 48h: what capacity data do you have?",
            "S2 — What data is MISSING today and would have to be collected manually now?",
            "S3 — Carlos asks cost per sub-team this quarter: can you answer — in how long, with what precision?",
            "S3 — Which of the 4 costs (int / SWF / Accenture / Vendor) are opaque?",
         ],
         [
            "For each scenario: what would the TOM enable that today requires manual work?",
            "Which scenario exposes the largest gap between ideal response and today's reality?",
            "Map each gap → the block that generated it (Capacity · Resource Mix · Reporting · Metrics).",
            "What is the smallest reporting change that resolves the most scenarios at once?",
         ]),
        ("06", "Outputs & Commitments",
         "Dynamic: Artifact Review  —  walk every wall-output, name an owner, set a follow-up date",
         "10 min", CYAN,
         [
            "Block 1 artifact — Capacity map (3×3 by sub-team).  Owner? Consolidation date?",
            "Block 2 artifact — Ranked list of top-5 mix decisions + cost gaps.  Owner?",
            "Block 3 artifact — Kanban report inventory (Used / Ornamental / Missing).  Owner?",
            "Block 4 artifact — MoSCoW grid with 3–5 MUST metrics per sub-team.  Owner?",
            "Block 5 artifact — Scenario gap list → TOM instrumentation backlog.  Owner?",
            "Every HYPOTHESIS tagged item — who verifies it, and by when?",
         ],
         [
            "Every artifact has a named owner and a consolidation date (max 5 business days).",
            "Open hypotheses get a data source assigned + verification deadline.",
            "Next session scheduled on the spot — second-pass to close unanswered questions.",
            "Debrief reads artifacts, not transcript — confirm all are photographed before closing.",
         ]),
    ]

    body_h = 1400
    for num, title, subtitle, duration, color, individual, synergy in blocks:
        b_h = 150 + body_h + 20
        rrect(d,M,Y,FW,b_h,outline=BORD)

        # Header (same style as Carlos)
        rrect(d,M,Y,FW,130,fill=INNER,outline=color,r=18)
        d.rectangle([M,Y,M+10,Y+130],fill=color)

        # Number badge
        FQNUM = font(54,True)
        rrect(d,M+28,Y+28,72,72,fill=color,r=14)
        nw=d.textlength(num,font=FQNUM)
        d.text((M+28+(72-nw)//2,Y+28),num,fill=WHITE,font=FQNUM)

        # Title + subtitle
        d.text((M+120,Y+22),title,fill=WHITE,font=FS)
        d.text((M+120,Y+78),subtitle,fill=DIM,font=FSB)

        # Timer on right
        timer_text=f"⏱  {duration}"
        tw=d.textlength(timer_text,font=FL)
        d.text((M+FW-tw-30,Y+46),timer_text,fill=WHITE,font=FL)

        # Workspace (same style as Carlos) — split into 2 columns for Individual + Synergy
        cy = Y + 150
        rrect(d,M+INN,cy,FW-INN*2,body_h,fill=INNER,outline=color,r=14)
        d.rectangle([M+INN,cy,M+INN+8,cy+body_h],fill=color)

        col_w = (FW - INN*2 - 30) // 2
        # Left column: Individual
        lx = M + INN + 30
        d.text((lx,cy+24),"INDIVIDUAL QUESTIONS  —  same format, run as batch",fill=color,font=FL)
        # Divider
        d.rectangle([lx + col_w - 15, cy+20, lx + col_w - 12, cy + body_h - 20],fill=BORD)
        qy = cy + 80
        for i,q in enumerate(individual):
            rrect(d,lx,qy,34,34,fill=color,r=8)
            nb=f"{i+1}"
            nw=d.textlength(nb,font=FSB)
            d.text((lx+(34-nw)//2,qy+2),nb,fill=WHITE,font=FSB)
            # Wrap
            words=q.split()
            line=""; ly=qy
            max_w = col_w - 70
            for w_ in words:
                test = line + (" " if line else "") + w_
                if d.textlength(test,font=FB) > max_w:
                    d.text((lx+50,ly),line,fill=WHITE,font=FB)
                    ly += 32
                    line = w_
                else:
                    line = test
            if line:
                d.text((lx+50,ly),line,fill=WHITE,font=FB)
                ly += 32
            qy = max(ly, qy + 60) + 30

        # Right column: Synergy
        rx = M + INN + 30 + col_w
        d.text((rx,cy+24),"SYNERGY QUESTIONS  —  cross-topic connections",fill=ORANGE if color != ORANGE else RED,font=FL)
        qy = cy + 80
        syn_color = ORANGE if color != ORANGE else RED
        for q in synergy:
            rrect(d,rx,qy,34,34,fill=syn_color,r=8)
            d.text((rx+8,qy+2),"→",fill=WHITE,font=FSB)
            words=q.split()
            line=""; ly=qy
            max_w = col_w - 70
            for w_ in words:
                test = line + (" " if line else "") + w_
                if d.textlength(test,font=FB) > max_w:
                    d.text((rx+50,ly),line,fill=WHITE,font=FB)
                    ly += 32
                    line = w_
                else:
                    line = test
            if line:
                d.text((rx+50,ly),line,fill=WHITE,font=FB)
                ly += 32
            qy = max(ly, qy + 60) + 30

        Y += b_h + GAP

    return img.crop((0,0,W,Y))



def build_heather():
    """Heather — 30 min stakeholder interview canvas. 4 cols × 2 rows grid."""
    n_cols = 4
    n_rows = 2
    gap_x = 28
    col_w = (FW - gap_x * (n_cols - 1)) // n_cols
    header_h = 140
    body_h = 1200
    block_h = header_h + body_h + 10

    total_h = M + 240 + n_rows * block_h + (n_rows - 1) * GAP + M
    img = Image.new("RGB", (W, total_h), BG)
    d = ImageDraw.Draw(img)

    Y = M
    d.text((M, Y), "Heather — Stakeholder Discovery Canvas", fill=WHITE, font=FT)
    Y += 70
    d.text((M, Y), "1-on-1 Stakeholder Conversation  |  30 min  |  Digital CoE \u00b7 BRS \u00b7 Travis \u00b7 Pilot \u00b7 Prioritization",
           fill=DIM, font=FM)
    Y += 90

    questions = [
        ("01", "Role & Day-to-Day",
         "How do you work in your new role today? Typical day/week — priorities, decisions, escalations?",
         BLUE),
        ("02", "Digital CoE Gaps",
         "What is NOT happening with the Digital CoE that we should strengthen? What should be connected but isn't?",
         GREEN),
        ("03", "BRS Engagement",
         "Walk us through BRS work — meeting cadence, dashboards/panels, conceptual models (the ones Vitor validated).",
         ORANGE),
        ("04", "Working with Travis",
         "How do we sharpen focus to work better with Travis? That's where prioritization lives — what do we need to get right?",
         PINK),
        ("05", "Pilot Scoping",
         "WHERE should we run the pilot? Which team / sub-scope / project makes the best first pilot — and why there?",
         RED),
        ("06", "Prioritization Today",
         "How is prioritization actually done today? Which criteria, which forums, who decides, which gaps?",
         CYAN),
        ("07", "Fixed Recurrence",
         "Proposed fixed recurrence with Heather — cadence, format, artifacts, standing agenda.",
         LIME),
    ]

    FQNUM = font(54, True)

    for idx, (num, title, subtitle, color) in enumerate(questions):
        row = idx // n_cols
        col = idx % n_cols
        x = M + col * (col_w + gap_x)
        y = Y + row * (block_h + GAP)

        # Outer frame
        rrect(d, x, y, col_w, block_h, outline=BORD)

        # Header band
        rrect(d, x, y, col_w, header_h, fill=INNER, outline=color, r=18)
        d.rectangle([x, y, x+10, y+header_h], fill=color)

        # Number badge
        rrect(d, x+22, y+22, 62, 62, fill=color, r=14)
        nw = d.textlength(num, font=FQNUM)
        d.text((x+22+(62-nw)//2, y+26), num, fill=WHITE, font=FQNUM)

        # Title
        d.text((x+102, y+22), title, fill=WHITE, font=FS)

        # Subtitle (wrapped)
        max_w = col_w - 120
        line = ""
        line_y = y + 76
        for w_ in subtitle.split():
            test = line + (" " if line else "") + w_
            if d.textlength(test, font=FSB) > max_w:
                d.text((x+102, line_y), line, fill=DIM, font=FSB)
                line_y += 32
                line = w_
            else:
                line = test
        if line:
            d.text((x+102, line_y), line, fill=DIM, font=FSB)

        # Large empty workspace for stickies / notes
        cy = y + header_h + 4
        rrect(d, x + INN, cy, col_w - INN*2, body_h - 14, fill=INNER, outline=color, r=14)
        d.rectangle([x + INN, cy, x + INN + 8, cy + body_h - 14], fill=color)

    return img.crop((0, 0, W, Y + n_rows * block_h + (n_rows - 1) * GAP + M))


import math as _math
TARGET_PX = 28_000_000   # Miro limit ~30MP; keep 2MP headroom

for name,builder in [("Stephanie_IT_CS",build_stephanie),
                     ("Vitor_Strategic",build_vitor),
                     ("Vitor_Drivers_Reference",build_drivers_reference),
                     ("Carlos_Strategic",build_carlos),
                     ("Kumar_Strategic",build_kumar),
                     ("Heather_Strategic",build_heather),
                     ("Applications_Team_Dynamic",build_applications_team)]:
    result=builder()
    # Downscale if above Miro ceiling
    if result.size[0] * result.size[1] > TARGET_PX:
        scale = _math.sqrt(TARGET_PX / (result.size[0] * result.size[1]))
        new_size = (int(result.size[0] * scale), int(result.size[1] * scale))
        result = result.resize(new_size, Image.LANCZOS)
    out=f"{base}/SESSION_1_{name}_Canvas.png"
    result.save(out,"PNG",optimize=True)
    px=result.size[0]*result.size[1]
    print(f"[{name}] {result.size[0]}x{result.size[1]} — {px/1e6:.1f}MP → {os.path.basename(out)}")
