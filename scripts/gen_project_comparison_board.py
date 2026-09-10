"""Cross-Project Comparison Board — KB-verified slides, split into 2 parts"""
from PIL import Image, ImageDraw, ImageFont
import os, glob

def font(size, bold=False):
    for p in ["/System/Library/Fonts/Helvetica.ttc", "/System/Library/Fonts/SFNSText.ttf"]:
        if os.path.exists(p):
            try: return ImageFont.truetype(p, size, index=1 if bold and p.endswith(".ttc") else 0)
            except: return ImageFont.truetype(p, size)
    return ImageFont.load_default()

FT=font(64,True); FM=font(32); FL=font(30,True)
FB=font(26); FTAG=font(24,True); FPROJ=font(32,True); FDIM=font(30,True)
FDIM2=font(24); FGAP=font(28,True); FNUM=font(56,True)

BG=(2,17,30); CARD=(4,22,44); WHITE=(255,255,255); DIM_C=(160,175,195)
MUTED=(100,120,145); BORD=(40,60,85); TBD_C=(200,150,40)
DIM_ACCENT=(20,40,60)

PROJ_COLORS = {
    "P&S_P&T":(120,190,32), "Digital_CoE":(0,82,155),
    "Renewables_DTO":(247,148,29), "Infrastructure":(160,100,180),
    "Digital_Products":(200,60,60),
}

SLIDE_DIR = "/tmp/slide_renders/slides_final"

dimensions = [
    ("Lifecycle","Operating Model E2E","lifecycle"),
    ("Roles","Team Structure","roles"),
    ("Demand Intake","Channels & Classification","demand"),
    ("Portfolio","Framework & Kanban","portfolio"),
    ("Prioritization","Portfolio Level","prio_portfolio"),
    ("Prioritization","Team / Backlog","prio_team"),
    ("Governance","Forums & Cadences","governance"),
    ("Metrics / KPIs","Dashboards","metrics"),
    ("Capacity","Resource Model","capacity"),
]

M=40; SIDEBAR_W=250; COL_GAP=10; ROW_GAP=14
THUMB_W=760; THUMB_H=428
PROJ_HDR_H=54; CAPTION_H=34

def find_slides(prefix, dim_key):
    """Find all slides for a project+dimension, supporting any number of suffixes."""
    matches = []
    for suffix in ["", "2", "3", "4", "5", "6"]:
        matches.extend(glob.glob(f"{SLIDE_DIR}/{prefix}_{dim_key}{suffix}_*"))
    matches.sort()
    return matches

def build_board(part_projects, part_label, out_path):
    N_COLS = len(part_projects)
    N_ROWS = len(dimensions)

    # Calculate max slides per row for this part
    max_per_row = []
    for _,_,dk in dimensions:
        mx = 0
        for _,prefix in part_projects:
            if prefix == "Digital_Products": continue
            n = len(find_slides(prefix, dk))
            mx = max(mx, n)
        max_per_row.append(max(mx, 1))  # at least 1 for TBD column

    def cell_h(ri):
        n = max_per_row[ri]
        return n * (THUMB_H + CAPTION_H) + max(0, (n-1)*6)

    # Check which rows have ANY content (skip empty rows)
    active_rows = []
    for ri, (_,_,dk) in enumerate(dimensions):
        has_content = False
        for _,prefix in part_projects:
            if prefix == "Digital_Products":
                has_content = True  # always show TBD
                break
            if find_slides(prefix, dk):
                has_content = True
                break
        if has_content:
            active_rows.append(ri)

    GRID_X = M + SIDEBAR_W + 12
    TOTAL_W = GRID_X + N_COLS * (THUMB_W + COL_GAP) - COL_GAP + M
    total_h = M + 120 + PROJ_HDR_H + 8
    for ri in active_rows:
        total_h += cell_h(ri) + ROW_GAP
    total_h += M + 50

    img = Image.new("RGB", (TOTAL_W, total_h), BG)
    d = ImageDraw.Draw(img)

    def rrect(x,y,w,h,fill=CARD,outline=None,r=12):
        d.rounded_rectangle([x,y,x+w,y+h],r,fill=fill,outline=outline,width=2 if outline else 0)

    # Title
    Y = M
    d.text((M,Y), f"Operating Model Comparison — {part_label}", fill=WHITE, font=FT)
    Y += 72
    d.text((M,Y), "KB-verified slides from delivered materials  |  Zoom-friendly resolution", fill=DIM_C, font=FM)
    Y += 42

    # Column headers
    rrect(M,Y,SIDEBAR_W,PROJ_HDR_H-4,fill=DIM_ACCENT,outline=BORD,r=10)
    d.text((M+14,Y+12),"DIMENSION",fill=DIM_C,font=FL)

    for ci,(name,prefix) in enumerate(part_projects):
        cx = GRID_X + ci*(THUMB_W+COL_GAP)
        color = PROJ_COLORS.get(prefix,(100,100,100))
        rrect(cx,Y,THUMB_W,PROJ_HDR_H-4,fill=color,r=10)
        tw = d.textlength(name, font=FPROJ)
        d.text((cx+(THUMB_W-tw)//2, Y+12), name, fill=WHITE, font=FPROJ)
    Y += PROJ_HDR_H + 6

    # Rows (only active ones)
    dim_num = 0
    for ri in active_rows:
        dim_main, dim_sub, dim_key = dimensions[ri]
        ch = cell_h(ri)
        row_y = Y
        dim_num += 1

        # Sidebar
        rrect(M,row_y,SIDEBAR_W,ch,fill=DIM_ACCENT,outline=BORD,r=10)
        d.text((M+14,row_y+10), f"{dim_num:02d}", fill=(50,90,130), font=FNUM)
        d.text((M+14,row_y+70), dim_main, fill=WHITE, font=FDIM)
        d.text((M+14,row_y+104), dim_sub, fill=DIM_C, font=FDIM2)

        for ci,(proj_name,prefix) in enumerate(part_projects):
            cx = GRID_X + ci*(THUMB_W+COL_GAP)
            color = PROJ_COLORS.get(prefix,(100,100,100))

            # Digital Products = TO BE BUILT
            if prefix == "Digital_Products":
                rrect(cx,row_y,THUMB_W,ch,fill=(15,25,40),outline=TBD_C,r=6)
                t1="TO BE BUILT"; tw1=d.textlength(t1,font=FGAP)
                d.text((cx+(THUMB_W-tw1)//2, row_y+ch//2-22), t1, fill=TBD_C, font=FGAP)
                t2="Wave 1 deliverable"; tw2=d.textlength(t2,font=FB)
                d.text((cx+(THUMB_W-tw2)//2, row_y+ch//2+12), t2, fill=MUTED, font=FB)
                continue

            matches = find_slides(prefix, dim_key)

            if matches:
                sy = row_y
                for match_path in matches:
                    slide_img = Image.open(match_path)
                    slide_img = slide_img.resize((THUMB_W, THUMB_H), Image.LANCZOS)
                    rrect(cx-1,sy-1,THUMB_W+2,THUMB_H+2,fill=BG,outline=(*color,80),r=3)
                    img.paste(slide_img, (cx, sy))

                    # Caption
                    fname = os.path.basename(match_path)
                    parts = fname.replace(".png","").split("_")
                    slide_num = "?"
                    for p in parts:
                        if p.isdigit(): slide_num=p; break
                    found=False; lw=[]
                    for p in parts:
                        if found: lw.append(p)
                        if p==slide_num: found=True
                    sl = " ".join(lw).replace("-"," ").replace("plus","+").replace("and","&")
                    cap = f"Slide {slide_num}  —  {sl}" if sl else f"Slide {slide_num}"

                    d.rectangle([cx,sy+THUMB_H,cx+THUMB_W,sy+THUMB_H+CAPTION_H],fill=(8,20,35))
                    d.text((cx+8,sy+THUMB_H+6), cap, fill=DIM_C, font=FTAG)
                    sy += THUMB_H + CAPTION_H + 6
            else:
                # Empty — no content for this project in this dimension
                # Just leave blank (dark background)
                pass

        Y += ch + ROW_GAP

    # Legend
    Y += 6
    rrect(M,Y,TOTAL_W-M*2,44,fill=CARD,outline=BORD,r=10)
    lx = M+16
    for label,col in [("KB-verified slide",(100,180,100)),("TO BE BUILT",TBD_C)]:
        d.ellipse([lx,Y+14,lx+14,Y+28],fill=col)
        d.text((lx+20,Y+10),label,fill=DIM_C,font=FB)
        lx += d.textlength(label,font=FB) + 40
    Y += 44 + M

    img = img.crop((0,0,TOTAL_W,Y))
    img.save(out_path, "PNG", optimize=True)
    px = img.size[0]*img.size[1]
    print(f"  {part_label}: {img.size[0]}x{img.size[1]} — {px/1e6:.1f}MP → {os.path.basename(out_path)}")

# ═══════════════════════════════════════════════════════════════════════
base = "/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento"

# Part 1: The 3 most mature projects
build_board(
    [("P&S / P&T","P&S_P&T"), ("Digital CoE","Digital_CoE"), ("Renewables DTO","Renewables_DTO")],
    "Part 1 — P&S, Digital CoE, Renewables DTO",
    f"{base}/Cross_Project_Comparison_Part1.png"
)

# Part 2: Infrastructure + Digital Products (our scope)
build_board(
    [("Infrastructure","Infrastructure"), ("Digital Products","Digital_Products")],
    "Part 2 — Infrastructure & Digital Products",
    f"{base}/Cross_Project_Comparison_Part2.png"
)
