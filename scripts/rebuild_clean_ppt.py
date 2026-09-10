from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pathlib import Path

src = Path('/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/Digital Products - Prior Materials Review .pptx')
out = Path('/Users/gustavohenriquecastellano/Desktop/DP-Prior-Materials-Review-rebuilt-clean.pptx')

GREEN = RGBColor(0x00,0xA4,0x43)
BLUE = RGBColor(0x00,0x20,0x60)
WHITE = RGBColor(0xFF,0xFF,0xFF)
LIGHT = RGBColor(0xF5,0xF5,0xF5)
CHAR = RGBColor(0x2C,0x3E,0x50)
DP = RGBColor(0x00,0x5B,0x96)

src_prs = Presentation(str(src))
new = Presentation()
# set widescreen 16:9
new.slide_width = src_prs.slide_width
new.slide_height = src_prs.slide_height

# helper

def add_title_slide(title, subtitle=''):
    s = new.slides.add_slide(new.slide_layouts[6])
    t = s.shapes.add_textbox(Inches(0.8), Inches(2.0), Inches(11.5), Inches(1.0))
    p=t.text_frame.paragraphs[0]
    r=p.add_run(); r.text=title; r.font.size=Pt(44); r.font.color.rgb=GREEN
    if subtitle:
        p2=t.text_frame.add_paragraph(); r2=p2.add_run(); r2.text=subtitle; r2.font.size=Pt(18); r2.font.color.rgb=CHAR

# slide 1
texts=[]
for sh in src_prs.slides[0].shapes:
    if sh.has_text_frame:
        tx=sh.text_frame.text.strip()
        if tx: texts.append(tx)
add_title_slide(texts[0] if texts else 'Digital Products', texts[1] if len(texts)>1 else 'Review of existing assets')

# slide 2-3 simplified from existing
for i in [1,2]:
    s = new.slides.add_slide(new.slide_layouts[6])
    y=Inches(0.4)
    for sh in src_prs.slides[i].shapes:
        if sh.has_text_frame:
            tx=sh.text_frame.text.strip()
            if tx:
                box=s.shapes.add_textbox(Inches(0.8), y, Inches(11.7), Inches(0.55))
                p=box.text_frame.paragraphs[0]
                r=p.add_run(); r.text=tx[:220]
                r.font.size=Pt(14 if y<Inches(1) else 10)
                r.font.color.rgb = GREEN if y<Inches(1) else CHAR
                y += Inches(0.5)
                if y > Inches(6.8):
                    break

# table slides copied content only
for si in [3,4,5]:
    src_slide = src_prs.slides[si]
    # title
    title='Table'
    for sh in src_slide.shapes:
        if sh.has_text_frame:
            tx=sh.text_frame.text.strip()
            if tx:
                title=tx.split('\n')[0]
                break
    s = new.slides.add_slide(new.slide_layouts[6])
    tb=s.shapes.add_textbox(Inches(0.4), Inches(0.15), Inches(12.5), Inches(0.5))
    p=tb.text_frame.paragraphs[0]; r=p.add_run(); r.text=title; r.font.size=Pt(14); r.font.color.rgb=GREEN

    # find source table
    st=None
    for sh in src_slide.shapes:
        if getattr(sh,'has_table',False) and sh.has_table:
            st=sh.table
            break
    if st is None:
        continue
    rows=len(st.rows); cols=len(st.columns)
    shape=s.shapes.add_table(rows, cols, Inches(0.2), Inches(0.8), Inches(12.9), Inches(6.4))
    t=shape.table

    # widths
    if cols==5:
        widths=[Inches(1.1),Inches(3.0),Inches(2.8),Inches(2.2),Inches(3.8)]
    else:
        widths=[Inches(1.4),Inches(4.2),Inches(3.9),Inches(3.0)]
    for ci,w in enumerate(widths[:cols]):
        t.columns[ci].width=int(w)

    for r in range(rows):
        for c in range(cols):
            cell=t.cell(r,c)
            txt=st.cell(r,c).text
            cell.text=txt
            if r==0:
                cell.fill.solid(); cell.fill.fore_color.rgb = DP if (cols==5 and c==4) else BLUE
                for p in cell.text_frame.paragraphs:
                    for run in p.runs:
                        run.font.size=Pt(7); run.font.bold=True; run.font.color.rgb=WHITE
            else:
                cell.fill.solid(); cell.fill.fore_color.rgb = WHITE if r%2 else LIGHT
                if cols==5 and c==4:
                    cell.fill.fore_color.rgb = RGBColor(0xE8,0xF0,0xFE) if r%2 else RGBColor(0xD6,0xE4,0xF0)
                for p in cell.text_frame.paragraphs:
                    for run in p.runs:
                        run.font.size=Pt(7); run.font.color.rgb=CHAR

new.save(str(out))
print(out)
