"""
Fix slide 47:
  - Remove the Gustavo speech bubble (shape [37])
  - Re-add the "Eight principles..." subtitle that got accidentally deleted
"""
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

SRC = '/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/2604_AS-IS deliverable.pptx'
prs = Presentation(SRC)
slide = prs.slides[46]

# 1. Remove the Gustavo comment bubble — identify by text
to_remove = []
for sh in slide.shapes:
    if sh.has_text_frame and 'vitor told us about' in sh.text_frame.text.lower():
        to_remove.append(sh)
for sh in to_remove:
    sh._element.getparent().remove(sh._element)

# 2. Re-add the subtitle "Eight principles..."
tb = slide.shapes.add_textbox(Emu(246506), Emu(843976), Emu(11274552), Emu(430887))
tf = tb.text_frame
tf.word_wrap = True
tf.margin_left = tf.margin_right = Emu(0)
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.LEFT
run = p.add_run()
run.text = ("Eight principles guide the Target Operating Model — each answers one of the structural issues "
            "and is grounded in prior NTT work with Avangrid. From this Baseline: BRANCH-UNIFIED (Apps + CS) "
            "and PRODUCT-ORIENTED (Value Streams, not silos).")
run.font.name = 'Helvetica Neue'
run.font.size = Pt(14)
run.font.bold = True
run.font.color.rgb = RGBColor(0x2E, 0x6C, 0xB3)  # navy-blue like existing headers

prs.save(SRC)
print("Fixed")
