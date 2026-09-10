"""
Apply principle revisions to slide 47 of the AS-IS deliverable deck:
  1. Update #04 title + description (Standardized Execution → Execution Platforms, Jira+PPM Now)
  2. Add #08 Hybrid Delivery Model in the right-column / row-4 slot
  3. Relocate the green "+ From this Baseline" callout to the top subtitle
  4. Delete the in-slide comment speech bubble
  5. Update the header subtitle from "Seven" to "Eight" principles
"""
import copy
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from lxml import etree

SRC = '/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/2604_AS-IS deliverable.pptx'
OUT = '/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/2604_AS-IS deliverable.pptx'

prs = Presentation(SRC)
slide = prs.slides[46]  # zero-indexed, slide 47

# ---------- 1. Edit card 04 ----------
def set_text(tf, text):
    # preserve first run's formatting
    p = tf.paragraphs[0]
    if not p.runs:
        run = p.add_run()
    else:
        run = p.runs[0]
        # drop any extra runs
        for extra in p.runs[1:]:
            extra._r.getparent().remove(extra._r)
    run.text = text
    # remove extra paragraphs
    for extra in list(tf.paragraphs[1:]):
        extra._p.getparent().remove(extra._p)

set_text(slide.shapes[18].text_frame, "Standardized Execution Platforms")
set_text(slide.shapes[19].text_frame,
         "Jira for Agile delivery (small / medium + evolutive work); PPM Now for Waterfall delivery "
         "(large projects + CAPEX). Consistent tracking and management across both tracks.")

# ---------- 2. Add #08 Hybrid Delivery Model ----------
# Clone card 04's four shapes: background [15], accent [16], number [17], title [18], description [19]
# Offset to right column: delta_x = 5988938 - 246506 = 5742432
DX = 5988938 - 246506

def clone_shape(src_shape):
    new_el = copy.deepcopy(src_shape._element)
    src_shape._element.getparent().append(new_el)
    return new_el

# Clone background rect
src_idx_list = [15, 16, 17, 18, 19]
cloned_elements = []
for idx in src_idx_list:
    el = clone_shape(slide.shapes[idx])
    cloned_elements.append(el)

# Shift each clone horizontally
def shift_x(el, dx):
    # Find a:off in spPr/xfrm
    ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
          'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
    off = el.find('.//a:xfrm/a:off', ns)
    if off is not None:
        cur_x = int(off.get('x'))
        off.set('x', str(cur_x + dx))

for el in cloned_elements:
    shift_x(el, DX)

# Refresh shape list — cloned shapes are now at the end
new_shapes = [sh for sh in slide.shapes][-5:]

# Update their text (number, title, description)
# new_shapes[0] = background (no text)
# new_shapes[1] = accent (no text)
# new_shapes[2] = "04" → "08"
# new_shapes[3] = title → "Hybrid Delivery Model"
# new_shapes[4] = description → new

set_text(new_shapes[2].text_frame, "08")
set_text(new_shapes[3].text_frame, "Hybrid Delivery Model")
set_text(new_shapes[4].text_frame,
         "Delivery method matched to demand type: Agile (Jira / squads) for small–medium + evolutive work; "
         "Waterfall (PDD + PPM Now) for large, CAPEX or regulatory projects. Shared intake and governance.")

# ---------- 3. Remove the green "+ From this Baseline" callout (shapes 35 + 36)
# We keep the message but move it into the header subtitle [41]
green_bg = slide.shapes[35]
green_tx = slide.shapes[36]
green_bg._element.getparent().remove(green_bg._element)
green_tx._element.getparent().remove(green_tx._element)

# ---------- 4. Remove the Gustavo comment speech bubble [39] ----------
bubble = slide.shapes[39]
bubble._element.getparent().remove(bubble._element)

# ---------- 5. Update header subtitle to "Eight" and add BRANCH / PRODUCT-ORIENTED ----------
# Slide 41 (index 41 in original list) is now at different index after deletions — re-fetch by text
for sh in slide.shapes:
    if sh.has_text_frame and sh.text_frame.text.startswith("Seven principles"):
        set_text(sh.text_frame,
                 "Eight principles guide the Target Operating Model — each answers one of the structural issues "
                 "and is grounded in prior NTT work with Avangrid. From this Baseline: BRANCH-UNIFIED (Apps + CS) "
                 "and PRODUCT-ORIENTED (Value Streams, not silos).")
        break

prs.save(OUT)
print("Saved:", OUT)
