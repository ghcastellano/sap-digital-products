"""Append a 'Value Lens' slide: Bain 30 Elements of Value pyramid + Heather's metrics overlay."""
from copy import deepcopy
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

FILE = "Material em andamento/Customer Services - Proposed Value Streams - v2.pptx"

NAVY  = RGBColor(0x0F, 0x1B, 0x4F)
BLUE2 = RGBColor(0x2C, 0x57, 0x8B)
BLUE3 = RGBColor(0x5B, 0x7F, 0xB5)
BLUE4 = RGBColor(0xA5, 0xC2, 0xDA)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK  = RGBColor(0x2C, 0x3E, 0x50)
GRAY  = RGBColor(0x95, 0xA5, 0xA6)
HIGHLIGHT = RGBColor(0xFF, 0xC4, 0x3D)  # mustard for Heather highlights


def clone_slide_keep_bg(prs, source_idx):
    """Clone slide; keep only header/background-related shapes (top bar + footer logos)."""
    src = prs.slides[source_idx]
    new = prs.slides.add_slide(src.slide_layout)
    for shp in list(new.shapes):
        sp = shp._element
        sp.getparent().remove(sp)
    # copy bg
    src_bg = src._element.find(qn('p:cSld') + '/' + qn('p:bg'))
    new_cSld = new._element.find(qn('p:cSld'))
    if src_bg is not None and new_cSld is not None:
        existing_bg = new_cSld.find(qn('p:bg'))
        if existing_bg is not None:
            new_cSld.remove(existing_bg)
        new_cSld.insert(0, deepcopy(src_bg))
    # copy only top header & subtitle textboxes (idx 0, 1) from source slide 4 (a Em Branco slide)
    for el_idx, el in enumerate(src.shapes._spTree):
        tag = el.tag.split('}', 1)[-1]
        if tag in ('nvGrpSpPr', 'grpSpPr'):
            continue
    return new


def add_text(slide, x, y, w, h, text, *, size=10, bold=False, italic=False,
             color=DARK, name="Arial", align=PP_ALIGN.LEFT, anchor=None, wrap=True):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.margin_left = Inches(0.04); tf.margin_right = Inches(0.04)
    tf.margin_top = Inches(0.02); tf.margin_bottom = Inches(0.02)
    if anchor: tf.vertical_anchor = anchor
    lines = text.split('\n')
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run(); r.text = line
        r.font.name = name
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color
    return tb


def add_rect(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None):
    sh = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(0.5)
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        sh.adjustments[0] = radius
    return sh


def add_chip(slide, x, y, w, h, text, *, fill=WHITE, txt_color=DARK, border=None,
             font_size=7.0, bold=False):
    """Small rounded rectangle chip with centered text."""
    add_rect(slide, x, y, w, h, fill=fill, line=border or NAVY,
             shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.30)
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    tf.margin_left = Inches(0.02); tf.margin_right = Inches(0.02)
    tf.margin_top = Inches(0.01); tf.margin_bottom = Inches(0.01)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text
    r.font.name = "Arial"; r.font.size = Pt(font_size); r.font.bold = bold; r.font.color.rgb = txt_color


def add_chips_row(slide, x_start, y, total_w, h, items, *, gap=0.06, **chip_kwargs):
    n = len(items)
    chip_w = (total_w - (n - 1) * gap) / n
    for i, label in enumerate(items):
        cx = x_start + i * (chip_w + gap)
        add_chip(slide, cx, y, chip_w, h, label, **chip_kwargs)


def main():
    prs = Presentation(FILE)
    # Clone slide 4 (a "Em Branco" + has header bar). Then strip everything except header.
    src = prs.slides[3]  # slide 4 = framing slide (we'll wipe its body)
    # Actually we want a fresh layout. Easier: clone the OVS template slide (slide 6) and wipe.
    src = prs.slides[5]
    new_slide = prs.slides.add_slide(src.slide_layout)
    # remove default placeholders
    for shp in list(new_slide.shapes):
        sp = shp._element
        sp.getparent().remove(sp)
    # copy background
    src_bg = src._element.find(qn('p:cSld') + '/' + qn('p:bg'))
    new_cSld = new_slide._element.find(qn('p:cSld'))
    if src_bg is not None and new_cSld is not None:
        existing_bg = new_cSld.find(qn('p:bg'))
        if existing_bg is not None:
            new_cSld.remove(existing_bg)
        new_cSld.insert(0, deepcopy(src_bg))

    slide = new_slide

    # --- Header (mirror existing slides) ---
    add_text(slide, 0.40, 0.10, 12.53, 0.38, "VALUE LENS",
             size=20, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    add_text(slide, 0.40, 0.58, 12.53, 0.23,
             "BAIN 30 ELEMENTS OF VALUE  ·  HEATHER'S METRICS OVERLAY (CS PRODUCTS · SESSION INPUT)",
             size=11, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # --- Layout: pyramid on left (x 0.40..8.30, w 7.90), Heather panel right (x 8.45..12.93, w 4.48) ---

    # ---------- LEFT: Bain Pyramid ----------
    px = 0.40
    pw = 7.90
    py_start = 1.00

    # Tier definitions (label, color, items)
    TIERS = [
        ("SOCIAL IMPACT", NAVY,  ["Self-transcendence"], 1),
        ("LIFE CHANGING", BLUE2, ["Provides hope", "Self-actualization", "Motivation",
                                  "Heirloom", "Affiliation / belonging"], 1),
        ("EMOTIONAL",     BLUE3, ["Reduces anxiety", "Rewards me", "Nostalgia",
                                  "Design / aesthetics", "Badge value",
                                  "Wellness", "Therapeutic value", "Fun / entertainment",
                                  "Attractiveness", "Provides access"], 2),
        ("FUNCTIONAL",    BLUE4, ["Saves time", "Simplifies", "Makes money", "Reduces risk",
                                  "Organizes", "Integrates", "Connects",
                                  "Reduces effort", "Avoids hassles", "Reduces cost",
                                  "Quality", "Variety", "Sensory appeal", "Informs"], 2),
    ]

    # Heather's items map to specific Bain elements — used to highlight them
    HEATHER_HIGHLIGHTS = {
        "Reduces anxiety", "Wellness", "Therapeutic value",   # CSAT / customer satisfaction
        "Saves time", "Simplifies", "Reduces effort",         # outcome > outputs / shorten cycle
        "Quality", "Reduces risk", "Avoids hassles",          # reliability
        "Connects", "Integrates", "Informs",                  # open channels / short feedback
    }

    label_w = 1.20
    chips_x = px + label_w + 0.10
    chips_w = pw - label_w - 0.10
    chip_h = 0.35
    chip_v_gap = 0.08
    band_v_gap = 0.10

    cur_y = py_start
    for tier_name, tier_color, items, rows in TIERS:
        # split items into rows
        per_row = -(-len(items) // rows)  # ceil
        rows_items = [items[i*per_row:(i+1)*per_row] for i in range(rows)]
        band_h = rows * chip_h + (rows - 1) * chip_v_gap + 0.20

        # Tier label (left colored block)
        add_rect(slide, px, cur_y, label_w, band_h, fill=tier_color,
                 shape=MSO_SHAPE.RECTANGLE)
        tb = slide.shapes.add_textbox(Inches(px), Inches(cur_y), Inches(label_w), Inches(band_h))
        tf = tb.text_frame; tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Inches(0.05); tf.margin_right = Inches(0.05)
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = tier_name
        r.font.name = "Arial"; r.font.size = Pt(9.5); r.font.bold = True; r.font.color.rgb = WHITE

        # Chips
        for ri, row_items in enumerate(rows_items):
            n = len(row_items)
            row_y = cur_y + 0.10 + ri * (chip_h + chip_v_gap)
            gap = 0.08
            chip_w = (chips_w - (n - 1) * gap) / n
            for i, label in enumerate(row_items):
                cx = chips_x + i * (chip_w + gap)
                if label in HEATHER_HIGHLIGHTS:
                    add_chip(slide, cx, row_y, chip_w, chip_h, label,
                             fill=HIGHLIGHT, txt_color=NAVY, border=NAVY,
                             font_size=7.5, bold=True)
                else:
                    add_chip(slide, cx, row_y, chip_w, chip_h, label,
                             fill=WHITE, txt_color=DARK, border=tier_color,
                             font_size=7.0)

        cur_y += band_h + band_v_gap

    # Pyramid footer source line
    add_text(slide, px, cur_y + 0.05, pw, 0.18,
             "Source: Bain & Company / HBR — \"The Elements of Value\" (Sept 2016)  ·  "
             "Highlighted (▮ mustard) = elements aligned to Heather's stated metrics",
             size=7.0, italic=True, color=GRAY, align=PP_ALIGN.LEFT)

    # ---------- RIGHT: Heather's metrics box ----------
    hx = 8.45
    hy = 1.00
    hw = 4.48
    hh = 5.95

    # Outer box (white body + navy header)
    add_rect(slide, hx, hy + 0.05, hw, hh - 0.05,
             fill=WHITE, line=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
    add_rect(slide, hx - 0.02, hy, hw + 0.04, 0.55,
             fill=NAVY, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)

    add_text(slide, hx + 0.10, hy + 0.10, hw - 0.20, 0.40,
             "HEATHER'S METRICS — SESSION INPUT",
             size=11.5, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    add_text(slide, hx + 0.15, hy + 0.65, hw - 0.30, 0.22,
             "Heather LaJoie · Director Customer Service Products  ·  flagged as her own ask",
             size=8.0, italic=True, color=DARK, align=PP_ALIGN.LEFT)

    # Grouped bullets
    sec_y = hy + 0.95
    sections = [
        ("CUSTOMER OUTCOME METRICS", [
            "NPS",
            "CSAT",
            "Customer satisfaction",
            "Reliability",
        ]),
        ("LISTENING & FEEDBACK", [
            "Open channels to listen end-users",
            "Short feedback loops",
            "Feedback loops to improve the solutions",
        ]),
        ("DELIVERY MODEL", [
            "Customer-centric · customer-service first",
            "Customer approach service — deliver features → outcomes",
            "Outcome > Outputs",
            "Effectiveness",
        ]),
        ("SPEED & FLOW", [
            "Ideation-to-delivery is too long → shorten the cycle",
            "Reduce handoff delays",
        ]),
        ("CONSTRAINT TO REMOVE", [
            "Siloed demand intake — no global view",
        ]),
    ]

    cur = sec_y
    for sec_title, items in sections:
        add_text(slide, hx + 0.15, cur, hw - 0.30, 0.22, sec_title,
                 size=8.5, bold=True, color=NAVY, align=PP_ALIGN.LEFT)
        cur += 0.25
        body = "\n".join(f"•  {b}" for b in items)
        h_block = 0.20 * len(items) + 0.05
        add_text(slide, hx + 0.20, cur, hw - 0.40, h_block, body,
                 size=8.5, color=DARK, align=PP_ALIGN.LEFT)
        cur += h_block + 0.10

    # Footer attribution / how-to-use
    add_text(slide, hx + 0.15, hy + hh - 0.55, hw - 0.30, 0.50,
             "How to use: every DVS backlog item should map to ≥ 1 Bain element. "
             "Heather's metrics become the OVS-level KPI panel for Meter-to-Cash.",
             size=7.5, italic=True, color=GRAY, align=PP_ALIGN.LEFT)

    prs.save(FILE)
    print(f"Saved → {FILE}")
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
