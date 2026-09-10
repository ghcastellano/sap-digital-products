"""
TOM MAP v3 — In-content red edits (rebuild from workbook-2).

Replaces workbook-3.pptx (which had ANA badges) with proper in-content changes:
  - Each modified text run is colored RED · same font/size kept
  - No "ANA · " explanation badges
  - 6 new slides preserved (already content)
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

SRC = "/Users/gustavohenriquecastellano/Downloads/Base Apresentações Projeto Digital Product/TOM MAP workbook-2.pptx"
DST = "/Users/gustavohenriquecastellano/Downloads/Base Apresentações Projeto Digital Product/TOM MAP workbook - final.pptx"

# === Palette ===
HEADER_BLUE      = RGBColor(0x2C, 0x57, 0x8B)
AVANGRID_GREEN   = RGBColor(0x00, 0x8C, 0x38)
GREEN_CARD       = RGBColor(0x00, 0x7A, 0x32)
GREEN_DARK       = RGBColor(0x00, 0x6B, 0x3F)
GREEN_MID        = RGBColor(0x00, 0x89, 0x50)
GREEN_LIGHT      = RGBColor(0x00, 0x9F, 0x5D)
GREEN_BG         = RGBColor(0xF2, 0xF8, 0xF4)
GREY_BG          = RGBColor(0xF4, 0xF5, 0xF7)  # light gray bg for form strips
GREEN_PASTEL     = RGBColor(0xC8, 0xE0, 0xCB)
RED_EDIT         = RGBColor(0x80, 0x60, 0x00)  # subtle dark amber/olive · indicates edit without screaming
WHITE            = RGBColor(0xFF, 0xFF, 0xFF)
DARK_TEXT        = RGBColor(0x2C, 0x3E, 0x50)
GREY_LABEL       = RGBColor(0x6B, 0x77, 0x8C)

TITLE_FONT = "IberPangea Text Medium"
BODY_FONT  = "Arial"

# === Helpers (no-stroke filled boxes) ===
def _no_line(shape):
    shape.line.fill.background()

def rect(slide, x, y, w, h, fill=None):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    s.shadow.inherit = False
    if fill is None: s.fill.background()
    else: s.fill.solid(); s.fill.fore_color.rgb = fill
    _no_line(s)
    return s

def round_rect(slide, x, y, w, h, fill=None, radius=0.04):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    s.shadow.inherit = False
    s.adjustments[0] = radius
    if fill is None: s.fill.background()
    else: s.fill.solid(); s.fill.fore_color.rgb = fill
    _no_line(s)
    return s

def text(slide, x, y, w, h, txt, size=10, bold=False, italic=False,
         color=DARK_TEXT, font=BODY_FONT,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    """Create textbox · supports \\n for multi-paragraph content (each line a separate paragraph)."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.margin_left = Emu(36000); tf.margin_right = Emu(36000)
    tf.margin_top = Emu(18000); tf.margin_bottom = Emu(18000)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    lines = txt.split("\n") if isinstance(txt, str) else [str(txt)]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run(); r.text = line
        r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
        r.font.name = font
        r.font.color.rgb = color
    return tb

def header_bar(slide, w_total, line1, line2):
    rect(slide, 0, 0, w_total, Inches(0.95), HEADER_BLUE)
    rect(slide, 0, Inches(0.95), w_total, Emu(15000), AVANGRID_GREEN)
    text(slide, Inches(0.40), Inches(0.10), Inches(11), Inches(0.42),
         line1, size=22, bold=True, color=WHITE, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
    text(slide, Inches(0.40), Inches(0.55), Inches(11), Inches(0.32),
         line2, size=10, bold=True, color=WHITE, font=BODY_FONT)
    text(slide, Inches(11.4), Inches(0.30), Inches(1.7), Inches(0.4),
         "NTT  DATA", size=14, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.RIGHT)

def footer(slide, label):
    text(slide, Inches(0.40), Inches(7.18), Inches(8), Inches(0.18),
         "Avangrid · Digital Products · Target Operating Model",
         size=8, italic=True, color=GREY_LABEL, font=BODY_FONT)
    text(slide, Inches(11.5), Inches(7.18), Inches(1.6), Inches(0.18),
         label, size=8, color=GREY_LABEL, font=BODY_FONT, align=PP_ALIGN.RIGHT)

def cs_itba_badge(slide, x, y, applies_to):
    color = {"CS": AVANGRID_GREEN, "ITBA": HEADER_BLUE, "BOTH": GREY_LABEL}[applies_to]
    w, h = Inches(0.95), Inches(0.22)
    round_rect(slide, x, y, w, h, color, radius=0.30)
    text(slide, x, y, w, h, applies_to, size=7, bold=True, color=WHITE,
         font=BODY_FONT, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

def move_slide(prs, old_index, new_index):
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    s = slides[old_index]
    xml_slides.remove(s)
    xml_slides.insert(new_index, s)


def delete_slide(prs, idx):
    """Delete slide at 0-based index."""
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    xml_slides.remove(slides[idx])


def delete_shapes_with_text(slide, text_substrings):
    """Delete shapes whose text frame contains any of the given substrings."""
    deleted = 0
    for sh in list(slide.shapes):
        if not sh.has_text_frame: continue
        txt = sh.text_frame.text
        for sub in text_substrings:
            if sub and sub in txt:
                sp = sh._element
                sp.getparent().remove(sp)
                deleted += 1
                break
    return deleted


def delete_shapes_in_region(slide, x_min, y_min, x_max, y_max, only_empty=True):
    """Delete shapes positioned within a bounding box. If only_empty, restrict to text-empty shapes."""
    deleted = 0
    for sh in list(slide.shapes):
        try:
            x = sh.left/914400 if sh.left else 0
            y = sh.top/914400 if sh.top else 0
        except: continue
        if not (x_min <= x <= x_max and y_min <= y <= y_max): continue
        if only_empty and sh.has_text_frame:
            if sh.text_frame.text.strip(): continue
        sp = sh._element
        sp.getparent().remove(sp)
        deleted += 1
    return deleted


def find_idx_by_unique(prs, sub):
    for i, slide in enumerate(prs.slides):
        for sh in slide.shapes:
            if sh.has_text_frame and sub.upper() in sh.text_frame.text.upper():
                return i
    return -1


# ============================================================================
# IN-PLACE TEXT EDITS — find specific runs and replace, coloring red
# ============================================================================

def edit_run_in_slide(slide, target_substr, new_text):
    """Find run containing target_substr · replace its text · color RED · keep original font/size."""
    for sh in slide.shapes:
        if not sh.has_text_frame: continue
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if target_substr in r.text:
                    r.text = new_text
                    r.font.color.rgb = RED_EDIT
                    return True
    return False


def add_red_text_box(slide, x, y, w, h, txt, size=9, bold=False, italic=False,
                     align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    """Add a new text box with red text — used when adding NEW content (not annotation)."""
    return text(slide, x, y, w, h, txt, size=size, bold=bold, italic=italic,
                color=RED_EDIT, font=BODY_FONT, align=align, anchor=anchor)


def global_replace_run(prs, slide_range, old, new, color_red=True):
    """Walk all runs in slides within range; replace old text with new."""
    count = 0
    for i, slide in enumerate(prs.slides, 1):
        if not (slide_range[0] <= i <= slide_range[1]): continue
        for sh in slide.shapes:
            if not sh.has_text_frame: continue
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if old in r.text:
                        r.text = r.text.replace(old, new)
                        if color_red: r.font.color.rgb = RED_EDIT
                        count += 1
    return count


def opco_mult_panel(slide, x, y, w, h):
    """Add red panel explaining OPCOmult — used on slides 31, 32."""
    round_rect(slide, x, y, w, h, GREEN_BG, radius=0.04)
    text(slide, x + Inches(0.12), y + Inches(0.06), w - Inches(0.20), Inches(0.22),
         "OPCOmult  ·  multiplier on BV based on # OPCOs impacted",
         size=10, bold=True, color=RED_EDIT, font=BODY_FONT)
    text(slide, x + Inches(0.12), y + Inches(0.30), w - Inches(0.20), h - Inches(0.36),
         "1.0  ·  single OPCO          1.3  ·  regional (2–4 OPCOs)          1.6  ·  all 5+ OPCOs\n"
         "A1 detects duplicates across OPCOs and merges into one card with tag [N OPCOs]",
         size=9, color=RED_EDIT, font=BODY_FONT)


def replace_shape_text_red(slide, anchor_substring, new_text, size=None, bold=None):
    """Find shape containing anchor_substring · clear · write new text in RED · keep one paragraph."""
    for sh in slide.shapes:
        if not sh.has_text_frame: continue
        if anchor_substring not in sh.text_frame.text: continue
        tf = sh.text_frame
        # Capture original font properties from first run
        orig_size = orig_bold = orig_font = None
        for p in tf.paragraphs:
            for r in p.runs:
                try: orig_size = r.font.size
                except: pass
                try: orig_bold = r.font.bold
                except: pass
                try: orig_font = r.font.name
                except: pass
                break
            break
        # Clear and rewrite
        for p in list(tf.paragraphs):
            for r in list(p.runs):
                r.text = ""
        # Reset to single paragraph with new text
        p = tf.paragraphs[0]
        # remove other runs in p
        for r in list(p.runs)[1:]:
            r._r.getparent().remove(r._r)
        if not p.runs:
            r = p.add_run()
        else:
            r = p.runs[0]
        r.text = new_text
        if size: r.font.size = Pt(size)
        elif orig_size: r.font.size = orig_size
        if bold is not None: r.font.bold = bold
        elif orig_bold is not None: r.font.bold = orig_bold
        if orig_font: r.font.name = orig_font
        r.font.color.rgb = RED_EDIT
        return True
    return False


def cleanup_opco_panels_on_slide(slide):
    """Remove the OPCO panels I previously added that overlap content."""
    targets = [
        "OPCOmult  ·  multiplier on BV",
        "1.0  ·  single OPCO",
        "A1 detects duplicates across OPCOs",
        "EXAMPLE  ·  same BV=4 with different OPCO",
        "1 OPCO (UIL)",
        "Regional (UIL + CMP + NY)",
        "All 5+ OPCOs (cross-OPCO)",
        "+60% boost · justifies cross-OPCO",
    ]
    deleted = 0
    for sh in list(slide.shapes):
        if not sh.has_text_frame: continue
        for t in targets:
            if t in sh.text_frame.text:
                sh._element.getparent().remove(sh._element)
                deleted += 1
                break
    # Also remove GREEN_BG rounded rect containers at y > 6.40 (the panel backgrounds)
    for sh in list(slide.shapes):
        try:
            y = sh.top/914400 if sh.top else 0
            w = sh.width/914400 if sh.width else 0
        except: continue
        if y > 6.40 and w > 10 and y < 7.6:
            try:
                if sh.fill.type == 1:
                    rgb = str(sh.fill.fore_color.rgb)
                    # GREEN_BG = F2F8F4 — that's the panel background color I used
                    if rgb == "F2F8F4":
                        if not sh.has_text_frame or sh.text_frame.text.strip() == "":
                            sh._element.getparent().remove(sh._element)
                            deleted += 1
            except: pass
    return deleted


def add_opcomult_5th_box_to_slide(prs, anchor):
    """Resize the 4 existing boxes (BV/UX/Safety/Effort) and add OPCOmult as 5th box."""
    idx = find_idx_by_unique(prs, anchor)
    if idx < 0:
        return False
    slide = prs.slides[idx]
    # First clean any existing OPCO panel additions
    cleanup_opco_panels_on_slide(slide)

    # Original 4 box layout: x_starts = 0.40, 3.53, 6.66, 9.79; each w=2.95
    OLD_X = [0.40, 3.53, 6.66, 9.79]
    OLD_W = 2.95
    BOX_Y = 2.90
    BOX_H = 2.40

    # New 5-column layout: 5 boxes width 2.43, gap 0.10, margin 0.40
    NEW_W = 2.43
    NEW_GAP = 0.10
    NEW_X = [0.40 + i * (NEW_W + NEW_GAP) for i in range(5)]

    # Resize and reposition each of the 4 existing boxes (body + header + text shapes)
    scale = NEW_W / OLD_W
    for box_i, old_x in enumerate(OLD_X):
        new_x = NEW_X[box_i]
        for sh in slide.shapes:
            try:
                sx = sh.left/914400 if sh.left else 0
                sw = sh.width/914400 if sh.width else 0
                sy = sh.top/914400 if sh.top else 0
                sh_h = sh.height/914400 if sh.height else 0
            except: continue
            # Skip full-width shapes (header bars, formula bar)
            if sw > 10: continue
            # Skip shapes outside the y-range of the 4 boxes
            if sy < BOX_Y - 0.05 or sy > BOX_Y + BOX_H + 0.10: continue
            # Determine if shape belongs to this box (center within box x range)
            center = sx + sw / 2
            if not (old_x - 0.05 <= center <= old_x + OLD_W + 0.05): continue
            # Reposition + resize
            offset_within = sx - old_x
            new_sx = new_x + offset_within * scale
            new_sw = sw * scale
            sh.left = Inches(new_sx)
            sh.width = Inches(new_sw)

    # Now add 5th box (OPCOmult) at NEW_X[4]
    box5_x = NEW_X[4]
    OPCO_FILL = RGBColor(0x00, 0xC0, 0x00)  # green tone for OPCOmult — actually use red-tinted
    OPCO_FILL = RED_EDIT  # since OPCOmult is the red edit, use red header
    # Body
    body = round_rect(slide, Inches(box5_x), Inches(BOX_Y), Inches(NEW_W), Inches(BOX_H), WHITE, radius=0.04)
    # Header
    header = round_rect(slide, Inches(box5_x), Inches(BOX_Y), Inches(NEW_W), Inches(0.65), OPCO_FILL, radius=0.04)
    # Title "OPCOmult"
    text(slide, Inches(box5_x), Inches(BOX_Y + 0.04), Inches(NEW_W), Inches(0.32),
         "OPCOmult", size=18, bold=True, color=WHITE, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Subtitle "multiplier"
    text(slide, Inches(box5_x), Inches(BOX_Y + 0.40), Inches(NEW_W), Inches(0.22),
         "multiplier", size=11, color=WHITE, font=BODY_FONT,
         align=PP_ALIGN.CENTER)
    # Field label
    text(slide, Inches(box5_x + 0.12), Inches(BOX_Y + 0.78), Inches(NEW_W - 0.24), Inches(0.22),
         "OPCO scope multiplier",
         size=10, bold=True, color=RED_EDIT, font=BODY_FONT)
    # Scoring location
    text(slide, Inches(box5_x + 0.12), Inches(BOX_Y + 0.98), Inches(NEW_W - 0.24), Inches(0.22),
         "applied to BV at Analyzing",
         size=9, italic=True, color=RED_EDIT, font=BODY_FONT)
    # Body bullets
    text(slide, Inches(box5_x + 0.12), Inches(BOX_Y + 1.28), Inches(NEW_W - 0.24), Inches(1.10),
         "Values:\n· 1.0  =  single OPCO\n· 1.3  =  regional (2–4)\n· 1.6  =  all 5+ OPCOs\n· A1 dedup auto",
         size=9, color=DARK_TEXT, font=BODY_FONT)
    return True


def opco_example_panel(slide, x, y, w, h):
    """Add red panel with numerical OPCO example — slide 32."""
    round_rect(slide, x, y, w, h, GREEN_BG, radius=0.04)
    text(slide, x + Inches(0.12), y + Inches(0.06), w - Inches(0.20), Inches(0.22),
         "EXAMPLE  ·  same BV=4 with different OPCO scope (ITBA · w₁=0.45)",
         size=9, bold=True, color=RED_EDIT, font=BODY_FONT)
    text(slide, x + Inches(0.12), y + Inches(0.28), w - Inches(0.20), h - Inches(0.34),
         "1 OPCO (UIL)                       →  4 × 1.0 × 0.45 = 1.80\n"
         "Regional (UIL + CMP + NY)  →  4 × 1.3 × 0.45 = 2.34\n"
         "All 5+ OPCOs (cross-OPCO) →  4 × 1.6 × 0.45 = 2.88     +60% boost · justifies cross-OPCO priority",
         size=9, color=RED_EDIT, font=BODY_FONT)


# ============================================================================
# Apply edits per slide
# ============================================================================

EDITS = {
    "BV (slide 24)": [
        # [08] subtitle change
        ("Used in Pain Score formula · scored at triage with the Champion",
         "ONE of 4 inputs into Pain Score · scored at Analyzing by Architect"),
        # [13] TRANSFORMATIONAL desc — remove regulatory non-compliance (covered by Regulatory flag)
        ("Multi-million  impact · safety · regulatory non-compliance · existential",
         "Multi-million impact · safety · existential business risk"),
        # [59] WHY IT MATTERS — Champion → Reviewer
        ("Champions can't game ranking",
         "Reviewers can't game ranking"),
    ],
    "CoS (slide 27)": [
        # [13] CAPACITY ALLOCATION values — softened (no hardcoded %)
        ("60–70% of squad capacity",
         "Majority share · % set quarterly at BRP"),
        ("15–25% of capacity (highest in Corporate)",
         "Regulatory-driven · % set quarterly at BRP"),
        ("5–10% capacity reserve · blocked-out per quarter",
         "Reserved for P1/P2 · % set quarterly at BRP"),
        ("10–15% capacity reserved per quarter",
         "Tech debt allocation · % set quarterly at BRP"),
        # [35] Expedite WHEN-TO-USE — inline define CIT P1/P2 (rest of P3/P4 added separately)
        ("P1/P2 customer impact · production outage · regulatory deadline imminent (< 30 days)",
         "P1 (>100k customers OR full OPCO outage) · P2 (>10k customers · degraded service) · regulatory deadline < 30 days"),
    ],
    "Type+Size (slide 28)": [
        # [03] subtitle
        ("TYPE  ·  Type: Demand · Incident · Change  ·  Size: S · M · L · XL",
         "TYPE  ·  Type: Demand · Change (Incident auto from IT Now)  ·  Size: S · M · L · XL"),
        # [08] heading
        ("TYPE  ·  Demand · Incident · Change",
         "TYPE  ·  Demand · Change   (Incident auto from IT Now → not user-selectable)"),
        # [20] Incident description — explain auto from IT Now
        ("Production outage · break-fix · service degradation",
         "Production outage · break-fix · NOT user-selectable — auto-created from ServiceNow (IT Now)"),
        # [22] BEHAVIOR for Incident
        ("Bypasses Gates. Auto-converts to Bug. Routed directly to Squad Expedite swimlane. CoS = Expedite.",
         "Auto-created when ServiceNow P1/P2 ticket fires · API → Jira creates Bug · A2 sets CoS=Expedite + WIP-1"),
        # [30] BEHAVIOR for Change — remove ambiguity with regulatory flag
        ("Goes through Triage but with lighter DoR. Still requires Gate 2 if Size ≥ M. Common for Upgrades.",
         "SAP kernel · vendor refresh · config change. Lighter DoR. Gate 2 if Size ≥ M. (Regulatory uses separate flag → Fixed-Date)"),
    ],
    "Work Type (slide 29)": [
        # [24] Upgrade description — remove "regulatory mandates"
        ("SAP kernel patches · vendor refresh · regulatory mandates · forced by external",
         "SAP kernel patches · vendor refresh · forced by external (Regulatory uses separate flag)"),
    ],
    "Pain Origin (slide 30)": [
        # Title of right panel — show TWO variants
        ("CS ADAPTATION — Customer Services",
         "DIGITAL PRODUCTS · TWO VARIANTS · CS + ITBA"),
        ("Source · CS Proposed Value Streams Slide 3 · 4-factor · UX-heavy",
         "CS: 2 components (BV+UX · UX-heavy)  ·  ITBA: 3 components (BV+UX+Safety · BV-heavy)"),
        # Replace formula with both variants
        ("H(x)  =  ( BV·0.35 + UX·0.40 + Safety·0.25 )  /  Effort",
         "CS:  H(x) = (BV·OPCOmult·0.40 + UX·0.60) / Effort        ITBA:  H(x) = (BV·OPCOmult·0.45 + UX·0.30 + Safety·0.25) / Effort"),
        # Header — present clearly without "REMOVED" wording
        ("WHY 4 FACTORS · WHY UX-HEAVY",
         "WHY DIFFERENT WEIGHTS PER VS · CS=customer-first · ITBA=grid-first"),
        # 3 weight boxes — show CS and ITBA values side by side
        ("BV (35%)",
         "BV  ·  CS 0.40 / ITBA 0.45"),
        ("UX (40%)",
         "UX  ·  CS 0.60 / ITBA 0.30"),
        ("Safety (25%)",
         "Safety  ·  ITBA 0.25 only"),
        # Description boxes
        ("kept (utility realities)",
         "applies only to ITBA"),
        ("Field crew · public · regulatory",
         "ITBA: utility safety · field crew · public hazard · regulatory breach. CS: customer-facing · no crew exposure."),
    ],
    "Pain Decoded (slide 32)": [
        # [23] Safety description — clean ITBA/CS variant explanation
        ("Safety dimensions specific to utilities · field crew exposure · public hazard · regulatory breach · cyber",
         "ITBA: utility safety · field crew · public hazard · regulatory breach.  CS variant: H(x) = (BV·OPCOmult·0.40 + UX·0.60) / Effort"),
    ],
    "After Submission (slide 33)": [
        # [17] Triage cast align with Gate 1
        ("TRIAGE  ·  Demand Manager + Champion",
         "TRIAGE (Gate 1)  ·  Demand Manager (chair) + PM + Architect"),
        # [18] @notify Champion → @notify PM
        ("@notify Champion",
         "@notify PM"),
    ],
    # Field Progression (v1-built) — fix old "Demand · Incident · Change" triple
    "Field Progression (anchor)": [
        ("Type (Demand · Incident · Change)",
         "Type (Demand · Change · Incident auto from IT Now)"),
    ],
    # Value Stream Definition (slide 26) — @notify at intake should be PM, not Champion
    "Value Stream Definition (slide 26)": [
        ("@notify Champion · sets squad swimlane · routes to portfolio board",
         "@notify PM · sets squad swimlane · routes to portfolio board"),
        # WHY IT MATTERS — "Champion never gets notified" → "PM never gets notified"
        ("Champion never gets notified. Triage never happens.",
         "PM never gets notified. Triage never happens."),
    ],
    # Pain Origin (slide 30) — final bullet "Champion can't game" → "Reviewers"
    "Pain Origin extra (slide 30)": [
        ("Auto-computed removes politics · Champion can't game ranking",
         "Auto-computed removes politics · Reviewers can't game ranking"),
    ],
    # Pain Auto-Comp (slide 31)
    "Pain Auto-Comp (slide 31)": [
        ("scored at TRIAGE with Champion",
         "scored at Analyzing by Architect"),
        ("Champion + PM",
         "PM + Architect"),
    ],
    # Pain Decoded (slide 32) — variable description fixes
    "Pain Decoded extra (slide 32)": [
        ("▸ scored by Champion at Triage",
         "▸ scored by Architect at Analyzing"),
        ("▸ scored by Champion + Architect at Triage",
         "▸ scored by Architect at Analyzing"),
    ],
    # Class of Service Type (slide 28) — remove "regulated change"
    "CoS Type extra (slide 28)": [
        ("Config change · vendor refresh · regulated change",
         "Config change · vendor refresh · upgrade"),
    ],
    "Governance (slide 42)": [
        # [42] BRP participants — couple to existing
        ("All VSs + business sponsors + Iberdrola",
         "Couples to existing P&T BRP + Digital CoE BRP — no separate event"),
        # [51] BRP description — already says cross-portfolio · keep but add couple note
        ("Cross-portfolio alignment. Surfaces dependencies early and locks quarterly OKRs.",
         "quarterly priorities summary enters EXISTING P&T BRP (Carlos+Asier) and Digital CoE BRP · Vitor liaises"),
    ],
}


# ============================================================================
# In-place edits to slides built by us in v1 (Team Kanban Deep + Value & Prio + BRP Deep)
# ============================================================================

V1_BUILT_EDITS = {
    "Team Kanban Deep": {
        "anchor": "TEAM KANBAN  ·  PHASE DEEP DIVE",
        "edits": [
            # Phase 0 OWNER — Ana: discovery side wont be us
            ("Product Manager", "Business / vendor PM   ·   NOT Digital Products"),
        ],
    },
    "BRP Deep": {
        "anchor": "BIG ROOM PLANNING  ·  DEEP DIVE",
        "edits": [
            # WHAT cell content
            ("Quarterly half-day session aligning all Value Streams: lock OKRs · "
             "map cross-VS dependencies · commit capacity per quarter.",
             "Digital Products does NOT run its own BRP · injects quarterly priorities summary into the existing "
             "P&T BRP (Carlos+Asier) and Digital CoE BRP."),
            # WHY cell content
            ("Surfaces dependencies BEFORE they become blockers · aligns portfolio with "
             "corporate OKRs · stops 'reg always wins' bypass.",
             "Avoids creating a parallel ritual · respects governance already in place · "
             "ensures cross-portfolio alignment without overhead."),
            # WHEN cell content
            ("First week of each quarter · 13:00–17:00 ET · 4 hours · "
             "in-person preferred · pre-read circulated 5 business days ahead.",
             "Aligned to existing P&T BRP cadence · quarterly priorities summary prepared 1 week before each event · "
             "no own meeting on Digital Products calendar."),
            # WHO cell content
            ("CHAIR  ·  VP Vitor Cruz   ·   DECIDERS  ·  Carlos Garbin · Kumar Jaladanki   ·   "
             "CHAMPIONS  ·  1 per VS   ·   SPONSORS  ·  business CS + ITBA   ·   "
             "Iberdrola liaison · Architect · Travis (PMO Lead)",
             "Digital Products has liaisons to existing BRPs · Carlos to P&T BRP · Heather to Digital CoE BRP · "
             "Vitor cross-portfolio liaison   ·   Architect supports both events. Kumar (CS) + Carlos (ITBA) are Digital Products Directors."),
        ],
    },
    "Value & Prio": {
        "anchor": "VALUE  &  PRIORITIZATION",
        "edits": [
            # Pain Score formula — add OPCO multiplier
            ("H(x)  =  ( BV · w₁  +  UX · w₂  +  Safety · w₃ )  ÷  Effort",
             "H(x)  =  ( BV · OPCOmult · w₁  +  UX · w₂  +  Safety · w₃ )  ÷  Effort"),
        ],
    },
    "Automation Catalog": {
        "anchor": "AUTOMATION CATALOG  ·  A1 → A6",
        "edits": [
            # A1 action — @notifies Champion → @notifies PM
            ("Reads Value Stream + OPCO → routes ticket to correct Portfolio Board · @notifies Champion",
             "Reads Value Stream + OPCO → routes ticket to correct Portfolio Board · @notifies PM"),
        ],
    },
    "Gates Deep Dive": {
        "anchor": "GATES DEEP DIVE",
        "edits": [
            # Gate 1 step 4 — @notify Champion → @notify PM (Champion only at Gate 2 per our model)
            ("Auto-routing recorded in Jira · @notify Champion of accepted items",
             "Auto-routing recorded in Jira · @notify PM of accepted items"),
        ],
    },
    "Field Progression": {
        "anchor": "FIELD PROGRESSION",
        "edits": [
            # Old Type triple
            ("Type (Demand · Incident · Change)",
             "Type (Demand · Change · Incident auto from IT Now)"),
            # Pain Score formula — add OPCO + CS/ITBA split
            ("H(x) = (BV·w₁ + UX·w₂ + Safety·w₃) / Effort",
             "ITBA: (BV·OPCOmult·w₁ + UX·w₂ + Safety·w₃) / Effort   ·   CS: (BV·OPCOmult·0.40 + UX·0.60) / Effort"),
            # Scores label — drop Safety for CS
            ("BV · UX · Safety · Effort — set at Analyzing",
             "ITBA: BV · UX · Safety · Effort   ·   CS: BV · UX · Effort (no Safety)"),
            # CRITICAL: Reviewing/Analyzing owner is PM (per VS), NOT Demand Manager
            # Demand Manager owns INTAKE QUALITY at Triage only; PM owns the VS backlog
            ("Demand Manager + Sponsor",
             "PM (per VS) + Sponsor"),
            ("Demand Manager + Architect",
             "PM (per VS) + Architect"),
        ],
    },
    # === V3.2 — comprehensive Safety/OPCO/PM-steward sweep ===
    "Automation Catalog A4": {
        "anchor": "AUTOMATION CATALOG  ·  A1 → A6",
        "edits": [
            ("Computes  H(x) = (BV·w₁ + UX·w₂ + Safety·w₃) / Effort  ·  writes Pain Score field",
             "Computes Pain Score · ITBA: (BV·OPCOmult·w₁+UX·w₂+Safety·w₃)/E · CS: (BV·OPCOmult·0.40+UX·0.60)/E · writes Pain Score"),
            ("BV · UX · Safety · Effort scores filled at Analyzing",
             "BV · UX · (Safety: ITBA only) · Effort  ·  filled at Analyzing"),
            # Trigger of A2 — Customer Impact Tier reference (already mentioned, but reinforce CIT scope)
            # already correct
        ],
    },
    # Pain Score Formula Decoded (slide 32) — main formula needs OPCO
    "Pain Decoded formula": {
        "anchor": "FORMULA DECODED",
        "edits": [
            ("H(x)  =  ( BV·w₁  +  UX·w₂  +  Safety·w₃ )  /  Effort",
             "ITBA: ( BV·OPCOmult·w₁ + UX·w₂ + Safety·w₃ ) / Effort   ·   CS: ( BV·OPCOmult·0.40 + UX·0.60 ) / Effort"),
        ],
    },
    # Pain Auto-Comp (slide 31) — main formula
    "Pain Auto-Comp formula": {
        "anchor": "PAIN SCORE · AUTO-COMPUTATION",
        "edits": [
            ("H(x)  =  ( BV·w₁  +  UX·w₂  +  Safety·w₃ )  /  Effort",
             "ITBA: ( BV·OPCOmult·w₁ + UX·w₂ + Safety·w₃ ) / Effort   ·   CS: ( BV·OPCOmult·0.40 + UX·0.60 ) / Effort"),
        ],
    },
    # Gates Deep Dive (slide 47) — replace "PM steward" with "PM observers" (PM steward was legacy)
    "Gates Deep Dive PM steward": {
        "anchor": "GATES DEEP DIVE",
        "edits": [
            ("▸ OBSERVERS · PM stewards (Heather + P&T pool) · Sponsor on request",
             "▸ OBSERVERS · PMs (CS + ITBA pool) · Sponsor on request"),
            ("▸ ATTENDEES · PM stewards · Architect · Travis (PMO Lead · for L/XL only)",
             "▸ ATTENDEES · PMs (per VS) · Architect · Travis (PMO Lead · for L/XL only)"),
            # CRITICAL: Champion is at GATE 2 ONLY — remove from Gate 1 attendees
            ("▸ ATTENDEES · Champions (one per VS) · Architect (cross-VS standards)",
             "▸ ATTENDEES · PMs (one per VS) · Architect (cross-VS standards)"),
            # Gate 1 step 4 already changed "@notify Champion" → "@notify PM" earlier
        ],
    },
    # === V3.3 — split-run formulas + Safety annotations + manager slides ===
    # Slide 31, 32 — split-run formula: modify the "BV·w" run to add OPCOmult
    "Pain Auto-Comp split formula": {
        "anchor": "PAIN SCORE · AUTO-COMPUTATION",
        "edits": [
            ("BV·w", "BV·OPCOmult·w"),  # only first occurrence; that's the formula run
        ],
    },
    "Pain Decoded split formula": {
        "anchor": "FORMULA DECODED",
        "edits": [
            ("BV·w", "BV·OPCOmult·w"),
        ],
    },
    # Slide 32 — Safety variable description: add "(ITBA only)"
    "Pain Decoded Safety annotations": {
        "anchor": "FORMULA DECODED",
        "edits": [
            ("weight of Safety", "weight of Safety (ITBA only)"),
            ("= how much Safety matters", "= how much Safety matters (ITBA only · CS drops it)"),
            ("Different Value Streams have different priorities. CS is customer-facing → UX ma",
             "Different Value Streams have different priorities. CS is customer-facing · UX-heavy and Safety REMOVED. ITBA / Corporate keep all 3 components."),
        ],
    },
    # Slide 24 — BV "INCLUDE NUMBERS" — Safety category example
    "BV Safety example annotation": {
        "anchor": "FIELD GUIDE – BUSINESS VALUE",
        "edits": [
            ("Field crew exposure · public hazard · near-miss",
             "Field crew exposure · public hazard · near-miss   (ITBA only · CS uses Customer Reliability instead)"),
        ],
    },
    # Slide 31 — re-computation triggers: BV/UX/Safety/Effort → annotate
    "Pain Auto-Comp triggers": {
        "anchor": "PAIN SCORE · AUTO-COMPUTATION",
        "edits": [
            ("Any BV / UX / Safety / Effort score edited",
             "Any BV / UX / Safety (ITBA only) / Effort score edited"),
            ("re-computed if BV/UX/Safety/Effort",
             "re-computed if BV/UX/Safety(ITBA)/Effort"),
        ],
    },
    # Manager slides 51, 53, 54, 55 — "Project Manager" → "Product Manager (PM)"
    "Manager Slide 51 PM": {
        "anchor": "DEMAND MANAGEMENT – INITIAL CLASSIFICATION",
        "edits": [
            ("Project Manager", "Product Manager"),
        ],
    },
    "Manager Slide 53 PM": {
        "anchor": "PREPARATION & PLANNING – PRIORITY BACKLOG",
        "edits": [
            ("Product Owner / Project Manager / Technical Leads",
             "Product Manager / Technical Leads"),
            ("Product Owner / Project Manager ", "Product Manager "),
        ],
    },
    "Manager Slide 54 PM": {
        "anchor": "PREPARATION & PLANNING – CAPACITY ALLOCATION",
        "edits": [
            ("Project Manager", "Product Manager"),
        ],
    },
    "Manager Slide 55 PM": {
        "anchor": "PREPARATION & PLANNING –  WORK ALLOCATION",
        "edits": [
            ("Project Manager", "Product Manager"),
        ],
    },
    # Slide 37 (PORTFOLIO BOARD) — Analyzing owner = PM, not Demand Manager
    "Portfolio Board owners": {
        "anchor": "PORTFOLIO BOARD",
        "edits": [
            ("Demand Manager + Architect",
             "PM (per VS) + Architect"),
            # Reviewing on Portfolio Board (if exists)
            ("Sponsor + Demand Manager",
             "Sponsor + PM (per VS)"),
        ],
    },
    # Slide 42 (STEP-LEVEL POLICIES) — same fix
    "Step-Level owners": {
        "anchor": "STEP-LEVEL POLICIES",
        "edits": [
            ("Owner: Demand Manager + Architect",
             "Owner: PM (per VS) + Architect"),
            ("Owner: Sponsor + Demand Manager",
             "Owner: Sponsor + PM (per VS)"),
        ],
    },
    # === V3.4 — TYPE TAXONOMY (Caminho 2 P&T-aligned) ===
    # Type = { Production Support · Regulatory · Non-Regulatory }
    # Work Type field REMOVED · absorbed into Type
    # Regulatory flag REMOVED · now Type=Regulatory triggers A3
    #
    # Slide 20 (form mockup detailed) — clean removal of Work Type + Regulatory flag
    "Form mockup detailed (slide 20)": {
        "anchor": "Avangrid · Digital Products  /  Create Issue",
        "edits": [
            # Type field value
            ("Demand", "Production Support · Regulatory · Non-Regulatory"),
            # Empty out Work Type field entirely (no historical wording)
            ("Work Type *", ""),
            ("▼ BAU · Build · Upgrade", ""),
            # Empty out Regulatory flag field entirely
            ("Regulatory flag  ·  if true · auto-Fixed-Date + Compliance sponsor", ""),
            ("☐ Yes (regulatory mandate)", ""),
            # Pain Score formula — show CS + ITBA variants
            ("H(x) = (BV·w₁ + UX·w₂ + Safety·w₃) / Effort",
             "ITBA: (BV·OPCOmult·w₁+UX·w₂+Safety·w₃)/E   ·   CS: (BV·OPCOmult·0.40+UX·0.60)/E"),
        ],
    },
    # Slide 19 (form mockup overview)
    "Form overview (slide 19)": {
        "anchor": "ONE JIRA FRONT DOOR",
        "edits": [
            # Old "Demand type · Run/Grow/Transform" — separate Type from R/G/T
            ("Run · Grow · Transform",
             "Production Support · Regulatory · Non-Regulatory"),  # this is Type now
            ("BV · UX · Safety · Effort",
             "ITBA: BV·UX·Safety·Effort   ·   CS: BV·UX·Effort (no Safety)"),
        ],
    },
    # Slide 21 Field Progression — drop Work Type lens + Regulatory flag
    "Field Progression Type/Flags": {
        "anchor": "FIELD PROGRESSION",
        "edits": [
            ("Work Type (BAU / Build / Upgrade)  ·  Run / Grow / Transform",
             "Run / Grow / Transform   (Work Type removed · absorbed into Type)"),
            ("Regulatory flag (Yes/No) · Customer Impact Tier (P1–P4 if known)",
             "Customer Impact Tier (P1–P4 if known)   (Regulatory flag removed · now Type=Regulatory)"),
        ],
    },
    # Slide 28 Type+Size — relabel 3 cards from {Demand, Incident, Change} to {Production Support, Regulatory, Non-Regulatory}
    "Slide 28 Type taxonomy rewrite": {
        "anchor": "Two operational fields that branch the path through gates",
        "edits": [
            # Subtitle
            ("TYPE  ·  Type: Demand · Change (Incident auto from IT Now)  ·  Size: S · M · L · XL",
             "TYPE  ·  Type: Production Support · Regulatory · Non-Regulatory  ·  Size: S · M · L · XL"),
            ("TYPE  ·  Demand · Change   (Incident auto from IT Now → not user-selectable)",
             "TYPE  ·  Production Support · Regulatory · Non-Regulatory   (Incident auto from IT Now · not in Type)"),
            # Card 1 (was Demand) → Production Support
            ("Demand",
             "Production Support"),
            ("New work · features · projects · enhancements",
             "Operational maintenance · break-fix manual · routine work · investigations · tuning"),
            ("Most submissions. Goes through full path: NEW → ROUTED → TRIAGED → BACKLOG → Gates → Squad.",
             "Operational work · light DoR · CoS=Standard or Intangible · Gate 2 only if Size ≥ M · most go to squad direct."),
            ("Standard intake flow",
             "Light DoR · Standard CoS · Squad direct (S) · Gate 2 (M+)"),
            # Card 2 (was Incident) → Regulatory
            ("Incident",
             "Regulatory"),
            ("Production outage · break-fix · NOT user-selectable — auto-created from ServiceNow (IT Now)",
             "Mandatory · external mandate · PUC · SOX · Iberdrola · vendor SLA · regulator-driven deadline"),
            ("Auto-created when ServiceNow P1/P2 ticket fires · API → Jira creates Bug · A2 sets CoS=Expedite + WIP-1",
             "Triggers A3 Auto-Fixed-Date · CoS=Fixed-Date · mandatory Gate 2 review · compliance sponsor required"),
            ("Skip BACKLOG · skip Gates · go to Squad Expedite",
             "Full Triage · auto Fixed-Date · Gate 2 mandatory · scheduled backwards from deadline"),
            # Card 3 (was Change) → Non-Regulatory
            ("Change",
             "Non-Regulatory"),
            ("SAP kernel · vendor refresh · config change. Lighter DoR. Gate 2 if Size ≥ M. (Regulatory uses separate flag → Fixed-Date)",
             "Business demand · new features · projects · enhancements · innovations · most common type"),
            ("Light DoR · Gate 2 if M+",
             "Full DoR · Gate 2 · CoS=Standard or Expedite (if P1/P2)"),
        ],
    },
    # Slide 29 — relabel from Work Type to Type taxonomy
    "Slide 29 rewrite Type vs RGT": {
        "anchor": "WORK TYPE VS RUN/GROW/TRANSFORM",
        "edits": [
            # Title
            ("FIELD GUIDE – WORK TYPE VS RUN/GROW/TRANSFORM",
             "FIELD GUIDE – TYPE VS RUN/GROW/TRANSFORM"),
            # Header card
            ("Work Type vs Run/Grow/Transform",
             "Type vs Run/Grow/Transform"),
            ("TYPE  ·  Two single-select fields · operational lens + strategic lens",
             "TYPE  ·  Two complementary lenses · WHAT (Type) + WHY (R/G/T)"),
            # Left column
            ("Work Type",
             "Type"),
            ("ANSWERS:  How will this work be executed?",
             "ANSWERS:  WHAT category of work?"),
            # 3 boxes
            ("BAU",
             "Production Support"),
            ("Business as Usual",
             "Operational maintenance"),
            ("Recurring · keep-the-lights-on · bug fixes · incidents · no capability change",
             "Recurring · break-fix · maintenance · routine investigations · keep-the-lights-on"),
            ("Build",
             "Non-Regulatory"),
            ("New development",
             "Business demand"),
            ("Net-new features · new code · integrations · capability extension · roadmap-driven",
             "Net-new features · projects · enhancements · innovations · most common Type"),
            ("Upgrade",
             "Regulatory"),
            ("Version refresh",
             "External mandate"),
            ("SAP kernel patches · vendor refresh · forced by external (Regulatory uses separate flag)",
             "PUC · SOX · Iberdrola · vendor SLA · regulator-driven deadline → triggers A3 Auto-Fixed-Date"),
        ],
    },
    # Slide 35 Automation Catalog A3 trigger — Regulatory flag is gone, now Type=Regulatory
    "Automation Catalog A3 trigger": {
        "anchor": "AUTOMATION CATALOG  ·  A1 → A6",
        "edits": [
            ("Regulatory flag = Yes",
             "Type = Regulatory"),
        ],
    },
    # Slide 25 Acceptance Criteria — "if Build/Transform" should now reference Type
    "Slide 25 Build reference": {
        "anchor": "FIELD GUIDE – ACCEPTANCE CRITERIA",
        "edits": [
            ("Definition of Done · UAT signoff · Gate 1 reference (if Build/Transform)",
             "Definition of Done · UAT signoff · Gate 1 reference (if Type=Non-Regulatory and R/G/T=Transform)"),
        ],
    },
}


# ============================================================================
# 6 new slides (same as v2 but with content-only red — no ANA badges)
# ============================================================================

def build_routing_logic(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_bar(s, prs.slide_width,
               "ROUTING LOGIC BY SIZE",
               "What goes to TEAM directly · what goes to PORTFOLIO · what bypasses everything")
    flow_y = Inches(1.30); flow_h = Inches(2.50)
    cols = [
        ("Size = S\n< 5 days", "Story (no Epic)", "TEAM BACKLOG\nDIRECT", "Skip portfolio · pulled by squad", GREEN_LIGHT),
        ("Size = M / L", "Epic + Stories", "PORTFOLIO BACKLOG\nGATE 2 next", "Sequenced by Pain + CoS", GREEN_MID),
        ("Size = XL\ncross-team", "Initiative + Epics", "ENTERPRISE DASHBOARD\nGATE 2 + tracked at VP layer", "Owns roadmap horizon · multi-quarter", GREEN_DARK),
        ("Type = Incident\n(auto from IT Now)", "Bug · CoS=Expedite", "SQUAD EXPEDITE LANE\nWIP-1", "Bypasses Triage · A2 fires", RED_EDIT),
    ]
    cw = (Inches(12.55) - 3 * Inches(0.10)) / 4
    for i, (src, what, where, why, fill) in enumerate(cols):
        cx = Inches(0.40) + i * (cw + Inches(0.10))
        round_rect(s, cx, flow_y, cw, flow_h, WHITE, radius=0.04)
        round_rect(s, cx, flow_y, cw, Inches(0.55), fill, radius=0.04)
        text(s, cx, flow_y + Inches(0.10), cw, Inches(0.35),
             src, size=11, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, cx + Inches(0.15), flow_y + Inches(0.65), cw - Inches(0.30), Inches(0.18),
             "AUTO-CONVERTS TO", size=7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, cx + Inches(0.15), flow_y + Inches(0.83), cw - Inches(0.30), Inches(0.30),
             what, size=11, bold=True, color=GREEN_DARK, font=BODY_FONT)
        text(s, cx + Inches(0.15), flow_y + Inches(1.20), cw - Inches(0.30), Inches(0.18),
             "GOES TO", size=7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, cx + Inches(0.15), flow_y + Inches(1.38), cw - Inches(0.30), Inches(0.50),
             where, size=10, bold=True, color=DARK_TEXT, font=BODY_FONT)
        text(s, cx + Inches(0.15), flow_y + Inches(1.95), cw - Inches(0.30), Inches(0.50),
             why, size=8.5, italic=True, color=GREEN_DARK, font=BODY_FONT)
    mid_y = Inches(4.05)
    round_rect(s, Inches(0.40), mid_y, Inches(12.55), Inches(1.30), GREEN_BG, radius=0.04)
    text(s, Inches(0.55), mid_y + Inches(0.08), Inches(12.2), Inches(0.22),
         "ONE LOGIC · TWO LEVELS — why size dictates routing",
         size=11, bold=True, color=GREEN_DARK, font=BODY_FONT)
    rules = [
        ("TEAM-LEVEL  ·  Squad pulls by capacity",
         "Size S items DON'T need cross-team coordination · go straight to Squad's Team Kanban after Triage. Squad pulls when WIP slot available."),
        ("PORTFOLIO-LEVEL  ·  Director sequences",
         "Size M/L/XL touch capacity slots · multiple squads · vendor scope · need Gate 2 approval for ROI + capacity reservation."),
    ]
    by = mid_y + Inches(0.32)
    for label, body in rules:
        text(s, Inches(0.55), by, Inches(4.5), Inches(0.22), label, size=9, bold=True, color=DARK_TEXT, font=BODY_FONT)
        text(s, Inches(5.10), by, Inches(7.8), Inches(0.42), body, size=9, color=DARK_TEXT, font=BODY_FONT)
        by += Inches(0.45)
    bot_y = Inches(5.55)
    round_rect(s, Inches(0.40), bot_y, Inches(12.55), Inches(1.50), WHITE, radius=0.04)
    text(s, Inches(0.55), bot_y + Inches(0.10), Inches(12.2), Inches(0.22),
         "CS vs ITBA  ·  WHERE THE LOGIC DIFFERS",
         size=11, bold=True, color=HEADER_BLUE, font=BODY_FONT)
    diff_y = bot_y + Inches(0.40)
    cs_itba_badge(s, Inches(0.55), diff_y, "CS")
    text(s, Inches(1.60), diff_y, Inches(11.0), Inches(0.50),
         "1 OVS Meter-to-Cash · 4 DVS + 1 Enabling · weights w₁=0.40 w₂=0.60 (no Safety) · Customer Reliability replaces Safety · OVS KPIs (NPS · CSAT · Reliability)",
         size=9, color=DARK_TEXT, font=BODY_FONT)
    cs_itba_badge(s, Inches(0.55), diff_y + Inches(0.50), "ITBA")
    text(s, Inches(1.60), diff_y + Inches(0.50), Inches(11.0), Inches(0.50),
         "5 teams unchanged · weights w₁=0.45 w₂=0.30 w₃=0.25 (BV-heavy · safety relevant for grid systems) · Sarath OR new PMO-Delivery role · Carlos approves at Gate 2",
         size=9, color=DARK_TEXT, font=BODY_FONT)
    cs_itba_badge(s, Inches(0.55), diff_y + Inches(1.00), "BOTH")
    text(s, Inches(1.60), diff_y + Inches(1.00), Inches(11.0), Inches(0.50),
         "Same intake form · same A1–A6 automations · same Triage cadence · same Gate 2 cadence · same Class of Service taxonomy",
         size=9, color=DARK_TEXT, font=BODY_FONT)
    footer(s, "Routing Logic by Size")
    return s


def build_cos_maintain(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_bar(s, prs.slide_width,
               "HOW WE MAINTAIN CLASS OF SERVICE",
               "The PROCESS that keeps the CoS mix healthy · % values validated each quarter")
    cycle_y = Inches(1.30); cycle_h = Inches(3.20)
    steps = [
        ("01 OBSERVE", "MEASURE flow per CoS",
         "Squad cycle-time + WIP + completion-rate per CoS lane · weekly snapshot",
         GREEN_DARK),
        ("02 ASSESS", "DETECT imbalances",
         "Are Expedites > planned cap? Fixed-Date items at risk? Standard starving? Aging tags rising?",
         GREEN_MID),
        ("03 DECIDE", "REBALANCE proposed",
         "PM + Demand Manager propose adjustments for next quarter · Director endorses at Gate 2",
         GREEN_LIGHT),
        ("04 ADJUST", "APPLY new mix",
         "WIP limits updated in Jira · automations re-tuned · KPIs re-baselined · loop back to OBSERVE",
         GREEN_CARD),
    ]
    sw = (Inches(12.55) - 3 * Inches(0.10)) / 4
    for i, (step_id, action, desc, fill) in enumerate(steps):
        sx = Inches(0.40) + i * (sw + Inches(0.10))
        round_rect(s, sx, cycle_y, sw, cycle_h, WHITE, radius=0.04)
        round_rect(s, sx, cycle_y, sw, Inches(0.65), fill, radius=0.04)
        text(s, sx, cycle_y + Inches(0.06), sw, Inches(0.22),
             step_id, size=9, bold=True, color=GREEN_PASTEL, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, sx, cycle_y + Inches(0.28), sw, Inches(0.32),
             action, size=12, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, sx + Inches(0.15), cycle_y + Inches(0.85), sw - Inches(0.30), Inches(2.0),
             desc, size=9.5, color=DARK_TEXT, font=BODY_FONT)
    own_y = Inches(4.75)
    round_rect(s, Inches(0.40), own_y, Inches(12.55), Inches(1.30), GREEN_BG, radius=0.04)
    text(s, Inches(0.55), own_y + Inches(0.10), Inches(12.2), Inches(0.22),
         "WHO OWNS WHAT  ·  THE LOOP THAT NEVER STOPS",
         size=11, bold=True, color=GREEN_DARK, font=BODY_FONT)
    rows = [
        ("OBSERVE",  "PM (per VS)",                     "Weekly · automated dashboard"),
        ("ASSESS",   "Demand Manager",                  "Weekly · Triage meeting"),
        ("DECIDE",   "Director (Carlos · Kumar) at Gate 2", "Monthly · Portfolio Committee"),
        ("ADJUST",   "Squad Lead",                      "End of quarter · BRP outcome"),
    ]
    by = own_y + Inches(0.40)
    for step, who, when in rows:
        text(s, Inches(0.55), by, Inches(2.0), Inches(0.20),
             step, size=8.5, bold=True, color=GREEN_DARK, font=BODY_FONT)
        text(s, Inches(2.65), by, Inches(5.5), Inches(0.20),
             who, size=9, color=DARK_TEXT, font=BODY_FONT)
        text(s, Inches(8.20), by, Inches(4.6), Inches(0.20),
             when, size=9, italic=True, color=GREY_LABEL, font=BODY_FONT)
        by += Inches(0.22)
    # CoS lanes + assignment rules — NO committed %
    cap_y = Inches(6.10)
    round_rect(s, Inches(0.40), cap_y, Inches(12.55), Inches(1.10), WHITE, radius=0.04)
    text(s, Inches(0.55), cap_y + Inches(0.04), Inches(12.20), Inches(0.20),
         "FOUR CoS LANES  ·  ASSIGNMENT RULES  ·  capacity % to be defined per quarter at Gate 2 (not committed yet)",
         size=8, bold=True, color=GREEN_DARK, font=BODY_FONT)
    # 4 EQUAL-WIDTH lanes (visual: rules, not %)
    lane_y = cap_y + Inches(0.28)
    lane_h = Inches(0.36)
    n = 4
    margin = 0.55
    gap = 0.10
    lane_w = (12.55 - (n - 1) * gap - 0.30) / n
    cos_data = [
        ("Standard",   GREEN_MID,   "default · no flag",                "Bulk of demand · ranked by Pain Score"),
        ("Fixed-Date", GREEN_DARK,  "Type = Regulatory  →  A3 auto",    "External deadline · scheduled backwards"),
        ("Expedite",   GREEN_CARD,  "CIT P1/P2  →  A2 auto",            "Customer outage · jump-the-queue"),
        ("Intangible", GREEN_LIGHT, "tech debt · opt-in by PM",         "Refactor · platform · enablers"),
    ]
    for i, (label, fill, rule, ex) in enumerate(cos_data):
        lx = Inches(margin + i * (lane_w + gap))
        round_rect(s, lx, lane_y, Inches(lane_w), lane_h, fill, radius=0.04)
        text(s, lx, lane_y, Inches(lane_w), lane_h,
             label, size=11, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Rule below
        text(s, lx, lane_y + Inches(0.40), Inches(lane_w), Inches(0.20),
             rule, size=8, bold=True, italic=True, color=GREEN_DARK, font=BODY_FONT,
             align=PP_ALIGN.CENTER)
        # Example below rule
        text(s, lx, lane_y + Inches(0.60), Inches(lane_w), Inches(0.22),
             ex, size=7.5, color=DARK_TEXT, font=BODY_FONT,
             align=PP_ALIGN.CENTER)
    footer(s, "How We Maintain CoS")
    return s


def build_incident_flow(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_bar(s, prs.slide_width,
               "INCIDENT FLOW  ·  IT NOW → JIRA",
               "Why Incident is NOT a Type selector · how it lands in the right Squad automatically")
    fl_y = Inches(1.30); fl_h = Inches(2.30)
    steps = [
        ("01", "INCIDENT RAISED", "User opens ticket in IT Now (ServiceNow) — only path for incidents",
         "ServiceNow form · auto-fields filled by user", HEADER_BLUE),
        ("02", "P1/P2 FILTER", "ServiceNow checks priority — only P1/P2 incidents flow to Jira (P3/P4 stay in IT Now)",
         "Webhook fires only on P1/P2 events", HEADER_BLUE),
        ("03", "API CALL", "ServiceNow webhook → Jira Cloud REST API → creates Bug ticket",
         "JSON payload includes priority + category + OPCO + system", AVANGRID_GREEN),
        ("04", "A2 + A1 RUN", "A1 routes to correct Squad · A2 sets CoS=Expedite + WIP-1",
         "Status TRIAGED skipped · already validated by ServiceNow", GREEN_DARK),
        ("05", "STATUS SYNC", "Squad works the Bug in Jira · status updates sync back to ServiceNow",
         "Bidirectional sync via API · ServiceNow remains record-of-truth for SLA", GREEN_CARD),
    ]
    sw = (Inches(12.55) - 4 * Inches(0.08)) / 5
    for i, (sid, title, desc, tech, fill) in enumerate(steps):
        sx = Inches(0.40) + i * (sw + Inches(0.08))
        round_rect(s, sx, fl_y, sw, fl_h, WHITE, radius=0.03)
        round_rect(s, sx, fl_y, sw, Inches(0.65), fill, radius=0.03)
        text(s, sx, fl_y + Inches(0.05), sw, Inches(0.22),
             sid, size=8, bold=True, color=GREEN_PASTEL, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, sx, fl_y + Inches(0.27), sw, Inches(0.32),
             title, size=11, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, sx + Inches(0.10), fl_y + Inches(0.78), sw - Inches(0.20), Inches(0.85),
             desc, size=9.5, color=DARK_TEXT, font=BODY_FONT)
        text(s, sx + Inches(0.10), fl_y + Inches(1.65), sw - Inches(0.20), Inches(0.55),
             tech, size=8, italic=True, color=GREEN_DARK, font=BODY_FONT)
    why_y = Inches(3.95)
    round_rect(s, Inches(0.40), why_y, Inches(12.55), Inches(2.10), GREEN_BG, radius=0.04)
    text(s, Inches(0.55), why_y + Inches(0.10), Inches(12.2), Inches(0.22),
         "WHY INCIDENT IS NOT IN THE INTAKE FORM",
         size=11, bold=True, color=GREEN_DARK, font=BODY_FONT)
    bullets = [
        ("Single source of truth",
         "Incidents have their own SLA framework in ServiceNow (P1 4h · P2 8h · P3 24h · P4 72h) — Jira would duplicate this."),
        ("ServiceNow is record of truth for incidents",
         "Audit · compliance · regulator-facing reports come from ServiceNow — Jira mirrors only the actionable subset (P1/P2)."),
        ("Avoid manual mis-classification",
         "Users routinely picked 'Incident' for everything urgent — flooding the incident queue. Removing the option forces correct path."),
        ("Restricted scope today",
         "ServiceNow → Jira link active for CCC + Smart Metering (CS); ITBA scope expanding next quarter post Heather validation."),
    ]
    by = why_y + Inches(0.40)
    for label, body in bullets:
        rect(s, Inches(0.65), by + Inches(0.04), Inches(0.10), Inches(0.10), AVANGRID_GREEN)
        text(s, Inches(0.85), by, Inches(4.0), Inches(0.36), label, size=10, bold=True, color=DARK_TEXT, font=BODY_FONT)
        text(s, Inches(4.95), by, Inches(8.0), Inches(0.36), body, size=9, color=DARK_TEXT, font=BODY_FONT)
        by += Inches(0.40)
    cs_itba_badge(s, Inches(11.10), Inches(6.20), "CS")
    text(s, Inches(8.00), Inches(6.20), Inches(3.0), Inches(0.22),
         "Active scope today:", size=9, italic=True, color=GREY_LABEL,
         font=BODY_FONT, align=PP_ALIGN.RIGHT)
    cs_itba_badge(s, Inches(12.10), Inches(6.20), "ITBA")
    text(s, Inches(8.00), Inches(6.45), Inches(5.0), Inches(0.22),
         "ITBA expansion: post-validation · later quarter",
         size=8, italic=True, color=GREY_LABEL, font=BODY_FONT, align=PP_ALIGN.RIGHT)
    footer(s, "Incident Flow")
    return s


def build_existing_project(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_bar(s, prs.slide_width,
               "DEMAND ON EXISTING PROJECT",
               "Scope-add to a live Initiative · skip Funnel/Reviewing/Analyzing · Story shortcut under existing Epic")
    top_y = Inches(1.20)
    round_rect(s, Inches(0.40), top_y, Inches(12.55), Inches(0.65), GREEN_BG, radius=0.04)
    text(s, Inches(0.55), top_y + Inches(0.10), Inches(12.2), Inches(0.22),
         "WHEN  ·  Initiative is in IMPLEMENTING · new related ask comes from same Sponsor / OPCO / business area",
         size=11, bold=True, color=GREEN_DARK, font=BODY_FONT)
    text(s, Inches(0.55), top_y + Inches(0.36), Inches(12.2), Inches(0.22),
         "EXAMPLE  ·  CRM unified-record Initiative live · Sponsor asks to add 'phone number masking for CCAR call recordings'",
         size=9, italic=True, color=DARK_TEXT, font=BODY_FONT)
    fl_y = Inches(2.05); fl_h = Inches(3.10)
    steps = [
        ("01", "LINK",
         "Requester knows parent Initiative key  ·  parent is in IMPLEMENTING",
         "Requester opens NEW Demand · types Parent Initiative key · A1 detects link",
         "Form auto-fills VS · OPCO · Sponsor from parent", AVANGRID_GREEN),
        ("02", "BYPASS",
         "A1 confirmed valid link  ·  same VS / Sponsor / OPCO as parent",
         "Skips Funnel · Reviewing · Analyzing · Gate 2 — already approved as part of parent",
         "No Pain Score recompute · keeps parent's CoS", GREEN_DARK),
        ("03", "SQUAD APPROVE",
         "Bypass triggered  ·  Squad PM gets notification of incoming scope-add",
         "Squad PM accepts (or rejects) the scope-add — protects WIP and timeline",
         "PM evaluates impact · estimates delta · raises if blocking", GREEN_MID),
        ("04", "STORY",
         "Squad PM ACCEPTED in step 03  ·  scope-add fits within current sprint",
         "Becomes a STORY under the parent Initiative's Epic · pulled into Squad Team Kanban",
         "Story status syncs with parent Initiative · cycle-time tracked separately", GREEN_LIGHT),
    ]
    sw = (Inches(12.55) - 3 * Inches(0.10)) / 4
    for i, (sid, title, condition, desc, tech, fill) in enumerate(steps):
        sx = Inches(0.40) + i * (sw + Inches(0.10))
        round_rect(s, sx, fl_y, sw, fl_h, WHITE, radius=0.04)
        round_rect(s, sx, fl_y, sw, Inches(0.65), fill, radius=0.04)
        text(s, sx, fl_y + Inches(0.05), sw, Inches(0.22),
             sid, size=8, bold=True, color=GREEN_PASTEL, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, sx, fl_y + Inches(0.27), sw, Inches(0.32),
             title, size=14, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # CONDITION label + text — explains WHEN this step happens
        text(s, sx + Inches(0.12), fl_y + Inches(0.75), sw - Inches(0.24), Inches(0.18),
             "CONDITION", size=7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, sx + Inches(0.12), fl_y + Inches(0.92), sw - Inches(0.24), Inches(0.55),
             condition, size=8, italic=True, color=GREEN_DARK, font=BODY_FONT)
        # WHAT HAPPENS · description
        text(s, sx + Inches(0.12), fl_y + Inches(1.55), sw - Inches(0.24), Inches(0.18),
             "WHAT HAPPENS", size=7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, sx + Inches(0.12), fl_y + Inches(1.72), sw - Inches(0.24), Inches(0.85),
             desc, size=9, color=DARK_TEXT, font=BODY_FONT)
        # TECHNICAL DETAIL · footer italic
        text(s, sx + Inches(0.12), fl_y + Inches(2.65), sw - Inches(0.24), Inches(0.40),
             tech, size=8, italic=True, color=GREEN_DARK, font=BODY_FONT)
    # STATUS-BASED DECISION MATRIX — when each path applies
    matrix_y = Inches(5.30)
    round_rect(s, Inches(0.40), matrix_y, Inches(12.55), Inches(1.05), WHITE, radius=0.04)
    round_rect(s, Inches(0.40), matrix_y, Inches(12.55), Inches(0.30), AVANGRID_GREEN, radius=0.04)
    text(s, Inches(0.55), matrix_y + Inches(0.04), Inches(12.20), Inches(0.22),
         "DECISION MATRIX  ·  When each path applies based on parent Initiative status",
         size=10, bold=True, color=WHITE, font=BODY_FONT)
    # Matrix rows
    rows = [
        ("Funnel / Reviewing / Analyzing",  "Update the EXISTING Demand card directly (not yet committed) · no new card needed"),
        ("Backlog (post-Gate 2 GO · awaiting capacity)",  "Update parent Initiative scope · re-evaluate at next Gate 2 if scope grows"),
        ("Implementing (active in squad)",  "USE SHORTCUT (this slide) → Story under existing Epic · Squad PM approves"),
        ("Done (completed)",                 "Open NEW Demand through Funnel · normal flow · may become new Initiative"),
    ]
    by = matrix_y + Inches(0.32)
    for status, action in rows:
        text(s, Inches(0.55), by, Inches(4.50), Inches(0.18),
             status, size=8.5, bold=True, color=GREEN_DARK, font=BODY_FONT)
        text(s, Inches(5.10), by, Inches(7.80), Inches(0.18),
             action, size=8.5, color=DARK_TEXT, font=BODY_FONT)
        by += Inches(0.18)

    # GUARDRAILS section (compact)
    grd_y = Inches(6.42)
    round_rect(s, Inches(0.40), grd_y, Inches(12.55), Inches(0.70), GREEN_BG, radius=0.04)
    text(s, Inches(0.55), grd_y + Inches(0.06), Inches(12.2), Inches(0.18),
         "GUARDRAILS  ·  EVEN WHEN PARENT IS IMPLEMENTING, SHORTCUT MAY BE REJECTED",
         size=9, bold=True, color=GREEN_DARK, font=BODY_FONT)
    rules = [
        "▸  Scope > 5 days OR > 1 squad sprint → open new Demand through Funnel    ▸  Different Sponsor / OPCO / VS than parent → normal flow",
        "▸  Squad PM has full authority to reject if WIP / timeline at risk · Sponsor escalates to Director if needed",
    ]
    by = grd_y + Inches(0.26)
    for r in rules:
        text(s, Inches(0.55), by, Inches(12.2), Inches(0.20), r, size=8.5, color=DARK_TEXT, font=BODY_FONT)
        by += Inches(0.20)
    footer(s, "Demand on Existing Initiative")
    return s


def build_digital_coe(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_bar(s, prs.slide_width,
               "DIGITAL CoE  ×  DIGITAL PRODUCTS",
               "How the two organizations connect · same Jira intake · separate governance · joint OKR sync")
    col_y = Inches(1.20); col_h = Inches(5.45); col_w = Inches(6.18); col_gap = Inches(0.18)
    round_rect(s, Inches(0.40), col_y, col_w, col_h, WHITE, radius=0.04)
    round_rect(s, Inches(0.40), col_y, col_w, Inches(0.65), AVANGRID_GREEN, radius=0.04)
    text(s, Inches(0.55), col_y + Inches(0.12), col_w - Inches(0.30), Inches(0.32),
         "1.  DEMAND CHANNEL  ·  via Jira", size=14, bold=True, color=WHITE, font=BODY_FONT)
    text(s, Inches(0.55), col_y + Inches(0.40), col_w - Inches(0.30), Inches(0.22),
         "Digital CoE submits demand the same way as any business area",
         size=9, italic=True, color=GREEN_PASTEL, font=BODY_FONT)
    bullets_left = [
        ("ENTRY POINT", "Digital CoE opens a Jira Demand via the SAME intake form  ·  field 'Sponsor area' = 'Digital CoE'"),
        ("ROUTING", "A1 routes to the correct VS based on impacted product (CRM · Smart Metering · Genesys · etc.)"),
        ("PRIORITIZATION", "Same Pain Score formula applies · Digital CoE OKRs feed BV component"),
        ("GATE 2 SEAT", "Digital CoE Champion participates as VOTING MEMBER for items where Digital CoE is sponsor"),
        ("CARD FLOW", "Same lifecycle: Funnel → Reviewing → Analyzing → Backlog → Implementing → Done"),
    ]
    by = col_y + Inches(0.85)
    for label, body in bullets_left:
        text(s, Inches(0.55), by, Inches(1.65), Inches(0.20),
             label, size=8, bold=True, color=GREEN_DARK, font=BODY_FONT)
        text(s, Inches(2.25), by, col_w - Inches(2.00), Inches(0.65),
             body, size=9, color=DARK_TEXT, font=BODY_FONT)
        by += Inches(0.72)
    rx = Inches(0.40) + col_w + col_gap
    round_rect(s, rx, col_y, col_w, col_h, WHITE, radius=0.04)
    round_rect(s, rx, col_y, col_w, Inches(0.65), HEADER_BLUE, radius=0.04)
    text(s, rx + Inches(0.15), col_y + Inches(0.12), col_w - Inches(0.30), Inches(0.32),
         "2.  STRATEGIC ALIGNMENT  ·  via OKR sync", size=14, bold=True, color=WHITE, font=BODY_FONT)
    text(s, rx + Inches(0.15), col_y + Inches(0.40), col_w - Inches(0.30), Inches(0.22),
         "Existing BRP outputs flow IN as input to our monthly Gate 2 · we don't run a quarterly event",
         size=9, italic=True, color=GREEN_PASTEL, font=BODY_FONT)
    bullets_right = [
        ("DIRECTION",       "BRPs (existing) → us  ·  we RECEIVE outputs · refined priorities · OKRs · capacity decisions"),
        ("OWNERS",          "Kumar (CS) + Carlos (ITBA) + Vitor (VP)  ·  liaisons to Digital CoE BRP and P&T BRP respectively"),
        ("CONTENT IN",      "Refined OKRs · regulatory mandates · cross-org dependencies · headcount/capacity decisions from BRPs"),
        ("OUTCOME",         "BRP outputs feed our monthly Gate 2 · Gate 2 capacity & priority decisions incorporate BRP inputs"),
        ("NO OWN BRP",      "We don't run a Digital Products BRP · couple to existing P&T BRP + Digital CoE BRP via the liaisons"),
    ]
    by = col_y + Inches(0.85)
    for label, body in bullets_right:
        text(s, rx + Inches(0.15), by, Inches(1.65), Inches(0.20),
             label, size=8, bold=True, color=HEADER_BLUE, font=BODY_FONT)
        text(s, rx + Inches(1.85), by, col_w - Inches(2.00), Inches(0.65),
             body, size=9, color=DARK_TEXT, font=BODY_FONT)
        by += Inches(0.72)
    footer(s, "Digital CoE × Digital Products")
    return s


def build_cit_field_guide(prs):
    """NEW Field Guide slide dedicated to Customer Impact Tier (CIT)."""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_bar(s, prs.slide_width,
               "FIELD GUIDE – CUSTOMER IMPACT TIER",
               "Categorical severity flag · CS only · drives A2 Auto-Expedite · NOT a Pain Score input")

    # Top metadata strip (mimics other Field Guide slides)
    meta_y = Inches(1.20)
    round_rect(s, Inches(0.40), meta_y, Inches(12.55), Inches(0.55), GREEN_DARK, radius=0.04)
    text(s, Inches(0.55), meta_y + Inches(0.04), Inches(8.0), Inches(0.20),
         "TYPE  ·  Single-select · 4 options · CS only · optional at submit",
         size=8, color=GREEN_PASTEL, font=BODY_FONT)
    text(s, Inches(0.55), meta_y + Inches(0.22), Inches(8.0), Inches(0.30),
         "Customer Impact Tier",
         size=18, bold=True, color=WHITE, font=BODY_FONT)
    # Right-side role badge — standard green (matches other Field Guide slides)
    round_rect(s, Inches(8.60), meta_y + Inches(0.13), Inches(4.30), Inches(0.30), GREEN_DARK, radius=0.30)
    text(s, Inches(8.60), meta_y + Inches(0.13), Inches(4.30), Inches(0.30),
         "● REQUESTER (if known)  /  PM at Reviewing",
         size=8, bold=True, color=WHITE, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 4 vertical cards for P1/P2/P3/P4 — green gradient (no red, no olive)
    card_y = Inches(1.95)
    card_h = Inches(3.85)
    card_w = (Inches(12.55) - 3 * Inches(0.10)) / 4
    cards = [
        ("P1", "CRITICAL", "> 100k customers OR full OPCO outage", GREEN_DARK, [
            ("WHEN TO USE",  "Production outage affecting full customer base of an OPCO · service unavailable · revenue impact > $100k/day"),
            ("AUTO-TRIGGER", "A2 Auto-Expedite · CoS=Expedite · WIP-1 · 4-hour SLA"),
            ("EXAMPLES",     "Billing engine down · IVR fully offline · Iberdrola month-close blocker"),
        ]),
        ("P2", "HIGH", "> 10k customers · degraded service", GREEN_CARD, [
            ("WHEN TO USE",  "Significant degradation but service continues · workarounds available · regulatory deadline imminent"),
            ("AUTO-TRIGGER", "A2 Auto-Expedite · CoS=Expedite · WIP-1 · 8-hour SLA"),
            ("EXAMPLES",     "CRM partial outage · Genesys queue routing failure · payment gateway slow"),
        ]),
        ("P3", "MEDIUM", "single-OPCO non-critical issue", GREEN_MID, [
            ("WHEN TO USE",  "One OPCO affected · workarounds in place · low complaint volume · not customer-facing severity"),
            ("AUTO-TRIGGER", "No automation · Standard CoS · 24-hour SLA · normal triage flow"),
            ("EXAMPLES",     "Localized billing display issue · UI glitch on a single OPCO · dashboard slow"),
        ]),
        ("P4", "LOW", "internal-only · no customer impact", GREEN_LIGHT, [
            ("WHEN TO USE",  "Internal users only · agent-side or back-office tooling · no customer-facing implications"),
            ("AUTO-TRIGGER", "No automation · Standard or Intangible CoS · 72-hour SLA"),
            ("EXAMPLES",     "Internal report formatting · agent dashboard tweak · documentation update"),
        ]),
    ]
    for i, (level, severity, summary, fill, blocks) in enumerate(cards):
        cx = Inches(0.40) + i * (card_w + Inches(0.10))
        round_rect(s, cx, card_y, card_w, card_h, WHITE, radius=0.04)
        round_rect(s, cx, card_y, card_w, Inches(0.85), fill, radius=0.04)
        text(s, cx, card_y + Inches(0.06), card_w, Inches(0.32),
             level, size=22, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER)
        text(s, cx, card_y + Inches(0.40), card_w, Inches(0.22),
             severity, size=10, bold=True, color=GREEN_PASTEL, font=BODY_FONT,
             align=PP_ALIGN.CENTER)
        text(s, cx + Inches(0.08), card_y + Inches(0.62), card_w - Inches(0.16), Inches(0.22),
             summary, size=8, italic=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER)

        by = card_y + Inches(1.00)
        for label, content in blocks:
            text(s, cx + Inches(0.12), by, card_w - Inches(0.24), Inches(0.18),
                 label, size=8, bold=True, color=GREY_LABEL, font=BODY_FONT)
            text(s, cx + Inches(0.12), by + Inches(0.20), card_w - Inches(0.24), Inches(0.70),
                 content, size=8.5, color=DARK_TEXT, font=BODY_FONT)
            by += Inches(0.92)

    # HOW IT AFFECTS CALCULATION section — standard green palette (matches other Field Guide slides)
    eff_y = Inches(5.95)
    round_rect(s, Inches(0.40), eff_y, Inches(12.55), Inches(0.95), GREEN_BG, radius=0.04)
    text(s, Inches(0.55), eff_y + Inches(0.08), Inches(12.20), Inches(0.22),
         "HOW IT AFFECTS THE CALCULATION",
         size=11, bold=True, color=GREEN_DARK, font=BODY_FONT)
    text(s, Inches(0.55), eff_y + Inches(0.30), Inches(12.20), Inches(0.30),
         "▸  CIT does NOT enter the Pain Score formula  —  it is a CATEGORICAL FLAG, not a numeric weight",
         size=10, color=DARK_TEXT, font=BODY_FONT)
    text(s, Inches(0.55), eff_y + Inches(0.52), Inches(12.20), Inches(0.30),
         "▸  P1 / P2  →  triggers A2 Auto-Expedite  →  forces CoS = Expedite + WIP-1  →  jumps the queue regardless of Pain Score",
         size=10, color=DARK_TEXT, font=BODY_FONT)
    text(s, Inches(0.55), eff_y + Inches(0.74), Inches(12.20), Inches(0.20),
         "▸  P3 / P4  →  no automation  →  follows Pain Score ranking + CoS lane normally",
         size=10, color=DARK_TEXT, font=BODY_FONT)

    # Mandatory rule footer — standard dark text
    rule_y = Inches(7.00)
    text(s, Inches(0.55), rule_y, Inches(12.20), Inches(0.20),
         "MANDATORY?  ·  Optional at submit (Requester fills if known) · Reviewing PM may add later  ·  empty = Standard CoS default",
         size=9, italic=True, bold=True, color=GREEN_DARK, font=BODY_FONT)

    footer(s, "Customer Impact Tier · NEW")
    return s


def build_post_gate2_flow(prs):
    """NEW slide showing what happens after Gate 2: GO / NO-GO / DEFER paths."""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_bar(s, prs.slide_width,
               "POST-GATE 2  ·  WHAT HAPPENS NEXT",
               "Three paths from Gate 2 decision · GO → execution · NO-GO → archive · DEFER → next month")

    # Top: Gate 2 decision node
    g2_y = Inches(1.20)
    round_rect(s, Inches(5.20), g2_y, Inches(2.95), Inches(0.85), GREEN_DARK, radius=0.04)
    text(s, Inches(5.20), g2_y + Inches(0.06), Inches(2.95), Inches(0.20),
         "GATE 2 · DECISION", size=8, bold=True, color=GREEN_PASTEL,
         font=BODY_FONT, align=PP_ALIGN.CENTER)
    text(s, Inches(5.20), g2_y + Inches(0.28), Inches(2.95), Inches(0.30),
         "Director (Carlos · Kumar)",
         size=12, bold=True, color=WHITE, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(5.20), g2_y + Inches(0.58), Inches(2.95), Inches(0.22),
         "Monthly · 90 min", size=9, italic=True, color=GREEN_PASTEL,
         font=BODY_FONT, align=PP_ALIGN.CENTER)

    # 3 outcome paths
    path_y = Inches(2.30)
    path_h = Inches(4.50)
    path_w = (Inches(12.55) - 2 * Inches(0.20)) / 3
    paths = [
        ("GO", "Initiative approved",
         AVANGRID_GREEN,
         [
            ("01 STATUS", "Issue → 'Ready For Execution' · capacity slot reserved · Backlog rank assigned"),
            ("02 NOTIFY", "@Sponsor + @Squad Lead · Layer 0 VP Dashboard updated · published in Confluence"),
            ("03 HANDOFF", "PM (Portfolio) hands off to Tech Lead + squad · DoR re-checked at squad level"),
            ("04 EXECUTION", "Enters Team Kanban · Phase 0 Discovery → Phase 1 Materialization → Phase 2 Creation → Phase 3 Ongoing → Phase 4 Maintenance"),
            ("05 KPIs", "Cycle time · throughput · quality · NPS/CSAT delta · OVS KPI panel monitors outcomes"),
         ]),
        ("NO-GO", "Initiative rejected",
         GREEN_CARD,
         [
            ("01 STATUS", "Issue → 'Rejected' · removed from active backlog · archived for traceability"),
            ("02 RATIONALE", "Director records reason: capacity · ROI · strategic fit · technical feasibility"),
            ("03 NOTIFY", "@Sponsor receives written rationale · Confluence post · learning captured"),
            ("04 RE-SUBMIT", "Sponsor may re-submit with different framing/scope · counts as new Demand · full Triage"),
            ("05 LEARNING", "Quarterly review of NO-GO patterns to refine intake quality + scoring rubrics"),
         ]),
        ("DEFER", "Decision postponed",
         GREEN_MID,
         [
            ("01 STATUS", "Issue → 'Deferred' · stays in portfolio backlog · timestamp + reason captured"),
            ("02 REASON", "Insufficient info · capacity tight · waiting on dependency · seasonal timing"),
            ("03 NOTIFY", "@Sponsor + @PM · expected return date communicated"),
            ("04 RE-EVALUATE", "Auto-surfaces at next monthly Gate 2 · PM may add missing info between gates"),
            ("05 ESCALATION", "If deferred 3+ times → escalate to BRP for cross-portfolio decision"),
         ]),
    ]
    for i, (label, sub, fill, blocks) in enumerate(paths):
        cx = Inches(0.40) + i * (path_w + Inches(0.20))
        round_rect(s, cx, path_y, path_w, path_h, WHITE, radius=0.03)
        round_rect(s, cx, path_y, path_w, Inches(0.80), fill, radius=0.03)
        text(s, cx, path_y + Inches(0.10), path_w, Inches(0.32),
             label, size=20, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, cx, path_y + Inches(0.46), path_w, Inches(0.22),
             sub, size=10, italic=True, color=GREEN_PASTEL, font=BODY_FONT,
             align=PP_ALIGN.CENTER)
        by = path_y + Inches(0.95)
        for label_b, content in blocks:
            text(s, cx + Inches(0.15), by, path_w - Inches(0.30), Inches(0.20),
                 label_b, size=8, bold=True, color=GREEN_DARK, font=BODY_FONT)
            text(s, cx + Inches(0.15), by + Inches(0.18), path_w - Inches(0.30), Inches(0.45),
                 content, size=8, color=DARK_TEXT, font=BODY_FONT)
            by += Inches(0.68)

    # Bottom callout
    note_y = Inches(6.95)
    round_rect(s, Inches(0.40), note_y, Inches(12.55), Inches(0.20), GREEN_BG, radius=0.04)
    text(s, Inches(0.55), note_y, Inches(12.20), Inches(0.20),
         "All Gate 2 decisions are LOGGED in Jira · auditable · feed quarterly portfolio review at the BRPs (P&T + Digital CoE)",
         size=8, italic=True, bold=True, color=GREEN_DARK, font=BODY_FONT,
         anchor=MSO_ANCHOR.MIDDLE)

    footer(s, "Post-Gate 2 Flow")
    return s


def build_preparation_planning(prs):
    """NEW slide for Preparation & Planning · 2 levels: Portfolio practices + Project planning."""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_bar(s, prs.slide_width,
               "PREPARATION & PLANNING",
               "Level 1 · Portfolio practices  |  Level 2 · Project planning  ·  what happens between Gate 2 GO and execution start")

    # Two-column layout
    col_y = Inches(1.20)
    col_h = Inches(5.85)
    col_w = (Inches(12.55) - Inches(0.20)) / 2

    # LEVEL 1 — Portfolio practices
    round_rect(s, Inches(0.40), col_y, col_w, col_h, WHITE, radius=0.04)
    round_rect(s, Inches(0.40), col_y, col_w, Inches(0.85), AVANGRID_GREEN, radius=0.04)
    text(s, Inches(0.55), col_y + Inches(0.10), col_w - Inches(0.30), Inches(0.30),
         "LEVEL 1  ·  PORTFOLIO PRACTICES",
         size=14, bold=True, color=WHITE, font=BODY_FONT)
    text(s, Inches(0.55), col_y + Inches(0.40), col_w - Inches(0.30), Inches(0.22),
         "Director-level · monthly Gate 2 cadence · informed by external BRP outputs",
         size=9, italic=True, color=GREEN_PASTEL, font=BODY_FONT)
    text(s, Inches(0.55), col_y + Inches(0.62), col_w - Inches(0.30), Inches(0.22),
         "OWNER · Director (Carlos · Kumar)  ·  cadence: monthly Gate 2  ·  no own quarterly event",
         size=8, bold=True, color=WHITE, font=BODY_FONT)
    blocks_l1 = [
        ("CAPACITY PLANNING",
         "Squad capacity FIXED quarterly (squads × FTE × velocity) · % allocation per CoS DECIDED monthly at Gate 2 · informed by P&T BRP + Digital CoE BRP outputs"),
        ("PORTFOLIO BACKLOG",
         "Long-term roadmap · ranked by Pain Score + CoS + OPCO multiplier · re-ranked at every monthly Gate 2"),
        ("APPLICATION LICENSES",
         "License inventory · renewal calendar · cost optimization · vendor management · auto-flagged at intake"),
        ("CROSS-VS DEPENDENCIES",
         "Identified from P&T BRP + Digital CoE BRP OUTPUTS (we receive, not produce) · feed into our Gate 2 decisions"),
        ("OKRs CASCADE",
         "Corporate OKRs → Director OKRs → VS OKRs · arrive via existing BRPs · drives BV scoring + Gate 2 prioritization"),
    ]
    by = col_y + Inches(0.95)
    for label_b, content in blocks_l1:
        text(s, Inches(0.55), by, col_w - Inches(0.30), Inches(0.20),
             label_b, size=9, bold=True, color=GREEN_DARK, font=BODY_FONT)
        text(s, Inches(0.55), by + Inches(0.20), col_w - Inches(0.30), Inches(0.65),
             content, size=9, color=DARK_TEXT, font=BODY_FONT)
        by += Inches(0.92)

    # LEVEL 2 — Value Stream planning (per VS · within Product context)
    rx = Inches(0.40) + col_w + Inches(0.20)
    round_rect(s, rx, col_y, col_w, col_h, WHITE, radius=0.04)
    round_rect(s, rx, col_y, col_w, Inches(0.85), GREEN_DARK, radius=0.04)
    text(s, rx + Inches(0.15), col_y + Inches(0.10), col_w - Inches(0.30), Inches(0.30),
         "LEVEL 2  ·  VALUE STREAM PLANNING",
         size=14, bold=True, color=WHITE, font=BODY_FONT)
    text(s, rx + Inches(0.15), col_y + Inches(0.40), col_w - Inches(0.30), Inches(0.22),
         "PM-level · post-Gate 2 GO · per Initiative within the Value Stream",
         size=9, italic=True, color=GREEN_PASTEL, font=BODY_FONT)
    text(s, rx + Inches(0.15), col_y + Inches(0.62), col_w - Inches(0.30), Inches(0.22),
         "OWNER · PM (per VS)  ·  cadence: per Initiative · continuous within VS",
         size=8, bold=True, color=WHITE, font=BODY_FONT)
    blocks_l2 = [
        ("KICK-OFF",
         "Within 1 week of Gate 2 GO · PM + Tech Lead + Sponsor · scope confirmed · success metrics agreed in VS context"),
        ("DEEP DISCOVERY",
         "Phase 0 deepens · users / personas validated · solution shape · risks · architecture review (run by domain product team / vendor · NOT Digital Products)"),
        ("SIZING & MATERIALIZATION",
         "Phase 1 · t-shirt confirmed · roadmap · budget · vendor model finalized · feeds VS roadmap"),
        ("DEFINITION OF READY",
         "Squad-level DoR check before pulling to Phase 2/3 (Delivery · Digital Products) · DoR mismatch → back to Materialization"),
        ("EXISTING INITIATIVES",
         "Scope-add via 'Demand on Existing Project' shortcut · Squad PM evaluates impact · sub-task linked to parent Initiative"),
    ]
    by = col_y + Inches(0.95)
    for label_b, content in blocks_l2:
        text(s, rx + Inches(0.15), by, col_w - Inches(0.30), Inches(0.20),
             label_b, size=9, bold=True, color=GREEN_DARK, font=BODY_FONT)
        text(s, rx + Inches(0.15), by + Inches(0.20), col_w - Inches(0.30), Inches(0.65),
             content, size=9, color=DARK_TEXT, font=BODY_FONT)
        by += Inches(0.92)

    footer(s, "Preparation & Planning")
    return s


def build_form_view_per_stage(prs, stage_id, stage_name, owner, active_sections, header_color):
    """Build a form mockup slide tailored to a specific stage. Matches original TO-BE DEMAND FLOW layout exactly."""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    # Title matching original slide 20 (Avangrid green title + dark subtitle)
    text(s, Inches(0.40), Inches(0.10), Inches(11.0), Inches(0.42),
         f"TO-BE DEMAND FLOW  ·  {stage_name}",
         size=22, bold=True, color=AVANGRID_GREEN, font=TITLE_FONT,
         anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(0.40), Inches(0.50), Inches(11.0), Inches(0.32),
         f"OWNER · {owner}  ·  previous fields locked (gray) · future fields not visible yet",
         size=11, bold=True, color=DARK_TEXT, font=BODY_FONT)
    text(s, Inches(11.4), Inches(0.20), Inches(1.7), Inches(0.4),
         "NTT  DATA", size=14, bold=True, color=HEADER_BLUE,
         font=BODY_FONT, align=PP_ALIGN.RIGHT)
    # Jira green banner
    bar_y = Inches(0.95)
    rect(s, Inches(0.40), bar_y, Inches(12.55), Inches(0.32), AVANGRID_GREEN)
    text(s, Inches(0.55), bar_y, Inches(8), Inches(0.32),
         "▦  Jira Software  /  Avangrid · Digital Products  /  Create Issue",
         size=10, bold=True, color=WHITE, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
    # Type indicator + ticket ID (Reviewing+ shows actual values, Funnel shows placeholder)
    type_strip_y = Inches(1.30)
    rect(s, Inches(0.40), type_strip_y, Inches(12.55), Inches(0.30), GREY_BG)
    sc_state = active_sections.get("class")
    if sc_state == "active":
        text(s, Inches(0.55), type_strip_y, Inches(4.0), Inches(0.30),
             "▼  Production Support · Regulatory · Non-Regulatory",
             size=9, bold=True, color=AVANGRID_GREEN, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
    elif sc_state == "locked":
        text(s, Inches(0.55), type_strip_y, Inches(4.0), Inches(0.30),
             "Non-Regulatory  (locked)",
             size=9, italic=True, color=RGBColor(0x66, 0x66, 0x66), font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(8.0), type_strip_y, Inches(4.5), Inches(0.30),
         "DGTL-PROD-NNNN",
         size=9, italic=True, color=GREY_LABEL, font=BODY_FONT,
         anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT)

    # === Helper: render a field box · skip rendering if state == 'hidden' ===
    def field(x, y, w, h, label, value, state):
        if state == "hidden":
            return  # don't render at all — the field is not visible at this stage
        if state == "active":
            label_color = GREEN_DARK
            bg = WHITE
            value_color = DARK_TEXT
        elif state == "locked":
            label_color = RGBColor(0x99, 0x99, 0x99)
            bg = RGBColor(0xEE, 0xEE, 0xEE)
            value_color = RGBColor(0x66, 0x66, 0x66)
        else:
            return
        # Label
        text(s, Inches(x), Inches(y), Inches(w), Inches(0.18),
             label, size=8, bold=True, color=label_color, font=BODY_FONT)
        # Field box
        round_rect(s, Inches(x), Inches(y + 0.18), Inches(w), Inches(h - 0.18), bg, radius=0.04)
        # Value
        text(s, Inches(x + 0.08), Inches(y + 0.22), Inches(w - 0.16), Inches(h - 0.26),
             value, size=8, italic=(state == "locked"), color=value_color, font=BODY_FONT)

    def section_header(x, y, w, label, state):
        if state == "hidden": return
        text(s, Inches(x), Inches(y), Inches(w), Inches(0.20),
             label, size=11, bold=True,
             color=AVANGRID_GREEN if state == "active" else RGBColor(0x99, 0x99, 0x99),
             font=BODY_FONT)

    # === Define field layout (consistent across all 3 stage views) ===
    # WHAT section (left side)
    sec_y = Inches(1.85)
    text(s, Inches(0.40), sec_y, Inches(6.0), Inches(0.20),
         "WHAT",
         size=11, bold=True, color=AVANGRID_GREEN if active_sections.get("what") == "active" else RGBColor(0xAA, 0xAA, 0xAA),
         font=BODY_FONT)

    sw = active_sections.get("what")
    field(0.40, 2.10, 6.00, 0.55, "Title *",
          'e.g. "Reduce billing-correction cycle from 14 to 5 days"' if sw == "active"
          else "Reduce billing-correction cycle from 14 to 5 days" if sw == "locked"
          else "(filled at Funnel by Requester)", sw)
    field(0.40, 2.70, 6.00, 0.95, "Description *",
          "Context · current pain · who is affected · what blocks today" if sw == "active"
          else "Customer Care reps reconcile billing errors manually · 240 hrs/cycle · 8% inbound calls" if sw == "locked"
          else "(filled at Funnel by Requester)", sw)
    field(0.40, 3.70, 2.90, 0.55, "Business Value *",
          "$ impact · % delta · NPS · CSAT · time saved" if sw == "active"
          else "$285k/yr labor · NPS +5 · 240 hrs/cycle" if sw == "locked"
          else "(filled at Funnel by Requester)", sw)
    field(3.50, 3.70, 2.90, 0.55, "Acceptance Criteria *",
          "list of measurable conditions · UAT testable" if sw == "active"
          else "Cycle ≤ 5 days · UAT signed by Sponsor" if sw == "locked"
          else "(filled at Funnel by Requester)", sw)

    # ROUTING section
    text(s, Inches(0.40), Inches(4.35), Inches(6.0), Inches(0.20),
         "ROUTING",
         size=11, bold=True,
         color=AVANGRID_GREEN if active_sections.get("routing") == "active" else RGBColor(0xAA, 0xAA, 0xAA),
         font=BODY_FONT)
    sr = active_sections.get("routing")
    field(0.40, 4.60, 6.00, 0.45, "Value Stream destination *",
          "▼ Select Value Stream / Team..." if sr == "active"
          else "Billing Operations (CS · Kumar)" if sr == "locked"
          else "(filled at Funnel)", sr)
    field(0.40, 5.15, 1.95, 0.45, "OPCO *",
          "▼ multi-select" if sr == "active"
          else "UIL · CMP · NY (3)" if sr == "locked"
          else "(filled at Funnel)", sr)
    field(2.45, 5.15, 1.95, 0.45, "Sponsor *",
          "@ user picker" if sr == "active"
          else "Heather LaJoie" if sr == "locked"
          else "(filled at Funnel)", sr)
    field(4.50, 5.15, 1.90, 0.45, "Customer Impact Tier",
          "P1 / P2 / P3 / P4" if sr == "active"
          else "P2 (high)" if sr == "locked"
          else "—", sr)

    # CLASSIFICATION section (right side)
    text(s, Inches(6.65), Inches(1.85), Inches(6.30), Inches(0.20),
         "CLASSIFICATION",
         size=11, bold=True,
         color=AVANGRID_GREEN if active_sections.get("class") == "active" else RGBColor(0xAA, 0xAA, 0xAA),
         font=BODY_FONT)
    sc = active_sections.get("class")
    field(6.65, 2.10, 3.10, 0.45, "Type",
          "▼ Production Support / Regulatory / Non-Reg" if sc == "active"
          else "Non-Regulatory" if sc == "locked"
          else "(filled at Reviewing)", sc)
    field(9.85, 2.10, 3.10, 0.45, "Size",
          "▼ S · M · L  (XL → PPM Now)" if sc == "active"
          else "M" if sc == "locked"
          else "(filled at Reviewing)", sc)
    field(6.65, 2.65, 3.10, 0.45, "Class of Service",
          "Standard / Fixed-Date / Expedite / Intangible" if sc == "active"
          else "Standard" if sc == "locked"
          else "(filled at Reviewing)", sc)
    field(9.85, 2.65, 3.10, 0.45, "Run / Grow / Transform",
          "▼ Run · Grow · Transform" if sc == "active"
          else "Grow" if sc == "locked"
          else "(filled at Reviewing)", sc)
    field(6.65, 3.20, 3.10, 0.45, "Due Date",
          "📅 YYYY-MM-DD" if sc == "active"
          else "2026-Q3" if sc == "locked"
          else "—", sc)

    # SCORES (1-5) section — only active at Analyzing
    text(s, Inches(6.65), Inches(3.85), Inches(6.30), Inches(0.20),
         "SCORES (1–5)  ·  PM scores BV/UX  ·  Architect scores Safety/Effort",
         size=11, bold=True,
         color=AVANGRID_GREEN if active_sections.get("scores") == "active" else RGBColor(0xAA, 0xAA, 0xAA),
         font=BODY_FONT)
    ss = active_sections.get("scores")
    field(6.65, 4.10, 1.50, 0.45, "BV (1–5)",
          "▼ 1-5" if ss == "active" else "4" if ss == "locked" else "—", ss)
    field(8.25, 4.10, 1.50, 0.45, "UX (1–5)",
          "▼ 1-5" if ss == "active" else "5" if ss == "locked" else "—", ss)
    field(9.85, 4.10, 1.50, 0.45, "Safety (1–5)  · ITBA",
          "▼ 1-5" if ss == "active" else "n/a CS" if ss == "locked" else "—", ss)
    field(11.45, 4.10, 1.50, 0.45, "Effort (1–5)",
          "▼ S/M/L" if ss == "active" else "M" if ss == "locked" else "—", ss)

    # PAIN SCORE box
    ps_y = Inches(4.75)
    ps_state = active_sections.get("painscore")
    ps_fill = GREEN_DARK if ps_state == "active" else RGBColor(0xCC, 0xCC, 0xCC)
    ps_text_color = WHITE if ps_state == "active" else RGBColor(0xEE, 0xEE, 0xEE)
    round_rect(s, Inches(6.65), ps_y, Inches(6.30), Inches(0.85), ps_fill, radius=0.04)
    text(s, Inches(6.65), ps_y + Inches(0.04), Inches(6.30), Inches(0.20),
         "PAIN SCORE  ·  auto-computed",
         size=9, bold=True, color=ps_text_color, font=BODY_FONT, align=PP_ALIGN.CENTER)
    if ps_state == "active":
        text(s, Inches(6.70), ps_y + Inches(0.26), Inches(6.20), Inches(0.25),
             "Pain Score · ITBA = (BV·OPCOmult·w₁ + UX·w₂ + Safety·w₃) / Effort",
             size=9, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, Inches(6.70), ps_y + Inches(0.50), Inches(6.20), Inches(0.25),
             "Pain Score · CS = (BV·OPCOmult·0.40 + UX·0.60) / Effort",
             size=9, bold=True, color=WHITE, font=BODY_FONT, align=PP_ALIGN.CENTER)
    else:
        text(s, Inches(6.65), ps_y + Inches(0.30), Inches(6.30), Inches(0.40),
             "(computed automatically after scores entered at Analyzing)",
             size=9, italic=True, color=ps_text_color, font=BODY_FONT, align=PP_ALIGN.CENTER)

    # Bottom warning — neutral light bg + green dark text (no olive)
    warn_y = Inches(5.75)
    round_rect(s, Inches(0.40), warn_y, Inches(12.55), Inches(0.32),
               GREEN_BG, radius=0.04)
    text(s, Inches(0.55), warn_y, Inches(12.20), Inches(0.32),
         "⚠  IT Now is reserved for incidents only — not for new demand",
         size=9, italic=True, color=GREEN_DARK, font=BODY_FONT,
         anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.CENTER)

    # Legend
    leg_y = Inches(6.30)
    text(s, Inches(0.40), leg_y, Inches(12.55), Inches(0.20),
         "LEGEND",
         size=8, bold=True, color=GREY_LABEL, font=BODY_FONT)
    rect(s, Inches(0.40), leg_y + Inches(0.25), Inches(0.20), Inches(0.20), AVANGRID_GREEN)
    text(s, Inches(0.65), leg_y + Inches(0.22), Inches(2.0), Inches(0.22),
         "ACTIVE · fill now",
         size=9, color=DARK_TEXT, font=BODY_FONT)
    rect(s, Inches(2.80), leg_y + Inches(0.25), Inches(0.20), Inches(0.20), RGBColor(0xEE, 0xEE, 0xEE))
    text(s, Inches(3.05), leg_y + Inches(0.22), Inches(3.0), Inches(0.22),
         "LOCKED · filled in earlier stage",
         size=9, color=DARK_TEXT, font=BODY_FONT)
    rect(s, Inches(6.20), leg_y + Inches(0.25), Inches(0.20), Inches(0.20), RGBColor(0xF8, 0xF8, 0xF8))
    text(s, Inches(6.45), leg_y + Inches(0.22), Inches(6.0), Inches(0.22),
         "FUTURE · not yet visible at this stage",
         size=9, color=DARK_TEXT, font=BODY_FONT)

    footer(s, f"Form View · {stage_name}")
    return s


def build_form_view_funnel(prs):
    return build_form_view_per_stage(prs, "Stage 1 · FUNNEL", "FUNNEL · SUBMIT", "Requester (business user)",
        {"what": "active", "routing": "active", "class": "hidden", "scores": "hidden", "painscore": "hidden"},
        AVANGRID_GREEN)


def build_form_view_reviewing(prs):
    return build_form_view_per_stage(prs, "Stage 2 · REVIEWING", "REVIEWING · post-Gate 1", "PM (Portfolio · per VS)",
        {"what": "locked", "routing": "locked", "class": "active", "scores": "hidden", "painscore": "hidden"},
        GREEN_CARD)


def build_form_view_analyzing(prs):
    return build_form_view_per_stage(prs, "Stage 3 · ANALYZING", "ANALYZING · pre-Gate 2", "PM (Portfolio · per VS) + Architect",
        {"what": "locked", "routing": "locked", "class": "locked", "scores": "active", "painscore": "active"},
        GREEN_DARK)


def build_form_views_by_role(prs):
    """3 form mockup views side-by-side: Requester · PM · Director."""
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_bar(s, prs.slide_width,
               "FORM VIEWS BY ROLE & STAGE",
               "Progressive elaboration · Jira conditional field visibility · permission scheme per role")

    col_y = Inches(1.20); col_h = Inches(5.55)
    col_w = Inches(4.10); col_gap = Inches(0.10)

    views = [
        ("REQUESTER VIEW", "FUNNEL · SUBMIT", "Sponsor / business user", AVANGRID_GREEN,
         [
            ("BUSINESS CONTEXT", [
                "Title *",
                "Description *",
                "Business Value · narrative *",
                "Acceptance Criteria *",
            ]),
            ("ROUTING", [
                "Value Stream destination *",
                "OPCO *  ·  multi-select",
                "Sponsor *  ·  user picker",
            ]),
            ("FLAGS · optional", [
                "Customer Impact Tier (P1–P4 if known)",
            ]),
            ("ABOUT SPONSOR", [
                "Sponsor = business owner who validates demand · accepts at Done",
                "Requester ≠ Sponsor · Sponsor mandatory user-picker",
            ]),
         ]),
        ("PM VIEW", "REVIEWING → ANALYZING", "PM (per VS) + Architect", GREEN_DARK,
         [
            ("REQUESTER FIELDS", [
                "All Funnel-stage fields · read-only",
            ]),
            ("CLASSIFICATION  ·  Reviewing", [
                "Type *  ·  PS / Reg / Non-Reg",
                "Size *  ·  initial guess",
                "CoS *  ·  initial guess",
                "Run / Grow / Transform *",
            ]),
            ("SCORES  ·  Analyzing", [
                "BV score 1–5 *  ·  by PM",
                "UX score 1–5 *  ·  by PM",
                "Safety score 1–5 (ITBA only) *  ·  by Architect",
                "Effort *  ·  by Architect · t-shirt size",
            ]),
            ("ARTIFACTS", [
                "Discovery Report attached",
                "Concept Prototype if Transform",
            ]),
         ]),
        ("DIRECTOR VIEW", "GATE 2 · PORTFOLIO", "Director (Carlos · Kumar)", HEADER_BLUE,
         [
            ("FULL CARD", [
                "All previous fields · read-only",
            ]),
            ("AUTO-COMPUTED · system", [
                "Pain Score (A4 auto-fires)",
                "Backlog rank (after decision)",
            ]),
            ("DECISION", [
                "Go / No-Go / Defer *",
                "Capacity slot · quarter",
                "Comments · rationale",
            ]),
            ("OUTPUTS", [
                "Status → Ready For Execution",
                "Notify Sponsor + Squad Lead",
            ]),
         ]),
    ]

    for i, (label, when, who, fill, sections) in enumerate(views):
        cx = Inches(0.40) + i * (col_w + col_gap)
        round_rect(s, cx, col_y, col_w, col_h, WHITE, radius=0.03)
        round_rect(s, cx, col_y, col_w, Inches(0.85), fill, radius=0.03)
        text(s, cx + Inches(0.15), col_y + Inches(0.08), col_w - Inches(0.30), Inches(0.30),
             label, size=14, bold=True, color=WHITE, font=BODY_FONT)
        text(s, cx + Inches(0.15), col_y + Inches(0.38), col_w - Inches(0.30), Inches(0.22),
             "▸ " + when, size=9, italic=True, color=GREEN_PASTEL, font=BODY_FONT)
        text(s, cx + Inches(0.15), col_y + Inches(0.60), col_w - Inches(0.30), Inches(0.22),
             "OWNER · " + who, size=9, bold=True, color=WHITE, font=BODY_FONT)
        by = col_y + Inches(1.00)
        for sec_label, items in sections:
            text(s, cx + Inches(0.15), by, col_w - Inches(0.30), Inches(0.20),
                 sec_label, size=8, bold=True, color=GREY_LABEL, font=BODY_FONT)
            by += Inches(0.22)
            for item in items:
                text(s, cx + Inches(0.22), by, col_w - Inches(0.44), Inches(0.20),
                     "▪ " + item, size=8.5, color=DARK_TEXT, font=BODY_FONT)
                by += Inches(0.20)
            by += Inches(0.06)

    note_y = Inches(6.85)
    round_rect(s, Inches(0.40), note_y, Inches(12.55), Inches(0.30), GREEN_BG, radius=0.04)
    text(s, Inches(0.55), note_y, Inches(12.2), Inches(0.30),
         "Jira conditional field visibility by issue status + user role  ·  permission scheme controls who can edit what  ·  asterisk (*) = mandatory at THAT stage",
         size=9, italic=True, color=GREEN_DARK, font=BODY_FONT, anchor=MSO_ANCHOR.MIDDLE)
    footer(s, "Form Views by Role & Stage")
    return s


def build_validation_roadmap(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    header_bar(s, prs.slide_width,
               "VALIDATION ROADMAP  ·  NEXT STAGE PLANNING",
               "Four phases · CS first → ITBA pilot → tooling → org rollout · exit criteria per phase")
    fl_y = Inches(1.20); fl_h = Inches(4.50)
    phases = [
        ("PHASE 1", "VALIDATE WITH CS",
         "Heather + Kumar review the model against CS reality · iterate on weights · CIT thresholds · CoS rules",
         ["Director sign-off (Kumar)",
          "CS Director endorsement of CS adaptations",
          "Pain Score weights validated for M2C-SL · M2C-RO · CCC",
          "Customer Reliability rubric defined"],
         "Weeks 1–3", AVANGRID_GREEN),
        ("PHASE 2", "ITBA PILOT",
         "Pick 1 ITBA value stream (suggest GIS — has working PDD process) · run Gate 1+Gate 2 for 1 quarter",
         ["Carlos selects pilot VS",
          "Sarath OR new PMO-Delivery role assigned",
          "1 full Gate 2 cycle completed",
          "Lessons learned captured for ITBA-wide rollout"],
         "Weeks 4–16", GREEN_DARK),
        ("PHASE 3", "TOOLING ROLLOUT",
         "Stand up Jira intake form · ServiceNow ↔ Jira API · automations A1–A6 · dashboards",
         ["Jira request types deployed (1 per VS)",
          "ServiceNow webhook live for CCC + Smart Metering",
          "Automations A1–A6 in production",
          "Layer 0 VP Dashboard live"],
         "Weeks 4–20 (parallel to Phase 2)", GREEN_MID),
        ("PHASE 4", "ORG ROLLOUT",
         "PMs onboarded · all VSs running model · Demand Manager rotation set · Champions named",
         ["3–4 PMs hired (or BRS people acting as PMs)",
          "Demand Manager assigned per VS",
          "Champions named per VS",
          "All VSs running 1 quarter at TOM"],
         "Weeks 16–32", GREEN_LIGHT),
    ]
    sw = (Inches(12.55) - 3 * Inches(0.10)) / 4
    for i, (pid, title, desc, exits, dur, fill) in enumerate(phases):
        sx = Inches(0.40) + i * (sw + Inches(0.10))
        round_rect(s, sx, fl_y, sw, fl_h, WHITE, radius=0.03)
        round_rect(s, sx, fl_y, sw, Inches(0.78), fill, radius=0.03)
        text(s, sx, fl_y + Inches(0.06), sw, Inches(0.22),
             pid, size=9, bold=True, color=GREEN_PASTEL, font=BODY_FONT, align=PP_ALIGN.CENTER)
        text(s, sx, fl_y + Inches(0.30), sw, Inches(0.36),
             title, size=12, bold=True, color=WHITE, font=BODY_FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        text(s, sx + Inches(0.12), fl_y + Inches(0.92), sw - Inches(0.24), Inches(0.85),
             desc, size=9, italic=True, color=GREEN_DARK, font=BODY_FONT)
        text(s, sx + Inches(0.12), fl_y + Inches(1.85), sw - Inches(0.24), Inches(0.18),
             "EXIT CRITERIA", size=7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        by = fl_y + Inches(2.05)
        for ex in exits:
            text(s, sx + Inches(0.12), by, sw - Inches(0.24), Inches(0.42),
                 "▪ " + ex, size=8.5, color=DARK_TEXT, font=BODY_FONT)
            by += Inches(0.42)
        text(s, sx + Inches(0.12), fl_y + Inches(4.55), sw - Inches(0.24), Inches(0.18),
             "TIMEBOX", size=7, bold=True, color=GREY_LABEL, font=BODY_FONT)
        text(s, sx + Inches(0.12), fl_y + Inches(4.73), sw - Inches(0.24), Inches(0.30),
             dur, size=10, bold=True, color=GREEN_DARK, font=BODY_FONT)
    footer(s, "Validation Roadmap · v1")
    return s


# ============================================================================
# MAIN
# ============================================================================

print("Loading TOM MAP workbook-2.pptx ...")
prs = Presentation(SRC)
total_before = len(prs.slides)
print(f"  slides BEFORE: {total_before}")

# === PRE-EDIT: DELETE shapes that should not exist anymore (Work Type, Regulatory flag) ===
print("\n[A0] Deleting obsolete shapes from form mockup (slide 20) ...")
form_idx = find_idx_by_unique(prs, "Avangrid · Digital Products  /  Create Issue")
if form_idx >= 0:
    n = delete_shapes_with_text(prs.slides[form_idx], [
        "Work Type *",
        "▼ BAU · Build · Upgrade",
        "Regulatory flag  ·  if true · auto-Fixed-Date + Compliance sponsor",
        "☐ Yes (regulatory mandate)",
    ])
    print(f"   {n} shapes deleted from form mockup")
    # Also clean up adjacent empty shapes (dropdown boxes / icons left behind) in those regions
    # Work Type field area: roughly x=8.0-10.0, y=3.5-3.8
    n2 = delete_shapes_in_region(prs.slides[form_idx], 7.9, 3.45, 10.4, 3.85)
    # Regulatory flag area: roughly x=8.0-12.5, y=5.55-5.95
    n3 = delete_shapes_in_region(prs.slides[form_idx], 7.9, 5.55, 12.8, 5.95)
    print(f"   {n2 + n3} adjacent empty shapes cleaned in those regions")

# --- Apply in-place edits ---
print("\n[A] In-place text edits (red color · same font/size) ...")

# Slides addressed by 1-based index
SLIDE_INDEX = {
    "BV (slide 24)":            23,  # 0-based
    "CoS (slide 27)":           26,
    "Type+Size (slide 28)":     27,
    "Work Type (slide 29)":     28,
    "Pain Origin (slide 30)":   29,
    "Pain Decoded (slide 32)":  31,
    "After Submission (slide 33)": 32,
    # Additional v3.1 corrections from QA pass
    "Value Stream Definition (slide 26)": 25,
    "Pain Origin extra (slide 30)":       29,
    "Pain Auto-Comp (slide 31)":          30,
    "Pain Decoded extra (slide 32)":      31,
    "CoS Type extra (slide 28)":          27,
}

# Anchor-based slides (v1-built — not at fixed index)
ANCHOR_SLIDES = {
    "Field Progression (anchor)": "FIELD PROGRESSION",
}

for slide_label, edits in EDITS.items():
    if slide_label not in SLIDE_INDEX:
        # Special — Governance slide find by anchor
        if slide_label == "Governance (slide 42)":
            idx = find_idx_by_unique(prs, "TWO GATES ABSORB")
            slide = prs.slides[idx] if idx >= 0 else None
        else:
            slide = None
    else:
        slide = prs.slides[SLIDE_INDEX[slide_label]]
    if slide is None:
        print(f"   {slide_label:35s}  NOT FOUND")
        continue
    applied = 0
    for old, new in edits:
        if edit_run_in_slide(slide, old, new):
            applied += 1
    print(f"   {slide_label:35s}  {applied}/{len(edits)} edits applied")

# Governance edit (special)
gov_idx = find_idx_by_unique(prs, "TWO GATES ABSORB")
if gov_idx >= 0:
    gov_applied = 0
    for old, new in EDITS.get("Governance (slide 42)", []):
        pass
    # Apply governance edits explicitly
    gov_slide = prs.slides[gov_idx]
    for old, new in [
        ("All VSs + business sponsors + Iberdrola",
         "Couples to existing P&T BRP + Digital CoE BRP — no separate event"),
        ("Cross-portfolio alignment. Surfaces dependencies early and locks quarterly OKRs.",
         "quarterly priorities summary enters EXISTING P&T BRP (Carlos+Asier) and Digital CoE BRP · Vitor liaises"),
    ]:
        if edit_run_in_slide(gov_slide, old, new):
            gov_applied += 1
    print(f"   {'Governance (slide 42)':35s}  {gov_applied}/2 edits applied")

# --- Apply edits to v1-built slides (Team Kanban Deep / BRP Deep / Value & Prio) ---
print("\n[B] Edits to v1-built slides ...")
for label, info in V1_BUILT_EDITS.items():
    idx = find_idx_by_unique(prs, info["anchor"])
    if idx < 0:
        print(f"   {label:30s}  NOT FOUND")
        continue
    slide = prs.slides[idx]
    applied = 0
    for old, new in info["edits"]:
        if edit_run_in_slide(slide, old, new):
            applied += 1
    print(f"   {label:30s}  {applied}/{len(info['edits'])} edits applied")

# --- Build 6 new slides ---
print("\n[C] Building 6 new slides ...")
build_routing_logic(prs);     print("   01 ROUTING LOGIC BY SIZE")
build_cos_maintain(prs);      print("   02 HOW WE MAINTAIN COS")
build_incident_flow(prs);     print("   03 INCIDENT FLOW · IT NOW → JIRA")
build_existing_project(prs);  print("   04 DEMAND ON EXISTING PROJECT")
build_digital_coe(prs);       print("   05 DIGITAL CoE × DIGITAL PRODUCTS")
build_validation_roadmap(prs);print("   06 VALIDATION ROADMAP")
build_form_views_by_role(prs);print("   07 FORM VIEWS BY ROLE & STAGE")
build_cit_field_guide(prs);   print("   08 FIELD GUIDE – CUSTOMER IMPACT TIER")
build_form_view_funnel(prs);  print("   09 FORM VIEW · FUNNEL")
build_form_view_reviewing(prs);print("   10 FORM VIEW · REVIEWING")
build_form_view_analyzing(prs);print("   11 FORM VIEW · ANALYZING")
build_post_gate2_flow(prs);   print("   12 POST-GATE 2 FLOW")
build_preparation_planning(prs);print("   13 PREPARATION & PLANNING")

# --- Reorder ---
print("\n[D] Reordering ...")
ri  = find_idx_by_unique(prs, "ROUTING LOGIC BY SIZE")
ci  = find_idx_by_unique(prs, "HOW WE MAINTAIN CLASS OF SERVICE")
ii  = find_idx_by_unique(prs, "INCIDENT FLOW  ·  IT NOW")
ei  = find_idx_by_unique(prs, "DEMAND ON EXISTING PROJECT")
di  = find_idx_by_unique(prs, "DIGITAL CoE  ×  DIGITAL PRODUCTS")
vi  = find_idx_by_unique(prs, "VALIDATION ROADMAP  ·  NEXT STAGE")

# 04 EXISTING_PROJECT after AFTER SUBMISSION
target = find_idx_by_unique(prs, "Auto-routing → Triage → Auto-conversion") + 1
move_slide(prs, ei, target)
print(f"   04 EXISTING_PROJECT → position {target+1}")

# 03 INCIDENT after AUTOMATION CATALOG
target = find_idx_by_unique(prs, "AUTOMATION CATALOG  ·  A1 → A6") + 1
ii = find_idx_by_unique(prs, "INCIDENT FLOW  ·  IT NOW")
move_slide(prs, ii, target)
print(f"   03 INCIDENT_FLOW    → position {target+1}")

# 01 ROUTING after PORTFOLIO BOARD (use unique)
target = find_idx_by_unique(prs, "CARDS ARE INITIATIVES") + 1
ri = find_idx_by_unique(prs, "ROUTING LOGIC BY SIZE")
move_slide(prs, ri, target)
print(f"   01 ROUTING_LOGIC    → position {target+1}")

# 02 COS_MAINTAIN after STEP-LEVEL POLICIES
target = find_idx_by_unique(prs, "STEP-LEVEL POLICIES") + 1
ci = find_idx_by_unique(prs, "HOW WE MAINTAIN CLASS OF SERVICE")
move_slide(prs, ci, target)
print(f"   02 COS_MAINTAIN     → position {target+1}")

# 05 DIGITAL_COE after BRP DEEP DIVE
target = find_idx_by_unique(prs, "BIG ROOM PLANNING  ·  DEEP DIVE") + 1
di = find_idx_by_unique(prs, "DIGITAL CoE  ×  DIGITAL PRODUCTS")
move_slide(prs, di, target)
print(f"   05 DIGITAL_COE      → position {target+1}")

# 06 VALIDATION_ROADMAP after WORK ALLOCATION (closes section)
target = find_idx_by_unique(prs, "PREPARATION & PLANNING –  WORK ALLOCATION")
if target < 0:
    target = find_idx_by_unique(prs, "WORK ALLOCATION")
target += 1
vi = find_idx_by_unique(prs, "VALIDATION ROADMAP  ·  NEXT STAGE")
move_slide(prs, vi, target)
print(f"   06 VALIDATION_ROADMAP → position {target+1}")

# 07 FORM_VIEWS_BY_ROLE between form mockup (slide 20) and FIELD PROGRESSION (slide 21)
form_pos = find_idx_by_unique(prs, "Avangrid · Digital Products  /  Create Issue")
fv = find_idx_by_unique(prs, "FORM VIEWS BY ROLE & STAGE")
target_fv = form_pos + 1
move_slide(prs, fv, target_fv)
print(f"   07 FORM_VIEWS_BY_ROLE → after form mockup → position {target_fv+1}")

# 08 CIT FIELD GUIDE — between Type+Size slide and Type vs R/G/T slide (after CoS detail)
ts_pos = find_idx_by_unique(prs, "Two operational fields that branch the path through gates")
cit = find_idx_by_unique(prs, "FIELD GUIDE – CUSTOMER IMPACT TIER")
target_cit = ts_pos + 1
move_slide(prs, cit, target_cit)
print(f"   08 CIT_FIELD_GUIDE → after Type+Size → position {target_cit+1}")

# 09-11 FORM VIEW per stage — insert after Form Views by Role overview
fvr_pos = find_idx_by_unique(prs, "FORM VIEWS BY ROLE & STAGE")
for label_anchor in ["TO-BE DEMAND FLOW  ·  FUNNEL · SUBMIT",
                      "TO-BE DEMAND FLOW  ·  REVIEWING · post-Gate 1",
                      "TO-BE DEMAND FLOW  ·  ANALYZING · pre-Gate 2"]:
    src_idx = find_idx_by_unique(prs, label_anchor)
    if src_idx >= 0 and fvr_pos >= 0:
        target = fvr_pos + 1
        move_slide(prs, src_idx, target)
        fvr_pos = target  # next slide goes after this one
        print(f"   {label_anchor[:50]} → position {target+1}")

# 12-13 POST-GATE 2 + PREPARATION & PLANNING — insert after Gates Deep Dive (Gate 2 details lead to next steps)
gd_pos = find_idx_by_unique(prs, "GATES DEEP DIVE")
if gd_pos >= 0:
    pg2 = find_idx_by_unique(prs, "POST-GATE 2  ·  WHAT HAPPENS NEXT")
    if pg2 >= 0:
        target = gd_pos + 1
        move_slide(prs, pg2, target)
        print(f"   POST-GATE 2 → position {target+1}")
    # Use unique anchor for the NEW Preparation & Planning slide (the subtitle is unique)
    pp = find_idx_by_unique(prs, "Level 1 · Portfolio practices  |  Level 2 · Project planning")
    if pp >= 0:
        target = find_idx_by_unique(prs, "POST-GATE 2  ·  WHAT HAPPENS NEXT") + 1
        move_slide(prs, pp, target)
        print(f"   PREPARATION & PLANNING (NEW) → position {target+1}")

# Team Kanban Deep Dive — add visual zone markers (Discovery vs Delivery)
tk_pos = find_idx_by_unique(prs, "TEAM KANBAN  ·  PHASE DEEP DIVE")
if tk_pos >= 0:
    tk_slide = prs.slides[tk_pos]
    zone_y = Inches(1.95)
    round_rect(tk_slide, Inches(0.40), zone_y, Inches(2.45), Inches(0.20),
               RGBColor(0xCC, 0xCC, 0xCC), radius=0.04)
    text(tk_slide, Inches(0.40), zone_y, Inches(2.45), Inches(0.20),
         "DISCOVERY · NOT Digital Products",
         size=8, bold=True, color=WHITE, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    delivery_x = Inches(2.92)
    delivery_w = Inches(13.33 - 2.92 - 0.40)
    round_rect(tk_slide, delivery_x, zone_y, delivery_w, Inches(0.20),
               AVANGRID_GREEN, radius=0.04)
    text(tk_slide, delivery_x, zone_y, delivery_w, Inches(0.20),
         "DELIVERY · Digital Products squads (orange zone)",
         size=8, bold=True, color=WHITE, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    print(f"   Team Kanban Deep zone markers added")

# Gates Deep Dive — restructure WHO PARTICIPATES with explicit roles (CHAIR · DECIDERS · ATTENDEES · OBSERVERS)
gd_idx_struct = find_idx_by_unique(prs, "GATES DEEP DIVE")
if gd_idx_struct >= 0:
    gd_slide = prs.slides[gd_idx_struct]
    # Find each WHO PARTICIPATES content shape (Gate 1 + Gate 2) by position
    # Gate 1 box is on left half, Gate 2 on right half
    for sh in gd_slide.shapes:
        if not sh.has_text_frame: continue
        try:
            x = sh.left/914400 if sh.left else 0
        except: continue
        txt = sh.text_frame.text.strip()
        # Gate 1 WHO PARTICIPATES (left side, x < 6.5)
        if x < 6.5 and ("OBSERVERS" in txt or "ATTENDEES" in txt or "CHAIR" in txt) and "PARTICIPATES" not in txt:
            # Clear and rewrite with proper 3-role structure for Gate 1
            tf = sh.text_frame
            for p in list(tf.paragraphs):
                for r in list(p.runs):
                    r.text = ""
            for p in list(tf.paragraphs)[1:]:
                p._p.getparent().remove(p._p)
            lines = [
                "▸ CHAIR · Lead PM (rotating among VS PMs)",
                "▸ ATTENDEES · PMs (one per VS) · Architect",
                "▸ OBSERVERS · Sponsor on request",
            ]
            p0 = tf.paragraphs[0]
            for i, line in enumerate(lines):
                p = p0 if i == 0 else tf.add_paragraph()
                # remove existing runs
                for r in list(p.runs):
                    r.text = ""
                if not p.runs:
                    r = p.add_run()
                else:
                    r = p.runs[0]
                r.text = line
                r.font.size = Pt(8)
                r.font.color.rgb = DARK_TEXT
                r.font.name = BODY_FONT
        # Gate 2 WHO PARTICIPATES (right side, x > 6.5)
        elif x > 6.5 and ("ATTENDEES" in txt or "Director" in txt) and "PARTICIPATES" not in txt:
            tf = sh.text_frame
            for p in list(tf.paragraphs):
                for r in list(p.runs):
                    r.text = ""
            for p in list(tf.paragraphs)[1:]:
                p._p.getparent().remove(p._p)
            lines = [
                "▸ CHAIR · VP (Vitor Cruz)",
                "▸ DECIDERS · Directors — Carlos Garbin (ITBA) · Kumar Jaladanki (CS) · Heather LaJoie (CS)",
                "▸ ATTENDEES · PMs (per VS) · Architect · IT Program Delivery Manager",
                "▸ OBSERVERS · Sponsors (on relevant items only)",
            ]
            p0 = tf.paragraphs[0]
            for i, line in enumerate(lines):
                p = p0 if i == 0 else tf.add_paragraph()
                for r in list(p.runs):
                    r.text = ""
                if not p.runs:
                    r = p.add_run()
                else:
                    r = p.runs[0]
                r.text = line
                r.font.size = Pt(8)
                r.font.color.rgb = DARK_TEXT
                r.font.name = BODY_FONT
    print(f"   Gates Deep Dive WHO PARTICIPATES restructured (CHAIR · DECIDERS · ATTENDEES · OBSERVERS)")
    # Also rebuild HOW · STEP-BY-STEP for Gate 1 (it has only 1 step now, should have 5)
    for sh in gd_slide.shapes:
        if not sh.has_text_frame: continue
        try:
            x = sh.left/914400 if sh.left else 0
        except: continue
        txt = sh.text_frame.text.strip()
        # Gate 1 HOW STEP-BY-STEP (left side, contains "Auto-routing" or partial steps)
        if x < 6.5 and "Auto-routing recorded" in txt and "1." not in txt:
            tf = sh.text_frame
            for p in list(tf.paragraphs):
                for r in list(p.runs):
                    r.text = ""
            for p in list(tf.paragraphs)[1:]:
                p._p.getparent().remove(p._p)
            steps_g1 = [
                "1. Lead PM pulls all NEW + ROUTED tickets from prior 7 days",
                "2. For each: validate DoR (Title · Description · BV narrative · AC · VS · OPCO · Sponsor)",
                "3. Decision  ·  ACCEPT (→ Reviewing)  ·  BOUNCE (→ Requestor with reason)  ·  MERGE (→ existing item)",
                "4. Auto-routing recorded in Jira · @notify PM of accepted items",
                "5. KPI captured  ·  DoR pass rate · bounce rate · channel adoption",
            ]
            p0 = tf.paragraphs[0]
            for i, line in enumerate(steps_g1):
                p = p0 if i == 0 else tf.add_paragraph()
                for r in list(p.runs):
                    r.text = ""
                if not p.runs:
                    r = p.add_run()
                else:
                    r = p.runs[0]
                r.text = line
                r.font.size = Pt(8)
                r.font.color.rgb = DARK_TEXT
                r.font.name = BODY_FONT
            print(f"   Gate 1 HOW · STEP-BY-STEP rebuilt (5 steps)")
            break

# ORIGINAL Team Kanban slide (FIVE PHASES BY TWO TRACKS) — Ana feedback explicit
tk_orig_pos = find_idx_by_unique(prs, "FIVE PHASES BY TWO TRACKS")
if tk_orig_pos >= 0:
    tk_orig = prs.slides[tk_orig_pos]
    # Add prominent banner explaining ORANGE = us, BLUE = others
    # Phase track section is around y=2.0 to 6.5, with DISCOVERY blue at top half, DELIVERY orange at bottom
    # Add zone-purpose banners in the empty space at top-left
    banner_y = Inches(1.55)
    # Banner 1: Discovery (NOT us) explanation
    round_rect(tk_orig, Inches(4.30), banner_y, Inches(8.65), Inches(0.30),
               RGBColor(0x33, 0x55, 0x88), radius=0.04)
    text(tk_orig, Inches(4.30), banner_y, Inches(8.65), Inches(0.30),
         "BLUE · DISCOVERY TRACK — run by domain product team / vendor · NOT Digital Products",
         size=9, bold=True, color=WHITE, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Banner 2: Delivery (us) explanation
    banner2_y = Inches(1.85)
    round_rect(tk_orig, Inches(4.30), banner2_y, Inches(8.65), Inches(0.30),
               RGBColor(0xE8, 0x80, 0x30), radius=0.04)
    text(tk_orig, Inches(4.30), banner2_y, Inches(8.65), Inches(0.30),
         "ORANGE · DELIVERY TRACK — Digital Products squads · we are almost ALWAYS here (Phase 2-4)",
         size=9, bold=True, color=WHITE, font=BODY_FONT,
         align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Handoff suggestion: how worlds connect
    connect_y = Inches(7.05)
    round_rect(tk_orig, Inches(0.40), connect_y, Inches(12.55), Inches(0.40),
               GREEN_BG, radius=0.04)
    text(tk_orig, Inches(0.55), connect_y + Inches(0.04), Inches(12.20), Inches(0.18),
         "HOW THE TWO WORLDS CONNECT  ·  Digital Products PM advisor in Discovery + DoR gate at handoff",
         size=9, bold=True, color=GREEN_DARK, font=BODY_FONT)
    text(tk_orig, Inches(0.55), connect_y + Inches(0.20), Inches(12.20), Inches(0.18),
         "1) PM (Digital Products) participates as ADVISOR in Phase 0/1 reviews  ·  2) DoR check gates Phase 1→2 (Materialization → Creation)  ·  3) Handoff package: Validated problem · Roadmap · Sizing · Budget",
         size=8, italic=True, color=DARK_TEXT, font=BODY_FONT)
    print(f"   Original Team Kanban (slide {tk_orig_pos+1}) — orange/blue zone explanation + handoff added")

# === GLOBAL CLEANUP PASSES (after all targeted edits + new slides) ===
print("\n[E] Global cleanup passes ...")

# Pass 1: NO Demand Manager role — replace with PM globally in slides 17-60
n = global_replace_run(prs, (17, 60), "Demand Manager", "PM")
print(f"   'Demand Manager' → 'PM' in {n} runs (slides 17-60)")

# Pass 2: Remove "REMOVED" / "absorbed" / "now in" language — keep clean new content
cleanups = [
    ("(REMOVED · absorbed into Type)", ""),
    ("(REMOVED · now in Type)", ""),
    ("(REMOVED · use Type = Regulatory to trigger A3)", ""),
    ("Work Type (REMOVED · now in Type)", ""),
    ("Work Type *", ""),
    ("(absorbed into Type field above)", ""),
    ("▼ (absorbed into Type field above)", ""),
    ("(field removed · Type=Regulatory now triggers A3 Auto-Fixed-Date)", ""),
    ("Regulatory flag (REMOVED · use Type = Regulatory to trigger A3)", ""),
    # parenthetical historical explanations on Field Progression
    ("   (Work Type removed · absorbed into Type)", ""),
    ("   (Regulatory flag removed · now Type=Regulatory)", ""),
    # Pain Decoded "(ITBA only · CS drops it)" → just "(ITBA only)"
    ("(ITBA only · CS drops it)", "(ITBA only)"),
    # Slide 28 Type body — clean up Incident references
    ("(Incident auto from IT Now · not in Type)", ""),
    ("(Incident auto from IT Now)", ""),
    # Slide 27 "blocked-out per quarter" historical
    ("Auto-routing only works", "Auto-routing only works"),  # placeholder
    # Old "Safety REMOVED" → just don't mention Safety on CS card
    ("Safety REMOVED for CS", ""),
    ("Safety REMOVED", "Safety (ITBA)"),
]
for old, new in cleanups:
    n = global_replace_run(prs, (17, 60), old, new, color_red=(new != ""))
    if n > 0:
        print(f"   removed '{old[:45]}'  in {n} runs")

# Pass 3: Replace H(x) with Pain Score globally (slides 17-60)
n = global_replace_run(prs, (17, 60), "H(x)", "Pain Score")
print(f"   'H(x)' → 'Pain Score' in {n} runs")

# Pass 4: Replace form mockup formula shape (handles split runs)
form_idx = find_idx_by_unique(prs, "Avangrid · Digital Products  /  Create Issue")
if form_idx >= 0:
    ok = replace_shape_text_red(
        prs.slides[form_idx],
        "Pain Score = (",
        "Pain Score · ITBA = (BV·OPCOmult·w₁ + UX·w₂ + Safety·w₃) / Effort\n"
        "Pain Score · CS    = (BV·OPCOmult·0.40 + UX·0.60) / Effort",
        size=10, bold=True)
    if not ok:
        # try alternative (in case H(x) global replace already ran but shape still has BV·w)
        ok = replace_shape_text_red(
            prs.slides[form_idx], "BV·w",
            "Pain Score · ITBA = (BV·OPCOmult·w₁ + UX·w₂ + Safety·w₃) / Effort\n"
            "Pain Score · CS    = (BV·OPCOmult·0.40 + UX·0.60) / Effort",
            size=10, bold=True)
    print(f"   Form mockup formula → CS+ITBA   {'OK' if ok else 'NOT FOUND'}")

# Pass 5: Replace formula shape on slide 31 (Pain Auto-Comp) — split runs
auto_idx = find_idx_by_unique(prs, "PAIN SCORE · AUTO-COMPUTATION")
if auto_idx >= 0:
    ok = replace_shape_text_red(
        prs.slides[auto_idx],
        "BV·OPCOmult·w",
        "Pain Score · ITBA = (BV·OPCOmult·w₁ + UX·w₂ + Safety·w₃) / Effort     ·     "
        "Pain Score · CS = (BV·OPCOmult·0.40 + UX·0.60) / Effort",
        size=14, bold=True)
    print(f"   Slide 31 formula → CS+ITBA      {'OK' if ok else 'NOT FOUND'}")

# Pass 6: Replace formula shape on slide 32 (Pain Decoded)
dec_idx = find_idx_by_unique(prs, "FORMULA DECODED")
if dec_idx >= 0:
    ok = replace_shape_text_red(
        prs.slides[dec_idx],
        "BV·OPCOmult·w",
        "Pain Score · ITBA = (BV·OPCOmult·w₁ + UX·w₂ + Safety·w₃) / Effort     ·     "
        "Pain Score · CS = (BV·OPCOmult·0.40 + UX·0.60) / Effort",
        size=14, bold=True)
    print(f"   Slide 32 formula → CS+ITBA      {'OK' if ok else 'NOT FOUND'}")

# Pass 7: Add OPCOmult as 5th BOX (matching style of BV/UX/Safety/Effort) on slides 31 and 32
# This replaces the previous bottom panel approach (which overlapped content)
for anchor in ["PAIN SCORE · AUTO-COMPUTATION", "FORMULA DECODED"]:
    ok = add_opcomult_5th_box_to_slide(prs, anchor)
    print(f"   OPCOmult 5th box added to slide ({anchor[:30]})  {'OK' if ok else 'NOT FOUND'}")

# Pass 9: DELETE Incident Flow slide entirely (incidents always in IT Now — no Jira flow)
inc_idx = find_idx_by_unique(prs, "INCIDENT FLOW  ·  IT NOW")
if inc_idx >= 0:
    delete_slide(prs, inc_idx)
    print(f"   Incident Flow slide DELETED (was slide {inc_idx+1})")

# Pass 10: Clean Incident references on slide 28 (Type+Size) — relabel Incident card to Regulatory
def force_replace_in_slide(slide, replacements):
    """Apply text replacements per run, color RED, count."""
    n = 0
    for sh in slide.shapes:
        if not sh.has_text_frame: continue
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                for old, new in replacements:
                    if old in r.text:
                        r.text = r.text.replace(old, new)
                        r.font.color.rgb = RED_EDIT
                        n += 1
                        break
    return n

slide28_idx = find_idx_by_unique(prs, "Two operational fields that branch the path through gates")
if slide28_idx >= 0:
    n = force_replace_in_slide(prs.slides[slide28_idx], [
        ("Incident", "Regulatory"),
        ("Production outage · break-fix · NOT user-selectable — auto-created from ServiceNow (IT Now)",
         "Mandatory · external mandate · PUC · SOX · Iberdrola · vendor SLA · regulator-driven deadline"),
        ("Production outage · break-fix · service degradation",
         "Mandatory · external mandate · PUC · SOX · Iberdrola · vendor SLA"),
        ("Auto-created when ServiceNow P1/P2 ticket fires · API → Jira creates Bug · A2 sets CoS=Expedite + WIP-1",
         "Triggers A3 Auto-Fixed-Date · CoS=Fixed-Date · mandatory Gate 2 · compliance sponsor required"),
        ("Bypasses Gates. Auto-converts to Bug. Routed directly to Squad Expedite swimlane. CoS = Expedite.",
         "Full Triage with auto Fixed-Date CoS · mandatory Gate 2 · scheduled backwards from deadline"),
        ("Skip BACKLOG · skip Gates · go to Squad Expedite",
         "Full Triage · auto Fixed-Date · Gate 2 mandatory"),
    ])
    print(f"   Slide 28 Incident → Regulatory  ({n} runs)")

# Pass 11: Clean Incident column references on slide 33 (After Submission) — mark as IT Now only
sub_idx = find_idx_by_unique(prs, "Auto-routing → Triage → Auto-conversion")
if sub_idx >= 0:
    n = force_replace_in_slide(prs.slides[sub_idx], [
        ("Type = Incident", "(Incidents → IT Now only)"),
        ("BUG · CoS = Expedite", "(out of Jira scope)"),
        ("Squad Expedite swimlane · WIP-1", "(handled in ServiceNow)"),
        ("BUG ·", ""),
        ("URGENT", ""),
    ])
    print(f"   Slide 33 Incident column cleaned  ({n} runs)")

# Pass 12: Clean Incident column on slide 38 (Routing Logic)
rl_idx = find_idx_by_unique(prs, "ROUTING LOGIC BY SIZE")
if rl_idx >= 0:
    n = force_replace_in_slide(prs.slides[rl_idx], [
        ("Type = Incident\n(auto from IT Now)", "(Incidents → IT Now only)"),
        ("Type = Incident", "(Incidents → IT Now only)"),
        ("Bug · CoS=Expedite", "(out of Jira scope)"),
        ("SQUAD EXPEDITE LANE\nWIP-1", "(handled in ServiceNow)"),
        ("Bypasses Triage · A2 fires", "(no Jira route)"),
    ])
    print(f"   Slide 38 Incident column cleaned  ({n} runs)")

# Pass 13: Final brute-force — any remaining "Incident" text in slides 17-60 cleanup
n = global_replace_run(prs, (17, 60), "Incident", "(Incidents handled in IT Now only)")
if n > 0:
    print(f"   Residual 'Incident' refs cleaned  ({n} runs)")

# Pass 14: Remove residual Work Type mentions in manager slides
for old, new in [
    ("Work Type (BAU / Build / Upgrade)", "Type"),
    ("Work Type ", "Type "),
    ("Work Type", "Type"),
    ("BAU · Build · Upgrade", "Production Support · Regulatory · Non-Regulatory"),
]:
    n = global_replace_run(prs, (17, 60), old, new)
    if n > 0:
        print(f"   Residual '{old[:40]}' cleaned  ({n} runs)")

# Pass 14b: Extend cleanup to ALL slides for Demand Manager + H(x) (slides > 60 still have residuals)
n = global_replace_run(prs, (1, 200), "Demand Manager", "PM")
if n > 0: print(f"   'Demand Manager' → 'PM' (extended)  ({n} runs)")
# Case variants
for variant in ["DEMAND MANAGER", "demand manager"]:
    n = global_replace_run(prs, (1, 200), variant, "PM")
    if n > 0: print(f"   '{variant}' → 'PM'  ({n} runs)")

# Fix PM duplications caused by global replace
for old, new in [
    ("PM (chair) + PM + Architect",   "PM (chair) + Architect"),
    ("PM (chair) + PM",                "PM (chair)"),
    ("PM + PM propose",                "PM proposes"),
    ("PM + PM",                        "PM"),
]:
    n = global_replace_run(prs, (1, 200), old, new)
    if n > 0: print(f"   PM dedup: '{old}' → '{new}'  ({n} runs)")

# Clarify PM works at PORTFOLIO level during Reviewing/Analyzing (not team-level)
for old, new in [
    ("PM (per VS) + Architect",        "PM (Portfolio · per VS) + Architect"),
    ("PM (per VS) + Sponsor",          "PM (Portfolio · per VS) + Sponsor"),
    ("PM (per VS)",                    "PM (Portfolio · per VS)"),
]:
    n = global_replace_run(prs, (1, 200), old, new)
    if n > 0: print(f"   PM clarify: '{old}' → '{new}'  ({n} runs)")
# H(x) → Pain Score globally EXCEPT slide 30 (P&S Networks historical formula)
slide30_idx = 29  # 0-based slide 30 (Pain Score Origin) — keep historical P&S H(x) formula
for i, slide in enumerate(prs.slides, 1):
    if i == 30: continue  # skip P&S historical
    for sh in slide.shapes:
        if not sh.has_text_frame: continue
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if "H(x)" in r.text:
                    r.text = r.text.replace("H(x)", "Pain Score")
                    r.font.color.rgb = RED_EDIT
print(f"   'H(x)' → 'Pain Score' globally (skipping slide 30 P&S historical)")

# Pass 14c: Slide 20 form mockup asterisks — remove * from non-Funnel-stage fields
form_idx2 = find_idx_by_unique(prs, "Avangrid · Digital Products  /  Create Issue")
if form_idx2 >= 0:
    n = force_replace_in_slide(prs.slides[form_idx2], [
        ("Type *", "Type"),
        ("Size *", "Size"),
        ("Class of Service *", "Class of Service"),
        ("Run / Grow / Transform *", "Run / Grow / Transform"),
    ])
    print(f"   Slide 20 asterisks adjusted (Funnel-only)  ({n} runs)")

# Pass 14d: Slide 24 BV — modify Safety bullet IN-PLACE to show both ITBA and CS variants
bv_idx = find_idx_by_unique(prs, "FIELD GUIDE – BUSINESS VALUE")
if bv_idx >= 0:
    bv_slide = prs.slides[bv_idx]
    # Modify "▸ Safety" header → split label
    force_replace_in_slide(bv_slide, [
        # Replace Safety header to mention both ITBA and CS-Customer Reliability
        ("▸ Safety", "▸ Safety (ITBA) / Customer Reliability (CS)"),
        # Replace Safety description to include CS variant inline
        ("Field crew exposure · public hazard · near-miss",
         "ITBA · field crew · public hazard · near-miss     CS · uptime · NPS hits · service availability"),
    ])
    # Update TRIGGERS / DRIVES with weights per VS
    force_replace_in_slide(bv_slide, [
        ("▸ Pain Score (BV component · weight 0.35 in CS) · Gate 2 review · Portfolio Committee defense",
         "▸ Pain Score · BV component  ·  ITBA: w₁=0.45 · CS: w₁=0.40  ·  Gate 2 review · Portfolio Committee defense"),
    ])
    print(f"   Slide 24 BV split (Safety ITBA / Customer Reliability CS) + weights per VS — in-place")

# Pass 14e: Slides 30 & 32 — BV/UX scored by PM, Safety/Effort by Architect (position-aware)
def position_aware_score_fix(slide, expected_x_ranges):
    """For each variable column, fix the 'scored by' line based on column position."""
    for sh in slide.shapes:
        if not sh.has_text_frame: continue
        try: x = sh.left/914400 if sh.left else 0
        except: continue
        # Determine which column this shape is in
        column = None
        for col_name, (xmin, xmax) in expected_x_ranges.items():
            if xmin <= x <= xmax:
                column = col_name; break
        if column is None: continue
        # Apply column-specific replacements
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                txt = r.text
                if column in ("BV", "UX"):
                    if "scored by Architect at Analyzing" in txt:
                        r.text = txt.replace("scored by Architect at Analyzing", "scored by PM at Analyzing")
                        r.font.color.rgb = RED_EDIT
                    elif "scored at Triage" in txt:
                        r.text = txt.replace("scored at Triage", "scored by PM at Analyzing")
                        r.font.color.rgb = RED_EDIT
                elif column == "Effort":
                    if "sized by SMEs at TRIAGE" in txt or "sized by SMEs at Triage" in txt:
                        r.text = txt.replace("sized by SMEs at TRIAGE", "sized by Architect at Analyzing").replace("sized by SMEs at Triage", "sized by Architect at Analyzing")
                        r.font.color.rgb = RED_EDIT

# Apply to slide 32 (Pain Score Decoded)
dec_idx2 = find_idx_by_unique(prs, "FORMULA DECODED")
if dec_idx2 >= 0:
    position_aware_score_fix(prs.slides[dec_idx2], {
        "BV":     (0.0, 3.3),
        "UX":     (3.3, 6.4),
        "Safety": (6.4, 9.5),
        "Effort": (9.5, 13.5),
    })
    print(f"   Slide 32 BV/UX scored by PM · Safety/Effort by Architect")

# Pass 14f: CIT card on slide 27 (CoS) REMOVED — now CIT has dedicated Field Guide slide

# Pass 14j: CoS slide badge — DEMAND MANAGER + ARCHITECT → PM + ARCHITECT (slide 27 was missed earlier)
n = global_replace_run(prs, (17, 60), "REVIEWING / ANALYZING  ·  DEMAND MANAGER + ARCHITECT",
                       "REVIEWING / ANALYZING  ·  PM + ARCHITECT")
if n > 0: print(f"   CoS badge: 'DEMAND MANAGER' → 'PM'  ({n} runs)")
# Also handle other variants
for old, new in [
    ("DEMAND MANAGER + ARCHITECT", "PM + ARCHITECT"),
    ("DEMAND MANAGER (chair) + PM + Architect", "PM (chair) + Architect"),
    ("Demand Manager (chair)", "PM (chair)"),
]:
    n = global_replace_run(prs, (1, 200), old, new)
    if n > 0: print(f"   '{old[:40]}' → cleaned  ({n} runs)")

# Pass 14k: XL → PPM Now (XL demands handled in PPM Now, NOT in our Jira flow)
# Field Progression slide 22 + Type+Size slide 28 + various
for old, new in [
    ("T-shirt: S · M · L · XL — best-guess",                   "T-shirt: S · M · L — best-guess  (XL → PPM Now)"),
    ("T-shirt: S · M · L · XL",                                 "T-shirt: S · M · L  (XL → PPM Now)"),
    ("S · M · L · XL — best-guess",                             "S · M · L — best-guess  (XL → PPM Now)"),
    ("Size: S · M · L · XL",                                    "Size: S · M · L  (XL → PPM Now)"),
    ("Type (Production Support · Regulatory · Non - Regulatory)",
     "Type (Production Support · Regulatory · Non-Regulatory)"),  # fix typo
    # Update XL examples / mentions in body
    ("Size = XL  ·  cross-team",                                "(XL → handled in PPM Now · NOT Digital Products Jira)"),
    ("Initiative sliced into Epics (Initiative > 3 months rule)",
     "(XL → PPM Now)  ·  not in Digital Products Jira flow"),
    ("Triage SLA: 10 business days · sponsor sign-off required",
     "(XL → PPM Now)  ·  out of scope for this flow"),
    ("Sponsor: Vitor Cruz · Iberdrola for cross-OPCO XL",       "(XL → PPM Now)"),
    ("Extra-Large  ·  > 60 days · cross-area",                  "(XL → handled in PPM Now)"),
    ("Full DoR + business case + capital plan alignment",       "(XL → PPM Now)  ·  out of Digital Products scope"),
]:
    n = global_replace_run(prs, (17, 60), old, new)
    if n > 0: print(f"   XL→PPM Now: '{old[:50]}' cleaned  ({n} runs)")

# Pass 14l: Slide 33 (After Submission) — Step 2 clarify board + handle XL column
sub_idx = find_idx_by_unique(prs, "Auto-routing → Triage → Auto-conversion")
if sub_idx >= 0:
    n = force_replace_in_slide(prs.slides[sub_idx], [
        # Step 2 description: clarify which board
        ("reads Value Stream → assigns board",
         "reads Value Stream → assigns CS Portfolio Board OR ITBA Portfolio Board"),
        # XL column entire content → PPM Now
        ("Size = XL  ·  cross-team", "Size = XL"),
        ("INITIATIVE + placeholder Epics", "(XL → PPM Now)"),
        ("Enterprise Dashboard · Gate 2 next", "Out of Digital Products Jira scope"),
        ("Enterprise Dashboard · Gate 1 first", "Out of Digital Products Jira scope"),
    ])
    print(f"   Slide 33 Step 2 board + XL→PPM Now  ({n} runs)")

# Pass 14m: Routing Logic by Size slide 38 — XL column → PPM Now
rl_idx = find_idx_by_unique(prs, "ROUTING LOGIC BY SIZE")
if rl_idx >= 0:
    n = force_replace_in_slide(prs.slides[rl_idx], [
        ("Size = XL\ncross-team", "Size = XL\n(out of scope)"),
        ("Size = XL", "Size = XL"),
        ("Initiative + Epics", "(handled in PPM Now)"),
        ("ENTERPRISE DASHBOARD\nGATE 2 + tracked at VP layer", "PPM Now\n(NOT in Digital Products Jira)"),
        ("Owns roadmap horizon · multi-quarter", "PPM Now manages cross-quarter XL initiatives"),
    ])
    print(f"   Slide 38 Routing Logic XL → PPM Now  ({n} runs)")

# Pass 14n: Form mockup slide 20 — minimal field annotations (stage · owner) on EVERY field
form_idx3 = find_idx_by_unique(prs, "Avangrid · Digital Products  /  Create Issue")
if form_idx3 >= 0:
    fslide = prs.slides[form_idx3]
    n = force_replace_in_slide(fslide, [
        # FUNNEL · REQUESTER fields
        ("Title *  ·  outcome-based, not task-based",
         "Title *  ·  outcome-based, not task-based  ·  Funnel · Requester"),
        ("Description *  ·  describe the problem, not the solution",
         "Description *  ·  describe the problem, not the solution  ·  Funnel · Requester"),
        ("Business Value *  ·  why it matters",
         "Business Value *  ·  why it matters  ·  Funnel narrative (Requester) + 1–5 score at Analyzing (PM)"),
        ("Acceptance Criteria *  ·  measurable",
         "Acceptance Criteria *  ·  measurable  ·  Funnel · Requester"),
        ("Value Stream destination *",
         "Value Stream destination *  ·  drives auto-routing  ·  Funnel · Requester"),
        ("Customer Impact Tier  ·  CS only · drives Auto-Expedite",
         "Customer Impact Tier  ·  CS only · drives Auto-Expedite  ·  Funnel · Requester (optional)"),
        # REVIEWING · PM fields
        ("Class of Service  ·  swimlane assignment",
         "Class of Service  ·  swimlane assignment  ·  Reviewing · PM (initial guess)"),
        ("Class of Service *  ·  swimlane assignment",
         "Class of Service *  ·  swimlane assignment  ·  Reviewing · PM (initial guess)"),
        # OPCO and Sponsor (Funnel · Requester)
        ("OPCO *", "OPCO *  ·  multi-select  ·  Funnel · Requester"),
        ("Sponsor *", "Sponsor *  ·  business owner  ·  Funnel · Requester"),
        ("Due Date", "Due Date  ·  auto if Type=Regulatory  ·  Reviewing · PM"),
        ("Run / Grow / Transform *", "Run / Grow / Transform *  ·  Reviewing · PM"),
    ])
    print(f"   Form mockup slide 20 field annotations (full sweep)  ({n} runs)")
    # Add small floating "Reviewing · PM" labels next to Type and Size fields (which don't have a long label)
    text(fslide, Inches(7.95), Inches(3.86), Inches(2.40), Inches(0.16),
         "Type  ·  Reviewing · PM",
         size=8, italic=True, color=RED_EDIT, font=BODY_FONT)
    text(fslide, Inches(10.38), Inches(3.86), Inches(2.40), Inches(0.16),
         "Size  ·  Reviewing · PM",
         size=8, italic=True, color=RED_EDIT, font=BODY_FONT)
    # Add placeholder text inside BV and AC empty input fields (matching Description placeholder style)
    text(fslide, Inches(0.70), Inches(3.55), Inches(3.40), Inches(0.28),
         "$ impact · % delta · NPS · CSAT · time saved · volume · compliance · safety / customer reliability",
         size=8, italic=True, color=GREY_LABEL, font=BODY_FONT)
    text(fslide, Inches(4.40), Inches(3.55), Inches(3.40), Inches(0.28),
         "list of measurable conditions · binary pass/fail · who confirms (UAT)",
         size=8, italic=True, color=GREY_LABEL, font=BODY_FONT)
    # Add a SCORES (1-5) section as a clean banner ABOVE the Pain Score formula box
    # Position chosen between Customer Impact Tier (y~5.85) and CONTEXT (y~6.20) — small banner
    # Actually best place: below CONTEXT row, above Pain Score box
    scores_y = Inches(6.55)
    round_rect(fslide, Inches(7.95), scores_y, Inches(4.80), Inches(0.32), GREEN_BG, radius=0.04)
    text(fslide, Inches(8.05), scores_y + Inches(0.04), Inches(4.65), Inches(0.18),
         "SCORES (1–5) — added at Analyzing",
         size=8, bold=True, color=RED_EDIT, font=BODY_FONT)
    text(fslide, Inches(8.05), scores_y + Inches(0.16), Inches(4.65), Inches(0.16),
         "BV ▢   ·   UX ▢   ·   Safety ▢ (ITBA)   ·   Effort ▢      |  BV/UX by PM · Safety/Effort by Architect",
         size=7, color=DARK_TEXT, font=BODY_FONT)
    print(f"   Form mockup slide 20: BV/AC placeholders + SCORES (1-5) section added")

# Pass 14g: Add Discovery → flow connection annotation on slide 18
disc_idx = find_idx_by_unique(prs, "Four core practices that reshape the intake conversation")
if disc_idx >= 0:
    disc_slide = prs.slides[disc_idx]
    # Find the LIGHTWEIGHT BUT FORMAL DISCOVERY card and add connection note below it
    # Position at bottom of slide where there's space
    note_y = Inches(6.55)
    text(disc_slide, Inches(8.96), note_y, Inches(2.37), Inches(0.50),
         "Discovery feeds: Description (Funnel)  ·  BV/UX scoring (Analyzing)  ·  DoR check at Gate 1",
         size=8, italic=True, color=RED_EDIT, font=BODY_FONT)
    print(f"   Discovery → flow connection annotation added to slide {disc_idx+1}")

# Pass 14h: Add UX visibility link on Field Progression (slide 22)
fp_idx = find_idx_by_unique(prs, "FIELD PROGRESSION")
if fp_idx >= 0:
    fp_slide = prs.slides[fp_idx]
    # Add small text near bottom referencing UX rubric location
    text(fp_slide, Inches(0.55), Inches(7.05), Inches(8.0), Inches(0.18),
         "▸ UX rubric defined on Pain Score · Formula Decoded slide  ·  scored 1–5 by PM at Analyzing",
         size=8, italic=True, color=RED_EDIT, font=BODY_FONT)
    print(f"   UX visibility link added to slide {fp_idx+1}")

# Pass 14i: Clean residual Work Type on slides 68, 75, 76 (manager appendix)
for old, new in [
    ("Work Type (BAU / Build / Upgrade) ", "Type "),
    ("Work Type (BAU / Build / Upgrade)", "Type"),
    ("Competing Demand Across Work Types", "Competing Demand Across Types"),
]:
    n = global_replace_run(prs, (60, 92), old, new)
    if n > 0: print(f"   '{old[:35]}' cleaned in appendix  ({n} runs)")

# Pass 14z: ServiceNow integration removal — incidents stay in ServiceNow · NO Jira integration
# Per user decision: all incidents are managed in ServiceNow · we don't track them in Jira at all
sn_removals = [
    "(Incidents handled in IT Now only)",
    "(Incidents → IT Now only)",
    "Incidents handled in IT Now only",
    "Incidents → IT Now only",
    "ServiceNow → Jira link active for CCC + Smart Metering",
    "ServiceNow webhook live for CCC + Smart Metering",
    "ServiceNow ↔ Jira API",
    "auto from IT Now",
    "(auto from IT Now)",
]
for pat in sn_removals:
    n = global_replace_run(prs, (1, 200), pat, "")
    if n > 0:
        print(f"   ServiceNow ref removed: '{pat[:45]}'  ({n} runs)")

# Also delete any standalone shapes containing only the warning IT Now text on form mockup slides
form_idx_sn = find_idx_by_unique(prs, "Avangrid · Digital Products  /  Create Issue")
if form_idx_sn >= 0:
    fs = prs.slides[form_idx_sn]
    for sh in list(fs.shapes):
        if not sh.has_text_frame: continue
        if "IT Now is reserved for incidents only" in sh.text_frame.text:
            try:
                sh._element.getparent().remove(sh._element)
                print(f"   ServiceNow warning shape removed from form mockup")
            except: pass

# Pass 15: BRP usage cleanup — Digital Products has NO own BRP (couples to P&T BRP + Digital CoE BRP)
# Replace ambiguous "BRP" references on CoS / capacity slides with "Gate 2 (monthly)"
for old, new in [
    ("% set quarterly at BRP", "% set at Gate 2 (monthly)"),
    ("validated each quarter at BRP", "validated each Gate 2 (monthly Portfolio Committee)"),
    ("End of quarter · BRP outcome", "End of quarter · feeds quarterly priorities summary to existing P&T BRP / Digital CoE BRP"),
    ("BRP outcome", "quarterly priorities summary to existing BRPs"),
]:
    n = global_replace_run(prs, (17, 60), old, new)
    if n > 0:
        print(f"   '{old[:40]}' → cleaned  ({n} runs)")

# --- Save ---
print(f"\n[F] Saving → {DST}")
prs.save(DST)
print(f"   slides AFTER: {len(prs.slides)}")
print("\nDONE.")
