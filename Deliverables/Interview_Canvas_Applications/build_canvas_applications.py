"""
Applications Team Interview Canvas — Avangrid SAP Digital Products
NTT DATA | Wave 1 Discovery Dynamic

Focus topics (extends the 5 lifecycle stages already mapped in the Demand Map):
  6. Capacity & Resource Model
  7. Reporting & Metrics Review

Designed as a live workshop canvas: 6 blocks × timeboxed dynamics. Each block groups
questions of SIMILAR FORMAT (so the dynamic doesn't break rhythm) and adds
SYNERGY questions that link across topics.

Total duration: ~90 minutes.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ========== NTT BRAND TOKENS ==========
NAVY = RGBColor(0x0F, 0x1C, 0x50)
BLUE = RGBColor(0x67, 0x85, 0xC1)
ACCENT = RGBColor(0xBC, 0x43, 0x28)
HIGHLIGHT = RGBColor(0xE6, 0xB6, 0x00)

WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY = RGBColor(0xF5, 0xF5, 0xF5)
MID_GRAY = RGBColor(0x95, 0xA5, 0xA6)
CHARCOAL = RGBColor(0x2C, 0x3E, 0x50)

GREEN_OK = RGBColor(0x2E, 0x7D, 0x32)
YELLOW_GAP = RGBColor(0xF9, 0xA8, 0x25)
RED_CRITICAL = RGBColor(0xBC, 0x43, 0x28)

FONT = "Helvetica Neue"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)


# ========== HELPERS ==========

def set_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_text(slide, left, top, width, height, text, *,
             size=11, bold=False, color=CHARCOAL, font=FONT,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    lines = text.split("\n") if isinstance(text, str) else text
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        p.line_spacing = 1.15
        run = p.add_run()
        run.text = line
        run.font.name = font
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = color
    return tb


def add_rect(slide, left, top, width, height, fill_color, *,
             line_color=None, line_weight=None, corner=False):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if corner else MSO_SHAPE.RECTANGLE
    shp = slide.shapes.add_shape(shape_type, left, top, width, height)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill_color
    if line_color is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line_color
        if line_weight:
            shp.line.width = line_weight
    shp.shadow.inherit = False
    return shp


def add_action_title(slide, title_text):
    add_text(slide, Inches(0.5), Inches(0.35), Inches(12.33), Inches(0.9),
             title_text, size=18, bold=True, color=NAVY, font=FONT)
    add_rect(slide, Inches(0.5), Inches(1.25), Inches(12.33), Inches(0.015), LIGHT_GRAY)


def add_takeaway(slide, text, color=NAVY):
    add_rect(slide, Inches(0.5), Inches(6.55), Inches(12.33), Inches(0.55), color)
    add_text(slide, Inches(0.75), Inches(6.60), Inches(11.83), Inches(0.45),
             text, size=11, bold=False, color=WHITE, font=FONT,
             anchor=MSO_ANCHOR.MIDDLE)


def add_footer(slide, page_label):
    add_text(slide, Inches(10.0), Inches(7.18), Inches(2.83), Inches(0.25),
             page_label, size=9, color=MID_GRAY, font=FONT, align=PP_ALIGN.RIGHT)


def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


# ========== NAV STRIP ==========

NAV_SECTIONS = [
    ("01", "SETUP", "How to run"),
    ("02", "CAPACITY & RESOURCE", "Blocks 1–2"),
    ("03", "REPORTING & METRICS", "Blocks 3–4"),
    ("04", "SYNERGY", "Block 5"),
    ("05", "OUTPUTS", "Block 6"),
]


def add_nav_strip(slide, current_idx):
    if current_idx is None:
        return
    strip_y = 0.08
    strip_h = 0.22
    pill_gap = 0.05
    total_w = 12.33
    n = len(NAV_SECTIONS)
    pill_w = (total_w - (n - 1) * pill_gap) / n
    for i, (num, name, _sub) in enumerate(NAV_SECTIONS):
        x = 0.5 + i * (pill_w + pill_gap)
        is_current = (i == current_idx)
        bg = NAVY if is_current else LIGHT_GRAY
        fg = WHITE if is_current else MID_GRAY
        add_rect(slide, Inches(x), Inches(strip_y), Inches(pill_w), Inches(strip_h), bg, corner=True)
        add_text(slide, Inches(x), Inches(strip_y), Inches(pill_w), Inches(strip_h),
                 f"{num} · {name}",
                 size=8, bold=is_current, color=fg, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)


def _nav(prs, idx):
    add_nav_strip(prs.slides[-1], idx)


# ========== BLOCK RENDERER ==========

def render_block(slide, *, block_num, block_name, duration, dynamic_name, dynamic_format,
                 individual_questions, synergy_questions, materials, output_spec):
    """Canonical block layout — used across all 6 interview blocks.

    Top banner: block number + name + duration + dynamic type.
    Two columns: INDIVIDUAL QUESTIONS (similar format) + SYNERGY QUESTIONS (cross-topic).
    Bottom banner: expected output + materials.
    """
    # Banner
    banner_y = 1.42
    banner_h = 0.75
    add_rect(slide, Inches(0.5), Inches(banner_y), Inches(12.33), Inches(banner_h), NAVY)
    # Block number badge
    add_rect(slide, Inches(0.5), Inches(banner_y), Inches(1.1), Inches(banner_h), ACCENT)
    add_text(slide, Inches(0.5), Inches(banner_y), Inches(1.1), Inches(banner_h),
             f"BLOCK\n{block_num}", size=14, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # Block title + dynamic
    add_text(slide, Inches(1.80), Inches(banner_y + 0.08), Inches(7.8), Inches(0.30),
             block_name, size=14, bold=True, color=WHITE, font=FONT)
    add_text(slide, Inches(1.80), Inches(banner_y + 0.42), Inches(7.8), Inches(0.28),
             f"Dynamic:  {dynamic_name}  —  {dynamic_format}",
             size=9, italic=True, color=BLUE, font=FONT)
    # Duration pill
    add_rect(slide, Inches(10.80), Inches(banner_y + 0.18), Inches(1.9), Inches(0.38),
             BLUE, corner=True)
    add_text(slide, Inches(10.80), Inches(banner_y + 0.18), Inches(1.9), Inches(0.38),
             f"⏱  {duration}", size=10, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Two columns: individual + synergy
    col_top = 2.32
    col_h = 3.85
    left_col_x = 0.5
    right_col_x = 6.80
    col_w = 6.1

    # Left column: individual questions
    add_rect(slide, Inches(left_col_x), Inches(col_top), Inches(col_w), Inches(0.40), BLUE)
    add_text(slide, Inches(left_col_x + 0.20), Inches(col_top), Inches(col_w - 0.40), Inches(0.40),
             "INDIVIDUAL QUESTIONS  —  same format, run as a batch",
             size=10, bold=True, color=WHITE, font=FONT, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(slide, Inches(left_col_x), Inches(col_top + 0.40), Inches(col_w), Inches(col_h - 0.40),
             LIGHT_GRAY)
    q_text = "\n".join([f"{i+1}.  {q}" for i, q in enumerate(individual_questions)])
    add_text(slide, Inches(left_col_x + 0.25), Inches(col_top + 0.55),
             Inches(col_w - 0.50), Inches(col_h - 0.70),
             q_text, size=10, color=CHARCOAL, font=FONT)

    # Right column: synergy questions
    add_rect(slide, Inches(right_col_x), Inches(col_top), Inches(col_w), Inches(0.40), ACCENT)
    add_text(slide, Inches(right_col_x + 0.20), Inches(col_top), Inches(col_w - 0.40), Inches(0.40),
             "SYNERGY QUESTIONS  —  connect to other blocks / topics",
             size=10, bold=True, color=WHITE, font=FONT, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(slide, Inches(right_col_x), Inches(col_top + 0.40), Inches(col_w), Inches(col_h - 0.40),
             LIGHT_GRAY)
    s_text = "\n".join([f"→  {q}" for q in synergy_questions])
    add_text(slide, Inches(right_col_x + 0.25), Inches(col_top + 0.55),
             Inches(col_w - 0.50), Inches(col_h - 0.70),
             s_text, size=10, color=CHARCOAL, font=FONT, italic=True)

    # Bottom strip: materials + output
    bot_y = col_top + col_h + 0.12
    bot_h = 0.35
    # Materials
    add_rect(slide, Inches(0.5), Inches(bot_y), Inches(6.1), Inches(bot_h), NAVY)
    add_text(slide, Inches(0.70), Inches(bot_y), Inches(1.4), Inches(bot_h),
             "MATERIALS", size=8, bold=True, color=BLUE, font=FONT,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Inches(2.10), Inches(bot_y), Inches(4.35), Inches(bot_h),
             materials, size=9, color=WHITE, font=FONT, anchor=MSO_ANCHOR.MIDDLE)
    # Output
    add_rect(slide, Inches(6.80), Inches(bot_y), Inches(6.03), Inches(bot_h), ACCENT)
    add_text(slide, Inches(7.00), Inches(bot_y), Inches(1.4), Inches(bot_h),
             "OUTPUT", size=8, bold=True, color=WHITE, font=FONT,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Inches(8.40), Inches(bot_y), Inches(4.38), Inches(bot_h),
             output_spec, size=9, color=WHITE, font=FONT, anchor=MSO_ANCHOR.MIDDLE)


# ========== SLIDES ==========

def slide_cover(prs):
    slide = blank(prs)
    set_bg(slide, NAVY)
    add_rect(slide, Inches(0.75), Inches(0.75), Inches(0.04), Inches(6), BLUE)
    add_text(slide, Inches(1.1), Inches(0.85), Inches(11.5), Inches(0.3),
             "AVANGRID  —  SAP DIGITAL PRODUCTS  •  APPLICATIONS BRANCH",
             size=10, bold=True, color=BLUE, font=FONT)
    add_text(slide, Inches(1.1), Inches(1.2), Inches(11.5), Inches(0.3),
             "Wave 1 Discovery  •  Capacity & Resource  ×  Reporting & Metrics  •  April 2026",
             size=10, color=MID_GRAY, font=FONT)
    add_rect(slide, Inches(1.1), Inches(3.0), Inches(0.5), Inches(0.05), ACCENT)
    add_text(slide, Inches(1.1), Inches(3.15), Inches(11.5), Inches(1.3),
             "Interview Canvas", size=54, bold=True, color=WHITE, font=FONT)
    add_text(slide, Inches(1.1), Inches(4.45), Inches(11.5), Inches(0.5),
             "Live workshop dynamic for Carlos Garbin's Applications team",
             size=18, color=BLUE, font=FONT)
    add_text(slide, Inches(1.1), Inches(5.10), Inches(11.5), Inches(1.4),
             "Six timeboxed blocks — each runs a single dynamic format (mapping · ranking · inventory · prioritization · scenario).\n"
             "Individual questions follow the same format so the rhythm never breaks; synergy questions connect topics so\n"
             "capacity evidence informs the reporting design and vice-versa. Total runtime: ~90 minutes.",
             size=12, color=WHITE, font=FONT, italic=True)
    add_rect(slide, Inches(1.1), Inches(6.7), Inches(0.3), Inches(0.03), BLUE)
    add_text(slide, Inches(1.1), Inches(6.82), Inches(6), Inches(0.3),
             "NTT DATA  —  Consulting", size=10, bold=True, color=WHITE, font=FONT)
    add_text(slide, Inches(10.5), Inches(6.82), Inches(2.5), Inches(0.3),
             "8 slides  ·  6 blocks", size=9, color=MID_GRAY, font=FONT, align=PP_ALIGN.RIGHT)
    add_text(slide, Inches(10.5), Inches(7.12), Inches(2.5), Inches(0.25),
             "Version 1.0", size=8, color=MID_GRAY, font=FONT, align=PP_ALIGN.RIGHT, italic=True)


def slide_how_to_run(prs):
    slide = blank(prs)
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Setup — five rules that keep the 90-minute dynamic focused, symmetric across blocks, and output-driven")

    # Left: Five rules of engagement
    add_text(slide, Inches(0.5), Inches(1.50), Inches(6.1), Inches(0.35),
             "FIVE RULES OF ENGAGEMENT", size=11, bold=True, color=ACCENT, font=FONT)
    rules = [
        ("01",
         "One format per block",
         "Each block runs a single dynamic (mapping, ranking, inventory, prioritization, scenarios). Individual questions share that format so answers fit on the same canvas and compare side-by-side."),
        ("02",
         "Symmetry: individual vs. synergy",
         "Every block has INDIVIDUAL questions (same-format batch) on the left and SYNERGY questions (cross-topic) on the right. Synergy reveals how Capacity decisions land in Reporting — and vice-versa."),
        ("03",
         "Timebox aggressively",
         "15 minutes per block. Stop when time runs out, even mid-discussion — un-answered questions become backlog for a second session rather than breaking rhythm."),
        ("04",
         "Evidence beats opinion",
         "Prefer numbers, artifacts, screenshots. If the answer is \"I think…\", tag it as HYPOTHESIS and mark the data source needed to confirm."),
        ("05",
         "Exit with artifacts, not minutes",
         "Every block produces a concrete output (map, ranked list, inventory table, MoSCoW grid, scenario walk-through). The debrief reads artifacts, not a transcript."),
    ]
    y = 1.90
    for num, title, desc in rules:
        add_rect(slide, Inches(0.5), Inches(y), Inches(0.50), Inches(0.80), NAVY)
        add_text(slide, Inches(0.5), Inches(y), Inches(0.50), Inches(0.80),
                 num, size=14, bold=True, color=WHITE, font=FONT,
                 align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, Inches(1.10), Inches(y), Inches(5.4), Inches(0.30),
                 title, size=11, bold=True, color=NAVY, font=FONT)
        add_text(slide, Inches(1.10), Inches(y + 0.28), Inches(5.4), Inches(0.55),
                 desc, size=9, color=CHARCOAL, font=FONT)
        y += 0.90

    # Right: 90-minute agenda
    add_text(slide, Inches(6.85), Inches(1.50), Inches(6.0), Inches(0.35),
             "90-MINUTE AGENDA", size=11, bold=True, color=NAVY, font=FONT)

    agenda = [
        ("00:00 – 00:05", "Opening", "Goals + rules + assign note-taker", BLUE),
        ("00:05 – 00:20", "Block 1", "Capacity Landscape  (Mural mapping)", ACCENT),
        ("00:20 – 00:35", "Block 2", "Resource Mix & Cost  (Dot voting)", ACCENT),
        ("00:35 – 00:50", "Block 3", "Reporting Inventory  (Kanban audit)", BLUE),
        ("00:50 – 01:05", "Block 4", "Metrics That Matter  (MoSCoW ranking)", BLUE),
        ("01:05 – 01:20", "Block 5", "Synergy Scenarios  (Scenario walk-through)", ACCENT),
        ("01:20 – 01:30", "Block 6", "Outputs  (Artifact review + commitments)", NAVY),
    ]
    y = 1.90
    for slot, label, title, color in agenda:
        add_rect(slide, Inches(6.85), Inches(y), Inches(0.10), Inches(0.50), color)
        add_text(slide, Inches(7.05), Inches(y + 0.02), Inches(1.5), Inches(0.25),
                 slot, size=8, bold=True, color=color, font=FONT)
        add_text(slide, Inches(7.05), Inches(y + 0.25), Inches(1.5), Inches(0.25),
                 label, size=8, color=MID_GRAY, font=FONT)
        add_text(slide, Inches(8.70), Inches(y + 0.02), Inches(4.2), Inches(0.50),
                 title, size=10, color=CHARCOAL, font=FONT, anchor=MSO_ANCHOR.MIDDLE)
        y += 0.55

    # Facilitation roles
    fac_y = 5.95
    add_rect(slide, Inches(6.85), Inches(fac_y), Inches(6.0), Inches(0.52), LIGHT_GRAY)
    add_rect(slide, Inches(6.85), Inches(fac_y), Inches(0.08), Inches(0.52), NAVY)
    add_text(slide, Inches(7.05), Inches(fac_y + 0.03), Inches(5.8), Inches(0.22),
             "ROLES", size=8, bold=True, color=NAVY, font=FONT)
    add_text(slide, Inches(7.05), Inches(fac_y + 0.25), Inches(5.8), Inches(0.25),
             "Facilitator · Timekeeper · Note-taker · Evidence-tagger (marks HYPOTHESIS vs. EVIDENCE)",
             size=9, color=CHARCOAL, font=FONT)

    add_takeaway(slide,
        "Rules and agenda are non-negotiable — the dynamic's value comes from rhythm and symmetry; broken rhythm makes the outputs non-comparable",
        color=NAVY)
    add_footer(slide, "Setup  |  02 / 08")


def slide_block_1_capacity_landscape(prs):
    slide = blank(prs)
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Block 1 — Capacity Landscape:  map where the ~26 FTEs of the Applications branch actually spend their time today")
    render_block(slide,
        block_num="01",
        block_name="CAPACITY LANDSCAPE",
        duration="15 min",
        dynamic_name="Mural Mapping",
        dynamic_format="each participant places post-its on a 3×3 grid — BAU vs. Strategic (x-axis) × Internal vs. Vendor vs. SWF (y-axis)",
        individual_questions=[
            "Quantos FTEs seu sub-time tem hoje  (internos · SWF · Accenture · vendor) ?",
            "Quanto % do seu time está em BAU / incidentes vs. strategic work ?",
            "Qual é seu % de utilização médio nos últimos 3 meses ?",
            "Quantas vagas abertas e quanto tempo médio para preencher ?",
            "Quanto tempo de KT quando alguém sai  (onboarding + ramp-up) ?",
            "Quem são os pontos de key-person risk  (conhecimento concentrado) ?",
        ],
        synergy_questions=[
            "Gargalos que NÃO são de pessoas — licenças · acessos · vendor availability?",
            "Quando regulatório entra, o que é reprogramado e quem decide?",
            "Existe alguma capacidade ociosa hoje que o reporting não mostra?",
            "Se você dobrasse o time, qual stage do demand lifecycle ficaria melhor primeiro?",
        ],
        materials="Mural board · post-its coloridos · cronômetro",
        output_spec="Capacity map 3×3 por sub-time  +  lista de key-person risks",
    )
    add_takeaway(slide,
        "A Capacity Landscape é o baseline quantitativo — sem ele, toda a conversa de reporting e metrics vira opinião",
        color=NAVY)
    add_footer(slide, "Block 1  —  Capacity Landscape  |  03 / 08")


def slide_block_2_resource_mix(prs):
    slide = blank(prs)
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Block 2 — Resource Mix & Cost:  expose the economics behind staffing decisions — when do you hire, when do you buy, when do you outsource?")
    render_block(slide,
        block_num="02",
        block_name="RESOURCE MIX & COST",
        duration="15 min",
        dynamic_name="Dot Voting + Cost Ranking",
        dynamic_format="cada participante coloca 3 dots nas 5 mais decisões que travaram ou aceleraram entrega — depois ranqueia custos por categoria",
        individual_questions=[
            "Distribuição de custo  %  entre internos · SWF · Accenture · Vendor  no seu sub-time?",
            "Quem decide quando usar SWF  vs.  contratar interno  vs.  vendor ?",
            "Como é feita a estimativa de esforço — quem estima, com que base?",
            "Você tem visibilidade de burn rate  (mensal · por projeto) ?",
            "Qual é o custo médio por ticket · por PDD · por release ?",
            "Quais decisões de mix você gostaria de voltar atrás nos últimos 12 meses?",
        ],
        synergy_questions=[
            "A decisão de mix depende de quê  (criticality · tech · duration · regulatory)?",
            "Como você prova ROI de uma expansão de capacity para Carlos / Vitor hoje?",
            "Vendor controla timeline ou você controla — e o que o reporting mostra?",
            "Se cortassem 20% do budget amanhã, o que sairia primeiro — e por quê?",
        ],
        materials="Template de ranking  (Excel / Mural)  ·  dots coloridos  ·  dados de custo se disponíveis",
        output_spec="Ranked list das 5 decisões de mix mais impactantes  +  gaps de visibilidade de custo",
    )
    add_takeaway(slide,
        "Mix & Cost é onde a conversa de capacity vira decisão financeira — sem dados aqui, o TOM não justifica investimento",
        color=ACCENT)
    add_footer(slide, "Block 2  —  Resource Mix & Cost  |  04 / 08")


def slide_block_3_reporting_inventory(prs):
    slide = blank(prs)
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Block 3 — Reporting Inventory:  catalogue every dashboard, report, and spreadsheet that circulates today — and separate used from ornamental")
    render_block(slide,
        block_num="03",
        block_name="REPORTING INVENTORY",
        duration="15 min",
        dynamic_name="Kanban-Style Audit",
        dynamic_format="cada relatório / dashboard vira um card em 3 colunas — USED · ORNAMENTAL · MISSING. Cada participante coloca até 5 cards por coluna",
        individual_questions=[
            "Quais relatórios · dashboards você gera ou recebe hoje?",
            "Quem consome cada um  (Carlos · Vitor · business · auditor)  e em que cadência?",
            "Onde os dados moram — Jira · IT Now · PDD workbook · Excel · SharePoint?",
            "Quanto tempo  h / mês  você gasta consolidando relatório manualmente?",
            "Você confia nos números que reporta? Se não, por quê?",
            "Qual relatório seu time gera mas NINGUÉM lê?",
        ],
        synergy_questions=[
            "Quais números do reporting atual refletem a Capacity real do Bloco 1?",
            "Há relatório hoje que rastreia vendor performance  (Bloco 2)? Se sim, quão granular?",
            "Se o TOM trouxesse SSO / dashboards automáticos, qual relatório sumiria primeiro?",
            "Qual dado você REPORTA mas não USA na sua decisão — e qual USA mas não reporta?",
        ],
        materials="Template Kanban  (3 colunas)  ·  cards vazios  ·  screenshots de dashboards existentes se possível",
        output_spec="Inventário categorizado  (USED · ORNAMENTAL · MISSING)  +  manual-work hour estimate por relatório",
    )
    add_takeaway(slide,
        "O inventário quase sempre expõe trabalho duplicado e números em que ninguém confia — essa é a munição para redesenhar o reporting no TOM",
        color=NAVY)
    add_footer(slide, "Block 3  —  Reporting Inventory  |  05 / 08")


def slide_block_4_metrics_that_matter(prs):
    slide = blank(prs)
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Block 4 — Metrics That Matter:  distinguish between what you MEASURE, what you SHOULD measure, and what you REPORT to leadership today")
    render_block(slide,
        block_num="04",
        block_name="METRICS THAT MATTER",
        duration="15 min",
        dynamic_name="MoSCoW Prioritization",
        dynamic_format="lista de métricas candidatas — cada participante classifica em MUST · SHOULD · COULD · WON'T. Debate em cima dos MUSTs para chegar a 3–5 métricas de time",
        individual_questions=[
            "Quais 3 KPIs são mais importantes para o seu sub-time hoje?",
            "Quais KPIs você efetivamente reporta para Carlos · Vitor?",
            "Qual métrica mostra a saúde do seu time  (velocidade · backlog · tickets)?",
            "Você mede Lead Time · Cycle Time · Throughput · WIP · rework? Quais?",
            "Como você define  \"trimestre bem-sucedido\"  para o seu sub-time?",
            "Qual métrica atual você tem certeza que NÃO reflete a realidade?",
        ],
        synergy_questions=[
            "Se só pudesse reportar 1 métrica para o board, qual seria — e por que ela?",
            "Que métrica ligaria Capacity  (Bloco 1)  com Reporting  (Bloco 3)  hoje?",
            "Qual métrica você gostaria que existisse mas a instrumentação impede?",
            "Métricas que o time USA  ≠  métricas que a liderança PEDE — onde está a lacuna?",
        ],
        materials="MoSCoW grid  ·  lista candidata de métricas pré-populada  ·  markers",
        output_spec="3–5 métricas MUST por sub-time  +  lista de métricas WON'T justificadas",
    )
    add_takeaway(slide,
        "A métrica que sobrevive ao MoSCoW é a que deve ser instrumentada primeiro — as demais entram no backlog do TOM",
        color=ACCENT)
    add_footer(slide, "Block 4  —  Metrics That Matter  |  06 / 08")


def slide_block_5_synergy_scenarios(prs):
    slide = blank(prs)
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Block 5 — Synergy Scenarios:  stress-test the Capacity × Reporting connection with three real scenarios — the gap between ideal response and today's is the TOM backlog")

    # Banner
    banner_y = 1.42
    banner_h = 0.75
    add_rect(slide, Inches(0.5), Inches(banner_y), Inches(12.33), Inches(banner_h), NAVY)
    add_rect(slide, Inches(0.5), Inches(banner_y), Inches(1.1), Inches(banner_h), ACCENT)
    add_text(slide, Inches(0.5), Inches(banner_y), Inches(1.1), Inches(banner_h),
             "BLOCK\n05", size=14, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Inches(1.80), Inches(banner_y + 0.08), Inches(7.8), Inches(0.30),
             "SYNERGY SCENARIOS", size=14, bold=True, color=WHITE, font=FONT)
    add_text(slide, Inches(1.80), Inches(banner_y + 0.42), Inches(7.8), Inches(0.28),
             "Dynamic:  Scenario Walk-through  —  group narrates step-by-step: what they'd do, what data they'd need, what they wouldn't have",
             size=9, italic=True, color=BLUE, font=FONT)
    add_rect(slide, Inches(10.80), Inches(banner_y + 0.18), Inches(1.9), Inches(0.38),
             BLUE, corner=True)
    add_text(slide, Inches(10.80), Inches(banner_y + 0.18), Inches(1.9), Inches(0.38),
             "⏱  15 min", size=10, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    scenarios = [
        {
            "num": "S1",
            "trigger": "Regulatory spike",
            "prompt": "Um pico de demanda regulatória inesperado chega hoje (ex: FERC / compliance mandate).",
            "probe": [
                "Quanto tempo até você enxergar o impacto em capacity?",
                "Qual relatório te alertaria — e qual não existe e deveria?",
                "Quem é reprogramado — e o reporting conta esse trade-off para Carlos?",
            ],
        },
        {
            "num": "S2",
            "trigger": "Hiring request",
            "prompt": "Vanessa  (Global SAP)  pede mais 2 FTEs. Vitor pede justificativa em 48h.",
            "probe": [
                "Que dados de capacity você tem hoje para justificar?",
                "Que dados FALTAM e teriam que ser coletados manualmente agora?",
                "Se o TOM estivesse rodando, qual dashboard resolveria isso em 15 min?",
            ],
        },
        {
            "num": "S3",
            "trigger": "Executive cost question",
            "prompt": "Carlos pergunta:  \"Quanto custa rodar a branch Applications este trimestre, por sub-time?\"",
            "probe": [
                "Você consegue responder hoje — em quanto tempo · com que precisão?",
                "Quais dos 4 custos  (int / SWF / Accenture / Vendor)  são opacos?",
                "Qual é a menor mudança de reporting que resolveria essa pergunta para sempre?",
            ],
        },
    ]

    # 3 scenario cards side-by-side
    card_w = 4.08
    card_h = 3.85
    gap = 0.06
    y0 = 2.32
    for i, sc in enumerate(scenarios):
        x = 0.5 + i * (card_w + gap)
        add_rect(slide, Inches(x), Inches(y0), Inches(card_w), Inches(card_h), LIGHT_GRAY)
        # Top scenario header
        add_rect(slide, Inches(x), Inches(y0), Inches(card_w), Inches(0.55), ACCENT)
        add_text(slide, Inches(x + 0.20), Inches(y0 + 0.05), Inches(1.0), Inches(0.20),
                 sc["num"], size=9, bold=True, color=WHITE, font=FONT)
        add_text(slide, Inches(x + 0.20), Inches(y0 + 0.25), Inches(card_w - 0.40), Inches(0.25),
                 sc["trigger"], size=13, bold=True, color=WHITE, font=FONT)
        # Trigger prompt
        add_text(slide, Inches(x + 0.20), Inches(y0 + 0.70), Inches(card_w - 0.40), Inches(0.22),
                 "TRIGGER", size=8, bold=True, color=ACCENT, font=FONT)
        add_text(slide, Inches(x + 0.20), Inches(y0 + 0.92), Inches(card_w - 0.40), Inches(1.05),
                 sc["prompt"], size=10, color=CHARCOAL, font=FONT, italic=True)
        # Probe questions
        add_text(slide, Inches(x + 0.20), Inches(y0 + 2.05), Inches(card_w - 0.40), Inches(0.22),
                 "PROBE", size=8, bold=True, color=NAVY, font=FONT)
        probe_text = "\n".join([f"▸  {p}" for p in sc["probe"]])
        add_text(slide, Inches(x + 0.20), Inches(y0 + 2.30), Inches(card_w - 0.40), Inches(1.45),
                 probe_text, size=9, color=CHARCOAL, font=FONT)

    # Bottom strip
    bot_y = y0 + card_h + 0.12
    bot_h = 0.35
    add_rect(slide, Inches(0.5), Inches(bot_y), Inches(6.1), Inches(bot_h), NAVY)
    add_text(slide, Inches(0.70), Inches(bot_y), Inches(1.4), Inches(bot_h),
             "MATERIALS", size=8, bold=True, color=BLUE, font=FONT, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Inches(2.10), Inches(bot_y), Inches(4.35), Inches(bot_h),
             "Flipchart  ·  scribe capturing \"ideal  vs.  today\"  gaps",
             size=9, color=WHITE, font=FONT, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(slide, Inches(6.80), Inches(bot_y), Inches(6.03), Inches(bot_h), ACCENT)
    add_text(slide, Inches(7.00), Inches(bot_y), Inches(1.4), Inches(bot_h),
             "OUTPUT", size=8, bold=True, color=WHITE, font=FONT, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Inches(8.40), Inches(bot_y), Inches(4.38), Inches(bot_h),
             "Gap list  per scenario  →  TOM instrumentation backlog",
             size=9, color=WHITE, font=FONT, anchor=MSO_ANCHOR.MIDDLE)

    add_takeaway(slide,
        "Scenarios force the group to SIMULATE — the gap between what they'd want to do and what they can do today is the TOM's measurable value",
        color=NAVY)
    add_footer(slide, "Block 5  —  Synergy Scenarios  |  07 / 08")


def slide_block_6_outputs(prs):
    slide = blank(prs)
    set_bg(slide, WHITE)
    add_action_title(slide,
        "Block 6 — Outputs & Commitments:  close the session with named artifacts, owners, and a dated follow-up — otherwise the dynamic evaporates")

    # Banner
    banner_y = 1.42
    banner_h = 0.75
    add_rect(slide, Inches(0.5), Inches(banner_y), Inches(12.33), Inches(banner_h), NAVY)
    add_rect(slide, Inches(0.5), Inches(banner_y), Inches(1.1), Inches(banner_h), NAVY)
    add_text(slide, Inches(0.5), Inches(banner_y), Inches(1.1), Inches(banner_h),
             "BLOCK\n06", size=14, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Inches(1.80), Inches(banner_y + 0.08), Inches(7.8), Inches(0.30),
             "OUTPUTS & COMMITMENTS", size=14, bold=True, color=WHITE, font=FONT)
    add_text(slide, Inches(1.80), Inches(banner_y + 0.42), Inches(7.8), Inches(0.28),
             "Dynamic:  Artifact Review  —  walk every wall-output, name an owner, set a follow-up date",
             size=9, italic=True, color=BLUE, font=FONT)
    add_rect(slide, Inches(10.80), Inches(banner_y + 0.18), Inches(1.9), Inches(0.38),
             BLUE, corner=True)
    add_text(slide, Inches(10.80), Inches(banner_y + 0.18), Inches(1.9), Inches(0.38),
             "⏱  10 min", size=10, bold=True, color=WHITE, font=FONT,
             align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Artifacts table
    add_text(slide, Inches(0.5), Inches(2.40), Inches(12.33), Inches(0.30),
             "EXPECTED ARTIFACTS  (one per block)", size=10, bold=True, color=NAVY, font=FONT)

    artifacts = [
        ("Block 1", "Capacity map 3×3 by sub-team", "Capacity Baseline snapshot", "NTT Analyst"),
        ("Block 2", "Ranked list of 5 mix decisions", "Cost visibility heatmap", "NTT + Carlos"),
        ("Block 3", "Kanban report inventory", "Used vs. Ornamental vs. Missing", "NTT + sub-team lead"),
        ("Block 4", "MoSCoW metrics grid", "3–5 MUST metrics per sub-team", "NTT + Carlos"),
        ("Block 5", "Scenario gap list", "TOM instrumentation backlog", "NTT"),
    ]

    header_y = 2.75
    col_widths = [1.2, 4.5, 4.2, 2.43]
    x_positions = [0.5]
    for w in col_widths[:-1]:
        x_positions.append(x_positions[-1] + w)
    headers = ["FROM", "ARTIFACT", "FEEDS INTO", "OWNER"]
    add_rect(slide, Inches(0.5), Inches(header_y), Inches(12.33), Inches(0.40), ACCENT)
    for i, h in enumerate(headers):
        add_text(slide, Inches(x_positions[i] + 0.15), Inches(header_y), Inches(col_widths[i]), Inches(0.40),
                 h, size=9, bold=True, color=WHITE, font=FONT, anchor=MSO_ANCHOR.MIDDLE)
    y = header_y + 0.45
    for i, (block, artifact, feeds, owner) in enumerate(artifacts):
        bg = LIGHT_GRAY if i % 2 == 0 else WHITE
        add_rect(slide, Inches(0.5), Inches(y), Inches(12.33), Inches(0.42), bg)
        add_text(slide, Inches(x_positions[0] + 0.15), Inches(y), Inches(col_widths[0]), Inches(0.42),
                 block, size=9, bold=True, color=NAVY, font=FONT, anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, Inches(x_positions[1] + 0.15), Inches(y), Inches(col_widths[1]), Inches(0.42),
                 artifact, size=9, color=CHARCOAL, font=FONT, anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, Inches(x_positions[2] + 0.15), Inches(y), Inches(col_widths[2]), Inches(0.42),
                 feeds, size=9, color=BLUE, font=FONT, italic=True, anchor=MSO_ANCHOR.MIDDLE)
        add_text(slide, Inches(x_positions[3] + 0.15), Inches(y), Inches(col_widths[3]), Inches(0.42),
                 owner, size=9, color=CHARCOAL, font=FONT, anchor=MSO_ANCHOR.MIDDLE)
        y += 0.44

    # Commitments / follow-up
    commit_y = 5.35
    add_rect(slide, Inches(0.5), Inches(commit_y), Inches(12.33), Inches(1.10), LIGHT_GRAY)
    add_rect(slide, Inches(0.5), Inches(commit_y), Inches(0.08), Inches(1.10), ACCENT)
    add_text(slide, Inches(0.70), Inches(commit_y + 0.08), Inches(12), Inches(0.28),
             "COMMITMENTS BEFORE WE LEAVE THE ROOM", size=10, bold=True, color=NAVY, font=FONT)
    add_text(slide, Inches(0.70), Inches(commit_y + 0.38), Inches(12), Inches(0.70),
             "▸  Every artifact has a named owner and a date to consolidate  (max 5 business days).\n"
             "▸  Open hypotheses tagged HYPOTHESIS get a data source assigned and a verification deadline.\n"
             "▸  Next session scheduled on the spot — second-pass to close un-answered questions carried over.",
             size=10, color=CHARCOAL, font=FONT)

    add_takeaway(slide,
        "A good workshop is measured 5 days later, not in the room — artifacts + owners + dates are the only way the dynamic becomes decision",
        color=NAVY)
    add_footer(slide, "Block 6  —  Outputs & Commitments  |  08 / 08")


# ========== MAIN ==========

def build():
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    slide_cover(prs)                                  # 01 — no nav (cover)
    slide_how_to_run(prs);           _nav(prs, 0)     # 02 SETUP
    slide_block_1_capacity_landscape(prs); _nav(prs, 1)  # 03 CAPACITY
    slide_block_2_resource_mix(prs); _nav(prs, 1)     # 04 CAPACITY
    slide_block_3_reporting_inventory(prs); _nav(prs, 2) # 05 REPORTING
    slide_block_4_metrics_that_matter(prs); _nav(prs, 2) # 06 REPORTING
    slide_block_5_synergy_scenarios(prs); _nav(prs, 3) # 07 SYNERGY
    slide_block_6_outputs(prs);       _nav(prs, 4)    # 08 OUTPUTS

    out = "/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Deliverables/Interview_Canvas_Applications/Interview_Canvas_Applications_v1.pptx"
    prs.save(out)
    print(f"Saved: {out}")
    print(f"Slides: {len(prs.slides)}")


if __name__ == "__main__":
    build()
