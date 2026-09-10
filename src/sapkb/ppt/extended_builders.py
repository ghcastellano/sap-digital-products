"""
Extended Slide Builders — Additional consulting slide types
============================================================
Matrix 2x2, Timeline, SWOT, Process Flow, Gantt, Harvey Balls, Comparison Table
"""
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

from .slide_builders import DeckBuilder, SLIDE_W, SLIDE_H, FONT


class ExtendedDeckBuilder(DeckBuilder):
    """Extends DeckBuilder with additional consulting slide types."""

    def add_matrix_2x2(self, title: str, x_axis: str, y_axis: str,
                       quadrants: list[dict], subtitle: str = None) -> None:
        """
        2x2 strategic matrix (Effort vs Impact, etc.)

        quadrants: [
            {'label': 'Quick Wins', 'items': ['Item 1', 'Item 2'], 'position': 'top_left'},
            {'label': 'Strategic', 'items': ['Item A'], 'position': 'top_right'},
            {'label': 'Deprioritize', 'items': [], 'position': 'bottom_left'},
            {'label': 'Reassess', 'items': ['Item X'], 'position': 'bottom_right'},
        ]
        """
        slide = self._blank()
        self._action_title(slide, title, subtitle)

        # Matrix area
        mx, my = Inches(2.5), Inches(1.5)
        mw, mh = Inches(8.5), Inches(5.0)
        half_w, half_h = mw // 2, mh // 2

        # Quadrant colors (light fills)
        q_colors = {
            "top_left": self.brand.tint_05,
            "top_right": self.brand.tint_15,
            "bottom_left": self.brand.light_gray,
            "bottom_right": self.brand.tint_30,
        }
        q_positions = {
            "top_left": (mx, my),
            "top_right": (mx + half_w, my),
            "bottom_left": (mx, my + half_h),
            "bottom_right": (mx + half_w, my + half_h),
        }

        for q in quadrants:
            pos = q.get("position", "top_left")
            x, y = q_positions[pos]
            color = q_colors.get(pos, self.brand.light_gray)

            # Quadrant box
            box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, half_w, half_h)
            box.fill.solid()
            box.fill.fore_color.rgb = color
            box.line.color.rgb = self.brand.mid_gray
            box.line.width = Pt(0.5)

            # Label
            lbl = slide.shapes.add_textbox(x + Inches(0.2), y + Inches(0.1),
                                           half_w - Inches(0.4), Inches(0.4))
            p = lbl.text_frame.paragraphs[0]
            p.text = q.get("label", "")
            p.font.size = Pt(12)
            p.font.bold = True
            p.font.color.rgb = self.brand.primary
            p.font.name = self.brand.font_family

            # Items
            items = q.get("items", [])
            if items:
                items_box = slide.shapes.add_textbox(
                    x + Inches(0.2), y + Inches(0.5),
                    half_w - Inches(0.4), half_h - Inches(0.6),
                )
                tf = items_box.text_frame
                tf.word_wrap = True
                for j, item in enumerate(items[:6]):
                    p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
                    p.text = f"• {item}"
                    p.font.size = Pt(9)
                    p.font.color.rgb = self.brand.charcoal
                    p.font.name = self.brand.font_family

        # Axis labels
        # X-axis
        x_lbl = slide.shapes.add_textbox(mx, my + mh + Inches(0.1), mw, Inches(0.3))
        p = x_lbl.text_frame.paragraphs[0]
        p.text = f"← Low {x_axis}                                                          High {x_axis} →"
        p.font.size = Pt(9)
        p.font.color.rgb = self.brand.dark_gray
        p.font.name = self.brand.font_family
        p.alignment = PP_ALIGN.CENTER

        # Y-axis
        y_lbl = slide.shapes.add_textbox(mx - Inches(2.0), my + half_h - Inches(0.15),
                                         Inches(1.8), Inches(0.3))
        p = y_lbl.text_frame.paragraphs[0]
        p.text = f"↑ High {y_axis}"
        p.font.size = Pt(9)
        p.font.color.rgb = self.brand.dark_gray
        p.font.name = self.brand.font_family
        p.alignment = PP_ALIGN.RIGHT

    def add_timeline(self, title: str, phases: list[dict],
                     subtitle: str = None) -> None:
        """
        Horizontal timeline with colored phase bars.

        phases: [
            {'name': 'Phase 1', 'duration': 'W1-W2', 'description': '...', 'deliverables': ['...']},
            ...
        ]
        """
        slide = self._blank()
        self._action_title(slide, title, subtitle)

        n = len(phases)
        bar_y = Inches(2.8)
        bar_h = Inches(0.6)
        total_w = Inches(11.5)
        phase_w = total_w // n
        start_x = Inches(0.8)

        colors = [self.brand.secondary, self.brand.primary, self.brand.shade_50,
                  self.brand.accent, self.brand.shade_70, self.brand.tint_30]

        # Timeline line
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                      start_x, bar_y + bar_h // 2 - Pt(2),
                                      total_w, Pt(4))
        line.fill.solid()
        line.fill.fore_color.rgb = self.brand.mid_gray
        line.line.fill.background()

        for i, phase in enumerate(phases):
            x = start_x + phase_w * i
            color = colors[i % len(colors)]

            # Phase bar
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x + Inches(0.05), bar_y,
                                         phase_w - Inches(0.1), bar_h)
            bar.fill.solid()
            bar.fill.fore_color.rgb = color
            bar.line.fill.background()
            tf = bar.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.text = phase["name"]
            p.font.size = Pt(11)
            p.font.bold = True
            p.font.color.rgb = self.brand.text_on(color)
            p.font.name = self.brand.font_family
            p.alignment = PP_ALIGN.CENTER

            # Duration label above
            dur = slide.shapes.add_textbox(x, bar_y - Inches(0.35), phase_w, Inches(0.3))
            p = dur.text_frame.paragraphs[0]
            p.text = phase.get("duration", "")
            p.font.size = Pt(8)
            p.font.color.rgb = self.brand.dark_gray
            p.font.name = self.brand.font_family
            p.alignment = PP_ALIGN.CENTER

            # Description below
            desc_y = bar_y + bar_h + Inches(0.15)
            desc = slide.shapes.add_textbox(x, desc_y, phase_w - Inches(0.1), Inches(1.5))
            tf = desc.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = phase.get("description", "")
            p.font.size = Pt(8)
            p.font.color.rgb = self.brand.charcoal
            p.font.name = self.brand.font_family

            # Deliverables
            for d in phase.get("deliverables", [])[:4]:
                p2 = tf.add_paragraph()
                p2.text = f"✓ {d}"
                p2.font.size = Pt(7)
                p2.font.color.rgb = self.brand.secondary
                p2.font.name = self.brand.font_family

    def add_swot(self, title: str, strengths: list[str], weaknesses: list[str],
                 opportunities: list[str], threats: list[str],
                 subtitle: str = None) -> None:
        """Four-quadrant SWOT analysis."""
        slide = self._blank()
        self._action_title(slide, title, subtitle)

        mx, my = Inches(0.8), Inches(1.4)
        qw, qh = Inches(5.65), Inches(2.7)
        gap = Inches(0.2)

        quadrants = [
            ("STRENGTHS", strengths, mx, my, self.brand.secondary),
            ("WEAKNESSES", weaknesses, mx + qw + gap, my, self.brand.accent),
            ("OPPORTUNITIES", opportunities, mx, my + qh + gap, self.brand.primary),
            ("THREATS", threats, mx + qw + gap, my + qh + gap, self.brand.shade_90),
        ]

        for label, items, x, y, color in quadrants:
            # Header bar
            hdr = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, qw, Inches(0.4))
            hdr.fill.solid()
            hdr.fill.fore_color.rgb = color
            hdr.line.fill.background()
            p = hdr.text_frame.paragraphs[0]
            p.text = label
            p.font.size = Pt(11)
            p.font.bold = True
            p.font.color.rgb = self.brand.text_on(color)
            p.font.name = self.brand.font_family
            hdr.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

            # Content box
            content = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                              x, y + Inches(0.4), qw, qh - Inches(0.4))
            content.fill.solid()
            content.fill.fore_color.rgb = self.brand.light_gray
            content.line.color.rgb = color
            content.line.width = Pt(1)

            # Items
            items_box = slide.shapes.add_textbox(
                x + Inches(0.15), y + Inches(0.5),
                qw - Inches(0.3), qh - Inches(0.6),
            )
            tf = items_box.text_frame
            tf.word_wrap = True
            for j, item in enumerate(items[:6]):
                p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
                p.text = f"• {item}"
                p.font.size = Pt(9)
                p.font.color.rgb = self.brand.charcoal
                p.font.name = self.brand.font_family
                p.space_after = Pt(4)

    def add_process_flow(self, title: str, steps: list[dict],
                         subtitle: str = None) -> None:
        """
        Horizontal process flow with chevron-like shapes.

        steps: [{'name': 'Step 1', 'description': '...'}, ...]
        """
        slide = self._blank()
        self._action_title(slide, title, subtitle)

        n = len(steps)
        start_x = Inches(0.8)
        total_w = Inches(11.5)
        step_w = total_w // n
        step_h = Inches(1.2)
        y = Inches(2.5)

        colors = [self.brand.primary, self.brand.secondary, self.brand.shade_50,
                  self.brand.accent, self.brand.shade_70]

        for i, step in enumerate(steps):
            x = start_x + step_w * i
            color = colors[i % len(colors)]

            # Step box (using pentagon/chevron shape)
            box = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, x, y, step_w - Inches(0.1), step_h)
            box.fill.solid()
            box.fill.fore_color.rgb = color
            box.line.fill.background()
            tf = box.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.text = step["name"]
            p.font.size = Pt(10)
            p.font.bold = True
            p.font.color.rgb = self.brand.text_on(color)
            p.font.name = self.brand.font_family
            p.alignment = PP_ALIGN.CENTER

            # Description below
            desc = slide.shapes.add_textbox(x, y + step_h + Inches(0.2),
                                            step_w - Inches(0.1), Inches(2.0))
            tf = desc.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = step.get("description", "")
            p.font.size = Pt(8)
            p.font.color.rgb = self.brand.charcoal
            p.font.name = self.brand.font_family

    def add_comparison_table(self, title: str, headers: list[str],
                              rows: list[dict], highlight_column: int | None = None,
                              subtitle: str = None, source: str = None) -> None:
        """
        Enhanced comparison table with optional highlight column.

        rows: [{'label': 'Row Name', 'values': ['val1', 'val2', ...]}, ...]
        """
        slide = self._blank()
        self._action_title(slide, title, subtitle)

        n_rows = len(rows) + 1
        n_cols = len(headers)
        table_w = Inches(11.5)
        col_w = table_w // n_cols

        ts = slide.shapes.add_table(n_rows, n_cols,
                                    Inches(0.8), Inches(1.3), table_w, Inches(0.35) * n_rows)
        table = ts.table

        # Headers
        for i, h in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = h
            bg = self.brand.accent if (highlight_column is not None and i == highlight_column) else self.brand.primary
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            for p in cell.text_frame.paragraphs:
                p.font.size = FONT['table_header']
                p.font.bold = True
                p.font.color.rgb = self.brand.white
                p.font.name = self.brand.font_family
                p.alignment = PP_ALIGN.CENTER
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE

        # Data rows
        for r_idx, row in enumerate(rows):
            values = [row.get("label", "")] + row.get("values", []) if "label" in row else row.get("values", [])
            for c_idx, val in enumerate(values[:n_cols]):
                cell = table.cell(r_idx + 1, c_idx)
                cell.text = str(val)
                bg = self.brand.white if r_idx % 2 == 0 else self.brand.light_gray
                if highlight_column is not None and c_idx == highlight_column:
                    bg = self.brand.tint_05
                cell.fill.solid()
                cell.fill.fore_color.rgb = bg
                for p in cell.text_frame.paragraphs:
                    p.font.size = FONT['table_body']
                    p.font.color.rgb = self.brand.charcoal
                    p.font.name = self.brand.font_family
                    p.alignment = PP_ALIGN.CENTER if c_idx > 0 else PP_ALIGN.LEFT
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE

        for i in range(n_cols):
            table.columns[i].width = int(col_w)

        if source:
            self._add_source(slide, source)

    def add_harvey_balls(self, title: str, categories: list[str],
                         items: list[dict], subtitle: str = None) -> None:
        """
        Maturity assessment grid with filled circles.

        items: [{'name': 'Capability A', 'scores': [100, 75, 50, 25, 0]}, ...]
        Scores: 0=empty, 25=quarter, 50=half, 75=three-quarter, 100=full
        """
        slide = self._blank()
        self._action_title(slide, title, subtitle)

        n_cols = len(categories) + 1  # +1 for item names
        n_rows = len(items) + 1  # +1 for header
        table_w = Inches(11.5)
        col_w = table_w // n_cols

        ts = slide.shapes.add_table(n_rows, n_cols,
                                    Inches(0.8), Inches(1.3), table_w, Inches(0.4) * n_rows)
        table = ts.table

        # Header
        headers = [""] + categories
        for i, h in enumerate(headers):
            cell = table.cell(0, i)
            cell.text = h
            cell.fill.solid()
            cell.fill.fore_color.rgb = self.brand.primary
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(9)
                p.font.bold = True
                p.font.color.rgb = self.brand.white
                p.font.name = self.brand.font_family
                p.alignment = PP_ALIGN.CENTER if i > 0 else PP_ALIGN.LEFT
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE

        # Data
        ball_chars = {0: "○", 25: "◔", 50: "◑", 75: "◕", 100: "●"}
        for r_idx, item in enumerate(items):
            # Name cell
            cell = table.cell(r_idx + 1, 0)
            cell.text = item["name"]
            bg = self.brand.white if r_idx % 2 == 0 else self.brand.light_gray
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(9)
                p.font.color.rgb = self.brand.charcoal
                p.font.name = self.brand.font_family

            # Score cells
            for c_idx, score in enumerate(item.get("scores", [])):
                cell = table.cell(r_idx + 1, c_idx + 1)
                cell.text = ball_chars.get(score, ball_chars[min(ball_chars.keys(), key=lambda k: abs(k - score))])
                cell.fill.solid()
                cell.fill.fore_color.rgb = bg
                for p in cell.text_frame.paragraphs:
                    p.font.size = Pt(16)
                    p.font.color.rgb = self.brand.primary
                    p.font.name = self.brand.font_family
                    p.alignment = PP_ALIGN.CENTER
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE

        for i in range(n_cols):
            table.columns[i].width = int(col_w)
