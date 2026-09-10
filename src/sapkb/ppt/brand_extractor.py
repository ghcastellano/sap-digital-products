"""
Brand Extractor — Extract dominant colors and fonts from existing PPTX files
"""
from collections import Counter
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor


def extract_brand_from_pptx(file_path: Path) -> dict:
    """
    Analyze a PPTX file to extract the dominant brand palette.

    Returns:
        {'primary': '#HEX', 'secondary': '#HEX', 'accent': '#HEX',
         'font': 'FontName', 'all_colors': Counter}
    """
    prs = Presentation(str(file_path))
    color_counter = Counter()
    font_counter = Counter()

    for slide in prs.slides:
        for shape in slide.shapes:
            # Extract fill colors from shapes
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        # Font tracking
                        if run.font.name:
                            font_counter[run.font.name] += 1
                        # Text color
                        try:
                            if run.font.color and run.font.color.type is not None and run.font.color.rgb:
                                hex_color = str(run.font.color.rgb)
                                # Skip pure black and white (too common)
                                if hex_color not in ("000000", "FFFFFF"):
                                    color_counter[hex_color] += 1
                        except Exception:
                            pass

            # Shape fill colors
            try:
                if hasattr(shape, 'fill') and shape.fill.type is not None:
                    fill = shape.fill
                    if fill.fore_color and fill.fore_color.type is not None:
                        hex_color = str(fill.fore_color.rgb)
                        if hex_color not in ("000000", "FFFFFF", "F5F5F5", "E7E6E6"):
                            color_counter[hex_color] += len(str(shape.width or 0))  # weight by size
            except Exception:
                pass

    # Group colors by hue family
    hue_groups = _group_by_hue(color_counter)

    # Top 3 hue groups = primary, secondary, accent
    top_groups = sorted(hue_groups.items(), key=lambda x: x[1]["total"], reverse=True)[:3]

    result = {
        "primary": f"#{top_groups[0][1]['representative']}" if len(top_groups) > 0 else "#002060",
        "secondary": f"#{top_groups[1][1]['representative']}" if len(top_groups) > 1 else "#0070C0",
        "accent": f"#{top_groups[2][1]['representative']}" if len(top_groups) > 2 else "#19A3FC",
        "font": font_counter.most_common(1)[0][0] if font_counter else "Arial",
        "all_colors": color_counter.most_common(20),
    }
    return result


def _hex_to_hsl(hex_str: str) -> tuple:
    """Convert hex color to HSL for hue grouping."""
    r = int(hex_str[0:2], 16) / 255.0
    g = int(hex_str[2:4], 16) / 255.0
    b = int(hex_str[4:6], 16) / 255.0

    mx = max(r, g, b)
    mn = min(r, g, b)
    l = (mx + mn) / 2.0

    if mx == mn:
        h = s = 0.0
    else:
        d = mx - mn
        s = d / (2.0 - mx - mn) if l > 0.5 else d / (mx + mn)
        if mx == r:
            h = (g - b) / d + (6 if g < b else 0)
        elif mx == g:
            h = (b - r) / d + 2
        else:
            h = (r - g) / d + 4
        h /= 6.0

    return (h * 360, s, l)


def _group_by_hue(color_counter: Counter) -> dict:
    """Group colors into hue families (30-degree buckets)."""
    groups = {}
    for hex_color, count in color_counter.items():
        try:
            h, s, l = _hex_to_hsl(hex_color)
            # Skip very gray colors (low saturation)
            if s < 0.1:
                continue
            bucket = int(h / 30) * 30
            if bucket not in groups:
                groups[bucket] = {"total": 0, "colors": [], "representative": hex_color}
            groups[bucket]["total"] += count
            groups[bucket]["colors"].append((hex_color, count))
            # Update representative to the most used color in this bucket
            if count > max(c for _, c in groups[bucket]["colors"][:-1], default=0):
                groups[bucket]["representative"] = hex_color
        except Exception:
            continue
    return groups
