"""Reorder slides in build_tom_deck_v2.py into v3 with logical flow."""
from pathlib import Path
import re

src = Path("build_tom_deck_v2.py").read_text()

# Find all slide block start positions
markers = [
    (m.start(), m.group(0))
    for m in re.finditer(r"# ════+\n# SLIDE.*\n# ════+", src)
]
end_marker = re.search(r"# ════+\nout_file = OUT", src)

# Header before first slide
header = src[:markers[0][0]]
# Footer (save block)
footer = src[end_marker.start():]

blocks = []
for i, (pos, _) in enumerate(markers):
    end = markers[i + 1][0] if i + 1 < len(markers) else end_marker.start()
    blocks.append(src[pos:end])

# Identify each block by content
def label(b):
    if "SLIDE 1 — COVER" in b: return "cover"
    if "SLIDE 2 — Executive" in b: return "execsum"
    if "SLIDE 3 — Current State" in b: return "current"
    if "SLIDE 4 — TOM Overview" in b: return "tom"
    if "SLIDE 5 — Stage 01" in b: return "stage01_overview"
    if "SLIDE 6 — CAPABILITIES" in b: return "cap"
    if "SLIDE 7 — ROLES" in b: return "roles"
    if "SLIDE 8 — TOOLS" in b: return "tools"
    if "SLIDE 9 — GOVERNANCE" in b: return "gov"
    if "SLIDE 10 — PERFORMANCE" in b: return "perf"
    if "SLIDE 11 — Roles & Responsibilities" in b: return "rr"
    if "SLIDE 12 — Forums" in b: return "forums"
    if "SLIDE 13 — Portfolio Architecture" in b: return "portfolio"
    if "SLIDE 14 — Portfolio Board mockup" in b: return "portfolio_mock"
    if "SLIDE 15 — Team Kanban" in b: return "team_kanban"
    if "SLIDE 16 — Class of Service" in b: return "cos"
    if "SLIDE 17 — Roadmap" in b: return "roadmap"
    if "Guiding Principles" in b: return "principles"
    if "Demand Management Practices" in b: return "practices"
    if "Step policies" in b: return "step_policies"
    return "unknown"

labeled = {label(b): b for b in blocks}

# New order
new_order = [
    "cover",
    "execsum",
    "current",
    "principles",       # MOVED UP
    "practices",        # MOVED UP
    "tom",
    "stage01_overview",
    "cap",
    "roles",
    "tools",
    "gov",
    "perf",
    "rr",
    "forums",
    "portfolio",
    "step_policies",    # MOVED UP
    "portfolio_mock",
    "team_kanban",
    "cos",
    "roadmap",
]

# Renumber the SLIDE comments
out_blocks = []
for i, name in enumerate(new_order, start=1):
    block = labeled[name]
    # Replace the slide number
    block = re.sub(r"# SLIDE \d+ — ", f"# SLIDE {i} — ", block)
    # If it doesn't have a number (the 3 new slides), insert one
    block = re.sub(r"# SLIDE — ", f"# SLIDE {i} — ", block)
    out_blocks.append(block)

# Update footer to v3 filename
footer = footer.replace("v2.0.pptx", "v3.0.pptx")

result = header + "\n".join(out_blocks) + footer
Path("build_tom_deck_v3.py").write_text(result)
print("✓ Wrote build_tom_deck_v3.py")
