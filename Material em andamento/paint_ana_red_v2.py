"""
Fix: insert <a:solidFill> in CORRECT position (after <a:ln>, before everything else)
per OOXML schema CT_TextCharacterProperties. PowerPoint silently ignores out-of-order fills.
"""
from pptx import Presentation
from pathlib import Path
from lxml import etree

SRC = Path("/Users/gustavohenriquecastellano/Projeto-SAP-Digital-Products/Material em andamento/"
           "2604_AS-IS deliverable vrevisada - Final 3 (Ana integrated).pptx")
a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
p_ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": a_ns, "p": p_ns}
RED_HEX = "C00000"
prs = Presentation(str(SRC))


def color_run_red(run_elem, hex_color=RED_HEX):
    rPr = run_elem.find("a:rPr", NS)
    if rPr is None:
        rPr = etree.SubElement(run_elem, f"{{{a_ns}}}rPr")
        run_elem.insert(0, rPr)
    # Remove existing solidFill anywhere
    for sf in rPr.findall("a:solidFill", NS):
        rPr.remove(sf)
    # Build solidFill
    sf = etree.Element(f"{{{a_ns}}}solidFill")
    sc = etree.SubElement(sf, f"{{{a_ns}}}srgbClr")
    sc.set("val", hex_color)
    # Insert in correct OOXML order — after <a:ln> if present, otherwise as first child
    ln = rPr.find("a:ln", NS)
    if ln is not None:
        ln.addnext(sf)
    else:
        rPr.insert(0, sf)
    rPr.set("b", "1")


def shape_full_text(sp_xml):
    return "".join((t.text or "") for t in sp_xml.findall(".//a:t", NS))


def find_shape(slide, contains, occ=0):
    matches = [sp for sp in slide.element.findall(".//p:sp", NS) if contains in shape_full_text(sp)]
    return matches[occ] if occ < len(matches) else None


def paint_para(slide, shape_substr, target_substr):
    sh = find_shape(slide, shape_substr)
    if sh is None: return False
    for pa in sh.findall(".//a:p", NS):
        ptxt = "".join((t.text or "") for t in pa.findall(".//a:t", NS))
        if target_substr in ptxt:
            for r in pa.findall("a:r", NS):
                color_run_red(r)
            return True
    return False


targets = [
    (12, "Limited and slow-to-scale capacity", "Run and Build are blended"),
    (12, "Limited and slow-to-scale capacity", "Business-side roles missing"),
    (39, "Limited and slow-to-scale capacity", "Run and Build are blended"),
    (39, "Limited and slow-to-scale capacity", "Business-side roles missing"),
    (11, "These are the structural and operational limitations", "Run/Build blended"),
    (38, "These are the structural and operational limitations", "Run/Build blended"),
    (11, "Our assessment highlights opportunities", "six operating-model blocks"),
    (38, "Our assessment highlights opportunities", "six operating-model blocks"),
    (32, "No portfolio steering forum", "ERP Steering Committee"),
    (40, "demand is owned at the application level", "supporting roles that anchor"),
    (41, "Key leadership roles remain unfilled", "Run side is also undefined"),
    (14, "Defined roles (Product Owner / Product Manager / Tech Lead)", "domain-based organization"),
    (43, "Defined roles (Product Owner / Product Manager / Tech Lead)", "domain-based organization"),
]
for idx, sh_sub, tg_sub in targets:
    ok = paint_para(prs.slides[idx], sh_sub, tg_sub)
    print(f"Slide {idx+1} · {tg_sub[:35]:35} : {'OK' if ok else 'MISS'}")

prs.save(str(SRC))
print(f"\n✓ Saved (red fixed)")
