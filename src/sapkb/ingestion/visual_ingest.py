"""
Visual Ingest — Claude Vision for slide description (Haiku for text, Sonnet for complex visuals)

Cost model:
- Haiku: ~$0.002/slide (text-heavy, simple layouts)
- Sonnet: ~$0.015/slide (diagrams, charts, Venn, flows, org charts)

The pipeline:
1. Render each slide as an image (via LibreOffice → PDF → pdftoppm, or python-pptx text fallback)
2. Classify complexity: text-heavy → Haiku, visually complex → Sonnet
3. Claude describes each slide as rich semantic text
4. The description becomes the chunk content (searchable, embeddable)
"""
import base64
import io
import os
import subprocess
import tempfile
from pathlib import Path
from typing import Optional

from PIL import Image

from ..models import DocumentChunk, ChunkMetadata, DocumentType


HAIKU = "claude-haiku-4-5-20251001"
SONNET = "claude-sonnet-4-6"

# Keywords that indicate a slide has complex visuals needing Sonnet
_COMPLEX_VISUAL_KEYWORDS = [
    "diagram", "chart", "graph", "flow", "venn", "sankey", "architecture",
    "org chart", "organogram", "matrix", "heatmap", "radar", "spider",
    "timeline", "roadmap", "gantt", "swimlane", "process map",
    "infographic", "dashboard", "wireframe", "mockup",
]

# Shape count threshold: slides with many shapes likely have complex visuals
_SHAPE_COUNT_THRESHOLD = 15


def _get_anthropic_client():
    """Lazy-load Anthropic client to avoid import cost on non-visual ingests."""
    import anthropic
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        raise RuntimeError(
            "ANTHROPIC_API_KEY not set. Required for visual ingest.\n"
            "Set it in .env or export ANTHROPIC_API_KEY=sk-ant-..."
        )
    return anthropic.Anthropic(api_key=api_key)


def _pil_to_base64(img: Image.Image, max_px: int = 1024) -> str:
    """Resize and convert PIL image to base64 JPEG."""
    img.thumbnail((max_px, max_px), Image.LANCZOS)
    buf = io.BytesIO()
    img.convert("RGB").save(buf, format="JPEG", quality=85)
    return base64.standard_b64encode(buf.getvalue()).decode()


def _render_slides_to_images(pptx_path: Path) -> list[Optional[Image.Image]]:
    """
    Render PPTX slides as images via LibreOffice → PDF → pdftoppm.
    Returns list of PIL Images (or None for failed slides).
    """
    images = []

    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = Path(tmpdir) / "slides.pdf"

        # Try LibreOffice conversion
        try:
            # Find soffice
            soffice_paths = [
                "/Applications/LibreOffice.app/Contents/MacOS/soffice",
                "soffice",
                "/usr/bin/soffice",
            ]
            soffice = None
            for sp in soffice_paths:
                try:
                    result = subprocess.run([sp, "--version"], capture_output=True, timeout=5)
                    if result.returncode == 0:
                        soffice = sp
                        break
                except (FileNotFoundError, subprocess.TimeoutExpired):
                    continue

            if soffice:
                subprocess.run(
                    [soffice, "--headless", "--convert-to", "pdf",
                     "--outdir", tmpdir, str(pptx_path)],
                    capture_output=True, timeout=120,
                )
                # Find the PDF (name matches input but with .pdf extension)
                pdfs = list(Path(tmpdir).glob("*.pdf"))
                if pdfs:
                    pdf_path = pdfs[0]

                    # Convert PDF to images with pdftoppm
                    result = subprocess.run(
                        ["pdftoppm", "-jpeg", "-r", "150", str(pdf_path),
                         str(Path(tmpdir) / "slide")],
                        capture_output=True, timeout=120,
                    )
                    if result.returncode == 0:
                        jpgs = sorted(Path(tmpdir).glob("slide-*.jpg"))
                        for jpg in jpgs:
                            images.append(Image.open(jpg).copy())
                        return images
        except Exception:
            pass

    # Fallback: no images available
    return []


def _classify_slide_complexity(slide, raw_text: str) -> str:
    """
    Classify a slide as 'text' (Haiku) or 'complex' (Sonnet).

    Uses heuristics:
    - Shape count > threshold → complex
    - Has charts/diagrams → complex
    - Has grouped shapes → complex
    - Contains complexity keywords → complex
    - Otherwise → text (Haiku)
    """
    shape_count = len(slide.shapes)
    text_lower = raw_text.lower()

    # Check for complex visual keywords in slide text
    for kw in _COMPLEX_VISUAL_KEYWORDS:
        if kw in text_lower:
            return "complex"

    # High shape count indicates visual complexity
    if shape_count >= _SHAPE_COUNT_THRESHOLD:
        return "complex"

    # Check for charts
    for shape in slide.shapes:
        if shape.has_chart:
            return "complex"
        # Check for grouped shapes (common in diagrams)
        if hasattr(shape, "shapes"):  # GroupShape
            return "complex"

    # Check for many images (infographic-like)
    image_count = sum(1 for s in slide.shapes if s.shape_type == 13)  # MSO_SHAPE_TYPE.PICTURE
    if image_count >= 3:
        return "complex"

    return "text"


def _describe_with_vision(client, b64_image: str, slide_num: int,
                          raw_text: str, model: str) -> str:
    """Use Claude Vision to describe a slide image."""
    prompt = f"""You are analyzing slide {slide_num} of a consulting presentation.

The raw text extracted from this slide is:
---
{raw_text[:2000]}
---

Describe ALL content visible in this slide image:
- Exact text (titles, bullets, labels, annotations)
- Data in any charts, tables, or matrices (include numbers)
- Visual structure: layout, diagrams, flows, relationships shown
- Icons, colors, or visual elements that convey meaning
- Speaker notes or annotations if visible

Be thorough and specific — this description will be used for semantic search.
Respond in the same language as the slide content (English or Portuguese)."""

    response = client.messages.create(
        model=model,
        max_tokens=1000,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": b64_image}},
                {"type": "text", "text": prompt},
            ],
        }],
    )
    return response.content[0].text


def _describe_text_only(client, raw_text: str, slide_num: int, model: str) -> str:
    """Use Claude (text-only) to enrich a slide description when no image is available."""
    prompt = f"""Slide {slide_num} content:
{raw_text[:3000]}

Rewrite this as a rich semantic description suitable for search indexing.
Include all key information, numbers, concepts, frameworks mentioned.
Be thorough — this will be the only searchable representation of this slide.
Respond in the same language as the input."""

    response = client.messages.create(
        model=model,
        max_tokens=800,
        messages=[{"role": "user", "content": prompt}],
    )
    return response.content[0].text


def visual_ingest_pptx(file_path: Path, source_file: str, source_path: str,
                       project: str = "", tags: list[str] | None = None,
                       ) -> list[DocumentChunk]:
    """
    Visual ingest pipeline for PPTX:
    1. Parse slides for raw text
    2. Render slides as images (if LibreOffice available)
    3. Classify each slide: text → Haiku, complex → Sonnet
    4. Generate rich descriptions via Claude
    5. Return DocumentChunks with enriched content
    """
    from pptx import Presentation

    prs = Presentation(str(file_path))
    client = _get_anthropic_client()

    # Try to render all slides as images
    slide_images = _render_slides_to_images(file_path)
    has_images = len(slide_images) == len(prs.slides)

    # File metadata
    date_created = ""
    try:
        cp = prs.core_properties
        if cp.modified:
            date_created = cp.modified.isoformat().split("T")[0]
        elif cp.created:
            date_created = cp.created.isoformat().split("T")[0]
    except Exception:
        pass

    chunks = []
    haiku_count = 0
    sonnet_count = 0

    for slide_idx, slide in enumerate(prs.slides):
        slide_num = slide_idx + 1

        # Extract raw text
        texts = []
        title = ""
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                texts.append(shape.text.strip())
        try:
            if slide.shapes.title:
                title = slide.shapes.title.text.strip()
        except Exception:
            pass
        if not title and texts:
            title = texts[0][:80] if texts[0] else ""

        raw_text = "\n".join(texts)
        if len(raw_text.strip()) < 20:
            continue  # Skip nearly empty slides

        # Classify complexity
        complexity = _classify_slide_complexity(slide, raw_text)
        model = SONNET if complexity == "complex" else HAIKU

        # Generate description
        if has_images and slide_idx < len(slide_images) and slide_images[slide_idx]:
            b64 = _pil_to_base64(slide_images[slide_idx])
            description = _describe_with_vision(client, b64, slide_num, raw_text, model)
        else:
            # No image — text-only enrichment
            description = _describe_text_only(client, raw_text, slide_num, model)

        # Track model usage
        if model == SONNET:
            sonnet_count += 1
        else:
            haiku_count += 1

        # Build chunk: enriched description + raw text as backup
        chunk_text = f"[Slide {slide_num}: {title}]\n{description}\n\n[Raw Text]\n{raw_text}"

        # Speaker notes
        notes = ""
        if slide.has_notes_slide:
            try:
                notes = slide.notes_slide.notes_text_frame.text.strip()
            except Exception:
                pass
        if notes:
            chunk_text += f"\n\n[Speaker Notes]\n{notes}"

        meta = ChunkMetadata(
            document_type=DocumentType.PPTX,
            source_file=source_file,
            source_path=source_path,
            date_created=date_created,
            project=project,
            tags=tags or [],
            slide_number=slide_num,
            slide_title=title,
            chunk_index=slide_num - 1,
            total_chunks=len(prs.slides),
        )
        chunks.append(DocumentChunk(text=chunk_text, metadata=meta))

    # Print cost summary
    print(f"  Visual ingest: {haiku_count} slides → Haiku, {sonnet_count} slides → Sonnet")
    est_cost = haiku_count * 0.002 + sonnet_count * 0.015
    print(f"  Estimated cost: ~${est_cost:.3f}")

    return chunks
