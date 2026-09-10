"""
PPTX Chunker — One chunk per slide, with slide title as context prefix
"""
from ...models import DocumentChunk, ChunkMetadata, DocumentType


def chunk_pptx(parsed: dict, source_file: str, source_path: str,
               project: str = "", tags: list[str] | None = None) -> list[DocumentChunk]:
    """
    Create one DocumentChunk per slide.
    Chunk text format: [Slide N: Title]\n{body}\n[Speaker Notes: {notes}]
    """
    slides = parsed["slides"]
    file_meta = parsed["metadata"]
    date_created = file_meta.get("modified") or file_meta.get("created") or ""
    # Truncate datetime to date
    if date_created and "T" in date_created:
        date_created = date_created.split("T")[0]

    chunks = []
    for slide in slides:
        parts = []
        header = f"[Slide {slide['number']}"
        if slide["title"]:
            header += f": {slide['title']}"
        header += "]"
        parts.append(header)

        if slide["body"]:
            parts.append(slide["body"])
        if slide["notes"]:
            parts.append(f"[Speaker Notes: {slide['notes']}]")

        text = "\n".join(parts)
        # Skip nearly empty slides (just index/divider pages)
        if len(text.strip()) < 30:
            continue

        meta = ChunkMetadata(
            document_type=DocumentType.PPTX,
            source_file=source_file,
            source_path=source_path,
            date_created=date_created,
            project=project,
            tags=tags or [],
            slide_number=slide["number"],
            slide_title=slide["title"],
            chunk_index=slide["number"] - 1,
            total_chunks=len(slides),
        )
        chunks.append(DocumentChunk(text=text, metadata=meta))

    return chunks
