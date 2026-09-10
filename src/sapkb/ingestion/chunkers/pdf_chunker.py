"""
PDF Chunker — One chunk per page with page number context
"""
from ...models import DocumentChunk, ChunkMetadata, DocumentType


def chunk_pdf(parsed: dict, source_file: str, source_path: str,
              project: str = "", tags: list[str] | None = None) -> list[DocumentChunk]:
    """Create one DocumentChunk per page."""
    pages = parsed["pages"]
    chunks = []
    for page in pages:
        if len(page["text"].strip()) < 20:
            continue
        text = f"[Page {page['number']}]\n{page['text']}"
        meta = ChunkMetadata(
            document_type=DocumentType.PDF,
            source_file=source_file,
            source_path=source_path,
            project=project,
            tags=tags or [],
            chunk_index=page["number"] - 1,
            total_chunks=len(pages),
        )
        chunks.append(DocumentChunk(text=text, metadata=meta))
    return chunks
