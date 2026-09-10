"""
Data Models — Pydantic models for documents, chunks, search results, provenance
"""
from datetime import date
from enum import Enum
from hashlib import sha256
from typing import Optional

from pydantic import BaseModel, Field, computed_field


class DocumentType(str, Enum):
    PPTX = "pptx"
    PDF = "pdf"
    DOCX = "docx"
    XLSX = "xlsx"
    TRANSCRIPT = "transcript"
    NOTES = "notes"


class ProvenanceType(str, Enum):
    DADO = "DADO"                # Extracted directly from a source document
    INFERENCIA = "INFERÊNCIA"    # Logical conclusion from multiple real data points
    SUPOSICAO = "SUPOSIÇÃO"      # AI-created/assumed without documentary basis
    DIVERGENCIA = "DIVERGÊNCIA"  # Conflicting information between sources


class ChunkMetadata(BaseModel):
    document_type: DocumentType
    source_file: str
    source_path: str
    date_created: Optional[str] = None
    date_ingested: str = Field(default_factory=lambda: date.today().isoformat())
    project: str = ""
    tags: list[str] = Field(default_factory=list)

    # PPTX-specific
    slide_number: Optional[int] = None
    slide_title: Optional[str] = None

    # Transcript-specific
    speaker: Optional[str] = None
    timestamp_start: Optional[str] = None
    timestamp_end: Optional[str] = None

    # Chunk position
    chunk_index: int = 0
    total_chunks: int = 1

    # Divergence tracking
    has_divergence: bool = False
    divergent_chunk_ids: list[str] = Field(default_factory=list)

    def to_chroma_metadata(self) -> dict:
        """Convert to ChromaDB-compatible flat dict (str/int/float/bool only)."""
        d = {
            "document_type": self.document_type.value,
            "source_file": self.source_file,
            "source_path": self.source_path,
            "date_ingested": self.date_ingested,
            "project": self.project,
            "tags": ",".join(self.tags) if self.tags else "",
            "chunk_index": self.chunk_index,
            "total_chunks": self.total_chunks,
            "has_divergence": self.has_divergence,
        }
        if self.date_created:
            d["date_created"] = self.date_created
        if self.slide_number is not None:
            d["slide_number"] = self.slide_number
        if self.slide_title:
            d["slide_title"] = self.slide_title
        if self.speaker:
            d["speaker"] = self.speaker
        if self.timestamp_start:
            d["timestamp_start"] = self.timestamp_start
        if self.timestamp_end:
            d["timestamp_end"] = self.timestamp_end
        if self.divergent_chunk_ids:
            d["divergent_chunk_ids"] = ",".join(self.divergent_chunk_ids)
        return d

    @staticmethod
    def from_chroma_metadata(d: dict) -> "ChunkMetadata":
        """Reconstruct from ChromaDB metadata dict."""
        tags = d.get("tags", "").split(",") if d.get("tags") else []
        div_ids = d.get("divergent_chunk_ids", "").split(",") if d.get("divergent_chunk_ids") else []
        return ChunkMetadata(
            document_type=DocumentType(d["document_type"]),
            source_file=d["source_file"],
            source_path=d["source_path"],
            date_created=d.get("date_created"),
            date_ingested=d.get("date_ingested", ""),
            project=d.get("project", ""),
            tags=tags,
            slide_number=d.get("slide_number"),
            slide_title=d.get("slide_title"),
            speaker=d.get("speaker"),
            timestamp_start=d.get("timestamp_start"),
            timestamp_end=d.get("timestamp_end"),
            chunk_index=d.get("chunk_index", 0),
            total_chunks=d.get("total_chunks", 1),
            has_divergence=d.get("has_divergence", False),
            divergent_chunk_ids=div_ids,
        )


class DocumentChunk(BaseModel):
    text: str
    metadata: ChunkMetadata

    @computed_field
    @property
    def id(self) -> str:
        raw = f"{self.metadata.source_path}:{self.metadata.chunk_index}"
        return sha256(raw.encode()).hexdigest()[:16]


class SearchResult(BaseModel):
    chunk: DocumentChunk
    score: float
    provenance: ProvenanceType = ProvenanceType.DADO
    search_method: str = "hybrid"  # "vector", "bm25", "hybrid"

    def format_source(self) -> str:
        m = self.chunk.metadata
        loc = ""
        if m.slide_number is not None:
            loc = f" | Slide {m.slide_number}"
            if m.slide_title:
                loc += f": {m.slide_title}"
        elif m.speaker:
            loc = f" | {m.speaker}"
            if m.timestamp_start:
                loc += f" [{m.timestamp_start}]"
        return f"[{self.provenance.value}] {m.source_file}{loc} (score: {self.score:.3f})"


class Divergence(BaseModel):
    chunk_a_id: str
    chunk_b_id: str
    chunk_a_text: str
    chunk_b_text: str
    source_a: str
    source_b: str
    similarity: float
    topic: str = ""


class IngestionResult(BaseModel):
    source_file: str
    chunks_added: int
    document_type: DocumentType
    project: str = ""
    warnings: list[str] = Field(default_factory=list)
