"""
Ingestion Orchestrator — Pipeline coordinator for all document types
"""
import os
from pathlib import Path

from ..models import DocumentChunk, DocumentType, IngestionResult
from ..store import KnowledgeStore
from .parsers.pptx_parser import parse_pptx
from .parsers.pdf_parser import parse_pdf
from .parsers.docx_parser import parse_docx
from .parsers.xlsx_parser import parse_xlsx
from .chunkers.pptx_chunker import chunk_pptx
from .chunkers.pdf_chunker import chunk_pdf
from .chunkers.text_chunker import chunk_notes, chunk_transcript


# Area classification heuristics
_PROJECT_KEYWORDS = {
    # P&S / P&T area
    "nwow": "P&S / P&T",
    "p&s": "P&S / P&T",
    "brp": "P&S / P&T",
    "product discovery": "P&S / P&T",
    "project management playbook": "P&S / P&T",
    "pm playbook": "P&S / P&T",
    # Digital CoE area
    "digital coe": "Digital CoE",
    "agility playbook": "Digital CoE",
    "nextgen": "Digital CoE",
    "nexgen": "Digital CoE",
    "nttdata": "Digital CoE",
    "jira agile management": "Digital CoE",
    "cx": "Digital CoE",
    # Infrastructure area
    "avangrid": "Infrastructure",
    "tpo": "Infrastructure",
    "project_book": "Infrastructure",
    "project book": "Infrastructure",
    # Renewables DTO area
    "renewables": "Renewables DTO",
    "dto": "Renewables DTO",
    "nextgen_dto": "Renewables DTO",
    "deliverable1": "Renewables DTO",
    # SAP Digital Products (the proposal itself)
    "sap coe": "SAP Digital Products",
    "digital products": "SAP Digital Products",
    "sap": "SAP Digital Products",
}


def _detect_project(filename: str) -> str:
    """Auto-detect project from filename."""
    lower = filename.lower()
    for keyword, project in _PROJECT_KEYWORDS.items():
        if keyword in lower:
            return project
    return ""


def _detect_type(file_path: Path) -> DocumentType:
    ext = file_path.suffix.lower()
    return {
        ".pptx": DocumentType.PPTX,
        ".pdf": DocumentType.PDF,
        ".docx": DocumentType.DOCX,
        ".xlsx": DocumentType.XLSX,
        ".txt": DocumentType.NOTES,
        ".md": DocumentType.NOTES,
    }.get(ext, DocumentType.NOTES)


class IngestionOrchestrator:
    """Coordinates parsing, chunking, and storage of documents."""

    def __init__(self, store: KnowledgeStore):
        self.store = store

    def ingest_file(self, file_path: Path, project: str = "",
                    tags: list[str] | None = None,
                    visual: bool = False) -> IngestionResult:
        """Ingest a single file into the knowledge base."""
        file_path = Path(file_path)
        source_file = file_path.name
        source_path = str(file_path.absolute())
        doc_type = _detect_type(file_path)
        project = project or _detect_project(source_file)
        warnings = []

        try:
            chunks = self._parse_and_chunk(
                file_path, doc_type, source_file, source_path, project, tags,
                visual=visual,
            )
        except Exception as e:
            return IngestionResult(
                source_file=source_file,
                chunks_added=0,
                document_type=doc_type,
                project=project,
                warnings=[f"Parse error: {e}"],
            )

        if not chunks:
            warnings.append("No chunks produced (file may be empty or unsupported)")
            return IngestionResult(
                source_file=source_file,
                chunks_added=0,
                document_type=doc_type,
                project=project,
                warnings=warnings,
            )

        count = self.store.add_chunks(chunks)
        return IngestionResult(
            source_file=source_file,
            chunks_added=count,
            document_type=doc_type,
            project=project,
            warnings=warnings,
        )

    def ingest_directory(self, dir_path: Path, glob_pattern: str = "*.*",
                         project: str = "",
                         tags: list[str] | None = None) -> list[IngestionResult]:
        """Batch ingest all matching files in a directory."""
        dir_path = Path(dir_path)
        results = []
        for fp in sorted(dir_path.glob(glob_pattern)):
            if fp.is_file() and not fp.name.startswith("."):
                result = self.ingest_file(fp, project=project, tags=tags)
                results.append(result)
        return results

    def ingest_raw_text(self, text: str, doc_type: DocumentType = DocumentType.NOTES,
                        source_name: str = "raw_input",
                        project: str = "", speaker: str = "",
                        tags: list[str] | None = None) -> IngestionResult:
        """Quick ingestion of pasted text (transcripts or notes)."""
        if doc_type == DocumentType.TRANSCRIPT:
            chunks = chunk_transcript(
                text, source_name, source_name, project, tags, default_speaker=speaker
            )
        else:
            chunks = chunk_notes(text, source_name, source_name, project, tags)

        count = self.store.add_chunks(chunks)
        return IngestionResult(
            source_file=source_name,
            chunks_added=count,
            document_type=doc_type,
            project=project,
        )

    def _parse_and_chunk(self, file_path: Path, doc_type: DocumentType,
                         source_file: str, source_path: str,
                         project: str, tags: list[str] | None,
                         visual: bool = False) -> list[DocumentChunk]:
        if doc_type == DocumentType.PPTX:
            if visual:
                from .visual_ingest import visual_ingest_pptx
                return visual_ingest_pptx(
                    file_path, source_file, source_path, project, tags
                )
            parsed = parse_pptx(file_path)
            return chunk_pptx(parsed, source_file, source_path, project, tags)
        elif doc_type == DocumentType.PDF:
            parsed = parse_pdf(file_path)
            return chunk_pdf(parsed, source_file, source_path, project, tags)
        elif doc_type == DocumentType.DOCX:
            parsed = parse_docx(file_path)
            # Use notes chunker for DOCX sections
            full_text = "\n\n".join(
                f"## {s['heading']}\n{s['text']}" if s['heading'] else s['text']
                for s in parsed["sections"]
            )
            return chunk_notes(full_text, source_file, source_path, project, tags)
        elif doc_type == DocumentType.XLSX:
            parsed = parse_xlsx(file_path)
            # Render each sheet as a markdown-heading section + table
            full_text = "\n\n".join(
                f"## Sheet: {s['heading']}\n{s['text']}"
                for s in parsed["sections"]
            )
            return chunk_notes(full_text, source_file, source_path, project, tags)
        elif doc_type in (DocumentType.NOTES, DocumentType.TRANSCRIPT):
            text = file_path.read_text(encoding="utf-8", errors="replace")
            if doc_type == DocumentType.TRANSCRIPT:
                return chunk_transcript(text, source_file, source_path, project, tags)
            return chunk_notes(text, source_file, source_path, project, tags)
        else:
            return []
