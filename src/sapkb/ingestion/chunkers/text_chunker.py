"""
Text Chunker — Sliding window for notes; speaker-aware for transcripts
"""
import re

from ...models import DocumentChunk, ChunkMetadata, DocumentType


def chunk_notes(text: str, source_file: str, source_path: str,
                project: str = "", tags: list[str] | None = None,
                window_size: int = 500, overlap_pct: float = 0.20) -> list[DocumentChunk]:
    """Sliding window chunker for unstructured notes."""
    words = text.split()
    if not words:
        return []

    overlap = int(window_size * overlap_pct)
    step = window_size - overlap
    chunks = []
    idx = 0

    while idx < len(words):
        window = words[idx:idx + window_size]
        chunk_text = " ".join(window)
        if len(chunk_text.strip()) < 20:
            idx += step
            continue

        meta = ChunkMetadata(
            document_type=DocumentType.NOTES,
            source_file=source_file,
            source_path=source_path,
            project=project,
            tags=tags or [],
            chunk_index=len(chunks),
            total_chunks=0,  # will be updated
        )
        chunks.append(DocumentChunk(text=chunk_text, metadata=meta))
        idx += step

    # Update total_chunks
    for c in chunks:
        c.metadata.total_chunks = len(chunks)
    return chunks


# Speaker line patterns
_SPEAKER_PATTERNS = [
    re.compile(r'^\[?(\d{1,2}:\d{2}(?::\d{2})?)\]?\s*[-–]?\s*(.+?):\s*(.+)', re.MULTILINE),  # [00:05] John: text
    re.compile(r'^(.+?):\s+(.+)', re.MULTILINE),  # John: text
]


def chunk_transcript(text: str, source_file: str, source_path: str,
                     project: str = "", tags: list[str] | None = None,
                     default_speaker: str = "") -> list[DocumentChunk]:
    """Speaker-aware chunker for interview transcripts."""
    # Try to parse speaker turns
    turns = _parse_turns(text, default_speaker)

    if not turns:
        # Fallback to notes chunker
        return chunk_notes(text, source_file, source_path, project, tags)

    chunks = []
    # Group consecutive turns into chunks (~500 words max)
    current_text_parts = []
    current_speaker = turns[0]["speaker"]
    current_timestamp = turns[0].get("timestamp", "")
    current_words = 0

    for turn in turns:
        turn_words = len(turn["text"].split())
        if current_words + turn_words > 500 and current_text_parts:
            # Flush current chunk
            meta = ChunkMetadata(
                document_type=DocumentType.TRANSCRIPT,
                source_file=source_file,
                source_path=source_path,
                project=project,
                tags=tags or [],
                speaker=current_speaker,
                timestamp_start=current_timestamp,
                chunk_index=len(chunks),
                total_chunks=0,
            )
            chunks.append(DocumentChunk(
                text="\n".join(current_text_parts),
                metadata=meta,
            ))
            current_text_parts = []
            current_words = 0
            current_speaker = turn["speaker"]
            current_timestamp = turn.get("timestamp", "")

        prefix = f"[{turn['speaker']}]"
        if turn.get("timestamp"):
            prefix = f"[{turn['timestamp']}] {prefix}"
        current_text_parts.append(f"{prefix}: {turn['text']}")
        current_words += turn_words

    # Flush last chunk
    if current_text_parts:
        meta = ChunkMetadata(
            document_type=DocumentType.TRANSCRIPT,
            source_file=source_file,
            source_path=source_path,
            project=project,
            tags=tags or [],
            speaker=current_speaker,
            timestamp_start=current_timestamp,
            chunk_index=len(chunks),
            total_chunks=0,
        )
        chunks.append(DocumentChunk(text="\n".join(current_text_parts), metadata=meta))

    for c in chunks:
        c.metadata.total_chunks = len(chunks)
    return chunks


def _parse_turns(text: str, default_speaker: str) -> list[dict]:
    """Try to parse speaker turns from text."""
    turns = []
    lines = text.strip().split("\n")

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Try timestamped pattern first
        m = _SPEAKER_PATTERNS[0].match(line)
        if m:
            turns.append({
                "timestamp": m.group(1),
                "speaker": m.group(2).strip(),
                "text": m.group(3).strip(),
            })
            continue

        # Try simple speaker pattern
        m = _SPEAKER_PATTERNS[1].match(line)
        if m:
            speaker = m.group(1).strip()
            # Heuristic: speaker names are short
            if len(speaker) < 40 and not any(c in speaker for c in ".;,!?"):
                turns.append({
                    "speaker": speaker,
                    "text": m.group(2).strip(),
                })
                continue

        # Continuation of previous turn or unstructured
        if turns:
            turns[-1]["text"] += " " + line
        elif default_speaker:
            turns.append({"speaker": default_speaker, "text": line})

    return turns
