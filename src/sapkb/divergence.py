"""
Divergence Detection — Find contradictory information across sources
"""
from .models import DocumentChunk, Divergence
from .store import KnowledgeStore


class DivergenceDetector:
    """Detect contradictions between document chunks."""

    def __init__(self, store: KnowledgeStore, similarity_threshold: float = 0.85):
        self.store = store
        self.similarity_threshold = similarity_threshold

    def check_on_ingest(self, new_chunks: list[DocumentChunk]) -> list[Divergence]:
        """
        After ingesting new chunks, check for potential divergences with
        existing chunks that are semantically similar but from different sources.
        """
        divergences = []
        for chunk in new_chunks:
            # Search for similar existing chunks
            results = self.store.vector_search(chunk.text, n_results=5)
            for existing_chunk, similarity in results:
                # Skip chunks from the same source file
                if existing_chunk.metadata.source_file == chunk.metadata.source_file:
                    continue
                # High similarity but different sources = potential divergence
                if similarity >= self.similarity_threshold:
                    divergences.append(Divergence(
                        chunk_a_id=chunk.id,
                        chunk_b_id=existing_chunk.id,
                        chunk_a_text=chunk.text[:500],
                        chunk_b_text=existing_chunk.text[:500],
                        source_a=chunk.metadata.source_file,
                        source_b=existing_chunk.metadata.source_file,
                        similarity=similarity,
                    ))
        return divergences

    def check_query_results(self, results: list) -> list[Divergence]:
        """
        Check if query results contain contradictory information.
        Groups results by topic similarity and flags conflicts.
        """
        if len(results) < 2:
            return []

        divergences = []
        seen = set()

        for i, r1 in enumerate(results):
            for j, r2 in enumerate(results[i + 1:], i + 1):
                # Skip same source
                if r1.chunk.metadata.source_file == r2.chunk.metadata.source_file:
                    continue

                pair = tuple(sorted([r1.chunk.id, r2.chunk.id]))
                if pair in seen:
                    continue
                seen.add(pair)

                # Use embedding similarity to detect topically related but potentially
                # conflicting chunks. We check if both are highly relevant to the same
                # query (both scored high) but from different sources.
                if r1.score > 0.02 and r2.score > 0.02:
                    # Both are relevant — check if they might contradict
                    # This is a heuristic; proper contradiction detection needs NLI
                    emb1 = self.store.embed([r1.chunk.text])
                    emb2 = self.store.embed([r2.chunk.text])
                    import numpy as np
                    sim = float(np.dot(emb1[0], emb2[0]) / (
                        np.linalg.norm(emb1[0]) * np.linalg.norm(emb2[0])
                    ))
                    if sim >= self.similarity_threshold:
                        divergences.append(Divergence(
                            chunk_a_id=r1.chunk.id,
                            chunk_b_id=r2.chunk.id,
                            chunk_a_text=r1.chunk.text[:500],
                            chunk_b_text=r2.chunk.text[:500],
                            source_a=r1.chunk.metadata.source_file,
                            source_b=r2.chunk.metadata.source_file,
                            similarity=sim,
                        ))
        return divergences
