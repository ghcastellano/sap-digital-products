"""
Hybrid Search — Vector (ChromaDB) + BM25 keyword search merged with RRF
"""
import pickle
import re
from pathlib import Path

from rank_bm25 import BM25Okapi

from .models import DocumentChunk, SearchResult, ProvenanceType
from .store import KnowledgeStore


def _tokenize(text: str) -> list[str]:
    """Simple whitespace + lowercase tokenizer for BM25."""
    return re.findall(r'\w+', text.lower())


class HybridSearcher:
    """Combines ChromaDB vector search with BM25 keyword search via RRF."""

    def __init__(self, store: KnowledgeStore, bm25_index_path: str,
                 rrf_k: int = 60):
        self.store = store
        self.bm25_index_path = Path(bm25_index_path)
        self.rrf_k = rrf_k
        self._bm25: BM25Okapi | None = None
        self._bm25_chunks: list[DocumentChunk] = []
        self._load_bm25_index()

    def _load_bm25_index(self):
        corpus_path = self.bm25_index_path / "bm25_corpus.pkl"
        if corpus_path.exists():
            with open(corpus_path, "rb") as f:
                data = pickle.load(f)
            self._bm25 = data["bm25"]
            self._bm25_chunks = data["chunks"]

    def build_bm25_index(self):
        """Build/rebuild BM25 index from all chunks in ChromaDB."""
        self._bm25_chunks = self.store.get_all_chunks()
        if not self._bm25_chunks:
            self._bm25 = None
            return

        tokenized = [_tokenize(c.text) for c in self._bm25_chunks]
        self._bm25 = BM25Okapi(tokenized)

        self.bm25_index_path.mkdir(parents=True, exist_ok=True)
        with open(self.bm25_index_path / "bm25_corpus.pkl", "wb") as f:
            pickle.dump({"bm25": self._bm25, "chunks": self._bm25_chunks}, f)

    def _bm25_search(self, query: str, top_k: int = 20) -> list[tuple[DocumentChunk, float]]:
        if self._bm25 is None or not self._bm25_chunks:
            return []
        tokens = _tokenize(query)
        scores = self._bm25.get_scores(tokens)
        ranked = sorted(enumerate(scores), key=lambda x: x[1], reverse=True)[:top_k]
        return [(self._bm25_chunks[i], s) for i, s in ranked if s > 0]

    def search(self, query: str, top_k: int = 10,
               filters: dict | None = None) -> list[SearchResult]:
        """
        Hybrid search:
        1. ChromaDB vector search (top 20)
        2. BM25 keyword search (top 20)
        3. Merge with Reciprocal Rank Fusion
        4. Return top_k results
        """
        # Vector search
        vector_results = self.store.vector_search(query, n_results=20, where=filters)

        # BM25 search (no metadata filtering — filtered post-merge if needed)
        bm25_results = self._bm25_search(query, top_k=20)

        # If only one method has results, use it directly
        if not vector_results:
            return self._to_search_results(bm25_results[:top_k], "bm25")
        if not bm25_results:
            return self._to_search_results(vector_results[:top_k], "vector")

        # RRF merge
        merged = self._rrf_merge(vector_results, bm25_results)

        # Apply metadata filters to BM25 results post-merge
        if filters:
            merged = self._apply_filters(merged, filters)

        return merged[:top_k]

    def _rrf_merge(self, vector_results: list[tuple[DocumentChunk, float]],
                   bm25_results: list[tuple[DocumentChunk, float]]) -> list[SearchResult]:
        """Reciprocal Rank Fusion: score = sum(1/(k+rank)) across result lists."""
        scores: dict[str, float] = {}
        chunk_map: dict[str, DocumentChunk] = {}

        for rank, (chunk, _) in enumerate(vector_results):
            cid = chunk.id
            scores[cid] = scores.get(cid, 0) + 1.0 / (self.rrf_k + rank + 1)
            chunk_map[cid] = chunk

        for rank, (chunk, _) in enumerate(bm25_results):
            cid = chunk.id
            scores[cid] = scores.get(cid, 0) + 1.0 / (self.rrf_k + rank + 1)
            chunk_map[cid] = chunk

        ranked = sorted(scores.items(), key=lambda x: x[1], reverse=True)
        results = []
        for cid, score in ranked:
            chunk = chunk_map[cid]
            results.append(SearchResult(
                chunk=chunk,
                score=score,
                provenance=ProvenanceType.DADO,
                search_method="hybrid",
            ))
        return results

    def _to_search_results(self, items: list[tuple[DocumentChunk, float]],
                           method: str) -> list[SearchResult]:
        return [
            SearchResult(chunk=chunk, score=score, provenance=ProvenanceType.DADO,
                         search_method=method)
            for chunk, score in items
        ]

    def _apply_filters(self, results: list[SearchResult],
                       filters: dict) -> list[SearchResult]:
        """Post-filter results by metadata (for BM25 results that weren't pre-filtered)."""
        filtered = []
        for r in results:
            meta = r.chunk.metadata
            match = True
            if "project" in filters and meta.project != filters["project"]:
                match = False
            if "document_type" in filters and meta.document_type.value != filters["document_type"]:
                match = False
            if match:
                filtered.append(r)
        return filtered
