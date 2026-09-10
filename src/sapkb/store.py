"""
Knowledge Store — ChromaDB wrapper with embedded local persistence
"""
import chromadb
from sentence_transformers import SentenceTransformer

from .models import DocumentChunk, ChunkMetadata


class KnowledgeStore:
    """ChromaDB-backed document store with sentence-transformer embeddings."""

    def __init__(self, db_path: str, collection_name: str = "consulting_documents",
                 embedding_model: str = "sentence-transformers/all-MiniLM-L6-v2"):
        self._client = chromadb.PersistentClient(path=db_path)
        self._model = SentenceTransformer(embedding_model)
        self._collection = self._client.get_or_create_collection(
            name=collection_name,
            metadata={"hnsw:space": "cosine"},
        )

    @property
    def collection(self):
        return self._collection

    @property
    def model(self) -> SentenceTransformer:
        return self._model

    def embed(self, texts: list[str]) -> list[list[float]]:
        return self._model.encode(texts, show_progress_bar=False).tolist()

    def add_chunks(self, chunks: list[DocumentChunk]) -> int:
        if not chunks:
            return 0

        ids = [c.id for c in chunks]
        texts = [c.text for c in chunks]
        metadatas = [c.metadata.to_chroma_metadata() for c in chunks]
        embeddings = self.embed(texts)

        # Upsert to handle re-ingestion gracefully
        self._collection.upsert(
            ids=ids,
            documents=texts,
            metadatas=metadatas,
            embeddings=embeddings,
        )
        return len(ids)

    def vector_search(self, query: str, n_results: int = 20,
                      where: dict | None = None) -> list[tuple[DocumentChunk, float]]:
        query_embedding = self.embed([query])
        kwargs = {
            "query_embeddings": query_embedding,
            "n_results": min(n_results, self._collection.count() or 1),
        }
        if where:
            kwargs["where"] = where

        results = self._collection.query(**kwargs)

        output = []
        for i in range(len(results["ids"][0])):
            meta = ChunkMetadata.from_chroma_metadata(results["metadatas"][0][i])
            chunk = DocumentChunk(text=results["documents"][0][i], metadata=meta)
            # ChromaDB returns distances; for cosine, distance = 1 - similarity
            distance = results["distances"][0][i]
            score = 1.0 - distance
            output.append((chunk, score))
        return output

    def get_all_chunks(self) -> list[DocumentChunk]:
        """Retrieve all chunks (for BM25 index building)."""
        count = self._collection.count()
        if count == 0:
            return []
        results = self._collection.get(limit=count, include=["documents", "metadatas"])
        chunks = []
        for i in range(len(results["ids"])):
            meta = ChunkMetadata.from_chroma_metadata(results["metadatas"][i])
            chunk = DocumentChunk(text=results["documents"][i], metadata=meta)
            chunks.append(chunk)
        return chunks

    def get_stats(self) -> dict:
        count = self._collection.count()
        if count == 0:
            return {"total_chunks": 0, "by_project": {}, "by_type": {}, "sources": []}

        results = self._collection.get(limit=count, include=["metadatas"])
        by_project: dict[str, int] = {}
        by_type: dict[str, int] = {}
        sources: set[str] = set()

        for meta in results["metadatas"]:
            proj = meta.get("project", "unknown")
            dtype = meta.get("document_type", "unknown")
            by_project[proj] = by_project.get(proj, 0) + 1
            by_type[dtype] = by_type.get(dtype, 0) + 1
            sources.add(meta.get("source_file", ""))

        return {
            "total_chunks": count,
            "total_documents": len(sources),
            "by_project": by_project,
            "by_type": by_type,
            "sources": sorted(sources),
        }

    def delete_by_source(self, source_file: str) -> int:
        results = self._collection.get(
            where={"source_file": source_file},
            include=[],
        )
        ids = results["ids"]
        if ids:
            self._collection.delete(ids=ids)
        return len(ids)
