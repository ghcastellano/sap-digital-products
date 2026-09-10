"""
Settings — Pydantic Settings with .env support
"""
from pathlib import Path
from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    # Paths
    project_root: Path = Path(__file__).parent.parent
    chroma_db_path: Path = Path(__file__).parent.parent / "data" / "chroma_db"
    bm25_index_path: Path = Path(__file__).parent.parent / "data" / "bm25_index"
    generated_output_path: Path = Path(__file__).parent.parent / "data" / "generated"

    # Source document paths
    base_presentations_path: Path = Path.home() / "Downloads" / "Base Apresentações Projeto Digital Product"
    proposal_path: Path = Path(__file__).parent.parent / "Proposta Digital Products" / "SAP CoE New Ways of Working v2.0 (1).pptx"

    # Embeddings
    embedding_model: str = "sentence-transformers/all-MiniLM-L6-v2"
    embedding_dimension: int = 384

    # ChromaDB
    chroma_collection_name: str = "consulting_documents"
    chroma_notes_collection: str = "field_notes"

    # Search
    search_top_k: int = 20
    rrf_k: int = 60
    final_results: int = 10

    # LLM
    openai_api_key: str = ""
    openai_model: str = "gpt-4o"
    openai_temperature: float = 0.3
    max_context_tokens: int = 8000

    # Anthropic (for visual ingest)
    anthropic_api_key: str = ""

    # Brand (extracted from Proposal)
    brand_primary: str = "#002060"
    brand_secondary: str = "#0070C0"
    brand_accent: str = "#19A3FC"
    brand_font: str = "Arial"

    model_config = {"env_file": ".env", "extra": "ignore"}


def get_settings() -> Settings:
    return Settings()
