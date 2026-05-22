"""
Embedding Engine — BGE-M3 Sentence Transformer
================================================
Lazy-loaded BGE-M3 model for generating 1024-dim dense embeddings.
Supports batch encoding with progress bars and normalization.
"""

from sentence_transformers import SentenceTransformer
from config import EMBED_MODEL_NAME, EMBED_MAX_SEQ_LENGTH

# Lazy singleton
_embedder: SentenceTransformer | None = None


def get_embedder() -> SentenceTransformer:
    """Return the BGE-M3 embedding model (lazy-loaded singleton)."""
    global _embedder
    if _embedder is None:
        print("   -> Loading BGE-M3 embedding model...")
        _embedder = SentenceTransformer(EMBED_MODEL_NAME, device="cpu")
        _embedder.max_seq_length = EMBED_MAX_SEQ_LENGTH
        print("   -> BGE-M3 ready")
    return _embedder


def embed_texts(texts: list[str]) -> list[list[float]]:
    """Encode a list of texts into normalized BGE-M3 embeddings."""
    embedder = get_embedder()
    embs = embedder.encode(
        texts,
        show_progress_bar=True,
        normalize_embeddings=True,
    )
    return [e.tolist() for e in embs]


def embed_query(query: str) -> list[float]:
    """Encode a single query into a normalized BGE-M3 embedding vector."""
    embedder = get_embedder()
    return embedder.encode(query, normalize_embeddings=True).tolist()
