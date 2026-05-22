"""
Reranker — BGE-reranker-v2-m3 Cross-Encoder
=============================================
Reranks initial vector search results using a cross-encoder model
that scores (query, document) pairs for fine-grained relevance.
"""

from sentence_transformers import CrossEncoder
from config import RERANKER_MODEL_NAME, RERANKER_MAX_LENGTH, TOP_K_RERANK

# Lazy singleton
_reranker: CrossEncoder | None = None


def get_reranker() -> CrossEncoder:
    """Return the BGE reranker cross-encoder model (lazy-loaded singleton)."""
    global _reranker
    if _reranker is None:
        print("   -> Loading BGE-reranker-v2-m3...")
        _reranker = CrossEncoder(
            RERANKER_MODEL_NAME,
            device="cpu",
            max_length=RERANKER_MAX_LENGTH,
        )
        print("   -> BGE-reranker ready")
    return _reranker


def rerank_results(
    query: str,
    hits: list,
    top_k: int = TOP_K_RERANK,
) -> list[tuple]:
    """
    Rerank search results using cross-encoder scoring.

    Args:
        query: The user's search query
        hits: List of Qdrant search hits (with .payload["text"])
        top_k: Number of results to return after reranking

    Returns:
        List of (hit, reranker_score) tuples, sorted by score descending
    """
    reranker = get_reranker()
    pairs = [(query, hit.payload["text"]) for hit in hits]
    scores = reranker.predict(pairs, show_progress_bar=False)

    combined = list(zip(hits, scores))
    combined.sort(key=lambda x: x[1], reverse=True)
    return combined[:top_k]
