"""
Query Pipeline — Orchestrates the full Advanced RAG flow
=========================================================
Steps:
  1. User query → Smart module detection
  2. BGE-M3 embedding of query
  3. Qdrant vector search (or module-filtered retrieval)
  4. BGE-reranker-v2-m3 cross-encoder reranking
  5. Groq LLM answer generation
  6. Return structured response with before/after comparisons
"""

import numpy as np

from config import MODULE_KEYWORDS, TOP_K_RERANK, TOP_K_RETRIEVE
from embeddings import embed_query
from reranker import rerank_results
from vector_store import search_by_module, search_qdrant
from llm import generate_answer


def detect_module(query: str) -> str | None:
    """
    Detect the most likely module from the user's query
    using keyword matching with longest-match-wins scoring.
    """
    q = query.lower()
    best_match = None
    best_score = 0

    for module, keywords in MODULE_KEYWORDS.items():
        for kw in keywords:
            if kw in q:
                score = len(kw)
                if score > best_score:
                    best_score = score
                    best_match = module

    return best_match


def execute_query(
    query: str,
    qdrant_client,
    test_cases: list[dict],
) -> dict:
    """
    Execute the full Advanced RAG pipeline for a given query.

    Returns:
        dict with: answer, retrieved_before, retrieved_after,
                   module_detected, total_cases
    """
    module_name = detect_module(query)
    query_vector = embed_query(query)

    if module_name:
        # ── Module-specific retrieval ──
        # Get all test cases from the detected module
        module_hits = search_by_module(
            qdrant_client,
            module_name,
            limit=TOP_K_RERANK * 4,
        )

        # Score by cosine similarity within the module
        qv = np.array(query_vector)
        scored = []
        for hit in module_hits:
            # For scroll results, vectors aren't returned — use payload match
            # We compute similarity via embedding dot product
            sim = 0.0
            if hasattr(hit, "vector") and hit.vector is not None:
                sim = float(np.dot(qv, np.array(hit.vector)))
            scored.append((hit, sim))

        scored.sort(key=lambda x: x[1], reverse=True)
        hits = [s[0] for s in scored[:TOP_K_RETRIEVE]]

        # For module queries, rerank the hits
        reranked = rerank_results(query, hits)

    else:
        # ── General semantic search ──
        hits = search_qdrant(qdrant_client, query_vector, TOP_K_RETRIEVE)
        reranked = rerank_results(query, hits)

    # Build before-reranking response
    retrieved_before = []
    for hit in hits:
        p = hit.payload
        sim = getattr(hit, "score", 0.0)
        retrieved_before.append({
            "chunk_id": p["chunk_id"],
            "tc_id": p["tc_id"],
            "module": p["module"],
            "category": p["category"],
            "priority": p["priority"],
            "similarity": round(float(sim), 4),
            "text": p["text"],
            "length": p["length"],
        })

    # Build after-reranking response
    retrieved_after = []
    for hit, score in reranked:
        p = hit.payload
        sim_before = getattr(hit, "score", 0.0)
        retrieved_after.append({
            "chunk_id": p["chunk_id"],
            "tc_id": p["tc_id"],
            "module": p["module"],
            "category": p["category"],
            "priority": p["priority"],
            "similarity_before": round(float(sim_before), 4),
            "similarity_after": round(float(score), 4),
            "score_improvement": round(float(score) - float(sim_before), 4),
            "text": p["text"],
            "length": p["length"],
        })

    # Generate LLM answer
    answer = generate_answer(query, reranked, test_cases)

    return {
        "answer": answer,
        "retrieved_before": retrieved_before,
        "retrieved_after": retrieved_after,
        "module_detected": module_name,
        "total_cases": len(test_cases),
    }
