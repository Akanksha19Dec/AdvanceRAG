"""
LLM Integration — Groq Cloud (Llama 3.3 70B Versatile)
=========================================================
Handles answer generation and Naive vs Advanced RAG comparison
via the Groq API.
"""

import time

import requests

from config import (
    GROQ_API_KEY,
    GROQ_API_URL,
    GROQ_MAX_TOKENS,
    GROQ_MODEL,
    GROQ_TEMPERATURE,
)

# Cache for dataset summary to avoid recomputing
_stats_cache: dict = {"summary": None, "timestamp": 0}


def _build_stats_summary(test_cases: list[dict]) -> str:
    """Build a condensed statistics summary of the test case dataset."""
    module_counts: dict[str, int] = {}
    cat_counts: dict[str, int] = {}
    pri_counts: dict[str, int] = {}

    for tc in test_cases:
        module_counts[tc["module"]] = module_counts.get(tc["module"], 0) + 1
        cat_counts[tc["category"]] = cat_counts.get(tc["category"], 0) + 1
        pri_counts[tc["priority"]] = pri_counts.get(tc["priority"], 0) + 1

    return (
        f"Total test cases: {len(test_cases)}\n"
        f"Modules ({len(module_counts)}): "
        + ", ".join(f"{m} ({c})" for m, c in sorted(module_counts.items()))
        + "\n"
        f"Categories: "
        + ", ".join(f"{c} ({n})" for c, n in sorted(cat_counts.items()))
        + "\n"
        f"Priorities: "
        + ", ".join(f"{p} ({n})" for p, n in sorted(pri_counts.items()))
    )


def _get_stats(test_cases: list[dict], force: bool = False) -> str:
    """Get cached or fresh dataset stats summary (5-min cache TTL)."""
    now = time.time()
    if _stats_cache["summary"] is None or now - _stats_cache["timestamp"] > 300 or force:
        _stats_cache["summary"] = _build_stats_summary(test_cases)
        _stats_cache["timestamp"] = now
    return _stats_cache["summary"]


def generate_answer(
    query: str,
    reranked: list[tuple],
    test_cases: list[dict],
) -> str:
    """
    Generate a comprehensive answer using Groq LLM based on reranked results.

    Args:
        query: The user's question
        reranked: List of (hit, score) tuples from the reranker
        test_cases: Full dataset for statistics context

    Returns:
        LLM-generated answer string
    """
    stats = _get_stats(test_cases)

    # Build context from reranked results
    context = "\n\n---\n\n".join(
        f"[{r[0].payload['tc_id']}] (Relevance: {r[1]:.4f}):\n{r[0].payload['text']}"
        for r in reranked
    )

    system_msg = (
        "You are a Senior QA Lead for VWO (vwo.com), a digital experience platform. "
        "Answer questions about VWO Login Dashboard test cases based on the provided context. "
        "Always reference specific Test Case IDs. Be thorough and list ALL relevant matches. "
        "Include module, category, and priority information.\n\n"
        f"DATASET OVERVIEW:\n{stats}"
    )

    payload = {
        "model": GROQ_MODEL,
        "messages": [
            {"role": "system", "content": system_msg},
            {
                "role": "user",
                "content": (
                    f"CONTEXT (Re-ranked Test Cases — most relevant first):\n{context}\n\n"
                    f"QUESTION: {query}\n\n"
                    "List ALL relevant Test Case IDs from the context with brief explanation."
                ),
            },
        ],
        "temperature": GROQ_TEMPERATURE,
        "max_tokens": GROQ_MAX_TOKENS,
    }

    resp = requests.post(
        GROQ_API_URL,
        headers={
            "Authorization": f"Bearer {GROQ_API_KEY}",
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=60,
    )
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"]


def generate_comparison() -> str:
    """
    Generate a Naive RAG vs Advanced RAG comparison using Groq LLM.
    Falls back to a static comparison if the API call fails.
    """
    system_msg = (
        "You are a technical architect. Compare Naive RAG vs Advanced RAG in 3-4 bullet points. "
        "Be concise and specific about the technical improvements "
        "(BGE-M3, Qdrant, BGE Reranker, 5000 test cases)."
    )

    payload = {
        "model": GROQ_MODEL,
        "messages": [
            {"role": "system", "content": system_msg},
            {
                "role": "user",
                "content": (
                    "Compare NaiveRAG (nomic-embed-text + ChromaDB + 100 test cases + no reranker) "
                    "vs AdvancedRAG (BGE-M3 + Qdrant + BGE-reranker-v2-m3 + 5000 test cases). "
                    "Explain why AdvancedRAG produces better results."
                ),
            },
        ],
        "temperature": 0.3,
        "max_tokens": 500,
    }

    try:
        resp = requests.post(
            GROQ_API_URL,
            headers={
                "Authorization": f"Bearer {GROQ_API_KEY}",
                "Content-Type": "application/json",
            },
            json=payload,
            timeout=30,
        )
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"]
    except Exception:
        # Static fallback
        return (
            "- **Better Embeddings**: BGE-M3 (1024-dim, MTEB 63.0) vs nomic-embed-text (768-dim) "
            "— captures semantic nuances better\n"
            "- **Faster Vector DB**: Qdrant (Rust-based) vs ChromaDB (Python-based) "
            "— faster search with better filtering\n"
            "- **Cross-Encoder Reranking**: BGE-reranker-v2-m3 reorders top-10 results "
            "for maximum relevance vs Naive RAG's raw vector similarity only\n"
            "- **50× More Data**: 5000 test cases across 20 modules vs 100 across 8 "
            "— comprehensive coverage"
        )
