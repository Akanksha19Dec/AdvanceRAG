"""
Advanced RAG — Flask Application (Interactive Pipeline)
========================================================
Main entry point. Starts instantly — NO auto-ingestion.
All pipeline stages are driven by the user through the UI.

Pipeline Steps (user-driven):
  1. Upload Excel file
  2. View chunking results
  3. Embed & store in Qdrant
  4. Query → Top-10 vector retrieval
  5. Re-rank → Top-5
  6. LLM answer generation
"""

import os
import sys
import json
import time
import tempfile

from dotenv import load_dotenv
from flask import Flask, jsonify, request, send_from_directory
from werkzeug.utils import secure_filename

load_dotenv()

from config import (
    COLLECTION_NAME,
    EMBED_DIM,
    EMBED_MODEL_NAME,
    FLASK_DEBUG,
    FLASK_PORT,
    GROQ_API_KEY,
    GROQ_MODEL,
    RERANKER_MODEL_NAME,
    QDRANT_PATH,
    TOP_K_RETRIEVE,
    TOP_K_RERANK,
)
from data_ingestion import load_test_cases_from_file, compute_statistics
from embeddings import get_embedder, embed_texts, embed_query
from llm import generate_answer
from query_pipeline import execute_query
from reranker import get_reranker, rerank_results
from vector_store import init_qdrant, populate_qdrant, search_qdrant

# ── Flask App ──
app = Flask(__name__, static_folder=".", static_url_path="")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB max upload

# ── In-memory pipeline state ──
pipeline_state = {
    "test_cases": [],
    "stats": {},
    "qdrant_client": None,
    "db_info": None,
    "upload_filename": None,
    "embeddings_ready": False,
    "last_query_results": None,
}


# ══════════════════════════════════════════════════════════════
#  STARTUP — Lightweight, no auto-ingestion
# ══════════════════════════════════════════════════════════════
print("=" * 60)
print("  Advanced RAG Pipeline — Interactive Mode")
print("  BGE-M3 + Qdrant + BGE Reranker + Groq LLM")
print("=" * 60)
print("\n  ✓ Server starting — no auto-ingestion.")
print("  → Upload your Excel file through the UI to begin.\n")


# ══════════════════════════════════════════════════════════════
#  ROUTES — Frontend
# ══════════════════════════════════════════════════════════════

@app.route("/")
def index():
    """Serve the frontend."""
    return send_from_directory(".", "index.html")


# ══════════════════════════════════════════════════════════════
#  STEP 1: Upload Excel File
# ══════════════════════════════════════════════════════════════

@app.route("/api/upload", methods=["POST"])
def api_upload():
    """
    Upload an Excel file and parse test cases.
    Returns chunking preview data.
    """
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    if not file.filename.endswith((".xlsx", ".xls")):
        return jsonify({"error": "Please upload an Excel file (.xlsx or .xls)"}), 400

    # Save to temp location
    filename = secure_filename(file.filename)
    upload_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
    os.makedirs(upload_dir, exist_ok=True)
    filepath = os.path.join(upload_dir, filename)
    file.save(filepath)

    try:
        test_cases = load_test_cases_from_file(filepath)
        stats = compute_statistics(test_cases)

        pipeline_state["test_cases"] = test_cases
        pipeline_state["stats"] = stats
        pipeline_state["upload_filename"] = filename
        pipeline_state["embeddings_ready"] = False
        pipeline_state["qdrant_client"] = None
        pipeline_state["db_info"] = None

        return jsonify({
            "success": True,
            "filename": filename,
            "total_cases": len(test_cases),
            "modules": stats["modules"],
            "categories": stats["categories"],
            "priorities": stats["priorities"],
        })
    except Exception as e:
        return jsonify({"error": f"Failed to parse Excel: {str(e)}"}), 500


# ══════════════════════════════════════════════════════════════
#  STEP 2: View Chunking
# ══════════════════════════════════════════════════════════════

@app.route("/api/chunks")
def api_chunks():
    """Return chunked test cases with metadata."""
    if not pipeline_state["test_cases"]:
        return jsonify({"error": "No data uploaded yet. Please upload an Excel file first."}), 400

    cases = pipeline_state["test_cases"]
    # Return a sample of chunks with detailed info
    chunks_preview = []
    for tc in cases[:50]:  # First 50 for preview
        chunks_preview.append({
            "id": tc["id"],
            "tc_id": tc["tc_id"],
            "module": tc["module"],
            "category": tc["category"],
            "description": tc["description"],
            "preconditions": tc["preconditions"],
            "steps": tc["steps"],
            "expected": tc["expected"],
            "priority": tc["priority"],
            "text": tc["text"],
            "length": tc["length"],
        })

    return jsonify({
        "total": len(cases),
        "chunks": chunks_preview,
        "all_chunks_summary": {
            "avg_length": round(sum(c["length"] for c in cases) / len(cases), 1),
            "min_length": min(c["length"] for c in cases),
            "max_length": max(c["length"] for c in cases),
            "total_tokens_approx": sum(c["length"] for c in cases) // 4,
        },
        "modules": pipeline_state["stats"]["modules"],
        "categories": pipeline_state["stats"]["categories"],
        "priorities": pipeline_state["stats"]["priorities"],
    })


# ══════════════════════════════════════════════════════════════
#  STEP 3: Embed & Store in Qdrant
# ══════════════════════════════════════════════════════════════

@app.route("/api/embed-and-store", methods=["POST"])
def api_embed_and_store():
    """Embed all test cases with BGE-M3 and store in Qdrant."""
    if not pipeline_state["test_cases"]:
        return jsonify({"error": "No data uploaded. Upload first."}), 400

    try:
        cases = pipeline_state["test_cases"]

        # Initialize Qdrant
        print("\n[Qdrant] Initializing vector database...")
        qdrant_client = init_qdrant()

        # Populate (embed + store)
        print(f"[Embed] Embedding {len(cases)} test cases with BGE-M3...")
        db_info = populate_qdrant(qdrant_client, cases)

        pipeline_state["qdrant_client"] = qdrant_client
        pipeline_state["db_info"] = db_info
        pipeline_state["embeddings_ready"] = True

        # Grab a sample embedding for visualization
        sample_embedding = None
        try:
            from embeddings import embed_texts as _et
            sample_text = cases[0]["text"]
            sample_emb = _et([sample_text])[0]
            sample_embedding = {
                "first_10_dims": [round(v, 6) for v in sample_emb[:10]],
                "dimension": len(sample_emb),
                "norm": round(sum(v ** 2 for v in sample_emb) ** 0.5, 6),
            }
        except Exception:
            pass

        return jsonify({
            "success": True,
            "points_count": db_info["points_count"],
            "status": db_info["status"],
            "embed_model": EMBED_MODEL_NAME,
            "embed_dim": EMBED_DIM,
            "vector_db": "Qdrant",
            "collection_name": COLLECTION_NAME,
            "sample_embedding": sample_embedding,
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"Embedding failed: {str(e)}"}), 500


# ══════════════════════════════════════════════════════════════
#  STEP 4 & 5: Search (Top-10) + Rerank (Top-5)
# ══════════════════════════════════════════════════════════════

@app.route("/api/search", methods=["POST"])
def api_search():
    """
    Perform vector search (top-10) and return raw results.
    This is the retrieval-only step before reranking.
    """
    if not pipeline_state["embeddings_ready"]:
        return jsonify({"error": "Embeddings not ready. Complete Step 3 first."}), 400

    data = request.get_json()
    query = data.get("question", "").strip()
    if not query:
        return jsonify({"error": "Please provide a question."}), 400

    try:
        # Embed the query
        query_vector = embed_query(query)

        # Search Qdrant (top-10)
        hits = search_qdrant(
            pipeline_state["qdrant_client"],
            query_vector,
            TOP_K_RETRIEVE,
        )

        # Build results
        retrieved = []
        for rank, hit in enumerate(hits, 1):
            p = hit.payload
            retrieved.append({
                "rank": rank,
                "chunk_id": p["chunk_id"],
                "tc_id": p["tc_id"],
                "module": p["module"],
                "category": p["category"],
                "priority": p["priority"],
                "similarity": round(float(hit.score), 4),
                "text": p["text"],
                "length": p["length"],
                "description": p.get("description", ""),
            })

        # Store for reranking step
        pipeline_state["last_query_results"] = {
            "query": query,
            "hits": hits,
            "retrieved": retrieved,
        }

        return jsonify({
            "success": True,
            "query": query,
            "top_k": TOP_K_RETRIEVE,
            "results": retrieved,
        })
    except Exception as e:
        return jsonify({"error": f"Search failed: {str(e)}"}), 500


@app.route("/api/rerank", methods=["POST"])
def api_rerank():
    """
    Rerank the last search results using BGE-reranker-v2-m3.
    Returns top-5 after cross-encoder scoring.
    """
    if not pipeline_state["last_query_results"]:
        return jsonify({"error": "No search results to rerank. Run search first."}), 400

    try:
        query = pipeline_state["last_query_results"]["query"]
        hits = pipeline_state["last_query_results"]["hits"]

        # Rerank with cross-encoder
        reranked = rerank_results(query, hits, TOP_K_RERANK)

        # Build before/after comparison
        before_results = pipeline_state["last_query_results"]["retrieved"]
        after_results = []
        for rank, (hit, score) in enumerate(reranked, 1):
            p = hit.payload
            sim_before = getattr(hit, "score", 0.0)
            after_results.append({
                "rank": rank,
                "chunk_id": p["chunk_id"],
                "tc_id": p["tc_id"],
                "module": p["module"],
                "category": p["category"],
                "priority": p["priority"],
                "similarity_before": round(float(sim_before), 4),
                "reranker_score": round(float(score), 4),
                "text": p["text"],
                "length": p["length"],
                "description": p.get("description", ""),
            })

        # Store reranked for LLM step
        pipeline_state["last_query_results"]["reranked"] = reranked
        pipeline_state["last_query_results"]["after_results"] = after_results

        return jsonify({
            "success": True,
            "query": query,
            "reranker_model": RERANKER_MODEL_NAME,
            "top_k_before": TOP_K_RETRIEVE,
            "top_k_after": TOP_K_RERANK,
            "before": before_results,
            "after": after_results,
        })
    except Exception as e:
        return jsonify({"error": f"Reranking failed: {str(e)}"}), 500


# ══════════════════════════════════════════════════════════════
#  STEP 6: LLM Answer
# ══════════════════════════════════════════════════════════════

@app.route("/api/answer", methods=["POST"])
def api_answer():
    """Generate LLM answer from reranked results."""
    if not pipeline_state["last_query_results"] or "reranked" not in pipeline_state["last_query_results"]:
        return jsonify({"error": "No reranked results. Complete reranking first."}), 400

    if not GROQ_API_KEY:
        return jsonify({"error": "GROQ_API_KEY not set in .env"}), 500

    try:
        query = pipeline_state["last_query_results"]["query"]
        reranked = pipeline_state["last_query_results"]["reranked"]
        test_cases = pipeline_state["test_cases"]

        answer = generate_answer(query, reranked, test_cases)

        return jsonify({
            "success": True,
            "query": query,
            "answer": answer,
            "llm_model": GROQ_MODEL,
            "llm_provider": "Groq",
        })
    except Exception as e:
        return jsonify({"error": f"LLM generation failed: {str(e)}"}), 500


# ══════════════════════════════════════════════════════════════
#  UTILITY: Full query (for the final Q&A box)
# ══════════════════════════════════════════════════════════════

@app.route("/api/ask", methods=["POST"])
def api_ask():
    """
    Full Advanced RAG query endpoint (search + rerank + answer in one call).
    For the final Q&A section after pipeline is set up.
    """
    if not pipeline_state["embeddings_ready"]:
        return jsonify({"error": "Pipeline not ready. Complete all setup steps first."}), 400

    data = request.get_json()
    query = data.get("question", "").strip()
    if not query:
        return jsonify({"error": "Please provide a question."}), 400

    if not GROQ_API_KEY:
        return jsonify({"error": "GROQ_API_KEY not set in .env"}), 500

    try:
        result = execute_query(query, pipeline_state["qdrant_client"], pipeline_state["test_cases"])
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ══════════════════════════════════════════════════════════════
#  HEALTH
# ══════════════════════════════════════════════════════════════

@app.route("/api/health")
def api_health():
    """Health check endpoint."""
    return jsonify({
        "status": "healthy",
        "pipeline": "Advanced RAG — Interactive Mode",
        "upload_ready": bool(pipeline_state["test_cases"]),
        "embeddings_ready": pipeline_state["embeddings_ready"],
        "components": {
            "embedding": EMBED_MODEL_NAME,
            "reranker": RERANKER_MODEL_NAME,
            "vector_db": "Qdrant",
            "llm": f"Groq/{GROQ_MODEL}",
        },
    })


# ══════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print(f"[START] Advanced RAG at http://localhost:{FLASK_PORT}")
    app.run(debug=FLASK_DEBUG, port=FLASK_PORT)
