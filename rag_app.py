"""
RAG Application — ChromaDB + Nomic Embed Text (Ollama) + Groq LLM
==================================================================
- Reads a markdown document and chunks it (800 chars, 120 char overlap)
- Generates embeddings using nomic-embed-text via local Ollama
- Stores embeddings in ChromaDB (local persistent storage)
- Answers user questions using Groq LLM (openai/gpt-oss-120b)
- Serves an HTML UI for chunk visualization and Q&A
"""

import json
import math
import os
import requests
import chromadb
from dotenv import load_dotenv
from flask import Flask, jsonify, request, send_from_directory
from PyPDF2 import PdfReader

# Load environment variables from .env file
load_dotenv()

# ──────────────────────────────────────────────────────────────
# Configuration
# ──────────────────────────────────────────────────────────────
OLLAMA_BASE_URL = "http://localhost:11434"
EMBED_MODEL = "nomic-embed-text"

GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "")  # Loaded from .env file
GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_MODEL = "openai/gpt-oss-120b"

DOCUMENT_PATH = os.path.join(os.path.dirname(__file__), "data", "Product Requirements Document_ VWO Login Dashboard.pdf")
CHROMA_DB_PATH = os.path.join(os.path.dirname(__file__), "chroma_db")

CHUNK_SIZE = 800
CHUNK_OVERLAP = 120
TOP_K = 3  # Number of top chunks to retrieve

COLLECTION_NAME = "vwo_prd_pdf_chunks"

app = Flask(__name__, static_folder=".", static_url_path="")


# ──────────────────────────────────────────────────────────────
# Chunking Logic
# ──────────────────────────────────────────────────────────────
def load_document(path: str) -> str:
    """Load and return the full text of the document (supports PDF and text files)."""
    if path.lower().endswith(".pdf"):
        reader = PdfReader(path)
        text = ""
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
        return text
    else:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()


def chunk_document(text: str, chunk_size: int = CHUNK_SIZE, overlap: int = CHUNK_OVERLAP):
    """
    Split *text* into chunks of ~chunk_size characters with *overlap*.

    Strategy:
    1. Walk through the text with a target window of chunk_size.
    2. For each window, try to break at the last sentence-ending punctuation
       (., !, ?) or newline that falls within the window.
    3. If no such break is found, break at the last whitespace within chunk_size.
    4. Advance the cursor by (break_point − overlap) so successive chunks share
       *overlap* characters.
    """
    chunks = []
    start = 0
    text_len = len(text)
    min_size = int(chunk_size * 0.5)  # Allow chunks down to 50% of target

    while start < text_len:
        end = min(start + chunk_size, text_len)
        window = text[start:end]

        if end >= text_len:
            # Last chunk – take whatever is left
            if len(window.strip()) > 0:
                chunks.append({
                    "id": len(chunks),
                    "start": start,
                    "end": end,
                    "text": window,
                    "length": len(window),
                })
            break

        # Try to find a natural break within the window
        break_point = None
        # Prefer sentence boundaries
        for i in range(len(window) - 1, min_size - 1, -1):
            if window[i] in (".", "!", "?", "\n") and (i + 1 >= len(window) or window[i + 1] in (" ", "\n", "\r")):
                break_point = i + 1
                break

        # Fallback: break at last whitespace
        if break_point is None:
            for i in range(len(window) - 1, min_size - 1, -1):
                if window[i] in (" ", "\t", "\n"):
                    break_point = i + 1
                    break

        # Fallback: hard break at chunk_size
        if break_point is None:
            break_point = len(window)

        chunk_text = window[:break_point]
        chunks.append({
            "id": len(chunks),
            "start": start,
            "end": start + break_point,
            "text": chunk_text,
            "length": len(chunk_text),
        })

        # Advance cursor with overlap
        start = start + break_point - overlap

    return chunks


# ──────────────────────────────────────────────────────────────
# Embedding via Ollama (nomic-embed-text)
# ──────────────────────────────────────────────────────────────
def get_embedding(text: str) -> list:
    """Get embedding vector from Ollama nomic-embed-text."""
    resp = requests.post(
        f"{OLLAMA_BASE_URL}/api/embeddings",
        json={"model": EMBED_MODEL, "prompt": text},
        timeout=60,
    )
    resp.raise_for_status()
    return resp.json()["embedding"]


# ──────────────────────────────────────────────────────────────
# ChromaDB Setup & Operations
# ──────────────────────────────────────────────────────────────
def init_chromadb():
    """Initialize ChromaDB client with persistent local storage."""
    client = chromadb.PersistentClient(path=CHROMA_DB_PATH)
    return client


def populate_chromadb(client, chunks: list, force_rebuild: bool = False):
    """
    Store chunk embeddings in ChromaDB.
    If the collection already exists with the right count, skip re-embedding.
    """
    # Check if collection exists and is populated
    existing_collections = [c.name for c in client.list_collections()]

    if COLLECTION_NAME in existing_collections and not force_rebuild:
        collection = client.get_collection(name=COLLECTION_NAME)
        if collection.count() == len(chunks):
            print(f"   -> ChromaDB collection '{COLLECTION_NAME}' already populated with {collection.count()} chunks. Skipping.")
            return collection

    # Delete old collection if rebuilding
    if COLLECTION_NAME in existing_collections:
        client.delete_collection(name=COLLECTION_NAME)
        print("   -> Deleted old ChromaDB collection.")

    collection = client.create_collection(
        name=COLLECTION_NAME,
        metadata={"hnsw:space": "cosine"}
    )

    print("   -> Generating embeddings and storing in ChromaDB...")
    for i, chunk in enumerate(chunks):
        emb = get_embedding(chunk["text"])
        collection.add(
            ids=[f"chunk_{chunk['id']}"],
            embeddings=[emb],
            documents=[chunk["text"]],
            metadatas=[{
                "chunk_id": chunk["id"],
                "start": chunk["start"],
                "end": chunk["end"],
                "length": chunk["length"],
            }],
        )
        print(f"   -> Embedded & stored chunk {i + 1}/{len(chunks)}")

    return collection


def query_chromadb(collection, query: str, top_k: int = TOP_K):
    """Query ChromaDB with the embedded query to retrieve top-k similar chunks."""
    query_emb = get_embedding(query)
    results = collection.query(
        query_embeddings=[query_emb],
        n_results=top_k,
        include=["documents", "metadatas", "distances"],
    )
    return results


# ──────────────────────────────────────────────────────────────
# Groq LLM Answer Generation
# ──────────────────────────────────────────────────────────────
def generate_answer(query: str, retrieved_chunks: list) -> str:
    """Use Groq LLM to answer the question using retrieved context."""
    context = "\n\n---\n\n".join(
        [f"[Chunk {c['chunk_id']}]:\n{c['text']}" for c in retrieved_chunks]
    )
    system_prompt = """You are a helpful assistant. Answer the user's question based ONLY on the following context chunks from a VWO Project Requirement Document.

If the answer is not in the context, say "I could not find the answer in the provided document."

Always mention which chunk(s) you used to derive your answer."""

    user_prompt = f"""CONTEXT:
{context}

USER QUESTION: {query}

Provide a clear, concise answer and mention which chunk(s) you used."""

    headers = {
        "Authorization": f"Bearer {GROQ_API_KEY}",
        "Content-Type": "application/json",
    }

    payload = {
        "model": GROQ_MODEL,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": 0.3,
        "max_tokens": 1024,
    }

    resp = requests.post(GROQ_API_URL, headers=headers, json=payload, timeout=60)
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"]


# ──────────────────────────────────────────────────────────────
# Pre-compute on startup
# ──────────────────────────────────────────────────────────────
print("=" * 60)
print("  RAG System — ChromaDB + Nomic Embed + Groq LLM")
print("=" * 60)

print("\n[1/4] Loading document...")
raw_text = load_document(DOCUMENT_PATH)
print(f"   -> Document loaded: {len(raw_text)} characters")

print("\n[2/4] Chunking document (size={}, overlap={})...".format(CHUNK_SIZE, CHUNK_OVERLAP))
chunks = chunk_document(raw_text)
print(f"   -> {len(chunks)} chunks created")

print("\n[3/4] Initializing ChromaDB (persistent at '{}'')...".format(CHROMA_DB_PATH))
chroma_client = init_chromadb()
collection = populate_chromadb(chroma_client, chunks)
print(f"   -> ChromaDB ready. Collection: '{COLLECTION_NAME}', Count: {collection.count()}")

print("\n[4/4] Checking Groq API key...")
if GROQ_API_KEY:
    print(f"   -> Groq API key loaded (model: {GROQ_MODEL})")
else:
    print("   ⚠️  No GROQ_API_KEY set! Set it via environment variable or in the script.")
    print("      export GROQ_API_KEY='your-key-here'  (Linux/Mac)")
    print("      $env:GROQ_API_KEY='your-key-here'    (PowerShell)")

print("\n[OK] RAG system ready!\n")


# ──────────────────────────────────────────────────────────────
# Flask Routes
# ──────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return send_from_directory(".", "index.html")


@app.route("/api/chunks")
def api_chunks():
    """Return all chunks with embedding metadata for visualization."""
    # Retrieve embeddings from ChromaDB for display
    all_data = collection.get(
        ids=[f"chunk_{c['id']}" for c in chunks],
        include=["embeddings"],
    )

    enriched_chunks = []
    has_embeddings = all_data["embeddings"] is not None and len(all_data["embeddings"]) > 0
    for i, chunk in enumerate(chunks):
        emb = list(all_data["embeddings"][i]) if has_embeddings else []
        if len(emb) > 0:
            n = len(emb)
            emb_min = min(emb)
            emb_max = max(emb)
            emb_mean = sum(emb) / n
            emb_std = math.sqrt(sum((x - emb_mean) ** 2 for x in emb) / n)
            magnitude = math.sqrt(sum(x * x for x in emb))
            embedding_info = {
                "dimension": n,
                "magnitude": round(magnitude, 4),
                "min": round(emb_min, 6),
                "max": round(emb_max, 6),
                "mean": round(emb_mean, 6),
                "std": round(emb_std, 6),
                "preview": [round(v, 6) for v in emb[:8]],
            }
        else:
            embedding_info = {}

        enriched_chunks.append({
            **chunk,
            "embedding": embedding_info,
        })

    return jsonify({
        "total_chars": len(raw_text),
        "num_chunks": len(chunks),
        "chunk_size": CHUNK_SIZE,
        "overlap": CHUNK_OVERLAP,
        "embed_model": EMBED_MODEL,
        "llm_model": GROQ_MODEL,
        "llm_provider": "Groq",
        "vector_store": "ChromaDB",
        "embed_dim": enriched_chunks[0]["embedding"].get("dimension", 0) if enriched_chunks else 0,
        "chunks": enriched_chunks,
    })


@app.route("/api/ask", methods=["POST"])
def api_ask():
    """Answer a question using RAG with ChromaDB retrieval and Groq LLM."""
    data = request.get_json()
    query = data.get("question", "").strip()
    if not query:
        return jsonify({"error": "Please provide a question."}), 400

    if not GROQ_API_KEY:
        return jsonify({"error": "GROQ_API_KEY is not set. Please set it as an environment variable."}), 500

    try:
        # Retrieve from ChromaDB
        results = query_chromadb(collection, query, top_k=TOP_K)

        retrieved_chunks = []
        for i in range(len(results["ids"][0])):
            meta = results["metadatas"][0][i]
            # ChromaDB cosine distance: similarity = 1 - distance
            distance = results["distances"][0][i]
            similarity = 1 - distance

            retrieved_chunks.append({
                "chunk_id": meta["chunk_id"],
                "similarity": round(similarity, 4),
                "text": results["documents"][0][i],
                "start": meta["start"],
                "end": meta["end"],
                "length": meta["length"],
            })

        # Generate answer via Groq
        answer = generate_answer(query, retrieved_chunks)

        return jsonify({
            "answer": answer,
            "retrieved_chunks": retrieved_chunks,
        })
    except requests.exceptions.ConnectionError as e:
        error_msg = str(e)
        if "11434" in error_msg:
            return jsonify({"error": "Cannot connect to Ollama. Make sure Ollama is running on localhost:11434."}), 503
        elif "groq" in error_msg.lower():
            return jsonify({"error": "Cannot connect to Groq API. Check your internet connection."}), 503
        return jsonify({"error": f"Connection error: {error_msg}"}), 503
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/config")
def api_config():
    """Return current configuration for display in the UI."""
    return jsonify({
        "embed_model": EMBED_MODEL,
        "llm_model": GROQ_MODEL,
        "llm_provider": "Groq",
        "vector_store": "ChromaDB (local)",
        "chunk_size": CHUNK_SIZE,
        "chunk_overlap": CHUNK_OVERLAP,
        "top_k": TOP_K,
        "groq_api_key_set": bool(GROQ_API_KEY),
        "chroma_db_path": CHROMA_DB_PATH,
        "collection_name": COLLECTION_NAME,
        "total_chunks": collection.count(),
    })


if __name__ == "__main__":
    print("[START] RAG server at http://localhost:5000")
    app.run(debug=False, port=5000)
