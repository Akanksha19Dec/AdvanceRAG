"""
Vector Store — Qdrant Integration
====================================
Handles collection creation, batch upsert, vector search, and filtered scroll.
Uses local Qdrant (embedded mode, no server needed).
"""

from qdrant_client import QdrantClient
from qdrant_client.models import (
    Distance,
    FieldCondition,
    Filter,
    MatchValue,
    PointStruct,
    VectorParams,
)

from config import (
    COLLECTION_NAME,
    EMBED_DIM,
    QDRANT_BATCH_SIZE,
    QDRANT_PATH,
    TOP_K_RETRIEVE,
)
from embeddings import embed_texts


def init_qdrant() -> QdrantClient:
    """Initialize local Qdrant client (embedded mode)."""
    return QdrantClient(path=QDRANT_PATH)


def populate_qdrant(client: QdrantClient, cases: list[dict]) -> dict:
    """
    Create collection and upsert all test case embeddings.
    Skips if collection already has the expected number of points.
    """
    existing = client.collection_exists(COLLECTION_NAME)
    if existing:
        info = client.get_collection(COLLECTION_NAME)
        if info.points_count == len(cases):
            print(
                f"   -> Qdrant collection already has {info.points_count} points. Skipping."
            )
            return {
                "points_count": info.points_count,
                "status": "cached",
            }
        # Collection exists but count doesn't match — rebuild
        client.delete_collection(COLLECTION_NAME)

    client.create_collection(
        collection_name=COLLECTION_NAME,
        vectors_config=VectorParams(size=EMBED_DIM, distance=Distance.COSINE),
    )

    # Embed all test case texts
    print(f"   -> Embedding {len(cases)} test cases with BGE-M3...")
    texts = [tc["text"] for tc in cases]
    embs = embed_texts(texts)

    # Batch upsert
    print("   -> Storing in Qdrant...")
    total = len(cases)
    for start in range(0, total, QDRANT_BATCH_SIZE):
        end = min(start + QDRANT_BATCH_SIZE, total)
        points = []
        for i in range(start, end):
            tc = cases[i]
            points.append(
                PointStruct(
                    id=tc["id"],
                    vector=embs[i],
                    payload={
                        "chunk_id": tc["id"],
                        "tc_id": tc["tc_id"],
                        "module": tc["module"],
                        "category": tc["category"],
                        "priority": tc["priority"],
                        "length": tc["length"],
                        "description": tc["description"],
                        "text": tc["text"],
                    },
                )
            )
        client.upsert(collection_name=COLLECTION_NAME, points=points)
        print(f"      -> Stored {end}/{total}", end="\r")
    print()

    info = client.get_collection(COLLECTION_NAME)
    return {
        "points_count": info.points_count,
        "status": "created",
    }


def search_qdrant(
    client: QdrantClient,
    query_vector: list[float],
    top_k: int = TOP_K_RETRIEVE,
) -> list:
    """Perform vector similarity search in Qdrant."""
    response = client.query_points(
        collection_name=COLLECTION_NAME,
        query=query_vector,
        limit=top_k,
        with_payload=True,
    )
    return response.points


def search_by_module(
    client: QdrantClient,
    module_name: str,
    limit: int = 40,
) -> list:
    """Retrieve test cases from a specific module using Qdrant filter."""
    results, _ = client.scroll(
        collection_name=COLLECTION_NAME,
        scroll_filter=Filter(
            must=[FieldCondition(key="module", match=MatchValue(value=module_name))]
        ),
        limit=limit,
        with_payload=True,
    )
    return results
