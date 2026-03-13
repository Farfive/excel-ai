"""
ColBERT-style Late Interaction Scoring for spreadsheet chunks.

Instead of comparing single query vector vs single chunk vector,
splits both query and chunk into segments (lines/phrases) and computes
MaxSim: for each query segment, find the max similarity to any chunk segment.
Final score = mean of all MaxSim values.

This catches cases where "WACC" in the query matches "Weighted Average Cost of Capital"
in one line and "discount rate" in another line of the same chunk.
"""

import logging
import re
from typing import Any, Dict, List, Optional

import numpy as np

logger = logging.getLogger(__name__)

# Min segments to justify late interaction overhead
_MIN_SEGMENTS = 2
# Max segments per chunk to keep it fast
_MAX_CHUNK_SEGMENTS = 20
# Max segments per query
_MAX_QUERY_SEGMENTS = 6


def _split_into_segments(text: str, max_segments: int) -> List[str]:
    """Split text into meaningful segments (lines, then sentences)."""
    # First split by newlines
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    # If too few lines, try splitting by sentences
    if len(lines) < _MIN_SEGMENTS:
        lines = re.split(r"[.;,]\s+", text)
        lines = [l.strip() for l in lines if l.strip() and len(l.strip()) > 5]

    # If still too few, split by chunks of ~50 chars
    if len(lines) < _MIN_SEGMENTS and len(text) > 100:
        lines = [text[i:i+50] for i in range(0, len(text), 50)]

    return lines[:max_segments]


def colbert_score(
    query: str,
    chunk_text: str,
    embedder: Any,
    query_embedding: Optional[np.ndarray] = None,
) -> float:
    """Compute ColBERT-style MaxSim score between query and chunk.

    Args:
        query: User query string
        chunk_text: Chunk text content
        embedder: LocalEmbedder with .embed() method
        query_embedding: Pre-computed query embedding (optional, for single-vector fallback)

    Returns:
        MaxSim score in [0, 1]
    """
    query_segments = _split_into_segments(query, _MAX_QUERY_SEGMENTS)
    chunk_segments = _split_into_segments(chunk_text, _MAX_CHUNK_SEGMENTS)

    if len(query_segments) < _MIN_SEGMENTS or len(chunk_segments) < _MIN_SEGMENTS:
        # Fall back to single-vector cosine if not enough segments
        if query_embedding is not None:
            chunk_emb = np.array(embedder.embed_single(chunk_text[:500]), dtype=np.float32)
            q = query_embedding if isinstance(query_embedding, np.ndarray) else np.array(query_embedding, dtype=np.float32)
            return float(np.dot(q, chunk_emb) / (np.linalg.norm(q) * np.linalg.norm(chunk_emb) + 1e-10))
        return 0.0

    # Embed all segments in batch
    all_texts = query_segments + chunk_segments
    all_embeddings = np.array(embedder.embed(all_texts), dtype=np.float32)

    n_q = len(query_segments)
    q_embs = all_embeddings[:n_q]       # (n_q, d)
    c_embs = all_embeddings[n_q:]       # (n_c, d)

    # Normalize
    q_norms = np.linalg.norm(q_embs, axis=1, keepdims=True) + 1e-10
    c_norms = np.linalg.norm(c_embs, axis=1, keepdims=True) + 1e-10
    q_embs = q_embs / q_norms
    c_embs = c_embs / c_norms

    # MaxSim: for each query segment, max cosine similarity to any chunk segment
    sim_matrix = q_embs @ c_embs.T  # (n_q, n_c)
    max_sims = sim_matrix.max(axis=1)  # (n_q,)

    return float(np.mean(max_sims))


def colbert_rerank(
    query: str,
    chunks: List[Dict],
    embedder: Any,
    query_embedding: Optional[np.ndarray] = None,
    top_k: int = 15,
    min_segments: int = 3,
) -> List[Dict]:
    """Re-rank chunks using ColBERT-style late interaction scoring.

    Only applies to chunks with enough text to benefit from segment-level matching.
    Chunks below the segment threshold keep their original order.

    Args:
        query: User query
        chunks: List of chunk dicts with 'text'
        embedder: LocalEmbedder
        query_embedding: Pre-computed query embedding
        top_k: Number of results to return
        min_segments: Min lines in chunk to apply ColBERT (otherwise skip)

    Returns:
        Re-ranked list of chunks with 'colbert_score' added to metadata
    """
    if not chunks:
        return []

    scored = []
    skipped = []

    for chunk in chunks:
        text = chunk.get("text", "")
        line_count = text.count("\n") + 1

        if line_count >= min_segments and len(text) > 100:
            try:
                score = colbert_score(query, text, embedder, query_embedding)
                new_chunk = dict(chunk)
                new_chunk.setdefault("metadata", {})
                new_chunk["metadata"]["colbert_score"] = round(score, 4)
                scored.append((score, new_chunk))
            except Exception as e:
                logger.debug(f"ColBERT scoring failed for chunk: {e}")
                skipped.append(chunk)
        else:
            skipped.append(chunk)

    # Sort scored chunks by ColBERT score
    scored.sort(key=lambda x: x[0], reverse=True)
    result = [c for _, c in scored] + skipped

    if scored:
        logger.info(
            f"ColBERT rerank: {len(scored)} scored, {len(skipped)} skipped, "
            f"top={scored[0][0]:.3f}, bottom={scored[-1][0]:.3f}"
        )

    return result[:top_k]
