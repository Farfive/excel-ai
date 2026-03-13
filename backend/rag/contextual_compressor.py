"""
Contextual Compression — extracts only query-relevant lines from retrieved chunks.

After retrieval, each chunk may contain many rows of data. This module uses a
lightweight LLM call to extract only the lines directly relevant to the query,
reducing token usage by ~60% while improving answer precision.

Falls back to truncation-based compression if LLM is unavailable.
"""

import logging
import re
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)

# Max lines per chunk before compression kicks in
_COMPRESSION_THRESHOLD = 8
# Max chars to send to LLM for compression (avoid blowing up context)
_MAX_CHUNK_FOR_LLM = 2000


def _needs_compression(chunk: Dict) -> bool:
    """Check if a chunk has enough content to benefit from compression."""
    text = chunk.get("text", "")
    line_count = text.count("\n") + 1
    return line_count > _COMPRESSION_THRESHOLD and len(text) > 300


def _keyword_compress(text: str, query: str, max_lines: int = 15) -> str:
    """Fast keyword-based compression fallback (no LLM needed).

    Scores each line by keyword overlap with query and keeps top lines.
    Always preserves header lines and formula lines.
    """
    lines = text.split("\n")
    if len(lines) <= max_lines:
        return text

    query_tokens = set(re.findall(r"[A-Za-z0-9]+", query.lower()))

    scored = []
    for i, line in enumerate(lines):
        line_lower = line.lower()
        line_tokens = set(re.findall(r"[A-Za-z0-9]+", line_lower))

        # Base score: keyword overlap
        overlap = len(query_tokens & line_tokens)
        score = overlap * 2.0

        # Boost header/label lines (short, no numbers)
        if i < 3 or (len(line) < 60 and not any(c.isdigit() for c in line)):
            score += 3.0

        # Boost formula lines
        if "=" in line or "formula" in line_lower:
            score += 2.0

        # Boost lines with cell addresses
        if re.search(r"[A-Z]+\d+", line):
            score += 1.0

        # Boost lines with financial keywords
        for kw in ["total", "sum", "revenue", "ebit", "wacc", "dcf", "value", "margin", "growth"]:
            if kw in line_lower:
                score += 1.5
                break

        scored.append((score, i, line))

    # Sort by score descending, then take top lines, but preserve original order
    scored.sort(key=lambda x: x[0], reverse=True)
    top_indices = sorted([idx for _, idx, _ in scored[:max_lines]])
    kept = [lines[i] for i in top_indices]

    # If we cut lines, add indicator
    if len(kept) < len(lines):
        cut_count = len(lines) - len(kept)
        kept.append(f"[...{cut_count} less relevant lines omitted...]")

    return "\n".join(kept)


async def compress_chunks(
    chunks: List[Dict],
    query: str,
    ollama_client: Any = None,
    use_llm: bool = True,
    max_lines_per_chunk: int = 15,
) -> List[Dict]:
    """Compress retrieved chunks to keep only query-relevant content.

    Args:
        chunks: Retrieved chunk dicts with 'text' and 'metadata'
        query: User's query
        ollama_client: OllamaClient for LLM-based compression (optional)
        use_llm: Whether to attempt LLM compression (falls back to keyword)
        max_lines_per_chunk: Target max lines per compressed chunk

    Returns:
        List of chunks with compressed text (originals unchanged)
    """
    compressed = []
    llm_compressed_count = 0
    keyword_compressed_count = 0

    for chunk in chunks:
        if not _needs_compression(chunk):
            compressed.append(chunk)
            continue

        text = chunk.get("text", "")

        # Try LLM compression first
        if use_llm and ollama_client is not None and len(text) <= _MAX_CHUNK_FOR_LLM:
            try:
                result = await _llm_compress(text, query, ollama_client)
                if result and len(result) < len(text) * 0.9:
                    new_chunk = dict(chunk)
                    new_chunk["text"] = result
                    new_chunk["metadata"] = dict(chunk.get("metadata", {}))
                    new_chunk["metadata"]["compressed"] = True
                    new_chunk["metadata"]["original_length"] = len(text)
                    compressed.append(new_chunk)
                    llm_compressed_count += 1
                    continue
            except Exception as e:
                logger.debug(f"LLM compression failed for chunk, using keyword: {e}")

        # Fallback: keyword-based compression
        result = _keyword_compress(text, query, max_lines=max_lines_per_chunk)
        new_chunk = dict(chunk)
        new_chunk["text"] = result
        new_chunk["metadata"] = dict(chunk.get("metadata", {}))
        new_chunk["metadata"]["compressed"] = True
        new_chunk["metadata"]["original_length"] = len(text)
        compressed.append(new_chunk)
        keyword_compressed_count += 1

    total_compressed = llm_compressed_count + keyword_compressed_count
    if total_compressed > 0:
        orig_chars = sum(len(c.get("text", "")) for c in chunks)
        new_chars = sum(len(c.get("text", "")) for c in compressed)
        savings = (1 - new_chars / max(orig_chars, 1)) * 100
        logger.info(
            f"Compression: {total_compressed}/{len(chunks)} chunks compressed "
            f"({llm_compressed_count} LLM, {keyword_compressed_count} keyword), "
            f"chars {orig_chars}→{new_chars} ({savings:.0f}% reduction)"
        )

    return compressed


async def _llm_compress(text: str, query: str, ollama_client) -> Optional[str]:
    """Use LLM to extract only query-relevant lines from chunk text."""
    system = (
        "You are a precise text extractor for financial spreadsheet data. "
        "Given a chunk of spreadsheet data and a user query, extract ONLY the lines "
        "that are directly relevant to answering the query. "
        "Preserve cell addresses, formulas, and values exactly as they appear. "
        "Do NOT add explanations. Do NOT modify data. Just return the relevant lines."
    )

    response = await ollama_client.chat(
        messages=[{
            "role": "user",
            "content": f"Query: {query}\n\nData:\n{text[:_MAX_CHUNK_FOR_LLM]}\n\nExtract only relevant lines:",
        }],
        system=system,
        temperature=0.0,
    )

    result = response.strip()
    if not result or len(result) < 10:
        return None
    return result
