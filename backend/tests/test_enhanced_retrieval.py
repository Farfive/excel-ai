"""
End-to-end test for all 5 enhanced retrieval modules:
1. Query Decomposition
2. Contextual Compression
3. Formula Chain Unrolling
4. ColBERT Late Interaction
5. CGASR Index Persistence
"""

import asyncio
import os
import sys
import tempfile
import time

import numpy as np

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from parser.xlsx_parser import XLSXParser
from parser.graph_builder import DependencyGraphBuilder
from rag.chunker import ChunkBuilder
from rag.cgasr_index import build_cgasr_index, CGASRIndex
from rag.cgasr_retriever import CGASRRetriever
from rag.query_decomposer import is_complex_query, decompose_query
from rag.contextual_compressor import compress_chunks, _keyword_compress
from rag.formula_chain import unroll_formula_chain, extract_seed_addresses
from rag.colbert_scorer import colbert_score, colbert_rerank


# ── Fake embedder (deterministic, fast) ──────────────────────────────────
class FakeEmbedder:
    def __init__(self, dim=768):
        self.dim = dim

    def embed(self, texts):
        rng = np.random.RandomState(42)
        return [self._hash_embed(t, rng) for t in texts]

    def embed_single(self, text):
        return self._hash_embed(text, np.random.RandomState(42))

    def _hash_embed(self, text, rng):
        seed = hash(text) % (2**31)
        local_rng = np.random.RandomState(seed)
        v = local_rng.randn(self.dim).astype(np.float32)
        v /= np.linalg.norm(v) + 1e-10
        return v.tolist()


# ── Fake OllamaClient for query decomposition ───────────────────────────
class FakeOllama:
    async def chat(self, messages, system=None, temperature=0.1):
        query = messages[0]["content"]
        if "decompose" in system.lower() or "decompos" in query.lower():
            return '["What is WACC?", "How does WACC affect Enterprise Value?", "What cells depend on WACC?"]'
        return "Hypothetical answer about the workbook data."


def get_test_workbook():
    xlsx_path = os.path.join(
        os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
        "advanced_financial_model.xlsx",
    )
    if not os.path.exists(xlsx_path):
        return None, None, None, None
    parser = XLSXParser()
    wb = parser.parse(xlsx_path)
    builder = DependencyGraphBuilder()
    graph = builder.build(wb)
    graph = builder.run_algorithms(graph, wb)
    chunk_builder = ChunkBuilder()
    chunks = chunk_builder.build_chunks(wb, graph, "test")
    chunk_dicts = [{"chunk_id": c.chunk_id, "text": c.text, "metadata": c.metadata} for c in chunks]
    return wb, graph, chunks, chunk_dicts


def test_query_decomposition():
    print("\n" + "=" * 60)
    print("TEST 1: Query Decomposition")
    print("=" * 60)

    # Simple queries should NOT be decomposed
    assert not is_complex_query("What is WACC?"), "Simple query detected as complex"
    assert not is_complex_query("Show the P&L sheet"), "Simple query detected as complex"

    # Complex queries SHOULD be decomposed
    assert is_complex_query("How does WACC affect the enterprise value?"), "Complex query not detected"
    assert is_complex_query("What if Revenue grows 5% faster and COGS decrease?"), "Complex query not detected"
    assert is_complex_query("Compare the sensitivity of EV to WACC vs TGR"), "Complex query not detected"
    assert is_complex_query("Explain how Revenue flows through the P&L to FCF"), "Complex query not detected"

    # Test actual decomposition with fake LLM
    ollama = FakeOllama()
    sub_queries = asyncio.get_event_loop().run_until_complete(
        decompose_query("How does WACC affect the enterprise value?", ollama)
    )
    assert len(sub_queries) > 1, f"Expected >1 sub-queries, got {len(sub_queries)}"
    print(f"  ✅ Complex query decomposed into {len(sub_queries)} sub-queries: {sub_queries}")

    # Simple query should pass through
    sub_queries_simple = asyncio.get_event_loop().run_until_complete(
        decompose_query("What is WACC?", ollama)
    )
    assert len(sub_queries_simple) == 1, f"Simple query should not decompose, got {len(sub_queries_simple)}"
    print(f"  ✅ Simple query passed through unchanged")

    print("  ✅ All query decomposition tests passed")


def test_contextual_compression():
    print("\n" + "=" * 60)
    print("TEST 2: Contextual Compression")
    print("=" * 60)

    # Create a chunk with many lines
    long_text = "\n".join([
        "Revenue!A1: Revenue Breakdown",
        "Revenue!B1: 2021",
        "Revenue!A2: Product A",
        "Revenue!B2: 1000000",
        "Revenue!A3: Product B",
        "Revenue!B3: 500000",
        "Revenue!A4: Product C",
        "Revenue!B4: 250000",
        "Revenue!A5: Services",
        "Revenue!B5: 750000",
        "Revenue!A6: Consulting",
        "Revenue!B6: 300000",
        "Revenue!A7: Total Revenue = SUM(B2:B6)",
        "Revenue!B7: 2800000",
        "P&L!A10: COGS",
        "P&L!B10: 1200000",
        "P&L!A11: Gross Profit = Revenue - COGS",
        "P&L!B11: 1600000",
    ])

    chunks = [{"chunk_id": "test1", "text": long_text, "metadata": {"sheet": "Revenue"}}]

    # Keyword compression
    result = _keyword_compress(long_text, "What is Total Revenue?", max_lines=8)
    lines = result.split("\n")
    assert len(lines) <= 10, f"Expected ≤10 lines, got {len(lines)}"
    assert any("Total Revenue" in l for l in lines), "Total Revenue line should be preserved"
    print(f"  ✅ Keyword compression: {long_text.count(chr(10))+1} lines → {len(lines)} lines")

    # Async compression
    compressed = asyncio.get_event_loop().run_until_complete(
        compress_chunks(chunks, "What is Total Revenue?", use_llm=False)
    )
    assert len(compressed) == 1
    assert len(compressed[0]["text"]) <= len(long_text)
    print(f"  ✅ Chunk compression: {len(long_text)} → {len(compressed[0]['text'])} chars")

    # Short chunks should not be compressed
    short_chunks = [{"chunk_id": "short", "text": "WACC = 10%", "metadata": {}}]
    compressed_short = asyncio.get_event_loop().run_until_complete(
        compress_chunks(short_chunks, "What is WACC?", use_llm=False)
    )
    assert compressed_short[0]["text"] == "WACC = 10%", "Short chunk should not be compressed"
    print(f"  ✅ Short chunks preserved unchanged")

    print("  ✅ All contextual compression tests passed")


def test_formula_chain():
    print("\n" + "=" * 60)
    print("TEST 3: Formula Chain Unrolling")
    print("=" * 60)

    wb, graph, chunks, chunk_dicts = get_test_workbook()
    if wb is None:
        print("  ⚠️  Skipped (no test workbook)")
        return

    # Get seed addresses from top chunks
    seeds = extract_seed_addresses(chunk_dicts[:5], max_seeds=8)
    assert len(seeds) > 0, "Should find seed addresses"
    print(f"  Seeds: {seeds[:5]}...")

    # Unroll formula chain
    chain = unroll_formula_chain(seeds, graph, wb, max_depth=3, direction="both")
    assert len(chain) > 0, "Chain should not be empty"
    assert "FORMULA CHAIN" in chain, "Should have header"

    lines = chain.split("\n")
    print(f"  ✅ Formula chain: {len(lines)} lines from {len(seeds)} seeds")
    # Print first 5 lines as sample
    for line in lines[:5]:
        print(f"      {line}")

    # Test upstream-only direction
    chain_up = unroll_formula_chain(seeds[:3], graph, wb, max_depth=2, direction="up")
    assert len(chain_up) > 0
    print(f"  ✅ Upstream chain: {len(chain_up.split(chr(10)))} lines")

    print("  ✅ All formula chain tests passed")


def test_colbert_scorer():
    print("\n" + "=" * 60)
    print("TEST 4: ColBERT Late Interaction")
    print("=" * 60)

    embedder = FakeEmbedder()

    # Multi-line chunk should use MaxSim
    chunk_text = "\n".join([
        "Assumptions!B13: WACC = 10%",
        "Assumptions!B14: Terminal Growth Rate = 2.5%",
        "Assumptions!B15: Shares Outstanding = 100M",
        "DCF!B20: Enterprise Value = SUM(PV FCFs) + PV Terminal Value",
        "DCF!B21: Equity Value = EV - Net Debt",
    ])

    score = colbert_score("What is WACC?", chunk_text, embedder)
    assert 0 <= score <= 1, f"Score should be in [0,1], got {score}"
    print(f"  ✅ ColBERT score for WACC query: {score:.4f}")

    # Reranking
    chunks = [
        {"chunk_id": "a", "text": chunk_text, "metadata": {}},
        {"chunk_id": "b", "text": "Revenue!A1: Revenue\nRevenue!B1: 1000000\nRevenue!C1: 1200000", "metadata": {}},
        {"chunk_id": "c", "text": "WACC", "metadata": {}},  # Too short for ColBERT
    ]

    reranked = colbert_rerank("What is WACC?", chunks, embedder, top_k=3)
    assert len(reranked) == 3
    # Check that scored chunks have colbert_score in metadata
    scored_count = sum(1 for c in reranked if c.get("metadata", {}).get("colbert_score") is not None)
    print(f"  ✅ ColBERT reranked: {scored_count}/3 chunks scored, {3-scored_count} skipped (too short)")

    print("  ✅ All ColBERT tests passed")


def test_cgasr_persistence():
    print("\n" + "=" * 60)
    print("TEST 5: CGASR Index Persistence")
    print("=" * 60)

    wb, graph, chunks, chunk_dicts = get_test_workbook()
    if wb is None:
        print("  ⚠️  Skipped (no test workbook)")
        return

    embedder = FakeEmbedder()

    # Build index
    t0 = time.time()
    idx = build_cgasr_index(wb, graph, embedder=embedder)
    build_ms = (time.time() - t0) * 1000
    print(f"  Build: N={idx.N}, K={idx.K} in {build_ms:.0f}ms")

    # Save to temp file
    with tempfile.NamedTemporaryFile(suffix=".cgasr", delete=False) as tmp:
        tmp_path = tmp.name

    try:
        t0 = time.time()
        idx.save(tmp_path)
        save_ms = (time.time() - t0) * 1000
        size_mb = os.path.getsize(tmp_path) / (1024 * 1024)
        print(f"  Save: {size_mb:.1f}MB in {save_ms:.0f}ms")

        # Load from disk
        t0 = time.time()
        idx2 = CGASRIndex.load(tmp_path)
        load_ms = (time.time() - t0) * 1000
        print(f"  Load: N={idx2.N}, K={idx2.K} in {load_ms:.0f}ms")

        # Verify data integrity
        assert idx2.N == idx.N, f"N mismatch: {idx2.N} vs {idx.N}"
        assert idx2.K == idx.K, f"K mismatch: {idx2.K} vs {idx.K}"
        assert np.allclose(idx2.cell_embeddings, idx.cell_embeddings, atol=1e-6), "Embeddings mismatch"
        assert np.allclose(idx2.eigenvalues, idx.eigenvalues, atol=1e-6), "Eigenvalues mismatch"
        assert idx2.roles == idx.roles, "Roles mismatch"
        assert idx2.cell_to_idx == idx.cell_to_idx, "cell_to_idx mismatch"
        print(f"  ✅ Data integrity verified (embeddings, eigenvalues, roles, cell_to_idx)")

        # Verify retrieval gives same results
        q_emb = np.array(embedder.embed_single("What is WACC?"), dtype=np.float32)
        r1 = CGASRRetriever(idx).retrieve(q_emb, chunk_dicts[:50], k=5)
        r2 = CGASRRetriever(idx2).retrieve(q_emb, chunk_dicts[:50], k=5)

        ids1 = [r["chunk"]["chunk_id"] for r in r1]
        ids2 = [r["chunk"]["chunk_id"] for r in r2]
        assert ids1 == ids2, f"Retrieval results differ after load: {ids1} vs {ids2}"
        print(f"  ✅ Retrieval results identical after save/load")

        # Load should be much faster than build
        assert load_ms < build_ms * 0.5, f"Load ({load_ms:.0f}ms) should be <50% of build ({build_ms:.0f}ms)"
        print(f"  ✅ Load {load_ms:.0f}ms << Build {build_ms:.0f}ms ({load_ms/build_ms*100:.0f}%)")

    finally:
        os.unlink(tmp_path)

    print("  ✅ All persistence tests passed")


def test_full_pipeline_integration():
    print("\n" + "=" * 60)
    print("TEST 6: Full Pipeline Integration")
    print("=" * 60)

    wb, graph, chunks, chunk_dicts = get_test_workbook()
    if wb is None:
        print("  ⚠️  Skipped (no test workbook)")
        return

    embedder = FakeEmbedder()

    # 1. Query decomposition
    ollama = FakeOllama()
    sub_queries = asyncio.get_event_loop().run_until_complete(
        decompose_query("How does WACC affect the enterprise value and what are the key assumptions?", ollama)
    )
    print(f"  1. Decomposed into {len(sub_queries)} sub-queries")

    # 2. Formula chain from top chunks
    seeds = extract_seed_addresses(chunk_dicts[:5])
    chain = unroll_formula_chain(seeds, graph, wb, max_depth=3)
    chain_lines = chain.count("\n") + 1 if chain else 0
    print(f"  2. Formula chain: {chain_lines} lines from {len(seeds)} seeds")

    # 3. ColBERT rerank
    q_emb = np.array(embedder.embed_single("What is WACC?"), dtype=np.float32)
    reranked = colbert_rerank("What is WACC?", chunk_dicts[:20], embedder, query_embedding=q_emb, top_k=10)
    colbert_scored = sum(1 for c in reranked if c.get("metadata", {}).get("colbert_score") is not None)
    print(f"  3. ColBERT reranked: {colbert_scored}/{len(reranked)} scored")

    # 4. Contextual compression
    compressed = asyncio.get_event_loop().run_until_complete(
        compress_chunks(reranked[:10], "What is WACC?", use_llm=False)
    )
    orig_chars = sum(len(c.get("text", "")) for c in reranked[:10])
    comp_chars = sum(len(c.get("text", "")) for c in compressed)
    savings = (1 - comp_chars / max(orig_chars, 1)) * 100
    print(f"  4. Compression: {orig_chars} → {comp_chars} chars ({savings:.0f}% reduction)")

    # 5. CGASR index build + persistence
    idx = build_cgasr_index(wb, graph, embedder=embedder)
    with tempfile.NamedTemporaryFile(suffix=".cgasr", delete=False) as tmp:
        idx.save(tmp.name)
        idx2 = CGASRIndex.load(tmp.name)
        os.unlink(tmp.name)
    print(f"  5. CGASR persistence: OK (N={idx.N})")

    print("\n  ✅ Full pipeline integration test passed!")


if __name__ == "__main__":
    print("╔══════════════════════════════════════════════════════════════════╗")
    print("║   Enhanced Retrieval Pipeline — Full Test Suite                 ║")
    print("╚══════════════════════════════════════════════════════════════════╝")

    test_query_decomposition()
    test_contextual_compression()
    test_formula_chain()
    test_colbert_scorer()
    test_cgasr_persistence()
    test_full_pipeline_integration()

    print("\n" + "=" * 60)
    print("ALL TESTS PASSED ✅")
    print("=" * 60)
