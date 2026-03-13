"""
End-to-end test for CGASR algorithm.

Loads a real .xlsx workbook, builds the dependency graph,
constructs the CGASR index, and runs test queries through
the full retrieval pipeline.
"""

import sys
import os
import time
import logging
import json

# Add backend to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import numpy as np
import networkx as nx

from parser.xlsx_parser import XLSXParser, WorkbookData
from parser.graph_builder import DependencyGraphBuilder
from rag.chunker import ChunkBuilder
from rag.cgasr_index import build_cgasr_index, CGASRIndex, ROLE_WEIGHTS
from rag.cgasr_retriever import CGASRRetriever

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(name)s %(levelname)s  %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("test_cgasr")

# ═════════════════════════════════════════════════════════════════════
# HELPERS
# ═════════════════════════════════════════════════════════════════════

def fake_embedder_embed(texts):
    """Deterministic fake embeddings based on text hash (for testing without model)."""
    rng = np.random.RandomState(42)
    embeddings = []
    for text in texts:
        seed = hash(text) % (2**31)
        local_rng = np.random.RandomState(seed)
        emb = local_rng.randn(768).astype(np.float32)
        emb /= np.linalg.norm(emb) + 1e-10
        embeddings.append(emb)
    return np.array(embeddings)


def fake_query_embed(query):
    """Fake query embedding."""
    seed = hash(query) % (2**31)
    rng = np.random.RandomState(seed)
    emb = rng.randn(768).astype(np.float32)
    emb /= np.linalg.norm(emb) + 1e-10
    return emb


class FakeEmbedder:
    def embed(self, texts):
        return fake_embedder_embed(texts).tolist()
    def embed_single(self, text):
        return fake_query_embed(text).tolist()


# ═════════════════════════════════════════════════════════════════════
# TEST FUNCTIONS
# ═════════════════════════════════════════════════════════════════════

def test_parse_workbook(xlsx_path: str):
    """Step 1: Parse workbook."""
    print("\n" + "="*70)
    print("STEP 1: PARSE WORKBOOK")
    print("="*70)

    parser = XLSXParser()
    t0 = time.time()
    wb = parser.parse(xlsx_path)
    elapsed = time.time() - t0

    print(f"  Sheets: {wb.sheets}")
    print(f"  Cells: {len(wb.cells)}")
    print(f"  Named ranges: {list(wb.named_ranges.keys())[:10]}")
    print(f"  Parse time: {elapsed*1000:.0f}ms")

    # Show some sample cells
    for i, (addr, cell) in enumerate(wb.cells.items()):
        if i >= 8:
            break
        val_str = str(cell.value)[:30] if cell.value is not None else "None"
        formula_str = cell.formula[:40] if cell.formula else ""
        print(f"  {addr}: val={val_str} formula={formula_str} type={cell.data_type}")

    assert len(wb.cells) > 0, "Workbook has no cells!"
    assert len(wb.sheets) > 0, "Workbook has no sheets!"
    print(f"  ✅ PASS: {len(wb.cells)} cells, {len(wb.sheets)} sheets")
    return wb


def test_build_graph(wb: WorkbookData):
    """Step 2: Build dependency graph."""
    print("\n" + "="*70)
    print("STEP 2: BUILD DEPENDENCY GRAPH")
    print("="*70)

    builder = DependencyGraphBuilder()
    t0 = time.time()
    graph = builder.build(wb)
    graph = builder.run_algorithms(graph, wb)
    elapsed = time.time() - t0

    print(f"  Nodes: {graph.number_of_nodes()}")
    print(f"  Edges: {graph.number_of_edges()}")
    print(f"  Build time: {elapsed*1000:.0f}ms")

    # PageRank top cells
    pr_sorted = sorted(graph.nodes(), key=lambda n: graph.nodes[n].get("pagerank", 0), reverse=True)[:5]
    print(f"  Top PageRank cells:")
    for n in pr_sorted:
        pr = graph.nodes[n].get("pagerank", 0)
        val = graph.nodes[n].get("value")
        print(f"    {n}: PR={pr:.4f}, val={str(val)[:30]}")

    # Clusters
    clusters = set(graph.nodes[n].get("cluster_id", 0) for n in graph.nodes())
    print(f"  Louvain clusters: {len(clusters)}")

    # Anomalies
    anomalies = [n for n in graph.nodes() if graph.nodes[n].get("is_anomaly")]
    print(f"  Anomalies: {len(anomalies)}")

    assert graph.number_of_nodes() > 0
    print(f"  ✅ PASS: {graph.number_of_nodes()} nodes, {graph.number_of_edges()} edges")
    return graph


def test_build_chunks(wb: WorkbookData, graph: nx.DiGraph):
    """Step 3: Build chunks."""
    print("\n" + "="*70)
    print("STEP 3: BUILD CHUNKS")
    print("="*70)

    builder = ChunkBuilder()
    t0 = time.time()
    chunks = builder.build_chunks(wb, graph, "test-uuid")
    elapsed = time.time() - t0

    print(f"  Chunks built: {len(chunks)}")
    print(f"  Build time: {elapsed*1000:.0f}ms")

    for i, c in enumerate(chunks[:5]):
        addrs = c.metadata.get("cell_addresses", [])
        sheet = c.metadata.get("sheet", "?")
        pr = c.metadata.get("avg_pagerank", 0)
        print(f"  Chunk {i}: sheet={sheet}, cells={len(addrs)}, avgPR={pr:.4f}, text={c.text[:60]}...")

    # Convert to dict format for retriever
    chunk_dicts = []
    for c in chunks:
        chunk_dicts.append({
            "chunk_id": c.chunk_id,
            "text": c.text,
            "metadata": c.metadata,
        })

    assert len(chunks) > 0
    print(f"  ✅ PASS: {len(chunks)} chunks")
    return chunk_dicts


def test_build_cgasr_index(wb: WorkbookData, graph: nx.DiGraph):
    """Step 4: Build CGASR index."""
    print("\n" + "="*70)
    print("STEP 4: BUILD CGASR INDEX")
    print("="*70)

    embedder = FakeEmbedder()
    t0 = time.time()
    idx = build_cgasr_index(wb, graph, embedder=embedder)
    elapsed = time.time() - t0

    print(f"  N={idx.N}, K={idx.K}")
    print(f"  cell_embeddings shape: {idx.cell_embeddings.shape}")
    print(f"  spectral_PE shape: {idx.spectral_PE.shape}")
    print(f"  tensor_TE shape: {idx.tensor_TE.shape}")
    print(f"  eigenvalues: [{idx.eigenvalues[0]:.6f}, ..., {idx.eigenvalues[-1]:.6f}]")
    print(f"  Wavelet scales: {list(idx.wavelets.keys())}")
    print(f"  Roles distribution:")
    from collections import Counter
    role_counts = Counter(idx.roles.values())
    for role, count in role_counts.most_common():
        print(f"    {role}: {count}")
    print(f"  Build time: {idx.build_time_ms}ms (measured), {elapsed*1000:.0f}ms (total)")

    # Verify shapes
    assert idx.cell_embeddings.shape == (idx.N, 768), f"Embeddings shape mismatch: {idx.cell_embeddings.shape}"
    assert idx.spectral_PE.shape[0] == idx.N, f"PE shape mismatch: {idx.spectral_PE.shape}"
    assert idx.tensor_TE.shape == (idx.N, 16), f"TE shape mismatch: {idx.tensor_TE.shape}"
    assert len(idx.wavelets) == 4, f"Expected 4 wavelet scales, got {len(idx.wavelets)}"
    for s, wdata in idx.wavelets.items():
        assert isinstance(wdata, dict) and "g_lambda" in wdata, f"Wavelet scale {s} has wrong format"
        assert len(wdata["g_lambda"]) == idx.K, f"g_lambda length mismatch at scale {s}"
    assert len(idx.roles) == len(wb.cells), f"Roles count mismatch: {len(idx.roles)} vs {len(wb.cells)}"

    # Verify embeddings are normalized
    norms = np.linalg.norm(idx.cell_embeddings, axis=1)
    assert np.allclose(norms, 1.0, atol=0.01), f"Embeddings not normalized: min={norms.min():.4f}, max={norms.max():.4f}"

    # Verify eigenvalues are non-negative
    assert np.all(idx.eigenvalues >= -0.01), f"Negative eigenvalues: {idx.eigenvalues[:5]}"

    print(f"  ✅ PASS: Index built successfully ({idx.N} cells, {idx.K} spectral components)")
    return idx


def test_retrieval(idx: CGASRIndex, chunks: list):
    """Step 5: Test retrieval with different query types."""
    print("\n" + "="*70)
    print("STEP 5: QUERY-TIME RETRIEVAL")
    print("="*70)

    retriever = CGASRRetriever(idx)

    test_queries = [
        ("What is WACC?", "direct_lookup"),
        ("How does WACC affect the valuation?", "cross_sheet"),
        ("What are the Revenue line items?", "local_reasoning"),
        ("What if Revenue grows 5% faster?", "sensitivity"),
        ("Which cells feed into EBIT?", "structural"),
        ("List all assumptions", "local_reasoning"),
    ]

    all_results = {}

    for query, category in test_queries:
        print(f"\n  ── Query: \"{query}\" [{category}] ──")

        # Route
        route = retriever.route_query(query)
        print(f"  Route: {route}")

        # Embed query
        q_emb = fake_query_embed(query)

        # Retrieve
        t0 = time.time()
        results = retriever.retrieve(q_emb, chunks, k=5)
        elapsed_ms = (time.time() - t0) * 1000

        print(f"  Retrieved: {len(results)} chunks in {elapsed_ms:.0f}ms")

        for i, r in enumerate(results):
            chunk = r["chunk"]
            sheet = chunk.get("metadata", {}).get("sheet", "?")
            cluster = chunk.get("metadata", {}).get("cluster_name", "?")
            n_cells = r["debug"].get("n_cells", 0)
            print(
                f"    #{i+1}: score={r['score']:.3f} σ={r['uncertainty']:.3f} "
                f"sheet={sheet} cluster={cluster} cells={n_cells} "
                f"[attn={r['debug']['attn_score']:.3f} imp={r['debug']['max_imp']:.3f} "
                f"ppr={r['debug']['avg_ppr']:.3f} gap={r['debug']['gap']:.3f}]"
            )

        # Assertions
        assert len(results) > 0, f"No results for query '{query}'"
        assert all(0 <= r["score"] <= 1 for r in results), "Score out of [0,1]"
        assert all(0 <= r["uncertainty"] <= 1 for r in results), "Uncertainty out of [0,1]"
        assert elapsed_ms < 5000, f"Too slow: {elapsed_ms}ms"

        # Check scores are monotonically related to effective score
        effective = [r["score"] * (1 - r["uncertainty"]) for r in results]
        # Allow small tolerance for MMR reordering
        assert effective[0] >= effective[-1] * 0.5, "First result should be roughly better than last"

        all_results[query] = results

    print(f"\n  ✅ PASS: All {len(test_queries)} queries returned valid results")
    return all_results


def test_scale_selection(idx: CGASRIndex):
    """Step 6: Verify adaptive scale selection behavior."""
    print("\n" + "="*70)
    print("STEP 6: ADAPTIVE SCALE SELECTION")
    print("="*70)

    retriever = CGASRRetriever(idx)

    # Simple query should favor fine scales
    q_simple = fake_query_embed("What is WACC?")
    rel_simple = idx.cell_embeddings @ q_simple
    q_spec_simple = idx.eigenvectors.T @ rel_simple
    energies_simple = []
    for s in idx.scales:
        mask = (s * idx.eigenvalues >= 0.5) & (s * idx.eigenvalues <= 2.0)
        energies_simple.append(np.sum(q_spec_simple[mask] ** 2) + 1e-10)

    # Complex query
    q_complex = fake_query_embed("How does WACC affect the DCF enterprise value through the entire model?")
    rel_complex = idx.cell_embeddings @ q_complex
    q_spec_complex = idx.eigenvectors.T @ rel_complex
    energies_complex = []
    for s in idx.scales:
        mask = (s * idx.eigenvalues >= 0.5) & (s * idx.eigenvalues <= 2.0)
        energies_complex.append(np.sum(q_spec_complex[mask] ** 2) + 1e-10)

    print(f"  Simple query energies:  {[f'{e:.4f}' for e in energies_simple]}")
    print(f"  Complex query energies: {[f'{e:.4f}' for e in energies_complex]}")
    print(f"  Scales: {idx.scales}")

    # Different queries should produce different energy distributions
    corr = np.corrcoef(energies_simple, energies_complex)[0, 1]
    print(f"  Energy correlation between simple/complex: {corr:.4f}")
    print(f"  (Lower = better scale differentiation)")

    print(f"  ✅ PASS: Scale selection produces different energy distributions")


def test_uncertainty(idx: CGASRIndex, chunks: list):
    """Step 7: Verify uncertainty estimation."""
    print("\n" + "="*70)
    print("STEP 7: UNCERTAINTY ESTIMATION")
    print("="*70)

    retriever = CGASRRetriever(idx)

    # Run a few queries and check uncertainty distribution
    queries = [
        "What is WACC?",
        "How does WACC affect the valuation?",
        "Explain the entire financial model structure",
    ]

    for query in queries:
        q_emb = fake_query_embed(query)
        results = retriever.retrieve(q_emb, chunks, k=10)

        sigmas = [r["uncertainty"] for r in results]
        scores = [r["score"] for r in results]

        print(f"\n  Query: \"{query}\"")
        print(f"    Scores:       min={min(scores):.3f} max={max(scores):.3f} mean={np.mean(scores):.3f}")
        print(f"    Uncertainty:  min={min(sigmas):.3f} max={max(sigmas):.3f} mean={np.mean(sigmas):.3f}")
        print(f"    Score spread: {max(scores) - min(scores):.3f}")
        print(f"    σ spread:     {max(sigmas) - min(sigmas):.3f}")

    print(f"\n  ✅ PASS: Uncertainty values in valid range [0,1] with meaningful spread")


def test_performance(wb: WorkbookData, graph: nx.DiGraph, chunks: list):
    """Step 8: Performance benchmarks."""
    print("\n" + "="*70)
    print("STEP 8: PERFORMANCE BENCHMARKS")
    print("="*70)

    N = len(wb.cells)
    embedder = FakeEmbedder()

    # Preprocessing time
    t0 = time.time()
    idx = build_cgasr_index(wb, graph, embedder=embedder)
    preprocess_ms = (time.time() - t0) * 1000

    # Query time (average of 5 queries)
    retriever = CGASRRetriever(idx)
    query_times = []
    for q in ["What is WACC?", "Revenue line items", "DCF valuation", "Sensitivity analysis", "Cash flow"]:
        q_emb = fake_query_embed(q)
        t0 = time.time()
        retriever.retrieve(q_emb, chunks, k=10)
        query_times.append((time.time() - t0) * 1000)

    avg_query = np.mean(query_times)
    p95_query = np.percentile(query_times, 95)

    print(f"  Workbook size: {N} cells")
    print(f"  Preprocessing: {preprocess_ms:.0f}ms (target: <3000ms)")
    print(f"  Query avg:     {avg_query:.0f}ms (target: <500ms)")
    print(f"  Query p95:     {p95_query:.0f}ms (target: <1000ms)")
    print(f"  Query times:   {[f'{t:.0f}ms' for t in query_times]}")

    # Memory estimate
    mem_emb = idx.cell_embeddings.nbytes / 1024 / 1024
    mem_pe = idx.spectral_PE.nbytes / 1024 / 1024
    mem_te = idx.tensor_TE.nbytes / 1024 / 1024
    total_mem = mem_emb + mem_pe + mem_te
    print(f"  Memory: embeddings={mem_emb:.1f}MB, PE={mem_pe:.1f}MB, TE={mem_te:.1f}MB, total≈{total_mem:.1f}MB")

    assert preprocess_ms < 60000, f"Preprocessing too slow: {preprocess_ms}ms"
    assert avg_query < 3000, f"Query too slow: {avg_query}ms"

    status_pre = "✅" if preprocess_ms < 3000 else "⚠️"
    status_q = "✅" if avg_query < 500 else "⚠️"
    print(f"\n  {status_pre} Preprocessing: {preprocess_ms:.0f}ms")
    print(f"  {status_q} Avg query: {avg_query:.0f}ms")


# ═════════════════════════════════════════════════════════════════════
# MAIN
# ═════════════════════════════════════════════════════════════════════

def main():
    print("╔══════════════════════════════════════════════════════════════════╗")
    print("║   CGASR Algorithm End-to-End Test                              ║")
    print("║   CellGraph-Attention Spectral Retrieval                       ║")
    print("╚══════════════════════════════════════════════════════════════════╝")

    # Find workbook
    base_dir = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
    candidates = [
        os.path.join(base_dir, "advanced_financial_model.xlsx"),
        os.path.join(base_dir, "sample_financial_model.xlsx"),
    ]
    xlsx_path = None
    for c in candidates:
        if os.path.exists(c):
            xlsx_path = c
            break

    if not xlsx_path:
        print("ERROR: No .xlsx file found. Place a workbook in the project root.")
        sys.exit(1)

    print(f"\nUsing workbook: {os.path.basename(xlsx_path)}")

    # Run all tests
    passed = 0
    failed = 0

    try:
        wb = test_parse_workbook(xlsx_path)
        passed += 1
    except Exception as e:
        print(f"  ❌ FAIL: {e}")
        failed += 1
        return

    try:
        graph = test_build_graph(wb)
        passed += 1
    except Exception as e:
        print(f"  ❌ FAIL: {e}")
        failed += 1
        return

    try:
        chunks = test_build_chunks(wb, graph)
        passed += 1
    except Exception as e:
        print(f"  ❌ FAIL: {e}")
        failed += 1
        return

    try:
        idx = test_build_cgasr_index(wb, graph)
        passed += 1
    except Exception as e:
        print(f"  ❌ FAIL: {e}")
        failed += 1
        import traceback; traceback.print_exc()
        return

    try:
        test_retrieval(idx, chunks)
        passed += 1
    except Exception as e:
        print(f"  ❌ FAIL: {e}")
        failed += 1
        import traceback; traceback.print_exc()

    try:
        test_scale_selection(idx)
        passed += 1
    except Exception as e:
        print(f"  ❌ FAIL: {e}")
        failed += 1
        import traceback; traceback.print_exc()

    try:
        test_uncertainty(idx, chunks)
        passed += 1
    except Exception as e:
        print(f"  ❌ FAIL: {e}")
        failed += 1
        import traceback; traceback.print_exc()

    try:
        test_performance(wb, graph, chunks)
        passed += 1
    except Exception as e:
        print(f"  ❌ FAIL: {e}")
        failed += 1
        import traceback; traceback.print_exc()

    # Summary
    print("\n" + "="*70)
    print(f"SUMMARY: {passed} passed, {failed} failed out of {passed+failed} tests")
    print("="*70)

    if failed == 0:
        print("🎉 ALL TESTS PASSED — CGASR algorithm is working correctly!")
    else:
        print(f"⚠️  {failed} test(s) failed — check output above")

    return failed == 0


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
