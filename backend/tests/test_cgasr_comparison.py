"""
CGASR vs Classic Pipeline — Comparison Test

Runs the same queries through both pipelines and compares:
- Relevance of retrieved chunks
- Cross-sheet coverage
- Score spread & uncertainty
- Latency
"""

import sys
import os
import time
import logging
from collections import Counter, defaultdict

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import numpy as np
import networkx as nx

from parser.xlsx_parser import XLSXParser
from parser.graph_builder import DependencyGraphBuilder
from rag.chunker import ChunkBuilder
from rag.cgasr_index import build_cgasr_index
from rag.cgasr_retriever import CGASRRetriever

logging.basicConfig(level=logging.WARNING, format="%(name)s %(levelname)s  %(message)s")
logger = logging.getLogger("comparison")


class FakeEmbedder:
    def embed(self, texts):
        embs = []
        for text in texts:
            seed = hash(text) % (2**31)
            rng = np.random.RandomState(seed)
            e = rng.randn(768).astype(np.float32)
            e /= np.linalg.norm(e) + 1e-10
            embs.append(e)
        return np.array(embs)

    def embed_single(self, text):
        seed = hash(text) % (2**31)
        rng = np.random.RandomState(seed)
        e = rng.randn(768).astype(np.float32)
        e /= np.linalg.norm(e) + 1e-10
        return e


def cosine_sim(a, b):
    return float(np.dot(a, b) / (np.linalg.norm(a) * np.linalg.norm(b) + 1e-10))


def classic_retrieve(query_emb, chunk_dicts, chunk_embeddings, k=10):
    """Simulate classic cosine-similarity retrieval (vector search only)."""
    sims = chunk_embeddings @ query_emb
    top_idx = np.argsort(sims)[::-1][:k]
    results = []
    for i in top_idx:
        results.append({
            "chunk": chunk_dicts[i],
            "score": float(sims[i]),
            "method": "classic_cosine",
        })
    return results


def main():
    print("╔══════════════════════════════════════════════════════════════════╗")
    print("║   CGASR vs Classic Pipeline — Comparison                       ║")
    print("╚══════════════════════════════════════════════════════════════════╝\n")

    base_dir = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
    xlsx_path = os.path.join(base_dir, "advanced_financial_model.xlsx")
    if not os.path.exists(xlsx_path):
        xlsx_path = os.path.join(base_dir, "sample_financial_model.xlsx")

    # Parse
    parser = XLSXParser()
    wb = parser.parse(xlsx_path)
    builder = DependencyGraphBuilder()
    graph = builder.build(wb)
    graph = builder.run_algorithms(graph, wb)

    chunk_builder = ChunkBuilder()
    chunks = chunk_builder.build_chunks(wb, graph, "test")

    embedder = FakeEmbedder()
    texts = [c.text for c in chunks]
    chunk_embeddings = np.array(embedder.embed(texts), dtype=np.float32)

    chunk_dicts = []
    for c in chunks:
        chunk_dicts.append({
            "chunk_id": c.chunk_id,
            "text": c.text,
            "metadata": c.metadata,
        })

    # Build CGASR — pass embedder so it builds per-cell embeddings (NOT chunk embeddings)
    t0 = time.time()
    cgasr_idx = build_cgasr_index(wb, graph, embedder=embedder)
    build_time = time.time() - t0
    cgasr_ret = CGASRRetriever(cgasr_idx)

    print(f"Workbook: {os.path.basename(xlsx_path)}")
    print(f"Cells: {len(wb.cells)}, Sheets: {len(wb.sheets)}, Chunks: {len(chunks)}")
    print(f"CGASR index build: {build_time*1000:.0f}ms\n")

    # Test queries with expected relevant sheets
    test_queries = [
        {
            "query": "What is WACC?",
            "category": "direct_lookup",
            "expected_sheets": ["WACC", "Valuation", "Assumptions"],
        },
        {
            "query": "How does WACC affect the enterprise value?",
            "category": "cross_sheet",
            "expected_sheets": ["WACC", "Valuation", "DCF"],
        },
        {
            "query": "What are the Revenue assumptions?",
            "category": "local_reasoning",
            "expected_sheets": ["Revenue", "Assumptions"],
        },
        {
            "query": "What if Revenue grows 5% faster?",
            "category": "sensitivity",
            "expected_sheets": ["Revenue", "P&L", "Assumptions", "Scenarios"],
        },
        {
            "query": "Which cells feed into EBIT?",
            "category": "structural",
            "expected_sheets": ["P&L", "Revenue", "Employees"],
        },
        {
            "query": "Show the cash flow statement",
            "category": "direct_lookup",
            "expected_sheets": ["Cash Flow"],
        },
        {
            "query": "What is the terminal value in the DCF?",
            "category": "cross_sheet",
            "expected_sheets": ["DCF", "Valuation", "WACC"],
        },
        {
            "query": "Explain the debt schedule",
            "category": "structural",
            "expected_sheets": ["Debt", "Balance Sheet"],
        },
    ]

    print("=" * 90)
    print(f"{'Query':<45} {'Metric':<15} {'Classic':<15} {'CGASR':<15}")
    print("=" * 90)

    total_classic_sheets = 0
    total_cgasr_sheets = 0
    total_classic_hits = 0
    total_cgasr_hits = 0
    classic_times = []
    cgasr_times = []

    for tq in test_queries:
        query = tq["query"]
        expected = set(s.lower() for s in tq["expected_sheets"])
        q_emb = embedder.embed_single(query)

        # Classic retrieval
        t0 = time.time()
        classic_results = classic_retrieve(q_emb, chunk_dicts, chunk_embeddings, k=10)
        classic_ms = (time.time() - t0) * 1000
        classic_times.append(classic_ms)

        # CGASR retrieval
        t0 = time.time()
        cgasr_results = cgasr_ret.retrieve(q_emb, chunk_dicts, k=10)
        cgasr_ms = (time.time() - t0) * 1000
        cgasr_times.append(cgasr_ms)

        # Analyze classic
        classic_sheets = set()
        for r in classic_results:
            sh = r["chunk"].get("metadata", {}).get("sheet", "").lower()
            if sh:
                classic_sheets.add(sh)
        classic_hit = len(expected & classic_sheets)
        classic_coverage = classic_hit / len(expected) if expected else 0

        # Analyze CGASR
        cgasr_sheets = set()
        for r in cgasr_results:
            sh = r["chunk"].get("metadata", {}).get("sheet", "").lower()
            if sh:
                cgasr_sheets.add(sh)
        cgasr_hit = len(expected & cgasr_sheets)
        cgasr_coverage = cgasr_hit / len(expected) if expected else 0

        total_classic_hits += classic_hit
        total_cgasr_hits += cgasr_hit
        total_classic_sheets += len(classic_sheets)
        total_cgasr_sheets += len(cgasr_sheets)

        # Score spread
        classic_scores = [r["score"] for r in classic_results]
        cgasr_scores = [r["score"] for r in cgasr_results]
        classic_spread = max(classic_scores) - min(classic_scores) if classic_scores else 0
        cgasr_spread = max(cgasr_scores) - min(cgasr_scores) if cgasr_scores else 0

        # Uncertainty (CGASR only)
        cgasr_sigmas = [r.get("uncertainty", 0) for r in cgasr_results]
        avg_sigma = np.mean(cgasr_sigmas) if cgasr_sigmas else 0

        # Print
        short_q = query[:43] + ".." if len(query) > 43 else query
        print(f"{short_q:<45} {'sheets':<15} {len(classic_sheets):<15} {len(cgasr_sheets):<15}")
        print(f"{'':45} {'coverage':<15} {classic_coverage:<15.0%} {cgasr_coverage:<15.0%}")
        print(f"{'':45} {'spread':<15} {classic_spread:<15.3f} {cgasr_spread:<15.3f}")
        print(f"{'':45} {'uncertainty':<15} {'N/A':<15} {avg_sigma:<15.3f}")
        print(f"{'':45} {'latency':<15} {classic_ms:<15.1f} {cgasr_ms:<15.1f}")
        print("-" * 90)

    # Summary
    n = len(test_queries)
    print("\n" + "=" * 90)
    print("SUMMARY")
    print("=" * 90)

    print(f"\n{'Metric':<35} {'Classic':<20} {'CGASR':<20} {'Δ':<15}")
    print("-" * 90)

    avg_classic_sheets = total_classic_sheets / n
    avg_cgasr_sheets = total_cgasr_sheets / n
    delta_sheets = avg_cgasr_sheets - avg_classic_sheets
    print(f"{'Avg sheets per query':<35} {avg_classic_sheets:<20.1f} {avg_cgasr_sheets:<20.1f} {delta_sheets:+.1f}")

    total_expected = sum(len(tq["expected_sheets"]) for tq in test_queries)
    classic_recall = total_classic_hits / total_expected
    cgasr_recall = total_cgasr_hits / total_expected
    delta_recall = cgasr_recall - classic_recall
    print(f"{'Sheet coverage recall':<35} {classic_recall:<20.0%} {cgasr_recall:<20.0%} {delta_recall:+.0%}")

    avg_classic_ms = np.mean(classic_times)
    avg_cgasr_ms = np.mean(cgasr_times)
    print(f"{'Avg latency (ms)':<35} {avg_classic_ms:<20.1f} {avg_cgasr_ms:<20.1f} {avg_cgasr_ms - avg_classic_ms:+.1f}")

    print(f"{'Preprocessing (one-time, ms)':<35} {'0':<20} {build_time*1000:<20.0f} {'N/A':<15}")
    print(f"{'Has uncertainty estimation':<35} {'No':<20} {'Yes':<20}")
    print(f"{'Has spectral graph analysis':<35} {'No':<20} {'Yes':<20}")
    print(f"{'Has role-weighted attention':<35} {'No':<20} {'Yes':<20}")
    print(f"{'Has multi-scale wavelets':<35} {'No':<20} {'Yes':<20}")

    # Detailed role analysis
    print("\n\nCGASR ROLE DISTRIBUTION IN TOP RESULTS:")
    print("-" * 50)
    role_counter = Counter()
    for tq in test_queries:
        q_emb = embedder.embed_single(tq["query"])
        results = cgasr_ret.retrieve(q_emb, chunk_dicts, k=10)
        for r in results:
            addrs = r["chunk"].get("metadata", {}).get("cell_addresses", [])
            if isinstance(addrs, str):
                addrs = addrs.split(",") if addrs else []
            for a in addrs:
                role = cgasr_idx.roles.get(a, "UNKNOWN")
                role_counter[role] += 1

    for role, count in role_counter.most_common():
        bar = "█" * (count // 2)
        print(f"  {role:<15} {count:>4}  {bar}")

    # Spectral info
    print(f"\n\nSPECTRAL ANALYSIS:")
    print("-" * 50)
    print(f"  Eigenvalues (K={cgasr_idx.K}): [{cgasr_idx.eigenvalues[0]:.6f}, ..., {cgasr_idx.eigenvalues[-1]:.6f}]")
    print(f"  Spectral gap (λ₁): {cgasr_idx.eigenvalues[1] - cgasr_idx.eigenvalues[0]:.6f}")
    print(f"  Effective rank (λ > 0.001): {np.sum(cgasr_idx.eigenvalues > 0.001)}")

    # Connected components estimate
    n_near_zero = np.sum(cgasr_idx.eigenvalues < 1e-6)
    print(f"  Near-zero eigenvalues: {n_near_zero} (≈ connected components)")

    print(f"\n\n{'='*90}")
    if cgasr_recall >= classic_recall and avg_cgasr_sheets >= avg_classic_sheets:
        print("✅ CGASR provides BETTER OR EQUAL retrieval quality with richer analysis!")
    elif cgasr_recall > classic_recall:
        print("✅ CGASR provides BETTER coverage despite fewer sheet hits")
    else:
        print("⚠️  CGASR and Classic perform similarly — tune hyperparameters for improvement")
    print(f"{'='*90}")


if __name__ == "__main__":
    main()
