# CellGraph-Attention Spectral Retrieval (CGASR) — Part 2: Implementation

---

## 5. PSEUDOCODE

### 5.1 Preprocessing (`cgasr_preprocess`)

```python
import numpy as np
import scipy.sparse
import scipy.sparse.linalg
from collections import defaultdict

def cgasr_preprocess(workbook_data, graph_f, embedder):
    """Runs once at upload. Returns CGASRIndex. Target: <2s for 5000 cells."""
    cells = list(workbook_data.cells.values())
    cell_ids = list(workbook_data.cells.keys())
    N = len(cells)
    cell_to_idx = {addr: i for i, addr in enumerate(cell_ids)}

    # ── 1. Spatial adjacency via spatial hashing (O(N)) ───────────────
    spatial_hash = defaultdict(list)
    for i, ci in enumerate(cells):
        spatial_hash[(ci.sheet_name, ci.row, ci.col)].append(i)

    rows_s, cols_s, vals_s = [], [], []
    for i, ci in enumerate(cells):
        for dr in [-1, 0, 1]:
            for dc in [-1, 0, 1]:
                if dr == 0 and dc == 0:
                    continue
                key = (ci.sheet_name, ci.row + dr, ci.col + dc)
                for j in spatial_hash.get(key, []):
                    w = 0.5 if (abs(dr) == 1 and abs(dc) == 1) else 1.0
                    rows_s.append(i); cols_s.append(j); vals_s.append(w)
        # Cross-sheet same position
        for sheet in workbook_data.sheets:
            if sheet != ci.sheet_name:
                for j in spatial_hash.get((sheet, ci.row, ci.col), []):
                    rows_s.append(i); cols_s.append(j); vals_s.append(0.3)

    A_s = scipy.sparse.csr_matrix(
        (vals_s, (rows_s, cols_s)), shape=(N, N)
    )

    # ── 2. Formula adjacency with semantic weights ────────────────────
    rows_f, cols_f, vals_f = [], [], []
    for src, dst in graph_f.edges():
        si, di = cell_to_idx.get(src), cell_to_idx.get(dst)
        if si is not None and di is not None:
            w = _formula_weight(graph_f, src, dst)
            rows_f.append(si); cols_f.append(di); vals_f.append(w)
            rows_f.append(di); cols_f.append(si); vals_f.append(w)

    A_f = scipy.sparse.csr_matrix(
        (vals_f, (rows_f, cols_f)), shape=(N, N)
    )

    # ── 3. Combined Laplacian ─────────────────────────────────────────
    alpha = 0.7
    L_f = scipy.sparse.csgraph.laplacian(A_f, normed=True)
    L_s = scipy.sparse.csgraph.laplacian(A_s, normed=True)
    L = alpha * L_f + (1 - alpha) * L_s

    # ── 4. Spectral decomposition ─────────────────────────────────────
    K = min(64, max(N // 10, 4))
    try:
        eigenvalues, eigenvectors = scipy.sparse.linalg.eigsh(
            L, k=K, which='SM', sigma=1e-6
        )
        sort_idx = np.argsort(eigenvalues)
        eigenvalues = eigenvalues[sort_idx]
        eigenvectors = eigenvectors[:, sort_idx]
    except Exception:
        eigenvalues = np.zeros(K)
        eigenvectors = np.random.randn(N, K) * 0.01

    PE = eigenvectors  # (N, K)

    # ── 5. 4D Tensor + CP decomposition ───────────────────────────────
    R_max = min(max((c.row for c in cells), default=1), 100)
    C_max = min(max((c.col for c in cells), default=1), 26)
    S = len(workbook_data.sheets)
    sheet_idx = {s: i for i, s in enumerate(workbook_data.sheets)}
    T = 4  # numeric=0, formula=1, text=2, empty=3

    tensor = np.zeros((R_max, C_max, S, T), dtype=np.float32)
    for ci in cells:
        r, c, s = ci.row - 1, ci.col - 1, sheet_idx.get(ci.sheet_name, 0)
        if r < R_max and c < C_max:
            t = _cell_type_id(ci)
            tensor[r, c, s, t] = _normalize_value(ci)

    try:
        import tensorly
        from tensorly.decomposition import parafac
        factors = parafac(tensorly.tensor(tensor), rank=16, n_iter_max=50)
        fa, fb, fc, fd = [np.array(f) for f in factors.factors]
    except Exception:
        fa = np.random.randn(R_max, 16) * 0.01
        fb = np.random.randn(C_max, 16) * 0.01
        fc = np.random.randn(S, 16) * 0.01
        fd = np.random.randn(T, 16) * 0.01

    TE = np.zeros((N, 16), dtype=np.float32)
    for i, ci in enumerate(cells):
        r, c, s = ci.row - 1, ci.col - 1, sheet_idx.get(ci.sheet_name, 0)
        t = _cell_type_id(ci)
        if r < R_max and c < C_max:
            TE[i] = fa[r] * fb[c] * fc[s] * fd[t]

    # ── 6. Fused cell embeddings ──────────────────────────────────────
    cell_texts = [_cell_to_text(ci) for ci in cells]
    SE = np.array(embedder.embed(cell_texts), dtype=np.float32)  # (N, 768)

    combined = np.hstack([PE, TE, SE])  # (N, K+16+768)
    rng = np.random.RandomState(42)
    W_fuse = rng.randn(combined.shape[1], 768).astype(np.float32)
    W_fuse /= np.sqrt(combined.shape[1])
    E = combined @ W_fuse  # (N, 768)
    norms = np.linalg.norm(E, axis=1, keepdims=True) + 1e-10
    E = E / norms

    # ── 7. Wavelet coefficients (4 scales) ────────────────────────────
    scales = [0.5, 2.0, 8.0, 32.0]
    wavelets = {}
    for s in scales:
        g_lambda = s * eigenvalues * np.exp(-s * eigenvalues)
        W_mat = eigenvectors @ np.diag(g_lambda) @ eigenvectors.T
        W_mat[np.abs(W_mat) < 1e-4] = 0
        wavelets[s] = scipy.sparse.csr_matrix(W_mat)

    # ── 8. Cell roles ─────────────────────────────────────────────────
    roles = {}
    for addr, ci in workbook_data.cells.items():
        roles[addr] = _classify_role(ci, graph_f)

    # ── 9. Transition matrix for PPR ──────────────────────────────────
    col_sums = np.array(A_f.sum(axis=0)).flatten() + 1e-10
    P = A_f.multiply(1.0 / col_sums)

    return CGASRIndex(
        cell_embeddings=E, spectral_PE=PE, tensor_TE=TE,
        eigenvalues=eigenvalues, eigenvectors=eigenvectors,
        wavelets=wavelets, roles=roles, cell_to_idx=cell_to_idx,
        L=L, P=P, scales=scales, cell_ids=cell_ids,
    )


# ── Helper functions ──────────────────────────────────────────────────

def _formula_weight(graph, src, dst):
    formula = (graph.nodes[dst].get("formula") or "").upper()
    if any(f in formula for f in ["SUM(", "SUMIF(", "SUMPRODUCT("]):
        beta = 1.5
    elif any(f in formula for f in ["AVERAGE(", "AVG("]):
        beta = 1.3
    elif any(f in formula for f in ["IF(", "IFERROR("]):
        beta = 0.8
    elif any(f in formula for f in ["VLOOKUP(", "INDEX(", "MATCH("]):
        beta = 0.6
    else:
        beta = 1.0
    fan_in = max(graph.in_degree(dst), 1)
    return beta * (fan_in ** -0.5)


def _cell_type_id(cell):
    if cell.formula:
        return 1
    if isinstance(cell.value, (int, float)):
        return 0
    if isinstance(cell.value, str) and cell.value.strip():
        return 2
    return 3


def _normalize_value(cell):
    if isinstance(cell.value, (int, float)):
        return np.tanh(cell.value / 1e6)  # soft normalization
    return 0.5 if cell.formula else (0.1 if isinstance(cell.value, str) else 0.0)


def _cell_to_text(cell):
    parts = [f"{cell.sheet_name}!{cell.cell_address}"]
    if cell.named_range:
        parts.append(cell.named_range)
    if cell.formula:
        parts.append(f"={cell.formula}")
    if cell.value is not None:
        parts.append(str(cell.value)[:50])
    return " ".join(parts)


ROLE_WEIGHTS = {
    "TOTAL": 1.5, "SUBTOTAL": 1.3, "CROSS_REF": 1.2,
    "ASSUMPTION": 1.4, "LINE_ITEM": 1.0, "HEADER": 0.8, "COMPUTED": 1.0
}

def _classify_role(cell, graph):
    formula = (cell.formula or "").upper()
    addr = f"{cell.sheet_name}!{cell.cell_address}" if hasattr(cell, 'cell_address') else ""
    if cell.sheet_name == "Assumptions" and cell.is_hardcoded:
        return "ASSUMPTION"
    if isinstance(cell.value, str) and not cell.formula:
        return "HEADER"
    if "!" in (cell.formula or ""):
        return "CROSS_REF"
    if "SUM" in formula:
        if addr in graph:
            for pred in graph.predecessors(addr):
                if "SUM" in (graph.nodes[pred].get("formula") or "").upper():
                    return "TOTAL"
        return "SUBTOTAL"
    if isinstance(cell.value, (int, float)):
        return "LINE_ITEM"
    return "COMPUTED"
```

### 5.2 Query-Time Retrieval (`cgasr_retrieve`)

```python
def cgasr_retrieve(query, idx, chunks, embedder, k=10):
    """Per-query retrieval. Target: <500ms for 5000 cells."""
    N = idx.cell_embeddings.shape[0]
    q = np.array(embedder.embed_single(query), dtype=np.float32)

    # ── 1. Cell relevances ────────────────────────────────────────────
    rel = idx.cell_embeddings @ q  # (N,) cosine similarities

    # ── 2. Adaptive scale selection ───────────────────────────────────
    q_spectral = idx.eigenvectors.T @ rel  # (K,)
    scale_energies = []
    for s in idx.scales:
        mask = (s * idx.eigenvalues >= 0.5) & (s * idx.eigenvalues <= 2.0)
        scale_energies.append(np.sum(q_spectral[mask] ** 2) + 1e-10)
    w_scales = _softmax(np.array(scale_energies) / 0.1)

    # ── 3. Personalized PageRank (25 iterations) ─────────────────────
    teleport = _softmax(rel / 0.5)
    pi = teleport.copy()
    P_dense = idx.P.toarray() if scipy.sparse.issparse(idx.P) else idx.P
    for _ in range(25):
        pi = 0.85 * P_dense @ pi + 0.15 * teleport
    pi /= (pi.sum() + 1e-10)

    # ── 4. Multi-scale wavelet importance ─────────────────────────────
    importance = np.zeros(N)
    for i, s in enumerate(idx.scales):
        imp_s = idx.wavelets[s] @ pi
        importance += w_scales[i] * imp_s

    # ── 5. Score each chunk ───────────────────────────────────────────
    results = []
    for chunk in chunks:
        addrs = chunk.get("metadata", {}).get("cell_addresses", [])
        if isinstance(addrs, str):
            addrs = addrs.split(",") if addrs else []
        ci = [idx.cell_to_idx[a] for a in addrs if a in idx.cell_to_idx]

        if not ci:
            results.append((chunk, 0.0, 1.0))
            continue

        ci = np.array(ci)

        # Attention-pooled chunk repr
        rw = np.array([ROLE_WEIGHTS.get(idx.roles.get(addrs[j], "COMPUTED"), 1.0)
                       for j in range(len(ci))])
        attn_w = _softmax(importance[ci] * rw)
        h_chunk = attn_w @ idx.cell_embeddings[ci]  # (768,)

        # Cross-attention score
        h_norm = np.linalg.norm(h_chunk) + 1e-10
        attn_score = float(np.dot(q, h_chunk) / h_norm)

        # Max importance + avg PPR
        max_imp = float(np.max(importance[ci]))
        avg_ppr = float(np.mean(pi[ci]))

        # Spectral gap
        if len(ci) >= 3:
            sub = idx.L[np.ix_(ci, ci)]
            sub_dense = sub.toarray() if scipy.sparse.issparse(sub) else sub
            try:
                eigs = np.sort(np.linalg.eigvalsh(sub_dense))
                gap = float(eigs[1] - eigs[0]) if len(eigs) > 1 else 0.0
            except Exception:
                gap = 0.0
        else:
            gap = 0.5

        # Uncertainty
        p = _softmax(rel[ci])
        entropy = -np.sum(p * np.log(p + 1e-10))
        max_ent = np.log(len(ci) + 1e-10)
        sigma = float(entropy / max_ent) if max_ent > 0 else 1.0

        # Final sigmoid score
        raw = 0.4 * attn_score + 0.25 * max_imp + 0.2 * avg_ppr + 0.15 * gap
        score = float(1.0 / (1.0 + np.exp(-5.0 * raw)))  # scaled sigmoid

        results.append((chunk, score, sigma))

    # ── 6. MMR with uncertainty ───────────────────────────────────────
    results.sort(key=lambda x: x[1] * (1 - x[2]), reverse=True)
    selected = []
    for chunk, score, unc in results:
        if len(selected) >= k:
            break
        if not selected:
            selected.append({"chunk": chunk, "score": score, "uncertainty": unc})
            continue
        # Max sim to selected
        addrs = chunk.get("metadata", {}).get("cell_addresses", [])
        if isinstance(addrs, str):
            addrs = addrs.split(",") if addrs else []
        ci = [idx.cell_to_idx[a] for a in addrs if a in idx.cell_to_idx]
        if ci:
            h = np.mean(idx.cell_embeddings[ci], axis=0)
            max_sim = 0.0
            for s in selected:
                sa = s["chunk"].get("metadata", {}).get("cell_addresses", [])
                if isinstance(sa, str):
                    sa = sa.split(",") if sa else []
                si = [idx.cell_to_idx[a] for a in sa if a in idx.cell_to_idx]
                if si:
                    hs = np.mean(idx.cell_embeddings[si], axis=0)
                    sim = float(np.dot(h, hs) / (np.linalg.norm(h) * np.linalg.norm(hs) + 1e-10))
                    max_sim = max(max_sim, sim)
        else:
            max_sim = 0.0
        mmr = 0.6 * score * (1 - unc) - 0.4 * max_sim
        if mmr > -0.1:
            selected.append({"chunk": chunk, "score": score, "uncertainty": unc})

    return selected


def _softmax(x):
    e = np.exp(x - np.max(x))
    return e / (e.sum() + 1e-10)
```

---

## 6. COMPLEXITY ANALYSIS

### 6.1 Preprocessing (once per upload)

| Step | Time | Space | N=5000 estimate |
|------|------|-------|-----------------|
| Spatial graph (hash) | O(N) | O(N) | ~5ms |
| Formula weights | O(\|E\|) | O(\|E\|) | ~2ms |
| Combined Laplacian | O(N+\|E\|) | O(N+\|E\|) | ~3ms |
| **Spectral decomp** | **O(N·K²)** | O(N·K) | **~500ms** |
| 4D Tensor + CP | O(R·C·S·T·16·50) | O(Σ dims·16) | ~200ms |
| Cell embeddings | O(N·768) | O(N·768) | ~700ms (shared) |
| Fused projection | O(N·848·768) | O(N·768) | ~50ms |
| Wavelet matrices | O(4·N·K) | O(4·nnz) | ~100ms |
| Role classification | O(N·deg) | O(N) | ~10ms |
| **TOTAL** | | | **~1.5s** |

### 6.2 Per-Query

| Step | Time | N=5000 estimate |
|------|------|-----------------|
| Query embed | O(d) | ~10ms |
| Cell relevances | O(N·d) | ~5ms |
| Scale selection | O(N·K) | ~2ms |
| **PPR (25 iter)** | **O(25·N²)** or O(25·\|E\|) sparse | **~150ms** |
| Wavelet propagation | O(4·nnz) | ~20ms |
| Chunk scoring | O(M·\|C_avg\|·d) | ~50ms |
| Spectral gaps | O(M·\|C_avg\|³) | ~100ms |
| MMR ranking | O(k²·d) | ~5ms |
| **TOTAL** | | **~350ms** |

### 6.3 Space: ~20 MB for 5000 cells

---

## 7. COMPARISON TABLE

| Weakness | Current System | CGASR Solution | Expected Δ |
|----------|---------------|----------------|------------|
| **W1: Spatial** | Text embeddings lose grid | Dual Laplacian spectral PE | +25% context precision |
| **W2: Static traversal** | 1-hop BFS unweighted | Multi-scale wavelets + semantic edges | +30% cross-sheet recall |
| **W3: Query-agnostic** | Static PageRank | Personalized PPR from query | +20% MRR |
| **W4: Flat chunks** | Equal cells in cluster | Role-weighted attention pooling | +15% nDCG@5 |
| **W5: No formula sem.** | Formula as text string | β_type edge weights + propagation | +10% formula query accuracy |
| **W6: Cross-sheet** | 1-hop misses chains | Coarse wavelet s=32 | +35% cross-sheet recall |
| **W7: No confidence** | All equally certain | Spectral gap + entropy → σ | New capability |

---

## 8. ABLATION STUDY DESIGN

### Components to isolate:

| ID | Ablation | Tests |
|----|----------|-------|
| A1 | Remove spectral PE | Spatial awareness contribution |
| A2 | Remove tensor TE | 4D structure contribution |
| A3 | All edge weights = 1.0 | Formula semantic contribution |
| A4 | Static PageRank (no PPR) | Query conditioning contribution |
| A5 | Single wavelet scale | Multi-resolution contribution |
| A6 | All role weights = 1.0 | Hierarchy contribution |
| A7 | Remove σ from MMR | Uncertainty contribution |
| A8 | Replace cross-attn w/ cosine | Attention mechanism contribution |
| A9 | Full CGASR | Ceiling |
| A10 | Current pipeline | Floor / baseline |

### Metrics:
- **nDCG@5**, **MRR**, **Hit@1** — retrieval quality
- **Context Precision** — fraction of relevant tokens retrieved
- **Cross-Sheet Recall** — multi-sheet chain coverage
- **Answer Accuracy** — end-to-end LLM correctness
- **Uncertainty Calibration** — Spearman(σ, actual_error)
- **Latency p95** — must stay < 1000ms

### Evaluation set: 100 queries across 5 workbooks
- 20 direct lookup, 20 local reasoning, 20 cross-sheet, 15 sensitivity, 15 structural, 10 anomaly

---

## 9. IMPLEMENTATION ROADMAP

### Phase 1: Core Spectral (Week 1-2)
**File**: `backend/rag/cgasr_index.py`
- `CGASRIndex` dataclass
- `build_spatial_graph()` with hash optimization
- `build_combined_laplacian()`
- `spectral_decompose()`
- `classify_cell_roles()`
- **Integration**: called from `workbook.py` upload after `run_algorithms()`

### Phase 2: Tensor + Wavelet + Fusion (Week 2-3)
**File**: `backend/rag/cgasr_index.py` continued
- `build_cell_tensor()` + `cp_decompose()`
- `compute_wavelet_coefficients()`
- `compute_fused_embeddings()`
- `compute_formula_weights()`
- **Dep**: add `tensorly` to `requirements.txt`

### Phase 3: Query-Time Retriever (Week 3-4)
**File**: `backend/rag/cgasr_retriever.py`
- `CGASRRetriever` class
- `select_scale()`, `personalized_pagerank()`, `wavelet_propagate()`
- `score_chunks()`, `mmr_rank()`
- **Integration**: replace `retriever.retrieve()` in `excel_agent.py:1015`

### Phase 4: Eval & Tune (Week 4-5)
**File**: `backend/rag/cgasr_eval.py`
- Benchmark dataset (100 queries, 5 workbooks)
- Metric computation harness
- Ablation runner
- Hyperparameter sweep

---

## 10. REFERENCES & DIFFERENTIATION

| Reference | How CGASR Differs |
|-----------|-------------------|
| **TabRAG** (Si et al., 2025) | TabRAG parses tables from PDFs via VLM; CGASR operates on live spreadsheets with formula graphs — fundamentally different modality. |
| **TableRAG** (Zhang et al., 2025) | Treats tables as relational DB; CGASR treats spreadsheets as computational graphs with formula semantics, cross-sheet deps that SQL cannot capture. |
| **SpreadsheetLLM/SheetCompressor** (Chen et al., 2024) | Compresses for token efficiency in prompts; CGASR optimizes *retrieval* quality with query-conditioned scoring — complementary approaches. |
| **GAT** (Veličković et al., 2018) | Attention on node features within graph. CGASR adds cross-attention with *external* query + multi-scale wavelets + tensor decomposition. |
| **Spectral Graph Wavelets** (Hammond, 2011) | Static multi-resolution filters. CGASR introduces query-conditioned scale mixture — novel extension. |
| **WaveGC** (ICML 2025) | Learnable wavelet bases for GNN classification. CGASR uses fixed kernels with query-adaptive scale selection for retrieval. |
| **Personalized PageRank** (Haveliwala, 2002) | Fixed topic teleportation. CGASR derives teleportation from real-time query embeddings. |
