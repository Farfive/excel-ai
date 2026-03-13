# Cell Entropy Attention (CEA) — Fast Fallback Algorithm

**Companion to CGASR | Target: < 500ms per query | No training required**

---

## 1. NAME & DESCRIPTION

**Cell Entropy Attention (CEA)** — A lightweight algorithm using Shannon entropy of cell distributions and a 2-layer parameter-free Graph Attention Network to compute query-dependent cell importance. Serves as fast fallback when CGASR preprocessing is incomplete or latency budget is tight.

---

## 2. MATHEMATICAL FORMULATION

### 2.1 Column/Row Entropy (Precomputed at Upload)

**(Eq. C1)** Shannon entropy of column c in sheet s:
```
H_col(c, s) = -Σ_{t ∈ types} p_t · log₂(p_t + ε)
```
where p_t = fraction of cells in column (c, s) with type t ∈ {numeric, formula, text, empty, date, percentage}.

**(Eq. C2)** Row entropy:
```
H_row(r, s) = -Σ_{t ∈ types} p_t · log₂(p_t + ε)
```

High entropy = mixed types (header + values + formulas) = **information-dense** boundary regions.

**(Eq. C3)** Cell information density:
```
ID(c_i) = H_col(col_i, sheet_i) + H_row(row_i, sheet_i)
```

### 2.2 Conditional Mutual Information (Per Query)

**(Eq. C4)** Approximate I(c_i; q | N_i) without explicit probability distributions:
```
I(c_i; q | N_i) = max(0, cos(e_i, q) - (1/|N_i|) · Σ_{j ∈ N_i} cos(e_j, q))
```

where e_i is the existing SentenceTransformer embedding, N_i = predecessors ∪ successors in dependency graph.

**Interpretation**: Cell is important if relevant to query AND neighbors DON'T already cover that relevance. This captures unique information contribution.

### 2.3 2-Layer Graph Attention (Parameter-Free)

No learned parameters — attention vectors are fixed random projections. The mechanism still provides adaptive neighbor weighting based on embedding geometry.

**(Eq. C5)** Layer 1 — neighbor aggregation:
```
α_{ij}^{(1)} = softmax_j(LeakyReLU(a_1^T · [e_i ‖ e_j]))
h_i^{(1)} = ReLU(Σ_{j ∈ N_i ∪ {i}} α_{ij}^{(1)} · e_j)
```

where a_1 ∈ ℝ^{2d} is a fixed unit-norm random vector (seed=42).

**(Eq. C6)** Layer 2 — query-conditioned propagation:
```
α_{ij}^{(2)} = softmax_j(LeakyReLU(a_2^T · [h_i^{(1)} ‖ h_j^{(1)} ‖ q]))
h_i^{(2)} = ReLU(Σ_{j ∈ N_i ∪ {i}} α_{ij}^{(2)} · h_j^{(1)})
```

where a_2 ∈ ℝ^{3d} is a fixed unit-norm random vector (seed=43).

**Why this works without training**: The concatenation [e_i ‖ e_j] already encodes semantic similarity via the pre-trained SentenceTransformer. The random projection a_1 creates a consistent (reproducible) linear combination. LeakyReLU + softmax then adaptively weights neighbors by this projected similarity. Layer 2 adds query conditioning via the [... ‖ q] concatenation.

### 2.4 CEA Cell Importance Score

**(Eq. C7)**
```
IMP_CEA(c_i, q) = δ_1 · ID(c_i) + δ_2 · I(c_i; q | N_i) + δ_3 · cos(h_i^{(2)}, q)
```
where δ = [0.2, 0.3, 0.5] — GNN output dominates, entropy provides structural prior, CMI adds uniqueness.

### 2.5 Chunk Scoring

**(Eq. C8)** Max-mean hybrid:
```
score_CEA(C_j, q) = 0.7 · max_{c_i ∈ C_j} IMP_CEA(c_i, q) + 0.3 · mean_{c_i ∈ C_j} IMP_CEA(c_i, q)
```

**(Eq. C9)** Variance-based uncertainty:
```
σ_CEA(C_j, q) = std_{c_i ∈ C_j}[IMP_CEA(c_i, q)] / (range(IMP_CEA) + ε)
```

High within-chunk variance = cells disagree on relevance = uncertain.

---

## 3. CEA PSEUDOCODE

### 3.1 Preprocessing

```python
from collections import defaultdict
import numpy as np

def cea_preprocess(workbook_data):
    """O(N) preprocessing. Returns CEAIndex."""
    cells = workbook_data.cells

    # ── Column and row entropy ────────────────────────────────────────
    col_types = defaultdict(lambda: defaultdict(int))
    row_types = defaultdict(lambda: defaultdict(int))

    for cell in cells.values():
        t = _cell_type_id(cell)  # 0-5: numeric, formula, text, empty, date, pct
        col_types[(cell.sheet_name, cell.col)][t] += 1
        row_types[(cell.sheet_name, cell.row)][t] += 1

    def entropy(counts):
        total = sum(counts.values())
        if total == 0:
            return 0.0
        probs = [c / total for c in counts.values()]
        return -sum(p * np.log2(p + 1e-10) for p in probs)

    col_H = {k: entropy(v) for k, v in col_types.items()}
    row_H = {k: entropy(v) for k, v in row_types.items()}

    info_density = {}
    for addr, cell in cells.items():
        info_density[addr] = (
            col_H.get((cell.sheet_name, cell.col), 0.0) +
            row_H.get((cell.sheet_name, cell.row), 0.0)
        )

    # ── Fixed random attention vectors ────────────────────────────────
    d = 768
    rng = np.random.RandomState(42)
    a1 = rng.randn(2 * d).astype(np.float32)
    a1 /= np.linalg.norm(a1)
    rng2 = np.random.RandomState(43)
    a2 = rng2.randn(3 * d).astype(np.float32)
    a2 /= np.linalg.norm(a2)

    return CEAIndex(info_density=info_density, a1=a1, a2=a2)
```

### 3.2 Query-Time Retrieval

```python
def cea_retrieve(query, cea_idx, cell_embeddings, graph, chunks, embedder, k=10):
    """
    Per-query CEA retrieval. Target: <500ms.
    
    Args:
        cell_embeddings: (N, 768) pre-computed SentenceTransformer embeddings
        graph: NetworkX DiGraph (existing dependency graph)
    """
    cell_ids = list(graph.nodes())
    cell_to_idx = {addr: i for i, addr in enumerate(cell_ids)}
    N = len(cell_ids)
    E = cell_embeddings  # (N, 768)
    d = 768

    q = np.array(embedder.embed_single(query), dtype=np.float32)

    # ── 1. Cosine relevances ──────────────────────────────────────────
    cos_q = E @ q  # (N,)

    # ── 2. Conditional Mutual Information ─────────────────────────────
    cmi = np.zeros(N, dtype=np.float32)
    for i, addr in enumerate(cell_ids):
        nbrs = list(graph.predecessors(addr)) + list(graph.successors(addr))
        nbr_idx = [cell_to_idx[n] for n in nbrs if n in cell_to_idx]
        if nbr_idx:
            avg_nbr_cos = np.mean(cos_q[nbr_idx])
            cmi[i] = max(0.0, cos_q[i] - avg_nbr_cos)
        else:
            cmi[i] = max(0.0, cos_q[i])

    # ── 3. Layer 1: Neighbor attention ────────────────────────────────
    h1 = np.zeros_like(E)  # (N, 768)
    a1 = cea_idx.a1

    for i, addr in enumerate(cell_ids):
        nbrs = list(graph.predecessors(addr)) + list(graph.successors(addr))
        nbr_idx = [cell_to_idx[n] for n in nbrs if n in cell_to_idx]
        all_idx = [i] + nbr_idx  # include self

        if len(all_idx) == 1:
            h1[i] = E[i]
            continue

        # Compute attention coefficients
        concat = np.hstack([
            np.tile(E[i], (len(all_idx), 1)),  # (|N|+1, d)
            E[all_idx]                           # (|N|+1, d)
        ])  # (|N|+1, 2d)
        
        logits = concat @ a1  # (|N|+1,)
        logits = np.where(logits > 0, logits, 0.2 * logits)  # LeakyReLU
        alpha = _softmax(logits)  # (|N|+1,)
        
        h1[i] = np.maximum(0, alpha @ E[all_idx])  # ReLU

    # ── 4. Layer 2: Query-conditioned ─────────────────────────────────
    h2 = np.zeros_like(E)  # (N, 768)
    a2 = cea_idx.a2

    for i, addr in enumerate(cell_ids):
        nbrs = list(graph.predecessors(addr)) + list(graph.successors(addr))
        nbr_idx = [cell_to_idx[n] for n in nbrs if n in cell_to_idx]
        all_idx = [i] + nbr_idx

        if len(all_idx) == 1:
            h2[i] = h1[i]
            continue

        # [h_i || h_j || q]
        concat = np.hstack([
            np.tile(h1[i], (len(all_idx), 1)),  # (|N|+1, d)
            h1[all_idx],                          # (|N|+1, d)
            np.tile(q, (len(all_idx), 1))         # (|N|+1, d)
        ])  # (|N|+1, 3d)

        logits = concat @ a2
        logits = np.where(logits > 0, logits, 0.2 * logits)
        alpha = _softmax(logits)
        
        h2[i] = np.maximum(0, alpha @ h1[all_idx])

    # ── 5. CEA importance ─────────────────────────────────────────────
    gnn_cos = np.array([np.dot(h2[i], q) / (np.linalg.norm(h2[i]) + 1e-10) for i in range(N)])
    
    # Normalize components to [0, 1]
    id_arr = np.array([cea_idx.info_density.get(addr, 0.0) for addr in cell_ids])
    id_norm = id_arr / (id_arr.max() + 1e-10)
    cmi_norm = cmi / (cmi.max() + 1e-10)
    gnn_norm = (gnn_cos - gnn_cos.min()) / (gnn_cos.max() - gnn_cos.min() + 1e-10)

    imp = 0.2 * id_norm + 0.3 * cmi_norm + 0.5 * gnn_norm  # (N,)

    # ── 6. Chunk scoring ──────────────────────────────────────────────
    results = []
    for chunk in chunks:
        addrs = chunk.get("metadata", {}).get("cell_addresses", [])
        if isinstance(addrs, str):
            addrs = addrs.split(",") if addrs else []
        ci = [cell_to_idx[a] for a in addrs if a in cell_to_idx]

        if not ci:
            results.append((chunk, 0.0, 1.0))
            continue

        ci = np.array(ci)
        chunk_imp = imp[ci]

        # Score: max-mean hybrid
        score = 0.7 * np.max(chunk_imp) + 0.3 * np.mean(chunk_imp)

        # Uncertainty: normalized std
        imp_range = imp.max() - imp.min() + 1e-10
        sigma = float(np.std(chunk_imp) / imp_range)

        results.append((chunk, float(score), sigma))

    # ── 7. Rank by score*(1-σ) ────────────────────────────────────────
    results.sort(key=lambda x: x[1] * (1 - x[2]), reverse=True)

    return [
        {"chunk": chunk, "score": score, "uncertainty": sigma}
        for chunk, score, sigma in results[:k]
    ]


def _softmax(x):
    e = np.exp(x - np.max(x))
    return e / (e.sum() + 1e-10)
```

### 3.3 Optimization: Batched GAT for Speed

The per-node loop in Layer 1/2 is the bottleneck. For production, batch using sparse adjacency:

```python
def _batched_gat_layer(E, A_sparse, a_vec, d):
    """
    Vectorized GAT layer using sparse adjacency.
    E: (N, d) node features
    A_sparse: (N, N) sparse adjacency (binary + self-loops)
    a_vec: (2d,) or (3d,) attention vector
    Returns: (N, d) updated features
    """
    N = E.shape[0]
    
    # For each edge (i,j), compute a^T [e_i || e_j]
    # This is: a[:d] @ e_i + a[d:2d] @ e_j for the 2d case
    a_left = a_vec[:d]    # (d,)
    a_right = a_vec[d:2*d]  # (d,)
    
    left_scores = E @ a_left    # (N,) — each node's "query" score
    right_scores = E @ a_right  # (N,) — each node's "key" score
    
    # For each edge (i,j): logit = left_scores[i] + right_scores[j]
    rows, cols = A_sparse.nonzero()
    logits = left_scores[rows] + right_scores[cols]
    logits = np.where(logits > 0, logits, 0.2 * logits)  # LeakyReLU
    
    # Sparse softmax per row
    from scipy.special import softmax as sp_softmax
    # Group by row, softmax within each group
    alpha_sparse = scipy.sparse.csr_matrix((logits, (rows, cols)), shape=(N, N))
    
    # Row-wise softmax on sparse matrix
    for i in range(N):
        start, end = alpha_sparse.indptr[i], alpha_sparse.indptr[i+1]
        if end > start:
            vals = alpha_sparse.data[start:end]
            vals = np.exp(vals - vals.max())
            alpha_sparse.data[start:end] = vals / (vals.sum() + 1e-10)
    
    # Propagate: H = alpha @ E
    H = alpha_sparse @ E  # (N, d) sparse matmul
    return np.maximum(0, H)  # ReLU
```

**With batched GAT**: Layer 1 + Layer 2 run in O(|E| · d) instead of O(N · deg · d), bringing total to ~100ms.

---

## 4. CEA COMPLEXITY

### Preprocessing: O(N)
| Step | Time | Space |
|------|------|-------|
| Column/row entropy | O(N) | O(sheets × max(row,col)) |
| Random vectors | O(d) | O(d) |
| **TOTAL** | O(N) ≈ 5ms | ~10 KB |

### Per-Query: O(|E| · d)
| Step | Time | N=5000, \|E\|=8000 |
|------|------|---------------------|
| Query embed | O(d) | 10ms |
| Cosine relevances | O(N·d) | 5ms |
| CMI computation | O(N·deg) | 20ms |
| **GAT Layer 1 (batched)** | O(\|E\|·d) | **50ms** |
| **GAT Layer 2 (batched)** | O(\|E\|·d) | **50ms** |
| Importance scores | O(N) | 2ms |
| Chunk scoring | O(M·\|C_avg\|) | 10ms |
| Sorting | O(M·log M) | 1ms |
| **TOTAL** | | **~150ms** |

**Well within 500ms budget**, even on CPU.

---

## 5. CEA vs CGASR COMPARISON

| Aspect | CEA | CGASR |
|--------|-----|-------|
| **Preprocessing** | O(N), 5ms | O(N·K²), 1.5s |
| **Per-query** | ~150ms | ~350ms |
| **Spatial awareness** | None (uses existing embeddings) | Full (spectral PE) |
| **Multi-scale** | No (single aggregation) | Yes (4 wavelet scales) |
| **Cross-sheet** | Partial (2-hop via GAT) | Full (coarse wavelets) |
| **Uncertainty** | Variance-based (simple) | Spectral gap + entropy (rigorous) |
| **Dependencies** | numpy only | numpy, scipy, tensorly |
| **Best for** | Simple queries, low latency | Complex reasoning, cross-sheet |

### When to Use Which

```python
def select_retriever(query, workbook_size):
    """Route to CGASR or CEA based on query complexity and constraints."""
    # Simple lookup → neither (use direct lookup)
    if is_direct_lookup(query):
        return "direct"
    
    # Cross-sheet or sensitivity query → CGASR
    if has_cross_sheet_keywords(query) or has_sensitivity_keywords(query):
        return "cgasr"
    
    # Large workbook + simple question → CEA for speed
    if workbook_size > 8000 and not has_reasoning_keywords(query):
        return "cea"
    
    # Default to CGASR
    return "cgasr"
```

---

## 6. CEA THEORETICAL JUSTIFICATION

### Why entropy identifies structural regions:

In financial models, **information-dense** regions are structurally important:
- A header row has mixed types: text labels + empty → H ≈ 1.0 bit
- A pure data column has one type: all numeric → H ≈ 0 bits
- A boundary row (Total) has: text label + formula + sometimes hardcoded → H ≈ 1.5 bits
- Assumption sheets have: text + numeric + named ranges → H ≈ 1.8 bits

High entropy regions correlate with **structural anchors** (exactly what SheetCompressor identifies heuristically). CEA discovers these automatically via information theory.

### Why random-projection GAT works:

The Johnson-Lindenstrauss lemma guarantees that random projections preserve pairwise distances with high probability. For d=768, a random vector a ∈ ℝ^{2d} creates a 1-dimensional projection where similar [e_i ‖ e_j] pairs (semantically related cells) will have similar projected values. The softmax then gives higher attention to semantically similar neighbors.

The key insight: **we don't need learned attention** because the SentenceTransformer embeddings are already semantically meaningful. The random projection + softmax is sufficient to create adaptive neighbor weighting.

### Convergence:

CEA has no iterative computation (no PageRank, no eigendecomposition). The 2-layer GAT is a single forward pass. There are no convergence concerns — the output is deterministic given the fixed random seed.

---

## 7. CEA ABLATION COMPONENTS

| ID | Remove | Expected Impact |
|----|--------|-----------------|
| CE1 | Remove entropy (δ_1=0) | -5% nDCG on structural queries |
| CE2 | Remove CMI (δ_2=0) | -10% nDCG on precision-critical queries |
| CE3 | Remove GAT (δ_3=0, use raw cosine) | -15% overall — GAT is main contributor |
| CE4 | Remove Layer 2 (single-layer GAT) | -8% on cross-sheet queries |
| CE5 | Learned a_1, a_2 (fine-tuned) | +5% potential, but requires training data |
| CE6 | Full CEA | Baseline |
