# CGASR Implementation Plan

**Total estimated time: 5 weeks | 4 phases**

---

## PHASE 1: Core Spectral Infrastructure (Week 1-2)

### Step 1.1 — Add dependency
**File**: `backend/requirements.txt`
**Action**: Add `tensorly==0.8.1`
**Time**: 5 min

### Step 1.2 — Create CGASRIndex dataclass
**File**: `backend/rag/cgasr_index.py` (NEW)
**Action**: Define the `CGASRIndex` dataclass that holds all precomputed structures:
```
- cell_embeddings: np.ndarray     # (N, 768) fused embeddings
- spectral_PE: np.ndarray         # (N, K) spectral positional encodings
- tensor_TE: np.ndarray           # (N, 16) tensor embeddings
- eigenvalues: np.ndarray         # (K,)
- eigenvectors: np.ndarray        # (N, K)
- wavelets: dict[float, sparse]   # scale → sparse wavelet matrix
- roles: dict[str, str]           # cell_addr → role name
- cell_to_idx: dict[str, int]     # cell_addr → integer index
- cell_ids: list[str]             # ordered cell addresses
- L: sparse matrix                # combined Laplacian
- P: sparse matrix                # column-normalized transition matrix
- scales: list[float]             # [0.5, 2.0, 8.0, 32.0]
```
**Time**: 30 min

### Step 1.3 — Spatial graph builder
**File**: `backend/rag/cgasr_index.py`
**Function**: `build_spatial_graph(workbook_data) → scipy.sparse.csr_matrix`
**Logic**:
1. Create spatial hash: `(sheet, row, col) → [cell_indices]`
2. For each cell, check 8 neighbors via hash (O(1) per neighbor)
3. Add cross-sheet same-position edges (weight 0.3)
4. Return sparse adjacency matrix A_s
**Depends on**: WorkbookData.cells
**Time**: 1 hour

### Step 1.4 — Formula semantic weights
**File**: `backend/rag/cgasr_index.py`
**Function**: `compute_formula_weights(graph_f, cell_to_idx) → scipy.sparse.csr_matrix`
**Logic**:
1. Iterate all edges in NetworkX DiGraph
2. Classify formula type → β_type (SUM=1.5, AVG=1.3, ref=1.0, IF=0.8, VLOOKUP=0.6)
3. Compute fan_in normalization
4. Build sparse weighted adjacency A_f
**Depends on**: existing `DependencyGraphBuilder` output
**Time**: 1 hour

### Step 1.5 — Combined Laplacian
**File**: `backend/rag/cgasr_index.py`
**Function**: `build_combined_laplacian(A_f, A_s, alpha=0.7) → sparse matrix`
**Logic**:
1. `L_f = scipy.sparse.csgraph.laplacian(A_f, normed=True)`
2. `L_s = scipy.sparse.csgraph.laplacian(A_s, normed=True)`
3. `L = alpha * L_f + (1 - alpha) * L_s`
**Depends on**: Steps 1.3 + 1.4
**Time**: 30 min

### Step 1.6 — Spectral decomposition
**File**: `backend/rag/cgasr_index.py`
**Function**: `spectral_decompose(L, K) → (eigenvalues, eigenvectors)`
**Logic**:
1. K = min(64, N // 10)
2. `scipy.sparse.linalg.eigsh(L, k=K, which='SM', sigma=1e-6)`
3. Sort by eigenvalue ascending
4. Handle failure gracefully (fallback to random PE)
**Depends on**: Step 1.5
**Performance target**: < 500ms for N=5000
**Time**: 1 hour

### Step 1.7 — Cell role classification
**File**: `backend/rag/cgasr_index.py`
**Function**: `classify_cell_roles(workbook_data, graph_f) → dict[str, str]`
**Logic**:
1. For each cell, check in order: ASSUMPTION → HEADER → CROSS_REF → TOTAL → SUBTOTAL → LINE_ITEM → COMPUTED
2. TOTAL = SUM formula whose inputs include other SUM cells
3. ASSUMPTION = Assumptions sheet + is_hardcoded
**Depends on**: WorkbookData + DiGraph
**Time**: 1 hour

### Step 1.8 — Unit tests for Phase 1
**File**: `backend/tests/test_cgasr_index.py` (NEW)
**Tests**:
- Spatial graph: 3×3 grid → each interior cell has 8 neighbors
- Formula weights: SUM formula → β=1.5
- Laplacian: eigenvalues are non-negative, smallest ≈ 0
- Spectral PE: dimension = (N, K)
- Role classification: known cell → expected role
**Time**: 2 hours

**Phase 1 Total: ~8 hours**

---

## PHASE 2: Tensor + Wavelet + Embedding Fusion (Week 2-3)

### Step 2.1 — 4D Tensor construction
**File**: `backend/rag/cgasr_index.py`
**Function**: `build_cell_tensor(workbook_data) → np.ndarray`
**Logic**:
1. Determine dimensions: R_max=min(max_row,100), C_max=min(max_col,26), S=sheets, T=4
2. Classify each cell type: numeric(0), formula(1), text(2), empty(3)
3. Normalize values: `tanh(value / 1e6)` for numeric, 0.5 for formula, 0.1 for text
4. Fill tensor[r-1, c-1, s, t]
**Time**: 1 hour

### Step 2.2 — CP Decomposition
**File**: `backend/rag/cgasr_index.py`
**Function**: `cp_decompose(tensor, rank=16) → (factor_row, factor_col, factor_sheet, factor_type)`
**Logic**:
1. `tensorly.decomposition.parafac(tensor, rank=16, n_iter_max=50)`
2. Extract 4 factor matrices
3. For each cell, compute TE[i] = fa[row] * fb[col] * fc[sheet] * fd[type]
4. Graceful fallback if tensorly fails
**Depends on**: Step 2.1
**Performance target**: < 200ms
**Time**: 1 hour

### Step 2.3 — Fused embedding computation
**File**: `backend/rag/cgasr_index.py`
**Function**: `compute_fused_embeddings(PE, TE, SE) → np.ndarray`
**Logic**:
1. Concatenate [PE(N,K) | TE(N,16) | SE(N,768)] → (N, K+16+768)
2. Fixed random projection W ∈ ℝ^{(K+16+768) × 768} (seed=42)
3. E = concat @ W, then L2 normalize
**Depends on**: Phase 1 PE + Step 2.2 TE + existing embedder SE
**Time**: 30 min

### Step 2.4 — Wavelet coefficient precomputation
**File**: `backend/rag/cgasr_index.py`
**Function**: `compute_wavelets(eigenvalues, eigenvectors, scales) → dict`
**Logic**:
1. For each scale s ∈ {0.5, 2.0, 8.0, 32.0}:
   - g_lambda = s * eigenvalues * exp(-s * eigenvalues)
   - W = U @ diag(g_lambda) @ U^T
   - Sparsify: zero out |W[i,j]| < 1e-4
   - Store as scipy.sparse.csr_matrix
**Depends on**: Phase 1 spectral decomposition
**Performance target**: < 100ms total
**Time**: 1 hour

### Step 2.5 — Assemble CGASRIndex builder
**File**: `backend/rag/cgasr_index.py`
**Function**: `build_cgasr_index(workbook_data, graph_f, embedder) → CGASRIndex`
**Logic**: Orchestrator function calling Steps 1.3 → 1.6 → 2.1 → 2.4 in order
**Performance target**: < 2s total for N=5000 (excluding SentenceTransformer embed)
**Time**: 1 hour

### Step 2.6 — Wire into upload route
**File**: `backend/api/routes/workbook.py`
**Action**: After `builder.run_algorithms()`, call `build_cgasr_index()` and store in a new `cgasr_indices` dict
**Changes**:
1. Add `cgasr_indices = Depends(get_cgasr_indices)` parameter
2. After line 63: `cgasr_idx = build_cgasr_index(workbook_data, graph, embedder)`
3. `cgasr_indices[workbook_uuid] = cgasr_idx`
**Time**: 30 min

### Step 2.7 — Unit tests for Phase 2
**File**: `backend/tests/test_cgasr_index.py` (extend)
**Tests**:
- Tensor shape = (R_max, C_max, S, 4)
- CP factors: fa.shape[1] == 16
- Fused embedding: shape = (N, 768), L2 normalized
- Wavelet matrices: 4 scales, each sparse (N, N)
- Full build_cgasr_index: runs without error on sample workbook
**Time**: 2 hours

**Phase 2 Total: ~7 hours**

---

## PHASE 3: Query-Time Retrieval (Week 3-4)

### Step 3.1 — CGASRRetriever class
**File**: `backend/rag/cgasr_retriever.py` (NEW)
**Class**: `CGASRRetriever`
**Constructor**: Takes `embedder`, `store`, `ollama_client`, `graph`, `cgasr_index`
**Time**: 30 min

### Step 3.2 — Query routing
**File**: `backend/rag/cgasr_retriever.py`
**Function**: `_route_query(query) → "direct" | "cea" | "cgasr"`
**Logic**:
1. Check direct lookup regex → "direct"
2. Check cross-sheet/sensitivity/reasoning keywords → "cgasr"
3. Check workbook size > 8000 + simple question → "cea"
4. Default → "cgasr"
**Time**: 30 min

### Step 3.3 — Adaptive scale selection
**File**: `backend/rag/cgasr_retriever.py`
**Function**: `_select_scales(q, cgasr_index) → np.ndarray`
**Logic**:
1. cell_relevances = E @ q
2. q_spectral = U^T @ cell_relevances
3. For each scale: energy = sum of q_spectral² where s*λ ∈ [0.5, 2.0]
4. w_scales = softmax(energies / 0.1)
**Depends on**: CGASRIndex
**Time**: 30 min

### Step 3.4 — Personalized PageRank
**File**: `backend/rag/cgasr_retriever.py`
**Function**: `_personalized_pagerank(relevances, P, alpha=0.85, iters=25) → np.ndarray`
**Logic**:
1. teleport = softmax(relevances / 0.5)
2. pi = teleport
3. For 25 iterations: pi = alpha * P @ pi + (1-alpha) * teleport
4. Normalize pi
**Performance target**: < 150ms (use sparse P if possible, dense fallback for small N)
**Time**: 1 hour

### Step 3.5 — Multi-scale wavelet propagation
**File**: `backend/rag/cgasr_retriever.py`
**Function**: `_wavelet_propagate(pi, wavelets, w_scales) → np.ndarray`
**Logic**:
1. For each scale: imp_s = wavelet_matrix @ pi
2. importance = Σ w_s * imp_s
**Time**: 30 min

### Step 3.6 — Chunk scoring with attention + uncertainty
**File**: `backend/rag/cgasr_retriever.py`
**Function**: `_score_chunks(chunks, importance, pi, E, q, roles, L) → list[dict]`
**Logic**:
1. Per chunk: get cell indices from metadata
2. Role-weighted attention pooling → h_chunk
3. Cross-attention score = q^T · h_chunk / ||h_chunk||
4. Max importance + avg PPR
5. Spectral gap of induced subgraph (small eigendecomp)
6. Entropy-based uncertainty σ
7. Final: sigmoid(0.4*attn + 0.25*max_imp + 0.2*avg_ppr + 0.15*gap)
**Time**: 2 hours

### Step 3.7 — MMR ranking with uncertainty
**File**: `backend/rag/cgasr_retriever.py`
**Function**: `_mmr_rank(scored_chunks, E, cell_to_idx, k=10, lambda_=0.6) → list[dict]`
**Logic**:
1. Sort by score * (1 - σ)
2. Greedy MMR: add chunk if λ*score*(1-σ) - (1-λ)*max_sim > threshold
**Time**: 30 min

### Step 3.8 — Main retrieve() method
**File**: `backend/rag/cgasr_retriever.py`
**Function**: `async retrieve(query, workbook_uuid, k=15, workbook_data=None) → list[dict]`
**Logic**: Orchestrate steps 3.2 → 3.7, keep valuation injection from existing pipeline
**Time**: 1 hour

### Step 3.9 — CEA retriever
**File**: `backend/rag/cea_retriever.py` (NEW)
**Class**: `CEARetriever`
**Functions**:
1. `cea_preprocess(workbook_data)` → CEAIndex (entropy + fixed random vectors)
2. `retrieve(query, ...)` → chunks via entropy + CMI + 2-layer GAT
3. Batched sparse GAT for speed
**Performance target**: < 150ms per query
**Time**: 3 hours

### Step 3.10 — Swap retriever in ExcelAgent
**File**: `backend/agent/excel_agent.py`
**Action**: 
1. ExcelAgent.__init__ accepts optional `cgasr_index` parameter
2. In execute() at line 1015: use CGASRRetriever if cgasr_index available, else fallback to RAGRetriever
3. Pass uncertainty scores through to context builder
**Time**: 1 hour

### Step 3.11 — Wire into /ask endpoint
**File**: `backend/api/routes/workbook.py`
**Action**: When creating ExcelAgent for /ask, pass cgasr_index from the indices dict
**Time**: 30 min

### Step 3.12 — Unit + integration tests
**File**: `backend/tests/test_cgasr_retriever.py` (NEW)
**Tests**:
- Scale selection: "what is WACC?" → fine scale, "how does WACC affect..." → coarse scale
- PPR: query-relevant cells get higher π_q than irrelevant
- Chunk scoring: DCF chunks score high for valuation queries
- Uncertainty: well-connected chunks have low σ
- End-to-end: retrieve() returns k chunks with scores + σ
- CEA: retrieve completes in < 500ms
**Time**: 3 hours

**Phase 3 Total: ~14 hours**

---

## PHASE 4: Evaluation & Tuning (Week 4-5)

### Step 4.1 — Benchmark dataset
**File**: `backend/rag/cgasr_eval.py` (NEW)
**Action**: Create 100 queries across categories:
- 20 direct lookup ("What is WACC?", "Read Revenue for 2025")
- 20 local reasoning ("What are the Revenue line items?", "List all assumptions")
- 20 cross-sheet ("How does WACC affect price per share?", "Trace Revenue to EV")
- 15 sensitivity ("What if Revenue grows 5% faster?", "Impact of tax rate change")
- 15 structural ("Which cells feed into EBIT?", "What formulas use Assumptions!B13?")
- 10 anomaly ("Any unusual values in P&L?", "Check for outliers")
**Annotation**: For each query, mark ground-truth relevant chunks
**Time**: 4 hours

### Step 4.2 — Metric computation
**File**: `backend/rag/cgasr_eval.py`
**Functions**:
- `ndcg_at_k(predicted, truth, k=5)`
- `mrr(predicted, truth)`
- `hit_at_k(predicted, truth, k=1)`
- `context_precision(retrieved_text, relevant_cells)`
- `cross_sheet_recall(retrieved_chunks, required_sheets)`
- `uncertainty_calibration(scores, sigmas, actual_relevance)`
**Time**: 2 hours

### Step 4.3 — Ablation runner
**File**: `backend/rag/cgasr_eval.py`
**Function**: `run_ablation(workbook_path, queries, config) → results`
**Configs to test**:
1. Full CGASR (all components)
2. Remove spectral PE (use SE + TE only)
3. Remove tensor TE (use SE + PE only)
4. Static PageRank (no PPR)
5. Single wavelet scale (s=2 only)
6. No role weighting
7. No uncertainty in MMR
8. Replace attention with cosine
9. CEA only
10. Current pipeline (baseline)
**Time**: 3 hours

### Step 4.4 — Hyperparameter sweep
**Parameters to tune**:
- α (Laplacian mix): [0.5, 0.6, 0.7, 0.8]
- K (spectral components): [16, 32, 64]
- CP rank: [8, 16, 32]
- τ (scale temperature): [0.05, 0.1, 0.2]
- γ weights: grid search
- PPR iterations: [15, 25, 35]
**Method**: Grid search on benchmark, pick best nDCG@5
**Time**: 3 hours

### Step 4.5 — Latency profiling
**Action**: Profile each step on M-series Mac with N=1000, 3000, 5000, 8000 cells
**Target**: preprocessing < 2s, per-query < 500ms at all sizes
**Optimize**: If PPR too slow → use sparse matrix; if spectral too slow → reduce K
**Time**: 2 hours

### Step 4.6 — Results report
**File**: `backend/rag/cgasr_eval_results.json`
**Content**: All metrics per ablation, latency numbers, best hyperparameters
**Time**: 1 hour

**Phase 4 Total: ~15 hours**

---

## DEPENDENCY ORDER GRAPH

```
Phase 1:                    Phase 2:                Phase 3:               Phase 4:
                                                    
1.1 requirements.txt        2.1 tensor ──┐          3.1 class              4.1 benchmark
       │                    2.2 CP ──────┤          3.2 router             4.2 metrics
1.2 dataclass               2.3 fuse ◄───┤          3.3 scales             4.3 ablation
       │                    2.4 wavelets  │          3.4 PPR                4.4 hyperparam
1.3 spatial ─┐               │            │          3.5 wavelet prop       4.5 latency
1.4 formula ─┤              2.5 builder ◄─┘          3.6 scoring            4.6 report
1.5 laplacian◄┘              │                       3.7 MMR
       │                    2.6 upload wire           3.8 main retrieve
1.6 spectral                 │                       3.9 CEA
       │                    2.7 tests                3.10 agent swap
1.7 roles                                           3.11 endpoint wire
       │                                            3.12 tests
1.8 tests
```

---

## FILES TOUCHED — SUMMARY

| File | Action | Phase |
|------|--------|-------|
| `backend/requirements.txt` | Add `tensorly==0.8.1` | 1 |
| `backend/rag/cgasr_index.py` | **NEW** — CGASRIndex + all preprocessing | 1-2 |
| `backend/rag/cgasr_retriever.py` | **NEW** — CGASRRetriever + query-time | 3 |
| `backend/rag/cea_retriever.py` | **NEW** — CEA fast fallback | 3 |
| `backend/rag/cgasr_eval.py` | **NEW** — evaluation harness | 4 |
| `backend/api/routes/workbook.py` | **MODIFY** — add CGASR init in upload | 2 |
| `backend/agent/excel_agent.py` | **MODIFY** — swap retriever | 3 |
| `backend/tests/test_cgasr_index.py` | **NEW** — unit tests | 1-2 |
| `backend/tests/test_cgasr_retriever.py` | **NEW** — integration tests | 3 |

**4 new files, 2 modified files, 2 new test files**

---

## RISK MITIGATION

| Risk | Mitigation |
|------|------------|
| `eigsh` fails on disconnected graph | Try/catch, fallback to random PE |
| `tensorly` unavailable | Try/catch, fallback to zero TE |
| PPR too slow (dense P) | Use sparse matrix multiplication |
| Spectral gap eigendecomp on tiny chunks | Default gap=0.5 for |C| < 3 |
| Memory > 100MB for large workbooks | Cap N at 10000, R_max at 100, C_max at 26 |
| Existing tests break | CGASR is additive — RAGRetriever unchanged, only swap at agent level |

---

## ROLLBACK STRATEGY

The implementation is **fully additive** — no existing code is deleted. If CGASR causes issues:
1. In `excel_agent.py`: simply don't pass `cgasr_index` → falls back to existing RAGRetriever
2. In `workbook.py`: skip `build_cgasr_index()` call → no preprocessing overhead
3. All existing tests pass unchanged
