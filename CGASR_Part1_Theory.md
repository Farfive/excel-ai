# CellGraph-Attention Spectral Retrieval (CGASR) — Part 1: Theory

**Version**: 1.0 | **Date**: 2026-03-13 | **Target**: Excel AI Financial Modeling Assistant

---

## 1. ALGORITHM NAME & ONE-LINE DESCRIPTION

**CellGraph-Attention Spectral Retrieval (CGASR)** — A query-conditioned, uncertainty-aware retrieval algorithm that fuses spectral graph wavelets, information-theoretic cell importance, cross-attention mechanisms, and tensor decomposition over spreadsheet cells for spatially-aware, formula-semantic, hierarchically-structured financial model retrieval.

---

## 2. MATHEMATICAL FOUNDATIONS

### 2.1 Frameworks Combined

| Framework | Role in CGASR | Solves Weakness |
|-----------|---------------|-----------------|
| **(a) Spectral Graph Theory** | Dual-graph Laplacian eigenvectors as positional encodings; Chebyshev wavelet filters for multi-scale dependency propagation | W1 (spatial blindness), W2 (static traversal) |
| **(c) Information Theory** | Conditional mutual information for query-dependent cell importance; entropy-based uncertainty | W3 (query-agnostic importance), W7 (no confidence) |
| **(d) Attention Mechanisms** | Multi-head cross-attention between query and spectral cell embeddings; role-weighted self-attention within chunks | W4 (flat chunks), W5 (formula semantics) |
| **(h) Tensor Decomposition** | CP decomposition of 4D cell tensor (row × col × sheet × type) | W1 (spatial), W6 (cross-sheet collapse) |

### 2.2 Novel Theoretical Contribution

**No prior work combines spectral graph wavelets with information-theoretic query conditioning over a tensor-decomposed cell representation for spreadsheet retrieval.**

1. **Dual-Graph Spectral Encoding**: Constructs two Laplacians — formula dependency + spatial adjacency — and fuses them. Spreadsheets uniquely have both explicit (formula) and implicit (grid proximity) structure simultaneously.

2. **Query-Conditioned Wavelets**: Hammond (2011) wavelets are static. CGASR dynamically sets wavelet scale mixture from query embedding's spectral energy — "what is WACC?" activates fine-scale wavelets; "how does WACC affect valuation?" activates coarse-scale wavelets.

3. **Formula Semantic Propagation**: Edge weights encode formula type (SUM=1.5, IF=0.8, VLOOKUP=0.6), fan-in normalization, and depth — the formula's mathematical MEANING propagates through the spectral decomposition.

4. **Spectral Gap Confidence**: The spectral gap (λ₂ - λ₁) of retrieved chunk subgraphs serves as confidence measure — large gap = coherent retrieval, small gap = fragmented/uncertain.

---

## 3. FORMAL ALGORITHM — NOTATION

| Symbol | Definition |
|--------|-----------|
| G = (V, E) | Dependency graph. V = cells, E = formula refs |
| N = \|V\| | Number of cells |
| c_i ∈ V | Cell node i |
| q ∈ ℝ^d | Query embedding (d = 768) |
| L_f, L_s | Formula / Spatial Laplacians |
| λ_k, u_k | k-th eigenvalue/eigenvector of combined L |
| K | Spectral components retained (min(64, N/10)) |
| Ψ_s(c_i) | Wavelet coefficient at scale s |
| π_q | Query-conditioned Personalized PageRank |
| σ_i | Uncertainty estimate for chunk i |

---

## 3.1 PREPROCESSING PHASE (runs once at upload)

### 3.1.1 Dual-Graph Laplacian

**Spatial Adjacency** G_s: connect cells to 8-neighbors (weight 1.0 orthogonal, 0.5 diagonal) + cross-sheet same-position (weight 0.3).

**(Eq. 1)** Combined normalized Laplacian:
```
L_combined = α · L̃_f + (1 - α) · L̃_s,  α = 0.7
```

### 3.1.2 Spectral Decomposition

**(Eq. 2)** K smallest eigenpairs via `scipy.sparse.linalg.eigsh`:
```
L_combined · u_k = λ_k · u_k,  k = 0..K-1
```

**(Eq. 3)** Spectral Positional Encoding:
```
PE(c_i) = [u_1[i], ..., u_K[i]] ∈ ℝ^K
```
Cells close in formula+spatial senses → similar PE. **Solves W1.**

### 3.1.3 Formula Semantic Edge Weights

**(Eq. 4)** For edge (i→j) in formula graph:
```
w_f(i,j) = β_type(j) · fan_in(j)^{-0.5}
```
where β_type: SUM=1.5, AVERAGE=1.3, reference=1.0, IF=0.8, VLOOKUP=0.6. **Solves W5.**

### 3.1.4 4D Cell Tensor + CP Decomposition

**(Eq. 5)** Tensor 𝒯 ∈ ℝ^{R_max × C_max × S × 4}:
```
𝒯[r, c, s, t] = normalized_value(cell(r,c,s)) if type matches t, else 0
```

**(Eq. 6)** CP decomposition (rank R=16):
```
𝒯 ≈ Σ_{r=1}^{16} a_r ⊗ b_r ⊗ c_r ⊗ d_r
```

Factor vectors capture: **a** = row patterns (header/data/total → W4), **b** = temporal columns, **c** = sheet similarity (P&L↔CF → W6), **d** = type regions.

**(Eq. 7)** Cell tensor embedding:
```
TE(c_i) = [a_{row_i} ∘ b_{col_i} ∘ c_{sheet_i} ∘ d_{type_i}]_{r=1}^{16} ∈ ℝ^{16}
```

### 3.1.5 Fused Cell Embedding

**(Eq. 8)**
```
E(c_i) = W_fuse · [PE(c_i) ‖ TE(c_i) ‖ SE(c_i)] + b_fuse
```
where SE = SentenceTransformer (768-dim). W_fuse is a fixed random projection (no training needed).

### 3.1.6 Wavelet Coefficients (4 scales)

**(Eq. 9)** Using wavelet kernel g(x) = x·exp(-x):
```
Ψ_s = U · diag(g(s·λ)) · U^T
```
Scales s ∈ {0.5, 2.0, 8.0, 32.0} = {ultra-local, cluster, sheet, cross-sheet}.

Chebyshev order-6 approximation avoids full eigendecomposition: O(6·|E|) per scale.

### 3.1.7 Cell Role Classification

**(Eq. 10)**
```
role(c_i) ∈ {HEADER(0.8), LINE_ITEM(1.0), SUBTOTAL(1.3),
              TOTAL(1.5), CROSS_REF(1.2), ASSUMPTION(1.4)}
```
Based on formula type, sheet name, is_hardcoded, SUM-of-SUMs detection. **Solves W4.**

---

## 3.2 QUERY-TIME PHASE (per question)

### 3.2.1 Adaptive Scale Selection

**(Eq. 11)** Project query onto spectral basis:
```
q_spectral = U^T · (E · q)  ∈ ℝ^K
```

**(Eq. 12)** Scale weights:
```
w_s = softmax(E_scale(s) / τ),  τ = 0.1
```
where E_scale(s) = Σ_{k: s·λ_k ∈ [0.5,2.0]} |q_spectral[k]|².

### 3.2.2 Query-Conditioned Personalized PageRank

**(Eq. 13)** Teleportation from query relevance:
```
v_q[i] = softmax(cos(E(c_i), q) / 0.5)
π_q = (1-α)·(I - α·P)^{-1} · v_q,  α = 0.85
```
Power iteration ~25 steps. **Solves W3** — importance is now query-dependent.

### 3.2.3 Multi-Scale Wavelet Propagation

**(Eq. 14)** Per-scale importance:
```
imp_s(c_i) = Σ_j Ψ_s[i,j] · π_q[j]
```

**(Eq. 15)** Fused:
```
IMP(c_i, q) = Σ_s w_s · imp_s(c_i)
```
Coarse wavelets (s=32) propagate Assumptions!B13 through entire DCF chain. **Solves W6.**

### 3.2.4 Chunk Scoring with Cross-Attention

**(Eq. 16)** Attention-pooled chunk representation:
```
h(C_j) = Σ_{c_i ∈ C_j} softmax(IMP(c_i,q) · role_weight(c_i)) · E(c_i)
```

**(Eq. 17)** Cross-attention score:
```
ATTN(C_j, q) = q^T · h(C_j) / ‖h(C_j)‖
```

### 3.2.5 Uncertainty Estimation

**(Eq. 18)** Spectral gap confidence:
```
gap(C_j) = λ_2(L_{C_j}) - λ_1(L_{C_j})
```

**(Eq. 19)** Entropy uncertainty:
```
σ(C_j, q) = H(softmax(rel(cells, q))) / log(|C_j|)  ∈ [0, 1]
```
**Solves W7.**

### 3.2.6 Final Score

**(Eq. 20)**
```
score(C_j, q) = sigmoid(0.4·ATTN + 0.25·max_IMP + 0.2·avg_π_q + 0.15·gap)
```

**(Eq. 21)** MMR with uncertainty penalty:
```
MMR(C_j) = 0.6 · score · (1 - σ) - 0.4 · max_sim_to_selected
```

---

## 4. EQUATIONS SUMMARY TABLE

| # | Equation | Purpose |
|---|----------|---------|
| 1 | L_combined = α·L̃_f + (1-α)·L̃_s | Dual-graph fusion |
| 2 | Eigendecomposition | Spectral basis |
| 3 | PE(c_i) = eigenvectors[i,:] | Positional encoding |
| 4 | w_f = β_type · fan_in^{-0.5} | Formula semantic weights |
| 5-6 | 4D tensor + CP decomp | Structural embedding |
| 7 | TE from CP factors | Tensor cell embedding |
| 8 | E = W·[PE‖TE‖SE]+b | Fused embedding |
| 9 | Ψ_s wavelet matrix | Multi-scale filters |
| 10 | role classification | Hierarchy encoding |
| 11-12 | q_spectral → w_s | Adaptive scale selection |
| 13 | Personalized PageRank | Query-conditioned importance |
| 14-15 | Wavelet propagation → IMP | Multi-scale importance |
| 16-17 | Attention-pooled → ATTN | Chunk-query matching |
| 18-19 | gap + entropy → σ | Uncertainty estimation |
| 20-21 | score + MMR | Final ranking |
