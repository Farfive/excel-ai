"""
CellGraph-Attention Spectral Retrieval (CGASR) — Preprocessing Index

Builds all precomputed structures at workbook upload time:
- Dual-graph Laplacian (formula + spatial)
- Spectral positional encodings
- 4D tensor CP decomposition
- Fused cell embeddings
- Multi-scale wavelet coefficients
- Cell role classification
"""

import json
import logging
import os
import pickle
import time
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

import networkx as nx
import numpy as np
import scipy.sparse
import scipy.sparse.linalg
import scipy.sparse.csgraph

logger = logging.getLogger(__name__)

ROLE_WEIGHTS = {
    "TOTAL": 1.5,
    "SUBTOTAL": 1.3,
    "CROSS_REF": 1.2,
    "ASSUMPTION": 1.4,
    "LINE_ITEM": 1.0,
    "HEADER": 0.8,
    "COMPUTED": 1.0,
}

FORMULA_BETA = {
    "SUM": 1.5, "SUMIF": 1.5, "SUMPRODUCT": 1.5,
    "AVERAGE": 1.3, "AVG": 1.3,
    "IF": 0.8, "IFERROR": 0.8, "IFS": 0.8,
    "VLOOKUP": 0.6, "INDEX": 0.6, "MATCH": 0.6, "HLOOKUP": 0.6,
}


@dataclass
class CGASRIndex:
    cell_embeddings: np.ndarray       # (N, 768) fused
    spectral_PE: np.ndarray           # (N, K)
    tensor_TE: np.ndarray             # (N, 16)
    eigenvalues: np.ndarray           # (K,)
    eigenvectors: np.ndarray          # (N, K)
    wavelets: Dict[float, Any]        # scale → sparse matrix
    roles: Dict[str, str]             # addr → role
    cell_to_idx: Dict[str, int]       # addr → int
    cell_ids: List[str]               # ordered addresses
    L: Any                            # combined Laplacian (sparse)
    P: Any                            # transition matrix (sparse)
    scales: List[float]
    N: int
    K: int
    build_time_ms: int = 0

    def save(self, path: str) -> None:
        """Save CGASR index to disk for fast reload after server restart."""
        os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
        t0 = time.time()
        data = {
            "arrays": {
                "cell_embeddings": self.cell_embeddings,
                "spectral_PE": self.spectral_PE,
                "tensor_TE": self.tensor_TE,
                "eigenvalues": self.eigenvalues,
                "eigenvectors": self.eigenvectors,
            },
            "sparse": {
                "L": self.L,
                "P": self.P,
            },
            "meta": {
                "wavelets": self.wavelets,
                "roles": self.roles,
                "cell_to_idx": self.cell_to_idx,
                "cell_ids": self.cell_ids,
                "scales": self.scales,
                "N": self.N,
                "K": self.K,
                "build_time_ms": self.build_time_ms,
            },
        }
        with open(path, "wb") as f:
            pickle.dump(data, f, protocol=pickle.HIGHEST_PROTOCOL)
        ms = int((time.time() - t0) * 1000)
        size_mb = os.path.getsize(path) / (1024 * 1024)
        logger.info(f"CGASR index saved: {path} ({size_mb:.1f}MB) in {ms}ms")

    @classmethod
    def load(cls, path: str) -> "CGASRIndex":
        """Load CGASR index from disk."""
        t0 = time.time()
        with open(path, "rb") as f:
            data = pickle.load(f)
        idx = cls(
            cell_embeddings=data["arrays"]["cell_embeddings"],
            spectral_PE=data["arrays"]["spectral_PE"],
            tensor_TE=data["arrays"]["tensor_TE"],
            eigenvalues=data["arrays"]["eigenvalues"],
            eigenvectors=data["arrays"]["eigenvectors"],
            wavelets=data["meta"]["wavelets"],
            roles=data["meta"]["roles"],
            cell_to_idx=data["meta"]["cell_to_idx"],
            cell_ids=data["meta"]["cell_ids"],
            L=data["sparse"]["L"],
            P=data["sparse"]["P"],
            scales=data["meta"]["scales"],
            N=data["meta"]["N"],
            K=data["meta"]["K"],
            build_time_ms=data["meta"]["build_time_ms"],
        )
        ms = int((time.time() - t0) * 1000)
        logger.info(f"CGASR index loaded: {path} (N={idx.N}, K={idx.K}) in {ms}ms")
        return idx


def _cell_type_id(cell) -> int:
    if cell.formula:
        return 1
    if isinstance(cell.value, (int, float)):
        return 0
    if isinstance(cell.value, str) and cell.value.strip():
        return 2
    return 3


def _normalize_value(cell) -> float:
    if isinstance(cell.value, (int, float)):
        return float(np.tanh(cell.value / 1e6))
    return 0.5 if cell.formula else (0.1 if isinstance(cell.value, str) else 0.0)


def _cell_to_text(cell) -> str:
    parts = [cell.cell_address]
    if cell.named_range:
        parts.append(cell.named_range)
    if cell.formula:
        parts.append(f"={cell.formula}")
    if cell.value is not None:
        parts.append(str(cell.value)[:50])
    return " ".join(parts)


def _get_formula_beta(formula: str) -> float:
    fu = formula.upper()
    for key, beta in FORMULA_BETA.items():
        if key + "(" in fu:
            return beta
    return 1.0


def _classify_role(cell, graph_f) -> str:
    formula = (cell.formula or "").upper()
    addr = cell.cell_address

    if cell.sheet_name == "Assumptions" and cell.is_hardcoded:
        return "ASSUMPTION"
    if isinstance(cell.value, str) and not cell.formula:
        return "HEADER"
    if "!" in (cell.formula or ""):
        return "CROSS_REF"
    if "SUM" in formula:
        if addr in graph_f:
            for pred in graph_f.predecessors(addr):
                pf = (graph_f.nodes[pred].get("formula") or "").upper()
                if "SUM" in pf:
                    return "TOTAL"
        return "SUBTOTAL"
    if isinstance(cell.value, (int, float)):
        return "LINE_ITEM"
    return "COMPUTED"


def build_cgasr_index(
    workbook_data,
    graph_f: nx.DiGraph,
    existing_embeddings: Optional[np.ndarray] = None,
    embedder=None,
) -> CGASRIndex:
    """
    Build CGASR index from workbook data and dependency graph.
    
    Args:
        workbook_data: WorkbookData with .cells, .sheets
        graph_f: NetworkX DiGraph (formula dependency graph)
        existing_embeddings: (N, 768) pre-computed embeddings (optional)
        embedder: LocalEmbedder instance (used if existing_embeddings is None)
    """
    t0 = time.time()
    cells_list = list(workbook_data.cells.values())
    cell_ids = list(workbook_data.cells.keys())
    N = len(cells_list)

    if N == 0:
        logger.warning("Empty workbook, returning empty CGASR index")
        return _empty_index()

    cell_to_idx = {addr: i for i, addr in enumerate(cell_ids)}
    logger.info(f"CGASR: Building index for {N} cells across {len(workbook_data.sheets)} sheets")

    # ── 1. Spatial adjacency graph (O(N) via hashing) ─────────────────
    t1 = time.time()
    spatial_hash: Dict[Tuple, List[int]] = defaultdict(list)
    for i, ci in enumerate(cells_list):
        spatial_hash[(ci.sheet_name, ci.row, ci.col)].append(i)

    rows_s, cols_s, vals_s = [], [], []
    for i, ci in enumerate(cells_list):
        for dr in [-1, 0, 1]:
            for dc in [-1, 0, 1]:
                if dr == 0 and dc == 0:
                    continue
                key = (ci.sheet_name, ci.row + dr, ci.col + dc)
                for j in spatial_hash.get(key, []):
                    w = 0.5 if (abs(dr) == 1 and abs(dc) == 1) else 1.0
                    rows_s.append(i)
                    cols_s.append(j)
                    vals_s.append(w)
        # Cross-sheet same position
        for sheet in workbook_data.sheets:
            if sheet != ci.sheet_name:
                for j in spatial_hash.get((sheet, ci.row, ci.col), []):
                    rows_s.append(i)
                    cols_s.append(j)
                    vals_s.append(0.3)

    A_s = scipy.sparse.csr_matrix(
        (vals_s, (rows_s, cols_s)), shape=(N, N)
    ) if vals_s else scipy.sparse.csr_matrix((N, N))
    logger.info(f"CGASR: Spatial graph: {len(vals_s)} edges in {(time.time()-t1)*1000:.0f}ms")

    # ── 2. Formula adjacency with semantic weights ────────────────────
    t1 = time.time()
    rows_f, cols_f, vals_f = [], [], []
    for src, dst in graph_f.edges():
        si = cell_to_idx.get(src)
        di = cell_to_idx.get(dst)
        if si is not None and di is not None:
            dst_formula = (graph_f.nodes[dst].get("formula") or "")
            beta = _get_formula_beta(dst_formula)
            fan_in = max(graph_f.in_degree(dst), 1)
            w = beta * (fan_in ** -0.5)
            rows_f.append(si); cols_f.append(di); vals_f.append(w)
            rows_f.append(di); cols_f.append(si); vals_f.append(w)

    A_f = scipy.sparse.csr_matrix(
        (vals_f, (rows_f, cols_f)), shape=(N, N)
    ) if vals_f else scipy.sparse.csr_matrix((N, N))
    logger.info(f"CGASR: Formula graph: {len(vals_f)//2} edges in {(time.time()-t1)*1000:.0f}ms")

    # ── 3. Combined Laplacian ─────────────────────────────────────────
    t1 = time.time()
    alpha = 0.7
    try:
        L_f = scipy.sparse.csgraph.laplacian(A_f, normed=True)
        L_s = scipy.sparse.csgraph.laplacian(A_s, normed=True)
        L = alpha * L_f + (1 - alpha) * L_s
    except Exception as e:
        logger.warning(f"CGASR: Laplacian construction failed: {e}")
        L = scipy.sparse.eye(N)
    logger.info(f"CGASR: Combined Laplacian in {(time.time()-t1)*1000:.0f}ms")

    # ── 4. Spectral decomposition ─────────────────────────────────────
    t1 = time.time()
    K = min(32, max(N // 20, 4))
    K = min(K, N - 2) if N > 3 else 2

    try:
        if N < 500:
            L_dense = L.toarray() if scipy.sparse.issparse(L) else np.array(L)
            all_evals, all_evecs = np.linalg.eigh(L_dense.astype(np.float64))
            eigenvalues = np.abs(all_evals[:K])
            eigenvectors = all_evecs[:, :K]
        else:
            eigenvalues, eigenvectors = scipy.sparse.linalg.eigsh(
                L.astype(np.float64), k=K, which='SA', maxiter=300,
            )
            sort_idx = np.argsort(eigenvalues)
            eigenvalues = np.abs(eigenvalues[sort_idx])
            eigenvectors = eigenvectors[:, sort_idx]
        logger.info(f"CGASR: Spectral decomp K={K}, λ range [{eigenvalues[0]:.4f}, {eigenvalues[-1]:.4f}] in {(time.time()-t1)*1000:.0f}ms")
    except Exception as e:
        logger.warning(f"CGASR: eigsh failed ({e}), using random PE")
        rng = np.random.RandomState(42)
        eigenvalues = np.sort(np.abs(rng.randn(K)))
        eigenvectors = rng.randn(N, K) * 0.01

    PE = eigenvectors  # (N, K)

    # ── 5. 4D Tensor + CP decomposition ───────────────────────────────
    t1 = time.time()
    R_max = min(max((c.row for c in cells_list), default=1), 100)
    C_max = min(max((c.col for c in cells_list), default=1), 26)
    S = len(workbook_data.sheets)
    sheet_idx_map = {s: i for i, s in enumerate(workbook_data.sheets)}
    T_types = 4

    tensor = np.zeros((R_max, C_max, S, T_types), dtype=np.float32)
    for ci in cells_list:
        r = ci.row - 1
        c = ci.col - 1
        s = sheet_idx_map.get(ci.sheet_name, 0)
        if r < R_max and c < C_max:
            t = _cell_type_id(ci)
            tensor[r, c, s, t] = _normalize_value(ci)

    try:
        import tensorly
        from tensorly.decomposition import parafac
        cp_rank = 16
        factors = parafac(tensorly.tensor(tensor), rank=cp_rank, n_iter_max=50, random_state=42)
        fa, fb, fc, fd = [np.array(f) for f in factors.factors]

        TE = np.zeros((N, cp_rank), dtype=np.float32)
        for i, ci in enumerate(cells_list):
            r = ci.row - 1
            c = ci.col - 1
            s = sheet_idx_map.get(ci.sheet_name, 0)
            t_id = _cell_type_id(ci)
            if r < R_max and c < C_max:
                TE[i] = fa[r] * fb[c] * fc[s] * fd[t_id]
        logger.info(f"CGASR: Tensor CP decomp ({R_max}×{C_max}×{S}×{T_types}, rank={cp_rank}) in {(time.time()-t1)*1000:.0f}ms")
    except Exception as e:
        logger.warning(f"CGASR: Tensor decomp failed ({e}), using zero TE")
        TE = np.zeros((N, 16), dtype=np.float32)

    # ── 6. Fused cell embeddings ──────────────────────────────────────
    t1 = time.time()
    if existing_embeddings is not None:
        SE = np.array(existing_embeddings, dtype=np.float32)
        if SE.shape[0] != N:
            logger.warning(f"CGASR: Embedding count mismatch ({SE.shape[0]} vs {N}), recomputing")
            SE = None
    else:
        SE = None

    if SE is None and embedder is not None:
        texts = [_cell_to_text(ci) for ci in cells_list]
        SE = np.array(embedder.embed(texts), dtype=np.float32)
    elif SE is None:
        logger.warning("CGASR: No embeddings available, using random SE")
        rng = np.random.RandomState(42)
        SE = rng.randn(N, 768).astype(np.float32) * 0.01

    combined = np.hstack([PE.astype(np.float32), TE, SE])  # (N, K+16+768)
    rng = np.random.RandomState(42)
    d_in = combined.shape[1]
    d_out = 768
    W_fuse = (rng.randn(d_in, d_out) / np.sqrt(d_in)).astype(np.float32)
    E = combined @ W_fuse
    norms = np.linalg.norm(E, axis=1, keepdims=True) + 1e-10
    E = (E / norms).astype(np.float32)
    logger.info(f"CGASR: Fused embeddings ({N}, {d_out}) from ({N}, {d_in}) in {(time.time()-t1)*1000:.0f}ms")

    # ── 7. Wavelet coefficients ───────────────────────────────────────
    t1 = time.time()
    scales = [0.5, 2.0, 8.0, 32.0]
    wavelets = {}
    for s in scales:
        try:
            g_lambda = s * eigenvalues * np.exp(-s * eigenvalues)
            wavelets[s] = {"g_lambda": g_lambda.astype(np.float32)}
        except Exception as e:
            logger.warning(f"CGASR: Wavelet scale {s} failed: {e}")
            wavelets[s] = {"g_lambda": np.ones(K, dtype=np.float32)}
    logger.info(f"CGASR: Wavelets ({len(scales)} scales) in {(time.time()-t1)*1000:.0f}ms")

    # ── 8. Cell roles ─────────────────────────────────────────────────
    roles = {}
    for addr, ci in workbook_data.cells.items():
        roles[addr] = _classify_role(ci, graph_f)

    role_counts = defaultdict(int)
    for r in roles.values():
        role_counts[r] += 1
    logger.info(f"CGASR: Roles: {dict(role_counts)}")

    # ── 9. Transition matrix for PPR ──────────────────────────────────
    col_sums = np.array(A_f.sum(axis=0)).flatten() + 1e-10
    P = A_f.multiply(1.0 / col_sums)

    build_ms = int((time.time() - t0) * 1000)
    logger.info(f"CGASR: Index built in {build_ms}ms (N={N}, K={K})")

    return CGASRIndex(
        cell_embeddings=E,
        spectral_PE=PE,
        tensor_TE=TE,
        eigenvalues=eigenvalues,
        eigenvectors=eigenvectors,
        wavelets=wavelets,
        roles=roles,
        cell_to_idx=cell_to_idx,
        cell_ids=cell_ids,
        L=L,
        P=P,
        scales=scales,
        N=N,
        K=K,
        build_time_ms=build_ms,
    )


def _empty_index() -> CGASRIndex:
    return CGASRIndex(
        cell_embeddings=np.zeros((0, 768)),
        spectral_PE=np.zeros((0, 2)),
        tensor_TE=np.zeros((0, 16)),
        eigenvalues=np.zeros(2),
        eigenvectors=np.zeros((0, 2)),
        wavelets={},
        roles={},
        cell_to_idx={},
        cell_ids=[],
        L=scipy.sparse.eye(1),
        P=scipy.sparse.eye(1),
        scales=[0.5, 2.0, 8.0, 32.0],
        N=0,
        K=2,
    )
