"""
CellGraph-Attention Spectral Retrieval (CGASR) — Query-Time Retriever

Implements query-conditioned retrieval:
- Adaptive wavelet scale selection
- Personalized PageRank with query teleportation
- Multi-scale wavelet importance propagation
- Cross-attention chunk scoring with uncertainty
- MMR diversity ranking
"""

import logging
import re
import time
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import scipy.sparse

from rag.cgasr_index import CGASRIndex, ROLE_WEIGHTS

logger = logging.getLogger(__name__)

_CROSS_SHEET_RE = re.compile(
    r"\b(wacc|dcf|valuation|sensitivity|propagat|affect|impact|chain|trace|flow|"
    r"enterprise value|equity value|price per share|terminal|discount|npv|irr|"
    r"wycena|wpływ|wpływa|zmiana|zmień)\b",
    re.IGNORECASE,
)

_SIMPLE_RE = re.compile(
    r"^(?:what is|jaki jest|jaka jest|ile wynosi|podaj|read|odczytaj|list|pokaż)\b",
    re.IGNORECASE,
)


def _softmax(x: np.ndarray) -> np.ndarray:
    e = np.exp(x - np.max(x))
    return e / (e.sum() + 1e-10)


def _sigmoid(x: float) -> float:
    return 1.0 / (1.0 + np.exp(-x))


class CGASRRetriever:
    """Query-time CGASR retrieval engine."""

    def __init__(self, cgasr_index: CGASRIndex):
        self.idx = cgasr_index

    def route_query(self, query: str) -> str:
        """Decide: 'cgasr' (full) or 'cea' (fast)."""
        if _CROSS_SHEET_RE.search(query):
            return "cgasr"
        if self.idx.N > 8000 and _SIMPLE_RE.match(query):
            return "cea"
        return "cgasr"

    def retrieve(
        self,
        query_embedding: np.ndarray,
        chunks: List[Dict],
        k: int = 10,
    ) -> List[Dict]:
        """
        Full CGASR retrieval pipeline.
        
        Args:
            query_embedding: (768,) query vector from SentenceTransformer
            chunks: list of chunk dicts with 'metadata.cell_addresses'
            k: number of chunks to return
            
        Returns:
            list of dicts with 'chunk', 'score', 'uncertainty', 'debug'
        """
        t0 = time.time()
        idx = self.idx
        N = idx.N
        q = np.array(query_embedding, dtype=np.float32)

        if N == 0 or not chunks:
            return []

        # ── 1. Cell relevances ────────────────────────────────────────
        rel = idx.cell_embeddings @ q  # (N,) cosine similarities

        # ── 2. Adaptive scale selection ───────────────────────────────
        q_spectral = idx.eigenvectors.T @ rel  # (K,)
        scale_energies = []
        for s in idx.scales:
            mask = (s * idx.eigenvalues >= 0.5) & (s * idx.eigenvalues <= 2.0)
            e = np.sum(q_spectral[mask] ** 2) + 1e-10
            scale_energies.append(e)
        w_scales = _softmax(np.array(scale_energies) / 0.1)

        dominant_scale = idx.scales[np.argmax(w_scales)]
        logger.info(f"CGASR: Scale weights {dict(zip(idx.scales, w_scales.tolist()))}, dominant={dominant_scale}")

        # ── 3. Personalized PageRank (25 iterations) ─────────────────
        teleport = _softmax(rel / 0.5)
        pi = teleport.copy()

        P_arr = idx.P
        if scipy.sparse.issparse(P_arr):
            P_arr = P_arr.toarray()
        P_arr = np.array(P_arr, dtype=np.float64)

        for _ in range(25):
            pi = 0.85 * (P_arr @ pi) + 0.15 * teleport
        pi_sum = pi.sum()
        if pi_sum > 0:
            pi = pi / pi_sum

        # Top cells by PPR
        top_ppr_idx = np.argsort(pi)[-5:][::-1]
        top_ppr = [(idx.cell_ids[i], float(pi[i])) for i in top_ppr_idx]
        logger.info(f"CGASR: Top PPR cells: {top_ppr}")

        # ── 4. Multi-scale wavelet importance (spectral domain) ────────
        importance = np.zeros(N, dtype=np.float64)
        pi_spectral = idx.eigenvectors.T @ pi  # (K,) project pi to spectral
        for i, s in enumerate(idx.scales):
            wdata = idx.wavelets[s]
            if isinstance(wdata, dict) and "g_lambda" in wdata:
                g_lam = wdata["g_lambda"]
                filtered = g_lam * pi_spectral  # (K,) element-wise
                imp_s = idx.eigenvectors @ filtered  # (N,) back to node domain
            elif scipy.sparse.issparse(wdata):
                imp_s = wdata.dot(pi)
            else:
                imp_s = np.array(wdata) @ pi
            importance += w_scales[i] * imp_s

        # Normalize importance to [0, 1]
        imp_range = importance.max() - importance.min()
        if imp_range > 0:
            importance = (importance - importance.min()) / imp_range

        # ── 5. Score each chunk ───────────────────────────────────────
        results = []
        for chunk in chunks:
            addrs = chunk.get("metadata", {}).get("cell_addresses", [])
            if isinstance(addrs, str):
                addrs = addrs.split(",") if addrs else []
            ci = [idx.cell_to_idx[a] for a in addrs if a in idx.cell_to_idx]

            if not ci:
                results.append({
                    "chunk": chunk,
                    "score": 0.0,
                    "uncertainty": 1.0,
                    "debug": {"reason": "no_cells_mapped"},
                })
                continue

            ci_arr = np.array(ci)

            # Role-weighted attention pooling → chunk representation
            rw = np.array([
                ROLE_WEIGHTS.get(idx.roles.get(addrs[j] if j < len(addrs) else "", "COMPUTED"), 1.0)
                for j in range(len(ci))
            ])
            attn_raw = importance[ci_arr] * rw
            attn_w = _softmax(attn_raw * 5.0)  # temperature sharpening
            h_chunk = attn_w @ idx.cell_embeddings[ci_arr]  # (768,)

            # Cross-attention score
            h_norm = np.linalg.norm(h_chunk) + 1e-10
            attn_score = float(np.dot(q, h_chunk) / h_norm)

            # Max importance + avg PPR
            max_imp = float(np.max(importance[ci_arr]))
            avg_ppr = float(np.mean(pi[ci_arr]))

            # Spectral gap of induced subgraph
            if len(ci) >= 3:
                try:
                    sub_L = idx.L[np.ix_(ci_arr, ci_arr)]
                    if scipy.sparse.issparse(sub_L):
                        sub_L = sub_L.toarray()
                    eigs = np.sort(np.linalg.eigvalsh(sub_L))
                    gap = float(eigs[1] - eigs[0]) if len(eigs) > 1 else 0.0
                    gap = max(0.0, min(gap, 2.0))  # clamp
                except Exception:
                    gap = 0.5
            else:
                gap = 0.5

            # Uncertainty: variance of importance + entropy of relevance
            if len(ci) >= 2:
                imp_vals = importance[ci_arr]
                imp_std = float(np.std(imp_vals))
                imp_rng = float(importance.max() - importance.min()) + 1e-10
                var_component = imp_std / imp_rng

                p = _softmax(rel[ci_arr])
                entropy = -np.sum(p * np.log(p + 1e-10))
                max_entropy = np.log(len(ci))
                ent_component = float(entropy / max_entropy) if max_entropy > 0 else 0.5

                sigma = 0.5 * var_component + 0.5 * ent_component
            else:
                # Single cell: uncertainty based on how much the cell stands out
                cell_rel = float(rel[ci_arr[0]])
                rel_rank = float(np.sum(rel > cell_rel)) / N
                sigma = rel_rank  # high rank = low uncertainty
            sigma = max(0.01, min(sigma, 0.99))

            # Final score (sigmoid)
            gamma = [0.4, 0.25, 0.2, 0.15]
            raw = (gamma[0] * attn_score + gamma[1] * max_imp +
                   gamma[2] * avg_ppr + gamma[3] * gap)
            score = _sigmoid(5.0 * raw)  # scale for sigmoid sensitivity

            results.append({
                "chunk": chunk,
                "score": float(score),
                "uncertainty": float(sigma),
                "debug": {
                    "attn_score": round(attn_score, 4),
                    "max_imp": round(max_imp, 4),
                    "avg_ppr": round(avg_ppr, 4),
                    "gap": round(gap, 4),
                    "sigma": round(sigma, 4),
                    "n_cells": len(ci),
                    "dominant_scale": dominant_scale,
                },
            })

        # ── 6. MMR with uncertainty + sheet diversity ────────────────
        results.sort(key=lambda x: x["score"] * (1 - x["uncertainty"]), reverse=True)

        selected: List[Dict] = []
        selected_sheets: set = set()
        lambda_mmr = 0.7
        sheet_diversity_bonus = 0.15
        for r in results:
            if len(selected) >= k:
                break
            if not selected:
                selected.append(r)
                sh = r["chunk"].get("metadata", {}).get("sheet", "")
                if sh:
                    selected_sheets.add(sh)
                continue

            # Sheet diversity bonus: new sheets get boosted
            chunk_sheet = r["chunk"].get("metadata", {}).get("sheet", "")
            is_new_sheet = chunk_sheet and chunk_sheet not in selected_sheets

            # Compute max similarity to already selected
            addrs = r["chunk"].get("metadata", {}).get("cell_addresses", [])
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

            eff_score = r["score"] * (1 - r["uncertainty"])
            if is_new_sheet:
                eff_score += sheet_diversity_bonus
            mmr = lambda_mmr * eff_score - (1 - lambda_mmr) * max_sim
            if mmr > -0.1:
                selected.append(r)
                if chunk_sheet:
                    selected_sheets.add(chunk_sheet)

        elapsed_ms = int((time.time() - t0) * 1000)
        logger.info(
            f"CGASR: Retrieved {len(selected)}/{len(chunks)} chunks in {elapsed_ms}ms | "
            f"top score={selected[0]['score']:.3f} σ={selected[0]['uncertainty']:.3f}" if selected else
            f"CGASR: No chunks selected in {elapsed_ms}ms"
        )

        return selected
