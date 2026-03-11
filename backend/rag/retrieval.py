import logging
from typing import Any, Dict, List

import networkx as nx
import numpy as np

from rag.local_embedder import LocalEmbedder
from rag.chroma_store import ChromaStore

logger = logging.getLogger(__name__)


class RAGRetriever:
    def __init__(
        self,
        embedder: LocalEmbedder,
        store: ChromaStore,
        ollama_client: Any,
        graph: nx.DiGraph,
    ) -> None:
        self.embedder = embedder
        self.store = store
        self.ollama = ollama_client
        self.graph = graph

    async def hyde_embed(self, query: str) -> List[float]:
        system = (
            "You are a financial model expert. Generate a 2-3 sentence hypothetical answer "
            "describing what the answer to the following question about an Excel financial model might look like. "
            "Be specific, reference cell addresses, formulas, and financial concepts."
        )
        try:
            hypothesis = await self.ollama.chat(
                messages=[{"role": "user", "content": query}],
                system=system,
                temperature=0.3,
            )
            return self.embedder.embed_single(hypothesis)
        except Exception as e:
            logger.warning(f"HyDE generation failed, falling back to direct query embedding: {e}")
            return self.embedder.embed_single(query)

    def graph_expand(self, chunks: List[Dict], n_hops: int = 1) -> List[Dict]:
        candidate_ids = {c["chunk_id"] for c in chunks}
        extra_cell_addresses: set = set()

        for chunk in chunks:
            cell_addresses = chunk.get("metadata", {}).get("cell_addresses", [])
            if isinstance(cell_addresses, str):
                cell_addresses = cell_addresses.split(",") if cell_addresses else []

            for addr in cell_addresses:
                if addr in self.graph:
                    for _ in range(n_hops):
                        for pred in self.graph.predecessors(addr):
                            extra_cell_addresses.add(pred)
                        for succ in self.graph.successors(addr):
                            extra_cell_addresses.add(succ)

        additional = []
        seen_ids = set(candidate_ids)
        for chunk in chunks:
            for addr in extra_cell_addresses:
                meta = chunk.get("metadata", {})
                chunk_addrs = meta.get("cell_addresses", [])
                if isinstance(chunk_addrs, str):
                    chunk_addrs = chunk_addrs.split(",") if chunk_addrs else []
                if addr in chunk_addrs and chunk["chunk_id"] not in seen_ids:
                    additional.append(chunk)
                    seen_ids.add(chunk["chunk_id"])

        return chunks + additional

    def mmr_select(
        self,
        query_emb: List[float],
        candidates: List[Dict],
        k: int = 5,
        lambda_: float = 0.6,
    ) -> List[Dict]:
        if not candidates:
            return []

        def get_embedding(chunk: Dict) -> np.ndarray:
            emb = chunk.get("embedding")
            if emb is not None:
                return np.array(emb, dtype=float)
            return np.zeros(len(query_emb), dtype=float)

        q = np.array(query_emb, dtype=float)
        selected: List[Dict] = []
        remaining = list(candidates)

        while len(selected) < k and remaining:
            best_chunk = None
            best_score = float("-inf")

            for chunk in remaining:
                v = get_embedding(chunk)
                norm_v = np.linalg.norm(v)
                norm_q = np.linalg.norm(q)
                relevance = float(np.dot(q, v) / (norm_q * norm_v + 1e-10))

                if selected:
                    max_sim = max(
                        float(
                            np.dot(get_embedding(s), v)
                            / (np.linalg.norm(get_embedding(s)) * norm_v + 1e-10)
                        )
                        for s in selected
                    )
                else:
                    max_sim = 0.0

                score = lambda_ * relevance - (1 - lambda_) * max_sim
                if score > best_score:
                    best_score = score
                    best_chunk = chunk

            if best_chunk is not None:
                selected.append(best_chunk)
                remaining.remove(best_chunk)

        return selected

    async def retrieve(self, query: str, workbook_uuid: str, k: int = 5) -> List[Dict]:
        try:
            query_emb = await self.hyde_embed(query)
        except Exception as e:
            logger.warning(f"Embedding failed: {e}")
            return []

        try:
            raw_candidates = self.store.query(workbook_uuid, query_emb, n=20)
        except Exception as e:
            logger.warning(f"Chroma query failed: {e}")
            return []

        expanded = self.graph_expand(raw_candidates, n_hops=1)

        for chunk in expanded:
            text = chunk.get("text", "")
            if text:
                chunk["embedding"] = self.embedder.embed_single(text)
            else:
                chunk["embedding"] = query_emb

        selected = self.mmr_select(query_emb, expanded, k=k)
        return selected

    def build_context(self, chunks: List[Dict], query: str, max_chars: int = 3000) -> str:
        parts = [f"QUESTION: {query}\n\nCONTEXT FROM WORKBOOK:\n"]
        total = len(parts[0])
        for i, chunk in enumerate(chunks):
            text = chunk.get("text", "")
            addition = f"\n--- CHUNK {i+1} ---\n{text}\n"
            if total + len(addition) > max_chars:
                break
            parts.append(addition)
            total += len(addition)
        return "".join(parts)
