import logging
import re
from typing import Any, Dict, List, Optional, Tuple

import networkx as nx
import numpy as np
from rank_bm25 import BM25Okapi

from rag.local_embedder import LocalEmbedder
from rag.chroma_store import ChromaStore

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_SHEET_NAME_RE = re.compile(
    r"(?:in|w|na|z|from|sheet|arkusz|arkuszu)\s+['\"]?([A-Za-z][A-Za-z0-9 &_-]+)['\"]?",
    re.IGNORECASE,
)

_DIRECT_LOOKUP_RE = re.compile(
    r"(?:what is|jaki jest|jaka jest|ile wynosi|podaj|read|odczytaj|value of)\s+(.+)",
    re.IGNORECASE,
)


def _extract_sheet_hint(query: str) -> Optional[str]:
    """Try to extract a sheet name mentioned in the user query."""
    m = _SHEET_NAME_RE.search(query)
    return m.group(1).strip().rstrip("?.,!") if m else None


def _tokenize(text: str) -> List[str]:
    """Simple whitespace + punctuation tokenizer for BM25."""
    return re.findall(r"[A-Za-z0-9]+", text.lower())


def _rrf_fuse(
    ranked_lists: List[List[Tuple[str, float]]],
    k: int = 60,
) -> List[Tuple[str, float]]:
    """Reciprocal Rank Fusion across multiple ranked lists."""
    scores: Dict[str, float] = {}
    for ranked in ranked_lists:
        for rank, (doc_id, _score) in enumerate(ranked):
            scores[doc_id] = scores.get(doc_id, 0.0) + 1.0 / (k + rank + 1)
    return sorted(scores.items(), key=lambda x: x[1], reverse=True)


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
        self._bm25: Optional[BM25Okapi] = None
        self._bm25_chunks: List[Dict] = []
        self._cross_encoder = None

    # ------------------------------------------------------------------
    # BM25 index (built once per workbook, rebuilt on delta updates)
    # ------------------------------------------------------------------

    def build_bm25_index(self, workbook_uuid: str) -> None:
        """Build a BM25 index over all chunks for a workbook."""
        try:
            collection = self.store.get_or_create_collection(workbook_uuid)
            all_data = collection.get(include=["documents", "metadatas"])
            if not all_data["ids"]:
                self._bm25 = None
                self._bm25_chunks = []
                return

            chunks = []
            corpus = []
            for i, doc_id in enumerate(all_data["ids"]):
                text = all_data["documents"][i] if all_data["documents"] else ""
                meta = dict(all_data["metadatas"][i]) if all_data["metadatas"] else {}
                if "cell_addresses" in meta and isinstance(meta["cell_addresses"], str):
                    meta["cell_addresses"] = meta["cell_addresses"].split(",") if meta["cell_addresses"] else []
                chunks.append({"chunk_id": doc_id, "text": text, "metadata": meta})
                corpus.append(_tokenize(text))

            self._bm25 = BM25Okapi(corpus)
            self._bm25_chunks = chunks
            logger.info(f"BM25 index built: {len(chunks)} chunks for {workbook_uuid}")
        except Exception as e:
            logger.warning(f"BM25 index build failed: {e}")
            self._bm25 = None
            self._bm25_chunks = []

    def _bm25_search(self, query: str, n: int = 20) -> List[Tuple[str, float]]:
        """Keyword search via BM25. Returns (chunk_id, score) pairs."""
        if self._bm25 is None or not self._bm25_chunks:
            return []
        tokens = _tokenize(query)
        if not tokens:
            return []
        scores = self._bm25.get_scores(tokens)
        top_indices = np.argsort(scores)[::-1][:n]
        return [
            (self._bm25_chunks[i]["chunk_id"], float(scores[i]))
            for i in top_indices
            if scores[i] > 0
        ]

    # ------------------------------------------------------------------
    # Direct cell lookup (Informer) — bypasses RAG for simple queries
    # ------------------------------------------------------------------

    def direct_lookup(self, query: str, workbook_data: Any) -> Optional[str]:
        """For simple value queries, look up directly without LLM RAG.

        Returns a context string if a direct match was found, None otherwise.
        """
        if workbook_data is None:
            return None

        m = _DIRECT_LOOKUP_RE.match(query.strip())
        if not m:
            return None

        search_term = m.group(1).strip().lower().rstrip("?.,!")
        if len(search_term) < 2:
            return None

        from openpyxl.utils import get_column_letter

        matches: List[str] = []
        for cell_addr, cell in workbook_data.cells.items():
            if cell.value is None:
                continue
            label = str(cell.value).strip()
            label_lower = label.lower()
            # Require exact or near-exact match (not substring of long text)
            if label_lower == search_term or (
                len(label_lower) < 40 and search_term in label_lower
            ):
                sheet = cell.sheet_name
                row = cell.row
                col = cell.col
                for vc in range(col + 1, col + 4):
                    val_addr = f"{sheet}!{get_column_letter(vc)}{row}"
                    val_cell = workbook_data.cells.get(val_addr)
                    if val_cell and val_cell.value is not None:
                        matches.append(
                            f"{cell_addr} = '{label}' → {val_addr} = {val_cell.value}"
                        )
                        break
                else:
                    matches.append(f"{cell_addr} = '{label}'")

        # Only use direct lookup for small, precise result sets
        if not matches or len(matches) > 5:
            return None

        context = "DIRECT LOOKUP RESULTS:\n" + "\n".join(matches[:10])
        logger.info(f"Direct lookup matched {len(matches)} cells for '{search_term}'")
        return context

    # ------------------------------------------------------------------
    # Cross-encoder reranking
    # ------------------------------------------------------------------

    def _load_cross_encoder(self):
        """Lazy-load cross-encoder for reranking."""
        if self._cross_encoder is not None:
            return
        try:
            from sentence_transformers import CrossEncoder
            self._cross_encoder = CrossEncoder(
                "cross-encoder/ms-marco-MiniLM-L-6-v2",
                max_length=512,
            )
            logger.info("Cross-encoder loaded: ms-marco-MiniLM-L-6-v2")
        except Exception as e:
            logger.warning(f"Cross-encoder load failed (reranking disabled): {e}")
            self._cross_encoder = False

    def _cross_encoder_rerank(
        self, query: str, chunks: List[Dict], top_k: int = 5
    ) -> List[Dict]:
        """Re-rank chunks using cross-encoder for better precision."""
        self._load_cross_encoder()
        if not self._cross_encoder or self._cross_encoder is False:
            return chunks[:top_k]

        pairs = [(query, c.get("text", "")[:400]) for c in chunks]
        try:
            scores = self._cross_encoder.predict(pairs)
            scored = list(zip(chunks, scores))
            scored.sort(key=lambda x: x[1], reverse=True)
            return [c for c, _s in scored[:top_k]]
        except Exception as e:
            logger.warning(f"Cross-encoder rerank failed: {e}")
            return chunks[:top_k]

    # ------------------------------------------------------------------
    # Metadata filtering
    # ------------------------------------------------------------------

    def _filter_by_sheet(
        self, chunks: List[Dict], sheet_hint: Optional[str]
    ) -> List[Dict]:
        """Boost chunks matching the sheet hint. Don't discard others."""
        if not sheet_hint:
            return chunks

        hint_lower = sheet_hint.lower()
        boosted = []
        rest = []
        for c in chunks:
            sheet = c.get("metadata", {}).get("sheet", "")
            if sheet and sheet.lower() == hint_lower:
                boosted.append(c)
            else:
                rest.append(c)

        if boosted:
            logger.info(f"Sheet filter: boosted {len(boosted)} chunks for '{sheet_hint}'")
        return boosted + rest

    # ------------------------------------------------------------------
    # HyDE embedding (kept from original)
    # ------------------------------------------------------------------

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

    # ------------------------------------------------------------------
    # Graph expansion (kept from original)
    # ------------------------------------------------------------------

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

    # ------------------------------------------------------------------
    # MMR selection (kept from original)
    # ------------------------------------------------------------------

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

    # ------------------------------------------------------------------
    # Main retrieve pipeline (upgraded)
    # ------------------------------------------------------------------

    async def retrieve(
        self,
        query: str,
        workbook_uuid: str,
        k: int = 5,
        workbook_data: Any = None,
    ) -> List[Dict]:
        """Enhanced retrieval pipeline:
        1. Direct lookup (Informer) for simple queries
        2. HyDE embedding
        3. Vector search (ChromaDB)
        4. BM25 keyword search
        5. RRF fusion
        6. Sheet metadata boost
        7. Graph expansion
        8. Cross-encoder reranking
        """

        # --- Step 1: Direct lookup for simple queries ---
        direct_ctx = self.direct_lookup(query, workbook_data)
        if direct_ctx:
            return [{"chunk_id": "direct_lookup", "text": direct_ctx, "metadata": {}}]

        # --- Step 2: HyDE embedding ---
        try:
            query_emb = await self.hyde_embed(query)
        except Exception as e:
            logger.warning(f"Embedding failed: {e}")
            return []

        # --- Step 3: Vector search ---
        try:
            vector_candidates = self.store.query(workbook_uuid, query_emb, n=20)
        except Exception as e:
            logger.warning(f"Chroma query failed: {e}")
            vector_candidates = []

        # --- Step 4: BM25 keyword search ---
        bm25_results = self._bm25_search(query, n=20)

        # --- Step 5: RRF fusion ---
        vector_ranked = [
            (c["chunk_id"], 1.0 - c.get("distance", 0.0))
            for c in vector_candidates
        ]
        fused_ranking = _rrf_fuse([vector_ranked, bm25_results], k=60)

        chunk_map: Dict[str, Dict] = {}
        for c in vector_candidates:
            chunk_map[c["chunk_id"]] = c
        for c in self._bm25_chunks:
            if c["chunk_id"] not in chunk_map:
                chunk_map[c["chunk_id"]] = c

        fused_chunks = []
        for doc_id, rrf_score in fused_ranking[:30]:
            if doc_id in chunk_map:
                chunk = dict(chunk_map[doc_id])
                chunk["rrf_score"] = rrf_score
                fused_chunks.append(chunk)

        if not fused_chunks:
            fused_chunks = vector_candidates

        logger.info(
            f"Hybrid search: {len(vector_candidates)} vector + {len(bm25_results)} BM25 → {len(fused_chunks)} fused"
        )

        # --- Step 6: Sheet metadata boost ---
        sheet_hint = _extract_sheet_hint(query)
        fused_chunks = self._filter_by_sheet(fused_chunks, sheet_hint)

        # --- Step 7: Graph expansion ---
        expanded = self.graph_expand(fused_chunks, n_hops=1)

        # --- Step 8: Cross-encoder reranking ---
        reranked = self._cross_encoder_rerank(query, expanded, top_k=k)

        return reranked

    # ------------------------------------------------------------------
    # Context builder (kept from original)
    # ------------------------------------------------------------------

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
