import logging
import os
import random
from typing import List

import numpy as np

logger = logging.getLogger(__name__)


class LocalEmbedder:
    def __init__(self, models_dir: str, model_name: str) -> None:
        self.models_dir = models_dir
        self.model_name = model_name
        self._model = None

    def load(self) -> None:
        from sentence_transformers import SentenceTransformer
        cache_dir = os.path.abspath(self.models_dir)
        os.makedirs(cache_dir, exist_ok=True)
        logger.info(f"Loading embedding model: {self.model_name} from {cache_dir}")
        self._model = SentenceTransformer(self.model_name, cache_folder=cache_dir)
        logger.info(f"Embedding model loaded. Dimension: {self.dimension}")

    def embed(self, texts: List[str]) -> List[List[float]]:
        if self._model is None:
            raise RuntimeError("Embedder not loaded. Call load() first.")
        embeddings = self._model.encode(
            texts,
            normalize_embeddings=True,
            show_progress_bar=False,
            convert_to_numpy=True,
        )
        return embeddings.tolist()

    def embed_single(self, text: str) -> List[float]:
        return self.embed([text])[0]

    def similarity(self, a: List[float], b: List[float]) -> float:
        va = np.array(a, dtype=float)
        vb = np.array(b, dtype=float)
        norm_a = np.linalg.norm(va)
        norm_b = np.linalg.norm(vb)
        if norm_a == 0 or norm_b == 0:
            return 0.0
        return float(np.dot(va, vb) / (norm_a * norm_b))

    @property
    def dimension(self) -> int:
        if self._model is None:
            return 768
        return self._model.get_sentence_embedding_dimension()


class LSHIndex:
    def __init__(self, n_planes: int = 10, dim: int = 768) -> None:
        self.n_planes = n_planes
        self.dim = dim
        self._hyperplanes = np.random.randn(n_planes, dim)
        self._index: dict[str, str] = {}

    def _hash(self, vector: List[float]) -> str:
        v = np.array(vector, dtype=float)
        projections = self._hyperplanes @ v
        return "".join("1" if p >= 0 else "0" for p in projections)

    def add(self, chunk_id: str, vector: List[float]) -> None:
        h = self._hash(vector)
        self._index[chunk_id] = h

    def find_candidates(self, vector: List[float], hamming_threshold: int = 2) -> List[str]:
        query_hash = self._hash(vector)
        results = []
        for chunk_id, stored_hash in self._index.items():
            dist = sum(a != b for a, b in zip(query_hash, stored_hash))
            if dist <= hamming_threshold:
                results.append(chunk_id)
        return results

    def clear(self) -> None:
        self._index.clear()
