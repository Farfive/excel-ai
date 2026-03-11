import logging
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import chromadb
from chromadb.config import Settings as ChromaSettings

logger = logging.getLogger(__name__)


@dataclass
class ChunkRecord:
    chunk_id: str
    text: str
    embedding: List[float]
    metadata: Dict[str, Any] = field(default_factory=dict)


class ChromaStore:
    def __init__(self, chroma_path: str) -> None:
        self.chroma_path = chroma_path
        self._client: Optional[chromadb.PersistentClient] = None

    def _get_client(self) -> chromadb.PersistentClient:
        if self._client is None:
            self._client = chromadb.PersistentClient(
                path=self.chroma_path,
                settings=ChromaSettings(anonymized_telemetry=False),
            )
        return self._client

    def _collection_name(self, workbook_uuid: str) -> str:
        return f"wb_{workbook_uuid[:12]}"

    def get_or_create_collection(self, workbook_uuid: str):
        client = self._get_client()
        name = self._collection_name(workbook_uuid)
        return client.get_or_create_collection(
            name=name,
            metadata={"hnsw:space": "cosine"},
        )

    def upsert_chunks(self, workbook_uuid: str, chunks: List[ChunkRecord], batch: int = 100) -> None:
        collection = self.get_or_create_collection(workbook_uuid)
        for i in range(0, len(chunks), batch):
            batch_chunks = chunks[i:i + batch]
            ids = [c.chunk_id for c in batch_chunks]
            embeddings = [c.embedding for c in batch_chunks]
            documents = [c.text for c in batch_chunks]
            metadatas = []
            for c in batch_chunks:
                meta = dict(c.metadata)
                if "cell_addresses" in meta and isinstance(meta["cell_addresses"], list):
                    meta["cell_addresses"] = ",".join(meta["cell_addresses"])
                metadatas.append(meta)
            collection.upsert(ids=ids, embeddings=embeddings, documents=documents, metadatas=metadatas)
        logger.info(f"Upserted {len(chunks)} chunks for workbook {workbook_uuid}")

    def query(self, workbook_uuid: str, embedding: List[float], n: int = 20) -> List[Dict]:
        collection = self.get_or_create_collection(workbook_uuid)
        results = collection.query(
            query_embeddings=[embedding],
            n_results=min(n, collection.count() or 1),
            include=["documents", "metadatas", "distances"],
        )
        output = []
        if not results["ids"] or not results["ids"][0]:
            return output
        for idx in range(len(results["ids"][0])):
            meta = dict(results["metadatas"][0][idx]) if results["metadatas"] else {}
            if "cell_addresses" in meta and isinstance(meta["cell_addresses"], str):
                meta["cell_addresses"] = meta["cell_addresses"].split(",") if meta["cell_addresses"] else []
            output.append({
                "chunk_id": results["ids"][0][idx],
                "text": results["documents"][0][idx] if results["documents"] else "",
                "metadata": meta,
                "distance": results["distances"][0][idx] if results["distances"] else 0.0,
            })
        return output

    def delta_upsert(self, workbook_uuid: str, chunk_ids: List[str], new_chunks: List[ChunkRecord]) -> None:
        collection = self.get_or_create_collection(workbook_uuid)
        if chunk_ids:
            try:
                collection.delete(ids=chunk_ids)
            except Exception as e:
                logger.warning(f"Delta delete failed for some ids: {e}")
        self.upsert_chunks(workbook_uuid, new_chunks)

    def delete_workbook(self, workbook_uuid: str) -> None:
        client = self._get_client()
        name = self._collection_name(workbook_uuid)
        try:
            client.delete_collection(name)
            logger.info(f"Deleted collection {name}")
        except Exception as e:
            logger.warning(f"Could not delete collection {name}: {e}")

    def workbook_exists(self, workbook_uuid: str) -> bool:
        client = self._get_client()
        name = self._collection_name(workbook_uuid)
        try:
            client.get_collection(name)
            return True
        except Exception:
            return False

    def get_stats(self, workbook_uuid: str) -> Dict[str, Any]:
        collection = self.get_or_create_collection(workbook_uuid)
        count = collection.count()
        return {"collection_name": self._collection_name(workbook_uuid), "chunk_count": count}
