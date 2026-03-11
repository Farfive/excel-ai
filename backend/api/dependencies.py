from typing import Dict, Any
import networkx as nx

from rag.local_embedder import LocalEmbedder, LSHIndex
from rag.chroma_store import ChromaStore
from agent.ollama_client import OllamaClient
from agent.groq_client import GroqClient
from config.settings import get_settings
from parser.xlsx_parser import WorkbookData
from analysis.audit_trail import AuditTrail
from analysis.scenarios import ScenarioManager

settings = get_settings()

embedder = LocalEmbedder(
    models_dir=settings.models_dir,
    model_name=settings.embedding_model,
)

chroma = ChromaStore(chroma_path=settings.chroma_path)

if settings.groq_api_key:
    ollama = GroqClient(
        api_key=settings.groq_api_key,
        model=settings.groq_model,
    )
else:
    ollama = OllamaClient(
        base_url=settings.ollama_base_url,
        model=settings.ollama_model,
        fallback_model=settings.ollama_fallback_model,
    )

lsh_index = LSHIndex(n_planes=10, dim=768)

workbook_graphs: Dict[str, nx.DiGraph] = {}
workbook_data_cache: Dict[str, WorkbookData] = {}
workbook_states: Dict[str, Dict[str, Any]] = {}
audit_trails: Dict[str, AuditTrail] = {}
scenario_managers: Dict[str, ScenarioManager] = {}


def get_embedder() -> LocalEmbedder:
    return embedder


def get_chroma() -> ChromaStore:
    return chroma


def get_ollama() -> OllamaClient:
    return ollama


def get_lsh() -> LSHIndex:
    return lsh_index


def get_workbook_graphs() -> Dict[str, nx.DiGraph]:
    return workbook_graphs


def get_workbook_data_cache() -> Dict[str, WorkbookData]:
    return workbook_data_cache


def get_workbook_states() -> Dict[str, Dict[str, Any]]:
    return workbook_states


def get_audit_trails() -> Dict[str, AuditTrail]:
    return audit_trails


def get_scenario_managers() -> Dict[str, ScenarioManager]:
    return scenario_managers
