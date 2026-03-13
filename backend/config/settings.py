from functools import lru_cache
from typing import List
import logging
from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    database_url: str = "postgresql+asyncpg://postgres:postgres@localhost:5432/excelai"
    ollama_base_url: str = "http://localhost:11434"
    ollama_model: str = "llama3.1:8b"
    ollama_fallback_model: str = "mistral"
    embedding_model: str = "paraphrase-multilingual-mpnet-base-v2"
    models_dir: str = "./models"
    chroma_path: str = "./chroma_db"
    log_level: str = "INFO"
    cors_origins: List[str] = ["http://localhost:3000", "https://localhost:3000"]
    groq_api_key: str = ""
    groq_model: str = "llama-3.3-70b-versatile"

    model_config = {"env_file": ".env", "env_file_encoding": "utf-8"}

    def configure_logging(self) -> None:
        logging.basicConfig(
            level=getattr(logging, self.log_level.upper(), logging.INFO),
            format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        )


def get_settings() -> Settings:
    settings = Settings()
    settings.configure_logging()
    return settings
