import json
import logging
from typing import AsyncGenerator, List, Dict, Any, Optional

import aiohttp

logger = logging.getLogger(__name__)


class OllamaClient:
    def __init__(self, base_url: str, model: str, fallback_model: str = "mistral") -> None:
        self.base_url = base_url.rstrip("/")
        self.model = model
        self.fallback_model = fallback_model
        self._session: Optional[aiohttp.ClientSession] = None

    def _get_session(self) -> aiohttp.ClientSession:
        if self._session is None or self._session.closed:
            self._session = aiohttp.ClientSession()
        return self._session

    async def chat(
        self,
        messages: List[Dict[str, str]],
        system: Optional[str] = None,
        temperature: float = 0.1,
    ) -> str:
        payload: Dict[str, Any] = {
            "model": self.model,
            "messages": messages,
            "stream": False,
            "options": {"temperature": temperature},
        }
        if system:
            payload["system"] = system

        session = self._get_session()
        try:
            async with session.post(
                f"{self.base_url}/api/chat",
                json=payload,
                timeout=aiohttp.ClientTimeout(total=120),
            ) as resp:
                resp.raise_for_status()
                data = await resp.json()
                return data["message"]["content"]
        except Exception as e:
            logger.error(f"Ollama chat failed with model {self.model}: {e}")
            if self.fallback_model and self.fallback_model != self.model:
                logger.info(f"Retrying with fallback model {self.fallback_model}")
                payload["model"] = self.fallback_model
                async with session.post(
                    f"{self.base_url}/api/chat",
                    json=payload,
                    timeout=aiohttp.ClientTimeout(total=120),
                ) as resp:
                    resp.raise_for_status()
                    data = await resp.json()
                    return data["message"]["content"]
            raise

    async def stream_chat(
        self,
        messages: List[Dict[str, str]],
        system: Optional[str] = None,
    ) -> AsyncGenerator[str, None]:
        payload: Dict[str, Any] = {
            "model": self.model,
            "messages": messages,
            "stream": True,
            "options": {"temperature": 0.1},
        }
        if system:
            payload["system"] = system

        session = self._get_session()
        async with session.post(
            f"{self.base_url}/api/chat",
            json=payload,
            timeout=aiohttp.ClientTimeout(total=300),
        ) as resp:
            resp.raise_for_status()
            async for line in resp.content:
                line = line.strip()
                if not line:
                    continue
                try:
                    data = json.loads(line)
                    content = data.get("message", {}).get("content", "")
                    if content:
                        yield content
                    if data.get("done"):
                        break
                except json.JSONDecodeError as e:
                    logger.warning(f"Failed to parse NDJSON line: {e}")
                    continue

    async def is_running(self) -> bool:
        session = self._get_session()
        try:
            async with session.get(
                f"{self.base_url}/api/tags",
                timeout=aiohttp.ClientTimeout(total=3),
            ) as resp:
                return resp.status == 200
        except Exception:
            return False

    async def ensure_model(self, name: str) -> None:
        session = self._get_session()
        payload = {"name": name, "stream": True}
        async with session.post(
            f"{self.base_url}/api/pull",
            json=payload,
            timeout=aiohttp.ClientTimeout(total=600),
        ) as resp:
            resp.raise_for_status()
            async for line in resp.content:
                line = line.strip()
                if not line:
                    continue
                try:
                    data = json.loads(line)
                    status = data.get("status", "")
                    logger.info(f"Pull {name}: {status}")
                    if data.get("done") or status == "success":
                        break
                except json.JSONDecodeError:
                    continue

    async def close(self) -> None:
        if self._session and not self._session.closed:
            await self._session.close()
            self._session = None
