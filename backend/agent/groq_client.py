import json
import logging
from typing import AsyncGenerator, List, Dict, Any, Optional

import aiohttp

logger = logging.getLogger(__name__)

GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions"


class GroqClient:
    def __init__(self, api_key: str, model: str = "llama-3.3-70b-versatile") -> None:
        self.api_key = api_key
        self.model = model
        self._session: Optional[aiohttp.ClientSession] = None

    def _get_session(self) -> aiohttp.ClientSession:
        if self._session is None or self._session.closed:
            self._session = aiohttp.ClientSession()
        return self._session

    def _headers(self) -> Dict[str, str]:
        return {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }

    def _build_messages(
        self,
        messages: List[Dict[str, str]],
        system: Optional[str],
    ) -> List[Dict[str, str]]:
        out: List[Dict[str, str]] = []
        if system:
            out.append({"role": "system", "content": system})
        out.extend(messages)
        return out

    async def chat(
        self,
        messages: List[Dict[str, str]],
        system: Optional[str] = None,
        temperature: float = 0.1,
    ) -> str:
        payload: Dict[str, Any] = {
            "model": self.model,
            "messages": self._build_messages(messages, system),
            "temperature": temperature,
            "stream": False,
        }
        session = self._get_session()
        async with session.post(
            GROQ_API_URL,
            json=payload,
            headers=self._headers(),
            timeout=aiohttp.ClientTimeout(total=60),
        ) as resp:
            if not resp.ok:
                body = await resp.text()
                raise RuntimeError(f"Groq API error {resp.status}: {body}")
            data = await resp.json()
            return data["choices"][0]["message"]["content"]

    async def stream_chat(
        self,
        messages: List[Dict[str, str]],
        system: Optional[str] = None,
    ) -> AsyncGenerator[str, None]:
        payload: Dict[str, Any] = {
            "model": self.model,
            "messages": self._build_messages(messages, system),
            "temperature": 0.1,
            "stream": True,
        }
        session = self._get_session()
        async with session.post(
            GROQ_API_URL,
            json=payload,
            headers=self._headers(),
            timeout=aiohttp.ClientTimeout(total=120),
        ) as resp:
            if not resp.ok:
                body = await resp.text()
                raise RuntimeError(f"Groq API error {resp.status}: {body}")
            async for raw_line in resp.content:
                line = raw_line.decode("utf-8").strip()
                if not line or not line.startswith("data:"):
                    continue
                data_str = line[5:].strip()
                if data_str == "[DONE]":
                    break
                try:
                    chunk = json.loads(data_str)
                    delta = chunk["choices"][0].get("delta", {})
                    content = delta.get("content", "")
                    if content:
                        yield content
                except (json.JSONDecodeError, KeyError, IndexError):
                    continue

    async def is_running(self) -> bool:
        session = self._get_session()
        try:
            async with session.get(
                "https://api.groq.com/openai/v1/models",
                headers=self._headers(),
                timeout=aiohttp.ClientTimeout(total=5),
            ) as resp:
                return resp.status == 200
        except Exception:
            return False

    async def close(self) -> None:
        if self._session and not self._session.closed:
            await self._session.close()
            self._session = None
