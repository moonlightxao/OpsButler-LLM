import json
import re
import logging
import requests
from opsbutler.config import LLMConfig

logger = logging.getLogger(__name__)


class LLMClient:
    """Base LLM client interface."""

    def chat(self, messages: list[dict]) -> str:
        """Send messages and return text response."""
        raise NotImplementedError

    def chat_json(self, messages: list[dict]) -> dict:
        """Send messages and parse JSON from response."""
        response_text = self.chat(messages)
        return extract_json(response_text)


class OpenAICompatibleClient(LLMClient):
    """Client for OpenAI-compatible APIs (OpenAI, DeepSeek, Azure, etc.)."""

    def __init__(self, config: LLMConfig):
        self.base_url = config.base_url.rstrip("/")
        self.api_key = config.api_key
        self.model = config.model
        self.temperature = config.temperature
        self.max_tokens = config.max_tokens
        self.retry_count = config.retry_count

    def chat(self, messages: list[dict]) -> str:
        url = f"{self.base_url}/chat/completions"
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}",
        }
        payload = {
            "model": self.model,
            "messages": messages,
            "temperature": self.temperature,
            "max_tokens": self.max_tokens,
        }

        last_error = None
        for attempt in range(self.retry_count + 1):
            try:
                resp = requests.post(url, headers=headers, json=payload, timeout=120)
                resp.raise_for_status()
                data = resp.json()
                return data["choices"][0]["message"]["content"]
            except Exception as e:
                last_error = e
                logger.warning(f"LLM request attempt {attempt + 1} failed: {e}")
                if attempt < self.retry_count:
                    import time
                    time.sleep(2 ** attempt)

        raise RuntimeError(f"LLM request failed after {self.retry_count + 1} attempts: {last_error}")


class OllamaClient(LLMClient):
    """Client for Ollama local models."""

    def __init__(self, config: LLMConfig):
        self.host = config.base_url.rstrip("/") or "http://localhost:11434"
        self.model = config.model
        self.temperature = config.temperature
        self.retry_count = config.retry_count

    def chat(self, messages: list[dict]) -> str:
        url = f"{self.host}/api/chat"
        payload = {
            "model": self.model,
            "messages": messages,
            "stream": False,
            "options": {
                "temperature": self.temperature,
            },
        }

        last_error = None
        for attempt in range(self.retry_count + 1):
            try:
                resp = requests.post(url, json=payload, timeout=300)
                resp.raise_for_status()
                data = resp.json()
                return data["message"]["content"]
            except Exception as e:
                last_error = e
                logger.warning(f"Ollama request attempt {attempt + 1} failed: {e}")
                if attempt < self.retry_count:
                    import time
                    time.sleep(2 ** attempt)

        raise RuntimeError(f"Ollama request failed after {self.retry_count + 1} attempts: {last_error}")


def create_llm_client(config: LLMConfig) -> LLMClient:
    """Factory: create LLM client based on config.provider."""
    if config.provider == "ollama":
        return OllamaClient(config)
    return OpenAICompatibleClient(config)


def extract_json(text: str) -> dict:
    """Extract JSON object from LLM response text.

    Handles cases where LLM wraps JSON in markdown code blocks
    or adds extra text around it.
    """
    # Try direct parse first
    text = text.strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Try to extract from markdown code block
    pattern = r'```(?:json)?\s*\n?(.*?)\n?```'
    match = re.search(pattern, text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(1).strip())
        except json.JSONDecodeError:
            pass

    # Try to find first { ... } block
    pattern = r'\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\}'
    match = re.search(pattern, text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group(0))
        except json.JSONDecodeError:
            pass

    # Last resort: find balanced braces
    start = text.find('{')
    if start != -1:
        depth = 0
        for i in range(start, len(text)):
            if text[i] == '{':
                depth += 1
            elif text[i] == '}':
                depth -= 1
                if depth == 0:
                    try:
                        return json.loads(text[start:i+1])
                    except json.JSONDecodeError:
                        continue

    raise ValueError(f"Could not extract JSON from response: {text[:200]}...")
