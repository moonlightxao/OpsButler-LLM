import json
import re
import time
import logging
import requests
from opsbutler.config import LLMConfig

logger = logging.getLogger(__name__)


class LLMClient:
    """Base LLM client interface."""

    def chat(self, messages: list[dict]) -> str:
        """Send messages and return text response."""
        raise NotImplementedError

    def chat_json(self, messages: list[dict], json_retry: int | None = None) -> dict | list:
        """Send messages and parse JSON from response.

        Retries on JSON extraction failure with exponential backoff.
        Uses self.json_retry_count if json_retry not specified.
        """
        if json_retry is None:
            json_retry = getattr(self, 'json_retry_count', 2)

        last_error = None
        for attempt in range(json_retry + 1):
            try:
                response_text = self.chat(messages)
                return extract_json(response_text)
            except ValueError as e:
                last_error = e
                logger.warning(
                    "JSON extraction attempt %d/%d failed: %s",
                    attempt + 1, json_retry + 1, e
                )
                if attempt < json_retry:
                    time.sleep(2 ** attempt)

        raise ValueError(
            f"JSON extraction failed after {json_retry + 1} attempts: {last_error}"
        )


class OpenAICompatibleClient(LLMClient):
    """Client for OpenAI-compatible APIs (OpenAI, DeepSeek, Azure, etc.)."""

    def __init__(self, config: LLMConfig):
        self.base_url = config.base_url.rstrip("/")
        self.api_key = config.api_key
        self.model = config.model
        self.temperature = config.temperature
        self.max_tokens = config.max_tokens
        self.retry_count = config.retry_count
        self.debug = config.debug
        self.timeout = config.timeout
        self.json_retry_count = config.json_retry_count

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

        if self.debug:
            logger.info("===== [DEBUG] LLM Request =====")
            logger.info("  Model: %s", self.model)
            for msg in messages:
                role = msg.get("role", "unknown")
                content = msg.get("content", "")
                logger.info("  [%s]: %s", role, content[:5000])
            logger.info("===== End Prompt =====")

        last_error = None
        for attempt in range(self.retry_count + 1):
            try:
                start_time = time.time()
                resp = requests.post(url, headers=headers, json=payload, timeout=self.timeout)
                resp.raise_for_status()
                data = resp.json()
                duration = time.time() - start_time

                if self.debug:
                    usage = data.get("usage", {})
                    response_content = data["choices"][0]["message"]["content"]
                    logger.info("===== [DEBUG] LLM Response =====")
                    logger.info("  Duration: %.2fs", duration)
                    logger.info("  Prompt tokens: %s", usage.get("prompt_tokens", "N/A"))
                    logger.info("  Completion tokens: %s", usage.get("completion_tokens", "N/A"))
                    logger.info("  Total tokens: %s", usage.get("total_tokens", "N/A"))
                    logger.info("  Response: %s", response_content[:5000])
                    logger.info("===== End Response =====")

                return data["choices"][0]["message"]["content"]
            except Exception as e:
                last_error = e
                logger.warning(f"LLM request attempt {attempt + 1} failed: {e}")
                if attempt < self.retry_count:
                    time.sleep(2 ** attempt)

        raise RuntimeError(f"LLM request failed after {self.retry_count + 1} attempts: {last_error}")


class OllamaClient(LLMClient):
    """Client for Ollama local models."""

    def __init__(self, config: LLMConfig):
        self.host = config.base_url.rstrip("/") or "http://localhost:11434"
        self.model = config.model
        self.temperature = config.temperature
        self.retry_count = config.retry_count
        self.think = config.think
        self.debug = config.debug
        self.timeout = config.timeout
        self.json_retry_count = config.json_retry_count

    def chat(self, messages: list[dict]) -> str:
        url = f"{self.host}/api/chat"
        payload = {
            "model": self.model,
            "messages": messages,
            "stream": False,
            "options": {
                "temperature": self.temperature,
            },
            "think": self.think,
        }

        if self.debug:
            logger.info("===== [DEBUG] LLM Request =====")
            logger.info("  Model: %s", self.model)
            for msg in messages:
                role = msg.get("role", "unknown")
                content = msg.get("content", "")
                logger.info("  [%s]: %s", role, content[:5000])
            logger.info("===== End Prompt =====")

        last_error = None
        for attempt in range(self.retry_count + 1):
            try:
                start_time = time.time()
                resp = requests.post(url, json=payload, timeout=self.timeout)
                resp.raise_for_status()
                data = resp.json()
                duration = time.time() - start_time

                content = data["message"]["content"]

                if self.debug:
                    prompt_tokens = data.get("prompt_eval_count", "N/A")
                    completion_tokens = data.get("eval_count", "N/A")
                    logger.info("===== [DEBUG] LLM Response =====")
                    logger.info("  Duration: %.2fs", duration)
                    logger.info("  Prompt tokens: %s", prompt_tokens)
                    logger.info("  Completion tokens: %s", completion_tokens)
                    logger.info("  Response: %s", content[:5000])
                    logger.info("===== End Response =====")

                if not content or content.strip() in ("", "..."):
                    raise ValueError(
                        f"Ollama returned empty or placeholder content, "
                        f"response keys: {list(data.keys())}, "
                        f"message keys: {list(data['message'].keys())}"
                    )
                return content
            except Exception as e:
                last_error = e
                logger.warning(f"Ollama request attempt {attempt + 1} failed: {e}")
                if attempt < self.retry_count:
                    time.sleep(2 ** attempt)

        raise RuntimeError(f"Ollama request failed after {self.retry_count + 1} attempts: {last_error}")


def create_llm_client(config: LLMConfig) -> LLMClient:
    """Factory: create LLM client based on config.provider."""
    if config.provider == "ollama":
        return OllamaClient(config)
    return OpenAICompatibleClient(config)


def _fix_json_newlines(text: str) -> str:
    """Replace literal newlines inside JSON string values with escaped \\n.

    LLMs sometimes return JSON where string values contain literal newlines
    instead of escaped \\n, which is invalid JSON.
    """
    result = []
    in_string = False
    i = 0
    while i < len(text):
        char = text[i]
        # Handle escape sequences inside strings — skip the next char
        if char == '\\' and in_string:
            result.append(char)
            i += 1
            if i < len(text):
                result.append(text[i])
            i += 1
            continue
        if char == '"':
            in_string = not in_string
        elif char == '\n' and in_string:
            result.append('\\n')
            i += 1
            continue
        result.append(char)
        i += 1
    return ''.join(result)


def _try_parse_json(text: str) -> dict | list | None:
    """Try to parse text as JSON, with newline repair fallback."""
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    try:
        return json.loads(_fix_json_newlines(text))
    except json.JSONDecodeError:
        pass
    return None


def _find_all_json_candidates(text: str) -> list[tuple[int, object]]:
    """Find all valid, non-empty JSON objects/arrays in text.

    Returns list of (end_position, parsed_json) sorted by position.
    Skips empty {} and [].
    """
    candidates = []
    i = 0
    while i < len(text):
        if text[i] not in ('{', '['):
            i += 1
            continue

        open_char = text[i]
        close_char = '}' if open_char == '{' else ']'
        depth = 0
        in_string = False
        j = i
        while j < len(text):
            c = text[j]
            if c == '\\' and in_string:
                j += 2
                continue
            if c == '"':
                in_string = not in_string
            elif not in_string:
                if c == open_char:
                    depth += 1
                elif c == close_char:
                    depth -= 1
                    if depth == 0:
                        candidate = text[i:j+1]
                        result = _try_parse_json(candidate)
                        if result is not None:
                            is_empty = (isinstance(result, dict) and len(result) == 0) or \
                                       (isinstance(result, list) and len(result) == 0)
                            if not is_empty:
                                candidates.append((j, result))
                        i = j
                        break
            j += 1
        i += 1
    return candidates


def extract_json(text: str) -> dict | list:
    """Extract JSON object or array from LLM response text.

    Handles cases where LLM wraps JSON in markdown code blocks,
    adds extra text (including thinking/chain-of-thought), or
    returns empty content.

    Strategy: find all valid JSON candidates in the response and
    take the LAST non-empty one - since thinking/CoT comes before
    the actual answer in most reasoning models.
    """
    if not text:
        raise ValueError("LLM returned empty or None response")

    text = text.strip()

    # Strategy 1: direct parse
    result = _try_parse_json(text)
    if result is not None:
        is_empty = (isinstance(result, dict) and len(result) == 0) or \
                   (isinstance(result, list) and len(result) == 0)
        if not is_empty:
            return result

    # Strategy 2: extract from markdown code blocks (take last valid)
    last_md_result = None
    for match in re.finditer(r'```(?:json)?\s*\n?(.*?)\n?```', text, re.DOTALL):
        result = _try_parse_json(match.group(1).strip())
        if result is not None:
            is_empty = (isinstance(result, dict) and len(result) == 0) or \
                       (isinstance(result, list) and len(result) == 0)
            if not is_empty:
                last_md_result = result
    if last_md_result is not None:
        return last_md_result

    # Strategy 3: find ALL JSON candidates, take the last non-empty one
    candidates = _find_all_json_candidates(text)
    if candidates:
        _, result = candidates[-1]
        return result

    raise ValueError(f"Could not extract JSON from response: {text[:200]}...")
