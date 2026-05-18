from __future__ import annotations

import json
import os
import urllib.error
import urllib.request
from typing import Any, Dict, List


def llm_enabled(flag_name: str) -> bool:
    return os.getenv(flag_name, "false").strip().lower() in {"1", "true", "yes", "on"}


def _first_env(*names: str) -> str | None:
    for name in names:
        value = os.getenv(name)
        if value:
            return value.strip()
    return None


def _minimax_messages_url() -> str:
    base_url = _first_env("MINIMAX_BASE_URL", "ANTHROPIC_BASE_URL") or "https://api.minimaxi.com/anthropic"
    base_url = base_url.rstrip("/")
    if base_url.endswith("/v1/messages"):
        return base_url
    if base_url.endswith("/v1"):
        return f"{base_url}/messages"
    return f"{base_url}/v1/messages"


def _timeout_seconds() -> int:
    timeout_ms = _first_env("API_TIMEOUT_MS")
    if timeout_ms:
        try:
            return max(1, int(timeout_ms) // 1000)
        except ValueError:
            pass
    try:
        return max(1, int(os.getenv("LLM_TIMEOUT_SECONDS", "60")))
    except ValueError:
        return 60


def _max_tokens() -> int:
    try:
        return max(1, int(os.getenv("LLM_MAX_TOKENS", "2000")))
    except ValueError:
        return 2000


def _temperature() -> float:
    try:
        value = float(os.getenv("LLM_TEMPERATURE", "0.2"))
    except ValueError:
        return 0.2
    return min(max(value, 0.01), 1.0)


def _parse_json_object(text: str) -> Dict[str, Any] | None:
    stripped = text.strip()
    if stripped.startswith("```"):
        lines = stripped.splitlines()
        if len(lines) >= 3 and lines[-1].strip().startswith("```"):
            stripped = "\n".join(lines[1:-1]).strip()

    try:
        parsed = json.loads(stripped)
        return parsed if isinstance(parsed, dict) else None
    except json.JSONDecodeError:
        pass

    decoder = json.JSONDecoder()
    for idx, char in enumerate(stripped):
        if char != "{":
            continue
        try:
            parsed, _ = decoder.raw_decode(stripped[idx:])
        except json.JSONDecodeError:
            continue
        if isinstance(parsed, dict):
            return parsed
    return None


def call_llm_json(system_prompt: str, user_payload: Dict[str, Any]) -> Dict[str, Any] | None:
    """Small optional MiniMax client.

    The app can run fully without network or API keys. This helper is used only
    when ENABLE_LLM_MAPPING or ENABLE_LLM_QUERY is explicitly enabled.
    """
    api_key = _first_env("MINIMAX_API_KEY", "ANTHROPIC_AUTH_TOKEN", "ANTHROPIC_API_KEY")
    if not api_key:
        return None

    model = _first_env("MINIMAX_MODEL", "ANTHROPIC_MODEL", "LLM_MODEL") or "MiniMax-M2.7-highspeed"
    body = {
        "model": model,
        "system": (
            f"{system_prompt}\n\n"
            "你必须只输出一个合法 JSON 对象，不要输出 Markdown、解释、代码块或多余文本。"
        ),
        "messages": [
            {
                "role": "user",
                "content": json.dumps(user_payload, ensure_ascii=False),
            },
        ],
        "max_tokens": _max_tokens(),
        "temperature": _temperature(),
        "stream": False,
    }
    request = urllib.request.Request(
        _minimax_messages_url(),
        data=json.dumps(body).encode("utf-8"),
        headers={
            "X-Api-Key": api_key,
            "Content-Type": "application/json",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=_timeout_seconds()) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError):
        return None

    text_chunks: List[str] = []
    for content in payload.get("content", []):
        if content.get("type") == "text" and content.get("text"):
            text_chunks.append(content["text"])
    if not text_chunks:
        return None

    return _parse_json_object("\n".join(text_chunks))
