from __future__ import annotations

from copy import deepcopy
from datetime import datetime
from typing import Any, Dict
from uuid import uuid4


class SessionStore:
    def __init__(self) -> None:
        self._items: Dict[str, Dict[str, Any]] = {}

    def save(self, result: Dict[str, Any]) -> Dict[str, Any]:
        session_id = uuid4().hex
        payload = deepcopy(result)
        payload["session_id"] = session_id
        payload["created_at"] = datetime.now().isoformat(timespec="seconds")
        self._items[session_id] = payload
        return deepcopy(payload)

    def get(self, session_id: str) -> Dict[str, Any] | None:
        item = self._items.get(session_id)
        return deepcopy(item) if item else None

    def latest(self, scene: str | None = None) -> Dict[str, Any] | None:
        values = list(self._items.values())
        if scene:
            values = [item for item in values if item.get("scene") == scene]
        if not values:
            return None
        values.sort(key=lambda item: item.get("created_at", ""), reverse=True)
        return deepcopy(values[0])


store = SessionStore()
