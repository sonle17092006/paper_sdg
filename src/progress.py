"""Ghi tiến trình pipeline ra data/progress.json để lần sau biết đã chạy tới đâu."""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from config import PROGRESS_JSON


def _now() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def load_progress() -> dict[str, Any]:
    if not PROGRESS_JSON.exists():
        return {"created_at": _now(), "steps": {}}
    with PROGRESS_JSON.open(encoding="utf-8") as f:
        return json.load(f)


def save_progress(data: dict[str, Any]) -> None:
    PROGRESS_JSON.parent.mkdir(parents=True, exist_ok=True)
    data["updated_at"] = _now()
    with PROGRESS_JSON.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def mark_step(name: str, **payload: Any) -> dict[str, Any]:
    data = load_progress()
    steps = data.setdefault("steps", {})
    rec = {"status": payload.pop("status", "done"), "at": _now()}
    rec.update(payload)
    steps[name] = rec
    save_progress(data)
    return data
