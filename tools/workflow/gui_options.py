"""Remember the explicitly saved options of the unified QA page."""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

from tools.header_aliases import default_header_aliases_path


def _restore_values(saved: Any, defaults: Any) -> Any:
    """Use the page's defaults to validate types and fill newly added options."""
    if isinstance(defaults, dict) and isinstance(saved, dict):
        return {
            key: _restore_values(saved.get(key, value), value)
            for key, value in defaults.items()
        }
    if isinstance(defaults, tuple) and isinstance(saved, (list, tuple)):
        item_type = type(defaults[0]) if defaults else str
        if all(type(value) is item_type for value in saved):
            if item_type is not bool or len(saved) == len(defaults):
                return tuple(saved)
    elif type(saved) is type(defaults):
        if type(saved) is int and not 1 <= saved <= 1_000_000:
            raise ValueError("记住的开始行超出有效范围。")
        return saved
    raise ValueError("记住的选项格式不正确。")


class WorkflowOptionsStore:
    def __init__(self, config_path: str | os.PathLike[str] | None = None) -> None:
        self.config_path = (
            Path(config_path).expanduser()
            if config_path is not None
            else default_header_aliases_path().with_name("workflow_options.json")
        )

    def load(self, defaults: dict[str, Any]) -> dict[str, Any]:
        if not self.config_path.is_file():
            return defaults
        try:
            payload = json.loads(self.config_path.read_text(encoding="utf-8"))
        except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ValueError(f"无法读取记住的选项：{exc}") from exc
        if not isinstance(payload, dict) or payload.get("version") != 1:
            raise ValueError("记住的选项版本不支持。")
        return _restore_values(payload.get("options"), defaults)

    def save(self, options: dict[str, Any]) -> None:
        payload = {"version": 1, "options": options}
        self.config_path.parent.mkdir(parents=True, exist_ok=True)
        temporary_path = self.config_path.with_suffix(".json.tmp")
        try:
            temporary_path.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
            os.replace(temporary_path, self.config_path)
        except OSError:
            temporary_path.unlink(missing_ok=True)
            raise
