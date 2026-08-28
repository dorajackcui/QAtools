"""Persistent aliases used to recognize source and target Excel headers."""

from __future__ import annotations

import json
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable


SOURCE_HEADER = "source"
TARGET_HEADER = "target"


def _clean_aliases(values: Iterable[object]) -> tuple[str, ...]:
    aliases: list[str] = []
    seen: set[str] = set()
    for value in values:
        if not isinstance(value, str):
            raise ValueError("表头别名必须是字符串。")
        alias = value.strip()
        normalized = alias.casefold()
        if not normalized or normalized in seen:
            continue
        seen.add(normalized)
        aliases.append(alias)
    return tuple(aliases)


@dataclass(frozen=True)
class HeaderAliases:
    """Custom header labels; the built-in source/target labels are implicit."""

    source: tuple[str, ...] = ()
    target: tuple[str, ...] = ()

    @classmethod
    def create(
        cls,
        *,
        source: Iterable[object] = (),
        target: Iterable[object] = (),
    ) -> HeaderAliases:
        source_aliases = _clean_aliases(source)
        target_aliases = _clean_aliases(target)
        source_headers = {SOURCE_HEADER, *(alias.casefold() for alias in source_aliases)}
        target_headers = {TARGET_HEADER, *(alias.casefold() for alias in target_aliases)}
        overlapping = source_headers & target_headers
        if overlapping:
            names = "、".join(sorted(overlapping))
            raise ValueError(f"Source 和 Target 表头别名不能重复：{names}")

        return cls(
            source=tuple(
                alias for alias in source_aliases if alias.casefold() != SOURCE_HEADER
            ),
            target=tuple(
                alias for alias in target_aliases if alias.casefold() != TARGET_HEADER
            ),
        )

    @classmethod
    def from_dict(cls, payload: Any) -> HeaderAliases:
        if not isinstance(payload, dict):
            raise ValueError("表头别名配置需要是 JSON 对象。")
        source = payload.get("source", [])
        target = payload.get("target", [])
        if not isinstance(source, list) or not isinstance(target, list):
            raise ValueError("表头别名配置中的 source 和 target 需要是数组。")
        return cls.create(source=source, target=target)

    @property
    def source_headers(self) -> frozenset[str]:
        return frozenset((SOURCE_HEADER, *(alias.casefold() for alias in self.source)))

    @property
    def target_headers(self) -> frozenset[str]:
        return frozenset((TARGET_HEADER, *(alias.casefold() for alias in self.target)))


def default_header_aliases_path() -> Path:
    if sys.platform == "darwin":
        return (
            Path.home()
            / "Library"
            / "Application Support"
            / "Toolshub"
            / "header_aliases.json"
        )
    if os.name == "nt":
        app_data = os.environ.get("APPDATA")
        base_path = Path(app_data) if app_data else Path.home() / "AppData" / "Roaming"
        return base_path / "Toolshub" / "header_aliases.json"
    config_home = os.environ.get("XDG_CONFIG_HOME")
    base_path = Path(config_home) if config_home else Path.home() / ".config"
    return base_path / "toolshub" / "header_aliases.json"


class HeaderAliasStore:
    def __init__(self, config_path: str | os.PathLike[str] | None = None) -> None:
        self.config_path = (
            Path(config_path).expanduser()
            if config_path is not None
            else default_header_aliases_path()
        )

    def load(self) -> HeaderAliases:
        if not self.config_path.is_file():
            return HeaderAliases()
        try:
            payload = json.loads(self.config_path.read_text(encoding="utf-8"))
        except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ValueError(f"无法读取表头别名配置：{exc}") from exc
        try:
            aliases_payload = payload["header_aliases"]
        except (KeyError, TypeError) as exc:
            raise ValueError("表头别名配置缺少 header_aliases 对象。") from exc
        return HeaderAliases.from_dict(aliases_payload)

    def save(self, aliases: HeaderAliases) -> None:
        validated = HeaderAliases.create(source=aliases.source, target=aliases.target)
        payload = {
            "version": 1,
            "header_aliases": {
                "source": list(validated.source),
                "target": list(validated.target),
            },
        }
        self.config_path.parent.mkdir(parents=True, exist_ok=True)
        temporary_path = self.config_path.with_suffix(f"{self.config_path.suffix}.tmp")
        try:
            temporary_path.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
            os.replace(temporary_path, self.config_path)
        except OSError:
            try:
                temporary_path.unlink()
            except FileNotFoundError:
                pass
            raise


__all__ = [
    "HeaderAliases",
    "HeaderAliasStore",
    "SOURCE_HEADER",
    "TARGET_HEADER",
    "default_header_aliases_path",
]
