from __future__ import annotations

import hashlib
import json
import tomllib
from dataclasses import dataclass
from functools import lru_cache
from importlib import resources
from pathlib import Path
from typing import Any

from .errors import ConfigError

__all__ = [
    "TagRules",
    "default_tag_rules",
    "load_tag_rules",
    "normalized_tag_rules_hash",
]


@dataclass(frozen=True)
class TagRules:
    version: int
    angle_allowed: frozenset[str]
    bbcode_allowed: frozenset[str]
    protect_raw_braces: bool
    source: str = "default"

    def __post_init__(self) -> None:
        object.__setattr__(
            self,
            "angle_allowed",
            frozenset(name.lower() for name in self.angle_allowed),
        )
        object.__setattr__(
            self,
            "bbcode_allowed",
            frozenset(name.lower() for name in self.bbcode_allowed),
        )

    def allows_angle(self, name: str) -> bool:
        return name.lower() in self.angle_allowed

    def allows_bbcode(self, name: str) -> bool:
        return name.lower() in self.bbcode_allowed


def default_tag_rules() -> TagRules:
    return _default_tag_rules()


@lru_cache(maxsize=1)
def _default_tag_rules() -> TagRules:
    path = resources.files("phraseloom").joinpath("tag_rules.toml")
    with path.open("rb") as handle:
        data = tomllib.load(handle)
    return _parse_tag_rules(data, source="default")


def load_tag_rules(path: str | Path | None = None) -> TagRules:
    if path is None:
        return default_tag_rules()

    tag_rules_path = Path(path)
    try:
        data = tomllib.loads(tag_rules_path.read_text(encoding="utf-8"))
    except tomllib.TOMLDecodeError as exc:
        raise ConfigError(f"{tag_rules_path}: invalid TOML: {exc}") from exc

    return _parse_tag_rules(data, source=str(tag_rules_path))


def normalized_tag_rules_hash(rules: TagRules) -> str:
    normalized = {
        "version": rules.version,
        "angle_allowed": sorted(name.lower() for name in rules.angle_allowed),
        "bbcode_allowed": sorted(name.lower() for name in rules.bbcode_allowed),
        "protect_raw_braces": rules.protect_raw_braces,
    }
    payload = json.dumps(normalized, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _parse_tag_rules(data: dict[str, Any], *, source: str) -> TagRules:
    if data.get("version") != 1:
        raise ConfigError("tag rules version must be exactly 1")

    angle_allowed = _read_allowlist_section(data, "angle_tags")
    bbcode_allowed = _read_allowlist_section(data, "bbcode_tags")

    raw_braces = data.get("raw_braces")
    if not isinstance(raw_braces, dict):
        raise ConfigError("raw_braces section is required")

    protect_raw_braces = raw_braces.get("protect_all")
    if not isinstance(protect_raw_braces, bool):
        raise ConfigError("raw_braces.protect_all must be a bool")

    return TagRules(
        version=1,
        angle_allowed=frozenset(angle_allowed),
        bbcode_allowed=frozenset(bbcode_allowed),
        protect_raw_braces=protect_raw_braces,
        source=source,
    )


def _read_allowlist_section(data: dict[str, Any], section_name: str) -> frozenset[str]:
    section = data.get(section_name)
    if not isinstance(section, dict):
        raise ConfigError(f"{section_name} section is required")

    if section.get("mode") != "allowlist":
        raise ConfigError(f"{section_name}.mode must be 'allowlist'")

    allowed = section.get("allowed")
    if not isinstance(allowed, list):
        raise ConfigError(f"{section_name}.allowed must be a list of non-empty strings")

    normalized: set[str] = set()
    for item in allowed:
        if not isinstance(item, str) or not item.strip():
            raise ConfigError(
                f"{section_name}.allowed must be a list of non-empty strings"
            )
        normalized.add(item.strip().lower())

    return frozenset(normalized)
