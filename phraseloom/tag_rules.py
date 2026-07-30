from __future__ import annotations

import hashlib
import json
import tomllib
from dataclasses import dataclass, field
from functools import lru_cache
from importlib import resources
from pathlib import Path
from typing import Any, Mapping

from .errors import ConfigError

__all__ = [
    "TagRules",
    "default_tag_rules",
    "load_tag_rules",
    "normalized_tag_rules_hash",
    "tag_rules_from_payload",
    "tag_rules_payload",
]


@dataclass(frozen=True)
class TagRules:
    version: int
    angle_allowed: frozenset[str]
    bbcode_allowed: frozenset[str]
    protect_raw_braces: bool
    source: str = "default"
    angle_aliases: Mapping[str, str] = field(default_factory=dict)
    angle_single: frozenset[str] = field(default_factory=frozenset)
    angle_optional_pair: frozenset[str] = field(default_factory=frozenset)

    def __post_init__(self) -> None:
        aliases = {
            alias.strip().lower(): target.strip().lower()
            for alias, target in self.angle_aliases.items()
            if alias.strip() and target.strip()
        }

        def canonicalize(name: str) -> str:
            current = name.strip().lower()
            seen: set[str] = set()
            while current in aliases and current not in seen:
                seen.add(current)
                current = aliases[current]
            return current

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
        object.__setattr__(self, "angle_aliases", aliases)
        object.__setattr__(
            self,
            "angle_single",
            frozenset(canonicalize(name) for name in self.angle_single),
        )
        object.__setattr__(
            self,
            "angle_optional_pair",
            frozenset(canonicalize(name) for name in self.angle_optional_pair),
        )

    def canonical_angle(self, name: str) -> str:
        current = name.strip().lower()
        seen: set[str] = set()
        while current in self.angle_aliases and current not in seen:
            seen.add(current)
            current = self.angle_aliases[current]
        return current

    def allows_angle(self, name: str) -> bool:
        raw_name = name.strip().lower()
        return raw_name in self.angle_allowed or self.canonical_angle(name) in self.angle_allowed

    def allows_bbcode(self, name: str) -> bool:
        return name.lower() in self.bbcode_allowed

    def is_angle_single(self, name: str) -> bool:
        return self.canonical_angle(name) in self.angle_single

    def is_angle_optional_pair(self, name: str) -> bool:
        return self.canonical_angle(name) in self.angle_optional_pair


def default_tag_rules() -> TagRules:
    return _default_tag_rules()


@lru_cache(maxsize=1)
def _default_tag_rules() -> TagRules:
    path = resources.files("phraseloom").joinpath("tag_rules.toml")
    try:
        with path.open("rb") as handle:
            data = tomllib.load(handle)
    except (OSError, UnicodeDecodeError) as exc:
        raise ConfigError(f"could not read default tag rules: {exc}") from exc
    except tomllib.TOMLDecodeError as exc:
        raise ConfigError(f"default tag rules contain invalid TOML: {exc}") from exc
    return _parse_tag_rules(data, source="default")


def load_tag_rules(path: str | Path | None = None) -> TagRules:
    if path is None:
        return default_tag_rules()

    tag_rules_path = Path(path)
    try:
        raw_config = tag_rules_path.read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError) as exc:
        raise ConfigError(f"{tag_rules_path}: could not read tag rules: {exc}") from exc

    try:
        data = tomllib.loads(raw_config)
    except tomllib.TOMLDecodeError as exc:
        raise ConfigError(f"{tag_rules_path}: invalid TOML: {exc}") from exc

    return _parse_tag_rules(data, source=str(tag_rules_path))


def normalized_tag_rules_hash(rules: TagRules) -> str:
    normalized = _normalized_tag_rules(rules)
    payload = json.dumps(normalized, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def tag_rules_payload(rules: TagRules) -> str:
    return json.dumps(
        _normalized_tag_rules(rules),
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )


def tag_rules_from_payload(payload: str, *, source: str = "embedded") -> TagRules:
    try:
        normalized = json.loads(payload)
    except (TypeError, json.JSONDecodeError) as exc:
        raise ConfigError("embedded tag rules are invalid") from exc
    if not isinstance(normalized, dict):
        raise ConfigError("embedded tag rules are invalid")
    try:
        return TagRules(
            version=normalized["version"],
            angle_allowed=frozenset(normalized["angle_allowed"]),
            bbcode_allowed=frozenset(normalized["bbcode_allowed"]),
            protect_raw_braces=normalized["protect_raw_braces"],
            source=source,
            angle_aliases=normalized.get("angle_aliases", {}),
            angle_single=frozenset(normalized.get("angle_single", [])),
            angle_optional_pair=frozenset(
                normalized.get("angle_optional_pair", [])
            ),
        )
    except (KeyError, TypeError, ValueError) as exc:
        raise ConfigError("embedded tag rules are invalid") from exc


def _normalized_tag_rules(rules: TagRules) -> dict[str, Any]:
    return {
        "version": rules.version,
        "angle_allowed": sorted(name.lower() for name in rules.angle_allowed),
        "angle_aliases": {
            alias: rules.angle_aliases[alias] for alias in sorted(rules.angle_aliases)
        },
        "angle_single": sorted(rules.angle_single),
        "angle_optional_pair": sorted(rules.angle_optional_pair),
        "bbcode_allowed": sorted(name.lower() for name in rules.bbcode_allowed),
        "protect_raw_braces": rules.protect_raw_braces,
    }


def _parse_tag_rules(data: dict[str, Any], *, source: str) -> TagRules:
    version = data.get("version")
    if type(version) is not int or version != 1:
        raise ConfigError("tag rules version must be exactly 1")

    angle_allowed = _read_allowlist_section(data, "angle_tags")
    angle_section = _read_required_section(data, "angle_tags")
    angle_aliases = _read_alias_section(angle_section, "angle_tags.aliases")
    angle_single = _read_nested_tag_list(angle_section, "angle_tags.single")
    angle_optional_pair = _read_nested_tag_list(
        angle_section, "angle_tags.optional_pair"
    )
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
        angle_aliases=angle_aliases,
        angle_single=angle_single,
        angle_optional_pair=angle_optional_pair,
    )


def _read_allowlist_section(data: dict[str, Any], section_name: str) -> frozenset[str]:
    section = _read_required_section(data, section_name)

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


def _read_required_section(data: dict[str, Any], section_name: str) -> dict[str, Any]:
    section = data.get(section_name)
    if not isinstance(section, dict):
        raise ConfigError(f"{section_name} section is required")
    return section


def _read_alias_section(
    parent: dict[str, Any], section_name: str
) -> dict[str, str]:
    section_key = section_name.rsplit(".", 1)[-1]
    aliases = parent.get(section_key, {})
    if not isinstance(aliases, dict):
        raise ConfigError(f"{section_name} must be a table")

    normalized: dict[str, str] = {}
    for alias, target in aliases.items():
        if not isinstance(alias, str) or not alias.strip():
            raise ConfigError(f"{section_name} keys must be non-empty strings")
        if not isinstance(target, str) or not target.strip():
            raise ConfigError(f"{section_name} values must be non-empty strings")
        normalized[alias.strip().lower()] = target.strip().lower()

    return normalized


def _read_nested_tag_list(
    parent: dict[str, Any], section_name: str
) -> frozenset[str]:
    section_key = section_name.rsplit(".", 1)[-1]
    section = parent.get(section_key)
    if section is None:
        return frozenset()
    if not isinstance(section, dict):
        raise ConfigError(f"{section_name} must be a table")

    tags = section.get("tags")
    if not isinstance(tags, list):
        raise ConfigError(f"{section_name}.tags must be a list of non-empty strings")

    normalized: set[str] = set()
    for item in tags:
        if not isinstance(item, str) or not item.strip():
            raise ConfigError(
                f"{section_name}.tags must be a list of non-empty strings"
            )
        normalized.add(item.strip().lower())

    return frozenset(normalized)
