"""Persistent, shared history-TB project presets."""

from __future__ import annotations

import json
import os
import sys
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class TbProject:
    name: str
    file_path: str
    sheet: str
    source_column: str
    target_column: str
    start_row: int = 2

    @classmethod
    def from_dict(cls, data: Any) -> TbProject | None:
        if not isinstance(data, dict):
            return None
        try:
            project = cls(
                name=str(data["name"]).strip(),
                file_path=str(data["file_path"]).strip(),
                sheet=str(data["sheet"]).strip(),
                source_column=str(data["source_column"]).strip(),
                target_column=str(data["target_column"]).strip(),
                start_row=int(data.get("start_row", 2)),
            )
        except (KeyError, TypeError, ValueError):
            return None
        if (
            not project.name
            or not project.file_path
            or not project.sheet
            or not project.source_column
            or not project.target_column
            or project.start_row < 1
        ):
            return None
        return project


def default_tb_projects_path() -> Path:
    if sys.platform == "darwin":
        return (
            Path.home()
            / "Library"
            / "Application Support"
            / "Toolshub"
            / "tb_projects.json"
        )
    if os.name == "nt":
        app_data = os.environ.get("APPDATA")
        base_path = Path(app_data) if app_data else Path.home() / "AppData" / "Roaming"
        return base_path / "Toolshub" / "tb_projects.json"
    config_home = os.environ.get("XDG_CONFIG_HOME")
    base_path = Path(config_home) if config_home else Path.home() / ".config"
    return base_path / "toolshub" / "tb_projects.json"


class TbProjectStore:
    def __init__(self, config_path: str | os.PathLike[str] | None = None) -> None:
        self.config_path = (
            Path(config_path).expanduser()
            if config_path is not None
            else default_tb_projects_path()
        )

    def list_projects(self) -> tuple[TbProject, ...]:
        if not self.config_path.is_file():
            return ()
        try:
            payload = json.loads(self.config_path.read_text(encoding="utf-8"))
        except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ValueError(f"无法读取 TB 项目配置：{exc}") from exc

        raw_projects = payload.get("projects", ()) if isinstance(payload, dict) else ()
        if not isinstance(raw_projects, list):
            raw_projects = ()
        projects = [
            project
            for raw_project in raw_projects
            if (project := TbProject.from_dict(raw_project)) is not None
        ]
        return tuple(sorted(projects, key=lambda project: project.name.casefold()))

    def find_project(self, name: str) -> TbProject | None:
        normalized_name = name.strip().casefold()
        return next(
            (
                project
                for project in self.list_projects()
                if project.name.casefold() == normalized_name
            ),
            None,
        )

    def save_project(self, project: TbProject) -> None:
        validated_project = TbProject.from_dict(asdict(project))
        if validated_project is None:
            raise ValueError("TB 项目信息不完整。")

        projects = list(self.list_projects())
        normalized_name = validated_project.name.casefold()
        projects = [
            existing
            for existing in projects
            if existing.name.casefold() != normalized_name
        ]
        projects.append(validated_project)
        self._write_projects(projects)

    def delete_project(self, name: str) -> bool:
        projects = list(self.list_projects())
        normalized_name = name.strip().casefold()
        remaining_projects = [
            project
            for project in projects
            if project.name.casefold() != normalized_name
        ]
        if len(remaining_projects) == len(projects):
            return False
        self._write_projects(remaining_projects)
        return True

    def _write_projects(self, projects: list[TbProject]) -> None:
        self.config_path.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "version": 1,
            "projects": [
                asdict(project)
                for project in sorted(
                    projects,
                    key=lambda item: item.name.casefold(),
                )
            ],
        }
        temporary_path = self.config_path.with_suffix(
            f"{self.config_path.suffix}.tmp"
        )
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
