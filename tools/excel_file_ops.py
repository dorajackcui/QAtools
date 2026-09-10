"""Shared file enumeration, output directories and in-memory operation summaries."""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass, field
import json
import os
from pathlib import Path
import shutil
import tempfile


REPORT_NAME = "qatools_tool_report.json"
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls", ".xlsb"}
BACKUP_DIRECTORY = ".qatools-backups"
_session_outputs: set[Path] = set()


@dataclass
class OperationResult:
    source: str
    status: str
    output: str | None = None
    message: str = ""
    details: dict = field(default_factory=dict)


@dataclass
class BatchSummary:
    tool: str
    output_dir: Path
    results: list[OperationResult] = field(default_factory=list)
    details: dict = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)

    @property
    def succeeded_files(self) -> int:
        return sum(row.status in {"updated", "unchanged", "read", "replaced", "copied"} for row in self.results)

    @property
    def failed_files(self) -> int:
        return sum(row.status == "failed" for row in self.results)

    @property
    def skipped_files(self) -> int:
        return sum(row.status == "skipped" for row in self.results)

    def describe(self) -> str:
        lines = [f"输出目录: {self.output_dir}",
                 f"成功: {self.succeeded_files}；失败: {self.failed_files}；跳过: {self.skipped_files}"]
        lines.extend(self.warnings)
        return "\n".join(lines)


def create_output_directory(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=False)
    _session_outputs.add(path.resolve())


def generated_directory(folder: Path) -> bool:
    if folder.resolve() in _session_outputs:
        return True
    # Recognize earlier versions' outputs without creating new marker/report files.
    for name, field_name in ((REPORT_NAME, "tool"), ("content_sync_report.json", "direction")):
        marker = folder / name
        if not marker.is_file():
            continue
        try:
            content = json.loads(marker.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        if isinstance(content, dict) and content.get("schema_version") == 1 and content.get(field_name):
            return True
    return False


def excel_files(folder: str | Path, *, extensions: Iterable[str] = EXCEL_EXTENSIONS,
                exclude: Iterable[Path] = ()) -> tuple[list[Path], list[Path]]:
    root = Path(folder).expanduser().resolve()
    if not root.is_dir():
        raise ValueError(f"目录不存在: {root}")
    excluded = tuple(path.resolve() for path in exclude)
    accepted, skipped = [], []
    extensions = set(extensions)

    def excluded_path(path: Path) -> bool:
        resolved = path.resolve()
        return (not resolved.is_relative_to(root)
                or any(resolved == item or resolved.is_relative_to(item) for item in excluded))

    def on_error(error: OSError) -> None:
        raise error

    for directory, directories, names in os.walk(root, followlinks=False, onerror=on_error):
        parent = Path(directory)
        directories[:] = [name for name in directories
                          if name.casefold() != BACKUP_DIRECTORY and not excluded_path(parent / name)
                          and not generated_directory(parent / name)]
        for name in names:
            path = parent / name
            if name.startswith("~$") or excluded_path(path):
                continue
            suffix = path.suffix.lower()
            if suffix in extensions:
                accepted.append(path)
            elif suffix in EXCEL_EXTENSIONS:
                skipped.append(path)
    key = lambda p: (str(p).casefold(), str(p))
    return sorted(accepted, key=key), sorted(skipped, key=key)


def new_output_directory(output: str | Path, *, inputs: Iterable[Path]) -> Path:
    """Validate now; callers reserve with mkdir only after their other preflight work."""
    path = Path(output).expanduser().resolve()
    for source in inputs:
        source = source.resolve()
        if source == path or source.is_relative_to(path):
            raise ValueError("输出目录不能包含或覆盖输入。")
    if path.exists():
        raise ValueError(f"输出目录已存在，请指定新目录: {path}")
    return path


def optional_output_directory(output: str | Path | None, *, default: Path,
                              inputs: Iterable[Path]) -> tuple[Path, bool]:
    """An omitted/blank output explicitly means operation at the input location."""
    if output is None or isinstance(output, str) and not output.strip():
        return default.resolve(), True
    return new_output_directory(output, inputs=inputs), False


def atomic_copy(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    handle, name = tempfile.mkstemp(dir=destination.parent, suffix=destination.suffix)
    os.close(handle)
    temporary = Path(name)
    try:
        shutil.copy2(source, temporary)
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)
