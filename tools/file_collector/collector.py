"""Read-only copy plans and guarded, byte-preserving file extraction."""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, replace
import os
from pathlib import Path
import re
import shutil
import tempfile

from tools.excel_file_ops import (
    BatchSummary, EXCEL_EXTENSIONS, OperationResult, create_output_directory, excel_files,
)
from tools.operation_logs import emit_log, log_result, log_summary


@dataclass(frozen=True)
class CopyEntry:
    names: tuple[str, ...]
    source: Path  # Relative to the source root.
    destination: Path  # Relative to the output root.
    fingerprint: tuple[int, ...] | None
    status: str = "ready"
    message: str = ""


@dataclass(frozen=True)
class CopyPlan:
    source_dir: Path
    output_dir: Path
    names: tuple[str, ...]
    preserve_tree: bool
    entries: tuple[CopyEntry, ...]
    not_found: tuple[str, ...]

    @property
    def ready_count(self) -> int:
        return sum(entry.status == "ready" for entry in self.entries)

    @property
    def conflict_count(self) -> int:
        return sum(entry.status == "conflict" for entry in self.entries)

    def describe(self) -> str:
        unavailable = sum(entry.status == "unavailable" for entry in self.entries)
        return (f"清单 {len(self.names)} 项；匹配 {len(self.entries)} 个文件；可复制 {self.ready_count}；"
                f"冲突 {self.conflict_count}；不可读 {unavailable}；未找到 {len(self.not_found)} 项")


def parse_names(text: str, *, comma_separated: bool = False) -> tuple[str, ...]:
    """Keep spaces and punctuation inside names; never accept paths or patterns."""
    names, seen = [], set()
    for raw in re.split(r"[\r\n\t,，]+" if comma_separated else r"[\r\n\t]+", text.lstrip("\ufeff")):
        name = raw.strip()
        if not name:
            continue
        if (name in {".", ".."} or any(ord(char) < 32 or char in '/\\:*?"<>|' for char in name)
                or name.endswith(".")):
            raise ValueError(f"请只填写文件名，不含路径、通配符或非法字符：{name}")
        key = name.casefold()
        if key not in seen:
            names.append(name)
            seen.add(key)
    if not names:
        raise ValueError("请输入至少一个文件名。")
    return tuple(names)


def read_names_file(path: str | Path) -> str:
    try:
        return Path(path).read_text(encoding="utf-8-sig")
    except UnicodeError as exc:
        raise ValueError("文件名清单需使用 UTF-8 编码的 TXT 文件。") from exc


def _directories(source_dir, output_dir) -> tuple[Path, Path]:
    if not str(source_dir or "").strip() or not str(output_dir or "").strip():
        raise ValueError("请选择来源目录并指定新输出目录。")
    source, output = (Path(value).expanduser().resolve() for value in (source_dir, output_dir))
    if not source.is_dir():
        raise ValueError(f"来源目录不存在：{source}")
    if source == output or source.is_relative_to(output) or output.is_relative_to(source):
        raise ValueError("来源目录和输出目录不能相同或互相包含。")
    if output.exists() or output.is_symlink():
        raise ValueError(f"输出目录已存在，请指定新目录：{output}")
    return source, output


def _fingerprint(path: Path) -> tuple[int, ...]:
    info = path.stat()
    return (info.st_dev, info.st_ino, info.st_size, info.st_mtime_ns, info.st_ctime_ns)


def build_copy_plan(source_dir: str | Path, names_text: str, *, output_dir: str | Path,
                    preserve_tree: bool = False, comma_separated: bool = False,
                    log_callback=None) -> CopyPlan:
    names = parse_names(names_text, comma_separated=comma_separated)
    source, output = _directories(source_dir, output_dir)
    emit_log(log_callback, f"扫描来源目录：{source}")
    files, _ = excel_files(source, skip_links=True)
    full_names, stems = defaultdict(list), defaultdict(list)
    for path in files:
        full_names[path.name.casefold()].append(path)
        stems[path.stem.casefold()].append(path)
    matched, missing = {}, []
    for name in names:
        index = full_names if Path(name).suffix.casefold() in EXCEL_EXTENSIONS else stems
        candidates = index.get(name.casefold(), ())
        if not candidates:
            missing.append(name)
        for path in candidates:
            matched.setdefault(path, []).append(name)

    entries = []
    for path in sorted(matched, key=lambda p: (str(p).casefold(), str(p))):
        relative = path.relative_to(source)
        destination = relative if preserve_tree else Path(path.name)
        try:
            entry = CopyEntry(tuple(matched[path]), relative, destination, _fingerprint(path))
        except OSError as exc:
            entry = CopyEntry(tuple(matched[path]), relative, destination, None, "unavailable", str(exc))
        entries.append(entry)

    # Detect full-path collisions AND file/directory collisions, including case-only
    # differences on a case-sensitive source that will be copied to Windows.
    destinations = defaultdict(list)
    for index, entry in enumerate(entries):
        destinations[entry.destination.as_posix().casefold()].append(index)
    conflicts = {index for group in destinations.values() if len(group) > 1 for index in group}
    for index, entry in enumerate(entries):
        for parent in entry.destination.parents:
            group = destinations.get(parent.as_posix().casefold(), ())
            if group:
                conflicts.update((index, *group))
    for index in conflicts:
        entries[index] = replace(entries[index], status="conflict", message="输出路径冲突；请调整清单或目录结构选项。")
    plan = CopyPlan(source, output, names, preserve_tree, tuple(entries), tuple(missing))
    emit_log(log_callback, plan.describe())
    return plan


def _check_source(plan: CopyPlan, entry: CopyEntry) -> Path:
    source = plan.source_dir / entry.source
    if not source.resolve().is_relative_to(plan.source_dir) or _fingerprint(source) != entry.fingerprint:
        raise ValueError("来源文件已变化，请重新预览。")
    return source


def _check_destination(root: Path, destination: Path) -> None:
    if root.resolve() != root or not destination.resolve().is_relative_to(root):
        raise ValueError("输出路径已变化，请重新预览。")
    if destination.exists() or destination.is_symlink():
        raise FileExistsError(f"输出文件已存在，未覆盖：{destination.name}")


def _copy_new_file(plan: CopyPlan, entry: CopyEntry) -> None:
    source = _check_source(plan, entry)
    destination = plan.output_dir / entry.destination
    _check_destination(plan.output_dir, destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    _check_destination(plan.output_dir, destination)
    descriptor, name = tempfile.mkstemp(prefix=".qatools-", suffix=".tmp", dir=destination.parent)
    os.close(descriptor)
    temporary = Path(name)
    try:
        shutil.copyfile(source, temporary)
        _check_source(plan, entry)
        info = source.stat()
        os.utime(temporary, ns=(info.st_atime_ns, info.st_mtime_ns))
        _check_destination(plan.output_dir, destination)
        # Windows rename refuses an existing destination. POSIX link publishes a
        # complete staged file with the same no-replace guarantee. Never fall back
        # to os.replace, which could destroy a file created since preview.
        if os.name == "nt":
            os.rename(temporary, destination)
        else:
            os.link(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)


def execute_copy_plan(plan: CopyPlan, *, progress_callback=None, log_callback=None) -> BatchSummary:
    # Re-scan before any writes: newly added duplicates also invalidate the plan.
    fresh = build_copy_plan(plan.source_dir, "\n".join(plan.names), output_dir=plan.output_dir,
                            preserve_tree=plan.preserve_tree)
    if fresh != plan:
        raise ValueError("匹配结果或来源文件已变化，请重新预览。")
    summary = BatchSummary("collect-files", plan.output_dir, details={
        "request_count": len(plan.names), "matched_files": len(plan.entries),
        "conflict_files": plan.conflict_count, "not_found": list(plan.not_found),
    })
    summary.warnings.extend(f"未找到：{name}" for name in plan.not_found)
    emit_log(log_callback, f"文件提取输出目录：{plan.output_dir}")
    if plan.ready_count:
        create_output_directory(plan.output_dir)
    for index, entry in enumerate(plan.entries, 1):
        result = OperationResult(str(entry.source), "skipped", message=entry.message)
        if entry.status == "ready":
            try:
                _copy_new_file(plan, entry)
                result.status = "copied"
                result.output = str(plan.output_dir / entry.destination)
            except Exception as exc:
                result.status = "failed"
                result.message = str(exc)
        summary.results.append(result)
        log_result(log_callback, result)
        if progress_callback:
            progress_callback(index, len(plan.entries))
    summary.details["copied_files"] = summary.succeeded_files
    log_summary(log_callback, summary, f"未找到 {len(plan.not_found)} 项")
    return summary
