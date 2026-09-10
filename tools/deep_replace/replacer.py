from __future__ import annotations

from collections import defaultdict
from pathlib import Path
import tempfile

from tools.excel_file_ops import (
    BatchSummary, OperationResult, atomic_copy, excel_files, optional_output_directory, create_output_directory,
    BACKUP_DIRECTORY,
)
from tools.operation_logs import emit_log, log_result, log_summary


def replace_files(source_dir: str | Path, target_dir: str | Path, *,
                  output_dir: str | Path | None = None,
                  progress_callback=None, log_callback=None) -> BatchSummary:
    emit_log(log_callback, f"开始同名文件替换：{source_dir} → {target_dir}")
    source, target = (Path(path).expanduser().resolve() for path in (source_dir, target_dir))
    if source == target or source.is_relative_to(target) or target.is_relative_to(source):
        raise ValueError("源目录和目标目录不能相同或互相包含。")
    output, inplace = optional_output_directory(output_dir, default=target, inputs=(source, target))
    emit_log(log_callback, "保存方式：原位替换目标文件（先备份）" if inplace else f"保存方式：另存到 {output}")
    sources, _ = excel_files(source, exclude=() if inplace else (output,))
    targets, _ = excel_files(target, exclude=() if inplace else (output,))
    if not sources or not targets:
        raise ValueError("源目录和目标目录均需包含 Excel 文件。")
    source_index, target_index = defaultdict(list), defaultdict(list)
    for path in sources:
        source_index[path.name.casefold()].append(path)
    for path in targets:
        target_index[path.name.casefold()].append(path)
    summary = BatchSummary("deep-replace", output, details={"source_dir": str(source), "target_dir": str(target),
                                                            "inplace": inplace})
    if not inplace:
        create_output_directory(output)
    backup_dir = None
    for index, path in enumerate(targets, 1):
        relative = path.relative_to(target)
        candidates = source_index[path.name.casefold()]
        ambiguous = bool(candidates) and (len(candidates) != 1 or len(target_index[path.name.casefold()]) != 1)
        destination = path if inplace else output / relative
        result = OperationResult(str(relative), "failed")
        try:
            if ambiguous:
                if not inplace:
                    atomic_copy(path, destination)
                result.status = "skipped"
                result.message = "同名源或目标不唯一，保留目标原文件。"
                result.details = {"sources": [str(p) for p in candidates],
                                  "targets": [str(p) for p in target_index[path.name.casefold()]]}
            elif candidates:
                if inplace:
                    if backup_dir is None:
                        backup_root = target / BACKUP_DIRECTORY
                        if not backup_root.resolve().is_relative_to(target):
                            raise ValueError("备份目录不能指向目标目录之外。")
                        backup_root.mkdir(exist_ok=True)
                        backup_dir = Path(tempfile.mkdtemp(prefix="deep-replace-", dir=backup_root))
                        summary.details["backup_dir"] = str(backup_dir)
                        emit_log(log_callback, f"备份目录: {backup_dir}（保留原相对路径）")
                    backup = backup_dir / relative
                    atomic_copy(path, backup)
                    result.details["backup"] = str(backup)
                atomic_copy(candidates[0], destination)
                result.status = "replaced"
                result.details["replacement"] = str(candidates[0])
            else:
                if not inplace:
                    atomic_copy(path, destination)
                result.status = "unchanged"
            result.output = str(destination)
        except Exception as exc:
            result.message = str(exc)
        summary.results.append(result)
        log_result(log_callback, result)
        if progress_callback:
            progress_callback(index, len(targets))
    summary.details["unmatched_sources"] = [str(p.relative_to(source)) for p in sources
                                             if p.name.casefold() not in target_index]
    summary.details["replaced_files"] = sum(row.status == "replaced" for row in summary.results)
    for source_name in summary.details["unmatched_sources"]:
        emit_log(log_callback, f"来源文件没有对应目标: {source_name}")
    log_summary(log_callback, summary, f"替换 {summary.details['replaced_files']} 个文件")
    return summary
