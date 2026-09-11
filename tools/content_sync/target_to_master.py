"""Collect target translations in deterministic order and save a new Master."""

from __future__ import annotations

from dataclasses import asdict
from pathlib import Path

from tools.excel_file_ops import (
    BatchSummary, OperationResult, excel_files, optional_output_directory, create_output_directory,
)
from tools.operation_logs import emit_log, log_result, log_summary
from .master_to_target import ColumnMapping, FileResult, _read_master, _sync_file
from .parallel import ordered_results, validate_workers


def sync_targets_to_master(
    master_file: str | Path, target_dir: str | Path, *, output_dir: str | Path | None = None,
    master_columns: ColumnMapping = ColumnMapping("B", "C", "D"),
    target_columns: ColumnMapping = ColumnMapping("A", "B", "C"),
    master_sheet: str | None = None, target_sheet: str | None = None,
    master_header_rows: int = 1, target_header_rows: int = 1,
    fill_blank_only: bool = False, allow_blank_write: bool = False,
    workers: int = 1,
    progress_callback=None, log_callback=None,
) -> BatchSummary:
    validate_workers(workers)
    emit_log(log_callback, f"开始回填：{target_dir} → {master_file}")
    master_columns.indexes(1)
    target_columns.indexes(1)
    for value in (master_header_rows, target_header_rows):
        if isinstance(value, bool) or not isinstance(value, int) or not 0 <= value < 1048576:
            raise ValueError("表头行数必须是 0–1048575 的整数。")
    master = Path(master_file).expanduser().resolve()
    folder = Path(target_dir).expanduser().resolve()
    if not master.is_file() or master.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("Master 必须是存在的 .xlsx/.xlsm 文件。")
    output, inplace = optional_output_directory(output_dir, default=master.parent, inputs=(master, folder))
    emit_log(log_callback, "保存方式：原位更新 Master" if inplace else f"保存方式：另存到 {output}")
    files, unsupported = excel_files(folder, extensions={".xlsx", ".xlsm"},
                                     exclude=(master,) if inplace else (master, output))
    if not files:
        raise ValueError("目录中没有支持的小表。")
    summary = BatchSummary("target-to-master", output)
    merged, origins, conflicts = {}, {}, []

    def read(path):
        return _read_master(path, target_columns, 1, target_sheet, target_header_rows,
                            skip_blank_content=not allow_blank_write)

    for index, (data, error) in enumerate(ordered_results(read, files, min(workers, len(files))), 1):
        path = files[index - 1]
        relative = str(path.relative_to(folder))
        try:
            if error is not None:
                raise ValueError(error)
            records, duplicates, sheet_name = data
            for duplicate in duplicates:
                conflicts.append({"file": relative, "sheet": sheet_name, **duplicate})
                emit_log(log_callback, "小表重复身份（后行优先）", conflicts[-1])
            for identity, record in records.items():
                origin = {"file": relative, "sheet": sheet_name, "row": record[0]}
                if identity in merged:
                    conflicts.append({"key": identity[0], "match": identity[1],
                                      "previous": origins[identity], "selected": origin,
                                      "previous_content": merged[identity][1], "selected_content": record[1]})
                    emit_log(log_callback, "跨文件重复身份（后文件优先）", conflicts[-1])
                merged[identity] = record
                origins[identity] = origin
            summary.results.append(OperationResult(relative, "read", details={"identities": len(records)}))
        except Exception as exc:
            summary.results.append(OperationResult(relative, "failed", message=str(exc)))
        log_result(log_callback, summary.results[-1])
        if progress_callback:
            progress_callback(index, len(files))
    summary.results.extend(OperationResult(str(path.relative_to(folder)), "skipped", message="不支持的格式")
                           for path in unsupported)
    for result in summary.results:
        if result.status == "skipped":
            log_result(log_callback, result)
    if not inplace:
        create_output_directory(output)
    result = FileResult(str(master))
    identities = set()
    if any(item.status == "read" for item in summary.results):
        try:
            _sync_file(master, output / master.name, merged, master_columns, 1, master_sheet,
                       master_header_rows, fill_blank_only, allow_blank_write, result, identities)
        except Exception as exc:
            result.error = str(exc)
            result.updated_cells = 0
    else:
        result.error = "所有小表读取失败，未生成 Master。"
    summary.results.append(OperationResult(str(master), result.status, result.output, result.error or "",
                                           details=asdict(result)))
    log_result(log_callback, summary.results[-1])
    # Report candidates absent from Master, including origins, without appending new rows.
    unmatched = []
    if result.output:
        unmatched = [{"key": key[0], "match": key[1], **origins[key]} for key in merged if key not in identities]
    summary.details = {"master_file": str(master), "target_dir": str(folder),
                       "master_columns": asdict(master_columns), "target_columns": asdict(target_columns),
                       "master_sheet": master_sheet, "target_sheet": target_sheet,
                       "master_header_rows": master_header_rows, "target_header_rows": target_header_rows,
                       "fill_blank_only": fill_blank_only, "allow_blank_write": allow_blank_write,
                       "inplace": inplace,
                       "updated_cells": result.updated_cells, "conflicts": conflicts,
                       "unmatched_candidates": unmatched, "master_output": result.output}
    for candidate in unmatched:
        emit_log(log_callback, "Master 中未找到身份", candidate)
    log_summary(log_callback, summary, f"更新 {result.updated_cells} 格")
    return summary
