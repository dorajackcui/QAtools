from __future__ import annotations

from pathlib import Path

from tools.excel_com import edit_copy, excel_session
from tools.excel_file_ops import BatchSummary, OperationResult, excel_files, optional_output_directory, create_output_directory
from tools.operation_logs import emit_log, log_result, log_summary


def process_excel_copies(folder_path: str | Path, *, output_dir: str | Path | None = None,
                         action: str = "resave", column: int = 1, sheet: str | None = None,
                         header_rows: int = 1, inserted_header: str = "Translation",
                         progress_callback=None, log_callback=None) -> BatchSummary:
    if action not in {"resave", "clear", "insert", "delete"}:
        raise ValueError("未知 Excel 操作。")
    if isinstance(column, bool) or not isinstance(column, int) or not 1 <= column <= 16384:
        raise ValueError("列号必须在 A–XFD 范围内。")
    if isinstance(header_rows, bool) or not isinstance(header_rows, int) or not 0 <= header_rows < 1048576:
        raise ValueError("表头行数必须是 0–1048575 的整数。")
    if not isinstance(inserted_header, str) or len(inserted_header) > 32767:
        raise ValueError("插入列表头必须是长度不超过 32767 的文本。")
    action_label = {"resave": "兼容性重存", "clear": "清空列", "insert": "插入列", "delete": "删除列"}[action]
    emit_log(log_callback, f"开始{action_label}：{folder_path}")
    folder = Path(folder_path).expanduser().resolve()
    output, inplace = optional_output_directory(output_dir, default=folder, inputs=(folder,))
    emit_log(log_callback, "保存方式：原位更新文件" if inplace else f"保存方式：另存到 {output}")
    files, skipped = excel_files(folder, extensions={".xlsx", ".xlsm", ".xls"}, exclude=() if inplace else (output,))
    if not files:
        raise ValueError("目录中没有支持的 .xlsx/.xlsm/.xls 文件。")
    summary = BatchSummary("excel-compatibility" if action == "resave" else "column-tools", output,
                           details={"action": action, "column": column, "sheet": sheet,
                                    "header_rows": header_rows, "inserted_header": inserted_header, "inplace": inplace})
    initialized = False
    emit_log(log_callback, "正在启动独立 Excel 会话…")
    try:
        with excel_session() as application:
            # Only reserve outputs once Excel is usable.
            if not inplace:
                create_output_directory(output)
            initialized = True
            for index, path in enumerate(files, 1):
                relative = path.relative_to(folder)
                destination = output / relative
                try:
                    edit_copy(application, path, destination, action=action, column=column,
                              sheet=sheet, header_rows=header_rows, inserted_header=inserted_header)
                    result = OperationResult(str(relative), "updated", str(destination))
                except Exception as exc:
                    result = OperationResult(str(relative), "failed", message=str(exc))
                summary.results.append(result)
                log_result(log_callback, result)
                if progress_callback:
                    progress_callback(index, len(files))
    except Exception as exc:
        if not initialized:
            raise
        summary.warnings.append(f"Excel 会话未正常结束: {exc}")
    summary.results.extend(OperationResult(str(path.relative_to(folder)), "skipped", message="不支持的格式") for path in skipped)
    for result in summary.results:
        if result.status == "skipped":
            log_result(log_callback, result)
    log_summary(log_callback, summary)
    return summary


def resave_workbooks(folder_path: str | Path, **kwargs) -> BatchSummary:
    return process_excel_copies(folder_path, action="resave", **kwargs)
