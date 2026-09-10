from __future__ import annotations

from openpyxl.utils import column_index_from_string

from tools.excel_compatibility.processor import process_excel_copies


def operate_columns(folder_path, *, column: str, action: str = "clear", **kwargs):
    if action not in {"clear", "insert", "delete"}:
        raise ValueError("列操作必须为 clear、insert 或 delete。")
    return process_excel_copies(folder_path, column=column_index_from_string(column.strip().upper()),
                                action=action, **kwargs)
