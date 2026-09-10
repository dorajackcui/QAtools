"""Optional Excel COM session and atomic edits of isolated workbook copies."""

from __future__ import annotations

from contextlib import contextmanager
import os
from pathlib import Path
import shutil
import sys
import tempfile


def _load_com():
    if sys.platform != "win32":
        raise RuntimeError("此操作需要 Windows 和 Microsoft Excel。")
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError('缺少 Excel 支持组件。源码环境请运行 python -m pip install -e ".[excel-com]"。') from exc
    return pythoncom, win32com.client


@contextmanager
def excel_session():
    pythoncom, client = _load_com()
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    application = None
    body_failed = False
    try:
        try:
            application = client.DispatchEx("Excel.Application")
        except Exception as exc:
            raise RuntimeError(f"无法启动 Microsoft Excel，请确认已安装并可正常启动。{exc}") from exc
        application.Visible = False
        application.DisplayAlerts = False
        application.EnableEvents = False
        application.AskToUpdateLinks = False
        application.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        yield application
    except BaseException:
        body_failed = True
        raise
    finally:
        try:
            if application is not None:
                try:
                    application.Quit()
                except Exception:
                    if not body_failed:
                        raise
        finally:
            application = None
            pythoncom.CoUninitialize()


def edit_copy(application, source: Path, destination: Path, *, action: str = "resave",
              column: int = 1, sheet: str | None = None, header_rows: int = 1,
              inserted_header: str = "Translation") -> None:
    """Excel saves a temporary copy; an error leaves source/destination untouched."""
    destination.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".qatools-excel-", dir=destination.parent) as temporary:
        working = Path(temporary) / source.name
        shutil.copy2(source, working)
        workbook = None
        try:
            application.AutomationSecurity = 3
            workbook = application.Workbooks.Open(
                str(working), UpdateLinks=0, ReadOnly=False, Password="", WriteResPassword="",
                IgnoreReadOnlyRecommended=True, Notify=False, AddToMru=False,
            )
            if workbook.ReadOnly:
                raise PermissionError("Excel 以只读方式打开工作簿，未保存。")
            if action != "resave":
                worksheet = workbook.Worksheets(sheet) if sheet else workbook.ActiveSheet
                if column > worksheet.Columns.Count:
                    raise ValueError("配置列超出当前工作簿的列范围。")
                if action == "clear":
                    used = worksheet.UsedRange
                    last_row = used.Row + used.Rows.Count - 1
                    if last_row > header_rows:
                        worksheet.Range(worksheet.Cells(header_rows + 1, column),
                                        worksheet.Cells(last_row, column)).ClearContents()
                elif action == "insert":
                    worksheet.Columns(column).Insert()
                    header = worksheet.Cells(1, column)
                    header.NumberFormat = "@"
                    header.Value = inserted_header
                elif action == "delete":
                    worksheet.Columns(column).Delete()
                else:
                    raise ValueError(f"未知列操作: {action}")
            workbook.Save()
        finally:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
                workbook = None
        os.replace(working, destination)
