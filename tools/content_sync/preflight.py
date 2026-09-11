"""Read-only selection checks; advisory snapshots, never authorization to write."""
from dataclasses import dataclass
from pathlib import Path
import random
import stat
import sys


def is_readonly(path: Path) -> bool:
    metadata = path.stat()
    attributes = getattr(metadata, "st_file_attributes", None)
    if attributes is not None:
        return bool(attributes & stat.FILE_ATTRIBUTE_READONLY)
    return not bool(metadata.st_mode & (stat.S_IWUSR | stat.S_IWGRP | stat.S_IWOTH))


def _probe_windows_file(path: Path) -> None:
    """Probe sharing without changing contents; close the handle immediately.

    https://learn.microsoft.com/windows/win32/api/fileapi/nf-fileapi-createfilew
    """
    import ctypes
    from ctypes import wintypes

    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    create = kernel.CreateFileW
    create.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                       ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
    create.restype = wintypes.HANDLE
    close = kernel.CloseHandle
    close.argtypes = [wintypes.HANDLE]
    close.restype = wintypes.BOOL
    handle = create(str(path), 0x80000000, 0, None, 3, 0x80, None)  # READ, exclusive, OPEN_EXISTING
    if handle == ctypes.c_void_p(-1).value:
        raise ctypes.WinError(ctypes.get_last_error())
    close(handle)


def master_warnings(path: Path) -> tuple[str, ...]:
    warnings = []
    try:
        if is_readonly(path):
            warnings.append("Master 文件为只读；原位回填可能失败，作为来源读取不受此属性影响。")
        if sys.platform == "win32":
            _probe_windows_file(path)
    except OSError as exc:
        if getattr(exc, "winerror", None) in {32, 33}:
            warnings.append("Master 正被其他程序占用。请先保存并关闭 Excel 中的该文件，再进行同步。")
        else:
            warnings.append(f"无法确认 Master 的访问状态：{exc}")
    # A lock file can be stale; never describe it as proof of a live Excel owner.
    if not warnings and path.with_name("~$" + path.name).exists():
        warnings.append("发现 Master 的 Excel 临时锁文件，可能仍在编辑，也可能是残留文件。请确认已保存并关闭。")
    return tuple(warnings)


@dataclass(frozen=True)
class MasterCheck:
    choices: object | None
    warnings: tuple[str, ...]
    error: str = ""


def inspect_master(path: str) -> MasterCheck:
    from tools.excel_metadata import list_workbook_sheets

    source = Path(path).expanduser().resolve()
    warnings = master_warnings(source)
    try:
        return MasterCheck(list_workbook_sheets(source), warnings)
    except Exception as exc:
        return MasterCheck(None, warnings, str(exc))


@dataclass(frozen=True)
class DirectoryCheck:
    total: int
    unsupported: int
    sampled: tuple[str, ...]
    readonly: tuple[str, ...]
    errors: tuple[str, ...]

    def describe(self, *, reverse: bool, inplace: bool) -> str:
        lines = [f"可处理小表：{self.total} 个（.xlsx / .xlsm）",
                 f"不支持格式：{self.unsupported} 个",
                 f"随机抽检：{len(self.sampled)} / {self.total} 个；只读：{len(self.readonly)} 个"]
        if self.readonly:
            lines.append("只读文件：\n" + "\n".join(self.readonly))
            if reverse:
                lines.append("小表仅作为来源读取，只读属性不影响回填 Master。")
            elif inplace:
                lines.append("原位更新这些文件可能失败，可解除只读或另存。")
            else:
                lines.append("已选择另存目录，原小表不会被覆盖；输出仍以实际保存结果为准。")
        if self.errors:
            lines.append("无法检查：\n" + "\n".join(self.errors))
        if self.total == 0:
            lines.append("没有可处理的小表，请重新选择目录。")
        lines.append("仅检查文件只读属性；抽检结果不代表全部文件可写，也不检测 Excel 写保护或文件占用。")
        return "\n\n".join(lines)


def inspect_targets(folder: str, *, master: str = "", output: str = "", reverse: bool = False) -> DirectoryCheck:
    from tools.excel_file_ops import excel_files
    from .master_to_target import _input_files

    root = Path(folder).expanduser().resolve()
    if not root.is_dir():
        raise ValueError(f"小表目录不存在：{root}")
    master_path = Path(master).expanduser().resolve() if master.strip() else None
    output_path = Path(output).expanduser().resolve() if output.strip() else None
    if reverse:
        files, unsupported = excel_files(root, extensions={".xlsx", ".xlsm"},
                                         exclude=tuple(p for p in (master_path, output_path) if p is not None))
    else:
        files, unsupported = _input_files(root, output_path, master_path)
    sample = sorted(random.sample(files, min(20, len(files))), key=lambda p: (str(p).casefold(), str(p)))
    readonly, errors = [], []
    for path in sample:
        relative = str(path.relative_to(root))
        try:
            if is_readonly(path):
                readonly.append(relative)
        except OSError as exc:
            errors.append(f"{relative}：{exc}")
    return DirectoryCheck(len(files), len(unsupported), tuple(str(p.relative_to(root)) for p in sample),
                          tuple(readonly), tuple(errors))
