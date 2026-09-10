"""Reproducible synthetic sync benchmark; only touches its temporary directory.

Run without arguments for all cases. Each run uses a fresh child process so peak
RSS is comparable. Timings are exclusive wall time, including wrapper overhead;
they are diagnostics, not a CI performance threshold.
"""
from __future__ import annotations

import argparse
from collections import defaultdict
from contextlib import ExitStack, contextmanager
import ctypes
from io import BytesIO
import json
from pathlib import Path
import platform
import subprocess
import sys
import tempfile
from time import perf_counter
from unittest.mock import patch
from xml.etree import ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

CASES = {"large": (20000, 20, 1), "many": (20000, 200, 1),
         "multi": (10000, 20, 12), "formulas": (10000, 20, 1),
         "resave": (1000, 5, 1)}


class Timings:
    def __init__(self):
        self.seconds = defaultdict(float)
        self.stack = []

    @contextmanager
    def measure(self, name):
        frame = [perf_counter(), 0.0]
        self.stack.append(frame)
        try:
            yield
        finally:
            elapsed = perf_counter() - frame[0]
            self.seconds[name] += elapsed - frame[1]
            self.stack.pop()
            if self.stack:
                self.stack[-1][1] += elapsed

    def wrap(self, function, name):
        def wrapped(*args, **kwargs):
            with self.measure(name):
                return function(*args, **kwargs)
        return wrapped


def peak_rss_mib():
    if sys.platform == "win32":
        from ctypes import wintypes
        class Counters(ctypes.Structure):
            _fields_ = [("cb", wintypes.DWORD), ("PageFaultCount", wintypes.DWORD)] + [
                (name, ctypes.c_size_t) for name in ("PeakWorkingSetSize", "WorkingSetSize",
                "QuotaPeakPagedPoolUsage", "QuotaPagedPoolUsage", "QuotaPeakNonPagedPoolUsage",
                "QuotaNonPagedPoolUsage", "PagefileUsage", "PeakPagefileUsage")]
        kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel.GetCurrentProcess.restype = wintypes.HANDLE
        query = ctypes.WinDLL("psapi", use_last_error=True).GetProcessMemoryInfo
        query.argtypes = [wintypes.HANDLE, ctypes.POINTER(Counters), wintypes.DWORD]
        query.restype = wintypes.BOOL
        counters = Counters()
        counters.cb = ctypes.sizeof(counters)
        if not query(kernel.GetCurrentProcess(), ctypes.byref(counters), counters.cb):
            raise ctypes.WinError(ctypes.get_last_error())
        return counters.PeakWorkingSetSize / 1024**2
    import resource
    peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    return peak / (1024**2 if sys.platform == "darwin" else 1024)


def fixture(path, start, stop, columns, *, master, source, formulas):
    from openpyxl import Workbook
    book = Workbook(write_only=True)
    sheet = book.create_sheet("Data")
    sheet.append((["id"] if master else []) + ["key", "source"] + [f"text{i}" for i in range(columns)])
    for i in range(start, stop):
        content = [f"  translated {i}/{j}\n<tag>正文</tag>  " if source else "old" for j in range(columns)]
        if source and formulas and i % 10 == 0:
            content[0] = "=7"
        sheet.append(([i] if master else []) + [f"key{i}", f"Source {i}"] + content)
    try:
        book.save(path)
    finally:
        book.close()
    if source and formulas:
        # Supply a known cached value without requiring Excel for this fixture.
        replacement = BytesIO()
        with ZipFile(path) as original, ZipFile(replacement, "w", ZIP_DEFLATED) as output:
            for entry in original.infolist():
                data = original.read(entry.filename)
                if entry.filename == "xl/worksheets/sheet1.xml":
                    document = ET.fromstring(data)
                    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
                    for cell in document.iter(ns + "c"):
                        if cell.find(ns + "f") is not None:
                            value = cell.find(ns + "v")
                            if value is None:
                                value = ET.SubElement(cell, ns + "v")
                            value.text = "7"
                    data = ET.tostring(document, encoding="utf-8")
                output.writestr(entry, data)
        path.write_bytes(replacement.getvalue())


def run(case, direction):
    import openpyxl
    from tools.content_sync import master_to_target as forward, target_to_master as reverse
    from tools import excel_com
    from tools.qt_operation_logs import LogBuffer

    rows, files, columns = CASES[case]
    timings = Timings()
    with tempfile.TemporaryDirectory(prefix="qatools-sync-benchmark-") as directory:
        root = Path(directory)
        master, targets = root / "master.xlsx", root / "targets"
        targets.mkdir()
        fixture(master, 0, rows, columns, master=True, source=direction == "forward", formulas=case == "formulas")
        for index in range(files):
            fixture(targets / f"{index:04d}.xlsx", index * rows // files, (index + 1) * rows // files,
                    columns, master=False, source=direction == "reverse", formulas=case == "formulas")

        load = forward.load_workbook
        def timed_load(*args, **kwargs):
            cached = kwargs.get("data_only", False)
            label = "formula_cache" if cached else "source_read"
            with timings.measure(label):
                book = load(*args, **kwargs)
            for sheet in book.worksheets:
                original_rows = sheet.iter_rows
                def timed_rows(*args, _rows=original_rows, _label=label, **kwargs):
                    iterator = iter(_rows(*args, **kwargs))
                    while True:
                        with timings.measure(_label):
                            row = next(iterator, None)
                        if row is None:
                            break
                        yield row
                sheet.iter_rows = timed_rows
            return book

        logs = LogBuffer()
        session = excel_com.excel_session
        @contextmanager
        def timed_session():
            with ExitStack() as cleanup:
                with timings.measure("excel_session"):
                    application = cleanup.enter_context(session())
                try:
                    yield application
                finally:
                    with timings.measure("excel_session"):
                        cleanup.close()

        with ExitStack() as stack:
            stack.enter_context(patch.object(excel_com, "excel_session", timed_session))
            stack.enter_context(patch.object(forward, "load_workbook", timed_load))
            for module, names in ((forward, {"_input_files": "scan", "_read_master": "source_index",
                    "load_workbook_for_editing": "target_read", "_sync_file": "match_write",
                    "_save_output": "save"}),
                    (reverse, {"excel_files": "scan", "_read_master": "source_index", "_sync_file": "match_write"}),
                    (excel_com, {"edit_copy": "excel_resave"})):
                for name, label in names.items():
                    stack.enter_context(patch.object(module, name, timings.wrap(getattr(module, name), label)))
            for module in (forward, reverse):
                for name in ("emit_log", "log_result", "log_summary"):
                    stack.enter_context(patch.object(module, name, timings.wrap(getattr(module, name), "logs")))
            with timings.measure("orchestration"):
                if direction == "forward":
                    result = forward.sync_master_to_targets(master, targets, column_count=columns,
                        compatibility_resave=case == "resave", log_callback=logs.append)
                    updates = result.updated_cells
                else:
                    result = reverse.sync_targets_to_master(master, targets, log_callback=logs.append)
                    updates = result.details["updated_cells"]
        assert result.failed_files == 0, result
        assert updates == rows * (columns if direction == "forward" else 1), updates
        return {"case": case, "direction": direction, "rows": rows, "files": files,
                "columns": columns, "seconds": dict(timings.seconds),
                "total_seconds": sum(timings.seconds.values()), "peak_process_rss_mib": peak_rss_mib(),
                "python": platform.python_version(), "openpyxl": openpyxl.__version__,
                "platform": platform.platform(), "updated_cells": updates}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--case", choices=CASES)
    parser.add_argument("--direction", choices=("forward", "reverse"), default="forward")
    parser.add_argument("--repeat", type=int, default=3)
    parser.add_argument("--include-resave", action="store_true", help="Requires installed desktop Excel")
    args = parser.parse_args()
    if args.case:
        print(json.dumps(run(args.case, args.direction), ensure_ascii=False))
        return
    for case in CASES:
        if case == "resave" and not args.include_resave:
            continue
        for direction in (("forward",) if case in {"multi", "resave"} else ("forward", "reverse")):
            for _ in range(args.repeat):
                subprocess.run([sys.executable, str(Path(__file__).resolve()), "--case", case,
                                "--direction", direction], check=True)


if __name__ == "__main__":
    main()
