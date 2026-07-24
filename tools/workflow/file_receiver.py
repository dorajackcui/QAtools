"""Receive Excel files forwarded by the macOS Finder QA quick action."""

from __future__ import annotations

import json
import os
import uuid
from pathlib import Path
from typing import IO

try:
    import fcntl
except ImportError:  # pragma: no cover - Finder integration is macOS-only.
    fcntl = None  # type: ignore[assignment]


SUPPORTED_EXCEL_SUFFIXES = frozenset({".xlsx", ".xlsm"})


def normalize_workflow_input_file(
    file_path: str | os.PathLike[str],
    *,
    require_exists: bool = True,
) -> Path:
    """Return a normalized supported Excel path or raise ``ValueError``."""

    candidate = Path(file_path).expanduser()
    if candidate.suffix.lower() not in SUPPORTED_EXCEL_SUFFIXES:
        raise ValueError("QA workflow 只支持 .xlsx 或 .xlsm 文件。")
    if require_exists and not candidate.is_file():
        raise ValueError(f"Excel 文件不存在：{candidate}")
    return candidate.absolute()


def default_receiver_directory() -> Path:
    user_id = os.getuid() if hasattr(os, "getuid") else "user"
    return Path("/tmp") / f"tagexactor-qa-workflow-{user_id}"


def _receiver_paths(receiver_directory: Path) -> tuple[Path, Path]:
    return receiver_directory / "receiver.lock", receiver_directory / "inbox"


def _receiver_is_active(lock_path: Path) -> bool:
    if fcntl is None:
        return False
    try:
        lock_path.parent.mkdir(parents=True, exist_ok=True)
        lock_file = lock_path.open("a+", encoding="utf-8")
    except OSError:
        return False

    try:
        try:
            fcntl.flock(lock_file.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            return True
        fcntl.flock(lock_file.fileno(), fcntl.LOCK_UN)
        return False
    finally:
        lock_file.close()


def send_workflow_input_file(
    file_path: str | os.PathLike[str],
    *,
    receiver_directory: str | os.PathLike[str] | None = None,
) -> bool:
    """Forward a file to an existing Toolshub process.

    ``False`` means that no current Toolshub receiver could be reached.
    """

    normalized_path = normalize_workflow_input_file(file_path)
    receiver_path = (
        Path(receiver_directory)
        if receiver_directory
        else default_receiver_directory()
    )
    lock_path, inbox_path = _receiver_paths(receiver_path)
    if not _receiver_is_active(lock_path):
        return False

    request_id = uuid.uuid4().hex
    temporary_request = inbox_path / f".{request_id}.tmp"
    ready_request = inbox_path / f"{request_id}.json"
    payload = json.dumps(
        {"path": str(normalized_path)},
        ensure_ascii=False,
    )
    try:
        inbox_path.mkdir(parents=True, exist_ok=True)
        temporary_request.write_text(payload, encoding="utf-8")
        os.replace(temporary_request, ready_request)
    except OSError:
        try:
            temporary_request.unlink()
        except FileNotFoundError:
            pass
        return False
    return True


class WorkflowFileReceiver:
    """Own the current Toolshub lock and expose Finder requests to Tk polling."""

    def __init__(
        self,
        *,
        receiver_directory: str | os.PathLike[str] | None = None,
    ) -> None:
        self.receiver_directory = (
            Path(receiver_directory)
            if receiver_directory
            else default_receiver_directory()
        )
        self.lock_path, self.inbox_path = _receiver_paths(self.receiver_directory)
        self._lock_file: IO[str] | None = None

    def start(self) -> bool:
        """Start receiving, or return ``False`` if another receiver is active."""

        if fcntl is None:
            return False
        if self._lock_file is not None:
            return True

        lock_file: IO[str] | None = None
        try:
            self.receiver_directory.mkdir(parents=True, exist_ok=True)
            os.chmod(self.receiver_directory, 0o700)
            self.inbox_path.mkdir(parents=True, exist_ok=True)
            lock_file = self.lock_path.open("a+", encoding="utf-8")
            os.chmod(self.lock_path, 0o600)
            fcntl.flock(lock_file.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except (OSError, BlockingIOError):
            if lock_file is not None:
                lock_file.close()
            return False

        self._lock_file = lock_file
        return True

    def pop_pending_paths(self) -> tuple[str, ...]:
        paths: list[str] = []
        if self._lock_file is None:
            return ()

        for request_file in sorted(self.inbox_path.glob("*.json")):
            try:
                payload = json.loads(request_file.read_text(encoding="utf-8"))
                requested_path = payload["path"]
                if isinstance(requested_path, str):
                    paths.append(requested_path)
            except (KeyError, OSError, TypeError, ValueError, UnicodeDecodeError):
                pass
            finally:
                try:
                    request_file.unlink()
                except FileNotFoundError:
                    pass
        return tuple(paths)

    def close(self) -> None:
        if self._lock_file is None:
            return
        if fcntl is None:  # pragma: no cover - defensive cross-platform guard.
            self._lock_file.close()
            self._lock_file = None
            return
        try:
            fcntl.flock(self._lock_file.fileno(), fcntl.LOCK_UN)
        finally:
            self._lock_file.close()
            self._lock_file = None
