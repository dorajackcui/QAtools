"""Receive Excel files forwarded by macOS Finder quick actions."""

from __future__ import annotations

import json
import os
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import IO

try:
    import fcntl
except ImportError:  # pragma: no cover - Finder integration is macOS-only.
    fcntl = None  # type: ignore[assignment]


SUPPORTED_EXCEL_SUFFIXES = frozenset({".xlsx", ".xlsm"})
QA_WORKFLOW_ACTION = "qa_workflow"
FRENCH_NBSP_RESTORE_ACTION = "french_nbsp_restore"
SUPPORTED_FILE_ACTIONS = frozenset(
    {
        QA_WORKFLOW_ACTION,
        FRENCH_NBSP_RESTORE_ACTION,
    }
)


@dataclass(frozen=True)
class ToolFileRequest:
    action: str
    file_path: str


def normalize_excel_input_file(
    file_path: str | os.PathLike[str],
    *,
    action_name: str = "Toolshub",
    require_exists: bool = True,
) -> Path:
    """Return a normalized supported Excel path or raise ``ValueError``."""

    candidate = Path(file_path).expanduser()
    if candidate.suffix.lower() not in SUPPORTED_EXCEL_SUFFIXES:
        raise ValueError(f"{action_name} 只支持 .xlsx 或 .xlsm 文件。")
    if require_exists and not candidate.is_file():
        raise ValueError(f"Excel 文件不存在：{candidate}")
    return candidate.absolute()


def normalize_workflow_input_file(
    file_path: str | os.PathLike[str],
    *,
    require_exists: bool = True,
) -> Path:
    """Return a normalized supported Excel path or raise ``ValueError``."""

    return normalize_excel_input_file(
        file_path,
        action_name="QA workflow",
        require_exists=require_exists,
    )


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

    return send_tool_input_file(
        QA_WORKFLOW_ACTION,
        file_path,
        receiver_directory=receiver_directory,
    )


def send_tool_input_file(
    action: str,
    file_path: str | os.PathLike[str],
    *,
    receiver_directory: str | os.PathLike[str] | None = None,
) -> bool:
    """Forward an Excel file and requested action to the current Toolshub."""

    if action not in SUPPORTED_FILE_ACTIONS:
        raise ValueError(f"不支持的 Toolshub 文件操作：{action}")
    action_name = (
        "NBSP restore"
        if action == FRENCH_NBSP_RESTORE_ACTION
        else "QA workflow"
    )
    normalized_path = normalize_excel_input_file(
        file_path,
        action_name=action_name,
    )
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
        {
            "action": action,
            "path": str(normalized_path),
        },
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

    def pop_pending_requests(self) -> tuple[ToolFileRequest, ...]:
        requests: list[ToolFileRequest] = []
        if self._lock_file is None:
            return ()

        for request_file in sorted(self.inbox_path.glob("*.json")):
            try:
                payload = json.loads(request_file.read_text(encoding="utf-8"))
                requested_action = payload.get("action", QA_WORKFLOW_ACTION)
                requested_path = payload["path"]
                if (
                    requested_action in SUPPORTED_FILE_ACTIONS
                    and isinstance(requested_path, str)
                ):
                    requests.append(
                        ToolFileRequest(
                            action=requested_action,
                            file_path=requested_path,
                        )
                    )
            except (KeyError, OSError, TypeError, ValueError, UnicodeDecodeError):
                pass
            finally:
                try:
                    request_file.unlink()
                except FileNotFoundError:
                    pass
        return tuple(requests)

    def pop_pending_paths(self) -> tuple[str, ...]:
        """Compatibility helper for callers that only expect QA file paths."""

        return tuple(
            request.file_path
            for request in self.pop_pending_requests()
        )

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
