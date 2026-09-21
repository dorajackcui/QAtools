"""Forward page selections to the running Qt window on all desktop platforms."""

from __future__ import annotations

import hashlib
import json
import time

from PySide6.QtCore import QObject, QStandardPaths, QTimer, Signal
from PySide6.QtNetwork import QLocalServer, QLocalSocket


PATH_ACTIONS = frozenset({
    "qa_workflow", "french_nbsp_restore", "compatibility", "excel_merger", "untranslated_stats",
})


def send_path_request(action: str, path: str, *, timeout_ms: int = 5000) -> bool:
    """Wait for a starting instance, then send one acknowledged path request."""
    if action not in PATH_ACTIONS:
        return False
    payload = json.dumps({"action": action, "path": path}, ensure_ascii=False).encode() + b"\n"
    if len(payload) > 4096:
        return False
    deadline = time.monotonic() + timeout_ms / 1000
    while True:
        socket = QLocalSocket()
        try:
            socket.connectToServer(navigation_server_name())
            if socket.waitForConnected(200):
                socket.write(payload)
                if socket.bytesToWrite() and not socket.waitForBytesWritten(2000):
                    return False
                # Never resend after writing: NBSP requests must execute only once.
                if not socket.canReadLine() and not socket.waitForReadyRead(2000):
                    return False
                return bytes(socket.readLine()) == b"ok\n"
        finally:
            socket.disconnectFromServer()
        if time.monotonic() >= deadline:
            return False
        time.sleep(0.05)


def navigation_server_name() -> str:
    user_directory = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.GenericConfigLocation)
    user_key = hashlib.sha256(user_directory.encode()).hexdigest()[:16]
    return f"qatools-navigation-{user_key}"


def send_tool_selection(tool: str, checks: list[str] | None = None) -> bool:
    socket = QLocalSocket()
    try:
        socket.connectToServer(navigation_server_name())
        if not socket.waitForConnected(1000):
            return False
        payload = json.dumps({"tool": tool, "checks": checks or []}).encode() + b"\n"
        socket.write(payload)
        return socket.bytesToWrite() == 0 or socket.waitForBytesWritten(1000)
    finally:
        socket.disconnectFromServer()


class ToolNavigationServer(QObject):
    requested = Signal(str, object)
    path_requested = Signal(str, str)

    def __init__(self, parent: QObject | None = None, *, name: str | None = None) -> None:
        super().__init__(parent)
        self.name = name or navigation_server_name()
        self.server = QLocalServer(self)
        self.server.setSocketOptions(QLocalServer.SocketOption.UserAccessOption)
        self.server.newConnection.connect(self._accept)
        self._connections: dict[QLocalSocket, QTimer] = {}

    def start(self) -> bool:
        # Called only after taking the application's single-instance lock.
        QLocalServer.removeServer(self.name)
        return self.server.listen(self.name)

    def close(self) -> None:
        self.server.close()
        for socket in tuple(self._connections):
            self._drop(socket)

    def _drop(self, socket: QLocalSocket) -> None:
        timer = self._connections.pop(socket, None)
        if timer is None:
            return
        timer.stop()
        socket.abort()
        socket.deleteLater()

    def _accept(self) -> None:
        while self.server.hasPendingConnections():
            socket = self.server.nextPendingConnection()
            socket.setParent(self)
            timer = QTimer(socket)
            self._connections[socket] = timer
            socket.readyRead.connect(lambda socket=socket: self._read(socket))
            socket.disconnected.connect(lambda socket=socket: self._drop(socket))
            timer.setSingleShot(True)
            timer.timeout.connect(lambda socket=socket: self._drop(socket))
            timer.start(2000)
            self._read(socket)

    def _read(self, socket: QLocalSocket) -> None:
        if socket.bytesAvailable() > 4096:
            socket.disconnectFromServer()
            return
        if not socket.canReadLine():
            return
        try:
            payload = json.loads(bytes(socket.readLine()))
            if not isinstance(payload, dict):
                return
            if "action" in payload:
                action, path = payload["action"], payload.get("path")
                if isinstance(action, str) and action in PATH_ACTIONS and isinstance(path, str) and path and "\0" not in path:
                    socket.write(b"ok\n")
                    socket.flush()
                    # Acknowledge receipt before workbook reads or modal dialogs.
                    QTimer.singleShot(0, lambda: self.path_requested.emit(action, path))
                return
            tool, checks = payload["tool"], payload.get("checks", [])
            if isinstance(tool, str) and isinstance(checks, list) and all(isinstance(item, str) for item in checks):
                self.requested.emit(tool, checks)
        except (KeyError, TypeError, ValueError, UnicodeDecodeError):
            pass
        finally:
            socket.disconnectFromServer()
