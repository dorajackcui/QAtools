"""Bounded log queue with incremental GUI updates, independent of worker signals."""

from collections import deque
from datetime import datetime
from threading import Lock

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication, QDialog, QHBoxLayout, QPlainTextEdit, QPushButton, QVBoxLayout



class LogBuffer:
    def __init__(self, capacity=10000):
        self._items = deque(maxlen=capacity)
        self._lock = Lock()
        self._dropped = 0

    def append(self, message):
        # Bound even a single very long error or conflict message before queuing it.
        message = str(message)
        if len(message) > 4000:
            message = message[:4000] + " …（此条日志已截断）"
        line = f"[{datetime.now():%H:%M:%S}] {message}"
        with self._lock:
            if len(self._items) == self._items.maxlen:
                self._dropped += 1
            self._items.append(line)

    def take(self, limit=100):
        with self._lock:
            lines = []
            if self._dropped:
                lines.append(f"日志量较大，已省略 {self._dropped} 条较早记录。")
                self._dropped = 0
            # Bound both record count and text per tick, even with multiline input.
            size = 0
            while self._items and len(lines) < limit and size < 32000:
                line = self._items.popleft()
                lines.append(line)
                size += len(line)
            return lines

    def clear(self):
        with self._lock:
            self._items.clear()
            self._dropped = 0


class OperationLogDialog(QDialog):
    """Modeless dialog. Worker callbacks only touch LogBuffer, never Qt widgets."""

    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("运行日志")
        self.setModal(False)
        self.resize(740, 420)
        self.setMinimumSize(480, 280)
        self.buffer = LogBuffer()
        layout = QVBoxLayout(self)
        self.text = QPlainTextEdit()
        self.text.setReadOnly(True)
        self.text.setMaximumBlockCount(10000)
        self.text.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self.text.setPlaceholderText("运行工具后，可在此查看逐文件处理详情。")
        layout.addWidget(self.text, 1)
        row = QHBoxLayout()
        self.copy_button = QPushButton("复制日志")
        self.copy_button.clicked.connect(lambda: QApplication.clipboard().setText(self.text.toPlainText()))
        row.addWidget(self.copy_button)
        row.addStretch(1)
        close = QPushButton("关闭")
        close.clicked.connect(self.close)
        row.addWidget(close)
        layout.addLayout(row)
        self.timer = QTimer(self)
        self.timer.setInterval(100)
        self.timer.timeout.connect(self.flush)
        self.running = False

    def start_run(self, message=""):
        self.buffer.clear()
        self.text.clear()
        if message:
            self.buffer.append(message)
        self.running = True
        self.timer.start()

    def finish_run(self, message=""):
        if message:
            self.buffer.append(message)
        self.running = False
        self.timer.start()

    def open_logs(self):
        self.show()
        self.raise_()
        self.activateWindow()
        self.timer.start()
        self.flush()

    def flush(self):
        # Keep each GUI tick small; never drain an unbounded worker backlog here.
        lines = self.buffer.take()
        if not lines:
            if not self.running:
                self.timer.stop()
            return
        scrollbar = self.text.verticalScrollBar()
        follow = scrollbar.value() >= scrollbar.maximum() - 2
        position = scrollbar.value()
        self.text.appendPlainText("\n".join(lines))
        scrollbar.setValue(scrollbar.maximum() if follow else position)
