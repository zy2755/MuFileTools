"""界面状态与后台线程：文件载入、缩略图生成。

所有 Office / PyMuPDF 的重活都在这两个 QThread 里做，主线程只负责画界面。
"""
from __future__ import annotations

import queue
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from PIL import Image
from PySide6.QtCore import QThread, Signal
from PySide6.QtGui import QImage, QPixmap

from core.models import SUPPORTED_EXTS, Unit
from core.sources import open_source


def pil_to_qimage(img: Image.Image) -> QImage:
    rgb = img.convert("RGB")
    data = rgb.tobytes("raw", "RGB")
    return QImage(data, rgb.width, rgb.height, rgb.width * 3, QImage.Format.Format_RGB888).copy()


def pil_to_pixmap(img: Image.Image) -> QPixmap:
    return QPixmap.fromImage(pil_to_qimage(img))


@dataclass
class DocEntry:
    path: Path
    uid: str
    source: Any = None
    units: list[Unit] = field(default_factory=list)
    checked: set[str] = field(default_factory=set)
    error: str | None = None
    loading: bool = True
    thumbs: dict[str, QPixmap] = field(default_factory=dict)

    @property
    def name(self) -> str:
        return self.path.name

    @property
    def kind(self) -> str:
        return SUPPORTED_EXTS.get(self.path.suffix.lower(), "pdf")

    @property
    def kind_label(self) -> str:
        return {"pdf": "PDF", "word": "Word", "excel": "Excel"}.get(self.kind, self.kind)


class Loader(QThread):
    """串行解析待打开文件。Word/Excel 依赖 COM，串行最稳。"""

    loaded = Signal(object)          # DocEntry
    failed = Signal(object, str)     # DocEntry, message
    busy = Signal(object, str)       # DocEntry, 提示文本

    def __init__(self, parent=None):
        super().__init__(parent)
        self._q: queue.Queue[DocEntry] = queue.Queue()
        self._stop = threading.Event()

    def enqueue(self, entry: DocEntry) -> None:
        self._q.put(entry)

    def stop(self) -> None:
        self._stop.set()

    def run(self) -> None:  # noqa: D102
        while not self._stop.is_set():
            try:
                entry = self._q.get(timeout=0.2)
            except queue.Empty:
                continue
            try:
                entry.source = open_source(entry.path, progress=lambda m: self.busy.emit(entry, m))
                entry.units = entry.source.units()
                entry.loading = False
                # 单位不多的文件默认全选，多页文件默认不选，避免误导出几百页
                if len(entry.units) <= 40:
                    entry.checked = {u.key for u in entry.units}
                self.loaded.emit(entry)
            except Exception as exc:
                entry.loading = False
                entry.error = str(exc)
                self.failed.emit(entry, str(exc))


class Thumbnailer(QThread):
    """按需生成缩略图，避免一次性渲染几百页卡死界面。"""

    ready = Signal(str, str, QPixmap)   # uid, key, pixmap
    failed = Signal(str, str, str)      # uid, key, message

    def __init__(self, parent=None):
        super().__init__(parent)
        self._q: queue.Queue[tuple] = queue.Queue()
        self._stop = threading.Event()

    def request(self, uid: str, source, key: str, width: int = 300) -> None:
        self._q.put((uid, source, key, width))

    def stop(self) -> None:
        self._stop.set()

    def run(self) -> None:  # noqa: D102
        while not self._stop.is_set():
            try:
                uid, source, key, width = self._q.get(timeout=0.2)
            except queue.Empty:
                continue
            try:
                img = source.thumbnail(key, width)
                self.ready.emit(uid, key, pil_to_pixmap(img))
            except Exception as exc:
                self.failed.emit(uid, key, str(exc))
