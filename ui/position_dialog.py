# -*- coding: utf-8 -*-
"""「目录插在哪一页之后」对话框。

为什么需要它：封面识别是**猜**的。绝大多数文档靠「第一个正文标题之前」能猜对，
但下面这几种情况任何启发式都会错，只有人知道该插哪：

  · 封面和正文挤在同一页（前面还压着标题页、目录页）
  · 封面本身就套了「标题 1」样式
  · 文档开头有一页「评阅表 / 任务书 / 保密声明」，它既不是封面也不是正文

界面上不让人数页码，而是把「每页开头是什么字」列出来 ——
看见「第 1 页 · 调查研究报告」就知道该选第 1 页。

分页必须靠 Word 排版算出来（python-docx 读不出分页位置），
所以探测放在后台线程里跑，别把界面卡住。
"""
from __future__ import annotations

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

from core import office


class _PageWorker(QThread):
    """后台读分页信息 —— 要起一个 Word 进程，几百毫秒到几秒。"""

    done = Signal(bool, list, str)

    def __init__(self, src: str, parent=None):
        super().__init__(parent)
        self._src = src

    def run(self) -> None:  # noqa: D102
        try:
            pages = office.list_pages(self._src)
            if not pages:
                self.done.emit(False, [], "读不出分页信息（没装 Word，或文档打不开）")
            else:
                self.done.emit(True, pages, "")
        except Exception as exc:  # noqa: BLE001
            self.done.emit(False, [], str(exc))


class PositionDialog(QDialog):
    """列出每页的页首文字，让人选「目录插在这一页之后」。"""

    def __init__(self, src: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("选择目录位置")
        self.resize(760, 560)

        self._src = src
        self._pages: list[dict] = []
        self._page = 0
        self._worker: _PageWorker | None = None

        self._build_ui()
        self._start()

    # ------------------------------------------------------------------ 界面
    def _build_ui(self) -> None:
        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 12, 14, 12)
        lay.setSpacing(10)

        t = QLabel("目录插在哪一页之后")
        t.setObjectName("Title")
        lay.addWidget(t)

        sub = QLabel(
            "下面列出每一页开头的文字，选一页，目录就插在它后面。"
            "页码按「加目录之前」的排版算 —— 就是你手上这份稿子现在的分页。"
        )
        sub.setObjectName("Subtitle")
        sub.setWordWrap(True)
        lay.addWidget(sub)

        self.table = QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["页", "这一页开头是什么"])
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setShowGrid(False)
        head = self.table.horizontalHeader()
        head.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(0, 68)
        head.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table.itemSelectionChanged.connect(self._on_pick)
        # 不要焦点框：选中行已经有整行底色，再叠一个焦点矩形会在「页」那一格
        # 画出一块颜色相反的方块（浅色下白底蓝边、深色下蓝底白边），很扎眼
        self.table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        lay.addWidget(self.table, 1)

        self.hint = QLabel("正在读取分页…（要启动一下 Word，稍等）")
        self.hint.setObjectName("Muted")
        self.hint.setWordWrap(True)
        lay.addWidget(self.hint)

        bar = QHBoxLayout()
        bar.setSpacing(8)
        bar.addStretch(1)
        cancel = QPushButton("取消")
        cancel.clicked.connect(self.reject)
        bar.addWidget(cancel)
        self.ok = QPushButton("就这一页")
        self.ok.setObjectName("Primary")
        self.ok.setEnabled(False)
        self.ok.clicked.connect(self.accept)
        bar.addWidget(self.ok)
        lay.addLayout(bar)

    # ------------------------------------------------------------------ 数据
    def _start(self) -> None:
        self._worker = _PageWorker(self._src, self)
        self._worker.done.connect(self._on_pages)
        self._worker.start()

    def _on_pages(self, ok: bool, pages: list, msg: str) -> None:
        self._worker = None
        if not ok:
            self.hint.setText(msg or "读不出分页信息")
            return
        self._pages = list(pages)
        self.table.setRowCount(len(self._pages))
        for row, item in enumerate(self._pages):
            page_item = QTableWidgetItem(str(item.get("page", row + 1)))
            page_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(row, 0, page_item)
            text = (item.get("text") or "").strip() or "（这一页是空白的）"
            self.table.setItem(row, 1, QTableWidgetItem(text))
        self.table.selectRow(0)
        self.hint.setText(
            f"共 {len(self._pages)} 页。选一页，目录会插在这一页之后、"
            "下一页之前 —— 比如封面单独占第 1 页，就选第 1 页。"
        )
        self.ok.setEnabled(True)

    def _on_pick(self) -> None:
        row = self.table.currentRow()
        if 0 <= row < len(self._pages):
            self._page = int(self._pages[row].get("page", row + 1))
            self.ok.setEnabled(True)

    # ------------------------------------------------------------------ 结果
    def page(self) -> int:
        """选中的页码（目录插在这一页之后）；没选返回 0。"""
        return self._page

    def shutdown(self) -> None:
        """窗口关闭时把后台线程收掉，别留着 Word 进程。"""
        if self._worker is not None:
            try:
                self._worker.done.disconnect()
            except Exception:
                pass
            self._worker.wait(8000)
            self._worker = None

    def closeEvent(self, event) -> None:  # noqa: N802
        self.shutdown()
        super().closeEvent(event)
