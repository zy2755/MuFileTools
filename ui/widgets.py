"""界面控件：单位卡片（一页 / 一张工作表）、缩略图网格、大图预览。"""
from __future__ import annotations

from PySide6.QtCore import QSize, Qt, Signal
from PySide6.QtGui import QFontMetrics, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from core.models import Unit

CARD_W = 186
CARD_H = 292
THUMB_W = 166
THUMB_H = 218
GAP = 12


class _Thumb(QLabel):
    clicked = Signal()
    double = Signal()

    def mousePressEvent(self, event):  # noqa: N802
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()
        super().mousePressEvent(event)

    def mouseDoubleClickEvent(self, event):  # noqa: N802
        if event.button() == Qt.MouseButton.LeftButton:
            self.double.emit()
        super().mouseDoubleClickEvent(event)


class UnitCard(QFrame):
    toggled = Signal(str, bool)
    opened = Signal(str)

    def __init__(self, unit: Unit, checked: bool = False, parent=None):
        super().__init__(parent)
        self.unit = unit
        self.setObjectName("UnitCard")
        self.setFixedSize(CARD_W, CARD_H)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(9, 9, 9, 9)
        lay.setSpacing(6)

        self.thumb = _Thumb("加载中…")
        self.thumb.setObjectName("Thumb")
        self.thumb.setFixedSize(THUMB_W, THUMB_H)
        self.thumb.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.thumb.setCursor(Qt.CursorShape.PointingHandCursor)
        lay.addWidget(self.thumb, 0, Qt.AlignmentFlag.AlignHCenter)

        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(4)
        self.check = QCheckBox()
        self.check.setChecked(checked)
        self.check.setCursor(Qt.CursorShape.PointingHandCursor)
        label = QLabel()
        label.setObjectName("Muted")
        fm = QFontMetrics(label.font())
        label.setText(fm.elidedText(unit.label, Qt.TextElideMode.ElideRight, 130))
        label.setToolTip(unit.label)
        row.addWidget(self.check)
        row.addWidget(label, 1)
        lay.addLayout(row)

        meta_bits = []
        if unit.meta.get("size"):
            w, h = unit.meta["size"]
            meta_bits.append(f"{w}×{h}")
        if "rows" in unit.meta:
            meta_bits.append(f"{unit.meta['rows']}行×{unit.meta['cols']}列")
        if unit.meta.get("visible") is False:
            meta_bits.append("隐藏")
        self.meta_label = QLabel(" · ".join(meta_bits) if meta_bits else " ")
        self.meta_label.setObjectName("Hint")
        self.meta_label.setStyleSheet("font-size: 11px;")
        lay.addWidget(self.meta_label)

        self.thumb.clicked.connect(self._toggle)
        self.thumb.double.connect(lambda: self.opened.emit(self.unit.key))
        self.check.toggled.connect(self._on_check)
        self._sync()

    # ------------------------------------------------------------------
    def _toggle(self) -> None:
        self.check.setChecked(not self.check.isChecked())

    def _on_check(self, value: bool) -> None:
        self._sync()
        self.toggled.emit(self.unit.key, value)

    def _sync(self) -> None:
        on = self.check.isChecked()
        self.setProperty("checked", "true" if on else "false")
        self.style().unpolish(self)
        self.style().polish(self)

    def set_checked(self, value: bool) -> None:
        if self.check.isChecked() != value:
            self.check.setChecked(value)

    def set_thumb(self, pix: QPixmap) -> None:
        self.thumb.setPixmap(
            pix.scaled(THUMB_W - 6, THUMB_H - 6,
                       Qt.AspectRatioMode.KeepAspectRatio,
                       Qt.TransformationMode.SmoothTransformation)
        )

    def set_error(self, msg: str) -> None:
        self.thumb.setText("无法预览")
        self.thumb.setToolTip(msg)


class PreviewGrid(QScrollArea):
    toggled = Signal(str, bool)
    opened = Signal(str)
    want_thumb = Signal(str, str)     # uid, key

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWidgetResizable(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._canvas = QWidget()
        self._canvas.setObjectName("Root")
        self._grid = QGridLayout(self._canvas)
        self._grid.setContentsMargins(16, 16, 16, 16)
        self._grid.setSpacing(GAP)
        self._grid.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        self.setWidget(self._canvas)

        self._cards: list[UnitCard] = []
        self._cols = 0
        self._uid = ""
        self._placeholder: QLabel | None = QLabel("左侧添加文件后，在这里勾选要提取的页面")
        self._placeholder.setObjectName("Placeholder")
        self._placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._grid.addWidget(self._placeholder, 0, 0)

    # ------------------------------------------------------------------
    def clear(self) -> None:
        for card in self._cards:
            card.setParent(None)
            card.deleteLater()
        self._cards = []
        self._cols = 0
        self._uid = ""
        if self._placeholder is not None:
            self._placeholder.setParent(None)
            self._placeholder.deleteLater()
        self._placeholder = QLabel("把 PDF / Word / Excel 拖到窗口任意位置\n或点左侧「＋ 添加文件」，然后在这里勾选要提取的页面")
        self._placeholder.setObjectName("Placeholder")
        self._placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._grid.addWidget(self._placeholder, 0, 0)

    def show_message(self, text: str) -> None:
        self.clear()
        if self._placeholder is not None:
            self._placeholder.setText(text)

    def set_units(self, uid: str, units: list[Unit], checked: set[str]) -> None:
        self.clear()
        self._uid = uid
        if not units:
            self.show_message("这个文件里没有可提取的页面")
            return
        if self._placeholder is not None:
            self._placeholder.setParent(None)
            self._placeholder.deleteLater()
            self._placeholder = None
        for unit in units:
            card = UnitCard(unit, checked=unit.key in checked)
            card.toggled.connect(self.toggled.emit)
            card.opened.connect(self.opened.emit)
            self._cards.append(card)
            card.thumb.setText("排队中…")
        self._layout_cards(force=True)
        for card in self._cards:
            self.want_thumb.emit(uid, card.unit.key)

    def apply_thumb(self, uid: str, key: str, pix: QPixmap) -> None:
        if uid != self._uid:
            return
        for card in self._cards:
            if card.unit.key == key:
                card.set_thumb(pix)
                return

    def apply_thumb_error(self, uid: str, key: str, msg: str) -> None:
        if uid != self._uid:
            return
        for card in self._cards:
            if card.unit.key == key:
                card.set_error(msg)
                return

    def sync_checks(self, checked: set[str]) -> None:
        for card in self._cards:
            card.set_checked(card.unit.key in checked)

    # ------------------------------------------------------------------
    def resizeEvent(self, event):  # noqa: N802
        super().resizeEvent(event)
        if self._cards:
            self._layout_cards()

    def _layout_cards(self, force: bool = False) -> None:
        if not self._cards:
            return
        avail = max(self.viewport().width() - 32, CARD_W)
        cols = max(1, (avail + GAP) // (CARD_W + GAP))
        if cols == self._cols and not force:
            return
        self._cols = cols
        while self._grid.count():
            self._grid.takeAt(0)
        for i, card in enumerate(self._cards):
            self._grid.addWidget(card, i // cols, i % cols)


class PreviewDialog(QDialog):
    """双击卡片：按 150 DPI 渲染大图，可滚动查看。"""

    def __init__(self, source, key: str, title: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(900, 760)
        # 背景交给全局 QSS 的 QDialog 规则，别在这里写死颜色，
        # 否则切到深色主题时这个窗口会孤零零一块白

        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 12, 12, 12)

        bar = QHBoxLayout()
        name = QLabel(title)
        name.setObjectName("Title")
        bar.addWidget(name)
        bar.addStretch(1)
        close = QPushButton("关闭")
        close.clicked.connect(self.accept)
        bar.addWidget(close)
        lay.addLayout(bar)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        holder = QLabel("正在渲染…")
        holder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        scroll.setWidget(holder)
        lay.addWidget(scroll, 1)

        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            pages = source.render(key, 150)
            pix = QPixmap()
            from ui.state import pil_to_pixmap
            pix = pil_to_pixmap(pages[0])
            holder.setPixmap(pix)
            holder.setText("")
        except Exception as exc:
            holder.setText(f"渲染失败：{exc}")
        finally:
            QApplication.restoreOverrideCursor()
