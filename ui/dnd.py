# -*- coding: utf-8 -*-
"""全局文件拖放：拖到窗口任意位置都能用。

为什么不在各页面里写 dragEnterEvent / dropEvent 就够
--------------------------------------------------
Qt 的拖放事件只投递给「鼠标正下方那个控件」，而且只要祖先里有人
acceptDrops 就算数。问题在于有些控件**天生就接受拖放**：

  · QLineEdit      拖文字进来会插进输入框（「输出目录」那一栏就是）
  · QPlainTextEdit 日志框同理
  · QComboBox      下拉框也会吞掉
  · 顶栏的 QPushButton / QLabel 根本不在页面里，拖上去没人管

结果就是「拖到中间空白处能用，拖到输入框上变成一串路径，拖到顶栏没反应」，
看起来像是拖放根本没做。所以这里把「拖本地文件」这件事提到
QApplication 层统一拦截：事件过滤器比任何控件的处理都早，先抢下来再说。

约定：
  · 只有「带本地文件 / 文件夹的拖放」会被接管，拖别的东西（选中的文字等）
    原样放行，不改变原有行为
  · 只接管落在本窗口内的拖放，不影响弹出对话框
"""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QEvent, QObject, QRectF, Qt
from PySide6.QtGui import QColor, QFont, QFontMetrics, QPainter, QPen
from PySide6.QtWidgets import QApplication, QWidget


def local_paths(mime) -> list[str]:
    """取出 mime 里的本地路径（文件和文件夹都算），没有就返回空列表。"""
    if not mime.hasUrls():
        return []
    return [u.toLocalFile() for u in mime.urls() if u.isLocalFile()]


def collect_by_ext(raw_paths: list[str], exts: set[str] | dict) -> list[str]:
    """把路径筛成指定扩展名的文件列表；文件夹会递归展开。

    拖文件夹是很自然的操作（比如整个「合同」文件夹拖进来），所以这里
    顺手支持了——原来是点「添加文件夹」按钮才有的能力。
    """
    found: list[str] = []
    for raw in raw_paths:
        p = Path(raw)
        if p.is_dir():
            found.extend(
                str(f) for f in sorted(p.rglob("*"))
                if f.is_file() and f.suffix.lower() in exts
            )
        elif p.is_file() and p.suffix.lower() in exts:
            found.append(str(p))
    return found


def title_font_for(base: QFont) -> QFont:
    """在基础字体上放大出一档标题字体。

    Qt 的字体有两种度量方式，同一个 QFont 只会有一种有效：
      · 用磅设的（pointSizeF() > 0）
      · 用像素设的（pixelSize() > 0，此时 pointSizeF() 返回 -1）
    如果不管三七二十一按磅加 5，遇到像素字体就会算出 -1+5=4pt —— 6 像素高，
    字会糊成一团（这个坑当场踩过）。
    """
    font = QFont(base)
    if font.pointSizeF() > 0:
        font.setPointSizeF(font.pointSizeF() + 5)
    else:
        base_px = font.pixelSize()
        font.setPixelSize((base_px if base_px > 0 else 13) + 7)
    font.setBold(True)
    return font


class DropOverlay(QWidget):
    """拖拽悬停时盖在窗口上的提示层。

    鼠标事件全部穿透（WA_TransparentForMouseEvents），它只是画一层淡淡的
    高亮 + 一句「松手会发生什么」，让用户在松手前就知道结果。
    """

    def __init__(self, parent: QWidget, palette_getter):
        super().__init__(parent)
        self._palette_getter = palette_getter
        self._title = ""
        self._sub = ""
        # 不透传鼠标事件，否则盖上去之后拖放事件就找不到下面的控件了
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        # 不画系统底色，让下面的界面透出来
        self.setAttribute(Qt.WidgetAttribute.WA_NoSystemBackground, True)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self.hide()

    # ------------------------------------------------------------------
    def show_hint(self, title: str, sub: str = "") -> None:
        self._title = title
        self._sub = sub
        parent = self.parentWidget()
        if parent is not None:
            self.setGeometry(parent.rect())
        self.show()
        self.raise_()
        self.update()

    def hide_hint(self) -> None:
        if self.isVisible():
            self.hide()

    # ------------------------------------------------------------------
    def paintEvent(self, event):  # noqa: N802
        pal = self._palette_getter()
        accent = QColor(pal["accent"])
        text_color = QColor(pal["text"])
        muted = QColor(pal["muted"])
        surface = QColor(pal["surface"])
        surface.setAlpha(246)

        p = QPainter(self)
        p.setRenderHint(QPainter.RenderHint.Antialiasing)

        # 整块淡淡的着色，让用户知道"整个窗口都收文件"
        veil = QColor(accent)
        veil.setAlpha(32)
        p.fillRect(self.rect(), veil)

        # 虚线边框
        border = QColor(accent)
        border.setAlpha(150)
        pen = QPen(border)
        pen.setWidth(2)
        pen.setStyle(Qt.PenStyle.DashLine)
        p.setPen(pen)
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawRoundedRect(QRectF(self.rect()).adjusted(10, 10, -10, -10), 16, 16)

        if not self._title:
            return

        # 中间的文字卡片
        title_font = title_font_for(self.font())
        sub_font = QFont(self.font())

        fm_title = QFontMetrics(title_font)
        fm_sub = QFontMetrics(sub_font)
        title_w = fm_title.horizontalAdvance(self._title)
        sub_w = fm_sub.horizontalAdvance(self._sub) if self._sub else 0
        box_w = max(title_w, sub_w) + 64
        box_h = (fm_title.height() + 10) + (fm_sub.height() + 6 if self._sub else 0) + 34

        cx = self.width() / 2
        cy = self.height() / 2
        box = QRectF(cx - box_w / 2, cy - box_h / 2, box_w, box_h)

        p.setPen(Qt.PenStyle.NoPen)
        p.setBrush(surface)
        p.drawRoundedRect(box, 14, 14)
        p.setPen(QPen(border, 1))
        p.setBrush(Qt.BrushStyle.NoBrush)
        p.drawRoundedRect(box, 14, 14)

        p.setFont(title_font)
        p.setPen(text_color)
        p.drawText(
            QRectF(box.left(), box.top() + 14, box.width(), fm_title.height() + 4),
            Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter,
            self._title,
        )
        if self._sub:
            p.setFont(sub_font)
            p.setPen(muted)
            p.drawText(
                QRectF(box.left(), box.top() + 16 + fm_title.height() + 2, box.width(), fm_sub.height() + 4),
                Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter,
                self._sub,
            )
        p.end()


class FileDropRouter(QObject):
    """装在 QApplication 上的事件过滤器，统一处理「把文件拖进窗口」。

    调用方只需要提供两个回调：
      · hint_provider(paths) -> (标题, 副标题) —— 悬停时显示什么
      · handler(paths)                        —— 松手后干什么
    """

    def __init__(self, window: QWidget, hint_provider, handler, overlay_palette):
        super().__init__(window)
        self._window = window
        self._hint = hint_provider
        self._handle = handler
        self.overlay = DropOverlay(window, overlay_palette)
        self._active = False
        app = QApplication.instance()
        if app is not None:
            app.installEventFilter(self)

    # ------------------------------------------------------------------
    def detach(self) -> None:
        app = QApplication.instance()
        if app is not None:
            app.removeEventFilter(self)

    def _inside_window(self, obj) -> bool:
        return isinstance(obj, QWidget) and obj.window() is self._window

    def eventFilter(self, obj, event):  # noqa: N802
        etype = event.type()
        if etype not in (
            QEvent.Type.DragEnter,
            QEvent.Type.DragMove,
            QEvent.Type.DragLeave,
            QEvent.Type.Drop,
        ):
            return False
        # 只管本窗口内的拖放，别去干涉文件对话框之类
        if not self._inside_window(obj):
            return False

        if etype == QEvent.Type.DragLeave:
            if self._active:
                self._active = False
                self.overlay.hide_hint()
            return False   # 让控件自己也收到，方便未来做行内高亮

        paths = local_paths(event.mimeData())
        if not paths:
            return False   # 不是拖文件（比如拖选中的文字），原样放行

        if etype in (QEvent.Type.DragEnter, QEvent.Type.DragMove):
            title, sub = self._hint(paths)
            if not self._active or (title, sub) != (self.overlay._title, self.overlay._sub):
                self.overlay.show_hint(title, sub)
            self._active = True
            event.acceptProposedAction()
            return True    # 抢下来，别再让 QLineEdit 去插路径文字

        # Drop
        self._active = False
        self.overlay.hide_hint()
        self._handle(paths)
        event.acceptProposedAction()
        return True

    def sync_geometry(self) -> None:
        if self.overlay.isVisible() and self._window is not None:
            self.overlay.setGeometry(self._window.rect())
