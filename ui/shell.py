# -*- coding: utf-8 -*-
"""小睦文件转换器 · 主窗口壳：顶部切换「页面提取 / 图片转 ICO」两个工具页。"""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QColor, QIcon, QPainter, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from ui.extract_page import ExtractPage
from ui.ico_page import IcoPage

APP_NAME = "小睦文件转换器"


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(1280, 820)
        self.setWindowIcon(_make_icon())

        self.extract_page = ExtractPage()
        self.ico_page = IcoPage()

        self._build_ui()
        self._build_menu()

    # ------------------------------------------------------------------ 界面
    def _build_ui(self) -> None:
        central = QWidget()
        central.setObjectName("Root")
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # 顶栏：品牌 + 工具切换
        bar = QFrame()
        bar.setObjectName("TopBar")
        bl = QHBoxLayout(bar)
        bl.setContentsMargins(14, 8, 14, 8)
        bl.setSpacing(12)

        brand = QLabel(APP_NAME)
        brand.setObjectName("Brand")
        bl.addWidget(brand)
        brand_sub = QLabel("多格式文档提取 · 图片转 ICO")
        brand_sub.setObjectName("Muted")
        bl.addWidget(brand_sub)
        bl.addStretch(1)

        self.btn_extract = self._nav_button("页面提取")
        self.btn_ico = self._nav_button("图片转 ICO")
        self.btn_extract.setCheckable(True)
        self.btn_ico.setCheckable(True)
        self.btn_extract.setChecked(True)
        self.btn_extract.clicked.connect(lambda: self._switch(0))
        self.btn_ico.clicked.connect(lambda: self._switch(1))
        bl.addWidget(self.btn_extract)
        bl.addWidget(self.btn_ico)
        outer.addWidget(bar)

        # 工具页
        self.stack = QStackedWidget()
        self.stack.addWidget(self.extract_page)
        self.stack.addWidget(self.ico_page)
        outer.addWidget(self.stack, 1)

    def _nav_button(self, text: str) -> QPushButton:
        b = QPushButton(text)
        b.setObjectName("Nav")
        b.setCursor(Qt.CursorShape.PointingHandCursor)
        return b

    def _switch(self, index: int) -> None:
        self.stack.setCurrentIndex(index)
        self.btn_extract.setChecked(index == 0)
        self.btn_ico.setChecked(index == 1)

    def _build_menu(self) -> None:
        bar = self.menuBar()
        m_help = bar.addMenu("帮助(&H)")
        a_about = QAction("关于", self)
        a_about.triggered.connect(self._about)
        m_help.addAction(a_about)

        m_file = bar.addMenu("文件(&F)")
        a_quit = QAction("退出", self)
        a_quit.setShortcut("Ctrl+Q")
        a_quit.triggered.connect(self.close)
        m_file.addAction(a_quit)

    def _about(self) -> None:
        QMessageBox.about(
            self, f"关于 {APP_NAME}",
            f"{APP_NAME} (MuFileTools)\n\n"
            "· 页面提取：PDF / Word 按页提取，Excel 按工作表提取，\n"
            "  可存为 PNG / JPG / PDF / Word / Excel。\n"
            "· 图片转 ICO：圆角实时预览，一键导出七档尺寸的 .ico 图标。\n\n"
            "渲染引擎：PyMuPDF + Microsoft Office（未安装时降级 LibreOffice / openpyxl）\n"
            "图标引擎：Pillow（逐档缩放 + 抗锯齿圆角）",
        )

    def closeEvent(self, event):  # noqa: N802
        self.extract_page.shutdown()
        super().closeEvent(event)


def _make_icon() -> QIcon:
    """优先用项目自带的 app_icon.ico，缺失时用代码画一个兜底图标。"""
    from pathlib import Path

    ico = Path(__file__).resolve().parents[1] / "app_icon.ico"
    if ico.exists():
        icon = QIcon(str(ico))
        if not icon.isNull():
            return icon
    pix = QPixmap(64, 64)
    pix.fill(QColor(0, 0, 0, 0))
    p = QPainter(pix)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)
    p.setBrush(QColor("#2f6fed"))
    p.setPen(Qt.PenStyle.NoPen)
    p.drawRoundedRect(4, 6, 56, 52, 10, 10)
    p.setBrush(QColor("#ffffff"))
    p.drawRoundedRect(14, 14, 22, 26, 3, 3)
    p.setBrush(QColor("#bcd3ff"))
    p.drawRoundedRect(30, 24, 20, 16, 3, 3)
    p.end()
    return QIcon(pix)
