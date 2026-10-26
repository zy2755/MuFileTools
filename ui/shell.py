# -*- coding: utf-8 -*-
"""小睦文件转换器 · 主窗口壳：顶部切换「页面提取 / 图片转 ICO」两个工具页。"""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QActionGroup, QColor, QIcon, QPainter, QPixmap
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

from ui import settings
from ui import theme as theme_mod
from ui.extract_page import ExtractPage
from ui.ico_page import IcoPage

APP_NAME = "小睦文件转换器"


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(1280, 820)
        self.setWindowIcon(_make_icon())

        self._pref = settings.load()          # system / light / dark
        self._theme_actions: dict[str, QAction] = {}

        self.extract_page = ExtractPage()
        self.ico_page = IcoPage()

        self._build_ui()
        self._build_menu()
        self._sync_theme_ui()

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

        # 外观切换
        self.btn_theme = QPushButton()
        self.btn_theme.setObjectName("ThemeToggle")
        self.btn_theme.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_theme.clicked.connect(self._toggle_theme)
        bl.addWidget(self.btn_theme)

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

        # 视图：外观三选一（互斥）
        m_view = bar.addMenu("视图(&V)")
        group = QActionGroup(self)
        group.setExclusive(True)
        for key, label in (
            (settings.PREF_SYSTEM, "跟随系统外观"),
            (settings.PREF_LIGHT, "浅色"),
            (settings.PREF_DARK, "深色"),
        ):
            act = QAction(label, self)
            act.setCheckable(True)
            act.triggered.connect(lambda _checked=False, k=key: self._set_preference(k))
            group.addAction(act)
            m_view.addAction(act)
            self._theme_actions[key] = act
        m_view.addSeparator()
        a_toggle = QAction("切换深色 / 浅色", self)
        a_toggle.setShortcut("Ctrl+Shift+D")
        a_toggle.triggered.connect(self._toggle_theme)
        m_view.addAction(a_toggle)

        m_help = bar.addMenu("帮助(&H)")
        a_about = QAction("关于", self)
        a_about.triggered.connect(self._about)
        m_help.addAction(a_about)

        m_file = bar.addMenu("文件(&F)")
        a_quit = QAction("退出", self)
        a_quit.setShortcut("Ctrl+Q")
        a_quit.triggered.connect(self.close)
        m_file.addAction(a_quit)

    # ------------------------------------------------------------- 外观切换
    def _toggle_theme(self) -> None:
        """在深色 / 浅色之间切换。

        当前若是「跟随系统」，就按实际生效的那一套取反，并把结果固化
        成显式偏好——用户手动点过之后，就不该再被系统设置悄悄改回去。
        """
        current = theme_mod.resolve(self._pref)
        self._set_preference(
            settings.PREF_LIGHT if current == theme_mod.MODE_DARK else settings.PREF_DARK
        )

    def _set_preference(self, preference: str) -> None:
        self._pref = preference
        settings.save(preference)
        mode = theme_mod.resolve(preference)
        theme_mod.apply(QApplication.instance(), mode)
        theme_mod.apply_titlebar(self, mode == theme_mod.MODE_DARK)
        self._sync_theme_ui()

    def _sync_theme_ui(self) -> None:
        """同步按钮文案与菜单勾选。"""
        mode = theme_mod.resolve(self._pref)
        if mode == theme_mod.MODE_DARK:
            self.btn_theme.setText("☀ 浅色")
            self.btn_theme.setToolTip("切回浅色外观（Ctrl+Shift+D）")
        else:
            self.btn_theme.setText("☾ 深色")
            self.btn_theme.setToolTip("切换到深色外观（Ctrl+Shift+D）")

        act = self._theme_actions.get(self._pref)
        if act is not None and not act.isChecked():
            act.setChecked(True)

    def showEvent(self, event):  # noqa: N802
        super().showEvent(event)
        # 窗口句柄要等 show 之后才有效，标题栏配色放在这里
        theme_mod.apply_titlebar(self, theme_mod.resolve(self._pref) == theme_mod.MODE_DARK)

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
