# -*- coding: utf-8 -*-
"""小睦文件转换器 (MuFileTools) 入口。

集成两个工具：页面提取（PDF / Word / Excel）与图片转 ICO。
运行： python main.py
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from PySide6.QtGui import QFont  # noqa: E402
from PySide6.QtWidgets import QApplication  # noqa: E402

from core import office  # noqa: E402
from ui.shell import APP_NAME, MainWindow, _make_icon  # noqa: E402
from ui.theme import QSS  # noqa: E402


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("MuFileTools")
    app.setApplicationDisplayName(APP_NAME)
    app.setOrganizationName("MuFileTools")
    app.setWindowIcon(_make_icon())

    font = QFont("Microsoft YaHei UI", 9)
    app.setFont(font)
    app.setStyleSheet(QSS)

    office.cleanup_stale()

    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
