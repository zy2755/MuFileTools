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
from ui import settings  # noqa: E402
from ui import theme as theme_mod  # noqa: E402
from ui.shell import APP_NAME, MainWindow, _make_icon  # noqa: E402


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("MuFileTools")
    app.setApplicationDisplayName(APP_NAME)
    app.setOrganizationName("MuFileTools")
    app.setWindowIcon(_make_icon())

    font = QFont("Microsoft YaHei UI", 9)
    app.setFont(font)

    # 外观：默认跟随系统深浅色，用户在「视图」菜单里改过就以其选择为准
    preference = settings.load()
    theme_mod.apply(app, theme_mod.resolve(preference))

    office.cleanup_stale()

    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
