"""外观偏好的读写。

存在 QSettings 里（Windows 下是注册表 HKCU\\Software\\MuFileTools），
不往用户目录丢配置文件，卸载后也留不下垃圾。

取值：
    "system"  跟随 Windows 的应用深浅色（默认）
    "light"   强制浅色
    "dark"    强制深色
"""
from __future__ import annotations

from PySide6.QtCore import QSettings

KEY = "appearance/theme"
PREF_SYSTEM = "system"
PREF_LIGHT = "light"
PREF_DARK = "dark"
VALID = (PREF_SYSTEM, PREF_LIGHT, PREF_DARK)
DEFAULT = PREF_SYSTEM


def load() -> str:
    """读取外观偏好，非法值回退到跟随系统。"""
    value = QSettings().value(KEY, DEFAULT)
    value = str(value) if value is not None else DEFAULT
    return value if value in VALID else DEFAULT


def save(preference: str) -> None:
    """保存外观偏好。"""
    s = QSettings()
    s.setValue(KEY, preference if preference in VALID else DEFAULT)
    s.sync()
