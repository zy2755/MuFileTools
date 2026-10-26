"""界面配色与 QSS 样式：浅色 / 深色两套主题。

设计原则
--------
* 所有颜色都收在 ``PALETTES`` 的两个字典里，QSS 里只出现 ``{token}``
  占位符，切主题 = 换一本字典重新渲染，没有任何硬编码色值。
* 深色主题的 accent 比浅色更亮（#4c8dff vs #2f6fed）：低饱和的深蓝
  放在深色底上会发灰。同理 hover 态在深色下是继续提亮、浅色下是加深，
  方向相反但都符合各自的视觉逻辑。
* 除了设置 QSS，还会设置 QPalette。原生绘制的部分（文件对话框、菜单
  阴影、滚动条箭头等）不吃 QSS，只吃 palette，两边都设才不会有白块。
"""
from __future__ import annotations

import sys

MODE_LIGHT = "light"
MODE_DARK = "dark"

# ---------------------------------------------------------------- 调色板
_LIGHT = {
    "accent": "#2f6fed",          # 主色
    "accent_hover": "#1f55c4",    # 主色 hover（浅色下加深）
    "on_accent": "#ffffff",       # 落在主色上的文字
    "bg": "#f2f5fa",              # 窗口底色
    "surface": "#ffffff",         # 卡片 / 面板
    "surface_alt": "#fbfcfe",     # 凹陷区域（日志、缩略图槽）
    "surface_sunk": "#f7faff",    # 选中卡片底
    "border": "#dfe5ef",
    "border_strong": "#c3d2ee",   # hover / 强调边框
    "text": "#1b2333",
    "muted": "#6b7688",
    "hover_bg": "#eef3fe",        # 浅色高亮底
    "selected_bg": "#f2f7ff",
    "field_bg": "#ffffff",        # 输入框底（要和面板拉开层次）
    "track": "#e8edf6",           # 滑杆轨道
    "disabled_text": "#a9b2c2",
    "disabled_border": "#e9edf5",
    "primary_disabled": "#b9c9ea",
    "check_border": "#b6c0d2",
    "check_bg": "#ffffff",
    "scroll": "#cbd4e3",
    "scroll_hover": "#aebacd",
    "progress_bg": "#e8edf6",
    "log_text": "#4a5568",
    "tip_bg": "#30384a",
    "tip_text": "#ffffff",
    "menu_bg": "#ffffff",
    "menu_sel": "#eef3fe",
}

_DARK = {
    "accent": "#4c8dff",
    "accent_hover": "#6ea3ff",    # 深色下 hover 提亮
    "on_accent": "#0f1620",       # 亮蓝底上用深字，对比更稳
    "bg": "#1b1f27",
    "surface": "#242a34",
    "surface_alt": "#1f242d",
    "surface_sunk": "#2b3341",
    "border": "#333c4a",
    "border_strong": "#4a5a72",
    "text": "#e6ecf6",
    "muted": "#98a4b8",
    "hover_bg": "#2a3a55",
    "selected_bg": "#25334a",
    "field_bg": "#1c212a",        # 比面板更暗，形成凹陷感
    "track": "#3d4757",           # 滑杆轨道（比背景亮才看得见）
    "disabled_text": "#5f6a7d",
    "disabled_border": "#2c333f",
    "primary_disabled": "#3c4c68",
    "check_border": "#4d5868",
    "check_bg": "#2a313b",
    "scroll": "#3d4757",
    "scroll_hover": "#505c6e",
    "progress_bg": "#2c333f",
    "log_text": "#a7b3c7",
    "tip_bg": "#3b4554",
    "tip_text": "#f0f4fa",
    "menu_bg": "#242a34",
    "menu_sel": "#2f4a72",
}

PALETTES = {MODE_LIGHT: _LIGHT, MODE_DARK: _DARK}

# 兼容旧引用（其它模块可能 import 单个常量）
ACCENT = _LIGHT["accent"]
ACCENT_DARK = _LIGHT["accent_hover"]
BG = _LIGHT["bg"]
SURFACE = _LIGHT["surface"]
BORDER = _LIGHT["border"]
TEXT = _LIGHT["text"]
MUTED = _LIGHT["muted"]


def palette(mode: str) -> dict[str, str]:
    """取某一模式的调色板，未知模式回退浅色。"""
    return PALETTES.get(mode, _LIGHT)


# ---------------------------------------------------------------- QSS 模板
# 注意：`{{` 在这里是字面花括号（.format 转义），`{token}` 由调色板填充。
_QSS_TEMPLATE = """
* {{
    font-family: "Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI", sans-serif;
    font-size: 13px;
    color: {text};
}}
QMainWindow, QDialog, QWidget#Root {{ background: {bg}; }}
QStackedWidget {{ background: transparent; }}

/* ---------- 顶部栏 ---------- */
QFrame#TopBar {{
    background: {surface};
    border-bottom: 1px solid {border};
}}
QLabel#Brand {{ font-size: 17px; font-weight: 700; color: {accent}; }}
QPushButton#Nav {{
    background: transparent;
    border: 1px solid transparent;
    border-radius: 8px;
    padding: 6px 16px;
    color: {muted};
    font-weight: 600;
}}
QPushButton#Nav:hover {{ color: {accent}; background: {hover_bg}; }}
QPushButton#Nav:checked {{ background: {accent}; color: {on_accent}; }}
QPushButton#Nav:checked:hover {{ background: {accent_hover}; }}

/* ---------- 主题切换按钮 ---------- */
QPushButton#ThemeToggle {{
    background: transparent;
    border: 1px solid {border};
    border-radius: 8px;
    padding: 5px 11px;
    color: {muted};
    font-size: 12px;
}}
QPushButton#ThemeToggle:hover {{
    color: {accent};
    border-color: {accent};
    background: {hover_bg};
}}

/* ---------- 面板与文字层级 ---------- */
QWidget#Panel {{ background: {surface}; border-radius: 12px; border: 1px solid {border}; }}
QLabel#Title {{ font-size: 17px; font-weight: 600; }}
QLabel#Subtitle, QLabel#Muted, QLabel#Hint {{ color: {muted}; }}
QLabel#Section {{ font-weight: 600; font-size: 13px; }}
QLabel#Thumb {{
    background: {surface_alt};
    border: 1px solid {border};
    border-radius: 8px;
}}
QLabel#Placeholder {{ color: {muted}; font-size: 14px; }}

/* ---------- 按钮 ---------- */
QPushButton {{
    background: {surface};
    border: 1px solid {border};
    border-radius: 8px;
    padding: 7px 14px;
}}
QPushButton:hover {{ border-color: {accent}; color: {accent}; }}
QPushButton:pressed {{ background: {hover_bg}; }}
QPushButton:disabled {{ color: {disabled_text}; border-color: {disabled_border}; }}

QPushButton#Primary {{
    background: {accent};
    color: {on_accent};
    border: none;
    font-weight: 600;
    padding: 10px 18px;
    border-radius: 9px;
}}
QPushButton#Primary:hover {{ background: {accent_hover}; color: {on_accent}; }}
QPushButton#Primary:disabled {{ background: {primary_disabled}; color: {on_accent}; }}

QPushButton#Ghost {{ background: transparent; border: none; color: {accent}; padding: 4px 8px; }}
QPushButton#Ghost:hover {{ color: {accent_hover}; text-decoration: underline; }}

/* ---------- 输入控件 ---------- */
QComboBox, QLineEdit, QSpinBox {{
    background: {field_bg};
    border: 1px solid {border};
    border-radius: 8px;
    padding: 6px 9px;
    min-height: 20px;
}}
QComboBox:focus, QLineEdit:focus, QSpinBox:focus {{ border-color: {accent}; }}
QComboBox::drop-down {{ border: none; width: 22px; }}
QComboBox::down-arrow {{ image: {arrow_img}; width: 11px; height: 7px; }}
QComboBox QAbstractItemView {{
    background: {surface};
    border: 1px solid {border};
    selection-background-color: {selected_bg};
    selection-color: {text};
    outline: none;
}}
QSpinBox::up-button, QSpinBox::down-button {{ background: transparent; border: none; width: 16px; }}

/* ---------- 文件列表 ---------- */
QListWidget#Files {{ background: transparent; border: none; outline: none; }}
QListWidget#Files::item {{
    background: {surface};
    border: 1px solid {border};
    border-radius: 9px;
    padding: 8px 10px;
    margin: 3px 2px;
    /* 必须显式给颜色：不写的话选中行会去拿调色板的 HighlightedText
       （深色模式下那是"亮蓝底上的深字"#0f1620），压在 selected_bg 上几乎看不见 */
    color: {text};
}}
QListWidget#Files::item:selected {{
    border-color: {accent};
    background: {selected_bg};
    color: {text};
}}
QListWidget#Files::item:hover {{ border-color: {border_strong}; }}

/* ---------- 页面卡片 ---------- */
QFrame#UnitCard {{
    background: {surface};
    border: 1px solid {border};
    border-radius: 10px;
}}
QFrame#UnitCard:hover {{ border-color: {border_strong}; }}
QFrame#UnitCard[checked="true"] {{ border: 2px solid {accent}; background: {surface_sunk}; }}

/* ---------- 勾选框 ---------- */
QCheckBox {{ spacing: 6px; }}
QCheckBox::indicator {{
    width: 15px; height: 15px;
    border: 1px solid {check_border};
    border-radius: 4px;
    background: {check_bg};
}}
QCheckBox::indicator:checked {{
    background: {accent};
    border-color: {accent};
    image: {check_img};
}}
QCheckBox::indicator:hover {{ border-color: {accent}; }}
QCheckBox#AllCheck::indicator {{ border-radius: 4px; }}

/* ---------- 滑杆（圆角调节）---------- */
QSlider {{ background: transparent; }}
QSlider::groove:horizontal {{
    height: 6px;
    background: {track};
    border-radius: 3px;
}}
QSlider::sub-page:horizontal {{ background: {accent}; border-radius: 3px; }}
QSlider::handle:horizontal {{
    background: {surface};
    border: 2px solid {accent};
    width: 14px;
    margin: -6px 0;
    border-radius: 9px;
}}
QSlider::handle:horizontal:hover {{ background: {accent}; }}
QSlider::handle:horizontal:pressed {{ background: {accent_hover}; border-color: {accent_hover}; }}

/* ---------- 滚动 ---------- */
QScrollArea {{ background: transparent; border: none; }}
QScrollBar:vertical {{ background: transparent; width: 10px; margin: 2px; }}
QScrollBar::handle:vertical {{ background: {scroll}; border-radius: 5px; min-height: 30px; }}
QScrollBar::handle:vertical:hover {{ background: {scroll_hover}; }}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
QScrollBar::add-page, QScrollBar::sub-page {{ background: transparent; }}
QScrollBar:horizontal {{ background: transparent; height: 10px; margin: 2px; }}
QScrollBar::handle:horizontal {{ background: {scroll}; border-radius: 5px; min-width: 30px; }}
QScrollBar::handle:horizontal:hover {{ background: {scroll_hover}; }}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0; }}

/* ---------- 进度与日志 ---------- */
QProgressBar {{
    background: {progress_bg};
    border: none;
    border-radius: 4px;
    height: 8px;
    text-align: center;
}}
QProgressBar::chunk {{ background: {accent}; border-radius: 4px; }}

QPlainTextEdit#Log {{
    background: {surface_alt};
    border: 1px solid {border};
    border-radius: 8px;
    color: {log_text};
    font-family: "Cascadia Mono", Consolas, monospace;
    font-size: 12px;
}}

/* ---------- 菜单栏 / 菜单 ---------- */
QMenuBar {{
    background: {surface};
    color: {text};
    border-bottom: 1px solid {border};
    padding: 2px 4px;
}}
QMenuBar::item {{ background: transparent; color: {text}; padding: 4px 10px; border-radius: 6px; }}
QMenuBar::item:selected {{ background: {hover_bg}; color: {accent}; }}
QMenu {{
    background: {menu_bg};
    border: 1px solid {border};
    border-radius: 8px;
    padding: 4px;
}}
QMenu::item {{ padding: 6px 26px 6px 14px; border-radius: 6px; }}
QMenu::item:selected {{ background: {menu_sel}; color: {text}; }}
QMenu::separator {{ height: 1px; background: {border}; margin: 4px 8px; }}

/* ---------- 其它 ---------- */
QSplitter::handle {{ background: transparent; width: 8px; }}
QToolTip {{
    background: {tip_bg};
    color: {tip_text};
    border: 1px solid {border};
    padding: 5px 8px;
    border-radius: 6px;
}}
"""


def _render(mode: str, extra: dict[str, str] | None = None) -> str:
    """把调色板灌进模板。extra 用来塞图片 url 之类的运行期 token。"""
    tokens: dict[str, str] = dict(palette(mode))
    tokens["check_img"] = "none"
    tokens["arrow_img"] = "none"
    if extra:
        tokens.update(extra)
    return _QSS_TEMPLATE.format(**tokens)


def build_qss(mode: str) -> str:
    """按主题渲染完整 QSS（不含自绘图标，给外部/测试用）。"""
    return _render(mode)


# 兼容旧引用：默认浅色样式表
QSS = build_qss(MODE_LIGHT)


# ---------------------------------------------------------------- 自绘小图标
def _hex_rgb(color: str) -> tuple[int, int, int]:
    color = color.lstrip("#")
    return int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)


def _draw_tick(path) -> None:
    """白色对勾，4 倍超采样后缩小，边缘平滑。"""
    from PIL import Image, ImageDraw

    s, size = 4, 16
    big = Image.new("RGBA", (size * s, size * s), (0, 0, 0, 0))
    d = ImageDraw.Draw(big)
    pts = [(x * s, y * s) for x, y in ((3.6, 8.4), (6.7, 11.5), (12.4, 4.6))]
    d.line(pts, fill=(255, 255, 255, 255), width=int(2.2 * s), joint="curve")
    r = 1.1 * s
    for x, y in (pts[0], pts[-1]):          # 补圆头，避免端点被削平
        d.ellipse([x - r, y - r, x + r, y + r], fill=(255, 255, 255, 255))
    big.resize((size, size), Image.LANCZOS).save(path)


def _draw_arrow(path, color: str) -> None:
    """下拉框的小三角，颜色跟着主题走。"""
    from PIL import Image, ImageDraw

    s, w, h = 4, 12, 8
    big = Image.new("RGBA", (w * s, h * s), (0, 0, 0, 0))
    d = ImageDraw.Draw(big)
    d.polygon([(0, 0), (w * s, 0), (w * s // 2, h * s)], fill=(*_hex_rgb(color), 255))
    big.resize((w, h), Image.LANCZOS).save(path)


def _ensure_assets(mode: str) -> dict[str, str]:
    """生成勾号 / 箭头图标并返回 QSS 可用的 url。

    Qt 的 QSS 一旦给 ``::indicator`` 指定了 background，就不再绘制原生
    对勾了——所以勾得自己画。图片放临时目录，按主题分开命名，启动时
    检查存在性，缺了才重画。
    """
    try:
        import tempfile
        from pathlib import Path

        root = Path(tempfile.gettempdir()) / "MuFileTools" / "ui"
        root.mkdir(parents=True, exist_ok=True)

        tick = root / "check.png"
        if not tick.exists():
            _draw_tick(tick)

        arrow = root / f"arrow_{mode}.png"
        if not arrow.exists():
            _draw_arrow(arrow, palette(mode)["muted"])

        return {
            "check_img": f'url("{tick.as_posix()}")',
            "arrow_img": f'url("{arrow.as_posix()}")',
        }
    except Exception:
        # 图标生成失败就退回纯色方案，不该因此启动不了
        return {}


# ---------------------------------------------------------------- 系统检测
def system_is_dark() -> bool:
    """读取 Windows「应用模式」是否为深色（跟随系统用）。"""
    if sys.platform != "win32":
        return False
    try:
        import winreg

        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
        )
        with key:
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        return int(value) == 0
    except Exception:
        return False


def resolve(preference: str) -> str:
    """把偏好（system/light/dark）解析成实际生效的 light/dark。"""
    if preference == MODE_DARK:
        return MODE_DARK
    if preference == MODE_LIGHT:
        return MODE_LIGHT
    return MODE_DARK if system_is_dark() else MODE_LIGHT


# ---------------------------------------------------------------- 应用主题
def _build_palette(mode: str):
    from PySide6.QtGui import QColor, QPalette

    p = palette(mode)
    qp = QPalette()
    roles = {
        QPalette.ColorRole.Window: p["bg"],
        QPalette.ColorRole.WindowText: p["text"],
        QPalette.ColorRole.Base: p["surface"],
        QPalette.ColorRole.AlternateBase: p["surface_alt"],
        QPalette.ColorRole.Text: p["text"],
        QPalette.ColorRole.Button: p["surface"],
        QPalette.ColorRole.ButtonText: p["text"],
        QPalette.ColorRole.BrightText: p["text"],
        QPalette.ColorRole.Highlight: p["accent"],
        QPalette.ColorRole.HighlightedText: p["on_accent"],
        QPalette.ColorRole.ToolTipBase: p["tip_bg"],
        QPalette.ColorRole.ToolTipText: p["tip_text"],
        QPalette.ColorRole.PlaceholderText: p["muted"],
        QPalette.ColorRole.Link: p["accent"],
        QPalette.ColorRole.Mid: p["border"],
        QPalette.ColorRole.Midlight: p["border"],
        QPalette.ColorRole.Dark: p["border_strong"],
        QPalette.ColorRole.Shadow: p["border"],
    }
    for role, color in roles.items():
        qp.setColor(role, QColor(color))
    for role in (
        QPalette.ColorRole.Text,
        QPalette.ColorRole.ButtonText,
        QPalette.ColorRole.WindowText,
    ):
        qp.setColor(QPalette.ColorGroup.Disabled, role, QColor(p["disabled_text"]))
    return qp


def apply(app, mode: str) -> None:
    """把主题应用到整个应用（QSS + QPalette）。"""
    if app is None:
        return
    mode = MODE_DARK if mode == MODE_DARK else MODE_LIGHT
    app.setStyleSheet(_render(mode, _ensure_assets(mode)))
    app.setPalette(_build_palette(mode))


def applied_mode(app=None) -> str:
    """按应用当前的 QPalette 判断实际生效的是浅色还是深色。

    为什么不直接读偏好设置：偏好可能是「跟随系统」，也可能被别的代码路径
    改过样式，两者一旦不一致，自绘控件（例如拖放提示层）就会拿错配色——
    浅色界面上画出一块深色卡片。以实际调色板为准最稳。
    """
    from PySide6.QtGui import QPalette
    from PySide6.QtWidgets import QApplication

    app = app if app is not None else QApplication.instance()
    if app is None:
        return MODE_LIGHT

    c = app.palette().color(QPalette.ColorRole.Window)
    return MODE_DARK if (c.red() + c.green() + c.blue()) / 3 < 128 else MODE_LIGHT

def apply_titlebar(window, dark: bool) -> None:
    """让 Windows 标题栏跟着变深 / 变浅（Win10 2004+ 支持）。"""
    if sys.platform != "win32" or window is None:
        return
    try:
        import ctypes

        hwnd = int(window.winId())
        value = ctypes.c_int(1 if dark else 0)
        # 20 = DWMWA_USE_IMMERSIVE_DARK_MODE（新版），19 = 旧版系统
        for attr in (20, 19):
            if ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, attr, ctypes.byref(value), ctypes.sizeof(value)
            ) == 0:
                break
    except Exception:
        pass
