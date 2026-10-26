"""界面样式（QSS）与配色。"""
from __future__ import annotations

ACCENT = "#2f6fed"
ACCENT_DARK = "#1f55c4"
BG = "#f2f5fa"
SURFACE = "#ffffff"
BORDER = "#dfe5ef"
TEXT = "#1b2333"
MUTED = "#6b7688"

QSS = f"""
* {{
    font-family: "Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI", sans-serif;
    font-size: 13px;
    color: {TEXT};
}}
QMainWindow, QWidget#Root {{ background: {BG}; }}
QFrame#TopBar {{
    background: {SURFACE};
    border-bottom: 1px solid {BORDER};
}}
QLabel#Brand {{ font-size: 17px; font-weight: 700; color: {ACCENT}; }}
QPushButton#Nav {{
    background: transparent;
    border: 1px solid transparent;
    border-radius: 8px;
    padding: 6px 16px;
    color: {MUTED};
    font-weight: 600;
}}
QPushButton#Nav:hover {{ color: {ACCENT}; background: #eef3fe; }}
QPushButton#Nav:checked {{
    background: {ACCENT};
    color: #ffffff;
}}
QPushButton#Nav:checked:hover {{ background: {ACCENT_DARK}; }}
QWidget#Panel {{ background: {SURFACE}; border-radius: 12px; border: 1px solid {BORDER}; }}
QLabel#Title {{ font-size: 17px; font-weight: 600; }}
QLabel#Subtitle, QLabel#Muted, QLabel#Hint {{ color: {MUTED}; }}
QLabel#Section {{ font-weight: 600; font-size: 13px; }}
QLabel#Thumb {{
    background: #fbfcfe;
    border: 1px solid {BORDER};
    border-radius: 8px;
}}
QLabel#Placeholder {{ color: {MUTED}; font-size: 14px; }}

QPushButton {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 8px;
    padding: 7px 14px;
}}
QPushButton:hover {{ border-color: {ACCENT}; color: {ACCENT}; }}
QPushButton:pressed {{ background: #eef3fe; }}
QPushButton:disabled {{ color: #a9b2c2; border-color: #e9edf5; }}

QPushButton#Primary {{
    background: {ACCENT};
    color: #ffffff;
    border: none;
    font-weight: 600;
    padding: 10px 18px;
    border-radius: 9px;
}}
QPushButton#Primary:hover {{ background: {ACCENT_DARK}; }}
QPushButton#Primary:disabled {{ background: #b9c9ea; color: #ffffff; }}

QPushButton#Ghost {{ background: transparent; border: none; color: {ACCENT}; padding: 4px 8px; }}
QPushButton#Ghost:hover {{ color: {ACCENT_DARK}; text-decoration: underline; }}

QComboBox, QLineEdit, QSpinBox {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 8px;
    padding: 6px 9px;
    min-height: 20px;
}}
QComboBox:focus, QLineEdit:focus, QSpinBox:focus {{ border-color: {ACCENT}; }}
QComboBox::drop-down {{ border: none; width: 20px; }}
QComboBox QAbstractItemView {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    selection-background-color: #eef3fe;
    selection-color: {TEXT};
    outline: none;
}}

QListWidget#Files {{
    background: transparent;
    border: none;
    outline: none;
}}
QListWidget#Files::item {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 9px;
    padding: 8px 10px;
    margin: 3px 2px;
}}
QListWidget#Files::item:selected {{
    border-color: {ACCENT};
    background: #f2f7ff;
}}
QListWidget#Files::item:hover {{ border-color: #c3d2ee; }}

QFrame#UnitCard {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 10px;
}}
QFrame#UnitCard:hover {{ border-color: #b9c9ea; }}
QFrame#UnitCard[checked="true"] {{
    border: 2px solid {ACCENT};
    background: #f7faff;
}}

QCheckBox {{ spacing: 6px; }}
QCheckBox::indicator {{
    width: 15px; height: 15px;
    border: 1px solid #b6c0d2;
    border-radius: 4px;
    background: #ffffff;
}}
QCheckBox::indicator:checked {{
    background: {ACCENT};
    border-color: {ACCENT};
}}
QCheckBox#AllCheck::indicator {{ border-radius: 4px; }}

QScrollArea {{ background: transparent; border: none; }}
QScrollBar:vertical {{ background: transparent; width: 10px; margin: 2px; }}
QScrollBar::handle:vertical {{ background: #cbd4e3; border-radius: 5px; min-height: 30px; }}
QScrollBar::handle:vertical:hover {{ background: #aebacd; }}
QScrollBar::add-line, QScrollBar::sub-line {{ height: 0; }}
QScrollBar:horizontal {{ background: transparent; height: 10px; margin: 2px; }}
QScrollBar::handle:horizontal {{ background: #cbd4e3; border-radius: 5px; min-width: 30px; }}

QProgressBar {{
    background: #e8edf6;
    border: none;
    border-radius: 4px;
    height: 8px;
    text-align: center;
}}
QProgressBar::chunk {{ background: {ACCENT}; border-radius: 4px; }}

QPlainTextEdit#Log {{
    background: #fbfcfe;
    border: 1px solid {BORDER};
    border-radius: 8px;
    color: #4a5568;
    font-family: "Cascadia Mono", Consolas, monospace;
    font-size: 12px;
}}

QSplitter::handle {{ background: transparent; width: 8px; }}
QToolTip {{
    background: #30384a; color: #ffffff; border: none;
    padding: 5px 8px; border-radius: 6px;
}}
"""
