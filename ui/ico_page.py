# -*- coding: utf-8 -*-
"""「图片转 ICO」工具页。

选一张图片（JPG / PNG / BMP / GIF / WebP / TIFF 等），拖动滑杆调整圆角
（50% 即正圆），右侧棋盘格区域实时预览各尺寸 ICO 效果，一键导出包含
256/128/64/48/32/24/16 七档尺寸的 .ico 文件。算法在 core/iconizer.py。
"""
from __future__ import annotations

import os
from pathlib import Path

from PIL import Image
from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QCheckBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QSlider,
    QVBoxLayout,
    QWidget,
)

from core import iconizer
from ui.state import pil_to_pixmap


class _ClickLabel(QLabel):
    """双击可触发回调的图片占位标签。"""

    def __init__(self, on_double, parent=None):
        super().__init__(parent)
        self._on_double = on_double

    def mouseDoubleClickEvent(self, event):  # noqa: N802
        if event.button() == Qt.MouseButton.LeftButton:
            self._on_double()
        else:
            super().mouseDoubleClickEvent(event)


class IcoPage(QWidget):
    """选图 → 调圆角 → 实时预览 → 导出 ICO。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("Root")
        self.setAcceptDrops(True)

        self.original_small: Image.Image | None = None  # 降采样后的原图
        self.source_path: str | None = None
        self._source_size: tuple[int, int] = (0, 0)

        self._refresh_timer = QTimer(self)
        self._refresh_timer.setSingleShot(True)
        self._refresh_timer.setInterval(30)
        self._refresh_timer.timeout.connect(self._refresh)
        self._resize_timer = QTimer(self)
        self._resize_timer.setSingleShot(True)
        self._resize_timer.setInterval(120)
        self._resize_timer.timeout.connect(self._refresh)

        self._build_ui()
        self._refresh()

    # ------------------------------------------------------------------ 界面
    def _build_ui(self) -> None:
        outer = QVBoxLayout(self)
        outer.setContentsMargins(14, 12, 14, 10)
        outer.setSpacing(10)

        # 顶部标题
        head = QHBoxLayout()
        title_box = QVBoxLayout()
        title_box.setSpacing(2)
        t = QLabel("图片转 ICO")
        t.setObjectName("Title")
        sub = QLabel("把 JPG / PNG 等图片变成 Windows .ico 图标：圆角预览，一键导出七档尺寸")
        sub.setObjectName("Subtitle")
        title_box.addWidget(t)
        title_box.addWidget(sub)
        head.addLayout(title_box)
        head.addStretch(1)

        self.btn_choose = QPushButton("选择图片…")
        self.btn_choose.clicked.connect(self.choose_image)
        head.addWidget(self.btn_choose)
        outer.addLayout(head)

        # 主体：左原图 / 右 ICO 预览
        body = QHBoxLayout()
        body.setSpacing(10)

        left = QFrame()
        left.setObjectName("Panel")
        lv = QVBoxLayout(left)
        lv.setContentsMargins(12, 12, 12, 12)
        lv.setSpacing(8)
        lsec = QLabel("原图（拖进来或双击此处选择）")
        lsec.setObjectName("Section")
        lv.addWidget(lsec)
        self.src_label = _ClickLabel(self.choose_image, left)
        self.src_label.setText("把图片拖到这里\n或双击选择（JPG / PNG / BMP / GIF / WebP）")
        self.src_label.setObjectName("Placeholder")
        self.src_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.src_label.setCursor(Qt.CursorShape.PointingHandCursor)
        lv.addWidget(self.src_label, 1)
        self.src_panel = left
        body.addWidget(left, 1)

        right = QFrame()
        right.setObjectName("Panel")
        rv = QVBoxLayout(right)
        rv.setContentsMargins(12, 12, 12, 12)
        rv.setSpacing(8)
        rsec = QLabel("ICO 实时预览（棋盘格 = 透明区域）")
        rsec.setObjectName("Section")
        rv.addWidget(rsec)
        self.ico_label = QLabel("选择图片后\n这里实时预览 ICO 效果")
        self.ico_label.setObjectName("Placeholder")
        self.ico_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        rv.addWidget(self.ico_label, 1)

        smalls = QHBoxLayout()
        smalls.setSpacing(8)
        smalls.addStretch(1)
        self.small_labels: dict[int, QLabel] = {}
        for size in iconizer.PREVIEW_SIZES:
            cell = QVBoxLayout()
            cell.setSpacing(2)
            lbl = QLabel()
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl.setMinimumSize(size, size)
            self.small_labels[size] = lbl
            cap = QLabel(f"{size} × {size}")
            cap.setObjectName("Muted")
            cap.setAlignment(Qt.AlignmentFlag.AlignCenter)
            cell.addWidget(lbl)
            cell.addWidget(cap)
            holder = QWidget()
            holder.setLayout(cell)
            smalls.addWidget(holder)
        smalls.addStretch(1)
        self.smalls_holder = QWidget()
        self.smalls_holder.setLayout(smalls)
        rv.addWidget(self.smalls_holder)

        self.right_caption = QLabel("大图为导出 ICO 的 256 × 256 档，随窗口缩放；64/32/16 为真实像素")
        self.right_caption.setObjectName("Muted")
        self.right_caption.setAlignment(Qt.AlignmentFlag.AlignCenter)
        rv.addWidget(self.right_caption)
        self.ico_panel = right
        body.addWidget(right, 1)
        outer.addLayout(body, 1)

        # 底部：调整条 + 导出
        bar = QFrame()
        bar.setObjectName("Panel")
        bv = QHBoxLayout(bar)
        bv.setContentsMargins(12, 10, 12, 10)
        bv.setSpacing(10)
        bv.addWidget(QLabel("圆角："))
        self.radius_slider = QSlider(Qt.Orientation.Horizontal)
        self.radius_slider.setRange(0, 50)
        self.radius_slider.setValue(0)
        self.radius_slider.setMinimumWidth(180)
        self.radius_slider.valueChanged.connect(self._on_slider)
        bv.addWidget(self.radius_slider, 1)
        self.radius_text = QLabel("0%")
        self.radius_text.setMinimumWidth(34)
        bv.addWidget(self.radius_text)
        self.crop_check = QCheckBox("非正方形图片居中裁剪（取消则透明补边）")
        self.crop_check.setChecked(True)
        self.crop_check.toggled.connect(lambda _v: self._schedule_refresh())
        bv.addWidget(self.crop_check)
        self.btn_export = QPushButton("导出 ICO…")
        self.btn_export.setObjectName("Primary")
        self.btn_export.setMinimumHeight(36)
        self.btn_export.clicked.connect(self.export_ico)
        bv.addWidget(self.btn_export)
        outer.addWidget(bar)

        self.status = QLabel("就绪 — 拖一张图片进来，或点「选择图片…」")
        self.status.setObjectName("Muted")
        outer.addWidget(self.status)

    # ------------------------------------------------------------------ 事件
    def choose_image(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "选择图片", "", iconizer.FILE_FILTER)
        if path:
            self.load_path(path)

    def load_path(self, path: str) -> None:
        try:
            img = iconizer.load_source(path)
        except Exception as exc:  # 读取失败时如实提示，不假装成功
            QMessageBox.critical(self, "无法读取图片", f"{path}\n\n{exc}")
            return
        self._source_size = img.size
        self.original_small = iconizer.downscale(img)
        self.source_path = path
        name = os.path.basename(path)
        self.status.setText(f"已加载 {name}（{img.size[0]} × {img.size[1]} 像素），拖动「圆角」滑杆可实时预览")
        self._refresh()

    def _on_slider(self, value: int) -> None:
        self.radius_text.setText(f"{value}%")
        self._schedule_refresh()

    def _schedule_refresh(self) -> None:
        self._refresh_timer.start()

    def resizeEvent(self, event):  # noqa: N802
        super().resizeEvent(event)
        if self.original_small is not None:
            self._resize_timer.start()  # 窗口变化后延迟重绘，预览始终填满可用空间

    def dragEnterEvent(self, event):  # noqa: N802
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):  # noqa: N802
        # 拖放统一由 ui/dnd.FileDropRouter 接管，这里只是没有路由时的兜底
        paths = [u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()]
        handler = getattr(self.window(), "handle_dropped_paths", None)
        if handler is not None:
            handler(paths)
            return
        for raw in paths:
            p = Path(raw)
            if p.is_file() and p.suffix.lower() in iconizer.IMAGE_EXTS:
                self.load_path(str(p))
                return
        self.notify("拖入的文件不是可识别的图片格式")

    def notify(self, text: str) -> None:
        """给外部（主窗口的拖放路由）用的提示入口。"""
        self.status.setText(text)

    # ------------------------------------------------------------------ 渲染
    def _refresh(self) -> None:
        if self.original_small is None:
            self.src_label.clear()
            self.src_label.setText("把图片拖到这里\n或双击选择（JPG / PNG / BMP / GIF / WebP）")
            self.ico_label.clear()
            self.ico_label.setText("选择图片后\n这里实时预览 ICO 效果")
            for lbl in self.small_labels.values():
                lbl.clear()
            return

        radius = float(self.radius_slider.value())
        square = iconizer.to_square(self.original_small, crop=self.crop_check.isChecked())

        # 左：原图等比放进可用区域
        left_side = max(120, self.src_panel.height() - 76, self.src_panel.width() - 40)
        disp = iconizer.fit_contain(self.original_small, left_side)
        self.src_label.setPixmap(pil_to_pixmap(disp))

        # 右：主预览占满面板剩余高度（预留 64/32/16 小图行与说明文字）
        reserved = self.smalls_holder.sizeHint().height() + self.right_caption.sizeHint().height() + 74
        right_side = max(120, min(self.ico_panel.width() - 40, self.ico_panel.height() - reserved))
        main_icon = iconizer.on_checkerboard(
            iconizer.apply_round(iconizer.fit_square(square, right_side), radius), cell=8
        )
        self.ico_label.setPixmap(pil_to_pixmap(main_icon))

        for size in iconizer.PREVIEW_SIZES:
            icon = iconizer.on_checkerboard(
                iconizer.apply_round(iconizer.fit_square(square, size), radius), cell=4
            )
            self.small_labels[size].setPixmap(pil_to_pixmap(icon))

    # ------------------------------------------------------------------ 导出
    def _export_to(self, path: str) -> list[int]:
        """导出到指定路径，返回包含的尺寸（测试可直接调用）。"""
        assert self.original_small is not None
        square = iconizer.to_square(self.original_small, crop=self.crop_check.isChecked())
        icons = iconizer.build_icon_set(square, float(self.radius_slider.value()))
        iconizer.save_ico(icons, path)
        return sorted(icons, reverse=True)

    def export_ico(self) -> None:
        if self.original_small is None:
            QMessageBox.information(self, "提示", "请先选择一张图片")
            return
        stem = "icon"
        if self.source_path:
            stem = Path(self.source_path).stem
        default = str(Path.home() / "Desktop" / f"{stem}.ico")
        path, _ = QFileDialog.getSaveFileName(
            self, "导出 ICO", default, "ICO 图标 (*.ico);;所有文件 (*.*)"
        )
        if not path:
            return
        try:
            sizes = self._export_to(path)
        except Exception as exc:
            QMessageBox.critical(self, "导出失败", f"{path}\n\n{exc}")
            return
        sizes_text = "、".join(str(s) for s in sizes)
        self.status.setText(f"已导出 {path}")
        QMessageBox.information(self, "导出成功", f"已保存：{path}\n包含尺寸（px）：{sizes_text}")
