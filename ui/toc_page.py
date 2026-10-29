# -*- coding: utf-8 -*-
"""「Word 文档快捷生成目录」工具页。

拖入一个 .docx → 在正文前插入 Word 原生目录 → 另存为「原名_带目录.docx」。
源文件全程不参与写入（实现上是复制一份、只改副本，见 core/office.add_toc_to_docx）。

动手之前先用 core/toc.scan_headings 静态扫一遍标题（纯 python-docx、不启动 Word），
一条都扫不到就直接拦住 —— 否则 Word 会给你插一个只有「未找到目录项。」的空目录，
而且它不报错，白跑一趟还看不出来。
"""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QSettings, Qt, QThread, Signal
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from core import office, toc
from core.exporter import _unique

NO_TITLE_HINT = (
    "这份文档没有使用「标题 1 / 2 / 3」样式，Word 只会插出一个空目录。\n"
    "请先在 Word 里把章节标题套上标题样式，或用「页面提取」把页面导成图片。"
)


class _ClickLabel(QLabel):
    """双击可触发回调的占位标签。"""

    def __init__(self, on_double, parent=None):
        super().__init__(parent)
        self._on_double = on_double

    def mouseDoubleClickEvent(self, event):  # noqa: N802
        if event.button() == Qt.MouseButton.LeftButton:
            self._on_double()
        else:
            super().mouseDoubleClickEvent(event)


# --------------------------------------------------------------------------- 后台任务

class _TocWorker(QThread):
    """Word 冷启动 + 大文档重排可能十几秒，不能放主线程。"""

    progress = Signal(int, int, str)
    done = Signal(bool, str, str)          # ok, 消息, 输出路径

    def __init__(self, src: str, out: str, opts: dict, parent=None):
        super().__init__(parent)
        self._src = src
        self._out = out
        self._opts = dict(opts)

    def _report(self, done: int, total: int, msg: str) -> None:
        self.progress.emit(done, total, msg)

    def run(self) -> None:  # noqa: D102
        try:
            if office.has_word():
                res = office.add_toc_to_docx(
                    self._src, self._out, progress=self._report, **self._opts
                )
                msg = f"{res['entries']} 条目录"
                if res.get("replaced"):
                    msg += f"，替换了 {res['replaced']} 个旧目录"
                self.done.emit(True, msg, str(res["out"]))
            else:
                plain = {k: v for k, v in self._opts.items() if k != "replace_existing"}
                res = office.add_toc_to_docx_pydocx(
                    self._src, self._out, progress=self._report, **plain
                )
                self.done.emit(
                    True,
                    "已写入目录域 —— 本机没装 Word，打开文件后按 Ctrl+A 再按 F9 就会生成",
                    str(res["out"]),
                )
        except Exception as exc:  # noqa: BLE001
            self.done.emit(False, str(exc), "")


# --------------------------------------------------------------------------- 页面

class TocPage(QWidget):
    """拖入 Word → 生成带目录的副本。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("Root")
        self.setAcceptDrops(True)

        self.src_path: str | None = None
        self._rows: list[tuple[int, str]] = []
        self._worker: _TocWorker | None = None
        self._settings = QSettings("MuFileTools", "TocPage")

        self._build_ui()
        self._restore_settings()
        self._refresh_view()

    # ------------------------------------------------------------------ 界面
    def _build_ui(self) -> None:
        outer = QVBoxLayout(self)
        outer.setContentsMargins(14, 12, 14, 10)
        outer.setSpacing(10)

        # 顶部标题
        head = QHBoxLayout()
        title_box = QVBoxLayout()
        title_box.setSpacing(2)
        t = QLabel("Word 文档快捷生成目录")
        t.setObjectName("Title")
        sub = QLabel("拖入 .docx，在正文前插入 Word 原生目录；原文件不动，另存为「原名_带目录.docx」")
        sub.setObjectName("Subtitle")
        title_box.addWidget(t)
        title_box.addWidget(sub)
        head.addLayout(title_box)
        head.addStretch(1)
        self.btn_choose = QPushButton("选择 Word…")
        self.btn_choose.clicked.connect(self.choose_file)
        head.addWidget(self.btn_choose)
        self.btn_clear = QPushButton("清空")
        self.btn_clear.setObjectName("Ghost")
        self.btn_clear.clicked.connect(self.clear)
        head.addWidget(self.btn_clear)
        outer.addLayout(head)

        # 主体：左文件 / 右选项
        body = QHBoxLayout()
        body.setSpacing(10)

        left = QFrame()
        left.setObjectName("Panel")
        lv = QVBoxLayout(left)
        lv.setContentsMargins(12, 12, 12, 12)
        lv.setSpacing(8)
        lsec = QLabel("待处理文档")
        lsec.setObjectName("Section")
        lv.addWidget(lsec)
        self.file_label = _ClickLabel(self.choose_file, left)
        self.file_label.setObjectName("Placeholder")
        self.file_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.file_label.setWordWrap(True)
        self.file_label.setCursor(Qt.CursorShape.PointingHandCursor)
        lv.addWidget(self.file_label, 1)
        self.file_meta = QLabel("原文件只读不动；结果另存为「原名_带目录.docx」")
        self.file_meta.setObjectName("Muted")
        self.file_meta.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lv.addWidget(self.file_meta)
        body.addWidget(left, 3)

        right = QFrame()
        right.setObjectName("Panel")
        rv = QVBoxLayout(right)
        rv.setContentsMargins(12, 12, 12, 12)
        rv.setSpacing(8)
        rsec = QLabel("目录设置")
        rsec.setObjectName("Section")
        rv.addWidget(rsec)

        depth_row = QHBoxLayout()
        depth_row.setSpacing(8)
        depth_row.addWidget(QLabel("收录到"))
        self.depth_box = QComboBox()
        for label, value in (("1 级（章）", 1), ("1-2 级", 2), ("1-3 级", 3), ("1-4 级", 4)):
            self.depth_box.addItem(label, value)
        self.depth_box.currentIndexChanged.connect(self._on_depth_changed)
        depth_row.addWidget(self.depth_box, 1)
        rv.addLayout(depth_row)

        self.opt_title = QCheckBox("在目录上方加一行「目录」标题")
        self.opt_break = QCheckBox("目录单独起页（正文从下一页开始）")
        self.opt_replace = QCheckBox("替换文档里已有的目录")
        self.opt_title.setChecked(True)
        self.opt_break.setChecked(True)
        self.opt_replace.setChecked(True)
        for box in (self.opt_title, self.opt_break, self.opt_replace):
            rv.addWidget(box)

        stat_sec = QLabel("识别到的标题")
        stat_sec.setObjectName("Section")
        rv.addWidget(stat_sec)
        self.stat_label = QLabel("还没有载入文档")
        self.stat_label.setObjectName("Muted")
        rv.addWidget(self.stat_label)

        # 警告放在预览框**上方**：预览框有 stretch 会吃掉剩余高度，
        # 摆在下面会被挤成 0 高、看不见（踩过一次）
        self.warn_label = QLabel("")
        self.warn_label.setObjectName("Warn")
        self.warn_label.setWordWrap(True)
        self.warn_label.hide()
        rv.addWidget(self.warn_label)

        self.preview = QPlainTextEdit()
        self.preview.setObjectName("Log")
        self.preview.setReadOnly(True)
        self.preview.setMinimumHeight(110)
        self.preview.setPlaceholderText("载入文档后，这里列出识别到的章节目录")
        rv.addWidget(self.preview, 1)
        body.addWidget(right, 2)

        outer.addLayout(body, 1)

        # 底栏
        bar = QFrame()
        bar.setObjectName("Panel")
        bh = QHBoxLayout(bar)
        bh.setContentsMargins(12, 10, 12, 10)
        bh.setSpacing(8)
        bh.addWidget(QLabel("输出到"))
        self.out_dir = QLineEdit()
        self.out_dir.setPlaceholderText("留空 = 与源文件同目录")
        bh.addWidget(self.out_dir, 1)
        b_browse = QPushButton("…")
        b_browse.setFixedWidth(36)
        b_browse.clicked.connect(self._choose_outdir)
        bh.addWidget(b_browse)
        self.btn_go = QPushButton("生成带目录的 Word")
        self.btn_go.setObjectName("Primary")
        self.btn_go.setMinimumHeight(36)
        self.btn_go.clicked.connect(self.generate)
        bh.addWidget(self.btn_go)
        outer.addWidget(bar)

        self.status = QLabel("就绪 — 把一个 .docx 拖到窗口任意位置，或点「选择 Word…」")
        self.status.setObjectName("Muted")
        outer.addWidget(self.status)

    # ------------------------------------------------------------------ 设置
    def _restore_settings(self) -> None:
        s = self._settings
        try:
            levels = int(s.value("levels", 3))
        except Exception:
            levels = 3
        idx = self.depth_box.findData(max(1, min(4, levels)))
        if idx >= 0:
            self.depth_box.setCurrentIndex(idx)
        self.opt_title.setChecked(s.value("with_title", "true") in (True, "true"))
        self.opt_break.setChecked(s.value("page_break", "true") in (True, "true"))
        self.opt_replace.setChecked(s.value("replace_existing", "true") in (True, "true"))
        self.out_dir.setText(str(s.value("out_dir", "") or ""))

    def _save_settings(self) -> None:
        s = self._settings
        s.setValue("levels", int(self.depth_box.currentData()))
        s.setValue("with_title", "true" if self.opt_title.isChecked() else "false")
        s.setValue("page_break", "true" if self.opt_break.isChecked() else "false")
        s.setValue("replace_existing", "true" if self.opt_replace.isChecked() else "false")
        s.setValue("out_dir", self.out_dir.text().strip())

    # ------------------------------------------------------------------ 载入
    def choose_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "选择 Word 文档", str(Path.home()), "Word 文档 (*.docx);;所有文件 (*.*)"
        )
        if path:
            self.load_path(path)

    def add_paths(self, paths: list[str]) -> None:
        """给主窗口的拖放路由用。一次只处理一份，多的给个提示。"""
        docs = [p for p in paths if Path(p).suffix.lower() in toc.DOCX_EXTS]
        if not docs:
            if paths:
                self.notify("这里只认 .docx：老的 .doc / .rtf 请先在 Word 里另存为 .docx")
            return
        self.load_path(docs[0])
        if len(docs) > 1:
            self.notify(f"已载入 {Path(docs[0]).name}；目录一次只处理一份，其余 {len(docs) - 1} 份请逐个来")

    def load_path(self, path: str) -> None:
        p = Path(path)
        self.src_path = str(p)
        suffix = p.suffix.lower()
        if suffix not in toc.DOCX_EXTS:
            self._rows = []
            self._refresh_view()
            self.warn(f"只支持 .docx（这是 {suffix or '无扩展名'}）。请在 Word 里「另存为」一份 .docx 再来。")
            return

        try:
            self._rows = toc.scan_headings(p)
        except Exception as exc:  # noqa: BLE001
            self._rows = []
            self._refresh_view()
            self.warn(f"读不了这份文档：{exc}")
            return

        self._refresh_view()
        if not self._rows:
            self.warn(NO_TITLE_HINT)
        elif self._worker is None:
            depth = int(self.depth_box.currentData())
            count = sum(toc.level_counts(self._rows, depth).values())
            self.notify(f"已载入 {p.name}：{count} 条标题将进入目录")

    def clear(self) -> None:
        self.src_path = None
        self._rows = []
        self._refresh_view()
        self.notify("就绪 — 把一个 .docx 拖到窗口任意位置，或点「选择 Word…」")

    def notify(self, text: str) -> None:
        """给外部（主窗口的拖放路由）用的提示入口。"""
        self.status.setText(text)

    def warn(self, text: str) -> None:
        self.warn_label.setText(text)
        self.warn_label.show()

    def _clear_warn(self) -> None:
        self.warn_label.clear()
        self.warn_label.hide()

    # ------------------------------------------------------------------ 刷新
    def _on_depth_changed(self) -> None:
        # 只按新深度重新过滤，不重扫文件
        self._refresh_view()
        if self.src_path and self._rows:
            depth = int(self.depth_box.currentData())
            count = sum(toc.level_counts(self._rows, depth).values())
            self.notify(f"{count} 条标题将进入目录")

    def _refresh_view(self) -> None:
        if not self.src_path:
            self.file_label.setObjectName("Placeholder")
            self.file_label.setText("把 .docx 拖到这里\n或双击选择")
            self.file_meta.setText("原文件只读不动；结果另存为「原名_带目录.docx」")
            self.stat_label.setText("还没有载入文档")
            self.preview.clear()
            self._clear_warn()
            self.btn_go.setEnabled(False)
            self._repolish(self.file_label)
            return

        p = Path(self.src_path)
        depth = int(self.depth_box.currentData())

        if self._rows:
            self.file_label.setObjectName("FileName")
            self.file_label.setText(f"{p.name}\n\n{len(self._rows)} 条标题")
        else:
            self.file_label.setObjectName("Placeholder")
            self.file_label.setText(f"{p.name}\n\n没有识别到标题样式")
        self._repolish(self.file_label)

        try:
            size_kb = p.stat().st_size / 1024
            self.file_meta.setText(f"{size_kb:.1f} KB · 原文件只读不动，结果另存为「{p.stem}_带目录.docx」")
        except Exception:
            self.file_meta.setText(f"结果另存为「{p.stem}_带目录.docx」")

        counts = toc.level_counts(self._rows, depth)
        names = {1: "一级", 2: "二级", 3: "三级", 4: "四级"}
        self.stat_label.setText(
            " · ".join(f"{names[i]} {counts[i]} 个" for i in counts)
            if self._rows
            else "0 条（这份文档没有用标题样式）"
        )

        lines = toc.preview_lines(self._rows, depth)
        self.preview.setPlainText("\n".join(lines) if lines else "（空）")

        self.btn_go.setEnabled(bool(self._rows) and self._worker is None)
        if self._rows:
            self._clear_warn()

    @staticmethod
    def _repolish(widget: QWidget) -> None:
        """objectName 变了要手动重刷样式，否则不会跟着变。"""
        style = widget.style()
        style.unpolish(widget)
        style.polish(widget)

    # ------------------------------------------------------------------ 生成
    def _choose_outdir(self) -> None:
        start = self.out_dir.text().strip()
        if not start and self.src_path:
            start = str(Path(self.src_path).parent)
        folder = QFileDialog.getExistingDirectory(self, "选择输出目录", start or str(Path.home()))
        if folder:
            self.out_dir.setText(folder)

    def _out_path(self) -> Path:
        src = Path(self.src_path or "")
        folder_text = self.out_dir.text().strip()
        folder = Path(folder_text) if folder_text else src.parent
        return _unique(folder / f"{src.stem}_带目录.docx", "rename")

    def _opts(self) -> dict:
        return {
            "levels": int(self.depth_box.currentData()),
            "with_title": self.opt_title.isChecked(),
            "page_break": self.opt_break.isChecked(),
            "replace_existing": self.opt_replace.isChecked(),
        }

    def _generate_to(self, out_path: str) -> dict:
        """同步生成到指定路径，返回 office 的结果字典（测试可直接调用）。

        没装 Word 时自动走 python-docx 域代码降级。
        """
        if not self.src_path:
            raise RuntimeError("还没有载入文档。")
        opts = self._opts()
        if office.has_word():
            return office.add_toc_to_docx(self.src_path, out_path, **opts)
        plain = {k: v for k, v in opts.items() if k != "replace_existing"}
        return office.add_toc_to_docx_pydocx(self.src_path, out_path, **plain)

    def generate(self) -> None:
        if not self.src_path:
            QMessageBox.information(self, "提示", "请先拖入或选择一个 .docx 文档。")
            return
        if not self._rows:
            QMessageBox.warning(self, "无法生成目录", NO_TITLE_HINT)
            return
        if self._worker is not None:
            return

        self._save_settings()
        try:
            target = self._out_path()
            target.parent.mkdir(parents=True, exist_ok=True)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "输出目录不可用", f"{exc}\n\n换个目录，或把文件复制到本地磁盘再试。")
            return

        self.btn_go.setEnabled(False)
        self.btn_go.setText("生成中…")
        self.notify(f"正在处理 {Path(self.src_path).name}…")

        self._worker = _TocWorker(self.src_path, str(target), self._opts(), self)
        self._worker.progress.connect(lambda _d, _t, msg: self.notify(msg))
        self._worker.done.connect(self._on_done)
        self._worker.finished.connect(self._on_worker_finished)
        self._worker.start()

    def _on_done(self, ok: bool, message: str, out_path: str) -> None:
        self.btn_go.setText("生成带目录的 Word")
        if ok:
            name = Path(out_path).name if out_path else ""
            self.notify(f"完成：{name}（{message}）")
        else:
            self.notify("生成失败")
            QMessageBox.critical(self, "生成失败", message)

    def _on_worker_finished(self) -> None:
        self._worker = None
        self.btn_go.setEnabled(bool(self._rows))

    def shutdown(self) -> None:
        self._save_settings()
        worker = self._worker
        if worker is not None:
            try:
                worker.wait(3000)
            except Exception:
                pass
            self._worker = None

    # ------------------------------------------------------------------ 拖放兜底
    def dragEnterEvent(self, event):  # noqa: N802
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):  # noqa: N802
        # 正常情况下由 ui/dnd.FileDropRouter 接管，这里只是没有路由时的兜底
        paths = [u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()]
        handler = getattr(self.window(), "handle_dropped_paths", None)
        if handler is not None:
            handler(paths)
            return
        self.add_paths(paths)
