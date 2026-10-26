# -*- coding: utf-8 -*-
"""「页面提取」工具页（原 PagePull 主窗口，整合为小睦文件转换器的一个工具页）。"""
from __future__ import annotations

import uuid
from pathlib import Path

from PySide6.QtCore import QSettings, QSize, Qt, QThread, Signal
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QSplitter,
    QVBoxLayout,
    QWidget,
)

from core import office
from core.exporter import run_jobs
from core.models import SUPPORTED_EXTS, ExportSpec, Job
from ui.state import DocEntry, Loader, Thumbnailer
from ui.widgets import PreviewDialog, PreviewGrid


# --------------------------------------------------------------------------- 后台导出

class ExportWorker(QThread):
    progress = Signal(int, int, str)
    done = Signal(list)
    failed = Signal(str)

    def __init__(self, jobs: list[Job], spec: ExportSpec, parent=None):
        super().__init__(parent)
        self._jobs = jobs
        self._spec = spec

    def run(self) -> None:  # noqa: D102
        try:
            results = run_jobs(self._jobs, self._spec, progress=lambda d, t, m: self.progress.emit(d, t, m))
            self.done.emit(results)
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(str(exc))


class CapsWorker(QThread):
    ready = Signal(dict)

    def run(self) -> None:  # noqa: D102
        try:
            self.ready.emit(office.capabilities())
        except Exception:
            self.ready.emit({"word": False, "excel": False, "soffice": False})


# --------------------------------------------------------------------------- 工具页

class ExtractPage(QWidget):
    """从 PDF / Word 抽取指定页，从 Excel 抽取指定工作表，一键存为图片或文件。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)

        self.entries: dict[str, DocEntry] = {}
        self.order: list[str] = []
        self.current_uid: str | None = None
        self._export_worker: ExportWorker | None = None

        self.settings = QSettings("MuFileTools", "PageExtract")

        self.loader = Loader(self)
        self.loader.loaded.connect(self._on_loaded)
        self.loader.failed.connect(self._on_load_failed)
        self.loader.busy.connect(self._on_busy)
        self.loader.start()

        self.thumbs = Thumbnailer(self)
        self.thumbs.ready.connect(self._on_thumb_ready)
        self.thumbs.failed.connect(self._on_thumb_failed)
        self.thumbs.start()

        self._build_ui()
        self._build_shortcuts()
        self._restore_settings()

        self.caps_worker = CapsWorker(self)
        self.caps_worker.ready.connect(self._on_caps)
        self.caps_worker.start()

    # ------------------------------------------------------------------ 构建界面
    def _build_ui(self) -> None:
        root = QWidget()
        root.setObjectName("Root")
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(root)
        lay = QVBoxLayout(root)
        lay.setContentsMargins(14, 12, 14, 10)
        lay.setSpacing(10)

        # 顶部标题
        head = QHBoxLayout()
        title_box = QVBoxLayout()
        title_box.setSpacing(2)
        t = QLabel("页面提取")
        t.setObjectName("Title")
        sub = QLabel("从 PDF / Word 抽取指定页，从 Excel 抽取指定工作表，一键存为图片或文件")
        sub.setObjectName("Subtitle")
        title_box.addWidget(t)
        title_box.addWidget(sub)
        head.addLayout(title_box)
        head.addStretch(1)
        self.caps_label = QLabel("正在检测 Office…")
        self.caps_label.setObjectName("Muted")
        head.addWidget(self.caps_label)
        lay.addLayout(head)

        # 主体三栏
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # 左：文件队列
        left = QFrame()
        left.setObjectName("Panel")
        lv = QVBoxLayout(left)
        lv.setContentsMargins(12, 12, 12, 12)
        lv.setSpacing(8)
        row = QHBoxLayout()
        lb = QLabel("文件队列")
        lb.setObjectName("Section")
        self.queue_count = QLabel("0 个文件")
        self.queue_count.setObjectName("Muted")
        row.addWidget(lb)
        row.addStretch(1)
        row.addWidget(self.queue_count)
        lv.addLayout(row)

        self.file_list = QListWidget()
        self.file_list.setObjectName("Files")
        self.file_list.setUniformItemSizes(False)
        self.file_list.currentRowChanged.connect(self._on_file_selected)
        lv.addWidget(self.file_list, 1)

        add_row = QHBoxLayout()
        add_row.setSpacing(6)
        b_add = QPushButton("＋ 添加文件")
        b_add.clicked.connect(self._pick_files)
        b_dir = QPushButton("添加文件夹")
        b_dir.clicked.connect(self._pick_folder)
        b_del = QPushButton("移除")
        b_del.clicked.connect(self._remove_current)
        add_row.addWidget(b_add, 1)
        add_row.addWidget(b_dir, 1)
        add_row.addWidget(b_del)
        lv.addLayout(add_row)
        left.setMinimumWidth(310)
        left.setMaximumWidth(400)
        splitter.addWidget(left)

        # 中：预览网格
        center = QFrame()
        center.setObjectName("Panel")
        cv = QVBoxLayout(center)
        cv.setContentsMargins(0, 12, 0, 4)
        cv.setSpacing(6)
        ch = QHBoxLayout()
        ch.setContentsMargins(14, 0, 14, 0)
        self.file_title = QLabel("未选择文件")
        self.file_title.setObjectName("Section")
        ch.addWidget(self.file_title)
        ch.addStretch(1)
        self.sel_label = QLabel("")
        self.sel_label.setObjectName("Muted")
        ch.addWidget(self.sel_label)
        b_all = QPushButton("全选")
        b_all.setObjectName("Ghost")
        b_all.clicked.connect(lambda: self._check_all(True))
        b_inv = QPushButton("反选")
        b_inv.setObjectName("Ghost")
        b_inv.clicked.connect(self._invert)
        b_none = QPushButton("清空选择")
        b_none.setObjectName("Ghost")
        b_none.clicked.connect(lambda: self._check_all(False))
        ch.addWidget(b_all)
        ch.addWidget(b_inv)
        ch.addWidget(b_none)
        cv.addLayout(ch)

        self.grid = PreviewGrid()
        self.grid.toggled.connect(self._on_unit_toggled)
        self.grid.opened.connect(self._open_preview)
        self.grid.want_thumb.connect(self._on_want_thumb)
        cv.addWidget(self.grid, 1)
        splitter.addWidget(center)

        # 右：导出设置
        right = QFrame()
        right.setObjectName("Panel")
        rv = QVBoxLayout(right)
        rv.setContentsMargins(14, 14, 14, 14)
        rv.setSpacing(9)
        rb = QLabel("导出设置")
        rb.setObjectName("Section")
        rv.addWidget(rb)

        rv.addWidget(self._tag("输出格式"))
        self.fmt_box = QComboBox()
        for label, value in (("图片 · PNG", "png"), ("图片 · JPG", "jpg"),
                             ("PDF 文件", "pdf"), ("Word 文档", "docx"), ("Excel 工作簿", "xlsx")):
            self.fmt_box.addItem(label, value)
        self.fmt_box.currentIndexChanged.connect(self._sync_option_state)
        rv.addWidget(self.fmt_box)

        rv.addWidget(self._tag("清晰度"))
        self.dpi_box = QComboBox()
        for label, value in (("标准 150 DPI", 150), ("高清 200 DPI", 200),
                             ("印刷 300 DPI", 300), ("超清 600 DPI", 600), ("草稿 96 DPI", 96)):
            self.dpi_box.addItem(label, value)
        rv.addWidget(self.dpi_box)

        self.jpg_quality = QSpinBox()
        self.jpg_quality.setRange(50, 100)
        self.jpg_quality.setSuffix(" %")
        self.jpg_row = self._labeled("JPG 质量", self.jpg_quality)
        rv.addWidget(self.jpg_row)

        self.opt_long = QCheckBox("所选页拼成一张长图")
        self.opt_merge = QCheckBox("多个工作表合并为一个工作簿")
        rv.addWidget(self.opt_long)
        rv.addWidget(self.opt_merge)

        rv.addWidget(self._tag("输出目录"))
        dir_row = QHBoxLayout()
        dir_row.setSpacing(6)
        self.out_dir = QLineEdit()
        dir_row.addWidget(self.out_dir, 1)
        b_browse = QPushButton("…")
        b_browse.setFixedWidth(36)
        b_browse.clicked.connect(self._choose_outdir)
        dir_row.addWidget(b_browse)
        rv.addLayout(dir_row)

        name_row = QHBoxLayout()
        name_row.setSpacing(6)
        self.prefix = QLineEdit()
        self.prefix.setPlaceholderText("文件名前缀")
        self.suffix = QLineEdit()
        self.suffix.setPlaceholderText("文件名后缀")
        name_row.addWidget(self.prefix)
        name_row.addWidget(self.suffix)
        rv.addLayout(name_row)

        rv.addStretch(1)

        self.summary = QLabel("尚未选择任何页面")
        self.summary.setObjectName("Muted")
        self.summary.setWordWrap(True)
        rv.addWidget(self.summary)

        self.export_btn = QPushButton("开始导出")
        self.export_btn.setObjectName("Primary")
        self.export_btn.setMinimumHeight(42)
        self.export_btn.clicked.connect(self._start_export)
        rv.addWidget(self.export_btn)

        open_out = QPushButton("打开输出目录")
        open_out.clicked.connect(self._open_out_dir)
        rv.addWidget(open_out)
        right.setMinimumWidth(290)
        right.setMaximumWidth(360)
        splitter.addWidget(right)

        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setStretchFactor(2, 0)
        splitter.setSizes([330, 660, 300])
        lay.addWidget(splitter, 1)

        # 底部：进度 + 日志
        bottom = QFrame()
        bottom.setObjectName("Panel")
        bv = QVBoxLayout(bottom)
        bv.setContentsMargins(12, 10, 12, 10)
        bv.setSpacing(6)
        prow = QHBoxLayout()
        self.status_label = QLabel("就绪：把 PDF / Word / Excel 拖进窗口即可开始")
        self.status_label.setObjectName("Muted")
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setFixedWidth(220)
        self.progress.setTextVisible(False)
        prow.addWidget(self.status_label, 1)
        prow.addWidget(self.progress)
        bv.addLayout(prow)

        self.log = QPlainTextEdit()
        self.log.setObjectName("Log")
        self.log.setReadOnly(True)
        self.log.setMaximumBlockCount(500)
        self.log.setFixedHeight(84)
        bv.addWidget(self.log)
        lay.addWidget(bottom)

    @staticmethod
    def _tag(text: str) -> QLabel:
        lb = QLabel(text)
        lb.setObjectName("Muted")
        lb.setStyleSheet("font-size:12px;")
        return lb

    def _labeled(self, text: str, widget: QWidget) -> QWidget:
        box = QWidget()
        lay = QHBoxLayout(box)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(8)
        lb = self._tag(text)
        lb.setFixedWidth(64)
        lay.addWidget(lb)
        lay.addWidget(widget, 1)
        return box

    def _build_shortcuts(self) -> None:
        a_add = QAction("添加文件…", self)
        a_add.setShortcut("Ctrl+O")
        a_add.triggered.connect(self._pick_files)
        a_dir = QAction("添加文件夹…", self)
        a_dir.setShortcut("Ctrl+Shift+O")
        a_dir.triggered.connect(self._pick_folder)
        a_clear = QAction("清空队列", self)
        a_clear.triggered.connect(self._clear_all)
        for a in (a_add, a_dir, a_clear):
            self.addAction(a)

    # ------------------------------------------------------------------ 设置读写
    def _restore_settings(self) -> None:
        s = self.settings
        default_out = str(Path.home() / "Desktop" / "页面提取")
        self.out_dir.setText(str(s.value("out_dir", default_out)))
        self.prefix.setText(str(s.value("prefix", "")))
        self.suffix.setText(str(s.value("suffix", "")))
        fmt = str(s.value("fmt", "png"))
        idx = self.fmt_box.findData(fmt)
        self.fmt_box.setCurrentIndex(idx if idx >= 0 else 0)
        dpi = int(s.value("dpi", 200))
        j = self.dpi_box.findData(dpi)
        self.dpi_box.setCurrentIndex(j if j >= 0 else 1)
        self.jpg_quality.setValue(int(s.value("jpg_quality", 92)))
        self.opt_long.setChecked(s.value("long_image", "false") in (True, "true"))
        self.opt_merge.setChecked(s.value("merge_xlsx", "true") in (True, "true"))
        self._sync_option_state()

    def _save_settings(self) -> None:
        s = self.settings
        s.setValue("out_dir", self.out_dir.text())
        s.setValue("fmt", self.fmt_box.currentData())
        s.setValue("dpi", self.dpi_box.currentData())
        s.setValue("jpg_quality", self.jpg_quality.value())
        s.setValue("long_image", "true" if self.opt_long.isChecked() else "false")
        s.setValue("merge_xlsx", "true" if self.opt_merge.isChecked() else "false")
        s.setValue("prefix", self.prefix.text())
        s.setValue("suffix", self.suffix.text())

    def _sync_option_state(self) -> None:
        fmt = self.fmt_box.currentData()
        is_img = fmt in ("png", "jpg")
        self.dpi_box.setEnabled(fmt in ("png", "jpg", "docx"))
        self.jpg_row.setVisible(fmt == "jpg")
        self.opt_long.setVisible(is_img)
        self.opt_merge.setVisible(fmt == "xlsx")

    # ------------------------------------------------------------------ 添加文件
    def _pick_files(self) -> None:
        pattern = " ".join(f"*{ext}" for ext in SUPPORTED_EXTS)
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择 PDF / Word / Excel 文件", str(Path.home()), f"可提取文件 ({pattern});;所有文件 (*.*)"
        )
        if files:
            self.add_paths(files)

    def _pick_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹", str(Path.home()))
        if not folder:
            return
        found = sorted(
            str(p) for p in Path(folder).rglob("*") if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS
        )
        if not found:
            QMessageBox.information(self, "提示", "这个文件夹里没有 PDF / Word / Excel 文件。")
            return
        self.add_paths(found)

    def add_paths(self, paths: list[str]) -> None:
        added = 0
        for raw in paths:
            p = Path(raw)
            if p.suffix.lower() not in SUPPORTED_EXTS:
                continue
            key = str(p.resolve()).lower()
            if any(e.path == Path(key) for e in self.entries.values()):
                continue
            uid = uuid.uuid4().hex[:8]
            entry = DocEntry(path=p.resolve(), uid=uid)
            self.entries[uid] = entry
            self.order.append(uid)
            self._add_queue_item(entry)
            self.loader.enqueue(entry)
            added += 1
        if added:
            self._log(f"已加入 {added} 个文件")
            if self.current_uid is None:
                self.file_list.setCurrentRow(0)
        self._refresh_counts()

    def _add_queue_item(self, entry: DocEntry) -> None:
        item = QListWidgetItem(self._queue_text(entry))
        item.setData(Qt.ItemDataRole.UserRole, entry.uid)
        item.setSizeHint(QSize(0, 58))
        item.setToolTip(str(entry.path))
        self.file_list.addItem(item)

    def _queue_text(self, entry: DocEntry) -> str:
        if entry.loading:
            meta = f"{entry.kind_label} · 正在解析…"
        elif entry.error:
            meta = f"{entry.kind_label} · 打开失败"
        else:
            unit = "页" if entry.kind in ("pdf", "word") else "张表"
            meta = f"{entry.kind_label} · 共 {len(entry.units)} {unit} · 已选 {len(entry.checked)}"
        return f"{entry.name}\n{meta}"

    def _refresh_queue_item(self, entry: DocEntry) -> None:
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            if item.data(Qt.ItemDataRole.UserRole) == entry.uid:
                item.setText(self._queue_text(entry))
                if entry.error:
                    item.setToolTip(f"{entry.path}\n{entry.error}")
                return

    def _refresh_counts(self) -> None:
        total_files = len(self.order)
        total_units = sum(len(e.checked) for e in self.entries.values())
        self.queue_count.setText(f"{total_files} 个文件")
        self.summary.setText(
            f"已选 {total_files} 个文件中的 {total_units} 个页面/工作表" if total_units else "尚未选择任何页面"
        )
        self.export_btn.setEnabled(total_units > 0 and self._export_worker is None)
        if self.current_uid:
            entry = self.entries.get(self.current_uid)
            if entry and not entry.loading:
                total = len(entry.units)
                self.sel_label.setText(f"已选 {len(entry.checked)} / {total}")

    # ------------------------------------------------------------------ 载入回调
    def _on_busy(self, entry: DocEntry, msg: str) -> None:
        self.status_label.setText(f"{entry.name}：{msg}")

    def _on_loaded(self, entry: DocEntry) -> None:
        self._refresh_queue_item(entry)
        self._log(f"已解析 {entry.name}（{len(entry.units)} 个页面/工作表）")
        if self.current_uid == entry.uid:
            self._show_entry(entry)
        self._refresh_counts()
        if entry.error is None and len(entry.units) > 40:
            self._log(f"{entry.name} 页面较多，默认不勾选，请用「全选」或逐页选择")

    def _on_load_failed(self, entry: DocEntry, msg: str) -> None:
        self._refresh_queue_item(entry)
        self._log(f"❌ {entry.name}：{msg}")
        self.status_label.setText(f"{entry.name} 打开失败：{msg}")
        if self.current_uid == entry.uid:
            self._show_entry(entry)

    def _on_caps(self, caps: dict) -> None:
        word = "Word ✓" if caps.get("word") else "Word ✗"
        excel = "Excel ✓" if caps.get("excel") else "Excel ✗"
        extra = " · LibreOffice ✓" if caps.get("soffice") else ""
        self.caps_label.setText(f"转换引擎：{word} · {excel}{extra}")
        if not caps.get("word"):
            self._log("未检测到 Microsoft Word：Word 文件无法确定页码，请安装 Office 或 LibreOffice")

    # ------------------------------------------------------------------ 选择与预览
    def _on_file_selected(self, row: int) -> None:
        if row < 0 or row >= self.file_list.count():
            return
        uid = self.file_list.item(row).data(Qt.ItemDataRole.UserRole)
        self.current_uid = uid
        entry = self.entries.get(uid)
        if entry:
            self._show_entry(entry)

    def _show_entry(self, entry: DocEntry) -> None:
        self.file_title.setText(entry.name)
        if entry.loading:
            self.grid.show_message("正在解析文件…")
            self.sel_label.setText("")
            return
        if entry.error:
            self.grid.show_message(f"打开失败：{entry.error}")
            self.sel_label.setText("")
            return
        self.grid.set_units(entry.uid, entry.units, entry.checked)
        self.sel_label.setText(f"已选 {len(entry.checked)} / {len(entry.units)}")

    def _on_unit_toggled(self, key: str, value: bool) -> None:
        entry = self.entries.get(self.current_uid or "")
        if not entry:
            return
        if value:
            entry.checked.add(key)
        else:
            entry.checked.discard(key)
        self._refresh_queue_item(entry)
        self._refresh_counts()

    def _check_all(self, value: bool) -> None:
        entry = self.entries.get(self.current_uid or "")
        if not entry or not entry.units:
            return
        entry.checked = {u.key for u in entry.units} if value else set()
        self.grid.sync_checks(entry.checked)
        self._refresh_queue_item(entry)
        self._refresh_counts()

    def _invert(self) -> None:
        entry = self.entries.get(self.current_uid or "")
        if not entry or not entry.units:
            return
        all_keys = {u.key for u in entry.units}
        entry.checked = all_keys - entry.checked
        self.grid.sync_checks(entry.checked)
        self._refresh_queue_item(entry)
        self._refresh_counts()

    def _on_want_thumb(self, uid: str, key: str) -> None:
        entry = self.entries.get(uid)
        if not entry or not entry.source:
            return
        cached = entry.thumbs.get(key)
        if cached is not None:
            self.grid.apply_thumb(uid, key, cached)
            return
        self.thumbs.request(uid, entry.source, key, 300)

    def _on_thumb_ready(self, uid: str, key: str, pix) -> None:
        entry = self.entries.get(uid)
        if entry:
            entry.thumbs[key] = pix
        self.grid.apply_thumb(uid, key, pix)

    def _on_thumb_failed(self, uid: str, key: str, msg: str) -> None:
        self.grid.apply_thumb_error(uid, key, msg)

    def _open_preview(self, key: str) -> None:
        entry = self.entries.get(self.current_uid or "")
        if not entry or not entry.source:
            return
        unit = next((u for u in entry.units if u.key == key), None)
        title = f"{entry.name} · {unit.label if unit else key}"
        PreviewDialog(entry.source, key, title, self).exec()

    # ------------------------------------------------------------------ 导出
    def _choose_outdir(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "选择输出目录", self.out_dir.text() or str(Path.home()))
        if folder:
            self.out_dir.setText(folder)

    def _build_spec(self) -> ExportSpec:
        return ExportSpec(
            fmt=self.fmt_box.currentData(),
            out_dir=Path(self.out_dir.text().strip() or (Path.home() / "页面提取")),
            dpi=int(self.dpi_box.currentData()),
            jpg_quality=self.jpg_quality.value(),
            long_image=self.opt_long.isChecked(),
            merge_xlsx=self.opt_merge.isChecked(),
            prefix=self.prefix.text().strip(),
            suffix=self.suffix.text().strip(),
        )

    def _start_export(self) -> None:
        spec = self._build_spec()
        jobs: list[Job] = []
        for uid in self.order:
            entry = self.entries[uid]
            if not entry.source or not entry.checked:
                continue
            units = [u for u in entry.units if u.key in entry.checked]
            if units:
                jobs.append(Job(source=entry.source, units=units))
        if not jobs:
            QMessageBox.warning(self, "无法导出", "还没有勾选任何页面或工作表。")
            return
        if spec.fmt == "xlsx" and not any(getattr(j.source, "kind", "") == "excel" for j in jobs):
            QMessageBox.warning(self, "格式不匹配", "Excel 工作簿格式只能导出工作表，请改选 PDF / 图片 / Word。")
            return
        if not spec.out_dir.exists():
            try:
                spec.out_dir.mkdir(parents=True, exist_ok=True)
            except Exception as exc:
                QMessageBox.critical(self, "目录不可用", f"无法创建输出目录：{exc}")
                return

        self._save_settings()
        self.progress.setValue(0)
        self.export_btn.setEnabled(False)
        self.status_label.setText("正在导出…")
        self._log(f"开始导出：{spec.fmt.upper()} → {spec.out_dir}")

        self._export_worker = ExportWorker(jobs, spec, self)
        self._export_worker.progress.connect(self._on_export_progress)
        self._export_worker.done.connect(self._on_export_done)
        self._export_worker.failed.connect(self._on_export_failed)
        self._export_worker.start()

    def _on_export_progress(self, done: int, total: int, msg: str) -> None:
        self.progress.setValue(int(done / max(total, 1) * 100))
        self.status_label.setText(msg)

    def _on_export_done(self, results: list) -> None:
        self._export_worker = None
        self.progress.setValue(100)
        self.export_btn.setEnabled(True)
        self.status_label.setText(f"完成：生成 {len(results)} 个文件")
        for p in results[:60]:
            self._log(f"✔ {p}")
        if len(results) > 60:
            self._log(f"… 其余 {len(results) - 60} 个文件略")
        self._refresh_counts()

        box = QMessageBox(self)
        box.setWindowTitle("导出完成")
        box.setText(f"已生成 {len(results)} 个文件。\n输出目录：{self.out_dir.text()}")
        open_btn = box.addButton("打开文件夹", QMessageBox.ButtonRole.AcceptRole)
        box.addButton("好", QMessageBox.ButtonRole.RejectRole)
        box.exec()
        if box.clickedButton() is open_btn:
            self._open_out_dir()

    def _on_export_failed(self, msg: str) -> None:
        self._export_worker = None
        self.export_btn.setEnabled(True)
        self.progress.setValue(0)
        self.status_label.setText("导出失败")
        self._log(f"❌ 导出失败：{msg}")
        QMessageBox.critical(self, "导出失败", msg)

    def _open_out_dir(self) -> None:
        folder = Path(self.out_dir.text().strip() or ".")
        folder.mkdir(parents=True, exist_ok=True)
        try:
            import os
            os.startfile(str(folder))  # noqa: S606
        except Exception as exc:
            QMessageBox.warning(self, "打不开目录", str(exc))

    # ------------------------------------------------------------------ 其它
    def _remove_current(self) -> None:
        row = self.file_list.currentRow()
        if row < 0:
            return
        item = self.file_list.takeItem(row)
        uid = item.data(Qt.ItemDataRole.UserRole)
        entry = self.entries.pop(uid, None)
        if entry:
            if entry.source:
                try:
                    entry.source.close()
                except Exception:
                    pass
            self.order = [u for u in self.order if u != uid]
        self.current_uid = None
        if self.file_list.count():
            self.file_list.setCurrentRow(min(row, self.file_list.count() - 1))
        else:
            self.grid.clear()
            self.file_title.setText("未选择文件")
            self.sel_label.setText("")
        self._refresh_counts()

    def _clear_all(self) -> None:
        self.file_list.clear()
        for entry in self.entries.values():
            if entry.source:
                try:
                    entry.source.close()
                except Exception:
                    pass
        self.entries.clear()
        self.order.clear()
        self.current_uid = None
        self.grid.clear()
        self.file_title.setText("未选择文件")
        self.sel_label.setText("")
        self._refresh_counts()

    def _log(self, text: str) -> None:
        self.log.appendPlainText(text)

    # ------------------------------------------------------------------ 拖放
    def dragEnterEvent(self, event):  # noqa: N802
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):  # noqa: N802
        paths = [u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()]
        files: list[str] = []
        for p in paths:
            path = Path(p)
            if path.is_dir():
                files.extend(str(f) for f in path.rglob("*") if f.is_file() and f.suffix.lower() in SUPPORTED_EXTS)
            elif path.suffix.lower() in SUPPORTED_EXTS:
                files.append(str(path))
        if files:
            self.add_paths(files)
        else:
            self._log("拖入的文件类型不支持")

    # ------------------------------------------------------------------ 退出清理
    def shutdown(self) -> None:
        """关闭前后台线程与已打开的文档（由主窗口 closeEvent 调用）。"""
        self._save_settings()
        for worker in (self.loader, self.thumbs):
            try:
                worker.stop()
                worker.wait(1500)
            except Exception:
                pass
        if self._export_worker is not None:
            self._export_worker.wait(3000)
        for entry in self.entries.values():
            if entry.source:
                try:
                    entry.source.close()
                except Exception:
                    pass
