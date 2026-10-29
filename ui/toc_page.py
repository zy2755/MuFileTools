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
from PySide6.QtGui import QTextOption          # Qt6 把它挪到 QtGui 了，QtCore 里没有
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
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

from core import office, outline, toc
from core.exporter import _unique
from ui.mark_dialog import MarkDialog

NO_TITLE_HINT = (
    "没有识别到标题样式，Word 只会插一个空目录。\n"
    "点右上角的「手动标标题…」自己挑出章节（字体字号一个都不会变），"
    "或者先在 Word 里把章节标题套上「标题 1 / 2 / 3」样式。"
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

    def __init__(self, src: str, out: str, opts: dict, rows: list | None = None, parent=None):
        super().__init__(parent)
        self._src = src
        self._out = out
        self._opts = dict(opts)
        self._rows = list(rows or [])

    def _report(self, done: int, total: int, msg: str) -> None:
        self.progress.emit(done, total, msg)

    def run(self) -> None:  # noqa: D102
        try:
            if office.has_word():
                res = office.add_toc_to_docx(
                    self._src, self._out, progress=self._report, **self._opts
                )
                msg = f"{res['entries']} 条目录"
                if res.get("mode") == "updated":
                    levels = int(self._opts.get("levels", 3))
                    msg = (f"文档本来就有目录 —— 已在原位置就地更新为 1-{levels} 级"
                           f"（{res['entries']} 条），位置和排版都没动")
                elif res.get("mode") == "kept":
                    msg = f"文档已有目录，按你的设置保持原样没动（{res['entries']} 条）"
                elif res.get("replaced"):
                    msg += f"，替换了 {res['replaced']} 个旧目录"
                self.done.emit(True, msg, str(res["out"]))
            else:
                plain = {k: v for k, v in self._opts.items() if k != "replace_existing"}
                plain["rows"] = self._rows or None   # 手动标记的结果要填进域结果里
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
        self._has_toc = False          # 载入的文档里本来就有目录？
        self._marks: dict[int, int] = {}   # 手动标的 {段落下标: 级别}
        self._edited = False           # 这份文档的标题是不是手动标的
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
        self.opt_replace = QCheckBox("文档已有目录时就地更新（不挪位置）")
        self.opt_title.setChecked(True)
        self.opt_break.setChecked(True)
        self.opt_replace.setChecked(True)
        for box in (self.opt_title, self.opt_break, self.opt_replace):
            rv.addWidget(box)

        stat_row = QHBoxLayout()
        stat_row.setSpacing(6)
        stat_sec = QLabel("识别到的标题")
        stat_sec.setObjectName("Section")
        stat_row.addWidget(stat_sec)
        stat_row.addStretch(1)
        self.btn_mark = QPushButton("手动标标题…")
        # 文字本身以省略号结尾（表示「点开会继续」），别当成被截断的提示去加宽度。
        # 实测 sizeHint 只要 106px，给它一个下限防止窄面板挤压即可。
        self.btn_mark.setMinimumWidth(106)
        self.btn_mark.setToolTip(
            "文档里的标题是手打的、没套标题样式？自己指认哪些行是几级标题。\n"
            "只记大纲级别，字体字号颜色一个都不会变。"
        )
        self.btn_mark.setEnabled(False)
        self.btn_mark.clicked.connect(self.open_mark_dialog)
        stat_row.addWidget(self.btn_mark)
        rv.addLayout(stat_row)
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
        # 右侧面板窄，长标题（「2.1.1 校园闲置经济市场规模」这种）一换行就散成两行，
        # 缩进层级也跟着乱。改成不折行 + 横向滚动，一行一条看得清。
        self.preview.setWordWrapMode(QTextOption.WrapMode.NoWrap)
        self.preview.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
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
        self._marks = {}
        self._edited = False
        suffix = p.suffix.lower()
        if suffix not in toc.DOCX_EXTS:
            self._rows = []
            self._has_toc = False
            self._refresh_view()
            self.warn(f"只支持 .docx（这是 {suffix or '无扩展名'}）。请在 Word 里「另存为」一份 .docx 再来。")
            return

        try:
            self._rows = toc.scan_headings(p)
            self._has_toc = toc.has_toc(p)
        except Exception as exc:  # noqa: BLE001
            self._rows = []
            self._has_toc = False
            self._refresh_view()
            self.warn(f"读不了这份文档：{exc}")
            return

        self._refresh_view()
        if not self._rows:
            self.warn(NO_TITLE_HINT)
            # 状态行也要刷 —— 否则换一份没标题的文档时，底下还挂着上一份的「已载入…」
            self.notify(f"已载入 {p.name}：没有识别到标题样式 —— 点「手动标标题…」自己挑出章节")
        elif self._worker is None:
            depth = int(self.depth_box.currentData())
            count = sum(toc.level_counts(self._rows, depth).values())
            if self._has_toc:
                self.notify(f"已载入 {p.name}：文档本来就有目录，"
                            f"会在原位置就地更新成 {count} 条，位置不动")
            else:
                self.notify(f"已载入 {p.name}：文档没有目录，将新插入 {count} 条")

    def clear(self) -> None:
        self.src_path = None
        self._rows = []
        self._has_toc = False
        self._marks = {}
        self._edited = False
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

    # ------------------------------------------------------------------ 手动标标题
    def open_mark_dialog(self) -> None:
        """打开段落清单，让人自己指认哪一行是几级标题。

        第一次打开会自动猜一遍（core.outline.guess_levels），之后每次打开
        沿用上次的结果 —— 改完深度想再生成一次时，不用重新标。
        """
        if not self.src_path:
            return
        p = Path(self.src_path)
        if p.suffix.lower() not in toc.DOCX_EXTS:
            self.warn("手动标标题只支持 .docx。")
            return
        if self._worker is not None:
            return

        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            paras = outline.list_paragraphs(p)
        except Exception as exc:  # noqa: BLE001
            QApplication.restoreOverrideCursor()
            self.warn(f"读不了这份文档：{exc}")
            return
        finally:
            QApplication.restoreOverrideCursor()

        if not paras:
            self.warn("这份文档里没有可标记的文字段落（可能是扫描件或整页图片）。")
            return

        marks = dict(self._marks)
        if not marks:
            marks = outline.guess_levels(paras)      # 头一次打开才自动猜

        dlg = MarkDialog(paras, marks, p.name, self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return                                    # 取消就保持原样，不动之前的标记

        self._marks = dlg.marks()
        self._rows = dlg.marked_rows()
        self._edited = True
        self._refresh_view()

        count = len(self._rows)
        if count:
            self.notify(f"已手动标记 {count} 条标题 —— 点「生成带目录的 Word」")
        else:
            self.warn("一条标题都没标记。点「手动标标题…」挑出章节，或干脆套用 Word 的标题样式。")

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
            self.btn_mark.setEnabled(False)
            self._repolish(self.file_label)
            return

        p = Path(self.src_path)
        depth = int(self.depth_box.currentData())

        if self._rows:
            self.file_label.setObjectName("FileName")
            note = "（手动标记）" if self._edited else ""
            self.file_label.setText(f"{p.name}\n\n{len(self._rows)} 条标题{note}")
        else:
            self.file_label.setObjectName("Placeholder")
            self.file_label.setText(f"{p.name}\n\n没有识别到标题样式")
        self._repolish(self.file_label)
        self.btn_mark.setEnabled(True)

        try:
            size_kb = p.stat().st_size / 1024
            head = f"{size_kb:.1f} KB"
        except Exception:
            head = ""
        # 已有目录 / 没有目录，这两种情况的处理完全不同，得让人一眼看见
        if self._has_toc:
            mode_note = "文档已有目录 → 就地更新，位置和排版不动"
        else:
            mode_note = "文档没有目录 → 新插入一份"
        self.file_meta.setText(
            f"{head} · {mode_note} · 原文件只读不动，结果另存为「{p.stem}_带目录.docx」"
            if head else
            f"{mode_note} · 原文件只读不动，结果另存为「{p.stem}_带目录.docx」"
        )

        counts = toc.level_counts(self._rows, depth)
        names = {1: "一级", 2: "二级", 3: "三级", 4: "四级"}
        if self._rows:
            self.stat_label.setText(
                " · ".join(f"{names[i]} {counts[i]} 个" for i in counts)
                + ("　← 手动标记" if self._edited else "")
            )
        else:
            self.stat_label.setText("0 条 —— 点右边的「手动标标题…」自己挑出章节")

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
            "marks": dict(self._marks),
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
        plain["rows"] = self._rows or None
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

        self._worker = _TocWorker(self.src_path, str(target), self._opts(), self._rows, self)
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
