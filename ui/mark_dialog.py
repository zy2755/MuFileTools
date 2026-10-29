# -*- coding: utf-8 -*-
"""「手动标标题级别」对话框。

场景：文档里的标题全是手打的（WPS 用户的重灾区），一个「标题 1/2/3」样式都没套，
Word 的目录一条都认不出来。这个界面让人自己指认：这行是一级，那行是二级。

标完不改任何外观 —— 只往段落属性里写一个大纲级别（见 core/outline.py）。

操作上做了三层提速，因为一份 120 段的文档手点下来很累：
  1. 打开时已经按编号规律 + 字形特征自动猜好一遍，人只改错的那几条；
  2. 多选几行后按顶部的「1 级 / 2 级 / 3 级 / 4 级」，一次改一批；
  3. 双击某行可以在 正文 → 1 → 2 → 3 → 4 → 正文 之间循环，不用展开下拉。
"""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

from core.outline import MAX_LEVEL, Para

_LEVEL_LABELS = ["正文", "1 级", "2 级", "3 级", "4 级"]


class MarkDialog(QDialog):
    """段落清单 + 每行一个级别下拉。"""

    def __init__(self, paras: list[Para], marks: dict[int, int], title: str = "", parent=None):
        super().__init__(parent)
        self.setWindowTitle("标记标题级别")
        self.resize(1020, 720)
        # 背景交给全局 QSS 的 QDialog 规则，别在这里写死颜色

        self._paras = list(paras)
        self._marks: dict[int, int] = {p.index: int(marks.get(p.index, p.level)) for p in self._paras}
        self._guessed_total = sum(1 for v in marks.values() if v > 0)
        self._row_of = {p.index: r for r, p in enumerate(self._paras)}
        self._doc_title = title

        self._build_ui()
        self._fill_table()
        self._update_stat()

    # ------------------------------------------------------------------ 界面
    def _build_ui(self) -> None:
        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 12, 14, 12)
        lay.setSpacing(10)

        head = QHBoxLayout()
        title_box = QVBoxLayout()
        title_box.setSpacing(2)
        t = QLabel("标记哪些行是章节标题")
        t.setObjectName("Title")
        title_box.addWidget(t)
        sub = QLabel(
            "只记录「这行属于第几级」，文档的字体、字号、颜色、缩进一个都不会变 —— "
            "目录认得出它，但排版还是你原来那样。"
        )
        sub.setObjectName("Subtitle")
        sub.setWordWrap(True)
        title_box.addWidget(sub)
        head.addLayout(title_box, 1)
        lay.addLayout(head)

        # ---- 工具条
        bar = QHBoxLayout()
        bar.setSpacing(8)
        self.search = QLineEdit()
        self.search.setPlaceholderText("按文字筛选…")
        self.search.textChanged.connect(self._apply_filter)
        bar.addWidget(self.search, 1)

        self.only_cand = QCheckBox("只看候选")
        self.only_cand.setToolTip("隐藏明显是正文的长段落，只留下像标题的行")
        self.only_cand.toggled.connect(self._apply_filter)
        bar.addWidget(self.only_cand)

        bar.addSpacing(6)
        bar.addWidget(QLabel("选中行设为"))
        self._bulk_buttons: list[QPushButton] = []
        for level in range(1, MAX_LEVEL + 1):
            b = QPushButton(f"{level} 级")
            b.setFixedWidth(52)
            b.clicked.connect(lambda _c=False, lv=level: self._apply_bulk(lv))
            bar.addWidget(b)
            self._bulk_buttons.append(b)
        b_body = QPushButton("正文")
        b_body.setFixedWidth(52)
        b_body.clicked.connect(lambda: self._apply_bulk(0))
        bar.addWidget(b_body)
        lay.addLayout(bar)

        # ---- 表格
        self.table = QTableWidget(len(self._paras), 2, self)
        self.table.setHorizontalHeaderLabels(["文档内容", "级别"])
        self.table.verticalHeader().setVisible(False)
        self.table.setShowGrid(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setDefaultSectionSize(28)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        self.table.setColumnWidth(1, 96)
        self.table.cellDoubleClicked.connect(self._on_double_click)
        lay.addWidget(self.table, 1)

        # ---- 底部
        foot = QHBoxLayout()
        self.stat = QLabel("")
        self.stat.setObjectName("Muted")
        foot.addWidget(self.stat, 1)

        self.btn_clear = QPushButton("清除自动预判")
        self.btn_clear.setObjectName("Ghost")
        self.btn_clear.setToolTip("把所有行恢复成「正文」，从零开始自己标")
        self.btn_clear.clicked.connect(self._clear_guess)
        foot.addWidget(self.btn_clear)

        cancel = QPushButton("取消")
        cancel.clicked.connect(self.reject)
        foot.addWidget(cancel)

        ok = QPushButton("确定")
        ok.setObjectName("Primary")
        ok.setMinimumWidth(96)
        ok.clicked.connect(self.accept)
        foot.addWidget(ok)
        lay.addLayout(foot)

    def _fill_table(self) -> None:
        mono = QFont()
        mono.setPointSize(9)
        for row, para in enumerate(self._paras):
            item = QTableWidgetItem(para.text)
            item.setToolTip(para.text)
            if para.index < 0:
                item.setText("（表格内）" + para.text)
            if para.level > 0 and not para.candidate:
                item.setForeground(QColor("#8a8a8a"))
            self.table.setItem(row, 0, item)

            combo = QComboBox()
            for label in _LEVEL_LABELS:
                combo.addItem(label)
            combo.setCurrentIndex(max(0, min(MAX_LEVEL, self._marks.get(para.index, para.level))))
            combo.currentIndexChanged.connect(
                lambda _i, idx=para.index: self._on_level_changed(idx)
            )
            self.table.setCellWidget(row, 1, combo)

    # ------------------------------------------------------------------ 交互
    def _combo(self, index: int) -> QComboBox | None:
        row = self._row_of.get(index)
        if row is None:
            return None
        return self.table.cellWidget(row, 1)

    def _on_level_changed(self, index: int) -> None:
        combo = self._combo(index)
        if combo is None:
            return
        self._marks[index] = combo.currentIndex()
        self._update_stat()

    def _set_level(self, index: int, level: int) -> None:
        combo = self._combo(index)
        if combo is None:
            return
        combo.blockSignals(True)          # 别让 setCurrentIndex 反过来再触发一遍
        combo.setCurrentIndex(max(0, min(MAX_LEVEL, level)))
        combo.blockSignals(False)
        self._marks[index] = combo.currentIndex()

    def _apply_bulk(self, level: int) -> None:
        rows = sorted({i.row() for i in self.table.selectionModel().selectedRows()})
        if not rows:
            self.stat.setText("先在左边点一下要改的行（可以按住 Shift / Ctrl 多选）")
            return
        for row in rows:
            index = self._paras[row].index
            self._set_level(index, level)
        self._update_stat()
        self.stat.setText(self.stat.text() + f"　　刚把 {len(rows)} 行设成了「{_LEVEL_LABELS[level]}」")

    def _on_double_click(self, row: int, _col: int) -> None:
        """双击在 正文 → 1 → 2 → 3 → 4 → 正文 之间循环。"""
        index = self._paras[row].index
        cur = self._marks.get(index, 0)
        self._set_level(index, (cur + 1) % (MAX_LEVEL + 1))
        self._update_stat()

    def _apply_filter(self) -> None:
        keyword = self.search.text().strip().lower()
        only_cand = self.only_cand.isChecked()
        hidden = 0
        for row, para in enumerate(self._paras):
            show = True
            if keyword and keyword not in para.text.lower():
                show = False
            if show and only_cand:
                show = para.candidate or self._marks.get(para.index, 0) > 0
            if not show:
                hidden += 1
            self.table.setRowHidden(row, not show)

        if only_cand and hidden:
            self.stat.setText(f"已隐藏 {hidden} 行（看着像正文）；勾掉「只看候选」可以看全部")

    def _clear_guess(self) -> None:
        for para in self._paras:
            self._set_level(para.index, 0)
        self._update_stat()
        self.stat.setText("已清空全部标记，所有行都是「正文」—— 现在可以自己从头标")

    def _update_stat(self) -> None:
        counts = {lv: 0 for lv in range(1, MAX_LEVEL + 1)}
        for level in self._marks.values():
            if 1 <= level <= MAX_LEVEL:
                counts[level] += 1
        total = sum(counts.values())
        guess_note = f"（其中 {self._guessed_total} 条是自动猜的）" if self._guessed_total else ""
        self.stat.setText(
            f"已标 {total} 条：一级 {counts[1]} · 二级 {counts[2]} · "
            f"三级 {counts[3]} · 四级 {counts[4]}{guess_note}"
        )

    # ------------------------------------------------------------------ 结果
    def marks(self) -> dict[int, int]:
        """所有行的最终级别，只返回非正文的（正文不用写，省得覆盖样式级别）。"""
        return {idx: lv for idx, lv in self._marks.items() if lv > 0}

    def marked_rows(self) -> list[tuple[int, str]]:
        """按文档顺序的 [(级别, 文本)]，给界面预览用。"""
        out: list[tuple[int, str]] = []
        for para in self._paras:
            level = self._marks.get(para.index, 0)
            if level > 0:
                out.append((level, para.text))
        return out
