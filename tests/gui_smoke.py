"""离屏 GUI 冒烟测试：构建界面 → 入队 → 等解析 → 勾选 → 导出 → 校验。

运行： python tests/gui_smoke.py
不弹窗、不需要桌面会话（QT_QPA_PLATFORM=offscreen）。
"""
from __future__ import annotations

import os
import shutil
import sys
import time
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from PySide6.QtWidgets import QApplication  # noqa: E402

import ui.extract_page as ep  # noqa: E402
import ui.ico_page as ip  # noqa: E402
import ui.shell as mw  # noqa: E402
from ui.shell import MainWindow  # noqa: E402
from ui.theme import QSS  # noqa: E402

OUT = ROOT / "out_gui"
SAMPLES = ROOT / "samples"
FAILS: list[str] = []


class _FakeBox:
    """替换 QMessageBox，避免离屏测试被模态框卡住。"""

    ButtonRole = type("R", (), {"AcceptRole": 0, "RejectRole": 1})

    def __init__(self, *a, **k):
        pass

    def setWindowTitle(self, *a):
        pass

    def setText(self, *a):
        pass

    def addButton(self, *a, **k):
        return object()

    def exec(self):
        return 0

    def clickedButton(self):
        return None

    @staticmethod
    def warning(*a, **k):
        return 0

    @staticmethod
    def critical(*a, **k):
        return 0

    @staticmethod
    def information(*a, **k):
        return 0

    @staticmethod
    def about(*a, **k):
        return 0


def expect(cond: bool, label: str, detail: str = "") -> None:
    print(("  ✔ " if cond else "  ✘ ") + label + (f"  {detail}" if detail else ""))
    if not cond:
        FAILS.append(label)


def pump(app: QApplication, seconds: float) -> None:
    end = time.time() + seconds
    while time.time() < end:
        app.processEvents()
        time.sleep(0.03)


def wait_until(app: QApplication, cond, timeout: float = 180.0, label: str = "") -> bool:
    end = time.time() + timeout
    while time.time() < end:
        app.processEvents()
        if cond():
            return True
        time.sleep(0.05)
    print(f"  … 等待超时：{label}")
    return False


def main() -> int:
    if not (SAMPLES / "样例报告.pdf").exists():
        print("缺少样例，先运行 python tools/make_samples.py")
        return 2
    shutil.rmtree(OUT, ignore_errors=True)

    # 这个测试会改「输出目录 / 格式 / 清晰度」并在导出时写回注册表。
    # 先快照真实配置，结束时原样还原（踩过一次：测试把 out_dir 写成了 out_gui）。
    from PySide6.QtCore import QSettings

    real = QSettings("MuFileTools", "PageExtract")
    snapshot = {k: real.value(k) for k in real.allKeys()}
    try:
        return _run()
    finally:
        for k in real.allKeys():
            if k not in snapshot:
                real.remove(k)
        for k, v in snapshot.items():
            real.setValue(k, v)
        real.sync()


def _run() -> int:

    # 三个模块都换掉 QMessageBox，模态框全部变空操作
    ep.QMessageBox = _FakeBox  # type: ignore[assignment]
    ip.QMessageBox = _FakeBox  # type: ignore[assignment]
    mw.QMessageBox = _FakeBox  # type: ignore[assignment]

    app = QApplication(sys.argv)
    app.setStyleSheet(QSS)

    print("[1] 构建主窗口")
    win = MainWindow()
    win.show()
    pump(app, 0.5)
    page = win.extract_page
    expect(win.isVisible(), "窗口已创建")
    expect(win.windowTitle() == "小睦文件转换器", "窗口标题为「小睦文件转换器」", win.windowTitle())
    expect(win.stack.count() == 2, "壳窗口包含 2 个工具页", f"count={win.stack.count()}")
    expect(len(page.order) == 0, "初始队列为空")

    print("[2] 加入样例文件")
    page.add_paths([
        str(SAMPLES / "样例报告.pdf"),
        str(SAMPLES / "样例数据.xlsx"),
    ])
    expect(len(page.order) == 2, "队列有 2 个文件", f"order={len(page.order)}")

    ok = wait_until(app, lambda: all(not e.loading for e in page.entries.values()), 240, "解析完成")
    expect(ok, "两个文件解析完成")

    pdf_entry = next((e for e in page.entries.values() if e.kind == "pdf"), None)
    xls_entry = next((e for e in page.entries.values() if e.kind == "excel"), None)
    expect(pdf_entry is not None and len(pdf_entry.units) == 6, "PDF 解析出 6 页",
           f"{len(pdf_entry.units) if pdf_entry else 0}")
    expect(xls_entry is not None and len(xls_entry.units) == 3, "Excel 解析出 3 张工作表",
           f"{len(xls_entry.units) if xls_entry else 0}")
    expect(pdf_entry is not None and pdf_entry.checked == {u.key for u in pdf_entry.units},
           "≤40 单位默认全选")

    print("[3] 预览网格与缩略图")
    page.file_list.setCurrentRow(page.order.index(pdf_entry.uid))
    pump(app, 0.3)
    expect(len(page.grid._cards) == 6, "网格渲染 6 张卡片", f"{len(page.grid._cards)}")
    got = wait_until(app, lambda: sum(1 for c in page.grid._cards
                                      if c.thumb.pixmap() is not None and not c.thumb.pixmap().isNull()) >= 3,
                     120, "缩略图")
    expect(got, "缩略图已填充（≥3 张）")

    print("[4] 勾选联动")
    # 两个文件都清空，才能验证"无勾选时导出禁用"
    page.file_list.setCurrentRow(page.order.index(xls_entry.uid))
    pump(app, 0.2)
    page._check_all(False)
    page.file_list.setCurrentRow(page.order.index(pdf_entry.uid))
    pump(app, 0.2)
    page._check_all(False)
    expect(len(pdf_entry.checked) == 0 and len(xls_entry.checked) == 0, "清空选择生效")
    expect(not page.export_btn.isEnabled(), "无勾选时导出按钮禁用")
    page.grid._cards[0].check.setChecked(True)
    pump(app, 0.2)
    expect(len(pdf_entry.checked) == 1, "单卡勾选写回状态", f"{pdf_entry.checked}")
    page._invert()
    pump(app, 0.2)
    expect(len(pdf_entry.checked) == 5, "反选后剩 5 页", f"{len(pdf_entry.checked)}")

    print("[5] 通过界面导出 PNG")
    page._check_all(False)
    for key in ("p1", "p3"):
        next(c for c in page.grid._cards if c.unit.key == key).check.setChecked(True)
    pump(app, 0.2)
    page.fmt_box.setCurrentIndex(page.fmt_box.findData("png"))
    page.dpi_box.setCurrentIndex(page.dpi_box.findData(150))
    page.opt_long.setChecked(False)      # 别被注册表里的旧偏好带跑：这里要的是逐页出图
    page.out_dir.setText(str(OUT))
    page._sync_option_state()
    page._start_export()
    done = wait_until(app, lambda: page._export_worker is None and page.progress.value() == 100, 240, "导出完成")
    expect(done, "导出流程结束")
    files = sorted(OUT.glob("*.png"))
    expect(len(files) == 2, "生成 2 张 PNG", f"{[f.name for f in files]}")
    expect(all(f.stat().st_size > 20_000 for f in files), "图片体积正常",
           f"{[f.stat().st_size // 1024 for f in files]} KB")

    print("[6] 通过界面导出 XLSX")
    page.file_list.setCurrentRow(page.order.index(xls_entry.uid))
    pump(app, 0.4)
    page._check_all(True)
    page.fmt_box.setCurrentIndex(page.fmt_box.findData("xlsx"))
    page._sync_option_state()
    page._start_export()
    done = wait_until(app, lambda: page._export_worker is None, 240, "xlsx 导出")
    expect(done, "xlsx 导出结束")
    books = sorted(OUT.glob("*.xlsx"))
    expect(len(books) == 1, "生成 1 个工作簿", f"{[f.name for f in books]}")

    print("[7] 合并 Word：两份文档各取部分页 + 调顺序")
    page.add_paths([str(SAMPLES / "样例方案.docx"), str(SAMPLES / "样例合同.docx")])
    ok = wait_until(app, lambda: all(not e.loading for e in page.entries.values()), 240, "Word 解析")
    expect(ok, "两份 Word 解析完成")
    doc_a = next(e for e in page.entries.values() if e.name == "样例方案.docx")
    doc_b = next(e for e in page.entries.values() if e.name == "样例合同.docx")
    expect(len(doc_a.units) == 3 and len(doc_b.units) == 4, "页数正确（3 / 4）",
           f"{len(doc_a.units)} / {len(doc_b.units)}")

    # 全部清空，然后精确地勾：A 的第 1、3 页，B 的第 2 页
    for e in (pdf_entry, xls_entry, doc_a, doc_b):
        page.file_list.setCurrentRow(page.order.index(e.uid))
        pump(app, 0.15)
        page._check_all(False)
    page.file_list.setCurrentRow(page.order.index(doc_a.uid))
    pump(app, 0.2)
    for key in ("p1", "p3"):
        next(c for c in page.grid._cards if c.unit.key == key).check.setChecked(True)
    page.file_list.setCurrentRow(page.order.index(doc_b.uid))
    pump(app, 0.2)
    next(c for c in page.grid._cards if c.unit.key == "p2").check.setChecked(True)
    pump(app, 0.2)
    expect(len(page.ordered) == 3, "导出顺序里有 3 页", f"{page.ordered}")
    expect([k for _u, k in page.ordered] == ["p1", "p3", "p2"], "默认按勾选先后排序",
           f"{[k for _u, k in page.ordered]}")

    # 把最后一页（B 的第 2 页）挪到最前面
    page.order_list.setCurrentRow(2)
    page._move_order(-1)
    page._move_order(-1)
    pump(app, 0.2)
    expect([k for _u, k in page.ordered] == ["p2", "p1", "p3"], "前移生效",
           f"{[k for _u, k in page.ordered]}")
    expect(page.order_list.item(0).text().startswith("1."), "序号跟着刷新",
           page.order_list.item(0).text())

    page.fmt_box.setCurrentIndex(page.fmt_box.findData("docx"))
    page._sync_option_state()
    expect(page.opt_merge_docx.isVisibleTo(page), "Word 格式下显示合并开关")
    page.opt_merge_docx.setChecked(True)
    page._start_export()
    done = wait_until(app, lambda: page._export_worker is None and page.progress.value() == 100, 300, "合并完成")
    expect(done, "合并流程结束")
    merged = list(OUT.glob("*.docx"))
    expect(len(merged) == 1, "只生成 1 份 Word", f"{[f.name for f in merged]}")
    if merged:
        import pymupdf
        from core.office import word_to_pdf
        pdf = pymupdf.open(word_to_pdf(merged[0].resolve(), force=True))
        expect(pdf.page_count == 3, "合并结果 3 页", f"{pdf.page_count}")
        heads = [pdf[i].get_text().strip().split("\n")[0] for i in range(pdf.page_count)]
        expect(heads[0].startswith("第 2 章 合同条款"), "第 1 页是「合同条款 第2章」", heads[0])
        expect(heads[1].startswith("第 1 章 项目实施方案"), "第 2 页是「项目实施方案 第1章」", heads[1])
        expect(heads[2].startswith("第 3 章 项目实施方案"), "第 3 页是「项目实施方案 第3章」", heads[2])
        pdf.close()

    print("[8] 图片转 ICO 页签")
    win._switch(1)
    pump(app, 0.2)
    expect(win.stack.currentWidget() is win.ico_page, "切换到 ICO 工具页")
    expect(win.btn_ico.isChecked() and not win.btn_extract.isChecked(), "导航按钮选中态正确")
    ico_page = win.ico_page
    img_path = OUT / "ico_src.png"
    from PIL import Image, ImageDraw
    im = Image.new("RGB", (300, 220), (90, 140, 220))
    ImageDraw.Draw(im).ellipse([60, 40, 240, 180], fill=(250, 210, 90))
    im.save(img_path)
    ico_page.load_path(str(img_path))
    ico_page.radius_slider.setValue(20)
    ico_page._refresh()
    expect(ico_page.original_small is not None, "ICO 页加载图片成功")
    ico_out = OUT / "smoke.ico"
    sizes = ico_page._export_to(str(ico_out))
    expect(len(sizes) == 7, "导出 7 档尺寸 ICO", f"{sizes}")
    expect(ico_out.stat().st_size > 0, "ICO 文件已生成", f"{ico_out.stat().st_size} B")
    win._switch(0)
    pump(app, 0.2)
    expect(win.stack.currentWidget() is page, "切回页面提取")

    print("[9] 移除与清空")
    page.file_list.setCurrentRow(0)
    page._remove_current()
    pump(app, 0.3)
    expect(len(page.order) == 3, "移除后剩 3 个文件", f"{len(page.order)}")
    page._clear_all()
    pump(app, 0.3)
    expect(len(page.order) == 0, "清空队列生效")
    expect(not page.export_btn.isEnabled(), "清空后导出按钮禁用")

    win.close()
    pump(app, 0.3)

    print()
    if FAILS:
        print(f"GUI 冒烟测试失败 {len(FAILS)} 项：{FAILS}")
        return 1
    print("GUI 冒烟测试全部通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
