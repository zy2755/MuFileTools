"""离屏 GUI 冒烟测试：构建界面 → 入队 → 等解析 → 勾选 → 导出 → 校验。

运行： python tests/gui_smoke.py
不弹窗、不需要桌面会话（QT_QPA_PLATFORM=offscreen）。
"""
from __future__ import annotations

import hashlib
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
import ui.toc_page as tp  # noqa: E402
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

    namespaces = [("MuFileTools", "PageExtract"), ("MuFileTools", "TocPage")]
    snapshots = {}
    for org, app_name in namespaces:
        store = QSettings(org, app_name)
        snapshots[store] = {k: store.value(k) for k in store.allKeys()}
    try:
        return _run()
    finally:
        for store, snap in snapshots.items():
            for k in store.allKeys():
                if k not in snap:
                    store.remove(k)
            for k, v in snap.items():
                store.setValue(k, v)
            store.sync()


def _run() -> int:

    # 三个模块都换掉 QMessageBox，模态框全部变空操作
    ep.QMessageBox = _FakeBox  # type: ignore[assignment]
    ip.QMessageBox = _FakeBox  # type: ignore[assignment]
    mw.QMessageBox = _FakeBox  # type: ignore[assignment]
    tp.QMessageBox = _FakeBox  # type: ignore[assignment]

    app = QApplication(sys.argv)
    app.setStyleSheet(QSS)

    print("[1] 构建主窗口")
    win = MainWindow()
    win.show()
    pump(app, 0.5)
    page = win.extract_page
    expect(win.isVisible(), "窗口已创建")
    expect(win.windowTitle() == "小睦文件转换器", "窗口标题为「小睦文件转换器」", win.windowTitle())
    expect(win.stack.count() == 3, "壳窗口包含 3 个工具页", f"count={win.stack.count()}")
    expect(len(win._nav_buttons) == 3, "顶栏有 3 个导航按钮", f"{len(win._nav_buttons)}")
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

    print("[7b] 合并 + 自动插入目录")
    toc_dir = OUT / "toc_merge"
    toc_dir.mkdir(parents=True, exist_ok=True)
    page.out_dir.setText(str(toc_dir))
    page.opt_toc.setChecked(True)
    page.toc_depth.setCurrentIndex(page.toc_depth.findData(3))
    page._sync_option_state()
    expect(page.toc_row.isVisibleTo(page), "勾了合并后显示目录选项")
    expect(page._build_spec().toc_levels == 3, "spec 带上 toc_levels=3",
           str(page._build_spec().toc_levels))
    page._start_export()
    done = wait_until(app, lambda: page._export_worker is None and page.progress.value() == 100,
                      300, "合并+目录完成")
    expect(done, "合并+目录流程结束")
    with_toc = list(toc_dir.glob("*.docx"))
    expect(len(with_toc) == 1, "只生成 1 份带目录的 Word", f"{[f.name for f in with_toc]}")
    if with_toc:
        import pymupdf
        from core.office import word_to_pdf
        pdf = pymupdf.open(word_to_pdf(with_toc[0].resolve(), force=True))
        expect(pdf.page_count == 4, "3 页正文 + 1 页目录 = 4 页", f"{pdf.page_count}")
        first = pdf[0].get_text()
        expect("目录" in first, "第 1 页是目录页", first[:40])
        expect("第 2 章 合同条款" in first and "第 1 章 项目实施方案" in first,
               "目录里收录了两章的标题", first[:120])
        expect(pdf[1].get_text().strip().startswith("第 2 章 合同条款"), "第 2 页才是正文首页",
               pdf[1].get_text()[:40])
        pdf.close()
    page.opt_toc.setChecked(False)
    page.out_dir.setText(str(OUT))
    page._sync_option_state()

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

    print("[9] Word 目录页签")
    win._switch(2)
    pump(app, 0.2)
    expect(win.stack.currentWidget() is win.toc_page, "切换到 Word 目录页")
    expect(win.btn_toc.isChecked() and not win.btn_extract.isChecked()
           and not win.btn_ico.isChecked(), "三个导航按钮互斥选中正确")
    tpage = win.toc_page
    tpage.load_path(str(SAMPLES / "样例方案.docx"))
    pump(app, 0.2)
    expect(len(tpage._rows) == 6, "识别到 6 条标题", f"{len(tpage._rows)}")
    expect("一级 3" in tpage.stat_label.text() and "二级 3" in tpage.stat_label.text(),
           "统计文案正确", tpage.stat_label.text())
    expect(tpage.btn_go.isEnabled(), "有标题时可生成")
    expect(not tpage.warn_label.isVisibleTo(tpage), "有标题时不显示警告")

    # 没有标题样式的文档 → 红字 + 主按钮禁用
    from docx import Document as _Doc

    plain = OUT / "无标题.docx"
    plain_doc = _Doc()
    plain_doc.add_paragraph("第一段普通正文，没有套标题样式。")
    plain_doc.add_paragraph("第二段也一样。")
    plain_doc.save(str(plain))
    tpage.load_path(str(plain))
    pump(app, 0.2)
    expect(len(tpage._rows) == 0, "无标题文档扫到 0 条", f"{len(tpage._rows)}")
    expect(not tpage.btn_go.isEnabled(), "无标题时主按钮禁用")
    expect(tpage.warn_label.isVisibleTo(tpage), "无标题时显示红字提示")

    # 非 docx 也拦住
    tpage.load_path(str(SAMPLES / "样例报告.pdf"))
    pump(app, 0.2)
    expect(not tpage.btn_go.isEnabled(), "拖入 pdf 时主按钮禁用")

    # 真生成（走同步入口，避开线程时序）
    tpage.load_path(str(SAMPLES / "样例方案.docx"))
    pump(app, 0.2)
    src_md5 = hashlib.md5((SAMPLES / "样例方案.docx").read_bytes()).hexdigest()
    res = tpage._generate_to(str(OUT / "样例方案_带目录.docx"))
    expect(res["entries"] == 6, "生成 6 条目录", str(res.get("entries")))
    expect(hashlib.md5((SAMPLES / "样例方案.docx").read_bytes()).hexdigest() == src_md5,
           "源文件一个字节都没动")
    chk = _Doc(str(res["out"]))
    expect(chk.paragraphs[0].text.strip() == "目录", "首段是「目录」标题行",
           chk.paragraphs[0].text)
    expect("第 1 章" in chk.paragraphs[1].text, "次段是目录正文", chk.paragraphs[1].text[:40])

    print("[9b] 手动标标题")
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt

    from core import outline as _outline
    from ui.mark_dialog import MarkDialog

    hand = OUT / "手打标题.docx"
    hd = _Doc()
    cover = hd.add_paragraph()
    cover.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cr = cover.add_run("手打大标题")
    cr.bold = True
    cr.font.size = Pt(22)
    for lv, txt in ((1, "1 项目简介"), (2, "1.1 项目介绍"), (3, "1.1.1 平台定位"),
                    (0, "正文内容" * 20)):
        para = hd.add_paragraph()
        run = para.add_run(txt)
        if lv:
            run.bold = True
            run.font.size = Pt({1: 16, 2: 14, 3: 12}[lv])
    hd.save(str(hand))

    tpage.load_path(str(hand))
    pump(app, 0.2)
    expect(len(tpage._rows) == 0, "手打标题的文档扫到 0 条", f"{len(tpage._rows)}")
    expect(tpage.btn_mark.isEnabled(), "「手动标标题」按钮可用")
    expect(not tpage.btn_go.isEnabled(), "没标题时主按钮仍禁用")

    paras = _outline.list_paragraphs(hand)
    guessed = _outline.guess_levels(paras)
    expect(len(guessed) >= 3, "自动预判出标题", f"{len(guessed)}")

    dlg = MarkDialog(paras, guessed, hand.name, win)
    expect(dlg.table.rowCount() == len(paras), "对话框列出全部段落",
           f"{dlg.table.rowCount()} / {len(paras)}")
    expect(dlg.marks() == guessed, "对话框初始值 = 预判结果")
    dlg._clear_guess()
    expect(len(dlg.marks()) == 0, "「清除自动预判」把标记清空了")

    for i in (1, 2, 3):
        dlg._set_level(paras[i].index, i)
    got = dlg.marks()
    expect(got.get(paras[1].index) == 1 and got.get(paras[3].index) == 3,
           "手动设置 1/2/3 级生效", str(got))
    rows = dlg.marked_rows()
    expect(len(rows) == 3 and rows[0] == (1, "1 项目简介"), "marked_rows 顺序正确", str(rows))
    dlg.close()

    tpage._marks = got
    tpage._rows = rows
    tpage._edited = True
    tpage._refresh_view()
    pump(app, 0.2)
    expect(tpage.btn_go.isEnabled(), "标记之后主按钮可用")
    expect("手动标记" in tpage.file_label.text(), "文件区标注「手动标记」",
           tpage.file_label.text())
    expect("一级 1" in tpage.stat_label.text(), "统计行跟着变", tpage.stat_label.text())

    res2 = tpage._generate_to(str(OUT / "手打标题_带目录.docx"))
    expect(res2["entries"] == 3, "按标记生成 3 条目录", str(res2.get("entries")))
    chk2 = _Doc(str(res2["out"]))
    heads2 = [p.text.strip() for p in chk2.paragraphs[:3]]
    expect("目录" in heads2, "目录插在封面之后（不在封面前）", str(heads2))
    expect(heads2.index("目录") == 1, "页序：封面 → 目录 → 正文", str(heads2))

    win._switch(0)
    pump(app, 0.2)

    print("[10] 移除与清空")
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
