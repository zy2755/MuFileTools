# -*- coding: utf-8 -*-
"""拖放功能测试：模拟把文件拖到窗口各处，验证应用级路由是否真的生效。

这个功能手工验证很别扭（要真的拿鼠标拖），所以用 QApplication.sendEvent
合成 DragEnter / Drop 事件直接打到各个控件上——包括那些天生会吞拖放的
输入框和日志框。
"""
from __future__ import annotations

import os
import sys
import tempfile
import traceback
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import fitz
from PIL import Image
from PySide6.QtCore import QEvent, QMimeData, QPoint, QPointF, Qt, QUrl
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import QApplication

from ui import theme
from ui.shell import MainWindow

TMP = Path(tempfile.mkdtemp(prefix="dnd_probe_"))
ok = 0
fail = 0


def check(name: str, cond: bool, extra: str = "") -> None:
    global ok, fail
    if cond:
        ok += 1
        print(f"  PASS  {name}" + (f"   ({extra})" if extra else ""))
    else:
        fail += 1
        print(f"  FAIL  {name}" + (f"   ({extra})" if extra else ""))


def make_pdf(n_pages: int = 2) -> str:
    p = TMP / f"doc{n_pages}.pdf"
    doc = fitz.open()
    for i in range(n_pages):
        page = doc.new_page()
        page.insert_text((72, 100), f"Probe page {i + 1}", fontsize=24)
    doc.save(str(p))
    doc.close()
    return str(p)


def make_png() -> str:
    p = TMP / "pic.png"
    Image.new("RGB", (256, 256), (40, 110, 220)).save(p)
    return str(p)


def make_txt() -> str:
    p = TMP / "note.txt"
    p.write_text("hello", encoding="utf-8")
    return str(p)


def drag(target, paths, drop=True):
    md = QMimeData()
    md.setUrls([QUrl.fromLocalFile(p) for p in paths])
    enter = QDragEnterEvent(
        QPoint(30, 30),
        Qt.DropAction.CopyAction,
        md,
        Qt.MouseButton.LeftButton,
        Qt.KeyboardModifier.NoModifier,
    )
    QApplication.sendEvent(target, enter)
    if drop:
        ev = QDropEvent(
            QPointF(30, 30),
            Qt.DropAction.CopyAction,
            md,
            Qt.MouseButton.LeftButton,
            Qt.KeyboardModifier.NoModifier,
        )
        QApplication.sendEvent(target, ev)


def reset(win: MainWindow) -> None:
    win._switch(0)
    win.extract_page._clear_all()
    win.ico_page.original_small = None
    win.ico_page.source_path = None
    win.extract_page.out_dir.setText("原始目录-不该被改")
    win.extract_page.log.clear()


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("MuFileTools")
    app.setOrganizationName("MuFileTools")

    # ExtractPage 用 QSettings("MuFileTools", "PageExtract") 存偏好，而这个测试会
    # 改「输出目录」并在关闭时被写回。这里先快照真实配置，结束时原样还原，
    # 免得把测试用的假路径留在用户的注册表里（踩过一次）。
    from PySide6.QtCore import QSettings
    real = QSettings("MuFileTools", "PageExtract")
    snapshot = {k: real.value(k) for k in real.allKeys()}

    try:
        return _run(app)
    finally:
        for k in real.allKeys():
            if k not in snapshot:
                real.remove(k)
        for k, v in snapshot.items():
            real.setValue(k, v)
        real.sync()


def _run(app) -> int:
    theme.apply(app, theme.MODE_DARK)

    win = MainWindow()
    win.resize(1280, 820)
    win.show()
    app.processEvents()

    pdf = make_pdf()
    png = make_png()
    txt = make_txt()
    folder = TMP / "批量"
    folder.mkdir(exist_ok=True)
    for i in range(3):
        Image.new("RGB", (32, 32), (200, 60, 60)).save(folder / f"a{i}.png")

    ep = win.extract_page
    ip = win.ico_page

    print("\n[1] 页面提取页 —— 拖到不同控件上都应该收下")
    targets = [
        ("窗口中央区", win.centralWidget()),
        ("页面本身", ep),
        ("文件列表", ep.file_list),
        ("预览网格", ep.grid.viewport()),
        ("日志框(会吞文字的控件)", ep.log),
        ("输出目录输入框(会吞文字的控件)", ep.out_dir),
        ("输出格式下拉框", ep.fmt_box),
        ("左侧面板空白", ep),
        ("顶部导航栏", win.btn_extract),
        ("状态栏文字", ep.status_label),
    ]
    for name, target in targets:
        reset(win)
        drag(target, [pdf])
        check(name, len(ep.order) == 1, f"队列 {len(ep.order)}")
        if name.startswith("输出目录"):
            check("  输入框没有被写入路径", ep.out_dir.text() == "原始目录-不该被改", ep.out_dir.text()[:30])
        if name.startswith("日志"):
            check("  日志框没有插入路径文字", "doc2.pdf" not in ep.log.toPlainText())

    print("\n[2] 悬停提示层")
    reset(win)
    win.drop_router.overlay.hide()
    drag(win.centralWidget(), [pdf, png], drop=False)
    check("拖入时提示层出现", win.drop_router.overlay.isVisible())
    check("提示文案含文件数", "1 个文件" in win.drop_router.overlay._title, win.drop_router.overlay._title)
    from PySide6.QtGui import QFontMetrics
    from ui.dnd import title_font_for
    tf = title_font_for(win.drop_router.overlay.font())
    check("标题字号没有被算成小字", QFontMetrics(tf).height() >= 20, f"高 {QFontMetrics(tf).height()}px")
    drag(win.centralWidget(), [pdf], drop=True)
    check("松手后提示层消失", not win.drop_router.overlay.isVisible())

    print("\n[3] 拖文件夹（递归展开子文件）")
    reset(win)
    drag(win.centralWidget(), [str(folder)])
    check("  图片被识别（切到 ICO 页）", win.stack.currentIndex() == 1)
    check("  已载入其中一张", ip.source_path is not None, str(ip.source_path))

    print("\n[4] 拖图片到「页面提取」页 —— 自动切到 ICO 工具")
    reset(win)
    drag(ep.file_list, [png])
    check("切到 ICO 页", win.stack.currentIndex() == 1)
    check("图片已载入", ip.source_path == png, str(ip.source_path))

    print("\n[5] 图片转 ICO 页")
    reset(win)
    win._switch(1)
    drag(ip.src_label, [png])
    check("拖到原图占位区即可载入", ip.source_path == png)
    reset(win)
    win._switch(1)
    drag(ip.btn_choose, [png])
    check("拖到顶部按钮上也载入", ip.source_path == png)
    reset(win)
    win._switch(1)
    drag(ip.radius_slider, [png])
    check("拖到滑杆上也载入", ip.source_path == png)

    print("\n[6] 文档拖到 ICO 页 —— 切回页面提取")
    reset(win)
    win._switch(1)
    drag(ip.ico_label, [pdf])
    check("切回页面提取页", win.stack.currentIndex() == 0)
    check("文档已加入队列", len(ep.order) == 1)

    print("\n[7] 不支持的类型 / 多个文件")
    reset(win)
    drag(ep.file_list, [txt])
    check("txt 被拒但给了提示", len(ep.order) == 0 and "不支持" in ep.status_label.text(), ep.status_label.text()[:40])
    reset(win)
    win._switch(1)
    drag(ip.src_label, [png, make_png()])
    check("ICO 页多图只载入第一张且提示", ip.source_path is not None and "一次只能转换一张" in ip.status.text(), ip.status.text()[:40])
    reset(win)
    drag(ep.file_list, [pdf, pdf])
    check("同一文件拖两次不重复入队", len(ep.order) == 1, f"队列 {len(ep.order)}")

    print("\n[8] 截图：悬停提示层外观")
    reset(win)
    win._switch(0)
    app.processEvents()
    drag(win.centralWidget(), [pdf], drop=False)
    app.processEvents()
    shot = TMP / "overlay_dark.png"
    win.grab().save(str(shot))
    print("  已保存", shot)
    theme.apply(app, theme.MODE_LIGHT)
    app.processEvents()
    drag(win.centralWidget(), [pdf], drop=False)
    app.processEvents()
    shot2 = TMP / "overlay_light.png"
    win.grab().save(str(shot2))
    print("  已保存", shot2)

    print(f"\n结果：{ok} 通过 / {fail} 失败")
    win.close()
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        traceback.print_exc()
        sys.exit(2)
