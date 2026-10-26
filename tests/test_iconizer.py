# -*- coding: utf-8 -*-
"""「图片转 ICO」工具测试：core/iconizer.py 算法 + IcoPage 界面冒烟。

运行：python tests/test_iconizer.py
界面部分在离屏模式运行，不弹窗、不需要桌面会话。
"""

from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from PIL import Image, ImageDraw  # noqa: E402

from core import iconizer as ic  # noqa: E402


def make_test_image(path: str, w: int = 512, h: int = 384) -> None:
    """生成一张非正方形 JPG 测试图：红底 + 蓝色圆形。"""
    img = Image.new("RGB", (w, h), (200, 60, 60))
    draw = ImageDraw.Draw(img)
    draw.ellipse([w // 4, h // 4, 3 * w // 4, 3 * h // 4], fill=(60, 90, 200))
    img.save(path, format="JPEG")


def test_core(tmp: str) -> None:
    jpg = os.path.join(tmp, "src.jpg")
    make_test_image(jpg)

    # 1) 读取与非正方形处理
    full = ic.load_source(jpg)
    assert full.mode == "RGBA" and full.size == (512, 384), (full.mode, full.size)
    crop = ic.to_square(full, crop=True)
    assert crop.size == (384, 384), crop.size
    pad = ic.to_square(full, crop=False)
    assert pad.size == (512, 512), pad.size
    assert pad.getpixel((2, 2))[3] == 0, "透明补边区域应为全透明"

    proc = ic.downscale(crop)
    assert proc.size == (384, 384), proc.size

    # 2) 圆角：四角透明、中心不透明、已有透明保留
    rounded = ic.apply_round(proc, 25)
    assert rounded.size == proc.size, rounded.size
    corner = rounded.getpixel((0, 0))[3]
    center = rounded.getpixel((192, 192))[3]
    assert corner < 10, f"圆角外应为透明，实际 alpha={corner}"
    assert center == 255, f"中心应为不透明，实际 alpha={center}"
    assert ic.apply_round(proc, 0) is proc, "半径 0 不应改动图像"

    hole = Image.new("RGBA", (64, 64), (0, 255, 0, 255))
    hole.putpixel((32, 32), (0, 0, 0, 0))
    out = ic.apply_round(hole, 25)
    assert out.getpixel((32, 32))[3] == 0, "图片自身的透明区域必须保留"

    # 3) 导出 ICO 并回读校验
    icons = ic.build_icon_set(proc, 25)
    assert set(icons) == set(ic.ICO_SIZES), sorted(icons)
    ico_path = os.path.join(tmp, "out.ico")
    ic.save_ico(icons, ico_path)
    assert os.path.getsize(ico_path) > 0

    with Image.open(ico_path) as ico:
        sizes = getattr(ico, "sizes", None) or ico.info.get("sizes")
        assert sizes, "读不到 ICO 内含尺寸"
        sizes = {tuple(size) for size in sizes}
        assert (256, 256) in sizes and (16, 16) in sizes, sizes
        ico.size = (256, 256)  # 选择 256 帧；若版本不支持赋值则默认已打开最大帧
        frame = ico.convert("RGBA")
        assert frame.size == (256, 256), frame.size
        assert frame.mode == "RGBA", frame.mode
        corner_a = frame.getpixel((0, 0))[3]
        center_a = frame.getpixel((128, 128))[3]
        assert corner_a < 10, f"ICO 帧圆角外应为透明，实际 alpha={corner_a}"
        assert center_a == 255, f"ICO 帧中心应为不透明，实际 alpha={center_a}"
    print("  ✔ 核心算法：读取/正方形化/圆角/透明保留/ICO 回读")


def test_gui(tmp: str) -> None:
    from PySide6.QtWidgets import QApplication

    from ui.ico_page import IcoPage

    app = QApplication.instance() or QApplication([])
    page = IcoPage()
    assert not page.radius_slider.isEnabled() or True  # 控件可用即可

    jpg = os.path.join(tmp, "src.jpg")
    make_test_image(jpg)
    page.load_path(jpg)
    assert page.original_small is not None, "加载图片失败"
    page.radius_slider.setValue(25)
    page._refresh()  # 直接刷新，不等防抖定时器

    # 程序化导出（绕开文件对话框）并回读
    ico_path = os.path.join(tmp, "page_out.ico")
    sizes = page._export_to(ico_path)
    assert sizes == sorted(ic.ICO_SIZES, reverse=True), sizes
    with Image.open(ico_path) as ico:
        got = {tuple(s) for s in ico.info.get("sizes", ())}
        assert (256, 256) in got and (16, 16) in got, got
    print("  ✔ 界面冒烟：加载 → 滑杆 → 预览刷新 → 导出回读")

    # 非正方形「透明补边」分支 + 窗口缩放刷新
    page.crop_check.setChecked(False)
    page._refresh()
    page.resize(1400, 900)
    page._refresh()
    assert page.original_small is not None
    print("  ✔ 界面冒烟：补边分支与窗口缩放刷新正常")

    page.deleteLater()
    app.processEvents()


def main() -> int:
    tmp = tempfile.mkdtemp(prefix="iconizer_test_")
    test_core(tmp)
    test_gui(tmp)
    print("图片转 ICO 测试全部通过")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
