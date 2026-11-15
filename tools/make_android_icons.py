# -*- coding: utf-8 -*-
"""从 app_icon.ico 生成 Android 启动图标资源（mipmap 各密度 + 自适应图标）。

产物写入 android/app/src/main/res/ 下：
* mipmap-mdpi/hdpi/xhdpi/xxhdpi/xxxhdpi/ic_launcher.png（传统方形）
* 同目录 ic_launcher_round.png（圆形裁剪）
* 同目录 ic_launcher_foreground.png（自适应图标前景，内容缩到安全区）
自适应图标背景色由资源 android:background 引用 colors.xml 里的 launcher_bg。
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
ICO = ROOT / "app_icon.ico"
RES = ROOT / "android" / "app" / "src" / "main" / "res"

DENSITIES = {
    "mdpi": 1,
    "hdpi": 1.5,
    "xhdpi": 2,
    "xxhdpi": 3,
    "xxxhdpi": 4,
}
LEGACY_DP = 48          # 传统启动图标 48dp
ADAPTIVE_DP = 108       # 自适应图层 108dp
SAFE_RATIO = 0.62       # 前景内容占画布比例（启动器只显示中间 72dp/108dp ≈ 0.67，超过会切到帽子边缘）


def load_icon() -> Image.Image:
    """读 .ico 里最大的一帧，转 RGBA。"""
    im = Image.open(ICO)
    try:
        sizes = list(im.ico.sizes())
    except Exception:
        sizes = [im.size]
    best = max(sizes, key=lambda s: s[0] * s[1])
    try:
        im.size = best
        im.load()
    except Exception:
        pass
    img = im.convert("RGBA")
    if img.size[0] < 256:
        img = img.resize((256, 256), Image.LANCZOS)
    return img


def border_color(img: Image.Image) -> tuple[int, int, int]:
    """图标边缘最常出现的颜色，作为自适应图标背景色。

    用众数而不是平均值：边缘若有抗锯齿过渡像素，平均值会被拉灰；
    众数能拿到画布本身的底色（例如白底手绘稿的 #fefefe），
    背景层与画稿底色一致时，画稿的方块边缘才不会在圆形遮罩里露出来。
    """
    w, h = img.size
    px = img.convert("RGBA").load()
    cols = []
    for i in range(0, w, max(1, w // 64)):
        for y in (0, h - 1):
            r, g, b, a = px[i, y]
            if a > 128:
                cols.append((r, g, b))
    for i in range(0, h, max(1, h // 64)):
        for x in (0, w - 1):
            r, g, b, a = px[x, i]
            if a > 128:
                cols.append((r, g, b))
    if not cols:
        return (255, 255, 255)
    from collections import Counter
    return Counter(cols).most_common(1)[0][0]


def rounded(img: Image.Image) -> Image.Image:
    """圆形裁剪（用于 ic_launcher_round）。"""
    side = min(img.size)
    sq = img.crop(((img.size[0] - side) // 2, (img.size[1] - side) // 2,
                   (img.size[0] + side) // 2, (img.size[1] + side) // 2))
    mask = Image.new("L", (side * 4, side * 4), 0)
    from PIL import ImageDraw
    ImageDraw.Draw(mask).ellipse([0, 0, side * 4 - 1, side * 4 - 1], fill=255)
    mask = mask.resize((side, side), Image.LANCZOS)
    out = Image.new("RGBA", (side, side), (0, 0, 0, 0))
    out.paste(sq, (0, 0), mask)
    return out


def main() -> None:
    icon = load_icon()
    bg = border_color(icon)
    bg_hex = "#%02x%02x%02x" % bg
    print(f"自适应背景色: {bg_hex}")

    # 传统方形 + 圆形
    for name, folder in (("ic_launcher.png", "mipmap-%s"), ("ic_launcher_round.png", "mipmap-%s")):
        for dpi, scale in DENSITIES.items():
            side = round(LEGACY_DP * scale)
            target = RES / (folder % dpi)
            target.mkdir(parents=True, exist_ok=True)
            out = icon.resize((side, side), Image.LANCZOS)
            if name.endswith("round.png"):
                out = rounded(out)
            out.save(target / name)

    # 自适应前景：画布先铺满背景色，再把画稿贴上去。
    # 若画布留透明，画稿自带的方块边缘（半透明抗锯齿）叠在白底上会露出一圈浅灰轮廓。
    for dpi, scale in DENSITIES.items():
        canvas_side = round(ADAPTIVE_DP * scale)
        content = round(canvas_side * SAFE_RATIO)
        fg = icon.resize((content, content), Image.LANCZOS)
        canvas = Image.new("RGBA", (canvas_side, canvas_side), bg + (255,))
        canvas.paste(fg, ((canvas_side - content) // 2, (canvas_side - content) // 2), fg)
        target = RES / ("mipmap-%s" % dpi)
        target.mkdir(parents=True, exist_ok=True)
        canvas.save(target / "ic_launcher_foreground.png")

    # 把背景色写进 colors.xml（保留其它条目）
    colors_path = RES / "values" / "colors.xml"
    colors_path.parent.mkdir(parents=True, exist_ok=True)
    text = colors_path.read_text(encoding="utf-8") if colors_path.exists() else ""
    import re
    if "launcher_bg" in text:
        text = re.sub(r'<color name="launcher_bg">[^<]*</color>',
                      f'<color name="launcher_bg">{bg_hex}</color>', text)
    else:
        text = text.replace("</resources>", f'    <color name="launcher_bg">{bg_hex}</color>\n</resources>') \
            if "</resources>" in text else f'<?xml version="1.0" encoding="utf-8"?>\n<resources>\n    <color name="launcher_bg">{bg_hex}</color>\n</resources>\n'
    colors_path.write_text(text, encoding="utf-8")
    print("图标资源生成完毕")


if __name__ == "__main__":
    main()
