# -*- coding: utf-8 -*-
"""ICO 生成核心逻辑（「图片转 ICO」工具的算法层）。

从图片生成 Windows .ico：读取 → 纠正 EXIF 方向 → 正方形化 → 逐档缩放加圆角
→ 打包为含七档尺寸的 .ico 文件。全部为纯 Pillow 函数，不依赖任何界面框架。
"""

from __future__ import annotations

from PIL import Image, ImageChops, ImageDraw, ImageOps

RESAMPLE = Image.Resampling.LANCZOS

ICO_SIZES = (256, 128, 64, 48, 32, 24, 16)  # 导出的七档尺寸（由大到小）
PREVIEW_SIZES = (64, 32, 16)                # 预览区额外展示的小尺寸
SOURCE_MAX = 1024                           # 参与处理的图像最大边长
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp", ".tif", ".tiff"}
FILE_FILTER = "图片文件 (*.jpg *.jpeg *.png *.bmp *.gif *.webp *.tif *.tiff);;所有文件 (*.*)"


def load_source(path: str) -> Image.Image:
    """读取图片，纠正 EXIF 方向，统一转为 RGBA。"""
    with Image.open(path) as im:
        im = ImageOps.exif_transpose(im)
        im.load()
        return im.convert("RGBA")


def downscale(img: Image.Image, max_side: int = SOURCE_MAX) -> Image.Image:
    """等比缩小到最大边长以内（只缩小，不放大）。"""
    width, height = img.size
    longest = max(width, height)
    if longest <= max_side:
        return img
    scale = max_side / longest
    return img.resize((max(1, round(width * scale)), max(1, round(height * scale))), RESAMPLE)


def to_square(img: Image.Image, crop: bool = True) -> Image.Image:
    """变成正方形：crop=True 居中裁剪，否则用透明像素补边。"""
    width, height = img.size
    if width == height:
        return img
    if crop:
        side = min(width, height)
        left, top = (width - side) // 2, (height - side) // 2
        return img.crop((left, top, left + side, top + side))
    side = max(width, height)
    canvas = Image.new("RGBA", (side, side), (0, 0, 0, 0))
    canvas.paste(img, ((side - width) // 2, (side - height) // 2))
    return canvas


def fit_square(square: Image.Image, side: int) -> Image.Image:
    """正方形图缩放到指定边长（必要时允许放大）。"""
    if square.size[0] == side:
        return square
    return square.resize((side, side), RESAMPLE)


def fit_contain(img: Image.Image, box: int) -> Image.Image:
    """等比缩放到能放进 box×box 的框内（允许放大），保持原始宽高比。"""
    width, height = img.size
    scale = min(box / width, box / height)
    if scale == 1.0:
        return img.copy()
    return img.resize((max(1, round(width * scale)), max(1, round(height * scale))), RESAMPLE)


def rounded_mask(size: int, radius_pct: float) -> Image.Image:
    """抗锯齿圆角蒙版：按 4 倍超采样绘制圆角矩形后缩小回目标尺寸。"""
    pct = max(0.0, min(50.0, float(radius_pct)))
    big = size * 4
    mask = Image.new("L", (big, big), 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle([0, 0, big - 1, big - 1], radius=big * pct / 100.0, fill=255)
    return mask.resize((size, size), RESAMPLE)


def apply_round(img: Image.Image, radius_pct: float) -> Image.Image:
    """给 RGBA 图加圆角：圆角外透明；图片自身已有的透明区域保持不变。"""
    if radius_pct <= 0:
        return img
    side = img.size[0]
    if img.size[1] != side:
        img = to_square(img)
    out = img.copy()
    out.putalpha(ImageChops.darker(out.getchannel("A"), rounded_mask(side, radius_pct)))
    return out


def build_icon_set(square: Image.Image, radius_pct: float) -> dict[int, Image.Image]:
    """每一档尺寸分别缩放、加圆角，保证小尺寸的圆角比例一致、边缘清晰。"""
    icons: dict[int, Image.Image] = {}
    for size in ICO_SIZES:
        icons[size] = apply_round(fit_square(square, size), radius_pct)
    return icons


def save_ico(icons: dict[int, Image.Image], path: str) -> None:
    """把各尺寸写入单个 .ico 文件。"""
    order = sorted(icons, reverse=True)
    icons[order[0]].save(
        path,
        format="ICO",
        append_images=[icons[size] for size in order[1:]],
        sizes=[(size, size) for size in order],
    )


def checkerboard(side: int, cell: int = 8) -> Image.Image:
    """灰白棋盘格底图，用于衬托透明区域。"""
    board = Image.new("RGB", (side, side), (255, 255, 255))
    draw = ImageDraw.Draw(board)
    grey = (204, 204, 204)
    for y in range(0, side, cell):
        for x in range(0, side, cell):
            if (x // cell + y // cell) % 2:
                draw.rectangle(
                    [x, y, min(x + cell, side) - 1, min(y + cell, side) - 1], fill=grey
                )
    return board


def on_checkerboard(icon: Image.Image, cell: int = 8) -> Image.Image:
    """把透明底图标合成到同尺寸棋盘格上，得到可直接显示的预览图。"""
    board = checkerboard(icon.size[0], cell).convert("RGBA")
    board.alpha_composite(icon)
    return board
