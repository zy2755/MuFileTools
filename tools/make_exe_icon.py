# -*- coding: utf-8 -*-
"""生成打包 exe 用的程序图标 app_icon.ico（蓝色圆角底 + 白色 ICO 字样）。"""

from PIL import Image, ImageDraw, ImageFont

SIZE = 256
img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
draw = ImageDraw.Draw(img)
draw.rounded_rectangle([8, 8, SIZE - 8, SIZE - 8], radius=56, fill=(36, 101, 232, 255))

font = ImageFont.load_default(96)
left, top, right, bottom = draw.textbbox((0, 0), "ICO", font=font)
draw.text(
    ((SIZE - (right - left)) / 2 - left, (SIZE - (bottom - top)) / 2 - top),
    "ICO",
    font=font,
    fill=(255, 255, 255, 255),
)

img.save(
    "app_icon.ico",
    format="ICO",
    sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)],
)
print("app_icon.ico generated")
