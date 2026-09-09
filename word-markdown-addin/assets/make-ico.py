"""生成多尺寸 .ico 图标（供 exe 使用）"""
import os
import sys

sys.path.insert(0, r"C:\Users\Miracle R\.workbuddy\binaries\python\envs\default\Lib\site-packages")

from PIL import Image, ImageDraw, ImageFont

OUT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets"))

C1 = (76, 141, 255)
C2 = (123, 92, 255)
WHITE = (255, 255, 255, 255)


def gradient(size):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    for y in range(size):
        for x in range(size):
            t = (x + y) / (2 * (size - 1)) if size > 1 else 0
            r = int(C1[0] + (C2[0] - C1[0]) * t)
            g = int(C1[1] + (C2[1] - C1[1]) * t)
            b = int(C1[2] + (C2[2] - C1[2]) * t)
            d.point((x, y), (r, g, b, 255))
    return img


def rounded_mask(size, radius_ratio=0.22):
    mask = Image.new("L", (size, size), 0)
    d = ImageDraw.Draw(mask)
    r = max(1, int(size * radius_ratio))
    d.rounded_rectangle([0, 0, size - 1, size - 1], radius=r, fill=255)
    return mask


def make_icon(size):
    base = gradient(size)
    mask = rounded_mask(size)
    icon = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    icon.paste(base, (0, 0), mask)

    d = ImageDraw.Draw(icon)
    w = max(2, size // 12)
    cx = size // 2
    top = int(size * 0.26)
    bottom = int(size * 0.62)
    d.line([(cx, top), (cx, bottom)], fill=WHITE, width=w)
    ah = max(3, size // 7)
    aw = max(3, size // 7)
    d.polygon([(cx - aw, bottom - ah), (cx + aw, bottom - ah), (cx, bottom + int(ah * 0.2))], fill=WHITE)
    bar_y = int(size * 0.78)
    d.line([(int(size * 0.3), bar_y), (int(size * 0.7), bar_y)], fill=WHITE, width=w)
    return icon


sizes = [16, 24, 32, 48, 64, 128, 256]
images = [make_icon(s) for s in sizes]

ico_path = os.path.join(OUT, "icon.ico")
images[-1].save(ico_path, format="ICO", sizes=[(s, s) for s in sizes])
print(f"生成 {ico_path} ({os.path.getsize(ico_path)} bytes)，尺寸：{sizes}")
