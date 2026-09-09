"""生成 Office 加载项图标（16/32/80 px PNG）"""
import os
import sys

sys.path.insert(0, r"C:\Users\Miracle R\.workbuddy\binaries\python\envs\default\Lib\site-packages")

from PIL import Image, ImageDraw, ImageFont

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets")
OUT = os.path.normpath(OUT)
os.makedirs(OUT, exist_ok=True)

# 配色：与任务面板渐变一致
C1 = (76, 141, 255)    # #4c8dff
C2 = (123, 92, 255)    # #7b5cff
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


def load_font(size):
    candidates = [
        r"C:\Windows\Fonts\segoeuib.ttf",
        r"C:\Windows\Fonts\arialbd.ttf",
        r"C:\Windows\Fonts\calibrib.ttf",
        r"C:\Windows\Fonts\msyhbd.ttc",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                continue
    return ImageFont.load_default()


def make_icon(size):
    base = gradient(size)
    mask = rounded_mask(size)
    icon = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    icon.paste(base, (0, 0), mask)

    d = ImageDraw.Draw(icon)
    if size >= 32:
        # 画一个 Markdown 记号：下箭头 + 底部横线
        w = max(2, size // 12)
        cx = size // 2
        top = int(size * 0.26)
        bottom = int(size * 0.62)
        d.line([(cx, top), (cx, bottom)], fill=WHITE, width=w)
        # 箭头
        ah = max(3, size // 7)
        aw = max(3, size // 7)
        d.polygon([(cx - aw, bottom - ah), (cx + aw, bottom - ah), (cx, bottom + int(ah * 0.2))], fill=WHITE)
        # 底部横线
        bar_y = int(size * 0.78)
        d.line([(int(size * 0.3), bar_y), (int(size * 0.7), bar_y)], fill=WHITE, width=w)
    else:
        # 16px：只保留箭头与横线，避免糊成一团
        cx = size // 2
        d.line([(cx, 4), (cx, 9)], fill=WHITE, width=2)
        d.polygon([(cx - 2, 8), (cx + 2, 8), (cx, 11)], fill=WHITE)
        d.line([(4, 13), (size - 5, 13)], fill=WHITE, width=2)

    return icon


for s in (16, 32, 80):
    icon = make_icon(s)
    path = os.path.join(OUT, f"icon-{s}.png")
    icon.save(path, "PNG")
    print(f"生成 {path} ({os.path.getsize(path)} bytes)")
