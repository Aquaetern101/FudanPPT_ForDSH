# -*- coding: utf-8 -*-
"""把逐页 PNG 拼成预览：三列总览 + 逐页长图。

用法：python make_preview.py <png目录> <输出前缀>
产出：<前缀>-sheet.png（三列）、<前缀>-full.png（逐页 1200px 宽）
"""
import glob
import os
import sys

from PIL import Image, ImageDraw, ImageFont

FONT_CANDIDATES = [
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\msyhbd.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
]


def font(size):
    for p in FONT_CANDIDATES:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                continue
    return ImageFont.load_default()


def main():
    png_dir, prefix = sys.argv[1], sys.argv[2]
    files = sorted(glob.glob(os.path.join(png_dir, "*.PNG")) or
                   glob.glob(os.path.join(png_dir, "*.png")))
    if not files:
        raise SystemExit("没有可用的 PNG：%s" % png_dir)

    # 三列总览
    cols, tw, th, pad = 3, 520, 293, 12
    rows = (len(files) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * tw + pad * (cols + 1), rows * th + pad * (rows + 1)), (235, 235, 235))
    f13 = font(13)
    for i, p in enumerate(files):
        im = Image.open(p).convert("RGB").resize((tw, th))
        x = pad + (i % cols) * (tw + pad)
        y = pad + (i // cols) * (th + pad)
        sheet.paste(im, (x, y))
        ImageDraw.Draw(sheet).text((x + 5, y + th - 16), os.path.basename(p)[:-4], font=f13, fill=(255, 255, 255))
    sheet_path = prefix + "-sheet.png"
    sheet.save(sheet_path)

    # 逐页长图
    W, gap, lab = 1200, 18, 26
    H = int(W * 9 / 16)
    full = Image.new("RGB", (W + gap * 2, len(files) * (H + lab) + gap * 2), (245, 245, 245))
    f18 = font(18)
    for i, p in enumerate(files):
        im = Image.open(p).convert("RGB").resize((W, H))
        x, y = gap, gap + i * (H + lab)
        full.paste(im, (x, y))
        ImageDraw.Draw(full).text((x, y + H + 4), os.path.basename(p)[:-4], font=f18, fill=(60, 60, 60))
    full_path = prefix + "-full.png"
    full.save(full_path)

    print(sheet_path)
    print(full_path)


if __name__ == "__main__":
    main()
