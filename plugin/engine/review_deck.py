# -*- coding: utf-8 -*-
"""出稿后的四道校验 —— **建议模式**。

按你的要求：成品正常交付，校验只返回「下一步修改意见」，不阻断、不自动改。
（避免"校验不过 → 越改越奇怪"的死循环。）

四道：
  1 来源：配图是否来自 PDF 裁切（MD5 比对 figures.json）
  2 几何：图/表与文字是否重叠；图文是否只按左右或上下排布
  3 排版：字号是否在档、标题字数、重复文本、支撑句是否超过两行
  4 对比：文字压在装饰/色块上时的对比度

用法：
  python review_deck.py <pptx>            # 建议模式，永远 exit 0
  python review_deck.py <pptx> --strict   # CI 用：有问题 exit 2
"""
import hashlib
import io
import json
import os
import sys

from pptx import Presentation
from pptx.util import Emu

FONT_SCALE = {11, 13, 16, 22, 28, 32, 34, 40, 48, 54}
PICTURE = 13
TABLE = 19

HERE = os.path.dirname(os.path.abspath(__file__))
MANIFEST = os.path.join(os.path.dirname(HERE), "02-build", "fig-crop", "figures.json")


def md5(b):
    return hashlib.md5(b).hexdigest()


def load_crop_index(path):
    if not os.path.exists(path):
        return {}
    with io.open(path, encoding="utf-8") as f:
        m = json.load(f)
    idx = {}
    for e in m.get("entries", []):
        p = e.get("file")
        if p and os.path.exists(p):
            with open(p, "rb") as fh:
                idx[md5(fh.read())] = e
    return idx


def rect(sh):
    return (sh.left, sh.top, sh.left + sh.width, sh.top + sh.height)


def overlap(a, b):
    w = min(a[2], b[2]) - max(a[0], b[0])
    h = min(a[3], b[3]) - max(a[1], b[1])
    if w <= 0 or h <= 0:
        return 0.0
    return (w * h) / 914400.0 / 914400.0


def luminance(rgb):
    def ch(v):
        v = v / 255.0
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    r, g, b = rgb
    return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b)


def contrast(c1, c2):
    l1, l2 = luminance(c1), luminance(c2)
    hi, lo = max(l1, l2), min(l1, l2)
    return (hi + 0.05) / (lo + 0.05)


def shape_rgb(sh):
    try:
        if sh.fill.type == 1:  # solid
            c = sh.fill.fore_color.rgb
            return (c[0], c[1], c[2])
    except Exception:
        pass
    return None


def text_rgb(sh):
    try:
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if r.text.strip() and r.font.color and r.font.color.type is not None:
                    c = r.font.color.rgb
                    return (c[0], c[1], c[2])
    except Exception:
        pass
    return None



def text_extent(sh):
    """按文字实际渲染尺寸估算 bbox（顶对齐），避免把空白的框高算成遮挡。"""
    x, y, w, h = rect(sh)
    try:
        txt = sh.text_frame.text
    except Exception:
        return (x, y, x + w, y + h)
    lines, size, spacing = 0, 22.0, 1.4
    for p in sh.text_frame.paragraphs:
        pt = None
        for r in p.runs:
            if r.font.size is not None:
                pt = float(r.font.size.pt)
                break
        pt = pt or 22.0
        size = max(size, pt) if lines else pt
        spacing = float(p.line_spacing or 1.4) if p.line_spacing else 1.4
        chars_per_line = max(1.0, (w / 914400.0) / (pt * 0.0185))
        lines += max(1, int(-(-len(p.text) // int(chars_per_line))))
    if lines == 0:
        return (x, y, x + w, y + h)
    est_h = min(h, int(lines * size * spacing / 72.0 * 914400))
    return (x, y, x + w, y + est_h)


def rect_ink(sh):
    if sh.has_text_frame and sh.text_frame.text.strip():
        return text_extent(sh)
    return rect(sh)



def image_mean_rgb(sh):
    """取图片平均色（用于估算背景图上的文字对比度）。"""
    try:
        import io as _io
        from PIL import Image
        import numpy as np
        im = Image.open(_io.BytesIO(sh.image.blob)).convert("RGB")
        im.thumbnail((64, 64))
        a = np.asarray(im, dtype=float).reshape(-1, 3)
        m = a.mean(axis=0)
        return (float(m[0]), float(m[1]), float(m[2]))
    except Exception:
        return None



def shape_fill_rgb(sh):
    """实心色 → 直接返回；图片填充 → 取图片平均色；其余 None。"""
    try:
        if sh.fill.type == 1:
            c = sh.fill.fore_color.rgb
            return (c[0], c[1], c[2])
    except Exception:
        pass
    # 图片填充
    try:
        from pptx.oxml.ns import qn
        blip = sh._element.find(".//" + qn("a:blip"))
        if blip is None:
            return None
        rid = blip.get(qn("r:embed"))
        if not rid:
            return None
        part = sh.part.related_part(rid)
        import io as _io
        from PIL import Image
        import numpy as np
        im = Image.open(_io.BytesIO(part.blob)).convert("RGB")
        im.thumbnail((64, 64))
        a = np.asarray(im, dtype=float).reshape(-1, 3)
        m = a.mean(axis=0)
        return (float(m[0]), float(m[1]), float(m[2]))
    except Exception:
        return None


def review(pptx):
    prs = Presentation(pptx)
    idx = load_crop_index(MANIFEST)
    items = []

    def add(page, gate, msg, fix):
        items.append({"page": page, "gate": gate, "msg": msg, "fix": fix})

    # 非 13.333 画布的稿子，字号被 K 缩放；比较档位表前先还原设计字号
    k = (prs.slide_width / 914400.0) / 13.333
    if abs(k - 1.0) < 0.02:
        k = 1.0

    seen_text = {}
    for i, s in enumerate(prs.slides, 1):
        pics = [sh for sh in s.shapes if sh.shape_type == PICTURE]
        tables = [sh for sh in s.shapes if getattr(sh, "has_table", False)]
        texts = [sh for sh in s.shapes if sh.has_text_frame and sh.text_frame.text.strip()]

        # 1 来源
        for sh in pics:
            h = md5(sh.image.blob)
            if idx and h not in idx:
                add(i, "来源", "配图不是 PDF 裁切（%dx%d）" % sh.image.size,
                    "用 dppt_figures 从原文裁，或确认它是官方补充材料")

        # 2 几何
        for a in pics + tables:
            for b in texts:
                ov = overlap(rect(a), text_extent(b))
                if ov > 0.01:
                    add(i, "几何", "图/表与文字重叠 %.2f in²" % ov,
                        "把文字移到图的上方或下方，或缩图")
        if pics and texts:
            for a in pics:
                for b in texts:
                    ra, rb = rect(a), text_extent(b)
                    h_ok = ra[2] <= rb[0] or rb[2] <= ra[0]
                    v_ok = ra[3] <= rb[1] or rb[3] <= ra[1]
                    if not h_ok and not v_ok and overlap(ra, rb) <= 0.01:
                        add(i, "几何", "图文既非左右也非上下", "改成左图右文，或上图下文")

        # 3 排版
        for sh in texts:
            t = sh.text_frame.text.strip()
            if t in seen_text:
                add(i, "排版", "与第 %d 页重复：「%s」" % (seen_text[t], t[:18]), "删掉重复文本")
            else:
                seen_text[t] = i
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.size is None:
                        continue
                    pt = int(round(r.font.size.pt / k))
                    if pt not in FONT_SCALE:
                        add(i, "排版", "字号 %dpt 不在档位表" % pt, "回落到 11/13/16/22/28/32/48/54")
            if sh.is_placeholder and sh.placeholder_format.type in (1, 3) and len(t) > 34:
                add(i, "排版", "标题 %d 字，超出 34 字上限" % len(t), "压成名词短语，或拆两页")
            # 支撑句行数估算：22pt / 22.5 字每英寸（实测经验值）
            if t and not sh.is_placeholder:
                est_lines = max(1, int(len(t) / 22.5 / max(sh.width / 914400.0, 1)) + 1)
                if est_lines > 2:
                    add(i, "排版", "文本约 %d 行（%d 字）" % (est_lines, len(t)),
                        "按 A2：论证超过两句就拆页")

        # 3.5 整幅背景图（版式里的图也算：封面/封底图通常放在 layout 上）
        sw_, sh_ = prs.slide_width, prs.slide_height
        bg_pics = list(pics) + [sh for sh in s.slide_layout.shapes if sh.shape_type == PICTURE]
        for pic in bg_pics:
            ra = rect(pic)
            cover_ratio = ((min(ra[2], sw_) - max(ra[0], 0)) * (min(ra[3], sh_) - max(ra[1], 0))) / float(sw_ * sh_)
            if cover_ratio < 0.70:
                continue
            bg = image_mean_rgb(pic)
            if not bg:
                continue
            for ts in texts:
                fg = text_rgb(ts)
                if not fg:
                    continue
                ratio = contrast(fg, tuple(int(v) for v in bg))
                if ratio < 3.0:
                    add(i, "对比", "文字压在整幅背景图上（背景均色 #%02X%02X%02X，对比 %.1f:1）"
                        % (int(bg[0]), int(bg[1]), int(bg[2]), ratio),
                        "换白/深色文字，或给文字加半透明底；需人工看预览确认")

        # 4 对比
        deco = [sh for sh in list(s.shapes) + list(s.slide_layout.shapes)
                if (not sh.has_text_frame or not sh.text_frame.text.strip()) and shape_fill_rgb(sh)]
        for ts in texts:
            fg = text_rgb(ts)
            if not fg:
                continue
            for d in deco:
                ov = overlap(text_extent(ts), rect(d))
                if ov <= 0.01:
                    continue
                bg = shape_fill_rgb(d)
                if not bg:
                    continue
                ratio = contrast(fg, bg)
                if ratio < 3.0:
                    add(i, "对比", "文字压「%s」仅 %.1f:1" % (d.name, ratio),
                        "文字收进白底，或改白色文字（≥4.5:1）")
    return items


def main():
    if len(sys.argv) < 2:
        sys.exit("用法：python review_deck.py <pptx> [--strict]")
    pptx = sys.argv[1]
    strict = "--strict" in sys.argv
    items = review(pptx)
    print("=" * 60)
    print("校验结果：%s" % os.path.basename(pptx))
    print("=" * 60)
    if not items:
        print("四道全过：来源 / 几何 / 排版 / 对比。")
        return
    print("成品照常交付；以下仅作下一步修改意见（%d 条）\n" % len(items))
    for k, it in enumerate(items, 1):
        print("%2d. [%s] p%02d  %s" % (k, it["gate"], it["page"], it["msg"]))
        print("      建议：%s" % it["fix"])
    if strict:
        sys.exit(2)


if __name__ == "__main__":
    main()
