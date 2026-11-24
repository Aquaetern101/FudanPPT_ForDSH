# -*- coding: utf-8 -*-
"""从 PDF 原图裁切配图（PyMuPDF）。

配图优先级：PDF 原图裁切 > 官方补充材料 > 自绘示意图（须标注「示意」）。
本脚本负责第一档，支持：
  * --probe            列出页面尺寸、内嵌位图与矢量图形聚类的 bbox，供判断怎么裁
  * --box x0,y0,x1,y1  按页面比例裁一块（0–1）
  * --grid 2x2         把 --box 区域再切成 R×C 个子图（模型判断「需不需要分图」）
  * --trim             自动去掉四周纯白边（裁完再按墨迹 bbox 重渲一次，保清晰度）
  * 输出 PNG + figures.json 清单（页号、点位框、像素尺寸、宽高比、dpi）

用法示例：
  python extract_figures.py --pdf paper.pdf --probe --pages 3-5
  python extract_figures.py --pdf paper.pdf --page 4 --box 0.05,0.10,0.95,0.85 --trim \
      --outdir figs --prefix fig1 --dpi 300
  python extract_figures.py --pdf paper.pdf --page 4 --box 0.05,0.10,0.95,0.85 \
      --grid 2x2 --trim --outdir figs --prefix fig1 --dpi 300
"""
import argparse
import io
import json
import os
import re
import sys

try:                    # PyMuPDF 1.24+ 推荐名
    import pymupdf as fitz
except ImportError:     # 老版本
    import fitz


def parse_pages(spec):
    out = []
    for part in str(spec).split(","):
        part = part.strip()
        if not part:
            continue
        m = re.match(r"^(\d+)\s*-\s*(\d+)$", part)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            out.extend(range(a, b + 1))
        elif part.isdigit():
            out.append(int(part))
    return sorted(set(out))


def frac_box(page_rect, box):
    x0, y0, x1, y1 = box
    w, h = page_rect.width, page_rect.height
    return fitz.Rect(page_rect.x0 + x0 * w, page_rect.y0 + y0 * h,
                     page_rect.x0 + x1 * w, page_rect.y0 + y1 * h)


def ink_bbox(pix, threshold=250, pad=2):
    """返回非白墨迹的像素 bbox（含 pad），找不到返回 None。"""
    import numpy as np
    arr = np.frombuffer(pix.samples, dtype=np.uint8)
    arr = arr.reshape(pix.height, pix.width, pix.n)
    rgb = arr[:, :, :3]
    mask = (rgb < threshold).any(axis=2)
    ys, xs = np.where(mask)
    if ys.size == 0:
        return None
    y0 = max(0, int(ys.min()) - pad)
    y1 = min(pix.height, int(ys.max()) + 1 + pad)
    x0 = max(0, int(xs.min()) - pad)
    x1 = min(pix.width, int(xs.max()) + 1 + pad)
    return x0, y0, x1, y1


def auto_figure_rect(page, min_px=200, min_frac=0.01, pad=4.0):
    """用大位图 + 大矢量簇的并集框出图区；找不到返回 None。"""
    pr = page.rect
    rects = []
    for img in page.get_images(full=True):
        if img[2] >= min_px and img[3] >= min_px:
            for r in page.get_image_rects(img[0]):
                if r.width > 40 and r.height > 40:
                    rects.append(r)
    for c in cluster_boxes([d["rect"] for d in page.get_drawings()]):
        r = fitz.Rect(c)
        if (r.width * r.height) >= min_frac * pr.width * pr.height:
            rects.append(r)
    if not rects:
        return None
    x0 = min(r.x0 for r in rects) - pad
    y0 = min(r.y0 for r in rects) - pad
    x1 = max(r.x1 for r in rects) + pad
    y1 = max(r.y1 for r in rects) + pad
    return fitz.Rect(max(pr.x0, x0), max(pr.y0, y0), min(pr.x1, x1), min(pr.y1, y1))


def render(doc, page_no, rect, dpi):
    page = doc[page_no - 1]
    pix = page.get_pixmap(clip=rect, dpi=dpi, alpha=False)
    return pix


def trim_rect(doc, page_no, rect, dpi):
    """渲染一次找墨迹边界，换算回点位矩形。"""
    pix = render(doc, page_no, rect, dpi)
    bb = ink_bbox(pix)
    if bb is None:
        return rect, pix
    sx = rect.width / pix.width
    sy = rect.height / pix.height
    x0, y0, x1, y1 = bb
    return fitz.Rect(rect.x0 + x0 * sx, rect.y0 + y0 * sy,
                     rect.x0 + x1 * sx, rect.y0 + y1 * sy), pix


def cluster_boxes(rects, gap=6.0):
    """把一堆矩形按邻近关系合并成簇（贪心，够用即可）。"""
    boxes = [[r.x0, r.y0, r.x1, r.y1] for r in rects if r.width > 1 and r.height > 1]
    changed = True
    while changed:
        changed = False
        out = []
        while boxes:
            a = boxes.pop()
            i = 0
            while i < len(boxes):
                b = boxes[i]
                if (a[0] - gap <= b[2] and b[0] - gap <= a[2]
                        and a[1] - gap <= b[3] and b[1] - gap <= a[3]):
                    a = [min(a[0], b[0]), min(a[1], b[1]),
                         max(a[2], b[2]), max(a[3], b[3])]
                    boxes.pop(i)
                    changed = True
                else:
                    i += 1
            out.append(a)
        boxes = out
    return boxes


def probe(doc, pages):
    report = []
    for p in pages:
        page = doc[p - 1]
        pr = page.rect
        images = []
        seen = set()
        for img in page.get_images(full=True):
            xref = img[0]
            if min(img[2], img[3]) < 100:      # 3x2 这类是被拆碎的符号，直接忽略
                continue
            try:
                rects = page.get_image_rects(xref)
            except Exception:
                rects = []
            for r in rects:
                if r.width < 20 or r.height < 20:
                    continue
                key = (round(r.x0), round(r.y0), round(r.x1), round(r.y1))
                if key in seen:
                    continue
                seen.add(key)
                images.append({
                    "bbox_pt": [round(v, 1) for v in (r.x0, r.y0, r.x1, r.y1)],
                    "frac": [round((r.x0 - pr.x0) / pr.width, 3), round((r.y0 - pr.y0) / pr.height, 3),
                             round((r.x1 - pr.x0) / pr.width, 3), round((r.y1 - pr.y0) / pr.height, 3)],
                    "px": [img[2], img[3]],
                })
        drawings = page.get_drawings()
        clusters = cluster_boxes([d["rect"] for d in drawings])
        clusters = sorted([c for c in clusters if (c[2] - c[0]) > 20 and (c[3] - c[1]) > 20],
                          key=lambda c: (c[1], c[0]))
        report.append({
            "page": p,
            "size_pt": [round(pr.width, 1), round(pr.height, 1)],
            "images": images[:12],
            "vector_groups": [{
                "bbox_pt": [round(v, 1) for v in c],
                "frac": [round((c[0] - pr.x0) / pr.width, 3), round((c[1] - pr.y0) / pr.height, 3),
                         round((c[2] - pr.x0) / pr.width, 3), round((c[3] - pr.y0) / pr.height, 3)],
            } for c in clusters][:8],
            "text_chars": len(page.get_text("text")),
        })
    return report


def main():
    ap = argparse.ArgumentParser(description="PDF 原图裁切")
    ap.add_argument("--pdf", required=True)
    ap.add_argument("--page", type=int)
    ap.add_argument("--pages", default="")
    ap.add_argument("--box", default="", help="页面比例框 x0,y0,x1,y1（0–1）")
    ap.add_argument("--grid", default="", help="RxC，把 box 切成 R 行 C 列")
    ap.add_argument("--trim", action="store_true")
    ap.add_argument("--auto", action="store_true",
                    help="用页面上的大位图与矢量图形簇自动框出图区（避开正文栏）")
    ap.add_argument("--dpi", type=int, default=300)
    ap.add_argument("--outdir", default="")
    ap.add_argument("--prefix", default="fig")
    ap.add_argument("--probe", action="store_true")
    ap.add_argument("--manifest", default="")
    args = ap.parse_args()

    doc = fitz.open(args.pdf)
    if args.probe:
        pages = parse_pages(args.pages) or ([args.page] if args.page else [1])
        print(json.dumps({"pdf": args.pdf, "page_count": doc.page_count,
                          "pages": probe(doc, pages)}, ensure_ascii=False, indent=2))
        return

    if not args.page:
        raise SystemExit("裁切需要 --page")
    outdir = args.outdir or os.path.join(os.path.dirname(os.path.abspath(args.pdf)), "figs")
    os.makedirs(outdir, exist_ok=True)

    page = doc[args.page - 1]
    pr = page.rect
    base = frac_box(pr, [float(v) for v in args.box.split(",")]) if args.box else pr
    if args.auto:
        auto = auto_figure_rect(page)
        if auto is None:
            raise SystemExit("--auto 没找到图区，请改用 --box 手工框")
        base = (base & auto) if args.box else auto

    units = []
    if args.grid:
        m = re.match(r"^(\d+)\s*[xX×]\s*(\d+)$", args.grid.strip())
        if not m:
            raise SystemExit("--grid 形如 2x2")
        rows, cols = int(m.group(1)), int(m.group(2))
        w = base.width / cols
        h = base.height / rows
        for r in range(rows):
            for c in range(cols):
                rect = fitz.Rect(base.x0 + c * w, base.y0 + r * h,
                                 base.x0 + (c + 1) * w, base.y0 + (r + 1) * h)
                units.append(("r%dc%d" % (r + 1, c + 1), rect))
    else:
        units.append(("", base))

    entries = []
    for tag, rect in units:
        if args.trim:
            rect, _ = trim_rect(doc, args.page, rect, args.dpi)
        pix = render(doc, args.page, rect, args.dpi)
        name = "_".join([x for x in (args.prefix, ("p%d" % args.page), tag) if x]) + ".png"
        path = os.path.join(outdir, name)
        pix.save(path)
        entries.append({
            "file": path,
            "page": args.page,
            "tag": tag,
            "bbox_pt": [round(v, 1) for v in (rect.x0, rect.y0, rect.x1, rect.y1)],
            "frac": [round((rect.x0 - pr.x0) / pr.width, 3), round((rect.y0 - pr.y0) / pr.height, 3),
                     round((rect.x1 - pr.x0) / pr.width, 3), round((rect.y1 - pr.y0) / pr.height, 3)],
            "px": [pix.width, pix.height],
            "aspect": round(pix.width / pix.height, 3),
            "dpi": args.dpi,
            "kb": round(os.path.getsize(path) / 1024, 1),
            "source_pdf": os.path.abspath(args.pdf),
        })

    mpath = args.manifest or os.path.join(outdir, "figures.json")
    manifest = {"pdf": os.path.abspath(args.pdf), "entries": []}
    if os.path.exists(mpath):
        try:
            with io.open(mpath, encoding="utf-8") as f:
                manifest = json.load(f)
            manifest.setdefault("entries", [])
        except Exception:
            pass
    manifest["entries"].extend(entries)
    with io.open(mpath, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    print(json.dumps({"ok": True, "manifest": mpath, "entries": entries},
                     ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
