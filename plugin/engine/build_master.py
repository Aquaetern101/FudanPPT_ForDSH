# -*- coding: utf-8 -*-
"""母版模式出稿：只认版式与占位符，不搬设计、不估坐标。

用法：
  python build_master.py --template <pptx> --out <pptx> [--slides <json>]

产出一份"模板样张"（5 页）：封面 → 目录 → 章节页 → 正文页 → 封底。
页面内容取 --slides 里的题目/副题与四个小节名；缺省用内置示例。

排版来源（每个模板各自量）：
- 标题盒、正文盒：从该模板自己的幻灯片里量（最宽的标题文本块 + 其下方文本块的并集）
- 标题字级、正文字级、字体：取该模板里对应文本 run 的字号/字体
- 设计：只用带装饰的版式，文字一律写进版式占位符；没有占位符时才用上面量到的盒子补文本框
"""
import argparse
import glob
import json
import os
import statistics
import sys

from lxml import etree
from pptx import Presentation
from pptx.util import Emu, Pt

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def set_font(run, name, size, bold=False):
    run.font.size = Pt(size)
    run.font.bold = bold
    if name:
        run.font.name = name
        rPr = run._r.get_or_add_rPr()
        for tag in ("ea", "cs"):
            el = rPr.find(A + tag)
            if el is None:
                el = etree.SubElement(rPr, A + tag)
            el.set("typeface", name)


def fit_size(lines, w_pt, h_pt, pt, floor=7):
    size = pt
    while size > floor:
        per_line = max(1, int(w_pt / (size * 0.92)))
        rows = sum(max(1, -(-len(s) // per_line)) for s in lines)
        if rows * size * 1.30 <= h_pt:
            return size
        size -= 1
    return floor


def put(shape, lines, pt, font, bold=False, floor=7):
    if isinstance(lines, str):
        lines = [lines]
    tf = shape.text_frame
    tf.word_wrap = True
    size = fit_size(lines, shape.width / 12700.0, shape.height / 12700.0, pt, floor)
    for p in list(tf.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    for i, s in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        run = para.add_run()
        run.text = s
        set_font(run, font, size, bold)
    return size


def textbox(slide, box, lines, pt, font, bold=False):
    shp = slide.shapes.add_textbox(*[Emu(int(v * 914400)) for v in box])
    put(shp, lines, pt, font, bold)
    return shp


def learn(prs):
    SW, SH = prs.slide_width, prs.slide_height
    cand_t, cand_b = [], []
    for s in prs.slides:
        for shp in s.shapes:
            if not shp.has_text_frame or len(shp.text_frame.text.strip()) < 2:
                continue
            sz = font = None
            for para in shp.text_frame.paragraphs:
                for r in para.runs:
                    if r.font.size and sz is None:
                        sz = r.font.size.pt
                    if r.font.name and font is None:
                        font = r.font.name
                if sz and font:
                    break
            item = (shp.left, shp.top, shp.width, shp.height, sz, shp.text_frame.text.strip(), font)
            if shp.top < 0.36 * SH and shp.width > 0.42 * SW:
                cand_t.append(item)
            elif shp.width > 0.12 * SW:
                cand_b.append(item)

    def fmt(v):
        return round(v / 914400, 2)

    t = max((i for i in cand_t if i[4]), key=lambda i: i[2], default=None)
    title_box = [fmt(t[0]), fmt(t[1]), fmt(t[2]), fmt(t[3])] if t else \
        [fmt(0.06 * SW), fmt(0.08 * SH), fmt(0.88 * SW), fmt(0.14 * SH)]
    title_pt = t[4] if t else 40
    title_font = t[6] if (t and t[6] and not t[6].startswith("+")) else None

    below = [i for i in cand_b if i[1] > (title_box[1] + title_box[3]) * 914400 * 0.8]
    if below:
        x0 = min(i[0] for i in below)
        y0 = min(i[1] for i in below)
        x1 = max(i[0] + i[2] for i in below)
        y1 = max(i[1] + i[3] for i in below)
        body_box = [fmt(x0), fmt(y0), fmt(x1 - x0), fmt(y1 - y0)]
        sizes = [i[4] for i in below if i[4] and 10 <= i[4] <= 34]
        body_pt = int(statistics.median(sizes)) if sizes else 18
        # 正文字体取"正文区出现最多的字体"，避免把只有英文的字体（Abhaya Libre 等）当中文正文字体
        freq = {}
        for i in below:
            if i[6] and not i[6].startswith("+"):
                freq[i[6]] = freq.get(i[6], 0) + 1
        body_font = max(freq.items(), key=lambda kv: kv[1])[0] if freq else None
    else:
        body_box = [title_box[0], title_box[1] + title_box[3] + 0.3,
                    title_box[2], max(0.6, fmt(SH) - (title_box[1] + title_box[3]) - 0.9)]
        body_pt, body_font = 18, None

    fonts = {}
    for s in prs.slides:
        for shp in s.shapes:
            if not shp.has_text_frame:
                continue
            for para in shp.text_frame.paragraphs:
                for r in para.runs:
                    if r.font.name:
                        fonts[r.font.name] = fonts.get(r.font.name, 0) + 1
    top_fonts = [f for f, _ in sorted(fonts.items(), key=lambda kv: -kv[1]) if not f.startswith("+")]
    return {
        "title_box": title_box, "title_pt": title_pt, "title_font": title_font or (top_fonts[0] if top_fonts else None),
        "body_box": body_box, "body_pt": body_pt, "body_font": body_font or (top_fonts[0] if top_fonts else None),
        "fonts": top_fonts[:4],
    }


def pick_layouts(prs):
    lays = list(prs.slide_layouts)
    deco = [sum(1 for sh in l.shapes if not sh.is_placeholder) for l in lays]

    def find(keys):
        best, score = None, -1
        for i, l in enumerate(lays):
            if deco[i] == 0:
                continue
            if keys and not any(k.lower() in (l.name or "").lower() for k in keys):
                continue
            if deco[i] > score:
                best, score = i, deco[i]
        return best

    cover = find(["封面", "title slide", "标题幻灯片"])
    if cover is None:
        cover = max(range(len(lays)), key=lambda i: deco[i])
    toc = find(["目录", "content"])
    sect = find(["节标题", "小标题页", "section"])
    content = find(["标题和内容", "自定义版式", "defause", "内容", "picture"])
    if content is None:
        content = sect if sect is not None else None
    ending = find(["ending", "封底", "致谢"])
    return {
        "cover": cover,
        "toc": toc if toc is not None else content if content is not None else cover,
        "section": sect if sect is not None else cover,
        "content": content if content is not None else cover,
        "ending": ending if ending is not None else cover,
    }, deco


def ph_by_type(slide, types):
    return [shp for shp in slide.shapes
            if shp.is_placeholder and str(shp.placeholder_format.type).split()[0] in types]


def drop_empty(slide):
    for shp in list(slide.shapes):
        if not shp.is_placeholder:
            continue
        if str(shp.shape_type).startswith("PICTURE"):
            continue
        if shp._element.find(".//" + A + "blip") is not None:
            continue
        if shp.has_text_frame and shp.text_frame.text.strip():
            continue
        shp._element.getparent().remove(shp._element)


def clear_slides(prs):
    for sld in list(prs.slides._sldIdLst):
        rid = sld.get(RID)
        if rid:
            prs.part.drop_rel(rid)
        prs.slides._sldIdLst.remove(sld)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--slides")
    args = ap.parse_args()

    title = "深度突变扫描引导的宏基因组基序挖掘"
    subtitle = "Meta-SIFT 在 T7 受体结合蛋白上的应用"
    if args.slides and os.path.exists(args.slides):
        try:
            doc = json.load(open(args.slides, encoding="utf-8"))
            deck = doc.get("deck", {})
            title = deck.get("title") or title
            subtitle = deck.get("subtitle") or subtitle
        except Exception:
            pass

    prs = Presentation(args.template)
    info = learn(prs)
    picks, _ = pick_layouts(prs)
    lays = list(prs.slide_layouts)
    clear_slides(prs)

    tf, bf = info["title_font"], info["body_font"]

    def add(idx):
        return prs.slides.add_slide(lays[idx])

    def head(slide, text, ratio=1.0, bold=True):
        ttl = ph_by_type(slide, ("CENTER_TITLE", "TITLE"))
        pt = info["title_pt"] * ratio
        if ttl:
            put(ttl[0], text, pt, tf, bold)
        else:
            textbox(slide, info["title_box"], text, pt, tf, bold)

    def body(slide, text):
        got = ph_by_type(slide, ("BODY", "OBJECT", "SUBTITLE"))
        if got:
            put(got[0], text, info["body_pt"], bf)
        else:
            textbox(slide, info["body_box"], text, info["body_pt"], bf)

    s = add(picks["cover"])
    head(s, title)
    body(s, subtitle)
    drop_empty(s)

    s = add(picks["toc"])
    ents = ["空间有多大", "旧索引为何失效", "Meta-SIFT 怎么建库", "筛出来的是什么"]
    bodies = ph_by_type(s, ("BODY", "OBJECT"))
    if len(bodies) >= 4:
        for shp, t in zip(bodies[:4], ents):
            put(shp, t, info["body_pt"] + 6, bf)
    else:
        top = info["body_box"][1]
        step = max(0.55, (info["title_box"][1] + info["body_box"][3] - top) / 4)
        for i, t in enumerate(ents):
            textbox(s, [info["title_box"][0], top + i * step, info["title_box"][2] * 0.92, step * 0.8],
                    "%02d  %s" % (i + 1, t), info["body_pt"] + 4, bf)
    drop_empty(s)

    s = add(picks["section"])
    head(s, "01  空间有多大", 0.8)
    body(s, "6 个氨基酸的组合已达 6.4×10⁷，10 个位点是 10¹³ 量级，合计超过 10¹⁴。")
    drop_empty(s)

    s = add(picks["content"])
    head(s, "以实验功能分数作搜索种子", 0.62)
    body(s, "seed score 高的位点落在外环与朝外 β 片层，与已知侵染决定区吻合。\n"
            "给序列空间按功能排序，而不是按同源程度排序。\n"
            "结构蛋白命中显著多于非结构蛋白。")
    drop_empty(s)

    s = add(picks["ending"])
    head(s, "敬请批评指正")
    drop_empty(s)

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    prs.save(args.out)
    print("saved %s slides %d" % (args.out, len(prs.slides)))
    print("master: title %s pt/%s, body %s pt/%s, layouts %s" % (
        info["title_pt"], tf, info["body_pt"], bf, picks))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
