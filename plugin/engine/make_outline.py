# -*- coding: utf-8 -*-
"""C2：任意文档 → 大纲 JSON（骨架版）。

输入：pdf / docx / md / txt
输出：符合引擎 kind 的大纲 JSON；再交给 build_deck.py 出稿。

叙事骨架（按你的 (目的-手段-小论点) 逻辑）：封面 → 问题页 → 方法页（四步缎带）
→ 结果页（1~2 页，带数字的句子）→ 局限页 → 参考文献 → 谢谢聆听

说明：这是"骨架"，purpose 与论据还需要人或模型审一遍——它保证的是
"任何文档都能一路走到 pptx"，而不是"自动写得好"。
"""
import argparse
import io
import json
import os
import re
import sys

HEAD_PAT = [
    re.compile(r"^\s*(摘\s*要|Abstract)\s*$", re.I),
    re.compile(r"^\s*(引\s*言|前\s*言|绪\s*论|Introduction|Background)\s*$", re.I),
    re.compile(r"^\s*(方\s*法|材料与方法|Methods?|Materials and Methods)\s*$", re.I),
    re.compile(r"^\s*(结\s*果|Results?)\s*$", re.I),
    re.compile(r"^\s*(讨\s*论|结\s*论|Discussion|Conclusions?)\s*$", re.I),
    re.compile(r"^\s*(参考文献|References?|Bibliography)\s*$", re.I),
]
NUM = re.compile(r"\d")


def read_pdf(path):
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    doc = fitz.open(path)
    pages = [doc[i].get_text("text") for i in range(doc.page_count)]
    return pages


def read_docx(path):
    from docx import Document
    d = Document(path)
    paras = [p.text for p in d.paragraphs]
    out, cur = [], []
    for t in paras:
        if t.strip():
            cur.append(t.strip())
        elif cur:
            out.append("\n".join(cur))
            cur = []
    if cur:
        out.append("\n".join(cur))
    return out


def read_text(path):
    return io.open(path, encoding="utf-8", errors="ignore").read().split("\n\n")


def blocks_of(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        raw = read_pdf(path)
        text = "\n".join(raw)
        # 按空行/标题切段
        return [b.strip() for b in re.split(r"\n\s*\n", text) if b.strip()]
    if ext == ".docx":
        return read_docx(path)
    return [b.strip() for b in read_text(path) if b.strip()]


def norm(t):
    return re.sub(r"\s+", " ", t).strip()


def pick_sentences(text, maxn=3, prefer_num=False):
    sents = [s for s in re.split(r"(?<=[。！？.!?])\s*", norm(text)) if len(s) > 8]
    if prefer_num:
        with_num = [s for s in sents if NUM.search(s)]
        sents = with_num or sents
    return sents[:maxn]


def split_sections(blocks):
    """按标题词把正文切成 {名称: [块]}；没标题就整篇当一节。"""
    sections = {"正文": []}
    cur = "正文"
    for b in blocks:
        head = norm(b).strip("：: ")
        matched = None
        for pat in HEAD_PAT:
            if pat.match(head) and len(head) <= 20:
                matched = pat.pattern
                break
        if matched:
            cur = head
            sections.setdefault(cur, [])
        else:
            sections.setdefault(cur, []).append(b)
    return sections


def find(sections, *keys):
    for name, body in sections.items():
        for k in keys:
            if k.lower() in name.lower():
                return body
    return []


def title_of_doc(path, sections):
    base = os.path.splitext(os.path.basename(path))[0]
    for name in sections:
        if name not in ("正文",) and len(name) < 30:
            return name
    return base



MEETING_KW = ("会议记录", "会议纪要", "纪要", "议程", "主持人", "签到", "表决", "选民小组")
PAPER_KW = ("摘要", "Abstract", "关键词", "Methods", "参考文献", "Results", "Discussion")


def doc_type(text):
    if sum(1 for k in MEETING_KW if k in text) >= 3:
        return "meeting"
    if sum(1 for k in PAPER_KW if k in text) >= 3:
        return "paper"
    return "generic"


def _lines(blocks_text):
    out = []
    for line in blocks_text.split("\n"):
        line = norm(line)
        if len(line) >= 6:
            out.append(line)
    return out


def _hits(lines, kws, maxn=3, limit=42):
    got = []
    for ln in lines:
        if any(k in ln for k in kws):
            got.append(ln[:limit])
        if len(got) >= maxn:
            break
    return got


def build_meeting(path, blocks, title):
    text = "\n".join(blocks)
    lines = _lines(text)
    slides = []
    n = 1
    short = title
    for pre in ("复旦大学", "杨浦004选区", "杨浦区", "第"):
        if len(short) > 24 and pre in short and pre != "第":
            short = short[short.index(pre) + len(pre):]
    if len(short) > 24:
        short = short[-24:]
    slides.append({"n": n, "kind": "cover", "layout": "cover", "purpose": short,
                   "sub": "会议记录", "notes": "完整名称：%s。" % title})
    n += 1

    basic = _hits(lines, ("时间", "地点", "主持", "记录人", "应到", "实到", "出席"), 3, 46)
    slides.append({"n": n, "kind": "bullets", "layout": "bullets-3",
                   "purpose": "会议基本信息",
                   "bullets": basic or ["（记录未写时间地点）"],
                   "conclusion": "会议要素齐备，程序有效。", "notes": "骨架自动生成"})
    n += 1

    agenda = _hits(lines, ("议程", "议题", "会议开始", "宣布"), 3, 46)
    slides.append({"n": n, "kind": "bullets", "layout": "bullets-3",
                   "purpose": "会议议程",
                   "bullets": agenda or ["（记录未列议程）"],
                   "conclusion": "议程经主持人说明后依次进行。", "notes": "骨架自动生成"})
    n += 1

    report = _hits(lines, ("通报", "说明", "介绍", "汇报", "产生情况"), 3, 46)
    slides.append({"n": n, "kind": "bullets", "layout": "bullets-3",
                   "purpose": "主要通报事项",
                   "bullets": report or ["（无通报事项）"],
                   "conclusion": "情况通报构成讨论的基础。", "notes": "骨架自动生成；数字需回原文核对"})
    n += 1

    vote = _hits(lines, ("表决", "通过", "同意", "反对", "弃权", "选举", "推选"), 3, 46)
    slides.append({"n": n, "kind": "bullets", "layout": "bullets-3",
                   "purpose": "表决与结果",
                   "bullets": vote or ["（记录未载表决）"],
                   "conclusion": "表决结果按记录载明。", "notes": "骨架自动生成"})
    n += 1

    nxt = _hits(lines, ("下一步", "要求", "安排", "将于", "后续"), 3, 46)
    slides.append({"n": n, "kind": "bullets", "layout": "bullets-3",
                   "purpose": "后续安排",
                   "bullets": nxt or ["（无后续安排记录）"],
                   "conclusion": "事项按分工推进。", "notes": "骨架自动生成"})
    n += 1

    slides.append({"n": n, "kind": "thanks", "layout": "thanks", "purpose": "谢谢聆听", "notes": ""})
    return {"deck": {"title": title, "source": "meeting", "cite_style": "auto"}, "slides": slides}


def build(path):
    blocks = blocks_of(path)
    sections = split_sections(blocks)
    title = title_of_doc(path, sections)
    kind = doc_type("\n".join(blocks))
    if kind == "meeting":
        d = build_meeting(path, blocks, title)
        d["deck"]["doc_type"] = "meeting"
        return d

    intro = find(sections, "引言", "前言", "绪论", "Introduction", "Background", "摘要", "Abstract")
    method = find(sections, "方法", "Method")
    result = find(sections, "结果", "Result")
    disc = find(sections, "讨论", "结论", "Discussion", "Conclusion")

    body_text = "\n".join(blocks)
    if not intro:
        intro = blocks[:2]
    if not method:
        method = blocks[2:4]
    if not result:
        result = [b for b in blocks if NUM.search(b)][:3] or blocks[4:6]
    if not disc:
        disc = blocks[-2:] if len(blocks) > 2 else blocks[:1]

    slides = []
    n = 1
    slides.append({"n": n, "kind": "cover", "layout": "cover", "purpose": title,
                   "sub": "自动大纲（骨架）", "notes": "由 make_outline.py 生成，标题/论据待润色。"})
    n += 1

    intro_pts = []
    for b in intro:
        intro_pts.extend(pick_sentences(b, 2))
        if len(intro_pts) >= 3:
            break
    slides.append({"n": n, "kind": "bullets", "layout": "bullets-3",
                   "purpose": "研究背景与问题",
                   "bullets": [norm(s)[:40] for s in intro_pts[:3]] or ["（原始文档未给出明确背景段）"],
                   "conclusion": "问题的边界由此确定。", "notes": "骨架自动生成"})
    n += 1

    steps = []
    for b in method:
        s = pick_sentences(b, 1)
        if s:
            steps.append(norm(s[0])[:24])
        if len(steps) >= 4:
            break
    while len(steps) < 4:
        steps.append("步骤 %d" % (len(steps) + 1))
    slides.append({"n": n, "kind": "process", "layout": "ribbon-4",
                   "purpose": "研究方法与实施步骤",
                   "blocks": steps[:4],
                   "conclusion": "四步构成可复现的流程。", "notes": "骨架自动生成"})
    n += 1

    res_pts = []
    for b in result:
        res_pts.extend(pick_sentences(b, 3, prefer_num=True))
        if len(res_pts) >= 3:
            break
    slides.append({"n": n, "kind": "bullets", "layout": "bullets-3",
                   "purpose": "主要结果",
                   "bullets": [norm(s)[:42] for s in res_pts[:3]] or ["（未从原文提取到带数据的结果句）"],
                   "conclusion": "结果支持上述问题设定。", "notes": "骨架自动生成；数值需回原文核对"})
    n += 1

    lim_pts = []
    for b in disc:
        lim_pts.extend(pick_sentences(b, 2))
        if len(lim_pts) >= 3:
            break
    slides.append({"n": n, "kind": "bullets", "layout": "bullets-3",
                   "purpose": "讨论与局限",
                   "bullets": [norm(s)[:42] for s in lim_pts[:3]] or ["（原文讨论段较短）"],
                   "conclusion": "结论与限制同时成立。", "notes": "骨架自动生成"})
    n += 1

    refs_raw = find(sections, "参考文献", "References", "Bibliography")
    ref_texts = []
    for b in refs_raw:
        for line in b.split("\n"):
            line = norm(line)
            if len(line) > 20:
                ref_texts.append(line)
    refs = []
    for t in ref_texts[:6]:
        lang = "zh" if re.search(r"[\u4e00-\u9fff]", t) else "en"
        refs.append({"lang": lang, "runs": [{"t": t}]})
    if not refs:
        refs = [{"lang": "zh", "runs": [{"t": "（原文未提供参考文献列表，需补）"}]}]
    slides.append({"n": n, "kind": "refs", "layout": "refs", "purpose": "参考文献",
                   "cite_style": "auto", "refs": refs, "notes": "骨架自动生成；中文文献建议补 GB/T 7714 字段"})
    n += 1

    slides.append({"n": n, "kind": "thanks", "layout": "thanks", "purpose": "谢谢聆听", "notes": ""})

    return {"deck": {"title": title, "source": os.path.abspath(path), "cite_style": "auto",
                     "doc_type": kind},
            "slides": slides}


def main():
    ap = argparse.ArgumentParser(description="任意文档 → 大纲 JSON（骨架）")
    ap.add_argument("input")
    ap.add_argument("--out", default="")
    args = ap.parse_args()
    if not os.path.exists(args.input):
        sys.exit("找不到输入：%s" % args.input)
    d = build(args.input)
    out = args.out or os.path.join(os.path.dirname(os.path.abspath(args.input)),
                                   "outline-%s.json" % os.path.splitext(os.path.basename(args.input))[0])
    with io.open(out, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=2)
    print(json.dumps({"ok": True, "outline": out, "slides": len(d["slides"])}, ensure_ascii=False))


if __name__ == "__main__":
    main()
