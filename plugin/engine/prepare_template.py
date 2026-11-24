# -*- coding: utf-8 -*-
"""先把第二批的"母版/版式"补齐，再出稿。

问题：第二批 18 套模板的版式是空的（只有 5 个空占位符），设计全画在幻灯片上，
      且幻灯片没有任何占位符。直接 build 必然白页。

处理：把选定源幻灯片上的**装饰形状**（图片、色块、线条、小号编号文字）复制进版式，
      幻灯片背景也一并搬过去；同时用被删掉的正文文字区反推内容区，写回版式里的
      标题占位符几何，这样 --auto 就能量到正确的 CX/CW。

产出：<stem>-prepared.pptx（与原件同目录的 prepared 子目录），供 build --auto 使用。
"""
import argparse
import copy
import io
import json
import os

from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Emu

BLIP_TAGS = (".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip",)


def shape_text(sh):
    try:
        return sh.text_frame.text.strip() if sh.has_text_frame else ""
    except Exception:
        return ""


def is_decoration(sh):
    """文字长度 ≤4 的当装饰（编号 01、栏目名），其余当正文。"""
    t = shape_text(sh)
    if not t:
        return True
    return len(t) <= 4 and not any("\u4e00" <= c <= "\u9fff" for c in t)



_FRAME_TAGS = ("}graphicFrame", "}oleObj", "}media", "}video", "}audio")


def is_copyable(sh):
    """只搬 autoshape / picture / connector / 纯形状组合。"""
    xml = sh._element.xml[:4000]
    if any(t in xml for t in _FRAME_TAGS):
        return False
    if any(t in xml for t in (":diagram", "dgm:", "a:graphicData")):
        return False
    return True


NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"



def cleanup_dangling(el):
    """删掉因关系重绑失败而变成非法的空引用。"""
    def _drop(node, parent_tag):
        parent = node.getparent()
        if parent is not None and parent.tag.split("}")[-1] == parent_tag:
            gp = parent.getparent()
            if gp is not None:
                gp.remove(parent)

    for node in list(el.iter()):
        tag = node.tag.split("}")[-1]
        if tag == "blip":
            if not node.get("r:embed") and not any(k.endswith("}embed") for k in node.attrib):
                _drop(node, "blipFill")
        elif tag == "tags":
            if not any(k.endswith("}id") for k in node.attrib):
                _drop(node, "custDataLst")

def copy_shape(src_slide, dst_part, spTree, sh):
    el = copy.deepcopy(sh._element)
    for node in list(el.iter()):
        for attr in list(node.attrib.keys()):
            if not attr.startswith(NS_R):
                continue
            rid = node.get(attr)
            new_rid = None
            try:
                rel = src_slide.part.rels[rid]
                if rel.is_external:
                    new_rid = None
                else:
                    new_rid = dst_part.relate_to(rel.target_part, rel.reltype)
            except Exception:
                new_rid = None
            if new_rid:
                node.set(attr, new_rid)
            else:
                node.attrib.pop(attr, None)
    cleanup_dangling(el)
    spTree.append(el)
    used = set()
    for node in spTree.iter():
        if node.tag.endswith("}cNvPr"):
            try:
                used.add(int(node.get("id")))
            except Exception:
                pass
    nid = max(used) + 1 if used else 1
    for node in el.iter():
        if node.tag.endswith("}cNvPr"):
            node.set("id", str(nid))
            nid += 1


def slide_bg_xml(slide, dst_part=None):
    """复制幻灯片背景；带图片填充时把 rId 重绑到目标 part，失败就返回 None。"""
    bg = slide._element.find("{http://schemas.openxmlformats.org/presentationml/2006/main}bg")
    if bg is None:
        return None
    bg = copy.deepcopy(bg)
    for node in bg.iter():
        rid = node.get(NS_R + "embed") or node.get(NS_R + "id") or node.get(NS_R + "link")
        if not rid:
            continue
        if dst_part is None:
            return None
        try:
            part = slide.part.related_part(rid)
            new_rid = dst_part.relate_to(part, RT.IMAGE)
            for attr in ("embed", "id", "link"):
                if node.get(NS_R + attr):
                    node.set(NS_R + attr, new_rid)
        except Exception:
            return None
    return bg


def text_zones(slide):
    """返回 [(x, y, w, h, text)] —— 用于反推标题区与内容区。"""
    out = []
    for sh in slide.shapes:
        t = shape_text(sh)
        if not t or is_decoration(sh):
            continue
        out.append((sh.left / 914400, sh.top / 914400, sh.width / 914400, sh.height / 914400, t))
    return out


def pick_sources(prs):
    """封面源 = 第 0 页；正文源 = 装饰最丰富且留得住文字的那页。"""
    sh = prs.slide_height / 914400
    cover_i = 0
    body_i = 1 if len(prs.slides) > 1 else 0
    first_ok = None
    best, best_score = None, -1.0
    for i, s in enumerate(prs.slides):
        if i == 0:
            continue
        zs = text_zones(s)
        if not 2 <= len(zs) <= 6:
            continue
        if min(z[1] for z in zs) >= 0.35 * sh:
            continue
        if max(z[3] for z in zs) >= 0.18 * sh:
            continue
        if first_ok is None:
            first_ok = i
        deco = sum(1 for x in s.shapes if shape_text(x) == "" and is_copyable(x))
        pics = sum(1 for x in s.shapes if str(x.shape_type).startswith("PICTURE"))
        if not 6 <= deco <= 30:
            continue
        score = deco + 2 * pics
        if score > best_score:
            best, best_score = i, score
    if best is not None:
        body_i = best
    elif first_ok is not None:
        body_i = first_ok
    return cover_i, body_i


def build_prepared(src, out):
    prs = Presentation(src)
    cover_i, body_i = pick_sources(prs)
    sw, sh = prs.slide_width / 914400, prs.slide_height / 914400

    layouts = list(prs.slide_layouts)
    lay_cover = layouts[0]
    lay_body = layouts[1] if len(layouts) > 1 else layouts[0]

    meta = {"source": os.path.abspath(src), "canvas": [sw, sh],
            "cover_slide": cover_i, "body_slide": body_i}

    single = len(layouts) < 2
    meta["single_layout"] = single
    for lay, src_i in (() if single else ((lay_cover, cover_i), (lay_body, body_i))):
        slide = prs.slides[src_i]
        spTree = lay.shapes._spTree
        n = 0
        for shp in list(slide.shapes):
            if is_decoration(shp) and shape_text(shp) == "" and is_copyable(shp):
                copy_shape(slide, lay.part, spTree, shp)
                n += 1
        bg = slide_bg_xml(slide, lay.part)
        if bg is not None:
            csld = lay._element.find("{http://schemas.openxmlformats.org/presentationml/2006/main}cSld")
            if csld is not None:
                old_bg = csld.find("{http://schemas.openxmlformats.org/presentationml/2006/main}bg")
                if old_bg is not None:
                    csld.remove(old_bg)
                csld.insert(0, bg)
        meta["layout_%s_decoration" % lay.name] = n

        # 用文字区反推：标题区 = 最靠上的文字块；内容区 = 其余文字块的并集
        zs = sorted(text_zones(slide), key=lambda z: z[1])
        if zs:
            tz = zs[0]
            for ph in lay.placeholders:
                if str(ph.placeholder_format.type).startswith("TITLE") or \
                        str(ph.placeholder_format.type).startswith("CENTER_TITLE"):
                    ph.left, ph.top = Emu(int(tz[0] * 914400)), Emu(int(tz[1] * 914400))
                    ph.width, ph.height = Emu(int(tz[2] * 914400)), Emu(int(tz[3] * 914400))
                    meta["title_zone"] = list(tz[:4])
                    break
        rest = zs[1:]
        if rest:
            x0 = min(z[0] for z in rest)
            y0 = min(z[1] for z in rest)
            x1 = max(z[0] + z[2] for z in rest)
            y1 = max(z[1] + z[3] for z in rest)
            meta["content_zone"] = [round(x0, 2), round(y0, 2), round(x1 - x0, 2), round(y1 - y0, 2)]

    if single:
        zs = sorted(text_zones(prs.slides[body_i]), key=lambda z: z[1])
        if zs:
            meta["title_zone"] = list(zs[0][:4])
        rest = zs[1:]
        if rest:
            x0 = min(z[0] for z in rest)
            y0 = min(z[1] for z in rest)
            x1 = max(z[0] + z[2] for z in rest)
            y1 = max(z[1] + z[3] for z in rest)
            meta["content_zone"] = [round(x0, 2), round(y0, 2), round(x1 - x0, 2), round(y1 - y0, 2)]

    os.makedirs(os.path.dirname(out), exist_ok=True)
    prs.save(out)
    return meta


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("--out", default="")
    args = ap.parse_args()
    src = args.input
    base = os.path.splitext(os.path.basename(src))[0]
    out = args.out or os.path.join(os.path.dirname(os.path.abspath(src)), "prepared", base + "-prepared.pptx")
    meta = build_prepared(src, out)
    side = os.path.splitext(out)[0] + ".json"
    io.open(side, "w", encoding="utf-8").write(json.dumps(meta, ensure_ascii=False, indent=2))
    print(json.dumps({"ok": True, "prepared": out, "meta": meta}, ensure_ascii=False))


if __name__ == "__main__":
    main()
