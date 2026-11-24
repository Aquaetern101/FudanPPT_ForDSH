# -*- coding: utf-8 -*-
"""按「日冕-绿」模板生成完整 12 页（v3：重心居中 / 数值解析修正 / 单 panel 选图）。
输入：02-build/slides-sciadv.json
输出：03-out/sciadv-crown.pptx
"""
import io
import json
import math
import os
import re
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt
_BASE_INCHES = Inches
_BASE_PT = Pt
K = 1.0
BODY_DY = 0.0


def CANVAS(v):
    """画布派生长度（SW-1.0 这类），不参与 K 缩放。"""
    return Emu(int(v * 914400))


def Inches(v):  # noqa: N802
    return Emu(int(_BASE_INCHES(v) * K))


def BY(v):
    """正文纵向带：按模板内容区整体下移。"""
    return Emu(int(_BASE_INCHES(v + BODY_DY) * K))


def Pt(v):  # noqa: N802
    return Emu(int(_BASE_PT(v) * K))




HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.environ.get("DPPT_ROOT", r"D:\dsh-wx\fudan-ppt")
TEMPLATE = os.path.join(ROOT, r"00-template\复旦风景(日冕)-绿16：9.pptx")
SLIDES = os.environ.get("DPPT_SLIDES", os.path.join(HERE, "slides-sciadv.json"))
PANELS = os.path.join(ROOT, r"02-build\fig-sciadv\panels")
OUT = os.path.join(ROOT, r"03-out\sciadv-crown.pptx")

SERIES_ID = "R3"
CITE_ON_BAND = True
COVER = {}
WHITE_BG = False
TITLE_TOP = 0.08
THANKS_TOP = 3.02
THANKS_RULE = 4.42
THANKS_TEXT = "致谢"
ACCENT = RGBColor(0x77, 0x94, 0x3C)
ACCENT2 = RGBColor(0x4F, 0x63, 0x28)

PROFILES = {
    # 风景系列：7 个版式、坐标一致，仅配色不同 → 同一代码，只换 profile
    "crown": {
        "series": "R3",
        "template": os.path.join(ROOT, r"00-template\复旦风景(日冕)-绿16：9.pptx"),
        "accent": (0x77, 0x94, 0x3C),
        "out": os.path.join(ROOT, r"03-out\sciadv-crown.pptx"),
        "body_layout": 3, "cover_layout": 6,
    },
    "motto": {
        "series": "R50",
        "template": os.path.join(ROOT, r"00-template\复旦风景(校训墙)-红16：9.pptx"),
        "accent": (0xF9, 0x30, 0x25),
        "out": os.path.join(ROOT, r"03-out\sciadv-motto.pptx"),
        "body_layout": 3, "cover_layout": 6,
    },
    "jiangwan": {
        "series": "R2",
        "template": os.path.join(ROOT, r"00-template\复旦风景(江湾校区)-黑16：9.pptx"),
        "accent": (0x26, 0x26, 0x26),
        "out": os.path.join(ROOT, r"03-out\sciadv-jiangwan.pptx"),
        "body_layout": 3, "cover_layout": 6,
    },
    "xianghuitang": {
        "series": "R1",
        "template": os.path.join(ROOT, r"00-template\复旦风景(相辉堂)-蓝16：9.pptx"),
        "accent": (0x01, 0x71, 0xC5),
        "out": os.path.join(ROOT, r"03-out\sciadv-xianghuitang.pptx"),
        "body_layout": 3, "cover_layout": 6,
    },
    "laoxiaomen": {
        "series": "R200",
        "template": os.path.join(ROOT, r"00-template\复旦风景(老校门)-紫16：9.pptx"),
        "accent": (0x7E, 0x6C, 0xC3),
        "out": os.path.join(ROOT, r"03-out\sciadv-laoxiaomen.pptx"),
        "body_layout": 3, "cover_layout": 6,
    },
    # 模板01：6 版式，封面带图片占位符与默认 Office 配色图形；末尾自带封底版式
    "tpl01": {
        "series": "R1",
        "template": os.path.join(ROOT, r"00-template\复旦模板-01.pptx"),
        "accent": (0x00, 0xAD, 0xEF),
        "out": os.path.join(ROOT, r"03-out\sciadv-tpl01.pptx"),
        "body_layout": 3, "cover_layout": 0, "thanks_layout": 5, "cite_on_band": False,
        "content": (0.73, 11.87),
        "cover": {"keep_layout": True, "eyebrow": False},
    },
    # 简约-会议：16:9，但每个版式都铺满版深蓝 #2A5989 -> 文字必须反色
    "meeting": {
        "series": "R1",
        "template": os.path.join(ROOT, r"00-template\复旦简约-会议16：9.pptx"),
        "accent": (0x2A, 0x59, 0x89),
        "out": os.path.join(ROOT, r"03-out\sciadv-meeting.pptx"),
        "body_layout": 1, "cover_layout": 0,
        "content": (1.10, 11.32), "cite_on_band": False,
        # 校徽图在版式层 1.78–3.40in；标题从 3.98 起（用户指定：原副标题位置作标题顶）
        # 模板自带的 3.76in 装饰线正好当分隔线，所以不再另画
        "cover": {"eyebrow": False, "title_top": 4.55, "title_h": 1.20,
                   "rule": False, "rule_y": 3.76, "sub_y": 5.55},
        "title_top": 0.62, "emblem_top": 1.26,
        "thanks_top": 4.45, "thanks_rule": 5.95, "thanks_text": "谢谢聆听",
        # 深蓝只是外框，内容区是浅灰面板（实测 #E7E6E6 覆盖 0.14–13.19in），
        # 文字必须用深色；反白会整页看不见。
        "dark": False,
    },
    "meeting_white": {
        "series": "R1",
        "template": os.path.join(ROOT, r"00-template\复旦简约-会议16：9.pptx"),
        "accent": (0x2A, 0x59, 0x89),
        "out": os.path.join(ROOT, r"03-out\sciadv-meeting-white.pptx"),
        "body_layout": 1, "cover_layout": 0,
        "content": (1.10, 11.32), "cite_on_band": False,
        "cover": {"eyebrow_y": 1.28, "title_top": 1.72, "title_h": 1.72,
                   "rule": False, "rule_y": 3.80},
        "dark": False, "white_bg": True,
    },
    # 简约-蓝：4:3，封面整页铺色，正文右侧有装饰图 → 内容只用左半
    "jianyue": {
        "series": "R1",
        "template": os.path.join(ROOT, r"00-template\复旦简约-蓝4：3.pptx"),
        "accent": (0x2A, 0x59, 0x89),
        "out": os.path.join(ROOT, r"03-out\probe-jianyue.pptx"),
        "body_layout": 5, "cover_layout": 0,
        "content": (0.69, 5.75),
        "probe": [1, 5, 12, 13],
    },
}

_PAL_FILE = os.path.join(HERE, "色卡.json")
with io.open(_PAL_FILE, encoding="utf-8") as _f:
    _PAL = json.load(_f)
SERIES = {s["id"]: s["slots"] for s in _PAL["series"]}
SERIES_ORDER = [s["id"] for s in _PAL["series"] if s["kind"] == "make"]


def hex2rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def pick_colors(n, first):
    """按色卡 rules.contrast / monochrome 取 n 个颜色。first 为本模板所属色系。"""
    order = [first] + [x for x in SERIES_ORDER if x != first]
    out = []
    if n >= 4:
        idx = ["1", "2", "4"]
        for sid in order:
            for k in idx:
                out.append(SERIES[sid][{"1": 0, "2": 1, "4": 3}[k]])
                if len(out) == n:
                    return [hex2rgb(c) for c in out]
    else:
        for k in (0, 1, 3, 2, 4):
            c = SERIES[first][k]
            if c:
                out.append(c)
            if len(out) == n:
                break
    return [hex2rgb(c) for c in out]


BLACK = RGBColor(0x5C, 0x5C, 0x5C)
INK = RGBColor(0x2D, 0x46, 0x5B)
GRAY = RGBColor(0x88, 0x88, 0x88)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LINE = RGBColor(0xC8, 0xC8, 0xC8)
FONT = "微软雅黑"

SW, SH = 13.333, 7.5          # 画布，main 中按模板实际尺寸覆盖
CX = 1.095                    # 内容区左边界，main 中按版式标题占位符覆盖
CW = 11.15                    # 内容区宽度，同上

S_DISPLAY, S_H1, S_H2, S_BODY, S_SMALL, S_CAP, S_MICRO = 54, 32, 28, 22, 16, 13, 11


def sup2num(t):
    m = {"⁰": "0", "¹": "1", "²": "2", "³": "3", "⁴": "4", "⁵": "5",
         "⁶": "6", "⁷": "7", "⁸": "8", "⁹": "9", "⁻": "-"}
    return "".join(m.get(c, c) for c in t)


def parse_val(s):
    t = sup2num(s).replace(",", "").replace(">", "").replace("≈", "").strip()
    t = t.replace("×10", "e").replace("*10", "e").replace("x10", "e")
    try:
        if "e" in t:
            m, e = t.split("e")
            return float(m) * 10 ** int(float(e))
        return float(t)
    except Exception:
        return 1.0


D = {}
LATIN_FONT = "Times New Roman"


def apply_font(run, latin=None, ea=None):
    """A1 中英混排：西文/数字走 Times New Roman，中文走微软雅黑。

    同一 run 内用 a:latin / a:ea 两个声明区分，PowerPoint 会按字符自动取；
    比拆 run 省事，也不会破坏行内混排的断行。
    """
    latin = latin or LATIN_FONT
    ea = ea or FONT
    rPr = run._r.get_or_add_rPr()
    for tag, face in (("a:latin", latin), ("a:ea", ea), ("a:cs", latin)):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", face)


def txbox(s, x, y, w, h, text, size=S_BODY, bold=False, color=BLACK, spacing=1.8,
          align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
    box = s.shapes.add_textbox(Emu(int(x)), Emu(int(y)), Emu(int(w)), Emu(int(h)))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        r = p.add_run()
        r.text = line
        r.font.name = FONT
        apply_font(r)
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
    return box


def title_of(s, text):
    ph = None
    for sh in s.shapes:
        if sh.is_placeholder and sh.placeholder_format.type in (
                PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
            ph = sh
            break
    size = S_H1 if len(text) <= 24 else S_H2
    if ph is None:
        box = s.shapes.add_textbox(Inches(CX), Inches(TITLE_TOP), Inches(CW), Inches(1.18))
        tf = box.text_frame
    else:
        ph.left, ph.top = Inches(CX), Inches(TITLE_TOP)
        ph.width, ph.height = Inches(CW), Inches(1.18)
        tf = ph.text_frame
        tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = text
    r.font.name = FONT
    apply_font(r)
    r.font.size = Pt(size)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x33, 0x33, 0x33)


def conclusion(s, text):
    ln = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(CX), BY(5.92),
                            Emu(int(Inches(1.1))), Emu(int(Inches(0.035))))
    ln.fill.solid()
    ln.fill.fore_color.rgb = LINE
    ln.line.fill.background()
    txbox(s, Inches(CX), BY(6.05), Inches(CW), Inches(1.22), text, size=S_BODY)


def pagenum(s, n):
    txbox(s, CANVAS(SW - 1.20), CANVAS(SH - 0.52), Inches(0.45), Inches(0.3), f"{n:02d}",
          size=S_MICRO, color=GRAY, spacing=1.0, align=PP_ALIGN.RIGHT)


def rect(s, x, y, w, h, color, shape=MSO_SHAPE.RECTANGLE):
    shp = s.shapes.add_shape(shape, Emu(int(x)), Emu(int(y)), Emu(int(w)), Emu(int(h)))
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        shp.adjustments[0] = 0.04
    return shp


def resolve_fig(spec):
    # PDF 原图裁切优先；panel 只是历史兜底
    if spec.get("figure"):
        return os.path.join(ROOT, "02-build", spec["figure"])
    if spec.get("panel"):
        return os.path.join(PANELS, spec["panel"] + ".png")
    raise KeyError("配图缺失：既没有 figure 也没有 panel")


def place_figure(s, img_path, zone_x, zone_w, zone_y, zone_h):
    """按比例铺满可用区（contain）并居中：小图放大、大图缩小，绝不变形。"""
    pic = s.shapes.add_picture(img_path, Emu(int(zone_x)), Emu(int(zone_y)))
    scale = min(float(zone_w) / pic.width, float(zone_h) / pic.height)
    pic.width = Emu(int(pic.width * scale))
    pic.height = Emu(int(pic.height * scale))
    pic.left = Emu(int(zone_x + (zone_w - pic.width) / 2))
    pic.top = Emu(int(zone_y + (zone_h - pic.height) / 2))
    pic.left = Emu(int(zone_x + (zone_w - pic.width) / 2))
    pic.top = Emu(int(zone_y + (zone_h - pic.height) / 2))
    return pic


def p_cover(prs, spec, layout, closing=False):
    """封面：眉标 + 主标题 + 细线 + 副标题，底部绿带上放文献出处。
    本模板封面无背景图，白区 0–5.85in 必须被这套结构占住，否则整页空。
    """
    s = prs.slides.add_slide(layout)
    for sh in list(s.shapes):
        if not sh.is_placeholder:
            continue
        kind = sh.placeholder_format.type
        # 注意：不能用 "TITLE" in str(kind) —— "SUBTITLE" 也含 "TITLE"，
        # 会把副标题占位符当成标题再填一遍，出现两行完全重复的字。
        if kind in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
            if not COVER.get("keep_layout"):
                sh.left, sh.top = Inches(0.50), Inches(COVER.get("title_top", 2.18))
                sh.width, sh.height = CANVAS(SW - 1.00), Inches(COVER.get("title_h", 1.95))
            tf = sh.text_frame
            tf.clear()
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.TOP
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.line_spacing = 1.15
            r = p.add_run()
            r.text = spec["purpose"]
            r.font.name = FONT
            apply_font(r)
            cover_size = 40 if closing else 48
            if not closing:
                n_chars = len(spec["purpose"])
                cover_size = 48 if n_chars <= 18 else (40 if n_chars <= 24 else (34 if n_chars <= 30 else 28))
            r.font.size = Pt(cover_size)
            r.font.bold = True
            r.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
        else:
            sh._element.getparent().remove(sh._element)

    # 眉标
    if spec.get("eyebrow") and not closing and COVER.get("eyebrow", True):
        txbox(s, Inches(0.50), Inches(COVER.get("eyebrow_y", 1.62)), CANVAS(SW - 1.00), Inches(0.34),
              spec["eyebrow"], size=S_CAP, color=ACCENT, spacing=1.0, align=PP_ALIGN.CENTER)

    # 分隔线（模板自带同类装饰线时不重复画）
    rule_y = Inches(COVER.get("rule_y", 4.28)) if not closing else Inches(4.02)
    if COVER.get("rule", True):
        rl = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                CANVAS(SW / 2 - 0.62), Emu(int(rule_y)),
                                Emu(int(Inches(1.24))), Emu(int(Inches(0.045))))
        rl.fill.solid()
        rl.fill.fore_color.rgb = ACCENT
        rl.line.fill.background()

    # 副标题
    sub_y = Inches(COVER["sub_y"]) if COVER.get("sub_y") else rule_y + Inches(0.22)
    txbox(s, Inches(0.50), sub_y, CANVAS(SW - 1.00), Inches(0.62),
          spec.get("sub", ""), size=S_SMALL, color=RGBColor(0x5C, 0x5C, 0x5C),
          spacing=1.2, align=PP_ALIGN.CENTER)

    # 底部色带上的出处；没有色带的模板用深色字，否则白字会消失在浅底上
    if spec.get("cite") and not closing:
        txbox(s, Inches(0.62), Inches(6.62), CANVAS(SW - 1.2), Inches(0.34),
              spec["cite"], size=S_CAP,
              color=WHITE if CITE_ON_BAND else RGBColor(0x88, 0x88, 0x88), spacing=1.0)
    return s


def p_bars(prs, spec, layout, n):
    s = prs.slides.add_slide(layout)
    title_of(s, spec["purpose"])
    vals = [parse_val(b["value"]) for b in spec["bars"]]
    logs = [math.log10(max(v, 1)) for v in vals]
    top = max(logs) if max(logs) > 0 else 1
    lab_w = Inches(min(2.55, CW * 0.24))
    bar_x = Inches(CX + min(2.55, CW * 0.24) + 0.25)
    bar_max = max(CW - min(2.55, CW * 0.24) - 1.75, 1.2)
    y = Inches(1.95)
    bcols = pick_colors(len(spec["bars"]), SERIES_ID)
    for i, (b, lg) in enumerate(zip(spec["bars"], logs)):
        w = Inches(max(lg / top * bar_max, 0.35))
        txbox(s, Inches(CX), y + Inches(0.02), lab_w, Inches(0.5), b["label"],
              size=S_SMALL, color=INK, spacing=1.0)
        rect(s, bar_x, y, Emu(int(w)), Inches(0.38), bcols[i % len(bcols)])
        txbox(s, bar_x + Emu(int(w)) + Inches(0.12), y + Inches(0.02), Inches(1.9),
              Inches(0.5), b["value"], size=S_SMALL, bold=True, color=INK, spacing=1.0)
        y += Inches(0.92)
    txbox(s, Inches(CX), BY(4.55), Inches(CW), Inches(1.0), spec["support"],
          size=S_BODY, color=BLACK)
    conclusion(s, spec["conclusion"])
    pagenum(s, n)
    return s


def p_process(prs, spec, layout, n):
    s = prs.slides.add_slide(layout)
    title_of(s, spec["purpose"])
    steps = spec["blocks"]
    if CW < 9.0:
        # 窄栏（如 4:3 带侧图）：竖向缎带，一步一行
        y = Inches(1.60)
        for i, b in enumerate(steps):
            shp = s.shapes.add_shape(MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON,
                                     Inches(CX), Emu(int(y)), Inches(CW), Inches(0.52))
            shp.fill.solid()
            shp.fill.fore_color.rgb = ACCENT
            shp.line.fill.background()
            tf = shp.text_frame
            tf.word_wrap = True
            tf.margin_left = Emu(int(Inches(0.45 if i else 0.32)))
            tf.margin_right = Emu(int(Inches(0.35)))
            tf.margin_top = tf.margin_bottom = 0
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            r = p.add_run()
            r.text = f'{b["num"]}  {b["title"]}'
            r.font.name = FONT
            apply_font(r)
            r.font.size = Pt(S_SMALL)
            r.font.bold = True
            r.font.color.rgb = WHITE
            txbox(s, Inches(CX + 0.30), y + Inches(0.60), Inches(CW - 0.30), Inches(0.55),
                  b["desc"].replace("\n", " · "), size=S_CAP, color=BLACK, spacing=1.2)
            y += Inches(1.05)
        # 窄栏：支撑句只在放得下时上屏，放不下交给备注（缎带说明已承载细节）
        room = 5.92 - (y + 0.05)
        if spec.get("support") and room >= 0.75:
            sup_size = S_BODY if (len(spec["support"]) <= 26 and room >= 1.35) else S_CAP
            txbox(s, Inches(CX), y + Inches(0.05), Inches(CW),
                  Inches(min(room - 0.05, 1.35)),
                  spec["support"], size=sup_size, color=BLACK, spacing=1.45)
        conclusion(s, spec["conclusion"])
        pagenum(s, n)
        return s
    x0, total, overlap = Inches(CX), Inches(CW), Inches(0.18)
    bw = (total + 3 * overlap) / 4
    by, bh = Inches(1.52), Inches(0.80)
    for i, b in enumerate(steps):
        x = x0 + i * (bw - overlap)
        shp = s.shapes.add_shape(MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON,
                                 Emu(int(x)), Emu(int(by)), Emu(int(bw)), Emu(int(bh)))
        shp.fill.solid()
        shp.fill.fore_color.rgb = ACCENT
        shp.line.fill.background()
        tf = shp.text_frame
        tf.word_wrap = True
        tf.margin_left = Emu(int(Inches(0.42 if i else 0.30)))
        tf.margin_right = Emu(int(Inches(0.30)))
        tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = f'{b["num"]}  {b["title"]}'
        r.font.name = FONT
        apply_font(r)
        r.font.size = Pt(S_SMALL)
        r.font.bold = True
        r.font.color.rgb = WHITE
        txbox(s, x + Inches(0.04), by + Inches(0.92), bw - Inches(0.06), Inches(1.05),
              b["desc"], size=S_CAP, color=BLACK, spacing=1.45)
    txbox(s, Inches(CX), Inches(3.62), Inches(CW), Inches(1.3), spec["support"],
          size=S_BODY, color=BLACK)
    conclusion(s, spec["conclusion"])
    pagenum(s, n)
    return s


def p_funnel(prs, spec, layout, n):
    s = prs.slides.add_slide(layout)
    title_of(s, spec["purpose"])
    rows = spec["funnel"]
    zone_x, zone_w = Inches(CX), Inches(CW)
    widths = [Inches(10.2), Inches(8.4), Inches(5.6)]
    y = Inches(1.90)
    fcols = [PAL[0], PAL[1], PAL[min(2, len(PAL) - 1)]]
    for i, (r, bw) in enumerate(zip(rows, widths)):
        x = zone_x + (zone_w - bw) / 2
        rect(s, x, y, bw, Inches(0.78), fcols[i % len(fcols)], shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        txbox(s, x, y, bw, Inches(0.78), f'{r["label"]}   {r["value"]}',
              size=S_SMALL, bold=True, color=WHITE, spacing=1.0,
              align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        y += Inches(0.98)
    txbox(s, Inches(CX), Inches(4.58), Inches(CW), Inches(1.25), spec["support"],
          size=S_BODY, color=BLACK, spacing=1.5)
    conclusion(s, spec["conclusion"])
    pagenum(s, n)
    return s


def p_numbers(prs, spec, layout, n):
    s = prs.slides.add_slide(layout)
    title_of(s, spec["purpose"])
    gap = 0.35
    col_w = Inches(min(3.5, (CW - 2 * gap) / 3))
    total = 3 * col_w + 2 * gap
    x = (CANVAS(SW) - total) / 2
    for st in spec["stats"]:
        txbox(s, x, Inches(1.95), col_w, Inches(0.95), st["num"], size=S_DISPLAY, bold=True,
              color=INK, spacing=1.0)
        txbox(s, x, Inches(2.98), col_w, Inches(0.85), st["label"], size=S_CAP, color=BLACK,
              spacing=1.5)
        x += col_w + gap
    txbox(s, Inches(CX), Inches(4.05), Inches(CW), Inches(1.2), spec["support"],
          size=S_BODY, color=BLACK)
    conclusion(s, spec["conclusion"])
    pagenum(s, n)
    return s


def p_figure_full(prs, spec, layout, n):
    s = prs.slides.add_slide(layout)
    title_of(s, spec["purpose"])
    has_text = bool(spec.get("support"))
    zone_h = Inches(3.30) if has_text else Inches(4.15)
    place_figure(s, resolve_fig(spec), Inches(CX), Inches(CW), BY(1.45), zone_h)
    if has_text:
        txbox(s, Inches(CX), BY(4.86), Inches(CW), Inches(1.05), spec["support"],
              size=S_BODY, color=BLACK, spacing=1.5)
    conclusion(s, spec["conclusion"])
    pagenum(s, n)
    return s


def p_split(prs, spec, layout, n):
    s = prs.slides.add_slide(layout)
    title_of(s, spec["purpose"])
    text_w, gutter, img_w = BY(4.55), Inches(0.55), Inches(6.0)
    total = text_w + gutter + img_w
    x = (CANVAS(SW) - total) / 2
    for i, line in enumerate(spec["bullets"]):
        txbox(s, x, Inches(1.62) + i * Inches(1.02), text_w, Inches(1.05),
              "· " + line, size=S_BODY, color=BLACK, spacing=1.5)
    place_figure(s, resolve_fig(spec), x + text_w + gutter, img_w,
                 Inches(1.45), Inches(3.2))
    conclusion(s, spec["conclusion"])
    pagenum(s, n)
    return s


def p_bullets(prs, spec, layout, n):
    s = prs.slides.add_slide(layout)
    title_of(s, spec["purpose"])
    for i, line in enumerate(spec["bullets"]):
        y = Inches(1.95) + i * Inches(1.02)
        rect(s, Inches(CX), y + Inches(0.14), Inches(0.10), Inches(0.50), ACCENT)
        txbox(s, Inches(CX) + Inches(0.38), y, Inches(CW) - Inches(0.4), Inches(0.95),
              line, size=S_BODY, color=BLACK, spacing=1.4)
    if spec.get("conclusion"):
        conclusion(s, spec["conclusion"])
    pagenum(s, n)
    return s


def hanging(par, mar_l=0.45, indent=-0.45):
    """悬挂缩进：MLA 参考文献第二行起缩进。"""
    pPr = par._p.get_or_add_pPr()
    pPr.set("marL", str(int(mar_l * 914400)))
    pPr.set("indent", str(int(indent * 914400)))


GB_AUTHOR_SPLIT = re.compile(r"[,，、;；]\s*")


def _gb_authors(authors, lang):
    """GB/T 7714：≤3 位全列，>3 位列前 3 + 等 / et al.；西文姓全大写、名缩写。"""
    if not authors:
        return ""
    names = [a.strip() for a in GB_AUTHOR_SPLIT.split(authors) if a.strip()]
    over = len(names) > 3
    names = names[:3]
    if lang == "en":
        out = []
        for n in names:
            parts = n.split()
            if len(parts) >= 2:
                surname = parts[-1].upper()
                initials = " ".join(p[0].upper() for p in parts[:-1])
                out.append("%s %s" % (surname, initials))
            else:
                out.append(n.upper())
        tail = "et al" if over else ""
    else:
        out = names
        tail = "等" if over else ""
    s = ",".join(out)
    if tail:
        s = s + ", " + tail if False else s + "," + tail
    return s


def fmt_gbt7714(ref):
    """GB/T 7714-2015 顺序编码制，一行一条，返回文本。"""
    lang = ref.get("lang", "zh")
    au = _gb_authors(ref.get("authors", ""), lang)
    t = ref.get("title", "")
    kind = ref.get("type", "J")
    if kind == "J":
        body = "%s[J].%s,%s" % (t, ref.get("journal", ""), ref.get("year", ""))
        if ref.get("vol"):
            body += ",%s" % ref["vol"]
        if ref.get("issue"):
            body += "(%s)" % ref["issue"]
        if ref.get("pages"):
            body += ":%s" % ref["pages"]
    elif kind in ("M", "C", "D", "R"):
        place = ref.get("place", "")
        pub = ref.get("publisher", "")
        body = "%s[%s].%s:%s,%s" % (t, kind, place, pub, ref.get("year", ""))
        if ref.get("pages"):
            body += ":%s" % ref["pages"]
    else:  # EB/OL 等电子资源
        body = "%s[%s/OL].(%s)[%s].%s" % (t, kind, ref.get("issued", ""), ref.get("accessed", ""),
                                          ref.get("url", ""))
    if ref.get("doi"):
        body += ".DOI:%s" % ref["doi"]
    body = body.rstrip(".") + "."
    return (au + "." + body) if au else body


def pick_cite_style(spec):
    """A3：全中文 → GB/T 7714；中英混排 → MLA9；显式 cite_style 优先。"""
    style = spec.get("cite_style") or D.get("deck", {}).get("cite_style") or "auto"
    if style in ("mla9", "gbt7714"):
        return style
    refs = spec.get("refs") or []
    langs = [r.get("lang") for r in refs]
    if refs and all(l == "zh" for l in langs):
        return "gbt7714"
    return "mla9"


def _cell_border(cell, top=None, bottom=None):
    """三线表：只留顶/底线，其余 noFill（不用 #000 竖线，也不用框）。"""
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for el in tcPr.findall(qn(tag)):
            tcPr.remove(el)
    order = [("a:lnL", None), ("a:lnR", None), ("a:lnT", top), ("a:lnB", bottom)]
    idx = 0
    for tag, w_pt in order:
        ln = tcPr.makeelement(qn(tag), {"w": str(int((w_pt or 0) * 12700)), "cap": "flat",
                                        "cmpd": "sng", "algn": "ctr"})
        if w_pt:
            fill = ln.makeelement(qn("a:solidFill"), {})
            clr = fill.makeelement(qn("a:srgbClr"), {"val": "000000"})
            fill.append(clr)
            ln.append(fill)
        else:
            ln.append(ln.makeelement(qn("a:noFill"), {}))
        tcPr.insert(idx, ln)
        idx += 1


def p_table(prs, spec, layout, n):
    """三线表页：无竖线、无填充；顶/底 1.5pt，表头下线 0.75pt；表注只进备注。"""
    s = prs.slides.add_slide(layout)
    title_of(s, spec["purpose"])
    cols = spec["columns"]
    rows = spec["rows"]
    widths = spec.get("widths") or [1.0] * len(cols)
    total = float(sum(widths))
    nrow = len(rows) + 1
    top = Inches(spec.get("top", 1.85))
    tbl_w = Inches(CW)
    row_h = Inches(spec.get("row_h", 0.50))
    shape = s.shapes.add_table(nrow, len(cols), Inches(CX), top, tbl_w, row_h * nrow)
    table = shape.table
    table.first_row = False
    table.horz_banding = False
    for j, wgt in enumerate(widths):
        table.columns[j].width = Emu(int(tbl_w * (wgt / total)))
    for i in range(nrow):
        table.rows[i].height = row_h
    fs = spec.get("size", 16)
    for i, row in enumerate([cols] + rows):
        for j, val in enumerate(row):
            cell = table.cell(i, j)
            cell.fill.background()
            cell.margin_left = cell.margin_right = Inches(0.10)
            cell.margin_top = cell.margin_bottom = Inches(0.04)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if j else PP_ALIGN.LEFT
            r = p.add_run()
            r.text = str(val)
            r.font.name = FONT
            apply_font(r)
            r.font.size = Pt(fs)
            r.font.bold = (i == 0)
            r.font.color.rgb = BLACK
            top_w = 1.5 if i == 0 else None
            bot_w = 0.75 if i == 0 else (1.5 if i == nrow - 1 else None)
            _cell_border(cell, top=top_w, bottom=bot_w)
    if spec.get("note"):
        s.notes_slide.notes_text_frame.text = spec["note"]
    pagenum(s, n)
    return s



CHART_DIR = os.path.join(ROOT, "02-build", "charts")
RSCRIPT = os.environ.get("DPPT_RSCRIPT", r"D:\download\R\R-4.3.3\bin\Rscript.exe")
CHART_ENGINE = {"used": None, "detail": ""}


def _r_script_path():
    for base in (os.path.dirname(os.path.abspath(__file__)),
                 r"D:\dsh-wx\dppt-plugin\engine"):
        p = os.path.join(base, "plot_chart.R")
        if os.path.exists(p):
            return p
    return None


def _slug(s):
    out = "".join(ch if ch.isalnum() else "_" for ch in s)
    return out[:40]


def chart_png(spec, w_in=7.2, h_in=3.2):
    """返回 (png 路径 或 None, 引擎说明)。顺序：R → matplotlib → None(调用方走原生)。"""
    os.makedirs(CHART_DIR, exist_ok=True)
    kind = spec.get("kind", "bar")
    cats = spec.get("categories") or []
    series = spec.get("series") or []
    key = _slug(kind + "_" + "_".join(cats) + "_" + "_".join(s.get("name", "") for s in series))
    png = os.path.join(CHART_DIR, key + ".png")
    csv = os.path.join(CHART_DIR, key + ".csv")

    # --- R ---
    rscript = _r_script_path()
    if os.path.exists(RSCRIPT) and rscript:
        with io.open(csv, "w", encoding="utf-8", newline="") as f:
            f.write("label," + ",".join(s.get("name", "s%d" % i) for i, s in enumerate(series)) + "\n")
            for i, c in enumerate(cats):
                f.write(c + "," + ",".join(str(s["values"][i]) for s in series) + "\n")
        pal = ",".join(spec.get("palette") or ["#537B61", "#A8BC8F", "#354C3D", "#D7B4B1"])
        try:
            import subprocess
            r = subprocess.run([RSCRIPT, rscript, kind, csv, png, pal, str(w_in), str(h_in), "300"],
                               capture_output=True, text=True, timeout=120)
            if r.returncode == 0 and os.path.exists(png):
                CHART_ENGINE.update(used="R", detail="R 4.3.3 base graphics")
                return png, "R 4.3.3（base graphics，cairo，300dpi）"
            CHART_ENGINE.update(detail=(r.stderr or "")[-300:])
        except Exception as e:
            CHART_ENGINE.update(detail=str(e))

    # --- matplotlib ---
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib import font_manager
        for cand in (r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\simhei.ttf"):
            if os.path.exists(cand):
                font_manager.fontManager.addfont(cand)
                plt.rcParams["font.family"] = font_manager.FontProperties(fname=cand).get_name()
                break
        fig, ax = plt.subplots(figsize=(w_in, h_in), dpi=300)
        pal = spec.get("palette") or ["#537B61", "#A8BC8F", "#354C3D"]
        if kind == "pie":
            ax.pie([s["values"][0] for s in series] or [], labels=cats,
                   colors=pal * 3, wedgeprops=dict(color="white"))
        elif kind == "line":
            for i, s in enumerate(series):
                ax.plot(cats, s["values"], marker="o", color=pal[i % len(pal)], label=s.get("name", ""))
            if len(series) > 1:
                ax.legend(frameon=False)
        else:
            width = 0.8 / max(len(series), 1)
            for i, s in enumerate(series):
                ax.bar([x + i * width for x in range(len(cats))], s["values"], width=width,
                       color=pal[i % len(pal)], label=s.get("name", ""))
            ax.set_xticks(range(len(cats)))
            ax.set_xticklabels(cats)
            if len(series) > 1:
                ax.legend(frameon=False)
        for spine in ("top", "right"):
            ax.spines[spine].set_visible(False)
        fig.tight_layout()
        fig.savefig(png)
        plt.close(fig)
        CHART_ENGINE.update(used="matplotlib", detail="matplotlib")
        return png, "matplotlib（300dpi）"
    except Exception as e:
        CHART_ENGINE.update(detail=str(e))

    return None, "PowerPoint 原生图形"


def _native_chart(s, spec, x, y, w, h):
    """兜底：用 PowerPoint 原生 chart（可编辑）。"""
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    kind = spec.get("kind", "bar")
    ctype = {"bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
             "line": XL_CHART_TYPE.LINE_MARKERS,
             "pie": XL_CHART_TYPE.PIE}.get(kind, XL_CHART_TYPE.COLUMN_CLUSTERED)
    cd = CategoryChartData()
    cd.categories = spec.get("categories") or []
    for s_ in spec.get("series") or []:
        cd.add_series(s_.get("name", ""), s_.get("values") or [])
    gf = s.shapes.add_chart(ctype, Emu(int(x)), Emu(int(y)), Emu(int(w)), Emu(int(h)), cd)
    return gf.chart


def p_chart(prs, spec, layout, n):
    """图表页：R → matplotlib → 原生。引擎名进备注并在 stdout 说明。"""
    s = prs.slides.add_slide(layout)
    title_of(s, spec["purpose"])
    ch = spec.get("chart") or {}
    if spec.get("palette"):
        ch["palette"] = spec["palette"]
    png, engine = chart_png(ch, w_in=7.2, h_in=3.2)
    zone_y, zone_h = Inches(1.45), Inches(3.30)
    if png:
        place_figure(s, png, Inches(CX), Inches(CW), zone_y, zone_h)
    else:
        _native_chart(s, ch, Inches(CX), zone_y, Inches(CW), zone_h)
    if spec.get("support"):
        txbox(s, Inches(CX), BY(4.86), Inches(CW), Inches(1.05), spec["support"],
              size=S_BODY, color=BLACK, spacing=1.5)
    conclusion(s, spec["conclusion"])
    note = (spec.get("notes") or "").strip()
    note = (note + " " if note else "") + "图：%s。" % engine
    s.notes_slide.notes_text_frame.text = note
    print("  chart p%s 引擎：%s" % (spec.get("n"), engine))
    pagenum(s, n)
    return s


def p_refs(prs, spec, layout, n):
    """参考文献页：正文页形式，标题按模板默认左上；条目 MLA 9，悬挂缩进。"""
    s = prs.slides.add_slide(layout)
    title_of(s, spec["purpose"])
    y = Inches(1.62)
    for ref in spec["refs"]:
        est_w = Inches(CW) - Inches(0.45)
        chars_per_line = max(int(est_w / (16 * 12700)), 20)
        _segs = ref["runs"] if pick_cite_style(spec) == "mla9" else [{"t": fmt_gbt7714(ref)}]
        text_len = sum(len(r["t"]) for r in _segs)
        lines = max(1, -(-text_len // chars_per_line))
        h = Inches(0.34 * lines + 0.14)
        box = s.shapes.add_textbox(Inches(CX), Emu(int(y)), Inches(CW), Emu(int(h)))
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        par = tf.paragraphs[0]
        par.line_spacing = 1.55
        hanging(par)
        style = pick_cite_style(spec)
        segs = ref["runs"] if style == "mla9" else [{"t": fmt_gbt7714(ref)}]
        if style == "gbt7714":
            segs = [{"t": "[%d] " % (spec["refs"].index(ref) + 1)}] + segs
        for seg in segs:
            r = par.add_run()
            r.text = seg["t"]
            r.font.name = FONT
            apply_font(r)
            r.font.size = Pt(S_SMALL)
            r.font.italic = bool(seg.get("i"))
            r.font.color.rgb = BLACK
        y += Emu(int(h)) + Inches(0.20)
    pagenum(s, n)
    return s


def p_thanks(prs, spec, layout, n):
    """封底：只有「致谢」，居中。不要标题、副标题、出处、结论句。"""
    s = prs.slides.add_slide(layout)
    for sh in list(s.shapes):
        if sh.is_placeholder:
            sh._element.getparent().remove(sh._element)
    txbox(s, Inches(0.50), Inches(THANKS_TOP), CANVAS(SW - 1.00), Inches(1.30),
          THANKS_TEXT, size=48, bold=True, color=RGBColor(0x33, 0x33, 0x33),
          spacing=1.15, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    rl = s.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                            CANVAS(SW / 2 - 0.62), Emu(int(Inches(THANKS_RULE))),
                            Emu(int(Inches(1.24))), Emu(int(Inches(0.045))))
    rl.fill.solid()
    rl.fill.fore_color.rgb = ACCENT
    rl.line.fill.background()
    return s



def _theme_accent(path):
    """从 ppt/theme/theme1.xml 读 accent1，读不到返回 None。"""
    import re
    import zipfile
    try:
        with zipfile.ZipFile(path) as z:
            names = [n for n in z.namelist() if n.startswith("ppt/theme/theme") and n.endswith(".xml")]
            if not names:
                return None
            xml = z.read(sorted(names)[0]).decode("utf-8", "ignore")
        m = re.search(r"<a:accent1>.*?<a:srgbClr val=\"([0-9A-Fa-f]{6})\"", xml, re.S)
        if m:
            return "#" + m.group(1).upper()
        m = re.search(r"<a:accent1>.*?<a:sysClr[^>]*lastClr=\"([0-9A-Fa-f]{6})\"", xml, re.S)
        if m:
            return "#" + m.group(1).upper()
    except Exception:
        pass
    return None


def _nearest_series(hex_color):
    """按 accent 与各色系首槽的 RGB 距离挑最接近的色系。"""
    if not hex_color:
        return "R3", "R3"
    try:
        r = int(hex_color[1:3], 16)
        g = int(hex_color[3:5], 16)
        b = int(hex_color[5:7], 16)
    except Exception:
        return "R3", "R3"
    best, bestd = None, 1e9
    for sid, slots in SERIES.items():
        c = slots[0]
        if not c:
            continue
        cr, cg, cb = int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16)
        d = (r - cr) ** 2 + (g - cg) ** 2 + (b - cb) ** 2
        if d < bestd:
            best, bestd = sid, d
    return best or "R3", hex_color


def auto_profile(path):
    """从模板推断版式参数，供 --auto 使用。"""
    prs = Presentation(path)
    sw, sh = prs.slide_width / 914400, prs.slide_height / 914400
    body_i, best = 0, -1.0
    for i, lay in enumerate(prs.slide_layouts):
        types = [ph.placeholder_format.type for ph in lay.placeholders]
        has_title = any(t in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE) for t in types)
        has_body = any(t in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT) for t in types)
        score = (2.0 if has_title else 0) + (2.0 if has_body else 0) + min(len(types), 4) * 0.1
        if score > best:
            best, body_i = score, i
    cx, cw = None, None
    for ph in prs.slide_layouts[body_i].placeholders:
        if ph.placeholder_format.type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
            cx, cw = ph.left / 914400, ph.width / 914400
            break
    _k = sw / 13.333
    if cx is None or cw < 0.60 * sw:      # 标题占位符过窄时退回整幅内容区
        cx, cw = 0.06 * 13.333, 0.88 * 13.333
    else:
        cx, cw = cx / _k, cw / _k          # 画布英寸 → 设计英寸（build 里还要乘 K）
    accent = _theme_accent(path)
    series, accent_hex = _nearest_series(accent)
    return {
        "template": path,
        "body_layout": body_i,
        "cover_layout": 0,
        "thanks_layout": 0,
        "content": [round(cx, 3), round(cw, 3)],
        "series": series,
        "accent": accent_hex,
        "canvas": [round(sw, 2), round(sh, 2)],
        "out": "",
        "auto": True,
    }



def ensure_design(path):
    """版式里若没有装饰形状（设计全画在幻灯片上），先把它搬进版式再出稿。

    只看"有没有任意一个版式带装饰"不够：08 那套模板的第 3 个版式有 2 个孤立形状，
    但前两个版式是空的，于是会被误判成"已设计"，正文页套上去就是白页。
    判据改成前两个版式都带装饰才算已设计。
    """
    try:
        prs = Presentation(path)
        lays = list(prs.slide_layouts)

        def deco_of(lay):
            return sum(1 for sh in lay.shapes if not sh.is_placeholder)

        if len(lays) >= 2 and deco_of(lays[0]) >= 2 and deco_of(lays[1]) >= 2:
            return path, None
        if len(lays) == 1 and deco_of(lays[0]) >= 2:
            return path, None
        import prepare_template as PT
        cache = os.environ.get("DPPT_CACHE") or os.path.join(ROOT, "03-out", "prepared-cache")
        os.makedirs(cache, exist_ok=True)
        name = re.sub(r"[^\w\u4e00-\u9fff.-]+", "_", os.path.basename(path))[:80]
        out = os.path.join(cache, name + ".prepared.pptx")
        meta = PT.build_prepared(path, out)
        return out, meta
    except Exception:
        return path, None




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

def stamp_design(prs, raw_path, cover_i, body_i):
    """把原始模板的封面/正文设计逐页贴到底层（单版式模板用）。"""
    import copy as _copy
    raw = Presentation(raw_path)
    cover = raw.slides[cover_i]
    body = raw.slides[body_i]
    for idx, slide in enumerate(prs.slides):
        which = os.environ.get("DPPT_STAMP_WHICH", "both")
        if which == "cover" and idx != 0:
            continue
        if which == "body" and idx == 0:
            continue
        src = cover if idx == 0 else body
        spTree = slide.shapes._spTree
        pos = 2
        limit = int(os.environ.get("DPPT_STAMP_LIMIT", "0") or 0)
        taken = 0
        for shp in src.shapes:
            if not prepare_template_is_deco(shp):
                continue
            if limit and taken >= limit:
                break
            if str(shp.shape_type).startswith("PICTURE") and not os.environ.get("DPPT_STAMP_NOPIC"):
                try:
                    pic = slide.shapes.add_picture(io.BytesIO(shp.image.blob),
                                                   shp.left, shp.top, shp.width, shp.height)
                    bf = shp._element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}srcRect")
                    if bf is not None:
                        for side in ("l", "t", "r", "b"):
                            v = bf.get(side)
                            if v:
                                setattr(pic, "crop_" + {"l": "left", "t": "top", "r": "right", "b": "bottom"}[side],
                                        int(v) / 100000.0)
                except Exception:
                    pass
                continue
            if os.environ.get("DPPT_STAMP_PICONLY"):
                continue
            try:
                el = _copy.deepcopy(shp._element)
                taken += 1
                base = 0
                for node in spTree.iter():
                    if node.tag.endswith("}cNvPr"):
                        try:
                            base = max(base, int(node.get("id") or 0))
                        except Exception:
                            pass
                if base:
                    for node in el.iter():
                        tag = node.tag.split("}")[-1]
                        if tag in ("cNvPr", "stCxn", "endCxn"):
                            try:
                                node.set("id", str(int(node.get("id") or 0) + base))
                            except Exception:
                                pass
                for node in list(el.iter()):
                    for attr in list(node.attrib.keys()):
                        if not attr.startswith("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"):
                            continue
                        rid = node.get(attr)
                        try:
                            rel = src.part.rels[rid]
                            new_rid = slide.part.relate_to(rel.target_part, rel.reltype)
                            node.set(attr, new_rid)
                        except Exception:
                            node.attrib.pop(attr, None)
                cleanup_dangling(el)
                spTree.insert(pos, el)
                pos += 1
            except Exception:
                pass


def prepare_template_is_deco(shp):
    try:
        txt = shp.text_frame.text.strip() if shp.has_text_frame else ""
    except Exception:
        txt = ""
    if txt and len(txt) > 4:
        return False
    xml = shp._element.xml
    # 组合里含图片时整体跳过：组内图片无法用 add_picture 重建，搬 XML 会坏包
    if str(shp.shape_type).startswith("GROUP") and "}pic" in xml:
        return False
    for bad in ("}graphicFrame", "}oleObj", "}media", "a:graphicData", "a:blipFill"):
        if bad in xml:
            return False
    return True


def main():
    global ACCENT, OUT, SW, SH, CX, CW
    import argparse
    ap = argparse.ArgumentParser(description="复旦叙事 PPT 生成器")
    ap.add_argument("profile", nargs="?", default="crown",
                    help="版式 profile：" + ", ".join(sorted(PROFILES)))
    ap.add_argument("--template", help="覆盖模板 pptx 绝对路径")
    ap.add_argument("--slides", help="覆盖大纲 JSON 路径")
    ap.add_argument("--out", help="覆盖输出 pptx 路径")
    ap.add_argument("--list", action="store_true", help="只列 profile 名")
    ap.add_argument("--auto", action="store_true",
                    help="不写死 profile：从 --template 的版式/主题色推断（批量试新模板用）")
    args = ap.parse_args()
    if args.list:
        for k in sorted(PROFILES):
            print(k, "\t", PROFILES[k]["template"])
        return
    prof_name = args.profile
    if getattr(args, "auto", False) or prof_name == "auto":
        if not args.template:
            raise SystemExit("--auto 需要同时给 --template")
        tpl_path, meta = ensure_design(args.template)
        if meta:
            print("检测到空版式：已把设计搬进版式（封面取 p%d，正文取 p%d）"
                  % (meta.get("cover_slide", 0) + 1, meta.get("body_slide", 0) + 1))
            args.template = tpl_path
        prof = auto_profile(tpl_path)
        prof_name = "auto"
        if meta and meta.get("single_layout"):
            globals()["STAMP"] = (args.template, meta.get("cover_slide", 0), meta.get("body_slide", 0))
            os.environ["DPPT_STAMP_PICONLY"] = "1"   # 单版式模板只贴图片，矢量装饰跳过
        if meta:
            # prepare 的约定：版式 0 = 封面设计，版式 1 = 正文设计
            prof["cover_layout"] = 0
            prof["body_layout"] = 0 if meta.get("single_layout") else 1
            prof["thanks_layout"] = 0
            _k = prof["canvas"][0] / 13.333
            if meta.get("title_zone"):
                prof["title_top"] = round(meta["title_zone"][1] / _k, 3)
            if meta.get("content_zone"):
                prof["body_dy"] = round(meta["content_zone"][1] / _k - 1.45, 3)
    elif prof_name not in PROFILES:
        raise SystemExit("未知 profile：%s；可选：%s" % (prof_name, ", ".join(sorted(PROFILES))))
    else:
        prof = PROFILES[prof_name]
    global SERIES_ID, CITE_ON_BAND
    SERIES_ID = prof["series"]
    CITE_ON_BAND = prof.get("cite_on_band", True)
    globals()["COVER"] = prof.get("cover", {})
    globals()["WHITE_BG"] = prof.get("white_bg", False)
    globals()["TITLE_TOP"] = prof.get("title_top", 0.08)
    globals()["THANKS_TOP"] = prof.get("thanks_top", 3.02)
    globals()["THANKS_RULE"] = prof.get("thanks_rule", 4.42)
    globals()["THANKS_TEXT"] = prof.get("thanks_text", "致谢")
    globals()["BODY_DY"] = float(prof.get("body_dy", 0.0) or 0.0)
    ACCENT = hex2rgb(SERIES[prof["series"]][0])
    # 数据/图元只取 1、2、4、5 位：3 位是浅底，白字压不住，规则里列为保留槽
    slots = SERIES[prof["series"]]
    globals()["PAL"] = [hex2rgb(c) for i, c in enumerate(slots) if c and i != 2]
    OUT = args.out or prof.get("out") or os.path.join(ROOT, r"03-out\deck.pptx")
    if args.template:
        prof["template"] = args.template
    if args.slides:
        globals()["SLIDES"] = args.slides
    with io.open(SLIDES, encoding="utf-8") as f:
        globals()["D"] = d_local = json.load(f)
    if prof.get("auto"):
        print("auto profile: 画布 %.2fx%.2f  正文版式 #%d  内容区 %.2f+%.2f  accent %s  色系 %s"
              % (prof["canvas"][0], prof["canvas"][1], prof["body_layout"],
                 prof["content"][0], prof["content"][1], prof["accent"], prof["series"]))
    try:
        _sw = Presentation(prof["template"]).slide_width / 914400
        globals()["K"] = round(_sw / 13.333, 4)
    except Exception:
        globals()["K"] = 1.0
    prs = Presentation(prof["template"])
    for sldId in list(prs.slides._sldIdLst):
        prs.part.drop_rel(sldId.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"))
        prs.slides._sldIdLst.remove(sldId)

    # 画布与内容区：从模板实际几何推导，不写死
    SW, SH = prs.slide_width / 914400, prs.slide_height / 914400
    lay_body = prs.slide_layouts[prof["body_layout"]]
    lay_cover = prs.slide_layouts[prof["cover_layout"]]
    lay_thanks = prs.slide_layouts[prof.get("thanks_layout", prof["cover_layout"])]
    if prof.get("emblem_top") is not None:
        for lay in {id(l): l for l in (lay_cover, lay_thanks)}.values():
            for shp in lay.shapes:
                if shp.name.startswith("图片") or str(shp.shape_type).startswith("PICTURE"):
                    shp.top = Inches(prof["emblem_top"])
    if prof.get("content"):
        CX, CW = prof["content"]
    else:
        for sh in lay_body.shapes:
            if sh.is_placeholder and sh.placeholder_format.type in (
                    PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
                CX, CW = sh.left / 914400, sh.width / 914400
                break

    fn = {"cover": lambda sp: p_cover(prs, sp, lay_cover),
          "refs": lambda sp: p_refs(prs, sp, lay_body, sp["n"]),
          "thanks": lambda sp: p_thanks(prs, sp, lay_thanks, sp["n"]),
          "bars": lambda sp: p_bars(prs, sp, lay_body, sp["n"]),
          "process": lambda sp: p_process(prs, sp, lay_body, sp["n"]),
          "funnel": lambda sp: p_funnel(prs, sp, lay_body, sp["n"]),
          "numbers": lambda sp: p_numbers(prs, sp, lay_body, sp["n"]),
          "figure-full": lambda sp: p_figure_full(prs, sp, lay_body, sp["n"]),
          "figure-split": lambda sp: p_split(prs, sp, lay_body, sp["n"]),
          "bullets": lambda sp: p_bullets(prs, sp, lay_body, sp["n"]),
          "table": lambda sp: p_table(prs, sp, lay_body, sp["n"]),
          "chart": lambda sp: p_chart(prs, sp, lay_body, sp["n"])}

    if prof.get("probe"):
        want = set(prof["probe"])
        for spec in d_local["slides"]:
            if spec["n"] in want:
                fn[spec["kind"]](spec)
        print(f"probe {prof_name}: 画布 {SW:.2f}x{SH:.2f}  内容区 {CX:.2f}+{CW:.2f}  "
              f"生成 {len(want)} 页")
    else:
        for spec in d_local["slides"]:
            fn[spec["kind"]](spec)

    if WHITE_BG:
        for sl in prs.slides:
            try:
                sl.background.fill.solid()
                sl.background.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            except Exception as e:
                print("bg skip", e)

    # 清掉继承自版式的空占位符：它们的填充会盖出一块浅色矩形，把暗底和文字都毁掉
    for sl in prs.slides:
        for shp in list(sl.shapes):
            if not shp.is_placeholder:
                continue
            if shp.placeholder_format.type in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
                continue
            if not shp.has_text_frame or not shp.text_frame.text.strip():
                shp._element.getparent().remove(shp._element)

    if prof.get("dark"):
        LIGHT = RGBColor(0xF2, 0xF2, 0xF2)
        for sl in prs.slides:
            for shp in sl.shapes:
                if not shp.has_text_frame:
                    continue
                for par in shp.text_frame.paragraphs:
                    for r in par.runs:
                        try:
                            c = r.font.color.rgb
                        except Exception:
                            continue
                        if c is None:
                            continue
                        lum = (0.299 * c[0] + 0.587 * c[1] + 0.114 * c[2]) / 255
                        if lum < 0.55:
                            r.font.color.rgb = LIGHT

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    if globals().get("STAMP"):
        stamp_design(prs, *STAMP)
    prs.save(OUT)
    print("saved", OUT, "slides", len(prs.slides))


def _delegate_master():
    """profile=master 时转给 build_master.py（母版模式：只认版式与占位符）。"""
    import build_master
    sys.argv = [sys.argv[0]] + sys.argv[2:]
    build_master.main()


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "master":
        _delegate_master()
        raise SystemExit(0)
    main()









