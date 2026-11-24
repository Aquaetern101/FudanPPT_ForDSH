# 复旦风格 PPT 插件 · v0.1.0（第一批）

> 给 DSH 用的小插件：把一篇论文、一份报告、一个 Excel，变成一套能上台讲的复旦风格 PPT。
> 这件事原本要花一整晚，现在大部分时间花在「想清楚讲什么」上，剩下交给它。

**版本 v0.1.0** ｜ 只支持第一批 8 套模板 ｜ 模板独立仓库、按需自动下载 ｜ 第二批见 `SECOND-BATCH.md`

---

## 快速启动

**1. 装插件**

```bash
pnpm add file:<解压路径>/plugin
```

更建议丢给 DSH 自己装喔，直接告诉它：“我要装https://github.com/Aquaetern101/FudanPPT_ForDSH”的插件

```
plugin_manager install_bundle <解压路径>/plugin
```

**2. 模板：插件自己去下（不用管）**

模板单独放在仓库 `FudanPPT_ForDSH_Templates`。第一次用某套模板时，
插件会按 `templates.json` 里登记的 `raw` 地址，只下这一套（最大 25 MB），
存到模板目录里，以后直接用。clone 主仓库只有 1 MB。

想自己管文件也行，三种方式任选：

- 把 pptx 放到模板目录里（默认 `DPPT_ROOT/00-template/`，可用 `DPPT_TEMPLATE_DIR` 改）
- 整包下载 Releases 里的 `FudanPPT_ForDSH_templates.zip` 解压进去
- 离线环境设 `DPPT_TEMPLATE_FETCH=0`，插件就不再尝试联网

```bash
export DPPT_TEMPLATE_DIR=<你的模板目录>     # 可选，默认 DPPT_ROOT/00-template
```

文件名要保持原样（插件按名字找），想改名就改 `templates.json` 里的 `file`。

**3. 出一份试试**

```bash
export DPPT_ROOT=<你的 dataRoot>
export DPPT_TEMPLATE_DIR=<你的模板目录>

python plugin/engine/make_outline.py --input 论文.pdf --out outline.json
python plugin/engine/build_deck.py meeting \
       --template "$DPPT_TEMPLATE_DIR/复旦简约-会议16：9.pptx" \
       --slides outline.json --out out/deck.pptx
```

打开 `out/deck.pptx`，13 页。不对就回来改 `outline.json`——大纲里的一页对应成稿的一页，改起来很直接。

---

## 架构说明

```
plugin/
├── package.json          包名 @local/fudan-dppt，版本 0.1.0，入口 index.js
├── cordis.patch.yml      DSH 的插件声明（装上去靠它被识别）
├── index.js              6 个工具在这里注册；出稿、渲染、拼预览的调度也在这
├── figures-tool.js       dppt_figures 的实现：在原文 PDF 上找图块、裁、切
├── review-tools.js       dppt_review 的实现：四道校验
├── icon.svg
├── locale/
│   ├── zh.json           工具描述、报错文案
│   └── en.json
├── assets/
│   ├── templates.json    第一批 8 套的登记表：id、比例、配色、关键词、备注
│   └── refs/*.png        每套一张参考图（封面＋正文页），选模板时看这个
└── engine/
    ├── build_deck.py     主力。按 profile 把大纲排成 13 页
    ├── prepare_template.py  有的模板把设计画在幻灯片上而不是版式里，它负责搬进版式
    ├── build_master.py   母版模式（第二批用；第一批不走这条路，留着）
    ├── extract_figures.py   PDF 抽图：先 probe 看位置，再 box 精裁，grid 分图
    ├── make_outline.py   任意文档 → 大纲 JSON（判类型、选页型、抽句）
    ├── review_deck.py    来源 / 几何 / 排版 / 对比，四条线各查各的
    ├── make_preview.py   逐页 PNG → 一张联览图
    ├── render_deck.ps1   调本机 PowerPoint 导 PNG（COM，比 LibreOffice 准）
    ├── plot_chart.R      R 出图的接口。matplotlib 没装，是故意的
    ├── slides-sciadv.json  自带示例大纲：Sci Adv 那篇 Meta-SIFT
    └── 色卡.json          配色卡。不许现场调色
`

> `templates.json` 只存文件名、比例、配色、关键词和参考图；实际 pptx 从 Releases 按需下载，用 `DPPT_TEMPLATE_DIR` 指过去。

---

## 它是怎么干活的

一条线，从文档到成品，中间每一站都留了人为干预的口子。

**① 读文档 → 大纲**

`make_outline.py` 判断这是论文、会议记录还是普通报告，然后按页型抽句：问题、方法、数字、结论。
产出是 `outline.json`——**这一步的产物值得你亲手改**，因为后面所有排版都只是把这份 JSON 摆上页。

**② 挑模板**

`dppt_templates` 列全部；`dppt_recommend "严肃答辩"` 按文风打分。每套都配了参考图，看图比看关键词快。

**③ 配图**

这是硬规矩里最硬的一条：**能用原文原图，就绝不重绘。**

```
dppt_figures(pdf, probe=true)                     先看这一页有没有图、在哪
dppt_figures(pdf, page=5, box="0.1,0.2,0.9,0.8")  按页面比例裁
dppt_figures(pdf, page=5, grid="2x2")             一张图里分好几个子图
```

**④ 排版**

`build_deck.py` 按 profile 一行行摆：封面、目录、章节、正文、封底。
正文与图只有两种关系——**左右**，或者**上下**。想摆第三种，说明这一页该拆了。

**⑤ 校验**

出完自动跑 `review_deck.py`，四个维度：数字能不能回溯、元素有没有越界重叠、字号在不在档位、深色底上的字够不够白。
默认**只提意见、不拦你**。它见过太多"越改越怪"，所以只在你要它严格时才严格。

---

## 6 个工具

| 工具 | 干什么 | 什么时候用 |
|---|---|---|
| `dppt_templates` | 列模板 | 不知道有什么可选 |
| `dppt_recommend` | 按文风推荐 | 有话说不出该用哪套 |
| `dppt_outline` | 文档转大纲 | 每次开工第一件事 |
| `dppt_figures` | 原文 PDF 裁图 | 有图必须用图 |
| `dppt_build` | 出 pptx（可连渲染） | 主体 |
| `dppt_review` | 四道校验 | 出完、或改完 |

---

## 依赖

**必装**

- Python 3.10+
- `python-pptx` `lxml` `Pillow` `numpy`
- 本机 PowerPoint（导 PNG 走 COM，`render_deck.ps1` 会调它）

**按需**

- `PyMuPDF`：PDF 抽图。没有它 `dppt_figures` 用不了，其余照常
- `Rscript`（R 4.x）：`plot_chart.R` 自绘。不装就跳过，不会报错

**环境变量**

| 变量 | 用途 |
|---|---|
| `DPPT_ROOT` | dataRoot，模板从这里找 |
| `PYTHONIOENCODING=utf-8` | Windows 上不设会乱码，务必设 |

---

## 两条约定

**配色**一律取 `engine/色卡.json`。一套片子里三种红，是最容易露怯的地方。

**字号**在档位表里选：11 / 13 / 16 / 22 / 28 / 32 / 48 / 54。中间值不是不能写，是写了以后整本会显得散。

其他固化的东西（一页一论点、论证超两句就拆、三线表、图注进备注、中文 GB/T 7714-2015 / 中英混排 MLA 9、中文微软雅黑配西文 Times New Roman）都在 `build_deck.py` 和 `review_deck.py` 里，不用记。

---

## 边界，说在前面

- 4:3 那套（`jianyue`）的正文版式没做专门适配，能出但不够漂亮。
- 自绘图表只留了接口，没内置 matplotlib——想清楚再画，比随手画一张强。
- 第二批 6 套换了做法（母版模式），对应关系在 `SECOND-BATCH.md`；模板文件不在本包。棕色那套版面还在调，暂不发布。
- 这套东西的脾气：**大纲不对，排版救不回来。** 时间花在 ①②③，④⑤ 交给它。

---

## 声明

- 本插件只做**排版与流程编排**，不生产、不销售、不授权任何 PPT 模板。
- 模板（含版式、配色、图形、字体搭配）的著作权归各自作者或权利方所有。本插件不主张任何权利，也不对其合法性、授权状态、可商用性作任何保证。
- 主仓库不附带模板文件；模板在配套仓库 `FudanPPT_ForDSH_Templates` 内分发。使用者需自行确认合法使用权，并按原权利方的许可范围使用（个人学习、课程汇报、商业演示的条件可能不同）。
- 由模板使用、二次修改、对外发布引起的权利纠纷，本插件及其作者不承担责任。
- 复旦校名、校徽、校训等标识的权利归复旦大学所有；生成物的对外使用需遵守学校相关规定。
- 代码部分（plugin/）可自由修改自用；若再分发，请保留本声明。
