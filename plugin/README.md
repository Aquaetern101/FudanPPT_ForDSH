# 复旦叙事 PPT 插件（测试版 v0.1）

把大纲 JSON 生成为复旦风格 PPT，并提供"选模板前一瞥"的参考图。

## 工具

| 工具 | 作用 |
|---|---|
| `dppt_templates` | 列模板：id、比例、配色系列、状态、文风关键词、参考图绝对路径（封面 + 正文页并排） |
| `dppt_recommend` | 按文风关键词打分排序，返回推荐模板 + 参考图 |
| `dppt_figures` | 从原文 PDF 裁配图：probe 看图块位置、box 精确裁、grid 分图二次裁切、trim 去白边 |
| `dppt_review` | 四道校验建议：来源 / 几何 / 排版 / 对比；建议模式，不阻断交付 |
| `dppt_outline` | 任意文档（pdf/docx/md/txt）→ 大纲 JSON 骨架 |
| `dppt_build` | 指定模板 → 生成 pptx → 渲染逐页 PNG → 合成预览图 |

## 配图规则（硬性）

1. 原文 PDF 原图裁切（`dppt_figures`）——默认且首选；
2. 官方补充材料；
3. 自绘示意图——仅当前两者都无图，且须标注「示意」。

凡原文有对应图，禁止自绘替代。出稿后跑 `01-analyse\check_figures.py` 校验来源。

## 新模板（--auto）

不认识的新模板：dppt_build 会用 --auto 从模板推断画布、正文版式、内容区与主题强调色，
并按 accent 就近选色卡色系。非 13.333x7.5 的画布按 K=SW/13.333 等比缩放全部版面常数。

## 模板

`meeting`（简约-会议）、`crown`（日冕绿）、`motto`（校训墙红）、`jiangwan`（江湾黑）、
`xianghuitang`（相辉堂蓝）、`laoxiaomen`（老校门紫）、`tpl01`（几何）、`jianyue`（4:3，暂缓）。

参考图放在 `assets/refs/`，清单在 `assets/templates.json`（`keywords` 供文风推荐）。

## 目录

```
index.js              工具注册（Host）
engine/build_deck.py  排版引擎（python-pptx，按 profile 取版式与配色）
engine/slides-sciadv.json  示例大纲
engine/色卡.json      配色单一来源
engine/render_deck.ps1     PowerPoint COM 导出 PNG
engine/make_preview.py     拼三列总览 + 逐页长图
assets/templates.json 模板清单
assets/refs/*.png     模板参考图
```

## 配置

```yaml
- id: fudan-dppt
  name: '@local/fudan-dppt'
  config:
    dataRoot: 'D:\dsh-wx\fudan-ppt'   # 模板与大图所在目录
    outDir: 'D:\dsh-wx\dppt-plugin\out'
    render: true
    pythonPath: 'C:\...\python.exe'
    timeoutMs: 300000
```

## 已验证 / 未验证

- 已验证：引擎 CLI（`--template/--slides/--out`）、会议模板端到端生成、预览图产出。
- 依赖本机 PowerPoint（COM）才能渲染预览；无 PowerPoint 时 `render: false` 仅生成 pptx。
- 模板 pptx 仍从 `dataRoot` 读取，未打包进插件（测试版取舍，正式版做 pack）。

## 已知限制

- 大纲 schema 目前只认示例结构（cover/bars/process/funnel/numbers/figure-split/figure-full/bullets/refs/thanks）。
- `tpl01` 副标题与出处仍会伸进装饰色块，待定方案。
- 4:3（`jianyue`）状态是 `held`。
