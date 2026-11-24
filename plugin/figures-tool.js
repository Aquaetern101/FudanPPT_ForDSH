/**
 * 追加工具：dppt_figures —— 从原文 PDF 裁配图（含 probe / grid 分图 / trim）。
 *
 * 规范：配图必须优先取自原文 PDF；自绘示意图只能兜底且须标注「示意」。
 * 该工具就是"从 PDF 取图"的那一步，避免模型退回去画图。
 */
import { spawn } from 'node:child_process'
import { access } from 'node:fs/promises'
import { constants } from 'node:fs'
import path from 'node:path'

const PICTURE_SRC_RULE =
  '配图优先级：① 原文 PDF 原图裁切（本工具）② 官方补充材料 ③ 自绘示意图（须在版面或备注标注「示意」）。' +
  '凡原文有对应图，禁止自绘替代。'

export const FIGURE_SRC_RULE = PICTURE_SRC_RULE

async function exists(p) {
  try {
    await access(p, constants.F_OK)
    return true
  } catch {
    return false
  }
}

function exec(cmd, args, timeoutMs) {
  return new Promise((resolve, reject) => {
    const child = spawn(cmd, args, { windowsHide: true })
    let stdout = ''
    let stderr = ''
    const timer = setTimeout(() => {
      child.kill()
      reject(new Error(`超时 ${timeoutMs}ms`))
    }, timeoutMs)
    child.stdout?.on('data', d => { stdout += d.toString() })
    child.stderr?.on('data', d => { stderr += d.toString() })
    child.on('error', e => { clearTimeout(timer); reject(e) })
    child.on('close', code => { clearTimeout(timer); resolve({ code, stdout, stderr }) })
  })
}

/**
 * @param {import('node:child_process')} _spawn 保留签名一致性
 */
export function registerFiguresTool(ctx, cfg, ENGINE, textBlock) {
  ctx.tools.register({
    name: 'dppt_figures',
    description:
      '从原文 PDF 裁切配图。先 probe 看页面上的位图/矢量图块位置，再用 box（页面比例）精确裁；需要分图时用 grid（如 2x2）把一块再切成子图。' +
      PICTURE_SRC_RULE,
    parameters: {
      type: 'object',
      properties: {
        pdf: { type: 'string', description: '原文 PDF 绝对路径' },
        probe: { type: 'boolean', description: 'true 时只探测：列出页面尺寸、大位图与矢量图块的 bbox（比例），不裁切' },
        pages: { type: 'string', description: 'probe 模式下的页号，如 "3-5,7"（1 基）' },
        page: { type: 'number', description: '裁切页号（1 基）' },
        box: { type: 'string', description: '裁切框，页面比例 "x0,y0,x1,y1"（0–1）；省略等于整页' },
        grid: { type: 'string', description: '分图，"RxC" 例如 "2x2"：把 box 区域切成子图，用于二次裁切' },
        auto: { type: 'boolean', description: 'true 时用页面大位图+矢量块自动框出图区（避开正文栏），可与 box 取交集' },
        trim: { type: 'boolean', description: '自动去掉白边（默认 true）' },
        dpi: { type: 'number', description: '渲染分辨率，默认 300；线条图建议 400' },
        outdir: { type: 'string', description: '输出目录，默认 <pdf 同级>/figs' },
        prefix: { type: 'string', description: '文件名前缀，默认 fig' },
      },
      required: ['pdf'],
      additionalProperties: false,
    },
    output: {
      schema: {
        type: 'object',
        properties: {
          ok: { type: 'boolean' },
          mode: { type: 'string' },
          manifest: { type: 'string' },
          entries: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                file: { type: 'string' },
                page: { type: 'number' },
                px: { type: 'array', items: { type: 'number' } },
                aspect: { type: 'number' },
                frac: { type: 'array', items: { type: 'number' } },
              },
              additionalProperties: true,
            },
          },
          probe: { type: 'object' },
          log: { type: 'string' },
        },
        additionalProperties: false,
      },
      render: (_a, v) => {
        if (!v?.ok) return textBlock(`裁切失败：${v?.log || ''}`)
        if (v.mode === 'probe') {
          return textBlock(`PDF ${v.probe?.page_count} 页，探测结果：\n` + JSON.stringify(v.probe?.pages, null, 2))
        }
        return textBlock(
          `已裁 ${v.entries.length} 张：\n` +
            v.entries.map(e => `- ${e.file}  ${e.px[0]}×${e.px[1]}  aspect=${e.aspect}`).join('\n') +
            `\n清单：${v.manifest}`,
        )
      },
    },
    async execute(args) {
      const pdf = String(args?.pdf || '')
      if (!(await exists(pdf))) throw new Error(`找不到 PDF：${pdf}`)
      const python = cfg.pdfPython
      if (!(await exists(python))) {
        throw new Error(
          `找不到带 PyMuPDF 的 python：${python}。请安装 PyMuPDF（pip install pymupdf）或在插件 config 里指定 pdfPython`,
        )
      }
      const script = path.join(ENGINE, 'extract_figures.py')
      const cli = ['-X', 'utf8', script, '--pdf', pdf, '--dpi', String(args?.dpi || 300)]
      const outdir = args?.outdir || path.join(path.dirname(pdf), 'figs')
      if (args?.probe) {
        cli.push('--probe', '--pages', String(args?.pages || args?.page || '1'))
      } else {
        if (!args?.page) throw new Error('裁切需要 page（1 基）')
        cli.push('--page', String(args.page), '--outdir', outdir, '--prefix', String(args?.prefix || 'fig'))
        if (args?.box) cli.push('--box', String(args.box))
        if (args?.grid) cli.push('--grid', String(args.grid))
        if (args?.auto) cli.push('--auto')
        if (args?.trim !== false) cli.push('--trim')
      }
      const r = await exec(python, cli, cfg.timeoutMs)
      if (r.code !== 0) {
        return { ok: false, mode: args?.probe ? 'probe' : 'crop', manifest: '', entries: [], log: (r.stderr || r.stdout).slice(-2000) }
      }
      let data
      try {
        data = JSON.parse(r.stdout)
      } catch {
        return { ok: false, mode: 'crop', manifest: '', entries: [], log: r.stdout.slice(-2000) }
      }
      if (args?.probe) {
        return { ok: true, mode: 'probe', manifest: '', entries: [], probe: data, log: '' }
      }
      return { ok: true, mode: 'crop', manifest: data.manifest, entries: data.entries, log: '' }
    },
  })
}
