/**
 * 复旦叙事 PPT 插件（测试版 v0.1）
 *
 * 提供三个 Agent 工具：
 *   dppt_templates  列出可用模板与其参考图（封面 + 正文页）
 *   dppt_recommend  按文风关键词推荐模板
 *   dppt_build      用指定模板把大纲生成为 pptx，并渲染预览图
 *
 * 生成引擎是同目录 engine/build_deck.py（python-pptx），
 * 渲染走 PowerPoint COM（engine/render_deck.ps1 + make_preview.py）。
 */
import { spawn } from 'node:child_process'
import { access, mkdir, readFile } from 'node:fs/promises'
import { constants } from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

import { registerFiguresTool, FIGURE_SRC_RULE } from './figures-tool.js'
import { registerReviewTool, registerOutlineTool } from './review-tools.js'

const PKG = path.dirname(fileURLToPath(import.meta.url))
const ENGINE = path.join(PKG, 'engine')
const DEFAULT_PYTHON = 'C:\\Users\\1\\.dsh\\dsh-runtimes\\dsh-primary-runtime\\dependencies\\python\\python.exe'
const DEFAULT_ROOT = 'D:\\dsh-wx\\fudan-ppt'
const DEFAULT_PDF_PYTHON = 'C:\\Users\\1\\AppData\\Local\\Programs\\Python\\Python312\\python.exe'

export const inject = ['tools']

function settings(config = {}) {
  return {
    pythonPath: config.pythonPath || DEFAULT_PYTHON,
    pdfPython: config.pdfPython || DEFAULT_PDF_PYTHON,
    dataRoot: config.dataRoot || DEFAULT_ROOT,
    outDir: config.outDir || path.join(PKG, 'out'),
    render: config.render !== false,
    timeoutMs: Number(config.timeoutMs || 300000),
    shell: config.powershell || 'powershell.exe',
  }
}

async function exists(p) {
  try {
    await access(p, constants.F_OK)
    return true
  } catch {
    return false
  }
}

async function loadTemplates(cfg) {
  const raw = await readFile(path.join(PKG, 'assets', 'templates.json'), 'utf8')
  const doc = JSON.parse(raw)
  // 模板目录可外部指定：DPPT_TEMPLATE_DIR 或 dataRoot/00-template。
  // 模板文件不随插件分发，用户指向自己已有的那份即可。
  const tplRoot = process.env.DPPT_TEMPLATE_DIR || path.join(cfg.dataRoot, '00-template')
  return (doc.templates || []).map(t => ({
    ...t,
    ref: path.join(PKG, 'assets', t.ref),
    templatePath: path.isAbsolute(t.file) ? t.file : path.join(tplRoot, t.file),
  }))
}

function run(cmd, args, { cwd, env, timeoutMs }) {
  return new Promise((resolve, reject) => {
    const started = Date.now()
    const child = spawn(cmd, args, {
      cwd,
      env: { ...process.env, ...env },
      windowsHide: true,
    })
    let stdout = ''
    let stderr = ''
    const timer = setTimeout(() => {
      child.kill()
      reject(new Error(`超时 ${timeoutMs}ms：${path.basename(cmd)}`))
    }, timeoutMs)
    child.stdout?.on('data', d => { stdout += d.toString() })
    child.stderr?.on('data', d => { stderr += d.toString() })
    child.on('error', err => { clearTimeout(timer); reject(err) })
    child.on('close', code => {
      clearTimeout(timer)
      resolve({ code, stdout, stderr, elapsedMs: Date.now() - started })
    })
  })
}

function textBlock(value) {
  return [{ type: 'text', text: typeof value === 'string' ? value : JSON.stringify(value, null, 2) }]
}

export function apply(ctx, config = {}) {
  const cfg = settings(config)

  ctx.tools.register({
    name: 'dppt_templates',
    description:
      '列出可用于生成复旦风格 PPT 的模板：id、名称、比例、配色系列、状态、文风关键词，以及每套模板的参考图（封面 + 正文页）绝对路径。关键词参数可按文风过滤。',
    parameters: {
      type: 'object',
      properties: {
        keyword: { type: 'string', description: '按文风关键词过滤，例如「答辩」「会议」「红色」；省略则返回全部' },
      },
      additionalProperties: false,
    },
    output: {
      schema: {
        type: 'object',
        properties: {
          count: { type: 'number' },
          templates: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                id: { type: 'string' },
                name: { type: 'string' },
                ratio: { type: 'string' },
                series: { type: 'string' },
                status: { type: 'string' },
                keywords: { type: 'array', items: { type: 'string' } },
                notes: { type: 'string' },
                ref: { type: 'string' },
                templatePath: { type: 'string' },
              },
              additionalProperties: true,
            },
          },
        },
        additionalProperties: false,
      },
      render: (_args, value) => {
        const list = value?.templates || []
        const lines = list.map(
          t => `- ${t.id}（${t.name}，${t.ratio}，${t.series}）${t.status === 'ready' ? '' : '【' + t.status + '】'}\n  参考图：${t.ref}`,
        )
        return textBlock(`共 ${list.length} 套模板：\n${lines.join('\n')}`)
      },
    },
    async execute(args) {
      const all = await loadTemplates(cfg)
      const kw = String(args?.keyword || '').trim()
      const list = kw
        ? all.filter(t => t.keywords?.some(k => k.includes(kw) || kw.includes(k)) || t.name.includes(kw))
        : all
      return {
        count: list.length,
        templates: list.map(t => ({
          id: t.id,
          name: t.name,
          ratio: t.ratio,
          series: t.series,
          status: t.status,
          keywords: t.keywords,
          notes: t.notes,
          ref: t.ref,
          templatePath: t.templatePath,
        })),
      }
    },
  })

  ctx.tools.register({
    name: 'dppt_recommend',
    description:
      '按用户提供的文风关键词（如「严肃答辩」「校庆典礼」「创业路演」）给模板打分并排序，返回推荐模板及其参考图路径，供在套模板前给用户看。',
    parameters: {
      type: 'object',
      properties: {
        style: { type: 'string', description: '文风描述或关键词，可含多个词，用空格/逗号分隔' },
        limit: { type: 'number', description: '返回条数，默认 3' },
      },
      required: ['style'],
      additionalProperties: false,
    },
    output: {
      schema: {
        type: 'object',
        properties: {
          query: { type: 'string' },
          matches: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                id: { type: 'string' },
                name: { type: 'string' },
                score: { type: 'number' },
                reason: { type: 'string' },
                ref: { type: 'string' },
                status: { type: 'string' },
              },
              additionalProperties: true,
            },
          },
        },
        additionalProperties: false,
      },
      render: (_a, v) =>
        textBlock(
          `文风「${v?.query}」推荐：\n` +
            (v?.matches || []).map(m => `- ${m.id}（${m.name}）score=${m.score}｜${m.reason}\n  参考图：${m.ref}`).join('\n'),
        ),
    },
    async execute(args) {
      const q = String(args?.style || '')
      const tokens = q.split(/[\s,，、;；/]+/).filter(Boolean)
      const limit = Number(args?.limit || 3)
      const all = await loadTemplates(cfg)
      const scored = all
        .map(t => {
          const hits = tokens.filter(
            tok => t.keywords?.some(k => k.includes(tok) || tok.includes(k)) || t.name.includes(tok) || t.series.includes(tok),
          )
          return { t, hits, score: hits.length * 10 + (t.status === 'ready' ? 1 : 0) }
        })
        .filter(x => x.hits.length > 0)
        .sort((a, b) => b.score - a.score)
        .slice(0, limit)
      return {
        query: q,
        matches: scored.map(({ t, hits, score }) => ({
          id: t.id,
          name: t.name,
          score,
          reason: `命中：${hits.join('、')}；${t.notes || ''}`,
          ref: t.ref,
          status: t.status,
        })),
      }
    },
  })

  registerFiguresTool(ctx, cfg, ENGINE, textBlock)
  registerReviewTool(ctx, cfg, ENGINE, textBlock)
  registerOutlineTool(ctx, cfg, ENGINE, textBlock)

  ctx.tools.register({
    name: 'dppt_build',
    description:
      '用指定模板把大纲 JSON 生成为 pptx：调用引擎排版，可选渲染逐页 PNG 并合成预览图。返回 pptx、预览图与引擎日志。',
    parameters: {
      type: 'object',
      properties: {
        template: { type: 'string', description: '模板 id，见 dppt_templates，例如 meeting / crown / tpl01' },
        outline: { type: 'string', description: '大纲 JSON 绝对路径；省略则用插件内置的示例大纲' },
        out: { type: 'string', description: '输出 pptx 绝对路径；省略则写到插件 out 目录' },
        render: { type: 'boolean', description: '是否渲染预览图，默认 true' },
      },
      required: ['template'],
      additionalProperties: false,
    },
    output: {
      schema: {
        type: 'object',
        properties: {
          ok: { type: 'boolean' },
          template: { type: 'string' },
          pptx: { type: 'string' },
          previewSheet: { type: 'string' },
          previewFull: { type: 'string' },
          pngDir: { type: 'string' },
          slides: { type: 'number' },
          elapsedMs: { type: 'number' },
          advice: { type: 'string' },
          log: { type: 'string' },
        },
        additionalProperties: false,
      },
      render: (_a, v) => {
        if (!v?.ok) return textBlock(`生成失败：${v?.log || '未知错误'}`)
        return textBlock(
          [
            `已生成 ${v.slides} 页：${v.pptx}`,
            v.previewSheet ? `预览总览：${v.previewSheet}` : '',
            v.previewFull ? `逐页长图：${v.previewFull}` : '',
            v.pngDir ? `单页 PNG：${v.pngDir}` : '',
            `耗时 ${Math.round((v.elapsedMs || 0) / 1000)}s`,
            v.advice ? `\n${v.advice}` : '',
          ]
            .filter(Boolean)
            .join('\n'),
        )
      },
    },
    async execute(args, exec) {
      const templates = await loadTemplates(cfg)
      const tpl = templates.find(t => t.id === args?.template)
      if (!tpl) throw new Error(`未知模板 id：${args?.template}；可用：${templates.map(t => t.id).join(', ')}`)

      const python = cfg.pythonPath
      if (!(await exists(python))) throw new Error(`找不到 python：${python}`)
      if (!(await exists(tpl.templatePath))) throw new Error(`找不到模板文件：${tpl.templatePath}`)

      await mkdir(cfg.outDir, { recursive: true })
      const stamp = new Date().toISOString().replace(/[-:T]/g, '').slice(0, 12)
      const outPptx = args?.out || path.join(cfg.outDir, `${tpl.id}-${stamp}.pptx`)
      const outline = args?.outline || path.join(ENGINE, 'slides-sciadv.json')

      const masterMode = tpl.profile === 'master'
      const buildArgs = masterMode
        ? [
          path.join(ENGINE, 'build_master.py'),
          '--template', tpl.templatePath,
          '--slides', outline,
          '--out', outPptx,
        ]
        : [
          path.join(ENGINE, 'build_deck.py'),
          tpl.profile,
          '--template', tpl.templatePath,
          '--slides', outline,
          '--out', outPptx,
        ]
      const build = await run(python, buildArgs, {
        cwd: ENGINE,
        env: { DPPT_ROOT: cfg.dataRoot, PYTHONIOENCODING: 'utf-8' },
        timeoutMs: cfg.timeoutMs,
      })
      if (build.code !== 0) {
        return {
          ok: false,
          template: tpl.id,
          pptx: outPptx,
          slides: 0,
          elapsedMs: build.elapsedMs,
          log: `exit ${build.code}\n${build.stdout}\n${build.stderr}`.trim(),
        }
      }
      const slideMatch = /slides (\d+)/.exec(build.stdout)
      let pngDir
      let previewSheet
      let previewFull
      const doRender = args?.render === undefined ? cfg.render : !!args.render
      if (doRender) {
        pngDir = path.join(cfg.outDir, 'png-' + tpl.id)
        await mkdir(pngDir, { recursive: true })
        const ps = await run(
          cfg.shell,
          ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', path.join(ENGINE, 'render_deck.ps1'), '-Pptx', outPptx, '-OutDir', pngDir],
          { cwd: ENGINE, env: {}, timeoutMs: cfg.timeoutMs },
        )
        if (ps.code === 0) {
          const base = path.join(cfg.outDir, `preview-${tpl.id}`)
          const pv = await run(
            python,
            [path.join(ENGINE, 'make_preview.py'), pngDir, base],
            { cwd: ENGINE, env: { PYTHONIOENCODING: 'utf-8' }, timeoutMs: cfg.timeoutMs },
          )
          if (pv.code === 0) {
            const lines = pv.stdout.trim().split(/\r?\n/)
            previewSheet = lines[0]
            previewFull = lines[1]
          }
        }
      }

      let advice = ''
      try {
        const rv = await run(python, [path.join(ENGINE, 'review_deck.py'), outPptx],
          { cwd: ENGINE, env: { PYTHONIOENCODING: 'utf-8' }, timeoutMs: cfg.timeoutMs })
        advice = (rv.stdout || '').trim()
      } catch {
        advice = ''
      }

      return {
        ok: true,
        advice,
        template: tpl.id,
        pptx: outPptx,
        previewSheet: previewSheet || '',
        previewFull: previewFull || '',
        pngDir: pngDir || '',
        slides: slideMatch ? Number(slideMatch[1]) : 0,
        elapsedMs: build.elapsedMs,
        log: `${build.stdout}${build.stderr}`.trim(),
      }
    },
  })
}
