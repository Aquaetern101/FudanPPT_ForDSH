/**
 * 追加工具：dppt_review —— 出稿后的四道校验，**建议模式**。
 * 不阻断交付：成品正常给，同时返回下一步修改意见。
 *
 * 四道：来源（配图是否 PDF 裁切）/ 几何（图文只许左右或上下）/ 排版（字号档位、标题字数、重复文本）/ 对比（压色块或整幅底图）
 */
import { spawn } from 'node:child_process'
import { existsSync } from 'node:fs'
import path from 'node:path'

function exec(cmd, args, timeoutMs) {
  return new Promise((resolve, reject) => {
    const child = spawn(cmd, args, {
      windowsHide: true,
      env: { ...process.env, PYTHONIOENCODING: 'utf-8', PYTHONUTF8: '1' },
    })
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

export function registerReviewTool(ctx, cfg, ENGINE, textBlock) {
  ctx.tools.register({
    name: 'dppt_review',
    description:
      '对已生成的 pptx 跑四道校验并给出**下一步修改意见**（来源 / 几何 / 排版 / 对比）。建议模式：不阻断交付，只列问题与修法；strict 时才以非零退出。',
    parameters: {
      type: 'object',
      properties: {
        pptx: { type: 'string', description: '待校验的 pptx 绝对路径' },
        strict: { type: 'boolean', description: 'true 时有问题即返回失败（CI 用），默认 false' },
      },
      required: ['pptx'],
      additionalProperties: false,
    },
    output: {
      schema: {
        type: 'object',
        properties: {
          ok: { type: 'boolean' },
          items: { type: 'number' },
          report: { type: 'string' },
        },
        additionalProperties: false,
      },
      render: (_a, v) => textBlock(v.report || '（无输出）'),
    },
    async execute(args) {
      const pptx = String(args?.pptx || '')
      if (!existsSync(pptx)) throw new Error(`找不到文件：${pptx}`)
      const cli = [path.join(ENGINE, 'review_deck.py'), pptx]
      if (args?.strict) cli.push('--strict')
      const r = await exec(cfg.pythonPath, cli, cfg.timeoutMs)
      const report = (r.stdout || '').trim()
      const items = (report.match(/^\s*\d+\.\s\[/gm) || []).length
      return { ok: r.code === 0, items, report: report + (r.stderr ? `\n${r.stderr.slice(-500)}` : '') }
    },
  })
}

/**
 * 追加工具：dppt_outline —— 任意文档（pdf/docx/md/txt）→ 大纲 JSON 骨架。
 */
export function registerOutlineTool(ctx, cfg, ENGINE, textBlock) {
  ctx.tools.register({
    name: 'dppt_outline',
    description:
      '把任意文档转成 PPT 大纲 JSON（骨架）：自动判文档类型（论文 / 会议记录 / 通用），选页型并抽句。产出仍需人审，但保证"任何文档都能一路走到 pptx"。',
    parameters: {
      type: 'object',
      properties: {
        input: { type: 'string', description: 'pdf / docx / md / txt 绝对路径' },
        out: { type: 'string', description: '输出 JSON 路径；省略则写到文档同目录' },
      },
      required: ['input'],
      additionalProperties: false,
    },
    output: {
      schema: {
        type: 'object',
        properties: {
          ok: { type: 'boolean' },
          outline: { type: 'string' },
          slides: { type: 'number' },
          log: { type: 'string' },
        },
        additionalProperties: false,
      },
      render: (_a, v) =>
        textBlock(
          v.ok
            ? `大纲已生成：${v.outline}\n页数：${v.slides}\n下一步：dppt_build 用某套模板跑它`
            : `生成失败：${v.log}`,
        ),
    },
    async execute(args) {
      const input = String(args?.input || '')
      if (!existsSync(input)) throw new Error(`找不到输入：${input}`)
      const cli = [path.join(ENGINE, 'make_outline.py'), input]
      if (args?.out) cli.push('--out', String(args.out))
      const r = await exec(cfg.pdfPython, cli, cfg.timeoutMs)
      if (r.code !== 0) return { ok: false, outline: '', slides: 0, log: (r.stderr || r.stdout).slice(-800) }
      try {
        const data = JSON.parse(r.stdout)
        return { ok: true, outline: data.outline, slides: data.slides, log: '' }
      } catch {
        return { ok: false, outline: '', slides: 0, log: r.stdout.slice(-800) }
      }
    },
  })
}
