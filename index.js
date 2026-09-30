/**
 * dsh-export-office-noshell —— 生成 Word / Excel / PowerPoint / PDF / 图片的工具集
 *
 * 五个工具：`export_docx_noshell` / `export_xlsx_noshell` / `export_pptx_noshell` / `export_pdf_noshell` / `export_image_noshell`。
 * 生成逻辑跑在插件进程里（Node，零依赖），调用方只提交数据 ——
 * 不需要 shell、不需要 Python、不需要装 Office。
 * 例外是 PDF 与图片：前者靠本机 LibreOffice 转一道，后者靠本机无头 Chrome / Edge
 * 截图（HTML → 图片），两者都由插件自己拉子进程，模型手上依然没有 shell。
 *
 * 归属：本插件只往 host 的 `tools` 注册表注册工具，不发布任何服务，
 * 因此不需要 isolate realm，也不改动出厂的 host / preset 文件。
 * 装在 profile 层，所以这份能力对 profile 内每个会话都可见。
 *
 * 边界：只往「调用方会话工作区」的 `.ai-output/` 下写这些文件。调用方给的是**文件名**，
 * 给不了路径 —— 目录由插件配置决定，越界一律拒绝。
 *
 * @module dsh-export-office-noshell
 */

import { copyFileSync, existsSync, mkdirSync, readFileSync, renameSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { isAbsolute, join, relative, resolve } from 'node:path'

import { buildDocx } from './lib/docx.mjs'
import { cleanScratch, convertToPdf, findSoffice, makeScratch } from './lib/pdf.mjs'
import { buildPptx } from './lib/pptx.mjs'
import { clampInt, findChrome, shootHtml } from './lib/shot.mjs'
import { MAX_PARTS, clearStage, readParts, savePart, stageDir, stagingState } from './lib/staging.mjs'
import { buildXlsx } from './lib/xlsx.mjs'
import { imageSpec } from './lib/zip.mjs'

/* ────────────────────────────── 配置 ────────────────────────────── */

const DEFAULTS = {
  /** 兜底工作区根目录：仅在调用方会话拿不到 cwd 时使用；能拿到就以会话工作区为准。 */
  root: '',
  /** 输出目录（相对 root）：交付文件固定落在这里。 */
  outputDir: '.ai-output',
  /** 单个文件的单元格 / 段落 / 项目符号总数上限，防止一次塞爆。 */
  maxItems: 20000,
  /** LibreOffice 的 soffice 可执行文件路径；留空则自动探测常见安装位置。 */
  soffice: '',
  /** 一次 PDF 转换的超时（毫秒）。 */
  convertTimeoutMs: 120000,
  /** Chrome / Edge 可执行文件路径；留空则自动探测（也认 CHROME_PATH 环境变量）。 */
  chrome: '',
  /** 一次 HTML 截图的超时（毫秒）。 */
  shotTimeoutMs: 60000,
}

/**
 * Cordis 只认 Standard Schema：加载时会调用 `Config['~standard'].validate(raw)`。
 * 普通「字段→默认值」的 map 会在加载时报 `Cannot read properties of undefined (reading 'validate')`。
 */
function normalizeConfig(raw) {
  const input = raw !== null && typeof raw === 'object' && !Array.isArray(raw) ? raw : {}
  const issues = []
  const value = { ...DEFAULTS }

  // root 降级为兜底项：输出根目录优先取调用方会话的工作区，所以不再强制配置。
  const root = typeof input.root === 'string' ? input.root.trim() : ''
  if (root !== '') value.root = resolve(root)

  if (input.maxItems !== undefined && input.maxItems !== null) {
    const number = Number(input.maxItems)
    if (!Number.isFinite(number) || number <= 0) issues.push({ message: 'maxItems 必须是正数', path: ['maxItems'] })
    else value.maxItems = Math.floor(number)
  }

  if (typeof input.outputDir === 'string' && input.outputDir.trim() !== '') {
    const cleaned = input.outputDir.trim().replace(/\\/g, '/').replace(/^\.\//, '').replace(/^\/+/, '').replace(/\/+$/, '')
    if (cleaned === '' || cleaned.includes('..')) issues.push({ message: 'outputDir 必须是工作区内的相对目录', path: ['outputDir'] })
    else value.outputDir = cleaned
  }

  if (typeof input.soffice === 'string' && input.soffice.trim() !== '') value.soffice = input.soffice.trim()

  if (typeof input.chrome === 'string' && input.chrome.trim() !== '') value.chrome = input.chrome.trim()

  if (input.convertTimeoutMs !== undefined && input.convertTimeoutMs !== null) {
    const number = Number(input.convertTimeoutMs)
    if (!Number.isFinite(number) || number <= 0) issues.push({ message: 'convertTimeoutMs 必须是正数', path: ['convertTimeoutMs'] })
    else value.convertTimeoutMs = Math.floor(number)
  }

  if (input.shotTimeoutMs !== undefined && input.shotTimeoutMs !== null) {
    const number = Number(input.shotTimeoutMs)
    if (!Number.isFinite(number) || number <= 0) issues.push({ message: 'shotTimeoutMs 必须是正数', path: ['shotTimeoutMs'] })
    else value.shotTimeoutMs = Math.floor(number)
  }

  return issues.length > 0 ? { issues } : { value }
}

/** 插件配置（Standard Schema 形态）。 */
export const Config = {
  '~standard': {
    version: 1,
    vendor: 'dsh-export-office-noshell',
    validate: normalizeConfig,
  },
}

/** 只依赖工具注册表。 */
export const inject = ['tools']

/* ─────────────────────────── 公共零件 ─────────────────────────── */

/** 模块级配置，apply 时赋值。 */
let activeConfig = { ...DEFAULTS }

/**
 * 把作者写法的参数表编译成注册表要的形状。
 *
 * `ToolRuntime.register` 会对 `parameters` 直接跑 `assertSupportedJsonSchema`，
 * 只认原始 JSON Schema 子集：属性节点上**不能**写 `required: true`，
 * 必填只能落在对象根的 `required` 数组里。
 */
function compileParameters(spec) {
  const properties = {}
  const required = []
  for (const [key, raw] of Object.entries(spec ?? {})) {
    const node = { ...raw }
    if (node.required === true) {
      required.push(key)
      delete node.required
    }
    properties[key] = node
  }
  return { type: 'object', properties, ...(required.length > 0 ? { required } : {}) }
}

const CELLS = { type: 'array', items: { type: 'string' } }

/** Excel 里的一个原生图表（Excel 打开后可改类型、换数据源、跟随公式重算）。 */
const XLSX_CHART = {
  type: 'object',
  properties: {
    type: { type: 'string', enum: ['bar', 'line', 'pie'], description: '图表类型：柱状图 / 折线图 / 饼图，默认 bar' },
    title: { type: 'string', description: '图表标题，可省略' },
    categories: { type: 'string', description: '当分类轴的列名（表头文字），默认第一列' },
    series: { ...CELLS, description: '要画成系列的列名列表，默认「除分类列外的所有列」' },
    anchor: { type: 'string', description: '图表左上角落在哪个单元格，默认 A2' },
    width: { type: 'number', description: '图表宽度（像素），默认 720' },
    height: { type: 'number', description: '图表高度（像素），默认 420' },
  },
}

/** Excel 里的一个图片（落进该表的 drawing + media）。 */
const XLSX_IMAGE = {
  type: 'object',
  properties: {
    path: { type: 'string', description: '.ai-output/ 下已有的图片文件名（.png/.jpg/.jpeg/.gif）' },
    data: { type: 'string', description: '图片内容的 base64（可带 data: 前缀），与 path 二选一' },
    anchor: { type: 'string', description: '左上角落在哪个单元格，默认 A1，例如 "F2"' },
    width: { type: 'number', description: '显示宽度（像素），默认 420' },
    height: { type: 'number', description: '显示高度（像素），默认 300' },
  },
}

/**
 * 校验文件名并算出落盘路径。
 * @returns `{ ok: true, dir, target }` 或 `{ ok: false, error }`。
 */
/**
 * 决定输出根目录：优先用调用方会话的工作区，其次退回插件配置的 root。
 * profile 内的会话共用一个插件实例，硬编码 root 会把所有会话的产出塞进同一个工作区。
 */
function outputRoot(exec) {
  const cwd = exec?.agent?.session?.header?.cwd
  if (typeof cwd === 'string' && cwd.trim() !== '') return resolve(cwd)
  return activeConfig.root
}

/** 取会话标识：优先会话 id，退回工作区路径 —— 暂存目录靠它把不同会话隔开。 */
function sessionKey(exec) {
  const header = exec?.agent?.session?.header
  const id = header?.id
  return typeof id === 'string' && id.trim() !== '' ? id : String(header?.cwd ?? '')
}

function resolveTarget(fileName, extension, exec) {
  const root = outputRoot(exec)
  if (root === '') return { ok: false, error: '拿不到调用方会话的工作区，且插件未配置兜底 root，无法决定写到哪。' }

  const name = String(fileName ?? '').trim()
  const pattern = new RegExp(`^[^\\\\/:*?"<>|]+\\.${extension}$`, 'i')
  if (!pattern.test(name)) {
    return { ok: false, error: `fileName 必须是「文件名.${extension}」，且不含路径分隔符；收到「${name}」。` }
  }

  const dir = join(root, activeConfig.outputDir)
  const target = join(dir, name)
  const escape = relative(dir, target)
  if (escape.startsWith('..') || isAbsolute(escape)) return { ok: false, error: '目标路径越界，已拒绝。' }

  return { ok: true, dir, target, name }
}

/**
 * 把文件搬到目标位置。
 *
 * Windows 上 `fs.renameSync` 不能跨盘：临时目录（TEMP 通常在 C:）与工作区
 * （可能在 E:）不同卷时会抛 `EXDEV: cross-device link not permitted`。
 * 所以 rename 失败退回复制 + 删源文件，语义与 rename 一致。
 */
function moveFile(from, to) {
  try {
    renameSync(from, to)
  } catch (error) {
    if (error?.code !== 'EXDEV') throw error
    copyFileSync(from, to)
    rmSync(from, { force: true })
  }
}

/** 真正写盘，并把异常翻译成结果对象。 */
function writeOut(resolved, bytes) {
  try {
    mkdirSync(resolved.dir, { recursive: true })
    writeFileSync(resolved.target, bytes)
  } catch (error) {
    return { ok: false, error: `写文件失败：${error instanceof Error ? error.message : String(error)}` }
  }
  return {
    ok: true,
    path: resolved.target,
    relativePath: `${activeConfig.outputDir}/${resolved.name}`,
    bytes: statSync(resolved.target).size,
  }
}

/**
 * 收集并校验一组图片项，读成可嵌入的字节。
 *
 * docx 的 blocks 与 xlsx 每表的 images 都是同一套字段（`path` / `data` / 尺寸），
 * 所以共用这一个收集器：传 `'block'` 模式时只挑 `type === 'image'` 的块并回填
 * 块下标映射，传 `'entry'` 模式时整组都是图片。
 *
 * 图片有两种给法：`path`（.ai-output/ 下已有的图片文件名）或 `data`（base64，
 * 可带 data: 前缀）。两者都只认 png / jpg / jpeg / gif，且以**实际字节**判断
 * 真实格式（不看扩展名）；尺寸从文件头读，读不出来就拒绝 —— 读不出尺寸就没法
 * 按比例排版。
 *
 * 注意：**绝不能往入参对象上写字段**。运行时在派发前对参数做了 deepFreeze
 * （packages/core/tools 的 `arguments: deepFreeze(detached)`），冻结对象加属性
 * 在严格模式（ESM 天然是）直接抛 `Cannot add property …, object is not extensible`。
 * 图片与块的对应关系改为按块下标记录在 `mediaIndex` 数组里。
 *
 * @returns `{ ok: true, media, mediaIndex }` 或 `{ ok: false, error }`。
 *   `mediaIndex[i]` 是第 i 个块对应的 media 下标，非图片块为 `-1`。
 */
function collectMedia(items, exec, mode = 'block') {
  const list = Array.isArray(items) ? items : []
  const media = []
  const mediaIndex = list.map(() => -1)

  for (const [position, item] of list.entries()) {
    if (mode === 'block' && item?.type !== 'image') continue

    const index = media.length
    const ordinal = index + 1
    const path = String(item?.path ?? '').trim()
    const inline = String(item?.data ?? '').trim()
    let data
    let label

    if (inline !== '') {
      const base64 = inline.replace(/^data:image\/[a-z+]+;base64,/i, '')
      try {
        data = Buffer.from(base64, 'base64')
      } catch {
        return { ok: false, error: `第 ${ordinal} 张图片的 data 不是合法 base64。` }
      }
      label = 'data'
    } else if (path !== '') {
      if (!/^[^\\/:*?"<>|]+\.(png|jpe?g|gif)$/i.test(path)) {
        return { ok: false, error: `图片 path 只能是 .ai-output/ 下的 .png/.jpg/.jpeg/.gif 文件名；收到「${path}」。` }
      }
      const root = outputRoot(exec)
      if (root === '') return { ok: false, error: '拿不到调用方会话的工作区，无法读 .ai-output/ 下的图片。' }
      const full = join(root, activeConfig.outputDir, path)
      if (!existsSync(full)) return { ok: false, error: `.ai-output/ 下没有这张图片：${path}` }
      data = readFileSync(full)
      label = path
    } else {
      return { ok: false, error: `第 ${ordinal} 张图片要给 path（.ai-output/ 下已有图片）或 data（base64），二选一。` }
    }

    const spec = imageSpec(data)
    if (!spec) return { ok: false, error: `第 ${ordinal} 张图片（${label}）不是可识别的 png / jpeg / gif，或尺寸读不出来。` }

    mediaIndex[position] = index
    // 多表共用 xlsx 的 xl/media/ 命名空间，所以带上工作表序号避免重名。
    media.push({ name: `image${ordinal}.${spec.ext}`, data, spec })
  }

  return { ok: true, media, mediaIndex }
}

/**
 * 交付说明。
 *
 * `present` 走的是会话文件系统，只让 GUI / 客户端看到文件；IM 渠道
 * （飞书 / 企微 / 钉钉…）看不到，必须再用 dsh-im 插件提供的
 * `dsh_im_return_file` 把同一个文件登记为「本轮结束后投递」。
 * 该工具是否在册取决于会话，所以两步都写清楚，由模型按需执行。
 */
const DELIVERY_NOTE = '交付：调用 present 让 GUI / 客户端可见；若工具表里还有 dsh_im_return_file（IM 渠道会话），必须对同一个文件的绝对路径再调它登记投递，否则 IM 端收不到文件。'

/**
 * 去掉对象里值为 `undefined` 的键（递归）。
 *
 * 运行时会先把工具返回值做一次 lossless-JSON 快照，`undefined` 不在合法取值里，
 * 一个多余的 `undefined` 就会让整次调用以「value is not lossless JSON」失败 ——
 * 哪怕文件已经写好。这里兜一道，保证返回出去的值永远是可 JSON 化的。
 */
function dropUndefined(value) {
  if (Array.isArray(value)) return value.map(dropUndefined)
  if (value === null || typeof value !== 'object') return value
  const out = {}
  for (const [key, item] of Object.entries(value)) {
    if (item !== undefined) out[key] = dropUndefined(item)
  }
  return out
}

/** 统一的输出渲染。 */
const OUTPUT = {
  schema: { type: 'object', additionalProperties: true },
  render: (_args, value) => {
    const body = JSON.stringify(dropUndefined(value), null, 2)
    if (!value || value.ok !== true) return [{ type: 'text', text: body }]
    // 分批的中间批次与「只 reset」都只动暂存、没有文件可交付，别去催 present。
    if (value.saved === true) return [{ type: 'text', text: `${body}\n\n这一批已存下，还没生成文件：继续传下一批（part 递增），最后一批带 final: true 一次成文。` }]
    if (value.cleared === true) return [{ type: 'text', text: `${body}\n\n暂存已清空，这篇文档要重新从 part: 1 传起。` }]
    return [{ type: 'text', text: `${body}\n\n交付提醒：present 只覆盖 GUI；若工具表里有 dsh_im_return_file，请对上面这个路径再调一次，否则 IM 渠道收不到文件。` }]
  },
}

/* ─────────────────────────── 工具定义 ─────────────────────────── */

const DOCX_BLOCK = {
  type: 'object',
  properties: {
    type: { type: 'string', enum: ['heading', 'paragraph', 'bullet', 'table', 'image'], description: 'block 类型' },
    text: { type: 'string', description: 'heading / paragraph / bullet 的正文' },
    level: { type: 'number', description: 'heading 级别：1 / 2 / 3' },
    columns: { ...CELLS, description: 'table 的表头' },
    rows: { type: 'array', items: CELLS, description: 'table 的数据行，每行是字符串数组' },
    path: { type: 'string', description: 'image 专用：.ai-output/ 下已有的图片文件名（.png/.jpg/.jpeg/.gif），例如 MOM.png' },
    data: { type: 'string', description: 'image 专用：图片内容的 base64（可带 data:image/png;base64, 前缀），与 path 二选一' },
    width: { type: 'number', description: 'image 专用：显示宽度（像素），默认按版心 90% 排' },
    widthPct: { type: 'number', description: 'image 专用：显示宽度占版心百分比（1-100）；给了 width 就以 width 为准' },
  },
  required: ['type'],
}

const xlsxTool = {
  name: 'export_xlsx_noshell',
  description: [
    '生成真正的 Excel 工作簿（.xlsx），写到工作区的 .ai-output/ 目录。',
    DELIVERY_NOTE,
    '单元格值一律用字符串传：要数字就把列名填进 numberColumns；要公式就写 "=SUM(E2:E4)"，插件会写成 Excel 公式（打开时自动重算）。',
    '每张表还能放 images（图片，path 指 .ai-output/ 下已有图或 data 给 base64，可给 anchor 定位）与 charts（原生图表，type 可选 bar/line/pie，categories 指分类列名、series 指要画的列名，Excel 打开后可改类型换数据源）。',
    '只支持 .xlsx；fileName 只能是文件名，不能带路径。',
  ].join(' '),
  parameters: compileParameters({
    fileName: { type: 'string', required: true, description: '不含路径的文件名，例如 customer-list_20260925.xlsx' },
    sheets: {
      type: 'array',
      required: true,
      description: '一个或多个工作表',
      items: {
        type: 'object',
        properties: {
          name: { type: 'string', description: '工作表名，默认 Sheet1，Excel 上限 31 字' },
          columns: { ...CELLS, description: '表头，按列顺序，也决定列数' },
          rows: { type: 'array', items: CELLS, description: '数据行，每行是字符串数组' },
          numberColumns: { ...CELLS, description: '需要按数字写入的列名（可用列名或列字母），否则数字会变成文本' },
          images: { type: 'array', items: XLSX_IMAGE, description: '本工作表要嵌入的图片' },
          charts: { type: 'array', items: XLSX_CHART, description: '本工作表要插入的原生图表；数据取自本表（第一行为表头）' },
        },
        required: ['columns', 'rows'],
      },
    },
  }),
  output: OUTPUT,
  async execute(args, exec) {
    const sheets = Array.isArray(args.sheets) ? args.sheets : []
    if (sheets.length === 0) return { ok: false, error: 'sheets 不能为空。' }
    const count = sheets.reduce((total, sheet) => total + (sheet?.rows?.length ?? 0), 0)
    if (count > activeConfig.maxItems) return { ok: false, error: `数据共 ${count} 行，超过上限 ${activeConfig.maxItems} 行。` }

    const resolved = resolveTarget(args.fileName, 'xlsx', exec)
    if (!resolved.ok) return resolved

    // 逐表把图片读成可嵌入的字节；图表不需要额外读取（数据就在表里）。
    const media = []
    const prepared = []
    let imageCount = 0
    let chartCount = 0
    for (const [index, sheet] of sheets.entries()) {
      const collected = collectMedia(sheet?.images, exec, 'entry')
      if (!collected.ok) return { ok: false, error: `工作表「${sheet?.name ?? `Sheet${index + 1}`}」：${collected.error}` }
      const offset = media.length
      const images = collected.media.map((item, i) => {
        // xlsx 的 xl/media/ 是整个工作簿共享的，重命名避免多表重名。
        const source = Array.isArray(sheet.images) ? sheet.images[i] ?? {} : {}
        return {
          ...item,
          name: `image${offset + i + 1}.${item.spec.ext}`,
          anchor: source.anchor,
          width: source.width,
          height: source.height,
        }
      })
      media.push(...collected.media)
      imageCount += images.length
      chartCount += Array.isArray(sheet?.charts) ? sheet.charts.length : 0
      prepared.push({ ...sheet, images })
    }

    let bytes
    try {
      bytes = buildXlsx({ sheets: prepared })
    } catch (error) {
      return { ok: false, error: `生成失败：${error instanceof Error ? error.message : String(error)}` }
    }
    const written = writeOut(resolved, bytes)
    if (!written.ok) return written
    return {
      ...written,
      sheets: sheets.map((sheet) => String(sheet?.name ?? 'Sheet1')),
      ...(imageCount === 0 ? {} : { images: imageCount }),
      ...(chartCount === 0 ? {} : { charts: chartCount }),
    }
  },
}

/**
 * docx 生成三步：读图 → 拼包 → 写盘。所有批次最终都汇到这一段。
 * @returns 与工具返回值同形的对象（`ok: true` 时含路径、字节数、块数）。
 */
function buildDocxFile(resolved, { title, blocks }, exec) {
  const collected = collectMedia(blocks, exec)
  if (!collected.ok) return collected

  let bytes
  try {
    bytes = buildDocx({ title, blocks, media: collected.media, mediaIndex: collected.mediaIndex })
  } catch (error) {
    return { ok: false, error: `生成失败：${error instanceof Error ? error.message : String(error)}` }
  }

  const written = writeOut(resolved, bytes)
  if (!written.ok) return written
  return { ...written, blocks: blocks.length, ...(collected.media.length === 0 ? {} : { images: collected.media.length }) }
}

/**
 * 一批的推荐上限：块数 / 参数 JSON 字符数。
 *
 * 实测工具参数 JSON 超过 1.6 万字符就开始写坏（16K 以下 0 失败，16–24K 约 7% 崩），
 * 这里取一半留余量：20 个块的自然长度约 1000–2500 字，参数约 5–8K 字符。
 * 只按块数下指令 —— 模型数字数不准，边写边数块数是它做得到的。
 */
const DOCX_BATCH = { blocks: 20, chars: 8000 }

/** 超上限不拦（内容已经写出来了），只回一句 warning 让下一批写小些。 */
function batchWarning(blocks) {
  const length = JSON.stringify(blocks).length
  if (blocks.length <= DOCX_BATCH.blocks && length <= DOCX_BATCH.chars) return undefined
  return `本批 ${blocks.length} 块 / ${length} 字符，超过建议上限（${DOCX_BATCH.blocks} 块 / ${DOCX_BATCH.chars} 字符）：内容已收下，下一批请写小一些。`
}

const docxTool = {
  name: 'export_docx_noshell',
  description: [
    '生成真正的 Word 文档（.docx），写到工作区的 .ai-output/ 目录。',
    DELIVERY_NOTE,
    'blocks 按顺序排：heading（小标题，level 1-3）、paragraph（正文）、bullet（项目符号）、table（表格，需要 columns + rows）、image（图片，path 指 .ai-output/ 下已有的 png/jpg/gif，或 data 给 base64）。',
    '分批写：每次调用交一批 —— fileName + part（批次号，从 1 开始）+ blocks，先落暂存、不出文件；最后一批再加 final: true，插件按批次号顺序合并全部批次，一次性生成完整 docx。小文档也是一次调用：part: 1 + final: true。',
    '一批写 15-20 个块（一段 / 一条 / 一张表各算 1 块，正文合计 1000-2500 字）就停下发出来，剩下的下一次接着写；超过 20 块或 8000 字符时插件会回一句 warning。',
    '同一个 part 重传会覆盖，不会重复；漏了号就补那一批再 final，已存批次不会丢。收尾也可以只给 final: true（不带 blocks），重复调用是安全的。想推倒重写就给 reset: true 清空这篇文档的暂存。',
    '只支持 .docx；fileName 只能是文件名，不能带路径。',
  ].join(' '),
  parameters: compileParameters({
    fileName: { type: 'string', required: true, description: '不含路径的文件名，例如 客户清单_20260925.docx' },
    title: { type: 'string', description: '文档大标题，可省略（分批时只在 final 那批生效，取最后一次给的值）' },
    blocks: { type: 'array', items: DOCX_BLOCK, description: '本批内容块，按顺序排列；一批 15-20 个块（一段 / 一条 / 一张表各算 1 块）' },
    part: { type: 'number', description: '批次号，从 1 开始的整数；带 blocks 就必须给，同一个号重传会覆盖' },
    final: { type: 'boolean', description: 'true = 收尾：把已存的全部批次合并成一篇 docx；可以不带 blocks' },
    reset: { type: 'boolean', description: 'true = 先清空这篇文档已存的批次，再处理本批（推倒重写时用）' },
  }),
  output: OUTPUT,
  async execute(args, exec) {
    const resolved = resolveTarget(args.fileName, 'docx', exec)
    if (!resolved.ok) return resolved

    const blocks = Array.isArray(args.blocks) ? args.blocks : []
    const final = args.final === true
    const reset = args.reset === true
    const raw = args.part
    const part = raw === undefined || raw === null || raw === '' ? undefined : Number(raw)
    if (part !== undefined && (!Number.isInteger(part) || part < 1 || part > 9999)) {
      return { ok: false, error: `part 必须是 1..9999 的整数（批次号，从 1 开始）；收到「${raw}」。` }
    }
    if (blocks.length === 0 && !final && !reset) {
      return { ok: false, error: '写内容要给 fileName + part（批次号，从 1 开始）+ blocks；收尾就再带 final: true。' }
    }

    /* 只有一条路：这次调用就是一批。先落暂存，再决定要不要收尾。 */
    const session = sessionKey(exec)
    const dir = stageDir(outputRoot(exec), activeConfig.outputDir, session, resolved.name)
    if (reset) clearStage(dir)

    let saved = part
    let warning
    if (blocks.length > 0) {
      // 每批都必须报批次号：同一个号落到同一个文件上，重传才幂等。
      // 不做「自动编号」—— 重试时它会给出同号之外的新号，正好造出重复内容。
      if (saved === undefined) {
        return { ok: false, error: '每一批都要给 part（批次号，从 1 开始）：同一个号重传会覆盖，才不会写出两份。' }
      }
      // 只用「目录 + index.json」，不读历史批次正文 —— 读正文是 O(n²)，200 批能拖到几十秒。
      const before = stagingState(dir)
      if (!before.parts.includes(saved) && before.parts.length >= MAX_PARTS) {
        return { ok: false, error: `这篇文档已存 ${before.parts.length} 批，达到上限 ${MAX_PARTS}；请先 final 成文，或 reset 清空重来。` }
      }
      // 上限按「留下的批次 + 本批」算：同号重传时要把被替换那批排除掉。
      const kept = before.total - (before.blocks[saved] ?? 0)
      if (kept + blocks.length > activeConfig.maxItems) {
        return { ok: false, error: `已存批次加本批共 ${kept + blocks.length} 个块，超过上限 ${activeConfig.maxItems}。` }
      }
      savePart(dir, saved, { title: args.title, blocks, meta: { sessionId: session, fileName: resolved.name } })
      warning = batchWarning(blocks)
    }

    const state = stagingState(dir)

    if (!final) {
      if (blocks.length === 0) {
        // 只给了 reset：那件事（清空这篇文档的暂存）已经做完，没有别的可做。
        if (reset) return { ok: true, cleared: true, fileName: resolved.name, parts: state.parts, blocks: 0 }
        return { ok: false, error: '写内容要给 blocks + part；收尾就给 final: true。' }
      }
      return {
        ok: true,
        saved: true,
        fileName: resolved.name,
        part: saved,
        parts: state.parts,
        blocks: state.total,
        nextPart: Math.max(0, ...state.parts) + 1,
        ...(warning === undefined ? {} : { warning }),
      }
    }

    /* 收尾：到这里才把每批正文真正读回来合并。批次必须连续，缺号就停在这里，暂存原样保留。 */
    const snapshot = readParts(dir)
    if (snapshot.error) return { ok: false, error: snapshot.error }
    const numbers = snapshot.parts.map((item) => item.part)

    if (snapshot.parts.length === 0) {
      // 收尾重试：暂存已经清掉、文件已经在了 —— 把既有产出原样报回去，
      // 并写明「这次没重新生成」，免得模型以为又出了一版。
      if (existsSync(resolved.target)) {
        return {
          ok: true,
          already: true,
          path: resolved.target,
          relativePath: `${activeConfig.outputDir}/${resolved.name}`,
          bytes: statSync(resolved.target).size,
          note: '暂存里已没有待合并的批次（上次 final 已成文并清理了暂存）：本次没有重新生成，文件就是上次的产出。要重写请 reset: true 后从 part: 1 重传。',
        }
      }
      return { ok: false, error: '还没有任何已存批次，无法收尾成文；先传 part: 1 + blocks。' }
    }
    const missing = []
    for (let index = 1; index <= Math.max(...numbers); index += 1) if (!numbers.includes(index)) missing.push(index)
    if (missing.length > 0) {
      return { ok: false, error: `批次不连续，缺第 ${missing.join(' / ')} 批；补齐后再传 final: true（已存批次保留）。` }
    }

    const merged = snapshot.parts.flatMap((item) => item.blocks)
    if (merged.length === 0) return { ok: false, error: '已存批次里没有内容块。' }
    if (merged.length > activeConfig.maxItems) return { ok: false, error: `合并后 ${merged.length} 个块，超过上限 ${activeConfig.maxItems}。` }

    const title = String(args.title ?? '').trim() || snapshot.parts.find((item) => item.title.trim() !== '')?.title
    const result = buildDocxFile(resolved, { title, blocks: merged }, exec)
    if (result.ok !== true) return result
    clearStage(dir) // 成文成功才删暂存（part 文件 + index.json + meta.json 一次清掉）；写盘失败保留，final 可直接重试。
    return { ...result, parts: snapshot.parts.length, ...(warning === undefined ? {} : { warning }) }
  },
}

const pptxTool = {
  name: 'export_pptx_noshell',
  description: [
    '生成真正的 PowerPoint 演示文稿（.pptx），写到工作区的 .ai-output/ 目录。',
    DELIVERY_NOTE,
    'slides 每页给一个 title 和 bullets 项目符号列表；bullets 里每项可以是字符串，也可以写 {"text":"...","level":1} 表示缩进层级。',
    '只支持 .pptx；fileName 只能是文件名，不能带路径。',
  ].join(' '),
  parameters: compileParameters({
    fileName: { type: 'string', required: true, description: '不含路径的文件名，例如 汇报_20260925.pptx' },
    slides: {
      type: 'array',
      required: true,
      description: '演示文稿的每一页',
      items: {
        type: 'object',
        properties: {
          title: { type: 'string', description: '本页标题' },
          bullets: {
            type: 'array',
            description: '本页项目符号；每项可以是字符串，也可以是 {"text":"…","level":1} 表示缩进层级',
            items: {
              oneOf: [
                { type: 'string' },
                {
                  type: 'object',
                  properties: {
                    text: { type: 'string', description: '这一条的文字' },
                    level: { type: 'number', description: '缩进层级，0 起' },
                  },
                  required: ['text'],
                },
              ],
            },
          },
        },
        required: ['title'],
      },
    },
  }),
  output: OUTPUT,
  async execute(args, exec) {
    const slides = Array.isArray(args.slides) ? args.slides : []
    if (slides.length === 0) return { ok: false, error: 'slides 不能为空。' }
    const count = slides.reduce((total, slide) => total + (slide?.bullets?.length ?? 0), 0)
    if (count > activeConfig.maxItems) return { ok: false, error: `项目符号共 ${count} 条，超过上限 ${activeConfig.maxItems}。` }

    const resolved = resolveTarget(args.fileName, 'pptx', exec)
    if (!resolved.ok) return resolved

    let bytes
    try {
      bytes = buildPptx({ slides })
    } catch (error) {
      return { ok: false, error: `生成失败：${error instanceof Error ? error.message : String(error)}` }
    }
    const written = writeOut(resolved, bytes)
    if (!written.ok) return written
    return { ...written, slides: slides.length }
  },
}

const pdfTool = {
  name: 'export_pdf_noshell',
  description: [
    '生成 PDF，写到工作区的 .ai-output/ 目录。',
    DELIVERY_NOTE,
    '两种用法：给 blocks（和 export_docx_noshell 同格式，插件先生成 docx 再转 PDF），或者给 source 指向 .ai-output/ 里已有的 .docx/.xlsx/.pptx 转成 PDF。',
    'blocks 里可以放 image 块（path 指 .ai-output/ 下已有的 png/jpg/gif，或 data 给 base64），图片会真正嵌进 PDF，不是只留个文件名。',
    '转换由本机 LibreOffice 无头完成（中文、表格、排版都靠它）；装了 Office/LibreOffice 的机器上才可用。',
    '只支持 .pdf；fileName 只能是文件名，不能带路径。',
  ].join(' '),
  parameters: compileParameters({
    fileName: { type: 'string', required: true, description: '不含路径的文件名，例如 客户清单_20260925.pdf' },
    title: { type: 'string', description: '文档大标题，可省略（仅在给 blocks 时生效）' },
    blocks: { type: 'array', items: DOCX_BLOCK, description: '内容块，格式与 export_docx_noshell 相同' },
    source: { type: 'string', description: '.ai-output/ 里已有的文件名（.docx/.xlsx/.pptx），直接转它；与 blocks 二选一' },
  }),
  output: OUTPUT,
  async execute(args, exec) {
    const source = String(args.source ?? '').trim()
    const blocks = Array.isArray(args.blocks) ? args.blocks : []
    if (source === '' && blocks.length === 0) {
      return { ok: false, error: '要给 blocks（生成内容）或 source（.ai-output/ 里已有的文件），二选一。' }
    }
    if (source !== '' && !/^[^\\/:*?"<>|]+\.(docx|xlsx|pptx)$/i.test(source)) {
      return { ok: false, error: `source 只能是 .ai-output/ 下的 .docx/.xlsx/.pptx 文件名；收到「${source}」。` }
    }

    const resolved = resolveTarget(args.fileName, 'pdf', exec)
    if (!resolved.ok) return resolved

    const collected = collectMedia(blocks, exec)
    if (!collected.ok) return collected

    const found = findSoffice(activeConfig.soffice)
    if (found.error) return { ok: false, error: found.error }

    const scratch = makeScratch()
    try {
      const base = resolved.name.replace(/\.pdf$/i, '')
      let input
      let from
      if (source !== '') {
        input = join(resolved.dir, source)
        if (!existsSync(input)) return { ok: false, error: `.ai-output/ 下没有这个文件：${source}` }
        from = source.split('.').pop().toLowerCase()
      } else {
        input = join(scratch, `${base}.docx`)
        writeFileSync(input, buildDocx({ title: args.title, blocks, media: collected.media, mediaIndex: collected.mediaIndex }))
        from = 'docx'
      }

      const converted = convertToPdf({
        soffice: found.path,
        source: input,
        outDir: scratch,
        timeoutMs: activeConfig.convertTimeoutMs,
      })
      if (!converted.ok) return { ok: false, error: converted.error }

      mkdirSync(resolved.dir, { recursive: true })
      moveFile(converted.pdf, resolved.target)
      return {
        ok: true,
        path: resolved.target,
        relativePath: `${activeConfig.outputDir}/${resolved.name}`,
        bytes: statSync(resolved.target).size,
        from,
        soffice: found.path,
      }
    } catch (error) {
      return { ok: false, error: `转换失败：${error instanceof Error ? error.message : String(error)}` }
    } finally {
      cleanScratch(scratch)
    }
  },
}

/* ─────────────────────────── 出图（HTML → 截图） ─────────────────────────── */

const IMAGE_DEFAULTS = { width: 1200, height: 800, scale: 2, waitMs: 2000 }

/**
 * 把调用方给的 HTML 整理成可直接渲染的文档。
 *
 * 三种情况：已是完整文档（含 <html）就原样用；只是片段就套一层默认壳；给了
 * selector 就再注入两段 CSS —— 一是把 body 撑满，否则元素贴合视口左上角但整页
 * 截图会多出一块空白；二是把「非命中元素」压掉，让截图只框住那一个元素。
 */
function wrapHtml(html, clipped) {
  const source = clipped
    ? `<style>
  html, body { margin: 0 !important; padding: 0 !important; width: 100% !important; height: 100% !important; }
  body > *:not(${clipped}) { display: none !important; }
  ${clipped} { margin: 0 !important; }
</style>`
    : ''

  if (/<html[\s>]/i.test(html)) {
    if (source === '') return html
    return html.includes('</head>') ? html.replace('</head>', `${source}</head>`) : `${source}${html}`
  }

  return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  html, body { margin: 0; padding: 0; }
  body { font: 16px/1.6 "Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", sans-serif; color: #1f2937; background: #ffffff; }
</style>
${source}
</head>
<body>
${html}
</body>
</html>`
}

const imageTool = {
  name: 'export_image_noshell',
  description: [
    '把一段 HTML 渲染成图片（.png / .jpeg），写到工作区的 .ai-output/ 目录。',
    DELIVERY_NOTE,
    '出图路径是「HTML → 本机无头 Chrome / Edge 截图」，所以能用完整 CSS：渐变、阴影、圆角、flex / grid 布局、内联 SVG 都行；中文正常显示。',
    '要 CDN 图表库（如 ECharts）就在 html 里直接 <script src="https://...">，默认允许联网，按 waitMs 等它渲染完再截（联网不通时改成内联脚本或把库内容塞进 html）。',
    'html 给完整文档或片段都行（片段会被套一层默认样式壳）。width / height 是视口尺寸；只截某个元素就传 selector（支持 #id / .class / 标签）。',
    '要「MOM」这种文字图别用本工具硬凑：直接写 <div style="font:700 120px sans-serif">MOM</div> 更快更清。',
    '只支持 .png / .jpeg；fileName 只能是文件名，不能带路径。',
  ].join(' '),
  parameters: compileParameters({
    fileName: { type: 'string', required: true, description: '不含路径的文件名，例如 MOM.png / 客户分布_20260928.png' },
    html: { type: 'string', required: true, description: '要渲染成图片的 HTML（完整文档或片段均可）' },
    width: { type: 'number', description: '视口宽（像素），默认 1200' },
    height: { type: 'number', description: '视口高（像素），默认 800；fullPage 为 true 时只当最小高' },
    selector: { type: 'string', description: '只截这个 CSS 选择器命中的元素，例如 #chart；给了它就不截整页' },
    fullPage: { type: 'boolean', description: '默认 true，截整页（长图）；false 只截视口那一屏' },
    scale: { type: 'number', description: '设备像素比，默认 2（更清晰、文件更大）；1 表示原始像素' },
    waitMs: { type: 'number', description: '等页面渲染 / CDN 图表加载的毫秒数，默认 2000；纯静态 HTML 可给 300' },
    allowNetwork: { type: 'boolean', description: '是否允许页面联网取外部资源（CDN 图表库等），默认 true；设 false 则断网截图，只能用内联资源' },
    format: { type: 'string', enum: ['png', 'jpeg'], description: '图片格式，默认 png' },
    quality: { type: 'number', description: 'format 为 jpeg 时的画质 1-100，默认 92；png 忽略此项' },
  }),
  output: OUTPUT,
  async execute(args, exec) {
    const html = String(args.html ?? '').trim()
    if (html === '') return { ok: false, error: 'html 不能为空。' }

    const format = String(args.format ?? 'png').toLowerCase() === 'jpeg' ? 'jpeg' : 'png'

    // 不走 resolveTarget：图片的扩展名要同时认 .jpeg 和 .jpg（同一个格式的两种常见写法），
    // 而且这里必须在跑浏览器之前就把非法文件名挡掉，别白截一张图。
    const root = outputRoot(exec)
    if (root === '') return { ok: false, error: '拿不到调用方会话的工作区，且插件未配置兜底 root，无法决定写到哪。' }
    const wanted = String(args.fileName ?? '').trim()
    const name = format === 'jpeg' && /\.jpeg$/i.test(wanted) ? wanted.replace(/\.jpeg$/i, '.jpg') : wanted
    const extension = format === 'jpeg' ? '(jpg|jpeg)' : 'png'
    if (!new RegExp(`^[^\\\\/:*?"<>|]+\\.${extension}$`, 'i').test(name)) {
      return { ok: false, error: `fileName 必须是「文件名.${format === 'jpeg' ? 'jpg / .jpeg' : 'png'}」，且不含路径分隔符；收到「${wanted}」。` }
    }
    const dir = join(root, activeConfig.outputDir)
    const target = join(dir, name)
    const escape = relative(dir, target)
    if (escape.startsWith('..') || isAbsolute(escape)) return { ok: false, error: '目标路径越界，已拒绝。' }

    const found = findChrome(activeConfig.chrome)
    if (found.error) return { ok: false, error: found.error }

    const selector = String(args.selector ?? '').trim()
    const width = clampInt(args.width, 1, 10000, IMAGE_DEFAULTS.width)
    const height = clampInt(args.height, 1, 10000, IMAGE_DEFAULTS.height)
    const scale = clampInt(args.scale, 1, 4, IMAGE_DEFAULTS.scale)
    const waitMs = clampInt(args.waitMs, 0, 30000, IMAGE_DEFAULTS.waitMs)
    const quality = clampInt(args.quality, 1, 100, 92)
    const allowNetwork = args.allowNetwork !== false
    const fullPage = selector === '' && args.fullPage !== false

    const scratch = makeScratch()
    try {
      writeFileSync(join(scratch, 'page.html'), wrapHtml(html, selector !== ''), 'utf8')
      const shot = shootHtml({
        chrome: found.path,
        scratch,
        format,
        quality,
        width,
        height,
        deviceScaleFactor: scale,
        fullPage,
        allowNetwork,
        renderWaitMs: waitMs,
        timeoutMs: activeConfig.shotTimeoutMs,
      })
      if (!shot.ok) return shot

      mkdirSync(dir, { recursive: true })
      moveFile(shot.file, target)
      return {
        ok: true,
        path: target,
        relativePath: `${activeConfig.outputDir}/${name}`,
        bytes: statSync(target).size,
        format,
        width,
        height,
        fullPage,
        // 不能写成 `selector: selector === '' ? undefined : selector`：运行时会把这个值
        // 送进 lossless-JSON 快照（packages/util/values 的 snapshotJsonValue），而
        // `undefined` 不是合法 JSON 值 —— 一句 `value is not lossless JSON` 会把整次
        // 调用判成 invalid output，哪怕图已经写好。没有命中选择器就干脆不带这个键。
        ...(selector === '' ? {} : { selector }),
        chrome: found.path,
      }
    } catch (error) {
      return { ok: false, error: `截图失败：${error instanceof Error ? error.message : String(error)}` }
    } finally {
      cleanScratch(scratch)
    }
  },
}

/* ────────────────────────────── 装载 ────────────────────────────── */

/**
 * 注册五个工具。
 * @param ctx - 提供 `tools` 注册表的宿主上下文。
 * @param config - 已校验并填好默认值的插件配置。
 */
export function apply(ctx, config) {
  activeConfig = config ?? { ...DEFAULTS }
  ctx.tools.register(xlsxTool)
  ctx.tools.register(docxTool)
  ctx.tools.register(pptxTool)
  ctx.tools.register(pdfTool)
  ctx.tools.register(imageTool)
}
