/**
 * export_docx_noshell 分段暂存 —— 把「一次吐完整篇大 JSON」拆成多批提交。
 *
 * 为什么分段：整篇文档的 blocks 挤在一次工具调用里，内容越长越容易截断、漏引号、
 * 漏括号；JSON 是「一错全挂」的格式，坏一处就得整篇重来。分批之后每批都是小 JSON，
 * 单批失败只影响那一批。
 *
 * 重试怎么办：批次号就是文件名（`part-0001.json`），同一个号重传 = 覆盖写，
 * 所以重试天然幂等，不会写进两份内容。
 *
 * 存哪儿：`<工作区>/<outputDir>/.staging/docx/<哈希>/`。哈希由「会话 id + 文件名」
 * 算出来 —— 同一会话里的两个文档、两个会话里的同名文档都不会串味。收尾生成成功后
 * **整个目录（含 index.json）一次删掉**；失败就原样保留，方便补齐缺号再收尾。
 *
 * `index.json` 只记「哪一批多少块」，给落批时算上限用 —— 不然每落一批都要把所有
 * 历史批次读一遍，批次数一多就是 O(n²)（实测 200 批要 37 秒）。索引对不上目录
 * （丢了、被并发写坏、手删了 part 文件）就当场按文件重建，永远自愈。
 *
 * @module staging
 */

import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, readdirSync, renameSync, rmdirSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'

/** 暂存根目录（相对 outputDir）；带点前缀，免得混进正常交付清单。 */
const STAGING_ROOT = '.staging/docx'

/** 单篇文档的批次数上限，防止「一直传下去」停不下来。 */
export const MAX_PARTS = 200

/** 批次文件名固定 4 位补零，字典序即批次顺序。 */
const PART_PATTERN = /^part-(\d{4})\.json$/

/** 索引文件名。 */
const INDEX_FILE = 'index.json'

/** 算这篇文档的暂存目录绝对路径；同一个会话 + 同一个文件名永远算出同一个目录。 */
export function stageDir(root, outputDir, sessionId, fileName) {
  const hash = createHash('sha1').update(`${sessionId}\u0000${fileName}`).digest('hex').slice(0, 16)
  return join(root, outputDir, STAGING_ROOT, hash)
}

/** 目录里现存的批次号，按升序。只列文件名，不读内容。 */
function partNumbers(dir) {
  if (!existsSync(dir)) return []
  return readdirSync(dir)
    .filter((item) => PART_PATTERN.test(item))
    .map((item) => Number(PART_PATTERN.exec(item)[1]))
    .sort((left, right) => left - right)
}

/** 只读索引，不做一致性判断；坏了当没有。 */
function readIndexRaw(dir) {
  try {
    const parsed = JSON.parse(readFileSync(join(dir, INDEX_FILE), 'utf8'))
    return parsed?.blocks !== null && typeof parsed?.blocks === 'object' ? parsed.blocks : {}
  } catch {
    return {}
  }
}

/** 写索引；写不进去不算错（下次按文件重建）。 */
function writeIndex(dir, blocks) {
  try {
    writeFileSync(join(dir, INDEX_FILE), JSON.stringify({ blocks, updatedAt: new Date().toISOString() }), 'utf8')
  } catch {
    /* 忽略：索引只是加速用的 */
  }
}

/** 数一个 part 文件里有几个块；读不出来记 0，坏文件留给收尾时正式报错。 */
function countBlocks(file) {
  try {
    const parsed = JSON.parse(readFileSync(file, 'utf8'))
    return Array.isArray(parsed?.blocks) ? parsed.blocks.length : 0
  } catch {
    return 0
  }
}

/**
 * 落一批。
 *
 * 先写 `.tmp` 再 rename：中途失败也不会留下半截 JSON 被下一次读进来。
 * 同一个批次号已经存在就覆盖，并把 `replaced` 报给调用方（用来提示「这是重传」）。
 * 顺手把这一批的块数增量写进 `index.json`（不重读任何历史批次）。
 */
export function savePart(dir, part, { title, blocks, meta }) {
  mkdirSync(dir, { recursive: true })
  const target = join(dir, `part-${String(part).padStart(4, '0')}.json`)
  const replaced = existsSync(target)

  const temp = `${target}.tmp`
  writeFileSync(temp, JSON.stringify({ part, title: title ?? '', blocks, savedAt: new Date().toISOString() }), 'utf8')
  renameSync(temp, target)

  const index = readIndexRaw(dir)
  index[part] = blocks.length
  writeIndex(dir, index)

  // meta.json 只写一次：人翻暂存目录时能认出这是哪个会话的哪篇文档。
  const metaFile = join(dir, 'meta.json')
  if (!existsSync(metaFile)) writeFileSync(metaFile, JSON.stringify({ ...meta, createdAt: new Date().toISOString() }, null, 2), 'utf8')

  return { replaced }
}

/**
 * 只问「存了哪些批、一共多少块」，不读正文。
 *
 * 批次号来自文件名；块数走 `index.json`，索引与目录对不上（丢了 / 少了 / 多了）
 * 就按文件重建一次 —— 所以索引永远不是「真相」，只是缓存。
 */
export function stagingState(dir) {
  const parts = partNumbers(dir)
  if (parts.length === 0) return { parts: [], blocks: {}, total: 0 }

  const index = readIndexRaw(dir)
  const consistent = parts.every((part) => Number.isInteger(index[part])) && Object.keys(index).length === parts.length
  const blocks = consistent ? index : Object.fromEntries(parts.map((part) => [part, countBlocks(join(dir, `part-${String(part).padStart(4, '0')}.json`))]))
  if (!consistent) writeIndex(dir, blocks)

  return { parts, blocks, total: parts.reduce((sum, part) => sum + (blocks[part] ?? 0), 0) }
}

/** 读回已存批次的完整内容（只有收尾时才需要），按批次号升序。文件坏了就报错，不猜内容。 */
export function readParts(dir) {
  const parts = []
  for (const part of partNumbers(dir)) {
    const name = `part-${String(part).padStart(4, '0')}.json`
    try {
      const parsed = JSON.parse(readFileSync(join(dir, name), 'utf8'))
      parts.push({
        part,
        title: typeof parsed?.title === 'string' ? parsed.title : '',
        blocks: Array.isArray(parsed?.blocks) ? parsed.blocks : [],
      })
    } catch (error) {
      return {
        parts,
        error: `暂存批次 ${name} 读不出来（${error instanceof Error ? error.message : String(error)}）；可加 reset: true 清空重传。`,
      }
    }
  }
  return { parts }
}

/** 清掉这篇文档的全部暂存（part 文件 + index.json + meta.json 一起没）；清不掉不算错。 */
export function clearStage(dir) {
  try {
    rmSync(dir, { recursive: true, force: true })
  } catch {
    /* 忽略：暂存残留不影响交付 */
  }
  // 顺带收掉留空的上层目录（.staging/docx、.staging）：非空会抛 ENOTEMPTY，忽略即可。
  for (const parent of [join(dir, '..'), join(dir, '..', '..')]) {
    try {
      rmdirSync(parent)
    } catch {
      /* 忽略 */
    }
  }
}
