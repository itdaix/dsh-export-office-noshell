/**
 * PDF 转换：把生成好的 Office 文件交给本机 LibreOffice 无头转成 PDF。
 *
 * 为什么不自己写 PDF：PDF 里显示中文要嵌入字体（还要切子集、写 CID 映射），
 * 成本远高于「生成 docx 再让 LibreOffice 转一道」。所以这里只做调度。
 *
 * @module lib/pdf
 */

import { spawnSync } from 'node:child_process'
import { existsSync, mkdirSync, readdirSync, rmSync, statSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'

/** 本插件在系统临时目录里用的目录前缀（清理时只认它，不碰别家）。 */
const SCRATCH_PREFIX = 'dsh-export-office-noshell-'

/** LibreOffice 常见安装位置。 */
const CANDIDATES = [
  'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
  'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
  '/usr/bin/soffice',
  '/usr/local/bin/soffice',
]

/**
 * 找 LibreOffice。
 * @param configured - 插件配置里指定的路径，优先用。
 * @returns `{ path }`，找不到则是 `{ error }`。
 */
export function findSoffice(configured) {
  if (configured) {
    return existsSync(configured)
      ? { path: configured }
      : { error: `配置里指定的 LibreOffice 不存在：${configured}` }
  }
  const found = CANDIDATES.find((candidate) => existsSync(candidate))
  return found
    ? { path: found }
    : { error: `没找到 LibreOffice（找过：${CANDIDATES.join(' / ')}）。PDF 导出需要它，或把路径写进插件配置的 soffice。` }
}

/**
 * 清掉自己以前留下的临时工作目录。
 *
 * 转换失败（或进程被强杀）时 `cleanScratch` 删不掉，那些目录就会一直躺在 TEMP 里。
 * 这里只删**本插件自己造的**前缀目录，且只删超过一小时的 —— 不动别人、也不影响
 * 正在跑的另一次转换。
 */
function sweepOldScratch(maxAgeMs = 60 * 60 * 1000) {
  let entries
  try {
    entries = readdirSync(tmpdir(), { withFileTypes: true })
  } catch {
    return
  }
  const cutoff = Date.now() - maxAgeMs
  for (const entry of entries) {
    if (!entry.isDirectory()) continue
    if (!entry.name.startsWith(SCRATCH_PREFIX)) continue
    const full = join(tmpdir(), entry.name)
    try {
      if (statSync(full).mtimeMs > cutoff) continue
      rmSync(full, { recursive: true, force: true })
    } catch {
      /* 正被别人占着就下次再说 */
    }
  }
}

/** 一个独立的工作目录，用完就删。 */
export function makeScratch() {
  sweepOldScratch()
  const dir = join(tmpdir(), `dsh-export-office-noshell-${process.pid}-${Date.now().toString(36)}`)
  mkdirSync(dir, { recursive: true })
  return dir
}

export function cleanScratch(dir) {
  try {
    rmSync(dir, { recursive: true, force: true })
  } catch {
    /* 临时目录清不掉不影响结果 */
  }
}

/**
 * 转一个文件为 PDF。
 *
 * 配置沿用一份共享的 `-env:UserInstallation`（放在系统临时目录）。
 * 曾经试过「每次转换都用全新隔离配置」来规避 Windows 的「正在等待打印机连接」
 * 弹框，**实测无效** —— 那个弹框的真因是 WSD 打印机在 VPN 下失联、而 LibreOffice
 * 启动时会枚举打印机（详见 README「已知环境问题」），换配置目录拦不住枚举本身。
 * 所以这里回到原样，保持简单。
 *
 * @returns `{ ok: true, pdf }` 或 `{ ok: false, error }`。
 */
export function convertToPdf({ soffice, source, outDir, timeoutMs, capture = true }) {
  const profile = join(tmpdir(), 'dsh-export-office-noshell-loprofile')
  const args = [
    `-env:UserInstallation=${pathToFileURL(profile).href}`,
    '--headless',
    '--norestore',
    '--convert-to',
    'pdf',
    '--outdir',
    outDir,
    source,
  ]

  const result = capture
    ? spawnSync(soffice, args, { timeout: timeoutMs, encoding: 'utf8' })
    : spawnSync(soffice, args, { timeout: timeoutMs, stdio: 'ignore' })
  if (result.error?.code === 'EPERM') {
    // 受限沙箱里开不了管道（piped stdio 会被拒），退回不捕获输出再试一次。
    return convertToPdf({ soffice, source, outDir, timeoutMs, capture: false })
  }
  if (result.error) {
    return { ok: false, error: `调用 LibreOffice 失败：${result.error.message}` }
  }

  // soffice 有时退出码非 0，但文件已经产出 —— 以文件为准。
  const produced = readdirSync(outDir).filter((name) => name.toLowerCase().endsWith('.pdf'))
  if (produced.length === 0) {
    const detail = `${result.stdout ?? ''}${result.stderr ?? ''}`.trim().split('\n').slice(-3).join(' ')
    return { ok: false, error: `LibreOffice 没有产出 PDF（退出码 ${result.status}）。${detail}` }
  }
  return { ok: true, pdf: join(outDir, produced[0]) }
}
