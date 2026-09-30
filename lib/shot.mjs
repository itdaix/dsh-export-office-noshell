/**
 * HTML 截图：把一段 HTML 渲染成 PNG / JPEG。
 *
 * 为什么不自己画图：本机没有 Node 端图形栈（不引 canvas / sharp 的话），
 * 而「HTML + CSS」本身就是一套完整的排版语言，浏览器又是现成的渲染器 ——
 * 所以这里也只做调度：把 HTML 落成临时文件，交给无头 Chrome / Edge 截一张。
 *
 * 与 pdf.mjs 同构：找可执行文件 → 建临时目录 → 同步跑子进程 →
 * 以产出文件为准（不信任退出码）→ 开不了管道就退回不捕获输出重试。
 *
 * @module lib/shot
 */

import { spawnSync } from 'node:child_process'
import { existsSync, readdirSync, statSync } from 'node:fs'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'

/** 浏览器常见安装位置：Chrome 优先，Edge 兜底（Chromium 内核即可）。 */
const CANDIDATES = [
  'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
  'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
  'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
  'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
  '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
  '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
  '/usr/bin/google-chrome',
  '/usr/bin/google-chrome-stable',
  '/usr/bin/chromium',
  '/usr/bin/chromium-browser',
  '/usr/bin/microsoft-edge',
]

const REMEDY = '装 Chrome 或 Edge 即可；也可以把可执行文件路径写进插件配置的 chrome，或设 CHROME_PATH 环境变量。'

/**
 * 找浏览器。
 * @param configured - 插件配置里指定的路径，优先用。
 * @returns `{ path }`，找不到则是 `{ error }`。
 */
export function findChrome(configured) {
  const wanted = String(configured ?? '').trim()
  if (wanted !== '') {
    return existsSync(wanted) ? { path: wanted } : { error: `配置里指定的浏览器不存在：${wanted}。` }
  }

  const fromEnv = String(process.env.CHROME_PATH ?? '').trim()
  if (fromEnv !== '' && existsSync(fromEnv)) return { path: fromEnv }

  const found = CANDIDATES.find((candidate) => existsSync(candidate))
  return found ? { path: found } : { error: `没找到 Chrome / Edge（找过：${CANDIDATES.join(' / ')}）。截图需要它，${REMEDY}` }
}

/** 尺寸 / 倍率收敛到浏览器肯接受的区间，非法值退回默认。 */
export function clampInt(value, min, max, fallback) {
  const number = Number(value)
  if (!Number.isFinite(number)) return fallback
  return Math.min(max, Math.max(min, Math.round(number)))
}

/**
 * 截一张图。
 *
 * 退出码与 stdout 都不可信（headless 截完图有时仍报非 0），所以判定一律以
 * 「临时目录里有没有产出文件」为准，和 LibreOffice 那条路子保持一致。
 *
 * @returns `{ ok: true, file }` 或 `{ ok: false, error }`。
 */
export function shootHtml({
  chrome,
  scratch,
  format = 'png',
  quality = 92,
  width = 1200,
  height = 800,
  deviceScaleFactor,
  fullPage = true,
  allowNetwork = true,
  renderWaitMs = 2000,
  timeoutMs,
  capture = true,
}) {
  const page = join(scratch, 'page.html')
  const out = join(scratch, `shot.${format}`)

  const args = [
    '--headless=new',
    // 首启弹窗 / 同步 / 默认浏览器检查都会污染截图，逐个关掉。
    '--no-first-run',
    '--no-default-browser-check',
    '--disable-sync',
    '--disable-extensions',
    '--disable-background-networking',
    '--hide-scrollbars',
    '--force-color-profile=srgb',
    `--user-data-dir=${join(scratch, 'profile')}`,
    `--window-size=${width},${height}`,
    `--screenshot=${out}`,
  ]

  // 联网时给一个虚拟时间预算，等页面把外部资源（CDN 图表库）加载完再截；
  // 断网时反过来把网络组件整个关掉，页面只能吃内联资源。
  if (allowNetwork) args.push(`--virtual-time-budget=${renderWaitMs}`)
  else args.push('--disable-features=NetworkService', '--host-resolver-rules=MAP * ~NOTFOUND')
  if (format === 'jpeg') args.push('--screenshot-format=jpeg', `--screenshot-quality=${quality}`)
  if (typeof deviceScaleFactor === 'number') args.push(`--force-device-scale-factor=${deviceScaleFactor}`)
  if (fullPage) args.push('--screenshot-full-page')
  args.push(pathToFileURL(page).href)

  const result = capture
    ? spawnSync(chrome, args, { timeout: timeoutMs, encoding: 'utf8' })
    : spawnSync(chrome, args, { timeout: timeoutMs, stdio: 'ignore' })
  if (result.error?.code === 'EPERM') {
    // 受限沙箱里开不了管道（piped stdio 会被拒），退回不捕获输出再试一次。
    return shootHtml({ chrome, scratch, format, quality, width, height, deviceScaleFactor, fullPage, allowNetwork, renderWaitMs, timeoutMs, capture: false })
  }
  if (result.error) {
    if (result.error.code === 'ETIMEDOUT') return { ok: false, error: `截图超时（${timeoutMs} 毫秒）。页面里可能有拿不到的资源。` }
    return { ok: false, error: `调用浏览器失败：${result.error.message}` }
  }

  if (existsSync(out) && statSync(out).size > 0) return { ok: true, file: out }

  const stray = readdirSync(scratch).filter((name) => name.toLowerCase().startsWith('screenshot'))
  if (stray.length > 0) return { ok: true, file: join(scratch, stray[0]) }

  const detail = `${result.stdout ?? ''}${result.stderr ?? ''}`.trim().split('\n').slice(-3).join(' ')
  return { ok: false, error: `浏览器没有产出图片（退出码 ${result.status}）。${detail}` }
}
