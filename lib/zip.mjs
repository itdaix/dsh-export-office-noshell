/**
 * 共用零件：极简 ZIP 打包 + XML 转义。只用 node:zlib，零依赖。
 *
 * docx / xlsx / pptx 都是 OOXML 包，本质就是「一批 XML 打包成 zip」。
 *
 * @module lib/zip
 */

import { deflateRawSync } from 'node:zlib'

export const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
export const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
export const NS_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
export const NS_CT = 'http://schemas.openxmlformats.org/package/2006/content-types'

/** XML 文本转义。 */
export function escapeXml(value) {
  return String(value).replace(/[&<>"']/g, (ch) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;',
  })[ch])
}

/** 显示宽度：CJK 算 2 列。 */
export function widthOf(text) {
  return [...String(text)].reduce((w, ch) => w + (/[\u1100-\uFFE6]/.test(ch) ? 2 : 1), 0)
}

/**
 * 从图片字节里读出像素尺寸与真实格式。
 *
 * 只认三种最常见的位图（png / jpeg / gif）—— PDF 与 Word 的图片嵌入都靠它，
 * 尺寸读不出来就没法按比例排版。认不出来返回 `null`，由调用方决定怎么报错。
 *
 * @returns `{ kind, mime, ext, widthPx, heightPx }` 或 `null`。
 */
export function imageSpec(buffer) {
  if (!Buffer.isBuffer(buffer) || buffer.length < 16) return null

  // PNG：文件头 8 字节固定，紧跟 IHDR，宽高各 4 字节大端
  if (buffer.readUInt32BE(0) === 0x89504e47 && buffer.readUInt32BE(4) === 0x0d0a1a0a) {
    return { kind: 'png', mime: 'image/png', ext: 'png', widthPx: buffer.readUInt32BE(16), heightPx: buffer.readUInt32BE(20) }
  }

  // GIF：逻辑屏幕描述符里的宽高，小端 16 位
  if (buffer.subarray(0, 3).toString('latin1') === 'GIF' && buffer.length >= 10) {
    return { kind: 'gif', mime: 'image/gif', ext: 'gif', widthPx: buffer.readUInt16LE(6), heightPx: buffer.readUInt16LE(8) }
  }

  // JPEG：扫段找 SOFn，尺寸在段内偏移 5 / 7 处，大端
  if (buffer.readUInt16BE(0) === 0xffd8) {
    let cursor = 2
    while (cursor + 9 < buffer.length) {
      if (buffer[cursor] !== 0xff) {
        cursor += 1
        continue
      }
      const marker = buffer[cursor + 1]
      if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
        cursor += 2
        continue
      }
      const size = buffer.readUInt16BE(cursor + 2)
      const isFrame = (marker >= 0xc0 && marker <= 0xc3) || (marker >= 0xc5 && marker <= 0xc7) || (marker >= 0xc9 && marker <= 0xcb) || (marker >= 0xcd && marker <= 0xcf)
      if (isFrame) {
        return { kind: 'jpeg', mime: 'image/jpeg', ext: 'jpg', widthPx: buffer.readUInt16BE(cursor + 5), heightPx: buffer.readUInt16BE(cursor + 7) }
      }
      if (size < 2) break
      cursor += 2 + size
    }
  }

  return null
}

/** 按扩展名推断图片的 MIME，供「用户给了文件名但内容还没读」时使用。 */
export function mimeOfExtension(extension) {
  const ext = String(extension).toLowerCase()
  if (ext === 'jpg' || ext === 'jpeg') return 'image/jpeg'
  if (ext === 'png') return 'image/png'
  if (ext === 'gif') return 'image/gif'
  return ''
}

/* ────────────────────────────── ZIP ────────────────────────────── */

const CRC = (() => {
  const table = new Int32Array(256)
  for (let n = 0; n < 256; n += 1) {
    let c = n
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    table[n] = c
  }
  return table
})()

function crc32(buf) {
  let c = -1
  for (const byte of buf) c = CRC[(c ^ byte) & 0xff] ^ (c >>> 8)
  return (c ^ -1) >>> 0
}

/**
 * 把 `[名称, 内容]` 列表打成 zip。
 *
 * 内容可以是字符串（XML 之类的文本），也可以是 Buffer —— 图片这类二进制媒体
 * 必须走 Buffer，一旦先当 utf8 字符串处理就会被改写字节。
 *
 * @param entries - 条目列表，名称用 `/` 分隔，不带前导斜杠。
 * @returns 可直接写盘的 Buffer。
 */
export function zip(entries) {
  const parts = []
  const central = []
  let offset = 0

  for (const [name, content] of entries) {
    const nameBuf = Buffer.from(name, 'utf8')
    const raw = Buffer.isBuffer(content) ? content : Buffer.from(content, 'utf8')
    const deflated = deflateRawSync(raw, { level: 9 })
    const useDeflate = deflated.length < raw.length
    const body = useDeflate ? deflated : raw
    const method = useDeflate ? 8 : 0
    const crc = crc32(raw)

    const local = Buffer.alloc(30)
    local.writeUInt32LE(0x04034b50, 0)
    local.writeUInt16LE(20, 4)
    local.writeUInt16LE(0x0800, 6) // 文件名按 UTF-8
    local.writeUInt16LE(method, 8)
    local.writeUInt16LE(0, 10)
    local.writeUInt16LE(0x21, 12) // 1980-01-01
    local.writeUInt32LE(crc, 14)
    local.writeUInt32LE(body.length, 18)
    local.writeUInt32LE(raw.length, 22)
    local.writeUInt16LE(nameBuf.length, 26)
    local.writeUInt16LE(0, 28)
    parts.push(local, nameBuf, body)

    const entry = Buffer.alloc(46)
    entry.writeUInt32LE(0x02014b50, 0)
    entry.writeUInt16LE(20, 4)
    entry.writeUInt16LE(20, 6)
    entry.writeUInt16LE(0x0800, 8)
    entry.writeUInt16LE(method, 10)
    entry.writeUInt16LE(0, 12)
    entry.writeUInt16LE(0x21, 14)
    entry.writeUInt32LE(crc, 16)
    entry.writeUInt32LE(body.length, 20)
    entry.writeUInt32LE(raw.length, 24)
    entry.writeUInt16LE(nameBuf.length, 28)
    entry.writeUInt16LE(0, 30)
    entry.writeUInt16LE(0, 32)
    entry.writeUInt16LE(0, 34)
    entry.writeUInt16LE(0, 36)
    entry.writeUInt32LE(0, 38)
    entry.writeUInt32LE(offset, 42)
    central.push(entry, nameBuf)

    offset += local.length + nameBuf.length + body.length
  }

  const centralBuf = Buffer.concat(central)
  const end = Buffer.alloc(22)
  end.writeUInt32LE(0x06054b50, 0)
  end.writeUInt16LE(0, 4)
  end.writeUInt16LE(0, 6)
  end.writeUInt16LE(entries.length, 8)
  end.writeUInt16LE(entries.length, 10)
  end.writeUInt32LE(centralBuf.length, 12)
  end.writeUInt32LE(offset, 16)
  end.writeUInt16LE(0, 20)
  return Buffer.concat([...parts, centralBuf, end])
}

/** 公共部件：`_rels/.rels` 与 `[Content_Types].xml`。 */
export function packageParts({ mainPart, mainType, defaults, overrides }) {
  return [
    ['_rels/.rels', `${XML_HEAD}
<Relationships xmlns="${NS_PKG}">
<Relationship Id="rId1" Type="${NS_REL}/officeDocument" Target="${mainPart}"/>
</Relationships>`],
    ['[Content_Types].xml', `${XML_HEAD}
<Types xmlns="${NS_CT}">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
${(defaults ?? []).map(([ext, type]) => `<Default Extension="${ext}" ContentType="${type}"/>`).join('\n')}
<Override PartName="/${mainPart}" ContentType="${mainType}"/>
${overrides.map(([part, type]) => `<Override PartName="/${part}" ContentType="${type}"/>`).join('\n')}
</Types>`],
  ]
}
