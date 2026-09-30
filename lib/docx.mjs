/**
 * DOCX 生成器：标题 + 段落 / 小标题 / 项目符号 / 表格 / 图片 → 真正的 Word 文档。
 *
 * 图片不是「引用路径」，而是真嵌进包里：字节进 `word/media/`，正文里放一个
 * `<w:drawing>`，关系表里补一条 image 关系，内容类型里补一条图片默认类型。
 * 这也正是 PDF 能带图的原因 —— PDF 是把这份 docx 交给 LibreOffice 转的。
 *
 * @module lib/docx
 */

import { XML_HEAD, escapeXml, imageSpec, packageParts, zip } from './zip.mjs'

const NS_MAIN = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const NS_WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const NS_PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'

/** 版心宽度（A4 减左右页边距，单位 dxa）与 1 px 对应的 EMU。 */
const CONTENT_DXA = 9360
const EMU_PER_PX = 9525
/** 图片默认按版心的 90% 宽排，且高度不超过 620px，免得一张长图把版面顶爆。 */
const DEFAULT_IMAGE_WIDTH_PCT = 90
const MAX_IMAGE_HEIGHT_PX = 620

const STYLES = `${XML_HEAD}
<w:styles xmlns:w="${NS_MAIN}">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="微软雅黑"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="288" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="0" w:after="280"/></w:pPr><w:rPr><w:b/><w:sz w:val="48"/><w:szCs w:val="48"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="280" w:after="140"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/><w:szCs w:val="26"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="100"/><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="420" w:hanging="210"/><w:spacing w:after="60"/></w:pPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:left w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:right w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="BFBFBF"/></w:tblBorders></w:tblPr></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
</w:styles>`

const run = (text, bold) =>
  `<w:r>${bold ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`

const para = (text, style, bold) =>
  `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ''}${text === '' ? '' : run(text, bold)}</w:p>`

function table(block) {
  const columns = block.columns.map((c) => String(c))
  const width = Math.floor(CONTENT_DXA / Math.max(columns.length, 1))
  const grid = columns.map(() => `<w:gridCol w:w="${width}"/>`).join('')
  const row = (cells, bold) =>
    `<w:tr>${columns
      .map((_, ci) => `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${para(cells[ci] ?? '', undefined, bold)}</w:tc>`)
      .join('')}</w:tr>`
  const rows = [row(columns, true), ...block.rows.map((cells) => row(Array.isArray(cells) ? cells : [cells], false))]
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rows.join('')}</w:tbl>`
}

/** 算出图片的显示尺寸（px）：按版心比例定宽，再按原始比例推高，最后压高度上限。 */
function imageBox(spec, widthPct, requestedWidthPx) {
  const ratio = spec.widthPx > 0 && spec.heightPx > 0 ? spec.heightPx / spec.widthPx : 1
  const contentPx = Math.round((CONTENT_DXA / 20) * (96 / 72)) // 版心 dxa → 96dpi 像素（约 624px）
  const pct = Number.isFinite(widthPct) && widthPct > 0 ? Math.min(widthPct, 100) : DEFAULT_IMAGE_WIDTH_PCT
  let width = Number.isFinite(requestedWidthPx) && requestedWidthPx > 0 ? requestedWidthPx : Math.round((contentPx * pct) / 100)
  width = Math.max(16, Math.min(width, contentPx * 2))
  let height = Math.round(width * ratio)
  if (height > MAX_IMAGE_HEIGHT_PX) {
    width = Math.round((width * MAX_IMAGE_HEIGHT_PX) / height)
    height = MAX_IMAGE_HEIGHT_PX
  }
  return { width, height }
}

/** 一张居中排列的内联图片。 */
function image(block, documents, imageIndex) {
  const item = documents[imageIndex]
  if (!item) return ''
  const { width, height } = imageBox(item.spec, Number(block.widthPct), Number(block.width))
  const cx = width * EMU_PER_PX
  const cy = height * EMU_PER_PX
  const rid = `rIdImg-${imageIndex}`
  const drawing = `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">`
    + `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`
    + `<wp:docPr id="${imageIndex + 1}" name="图片 ${imageIndex + 1}"/>`
    + `<a:graphic><a:graphicData uri="${NS_PIC}"><pic:pic>`
    + `<pic:nvPicPr><pic:cNvPr id="${imageIndex + 1}" name="${escapeXml(item.name)}"/><pic:cNvPicPr/></pic:nvPicPr>`
    + `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
    + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`
    + `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`
    + `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`
  return `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="120" w:after="60"/></w:pPr><w:r>${drawing}</w:r></w:p>`
}

/**
 * 组装 DOCX。
 * @param options - `title`、`blocks`（heading / paragraph / bullet / table / image）与 `media`。
 *   `media` 是图片清单：`{ name, data, spec }`，其中 `data` 必须是 Buffer，
 *   `spec` 来自 `imageSpec()`；`mediaIndex[i]` 指出第 i 个块用第几张图（非图片块为 -1）。
 *   对应关系按**下标**传进来，而不是写回块对象 —— 工具参数是深冻结的，写不进去
 *   （原因见 index.js 的 collectMedia 注释）。
 * @returns 可直接写盘的 Buffer。
 */
export function buildDocx({ title, blocks, media, mediaIndex }) {
  const documents = media ?? []
  const indices = mediaIndex ?? []
  const body = []
  if (title) body.push(para(String(title), 'Title'))
  for (const [position, block] of (blocks ?? []).entries()) {
    const text = block.text === undefined ? '' : String(block.text)
    switch (block.type) {
      case 'heading': {
        const level = Math.min(Math.max(Number(block.level) || 1, 1), 3)
        body.push(para(text, `Heading${level}`))
        break
      }
      case 'bullet':
        body.push(para(`• ${text}`, 'ListParagraph'))
        break
      case 'table':
        if (Array.isArray(block.columns) && block.columns.length > 0) body.push(table(block))
        break
      case 'image':
        body.push(image(block, documents, indices[position]))
        break
      default:
        body.push(para(text))
    }
  }
  body.push('<w:p/>')

  // 每张图一条 image 关系（Id 与块里的 rIdImg-<i> 对齐），外加原来的 styles 关系。
  const imageDefaults = []
  const imageRels = []
  const mediaEntries = []
  for (const [index, item] of documents.entries()) {
    imageRels.push(`<Relationship Id="rIdImg-${index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${item.name}"/>`)
    if (!imageDefaults.some(([ext]) => ext === item.spec.ext)) imageDefaults.push([item.spec.ext, item.spec.mime])
    mediaEntries.push([`word/media/${item.name}`, item.data])
  }

  const [rels, types] = packageParts({
    mainPart: 'word/document.xml',
    mainType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    defaults: imageDefaults,
    overrides: [['word/styles.xml', 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml']],
  })

  const document = `${XML_HEAD}
<w:document xmlns:w="${NS_MAIN}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="${NS_WP}" xmlns:a="${NS_A}" xmlns:pic="${NS_PIC}">
<w:body>
${body.join('\n')}
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>
</w:body>
</w:document>`

  return zip([
    types,
    rels,
    ['word/document.xml', document],
    ['word/_rels/document.xml.rels', `${XML_HEAD}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
${imageRels.join('\n')}
</Relationships>`],
    ['word/styles.xml', STYLES],
    ...mediaEntries,
  ])
}
