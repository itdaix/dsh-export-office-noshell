/**
 * PPTX 生成器：每页「标题 + 项目符号」→ 真正的 PowerPoint 演示文稿。
 *
 * 版式用固定的空白母版 + 显式文本框，避免依赖占位符继承链。
 *
 * @module lib/pptx
 */

import { XML_HEAD, escapeXml, packageParts, zip } from './zip.mjs'

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'

const THEME = `${XML_HEAD}
<a:theme xmlns:a="${A}" name="Office 主题">
<a:themeElements>
<a:clrScheme name="Office">
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="Office">
<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface="微软雅黑"/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface="微软雅黑"/><a:cs typeface=""/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
<a:objectDefaults/>
<a:extraClrSchemeLst/>
</a:theme>`

const SP_TREE_HEAD = `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>`

const SLIDE_MASTER = `${XML_HEAD}
<p:sldMaster xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}">
<p:cSld>
<p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
<p:spTree>${SP_TREE_HEAD}</p:spTree>
</p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>`

const SLIDE_LAYOUT = `${XML_HEAD}
<p:sldLayout xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}" type="blank" preserve="1">
<p:cSld name="空白"><p:spTree>${SP_TREE_HEAD}</p:spTree></p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`

const rels = (items) => `${XML_HEAD}
<Relationships xmlns="${PKG}">
${items.map(([id, type, target]) => `<Relationship Id="${id}" Type="${type}" Target="${target}"/>`).join('\n')}
</Relationships>`

/** 一个文本框。 */
function textBox({ id, name, x, y, cx, cy, size, bold, paragraphs }) {
  if (paragraphs.length === 0) return ''
  const body = paragraphs
    .map(({ text, bullet, level }) => {
      const depth = Math.min(Math.max(Number(level) || 0, 0), 4)
      const marL = 342900 + depth * 342900
      const pPr = bullet
        ? `<a:pPr marL="${marL}" indent="-342900"><a:buFont typeface="Arial"/><a:buChar char="•"/></a:pPr>`
        : `<a:pPr marL="${marL}" indent="0"/>`
      return `<a:p>${pPr}<a:r><a:rPr lang="zh-CN" sz="${size}"${bold ? ' b="1"' : ''} dirty="0"/><a:t xml:space="preserve">${escapeXml(text)}</a:t></a:r></a:p>`
    })
    .join('')
  return `<p:sp>
<p:nvSpPr><p:cNvPr id="${id}" name="${escapeXml(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>
<p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:normAutofit/></a:bodyPr><a:lstStyle/>${body}</p:txBody>
</p:sp>`
}

function slideXml(slide) {
  const title = slide.title === undefined ? '' : String(slide.title)
  const bullets = (slide.bullets ?? []).map((item) =>
    typeof item === 'string' ? { text: item } : { text: String(item?.text ?? ''), level: item?.level })
  const shapes = [
    textBox({
      id: 2, name: 'Title', x: 838200, y: 457200, cx: 10515600, cy: 1325563,
      size: 3600, bold: true, paragraphs: title === '' ? [] : [{ text: title }],
    }),
    textBox({
      id: 3, name: 'Content', x: 838200, y: 1825625, cx: 10515600, cy: 4351338,
      size: 2000, bold: false, paragraphs: bullets.filter((item) => item.text !== '').map((item) => ({ ...item, bullet: true })),
    }),
  ].join('')
  return `${XML_HEAD}
<p:sld xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}">
<p:cSld><p:spTree>${SP_TREE_HEAD}${shapes}</p:spTree></p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`
}

/**
 * 组装 PPTX。
 * @param options - `slides`：`{ title, bullets }`，bullets 可带 `level`。
 * @returns 可直接写盘的 Buffer。
 */
export function buildPptx({ slides }) {
  const normalized = slides.map((slide) => ({
    title: slide.title === undefined ? '' : String(slide.title),
    bullets: Array.isArray(slide.bullets) ? slide.bullets : [],
  }))
  if (normalized.length === 0) throw new Error('slides 不能为空')

  const [pkgRels, types] = packageParts({
    mainPart: 'ppt/presentation.xml',
    mainType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
    overrides: [
      ['ppt/slideMasters/slideMaster1.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'],
      ['ppt/slideLayouts/slideLayout1.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'],
      ...normalized.map((_, index) => [
        `ppt/slides/slide${index + 1}.xml`,
        'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
      ]),
      ['ppt/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'],
    ],
  })

  const themeRelId = `rId${normalized.length + 2}`
  const presentation = `${XML_HEAD}
<p:presentation xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}" saveSubsetFonts="1">
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
<p:sldIdLst>${normalized
    .map((_, index) => `<p:sldId id="${256 + index}" r:id="rId${index + 2}"/>`)
    .join('')}</p:sldIdLst>
<p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>
<p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`

  return zip([
    types,
    pkgRels,
    ['ppt/presentation.xml', presentation],
    ['ppt/_rels/presentation.xml.rels', rels([
      ['rId1', `${R}/slideMaster`, 'slideMasters/slideMaster1.xml'],
      ...normalized.map((_, index) => [`rId${index + 2}`, `${R}/slide`, `slides/slide${index + 1}.xml`]),
      [themeRelId, `${R}/theme`, 'theme/theme1.xml'],
    ])],
    ['ppt/theme/theme1.xml', THEME],
    ['ppt/slideMasters/slideMaster1.xml', SLIDE_MASTER],
    ['ppt/slideMasters/_rels/slideMaster1.xml.rels', rels([
      ['rId1', `${R}/slideLayout`, '../slideLayouts/slideLayout1.xml'],
      ['rId2', `${R}/theme`, '../theme/theme1.xml'],
    ])],
    ['ppt/slideLayouts/slideLayout1.xml', SLIDE_LAYOUT],
    ['ppt/slideLayouts/_rels/slideLayout1.xml.rels', rels([
      ['rId1', `${R}/slideMaster`, '../slideMasters/slideMaster1.xml'],
    ])],
    ...normalized.map((slide, index) => [`ppt/slides/slide${index + 1}.xml`, slideXml(slide)]),
    ...normalized.map((_, index) => [`ppt/slides/_rels/slide${index + 1}.xml.rels`, rels([
      ['rId1', `${R}/slideLayout`, '../slideLayouts/slideLayout1.xml'],
    ])]),
  ])
}
