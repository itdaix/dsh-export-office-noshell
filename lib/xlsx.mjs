/**
 * XLSX 生成器：表头 + 数据行 → 真正的 Excel 工作簿。
 *
 * 除了数据，还支持往工作表里放**图片**与**原生图表**：两者都走 DrawingML
 * （`xl/drawings/drawingN.xml`），图片另需 `xl/media/`，图表另需 `xl/charts/chartN.xml`。
 * 图表是 Excel 原生图表对象（不是贴图）—— 打开后能改类型、换数据源、跟随公式重算。
 *
 * @module lib/xlsx
 */

import { XML_HEAD, escapeXml, packageParts, widthOf, zip } from './zip.mjs'

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const NS_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'
const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'

/** 1 像素 = 9525 EMU（96 dpi）。 */
const EMU_PER_PX = 9525
/** 默认列宽（字符数）→ 像素，以及默认行高（磅）→ 像素。 */
const DEFAULT_COL_WIDTH_CHARS = 9
const DEFAULT_ROW_HEIGHT_PT = 15
const COL_PX_PER_CHAR = 7
const ROW_PX_PER_PT = 4 / 3
/** 图片与图表的默认尺寸（像素）。 */
const DEFAULT_IMAGE_PX = { width: 420, height: 300 }
const DEFAULT_CHART_PX = { width: 720, height: 420 }

/** 0 → A，25 → Z，26 → AA。 */
function columnName(index) {
  let n = index
  let name = ''
  do {
    name = String.fromCharCode(65 + (n % 26)) + name
    n = Math.floor(n / 26) - 1
  } while (n >= 0)
  return name
}

/** `B3` → `{ col: 1, row: 2 }`（0 起）。认不出来回退到 A1。 */
function parseCellRef(ref) {
  const match = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(String(ref ?? '').trim())
  if (!match) return { col: 0, row: 0 }
  const letters = match[1].toUpperCase()
  let col = 0
  for (const ch of letters) col = col * 26 + (ch.charCodeAt(0) - 64)
  return { col: col - 1, row: Number(match[2]) - 1 }
}

/** 列宽（字符）→ 像素，跟 Excel 的换算式对齐。 */
function columnPx(chars) {
  return Math.round(chars * COL_PX_PER_CHAR + 5)
}

/** 按表头内容估算每列宽度（字符数），与写入 `<col>` 时的算法保持一致。 */
function columnWidths(sheet) {
  const grid = [sheet.columns, ...sheet.rows]
  return sheet.columns.map((_, ci) => {
    const longest = grid.reduce((w, row) => Math.max(w, widthOf(row[ci] ?? '')), 0)
    return Math.min(Math.max(longest + 2, DEFAULT_COL_WIDTH_CHARS), 40)
  })
}

/** 从锚点单元格算出起点偏移量。 */
function anchorOffset(anchor, widths) {
  const { col, row } = parseCellRef(anchor)
  let x = 0
  for (let i = 0; i < col; i += 1) x += columnPx(widths[i] ?? DEFAULT_COL_WIDTH_CHARS)
  const y = Math.round(row * DEFAULT_ROW_HEIGHT_PT * ROW_PX_PER_PT)
  return { col, row, x, y }
}

/* ─────────────────────────── 单元格 ─────────────────────────── */

const NUMBER = /^-?\d+(\.\d+)?([eE][+-]?\d+)?$/

/** `=` 开头 → 公式；命中数字列 → 数字；其余按文本。 */
function cell(ref, value, style, asNumber) {
  const attr = style ? ` s="${style}"` : ''
  const text = value === null || value === undefined ? '' : String(value)
  if (text.startsWith('=') && text.length > 1) {
    return `<c r="${ref}"${attr}><f>${escapeXml(text.slice(1))}</f></c>`
  }
  if (asNumber && NUMBER.test(text.trim())) {
    return `<c r="${ref}"${attr}><v>${text.trim()}</v></c>`
  }
  if (text === '') return `<c r="${ref}"${attr}/>`
  return `<c r="${ref}"${attr} t="inlineStr"><is><t xml:space="preserve">${escapeXml(text)}</t></is></c>`
}

/* ─────────────────────── 图表：定义与 XML ─────────────────────── */

/** 图表类型 → 中文名（写进报错，也用于默认标题）。 */
const CHART_KINDS = { bar: '柱状图', line: '折线图', pie: '饼图' }

/**
 * 把表头名字解析成列下标。
 * @returns 列下标，找不到是 -1。
 */
function columnIndexOf(sheet, name) {
  const wanted = String(name ?? '').trim()
  if (wanted === '') return -1
  const direct = sheet.columns.findIndex((column) => String(column).trim() === wanted)
  if (direct >= 0) return direct
  const upper = wanted.toUpperCase()
  const asLetter = sheet.columns.findIndex((column, index) => columnName(index) === upper)
  return asLetter
}

/** 一个系列（`<c:ser>`）：含数据缓存与名称。 */
function chartSeries({ categoriesIndex, seriesIndex, values, categories, sheetName, headerRow, lastRow }) {
  const column = columnName(seriesIndex)
  const categoryColumn = columnName(categoriesIndex)
  // 三处引用各就各位：系列名取表头那格；公式指向数据区；缓存里的行号必须与公式一致 ——
  // 第一份数据行是 headerRow + 1，缓存的第 i 项就落在 `${column}${headerRow + 1 + i}`。
  const nameRef = `${escapeXml(sheetName)}!$${column}$${headerRow}`
  const catRef = `${escapeXml(sheetName)}!$${categoryColumn}$${headerRow + 1}:$${categoryColumn}$${lastRow}`
  const valRef = `${escapeXml(sheetName)}!$${column}$${headerRow + 1}:$${column}$${lastRow}`
  const catCache = categories
    .map((value, i) => `<c:pt idx="${i}"><c:v>${escapeXml(value)}</c:v></c:pt>`)
    .join('')
  const valCache = values
    .map((value, i) => `<c:pt idx="${i}"><c:v>${value}</c:v></c:pt>`)
    .join('')
  return {
    nameRef,
    catRef,
    valRef,
    catCache,
    valCache,
    categoryCount: categories.length,
    valueCount: values.length,
  }
}

/** 生成一张图的 `chartN.xml`。 */
function chartXml({ kind, title, series, colors }) {
  const serXml = series
    .map((item, index) => {
      const color = colors[index % colors.length]
      const cat = `<c:cat><c:strRef><c:f>${item.catRef}</c:f><c:strCache><c:ptCount val="${item.categoryCount}"/>${item.catCache}</c:strCache></c:strRef></c:cat>`
      const val = `<c:val><c:numRef><c:f>${item.valRef}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${item.valueCount}"/>${item.valCache}</c:numCache></c:numRef></c:val>`
      const name = `<c:tx><c:strRef><c:f>${item.nameRef}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escapeXml(item.name ?? `系列${index + 1}`)}</c:v></c:pt></c:strCache></c:strRef></c:tx>`
      const fill = `<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr>`
      if (kind === 'pie') {
        // 饼图不吃逐点填色，交给 Excel 的默认配色；只保证数据标签可见。
        return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>${name}${cat}${val}`
          + `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/></c:dLbls>`
          + `</c:ser>`
      }
      return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>${name}${fill}${cat}${val}`
        + `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/></c:dLbls>`
        + `</c:ser>`
    })
    .join('')

  const titleXml = title
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="1"/></a:pPr><a:r><a:t>${escapeXml(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
    : '<c:autoTitleDeleted val="1"/>'

  if (kind === 'pie') {
    return `${XML_HEAD}
<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}">
<c:chart>${titleXml}<c:plotArea><c:layout/>
<c:pieChart><c:varyColors val="1"/>${serXml}<c:firstSliceAng val="0"/></c:pieChart>
</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>
</c:chartSpace>`
  }

  const barChart = kind === 'bar'
  const plot = barChart
    ? `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${serXml}<c:gapWidth val="120"/><c:axId val="111111111"/><c:axId val="222222222"/></c:barChart>`
    : `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${serXml}<c:marker val="1"/><c:axId val="111111111"/><c:axId val="222222222"/></c:lineChart>`

  return `${XML_HEAD}
<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}">
<c:chart>${titleXml}<c:plotArea><c:layout/>
${plot}
<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:catAx>
<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="111111111"/></c:valAx>
</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>
</c:chartSpace>`
}

/** 图表默认配色（跟 MOM 看板常用的商务蓝灰一致）。 */
const CHART_COLORS = ['5B7DAB', '16A3BE', '9FB2C9', 'F2A65A', '7A8FA6', '4C6E91']

/**
 * 按「一块数据区」生成一张图的图表定义。
 * 第一列当分类，其余数值列各成一个系列。
 */
function buildChart(sheet, chart, sheetName) {
  const kind = CHART_KINDS[chart.type] ? chart.type : 'bar'
  const headerRow = 1
  const lastRow = headerRow + sheet.rows.length

  const categoriesIndex = chart.categories === undefined ? 0 : columnIndexOf(sheet, chart.categories)
  if (categoriesIndex < 0) throw new Error(`图表「${chart.title ?? ''}」找不到分类列「${chart.categories}」`)

  const wanted = Array.isArray(chart.series) && chart.series.length > 0
    ? chart.series
    : sheet.columns.filter((_, index) => index !== categoriesIndex).map((column) => String(column))

  const series = wanted.map((name) => {
    const index = columnIndexOf(sheet, name)
    if (index < 0) throw new Error(`图表「${chart.title ?? ''}」找不到系列列「${name}」`)
    const values = sheet.rows.map((row) => {
      const raw = row[index]
      const numeric = Number(String(raw ?? '').trim())
      return Number.isFinite(numeric) ? numeric : 0
    })
    const categories = sheet.rows.map((row) => String(row[categoriesIndex] ?? ''))
    return {
      ...chartSeries({ categoriesIndex, seriesIndex: index, values, categories, sheetName, headerRow, lastRow }),
      name: String(name),
    }
  })

  return {
    xml: chartXml({ kind, title: chart.title, series, colors: CHART_COLORS }),
    anchor: chart.anchor ?? 'A2',
    width: Number.isFinite(Number(chart.width)) && Number(chart.width) > 0 ? Number(chart.width) : DEFAULT_CHART_PX.width,
    height: Number.isFinite(Number(chart.height)) && Number(chart.height) > 0 ? Number(chart.height) : DEFAULT_CHART_PX.height,
    kind,
  }
}

/* ─────────────────────── DrawingML：图片与图表 ─────────────────────── */

/** 像素 → EMU。 */
function EMU_PX(pixels) {
  return Math.round(pixels * EMU_PER_PX)
}

/**
 * 一个「左上角 + 明确尺寸」的锚点。
 *
 * 用 oneCellAnchor 而不是 twoCellAnchor：前者只锚一个角、尺寸由 `<a:ext>` 直接给定，
 * 不依赖目标区域的列宽行高；后者要写两个绝对坐标，早先写成「起点 + 1 格」就把图表
 * 压成了几十像素的一小块（用户实测的红框）。尺寸直给，这类错从根上没了。
 */
function anchorXml({ offset, width, height }, inner) {
  return `<xdr:oneCellAnchor>`
    + `<xdr:from><xdr:col>${offset.col}</xdr:col><xdr:colOff>${EMU_PX(offset.x)}</xdr:colOff><xdr:row>${offset.row}</xdr:row><xdr:rowOff>${EMU_PX(offset.y)}</xdr:rowOff></xdr:from>`
    + `<xdr:ext cx="${EMU_PX(width)}" cy="${EMU_PX(height)}"/>`
    + inner
    + `</xdr:oneCellAnchor>`
}

/** 图片：`rId` 指向该 drawing 自己的关系。 */
function pictureAnchor({ offset, width, height, name, rId, id }) {
  const inner = `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="${escapeXml(name)}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>`
    + `<xdr:blipFill><a:blip xmlns:r="${NS_REL}" r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>`
    + `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${EMU_PX(width)}" cy="${EMU_PX(height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>`
    + `</xdr:pic><xdr:clientData/>`
  return anchorXml({ offset, width, height }, inner)
}

/** 图表：一个 graphicFrame 引用 chartN.xml，尺寸同样写死在 xfrm 上。 */
function chartAnchor({ offset, width, height, rId, id }) {
  const inner = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="图表 ${id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>`
    + `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="${EMU_PX(width)}" cy="${EMU_PX(height)}"/></xdr:xfrm>`
    + `<a:graphic><a:graphicData uri="${NS_C}"><c:chart xmlns:c="${NS_C}" xmlns:r="${NS_REL}" r:id="${rId}"/></a:graphicData></a:graphic>`
    + `</xdr:graphicFrame><xdr:clientData/>`
  return anchorXml({ offset, width, height }, inner)
}

/* ─────────────────────────── 工作表 ─────────────────────────── */

function sheetXml(sheet, drawingRelId) {
  const grid = [sheet.columns, ...sheet.rows]
  const numeric = new Set(sheet.numberColumns)

  const widths = columnWidths(sheet)
  const cols = widths.map((width, ci) => `<col min="${ci + 1}" max="${ci + 1}" width="${width}" customWidth="1"/>`).join('')

  const body = grid
    .map((row, ri) => {
      const cells = sheet.columns
        .map((column, ci) => {
          const ref = `${columnName(ci)}${ri + 1}`
          if (ri === 0) return cell(ref, row[ci] ?? column, 1, false)
          return cell(ref, row[ci], 0, numeric.has(columnName(ci)) || numeric.has(column))
        })
        .join('')
      return `<row r="${ri + 1}">${cells}</row>`
    })
    .join('')

  const drawing = drawingRelId ? `<drawing r:id="${drawingRelId}"/>` : ''
  const namespaces = `xmlns="${NS_MAIN}"${drawing ? ` xmlns:r="${NS_REL}"` : ''}`

  return `${XML_HEAD}
<worksheet ${namespaces}>
<cols>${cols}</cols>
<sheetData>${body}</sheetData>${drawing}
</worksheet>`
}

const STYLES = `${XML_HEAD}
<styleSheet xmlns="${NS_MAIN}">
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="2"><border/><border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/><xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`

/**
 * 组装 XLSX。
 *
 * @param options - `sheets`：`{ name, columns, rows, numberColumns, images, charts }`。
 *   `images[i]`：`{ name, data, spec, anchor, width, height }` —— `data` 必须是 Buffer，
 *   `spec` 来自 `imageSpec()`；`charts[i]`：`{ type, title, categories, series, anchor, width, height }`。
 * @returns 可直接写盘的 Buffer。
 */
export function buildXlsx({ sheets }) {
  const normalized = sheets.map((sheet, index) => ({
    name: String(sheet.name ?? `Sheet${index + 1}`).slice(0, 31),
    columns: (sheet.columns ?? []).map((c) => String(c)),
    rows: (sheet.rows ?? []).map((row) => (Array.isArray(row) ? row : [row])),
    numberColumns: (sheet.numberColumns ?? []).map((c) => String(c)),
    images: sheet.images ?? [],
    charts: sheet.charts ?? [],
  }))
  if (normalized.some((sheet) => sheet.columns.length === 0)) {
    throw new Error('每个工作表都必须有 columns（表头）')
  }

  // 只有含图片/图表的工作表才建 drawing 部件，编号按出现顺序连续。
  let drawingCount = 0
  let chartCount = 0
  let mediaCount = 0
  const plan = normalized.map((sheet, index) => {
    const hasDrawing = sheet.images.length > 0 || sheet.charts.length > 0
    const drawing = hasDrawing ? ++drawingCount : 0
    const charts = sheet.charts.map(() => ++chartCount)
    const media = sheet.images.map(() => ++mediaCount)
    return { sheet, sheetIndex: index, drawing, charts, media, hasDrawing }
  })

  const overrides = [
    ...normalized.map((_, index) => [
      `xl/worksheets/sheet${index + 1}.xml`,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
    ]),
    ['xl/styles.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'],
    ...plan.filter((p) => p.hasDrawing).map((p) => [
      `xl/drawings/drawing${p.drawing}.xml`,
      'application/vnd.openxmlformats-officedocument.drawing+xml',
    ]),
    ...plan.flatMap((p) => p.charts.map((chartNumber) => [
      `xl/charts/chart${chartNumber}.xml`,
      'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
    ])),
  ]
  // 图片的默认内容类型（png / jpg 各一条，去重）。
  const imageDefaults = []
  for (const { sheet } of plan) {
    for (const image of sheet.images) {
      const ext = image.spec?.ext
      const mime = image.spec?.mime
      if (ext && mime && !imageDefaults.some(([e]) => e === ext)) imageDefaults.push([ext, mime])
    }
  }

  const stylesRelId = `rId${normalized.length + 1}`
  const [rels, types] = packageParts({
    mainPart: 'xl/workbook.xml',
    mainType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
    defaults: imageDefaults,
    overrides,
  })

  const parts = [
    types,
    rels,
    ['xl/workbook.xml', `${XML_HEAD}
<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">
<sheets>${normalized
      .map((sheet, index) => `<sheet name="${escapeXml(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`)
      .join('')}</sheets>
<calcPr calcId="0" fullCalcOnLoad="1"/>
</workbook>`],
    ['xl/_rels/workbook.xml.rels', `${XML_HEAD}
<Relationships xmlns="${NS_PKG}">
${normalized
      .map((_, index) => `<Relationship Id="rId${index + 1}" Type="${NS_REL}/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`)
      .join('\n')}
<Relationship Id="${stylesRelId}" Type="${NS_REL}/styles" Target="styles.xml"/>
</Relationships>`],
    ['xl/styles.xml', STYLES],
  ]

  for (const item of plan) {
    const { sheet, sheetIndex, drawing, charts, media } = item
    const widths = columnWidths(sheet)

    // 工作表：有 drawing 就补一条关系并在正文里引用它。
    parts.push([`xl/worksheets/sheet${sheetIndex + 1}.xml`, sheetXml(sheet, item.hasDrawing ? 'rId1' : '')])

    if (!item.hasDrawing) continue

    parts.push([`xl/worksheets/_rels/sheet${sheetIndex + 1}.xml.rels`, `${XML_HEAD}
<Relationships xmlns="${NS_PKG}">
<Relationship Id="rId1" Type="${NS_REL}/drawing" Target="../drawings/drawing${drawing}.xml"/>
</Relationships>`])

    const drawingRels = []
    const anchors = []
    let shapeId = 1

    for (const [i, image] of sheet.images.entries()) {
      if (!image?.spec || !Buffer.isBuffer(image?.data)) {
        throw new Error(`工作表「${sheet.name}」第 ${i + 1} 张图片缺少 spec / data（应由调用方用 imageSpec() 解析后传入）`)
      }
      const rId = `rIdImg${media[i]}`
      const name = `image${media[i]}.${image.spec.ext}`
      const width = Number.isFinite(Number(image.width)) && Number(image.width) > 0 ? Number(image.width) : DEFAULT_IMAGE_PX.width
      const height = Number.isFinite(Number(image.height)) && Number(image.height) > 0 ? Number(image.height) : DEFAULT_IMAGE_PX.height
      drawingRels.push(`<Relationship Id="${rId}" Type="${NS_REL}/image" Target="../media/${name}"/>`)
      anchors.push(pictureAnchor({ offset: anchorOffset(image.anchor ?? 'A1', widths), width, height, name, rId, id: shapeId++ }))
      parts.push([`xl/media/${name}`, image.data])
    }

    for (const [i, chart] of sheet.charts.entries()) {
      const chartNumber = charts[i]
      const rId = `rIdChart${chartNumber}`
      const built = buildChart(sheet, chart, sheet.name)
      drawingRels.push(`<Relationship Id="${rId}" Type="${NS_REL}/chart" Target="../charts/chart${chartNumber}.xml"/>`)
      anchors.push(chartAnchor({ offset: anchorOffset(built.anchor, widths), width: built.width, height: built.height, rId, id: shapeId++ }))
      parts.push([`xl/charts/chart${chartNumber}.xml`, built.xml])
    }

    parts.push([`xl/drawings/_rels/drawing${drawing}.xml.rels`, `${XML_HEAD}
<Relationships xmlns="${NS_PKG}">
${drawingRels.join('\n')}
</Relationships>`])
    parts.push([`xl/drawings/drawing${drawing}.xml`, `${XML_HEAD}
<xdr:wsDr xmlns:xdr="${NS_XDR}" xmlns:a="${NS_A}">
${anchors.join('\n')}
</xdr:wsDr>`])
  }

  return zip(parts)
}
