# dsh-export-office-noshell

生成 Word / Excel / PPT / PDF / 图片的工具集（`export_docx_noshell` / `export_xlsx_noshell` / `export_pptx_noshell` / `export_pdf_noshell` / `export_image_noshell`），零 npm 依赖，只把文件写进调用方会话工作区的 `.ai-output/` 目录。

## 技术选型

| 选型 | 方案 | 为什么 |
|---|---|---|
| Word / Excel / PPT 生成 | 手拼 OOXML（本质是「一批 XML 打包成 zip」），用 Node 内置 `zlib` 完成 | 零依赖，不引第三方库 |
| zip 打包 | 自写极简 zip（逐条 deflate level 9 + CRC32） | 同上 |
| PDF 生成 | 先拼 docx，再交本机 LibreOffice 无头转一道 | PDF 里中文要嵌字体、切子集、写 CID 映射，自写成本远高于转一道 |
| 图片生成 | 把 HTML 落成临时文件，交本机无头 Chrome / Edge（`--headless=new --screenshot`）截一张 | HTML + CSS 就是一套完整排版语言，浏览器是现成渲染器；自写位图字体、渐变、阴影不划算 |
| 运行环境 | 纯 JS（ESM），跑在插件进程 | 模型侧不需要 shell / Python / 装 Office；外部程序由插件自己拉子进程 |

## 核心流程（动作清单）

### 阶段一：装载（插件启动，同步）

1. 拿到插件配置，逐项校验：`maxItems` 必须正数、`outputDir` 必须是工作区内相对目录且不含 `..`、`convertTimeoutMs` / `shotTimeoutMs` 必须正数
2. 给缺的字段填默认值：`root` 空、`outputDir` `.ai-output`、`maxItems` 20000、`soffice` 空、`convertTimeoutMs` 120000、`chrome` 空、`shotTimeoutMs` 60000
3. 把 5 个工具（`export_docx_noshell` / `export_xlsx_noshell` / `export_pptx_noshell` / `export_pdf_noshell` / `export_image_noshell`）注册进 host 的 `tools` 表
4. 交回给「调用」

### 阶段二：公共前置（5 个工具都先走这 5 步）

5. 取出调用方会话的工作区路径；拿不到就退回配置里的兜底 `root`
6. 校验 `fileName`：必须形如「文件名.扩展名」，不含路径分隔符
7. 拼出落盘路径 `<工作区>/<outputDir>/<fileName>`
8. 算目标相对目录的相对路径，出现 `..` 或绝对路径就拒绝（越界防护，写不出工作区）
9. 交回给各自的「组装」

### 阶段三：组装 docx（export_docx_noshell）

10. 判路：带 `blocks` 就是「交一批」，带 `final: true` 就是「收尾」—— 只有这一条路，没有「一次性 / 分批」两套
11. 校验 `part`：带 `blocks` 就必须给，且是 1–9999 的整数（批次号，从 1 开始）；没给直接报错
12. 按「会话 id + 文件名」算出这篇文档的暂存目录 `<工作区>/<outputDir>/.staging/docx/<哈希>/`
13. `reset = true` 就把这个暂存目录整个删掉（推倒重写）
14. 校验上限：批次数 ≤ 200、留下的批次块数 + 本批块数 ≤ `maxItems`（块数走 `index.json`，不重读历史批次）
15. 把本批写成 `part-000N.json`（先写 `.tmp` 再 rename，同号覆盖 = 重传幂等），顺手记一份 `meta.json` 标出会话与文件名，并把本批块数增量写进 `index.json`
16. 量一下本批体量：超过 20 块或 8000 字符就记一句 warning（照收，只提醒下一批写小些 —— 实测 1.6 万字符才开始写坏）
17. 不是收尾（`final` 不为 true）就到此为止：返回已存批次号、块数、`nextPart`、warning，**不出文件**
18. 收尾时按批次号排序，缺号就报「缺第 N 批」并保留暂存 —— 不生成、不丢内容
19. 合并全部批次的 `blocks`；`title` 取本次给的，没给就用第一个带标题的批次
20. 有 `title` 就先拼一个标题段（Title 样式：加粗、48 半磅）
21. 逐块按类型分派：`heading` 级别 clamp 到 1–3 套 Heading 样式；`bullet` 文本前加「• 」套列表样式；`table` 按列数平分列宽、表头加粗；`image` 走图片四件套；其余当普通段落
22. 末尾补一个空段落 + 页面尺寸（A4：宽 11906、高 16838，页边距 1440）
23. 拼出正文、样式表、关系、内容类型这 4 份 XML
24. 全部条目打成 zip（逐条 deflate 压到 level 9，压不小就存原文，算 CRC32）
25. 交回给「写盘」
26. 写盘成功才删暂存目录（part 文件 + `index.json` + `meta.json` 一起清掉）；写盘失败保留暂存，`final` 可以直接重试
27. 收尾重试（暂存已空、文件已在）不重造：返回既有产出，并写明「本次没有重新生成」

### 阶段四：组装 xlsx（export_xlsx_noshell）

28. 拿到 `sheets` 工作表列表，空数组直接报错
29. 累计各表数据行总数，超 `maxItems` 就报错
30. 规范化每表：`name` 截到 31 字（默认 SheetN）、`columns` / `rows` / `numberColumns` 转字符串
31. 校验每表必须有表头（`columns`），缺就报错
32. 逐表拼 XML：按列算显示宽度（取该列最长值，CJK 算 2 列，+2 后 clamp 到 9–40）；单元格按 `=` 开头写公式、命中 `numberColumns` 且是数字写数字、空写空、其余写文本；表头行加粗
33. 拼出工作簿、样式表、关系、每表 XML
34. 全部条目打成 zip
35. 交回给「写盘」

### 阶段五：组装 pptx（export_pptx_noshell）

36. 拿到 `slides` 页列表，空数组直接报错
37. 累计各页项目符号总数，超 `maxItems` 就报错
38. 拼固定空白母版 + 空白布局（16:9，页面 12192000 × 6858000）
39. 逐页拼 XML：`title` 进标题框（加粗 36 号），`bullets` 进正文框（20 号），每条加「•」，`level` 决定缩进（0–4 级，每级缩 342900）
40. 拼出演示主文档、主题、母版、布局、每页及各自关系
41. 全部条目打成 zip
42. 交回给「写盘」

### 阶段六：写盘（docx / xlsx / pptx 共用）

43. 递归建输出目录（已有则跳过）
44. 把 zip 字节流同步写进目标文件
45. 读出写入后的文件大小（字节）
46. 返回：绝对路径 + 相对路径 + 字节数 + 数量（块数 / 表数 / 页数）
47. 附交付提醒：先调 `present` 让 GUI 可见，再调 `dsh_im_return_file` 登记 IM 投递

### 阶段七：出图（export_image_noshell）

48. 拿到 `html`，空白直接报错
49. 定 `format`（默认 png），走自己的落盘校验：文件名必须形如「文件名.png」；`format=jpeg` 时同时认 `.jpg` 与 `.jpeg`（统一按 `.jpg` 落盘），非法文件名在跑浏览器之前就挡掉
50. 找浏览器：优先用配置的 `chrome`，其次 `CHROME_PATH` 环境变量，再按候选位置（Chrome 两处 + Edge 两处 + mac / Linux 若干）探测，找不到就报错
51. 整理 HTML：已是完整文档（含 `<html`）就原样用；只是片段就套一层默认壳（UTF-8、微软雅黑字体栈、白底）；给了 `selector` 再注入两段 CSS —— body 撑满 + 非命中元素 `display:none`，让截图只框住那个元素
52. 收敛数值参数：`width` / `height`（1–10000，默认 1200×800）、`scale`（1–4，默认 2）、`waitMs`（0–30000，默认 2000）、`quality`（1–100，默认 92）
53. 建临时目录，把 HTML 写成 `page.html`
54. 同步调浏览器截图：`--headless=new --screenshot=<临时目录>/shot.<格式> --window-size=<宽>,<高> --user-data-dir=<临时目录>/profile`，关掉首启弹窗 / 同步 / 扩展 / 滚动条，锁 sRGB；给了 `selector` 或 `fullPage=false` 就不加 `--screenshot-full-page`；`format=jpeg` 再加 `--screenshot-format=jpeg --screenshot-quality=<画质>`；`scale>1` 加 `--force-device-scale-factor`
55. 按 `allowNetwork` 分流：默认允许联网，加 `--virtual-time-budget=<waitMs>` 等 CDN 图表加载完再截；设 `false` 则加 `--disable-features=NetworkService --host-resolver-rules=MAP * ~NOTFOUND` 断网截，页面只能用内联资源
56. 开不了管道就退回「不捕获输出」重试一次
57. 以产出文件为准：临时目录里有图就认（不信任退出码，headless 截完图有时仍报非 0）；超时单独给一句提示
58. 把产出图搬进目标路径（跨盘 rename 失败就退回复制 + 删源）
59. 返回：绝对路径 + 相对路径 + 字节数 + 格式 + 视口尺寸 + 是否整页 + 命中选择器 + 浏览器路径
60. 清掉临时目录（清不掉不影响结果）
61. 附交付提醒

### 阶段八：转 PDF（export_pdf_noshell）

62. 拿到 `source` 或 `blocks`，两者都空直接报错（二选一）
63. 校验 `source`：只能是 `.ai-output/` 下的 `.docx` / `.xlsx` / `.pptx` 文件名
64. 走公共前置，算出目标 `.pdf` 路径
65. 找 LibreOffice：优先用配置的 `soffice`，否则按候选位置（Windows 两个 + Linux 两个）探测，找不到就报错
66. 建独立临时目录（系统临时目录下，带进程号 + 时间戳）
67. 准备源文件：给了 `source` 就取已有文件；给了 `blocks` 就先生成一个临时 docx 进临时目录
68. 同步调 soffice 无头转 PDF（`--headless --convert-to pdf`，单次超时 `convertTimeoutMs`，默认 120 秒）
69. 开不了管道就退回「不捕获输出」重试一次
70. 以产出文件为准：扫临时目录找 `.pdf`，没有就报错（不信任退出码）
71. 把产出 pdf 搬进目标路径（跨盘 rename 失败就退回复制 + 删源）
72. 返回：绝对路径 + 相对路径 + 字节数 + 来源格式 + soffice 路径
73. 清掉临时目录（清不掉不影响结果）
74. 附交付提醒

## 工具与入参

| 工具 | 必填 | 可选 | 产出 |
|---|---|---|---|
| `export_docx_noshell` | `fileName` | `blocks`、`part`、`title`、`final` / `reset` | Word，块按顺序排：heading / paragraph / bullet / table / **image**；一次一批（15–20 块），最后一批 `final: true` 合并成文 |
| `export_xlsx_noshell` | `fileName`、`sheets` | 每表 `name`、`numberColumns`、**`images`**、**`charts`** | Excel，多工作表；`=公式` 打开自动重算；**可嵌图片、可插原生图表** |
| `export_pptx_noshell` | `fileName`、`slides` | 每页 `bullets` 的 `level` | PowerPoint，每页 title + bullets（支持缩进层级） |
| `export_pdf_noshell` | `fileName`，`blocks` 或 `source` 二选一 | `title`（仅 `blocks` 生效） | PDF，靠本机 LibreOffice 无头转；**blocks 里可带 image 块，图片会真嵌进 PDF** |
| `export_image_noshell` | `fileName`、`html` | `width`、`height`、`selector`、`fullPage`、`scale`、`waitMs`、`allowNetwork`、`format`、`quality` | PNG / JPEG，靠本机无头 Chrome / Edge 截图；可用完整 CSS 与 CDN 图表库 |

### docx 分批提交（`part` / `final` / `reset`）

`export_docx_noshell` 只有一条路：**一次调用交一批**，最后一批带 `final: true` 收尾。

| 字段 | 说明 |
|---|---|
| `blocks` | 本批内容块，按顺序排 |
| `part` | 批次号，从 1 开始；**带 `blocks` 就必须给**（批次号就是暂存文件名，同号重传 = 覆盖，所以重试幂等） |
| `final` | `true` = 收尾：按 `part` 顺序合并全部批次，一次性生成 docx；可以不带 `blocks`，重复调用安全 |
| `reset` | `true` = 先清空这篇文档已存的批次；单独给 `reset: true` 就是「清空重来」 |

调用范式：

1. 小文档：`{ fileName, part: 1, final: true, title, blocks: [...] }` —— 一次调用成文
2. 大文档：`{ fileName, part: 1, blocks: [...] }` → `{ saved: true, nextPart: 2 }`，**不出文件**
3. 接着 `{ fileName, part: 2, blocks: [...] }` → `nextPart: 3`，最后一轮 `{ fileName, final: true }`

一批写多少（给模型的硬规则）：

| 上限 | 值 | 依据 |
|---|---|---|
| 块数 | 15–20 块（一段 / 一条 / 一张表各算 1 块） | 模型能边写边数，比数字数可靠 |
| 正文 | 1000–2500 字 | 20 个块的自然长度 |
| 参数 JSON | ≤ 8000 字符 | 实测 1.6 万字符开始写坏（16K 以下 0 失败，16–24K 约 7% 崩），取一半留余量 |

超上限不拦（内容已经写出来了），只在返回值里带一句 `warning` 让下一批写小些。

其余要点：

- **漏号、失败都不丢内容**：缺第 2 批时 `final` 报「缺第 2 批」并保留已存批次，补传 `part: 2` 再 final 即可。
- **暂存在哪**：`<工作区>/<outputDir>/.staging/docx/<哈希>/`，哈希由「会话 id + 文件名」算出 —— 同一会话里两个文档、两个会话里同名文档都不会串味。目录里另有 `meta.json`（标出哪个会话的哪个文件）与 `index.json`（每批多少块的缓存，落批只读它）。成文成功后**整个目录一次删掉**；没成文就一直留着（要清就 `reset: true`）。
- **PDF 不参与分批**：要分批就先 `export_docx_noshell` 分批成文，再用 `export_pdf_noshell` 的 `source` 转过去。


### image 块（export_docx_noshell / export_pdf_noshell 共用）

| 字段 | 必填 | 说明 |
|---|---|---|
| `path` | 二选一 | `.ai-output/` 下已有的图片文件名（`.png` / `.jpg` / `.jpeg` / `.gif`），例如 `客户分组统计.png` |
| `data` | 二选一 | 图片内容的 base64，可带 `data:image/png;base64,` 前缀 |
| `width` | | 显示宽度（像素），默认按版心 90% 排 |
| `widthPct` | | 显示宽度占版心百分比（1–100）；给了 `width` 就以 `width` 为准 |

图片以**实际字节**判格式（不看扩展名），尺寸从文件头读；读不出尺寸就拒绝 —— 读不出就没法按比例排版。高度上限 620px，超出按比例回缩，免得一张长图顶爆版面。

### Excel 的 images 与 charts（每张工作表各自给）

`images[]` 字段：

| 字段 | 必填 | 说明 |
|---|---|---|
| `path` | 二选一 | `.ai-output/` 下已有的图片文件名 |
| `data` | 二选一 | 图片 base64 |
| `anchor` | | 左上角落在哪个单元格，默认 `A1`，例如 `"F2"` |
| `width` / `height` | | 显示尺寸（像素），默认 420 × 300 |

`charts[]` 字段：

| 字段 | 必填 | 说明 |
|---|---|---|
| `type` | | `bar`（柱状）/ `line`（折线）/ `pie`（饼图），默认 `bar` |
| `title` | | 图表标题 |
| `categories` | | 分类列的表头文字，默认第一列 |
| `series` | | 要画成系列的列名列表，默认「除分类列外的所有列」 |
| `anchor` | | 图表左上角落在哪个单元格，默认 `A2` |
| `width` / `height` | | 图表尺寸（像素），默认 720 × 420 |

图表是**原生图表对象**，不是贴图：Excel 打开后能改类型、换数据源、跟随公式重算。数据取本表（第一行为表头），所以图表的列名必须能在本表表头里找到，找不到直接报错、不猜。

## 配置项

| 键 | 默认 | 说明 |
|---|---|---|
| `root` | 空 | 兜底工作区根目录；正常无需配置，优先用调用方会话工作区 |
| `outputDir` | `.ai-output` | 输出目录（相对 root） |
| `maxItems` | 20000 | 单文件单元格 / 段落 / 项目符号总数上限 |
| `soffice` | 空 | LibreOffice 可执行文件路径，留空自动探测 |
| `convertTimeoutMs` | 120000 | 单次 PDF 转换超时（毫秒） |
| `chrome` | 空 | Chrome / Edge 可执行文件路径，留空自动探测（也认 `CHROME_PATH`） |
| `shotTimeoutMs` | 60000 | 单次 HTML 截图超时（毫秒） |

## 已知环境问题（不是插件的 bug）

### 挂 VPN 时导出 PDF 会弹「正在等待打印机连接」

| 项 | 说明 |
|---|---|
| 现象 | 挂着 VPN 时导出 PDF，Windows 弹「正在等待打印机连接或取消连接」 |
| 真因 | 本机默认打印机是一台 **WSD 端口**的 HP（WSD 靠网络发现定位设备）。VPN 一连上就改写路由、掐掉本地组播，Windows 发现不到打印机；而 **LibreOffice 启动时会枚举系统打印机**（即便 `--headless` 也照样枚举），一碰那台失联的打印机，驱动就进入等待状态并弹框 |
| 已验证无效的做法 | 给 LibreOffice 换**一次性隔离的用户配置目录**（`-env:UserInstallation`）—— 试过，照样弹。换配置目录拦不住「枚举打印机」这个动作本身 |
| 系统侧的对策 | ① 把该打印机设为「脱机使用打印机」；② 或把默认打印机改成 `Microsoft Print to PDF`。两者都能让 Windows 不再尝试连接它 |
| 与插件无关 | 插件全程没有任何打印调用，导出走的是 `soffice --convert-to pdf` 格式转换 |

### 临时目录会自清理

`makeScratch()` 每次顺手扫掉 `dsh-export-office-noshell-` 前缀下、且**超过一小时**的旧工作目录 ——
转换失败或进程被强杀时 `cleanScratch` 删不掉，靠这道兜底。只删旧的，不误伤正在跑的另一次转换。

## 关键设计点

- **输出目录绑定会话工作区**：优先取 `cwd`，不硬编码全局 root —— profile 内多会话共用一个插件实例，硬编码会把所有会话的产出塞进同一个目录。
- **文件名即安全边界**：`fileName` 只能是不带路径的文件名，`..` / 绝对路径 / 路径分隔符一律拒绝，保证写不出工作区。图片因为要认 `.jpg` / `.jpeg` 两种写法，自己走一份等价校验，且必须在拉起浏览器之前判完，免得白截一张。
- **单元格三类判定**：`=` 开头 → 公式（`<f>`，打开自动重算）；命中 `numberColumns` 且是数字 → 数字（`<v>`）；其余 → 文本。
- **PDF / 图片以产出为准**：不信任外部程序退出码（headless Chrome 截完图有时仍报非 0，soffice 也会），只认临时目录里是否真出了文件；EPERM（开不了管道）退回不捕获输出重试；EXDEV（跨盘）退回复制 + 删源。
- **出图借用浏览器而不是自写渲染**：HTML + CSS 直接换来渐变、阴影、圆角、flex / grid、内联 SVG 与中文字体，代价只是依赖本机装了 Chrome / Edge；插件自己拉子进程，模型侧依旧不需要 shell。
- **图片是「嵌进包」而不是「引用路径」**：字节进 `word/media/`（Excel 是 `xl/media/`），正文放 `<w:drawing>`（Excel 是 `xl/drawings/drawingN.xml` 里的锚点），关系表补 image 关系，内容类型补图片默认类型 —— 四件套齐了 Office 才认。PDF 因此也能带图：它就是这份 docx 交给 LibreOffice 转的，图随文档一起走，不依赖原图还在不在。
- **Excel 图表是原生对象**：`xl/charts/chartN.xml` 里写成 `<c:barChart>` / `<c:lineChart>` / `<c:pieChart>`，配 `<c:catAx>` / `<c:valAx>` 与逐系列的数据缓存。缓存不是可选装饰 —— 它让图表在**没重算之前**也能显示正确的柱子；`<c:f>` 公式则保证用户改了数据后 Excel 会跟着重算。写缓存时行号必须和公式一致（第一行数据落在 `headerRow + 1`），差一行就会出现「图里有柱子但数字对不上」的假象。
- **图片格式看字节不看扩展名**：`path` 只是入口，真实格式与尺寸都由 `imageSpec()` 从文件头读（png / jpeg / gif 三种）。读不出尺寸直接拒 —— 否则没法按比例排版，排出来是变形的。
- **工具参数是深冻结的，只能读不能写**：运行时派发前会做 `deepFreeze(detached)`，往 `args` 或里面的块上写字段会直接抛 `Cannot add property …, object is not extensible`（严格模式下冻结对象拒绝加属性）。所以块与图片的对应关系用下标数组 `mediaIndex` 旁路传递，**不要**把信息写回 block。测试脚本同样必须先 `deepFreeze` 再调工具，否则这类错测不出来。
- **联网是显式开关**：默认允许联网，好让 `<script src="https://cdn…">` 这种 CDN 图表库能加载（配 `--virtual-time-budget=<waitMs>` 等它渲染完）；要离线可复现就把 `allowNetwork` 设 `false`，网络组件会被整个关掉。
- **PDF 转换沿用一份共享的 LibreOffice 配置**（`%TEMP%\dsh-export-office-noshell-loprofile`）。曾为规避「等待打印机连接」弹框改成一次性隔离配置，**实测无效**且增加了临时文件负担，已回退 —— 那个弹框的真因见上面「已知环境问题」。
- **临时目录会自清理**：转换失败或进程被强杀时 `cleanScratch` 删不掉，配置文件就留在 TEMP 里（每个约 280 KB）。`makeScratch()` 每次顺手扫掉**自己前缀**下、且**超过一小时**的旧目录 —— 只删旧的，不误伤正在跑的另一次转换，也不碰别家的目录。
- **DrawingML 用 oneCellAnchor，不写 twoCellAnchor**：只锚一个角（左上单元格 + 像素偏移），尺寸由 `<xdr:ext>` 直给，不依赖目标区域的列宽行高。twoCellAnchor 要求写两个**绝对坐标**，算错一边就把图表压成几十像素的一小块（踩过：写成「起点 + 1 格」，用户实测图表小得像个点）。
- **docx 分批提交**：模型一次生成的 JSON 越长越容易坏（实测工具参数超 1.6 万字符开始写坏），所以 `export_docx_noshell` 只有一条路 —— 一次调用交一批（15–20 块），最后 `final: true` 合并成文。批次号即暂存文件名（`part-0001.json`），**同号覆盖 → 重试幂等**；缺号拒绝成文并保留暂存 → 单批失败不牵连其他批次；成文成功才清暂存 → 写盘失败可直接重试 `final`。**故意不做自动编号**（重试会给出新号，正好造出重复内容），**也不按字数下指令**（模型数字数不准，按「块数」下才有约束力）。
- **暂存索引 `index.json`**：落批时只需要「存了哪几批、一共多少块」，块数就记在暂存目录的 `index.json` 里，不再重读历史批次 —— 原先每落一批重读全部历史，是 O(n²)（实测 200 批 37 秒，改后 0.7 秒）。索引与目录对不上（丢了 / 手删了 part 文件 / 并发写坏）就当场按文件重建，所以它只是缓存、不是真相；收尾时随暂存目录一起删。
- **交付两步**：写盘 ≠ 用户可见，还要 `present`（GUI / 客户端）+ `dsh_im_return_file`（IM 渠道投递）。
- **返回值里绝不能出现 `undefined`**：运行时在交付结果前会跑一次 lossless-JSON 快照（`packages/util/values` 的 `snapshotJsonValue`），`undefined` 不是合法 JSON 值 —— 多写一个 `undefined` 就会让整次调用以 `value is not lossless JSON` 失败，**哪怕文件已经写好、也照常交付了**。要「可选字段」就用条件展开 `...(x ? { k: x } : {})`，别写 `k: undefined`；`OUTPUT.render` 另有一道 `dropUndefined` 兜底。
- **列宽估算法**：CJK 字符算 2 列宽，最长值 +2 后 clamp 到 9–40，避免中文挤成一列。

## 交付

1. `present` —— 让 GUI / 客户端看到文件；
2. `dsh_im_return_file` —— IM 渠道（飞书 / 企微 / 钉钉…）投递，对同一条绝对路径再登记一次，否则 IM 端收不到。

## 验收

```bash
node selftest.mjs "<输出目录>"          # 33 项：docx / xlsx / pptx / pdf / png
node e2e-image.mjs "<输出目录>"         # 7 项：真装插件 → 调 export_image_noshell → ECharts CDN 出图 + lossless 回归
node test-image-embed.mjs "<输出目录>"  # 21 项：image 块嵌入 docx 的结构 + 转 PDF 后图还在不在
node test-xlsx-chart.mjs "<输出目录>"   # 33 项：Excel 嵌图片 + 原生图表（结构 / 校验 / 转 PDF 验矢量）
node test-docx-stage.mjs "<输出目录>"    # 40 项：一次成文 / 分批 / 重传幂等 / 缺号 / 乱序 / 收尾重试 / warning / reset / 会话内隔离 / index.json 自愈
```

分工：`selftest.mjs` 直接打 `lib/` 里的构建函数（结构层面）；`e2e-image.mjs` 走「插件注册 → 工具 execute」的真实链路，
顺带验证 `wrapHtml` 套壳、非法文件名拦截、CDN 图表加载，以及返回值过不过得了运行时的 lossless-JSON 快照；
`test-image-embed.mjs` 专验图片嵌入（拆 docx 比对媒体字节 / 关系 ID / 内容类型，再用 LibreOffice 转 PDF 查图像对象）；
`test-xlsx-chart.mjs` 专验 Excel 的图片与图表（拆 xlsx 比对 media / drawing / chart 三类部件与关系，解压 PDF 内容流数绘图算子）。
`test-docx-stage.mjs` 专验分段导出（真实链路调 `export_docx_noshell.execute`，不打外部程序）。

两节里的 `pdf` / `image` 部分要真拉外部程序（LibreOffice / 浏览器）：
**在禁止创建命名管道的受限沙箱里会失败**（Chrome 报 `mojo platform_channel.cc:108 Check failed: 拒绝访问 (0x5)`），
那是环境限制、不是代码问题 —— 换正常档位再跑即可。

## 生效与回滚

- 改动需重启 `dsh web`（插件在启动时加载）。
- 备份：`index.js.bak-<时间戳>`、`cordis.patch.yml.bak-<时间戳>`。
