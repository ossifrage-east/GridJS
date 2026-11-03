# PPT 演示文稿功能实现方案

> **状态：方案稿，未改动任何代码。**
> 目标：能打开 `.ppt` / `.pptx` 两种文件，在主页面展示，呈现的内容与样式尽量与 Office PowerPoint 一致。

---

## 0. 一句话结论

复用现有 OPC/ZIP + OLE2 解析地基，新增一个 `slide` 工作区（第三类载体），走**「存档化 + 模型化 + DOM 绝对定位渲染」**三条路；与 Word 的关键差别是**模型不复用块树**（PPT 是绝对定位 + 三级样式继承，不是流式），但**工程骨架（路由 / 契约 / 工具栏档案 / 打印预览 / 滚动条 / 缩放条）全部复用**。

---

## 1. 现状盘点：哪些能直接复用，哪些必须新写

| 现有能力 | 位置 | PPT 能否复用 | 说明 |
|---|---|---|---|
| ZIP / OPC 解包 | `utils/parser/zip.ts` | ✅ 直接用 | pptx 就是 OPC 包 |
| OLE2 复合文档 | `utils/parser/cfb.ts` | ✅ 直接用 | ppt 是 CFB，流名 `PowerPoint Document` |
| 流式 XML 扫描 | `utils/parser/xmlScanner.ts` | ✅ 直接用 | 项目铁律：禁 DOMParser，大 XML 不全量建 DOM |
| Worker 卸载 | `parser.worker.ts` / `parserClient.ts` | ✅ 加分支 | 列式回包 + transferList 的通道已建好 |
| 扩展名映射 | `utils/fileName.ts` | ✅ 已预留 | `ppt/pptx → demo` 早已写进 `EXTENSION_TO_NATIVE` |
| 原生后缀 | `NATIVE_EXTENSIONS` | ✅ 已预留 | `grid / text / demo`，`.demo` 目前无工作区 |
| 存档键 | `toArchiveKey(path)` | ✅ 不用改 | 只去 `.grid`，`路演.demo` 原样保留 |
| 打印预览外壳 | `utils/printPreviewShell.ts` | ✅ 三载体共用 | 表格/文档已共用，PPT 注入「正文样式」即可 |
| 自定义滚动条 | `document/docScroller.ts` | ✅ 复用 | 缩略图 rail 纵向 + 画布横向 |
| 缩放条 / 快捷键 | `DocWorkspace` 的 Slider + Ctrl+滚轮 | ⚠️ 抽公共 | 抽成 `common/zoomBar.ts`，三载体共用 |
| 工作区路由 | `workspace/workspaceRouter.ts` | ⚠️ 二态→三态 | |
| 打开分派 | `workspace/openRouter.ts` | ⚠️ 加分支 | 删掉「演示文稿暂不支持编辑，已按表格打开」的降级 |
| 契约层 | `contract/editHost.ts` | ⚠️ 扩展 | `WorkspaceKind` 加 `'slide'` |
| 工具栏档案 | `toolbar/profile.ts` | ✅ 加一套档案 | 纯数据，改表不改算法 |
| 持久化 | `dataArchitecture/indexDB.ts` | ⚠️ 升级 | `DB_VERSION 3 → 4`，新增 `decks` + `deckmedia` 仓 |
| 块树模型 | `document/blockTree.ts` | ❌ 不复用 | **见 §3，这是本方案最关键的取舍** |
| Canvas 渲染 | `sheets/*.ts` | ❌ 不复用 | **见 §5.1，PPT 必须走 DOM** |
| 现有 ppt/pptx 解析 | `utils/parser/documentParser.ts` | ⚠️ 仅降级兜底 | 它只抽纯文本铺 A 列，无样式无位置 |

---

## 2. 「与 Office 一致」到底指什么：拆成可验收的矩阵

「一致」不拆开会变成一句无法验收的口号。按元素列出还原度目标与代价，**这就是验收表**：

| 元素 | 数据依据 | 还原度目标 | 阶段 |
|---|---|---|---|
| 页面尺寸 / 方向 | `presentation.xml` `sldSz cx/cy` | 像素级 | P0 |
| 形状位置尺寸 | `a:xfrm` 的 `off / ext`（EMU） | 像素级 | P0 |
| 旋转 / 翻转 | `a:xfrm rot / flipH / flipV` | 像素级 | P0 |
| 文本字号 | `a:rPr sz`（1/100 pt） | 像素级 | P0 |
| 字体族 | `a:latin / a:ea / a:cs` + 主题 `fontScheme` | ⚠️ 见 §8「字体缺失」 | P0/P1 |
| 粗 / 斜 / 删除线 | `b / i / strike` | 完全一致 | P0 |
| 下划线 | `u` + `a:uFill` / `a:uLn`（含线型） | 基本一致（线型 P1） | P0 |
| 文字颜色 | `a:solidFill` 或 `a:schemeClr` + 主题 | 完全一致 | P0（srgb）/ **P1（主题色链）** |
| 段落对齐 | `a:pPr algn` | 完全一致 | P0 |
| **行距** | `a:lnSpc spcPct / spcPts` | 完全一致（**有系数坑，见 §8**） | P0 |
| 段前后间距 | `spcBef / spcAft`（1/100 pt） | 完全一致 | P0 |
| 缩进 / 悬挂 | `marL / indent` | 完全一致 | P0 |
| 项目符号 / 编号 | `buChar / buAutoNum / buNone` | 基本一致 | P1 |
| 单色填充 | `a:solidFill` | 一致 | P0 |
| **主题色** | `a:schemeClr` + `clrMap` + `clrScheme` + 变换链 | 完全一致 | **P1（头号难点）** |
| 渐变填充 | `a:gradFill`（`gsLst` / `lin ang` / `path`） | 基本一致 | P1 |
| 图片填充 / 图片 | `a:blip r:embed` + `srcRect` 裁剪 | 一致 | P0（图片）/ P1（裁剪） |
| 边框 | `a:ln`（`w` / `prstDash` / 端点箭头） | 基本一致 | P1 |
| 阴影 / 发光 | `a:effectLst` | 近似 | P1 |
| 形状几何 | `a:prstGeom prst` + `avLst` | 常用 5 种 P0 / 40 种 P1 / 全量 P2 | 分阶段 |
| 文本框内边距 / 竖排 | `bodyPr` 的 `inset / vert / anchor` | 一致 | P0/P1 |
| **自动缩放** | `normAutofit fontScale` / `spAutoFit` | 静态缩放 P0 / 迭代收敛 P1 | P0/P1 |
| 表格 | `a:tbl`（合并 / 边框 / 底纹） | 一致 | P0/P1 |
| 图表 | `chart.xml` + 内嵌 xlsx 数据 | **P1 可做到**（复用现有 xlsx 解析） | P1 |
| SmartArt / 公式 / OLE | `diagram` / `oleObj` | 占位框 | P2 |
| 母版背景 / 水印 | slide→layout→master 三级 `bg` | 一致 | P1 |
| 分组 | `grpSp` 子坐标系仿射 | 一致（**易错，见 §8**） | P1 |
| 演讲者备注 | `notesSlideN.xml` | 一致 | P2 |
| 动画 / 切换 | `timing` / `transition` | 不播放，仅标记 | P2 |

**`.ppt`（二进制）的诚实天花板**：见 §4.3，它走的是 MS-PPT + Escher 双层结构，几何与填充的还原成本远高于 pptx。方案给它的目标是 **P0「文本 + 位置 + 基本字号粗斜体 + 图片」**，P1 再补几何与配色。这一点必须在产品上明示，不能承诺与 pptx 同等的还原度。

---

## 3. 数据模型：为什么不复用块树

Word 是**流式**的：block 顺序排列，分页是渲染期推出来的。
PPT 是**绝对定位 + 继承**的：每个形状自带 EMU 坐标，未指定的样式要从 `layout` 的占位符 → `master` 的 `txStyles` 逐级继承。

所以 `DocBlock` 装不下 PPT。但也不是完全另起炉灶——**文本 run 的样式字段口径沿用 `CharStylePatch`**（`utils/parser` 与 `document` 已就这套命名达成一致），将来「把 PPT 文字复制到文档」可以直接对接。

```
DeckSnapshot { v, meta{ title, slideSize{cx,cy}, notesSize, viewState }, slides: Slide[] }

Slide { id, layoutId, masterId, shapes: Shape[], background?, hidden, notes }

Shape {
  id, type: 'sp' | 'pic' | 'cxnSp' | 'graphicFrame' | 'grpSp',
  xfrm { off{x,y}, ext{cx,cy}, rot, flipH, flipV },
  geom { prst, avLst, path? },
  fill: Fill, line: Line, effects?,
  txBody?: { bodyPr{anchor,wrap,inset,vert,autofit}, paras: Para[] },
  table?, chart?, pic{ embed, srcRect }, children?: Shape[]   // grpSp
}

Para { pPr{ algn, lvl, marL, indent, bullet, lnSpc, spcBef, spcAft }, runs: Run[] }

Run  { text, rPr{ sz, b, i, u, strike, cap, baseline, spc, font{latin,ea,cs},
                  fill: ColorRef, highlight?, lang } }

ColorRef = { kind:'srgb',  hex }
         | { kind:'scheme', val:'accent1'|'tx'|'bg1'|..., mods: Transform[] }
         | { kind:'prst', val } | { kind:'sys', val } | { kind:'scrgb'|'hsl' }
Transform = { op:'shade'|'tint'|'lumMod'|'lumOff'|'satMod'|'alpha', val }
```

**两条模型铁律：**

1. **颜色必须存「引用 + 变换链」，不能解析期就拍成 hex。** `a:schemeClr` 要等 `theme` 与 `master.clrMap` 到齐才知道真实 RGB。拍早了就永远对不上母版换色的场景。
2. **样式的解析必须走 `resolveTextStyle(shape, layout, master)` 三级链**，且按 `lvl`（0–8 级大纲）取 `txStyles.titleStyle / bodyStyle / otherStyle` 的 `lvlNpPr`。**这条不做，字号、字体、项目符号会全部出错** —— 这是自研 pptx 解析器的第一大坑。

---

## 4. 解析流水线

### 4.1 pptx（OPC）

```
File → ZipParser 解包（按 part 名索引，media 按需）
  → [Content_Types].xml + _rels/presentation.xml.rels   // 定位 presentation part（不固定叫 ppt/presentation.xml）
  → presentation.xml      sldSz / sldIdLst（顺序 + r:id）/ notesSz
  → 逐页：slideN.xml ──rels──> slideLayoutN.xml ──rels──> slideMasterN.xml ──rels──> themeN.xml
  → resolve：主题色 / 字体 / 文本样式三级继承
  → DeckSnapshot
```

要点：
- **必须读 rels**：`slide1.xml` 的母版可能是 `slideLayout7.xml`，按文件名猜必错。
- `theme/themeN.xml` 取 `clrScheme`（dk1/lt1/dk2/lt2/accent1-6/hlink/folHlink）、`fontScheme`（major/minor × latin/ea/cs）、`fmtScheme`。
- `slideMaster.clrMap` 是「逻辑名 → scheme 名」的映射（`tx→dk1`、`bg1→lt1`、`hlink→hlink`…），**主题色解析必须先过它**。
- media：rels 里的 image part → Blob，落 `deckmedia` 仓，键 = `toArchiveKey(path) + '#' + partName`。

### 4.2 Worker 卸载与回包

大 PPT（几百页 + 内嵌大图）在主线程解 ZIP 会卡死。复用 `parser.worker.ts`：

- 新增 request type `'pptx' | 'ppt'`，走既有 `extractInWorker` 通道。
- **回包列式（SoA）**：`offs: Float64Array`、`exts`、`typeIds: Uint8Array`、`textStart/textLen: Uint32Array`、文本拼成大字符串 + 偏移数组，`transferList` 零拷贝 —— 与既有「Worker 回包列式」铁律一致。
- **图片二进制不进回包**：否则一次几 MB~几十 MB。在 worker 内直接写 IndexedDB，只回传 part 名清单。
- 按 slide 分批，12ms 预算 + `yieldToUI`（复用 `utils/scheduler.ts`）。

### 4.3 ppt（OLE2：MS-PPT + Escher）

```
File → CompoundFile
  → 流 "PowerPoint Document"：记录流（DocumentContainer → SlideListWithText → SlideContainer …）
      文本：TextBytesAtom（8bit）/ TextCharsAtom（16bit），按 shapeId 归位
      样式：TextMasterStyleAtom（母版 9 级）+ StyleTextPropAtom（按 charCount 偏移表的段落/字符异常）
  → 流内嵌 Drawing：Escher（MS-ODRAW）OfficeArtSpContainer
      msofbtSp（spid / shapeType / flags）、msofbtOPT（填充、线条、阴影属性）、
      ClientAnchor / ChildAnchor（位置）、ClientTextbox（关联文本）
  → 流 "Pictures"：图片容器
```

- **坐标单位是 master units（576 /inch）**，不是 EMU：`1 master unit = 1587.5 EMU`。换算错了整页乱飞。
- 记录号与结构**实施时必须按 MS-PPT / MS-ODRAW 规范逐条核对**，不要凭记忆写常量。
- 成本评估：这是整个方案里单位产出最低的一块。建议 P0 只做「文本 + 位置 + 基本字符样式 + 图片」，几何与配色放 P1。

---

## 5. 渲染方案

### 5.1 为什么不用 Canvas

表格用 Canvas 是因为要画百万格子，绘制吞吐是第一矛盾。PPT 一页最多几百个形状，但每个形状要：

- 文本可选中、将来可编辑（Canvas 要自己实现命中测试与输入法）；
- 渐变、阴影、圆角、虚线边框、竖排、自动换行（Canvas 全要手写）；
- 富文本混排与行距度量（Canvas 的 `measureText` 拿不到行高）。

**结论：PPT 走 DOM + 绝对定位**，与 `docView.ts` 同源，样式映射函数抽公共模块复用。

### 5.2 DOM 结构

```
.deck-stage                     缩放原点 top left
 └─ .slide-card                 width/height = sldSz→px，overflow:hidden
     ├─ .slide-bg               三级背景（slide → layout → master）
     └─ .shape                  position:absolute; left/top/width/height = EMU→px
         │                      transform: rotate(θ) scaleX(±1) scaleY(±1)
         ├─ svg.shape-geom      几何：用 SVG path 画（fill/stroke 直接吃 CSS 色）
         └─ .shape-tx           文本层（div，承载富文本）
```

**形状一律用内联 SVG 画**，文本用叠加的 div 承载。这是最贴近 Office 的做法：`prstGeom` 的 160+ 预设本质是 path 定义，用 SVG 表达是天然匹配；而 CSS 的 `border-radius` / `clip-path` 表达不了箭头、星形、标注框。

### 5.3 关键换算（写死在一处，常量化）

| 量 | 单位 | 换算 |
|---|---|---|
| 长度 EMU → px | EMU | `px = emu / 9525`（既有 `EMU_PER_PIXEL`，96dpi） |
| 字号 `sz` | 1/100 pt | `px = sz / 100 * 96 / 72` |
| 行距 `spcPct` | 1/1000 % | `line-height = spcPct / 100000 * 1.2`（**1.2 不能省，见 §8**） |
| 段间距 `spcPts` | 1/100 pt | `px = spcPts / 100 * 96 / 72` |
| 基线 `baseline` | 0.1 % | `vertical-align: (baseline/1000)%` |
| 字距缩放 `fontScale` | 1/1000 % | `font-size × fontScale/100000` |
| 渐变角 `lin ang` | 1/60000 度，顺时针 | CSS `linear-gradient(ang/60000deg, …)` |
| 渐变节点 `gs pos` | 千分比 | `pos/1000 %` |
| 裁剪 `srcRect` | 千分比 | `clip-path: inset(…)` |
| 缩放 | — | `transform: scale(z)` + `transform-origin: top left` + **槽位按「纸×缩放」占位** |

最后一条是文档那边已经踩过的坑（`transform` 不占布局，没有槽位标尺就与纸对不齐），PPT 直接照抄结论。

### 5.4 主题色解析链

```
resolveColor(ref, theme, clrMap):
  scheme 名 → clrMap（tx/bg1/hlink…）→ clrScheme 取基础 RGB
  → 按 XML 子元素顺序依次套 shade / tint / lumMod / lumOff / satMod / alpha
```

**变换顺序必须严格按 XML 子元素顺序链式中缀**，顺序错了颜色就偏 —— Office 是严格顺序执行的，不存在交换律。

---

## 6. 主页面接入点（改哪些文件）

| 接入点 | 文件 | 改动 |
|---|---|---|
| 载体类型 | `contract/editHost.ts` | `WorkspaceKind` 加 `'slide'` |
| 工作区路由 | `workspace/workspaceRouter.ts` | 二态 → 三态，加 `showSlide()` |
| 打开分派 | `workspace/openRouter.ts` | 加 `isSlidePath`（`demo`）；删掉「暂不支持编辑，已按表格打开」降级；加 `openSlideFile` |
| 导入 | `common/fileTree/fileOpener.ts` | `SUPPORTED_EXTENSIONS` 已含 ppt/pptx；`openAsDocument` 旁加 `openAsSlide`，解析结果 → `onDeck(path, snapshot)` |
| 解析 | 新增 `utils/parser/deckParser.ts` | + `parser.worker.ts` 分支 |
| 持久化 | `dataArchitecture/indexDB.ts` | `DB_VERSION 3→4`，新增 `decks`（快照）+ `deckmedia`（图片）仓；只新建仓，旧仓不动，零迁移 |
| 存档键 | `utils/fileName.ts` | **不用改** |
| 工作区 | 新增 `core/deck/` | `DeckStore` / `DeckView` / `DeckWorkspace` / `DeckEditHost` / `SlideRail`（左侧缩略图） |
| 工具栏档案 | `toolbar/profile.ts` | 加 `slide` 档案：开始（剪贴板/字体/段落/绘图）、插入（新建页/图片/表格/形状）、设计（主题/变体/页面设置）、视图（普通/大纲/浏览/备注页）、放映 |
| 快速访问区 | `toolbar/quickAccess.ts` | 打印/打印预览已走注入回调，加 `slide` 分支 |
| 装配 | `core/app.ts` | `new DeckWorkspace(...)`、注入 router、打印回调 |
| 打印预览 | `utils/printPreviewShell.ts` | 三载体共用，PPT 注入「幻灯片样式」+ 讲义/备注布局选项 |
| 样式 | 新增 `core/style/deck.scss` | 复用 `scroller.scss` |

---

## 7. 分阶段路线

### P0 — 能打开、能对位（只读）

1. 新增 `slide` 载体与三态路由，`.demo` 不再降级成表格；
2. `deckParser` 解析 pptx：presentation / slide / layout / master / theme 五级关系，产出 `DeckSnapshot`；
3. `DeckView` DOM 渲染：位置、尺寸、旋转、字号、粗斜体下划线、单色、段落对齐、行距、段间距、图片、表格；
4. 形状几何先覆盖 5 种（`rect` / `roundRect` / `ellipse` / `line` / `triangle`）；
5. 左侧缩略图 rail + 翻页（复用 `DocScrollbar`）+ 缩放条（抽公共 `ZoomBar`）；
6. 打印 / 打印预览（复用 `printPreviewShell`）；
7. `.ppt` 走「文本 + 位置」降级通道。

### P1 — 像 Office

1. **主题色链 + `clrMap` + 变换链**（头号）；
2. **三级文本样式继承 `txStyles` × 9 级 `lvl`**（次号）；
3. 渐变填充、边框、阴影、竖排、文本框内边距、裁剪；
4. 项目符号 / 编号（`buChar` / `buAutoNum` 全类型）；
5. 形状预设扩到 40 种；
6. 分组子坐标系仿射；
7. 母版背景 / 隐藏页；
8. **图表**：`chart.xml` 的缓存数据来自内嵌 `.xlsx`，**可复用现有 `sheetParser`**，用 SVG 画柱/折线/饼；
9. `spAutoFit` 迭代收敛；
10. 字体缺失兜底（见 §8）。

### P2 — 完整与可写

1. 全量 160+ 形状预设；
2. SmartArt / 公式 / OLE 占位与降图；
3. 演讲者备注页、讲义视图；
4. 动画 / 切换：不播放，仅标记 + 属性面板；
5. 编辑能力（形状拖拽、文本编辑、新增页）+ 导出 pptx（复用 `createZipBlob` 思路，与 `docxCodec` 对称）；
6. `.ppt` 二进制几何与 Escher 填充还原。

---

## 8. 已知坑位清单（实施时逐条对照）

| # | 坑 | 后果 | 对策 |
|---|---|---|---|
| 1 | 主题色变换链顺序 | 颜色整体偏色 | 按 XML 子元素顺序链式执行，禁重排 |
| 2 | 跳过 `txStyles` 三级继承 | 字号/字体/项目符号全错 | `resolveTextStyle(shape, layout, master)` 必须实现 |
| 3 | **行距少乘 1.2** | 所有行距偏窄，视觉差异最明显 | `line-height = spcPct/100000 × 1.2`（Office 的「单倍行距」≈ 1.2 × 字号） |
| 4 | 分组 `grpSp` 子坐标系 | 组内形状整体错位 | 按 `chOff/chExt → off/ext` 做仿射，别直接用 `off` |
| 5 | `sz` 当成 px | 字号大一倍 | `sz` 是 1/100 pt |
| 6 | **字体缺失** | 换行位置全变，是「看起来不一致」的最大隐形来源 | 走 `utils/fontsLoader.ts` 的字体注册；缺失时按 metrics 近似并记 warning |
| 7 | 按文件名猜 layout/master | 用错母版 | 一律走 rels 的 r:id |
| 8 | 大图全进内存 | OOM | media 落 `deckmedia` 仓 + ObjectURL，卸载时 revoke |
| 9 | Worker 回包塞图片 | 主线程卡死 | 图片不进回包，worker 内直写 IndexedDB |
| 10 | `Char.toJSON()` 返回字符串 | 样式静默全丢 | 既有铁律：**禁对含 Char 的存档 `JSON.stringify`**；`DeckSnapshot` 保持纯对象 |
| 11 | `.ppt` 坐标单位 | 整页乱飞 | master units（576/inch），不是 EMU |
| 12 | 下划线复杂线型 | 双线/波浪显示为单线 | `u` 属性外还要读 `a:uFill` / `a:uLn` |
| 13 | 容器隐藏时量尺寸 | 幻灯片停在 0 尺寸 | 沿用文档结论：变可见时主动补算一次（`refreshFrame()`） |

---

## 9. 验证方案

1. **测试样本**：`test/` 下补 `.pptx` / `.ppt`（与既有 `.xlsx/.xls/.docx/.doc` 同列），建议三份：纯文本框型、图表 + 表格型、母版换色 + 渐变型。
2. **元素级断言**（项目既有风格：手写 DOM stub + 源码守卫）：解析结果断言具体数值 —— 位置 EMU→px、字号、行距、解析后的主题色 hex、继承后的字体族。
3. **新增守卫文件** `test/test-deck-feature.js`，覆盖：三态路由、`.demo` 不再降级、rels 而非文件名取母版、主题色链、`DB_VERSION` 升到 4 只建仓不迁移、`docParser` 与 `deckParser` 依赖方向。
4. **一致性验收表**：按 §2 那张矩阵逐项目视核对（与 Office 打开同一文件对照打勾）。
5. **回归基线**：`npx tsc --noEmit` + `npm run build` + 全量 `node test/test-*.js`，确认失败文件数不超过既有基线（当前 8 个既有失败文件：filter / render-opt / snapshot-v4 / persist-slice / filetree / xlsx-parse 等，与 PPT 无关）。
6. **反证**：每条守卫注入破坏串确认会红，**源码一律内存备份还原**（禁用 `git checkout`）。

---

## 10. 需要你拍板的两件事

1. **P0 只做只读浏览，编辑放 P2** —— 你的原话是「在主页面展示」，我按只读定位。若需要一上来就能改文字/拖形状，P0 的范围要扩（形状命中测试、输入法、脏标记、撤销栈都要提前）。
2. **`.ppt` 的还原度目标** —— 建议明示为「文本 + 位置 + 基本字符样式」，几何与配色 P1 补。若要求与 pptx 同等，需要单独评估 Escher 的工作量。
