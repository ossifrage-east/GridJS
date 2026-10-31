# Word 文档功能实现方案

> **文档性质：现状核对 + 剩余方案。** 本文不修改任何源码。
>
> **来源**：本文由《解耦优化与Word功能实现方案》拆分而来，并**按 2026-10-30 的实际代码重新核对**：
> 原方案的 P1（文档骨架）与 P2（编辑 + Markdown）已经落地，P3 中的 docx 互操作与打印也已落地；
> 本文因此把"设计意图"与"已实现事实"并列呈现，只对**尚未完成**的能力保留实施计划。
>
> **姊妹文档**：[代码分离解耦优化方案.md](./代码分离解耦优化方案.md)（原合并文档的第一部分）。
>
> **前置**：文档功能建立在解耦后的契约层（`EditHost` / `CommandBus`）与 `docs` 仓之上。

***

## 目录

- [1. 现状一览](#1-现状一览)
- [2. 六项设计决策与实际落地](#2-六项设计决策与实际落地)
- [3. 数据模型](#3-数据模型)
- [4. chars 复用](#4-chars-复用)
- [5. 三条转换链路](#5-三条转换链路)
- [6. 工具栏复用与 DocEditHost](#6-工具栏复用与-doceditHost)
- [7. 持久层：docs 仓](#7-持久层docs-仓)
- [8. 渲染与编辑](#8-渲染与编辑)
- [9. 工作区行为](#9-工作区行为)
- [10. 文件打开/导出链路](#10-文件打开导出链路)
- [11. 待办与差距](#11-待办与差距)
- [12. 风险与验收](#12-风险与验收)
- [13. 一页纸总结](#13-一页纸总结)

***

## 1. 现状一览

### 1.1 文档实现层文件

| 行数 | 文件 | 职责 |
| ---: | --- | --- |
| 680 | `core/document/blockTree.ts` | 块树模型 + 纯函数（Ch​ar 复用、样式、spans、克隆/规范化/恢复） |
| 497 | `core/document/docStore.ts` | 装载 / 脏标记 / 防抖落库 / 切换 / 独立撤销栈 |
| 1939 | `core/document/docView.ts` | DOM 渲染 + contenteditable 编辑 + 选区/光标 + 打印预览渲染 |
| 768 | `core/document/docWorkspace.ts` | 工作区装配：两条工具线、滚动条、缩放、字符计数、导出、页面设置 |
| 564 | `core/document/docEditHost.ts` | 实现 `EditHost`：命令 → 块树操作 |
| 477 | `core/document/docPrinter.ts` | 打印 / 打印预览（独立于表格 `Printer`） |
| 321 | `core/document/docScroller.ts` | 文档专用滚动条（读真实 `scrollLeft/Top`，非虚拟滚动） |
| 265 | `core/document/markdownView.ts` | Markdown 源码视图（textarea + 高亮层 + 行号） |
| 93 | `core/document/docRuler.ts` | 横向 / 纵向标尺 |
| 699 | `core/document/converter/docxCodec.ts` | docx ⇄ 块树（`xmlScanner` + `zip` + `createZipBlob`） |
| 660 | `core/document/converter/markdownCodec.ts` | 块树 ⇄ Markdown（含扩展语法与 `standardOnly`） |
| 84 | `core/document/converter/textCodec.ts` | 块树 ⇄ 纯文本 |

### 1.2 相关基础设施

| 文件 | 说明 |
| --- | --- |
| `core/contract/editHost.ts`、`commandBus.ts` | 命令契约与总线（见解耦文档） |
| `core/workspace/workspaceRouter.ts` | 表格 ⇄ 文档切换、宿主重绑、导航面板搬动 |
| `core/sheets/sheetEditHost.ts` | 表格侧宿主（同契约的另一实现） |
| `core/dataArchitecture/indexDB.ts` | `docs` 仓、`DB_VERSION = 3` |
| `core/common/fileOpener.ts` | `.docx` / `.md` / `.txt` / `.doc` 的解析分派 |
| `core/app.ts` | 打开分派、内容桥、导出、活动文件记忆 |
| `test/test-doc-feature.js` | 块树 / 转换器 / 存储 / 路由 / 工具栏并档的守卫 |

### 1.3 已具备的能力（自检清单）

- [x] 打开 `.text` / `.md` / `.docx` 进入文档工作区
- [x] 富文本编辑（字符样式、段落样式、列表、缩进、对齐）
- [x] 撤销 / 重做（独立 `UndoManager` 实例，连续输入合并）
- [x] Markdown 源码视图，双向同步
- [x] 文档即存即取：`docs` 仓 + 防抖落库 + 切换前 flush
- [x] 缩放（50%–400%）、Ctrl+滚轮、右侧垂直滑块
- [x] 视图状态持久化：缩放 / 滚动位置随文档存档恢复
- [x] 底部即时计算栏：默认全文，选区时显示选区字符数
- [x] 底部 / 右侧工具线（与表格同构的 `Split` + 滚动条 + 导航面板）
- [x] 文档打印 / 打印预览 / 页面设置（纸张 / 方向 / 页边距 / 缩放）
- [x] docx 导入（`.docx` → 块树）与导出（块树 → `.docx`）
- [x] 图片块与超链接区间的模型、渲染、docx 往返
- [ ] Python `doc` 命名空间（见第 11 节）
- [ ] 文档查找 / 替换
- [ ] Markdown 文件导出入口
- [ ] `.doc` / `.pptx` / `.ppt` 的现代解析链路（现为纯文本尽力提取）

***

## 2. 六项设计决策与实际落地

| # | 决策 | 结论 | 落地情况 |
| --- | --- | --- | --- |
| **W1** | toolbar UI 复用方式 | UI 外壳共享 + `CommandBus` 派发 | ✅ `toolbar/profile.ts` 声明两套档案；表格分区在文档里复用并按需隐藏按钮 |
| **W2** | Markdown 定位 | 交换格式 + 可编辑源码视图，非运行时真源 | ✅ `markdownView.ts` + `markdownCodec.ts`，块树仍是唯一真源 |
| **W3** | 持久层 | 新增 `docs` 仓，`Char` 直接复用为 run 样式载体 | ✅ `indexDB` `DOCS` 仓；块树行内内容就是 `Char[]` |
| **W4** | 转换链 | 以块树为中枢 | ✅ 6 条边全部实现（docx / text / Markdown） |
| **W5** | 渲染 | DOM，不走 Canvas | ✅ `docView.ts` 用 DOM；滚动条按原生 `scrollLeft/Top` |
| **W6** | docx 导出 | 复用 `xlsxWriter` 的 `createZipBlob` | ✅ `docxCodec.snapshotToDocx` 直接调 `createZipBlob` |

**与原方案的唯一实质偏差**：原方案写的是"严格受控渲染（DOM 只读、不反写模型）"，
实际实现选择**contenteditable 作为输入层 + 输入/粘贴/选区事件把 DOM 同步回块树**。
理由与代价见第 8 节——这是本功能最需要留意的技术点。

***

## 3. 数据模型

```ts
// core/document/blockTree.ts
export type BlockType =
    | 'paragraph' | 'heading' | 'listItem' | 'quote'
    | 'code' | 'table' | 'image' | 'hr';

export type DocAlign = 'left' | 'center' | 'right' | 'justify' | 'distribute';
export type DocListStyle = 'bullet' | 'ordered' | 'none';

export interface DocSpan {
    start: number;   // 码点下标（含）
    end: number;     // 码点下标（不含）
    link?: string;
    highlight?: string;
    vertAlign?: 'superscript' | 'subscript';
}

export interface DocBlock {
    id: string;
    type: BlockType;
    /** 行内内容：直接复用表格的 Char */
    chars: Char[];
    /** 扩展属性区间（链接 / 高亮 / 上下标） */
    spans?: DocSpan[];
    align?: DocAlign;
    level?: number;
    listStyle?: DocListStyle;
    indent?: number;
    /** 表格：行 × 单元格 */
    rows?: DocBlock[][];
    colSpan?: number;
    rowSpan?: number;
    /** 图片：二进制另存 images 仓，这里只留指针与尺寸 */
    imageRef?: { assetId: string; width: number; height: number };
}

export interface DocPageSetup {
    orientation?: 'portrait' | 'landscape';
    paperSize?: 'A4' | 'A5' | 'A3' | 'Letter' | 'Legal';
    marginPreset?: 'normal' | 'wide' | 'narrow';
    scaleMode?: 'none' | 'fit-page' | 'custom';
    scaleRatio?: number;
}

/** 每份文档各自记住的「看的姿势」 */
export interface DocViewState {
    zoom?: number;       // 0.5–4
    scrollLeft?: number;
    scrollTop?: number;
}

export interface DocMeta {
    title: string;
    view: 'rich' | 'markdown';
    pageSetup: DocPageSetup;
    viewState?: DocViewState;
}

export interface DocSnapshot {
    v: 1;
    meta: DocMeta;
    blocks: DocBlock[];
}
```

### 3.1 纯函数工具（已实现）

| 函数 | 用途 |
| --- | --- |
| `charsFromText(text, style?)` | 文本 → `Char[]` |
| `blockText(block)` / `blockLength(block)` / `setBlockText` | 取文本 / 长度 / 整体替换 |
| `applyCharStyle(chars, start, end, patch)` | 区间套样式（稀疏赋值） |
| `queryCharStyle(chars, start, end)` | 区间样式代表值（按钮高亮） |
| `spanAt` / `addSpan` / `clearSpans` | 扩展区间读写 |
| `cloneBlocks` / `restoreBlocks` / `normalizeBlocks` | 撤销快照 / 原型恢复 / id 去重 |
| `createBlankSnapshot` / `isEmptySnapshot` / `isCharTreeIntact` | 空档 / 对账 / 完整性守卫 |

***

## 4. chars 复用

### 4.1 为什么能直接复用

| `Char` 字段 | `FontStyle` 字段 | 语义 |
| --- | --- | --- |
| `fontFamily` | `fontFamily` | 字体族 |
| `fontSize` | `fontSize` | 字号 |
| `fontWeight` | `fontWeight` | 加粗 |
| `fontStyle` | `fontStyle` | 斜体 |
| `fontColor` | `fontColor` | 字体颜色 |
| `underline` | `underline` | 下划线 |
| `strikethrough` | `strikethrough` | 删除线 |

七个字段完全一致，且 `Char`：

1. 是**纯数据类**（非 `EventEmitter`）——大量块 / run 时无监听表开销；
2. `fromRaw` 是**稀疏复制**——未赋值字段不成为自有属性，稀疏性天然成立；
3. 已提供 `toJSON` / `fromJSON`——存档序列化直接复用。

### 4.2 硬约束：`Char` 零改动

文档需要而 `Char` 没有的属性（`highlight` / `vertAlign` / `link`）**一律落在 `DocBlock.spans`**，不扩 `Char`。
原因：`Char.toJSON` 的输出形状参与表格存档 v4 的字节比对与快照相等判断
（`_sameCellSnapshot` / `_valuesSnapshotsEqual`），改字段等于动存量数据。

### 4.3 三个复用接口（已落地）

- `applyCharStyle(chars, start, end, patch)` —— 与 `Cell.setSelectedCharsFont` 同构；
- `queryCharStyle(chars, start, end)` —— 供 `DocEditHost.snapshot()` 回填按钮高亮；
- `blockText(block)` —— 与 `Cell.getText` 同构，供 Markdown / docx / 字符计数共用。

***

## 5. 三条转换链路

以**块树为中枢**，一共 6 条边：

```
                    ┌──────────┐
                    │  块树     │  DocBlock[]（chars: Char[]）
                    └────┬─────┘
             ┌───────────┼───────────┐
             ▼           ▼           ▼
        ┌─────────┐ ┌─────────┐ ┌─────────┐
        │  docx   │ │  text   │ │Markdown │
        └─────────┘ └─────────┘ └─────────┘
```

| 边 | 实现 | 说明 |
| --- | --- | --- |
| docx → 块树 | `converter/docxCodec.ts` `docxToSnapshot()` | `ZipParser` 解压 → `scanXml` 扫描 `word/document.xml` → 块树（**禁 DOMParser**） |
| 块树 → docx | `converter/docxCodec.ts` `snapshotToDocx()` | 生成 OOXML 部件 → `createZipBlob()` 打包 |
| 块树 → text | `converter/textCodec.ts` `blocksToText()` | 逐块 `blockText`，块间 `\n` |
| text → 块树 | `converter/textCodec.ts` `textToBlocks()` | 按 `\n` 切段，每行一个 `paragraph` |
| 块树 → Markdown | `converter/markdownCodec.ts` `blocksToMarkdown(blocks, standardOnly?)` | 行内样式合并成 `**` / `*` / `~~` / `<u>`，扩展语法见下 |
| Markdown → 块树 | `converter/markdownCodec.ts` `markdownToBlocks()` | 自研解析器（不引第三方库） |

**Markdown 扩展语法**（往返无损，但对标准阅读器只是原样文本）：

| 样式 | 输出 |
| --- | --- |
| 加粗 / 斜体 / 删除线 / 下划线 | `**x**` / `*x*` / `~~x~~` / `<u>x</u>` |
| 字体颜色 | `x{#FF0000}` |
| 高亮 | `==x==` |
| 链接 | `[x](url)` |

`blocksToMarkdown(blocks, true)` 会**只输出标准语法**（丢掉颜色/高亮扩展），用于"导出给外部工具"。

**docx 部件**（`snapshotToDocx` 产出）：

```
[Content_Types].xml
_rels/.rels
word/document.xml          ← 由块树生成
word/styles.xml            ← 由用到的样式生成
word/_rels/document.xml.rels   ← 有图片 / 超链接时
word/media/imageN.png          ← 有图片时
```

**测试覆盖**（`test/test-doc-feature.js`）：
纯函数稀疏性 → 拷贝/原型/规范化 → `textCodec` → Markdown 结构往返 → 行内样式与扩展语法 →
**docx 往返（含高亮/颜色/字号、表格）** → 依赖方向守卫 → 存储/路由守卫 → 工具条并档/导出。

***

## 6. 工具栏复用与 DocEditHost

### 6.1 两套档案（`toolbar/profile.ts`）

| 工作区 | 页签与分区 |
| --- | --- |
| 表格 | 开始：格式刷 / 字体样式 / 对齐方式 / 数据处理；页面：打印设置 |
| 文档 | 开始：剪贴板（格式刷分区改名）/ 字体 / 段落对齐 / 段落样式；视图：文档视图（Markdown 开关）；页面：页面设置 |

**同一个分区实例**在两种工作区里使用，差异只有：

1. **档案里的名字**（如"格式刷"→"剪贴板"）；
2. **`canExecute` 置灰**：表格专属命令在文档侧返回 false。

### 6.2 `DocEditHost` 支持的命令（已实现）

| 类别 | 命令 |
| --- | --- |
| 字符 | `setFontFamily` / `setFontSize` / `toggleBold` / `toggleItalic` / `toggleUnderline` / `toggleStrike` / `setTextColor` / `setBackground` |
| 段落 | `setAlign` / `indentStep` / `setParaStyle` / `toggleList` |
| 历史 | `undo` / `redo` |
| 剪贴板 | `copy` / `cut` / `paste` / `pasteValue` / `pasteStyle` |
| 格式刷 | `brushClick` / `brushDoubleClick` |
| 表格专属 | `setAlign`（水平对齐可用，垂直对齐不可用）/ `toggleMergeCells` / `toggleFilter` / `sortUp` / `sortDown` / `setBorder` / `toggleWrap` → `canExecute` false |

### 6.3 状态同步

- 分区不再订阅 `DataCollection` 的内部事件，统一由 `CommandBus` 广播 `EDIT_STATE_CHANGED`；
- `DocEditHost.snapshot()` 从当前选区 / 折叠光标处的 `queryCharStyle` / 段落属性回填快照；
- 按钮的 `isDisabledOnEditting` 依赖总线**原样转发** `IS_EDITTING_CHANGED`（`test-edit-decouple.js` 有断言）。

***

## 7. 持久层：docs 仓

### 7.1 改动（已完成）

| 位置 | 内容 |
| --- | --- |
| `indexDB.ts` `STORE_NAMES` | `DOCS: 'docs'` |
| `indexDB.ts` `DB_VERSION` | 2 → **3** |
| `indexDB.ts` `OBJECT_STORE_CONFIG` | `{ keyPath: 'docKey', autoIncrement: false }` |
| `indexDB.ts` | `setDoc / getDoc / deleteDoc / getAllDocs` |
| `createObjectStores` | 已是幂等（`if (!contains)`）→ **存量零迁移** |

### 7.2 存档结构

```ts
export interface DocSnapshot {
    v: 1;
    meta: DocMeta;
    blocks: DocBlock[];
}
```

**为什么结构化克隆直存、不分片**：

| 维度 | 表格 `values` | 文档 `docs` |
| --- | --- | --- |
| 体量 | 10⁵～10⁶ 单元格 | 10²～10⁴ 块 |
| 序列化 | 分片 JSON（4000 格/片）+ `SparseAxis` | 结构化克隆直存 |
| 理由 | 分片为避免百万级字符串拼接造成主线程长任务 | 文档规模下引入分片只增复杂度 |

> ⚠️ **结构化克隆会丢原型**：`Char` 实例读回后是普通对象。
> `DocStore` 装载时必须经 `restoreBlocks()`（内部 `Char.fromRaw`）重建原型，
> 否则 `chars` 上的样式方法不可用。`test-doc-feature.js` 的「模拟结构化克隆」段与
> 「`Char.toJSON()` 返回字符串会导致样式静默全丢」段正是在防这件事。

### 7.3 存档键规则

沿用 `toArchiveKey(path)` 纯函数（`utils/fileName.ts`）：

| 路径 | 存档键 | 仓 |
| --- | --- | --- |
| `报表.grid` | `报表` | `values` |
| `报表.text` | `报表.text` | `docs` |
| `报表.md` | `报表.md` | `docs` |
| `归档/销售.grid` | `归档/销售` | `values` |
| `归档/说明.text` | `归档/说明.text` | `docs` |

**同名不同后缀 = 两个独立文件，各自独立存档** —— 项目既有铁律，`docs` 仓天然继承。

### 7.4 落库时机

| 时机 | 行为 |
| --- | --- |
| 每次 `update()` / `updateMeta()` | `markDirty()`，500ms 防抖落库 |
| 连续输入 | 700ms 静默窗口合并成一条撤销记录 |
| 切换文档 / 关闭 / 页面隐藏 / 导出前 | `flush()` 立即落库 |
| 打开新文档 | 先 `flush()` 旧文档，再装载新文档 |

`DocStore` 对外提供：`load` / `switchTo` / `close` / `update` / `updateMeta` / `replaceBlocks` / `flush` / `readSnapshot` / `writeSnapshot` / `deleteSnapshot`。
`readSnapshot/writeSnapshot/deleteSnapshot` 供内容桥（复制/移动/删除）与导出使用。

### 7.5 视图状态持久化（本次新增）

`DocMeta.viewState` 记录 `zoom / scrollLeft / scrollTop`：

- 滚动与缩放触发 **300ms 防抖**写入 meta（`updateMeta` 会顺带触发落库）；
- 切换文档、隐藏、`flush()` 前**立即补存**；
- 打开时先恢复缩放（无存档则退回「适应页宽」），下一帧写回滚动位置并刷新滚动条；
- 老存档没有 `viewState` 时按默认（适应页宽 + 页首）打开。

***

## 8. 渲染与编辑

### 8.1 为什么用 DOM 而不是 Canvas

| 维度 | 表格（Canvas 合理） | 文档（DOM 更合适） |
| --- | --- | --- |
| 节点规模 | 10⁵～10⁶ 单元格 | 10²～10⁴ 块 / run |
| 排版 | 每格单行或整体换行 | 行内混排、中西文混排、行内样式、基线对齐 |
| 选区 | 矩形选区 | 任意字符区间、跨块选区 |
| 输入法 | 单元格单行编辑 | 长文本 IME 组字、候选窗定位 |
| 命中测试 | 坐标算行列 | `caretRangeFromPoint` 免费提供 |

### 8.2 结构

```
.doc-workspace
├── .doc-body
│   ├── 标尺（docRuler：上 / 左）
│   ├── .doc-scroller（真实滚动容器）
│   │   └── 页面容器 .doc-page（纸张 + 页边距 + 缩放 transform）
│   │       └── 块容器（contenteditable 宿主）
│   │           └── 每块 <div data-block-id="...">
│   │               └── 行内内容按 chars + spans 分段渲染成 <span>
│   ├── Markdown 源码视图（textarea + 高亮层 + 行号，默认隐藏）
│   └── 空文档占位
├── .bottomTools（左：字符计数；右：横向滚动条）
└── .rightTools（上：导航面板；下：缩放滑块；最下：纵向滚动条）
```

### 8.3 编辑循环（实际实现）

| 环节 | 实现 |
| --- | --- |
| 输入层 | **一个** `contenteditable` 承载全部正文（避免多容器跨块光标不稳） |
| 模型同步 | `input` 事件 → 从 DOM 读 `data-block-id` + 样式 → 重建块树 → `store.update()` |
| 选区 | 浏览器原生 `Range`；折叠光标上用 `pendingStyle` 记录"待生效样式" |
| 粘贴 | 一律按纯文本插入（拒绝 HTML，避免模型外结构被静默丢弃） |
| 单击定位 | 按下/抬起位移 ≤3px 视为单击 → `caretRangeFromPoint` 落光标 |
| 输入法 | `beforeinput` 中检查 `isComposing`；组字期间不触发"待生效样式"的自行插入 |
| 撤销 | 复用 `UndoManager` 类、**独立实例**；块树快照（`cloneBlocks` / `restoreBlocks`） |
| 样式回填 | `applyRunStyle`：字体/字号/粗斜/下划线/删除线/颜色；`spans` 负责高亮/上下标/链接 |

**取舍说明**：contenteditable 换来成熟的 IME 与选区体验，代价是浏览器会对 DOM 做归一化
（合并文本节点、插入 `<b>`、丢空格）。因此：

- 每次同步都以**块树的字符序列**为准重新对齐，多出来的结构直接丢弃；
- 链接只上色不加下划线（下划线会被同步读成"字符带下划线"，往返一次就多一层 `<u>`）；
- 上下标只改 `verticalAlign`、不改字号（改字号会让"导出→再导入"逐次缩小）；
- `paste` 强制 `text/plain`。

`test-doc-feature.js` 对上述往返行为有专门断言，改动同步逻辑前先跑它。

### 8.4 缩放与滚动

| 项 | 实现 |
| --- | --- |
| 缩放 | `.doc-page` 上 CSS `transform: scale()`；范围 **0.5–4**（滑块 50–400，同源常量） |
| 适应页宽 | `setZoom(可用宽度 / 纸张宽度)`，钳制到 0.5–4 |
| Ctrl+滚轮 | 每格 ±10%，`preventDefault` 防止浏览器整页缩放 |
| 滚动条 | `DocScrollbar` 读真实 `scrollLeft/Top`，不伪造虚拟滚动 |
| 标尺 | 挂在 `.doc-body`（不是滚动区），监听 scroll 反向平移刻度 |

### 8.5 打印

- `DocPrinter` 独立于表格 `Printer`（后者绑在 `DataCollection` 上，文档里点不动）。
- 预览走 iframe 副本，副本内 `contenteditable = false`（避免预览里出现光标与拼写提示）。
- 页面设置写 `DocMeta.pageSetup`，与 `DOC_PAPER_SIZES` / `DOC_MARGIN_PRESETS` 一一对应。

***

## 9. 工作区行为

### 9.1 打开 / 切换 / 关闭

```
打开 .text/.md/.docx
  └─ App.openDocFile()
       ├─ 若从表格切来：datas 落库（switchToSheet 顺序语义）
       ├─ WorkspaceRouter.showDoc()（显隐容器 + bus.bind(DocEditHost) + 导航面板搬入）
       └─ DocWorkspace.open(key, title)
            ├─ flush 上一份的 viewState
            ├─ DocStore.load()（无存档则建空白存档并落库）
            ├─ 恢复缩放（无存档 → 适应页宽）
            ├─ 下一帧写回 scrollLeft/Top + 刷新滚动条
            └─ 富文本模式时聚焦
```

切换回表格时：`DocWorkspace.hide()`（先 flush viewState）、`Sheets.refreshLayout()`、导航面板搬回表格右侧工具线。

### 9.2 底部 / 右侧工具线（与表格同构）

| 位置 | 内容 | 持久化 |
| --- | --- | --- |
| 底部左半 | 即时计算栏：`字符 N` | — |
| 底部右半 | 横向滚动条 | 滚动位置写入 `viewState` |
| 右侧上半 | 导航面板（与表格**共用同一个** `Navigator` 实例，切换时搬动 DOM） | 导航宽度 / 选中项由 `Navigator` 自身持久化 |
| 右侧下半 | 垂直细线缩放滑块（50–400） | 缩放写入 `viewState` |
| 右侧最下 | 纵向滚动条 | 滚动位置写入 `viewState` |

两条工具线都是 `Split`，位置按 `storageKey`（`doc-bottomTools` / `doc-rightTools`）持久化；
`Split.refreshLayout()` 在容器重新可见时把隐藏期间读回的位置落地（切工作区不会丢分割位置）。

### 9.3 即时计算栏

- 默认统计**全文**字符数（去空白）；
- 当文档滚动区内有选区时，只统计**选区**字符数；
- `selectionchange` + `DOC_CHANGED` / `DOC_RESET`，80ms 防抖。

***

## 10. 文件打开/导出链路

### 10.1 打开（`fileOpener.ts`）

| 扩展名 | 走向 | 说明 |
| --- | --- | --- |
| `.docx` | `docxToSnapshot()` → 块树 | 流式 XML 扫描，**不用 DOMParser**，为将来 Worker 化留路 |
| `.md` | 读文本 → `textToBlocks()`，`meta.view = 'markdown'` | 直接落在源码视图 |
| `.txt` | 读文本 → `textToBlocks()` | 纯文本还原段落 |
| `.doc` | `DocumentParser` 提取纯文本 → `textToBlocks()` | 二进制格式，尽力而为（见第 11 节） |
| `.grid` | 表格链路 | 不经 `DataCollection` 以外的路径 |
| `.pptx` / `.ppt` | 表格链路（`DocumentParser` 提取文本铺到 A 列） | 目前没有演示文稿工作区 |

解析结果由 `App.openParsedDocument()` **先写入 `docs` 仓再打开**，
使"刚打开的文件"与"刷新后恢复的文件"走同一条路。

### 10.2 导出

| 入口 | 实现 |
| --- | --- |
| 文件树右键「另存为docx」 | 与「另存为xlsx」同一套导出清单对话框：列出选中的 `.text` 文件，由用户逐行 / 批量「开始转换」后才生成 |
| docx 转换作业 | `FileTree.convertDocx` 回调 → `App.convertTextToDocx()`：先 `flush()` 当前文档 → 读快照 → `snapshotToDocx()`（无存档时导出空白 Word）→ 作业产物由面板「下载 / 打包下载」写盘 |
| Markdown | `DocWorkspace.toMarkdown(standardOnly)` 已提供字符串，**尚无下载入口** |

***

## 11. 待办与差距

| # | 差距 | 现状 | 建议 |
| --- | --- | --- | --- |
| 1 | **Python `doc` 命名空间** | `utils/python/` 只有表格的 `silkspaces` 命名空间（`sheetNameSpace.ts`）；没有 `open_doc()` / 写回 | 新增 `docNameSpace.ts`，复用同一 Worker 通道：`open_doc()` 读 `docs` 仓快照 → Python 对象；写回时打包成 `DocSnapshot` 并进撤销栈。载荷只传"当前文档"，避免把全部文档预加载进 Worker |
| 2 | **`.doc` 解析仍走 `DOMParser`** | `documentParser.ts:207`（仅 `.doc` 二进制文本提取路径用到） | 两条路：① 改用 `xmlScanner`/`cfb` 的纯字节扫描（与"XML 一律 xmlScanner"铁律完全一致）；② 明确标注"尽力提取、结构丢失"并在 UI 提示。**不要**让这条路径进入 docx 富文本链路 |
| 3 | **docx 解析未 Worker 化** | 目前主线程 `docxToSnapshot()`，注释里写明"为将来搬进 Worker 留路" | 并入 `parser.worker`：`decode/extract/generate` 三步法与 `sheetParser` 同构；验收目标：5 MB docx 打开 < 1.5 s、主线程单任务 < 50 ms（**当前未验证，是目标不是结论**） |
| 4 | **文档查找 / 替换** | `CommandId` 无 `findAll` / `replaceAll`；`searchDialog` 仍直连表格 `DataCollection` | 与解耦文档 K1 一起做：先补命令契约，再让对话框走总线；文档侧实现"块内范围查找 + 选区替换" |
| 5 | **Markdown 导出入口** | `toMarkdown()` 已有，`standardOnly` 也支持 | 文件树"另存为 Markdown"接 `exportMd` 回调；默认提供"标准 Markdown / 带扩展语法"两个选项 |
| 6 | **图片块端到端** | 模型 / docx 往返 / 渲染（占位）已通；`images` 仓已有 | 渲染从占位文本换成真实图片：`assetId` → `images` 仓 → `Blob URL`；插入图片的 UI 入口（「插入」页签目前未装配） |
| 7 | **文档内表格编辑** | 渲染与 docx 往返已通；`docView` 有表格同步逻辑 | 补单元格内换行 / 增删行列的操作入口；若暂不做，至少保证"外部 docx 导入的表格可读、可导出、不丢结构" |
| 8 | **「插入」页签** | 档案里没有 | 表格 / 图片 / 链接 / 代码块 / 水平线五个按钮，全部走 `CommandId`（新增 `insertTable` / `insertImage` / `insertLink` / `insertHr`） |

> 优先级建议：**查找替换（与 K1 合并）→ Python 命名空间 → 图片端到端 → docx Worker 化 → Markdown 导出**。
> 前两项有明确的复用路径与守卫基础；docx Worker 化收益大但风险也大，放在有稳定的回归基线之后。

***

## 12. 风险与验收

### 12.1 风险

| # | 风险 | 影响 | 缓解 |
| --- | --- | --- | --- |
| 1 | **contenteditable 的归一化 / IME** | 光标跳动、丢字、样式漂移 | 已按第 8.3 的取舍处理；任何同步逻辑改动前先跑 `test-doc-feature.js`；新增用例覆盖输入法组字 |
| 2 | **`Char` 被误改** | 表格存档 v4 字节变化 → 快照比对失效 | `Char` **零改动**是硬约束；扩展属性一律走 `DocBlock.spans` |
| 3 | **结构化克隆丢原型** | 读回后 `chars` 样式方法不可用 | `DocStore` 装载走 `restoreBlocks()`；测试已模拟克隆并断言原型 |
| 4 | **`viewState` 噪声写库** | 每次滚动停下都触发一次文档落库 | 300ms 防抖 + 只在停止后写；若文档体量大，可后续把 `viewState` 拆到独立 settings 键（暂不必要） |
| 5 | **`reconcileValueStore` 漏扫 `docs`** | 删文件后残留孤儿存档 | 复核并扩展为同时扫描 `docs`；**只删空存档，非空要迁移** |
| 6 | **`.doc` 路径的 DOMParser** | 与项目铁律相邻冲突、进不了 Worker | 见第 11 节第 2 条，二选一明确化 |
| 7 | **窗口缩放/容器隐藏期间的尺寸** | 纸张、标尺、工具线量到 0 | 已在 `show()` 中补 `refreshFrame()` / `refreshLayout()`；新增布局改动须复用这套"重新可见时补算"范式 |

### 12.2 验收口径

| 场景 | 验收 |
| --- | --- |
| 单元 / 结构 | `node test/test-doc-feature.js` 全绿（块树、三个 codec、docx 往返、依赖方向、存储/路由、工具栏并档） |
| 解耦联动 | `node test/test-edit-decouple.js` 全绿（契约零实现依赖、工具栏无直连、装配可重入） |
| 打开 | `.text` / `.md` / `.docx` 进入文档工作区；`.grid` 进入表格；互相不串味 |
| 编辑 | 中文输入法组字不丢字、不跳字；加粗 / 颜色 / 列表 / 缩进立即生效；撤销重做连选区一起回退 |
| 持久化 | 关闭标签页再打开：正文、缩放、滚动位置、工具线分割位置全部恢复 |
| 转换 | Markdown 往返结构无损；docx 导出 Word / WPS 打开无警告；样式（含高亮/颜色/字号/表格）保留 |
| 性能 | 常规文档（10²–10⁴ 块）输入无可感延迟；docx 导入暂以主线程实测为基线 |

***

## 13. 一页纸总结

| 议题 | 结论 |
| --- | --- |
| **总体状态** | P1 骨架 + P2 编辑/Markdown + P3 的 docx/打印**已落地**；剩余是 Python、查找替换、图片端到端、Worker 化、Markdown 导出 |
| **架构位置** | 与表格并列的工作区；通过 `EditHost` + `CommandBus` 复用同一套工具栏；通过 `WorkspaceRouter` 切换 |
| **真源** | 块树 `DocBlock[]`（行内 `Char[]` + 扩展 `DocSpan[]`）；Markdown 只是源码视图，不是运行时真源 |
| **复用 chars** | 七个样式字段与 `FontStyle` 完全一致；`Char` **零改动**，高亮/上下标/链接走 `spans` |
| **转换链** | 块树中枢，6 条边；docx 双向、text 双向、Markdown 双向（扩展语法可降级） |
| **docx 导出** | 复用 `xlsxWriter` 的 `createZipBlob`；无需自写 ZIP |
| **持久层** | `docs` 仓（`DB_VERSION = 3`，幂等升级，存量零迁移）；结构化克隆直存，读回重建 `Char` 原型；`viewState` 随档恢复 |
| **渲染** | DOM + 单个 contenteditable 输入层；渲染模型→DOM，输入 DOM→模型同步；拒绝 HTML 粘贴 |
| **工具栏** | 两套档案共用分区；表格专属命令在文档侧 `canExecute` false → 自动置灰 |
| **最大风险** | ① contenteditable 的 IME/归一化 ② `Char` 被误改 ③ 原型丢失 ④ `docs` 仓对账遗漏 |
| **下一步** | 查找替换（与解耦 K1 合并）→ Python `doc` 命名空间 → 图片端到端 → docx Worker 化 → Markdown 导出 |
