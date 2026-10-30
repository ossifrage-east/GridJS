# Word 文档功能实现方案

> **文档性质：纯方案建议。** 本文不修改任何源码，仅基于对当前实现的阅读给出设计与落地建议。
>
> 适用对象：`src/core/`（工作区装配层）、`src/core/toolbar/`（工具栏层）、`src/utils/parser/`（解析层）、`src/core/dataArchitecture/indexDB.ts`（持久化层）。
>
> 姊妹文档：`Excel多线程解析方案.md`（大表格解析链路）。本文沿用其分层、优先级与验收口径。

---

## 目录

1. [背景与目标](#1-背景与目标)
2. [现状盘点：文档在项目里现在是什么](#2-现状盘点文档在项目里现在是什么)
3. [五个关键设计决策](#3-五个关键设计决策)
4. [目标架构](#4-目标架构)
5. [文档数据模型](#5-文档数据模型)
6. [文档工作区实现](#6-文档工作区实现)
   - [6.1 渲染层：为什么文档不用 Canvas](#61-渲染层为什么文档不用-canvas)
   - [6.2 编辑层：选区、输入法、撤销](#62-编辑层选区输入法撤销)
   - [6.3 工具栏：命令总线与分区装配](#63-工具栏命令总线与分区装配)
   - [6.4 存储层：新增 docs 对象仓](#64-存储层新增-docs-对象仓)
   - [6.5 导入导出：docx / Markdown](#65-导入导出docx--markdown)
   - [6.6 Python 编程接口](#66-python-编程接口)
   - [6.7 打印](#67-打印)
7. [实施路线图](#7-实施路线图)
8. [风险与已知坑](#8-风险与已知坑)
9. [附录：接口草案](#9-附录接口草案)

---

## 1. 背景与目标

**诉求**：在 silkspaces 里实现 Word 文档功能 —— 有工具栏、有 Markdown、有富文本。

**为什么现在做是合适的**：项目已经具备三块基础设施，且都已按「表格」这一种载体调优完毕：

| 基础设施 | 现状 | 对文档功能的价值 |
| --- | --- | --- |
| 文件树 + 三种原生后缀 | `grid` / `text` / `demo`（`utils/fileName.ts`），「新建文档」菜单项已存在 | 文档的**文件身份**已经打通，缺的是内容载体 |
| 工具栏三级折叠 | section → tabs → root-compact 已跑通，含紧凑模式与量测 | 文档工具栏**外壳**直接复用 |
| 解析三步架构 | decode → extract → generate + Worker + 列式打包 | 文档解析**沿用同一范式** |
| 持久化 | IndexedDB 三仓 + 存档键纯函数 `toArchiveKey(path)` | 文档存档**天然分键**，不污染表格存档 |
| Python 编程 | Pyodide Worker + `sheetNameSpace.ts` 命名空间 | 文档可获得同等的宏编程能力 |

**目标（可度量）**：

| 指标 | 目标值 |
| --- | --- |
| 新建 `.text` 文件 → 出现可编辑文档视图 | < 100 ms（不含字体加载） |
| 打开 5 MB docx → 全文可见 | < 1.5 s |
| 打开 20 MB docx → 主线程单次任务占用 | < 50 ms（解析全程在 Worker） |
| 富文本编辑输入延迟 | < 16 ms（每帧一次 DOM patch） |
| Markdown ⇄ 富文本 往返 | 块级无损（行内样式保留率 ≥ 95%） |
| docx 导出 → Word/WPS 打开 | 无格式错乱、无「内容已损坏」提示 |

**非目标（本轮明确不做）**：多人协同、光标共享、修订模式、批注、公式域、目录自动生成、页眉页脚、节（section）级页面设置。

---

## 2. 现状盘点：文档在项目里现在是什么

### 2.1 现有链路：文档被降级成「A 列的纯文本行」

```
用户拖入 report.docx
   → FileOpener.openFile()                       （common/fileOpener.ts）
   → parse(): extension='docx' → new DocumentParser().parse(file)
   → DocumentParser 抽出 word/document.xml 里每个 <w:p> 的 <w:t> 文本
   → result.lines: string[]
   → 每行包成一个 Cell，落在 A1 / A2 / A3 …
   → applyDocument() 写进 values 仓，activedSheetName = 存档键
   → 主页面仍然渲染成「表格」
```

即：**当前没有文档视图**。`.text` / `.demo` 后缀的原生文件，打开后呈现的是一张只有 A 列有字的空表格。

### 2.2 逐项差距

| 能力 | 表格侧现状 | 文档侧现状 | 差距性质 |
| --- | --- | --- | --- |
| 渲染 | Canvas 5 层画布（`sheets/`） | 无 | **需新建** |
| 编辑 | `EditCanvas` + textarea 承接输入 + IME | 无 | **需新建** |
| 数据模型 | `Cell`（含 `Char[]` 字符级样式） | 只有 `string[]` | **需新建** |
| 样式 | `CellStyleTable` + 单元格级 + 字符级双层 | 无 | **需新建** |
| 工具栏 | 6 个页签 / 5 个分区 | 只复用外壳，内容面向单元格 | **需新建内容** |
| 持久化 | `values` 仓，存档 v4 分片 JSON | 与表格共用同一份 `values` 结构 | **需隔离** |
| 解析 | `SheetParser` 三步 + Worker | `DocumentParser` 纯文本 + 主线程 DOMParser | **需升级** |
| 导出 | 「另存为 xlsx」为模拟进度条 | `saveAsDocx` 菜单项 `disabled: true` | **需新建**（且缺 zip 写入器） |
| 打印 | `printer.ts` 表格分页 | 无 | **需新建** |
| Python | `sheet.*` 命名空间 | 无 | **需新建** |

### 2.3 三处「必须先拆的地基」

在写任何文档功能之前，有三处耦合必须先松动，否则文档与表格只能二选一：

**① `Toolbar` 的分区直接绑定 `DataCollection`**

`sectionFont.ts` 的做法是直接调 `this.data.setSelectedCharsFont(...)` / `setSelectedCellsFont(...)`；`sectionAlign.ts` 直接调 `setSelectedCellsFont({textAlign})` / `toggleMergeCells()`。字体分区里**没有任何一行是「通用的加粗按钮」** —— 它是「给选中单元格加粗」。

结论：**工具栏的 UI 外壳（`BtnBase` / `Select` / `ColorPicker` / `Menu`）可以 100% 复用，命令派发必须抽象出去。**

**② `App` 假定「主页面永远是一张表」**

`app.ts` 的 `initSheets()` 无条件创建 `Sheets`；`activateFile()` 无条件走 `data.switchToSheet()`；`setupDataListeners()` 监听的全是一张表的事件。文档视图需要在这三处各开一个分支。

**③ 存档键虽已分后缀，但仓是共用的**

`toArchiveKey(path)` 只去掉 `.grid`，保留其它后缀 —— 因此 `报表.grid` → 键 `报表`，`报表.text` → 键 `报表.text`，**两者天然不冲突**。但 `values` 仓的结构是「单元格数组」，塞不进块树。

---

## 3. 五个关键设计决策

### 决策 1：Markdown 的定位 —— 「交换格式 + 可编辑源码视图」，不是运行时真源

| 方案 | 描述 | 评价 |
| --- | --- | --- |
| **A. Markdown 为唯一真源** | 存盘只存 Markdown 字符串，富文本是它的渲染 | 实现最省；但**丢掉 Word 语义**（字号、颜色、表格合并、多级列表编号），docx 往返必然降级；且无法表达「同一段里前 3 个字是蓝色」 |
| **B. 块树为真源，Markdown 是导入导出格式之一** ⭐ | 内部是「块 + 行内 run」树；Markdown 作为 `import/export` 通道 | 富文本能力完整；docx → 块树 → docx 往返保真度高；代价是要写三条转换（md↔tree、docx↔tree） |
| **C. 双轨并存** | 一份 Markdown、一份块树，双向同步 | 双份真源必然出现同步分歧，不考虑 |

**选 B。** 理由：用户要「Word 功能」，Word 的语义模型就是「段落 + 行内 run 级样式」，块树是这个模型的直接映射；而 Markdown 是**面向人的交换格式**，天然是导入导出的载体 —— 它同时满足「有 Markdown」和「不牺牲富文本」。

在此之上额外提供**Markdown 源码视图**（分栏或切换式编辑器），满足「像 Typora 一样直接写 Markdown」的诉求。该视图是**块树的一个投影**，双向可编辑（P2）。

### 决策 2：渲染层 —— 文档走 DOM，不走 Canvas

表格用 Canvas 是因为「十万个同构单元格」的场景下 DOM 节点数会成为瓶颈。文档**不具备这个特征**：

| 维度 | 表格（Canvas 的理由成立） | 文档（DOM 更合适） |
| --- | --- | --- |
| 节点规模 | 10⁵～10⁶ 个单元格 | 10²～10⁴ 个块/run |
| 排版复杂度 | 每格单行或整体换行，规则简单 | **行内混排**、中西文混排、行内样式、混排基线对齐 —— 极难自研 |
| 选区 | 矩形选区，规则简单 | 任意字符区间、跨块选区、DOM Range |
| 输入法 | 单元格内单行编辑 | 长文本 IME 组字、候选窗定位 |
| 命中测试 | 坐标算行列 | 浏览器 `Range`/`caretRangeFromPoint` 免费给 |

**结论**：文档视图用 DOM，`contenteditable` + `document.execCommand` 已被证明会产出「浏览器归一化后的 DOM ≠ 模型」的经典难题，因此采用 **受控渲染**（模型 → DOM diff，DOM 不反写模型），输入走 `beforeinput` 事件 + `Selection` API 自行落点。

> ⚠️ 这与项目「Canvas 高性能渲染」的口号并不冲突：口号针对的是**表格**。文档是另一种载体，用另一个渲染器是分层架构的正常结果。

### 决策 3：文档存储 —— 新增 `docs` 对象仓，与 `values` 完全隔离

`values` 仓的结构是「单元格数组 + 行列头 + 打印设置」，与块树不兼容。把两者塞进一个仓会带来「读到不知道是哪种格式」的分支地狱。

**新增第四个仓 `docs`**，键规则与 `values` 一致（`toArchiveKey(path)` 纯函数），因此：

- `报表.grid` → values 键 `报表`，docs 键 `报表`（**注意会撞**）

需要修正：`toArchiveKey` 目前只去 `.grid` 后缀，所以 `报表.grid` → `报表`，而 `报表.text` → `报表.text`。两者**不撞**（前者去后缀、后者留后缀）。但若将来出现 `报表.md` 与 `报表.grid`，键分别是 `报表.md` 与 `报表` —— 仍不撞。**结论：现有 `toArchiveKey` 规则已足够，无需改动。**

新增仓需要对 `IndexDB` 做两件事（DB_VERSION 2 → 3）：

- `STORE_NAMES.DOCS = 'docs'`，`keyPath: 'docKey'`，索引 `{ name: 'docKey', unique: true }`
- `setDoc / getDoc / deleteDoc / getAllDocs` 四个方法，**结构化克隆直存**（与 `images` 库同款，不做 JSON.stringify）

`createObjectStores` 是幂等的（`if (!database.objectStoreNames.contains(...))`），升级时只需把 `DB_VERSION` 从 2 改 3 并加入新仓定义，**存量数据零迁移**。

### 决策 4：docx 解析沿用三步架构，但换掉 DOMParser

项目已有铁律：`xlsx` 的 XML 一律走 `xmlScanner.scanXml`，禁用 `DOMParser`（`MEMORY.md` §4）。而现有 `documentParser.ts` 恰恰用了 `DOMParser`（第 207 行）—— 它只在解析小文本文件时没暴露问题，一旦做真正的 docx 富文本解析，`DOMParser` 的「整棵驻留 DOM」会成为大文件的内存峰值来源，且**无法在 Worker 里运行**（`DOMParser` 是 Window 专有 API）。

**因此文档解析必须从第一天就走 `xmlScanner`**，并复用 `parser.worker.ts` 通道新增 `kind: 'doc'` 任务类型。

### 决策 5：导出 docx 需要先补一个 zip 写入器

项目只有 zip **读**（`utils/parser/zip.ts` 的 `ZipParser`）和 inflate（`inflate.ts`，走原生 `DecompressionStream('deflate-raw')`），**没有 zip 写**。而 docx 就是一个 zip 包。

- 浏览器原生有 `CompressionStream('deflate-raw')` → 可用于压缩条目
- 中央目录 / local header / EOCD 需自行拼装（`ZipParser` 的解析逻辑反向即可，约 150 行）
- 现有 `inflate.ts` 已有原生 `DecompressionStream` 优先、纯 JS 回退的双路径范式，压缩侧照抄即可

**结论：新增 `utils/parser/zipWriter.ts`，作为 docx 导出的前置依赖。** 这是整个方案里唯一一块「与 Word 功能无直接关系但必须先做」的基建。

---

## 4. 目标架构

```
┌──────────────────────────────────────────────────────────────┐
│                        App 层（app.ts）                        │
│  WorkspaceRouter：按文件后缀决定激活 Sheet / Doc 工作区       │
│  FileTree / QuickAccess / Printer 三个外壳对两种工作区都存在  │
├───────────────────────────┬──────────────────────────────────┤
│   SheetWorkspace（现有）   │   DocWorkspace（新建）             │
│   Sheets + 5 块 Canvas    │   DocView（DOM）+ BlockTree 模型   │
│   DataCollection           │   DocStore                        │
├───────────────────────────┴──────────────────────────────────┤
│              工具栏层（toolbar/）：外壳共用，内容按工作区装配   │
│  ToolbarShell（tabs/root-compact/section-nav）                 │
│    ├── SheetSections（sectionFont / sectionAlign / …）         │
│    └── DocSections（sectionDocFont / sectionDocBlock / …）     │
├──────────────────────────────────────────────────────────────┤
│              命令总线（core/command/）：新引入                  │
│  CommandBus：Section 只发命令，Host 决定谁执行                 │
├──────────────────────────────────────────────────────────────┤
│              持久化（indexDB.ts）：新增 docs 仓                 │
│  settings │ values（表格存档 v4）│ images │ docs（块树 + md）  │
├──────────────────────────────────────────────────────────────┤
│              解析层（utils/parser/）                           │
│  SheetParser（现有）│ DocParser（新，走 xmlScanner + Worker）  │
│  ZipParser（读）│ ZipWriter（新）│ CFB / inflate / xmlScanner  │
└──────────────────────────────────────────────────────────────┘
```

### 4.1 命令总线（新增，最关键的解耦点）

现状 `SectionFont` 直接调 `data.setSelectedCellsFont(...)`，文档工具栏无法复用。新增一层薄抽象：

```ts
// core/command/commandBus.ts
export interface EditHost {
    readonly kind: 'sheet' | 'doc';
    canExecute(cmd: string): boolean;
    execute(cmd: string, payload?: unknown): void;
    queryState(cmd: string): unknown;   // 供按钮高亮/置灰
}

export class CommandBus {
    private host: EditHost | null = null;
    bind(host: EditHost): void;
    dispatch(cmd: string, payload?: unknown): void;
    state(cmd: string): unknown;
}
```

改造后的 `SectionFont` 分成两层：

| 层 | 内容 | 复用性 |
| --- | --- | --- |
| **UI 外壳** | 字体下拉、字号下拉、粗/斜/下划/删除线按钮、文字色、背景色 | 表格与文档**共用同一份组件** |
| **命令派发** | `bus.dispatch('toggleBold')` | 由 `SheetHost` / `DocHost` 各自实现 |

`SheetHost` 内部转调现有 `data.setSelectedCharsFont(...)` / `setSelectedCellsFont(...)`；`DocHost` 内部改块树里选区 run 的样式。**现有表格行为零变化**。

### 4.2 工作区路由

`App` 新增 `WorkspaceRouter`，依据文件的**原生后缀**路由（`utils/fileName.ts` 已有映射）：

| 原生后缀 | 载体 | 激活工作区 |
| --- | --- | --- |
| `.grid` | 表格 | `SheetWorkspace` |
| `.text` | 文档 | `DocWorkspace` |
| `.demo` | 演示文稿 | （本轮不实现，暂回退为表格模式并提示） |

路由点共 4 处，全部收敛到一个 `switchWorkspace(kind)` 方法：

1. `initSheets()` → 改为 `initWorkspaces()`，两套视图都建好，按需显隐（`hidden` 属性 + `[hidden]{display:none}`，项目已有此约定）
2. `activateFile(path)` → 先按后缀判断，再走对应装载路径
3. `setupDataListeners()` → 表格事件订阅随工作区切换解绑/重绑
4. `FileOpener.applyDocument()` → 文档文件走 `DocStore` 装载，不再写 `values`

---

## 5. 文档数据模型

### 5.1 块（Block）与行内 run

```
Document
├── meta: { title, view: 'rich' | 'markdown', pageSetup }
├── blocks: Block[]
└── (派生) markdown: string        // 由 blocks 序列化而来，不单独存储
```

**Block**（块级）：

| 字段 | 类型 | 说明 |
| --- | --- | --- |
| `id` | `string` | 稳定标识，撤销/协同/增量 patch 用 |
| `type` | `'paragraph' \| 'heading' \| 'listItem' \| 'quote' \| 'code' \| 'table' \| 'image' \| 'hr'` | 块类型 |
| `level` | `number` | heading 的 1–6 级；listItem 的列表层级 0–4 |
| `listStyle` | `'bullet' \| 'ordered' \| 'none'` | 列表样式（对应 docx `numPr`） |
| `align` | `'left' \| 'center' \| 'right' \| 'justify'` | 段落对齐（docx `w:jc`） |
| `indent` | `number` | 缩进级数 |
| `runs` | `Run[]` | 行内内容 |
| `children` | `Block[]` | 代码块 / 表格单元格的子块 |
| `rows` / `colSpan` / `rowSpan` | | 仅表格用（docx `w:gridSpan` / `w:vMerge`） |
| `imageRef` | `{ assetId, width, height }` | 仅图片用 |

**Run**（行内，字符级样式）：

| 字段 | 对应 docx | 对应 Markdown |
| --- | --- | --- |
| `text` | `w:t` | 文本 |
| `bold` | `w:rPr/w:b` | `**…**` |
| `italic` | `w:rPr/w:i` | `*…*` |
| `underline` | `w:rPr/w:u` | `<u>…</u>`（扩展语法） |
| `strike` | `w:rPr/w:strike` | `~~…~~` |
| `color` | `w:rPr/w:color@w:val` | `{#RRGGBB}`（扩展语法） |
| `highlight` | `w:rPr/w:highlight@w:val` | `==…==`（扩展语法） |
| `fontSize` | `w:rPr/w:sz`（半磅） | `^字号^`（扩展语法） |
| `fontFamily` | `w:rPr/w:rFonts@w:ascii` | （扩展语法） |
| `vertAlign` | `w:rPr/w:vertAlign` | `^…^` / `~…~` |
| `link` | `w:hyperlink` | `[文字](url)` |

> **设计取舍**：Markdown 标准语法覆盖不了颜色 / 字号 / 字体 / 高亮。为了「Markdown 往返无损」，本方案约定一组**扩展语法**（用 `{#…}`、`^…^`、`==…==`），并在导出时保证标准 Markdown 阅读器也能正常显示（扩展标记退化为纯文本或被安全忽略）。若不接受扩展语法，则这些样式在 Markdown 通道中降级丢失 —— 需要产品侧拍板。**默认建议：接受扩展语法。**

### 5.2 与表格 `Char` 的关系

表格的 `Char` 已有 `fontFamily / fontSize / fontWeight / fontStyle / fontColor / underline / strikethrough` 七个样式字段 —— 与 `Run` 的字段集**高度重合**。

**建议：`Run` 不复用 `Char` 类**，但**复用字段命名**。理由：

- `Char` 是「一个字符」，`Run` 是「一段同样式文本」，粒度不同
- `Char` 被 `values` 存档格式深度绑定（`Char.toJSON` / `fromRaw` 的稀疏字段约定），改动会波及表格存档 v4 的兼容性
- 但字段名一致，能让「样式面板」的读写逻辑在两套工作区间复用

### 5.3 文档存档格式（`docs` 仓）

```ts
interface DocSnapshot {
    v: 1;
    meta: { title: string; pageSetup: PageSetup };
    blocks: Block[];       // 结构化克隆直存，不 JSON.stringify
    markdown?: string;      // 可选：缓存最近一次导出的 Markdown（纯派生，可随时重算）
}
```

与表格存档 v4 的取舍对照：

| 维度 | 表格 `values` | 文档 `docs` |
| --- | --- | --- |
| 体量 | 10⁵～10⁶ 单元格 → **必须分片** | 10²～10⁴ 块 → **不分片** |
| 序列化 | 分片 JSON（4000 格/片）+ `SparseAxis` | 结构化克隆直存（`IndexedDB` 原生支持嵌套对象） |
| 样式表 | `CellStyleTable` 显式表 + 下标引用 | 块树内联样式（块数少，内联成本可接受） |

**结论：文档不做分片。** 表格的分片机制是为「百万级字符串拼接导致主线程长任务」设计的，文档规模下引入分片只会徒增复杂度。

---

## 6. 文档工作区实现

### 6.1 渲染层：为什么文档不用 Canvas

见 [决策 2](#决策-2渲染层--文档走-dom不走-canvas)。落到实现上：

```
DocView（core/document/docView.ts）
├── 页面容器（.doc-page，纸张宽度 + 页边距）
├── 块渲染器：blocks → DOM（增量 patch，不整页重建）
├── 选区层：自绘选区矩形 + 跨块选区
├── 光标层：自绘 caret（闪烁）
└── 滚动/缩放：复用 Scroller 的滚动模型，缩放用 CSS transform
```

**增量 patch 策略**：每块一个 DOM 节点，`data-block-id` 标记。变更时只对受影响的块做「重建该块子节点」或「原地改属性」，不动其它块。整页重建在 10⁴ 块时也才 10 ms 级，但输入时每帧重建会导致**光标丢失与滚动跳动**，因此必须增量。

### 6.2 编辑层：选区、输入法、撤销

| 能力 | 方案 |
| --- | --- |
| 光标落点 | `document.caretRangeFromPoint()` / `caretPositionFromPoint()` + 自建 `offset ↔ (blockId, runIndex, charOffset)` 映射表 |
| 选区 | 浏览器 `Range` 对象 + 自建逻辑选区（跨块时首块尾块各持一段） |
| 输入 | `beforeinput` 事件拦截 → 走模型 → 重绘 → 用 `Selection.setBaseAndExtent` 复位光标 |
| 中文 IME | `compositionstart/update/end` 三态：组字期间**只更新插入符预览，不改模型**；`compositionend` 落一次模型变更。**不可用 `beforeinput` 逐字符处理组字**（会打断候选窗） |
| 复制/剪切/粘贴 | `ClipboardEvent`：`text/plain` 走纯文本；内部格式用自定义 MIME `application/x-silkspaces-doc`（存块树 JSON） |
| 撤销 | 复用 `utils/undoManager.ts` 的 `UndoManager`，新增 `type: 'docBlocks'` 的快照类型。**不复用 `DataCollection.runWithFullStateUndo`**（那是单元格语义） |
| 格式刷 | 复用 `SectionBrush` 的交互模式，但刷的是「run 样式集」 |

**关键约束**：`UndoManager` 的 `UndoRestoreCallback` 是注入式的（`undoManager.undo(callback)`），因此 `DocStore` 可以**共用同一个 `UndoManager` 类**、持有**自己的实例**，把表格与文档的撤销栈彻底隔开。

### 6.3 工具栏：命令总线与分区装配

**外壳复用**：`ToolbarShell`（tabs 行 + `section-nav` + root-compact 折叠 + `measureTabsNeed` 量测）完全复用现有 `toolbar.ts`，不改动。

**内容分区**（新建 `core/toolbar/doc/`）：

| 页签 | 分区 | 内容 |
| --- | --- | --- |
| **开始** | `sectionDocFont` | 字体、字号、粗、斜、下划、删除线、文字色、高亮色（**复用 `sectionFont` 的 UI 外壳**） |
| | `sectionDocBlock` | 段落对齐、行距、缩进、列表（项目符号/编号）、引用、分割线、清除格式 |
| | `sectionDocStyle` | 标题 H1–H6、正文、代码块 |
| **插入** | `sectionDocInsert` | 表格、图片、链接、代码块、水平线、日期 |
| **视图** | `sectionDocView` | 富文本 ⇄ Markdown 源码切换、显示格式标记、缩放、专注模式 |
| **页面** | `sectionDocPage` | 纸张、页边距、方向（**复用 `printer.ts` 的 `PrinterSetting` 结构**） |

**装配时机**：`Toolbar.setTabContent()` 现在在 `init()` 里一次性写死。改为**按工作区切换时重建**：

```ts
// 表格 → 文档 切换时
toolbar.clearTabContents();               // 清空各 group 的内容（保留 tabs 行与折叠状态）
toolbar.setTabContent('开始', [ ...docSections ]);
```

`Toolbar` 的 `groups` / `navs` 两个 `Map` 与 `rootTabs` 数组需要支持「重复装配」：同一页签文本重复 `setTabContent` 时，先清空该 group 的子节点再挂载（现有实现会直接 append，重复调用会叠加 —— 这是改造时要顺手修掉的一个既有隐患）。

**图标**：`assets/images/toolbar/` 已有 `bold.svg` / `italic.svg` / `text-left.svg` / `text-center.svg` / `text-right.svg` / `justify.svg` / `text-indent-left.svg` / `text-indent-right.svg` / `table.svg` / `line-break.svg` 等，**大部分可直接复用**；需新增：`quote.svg`、`code.svg`、`h1.svg`、`bullet-list.svg`、`ordered-list.svg`、`link.svg`、`image.svg`、`markdown.svg`、`hr.svg`。

### 6.4 存储层：新增 docs 对象仓

见 [决策 3](#决策-3文档存储--新增-docs-对象仓)。落地清单：

| 改动点 | 内容 |
| --- | --- |
| `indexDB.ts` | `DB_VERSION` 2 → 3；新增 `STORE_NAMES.DOCS`；`OBJECT_STORE_CONFIG` / `OBJECT_STORE_INDEXES` 各加一条；新增 4 个方法 |
| `fileName.ts` | **不改**（`toArchiveKey` 规则已足够分键） |
| `fileTree.ts` | `FileContentBridge` 增加文档侧的 `copyDoc / moveDoc / removeDoc`（或复用同一组回调，内部按后缀分派） |
| `app.ts` | `copyFileContents` / `moveFileContents` / `removeFileContents` / `ensureFileContents` 需**同时处理 values 与 docs 两个仓** |

**落库节奏**：沿用表格的防抖 + 空闲落库（`AUTO_SAVE_DEBOUNCE_MS = 500`）。文档体量小，防抖窗口可短一些（建议 300 ms），编辑体验更好。

### 6.5 导入导出：docx / Markdown

#### 6.5.1 docx 导入（`utils/parser/docParser.ts`）

沿用 `SheetParser` 的三步法，改造成文档版：

| 步骤 | 表格版 | 文档版 |
| --- | --- | --- |
| **decode** | 解 zip → 读 `xl/*.xml` | 解 zip → 读 `word/document.xml`、`word/styles.xml`、`word/numbering.xml`、`word/_rels/document.xml.rels`、`word/media/*` |
| **extract** | 扫 XML → 列式 `PackedSheet` | 扫 XML → `Block[]` + `Run[]` + 图片资产列表 |
| **generate** | 列式 → `Cell`/`Char` 实例 | 块树 → `DocSnapshot`（Worker 内产出的应当是**纯数据**，回主线程后再实例化，避免跨线程丢原型 —— 与现有 `generatePacked` 的处理一致） |

**需要解析的 docx XML 要素清单**：

| 要素 | 含义 | 映射到 |
| --- | --- | --- |
| `w:p` | 段落 | Block |
| `w:pPr/w:pStyle@w:val` | 段落样式名 | 经 `styles.xml` 查 `w:style[@w:styleId]` → `w:name` → 映射为 `heading`/`quote`/`code` 等 |
| `w:pPr/w:jc@w:val` | 对齐 | `align` |
| `w:pPr/w:numPr`（`w:numId` + `w:ilvl`） | 列表编号 | `listStyle` + `level`（经 `numbering.xml` 判 bullet/ordered） |
| `w:pPr/w:ind@w:left` | 缩进 | `indent` |
| `w:r` | 行内 run | Run |
| `w:rPr/w:b` `w:i` `w:u` `w:strike` `w:vertAlign` | 字符样式 | Run 对应字段 |
| `w:rPr/w:color@w:val` | 颜色 | `color`（注意 `auto` / 主题色 `themeColor` 要落到具体 RGB） |
| `w:rPr/w:sz@w:val` | 字号（半磅） | `fontSize`（÷2） |
| `w:rPr/w:rFonts@w:ascii` | 字体 | `fontFamily` |
| `w:rPr/w:highlight@w:val` | 高亮 | `highlight` |
| `w:t` | 文本（注意 `xml:space="preserve"`） | `Run.text` |
| `w:tab` / `w:br` | 制表 / 换行 | `\t` / `\n` |
| `w:tbl` / `w:tr` / `w:tc` | 表格 | Block `table`（`w:gridSpan` → `colSpan`，`w:vMerge` → `rowSpan`） |
| `w:hyperlink@r:id` | 超链接 | 经 rels 解析 URL → Run `link` |
| `w:drawing` → `wp:inline` → `a:blip@r:embed` | 图片 | 经 rels 找 `word/media/*` → `imageRef` |

**必须注意的坑**：

1. `w:t` 的 `xml:space="preserve"` —— 前后空格不能被 XML 解析器吃掉
2. `styles.xml` 的 `w:docDefaults` 提供默认字号/字体，块上没有显式样式时要回落到这里
3. 主题色（`w:themeColor`）需要 `theme1.xml` 解析才能落成 RGB；本轮可先降级为 `auto`（黑色），P2 补齐
4. `w:sectPr`（节属性，含页面尺寸/页边距）应解析进 `meta.pageSetup`，这样「页面」页签才有内容
5. `xmlScanner.scanXml` 的 `attrs` 参数**仅回调期间有效**（扫描器复用该对象），取值必须在回调内取出 —— 这条已在 `xmlScanner.ts` 的注释里写明，解析器实现时务必遵守

#### 6.5.2 Markdown 导入导出

**自研，不引第三方库**（项目 `dependencies` 目前只有 `pyodide`，保持一致）。

| 方向 | 范围 |
| --- | --- |
| md → 块树 | 标题（ATX `#`~`######` + Setext）、段落、有序/无序/任务列表、引用、代码块（围栏）、水平线、表格（GFM pipe table）、行内：`**粗**` `*斜*` `~~删~~` `` `码` `` `[链接](url)` `![图](url)` + 本方案的扩展语法 |
| 块树 → md | 上述的逆变换，行内样式按「连续同 run 合并」输出，避免 `**a****b**` 这种碎裂 |

**扩展语法的降级策略**：导出的 md 里，颜色写成 `文字{#FF0000}`。标准 Markdown 阅读器会原样显示这段字（不报错、不丢内容），本项目自己的阅读器解析扩展语法后还原样式。

#### 6.5.3 docx 导出

依赖新增的 `ZipWriter`（见 [决策 5](#决策-5导出-docx-需要先补一个-zip-写入器)）。产出的包结构（最小可用集）：

```
[Content_Types].xml
_rels/.rels
word/document.xml
word/styles.xml
word/_rels/document.xml.rels     （仅当有图片/超链接时）
word/media/image1.png            （仅当有图片时）
docProps/core.xml、app.xml        （可选）
```

`word/document.xml` 由块树序列化而来（`w:p` / `w:r` / `w:pPr` / `w:rPr` / `w:tbl` / `w:drawing`）。

**与「另存为」对话框的衔接**：`fileTree.ts` 的 `saveAsDocx` 菜单项目前是 `disabled: true`，本轮把它接上，并复用现有 `showSaveAsXlsxDialog()` 的清单式对话框骨架（表格行 + 转换进度 + 单个/批量下载）。

> 现有的「另存为 xlsx」是**模拟进度条**（`SaveAsRowState` 的 `converting/paused/done` 由定时器推进），并非真实导出。文档导出要不要做成真导出，需要产品侧明确；建议 **文档侧做真导出**（价值高、技术路径清晰），表格侧的 xlsx 导出保持现状不动。

### 6.6 Python 编程接口

沿用 `utils/python/sheetNameSpace.ts` 的模式，新增 `docNameSpace.ts`，在 Worker 里注入 `doc` 命名空间：

```python
doc.title                      # 文档标题
doc.blocks                     # 块列表（只读视图）
doc.text                       # 全文纯文本
doc.append("段落文本")          # 追加段落
doc.insert_markdown("# 标题\n正文")   # 插入 Markdown（自动解析为块树）
doc.find("关键词")              # 返回命中位置列表
doc.set_style(block_index, bold=True)   # 设置块/行内样式
doc.export_markdown()          # 返回 Markdown 源码
```

实现要点（复用现有约定）：

- **快照分片**：文档块数少，但仍按 `chunkObjectJson` 的 4000 单位分片传递（保持通道协议一致，避免 Worker 侧出现第二套协议）
- **写回合并**：与 `applyWritesAll` 同款策略 —— 文档侧只调一次 `applyDocWrites`，撤销栈不碎片化
- **`usesFilesBinding` 判定**：文档的 `open_doc()` 同样会触发预加载，判定逻辑必须走同一个 `usesFilesBinding`（先剥注释），否则默认模板的 API 注释会命中裸 `/\bfiles\b/` 而预加载整个工作区
- **与表格脚本共存**：同一份 `py-editor:script` 存档键不变，但脚本内容按当前工作区注入不同的命名空间（`sheet` vs `doc`）

### 6.7 打印

表格打印走 `utils/printer.ts` 的「按行列切页 → 生成 HTML table → `@page` CSS」。文档分页规则完全不同（按内容流自然分页），**不应复用 `printer.ts` 的分页算法**，但**应复用其设置结构**：

| 复用 | 新建 |
| --- | --- |
| `PrinterSetting`（纸张 / 方向 / 页边距 / 缩放 / 双面 / 页码范围） | 文档分页算法（按块流高度累加 + 孤行寡行控制） |
| `PAPER_SIZES` / `MARGIN_PRESETS` 常量 | 文档页眉页脚（P3） |
| 打印预览浮层（`gs-print-overlay` 样式族） | 文档预览页的块渲染（复用 `DocView` 的渲染器，纸张宽度固定） |
| `@page { size; margin }` 输出机制 | |

---

## 7. 实施路线图

### P0 — 打通「文档是一种载体」（无富文本，先有骨架）

| # | 任务 | 涉及文件 | 验收 |
| --- | --- | --- | --- |
| 1 | `docs` 对象仓 + 4 个存取方法 | `dataArchitecture/indexDB.ts` | DB_VERSION 升 3，存量表格存档全部可读 |
| 2 | 文档数据模型（`Block` / `Run` / `DocSnapshot`） | 新建 `core/document/model.ts` | 类型完备，`tsc --noEmit` 通过 |
| 3 | `DocStore`（类 `DataCollection` 的文档侧：装载、脏标记、防抖落库、切换） | 新建 `core/document/docStore.ts` | 切换文档不丢表格内容 |
| 4 | `DocView` 只读渲染（块 → DOM） | 新建 `core/document/docView.ts` | 段落/标题/列表/引用/代码块/表格正确渲染 |
| 5 | `WorkspaceRouter` + `App` 四处路由点改造 | `core/app.ts` | `.grid`/`.text` 双击分别进各自工作区，互不串味 |
| 6 | `Toolbar` 分区装配改造 + 文档「视图」分区（视图切换/缩放） | `core/toolbar/toolbar.ts`、`core/toolbar/doc/` | 切换工作区工具栏内容正确换装 |
| 7 | `FileTree` / `App` 的文档侧增删改查（复制/移动/删除/重命名随文档走） | `common/fileTree.ts`、`core/app.ts` | 文件树操作后文档内容不丢 |

### P1 — 富文本编辑 + Markdown 视图

| # | 任务 | 验收 |
| --- | --- | --- |
| 8 | 命令总线 `CommandBus` + `EditHost` 接口，改造 `sectionFont` 为「外壳 + 派发」 | 表格侧字体功能行为零变化（跑通既有测试） |
| 9 | 编辑层：光标 / 选区 / `beforeinput` / IME 三态 | 中文输入法组字不丢字、不跳字 |
| 10 | 文档「开始」「插入」分区（字体、段落、样式、插入） | 工具栏按钮真实生效，按钮高亮反映当前选区状态 |
| 11 | 文档撤销/重做（`UndoManager` + `docBlocks` 快照） | Ctrl+Z / Ctrl+Y 正确回退块树与选区 |
| 12 | Markdown 序列化 + 反序列化（自研，含扩展语法） | 往返无损（行内样式保留率 ≥ 95%） |
| 13 | Markdown 源码视图（分栏编辑器，复用 `pyEditor` 的 `pre + textarea` 叠层 + 高亮范式） | 源码可编辑，改动同步回块树 |
| 14 | 复制/剪切/粘贴（纯文本 + 内部块树 MIME） | 跨应用粘贴保留富文本 |

### P2 — docx 互操作 + 打印 + Python

| # | 任务 | 验收 |
| --- | --- | --- |
| 15 | `ZipWriter`（`CompressionStream('deflate-raw')` + 中央目录拼装） | 能产出合法 zip，`ZipParser` 可回读 |
| 16 | `DocParser` 三步法（decode/extract/generate），走 `xmlScanner` + Worker | 5 MB docx 打开 < 1.5 s，主线程单任务 < 50 ms |
| 17 | docx 富文本解析（样式 / 列表 / 表格 / 超链接 / 图片 / sectPr） | 见 §6.5.1 要素清单逐项覆盖 |
| 18 | docx 导出（块树 → `word/document.xml` → zip） | Word / WPS 打开无警告、样式正确 |
| 19 | 「另存为 docx」接入 `fileTree` 的清单对话框 | 批量下载可用 |
| 20 | 文档打印与打印预览（复用 `PrinterSetting`，新写分页） | A4 / 横向 / 页边距预设生效 |
| 21 | Python `doc` 命名空间 + `open_doc()` + 写回 | 脚本能读写文档，落点进撤销栈 |
| 22 | 主题色（`theme1.xml`）解析、页眉页脚、节级页面设置 | Word 还原度高 |

### P3 — 增强（不排期）

修订模式 / 批注 / 查找替换（复用表格的 `find/replace` 引擎，但作用于块树）/ 表格内嵌图片 / 文档协同（预留 `Block.id` + `patch` 协议）/ 演示文稿 `.demo` 工作区。

---

## 8. 风险与已知坑

| # | 风险 | 影响 | 缓解 |
| --- | --- | --- | --- |
| 1 | **受控富文本编辑器是本方案最大的不确定性** | contenteditable 的浏览器归一化、IME 组字、光标复位都可能反复出问题 | 严格「模型为准、DOM 只读渲染」；输入走 `beforeinput`；IME 组字期间只更新插入符不落模型；先做 P0 只读骨架把数据链路跑通再上编辑 |
| 2 | **`DOMParser` 依赖残留** | `documentParser.ts` 现用 `DOMParser`，做富文本后必然 OOM 且无法进 Worker | P2 起强制改 `xmlScanner.scanXml`；P1 期间仅复用其纯文本结果时保留旧实现 |
| 3 | **zip 写入器缺失** | docx 导出被硬阻塞 | P2 第 15 项单独立项，提前做；`ZipParser` 解析逻辑反向拼装 |
| 4 | **表格与文档两套撤销栈共存** | 用户在两个工作区间切换时误撤销 | 两个 `UndoManager` 实例完全隔离；切工作区即清空另一侧的栈 |
| 5 | **`Toolbar.setTabContent` 重复调用会叠加** | 切换工作区后分区显示两份 | 改造为「重复装配先清 group 子节点」；这是既有隐患，顺手修 |
| 6 | **扩展 Markdown 语法不被外部工具识别** | 导出的 md 在别处显示异常 | 扩展语法设计成「不破坏标准渲染」的形式；提供「仅标准 Markdown」导出选项 |
| 7 | **文件树与 docs 仓的搬迁遗漏** | 移动/重命名文件后文档内容留在旧键 | `FileContentBridge` 的 copy/move/remove 三个回调同时处理两个仓；补一条守卫测试：移动 `.text` 文件后 `getDoc(新键)` 有内容、旧键为空 |
| 8 | **`reconcileValueStore` 的清理逻辑只认 values** | 删了文档文件后 docs 仓留孤儿存档 | 扩展清理逻辑同时扫描 `docs` 仓；注意**只删空存档**，非空存档要迁 |
| 9 | **`.demo` 无工作区** | 演示文稿文件打开后行为未定义 | 明确降级：暂按表格模式打开并提示「演示文稿尚未支持」；不要静默 |
| 10 | **大 docx 的图片资产** | `word/media/*` 解压后可能上百 MB | 图片按需读取（先取尺寸再决定是否解码）；沿用表格图片的 `createImageBitmap` 通道 |

### 8.1 与既有铁律的相容性自查

| 既有约定（`MEMORY.md`） | 本方案是否冲突 | 说明 |
| --- | --- | --- |
| 存档键 = `toArchiveKey(path)` 纯函数 | ✅ 不冲突 | 新增 docs 仓沿用同一键规则，不引入第二套换算 |
| 「同名不同后缀 = 两个独立文件」 | ✅ 不冲突 | `.grid` / `.text` 键本就不同 |
| xlsx XML 一律 `xmlScanner`，禁 `DOMParser` | ✅ 主动遵守 | 文档解析同样走 `xmlScanner` |
| 改 values 一律包 `runWithFullStateUndo` | ✅ 不适用 | 文档侧自建 `docBlocks` 快照，不碰表格撤销语义 |
| 禁全局 `getElementById` 找自身子元素 | ⚠️ 需注意 | 文档视图用 `querySelector` 找块时必须**容器内 scoped**，且不得用 `getElementById` |
| `el.classList` 只能 add/remove/toggle | ⚠️ 需注意 | `classList = 'x'` 会抛 TypeError |
| `hidden` 显隐须补 `[hidden]{display:none}` | ⚠️ 需注意 | 工作区切换用 `hidden`，父容器若有 `display:flex` 需补该规则 |
| 异步逻辑禁包进同步 undo 回调 | ⚠️ 需注意 | 文档 undo 回调内不得 `await` |
| 测试共享同一 `loadModule` 缓存 | ⚠️ 需注意 | 新模块的测试同样要共享缓存，否则 `Block`/`Run` 是另一模块实例 |
| 守卫写完**必须反证一次** | ⚠️ 需注意 | docs 仓、命令总线、工作区路由都要有守卫测试 |

---

## 9. 附录：接口草案

### 9.1 `core/document/model.ts`

```ts
export type BlockType =
    | 'paragraph' | 'heading' | 'listItem' | 'quote'
    | 'code' | 'table' | 'image' | 'hr';

export interface RunStyle {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strike?: boolean;
    color?: string;
    highlight?: string;
    fontSize?: number;
    fontFamily?: string;
    vertAlign?: 'superscript' | 'subscript';
    link?: string;
}

export interface Run {
    text: string;
    style?: RunStyle;
}

export interface Block {
    id: string;
    type: BlockType;
    level?: number;
    listStyle?: 'bullet' | 'ordered' | 'none';
    align?: 'left' | 'center' | 'right' | 'justify';
    indent?: number;
    runs: Run[];
    children?: Block[];
    rows?: Block[][];
    colSpan?: number;
    rowSpan?: number;
    imageRef?: { assetId: string; width: number; height: number };
}

export interface DocMeta {
    title: string;
    view: 'rich' | 'markdown';
    pageSetup: PrinterSetting;
}

export interface DocSnapshot {
    v: 1;
    meta: DocMeta;
    blocks: Block[];
    markdown?: string;
}
```

### 9.2 `core/command/commandBus.ts`

```ts
export interface EditHost {
    readonly kind: 'sheet' | 'doc';
    canExecute(cmd: string): boolean;
    execute(cmd: string, payload?: unknown): void;
    queryState(cmd: string): unknown;
}

export const DOC_COMMANDS = [
    'toggleBold', 'toggleItalic', 'toggleUnderline', 'toggleStrike',
    'setFontFamily', 'setFontSize', 'setColor', 'setHighlight',
    'setAlign', 'setIndent', 'setLineHeight',
    'setBlockType', 'toggleList', 'toggleQuote', 'insertLink',
    'insertTable', 'insertImage', 'insertCodeBlock', 'insertHr',
    'clearFormat', 'toggleView', 'setZoom',
] as const;
```

### 9.3 `utils/parser/docParser.ts`

```ts
export interface DocParserOptions {
    readImages?: boolean;
    readStyles?: boolean;
}

export interface DocParseResult {
    blocks: Block[];
    meta: DocMeta;
    images: ExtractedImage[];
}

export class DocParser {
    public async decode(file: File, options?: DocParserOptions): Promise<DecodedDoc>;
    public async extract(decoded: DecodedDoc, options?: DocParserOptions): Promise<DocParseResult>;
    public generate(result: DocParseResult): DocSnapshot;
}
```

### 9.4 `utils/parser/zipWriter.ts`

```ts
export interface ZipWriteEntry {
    path: string;
    data: Uint8Array;
    compress?: boolean;   // 默认 true（deflate-raw）
}

export class ZipWriter {
    public add(path: string, data: Uint8Array | string): void;
    public async toBlob(): Promise<Blob>;
}
```

### 9.5 `core/document/docStore.ts`

```ts
export class DocStore extends EventEmitter {
    public docKey: string;
    public blocks: Block[];
    public meta: DocMeta;
    public undoManager: UndoManager;
    public selection: DocSelection;

    public async open(docKey: string): Promise<boolean>;
    public async save(): Promise<void>;
    public toMarkdown(): string;
    public applyMarkdown(md: string): void;
    public runWithDocUndo(label: string, action: () => void): void;
}
```

---

## 10. 一页纸总结

| 维度 | 结论 |
| --- | --- |
| 真源 | **块树**（`Block[]` + `Run[]`），不是 Markdown 字符串 |
| Markdown | **交换格式 + 可编辑源码视图**；行内扩展语法保证往返无损 |
| 渲染 | **DOM**（受控渲染 + 增量 patch），不用 Canvas |
| 编辑 | `beforeinput` + `Selection` API + IME 三态；自绘选区与光标 |
| 工具栏 | **外壳复用**（tabs / section-nav / root-compact），内容按工作区换装；命令总线解耦 |
| 存储 | **新增 `docs` 仓**；键沿用 `toArchiveKey`；不分片 |
| 解析 | `xmlScanner` + Worker，沿用 decode/extract/generate 三步法 |
| 导出 | 先补 `ZipWriter`，再出 docx；`saveAsDocx` 菜单项接上 |
| 撤销 | 独立 `UndoManager` 实例 + `docBlocks` 快照类型 |
| Python | 新增 `doc` 命名空间，通道协议与表格同构 |
| 打印 | 复用 `PrinterSetting` 与 `@page` 机制，分页算法新写 |
| 最大风险 | 受控富文本编辑器的光标/IME 稳定性 |
| 前置依赖 | `ZipWriter`（docx 导出的硬前提） |
| 建议起步 | P0 七项：先把「文档是一种载体」的数据链路跑通，再上编辑 |
