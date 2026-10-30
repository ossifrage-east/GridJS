# 代码分离解耦优化 + Word 文档功能实现方案

> **文档性质：纯方案建议。** 本文不修改任何源码，仅基于对当前实现的阅读与度量给出设计与落地建议。
>
> 适用对象：`src/core/`（工作区装配层、工具栏层、数据架构层）、`src/utils/`（工具层、解析层）。
>
> 姊妹文档：`Excel多线程解析方案.md`（大表格解析链路）、`Word文档功能实现方案.md`（本文的前身，后按本文的「解耦优先 + 复用 chars」结论重写）。

---

## 目录

- [第一部分：现有代码的分离解耦优化](#第一部分现有代码的分离解耦优化)
  - [1. 现状度量](#1-现状度量)
  - [2. 六处结构性耦合](#2-六处结构性耦合)
  - [3. 解耦总纲：三层夹心架构](#3-解耦总纲三层夹心架构)
  - [4. 优化后的代码路径图](#4-优化后的代码路径图)
  - [5. 六项解耦改造详案](#5-六项解耦改造详案)
  - [6. 解耦改造的落地顺序与风险](#6-解耦改造的落地顺序与风险)
- [第二部分：Word 文档功能实现](#第二部分word-文档功能实现)
  - [7. 六项设计决策](#7-六项设计决策)
  - [8. 块树模型与 chars 复用](#8-块树模型与-chars-复用)
  - [9. 三条转换链路](#9-三条转换链路)
  - [10. 工具栏复用：UI 外壳共享](#10-工具栏复用ui-外壳共享)
  - [11. 持久层：docs 仓](#11-持久层docs-仓)
  - [12. 渲染层：DOM](#12-渲染层dom)
  - [13. 实施路线图](#13-实施路线图)
  - [14. 风险与已知坑](#14-风险与已知坑)
  - [15. 附录：接口草案](#15-附录接口草案)

---

# 第一部分：现有代码的分离解耦优化

## 1. 现状度量

### 1.1 文件规模（`wc -l`，全仓 52434 行 TS+SCSS）

| 行数 | 文件 | 归属 | 职责数（估） |
| ---: | --- | --- | ---: |
| 8686 | `core/dataArchitecture/dataCollection.ts` | 数据架构层 | **8+** |
| 2961 | `core/common/fileTree.ts` | 通用组件层 | **7+** |
| 2384 | `core/sheets/canvas.ts` | 渲染层 | **6+** |
| 2251 | `utils/parser/sheetParser.ts` | 解析层 | 3（三步法，尚可） |
| 1949 | `utils/printer.ts` | 工具层 | 4 |
| 1738 | `utils/python/sheetNameSpace.ts` | 工具层 | 1（可接受） |
| 1612 | `core/sheets/editCanvas.ts` | 渲染层 | 5 |
| 1191 | `core/dataArchitecture/cell.ts` | 数据架构层 | 2 |
| 1152 | `core/app.ts` | 应用层 | **8+** |
| 996 | `utils/parser/xlsxWriter.ts` | 解析层 | 3 |
| 910 | `core/toolbar/start/searchDialog.ts` | 工具栏层 | 3 |
| 839 | `core/assistant/filter.ts` | 助手层 | 3 |
| 510 | `core/dataArchitecture/indexDB.ts` | 数据架构层 | 1（可接受） |

**两个"巨型文件"**：`dataCollection.ts`（8686 行）占全仓 16.5%，`fileTree.ts`（2961 行）占 5.6%。二者合计 22%。

### 1.2 `dataCollection.ts` 的职责面

该类有 **370 个成员**（public/private/getter/setter）。按关键词归类统计其方法分布：

| 职责域 | 相关方法数 | 代表方法 |
| --- | ---: | --- |
| 单元格读写 / 索引 | 98 | `getCellByName` / `ensureWritableCell` / `invalidateCellIndex` |
| 合并单元格 | 19 | `toggleMergeCells` / `mergeCells` / `executeMergeCells` / `getMergedOutlineBorders` |
| 撤销 / 快照 | 23 | `runWithFullStateUndo` / `runWithCellsUndo` / `_applyFullStateSnapshot` |
| 筛选 / 隐藏 | 18 | `setFilter` / `validataFilter` / `setRowsHidden` / `restoreRowsHiddenByFilter` |
| 查找 / 替换 | 5+ | `findAllInRect` / `replaceInRect` / `resolveSearchRect` |
| 图片 | 5 | `_loadImages` / `images` |
| 持久化 | 4 | `syncValuesToDB` / `toSnapshotChunked` / `switchToSheet` |

**问题不在于方法多，而在于这些方法共享同一批私有字段**（`_values` / `_cellIndex` / `_mergedCells` / `_filterConditions` / `_hiddenRowsSet` / `_undoManager` …）。要拆分就必须先把这些字段的归属说清楚。

### 1.3 工具栏分区对 `DataCollection` 的依赖面

| 分区文件 | 调用的 `this.data.*` |
| --- | --- |
| `sectionFont.ts` | `activedCell` `cursorAbsolutePosition` `getCellByName` `isEditting` `on` `setSelectedCellsBorder` `setSelectedCellsFont` `setSelectedCharsFont` `syncValuesFromDB` |
| `sectionAlign.ts` | `activedCell` `changeSelectedCellsIndent` `getCellByName` `isEditting` `on` `setSelectedCellsFont` `syncValuesFromDB` `toggleMergeCells` `zoom` |
| `sectionBrush.ts` | `brushMode` `clipboardManager` `copyChars` `copySelection` `cutChars` `cutSelection` `cursorAbsolutePosition` `emit` `getSelectedCharsStartIndex` `isEditting` `on` `pasteChars` `pasteFormatOnly` `pastePlainText` `pasteSelection` `runWithFullStateUndo` `selectionOnEditting` `values` |
| `sectionProcess.ts` | `isEditting` `on` `setFilter` `sortDown` `sortUp` `syncValuesFromDB` `validataFilter` |
| `searchDialog.ts` | `cellMatches` `expandSingleCellRect` `findAllInRect` `formatRect` `getCellName` `gotoSearchHit` `off` `on` `replaceInRect` `resolveSearchRect` `selection` |

**共 51 个不同的 `data.*` 调用点。** 其中真正"表格专属"的只有约 12 个（`toggleMergeCells` / `setFilter` / `sortUp` / `sortDown` / `validataFilter` / `getCellByName` / `copySelection` / `pasteSelection` …），**其余约 39 个是"编辑宿主"的通用语义**：加粗、斜体、下划线、字体、字号、颜色、背景、对齐、缩进、复制、剪切、粘贴、撤销、监听变更。

> **这是本次解耦的核心发现：工具栏里 76% 的调用是通用的，24% 是表格专属的。** 现有的分区文件没有把这两类分开，导致每个分区都绑死 `DataCollection`。

### 1.4 工具栏装配点

`toolbar.ts` 的 `init()` 里硬编码：

```ts
this.setTabContent(TOOLBAR_TABS.START, [ sectionBrush, sectionFont, sectionAlign, sectionProcess ]);
this.setTabContent(TOOLBAR_TABS.PAGE,  [ sectionPrint ]);
```

`setTabs()` 遍历 `Object.values(TOOLBAR_TABS)`（开始/插入/页面/数据/视图/工具）创建 6 个页签，但**只装配了 2 个**，其余 4 个是空 group。

`setTabContent()` 的实现是**纯 append**（第 286 行起：`group.appendChild(section.el)`），**没有清空逻辑** —— 重复调用会叠加分区。这是切换工作区时必然踩到的坑。

### 1.5 其它已确认的事实

| 事实 | 位置 | 对解耦的影响 |
| --- | --- | --- |
| `Char` 有 7 个稀疏样式字段 | `char.ts` 27-41 | **可直接复用为文档 run 的样式载体**（见第二部分） |
| `FontStyle` 前 7 个字段与 `Char` **完全一致** | `cellContent.ts` 71-84 | 样式协议天然通用，无需新定义 |
| `Char.fromRaw` 是稀疏复制 | `char.ts` 98 | 复用时稀疏性不丢 |
| `Char` 非 `EventEmitter` | `char.ts` 25 | 复用无监听表开销 |
| `undoManager` 的恢复回调是**注入式** | `undoManager.ts` 295 | 文档可**共用类、持有独立实例** |
| `printer` 用 `getSharedPrinter(data)` 单例 | `printer.ts` | 打印设置也是"宿主"语义，应随宿主切换 |
| **`xlsxWriter.ts` 已有通用 ZIP 写入器** | `xlsxWriter.ts` 555 `createZipBlob(entries)` | **docx 导出无需新写 zip writer**（修正了前版方案的判断） |
| `fileTree.ts` 2961 行含 7 类职责 | — | 需要拆分 |
| `app.ts` 1152 行含 8 类职责 | — | 需要拆分 |
| `documentParser.ts` 用 `DOMParser` | `documentParser.ts` 207 | 与项目铁律冲突，且进不了 Worker |

---

## 2. 六处结构性耦合

| # | 耦合 | 现状 | 后果 | 严重度 |
| --- | --- | --- | --- | --- |
| **C1** | **工具栏分区 ↔ DataCollection** | 51 个 `data.*` 调用点直连 | 文档工具栏无法复用 UI；表格分区无法用于其它宿主 | 🔴 高 |
| **C2** | **App ↔ 单一工作区假设** | `initSheets` / `activateFile` / `setupDataListeners` / `FileOpener.applyDocument` 四处假定"主页面永远是表" | 无法并存第二种载体 | 🔴 高 |
| **C3** | **dataCollection 内部 8+ 职责混居** | 8686 行 / 370 成员 / 共享私有字段 | 改一处要读全文；无法独立测试 | 🟠 中高 |
| **C4** | **fileTree ↔ 内容桥语义过窄** | `FileContentBridge` 只有 create/copy/move/remove 四组回调，且都指向 `values`/`images` | 第二种载体（docs）的增删改查无处安放 | 🟠 中高 |
| **C5** | **Toolbar 装配不可重入** | `setTabContent` 纯 append，无清空 | 切换上下文必然叠加分区 | 🟡 中 |
| **C6** | **持久层只有一种内容形态** | `values` 仓结构固定为单元格数组 | 块树无处存放 | 🟡 中 |

---

## 3. 解耦总纲：三层夹心架构

```
┌──────────────────────────────────────────────────────────────┐
│  外壳层（Shell）—— 与内容类型无关，两种工作区共用              │
│  ToolbarShell · QuickAccess · FileTree · Navigator · Menu      │
│  Printer · Dialog · MessageBox · Split · Scroller            │
├──────────────────────────────────────────────────────────────┤
│  契约层（Contract）—— 新增，纯接口，零依赖                    │
│  EditHost · CommandBus · ContentStore<T> · WorkspaceRouter   │
├──────────────────────────────────────────────────────────────┤
│  实现层（Impl）—— 可插拔，各自独立                            │
│  SheetHost(DocHost) · SheetStore(DocStore) · SheetView(DocView) │
└──────────────────────────────────────────────────────────────┘
```

**核心思想**：把"**是什么**"（外壳 + 契约）与"**装什么**"（实现）彻底分开。契约层不引用任何具体实现，实现层只依赖契约层。依赖方向单向向下，零反向引用。

---

## 4. 优化后的代码路径图

### 4.1 目标目录结构

```
src/
├── core/
│   ├── app.ts                          ← 瘦身为装配器（目标 ≤ 350 行）
│   ├── constant.ts                     ← 常量集中（新增 DOC_* / COMMAND_*）
│   │
│   ├── contract/                       ← 🆕 契约层：纯接口，零实现依赖
│   │   ├── editHost.ts                 ← EditHost 接口 + 命令名常量
│   │   ├── commandBus.ts               ← CommandBus（事件派发 + 状态查询）
│   │   ├── contentStore.ts             ← ContentStore<T> 泛型存档契约
│   │   └── workspace.ts                ← Workspace 生命周期契约 + WorkspaceKind
│   │
│   ├── shell/                          ← 🆕 外壳层：从现有目录迁入
│   │   ├── toolbar/
│   │   │   ├── toolbarShell.ts         ← 从 toolbar.ts 拆出的外壳（tabs/折叠/量测）
│   │   │   ├── toolbarRegistry.ts      ← 🆕 分区注册表（按宿主类型装配）
│   │   │   ├── quickAccess.ts          ← 迁入
│   │   │   └── sections/               ← 🆕 通用分区组件（外壳与命令分离）
│   │   │       ├── fontSection.ts      ← 由 sectionFont.ts 拆出（UI 外壳）
│   │   │       ├── alignSection.ts     ← 由 sectionAlign.ts 拆出
│   │   │       ├── brushSection.ts     ← 由 sectionBrush.ts 拆出
│   │   │       └── ...
│   │   ├── fileTree/                   ← 由 fileTree.ts 拆分
│   │   │   ├── fileTree.ts             ← 树 UI（目标 ≤ 900 行）
│   │   │   ├── fileTreeModel.ts        ← 🆕 树结构与路径计算
│   │   │   ├── fileContentBridge.ts    ← 🆕 内容桥（多载体）
│   │   │   └── saveAsDialog.ts         ← 由 fileTree.ts 拆出
│   │   └── ...
│   │
│   ├── sheet/                          ← 表格实现层（由 sheets/ 迁移并改名）
│   │   ├── sheetWorkspace.ts           ← 🆕 表格工作区（实现 Workspace）
│   │   ├── sheetHost.ts                ← 🆕 实现 EditHost（命令 → DataCollection）
│   │   ├── sheetStore.ts               ← 🆕 从 dataCollection 拆出存档职责
│   │   ├── dataCollection.ts           ← 瘦身为纯数据模型（目标 ≤ 3500 行）
│   │   └── canvas/                     ← 原 sheets/ 的画布部分
│   │
│   ├── document/                       ← 🆕 文档实现层
│   │   ├── docWorkspace.ts             ← 文档工作区
│   │   ├── docHost.ts                  ← 实现 EditHost
│   │   ├── docStore.ts                 ← docs 仓读写
│   │   ├── docView.ts                  ← DOM 渲染
│   │   ├── blockTree.ts                ← 块树模型 + 选区
│   │   └── converter/                  ← 三条转换链
│   │       ├── blockTreeCodec.ts       ← 块树 ⇄ 序列化
│   │       ├── markdownCodec.ts        ← 块树 ⇄ Markdown
│   │       ├── textCodec.ts            ← 块树 ⇄ 纯文本
│   │       └── docxCodec.ts            ← 块树 ⇄ docx
│   │
│   └── dataArchitecture/               ← 共享数据基元
│       ├── char.ts                     ← ✅ 保持不变（文档复用）
│       ├── cell.ts                     ← 表格专用
│       ├── cellStyleTable.ts           ← 由 cell.ts 拆出
│       ├── header.ts                   ← 表格专用
│       ├── indexDB.ts                  ← 新增 docs 仓
│       └── ...
│
└── utils/
    ├── parser/
    │   ├── sheetParser.ts              ← 保持
    │   ├── docParser.ts                ← 🆕 docx 解析（走 xmlScanner）
    │   ├── zip.ts                      ← 读
    │   ├── xlsxWriter.ts                ← 保持（含 createZipBlob，docx 复用）
    │   ├── docxWriter.ts               ← 🆕 docx 写出（复用 createZipBlob）
    │   └── xmlScanner.ts               ← 保持
    └── ...
```

### 4.2 依赖方向图（箭头 = import 方向，禁止反向）

```
                    ┌─────────────┐
                    │  core/app   │  装配器：只做 new + 接线
                    └──────┬──────┘
                           │
              ┌────────────┼────────────┐
              ▼            ▼            ▼
        ┌──────────┐ ┌──────────┐ ┌──────────┐
        │  shell   │ │contract  │ │  impl    │
        │ (外壳)   │ │ (契约)   │ │ (实现)   │
        └────┬─────┘ └────┬─────┘ └────┬─────┘
             │            │            │
             └────────────┴────────────┘
                          ▲
                          │ 只允许向上依赖契约
              ┌───────────┴───────────┐
              ▼                       ▼
      ┌───────────────┐      ┌───────────────┐
      │  sheet/*      │      │  document/*   │
      │ (表格实现)    │      │ (文档实现)    │
      └───────────────┘      └───────────────┘
```

**铁律**：`contract/` 目录**不得 import 任何 `sheet/` `document/` `shell/` 的东西**，也不得 import `dataArchitecture/cell.ts`（它含表格语义）。`contract/` 只允许 import `utils/eventEmitter.ts` 与 `core/constant.ts` 中的常量。

### 4.3 契约层三个接口（解耦的支点）

```ts
// core/contract/editHost.ts

/** 编辑宿主：工具栏与对话框通过它操作「当前正在编辑的东西」 */
export interface EditHost {
    readonly kind: 'sheet' | 'doc';
    /** 该命令在当前宿主下是否可执行（工具栏据此置灰） */
    canExecute(cmd: CommandId): boolean;
    /** 执行命令 */
    execute(cmd: CommandId, payload?: CommandPayload): void;
    /** 查询命令状态（工具栏据此高亮按钮） */
    queryState(cmd: CommandId): unknown;
    /** 宿主数据变更事件（工具栏据此刷新外观） */
    readonly events: EventEmitter;
}

export type CommandId =
    // 字符级
    | 'toggleBold' | 'toggleItalic' | 'toggleUnderline' | 'toggleStrike'
    | 'setFontFamily' | 'setFontSize' | 'setFontColor' | 'setHighlight'
    // 段落级
    | 'setAlign' | 'setIndent' | 'setLineHeight' | 'setBlockType'
    | 'toggleList' | 'toggleQuote' | 'clearFormat'
    // 剪贴板
    | 'copy' | 'cut' | 'paste' | 'pastePlainText' | 'pasteFormatOnly'
    // 表格专属（文档侧 canExecute 返回 false）
    | 'toggleMergeCells' | 'setFilter' | 'sortUp' | 'sortDown'
    | 'setBorder' | 'setNumberFormat'
    // 视图
    | 'setZoom' | 'toggleViewMode' | 'find' | 'replace';

export interface CommandPayload {
    value?: unknown;
    toggle?: boolean;
}
```

```ts
// core/contract/commandBus.ts

/**
 * 命令总线：工具栏分区只认识它，不认识任何具体宿主
 */
export class CommandBus extends EventEmitter {
    private host: EditHost | null = null;

    /** 切换宿主（工作区切换时调用） */
    bind(host: EditHost): void {
        this.host = host;
        this.emit('host:changed', host);
    }

    dispatch(cmd: CommandId, payload?: CommandPayload): void {
        if (!this.host || !this.host.canExecute(cmd)) return;
        this.host.execute(cmd, payload);
    }

    state(cmd: CommandId): unknown {
        return this.host?.queryState(cmd);
    }

    get current(): EditHost | null { return this.host; }
}
```

```ts
// core/contract/contentStore.ts

/**
 * 内容存档契约：表格与文档各自实现，文件树只认这个接口
 */
export interface ContentStore<TContent> {
    /** 该载体使用的 IndexedDB 仓名 */
    readonly storeName: string;
    /** 存档键（路径的纯函数） */
    keyOf(path: string): string;
    /** 读取 */
    get(path: string): Promise<TContent | null>;
    /** 写入 */
    set(path: string, content: TContent): Promise<void>;
    /** 删除 */
    remove(path: string): Promise<void>;
    /** 列出全部存档键 */
    keys(): Promise<string[]>;
}
```

### 4.4 改造前后的依赖面对比

| 指标 | 改造前 | 改造后 |
| --- | ---: | ---: |
| `dataCollection.ts` 行数 | 8686 | ~3500（存档职责移出） |
| `app.ts` 行数 | 1152 | ~350 |
| `fileTree.ts` 行数 | 2961 | ~900（+ 3 个新文件） |
| 工具栏分区的 `data.*` 直连数 | 51 | **0**（全部走 `CommandBus`） |
| `contract/` 对实现的依赖 | — | **0**（单向） |
| 工具栏 UI 组件在两种宿主间的复用率 | 0% | **~90%** |

---

## 5. 六项解耦改造详案

### 5.1 D1：抽出契约层（新增 `core/contract/`）

**内容**：3 个文件，纯接口 + 常量，零实现依赖。

**做法**：
1. 新建 `core/contract/editHost.ts`、`commandBus.ts`、`contentStore.ts`、`workspace.ts`
2. `CommandId` 的命令名清单**从现有工具栏调用点反推**（第一节 1.3 的 51 个调用点已分类完毕）
3. `contract/` 里**不放任何实现**，只放接口与类型

**验收**：`npx tsc --noEmit` 通过；`grep -rn "from.*sheet\|from.*document" src/core/contract/` 返回空。

### 5.2 D2：工具栏分区拆为「UI 外壳 + 命令派发」

这是**复用 toolbar UI 实现文档富文本**的前提，也是解耦收益最大的一项。

**改造前**（`sectionFont.ts`）：

```ts
onClick: (todo) => {
    if (this.data.isEditting) {
        this.data.setSelectedCharsFont({ fontWeight: bool });   // ← 直连
    } else {
        this.data.setSelectedCellsFont({ fontWeight: bool });   // ← 直连
    }
}
```

**改造后**（三段式）：

```ts
// ① core/contract/editHost.ts —— 命令契约
export type CommandId = 'toggleBold' | 'setFontFamily' | /* ... */;

// ② core/toolbar/shell/sections/fontSection.ts —— UI 外壳（不知道宿主是谁）
export class FontSection {
    constructor(options: { bus: CommandBus; menu: Menu; title?: string }) { /* ... */ }
    private toggleBold(): void {
        this.bus.dispatch('toggleBold', { toggle: true });   // 只发命令
    }
    private setFontFamily(family: string): void {
        this.bus.dispatch('setFontFamily', { value: family });
    }
    // 按钮高亮：订阅 bus 的状态变化
    private syncState(): void {
        const bold = this.bus.state('toggleBold');
        this.boldBtn.setActived(!!bold);
    }
}

// ③ core/sheet/sheetHost.ts —— 表格侧实现
export class SheetHost implements EditHost {
    readonly kind = 'sheet';
    constructor(private data: DataCollection) {}
    canExecute(cmd: CommandId): boolean {
        return !TABLE_ONLY.has(cmd) || this.data.getSelectedCells().length > 0;
    }
    execute(cmd: CommandId, payload?: CommandPayload): void {
        switch (cmd) {
            case 'toggleBold': {
                const bool = !this.data.queryState('toggleBold');
                if (this.data.isEditting) this.data.setSelectedCharsFont({ fontWeight: bool });
                else this.data.setSelectedCellsFont({ fontWeight: bool });
                break;
            }
            case 'setFontFamily':
                // 原 sectionFont 的分支逻辑原样搬过来，**行为零变化**
                break;
            // ...
        }
    }
}
```

**关键约束**：
- **搬迁而非重写**。`sectionFont.ts` 现有的每一个分支（编辑态 / 非编辑态、`isDisabledOnEditting`、菜单开合、`setSelectedCellsBorder` 的 7 个 border 分支）都**原样搬进 `SheetHost.execute`**，只把 `this.data.X(...)` 换成 `this.data.X(...)` 的等价调用。
- **UI 外壳不含任何 `data` 引用**。改造后 `grep -c "this\.data\." src/core/toolbar/shell/sections/*.ts` 应为 0。
- 现有测试 `test-button-menu.js`（22/22）必须保持全绿 —— 它测的正是按钮菜单开合链路。

**收益**：`sectionFont` 的 UI 外壳（字体下拉、字号下拉、粗/斜/下划/删除线 5 个按钮、文字色、背景色）**一次实现，两种宿主共用**。文档侧只需写 `DocHost.execute('toggleBold')` 即可获得完全一致的按钮外观与交互。

### 5.3 D3：`dataCollection.ts` 瘦身（8686 → ~3500）

**拆分依据**：按第一节 1.2 的职责面，把**不共享表格语义**的部分移出。

| 移出内容 | 去向 | 理由 |
| --- | --- | --- |
| 存档序列化 | `sheet/sheetStore.ts` | 存档格式（v3/v4、分片 JSON）是**存储契约**，不属于数据模型 |
| 落库节流 | `sheet/sheetStore.ts` | `_scheduleAutoSave` / `_runIdleSave` / `flushPendingSave` |
| 筛选 / 隐藏 | `sheet/sheetFilter.ts` | 筛选有独立的索引体系（`_filterCells` / `_filterIndex`）与热路径约束 |
| 查找 / 替换 | `sheet/sheetSearch.ts` | 纯算法，零表格状态依赖，可独立测试 |
| 图片 | `sheet/sheetImages.ts` | `_loadImages` / `_images` / `_imagesDirty` |

**保留在 `dataCollection.ts`**：单元格模型、索引、合并、选区、撤销编排、脚本写回 —— 这些**共享 `_values` / `_cellIndex` / `_mergedCells` 等核心字段**，拆开反而会引入大量跨模块调用。

**做法**：
1. 先**纯搬迁**：把选定区间的代码整段移到新文件，`this.xxx` → `this.store.xxx` 或通过构造注入
2. **不改任何一行逻辑**（搬迁 PR 单独提交，便于 `git bisect`）
3. 搬迁后跑全量测试，对比基线（当前基线见 `MEMORY-DETAILS.md` §8：5 个预先存在的失败不得当成回归）

**注意**：`dataCollection.ts` 已有 8686 行，任何"顺手优化"都会让 review 失效。**本步只搬不改。**

### 5.4 D4：App 瘦身为装配器 + Workspace 路由

`app.ts` 当前 1152 行、8 类职责。拆分：

| 职责 | 去向 |
| --- | --- |
| 布局装配（split / menu / toolbar / navigator / scroller） | 保留在 `app.ts` |
| 文件树交互 + 内容搬迁 + 存档对账 | `app.ts` 保留调用，具体逻辑下沉到 `shell/fileTree/*` 与 `sheet/sheetStore` |
| 缩放滑块 / 底部状态栏 | `app.ts` 保留 |
| **工作区路由** | 🆕 `core/contract/workspace.ts` 的 `WorkspaceRouter` |

**Workspace 契约**：

```ts
// core/contract/workspace.ts
export type WorkspaceKind = 'sheet' | 'doc';

export interface Workspace {
    readonly kind: WorkspaceKind;
    readonly host: EditHost;
    /** 装载指定路径的内容 */
    open(path: string): Promise<boolean>;
    /** 挂载到主区域 */
    mount(parent: HTMLElement): void;
    /** 从主区域卸载（切到另一种工作区时） */
    unmount(): void;
    /** 当前是否可编辑 */
    readonly editable: boolean;
}
```

**`App` 侧只留四行路由**：

```ts
private async activateFile(path: string): Promise<void> {
    const kind = kindOfPath(path);                    // 按原生后缀判定
    if (this.workspace.kind !== kind) {
        this.workspace.unmount();
        this.workspace = this.createWorkspace(kind);   // sheet → SheetWorkspace / doc → DocWorkspace
        this.workspace.mount(this.splitPrev);
        this.toolbarShell.rebind(this.workspace.host); // 工具栏换装
    }
    await this.workspace.open(path);
    // 文件树高亮 / 记忆 / 脚本切换等后续逻辑不变
}
```

**注意**：现有 `activateFile` 里「先 `switchToSheet` 再 `showActivedFile`」的顺序语义要保留 —— 它保证「重复双击同一文件不产生副作用」（`switchToSheet` 返回 false 表示未切换）。

### 5.5 D5：工具栏装配可重入（修 C5）

`setTabContent()` 改为**先清空再挂载**：

```ts
public setTabContent(tabText: string, sections: ToolbarSectionItem[]): void {
    const group = this.groups.get(this.hashOf(tabText));
    if (!group) return;
    // 🆕 先清空：重复装配（切换工作区）会叠加分区
    group.innerHTML = '';
    this.navs.delete(group.id);
    // ...原有 append 逻辑不变
}
```

**配套**：新增 `ToolbarShell.rebind(host: EditHost)` —— 解绑旧分区的监听、绑定新宿主的命令总线、清空并重新装配各页签分区。

**注意**：`group.innerHTML = ''` 会一并清掉 `section-nav` 标签（它是 `group.insertBefore(nav, group.firstChild)` 插进去的），原逻辑在后面会重新插入，顺序正确。但**必须确认 `nav` 的重建逻辑在清空之后执行** —— 现有代码正是这个顺序（先 append sections，再 insertBefore nav），所以只需在函数开头加清空即可。

### 5.6 D6：内容桥泛化（修 C4）

`FileContentBridge` 从"只管 `values`/`images`"泛化为"管**多个载体**"：

```ts
export interface ContentBridge {
    /** 按路径后缀分派到具体载体 */
    create?(paths: string[]): void;
    copy?(plans: FileMovePlan[]): void;
    move?(plans: FileMovePlan[]): void;
    remove?(paths: string[]): void;
}
```

**实现侧**：`App.copyFileContents` 内部按路径后缀分派：

```ts
private async copyFileContents(plans: FileMovePlan[]): Promise<void> {
    for (const plan of plans) {
        const key = this.toArchiveKey(plan.from);
        const raw = await this.settingsDb.getValue(key);
        if (raw !== null) await this.settingsDb.setValue(this.toArchiveKey(plan.to), raw);
        // 🆕 文档载体
        const doc = await this.settingsDb.getDoc(key);
        if (doc !== null) await this.settingsDb.setDoc(this.toArchiveKey(plan.to), doc);
    }
}
```

**关键**：`toArchiveKey` 对 `.grid` 去后缀、对 `.text` 保留后缀，因此 `报表.grid` → `报表`、`报表.text` → `报表.text`，**两个载体天然分键**，共用同一个 `copyFileContents` 循环即可，不需要分派逻辑。**这是零成本的**。

---

## 6. 解耦改造的落地顺序与风险

### 6.1 顺序（严格串行，每步独立可回滚）

| 步 | 内容 | 依赖 | 风险 |
| --- | --- | --- | --- |
| **S1** | 新建 `core/contract/`（纯新增，无改动） | — | 🟢 零风险 |
| **S2** | D5：工具栏装配可重入（`setTabContent` 加清空） | — | 🟢 低（现有只调一次，行为不变） |
| **S3** | D2：工具栏分区拆 UI 外壳 + `CommandBus` + `SheetHost` | S1 | 🟡 中（51 个调用点搬迁，须逐条比对） |
| **S4** | D6：内容桥泛化 + `indexDB` 加 docs 仓 | S1 | 🟡 中（存量存档零迁移，须验全量文件树操作） |
| **S5** | D3：`dataCollection` 拆分（纯搬迁） | S3、S4 | 🟠 中高（8686 行大手术，须逐 PR） |
| **S6** | D4：App 瘦身 + `WorkspaceRouter` 接入 | S3、S5 | 🟠 中高（触及主链路） |

**S3 是整个改造的关键节点**：做完它，工具栏 UI 就已经与表格解耦，文档功能可以并行开工，不必等 S5/S6。

### 6.2 每步的验证口径

| 步 | 验证 |
| --- | --- |
| S1 | `npx tsc --noEmit`；`grep -rn "from.*\(sheet\|document\|shell\)" src/core/contract/` 为空 |
| S2 | `test-toolbar-compact.js` 全绿；现有 `test-button-menu.js` 22/22 |
| S3 | **全量 `node test/test-*.js`**，对比基线（5 个预先存在的失败不得当成回归）；**新增守卫**：`sections/*.ts` 中 `this.data.` 出现次数为 0 |
| S4 | `test-filetree.js` 全绿；**新增守卫**：移动 `.text` 文件后新键有内容、旧键为空 |
| S5 | 全量测试；**新增守卫**：`dataCollection.ts` 行数 < 4000（防回胖） |
| S6 | 全量测试；**新增守卫**：`.grid` 与 `.text` 双击分别进入各自工作区，互不串味 |

### 6.3 风险

| # | 风险 | 缓解 |
| --- | --- | --- |
| 1 | **S3 搬迁时改变行为** | 「搬迁而非重写」；每个命令分支逐条与原代码 diff 比对；`SheetHost.execute` 的 switch 分支必须与原 `sectionXxx` 的分支**一一对应** |
| 2 | **S5 大手术引入隐性回归** | 纯搬迁，单 PR 单职责；每 PR 跑全量测试；不夹带任何"顺手优化" |
| 3 | **契约层反向依赖** | 加一条**源码级守卫测试**：正则扫 `src/core/contract/*.ts` 不得出现 `import ... from "../sheet"` 等 |
| 4 | **既有隐患被放大** | `setTabContent` 叠加（C5）在 S2 先修，否则 S3 的工作区切换会立刻暴露 |
| 5 | **护栏缺失** | 每步补源码级守卫 + **反证一次**（故意改坏，确认守卫 FAIL）—— 这是项目既有铁律 |

---

# 第二部分：Word 文档功能实现

> 本部分按「复用 toolbar UI / 支持 Markdown / 新增 docs 仓并复用 chars / 三条转换链 / DOM 渲染」五项要求设计，
> 并直接建立在第一部分解耦后的架构之上。

## 7. 六项设计决策

| # | 决策 | 结论 | 依据 |
| --- | --- | --- | --- |
| **W1** | toolbar UI 复用方式 | **UI 外壳共享 + CommandBus 派发**（S3 已就位） | 第一节 1.3：51 个 `data.*` 调用点里 39 个是通用编辑语义 |
| **W2** | Markdown 定位 | **交换格式 + 可编辑源码视图**，非运行时真源 | Markdown 表达不了字号/颜色/表格合并 |
| **W3** | 持久层 | **新增 `docs` 仓**，`Char` 直接复用为 run 样式载体 | `Char` 的 7 个字段与 `FontStyle` 前 7 个**完全一致** |
| **W4** | 转换链 | **以块树为中枢**：`docx ⇄ 块树 ⇄ text`、`docx ⇄ 块树 ⇄ Markdown`、`text ⇄ 块树 ⇄ Markdown` | 中枢式比 N×N 直连少一半转换器 |
| **W5** | 渲染 | **DOM，不走 Canvas** | 行内混排/选区/IME/命中测试四项 Canvas 全不占优 |
| **W6** | docx 导出 | **复用 `xlsxWriter.ts` 的 `createZipBlob`** | 修正前版判断：项目**已有**通用 ZIP 写入器 |

### 7.1 W6 的修正说明

上一版方案称"项目没有 zip 写入器，docx 导出硬阻塞" —— **这是错的**。`src/utils/parser/xlsxWriter.ts` 第 555 行已导出通用的：

```ts
export function createZipBlob(entries: ZipEntryInput[]): Blob
```

它实现了完整的 ZIP 写入（local header + central directory + EOCD + CRC32），且**条目名与内容是参数化的**（`ZipEntryInput = { name, data }`），与 xlsx 无关。当前 `fileTree.ts:1776` 已在用它打包下载。

**结论**：docx 导出只需新增 `docxWriter.ts`（负责生成 `word/document.xml` 等条目内容），**打包直接调 `createZipBlob`**。这砍掉了原方案里一个整块工作量。

### 7.2 W3 的 chars 复用（本方案的关键设计）

用户要求"持久层新增 docs 库，能复用 chars"。核对后确认这不是勉强的复用，而是**天然契合**：

| `Char` 字段 | `FontStyle` 字段 | 语义 |
| --- | --- | --- |
| `fontFamily` | `fontFamily` | 字体族 |
| `fontSize` | `fontSize` | 字号 |
| `fontWeight` | `fontWeight` | 加粗 |
| `fontStyle` | `fontStyle` | 斜体 |
| `fontColor` | `fontColor` | 字体颜色 |
| `underline` | `underline` | 下划线 |
| `strikethrough` | `strikethrough` | 删除线 |

**七字段完全一致**。`Char` 还额外满足三个复用前提：

1. **纯数据类，非 EventEmitter** → 无监听表开销，文档块多时内存友好
2. **`fromRaw` 稀疏复制** → 未赋值的字段不成为自有属性，稀疏性天然成立
3. **已有 `toJSON` / `fromJSON`** → 存档序列化直接复用

**设计**：文档的行内内容**直接用 `Char[]` 承载**，不另造 `Run` 类：

```ts
export interface DocBlock {
    id: string;
    type: BlockType;
    chars: Char[];        // ← 复用！与 Cell.chars 同一形态
    align?: 'left' | 'center' | 'right' | 'justify';
    level?: number;
    listStyle?: 'bullet' | 'ordered' | 'none';
    indent?: number;
    children?: DocBlock[];
    imageRef?: { assetId: string; width: number; height: number };
}
```

**这带来三个额外好处**：

| 好处 | 说明 |
| --- | --- |
| **样式应用逻辑可复用** | `Cell` 的 `setSelectedCharsFont` 遍历 `chars` 改样式；文档侧可写一份**同构**的 `applyCharStyle(chars, style)` |
| **Markdown 序列化可复用** | 遍历 `chars` 判断样式组合 → 输出 `**粗**` / `*斜*` / `~~删~~`，与 `Cell.getText()` 的拼接方式一致 |
| **docx 解析可复用** | `<w:r>` → 一个 `Char`，`<w:rPr>` → `Char` 的 7 个字段，与 `biff.ts` 解析 `<w:rPr>` 的手法同构 |

**扩展字段的处理**：`Char` 只有 7 个字段，而文档还需要 `highlight` / `vertAlign` / `link` / `fontSize`（有）。三种选择：

| 方案 | 说明 | 评价 |
| --- | --- | --- |
| **扩 `Char`** | 加 3 个可选字段 | ⚠️ `Char.toJSON` 会多写 3 个键 → **表格存档 v4 的字节内容变化** → 触发快照比对（`_sameCellSnapshot`）的字段级差异 → 存量快照需重算。**不推荐** |
| **文档侧扩展映射** | `DocBlock` 上挂 `links: Map<number, string>`（字符区间 → URL） | ✅ `Char` **零改动**，表格存档完全不受影响 |
| **块级旁挂** | `DocBlock.meta` 存块级特殊属性 | ✅ 最简单，但行内级的高亮/上下标会丢失 |

**推荐：方案二（文档侧扩展映射）** —— `Char` 保持 100% 不动，`DocBlock` 额外持有：

```ts
export interface DocBlock {
    // ...
    chars: Char[];
    /** 字符区间 → 扩展属性（链接 / 高亮 / 上下标），避开 Char 的字段边界 */
    spans?: DocSpan[];
}

export interface DocSpan {
    start: number;      // 起始字符下标（含）
    end: number;        // 结束字符下标（不含）
    link?: string;
    highlight?: string;
    vertAlign?: 'superscript' | 'subscript';
}
```

> **`Char` 保持零改动是硬约束** —— 表格存档 v4 的兼容性与快照比对逻辑（`_sameCellSnapshot` / `_valuesSnapshotsEqual`）都建立在 `Char` 的字段集合与 `toJSON` 输出形状上，动它就是动存量数据。

---

## 8. 块树模型与 chars 复用

### 8.1 模型定义

```ts
// core/document/blockTree.ts
export type BlockType =
    | 'paragraph' | 'heading' | 'listItem' | 'quote'
    | 'code' | 'table' | 'image' | 'hr';

export interface DocBlock {
    id: string;
    type: BlockType;
    /** 行内内容：复用表格的 Char（7 个样式字段） */
    chars: Char[];
    /** 扩展属性区间（链接 / 高亮 / 上下标） */
    spans?: DocSpan[];
    align?: 'left' | 'center' | 'right' | 'justify';
    /** heading 的 1–6 级；listItem 的列表层级 0–4 */
    level?: number;
    listStyle?: 'bullet' | 'ordered' | 'none';
    indent?: number;
    /** 代码块 / 表格单元格的内容 */
    children?: DocBlock[];
    /** 表格专用 */
    rows?: DocBlock[][];
    colSpan?: number;
    rowSpan?: number;
    imageRef?: { assetId: string; width: number; height: number };
}

export interface DocMeta {
    title: string;
    view: 'rich' | 'markdown';
    pageSetup: PrinterSetting;    // 复用 utils/printer 的类型
}

export interface DocSnapshot {
    v: 1;
    meta: DocMeta;
    blocks: DocBlock[];
}
```

### 8.2 chars 复用的三个具体接口

```ts
// core/document/blockTree.ts

/** 取块内纯文本（与 Cell.getText 同构） */
export function blockText(block: DocBlock): string {
    let text = '';
    for (let i = 0; i < block.chars.length; i++) {
        const c = block.chars[i].char;
        if (c != null) text += c;
    }
    return text;
}

/**
 * 对块内字符区间应用样式（与 Cell.setSelectedCharsFont 同构）
 *
 * 复用 `Char` 的稀疏赋值约定：只写实际存在的字段，未赋值的不物化。
 */
export function applyCharStyle(
    chars: Char[], start: number, end: number, style: CharStylePatch
): void {
    for (let i = start; i < end && i < chars.length; i++) {
        const ch = chars[i];
        if (style.fontFamily !== undefined) ch.fontFamily = style.fontFamily;
        if (style.fontSize !== undefined) ch.fontSize = style.fontSize;
        if (style.fontWeight !== undefined) ch.fontWeight = style.fontWeight;
        if (style.fontStyle !== undefined) ch.fontStyle = style.fontStyle;
        if (style.fontColor !== undefined) ch.fontColor = style.fontColor;
        if (style.underline !== undefined) ch.underline = style.underline;
        if (style.strikethrough !== undefined) ch.strikethrough = style.strikethrough;
    }
}

/** 块内某字符区间的样式是否统一（供按钮高亮判断） */
export function queryCharStyle(chars: Char[], start: number, end: number): CharStylePatch {
    // 取区间内首个字符的样式作为代表值（与表格的 queryState 口径一致）
}
```

**注意**：`CharStylePatch` 直接复用 `cellContent.ts` 的 `FontStyle`（前 7 个字段）—— 不新定义类型。

---

## 9. 三条转换链路

### 9.1 中枢式拓扑

以**块树为中枢**，三条链路两两互通：

```
                    ┌──────────┐
                    │  块树     │  DocBlock[]（chars: Char[]）
                    └────┬─────┘
             ┌───────────┼───────────┐
             ▼           ▼           ▼
        ┌─────────┐ ┌─────────┐ ┌─────────┐
        │  docx   │ │  text   │ │Markdown │
        │(二进制) │ │(纯文本) │ │ (源码)  │
        └─────────┘ └─────────┘ └─────────┘
             │           │           │
             └───────────┴───────────┘
                  共 6 条转换边
```

**6 条边**（而非 9 条全直连）：

| 边 | 转换器 | 方向 | 复杂度 |
| --- | --- | --- | --- |
| docx → 块树 | `docParser.ts`（Worker 内） | 单向 | 🟠 高（XML 解析） |
| 块树 → docx | `docxWriter.ts` | 单向 | 🟠 中（XML 生成 + `createZipBlob`） |
| 块树 → text | `textCodec.ts` | 单向 | 🟢 低 |
| text → 块树 | `textCodec.ts` | 单向 | 🟢 低 |
| 块树 → Markdown | `markdownCodec.ts` | 单向 | 🟡 中 |
| Markdown → 块树 | `markdownCodec.ts` | 单向 | 🟡 中 |

**text 与 Markdown 之间的转换不需要直接实现** —— 走 `text → 块树 → Markdown` 即可。这正是"以块树为中枢"的收益：**6 条边而非 9 条，且 text↔Markdown 的往返不损失块结构**。

### 9.2 各转换器的实现要点

#### `textCodec.ts`（块树 ⇄ 纯文本）

| 方向 | 规则 |
| --- | --- |
| 块树 → text | 逐块取 `blockText()`，块间以 `\n` 分隔；表格块以制表符分列、行以 `\n` 分隔（**对齐 TSV**，与 `ClipboardManager.cellsToTSV` 同款） |
| text → 块树 | 按 `\n` 切分 → 每行一个 `paragraph` 块；行内 `\t` → 制表符（`Char.char = '\t'`） |

**复用点**：TSV 的约定与 `clipboardManager.ts` 的 `cellsToTSV` 一致 → 文档与表格之间的复制粘贴可以互通。

#### `markdownCodec.ts`（块树 ⇄ Markdown）

**块树 → Markdown**：

| 块类型 | 输出 |
| --- | --- |
| `heading` level N | `#` × N + 空格 + 文本 |
| `paragraph` | 文本 |
| `listItem` bullet | `- 文本`（按 `indent` 缩进 2 空格/级） |
| `listItem` ordered | `1. 文本` |
| `quote` | `> 文本` |
| `code` | ` ```lang ` 围栏 |
| `hr` | `---` |
| `table` | GFM pipe table |

**行内序列化**（遍历 `chars`，合并连续同样式的连续段）：

| 样式组合 | 输出 |
| --- | --- |
| `fontWeight` | `**文本**` |
| `fontStyle` | `*文本*` |
| `strikethrough` | `~~文本~~` |
| `underline` | `<u>文本</u>` |
| `fontColor` | `文本{#FF0000}`（扩展语法） |
| `highlight` | `==文本==`（扩展语法） |
| 链接（span） | `[文本](url)` |

**Markdown → 块树**：解析器自研（项目 `dependencies` 只有 `pyodide`，保持不引第三方库）。需支持：ATX 标题、Setext 标题、段落、有序/无序/任务列表、引用、围栏代码块、水平线、GFM pipe table、行内 `**` `*` `~~` `` ` `` `[](url)` `![]()` + 本方案的扩展语法。

**扩展语法的降级**：导出的 `{#FF0000}` 在标准 Markdown 阅读器里原样显示（不报错、不丢内容）→ 满足"外部工具可读"。

#### `docxCodec`（docx ⇄ 块树）

**docx → 块树**（`utils/parser/docParser.ts`，走 `xmlScanner` + Worker）：

| docx 要素 | → 块树 |
| --- | --- |
| `w:p` | `DocBlock`（type 由样式决定） |
| `w:pPr/w:pStyle@w:val` | 经 `styles.xml` 查 `w:style[@w:styleId]` → `w:name` → `heading` / `quote` / `code` |
| `w:pPr/w:jc@w:val` | `align` |
| `w:pPr/w:numPr` | `listStyle` + `level`（经 `numbering.xml` 判 bullet/ordered） |
| `w:r` | **`Char`**（`w:t` → `char`） |
| `w:rPr/w:b` `w:i` `w:u` `w:strike` | → `Char.fontWeight` / `fontStyle` / `underline` / `strikethrough` |
| `w:rPr/w:color@w:val` | → `Char.fontColor` |
| `w:rPr/w:sz@w:val` | → `Char.fontSize`（半磅 ÷2） |
| `w:rPr/w:rFonts@w:ascii` | → `Char.fontFamily` |
| `w:hyperlink@r:id` | → `DocSpan.link`（经 rels 解析 URL） |
| `w:tbl` / `w:tr` / `w:tc` | `DocBlock` type=`table`，`w:gridSpan` → `colSpan`，`w:vMerge` → `rowSpan` |
| `w:drawing` → `a:blip@r:embed` | `imageRef`（经 rels 找 `word/media/*`） |
| `w:sectPr` | `meta.pageSetup` |

**复用点**：`biff.ts` 解析 `<w:rPr>` 的手法可直接迁移 —— 都是「标签 → 布尔属性」的映射。

**docx 块树 → docx**（`utils/parser/docxWriter.ts`）：

产出条目交给 `createZipBlob`：

```
[Content_Types].xml
_rels/.rels
word/document.xml          ← 由块树生成
word/styles.xml            ← 由用到的样式生成
word/_rels/document.xml.rels   ← 有图片/超链接时
word/media/imageN.png           ← 有图片时
```

```ts
// utils/parser/docxWriter.ts
import { createZipBlob, type ZipEntryInput } from './xlsxWriter';

export function writeDocx(snapshot: DocSnapshot): Blob {
    const entries: ZipEntryInput[] = [
        { name: '[Content_Types].xml', data: enc(contentTypesXml()) },
        { name: '_rels/.rels', data: enc(relsXml()) },
        { name: 'word/document.xml', data: enc(documentXml(snapshot)) },
        { name: 'word/styles.xml', data: enc(stylesXml(snapshot)) },
    ];
    // 有图片/超链接时追加 rels 与 media
    return createZipBlob(entries);   // ← 复用 xlsxWriter 的通用 ZIP 写入器
}
```

---

## 10. 工具栏复用：UI 外壳共享

### 10.1 复用清单（直接回答"复用原有 toolbar 的 UI"）

| 现有 UI 组件 | 表格侧 | 文档侧 | 复用方式 |
| --- | --- | --- | --- |
| 字体下拉 `Select` | ✅ | ✅ | **同一组件实例工厂**，数据源换成 `bus.state('fontFamily')` |
| 字号下拉 `Select` | ✅ | ✅ | 同上 |
| 粗/斜/下划/删除线 `BtnBase` ×4 | ✅ | ✅ | **同一 UI 外壳**，点击改为 `bus.dispatch('toggleBold')` |
| 文字色 `BtnBase` + `ColorPicker` | ✅ | ✅ | 同上 |
| 背景色 `BtnBase` + `ColorPicker` | ✅ | ✅ | 同上 |
| 对齐按钮 ×4 | ✅ | ✅ | 同上（`toggleMergeCells` 在文档侧 `canExecute` 返回 false → 置灰） |
| 缩进 ± | ✅ | ✅ | 同上 |
| 自动换行 | ✅ | ❌ | 文档用 CSS `white-space`，不设"换行"开关 → `canExecute` false |
| 边框 `BtnBase` + 菜单 | ✅ | ❌ | 文档无单元格边框 → `canExecute` false |
| 合并单元格 | ✅ | ❌ | 同上 |
| 筛选 / 排序 | ✅ | ❌ | 同上 |
| 格式刷 `SectionBrush` | ✅ | ✅ | **交互模式复用**，刷的内容从"单元格样式"变成"字符样式集" |
| 打印 / 打印预览 | ✅ | ✅ | 复用 `QuickAccess` 按钮；`Printer` 换文档实现 |

**结论：约 90% 的 toolbar UI 可直接复用。**

### 10.2 文档工作区的工具栏装配

```ts
// core/toolbar/shell/toolbarRegistry.ts
export function buildDocSections(bus: CommandBus, menu: Menu): Map<string, ToolbarSectionItem[]> {
    return new Map([
        ['开始', [
            { el: new FontSection({ bus, menu, title: '字体样式' }).getElement(),
              title: '字体样式' },
            { el: new ParagraphSection({ bus, menu, title: '段落' }).getElement(),
              title: '段落' },
            { el: new BrushSection({ bus, menu, title: '格式刷' }).getElement(),
              title: '格式刷' },
        ]],
        ['插入', [ /* 表格、图片、链接、代码块、水平线 */ ]],
        ['视图', [ /* 富文本⇄Markdown 切换、缩放、专注模式 */ ]],
        ['页面', [ /* 纸张、页边距、方向（复用 PrinterSetting） */ ]],
    ]);
}
```

**`ToolbarShell.rebind(host)`** 在工作区切换时：
1. 解绑旧分区的 `bus` 监听
2. `bus.bind(host)`
3. 清空并重新装配各页签分区（D5 已把 `setTabContent` 改为可重入）

### 10.3 按钮状态同步

现有 `sectionAlign.ts` 用 `this.data.on(DataEvents.ACTIVED_CELL_CHANGED, () => this.setToolbarsState())` 刷新按钮状态。文档侧对应：

```ts
// core/document/docHost.ts
export class DocHost implements EditHost {
    readonly kind = 'doc';
    readonly events = new EventEmitter();

    queryState(cmd: CommandId): unknown {
        switch (cmd) {
            case 'toggleBold': return this.store.queryBold();
            case 'toggleItalic': return this.store.queryItalic();
            // ...
        }
    }
}
```

`FontSection` 订阅 `bus` 的 `state:changed` 事件刷新高亮 —— **这是 D2 改造带来的直接收益**：现有分区需要各自 `this.data.on(...)`，改造后统一由 `CommandBus` 广播。

---

## 11. 持久层：docs 仓

### 11.1 改动清单

| 位置 | 改动 |
| --- | --- |
| `indexDB.ts` `STORE_NAMES` | 🆕 `DOCS: 'docs'` |
| `indexDB.ts` `DB_VERSION` | 2 → **3** |
| `indexDB.ts` `OBJECT_STORE_CONFIG` | 🆕 `{ [STORE_NAMES.DOCS]: { keyPath: 'docKey', autoIncrement: false } }` |
| `indexDB.ts` `OBJECT_STORE_INDEXES` | 🆕 `{ name: 'docKey', unique: true }` |
| `indexDB.ts` | 🆕 `setDoc / getDoc / deleteDoc / getAllDocs` |
| `indexDB.ts` `createObjectStores` | 已是幂等（`if (!contains)`）→ **存量零迁移** |

### 11.2 存档结构

```ts
export interface DocSnapshot {
    v: 1;
    meta: DocMeta;
    blocks: DocBlock[];
}
```

**为什么结构化克隆直存、而不是像 `values` 那样分片 JSON**：

| 维度 | 表格 `values` | 文档 `docs` |
| --- | --- | --- |
| 体量 | 10⁵～10⁶ 单元格 | 10²～10⁴ 块 |
| 序列化 | 分片 JSON（4000 格/片）+ `SparseAxis` | **结构化克隆直存**（`Char` 实例可直接存） |
| 理由 | 分片是为避免百万级字符串拼接造成主线程长任务 | 文档规模下引入分片只增复杂度 |

> ⚠️ **结构化克隆会丢失原型**：`Char` 实例存进 IndexedDB 后读回来是**普通对象**（没有 `Char.prototype`）。因此 `getDoc` 必须用 `Char.fromRaw` 重建原型 —— 这与 `parserClient.ts` 里"generate 产出 Cell/Char 类实例，必须在主线程完成（跨线程会丢原型）"是同一个坑，**docs 仓同理**。

```ts
// indexDB.ts
public async getDoc(docKey: string): Promise<DocSnapshot | null> {
    const record = await this.wrapRequest<DocRecord | undefined>(
        this.getObjectStore(this.getTransaction('readonly', STORE_NAMES.DOCS), STORE_NAMES.DOCS).get(docKey)
    );
    if (!record || !record.snapshot) return null;
    // 🆕 重建 Char 原型（结构化克隆会丢原型）
    const snapshot = record.snapshot as DocSnapshot;
    snapshot.blocks = snapshot.blocks.map(restoreBlock);
    return snapshot;
}

function restoreBlock(block: DocBlock): DocBlock {
    if (Array.isArray(block.chars)) {
        block.chars = block.chars.map(c => Char.fromRaw(c));
    }
    if (Array.isArray(block.children)) block.children = block.children.map(restoreBlock);
    return block;
}
```

### 11.3 存档键规则

**沿用 `toArchiveKey(path)`，无需改动**：

| 路径 | 存档键 | values 仓 | docs 仓 |
| --- | --- | --- | --- |
| `报表.grid` | `报表` | ✅ | — |
| `报表.text` | `报表.text` | — | ✅ |
| `报表.md` | `报表.md` | — | ✅ |
| `归档/销售.grid` | `归档/销售` | ✅ | — |
| `归档/说明.text` | `归档/说明.text` | — | ✅ |

**同名不同后缀 = 两个独立文件，各自独立存档** —— 这是项目既有铁律，docs 仓天然继承。

### 11.4 文件树与 docs 仓的联动

`App.copyFileContents` / `moveFileContents` / `removeFileContents` / `ensureFileContents` 四个方法各加一段 docs 处理（详见第一部分 D6）。

**`reconcileValueStore` 的扩展**：该方法当前只扫描 `values` 仓做孤儿清理与历史键迁移。需扩展为同时扫描 `docs` 仓：

- **只删空存档**（与表格侧同口径）
- 非空存档要**迁移**（历史键 → 新键），不能删

---

## 12. 渲染层：DOM

### 12.1 为什么不走 Canvas（回答"渲染用 DOM"的技术依据）

| 维度 | 表格（Canvas 的理由成立） | 文档（DOM 更合适） |
| --- | --- | --- |
| 节点规模 | 10⁵～10⁶ 单元格 | 10²～10⁴ 块/run |
| 排版复杂度 | 每格单行或整体换行 | **行内混排**、中西文混排、行内样式、混排基线对齐 |
| 选区 | 矩形选区，规则简单 | 任意字符区间、跨块选区、DOM Range |
| 输入法 | 单元格内单行编辑 | 长文本 IME 组字、候选窗定位 |
| 命中测试 | 坐标算行列 | `caretRangeFromPoint` 免费给 |

### 12.2 结构

```
DocView（core/document/docView.ts）
├── 页面容器 .doc-page（纸张宽度 + 页边距）
├── 块容器 .doc-blocks
│   └── 每个块一个 <div data-block-id="...">
│       └── 行内内容按 chars 的样式分段渲染成 <span>
├── 选区层（自绘选区矩形，跨块时每块各一段）
├── 光标层（自绘 caret，闪烁）
└── 滚动/缩放（复用 Scroller 的滚动模型，缩放用 CSS transform）
```

### 12.3 编辑循环

| 环节 | 实现 |
| --- | --- |
| 光标落点 | `document.caretRangeFromPoint()` → 由 `data-block-id` + DOM 路径反查 `(blockIndex, charIndex)` |
| 选区 | 浏览器 `Range` + 自建逻辑选区（跨块时首尾各持一段） |
| 输入 | `beforeinput` 拦截 → 改块树 → 增量重绘 → `Selection.setBaseAndExtent` 复位光标 |
| IME | `compositionstart/update/end` 三态：组字期间**只更新插入符预览不改模型**；`compositionend` 落一次变更 |
| 撤销 | 复用 `UndoManager` 类，**独立实例**，新增 `type: 'docBlocks'` 快照 |

**受控渲染原则**：DOM 只由模型单向生成，**DOM 不反写模型**。这规避了 contenteditable 的浏览器归一化问题（浏览器可能插入 `<b>`、合并文本节点、删空格）。

### 12.4 Markdown 源码视图

复用 `pyEditor.ts` 的 **`pre`(高亮层) + `textarea`(输入层) 叠层 + gutter 同步** 范式：

```ts
// core/document/markdownView.ts
export class MarkdownView {
    private highlight: HTMLPreElement;   // 语法高亮层
    private textarea: HTMLTextAreaElement; // 输入层
    private gutter: HTMLPreElement;        // 行号

    // pyEditor.ts 的 highlightPython → highlightMarkdown
    // 双向：源码改动 → 解析 → 块树 → 反向刷新富文本视图
}
```

高亮器只需一个轻量 tokenizer（标题/粗斜/代码/链接/列表/引用）—— 比 `highlightPython` 的正则简单得多。

---

## 13. 实施路线图

> 前置：先完成第一部分的 **S1–S3**（契约层 + 工具栏解耦），再开工文档功能。

### P0 — 解耦地基（不产出用户可见功能）

| # | 任务 | 涉及 | 验收 |
| --- | --- | --- | --- |
| 1 | `core/contract/` 四个契约文件 | 新建 | `tsc` 通过；契约层零实现依赖（源码级守卫） |
| 2 | `setTabContent` 改为可重入 | `toolbar.ts` | `test-toolbar-compact.js` 全绿 |
| 3 | 工具栏分区拆 UI 外壳 + `CommandBus` + `SheetHost` | `toolbar/`、`sheet/` | 51 个 `data.*` 直连降到 0；**表格行为零变化**（全量测试对比基线） |

### P1 — 文档骨架（只读）

| # | 任务 | 验收 |
| --- | --- | --- |
| 4 | `docs` 仓 + 4 个存取方法（`getDoc` 重建 `Char` 原型） | DB_VERSION 升 3；**存量表格存档全部可读** |
| 5 | 块树模型 + `blockText` / `applyCharStyle` / `queryCharStyle` | 复用 `Char` 的 7 字段；`Char` **零改动** |
| 6 | `DocStore`（装载 / 脏标记 / 防抖落库 / 切换） | 切换文档不丢表格内容 |
| 7 | `DocView` 只读 DOM 渲染 | 段落/标题/列表/引用/代码块/表格正确渲染 |
| 8 | `Workspace` 契约 + `WorkspaceRouter` + `App` 四处路由点改造 | `.grid` / `.text` 双击分别进各自工作区 |
| 9 | `DocHost` 实现 `EditHost` | 文档侧 `canExecute` 对表格专属命令返回 false |

### P2 — 编辑 + Markdown

| # | 任务 | 验收 |
| --- | --- | --- |
| 10 | 编辑层：光标 / 选区 / `beforeinput` / IME 三态 | 中文输入法组字不丢字、不跳字 |
| 11 | 文档工具栏分区（字体/段落/样式/插入/视图） | 按钮生效，高亮反映选区状态 |
| 12 | 文档撤销/重做（`UndoManager` 独立实例 + `docBlocks`） | Ctrl+Z / Ctrl+Y 正确回退块树与选区 |
| 13 | `textCodec`（块树 ⇄ 纯文本） | 与 `clipboardManager.cellsToTSV` 的 TSV 约定一致 |
| 14 | `markdownCodec`（块树 ⇄ Markdown，含扩展语法） | 往返无损（行内样式保留率 ≥ 95%） |
| 15 | Markdown 源码视图（复用 `pyEditor` 叠层范式） | 源码可编辑，双向同步 |

### P3 — docx 互操作 + 打印 + Python

| # | 任务 | 验收 |
| --- | --- | --- |
| 16 | `docParser`（`xmlScanner` + Worker，沿用 decode/extract/generate） | 5 MB docx 打开 < 1.5 s；主线程单任务 < 50 ms |
| 17 | `docxWriter`（块树 → `word/document.xml` → **复用 `createZipBlob`**） | Word / WPS 打开无警告、样式正确 |
| 18 | 「另存为 docx」接入 `fileTree` 清单对话框 | 批量下载可用 |
| 19 | 文档打印与预览（复用 `PrinterSetting`，新写分页） | A4 / 横向 / 页边距预设生效 |
| 20 | Python `doc` 命名空间 + `open_doc()` + 写回 | 脚本能读写文档，落点进撤销栈 |

---

## 14. 风险与已知坑

| # | 风险 | 影响 | 缓解 |
| --- | --- | --- | --- |
| 1 | **S3 工具栏搬迁改变表格行为** | 表格核心交互回归 | 「搬迁而非重写」；逐命令分支 diff 比对；全量测试对比基线 |
| 2 | **受控富文本编辑器的不确定性** | 光标/IME 反复出问题 | 严格"模型为准、DOM 只读渲染"；IME 组字期间不落模型；P1 先跑通只读链路 |
| 3 | **`Char` 原型在 IndexedDB 中丢失** | 文档读回后样式方法不可用 | `getDoc` 用 `Char.fromRaw` 重建；**加守卫测试**：读回的 `chars[0]` 必须是 `Char` 实例 |
| 4 | **`Char` 被误改** | 表格存档 v4 字节变化 → 快照比对失效 → 存量重算 | **`Char` 零改动**是硬约束；扩展属性走 `DocBlock.spans` |
| 5 | **`documentParser.ts` 用 `DOMParser`** | 大 docx OOM 且进不了 Worker | P3 起强制改 `xmlScanner.scanXml` |
| 6 | **`Toolbar.setTabContent` 叠加** | 切工作区后分区显示两份 | D5 先修（重复装配先清 group 子节点） |
| 7 | **扩展 Markdown 语法不被外部工具识别** | 导出的 md 在别处显示异常 | 扩展语法设计成不破坏标准渲染；提供"仅标准 Markdown"导出选项 |
| 8 | **`reconcileValueStore` 只认 values** | 删文档文件后 docs 留孤儿存档 | 扩展为同时扫描 docs 仓；**只删空存档**，非空要迁移 |
| 9 | **`fileTree` 搬运遗漏 docs** | 移动文件后文档内容留在旧键 | D6 泛化内容桥；**守卫测试**：移动 `.text` 后新键有内容、旧键为空 |
| 10 | **`.demo` 无工作区** | 演示文稿行为未定义 | 明确降级为表格模式并提示，不静默 |
| 11 | **契约层反向依赖** | 架构退化，`contract/` 变成垃圾场 | 源码级守卫：正则扫 `contract/*.ts` 不得 import 实现层 |

### 14.1 与既有铁律的相容性自查

| 既有约定（`MEMORY.md`） | 是否冲突 | 说明 |
| --- | --- | --- |
| 存档键 = `toArchiveKey(path)` 纯函数 | ✅ | docs 仓沿用同一规则，不引入第二套换算 |
| 「同名不同后缀 = 两个独立文件」 | ✅ | `.grid` / `.text` 键本就不同 |
| xlsx XML 一律 `xmlScanner`，禁 `DOMParser` | ✅ 主动遵守 | 文档解析同样走 `xmlScanner` |
| 禁全局 `getElementById` 找自身子元素 | ⚠️ | 文档视图找块必须**容器内 scoped** |
| `el.classList` 只能 add/remove/toggle | ⚠️ | `classList = 'x'` 会抛 TypeError |
| `hidden` 显隐须补 `[hidden]{display:none}` | ⚠️ | 工作区切换用 `hidden`，父容器有 `display:flex` 须补规则 |
| 异步逻辑禁包进同步 undo 回调 | ⚠️ | 文档 undo 回调内不得 `await` |
| `Char` 纯数据类、非 EventEmitter、稀疏字段 | ✅ **完全遵守** | 文档直接复用 `Char`，不新建 Run 类 |
| 测试共享同一 `loadModule` 缓存 | ⚠️ | 新模块测试同样共享缓存，否则 `Char` 是另一模块实例 |
| 守卫写完**必须反证一次** | ⚠️ | 契约层、docs 仓、chars 复用、工作区路由都要有守卫 |
| 不写「宽松」DOM 桩 | ⚠️ | 文档视图的 DOM 桩需含 `Range` / `Selection` / `getSelection` |

---

## 15. 附录：接口草案

### 15.1 `core/contract/editHost.ts`

```ts
export type WorkspaceKind = 'sheet' | 'doc';

export type CommandId =
    | 'toggleBold' | 'toggleItalic' | 'toggleUnderline' | 'toggleStrike'
    | 'setFontFamily' | 'setFontSize' | 'setFontColor' | 'setHighlight'
    | 'setAlign' | 'setIndent' | 'setLineHeight' | 'setBlockType'
    | 'toggleList' | 'toggleQuote' | 'clearFormat'
    | 'copy' | 'cut' | 'paste' | 'pastePlainText' | 'pasteFormatOnly'
    | 'toggleMergeCells' | 'setFilter' | 'sortUp' | 'sortDown'
    | 'setBorder' | 'setNumberFormat'
    | 'setZoom' | 'toggleViewMode' | 'find' | 'replace';

export interface CommandPayload {
    value?: unknown;
    toggle?: boolean;
}

export interface EditHost {
    readonly kind: WorkspaceKind;
    canExecute(cmd: CommandId): boolean;
    execute(cmd: CommandId, payload?: CommandPayload): void;
    queryState(cmd: CommandId): unknown;
    readonly events: EventEmitter;
}
```

### 15.2 `core/contract/workspace.ts`

```ts
export interface Workspace {
    readonly kind: WorkspaceKind;
    readonly host: EditHost;
    open(path: string): Promise<boolean>;
    mount(parent: HTMLElement): void;
    unmount(): void;
    readonly editable: boolean;
}

export interface WorkspaceRouter {
    readonly current: Workspace;
    switchTo(kind: WorkspaceKind): Workspace;
    onChange(fn: (kind: WorkspaceKind) => void): void;
}
```

### 15.3 `core/contract/contentStore.ts`

```ts
export interface ContentStore<TContent> {
    readonly storeName: string;
    keyOf(path: string): string;
    get(path: string): Promise<TContent | null>;
    set(path: string, content: TContent): Promise<void>;
    remove(path: string): Promise<void>;
    keys(): Promise<string[]>;
}
```

### 15.4 `core/document/blockTree.ts`

```ts
import { Char } from '../dataArchitecture/char';
import { FontStyle } from '../dataArchitecture/cellContent';

export type BlockType =
    | 'paragraph' | 'heading' | 'listItem' | 'quote'
    | 'code' | 'table' | 'image' | 'hr';

export interface DocSpan {
    start: number;
    end: number;
    link?: string;
    highlight?: string;
    vertAlign?: 'superscript' | 'subscript';
}

export interface DocBlock {
    id: string;
    type: BlockType;
    chars: Char[];
    spans?: DocSpan[];
    align?: 'left' | 'center' | 'right' | 'justify';
    level?: number;
    listStyle?: 'bullet' | 'ordered' | 'none';
    indent?: number;
    children?: DocBlock[];
    rows?: DocBlock[][];
    colSpan?: number;
    rowSpan?: number;
    imageRef?: { assetId: string; width: number; height: number };
}

export interface DocSnapshot {
    v: 1;
    meta: DocMeta;
    blocks: DocBlock[];
}

/** 取块内纯文本（与 Cell.getText 同构） */
export function blockText(block: DocBlock): string;

/** 对块内字符区间应用样式（复用 Char 的稀疏赋值约定） */
export function applyCharStyle(
    chars: Char[], start: number, end: number, style: FontStyle
): void;

/** 块内某字符区间的样式（供按钮高亮判断） */
export function queryCharStyle(chars: Char[], start: number, end: number): FontStyle;
```

### 15.5 `core/document/converter/`

```ts
// markdownCodec.ts
export function blocksToMarkdown(blocks: DocBlock[]): string;
export function markdownToBlocks(md: string): DocBlock[];

// textCodec.ts
export function blocksToText(blocks: DocBlock[]): string;
export function textToBlocks(text: string): DocBlock[];

// docxCodec（解析侧在 utils/parser/docParser.ts）
export function blocksToDocxXml(snapshot: DocSnapshot): string;
```

### 15.6 `utils/parser/docxWriter.ts`

```ts
import { createZipBlob, type ZipEntryInput } from './xlsxWriter';

export function writeDocx(snapshot: DocSnapshot): Blob {
    const entries: ZipEntryInput[] = [ /* ... */ ];
    return createZipBlob(entries);   // 复用 xlsxWriter 的通用 ZIP 写入器
}
```

---

## 16. 一页纸总结

| 议题 | 结论 |
| --- | --- |
| **解耦总纲** | 三层夹心：外壳层（与内容无关）+ 契约层（纯接口）+ 实现层（可插拔） |
| **解耦支点** | `EditHost` + `CommandBus` + `ContentStore<T>` + `Workspace` 四个契约 |
| **解耦收益** | 工具栏 `data.*` 直连 51 → 0；`dataCollection` 8686 → ~3500；`app` 1152 → ~350 |
| **解耦顺序** | S1 契约 → S2 装配可重入 → **S3 工具栏解耦（关键节点）** → S4 内容桥 → S5 大类拆分 → S6 App 瘦身 |
| **toolbar UI 复用** | UI 外壳共享 + `CommandBus` 派发；**约 90% 的 UI 直接复用** |
| **Markdown** | 交换格式 + 可编辑源码视图；行内扩展语法 `{#色}` `==高亮==` 保证往返无损 |
| **docs 仓** | `DB_VERSION` 2→3，幂等升级，存量零迁移；结构化克隆直存，**不分片** |
| **复用 chars** | **直接用 `Char[]` 承载行内内容**（7 字段与 `FontStyle` 完全一致）；`Char` **零改动**，扩展属性走 `DocBlock.spans` |
| **转换链** | 以块树为中枢，**6 条边而非 9 条**；text↔Markdown 走块树中转，不需直连 |
| **docx 导出** | **复用 `xlsxWriter.ts` 的 `createZipBlob`**（修正前版"缺 zip 写入器"的判断） |
| **渲染** | DOM 受控渲染（模型 → DOM 单向），增量 patch + 光标复位；不走 Canvas |
| **最大风险** | ① 受控富文本编辑器的光标/IME 稳定性 ② S3 搬迁改变表格行为 ③ `Char` 被误改导致存档不兼容 |
| **硬约束** | `Char` 零改动；契约层零实现依赖；`setTabContent` 先修可重入 |
