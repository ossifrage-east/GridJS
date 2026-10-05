# SilkSpaces Code Wiki

> 藕丝空间 (Lotus Filament Space) — 基于 HTML5 Canvas 的 Web 电子表格组件  
> 会道者，一缕藕丝牵大象。寡道者，千钧铁棒打苍蝇。

---

## 目录

- [1. 项目概述](#1-项目概述)
- [2. 技术栈与构建工具](#2-技术栈与构建工具)
- [3. 项目目录结构](#3-项目目录结构)
- [4. 整体架构](#4-整体架构)
- [5. 核心模块详解](#5-核心模块详解)
  - [5.1 数据架构层 (dataArchitecture)](#51-数据架构层-dataarchitecture)
  - [5.2 Canvas 渲染层 (sheets)](#52-canvas-渲染层-sheets)
  - [5.3 UI 组件层 (common)](#53-ui-组件层-common)
  - [5.4 工具栏层 (toolbar)](#54-工具栏层-toolbar)
  - [5.5 工具模块 (utils)](#55-工具模块-utils)
  - [5.6 常量与事件 (constant)](#56-常量与事件-constant)
- [6. 类继承关系图](#6-类继承关系图)
- [7. 数据流与事件驱动](#7-数据流与事件驱动)
- [8. 数据持久化 (IndexedDB)](#8-数据持久化-indexeddb)
- [9. 依赖关系](#9-依赖关系)
- [10. 项目运行方式](#10-项目运行方式)

---

## 1. 项目概述

SilkSpaces（藕丝空间）是一个轻量级的 Web 电子表格组件，使用原生 JavaScript Canvas 实现，尽量不依赖第三方库。项目目标是在浏览器中提供类似 Excel 的表格编辑体验，支持单元格编辑、样式设置、字体管理、行列操作等功能，数据通过 IndexedDB 进行本地持久化。

**核心特性：**
- 基于 Canvas 的高性能表格渲染
- 单元格级别的样式和字符级格式控制
- 响应式布局，支持面板拖拽调整
- IndexedDB 本地数据持久化
- 事件驱动的数据变更通知机制
- 支持缩放、选区、编辑等交互

---

## 2. 技术栈与构建工具

| 分类 | 技术 |
|------|------|
| 语言 | TypeScript (ES6 模块, 输出目标 ES5) |
| 构建 | Webpack 5 + ts-loader |
| 样式 | SCSS + PostCSS + Autoprefixer |
| 数据存储 | 浏览器原生 IndexedDB API |
| 渲染 | HTML5 Canvas 2D API |
| 包管理 | npm |
| 开发服务器 | webpack-dev-server |
| CSS 提取 | mini-css-extract-plugin |
| SVG 处理 | mini-svg-data-uri |
| 哈希 | murmurhash-js（工具栏标签 ID 生成） |

---

## 3. 项目目录结构

```
silkspaces/
├── dist/                          # 构建输出目录
├── src/
│   ├── assets/                    # 静态资源
│   │   ├── fonts/                 # HarmonyOS Sans SC 字体文件
│   │   ├── icons/                 # 图标字体 (gridjs)
│   │   └── images/                # SVG 图标资源
│   │       ├── blockly/           # 边框样式图标
│   │       ├── cursor/            # 光标图标
│   │       └── toolbar/           # 工具栏按钮图标
│   ├── core/                      # 核心业务代码
│   │   ├── common/                # 通用 UI 组件
│   │   │   ├── button.ts          # 按钮组件
│   │   │   ├── menu.ts            # 菜单组件
│   │   │   ├── message.ts         # 消息提示组件
│   │   │   ├── navigator.ts       # 导航面板组件
│   │   │   ├── scroller.ts        # 滚动条组件
│   │   │   ├── select.ts          # 选择器组件
│   │   │   └── split.ts           # 面板分割组件
│   │   ├── dataArchitecture/      # 数据架构层
│   │   │   ├── cell.ts            # 单元格类
│   │   │   ├── cellContent.ts     # 单元格内容排版接口
│   │   │   ├── char.ts            # 字符类
│   │   │   ├── dataCollection.ts  # 数据集合管理器
│   │   │   ├── header.ts          # 行列头管理
│   │   │   ├── indexDB.ts         # IndexedDB 封装
│   │   │   ├── values.ts          # 值对象（预留扩展）
│   │   │   └── visibleView.ts     # 可见视图尺寸管理
│   │   ├── sheets/                # Canvas 渲染层
│   │   │   ├── allSelect.ts       # 全选按钮
│   │   │   ├── canvas.ts          # Canvas 基类
│   │   │   ├── colHeaderCanvas.ts # 列头画布
│   │   │   ├── editCanvas.ts      # 编辑画布
│   │   │   ├── rowHeaderCanvas.ts # 行头画布
│   │   │   ├── sheetCanvas.ts     # 主表格画布
│   │   │   └── sheets.ts          # 工作表容器
│   │   ├── style/                 # SCSS 样式文件
│   │   │   ├── brand.scss         # 品牌样式
│   │   │   ├── menu.scss          # 菜单样式
│   │   │   ├── nav.scss           # 导航样式
│   │   │   ├── scroller.scss      # 滚动条样式
│   │   │   ├── sheets.scss        # 工作表样式
│   │   │   ├── split.scss         # 分割面板样式
│   │   │   └── toolbar.scss       # 工具栏样式
│   │   ├── toolbar/               # 工具栏层
│   │   │   ├── quickAccess.ts     # 快速访问栏
│   │   │   ├── toolbar.ts         # 工具栏容器
│   │   │   └── start/             # "开始"标签页分区
│   │   │       ├── sectionAlign.ts   # 对齐方式分区
│   │   │       ├── sectionBrush.ts   # 格式刷分区
│   │   │       ├── sectionFont.ts    # 字体样式分区
│   │   │       └── sectionProcess.ts # 数据处理分区
│   │   ├── app.ts                 # 应用入口类
│   │   └── constant.ts            # 常量与事件定义
│   ├── utils/                     # 工具模块
│   │   ├── clipboardManager.ts    # 剪贴板管理器
│   │   ├── debounce.ts            # 防抖函数
│   │   ├── dom.ts                 # DOM 操作工具
│   │   ├── eventEmitter.ts        # 事件发射器
│   │   ├── fontsLoader.ts         # 字体检测与加载
│   │   └── undoManager.ts         # 撤销/恢复管理器
│   ├── declareModule.d.ts         # 模块类型声明
│   ├── index.html                 # HTML 模板
│   ├── index.scss                 # 全局样式入口
│   └── index.ts                   # 应用启动入口
├── .browserslistrc                # 浏览器兼容配置
├── Introduction.txt               # 项目简介
├── package.json                   # 项目配置
├── postcss.config.js              # PostCSS 配置
├── tsconfig.json                  # TypeScript 配置
└── webpack.config.js              # Webpack 配置
```

---

## 4. 整体架构

项目采用 **分层架构** 设计，从下到上分为四层：

```
┌──────────────────────────────────────────────────┐
│                    App 层                         │
│  应用入口，组装所有模块，监听数据事件触发重绘        │
├──────────────────────────────────────────────────┤
│               UI 组件层 (common)                  │
│  Split / Menu / Navigator / Scroller / Button     │
│  / Select / Message / Toolbar / QuickAccess       │
├──────────────────────────────────────────────────┤
│              Canvas 渲染层 (sheets)               │
│  Sheets → SheetCanvas / RowHeaderCanvas           │
│         / ColHeaderCanvas / EditCanvas / AllSelect │
├──────────────────────────────────────────────────┤
│             数据架构层 (dataArchitecture)          │
│  DataCollection → Cell / Char / Header            │
│                  / VisibleView / IndexDB           │
├──────────────────────────────────────────────────┤
│              工具模块 (utils)                      │
│  EventEmitter / UndoManager / ClipboardManager    │
│  / dom / debounce / fontsLoader                   │
└──────────────────────────────────────────────────┘
```

**核心设计理念：**
- **事件驱动**：所有数据变更通过 EventEmitter 发布事件，UI 层订阅响应
- **数据与渲染分离**：DataCollection 管理数据状态，Canvas 层负责绘制
- **Canvas 高性能渲染**：表格内容完全通过 Canvas 2D 绘制，避免 DOM 节点过多

---

## 5. 核心模块详解

### 5.1 数据架构层 (dataArchitecture)

数据架构层是整个项目的数据核心，管理电子表格的所有数据状态。

#### 5.1.1 DataCollection

**文件**：[dataCollection.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/dataArchitecture/dataCollection.ts)  
**继承**：`EventEmitter`

数据集合管理器，是整个应用的数据中心。协调列头、行头、单元格、可见视图等子模块，提供数据操作、坐标转换和矩形计算等功能。

| 属性 | 类型 | 说明 |
|------|------|------|
| `db` | `IndexDB` | IndexedDB 实例（readonly） |
| `visibleView` | `VisibleView` | 可见视图尺寸实例 |
| `colHeaders` | `ColHeaders` | 列头集合 |
| `rowHeaders` | `RowHeaders` | 行头集合 |
| `values` (getter/setter) | `Cell[]` | 当前工作表所有单元格数据（通过 JSON 序列化/反序列化实现持久化） |
| `activedSheetName` | `string` | 当前选中工作表名称 |
| `isEditting` | `boolean` | 是否正在编辑 |
| `showGridLines` | `boolean` | 是否显示网格线 |
| `activedCell` | `string` | 当前选中单元格（如 "A1"） |
| `selection` | `string` | 选区范围（如 "A1:B3"） |
| `anchorSelection` | `string \| null` | 锚点选区 |
| `zoom` | `number` | 缩放比例（默认 1.2） |

**关键方法：**

| 方法 | 签名 | 说明 |
|------|------|------|
| `syncValuesFromDB` | `() => Promise<void>` | 从 IndexedDB 同步数据到内存 |
| `syncValuesToDB` | `() => Promise<void>` | 将内存数据同步到 IndexedDB |
| `getCellValue` | `(cell: string) => string` | 获取指定单元格的文本值 |
| `setCellValue` | `(cell: string, value: string) => void` | 设置指定单元格的文本值 |
| `addCell` | `(cell: Cell) => void` | 添加单元格到数据集合 |
| `createSheetName` | `() => Promise<void>` | 创建新的工作表 |
| `fromJSON` | `static (json: string) => Cell[]` | 将 JSON 字符串反序列化为 Cell 数组 |

> **重要**：`values` 属性使用 getter/setter 模式。getter 返回 `DataCollection.fromJSON()` 的结果（新数组），setter 通过 `updateProperty` 触发 `VALUES_CHANGED` 事件。直接对 getter 结果执行 `push()` 不会持久化变更，必须通过 setter 赋值。

#### 5.1.2 Cell

**文件**：[cell.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/dataArchitecture/cell.ts)  
**继承**：`EventEmitter`

表示电子表格的单元格，包含样式和值属性。

| 属性 | 类型 | 说明 |
|------|------|------|
| `cell` | `string` | 单元格名称（如 "A1"） |
| `isEdit` | `boolean` | 是否正在编辑 |
| `colspan` | `number?` | 跨列数 |
| `rowspan` | `number?` | 跨行数 |
| `fontFamily` | `string?` | 字体族 |
| `fontSize` | `number?` | 字体大小 |
| `fontWeight` | `boolean?` | 是否加粗 |
| `fontStyle` | `boolean?` | 是否斜体 |
| `fontColor` | `string?` | 字体颜色 |
| `underline` | `boolean?` | 下划线 |
| `strikethrough` | `boolean?` | 删除线 |
| `borderTopWidth` | `number?` | 上边框宽度 |
| `borderBottomWidth` | `number?` | 下边框宽度 |
| `borderLeftWidth` | `number?` | 左边框宽度 |
| `borderRightWidth` | `number?` | 右边框宽度 |
| `textAlign` | `TextAlign?` | 水平对齐（`left`/`center`/`right`） |
| `alignItems` | `VerticalAlign?` | 垂直对齐（`top`/`middle`/`bottom`） |
| `backgroundColor` | `string?` | 背景颜色 |
| `chars` | `Char[]` | 字符数组 |
| `filter` | `boolean?` | 是否过滤 |
| `wrap` | `boolean?` | 自动换行 |
| `letterSpacing` | `number?` | 字间距 |
| `lineSpacing` | `number?` | 行间距 |

所有属性均通过 getter/setter + `updateProperty` 实现变更事件通知。

#### 5.1.3 Char

**文件**：[char.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/dataArchitecture/char.ts)  
**继承**：`EventEmitter`

表示单元格中的单个字符，支持字符级别的样式控制。

| 属性 | 类型 | 说明 |
|------|------|------|
| `char` | `any` | 字符内容 |
| `fontFamily` | `string?` | 字体族 |
| `fontSize` | `number?` | 字体大小 |
| `fontWeight` | `boolean?` | 是否加粗 |
| `fontStyle` | `boolean?` | 是否斜体 |
| `fontColor` | `string?` | 字体颜色 |
| `underline` | `boolean?` | 下划线 |
| `strikethrough` | `boolean?` | 删除线 |

#### 5.1.4 ColHeaders / RowHeaders

**文件**：[header.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/dataArchitecture/header.ts)  
**继承**：`EventEmitter`

管理列头和行头的位置与尺寸信息。

**ColHeaders 关键属性：**

| 属性 | 类型 | 说明 |
|------|------|------|
| `offsetWidth` | `number` | 水平偏移量（滚动位置） |
| `allColWidth` | `number` | 列宽度总和 |

**ColHeaders 关键方法：**

| 方法 | 说明 |
|------|------|
| `getAllCols()` | 获取所有列头数据副本 |
| `getAt(index)` | 获取指定索引的列头数据（从 1 开始） |
| `addCol(index, width)` | 添加列 |
| `getVisibleCols(startX, endX)` | 获取可见区域内的列 |

**RowHeaders** 与 ColHeaders 结构对称，管理行高和垂直偏移量。

#### 5.1.5 VisibleView

**文件**：[visibleView.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/dataArchitecture/visibleView.ts)  
**继承**：`EventEmitter`

管理可见视图的尺寸参数，为 Canvas 渲染提供计算依据。

| 属性 | 类型 | 说明 |
|------|------|------|
| `rowHeaderWidth` | `number` | 行头宽度 |
| `colHeaderHeight` | `number` | 列头高度 |
| `sheetWidth` | `number` | 工作表可见宽度 |
| `sheetHeight` | `number` | 工作表可见高度 |

#### 5.1.6 CellContent / LineText

**文件**：[cellContent.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/dataArchitecture/cellContent.ts)

纯接口定义，描述单元格内容的排版信息。

**LineText 接口：**

| 字段 | 类型 | 说明 |
|------|------|------|
| `text` | `any` | 文本内容 |
| `lineWidth` | `number` | 行宽（含字间距） |
| `lineHeight` | `number` | 行高（含行间距） |
| `x` | `number` | 相对于单元格左边界的 X 坐标 |
| `y` | `number` | 相对于单元格顶部边界的 Y 坐标 |
| `editX` | `number` | 相对于编辑框左边界的 X 坐标 |
| `editY` | `number` | 相对于编辑框顶部边界的 Y 坐标 |

**CellContent 接口：**

| 字段 | 类型 | 说明 |
|------|------|------|
| `editLeft` | `number` | 编辑框左偏移 |
| `editTop` | `number` | 编辑框上偏移 |
| `charsWidth` | `number` | 文本总宽 |
| `charsHeight` | `number` | 文本总高 |
| `contentWidth` | `number` | 内容总宽 |
| `contentHeight` | `number` | 内容总高 |
| `lines` | `LineText[]` | 行文本数组 |

#### 5.1.7 Values

**文件**：[values.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/dataArchitecture/values.ts)  
**继承**：`EventEmitter`

预留的值对象类，目前仅有基本结构，为未来扩展保留。

---

### 5.2 Canvas 渲染层 (sheets)

Canvas 渲染层负责将数据绘制到 HTML5 Canvas 上，实现高性能表格渲染。

#### 5.2.1 Canvas (基类)

**文件**：[canvas.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/sheets/canvas.ts)  
**继承**：`EventEmitter`

所有 Canvas 组件的基类，提供画布基础操作和文本测量功能。

| 属性 | 类型 | 说明 |
|------|------|------|
| `canvas` | `HTMLCanvasElement` | Canvas 元素 |
| `ctx` | `CanvasRenderingContext2D` | 2D 绘图上下文 |
| `parentElement` | `HTMLElement` | 父容器 |
| `data` | `DataCollection` | 数据集合实例 |

**关键方法：**

| 方法 | 说明 |
|------|------|
| `setupContext()` | 初始化画布上下文参数（imageSmoothingEnabled、字体、对齐等） |
| `updateCanvasSize()` | 更新画布尺寸，适配容器 |
| `draw()` | 绘制画布内容（子类重写） |

#### 5.2.2 SheetCanvas

**文件**：[sheetCanvas.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/sheets/sheetCanvas.ts)  
**继承**：`Canvas`

主表格画布，负责绘制单元格内容、背景、边框、网格线和选区。

**特有功能：**
- 字体加载等待（`waitForFontsLoaded`）：使用 `Promise.race` 设置 3 秒超时
- 加载动画（`drawLoadingAnimation`）：字体加载期间显示旋转虚线圆圈
- 单元格绘制：背景色、文字内容、边框、选区高亮
- 网格线绘制

> **重要约束**：Canvas 文本渲染坐标必须使用 `Math.round()` 避免高 DPI 屏幕文字重影。OffscreenCanvas 绘制必须从 `(0,0)` 开始。

#### 5.2.3 ColHeaderCanvas

**文件**：[colHeaderCanvas.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/sheets/colHeaderCanvas.ts)  
**继承**：`Canvas`

列头画布，绘制列标题（A, B, C...）、选中状态和列宽调整手柄。

#### 5.2.4 RowHeaderCanvas

**文件**：[rowHeaderCanvas.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/sheets/rowHeaderCanvas.ts)  
**继承**：`Canvas`

行头画布，绘制行号（1, 2, 3...）、选中状态和行高调整手柄。

#### 5.2.5 EditCanvas

**文件**：[editCanvas.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/sheets/editCanvas.ts)  
**继承**：`Canvas`

编辑画布，实现单元格编辑状态下的输入框和光标显示。

#### 5.2.6 AllSelect

**文件**：[allSelect.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/sheets/allSelect.ts)  
**继承**：`EventEmitter`

全选按钮组件，位于行头与列头交汇处，点击选中所有单元格。

#### 5.2.7 Sheets (容器)

**文件**：[sheets.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/sheets/sheets.ts)  
**继承**：`EventEmitter`

工作表容器，是所有 Canvas 组件的父容器和管理者。

| 属性 | 类型 | 说明 |
|------|------|------|
| `sheetCanvas` | `SheetCanvas` | 主表格画布 |
| `colHeaderCanvas` | `ColHeaderCanvas` | 列头画布 |
| `rowHeaderCanvas` | `RowHeaderCanvas` | 行头画布 |
| `editCanvas` | `EditCanvas` | 编辑画布 |
| `allSelect` | `AllSelect` | 全选按钮 |
| `data` | `DataCollection` | 数据实例 |

**关键方法：**

| 方法 | 说明 |
|------|------|
| `draw()` | 绘制所有画布内容 |
| `updateCanvasSize()` | 更新所有画布尺寸 |
| `initVisibleView()` | 初始化可见视图尺寸 |
| `setupResizeObserver()` | 监听容器尺寸变化（仅观察 parentElement） |

**初始化流程：**
1. `createSheets()` — 创建容器 DOM
2. `initVisibleView()` — 计算可见区域尺寸
3. `createCanvas()` — 创建所有 Canvas 实例
4. `createAllSelect()` — 创建全选按钮
5. `setupResizeObserver()` — 监听尺寸变化
6. `setupEventListeners()` — 绑定事件监听

---

### 5.3 UI 组件层 (common)

通用 UI 组件，提供菜单、按钮、滚动条、面板分割等基础交互能力。

#### 5.3.1 Split

**文件**：[split.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/common/split.ts)  
**继承**：`EventEmitter`

面板分割组件，支持水平和垂直分割，通过拖拽调整面板比例。使用 IndexedDB 持久化存储布局状态。

| 特性 | 说明 |
|------|------|
| 方向 | 水平（默认）/ 垂直 |
| 持久化 | 分割位置通过 IndexedDB 保存 |
| 可见性 | `handleVisibility()` 监控元素是否可见 |
| 响应式 | 自动适配容器变化 |

#### 5.3.2 Menu

**文件**：[menu.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/common/menu.ts)  
**继承**：`EventEmitter`

下拉菜单组件，支持菜单项、分隔线、子菜单和自定义元素。

#### 5.3.3 Button (BtnBase)

**文件**：[button.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/common/button.ts)  
**继承**：`EventEmitter`

按钮组件，支持图标、文字、下拉箭头、菜单和状态管理（选中、禁用）。

#### 5.3.4 Select

**文件**：[select.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/common/select.ts)  
**继承**：`EventEmitter`

选择器组件，带下拉菜单的选项选择，用于字体选择器、字号选择器等。

#### 5.3.5 Scroller

**文件**：[scroller.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/common/scroller.ts)  
**继承**：`EventEmitter`

自定义滚动条组件，支持水平和垂直方向，可拖拽调整面板大小。

| 特性 | 说明 |
|------|------|
| 方向 | 水平 / 垂直 |
| 滑块最小尺寸 | 20px |
| 缩放因子 | 比例计算需一致应用缩放系数 |
| 防抖 | 内置防抖更新机制 |

#### 5.3.6 Navigator

**文件**：[navigator.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/common/navigator.ts)  
**继承**：`EventEmitter`

导航面板组件，支持面板切换和内容显示，使用 IndexedDB 持久化选中状态。默认导航项：`['文件', '工具', '脚本']`。

#### 5.3.7 Message

**文件**：[message.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/common/message.ts)  
**继承**：`EventEmitter`

消息提示框组件，支持成功、错误、警告、信息等类型，以及确认对话框。

---

### 5.4 工具栏层 (toolbar)

#### 5.4.1 Toolbar

**文件**：[toolbar.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/toolbar/toolbar.ts)  
**继承**：`EventEmitter`

顶部工具栏容器，管理标签页和内容区域。

| 标签页 | 名称 |
|--------|------|
| START | 开始 |
| INSERT | 插入 |
| PAGE | 页面 |
| DATA | 数据 |
| VIEW | 视图 |
| TOOLS | 工具 |

标签 ID 使用 murmurhash3 哈希生成，支持点击切换和双击折叠。

#### 5.4.2 QuickAccess

**文件**：[quickAccess.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/toolbar/quickAccess.ts)

快速访问栏，包含品牌标识、撤销、恢复、打印等快捷按钮。

#### 5.4.3 Section 分区组件

"开始"标签页下的功能分区：

| 文件 | 分区 | 功能 |
|------|------|------|
| `sectionBrush.ts` | 格式刷 | 格式复制/粘贴 |
| `sectionFont.ts` | 字体样式 | 字体选择器、字号、粗体、斜体、下划线、删除线、背景色、文本颜色 |
| `sectionAlign.ts` | 对齐方式 | 水平对齐（左/中/右）、垂直对齐（上/中/下） |
| `sectionProcess.ts` | 数据处理 | 筛选、排序、填充等 |

---

### 5.5 工具模块 (utils)

#### 5.5.1 EventEmitter

**文件**：[eventEmitter.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/utils/eventEmitter.ts)

发布-订阅模式的事件发射器，是整个项目的核心基础设施。

| 方法 | 签名 | 说明 |
|------|------|------|
| `on` | `(event: string, handler: EventHandler) => this` | 注册事件监听器 |
| `once` | `(event: string, handler: EventHandler) => this` | 注册一次性监听器 |
| `off` | `(event: string, handler?: EventHandler) => this` | 移除监听器 |
| `emit` | `(event: string, ...args: any[]) => this` | 触发事件 |
| `updateProperty` | `(propName: DataEvents, internalProp: string, value: any) => void` | 更新属性并触发变更事件 |
| `listeners` | `(event: string) => EventHandler[]` | 获取指定事件的监听器 |
| `eventNames` | `() => string[]` | 获取所有已注册事件名 |
| `removeAllListeners` | `() => this` | 移除所有监听器 |

> `updateProperty` 是数据层的核心方法：仅当新旧值不同时更新属性并 emit 变更事件，实现响应式数据绑定。

#### 5.5.2 UndoManager

**文件**：[undoManager.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/utils/undoManager.ts)

撤销/恢复管理器，基于双栈结构实现。

| 特性 | 说明 |
|------|------|
| 最大栈大小 | 100（可配置 0-100） |
| 操作组合 | `beginGroup()` / `endGroup()` 合并多步为一步 |
| 操作类型 | `cell` / `char` / `style` / `format` / `group` / `batchOperation` |
| 新操作清空 | push 新操作时自动清空重做栈 |

#### 5.5.3 ClipboardManager

**文件**：[clipboardManager.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/utils/clipboardManager.ts)

剪贴板管理器，提供内部剪贴板和系统剪贴板的读写接口。

| 方法 | 说明 |
|------|------|
| `copyToSystemClipboard(text)` | 复制文本到系统剪贴板 |
| `readFromSystemClipboard()` | 从系统剪贴板读取文本 |
| `copy(data)` / `cut(data)` / `paste(targetCell?)` | 内部剪贴板操作 |

支持两种数据类型：`cell`（单元格级）和 `char`（字符级）。

#### 5.5.4 dom

**文件**：[dom.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/utils/dom.ts)

DOM 操作工具函数集。

| 函数 | 说明 |
|------|------|
| `createDiv(options)` | 创建 div 元素 |
| `createSpan(options)` | 创建 span 元素 |
| `createInput(options)` | 创建 input 元素 |
| `createButton(options)` | 创建 button 元素 |
| `createElement(tagName, options)` | 创建任意 HTML 元素 |
| `removeElement(element)` | 移除元素 |
| `clearElement(element)` | 清空元素内容 |
| `getElementPosition(element)` | 获取元素绝对位置 |
| `hasClass` / `addClass` / `removeClass` / `toggleClass` | CSS 类操作 |

`CreateElementOptions` 接口支持：className、id、textContent、innerHTML、style、attributes、dataset、children、contentEditable、tabindex、events。

#### 5.5.5 debounce

**文件**：[debounce.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/utils/debounce.ts)

通用防抖函数。

```typescript
debounce<T extends (...args: any[]) => void>(func: T, wait: number): T
```

#### 5.5.6 fontsLoader

**文件**：[fontsLoader.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/utils/fontsLoader.ts)

字体检测与加载工具，提供两个核心功能：

**1. `detectSystemFonts(fontList?, testString?, timeout=3000)`**
- 通过 Canvas 测量文本宽度检测系统已安装字体
- 支持中英文常见字体列表（含鸿蒙黑体、宋体、微软雅黑等）
- 默认 3 秒超时

**2. `BatchFontLoader` 类**
- 批量注册和加载自定义字体（FontFace API）
- 支持超时控制（默认 30 秒）
- 支持 `applyFont()` 将字体应用到 DOM 元素

**3. `waitForFontsLoaded(fontNames?, timeout=5000)`**
- 等待字体加载完成
- 使用 `Promise.race` 防止无限等待
- 默认 5 秒超时

> **重要**：`document.fonts.ready` 可能因字体加载失败、网络问题等原因永不 resolve，必须配合超时使用。

**`fontsZhData`** 导出变量提供了常用中文字体的元数据（英文名、中文名、分类、用途描述）。

---

### 5.6 常量与事件 (constant)

**文件**：[constant.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/constant.ts)

定义了项目所有常量和事件枚举。

#### 核心常量

| 常量 | 值 | 说明 |
|------|-----|------|
| `DEFAULT_CELL_WIDTH` | 78 | 默认列宽 |
| `DEFAULT_CELL_HEIGHT` | 24 | 默认行高 |
| `MIN_WIDTH` | 20 | 最小列宽 |
| `MIN_HEIGHT` | 15 | 最小行高 |
| `DEFAULT_FONT_SIZE` | 16 | 默认字号 |
| `DEFAULT_FONT_FAMILY` | `'HarmonyOS Sans SC Regular'` | 默认字体 |
| `ROW_HEADER_PADDING` | 12 | 行头内边距 |
| `CELL_PADDING` | 2 | 单元格内边距 |
| `GRID_LINE_COLOR` | `'#ccc'` | 网格线颜色 |
| `HEADER_BG_COLOR` | `'whitesmoke'` | 表头背景色 |
| `SCROLLER_THUMB_MIN_SIZE` | 20 | 滚动条滑块最小尺寸 |

#### 事件枚举 DataEvents

由四个枚举合并而成：

**SheetProperty（工作表事件）：**
- `IS_EDITTING_CHANGED` — 编辑状态变更
- `SELECTION_CHANGED` — 选区变更
- `ANCHOR_SELECTION_CHANGED` — 锚点选区变更
- `ZOOM_CHANGED` — 缩放变更
- `VALUES_CHANGED` — 值变更
- `ALL_COL_WIDTH_CHANGED` / `ALL_ROW_HEIGHT_CHANGED` — 列宽/行高全局变更
- `OFFSETX_CHANGED` / `OFFSETY_CHANGED` — 水平/垂直偏移变更

**CellProperty（单元格事件）：**
- `ACTIVED_CELL_CHANGED` — 活动单元格变更
- `FONT_FAMILY_CHANGED` / `FONT_SIZE_CHANGED` — 字体属性变更
- `BORDER_*_CHANGED` — 边框宽度变更
- `TEXT_ALIGN_CHANGED` / `ALIGN_ITEMS_CHANGED` — 对齐变更
- `BACKGROUND_COLOR_CHANGED` — 背景色变更

**ValueProperty（字符事件）：**
- `CHAR_CHANGED` — 字符内容变更
- `CHAR_FONT_*_CHANGED` — 字符级样式变更

**VisibleViewProperty（视图事件）：**
- `ROW_HEADER_WIDTH_CHANGED` / `COL_HEADER_HEIGHT_CHANGED`
- `SHEET_WIDTH_CHANGED` / `SHEET_HEIGHT_CHANGED`

**MouseLocation 枚举：**
- `IN_All_SELECT` / `IN_COL_HEADER` / `IN_ROW_HEADER` / `IN_CELL` / `IN_ANCHOR` 等

---

## 6. 类继承关系图

```
EventEmitter (utils/eventEmitter.ts)
├── DataCollection (dataArchitecture/dataCollection.ts)
├── Char (dataArchitecture/char.ts)
├── Cell (dataArchitecture/cell.ts) — 注: Cell 继承 EventEmitter，组合 Char[]
├── ColHeaders (dataArchitecture/header.ts)
├── RowHeaders (dataArchitecture/header.ts)
├── VisibleView (dataArchitecture/visibleView.ts)
├── Value (dataArchitecture/values.ts)
├── Canvas (sheets/canvas.ts)
│   ├── SheetCanvas (sheets/sheetCanvas.ts)
│   ├── ColHeaderCanvas (sheets/colHeaderCanvas.ts)
│   ├── RowHeaderCanvas (sheets/rowHeaderCanvas.ts)
│   └── EditCanvas (sheets/editCanvas.ts)
├── AllSelect (sheets/allSelect.ts)
├── Sheets (sheets/sheets.ts)
├── Split (common/split.ts)
├── Menu (common/menu.ts)
├── BtnBase (common/button.ts)
├── Select (common/select.ts)
├── Scroller (common/scroller.ts)
├── Navigator (common/navigator.ts)
├── Message (common/message.ts)
├── Toolbar (toolbar/toolbar.ts)
├── QuickAccess (toolbar/quickAccess.ts)
└── Section* (toolbar/start/section*.ts)
```

---

## 7. 数据流与事件驱动

项目采用 **事件驱动的单向数据流** 模式：

```
用户交互 → DataCollection 属性变更 → emit(DataEvents) → App.setupDataListeners() → Sheets.draw()
                                              ↓
                                         syncValuesToDB()
```

**App 事件监听映射：**

| 事件 | 响应行为 |
|------|----------|
| `ACTIVED_CELL_CHANGED` | 设置 `isEditting = false` |
| `OFFSETX_CHANGED` (colHeaders) | `sheets.draw()` |
| `OFFSETY_CHANGED` (rowHeaders) | `sheets.draw()` |
| `ALL_ROW_HEIGHT_CHANGED` | `sheets.draw()` |
| `SELECTION_CHANGED` | `sheets.draw()` |
| `ANCHOR_SELECTION_CHANGED` | `sheets.draw()` |
| `IS_EDITTING_CHANGED` | `if (!isEditting) sheets.draw()` |
| `ZOOM_CHANGED` | `sheets.draw()` |
| `COLUMNS_CHANGED` | `sheets.draw()` |
| `ROWS_CHANGED` | `sheets.draw()` |
| `VALUES_CHANGED` | `sheets.draw()` + `data.syncValuesToDB()` |

**数据读写流程：**

1. **启动时**：`DataCollection` 构造函数 → `syncValuesFromDB()` → 从 IndexedDB 读取 → 反序列化为 `Cell[]`
2. **用户编辑**：修改 Cell/Char → setter 调用 `updateProperty()` → emit 事件 → App 监听 → `draw()` 重绘
3. **值变更**：`DataCollection.values` setter → `updateProperty(VALUES_CHANGED)` → emit → `syncValuesToDB()` 持久化

---

## 8. 数据持久化 (IndexedDB)

**文件**：[indexDB.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/dataArchitecture/indexDB.ts)

| 配置 | 值 |
|------|-----|
| 数据库名称 | `silkspaces` |
| 版本号 | `1` |

**对象存储（Object Store）：**

| 存储名 | 主键 | 索引 | 结构 |
|--------|------|------|------|
| `settings` | `key` | `type`(非唯一), `key`(唯一) | `{ key, value, type }` |
| `values` | `sheetName` | `sheetName`(唯一) | `{ sheetName, value }` |

**主要方法：**

| 方法 | 说明 |
|------|------|
| `init()` | 初始化数据库连接 |
| `setSetting(key, value)` | 保存配置项 |
| `getSetting(key)` | 获取配置项 |
| `setValue(sheetName, value)` | 保存工作表数据（JSON 字符串） |
| `getValue(sheetName)` | 获取工作表数据 |
| `getAllValues()` | 获取所有工作表数据 |
| `deleteValue(sheetName)` | 删除工作表数据 |

---

## 9. 依赖关系

### 外部依赖

| 依赖 | 用途 | 类型 |
|------|------|------|
| `murmurhash-js` | 工具栏标签 ID 哈希生成 | devDependency |

> 项目设计理念为尽量不使用第三方库，`murmurhash-js` 是唯一的外部运行时依赖。

### 内部模块依赖图

```
index.ts
  └── App
        ├── Split ← dom, EventEmitter, IndexDB
        ├── Menu ← dom, EventEmitter
        ├── Toolbar ← dom, EventEmitter, DataCollection, Menu, murmurhash-js
        │   ├── QuickAccess ← dom, DataCollection, Menu, Button
        │   └── Section* ← dom, DataCollection, Menu, Button, Select
        ├── Sheets ← dom, EventEmitter, debounce, DataCollection
        │   ├── SheetCanvas ← Canvas ← DataCollection, CellContent, EventEmitter
        │   ├── ColHeaderCanvas ← Canvas
        │   ├── RowHeaderCanvas ← Canvas
        │   ├── EditCanvas ← Canvas
        │   └── AllSelect ← dom, DataCollection, EventEmitter
        ├── Scroller ← dom, EventEmitter, debounce, DataCollection
        └── Navigator ← dom, EventEmitter, IndexDB

DataCollection ← EventEmitter, IndexDB, Cell, Char, ColHeaders, RowHeaders, VisibleView
Cell ← EventEmitter, Char
Char ← EventEmitter
IndexDB ← 原生 IDB API
```

### 循环依赖注意事项

`Cell` 类与 `Char` 类存在组合关系（Cell 包含 `Char[]`），但无循环引用。`DataCollection` 持有所有数据类的实例作为属性。

---

## 10. 项目运行方式

### 环境要求

- Node.js（建议 16+）
- npm

### 安装依赖

```bash
npm install
```

### 开发模式

```bash
npm run dev
```

启动 webpack-dev-server，支持热更新。默认访问 `http://localhost:8080`。

### 生产构建

```bash
npm run build
```

输出到 `dist/` 目录，包含：
- `index.html` — 入口页面
- `main.[hash].js` — 打包后的 JS
- `main.[contenthash].css` — 提取的 CSS
- `assets/` — 字体、图片等静态资源

### 应用初始化流程

```typescript
// src/index.ts
import { App } from './core/app';
const app = new App("#worktop");
```

1. 获取 `#worktop` DOM 元素
2. `initContainerSplit()` — 创建容器分割
3. `initMenu()` — 创建菜单
4. `initToolbar()` — 创建工具栏
5. `initSheets()` — 创建工作表
6. `initBottomToolsSplit()` / `initRightToolsSplit()` — 创建工具分割线
7. `initScroller()` — 创建滚动条
8. `setupDataListeners()` — 注册数据事件监听
9. `initNavigator()` — 创建导航面板（最后初始化）

---

## 附录：工程约束与注意事项

1. **Canvas 坐标**：必须使用 `Math.round()` 而非 `Math.floor()`，避免高 DPI 屏幕文字重影
2. **OffscreenCanvas**：绘制起点必须为 `(0,0)`，否则产生位置偏移
3. **动画冲突**：启动新动画前必须取消已有动画，防止视觉闪烁
4. **DataCollection.values**：getter 返回新数组，必须通过 setter 赋值才能触发持久化
5. **字体加载**：`document.fonts.ready` 必须配合超时（3 秒），防止字体加载失败导致永久等待
6. **ResizeObserver**：仅观察 `parentElement`，避免重复触发 resize 事件
7. **拖拽中的动画**：拖拽操作期间必须禁用动画，防止阴影效果
