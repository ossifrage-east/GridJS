# 为 silkspaces 增加 Pyodide 支持（实现方案）

## 1. 概述（Summary）

项目定位是「使用 canvas 制作的 web 表格，可以使用 python 进行宏编程」（见 [package.json](file:///c:/Users/Lenovo/Desktop/silkspaces/package.json#L4)），`src/core/assistant/pyEditor.ts` 已经是一个可用的 Python 代码编辑器（语法高亮 + 行号 + 缩进辅助 + 按工作表分键持久化），但**只能写不能跑**。

本次改动把 Pyodide（浏览器内 CPython，WebAssembly）接进项目，让「脚本」页签里的代码真正可以执行，并能让脚本直接读写当前工作表的单元格。

已与用户确认的四项决策：

| 决策项             | 选定方案                                                                                      |
| --------------- | ----------------------------------------------------------------------------------------- |
| Pyodide 运行时加载方式 | npm 依赖打包进项目（`node_modules/pyodide` 经 copy-webpack-plugin 拷进 `dist/pyodide/`），不走 CDN、不手工放置 |
| Python 执行环境     | Web Worker                                                                                |
| 脚本对工作表数据的访问     | 可读**也可写回**（写回纳入撤销栈）                                                                       |
| 运行入口与输出         | 「脚本」页签工具栏加「运行 / 停止」按钮 + 页签内输出面板                                                           |

本轮补充确认的三项：

| 决策项                | 选定方案                                                                                                                                |
| ------------------ | ----------------------------------------------------------------------------------------------------------------------------------- |
| Python 侧工作表 API 风格 | **函数式**：`get_cell("A1")` / `set_cell("A1", v)` / `get_range("A1:C3")` / `set_range(...)` / `sheet_name` / `row_count` / `col_count` |
| 运行与输出交互            | **实时流式输出 + 停止按钮**，并内置 30 秒超时保护                                                                                                      |
| 第三方包（numpy/pandas） | **不支持**，只用 Pyodide 自带 Python 标准库（完全离线）                                                                                              |

## 2. 现状分析（Current State Analysis）

### 2.1 已有的 Worker 范式（可直接镜像）

[parserClient.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/utils/parser/parserClient.ts) + [parser.worker.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/utils/parser/parser.worker.ts) 构成项目既定的 Worker 模式，本项目新增的任何 Worker 都应沿用它：

* Worker 脚本**固定文件名、相对文档路径加载**：`const WORKER_URL = 'parser.worker.js'` → `new Worker(WORKER_URL)`（[parserClient.ts L78-L157](file:///c:/Users/Lenovo/Desktop/silkspaces/src/utils/parser/parserClient.ts#L78-L157)）。这样支持部署在任意子路径。

* `id` 序号请求 / 响应协议 + `pending: Map<number, {resolve, reject}>`（[L96-L101](file:///c:/Users/Lenovo/Desktop/silkspaces/src/utils/parser/parserClient.ts#L96-L101)）。

* 惰性单例 Worker + `w.onerror` 兜底 + `failAllPending(reason)`（[L134-L194](file:///c:/Users/Lenovo/Desktop/silkspaces/src/utils/parser/parserClient.ts#L134-L194)）。

* Worker 侧不引入 `WebWorker` lib（会与 DOM lib 冲突），改用最小类型视图 `const ctx = self as unknown as { onmessage: ...; postMessage: ... }`（[parser.worker.ts L38-L46](file:///c:/Users/Lenovo/Desktop/silkspaces/src/utils/parser/parser.worker.ts#L38-L46)）。

* 入口异步逻辑包在 `void (async () => { ... })()` 里（[L61-L86](file:///c:/Users/Lenovo/Desktop/silkspaces/src/utils/parser/parser.worker.ts#L61-L86)），避免顶层 await（tsconfig `target: ES5`）。

### 2.2 构建配置

[webpack.config.js](file:///c:/Users/Lenovo/Desktop/silkspaces/webpack.config.js)：

* 导出**数组** `[mainConfig, workerConfig]`（[L182](file:///c:/Users/Lenovo/Desktop/silkspaces/webpack.config.js#L182)），workerConfig 见 [L149-L180](file:///c:/Users/Lenovo/Desktop/silkspaces/webpack.config.js#L149-L180)：`target: 'webworker'`、`output.filename` 固定名、`clean: false`。

* `dist` 的清空放在**构建前一次性** **`fs.rmSync`**（[L9-L13](file:///c:/Users/Lenovo/Desktop/silkspaces/webpack.config.js#L9-L13)），两个配置共用 dist 且都不开 `output.clean`。

* 主配置 resolve.fallback 关掉 `fs`/`path`/`vm`（[L139-L146](file:///c:/Users/Lenovo/Desktop/silkspaces/webpack.config.js#L139-L146)）。

* 没有 `devServer` 配置、没有 CSP（[src/index.html](file:///c:/Users/Lenovo/Desktop/silkspaces/src/index.html) 无 CSP meta）→ wasm 的 `wasm-unsafe-eval` 不会被拦。

* tsconfig：`target: ES5`、`module: ES2020`（动态 `import()` 要求 ≥ es2020，见假设 2）、`lib: ["ES6","DOM","ES2022"]`、`moduleResolution: node`、`skipLibCheck: true`（[tsconfig.json](file:///c:/Users/Lenovo/Desktop/silkspaces/tsconfig.json)）。

### 2.3 脚本页签现状

[pyEditor.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/assistant/pyEditor.ts)：

* 双层结构：`pre.py-editor-highlight`（高亮层）+ `textarea.py-editor-input`（透明输入层），外层 `div.py-editor-body` 里还有 `pre.py-editor-gutter` 行号栏，顶部是 `div.py-editor-toolbar`（[L193-L224](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/assistant/pyEditor.ts#L193-L224)）。

* 工具栏当前只有 `TITLE`（"Python 脚本"）与 `STATUS`（行/列指示）（[L193-L200](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/assistant/pyEditor.ts#L193-L200)）→ **没有按钮位、没有输出区**。

* 公开 API 只有 `getValue()` / `setValue()` / `getElements()`（[L236-L259](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/assistant/pyEditor.ts#L236-L259)）。

* 持久化：IndexedDB 设置表，键 `py-editor:script:<工作表名>`，防抖 400ms（[L126-L131](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/assistant/pyEditor.ts#L126-L131)、[L595-L640](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/assistant/pyEditor.ts#L595-L640)）。

* 默认脚本 `defaultScript(sheetName)` 以 `def main()` + `if __name__ == "__main__": main()` 结尾（[L172-L186](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/assistant/pyEditor.ts#L172-L186)）→ **运行环境里必须把** **`__name__`** **设为** **`'__main__'`，否则** **`main()`** **不会被调用、什么都不打印**。

* 监听 `DataEvents.ACTIVED_SHEET_NAME_CHANGED` 处理工作表切换（[L231-L233](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/assistant/pyEditor.ts#L231-L233)）；`sheetNameSpace.ts` 目前在仓库中**是空文件**、无任何引用（本次新建内容）。

挂载点在 [app.ts L420-423](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/app.ts#L420-L423)：

```ts
this.pyEditor = new PYEditor({ data: this.data });
navigator.navContentItemMount(NAV_PANEL[2], this.pyEditor.getElements().container);
```

### 2.4 DataCollection 的读写 API（写回路径的依据）

[dataCollection.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/dataArchitecture/dataCollection.ts) 中**没有** `setCell` / `setCellValue` / `updateCell`。可用的读写入口：

| 用途                | 入口                                                                                                                                                                                                                                                                                  |
| ----------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 读全部内容             | `get values(): Cell[]`（活引用，[L1072-L1074](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/dataArchitecture/dataCollection.ts#L1072-L1074)）                                                                                                                                       |
| 按名读一格             | `getCellByName(name)`（O(1) 索引，[L1545](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/dataArchitecture/dataCollection.ts#L1545)）                                                                                                                                                |
| 表尺寸               | `rowHeaders.length` / `colHeaders.length`（公开字段，[L187-L189](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/dataArchitecture/dataCollection.ts#L187-L189)；用法见 [L2797-L2798](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/dataArchitecture/dataCollection.ts#L2797-L2798)） |
| 名 ↔ 坐标            | `getCellName(col, row)`（[L3342](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/dataArchitecture/dataCollection.ts#L3342)）、`getCellColAndRow(cell)`（[L3358](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/dataArchitecture/dataCollection.ts#L3358)）                      |
| **稀疏写回 + 撤销**（推荐） | `runWithCellsUndo(label, cells: Cell[], action)`（[L2657-L2683](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/dataArchitecture/dataCollection.ts#L2657-L2683)）——代价与影响面成正比、不受 `LARGE_TABLE_UNDO_CELLS` 阈值限制，查找替换正在用                                                             |
| 矩形写回 + 撤销         | `runWithCellPatchUndo(label, rect, action)`（[L2608-L2639](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/dataArchitecture/dataCollection.ts#L2608-L2639)）——代价与矩形面积成正比，且会把区域内不存在的格子记成 `null`（撤销能删掉新增格）                                                                          |
| 原地改后提交            | `commitValues()`（[L1758](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/dataArchitecture/dataCollection.ts#L1758)）                                                                                                                                                             |
| 显式广播              | `notifyValuesChanged()`（[L679-L682](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/dataArchitecture/dataCollection.ts#L679-L682)）                                                                                                                                              |

`runWithCellsUndo` 内部已包含 `_pruneEmptyCells()` + `markContentChanged()` + `emit(VALUES_CHANGED)` + 撤销入栈（[L2665-L2682](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/dataArchitecture/dataCollection.ts#L2665-L2682)），因此**写回只需套一层** **`runWithCellsUndo`**，不必再手动 `commitValues()` / `notifyValuesChanged()`。

新建格子的既有写法（照抄 `pastePlainText`，[L5060-L5076](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/dataArchitecture/dataCollection.ts#L5060-L5076)）：

```ts
const cell = new Cell();
cell.cell = 'A1';
cell.addValue('文本');      // cell.ts L622
data.values.push(cell);     // values 是活引用，原地 push 合法
```

`Cell.addValue(value)`（[cell.ts L622-L629](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/dataArchitecture/cell.ts#L622-L629)）会用 `Array.from` 拆字符再建 `Char`，**不带任何样式**；`cell.clearValue()`（[L634-L636](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/dataArchitecture/cell.ts#L634-L636)）清空文本。

## 3. 方案设计

### 3.1 数据流

```
「运行」按钮（pyEditor，主线程）
   ↓ 取脚本文本 + buildSnapshot(data) 生成快照 JSON
pythonClient.runPython(...)  ──►  python.worker.js（Web Worker）
   │                                  ↓ import('./pyodide/pyodide.mjs')  ← module worker 内动态 import
   │                                  ↓ loadPyodide({ indexURL: './pyodide/' })
   │   ◄── {type:'stdout'|'stderr'} ──  │ setStdout/setStderr({batched}) 流式回调
   │                                  ↓ 注入 snapshot → 执行脚本 → 收集 _WRITES
   ◄── {type:'result', writesJson} ───┘
   ↓ applyWrites(data, writesJson)  → data.runWithCellsUndo(...)（纳入撤销 + 重绘）
   ↓ 输出面板追加「已写入 N 个单元格」
```

**为什么两边都传 JSON 字符串而不是直接传对象**：Pyodide 的 `globals.set` 把 JS 对象包成 `JsProxy`，需要在 Python 侧显式 `.destroy()` 才不泄漏；JSON 字符串是纯值，两端解析/序列化一次即可，规避代理生命周期问题。快照回主线程同理。

**为什么不回退主线程**：`parserClient` 在 Worker 不可用时回退主线程，是因为「打开文件」是核心路径、回退后行为一致。但 Pyodide 跑在主线程会**冻结整个 UI**（WASM 计算不释放事件循环），与「Worker 执行」这一决策直接冲突。故 `pythonClient` 的降级策略是**报错**：Worker 建不起来（如 `file://` 下同源限制）时，输出面板显示明确的错误提示，而不是悄悄改用主线程。

### 3.2 Worker 消息协议

```ts
// src/utils/python/pythonClient.ts 导出，worker 侧 import type 复用
export interface PythonRunRequest {
    id: number;
    script: string;          // 用户脚本全文
    sheetName: string;       // 当前工作表名（注入到命名空间的 sheet_name）
    snapshotJson: string;    // {"A1": "文本", ...}：当前工作表非空格子
    rowCount: number;
    colCount: number;
    namespaceSource: string; // silkspaces 模块的 Python 源码（由主线程随请求下发）
}

export type PythonWorkerMessage =
    | { id: number; type: 'stdout'; text: string }
    | { id: number; type: 'stderr'; text: string }
    | { id: number; type: 'result'; writesJson: string }   // {"A1": "新值", ...}
    | { id: number; type: 'error'; message: string };
```

`namespaceSource` 随请求下发（而不是 worker 里 `import` 那个 TS 模块），是为了**避免 worker bundle 被** **`dataCollection.ts`** **拖进来**：`sheetNameSpace.ts` 要在主线程用 `DataCollection` / `Cell`（运行时依赖），若 worker 静态 import 它，开发模式下 webpack 会把整条数据层依赖链打进 `python.worker.js`。

### 3.3 Python 侧命名空间 API（函数式）

模块名 `silkspaces`，用户脚本头部注入 `from silkspaces import *`，因此脚本里可直接写 `get_cell("A1")`。

| API                       | 说明                                       |
| ------------------------- | ---------------------------------------- |
| `sheet_name`              | `str`，当前工作表名（＝脚本存储键）                     |
| `row_count` / `col_count` | 当前表格的行数 / 列数                             |
| `get_cell(name)`          | 取文本，空格子返回 `""`；已写入未提交的值优先                |
| `set_cell(name, value)`   | 写文本；`value is None` 等价清空；其余 `str(value)` |
| `clear_cell(name)`        | 清空                                       |
| `get_range(a1)`           | `"A1:C3"` → `list[list[str]]`（越界补 `""`）  |
| `set_range(a1, values)`   | `values` 为二维可迭代；单格 `"A1"` 时接受一维序列        |
| `iter_cells()`            | 迭代 `(name, text)`，含快照与本次写入               |

单元格名归一化：列字母统一大写、去掉 `$`（`a1` / `$A$1` → `A1`）。

## 4. 具体改动（Proposed Changes）

### 4.1 `package.json`

* `dependencies`（新增字段）加入 `"pyodide": "<最新稳定版，≥0.26>"`。**安装后必须以** **`node_modules/pyodide/package.json`** **实际内容为准核对文件清单与版本号**，再据实调整 4.2 的拷贝范围。

* `devDependencies` 加入 `"copy-webpack-plugin": "^12.x"`。

### 4.2 `webpack.config.js`

1. 顶部新增 `const CopyWebpackPlugin = require('copy-webpack-plugin');`
2. 主配置 `plugins` 增加拷贝规则，把 pyodide 运行时资产放到固定子目录 `dist/pyodide/`：

```js
new CopyWebpackPlugin({
  patterns: [{
    from: resolve(__dirname, './node_modules/pyodide'),
    to: 'pyodide',
    // 只保留运行时会用到的 5 个文件，其余（类型声明 / sourcemap / 文档 / 包描述 /
    // 经典脚本 pyodide.js / console 示例页）剔除
    globOptions: { ignore: ['**/*.d.ts', '**/*.map', '**/README.md', '**/package.json', '**/console*.html', '**/pyodide.js'] },
  }],
})
```

1. 新增第三个配置 `pythonWorkerConfig`（与 `workerConfig` 同构，只改 name / entry / 输出名）：

```js
const pythonWorkerConfig = {
  name: 'python-worker',
  mode: 'development',
  target: 'webworker',
  experiments: { outputModule: true },   // Pyodide 314 拒绝 classic worker，必须产出 ESM
  entry: resolve(__dirname, './src/utils/python/python.worker.ts'),
  output: { filename: 'python.worker.js', path: DIST, clean: false, module: true },
  module: { rules: [{ test: /\.tsx?$/, use: 'ts-loader', exclude: /node_modules/ }] },
  resolve: { extensions: ['.ts', '.js'], fallback: { fs: false, path: false, vm: false } },
};
```

1. 导出改为 `module.exports = [mainConfig, workerConfig, pythonWorkerConfig];`

> 注意：主配置的 `output.filename` 带 hash，拷贝出来的 `dist/pyodide/*` 不带 hash —— 这是有意的，`indexURL` 必须是稳定路径。

### 4.3 `src/utils/python/python.worker.ts`（新建）

职责：加载 Pyodide → 注册 `silkspaces` 模块 → 执行脚本 → 流式回传输出 → 收集回写。

关键实现点：

* **类型视图**：沿用 parser worker 的写法，不引 `WebWorker` lib：

  ```ts
  const ctx = self as unknown as {
      onmessage: ((e: MessageEvent<PythonRunRequest>) => void) | null;
      postMessage: (m: PythonWorkerMessage) => void;
  };
  ```

* **加载 Pyodide**：`const { loadPyodide } = await import(/* webpackIgnore: true */ new URL('pyodide.mjs', PYODIDE_DIR).href)`，其中 `PYODIDE_DIR = new URL('./pyodide/', self.location.href).href`（以 worker 自身 URL 为基准算绝对地址，子路径部署同样成立）；`indexURL` 必须显式传入，否则 worker 环境反推模块位置会失败。类型用 `import type { PyodideAPI } from 'pyodide'`（编译期擦除，不进 bundle）。

* **惰性初始化 + 串行**：`let pyodidePromise: Promise<PyodideInterface> | null = null`，首个 `run` 消息时 `loadPyodide({ indexURL: './pyodide/' })`；初始化期间到达的后续请求排队（或直接拒绝，因为 pyEditor 会禁用「运行」按钮——采用后者，简单）。

* **注册命名空间**：把 `namespaceSource` 经 `pyodide.globals.set('__NS_SOURCE__', src)` 注入，再执行一段固定的 bootstrap：

  ```python
  import sys, types
  _mod = types.ModuleType('silkspaces')
  exec(__NS_SOURCE__, _mod.__dict__)
  sys.modules['silkspaces'] = _mod
  del __NS_SOURCE__
  ```

* **流式输出**：`pyodide.setStdout({ batched: (s) => ctx.postMessage({ id, type: 'stdout', text: s }) })`、`setStderr` 同理（每次 run 前重设以带上当前 id）。

* **执行**：注入快照 → 调命名空间的 `_load(...)` → 运行 `'__name__ = "__main__"\nfrom silkspaces import *\n' + script` → 调 `_dump()` 取回写 JSON：

  ```python
  import silkspaces as _ns, json as _json
  _ns._load(_json.loads(__SNAPSHOT__), __SHEET__, __ROWS__, __COLS__)
  ```

  （快照同样以 JSON 字符串经 globals 注入。）

* **收尾**：`postMessage({ id, type: 'result', writesJson })`；异常 `{ id, type: 'error', message }`（把 Python traceback 文本原样带回，输出面板按 stderr 样式显示）。

* 整个入口异步逻辑包在 `void (async () => { ... })()` 内（**不能用顶层 await**，tsconfig `target: ES5`）。

### 4.4 `src/utils/python/pythonClient.ts`（新建）

镜像 [parserClient.ts](file:///c:/Users/Lenovo/Desktop/silkspaces/src/utils/parser/parserClient.ts) 的结构，但**无主线程回退**：

* `const WORKER_URL = 'python.worker.js'`；`pending: Map<number, { resolve, reject, onStdout, onStderr }>`（流式需要按 id 分发多条消息，故 entry 带回调）。

* `export class PythonWorkerUnavailableError extends Error`

* `export async function runPython(options): Promise<{ writesJson: string }>`

  * `options: { script, sheetName, snapshotJson, rowCount, colCount, namespaceSource, onStdout?, onStderr?, timeoutMs? }`

  * 内部 `setTimeout` 超时（默认 `PythonClient.TIMEOUT_MS = 30000`）：超时 → `stopPython()` + `reject(new Error('Python 脚本执行超时（30 秒），已强制停止'))`

* `export function stopPython(): void`：`worker.terminate()` → 清 `pending`（全部 reject `'已停止'`）→ `worker = null`（**保留** **`pyodidePromise`** **置空**，下次运行重新加载 Pyodide 并提示「正在加载 Python 运行环境…」）

* `export function isPythonWorkerActive(): boolean`（诊断用，与 parser 侧命名一致）

* `w.onerror` / `w.onmessageerror` → 拒绝所有在途请求并置 `workerUnavailable = true`；`onerror` 的告警文案要点明「Pyodide 不会回退主线程」与 `file://` 同源限制这两个最常见原因。

### 4.5 `src/core/assistant/sheetNameSpace.ts`（填充空文件）

主线程侧桥接 + Python 模块源码，三块导出：

1. `export const PY_NAMESPACE_SOURCE: string` —— 3.3 表的 Python 源码（模块体，不含 `ModuleType` 注册逻辑；注册由 worker bootstrap 负责）。内部私有状态：

   * `_SNAPSHOT: dict[str,str]`、`_WRITES: dict[str,str]`

   * `_load(snapshot_json, sheet_name, rows, cols)` / `_dump()`（worker 调用；下划线开头，`from silkspaces import *` 不会导出它们）

   * `_normalize(name)`：列字母大写、去 `$`，非法（不含行号）时抛 `ValueError`

   * `get_cell` / `set_cell` / `clear_cell` / `get_range` / `set_range` / `iter_cells`

   * 用 JSON 解析而非 `eval`；`get_range` 用 A1 范围解析（可复用同一套列字母换算）

2. `export interface SheetSnapshot { sheetName: string; rowCount: number; colCount: number; cells: [string, string][]; json: string }`
   `export function buildSnapshot(data: DataCollection): SheetSnapshot`

   * 遍历 `data.values`，跳过 `cell.getText()` 为空的格子 → `cells`

   * `rowCount = data.rowHeaders.length`、`colCount = data.colHeaders.length`

   * `json = JSON.stringify(Object.fromEntries(cells))`（同名只可能一个，对象即键表）

3. `export function applyWrites(data: DataCollection, writesJson: string, label: string): number`

   * `JSON.parse` 得 `Record<string, string>`；空对象直接返回 0（不产生空撤销记录）

   * 对每个 `name`：`data.getCellByName(name)` 命中则复用，未命中则 `new Cell()` + `cell.cell = name` 并记入 `created` 数组

   * `data.runWithCellsUndo(label, cells, () => { ... })`：`cells = 既有 + 新建`

     * action 内：值非空 → `cell.addValue(value)`；值为空串 → `cell.clearValue()`；`created` 里的格子 `data.values.push(cell)`

   * 返回写入格数

   > 说明：`runWithCellsUndo` 的 before 快照对「本次新建的空格子」是空内容克隆，撤销会把空格子放回 `values`（下一次 `commitValues()` 的 `_pruneEmptyCells` 会清掉）。这是该 API 的既有语义（矩形版 `runWithCellPatchUndo` 用 `null` 表达「不存在」，能精确删格，但代价是矩形面积）。宏脚本的写入通常是稀疏的，故统一走稀疏版，**不**为这点边界额外引入矩形版分支。

### 4.6 `src/core/assistant/pyEditor.ts`

* 新增私有字段：`runButton`、`stopButton`、`output`（输出面板容器）、`isRunning = false`、`runTimer`（可选，若超时交给 client 则不需要）。

* 工具栏结构改为：`TITLE` + `ACTIONS`（内含 `RUN`、`STOP`）+ `STATUS`；`STOP` 初始 `disabled`（用 `attributes: { disabled: 'true' }`）。

* 新增输出面板：`div.py-editor-output`，挂在 `body` 之后（`container` 内、`body` 外），固定高度、`overflow: auto`；每次 run 开始时清空并显示。

* `handleRun()`：

  1. 若 `isRunning` 直接返回；
  2. 置 `isRunning`、按钮状态互换；
  3. `const snapshot = buildSnapshot(this.data)`；
  4. `await runPython({ script: this.getValue(), ...snapshot, namespaceSource: PY_NAMESPACE_SOURCE, onStdout: t => this.appendOutput(t, false), onStderr: t => this.appendOutput(t, true) })`
  5. 成功 → `const n = applyWrites(this.data, res.writesJson, '运行 Python 脚本')` → `appendOutput(n ? ` 已写入 \${n} 个单元格 ` : '运行完成（未写入单元格）', false)`
  6. `catch` → `appendOutput(err.message, true)`
  7. `finally` → 恢复按钮状态

* `handleStop()` → `stopPython()` → `appendOutput('已停止', true)`（`runPython` 的 reject 回调会各自补一条，注意不要重复刷屏：stop 时把 `isRunning` 置位为「用户主动停止」以免 finally 再追加）。

* `appendOutput(text, isError)`：按 `\n` 拆行、逐行建 `div.py-editor-output-line`（错误加 `py-editor-output-error`），追加后把面板 `scrollTop` 拉到底（实时流式必须滚动到底，否则用户看不到最新输出）。

* **默认脚本模板补充 API 注释**：在 `defaultScript`（[L172-L186](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/assistant/pyEditor.ts#L172-L186)）里加两三行注释列出可用 API（`sheet_name` / `row_count` / `col_count` / `get_cell` / `set_cell` / `get_range` / `set_range`）。`LEGACY_DEFAULT_SCRIPT` 与迁移逻辑**保持不动**（它迁移的是「首行固定标题」的历史模板，与本次新增注释不冲突）。

### 4.7 `src/core/constant.ts`

在 `PY_EDITOR_CLASS_NAMES`（[L150-L160](file:///c:/Users/Lenovo/Desktop/silkspaces/src/core/constant.ts#L150-L160)）末尾追加：

```ts
ACTIONS: 'py-editor-actions',
RUN: 'py-editor-run',
STOP: 'py-editor-stop',
OUTPUT: 'py-editor-output',
OUTPUT_LINE: 'py-editor-output-line',
OUTPUT_ERROR: 'py-editor-output-error',
```

### 4.8 `src/core/style/pyEditor.scss`

* `.py-editor-toolbar`：保持 space-between；新增 `.py-editor-actions`（`display: flex; gap: 6px`）。

* `.py-editor-run` / `.py-editor-stop`：小尺寸按钮（`font-size: 12px`、`padding: 2px 10px`、圆角、边框），`[disabled]` 用灰度 + `cursor: default`。

* `.py-editor-output`：`flex: none; height: 140px; overflow: auto; border-top: 1px solid #e0e0e0; background: #fbfbfb; font-family: Consolas,'Courier New',monospace; font-size: 12px; line-height: 18px; padding: 6px 8px; white-space: pre-wrap; word-break: break-all;`

* `.py-editor-output-line { min-height: 18px; }`、`.py-editor-output-error { color: #c62828; }`

## 5. 假设与决策（Assumptions & Decisions）

1. **pyodide npm 包的文件布局**（已实测核对，版本 314.0.7）：运行时文件直接在包根目录，共 5 个 —— `pyodide.mjs`、`pyodide.asm.mjs`、`pyodide.asm.wasm`（约 9.6 MB）、`python_stdlib.zip`（约 2.5 MB）、`pyodide-lock.json`。**包内没有 `pyodide.asm.js`**。
2. **必须用 ESM 入口 + module worker**（原「`importScripts` + 经典脚本」方案已被实测推翻）：Pyodide 314 初始化时做 `isClassicWorker()` 检测 —— 试调 `importScripts("data:text/javascript,")`，成功即判定当前是 classic worker，随即抛 `Classic web workers are not supported`；这与「是否自己调用过 `importScripts`」无关（classic worker 里该函数本就存在）。由此：(a) `pythonWorkerConfig` 开 `experiments.outputModule` + `output.module`，`pythonClient` 用 `new Worker(url, { type: 'module' })` 构造；(b) worker 内以运行时动态 `import()` 拉 `pyodide.mjs`（配 `/* webpackIgnore: true */`，避免 webpack 试图打包它）。动态 `import()` 要求 tsconfig `module ≥ es2020`，故 `module` 由 `ES6` 提到 `ES2020`（项目此前无其它动态 import，emit 无差异）。
3. 不支持第三方包（`loadPackage` / `micropip` 均不接入），完全离线。
4. 超时 30 秒、「停止」＝ `worker.terminate()` 后按需重建（Pyodide 重新初始化需要重新加载 wasm，属可接受代价；停止时输出面板给出「正在加载 Python 运行环境…」的提示由下一次 run 触发）。
5. 脚本作用域＝**当前活动工作表**：快照取自活动工作表，回写也写回活动工作表；与既有「脚本按工作表名分键存储」一致。
6. 不做输出面板折叠 / 拖拽调高、不做 stdin 交互输入（`setStdin` 不接入）——避免过度设计。
7. 不改 `sheetNameSpace.ts` 的文件位置（用户已预留 `src/core/assistant/sheetNameSpace.ts`）。

## 6. 验证步骤（Verification）

1. **类型检查**：`npx tsc --noEmit` 退出码 0。
2. **构建产物**：`npm run build` 后确认：

   * `dist/python.worker.js` 存在（固定文件名，无 hash）；

   * `dist/pyodide/` 下恰好 5 个文件：`pyodide.mjs`、`pyodide.asm.mjs`、`pyodide.asm.wasm`、`python_stdlib.zip`、`pyodide-lock.json`（静态服务须把 `.mjs` 映射为 `text/javascript`、`.wasm` 映射为 `application/wasm`）；

   * `dist/index.html` 正常生成。
3. **浏览器实测**（**必须 http(s) 打开，不能用** **`file://`** —— `file://` 下 Worker 受同源策略禁止，本方案不提供主线程回退）：

   * 打开页面 → 切到「脚本」页签 → 点「运行」→ 默认脚本的 `Hello, SilkSpaces!` 实时出现在输出面板；「运行」变禁用、「停止」可用，结束后恢复。

   * 读：脚本写 `print(get_cell("A1"))` / `print(sheet_name, row_count, col_count)`，输出面板显示正确值。

   * 写回：脚本写 `set_cell("A1", "来自 Python")` 与 `set_range("B2:C3", [["b2","c2"],["b3","c3"]])` → 运行后表格对应格显示新文本且立即重绘；按 `Ctrl+Z` 能撤销这批写入，`Ctrl+Y` 能重做。

   * 覆盖写：先给 A1 手动输入内容，再用 `set_cell("A1","x")` 覆盖 → 撤销后恢复原内容。

   * 流式 + 停止：脚本写 `import time` + `for i in range(1000): print(i); time.sleep(0.01)` → 输出逐行实时追加；改用 `while True: pass` 后点「停止」→ 立即中断、按钮恢复、表格可继续操作；再跑一次仍正常（验证 Worker 重建）。

   * 超时：`while True: pass` 后不点停止，等 30 秒 → 自动停止并提示超时。

   * 异常：脚本写 `1/0` → 输出面板红色显示 traceback，按钮状态正常恢复。

   * 工作表隔离：切到另一工作表运行，脚本内容与输出互不影响；切回原表脚本内容仍在。

