/*  # Pyodide 执行 Worker TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

/**
 * Pyodide 执行 Worker 入口
 *
 * ## 职责
 *
 * 1. 按 {@link PYODIDE_DIR} 加载 Pyodide 运行时（运行时动态 `import()` 其 ESM 入口
 *    `pyodide.mjs`）——**惰性**：第一次收到运行请求才加载，避免页面一打开就吃掉
 *    十几 MB 的下载与 wasm 编译；
 * 2. 把主线程下发的 `silkspaces` 命名空间源码注册成真正的 Python 模块；
 * 3. 注入工作表快照 → 执行用户脚本 → 取回写入改动；
 * 4. stdout / stderr 逐批**流式**回传（`setStdout({ batched })`），而不是等运行结束
 *    一次性返回：长脚本运行期间用户要能看到进度，否则界面看起来像卡死了。
 *
 * ## 为什么不把运行时打包进来
 *
 * Pyodide 的 wasm（约 10 MB）与 stdlib zip 由 Pyodide 自身按 `indexURL` 在运行时
 * 用 `fetch` / 动态 `import()` 拉取，webpack 无法也不应把它们当模块处理。
 * 因此它们是「静态资产」：构建时由 copy-webpack-plugin 拷到 `dist/pyodide/`
 * （见 webpack.config.js），运行时按下面的 {@link PYODIDE_DIR} 定位。
 *
 * ## indexURL 必须显式传入
 *
 * Pyodide 在拿不到 indexURL 时会尝试从「模块自身的位置」反推，而 Worker 环境里
 * 反推会失败并抛 `Could not extract indexURL path from pyodide module location`。
 * 这里用 **Worker 脚本自身的 URL** 作为基准算绝对地址，因此部署在任意子路径下都成立。
 *
 * ## 数据逐字段注入（不整包 JSON）

 * 早期版本把「命名空间源码 + 用户脚本 + 快照」`JSON.stringify` 成一整包
 * `__ARGS__` 下发、Python 侧 `json.loads` 展开 —— 那会在 worker 的 JS 堆里
 * 凭空多出一份几十 MB 的拼接串，且 Python 侧「原串 + 解析结果」长期同时驻留。
 * 现在 {@link runScript} 逐字段 `globals.set`：Pyodide 会把 JS 原始类型
 * （string / number）转换成对应的 Python 值（不产生 JsProxy），worker 侧零额外拷贝；
 * Python 侧（见 {@link RUN_PROGRAM}）用完一个字段就 `del` 一个，巨型快照原串
 * 不会陪你跑完整个脚本。仍以 JSON 字符串形态传递**格子载荷**（snapshotJson 等），
 * 由命名空间 `_load` 自行解析 —— 避开「把代码拼进 Python 字符串字面量」的转义问题。
 *
 * ## 无法从内部打断
 *
 * 用户脚本一旦进入死循环，同线程内没有任何抢占点，`runPython` 不会返回。
 * 打断只能靠外部 `Worker.terminate()`（见 `pythonClient.stopPython`），
 * 本文件因此**不实现**任何超时/中断逻辑。
 *
 * ## 并发约定
 *
 * 同一时刻只处理一次运行：stdout/stderr 的绑定是「按当前请求 id」的，
 * 两次运行重叠会让输出串到错误的请求上。调用方（pyEditor）以 `isRunning` 保证串行。
 *
 * @module python.worker
 */

import type { PyodideAPI } from 'pyodide';
import type { PythonRunRequest, PythonWorkerMessage } from './pythonClient';

/**
 * Worker 全局作用域的最小类型视图
 *
 * 不引入 `WebWorker` lib（会与项目的 DOM lib 冲突），仅声明实际用到的成员。
 */
const ctx = self as unknown as {
    onmessage: ((event: MessageEvent<PythonRunRequest>) => void) | null;
    postMessage: (message: PythonWorkerMessage) => void;
};

/**
 * Pyodide 资产目录（绝对 URL，以 Worker 脚本自身位置为基准）
 *
 * 例如 `https://host/dist/python.worker.js` → `https://host/dist/pyodide/`。
 */
const PYODIDE_DIR = new URL('./pyodide/', self.location.href).href;

/** `silkspaces` 命名空间在 Python 侧的模块名（与 sheetNameSpace.ts 的源码约定一致） */
const NAMESPACE_MODULE = 'silkspaces';

/** 惰性加载的运行时 Promise（失败时置回 null，允许下次重试） */
let runtimePromise: Promise<PyodideAPI> | null = null;

/**
 * 加载 Pyodide 运行时（运行时动态 `import()` 其 ESM 入口 `pyodide.mjs`）
 *
 * **必须用 ESM 入口**：Pyodide 314 初始化时会做 `isClassicWorker()` 检测 ——
 * 试调 `importScripts("data:text/javascript,")`，成功即判定当前是 classic worker，
 * 随即抛 `Classic web workers are not supported`（与「是否自己调用过
 * `importScripts`」无关：classic worker 里该函数就是存在的）。因此：
 *   1. 本 Worker 必须由 `new Worker(url, { type: 'module' })` 构造（见 pythonClient），
 *      并由 webpack 以 `output.module` 产出 ESM；
 *   2. 运行时入口取 `pyodide.mjs`（ESM）而非 `pyodide.js`（经典脚本，靠
 *      `importScripts` 往全局注入 `loadPyodide`，在 module worker 里无从使用）。
 * 两者是同一套实现的两种打包形态，行为一致。
 *
 * `/* webpackIgnore: true *​/` 必须保留：URL 是运行时算出来的，webpack 既无法
 * 静态解析，也不该把 `pyodide.mjs`（及其后续 `pyodide.asm.mjs`）打进 bundle。
 * @returns {Promise<PyodideAPI>} 运行时 API 对象
 */
async function loadPyodideRuntime(): Promise<PyodideAPI> {
    const mod = (await import(/* webpackIgnore: true */ new URL('pyodide.mjs', PYODIDE_DIR).href)) as {
        loadPyodide: (options?: { indexURL?: string }) => Promise<PyodideAPI>;
    };
    // indexURL 必须显式传入（见模块头注释）
    return mod.loadPyodide({ indexURL: PYODIDE_DIR });
}

/**
 * 取得（必要时加载）Pyodide 运行时
 * @returns {Promise<PyodideAPI>} 运行时 API 对象
 */
function getRuntime(): Promise<PyodideAPI> {
    if (!runtimePromise) {
        runtimePromise = loadPyodideRuntime().catch((err: unknown) => {
            runtimePromise = null;
            throw err;
        });
    }
    return runtimePromise;
}

/**
 * 一次运行对应的 Python 程序（固定部分）
 *
 * 末尾把写入结果（`_dump_chunks` 的分片数组）写进全局 `__RESULT__`，
 * 由 JS 侧 `globals.get('__RESULT__')` 取出。刻意用「赋值给全局再取」而不是
 * 「以表达式结尾靠返回值」：前者语义与 `runPython` 版本无关，更稳。
 *
 * 注入的全局变量逐字段命名（见 {@link runScript}）：快照与样式是**分片**下发的
 * （`__SNAP_N__` 计数 + `__SNAP_C0__…` 各片、`__STYLES_N__` / `__STYLES_C0__…`），
 * 这里逐片 `json.loads` 并进结果字典后立刻 `del` 片串 —— 任何时刻堆里最多一片；
 * 剩余字段用完即 `del`，且这些 `del` 都排在 `_user = dict(globals())` 之前 ——
 * 否则用户脚本的 globals 快照会把巨型载荷一起抄走。
 * 字段对应关系：`__NS_ARGS__`（命名空间源码，明文）→ `__SCRIPT__`（用户脚本，明文）→
 * `__SNAP_*` / `__STYLES_*`（分片 JSON 载荷）→ `__OPENED_JSON__` / `__META_JSON__`
 * （JSON 载荷，体量受 preloadAll 预算约束，保持单串）→ `__SHEET_NAME__` / `__ROWS__` / `__COLS__`（标量）。
 */
const RUN_PROGRAM = [
    'import json, sys, types',
    // 0) 命名空间源码是明文（不经 JSON），取到即摘
    '_ns_src = __NS_ARGS__',
    'del __NS_ARGS__',
    // 1) 注册命名空间模块：exec 进一个独立模块字典，再挂到 sys.modules，
    //    用户脚本随后用 `from silkspaces import *` 就能取到 API
    '_mod = types.ModuleType("' + NAMESPACE_MODULE + '")',
    'exec(compile(_ns_src, "' + NAMESPACE_MODULE + '", "exec"), _mod.__dict__)',
    'del _ns_src',
    'sys.modules["' + NAMESPACE_MODULE + '"] = _mod',
    // 2) 逐片解析快照与样式：每片到手即 update 进结果字典、立刻 del 片串，
    //    任何时刻堆里最多一片 —— 单条巨型 JSON 串的常驻峰值就此消失。
    //    _load 直接收已解析的字典（见 sheetNameSpace.PY_NAMESPACE_SOURCE 的 _load）。
    '_snapshot = {}',
    '_styles = {}',
    '_g = globals()',
    'for _i in range(_g["__SNAP_N__"]):',
    '    _snapshot.update(json.loads(_g["__SNAP_C%d__" % _i]))',
    '    del _g["__SNAP_C%d__" % _i]',
    'del _g["__SNAP_N__"]',
    'for _i in range(_g["__STYLES_N__"]):',
    '    _styles.update(json.loads(_g["__STYLES_C%d__" % _i]))',
    '    del _g["__STYLES_C%d__" % _i]',
    'del _g["__STYLES_N__"]',
    'del _g',
    '_mod._load(_snapshot, _styles, __OPENED_JSON__, __SHEET_NAME__, __ROWS__, __COLS__, __META_JSON__)',
    //    _load 已把载荷装进命名空间 —— 字典与原串立刻释放，后面几十秒的脚本执行期间不留副本
    'del _snapshot, _styles, __OPENED_JSON__, __META_JSON__, __SHEET_NAME__, __ROWS__, __COLS__',
    'from ' + NAMESPACE_MODULE + ' import *',
    // 3) 执行用户脚本：独立字典 + __name__='__main__'，使脚本末尾的
    //    `if __name__ == "__main__": main()` 能按预期触发。
    //    脚本先经 _mod._compile 改写：`A1 = 1` / `x = A1` 这类标识符会被换成
    //    对 __cells__ 的下标访问，于是模块级与 def 函数体内都能写 / 读单元格
    //    （函数体内的赋值本是 STORE_FAST，写进帧的局部数组，任何 globals 钩子
    //    都拦不到 —— 这也是原先只有模块级赋值生效的原因）。
    //    __cells__ 是模块内部名（下划线名不会被 `import *` 带出），故显式注入。
    //    注意顺序：到这里巨型字段已全部 del，_user 抄不到它们。
    '_user = dict(globals())',
    '_user["__name__"] = "__main__"',
    '_user["__cells__"] = _mod.__cells__',
    'exec(_mod._compile(__SCRIPT__, "<script>"), _user)',
    'del __SCRIPT__',
    // 4) 取回写入改动：分片数组（每片 ≤ 4000 格），JSON 序列化成一条后由 JS 侧拆开
    '__RESULT__ = json.dumps(_mod._dump_chunks())',
].join('\n');

/**
 * 把分片载荷逐片注入为独立的全局变量
 *
 * 命名规则：`<prefix>N__` = 片数，`<prefix>C0__`、`<prefix>C1__`… = 各片 JSON 串。
 * 刻意不传 JS 数组（那会包成 JsProxy，还得管 destroy 生命周期）：
 * 纯字符串全局量与 RUN_PROGRAM 的「逐片 loads → update → del」契约直接对上，
 * 也让整条链路可以用真实 CPython 做端到端测试。
 * @param globals - Pyodide 全局命名空间
 * @param prefix - 变量名前缀（如 `'__SNAP_'`）
 * @param chunks - JSON 片数组（空数组 → N__ = 0，Python 侧按空字典处理）
 */
function setChunked(globals: PyodideAPI['globals'], prefix: string, chunks: string[]): void {
    globals.set(prefix + 'N__', chunks.length);
    for (let i = 0; i < chunks.length; i++) globals.set(prefix + 'C' + i + '__', chunks[i]);
}

/**
 * 执行一次用户脚本
 * @param pyodide - 已就绪的运行时
 * @param request - 主线程下发的运行请求
 * @returns {Promise<string>} `_dump_chunks()` 分片数组的 JSON 序列化
 */
async function runScript(pyodide: PyodideAPI, request: PythonRunRequest): Promise<string> {
    // 输出绑定按当前请求 id 分流：流式回传的多条消息据此归到正确的运行
    pyodide.setStdout({ batched: (text: string) => ctx.postMessage({ id: request.id, type: 'stdout', text }) });
    pyodide.setStderr({ batched: (text: string) => ctx.postMessage({ id: request.id, type: 'stderr', text }) });

    // 逐字段注入（字段对应关系见 RUN_PROGRAM）：Pyodide 把 JS 的 string / number
    // 原生转换成 Python 值，不产生 JsProxy；对比旧的「整包 JSON.stringify 成 __ARGS__」，
    // worker 的 JS 堆里不再多拼一份几十 MB 的串。
    const g = pyodide.globals;
    g.set('__NS_ARGS__', request.namespaceSource);
    g.set('__SCRIPT__', request.script);
    setChunked(g, '__SNAP_', request.snapshotChunks);
    setChunked(g, '__STYLES_', request.stylesChunks);
    g.set('__OPENED_JSON__', request.openedJson);
    g.set('__META_JSON__', request.metaJson);
    g.set('__SHEET_NAME__', request.sheetName);
    g.set('__ROWS__', request.rowCount);
    g.set('__COLS__', request.colCount);

    // 注入完成后消息体就没用了（后面只读 request.id）：立刻摘掉巨型字段。
    // 不摘的话，postMessage 时 structured clone 出来的这份载荷会以「事件数据」
    // 的身份在 worker 堆里常驻到整次运行结束。
    const mutable = request as unknown as Record<string, unknown>;
    mutable.namespaceSource = undefined;
    mutable.script = undefined;
    mutable.snapshotChunks = undefined;
    mutable.stylesChunks = undefined;
    mutable.openedJson = undefined;
    mutable.metaJson = undefined;

    await pyodide.runPythonAsync(RUN_PROGRAM);
    return String(pyodide.globals.get('__RESULT__'));
}

/**
 * 把任意异常整理成可读文本
 *
 * Python 异常（`PythonError`）的 `message` 里已经带有完整 traceback，
 * 直接原样交给输出面板即可；其余情况退化为 `String(err)`。
 * 例外：MemoryError 对用户是黑话 —— traceback 满屏看不出重点，
 * 单独给一句「表格过大 / 分批处理」的人话提示。
 * @param err - 抛出的异常
 * @returns {string} 错误文本
 */
function errorText(err: unknown): string {
    if (err instanceof Error && err.message) {
        if (err.message.indexOf('MemoryError') >= 0) {
            return err.message + '\n（Python 运行内存不足：脚本一次操作的表格太大。'
                + '可改为按行/按块分批处理，或减少本次打开的表格数量后重试。）';
        }
        return err.message;
    }
    return String(err);
}

ctx.onmessage = (event: MessageEvent<PythonRunRequest>) => {
    const request = event.data;

    void (async () => {
        try {
            const pyodide = await getRuntime();
            // __RESULT__ 是 _dump_chunks() 分片数组的 JSON 序列化 —— 这里拆成真正的
            // 片数组再 postMessage：主线程与 worker 都不再出现单条巨型 writesJson
            const resultJson = await runScript(pyodide, request);
            const writesChunks = JSON.parse(resultJson) as string[];
            ctx.postMessage({ id: request.id, type: 'result', writesChunks });
        } catch (err) {
            ctx.postMessage({ id: request.id, type: 'error', message: errorText(err) });
        }
    })();
};
