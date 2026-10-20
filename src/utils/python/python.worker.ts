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
 * ## 数据传递一律用 JSON 字符串
 *
 * `pyodide.globals.set` 会把 JS 对象包成 `JsProxy`（需要 Python 侧 `.destroy()` 才不泄漏）。
 * 这里把「命名空间源码 + 用户脚本 + 快照」打包成一个 JSON 字符串一次性下发，
 * Python 侧 `json.loads` 展开即可 —— 既避开代理生命周期，也避开「把代码拼进
 * Python 字符串字面量」所需的层层转义。
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
 * 末尾把写入结果写进全局 `__RESULT__`，由 JS 侧 `globals.get('__RESULT__')` 取出。
 * 刻意用「赋值给全局再取」而不是「以表达式结尾靠返回值」：前者语义与
 * `runPython` 版本无关，更稳。
 *
 * 注入的全局变量只有 `__ARGS__`（JSON 字符串），字段见 {@link RunArgs}。
 */
const RUN_PROGRAM = [
    'import json, sys, types',
    '_scan = json.loads(__ARGS__)',
    // 1) 注册命名空间模块：exec 进一个独立模块字典，再挂到 sys.modules，
    //    用户脚本随后用 `from silkspaces import *` 就能取到 API
    '_mod = types.ModuleType("' + NAMESPACE_MODULE + '")',
    'exec(compile(_scan["namespace"], "' + NAMESPACE_MODULE + '", "exec"), _mod.__dict__)',
    'sys.modules["' + NAMESPACE_MODULE + '"] = _mod',
    // 2) 注入工作表快照与表信息（内部接口 _load 见 sheetNameSpace.PY_NAMESPACE_SOURCE）
    '_mod._load(_scan["snapshotJson"], _scan["sheet"], _scan["rows"], _scan["cols"])',
    'from ' + NAMESPACE_MODULE + ' import *',
    // 3) 执行用户脚本：独立字典 + __name__='__main__'，使脚本末尾的
    //    `if __name__ == "__main__": main()` 能按预期触发。
    //    脚本先经 _mod._compile 改写：`A1 = 1` / `x = A1` 这类标识符会被换成
    //    对 __cells__ 的下标访问，于是模块级与 def 函数体内都能写 / 读单元格
    //    （函数体内的赋值本是 STORE_FAST，写进帧的局部数组，任何 globals 钩子
    //    都拦不到 —— 这也是原先只有模块级赋值生效的原因）。
    //    __cells__ 是模块内部名（下划线名不会被 `import *` 带出），故显式注入。
    '_user = dict(globals())',
    '_user["__name__"] = "__main__"',
    '_user["__cells__"] = _mod.__cells__',
    'exec(_mod._compile(_scan["script"], "<script>"), _user)',
    // 4) 取回写入改动
    '__RESULT__ = _mod._dump()',
].join('\n');

/** 注入给 {@link RUN_PROGRAM} 的 `__ARGS__` 字段 */
interface RunArgs {
    /** silkspaces 命名空间模块源码 */
    namespace: string;
    /** 用户脚本全文 */
    script: string;
    /** 工作表快照 JSON（`{"A1": "文本", ...}`） */
    snapshotJson: string;
    /** 当前工作表名 */
    sheet: string;
    /** 表格行数 */
    rows: number;
    /** 表格列数 */
    cols: number;
}

/**
 * 执行一次用户脚本
 * @param pyodide - 已就绪的运行时
 * @param request - 主线程下发的运行请求
 * @returns {Promise<string>} 写回改动的 JSON 字符串
 */
async function runScript(pyodide: PyodideAPI, request: PythonRunRequest): Promise<string> {
    // 输出绑定按当前请求 id 分流：流式回传的多条消息据此归到正确的运行
    pyodide.setStdout({ batched: (text: string) => ctx.postMessage({ id: request.id, type: 'stdout', text }) });
    pyodide.setStderr({ batched: (text: string) => ctx.postMessage({ id: request.id, type: 'stderr', text }) });

    const args: RunArgs = {
        namespace: request.namespaceSource,
        script: request.script,
        snapshotJson: request.snapshotJson,
        sheet: request.sheetName,
        rows: request.rowCount,
        cols: request.colCount,
    };
    pyodide.globals.set('__ARGS__', JSON.stringify(args));

    await pyodide.runPythonAsync(RUN_PROGRAM);
    return String(pyodide.globals.get('__RESULT__'));
}

/**
 * 把任意异常整理成可读文本
 *
 * Python 异常（`PythonError`）的 `message` 里已经带有完整 traceback，
 * 直接原样交给输出面板即可；其余情况退化为 `String(err)`。
 * @param err - 抛出的异常
 * @returns {string} 错误文本
 */
function errorText(err: unknown): string {
    if (err instanceof Error && err.message) return err.message;
    return String(err);
}

ctx.onmessage = (event: MessageEvent<PythonRunRequest>) => {
    const request = event.data;

    void (async () => {
        try {
            const pyodide = await getRuntime();
            const writesJson = await runScript(pyodide, request);
            ctx.postMessage({ id: request.id, type: 'result', writesJson });
        } catch (err) {
            ctx.postMessage({ id: request.id, type: 'error', message: errorText(err) });
        }
    })();
};
