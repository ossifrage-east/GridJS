/*  # Pyodide 执行客户端 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

/**
 * Python（Pyodide）执行客户端
 *
 * ## 与 {@link module:parserClient} 的关系：同构，但**不回退主线程**
 *
 * 两者都以「固定文件名的 Worker + id 序号请求/响应协议 + 惰性单例」为骨架
 * （见 `src/utils/parser/parserClient.ts`），差别只在降级策略：
 *
 * - 解析 Worker 在 Worker 不可用时回退主线程，因为「打开文件」是核心路径、
 *   且解压/扫描在主线程上跑出的结果完全一致，只是慢；
 * - Pyodide 的 CPython 是 WebAssembly 计算，**跑在主线程会冻结整个 UI**
 *   （WebAssembly 执行期间不释放事件循环，页面完全无响应）。因此这里**不做回退**：
 *   Worker 建不起来时直接把错误抛给调用方，由「脚本」页签的输出面板向用户说明原因。
 *
 * 最常见的建不起来的原因：以 `file://` 直接打开 `dist/index.html` —— Worker 脚本
 * 受同源策略限制，`file://` 页面构造 Worker 会直接抛 SecurityError。
 *
 * ## 流式输出
 *
 * 一次运行会在 Worker 侧产生**多条**消息（stdout / stderr 逐批回传），最后以
 * `result` 或 `error` 收尾。因此 {@link pending} 的条目除了 resolve/reject，
 * 还持有本次运行的流式回调，按 id 分发。
 *
 * ## 停止与超时
 *
 * `pyodide.runPython` 一旦进入死循环无法从外部打断（同一线程内没有抢占点），
 * 唯一可靠的打断方式是 `Worker.terminate()`。因此：
 * - 「停止」按钮 = {@link stopPython}（终止 Worker，下次运行重新初始化 Pyodide）；
 * - 超时保护默认 {@link PythonClient.TIMEOUT_MS}，超时同样走终止路径。
 *
 * ⚠️ 并发约定：同一时刻只允许一次运行。Worker 侧每次运行都会重绑 stdout/stderr，
 * 若两次运行重叠，输出会串到错误的请求 id 上。调用方（pyEditor）以
 * `isRunning` 标志保证串行，这里不重复加锁。
 *
 * @module pythonClient
 */

/** Worker 请求：主线程 → Worker */
export interface PythonRunRequest {
    /** 请求序号（同一次请求的流式消息与最终结果据此对齐） */
    id: number;
    /** 用户脚本全文 */
    script: string;
    /** 当前工作表名（注入 Python 命名空间的 `sheet_name`） */
    sheetName: string;
    /** 当前工作表快照（`{"A1": "文本", ...}` 的 JSON 字符串） */
    snapshotJson: string;
    /** 当前表格行数 */
    rowCount: number;
    /** 当前表格列数 */
    colCount: number;
    /**
     * `silkspaces` 命名空间的 Python 模块源码
     *
     * 随请求下发而不是让 Worker `import` 那个 TS 模块：`sheetNameSpace.ts` 在主线程侧
     * 依赖 `DataCollection` / `Cell`，若 Worker 静态引用它，整条数据层依赖链都会被
     * 打进 `python.worker.js`。
     */
    namespaceSource: string;
}

/** Worker 响应：Worker → 主线程 */
export type PythonWorkerMessage =
    | { id: number; type: 'stdout'; text: string }
    | { id: number; type: 'stderr'; text: string }
    | { id: number; type: 'result'; writesJson: string }
    | { id: number; type: 'error'; message: string };

/**
 * 「Worker 基础设施不可用」错误
 *
 * 仅用于区分「Worker 建不起来 / 通信断了」与「脚本自身出错」：
 * 前者需要用户检查部署方式（最常见是 `file://`），后者只需看输出面板里的 traceback。
 */
export class PythonWorkerUnavailableError extends Error {
    constructor(message: string) {
        super(message);
        this.name = 'PythonWorkerUnavailableError';
    }
}

/**
 * Worker 脚本路径
 *
 * 由 webpack 的独立 worker 配置产出为固定文件名（见 webpack.config.js），
 * 相对当前文档解析，故支持部署在任意子路径下。
 */
const WORKER_URL = 'python.worker.js';

/** 惰性单例 Worker */
let worker: Worker | null = null;
/** 已判定不可用（建不起来或通信中断）：后续请求直接失败，不反复尝试 */
let workerUnavailable = false;
/** 不可用的原因（仅用于诊断输出；可用时为 null） */
let unavailableReason: string | null = null;

/** 请求序号分配器 */
let nextRequestId = 1;

/** 在途请求（含本次运行的流式回调） */
const pending = new Map<number, {
    resolve: (writesJson: string) => void;
    reject: (error: Error) => void;
    onStdout?: (text: string) => void;
    onStderr?: (text: string) => void;
    /** 超时定时器句柄（结束时清除） */
    timer: ReturnType<typeof setTimeout>;
}>();

/**
 * 运行选项
 */
export interface PythonRunOptions {
    /** 用户脚本全文 */
    script: string;
    /** 当前工作表名 */
    sheetName: string;
    /** 当前工作表快照 JSON（见 sheetNameSpace.buildSnapshot） */
    snapshotJson: string;
    /** 当前表格行数 */
    rowCount: number;
    /** 当前表格列数 */
    colCount: number;
    /** `silkspaces` 命名空间模块源码（见 sheetNameSpace.PY_NAMESPACE_SOURCE） */
    namespaceSource: string;
    /** stdout 实时回调（按 Python 的批次，通常是整行） */
    onStdout?: (text: string) => void;
    /** stderr 实时回调（Python 的 traceback / warning 都走这里） */
    onStderr?: (text: string) => void;
}

/**
 * 「Pyodide 执行 Worker」的运行参数与「不可用」判定
 *
 * 提供这些常量与方法集中在此，避免超时值与诊断文案散落。
 */
export const PythonClient = {
    /** 单次运行的超时上限（毫秒）：超过即终止 Worker，避免死循环永久占住资源 */
    TIMEOUT_MS: 30000,
} as const;

/**
 * 当前是否已有可用的 Pyodide Worker
 *
 * 供诊断页面 / 自动化检查读取「Pyodide 是否真的跑在 Worker 中」。
 * @returns {boolean} true = Worker 可用；false = 不可用
 */
export function isPythonWorkerActive(): boolean {
    return worker !== null && !workerUnavailable;
}

/**
 * Worker 不可用的原因；可用时为 null
 * @returns {string | null} 原因描述
 */
export function getPythonWorkerUnavailableReason(): string | null {
    return unavailableReason;
}

/** 记录一次不可用（同一会话内只保留首个原因） */
function markUnavailable(reason: string): void {
    workerUnavailable = true;
    if (unavailableReason === null) unavailableReason = reason;
}

/** 让所有在途请求失败（基础设施故障或主动停止时调用） */
function failAllPending(error: Error): void {
    for (const [, entry] of pending) {
        clearTimeout(entry.timer);
        entry.reject(error);
    }
    pending.clear();
}

/**
 * 取得可用 Worker；不可用时返回 null（并记住，避免反复尝试）
 */
function getWorker(): Worker | null {
    if (workerUnavailable) return null;
    if (worker) return worker;

    // 非浏览器环境（如 Node 测试）：没有 Worker 构造器
    if (typeof Worker === 'undefined') {
        markUnavailable('当前环境没有 Worker 构造器（非浏览器环境）');
        return null;
    }

    try {
        // type: 'module' 是**硬性要求**，不是优化：Pyodide 314 在初始化时会用
        // `isClassicWorker()`（试调 `importScripts("data:text/javascript,")` 是否成功）
        // 判定当前 Worker 类型，classic worker 一律抛 `Classic web workers are not
        // supported`。module worker 里没有 `importScripts`，该判定自然为 false。
        // 因此 python.worker.js 也必须由 webpack 以 output.module 产出（见 webpack.config.js）。
        const w = new Worker(WORKER_URL, { type: 'module' });

        w.onmessage = (event: MessageEvent<PythonWorkerMessage>) => {
            const message = event.data;
            const entry = pending.get(message.id);
            if (!entry) return;
            switch (message.type) {
                case 'stdout':
                    entry.onStdout?.(message.text);
                    return;
                case 'stderr':
                    entry.onStderr?.(message.text);
                    return;
                case 'result':
                    pending.delete(message.id);
                    clearTimeout(entry.timer);
                    entry.resolve(message.writesJson);
                    return;
                case 'error':
                    pending.delete(message.id);
                    clearTimeout(entry.timer);
                    // 脚本自身出错：不回退、不重试，直接把信息交给调用方展示
                    entry.reject(new Error(message.message));
                    return;
            }
        };

        w.onerror = (event?: unknown) => {
            // 脚本 404 / 语法错误 / file:// 同源拒绝 / 运行时崩溃 → 视为基础设施不可用
            const detail = (event as { message?: string } | undefined)?.message;
            markUnavailable(`Worker 脚本加载或运行失败（${WORKER_URL}${detail ? '：' + detail : ''}）`);
            if (worker === w) worker = null;
            try { w.terminate(); } catch { /* 忽略 */ }
            failAllPending(new PythonWorkerUnavailableError(unavailableReason as string));
        };

        w.onmessageerror = () => {
            failAllPending(new PythonWorkerUnavailableError('Python Worker 消息反序列化失败'));
        };

        worker = w;
        return w;
    } catch (err) {
        const detail = err instanceof Error ? err.message : String(err);
        markUnavailable(`构造 Worker 失败（${detail}）`);
        return null;
    }
}

/**
 * 终止当前 Worker（「停止」按钮与超时保护共用）
 *
 * 无法从外部打断 Pyodide 里的死循环，只能连 Worker 一起丢弃；
 * 若之前判定为「不可用」则一并清掉该标记，使用户修正部署方式后（如改用 http 打开）
 * 下一次运行有机会重新建起 Worker。
 */
export function stopPython(): void {
    const w = worker;
    worker = null;
    workerUnavailable = false;
    unavailableReason = null;
    if (w) {
        try { w.terminate(); } catch { /* 忽略 */ }
    }
    failAllPending(new Error('已停止'));
}

/**
 * 在 Worker 中执行一段 Python 脚本
 *
 * @param options - 运行参数（脚本文本、工作表快照、流式回调等）
 * @returns 本次运行写回工作表的改动（`{"A1": "新值", ...}` 的 JSON 字符串）
 * @throws {PythonWorkerUnavailableError} Worker 基础设施不可用（调用方应提示用户检查部署方式）
 * @example
 * const { writesJson } = await runPython({ script, sheetName, snapshotJson, rowCount, colCount, namespaceSource });
 */
export function runPython(options: PythonRunOptions): Promise<{ writesJson: string }> {
    const w = getWorker();
    if (!w) {
        const reason = unavailableReason || 'Pyodide Worker 不可用';
        return Promise.reject(new PythonWorkerUnavailableError(
            `${reason}。请确认页面以 http(s) 打开（file:// 下 Worker 被同源策略禁止），`
            + `且 dist/python.worker.js 与 dist/pyodide/ 均已生成。`
        ));
    }

    const id = nextRequestId++;
    const request: PythonRunRequest = {
        id,
        script: options.script,
        sheetName: options.sheetName,
        snapshotJson: options.snapshotJson,
        rowCount: options.rowCount,
        colCount: options.colCount,
        namespaceSource: options.namespaceSource,
    };

    return new Promise<{ writesJson: string }>((resolve, reject) => {
        const timer = setTimeout(() => {
            // 超时：与「停止」同路（丢弃 Worker），并给出区别于用户主动停止的文案
            pending.delete(id);
            stopPython();
            reject(new Error(`Python 脚本执行超时（${PythonClient.TIMEOUT_MS / 1000} 秒），已强制停止`));
        }, PythonClient.TIMEOUT_MS);

        pending.set(id, {
            resolve: (writesJson: string) => resolve({ writesJson }),
            reject,
            onStdout: options.onStdout,
            onStderr: options.onStderr,
            timer,
        });

        try {
            w.postMessage(request);
        } catch (err) {
            pending.delete(id);
            clearTimeout(timer);
            const detail = err instanceof Error ? err.message : String(err);
            markUnavailable(`向 Worker 投递消息失败（${detail}）`);
            reject(new PythonWorkerUnavailableError(unavailableReason as string));
        }
    });
}
