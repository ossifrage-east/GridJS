/**
 * Excel 解析客户端封装（Worker 优先，主线程回退）
 *
 * ## 为什么要分线程（性能约定，请勿回退为「一律主线程」）
 *
 * `SheetParser` 的前两步（`decode` 解压 + `extract` 流式扫描 XML）是**纯计算**：
 * 大文件下 ZIP 解压与数百 MB XML 的逐字节扫描会长时间占满主线程，
 * 期间页面无法响应任何交互（表现为「打开大文件即卡死」）。
 *
 * 把这两步放进 Web Worker 后：
 * - 主线程只保留第三步 `generate`（映射为 Cell/Char）与后续渲染；
 * - 传输走 **ArrayBuffer 转移**（零拷贝）：文件字节 `transfer` 给 Worker，
 *   解析出的工作表结构 `transfer` 回来（图片二进制随之转移）；
 * - Worker 不可用（老浏览器 / 构建产物缺失 / 运行在 Node 测试环境）时
 *   **自动回退主线程**，行为与优化前完全一致，不会让「打开文件」这一核心路径失效。
 *
 * ## 为什么 Worker 只做 decode+extract+pack，generate 留在主线程
 *
 * `generate` 产出的是 `Cell` / `Char` **类实例**，跨线程传递会丢失原型
 * （`getText()` 等方法不再可用），因此固定在主线程执行。
 *
 * 但"把 `ExtractedSheet[]` 原样传回来"曾是一个**被低估的瓶颈**：结构化克隆要逐对象
 * 重建，实测 6.3 万格（6.2 万格带样式）在主线程反序列化要 489 ms —— 比它取代的
 * `extract`（275 ms）更贵。因此 Worker 侧先做**列式打包**（见 `cellSource.ts`），
 * 回传的是 typed array + 样式去重表 + 文本池，克隆成本降到 48 ms。
 *
 * @module parserClient
 */

import {
    SheetParser,
    SheetParserOptions,
    SheetParserResult,
} from './sheetParser';
import type { PackedSheet } from './cellSource';

/** Worker 请求：主线程 → Worker */
export interface ParserWorkerRequest {
    /** 请求序号（同一次请求的响应对齐用） */
    id: number;
    /** 文件名（解析器据扩展名分派策略） */
    name: string;
    /** 文件字节（转移所有权，零拷贝） */
    buffer: ArrayBuffer;
    /** 解析选项 */
    options?: SheetParserOptions;
}

/** Worker 响应：Worker → 主线程 */
export interface ParserWorkerResponse {
    /** 与请求对应的序号 */
    id: number;
    /** 成功：**列式打包**的工作表（typed array 已通过 transfer 零拷贝回传） */
    sheets?: PackedSheet[];
    /** 失败：错误信息 */
    error?: string;
}

/**
 * 「Worker 基础设施不可用」错误
 *
 * 仅用于区分「Worker 建不起来 / 通信断了」（应回退主线程）与
 * 「文件本身解析失败」（应直接把错误抛给用户，不回退、不重复解析）。
 */
export class ParserWorkerUnavailableError extends Error {
    constructor(message: string) {
        super(message);
        this.name = 'ParserWorkerUnavailableError';
    }
}

/**
 * Worker 脚本路径
 *
 * 由 webpack 的独立 worker 配置产出为固定文件名（见 webpack.config.js），
 * 相对当前文档解析，故支持部署在任意子路径下。
 */
const WORKER_URL = 'parser.worker.js';

/** 惰性单例 Worker */
let worker: Worker | null = null;
/** 已判定不可用（建不起来或通信中断）：后续请求直接走主线程 */
let workerUnavailable = false;
/**
 * 回退主线程的原因（仅用于诊断输出；未回退时为 null）
 *
 * 之所以要把它记下来并打印：回退是**静默**的，一旦 Worker 建不起来
 * （最常见的是用 `file://` 直接打开 `dist/index.html` —— Worker 脚本受同源策略
 * 限制，`file://` 页面构造 Worker 会直接抛 SecurityError），
 * `decode + extract` 就仍在主线程跑，表现就是「优化做了但打开大文件还是卡」。
 * 没有日志时这个状态完全不可见，是排查时最费时的一种「假象」。
 */
let fallbackReason: string | null = null;

/** 请求序号分配器 */
let nextRequestId = 1;
/** 在途请求 */
const pending = new Map<number, {
    resolve: (sheets: PackedSheet[]) => void;
    reject: (error: Error) => void;
}>();

/** 打印一次回退告警（同一进程内只打一次，避免刷屏） */
function warnFallbackOnce(reason: string): void {
    if (fallbackReason !== null) return;
    fallbackReason = reason;
    // 面向开发者：这不是业务错误，而是「性能优化未生效」的信号
    console.warn(
        `[parserClient] 解析未使用 Web Worker，已回退到主线程（原因：${reason}）。`
        + `\n  影响：解压与 XML 扫描会在主线程执行，打开大文件时页面可能卡顿/无响应。`
        + `\n  请检查：1) 是否以 http(s) 而非 file:// 打开页面（file:// 下 Worker 被同源策略禁止）；`
        + `2) dist/parser.worker.js 是否存在且可访问；3) 浏览器控制台是否有脚本加载错误。`
    );
}

/**
 * 当前解析是否真正运行在 Web Worker 中
 *
 * 供诊断页面 / 自动化检查读取「优化是否生效」，避免只能靠肉眼观察卡顿来猜。
 * @returns {boolean} true = 使用 Worker；false = 已回退主线程
 */
export function isParserWorkerActive(): boolean {
    return worker !== null && !workerUnavailable;
}

/**
 * 回退主线程的原因；未回退时为 null
 * @returns {string | null} 原因描述
 */
export function getParserWorkerFallbackReason(): string | null {
    return fallbackReason;
}

/** 让所有在途请求以「基础设施不可用」失败，触发各自回退 */
function failAllPending(reason: string): void {
    const error = new ParserWorkerUnavailableError(reason);
    for (const [, entry] of pending) entry.reject(error);
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
        workerUnavailable = true;
        warnFallbackOnce('当前环境没有 Worker 构造器（非浏览器环境）');
        return null;
    }

    try {
        const w = new Worker(WORKER_URL);

        w.onmessage = (event: MessageEvent<ParserWorkerResponse>) => {
            const message = event.data;
            const entry = pending.get(message.id);
            if (!entry) return;
            pending.delete(message.id);
            if (message.error !== undefined) {
                // 解析失败：不回退（回退也会同样失败，只会白跑一遍并掩盖真实原因）
                entry.reject(new Error(message.error));
            } else {
                entry.resolve(message.sheets || []);
            }
        };

        w.onerror = (event?: unknown) => {
            // 脚本 404 / 语法错误 / file:// 同源拒绝 / 运行时崩溃 → 视为基础设施不可用
            const detail = (event as { message?: string } | undefined)?.message;
            workerUnavailable = true;
            if (worker === w) worker = null;
            try { w.terminate(); } catch { /* 忽略 */ }
            warnFallbackOnce(`Worker 脚本加载或运行失败（${WORKER_URL}${detail ? '：' + detail : ''}）`);
            failAllPending('解析 Worker 发生错误');
        };

        w.onmessageerror = () => {
            warnFallbackOnce('Worker 消息反序列化失败');
            failAllPending('解析 Worker 消息反序列化失败');
        };

        worker = w;
        return w;
    } catch (err) {
        workerUnavailable = true;
        const detail = err instanceof Error ? err.message : String(err);
        warnFallbackOnce(`构造 Worker 失败（${detail}）`);
        return null;
    }
}

/**
 * 在 Worker 中执行 decode + extract + 列式打包
 *
 * @param w - 可用的 Worker
 * @param file - 待解析文件
 * @param options - 解析选项
 * @returns 列式打包的工作表（typed array 已零拷贝回传）
 * @throws {ParserWorkerUnavailableError} 基础设施失败（调用方应回退主线程）
 */
function extractInWorker(
    w: Worker,
    file: File,
    options?: SheetParserOptions
): Promise<PackedSheet[]> {
    return new Promise<PackedSheet[]>((resolve, reject) => {
        const id = nextRequestId++;
        pending.set(id, { resolve, reject });

        file.arrayBuffer().then((buffer) => {
            try {
                const request: ParserWorkerRequest = { id, name: file.name, buffer, options };
                // 转移 buffer 所有权：Worker 侧拿到同一块内存，主线程零拷贝
                w.postMessage(request, [buffer]);
            } catch (err) {
                pending.delete(id);
                reject(new ParserWorkerUnavailableError(
                    err instanceof Error ? err.message : String(err)
                ));
            }
        }).catch((err) => {
            pending.delete(id);
            reject(new ParserWorkerUnavailableError(
                err instanceof Error ? err.message : String(err)
            ));
        });
    });
}

/**
 * 解析 Excel 文件为项目数据格式（Worker 优先，主线程回退）
 *
 * 与 `new SheetParser().parse(file, options)` 的返回值完全一致，
 * 区别只在于「解压 + 提取 + 打包」三步被搬到了 Worker。
 *
 * @param file - 待解析的 .xlsx / .xls 文件
 * @param options - 解析选项
 * @returns 与项目格式一致的数据
 * @example
 * const result = await parseWorkbook(file);
 */
export async function parseWorkbook(
    file: File,
    options?: SheetParserOptions
): Promise<SheetParserResult> {
    const parser = new SheetParser();
    const w = getWorker();

    if (w) {
        try {
            const packed = await extractInWorker(w, file, options);
            // generate 产出 Cell/Char 类实例，必须在主线程完成（跨线程会丢原型），
            // 但输入是列式打包数据，反序列化代价比原样回传低一个数量级
            return await parser.generatePacked(packed, options);
        } catch (err) {
            if (!(err instanceof ParserWorkerUnavailableError)) throw err;
            // 基础设施不可用 → 落到下面的主线程路径
            warnFallbackOnce(err.message);
        }
    }

    return parser.parse(file, options);
}
