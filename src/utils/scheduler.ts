/**
 * 分片调度工具：把长任务切成小片、片间让出主线程
 *
 * ## 为什么需要它（性能约定，勿改成「一口气跑完」）
 *
 * 打开大文件时主线程上原本有几段「一口气跑完」的长任务（映射为 Cell/Char、
 * 整表序列化、快照解码），单段动辄数百毫秒到 1.4 秒。这期间浏览器无法处理
 * 输入与渲染，用户看到的就是「卡死」。
 *
 * 把每段切成 {@link DEFAULT_SLICE_BUDGET_MS} 毫秒以内的小片、片间让出，
 * 总耗时几乎不变（让出本身的开销很小），但期间窗口可响应、可关闭、可取消。
 *
 * @module scheduler
 */

/** 单片最长占用主线程的时间（毫秒）。取得偏小以保证交互流畅 */
export const DEFAULT_SLICE_BUDGET_MS = 12;

/**
 * 每处理多少项才采样一次耗时
 *
 * 逐项调用 `Date.now()` 本身在大表上就是可观的成本（百万次），
 * 因此按固定间隔采样；间隔取 2 的幂，让判断退化为按位与。
 */
export const SLICE_CHECK_INTERVAL = 1024;

/** 复用的 MessageChannel（见 {@link yieldToUI}） */
let yieldChannel: MessageChannel | null = null;

/**
 * 等待「让出完成」的 resolver 队列（**必须排队，不能只存一个**）
 *
 * 历史实现把 `port1.onmessage` 直接指向本次调用的 `resolve`。只要有两处
 * 分片任务**并发**运行（例如自动保存的分片序列化与一次强制落库 / 切文件时
 * 的分片解码重叠），后一次调用就会覆盖前一次的 `onmessage`：
 * 前一个 `yieldToUI()` 的 promise **永远不会 settle**，那一段分片任务从此
 * 静默停住 —— 浏览器里表现为「这次自动保存再也没完成」，Node 里表现为
 * 进程无事可做而静默退出（基准脚本就踩到了这个坑）。
 *
 * 改为 FIFO 队列：每次 `postMessage` 送达时弹出最前面的 resolver，
 * 调用次数与送达次数一一对应，任意并发都各自推进。
 */
const yieldQueue: Array<() => void> = [];

/**
 * 取得（必要时创建）复用的让出通道
 *
 * `onmessage` 只装配一次：从队列头部取出 resolver 并调用。
 * @returns {MessageChannel} 通道
 */
function ensureYieldChannel(): MessageChannel {
    if (!yieldChannel) {
        yieldChannel = new MessageChannel();
        // 部分实现需要显式 start 才会派发消息
        (yieldChannel.port1 as unknown as { start?: () => void }).start?.();
        yieldChannel.port1.onmessage = () => {
            const resolve = yieldQueue.shift();
            if (resolve) resolve();
        };
    }
    return yieldChannel;
}

/**
 * 让出主线程一个宏任务，使浏览器有机会处理输入与渲染
 *
 * 实现上优先用 `MessageChannel`：`setTimeout(0)` 在浏览器里被钳到约 4ms，
 * 若让出几十次就会凭空多出上百毫秒；MessageChannel 的回调是普通宏任务，
 * 没有这个下限。环境不支持时（老浏览器 / Node）退回 `setTimeout`。
 *
 * **并发安全**：多处分片任务可以同时等待让出（各自入队、按序唤醒），
 * 不会互相吞掉 resolver（见 {@link yieldQueue}）。
 * @returns {Promise<void>} 让出后继续
 */
export function yieldToUI(): Promise<void> {
    return new Promise<void>((resolve) => {
        if (typeof MessageChannel === 'function') {
            const channel = ensureYieldChannel();
            yieldQueue.push(resolve);
            channel.port2.postMessage(0);
            return;
        }
        setTimeout(resolve, 0);
    });
}

/**
 * 分片游标：记录「本片从何时开始」，据此判断是否该让出
 *
 * 典型用法：
 * ```ts
 * const slice = new SliceBudget();
 * for (let i = 0; i < n; i++) {
 *     work(i);
 *     if (slice.shouldYield(i)) await slice.yieldNow();
 * }
 * ```
 */
export class SliceBudget {
    /** 本片起始时刻 */
    private started: number = Date.now();

    /** 重新开始计时（让出后、或进入新阶段时调用） */
    public restart(): void {
        this.started = Date.now();
    }

    /**
     * 本片时间是否已用尽（不做采样，每次都查时钟）
     *
     * 用于**循环体本身就是一大块工作**的场景（如每轮序列化 4000 个单元格 ≈ 5ms）：
     * 那里按间隔采样没有意义，直接每次查时钟即可。
     * @param {number} [budgetMs] - 本片预算；默认 {@link DEFAULT_SLICE_BUDGET_MS}
     * @returns {boolean} true = 该让出了
     */
    public expired(budgetMs: number = DEFAULT_SLICE_BUDGET_MS): boolean {
        return Date.now() - this.started >= budgetMs;
    }

    /**
     * 是否需要让出（按间隔采样，适用于逐项循环）
     * @param {number} index - 当前已处理的项下标（用于按间隔采样）
     * @param {number} [budgetMs] - 本片预算；默认 {@link DEFAULT_SLICE_BUDGET_MS}
     * @returns {boolean} true = 该让出了
     */
    public shouldYield(index: number, budgetMs: number = DEFAULT_SLICE_BUDGET_MS): boolean {
        if ((index & (SLICE_CHECK_INTERVAL - 1)) !== 0) return false;
        return Date.now() - this.started >= budgetMs;
    }

    /** 让出并重新计时 */
    public async yieldNow(): Promise<void> {
        await yieldToUI();
        this.restart();
    }
}
