/**
 * 输出面板的行缓冲（带上限裁剪）
 *
 * ## 为什么需要
 *
 * Python 脚本的 stdout 是按批流式回传的，一批可能是 1 行，也可能是几十万行
 * （`for i in range(1_000_000): print(i)` 就属于后者）。编辑器原先的实现是
 * 「一行一个 div 直接挂进输出面板」，没有任何上限 —— 百万行 `print` 就是百万个
 * DOM 节点，主线程内存与布局成本双双爆炸（浏览器直接 Out of Memory），
 * 而用户实际能看的只有面板里那十几行。
 *
 * 因此把「保留哪些行」这件纯逻辑单独抽到这里：只保留**最近的** {@link maxLines} 行，
 * 更早的行计入 {@link omitted}，由调用方在面板顶部显示「已省略 N 行」。
 *
 * ## 与调用方的契约
 *
 * {@link push} 返回的 `added` 是**缓冲区里现成的对象**（不是副本），调用方按序遍历
 * 追加节点即可；`removed` 是要从面板头部删掉的节点数。二者合起来保证
 * 「DOM 子节点（不含提示行）=== 缓冲区」这一不变式。
 *
 * @module outputBuffer
 */

/** 输出面板里的一行 */
export interface OutputLine {
    /** 行文本（不含换行符） */
    text: string;
    /** 是否为错误输出（stderr / traceback，红色显示） */
    isError: boolean;
}

/** {@link OutputBuffer.push} 的结果 */
export interface OutputPushResult {
    /** 本次追加后仍在缓冲区里的新行（缓冲区内的对象，按顺序） */
    added: OutputLine[];
    /**
     * 本次从**缓冲区头部**摘掉的行数
     *
     * 恰好等于调用方要从 DOM 头部删掉的节点数；批内被上限直接丢掉的行**不计入**
     * 这里（它们从未进过缓冲区、也从没进过 DOM），只进 {@link OutputBuffer.omitted}。
     */
    removed: number;
}

/**
 * 只保留最近 N 行的输出缓冲
 *
 * 不是环形队列：行数上限通常在千级，`splice` 的搬移成本可忽略，而数组顺序
 * 直接等于显示顺序，调用方不需要维护头尾指针。
 */
export class OutputBuffer {
    /** 当前保留的行（顺序 = 显示顺序） */
    private readonly lines: OutputLine[] = [];
    /** 累计被裁掉的行数（用于「已省略 N 行」提示） */
    private droppedCount = 0;
    /** 行数上限（至少 1） */
    private readonly maxLines: number;

    /**
     * @param maxLines - 最多保留多少行（<=0 时按 1 处理）
     */
    constructor(maxLines: number) {
        this.maxLines = maxLines > 0 ? Math.floor(maxLines) : 1;
    }

    /** 累计被裁掉的行数 */
    public get omitted(): number {
        return this.droppedCount;
    }

    /** 当前缓冲区里的行数（即调用方应有的 DOM 行节点数） */
    public get size(): number {
        return this.lines.length;
    }

    /**
     * 追加一段输出
     *
     * 换行语义与终端一致：末尾的换行不产生额外的空行（否则每次追加都凭空多一行），
     * 中间的空行则保留（否则脚本输出里的分段会被吃掉）。
     *
     * 【先截批次再建对象】一批可能远超上限，若先为每行建 `{text, isError}` 再裁，
     * 百万行批次等于先造百万个对象 —— 那正是要避免的开销。故先按 `parts.length`
     * 算出「本批只有最后 maxLines 行有机会留下」，从幸存区段开始建对象。
     *
     * 【为什么 `removed` 只会是历史行】批内截断保证「本批推进缓冲区的行数 ≤ maxLines」，
     * 于是二次裁剪要摘掉的行数 ≤ 追加前的行数 —— 被摘掉的一定是旧行，
     * `added` 因此就是缓冲区尾部 `pushed` 个对象。
     * @param text - 待追加的文本（可能含多行）
     * @param isError - 是否为错误输出
     * @returns {OutputPushResult} 需要追加到 DOM 的行与需要从头部摘掉的行数
     */
    public push(text: string, isError: boolean): OutputPushResult {
        if (!text) return { added: [], removed: 0 };

        const parts = text.split('\n');
        // 文本以换行结尾时 split 会多出一个空串，那只是行终止符，不是空行
        if (parts[parts.length - 1] === '') parts.pop();
        if (parts.length === 0) return { added: [], removed: 0 };

        // 本批就超限：只有最后 maxLines 行可能留下，前面的直接在字符串层丢掉
        const batchOverflow = parts.length > this.maxLines ? parts.length - this.maxLines : 0;
        const pushed = parts.length - batchOverflow;
        for (let i = batchOverflow; i < parts.length; i++) {
            this.lines.push({ text: parts[i], isError });
        }

        let removed = 0;
        if (this.lines.length > this.maxLines) {
            removed = this.lines.length - this.maxLines;
            this.lines.splice(0, removed);
        }
        // 省略计数 = 批内丢弃 + 被挤出缓冲区的历史行
        this.droppedCount += batchOverflow + removed;

        return { added: this.lines.slice(this.lines.length - pushed), removed };
    }

    /** 清空（连同「已省略」计数一起复位） */
    public clear(): void {
        this.lines.length = 0;
        this.droppedCount = 0;
    }
}
