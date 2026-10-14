/**
 * `generate` 的数据源抽象：数组式（提取结果原样）与列式打包式（跨线程回传）
 *
 * ## 为什么需要这一层（性能约定，请勿把两者合并回「只认 ExtractedSheet[]」）
 *
 * `generate` 的输入有两种来源：
 *
 * 1. **数组式** —— 主线程回退路径（Worker 不可用）与 `xls` 直连路径：
 *    `extract()` 就地产出 `ExtractedCell[]`（每格一个普通对象，样式是嵌套对象）；
 * 2. **列式打包式** —— Worker 回传路径：数据先 {@link packSheets} 成「平行数组 +
 *    样式表去重 + 文本池」，typed array 走 `transfer` 零拷贝回主线程。
 *
 * 第 2 条是**打开大文件不再卡死的关键**，原因是结构化克隆的代价：
 * 实测一份 648 KB 的 xlsx（6.3 万格，其中 6.2 万格带样式）——
 *
 * | 回传形态 | 主线程反序列化耗时 | 常驻内存增量 |
 * | --- | --- | --- |
 * | `ExtractedCell[]` 原样 | **489 ms** | 58.3 MB |
 * | 列式打包 + transfer | **48 ms** | 11.6 MB |
 *
 * 差距来自 V8 的结构化克隆要**逐对象**重建（每个普通对象 + 每个嵌套样式对象
 * 都是一次分配 + 一次属性表遍历），而 typed array 只是 `memcpy`。
 * 原形态下这 489 ms 是「打开文件」时主线程上最长的一段同步阻塞 —— 比它要取代的
 * `extract`（275 ms）还贵，也比 `generate`（39 ms）贵一个数量级。
 *
 * 因此：**Worker 侧一律打包，主线程只解包**。而两种来源在 `generate` 里共用
 * 同一份映射代码（{@link SheetSource} 接口），不会出现「两条路径行为不一致」。
 *
 * @module cellSource
 */

import type { CellStyleInfo, ExtractedImage } from './sheetParser';

/** 合并区域（0-based，闭区间） */
export interface ExtractedMerge {
    s: { r: number; c: number };
    e: { r: number; c: number };
}
/** 列宽（0-based 索引，像素） */
export interface ExtractedCol {
    index: number;
    width?: number;
    customWidth?: boolean;
}
/** 行高（0-based 索引，像素） */
export interface ExtractedRow {
    index: number;
    height?: number;
    customHeight?: boolean;
}
/** 数据范围（0-based，闭区间） */
export interface ExtractedRange {
    minRow: number;
    maxRow: number;
    minCol: number;
    maxCol: number;
}

/** 提取阶段的单元格（纯数据，`ExtractedCell` 的结构视图） */
export interface ExtractedCellLike {
    /** 0-based 行号 */
    row: number;
    /** 0-based 列号 */
    col: number;
    /** 单元格值 */
    value: string | number | boolean | null;
    /** 样式信息 */
    style?: CellStyleInfo;
    /**
     * 样式内部编号（提取阶段按编号内联缓存，同编号必同对象）
     *
     * 有了它，{@link packSheet} 的样式去重就是一次数字查表 —— 不必对每格
     * `JSON.stringify(style)`。缺失时打包侧退回按内容比较。
     */
    styleId?: number;
    /** 数字格式码 */
    numberFormat?: string;
}

/** 提取阶段的工作表（纯数据，`ExtractedSheet` 的结构视图） */
export interface ExtractedSheetLike {
    /** 工作表名称 */
    name: string;
    /** 所有单元格 */
    cells: ExtractedCellLike[];
    /** 合并区域 */
    merges: ExtractedMerge[];
    /** 列宽信息 */
    cols: ExtractedCol[];
    /** 行高信息 */
    rows: ExtractedRow[];
    /** 图片 */
    images: ExtractedImage[];
    /** 数据范围 */
    range: ExtractedRange;
}

/**
 * `generate` 所能接受的「一个工作表」的最小视图
 *
 * 刻意保持**基础类型**（number/string/undefined），不带 `Cell`/`Char` 实例：
 * 这样两种来源都能实现它，而 `generate` 只需一份映射逻辑。
 */
export interface SheetSource {
    /** 单元格数量 */
    readonly length: number;
    /** 合并区域 */
    readonly merges: ReadonlyArray<ExtractedMerge>;
    /** 列宽信息 */
    readonly cols: ReadonlyArray<ExtractedCol>;
    /** 行高信息 */
    readonly rows: ReadonlyArray<ExtractedRow>;
    /** 数据范围 */
    readonly range: ExtractedRange;
    /** 图片 */
    readonly images: ExtractedImage[];
    /** 第 i 格的 0-based 行号 */
    rowAt(index: number): number;
    /** 第 i 格的 0-based 列号 */
    colAt(index: number): number;
    /** 第 i 格的显示文本（布尔 → TRUE/FALSE；空值 → ''） */
    textAt(index: number): string;
    /** 第 i 格的样式；无样式返回 undefined */
    styleAt(index: number): CellStyleInfo | undefined;
    /** 第 i 格的数字格式码；无则返回 undefined */
    formatAt(index: number): string | undefined;
}

/**
 * 单元格值 → 显示文本
 *
 * 三条规则纯属业务约定，两种来源必须完全一致（历史行为）：
 * 布尔值转 Excel 风格的 `TRUE`/`FALSE`；`null`/`undefined` 视为空串；
 * 其余走 `String()` 自然转换。
 * @param value - 单元格值
 * @returns {string} 用于渲染的文本
 */
export function cellDisplayText(value: string | number | boolean | null | undefined): string {
    if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
    if (value === null || value === undefined) return '';
    return String(value);
}

/** 数组式数据源：直接以 `extract()` 产出的 `ExtractedCell[]` 为底 */
export class ArraySheetSource implements SheetSource {
    /** 单元格数组（引用，不复制） */
    private readonly cells: ExtractedCellLike[];
    /** 合并区域 */
    public readonly merges: ReadonlyArray<ExtractedMerge>;
    /** 列宽信息 */
    public readonly cols: ReadonlyArray<ExtractedCol>;
    /** 行高信息 */
    public readonly rows: ReadonlyArray<ExtractedRow>;
    /** 数据范围 */
    public readonly range: ExtractedRange;
    /** 图片 */
    public readonly images: ExtractedImage[];

    constructor(sheet: ExtractedSheetLike) {
        this.cells = sheet.cells;
        this.merges = sheet.merges;
        this.cols = sheet.cols;
        this.rows = sheet.rows;
        this.range = sheet.range;
        this.images = sheet.images;
    }

    public get length(): number {
        return this.cells.length;
    }

    public rowAt(index: number): number {
        return this.cells[index].row;
    }

    public colAt(index: number): number {
        return this.cells[index].col;
    }

    public textAt(index: number): string {
        return cellDisplayText(this.cells[index].value);
    }

    public styleAt(index: number): CellStyleInfo | undefined {
        return this.cells[index].style;
    }

    public formatAt(index: number): string | undefined {
        return this.cells[index].numberFormat;
    }
}

// ════════════════════════════════════════════════════════════
//  列式打包（Worker → 主线程）
// ════════════════════════════════════════════════════════════

/** 值类型标记：见 {@link PackedSheet.kind} */
export const PACK_KIND = {
    /** 空值（null / undefined）→ 文本为空 */
    EMPTY: 0,
    /** 文本（取自文本池） */
    TEXT: 1,
    /** 数值（存于 `num`，显示文本由 `String()` 还原） */
    NUMBER: 2,
    /** 布尔（存于 `num`，1=TRUE / 0=FALSE） */
    BOOLEAN: 3,
} as const;

/**
 * 列式打包的工作表
 *
 * 设计要点：
 * - `row/col/kind/num/styleIndex/formatIndex/textStart` 长度均为 `count`，
 *   彼此**下标对齐**，故任取一格都是 O(1)；
 * - `styles`/`formats` 是**去重表**：6 万格实测只剩 18 个样式，因此嵌套样式对象
 *   在克隆时只走 18 次；
 * - `text` 是**单一字符串池**（所有文本顺序拼接），`textStart[i]..textStart[i+1]`
 *   即第 i 格的文本切片 —— 把「6 万个字符串对象」压成「1 个字符串 + 一段 Int32」；
 * - 除 `text`/去重表外全部是 typed array，可整块 `transfer`（零拷贝）。
 */
export interface PackedSheet {
    /** 工作表名称 */
    name: string;
    /** 单元格数量（各平行数组的长度） */
    count: number;
    /** 0-based 行号 */
    row: Int32Array;
    /** 0-based 列号 */
    col: Int32Array;
    /** 值类型（{@link PACK_KIND}） */
    kind: Uint8Array;
    /** 数值（kind=NUMBER）或布尔（kind=BOOLEAN，1/0） */
    num: Float64Array;
    /** 样式表下标；-1 = 无样式 */
    styleIndex: Int32Array;
    /** 数字格式表下标；-1 = 无 */
    formatIndex: Int32Array;
    /** 文本池（所有 kind=TEXT 的格子顺序拼接） */
    text: string;
    /** 文本池切片起点，长度 `count + 1`（末位 = 文本池总长） */
    textStart: Int32Array;
    /** 去重后的样式表 */
    styles: Array<CellStyleInfo | null>;
    /** 去重后的数字格式表 */
    formats: string[];
    /** 合并区域 */
    merges: ExtractedMerge[];
    /** 列宽信息 */
    cols: ExtractedCol[];
    /** 行高信息 */
    rows: ExtractedRow[];
    /** 图片（其 `data` 的 buffer 随回包一起 transfer） */
    images: ExtractedImage[];
    /** 数据范围 */
    range: ExtractedRange;
}

/**
 * 把 `extract()` 的产出打包成列式结构
 *
 * 本函数应在 **Worker 内**调用（实测 6.3 万格约 91 ms，主线程无感）。
 * 打包一次的成本远小于「原样克隆省下的那 400+ ms」，这是净收益的来源。
 * @param sheet - 提取出的工作表
 * @returns {PackedSheet} 列式打包结果
 */
export function packSheet(sheet: ExtractedSheetLike): PackedSheet {
    const cells = sheet.cells;
    const count = cells.length;

    const row = new Int32Array(count);
    const col = new Int32Array(count);
    const kind = new Uint8Array(count);
    const num = new Float64Array(count);
    const styleIndex = new Int32Array(count).fill(-1);
    const formatIndex = new Int32Array(count).fill(-1);
    const textStart = new Int32Array(count + 1);

    const styles: Array<CellStyleInfo | null> = [];
    /** 样式编号 → 样式表下标（提取阶段的内联缓存保证「同编号同对象」） */
    const styleIdKeys = new Map<number, number>();
    /** 兜底：样式内容序列化串 → 样式表下标（仅用于没有 `styleId` 的格子） */
    const styleKeys = new Map<string, number>();
    const formats: string[] = [];
    const formatKeys = new Map<string, number>();
    const textParts: string[] = [];
    let textLength = 0;

    for (let i = 0; i < count; i++) {
        const cell = cells[i];
        row[i] = cell.row;
        col[i] = cell.col;
        textStart[i] = textLength;

        const value = cell.value;
        if (typeof value === 'boolean') {
            kind[i] = PACK_KIND.BOOLEAN;
            num[i] = value ? 1 : 0;
        } else if (typeof value === 'number') {
            kind[i] = PACK_KIND.NUMBER;
            num[i] = value;
        } else if (value === null || value === undefined) {
            kind[i] = PACK_KIND.EMPTY;
        } else {
            kind[i] = PACK_KIND.TEXT;
            const text = String(value);
            if (text) {
                textParts.push(text);
                textLength += text.length;
            }
        }

        const style = cell.style;
        if (style) {
            // 样式去重（性能约定，勿改回「一律 JSON.stringify」）：
            //
            // 提取阶段已按样式编号做了内联缓存 —— **同一 `styleId` 的 `style`
            // 必定是同一个对象**，所以拿数字编号直接查表即可，O(1) 且不产生字符串。
            // 历史实现对每格执行 `JSON.stringify(style)`，20 万格带样式时这一项就要
            // 620 ms（还要为每格分配一个字符串），而结果只是 5000 条去重表。
            //
            // 手工构造的 `ExtractedSheet`（测试、旧调用方）可能没有 `styleId`，
            // 那时退回按内容比较，行为与历史实现完全一致。
            let index: number | undefined;
            const styleId = cell.styleId;
            if (styleId !== undefined) {
                const hit = styleIdKeys.get(styleId);
                if (hit !== undefined) {
                    // 同一编号必须是同一个对象（内联缓存保证）；万一不是（异常来源
                    // 复用了编号），退回内容键，不让错误的合并把样式张冠李戴
                    if (styles[hit] === style) index = hit;
                } else {
                    index = styles.length;
                    styles.push(style);
                    styleIdKeys.set(styleId, index);
                }
            }
            if (index === undefined) {
                // 以序列化串为去重键：样式是「值对象」，内容相同即等价（buildCell 只读不改）
                const key = JSON.stringify(style);
                const hit = styleKeys.get(key);
                if (hit === undefined) {
                    index = styles.length;
                    styles.push(style);
                    styleKeys.set(key, index);
                } else {
                    index = hit;
                }
            }
            styleIndex[i] = index;
        }

        const format = cell.numberFormat;
        if (format) {
            let index = formatKeys.get(format);
            if (index === undefined) {
                index = formats.length;
                formats.push(format);
                formatKeys.set(format, index);
            }
            formatIndex[i] = index;
        }
    }
    textStart[count] = textLength;

    return {
        name: sheet.name,
        count,
        row,
        col,
        kind,
        num,
        styleIndex,
        formatIndex,
        text: textParts.join(''),
        textStart,
        styles,
        formats,
        merges: sheet.merges,
        cols: sheet.cols,
        rows: sheet.rows,
        images: sheet.images,
        range: sheet.range,
    };
}

/**
 * 打包整本工作簿
 * @param sheets - 提取出的所有工作表
 * @returns {PackedSheet[]} 列式打包结果
 */
export function packSheets(sheets: ExtractedSheetLike[]): PackedSheet[] {
    return sheets.map((sheet) => packSheet(sheet));
}

/**
 * 收集回包时需要 `transfer` 的 ArrayBuffer（零拷贝）
 *
 * 覆盖全部 typed array 与图片二进制。**注意**：`text` 是字符串，无法转移，
 * 只能按值复制（实测 6 万格约 0.3 MB，代价可忽略）。
 * @param packed - 打包结果
 * @returns {Transferable[]} 可转移对象列表
 */
export function packTransferables(packed: PackedSheet[]): Transferable[] {
    const transfer: Transferable[] = [];
    for (const sheet of packed) {
        // `Int32Array.buffer` 的静态类型是 `ArrayBufferLike`（含 SharedArrayBuffer），
        // 而 `Transferable` 只认 `ArrayBuffer`；这里断言是安全的：数组一律由本模块新建
        transfer.push(
            sheet.row.buffer as ArrayBuffer,
            sheet.col.buffer as ArrayBuffer,
            sheet.kind.buffer as ArrayBuffer,
            sheet.num.buffer as ArrayBuffer,
            sheet.styleIndex.buffer as ArrayBuffer,
            sheet.formatIndex.buffer as ArrayBuffer,
            sheet.textStart.buffer as ArrayBuffer
        );
        for (const image of sheet.images) {
            if (image.data && image.data.buffer) {
                transfer.push(image.data.buffer as ArrayBuffer);
            }
        }
    }
    return transfer;
}

/** 列式数据源：以 {@link PackedSheet} 为底，按需切片/查表还原每格 */
export class PackedSheetSource implements SheetSource {
    private readonly sheet: PackedSheet;

    constructor(sheet: PackedSheet) {
        this.sheet = sheet;
    }

    public get length(): number {
        return this.sheet.count;
    }

    public get merges(): ReadonlyArray<ExtractedMerge> {
        return this.sheet.merges;
    }

    public get cols(): ReadonlyArray<ExtractedCol> {
        return this.sheet.cols;
    }

    public get rows(): ReadonlyArray<ExtractedRow> {
        return this.sheet.rows;
    }

    public get range(): ExtractedRange {
        return this.sheet.range;
    }

    public get images(): ExtractedImage[] {
        return this.sheet.images;
    }

    public rowAt(index: number): number {
        return this.sheet.row[index];
    }

    public colAt(index: number): number {
        return this.sheet.col[index];
    }

    public textAt(index: number): string {
        const kind = this.sheet.kind[index];
        if (kind === PACK_KIND.TEXT) {
            // slice 只产生「切片字符串」视图，不复制字符数据（短串由 V8 自行决定）
            return this.sheet.text.slice(this.sheet.textStart[index], this.sheet.textStart[index + 1]);
        }
        if (kind === PACK_KIND.NUMBER) {
            return String(this.sheet.num[index]);
        }
        if (kind === PACK_KIND.BOOLEAN) {
            return this.sheet.num[index] === 1 ? 'TRUE' : 'FALSE';
        }
        return '';
    }

    public styleAt(index: number): CellStyleInfo | undefined {
        const style = this.sheet.styleIndex[index];
        if (style < 0) return undefined;
        return this.sheet.styles[style] ?? undefined;
    }

    public formatAt(index: number): string | undefined {
        const format = this.sheet.formatIndex[index];
        if (format < 0) return undefined;
        return this.sheet.formats[format];
    }
}
