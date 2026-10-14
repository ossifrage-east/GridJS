import { 
    DataEvents, 
    DEFAULT_CELL_HEIGHT, 
    DEFAULT_CELL_WIDTH, 
    MouseLocation
} from "../constant";
import { EventEmitter } from "../../utils/eventEmitter";

/** Excel 的行数上限（稀疏编码读回时的防御性上限） */
export const MAX_ROW_COUNT = 1048576;
/** Excel 的列数上限（稀疏编码读回时的防御性上限） */
export const MAX_COL_COUNT = 16384;

/**
 * 列头部数据结构
 * @interface ColHeader
 */
export interface ColHeader {
    /** 列的左侧位置 */
    left: number;
    /** 列的宽度 */
    width: number;
    /** 是否隐藏 */
    isHidden?: boolean;
}

/**
 * 行头部数据结构
 * @interface RowHeader
 */
export interface RowHeader {
    /** 行的顶部位置 */
    top: number;
    /** 行的高度 */
    height: number;
    /** 是否隐藏 */
    isHidden?: boolean;
}

/**
 * ColHeaders 类
 * 管理电子表格的列头部，包括列的位置、宽度等信息
 * 提供列宽、偏移量等属性的监听和变更通知
 * 
 * @class ColHeaders
 * @extends EventEmitter
 * @example
 * const colHeaders = new ColHeaders();
 * colHeaders.addCol(0, 10);
 * colHeaders.on('offsetx:changed', (width) => console.log('OffsetX:', width));
 * colHeaders.allColWidth = 1000;
 */
export class ColHeaders extends EventEmitter {
    /** 列头部数据数组 */
    public colHeaders: ColHeader[] = [];
    /** 水平偏移量 */
    private _offsetWidth: number = 0;
    /** 列宽度总和 */
    private _allColWidth: number = 0;

    /**
     * 构造函数
     */
    constructor() {
        super();
    }

    /**
     * 获取所有列头部数据
     * @returns {ColHeader[]} 列头部数组的副本
     */
    public getAllCols(): ColHeader[] {
        return [...this.colHeaders];
    }

    /**
     * 获取指定索引的列头部数据
     * @param {number} index - 列索引, 从0开始
     * @returns {ColHeader | undefined} 列头部数据
     */
    public getAt(index: number): ColHeader | undefined {
        return this.colHeaders[index];
    }

    /**
     * 获取列头部数量
     * @returns {number} 列数量
     */
    public get length(): number {
        return this.colHeaders.length;
    }

    /**
     * 设置水平偏移量
     * @param {number} width - 偏移宽度
     */
    public set offsetWidth(width: number) {
        this.updateProperty(DataEvents.OFFSETX_CHANGED, '_offsetWidth', width);
    }

    /**
     * 获取水平偏移量
     * @returns {number} 偏移宽度
     */
    public get offsetWidth(): number {
        return this._offsetWidth;
    }

    /**
     * 获取所有列宽度的总和（排除隐藏列）
     * @returns {number} 可见列宽度总和
     */
    public get allColWidth(): number {
        return this._allColWidth;
    }
    /**
     * 设置所有列宽度的总和
     * @param {number} width - 列宽度总和
     */
    public set allColWidth(width: number) {
        this.updateProperty(DataEvents.ALL_COL_WIDTH_CHANGED, '_allColWidth', width);
    }

    /**
     * 添加列头部
     * @param {number} [startCol=0] - 起始列索引
     * @param {number} endCol - 结束列索引
     */
    public addCol(startCol: number, endCol: number): void {
        for (let c = startCol; c < endCol; c++) {
            const colHeader = this.createColHeader(c);
            this.colHeaders.push(colHeader);
        }
        this.recalcAllColWidth();
    }

    /**
     * 重新计算可见列的宽度总和，排除隐藏列
     * @returns {number} 可见列宽度总和
     */
    public recalcAllColWidth(): number {
        const total = this.colHeaders
            .filter((cur) => !cur.isHidden)
            .reduce((acc, cur) => acc + cur.width, 0);
        // 走 setter 发出 ALL_COL_WIDTH_CHANGED 事件，让 scroller 等监听方感知总宽变化
        this.allColWidth = total;
        return total;
    }

    /**
     * 重算所有列的 left 位置（跳过隐藏列的 width 累加），使可见列紧密排列
     * 隐藏列的 left 被设为前一个可见列的 right（即被后续可见列覆盖位置），
     * 其 width 不再计入累计偏移，从而在视觉上"让位"给右侧可见列。
     * 同时重算可见列宽度总和。
     */
    public recalcColPositions(): void {
        applyColPositions(this.colHeaders);
        this.recalcAllColWidth();
    }    

    /**
     * 获取可见列的起始索引
     *
     * 语义与原先的「从头线性扫描」完全一致（返回第一个满足
     * `offsetWidth >= left && offsetWidth < left + width` 的列，1-based；
     * 无命中时返回 1），但利用 `left` 随索引单调不减的特性改为二分定位，
     * 复杂度 O(log N) —— 原实现在滚动到深层列时每帧要扫描上万次。
     * @returns {{startCol: number}} 可见列的起始索引, 从1开始
     */
    public getVisibleStartCol(): {startCol: number} {
        const len = this.colHeaders.length;
        if (len === 0) return { startCol: 1 };
        const offset = this.offsetWidth;
        // 二分：最后一个 `left <= offset` 的位置（无则说明 offset 在首列之前）
        let lo = 0, hi = len - 1, k = -1;
        while (lo <= hi) {
            const mid = (lo + hi) >> 1;
            if (this.colHeaders[mid].left <= offset) { k = mid; lo = mid + 1; }
            else hi = mid - 1;
        }
        if (k < 0) return { startCol: 1 };
        // 命中列必须满足右边缘 > offset；从 k 向前收缩到该连续区间的起点
        if (!(this.colHeaders[k].left + this.colHeaders[k].width > offset)) return { startCol: 1 };
        let i = k;
        while (i > 0 && this.colHeaders[i - 1].left + this.colHeaders[i - 1].width > offset) i--;
        return { startCol: i + 1 };
    }

    /**
     * 创建单个列头部
     * @param {number} index - 列索引
     * @returns {ColHeader} 列头部数据
     */
    private createColHeader(index: number): ColHeader {
        const left = this.calculateNewColLeft(index);
        return { left, width: DEFAULT_CELL_WIDTH };
    }

    /**
     * 计算新创建列的左侧位置
     * @param {number} index - 列索引, 从1开始
     * @returns {number} 左侧位置
     */
    private calculateNewColLeft(index: number): number {
        if (index < 1) {
            return 0;
        }
        const prevHeader = this.colHeaders[index - 1];
        return prevHeader.left + prevHeader.width;
    }

    /**
     * 获取指定位置的行和列索引, 排除隐藏列
     * @param {number} x - 水平偏移量
     * @returns {{location: MouseLocation, col: number, row: number}} 位置信息
     */

    public getColHeaderCell(x: number): {location: MouseLocation, col: number, row: number} {
        let location = MouseLocation.IN_COL_HEADER;
        let col = 1;
        for (let c = 0; c < this.colHeaders.length; c++) {
            // 隐藏列不可命中：其位置已被右侧可见列覆盖（recalcColPositions 后 left 紧凑）
            if (this.colHeaders[c].isHidden) continue;
            if (x + this.offsetWidth >= this.colHeaders[c].left + 3
                && x + this.offsetWidth < this.colHeaders[c].left + this.colHeaders[c].width - 3) {
                col = c + 1;
                break;
            }  else if (x + this.offsetWidth >= this.colHeaders[c].left + this.colHeaders[c].width - 3
                && c + 1 < this.colHeaders.length
                && x + this.offsetWidth < this.colHeaders[c + 1].left + 3
            ) {
                location = MouseLocation.IN_COL_HEADER_SPLIT;
                col = c + 1;
                break;
            } else if (x < 3) {
                location = MouseLocation.IN_COL_FROZEN_SPLIT;
                break;
            }
        }
        // left 已由 recalcColPositions 重算为紧凑排列，匹配到的索引即正确列号，无需隐藏补偿
        return { location, col, row: 1 };
    }

    /**
     * 清空所有列头部数据
     */
    public clear(): void {
        this.colHeaders = [];
        this._offsetWidth = 0;
    }
}

/**
 * RowHeaders 类
 * 管理电子表格的行头部，包括行的位置、高度等信息
 * 提供行高、偏移量等属性的监听和变更通知
 * 
 * @class RowHeaders
 * @extends EventEmitter
 * @example
 * const rowHeaders = new RowHeaders();
 * rowHeaders.addRow(0, 100);
 * rowHeaders.allRowHeight = 2400;
 * rowHeaders.on(DataEvents.ALL_ROW_HEIGHT_CHANGED, (height) => console.log('All Row Height:', height));
 * rowHeaders.offsetHeight = 100;
 * rowHeaders.on(DataEvents.OFFSETY_CHANGED, (height) => console.log('OffsetY:', height));
 * const rowHeader = rowHeaders.getAt(1);
 */
export class RowHeaders extends EventEmitter {
    /** 行头部数据数组 */
    public rowHeaders: RowHeader[] = [];
    /** 垂直偏移量 */
    private _offsetHeight: number = 0;
    /** 行高度总和 */
    private _allRowHeight: number = 0;

    /**
     * 构造函数
     */
    constructor() {
        super();
    }

    /**
     * 获取所有行头部数据
     * @returns {RowHeader[]} 行头部数组的副本
     */
    public getAll(): RowHeader[] {
        return [...this.rowHeaders];
    }
    
    /**
     * 获取所有行高度的总和（排除隐藏行）
     * @returns {number} 可见行高度总和
     */
    public get allRowHeight(): number {
        return this._allRowHeight;
    }
    /**
     * 设置所有行高度的总和
     * @param {number} height - 行高度总和
     */
    public set allRowHeight(height: number) {
        this.updateProperty(DataEvents.ALL_ROW_HEIGHT_CHANGED, '_allRowHeight', height);
    }
    /**
     * 获取指定索引的行头部数据
     * @param {number} index - 行索引, 从1开始
     * @returns {RowHeader | undefined} 行头部数据
     */
    public getAt(index: number): RowHeader | undefined {
        return this.rowHeaders[index];
    }

    /**
     * 获取行头部数量
     * @returns {number} 行数量
     */
    public get length(): number {
        return this.rowHeaders.length;
    }

    /**
     * 添加行头部
     * @param {RowHeader} rowHeader - 行头部数据
     */
    public addRowHeader(rowHeader: RowHeader): void {
        this.rowHeaders.push(rowHeader);
    }

    /**
     * 设置垂直偏移量
     * @param {number} height - 偏移高度
     */
    public set offsetHeight(height: number) {
        this.updateProperty(DataEvents.OFFSETY_CHANGED, '_offsetHeight', height);
    }

    /**
     * 获取垂直偏移量
     * @returns {number} 偏移高度
     */
    public get offsetHeight(): number {
        return this._offsetHeight;
    }    

    /**
     * 添加行头部
     * @param {number} [startRow=0] - 起始行索引
     * @param {number} endRow - 结束行索引
     */
    public addRow(startRow: number, endRow: number): void {
        for (let r = startRow; r < endRow; r++) {
            const rowHeader = this.createRowHeader(r);
            this.rowHeaders.push(rowHeader);
        }
        this.recalcAllRowHeight();
    }

    /**
     * 重新计算可见行的高度总和，排除隐藏行
     * @returns {number} 可见行高度总和
     */
    public recalcAllRowHeight(): number {
        const total = this.rowHeaders
            .filter((cur) => !cur.isHidden)
            .reduce((acc, cur) => acc + cur.height, 0);
        // 走 setter 发出 ALL_ROW_HEIGHT_CHANGED 事件，让 scroller 等监听方感知总高变化
        this.allRowHeight = total;
        return total;
    }

    /**
     * 重算所有行的 top 位置（跳过隐藏行的 height 累加），使可见行紧密排列
     * 隐藏行的 top 被设为前一个可见行的 bottom（即被后续可见行覆盖位置），
     * 其 height 不再计入累计偏移，从而在视觉上"让位"给下方可见行。
     * 同时重算可见行高度总和。
     */
    public recalcRowPositions(): void {
        applyRowPositions(this.rowHeaders);
        this.recalcAllRowHeight();
    }
    
    /**
     * 获取可见行的起始索引
     *
     * 语义与原先的「从头线性扫描」完全一致（返回第一个满足
     * `offsetHeight >= top && offsetHeight <= top + height` 的行，1-based；
     * 无命中时返回 1），但利用 `top` 随索引单调不减的特性改为二分定位，
     * 复杂度 O(log N) —— 原实现在滚动到深层行时每帧要扫描上万次。
     * @returns {{startRow: number}} 可见行的起始索引, 从1开始
     */
    public getVisibleStartRow(): {startRow: number} {
        const len = this.rowHeaders.length;
        if (len === 0) return { startRow: 1 };
        const offset = this.offsetHeight;
        let lo = 0, hi = len - 1, k = -1;
        while (lo <= hi) {
            const mid = (lo + hi) >> 1;
            if (this.rowHeaders[mid].top <= offset) { k = mid; lo = mid + 1; }
            else hi = mid - 1;
        }
        if (k < 0) return { startRow: 1 };
        // 命中行必须满足下边缘 >= offset；否则说明偏移已越过内容底部（与旧实现一致返回 1）
        if (!(this.rowHeaders[k].top + this.rowHeaders[k].height >= offset)) return { startRow: 1 };
        let i = k;
        while (i > 0 && this.rowHeaders[i - 1].top + this.rowHeaders[i - 1].height >= offset) i--;
        return { startRow: i + 1 };
    }

    /**
     * 创建单个新的行头部
     * @param {number} index - 行索引, 从1开始
     * @returns {RowHeader} 行头部数据
     */
    private createRowHeader(index: number): RowHeader {
        const top = this.calculateNewRowTop(index);
        return { top, height: DEFAULT_CELL_HEIGHT };
    }

    /**
     * 计算新创建的行的顶部位置
     * @param {number} index - 行索引, 从1开始
     * @returns {number} 顶部位置
     */
    private calculateNewRowTop(index: number): number {
        if (index === 0) {
            return 0;
        }
        const prevHeader = this.rowHeaders[index - 1];
        return prevHeader.top + prevHeader.height;
    }

    /**
     * 获取指定位置的行和列索引
     * @param {number} y - 鼠标Y坐标（从0开始）
     * @returns {{location: MouseLocation, col: number, row: number}} 位置信息
     */
    public getRowHeaderCell(y: number): {location: MouseLocation, col: number, row: number} {
        let location = MouseLocation.IN_ROW_HEADER;
        let row = 1;
        for (let r = 0; r < this.rowHeaders.length; r++) {
            // 隐藏行不可命中：其位置已被下方可见行覆盖（recalcRowPositions 后 top 紧凑）
            if (this.rowHeaders[r].isHidden) continue;
            if (y >= 3 && y + this.offsetHeight >= this.rowHeaders[r].top + 3
                && y + this.offsetHeight <= this.rowHeaders[r].top + this.rowHeaders[r].height - 3) {
                row = r + 1;
                break;
            } else if (y + this.offsetHeight >= this.rowHeaders[r].top + this.rowHeaders[r].height - 3
                && r + 1 < this.rowHeaders.length
                && y + this.offsetHeight < this.rowHeaders[r + 1].top + 3 ) {
                location = MouseLocation.IN_ROW_HEADER_SPLIT;
                row = r + 1;
                break;
            } else if (y < 3) {
                location = MouseLocation.IN_ROW_FROZEN_SPLIT;
                break;
            }
        }
        // top 已由 recalcRowPositions 重算为紧凑排列，匹配到的索引即正确行号，无需隐藏补偿
        return { location, col: 1, row };
    }

    /**
     * 清空所有行头部数据
     */
    public clear(): void {
        this.rowHeaders = [];
        this._offsetHeight = 0;
    }
}

// ════════════════════════════════════════════════════════════
//  稀疏轴编码（存档 v4 用）
// ════════════════════════════════════════════════════════════

/**
 * 一个轴（行或列）的稀疏编码
 *
 * ## 为什么必须稀疏（性能约定，勿回退为「存整个对象数组」）
 *
 * 膨胀维度的工作簿（`maxRow = 1048576`）在内存里就是 104 万个
 * `{ top, height }` 对象。把这一整串写进存档会凭空多出 **28.6 MB**，
 * 而其中 104 万行里只有几百行真的被改过尺寸 —— 剩下的全是 `defaultSize`。
 * 存档变大不是「多占点磁盘」这么轻描淡写：装载路径要先把这段 JSON
 * `JSON.parse` 出来，实测这一项就要 232 ms 的**整段同步**。
 *
 * 稀疏编码只写「例外」：
 * - `count` —— 轴长度（数量本身必须保留，否则滚动条与可视区计算全错）；
 * - `defaultSize` —— 绝大多数项的尺寸，不逐项写；
 * - `sizeFlat` —— 非默认尺寸的 `[索引, 尺寸, 索引, 尺寸, …]` 扁平数组；
 * - `hidden` —— 隐藏项索引（数量少，单独列）。
 *
 * 位置（`top`/`left`）**不存**：它由尺寸累加派生，读回时用同一套规则重算
 * （{@link applyColPositions} / {@link applyRowPositions}），不会出现
 * 「位置与尺寸互相矛盾」的存档。
 */
export interface SparseAxis {
    /** 轴长度（列数 / 行数） */
    count: number;
    /** 默认尺寸（等于它的项不进 `sizeFlat`） */
    defaultSize: number;
    /** 非默认尺寸的 `[索引, 尺寸, …]` 扁平数组 */
    sizeFlat: number[];
    /** 隐藏项索引 */
    hidden: number[];
}

/** 任一尺寸字段的头部视图 */
type SizedHeader = { isHidden?: boolean; width?: number; height?: number };

/**
 * 把一组头部编码成稀疏轴
 * @param headers - 头部数组（读 `width` 或 `height`，缺失视为默认值）
 * @param defaultSize - 默认尺寸（列宽 / 行高）
 * @param sizeKey - 取尺寸的字段名
 * @param limit - 数量上限（防御性）
 * @returns {SparseAxis} 稀疏编码
 */
function encodeAxis(
    headers: readonly SizedHeader[],
    defaultSize: number,
    sizeKey: 'width' | 'height',
    limit: number
): SparseAxis {
    const count = Math.min(headers.length, limit);
    const sizeFlat: number[] = [];
    const hidden: number[] = [];
    for (let i = 0; i < count; i++) {
        const h = headers[i];
        if (!h) continue;
        const size = h[sizeKey];
        if (typeof size === 'number' && size !== defaultSize) sizeFlat.push(i, size);
        if (h.isHidden) hidden.push(i);
    }
    return { count, defaultSize, sizeFlat, hidden };
}

/**
 * 把稀疏轴解码回头部数组（位置按「跳过隐藏项累加」规则就地算好）
 *
 * 读回的是外部数据，先做形状校验：任何非法字段按「该字段缺省」处理，
 * 并且数量被 Excel 上限截断 —— 一个损坏的 `count: 1e9` 不允许变成内存炸弹。
 * @param axis - 稀疏编码（可能来自损坏的存档）
 * @param defaultSize - 默认尺寸（编码中的 `defaultSize` 非法时用它兜底）
 * @param sizeKey - 写尺寸的字段名
 * @param limit - 数量上限
 * @returns {Array} 头部数组（位置已算好）
 */
function decodeAxis(
    axis: unknown,
    defaultSize: number,
    sizeKey: 'width' | 'height',
    limit: number
): Array<Record<string, number | boolean>> {
    const raw = (axis ?? {}) as Partial<SparseAxis>;
    const defaulted = Number.isFinite(raw.defaultSize) ? (raw.defaultSize as number) : defaultSize;
    const count = Number.isFinite(raw.count)
        ? Math.max(0, Math.min(Math.floor(raw.count as number), limit))
        : 0;

    const sizes = new Map<number, number>();
    const flat = raw.sizeFlat;
    if (Array.isArray(flat)) {
        for (let i = 0; i + 1 < flat.length; i += 2) {
            const index = flat[i];
            const size = flat[i + 1];
            if (Number.isFinite(index) && Number.isFinite(size) && index >= 0 && index < count) {
                sizes.set(index, size);
            }
        }
    }
    const hidden = new Set<number>();
    if (Array.isArray(raw.hidden)) {
        for (const index of raw.hidden) {
            if (Number.isFinite(index) && index >= 0 && index < count) hidden.add(index);
        }
    }

    const posKey = sizeKey === 'width' ? 'left' : 'top';
    const headers: Array<Record<string, number | boolean>> = new Array(count);
    let position = 0;
    for (let i = 0; i < count; i++) {
        const size = sizes.get(i) ?? defaulted;
        const isHidden = hidden.has(i);
        // 位置与 `applyColPositions`/`applyRowPositions` 同规则：
        // 隐藏项的位置与前一项相同（被后继可见项覆盖），且不计入累加
        headers[i] = isHidden
            ? { [posKey]: position, [sizeKey]: size, isHidden: true }
            : { [posKey]: position, [sizeKey]: size };
        if (!isHidden) position += size;
    }
    return headers;
}

/**
 * 重算列位置（`ColHeaders.recalcColPositions` 与稀疏解码共用同一套规则）
 * @param headers - 列头数组（就地修改 `left`）
 */
export function applyColPositions(headers: ColHeader[]): void {
    let left = 0;
    for (const h of headers) {
        h.left = left;
        if (!h.isHidden) left += h.width;
    }
}

/**
 * 重算行位置（`RowHeaders.recalcRowPositions` 与稀疏解码共用同一套规则）
 * @param headers - 行头数组（就地修改 `top`）
 */
export function applyRowPositions(headers: RowHeader[]): void {
    let top = 0;
    for (const h of headers) {
        h.top = top;
        if (!h.isHidden) top += h.height;
    }
}

/** 列头 → 稀疏轴 */
export function encodeColAxis(headers: readonly ColHeader[]): SparseAxis {
    return encodeAxis(headers, DEFAULT_CELL_WIDTH, 'width', MAX_COL_COUNT);
}

/** 行头 → 稀疏轴 */
export function encodeRowAxis(headers: readonly RowHeader[]): SparseAxis {
    return encodeAxis(headers, DEFAULT_CELL_HEIGHT, 'height', MAX_ROW_COUNT);
}

/** 稀疏轴 → 列头（位置已算好） */
export function decodeColAxis(axis: unknown): ColHeader[] {
    return decodeAxis(axis, DEFAULT_CELL_WIDTH, 'width', MAX_COL_COUNT) as unknown as ColHeader[];
}

/** 稀疏轴 → 行头（位置已算好） */
export function decodeRowAxis(axis: unknown): RowHeader[] {
    return decodeAxis(axis, DEFAULT_CELL_HEIGHT, 'height', MAX_ROW_COUNT) as unknown as RowHeader[];
}