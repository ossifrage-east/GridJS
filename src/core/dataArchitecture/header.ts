import { 
    DataEvents, 
    DEFAULT_CELL_HEIGHT, 
    DEFAULT_CELL_WIDTH, 
    MouseLocation
} from "../constant";
import { EventEmitter } from "../../utils/eventEmitter";

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
        let left = 0;
        for (const h of this.colHeaders) {
            h.left = left;
            if (!h.isHidden) left += h.width;
        }
        this.recalcAllColWidth();
    }    

    /**
     * 获取可见列的起始索引
     * @returns {{startCol: number}} 可见列的起始索引, 从1开始
     */
    public getVisibleStartCol(): {startCol: number} {
        let startCol = 0;
        for (let i = 0; i < this.colHeaders.length; i++) {
            if (this.offsetWidth >= this.colHeaders[i].left
                && this.offsetWidth < this.colHeaders[i].left + this.colHeaders[i].width
            ) {
                startCol = i;
                break;
            }
        }
        return { startCol: startCol + 1 };
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
        let top = 0;
        for (const h of this.rowHeaders) {
            h.top = top;
            if (!h.isHidden) top += h.height;
        }
        this.recalcAllRowHeight();
    }
    
    /**
     * 获取可见行的起始索引
     * @returns {{startRow: number}} 可见行的起始索引, 从1开始
     */
    public getVisibleStartRow(): {startRow: number} {
        let startRow = 0;
        for (let i = 0; i < this.rowHeaders.length; i++) {
            if (this.offsetHeight >= this.rowHeaders[i].top
                && this.offsetHeight <= this.rowHeaders[i].top + this.rowHeaders[i].height
            ) {
                startRow = i;
                break;
            }
        }
        return { startRow: startRow + 1 };
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