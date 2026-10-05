import { EventEmitter } from "../../utils/eventEmitter";
import { DataEvents } from "../constant";

/**
 * VisibleView 类
 * 管理电子表格可见视图的尺寸信息
 * 包括行头宽度、列头高度、工作表宽度和高度
 * 属性变更时会触发相应的事件通知
 * 
 * @class VisibleView
 * @extends EventEmitter
 * @example
 * const visibleView = new VisibleView();
 * visibleView.rowHeaderWidth = 80;
 * visibleView.on('rowHeaderWidthChange', (width) => {
 *     console.log('Row header width changed:', width);
 * });
 */
export class VisibleView extends EventEmitter {
    /** 行头宽度 */
    private _rowHeaderWidth: number = 0;
    /** 列头高度 */
    private _colHeaderHeight: number = 0;
    /** 工作表宽度 */
    private _sheetWidth: number = 0;
    /** 工作表高度 */
    private _sheetHeight: number = 0;

    /**
     * 构造函数
     * @constructor
     */
    constructor() {
        super();
    }

    /**
     * 获取行头宽度
     * @returns {number} 行头宽度
     */
    get rowHeaderWidth(): number {
        return this._rowHeaderWidth;
    }

    /**
     * 设置行头宽度
     * @param {number} value - 行头宽度
     * @fires VisibleView#rowHeaderWidthChange
     */
    set rowHeaderWidth(value: number) {
        this.updateProperty(DataEvents.ROW_HEADER_WIDTH_CHANGED, '_rowHeaderWidth', value);
    }

    /**
     * 获取列头高度
     * @returns {number} 列头高度
     */
    get colHeaderHeight(): number {
        return this._colHeaderHeight;
    }

    /**
     * 设置列头高度
     * @param {number} value - 列头高度
     * @fires VisibleView#colHeaderHeightChange
     */
    set colHeaderHeight(value: number) {
        this.updateProperty(DataEvents.COL_HEADER_HEIGHT_CHANGED, '_colHeaderHeight', value);
    }

    /**
     * 获取工作表宽度
     * @returns {number} 工作表宽度
     */
    get sheetWidth(): number {
        return this._sheetWidth;
    }

    /**
     * 设置工作表宽度
     * @param {number} value - 工作表宽度
     * @fires VisibleView#sheetWidthChange
     */
    set sheetWidth(value: number) {
        this.updateProperty(DataEvents.SHEET_WIDTH_CHANGED, '_sheetWidth', value);
    }

    /**
     * 获取工作表高度
     * @returns {number} 工作表高度
     */
    get sheetHeight(): number {
        return this._sheetHeight;
    }

    /**
     * 设置工作表高度
     * @param {number} value - 工作表高度
     * @fires VisibleView#sheetHeightChange
     */
    set sheetHeight(value: number) {
        this.updateProperty(DataEvents.SHEET_HEIGHT_CHANGED, '_sheetHeight', value);
    }
}