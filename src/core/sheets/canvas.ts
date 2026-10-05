/*  # canvas 组件 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import downFill from '../../assets/images/cursor/down-fill.svg';
import forwardFill from '../../assets/images/cursor/forward-fill.svg';
import crossEmpty from '../../assets/images/cursor/cross-empty.svg'
import crosshair from '../../assets/images/cursor/plus-lg.svg'
import { CellContent, LineText } from "../dataArchitecture/cellContent";
import { EventEmitter } from "../../utils/eventEmitter";
import { CELL_PADDING, DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE, GRID_LINE_COLOR, KEYWORDS, MouseLocation, TEXT_COLOR } from "../constant";
import { DataCollection } from "../dataArchitecture/dataCollection";
import { Menu, MenuContent } from "../common/menu";


/**
 * Canvas 初始化选项接口
 * @interface CanvasOptions
 */
export interface CanvasOptions {
    /** 父容器元素 */
    parentElement: HTMLElement;
    /** 数据集合实例 */
    data: DataCollection;
    /** 菜单实例 */
    menu?: Menu;
}

/**
 * Canvas 基类
 * 提供画布的基础操作和文本测量功能
 * 
 * @class Canvas
 * @extends EventEmitter
 * @example
 * const canvas = new Canvas({
 *     parentElement: document.getElementById('container'),
 *     data: new DataCollection()
 * });
 */
export class Canvas extends EventEmitter {
    /** Canvas 元素 */
    public canvas: HTMLCanvasElement;
    /** 2D 绘图上下文 */
    public ctx: CanvasRenderingContext2D;
    /** 父容器元素 */
    public parentElement: HTMLElement;
    /** 数据集合实例 */
    public data: DataCollection;
    /** 菜单实例 */
    public menu?: Menu;
    /** 记录当前是否应用了半像素偏移 */
    private hasHalfPixelOffset: boolean = false;
    /** 右键菜单的操作对象类型：行头（row）/列头（col）/单元格（cell） */
    protected _contextTarget: 'row' | 'col' | 'cell' = 'cell';
    menuContent: MenuContent;

    /**
     * 构造函数
     * @constructor
     * @param {CanvasOptions} options - 初始化选项
     */
    constructor(options: CanvasOptions) {
        super();
        this.data = options.data;
        this.menu = options.menu;
        this.setupMenuContent();
        this.canvas = document.createElement('canvas');
        const ctx = this.canvas.getContext('2d');
        if (!ctx) {
            throw new Error("Failed to get 2d context for editCanvas");
        }
        this.ctx = ctx;
        this.setupContext();
    }

    /**
     * 设置画布上下文参数
     * 所有绘制坐标均为 CSS 像素，内部自动转换为物理像素
     * @public
     */
    public setupContext(): void {
        this.ctx.imageSmoothingEnabled = true;
        this.ctx.imageSmoothingQuality = 'high';
        this.ctx.strokeStyle = GRID_LINE_COLOR;
        this.ctx.lineWidth = this.toDevicePixels(1);  // 1 CSS像素 = DPR物理像素
        this.ctx.font = `${Math.round(DEFAULT_FONT_SIZE * this.data.zoom * window.devicePixelRatio)}px ${DEFAULT_FONT_FAMILY}`;
        this.ctx.fillStyle = TEXT_COLOR;
        this.ctx.textBaseline = 'middle';
        this.ctx.textAlign = 'center';
    }

    /**
     * 将 CSS 像素转换为物理像素
     * @param {number} cssPixels - CSS 像素值
     * @returns {number} 物理像素值（已四舍五入）
     * @public
     */
    private toDevicePixels(cssPixels: number): number {
        return Math.round(cssPixels * window.devicePixelRatio);
    }

    /**
     * 设置线条宽度（CSS 像素）
     * 自动处理 devicePixelRatio 和半像素偏移
     * @param {number} widthCSS - 线条宽度（CSS 像素）
     * @public
     */
    public setLineWidth(widthCSS: number): void {
        const widthDevice = this.toDevicePixels(widthCSS);
        this.ctx.lineWidth = widthDevice;

        // 奇数像素宽度时添加半像素偏移
        if (widthDevice % 2 !== 0) {
            const halfPixel = 0.5 * window.devicePixelRatio;
            this.ctx.translate(halfPixel, halfPixel);
            this.hasHalfPixelOffset = true;
        }
    }

    public removeHalfPixelOffset(): void {
        // 先撤销之前的偏移（如果有）
        if (this.hasHalfPixelOffset) {
            const halfPixel = 0.5 * window.devicePixelRatio;
            this.ctx.translate(-halfPixel, -halfPixel);
            this.hasHalfPixelOffset = false;
        }
    }
    
    /**
     * 清空画布内容
     */
    public clear(): void {
        this.ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);
    }

    /**
     * 初始化画布尺寸
     * @param {number} left - 画布左边界位置
     * @param {number} top - 画布顶部边界位置
     * @param {number} width - 画布宽度
     * @param {number} height - 画布高度
     */
    public initCanvas(left: number, top: number, width: number, height: number): void {
        this.canvas.width = width * window.devicePixelRatio;
        this.canvas.height = height * window.devicePixelRatio;
        this.canvas.style.width = `${width}px`;
        this.canvas.style.height = `${height}px`;
    }

    /**
     * 测量单元格文本的宽度和高度
     * @param {string} cell - 单元格ID
     * @returns {CellContent} 包含单元格内容位置、内容尺寸和行文本信息的对象
     */
    public measureCellText(cell: string): CellContent {
        const { col, row } = this.data.getCellColAndRow(cell);
        const { left: cellLeft, top: cellTop, width: cellWidth, height: cellHeight } = this.data.calculateCellRect(col, row);
        
        const index = this.data.values.findIndex((item) => item.cell === cell);
        if (index === -1) {
            const x = Math.round(CELL_PADDING + cellWidth * this.data.zoom / 2);
            const y = Math.round(CELL_PADDING + (cellHeight * this.data.zoom - DEFAULT_FONT_SIZE * this.data.zoom) / 2);

            return {
                editLeft: Math.round(cellLeft * this.data.zoom),
                editTop: Math.round(cellTop * this.data.zoom),
                charsWidth: 0,
                charsHeight: 0,
                contentWidth: Math.round(cellWidth * this.data.zoom),
                contentHeight: Math.round(cellHeight * this.data.zoom),
                lines: [{ text: '', lineWidth: 0, 
                        lineHeight: Math.round(DEFAULT_FONT_SIZE * this.data.zoom), x: x, y: y, editX: x, editY: y }] as LineText[]
            };
        }
        
        const value = this.data.getCellValue(cell);
        const fontSize = Math.round((this.data.values[index].fontSize || DEFAULT_FONT_SIZE) 
            * this.data.zoom);
        const letterSpacing = this.data.values[index].letterSpacing || 0;
        const wrap = this.data.values[index].wrap || false;
        const lineSpacing = this.data.values[index].lineSpacing || 0;
        
        const paragraphs = value.split('\n');
        let cellValueWidth = 0, cellValueHeight = 0, charStartIndex = 0;
        const lines = [] as LineText[];
        
        for (let [paragraphIndex, paragraph] of paragraphs.entries()) {
            paragraph = paragraphIndex < paragraphs.length - 1 ? paragraph + '\n' : paragraph; 
            if (!paragraph || paragraph === '\n') {
                const lineHeight = fontSize + (paragraphIndex < paragraphs.length - 1 ? lineSpacing : 0);
                cellValueHeight += lineHeight;
                lines.push({ text: paragraph, lineWidth: 0, lineHeight: lineHeight, x: 0, y: 0, editX: 0, editY: 0 });
                charStartIndex += paragraph.length;
                continue;
            }
            
            const { lineWidth: paragraphWidth, lineHeight: paragraphHeight } 
                = this.calculateSingleLineGeometry(cell, charStartIndex, paragraph, letterSpacing);
            
            if (wrap && paragraphWidth + CELL_PADDING * 2 > cellWidth * this.data.zoom) {
                let startIndex = 0;
                while (startIndex < paragraph.length) {
                    let endIndex = this.findOptimalLineBreak(cell, paragraph, startIndex, cellWidth * this.data.zoom, letterSpacing, charStartIndex);
                    const lineText = paragraph.substring(startIndex, endIndex);
                    const { lineWidth, lineHeight } = this.calculateSingleLineGeometry(cell, charStartIndex + startIndex, lineText, letterSpacing);
                    cellValueWidth = Math.max(cellValueWidth, lineWidth);
                    cellValueHeight += lineHeight + (paragraphIndex < paragraphs.length - 1 ? lineSpacing : 0);
                    lines.push({ text: lineText, lineWidth: lineWidth, 
                        lineHeight: lineHeight + (paragraphIndex < paragraphs.length - 1 ? lineSpacing : 0), x: 0, y: 0, editX: 0, editY: 0 });                    
                    startIndex = endIndex;
                }
            } else {
                cellValueWidth = Math.max(cellValueWidth, paragraphWidth);
                cellValueHeight += paragraphHeight + (paragraphIndex < paragraphs.length - 1 ? lineSpacing : 0);
                lines.push({ text: paragraph, lineWidth: paragraphWidth, 
                    lineHeight: paragraphHeight + (paragraphIndex < paragraphs.length - 1 ? lineSpacing : 0), x: 0, y: 0, editX: 0, editY: 0 });
            }
            
            charStartIndex += paragraph.length;
        }
        
        const contentWidth = Math.max(cellValueWidth, cellWidth * this.data.zoom);
        const contentHeight = Math.max(cellValueHeight, cellHeight * this.data.zoom);
        
        let currentY = 0;
        for (const [lineIndex, line] of lines.entries()) {
            if (this.data.values.find((item) => item.cell === cell)?.textAlign === 'left') {
                lines[lineIndex].x = CELL_PADDING;
            } else if (this.data.values.find((item) => item.cell === cell)?.textAlign === 'right') {
                lines[lineIndex].x = Math.round(cellWidth * this.data.zoom - line.lineWidth - CELL_PADDING);
            } else {
                lines[lineIndex].x = Math.round((cellWidth * this.data.zoom - line.lineWidth) / 2);
            }
            
            if (this.data.values.find((item) => item.cell === cell)?.alignItems === 'top') {
                lines[lineIndex].y = Math.round(CELL_PADDING + currentY);
            } else if (this.data.values.find((item) => item.cell === cell)?.alignItems === 'bottom') {
                lines[lineIndex].y = Math.round(currentY
                    + cellHeight * this.data.zoom - cellValueHeight - CELL_PADDING);
            } else {
                lines[lineIndex].y = Math.round(currentY
                    + (cellHeight * this.data.zoom - cellValueHeight) / 2);
            }
            currentY += line.lineHeight;
        }
        
        const minLeft = Math.min(...lines.map(line => line.x));
        const minTop = Math.min(...lines.map(line => line.y));
        
        const editLeft = Math.round(minLeft < CELL_PADDING ? cellLeft * this.data.zoom + minLeft : cellLeft * this.data.zoom);
        const editTop = Math.round(minTop < CELL_PADDING ? cellTop * this.data.zoom + minTop : cellTop * this.data.zoom);
        
        for (let i = 0; i < lines.length; i++) {
            lines[i].editX = Math.round(cellLeft * this.data.zoom - editLeft + lines[i].x);
            lines[i].editY = Math.round(cellTop * this.data.zoom - editTop + lines[i].y);
        } 
        
        return {
            editLeft: editLeft,
            editTop: editTop,
            charsWidth: Math.round(cellValueWidth + CELL_PADDING * 4),
            charsHeight: Math.round(cellValueHeight + CELL_PADDING * 4),
            contentWidth: Math.round(contentWidth + (contentWidth > cellWidth * this.data.zoom + CELL_PADDING * 2 ? CELL_PADDING : 0)),
            contentHeight: Math.round(contentHeight),
            lines: lines
        };
    }
    
    /**
     * 计算单行文本在考虑字间距后的总宽度和高度
     * @private
     * @param {string} cell - 单元格ID
     * @param {number} charStartIndex - 字符起始索引
     * @param {string} paragraph - 段落文本
     * @param {number} letterSpacing - 字间距
     * @returns {{ lineWidth: number, lineHeight: number }} 行宽和行高
     */
    private calculateSingleLineGeometry(cell: string, charStartIndex: number, paragraph: string, letterSpacing: number): { lineWidth: number; lineHeight: number } {
        if (!paragraph || paragraph === '\n') return { lineWidth: 0, lineHeight: 0 };

        let totalWidth = 0, fontMaxHeight = 0;
        // letterSpacing 仅加在"当前字符是可见字符，且下一字符也是可见字符"之间
        // \n 不贡献宽度，也不参与 letterSpacing 的累加
        for (let i = 0; i < paragraph.length; i++) {
            const ch = paragraph[i];
            if (ch === '\n') continue;
            const { charWidth, charHeight } = this.calculateCharGeometry(cell, charStartIndex + i);
            const hasNextVisible = i + 1 < paragraph.length && paragraph[i + 1] !== '\n';
            totalWidth += charWidth + (hasNextVisible ? letterSpacing : 0);
            fontMaxHeight = Math.max(fontMaxHeight, charHeight);
        }
        return { lineWidth: Math.round(totalWidth), lineHeight: Math.round(fontMaxHeight) };
    }

    /**
     * 查找最佳换行点，使行宽尽可能接近最大宽度
     * @private
     * @param {string} cell - 单元格ID
     * @param {string} paragraph - 段落文本
     * @param {number} startIndex - 段落内起始索引（相对段落）
     * @param {number} maxWidth - 最大宽度
     * @param {number} letterSpacing - 字间距
     * @param {number} charStartIndex - 该段在 chars 数组中的全局起始索引
     * @returns {number} 最佳换行位置索引（相对段落）
     */
    private findOptimalLineBreak(cell: string, paragraph: string, startIndex: number, maxWidth: number, letterSpacing: number, charStartIndex: number = 0): number {
        if (startIndex + 1 >= paragraph.length) {
            return paragraph.length;
        }

        let bestBreakIndex = startIndex + 1;
        let currentWidth = 0;
        let currentIndex = startIndex;

        while (currentIndex < paragraph.length) {
            // 使用全局绝对字符索引查询样式，确保多段文本下查询到正确的字符
            const { charWidth } = this.calculateCharGeometry(cell, charStartIndex + currentIndex);

            if (currentWidth + charWidth + (currentIndex > startIndex ? letterSpacing : 0) > maxWidth) {
                return bestBreakIndex > startIndex ? bestBreakIndex : currentIndex + 1;
            }

            currentWidth += charWidth + (currentIndex > startIndex ? letterSpacing : 0);
            currentIndex++;
            bestBreakIndex = currentIndex;
        }

        return paragraph.length;
    }

    /**
     * 计算单个字符的宽度和高度
     * 样式继承优先级：字符级样式 → 默认样式
     * @param {string} cell - 单元格ID
     * @param {number} charIndex - 字符索引
     * @returns {{ charWidth: number, charHeight: number, fontColor: string, fontWeight: string, fontStyle: string, fontSize: number, fontFamily: string, underline: boolean, strikethrough: boolean }} 
     *          字符几何信息和样式信息
     */
    public calculateCharGeometry(cell: string, charIndex: number): { charWidth: number; charHeight: number; fontColor: string; fontWeight: string; fontStyle: string; fontSize: number; fontFamily: string; underline: boolean; strikethrough: boolean } {
        const cellObj = this.data.values.find((item) => item.cell === cell);
        // chars 数组是字符内容与样式的唯一数据源，避免额外重建整串字符串
        const charObj = cellObj?.chars?.[charIndex];

        // 样式继承：字符级样式 → 默认样式
        const pick = <T>(charVal: T | undefined, defaultVal: T): T =>
            charVal !== undefined ? charVal : defaultVal;

        const fontColor = pick(charObj?.fontColor, TEXT_COLOR);
        const fontFamily = pick(charObj?.fontFamily, DEFAULT_FONT_FAMILY);
        const underline = pick(charObj?.underline, false);
        const strikethrough = pick(charObj?.strikethrough, false);
        const fontWeight = pick(charObj?.fontWeight, false) ? 'bold' : '';
        const fontStyle = pick(charObj?.fontStyle, false) ? 'italic' : '';
        const fontSize = Math.round(pick(charObj?.fontSize, DEFAULT_FONT_SIZE) * this.data.zoom);

        // 从 chars 数组读取字符内容，索引越界或换行等关键字不占宽度
        const char = charObj?.char;
        const isMeasurable = typeof char === 'string' && char.length > 0 && !KEYWORDS.includes(char);

        this.ctx.save();
        const fontParts: string[] = [];
        if (fontStyle) fontParts.push(fontStyle);
        if (fontWeight) fontParts.push(fontWeight);
        fontParts.push(`${fontSize}px`, fontFamily);
        this.ctx.font = fontParts.join(' ');
        const charWidth = isMeasurable ? this.ctx.measureText(char).width : 0;
        this.ctx.restore();

        return { charWidth, charHeight: fontSize, fontColor, fontWeight, fontStyle, fontSize, fontFamily, underline, strikethrough };
    }

    /**
     * 绘制线段
     * @param {number} x1 - 起点X坐标（CSS像素）
     * @param {number} y1 - 起点Y坐标（CSS像素）
     * @param {number} x2 - 终点X坐标（CSS像素）
     * @param {number} y2 - 终点Y坐标（CSS像素）
     * @param {number} [lineWidth] - 线条宽度（CSS像素），不传则使用当前 ctx.lineWidth
     * @public
     */
    public drawLine(x1: number, y1: number, x2: number, y2: number, lineWidth?: number): void {
        if (lineWidth !== undefined) {
            this.setLineWidth(lineWidth);
        }
        this.ctx.beginPath();
        this.moveTo(x1, y1);
        this.lineTo(x2, y2);
        this.ctx.stroke();
        // 恢复到默认线宽
        this.removeHalfPixelOffset();
    }

    /**
     * 绘制文本
     * @param {string} text - 文本内容
     * @param {number} x - X坐标
     * @param {number} y - Y坐标
     */
    public fillText(text: string, x: number, y: number): void {
        const _x = Math.round(x * window.devicePixelRatio);
        const _y = Math.round(y * window.devicePixelRatio);
        this.ctx.fillText(text, _x, _y);
    }

    /**
     * 填充矩形
     * @param {number} x - 左上角X坐标
     * @param {number} y - 左上角Y坐标
     * @param {number} width - 宽度
     * @param {number} height - 高度
     */
    public fillRect(x: number, y: number, width: number, height: number): void {
        const _x = Math.round(x * window.devicePixelRatio);
        const _y = Math.round(y * window.devicePixelRatio);
        const _width = Math.round(width * window.devicePixelRatio);
        const _height = Math.round(height * window.devicePixelRatio);
        this.ctx.fillRect(_x, _y, _width, _height);
    }

    /**
     * 绘制矩形边框
     * @param {number} x - 左上角X坐标
     * @param {number} y - 左上角Y坐标
     * @param {number} width - 宽度
     * @param {number} height - 高度
     */
    public strokeRect(x: number, y: number, width: number, height: number): void {
        const _x = Math.round(x * window.devicePixelRatio);
        const _y = Math.round(y * window.devicePixelRatio);
        const _width = Math.round(width * window.devicePixelRatio);
        const _height = Math.round(height * window.devicePixelRatio);
        this.ctx.strokeRect(_x, _y, _width, _height);
    }

    /**
     * 移动画笔到指定位置
     * @param {number} x - X坐标
     * @param {number} y - Y坐标
     */
    public moveTo(x: number, y: number): void {
        const _x = Math.round(x * window.devicePixelRatio);
        const _y = Math.round(y * window.devicePixelRatio);
        this.ctx.moveTo(_x, _y);
    }

    /**
     * 从当前位置绘制线段到指定位置
     * @param {number} x - X坐标
     * @param {number} y - Y坐标
     */
    public lineTo(x: number, y: number): void {
        const _x = Math.round(x * window.devicePixelRatio);
        const _y = Math.round(y * window.devicePixelRatio);
        this.ctx.lineTo(_x, _y);
    }

    public offscreen(width: number, height: number): {
        offscreenCanvas: OffscreenCanvas,
        offscreenCtx: OffscreenCanvasRenderingContext2D 
    } {
        const offscreenCanvas = new OffscreenCanvas(width, height);
        const offscreenCtx = offscreenCanvas.getContext('2d');
        offscreenCtx.strokeStyle = GRID_LINE_COLOR;
        offscreenCtx.lineWidth = 1;
        offscreenCtx.font = `${Math.round(DEFAULT_FONT_SIZE * this.data.zoom * window.devicePixelRatio)}px ${DEFAULT_FONT_FAMILY}`;
        offscreenCtx.fillStyle = TEXT_COLOR;
        offscreenCtx.textBaseline = 'middle';
        offscreenCtx.textAlign = 'center';
        return { offscreenCanvas, offscreenCtx };
    }
    
       /**
     * 获取鼠标位置对应的单元格信息
     * @param {number} clientX - 鼠标X坐标（从0开始）
     * @param {number} clientY - 鼠标Y坐标（从0开始）
     * @returns { x: number, y: number } 包含鼠标位置的对象
     */
    public getGeometry(clientX: number, clientY: number): { x: number, y: number } {
        const rect = this.canvas.getBoundingClientRect();
        const x = clientX - rect.left;
        const y = clientY - rect.top;
        return { x, y };
    }

    /**
     * 获取选中的列和行
     * @returns {Object} - 包含sheet页面对应的起始列、结束列、起始行、结束行的对象
     */
    public getSelectedColsAndRows(): { startCol: number, endCol: number, startRow: number, endRow: number } {
        if (!this.data.selection) return { startCol: 1, endCol: 1, startRow: 1, endRow: 1 };
        const [start, end] = this.data.selection ? this.data.selection.split(':') : [this.data.activedCell, this.data.activedCell];
        const { col: startCol, row: startRow } = this.data.getCellColAndRow(start);
        const { col: endCol, row: endRow } = this.data.getCellColAndRow(end);
        // 行/列取消隐藏后尾部空行/列会被收缩删除，选区可能超出当前数组范围；
        // clamp 到有效范围，防止绘制选中框/高亮时 getAt(...).left/.top 读取 undefined
        const maxCol = Math.max(1, this.data.colHeaders.length);
        const maxRow = Math.max(1, this.data.rowHeaders.length);
        return {
            startCol: Math.min(startCol, maxCol),
            endCol: Math.min(endCol, maxCol),
            startRow: Math.min(startRow, maxRow),
            endRow: Math.min(endRow, maxRow),
        };
    }

    /**
     * 初始化右键菜单
     *
     * 在 canvas 上监听 contextmenu 事件：
     * 1. 阻止浏览器默认菜单，解析鼠标位置对应的行列（行头画布上 x<0 命中 IN_ROW_HEADER，
     *    列头画布上 y<0 命中 IN_COL_HEADER，由 getSheetCell 的跳过隐藏逻辑保证）；
     * 2. 行头/列头右键 → 右键的行/列已在对应的整行/整列选区内时保持现有选区，
     *    在选区外时选区设为该整行/整列；单元格右键 → 选区内保持选区、选区外选中该单元格；
     * 3. 构建 MenuContent（删除/隐藏/取消隐藏/行高/分割线/清除内容/设置单元格格式）
     *    并通过 Menu.openContextMenu 在鼠标位置弹出。
     *
     * 子类在构造完成后调用本方法启用右键菜单；
     * 动作执行后的重绘通过重写 {@link refreshAfterContextMenu} 实现。
     */
    protected setupContextMenu(): void {
        this.canvas.addEventListener('contextmenu', (e) => {
            e.preventDefault();
            e.stopPropagation();
            const menu = this.menu;
            if (!menu) return;

            const { x, y } = this.getGeometry(e.clientX, e.clientY);
            const { location, col, row } = this.resolveContextMenuCell(x, y);

            // 按右键位置确定选区：行头 → 整行；列头 → 整列；
            // 右键的行/列已落在对应的整行/整列选区内时保持现有选区（不重新选择），在选区外时重新选择该行/列；
            // 单元格 → 在现有选区内保持选区（便于对整块操作），在选区外选中右键位置的单元格
            const lastColName = this.data.getColName(this.data.colHeaders.length);
            const lastRow = this.data.rowHeaders.length;
            const { startCol, endCol, startRow, endRow } = this.getSelectedColsAndRows();
            if (location === MouseLocation.IN_ROW_HEADER) {
                this._contextTarget = 'row';
                // 现有选区为整行选区且右键的行在其中：保持选区（不重新选择）
                if (!(this.data.selection && startCol === 1 && endCol === this.data.colHeaders.length && row >= startRow && row <= endRow)) {
                    this.data.selection = `A${row}:${lastColName}${row}`;
                }
            } else if (location === MouseLocation.IN_COL_HEADER) {
                this._contextTarget = 'col';
                const colName = this.data.getColName(col);
                // 现有选区为整列选区且右键的列在其中：保持选区（不重新选择）
                if (!(this.data.selection && startRow === 1 && endRow === this.data.rowHeaders.length && col >= startCol && col <= endCol)) {
                    this.data.selection = `${colName}1:${colName}${lastRow}`;
                }
            } else {
                this._contextTarget = 'cell';
                if (col < startCol || col > endCol || row < startRow || row > endRow) {
                    // 右键位置在选区外：选中该单元格（selection setter 触发 SELECTION_CHANGED 联动重绘）
                    const cellName = this.data.getCellName(col, row);
                    this.data.activedCell = cellName;
                    this.data.selection = `${cellName}:${cellName}`;
                }
            }
            const { container } = this.menuContent.getElements();
            menu.openContextMenu(e.clientX, e.clientY, container);
        });
    }

    /**
     * 初始化右键菜单内容
     * @private
     * @returns {void} - 无返回值，仅初始化属性 menuContent
     */
    private setupMenuContent(): void {
        this.menuContent = new MenuContent({
            items: [
                { todo: 'undo', text: '撤销' },
                { todo: 'redo', text: '重做' },
                'separator',
                { todo: 'delete', text: '删除' },
                'separator',
                { todo: 'hide', text: '隐藏' },
                { todo: 'unhide', text: '取消隐藏' },
                { todo: 'rowHeight', text: '行高' },
                'separator',
                { todo: 'clearContent', text: '清除内容' },
                { todo: 'format', text: '设置单元格格式' },
            ],
            onClick: (todo: string) => {
                this.menu.closeMenu();
                // 撤销/重做：恢复由 DataCollection._undoRestoreCallback 完成，
                // 恢复期间广播的 VALUES_CHANGED/SELECTION_CHANGED 等事件会联动重绘主画布，
                // 此处再经 refreshAfterContextMenu 重绘当前画布（行头/列头）。
                if (todo === 'undo' || todo === 'redo') {
                    const canApply = todo === 'undo' ? this.data.undoManager.canUndo() : this.data.undoManager.canRedo();
                    if (canApply) {
                        todo === 'undo' ? this.data.undo() : this.data.redo();
                        this.refreshAfterContextMenu();
                    }
                    return;
                }
                this.handleContextMenuAction(todo);
            },
        });
    }

       /**
     * 解析右键位置的单元格信息
     *
     * 基类默认使用 getSheetCell（sheet 内容区坐标系：x<0 视为行头、y<0 视为列头）。
     * 行头/列头画布上右键时 getGeometry 返回的 x/y 恒为正值（落在各自表头宽度内），
     * 无法命中 IN_ROW_HEADER/IN_COL_HEADER 分支，
     * 因此行头/列头画布需重写本方法强制走对应分支。
     * @param {number} x - 相对当前画布的X坐标
     * @param {number} y - 相对当前画布的Y坐标
     * @returns {{ location: MouseLocation, col: number, row: number }} 单元格信息
     */
    protected resolveContextMenuCell(x: number, y: number): { location: MouseLocation, col: number, row: number } {
        return this.data.getSheetCell(x, y);
    }

    /**
     * 处理右键菜单动作
     *
     * 动作对象由 {@link _contextTarget} 决定（行头/列头/单元格），范围取当前选区：
     * - 删除：行头删行、列头删列、单元格清除内容；删除后选区保持原位置（越界收缩）；
     * - 隐藏：隐藏选区行/列（单元格上无隐藏语义，忽略）；
     * - 取消隐藏：取消选区范围内隐藏的行/列（非全表）；
     * - 行高：prompt 输入行高，应用到选区内所有行；
     * - 分割线：切换网格线显示（showGridLines）；
     * - 清除内容：清空选区内容（保留样式）；
     * - 设置单元格格式：prompt 输入背景颜色，应用到选区内所有单元格。
     *
     * 所有变更操作均通过 data.runWithFullStateUndo 包裹：
     * 操作前捕获全量状态快照（values + 行/列表头 + 网格线 + 选区），
     * 操作后推入撤销栈，使右键菜单的每项功能都可整体撤销/重做。
     * @param {string} todo - 动作名称
     */
    protected handleContextMenuAction(todo: string): void {
        const data = this.data;
        const { startCol, endCol, startRow, endRow } = this.getSelectedColsAndRows();
        switch (todo) {
            case 'delete': {
                if (this._contextTarget === 'row') {
                    data.runWithFullStateUndo('删除行', () => {
                        data.deleteRows(startRow, endRow);
                        // 删除后选区保持原位置（下方行上移补位），越界部分收缩到有效范围
                        const newEndRow = Math.min(endRow, data.rowHeaders.length);
                        const newStartRow = Math.min(startRow, newEndRow);
                        data.selection = `A${newStartRow}:${data.getColName(data.colHeaders.length)}${newEndRow}`;
                    });
                } else if (this._contextTarget === 'col') {
                    data.runWithFullStateUndo('删除列', () => {
                        data.deleteCols(startCol, endCol);
                        // 删除后选区保持原位置（右侧列左移补位），越界部分收缩到有效范围
                        const newEndCol = Math.min(endCol, data.colHeaders.length);
                        const newStartCol = Math.min(startCol, newEndCol);
                        data.selection = `${data.getColName(newStartCol)}1:${data.getColName(newEndCol)}${data.rowHeaders.length}`;
                    });
                } else {
                    // 单元格清除内容：选区保持原位置不动
                    data.runWithFullStateUndo('删除内容', () => {
                        data.clearContents(startCol, endCol, startRow, endRow);
                    });
                }
                break;
            }
            case 'hide':
                if (this._contextTarget === 'row') {
                    data.runWithFullStateUndo('隐藏行', () => data.setRowsHidden(startRow, endRow, true));
                } else if (this._contextTarget === 'col') {
                    data.runWithFullStateUndo('隐藏列', () => data.setColsHidden(startCol, endCol, true));
                } else {
                    return; // 单元格上无隐藏语义
                }
                break;
            case 'unhide':
                // 只取消当前选区行/列范围内的隐藏行/列（而非全表）：
                // 选区范围由 selection 字符串直接解析，包含其中被隐藏的行/列
                if (this._contextTarget === 'row') {
                    data.runWithFullStateUndo('取消隐藏行', () => data.setRowsHidden(startRow, endRow, false));
                } else if (this._contextTarget === 'col') {
                    data.runWithFullStateUndo('取消隐藏列', () => data.setColsHidden(startCol, endCol, false));
                } else {
                    return;
                }
                break;
            case 'rowHeight': {
                const current = data.rowHeaders.getAt(startRow - 1)?.height || 20;
                const input = window.prompt('请输入行高（像素）：', String(Math.round(current)));
                if (input === null) return;
                const height = Number(input);
                if (!Number.isFinite(height) || height <= 0) return;
                data.runWithFullStateUndo('调整行高', () => {
                    for (let r = startRow; r <= endRow; r++) {
                        data.adjustRowHeight(r, height);
                    }
                });
                break;
            }
            case 'gridLines':
                data.runWithFullStateUndo('切换分割线', () => {
                    data.showGridLines = !data.showGridLines;
                });
                break;
            case 'clearContent':
                data.runWithFullStateUndo('清除内容', () => {
                    data.clearContents(startCol, endCol, startRow, endRow);
                });
                break;
            case 'format': {
                const anchorCell = this.data.getCellName(startCol, startRow);
                const current = data.values.find(v => v.cell === anchorCell)?.backgroundColor || '#ffffff';
                const color = window.prompt('请输入背景颜色（如 #ff0000）：', current);
                if (color === null || color.trim() === '') return;
                data.runWithFullStateUndo('设置单元格格式', () => {
                    data.setCellsBackgroundColor(startCol, endCol, startRow, endRow, color.trim());
                });
                break;
            }
        }
        this.refreshAfterContextMenu();
    }

    /**
     * 右键菜单动作执行后的刷新钩子
     *
     * 基类默认无操作，子类重写以触发自身重绘（如调用各自的 draw 方法）。
     */
    protected refreshAfterContextMenu(): void {
    }

    /**
     * 滚动到指定单元格，确保单元格完全可见
     * @param {number} col - 单元格列索引
     * @param {number} row - 单元格行索引
     */
    public scrollCellIntoView(col: number, row: number): void {
        if (col < 0 || row < 0) {
            return;
        }

        const { left, top, width, height } = this.data.calculateCellRect(col, row);
        const { sheetWidth, sheetHeight } = this.data.visibleView;

        let needScroll = false;
        let newOffsetX = this.data.colHeaders.offsetWidth;
        let newOffsetY = this.data.rowHeaders.offsetHeight;

        const totalWidth = this.data.colHeaders.allColWidth;
        const totalHeight = this.data.rowHeaders.allRowHeight;

        const visibleWidth = Math.round(sheetWidth / this.data.zoom);
        const visibleHeight = Math.round(sheetHeight / this.data.zoom);

        const originalLeft = left + this.data.colHeaders.offsetWidth;
        const originalTop = top + this.data.rowHeaders.offsetHeight;

        if (left < 0) {
            newOffsetX = Math.max(0, originalLeft);
            needScroll = true;
        } else if (left + width > visibleWidth) {
            const requiredOffset = originalLeft + width - visibleWidth;
            newOffsetX = Math.min(totalWidth - visibleWidth, requiredOffset);
            newOffsetX = Math.max(0, newOffsetX);
            needScroll = true;
        }

        if (top < 0) {
            newOffsetY = Math.max(0, originalTop);
            needScroll = true;
        } else if (top + height > visibleHeight) {
            const requiredOffset = originalTop + height - visibleHeight;
            newOffsetY = Math.min(totalHeight - visibleHeight, requiredOffset);
            newOffsetY = Math.max(0, newOffsetY);
            needScroll = true;
        }

        if (needScroll) {
            this.data.colHeaders.offsetWidth = newOffsetX;
            this.data.rowHeaders.offsetHeight = newOffsetY;
        }
    }

    /**
     * 更新鼠标光标样式
     * @param {string} location - 鼠标位置，'inCell'、'inColHeader'、'inColSplit'、'inRowHeader'、'inRowSplit'、'inBrand'、'inAnchor'
     */
    public updateCursorStyle(location: string): void {
        switch(location) {
            case MouseLocation.IN_CELL:
                this.canvas.style.cursor = `url("${crossEmpty}") 8 8, auto`;
                break;
            case MouseLocation.IN_COL_HEADER:
                this.canvas.style.cursor = `url("${downFill}") 8 4, auto`;
                break;
            case MouseLocation.IN_COL_HEADER_SPLIT:
                this.canvas.style.cursor = 'col-resize';
                break;
            case MouseLocation.IN_ROW_HEADER:
                this.canvas.style.cursor = `url("${forwardFill}") 8 4, auto`;
                break;
            case MouseLocation.IN_ROW_HEADER_SPLIT:
                this.canvas.style.cursor = 'row-resize';
                break;
            case MouseLocation.IN_ANCHOR:
                this.canvas.style.cursor = `url("${crosshair}") 8 8, auto`;
                break;
        }
    }
}