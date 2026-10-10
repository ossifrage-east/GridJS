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
import { CellContent, LineText, CellBorderStyle } from "../dataArchitecture/cellContent";
import { EventEmitter } from "../../utils/eventEmitter";
import { CELL_PADDING, DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE, GRID_LINE_COLOR, KEYWORDS, MouseLocation, TEXT_COLOR } from "../constant";
import { DataCollection } from "../dataArchitecture/dataCollection";
import { Menu, MenuContent, createCompositeIcon } from "../common/menu";
import { Dialog } from "../common/dialog";
import { createDiv, createInput, createButton } from "../../utils/dom";
import { getCellDisplayText, formatNumberValue, isDateTimeCode, formatDateTimeValue } from "../../utils/numberFormat";


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
        const cellObj = this.data.values[index];
        const displayValue = getCellDisplayText(value, cellObj?.numberFormat);
        const fontSize = Math.round((this.data.values[index].fontSize || DEFAULT_FONT_SIZE) 
            * this.data.zoom);
        const letterSpacing = this.data.values[index].letterSpacing || 0;
        const wrap = this.data.values[index].wrap || false;
        const lineSpacing = this.data.values[index].lineSpacing || 0;
        
        const paragraphs = displayValue.split('\n');
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
        // 若单元格设置了数字格式，则从格式化后的显示文本中取字符进行测量
        let char = charObj?.char;
        const numberFormat = cellObj?.numberFormat;
        if (numberFormat && numberFormat !== 'General' && numberFormat !== '@') {
            const rawText = this.data.getCellValue(cell);
            const displayText = getCellDisplayText(rawText, numberFormat);
            char = displayText[charIndex];
        }
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
        // 横向图标按钮行：复制 / 剪切 / 粘贴 / 只粘贴文本 / 只粘贴格式
        const createClipboardRow = (): HTMLDivElement => {
            const row = createDiv({
                style: { display: 'flex', flexDirection: 'row', padding: '4px 2px', gap: '1px' }
            });
            const icons = [
                { todo: 'copy', icon: 'icon-copy', title: '复制 (Ctrl+C)' },
                { todo: 'cut', icon: 'icon-scissors', title: '剪切 (Ctrl+X)' },
                { todo: 'paste', icon: 'icon-paste', title: '粘贴 (Ctrl+V)' },
                { todo: 'pasteText', icon: 'icon-paste', badge: 'T' as const, title: '只粘贴文本' },
                { todo: 'pasteFormat', icon: 'icon-paste', badge: 'brush' as const, title: '只粘贴格式' },
            ];
            for (const it of icons) {
                const btn = createButton({
                    attributes: { title: it.title },
                    style: {
                        width: '28px', height: '28px', padding: '0',
                        display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                        border: 'none', borderRadius: '3px', backgroundColor: 'transparent',
                        cursor: 'pointer', fontSize: '14px', color: '#333',
                        transition: 'background-color 0.15s',
                    }
                });
                // 用 i 图标而非 emoji——icon css 已通过 @font-face 加载
                if (it.badge) {
                    btn.appendChild(createCompositeIcon(it.icon, it.badge));
                } else {
                    const iconEl = document.createElement('i');
                    iconEl.className = it.icon;
                    btn.appendChild(iconEl);
                }
                btn.addEventListener('mouseenter', () => { btn.style.backgroundColor = '#e9ecef'; });
                btn.addEventListener('mouseleave', () => { btn.style.backgroundColor = 'transparent'; });
                btn.addEventListener('click', (e) => {
                    e.preventDefault(); e.stopPropagation();
                    this.menu.closeMenu();
                    this.handleContextMenuAction(it.todo);
                });
                row.appendChild(btn);
            }
            return row;
        };

        this.menuContent = new MenuContent({
            items: [
                // 横向剪贴板按钮行（自定义元素，内联 5 个图标按钮）
                createClipboardRow,
                'separator',
                // 删除：hover 弹出二级子菜单（独立容器 append 到 body，避免主菜单 overflow 裁剪）
                () => {
                    const wrap = createDiv({
                        className: 'menu-item menu-item-has-submenu',
                        style: { position: 'relative' }
                    });
                    const iconEl = document.createElement('i');
                    iconEl.className = 'icon-bin';
                    const textSpan = document.createElement('span');
                    textSpan.className = 'text';
                    textSpan.textContent = '删除';
                    const arrowSpan = document.createElement('span');
                    arrowSpan.textContent = '▶';
                    arrowSpan.style.cssText = 'margin-left:auto; font-size:10px; color:#999;';
                    wrap.appendChild(iconEl);
                    wrap.appendChild(textSpan);
                    wrap.appendChild(arrowSpan);

                    // 子菜单容器引用（用于 hover 关闭 + 主菜单关闭时清理）
                    let submenuEl: HTMLElement | null = null;
                    // mouseleave 延迟关闭定时器：给鼠标从「删除项」移动到子菜单（穿过 gap）留时间
                    let closeTimer: ReturnType<typeof setTimeout> | null = null;
                    const cancelPendingClose = (): void => {
                        if (closeTimer !== null) {
                            clearTimeout(closeTimer);
                            closeTimer = null;
                        }
                    };

                    // hover 展开：在 body 上创建独立 MenuContent，定位到 delete 项右侧
                    wrap.addEventListener('mouseenter', () => {
                        cancelPendingClose(); // 快速来回移动时取消未生效的关闭
                        if (submenuEl) return; // 已展开则跳过
                        // 构建子菜单内容（复用 MenuContent 统一渲染，样式一致）
                        const subItems = [
                            { todo: 'deleteShiftLeft', icon: 'icon-text-indent-right', text: '单元格左移' },
                            { todo: 'deleteShiftUp', icon: 'icon-align-top', text: '单元格上移' },
                            { todo: 'deleteRow', icon: 'icon-bin', text: '删除整行' },
                            { todo: 'deleteCol', icon: 'icon-bin', text: '删除整列' },
                        ];
                        const subContent = new MenuContent({
                            items: subItems,
                            onClick: (todo: string) => {
                                closeSubmenu();
                                this.menu.closeMenu();
                                this.handleContextMenuAction(todo);
                            },
                        });
                        submenuEl = subContent.getElements().container;
                        submenuEl.classList.add('menu');
                        // 强制弹出样式：必须内联覆盖 .menu 类的 opacity:0 / visibility:hidden /
                        // transition 0.2s（否则子菜单要等淡入动画、感知为延迟显示）
                        submenuEl.style.cssText = 'position:absolute; display:flex; opacity:1; visibility:visible; transition:none; min-width:160px; box-shadow:0 4px 8px rgba(0,0,0,0.15); z-index:1002;';
                        document.body.appendChild(submenuEl);
                        // 定位：delete 项右边缘（重叠 1px 消除 hover gap，防止 mouseleave 误关）+ 顶部对齐
                        const wrapRect = wrap.getBoundingClientRect();
                        const subRect = submenuEl.getBoundingClientRect();
                        let left = wrapRect.right - 1;
                        let top = wrapRect.top;
                        if (left + subRect.width > window.innerWidth) {
                            // 右侧空间不足：改弹到左侧（同样重叠 1px）
                            left = wrapRect.left - subRect.width + 1;
                        }
                        if (top + subRect.height > window.innerHeight) {
                            top = window.innerHeight - subRect.height - 4;
                        }
                        if (top < 0) top = 4;
                        submenuEl.style.left = `${left}px`;
                        submenuEl.style.top = `${top}px`;
                        // 鼠标移入子菜单时取消待执行的关闭
                        submenuEl.addEventListener('mouseenter', cancelPendingClose);
                    });

                    // 关闭子菜单（从 body 移除 DOM）
                    const closeSubmenu = (): void => {
                        cancelPendingClose();
                        if (submenuEl && submenuEl.isConnected) {
                            submenuEl.remove();
                        }
                        submenuEl = null;
                    };
                    // 主菜单关闭时清理子菜单引用（兜底：防止 mouseleave 期间子菜单残留）
                    this.menu.on('close', closeSubmenu);
                    // mouseleave 延迟 120ms 关闭：鼠标穿过 wrap 与子菜单间的缝隙或子菜单内移动时不会误关
                    wrap.addEventListener('mouseleave', () => {
                        cancelPendingClose();
                        closeTimer = setTimeout(closeSubmenu, 120);
                    });
                    return wrap;
                },
                // 插入：hover 弹出二级子菜单（与删除子菜单同构，独立容器 append 到 body）
                () => {
                    const wrap = createDiv({
                        className: 'menu-item menu-item-has-submenu',
                        style: { position: 'relative' }
                    });
                    const iconEl = document.createElement('i');
                    iconEl.className = 'icon-plus';
                    const textSpan = document.createElement('span');
                    textSpan.className = 'text';
                    textSpan.textContent = '插入';
                    const arrowSpan = document.createElement('span');
                    arrowSpan.textContent = '▶';
                    arrowSpan.style.cssText = 'margin-left:auto; font-size:10px; color:#999;';
                    wrap.appendChild(iconEl);
                    wrap.appendChild(textSpan);
                    wrap.appendChild(arrowSpan);

                    let submenuEl: HTMLElement | null = null;
                    let closeTimer: ReturnType<typeof setTimeout> | null = null;
                    const cancelPendingClose = (): void => {
                        if (closeTimer !== null) {
                            clearTimeout(closeTimer);
                            closeTimer = null;
                        }
                    };

                    // 构建带「份数式」数字步进器的插入项（？处填入行数/列数；样式参照打印设置份数输入框，见 menu.scss）
                    const buildCountItem = (label: string, icon: string, todo: string): HTMLElement => {
                        const item = createDiv({ className: 'menu-item', style: { padding: '6px 10px', gap: '8px' } });
                        const ic = document.createElement('i');
                        ic.className = icon;
                        const txt = document.createElement('span');
                        txt.className = 'text';
                        txt.textContent = label;
                        // 输入框 + 右侧纵向 +/− 步进按钮（与打印设置份数输入框同构，不使用独立 ▲▼）
                        const grp = document.createElement('span');
                        grp.className = 'menu-count-stepper';
                        grp.style.marginLeft = 'auto';
                        const input = createInput({ type: 'number' });
                        input.className = 'menu-count-input';
                        input.value = '1';
                        input.min = '1';
                        input.max = '999';
                        input.step = '1';
                        const clampCount = (): void => {
                            const v = parseInt(input.value);
                            input.value = String(Number.isNaN(v) ? 1 : Math.min(999, Math.max(1, v)));
                        };
                        const btns = document.createElement('span');
                        btns.className = 'menu-count-btns';
                        // 步进：单击 ±1；按住持续 ±1（初始延迟 400ms，之后每 80ms 重复），抬起/移出即停止
                        const applyStep = (delta: number): void => {
                            const v = parseInt(input.value);
                            const base = Number.isNaN(v) ? 1 : v;
                            input.value = String(Math.min(999, Math.max(1, base + delta)));
                        };
                        const startRepeat = (delta: number): void => {
                            applyStep(delta);
                            let timer: ReturnType<typeof setTimeout> | null = null;
                            const tick = () => {
                                applyStep(delta);
                                timer = setTimeout(tick, 80);
                            };
                            timer = setTimeout(tick, 400);
                            // 在 window 上监听抬起/取消：即使指针移出按钮也能停止连发
                            const stop = () => {
                                if (timer !== null) {
                                    clearTimeout(timer);
                                    timer = null;
                                }
                                window.removeEventListener('pointerup', stop);
                                window.removeEventListener('pointercancel', stop);
                            };
                            window.addEventListener('pointerup', stop);
                            window.addEventListener('pointercancel', stop);
                        };
                        const bindStepButton = (button: HTMLButtonElement, delta: number): void => {
                            button.addEventListener('pointerdown', (e) => {
                                if (e.button !== 0) return;
                                e.preventDefault(); // 阻止焦点转移与文本选择，保证按住期间连发不被打断
                                e.stopPropagation();
                                startRepeat(delta);
                            });
                            // click 仅拦截冒泡：防止行体 click 触发插入动作（步进逻辑全部走 pointerdown）
                            button.addEventListener('click', (e) => { e.preventDefault(); e.stopPropagation(); });
                        };
                        const incBtn = createButton({ textContent: '+' });
                        const decBtn = createButton({ textContent: '−' });
                        bindStepButton(incBtn, 1);
                        bindStepButton(decBtn, -1);
                        input.addEventListener('click', (e) => { e.stopPropagation(); });
                        input.addEventListener('change', clampCount);
                        btns.appendChild(incBtn);
                        btns.appendChild(decBtn);
                        grp.appendChild(input);
                        grp.appendChild(btns);
                        item.appendChild(ic);
                        item.appendChild(txt);
                        item.appendChild(grp);
                        // 点击行体（按钮/输入框已 stopPropagation）触发插入
                        item.addEventListener('click', (e) => {
                            e.preventDefault();
                            e.stopPropagation();
                            const count = Math.max(1, parseInt(input.value) || 1);
                            closeSubmenu();
                            this.menu.closeMenu();
                            this.doInsertCountAction(todo, count);
                        });
                        return item;
                    };

                    wrap.addEventListener('mouseenter', () => {
                        cancelPendingClose();
                        if (submenuEl) return;
                        const subItems: ({ todo?: string, icon?: string | null, text?: string } | 'separator' | (() => HTMLElement))[] = [
                            { todo: 'insertShiftRight', icon: 'icon-text-indent-left', text: '插入单元格，活动单元格右移' },
                            { todo: 'insertShiftDown', icon: 'icon-align-bottom', text: '插入单元格，活动单元格下移' },
                            'separator',
                            () => buildCountItem('在上方插入行', 'icon-align-top', 'insertRowsAbove'),
                            () => buildCountItem('在下方插入行', 'icon-align-bottom', 'insertRowsBelow'),
                            () => buildCountItem('在左侧插入列', 'icon-text-indent-left', 'insertColsLeft'),
                            () => buildCountItem('在右侧插入列', 'icon-text-indent-right', 'insertColsRight'),
                        ];
                        const subContent = new MenuContent({
                            items: subItems,
                            onClick: (todo: string) => {
                                closeSubmenu();
                                this.menu.closeMenu();
                                this.handleContextMenuAction(todo);
                            },
                        });
                        submenuEl = subContent.getElements().container;
                        submenuEl.classList.add('menu');
                        submenuEl.style.cssText = 'position:absolute; display:flex; opacity:1; visibility:visible; transition:none; min-width:220px; box-shadow:0 4px 8px rgba(0,0,0,0.15); z-index:1002;';
                        document.body.appendChild(submenuEl);
                        const wrapRect = wrap.getBoundingClientRect();
                        const subRect = submenuEl.getBoundingClientRect();
                        let left = wrapRect.right - 1;
                        let top = wrapRect.top;
                        if (left + subRect.width > window.innerWidth) {
                            left = wrapRect.left - subRect.width + 1;
                        }
                        if (top + subRect.height > window.innerHeight) {
                            top = window.innerHeight - subRect.height - 4;
                        }
                        if (top < 0) top = 4;
                        submenuEl.style.left = `${left}px`;
                        submenuEl.style.top = `${top}px`;
                        submenuEl.addEventListener('mouseenter', cancelPendingClose);
                    });

                    const closeSubmenu = (): void => {
                        cancelPendingClose();
                        if (submenuEl && submenuEl.isConnected) {
                            submenuEl.remove();
                        }
                        submenuEl = null;
                    };
                    this.menu.on('close', closeSubmenu);
                    wrap.addEventListener('mouseleave', () => {
                        cancelPendingClose();
                        closeTimer = setTimeout(closeSubmenu, 120);
                    });
                    return wrap;
                },
                'separator',
                { todo: 'hide', icon: 'icon-eye-blocked', text: '隐藏' },
                { todo: 'unhide', icon: 'icon-eye', text: '取消隐藏' },
                { todo: 'rowHeight', icon: 'icon-align-middle', text: '行高' },
                { todo: 'colWidth', icon: 'icon-justify', text: '列宽' },
                'separator',
                { todo: 'clearContent', icon: 'icon-cross', text: '清除内容' },
                { todo: 'format', icon: 'icon-hammer', text: '设置单元格格式' },
            ],
            onClick: (todo: string) => {
                this.menu.closeMenu();
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
     * - 行高：弹出行高对话框（Bootstrap 风格），确认后应用到选区内所有行；
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
            case 'copy':
                data.copySelection();
                break;
            case 'cut':
                data.runWithFullStateUndo('剪切', () => data.cutSelection());
                break;
            case 'paste':
                data.runWithFullStateUndo('粘贴', () => data.pasteSelection());
                break;
            case 'pasteText':
                data.runWithFullStateUndo('只粘贴文本', () => data.pasteTextOnly());
                break;
            case 'pasteFormat':
                data.runWithFullStateUndo('只粘贴格式', () => data.pasteFormatOnly());
                break;
            case 'deleteShiftLeft': {
                // 清除选区内容 + 右侧可见单元格向左紧凑补位（单元格级，非整列）
                data.runWithFullStateUndo('删除并左移', () => {
                    data.shiftCellsLeft(startCol, endCol, startRow, endRow);
                });
                break;
            }
            case 'deleteShiftUp': {
                // 清除选区内容 + 下方可见单元格向上紧凑补位（单元格级，非整行）
                data.runWithFullStateUndo('删除并上移', () => {
                    data.shiftCellsUp(startCol, endCol, startRow, endRow);
                });
                break;
            }
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
            case 'deleteRow': {
                // 二级菜单「删除整行」：不依赖右键目标（行头/单元格均可），删除选区覆盖的整行
                data.runWithFullStateUndo('删除行', () => {
                    data.deleteRows(startRow, endRow);
                    // 删除后选区保持原位置（下方行上移补位），越界部分收缩到有效范围
                    const newEndRow = Math.min(endRow, data.rowHeaders.length);
                    const newStartRow = Math.min(startRow, newEndRow);
                    data.selection = `A${newStartRow}:${data.getColName(data.colHeaders.length)}${newEndRow}`;
                });
                break;
            }
            case 'deleteCol': {
                // 二级菜单「删除整列」：不依赖右键目标（列头/单元格均可），删除选区覆盖的整列
                data.runWithFullStateUndo('删除列', () => {
                    data.deleteCols(startCol, endCol);
                    // 删除后选区保持原位置（右侧列左移补位），越界部分收缩到有效范围
                    const newEndCol = Math.min(endCol, data.colHeaders.length);
                    const newStartCol = Math.min(startCol, newEndCol);
                    data.selection = `${data.getColName(newStartCol)}1:${data.getColName(newEndCol)}${data.rowHeaders.length}`;
                });
                break;
            }
            case 'insertShiftRight': {
                // 选区位置插入空白单元格 + 右侧可见单元格右移让位（单元格级）
                data.runWithFullStateUndo('插入并右移', () => {
                    data.insertCellsRight(startCol, endCol, startRow, endRow);
                });
                break;
            }
            case 'insertShiftDown': {
                // 选区位置插入空白单元格 + 下方可见单元格下移让位（单元格级）
                data.runWithFullStateUndo('插入并下移', () => {
                    data.insertCellsDown(startCol, endCol, startRow, endRow);
                });
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
            case 'rowHeight':
                // 弹出行高对话框：确认后经 runWithFullStateUndo 应用并重绘（见 openRowHeightDialog）
                this.openRowHeightDialog(startRow, endRow);
                return;
            case 'colWidth':
                // 弹出列宽对话框：镜像 openRowHeightDialog，经 runWithFullStateUndo 应用并重绘
                this.openColWidthDialog(startCol, endCol);
                return;
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
                // 弹出 Excel 风格的「设置单元格格式」对话框（数字 / 边框两个 Tab）
                this.openFormatCellsDialog(startCol, endCol, startRow, endRow);
                return;
            }
        }
        this.refreshAfterContextMenu();
    }

    /**
     * 处理带数量参数的插入动作（上方/下方插入行、左侧/右侧插入列）。
     *
     * 数量来自二级菜单中各输入框的当前值；操作经 runWithFullStateUndo 包裹（选区变更一并进入快照），
     * 操作后将选区重置为新插入的行/列范围（Excel 习惯：选中刚插入的部分以便后续操作）。
     * @param {string} todo - 动作名称（insertRowsAbove/insertRowsBelow/insertColsLeft/insertColsRight）
     * @param {number} count - 插入数量（调用前已确保 >= 1，此处再次钳制兜底）
     * @private
     */
    private doInsertCountAction(todo: string, count: number): void {
        const data = this.data;
        const { startCol, endCol, startRow, endRow } = this.getSelectedColsAndRows();
        const safeCount = Math.max(1, Math.floor(count));
        switch (todo) {
            case 'insertRowsAbove':
                data.runWithFullStateUndo('上方插入行', () => {
                    data.insertRows(startRow, safeCount);
                    data.selection = `A${startRow}:${data.getColName(data.colHeaders.length)}${startRow + safeCount - 1}`;
                });
                break;
            case 'insertRowsBelow':
                data.runWithFullStateUndo('下方插入行', () => {
                    data.insertRows(endRow + 1, safeCount);
                    data.selection = `A${endRow + 1}:${data.getColName(data.colHeaders.length)}${endRow + safeCount}`;
                });
                break;
            case 'insertColsLeft':
                data.runWithFullStateUndo('左侧插入列', () => {
                    data.insertCols(startCol, safeCount);
                    data.selection = `${data.getColName(startCol)}1:${data.getColName(startCol + safeCount - 1)}${data.rowHeaders.length}`;
                });
                break;
            case 'insertColsRight':
                data.runWithFullStateUndo('右侧插入列', () => {
                    data.insertCols(endCol + 1, safeCount);
                    data.selection = `${data.getColName(endCol + 1)}1:${data.getColName(endCol + safeCount)}${data.rowHeaders.length}`;
                });
                break;
            default:
                return;
        }
        this.refreshAfterContextMenu();
    }

    /**
     * 维度设置对话框（行高/列宽共用）
     *
     * Bootstrap input-group 风格：左侧标签 + 右侧输入框 + ▲▼ 步进按钮 + 取消/确定。
     * 校验逻辑：空/非数字/小于 minValue 视为非法，点确定时对话框保持打开不产生任何变更。
     * 交互特性：
     * - type=text + inputmode=numeric + input 事件过滤非数字字符（含全角数字），
     *   规避 type=number 不支持选区 API（selectionStart 恒 null、setSelectionRange 抛错）；
     * - ▲▼ 单击 ±1，按住 400ms 后每 80ms 重复，window pointerup/cancel 停止连发；
     * - 鼠标点击定位光标 + 拖拽选区：Chromium 怪癖守卫（非折叠选区时 preventDefault +
     *   手动定位），无选区时走原生行为；双击全选；
     * - Enter 确认，ESC 由 Dialog 文档级监听处理；
     * - 对话框关闭后自毁（destroy），避免 DOM 与事件监听累积。
     * @param opts - 配置选项
     * @private
     */
    private openDimensionDialog(opts: {
        /** 对话框标题（如「行高」「列宽」） */
        title: string;
        /** 输入框前标签文字（如「行高（像素）」） */
        labelText: string;
        /** 最小允许值，与数据层 adjust* 方法的钳制保持一致（行高/列宽均为 24） */
        minValue: number;
        /** 打开对话框时输入框的初始值 */
        currentValue: number;
        /**
         * 确认时应用输入值到选区内所有行/列。
         * 调用方内部应包 runWithFullStateUndo + 循环 adjust* + refreshAfterContextMenu。
         * @param value - 校验通过的数值（已 ≥ minValue）
         */
        applyValue: (value: number) => void;
    }): void {
        const { title, labelText, minValue, currentValue, applyValue } = opts;

        // 内容区：标签 + 数字输入框（Bootstrap form-control 视觉）
        const label = createDiv({
            textContent: labelText,
            style: { fontSize: '0.875rem', color: '#212529' }
        });
        // 输入框：自身无边框透明底，边框/圆角/聚焦高亮由外层 group 统一提供；
        // 用 text + inputmode=numeric 替代 type=number：后者不支持任何选区 API
        // （selectionStart 恒为 null、setSelectionRange 抛 InvalidStateError），
        // 无法实现鼠标点击定位光标与拖拽选区；数字约束改由 input 事件过滤保证
        const input = createInput({
            type: 'text',
            attributes: { inputmode: 'numeric', autocomplete: 'off', spellcheck: 'false' },
            style: {
                flex: '1',
                minWidth: '0',
                boxSizing: 'border-box',
                padding: '0.25rem 0.5rem',
                fontSize: '0.875rem',
                lineHeight: '1.5',
                color: '#212529',
                backgroundColor: 'transparent',
                border: 'none',
                outline: 'none',
            }
        });
        input.value = String(Math.round(currentValue));
        // 数字过滤：剥离非数字字符（含中文输入法产生的全角数字），保持等价于 type=number 的输入约束
        input.addEventListener('input', () => {
            const cleaned = input.value.replace(/\D/g, '');
            if (cleaned !== input.value) {
                input.value = cleaned;
            }
        });

        // 自定义步进：单击 ±1；按住持续 ±1（初始延迟 400ms，之后每 80ms 重复），抬起即停止
        const applyStep = (delta: number): void => {
            const value = Number(input.value);
            const base = Number.isFinite(value) ? value : 0;
            input.value = String(Math.max(minValue, base + delta));
        };
        const startRepeat = (delta: number): void => {
            applyStep(delta);
            let timer: ReturnType<typeof setTimeout> | null = null;
            const tick = () => {
                applyStep(delta);
                timer = setTimeout(tick, 80);
            };
            timer = setTimeout(tick, 400);
            // 在 window 上监听抬起/取消：即使指针移出按钮也能停止连发
            const stop = () => {
                if (timer !== null) {
                    clearTimeout(timer);
                    timer = null;
                }
                window.removeEventListener('pointerup', stop);
                window.removeEventListener('pointercancel', stop);
            };
            window.addEventListener('pointerup', stop);
            window.addEventListener('pointercancel', stop);
        };
        // 步进按钮：无独立边框，内嵌于 group 右侧，仅以细分隔线与输入框分界（Bootstrap input-group 视觉）；
        // flex: 1 1 0 让每个按钮各占列高的一半，hover 背景色因此铺满整个按钮区域
        const createSpinnerButton = (glyph: string, withDivider: boolean): HTMLButtonElement => createButton({
            textContent: glyph,
            style: {
                flex: '1 1 0',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                width: '1.4rem',
                padding: '0',
                fontSize: '0.55rem',
                lineHeight: '1',
                color: '#6c757d',
                backgroundColor: 'transparent',
                border: 'none',
                borderTop: withDivider ? '1px solid #dee2e6' : 'none',
                cursor: 'pointer',
                userSelect: 'none',
            }
        });
        const bindSpinner = (button: HTMLButtonElement, delta: number): void => {
            button.addEventListener('pointerdown', (e) => {
                if (e.button !== 0) return;
                e.preventDefault(); // 阻止焦点转移与文本选择，保证按住期间连发不被打断
                startRepeat(delta);
            });
            button.addEventListener('mouseenter', () => { button.style.backgroundColor = '#e9ecef'; });
            button.addEventListener('mouseleave', () => { button.style.backgroundColor = 'transparent'; });
        };

        const upButton = createSpinnerButton('▲', false);
        const downButton = createSpinnerButton('▼', true);
        bindSpinner(upButton, 1);
        bindSpinner(downButton, -1);
        const spinnerColumn = createDiv({
            children: [upButton, downButton],
            style: {
                display: 'flex',
                flexDirection: 'column',
                borderLeft: '1px solid #ced4da', // 输入框与步进按钮之间的分隔线
            }
        });

        // 组合容器（Bootstrap input-group）：输入框与步进按钮共用同一边框/圆角/白底，连为一体；
        // overflow hidden 让子元素随容器圆角裁切
        const group = createDiv({
            children: [input, spinnerColumn],
            style: {
                display: 'flex',
                alignItems: 'stretch',
                backgroundColor: '#fff',
                border: '1px solid #ced4da',
                borderRadius: '0.375rem',
                overflow: 'hidden',
                transition: 'border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out',
            }
        });
        // 聚焦态高亮作用到整个组合容器（对应 Bootstrap input-group 的 focus 视觉）
        input.addEventListener('focus', () => {
            group.style.borderColor = '#86b7fe';
            group.style.boxShadow = '0 0 0 0.25rem rgba(13, 110, 253, 0.25)';
        });
        input.addEventListener('blur', () => {
            group.style.borderColor = '#ced4da';
            group.style.boxShadow = 'none';
        });
        // —— 鼠标点击定位光标 + 拖拽选区 ——
        // Chromium 怪癖：输入框持有非折叠选区时（如打开对话框时的全选），mousedown 重新聚焦
        // 会「恢复原选区」而不是把光标定位到点击处（表现为全选不消选，只有方向键能移动光标）。
        // 处理策略：仅在存在选区时接管鼠标行为（preventDefault + 手动定位/拖拽选区）；
        // 无选区时完全不干预，点击定位与拖拽选择均走原生行为。
        const measureCtx = document.createElement('canvas').getContext('2d');
        /** 依据点击横坐标计算输入框内最近的字符边界索引（镜像 canvas 测量字符宽度） */
        const caretIndexFromClientX = (clientX: number): number => {
            const rect = input.getBoundingClientRect();
            const style = window.getComputedStyle(input);
            measureCtx.font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
            const x = clientX - rect.left - parseFloat(style.paddingLeft) + input.scrollLeft;
            let acc = 0;
            for (let i = 0; i < input.value.length; i++) {
                const w = measureCtx.measureText(input.value[i]).width;
                if (x < acc + w / 2) return i; // 落在字符前半段则定位到该字符之前
                acc += w;
            }
            return input.value.length;
        };
        let selectionAnchor = 0; // 拖拽起点（按下时的光标索引）
        const handleSelectionDrag = (e: MouseEvent): void => {
            const idx = caretIndexFromClientX(e.clientX);
            input.setSelectionRange(Math.min(selectionAnchor, idx), Math.max(selectionAnchor, idx));
        };
        const handleSelectionEnd = (): void => {
            window.removeEventListener('mousemove', handleSelectionDrag);
            window.removeEventListener('mouseup', handleSelectionEnd);
        };
        input.addEventListener('mousedown', (e) => {
            if (e.button !== 0) return;
            if (input.selectionStart === input.selectionEnd) return; // 无选区：原生行为即可
            e.preventDefault(); // 阻止 Chromium 恢复原选区，改由下方手动定位
            input.focus();
            selectionAnchor = caretIndexFromClientX(e.clientX);
            input.setSelectionRange(selectionAnchor, selectionAnchor);
            window.addEventListener('mousemove', handleSelectionDrag);
            window.addEventListener('mouseup', handleSelectionEnd);
        });
        // 双击全选（mousedown 被接管后原生双击选词不再生效，此处显式补齐该惯例行为）
        input.addEventListener('dblclick', () => {
            input.setSelectionRange(0, input.value.length);
        });

        const wrapper = createDiv({
            children: [label, group],
            style: { display: 'flex', flexDirection: 'column', gap: '0.5rem' }
        });

        // 校验并应用输入值；非法输入返回 false（由调用方决定是否保持对话框打开）
        const apply = (): boolean => {
            const value = input.value.trim();
            const num = Number(value);
            if (value === '' || !Number.isFinite(num) || num < minValue) return false;
            applyValue(num);
            return true;
        };

        const dialog = new Dialog({
            title,
            content: wrapper,
            size: 'sm',
            buttons: [
                { text: '取消', variant: 'secondary' },
                {
                    text: '确定',
                    variant: 'primary',
                    close: false, // 校验通过才手动关闭，非法输入时保持打开
                    onClick: (d) => {
                        if (apply()) d.hide();
                    }
                },
            ],
            onClose: () => dialog.destroy()
        });
        dialog.show();
        // 打开后聚焦并全选输入内容，便于直接键入新值；Enter 确认，ESC 由 Dialog 的文档级监听处理
        input.focus();
        input.select();
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                if (apply()) dialog.hide();
            }
        });
    }

    /**
     * 弹出行高设置对话框（委托 openDimensionDialog）
     * @param {number} startRow - 选区起始行（1-based）
     * @param {number} endRow - 选区结束行（1-based）
     * @private
     */
    private openRowHeightDialog(startRow: number, endRow: number): void {
        const data = this.data;
        const MIN_ROW_HEIGHT = 24; // 最小行高，与 DataCollection.adjustRowHeight 的钳制及行表头拖拽调高保持一致
        this.openDimensionDialog({
            title: '行高',
            labelText: '行高（像素）',
            minValue: MIN_ROW_HEIGHT,
            currentValue: data.rowHeaders.getAt(startRow - 1)?.height || MIN_ROW_HEIGHT,
            applyValue: (height) => {
                data.runWithFullStateUndo('调整行高', () => {
                    for (let r = startRow; r <= endRow; r++) {
                        data.adjustRowHeight(r, height);
                    }
                });
                this.refreshAfterContextMenu();
            }
        });
    }

    /**
     * 弹出列宽设置对话框（委托 openDimensionDialog）
     * @param {number} startCol - 选区起始列（1-based）
     * @param {number} endCol - 选区结束列（1-based）
     * @private
     */
    private openColWidthDialog(startCol: number, endCol: number): void {
        const data = this.data;
        const MIN_COL_WIDTH = 24; // 最小列宽，与 DataCollection.adjustColumnWidth 的钳制及列表头拖拽调高保持一致
        this.openDimensionDialog({
            title: '列宽',
            labelText: '列宽（像素）',
            minValue: MIN_COL_WIDTH,
            currentValue: data.colHeaders.getAt(startCol - 1)?.width || MIN_COL_WIDTH,
            applyValue: (width) => {
                data.runWithFullStateUndo('调整列宽', () => {
                    for (let c = startCol; c <= endCol; c++) {
                        data.adjustColumnWidth(c, width);
                    }
                });
                this.refreshAfterContextMenu();
            }
        });
    }

    /**
     * 设置单元格格式对话框（数字 / 边框两个 Tab）
     *
     * 参照 Excel 的「设置单元格格式」对话框，顶部两个导航 Tab：
     * - 数字：分类列表（常规/数值/货币/会计专用/百分比/科学记数/文本，日期/时间/分数列出但简化）+
     *   对应控件（小数位数/千分位/负数样式/货币符号）+ 示例预览
     * - 边框：预设按钮 + 线条样式 + 颜色 + 位置按钮
     * 确定后经 runWithFullStateUndo 写入选区，可撤销/重做。
     *
     * @param {number} startCol - 选区起始列（1-based，含）
     * @param {number} endCol - 选区结束列（1-based，含）
     * @param {number} startRow - 选区起始行（1-based，含）
     * @param {number} endRow - 选区结束行（1-based，含）
     * @private
     */
    private openFormatCellsDialog(startCol: number, endCol: number, startRow: number, endRow: number): void {
        const data = this.data;
        const anchorCellName = data.getCellName(startCol, startRow);
        const anchorCell = data.values.find(v => v.cell === anchorCellName);
        const currentFormat = anchorCell?.numberFormat;
        const currentBorderColor = anchorCell?.borderColor || '#000000';

        // —— 对话框状态（闭包变量）——
        let numberCategory = 'general';
        let decimalPlaces = 2;
        let useThousands = false;
        let negativeStyle = 0; // 0:-1234.10  1:(1234.10)  2:-1234.10红  3:(1234.10)红
        let currencySymbol = '¥';
        let lineStyle = 1;   // 1: 细实线  2: 粗实线
        let borderColor = currentBorderColor;
        let borderAction = '';  // '', 'none', 'outer', 'all', 'top', 'bottom', 'left', 'right'
        let dateFormatCode = 'yyyy/m/d';  // 日期分类当前选中的格式码
        let timeFormatCode = 'h:mm';      // 时间分类当前选中的格式码

        // 日期/时间预设格式码列表（格式码 + 示例预览标签）
        const datePresets: { code: string; label: string }[] = [
            { code: 'yyyy/m/d', label: '2026/10/10' },
            { code: 'yyyy-mm-dd', label: '2026-10-10' },
            { code: 'yyyy"年"m"月"d"日"', label: '2026年10月10日' },
            { code: 'm/d/yyyy', label: '10/10/2026' },
            { code: 'dddd, yyyy"年"m"月"d"日"', label: '星期六, 2026年10月10日' },
        ];
        const timePresets: { code: string; label: string }[] = [
            { code: 'h:mm', label: '14:30' },
            { code: 'h:mm:ss', label: '14:30:45' },
            { code: 'h:mm AM/PM', label: '2:30 PM' },
            { code: 'mm:ss', label: '30:45' },
            { code: 'yyyy-mm-dd h:mm:ss', label: '2026-10-10 14:30:45' },
        ];
        // 示例日期（预览用固定值）
        const exampleDate = new Date(2026, 9, 10, 14, 30, 45);

        // 从当前格式码推断初始状态
        if (currentFormat && currentFormat !== 'General') {
            if (currentFormat === '@') {
                numberCategory = 'text';
            } else if (currentFormat.includes('%')) {
                numberCategory = 'percentage';
                const m = currentFormat.match(/\.(0+)/);
                if (m) decimalPlaces = m[1].length;
            } else if (/E/i.test(currentFormat)) {
                numberCategory = 'scientific';
                const m = currentFormat.match(/\.(0+)/);
                if (m) decimalPlaces = m[1].length;
            } else if (/^[¥$€£]/.test(currentFormat)) {
                numberCategory = 'currency';
                currencySymbol = currentFormat[0];
                useThousands = true;
                const m = currentFormat.match(/\.(0+)/);
                if (m) decimalPlaces = m[1].length;
            } else if (isDateTimeCode(currentFormat)) {
                // 日期/时间格式码：含 h/s/AM-PM → 时间分类，否则 → 日期分类
                if (/[hs]/i.test(currentFormat) || /AM\/PM/i.test(currentFormat)) {
                    numberCategory = 'time';
                    timeFormatCode = currentFormat;
                } else {
                    numberCategory = 'date';
                    dateFormatCode = currentFormat;
                }
            } else {
                numberCategory = 'number';
                useThousands = currentFormat.includes('#,##');
                if (currentFormat.includes(';(')) negativeStyle = 1;
                const m = currentFormat.match(/\.(0+)/);
                if (m) decimalPlaces = m[1].length;
            }
        }

        // —— 合成格式码 ——
        const buildFormatCode = (): string => {
            switch (numberCategory) {
                case 'general': return 'General';
                case 'text': return '@';
                case 'number': {
                    const intPart = useThousands ? '#,##0' : '0';
                    const dec = decimalPlaces > 0 ? '.' + '0'.repeat(decimalPlaces) : '';
                    const fmt = intPart + dec;
                    if (negativeStyle === 1 || negativeStyle === 3) {
                        return fmt + ';(' + fmt + ')';
                    }
                    return fmt;
                }
                case 'currency':
                case 'accounting': {
                    const dec = decimalPlaces > 0 ? '.' + '0'.repeat(decimalPlaces) : '';
                    return currencySymbol + '#,##0' + dec;
                }
                case 'percentage': {
                    const dec = decimalPlaces > 0 ? '.' + '0'.repeat(decimalPlaces) : '';
                    return '0' + dec + '%';
                }
                case 'scientific': {
                    const dec = decimalPlaces > 0 ? '.' + '0'.repeat(decimalPlaces) : '';
                    return '0' + dec + 'E+00';
                }
                case 'date': return dateFormatCode;
                case 'time': return timeFormatCode;
                default: return 'General';
            }
        };

        // —— 示例预览 ——
        const previewPos = createDiv({ style: { fontSize: '0.9rem', fontFamily: 'monospace', color: '#212529' } });
        const previewNeg = createDiv({ style: { fontSize: '0.9rem', fontFamily: 'monospace', color: '#212529' } });
        const updatePreview = (): void => {
            const code = buildFormatCode();
            if (code === 'General' || code === '@') {
                previewPos.textContent = '1234.561';
                previewNeg.textContent = '-1234.561';
                previewNeg.style.display = '';
            } else if (isDateTimeCode(code)) {
                // 日期/时间格式码：用固定示例日期预览，隐藏负数预览
                previewPos.textContent = formatDateTimeValue(exampleDate, code);
                previewNeg.style.display = 'none';
            } else {
                previewPos.textContent = formatNumberValue(1234.561, code);
                previewNeg.textContent = formatNumberValue(-1234.561, code);
                previewNeg.style.display = '';
            }
            previewNeg.style.color = (negativeStyle === 2 || negativeStyle === 3) ? '#ff0000' : '#212529';
        };

        /**
         * 构建带步进按钮的数字输入组
         *
         * 参照行高对话框（openDimensionDialog）的 Bootstrap input-group 风格：
         * 透明底输入框 + 右侧纵向 ▲/▼ 步进按钮，共用边框/圆角/聚焦高亮。
         * 交互：单击 ±1，按住 400ms 后每 80ms 重复，抬起即停止。
         *
         * @param opts.value    - 初始值
         * @param opts.min      - 最小值
         * @param opts.max      - 最大值
         * @param opts.onChange - 值变更回调（已钳制到 min..max）
         * @returns 组合容器 HTMLElement
         */
        const buildStepperInput = (opts: {
            value: number; min: number; max: number; onChange: (v: number) => void;
        }): HTMLElement => {
            const { min, max, onChange } = opts;
            // type=text + inputmode=numeric：规避 type=number 不支持选区 API
            const input = createInput({
                type: 'text',
                attributes: { inputmode: 'numeric', autocomplete: 'off', spellcheck: 'false' },
                style: {
                    flex: '1', minWidth: '0', boxSizing: 'border-box',
                    padding: '0.125rem 0.5rem', fontSize: '0.8rem', lineHeight: '1.5',
                    color: '#212529', backgroundColor: 'transparent',
                    border: 'none', outline: 'none', textAlign: 'center',
                }
            });
            input.value = String(opts.value);
            // 数字过滤（含全角数字）
            input.addEventListener('input', () => {
                const cleaned = input.value.replace(/\D/g, '');
                if (cleaned !== input.value) input.value = cleaned;
            });
            // 失焦/回车时钳制并回调
            const applyValue = (): void => {
                const v = parseInt(input.value, 10);
                const clamped = Number.isNaN(v) ? min : Math.min(max, Math.max(min, v));
                input.value = String(clamped);
                onChange(clamped);
            };
            input.addEventListener('change', applyValue);
            input.addEventListener('blur', applyValue);

            // 步进：单击 ±1；按住持续 ±1（初始延迟 400ms，之后每 80ms 重复），抬起即停止
            const applyStep = (delta: number): void => {
                const v = parseInt(input.value, 10);
                const base = Number.isNaN(v) ? min : v;
                const clamped = Math.min(max, Math.max(min, base + delta));
                input.value = String(clamped);
                onChange(clamped);
            };
            const startRepeat = (delta: number): void => {
                applyStep(delta);
                let timer: ReturnType<typeof setTimeout> | null = null;
                const tick = (): void => { applyStep(delta); timer = setTimeout(tick, 80); };
                timer = setTimeout(tick, 400);
                const stop = (): void => {
                    if (timer !== null) { clearTimeout(timer); timer = null; }
                    window.removeEventListener('pointerup', stop);
                    window.removeEventListener('pointercancel', stop);
                };
                window.addEventListener('pointerup', stop);
                window.addEventListener('pointercancel', stop);
            };
            const createSpinnerButton = (glyph: string, withDivider: boolean): HTMLButtonElement => createButton({
                textContent: glyph,
                style: {
                    flex: '1 1 0', display: 'flex', alignItems: 'center', justifyContent: 'center',
                    width: '1.2rem', padding: '0', fontSize: '0.5rem', lineHeight: '1',
                    color: '#6c757d', backgroundColor: 'transparent', border: 'none',
                    borderTop: withDivider ? '1px solid #dee2e6' : 'none',
                    cursor: 'pointer', userSelect: 'none',
                }
            });
            const bindSpinner = (button: HTMLButtonElement, delta: number): void => {
                button.addEventListener('pointerdown', (e) => {
                    if (e.button !== 0) return;
                    e.preventDefault(); // 阻止焦点转移，保证连发不被打断
                    e.stopPropagation();
                    startRepeat(delta);
                });
                button.addEventListener('mouseenter', () => { button.style.backgroundColor = '#e9ecef'; });
                button.addEventListener('mouseleave', () => { button.style.backgroundColor = 'transparent'; });
            };
            const upButton = createSpinnerButton('▲', false);
            const downButton = createSpinnerButton('▼', true);
            bindSpinner(upButton, 1);
            bindSpinner(downButton, -1);
            const spinnerColumn = createDiv({
                children: [upButton, downButton],
                style: { display: 'flex', flexDirection: 'column', borderLeft: '1px solid #dee2e6' }
            });
            const group = createDiv({
                children: [input, spinnerColumn],
                style: {
                    display: 'flex', alignItems: 'stretch', backgroundColor: '#fff',
                    border: '1px solid #ced4da', borderRadius: '0.25rem', overflow: 'hidden',
                    width: '5rem',
                    transition: 'border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out',
                }
            });
            input.addEventListener('focus', () => {
                group.style.borderColor = '#86b7fe';
                group.style.boxShadow = '0 0 0 0.2rem rgba(13, 110, 253, 0.25)';
            });
            input.addEventListener('blur', () => {
                group.style.borderColor = '#ced4da';
                group.style.boxShadow = 'none';
            });
            return group;
        };

        // —— 数字面板 ——
        const categories: { key: string; label: string }[] = [
            { key: 'general', label: '常规' },
            { key: 'number', label: '数值' },
            { key: 'currency', label: '货币' },
            { key: 'accounting', label: '会计专用' },
            { key: 'percentage', label: '百分比' },
            { key: 'scientific', label: '科学记数' },
            { key: 'text', label: '文本' },
            { key: 'date', label: '日期' },
            { key: 'time', label: '时间' },
            { key: 'fraction', label: '分数' },
        ];
        const categoryList = createDiv({
            style: {
                width: '10rem', flexShrink: '0',
                borderRight: '1px solid #dee2e6',
                overflowY: 'auto', maxHeight: '16rem',
            }
        });
        const categoryButtons: HTMLButtonElement[] = [];
        for (const cat of categories) {
            const btn = createButton({
                textContent: cat.label,
                style: {
                    display: 'block', width: '100%',
                    padding: '0.25rem 0.5rem', fontSize: '0.8rem',
                    textAlign: 'left', cursor: 'pointer',
                    border: 'none', borderRadius: '0',
                    backgroundColor: 'transparent', color: '#212529',
                }
            });
            btn.addEventListener('click', () => {
                numberCategory = cat.key;
                categoryButtons.forEach(b => {
                    b.style.backgroundColor = 'transparent';
                    b.style.color = '#212529';
                });
                btn.style.backgroundColor = '#0d6efd';
                btn.style.color = '#fff';
                rebuildControls();
                updatePreview();
            });
            categoryButtons.push(btn);
            categoryList.appendChild(btn);
        }
        // 初始分类高亮
        {
            const idx = categories.findIndex(c => c.key === numberCategory);
            if (idx >= 0) {
                categoryButtons[idx].style.backgroundColor = '#0d6efd';
                categoryButtons[idx].style.color = '#fff';
            }
        }

        // 控件区（随分类变化重建）
        const controlsContainer = createDiv({
            style: { flex: '1', padding: '0 0.75rem', minWidth: '0' }
        });

        /** 重建右侧控件区，根据当前分类显示对应的设置项 */
        const rebuildControls = (): void => {
            controlsContainer.innerHTML = '';
            const cat = numberCategory;
            if (cat === 'general') {
                controlsContainer.appendChild(createDiv({
                    textContent: '常规格式：不应用特定的数字格式。',
                    style: { fontSize: '0.8rem', color: '#6c757d', padding: '0.5rem 0' }
                }));
                return;
            }
            if (cat === 'text') {
                controlsContainer.appendChild(createDiv({
                    textContent: '文本格式：单元格内容作为文本处理，输入内容原样显示。',
                    style: { fontSize: '0.8rem', color: '#6c757d', padding: '0.5rem 0' }
                }));
                return;
            }
            if (cat === 'date' || cat === 'time') {
                // 日期/时间分类：渲染预设格式码列表（单选，每项展示格式码 + 预览）
                const presets = cat === 'date' ? datePresets : timePresets;
                const currentCode = cat === 'date' ? dateFormatCode : timeFormatCode;
                const presetButtons: HTMLButtonElement[] = [];
                for (const preset of presets) {
                    const isSel = preset.code === currentCode;
                    const btn = createButton({
                        textContent: `${preset.label}`,
                        attributes: { title: preset.code },
                        style: {
                            display: 'block', width: '100%',
                            padding: '0.25rem 0.5rem', fontSize: '0.8rem',
                            textAlign: 'left', cursor: 'pointer',
                            border: '1px solid ' + (isSel ? '#0d6efd' : '#ced4da'),
                            borderRadius: '0.25rem', marginBottom: '0.25rem',
                            backgroundColor: isSel ? '#0d6efd' : 'transparent',
                            color: isSel ? '#fff' : '#212529',
                        }
                    });
                    btn.addEventListener('click', () => {
                        if (cat === 'date') dateFormatCode = preset.code;
                        else timeFormatCode = preset.code;
                        presetButtons.forEach(b => {
                            b.style.backgroundColor = 'transparent';
                            b.style.color = '#212529';
                            b.style.borderColor = '#ced4da';
                        });
                        btn.style.backgroundColor = '#0d6efd';
                        btn.style.color = '#fff';
                        btn.style.borderColor = '#0d6efd';
                        updatePreview();
                    });
                    presetButtons.push(btn);
                    controlsContainer.appendChild(btn);
                }
                return;
            }
            if (cat === 'fraction') {
                controlsContainer.appendChild(createDiv({
                    textContent: '此分类暂不支持自定义格式，确定后将使用常规格式。',
                    style: { fontSize: '0.8rem', color: '#6c757d', padding: '0.5rem 0' }
                }));
                return;
            }

            // 小数位数（步进输入组，参照行高对话框的 Bootstrap input-group 风格）
            const decLabel = createDiv({ textContent: '小数位数：', style: { fontSize: '0.8rem', flexShrink: '0', paddingTop: '0.125rem' } });
            const decStepper = buildStepperInput({
                value: decimalPlaces, min: 0, max: 30,
                onChange: (v: number) => { decimalPlaces = v; updatePreview(); }
            });
            controlsContainer.appendChild(createDiv({
                children: [decLabel, decStepper],
                style: { display: 'flex', alignItems: 'center', gap: '0.5rem', marginBottom: '0.5rem' }
            }));

            // 使用千位分隔符（仅数值）
            if (cat === 'number') {
                const thousandsCheck = createInput({ type: 'checkbox' });
                thousandsCheck.checked = useThousands;
                thousandsCheck.addEventListener('change', () => {
                    useThousands = thousandsCheck.checked;
                    updatePreview();
                });
                controlsContainer.appendChild(createDiv({
                    children: [thousandsCheck, createDiv({
                        textContent: '使用千位分隔符 (,)', style: { fontSize: '0.8rem' }
                    })],
                    style: { display: 'flex', alignItems: 'center', gap: '0.25rem', marginBottom: '0.5rem' }
                }));
            }

            // 负数样式（数值/百分比/科学记数）
            if (cat === 'number' || cat === 'percentage' || cat === 'scientific') {
                controlsContainer.appendChild(createDiv({
                    textContent: '负数样式：', style: { fontSize: '0.8rem', marginBottom: '0.25rem' }
                }));
                const negOptions = [
                    { label: '-1234.10', value: 0 },
                    { label: '(1234.10)', value: 1 },
                    { label: '-1234.10 红', value: 2 },
                    { label: '(1234.10) 红', value: 3 },
                ];
                const negButtons: HTMLButtonElement[] = [];
                const negContainer = createDiv({
                    style: { display: 'flex', flexWrap: 'wrap', gap: '0.25rem', marginBottom: '0.5rem' }
                });
                for (const opt of negOptions) {
                    const btn = createButton({
                        textContent: opt.label,
                        style: {
                            padding: '0.125rem 0.5rem', fontSize: '0.75rem',
                            border: '1px solid #ced4da', borderRadius: '0.25rem',
                            cursor: 'pointer', backgroundColor: 'transparent', color: '#212529',
                        }
                    });
                    btn.addEventListener('click', () => {
                        negativeStyle = opt.value;
                        negButtons.forEach(b => {
                            b.style.backgroundColor = 'transparent';
                            b.style.color = '#212529';
                            b.style.borderColor = '#ced4da';
                        });
                        btn.style.backgroundColor = '#0d6efd';
                        btn.style.color = '#fff';
                        btn.style.borderColor = '#0d6efd';
                        updatePreview();
                    });
                    if (opt.value === negativeStyle) {
                        btn.style.backgroundColor = '#0d6efd';
                        btn.style.color = '#fff';
                        btn.style.borderColor = '#0d6efd';
                    }
                    negButtons.push(btn);
                    negContainer.appendChild(btn);
                }
                controlsContainer.appendChild(negContainer);
            }

            // 货币符号（货币/会计专用）
            if (cat === 'currency' || cat === 'accounting') {
                controlsContainer.appendChild(createDiv({
                    textContent: '货币符号：', style: { fontSize: '0.8rem', marginBottom: '0.25rem' }
                }));
                const symbols = ['¥', '$', '€', '£'];
                const curButtons: HTMLButtonElement[] = [];
                const curContainer = createDiv({
                    style: { display: 'flex', gap: '0.25rem', marginBottom: '0.5rem' }
                });
                for (const sym of symbols) {
                    const btn = createButton({
                        textContent: sym,
                        style: {
                            padding: '0.125rem 0.5rem', fontSize: '0.9rem',
                            border: '1px solid #ced4da', borderRadius: '0.25rem',
                            cursor: 'pointer', backgroundColor: 'transparent', color: '#212529',
                        }
                    });
                    btn.addEventListener('click', () => {
                        currencySymbol = sym;
                        curButtons.forEach(b => {
                            b.style.backgroundColor = 'transparent';
                            b.style.color = '#212529';
                            b.style.borderColor = '#ced4da';
                        });
                        btn.style.backgroundColor = '#0d6efd';
                        btn.style.color = '#fff';
                        btn.style.borderColor = '#0d6efd';
                        updatePreview();
                    });
                    if (sym === currencySymbol) {
                        btn.style.backgroundColor = '#0d6efd';
                        btn.style.color = '#fff';
                        btn.style.borderColor = '#0d6efd';
                    }
                    curButtons.push(btn);
                    curContainer.appendChild(btn);
                }
                controlsContainer.appendChild(curContainer);
            }
        };
        rebuildControls();

        // 预览区
        const previewBox = createDiv({
            children: [
                createDiv({ textContent: '示例', style: { fontSize: '0.8rem', color: '#6c757d', marginBottom: '0.25rem' } }),
                previewPos,
                previewNeg,
            ],
            style: {
                marginTop: '0.5rem', padding: '0.5rem',
                border: '1px solid #dee2e6', borderRadius: '0.25rem',
                backgroundColor: '#f8f9fa',
            }
        });
        updatePreview();

        const numberPanel = createDiv({
            children: [
                createDiv({
                    children: [categoryList, controlsContainer],
                    style: { display: 'flex', minHeight: '10rem' }
                }),
                previewBox,
            ],
            style: { padding: '0.5rem 0' }
        });

        // —— 边框面板 ——
        // 辅助：创建一组互斥的可选按钮
        const createButtonRow = (
            items: { label: string; value: string }[],
            getValue: () => string,
            setValue: (v: string) => void
        ): HTMLElement => {
            const buttons: HTMLButtonElement[] = [];
            const container = createDiv({
                style: { display: 'flex', flexWrap: 'wrap', gap: '0.25rem', marginBottom: '0.5rem' }
            });
            for (const item of items) {
                const btn = createButton({
                    textContent: item.label,
                    style: {
                        padding: '0.125rem 0.5rem', fontSize: '0.8rem',
                        border: '1px solid #ced4da', borderRadius: '0.25rem',
                        cursor: 'pointer', backgroundColor: 'transparent', color: '#212529',
                    }
                });
                btn.addEventListener('click', () => {
                    setValue(item.value);
                    buttons.forEach(b => {
                        b.style.backgroundColor = 'transparent';
                        b.style.color = '#212529';
                        b.style.borderColor = '#ced4da';
                    });
                    btn.style.backgroundColor = '#0d6efd';
                    btn.style.color = '#fff';
                    btn.style.borderColor = '#0d6efd';
                });
                if (item.value === getValue()) {
                    btn.style.backgroundColor = '#0d6efd';
                    btn.style.color = '#fff';
                    btn.style.borderColor = '#0d6efd';
                }
                buttons.push(btn);
                container.appendChild(btn);
            }
            return container;
        };

        // 预设按钮
        const presetRow = createDiv({
            children: [
                createDiv({ textContent: '预设：', style: { fontSize: '0.8rem', width: '4.5rem', flexShrink: '0', paddingTop: '0.125rem' } }),
                createButtonRow(
                    [
                        { label: '无', value: 'none' },
                        { label: '外边框', value: 'outer' },
                        { label: '内边框', value: 'all' },
                    ],
                    () => borderAction,
                    (v) => { borderAction = v; }
                ),
            ],
            style: { display: 'flex', alignItems: 'flex-start' }
        });

        // 线条样式
        const lineStyleRow = createDiv({
            children: [
                createDiv({ textContent: '线条样式：', style: { fontSize: '0.8rem', width: '4.5rem', flexShrink: '0', paddingTop: '0.125rem' } }),
                createDiv({
                    children: [
                        createButton({
                            textContent: '│ 细实线',
                            style: {
                                padding: '0.125rem 0.5rem', fontSize: '0.8rem',
                                border: '1px solid #ced4da', borderRadius: '0.25rem',
                                cursor: 'pointer',
                                backgroundColor: lineStyle === 1 ? '#0d6efd' : 'transparent',
                                color: lineStyle === 1 ? '#fff' : '#212529',
                                borderColor: lineStyle === 1 ? '#0d6efd' : '#ced4da',
                            }
                        }),
                        createButton({
                            textContent: '┃ 粗实线',
                            style: {
                                padding: '0.125rem 0.5rem', fontSize: '0.8rem',
                                border: '1px solid #ced4da', borderRadius: '0.25rem',
                                cursor: 'pointer',
                                backgroundColor: lineStyle === 2 ? '#0d6efd' : 'transparent',
                                color: lineStyle === 2 ? '#fff' : '#212529',
                                borderColor: lineStyle === 2 ? '#0d6efd' : '#ced4da',
                            }
                        }),
                    ],
                    style: { display: 'flex', gap: '0.25rem' }
                }),
            ],
            style: { display: 'flex', alignItems: 'flex-start', marginBottom: '0.5rem' }
        });
        // 线条样式按钮事件
        {
            const btns = lineStyleRow.querySelectorAll('button');
            btns[0].addEventListener('click', () => {
                lineStyle = 1;
                btns.forEach((b, i) => {
                    (b as HTMLButtonElement).style.backgroundColor = (i === 0) ? '#0d6efd' : 'transparent';
                    (b as HTMLButtonElement).style.color = (i === 0) ? '#fff' : '#212529';
                    (b as HTMLButtonElement).style.borderColor = (i === 0) ? '#0d6efd' : '#ced4da';
                });
            });
            btns[1].addEventListener('click', () => {
                lineStyle = 2;
                btns.forEach((b, i) => {
                    (b as HTMLButtonElement).style.backgroundColor = (i === 1) ? '#0d6efd' : 'transparent';
                    (b as HTMLButtonElement).style.color = (i === 1) ? '#fff' : '#212529';
                    (b as HTMLButtonElement).style.borderColor = (i === 1) ? '#0d6efd' : '#ced4da';
                });
            });
        }

        // 颜色选择
        const colorInput = createInput({
            type: 'color',
            style: { width: '2rem', height: '1.5rem', padding: '0', border: '1px solid #ced4da', borderRadius: '0.25rem', cursor: 'pointer' }
        });
        colorInput.value = borderColor;
        colorInput.addEventListener('input', () => { borderColor = colorInput.value; });
        const swatchColors = ['#000000', '#ff0000', '#008000', '#0000ff', '#ffc000', '#ff00ff', '#00ffff', '#ffffff'];
        const swatchContainer = createDiv({ style: { display: 'flex', gap: '0.25rem', flexWrap: 'wrap' } });
        for (const sc of swatchColors) {
            const sw = createButton({
                style: {
                    width: '1.2rem', height: '1.2rem', padding: '0',
                    border: '1px solid #ced4da', borderRadius: '0.2rem',
                    cursor: 'pointer', backgroundColor: sc,
                }
            });
            sw.addEventListener('click', () => {
                borderColor = sc;
                colorInput.value = sc;
            });
            swatchContainer.appendChild(sw);
        }
        const colorRow = createDiv({
            children: [
                createDiv({ textContent: '颜色：', style: { fontSize: '0.8rem', width: '4.5rem', flexShrink: '0', paddingTop: '0.125rem' } }),
                colorInput,
                swatchContainer,
            ],
            style: { display: 'flex', alignItems: 'center', gap: '0.5rem', marginBottom: '0.5rem' }
        });

        // 位置按钮
        const positionRow = createDiv({
            children: [
                createDiv({ textContent: '位置：', style: { fontSize: '0.8rem', width: '4.5rem', flexShrink: '0', paddingTop: '0.125rem' } }),
                createButtonRow(
                    [
                        { label: '上', value: 'top' },
                        { label: '下', value: 'bottom' },
                        { label: '左', value: 'left' },
                        { label: '右', value: 'right' },
                        { label: '外边框', value: 'outer' },
                        { label: '内边框', value: 'all' },
                        { label: '无', value: 'none' },
                    ],
                    () => borderAction,
                    (v) => { borderAction = v; }
                ),
            ],
            style: { display: 'flex', alignItems: 'flex-start' }
        });

        const borderPanel = createDiv({
            children: [presetRow, lineStyleRow, colorRow, positionRow],
            style: { padding: '0.5rem 0' }
        });

        // —— Tab 栏 ——
        const tabNumber = createButton({
            textContent: '数字',
            style: {
                padding: '0.375rem 1rem', fontSize: '0.875rem',
                border: '1px solid #dee2e6', borderBottom: 'none',
                borderRadius: '0.25rem 0.25rem 0 0',
                cursor: 'pointer', position: 'relative',
                top: '1px', marginBottom: '-1px',
            }
        });
        const tabBorder = createButton({
            textContent: '边框',
            style: {
                padding: '0.375rem 1rem', fontSize: '0.875rem',
                border: '1px solid #dee2e6', borderBottom: 'none',
                borderRadius: '0.25rem 0.25rem 0 0',
                cursor: 'pointer', position: 'relative',
                top: '1px', marginBottom: '-1px',
            }
        });
        const tabBar = createDiv({
            children: [tabNumber, tabBorder],
            style: { display: 'flex', borderBottom: '1px solid #dee2e6' }
        });

        // Tab 切换
        const showTab = (tab: 'number' | 'border'): void => {
            const isActive = (t: 'number' | 'border') => t === tab;
            const updateTabStyle = (btn: HTMLButtonElement, active: boolean): void => {
                btn.style.backgroundColor = active ? '#fff' : '#f8f9fa';
                btn.style.color = active ? '#0d6efd' : '#212529';
                btn.style.borderBottom = active ? '1px solid #fff' : '1px solid #dee2e6';
            };
            updateTabStyle(tabNumber, isActive('number'));
            updateTabStyle(tabBorder, isActive('border'));
            numberPanel.style.display = tab === 'number' ? '' : 'none';
            borderPanel.style.display = tab === 'border' ? '' : 'none';
        };
        tabNumber.addEventListener('click', () => showTab('number'));
        tabBorder.addEventListener('click', () => showTab('border'));
        showTab('number');

        // —— 根容器 ——
        const root = createDiv({
            children: [tabBar, numberPanel, borderPanel],
            style: { display: 'flex', flexDirection: 'column' }
        });

        // —— 确定时应用变更 ——
        const applyChanges = (): void => {
            data.runWithFullStateUndo('设置单元格格式', () => {
                // 数字格式（分数分类仍 fallback 到常规，日期/时间存储实际格式码）
                let formatCode = buildFormatCode();
                if (numberCategory === 'fraction') {
                    formatCode = 'General';
                }
                data.setCellsNumberFormat(startCol, endCol, startRow, endRow, formatCode);

                // 边框：按选中的预设/位置调 setSelectedCellsBorder 设宽度，setCellsBorderColor 设颜色
                if (borderAction) {
                    let style: CellBorderStyle | undefined;
                    switch (borderAction) {
                        case 'none': style = CellBorderStyle.BorderNone; break;
                        case 'outer': style = lineStyle === 2 ? CellBorderStyle.BorderWideOuter : CellBorderStyle.BorderOuter; break;
                        case 'all': style = CellBorderStyle.BorderAll; break;
                        case 'top': style = CellBorderStyle.BorderTop; break;
                        case 'bottom': style = CellBorderStyle.BorderBottom; break;
                        case 'left': style = CellBorderStyle.BorderLeft; break;
                        case 'right': style = CellBorderStyle.BorderRight; break;
                    }
                    if (style !== undefined) {
                        data.setSelectedCellsBorder(style);
                    }
                    if (borderAction !== 'none') {
                        data.setCellsBorderColor(startCol, endCol, startRow, endRow, borderColor);
                    }
                }
            });
            this.refreshAfterContextMenu();
        };

        // —— 对话框 ——
        const dialog = new Dialog({
            title: '设置单元格格式',
            content: root,
            size: 'lg',
            buttons: [
                { text: '取消', variant: 'secondary' },
                {
                    text: '确定',
                    variant: 'primary',
                    close: false,
                    onClick: (d) => {
                        applyChanges();
                        d.hide();
                    }
                },
            ],
            onClose: () => dialog.destroy()
        });
        dialog.show();
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