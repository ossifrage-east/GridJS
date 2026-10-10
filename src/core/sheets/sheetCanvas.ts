/*  # 表格的 canvas 组件 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { CELL_PADDING, DEFAULT_ANCHOR_SELECTED_BORDER_COLOR, DEFAULT_CELL_BG_COLOR, DEFAULT_CELL_HEIGHT, DEFAULT_CELL_WIDTH, DEFAULT_FONT_FAMILY,
    DEFAULT_FONT_SIZE, DEFAULT_SELECTED_BG_COLOR, DEFAULT_SELECTED_BORDER_COLOR, DEFAULT_SELECTED_BORDER_WIDTH, GRID_LINE_COLOR,
    HEADER_BG_COLOR, TEXT_COLOR, DataEvents,
    SHEETS_NAME,
    MouseLocation} from "../constant";
import { Canvas, CanvasOptions } from "./canvas";
import { Filter } from "../assistant/filter";

export class SheetCanvas extends Canvas {
    private filters: {cellName: string, filter: Filter}[] = [];
    constructor(options: CanvasOptions) {
        super(options);
        this.parentElement = options.parentElement;
        this.data = options.data;
        this.menu = options.menu;
        this.setupContextMenu();
        this.init();
    }

    /**
     * 右键菜单动作执行后的刷新钩子（重写）
     *
     * 基类默认空操作；主数据画布需在复制/剪切/粘贴/清除内容/删除等动作后立即重绘，
     * 否则用户看到的是旧画面（数据已改但画布未刷新）。
     * 与 ColHeaderCanvas / RowHeaderCanvas 的重写保持一致。
     */
    protected refreshAfterContextMenu(): void {
        this.draw();
    }    

    /**
     * 绘制加载动画（旋转的菊花瓣图形）
     * @param {number} rotation - 旋转角度（弧度）
     */
    public drawLoadingAnimation(rotation: number): void {
        const width = this.canvas.width / window.devicePixelRatio;
        const height = this.canvas.height / window.devicePixelRatio;
        const centerX = width / 2;
        const centerY = height / 2;
        const radius = 20;
        const petalCount = 12;
        const petalLength = 12;

        this.ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);
        this.ctx.save();
        this.ctx.translate(centerX * window.devicePixelRatio, centerY * window.devicePixelRatio);
        this.ctx.rotate(rotation);

        for (let i = 0; i < petalCount; i++) {
            const angle = (Math.PI * 2 / petalCount) * i;
            const opacity = 1 - i / petalCount;

            this.ctx.save();
            this.ctx.rotate(angle);
            this.ctx.beginPath();
            this.ctx.moveTo(0, -radius * window.devicePixelRatio);
            this.ctx.lineTo(0, -(radius + petalLength) * window.devicePixelRatio);
            this.ctx.strokeStyle = `rgba(51, 51, 51, ${opacity})`;
            this.setLineWidth(3);
            this.ctx.lineCap = 'round';
            this.ctx.stroke();
            this.ctx.restore();
        }

        this.ctx.restore();
    }

    

    /**
     * 初始化 canvas
     */
    private init() {
        this.canvas.setAttribute('tabindex', '0');
        this.canvas.id = SHEETS_NAME.SHEET;
        this.updateCanvasSize();
        this.parentElement.appendChild(this.canvas);
    }

    /**
     * 更新 canvas 大小
     */
    public updateCanvasSize() {
        this.canvas.width = Math.round(this.data.visibleView.sheetWidth) * window.devicePixelRatio;
        this.canvas.height = Math.round(this.data.visibleView.sheetHeight) * window.devicePixelRatio;
        this.canvas.style.width = `${this.data.visibleView.sheetWidth}px`;
        this.canvas.style.height = `${this.data.visibleView.sheetHeight}px`;
        this.canvas.style.left = `${this.data.visibleView.rowHeaderWidth}px`;
        this.canvas.style.top = `${this.data.visibleView.colHeaderHeight}px`;
    }

    /**
     * 绘制表格内容
     */
    public async draw() {
        // await this.startLoadingAnimation();
        this.setupContext();
        this.ctx.save();
        this.clear();
        this.drawGridLines();
        this.drawCellBgColor();
        this.drawAllMergedCellBackground();
        this.drawCellsText();
        this.drawCellBorder();
        if (this.data.startLocation === MouseLocation.IN_ANCHOR) {
            this.drawAnchorSelcetedRect();
        }
        this.drawSelectedRect();
        this.ctx.restore();
        this.setFilters();
    }

    /**
     * 设置筛选组件实例, 如果存在筛选组件实例，先移除旧的实例
     */
    private setFilters() {
        const dataFilters = this.data.values.filter(v => v.filter !== undefined);
        if (dataFilters.length === 0) {  // 如果data中没有筛选数据，则销毁组件实例直接返回
            this.filters.forEach(filter => filter.filter.destroy());
            this.filters = [];
            return;
        }
        // 反向清理：cell 已被删除（如删除行/列）的 filter 实例需销毁，
        // 否则按钮 DOM 拖留、停留在旧位置显示
        this.filters = this.filters.filter(f => {
            if (!dataFilters.find(c => c.cell === f.cellName)) {
                f.filter.destroy();
                return false;
            }
            return true;
        });

        for (const cell of dataFilters) {
            const filter = this.filters.find(f => f.cellName === cell.cell);
            if (filter) { // 如果存在筛选组件实例，直接设置位置
                filter.filter.setPosition(cell);
                continue;
            }
            const _filter = new Filter({parentElement: this.parentElement, data: this.data, menu: this.menu});
            _filter.setPosition(cell);
            this.filters.push({cellName: cell.cell, filter: _filter});
        }
    }
    
    /**
     * 绘制网格线
     */
    private drawGridLines() {
        if (this.data.showGridLines === false) return; // 如果不显示网格线，直接返回
        this.ctx.save();
        this.ctx.strokeStyle = GRID_LINE_COLOR;

        // 绘制网格的垂直线条（跳过隐藏列）
        const { startCol, endCol } = this.data.getVisibleCols();
        for (let c = startCol - 1; c <= endCol; c++) {
            if (this.data.colHeaders.getAt(c).isHidden) continue;
            this.setLineWidth(1);  // 使用 CSS 像素，内部自动转换为物理像素
            const x = (this.data.colHeaders.getAt(c).left - this.data.colHeaders.offsetWidth) * this.data.zoom;
            this.drawLine(x, 0, x, this.data.visibleView.sheetHeight);
        }
        
        // 绘制网格的水平线条（跳过隐藏行）
        const { startRow, endRow } = this.data.getVisibleRows();
        for (let r = startRow - 1; r <= endRow; r++) {
            if (this.data.rowHeaders.getAt(r).isHidden) continue;
            this.setLineWidth(1);  // 使用 CSS 像素，内部自动转换为物理像素
            const y = (this.data.rowHeaders.getAt(r).top - this.data.rowHeaders.offsetHeight) * this.data.zoom;
            this.drawLine(0, y, this.data.visibleView.sheetWidth, y);
        }

        // 筛选/隐藏行后（或行数不足一屏时），内容总高可能小于可视高度，屏幕底部会露出无网格线的空白。
        // 从内容底部边缘开始按默认行高补齐虚拟横向网格线，直到铺满可视高度（只画线，不代表实际数据行）。
        const contentBottom = (this.data.rowHeaders.allRowHeight - this.data.rowHeaders.offsetHeight) * this.data.zoom;
        if (contentBottom < this.data.visibleView.sheetHeight) {
            this.setLineWidth(1);
            const virtualStep = DEFAULT_CELL_HEIGHT * this.data.zoom;
            for (let y = contentBottom; y <= this.data.visibleView.sheetHeight; y += virtualStep) {
                this.drawLine(0, y, this.data.visibleView.sheetWidth, y);
            }
        }

        // 隐藏列后（或列数不足一屏时），内容总宽可能小于可视宽度，屏幕右侧会露出无网格线的空白。
        // 从内容右缘开始按默认列宽补齐虚拟竖向网格线，直到铺满可视宽度（只画线，不代表实际数据列）。
        const contentRight = (this.data.colHeaders.allColWidth - this.data.colHeaders.offsetWidth) * this.data.zoom;
        if (contentRight < this.data.visibleView.sheetWidth) {
            this.setLineWidth(1);
            const virtualStepX = DEFAULT_CELL_WIDTH * this.data.zoom;
            for (let x = contentRight; x <= this.data.visibleView.sheetWidth; x += virtualStepX) {
                this.drawLine(x, 0, x, this.data.visibleView.sheetHeight);
            }
        }
        this.ctx.restore();
    }

    /**
     * 计算单元格考虑隐藏行列后的可见绘制矩形（视图坐标，已减去滚动偏移）
     *
     * 规则：
     * - 单元格所在列/行全部隐藏 → 返回 null（整个单元格不绘制）；
     * - 合并单元格跨越隐藏列/行 → 只保留可见部分：left/top 取第一个可见列/行的位置，
     *   width/height 只累加可见列/行的尺寸，被隐藏的部分被裁掉；
     * - 普通单元格（colspan/rowspan 为 1）所在列/行隐藏 → 同样返回 null。
     *
     * 合并展开逻辑与 DataCollection.calculateCellRect 保持一致：
     * 命中合并区域时取左上角起点，colspan/rowspan 从 values 中该单元格读取。
     * @param {number} col - 单元格列号（1-based）
     * @param {number} row - 单元格行号（1-based）
     * @returns {{ left: number, top: number, width: number, height: number } | null} 可见矩形；完全隐藏时返回 null
     * @private
     */
    private getVisibleCellRect(col: number, row: number): { left: number, top: number, width: number, height: number } | null {
        // 合并展开：命中合并区域时取左上角起点
        let startCol = col;
        let startRow = row;
        for (const mergedCell of this.data.getMergedCells()) {
            const [startCell, endCell] = mergedCell.split(':');
            const start = this.data.getCellColAndRow(startCell);
            const end = this.data.getCellColAndRow(endCell);
            if (col >= start.col && col <= end.col && row >= start.row && row <= end.row) {
                startCol = start.col;
                startRow = start.row;
                break;
            }
        }

        // 读取合并跨度
        const cellName = this.data.getCellName(startCol, startRow);
        const cell = this.data.values.find(v => v.cell === cellName);
        const colspan = cell?.colspan || 1;
        const rowspan = cell?.rowspan || 1;

        // 水平方向：只累加可见列，left 取第一个可见列
        let left = 0;
        let width = 0;
        let colFound = false;
        for (let c = 0; c < colspan; c++) {
            const header = this.data.colHeaders.getAt(startCol + c - 1);
            if (!header) break;
            if (header.isHidden) continue;
            if (!colFound) {
                left = header.left;
                colFound = true;
            }
            width += header.width;
        }
        if (!colFound) return null;

        // 垂直方向：只累加可见行，top 取第一个可见行
        let top = 0;
        let height = 0;
        let rowFound = false;
        for (let r = 0; r < rowspan; r++) {
            const header = this.data.rowHeaders.getAt(startRow + r - 1);
            if (!header) break;
            if (header.isHidden) continue;
            if (!rowFound) {
                top = header.top;
                rowFound = true;
            }
            height += header.height;
        }
        if (!rowFound) return null;

        return {
            left: left - this.data.colHeaders.offsetWidth,
            top: top - this.data.rowHeaders.offsetHeight,
            width,
            height,
        };
    }

    /**
     * 绘制单元格背景颜色
     */
    private drawCellBgColor() {
        for (let i = 0; i < this.data.values.length; i++) {
            const cell = this.data.values[i];
            const { col, row } = this.data.getCellColAndRow(cell.cell);
            // 隐藏行列不绘制；合并单元格只绘制可见部分
            const rect = this.getVisibleCellRect(col, row);
            if (!rect) continue;
            const cellLeft = (rect.left + 1) * this.data.zoom;
            const cellTop = (rect.top + 1) * this.data.zoom;
            const cellWidth = (rect.width - 1) * this.data.zoom;
            const cellHeight = (rect.height - 1) * this.data.zoom;
            if (!cell.backgroundColor) continue;
            this.ctx.save();
            this.ctx.fillStyle = cell.backgroundColor;
            this.fillRect(cellLeft, cellTop, cellWidth, cellHeight);
            this.ctx.restore();        
        }
    }
    
    /**
     * 绘制单元格边框
     */
    private drawCellBorder() {
        for (let i = 0; i < this.data.values.length; i++) {
            const cell = this.data.values[i];
            const { col, row } = this.data.getCellColAndRow(cell.cell);
            // 隐藏行列不绘制；合并单元格边框只沿可见部分绘制
            const rect = this.getVisibleCellRect(col, row);
            if (!rect) continue;
            const cellLeft = rect.left * this.data.zoom;
            const cellTop = rect.top * this.data.zoom;
            const cellWidth = rect.width * this.data.zoom;
            const cellHeight = rect.height * this.data.zoom;

            this.ctx.save();
            this.ctx.strokeStyle = cell.borderColor || 'black';
            if (cell.borderTopWidth) {
                this.drawLine(cellLeft, cellTop, cellLeft + cellWidth, cellTop, cell.borderTopWidth);
            }
            if (cell.borderLeftWidth) {
                this.drawLine(cellLeft, cellTop, cellLeft, cellTop + cellHeight, cell.borderLeftWidth);
            }
            if (cell.borderRightWidth) {
                this.drawLine(cellLeft + cellWidth, cellTop, cellLeft + cellWidth, cellTop + cellHeight, cell.borderRightWidth);
            }
            if (cell.borderBottomWidth) {
                this.drawLine(cellLeft, cellTop + cellHeight, cellLeft + cellWidth, cellTop + cellHeight, cell.borderBottomWidth);
            }
            this.ctx.restore();        
        }
    }

    /**
     * 绘制选中区域边框及activedCell透明背景
     */
    private drawSelectedRect() {
        const { left, top, width, height } = this.data.getSelectedRect();
        // 选区完全落在隐藏行列上（或无有效选区）时宽高为 0：
        // 跳过绘制，否则零尺寸 OffscreenCanvas 会导致 drawImage 抛出异常
        if (width <= 0 || height <= 0) return;
        const { activedLeft, activedTop, activedWidth, activedHeight } = this.data.getActivedRect();
        // 绘制选中区域背景及activedCell透明背景
        const { offscreenCanvas, offscreenCtx } = this.offscreen(width * this.data.zoom * window.devicePixelRatio,
            height * this.data.zoom * window.devicePixelRatio);
        offscreenCtx.clearRect(0, 0, 
            width * this.data.zoom * window.devicePixelRatio,
            height * this.data.zoom * window.devicePixelRatio);
        // 绘制选中区域背景（从 offscreenCanvas 的原点开始绘制）
        offscreenCtx.fillStyle = DEFAULT_SELECTED_BG_COLOR;
        offscreenCtx.fillRect(0, 0,
            offscreenCanvas.width,
            offscreenCanvas.height);
        
        // 绘制activedCell透明背景（使用相对坐标）
        offscreenCtx.clearRect(
            (activedLeft - left) * this.data.zoom * window.devicePixelRatio,
            (activedTop - top) * this.data.zoom * window.devicePixelRatio,
            activedWidth * this.data.zoom * window.devicePixelRatio,
            activedHeight * this.data.zoom * window.devicePixelRatio);

        // 绘制选中区域背景
        this.ctx.save();
        this.ctx.drawImage(offscreenCanvas, left * this.data.zoom * window.devicePixelRatio,
            top * this.data.zoom * window.devicePixelRatio);

        this.ctx.strokeStyle = DEFAULT_SELECTED_BORDER_COLOR;
        this.setLineWidth(DEFAULT_SELECTED_BORDER_WIDTH);
        this.strokeRect(left * this.data.zoom ,
            top * this.data.zoom,
            width * this.data.zoom,
            height * this.data.zoom);
        this.ctx.restore();
        this.drawAnchor();
    }

    /**
     * 绘制选中框右下角的Anchor
     */
    public drawAnchor() {
        const { left: AnchorLeft, top: AnchorTop,
            width: AnchorWidth, height: AnchorHeight } = this.data.getAnchorRect();
        this.ctx.save();
        // 绘制选中框右下角的Anchor
        this.ctx.strokeStyle = 'white';
        this.setLineWidth(DEFAULT_SELECTED_BORDER_WIDTH);
        this.strokeRect(AnchorLeft * this.data.zoom,
            AnchorTop * this.data.zoom,
            AnchorWidth * this.data.zoom,
            AnchorHeight * this.data.zoom);
        this.ctx.fillStyle = DEFAULT_SELECTED_BORDER_COLOR;
        this.fillRect(AnchorLeft * this.data.zoom,
            AnchorTop * this.data.zoom,
            AnchorWidth * this.data.zoom,
            AnchorHeight * this.data.zoom);
        this.ctx.restore();
    }

    /**
     * 绘制锚点选择框边框及透明背景
     */
    public drawAnchorSelcetedRect() {
        if (!this.data.anchorSelection) return;
        const selected = this.data.anchorSelection.split(':');
        const { col: startCol, row: startRow } = this.data.getCellColAndRow(selected[0]);
        const { left: startLeft, top: startTop, width: startWidth, height: startHeight } = this.data.calculateCellRect(startCol, startRow);
        const { col: endCol, row: endRow } = this.data.getCellColAndRow(selected[1]);
        const { left: endLeft, top: endTop, width: endWidth, height: endHeight } = this.data.calculateCellRect(endCol, endRow);

        let _left = Math.min(startLeft, endLeft);
        let _top = Math.min(startTop, endTop);
        let _width = Math.max(startLeft, endLeft) - _left + (Math.max(startLeft, endLeft) === endLeft ? endWidth : startWidth);
        let _height = Math.max(startTop, endTop) - _top + (Math.max(startTop, endTop) === endTop ? endHeight : startHeight);
        this.ctx.save();
        this.ctx.strokeStyle = DEFAULT_ANCHOR_SELECTED_BORDER_COLOR;
        this.setLineWidth(DEFAULT_SELECTED_BORDER_WIDTH);
        this.strokeRect(_left * this.data.zoom, _top * this.data.zoom,
            _width * this.data.zoom, _height * this.data.zoom);
        this.ctx.restore();
    }

    /**
     * 绘制单元格内容
     */
    private drawCellsText() {
        this.ctx.save();
        for (let i=0; i<this.data.values.length; i++) {
            const cell = this.data.values[i].cell;
            const { col, row } = this.data.getCellColAndRow(cell);

            // 计算考虑隐藏行列后的可见矩形；完全隐藏时跳过绘制
            const rect = this.getVisibleCellRect(col, row);
            if (!rect) continue;
            const cellWidth = rect.width;
            const cellHeight = rect.height;

            // 获取单元格位置
            const cellLeft = rect.left;
            const cellTop = rect.top;
            const { lines } = this.measureCellText(cell);
            const { offscreenCanvas, offscreenCtx } = this.offscreen(
                Math.round(cellWidth * this.data.zoom * window.devicePixelRatio),
                Math.round(cellHeight * this.data.zoom * window.devicePixelRatio)
            );
            offscreenCtx.clearRect(0, 0, offscreenCanvas.width, offscreenCanvas.height);
            
            // 复制主canvas的绘图上下文设置
            // const wrap = this.data.values[i].wrap || false || lines.length > 1;
            const letterSpacing = this.data.values[i].letterSpacing || 0;
            const lineSpacing = this.data.values[i].lineSpacing || 0;
            offscreenCtx.textAlign = 'left';
            offscreenCtx.textBaseline = 'bottom';
            // 在offscreenCanvas上绘制文本
            let k = 0; // 特殊字符索引
            for (let line of lines) {
                let currentX = line.x;
                let currentY = line.y + line.lineHeight - lineSpacing;
                for (let j=0; j < line.text.length; j++) {
                    const char = line.text[j] === '\n' ? '' : line.text[j];
                    const { 
                        charWidth,
                        fontColor,
                        fontWeight,
                        fontStyle,
                        fontSize,
                        fontFamily,
                        underline,
                        strikethrough 
                    } = this.calculateCharGeometry(cell, k);
                    
                    offscreenCtx.save();
                    offscreenCtx.fillStyle = `${fontColor}`;
                    const fontParts = [];
                    if (fontStyle) fontParts.push(fontStyle);
                    if (fontWeight) fontParts.push(fontWeight);
                    fontParts.push(`${Math.round(fontSize * window.devicePixelRatio)}px`);
                    fontParts.push(fontFamily);
                    offscreenCtx.font = fontParts.join(' ');
                    offscreenCtx.fillText(char, Math.round(currentX * window.devicePixelRatio), Math.round(currentY * window.devicePixelRatio));
                    
                    // 绘制下划线
                    if (underline) {
                        offscreenCtx.strokeStyle = fontColor;
                        offscreenCtx.lineWidth = Math.max(1, Math.round(fontSize / 10 * window.devicePixelRatio));
                        const underlineY = Math.round(currentY * window.devicePixelRatio - 1);
                        offscreenCtx.beginPath();
                        offscreenCtx.moveTo(Math.round(currentX * window.devicePixelRatio), underlineY);
                        offscreenCtx.lineTo(Math.round((currentX + charWidth) * window.devicePixelRatio), underlineY);
                        offscreenCtx.stroke();
                    }
                    
                    // 绘制删除线
                    if (strikethrough) {
                        offscreenCtx.strokeStyle = fontColor;
                        offscreenCtx.lineWidth = Math.max(1, Math.round(fontSize / 10 * window.devicePixelRatio));
                        const strikethroughY = Math.round((currentY - fontSize * 0.48) * window.devicePixelRatio);
                        offscreenCtx.beginPath();
                        offscreenCtx.moveTo(Math.round(currentX * window.devicePixelRatio), strikethroughY);
                        offscreenCtx.lineTo(Math.round((currentX + charWidth) * window.devicePixelRatio), strikethroughY);
                        offscreenCtx.stroke();
                    }
                    
                    offscreenCtx.restore();
                    k += 1;
                    currentX += charWidth + (j < line.text.length - 1 ? letterSpacing : 0);
                }
            }
            // 将 offscreenCanvas 内容合并到主canvas
            const x = Math.round(cellLeft * this.data.zoom);
            const y = Math.round(cellTop * this.data.zoom);
            this.ctx.drawImage(offscreenCanvas, 
                Math.round(x * window.devicePixelRatio), Math.round(y * window.devicePixelRatio), 
                Math.round(cellWidth * this.data.zoom * window.devicePixelRatio), 
                Math.round(cellHeight * this.data.zoom * window.devicePixelRatio));
        }
        
        this.ctx.restore();
    }

    /**
     * 绘制所有合并单元格背景
     *
     * 隐藏行列不绘制；合并单元格跨越隐藏列/行时只绘制可见部分。
     */
    public drawAllMergedCellBackground() {
        const MergedCells = this.data.getMergedCells();
        for (const mergedCell of MergedCells) {
            const [ startCell ] = mergedCell.split(':');  // 获取左上角单元格地址
            const { col, row } = this.data.getCellColAndRow(startCell);
            // 隐藏行列不绘制；跨隐藏列/行时只取可见部分
            const rect = this.getVisibleCellRect(col, row);
            if (!rect) continue;
            // 绘制合并单元格背景
            this.ctx.save();
            this.ctx.clearRect(Math.round(rect.left * this.data.zoom + 1),
                                Math.round(rect.top * this.data.zoom + 1),
                                Math.round(rect.width * this.data.zoom - 1), 
                                Math.round(rect.height * this.data.zoom - 1));
            const backgroundColor = this.data.values.find(cell => cell.cell === startCell)?.backgroundColor || DEFAULT_CELL_BG_COLOR;
            this.ctx.fillStyle = backgroundColor;
            this.fillRect(rect.left * this.data.zoom + 1,
                            rect.top * this.data.zoom + 1,
                            rect.width * this.data.zoom - 1, 
                            rect.height * this.data.zoom - 1);
            this.ctx.restore();
        }
    }

    /**
     * 处理鼠标按下事件
     * @param e 鼠标事件对象
     */
    public handleMouseDown(e: MouseEvent) {
        if (e.button !== 0 || e.target !== this.canvas) return;
        e.preventDefault();
        this.data.isMouseDown = true;
        const { offsetX, offsetY } = e;
        const { location, col, row } = this.data.getSheetCell(offsetX, offsetY);
        this.data.startLocation = location;
        if (this.data.startLocation === MouseLocation.IN_ANCHOR) {
            this.data.anchorSelection = this.data.selection;
            this.drawAnchorSelcetedRect();
        } else {
            this.data.activedCell = `${this.data.getColName(col)}${row}`;
            const mergedCells = this.data.getMergedCells();
            for (const mergedCell of mergedCells) {
                if (mergedCell.includes(`${this.data.getColName(col)}${row}`)) {
                    this.data.selection = mergedCell;
                    return;
                }
            }
            this.data.selection = `${this.data.activedCell}:${this.data.activedCell}`;
        }
        this.scrollCellIntoView(col, row);
    }

    /**
     * 处理鼠标移动事件
     * @param e 鼠标事件对象
     */
    public handleMouseMove(e: MouseEvent) {
        const { x: sheetX, y: sheetY } = this.getGeometry(e.clientX, e.clientY);
        const { location, col, row } = this.data.getSheetCell(sheetX, sheetY);
        this.updateCursorStyle(this.data.isMouseDown ? this.data.startLocation : location);  // 更新光标样式
        if (e.button !== 0 || !this.data.isMouseDown
            || this.data.isEditting
            || this.data.startLocation === MouseLocation.IN_All_SELECT
            || (this.data.overCell.col === col && this.data.overCell.row === row)
        ) return;
        this.data.overCell = { col, row };
        if (this.data.startLocation === MouseLocation.IN_COL_HEADER_SPLIT) {
            const newWidth = Math.round((e.clientX - this.data.MouseStartX) / this.data.zoom + this.data.startWidth);
            this.data.adjustColumnWidth(this.data.startCol, newWidth);
        } else if (this.data.startLocation === MouseLocation.IN_ROW_HEADER_SPLIT) {
            const newHeight = Math.round((e.clientY - this.data.MouseStartY) / this.data.zoom + this.data.startHeight);
            this.data.adjustRowHeight(this.data.startRow, newHeight);
        } else {
            this.handleAnchorSelect(sheetX, sheetY);
            this.data.setSelection(location, col, row);
            this.scrollCellIntoView(col, row);
        }        
    }

    /**
     * 处理鼠标松开事件
     * @param e 鼠标事件对象
     */
    public handleMouseUp(e: MouseEvent) {
        if (e.button !== 0 || !this.data.isMouseDown) return;
        this.handleAnchorSelected(); // 处理锚点选择完成之后鼠标抬起后操作
        this.data.isMouseDown = false;
        this.data.startLocation = '';
        this.data.overCell = { col: 0, row: 0 };
        this.data.MouseStartX = 0;
        this.data.startCol = 1;
        this.data.startWidth = 0;
        this.data.MouseStartY = 0;
        this.data.startRow = 1;
        this.data.startHeight = 0;
        if (this.data.brushMode) this.data.exeBrush();
    }

    /**
     * 处理双击事件
     * @param e 鼠标事件对象
     */
    public handleDoubleClick(e: MouseEvent) {
        if (e.button !== 0 || e.detail !== 2 || e.target !== this.canvas) return;
        this.data.isEditting = true;
    }

    /**
     * 处理锚点选择事件
     * @param sheetX 工作表坐标X
     * @param sheetY 工作表坐标Y
     */
    private handleAnchorSelect(sheetX: number, sheetY: number){ 
        if (this.data.startLocation !== MouseLocation.IN_ANCHOR)  return;
        let { col: startCol, row: startRow} = this.data.getCellColAndRow(this.data.selection.split(':')[0]);
        let { col: endCol, row: endRow} = this.data.getCellColAndRow(this.data.selection.split(':')[1]);
        const { left: anchorX, top: anchorY, width: anchorWidth, height: anchorHeight } = this.data.getAnchorRect();  // 获取锚点位置, 以锚点为中心，分四个象限，判断选择框是水平方向还是垂直方向
        const { col, row } = this.data.getSheetCell(sheetX, sheetY);
        if (Math.abs(sheetX - anchorX - anchorWidth / 2) >= Math.abs(sheetY - anchorY - anchorHeight / 2)) {  // 水平方向选择框
            if (col < startCol) {
                startCol = col;
            } else {
                endCol = col;
            }
        } else {   // 垂直方向选择框
            if (row < startRow) {
                startRow = row;
            } else {
                endRow = row;
            }
        }

        this.data.anchorSelection = `${this.data.getColName(startCol)}${startRow}:${this.data.getColName(endCol)}${endRow}`;

    }

    /**
     * 处理锚点选择完成之后鼠标抬起后操作
     */
    private handleAnchorSelected(){ 
        if (this.data.startLocation !== MouseLocation.IN_ANCHOR)  return;
        const _activeCol = this.data.getCellColAndRow(this.data.activedCell).col;
        const _activeRow = this.data.getCellColAndRow(this.data.activedCell).row;
        const selected = this.data.anchorSelection.split(':');
        const startCol = this.data.getCellColAndRow(selected[0]).col;
        const startRow = this.data.getCellColAndRow(selected[0]).row; 
        const { col, row } = this.data.getCellColAndRow(selected[1]);
        if (_activeCol < startCol || _activeCol > col || _activeRow < startRow || _activeRow > row) {
            this.data.activedCell = `${this.data.getColName(startCol)}${startRow}`;
        }
        this.data.selection = this.data.anchorSelection;
        this.data.anchorSelection = null;
    }
}
