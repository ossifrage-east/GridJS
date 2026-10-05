import { CELL_PADDING, DEFAULT_ANCHOR_SELECTED_BORDER_COLOR, DEFAULT_CELL_BG_COLOR, DEFAULT_FONT_FAMILY,
    DEFAULT_FONT_SIZE, DEFAULT_SELECTED_BG_COLOR, DEFAULT_SELECTED_BORDER_COLOR, DEFAULT_SELECTED_BORDER_WIDTH, GRID_LINE_COLOR,
    HEADER_BG_COLOR, TEXT_COLOR, SHEETS_NAME,
    DEFAULT_CELL_HEIGHT,
    ROW_HEADER_PADDING,
    MouseLocation} from "../constant";
import { Canvas, CanvasOptions } from "./canvas";
import { FullStateSnapshot } from "../../utils/undoManager";
import { Hidden } from "../assistant/hidden";


export class RowHeaderCanvas extends Canvas {
    /** 拖拽行高：待应用的目标高度（null 表示无待应用值） */
    private _pendingHeight: number | null = null;
    /** 拖拽行高：rAF 回调 id（null 表示无排程） */
    private _resizeRafId: number | null = null;
    /** 拖拽行高：最近一次实际应用的高度（用于跳过未跨像素的重复更新） */
    private _lastAppliedHeight: number | null = null;
    /** 拖拽行高：手势开始时的全量状态快照（mouseup 时提交撤销） */
    private _resizeUndoBefore: FullStateSnapshot | null = null;
    /** 隐藏行悬停指示按钮（上下双三角）：鼠标悬停到「分割线且下方紧邻隐藏行」时显示 */
    private hiddenIndicator: Hidden;

    constructor(options: CanvasOptions) {
        super(options);
        this.parentElement = options.parentElement;
        this.data = options.data;
        this.hiddenIndicator = new Hidden({ parentElement: options.parentElement, data: this.data, axis: 'row' });
        // 鼠标离开行头画布时隐藏指示按钮；移入按钮本身除外——按钮已可交互，
        // 由其自身 mouseleave 负责隐藏，避免「按钮显示在光标下方拦截事件 →
        // 画布 mouseleave → 隐藏 → 重现」的闪烁循环
        this.canvas.addEventListener('mouseleave', (e: MouseEvent) => {
            if (e.relatedTarget && this.hiddenIndicator.contains(e.relatedTarget as Node)) return;
            this.hiddenIndicator.hide();
        });
        this.initVisibleView(); // 初始化可见视图, 计算行头宽度
        this.init();
        this.setupContextMenu(); // 启用右键菜单
    }

    /**
     * 右键菜单动作执行后的刷新钩子
     *
     * 重绘行头（行号、选中高亮）。
     */
    protected refreshAfterContextMenu(): void {
        this.draw();
    }

    /**
     * 解析右键位置的单元格信息
     *
     * 行头画布的 X 坐标恒落在行头宽度内（正值），无法触发 getSheetCell 的
     * x<0 行头分支，因此强制传入 x=-1：命中 IN_ROW_HEADER，
     * 行号仍由 y 经正常匹配推导（含跳过隐藏行逻辑）。
     * @param {number} _x - 相对当前画布的X坐标（忽略）
     * @param {number} y - 相对当前画布的Y坐标
     * @returns {{ location: MouseLocation, col: number, row: number }} 行头区域信息
     */
    protected resolveContextMenuCell(_x: number, y: number): { location: MouseLocation, col: number, row: number } {
        return this.data.getSheetCell(-1, y);
    }

    /**
     * 初始化行头画布
     */
    private init() {
        this.canvas.setAttribute('tabindex', '0');// 确保 canvas 可聚焦
        this.canvas.id = SHEETS_NAME.ROW_HEADER;
        this.updateCanvasSize();
        this.parentElement.appendChild(this.canvas);
    }

    /**
     * 更新行头画布大小
     */
    public updateCanvasSize() {
        this.canvas.width = Math.round(this.data.visibleView.rowHeaderWidth) * window.devicePixelRatio;
        this.canvas.height = Math.round(this.data.visibleView.sheetHeight) * window.devicePixelRatio;
        this.canvas.style.width = this.data.visibleView.rowHeaderWidth + 'px';
        this.canvas.style.height = this.data.visibleView.sheetHeight + 'px';
        this.canvas.style.left = '0';
        this.canvas.style.top = `${this.data.visibleView.colHeaderHeight}px`;
        this.updateRowHeaders();
    }

    /**
     * 计算行头中行最大数的宽度
     */
    public calculateRowHeaderWidth(): number {
        const lastRow = this.data.rowHeaders.length;
        return this.ctx.measureText(`${lastRow}`).width;
    }

    /**
     * 初始化可见视图, 计算行头宽度
     * 表头数据由 sheets.ensureDefaultHeaders() 在创建画布前确保已初始化
     */
    private initVisibleView() {
        this.data.visibleView.rowHeaderWidth = Math.round(this.calculateRowHeaderWidth() + ROW_HEADER_PADDING * this.data.zoom);
    }

    /**
     * 更新行头画布大小
     */
    private updateRowHeaders() {
        const rows = Math.ceil(this.data.visibleView.sheetHeight / DEFAULT_CELL_HEIGHT / this.data.zoom);
        // 只更新新增的行
        if(rows > this.data.rowHeaders.length) this.data.rowHeaders.addRow(this.data.rowHeaders.length, rows);
    }

    /**
     * 绘制行头画布
     */
    public draw() {
        this.setupContext();
        this.ctx.save();
        this.clear(); // 清空画布
        // 背景不再整幅填充：改为循环内按可见行分段绘制，
        // 紧邻隐藏块时向内收缩，隐藏块边界（±1px 分割线位置）不填充背景形成缺口
        this.ctx.fillStyle = TEXT_COLOR;
        // 记录最后一个可见行的底缘（视图坐标），用于表格末尾不足整屏时的尾部背景填充
        let lastVisibleBottom = 0;
        const { startRow, endRow } = this.data.getVisibleRows();
        const { startRow: selectedStartRow, endRow: selectedEndRow } = this.getSelectedColsAndRows();
        // 绘制网格的水平线条
        for (let r = startRow - 1; r <= endRow; r++) {
            if(this.data.rowHeaders.getAt(r).isHidden) continue; // 跳过隐藏行
            // top 已由 recalcRowPositions 重算为紧凑排列（隐藏行不占位），无需手动补偿
            const y = (this.data.rowHeaders.getAt(r).top - this.data.rowHeaders.offsetHeight) * this.data.zoom;
            const rowHeight = this.data.rowHeaders.getAt(r).height * this.data.zoom;
            // 背景按可见行分段绘制：紧邻隐藏块时向内收缩，
            // 前边界线（bottomY-1）、中间缺口与后边界线（y+1）占位不填充，形成缺口标识隐藏区
            const prevHidden = this.data.rowHeaders.getAt(r - 1)?.isHidden === true;
            const nextHidden = this.data.rowHeaders.getAt(r + 1)?.isHidden === true;
            const bgStart = prevHidden ? y + 2 : y;  // 上一行隐藏：跳过中间缺口与后边界线（各 1px）
            const bgEnd = nextHidden ? y + rowHeight - 1 : y + rowHeight;  // 下一行隐藏：止于前边界线之前
            if (bgEnd > bgStart) {
                this.ctx.save();
                this.ctx.fillStyle = HEADER_BG_COLOR;
                this.fillRect(0, bgStart, this.data.visibleView.rowHeaderWidth, bgEnd - bgStart);
                this.ctx.restore();
            }
            lastVisibleBottom = y + rowHeight;
            // 绘制选中框
            if (r >= selectedStartRow - 1 && r < selectedEndRow) {
                // 绘制选中框的背景
                this.ctx.save();
                this.ctx.fillStyle = DEFAULT_SELECTED_BG_COLOR;
                this.fillRect(0, y,
                    this.data.visibleView.rowHeaderWidth,
                    this.data.rowHeaders.getAt(r).height * this.data.zoom);
                // 绘制选中框的右侧边框
                this.ctx.strokeStyle = DEFAULT_SELECTED_BORDER_COLOR;
                this.drawLine(this.data.visibleView.rowHeaderWidth, 
                    y,
                    this.data.visibleView.rowHeaderWidth,
                    y + this.data.rowHeaders.getAt(r).height * this.data.zoom,
                    DEFAULT_SELECTED_BORDER_WIDTH);
                this.ctx.restore();
            }
            this.ctx.save();
            // 绘制行头文本
            this.fillText(
                `${r + 1}`,
                this.data.visibleView.rowHeaderWidth / 2,
                y + this.data.rowHeaders.getAt(r).height * this.data.zoom / 2);
            this.ctx.restore();
            // 绘制网格的水平线条：紧邻隐藏行块时，前后边界线向外各偏移 1px，形成 2px 隐藏区视觉标识
            this.ctx.save();
            // 上一行隐藏时，本行顶边是隐藏块的后边界，向后（下）移 1px
            const lineY = this.data.rowHeaders.getAt(r - 1)?.isHidden ? y + 1 : y;
            this.drawLine(0, lineY, this.data.visibleView.rowHeaderWidth, lineY, 1);
            // 下一行隐藏时，本行底边是隐藏块的前边界，向前（上）移 1px 额外绘制
            if (this.data.rowHeaders.getAt(r + 1)?.isHidden) {
                const bottomY = y + this.data.rowHeaders.getAt(r).height * this.data.zoom;
                this.drawLine(0, bottomY - 1, this.data.visibleView.rowHeaderWidth, bottomY - 1, 1);
            }
            this.ctx.restore();
        }
        
        // 表格末尾不足整屏时，尾部区域背景整幅填充（不属于任何隐藏块边界）
        if (lastVisibleBottom < this.data.visibleView.sheetHeight) {
            this.ctx.fillStyle = HEADER_BG_COLOR;
            this.fillRect(0, lastVisibleBottom, this.data.visibleView.rowHeaderWidth,
                this.data.visibleView.sheetHeight - lastVisibleBottom);

            // 尾部虚拟行：序号按已有最大行号 +1 继续升序添加，绘制顶部分割线与序号，
            // 行高与虚拟网格线一致（默认行高 × zoom），与内容区补齐的横向线对齐。
            // 注意：此处 fillStyle 已被上方背景填充改为 HEADER_BG_COLOR，
            // fillText 使用当前 fillStyle 渲染，须显式恢复 TEXT_COLOR，否则序号与背景同色不可见
            const virtualStep = DEFAULT_CELL_HEIGHT * this.data.zoom;
            let virtualRowNo = this.data.rowHeaders.length;
            for (let y = lastVisibleBottom; y <= this.data.visibleView.sheetHeight; y += virtualStep) {
                virtualRowNo++;
                this.drawLine(0, y, this.data.visibleView.rowHeaderWidth, y, 1);
                this.ctx.save();
                this.ctx.fillStyle = TEXT_COLOR;
                this.fillText(`${virtualRowNo}`,
                    this.data.visibleView.rowHeaderWidth / 2,
                    y + virtualStep / 2);
                this.ctx.restore();
            }
        }
        this.ctx.restore();
    }

    /**
     * 处理鼠标按下事件
     * @param e 鼠标事件对象
     */
    public handleMouseDown(e: MouseEvent) {
        if (e.button !== 0 || e.target !== this.canvas) return;
        e.preventDefault();
        this.data.isMouseDown = true;
        const { offsetY } = e;
        const { location, col, row } = this.data.rowHeaders.getRowHeaderCell(Math.round(offsetY / this.data.zoom)); 
        this.data.startLocation = location;
        if (this.data.startLocation === MouseLocation.IN_ROW_HEADER_SPLIT) {
            this.data.MouseStartY = e.clientY;
            this.data.startRow = row;
            this.data.startHeight = this.data.rowHeaders.getAt(row - 1).height;
            this.hiddenIndicator.hide();  // 开始拖拽行高，隐藏指示按钮避免遮挡
            // 开始连续撤销捕获：整个拖拽手势折叠为一条撤销记录
            this._resizeUndoBefore = this.data.beginContinuousUndo();
            this._lastAppliedHeight = this.data.startHeight;
        }
        this.data.activedCell = `${this.data.getColName(col)}${row}`;
        this.data.selection = `${this.data.activedCell}:${this.data.getColName(this.data.colHeaders.length)}${row}`;
        this.scrollCellIntoView(col, row);
    }

    /**
     * 处理鼠标移动事件
     * @param e 鼠标事件对象
     */
    public handleMouseMove(e: MouseEvent) {
        // 拖拽行高分界线：一旦开始，即使鼠标移出画布也继续跟踪
        // （clientY 不依赖画布边界，绕过 e.target 检查），直到 mouseup
        if (this.data.isMouseDown && this.data.startLocation === MouseLocation.IN_ROW_HEADER_SPLIT) {
            this.updateCursorStyle(MouseLocation.IN_ROW_HEADER_SPLIT);
            this._queueRowHeight(e.clientY);
            return;
        }
        const { x: rowHeaderX, y: rowHeaderY } = this.getGeometry(e.clientX, e.clientY);
        // 悬停显隐：鼠标位于行分割线附近且该分割线下方紧邻隐藏行时显示指示按钮，否则隐藏；
        // 返回值表示鼠标是否落在指示按钮上（用于切换默认光标）
        const overIndicator = this._updateHiddenIndicator(rowHeaderX, rowHeaderY);
        const { location, col, row } = this.data.rowHeaders.getRowHeaderCell(Math.round(rowHeaderY / this.data.zoom));
        if (overIndicator) {
            this.canvas.style.cursor = 'default';  // 悬停在隐藏指示按钮上，显示默认光标
        } else {
            this.updateCursorStyle(this.data.isMouseDown ? this.data.startLocation : location);  // 更新光标样式
        }
        if (e.button !== 0 || !this.data.isMouseDown
            || e.target !== this.canvas
            || this.data.startLocation === MouseLocation.IN_All_SELECT
            || (this.data.overCell.col === col && this.data.overCell.row === row)
        ) return;
        if (this.data.startLocation === MouseLocation.IN_ROW_HEADER
            && (this.data.overCell.col !== col && this.data.overCell.row !== row) ) {
            this.data.overCell = { col, row };
            this.data.setSelection(location, col, row);
        }
    }

    /**
     * 悬停显隐检测：隐藏行指示按钮
     *
     * 鼠标位于某行分割线附近（视口 3px 容差）且该分割线下方紧邻隐藏行时，
     * 在行头右缘、垂直居中于分割线的位置显示上下双三角指示按钮；否则隐藏。
     *
     * 按钮已可交互（可点击取消隐藏对应行块）：鼠标进入按钮后画布不再收到
     * mousemove，因此仍用几何判断返回「按钮显示时鼠标是否落在其矩形范围内」，
     * 作为鼠标尚未进入按钮时的光标切换兜底（垂直方向按钮中心对齐分割线，
     * ±7px 覆盖 3px 容差，必然命中，故只需判 X 落点）。
     * @param {number} rowHeaderX - 相对行头画布的 X 坐标（CSS 像素）
     * @param {number} rowHeaderY - 相对行头画布的 Y 坐标（CSS 像素）
     * @returns {boolean} 鼠标是否悬停在指示按钮上（用于切换默认光标）
     * @private
     */
    private _updateHiddenIndicator(rowHeaderX: number, rowHeaderY: number): boolean {
        const zoom = this.data.zoom;
        const rows = this.data.rowHeaders;
        const yData = rowHeaderY / zoom + rows.offsetHeight;  // 换算为数据坐标
        const tolerance = 3 / zoom;  // 视口 3px 容差换算到数据坐标
        let splitTop: number | null = null;  // 命中的分割线（数据坐标，= 前一可见行 bottom = 隐藏块 top）
        let hiddenRange: { start: number; end: number } | null = null;  // 按钮对应的隐藏行块（1-based，含端点）
        for (let r = 0; r < rows.length - 1; r++) {
            const header = rows.getAt(r);
            if (!header) break;
            if (header.top > yData + tolerance) break;  // 已越过鼠标位置，后续行更靠下
            if (header.isHidden) continue;  // 隐藏行 top 被前一可见行覆盖，不构成分割线
            const bottom = header.top + header.height;
            if (Math.abs(bottom - yData) <= tolerance) {
                // 分割线下方紧邻隐藏行才显示指示按钮
                if (rows.getAt(r + 1)?.isHidden) {
                    splitTop = bottom;
                    // 向后扫描连续隐藏行块，换算为 1-based 含端点范围，供点击按钮时取消隐藏
                    let hiddenEnd0 = r + 1;  // 隐藏块最后一行的 0-based 索引
                    while (hiddenEnd0 + 1 < rows.length && rows.getAt(hiddenEnd0 + 1)?.isHidden) hiddenEnd0++;
                    hiddenRange = { start: r + 2, end: hiddenEnd0 + 1 };
                }
                break;
            }
        }
        if (splitTop === null) {
            this.hiddenIndicator.hide();
            return false;
        }
        const splitViewY = (splitTop - rows.offsetHeight) * zoom;  // 分割线的画布视图坐标
        // 按钮在父容器内定位：行头画布偏移 left=0/top=colHeaderHeight，
        // 水平贴行头右缘（按钮宽 10 + 2px 边距），垂直中心对齐分割线（按钮高 22）
        const left = this.data.visibleView.rowHeaderWidth - 12;
        const top = this.data.visibleView.colHeaderHeight + splitViewY - 11;
        // splitTop 非 null 时 hiddenRange 必已在同一分支同步赋值
        this.hiddenIndicator.show(left, top, hiddenRange!);
        // 判断鼠标 X 落点是否在按钮水平范围内（[rowHeaderWidth-12, rowHeaderWidth-2]）
        return rowHeaderX >= this.data.visibleView.rowHeaderWidth - 12;
    }

    /**
     * 排程一次行高应用（requestAnimationFrame 节流）
     * 同一帧内的多次 mousemove 只应用最后一次目标高度；
     * 高度与上次应用值相同（未跨像素）时直接跳过，避免无意义的重算与全量重绘。
     * @param {number} clientY - 鼠标视口 Y 坐标
     */
    private _queueRowHeight(clientY: number): void {
        const MIN_ROW_HEIGHT = 24;  // 与 DataCollection.adjustRowHeight 的最小行高一致
        const newHeight = Math.max(MIN_ROW_HEIGHT,
            Math.round((clientY - this.data.MouseStartY) / this.data.zoom + this.data.startHeight));
        if (newHeight === this._lastAppliedHeight) return;
        this._pendingHeight = newHeight;
        if (this._resizeRafId !== null) return;
        this._resizeRafId = requestAnimationFrame(() => {
            this._resizeRafId = null;
            if (this._pendingHeight === null) return;
            this._lastAppliedHeight = this._pendingHeight;
            this.data.adjustRowHeight(this.data.startRow, this._pendingHeight);
            this._pendingHeight = null;
        });
    }

    /**
     * 结束行高拖拽：取消未执行的 rAF、同步应用最终高度并提交连续撤销
     * 由 handleMouseUp 调用；不依赖 isMouseDown（sheetCanvas 可能已先重置共享状态）。
     */
    private _finishRowSplitDrag(): void {
        if (this._resizeRafId !== null) {
            cancelAnimationFrame(this._resizeRafId);
            this._resizeRafId = null;
        }
        const pending = this._pendingHeight;
        this._pendingHeight = null;
        if (pending !== null && pending !== this._lastAppliedHeight) {
            this._lastAppliedHeight = pending;
            this.data.adjustRowHeight(this.data.startRow, pending);  // 同步应用最终高度，保证落点精确
        }
        if (this._resizeUndoBefore !== null) {
            this.data.endContinuousUndo(this._resizeUndoBefore, '调整行高');
            this._resizeUndoBefore = null;
        }
    }

    /**
     * 处理鼠标松开事件
     * @param e 鼠标事件对象
     */
    public handleMouseUp(e: MouseEvent) {
        // 拖拽行高收尾（置于最前：不依赖 isMouseDown，避免被 sheetCanvas 先重置共享状态而漏提交）
        this._finishRowSplitDrag();
        if (e.button !== 0 || !this.data.isMouseDown) return;
        e.preventDefault();
        this.data.isMouseDown = false;
        this.data.startLocation = '';
        this.data.MouseStartY = 0;
        this.data.startRow = 1;
        this.data.startHeight = 0;
    }

    /**
     * 处理双击事件
     * @param e 鼠标事件对象
     */
    public handleDoubleClick(e: MouseEvent) {
        const { offsetY: rowHeaderY } = e;
        const { location, col, row } = this.data.rowHeaders.getRowHeaderCell(Math.round(rowHeaderY / this.data.zoom));
        if (e.button !== 0 || e.detail !== 2
            || e.target !== this.canvas
            || location !== MouseLocation.IN_ROW_HEADER_SPLIT
        ) return;
        
        this.autoAdjustRowHeight(row);
    }

    /**
     * 自动调整行高以适应内容高度
     * @param row 行索引
     */
    private autoAdjustRowHeight(row: number): void {
        if (row < 1 || row > this.data.rowHeaders.length) return;
        
        // 获取该行所有单元格内容
        const cellsInRow = this.data.values.filter(v => {
            const { row: cellRow } = this.data.getCellColAndRow(v.cell);
            return cellRow === row && v.cell !== '';
        });
        
        // 计算最大文本高度
        let minHeight = 20;
        for (const cell of cellsInRow) {
             const { charsHeight } = this.measureCellText(cell.cell);
            minHeight = Math.max(minHeight, charsHeight);
        }
        
        // 调整行高
        this.data.adjustRowHeight(row, Math.round(minHeight / this.data.zoom));        
    }
}