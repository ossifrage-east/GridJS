import { CELL_PADDING, DEFAULT_ANCHOR_SELECTED_BORDER_COLOR, DEFAULT_CELL_BG_COLOR, DEFAULT_FONT_FAMILY,
    DEFAULT_FONT_SIZE, DEFAULT_SELECTED_BG_COLOR, DEFAULT_SELECTED_BORDER_COLOR, DEFAULT_SELECTED_BORDER_WIDTH, GRID_LINE_COLOR,
    HEADER_BG_COLOR, TEXT_COLOR, SHEETS_NAME,
    DEFAULT_CELL_HEIGHT,
    DEFAULT_CELL_WIDTH,
    MouseLocation} from "../constant";
import { Canvas, CanvasOptions } from "./canvas";
import { ContinuousUndoSnapshot } from "../../utils/undoManager";
import { Hidden } from "../assistant/hidden";


export class ColHeaderCanvas extends Canvas {
    /** 拖拽列宽：待应用的目标宽度（null 表示无待应用值） */
    private _pendingWidth: number | null = null;
    /** 拖拽列宽：rAF 回调 id（null 表示无排程） */
    private _resizeRafId: number | null = null;
    /** 拖拽列宽：最近一次实际应用的宽度（用于跳过未跨像素的重复更新） */
    private _lastAppliedWidth: number | null = null;
    /** 拖拽列宽：手势开始时的基准快照（小表全量/大表 headers 补丁档，mouseup 时提交撤销） */
    private _resizeUndoBefore: ContinuousUndoSnapshot | null = null;
    /** 拖拽列宽：本次手势是否真正改变过宽度（用于把「纯单击」与「拖拽后回到原宽」区分开） */
    private _splitMoved = false;
    /** 拖拽列宽：手势对应的列号（分界线左侧那一列，按下时记录，不依赖随后可能被重置的共享状态） */
    private _splitCol = 0;
    /**
     * 拖拽列宽时要批量调整的列范围（null = 沿用原来的「选中该列并拖拽」语义）
     *
     * 仅在「当前选区跨多列且所点分界线属于选中列」时安排：此时按下分界线
     * 不改变选区，拖拽则对选中的全部列设成同一宽度（双击自动调宽同样作用于该范围）。
     */
    private _splitResizeRange: { startCol: number, endCol: number } | null = null;
    /** 隐藏列悬停指示按钮（左右双三角）：鼠标悬停到「分割线且右侧紧邻隐藏列」时显示 */
    private hiddenIndicator: Hidden;

    constructor(options: CanvasOptions) {
        super(options);
        this.parentElement = options.parentElement;
        this.data = options.data;
        this.hiddenIndicator = new Hidden({ parentElement: options.parentElement, data: this.data, axis: 'col' });
        // 鼠标离开列头画布时隐藏指示按钮；移入按钮本身除外——按钮已可交互，
        // 由其自身 mouseleave 负责隐藏，避免「按钮显示在光标下方拦截事件 →
        // 画布 mouseleave → 隐藏 → 重现」的闪烁循环
        this.canvas.addEventListener('mouseleave', (e: MouseEvent) => {
            if (e.relatedTarget && this.hiddenIndicator.contains(e.relatedTarget as Node)) return;
            this.hiddenIndicator.hide();
        });
        this.initVisibleView(); // 初始化可见视图, 计算列头高度
        this.init();
        this.setupContextMenu(); // 启用右键菜单
    }

    /**
     * 右键菜单动作执行后的刷新钩子
     *
     * 重绘列头（列名、选中高亮）。
     */
    protected refreshAfterContextMenu(): void {
        this.draw();
    }

    /**
     * 解析右键位置的单元格信息
     *
     * 列头画布的 Y 坐标恒落在列头高度内（正值），无法触发 getSheetCell 的
     * y<0 列头分支，因此强制传入 y=-1：命中 IN_COL_HEADER，
     * 列号仍由 x 经正常匹配推导（含跳过隐藏列逻辑）。
     *
     * **必须传 `rawGeometry`（纯几何）**：否则 getSheetCell 会把列号吸附到合并锚点。
     * 列头上解析出的行恒为第 1 行，而首行最常被横向合并（如 `A1:D1` 标题横幅），
     * 于是点 B / C / D 列头都会被解析成 `B1` / `C1` / `D1` 再吸附到锚点 `A1`
     * —— 右键任何列都只选中 A 列（实测 test_sample.xls）。点某列列头就该选中该列，
     * 与左键的 getColHeaderCell（本就不做合并吸附）语义一致。
     * @param {number} x - 相对当前画布的X坐标
     * @param {number} _y - 相对当前画布的Y坐标（忽略）
     * @returns {{ location: MouseLocation, col: number, row: number }} 列头区域信息
     */
    protected resolveContextMenuCell(x: number, _y: number): { location: MouseLocation, col: number, row: number } {
        return this.data.getSheetCell(x, -1, { rawGeometry: true });
    }

    /**
     * 初始化列头画布
     */
    private init() {
        this.canvas.setAttribute('tabindex', '0');// 确保 canvas 可聚焦
        this.canvas.id = SHEETS_NAME.COL_HEADER;
        this.updateCanvasSize();
        this.parentElement.appendChild(this.canvas);
    }

    /** 
     * 更新列头画布大小
     */
    public updateCanvasSize() {
        this.canvas.width = Math.round(this.data.visibleView.sheetWidth) * window.devicePixelRatio;
        this.canvas.height = Math.round(this.data.visibleView.colHeaderHeight * window.devicePixelRatio);
        this.canvas.style.width = `${this.data.visibleView.sheetWidth}px`;
        this.canvas.style.height = `${this.data.visibleView.colHeaderHeight}px`;
        this.canvas.style.left = `${this.data.visibleView.rowHeaderWidth}px`;
        this.canvas.style.top = '0';
        this.updateColHeaders();
    }

    /**
     * 初始化可见视图, 计算列头高度
     * 表头数据由 sheets.ensureDefaultHeaders() 在创建画布前确保已初始化
     */
    private initVisibleView() {
        this.data.visibleView.colHeaderHeight = Math.round(DEFAULT_CELL_HEIGHT * this.data.zoom);
    }

    /**
     * 更新列头画布大小
     */
    private updateColHeaders() {
        const cols = Math.ceil(this.data.visibleView.sheetWidth / DEFAULT_CELL_WIDTH / this.data.zoom);
        // 只更新新增的列
        if(cols > this.data.colHeaders.length) this.data.colHeaders.addCol(this.data.colHeaders.length, cols);
    }

    /**
     * 绘制列头画布
     */
    public draw() {
        this.setupContext();
        this.ctx.save();
        this.clear(); // 清空画布
        // 背景不再整幅填充：改为循环内按可见列分段绘制，
        // 紧邻隐藏块时向内收缩，隐藏块边界（±1px 分割线位置）不填充背景形成缺口
        this.ctx.fillStyle = TEXT_COLOR;
        // 记录最后一个可见列的右缘（视图坐标），用于表格末尾不足整屏时的尾部背景填充
        let lastVisibleRight = 0;
        const { startCol, endCol } = this.data.getVisibleCols();
        const { startCol: selectedStartCol, endCol: selectedEndCol } = this.getSelectedColsAndRows();
        // 绘制网格的垂直线条
        // 注意 getVisibleCols 返回 1-based 列号（含端点），此处 c 是 0-based 下标
        for (let c = startCol - 1; c < endCol; c++) {
            if(this.data.colHeaders.getAt(c).isHidden) continue; // 跳过隐藏列
            // left 已由 recalcColPositions 重算为紧凑排列（隐藏列不占位），无需手动补偿
            const x = (this.data.colHeaders.getAt(c).left - this.data.colHeaders.offsetWidth) * this.data.zoom;
            const colWidth = this.data.colHeaders.getAt(c).width * this.data.zoom;
            // 背景按可见列分段绘制：紧邻隐藏块时向内收缩，
            // 前边界线（rightX-1）、中间缺口与后边界线（x+1）占位不填充，形成缺口标识隐藏区
            const prevHidden = this.data.colHeaders.getAt(c - 1)?.isHidden === true;
            const nextHidden = this.data.colHeaders.getAt(c + 1)?.isHidden === true;
            const bgStart = prevHidden ? x + 2 : x;  // 上一列隐藏：跳过中间缺口与后边界线（各 1px）
            const bgEnd = nextHidden ? x + colWidth - 1 : x + colWidth;  // 下一列隐藏：止于前边界线之前
            if (bgEnd > bgStart) {
                this.ctx.save();
                this.ctx.fillStyle = HEADER_BG_COLOR;
                this.fillRect(bgStart, 0, bgEnd - bgStart, this.data.visibleView.colHeaderHeight);
                this.ctx.restore();
            }
            lastVisibleRight = x + colWidth;
            // 绘制选中框
            if (c >= selectedStartCol - 1 && c < selectedEndCol) {
                // 绘制选中框的背景
                this.ctx.save();
                this.ctx.fillStyle = DEFAULT_SELECTED_BG_COLOR;
                this.fillRect(
                    x,
                    0,
                    this.data.colHeaders.getAt(c).width * this.data.zoom,
                    this.data.visibleView.colHeaderHeight);
                // 绘制选中框的底部边框
                this.ctx.strokeStyle = DEFAULT_SELECTED_BORDER_COLOR;
                this.drawLine(x,
                    this.data.visibleView.colHeaderHeight,
                    x + this.data.colHeaders.getAt(c).width * this.data.zoom, 
                    this.data.visibleView.colHeaderHeight,
                    DEFAULT_SELECTED_BORDER_WIDTH);
                this.ctx.restore();
            }
            // 绘制列头文本
            this.fillText(
                this.data.getColName(c + 1),
                x + this.data.colHeaders.getAt(c).width * this.data.zoom / 2,
                this.data.visibleView.colHeaderHeight / 2);
            // 绘制网格的垂直线条：紧邻隐藏列块时，前后边界线向外各偏移 1px，形成 2px 隐藏区视觉标识
            this.ctx.save();
            this.ctx.strokeStyle = GRID_LINE_COLOR;
            // 上一列隐藏时，本列左边是隐藏块的后边界，向后（右）移 1px
            const lineX = this.data.colHeaders.getAt(c - 1)?.isHidden ? x + 1 : x;
            this.drawLine(lineX, 0, lineX, this.data.visibleView.colHeaderHeight, 1);
            // 下一列隐藏时，本列右边是隐藏块的前边界，向前（左）移 1px 额外绘制
            if (this.data.colHeaders.getAt(c + 1)?.isHidden) {
                const rightX = x + this.data.colHeaders.getAt(c).width * this.data.zoom;
                this.drawLine(rightX - 1, 0, rightX - 1, this.data.visibleView.colHeaderHeight, 1);
            }
            this.ctx.restore();
        }
        // 表格末尾不足整屏时，尾部区域背景整幅填充（不属于任何隐藏块边界）
        if (lastVisibleRight < this.data.visibleView.sheetWidth) {
            this.ctx.fillStyle = HEADER_BG_COLOR;
            this.fillRect(lastVisibleRight, 0, this.data.visibleView.sheetWidth - lastVisibleRight,
                this.data.visibleView.colHeaderHeight);

            // 尾部虚拟列：列名按已有最大列号 +1 继续升序添加（getColName 生成字母序号），
            // 绘制左侧分割线与列名，列宽与虚拟网格线一致（默认列宽 × zoom），与内容区补齐的竖向线对齐。
            // 注意：此处 fillStyle 已被上方背景填充改为 HEADER_BG_COLOR，strokeStyle 也非网格色，
            // 须显式设置 TEXT_COLOR / GRID_LINE_COLOR，否则列名与背景同色不可见、分割线颜色不对
            this.ctx.save();
            this.ctx.strokeStyle = GRID_LINE_COLOR;
            this.ctx.fillStyle = TEXT_COLOR;
            const virtualStepX = DEFAULT_CELL_WIDTH * this.data.zoom;
            let virtualColNo = this.data.colHeaders.length;
            for (let x = lastVisibleRight; x <= this.data.visibleView.sheetWidth; x += virtualStepX) {
                virtualColNo++;
                this.drawLine(x, 0, x, this.data.visibleView.colHeaderHeight, 1);
                this.fillText(this.data.getColName(virtualColNo),
                    x + virtualStepX / 2,
                    this.data.visibleView.colHeaderHeight / 2);
            }
            this.ctx.restore();
        }
        this.ctx.restore();
    }

    /**
     * 绘制列头高亮
     */
    public drawHighlightHeader() {
        const { startCol, endCol } = this.getSelectedColsAndRows();
        this.ctx.save();
        // 绘制高亮背景
        this.ctx.fillStyle = DEFAULT_SELECTED_BG_COLOR;
        this.fillRect(
            (this.data.colHeaders.getAt(startCol - 1).left - this.data.colHeaders.offsetWidth) * this.data.zoom,
            0,
            (this.data.colHeaders.getAt(endCol).left - this.data.colHeaders.offsetWidth) * this.data.zoom - (this.data.colHeaders.getAt(startCol - 1).left - this.data.colHeaders.offsetWidth) * this.data.zoom,
            this.data.visibleView.colHeaderHeight);
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
        const { offsetX } = e; // 获取colHeaderCanvas中的坐标X
        const { location, col, row } = this.data.colHeaders.getColHeaderCell(Math.round(offsetX / this.data.zoom));
        this.data.startLocation = location;
        if (this.data.startLocation === MouseLocation.IN_COL_HEADER_SPLIT) {
            this.data.MouseStartX = e.clientX;
            this.data.startCol = col;
            this.data.startWidth = this.data.colHeaders.getAt(col - 1).width;
            this.hiddenIndicator.hide();  // 开始拖拽列宽，隐藏指示按钮避免遮挡
            // 开始连续撤销捕获：整个拖拽手势折叠为一条撤销记录
            this._resizeUndoBefore = this.data.beginContinuousUndo();
            this._lastAppliedWidth = this.data.startWidth;
            this._splitMoved = false;
            this._splitCol = col;
            // 选中多列后点其中一列的分界线：选区保持不变（不重选该列、不挪活动格），
            // 拖拽时对选中的全部列设成同一宽度（见 _applyColWidth）。
            // 条件：选区跨多列，且所点分界线属于选中列（分界线的 col 是左侧那一列）。
            const { startCol: selStartCol, endCol: selEndCol } = this.getSelectedColsAndRows();
            this._splitResizeRange = selEndCol > selStartCol && col >= selStartCol && col <= selEndCol
                ? { startCol: selStartCol, endCol: selEndCol }
                : null;
            if (this._splitResizeRange) return;
        }
        // 活动格**跳过合并单元格**：列头点击恒取第 1 行，而首行最常被横向合并
        // （如表头横幅 A1:D1）；活动格落进合并区后，选中框上挖的透明洞会变成整个
        // 合并块那么大（getActivedRect 走合并锚点语义），整列高亮被挖掉一大片。
        this.data.activedCell = this.data.resolveHeaderActiveCell(col, row, 'col');
        // 列头点击 = **精确的**整列选区（`expandMerges = false`）
        // 【语义】点第 C 列头要的就是「C 列这一整列」，该列中间有没有横向合并
        // （如 C1:E1）与用户意图无关；若按合并区域归一化，选中的会变成 C~E 列，
        // 于是「查找范围」「复制内容」「删除列」全都跟着错一整块。
        // 高亮的正确性不依赖归一化：getSelectedRect 只用网格矩形取端点
        // （见那里的说明），因此单列选区照样精确地贴在所点的那一列上。
        this.data.setSelectedRange(col, 1, col, this.data.rowHeaders.length, false);
        // 滚到**活动格**（而非点击处的第 1 行）：活动格可能已被挪到别的行，
        // 保证它连同选中框一起落在屏幕里。
        const active = this.data.getCellColAndRow(this.data.activedCell);
        this.scrollCellIntoView(active.col, active.row);
    }

    /**
     * 处理鼠标移动事件
     * @param e 鼠标事件对象
     */
    public handleMouseMove(e: MouseEvent) {
        // 拖拽列宽分界线：一旦开始，即使鼠标移出画布也继续跟踪
        // （clientX 不依赖画布边界，绕过 e.target 检查），直到 mouseup
        if (this.data.isMouseDown && this.data.startLocation === MouseLocation.IN_COL_HEADER_SPLIT) {
            this.updateCursorStyle(MouseLocation.IN_COL_HEADER_SPLIT);
            this._queueColWidth(e.clientX);
            return;
        }
        const { x: colHeaderX, y: colHeaderY } = this.getGeometry(e.clientX, e.clientY);
        // 悬停显隐：鼠标位于列分割线附近且该分割线右侧紧邻隐藏列时显示指示按钮，否则隐藏；
        // 返回值表示鼠标是否落在指示按钮上（用于切换默认光标）
        const overIndicator = this._updateHiddenIndicator(colHeaderX, colHeaderY);
        const { location, col, row } = this.data.colHeaders.getColHeaderCell(Math.round(colHeaderX / this.data.zoom));
        if (overIndicator) {
            this.canvas.style.cursor = 'default';  // 悬停在隐藏指示按钮上，显示默认光标
        } else {
            this.updateCursorStyle(this.data.isMouseDown ? this.data.startLocation : location);  // 更新光标样式
        }
        if (e.button !== 0 || !this.data.isMouseDown
            || e.target !== this.canvas
            || this.data.startLocation === MouseLocation.IN_All_SELECT
        ) return;
        if (this.data.startLocation === MouseLocation.IN_COL_HEADER) {
            this.data.overCell = { col, row };
            this.data.setSelection(location, col, row);
        }
    }

    /**
     * 悬停显隐检测：隐藏列指示按钮
     *
     * 鼠标位于某列分割线附近（视口 3px 容差）且该分割线右侧紧邻隐藏列时，
     * 在列头下缘、水平居中于分割线的位置显示左右双三角指示按钮；否则隐藏。
     *
     * 按钮已可交互（可点击取消隐藏对应列块）：鼠标进入按钮后画布不再收到
     * mousemove，因此仍用几何判断返回「按钮显示时鼠标是否落在其矩形范围内」，
     * 作为鼠标尚未进入按钮时的光标切换兜底（水平方向按钮中心对齐分割线，
     * ±7px 覆盖 3px 容差，必然命中，故只需判 Y 落点）。
     * @param {number} colHeaderX - 相对列头画布的 X 坐标（CSS 像素）
     * @param {number} colHeaderY - 相对列头画布的 Y 坐标（CSS 像素）
     * @returns {boolean} 鼠标是否悬停在指示按钮上（用于切换默认光标）
     * @private
     */
    private _updateHiddenIndicator(colHeaderX: number, colHeaderY: number): boolean {
        const zoom = this.data.zoom;
        const cols = this.data.colHeaders;
        const xData = colHeaderX / zoom + cols.offsetWidth;  // 换算为数据坐标
        const tolerance = 3 / zoom;  // 视口 3px 容差换算到数据坐标
        let splitLeft: number | null = null;  // 命中的分割线（数据坐标，= 前一可见列 right = 隐藏块 left）
        let hiddenRange: { start: number; end: number } | null = null;  // 按钮对应的隐藏列块（1-based，含端点）
        for (let c = 0; c < cols.length - 1; c++) {
            const header = cols.getAt(c);
            if (!header) break;
            if (header.left > xData + tolerance) break;  // 已越过鼠标位置，后续列更靠右
            if (header.isHidden) continue;  // 隐藏列 left 被前一可见列覆盖，不构成分割线
            const right = header.left + header.width;
            if (Math.abs(right - xData) <= tolerance) {
                // 分割线右侧紧邻隐藏列才显示指示按钮
                if (cols.getAt(c + 1)?.isHidden) {
                    splitLeft = right;
                    // 向后扫描连续隐藏列块，换算为 1-based 含端点范围，供点击按钮时取消隐藏
                    let hiddenEnd0 = c + 1;  // 隐藏块最后一列的 0-based 索引
                    while (hiddenEnd0 + 1 < cols.length && cols.getAt(hiddenEnd0 + 1)?.isHidden) hiddenEnd0++;
                    hiddenRange = { start: c + 2, end: hiddenEnd0 + 1 };
                }
                break;
            }
        }
        if (splitLeft === null) {
            this.hiddenIndicator.hide();
            return false;
        }
        const splitViewX = (splitLeft - cols.offsetWidth) * zoom;  // 分割线的画布视图坐标
        // 按钮在父容器内定位：列头画布偏移 left=rowHeaderWidth/top=0，
        // 垂直贴列头下缘（按钮高 10 + 2px 边距），水平中心对齐分割线（按钮宽 22）
        const left = this.data.visibleView.rowHeaderWidth + splitViewX - 11;
        const top = this.data.visibleView.colHeaderHeight - 12;
        // splitLeft 非 null 时 hiddenRange 必已在同一分支同步赋值
        this.hiddenIndicator.show(left, top, hiddenRange!);
        // 判断鼠标 Y 落点是否在按钮垂直范围内（[colHeaderHeight-12, colHeaderHeight-2]）
        return colHeaderY >= this.data.visibleView.colHeaderHeight - 12;
    }

    /**
     * 排程一次列宽应用（requestAnimationFrame 节流）
     * 同一帧内的多次 mousemove 只应用最后一次目标宽度；
     * 宽度与上次应用值相同（未跨像素）时直接跳过，避免无意义的重算与全量重绘。
     * @param {number} clientX - 鼠标视口 X 坐标
     */
    private _queueColWidth(clientX: number): void {
        const MIN_COL_WIDTH = 24;  // 与 DataCollection.adjustColumnWidth 的最小列宽一致
        const newWidth = Math.max(MIN_COL_WIDTH,
            Math.round((clientX - this.data.MouseStartX) / this.data.zoom + this.data.startWidth));
        if (newWidth === this._lastAppliedWidth) return;
        this._pendingWidth = newWidth;
        if (this._resizeRafId !== null) return;
        this._resizeRafId = requestAnimationFrame(() => {
            this._resizeRafId = null;
            const pending = this._pendingWidth;
            if (pending === null) return;
            this._lastAppliedWidth = pending;
            this._splitMoved = true;
            this._applyColWidth(pending);
            this._pendingWidth = null;
        });
    }

    /**
     * 应用一次列宽：多列选区时把选中的全部列设为同一宽度，否则只调整被拖的那一列
     * @param {number} width - 目标列宽
     */
    private _applyColWidth(width: number): void {
        const range = this._splitResizeRange;
        if (range) {
            this.data.adjustColumnWidths(range.startCol, range.endCol, width);
        } else {
            this.data.adjustColumnWidth(this._splitCol, width);
        }
    }

    /**
     * 结束列宽拖拽：取消未执行的 rAF、同步应用最终宽度、补齐右侧列并提交连续撤销
     * 由 handleMouseUp 调用；不依赖 isMouseDown（sheetCanvas 可能已先重置共享状态）。
     *
     * **单击分界线不改变列宽**（自动调宽只由左键双击触发，见 {@link handleDoubleClick}），
     * 因此这里不再区分「单击」与「拖拽」。
     */
    private _finishColSplitDrag(): void {
        if (this._resizeRafId !== null) {
            cancelAnimationFrame(this._resizeRafId);
            this._resizeRafId = null;
        }
        const pending = this._pendingWidth;
        this._pendingWidth = null;
        const pendingApplied = pending !== null && pending !== this._lastAppliedWidth;
        if (pendingApplied) {
            this._lastAppliedWidth = pending;
            this._applyColWidth(pending);  // 同步应用最终宽度，保证落点精确
        }
        const range = this._splitResizeRange;
        this._splitResizeRange = null;
        if (this._resizeUndoBefore !== null) {
            this.data.endContinuousUndo(this._resizeUndoBefore, '调整列宽');
            this._resizeUndoBefore = null;
        }
        // 右侧补列：列宽调小后右边会露白，补齐/裁剪到恰好填满视口（本次调整的列不参与裁剪）
        if (pendingApplied || this._splitMoved) {
            if (range) this.data.fillColsToViewport(range.startCol, range.endCol);
            else if (this._splitCol > 0) this.data.fillColsToViewport(this._splitCol, this._splitCol);
        }
    }

    /**
     * 处理鼠标松开事件
     * @param e 鼠标事件对象
     */
    public handleMouseUp(e: MouseEvent) {
        // 拖拽列宽收尾（置于最前：不依赖 isMouseDown，避免被 sheetCanvas 先重置共享状态而漏提交）
        this._finishColSplitDrag();
        if (e.button !== 0 || !this.data.isMouseDown) return;
        this.data.isMouseDown = false;
        this.data.startLocation = '';
        this.data.MouseStartX = 0;
        this.data.startCol = 1;
        this.data.startWidth = 0;
    }

    /**
     * 处理双击事件
     * @param e 鼠标事件对象
     */
    public handleDoubleClick(e: MouseEvent) {
        const { offsetX: colHeaderX } = e;
        const { location, col, row } = this.data.colHeaders.getColHeaderCell(Math.round(colHeaderX / this.data.zoom));
        // 只有**左键双击**才自动调宽（单击、右键、按住拖动都不触发）
        if (e.button !== 0 || e.detail !== 2
            || e.target !== this.canvas
            || location !== MouseLocation.IN_COL_HEADER_SPLIT
        ) return;
        // 多列选区且双击的分界线属于选中列 → 对选中的全部列自动调宽
        const { startCol, endCol } = this.getSelectedColsAndRows();
        const inMultiSelection = endCol > startCol && col >= startCol && col <= endCol;
        const fitStart = inMultiSelection ? startCol : col;
        const fitEnd = inMultiSelection ? endCol : col;
        for (let c = fitStart; c <= fitEnd; c++) {
            this.autoAdjustColumnWidth(c);
        }
        // 列宽变化后补齐右侧列，保证网格填满视口
        this.data.fillColsToViewport(fitStart, fitEnd);
    }

    /**
     * 自动调整列宽以适应内容宽度
     *
     * **空列不做任何操作**：没有任何非空文本的列若也按 padding 宽度收窄，
     * 双击一下整列就被压到最小宽度，属于破坏性操作。
     *
     * **合并单元格不参与测量**：合并块（如 B22:D24）的文本宽度是整块共用的，
     * 若按整段文本给单列算宽，会把 B、C、D 同时撑到整块宽度（Excel 的自动调宽
     * 同样忽略合并单元格）。
     * @param col 列索引
     */
    private autoAdjustColumnWidth(col: number): void {
        if (col < 1 || col > this.data.colHeaders.length) return;
        
        // 获取该列所有**有文本且不在合并区域内**的单元格
        const cellsInColumn = this.data.values.filter(v => {
            const { col: cellCol, row: cellRow } = this.data.getCellColAndRow(v.cell);
            if (cellCol !== col) return false;
            if (this.data.findMergeRect(cellCol, cellRow) !== null) return false;
            return (this.data.getCellByName(v.cell)?.getText() ?? '') !== '';
        });
        if (cellsInColumn.length === 0) return;  // 空列不调整
        // 计算最大文本宽度
        let minWidth = 20;
        for (const cell of cellsInColumn) {
            const { charsWidth } = this.measureCellText(cell.cell);
            minWidth = Math.max(minWidth, charsWidth);
        }
        
        // 调整列宽
        this.data.adjustColumnWidth(col, Math.round(minWidth / this.data.zoom));
    }
}