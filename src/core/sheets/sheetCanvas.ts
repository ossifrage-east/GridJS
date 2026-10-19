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
import { Cell } from "../dataArchitecture/cell";
import { LineText } from "../dataArchitecture/cellContent";
import { Filter } from "../assistant/filter";
import { layoutVisibleImages, type ImageLayoutGrid } from "./imageLayout";
import { SheetImageLayer } from "./sheetImageLayer";

/**
 * 把选区矩形裁剪到视口范围（**纯函数**，便于单测）
 *
 * 选区可以远大于屏幕（全选、Ctrl+Shift+↓ 拖到表底、拖选几万行），此时直接用
 * 选区尺寸建离屏画布会超出浏览器画布上限（Chrome 单边 65535px、总像素 2^28），
 * 超限后画布静默变空白 —— 这就是「选区一大，选中框的半透明背景整片消失」的原因。
 * 裁剪后背板的尺寸恒不超过视口，既不会超限，也不必为屏外区域白填像素。
 *
 * @param rect - 选区矩形（工作表坐标系，相对视口原点）
 * @param view - 视口宽高
 * @returns 裁剪后的矩形；与视口无交集时返回 null
 */
export function clipRectToViewport(
    rect: { left: number; top: number; width: number; height: number },
    view: { width: number; height: number }
): { left: number; top: number; width: number; height: number } | null {
    const left = Math.max(rect.left, 0);
    const top = Math.max(rect.top, 0);
    const right = Math.min(rect.left + rect.width, view.width);
    const bottom = Math.min(rect.top + rect.height, view.height);
    if (right <= left || bottom <= top) return null;
    return { left, top, width: right - left, height: bottom - top };
}

export class SheetCanvas extends Canvas {
    /** 筛选按钮实例：单元格名 → Filter（每帧按名字定位，避免数组 find 的 O(F²)） */
    private filters: Map<string, Filter> = new Map();
    /**
     * 图片绘制层：负责位图异步解码缓存与绘制
     *
     * 图片解码是异步的，首次绘制那一帧拿不到位图，解码完成后需要再画一帧
     * （见 {@link scheduleImageRedraw}）。
     */
    private readonly imageLayer: SheetImageLayer;
    /** 是否已排期一次「图片解码完成」的重绘（合并同一批解码的多次回调） */
    private imageRedrawScheduled: boolean = false;

    /**
     * 单元格文本位图缓存：缓存键 → 已光栅化的单元格位图
     *
     * 滚动只改变格的位置，不改变内容与尺寸，因此视口内文本绝大多数可直接
     * drawImage 命中，省掉每帧的逐字符排版与绘制。条目超限时整体清空
     * （LRU 在这里收益不高，清空策略与基类离屏池一致）。
     */
    private readonly _textBitmapCache: Map<string, OffscreenCanvas> = new Map();
    /** 文本位图缓存的条目上限 */
    private static readonly TEXT_BITMAP_CACHE_LIMIT = 1500;
    constructor(options: CanvasOptions) {
        super(options);
        this.parentElement = options.parentElement;
        this.data = options.data;
        this.menu = options.menu;
        this.imageLayer = new SheetImageLayer({ requestRedraw: () => this.scheduleImageRedraw() });
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
        // 视口内的单元格集合 + 可见行列范围：一帧只算一次，供下面几遍绘制共用。
        // （历史实现让每遍绘制各自调用 getVisibleCells()，一帧算三遍，
        //   每遍都要新建上千个单元格名字符串并遍历合并区域。）
        const cells = this.data.getVisibleCells();
        const { startCol, endCol } = this.data.getVisibleCols();
        const { startRow, endRow } = this.data.getVisibleRows();
        this.drawCellBgColor(cells);
        this.drawCellBorder(cells);
        // 合并区遮挡必须在**边框之后、文字之前**：
        //  - 之后：合并区内部各格的边框此刻刚被 drawCellBorder 画上去，
        //    只有压在它上面才能被离屏画布整块盖掉（只留最外圈）；
        //  - 之前：合并区的文字要画在遮挡层之上，否则会被一起盖没。
        this.drawMergedCellCover(startCol, endCol, startRow, endRow);
        this.drawCellsText(cells);
        // 图片是浮动对象：画在网格内容之上、选区框之下（否则选中框会被图盖住看不见）
        this.drawImages();
        if (this.data.startLocation === MouseLocation.IN_ANCHOR) {
            this.drawAnchorSelcetedRect();
        }
        this.drawSelectedRect();
        this.ctx.restore();
        this.setFilters();
    }

    /**
     * 绘制当前工作表内嵌的图片
     *
     * 【性能约定，勿回退】先把矩形按**视口裁剪**（{@link layoutVisibleImages}），
     * 再交给绘制层：代价只与「屏上图片数」相关，与工作簿图片总数无关。
     * 无图片时（绝大多数表格）这里在第一行就返回，不产生任何分配。
     */
    private drawImages(): void {
        const images = this.data.images;
        if (images.length === 0) return;

        const { colHeaders, rowHeaders, visibleView } = this.data;
        // 表头几何按需查询：隐藏列在 recalcColPositions 后已被压成零视觉宽度，
        // 因此这里的 left/top 累计值天然避开了隐藏行列
        const grid: ImageLayoutGrid = {
            colLeft: (col) => colHeaders.getAt(col)?.left ?? 0,
            colWidth: (col) => colHeaders.getAt(col)?.width ?? DEFAULT_CELL_WIDTH,
            rowTop: (row) => rowHeaders.getAt(row)?.top ?? 0,
            rowHeight: (row) => rowHeaders.getAt(row)?.height ?? DEFAULT_CELL_HEIGHT,
        };
        const placed = layoutVisibleImages(images, grid, {
            left: colHeaders.offsetWidth,
            top: rowHeaders.offsetHeight,
            width: visibleView.sheetWidth,
            height: visibleView.sheetHeight,
        });
        this.imageLayer.draw(this.ctx, placed, this.data.zoom, window.devicePixelRatio);
    }

    /**
     * 排期一次「图片解码完成」的重绘（同一帧内的多次解码回调合并为一次）
     *
     * 解码完成时图片还没被画出来，不重绘就要等到用户滚动/点击才突然出现。
     * 但一批图会先后解码完（每次回调都重绘等于白画十几帧），故用一个标记合并。
     * `requestAnimationFrame` 不可用时回落 `setTimeout`（Node 测试环境）。
     * @private
     */
    private scheduleImageRedraw(): void {
        if (this.imageRedrawScheduled) return;
        this.imageRedrawScheduled = true;
        const repaint = (): void => {
            this.imageRedrawScheduled = false;
            void this.draw();
        };
        if (typeof requestAnimationFrame === 'function') requestAnimationFrame(repaint);
        else setTimeout(repaint, 16);
    }

    /**
     * 设置筛选组件实例, 如果存在筛选组件实例，先移除旧的实例
     *
     * 【性能约定，勿回退为「每帧先问一遍筛选单元格」】
     * 本方法每帧执行，而「本帧有没有筛选按钮」必须 O(1) 回答：
     * `getFilterCells()` 内部会 `_ensureCellIndex()`，一旦世代失效
     * （删除行/列、排序、改名、装载…）就**先重建整份单元格索引**
     * （遍历全部单元格 + 建 Map + 建合并倒排）才回答「其实一个都没有」。
     * 于是「表里根本没有筛选按钮」的绝大多数帧，也在为筛选白付一次索引重建。
     * {@link DataCollection.hasFilterCells} 按世代缓存了「上次数出 0 个」这一结论，
     * 同一世代内重复询问零成本。
     *
     * 【性能约定，勿回退为数组 + find】
     * 按钮数量为 F 时数组写法是 O(F²)（每个按钮都要 `find` 一遍现有实例与缓存列表）；
     * 改为按单元格名索引后为 O(F)。
     * 另外 {@link Filter.setPosition} 内部已做「位置未变则零 DOM 写入」，
     * 因此稳态帧不会产生强制同步布局。
     */
    private setFilters() {
        // O(1) 提前退出：本世代确定没有任何筛选按钮时，连索引都不必重建。
        // 首次调用必然为 true（_filterCellsEpoch 初始为 -1），不会漏掉已有按钮。
        if (!this.data.hasFilterCells()) {
            if (this.filters.size > 0) {
                this.filters.forEach(filter => filter.destroy());
                this.filters.clear();
            }
            return;
        }
        // 筛选单元格由 DataCollection 随单元格索引一并缓存，避免每帧对整表 filter
        const dataFilters = this.data.getFilterCells();
        if (dataFilters.length === 0) {  // 如果data中没有筛选数据，则销毁组件实例直接返回
            this.filters.forEach(filter => filter.destroy());
            this.filters.clear();
            return;
        }
        // 反向清理：cell 已被删除（如删除行/列）的 filter 实例需销毁，
        // 否则按钮 DOM 拖留、停留在旧位置显示。
        // 用 Set 做成员判定（O(F)），不要再写「对每个实例再遍历一遍筛选数组」
        // 的数组套数组写法 —— 那是 O(F²)，F 到几十上百时每帧都会白烧 CPU。
        if (this.filters.size > 0) {
            const live = new Set<string>();
            for (let i = 0; i < dataFilters.length; i++) live.add(dataFilters[i].cell);
            const stale: string[] = [];
            this.filters.forEach((filter, cellName) => {
                if (!live.has(cellName)) {
                    filter.destroy();
                    stale.push(cellName);
                }
            });
            for (const cellName of stale) this.filters.delete(cellName);
        }

        for (const cell of dataFilters) {
            const filter = this.filters.get(cell.cell);
            if (filter) { // 如果存在筛选组件实例，直接设置位置
                filter.setPosition(cell);
                continue;
            }
            const _filter = new Filter({parentElement: this.parentElement, data: this.data, menu: this.menu});
            _filter.setPosition(cell);
            this.filters.set(cell.cell, _filter);
        }
    }
    
    /**
     * 绘制网格线
     */
    private drawGridLines() {
        if (this.data.showGridLines === false) return; // 如果不显示网格线，直接返回
        this.ctx.save();
        this.ctx.strokeStyle = GRID_LINE_COLOR;

        // 绘制网格的垂直线条（只遍历视口内可见列，不再按列号跳过隐藏列）
        // 注意 getVisibleColNumbers 返回的是 1-based 列号，这里转回 0-based 下标取列头
        const colNumbers = this.data.getVisibleColNumbers();
        for (let ci = 0; ci < colNumbers.length; ci++) {
            const c = colNumbers[ci] - 1; // 1-based 转 0-based
            this.setLineWidth(1);  // 使用 CSS 像素，内部自动转换为物理像素
            const x = (this.data.colHeaders.getAt(c).left - this.data.colHeaders.offsetWidth) * this.data.zoom;
            this.drawLine(x, 0, x, this.data.visibleView.sheetHeight);
        }

        // 绘制网格的水平线条（只遍历视口内可见行，不再按行号跳过隐藏行）
        // 注意 getVisibleRowNumbers 返回的是 1-based 行号，这里转回 0-based 下标取行头
        const rowNumbers = this.data.getVisibleRowNumbers();
        for (let ri = 0; ri < rowNumbers.length; ri++) {
            const r = rowNumbers[ri] - 1; // 1-based 转 0-based
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
        // 合并展开：命中合并区域时取左上角起点。
        // 走 DataCollection 的合并倒排索引（O(该列合并块数)）—— 历史实现对本函数
        // 的**每次**调用都要遍历整张表的全部合并块并逐个 split(':') 解析行列号，
        // 而本函数又被「每个可见单元格 × 三遍绘制」调用，合并块上千时单帧上千万次比较。
        let startCol = col;
        let startRow = row;
        const mergeRect = this.data.findMergeRect(col, row);
        if (mergeRect !== null) {
            startCol = mergeRect.startCol;
            startRow = mergeRect.startRow;
        }

        // 读取合并跨度
        const cellName = this.data.getCellName(startCol, startRow);
        const cell = this.data.getCellByName(cellName);
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
     * @param {Cell[]} cells - 视口内的单元格（由 draw() 统一计算，避免每遍遍历重复取数）
     */
    private drawCellBgColor(cells: Cell[]) {
        // 视口裁剪：只遍历当前可见的单元格，复杂度 O(可见单元格数)，与表格总规模无关
        for (let i = 0; i < cells.length; i++) {
            const cell = cells[i];
            // 无背景色直接跳过：矩形计算（隐藏行列/合并展开）比一次判空贵得多，
            // 而大表里绝大多数单元格本就没有填充色，先判空能省掉大部分矩形计算
            if (!cell.backgroundColor) continue;
            const { col, row } = this.data.getCellColAndRow(cell.cell);
            // 隐藏行列不绘制；合并单元格只绘制可见部分
            const rect = this.getVisibleCellRect(col, row);
            if (!rect) continue;
            const cellLeft = (rect.left + 1) * this.data.zoom;
            const cellTop = (rect.top + 1) * this.data.zoom;
            const cellWidth = (rect.width - 1) * this.data.zoom;
            const cellHeight = (rect.height - 1) * this.data.zoom;
            this.ctx.save();
            this.ctx.fillStyle = cell.backgroundColor;
            this.fillRect(cellLeft, cellTop, cellWidth, cellHeight);
            this.ctx.restore();        
        }
    }
    
    /**
     * 绘制单元格边框
     * @param {Cell[]} cells - 视口内的单元格（由 draw() 统一计算）
     */
    private drawCellBorder(cells: Cell[]) {
        // 视口裁剪：只遍历当前可见的单元格
        for (let i = 0; i < cells.length; i++) {
            const cell = cells[i];
            // 四边都无边框的单元格直接跳过（网格线由 drawGridLines 负责，
            // 这里画不出任何东西，却会白做一次矩形计算与一次 ctx.save/restore）
            if (!cell.borderTopWidth && !cell.borderLeftWidth && !cell.borderRightWidth && !cell.borderBottomWidth) continue;
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
        const zoom = this.data.zoom;
        const dpr = window.devicePixelRatio;

        // 【视口裁剪，勿回退为「按整个选区建离屏画布」】
        // 选区常常远大于屏幕（全选、拖到第几万行、Ctrl+Shift+↓ 到底），此时
        // `w × h` 会超出浏览器画布上限（Chrome 单边 65535px、总像素 2^28），
        // 超限后 OffscreenCanvas 静默变成空白 —— 表现就是「选区一大，选中框的
        // 半透明背景整片消失，只剩一个描边框」。同时为屏外区域做几百万像素的
        // 填充也是纯浪费（每帧都做）。故只给「选区 ∩ 视口」建画布。
        const clip = clipRectToViewport(
            { left, top, width, height },
            { width: this.data.visibleView.sheetWidth, height: this.data.visibleView.sheetHeight }
        );

        this.ctx.save();
        if (clip !== null) {
            // 绘制选中区域背景及activedCell透明背景
            const { offscreenCanvas, offscreenCtx } = this.offscreen(clip.width * zoom * dpr, clip.height * zoom * dpr);
            offscreenCtx.clearRect(0, 0,
                offscreenCanvas.width,
                offscreenCanvas.height);
            // 绘制选中区域背景（从 offscreenCanvas 的原点开始绘制）
            offscreenCtx.fillStyle = DEFAULT_SELECTED_BG_COLOR;
            offscreenCtx.fillRect(0, 0,
                offscreenCanvas.width,
                offscreenCanvas.height);

            // 绘制activedCell透明背景（相对裁剪区左上角的坐标）
            offscreenCtx.clearRect(
                (activedLeft - clip.left) * zoom * dpr,
                (activedTop - clip.top) * zoom * dpr,
                activedWidth * zoom * dpr,
                activedHeight * zoom * dpr);

            // 绘制选中区域背景（贴到裁剪后的位置）
            this.ctx.drawImage(offscreenCanvas, clip.left * zoom * dpr, clip.top * zoom * dpr);
        }
        // 选区完全滚出视口时不画背景，但描边仍按完整选区画（Canvas 自行裁剪，
        // 超出部分不显示，也不必为它单独算裁剪）

        this.ctx.strokeStyle = DEFAULT_SELECTED_BORDER_COLOR;
        this.setLineWidth(DEFAULT_SELECTED_BORDER_WIDTH);
        this.strokeRect(left * zoom,
            top * zoom,
            width * zoom,
            height * zoom);
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
     *
     * 【性能约定，勿回退】
     * 一帧内每个可见单元格都要走一遍，因此这里刻意做了几件事：
     * 1. 空文本单元格直接跳过（大表里空格子占多数）；
     * 2. **单元格文本位图缓存**（{@link _textBitmapCache}）：滚动时单元格的
     *    内容、字号、尺寸都不变，只是位置变化。命中缓存时直接 drawImage，
     *    跳过最贵的 measureCellText（逐字符度量 + 换行计算）与逐字符 fillText；
     * 3. 未命中时走 {@link rasterizeCellText}，同样保持：
     *    - 不逐字符 save/restore；
     *    - font/fillStyle 仅在样式变化时赋值。
     *
     * 缓存键包含物理像素尺寸、zoom、dpr、对齐/换行/字距行距/数字格式以及
     * 每个字符的内容与字体签名，因此任何一项变化都会自然产生新键，无需手动失效。
     * @param {Cell[]} cells - 视口内的单元格（由 draw() 统一计算）
     */
    private drawCellsText(cells: Cell[]) {
        this.ctx.save();
        const dpr = window.devicePixelRatio;
        // 视口裁剪：只遍历当前可见的单元格
        for (let i=0; i<cells.length; i++) {
            const cellObj = cells[i];
            const cell = cellObj.cell;
            // 无文本的单元格不产生任何绘制指令，跳过（字符经 chars 数组承载；
            // 只有样式的单元格 chars 为空/未定义，同样无需走一遍离屏合成）
            const chars = cellObj.chars;
            if (!chars || chars.length === 0) continue;
            const { col, row } = this.data.getCellColAndRow(cell);

            // 计算考虑隐藏行列后的可见矩形；完全隐藏时跳过绘制
            const rect = this.getVisibleCellRect(col, row);
            if (!rect) continue;
            const cellWidth = rect.width;
            const cellHeight = rect.height;
            const cellLeft = rect.left;
            const cellTop = rect.top;

            const bitmapW = Math.round(cellWidth * this.data.zoom * dpr);
            const bitmapH = Math.round(cellHeight * this.data.zoom * dpr);
            const cacheKey = this.buildTextBitmapKey(cellObj, bitmapW, bitmapH);
            let bitmap = this._textBitmapCache.get(cacheKey);
            if (!bitmap) {
                // 未命中：测量布局并光栅化到一张专属位图，随后缓存供滚动复用
                const { lines } = this.measureCellText(cell);
                bitmap = new OffscreenCanvas(bitmapW, bitmapH);
                const bctx = bitmap.getContext('2d');
                if (!bctx) continue;
                this.rasterizeCellText(cell, lines, bctx);
                if (this._textBitmapCache.size >= SheetCanvas.TEXT_BITMAP_CACHE_LIMIT) {
                    this._textBitmapCache.clear();
                }
                this._textBitmapCache.set(cacheKey, bitmap);
            }

            // 将缓存位图合并到主 canvas（坐标/尺寸口径与原先逐字符绘制完全一致）
            const x = Math.round(cellLeft * this.data.zoom);
            const y = Math.round(cellTop * this.data.zoom);
            this.ctx.drawImage(bitmap,
                Math.round(x * dpr), Math.round(y * dpr),
                bitmapW, bitmapH);
        }

        this.ctx.restore();
    }

    /**
     * 构造单元格文本位图的缓存键
     *
     * 键覆盖所有能影响光栅化结果的输入：位图物理尺寸（含列宽行高/zoom/dpr）、
     * 单元格级排版属性（换行/对齐/字距/行距/数字格式/单元格级字体兜底）、
     * 以及逐字符的内容与 7 个字体字段。任一项变化都会得到不同的键。
     *
     * @param {Cell} cellObj - 单元格
     * @param {number} bitmapW - 位图物理宽
     * @param {number} bitmapH - 位图物理高
     * @returns {string} 缓存键
     * @private
     */
    private buildTextBitmapKey(cellObj: Cell, bitmapW: number, bitmapH: number): string {
        const dpr = window.devicePixelRatio;
        const z = Math.round(this.data.zoom * 1000);
        const chars = cellObj.chars ?? [];
        // 逐字符签名用单字符分隔符拼接；短字段名不参与以控制键长度
        let sig = '';
        for (let i = 0; i < chars.length; i++) {
            const c = chars[i];
            sig += c.char
                + '\u0001' + (c.fontFamily ?? '')
                + ',' + (c.fontSize ?? '')
                + ',' + (c.fontWeight ? 1 : 0)
                + ',' + (c.fontStyle ? 1 : 0)
                + ',' + (c.fontColor ?? '')
                + ',' + (c.underline ? 1 : 0)
                + ',' + (c.strikethrough ? 1 : 0)
                + '\u0002';
        }
        return `${bitmapW}x${bitmapH}|${z}|${dpr}|${cellObj.wrap ? 1 : 0}`
            + `|${cellObj.letterSpacing ?? 0}|${cellObj.lineSpacing ?? 0}`
            + `|${cellObj.numberFormat ?? ''}`
            + `|${cellObj.textAlign ?? ''}|${cellObj.alignItems ?? ''}`
            + `|${cellObj.fontFamily ?? ''}|${cellObj.fontSize ?? ''}|${sig}`;
    }

    /**
     * 把一个单元格的文本行光栅化到指定（物理像素）位图上下文
     *
     * 从原 drawCellsText 抽出的纯绘制逻辑，坐标系与主画布物理像素一致。
     * 保持两条性能约定：不逐字符 save/restore；font/fillStyle 仅在变化时赋值。
     *
     * @param {string} cell - 单元格名
     * @param {LineText[]} lines - measureCellText 产出的行布局
     * @param {OffscreenCanvasRenderingContext2D} ctx - 目标位图上下文（已按物理像素建图）
     * @private
     */
    private rasterizeCellText(
        cell: string,
        lines: LineText[],
        ctx: OffscreenCanvasRenderingContext2D
    ): void {
        const cellObj = this.data.getCellByName(cell);
        const letterSpacing = cellObj?.letterSpacing || 0;
        const lineSpacing = cellObj?.lineSpacing || 0;
        ctx.textAlign = 'left';
        ctx.textBaseline = 'bottom';
        // 连续相同样式的字符共享同一次 ctx 状态赋值（见 drawCellsText 注释）
        let lastFontStyle: string | null = null;
        let lastFontWeight = '';
        let lastFontSize = -1;
        let lastFontFamily = '';
        let lastFillColor = '';
        let k = 0; // 字符全局索引（与 calculateCharGeometry 的 charIndex 对齐）
        for (let li = 0; li < lines.length; li++) {
            const line = lines[li];
            if (line.text.length === 0) continue;
            let currentX = line.x;
            const currentY = line.y + line.lineHeight - lineSpacing;
            for (let j = 0; j < line.text.length; j++) {
                const ch = line.text[j] === '\n' ? '' : line.text[j];
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

                if (fontColor !== lastFillColor) {
                    ctx.fillStyle = fontColor;
                    lastFillColor = fontColor;
                }
                if (fontStyle !== lastFontStyle || fontWeight !== lastFontWeight
                    || fontSize !== lastFontSize || fontFamily !== lastFontFamily) {
                    const fontParts: string[] = [];
                    if (fontStyle) fontParts.push(fontStyle);
                    if (fontWeight) fontParts.push(fontWeight);
                    fontParts.push(`${Math.round(fontSize * window.devicePixelRatio)}px`);
                    fontParts.push(fontFamily);
                    ctx.font = fontParts.join(' ');
                    lastFontStyle = fontStyle;
                    lastFontWeight = fontWeight;
                    lastFontSize = fontSize;
                    lastFontFamily = fontFamily;
                }
                ctx.fillText(ch, Math.round(currentX * window.devicePixelRatio), Math.round(currentY * window.devicePixelRatio));

                // 绘制下划线 / 删除线：临时改描边参数，画完就地复位（无需 save/restore）
                if (underline || strikethrough) {
                    ctx.strokeStyle = fontColor;
                    ctx.lineWidth = Math.max(1, Math.round(fontSize / 10 * window.devicePixelRatio));
                    if (underline) {
                        const underlineY = Math.round(currentY * window.devicePixelRatio - 1);
                        ctx.beginPath();
                        ctx.moveTo(Math.round(currentX * window.devicePixelRatio), underlineY);
                        ctx.lineTo(Math.round((currentX + charWidth) * window.devicePixelRatio), underlineY);
                        ctx.stroke();
                    }
                    if (strikethrough) {
                        const strikethroughY = Math.round((currentY - fontSize * 0.48) * window.devicePixelRatio);
                        ctx.beginPath();
                        ctx.moveTo(Math.round(currentX * window.devicePixelRatio), strikethroughY);
                        ctx.lineTo(Math.round((currentX + charWidth) * window.devicePixelRatio), strikethroughY);
                        ctx.stroke();
                    }
                }

                k += 1;
                currentX += charWidth + (j < line.text.length - 1 ? letterSpacing : 0);
            }
        }
    }

    /**
     * 用离屏画布「盖住」每一个可见合并区域，只留最外圈的边框
     *
     * 合并单元格在视觉上就是一个大格，但数据层**并没有**删掉内部那些带边框/底色的
     * 格子（删了就再也还原不回来，正是「取消合并后线断续」的根因）。所以这里在渲染层
     * 做遮挡：把合并区整块不透明地铺上锚点底色，盖掉
     *   ① 穿过区域内部的网格线（drawGridLines 画的全表竖线/横线）、
     *   ② 内部各格自己的边框（drawCellBorder 刚画上去的）、
     *   ③ 内部各格与锚点不一致的底色，
     * 再沿区域**外轮廓**重画那四条边（线宽取轮廓上各格的边框，
     * 见 {@link DataCollection.getMergedOutlineBorders}）。
     *
     * 取消合并时因此什么都不用下发：清掉 colspan/rowspan 后，各格自己的边框原样还在，
     * 主屏的 drawCellBorder + drawGridLines 自然把竖线横线连续地画出来。
     *
     * 隐藏行列不绘制；合并单元格跨越隐藏列/行时只绘制可见部分。
     *
     * 【性能约定，勿回退】只处理与视口相交的合并区域（走列倒排索引），
     * 不要再对「全表所有合并块」逐个调用矩形计算 —— 那会退化成 O(合并数²)，
     * 实测 8000 处合并时单帧 3.8 秒。离屏画布按尺寸复用（{@link Canvas.offscreen}）。
     *
     * 【视口裁剪，勿回退为「按整个合并区建离屏画布」】合并区常常远大于屏幕，
     * 超限后 OffscreenCanvas 会静默变空白（Chrome 单边 65535px、总像素 2^28）。
     * @param {number} startCol - 视口起始列号（1-based）
     * @param {number} endCol - 视口结束列号（1-based）
     * @param {number} startRow - 视口起始行号（1-based）
     * @param {number} endRow - 视口结束行号（1-based）
     */
    public drawMergedCellCover(startCol: number, endCol: number, startRow: number, endRow: number) {
        const MergedCells = this.data.getVisibleMergeRects(startCol, endCol, startRow, endRow);
        if (MergedCells.length === 0) return;
        const zoom = this.data.zoom;
        const dpr = window.devicePixelRatio;
        const viewport = { width: this.data.visibleView.sheetWidth, height: this.data.visibleView.sheetHeight };
        for (const mergedCell of MergedCells) {
            // 隐藏行列不绘制；跨隐藏列/行时只取可见部分
            const rect = this.getVisibleCellRect(mergedCell.startCol, mergedCell.startRow);
            if (!rect) continue;
            const clip = clipRectToViewport(
                { left: rect.left, top: rect.top, width: rect.width, height: rect.height },
                viewport
            );
            if (!clip) continue;

            // ── 遮挡层：不透明整块铺底色，盖掉内部网格线与内部各格的边框/底色 ──
            // 内缩 1 CSS px：区域**边界**上的线属于「外面」（相邻格子的边框、
            // 以及下面要重画的外轮廓），不该被盖住；而内部的网格线离边界至少一整列/行，
            // 内缩 1px 完全不影响遮挡效果。
            const coverLeft = clip.left + 1;
            const coverTop = clip.top + 1;
            const coverWidth = clip.width - 1;
            const coverHeight = clip.height - 1;
            if (coverWidth > 0 && coverHeight > 0) {
                const anchor = this.data.getCellByName(mergedCell.anchor);
                const backgroundColor = anchor?.backgroundColor || DEFAULT_CELL_BG_COLOR;
                const { offscreenCanvas, offscreenCtx } =
                    this.offscreen(coverWidth * zoom * dpr, coverHeight * zoom * dpr);
                offscreenCtx.clearRect(0, 0, offscreenCanvas.width, offscreenCanvas.height);
                offscreenCtx.fillStyle = backgroundColor;
                offscreenCtx.fillRect(0, 0, offscreenCanvas.width, offscreenCanvas.height);
                this.ctx.drawImage(offscreenCanvas,
                    Math.round(coverLeft * zoom * dpr), Math.round(coverTop * zoom * dpr));
            }

            // ── 外轮廓：只画区域最外圈那四条边 ──
            const outline = this.data.getMergedOutlineBorders(
                mergedCell.startCol, mergedCell.endCol, mergedCell.startRow, mergedCell.endRow);
            if (!outline.top && !outline.bottom && !outline.left && !outline.right) continue;
            const left = rect.left * zoom;
            const top = rect.top * zoom;
            const width = rect.width * zoom;
            const height = rect.height * zoom;
            const anchor = this.data.getCellByName(mergedCell.anchor);
            this.ctx.save();
            this.ctx.strokeStyle = (anchor && anchor.borderColor) || outline.color || 'black';
            if (outline.top) this.drawLine(left, top, left + width, top, outline.top);
            if (outline.bottom) this.drawLine(left, top + height, left + width, top + height, outline.bottom);
            if (outline.left) this.drawLine(left, top, left, top + height, outline.left);
            if (outline.right) this.drawLine(left + width, top, left + width, top + height, outline.right);
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
            // 命中合并区域 → 选区取整块（**几何判定**，见 DataCollection.getMergedRangeAt）
            //
            // 【勿回退为 `mergedCells.find(v => v.includes(name))`】那是**字符串包含**：
            // 点 B2 会命中 'B22:D24'（含子串 'B2'）、点 D2 同样命中（'D24' 含 'D2'），
            // 于是选中框跑到二十行开外毫不相关的合并块上（test_sample.xls 实测）。
            const mergedRange = this.data.getMergedRangeAt(col, row);
            if (mergedRange !== null) {
                // 活动格必须落在**合并锚点**上，而不是被合并覆盖的那个格子：
                // 覆盖格的 colspan/rowspan 归零，编辑光标与填充柄无处安放；
                // 且 getActivedRect 会把透明洞挖成整个合并块那么大。
                this.data.activedCell = mergedRange.split(':')[0];
                this.data.selection = mergedRange;
                return;
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

        // 单元格拖选越过内容区左 / 上边沿（光标进入行头 / 列头区域）时，接管为
        // 向左 / 上自动滚动并把选区延伸到被滚出的上一列 / 上一行，直至第 1 列 / 行。
        // 【勿并入下面的越界守卫】那条守卫是给「行 / 列头拖选」用的：往左 / 上拖拽
        // 若也直接忽略，选区就会卡在首个可见列 / 行（往右 / 下靠 getSheetCell 钳到
        // 最后可见行列 + scrollCellIntoView 的 `left + width > 可见宽` 分支能自动延伸，
        // 左 / 上缺少等价分支）。行 / 列头拖选（startLocation 在表头）保持原样。
        if (this.data.isMouseDown
            && this.data.startLocation === MouseLocation.IN_CELL
            && (sheetX < 0 || sheetY < 0)) {
            this.handleSelectionEdgeScroll(sheetX, sheetY);
            // 复位命中缓存：回到内容区后同一格仍需重新触发（overCell 相同会短路）
            this.data.overCell = { col: 0, row: 0 };
            return;
        }

        // 【坐标越界守卫，勿删】mousemove 绑定在 document 上（sheets.ts），本方法会
        // 收到落在行头/列头区域的事件（x<0 / y<0）。此时 getSheetCell 走 x<0 / y<0
        // 边界分支返回 IN_ROW_HEADER / IN_COL_HEADER，且其「合并覆盖」会把行/列号
        // 吸附到合并锚点（如 A1:C54 覆盖下第 4 行 A4 → 锚点 A1），再经 setSelection
        // 的行/列头分支把选区改写成 A1:AO{起始行} / A{起始列}:AO1 —— 行/列头拖选
        // 时选区被重置到合并锚点（用户反馈：不能跳到合并单元格的位置作为选择依据）。
        // 行头/列头拖选由各自画布的 handleMouseMove 负责（行列号不经合并覆盖）；
        // 这里越界即忽略，保留向下/右越界时 getSheetCell 钳到最后可见行列的
        // 「拖选出视口自动延伸」行为。
        if (sheetX < 0 || sheetY < 0) return;
        // 行/列头拖选进行中（鼠标从行/列头拖入内容区，e.target 已变成本画布，
        // 行/列头画布的 mousemove 因 e.target 守卫不再更新选区，只能由这里接管）：
        // 必须用**纯几何**行列号，跳过 getSheetCell 的锚点/合并吸附 —— 首列纵向
        // 合并（A1:A3）会把经过行吸附到锚点行，首行横向合并同理，导致整行/整列
        // 选区端点被合并块错误锚定（用户需求：仅选对应行/列，忽略合并影响）。
        // setSelection 的行/列头分支本就传 expandMerges=false，两端语义一致。
        const headerDragging = this.data.isMouseDown
            && (this.data.startLocation === MouseLocation.IN_ROW_HEADER
                || this.data.startLocation === MouseLocation.IN_COL_HEADER);
        const { location, col, row } = this.data.getSheetCell(sheetX, sheetY, headerDragging ? { rawGeometry: true } : undefined);
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
     * 单元格拖选贴近内容区左 / 上边沿时的反向自动滚动
     *
     * 向右 / 下拖拽时 getSheetCell 会把越界坐标钳到「最后可见列 / 行」，再由
     * scrollCellIntoView 的 `left + width > 可见宽` 分支逐列 / 逐行把更后的列 / 行
     * 纳入选区；向左 / 上时首可见列 / 行的左 / 上边缘恒为 0、不满足该分支，于是
     * 选区卡在首可见列 / 行。这里补上镜像：滚动目标取「首可见列 / 行 - 1」，借
     * scrollCellIntoView 的 `left < 0` 分支把它对齐到视口左 / 上沿，逐次向前推进；
     * 若左侧 / 上方已无更多可见列 / 行则直接钳到第 1 列 / 行。全程按单元格语义，
     * 不转成整行 / 整列，故 location 固定传 {@link MouseLocation.IN_CELL}。
     *
     * 【勿改用 selection 字符串反推端点】选区经 setSelectedRange 归一化后，`':'`
     * 之后恒为**右下角**（max），向左 / 上拖拽时拿到的不是当前活动端点；且单格选区
     * （如首帧就拖出行头）根本没有 `':'`。这里一律用**纯几何**命中取当前光标落点，
     * 再对越界的那一侧做覆盖，既不怕单体格也不怕对角拖拽。
     * @param sheetX 内容区坐标X（<0 表示光标已越过左边沿进入行头）
     * @param sheetY 内容区坐标Y（<0 表示光标已越过上边沿进入列头）
     * @private
     */
    private handleSelectionEdgeScroll(sheetX: number, sheetY: number): void {
        // 纯几何落点：sheetX<0 时 getSheetCell 会把列钳到 1、sheetY<0 时把行钳到 1，
        // 未越界的一侧（如往左拖入行头时 sheetY>=0）则给出光标所在的实际行 / 列，
        // 正好是当前拖拽应保持的另一维端点。
        const geo = this.data.getSheetCell(sheetX, sheetY, { rawGeometry: true });
        let col = geo.col;
        let row = geo.row;

        if (sheetX < 0) {
            const prevCol = this.previousVisibleCol(this.data.getVisibleCols().startCol);
            col = prevCol > 0 ? prevCol : 1;   // 左侧已无更多可见列时钳到第 1 列
        }
        if (sheetY < 0) {
            const prevRow = this.previousVisibleRow(this.data.getVisibleRows().startRow);
            row = prevRow > 0 ? prevRow : 1;
        }

        this.data.setSelection(MouseLocation.IN_CELL, col, row);
        this.scrollCellIntoView(col, row);
        this.updateCursorStyle(MouseLocation.IN_CELL);
    }

    /**
     * 查找首个可见列之前最近的一个可见列（1-based）
     * @param firstShownCol - 首个可见列号（1-based，{@link DataCollection.getVisibleCols} 的 startCol）
     * @returns {number} 上一可见列号；其左侧已无可见列时返回 0
     * @private
     */
    private previousVisibleCol(firstShownCol: number): number {
        for (let c = firstShownCol - 2; c >= 0; c--) {
            if (!this.data.colHeaders.getAt(c).isHidden) return c + 1;
        }
        return 0;
    }

    /**
     * 查找首个可见行之前最近的一个可见行（1-based）
     * @param firstShownRow - 首个可见行号（1-based，{@link DataCollection.getVisibleRows} 的 startRow）
     * @returns {number} 上一可见行号；其上方已无可见行时返回 0
     * @private
     */
    private previousVisibleRow(firstShownRow: number): number {
        for (let r = firstShownRow - 2; r >= 0; r--) {
            if (!this.data.rowHeaders.getAt(r).isHidden) return r + 1;
        }
        return 0;
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
        if (this.data.brushMode) {
            // 格式刷落刷：影响面 = 当前目标选区，'selection' 档按选区捕获稀疏补丁，
            // 大表同样可撤销/重做（exeBrush 单格分支内部调 pasteFormatOnly，
            // 嵌套的记账捕获会自动折叠到本层，只产生一条记录）。
            this.data.runWithFullStateUndo('格式刷', () => this.data.exeBrush(), 'selection');
            // 包裹抑制了 exeBrush 内部 commitValues 的自动广播，补一次驱动重绘
            this.data.emit(DataEvents.VALUES_CHANGED, this.data.values);
        }
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
