/*  # 工作表数据层职责拆分（K3）
    # 从 dataCollection.ts 原样迁出，仅调整类继承关系与成员可见性，逻辑零改动
*/
import { Cell } from "../dataArchitecture/cell";
import { DataEvents, DEFAULT_CELL_HEIGHT, DEFAULT_CELL_WIDTH } from "../constant";
import { SheetImagesBase } from "./sheetImages";
import { SheetStoreBase } from "./sheetStore";

/**
 * 筛选 / 行列隐藏基类
 *
 * 筛选条件注册表、隐藏行列倒排索引、可见范围收尾（补足 / 收缩）与筛选按钮缓存。
 */
export abstract class SheetFilterBase extends SheetImagesBase {
    /* ---------- 仍定义在 DataCollection 中、供筛选逻辑调用的成员 ---------- */
    protected abstract _ensureCellIndex(): void;
    public abstract ensureWritableCell(name: string): Cell;
    public abstract getCellByName(name: string): Cell | undefined;
    public abstract getCellColAndRow(cell: string): { col: number, row: number };
    public abstract getCellName(col: number, row: number): string;
    public abstract getColName(num: number): string;
    public abstract getColNumber(char: string): number;
    public abstract getContentBounds(): { maxRow: number; maxCol: number };
    public abstract getColsWithContent(): Set<number>;
    public abstract getRowsWithText(): Set<number>;
    public abstract getEffectiveCellAt(row: number, col: number): Cell | undefined;
    public abstract runWithFullStateUndo(
        label: string,
        action: () => void,
        scope?: 'auto' | 'selection' | 'headers' | 'paste' | 'brush' | 'shiftLeft' | 'shiftUp' | 'filters',
        rect?: { startCol: number; endCol: number; startRow: number; endRow: number }
    ): void;
    /**
     * 重扫 `values` 收集带筛选标记的单元格（刷新 {@link _filterCells} 缓存）
     *
     * 只对已存在的单元格做一次线性扫描，不建 Map、不建合并倒排，
     * 因此代价远低于 {@link _rebuildCellIndex}。由 {@link _ensureCellIndex}
     * 在 {@link _filterCellsDirty} 置位时调用。
     * @private
     */
    protected _rebuildFilterCells(): void {
        const values = this._values;
        const filterCells: Cell[] = [];
        for (let i = 0; i < values.length; i++) {
            if (values[i].filter !== undefined) filterCells.push(values[i]);
        }
        this._filterCells = filterCells;
        this._filterCellsEpoch = this._valuesEpoch;
        this._filterCellsDirty = false;
    }

    /**
     * 获取带筛选标记的单元格列表（只读引用，请勿修改）
     * @returns {Cell[]} 带筛选标记的单元格
     */
    public getFilterCells(): Cell[] {
        this._ensureCellIndex();
        return this._filterCells;
    }

    /**
     * 「本世代是否存在筛选按钮」的廉价判定 —— 供每帧热路径提前退出（O(1)）
     *
     * `SheetCanvas.setFilters()` 每帧都要知道「这帧要不要处理筛选按钮」。
     * 历史实现先调 {@link getFilterCells}，而它会 `_ensureCellIndex()`：
     * 一旦世代失效（删除行/列、排序、改名、装载…）就**先重建整份单元格索引**
     * （遍历全部单元格 + 建 Map + 建合并倒排）才回答「其实一个筛选按钮都没有」。
     *
     * 这里用 {@link _filterCellsEpoch} 记录「上次数出 0 个按钮」的世代：
     * 同一世代内重复询问直接返回 false，不触发任何索引重建。
     * 世代失效即返回 true（保守：宁可让调用方去问一次，也不能漏掉新加的按钮）。
     * @returns {boolean} 本世代是否可能存在筛选按钮（false = 一定没有）
     */
    public hasFilterCells(): boolean {
        return this._filterCellsDirty
            || this._filterCellsEpoch !== this._valuesEpoch
            || this._filterCells.length > 0;
    }

    /**
     * 标记「筛选标记集合」已变化，使 {@link hasFilterCells} 与 {@link getFilterCells}
     * 下次都重新判定
     *
     * 增删筛选标记既不改 `values.length`、也不一定改 `_valuesEpoch`
     * （如给已存在的单元格挂 `filter`），若不显式调用：
     * - {@link hasFilterCells} 会一直沿用上一次「没有按钮」的结论；
     * - {@link getFilterCells} 里 `_ensureCellIndex()` 会认为索引仍是最新的，
     *   直接返回**不含新按钮**的旧缓存数组 —— 按钮就再也建不出来了。
     *
     * 因此这里同时置位 {@link _filterCellsDirty}（由 {@link _ensureCellIndex}
     * 负责真正重扫）与失效 {@link _filterCellsEpoch}，一次显式调用即覆盖
     * 「要不要问」与「问出来是什么」两个入口。
     *
     * ⚠️ **不要**再去动 `_indexEpoch`：它是「单元格索引内容世代」的标记，
     * 与筛选标记集合无关；置 -1 只会让 `_ensureCellIndex` 走「内容变了但结构没变」
     * 分支，而该分支会把 `_filterCellsEpoch` 刷回 `_valuesEpoch`，
     * 恰好抹掉这里的置脏意图（历史 bug：按钮不显示/不消失）。
     *
     * 同样注意**不要**在这里自增 `_valuesEpoch` —— 那会让整个单元格索引
     * （含合并倒排）一并失效重建，代价远大于收益。
     */
    public invalidateFilterCells(): void {
        this._filterCellsDirty = true;
        this._filterCellsEpoch = -1;
    }

    /**
     * 确保「筛选清单索引」处于最新状态（同一世代只构建一次）
     *
     * 产出三样东西，供筛选交互共用：
     * - {@link _contentMaxRow} / {@link _contentMaxCol}：内容范围。筛选下拉原先用
     *   `matrixValues()` 求「统计到第几行」—— 那会**分配一个 maxRow × maxCol 的稠密二维数组**
     *   （10 万行 × 50 列 ≈ 500 万个 `undefined` 槽位、几十 MB 临时内存），
     *   而调用方只需要一个行号上界。改为随本索引一次算出。
     * - {@link _rowsWithText}：筛选只可能隐藏有内容的行，用它把「逐行扫到总行数」
     *   收敛成「只扫有内容的行」。
     * - {@link _colsWithContent}：添加筛选按钮时要逐列判断「该列是否有内容」，
     *   用它把「每列重建一次稠密矩阵」收敛成 O(1) 查表（见该字段注释）。
     * @private
     */
    protected _ensureFilterIndex(): void {
        if (this._filterIndexEpoch === this._valuesEpoch
            && this._filterIndexLength === this._values.length) {
            return;
        }
        const values = this._values;
        const rowsWithText = new Set<number>();
        const colsWithContent = new Set<number>();
        let maxRow = 0;
        let maxCol = 0;
        for (let i = 0; i < values.length; i++) {
            const cell = values[i];
            // 无文本的单元格不计入内容范围（与「筛选只讨论内容」的语义一致），
            // 也省掉一次名字解析
            if (!SheetStoreBase.hasAnyText(cell)) continue;
            const { col, row } = this.getCellColAndRow(cell.cell);
            const endCol = col + (cell.colspan === undefined ? 1 : cell.colspan) - 1;
            const endRow = row + (cell.rowspan === undefined ? 1 : cell.rowspan) - 1;
            if (endRow > maxRow) maxRow = endRow;
            if (endCol > maxCol) maxCol = endCol;
            for (let r = row; r <= endRow; r++) rowsWithText.add(r);
            // 合并块覆盖到的每一列都算「有内容」——与历史 validataColInFilter
            // 「矩阵里该列任一格非 undefined」的口径一致
            for (let c = col; c <= endCol; c++) colsWithContent.add(c);
        }
        this._contentMaxRow = maxRow;
        this._contentMaxCol = maxCol;
        this._rowsWithText = rowsWithText;
        this._colsWithContent = colsWithContent;
        this._filterIndexEpoch = this._valuesEpoch;
        this._filterIndexLength = values.length;
    }

    /**
     * 添加过滤范围, 如果已存在过滤范围, 则清除所有过滤范围
     *
     * 开关操作整体包裹在 runWithFullStateUndo 中，作为一步可撤销/重做操作：
     * 快照包含 cell.filter 标记（values）、筛选条件注册表与行隐藏状态，
     * 撤销「清除筛选按钮」可完整恢复此前的筛选按钮、条件与隐藏行。
     * 包裹期间 commitValues 的自动快照与广播被抑制，故在动作内显式广播 VALUES_CHANGED。
     */
    public setFilter(): void {
        if (!this.activedCell) return;

        // 已存在筛选按钮 → 清除全部按钮（一步撤销）
        // 判定走世代缓存（O(1)）而不是整表 values.some(...)：本方法每次点击都会走
        // 这一行，大表下整表遍历一次实测 0.83 ms（基准 test/_bench-filter-add.js）。
        if (this.validataFilter()) {
            this.runWithFullStateUndo('清除筛选按钮', () => {
                // 必须在清条件**之前**算：被筛选隐藏的行要靠「当前条件重放」才认得出来
                const filterHidden = this.computeFilterHiddenRows();
                // 清除筛选按钮标记
                for (const cell of this.values) {
                    if (cell.filter !== undefined) cell.filter = undefined;
                }
                this.invalidateFilterCells();
                // 清除筛选条件注册表（按钮实例销毁时也会逐列清理，这里一次性清干净）
                this.clearAllFilterConditions();
                // 取消筛选 = 同时把被筛选隐藏的行恢复显示。
                // **只恢复「因筛选而隐藏」的行**：手动隐藏的行不在 filterHidden 中，保持隐藏。
                this.restoreRowsHiddenByFilter(filterHidden);
                this.emit(DataEvents.VALUES_CHANGED, this._values);
            }, 'filters');
            return;
        }

        const [start, end] = this._selection.split(':');
        const { col: startCol, row: startRow } = this.getCellColAndRow(start);
        const { col: endCol, row: endRow } = this.getCellColAndRow(end);
        // 只需知道「选中行这一行有没有内容」，不必为它建整张稠密矩阵：
        // 历史实现 `const rows = matrix[startRow - 1]` 只为拿到这一行，
        // 却先付了一次 O(maxRow × maxCol) 的矩阵分配（6 万格实测 8.9 ms）。
        // 改为在选中列范围内逐格 O(1) 查名索引。
        let rowHasContent = false;
        for (let c = startCol; c <= endCol; c++) {
            if (this.getCellByName(this.getCellName(c, startRow)) !== undefined) {
                rowHasContent = true;
                break;
            }
        }
        if (start !== end && startCol !== endCol && startRow === endRow && rowHasContent) {
            // 为选中行范围内的有效列逐列添加筛选按钮（一步撤销）
            this.runWithFullStateUndo('添加筛选按钮', () => {
                for (let c = startCol - 1; c < endCol; c++) {
                    if (!this.validataColInFilter(c)) continue;
                    // 合并区内部的列不单独立筛选按钮：那一格在合并态下被遮挡层盖住，
                    // 按钮根本画不出来。重定向到该合并块的锚点（见 ensureWritableCell）。
                    this.ensureWritableCell(this.getCellName(c + 1, startRow)).filter = startRow;
                }
                this.invalidateFilterCells();
                this.emit(DataEvents.VALUES_CHANGED, this._values);
                // 'selection' 档：只写选中行上那几个筛选按钮格，影响面 ⊆ 选区
            }, 'selection');
        } else {
            this.messageBox.confirm(
                `执行过滤操作时，请选择同一行的单元格范围。`,
                () => {
                },
                () => {
                }
            );
        }
    }

    /**
     * 恢复「被筛选隐藏」的行（取消筛选时调用）
     *
     * 连续行合并成一段调用 {@link setRowsHidden} —— 统一入口自带
     * 可见位置重算 / 滚动偏移 clamp / 尾部补足与收缩 / 广播，
     * **不要绕过它直接改 `isHidden`**（否则行位置与总高会和隐藏状态脱节）。
     * 已处于显示状态的行直接跳过，避免无谓的重算。
     * @param {Set<number>} rows - 因筛选而隐藏的行号（1-based）
     * @private
     */
    private restoreRowsHiddenByFilter(rows: Set<number>): void {
        if (rows.size === 0) return;
        const sorted = Array.from(rows).sort((a, b) => a - b);
        let segStart = -1;
        let segEnd = -1;
        const flush = (): void => {
            if (segStart >= 0) this.setRowsHidden(segStart, segEnd, false);
            segStart = -1;
            segEnd = -1;
        };
        for (let i = 0; i < sorted.length; i++) {
            const r = sorted[i];
            if (this.rowHeaders.getAt(r - 1)?.isHidden !== true) continue;  // 已显示：无需恢复
            if (segStart < 0) {
                segStart = r;
                segEnd = r;
            } else if (r === segEnd + 1) {
                segEnd = r;   // 与上一行连续 → 并入同一段
            } else {
                flush();
                segStart = r;
                segEnd = r;
            }
        }
        flush();
    }

    /**
     * 获取指定列的筛选条件
     * @param {number} col - 列号（1-based）
     * @returns {{ anchorRow: number; keep: Set<string> } | undefined} 筛选条件；该列未筛选时为 undefined
     */
    public getFilterCondition(col: number): { anchorRow: number; keep: Set<string> } | undefined {
        return this._filterConditions.get(col);
    }

    /**
     * 获取筛选操作顺序（已筛选列号按首次执行筛选的先后排列）
     * @returns {number[]} 列号数组副本
     */
    public getFilterOrder(): number[] {
        return [...this._filterOrder];
    }

    /**
     * 写入（或覆盖）指定列的筛选条件，并按首次筛选顺序登记列号
     * 仅供筛选组件在 runWithFullStateUndo 包裹内调用，条件随全量快照记录。
     * @param {number} col - 列号（1-based）
     * @param {number} anchorRow - 筛选按钮所在行号（1-based）
     * @param {Set<string>} keep - 保留内容集合
     */
    public setFilterCondition(col: number, anchorRow: number, keep: Set<string>): void {
        if (!this._filterConditions.has(col)) {
            this._filterOrder.push(col);
        }
        this._filterConditions.set(col, { anchorRow, keep });
    }

    /**
     * 清除指定列的筛选条件并移出操作顺序表
     * @param {number} col - 列号（1-based）
     */
    public clearFilterCondition(col: number): void {
        this._filterConditions.delete(col);
        this._filterOrder = this._filterOrder.filter(c => c !== col);
    }

    /**
     * 清除全部筛选条件（清空注册表与操作顺序表）
     */
    public clearAllFilterConditions(): void {
        if (this._filterConditions.size === 0 && this._filterOrder.length === 0) return;
        this._filterConditions = new Map();
        this._filterOrder = [];
    }

    /**
     * 按**当前筛选条件注册表**重放，算出「应当因筛选而隐藏的行」（1-based）
     *
     * 「执行筛选」与「清除筛选（恢复隐藏行）」两侧共用这一份判定，避免出现
     * 两份相似但不一致的实现（漂移的后果是「取消筛选后有些行再也回不来」）。
     *
     * 规则（与 Excel 筛选一致）：
     * - 只可能隐藏**有内容**的行（全空行始终显示）；
     * - 某列的约束只作用于「该列筛选行**以下**」的行（筛选行及其以上不受约束）；
     * - 跨越某列筛选行的合并块覆盖的行不受该列约束（合并块整体显示）；
     * - 行内容取「视觉生效」的文本：合并区域内取锚点文本（{@link getEffectiveTextAt}）。
     *
     * 复杂度 O(有内容的行 × 已筛选列)，与「总行数」无关 —— 大文件中总行数可达百万行，
     * 而真正有内容的往往只有几千行。
     * @returns {Set<number>} 应隐藏的行号集合（1-based）
     */
    public computeFilterHiddenRows(): Set<number> {
        const hidden = new Set<number>();
        if (this._filterConditions.size === 0) return hidden;
        const rowsWithText = this.getRowsWithText();
        if (rowsWithText.size === 0) return hidden;
        const totalRows = this.rowHeaders.length;
        // 各列约束只作用于自己的筛选行以下 → 全部筛选行之前的行不可能被隐藏
        let firstAnchor = Infinity;
        this._filterConditions.forEach((cond) => {
            if (cond.anchorRow < firstAnchor) firstAnchor = cond.anchorRow;
        });
        const order = this._filterOrder;
        rowsWithText.forEach((r) => {
            if (r <= firstAnchor || r > totalRows) return;
            for (let k = 0; k < order.length; k++) {
                const c = order[k];
                const cond = this._filterConditions.get(c);
                if (!cond || r <= cond.anchorRow) continue;   // 该列筛选行及以上不受此条件约束
                const cell = this.getEffectiveCellAt(r, c);
                if (cell !== undefined && this.getCellColAndRow(cell.cell).row <= cond.anchorRow) continue;  // 跨筛选行的合并块
                const text = cell === undefined ? '' : cell.getText();
                if (!cond.keep.has(text)) {
                    hidden.add(r);
                    return;
                }
            }
        });
        return hidden;
    }

    /**
     * 验证列是否有效（该列是否有内容且未隐藏）—— O(1)
     *
     * 历史实现对每一列都调用 `matrixValues()`，那会**分配整张 `maxRow × maxCol`
     * 稠密二维数组**并遍历全表两次；添加筛选按钮时它在逐列循环里被调用，
     * 于是「选中 N 列 = 重建 N 次稠密矩阵」（实测 6 万格下 200 列 = 1596 ms）。
     * 现在改为查 {@link _colsWithContent}（一次遍历、按世代缓存）。
     *
     * 语义与历史实现保持一致：
     * - `col` 越界（负数 / 超过内容最大列）→ false；
     * - 该列被隐藏 → false（不给隐藏列加筛选按钮）；
     * - 该列没有任何单元格（含被合并块覆盖的情况）→ false。
     * @param {number} col - 列索引（0-based）
     * @returns {boolean} 该列是否可加筛选按钮
     * @private
     */
    private validataColInFilter(col: number): boolean {
        if (col < 0) return false;
        const colNo = col + 1; // 内部索引 1-based，与 _colsWithContent 口径一致
        if (colNo > this.getContentBounds().maxCol) return false;
        if (this.colHeaders.getAt(col)?.isHidden === true) return false; // 跳过隐藏列
        return this.getColsWithContent().has(colNo);
    }

    /**
     * 验证「当前是否存在筛选按钮」
     *
     * 该接口被 `sectionProcess.setToolbarsState()` 挂在**每次 VALUES_CHANGED** 上
     * （含每次按键、每次滚动触发的广播），因此必须是 O(1)。
     * 历史实现 `this.values.filter(...)` 是整表遍历 + 建临时数组：
     * 基准 `test/_bench-filter-add.js` 实测 6 万格下 0.74 ms/次。
     *
     * 改用与 {@link hasFilterCells} 同一套世代缓存：同一世代内先问
     * `_filterCellsEpoch === _valuesEpoch`（不重建任何索引即可回答），
     * 世代失效才走一次 `getFilterCells()`（顺带产出后续渲染要用的缓存）。
     * @returns {boolean} 是否存在筛选按钮
     */
    public validataFilter(): boolean {
        if (!this._filterCellsDirty
            && this._filterCellsEpoch === this._valuesEpoch
            && this._values.length === this._indexedLength) {
            return this._filterCells.length > 0;
        }
        return this.getFilterCells().length > 0;
    }

    /**
     * 验证列是否隐藏
     * @param {number} col - 列索引（0-based）
     * @returns {boolean} 是否隐藏（索引越界或未设置时返回 false）
     */
    public isColHidden(col: number): boolean {
        return this.colHeaders.getAt(col)?.isHidden === true;
    }

    /**
     * 验证行是否隐藏
     * @param {number} row - 行索引（0-based）
     * @returns {boolean} 是否隐藏（索引越界或未设置时返回 false）
     */
    public isRowHidden(row: number): boolean {
        return this.rowHeaders.getAt(row)?.isHidden === true;
    }

    /**
     * 查找第一个可见的列（1-based 索引）
     * @returns {number} 第一个可见列号；所有列都隐藏时返回 0
     * @private
     */
    protected _findFirstVisibleCol(): number {
        for (let c = 0; c < this.colHeaders.length; c++) {
            if (!this.colHeaders.getAt(c).isHidden) return c + 1;
        }
        return 0;
    }

    /**
     * 查找最后一个可见的列（1-based 索引）
     * @returns {number} 最后一个可见列号；所有列都隐藏时返回 0
     * @private
     */
    protected _findLastVisibleCol(): number {
        for (let c = this.colHeaders.length - 1; c >= 0; c--) {
            if (!this.colHeaders.getAt(c).isHidden) return c + 1;
        }
        return 0;
    }

    /**
     * 查找第一个可见的行（1-based 索引）
     * @returns {number} 第一个可见行号；所有行都隐藏时返回 0
     * @private
     */
    protected _findFirstVisibleRow(): number {
        for (let r = 0; r < this.rowHeaders.length; r++) {
            if (!this.rowHeaders.getAt(r).isHidden) return r + 1;
        }
        return 0;
    }

    /**
     * 查找最后一个可见的行（1-based 索引）
     * @returns {number} 最后一个可见行号；所有行都隐藏时返回 0
     * @private
     */
    protected _findLastVisibleRow(): number {
        for (let r = this.rowHeaders.length - 1; r >= 0; r--) {
            if (!this.rowHeaders.getAt(r).isHidden) return r + 1;
        }
        return 0;
    }

    /**
     * 设置指定范围行的隐藏状态
     *
     * 同步维护 {@link _hiddenRowsSet} 倒排索引：隐藏=true 加入集合，=false 移除。
     * 筛选场景需遍历「当前已隐藏的行」做差量判定时，改为访问 {@link getHiddenRowsSet}
     * 而非扫到总行数（百万级大表的卡顿主因）。
     * @param {number} startRow - 起始行号（1-based，含）
     * @param {number} endRow - 结束行号（1-based，含）
     * @param {boolean} hidden - 是否隐藏
     */
    public setRowsHidden(startRow: number, endRow: number, hidden: boolean): void {
        for (let r = startRow; r <= endRow; r++) {
            const rh = this.rowHeaders.getAt(r - 1);
            if (rh) rh.isHidden = hidden;
            if (hidden) this._hiddenRowsSet.add(r);
            else this._hiddenRowsSet.delete(r);
        }
        // 重算可见行 top 位置（隐藏行不占位，下方行上移覆盖）+ 总高度
        this.rowHeaders.recalcRowPositions();
        const viewportHeight = this.visibleView.sheetHeight / this.zoom;
        // 隐藏后总高度变小，clamp 滚动偏移防止越界（避免 getVisibleRows 循环跑穿）
        const maxOffset = Math.max(0, this.rowHeaders.allRowHeight - viewportHeight);
        if (this.rowHeaders.offsetHeight > maxOffset) {
            this.rowHeaders.offsetHeight = maxOffset;
        }
        if (hidden) {
            // 隐藏：可见总高不足一屏时一次性补足（按可见总高判断，兼容自定义行高）
            // 【性能约定，勿回退为「while 逐行 addRow + recalcRowPositions」】
            // 每次循环 addRow 末尾 + recalcRowPositions 都是 O(N)；补几十行 = O(N×几十)
            const deficit = viewportHeight - this.rowHeaders.allRowHeight;
            if (deficit > 0) {
                const need = Math.ceil(deficit / DEFAULT_CELL_HEIGHT) + 1;
                this.rowHeaders.addRow(this.rowHeaders.length, this.rowHeaders.length + need);
                this.rowHeaders.recalcRowPositions();
            }
        } else {
            // 取消隐藏：可见总高变大，与屏幕比较后收缩尾部多余的无内容行（保证网格线仍能填满屏幕）。
            // 本次取消隐藏的行受保护不参与收缩，否则取消隐藏尾部空行会立刻被删导致操作无效
            const protectedRows = new Set<number>();
            for (let r = startRow; r <= endRow; r++) protectedRows.add(r);
            this.trimRowsToViewport(viewportHeight, protectedRows);
        }
        this.emit(DataEvents.VALUES_CHANGED, this._values);
    }

    /**
     * 取「当前已隐藏的行」集合引用（可变，筛选差量更新可直接增删）
     *
     * 由 {@link setRowsHidden} / {@link setRowHiddenFlag} 同步维护。
     * 筛选（{@link Filter.applyFilter}）差量更新时用它替代「扫到总行数」的
     * O(N) 全表扫描——百万行大表已隐藏行通常只有几千行。
     * @returns {Set<number>} 已隐藏行号集合（1-based）
     */
    public getHiddenRowsSet(): Set<number> {
        return this._hiddenRowsSet;
    }

    /**
     * 仅改单行隐藏标记 + 维护索引，不做收尾
     *
     * 供筛选批量应用时逐行调用，全部改完后由 {@link finalizeRowVisibility}
     * 统一做一次 recalc + clamp + 补足/收缩 + 广播。避免了「逐行调
     * setRowsHidden」产生 N 次 recalc + N 次广播/重绘的卡顿。
     * @param {number} row - 行号（1-based）
     * @param {boolean} hidden - 是否隐藏
     */
    public setRowHiddenFlag(row: number, hidden: boolean): void {
        const rh = this.rowHeaders.getAt(row - 1);
        if (rh) rh.isHidden = hidden;
        if (hidden) this._hiddenRowsSet.add(row);
        else this._hiddenRowsSet.delete(row);
    }

    /**
     * 批量应用行隐藏状态后做一次收尾
     *
     * 调用方应**先直接改 RowHeader.isHidden 标记**（同步维护 {@link _hiddenRowsSet}），
     * 全部改完后调本方法做一次 recalcRowPositions + clamp + 补足/收缩 + 广播。
     *
     * 这是为了让筛选（applyFilter）能一次性应用所有需要变化的行段，避免
     * 「分段调用 setRowsHidden」产生多次 recalc + 多次 emit/重绘——大表上
     * 多段连续隐藏/取消隐藏会产生可感卡顿。
     *
     * @param {number} protectedStart - 受保护行段起点（1-based，取消隐藏时不参与尾部收缩）
     * @param {number} protectedEnd - 受保护行段终点（1-based）
     * @param {boolean} hadHide - 本次是否有「隐藏」操作（决定是否走补足分支）
     * @param {boolean} hadUnhide - 本次是否有「取消隐藏」操作（决定是否走收缩分支）
     */
    public finalizeRowVisibility(
        protectedStart: number, protectedEnd: number,
        hadHide: boolean, hadUnhide: boolean
    ): void {
        this.rowHeaders.recalcRowPositions();
        const viewportHeight = this.visibleView.sheetHeight / this.zoom;
        const maxOffset = Math.max(0, this.rowHeaders.allRowHeight - viewportHeight);
        if (this.rowHeaders.offsetHeight > maxOffset) {
            this.rowHeaders.offsetHeight = maxOffset;
        }
        if (hadHide) {
            // 一次性补足：避免 while 逐行 addRow + recalcRowPositions 的 O(N×补足行数)
            const deficit = viewportHeight - this.rowHeaders.allRowHeight;
            if (deficit > 0) {
                const need = Math.ceil(deficit / DEFAULT_CELL_HEIGHT) + 1;
                this.rowHeaders.addRow(this.rowHeaders.length, this.rowHeaders.length + need);
                this.rowHeaders.recalcRowPositions();
            }
        }
        if (hadUnhide) {
            const protectedRows = new Set<number>();
            for (let r = protectedStart; r <= protectedEnd; r++) protectedRows.add(r);
            this.trimRowsToViewport(viewportHeight, protectedRows);
        }
        this.emit(DataEvents.VALUES_CHANGED, this._values);
    }

    /**
     * 收缩尾部无内容的行：从 rowHeaders 末尾删除无任何单元格内容的行，
     * 直到遇到有内容的行，或删除后可见总高将不足以填满视口为止。
     * 用于取消隐藏后清理历史补足/扩行产生的多余空行。
     * @param {number} viewportHeight - 视口高度（数据坐标，已除 zoom）
     * @param {Set<number>} protectedRows - 受保护的行号（1-based），本次操作涉及的行不参与收缩
     * @private
     * @returns {void}
     */
    protected trimRowsToViewport(viewportHeight: number, protectedRows: Set<number>): void {
        // 有内容的行集合（1-based）：被非空文本单元格覆盖的行（合并锚点 rowspan 覆盖行都算）
        const nonEmptyRows = new Set<number>();
        for (const cell of this.values) {
            if ((cell.chars?.map(c => c.char).join('') ?? '') === '') continue;
            const { row: r1 } = this.getCellColAndRow(cell.cell);
            for (let r = r1; r < r1 + (cell.rowspan ?? 1); r++) nonEmptyRows.add(r);
        }
        let total = this.rowHeaders.allRowHeight;
        while (this.rowHeaders.length > 0) {
            const rh = this.rowHeaders.getAt(this.rowHeaders.length - 1);
            if (!rh) break;
            const rowNo = this.rowHeaders.length;
            if (nonEmptyRows.has(rowNo)) break;    // 尾行有内容（行号 = 数组长度）
            if (protectedRows.has(rowNo)) break;   // 尾行在本次操作范围内，保护不删
            if (total - rh.height < viewportHeight) break;  // 删除后填不满视口，停止
            total -= rh.height;
            this.rowHeaders.rowHeaders.pop();
        }
        if (total !== this.rowHeaders.allRowHeight) {
            this.rowHeaders.recalcAllRowHeight();
            this.rowHeaders.recalcRowPositions();
            // 总高变小，clamp 滚动偏移防止越界
            const newMax = Math.max(0, this.rowHeaders.allRowHeight - viewportHeight);
            if (this.rowHeaders.offsetHeight > newMax) {
                this.rowHeaders.offsetHeight = newMax;
            }
            // 激活单元格/选区可能位于被删区域，收敛到当前范围防止后续读取越界
            this.constrainSelectionToBounds();
        }
    }

    /**
     * 将激活单元格与选区收敛到当前行/列数组范围内。
     * 行/列取消隐藏后尾部空行/列被收缩删除，位于被删区域的激活单元格/选区若不收敛，
     * 将导致编辑光标定位（calculateCellRect）、选区绘制（getSelectedColsAndRows）等读取越界。
     * @private
     * @returns {void}
     */
    private constrainSelectionToBounds(): void {
        const maxCol = Math.max(1, this.colHeaders.length);
        const maxRow = Math.max(1, this.rowHeaders.length);
        const clampCell = (name: string): string => {
            const { col, row } = this.getCellColAndRow(name);
            return `${this.getColName(Math.min(col, maxCol))}${Math.min(row, maxRow)}`;
        };
        // 激活单元格收敛（直接改私有字段，绕过 setter 的副作用）
        this._activedCell = clampCell(this._activedCell);
        // 选区起止分别收敛
        const [selStart, selEnd] = this._selection.split(':');
        this._selection = `${clampCell(selStart)}:${clampCell(selEnd)}`;
    }

    /**
     * 设置指定范围列的隐藏状态
     * @param {number} startCol - 起始列号（1-based，含）
     * @param {number} endCol - 结束列号（1-based，含）
     * @param {boolean} hidden - 是否隐藏
     */
    public setColsHidden(startCol: number, endCol: number, hidden: boolean): void {
        for (let c = startCol; c <= endCol; c++) {
            const ch = this.colHeaders.getAt(c - 1);
            if (ch) ch.isHidden = hidden;
        }
        // 重算可见列 left 位置（隐藏列不占位，右侧列左移覆盖）+ 总宽度
        this.colHeaders.recalcColPositions();
        const viewportWidth = this.visibleView.sheetWidth / this.zoom;
        // 隐藏后总宽度变小，clamp 滚动偏移防止越界（避免 getVisibleCols 循环跑穿）
        const maxOffset = Math.max(0, this.colHeaders.allColWidth - viewportWidth);
        if (this.colHeaders.offsetWidth > maxOffset) {
            this.colHeaders.offsetWidth = maxOffset;
        }
        if (hidden) {
            // 隐藏：可见总宽不足一屏时一次性补足（按可见总宽判断，兼容自定义列宽）
            // 【性能约定，勿回退为「while 逐列 addCol + recalcColPositions」】
            const deficit = viewportWidth - this.colHeaders.allColWidth;
            if (deficit > 0) {
                const need = Math.ceil(deficit / DEFAULT_CELL_WIDTH) + 1;
                this.colHeaders.addCol(this.colHeaders.length, this.colHeaders.length + need);
                this.colHeaders.recalcColPositions();
            }
        } else {
            // 取消隐藏：可见总宽变大，与屏幕比较后收缩尾部多余的无内容列（保证网格线仍能填满屏幕）。
            // 本次取消隐藏的列受保护不参与收缩，否则取消隐藏尾部空列会立刻被删导致操作无效
            const protectedCols = new Set<number>();
            for (let c = startCol; c <= endCol; c++) protectedCols.add(c);
            this.trimColsToViewport(viewportWidth, protectedCols);
        }
        this.emit(DataEvents.VALUES_CHANGED, this._values);
    }

    /**
     * 收缩尾部无内容的列：从 colHeaders 末尾删除无任何单元格内容的列，
     * 直到遇到有内容的列，或删除后可见总宽将不足以填满视口为止。
     * 用于取消隐藏后清理历史补足/扩列产生的多余空列。
     * @param {number} viewportWidth - 视口宽度（数据坐标，已除 zoom）
     * @param {Set<number>} protectedCols - 受保护的列号（1-based），本次操作涉及的列不参与收缩
     * @private
     * @returns {void}
     */
    protected trimColsToViewport(viewportWidth: number, protectedCols: Set<number>): void {
        // 有内容的列集合（1-based）：被非空文本单元格覆盖的列（合并锚点 colspan 覆盖列都算）
        const nonEmptyCols = new Set<number>();
        for (const cell of this.values) {
            if ((cell.chars?.map(c => c.char).join('') ?? '') === '') continue;
            const { col: c1 } = this.getCellColAndRow(cell.cell);
            for (let c = c1; c < c1 + (cell.colspan ?? 1); c++) nonEmptyCols.add(c);
        }
        let total = this.colHeaders.allColWidth;
        while (this.colHeaders.length > 0) {
            const ch = this.colHeaders.getAt(this.colHeaders.length - 1);
            if (!ch) break;
            const colNo = this.colHeaders.length;
            if (nonEmptyCols.has(colNo)) break;    // 尾列有内容（列号 = 数组长度）
            if (protectedCols.has(colNo)) break;   // 尾列在本次操作范围内，保护不删
            if (total - ch.width < viewportWidth) break;  // 删除后填不满视口，停止
            total -= ch.width;
            this.colHeaders.colHeaders.pop();
        }
        if (total !== this.colHeaders.allColWidth) {
            this.colHeaders.recalcAllColWidth();
            this.colHeaders.recalcColPositions();
            // 总宽变小，clamp 滚动偏移防止越界
            const newMax = Math.max(0, this.colHeaders.allColWidth - viewportWidth);
            if (this.colHeaders.offsetWidth > newMax) {
                this.colHeaders.offsetWidth = newMax;
            }
            // 激活单元格/选区可能位于被删区域，收敛到当前范围防止后续读取越界
            this.constrainSelectionToBounds();
        }
    }
}
