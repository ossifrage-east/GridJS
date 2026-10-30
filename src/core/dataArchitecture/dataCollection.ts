import { ColHeader, ColHeaders, RowHeader, RowHeaders, encodeColAxis, encodeRowAxis, decodeColAxis, decodeRowAxis, SparseAxis } from "./header";
import { Cell, CellStyleTable, TextAlign, VerticalAlign } from "./cell";
import { LineText, SelectionRange, FontStyle, BorderStyle, CellBorderStyle, CursorPosition, } from "./cellContent";
import { Char } from "./char";
import { VisibleView } from "./visibleView";
import { CELL_PADDING, DataEvents, DEFAULT_CELL_HEIGHT, DEFAULT_CELL_WIDTH, DEFAULT_SELECTED_BORDER_WIDTH, MouseLocation } from "../constant";
import { EventEmitter } from "../../utils/eventEmitter";
import { IndexDB } from "./indexDB";
import { UndoManager, UndoAction, FullStateSnapshot, CellPatchEntry, UndoAnchor, StatePatchSnapshot, StructuralSnapshot, ContinuousUndoSnapshot, SortPermutationSnapshot, ValuesSnapshot } from "../../utils/undoManager";
import { ClipboardManager } from "../../utils/clipboardManager";
import { MessageBox } from "../common/message";
import { MatrixSorter, type SortConfig } from "./matrixSorter";
import { PrinterSetting } from "../sheets/sheetPrinter";
import { SliceBudget } from "../../utils/scheduler";
import { SheetSearchBase } from "../sheets/sheetSearch";
import type { ExtractedImage } from "../../utils/parser/sheetParser";


/**
 * 合并区域的数值矩形（闭区间，1-based 行列号）
 *
 * 由 {@link DataCollection.getMergedCells} 的字符串形式（`A1:C3`）在
 * 索引重建时换算而来，供渲染热路径做纯整数比较，避免逐次字符串解析。
 * @interface MergeRect
 */
export interface MergeRect {
    /** 锚点单元格名（合并区域的左上角，如 `A1`） */
    anchor: string;
    /** 起始列号（1-based） */
    startCol: number;
    /** 起始行号（1-based） */
    startRow: number;
    /** 结束列号（1-based，含） */
    endCol: number;
    /** 结束行号（1-based，含） */
    endRow: number;
}

/**
 * 脚本（Python）可写的单元格字段白名单（值以外的样式 / 展示属性）
 *
 * 键名就是 `Cell` 上的公开属性名：Python 侧的别名（`color` / `bgColor` / `bold` …）
 * 在 `sheetNameSpace.PY_NAMESPACE_SOURCE` 里翻译成这些名字再回传，两边必须同时改。
 * 用**白名单**而不是「拿到的键全赋值」：脚本回传的是 JSON，键由用户脚本决定，
 * 直接赋值等于把 `cell` 身上的任意属性（含 `cell` 自身、`chars`）暴露给脚本。
 * 合并跨度（colspan / rowspan）单独处理，见 {@link DataCollection.applyScriptWrites}。
 */
const SCRIPT_WRITABLE_FIELDS: ReadonlySet<string> = new Set([
    'fontFamily', 'fontSize', 'fontWeight', 'fontStyle', 'fontColor',
    'underline', 'strikethrough',
    'backgroundColor', 'borderColor',
    'borderTopWidth', 'borderBottomWidth', 'borderLeftWidth', 'borderRightWidth',
    'textAlign', 'alignItems', 'wrap', 'letterSpacing', 'lineSpacing',
    'numberFormat',
]);

/**
 * 查询 / 替换的作用范围矩形（1-based 闭区间）
 *
 * 是**数值**矩形而非 `'A1:C9'` 字符串：扫描循环里每格都要用它做比较，
 * 字符串形态每次都得重新解析。字符串只在「输入/展示」两端出现
 * （见 {@link DataCollection.resolveSearchRect} / {@link DataCollection.formatRect}）。
 */
export interface SheetRect {
    /** 起始列号（1-based，含） */
    startCol: number;
    /** 起始行号（1-based，含） */
    startRow: number;
    /** 结束列号（1-based，含） */
    endCol: number;
    /** 结束行号（1-based，含） */
    endRow: number;
}

/**
 * 查询 / 替换的匹配选项
 *
 * 查询与替换共用同一套选项对象是硬要求：若两边算法不一致，
 * 会出现「查得到但替换不掉」（或反过来）的错位。
 */
export interface SearchMatchOptions {
    /** 区分大小写（默认 false） */
    matchCase?: boolean;
    /** 单元格内容**完全匹配**才算命中（默认 false，即「包含即命中」） */
    matchEntire?: boolean;
    /** 通配符模式：`?` 匹配单个字符、`*` 匹配任意长度字符串（默认 false，关键字按字面量处理） */
    useWildcard?: boolean;
}

/**
 * 一条查询命中
 *
 * `text` 是命中**当时**的单元格文本快照：替换发生之后（无论「替换」还是「全部替换」）
 * UI 都应当重新查一次，拿着旧快照继续展示会显示过期内容。
 */
export interface SearchHit {
    /** 单元格名称（如 `'B3'`） */
    cell: string;
    /** 列号（1-based） */
    col: number;
    /** 行号（1-based） */
    row: number;
    /** 命中时的单元格全文 */
    text: string;
}


/**
 * 工作表存档 v4（**默认落库形态**）
 *
 * ## 为什么不再是「一个 JSON 字符串」（性能约定，勿回退）
 *
 * v3 时代整份快照是一根字符串，装载时必然要一次 `JSON.parse(整份)`：
 * 40 万格实测约 140~290 ms、膨胀维度那份 42.3 MB 要 620 ms，而且是
 * **整段同步、不可让出**的 —— 分片（`SliceBudget`）能解决「解码单元格」
 * 的成本，解决不了这一次解析。它正是装载路径最后一块不可切分的地板。
 *
 * v4 把存档拆成「结构化克隆对象」直存 IndexedDB（不再需要字符串化），
 * 其中最大的部分 `chunks` 是**分片的 JSON 文本数组**（每片
 * {@link DataCollection.SNAPSHOT_VALUE_CHUNK} 格），装载时一片一片解析：
 * 单片最长约 1.8 ms，片间可让出主线程 —— 从「一次 290 ms 的冻结」
 * 变成「若干次 2 ms 的小块」。
 *
 * 实测（40 万格，`test/_bench-snapshot-form.js`）：
 *
 * | 存档形态 | 装载耗时 | 最大单次不可分片阻塞 |
 * | --- | --- | --- |
 * | v3 整份 JSON 字符串 | 428 ms | **290 ms** |
 * | 结构化克隆普通对象直存 | 951 ms | 560 ms |
 * | v4 分片 JSON + 稀疏行列头 | 约 320 ms | 42 ms（IDB 读回）/ 1.8 ms（解析） |
 *
 * 「结构化克隆普通对象」那一行说明**不能**天真地把整个 `values` 当对象数组存：
 * V8 的结构化克隆要逐对象重建，比 `JSON.parse` 还慢一倍。
 *
 * `styles` 存**数组**而不是字符串：表项数量是样式种类数（通常几十到几百条），
 * 直接用结构化克隆比再 `JSON.parse` 一次便宜。
 */
export interface SheetSnapshotV4 {
    /** 格式版本（固定 4，读取侧据此分派） */
    v: 4;
    /** 当前选中的单元格 */
    activedCell: string;
    /** 选区（如 `A1:C3`） */
    selection: string;
    /** 打印设置 */
    printSetting: PrinterSetting;
    /** 是否显示网格线 */
    showGridLines: boolean;
    /** 缩放比例 */
    zoom: number;
    /** 水平滚动偏移 */
    offsetWidth: number;
    /** 垂直滚动偏移 */
    offsetHeight: number;
    /** 列头（稀疏编码，见 `SparseAxis`） */
    cols: SparseAxis;
    /** 行头（稀疏编码，见 `SparseAxis`） */
    rows: SparseAxis;
    /** values 的 JSON 文本分片，每片是一个完整的 JSON 数组字符串 */
    chunks: string[];
    /** 样式表（`CellStyleTable.entries`） */
    styles: Array<Record<string, unknown>>;
    /**
     * 本表的 Python 编辑器脚本
     *
     * 键名沿用编辑器侧的约定（`py-editor:script`）：脚本属于**文件内容**，
     * 存在本表的 Value 记录里而不是 settings —— 重命名 / 复制 / 删除文件时
     * 跟着文件走，也不必再拼「表名」当键。未编写过脚本时缺省（不落库）。
     */
    'py-editor:script'?: string;
}

/**
 * DataCollection 类
 * 管理电子表格的数据集合，包括列头、行头、单元格等
 * 提供数据操作、坐标转换和矩形计算等功能
 *
 * @class DataCollection
 * @extends EventEmitter
 * @example
 * ```typescript
 * dataCollection.on(SheetProperty.Zoom, (cell, oldValue, newValue) => {
 *     console.log(`缩放比例从 ${oldValue} 变更为 ${newValue}`);
 * });
 * ```
 */
export class DataCollection extends SheetSearchBase {

    constructor() {
        super();
        this.visibleView = new VisibleView();
        this.colHeaders = new ColHeaders();
        this.rowHeaders = new RowHeaders();
        this._activedCell = 'A1';
        this._selection = `${this._activedCell}:${this._activedCell}`;
        this._anchorSelection = null;
        this._zoom = 1;
        this.undoManager = new UndoManager();
        this.clipboardManager = new ClipboardManager();
        this.dataLoaded = new Promise(resolve => { this._dataLoadedResolve = resolve; });
        this.messageBox = new MessageBox();
        this.matrixSorter = new MatrixSorter(this);
        // 自动保存改由「状态变更事件 → 置脏 → 防抖落库」驱动，
        // 取代原先 `Sheets.draw()` 每帧调用 `syncValuesToDB()`（每帧整表序列化 + 写 IndexedDB）。
        // 之所以自行订阅而非在各变更点散落调用：本类是所有状态的唯一持有者，
        // 无论内部还是外部（如 editCanvas）何时 emit，都能在这里统一捕获。
        // 订阅的事件与 `toJSON()` 快照字段一一对应：值/样式、列宽、行高、网格线、
        // 活动单元格、选区、缩放。`printSetting` 的 setter 不发事件，另行单独置脏。
        const dirtyEvents: DataEvents[] = [
            DataEvents.VALUES_CHANGED,
            DataEvents.COLUMNS_CHANGED,
            DataEvents.ROWS_CHANGED,
            DataEvents.SHOW_GRID_LINES_CHANGED,
            DataEvents.ZOOM_CHANGED,
            DataEvents.ACTIVED_CELL_CHANGED,
            DataEvents.SELECTION_CHANGED,
        ];
        dirtyEvents.forEach((event) => this.on(event, () => this._markDirty()));
    }

    /**
     * 创建工作表名称
     */
    // public async createSheetName() {
    //     await this.db.getAllValues().then((res) => {
    //         const names = res.map((item) => item.sheetName);
    //         let num = 1;
    //         while (num){
    //             if(!names.includes(`Sheet${num}`)){
    //                 this.db.setValue(`Sheet${num}`, JSON.stringify([]));
    //                 this.activedSheetName = `Sheet${num}`;
    //                 break;
    //             }
    //             num++;
    //         }
    //     });
    // }
    

    /**
     * 按存档键装载一张**临时**表（供 Python 脚本 `open()` 其它文件时使用）
     *
     * 与 {@link syncValuesFromDB} 的区别：那个装载的是当前表并会广播事件，
     * 这里只为「读数据 / 写回那一份存档」造一个离屏实例 —— 它不进 UI、不进撤销栈，
     * 也不碰当前表的数据。写回由调用方 `applyScriptWrites(...)` + `syncValuesToDB()` 完成。
     *
     * 刻意不装载图片：脚本能读写的只有值与样式。
     * @param {string} key - 存档键（见 `utils/fileName.toArchiveKey`）
     * @returns {Promise<DataCollection | null>} 临时实例；该键没有存档时 null
     */
    public static async loadScratch(key: string): Promise<DataCollection | null> {
        if (!key) return null;
        const scratch = new DataCollection();
        scratch.activedSheetName = key;
        const raw = await scratch.db.getValue(key);
        // 存档缺省时 IndexDB 返回 '{}'（见 IndexDB#getValue 的兜底），据此判定「没有这个文件」
        if (raw === undefined || raw === null || raw === '{}') return null;
        scratch._suppressValuesUndo = true;
        try {
            scratch._applyLoadedState(await DataCollection.fromSnapshotChunked(raw));
        } finally {
            scratch._refreshBaseline();
            scratch._suppressValuesUndo = false;
        }
        return scratch;
    }


    
    /**
     * 获取指定单元格的值
     * @param {string} cell - 单元格名称
     * @returns {string} 单元格值
     */
    public getCellValue(cell: string): string {
        return this.getCellByName(cell)?.getText() || '';
    }
    
    /**
     * 获取是否正在编辑
     * @returns {boolean} 是否正在编辑
     */
    get isEditting(): boolean {
        return this._isEditting;
    }

    /**
     * 设置是否正在编辑
     * @param {boolean} value - 是否正在编辑
     */
    set isEditting(value: boolean) {
        this.updateProperty(DataEvents.IS_EDITTING_CHANGED, '_isEditting', value);
    }

    /**
     * 获取是否显示网格线
     * @returns {boolean} 是否显示网格线
     */
    get showGridLines(): boolean {
        return this._showGridLines;
    }

    /**
     * 设置是否显示网格线
     * @param {boolean} value - 是否显示网格线
     */
    set showGridLines(value: boolean) {
        this.updateProperty(DataEvents.SHOW_GRID_LINES_CHANGED, '_showGridLines', value);
    }

    /**
     * 获取当前选中的工作表名称
     * @returns {string} 当前选中的工作表名称
     */
    get activedSheetName(): string {
        return this._activedSheetName;
    }

    /**
     * 设置当前选中的工作表名称
     * @param {string} value - 工作表名称
     */
    set activedSheetName(value: string) {
        this.updateProperty(DataEvents.ACTIVED_SHEET_NAME_CHANGED, '_activedSheetName', value);
    }

    /**
     * 获取当前工作表的 Python 脚本（没有编写过时为 null）
     * @returns {string | null} 脚本文本
     */
    get pyScript(): string | null {
        return this._pyScript;
    }

    /**
     * 设置当前工作表的 Python 脚本
     *
     * 脚本属于存档内容：赋值即标记脏，随后由自动保存写进本表的 Value 记录。
     * 与单元格改动一样，未打开文件（无存档键）时只留在内存，不会凭空造出存档。
     * @param {string | null} value - 脚本文本
     */
    set pyScript(value: string | null) {
        const next = value ?? null;
        if (next === this._pyScript) return;
        this._pyScript = next;
        this._markDirty();
    }

    /**
     * 获取当前活动单元格
     * @returns {string} 当前活动单元格名称
     */
    get activedCell(): string {
        return this._activedCell;
    }

    /**
     * 设置当前活动单元格
     * @param {string} value - 单元格名称
     */
    set activedCell(value: string) {
        this.updateProperty(DataEvents.ACTIVED_CELL_CHANGED, '_activedCell', value);
    }

    /**
     * 获取选区
     * @returns {string | null} 选区字符串
     */
    get selection(): string | null {
        return this._selection;
    }

    /**
     * 设置选区
     * @param {string | null} value - 选区字符串
     */
    set selection(value: string | null) {
        this.updateProperty(DataEvents.SELECTION_CHANGED, '_selection', value);
    }
    
    /**
     * 打印设置统一入口（getter/setter）：
     * 工具栏打印设置分区与打印预览面板均读写此处，随 toJSON/fromJSON 序列化持久化
     */
    public get printSetting(): PrinterSetting {
        return this._printSetting;
    }

    /**
     * 设置打印设置
     * @param {PrinterSetting} setting - 打印设置
     */
    public set printSetting(setting: PrinterSetting) {
        this._printSetting = setting;
        // 打印设置随 toJSON 持久化，但不像其它字段那样有对应事件；此处直接置脏，
        // 否则「只改了打印设置、没触发别的变更」时不会落库（原先靠每帧落库兜底）
        this._markDirty();
    }

    /**
     * 获取锚点选择框
     * @returns {string | null} 锚点选择框字符串
     */
    get anchorSelection(): string | null {
        return this._anchorSelection;
    }

    /**
     * 设置锚点选择框
     * @param {string | null} value - 锚点选择框字符串
     */
    set anchorSelection(value: string | null) {
        this.updateProperty(DataEvents.ANCHOR_SELECTION_CHANGED, '_anchorSelection', value);
    }

    /**
     * 获取缩放比例
     * @returns {number} 缩放比例
     */
    get zoom(): number {
        return this._zoom;
    }

    /**
     * 设置缩放比例
     * @param {number} value - 缩放比例
     */
    set zoom(value: number) {
        this.updateProperty(DataEvents.ZOOM_CHANGED, '_zoom', value);
    }

    /**
     * 获取当前工作表所有有内容的单元格数组
     * 注意：返回的是内部数组的活引用，外部只读访问安全；
     * 但原地修改（push/splice/下标赋值等）后必须调用 commitValues() 提交，
     * 才能把变更记入撤销栈并广播 VALUES_CHANGED。
     * @returns {Cell[]} 单元格数组
     */
    get values(): Cell[] {
        return this._values;
    }

    /**
     * 设置当前工作表所有有内容的单元格数组
     * 会以「变化前」的旧数组为基准记录一次撤销操作（装载与撤销回写期间除外），
     * 随后更新 _lastValuesSnapshot 基准并广播 VALUES_CHANGED。
     * 赋值后会先清扫其中既无内容也无样式及功能信息的空单元格（_pruneEmptyCells）。
     * @param {Cell[]} value - 新的单元格数组
     */
    set values(value: Cell[]) {
        const oldValue = this._values;
        this._values = value;
        this._pruneEmptyCells();
        this.invalidateCellIndex();
        // 「变化前」基准不能无条件取旧数组的引用：调用方可能先把旧数组里的 Cell
        // 原地改掉（chars.splice / Object.assign 等），再把同一个数组赋回来，
        // 此时旧引用与当前值逐字段相等，录制阶段会判成「无变化」而**静默丢掉**
        // 这次撤销记录 —— 随后的 Ctrl+Z 实际回退的是更早的一次操作，
        // 表现为「撤销后单元格内容变空、选区还留在原处」。
        // 旧引用已无法表达「变化前」时，退回上一次提交的基准快照
        // （_lastValuesSnapshot 是深拷贝，不会跟着原地修改一起变）。
        const before = this._lastValuesSnapshot
            && this._valuesSnapshotsEqual(oldValue, this._values)
            ? this._lastValuesSnapshot
            : oldValue;
        this._recordValuesChange(before);
    }

    /**
     * 获取当前工作表的内嵌图片（活引用，调用方只读）
     * @returns {ExtractedImage[]} 图片数组
     */
    get images(): ExtractedImage[] {
        return this._images;
    }

    /**
     * 设置当前工作表的内嵌图片
     *
     * 图片与单元格值共用「工作表名」作为存档键，但落在**独立的对象存储**里
     * （见 {@link IndexDB.setImages}）：值走 JSON 字符串、图走结构化克隆。
     *
     * 广播 `VALUES_CHANGED` 而非新增事件：图片是工作表内容的一部分，
     * 画布必须重绘才能显示它们，而该事件正是「重绘」的既有触发点；
     * 新增一个事件则要求所有订阅方（画布、筛选、文件树…）逐个适配，
     * 收益为零而漏改风险不为零。
     * @param {ExtractedImage[]} value - 新的图片数组
     */
    set images(value: ExtractedImage[]) {
        this._images = Array.isArray(value) ? value : [];
        this._imagesDirty = true;
        this.emit(DataEvents.VALUES_CHANGED, this._values);
    }


    /**
     * 使「名 → 单元格」索引失效（下次查找时按需重建）
     *
     * 任何会改变单元格名集合或名→对象映射的操作都必须调用：整体替换 values、
     * 重命名单元格、删除/插入行列、排序等。数组长度变化可由
     * {@link getCellByName} 自动侦测，**等长重命名（如删除行导致 A5→A4）
     * 则必须显式调用**，否则索引会返回过期对象。
     */
    public invalidateCellIndex(): void {
        this._valuesEpoch++;
        this._indexStructurallyInvalid = true;
    }

    /**
     * 标记单元格内容/属性已变化（如着色、格式、文本编辑），但不涉及名集合变更。
     *
     * 与 {@link invalidateCellIndex} 的区别：不设 {@link _indexStructurallyInvalid}，
     * 因此 {@link _ensureCellIndex} 不会触发全量索引重建——着色/格式操作在百万格
     * 大表上不再每次都重建索引（O(N) → O(1)）。
     *
     * 筛选按钮的世代缓存（{@link hasFilterCells}）仍因 _valuesEpoch 变化而失效，
     * 下次 draw 会走一次 getFilterCells → _ensureCellIndex，但后者发现长度未变且
     * 非结构失效，直接跳过重建。
     */
    public markContentChanged(): void {
        this._valuesEpoch++;
    }

    /**
     * 重建稀疏单元格索引与合并单元格缓存
     * 仅在索引失效后的首次查找时触发，稳态下为 O(1)。
     *
     * 同时重建合并区域的**数值矩形**与**列倒排索引**（见 {@link _mergeRects}、
     * {@link _mergeByCol}）：渲染热路径上对合并块的每次查找都靠它们降到近似 O(1)。
     * @private
     */
    private _rebuildCellIndex(): void {
        const index = new Map<string, Cell>();
        const merged: string[] = [];
        const filterCells: Cell[] = [];
        const mergeRects: MergeRect[] = [];
        const mergeByCol = new Map<number, MergeRect[]>();
        const wideMergeRects: MergeRect[] = [];
        const values = this._values;
        for (let i = 0; i < values.length; i++) {
            const cell = values[i];
            index.set(cell.cell, cell);
            if (cell.filter !== undefined) filterCells.push(cell);
            const colspan = cell.colspan || 1;
            const rowspan = cell.rowspan || 1;
            if (colspan === 1 && rowspan === 1) continue;
            const { col, row } = this.getCellColAndRow(cell.cell);
            const endCol = col + colspan - 1;
            const endRow = row + rowspan - 1;
            merged.push(`${cell.cell}:${this.getColName(endCol)}${endRow}`);
            const rect: MergeRect = { anchor: cell.cell, startCol: col, startRow: row, endCol, endRow };
            mergeRects.push(rect);
            if (colspan <= DataCollection.MERGE_COL_INDEX_SPAN_LIMIT) {
                for (let c = col; c <= endCol; c++) {
                    const list = mergeByCol.get(c);
                    if (list === undefined) mergeByCol.set(c, [rect]);
                    else list.push(rect);
                }
            } else {
                wideMergeRects.push(rect);
            }
        }
        this._cellIndex = index;
        this._mergedCells = merged;
        this._mergeRects = mergeRects;
        this._mergeByCol = mergeByCol;
        this._wideMergeRects = wideMergeRects;
        this._filterCells = filterCells;
        this._filterCellsEpoch = this._valuesEpoch;
        this._filterCellsDirty = false;   // 本方法已顺带重算出筛选列表
        this._indexedLength = values.length;
        this._indexEpoch = this._valuesEpoch;
    }

    /**
     * 确保索引与合并/筛选缓存处于最新状态（供缓存类只读接口使用）
     * @private
     */
    protected _ensureCellIndex(): void {
        if (this._values.length !== this._indexedLength || this._indexStructurallyInvalid) {
            this._rebuildCellIndex();
            this._indexStructurallyInvalid = false;
            return;   // _rebuildCellIndex 已同步筛选缓存并清脏
        }
        // 下面两个条件任一成立，都意味着「筛选标记集合可能已经变了」：
        // - `_filterCellsDirty`：显式失效（增删筛选标记既不改长度也不推进世代）；
        // - 世代推进：任何数据变化，含粘贴/格式刷等**只走 commitValues** 的路径
        //   （它们原地改写目标格的 filter，既不置脏也不改长度，只能靠世代兜底）。
        //
        // ⚠️ 历史实现在这里只刷世代标记、**不重扫 `_filterCells`**，
        // 于是缓存会长期停留在旧结论上：新加的按钮建不出来（不显示）、
        // 已删的按钮实例销毁不掉（不消失），且只在下一次整份索引重建时
        // 才突然对上 —— 这就是用户说的「有时候」。
        //
        // 重扫只是一次 values 线性扫描（逐个读 `filter` 字段），
        // 比 `_rebuildCellIndex`（建 Map + 合并倒排）便宜一个数量级；
        // 且仅在世代推进后的首次查询发生，稳态零成本。
        if (this._filterCellsDirty || this._indexEpoch !== this._valuesEpoch) {
            // 内容变了但结构没变：不重建整份索引，只刷世代标记并重扫筛选列表
            this._indexEpoch = this._valuesEpoch;
            this._rebuildFilterCells();
        }
    }


    /**
     * 判定「某列在给定行**以下**是否还有内容」（O(1)）
     *
     * 语义与 `Filter.updateFilterVisibility` 原先的整表遍历**完全一致**：
     * 存在一个「覆盖了该列、且下边缘在 `row` 之下、且文本非空」的单元格。
     * 合并单元格按其覆盖的全部列计入（跨本列的合并块也算本列有内容），
     * 文本判定只看是否存在非空字符（不产生中间字符串）。
     *
     * 「覆盖本列」用 `colspan`（缺省 1，`0` 视为不覆盖任何列，与历史实现一致）、
     * 「下边缘」用 `row + rowspan - 1`，都不含筛选行自身——Excel 的筛选行
     * （表头）不计入「以下的内容」。
     *
     * @param {number} col - 列号（1-based）
     * @param {number} row - 行号（1-based）；筛选按钮所在行，该行及以上不计入
     * @returns {boolean} 该列在 `row` 之下是否还有内容
     */
    public hasContentBelow(col: number, row: number): boolean {
        this._ensureColContentIndex();
        const bottom = this._colContentBottom.get(col);
        return bottom !== undefined && bottom > row;
    }

    /**
     * 按需构建 {@link _colContentBottom}（同一世代只构建一次）
     * @private
     */
    private _ensureColContentIndex(): void {
        if (this._colContentEpoch === this._valuesEpoch
            && this._colContentLength === this._values.length) {
            return;
        }
        const bottom = new Map<number, number>();
        const values = this._values;
        for (let i = 0; i < values.length; i++) {
            const cell = values[i];
            // 先做最便宜的非空文本判定（首字符即命中，通常 1 次比较就返回），
            // 再解析单元格名 —— 大表里这一句就省掉绝大多数格子的名字解析
            if (!DataCollection.hasAnyText(cell)) continue;
            const { col, row } = this.getCellColAndRow(cell.cell);
            const endCol = col + (cell.colspan === undefined ? 1 : cell.colspan) - 1;
            const lastRow = row + (cell.rowspan === undefined ? 1 : cell.rowspan) - 1;
            for (let c = col; c <= endCol; c++) {
                const prev = bottom.get(c);
                if (prev === undefined || lastRow > prev) bottom.set(c, lastRow);
            }
        }
        this._colContentBottom = bottom;
        this._colContentEpoch = this._valuesEpoch;
        this._colContentLength = this._values.length;
    }



    /**
     * 取「有内容的列」集合（只读引用，请勿修改）
     *
     * 与 {@link getRowsWithText} 同源同世代，供添加筛选按钮时逐列 O(1) 判定。
     * @returns {Set<number>} 列号集合（1-based）
     */
    public getColsWithContent(): Set<number> {
        this._ensureFilterIndex();
        return this._colsWithContent;
    }

    /**
     * 取「已用内容范围」：有文本单元格覆盖到的最大行列号（1-based）
     *
     * 是 `matrixValues()` 的**零分配**替代品 —— 筛选下拉只需要「统计到第几行」，
     * 不需要那张稠密矩阵（见 {@link _ensureFilterIndex}）。
     * 口径差异：只看有文本的单元格，仅有样式的单元格不会扩大范围。
     * @returns {{maxRow: number, maxCol: number}} 最大行/列号；无内容时为 0
     */
    public getContentBounds(): { maxRow: number; maxCol: number } {
        this._ensureFilterIndex();
        return { maxRow: this._contentMaxRow, maxCol: this._contentMaxCol };
    }

    /**
     * 取「有文本的行」集合（只读引用，请勿修改）
     * @returns {Set<number>} 行号集合（1-based）
     */
    public getRowsWithText(): Set<number> {
        this._ensureFilterIndex();
        return this._rowsWithText;
    }

    /**
     * 取某位置上「视觉上生效」的单元格 —— 合并区域内返回覆盖该位置的**锚点**单元格（O(1)）
     *
     * 是筛选侧 `buildAnchorMap()` 的替代品。历史实现每次打开下拉 / 每次执行筛选都要
     * 遍历整张表、为**每个被合并覆盖的位置**塞一条 `${row}:${col}` 字符串键
     * （400 万格量级直接上百毫秒）。这里改为「按需 O(1) 查一次」：
     * 先走合并倒排索引 {@link findMergeRect}（O(该列合并块数)），未命中则取本格。
     * @param {number} row - 行号（1-based）
     * @param {number} col - 列号（1-based）
     * @returns {Cell | undefined} 覆盖该位置的单元格；该位置为空时为 undefined
     */
    public getEffectiveCellAt(row: number, col: number): Cell | undefined {
        const rect = this.findMergeRect(col, row);
        if (rect !== null) return this.getCellByName(this.getCellName(rect.startCol, rect.startRow));
        return this.getCellByName(this.getCellName(col, row));
    }

    /**
     * 取某位置上视觉生效的文本（合并区域内为锚点文本）
     * @param {number} row - 行号（1-based）
     * @param {number} col - 列号（1-based）
     * @returns {string} 该位置的文本；为空时返回空字符串
     */
    public getEffectiveTextAt(row: number, col: number): string {
        const cell = this.getEffectiveCellAt(row, col);
        return cell === undefined ? '' : cell.getText();
    }

    /**
     * 按名称取单元格（O(1)，替代 `values.find(v => v.cell === name)`）
     *
     * 失效判定与 {@link _ensureCellIndex} **同源**，保证两类调用看到同一份索引：
     * 1. 结构失效（{@link invalidateCellIndex}）或 `_values` 长度变化 → 重建；
     * 2. 仅内容变化（{@link markContentChanged}）→ 只刷新世代标记，不重建；
     * 3. 命中后再校验 `cell.cell === name`，兜住「等长重命名」这类改了名集合
     *    却没改长度的情形。
     *
     * 稳态（未失效）下是一次 Map 查找，不会退化成每次都重建。
     * @param {string} name - 单元格名称（如 `A1`）
     * @returns {Cell | undefined} 单元格；不存在时 undefined
     */
    public getCellByName(name: string): Cell | undefined {
        // ⚠️ 命中分支**不能**跳过失效判定：_cellIndex 里可能残留「已从 _values 删除、
        //   但仍留在旧索引中」的条目（合并会删掉区域内除锚点外的所有格子）。
        //   历史实现是「命中且 name 相等就直接返回」，于是删除后仍能按名取到
        //   幽灵单元格 —— 补丁撤销据此捕获的 after 快照会把幽灵当成现存格子写回去，
        //   表现为「撤销后合并区还在 / 撤销无效」。
        // 判定与 {@link _ensureCellIndex} 同源：结构失效或长度变化才重建（O(N)，
        //   一次操作只付一次），单纯的内容变化只刷新世代标记（O(1)）。
        if (this._values.length !== this._indexedLength || this._indexStructurallyInvalid) {
            this._rebuildCellIndex();
            this._indexStructurallyInvalid = false;
        } else if (this._indexEpoch !== this._valuesEpoch) {
            this._indexEpoch = this._valuesEpoch;
            this._filterCellsEpoch = this._valuesEpoch;
        }
        const hit = this._cellIndex.get(name);
        return (hit !== undefined && hit.cell === name) ? hit : undefined;
    }

    /**
     * 深拷贝单元格数组，用于撤销快照
     * 复用 Cell.clone()（内部已对 chars 逐个深拷贝），保证快照与原数据完全解耦。
     * @param {Cell[]} cells - 待拷贝的单元格数组
     * @returns {Cell[]} 深拷贝后的新数组
     */
    private cloneValues(cells: Cell[]): Cell[] {
        return cells.map((c: Cell) => c.clone());
    }

    /**
     * 当前是否启用「整表撤销记账」
     *
     * 超过 {@link LARGE_TABLE_UNDO_CELLS} 的大表**关闭**逐次操作的整表撤销记账：
     * 撤销栈的每一条记录都要深拷贝整表（一次操作常需 2~3 份），
     * 几十万格时单次就上百毫秒、并让常驻内存翻倍，代价远大于收益。
     * 大表下：装载、删除行列、排序、格式刷等操作**不可撤销**（属文档级操作），
     * 编辑内容仍会正常保存与渲染，只是不进撤销栈。
     * @returns {boolean} true = 正常记账；false = 大表，跳过记账
     */
    private _undoTrackingEnabled(): boolean {
        return this._values.length < DataCollection.LARGE_TABLE_UNDO_CELLS;
    }

    /**
     * 刷新「变化前」基准快照
     *
     * 大表下**不做深拷贝**，直接置 null（无基准）：此时撤销记账已整体关闭，
     * 基准的唯一用途就是撤销记账，拷贝它纯属浪费内存与时间。
     * @private
     */
    protected _refreshBaseline(): void {
        this._lastValuesSnapshot = this._undoTrackingEnabled() ? this.cloneValues(this._values) : null;
    }

    /**
     * 清扫 _values 中的空单元格
     * 就地移除「既无内容也无样式及功能信息」的 Cell（Cell.isEmpty() 为 true），
     * 使 data.values 在稳态下只保留携带有效信息的单元格。
     * 调用时机（编辑会话进行中除外）：
     * - values setter：装载、撤销/重做回写、排序、筛选等重赋值后；
     * - commitValues()：原地变更提交前（非编辑期间）；
     * - runWithFullStateUndo()：结构操作结束后、捕获「变化后」快照前；
     * - endEditSession()：编辑会话结束时（编辑过程中为保护编辑内撤销栈快照
     *   与光标索引所依赖的单元格布局一致性，不做清扫）。
     * @returns {boolean} 是否移除了至少一个单元格
     */
    private _pruneEmptyCells(): boolean {
        let removed = false;
        for (let i = this._values.length - 1; i >= 0; i--) {
            if (this._values[i].isEmpty()) {
                this._values.splice(i, 1);
                removed = true;
            }
        }
        if (removed) this.invalidateCellIndex();
        return removed;
    }

    /**
     * 比较两个单元格的内容与样式是否等价（字段级深比较）
     *
     * 刻意**不做序列化**：不拼字符串、不分配对象，只做原始值比较。
     * 之所以不用「版本号/脏标记」做判定，是因为单元格存在 `Object.assign`、
     * `chars.splice` 等绕过 setter 的整体复制/原地修改路径，计数器极易漏记，
     * 一旦漏记就会出现「改动没进撤销栈」这类隐性缺陷；字段级比较虽然多花
     * 常数倍时间，但不依赖任何记账约定，行为始终正确。
     * @private
     */
    private _cellsEqual(a: Cell, b: Cell): boolean {
        if (a === b) return true;
        if (a.cell !== b.cell
            || a.colspan !== b.colspan || a.rowspan !== b.rowspan
            || a.borderTopWidth !== b.borderTopWidth || a.borderBottomWidth !== b.borderBottomWidth
            || a.borderLeftWidth !== b.borderLeftWidth || a.borderRightWidth !== b.borderRightWidth
            || a.borderColor !== b.borderColor
            || a.textAlign !== b.textAlign || a.alignItems !== b.alignItems
            || a.backgroundColor !== b.backgroundColor || a.numberFormat !== b.numberFormat
            || a.filter !== b.filter || a.wrap !== b.wrap
            || a.letterSpacing !== b.letterSpacing || a.lineSpacing !== b.lineSpacing
            || a.isEdit !== b.isEdit) {
            return false;
        }
        const charsA = a.chars;
        const charsB = b.chars;
        const lenA = charsA ? charsA.length : 0;
        const lenB = charsB ? charsB.length : 0;
        if (lenA !== lenB) return false;
        for (let i = 0; i < lenA; i++) {
            const x = charsA![i];
            const y = charsB![i];
            if (x === y) continue;
            if (x.char !== y.char || x.fontFamily !== y.fontFamily || x.fontSize !== y.fontSize
                || x.fontWeight !== y.fontWeight || x.fontStyle !== y.fontStyle
                || x.fontColor !== y.fontColor || x.underline !== y.underline
                || x.strikethrough !== y.strikethrough) {
                return false;
            }
        }
        return true;
    }

    /**
     * 比较两份快照是否等价
     *
     * 之所以不用「版本号/脏标记」做判定，是因为单元格存在 `Object.assign`、
     * `chars.splice` 等绕过 setter 的整体复制/原地修改路径，计数器极易漏记，
     * 一旦漏记就会出现「改动没进撤销栈」这类隐性缺陷；字段级比较虽然多花
     * 常数倍时间，但不依赖任何记账约定，行为始终正确。
     * @private
     */
    private _valuesSnapshotsEqual(a: Cell[], b: Cell[]): boolean {
        if (a === b) return true;
        if (a.length !== b.length) return false;
        for (let i = 0; i < a.length; i++) {
            if (!this._cellsEqual(a[i], b[i])) return false;
        }
        return true;
    }

    /**
     * 比较两份行列头是否等价（按条比对位置/尺寸/隐藏标记）
     * @private
     */
    private _headersEqual(a: ColHeader[] | RowHeader[], b: ColHeader[] | RowHeader[]): boolean {
        if (a === b) return true;
        if (!a || !b || a.length !== b.length) return false;
        for (let i = 0; i < a.length; i++) {
            const ha = a[i] as { left?: number, top?: number, width?: number, height?: number, isHidden?: boolean };
            const hb = b[i] as { left?: number, top?: number, width?: number, height?: number, isHidden?: boolean };
            if (ha.left !== hb.left || ha.top !== hb.top
                || ha.width !== hb.width || ha.height !== hb.height
                || ha.isHidden !== hb.isHidden) return false;
        }
        return true;
    }

    /**
     * 求两份整表快照中「首个发生变化」的单元格名
     *
     * 撤销/重做要把活动格与选区带到内容真正变化的位置，而整表快照本身不含
     * 「改的是哪一格」。遍历次序与 `values` 数组一致（行优先），
     * 返回的格子对撤销（before）与重做（after）两个方向都适用：
     * - 两快照都有该格 → 比字段；
     * - 仅 after 有 → 本次新增（撤销时会消失）；
     * - 仅 before 有 → 本次删除（同样需要把视图带过去）。
     * @param {Cell[]} before - 变化前快照
     * @param {Cell[]} after - 变化后快照
     * @returns {string | null} 首个变化单元格名；无差异时为 null
     * @private
     */
    private _findFirstChangedCell(before: Cell[], after: Cell[]): string | null {
        const beforeByName = new Map<string, Cell>();
        for (const cell of before) {
            if (cell && cell.cell) beforeByName.set(cell.cell, cell);
        }
        const afterNames = new Set<string>();
        for (const cell of after) {
            if (!cell || !cell.cell) continue;
            afterNames.add(cell.cell);
            const prev = beforeByName.get(cell.cell);
            if (!prev || !this._cellsEqual(prev, cell)) return cell.cell;
        }
        for (const cell of before) {
            if (cell && cell.cell && !afterNames.has(cell.cell)) return cell.cell;
        }
        return null;
    }

    /**
     * 记录一次 this.values 变更并广播
     * 以 before 为「变化前」基准，当前 _values 的深拷贝为「变化后」after，
     * 满足条件时推入撤销栈，随后刷新 _lastValuesSnapshot 基准并触发 VALUES_CHANGED。
     * - 装载期间（_suppressValuesUndo）：仅刷新基准、不广播、不入栈。
     * - 撤销/重做回写期间（_isApplyingUndo）：仅刷新基准并广播，不入栈。
     * - 大表（见 {@link _undoTrackingEnabled}）/ 无基准（before 为 null）：
     *   仅重建基准并广播，不入栈。
     *
     * 注意：**抑制分支不能省掉基准重建**。`_lastValuesSnapshot` 的消费方是
     * `commitValues()`，它把基准当作「变化前」；若这里只挂引用不拷贝，
     * 后续任意一次原地修改（`values.push/splice`、就地改 cell 字段）都会把
     * 基准与实时数据一起改掉，导致 `_valuesSnapshotsEqual` 恒真 → **静默丢撤销**。
     * 单元格对象在工具栏/编辑画布里被大量就地修改，故深拷贝不可省。
     * @param {Cell[] | null} before - 变更前的快照基准；null 表示无基准
     */
    private _recordValuesChange(before: Cell[] | null): void {
        if (this._suppressValuesUndo) {
            // 补丁撤销已自己捕获「变化前/后」，这里再刷一次整表基准纯属浪费
            // （小表是整表深拷贝，几十万格时单次就几十毫秒）。外层结束时会统一刷新。
            if (this._cellPatchUndoDepth === 0) {
                this._refreshBaseline();
            }
            return;
        }
        // 大表或不记账：不做「变化后」深拷贝，也不入栈
        if (before === null || !this._undoTrackingEnabled()) {
            this._refreshBaseline();
            this.emit(DataEvents.VALUES_CHANGED, this._values);
            return;
        }
        const after = this.cloneValues(this._values);
        if (!this._isApplyingUndo && !this._valuesSnapshotsEqual(before, after)) {
            // values setter 不碰选区，故前后锚点取同一份即可（仍挂上，供撤销后滚动定位）
            const anchor = this._captureAnchor();
            // 同时记下「实际变化的格子」：锚点只说明操作发生时选区在哪，
            // 撤销/重做时活动格与选区要落到内容真正变化处（见 _resolveUndoTargetCell）
            const focusCell = this._findFirstChangedCell(before, after);
            this.undoManager.push({
                type: 'values',
                data: {
                    before: this.cloneValues(before), after: this.cloneValues(after),
                    focusCell: focusCell ?? undefined
                } as ValuesSnapshot,
                timestamp: Date.now(), anchorBefore: anchor, anchorAfter: anchor
            });
        }
        this._lastValuesSnapshot = after;
        this.emit(DataEvents.VALUES_CHANGED, this._values);
    }

    /**
     * 提交一次原地修改
     * 所有对 this.values 数组进行原地变更（push/splice/下标赋值等）的方法，
     * 在变更完成后调用本方法：
     * - 编辑期间（isEditting）：记入编辑内撤销栈 _editUndoStack，不入 undoManager；
     *   整段编辑在 endEditSession() 时统一折叠为一条文档级撤销操作。
     * - 非编辑期间：先清扫空单元格（编辑期间不清扫，避免编辑内撤销栈快照
     *   与光标索引所依赖的单元格布局错位），再以 _lastValuesSnapshot 为「变化前」基准、
     *   当前 _values 为「变化后」记录撤销操作，并广播 VALUES_CHANGED。
     * - 大表：两处都不做深拷贝（编辑内逐字快照会让每次按键都拷贝整表，
     *   几万格以上直接不可交互），只清扫 + 广播。
     */
    public commitValues(): void {
        if (this.isEditting) {
            if (!this._undoTrackingEnabled()) {
                // 大表：整表快照关闭，改用「单格快照」记录编辑内撤销——
                // 编辑会话只触及活动单元格，单格深拷贝的成本与整表大小无关，
                // 因此百万格的表在编辑态改字体/打字依然可以 Ctrl+Z / Ctrl+Y。
                this._commitEditCellState();
                this.emit(DataEvents.VALUES_CHANGED, this._values);
                return;
            }
            const current = this.cloneValues(this._values);
            const last = this._editUndoStack[this._editUndoStack.length - 1];
            // 仅在内容确实变化时入栈，跳过无变化的提交，避免撤销栈产生空操作
            if (!last || !this._valuesSnapshotsEqual(last.values, current)) {
                // 时间窗口合并：500ms 内连续提交合并为栈顶一条，
                // 避免「输入 abcdef 要按 6 次 Ctrl+Z」的过细粒度。
                // 栈底（基准）永远不替换，只在 length > 1 时合并到栈顶。
                const now = performance.now();
                if (last && this._editUndoStack.length > 1
                    && now - this._lastEditCommitTime < 500) {
                    last.values = current;
                    last.cursor = this.cursorAbsolutePosition;
                } else {
                    this._editUndoStack.push({ values: current, cursor: this.cursorAbsolutePosition });
                }
                this._lastEditCommitTime = now;
                this._editRedoStack = [];
            }
            this.emit(DataEvents.VALUES_CHANGED, this._values);
            return;
        }
        // 提交前清扫空单元格：内容与样式信息均已不存在的 Cell 不再保留在 data.values 中
        this._pruneEmptyCells();
        this._recordValuesChange(this._lastValuesSnapshot);
    }

    /**
     * 大表编辑态下记录活动单元格的一次状态变化（单格版撤销栈）
     *
     * 仅当活动单元格相对栈顶确实发生变化时入栈并清空重做栈；
     * 无变化的提交（如重复点同一个格式按钮）不入栈。
     * @private
     */
    private _commitEditCellState(): void {
        const active = this.getCellByName(this.activedCell);
        const snapshot = active ? active.clone() : null;
        const top = this._editCellUndoStack[this._editCellUndoStack.length - 1];
        // 内容未变化：跳过，不入栈（避免空操作）
        if (top && this._sameCellSnapshot(top.cell, snapshot)) return;
        const now = performance.now();
        // 时间窗口合并：500ms 内连续提交合并为栈顶一条，
        // 避免「输入 abcdef 要按 6 次 Ctrl+Z」的过细粒度。
        // 栈底（基准）永远不替换，只在 length > 1 时合并到栈顶。
        if (top && this._editCellUndoStack.length > 1
            && now - this._lastEditCommitTime < 500) {
            top.cell = snapshot;
            top.cursor = this.cursorAbsolutePosition;
        } else {
            this._editCellUndoStack.push({ cell: snapshot, cursor: this.cursorAbsolutePosition });
            this._editCellRedoStack = [];
        }
        this._lastEditCommitTime = now;
    }

    /**
     * 比较两个单格快照是否等价（null 表示单元格不存在）
     *
     * 复用整表比较逻辑，把单个格包装成一格数组即可，避免另写一套字段比较。
     * @param {Cell | null} a - 快照 A
     * @param {Cell | null} b - 快照 B
     * @returns {boolean} 是否等价
     * @private
     */
    private _sameCellSnapshot(a: Cell | null, b: Cell | null): boolean {
        if (a === b) return true;
        if (a === null || b === null) return false;
        return this._valuesSnapshotsEqual([a], [b]);
    }

    /**
     * 把单格快照回写到当前 values（供大表单格撤销/重做使用）
     *
     * 快照为 null 表示该状态下活动单元格不存在 → 从 values 中移除；
     * 当前不存在而快照存在 → 追加回 values。回写后名称索引会在下次查询时惰性重建。
     * @param {Cell | null} snapshot - 目标状态
     * @private
     */
    private _applyEditCellState(snapshot: Cell | null): void {
        // 与 endEditSession 同源：编辑会话只改会话开始时那一格
        const name = this._editCellName ?? this.activedCell;
        const index = this._values.findIndex(c => c.cell === name);
        if (snapshot === null) {
            if (index >= 0) this._values.splice(index, 1);
        } else if (index >= 0) {
            this._values[index] = snapshot.clone();
        } else {
            this._values.push(snapshot.clone());
        }
        // 名称索引立即清空：否则 getCellByName 可能命中被替换格的旧引用
        this._cellIndex.clear();
        this._indexedLength = -1;
    }

    /**
     * 进入编辑会话
     * 捕获编辑前的 values 快照与光标位置作为编辑内撤销栈的基准（栈底 index 0），
     * 并清空编辑内重做栈。应在 isEditting 置为 true 时调用。
     * 大表下不建立基准（栈为空）：编辑内逐字撤销对几十万格的表代价过高，
     * 且 `_undoEditStep` 会因栈深不足直接返回 false，行为一致。
     */
    public beginEditSession(preEdit?: Cell | null): void {
        // 在编辑开始时捕获视图锚点：此刻 activedCell 与 selection 都指向被编辑的
        // 那一个单元格，二者一致；推迟到结束时再采会因鼠标已改动 activedCell 而错位
        // （见 _editSessionAnchor 字段注释）。
        this._editSessionAnchor = this._captureAnchor();
        // 编辑会话的作用对象也在此刻记死（见 _editCellName 字段注释）
        this._editCellName = this._activedCell || null;
        if (this._undoTrackingEnabled()) {
            const values = this.cloneValues(this._values);
            // 【refresh 覆盖输入的基准修正】调用方（editCanvas.startEditting）在
            // refresh 模式下会先清空 chars 再置 isEditting —— 那样本函数拍到的
            // 「编辑前」就已经是空壳，撤销会把原文撤成空白（大表实测）。
            // 故它在清空之前先把原文快照传进来，这里把快照里那一格换回去。
            if (preEdit !== undefined && this._editCellName !== null) {
                const index = values.findIndex((c: Cell) => c.cell === this._editCellName);
                if (preEdit === null) {
                    if (index >= 0) values.splice(index, 1);
                } else if (index >= 0) {
                    values[index] = preEdit.clone();
                } else {
                    values.push(preEdit.clone());
                }
            }
            this._editUndoStack = [{ values, cursor: this.cursorAbsolutePosition }];
        } else {
            // 大表：整表深拷贝关闭，改对活动单元格建立单格基准快照，
            // 编辑内撤销/重做依然可用（见 _editCellUndoStack）。
            // preEdit 显式给出时以它为准（同上：它才是清空前的原文）。
            const active = preEdit !== undefined ? preEdit : this.getCellByName(this.activedCell);
            this._editCellUndoStack = [{
                cell: active ? active.clone() : null,
                cursor: this.cursorAbsolutePosition
            }];
        }
        this._editRedoStack = [];
        this._editCellRedoStack = [];
        // 新编辑会话：重置时间戳与类型，确保会话首次输入作为独立条目入栈。
        this._lastEditCommitTime = 0;
        this._lastEditCommitType = null;
    }

    /**
     * 结束编辑会话
     * 把整段编辑折叠为 undoManager 的一条文档级撤销操作（编辑前→最终值），
     * 并刷新 _lastValuesSnapshot 基准；随后清空编辑内撤销/重做栈。
     * 应在 isEditting 置为 false 时调用。
     */
    public endEditSession(): void {
        // 编辑过程中不做清扫（保护编辑内撤销栈快照与光标索引的布局一致性），
        // 结束时统一清扫，使折叠入撤销栈的「最终值」与快照基准均不含空单元格。
        // 大表（栈为空）同样要清扫，否则空壳 Cell 会随编辑不断累积。
        this._pruneEmptyCells();
        // 用编辑开始时捕获的锚点（见 _editSessionAnchor）：结束时才采会因鼠标
        // 已把 activedCell 改到新点的格子上，与仍是旧选区的 selection 错位，
        // 导致撤销后按键编辑的不是被操作的单元格。兜底取当前锚点应对异常缺省。
        const anchor = this._editSessionAnchor ?? this._captureAnchor();
        if (this._editUndoStack.length > 0) {
            const preEdit = this._editUndoStack[0].values;
            const final = this.cloneValues(this._values);
            if (!this._valuesSnapshotsEqual(preEdit, final)) {
                this.undoManager.push({
                    type: 'values', data: { before: this.cloneValues(preEdit), after: final },
                    timestamp: Date.now(), label: '编辑内容', anchorBefore: anchor, anchorAfter: anchor
                });
            }
        } else if (this._editCellUndoStack.length > 0) {
            // 大表：整表快照关闭，编辑内用「单格快照」记账（见 _commitEditCellState）。
            // 编辑会话结束时必须把它**提升**为文档级撤销记录 ——
            // 否则退出编辑后这段编辑就从撤销栈里消失了，
            // 表现为「大表里改完一个格，Ctrl+Z 毫无反应」（test_sample.xlsx 实测）。
            // 只涉及活动格一个格子，成本 O(1)，与整表大小无关。
            const first = this._editCellUndoStack[0];
            // 目标格必须取**会话开始时**那一格（_editCellName）：结束时 activedCell
            // 可能已被鼠标点走，用它推目标会把补丁记到别的格子上 ——
            // 空白格的编辑于是「撤销无反应」（见 _editCellName 字段注释）。
            const name = this._editCellName ?? (first.cell ? first.cell.cell : this._activedCell);
            const current = this.getCellByName(name);
            const before = [{ cell: name, data: first.cell }];
            const after = [{ cell: name, data: current ? current.clone() : null }];
            if (this._cellPatchesEqual(before, after) === false) {
                this.undoManager.push({
                    type: 'cellPatch', data: { before, after },
                    timestamp: Date.now(), label: '编辑内容', anchorBefore: anchor, anchorAfter: anchor
                });
            }
        }
        this._refreshBaseline();
        this._editUndoStack = [];
        this._editRedoStack = [];
        this._editCellUndoStack = [];
        this._editCellRedoStack = [];
        // 会话结束即作废本会话锚点，避免被下一次编辑会话误用
        this._editSessionAnchor = null;
        this._editCellName = null;
    }

    /**
     * 编辑内撤销一步
     * 弹出 _editUndoStack 栈顶（当前状态）转入重做栈，回写新的栈顶状态，
     * 并记录待恢复的光标索引到 _editRestoreCursor。
     * 栈底（编辑前基准）不会被弹出，因此最多撤销到编辑前的状态。
     * 注意：此处仅替换 _values，不广播 VALUES_CHANGED、不设置光标——
     * 光标定位与广播由 editCanvas 在重测 lines 之后统一处理，
     * 否则 sheets.draw() 会在旧光标下绘制导致光标错乱。
     * @returns {boolean} 是否撤销成功
     */
    private _undoEditStep(): boolean {
        // 大表单格记账模式
        if (this._editUndoStack.length === 0) {
            if (this._editCellUndoStack.length <= 1) return false;
            const current = this._editCellUndoStack.pop()!;
            this._editCellRedoStack.push(current);
            const target = this._editCellUndoStack[this._editCellUndoStack.length - 1];
            this._editRestoreCursor = target.cursor;
            this._applyEditCellState(target.cell);
            // 撤销是明确的操作边界：重置时间戳，使撤销后的下一次输入作为新栈顶条目，
            // 不被合并到撤销前的最后操作（否则会丢失「撤销」这一意图）。
            this._lastEditCommitTime = 0;
            return true;
        }
        if (this._editUndoStack.length <= 1) return false;
        const current = this._editUndoStack.pop()!;
        this._editRedoStack.push(current);
        const target = this._editUndoStack[this._editUndoStack.length - 1];
        this._editRestoreCursor = target.cursor;
        this._applyEditState(target.values);
        this._lastEditCommitTime = 0;
        return true;
    }

    /**
     * 编辑内重做一步
     * 从 _editRedoStack 弹出状态回写，并转回 _editUndoStack。
     * 同样仅替换 _values，光标与广播交由 editCanvas 处理。
     * @returns {boolean} 是否重做成功
     */
    private _redoEditStep(): boolean {
        // 大表单格记账模式
        if (this._editUndoStack.length === 0) {
            if (this._editCellRedoStack.length === 0) return false;
            const state = this._editCellRedoStack.pop()!;
            this._editCellUndoStack.push(state);
            this._editRestoreCursor = state.cursor;
            this._applyEditCellState(state.cell);
            // 重做同样是明确的操作边界：重置时间戳。
            this._lastEditCommitTime = 0;
            return true;
        }
        if (this._editRedoStack.length === 0) return false;
        const state = this._editRedoStack.pop()!;
        this._editUndoStack.push(state);
        this._editRestoreCursor = state.cursor;
        this._applyEditState(state.values);
        this._lastEditCommitTime = 0;
        return true;
    }

    /**
     * 编辑内撤销/重做后应恢复到的光标绝对索引
     * editCanvas 在重测 lines、定位光标时读取本值。
     * @returns {number} 光标绝对索引
     */
    public get editRestoreCursor(): number {
        return this._editRestoreCursor;
    }

    /**
     * 回写编辑内撤销/重做的 values 到 _values
     * 直接替换 _values（绕过 setter，避免触发 undoManager 记录），
     * 快照会被再次深拷贝后赋值，确保栈中保存的快照永不被后续原地修改污染。
     * 不在此处广播 VALUES_CHANGED：编辑内撤销/重做后，由 editCanvas 在
     * 重新测量 lines、定位光标之后再统一广播，避免在旧光标下绘制导致错乱。
     * @param {Cell[]} values - 需要回写的快照
     */
    private _applyEditState(values: Cell[]): void {
        this._values = this.cloneValues(values);
    }

    /**
     * 回写撤销/重做快照到 this.values
     * 经 values setter 走统一流程；期间置 _isApplyingUndo 抑制再次入栈。
     * 快照会被 setter 内部再次深拷贝后赋值，确保栈中保存的快照
     * 永不被后续原地修改污染。
     * @param {Cell[]} snapshot - 需要回写的快照
     */
    private _applyValuesSnapshot(snapshot: Cell[]): void {
        this._isApplyingUndo = true;
        try {
            this.values = this.cloneValues(snapshot);
        } finally {
            this._isApplyingUndo = false;
        }
    }

    /**
     * 是否正在回写撤销/重做快照
     * 供事件处理器在回写期间跳过「派生状态自动变更」（如筛选条件自动清除），
     * 避免恢复过程中的中间态触发误入栈或状态覆盖。
     * @returns {boolean} 是否正在回写
     */
    public get isApplyingUndo(): boolean {
        return this._isApplyingUndo;
    }

    /**
     * 捕获当前「视图锚点」（活动单元格 + 选区）
     *
     * 撤销/重做必须把视口带回操作发生的那块区域：否则在百万行表里撤销第 8000 行
     * 的修改，画面仍停在第 1 行，用户看到的是「什么都没发生」。
     * 锚点刻意**不塞进**任何一种数据快照里，而是挂在 {@link UndoAction} 上 ——
     * 三种快照类型（values / fullState / cellPatch / statePatch）由此共用同一套
     * 捕获与回写逻辑，不必各自加字段。
     * @returns {UndoAnchor} 当前视图锚点
     */
    private _captureAnchor(): UndoAnchor {
        return { activedCell: this._activedCell, selection: this._selection };
    }

    /**
     * 回写视图锚点并把选区卷进可视区
     *
     * 两步都不能省：
     * 1. 恢复 `activedCell` / `selection` —— 走 setter，差量事件驱动工具栏同步；
     * 2. {@link scrollCellIntoView} 到活动格 —— 只恢复选区不滚动的话，
     *    百万行表里撤销后选择框仍在屏幕外，等于没恢复。
     * 滚动目标取**活动格**而非选区左上角：活动格就是用户当时的落点，
     * 且它必然落在选区内，滚到它即保证「选择框有一块在屏内」。
     * @param {UndoAnchor | undefined} anchor - 待恢复的锚点；undefined（旧快照）时不动选区
     * @returns {boolean} 是否发生了滚动
     */
    private _revealAnchor(anchor: UndoAnchor | undefined): boolean {
        if (!anchor || !anchor.activedCell) return false;
        this.activedCell = anchor.activedCell;
        if (anchor.selection) this.selection = anchor.selection;
        const { col, row } = this.getCellColAndRow(anchor.activedCell);
        if (!isFinite(col) || !isFinite(row) || col < 1 || row < 1) return false;
        return this.scrollCellIntoView(col, row);
    }

    /**
     * 捕获当前全量状态快照（values + 行/列表头 + 网格线 + 活动单元格 + 选区 + 筛选条件）
     * 所有字段均为深拷贝，与当前数据完全解耦。
     * @returns {FullStateSnapshot} 全量状态快照
     */
    private _captureFullState(): FullStateSnapshot {
        return {
            values: this.cloneValues(this._values),
            // 图片只做浅拷贝：对象本身是只读替换语义，深拷一份几十 MB 不值得
            images: this._images.slice(),
            rowHeaders: this.rowHeaders.rowHeaders.map(rh => ({ ...rh })),
            colHeaders: this.colHeaders.colHeaders.map(ch => ({ ...ch })),
            showGridLines: this._showGridLines,
            activedCell: this._activedCell,
            selection: this._selection,
            filters: this._serializeFilterConditions(),
        };
    }

    /**
     * 比较两份快照的图片集合
     *
     * 只比「长度 + 逐项引用」：图片对象在项目内是只读替换语义，
     * 引用相同即内容相同；引用不同即换过图（如打开文件）—— 这正是需要记入撤销的变更。
     * @param a - 快照 A 的图片数组（旧快照可能没有该字段）
     * @param b - 快照 B 的图片数组
     * @returns {boolean} 是否相同
     * @private
     */
    private static _imagesEqual(a: unknown[] | undefined, b: unknown[] | undefined): boolean {
        if (a === b) return true;
        const left = a ?? [];
        const right = b ?? [];
        if (left.length !== right.length) return false;
        for (let i = 0; i < left.length; i++) {
            if (left[i] !== right[i]) return false;
        }
        return true;
    }

    /**
     * 将筛选条件注册表序列化为可 JSON 比较的普通结构（按筛选操作顺序排列）
     * @returns {FullStateSnapshot['filters']} 筛选条件数组
     * @private
     */
    private _serializeFilterConditions(): FullStateSnapshot['filters'] {
        return this._filterOrder.map(col => {
            const cond = this._filterConditions.get(col)!;
            return { col, anchorRow: cond.anchorRow, keep: [...cond.keep] };
        });
    }

    /**
     * 从快照恢复筛选条件注册表（撤销/重做回写）
     * 按 filters 数组顺序重建 Map 与操作顺序表；快照无 filters 字段（旧格式）时清空注册表。
     * @param {FullStateSnapshot['filters']} filters - 快照中的筛选条件数组
     * @private
     */
    private _restoreFilterConditions(filters: FullStateSnapshot['filters']): void {
        this._filterConditions = new Map();
        this._filterOrder = [];
        if (!filters) return;
        for (const f of filters) {
            this._filterConditions.set(f.col, { anchorRow: f.anchorRow, keep: new Set(f.keep) });
            this._filterOrder.push(f.col);
        }
    }

    /**
     * 比较两份全量状态快照是否等价
     *
     * 逐字段比对（values 做字段级深比较、行列头逐条比对、其余为标量），
     * **不再 `JSON.stringify` 整个快照**——那会把整张表（含每个字符的两层 JSON）
     * 序列化两遍，大表上一次结构操作就能产生数百 MB 的临时字符串，
     * 是「打开大 Excel 页面崩溃」的直接推手之一。
     * @param {FullStateSnapshot} a - 快照 A
     * @param {FullStateSnapshot} b - 快照 B
     * @returns {boolean} 是否等价
     */
    private _fullStateSnapshotsEqual(a: FullStateSnapshot, b: FullStateSnapshot): boolean {
        if (a === b) return true;
        if (!this._valuesSnapshotsEqual(a.values as Cell[], b.values as Cell[])) return false;
        if (!DataCollection._imagesEqual(a.images, b.images)) return false;
        if (a.showGridLines !== b.showGridLines) return false;
        if (a.activedCell !== b.activedCell) return false;
        if (a.selection !== b.selection) return false;
        if (!this._headersEqual(a.rowHeaders, b.rowHeaders)) return false;
        if (!this._headersEqual(a.colHeaders, b.colHeaders)) return false;
        return JSON.stringify(a.filters) === JSON.stringify(b.filters);
    }

    /**
     * 以「全量状态快照」为单位执行一次可撤销操作
     *
     * 专为绕过 values setter 的结构性操作设计（右键菜单的删除行列、隐藏行列、
     * 清除内容、设置格式、调整行高、切换网格线、筛选操作等）：
     * 1. 捕获操作前全量快照；
     * 2. 抑制 values setter 的自动快照记录（避免 deleteRows/deleteCols 内部
     *    的 values 重赋值产生第二条不完整的撤销记录）；
     * 3. 执行 action（即使提前 return 也会经 finally 恢复抑制标记）；
     * 4. 捕获操作后快照，与操作前不等时推入 undoManager（type: 'fullState'）。
     *
     * 嵌套语义：action 内部再次调用本方法（如取消隐藏行触发筛选条件自动清除）
     * 时，**嵌套层不捕获任何快照**（其变更已包含在最外层快照的 after 中），
     * 整个复合操作折叠为一步撤销——同时也避免了对同一份数据重复深拷贝。
     *
     * 大表语义（见 {@link LARGE_TABLE_UNDO_CELLS}）：整表快照这条路被关掉，改由
     * `scope` 决定是否用**影响面补丁**记账（见 {@link StatePatchSnapshot}）：
     * - `'auto'`（默认）：大表下完全不记账 —— 与历史行为一致，
     *   用于影响面会溢出选区的操作（插入/删除整行整列会把下方/右侧所有格子重映射，
     *   只还原选区那一块会留下自相矛盾的数据）。
     * - `'selection'`：以**当前选区**为影响面记补丁，适用于「改的只有选区里的格子」
     *   的操作（清除内容、设置格式、剪切…）。影响面超过
     *   {@link LARGE_TABLE_PATCH_CELLS} 时放弃记账（退化为文档级操作），绝不硬扛。
     * - `'headers'`：只记行/列表头 + 网格线 + 图片 + 筛选，不记格子 ——
     *   适用于**根本不动 values** 的操作（隐藏/取消隐藏行列、调整行高列宽、切换网格线）。
     * - `'paste'`：以**粘贴落点**为影响面（见 {@link getPasteTargetRect}）。
     *   粘贴不能沿用 `'selection'`：用户常选中 1 格却粘贴一大块剪贴板，
     *   影响面会溢出选区，只还原选区那一块会留下自相矛盾的数据。
     * - `'shiftLeft'` / `'shiftUp'`：以**平移波及带**为影响面 ——
     *   「删除并左移/上移」「插入并右移/下移」会把选区右侧/下方的可见内容
     *   整体搬动（溢出端丢弃），只记选区那一块会留下自相矛盾的数据；
     *   影响面取「选区起点 → 内容边界」的矩形（内容边界用 O(1) 的
     *   {@link getContentBounds}，多记的格子前后相等、回写时是无害的空转），
     *   超 {@link LARGE_TABLE_PATCH_CELLS} 时放弃记账。
     * - `'filters'`：只记**挂着筛选按钮的格子**（稀疏，通常只有几列一个）+
     *   行/列表头 + 筛选条件 —— 「清除筛选按钮」改的正是这三样
     *   （cell.filter 标记、条件注册表、被筛选隐藏的行），格子影响面稀疏
     *   且无法用矩形表达，故单独设档。
     *
     * 前后都要判断阈值：装载第二个大文件时「操作前」就已经是大表，
     * 只判「操作后」仍会在捕获 before 时白拷一份整表。
     * @param {string} label - 操作标签（用于调试与撤销展示）
     * @param {() => void} action - 实际的变更操作
     * @param {'auto' | 'selection' | 'headers' | 'paste' | 'shiftLeft' | 'shiftUp' | 'filters'} [scope='auto'] - 大表下的记账策略（小表恒为整表快照）
     * @param {{ startCol: number; endCol: number; startRow: number; endRow: number }} [rect] - 选区矩形（`'shiftLeft'`/`'shiftUp'` 档必填，其余档忽略）
     * @example
     * data.runWithFullStateUndo('删除行', () => data.deleteRows(2, 3));
     * @example
     * // 大表下也要能撤销「清除内容」：影响面就是选区
     * data.runWithFullStateUndo('清除内容', () => data.clearContents(...), 'selection');
     */
    public runWithFullStateUndo(
        label: string,
        action: () => void,
        scope: 'auto' | 'selection' | 'headers' | 'paste' | 'brush' | 'shiftLeft' | 'shiftUp' | 'filters' = 'auto',
        rect?: { startCol: number; endCol: number; startRow: number; endRow: number }
    ): void {
        const isOutermost = this._fullStateUndoDepth === 0;
        const anchorBefore = isOutermost ? this._captureAnchor() : null;
        // 大表下按 scope 决定「怎么记」；小表恒为整表快照
        const sparse
            = (isOutermost && !this._undoTrackingEnabled())
                ? (scope === 'selection' ? this._selectionRectForPatch()
                    : scope === 'paste' ? this._pasteRectForPatch()
                        : scope === 'brush' ? this._brushRectForPatch()
                            : scope === 'shiftLeft' ? this._shiftLeftRectForPatch(rect)
                                : scope === 'shiftUp' ? this._shiftUpRectForPatch(rect)
                                    : null)
                : null;
        // 仅最外层需要「变化前」快照；嵌套层复用外层的捕获结果
        const before = (isOutermost && this._undoTrackingEnabled()) ? this._captureFullState() : null;
        // 'filters' 档：影响面是稀疏的筛选按钮格子（非矩形），单独捕获
        const sparseCells = (isOutermost && !this._undoTrackingEnabled() && scope === 'filters')
            ? this._captureFilterCellsPatch()
            : undefined;
        const sparseBefore = (isOutermost && !this._undoTrackingEnabled()
            && (scope === 'selection' || scope === 'headers' || scope === 'paste' || scope === 'brush'
                || scope === 'shiftLeft' || scope === 'shiftUp' || scope === 'filters'))
            ? this._captureStatePatch(sparse, scope === 'headers', sparseCells)
            : null;
        const prevSuppress = this._suppressValuesUndo;
        this._suppressValuesUndo = true;
        this._fullStateUndoDepth++;
        try {
            action();
        } finally {
            // 操作结束后清扫空单元格（如清除内容后残留的空 Cell），
            // 使「变化后」快照与稳态数据均不残留空单元格
            this._pruneEmptyCells();
            this._fullStateUndoDepth--;
            this._suppressValuesUndo = prevSuppress;
            // 着色/格式等属性操作不改名集合，只标记内容变化（O(1)），
            // 避免每次操作都触发百万格索引重建。
            // 结构操作（insertRows/insertCols 等重命名）自己调了 invalidateCellIndex。
            // _pruneEmptyCells 删空格时也会调 invalidateCellIndex（改变了名集合）。
            this.markContentChanged();
            if (!this._isApplyingUndo) {
                if (before !== null && this._undoTrackingEnabled()) {
                    // 仅最外层推入撤销记录；撤销/重做回写期间（_isApplyingUndo）不得入栈；
                    // 「变化后」超过阈值时同样跳过（如把空表装载成几十万格）
                    const after = this._captureFullState();
                    if (!this._fullStateSnapshotsEqual(before, after)) {
                        this.undoManager.push({
                            type: 'fullState', data: { before, after }, timestamp: Date.now(), label,
                            anchorBefore: anchorBefore || undefined, anchorAfter: this._captureAnchor()
                        });
                    }
                } else if (sparseBefore !== null) {
                    const after = this._captureStatePatch(sparse, scope === 'headers');
                    if (!this._statePatchSnapshotsEqual(sparseBefore, after)) {
                        this.undoManager.push({
                            type: 'statePatch', data: { before: sparseBefore, after }, timestamp: Date.now(), label,
                            anchorBefore: anchorBefore || undefined, anchorAfter: this._captureAnchor()
                        });
                    }
                }
            }
        }
    }

    /**
     * 求「按影响面记账」的捕获矩形（大表专用）
     *
     * 影响面取**当前选区**：右键菜单 / 工具栏的选区类操作改的正是这块。
     * 选区无法解析（无选区）或面积超过 {@link LARGE_TABLE_PATCH_CELLS} 时返回 null，
     * 调用方据此放弃记账 —— 宁可这一次不可撤销，也不能为它做一次整表深拷贝。
     * @returns 捕获矩形；null = 放弃记账
     * @private
     */
    private _selectionRectForPatch(): { startCol: number; endCol: number; startRow: number; endRow: number } | null {
        return this._capPatchRect(this._selection ? this._parseRect(this._selection) : null);
    }

    /**
     * 求「粘贴」的捕获矩形（大表专用）
     *
     * 取 {@link getPasteTargetRect} 的结果并套 {@link LARGE_TABLE_PATCH_CELLS} 上限。
     * 剪贴板整块超限（复制了几十万格再粘贴）时放弃记账，退化为文档级操作。
     * @returns 捕获矩形；null = 放弃记账
     * @private
     */
    private _pasteRectForPatch(): { startCol: number; endCol: number; startRow: number; endRow: number } | null {
        return this._capPatchRect(this.getPasteTargetRect());
    }

    /**
     * 求「格式刷」的捕获矩形（大表专用）
     *
     * 影响面 = **目标选区 ∪ 由源跨度撑开的块**。必须撑开：源是合并单元格（3×3）而
     * 目标只选了 1 格时，落点会变成一个 3×3 合并块，块内被覆盖的格子要被删除 ——
     * 那些格子在选区**之外**，只用 'selection' 档记账的话大表下撤销/重做撤不回它们
     * （表现为「撤一次只撤掉一半」）。
     * @returns 捕获矩形；null = 放弃记账
     * @private
     */
    private _brushRectForPatch(): { startCol: number; endCol: number; startRow: number; endRow: number } | null {
        return this._capPatchRect(this.getBrushTargetRect());
    }

    /** 给影响面矩形套 {@link LARGE_TABLE_PATCH_CELLS} 上限；超限返回 null */
    private _capPatchRect(
        rect: { startCol: number; endCol: number; startRow: number; endRow: number } | null
    ): { startCol: number; endCol: number; startRow: number; endRow: number } | null {
        if (!rect) return null;
        const area = (rect.endCol - rect.startCol + 1) * (rect.endRow - rect.startRow + 1);
        if (area > DataCollection.LARGE_TABLE_PATCH_CELLS) return null;
        return rect;
    }

    /**
     * 求「水平平移」（删除并左移 / 插入并右移）的捕获矩形（大表专用）
     *
     * 这类操作把选区**右侧**的可见内容整体向左/向右搬动（溢出最右可见列的丢弃），
     * 影响面 = 选区各行 × 从选区起始列到内容右边界。内容右边界取 O(1) 的
     * {@link getContentBounds().maxCol} 与选区末列的较大者 —— 可能比逐行真边界宽
     * （其它行内容更靠右时），多记的格子前后相等，回写时是无害的空转；
     * 超 {@link LARGE_TABLE_PATCH_CELLS} 时放弃记账。
     * @param sel - 选区矩形；缺失（无选区）时放弃记账
     * @returns 捕获矩形；null = 放弃记账
     * @private
     */
    private _shiftLeftRectForPatch(
        sel?: { startCol: number; endCol: number; startRow: number; endRow: number }
    ): { startCol: number; endCol: number; startRow: number; endRow: number } | null {
        if (!sel) return null;
        const maxCol = Math.max(sel.endCol, this.getContentBounds().maxCol);
        return this._capPatchRect({
            startCol: sel.startCol, endCol: maxCol, startRow: sel.startRow, endRow: sel.endRow
        });
    }

    /**
     * 求「垂直平移」（删除并上移 / 插入并下移）的捕获矩形（大表专用）
     *
     * 与 {@link _shiftLeftRectForPatch} 对称：操作把选区**下方**的可见内容整体
     * 上/下搬动，影响面 = 选区各列 × 从选区起始行到内容下边界。
     * @param sel - 选区矩形；缺失（无选区）时放弃记账
     * @returns 捕获矩形；null = 放弃记账
     * @private
     */
    private _shiftUpRectForPatch(
        sel?: { startCol: number; endCol: number; startRow: number; endRow: number }
    ): { startCol: number; endCol: number; startRow: number; endRow: number } | null {
        if (!sel) return null;
        const maxRow = Math.max(sel.endRow, this.getContentBounds().maxRow);
        return this._capPatchRect({
            startCol: sel.startCol, endCol: sel.endCol, startRow: sel.startRow, endRow: maxRow
        });
    }

    /**
     * 捕获「挂着筛选按钮」的格子补丁（'filters' 档，大表专用）
     *
     * 筛选按钮标记（`cell.filter`）散落在各列表头格上，无法用矩形表达；
     * 数量 = 已设筛选的列数（通常个位数到几十个），逐个克隆的代价可忽略。
     * 遍历整表 values 是 O(已用格数) —— 与「清除筛选按钮」操作自身的
     * 遍历同量级，不构成额外瓶颈。
     * @returns {CellPatchEntry[]} 筛选按钮格子的克隆补丁（可能为空数组）
     * @private
     */
    private _captureFilterCellsPatch(): CellPatchEntry[] {
        const out: CellPatchEntry[] = [];
        for (const cell of this._values) {
            if (cell.filter !== undefined) out.push({ cell: cell.cell, data: cell.clone() });
        }
        return out;
    }

    /**
     * 求粘贴落点矩形（从活动单元格起、按剪贴板的可见行列数铺开）
     *
     * 与 {@link pasteSelection} 的落点算法同源（同样跳过隐藏行列、同样在表尾截断），
     * 否则影响面算错就会出现「撤销只还原一半」。
     * @returns 粘贴落点矩形；无剪贴板或活动格非法时为 null
     */
    public getPasteTargetRect(): { startCol: number; endCol: number; startRow: number; endRow: number } | null {
        const clipboard = this.clipboardManager.getClipboard();
        if (!clipboard) return null;
        const { col: activedCol, row: activedRow } = this.getCellColAndRow(this.activedCell);
        if (!isFinite(activedCol) || !isFinite(activedRow) || activedCol < 1 || activedRow < 1) return null;
        // 纯文本剪贴板：只写活动格
        if (clipboard.type !== 'cell' || !clipboard.cells || !clipboard.sourceRange) {
            return { startCol: activedCol, endCol: activedCol, startRow: activedRow, endRow: activedRow };
        }
        const [start, end] = clipboard.sourceRange.split(':');
        const { col: startCol, row: startRow } = this.getCellColAndRow(start);
        const { col: endCol, row: endRow } = this.getCellColAndRow(end);
        if (!isFinite(startCol) || !isFinite(endCol) || !isFinite(startRow) || !isFinite(endRow)) return null;
        const colOffsets = clipboard.visibleColOffsets
            ?? Array.from({ length: endCol - startCol + 1 }, (_, i) => i);
        const rowOffsets = clipboard.visibleRowOffsets
            ?? Array.from({ length: endRow - startRow + 1 }, (_, i) => i);
        const targetColOffsets = this._collectVisibleOffsets(activedCol, Infinity, colOffsets.length, true);
        const targetRowOffsets = this._collectVisibleOffsets(activedRow, Infinity, rowOffsets.length, false);
        if (targetColOffsets.length === 0 || targetRowOffsets.length === 0) return null;
        return {
            startCol: activedCol,
            endCol: activedCol + targetColOffsets[targetColOffsets.length - 1],
            startRow: activedRow,
            endRow: activedRow + targetRowOffsets[targetRowOffsets.length - 1],
        };
    }

    /**
     * 求「格式刷落刷」的影响面矩形：目标选区按源的最大跨度向右下撑开
     *
     * 撑开量取剪贴板里**最大**的 colspan/rowspan，而不是逐格精确计算 ——
     * 精确值要为每个目标位置重算一遍源对应关系（回绕 + 隐藏行列），而偏大的
     * 影响面只是多克隆几个前后完全相同的空格子（无害空转），代价远低于漏记
     * （漏记 = 撤销后数据自相矛盾）。源里没有任何跨度时原样返回选区矩形。
     * @returns 影响面矩形；无选区时返回 null
     */
    public getBrushTargetRect(): { startCol: number; endCol: number; startRow: number; endRow: number } | null {
        const sel = this._selection ? this._parseRect(this._selection) : null;
        if (!sel) return null;
        const clipboard = this.clipboardManager.getClipboard();
        if (!clipboard || clipboard.type !== 'cell' || !clipboard.cells) return sel;
        let maxColspan = 1;
        let maxRowspan = 1;
        for (const cell of clipboard.cells) {
            if (typeof cell.colspan === 'number' && cell.colspan > maxColspan) maxColspan = cell.colspan;
            if (typeof cell.rowspan === 'number' && cell.rowspan > maxRowspan) maxRowspan = cell.rowspan;
        }
        if (maxColspan === 1 && maxRowspan === 1) return sel;
        return {
            startCol: sel.startCol,
            startRow: sel.startRow,
            endCol: sel.endCol + (maxColspan - 1),
            endRow: sel.endRow + (maxRowspan - 1),
        };
    }

    /**
     * 捕获稀疏状态快照（大表专用）
     *
     * 只遍历一次 `_values`（只读、不克隆不在影响面内的格子），
     * 命中矩形者才 `clone()` —— 成本与**影响面格数**成正比，与整表大小几乎无关。
     * 行/列表头、网格线、图片、筛选条件全量浅拷贝：一万多行表头也就几毫秒，
     * 而它们是「隐藏行列 / 行高列宽」这类操作**唯一**会变的东西。
     * @param rect - 影响面矩形；null 表示放弃格子记账（超上限）
     * @param headersOnly - true = 连格子都不记（'headers' 档）
     * @param presetCells - 预先捕获好的稀疏格子补丁（'filters' 档）；提供时优先于 rect 使用
     * @returns {StatePatchSnapshot | null} 快照；rect 为 null 且非 headersOnly 且无 presetCells 时为 null
     * @private
     */
    private _captureStatePatch(
        rect: { startCol: number; endCol: number; startRow: number; endRow: number } | null,
        headersOnly: boolean,
        presetCells?: CellPatchEntry[]
    ): StatePatchSnapshot | null {
        if (!headersOnly && rect === null && presetCells === undefined) return null;
        // 直接复用矩形版补丁捕获：它按**名字**逐格取，而不是遍历 _values 挑命中的 ——
        // 后者拿不到「操作后被删掉的格子」（清除内容会先清空再被剪掉），
        // 于是 after 侧会少一条记录，重做时无从删除，表现为「重做没反应」。
        // 按名取则把「此刻不存在」记成 null，补丁能表达「存在 ↔ 不存在」两个方向。
        const cells: CellPatchEntry[] = (presetCells !== undefined)
            ? presetCells
            : (!headersOnly && rect !== null)
                ? this._captureCellPatch(rect.startCol, rect.endCol, rect.startRow, rect.endRow)
                : [];
        return {
            cells,
            rowHeaders: this.rowHeaders.rowHeaders.map(rh => ({ ...rh })),
            colHeaders: this.colHeaders.colHeaders.map(ch => ({ ...ch })),
            showGridLines: this._showGridLines,
            images: this._images.slice(),
            filters: this._serializeFilterConditions(),
        };
    }

    /**
     * 比较两份稀疏状态快照是否等价
     * @private
     */
    private _statePatchSnapshotsEqual(a: StatePatchSnapshot, b: StatePatchSnapshot): boolean {
        if (a === b) return true;
        if (!this._cellPatchesEqual(a.cells, b.cells)) return false;
        if (a.showGridLines !== b.showGridLines) return false;
        if (!DataCollection._imagesEqual(a.images, b.images)) return false;
        if (!DataCollection._headersEqual(a.rowHeaders, b.rowHeaders)) return false;
        if (!DataCollection._headersEqual(a.colHeaders, b.colHeaders)) return false;
        if (!DataCollection._filtersEqual(a.filters, b.filters)) return false;
        return true;
    }

    /** 表头数组逐项浅比较（top/left/height/width/isHidden 等标量） */
    private static _headersEqual(a: any[], b: any[]): boolean {
        if (a === b) return true;
        if (!a || !b || a.length !== b.length) return false;
        for (let i = 0; i < a.length; i++) {
            const x = a[i];
            const y = b[i];
            if (x === y) continue;
            const kx = Object.keys(x);
            if (kx.length !== Object.keys(y).length) return false;
            for (let k = 0; k < kx.length; k++) {
                if (x[kx[k]] !== y[kx[k]]) return false;
            }
        }
        return true;
    }

    /** 筛选条件列表比较 */
    private static _filtersEqual(
        a: Array<{ col: number; anchorRow: number; keep: string[] }> | undefined,
        b: Array<{ col: number; anchorRow: number; keep: string[] }> | undefined
    ): boolean {
        if (a === b) return true;
        const left = a ?? [];
        const right = b ?? [];
        if (left.length !== right.length) return false;
        for (let i = 0; i < left.length; i++) {
            if (left[i].col !== right[i].col || left[i].anchorRow !== right[i].anchorRow) return false;
            if (left[i].keep.length !== right[i].keep.length) return false;
            for (let k = 0; k < left[i].keep.length; k++) {
                if (left[i].keep[k] !== right[i].keep[k]) return false;
            }
        }
        return true;
    }

    /**
     * 回写稀疏状态快照（撤销/重做恢复）
     *
     * 顺序刻意**先表头后格子**：行/列表头决定几何（top/left/隐藏），
     * 格子补丁只按名字替换内容，二者互不依赖，但表头先就位可以让随后的
     * 滚动定位与重绘拿到正确的坐标。
     * @param {StatePatchSnapshot} snapshot - 待恢复的快照
     * @private
     */
    private _applyStatePatchSnapshot(snapshot: StatePatchSnapshot): void {
        this._isApplyingUndo = true;
        try {
            this.rowHeaders.rowHeaders = snapshot.rowHeaders.map(rh => ({ ...rh }));
            this.rowHeaders.recalcAllRowHeight();
            this.colHeaders.colHeaders = snapshot.colHeaders.map(ch => ({ ...ch }));
            this.colHeaders.recalcAllColWidth();
            this.showGridLines = snapshot.showGridLines;
            this._restoreFilterConditions(snapshot.filters);
            if (Array.isArray(snapshot.images)) {
                this._images = snapshot.images.slice();
                this._imagesDirty = true;
            }
            if (snapshot.cells.length > 0) {
                this._applyCellPatch(snapshot.cells);
            }
            this._refreshBaseline();
            this.emit(DataEvents.VALUES_CHANGED, this._values);
        } finally {
            this._isApplyingUndo = false;
        }
    }

    /**
     * 捕获矩形区域内每个单元格的当前状态（局部补丁快照的「一半」）
     *
     * 顺序固定为「行优先」，回写时按名字匹配，因此与 `_values` 的实际排列无关。
     * **必须在改动前调用**：此时单元格索引处于稳态，`getCellByName` 是 O(1)；
     * 一旦在「边 push 边查」的循环里调用，每次 push 都会让索引世代失效，
     * 下一次查找就要重建整份索引（O(N)）——这正是取消合并 1000 格要 4.4 秒的原因。
     * @param {number} minCol - 起始列（1-based）
     * @param {number} maxCol - 结束列（1-based，含）
     * @param {number} minRow - 起始行（1-based）
     * @param {number} maxRow - 结束行（1-based，含）
     * @returns {CellPatchEntry[]} 区域内每个格子的克隆（不存在为 null）
     * @private
     */
    private _captureCellPatch(minCol: number, maxCol: number, minRow: number, maxRow: number): CellPatchEntry[] {
        const out: CellPatchEntry[] = [];
        for (let row = minRow; row <= maxRow; row++) {
            for (let col = minCol; col <= maxCol; col++) {
                const name = this.getCellName(col, row);
                const cell = this.getCellByName(name);
                out.push({ cell: name, data: cell ? cell.clone() : null });
            }
        }
        return out;
    }

    /**
     * 比较两份局部补丁是否等价
     *
     * 与 {@link _cellsEqual} 同源的字段级深比较，但多处理「null ↔ 非 null」
     * （格子被合并吞掉 / 被取消合并补回来）与「名字不同」三种情形。
     * @private
     */
    private _cellPatchesEqual(a: CellPatchEntry[], b: CellPatchEntry[]): boolean {
        if (a === b) return true;
        if (!a || !b || a.length !== b.length) return false;
        for (let i = 0; i < a.length; i++) {
            if (a[i].cell !== b[i].cell) return false;
            const ca = a[i].data;
            const cb = b[i].data;
            if (ca === cb) continue;
            if (!ca || !cb) return false;
            if (!this._cellsEqual(ca, cb)) return false;
        }
        return true;
    }

    /**
     * 以「局部单元格补丁」为单位执行一次可撤销操作
     *
     * 与 {@link runWithFullStateUndo} 的区别：只捕获**矩形区域内**的格子，
     * 代价与区域大小成正比、与整表大小无关，因此不受
     * {@link LARGE_TABLE_UNDO_CELLS} 阈值限制 —— 百万格的表里合并 50 个单元格
     * 依然可以 Ctrl+Z / Ctrl+Y（整表快照在这种规模下会被直接关掉）。
     *
     * 适用面：影响范围**可预先确定**的结构性操作（合并/取消合并、按行/列合并…）。
     * 范围不可预知的操作（删除行列会整体平移）仍应走全量快照。
     *
     * 流程与全量版一致：捕获 before → 抑制 values setter 自动记账 → 执行 action
     * → 清扫空单元格 → 捕获 after → 不等则入栈 → 广播。
     * @param {string} label - 操作标签
     * @param {{ minCol: number; maxCol: number; minRow: number; maxRow: number }} rect - 影响区域（1-based，闭区间）
     * @param {() => void} action - 实际的变更操作
     */
    public runWithCellPatchUndo(
        label: string,
        rect: { minCol: number; maxCol: number; minRow: number; maxRow: number },
        action: () => void
    ): void {
        const { minCol, maxCol, minRow, maxRow } = rect;
        const anchorBefore = this._captureAnchor();
        const before = this._captureCellPatch(minCol, maxCol, minRow, maxRow);
        const prevSuppress = this._suppressValuesUndo;
        this._suppressValuesUndo = true;
        this._cellPatchUndoDepth++;
        try {
            action();
        } finally {
            this._cellPatchUndoDepth--;
            this._suppressValuesUndo = prevSuppress;
            this._pruneEmptyCells();
            this.markContentChanged();
            if (!this._isApplyingUndo) {
                const after = this._captureCellPatch(minCol, maxCol, minRow, maxRow);
                if (!this._cellPatchesEqual(before, after)) {
                    this.undoManager.push({
                        type: 'cellPatch', data: { before, after }, timestamp: Date.now(), label,
                        anchorBefore, anchorAfter: this._captureAnchor()
                    });
                }
            }
            // 包裹期间跳过了基准刷新，这里补一次，避免后续 commitValues 拿过期基准
            this._refreshBaseline();
            this.emit(DataEvents.VALUES_CHANGED, this._values);
        }
    }

    /**
     * 以「一批指定单元格」为补丁单位执行一次可撤销操作
     *
     * 与 {@link runWithCellPatchUndo} 的关系：两者都是「局部补丁」，代价都与影响面
     * 成正比、都不受 {@link LARGE_TABLE_UNDO_CELLS} 限制。区别在**影响面怎么表达**：
     * - 矩形版按 `minCol..maxCol × minRow..maxRow` 逐格遍历，适合「影响面就是一块矩形」
     *   的结构性操作（合并 / 取消合并），并且会把区域内**不存在**的格子也记成 `null`，
     *   撤销时据此把操作中新增的格子删掉；
     * - 本方法按给定数组逐个捕获，适合**稀疏**的影响面。查找替换正是如此：整表范围内
     *   可能只有 3 处命中，矩形版却会为 10 万行 × 100 列每一格都克隆一份快照
     *   （既是几百毫秒的卡顿，也是成百 MB 的撤销栈内存）。
     * 替换只改写既有格子的 `chars`、不会新增格子，故不需要 `null` 项。
     * @param {string} label - 操作标签
     * @param {Cell[]} cells - 会被改动的单元格（`values` 里的活引用）
     * @param {() => void} action - 实际的变更操作
     */
    public runWithCellsUndo(label: string, cells: Cell[], action: () => void): void {
        const anchorBefore = this._captureAnchor();
        const before = this._captureCellsPatch(cells);
        const prevSuppress = this._suppressValuesUndo;
        this._suppressValuesUndo = true;
        this._cellPatchUndoDepth++;
        try {
            action();
        } finally {
            this._cellPatchUndoDepth--;
            this._suppressValuesUndo = prevSuppress;
            this._pruneEmptyCells();
            this.markContentChanged();
            if (!this._isApplyingUndo) {
                const after = this._captureCellsPatch(cells);
                if (!this._cellPatchesEqual(before, after)) {
                    this.undoManager.push({
                        type: 'cellPatch', data: { before, after }, timestamp: Date.now(), label,
                        anchorBefore, anchorAfter: this._captureAnchor()
                    });
                }
            }
            // 与矩形版同理：包裹期间跳过了基准刷新，这里补一次
            this._refreshBaseline();
            this.emit(DataEvents.VALUES_CHANGED, this._values);
        }
    }

    /**
     * 把一批「脚本写入」（值 + 样式 + 合并跨度）落到工作表上，整体纳入撤销栈
     *
     * ## 为什么要放在数据层，而不是留在 Python 桥里
     *
     * 脚本写表和手点写表**必须是同一套规则**，否则同一个动作两种结果：
     * - 写「被合并覆盖的位置」（D3:F5 里的 E4）要重定向到锚点（见 {@link ensureWritableCell}），
     *   否则会往 `values` 里塞一条永远被遮挡层盖住、却能被子查找/替换扫到的垃圾记录；
     * - 给锚点写 `colspan` / `rowspan` 等于「用脚本合并」，块内被覆盖的格子必须按
     *   {@link executeMergeCells} 的口径清理，并重算合并倒排；
     * - 被清掉的格子必须**先捕获进补丁**，否则撤销只回得来锚点、回不来被删的那几格。
     *
     * ## 两阶段分离
     *
     * 阶段一只读（解析落点、取格子、算跨度与覆盖格），阶段二才在撤销作用域内改数据；
     * 交替「改 + 按名查找」会让每格都重建一次整份索引（O(N)）。
     *
     * @param {string} label - 撤销栈里的操作标签
     * @param {Object} payload - 单元格名 → 字段字典（`value` / 样式 / `colspan` / `rowspan`）
     * @returns {number} 实际写回的格子数（被合并重定向折叠后）
     */
    public applyScriptWrites(label: string, payload: Record<string, Record<string, unknown>>): number {
        const names = Object.keys(payload);
        if (names.length === 0) return 0;

        // ── 阶段一：只读 ──
        // ① 合并重定向 + 同一锚点的多次写入合并（后写覆盖先写）
        const targets = new Map<string, Record<string, unknown>>();
        for (let i = 0; i < names.length; i++) {
            const name = names[i];
            const target = this.resolveWriteTargetName(name);
            const fields = payload[name];
            const prev = targets.get(target);
            if (prev === undefined) targets.set(target, Object.assign({}, fields));
            else Object.assign(prev, fields);
        }

        // ② 取落点格（必要时新建；只新建锚点）
        const byName = new Map<string, Cell>();
        const cells: Cell[] = [];
        targets.forEach((_fields, targetName) => {
            const cell = this.ensureWritableCell(targetName);
            byName.set(targetName, cell);
            cells.push(cell);
        });

        // ③ 跨度锚点：本次写入了 colspan / rowspan 的落点（先按「写入后」的跨度算清理范围）
        const spanTouched = new Set<string>();
        const anchors: { name: string; col: number; row: number; colspan: number; rowspan: number }[] = [];
        targets.forEach((fields, targetName) => {
            if (!('colspan' in fields) && !('rowspan' in fields)) return;
            spanTouched.add(targetName);
            const cell = byName.get(targetName) as Cell;
            const colspan = this._toSpan(fields.colspan, cell.colspan);
            const rowspan = this._toSpan(fields.rowspan, cell.rowspan);
            if (colspan <= 1 && rowspan <= 1) return;      // 改成 1 = 取消合并，无覆盖格
            const { col, row } = this.getCellColAndRow(targetName);
            if (!isFinite(col) || !isFinite(row) || col < 1 || row < 1) return;
            anchors.push({ name: targetName, col, row, colspan, rowspan });
        });

        // ④ 被覆盖的格子也必须进补丁：清理会把它们从 values 里摘掉，不捕获就撤不回来
        for (let i = 0; i < anchors.length; i++) {
            const a = anchors[i];
            for (let row = a.row; row <= a.row + a.rowspan - 1; row++) {
                for (let col = a.col; col <= a.col + a.colspan - 1; col++) {
                    if (col === a.col && row === a.row) continue;
                    const name = this.getCellName(col, row);
                    if (byName.has(name)) continue;         // 本身就是另一个落点，会在阶段二被写入
                    const cell = this.getCellByName(name);
                    if (cell === undefined) continue;
                    byName.set(name, cell);
                    cells.push(cell);
                }
            }
        }

        // ── 阶段二：结构性修改 ──
        this.runWithCellsUndo(label, cells, () => {
            targets.forEach((fields, targetName) => {
                const cell = byName.get(targetName) as Cell;
                if ('value' in fields) {
                    const value = fields.value;
                    if (value === undefined || value === null || value === '') cell.clearValue();
                    else cell.addValue(String(value));
                }
                // 样式：白名单驱动（键来自脚本，不能直接反射赋值）
                const raw = cell as unknown as Record<string, unknown>;
                for (const key in fields) {
                    if (key === 'value' || key === 'colspan' || key === 'rowspan') continue;
                    if (!SCRIPT_WRITABLE_FIELDS.has(key)) continue;
                    raw[key] = fields[key];
                }
                if ('colspan' in fields) cell.colspan = this._toSpan(fields.colspan, cell.colspan);
                if ('rowspan' in fields) cell.rowspan = this._toSpan(fields.rowspan, cell.rowspan);
            });
            // 与「手动合并」同口径：清掉落点块内被覆盖的格子（带边框/底色的留着只清文字）
            this._pruneCoveredCellsBySpans(anchors);
            // 跨度被改写过（含改回 1 = 取消合并）：合并倒排必须重算，
            // 否则 findMergeRect 还按旧跨度回答（既不改长度也不置脏，索引看不出变化）
            if (spanTouched.size > 0) this.invalidateCellIndex();
        });

        return targets.size;
    }

    /**
     * 解析跨度值（脚本传进来的可能是字符串 / 小数 / 缺省）
     * @param value - 写入值
     * @param fallback - 未给出时沿用的现值
     * @returns {number} 不小于 1 的整数跨度
     * @private
     */
    private _toSpan(value: unknown, fallback: number | undefined): number {
        const n = typeof value === 'number' ? value : Number(value);
        if (isFinite(n) && n >= 1) return Math.floor(n);
        return fallback === undefined ? 1 : fallback;
    }

    /**
     * 按给定格子逐个捕获补丁（{@link runWithCellsUndo} 的「一半」）
     *
     * 顺序 = 入参数组顺序，因此 before / after 两次捕获必然同序（引用同一份数组），
     * 满足 {@link _cellPatchesEqual} 的逐项比对前提。
     * 刻意**不按矩形遍历**：命中是稀疏的，矩形遍历会为区域内每一格都克隆一份快照。
     *
     * `_pruneEmptyCells()` 可能把「被替换成空」的格子从 `values` 里摘掉，但本数组
     * 持有的是对象引用，两次捕获读到的仍是同一个对象（内容为空）—— 于是 before/after
     * 不等、记录入栈，撤销时由 `_applyCellPatch` 把这个不存在的格子补回 `values`。
     * @param {Cell[]} cells - 待捕获的格子
     * @returns {CellPatchEntry[]} 补丁
     * @private
     */
    private _captureCellsPatch(cells: Cell[]): CellPatchEntry[] {
        const out: CellPatchEntry[] = [];
        for (let i = 0; i < cells.length; i++) {
            out.push({ cell: cells[i].cell, data: cells[i].clone() });
        }
        return out;
    }

    /**
     * 回写局部单元格补丁（撤销/重做恢复）
     *
     * 一次遍历 `_values` 完成「命中即替换、data 为 null 即删除」；
     * 补丁中未被命中的项（操作前不存在、操作后新增的格子）追加到末尾。
     * 全程只做一次 {@link invalidateCellIndex}，不像「边 push 边查」那样
     * 每个新增格子都触发一次整表索引重建。
     * @param {CellPatchEntry[]} patch - 待回写的补丁
     * @private
     */
    private _applyCellPatch(patch: CellPatchEntry[]): void {
        const map = new Map<string, Cell | null>();
        for (let i = 0; i < patch.length; i++) {
            const e = patch[i];
            // 克隆后再写入：撤销栈里的对象必须只读，否则后续编辑会改坏快照
            map.set(e.cell, e.data ? (e.data as Cell).clone() : null);
        }
        const values = this._values;
        let w = 0;
        for (let i = 0; i < values.length; i++) {
            const c = values[i];
            if (!map.has(c.cell)) {
                values[w++] = c;
                continue;
            }
            const replacement = map.get(c.cell);
            map.delete(c.cell);
            if (replacement) values[w++] = replacement;
        }
        values.length = w;
        // 补丁里没被消费掉的 = 该时刻新增的格子（补丁按行优先有序，追加顺序即行列序）
        map.forEach((v) => { if (v) values.push(v); });
        this.invalidateCellIndex();
    }

    /**
     * 撤销/重做期间回写局部补丁并广播
     *
     * 与 {@link _applyValuesSnapshot} 同构：`_isApplyingUndo` 期间回写，
     * 回写后刷新基准（下次 commitValues 的「变化前」）并广播 VALUES_CHANGED 驱动重绘。
     * @param {CellPatchEntry[]} patch - 待回写的补丁
     * @private
     */
    private _applyCellPatchAndBroadcast(patch: CellPatchEntry[]): void {
        this._isApplyingUndo = true;
        try {
            this._applyCellPatch(patch);
            this._refreshBaseline();
            this.emit(DataEvents.VALUES_CHANGED, this._values);
        } finally {
            this._isApplyingUndo = false;
        }
    }

    /**
     * 以「结构命令」为单位执行一次可撤销操作（插入/删除整行整列，大表专用）
     *
     * 与 {@link runWithFullStateUndo} 的区别：插入/删除整行整列会把下方/右侧
     * **所有**格子重命名，前后快照等价于两份整表深拷贝 —— 大表下这条路被
     * {@link LARGE_TABLE_UNDO_CELLS} 阈值整体关掉，历史表现是「大表里插入/删除
     * 行列后 Ctrl+Z 静默失效」。本方法改用**命令模式**（见 {@link StructuralSnapshot}）：
     * 只记「操作参数 + 被删带数据」，代价与被删带内的格子数成正比、与整表无关。
     *
     * 正确性依赖撤销栈的严格 LIFO：撤销某条记录时，其后的操作都已被先撤销，
     * 表格必然处于「该操作刚做完」的状态，逆操作因此是安全的。
     *
     * 记账上限：删除带内的格子克隆数超过 {@link LARGE_TABLE_PATCH_CELLS} 时
     * 放弃记账（宁可这一次不可撤销，也不能为它扛一份巨型快照）；
     * 操作实际未生效（行列数未变化）时同样不记。
     * @param {string} label - 操作标签
     * @param {'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols'} op - 操作类别
     * @param {number} at - 插入位置 / 删除带起始行（列）号（1-based）
     * @param {number} count - 插入 / 删除的数量（delete 时为**请求**数量，实际数量按表尾收敛后入账）
     * @param {() => void} action - 实际的变更操作（内部调用 insertRows/deleteRows/...）
     * @example
     * data.runWithStructuralUndo('删除行', 'deleteRows', 5, 2, () => data.deleteRows(5, 6));
     */
    public runWithStructuralUndo(
        label: string,
        op: StructuralSnapshot['op'],
        at: number,
        count: number,
        action: () => void
    ): void {
        const anchorBefore = this._captureAnchor();
        // delete 类操作先捕获被删带数据（撤销时的逆操作需要它还原）；insert 无需捕获
        const isDelete = op === 'deleteRows' || op === 'deleteCols';
        const before = isDelete ? this._captureStructuralBefore(op, at, count) : null;
        // 超上限放弃记账：快照仍要捕获表头/格子的情形只有 delete，insert 恒可记账
        const record = (before === null) || (before.removedCells.length <= DataCollection.LARGE_TABLE_PATCH_CELLS);
        const rowsLenBefore = this.rowHeaders.length;
        const colsLenBefore = this.colHeaders.length;
        const prevSuppress = this._suppressValuesUndo;
        this._suppressValuesUndo = true;
        try {
            action();
        } finally {
            this._suppressValuesUndo = prevSuppress;
            this._pruneEmptyCells();
            this.markContentChanged();
            // 包裹期间跳过了基准刷新，这里补一次，避免后续 commitValues 拿过期基准
            this._refreshBaseline();
            if (!this._isApplyingUndo && record) {
                // 操作未生效（行列数都没变）不记账，避免撤销时执行一次凭空的逆操作
                const changed = this.rowHeaders.length !== rowsLenBefore || this.colHeaders.length !== colsLenBefore;
                if (changed) {
                    const snapshot: StructuralSnapshot = (before !== null)
                        ? before
                        : { op, at, count, removedCells: [], removedHeaders: [], shrunkMerges: [] };
                    // 实际数量按表尾收敛（deleteRows/deleteCols 内部同款钳制），
                    // 逆操作 deleteRows(at, at+count-1) 才能删准同一段
                    if (op === 'deleteRows') snapshot.count = Math.max(0, Math.min(at + count - 1, rowsLenBefore) - at + 1);
                    if (op === 'deleteCols') snapshot.count = Math.max(0, Math.min(at + count - 1, colsLenBefore) - at + 1);
                    this.undoManager.push({
                        type: 'structural', data: snapshot, timestamp: Date.now(), label,
                        anchorBefore, anchorAfter: this._captureAnchor()
                    });
                }
            }
        }
    }

    /**
     * 捕获「删除整行/整列」的逆操作所需数据（{@link runWithStructuralUndo} 的记账一半）
     *
     * 遍历一次 values（O(已用格数)，与 deleteRows 自身的遍历同量级）：
     * - 落在删除带内的格子 → 克隆入 removedCells（名字为删除前的名字）；
     * - 锚点在删除带上方/左侧、跨距末端伸进删除带的合并 → 原始跨距入 shrunkMerges
     *   （删除会按重叠数收缩它们，重叠数可能小于 count，撤销时的结构插回
     *   只能按 count 还原「整段跨越」的合并，端点落在带内的必须靠这份记录）。
     * 表头按删除带切片（含自定义行高列宽与隐藏标记）。
     * @param op - 'deleteRows' | 'deleteCols'
     * @param at - 删除带起始行/列号（1-based）
     * @param count - 请求删除的数量
     * @returns {StructuralSnapshot} 快照（count 已按表尾收敛；带为空时 removed* 为空数组）
     * @private
     */
    private _captureStructuralBefore(op: 'deleteRows' | 'deleteCols', at: number, count: number): StructuralSnapshot {
        const snapshot: StructuralSnapshot = { op, at, count, removedCells: [], removedHeaders: [], shrunkMerges: [] };
        const removedCells: CellPatchEntry[] = [];
        const shrunkMerges: Array<{ cell: string; rowspan?: number; colspan?: number }> = [];
        if (op === 'deleteRows') {
            const clampedEnd = Math.min(at + count - 1, this.rowHeaders.length);
            snapshot.count = Math.max(0, clampedEnd - at + 1);
            for (const cell of this._values) {
                const { col, row } = this.getCellColAndRow(cell.cell);
                if (row >= at && row <= clampedEnd) {
                    removedCells.push({ cell: cell.cell, data: cell.clone() });
                } else if (cell.rowspan && cell.rowspan > 1 && row < at && row + cell.rowspan - 1 >= at) {
                    shrunkMerges.push({ cell: cell.cell, rowspan: cell.rowspan });
                }
            }
            snapshot.removedHeaders = this.rowHeaders.rowHeaders.slice(at - 1, clampedEnd);
        } else {
            const clampedEnd = Math.min(at + count - 1, this.colHeaders.length);
            snapshot.count = Math.max(0, clampedEnd - at + 1);
            for (const cell of this._values) {
                const { col, row } = this.getCellColAndRow(cell.cell);
                if (col >= at && col <= clampedEnd) {
                    removedCells.push({ cell: cell.cell, data: cell.clone() });
                } else if (cell.colspan && cell.colspan > 1 && col < at && col + cell.colspan - 1 >= at) {
                    shrunkMerges.push({ cell: cell.cell, colspan: cell.colspan });
                }
            }
            snapshot.removedHeaders = this.colHeaders.colHeaders.slice(at - 1, clampedEnd);
        }
        snapshot.removedCells = removedCells;
        snapshot.shrunkMerges = shrunkMerges;
        return snapshot;
    }

    /**
     * 回写结构性操作（撤销/重做恢复）
     *
     * 撤销 = 逆操作：insert → 删掉刚插入的那段；delete → 结构上插回 + 还原被删数据。
     * 重做 = 原操作重放：参数即可，无需任何快照数据。
     * @param {StructuralSnapshot} snap - 结构快照
     * @param {boolean} isUndo - true = 撤销（执行逆操作）；false = 重做（重放原操作）
     * @private
     */
    private _applyStructuralOp(snap: StructuralSnapshot, isUndo: boolean): void {
        this._isApplyingUndo = true;
        try {
            if (snap.op === 'insertRows') {
                if (isUndo) this.deleteRows(snap.at, snap.at + snap.count - 1);
                else this.insertRows(snap.at, snap.count);
            } else if (snap.op === 'deleteRows') {
                if (isUndo) this._restoreDeletedRows(snap);
                else this.deleteRows(snap.at, snap.at + snap.count - 1);
            } else if (snap.op === 'insertCols') {
                if (isUndo) this.deleteCols(snap.at, snap.at + snap.count - 1);
                else this.insertCols(snap.at, snap.count);
            } else {
                if (isUndo) this._restoreDeletedCols(snap);
                else this.deleteCols(snap.at, snap.at + snap.count - 1);
            }
            this._refreshBaseline();
        } finally {
            this._isApplyingUndo = false;
        }
    }

    /**
     * 还原被删除的行（deleteRows 的逆操作）
     *
     * 顺序刻意为「先结构、再表头、再格子、最后合并跨距」：
     * 1. {@link insertRows} 把下方格子整体移回原位、按 count 扩展整段跨越的合并
     *    （端点落在删除带内部的合并由第 4 步兜底）；
     * 2. 被删行的表头（自定义行高/隐藏标记） splice 回原位 —— insertRows 塞的是
     *    默认表头，直接整体替换掉这 count 个；
     * 3. 被删格子按删除前的名字克隆回填 —— 此时那些位置已被第 1 步清空，无名字冲突；
     * 4. shrunkMerges 里记录的合并锚点按**原始跨距**赋回（赋值语义，天然幂等，
     *    与第 1 步对「整段跨越」合并的扩展不冲突）。
     * @param {StructuralSnapshot} snap - 结构快照（op 必为 'deleteRows'）
     * @private
     */
    private _restoreDeletedRows(snap: StructuralSnapshot): void {
        // 1) 结构插回：下方格子下移归位 + 整段跨越的合并扩展 rowspan
        this.insertRows(snap.at, snap.count);
        // 2) 表头还原：整体替换 insertRows 塞入的默认表头（克隆入栈，快照保持只读）
        if (snap.removedHeaders.length > 0) {
            this.rowHeaders.rowHeaders.splice(
                snap.at - 1, snap.count, ...snap.removedHeaders.map(rh => ({ ...rh })));
            this.rowHeaders.recalcRowPositions();
        }
        // 3) 格子还原：按删除前的名字回填（此刻位置为空，push 即插入）
        if (snap.removedCells.length > 0) {
            for (const e of snap.removedCells) {
                if (e.data) this._values.push((e.data as Cell).clone());
            }
            this.invalidateCellIndex();
        }
        // 4) 跨距还原：端点落在删除带内的合并无法由第 1 步还原，按原始值赋回
        for (const m of snap.shrunkMerges) {
            const anchor = this.getCellByName(m.cell);
            if (!anchor) continue;
            if (m.rowspan !== undefined) anchor.rowspan = m.rowspan;
            if (m.colspan !== undefined) anchor.colspan = m.colspan;
        }
    }

    /**
     * 还原被删除的列（deleteCols 的逆操作）
     *
     * 与 {@link _restoreDeletedRows} 完全对称，仅轴向不同。
     * @param {StructuralSnapshot} snap - 结构快照（op 必为 'deleteCols'）
     * @private
     */
    private _restoreDeletedCols(snap: StructuralSnapshot): void {
        this.insertCols(snap.at, snap.count);
        if (snap.removedHeaders.length > 0) {
            this.colHeaders.colHeaders.splice(
                snap.at - 1, snap.count, ...snap.removedHeaders.map(ch => ({ ...ch })));
            this.colHeaders.recalcColPositions();
        }
        if (snap.removedCells.length > 0) {
            for (const e of snap.removedCells) {
                if (e.data) this._values.push((e.data as Cell).clone());
            }
            this.invalidateCellIndex();
        }
        for (const m of snap.shrunkMerges) {
            const anchor = this.getCellByName(m.cell);
            if (!anchor) continue;
            if (m.rowspan !== undefined) anchor.rowspan = m.rowspan;
            if (m.colspan !== undefined) anchor.colspan = m.colspan;
        }
    }

    /**
     * 开始一次跨事件的连续操作撤销捕获（如拖拽调整行高/列宽）
     * 立即捕获「变化前」基准，并抑制 values setter 的自动快照记录，
     * 直到调用 {@link endContinuousUndo} 提交。两次调用必须成对出现。
     *
     * 基准按表规模分档：
     * - 小表：整表全量快照（{@link FullStateSnapshot}）；
     * - 大表：'headers' 档稀疏快照（{@link StatePatchSnapshot}）—— 只浅拷贝
     *   行/列表头与标量状态、不克隆任何格子。拖拽行高列宽只改表头，
     *   一万多行表头的浅拷贝仅几毫秒，使大表下拖拽调宽高同样可以撤销。
     * @returns {ContinuousUndoSnapshot} 操作前基准快照（含档位标记）
     */
    public beginContinuousUndo(): ContinuousUndoSnapshot {
        this._suppressValuesUndo = true;
        if (this._undoTrackingEnabled()) {
            return { kind: 'full', snapshot: this._captureFullState() };
        }
        return { kind: 'patch', snapshot: this._captureStatePatch(null, true) as StatePatchSnapshot };
    }

    /**
     * 结束连续操作并提交撤销记录
     * 按 {@link beginContinuousUndo} 捕获时的同档基准捕获「变化后」快照，
     * 前后不同时推入 undoManager（小表 'fullState' / 大表 'statePatch'），
     * 并恢复 values setter 的自动记录。整个拖拽手势折叠为一条撤销记录。
     *
     * 抑制标记必须在 finally 中无条件恢复：早期实现在 try 之外提前 return，
     * 会让 `_suppressValuesUndo` 永久停在 true，此后所有变更都不再记账。
     * @param {ContinuousUndoSnapshot} before - {@link beginContinuousUndo} 返回的基准快照
     * @param {string} label - 操作标签
     */
    public endContinuousUndo(before: ContinuousUndoSnapshot, label: string): void {
        try {
            if (before.kind === 'full') {
                // 防御：拖拽期间表规模不应变化；若已切到大表则放弃，避免整表深拷贝
                if (!this._undoTrackingEnabled()) return;
                const after = this._captureFullState();
                if (!this._fullStateSnapshotsEqual(before.snapshot, after)) {
                    const anchor = this._captureAnchor();
                    this.undoManager.push({
                        type: 'fullState', data: { before: before.snapshot, after }, timestamp: Date.now(), label,
                        anchorBefore: anchor, anchorAfter: anchor
                    });
                }
            } else {
                if (this._undoTrackingEnabled()) return;
                const after = this._captureStatePatch(null, true) as StatePatchSnapshot;
                if (!this._statePatchSnapshotsEqual(before.snapshot, after)) {
                    const anchor = this._captureAnchor();
                    this.undoManager.push({
                        type: 'statePatch', data: { before: before.snapshot, after }, timestamp: Date.now(), label,
                        anchorBefore: anchor, anchorAfter: anchor
                    });
                }
            }
        } finally {
            this._suppressValuesUndo = false;
        }
    }

    /**
     * 回写全量状态快照（撤销/重做恢复）
     * 依次恢复筛选条件注册表、values（走 values setter，回写期间抑制自动记录并广播
     * VALUES_CHANGED）、图片、行/列表头（重算行列总尺寸）、网格线开关、活动单元格与选区；
     * 最后再次广播 VALUES_CHANGED，驱动筛选按钮外观等派生 UI 与恢复后的完整状态同步。
     * @param {FullStateSnapshot} snapshot - 待恢复的快照
     */
    private _applyFullStateSnapshot(snapshot: FullStateSnapshot): void {
        this._isApplyingUndo = true;
        try {
            // 先恢复筛选条件注册表：values 赋值触发的 VALUES_CHANGED 处理器将读到恢复后的条件
            this._restoreFilterConditions(snapshot.filters);
            this.values = this.cloneValues(snapshot.values as Cell[]);
            this.rowHeaders.rowHeaders = snapshot.rowHeaders.map(rh => ({ ...rh }));
            this.rowHeaders.recalcAllRowHeight();
            this.colHeaders.colHeaders = snapshot.colHeaders.map(ch => ({ ...ch }));
            this.colHeaders.recalcAllColWidth();
            this.showGridLines = snapshot.showGridLines;
            this.activedCell = snapshot.activedCell;
            this.selection = snapshot.selection;
            // 图片随快照一起回写（「打开文件」是可撤销的全量操作，
            // 只回写单元格会让撤销后的文件顶着另一个文件的图）。
            // 旧快照没有 images 字段时保持现状，不做「清空」这种破坏性推断。
            if (Array.isArray(snapshot.images)) {
                this._images = snapshot.images.slice();
                this._imagesDirty = true;
            }
            // 全部字段恢复完毕后广播一次，确保派生 UI（如筛选按钮激活色）与最终状态一致
            this.emit(DataEvents.VALUES_CHANGED, this._values);
        } finally {
            this._isApplyingUndo = false;
        }
    }

    /**
     * 判断单元格是否落在选区内（按行列号做数值矩形判定）
     *
     * **勿用字符串 `includes`**：`'B22:D24'` 会误判 `B2`、`D2`（见 {@link getMergedRangeAt}）。
     * @param {string} cell - 单元格名称
     * @param {string | null} selection - 选区串（`'A1:C3'`）；null 视为无选区
     * @returns {boolean} 是否在选区内
     * @private
     */
    private _isCellInsideSelection(cell: string, selection: string | null): boolean {
        if (!selection) return false;
        const [startName, endName] = selection.split(':');
        const start = this.getCellColAndRow(startName);
        const end = this.getCellColAndRow(endName ?? startName);
        const target = this.getCellColAndRow(cell);
        if (!isFinite(target.col) || !isFinite(target.row)) return false;
        return target.col >= Math.min(start.col, end.col)
            && target.col <= Math.max(start.col, end.col)
            && target.row >= Math.min(start.row, end.row)
            && target.row <= Math.max(start.row, end.row);
    }

    /**
     * 撤销/重做恢复回调
     * 处理两种操作类型：撤销回写 before，重做回写 after。
     * - 'values'：仅回写单元格数组；
     * - 'fullState'：回写 values + 行/列表头 + 网格线 + 活动单元格 + 选区。
     * @param {UndoAction} action - 操作对象
     * @param {boolean} isUndo - true=撤销，false=重做
     */
    private _undoRestoreCallback = (action: UndoAction, isUndo: boolean): void => {
        if (action.type === 'values' && action.data
            && Array.isArray((action.data as { before: Cell[]; after: Cell[] }).before)
            && Array.isArray((action.data as { before: Cell[]; after: Cell[] }).after)) {
            const snapshot = isUndo ? (action.data as { before: Cell[]; after: Cell[] }).before : (action.data as { before: Cell[]; after: Cell[] }).after;
            this._applyValuesSnapshot(snapshot);
        } else if (action.type === 'fullState' && action.data) {
            const snapshot = isUndo
                ? (action.data as { before: FullStateSnapshot; after: FullStateSnapshot }).before
                : (action.data as { before: FullStateSnapshot; after: FullStateSnapshot }).after;
            this._applyFullStateSnapshot(snapshot);
        } else if (action.type === 'cellPatch' && action.data) {
            const patch = isUndo
                ? (action.data as { before: CellPatchEntry[]; after: CellPatchEntry[] }).before
                : (action.data as { before: CellPatchEntry[]; after: CellPatchEntry[] }).after;
            this._applyCellPatchAndBroadcast(patch);
        } else if (action.type === 'statePatch' && action.data) {
            const snapshot = isUndo
                ? (action.data as { before: StatePatchSnapshot; after: StatePatchSnapshot }).before
                : (action.data as { before: StatePatchSnapshot; after: StatePatchSnapshot }).after;
            this._applyStatePatchSnapshot(snapshot);
        } else if (action.type === 'structural' && action.data) {
            // 结构撤销（插入/删除整行整列）：data 是单份命令快照（非 before/after 对），
            // 撤销 = 逆操作，重做 = 再执行一遍原操作
            this._applyStructuralOp(action.data as StructuralSnapshot, isUndo);
        } else if (action.type === 'sortPermutation' && action.data) {
            // 排序撤销：data 是排列命令（非状态快照），撤销 = 逆置换，重做 = 原置换
            this._applySortPermutation(action.data as SortPermutationSnapshot, isUndo);
        } else {
            return;
        }
        // 撤销/重做后把视口带回操作发生的区域：
        // 只还原数据不动选区的话，百万行表里撤销第 8000 行的修改后画面仍停在第 1 行，
        // 用户看到的是「什么都没发生」。这里统一恢复锚点并卷进可视区，
        // 随后补一次 VALUES_CHANGED —— 滚动偏移是在数据回写**之后**才变的，
        // 光靠回写时那次广播画不到新位置。
        const anchor = isUndo ? action.anchorBefore : action.anchorAfter;
        this._revealAnchor(anchor);
        // activedCell 指向「本次撤销/重做操作的单元格」（影响面首格）：
        // 锚点恢复的是操作前/后的光标位置，与「被操作格」未必同格（如在 A1 选中后
        // 撤销 F10 的输入，锚点会把活动格留在 A1）；活动格与选区都改落到操作发生处
        const targetCell = this._resolveUndoTargetCell(action, isUndo);
        if (targetCell) {
            this.activedCell = targetCell;
            const { col, row } = this.getCellColAndRow(targetCell);
            const valid = isFinite(col) && isFinite(row) && col >= 1 && row >= 1;
            // 选区必须跟着活动格一起落到「内容真正变化」的格子上：
            // 锚点恢复的是操作前/后的选区，与被改动的格子未必同处
            // （典型场景：A1 处选中后，撤销脚本对 D10 的写入）。
            // 只改 activedCell 的话，选择框会留在原处，视觉上「撤销没生效」，
            // 且选区与活动格不一致会让后续进入编辑态的落点判断跟着错位。
            // 目标格本就落在当前选区内时保持原选区（撤销 A1:C3 的批量格式不能塌成 A1）。
            if (!this._isCellInsideSelection(targetCell, this._selection)) {
                // 合并区内的目标格按整块选中，避免选择框只框住合并块的一角
                const merged = valid ? this.getMergedRangeAt(col, row) : null;
                this.selection = merged ?? `${targetCell}:${targetCell}`;
            }
            if (valid) this.scrollCellIntoView(col, row);
        }
        this.emit(DataEvents.VALUES_CHANGED, this._values);
        // 回写完成后广播快照恢复事件：activedCell/selection 与当前值相同时
        // updateProperty 不会发出差量事件（ACTIVED_CELL_CHANGED 等），
        // 依赖差量事件刷新外观的工具栏分区需监听本事件同步按钮状态
        this.emit(DataEvents.SNAPSHOT_RESTORED, action, isUndo);
    }

    /**
     * 推导撤销/重做「操作发生的单元格」（影响面首格），供 activedCell 跳转
     *
     * 各快照类型记录的数据不同，取法也不同：
     * - cellPatch / statePatch：补丁列表首格（行优先排序的第一个格子）。
     *   两侧都要看：区域原本**没有格子**时（给空白区域着色、合并空白区），
     *   撤销方向取 `before` 会得到空列表，而该区域本来也没有可跳的格子名，
     *   必须回退到 `after` 才能拿到目标格；
     * - structural：行操作取 A{at}（首列），列操作取 {at}列1行（首行）；
     * - sortPermutation：置换范围首位置（positions 为 0-based，需 +1）；
     * - fullState：快照内记录的 activedCell；
     * - values：入栈时对比前后快照求得的首格（{@link ValuesSnapshot.focusCell}）。
     *   旧快照没有该字段时返回 null，保持锚点恢复的活动格。
     * @param {UndoAction} action - 操作对象
     * @param {boolean} isUndo - true=撤销，false=重做
     * @returns {string | null} 目标单元格名，无法推导时为 null
     */
    private _resolveUndoTargetCell(action: UndoAction, isUndo: boolean): string | null {
        const side = isUndo ? 'before' : 'after';
        const data = action.data as Record<string, any>;
        switch (action.type) {
            case 'values':
                // before/after 两侧指向同一格（内容变化格），无需按方向取
                return (data?.focusCell as string | undefined) ?? null;
            case 'cellPatch': {
                // 优先取**本次方向**那一侧的补丁：撤销看 before、重做看 after。
                // 固定优先 before 会让重做落回「操作前」的格子；另一侧仅作回退
                // （区域原本没有格子时该侧为空列表，见方法注释）。
                const primary = data?.[side] as CellPatchEntry[] | undefined;
                const fallback = (isUndo ? data?.after : data?.before) as CellPatchEntry[] | undefined;
                return primary?.[0]?.cell ?? fallback?.[0]?.cell ?? null;
            }
            case 'statePatch': {
                const primary = data?.[side] as StatePatchSnapshot | undefined;
                const fallback = (isUndo ? data?.after : data?.before) as StatePatchSnapshot | undefined;
                return primary?.cells?.[0]?.cell ?? fallback?.cells?.[0]?.cell ?? null;
            }
            case 'structural': {
                const cmd = data as unknown as StructuralSnapshot;
                if (cmd.op === 'insertRows' || cmd.op === 'deleteRows') return this.getCellName(1, cmd.at);
                return this.getCellName(cmd.at, 1);
            }
            case 'sortPermutation': {
                const perm = data as unknown as SortPermutationSnapshot;
                const first = perm.positions?.[0];
                if (first === undefined) return null;
                return perm.axis === 'column' ? this.getCellName(1, first + 1) : this.getCellName(first + 1, 1);
            }
            case 'fullState': {
                const snapshot = data?.[side] as FullStateSnapshot | undefined;
                return snapshot?.activedCell ?? null;
            }
            default:
                return null;
        }
    }

    /**
     * 执行撤销
     * 编辑期间走编辑内逐字撤销（_undoEditStep）；非编辑期间走 undoManager。
     * @returns {boolean} 是否撤销成功
     */
    public undo(): boolean {
        if (this.isEditting) {
            return this._undoEditStep();
        }
        return this.undoManager.undo(this._undoRestoreCallback);
    }

    /**
     * 执行重做
     * 编辑期间走编辑内逐字重做（_redoEditStep）；非编辑期间走 undoManager。
     * @returns {boolean} 是否重做成功
     */
    public redo(): boolean {
        if (this.isEditting) {
            return this._redoEditStep();
        }
        return this.undoManager.redo(this._undoRestoreCallback);
    }

    /**
     * 开始一个撤销操作组
     * 组内的多次 commitValues() 会在 endUndoGroup() 时折叠为一步撤销，
     * 典型用途：包裹一段单元格编辑（多次按键）使其成为单步撤销。
     * @param {string} [label] - 组标签
     */
    public beginUndoGroup(label?: string): void {
        this.undoManager.beginGroup(label);
    }

    /**
     * 结束当前撤销操作组并折叠为一步撤销
     */
    public endUndoGroup(): void {
        this.undoManager.endGroup();
    }

    /**
     * 将结构化光标位置（行/字符索引）转换为单元格内的全局字符绝对索引。
     * 这是"结构化 → 绝对"唯一的转换实现，所有需要绝对索引的地方都应复用它。
     * @param {{ lineIndex: number; charIndex: number }} pos - 结构化位置
     * @returns {number} 全局字符绝对索引
     */
    public toAbsoluteIndex(pos: { lineIndex: number; charIndex: number }): number {
        return this.getPreLineCharsLength(pos.lineIndex) + pos.charIndex;
    }

    /**
     * 将单元格内的全局字符绝对索引转换为结构化光标位置（行/字符索引）。
     * 这是"绝对 → 结构化"唯一的转换实现，与 toAbsoluteIndex 互为逆运算。
     * 边界规则与 syncFromTextAreaCursor 一致：定位到第一个"末尾偏移 >= value"的行，
     * charIndex 为 value 相对该行起始偏移的差值，并做上下界裁剪。
     * @param {number} value - 全局字符绝对索引
     * @returns {{ lineIndex: number; charIndex: number }} 结构化位置
     */
    public fromAbsoluteIndex(value: number): { lineIndex: number; charIndex: number } {
        const clamped = Math.max(0, value);
        let totalIndex = 0;
        for (let i = 0; i < this.lines.length; i++) {
            const lineLength = this.lines[i].text.length;
            if (totalIndex + lineLength >= clamped) {
                return { lineIndex: i, charIndex: clamped - totalIndex };
            }
            totalIndex += lineLength;
        }
        // 超出全部文本，落在最后一行末尾
        const lastIndex = Math.max(0, this.lines.length - 1);
        return { lineIndex: lastIndex, charIndex: this.lines[lastIndex]?.text.length ?? 0 };
    }

    /**
     * 根据结构化光标位置解析屏幕坐标（editX/editY）。
     * 优先使用渲染层注入的 cursorCoordResolver（calcCursorX + line.editY），
     * 未注入时回退到行起始坐标，保证坐标口径统一。
     * @param {number} lineIndex - 行索引
     * @param {number} charIndex - 字符索引
     * @returns {{ editX: number; editY: number }} 屏幕坐标
     * @private
     */
    private resolveCursorCoord(lineIndex: number, charIndex: number): { editX: number; editY: number } {
        if (this.cursorCoordResolver) {
            return this.cursorCoordResolver(lineIndex, charIndex);
        }
        const line = this.lines[lineIndex];
        return { editX: line?.editX ?? 0, editY: line?.editY ?? 0 };
    }

    /**
     * 获取光标绝对位置
     * 由唯一真相源 cursorPosition 派生，无 cursorPosition 时返回上次缓存值。
     * @returns {number} 光标绝对位置
     */
    get cursorAbsolutePosition(): number {
        if (this.cursorPosition) {
            this._cursorAbsolutePosition = this.toAbsoluteIndex(this.cursorPosition);
        }
        return this._cursorAbsolutePosition;
    }

    /**
     * 设置光标绝对位置
     * 写入时基于"当前最新的 lines"反算出结构化 cursorPosition（唯一真相源），
     * 并统一解析屏幕坐标，最后广播 CURSOR_POSITION_CHANGED，使 textarea 的
     * setSelectionRange 与画布光标保持同步。
     *
     * 注意：这里不做 `value === 缓存值` 的提前返回。因为插入/删除/换行等操作会改变
     * lines 布局，即使绝对索引数值不变，其对应的 lineIndex/charIndex 也可能变化，
     * 必须无条件重算并广播，保证 cursorPosition 与 textarea 选区始终一致。
     * @param {number} value - 光标绝对位置
     */
    set cursorAbsolutePosition(value: number) {
        // if (value === this._cursorAbsolutePosition) return;
        // 无条件重算并广播，确保 cursorPosition 与 textarea 选区始终一致。
        const oldPosition = this.cursorPosition;
        this._cursorAbsolutePosition = Math.max(0, value);
        const { lineIndex, charIndex } = this.fromAbsoluteIndex(this._cursorAbsolutePosition);
        const { editX, editY } = this.resolveCursorCoord(lineIndex, charIndex);
        this.cursorPosition = { lineIndex, charIndex, editX, editY };
        // 直接广播结构化光标位置（唯一真相源），沿用 (oldValue, newValue) 约定；
        // 监听方据此同步 textarea.setSelectionRange 与画布光标。
        this.emit(DataEvents.CURSOR_POSITION_CHANGED, oldPosition, this.cursorPosition);
    }

    /**
     * 根据列号获取列名
     * @param {number} num - 列号（从1开始）
     * @returns {string} 列名（如 A, B, ..., Z, AA, AB...）
     */
    public getColName(num: number): string {
        let numeric = (num - 1) % 26;
        let char = String.fromCodePoint(numeric + 65);
        let num2 = Math.floor((num - 1) / 26);

        return num2 > 0 ? `${this.getColName(num2)}${char}` : char;
    }

    /**
     * 根据列名获取列号
     * @param {string} char - 列名（如 A, B, Z, AA）
     * @returns {number} 列号（从1开始）
     */
    public getColNumber(char: string): number {
        return char.toUpperCase().split('').reduce((num, c, i, arr) => {
            return num + (c.charCodeAt(0) - 64) * Math.pow(26, arr.length - 1 - i);
        }, 0);
    }

    /**
     * 根据列号和行号获取单元格名称
     * @param {number} col - 列号
     * @param {number} row - 行号
     * @returns {string} 单元格名称（如 A1, B2）
     */
    public getCellName(col: number, row: number): string {
        return `${this.getColName(col)}${row}`;
    }

    /**
     * 获取单元格对应的列号和行号
     *
     * 单趟扫描实现（性能约定）：旧写法是两次正则 `replace` + `split('')` + `reduce`，
     * 每次调用都要新建两个字符串、一个字符数组，并按幂次逐个累加。本函数处在
     * 渲染、筛选、命中测试等**按单元格调用**的热路径上，几十万格时这些临时分配
     * 本身就是可观的开销，故改为一次遍历直接累加。
     * 语义与旧实现保持一致：只认大写字母作列（`/[^A-Z]/`）、只认 ASCII 数字作行
     * （`/[^\d]/`）；没有任何数字时行号为 `NaN`（对齐 `parseInt('')`）。
     * @param {string} cell - 单元格名称
     * @returns {{ col: number, row: number }} 包含列号和行号的对象
     */
    public getCellColAndRow(cell: string): { col: number, row: number } {
        let col = 0;
        let row = 0;
        let hasRowDigit = false;
        for (let i = 0; i < cell.length; i++) {
            const code = cell.charCodeAt(i);
            if (code >= 65 && code <= 90) {
                col = col * 26 + (code - 64);
            } else if (code >= 48 && code <= 57) {
                row = row * 10 + (code - 48);
                hasRowDigit = true;
            }
        }
        return { col, row: hasRowDigit ? row : NaN };
    }

    /**
     * 获取鼠标位置对应的单元格信息
     *
     * 匹配过程中会跳过隐藏行/列：在遍历 colHeaders/rowHeaders 时
     * 遇到 isHidden=true 的 header 直接 continue；fallback 分支
     * （越界超出最后一列/行边界）返回最后一个可见列/行；
     * 兜底逻辑（循环结束未匹配到任何可见 header 且鼠标在 sheet 内）
     * 返回第一个可见列/行。
     * @param {number} _x - 鼠标X坐标（从0开始）
     * @param {number} _y - 鼠标Y坐标（从0开始）
     * @param {Object} [options] - 命中选项
     * @param {boolean} [options.rawGeometry=false] - 为 true 时返回**纯几何**命中结果：
     * 跳过锚点覆盖（IN_ANCHOR）与合并单元格锚点吸附。行/列头拖选进行中使用
     * （起点 startLocation 为 IN_ROW_HEADER / IN_COL_HEADER）：从行头向下拖入
     * 内容区时，首列的纵向合并（如 A1:A3）会把经过行的行号吸附到合并锚点，
     * 导致整行选区端点被错误锚定；列头横向拖选同理。行/列头选区恒为
     * expandMerges=false（见 setSelection / setSelectedRange 的说明），
     * 命中端也必须使用未经合并归一化的原始行列号，两端语义才一致。
     * 无合并单元格时该选项为幂等 no-op，不改变普通点击/拖选行为。
     * @returns {{ location: MouseLocation, col: number, row: number }} 包含鼠标位置、列号和行号的对象
     */
    public getSheetCell(_x: number, _y: number, options?: { rawGeometry?: boolean }): {location: MouseLocation, col: number, row: number} {
        const rawGeometry = options?.rawGeometry === true;
        let location = MouseLocation.IN_CELL;
        let col = 1;
        let row = 1;
        const x = Math.round(_x / this.zoom); // 转换为实际坐标
        const y = Math.round(_y / this.zoom); // 转换为实际坐标

        // 找第一个/最后一个可见列行索引（1-based），用于兜底
        const firstVisibleCol = this._findFirstVisibleCol();
        const lastVisibleCol = this._findLastVisibleCol();
        const firstVisibleRow = this._findFirstVisibleRow();
        const lastVisibleRow = this._findLastVisibleRow();

        let colFound = false;
        for (let c = 0; c < this.colHeaders.length; c++) {
            const ch = this.colHeaders.getAt(c);
            // 边界检查优先
            if (x < 0) {
                location = MouseLocation.IN_ROW_HEADER;
                break;
            } else if (x > this.visibleView.sheetWidth / this.zoom) {
                col = lastVisibleCol; // 越界时返回最后一个可见列；
                colFound = true;
                break;
            }
            // 隐藏列不可命中：其位置已被右侧可见列覆盖（recalcColPositions 后 left 紧凑）
            if (ch.isHidden) continue;
            // left 已紧凑重算，直接用文档坐标匹配，无需隐藏补偿
            if (x + this.colHeaders.offsetWidth >= ch.left
                && x + this.colHeaders.offsetWidth < ch.left + ch.width) {
                col = c + 1;
                colFound = true;
                break;
            }
        }
        // 循环走完未匹配到可见列，兜底取第一个可见列
        if (!colFound && x >= 0 && x <= this.visibleView.sheetWidth / this.zoom && firstVisibleCol > 0) {
            col = firstVisibleCol;
        }

        let rowFound = false;
        for (let r = 0; r < this.rowHeaders.length; r++) {
            const rh = this.rowHeaders.getAt(r);
            // 边界检查优先
            if (y < 0) {
                location = MouseLocation.IN_COL_HEADER;
                break;
            } else if (y > this.visibleView.sheetHeight / this.zoom) {
                row = lastVisibleRow; // 越界时返回最后一个可见行；
                rowFound = true;
                break;
            }
            // 隐藏行不可命中：其位置已被下方可见行覆盖（recalcRowPositions 后 top 紧凑）
            if (rh.isHidden) continue;
            // top 已紧凑重算，直接用文档坐标匹配，无需隐藏补偿
            if (y + this.rowHeaders.offsetHeight >= rh.top
                && y + this.rowHeaders.offsetHeight < rh.top + rh.height) {
                row = r + 1;
                rowFound = true;
                break;
            }
        }
        // 循环走完未匹配到可见行，兜底取第一个可见行
        if (!rowFound && y >= 0 && y <= this.visibleView.sheetHeight / this.zoom && firstVisibleRow > 0) {
            row = firstVisibleRow;
        }
            
        // 锚点覆盖：若锚点自身隐藏则跳过（行/列头拖选进行中要求纯几何命中，一并跳过）
        if (!rawGeometry) {
            const anchor = this.getAnchorRect();
            if (x >= anchor.left && x <= (anchor.left + anchor.width) 
                && y >= anchor.top && y <= (anchor.top + anchor.height)) {
                const anchorColHidden = this.isColHidden(anchor.col - 1);
                const anchorRowHidden = this.isRowHidden(anchor.row - 1);
                if (!anchorColHidden && !anchorRowHidden) {
                    location = MouseLocation.IN_ANCHOR;
                    col = anchor.col;
                    row = anchor.row;
                }
            }
        }

        // 合并单元格覆盖：若合并后的目标单元格隐藏则跳过（rawGeometry 下行/列头拖选
        // 不允许吸附到合并锚点，否则整行/整列选区端点会被合并块拖偏）
        if (!rawGeometry) {
            const { col: mergedCol, row: mergedRow } = this.getMergeColAndRow(this.getCellName(col, row));
            if (mergedCol !== col || mergedRow !== row) {
                const mergedColHidden = this.isColHidden(mergedCol - 1);
                const mergedRowHidden = this.isRowHidden(mergedRow - 1);
                if (!mergedColHidden && !mergedRowHidden) {
                    col = mergedCol;
                    row = mergedRow;
                }
            }
        }
        
        return { location, col, row };
    }

    /**
     * 判断单元格是否有内容或样式
     * @param {Cell} cell - 单元格数据
     * @returns {boolean} 是否有内容或样式
     */
    private hasCellContent(cell: Cell): boolean {
        return !!(
            (cell.chars && cell.chars.length > 0) ||
            cell.backgroundColor ||
            cell.borderTopWidth || cell.borderBottomWidth ||
            cell.borderLeftWidth || cell.borderRightWidth ||
            cell.fontFamily || cell.fontSize || cell.fontWeight || cell.fontStyle || cell.fontColor ||
            cell.underline || cell.strikethrough || cell.textAlign || cell.alignItems ||
            cell.wrap || cell.letterSpacing || cell.lineSpacing
        );
    }

    /**
     * 在指定区域中查找第一个有内容或样式的单元格作为源单元格
     * @param {number} minCol - 最小列
     * @param {number} maxCol - 最大列
     * @param {number} minRow - 最小行
     * @param {number} maxRow - 最大行
     * @returns {Cell | null} 源单元格的副本，若无则返回 null
     */
    private findSourceCell(minCol: number, maxCol: number, minRow: number, maxRow: number): Cell | null {
        for (let row = minRow; row <= maxRow; row++) {
            for (let col = minCol; col <= maxCol; col++) {
                const cellName = this.getCellName(col, row);
                const cell = this.getCellByName(cellName);
                if (cell && this.hasCellContent(cell)) {
                    return cell;
                }
            }
        }
        return null;
    }

    /**
     * 汇总合并区域**外轮廓**上的边框（线宽 + 颜色）
     *
     * 合并单元格在视觉上就是一个大格：内部各格的边框（以及网格线）由渲染层的
     * 遮挡层盖掉，**只有落在区域外轮廓上的那些边**才画出来（Excel 语义）。
     *
     * 【纯查询，勿改回「写进锚点」】本函数只**读**区域内各格的边框并聚合，不写
     * 任何数据 —— 合并因此不必改锚点的边框字段，取消合并也就不存在「还原」问题：
     * 各格自己的边框一直原样躺在 `values` 里，只需清掉 colspan/rowspan，主屏的
     * 画线函数（`drawCellBorder` + `drawGridLines`）就会把它们连续地画出来。
     *
     * 同一条边上多个格子线宽不一致时取第一个非空值（一条边只有一组字段，
     * 无法表达「逐段不同」）。
     * @param {number} minCol - 起始列（1-based）
     * @param {number} maxCol - 结束列（1-based，含）
     * @param {number} minRow - 起始行（1-based）
     * @param {number} maxRow - 结束行（1-based，含）
     * @returns {{ top?: number; bottom?: number; left?: number; right?: number; color?: string }}
     *   四条边的线宽（无边框的边为 undefined）与轮廓边框色
     */
    public getMergedOutlineBorders(minCol: number, maxCol: number, minRow: number, maxRow: number):
        { top?: number; bottom?: number; left?: number; right?: number; color?: string } {
        const outline: { top?: number; bottom?: number; left?: number; right?: number; color?: string } = {};
        const noteColor = (c: Cell): void => {
            if (outline.color === undefined && c.borderColor) outline.color = c.borderColor;
        };
        for (let col = minCol; col <= maxCol && outline.top === undefined; col++) {
            const c = this.getCellByName(this.getCellName(col, minRow));
            if (c && c.borderTopWidth) { outline.top = c.borderTopWidth; noteColor(c); }
        }
        for (let col = minCol; col <= maxCol && outline.bottom === undefined; col++) {
            const c = this.getCellByName(this.getCellName(col, maxRow));
            if (c && c.borderBottomWidth) { outline.bottom = c.borderBottomWidth; noteColor(c); }
        }
        for (let row = minRow; row <= maxRow && outline.left === undefined; row++) {
            const c = this.getCellByName(this.getCellName(minCol, row));
            if (c && c.borderLeftWidth) { outline.left = c.borderLeftWidth; noteColor(c); }
        }
        for (let row = minRow; row <= maxRow && outline.right === undefined; row++) {
            const c = this.getCellByName(this.getCellName(maxCol, row));
            if (c && c.borderRightWidth) { outline.right = c.borderRightWidth; noteColor(c); }
        }
        // 【锚点兜底，勿删】对合并区里**任一格**设边框都会重定向到锚点
        // （见 ensureWritableCell），于是「给整个合并块加外框」最终只在锚点 D3 上
        // 留下四条边框；而上面四段扫描分别只认 minRow / maxRow / minCol / maxCol
        // 上的格子 —— 锚点若不在那条边上（bottom 认第 5 行、right 认 F 列），
        // 就永远扫不到，合并块会缺下边框 / 右边框。故未取到时用锚点对应边兜底。
        // 纯加法：只填 undefined 的槽位，老数据（区域各格各自带边框）不受影响。
        if (outline.top === undefined || outline.bottom === undefined
            || outline.left === undefined || outline.right === undefined) {
            const anchor = this.getCellByName(this.getCellName(minCol, minRow));
            if (anchor) {
                if (outline.top === undefined && anchor.borderTopWidth) { outline.top = anchor.borderTopWidth; noteColor(anchor); }
                if (outline.bottom === undefined && anchor.borderBottomWidth) { outline.bottom = anchor.borderBottomWidth; noteColor(anchor); }
                if (outline.left === undefined && anchor.borderLeftWidth) { outline.left = anchor.borderLeftWidth; noteColor(anchor); }
                if (outline.right === undefined && anchor.borderRightWidth) { outline.right = anchor.borderRightWidth; noteColor(anchor); }
            }
        }
        return outline;
    }

    /**
     * 判断单元格是否携带「会被合并遮挡、且取消合并后需要复原」的绘制信息
     *
     * 即边框与底色 —— 这两样在合并态下被离屏遮挡层盖住，取消合并时要原样回来。
     * 用它决定合并时该格子是**留下**（保真）还是**删除**（纯空格子，留着只会让
     * 大表合并后 values 平白膨胀几十万条）。
     * @param {Cell} cell - 待判定的单元格
     * @returns {boolean} 有边框或底色时为 true
     * @private
     */
    private hasPaintStyle(cell: Cell): boolean {
        return !!(
            cell.backgroundColor ||
            cell.borderTopWidth || cell.borderBottomWidth ||
            cell.borderLeftWidth || cell.borderRightWidth
        );
    }

    /**
     * 按名字集合批量删除单元格（**一次**遍历、只失效一次索引）
     *
     * 【性能约定，勿回退为「逐个 splice + findIndex」】
     * 删除会让单元格索引世代失效，下一次 `getCellByName` 就要重建整份索引（O(N)）。
     * 原实现在循环里交替「查找 → splice」，取消合并 1000 格时重建了 999 次索引，
     * 5 万格实测 4.4 秒。这里改成「先收集名字 → 原地压缩一次」，重建次数恒为 1。
     * @param {Set<string>} names - 待删除的单元格名
     * @private
     */
    private _removeCellsByNames(names: Set<string>): void {
        if (names.size === 0) return;
        const values = this._values;
        let w = 0;
        for (let i = 0; i < values.length; i++) {
            const c = values[i];
            if (names.has(c.cell)) continue;
            values[w++] = c;
        }
        values.length = w;
        this.invalidateCellIndex();
    }

    /**
     * 判断单元格是否有内容
     */
    private hasContent(cellName: string): boolean {
        const cell = this.getCellByName(cellName);
        return !!(cell && cell.chars && cell.chars.length > 0 && !cell.chars.every(v => v.char === ' '));
    }

    /**
     * 计算单元格的矩形区域
     * 对于合并单元格，返回整个合并区域的矩形
     */
    public calculateCellRect(col: number, row: number): { left: number, top: number, width: number, height: number } {
        // 行/列取消隐藏后尾部空行/列可能被收缩删除：clamp 到当前有效范围，
        // 防止激活单元格/编辑光标位于被删区域时 getAt(...).left/.top 读取 undefined
        const maxCol = Math.max(1, this.colHeaders.length);
        const maxRow = Math.max(1, this.rowHeaders.length);
        let realCol = Math.min(col, maxCol);
        let realRow = Math.min(row, maxRow);
        // 合并展开：命中合并区域则取左上角（O(1) 索引查找，见 findMergeRect）
        const mergeRect = this.findMergeRect(col, row);
        if (mergeRect !== null) {
            realCol = mergeRect.startCol;
            realRow = mergeRect.startRow;
        }
        
        const left = this.colHeaders.getAt(realCol - 1).left - this.colHeaders.offsetWidth;
        const top = this.rowHeaders.getAt(realRow - 1).top - this.rowHeaders.offsetHeight;
        let width = this.colHeaders.getAt(realCol - 1).width;
        let height = this.rowHeaders.getAt(realRow - 1).height;
        
        const cellName = this.getCellName(realCol, realRow);
        const cell = this.getCellByName(cellName);
        if (cell) {
            const colspan = cell.colspan || 1;
            const rowspan = cell.rowspan || 1;
            
            if (colspan > 1) {
                width = 0;
                for (let c = 0; c < colspan; c++) {
                    const ch = this.colHeaders.getAt(realCol + c - 1);
                    // 跳过隐藏列：选区/单元格宽度只累计可见列
                    if (ch && !ch.isHidden) width += ch.width;
                }
            }

            if (rowspan > 1) {
                height = 0;
                for (let r = 0; r < rowspan; r++) {
                    const rh = this.rowHeaders.getAt(realRow + r - 1);
                    // 跳过隐藏行：选区/单元格高度只累计可见行
                    if (rh && !rh.isHidden) height += rh.height;
                }
            }
        }
        
        return { left, top, width, height };
    }

    /**
     * 获取可见列的起始索引和结束索引，从1开始计数
     *
     * 与旧实现（从最后一列向前回扫）结果等价，但改为**从可见起点向后推进**，
     * 复杂度 O(可见列数) 而非 O(总列数) —— 宽表上后者每帧都要空转上万次。
     *
     * 【约定，勿改回「下标」语义】`startCol` / `endCol` 都是 **1-based 列号**，
     * 且 `endCol` **包含「右边缘只露出一部分」的那一列**：视口右边界通常落在某列中间，
     * 该列必须参与绘制（内容会被画布边缘自然裁掉），否则右侧会缺一整列的文本/背景/边框。
     *
     * 历史实现把循环下标 `i`（0-based）直接写进了 `endCol`，而调用方全部按
     * 1-based 列号消费（`getCellName(c, r)` / `getVisibleMergeRects`），
     * 于是**视口最右一列永远不被返回**：它的内容、背景、边框都不绘制，
     * 只剩网格线 —— 表现为「右侧被裁一半的列没有内容」。
     * @returns {{startCol: number, endCol: number}} 可见列的起止列号（1-based，含端点）
     */
    public getVisibleCols(): {startCol: number, endCol: number} {
        const { startCol } = this.colHeaders.getVisibleStartCol();
        const len = this.colHeaders.length;
        const viewportRight = this.colHeaders.offsetWidth + this.visibleView.sheetWidth / this.zoom;
        // 二分找到第一个 left >= viewportRight 的列（0-based j），不存在则 j = len
        // left 单调不减（隐藏列 left = 前一个可见列 right，不累加），故可二分
        let lo = startCol - 1, hi = len - 1, j = len;
        while (lo <= hi) {
            const mid = (lo + hi) >> 1;
            const col = this.colHeaders.getAt(mid);
            if (col && col.left >= viewportRight) { j = mid; hi = mid - 1; }
            else lo = mid + 1;
        }
        return { startCol, endCol: Math.max(startCol, j) };
    }
    
    /**
     * 获取可见行的起始索引和结束索引，从1开始计数
     *
     * 与旧实现（从最后一行向前回扫）结果等价，但改为**从可见起点向后推进**，
     * 复杂度 O(可见行数) 而非 O(总行数) —— 高表上后者每帧都要空转上万次。
     *
     * 语义与 {@link getVisibleCols} 完全对称：`startRow` / `endRow` 都是
     * **1-based 行号**，`endRow` **包含「下边缘只露出一部分」的那一行**。
     *
     * 【性能约定，勿回退为「线性扫描跳过隐藏行」】
     * 筛选后视口内可能有大量隐藏行（top = 前一个可见行 bottom，不累加），
     * 线性扫描逐行 `continue` 会每帧跳过几万行。`top` 单调不减，改二分后 O(log N)。
     * @returns {{startRow: number, endRow: number}} 可见行的起止行号（1-based，含端点）
     */
    public getVisibleRows(): {startRow: number, endRow: number} {
        const { startRow } = this.rowHeaders.getVisibleStartRow();
        const len = this.rowHeaders.length;
        const viewportBottom = this.rowHeaders.offsetHeight + this.visibleView.sheetHeight / this.zoom;
        // 二分找到第一个 top >= viewportBottom 的行（0-based j），不存在则 j = len
        // top 单调不减（隐藏行 top = 前一个可见行 bottom，不累加），故可二分
        let lo = startRow - 1, hi = len - 1, j = len;
        while (lo <= hi) {
            const mid = (lo + hi) >> 1;
            const row = this.rowHeaders.getAt(mid);
            if (row && row.top >= viewportBottom) { j = mid; hi = mid - 1; }
            else lo = mid + 1;
        }
        // j = 第一个不在视口内的行（0-based）；视口最后一行行号 = j（1-based）
        // j 之前可能包含若干隐藏行（top < viewportBottom），由 getVisibleCells 跳过
        return { startRow, endRow: Math.max(startRow, j) };
    }

    /**
     * 获取视口内所有可见行的行号数组（1-based）
     *
     * 利用 `rowHeaders.visibleRowIndices` 在「纯可见行」上二分定位视口起止，
     * 只返回视口内的可见行行号。隐藏行不占位但行号不连续，按行号遍历会
     * 跳过大量隐藏行号；本方法返回纯可见行行号，调用方直接遍历无需跳过。
     *
     * 复杂度 O(log 可见行数 + 视口内可见行数)，与隐藏行数量无关。
     * @returns {number[]} 视口内可见行行号数组（1-based，升序）
     */
    public getVisibleRowNumbers(): number[] {
        const visIdx = this.rowHeaders.visibleRowIndices;
        if (visIdx.length === 0) return [];
        const offset = this.rowHeaders.offsetHeight;
        const viewportBottom = offset + this.visibleView.sheetHeight / this.zoom;
        const headers = this.rowHeaders;
        // 二分：在 visIdx 中找第一个 top >= viewportBottom 的（即第一个不在视口内的可见行）
        let lo = 0, hi = visIdx.length - 1, j = visIdx.length;
        while (lo <= hi) {
            const mid = (lo + hi) >> 1;
            const row = headers.getAt(visIdx[mid]);
            if (row && row.top >= viewportBottom) { j = mid; hi = mid - 1; }
            else lo = mid + 1;
        }
        // 二分：在 visIdx 中找最后一个 top <= offset 的（即视口起始可见行）
        let lo2 = 0, hi2 = visIdx.length - 1, k = -1;
        while (lo2 <= hi2) {
            const mid = (lo2 + hi2) >> 1;
            const row = headers.getAt(visIdx[mid]);
            if (row && row.top <= offset) { k = mid; lo2 = mid + 1; }
            else hi2 = mid - 1;
        }
        // k = 最后一个 top <= offset 的可见行在 visIdx 中的位置
        // 视口起始可见行：top 覆盖 offset 的行（top <= offset 但 top + height > offset）
        // 与 getVisibleStartRow 语义一致：从 k 开始（k 的 top <= offset，k+1 的 top > offset）
        // 若 k 的 top + height <= offset（k 已滚出视口），起始应为 k+1，但 k+1 仍在 [startIdx, endIdx) 区间内
        // +1 冗余：最下一行多画一个可见行，避免内容紧挨视口底边看不清
        const startIdx = Math.max(0, k < 0 ? 0 : k);
        const endIdx = Math.min(visIdx.length, j + 1);
        const result: number[] = [];
        for (let i = startIdx; i < endIdx; i++) result.push(visIdx[i] + 1);
        return result;
    }

    /**
     * 获取视口内所有可见列的列号数组（1-based）
     *
     * 与 {@link getVisibleRowNumbers} 对称：在 `colHeaders.visibleColIndices`
     * 上二分定位视口起止，只返回视口内的可见列列号。
     * @returns {number[]} 视口内可见列列号数组（1-based，升序）
     */
    public getVisibleColNumbers(): number[] {
        const visIdx = this.colHeaders.visibleColIndices;
        if (visIdx.length === 0) return [];
        const offset = this.colHeaders.offsetWidth;
        const viewportRight = offset + this.visibleView.sheetWidth / this.zoom;
        const headers = this.colHeaders;
        // 二分：在 visIdx 中找第一个 left >= viewportRight 的（即第一个不在视口内的可见列）
        let lo = 0, hi = visIdx.length - 1, j = visIdx.length;
        while (lo <= hi) {
            const mid = (lo + hi) >> 1;
            const col = headers.getAt(visIdx[mid]);
            if (col && col.left >= viewportRight) { j = mid; hi = mid - 1; }
            else lo = mid + 1;
        }
        // 二分：在 visIdx 中找最后一个 left <= offset 的（即视口起始可见列）
        let lo2 = 0, hi2 = visIdx.length - 1, k = -1;
        while (lo2 <= hi2) {
            const mid = (lo2 + hi2) >> 1;
            const col = headers.getAt(visIdx[mid]);
            if (col && col.left <= offset) { k = mid; lo2 = mid + 1; }
            else hi2 = mid - 1;
        }
        // +1 冗余：最右一列多画一个可见列，避免内容紧挨视口右边看不清
        const startIdx = Math.max(0, k < 0 ? 0 : k);
        const endIdx = Math.min(visIdx.length, j + 1);
        const result: number[] = [];
        for (let i = startIdx; i < endIdx; i++) result.push(visIdx[i] + 1);
        return result;
    }

    /**
     * 获取当前视口内的单元格（视口裁剪的唯一入口）
     *
     * 渲染层据此只遍历「真正看得见」的单元格，而不是整张表：
     * - 视口内的普通单元格按「行 × 列」名称用 O(1) 索引取出；
     * - 额外补上**左上角在视口外、但合并区域与视口相交**的合并锚点，
     *   否则滚动到合并区域中部时该区域会被漏画。
     *
     * 复杂度 O(可见单元格数 + 合并单元格数)，与表格总规模无关。
     *
     * 【性能约定，勿回退为「按行号遍历跳过隐藏行/列」】
     * 筛选后 startRow..endRow 之间可能夹杂大量隐藏行（top 不累加但仍占行号）。
     * 按行号遍历即使每次 `continue`，几万隐藏行 × 几十列 = 每帧百万次循环判断，
     * 是筛选后滚动卡顿的主因。改为遍历 {@link getVisibleRowNumbers} /
     * {@link getVisibleColNumbers} 返回的「纯可见行/列号」，循环次数只与
     * 视口内可见单元格数相关，与隐藏行数量无关。
     * @returns {Cell[]} 视口内需要绘制的单元格
     */
    public getVisibleCells(): Cell[] {
        const rowNumbers = this.getVisibleRowNumbers();
        const colNumbers = this.getVisibleColNumbers();
        const result: Cell[] = [];
        const seen = new Set<string>();
        for (let ri = 0; ri < rowNumbers.length; ri++) {
            const r = rowNumbers[ri];
            for (let ci = 0; ci < colNumbers.length; ci++) {
                const c = colNumbers[ci];
                const name = this.getCellName(c, r);
                const cell = this.getCellByName(name);
                if (cell !== undefined) {
                    result.push(cell);
                    seen.add(name);
                }
            }
        }
        // 合并区域：锚点可能位于视口左上方（如 A1:C3 而视口从 B2 开始）
        // 走列倒排索引 + 数值裁剪（不再对每个合并块做字符串 split/解析）
        const { startCol, endCol } = this.getVisibleCols();
        const { startRow, endRow } = this.getVisibleRows();
        const mergeRects = this.getVisibleMergeRects(startCol, endCol, startRow, endRow);
        for (let i = 0; i < mergeRects.length; i++) {
            const anchorName = mergeRects[i].anchor;
            if (seen.has(anchorName)) continue;
            const anchor = this.getCellByName(anchorName);
            if (anchor !== undefined) {
                result.push(anchor);
                seen.add(anchorName);
            }
        }
        return result;
    }

    /**
     * 判断列索引是否超出可见范围
     * @returns {boolean} 是否超出可见范围
     */
    private isOutOfColBound(): boolean {
        return this.colHeaders.allColWidth <= this.colHeaders.offsetWidth + this.visibleView.sheetWidth / this.zoom;
    }
    
    /**
     * 判断行索引是否超出可见范围
     * @returns {boolean} 是否超出可见范围
     */
    private isOutOfRowBound(): boolean {
        return this.rowHeaders.allRowHeight <= this.rowHeaders.offsetHeight + this.visibleView.sheetHeight / this.zoom;
    }

    // ====================================== 合并单元格操作 ======================================
    /**
     * 清除单元格值和样式（用于剪切操作）
     */
    private clearCellValue(cell: string): void {
        const target = this.getCellByName(cell);
        if (target) {
            target.chars = [];
            target.backgroundColor = undefined;
            target.borderTopWidth = undefined;
            target.borderBottomWidth = undefined;
            target.borderLeftWidth = undefined;
            target.borderRightWidth = undefined;
            this.emit(DataEvents.CELL_CHANGED, cell, { index: 0, changed: ['value', 'background', 'border'], special: {} });
        }
    }

    /**
     * 切换单元格合并状态
     * @param {'col' | 'row'} [axis] - 合并方向（'col' 按列合并，'row' 按行合并）
     */
    public toggleMergeCells(axis?: 'col' | 'row'): void {
        if (!this.selection) return;
        
        if (axis === 'col') {
            this.mergeCellsByCol();
        } else if (axis === 'row') {
            this.mergeCellsByRow();
        } else {
            if (this.isMerged()) {
                this.unmergeCells();
            } else {
                this.mergeCells();
            }
        }
    }

    /**
     * 判断选中的单元格是否合并
     * @returns {boolean} 是否合并
     */
    public isMerged(): boolean {
        const [ startCell, endCell ] = this.selection.split(':');
        const mergedCells = this.getMergedCells();
        if (mergedCells.includes(this.selection)) {
            return true;
        }
        if (startCell === endCell) {
            // 几何判定（勿用字符串 includes，见 getMergedRangeAt 的说明）：
            // 该格被某个合并区域覆盖（不管它是锚点还是内部的被覆盖格）即算合并
            const { col, row } = this.getCellColAndRow(startCell);
            if (this.findMergeRect(col, row) !== null) return true;
            const cell = this.getCellByName(startCell);
            if (cell) {
                return cell.colspan > 1 || cell.rowspan > 1;
            }
        }
        return false;
    }

    /**
     * 合并单元格
     * 将当前选择的单元格合并为一个单元格
     * 按照从上到下、从左到右的原则，保留第一个有内容或样式的单元格内容和样式
     */
    public mergeCells(): void {
        if (!this.selection) return;
        
        const [startCell, endCell] = this.selection.split(':');
        const startPos = this.getCellColAndRow(startCell);
        const endPos = this.getCellColAndRow(endCell);
        if (startPos === endPos) return;
        
        const minCol = Math.min(startPos.col, endPos.col);
        const maxCol = Math.max(startPos.col, endPos.col);
        const minRow = Math.min(startPos.row, endPos.row);
        const maxRow = Math.max(startPos.row, endPos.row);
        
        const colspan = maxCol - minCol + 1;
        const rowspan = maxRow - minRow + 1;        
        if (colspan === 1 && rowspan === 1) return;

        // 检查选择区域中有多少个单元格有内容
        const cellsWithContent = this.countCellsWithContent(minCol, maxCol, minRow, maxRow);
        
        // 如果有多个单元格有内容，显示确认对话框
        if (cellsWithContent > 1) {
            this.messageBox.confirm(
                '合并单元格时，仅保留左上角的值，而放弃其他值。',
                () => {
                    this.executeMergeCells(minCol, maxCol, minRow, maxRow, colspan, rowspan);
                },
                () => {
                    // 取消，不做任何操作
                }
            );
        } else {
            this.executeMergeCells(minCol, maxCol, minRow, maxRow, colspan, rowspan);
        }
    }

    /**
     * 检查选择区域中有多少个单元格有内容
     */
    private countCellsWithContent(minCol: number, maxCol: number, minRow: number, maxRow: number): number {
        let count = 0;
        for (let row = minRow; row <= maxRow; row++) {
            for (let col = minCol; col <= maxCol; col++) {
                if (this.hasContent(this.getCellName(col, row))) count++;
            }
        }
        return count;
    }

    /**
     * 执行合并单元格操作
     */
    private executeMergeCells(minCol: number, maxCol: number, minRow: number, maxRow: number, colspan: number, rowspan: number): void {
        // 合并只影响选区内的格子 → 走「局部补丁撤销」：
        // 代价与选区大小成正比，不受大表阈值限制（整表快照在 10 万格以上会被关掉，
        // 那正是「大表里合并完 Ctrl+Z 没反应」的根因）。
        this.runWithCellPatchUndo('合并单元格', { minCol, maxCol, minRow, maxRow }, () => {
            // ── 阶段一：只读（索引处于稳态，getCellByName 全是 O(1)）──
            // 顺序极关键：一旦在下面「边改结构边按名查找」，每次 push/splice 都会
            // 让索引世代失效，后续每次查找都要重建整份索引（O(N)）→ O(选区 × N)。
            const source = this.findSourceCell(minCol, maxCol, minRow, maxRow);
            const sourceCopy = source ? source.clone() : null;
            const anchorName = this.getCellName(minCol, minRow);
            const anchor = this.getCellByName(anchorName);
            // 【数据保真，勿改回「整格删掉 + 事后还原」】
            // 合并区内部那些**带边框 / 底色**的格子必须原样留在 values 里：
            // 它们是「取消合并后线要连起来」的唯一依据。合并态下它们看不见，
            // 是渲染层用离屏画布遮挡掉的（见 sheetCanvas.drawMergedCellCover），
            // 不是在这里删掉的；取消合并只需清掉 colspan/rowspan，
            // 主屏画线函数就会把它们和网格线一起连续地画出来。
            const toRemove = new Set<string>();
            const keepers: Cell[] = [];
            for (let row = minRow; row <= maxRow; row++) {
                for (let col = minCol; col <= maxCol; col++) {
                    if (col === minCol && row === minRow) continue;
                    const name = this.getCellName(col, row);
                    const cell = this.getCellByName(name);
                    if (!cell) continue;
                    if (this.hasPaintStyle(cell)) keepers.push(cell);
                    else toRemove.add(name);
                }
            }

            // ── 阶段二：结构性修改（此后不再按名查找）──
            this._removeCellsByNames(toRemove);
            // 留下的格子只清「会在合并态里露出来的东西」：文字与筛选按钮。
            // 边框/底色是取消合并时要复原的信息，一个字节都不能动。
            for (let i = 0; i < keepers.length; i++) {
                keepers[i].chars = [];
                keepers[i].filter = undefined;
            }
            const target = anchor || sourceCopy || new Cell();
            if (anchor) {
                // 源格内容整体覆盖锚点（源格可能就是被清空的某个格子，故用克隆）
                if (sourceCopy) {
                    Object.assign(target, sourceCopy);
                    target.cell = anchorName;
                }
            } else {
                target.cell = anchorName;
                this._values.push(target);
            }
            target.colspan = colspan;
            target.rowspan = rowspan;
            this.invalidateCellIndex();
        });
    }

    /**
     * 粘贴 / 格式刷把「源合并单元格」的跨度传播到目标后，清掉目标块内被覆盖的格子
     *
     * 语义与 {@link executeMergeCells} 完全同口径（两处必须保持一致，否则
     * 「手动合并」与「粘出来的合并」行为会分叉）：
     * - 块内**带边框 / 底色**的格子留在 `values` 里，只清空文字与筛选按钮 ——
     *   它们是「取消合并后线能连起来」的唯一依据（详见 executeMergeCells 的说明）；
     * - 其余（只有文字或全空）的格子整条删除 —— 它们在合并态下永远被遮挡层盖住，
     *   留着只会让「按名查找 / 查找替换」扫到一条看不见的垃圾记录。
     *
     * 【两阶段分离】先只读收集，再一次性改 `_values`，最后只 `invalidateCellIndex()`
     * 一次；交替「remove + 按名查找」会让每格重建整份索引（O(N)）。
     *
     * 【必须在撤销作用域内调用】删除格子属于 `values` 改动，不落在补丁影响面里
     * 就撤不回来（大表下表现尤其明显：撤销后块内旧内容不回来）。
     *
     * @param anchors - 落点锚点及其**最终**跨度；`col`/`row` 须与 `name` 对应
     *                  （写入被重定向时是锚点自己的坐标，不是被覆盖位置的）
     */
    private _pruneCoveredCellsBySpans(
        anchors: { name: string; col: number; row: number; colspan: number; rowspan: number }[]
    ): void {
        if (anchors.length === 0) return;
        // 同一锚点可能被登记多次（源按 % 回绕平铺时），按名去重，保留首次跨度
        const list: { name: string; col: number; row: number; colspan: number; rowspan: number }[] = [];
        const anchorNames = new Set<string>();
        for (const a of anchors) {
            if (anchorNames.has(a.name)) continue;
            anchorNames.add(a.name);
            list.push(a);
        }
        // ── 阶段一：只读（索引稳态，getCellByName 全 O(1)）──
        const toRemove = new Set<string>();
        const keepers: Cell[] = [];
        for (const a of list) {
            const endCol = a.col + a.colspan - 1;
            const endRow = a.row + a.rowspan - 1;
            for (let row = a.row; row <= endRow; row++) {
                for (let col = a.col; col <= endCol; col++) {
                    if (col === a.col && row === a.row) continue;
                    const name = this.getCellName(col, row);
                    if (anchorNames.has(name)) continue; // 另一个落点锚点，不删
                    const cell = this.getCellByName(name);
                    if (!cell) continue;
                    if (this.hasPaintStyle(cell)) keepers.push(cell);
                    else toRemove.add(name);
                }
            }
        }
        if (toRemove.size === 0 && keepers.length === 0) {
            // 【勿删】跨度**已经写进锚点了**（H4 从此是 3×3 合并块），合并倒排索引必须重算 ——
            // 哪怕块内原本一个格子都没有（无需删除）：`_mergeRects` 由 cell 的 colspan/rowspan
            // 推导，而改跨度既不改 `_values.length` 也不置 `_indexStructurallyInvalid`，
            // `_ensureCellIndex()` 看不出变化 → `findMergeRect` 给的是旧结论。
            // 少这一句的后果：粘出来的合并块当帧不生效，选中框 / 渲染 / 后续
            // 写入重定向全部按「没合并」处理，直到下一次碰巧的整份索引重建。
            this.invalidateCellIndex();
            return;
        }
        // ── 阶段二：结构性修改（此后不再按名查找）──
        this._removeCellsByNames(toRemove);
        for (let i = 0; i < keepers.length; i++) {
            keepers[i].chars = [];
            keepers[i].filter = undefined;
        }
        this.invalidateCellIndex();
    }

    /**
     * 登记一个「落点最终带跨度」的锚点，供 {@link _pruneCoveredCellsBySpans} 结算
     *
     * 只在跨度 >1 时才解析坐标（普通粘贴是绝对多数路径，不该为它付 `getCellColAndRow` 的代价）。
     * @param sink - 收集数组
     * @param cell - 落点单元格（**已被写入最终跨度**）；重定向时 `cell.cell` 是锚点名
     */
    private _collectSpanAnchor(
        sink: { name: string; col: number; row: number; colspan: number; rowspan: number }[],
        cell: Cell
    ): void {
        const colspan = cell.colspan ?? 1;
        const rowspan = cell.rowspan ?? 1;
        if (colspan <= 1 && rowspan <= 1) return;
        const { col, row } = this.getCellColAndRow(cell.cell);
        if (!isFinite(col) || !isFinite(row) || col < 1 || row < 1) return;
        sink.push({ name: cell.cell, col, row, colspan, rowspan });
    }

    /**
     * 一段字符是否「单一样式」（不存在富文本多段样式）
     *
     * 用于粘贴时决定**能不能**刷单元格级字体：Cell 的字体 setter 写的是**全部**字符，
     * 源是多段富文本时刷一遍就等于把分段抹平（Excel 的「粘贴」保留分段）。
     * 空数组 / 单字符视为单一（无可分歧之处）。
     * @param {Char[] | undefined} chars - 源单元格的字符序列
     * @returns {boolean} 全部字符样式一致时 true
     */
    private _hasUniformCharStyle(chars: Char[] | undefined): boolean {
        if (!chars || chars.length <= 1) return true;
        const first = chars[0];
        for (let i = 1; i < chars.length; i++) {
            const c = chars[i];
            if (c.fontFamily !== first.fontFamily
                || c.fontSize !== first.fontSize
                || c.fontWeight !== first.fontWeight
                || c.fontStyle !== first.fontStyle
                || c.fontColor !== first.fontColor
                || c.underline !== first.underline
                || c.strikethrough !== first.strikethrough) return false;
        }
        return true;
    }

    /**
     * 取消合并单元格
     * 将当前选中的合并单元格恢复为独立的单元格
     * 合并区域的内容和文字样式保留在左上角单元格中
     * 背景色和边框会应用到所有分解出来的单元格
     */
    public unmergeCells(): void {
        if (!this.selection) return;
        const [startCell] = this.selection.split(':');
        const cell = this.getCellByName(startCell);
        
        let topLeftCell: Cell | null = null;
        let topLeftCol = this.getCellColAndRow(startCell).col;
        let topLeftRow = this.getCellColAndRow(startCell).row;
        
        if (cell?.colspan === 0 && cell?.rowspan === 0) {
            // 锚点被合并吞掉（colspan/rowspan 归零）时，原点位置已无合并信息，
            // 只能回查合并表；同样必须按**行列号数值矩形**判定（见 getMergedRangeAt）
            const { col, row } = this.getCellColAndRow(startCell);
            const rect = this.findMergeRect(col, row);
            if (rect !== null) {
                topLeftCol = rect.startCol;
                topLeftRow = rect.startRow;
                topLeftCell = this.getCellByName(rect.anchor) || null;
            }
        } else if (cell && (cell.colspan > 1 || cell.rowspan > 1)) {
            topLeftCell = cell;
        }
        
        if (!topLeftCell) return;
        
        const colspan = topLeftCell.colspan || 1;
        const rowspan = topLeftCell.rowspan || 1;
        
        if (colspan === 1 && rowspan === 1) return;

        const minCol = topLeftCol;
        const maxCol = topLeftCol + colspan - 1;
        const minRow = topLeftRow;
        const maxRow = topLeftRow + rowspan - 1;

        // 取消合并同样只影响原合并区 → 走局部补丁撤销（大表下也可 Ctrl+Z / Ctrl+Y）
        this.runWithCellPatchUndo('取消合并单元格', { minCol, maxCol, minRow, maxRow }, () => {
            // ── 阶段一：只读（索引稳态，getCellByName 全 O(1)）──
            const existing: Cell[] = [];
            for (let r = 0; r < rowspan; r++) {
                for (let c = 0; c < colspan; c++) {
                    const cell = this.getCellByName(this.getCellName(minCol + c, minRow + r));
                    if (cell) existing.push(cell);
                }
            }

            // ── 阶段二：结构性修改（此后不再按名查找）──
            // 【只清跨度，不下发任何样式】区域内各格的边框与底色一直原样留在
            // values 里（合并时就没有删它们，见 executeMergeCells 的说明），
            // 所以取消合并后**不需要**还原、也不需要补建缺失的格子：
            // 主屏的 drawCellBorder（逐格自己的边框）+ drawGridLines（整行整列的
            // 网格线）会把竖线横线连续地画出来 —— 这正是「取消合并后线断续」的
            // 根治办法。历史上在这里「按外轮廓下发 / 逐格四边下发」的两种写法
            // 都必然错一半（外框区域凭空长网格、网格区域丢内部线），已全部删除。
            for (let i = 0; i < existing.length; i++) {
                existing[i].colspan = undefined;
                existing[i].rowspan = undefined;
            }
            this.invalidateCellIndex();
        });
    }

    /**
     * 按列合并单元格
     * 将当前选择的单元格区域中每一列的单元格分别合并为一个单元格
     * 例如：选择 A1:C3 区域，会分别合并 A1:A3、B1:B3、C1:C3
     */
    public mergeCellsByCol(): void {
        if (!this.selection) return;
        
        const [startCell, endCell] = this.selection.split(':');
        const startPos = this.getCellColAndRow(startCell);
        const endPos = this.getCellColAndRow(endCell);
        if (startPos === endPos) return;
        
        const minCol = Math.min(startPos.col, endPos.col);
        const maxCol = Math.max(startPos.col, endPos.col);
        const minRow = Math.min(startPos.row, endPos.row);
        const maxRow = Math.max(startPos.row, endPos.row);
        
        const rowspan = maxRow - minRow + 1;
        if (rowspan <= 1) return;

        // 检查选择区域中有多少个单元格有内容
        const cellsWithContent = this.countCellsWithContent(minCol, maxCol, minRow, maxRow);
        
        // 如果有多个单元格有内容，显示确认对话框
        if (cellsWithContent > 1) {
            this.messageBox.confirm(
                '合并单元格时，仅保留每列第一个有值的单元格，而放弃其他值。',
                () => {
                    this.executeMergeCellsByCol(minCol, maxCol, minRow, maxRow, rowspan);
                },
                () => {
                    // 取消，不做任何操作
                }
            );
        } else {
            this.executeMergeCellsByCol(minCol, maxCol, minRow, maxRow, rowspan);
        }
    }

    /**
     * 执行按列合并单元格操作
     */
    private executeMergeCellsByCol(minCol: number, maxCol: number, minRow: number, maxRow: number, rowspan: number): void {
        // 同 {@link executeMergeCells}：只读阶段与结构修改阶段严格分离，
        // 索引重建次数恒为 1（历史实现是「每列各来一遍」）。
        this.runWithCellPatchUndo('按列合并单元格', { minCol, maxCol, minRow, maxRow }, () => {
            // ── 阶段一：只读 ──
            const anchors: Array<Cell | undefined> = [];
            const sources = new Map<string, Cell>();
            const anchorNames: string[] = [];
            const toRemove = new Set<string>();
            const keepers: Cell[] = [];
            for (let col = minCol; col <= maxCol; col++) {
                const anchorName = this.getCellName(col, minRow);
                anchorNames.push(anchorName);
                const source = this.findSourceCell(col, col, minRow, maxRow);
                if (source) sources.set(anchorName, source.clone());
                anchors.push(this.getCellByName(anchorName));
                for (let row = minRow; row <= maxRow; row++) {
                    if (row === minRow) continue;
                    const name = this.getCellName(col, row);
                    const cell = this.getCellByName(name);
                    if (!cell) continue;
                    if (this.hasPaintStyle(cell)) keepers.push(cell);
                    else toRemove.add(name);
                }
            }

            // ── 阶段二：结构性修改 ──
            this._removeCellsByNames(toRemove);
            // 同 {@link executeMergeCells}：带边框/底色的格子留着保真，只清文字与筛选
            for (let i = 0; i < keepers.length; i++) {
                keepers[i].chars = [];
                keepers[i].filter = undefined;
            }
            for (let i = 0; i < anchorNames.length; i++) {
                const name = anchorNames[i];
                const anchor = anchors[i];
                const sourceCopy = sources.get(name);
                const target = anchor || sourceCopy || new Cell();
                if (anchor) {
                    if (sourceCopy) {
                        Object.assign(target, sourceCopy);
                        target.cell = name;
                    }
                } else {
                    target.cell = name;
                    this._values.push(target);
                }
                target.colspan = 1;
                target.rowspan = rowspan;
            }
            this.invalidateCellIndex();
        });
    }

    /**
     * 按行合并单元格
     * 将当前选择的单元格区域中每一行的单元格分别合并为一个单元格
     * 例如：选择 A1:C3 区域，会分别合并 A1:C1、A2:C2、A3:C3
     */
    public mergeCellsByRow(): void {
        if (!this.selection) return;
        
        const [startCell, endCell] = this.selection.split(':');
        const startPos = this.getCellColAndRow(startCell);
        const endPos = this.getCellColAndRow(endCell);
        if (startPos === endPos) return;
        
        const minCol = Math.min(startPos.col, endPos.col);
        const maxCol = Math.max(startPos.col, endPos.col);
        const minRow = Math.min(startPos.row, endPos.row);
        const maxRow = Math.max(startPos.row, endPos.row);
        
        const colspan = maxCol - minCol + 1;
        if (colspan <= 1) return;

        // 检查选择区域中有多少个单元格有内容
        const cellsWithContent = this.countCellsWithContent(minCol, maxCol, minRow, maxRow);
        
        // 如果有多个单元格有内容，显示确认对话框
        if (cellsWithContent > 1) {
            this.messageBox.confirm(
                '合并单元格时，仅保留每行第一个有值的单元格，而放弃其他值。',
                () => {
                    this.executeMergeCellsByRow(minCol, maxCol, minRow, maxRow, colspan);
                },
                () => {
                    // 取消，不做任何操作
                }
            );
        } else {
            this.executeMergeCellsByRow(minCol, maxCol, minRow, maxRow, colspan);
        }
    }

    /**
     * 执行按行合并单元格操作
     */
    private executeMergeCellsByRow(minCol: number, maxCol: number, minRow: number, maxRow: number, colspan: number): void {
        // 同 {@link executeMergeCellsByCol}，只是把「每列一个锚点」换成「每行一个锚点」
        this.runWithCellPatchUndo('按行合并单元格', { minCol, maxCol, minRow, maxRow }, () => {
            // ── 阶段一：只读 ──
            const anchors: Array<Cell | undefined> = [];
            const sources = new Map<string, Cell>();
            const anchorNames: string[] = [];
            const toRemove = new Set<string>();
            const keepers: Cell[] = [];
            for (let row = minRow; row <= maxRow; row++) {
                const anchorName = this.getCellName(minCol, row);
                anchorNames.push(anchorName);
                const source = this.findSourceCell(minCol, maxCol, row, row);
                if (source) sources.set(anchorName, source.clone());
                anchors.push(this.getCellByName(anchorName));
                for (let col = minCol; col <= maxCol; col++) {
                    if (col === minCol) continue;
                    const name = this.getCellName(col, row);
                    const cell = this.getCellByName(name);
                    if (!cell) continue;
                    if (this.hasPaintStyle(cell)) keepers.push(cell);
                    else toRemove.add(name);
                }
            }

            // ── 阶段二：结构性修改 ──
            this._removeCellsByNames(toRemove);
            // 同 {@link executeMergeCells}：带边框/底色的格子留着保真，只清文字与筛选
            for (let i = 0; i < keepers.length; i++) {
                keepers[i].chars = [];
                keepers[i].filter = undefined;
            }
            for (let i = 0; i < anchorNames.length; i++) {
                const name = anchorNames[i];
                const anchor = anchors[i];
                const sourceCopy = sources.get(name);
                const target = anchor || sourceCopy || new Cell();
                if (anchor) {
                    if (sourceCopy) {
                        Object.assign(target, sourceCopy);
                        target.cell = name;
                    }
                } else {
                    target.cell = name;
                    this._values.push(target);
                }
                target.colspan = colspan;
                target.rowspan = 1;
            }
            this.invalidateCellIndex();
        });
    }

    /**
     * 获取所有合并单元格的名称（形如 `A1:C3`）
     *
     * 结果来自随单元格索引一并维护的缓存 {@link _mergedCells}，
     * 复杂度为 O(1)（缓存失效后的首次调用为 O(N)）。
     * 原先每次调用都要全表扫描，而它又会被「单元格矩形计算」按单元格调用一次，
     * 构成第二处 O(N²)。
     *
     * **返回值是内部缓存数组的只读引用，请勿修改。**
     * @returns {string[]} 合并单元格的名称数组
     */
    public getMergedCells(): string[] {
        this._ensureCellIndex();
        return this._mergedCells;
    }

    /**
     * 获取全部合并区域的数值矩形（只读引用，请勿修改）
     *
     * 与 {@link getMergedCells} 同源（同一次索引重建产出），但已是行列号，
     * 供需要「遍历所有合并块并做数值裁剪」的调用方使用（如合并背景按视口绘制）。
     * @returns {MergeRect[]} 合并区域矩形数组
     */
    public getMergeRects(): MergeRect[] {
        this._ensureCellIndex();
        return this._mergeRects;
    }

    /**
     * 查找覆盖指定位置的合并区域
     *
     * 复杂度 O(该列上的合并块数)（跨列极大的整行合并另走 {@link _wideMergeRects} 线性兜底），
     * 与合并总数基本无关；相比「遍历全部合并块 + 字符串 split + 逐字符解析」的旧写法，
     * 这是 8000 处合并时主画布卡死的直接解药。
     * @param {number} col - 列号（1-based）
     * @param {number} row - 行号（1-based）
     * @returns {MergeRect | null} 覆盖该位置的合并区域；不在任何合并区域内时 null
     */
    public findMergeRect(col: number, row: number): MergeRect | null {
        this._ensureCellIndex();
        const list = this._mergeByCol.get(col);
        if (list !== undefined) {
            for (let i = 0; i < list.length; i++) {
                const rect = list[i];
                if (row >= rect.startRow && row <= rect.endRow) return rect;
            }
        }
        const wide = this._wideMergeRects;
        for (let i = 0; i < wide.length; i++) {
            const rect = wide[i];
            if (col >= rect.startCol && col <= rect.endCol
                && row >= rect.startRow && row <= rect.endRow) return rect;
        }
        return null;
    }

    /**
     * 取覆盖指定位置的合并区域**选区串**（如 `'B22:D24'`）；不在任何合并区域内时 `null`
     *
     * 【勿写成 `mergedCells.find(v => v.includes(name))`】—— 那是**字符串包含**而非几何判定：
     * `('B22:D24').includes('B2') === true`（`B22` 里就含 `B2`）、
     * `('B22:D24').includes('D2') === true`（`D24` 里就含 `D2`）。
     * 于是点第 2 行的 B2 / D2 会命中二十行开外毫不相关的合并块，选中框当场跑到
     * B22；反过来点合并块内部的 B23 却因为 `'B22:D24'.includes('B23') === false`
     * 而选不中。必须按行列号做**数值矩形**判定（走 {@link findMergeRect} 倒排索引）。
     * @param {number} col - 列号（1-based）
     * @param {number} row - 行号（1-based）
     * @returns {string | null} 覆盖该位置的合并区域选区串；不在任何合并区域内时 null
     */
    public getMergedRangeAt(col: number, row: number): string | null {
        const rect = this.findMergeRect(col, row);
        if (rect === null) return null;
        return `${rect.anchor}:${this.getColName(rect.endCol)}${rect.endRow}`;
    }

    /**
     * 粘贴 / 格式刷收尾：把选区对齐到「这次操作真正影响到的范围」
     *
     * 两件事，取**并集**（谁大听谁的）：
     * ① **整块粘贴范围**（Excel 行为）：粘完之后选中的是粘出来的那一片，而不是粘贴前
     *    那个可能毫不相干的旧选区。传入 {@link getPasteTargetRect} / {@link getBrushTargetRect}
     *    的矩形即可；矩形缺省（如「粘贴字符」这种只动活动格的路径）时本项不参与。
     * ② **活动格所在的整个合并块**：把 3×3 的 D3 粘到 H4 → H4 当场成为 H4:J6 的锚点，
     *    此时选区若只停在单个 H4，选中框只框住左上 1/9 格，紧接着的复制 / 删除 / 设格式
     *    也就跟着只作用于 1/9。落点本就在既有合并块内部时同理（写在 D3:F5 里的 E4，
     *    框必须是 D3:F5）。
     *
     * 【判定必须走数值矩形】用 {@link findMergeRect} 倒排索引 —— 勿写成
     * `mergedCells.find(v => v.includes(name))`（`'B22:D24'.includes('B2')` 的坑）。
     *
     * 【两者都落空时不动选区】没有传入矩形、活动格也不在任何合并区内 → 保持原选区。
     * @param {Object} [rect] - 本次操作影响到的矩形（粘贴落点 / 落刷影响面）
     * @returns {void}
     */
    public syncSelectionToMergedActive(
        rect?: { startCol: number; endCol: number; startRow: number; endRow: number }
    ): void {
        const active = this._activedCell;
        if (!active) return;
        const { col, row } = this.getCellColAndRow(active);
        if (!isFinite(col) || !isFinite(row) || col < 1 || row < 1) return;

        let startCol = Infinity, endCol = -Infinity, startRow = Infinity, endRow = -Infinity;
        const include = (c1: number, c2: number, r1: number, r2: number): void => {
            if (!isFinite(c1) || !isFinite(c2) || !isFinite(r1) || !isFinite(r2)) return;
            if (c1 < startCol) startCol = c1;
            if (c2 > endCol) endCol = c2;
            if (r1 < startRow) startRow = r1;
            if (r2 > endRow) endRow = r2;
        };
        const merged = this.findMergeRect(col, row);
        if (merged !== null) include(merged.startCol, merged.endCol, merged.startRow, merged.endRow);
        if (rect !== undefined) include(rect.startCol, rect.endCol, rect.startRow, rect.endRow);
        if (endCol < startCol || endRow < startRow) return;      // 两者都落空 → 保持原选区

        const range = `${this.getCellName(startCol, startRow)}:${this.getCellName(endCol, endRow)}`;
        if (range === this._selection) return;
        this.selection = range;
    }

    /**
     * 取覆盖该位置的合并区域**锚点名**（如 `'D3'`）；不在任何合并区内、或该位置**就是**锚点时 `null`
     *
     * 与 {@link getMergedRangeAt} 的分工：后者返回「整个区域」的选区串（`'D3:F5'`），
     * 供**选区展开**使用；这里只返回锚点，供**写入重定向**使用。
     * 判定同样走 {@link findMergeRect} 的数值矩形 —— 勿写成字符串包含
     * （`('B22:D24').includes('B2') === true`，理由见 {@link getMergedRangeAt}）。
     * @param {number} col - 列号（1-based）
     * @param {number} row - 行号（1-based）
     * @returns {string | null} 覆盖该位置的合并区域锚点名；不需要重定向时 null
     */
    public findMergeAnchorName(col: number, row: number): string | null {
        const rect = this.findMergeRect(col, row);
        if (rect === null) return null;
        if (col === rect.startCol && row === rect.startRow) return null;   // 自己就是锚点，无需重定向
        return rect.anchor;
    }

    /**
     * 该位置是否被合并区域**覆盖**（在某个合并区内，但不是锚点本身）
     * @param {number} col - 列号（1-based）
     * @param {number} row - 行号（1-based）
     * @returns {boolean} 是「被合并吞掉的格子」时 true
     */
    public isCoveredByMerge(col: number, row: number): boolean {
        return this.findMergeAnchorName(col, row) !== null;
    }

    /** @see resolveWriteTargetName(string) */
    public resolveWriteTargetName(name: string): string;
    /** @see resolveWriteTargetName(number, number) */
    public resolveWriteTargetName(col: number, row: number): string;
    /**
     * 解析「这次修改该落到哪个格子名上」—— 合并区域内的非锚点位置一律重定向到锚点
     *
     * 【为什么要重定向】D3 合并了 D3:F5（colspan=3 / rowspan=3）之后，D4 / D5 / E3 /
     * E4 / E5 / F3 / F4 / F5 在视觉上根本不存在：它们的底色、边框、文字全被
     * `drawMergedCellCover` 的离屏遮挡层盖掉。若照旧「按名取不到就 `new Cell` 再
     * push」，一次整片着色 / 刷格式就会往 `values` 里塞进成千上万条**永远看不见**
     * 的记录 —— 快照体积、`JSON.parse` 耗时、撤销补丁条数全部跟着膨胀，
     * 合并越多的表越明显。故这些位置的任何属性修改（值与样式）都落到锚点 D3。
     *
     * 两种入参（单元格名 / 行列号）只是入口差异，判定逻辑唯一：
     * 走 {@link findMergeAnchorName}，取不到锚点就原样返回。
     * @param {string | number} nameOrCol - 单元格名（如 `'E4'`）或列号（1-based）
     * @param {number} [row] - 行号（1-based）；第一个参数是列号时必须给出
     * @returns {string} 实际应写入的单元格名（未被覆盖时即入参本身）
     */
    public resolveWriteTargetName(nameOrCol: string | number, row?: number): string {
        if (typeof nameOrCol === 'string') {
            const pos = this.getCellColAndRow(nameOrCol);
            return this.findMergeAnchorName(pos.col, pos.row) || nameOrCol;
        }
        const col = nameOrCol;
        const r = row as number;
        return this.findMergeAnchorName(col, r) || this.getCellName(col, r);
    }

    /**
     * 取（必要时创建）「本次修改应落在的单元格」—— 返回活引用，改它即改表
     *
     * 合并区内的非锚点位置重定向到锚点（见 {@link resolveWriteTargetName}）：
     * - 锚点已在 `values` → 直接返回它，就地改；
     * - 锚点不在 `values` → 新建并 push（**只有锚点会被 push**，被覆盖的位置永不新建）。
     *
     * 【只读路径请勿使用】查询一律走 {@link getCellByName} / {@link getEffectiveCellAt}：
     * 本方法会在格子缺失时**凭空造一个**，把「这格是空的」查成「这格存在且为空」。
     * @param {string} name - 目标单元格名（**可以**是一个被合并覆盖的位置）
     * @returns {Cell} 应写入的单元格（活引用）
     */
    public ensureWritableCell(name: string): Cell {
        const target = this.resolveWriteTargetName(name);
        const hit = this.getCellByName(target);
        return hit !== undefined ? hit : this._createCell(target);
    }

    /**
     * 新建单元格并入 `values`（同时**增量登记**进名索引）
     *
     * 【勿回退为「只 push 不登记」】push 让 `_values.length` 与 `_indexedLength`
     * 拉开差距，下一次 `getCellByName` 就要重建整份索引（O(N)）；批量写入
     * （刷字体 / 刷边框 / 刷数字格式，一次动辄上万格）那就是「每格重建一次」。
     * 索引处于稳态时直接 `set` 一条并同步长度，代价 O(1)；本就失效时整体失效一次。
     *
     * 新格子没有 colspan/rowspan/filter，故合并倒排与筛选缓存都不受影响，
     * 无需 {@link invalidateCellIndex} 之外的额外失效。
     * @param {string} name - 单元格名
     * @returns {Cell} 新建的单元格（活引用）
     * @private
     */
    private _createCell(name: string): Cell {
        const cell = new Cell();
        cell.cell = name;
        this._values.push(cell);
        if (!this._indexStructurallyInvalid && this._values.length === this._indexedLength + 1) {
            this._cellIndex.set(name, cell);
            this._indexedLength = this._values.length;
        } else {
            this.invalidateCellIndex();
        }
        return cell;
    }

    /**
     * 求行/列头点击后**活动格**应落的位置（**跳过合并单元格**）
     *
     * 行头点击恒取第 1 列、列头点击恒取第 1 行，而第 1 行 / 第 1 列恰恰最常被整行整列合并
     * （表头横幅 `A1:D1`、首列纵向分组 `A5:A8`）。活动格一旦落进合并区会同时坏两件事：
     * - `getActivedRect()` 走**合并锚点语义**（`calculateCellRect`），在选中框上挖的透明洞
     *   会变成整个合并块那么大 —— 整行/整列的高亮被挖掉一大片，就是「选择框出问题」；
     * - 若活动格是被合并**覆盖**的格子（colspan/rowspan 归零），编辑光标、填充柄都无处安放。
     *
     * 故沿所点的那一行（那一列）**向后**找第一个「既不在任何合并区内、又未被隐藏」的位置。
     * 跳过隐藏行列是同一道理：活动格落在隐藏行列上时用户根本看不见它。
     * 整行/整列都被合并或隐藏时退回原位置 —— 活动格总得有个落点，
     * 且选区本身不受影响（选区仍按 `expandMerges=false` 精确地等于所点的整行/整列）。
     * @param {number} col - 点击处的列号（1-based；行头点击时恒为 1）
     * @param {number} row - 点击处的行号（1-based；列头点击时恒为 1）
     * @param {'row' | 'col'} axis - 'row' = 点了行头（沿该行横向找）；'col' = 点了列头（沿该列纵向找）
     * @returns {string} 活动格名称
     */
    public resolveHeaderActiveCell(col: number, row: number, axis: 'row' | 'col'): string {
        if (axis === 'row') {
            const last = this.colHeaders.length;
            for (let c = col; c <= last; c++) {
                if (this.isColHidden(c - 1)) continue;
                if (this.findMergeRect(c, row) !== null) continue;
                return this.getCellName(c, row);
            }
        } else {
            const last = this.rowHeaders.length;
            for (let r = row; r <= last; r++) {
                if (this.isRowHidden(r - 1)) continue;
                if (this.findMergeRect(col, r) !== null) continue;
                return this.getCellName(col, r);
            }
        }
        return this.getCellName(col, row);
    }

    /**
     * 取出与给定视口矩形相交的合并区域（供按视口绘制合并背景使用）
     *
     * 走列倒排索引：只遍历**可见列**上挂着的合并块，复杂度
     * O(可见列数 × 每列合并块数) —— 与「全表合并块数」无关。
     * 历史实现在这里对每个合并块调用一次含全量合并扫描的矩形计算，
     * 构成 O(合并数²)：8000 处合并时单帧 3.8 秒。
     * @param {number} startCol - 视口起始列号（1-based，含）
     * @param {number} endCol - 视口结束列号（1-based，含）
     * @param {number} startRow - 视口起始行号（1-based，含）
     * @param {number} endRow - 视口结束行号（1-based，含）
     * @returns {MergeRect[]} 与视口相交的合并区域（按锚点去重）
     */
    public getVisibleMergeRects(startCol: number, endCol: number, startRow: number, endRow: number): MergeRect[] {
        this._ensureCellIndex();
        const result: MergeRect[] = [];
        const seen = new Set<string>();
        const consider = (rect: MergeRect): void => {
            if (rect.endRow < startRow || rect.startRow > endRow) return;   // 行向不相交
            if (seen.has(rect.anchor)) return;
            seen.add(rect.anchor);
            result.push(rect);
        };
        for (let c = startCol; c <= endCol; c++) {
            const list = this._mergeByCol.get(c);
            if (list === undefined) continue;
            for (let i = 0; i < list.length; i++) consider(list[i]);
        }
        // 整行级合并块：列索引里没有它，单独按列区间判断
        const wide = this._wideMergeRects;
        for (let i = 0; i < wide.length; i++) {
            const rect = wide[i];
            if (rect.endCol < startCol || rect.startCol > endCol) continue;
            consider(rect);
        }
        return result;
    }

    /**
     * 获取单元格的合并后的列号和行号
     * 若请求的单元格在合并区域内（非左上角），返回左上角单元格的列号和行号
     * @param {string} cell - 单元格名称
     * @returns {{ col: number, row: number }} 合并单元格的左上角单元格的列号和行号
     */
    public getMergeColAndRow(cell: string): { col: number, row: number } {        
        const { col, row } = this.getCellColAndRow(cell);
        // 先直接查找单元格本身
        const cellData = this.getCellByName(cell);
        if (cellData && (cellData.colspan > 1 || cellData.rowspan > 1)) {
            return { col, row };
        }

        // 若不在合并单元格中，检查是否位于某个合并区域内（O(1) 索引查找）
        const mergeRect = this.findMergeRect(col, row);
        if (mergeRect !== null) {
            return { col: mergeRect.startCol, row: mergeRect.startRow };
        }
        return { col, row };
    }
    // ====================================== 以上是合并单元格操作 ======================================

    /**
     * 获取当前活动单元格的矩形区域
     * @returns {{ activedLeft: number, activedTop: number, activedWidth: number, activedHeight: number }} 活动单元格矩形区域
     */
    public getActivedRect(): { activedLeft: number, activedTop: number, activedWidth: number, activedHeight: number } {
        let { col, row } = this.getCellColAndRow(this.activedCell);
        const { left: activedLeft, top: activedTop, width: activedWidth, height: activedHeight } = this.calculateCellRect(col, row);
        return { activedLeft, activedTop, activedWidth, activedHeight };
    }

    /**
     * 获取选中的单元格名称
     * @returns {string[]} 选中的单元格名称数组
     */
    public getSelectedCells(): string[] {
        if (!this.selection) return [];
        const [start, end] = this.selection ? this.selection.split(':') : [this.activedCell, this.activedCell];
        const { col: startCol, row: startRow } = this.getCellColAndRow(start);
        const { col: endCol, row: endRow } = this.getCellColAndRow(end);
        const selectedCells: string[] = [];
        for (let col = startCol; col <= endCol; col++) {
            for (let row = startRow; row <= endRow; row++) {
                selectedCells.push(`${this.getColName(col)}${row}`);
            }
        }
        return selectedCells;
    }

    /**
     * 获取选中的矩形区域
     * @returns {{ left: number, top: number, width: number, height: number }} 选中的矩形区域
     */
    public getSelectedRect(): { left: number, top: number, width: number, height: number } {
        return this.getRangeRect(this.selection);
    }

    /**
     * 获取任意范围串对应的矩形区域
     *
     * 与 {@link getSelectedRect} 是同一套几何口径（网格矩形、跳过隐藏行列），
     * 供复制/剪切的流动虚线框复用：它必须跟着当前选区一起滚动/缩放，
     * 而不是把复制瞬间的屏幕坐标缓存下来。
     * @param {string | null} range - 范围串（如 'A1:C5'）；为空时退化为活动单元格
     * @returns {{ left: number, top: number, width: number, height: number }} 范围矩形区域
     */
    public getRangeRect(range: string | null): { left: number, top: number, width: number, height: number } {
        const [start, end] = range ? range.split(':') : [this.activedCell, this.activedCell];
        const { col: startCol0, row: startRow0 } = this.getCellColAndRow(start);
        const { col: endCol0, row: endRow0 } = this.getCellColAndRow(end);
        const minCol = Math.min(startCol0, endCol0);
        const maxCol = Math.max(startCol0, endCol0);
        const minRow = Math.min(startRow0, endRow0);
        const maxRow = Math.max(startRow0, endRow0);
        // 收缩到范围内的可见行列（隐藏项不占位，选区框不覆盖其尺寸）
        //
        // 【性能约定，勿回退为「从 min 到 max 全量扫一遍」】
        // 本方法每帧被 drawSelectedRect 调用，而选区可以大得离谱（全选 = A1:XFD1048576
        // → 106 万次迭代/帧）。改成**从两端向内找，命中即停**：无隐藏行列时一步命中
        // （绝大多数表格如此）；有隐藏时最多走「连续隐藏段」那么长，而不是整个选区。
        let startCol = -1;
        for (let c = minCol; c <= maxCol; c++) {
            if (!this.colHeaders.getAt(c - 1)?.isHidden) { startCol = c; break; }
        }
        let endCol = -1;
        for (let c = maxCol; c >= minCol; c--) {
            if (!this.colHeaders.getAt(c - 1)?.isHidden) { endCol = c; break; }
        }
        let startRow = -1;
        for (let r = minRow; r <= maxRow; r++) {
            if (!this.rowHeaders.getAt(r - 1)?.isHidden) { startRow = r; break; }
        }
        let endRow = -1;
        for (let r = maxRow; r >= minRow; r--) {
            if (!this.rowHeaders.getAt(r - 1)?.isHidden) { endRow = r; break; }
        }
        // 选区完全落在隐藏行列上：返回空矩形
        if (startCol < 0 || startRow < 0) {
            return { left: 0, top: 0, width: 0, height: 0 };
        }
        // 【勿回退为 calculateCellRect】
        // 选区矩形必须是「行列号本身对应的那块格子」的并集，端点的格子若落在合并区域内，
        // 由 calculateCellRect 取矩形会锚到合并区左上角并按 colspan/rowspan 放大 →
        // 整个高亮框被平移/撑大，跑到选区之外（行头/列头点击整行整列时最明显）。
        // 规范选区（见 setSelectedRange）必然完整包含所相交的合并区域，此时两种取法等价，
        // 故此处改用网格矩形：无论选区是否规范，高亮范围都恒等于所选行列范围。
        const { left: startLeft, top: startTop } = this.getCellGridRect(startCol, startRow);
        const { left: endLeft, top: endTop, width: endWidth, height: endHeight } = this.getCellGridRect(endCol, endRow);
        const width = endLeft + endWidth - startLeft;
        const height = endTop + endHeight - startTop;

        return { left: startLeft, top: startTop, width, height };
    }

    /**
     * 单元格的**网格矩形**：只按行列几何取位置与尺寸，**不做合并区域展开**
     *
     * 与 {@link calculateCellRect} 的唯一区别在合并区域上：后者命中合并区域时会把坐标锚到
     * 合并区左上角、并按 colspan/rowspan 累加宽高 —— 那是「绘制合并单元格本体」和
     * 「活动单元格高亮挖洞」需要的语义；而选区、边框这类「按行列号取矩形」的场景需要的是
     * 网格本身的几何，否则端点落在合并区内时会连带把矩形平移/放大（例：选区 = 第 2 行整行，
     * A 列存在 A1:A3 纵向合并 → 锚点到第 1 行且高度多算 2 行，高亮整体跑到选区上方）。
     * @param {number} col - 列号（1-based）
     * @param {number} row - 行号（1-based）
     * @returns {{ left: number, top: number, width: number, height: number }} 网格矩形（超出有效范围时收敛到边界；越界取不到则为全 0）
     */
    public getCellGridRect(col: number, row: number): { left: number, top: number, width: number, height: number } {
        const realCol = Math.min(Math.max(1, col), Math.max(1, this.colHeaders.length));
        const realRow = Math.min(Math.max(1, row), Math.max(1, this.rowHeaders.length));
        const colHeader = this.colHeaders.getAt(realCol - 1);
        const rowHeader = this.rowHeaders.getAt(realRow - 1);
        // 行/列数组可能在「切换/清空文件」后被收缩：取不到时返回空矩形，
        // 由调用方（drawSelectedRect）按零尺寸跳过绘制，避免读取 undefined 抛错
        if (!colHeader || !rowHeader) {
            return { left: 0, top: 0, width: 0, height: 0 };
        }
        return {
            left: colHeader.left - this.colHeaders.offsetWidth,
            top: rowHeader.top - this.rowHeaders.offsetHeight,
            width: colHeader.width,
            height: rowHeader.height
        };
    }

    /**
     * 设置选中的单元格范围
     *
     * ## 合并归一化（`expandMerges`）
     * 默认会把选区**向外扩张到完整包含**所有与之相交的合并区域，维持
     * 「选区永不半覆盖合并区域」这一不变式 —— 单元格点击与单元格拖选需要它
     * （否则选区端点落在合并区内部时，用户看到的高亮与「实际作用范围」是两回事）。
     *
     * 行头 / 列头点击是刻意的例外：用户点第 5 行的行头，要的就是「第 5 行这一整行」，
     * 这一行中间有没有纵向合并（如 A5:A8）与他的意图无关。若照旧扩张，选中的会变成
     * 第 5~8 行 —— 「查找范围」「复制的内容」「删除行」会全部跟着错一整块。
     * 故这两个入口显式传 `expandMerges = false`。
     * （高亮的正确性不依赖归一化：{@link getSelectedRect} 只用网格矩形取端点，
     * 不按合并锚点展开，见那里的说明。）
     * @param {number} startCol - 起始列索引
     * @param {number} startRow - 起始行索引
     * @param {number} endCol - 结束列索引
     * @param {number} endRow - 结束行索引
     * @param {boolean} [expandMerges=true] - 是否把选区扩张到完整包含相交的合并区域
     */
    public setSelectedRange(startCol: number, startRow: number, endCol: number, endRow: number, expandMerges: boolean = true) {
        let maxCol = Math.max(startCol, endCol);
        let minCol = Math.min(startCol, endCol);
        let maxRow = Math.max(startRow, endRow);
        let minRow = Math.min(startRow, endRow);
        if (expandMerges) {
            const mergeRects = this.getMergeRects();
            for (let i = 0; i < mergeRects.length; i++) {
                const rect = mergeRects[i];
                // 矩形相交即说明合并区域落在选区内（数值比较，不再逐格试探）
                if (rect.endCol < minCol || rect.startCol > maxCol) continue;
                if (rect.endRow < minRow || rect.startRow > maxRow) continue;
                if (rect.startCol < minCol) minCol = rect.startCol;
                if (rect.startRow < minRow) minRow = rect.startRow;
                if (rect.endCol > maxCol) maxCol = rect.endCol;
                if (rect.endRow > maxRow) maxRow = rect.endRow;
            }
        }
        if (this.startLocation !== MouseLocation.IN_ANCHOR) {
            this.selection = `${this.getColName(minCol)}${minRow}:${this.getColName(maxCol)}${maxRow}`;
            // this.anchorSelection = `${this.getColName(minCol)}${minRow}:${this.getColName(maxCol)}${maxRow}`;
        }
    }

    /**
     * 设置选中的单元格
     *
     * 起点在行头 / 列头上时（点击或拖选），产出的是**整行 / 整列选区**，与
     * {@link setSelectedRange} 的 `expandMerges` 说明一致：按行列号本身取，
     * 不按合并区域向外扩张（`false`）—— 用户点第 5 行头要的是第 5 行整行，
     * 不会被该行中间的合并块拖成第 5~8 行。
     * @param {string} location - 选中的单元格名称
     * @param {number} col - 单元格的列索引
     * @param {number} row - 单元格的行索引
     */
    public setSelection(location: string, col: number, row: number) {
        const { col: startCol, row: startRow } = this.getCellColAndRow(this.activedCell);
        // 起点在行头 / 列头 → 本次拖选产出整行 / 整列选区：不按合并区域扩张
        const inColHeader = this.startLocation === MouseLocation.IN_COL_HEADER;
        const inRowHeader = this.startLocation === MouseLocation.IN_ROW_HEADER;
        if (location.includes('inCol')) {
            if (inColHeader) {
                this.setSelectedRange(startCol, 1, col, this.rowHeaders.length, false);
            } else if (inRowHeader) {
                this.setSelectedRange(1, startRow, this.colHeaders.length, 1, false);
            } else {
                this.setSelectedRange(startCol, startRow, col, 1);
            }
        } else if (location.includes('inRow')) {
            if (inColHeader) {
                this.setSelectedRange(startCol, 1, 1, this.rowHeaders.length, false);
            } else if (inRowHeader) {
                this.setSelectedRange(1, startRow, this.colHeaders.length, row, false);
            } else {
                this.setSelectedRange(startCol, startRow, 1, row);
            }
        } else {
            if (inColHeader) {
                this.setSelectedRange(startCol, 1, col, this.rowHeaders.length, false);
            } else if (inRowHeader) {
                this.setSelectedRange(1, startRow, this.colHeaders.length, row, false);
            } else {
                this.setSelectedRange(startCol, startRow, col, row);
            }
        }
    }

    /**
     * 获取选中的单元格的锚点矩形区域
     * @returns {{ col: number, row: number, left: number, top: number, width: number, height: number }} 选中的单元格的锚点矩形区域
     */
    public getAnchorRect(): { col: number, row: number, left: number, top: number, width: number, height: number } {
        const endCell = this.selection!.split(':')[1] ? this.selection!.split(':')[1] : this.activedCell;
        const { col, row } = this.getCellColAndRow(endCell);
        // 锚点必须落在**选中框右下角**（见 getSelectedRect 的网格矩形取值），
        // 故同样按网格矩形取端点：用 calculateCellRect 会在端点落于合并区内时
        // 把锚点拖到合并区右下角，与画出来的选中框错位。
        const { left: endLeft, top: endTop, width: endWidth, height: endHeight } = this.getCellGridRect(col, row);
        const width = DEFAULT_SELECTED_BORDER_WIDTH * 2;
        const height = DEFAULT_SELECTED_BORDER_WIDTH * 2;
        const left = endLeft + endWidth - width + 2;
        const top = endTop + endHeight - height + 2;

        return { col, row, left, top, width, height };
    }

    /**
     * 调整单列列宽（宽度下限 24）
     * @param col 列索引（1-based）
     * @param newWidth 新列宽
     */
    public adjustColumnWidth(col: number, newWidth: number): void {
        this.adjustColumnWidths(col, col, newWidth);
    }

    /**
     * 批量调整列宽：把 [startCol, endCol] 内的列都设为同一宽度
     *
     * 供「选中多列后拖拽其中一列的分界线」使用——所有被选中的列按同一宽度设置。
     * 统一走 {@link ColHeaders.recalcColPositions}：一遍重算左缘位置与可见总宽
     * （`allColWidth`），后者是右侧补列与滚动条范围的依据；旧的逐列手写 left
     * 只挪位置、漏更新总宽，列宽缩小后无法判断是否需要补列。
     * @param {number} startCol - 起始列号（1-based，含）
     * @param {number} endCol - 结束列号（1-based，含）
     * @param {number} newWidth - 新列宽（自动钳到下限）
     */
    public adjustColumnWidths(startCol: number, endCol: number, newWidth: number): void {
        if (startCol < 1 || startCol > this.colHeaders.length) return;
        const width = Math.max(24, newWidth);
        const last = Math.min(endCol, this.colHeaders.length);
        for (let c = startCol; c <= last; c++) {
            const colHeader = this.colHeaders.getAt(c - 1);
            if (colHeader) colHeader.width = width;
        }
        this.colHeaders.recalcColPositions();
        this.emit(DataEvents.COLUMNS_CHANGED);
    }

    /**
     * 调整单行行高（高度下限 24）
     * @param row 行索引（1-based）
     * @param newHeight 新行高
     */
    public adjustRowHeight(row: number, newHeight: number): void {
        this.adjustRowHeights(row, row, newHeight);
    }

    /**
     * 批量调整行高：把 [startRow, endRow] 内的行都设为同一高度
     *
     * 供「选中多行后拖拽其中一行的分界线」使用，语义与
     * {@link adjustColumnWidths} 一致（统一 recalc，同步可见总高）。
     * @param {number} startRow - 起始行号（1-based，含）
     * @param {number} endRow - 结束行号（1-based，含）
     * @param {number} newHeight - 新行高（自动钳到下限）
     */
    public adjustRowHeights(startRow: number, endRow: number, newHeight: number): void {
        if (startRow < 1 || startRow > this.rowHeaders.length) return;
        const height = Math.max(24, newHeight);
        const last = Math.min(endRow, this.rowHeaders.length);
        for (let r = startRow; r <= last; r++) {
            const rowHeader = this.rowHeaders.getAt(r - 1);
            if (rowHeader) rowHeader.height = height;
        }
        this.rowHeaders.recalcRowPositions();
        this.emit(DataEvents.ROWS_CHANGED);
    }

    /**
     * 补齐/裁剪尾部行，使可见行总高与视口匹配
     *
     * 行高调整后调用：行高变小时下方会露白（补足空白行）；
     * 行高变大时尾部会多出整屏空行（裁剪掉，避免反复调整把行数越堆越多）。
     * @param {number} [protectedStart=0] - 本次操作涉及的行段起点（1-based，不参与尾部裁剪；0 表示无）
     * @param {number} [protectedEnd=0] - 本次操作涉及的行段终点（1-based，含；0 表示无）
     */
    public fillRowsToViewport(protectedStart: number = 0, protectedEnd: number = 0): void {
        const viewportHeight = this.visibleView.sheetHeight / this.zoom;
        if (viewportHeight <= 0) return;
        const protectedRows = new Set<number>();
        for (let r = protectedStart; r <= protectedEnd; r++) protectedRows.add(r);
        this.trimRowsToViewport(viewportHeight, protectedRows);
        const deficit = viewportHeight - this.rowHeaders.allRowHeight;
        if (deficit > 0) {
            const need = Math.ceil(deficit / DEFAULT_CELL_HEIGHT) + 1;
            this.rowHeaders.addRow(this.rowHeaders.length, this.rowHeaders.length + need);
            this.rowHeaders.recalcRowPositions();
        }
        this.emit(DataEvents.ROWS_CHANGED);
    }

    /**
     * 补齐/裁剪尾部列，使可见列总宽与视口匹配（与 {@link fillRowsToViewport} 同构）
     * @param {number} [protectedStart=0] - 本次操作涉及的列段起点（1-based，不参与尾部裁剪；0 表示无）
     * @param {number} [protectedEnd=0] - 本次操作涉及的列段终点（1-based，含；0 表示无）
     */
    public fillColsToViewport(protectedStart: number = 0, protectedEnd: number = 0): void {
        const viewportWidth = this.visibleView.sheetWidth / this.zoom;
        if (viewportWidth <= 0) return;
        const protectedCols = new Set<number>();
        for (let c = protectedStart; c <= protectedEnd; c++) protectedCols.add(c);
        this.trimColsToViewport(viewportWidth, protectedCols);
        const deficit = viewportWidth - this.colHeaders.allColWidth;
        if (deficit > 0) {
            const need = Math.ceil(deficit / DEFAULT_CELL_WIDTH) + 1;
            this.colHeaders.addCol(this.colHeaders.length, this.colHeaders.length + need);
            this.colHeaders.recalcColPositions();
        }
        this.emit(DataEvents.COLUMNS_CHANGED);
    }
    
    /**
     * 计算指定行之前的所有行的字符总数（即全局字符起始索引）
     * @param {number} lineIndex - 行索引
     * @returns {number} 该行之前所有行的字符总数
     * @private
     */
    public getPreLineCharsLength(lineIndex: number): number {
        let k = 0;
        for (let i = 0; i < lineIndex; i++) {
            k += this.lines[i].text.length;
        }
        return k;
    }
    
    /**
     * 获取选择区域的起始索引
     * @returns {number} 起始索引
     */
    public getSelectedCharsStartIndex(): number {
        if (!this.selectionOnEditting) return this.cursorAbsolutePosition;
        return this.toAbsoluteIndex(this.selectionOnEditting.start);
    }

    /**
     * 获取选择区域的结束索引
     * @returns {number} 结束索引
     */
    public getSelectedCharsEndIndex(): number {
        if (!this.selectionOnEditting) return this.cursorAbsolutePosition;
        return this.toAbsoluteIndex(this.selectionOnEditting.end);
    }

    /**
     * 当前复制/剪切的源范围（无则为 null）
     *
     * 记录的范围属于复制时所在的工作表：切换到其它工作表后对外返回 null，
     * 让流动虚线框只在原表显示；切回原表时再次返回该范围。
     * @returns {string | null} 源范围串
     */
    public get copyRange(): string | null {
        return this._copyRangeSheet === this._activedSheetName ? this._copyRange : null;
    }

    /**
     * 设置/清除复制源范围并广播
     *
     * 界面层（{@link CopyMarquee}）只监听这个事件来显示/隐藏流动虚线框，
     * 不订阅剪贴板内部数据，避免数据层与界面层互相依赖。
     * 设置时把范围绑定到当前工作表，之后切换表格只隐藏显示、不丢失标记。
     * @param {string | null} range - 源范围串；null 表示清除
     */
    public setCopyRange(range: string | null): void {
        if (range === null) {
            if (this._copyRange === null && this._copyRangeSheet === null) return;
            this._copyRange = null;
            this._copyRangeSheet = null;
            this.emit(DataEvents.COPY_RANGE_CHANGED, null);
            return;
        }
        if (this._copyRange === range && this._copyRangeSheet === this._activedSheetName) return;
        this._copyRange = range;
        this._copyRangeSheet = this._activedSheetName;
        this.emit(DataEvents.COPY_RANGE_CHANGED, range);
    }

    /**
     * 清除复制/剪切的流动虚线框（Esc）
     *
     * 只清界面标记，不清内部剪贴板 —— 与 Excel 一致：取消选取不影响已复制的内容。
     */
    public clearCopyRange(): void {
        this.setCopyRange(null);
    }

    /**
     * 复制选中的单元格内容
     *
     * 遍历选区时忽略隐藏的行/列：隐藏行列中的单元格不进入剪贴板，
     * 保证复制、格式刷等后续操作均不携带隐藏内容。
     */
    public copySelection(): void {
        if (!this.selection) return;
        const [start, end] = this.selection ? this.selection.split(':') : [this.activedCell, this.activedCell];
        const { col: startCol, row: startRow } = this.getCellColAndRow(start);
        const { col: endCol, row: endRow } = this.getCellColAndRow(end);
        // 记录选区内可见的相对偏移序列（升序），供粘贴/格式刷紧凑排列使用
        const visibleColOffsets: number[] = [];
        const visibleRowOffsets: number[] = [];
        for (let col = startCol; col <= endCol; col++) {
            if (!this.isColHidden(col - 1)) visibleColOffsets.push(col - startCol);
        }
        for (let row = startRow; row <= endRow; row++) {
            if (!this.isRowHidden(row - 1)) visibleRowOffsets.push(row - startRow);
        }
        const selectedCells: Cell[] = [];
        for (const colOff of visibleColOffsets) {
            for (const rowOff of visibleRowOffsets) {
                const _cell = this.getCellByName(`${this.getColName(colOff + startCol)}${rowOff + startRow}`);
                if ( _cell ) {
                    // ⚠️ 必须在**克隆体**上清筛选标记，不能直接改 _cell：
                    // `getCellByName` 返回的是 `values` 里的活引用，原地把 `filter`
                    // 置空等于「复制一下就把原表的筛选按钮删了」。因为这一路径既不改
                    // `values.length` 也不置 `_indexStructurallyInvalid`，`_filterCells`
                    // 缓存不会被刷新，按钮 DOM 仍留在屏上看似正常；之后任何一次
                    // 索引重建（编辑、插入行列、排序…）才会让它凭空消失 ——
                    // 表现为「筛选按钮有时候不显示」。
                    const snapshot = _cell.clone();
                    snapshot.filter = undefined; // 剪贴板不携带筛选条件
                    selectedCells.push(snapshot);
                }
            }
        }

        // 复制/剪切后源范围转为流动虚线框（Excel 行为），直到 Esc 或下一次复制/剪切
        this.setCopyRange(this.selection);
        this.clipboardManager.setClipboard({ type: 'cell', cells: selectedCells, sourceRange: this.selection, visibleColOffsets, visibleRowOffsets });
    }

    /**
     * 剪切选中的单元格内容
     *
     * 遍历选区时忽略隐藏的行/列：隐藏行列中的单元格既不进入剪贴板，
     * 也不会从数据集合中移除（保持原样不动）。
     */
    public cutSelection(): void {
        if (!this.selection) return;
        const [start, end] = this.selection ? this.selection.split(':') : [this.activedCell, this.activedCell];
        const { col: startCol, row: startRow } = this.getCellColAndRow(start);
        const { col: endCol, row: endRow } = this.getCellColAndRow(end);
        // 记录选区内可见的相对偏移序列（升序），供粘贴/格式刷紧凑排列使用
        const visibleColOffsets: number[] = [];
        const visibleRowOffsets: number[] = [];
        for (let col = startCol; col <= endCol; col++) {
            if (!this.isColHidden(col - 1)) visibleColOffsets.push(col - startCol);
        }
        for (let row = startRow; row <= endRow; row++) {
            if (!this.isRowHidden(row - 1)) visibleRowOffsets.push(row - startRow);
        }
        const selectedCells: Cell[] = [];
        for (const colOff of visibleColOffsets) {
            for (const rowOff of visibleRowOffsets) {
                const _cell = this.getCellByName(`${this.getColName(colOff + startCol)}${rowOff + startRow}`);
                if ( _cell ) {
                    this.values.splice(this.values.indexOf(_cell), 1); // 从数据集合中移除选中的单元格
                    selectedCells.push(_cell);
                }
            }
        }

        // 复制/剪切后源范围转为流动虚线框（Excel 行为），直到 Esc 或下一次复制/剪切
        this.setCopyRange(this.selection);
        this.clipboardManager.setClipboard({ type: 'cell', cells: selectedCells, sourceRange: this.selection, visibleColOffsets, visibleRowOffsets });
        this.commitValues();
    }

    /**
     * 从指定起始行/列开始按序收集可见偏移
     *
     * 隐藏行/列不占序列位置，返回相对 startPos 的升序偏移，
     * 供粘贴/格式刷把源可见位置与目标可见位置按顺序一一配对（紧凑排列）。
     * @param {number} startPos - 起始行/列号（1-based，含）
     * @param {number} endPos - 结束行/列号（1-based，含）；需要一直收集到表格末尾时传 Infinity
     * @param {number} count - 最多收集的可见位置数量
     * @param {boolean} isCol - true=收集列，false=收集行
     * @returns {number[]} 相对 startPos 的升序偏移数组
     * @private
     */
    private _collectVisibleOffsets(startPos: number, endPos: number, count: number, isCol: boolean): number[] {
        const result: number[] = [];
        const total = isCol ? this.colHeaders.length : this.rowHeaders.length;
        const last = Math.min(endPos, total);
        for (let pos = startPos; pos <= last && result.length < count; pos++) {
            if (isCol ? !this.isColHidden(pos - 1) : !this.isRowHidden(pos - 1)) {
                result.push(pos - startPos);
            }
        }
        return result;
    }

    /**
     * 粘贴选中的单元格内容
     *
     * 忽略隐藏的行/列并紧凑排列：源可见位置（复制时记录的偏移序列）
     * 与目标可见位置按顺序一一配对，隐藏行列不产生空位；
     * 目标可见位置不足时多余内容丢弃。
     */
    public pasteSelection(): void {
        const clipboard = this.clipboardManager.getClipboard();
        if (clipboard && clipboard.type === 'cell' && clipboard.cells) {
            const [start, end] = clipboard.sourceRange.split(':');
            const { col: startCol, row: startRow } = this.getCellColAndRow(start);
            const { col: endCol, row: endRow } = this.getCellColAndRow(end);
            const { col: activedCol, row: activedRow } = this.getCellColAndRow(this.activedCell);
            // 源可见偏移序列（复制时记录；旧剪贴板数据回退为全量序列）
            const colOffsets = clipboard.visibleColOffsets ?? Array.from({ length: endCol - startCol + 1 }, (_, i) => i);
            const rowOffsets = clipboard.visibleRowOffsets ?? Array.from({ length: endRow - startRow + 1 }, (_, i) => i);
            // 目标可见偏移序列：从激活单元格开始按序收集（隐藏行列被跳过，不产生空位）
            const targetColOffsets = this._collectVisibleOffsets(activedCol, Infinity, colOffsets.length, true);
            const targetRowOffsets = this._collectVisibleOffsets(activedRow, Infinity, rowOffsets.length, false);
            // 落点最终带跨度的锚点（源是合并单元格 / 目标本来就是合并块）
            const spanAnchors: { name: string; col: number; row: number; colspan: number; rowspan: number }[] = [];
            // 源可见位置与目标可见位置按序一一配对（紧凑排列）
            for (let i = 0; i < colOffsets.length && i < targetColOffsets.length; i++) {
                for (let j = 0; j < rowOffsets.length && j < targetRowOffsets.length; j++) {
                    const tgtCol = activedCol + targetColOffsets[i];
                    const tgtRow = activedRow + targetRowOffsets[j];
                    const item = clipboard.cells.find(c => c.cell === `${this.getColName(startCol + colOffsets[i])}${startRow + rowOffsets[j]}`);
                    if (item) {
                        // 【合并写入重定向】落在合并区**内部**（如 D3:F5 里的 E4）时，
                        // 照旧「按名取不到就 new Cell」只会造出一条永远被遮挡层盖住的
                        // 垃圾记录（整片粘贴 = 上千条）。一律经 ensureWritableCell
                        // 落到锚点；锚点不存在才新建，且只有它会被 push。
                        const targetName = this.getCellName(tgtCol, tgtRow);
                        const cell = this.ensureWritableCell(targetName);
                        cell.chars = [...item.chars.map(c => {
                            const _char = new Char();
                            _char.char = c.char;
                            _char.fontFamily = c.fontFamily;
                            _char.fontSize = c.fontSize;
                            _char.fontWeight = c.fontWeight;
                            _char.fontStyle = c.fontStyle;
                            _char.fontColor = c.fontColor;
                            _char.underline = c.underline;
                            _char.strikethrough = c.strikethrough;
                            return _char;
                        })];
                        cell.isEdit = item.isEdit;
                        // 【勿删】只在「写入落到本格 且 粘贴源自带跨度」时才改写跨度：
                        // ① 写入被重定向到锚点（E4 → D3）时，本格不是一个能独立定跨度的
                        //    位置，改它等于凭空改动整个合并块的几何；
                        // ② 源不带跨度时照抄 undefined 会把目标自己的合并拆掉 ——
                        //    「往合并块里粘一段普通内容」在 Excel 里不会取消合并。
                        if (cell.cell === targetName && (item.colspan !== undefined || item.rowspan !== undefined)) {
                            cell.colspan = item.colspan;
                            cell.rowspan = item.rowspan;
                        }
                        // 【字体】上方的 chars 复制已把**逐字符**样式带过来。这里的单元格级
                        // 字体只是补位，且**必须带守卫**：Cell 的字体 setter 写的是**全部**
                        // 字符（`_chars.forEach(...)`），无条件刷一遍会把源的多段富文本抹平
                        // —— 而「粘贴」在 Excel 里是保留分段样式的。源样式单一时刷与不刷等价。
                        if (this._hasUniformCharStyle(item.chars)) {
                            cell.fontFamily = item.fontFamily;
                            cell.fontSize = item.fontSize;
                            cell.fontWeight = item.fontWeight;
                            cell.fontStyle = item.fontStyle;
                            cell.fontColor = item.fontColor;
                            cell.underline = item.underline;
                            cell.strikethrough = item.strikethrough;
                        }
                        // 数字格式与边框颜色**不在 chars 上**（`Cell._numberFormat` /
                        // `_borderColor`），只能靠这里搬过去 —— 缺了这两行，「把合并单元格
                        // 粘到别处」会丢掉数字格式，正是本轮要补的洞。
                        cell.numberFormat = item.numberFormat;
                        cell.borderColor = item.borderColor;
                        cell.borderTopWidth = item.borderTopWidth;
                        cell.borderBottomWidth = item.borderBottomWidth;
                        cell.borderLeftWidth = item.borderLeftWidth;
                        cell.borderRightWidth = item.borderRightWidth;
                        cell.textAlign = item.textAlign;
                        cell.alignItems = item.alignItems;
                        cell.backgroundColor = item.backgroundColor;
                        cell.filter = item.filter;
                        cell.wrap = item.wrap;
                        cell.letterSpacing = item.letterSpacing;
                        cell.lineSpacing = item.lineSpacing;
                        // 登记跨度：源是合并单元格时目标也成为同尺寸合并块，
                        // 接下来要把块内被覆盖的格子清掉（否则旧内容被遮挡层盖住却仍在 values 里）
                        this._collectSpanAnchor(spanAnchors, cell);
                    }
                }
            }
            this._pruneCoveredCellsBySpans(spanAnchors);
        } else if (clipboard && clipboard.type === 'char' && clipboard.chars) {
            // 同上：活动格若落在合并区内部，写入落到锚点（ensureWritableCell）
            const cell = this.ensureWritableCell(this.activedCell);
            if (cell.isEdit) {
                cell.chars = clipboard.chars.map(c => {
                    const _char = new Char();
                    _char.char = c.char;
                    _char.fontFamily = c.fontFamily;
                    _char.fontSize = c.fontSize;
                    _char.fontWeight = c.fontWeight;
                    _char.fontStyle = c.fontStyle;
                    _char.fontColor = c.fontColor;
                    _char.underline = c.underline;
                    _char.strikethrough = c.strikethrough;
                    return _char;
                })
            }; // 粘贴选中的字符
        }
        // 收尾：选区扩成整块粘贴范围（Excel 行为），落点若成为合并块再并上整个合并块。
        // 字符型剪贴板只写活动格一个格子，不参与范围扩张（粘贴前的选择原样保留）。
        this.syncSelectionToMergedActive(
            (clipboard && clipboard.type === 'cell') ? (this.getPasteTargetRect() ?? undefined) : undefined);

        this.commitValues();
    }

    /**
     * 粘贴纯文本内容
     *
     * 忽略隐藏的行/列并紧凑排列：源可见位置与目标可见位置按顺序一一配对，
     * 隐藏行列不产生空位；目标可见位置不足时多余内容丢弃。
     */
    public pastePlainText(): void {
        const clipboard = this.clipboardManager.getClipboard();
        const [start, end] = clipboard.sourceRange.split(':');
        const { col: startCol, row: startRow } = this.getCellColAndRow(start);
        const { col: endCol, row: endRow } = this.getCellColAndRow(end);
        const { col: activedCol, row: activedRow } = this.getCellColAndRow(this.activedCell);
        // 源可见偏移序列（复制时记录；旧剪贴板数据回退为全量序列）
        const colOffsets = clipboard.visibleColOffsets ?? Array.from({ length: endCol - startCol + 1 }, (_, i) => i);
        const rowOffsets = clipboard.visibleRowOffsets ?? Array.from({ length: endRow - startRow + 1 }, (_, i) => i);
        // 目标可见偏移序列：从激活单元格开始按序收集（隐藏行列被跳过，不产生空位）
        const targetColOffsets = this._collectVisibleOffsets(activedCol, Infinity, colOffsets.length, true);
        const targetRowOffsets = this._collectVisibleOffsets(activedRow, Infinity, rowOffsets.length, false);
        // 源可见位置与目标可见位置按序一一配对（紧凑排列）
        for (let i = 0; i < colOffsets.length && i < targetColOffsets.length; i++) {
            for (let j = 0; j < rowOffsets.length && j < targetRowOffsets.length; j++) {
                const tgtCol = activedCol + targetColOffsets[i];
                const tgtRow = activedRow + targetRowOffsets[j];
                const item = clipboard.cells.find(c => c.cell === `${this.getColName(startCol + colOffsets[i])}${startRow + rowOffsets[j]}`);
                if (item) {
                    // 合并区内部的位置重定向到锚点（见 ensureWritableCell 的说明）
                    const cell = this.ensureWritableCell(this.getCellName(tgtCol, tgtRow));
                    cell.chars = [...item.chars.map(c => {
                        const _char = new Char();
                        _char.char = c.char;
                        _char.fontFamily = c.fontFamily;
                        _char.fontSize = c.fontSize;
                        _char.fontWeight = c.fontWeight;
                        _char.fontStyle = c.fontStyle;
                        _char.fontColor = c.fontColor;
                        _char.underline = c.underline;
                        _char.strikethrough = c.strikethrough;
                        return _char;
                    })];
                }
            }
        }
        // 同上：扩成整块粘贴范围 + 并上活动格所在的整个合并块
        this.syncSelectionToMergedActive(
            (clipboard && clipboard.type === 'cell') ? (this.getPasteTargetRect() ?? undefined) : undefined);

        this.commitValues();
    }

    /**
     * 只粘贴文本（值），不粘贴任何格式（样式/边框/对齐/合并等）。
     * 与 pasteSelection 的目标配对逻辑完全一致，但目标单元格只更新 chars 里的 char 字符，
     * 所有样式字段保留原值不变；剪贴板为 char 类型时同理。
     * 忽略隐藏的行/列并紧凑排列（与 pasteSelection 相同）。
     */
    public pasteTextOnly(): void {
        const clipboard = this.clipboardManager.getClipboard();
        if (clipboard && clipboard.type === 'cell' && clipboard.cells) {
            const [start, end] = clipboard.sourceRange.split(':');
            const { col: startCol, row: startRow } = this.getCellColAndRow(start);
            const { col: endCol, row: endRow } = this.getCellColAndRow(end);
            const { col: activedCol, row: activedRow } = this.getCellColAndRow(this.activedCell);
            const colOffsets = clipboard.visibleColOffsets ?? Array.from({ length: endCol - startCol + 1 }, (_, i) => i);
            const rowOffsets = clipboard.visibleRowOffsets ?? Array.from({ length: endRow - startRow + 1 }, (_, i) => i);
            const targetColOffsets = this._collectVisibleOffsets(activedCol, Infinity, colOffsets.length, true);
            const targetRowOffsets = this._collectVisibleOffsets(activedRow, Infinity, rowOffsets.length, false);
            for (let i = 0; i < colOffsets.length && i < targetColOffsets.length; i++) {
                for (let j = 0; j < rowOffsets.length && j < targetRowOffsets.length; j++) {
                    const tgtCol = activedCol + targetColOffsets[i];
                    const tgtRow = activedRow + targetRowOffsets[j];
                    const item = clipboard.cells.find(c => c.cell === `${this.getColName(startCol + colOffsets[i])}${startRow + rowOffsets[j]}`);
                    if (item) {
                        // 只拷贝字符内容，保留目标原有格式
                        const newChars = item.chars.map(c => {
                            const _char = new Char();
                            _char.char = c.char;
                            return _char;
                        });
                        // 合并区内部的位置重定向到锚点（见 ensureWritableCell 的说明）
                        this.ensureWritableCell(this.getCellName(tgtCol, tgtRow)).chars = newChars;
                    }
                }
            }
        } else if (clipboard && clipboard.type === 'char' && clipboard.chars) {
            this.ensureWritableCell(this.activedCell).chars = clipboard.chars.map(c => {
                const _char = new Char();
                _char.char = c.char;
                return _char;
            });
        }
        this.syncSelectionToMergedActive(
            (clipboard && clipboard.type === 'cell') ? (this.getPasteTargetRect() ?? undefined) : undefined);
        this.commitValues();
    }

    /**
     * 粘贴格式仅
     *
     * 忽略隐藏的行/列并紧凑排列：源可见位置与目标可见位置按顺序一一配对，
     * 隐藏行列不产生空位；目标可见位置不足时多余内容丢弃。
     */
    public pasteFormatOnly(): void {
        const clipboard = this.clipboardManager.getClipboard();
        const [start, end] = clipboard.sourceRange.split(':');
        const { col: startCol, row: startRow } = this.getCellColAndRow(start);
        const { col: endCol, row: endRow } = this.getCellColAndRow(end);
        const { col: activedCol, row: activedRow } = this.getCellColAndRow(this.activedCell);
        // 源可见偏移序列（复制时记录；旧剪贴板数据回退为全量序列）
        const colOffsets = clipboard.visibleColOffsets ?? Array.from({ length: endCol - startCol + 1 }, (_, i) => i);
        const rowOffsets = clipboard.visibleRowOffsets ?? Array.from({ length: endRow - startRow + 1 }, (_, i) => i);
        // 目标可见偏移序列：从激活单元格开始按序收集（隐藏行列被跳过，不产生空位）
        const targetColOffsets = this._collectVisibleOffsets(activedCol, Infinity, colOffsets.length, true);
        const targetRowOffsets = this._collectVisibleOffsets(activedRow, Infinity, rowOffsets.length, false);
        // 落点最终带跨度的锚点（源是合并单元格 / 目标本来就是合并块）
        const spanAnchors: { name: string; col: number; row: number; colspan: number; rowspan: number }[] = [];
        // 源可见位置与目标可见位置按序一一配对（紧凑排列）
        for (let i = 0; i < colOffsets.length && i < targetColOffsets.length; i++) {
            for (let j = 0; j < rowOffsets.length && j < targetRowOffsets.length; j++) {
                const tgtCol = activedCol + targetColOffsets[i];
                const tgtRow = activedRow + targetRowOffsets[j];
                const item = clipboard.cells.find(c => c.cell === `${this.getColName(startCol + colOffsets[i])}${startRow + rowOffsets[j]}`);
                if (item) {
                    // 合并区内部的位置重定向到锚点（见 ensureWritableCell 的说明）
                    const targetName = this.getCellName(tgtCol, tgtRow);
                    const cell = this.ensureWritableCell(targetName);
                    cell.isEdit = item.isEdit;
                    // 同 pasteSelection：只在「落到本格 且 源自带跨度」时才改写跨度
                    if (cell.cell === targetName && (item.colspan !== undefined || item.rowspan !== undefined)) {
                        cell.colspan = item.colspan;
                        cell.rowspan = item.rowspan;
                    }
                    // 【只粘贴格式】这里**必须**无条件刷单元格级字体：本方法刻意不触碰
                    // chars（见下方 exeBrush 单格分支的注释），字体只能靠 setter 写进
                    // 目标既有字符 —— 与 pasteSelection 不同（那边 chars 自带样式，刷会抹平分段）。
                    cell.fontFamily = item.fontFamily;
                    cell.fontSize = item.fontSize;
                    cell.fontWeight = item.fontWeight;
                    cell.fontStyle = item.fontStyle;
                    cell.fontColor = item.fontColor;
                    cell.underline = item.underline;
                    cell.strikethrough = item.strikethrough;
                    cell.borderTopWidth = item.borderTopWidth;
                    cell.borderBottomWidth = item.borderBottomWidth;
                    cell.borderLeftWidth = item.borderLeftWidth;
                    cell.borderRightWidth = item.borderRightWidth;
                    cell.textAlign = item.textAlign;
                    cell.alignItems = item.alignItems;
                    cell.backgroundColor = item.backgroundColor;
                    cell.filter = item.filter;
                    cell.wrap = item.wrap;
                    cell.letterSpacing = item.letterSpacing;
                    cell.lineSpacing = item.lineSpacing;
                    // 数字格式与边框颜色不在 chars 上，格式清单里漏了它们就等于「只粘贴格式」
                    // 之后数字格式丢失（与 pasteSelection 同步补齐）
                    cell.numberFormat = item.numberFormat;
                    cell.borderColor = item.borderColor;
                    // 同 pasteSelection：源是合并单元格时目标也成为同尺寸合并块
                    this._collectSpanAnchor(spanAnchors, cell);
                }
            }
        }
        // 清掉落点块内被覆盖的格子（旧内容若留在 values 里，会被遮挡层盖住却仍能被查到）
        this._pruneCoveredCellsBySpans(spanAnchors);
        // 落点成为合并块（或本就在某个合并块内）时，选中框要覆盖**整个**合并区域
        this.syncSelectionToMergedActive(
            (clipboard && clipboard.type === 'cell') ? (this.getPasteTargetRect() ?? undefined) : undefined);

        this.commitValues();
    }

    /**
     * 复制选中的字符内容
     */
    public copyChars(): void {
        const cell = this.getCellByName(this.activedCell);
        if (!cell) return;
        const startIndex = this.getSelectedCharsStartIndex();
        const endIndex = this.getSelectedCharsEndIndex();
        const chars = cell.chars.slice(startIndex, endIndex); // 复制选中的字符
        this.clipboardManager.setClipboard({ type: 'char', chars });
    }

    /**
     * 剪切选中的字符内容
     * 仅负责移除字符数据、写入剪贴板并广播 VALUES_CHANGED；光标的最终定位
     * （落在被删除区间起点）由调用方在重新测量 lines 之后统一处理。
     * @returns {number} 剪切完成后光标应处的绝对位置（被删除区间的起点）
     */
    public cutChars(): number {
        const cell = this.getCellByName(this.activedCell);
        if (!cell) return this.cursorAbsolutePosition;
        const startIndex = this.getSelectedCharsStartIndex();
        const endIndex = this.getSelectedCharsEndIndex();
        const chars = cell.chars.slice(startIndex, endIndex); // 复制选中的字符
        cell.chars.splice(startIndex, endIndex - startIndex); // 从单元格中移除选中的字符
        this.selectionOnEditting = null;
        this.clipboardManager.setClipboard({ type: 'char', chars: chars});
        this.commitValues();
        // 光标置于原所选字符之首的位置（即被删除区间的起点）
        return startIndex;
    }

    /**
     * 粘贴选中的字符内容
     * 有选区时替换选区内容，无选区时在光标位置插入。
     * 仅负责修改字符数据并广播 VALUES_CHANGED；光标的最终定位（光标落在
     * 新增字符串之后）由调用方在重新测量 lines 之后统一处理，以避免基于
     * 旧 lines 布局反算出错误的 lineIndex/charIndex。
     * @param {number} insertIndex - 插入位置（无选区时的光标位置）
     * @param {Char[]} chars - 要粘贴的字符数组
     * @returns {number} 粘贴完成后光标应处的绝对位置（插入起点 + 新增字符数）
     */
    public pasteChars(insertIndex: number, chars: Char[]): number {
        // 活动格若落在合并区内部，写入落到锚点（见 ensureWritableCell 的说明）
        const cell = this.ensureWritableCell(this.activedCell);
        const startIndex = this.getSelectedCharsStartIndex();
        const endIndex = this.getSelectedCharsEndIndex();
        const deleteCount = endIndex - startIndex;
        // 有选区时从选区起点替换，无选区时从光标位置插入
        const insertPos = deleteCount > 0 ? startIndex : insertIndex;
        if (cell.isEdit) {
            cell.chars.splice(insertPos, deleteCount, ...chars.map(c => {
                const _char = new Char();
                _char.char = c.char;
                _char.fontFamily = c.fontFamily;
                _char.fontSize = c.fontSize;
                _char.fontWeight = c.fontWeight;
                _char.fontStyle = c.fontStyle;
                _char.fontColor = c.fontColor;
                _char.underline = c.underline;
                _char.strikethrough = c.strikethrough;
                return _char;
            }));
        }
        // 光标落在粘贴内容之后
        const caretAbsolutePosition = insertPos + chars.length;
        this.selectionOnEditting = null;
        this.commitValues();
        return caretAbsolutePosition;
    }
    
    /**
     * 设置选中单元格的字体
     *
     * 撤销/重做：非编辑态下整体经 {@link runWithFullStateUndo}（'selection' 档）包裹 ——
     * 工具栏（字体/字号/粗斜体/下划线/删除线/颜色/对齐/换行/底色）在小表走整表快照、
     * 大表只按当前选区捕获稀疏格子补丁（影响面超 {@link LARGE_TABLE_PATCH_CELLS}
     * 时自动放弃记账），与右键菜单的选区类操作记账口径完全一致。
     * 编辑态（仅可能由活动格 cell.isEdit 分支进入）**裸执行**：字符级/单格撤销走
     * 独立的编辑内撤销栈，绝不能被全量/补丁记账包裹，否则会污染编辑栈
     * （工具栏在编辑态本就改调 {@link setSelectedCharsFont}，这里仅为防御）。
     * @param fontStyle 字体样式
     */
    public setSelectedCellsFont(fontStyle: FontStyle): void {
        const apply = (): void => {
            const selectedCells = this.getSelectedCells();
            selectedCells.forEach(_cell => {
                // 【合并写入重定向】选中整块合并区（D3:F5）时，getSelectedCells 会列出
                // D4/D5/E3…F5 八个「被吞掉」的位置。它们被遮挡层盖着，为它们新建 Cell
                // 只会制造永远看不见的记录（快照/撤销补丁/装载时间全部跟着膨胀），
                // 故一律经 ensureWritableCell 落到锚点 D3；锚点缺失才新建。
                const cell = this.ensureWritableCell(_cell);
                if(cell.isEdit) {
                    if (fontStyle.fontFamily !== cell.fontFamily && fontStyle.fontFamily !== undefined) cell.fontFamily = fontStyle.fontFamily;
                    if (fontStyle.fontSize !== cell.fontSize && fontStyle.fontSize !== undefined) cell.fontSize = fontStyle.fontSize;
                    if (fontStyle.fontWeight !== cell.fontWeight && fontStyle.fontWeight !== undefined) cell.fontWeight = fontStyle.fontWeight;
                    if (fontStyle.fontStyle !== cell.fontStyle && fontStyle.fontStyle !== undefined) cell.fontStyle = fontStyle.fontStyle;
                    if (fontStyle.fontColor !== cell.fontColor && fontStyle.fontColor !== undefined) cell.fontColor = fontStyle.fontColor;
                    if (fontStyle.underline !== cell.underline && fontStyle.underline !== undefined) cell.underline = fontStyle.underline;
                    if (fontStyle.strikethrough !== cell.strikethrough && fontStyle.strikethrough !== undefined) cell.strikethrough = fontStyle.strikethrough;
                    if (fontStyle.textAlign !== cell.textAlign && fontStyle.textAlign !== undefined) cell.textAlign = fontStyle.textAlign as TextAlign;
                    if (fontStyle.alignItems !== cell.alignItems && fontStyle.alignItems !== undefined) cell.alignItems = fontStyle.alignItems as VerticalAlign;
                    if (fontStyle.backgroundColor !== cell.backgroundColor && fontStyle.backgroundColor !== undefined) cell.backgroundColor = fontStyle.backgroundColor;
                    if (fontStyle.letterSpacing !== cell.letterSpacing && fontStyle.letterSpacing !== undefined) cell.letterSpacing = fontStyle.letterSpacing;
                    if (fontStyle.lineSpacing !== cell.lineSpacing && fontStyle.lineSpacing !== undefined) cell.lineSpacing = fontStyle.lineSpacing;
                    if (fontStyle.wrap !== cell.wrap && fontStyle.wrap !== undefined) cell.wrap = fontStyle.wrap;
                }
            });
            this.commitValues();
        };
        if (this.isEditting) {
            apply();
            return;
        }
        this.runWithFullStateUndo('设置字体格式', apply, 'selection');
        // 包裹期间 commitValues 的自动广播被抑制，这里补一次驱动重绘（app.ts 按 rAF 排程，重复发射会自动合并）
        this.emit(DataEvents.VALUES_CHANGED, this._values);
    }

    /**
     * 调整选中单元格的缩进（左缩进，单位 CSS px）
     *
     * 与 {@link setSelectedCellsFont} 的区别：字体样式是「设置绝对值」，重复点击
     * 结果不变；缩进是「按增量调整」，重复点击应当累积，故单独成方法。
     *
     * ## 去重合并锚点
     *
     * 选中整块合并区时，{@link getSelectedCells} 会列出 D4/D5/E3…F3 这些被吞掉的
     * 位置；它们经 {@link ensureWritableCell} 全部重定向到同一个锚点 D3。
     * 若逐名累加，锚点会被加 N 次（2×3 合并块 = 6 次），故先按**格子对象**去重再累加。
     *
     * ## 撤销
     *
     * 非编辑态整体经 `runWithFullStateUndo`（'selection' 档）包裹，口径与其它
     * 工具栏格式操作一致；编辑态裸执行，避免污染编辑内撤销栈。
     *
     * @param delta - 增量像素值（正数增加、负数减少）；结果小于 0 时归零
     */
    public changeSelectedCellsIndent(delta: number): void {
        const apply = (): void => {
            const targets = new Set<Cell>();
            this.getSelectedCells().forEach(_cell => targets.add(this.ensureWritableCell(_cell)));
            targets.forEach(cell => { cell.indent = (cell.indent ?? 0) + delta; });
            this.commitValues();
        };
        if (this.isEditting) {
            apply();
            return;
        }
        this.runWithFullStateUndo('调整缩进', apply, 'selection');
        // 包裹期间 commitValues 的自动广播被抑制，这里补一次驱动重绘
        this.emit(DataEvents.VALUES_CHANGED, this._values);
    }

    /**
     * 设置选中字符的字体
     *
     * 作用范围：
     * - 编辑态下有字符选区：仅作用于选区字符；
     * - 编辑态下只有光标（无选区）：作用于当前单元格的**全部字符**。
     *   否则点击字体按钮会静默无效，也无法撤销，用户感知为「字体改不了」。
     *
     * 两种分支都经过 {@link commitValues}，编辑态下自动进入编辑内撤销栈，
     * 可通过 Ctrl+Z / Ctrl+Y 撤销与重做。
     *
     * @param fontStyle 字体样式
     */
    public setSelectedCharsFont(fontStyle: FontStyle): void {
        let startIndex = this.getSelectedCharsStartIndex();
        let endIndex = this.getSelectedCharsEndIndex();
        const cell = this.getCellByName(this.activedCell);
        const chars = cell?.chars || [];
        // 无选区时退化为「整格字符」作用域
        if (startIndex === endIndex && chars.length > 0) {
            startIndex = 0;
            endIndex = chars.length;
        }
        for (let i = startIndex; i < endIndex; i++) {
            const char = chars[i];
            if (char) {
                if (fontStyle.fontFamily !== undefined) char.fontFamily = fontStyle.fontFamily;
                if (fontStyle.fontSize !== undefined) char.fontSize = fontStyle.fontSize;
                if (fontStyle.fontWeight !== undefined) char.fontWeight = fontStyle.fontWeight;
                if (fontStyle.fontStyle !== undefined) char.fontStyle = fontStyle.fontStyle;
                if (fontStyle.fontColor !== undefined) char.fontColor = fontStyle.fontColor;
                if (fontStyle.underline !== undefined) char.underline = fontStyle.underline;
                if (fontStyle.strikethrough !== undefined) char.strikethrough = fontStyle.strikethrough;
            }
        }
        this.commitValues();
    }

    /**
     * 把源单元格的格式应用到目标单元格（不含字符内容/值，仅复制样式）。
     * 复用于单格与多格两条分支，字段清单与 pasteFormatOnly / pasteSelection 完全对齐：
     * - 单元格级：colspan / rowspan / 四条边框 / 对齐方式 / 背景色 / filter / wrap / 字距行距
     *   以及单元格级镜像字体（fontFamily / fontSize / fontWeight / fontStyle / fontColor / underline / strikethrough）
     * - 字符级：目标 cell 的每个现有 char 都按源 cell 对应位置（若超出则取源末位）拷贝字体样式；
     *   目标无 chars 则不创建（避免凭空生成文本，保持格式刷"不复制内容"的约定）。
     *
     * @param {Cell} targetCell - 目标单元格（已存在或新建后传入）
     * @param {Cell} sourceCell - 源单元格（clipboard 中的）
     * @private
     */
    private _applyFormatToCell(targetCell: Cell, sourceCell: Cell): void {
        // —— 单元格级格式 ——
        targetCell.isEdit = sourceCell.isEdit;
        targetCell.colspan = sourceCell.colspan;
        targetCell.rowspan = sourceCell.rowspan;
        targetCell.fontFamily = sourceCell.fontFamily;
        targetCell.fontSize = sourceCell.fontSize;
        targetCell.fontWeight = sourceCell.fontWeight;
        targetCell.fontStyle = sourceCell.fontStyle;
        targetCell.fontColor = sourceCell.fontColor;
        targetCell.underline = sourceCell.underline;
        targetCell.strikethrough = sourceCell.strikethrough;
        targetCell.borderTopWidth = sourceCell.borderTopWidth;
        targetCell.borderBottomWidth = sourceCell.borderBottomWidth;
        targetCell.borderLeftWidth = sourceCell.borderLeftWidth;
        targetCell.borderRightWidth = sourceCell.borderRightWidth;
        targetCell.textAlign = sourceCell.textAlign;
        targetCell.alignItems = sourceCell.alignItems;
        targetCell.backgroundColor = sourceCell.backgroundColor;
        targetCell.filter = sourceCell.filter;
        targetCell.wrap = sourceCell.wrap;
        targetCell.letterSpacing = sourceCell.letterSpacing;
        targetCell.lineSpacing = sourceCell.lineSpacing;
        // 数字格式与边框颜色不在 chars 上：格式刷漏了它们，「把合并单元格的样式刷到别处」
        // 就丢掉数字格式（与 pasteSelection / pasteFormatOnly 同步补齐）
        targetCell.numberFormat = sourceCell.numberFormat;
        targetCell.borderColor = sourceCell.borderColor;

        // —— 字符级格式（仅对目标已存在字符逐个同步样式，不创建/删除字符）——
        const srcChars = sourceCell.chars ?? [];
        const dstChars = targetCell.chars ?? [];
        if (srcChars.length === 0 || dstChars.length === 0) return;
        for (let i = 0; i < dstChars.length; i++) {
            // 超出源长度时回退到源末位 char 的样式（与 Excel 行为一致：拖选多字符保持末位风格）
            const srcChar = srcChars[Math.min(i, srcChars.length - 1)];
            if (!srcChar || !dstChars[i]) continue;
            dstChars[i].fontFamily = srcChar.fontFamily;
            dstChars[i].fontSize = srcChar.fontSize;
            dstChars[i].fontWeight = srcChar.fontWeight;
            dstChars[i].fontStyle = srcChar.fontStyle;
            dstChars[i].fontColor = srcChar.fontColor;
            dstChars[i].underline = srcChar.underline;
            dstChars[i].strikethrough = srcChar.strikethrough;
        }
    }

    /**
     * 格式刷模式（单次模式或连续模式）
     * @type {'single' | 'continuous' | null}
     */
    public get brushMode(): 'single' | 'continuous' | null {
        return this._brushMode;
    }
    
    /**
     * 设置格式刷模式（单次模式或连续模式）
     * @param {'single' | 'continuous' | null} value - 新的格式刷模式
     */
    public set brushMode(value: 'single' | 'continuous' | null) {
        this._brushMode = value;
        this.emit(DataEvents.BRUSH_MODE_CHANGED, value);
    }


    /**
     * 执行格式刷（Excel 格式刷功能）
     * 使用剪贴板中最近一次 copySelection 保存的源范围 cells 作为样式来源，
     * 把当前 this.selection 覆盖的目标范围统一应用样式。支持：
     * - 目标=单格（start===end）：走 pasteFormatOnly 逻辑，同时额外补字符级样式。
     * - 目标=多格：目标按行列遍历，源格式按「目标相对范围 % 源尺寸」循环套用
     *   （与 Excel 相同：目标区域大于源区域时，源样式平铺回绕）。
     * 源与目标均忽略隐藏的行/列并紧凑排列：源可见位置与目标可见位置按序一一配对，
     * 隐藏行列不产生空位（隐藏位置既不作为样式来源，也不接收格式）。
     * 最后：单次模式（single）自动关闭 brushMode；连续模式（continuous）保持开启可继续刷。
     * 操作前后不会改动目标单元格的字符内容/值（仅样式），不存在的目标 cell 会被创建以承载格式。
     */
    public exeBrush(): void {
        if (!this.brushMode || !this.selection) return;

        const clipboard = this.clipboardManager.getClipboard();
        if (!clipboard || clipboard.type !== 'cell' || !clipboard.cells || !clipboard.sourceRange) return;

        const [sourceStart, sourceEnd] = clipboard.sourceRange.split(':');
        const { col: sStartCol, row: sStartRow } = this.getCellColAndRow(sourceStart);
        const { col: sEndCol, row: sEndRow } = this.getCellColAndRow(sourceEnd);
        // 源可见偏移序列（复制时记录；旧剪贴板数据回退为全量序列）
        const sourceColOffsets = clipboard.visibleColOffsets ?? Array.from({ length: sEndCol - sStartCol + 1 }, (_, i) => i);
        const sourceRowOffsets = clipboard.visibleRowOffsets ?? Array.from({ length: sEndRow - sStartRow + 1 }, (_, i) => i);
        const sourceW = sourceColOffsets.length;
        const sourceH = sourceRowOffsets.length;
        if (sourceW === 0 || sourceH === 0) return; // 源范围内无可见位置，无事可刷

        const [start, end] = this.selection.split(':');
        const { col: tStartCol, row: tStartRow } = this.getCellColAndRow(start);
        const { col: tEndCol, row: tEndRow } = this.getCellColAndRow(end);
        if (start === end) {
            // 单格目标：复用 pasteFormatOnly 覆盖单元格级字段，再在此基础上补字符级样式
            // （pasteFormatOnly 本身不触碰 chars；commitValues 由它内部自行调用一次）
            this.pasteFormatOnly();
            // 字符级格式补刀（循环取源首格，对应单目标单源最常见的使用路径）
            // 目标落在合并区内部时，pasteFormatOnly 已把格式写到锚点上 → 这里同样取锚点
            const targetCell = this.getCellByName(this.resolveWriteTargetName(start));
            const sourceCell = clipboard.cells.find(c => c.cell === `${this.getColName(sStartCol)}${sStartRow}`);
            if (targetCell && sourceCell && targetCell.chars && targetCell.chars.length > 0 && sourceCell.chars && sourceCell.chars.length > 0) {
                for (let i = 0; i < targetCell.chars.length; i++) {
                    const srcChar = sourceCell.chars[Math.min(i, sourceCell.chars.length - 1)];
                    if (!srcChar || !targetCell.chars[i]) continue;
                    targetCell.chars[i].fontFamily = srcChar.fontFamily;
                    targetCell.chars[i].fontSize = srcChar.fontSize;
                    targetCell.chars[i].fontWeight = srcChar.fontWeight;
                    targetCell.chars[i].fontStyle = srcChar.fontStyle;
                    targetCell.chars[i].fontColor = srcChar.fontColor;
                    targetCell.chars[i].underline = srcChar.underline;
                    targetCell.chars[i].strikethrough = srcChar.strikethrough;
                }
                this.commitValues();
            }
            if (this.brushMode === 'single') this.brushMode = null;
            return;
        }

        // 多格目标：目标可见位置与源可见网格按序配对，按源可见尺寸平铺回绕
        // （隐藏行/列不占序列位置，保证刷出的结果紧凑排列）
        const targetColOffsets = this._collectVisibleOffsets(tStartCol, tEndCol, Infinity, true);
        const targetRowOffsets = this._collectVisibleOffsets(tStartRow, tEndRow, Infinity, false);
        // 落点最终带跨度的锚点（源是合并单元格 / 目标本来就是合并块）
        const spanAnchors: { name: string; col: number; row: number; colspan: number; rowspan: number }[] = [];
        for (let i = 0; i < targetColOffsets.length; i++) {
            for (let j = 0; j < targetRowOffsets.length; j++) {
                const tCol = tStartCol + targetColOffsets[i];
                const tRow = tStartRow + targetRowOffsets[j];
                const targetCellName = `${this.getColName(tCol)}${tRow}`;
                // 源区域中对应偏移（按源可见尺寸回绕）
                const sourceCellName = `${this.getColName(sStartCol + sourceColOffsets[i % sourceW])}${sStartRow + sourceRowOffsets[j % sourceH]}`;
                const sourceCell = clipboard.cells.find(c => c.cell === sourceCellName);
                // 源偏移位置没有单元格（空格子）则不刷，与 Excel 保持一致
                if (!sourceCell) continue;

                // 【合并写入重定向】合并区内部的位置不单独立条目，格式刷到锚点上
                // （见 ensureWritableCell）。重定向时必须保住锚点的 colspan/rowspan：
                // _applyFormatToCell 会照抄源的跨度，而源通常不带跨度 → 合并被拆掉。
                const targetCell = this.ensureWritableCell(targetCellName);
                const spanColBefore = targetCell.colspan;
                const spanRowBefore = targetCell.rowspan;
                this._applyFormatToCell(targetCell, sourceCell);
                // 【勿删，三条分支的顺序也不能换】_applyFormatToCell 会照抄源的跨度：
                // ① 写入被重定向到锚点（E4 → D3）：锚点跨度是**既有合并块的几何**，
                //    绝不能被源的跨度改写 —— 否则刷一下整块合并的尺寸就变了；
                // ② 源不带跨度而目标原本是合并：照抄 undefined 会把目标自己的合并拆掉
                //    （「往合并块上刷普通格式」在 Excel 里不会取消合并）；
                // ③ 源自带跨度且落到本格：保留刚写入的源跨度 —— 这就是「把合并单元格的
                //    样式刷到别处时，目标也成为同尺寸合并块」，随后清理块内覆盖格。
                const sourceHasSpan = sourceCell.colspan !== undefined || sourceCell.rowspan !== undefined;
                if (targetCell.cell !== targetCellName) {
                    targetCell.colspan = spanColBefore;
                    targetCell.rowspan = spanRowBefore;
                } else if (!sourceHasSpan && (spanColBefore !== undefined || spanRowBefore !== undefined)) {
                    targetCell.colspan = spanColBefore;
                    targetCell.rowspan = spanRowBefore;
                }
                this._collectSpanAnchor(spanAnchors, targetCell);
            }
        }
        // 清掉落点块内被覆盖的格子（与 executeMergeCells 同口径）
        this._pruneCoveredCellsBySpans(spanAnchors);
        // 收尾：选区 = 用户圈定的目标选区 ∪ 落点所在的合并块。
        // 【勿用 getBrushTargetRect()】那是**撤销影响面**，刻意按源最大跨度向右下多撑开
        // 一格（宁可多克隆几个空格子也不能漏记）；拿它当选区会把框画到没刷到的地方。
        this.syncSelectionToMergedActive({ startCol: tStartCol, endCol: tEndCol, startRow: tStartRow, endRow: tEndRow });

        if (this.brushMode === 'single') this.brushMode = null;
        this.commitValues();
    }

    /**
     * 设置选中单元格的边框样式
     *
     * 撤销/重做：非编辑态经 {@link runWithFullStateUndo}（'selection' 档）包裹，
     * 小表整表快照、大表选区稀疏补丁（超 {@link LARGE_TABLE_PATCH_CELLS} 放弃记账）。
     * 「设置单元格格式」对话框已在外层包裹时本方法作为嵌套层自动折叠，
     * 只保留最外层一条撤销记录（嵌套深度机制保证不重复入栈）。
     * @param cellBorderStyle 边框样式
     */
    public setSelectedCellsBorder(cellBorderStyle: CellBorderStyle): void {
        const apply = (): void => {
            const selectedCells = this.getSelectedCells();
            const { col: startCol, row: startRow } = this.getCellColAndRow(selectedCells.at(0));
            const { col: endCol, row: endRow } = this.getCellColAndRow(selectedCells.at(-1));
            selectedCells.forEach(_cell => {
                const {col, row} = this.getCellColAndRow(_cell);
                // 用 boolean 标记本格四条边哪些需要本次设置；未标记的保留原值，避免误清空。
                // 旧实现把未赋值的 borderTopWidth 等当作 undefined 赋给 cell，会清掉选区内部
                // 已有宽度为 1 的边框，导致 BorderOuter/WideOuter/Bottom/Top/Left/Right
                // 这六种"只设外部/单边"的操作意外改写了内部单元格的边框样式。
                let setTop = false, setBottom = false, setLeft = false, setRight = false;
                let borderTopWidth: number | undefined, borderBottomWidth: number | undefined,
                    borderLeftWidth: number | undefined, borderRightWidth: number | undefined;
                if (cellBorderStyle === CellBorderStyle.BorderNone) {
                    setTop = setBottom = setLeft = setRight = true;
                    // 四条边都清空（undefined）
                } else if (cellBorderStyle === CellBorderStyle.BorderAll) {
                    setTop = setBottom = setLeft = setRight = true;
                    borderTopWidth = borderBottomWidth = borderLeftWidth = borderRightWidth = 1;
                } else if (cellBorderStyle === CellBorderStyle.BorderOuter) {
                    if (startRow === row) { setTop = true; borderTopWidth = 1; }
                    if (endRow === row) { setBottom = true; borderBottomWidth = 1; }
                    if (startCol === col) { setLeft = true; borderLeftWidth = 1; }
                    if (endCol === col) { setRight = true; borderRightWidth = 1; }
                } else if (cellBorderStyle === CellBorderStyle.BorderWideOuter) {
                    if (startRow === row) { setTop = true; borderTopWidth = 2; }
                    if (endRow === row) { setBottom = true; borderBottomWidth = 2; }
                    if (startCol === col) { setLeft = true; borderLeftWidth = 2; }
                    if (endCol === col) { setRight = true; borderRightWidth = 2; }
                } else if (cellBorderStyle === CellBorderStyle.BorderBottom) {
                    if (endRow === row) { setBottom = true; borderBottomWidth = 1; }
                } else if (cellBorderStyle === CellBorderStyle.BorderTop) {
                    if (startRow === row) { setTop = true; borderTopWidth = 1; }
                } else if (cellBorderStyle === CellBorderStyle.BorderLeft) {
                    if (startCol === col) { setLeft = true; borderLeftWidth = 1; }
                } else if (cellBorderStyle === CellBorderStyle.BorderRight) {
                    if (endCol === col) { setRight = true; borderRightWidth = 1; }
                }

                // 局部边框模式：本格四条边都不需要改时跳过，避免触碰选区内部格子的原有边框
                if (!setTop && !setBottom && !setLeft && !setRight) return;

                // 【合并写入重定向】同上：合并区内部的格子不单独立条目，四条边一律
                // 记到锚点上（外轮廓由 getMergedOutlineBorders 从锚点兜底取出，
                // 见该方法内的说明），否则「给整个合并块加外框」会只留下半个轮廓。
                const cell = this.ensureWritableCell(_cell);
                if (setTop) cell.borderTopWidth = borderTopWidth;
                if (setBottom) cell.borderBottomWidth = borderBottomWidth;
                if (setLeft) cell.borderLeftWidth = borderLeftWidth;
                if (setRight) cell.borderRightWidth = borderRightWidth;
            })
            this.commitValues();
        };
        if (this.isEditting) {
            apply();
            return;
        }
        this.runWithFullStateUndo('设置边框', apply, 'selection');
        // 包裹期间 commitValues 的自动广播被抑制，补一次驱动重绘（嵌套在外层记账内时仅多排程一帧，无副作用）
        this.emit(DataEvents.VALUES_CHANGED, this._values);
    }

    /**
     * 按所选列或行升序排序
     */
    public sortUp(): void {
        this._executeSort(true);
    }

    /**
     * 按所选列或行降序排序
     */
    public sortDown(): void {
        this._executeSort(false);
    }

    /**
     * 排序统一执行逻辑。
     * 根据选区判断是列排序（选区在同一列）还是行排序（选区在同一行），
     * 调用 MatrixSorter.sort() 统一接口，支持全量/连续/不连续三种模式。
     *
     * 撤销/重做（命令式排列，大表友好）：排序只是选定范围内的一次置换，
     * 不创建/销毁任何格子，因此不再像旧实现那样 cloneValues 两份整表
     * （含全部 chars 的深拷贝，百万格下单次数秒、上百 MB）。改为从
     * MatrixSorter 取出实际执行的排列（positions/origins 两条索引数组，
     * 排序被筛选锚点/合并一致性中止时为 null，则什么都不做），
     * 推一条 'sortPermutation' 记录；撤销/重做见 {@link _applySortPermutation}。
     *
     * @param {boolean} ascending - true=升序，false=降序
     * @private
     */
    private _executeSort(ascending: boolean): void {
        if (!this.selection) return;
        const [startCell, endCell] = this.selection.split(':');
        const { col: startCol, row: startRow } = this.getCellColAndRow(startCell);
        const { col: endCol, row: endRow } = this.getCellColAndRow(endCell);

        let sortConfig: SortConfig;
        if (startCol === endCol) {
            // 列排序：按选定列对选定行范围排序
            sortConfig = {
                axis: 'column',
                keyIndex: startCol - 1,
                ascending,
                startIndex: startRow - 1,
                endIndex: endRow - 1
            };
        } else if (startRow === endRow) {
            // 行排序：按选定行对选定列范围排序
            sortConfig = {
                axis: 'row',
                keyIndex: startRow - 1,
                ascending,
                startIndex: startCol - 1,
                endIndex: endCol - 1
            };
        } else {
            const label = ascending ? '升序' : '降序';
            this.messageBox.confirm(
                `执行${label}排列操作时，请选择同一列或同一行。`,
                () => {
                },
                () => {
                }
            );
            return;
        }

        const matrix = this.matrixValues();
        const sorter = this.matrixSorter.init(matrix);
        sorter.sort(sortConfig);
        // 中止（筛选锚点行/列、合并一致性失败、keyIndex 越界等）：矩阵未被改动，
        // 旧实现仍会 remap + 重赋 values（一次可观的无效遍历），这里直接收尾
        const permutation = sorter.lastPermutation;
        if (!permutation) return;
        const _matrix = sorter.getData();
        this.remapMatrixCellNames(_matrix);
        const sortedValues = _matrix.flat().filter((c): c is Cell => c !== undefined);

        // 抑制 setter 自动记录：撤销依据是排列命令，而非前后快照
        const prevSuppress = this._suppressValuesUndo;
        this._suppressValuesUndo = true;
        try {
            this.values = sortedValues;
        } finally {
            this._suppressValuesUndo = prevSuppress;
        }

        // 命令式撤销记录：只存两条索引数组，成本与参与排序的行/列数成正比，
        // 与整表格子数无关（数组拷一份入栈，sorter 实例上的引用后续会被复用）
        const anchor = this._captureAnchor();
        this.undoManager.push({
            type: 'sortPermutation',
            data: {
                axis: permutation.axis,
                positions: permutation.positions.slice(),
                origins: permutation.origins.slice()
            },
            timestamp: Date.now(),
            label: ascending ? '升序排序' : '降序排序',
            anchorBefore: anchor,
            anchorAfter: anchor
        });
        this._refreshBaseline();
        this.markContentChanged();
        this.emit(DataEvents.VALUES_CHANGED, this._values);
    }

    /**
     * 执行一次排序排列的撤销/重做回写
     *
     * 正向（重做）：positions[i] 放置当前 origins[i] 行/列；
     * 逆向（撤销）：交换两数组角色 —— origins[i] 放置当前 positions[i] 行/列。
     * 排列为同一集合上的双射，交换角色恰为逆置换。
     * @param {SortPermutationSnapshot} snap - 排序排列记录
     * @param {boolean} isUndo - true=撤销，false=重做
     * @private
     */
    private _applySortPermutation(snap: SortPermutationSnapshot, isUndo: boolean): void {
        this._isApplyingUndo = true;
        try {
            this._applyAxisPermutation(
                snap.axis,
                isUndo ? snap.origins : snap.positions,
                isUndo ? snap.positions : snap.origins
            );
            this._refreshBaseline();
            this.emit(DataEvents.VALUES_CHANGED, this._values);
        } finally {
            this._isApplyingUndo = false;
        }
    }

    /**
     * 沿指定轴对单元格做一次置换（排序撤销/重做的原语）
     *
     * 排列后 `destinations[i]`（0-based 绝对行/列号）位置上放着当前位于
     * `sources[i]` 的整行/整列格子 —— 实现只改格子名称里的行号或列号
     * （与排序主路径 remapMatrixCellNames 的改名语义完全一致），
     * 不移动对象、不增删格子。范围外的格子不动。
     * @param {'column' | 'row'} axis - 'column'=置换行（按列排序），'row'=置换列
     * @param {number[]} destinations - 目标位置序列（0-based）
     * @param {number[]} sources - 来源位置序列（0-based，与 destinations 等长）
     * @private
     */
    private _applyAxisPermutation(axis: 'column' | 'row', destinations: number[], sources: number[]): void {
        // 来源 → 目标 的改名映射；恒等项（位置不变的行/列）无需入表
        const move = new Map<number, number>();
        for (let i = 0; i < destinations.length; i++) {
            if (destinations[i] !== sources[i]) move.set(sources[i], destinations[i]);
        }
        if (move.size === 0) return;
        // 单遍扫描全部格子改名：O(实际格子数)，与排序自身的 matrix/remap 同阶
        for (const cell of this._values) {
            const { col, row } = this.getCellColAndRow(cell.cell);
            if (axis === 'column') {
                const to = move.get(row - 1);
                if (to !== undefined) cell.cell = this.getCellName(col, to + 1);
            } else {
                const to = move.get(col - 1);
                if (to !== undefined) cell.cell = this.getCellName(to + 1, row);
            }
        }
        // 名称大面积变化，单元格索引必须重建
        this.invalidateCellIndex();
        this.markContentChanged();
    }

    /**
     * 获取数据集合的矩阵表示。
     *
     * 从实际数据（this.values）推导矩阵维度（最大行号×最大列号），
     * 而非依赖渲染层的 rowHeaders/colHeaders（后者随视口滚动动态增长，
     * 与实际数据范围不一致），保证所有单元格都被纳入矩阵且位置正确。
     * 每个单元格放置在其行列号对应的网格位置（row→第一维, col→第二维），
     * 无内容的格子为 undefined。
     * @returns {(Cell | undefined)[][]} 数据集合的矩阵表示
     */
    public matrixValues(): (Cell | undefined)[][] {
        // 从实际数据推导矩阵维度，避免依赖渲染层 headers 导致维度与数据不匹配
        let maxRow = 0;
        let maxCol = 0;
        for (const cell of this.values) {
            const { col, row } = this.getCellColAndRow(cell.cell);
            if (row > maxRow) maxRow = row;
            if (col > maxCol) maxCol = col;
        }
        const rows = maxRow;
        const cols = maxCol;
        // 预分配密集二维数组，空位以 undefined 填充，避免稀疏数组导致的遍历/位置问题
        const matrix: (Cell | undefined)[][] = [];
        for (let r = 0; r < rows; r++) {
            matrix[r] = new Array(cols).fill(undefined);
        }
        for (const cell of this.values) {
            const { col, row } = this.getCellColAndRow(cell.cell);
            const r = row - 1;
            const c = col - 1;
            if (r < 0 || c < 0) continue;
            matrix[r][c] = cell;
        }
        return matrix;
    }

    /**
     * 排序后将矩阵中每个单元格的 cell 名称重映射为其在新矩阵中的坐标。
     * matrix[r][c] 的单元格名称更新为 getCellName(c+1, r+1)，
     * 保证后续 matrixValues() 能按新位置正确重建矩阵。
     * @param {(Cell | undefined)[][]} matrix - 排序后的矩阵（会被原地修改）
     * @private
     */
    private remapMatrixCellNames(matrix: (Cell | undefined)[][]): void {
        for (let r = 0; r < matrix.length; r++) {
            const row = matrix[r];
            if (!row) continue;
            for (let c = 0; c < row.length; c++) {
                const cell = row[c];
                if (cell) {
                    cell.cell = this.getCellName(c + 1, r + 1);
                }
            }
        }
    }


    /**
     * 删除指定范围的行
     *
     * - 行表头移除被删行，其后各行的 top 依次前移；
     * - values 中被删行上的单元格（含合并锚点）直接移除，对应合并随之消失；
     * - 位于删除区域下方的单元格整体上移（cell 名称重映射，内容与样式随对象保留）；
     * - 跨越删除区域的合并锚点收缩 rowspan。
     * @param {number} startRow - 起始行号（1-based，含）
     * @param {number} endRow - 结束行号（1-based，含）
     */
    public deleteRows(startRow: number, endRow: number): void {
        if (startRow < 1 || endRow < startRow || startRow > this.rowHeaders.length) return;
        const clampedEnd = Math.min(endRow, this.rowHeaders.length);
        const removeCount = clampedEnd - startRow + 1;
        if (removeCount <= 0) return;

        // 1. 行表头移除被删行，其后各行 top 按紧凑规则前移（隐藏行不占位，继续保持隐藏不显露）
        this.rowHeaders.rowHeaders.splice(startRow - 1, removeCount);
        this.rowHeaders.recalcRowPositions();

        // 2. values 重映射
        const survivors: Cell[] = [];
        for (const cell of this.values) {
            const { col, row } = this.getCellColAndRow(cell.cell);
            // 被删行上的单元格直接移除
            if (row >= startRow && row <= clampedEnd) continue;
            // 跨越删除区域的合并锚点收缩 rowspan
            if (cell.rowspan && cell.rowspan > 1) {
                const spanEnd = row + cell.rowspan - 1;
                const overlap = Math.min(spanEnd, clampedEnd) - Math.max(row, startRow) + 1;
                if (overlap > 0) cell.rowspan = cell.rowspan - overlap;
            }
            // 位于删除区域下方的单元格上移
            if (row > clampedEnd) {
                cell.cell = `${this.getColName(col)}${row - removeCount}`;
            }
            survivors.push(cell);
        }
        this.values = survivors;
        this.emit(DataEvents.VALUES_CHANGED, this._values);
    }

    /**
     * 删除指定范围的列
     *
     * - 列表头移除被删列，其后各列的 left 依次前移；
     * - values 中被删列上的单元格（含合并锚点）直接移除，对应合并随之消失；
     * - 位于删除区域右侧的单元格整体左移（cell 名称重映射，内容与样式随对象保留）；
     * - 跨越删除区域的合并锚点收缩 colspan。
     * @param {number} startCol - 起始列号（1-based，含）
     * @param {number} endCol - 结束列号（1-based，含）
     */
    public deleteCols(startCol: number, endCol: number): void {
        if (startCol < 1 || endCol < startCol || startCol > this.colHeaders.length) return;
        const clampedEnd = Math.min(endCol, this.colHeaders.length);
        const removeCount = clampedEnd - startCol + 1;
        if (removeCount <= 0) return;

        // 1. 列表头移除被删列，其后各列 left 按紧凑规则前移（隐藏列不占位，继续保持隐藏不显露）
        this.colHeaders.colHeaders.splice(startCol - 1, removeCount);
        this.colHeaders.recalcColPositions();

        // 2. values 重映射
        const survivors: Cell[] = [];
        for (const cell of this.values) {
            const { col, row } = this.getCellColAndRow(cell.cell);
            // 被删列上的单元格直接移除
            if (col >= startCol && col <= clampedEnd) continue;
            // 跨越删除区域的合并锚点收缩 colspan
            if (cell.colspan && cell.colspan > 1) {
                const spanEnd = col + cell.colspan - 1;
                const overlap = Math.min(spanEnd, clampedEnd) - Math.max(col, startCol) + 1;
                if (overlap > 0) cell.colspan = cell.colspan - overlap;
            }
            // 位于删除区域右侧的单元格左移
            if (col > clampedEnd) {
                cell.cell = `${this.getColName(col - removeCount)}${row}`;
            }
            survivors.push(cell);
        }
        this.values = survivors;
        this.emit(DataEvents.VALUES_CHANGED, this._values);
    }

    /**
     * 在指定位置插入若干整行（结构性插入，区别于 insertCellsRight 的单元格级右移）。
     *
     * - 行表头在 atRow 处插入 count 个默认高度的新行，其后各行 top 依次后移；
     * - values 中行号 >= atRow 的单元格整体下移 count（cell 名称重映射，内容与样式随对象保留）；
     * - 跨越插入位置的合并锚点 rowspan 扩展 count（合并随新行一起撑开）；
     * - 隐藏行同样下移（整行结构操作，隐藏行参与重映射，与 deleteRows 对称）。
     * @param {number} atRow - 插入位置（1-based，新行将占据 atRow..atRow+count-1）
     * @param {number} count - 插入行数
     */
    public insertRows(atRow: number, count: number): void {
        if (atRow < 1 || count < 1) return;
        // 1. 行表头插入新行（默认高度），重算 top
        const newHeaders: RowHeader[] = [];
        for (let i = 0; i < count; i++) newHeaders.push({ top: 0, height: DEFAULT_CELL_HEIGHT });
        this.rowHeaders.rowHeaders.splice(atRow - 1, 0, ...newHeaders);
        this.rowHeaders.recalcRowPositions();

        // 2. values 重映射：行号 >= atRow 的单元格下移 count；跨越插入点的合并扩展 rowspan
        for (const cell of this.values) {
            const { col, row } = this.getCellColAndRow(cell.cell);
            if (row >= atRow) {
                cell.cell = `${this.getColName(col)}${row + count}`;
            }
            if (cell.rowspan && cell.rowspan > 1 && row < atRow) {
                const spanEnd = row + cell.rowspan - 1;
                if (spanEnd >= atRow) cell.rowspan = cell.rowspan + count;
            }
        }
        this.invalidateCellIndex(); // 重命名了 cell.cell，索引名→对象映射失效
        this.emit(DataEvents.VALUES_CHANGED, this._values);
    }

    /**
     * 在指定位置插入若干整列（结构性插入）。
     *
     * - 列表头在 atCol 处插入 count 个默认宽度的新列，其后各列 left 依次后移；
     * - values 中列号 >= atCol 的单元格整体右移 count（cell 名称重映射）；
     * - 跨越插入位置的合并锚点 colspan 扩展 count；
     * - 隐藏列同样右移（与 deleteCols 对称）。
     * @param {number} atCol - 插入位置（1-based，新列将占据 atCol..atCol+count-1）
     * @param {number} count - 插入列数
     */
    public insertCols(atCol: number, count: number): void {
        if (atCol < 1 || count < 1) return;
        const newHeaders: ColHeader[] = [];
        for (let i = 0; i < count; i++) newHeaders.push({ left: 0, width: DEFAULT_CELL_WIDTH });
        this.colHeaders.colHeaders.splice(atCol - 1, 0, ...newHeaders);
        this.colHeaders.recalcColPositions();

        for (const cell of this.values) {
            const { col, row } = this.getCellColAndRow(cell.cell);
            if (col >= atCol) {
                cell.cell = `${this.getColName(col + count)}${row}`;
            }
            if (cell.colspan && cell.colspan > 1 && col < atCol) {
                const spanEnd = col + cell.colspan - 1;
                if (spanEnd >= atCol) cell.colspan = cell.colspan + count;
            }
        }
        this.invalidateCellIndex(); // 重命名了 cell.cell，索引名→对象映射失效
        this.emit(DataEvents.VALUES_CHANGED, this._values);
    }

    /**
     * 在选区位置插入空白单元格 + 右侧可见单元格向右让位（单元格级，非整列插入）。
     *
     * 对每一可见行 r in [startRow..endRow]：
     * 1. 选区内可见列的原有内容随该行可见单元格整体右移 k 个可见位（k = 选区内可见列数）；
     * 2. 选区内可见列位置变为空白（新插入的空单元格）；
     * 3. 被推出最右可见列之外的内容丢弃（与 Excel「活动单元格右移」一致，对称于 shiftCellsLeft 的右端清空）。
     *
     * 隐藏行/列的单元格内容与格式始终固定原位、不参与移动：
     * 可见单元格仅在可见位置之间右移让位（可跨越隐藏列），隐藏列既不搬出也不作为搬入目标。
     * 合并单元格不做特殊处理（暂由上层在调用前确保选区不含合并单元格）。
     * @param {number} startCol - 选区起始列（1-based，含）
     * @param {number} endCol - 选区结束列（1-based，含）
     * @param {number} startRow - 选区起始行（1-based，含）
     * @param {number} endRow - 选区结束行（1-based，含）
     */
    public insertCellsRight(startCol: number, endCol: number, startRow: number, endRow: number): void {
        const totalCols = this.colHeaders.length;
        // 可见列索引列表（隐藏列固定原位，既不搬出也不作为搬入目标）
        const visCols: number[] = [];
        for (let c = 1; c <= totalCols; c++) if (!this.colHeaders.getAt(c - 1)?.isHidden) visCols.push(c);
        const m = visCols.length;
        if (m === 0) return;
        // si = 选区起第一个可见列在 visCols 中的序号；k = 选区内可见列数（需让位的可见位数）
        let si = visCols.findIndex(c => c >= startCol);
        if (si < 0) si = m; // 选区完全在最后一个可见列之后：无可让位
        const k = visCols.filter(c => c >= startCol && c <= endCol).length;
        if (k === 0) return; // 选区内无可见列：无可插入空位
        for (let r = startRow; r <= endRow; r++) {
            if (this.rowHeaders.getAt(r - 1)?.isHidden) continue;
            // 自右向左搬运：保证目标列（序号更大）已先被清空，避免覆盖未搬的源
            for (let j = m - 1; j >= si; j--) {
                const srcCol = visCols[j];
                const tgtRank = j + k;
                if (tgtRank >= m) {
                    // 被推出最右可见列之外：内容丢弃（清空源）
                    const srcCell = this.getCellByName(this.getCellName(srcCol, r));
                    if (srcCell) srcCell.reset();
                    continue;
                }
                const tgtCol = visCols[tgtRank];
                const srcName = this.getCellName(srcCol, r);
                const tgtName = this.getCellName(tgtCol, r);
                const srcCell = this.getCellByName(srcName);
                const tgtCell = this.getCellByName(tgtName);
                if (srcCell) {
                    const charsCopy = srcCell.chars.map(ch => Object.assign(new Char(), ch));
                    if (tgtCell) {
                        Object.assign(tgtCell, srcCell, { cell: tgtName, chars: charsCopy });
                    } else {
                        const cloned = new Cell();
                        Object.assign(cloned, srcCell, { cell: tgtName, chars: charsCopy });
                        this.values.push(cloned);
                    }
                    srcCell.reset();
                } else if (tgtCell) {
                    // 源为空：空位同样右移，目标清空
                    tgtCell.reset();
                }
            }
        }
        this.emit(DataEvents.VALUES_CHANGED, this._values);
    }

    /**
     * 在选区位置插入空白单元格 + 下方可见单元格向下让位（单元格级，非整行插入）。
     *
     * 与 insertCellsRight 对称，仅方向不同；隐藏行/列的单元格内容与格式始终固定原位、不参与移动。
     * @param {number} startCol - 选区起始列（1-based，含）
     * @param {number} endCol - 选区结束列（1-based，含）
     * @param {number} startRow - 选区起始行（1-based，含）
     * @param {number} endRow - 选区结束行（1-based，含）
     */
    public insertCellsDown(startCol: number, endCol: number, startRow: number, endRow: number): void {
        const totalRows = this.rowHeaders.length;
        const visRows: number[] = [];
        for (let r = 1; r <= totalRows; r++) if (!this.rowHeaders.getAt(r - 1)?.isHidden) visRows.push(r);
        const m = visRows.length;
        if (m === 0) return;
        let si = visRows.findIndex(r => r >= startRow);
        if (si < 0) si = m;
        const k = visRows.filter(r => r >= startRow && r <= endRow).length;
        if (k === 0) return;
        for (let c = startCol; c <= endCol; c++) {
            if (this.colHeaders.getAt(c - 1)?.isHidden) continue;
            // 自下向上搬运：保证目标行（序号更大）已先被清空，避免覆盖未搬的源
            for (let j = m - 1; j >= si; j--) {
                const srcRow = visRows[j];
                const tgtRank = j + k;
                if (tgtRank >= m) {
                    const srcCell = this.getCellByName(this.getCellName(c, srcRow));
                    if (srcCell) srcCell.reset();
                    continue;
                }
                const tgtRow = visRows[tgtRank];
                const srcName = this.getCellName(c, srcRow);
                const tgtName = this.getCellName(c, tgtRow);
                const srcCell = this.getCellByName(srcName);
                const tgtCell = this.getCellByName(tgtName);
                if (srcCell) {
                    const charsCopy = srcCell.chars.map(ch => Object.assign(new Char(), ch));
                    if (tgtCell) {
                        Object.assign(tgtCell, srcCell, { cell: tgtName, chars: charsCopy });
                    } else {
                        const cloned = new Cell();
                        Object.assign(cloned, srcCell, { cell: tgtName, chars: charsCopy });
                        this.values.push(cloned);
                    }
                    srcCell.reset();
                } else if (tgtCell) {
                    tgtCell.reset();
                }
            }
        }
        this.emit(DataEvents.VALUES_CHANGED, this._values);
    }


    /**
     * 清除指定范围内单元格的内容（保留样式）
     * @param {number} startCol - 起始列号（1-based，含）
     * @param {number} endCol - 结束列号（1-based，含）
     * @param {number} startRow - 起始行号（1-based，含）
     * @param {number} endRow - 结束行号（1-based，含）
     */
    public clearContents(startCol: number, endCol: number, startRow: number, endRow: number): void {
        for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
                const cell = this.getCellByName(this.getCellName(c, r));
                if (cell) cell.reset();  // 清空所有内容+格式（保留功能性字段 colspan/rowspan/filter）
            }
        }
        this.emit(DataEvents.VALUES_CHANGED, this._values);
    }

    /**
     * 清除选区内容 + 右侧可见单元格向左紧凑补位（单元格级，非整列删除）。
     * 对每一可见行 r in [startRow..endRow]：
     * 1. 清空该行 [startCol..endCol] 范围内可见列的 chars（隐藏列内容与格式保留原样）；
     * 2. 从 endCol+1 向右遍历可见列 c'，将其单元格内容（chars + 样式 + 合并属性）
     *    依次搬到写入指针位置；写入指针从选区起第一个可见列开始，每搬一格前进到下一个可见列；
     * 3. 原位置（c'）的 Cell 内容清空。
     *
     * 隐藏行/列的单元格内容与格式始终固定原位、不参与移动：
     * 可见单元格仅在可见位置之间紧凑补位（可跨越隐藏列），隐藏列既不搬出也不作为搬入目标。
     * 合并单元格不做特殊处理（暂由上层在调用前确保选区不含合并单元格）。
     * 操作后不自动 commitValues/emit——调用方应包在 runWithFullStateUndo 内。
     * @param {number} startCol - 选区起始列（1-based，含）
     * @param {number} endCol - 选区结束列（1-based，含）
     * @param {number} startRow - 选区起始行（1-based，含）
     * @param {number} endRow - 选区结束行（1-based，含）
     */
    public shiftCellsLeft(startCol: number, endCol: number, startRow: number, endRow: number): void {
        const totalCols = this.colHeaders.length;
        for (let r = startRow; r <= endRow; r++) {
            // 跳过隐藏行：该行所有单元格（内容与格式）保持原位，不参与移动
            if (this.rowHeaders.getAt(r - 1)?.isHidden) continue;
            // 1. 清空选区内可见列的内容（隐藏列内容与格式保留原位不动）
            for (let c = startCol; c <= endCol; c++) {
                if (this.colHeaders.getAt(c - 1)?.isHidden) continue;
                const cell = this.getCellByName(this.getCellName(c, r));
                if (cell) cell.reset();
            }
            // 2. 可见单元格向左紧凑补位：写入指针从选区起第一个可见列开始，逐格接收右侧可见列的内容；
            //    隐藏列固定原位——既不搬出、也不作为搬入目标，可见内容可跨越隐藏列补位
            let writeCol = startCol;
            while (writeCol <= totalCols && this.colHeaders.getAt(writeCol - 1)?.isHidden) writeCol++;
            for (let c = endCol + 1; c <= totalCols; c++) {
                // 隐藏源列：内容与格式保持原位，不参与移动，也不消耗写入位置
                if (this.colHeaders.getAt(c - 1)?.isHidden) continue;
                // 写入指针追上源列：选区内已无可填充的可见空位，剩余单元格无需移动
                if (writeCol >= c) break;
                const srcName = this.getCellName(c, r);
                const tgtName = this.getCellName(writeCol, r);
                const srcCell = this.getCellByName(srcName);
                const tgtCell = this.getCellByName(tgtName);
                if (srcCell) {
                    const charsCopy = srcCell.chars.map(ch => Object.assign(new Char(), ch));
                    if (tgtCell) {
                        Object.assign(tgtCell, srcCell, { cell: tgtName, chars: charsCopy });
                    } else {
                        const cloned = new Cell();
                        Object.assign(cloned, srcCell, { cell: tgtName, chars: charsCopy });
                        this.values.push(cloned);
                    }
                    srcCell.reset();
                } else if (tgtCell) {
                    tgtCell.reset();
                }
                // 写入指针前进到下一个可见列（跳过隐藏列）
                do {
                    writeCol++;
                } while (writeCol <= totalCols && this.colHeaders.getAt(writeCol - 1)?.isHidden);
            }
        }
        this.emit(DataEvents.VALUES_CHANGED, this._values);
    }

    /**
     * 清除选区内容 + 下方可见单元格向上紧凑补位（单元格级，非整行删除）。
     * 对每一可见列 c in [startCol..endCol]：
     * 1. 清空该列 [startRow..endRow] 范围内可见行的 chars（隐藏行内容与格式保留原样）；
     * 2. 从 endRow+1 向下遍历可见行 r'，将其单元格内容依次搬到写入指针位置；
     * 3. 原位置（r'）的 Cell 内容清空。
     *
     * 与 shiftCellsLeft 对称，仅方向不同；隐藏行/列的单元格内容与格式始终固定原位、不参与移动。
     * @param {number} startCol - 选区起始列（1-based，含）
     * @param {number} endCol - 选区结束列（1-based，含）
     * @param {number} startRow - 选区起始行（1-based，含）
     * @param {number} endRow - 选区结束行（1-based，含）
     */
    public shiftCellsUp(startCol: number, endCol: number, startRow: number, endRow: number): void {
        const totalRows = this.rowHeaders.length;
        for (let c = startCol; c <= endCol; c++) {
            // 跳过隐藏列：该列所有单元格（内容与格式）保持原位，不参与移动
            if (this.colHeaders.getAt(c - 1)?.isHidden) continue;
            // 1. 清空选区内可见行的内容（隐藏行内容与格式保留原位不动）
            for (let r = startRow; r <= endRow; r++) {
                if (this.rowHeaders.getAt(r - 1)?.isHidden) continue;
                const cell = this.getCellByName(this.getCellName(c, r));
                if (cell) cell.reset();
            }
            // 2. 可见单元格向上紧凑补位：写入指针从选区起第一个可见行开始，逐格接收下方可见行的内容；
            //    隐藏行固定原位——既不搬出、也不作为搬入目标，可见内容可跨越隐藏行补位
            let writeRow = startRow;
            while (writeRow <= totalRows && this.rowHeaders.getAt(writeRow - 1)?.isHidden) writeRow++;
            for (let r = endRow + 1; r <= totalRows; r++) {
                // 隐藏源行：内容与格式保持原位，不参与移动，也不消耗写入位置
                if (this.rowHeaders.getAt(r - 1)?.isHidden) continue;
                // 写入指针追上源行：选区内已无可填充的可见空位，剩余单元格无需移动
                if (writeRow >= r) break;
                const srcName = this.getCellName(c, r);
                const tgtName = this.getCellName(c, writeRow);
                const srcCell = this.getCellByName(srcName);
                const tgtCell = this.getCellByName(tgtName);
                if (srcCell) {
                    const charsCopy = srcCell.chars.map(ch => Object.assign(new Char(), ch));
                    if (tgtCell) {
                        Object.assign(tgtCell, srcCell, { cell: tgtName, chars: charsCopy });
                    } else {
                        const cloned = new Cell();
                        Object.assign(cloned, srcCell, { cell: tgtName, chars: charsCopy });
                        this.values.push(cloned);
                    }
                    srcCell.reset();
                } else if (tgtCell) {
                    tgtCell.reset();
                }
                // 写入指针前进到下一个可见行（跳过隐藏行）
                do {
                    writeRow++;
                } while (writeRow <= totalRows && this.rowHeaders.getAt(writeRow - 1)?.isHidden);
            }
        }
        this.emit(DataEvents.VALUES_CHANGED, this._values);
    }

    /**
     * 设置指定范围内单元格的背景颜色（设置单元格格式）
     * @param {number} startCol - 起始列号（1-based，含）
     * @param {number} endCol - 结束列号（1-based，含）
     * @param {number} startRow - 起始行号（1-based，含）
     * @param {number} endRow - 结束行号（1-based，含）
     * @param {string} color - 背景颜色值（如 #ff0000）
     */
    public setCellsBackgroundColor(startCol: number, endCol: number, startRow: number, endRow: number, color: string): void {
        for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
                // 合并区内部的位置重定向到锚点；本方法不改「有没有这个格子」，
                // 故只重定向、不新建（新建是 ensureWritableCell 的事）
                const cell = this.getCellByName(this.resolveWriteTargetName(c, r));
                if (cell) cell.backgroundColor = color;
            }
        }
    }

    /**
     * 设置指定范围内单元格的数字格式码（设置单元格格式）
     * @param {number} startCol - 起始列号（1-based，含）
     * @param {number} endCol - 结束列号（1-based，含）
     * @param {number} startRow - 起始行号（1-based，含）
     * @param {number} endRow - 结束行号（1-based，含）
     * @param {string} format - 数字格式码（如 '0.00'/'#,##0.00'/'0.00%'/'¥#,##0.00'/'0.00E+00'/'@'/'General'）
     */
    public setCellsNumberFormat(startCol: number, endCol: number, startRow: number, endRow: number, format: string): void {
        for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
                // 合并区内部的位置重定向到锚点（见 ensureWritableCell 的说明）
                this.ensureWritableCell(this.getCellName(c, r)).numberFormat = format;
            }
        }
    }

    /**
     * 设置指定范围内单元格的边框颜色（设置单元格格式）
     * @param {number} startCol - 起始列号（1-based，含）
     * @param {number} endCol - 结束列号（1-based，含）
     * @param {number} startRow - 起始行号（1-based，含）
     * @param {number} endRow - 结束行号（1-based，含）
     * @param {string} color - 边框颜色值（如 '#000000'）
     */
    public setCellsBorderColor(startCol: number, endCol: number, startRow: number, endRow: number, color: string): void {
        for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
                // 同 setCellsBackgroundColor：只重定向到锚点，不新建格子
                const cell = this.getCellByName(this.resolveWriteTargetName(c, r));
                if (cell) cell.borderColor = color;
            }
        }
    }

    

    /**
     * 滚动表格使目标单元格完全可见（按需调整行列头滚动偏移）
     *
     * 与 `Canvas.scrollCellIntoView` 同算法，但坐标取**网格矩形**
     * （{@link getCellGridRect}，不做合并展开）：命中的格子若落在合并区内，
     * 按合并锚点算出的矩形会连带把旁边的行列卷进视野。
     * @param {number} col - 列号（1-based）
     * @param {number} row - 行号（1-based）
     * @returns {boolean} 是否发生了滚动
     */
    public scrollCellIntoView(col: number, row: number): boolean {
        const { left, top, width, height } = this.getCellGridRect(col, row);
        if (width === 0 && height === 0) return false;
        const { sheetWidth, sheetHeight } = this.visibleView;
        const visibleWidth = sheetWidth / this.zoom;
        const visibleHeight = sheetHeight / this.zoom;
        const totalWidth = this.colHeaders.allColWidth;
        const totalHeight = this.rowHeaders.allRowHeight;
        let offsetX = this.colHeaders.offsetWidth;
        let offsetY = this.rowHeaders.offsetHeight;
        let needScroll = false;

        const absoluteLeft = left + offsetX;
        const absoluteTop = top + offsetY;
        if (left < 0) {
            // 列号 2 起才允许向左滚动到露出它：第 1 列本来就该贴住视口左边，
            // 滚到 1px 只会把首列切掉一条边（像素级位移下最容易发生）
            offsetX = Math.max(0, col > 1 ? absoluteLeft : 0);
            needScroll = offsetX !== this.colHeaders.offsetWidth;
        } else if (left + width > visibleWidth) {
            const required = absoluteLeft + width - visibleWidth;
            offsetX = Math.min(Math.max(0, required), Math.max(0, totalWidth - visibleWidth));
            needScroll = true;
        }
        // 行号 2 起再往下滚：第 1 行恒在视口顶端，滚到 1px 只是让首行被裁掉一条边。
        // 同理第 1 列不接受「左上角露在外面」这种滑移（拖到像素级滚动时最容易发生）。
        if (top < 0) {
            offsetY = Math.max(0, row > 1 ? absoluteTop : 0);
            needScroll = offsetY !== this.rowHeaders.offsetHeight;
        } else if (top + height > visibleHeight) {
            const required = absoluteTop + height - visibleHeight;
            offsetY = Math.min(Math.max(0, required), Math.max(0, totalHeight - visibleHeight));
            needScroll = true;
        }
        if (needScroll) {
            this.colHeaders.offsetWidth = offsetX;
            this.rowHeaders.offsetHeight = offsetY;
        }
        return needScroll;
    }


    /**
     * 通知工具栏状态更新
     */
    public notifyToolbarStateChange() {
        this.emit(DataEvents.CURSOR_STATE_CHANGED);
    }

}
