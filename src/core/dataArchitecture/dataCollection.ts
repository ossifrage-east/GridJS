import { ColHeader, ColHeaders, RowHeader, RowHeaders } from "./header";
import { Cell, TextAlign, VerticalAlign } from "./cell";
import { LineText, SelectionRange, FontStyle, BorderStyle, CellBorderStyle, CursorPosition, } from "./cellContent";
import { Char } from "./char";
import { VisibleView } from "./visibleView";
import { CELL_PADDING, DataEvents, DEFAULT_CELL_HEIGHT, DEFAULT_CELL_WIDTH, DEFAULT_SELECTED_BORDER_WIDTH, MouseLocation } from "../constant";
import { EventEmitter } from "../../utils/eventEmitter";
import { IndexDB } from "./indexDB";
import { UndoManager, UndoAction, FullStateSnapshot } from "../../utils/undoManager";
import { ClipboardManager } from "../../utils/clipboardManager";
import { MessageBox } from "../common/message";
import { MatrixSorter, type SortConfig } from "./matrixSorter";
import { PrinterSetting } from "../../utils/printer";


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
export class DataCollection extends EventEmitter {
    /** 索引数据库实例 */
    readonly db: IndexDB;
    /** 可见视图实例 */
    visibleView: VisibleView;
    /** 列头集合 */
    colHeaders: ColHeaders;
    /** 行头集合 */
    rowHeaders: RowHeaders;
    /** 当前工作表sheet，所有有内容的JSON字符串（内部后备字段，统一通过 values getter/setter 访问） */
    private _values: Cell[] = [];
    /**
     * 上一次提交时的 this.values 深拷贝快照。
     * 作为下一次变更的「变化前」基准：每次原地修改后调用 commitValues()，
     * 便会以本快照为 before、当前 _values 为 after 记录一条撤销操作。
     */
    private _lastValuesSnapshot: Cell[] = [];
    /** 是否正在回写撤销/重做快照（期间禁止再次记录，避免递归入栈） */
    private _isApplyingUndo: boolean = false;
    /** 是否禁止记录撤销（初始从数据库装载时为 true，避免把首次装载误记为一次变更） */
    private _suppressValuesUndo: boolean = false;
    /**
     * 全局筛选条件注册表：列号(1-based) → { anchorRow: 筛选按钮所在行(1-based), keep: 保留内容 }。
     * 注册表随全量状态快照捕获/回写，使筛选操作可整体撤销/重做；
     * 筛选组件（assistant/filter）通过 get/set/clear API 读写，不再自持静态状态。
     */
    private _filterConditions: Map<number, { anchorRow: number; keep: Set<string> }> = new Map();
    /** 筛选操作顺序：已筛选列号按首次执行筛选的先后排列（重放筛选条件时使用） */
    private _filterOrder: number[] = [];
    /** runWithFullStateUndo 嵌套深度：仅最外层推入撤销记录，嵌套调用的变更折叠进外层快照 */
    private _fullStateUndoDepth: number = 0;
    /**
     * 编辑期间的撤销栈：保存每次按键后的 values 深拷贝快照及对应光标绝对索引。
     * 栈底（index 0）为编辑前的基准快照，供编辑内逐字撤销；
     * 整段编辑在 endEditSession() 时折叠为 undoManager 的一条文档级撤销操作。
     */
    private _editUndoStack: { values: Cell[]; cursor: number }[] = [];
    /** 编辑期间的重做栈，配合 _editUndoStack 使用 */
    private _editRedoStack: { values: Cell[]; cursor: number }[] = [];
    /**
     * 编辑内撤销/重做后应恢复到的光标绝对索引。
     * 由 _undoEditStep/_redoEditStep 写入，供 editCanvas 在重测 lines、
     * 定位光标时读取（光标定位必须发生在 lines 刷新之后）。
     */
    private _editRestoreCursor: number = 0;
    /** 当前选中的工作表名称 */
    private _activedSheetName: string = 'Sheet1';
    /** 是否正在编辑 */
    private _isEditting: boolean = false;
    /** 是否显示网格线 */
    private _showGridLines: boolean = true;
    /** 当前选中的单元格 */
    private _activedCell: string = 'A1';
    /** sheetCanvas选中区域 */
    private _selection: string;
    /** sheetCanvas锚点选择框 */
    private _anchorSelection: string | null;
    /** 缩放比例 */
    private _zoom: number = 1;
    /** 打印设置 */
    private _printSetting: PrinterSetting = {};
    /** sheetCanvas选中区域的起始位置 */
    startLocation: string = '';
    /** 是否正在按下鼠标键 */
    isMouseDown: boolean = false;
    /** 鼠标悬停的单元格 */
    overCell: { col: number, row: number } = { col: 0, row: 0 };
    /** 鼠标点击列头时的 x 坐标 */
    public MouseStartX: number = 0;
    /** 开始列索引 */
    public startCol: number = 1;  
    /** 开始列宽度 */
    public startWidth: number = 0;  
    /** 鼠标点击行头时的 y 坐标 */
    public MouseStartY: number = 0;
    /** 开始行索引 */
    public startRow: number = 1;  
    /** 开始行高度 */
    public startHeight: number = 0;
    /** 文本行数组 */
    public lines: LineText[] = [];
    /** 编辑状态下当前选择范围 */
    public selectionOnEditting: SelectionRange | null = null;
    /** 光标位置（唯一真相源，其余表示均由它派生） */
    public cursorPosition: CursorPosition | null = null;
    /** 光标绝对位置缓存（由 cursorPosition 派生，勿直接读写，统一走 cursorAbsolutePosition 读/写器） */
    public _cursorAbsolutePosition: number = 0;
    /**
     * 光标屏幕坐标解析器
     * 由渲染层（EditCanvas）注入，根据 lineIndex/charIndex 计算 editX/editY，
     * 统一走 calcCursorX + line.editY，保证坐标与文本渲染（含 letterSpacing）规则一致。
     * 未注入时回退到行起始坐标，保证纯数据环境下也能工作。
     */
    public cursorCoordResolver: ((lineIndex: number, charIndex: number) => { editX: number; editY: number }) | null = null;
    /** 撤销/恢复管理器实例 */
    public undoManager: UndoManager;
    /** 剪贴板管理器实例 */
    public clipboardManager: ClipboardManager;
    /** @type {'single' | 'continuous'} 格式刷模式（单次模式或连续模式） */
    private _brushMode: 'single' | 'continuous' | null = null;
    /** 数据加载完成的 Promise，供外部组件等待数据就绪 */
    public dataLoaded: Promise<void>;
    /** 数据加载完成的 Promise 解析函数 */
    private _dataLoadedResolve!: () => void;
    /** 消息框实例 */
    public messageBox: MessageBox;
    /** 矩阵排序器实例 */
    private matrixSorter: MatrixSorter;

    constructor() {
        super();
        this.db = new IndexDB();
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
     * 从数据库获取当前选中工作表的所有有内容的单元格数据
     */
    public async syncValuesFromDB(): Promise<void> {
        await this.db.getValue(this._activedSheetName).then(async (valueRes) => {
            // 装载数据属于初始化，不应被记为一次「变更」，期间禁止记录撤销
            this._suppressValuesUndo = true;
            try {
                if(valueRes) {
                    const data = DataCollection.fromJSON(valueRes);
                    this._activedCell = data.activedCell;
                    this._selection = data.selection || `${data.activedCell}:${data.activedCell}`;
                    this._anchorSelection = null;
                    // 直接赋值数组不会触发重算，需显式重算总尺寸，
                    // 否则 allColWidth/allRowHeight 保持 0，滚动条 thumb 会占满轨道导致无法拖动
                    this.colHeaders.colHeaders = data.colHeaders;
                    this.colHeaders.recalcAllColWidth();
                    this.colHeaders.offsetWidth = data.offsetWidth;
                    this.rowHeaders.rowHeaders = data.rowHeaders;
                    this.rowHeaders.recalcAllRowHeight();
                    this.rowHeaders.offsetHeight = data.offsetHeight;
                    this._showGridLines = data.showGridLines;
                    this._zoom = data.zoom;
                    // 打印设置统一入口随文档恢复（未保存过时保持接口默认值）
                    this._printSetting = data.printSetting || {};
                    this.values = data.values;
                }
            } finally {
                // 同步最新快照基准，解除抑制
                this._lastValuesSnapshot = this.cloneValues(this._values);
                this._suppressValuesUndo = false;
                this._dataLoadedResolve();
            }
        });
    }

    /**
     * 同步当前选中工作表的所有有内容的单元格数据到数据库中
     */
    public async syncValuesToDB(): Promise<void> {
        await this.db.setValue(this.activedSheetName, this.toJSON());
    }
    
    /**
     * 获取指定单元格的值
     * @param {string} cell - 单元格名称
     * @returns {string} 单元格值
     */
    public getCellValue(cell: string): string {
        return this.values.find((v) => v.cell === cell)?.chars?.map((v: Char) => v.char).join('') || '';
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
        this._recordValuesChange(oldValue);
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
        return removed;
    }

    /**
     * 比较两份快照是否等价（基于每格 toJSON 字符串）
     * 用于跳过无实际变化的提交，避免在撤销栈中产生空操作。
     * @param {Cell[]} a - 快照 A
     * @param {Cell[]} b - 快照 B
     * @returns {boolean} 是否等价
     */
    private _valuesSnapshotsEqual(a: Cell[], b: Cell[]): boolean {
        if (a === b) return true;
        if (a.length !== b.length) return false;
        for (let i = 0; i < a.length; i++) {
            if (a[i].toJSON() !== b[i].toJSON()) return false;
        }
        return true;
    }

    /**
     * 记录一次 this.values 变更并广播
     * 以 before 为「变化前」基准，当前 _values 的深拷贝为「变化后」after，
     * 满足条件时推入撤销栈，随后刷新 _lastValuesSnapshot 基准并触发 VALUES_CHANGED。
     * - 装载期间（_suppressValuesUndo）：仅刷新基准、不广播、不入栈。
     * - 撤销/重做回写期间（_isApplyingUndo）：仅刷新基准并广播，不入栈。
     * @param {Cell[]} before - 变更前的快照基准
     */
    private _recordValuesChange(before: Cell[]): void {
        if (this._suppressValuesUndo) {
            this._lastValuesSnapshot = this.cloneValues(this._values);
            return;
        }
        const after = this.cloneValues(this._values);
        if (!this._isApplyingUndo && !this._valuesSnapshotsEqual(before, after)) {
            this.undoManager.pushValuesSnapshot(this.cloneValues(before), this.cloneValues(after));
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
     */
    public commitValues(): void {
        if (this.isEditting) {
            const current = this.cloneValues(this._values);
            const last = this._editUndoStack[this._editUndoStack.length - 1];
            // 仅在内容确实变化时入栈，跳过无变化的提交，避免撤销栈产生空操作
            if (!last || !this._valuesSnapshotsEqual(last.values, current)) {
                this._editUndoStack.push({ values: current, cursor: this.cursorAbsolutePosition });
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
     * 进入编辑会话
     * 捕获编辑前的 values 快照与光标位置作为编辑内撤销栈的基准（栈底 index 0），
     * 并清空编辑内重做栈。应在 isEditting 置为 true 时调用。
     */
    public beginEditSession(): void {
        this._editUndoStack = [{ values: this.cloneValues(this._values), cursor: this.cursorAbsolutePosition }];
        this._editRedoStack = [];
    }

    /**
     * 结束编辑会话
     * 把整段编辑折叠为 undoManager 的一条文档级撤销操作（编辑前→最终值），
     * 并刷新 _lastValuesSnapshot 基准；随后清空编辑内撤销/重做栈。
     * 应在 isEditting 置为 false 时调用。
     */
    public endEditSession(): void {
        if (this._editUndoStack.length > 0) {
            // 编辑过程中不做清扫（保护编辑内撤销栈快照与光标索引的布局一致性），
            // 结束时统一清扫，使折叠入撤销栈的「最终值」与快照基准均不含空单元格
            this._pruneEmptyCells();
            const preEdit = this._editUndoStack[0].values;
            const final = this.cloneValues(this._values);
            if (!this._valuesSnapshotsEqual(preEdit, final)) {
                this.undoManager.pushValuesSnapshot(this.cloneValues(preEdit), final);
            }
            this._lastValuesSnapshot = this.cloneValues(this._values);
        }
        this._editUndoStack = [];
        this._editRedoStack = [];
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
        if (this._editUndoStack.length <= 1) return false;
        const current = this._editUndoStack.pop()!;
        this._editRedoStack.push(current);
        const target = this._editUndoStack[this._editUndoStack.length - 1];
        this._editRestoreCursor = target.cursor;
        this._applyEditState(target.values);
        return true;
    }

    /**
     * 编辑内重做一步
     * 从 _editRedoStack 弹出状态回写，并转回 _editUndoStack。
     * 同样仅替换 _values，光标与广播交由 editCanvas 处理。
     * @returns {boolean} 是否重做成功
     */
    private _redoEditStep(): boolean {
        if (this._editRedoStack.length === 0) return false;
        const state = this._editRedoStack.pop()!;
        this._editUndoStack.push(state);
        this._editRestoreCursor = state.cursor;
        this._applyEditState(state.values);
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
     * 捕获当前全量状态快照（values + 行/列表头 + 网格线 + 活动单元格 + 选区 + 筛选条件）
     * 所有字段均为深拷贝，与当前数据完全解耦。
     * @returns {FullStateSnapshot} 全量状态快照
     */
    private _captureFullState(): FullStateSnapshot {
        return {
            values: this.cloneValues(this._values),
            rowHeaders: this.rowHeaders.rowHeaders.map(rh => ({ ...rh })),
            colHeaders: this.colHeaders.colHeaders.map(ch => ({ ...ch })),
            showGridLines: this._showGridLines,
            activedCell: this._activedCell,
            selection: this._selection,
            filters: this._serializeFilterConditions(),
        };
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
     * 比较两份全量状态快照是否等价（基于 JSON 字符串）
     * 用于跳过无实际变化的操作（如取消的 prompt、对已隐藏行列再次取消隐藏），避免撤销栈空操作。
     * @param {FullStateSnapshot} a - 快照 A
     * @param {FullStateSnapshot} b - 快照 B
     * @returns {boolean} 是否等价
     */
    private _fullStateSnapshotsEqual(a: FullStateSnapshot, b: FullStateSnapshot): boolean {
        return JSON.stringify(a) === JSON.stringify(b);
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
     * 时，嵌套层不推入撤销记录，其变更已包含在最外层快照的 after 中，
     * 整个复合操作折叠为一步撤销。
     * @param {string} label - 操作标签（用于调试与撤销展示）
     * @param {() => void} action - 实际的变更操作
     * @example
     * data.runWithFullStateUndo('删除行', () => data.deleteRows(2, 3));
     */
    public runWithFullStateUndo(label: string, action: () => void): void {
        const before = this._captureFullState();
        const prevSuppress = this._suppressValuesUndo;
        this._suppressValuesUndo = true;
        this._fullStateUndoDepth++;
        try {
            action();
        } finally {
            // 操作结束后清扫空单元格（如清除内容后残留的空 Cell），
            // 使「变化后」快照与稳态数据均不残留空单元格
            this._pruneEmptyCells();
            const after = this._captureFullState();
            this._fullStateUndoDepth--;
            this._suppressValuesUndo = prevSuppress;
            // 仅最外层推入撤销记录；撤销/重做回写期间（_isApplyingUndo）不得入栈
            if (this._fullStateUndoDepth === 0 && !this._isApplyingUndo
                && !this._fullStateSnapshotsEqual(before, after)) {
                this.undoManager.push({ type: 'fullState', data: { before, after }, timestamp: Date.now(), label });
            }
        }
    }

    /**
     * 开始一次跨事件的连续操作撤销捕获（如拖拽调整行高/列宽）
     * 立即捕获全量状态快照作为「变化前」基准，并抑制 values setter 的自动快照记录，
     * 直到调用 {@link endContinuousUndo} 提交。两次调用必须成对出现。
     * @returns {FullStateSnapshot} 操作前全量状态快照
     */
    public beginContinuousUndo(): FullStateSnapshot {
        this._suppressValuesUndo = true;
        return this._captureFullState();
    }

    /**
     * 结束连续操作并提交撤销记录
     * 捕获当前全量状态为「变化后」快照，与 before 不同时推入 undoManager（type: 'fullState'），
     * 并恢复 values setter 的自动记录。整个拖拽手势折叠为一条撤销记录。
     * @param {FullStateSnapshot | null} before - {@link beginContinuousUndo} 返回的快照；为 null 时忽略
     * @param {string} label - 操作标签
     */
    public endContinuousUndo(before: FullStateSnapshot | null, label: string): void {
        if (!before) return;
        try {
            const after = this._captureFullState();
            if (!this._fullStateSnapshotsEqual(before, after)) {
                this.undoManager.push({ type: 'fullState', data: { before, after }, timestamp: Date.now(), label });
            }
        } finally {
            this._suppressValuesUndo = false;
        }
    }

    /**
     * 回写全量状态快照（撤销/重做恢复）
     * 依次恢复筛选条件注册表、values（走 values setter，回写期间抑制自动记录并广播
     * VALUES_CHANGED）、行/列表头（重算行列总尺寸）、网格线开关、活动单元格与选区；
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
            // 全部字段恢复完毕后广播一次，确保派生 UI（如筛选按钮激活色）与最终状态一致
            this.emit(DataEvents.VALUES_CHANGED, this._values);
        } finally {
            this._isApplyingUndo = false;
        }
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
        } else {
            return;
        }
        // 回写完成后广播快照恢复事件：activedCell/selection 与当前值相同时
        // updateProperty 不会发出差量事件（ACTIVED_CELL_CHANGED 等），
        // 依赖差量事件刷新外观的工具栏分区需监听本事件同步按钮状态
        this.emit(DataEvents.SNAPSHOT_RESTORED, action, isUndo);
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
     * @param {string} cell - 单元格名称
     * @returns {{ col: number, row: number }} 包含列号和行号的对象
     */
    public getCellColAndRow(cell: string): { col: number, row: number } {
        const col = this.getColNumber(cell.replace(/[^A-Z]/g, ''));
        const row = parseInt(cell.replace(/[^\d]/g, ''));
        return { col, row };
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
     * @returns {{ location: MouseLocation, col: number, row: number }} 包含鼠标位置、列号和行号的对象
     */
    public getSheetCell(_x: number, _y: number): {location: MouseLocation, col: number, row: number} {
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
            
        // 锚点覆盖：若锚点自身隐藏则跳过
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

        // 合并单元格覆盖：若合并后的目标单元格隐藏则跳过
        const { col: mergedCol, row: mergedRow } = this.getMergeColAndRow(this.getCellName(col, row));
        if (mergedCol !== col || mergedRow !== row) {
            const mergedColHidden = this.isColHidden(mergedCol - 1);
            const mergedRowHidden = this.isRowHidden(mergedRow - 1);
            if (!mergedColHidden && !mergedRowHidden) {
                col = mergedCol;
                row = mergedRow;
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
                const cell = this.values.find(c => c.cell === cellName);
                if (cell && this.hasCellContent(cell)) {
                    return cell;
                }
            }
        }
        return null;
    }

    /**
     * 判断单元格是否有内容
     */
    private hasContent(cellName: string): boolean {
        const cell = this.values.find(c => c.cell === cellName);
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
        const mergedCells = this.getMergedCells();
        for (let i = 0; i < mergedCells.length; i++) {
            const startCell = mergedCells[i].split(':')[0];
            const endCell = mergedCells[i].split(':')[1];
            const { col: startCol, row: startRow } = this.getCellColAndRow(startCell);
            const { col: endCol, row: endRow } = this.getCellColAndRow(endCell);

            if (col >= startCol && col <= endCol && row >= startRow && row <= endRow) {
                realCol = startCol;
                realRow = startRow;
                break;
            }
        }
        
        const left = this.colHeaders.getAt(realCol - 1).left - this.colHeaders.offsetWidth;
        const top = this.rowHeaders.getAt(realRow - 1).top - this.rowHeaders.offsetHeight;
        let width = this.colHeaders.getAt(realCol - 1).width;
        let height = this.rowHeaders.getAt(realRow - 1).height;
        
        const cellName = this.getCellName(realCol, realRow);
        const cellIndex = this.values.findIndex(v => v.cell === cellName);
        if (cellIndex >= 0) {
            const colspan = this.values[cellIndex].colspan || 1;
            const rowspan = this.values[cellIndex].rowspan || 1;
            
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
     * @returns {{startCol: number, endCol: number}} 可见列的起始索引和结束索引
     */
    public getVisibleCols(): {startCol: number, endCol: number} {
        const { startCol } = this.colHeaders.getVisibleStartCol();
        let endCol = this.colHeaders.length - 1;
        const viewportRight = this.colHeaders.offsetWidth + this.visibleView.sheetWidth / this.zoom;
        while (endCol > startCol) {
            const col = this.colHeaders.getAt(endCol - 1);
            if (!col) break;
            if (col.isHidden) { endCol--; continue; }  // 跳过隐藏列
            if (col.left + col.width < viewportRight) break;  // 列右边缘进入视口
            endCol--;
        }
        return { startCol, endCol: Math.max(startCol, endCol) };
    }
    
    /**
     * 获取可见行的起始索引和结束索引，从1开始计数
     * @returns {{startRow: number, endRow: number}} 可见行的起始索引和结束索引
     */
    public getVisibleRows(): {startRow: number, endRow: number} {
        const { startRow } = this.rowHeaders.getVisibleStartRow();
        let endRow = this.rowHeaders.length - 1;
        const viewportBottom = this.rowHeaders.offsetHeight + this.visibleView.sheetHeight / this.zoom;
        while (endRow > startRow) {
            const row = this.rowHeaders.getAt(endRow - 1);
            if (!row) break;
            if (row.isHidden) { endRow--; continue; }  // 跳过隐藏行
            if (row.top + row.height < viewportBottom) break;  // 行底边缘进入视口
            endRow--;
        }
        return { startRow, endRow: Math.max(startRow, endRow) };
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
        const cellIndex = this.values.findIndex(v => v.cell === cell);
        if (cellIndex >= 0) {
            this.values[cellIndex].chars = [];
            this.values[cellIndex].backgroundColor = undefined;
            this.values[cellIndex].borderTopWidth = undefined;
            this.values[cellIndex].borderBottomWidth = undefined;
            this.values[cellIndex].borderLeftWidth = undefined;
            this.values[cellIndex].borderRightWidth = undefined;
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
            if (mergedCells.find(v => v.includes(startCell))) {
                return true;
            }
            const cell = this.values.find(v => v.cell === startCell);
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
        try {
            const sourceCell = this.findSourceCell(minCol, maxCol, minRow, maxRow);

            for (let row = minRow; row <= maxRow; row++) {
                for (let col = minCol; col <= maxCol; col++) {
                    const cellName = this.getCellName(col, row);
                    const cellIndex = this.values.findIndex(c => c.cell === cellName);

                    if (col === minCol && row === minRow) {
                        if (cellIndex >= 0) {
                            if (sourceCell) {
                                Object.assign(this.values[cellIndex], sourceCell);
                                this.values[cellIndex].cell = cellName;
                            }
                            this.values[cellIndex].colspan = colspan;
                            this.values[cellIndex].rowspan = rowspan;
                        } else {
                            if (sourceCell) {
                                sourceCell.cell = cellName;
                                sourceCell.colspan = colspan;
                                sourceCell.rowspan = rowspan;
                                this.values.push(sourceCell);
                            } else {
                                const cell = new Cell();
                                cell.cell = cellName;
                                cell.colspan = colspan;
                                cell.rowspan = rowspan;
                                this.values.push(cell);
                            }
                        }
                    } else {
                        if (cellIndex >= 0) {
                            this.clearCellValue(cellName);
                            this.values.splice(cellIndex, 1);
                        }
                    }
                }
            }
        } finally {
            this.commitValues();
        }
    }

    /**
     * 取消合并单元格
     * 将当前选中的合并单元格恢复为独立的单元格
     * 合并区域的内容和文字样式保留在左上角单元格中
     * 背景色和边框会应用到所有分解出来的单元格
     */
    public unmergeCells(): void {
        if (!this.selection) return;
        const mergedCells = this.getMergedCells();
        const [startCell] = this.selection.split(':');
        const cell = this.values.find(c => c.cell === startCell);
        
        let topLeftCell: Cell | null = null;
        let topLeftCol = this.getCellColAndRow(startCell).col;
        let topLeftRow = this.getCellColAndRow(startCell).row;
        
        if (cell?.colspan === 0 && cell?.rowspan === 0) {
            for (const mergedGroup of mergedCells) {
                if (mergedGroup.includes(startCell)) {
                    const topLeftCellName = mergedGroup.split(':')[0];
                    const topLeftPos = this.getCellColAndRow(topLeftCellName);
                    topLeftCol = topLeftPos.col;
                    topLeftRow = topLeftPos.row;
                    topLeftCell = this.values.find(c => c.cell === topLeftCellName) || null;
                    break;
                }
            }
        } else if (cell && (cell.colspan > 1 || cell.rowspan > 1)) {
            topLeftCell = cell;
        }
        
        if (!topLeftCell) return;
        
        const colspan = topLeftCell.colspan || 1;
        const rowspan = topLeftCell.rowspan || 1;
        
        if (colspan === 1 && rowspan === 1) return;
        
        // 提取需要复制到所有单元格的样式（背景色和边框）
        const sharedStyles: Partial<Cell> = {
            backgroundColor: topLeftCell.backgroundColor,
            borderTopWidth: topLeftCell.borderTopWidth,
            borderBottomWidth: topLeftCell.borderBottomWidth,
            borderLeftWidth: topLeftCell.borderLeftWidth,
            borderRightWidth: topLeftCell.borderRightWidth
        };
        
        try {
            topLeftCell.colspan = 1;
            topLeftCell.rowspan = 1;
            
            for (let r = 0; r < rowspan; r++) {
                for (let c = 0; c < colspan; c++) {
                    const currentCol = topLeftCol + c;
                    const currentRow = topLeftRow + r;
                    const currentCellName = this.getCellName(currentCol, currentRow);
                    const existingCell = this.values.find(v => v.cell === currentCellName);
                    
                    if (r === 0 && c === 0) {
                        // 左上角单元格：保留内容和所有样式
                        if (existingCell) {
                            existingCell.colspan = undefined;
                            existingCell.rowspan = undefined;
                        }
                    } else {
                        // 其他单元格：只保留背景色和边框，内容清空
                        if (existingCell) {
                            existingCell.colspan = undefined;
                            existingCell.rowspan = undefined;
                            existingCell.chars = [];
                            Object.assign(existingCell, sharedStyles);
                        } else {
                            const cell = new Cell();
                            cell.cell = currentCellName;
                            cell.colspan = undefined;
                            cell.rowspan = undefined;
                            cell.chars = [];
                            Object.assign(cell, sharedStyles);
                            this.values.push(cell);
                        }
                    }
                }
            }
        } finally {
            this.commitValues();
        }
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
        try {
            for (let col = minCol; col <= maxCol; col++) {
                const sourceCell = this.findSourceCell(col, col, minRow, maxRow);

                for (let row = minRow; row <= maxRow; row++) {
                    const cellName = this.getCellName(col, row);
                    const cellIndex = this.values.findIndex(c => c.cell === cellName);

                    if (row === minRow) {
                        // 当前列的第一行：设置 rowspan
                        if (cellIndex >= 0) {
                            if (sourceCell) {
                                Object.assign(this.values[cellIndex], sourceCell);
                                this.values[cellIndex].cell = cellName;
                            }
                            this.values[cellIndex].colspan = 1;
                            this.values[cellIndex].rowspan = rowspan;
                        } else {
                            if (sourceCell) {
                                sourceCell.cell = cellName;
                                sourceCell.colspan = 1;
                                sourceCell.rowspan = rowspan;
                                this.values.push(sourceCell);
                            } else {
                                const cell = new Cell();
                                cell.cell = cellName;
                                cell.colspan = 1;
                                cell.rowspan = rowspan;
                                this.values.push(cell);
                            }
                        }
                    } else {
                        // 当前列的其他行：清除内容和样式
                        if (cellIndex >= 0) {
                            this.clearCellValue(cellName);
                            this.values.splice(cellIndex, 1);
                        }
                    }
                }
            }
        } finally {
            this.commitValues();
        }
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
        try {
            for (let row = minRow; row <= maxRow; row++) {
                const sourceCell = this.findSourceCell(minCol, maxCol, row, row);

                for (let col = minCol; col <= maxCol; col++) {
                    const cellName = this.getCellName(col, row);
                    const cellIndex = this.values.findIndex(c => c.cell === cellName);

                    if (col === minCol) {
                        // 当前行的第一列：设置 colspan
                        if (cellIndex >= 0) {
                            if (sourceCell) {
                                Object.assign(this.values[cellIndex], sourceCell);
                                this.values[cellIndex].cell = cellName;
                            }
                            this.values[cellIndex].colspan = colspan;
                            this.values[cellIndex].rowspan = 1;
                        } else {
                            if (sourceCell) {
                                sourceCell.cell = cellName;
                                sourceCell.colspan = colspan;
                                sourceCell.rowspan = 1;
                                this.values.push(sourceCell);
                            } else {
                                const cell = new Cell();
                                cell.cell = cellName;
                                cell.colspan = colspan;
                                cell.rowspan = 1;
                                this.values.push(cell);
                            }
                        }
                    } else {
                        // 当前行的其他列：清除内容和样式
                        if (cellIndex >= 0) {
                            this.clearCellValue(cellName);
                            this.values.splice(cellIndex, 1);
                        }
                    }
                }
            }
        } finally {
            this.commitValues();
        }
    }

    /**
     * 获取所有合并单元格的名称
     * @returns {string[]} 合并单元格的名称数组
     */
    public getMergedCells(): string[] {
        const mergeCells: string[] = [];
        for (let i = 0; i < this.values.length; i++) {
            const cell = this.values[i];
            const colspan = cell.colspan || 1;
            const rowspan = cell.rowspan || 1;
            if (colspan === 1 && rowspan === 1) continue;
            const { col, row } = this.getCellColAndRow(cell.cell);
            mergeCells.push(`${cell.cell}:${this.getColName(col + colspan - 1)}${row + rowspan - 1}`);
        }
        return mergeCells;
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
        const cellData = this.values.find(c => c.cell === cell);
        if (cellData && (cellData.colspan > 1 || cellData.rowspan > 1)) {
            return { col, row };
        }

        // 若不在合并单元格中，检查是否位于某个合并区域内
        const mergedCells = this.getMergedCells();
        for (const merged of mergedCells) {
            const [startCell, endCell] = merged.split(':');
            const { col: startCol, row: startRow } = this.getCellColAndRow(startCell);
            const { col: endCol, row: endRow } = this.getCellColAndRow(endCell);
            if (col >= startCol && col <= endCol && row >= startRow && row <= endRow) {
                const { col: topLeftCol, row: topLeftRow } = this.getCellColAndRow(startCell);
                return {
                    col: topLeftCol,
                    row: topLeftRow
                };
            }
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
        const [start, end] = this.selection ? this.selection.split(':') : [this.activedCell, this.activedCell];
        const { col: startCol0, row: startRow0 } = this.getCellColAndRow(start);
        const { col: endCol0, row: endRow0 } = this.getCellColAndRow(end);
        const minCol = Math.min(startCol0, endCol0);
        const maxCol = Math.max(startCol0, endCol0);
        const minRow = Math.min(startRow0, endRow0);
        const maxRow = Math.max(startRow0, endRow0);
        // 收缩到范围内的可见行列（隐藏项不占位，选区框不覆盖其尺寸）
        let startCol = -1;
        let endCol = -1;
        let startRow = -1;
        let endRow = -1;
        for (let c = minCol; c <= maxCol; c++) {
            if (!this.colHeaders.getAt(c - 1)?.isHidden) {
                if (startCol < 0) startCol = c;
                endCol = c;
            }
        }
        for (let r = minRow; r <= maxRow; r++) {
            if (!this.rowHeaders.getAt(r - 1)?.isHidden) {
                if (startRow < 0) startRow = r;
                endRow = r;
            }
        }
        // 选区完全落在隐藏行列上：返回空矩形
        if (startCol < 0 || startRow < 0) {
            return { left: 0, top: 0, width: 0, height: 0 };
        }
        const { left: startLeft, top: startTop } = this.calculateCellRect(startCol, startRow);
        const { left: endLeft, top: endTop, width: endWidth, height: endHeight } = this.calculateCellRect(endCol, endRow);
        const width = endLeft + endWidth - startLeft;
        const height = endTop + endHeight - startTop;

        return { left: startLeft, top: startTop, width, height };
    }

    /**
     * 设置选中的单元格范围
     * @param {number} startCol - 起始列索引
     * @param {number} startRow - 起始行索引
     * @param {number} endCol - 结束列索引
     * @param {number} endRow - 结束行索引
     */
    public setSelectedRange(startCol: number, startRow: number, endCol: number, endRow: number) {
        let maxCol = Math.max(startCol, endCol);
        let minCol = Math.min(startCol, endCol);
        let maxRow = Math.max(startRow, endRow);
        let minRow = Math.min(startRow, endRow);
        const mergedCells = this.getMergedCells();
        for (const mergedCell of mergedCells) {
            const startCell = mergedCell.split(':')[0];
            const { col: col0, row: row0 } = this.getCellColAndRow(startCell);
            const endCell = mergedCell.split(':')[1];
            const { col: col1, row: row1 } = this.getCellColAndRow(endCell);
            loop:
            for (let c = col0; c <= col1; c++) {
                for (let r = row0; r <= row1; r++) {
                    if (c >= minCol && c <= maxCol && r >= minRow && r <= maxRow) {
                        minCol = Math.min(minCol, col0);
                        minRow = Math.min(minRow, row0);
                        maxCol = Math.max(maxCol, col1);
                        maxRow = Math.max(maxRow, row1);
                        break loop;
                    }
                }
            }         
        }
        if (this.startLocation !== MouseLocation.IN_ANCHOR) {
            this.selection = `${this.getColName(minCol)}${minRow}:${this.getColName(maxCol)}${maxRow}`;
            // this.anchorSelection = `${this.getColName(minCol)}${minRow}:${this.getColName(maxCol)}${maxRow}`;
        }
    }

    /**
     * 设置选中的单元格
     * @param {string} location - 选中的单元格名称
     * @param {number} col - 单元格的列索引
     * @param {number} row - 单元格的行索引
     */
    public setSelection(location: string, col: number, row: number) {
        const { col: startCol, row: startRow } = this.getCellColAndRow(this.activedCell);
        if (location.includes('inCol')) {
            if (this.startLocation === MouseLocation.IN_COL_HEADER) {
                this.setSelectedRange(startCol, 1, col, this.rowHeaders.length);
            } else if (this.startLocation === MouseLocation.IN_ROW_HEADER) {
                this.setSelectedRange(1, startRow, this.colHeaders.length, 1);
            } else {
                this.setSelectedRange(startCol, startRow, col, 1);
            }
        } else if (location.includes('inRow')) {
            if (this.startLocation === MouseLocation.IN_COL_HEADER) {
                this.setSelectedRange(startCol, 1, 1, this.rowHeaders.length);
            } else if (this.startLocation === MouseLocation.IN_ROW_HEADER) {
                this.setSelectedRange(1, startRow, this.colHeaders.length, row);
            } else {
                this.setSelectedRange(startCol, startRow, 1, row);
            }
        } else {
            if (this.startLocation === MouseLocation.IN_COL_HEADER) {
                this.setSelectedRange(startCol, 1, col, this.rowHeaders.length);
            } else if (this.startLocation === MouseLocation.IN_ROW_HEADER) {
                this.setSelectedRange(1, startRow, this.colHeaders.length, row);
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
        const { left: endLeft, top: endTop, width: endWidth, height: endHeight } = this.calculateCellRect(col, row);
        const width = DEFAULT_SELECTED_BORDER_WIDTH * 2;
        const height = DEFAULT_SELECTED_BORDER_WIDTH * 2;
        const left = endLeft + endWidth - width + 2;
        const top = endTop + endHeight - height + 2;

        return { col, row, left, top, width, height };
    }

    /**
     * 调整列宽以适应内容宽度
     * @param col 列索引
     */
    public adjustColumnWidth(col: number, newWidth: number): void {
        if (col < 1 || col > this.colHeaders.length) return;
        const cols = this.colHeaders.length;
        const minWidth = 24;
        newWidth = Math.max(minWidth, newWidth);
        const colHeader = this.colHeaders.getAt(col - 1); // 获取列头部
        if (!colHeader) return;
        colHeader.width = newWidth;

        for (let i = col; i < cols; i++) {
            const currentColHeader = this.colHeaders.getAt(i - 1); // 获取当前列头部
            const nextColHeader = this.colHeaders.getAt(i); // 获取下一个列头部
            // 紧凑排列：隐藏列不占位（前进量为 0），保持隐藏块与其后可见列共享同一分割位置，
            // 避免拖拽/双击调整列宽时把隐藏列的宽度累加回去而重新显露
            nextColHeader.left = currentColHeader.left
                + (currentColHeader.isHidden ? 0 : currentColHeader.width);
        }
        this.emit(DataEvents.COLUMNS_CHANGED);
    }

    /**
     * 调整行高以适应内容高度
     * @param row 行索引
     */
    public adjustRowHeight(row: number, newHeight: number): void {
        if (row < 1 || row > this.rowHeaders.length) return;
        const rows = this.rowHeaders.length;
        const minHeight = 24;
        newHeight = Math.max(minHeight, newHeight);
        const rowHeader = this.rowHeaders.getAt(row - 1); // 获取行头部
        if (!rowHeader) return;
        rowHeader.height = newHeight;

        for (let i = row; i < rows; i++) {
            const currentRowHeader = this.rowHeaders.getAt(i - 1); // 获取当前行头部
            const nextRowHeader = this.rowHeaders.getAt(i); // 获取下一个行头部
            // 紧凑排列：隐藏行不占位（前进量为 0），保持隐藏块与其后可见行共享同一分割位置，
            // 避免拖拽/双击调整行高时把隐藏行的高度累加回去而重新显露
            nextRowHeader.top = currentRowHeader.top
                + (currentRowHeader.isHidden ? 0 : currentRowHeader.height);
        }
        this.emit(DataEvents.ROWS_CHANGED);
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
                const _cell = this.values.find(c => c.cell === `${this.getColName(colOff + startCol)}${rowOff + startRow}`);
                if ( _cell ) {
                    _cell.filter = undefined; // 移除筛选条件
                    selectedCells.push(_cell);
                }
            }
        }

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
                const _cell = this.values.find(c => c.cell === `${this.getColName(colOff + startCol)}${rowOff + startRow}`);
                if ( _cell ) {
                    this.values.splice(this.values.indexOf(_cell), 1); // 从数据集合中移除选中的单元格
                    selectedCells.push(_cell);
                }
            }
        }

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
            // 源可见位置与目标可见位置按序一一配对（紧凑排列）
            for (let i = 0; i < colOffsets.length && i < targetColOffsets.length; i++) {
                for (let j = 0; j < rowOffsets.length && j < targetRowOffsets.length; j++) {
                    const tgtCol = activedCol + targetColOffsets[i];
                    const tgtRow = activedRow + targetRowOffsets[j];
                    const cell = this.values.find(c => c.cell === `${this.getColName(tgtCol)}${tgtRow}`);
                    const item = clipboard.cells.find(c => c.cell === `${this.getColName(startCol + colOffsets[i])}${startRow + rowOffsets[j]}`);
                    if (item) {
                        if (cell) {
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
                            cell.colspan = item.colspan;
                            cell.rowspan = item.rowspan;
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
                        } else {
                            const _cell = new Cell();
                            _cell.cell = `${this.getColName(tgtCol)}${tgtRow}`;
                            _cell.colspan = item.colspan;
                            _cell.rowspan = item.rowspan;
                            _cell.borderTopWidth = item.borderTopWidth;
                            _cell.borderBottomWidth = item.borderBottomWidth;
                            _cell.borderLeftWidth = item.borderLeftWidth;
                            _cell.borderRightWidth = item.borderRightWidth;
                            _cell.textAlign = item.textAlign;
                            _cell.alignItems = item.alignItems;
                            _cell.backgroundColor = item.backgroundColor;
                            _cell.filter = item.filter;
                            _cell.wrap = item.wrap;
                            _cell.letterSpacing = item.letterSpacing;
                            _cell.lineSpacing = item.lineSpacing;
                            _cell.chars = item.chars.map(c => {
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
                            });
                            this.values.push(_cell);
                        }                        
                    }
                }
            }
        } else if (clipboard && clipboard.type === 'char' && clipboard.chars) {
            const cell = this.values.find(c => c.cell === this.activedCell);
            if (cell){
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
            } else {
                const _cell = new Cell();
                _cell.cell = this.activedCell;
                _cell.chars = clipboard.chars.map(c => {
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
                });
                this.values.push(_cell);
            }        
        }

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
                const cell = this.values.find(c => c.cell === `${this.getColName(tgtCol)}${tgtRow}`);
                const item = clipboard.cells.find(c => c.cell === `${this.getColName(startCol + colOffsets[i])}${startRow + rowOffsets[j]}`);
                if (item) {
                    if (cell) {
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
                    } else {
                        const _cell = new Cell();
                        _cell.cell = `${this.getColName(tgtCol)}${tgtRow}`;
                        _cell.chars = item.chars.map(c => {
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
                        });
                        this.values.push(_cell);
                    }                        
                }
            }
        }

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
                    const cell = this.values.find(c => c.cell === `${this.getColName(tgtCol)}${tgtRow}`);
                    const item = clipboard.cells.find(c => c.cell === `${this.getColName(startCol + colOffsets[i])}${startRow + rowOffsets[j]}`);
                    if (item) {
                        // 只拷贝字符内容，保留目标原有格式
                        const newChars = item.chars.map(c => {
                            const _char = new Char();
                            _char.char = c.char;
                            return _char;
                        });
                        if (cell) {
                            cell.chars = newChars;
                        } else {
                            const _cell = new Cell();
                            _cell.cell = `${this.getColName(tgtCol)}${tgtRow}`;
                            _cell.chars = newChars;
                            this.values.push(_cell);
                        }
                    }
                }
            }
        } else if (clipboard && clipboard.type === 'char' && clipboard.chars) {
            const cell = this.values.find(c => c.cell === this.activedCell);
            if (cell) {
                cell.chars = clipboard.chars.map(c => {
                    const _char = new Char();
                    _char.char = c.char;
                    return _char;
                });
            }
        }
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
        // 源可见位置与目标可见位置按序一一配对（紧凑排列）
        for (let i = 0; i < colOffsets.length && i < targetColOffsets.length; i++) {
            for (let j = 0; j < rowOffsets.length && j < targetRowOffsets.length; j++) {
                const tgtCol = activedCol + targetColOffsets[i];
                const tgtRow = activedRow + targetRowOffsets[j];
                const cell = this.values.find(c => c.cell === `${this.getColName(tgtCol)}${tgtRow}`);
                const item = clipboard.cells.find(c => c.cell === `${this.getColName(startCol + colOffsets[i])}${startRow + rowOffsets[j]}`);
                if (item) {
                    if (cell) {
                        cell.isEdit = item.isEdit;
                        cell.colspan = item.colspan;
                        cell.rowspan = item.rowspan;
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
                    } else {
                        const _cell = new Cell();
                        _cell.cell = `${this.getColName(tgtCol)}${tgtRow}`;
                        _cell.colspan = item.colspan;
                        _cell.rowspan = item.rowspan;
                        _cell.fontFamily = item.fontFamily;
                        _cell.fontSize = item.fontSize;
                        _cell.fontWeight = item.fontWeight;
                        _cell.fontStyle = item.fontStyle;
                        _cell.fontColor = item.fontColor;
                        _cell.underline = item.underline;
                        _cell.strikethrough = item.strikethrough;
                        _cell.borderTopWidth = item.borderTopWidth;
                        _cell.borderBottomWidth = item.borderBottomWidth;
                        _cell.borderLeftWidth = item.borderLeftWidth;
                        _cell.borderRightWidth = item.borderRightWidth;
                        _cell.textAlign = item.textAlign;
                        _cell.alignItems = item.alignItems;
                        _cell.backgroundColor = item.backgroundColor;
                        _cell.filter = item.filter;
                        _cell.wrap = item.wrap;
                        _cell.letterSpacing = item.letterSpacing;
                        _cell.lineSpacing = item.lineSpacing;
                        this.values.push(_cell);
                    }                        
                }
            }
        }

        this.commitValues();
    }

    /**
     * 复制选中的字符内容
     */
    public copyChars(): void {
        const cell = this.values.find(c => c.cell === this.activedCell);
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
        const cell = this.values.find(c => c.cell === this.activedCell);
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
        const cell = this.values.find(c => c.cell === this.activedCell);
        let caretAbsolutePosition: number;
        if (cell){
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
            caretAbsolutePosition = insertPos + chars.length;
        } else {
            const _cell = new Cell();
            _cell.cell = this.activedCell;
            _cell.chars = chars.map(c => {
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
            });
            this.values.push(_cell);
            // 新建单元格，光标在所有字符之后
            caretAbsolutePosition = chars.length;
        }
        this.selectionOnEditting = null;
        this.commitValues();
        return caretAbsolutePosition;
    }
    
    /**
     * 设置选中单元格的字体
     * @param fontStyle 字体样式
     */
    public setSelectedCellsFont(fontStyle: FontStyle): void {
        const selectedCells = this.getSelectedCells();
        selectedCells.forEach(_cell => {
            const cell = this.values.find(c => c.cell === _cell);
            if (cell) {
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
            } else {
                const newCell = new Cell();
                newCell.cell = _cell;
                newCell.fontFamily = fontStyle.fontFamily;
                newCell.fontSize = fontStyle.fontSize;
                newCell.fontWeight = fontStyle.fontWeight;
                newCell.fontStyle = fontStyle.fontStyle;
                newCell.fontColor = fontStyle.fontColor;
                newCell.underline = fontStyle.underline;
                newCell.strikethrough = fontStyle.strikethrough;
                newCell.textAlign = fontStyle.textAlign as TextAlign;
                newCell.alignItems = fontStyle.alignItems as VerticalAlign;
                newCell.backgroundColor = fontStyle.backgroundColor;
                newCell.letterSpacing = fontStyle.letterSpacing;
                newCell.lineSpacing = fontStyle.lineSpacing;
                newCell.wrap = fontStyle.wrap;
                this.values.push(newCell);
            }
        });
        this.commitValues();
    }

    /**
     * 设置选中字符的字体
     * @param fontStyle 字体样式
     */
    public setSelectedCharsFont(fontStyle: FontStyle): void {
        const startIndex = this.getSelectedCharsStartIndex();
        const endIndex = this.getSelectedCharsEndIndex();
        const chars = this.values.find(c => c.cell === this.activedCell)?.chars || [];
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
            const targetCell = this.values.find(c => c.cell === start);
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

                let targetCell = this.values.find(c => c.cell === targetCellName);
                if (!targetCell) {
                    targetCell = new Cell();
                    targetCell.cell = targetCellName;
                    this.values.push(targetCell);
                }
                this._applyFormatToCell(targetCell, sourceCell);
            }
        }

        if (this.brushMode === 'single') this.brushMode = null;
        this.commitValues();
    }

    /**
     * 设置选中单元格的边框样式
     * @param cellBorderStyle 边框样式
     */
    public setSelectedCellsBorder(cellBorderStyle: CellBorderStyle): void {
        const selectedCells = this.getSelectedCells();
        const { col: startCol, row: startRow } = this.getCellColAndRow(selectedCells.at(0));
        const { col: endCol, row: endRow } = this.getCellColAndRow(selectedCells.at(-1));
        selectedCells.forEach(_cell => {      
            const {col, row} = this.getCellColAndRow(_cell);
            let borderTopWidth, borderBottomWidth, borderLeftWidth, borderRightWidth;
            if (cellBorderStyle === CellBorderStyle.BorderNone) {
                borderTopWidth = undefined;
                borderBottomWidth = undefined;
                borderLeftWidth = undefined;
                borderRightWidth = undefined;
            } else if (cellBorderStyle === CellBorderStyle.BorderAll) {
                borderTopWidth = 1;
                borderBottomWidth = 1;
                borderLeftWidth = 1;
                borderRightWidth = 1;
            } else if (cellBorderStyle === CellBorderStyle.BorderOuter) {
                if (startRow === row) borderTopWidth = 1;
                if (endRow === row) borderBottomWidth = 1;
                if (startCol === col) borderLeftWidth = 1;
                if (endCol === col) borderRightWidth = 1;
            } else if (cellBorderStyle === CellBorderStyle.BorderWideOuter) {
                if (startRow === row) borderTopWidth = 2;
                if (endRow === row) borderBottomWidth = 2;
                if (startCol === col) borderLeftWidth = 2;
                if (endCol === col) borderRightWidth = 2;
            } else if (cellBorderStyle === CellBorderStyle.BorderBottom) {
                if (endRow === row) borderBottomWidth = 1;
            } else if (cellBorderStyle === CellBorderStyle.BorderTop) {
                if (startRow === row) borderTopWidth = 1;
            } else if (cellBorderStyle === CellBorderStyle.BorderLeft) {
                if (startCol === col) borderLeftWidth = 1;
            } else if (cellBorderStyle === CellBorderStyle.BorderRight) {
                if (endCol === col) borderRightWidth = 1;
            }

            const cell = this.values.find(c => c.cell === _cell);
            if (cell) {
                cell.borderTopWidth = borderTopWidth;
                cell.borderBottomWidth = borderBottomWidth;
                cell.borderLeftWidth = borderLeftWidth;
                cell.borderRightWidth = borderRightWidth;
            } else {
                const newCell = new Cell();
                newCell.cell = _cell;
                newCell.borderTopWidth = borderTopWidth;
                newCell.borderBottomWidth = borderBottomWidth;
                newCell.borderLeftWidth = borderLeftWidth;
                newCell.borderRightWidth = borderRightWidth;
                this.values.push(newCell);
            }
        })
        this.commitValues();
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
     * 撤销/重做：排序前深拷贝当前 values 作为 before 快照，
     * 抑制 setter 自动记录（因 remapMatrixCellNames 会就地修改 Cell 名称，
     * setter 的 oldValue 不再是真正的排序前状态），排序后手动推送快照。
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

        // 排序前深拷贝，作为撤销快照的 before（必须在 remap 之前捕获）
        const beforeSnapshot = this.cloneValues(this.values);
        const matrix = this.matrixValues();
        const _matrix = this.matrixSorter.init(matrix).sort(sortConfig).getData();
        this.remapMatrixCellNames(_matrix);
        const sortedValues = _matrix.flat().filter((c): c is Cell => c !== undefined);

        // 抑制 setter 自动记录（oldValue 中的 Cell 名称已被 remap 修改，不再可靠）
        this._suppressValuesUndo = true;
        try {
            this.values = sortedValues;
        } finally {
            this._suppressValuesUndo = false;
        }

        // 手动推送排序的撤销快照
        this.undoManager.pushValuesSnapshot(
            beforeSnapshot,
            this.cloneValues(this.values),
            ascending ? '升序排序' : '降序排序'
        );
        this._lastValuesSnapshot = this.cloneValues(this.values);
        this.emit(DataEvents.VALUES_CHANGED, this._values);
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
        const hasFilters = this.values.some((v: Cell | undefined) => v?.filter !== undefined);
        if (hasFilters) {
            this.runWithFullStateUndo('清除筛选按钮', () => {
                for (const cell of this.values) {
                    if (cell.filter !== undefined) cell.filter = undefined;
                }
                this.emit(DataEvents.VALUES_CHANGED, this._values);
            });
            return;
        }

        const [start, end] = this._selection.split(':');
        const { col: startCol, row: startRow } = this.getCellColAndRow(start);
        const { col: endCol, row: endRow } = this.getCellColAndRow(end);
        const matrix = this.matrixValues();
        const rows = matrix[startRow - 1];
        if (start !== end && startCol !== endCol && startRow === endRow && rows.some((v: Cell | undefined) => v !== undefined)) {
            // 为选中行范围内的有效列逐列添加筛选按钮（一步撤销）
            this.runWithFullStateUndo('添加筛选按钮', () => {
                for (let c = startCol - 1; c < endCol; c++) {
                    if (!this.validataColInFilter(c)) continue;
                    const cell = this.getCellName(c + 1, startRow);
                    const _cell = this.values.find((v: Cell | undefined) => v.cell === cell);
                    if (_cell) {
                        _cell.filter = startRow;
                    } else {
                        const newCell = new Cell();
                        newCell.cell = cell;
                        newCell.filter = startRow;
                        this.values.push(newCell);
                    }
                }
                this.emit(DataEvents.VALUES_CHANGED, this._values);
            });
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
     * 验证列是否有效
     * @param {number} col - 列索引
     * @returns {boolean} 是否有效
     */
    private validataColInFilter(col: number): boolean {
        const matrix = this.matrixValues();
        if (col < 0 || col >= matrix[0].length) return false;
        if (this.colHeaders.getAt(col).isHidden) return false; // 跳过隐藏列
        const cols = [];
        for (let r = 0; r < matrix.length; r++) {
            const row = matrix[r];
            if (row) {
                if (row[col] !== undefined) cols.push(row[col]);
            }
        }
        if (cols.length > 0) return true;
        return false;
    }

    /**
     * 验证过滤范围是否有效
     * @returns {boolean} 是否有效
     */
    public validataFilter(): boolean {
        const filters = this.values.filter((v: Cell | undefined) => v.filter !== undefined);
        if (filters.length > 0) return true;
        return false;
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
    private _findFirstVisibleCol(): number {
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
    private _findLastVisibleCol(): number {
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
    private _findFirstVisibleRow(): number {
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
    private _findLastVisibleRow(): number {
        for (let r = this.rowHeaders.length - 1; r >= 0; r--) {
            if (!this.rowHeaders.getAt(r).isHidden) return r + 1;
        }
        return 0;
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
                    const srcCell = this.values.find(v => v.cell === this.getCellName(srcCol, r));
                    if (srcCell) srcCell.reset();
                    continue;
                }
                const tgtCol = visCols[tgtRank];
                const srcName = this.getCellName(srcCol, r);
                const tgtName = this.getCellName(tgtCol, r);
                const srcCell = this.values.find(v => v.cell === srcName);
                const tgtCell = this.values.find(v => v.cell === tgtName);
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
                    const srcCell = this.values.find(v => v.cell === this.getCellName(c, srcRow));
                    if (srcCell) srcCell.reset();
                    continue;
                }
                const tgtRow = visRows[tgtRank];
                const srcName = this.getCellName(c, srcRow);
                const tgtName = this.getCellName(c, tgtRow);
                const srcCell = this.values.find(v => v.cell === srcName);
                const tgtCell = this.values.find(v => v.cell === tgtName);
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
     * 设置指定范围行的隐藏状态
     * @param {number} startRow - 起始行号（1-based，含）
     * @param {number} endRow - 结束行号（1-based，含）
     * @param {boolean} hidden - 是否隐藏
     */
    public setRowsHidden(startRow: number, endRow: number, hidden: boolean): void {
        for (let r = startRow; r <= endRow; r++) {
            const rh = this.rowHeaders.getAt(r - 1);
            if (rh) rh.isHidden = hidden;
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
            // 隐藏：可见总高不足一屏时逐行补足，新行真实加入 rowHeaders（按可见总高判断，兼容自定义行高）
            while (this.rowHeaders.allRowHeight < viewportHeight) {
                this.rowHeaders.addRow(this.rowHeaders.length, this.rowHeaders.length + 1);
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
     * 收缩尾部无内容的行：从 rowHeaders 末尾删除无任何单元格内容的行，
     * 直到遇到有内容的行，或删除后可见总高将不足以填满视口为止。
     * 用于取消隐藏后清理历史补足/扩行产生的多余空行。
     * @param {number} viewportHeight - 视口高度（数据坐标，已除 zoom）
     * @param {Set<number>} protectedRows - 受保护的行号（1-based），本次操作涉及的行不参与收缩
     * @private
     * @returns {void}
     */
    private trimRowsToViewport(viewportHeight: number, protectedRows: Set<number>): void {
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
            // 隐藏：可见总宽不足一屏时逐列补足，新列真实加入 colHeaders（按可见总宽判断，兼容自定义列宽）
            while (this.colHeaders.allColWidth < viewportWidth) {
                this.colHeaders.addCol(this.colHeaders.length, this.colHeaders.length + 1);
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
    private trimColsToViewport(viewportWidth: number, protectedCols: Set<number>): void {
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
                const cell = this.values.find(v => v.cell === this.getCellName(c, r));
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
                const cell = this.values.find(v => v.cell === this.getCellName(c, r));
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
                const srcCell = this.values.find(v => v.cell === srcName);
                const tgtCell = this.values.find(v => v.cell === tgtName);
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
                const cell = this.values.find(v => v.cell === this.getCellName(c, r));
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
                const srcCell = this.values.find(v => v.cell === srcName);
                const tgtCell = this.values.find(v => v.cell === tgtName);
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
                const cell = this.values.find(v => v.cell === this.getCellName(c, r));
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
                const cellName = this.getCellName(c, r);
                let cell = this.values.find(v => v.cell === cellName);
                if (cell) {
                    cell.numberFormat = format;
                } else {
                    cell = new Cell();
                    cell.cell = cellName;
                    cell.numberFormat = format;
                    this.values.push(cell);
                }
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
                const cell = this.values.find(v => v.cell === this.getCellName(c, r));
                if (cell) cell.borderColor = color;
            }
        }
    }

    /**
     * 将数据集合转换为JSON表示
     * @returns {string} 数据集合的JSON表示
     */
    public toJSON(): string {
        return JSON.stringify({
            activedCell: this.activedCell,
            selection: this._selection,
            printSetting: this._printSetting,
            colHeaders: this.colHeaders.colHeaders,
            offsetWidth: this.colHeaders.offsetWidth,
            allColWidth: this.colHeaders.allColWidth,
            rowHeaders: this.rowHeaders.rowHeaders,
            offsetHeight: this.rowHeaders.offsetHeight,
            allRowHeight: this.rowHeaders.allRowHeight,
            showGridLines: this.showGridLines,
            zoom: this.zoom,
            values: this.values.map((v: Cell) => v.toJSON()),
        });
    }

    /**
     * 从JSON表示创建数据集合
     * @param {string} json - 数据集合的JSON表示
     * @returns {Cell[]} 创建单元格实例数组
     */
    public static fromJSON(json: string): {activedCell: string,
        selection: string,
        printSetting: PrinterSetting,
        colHeaders: ColHeader[],
        offsetWidth: number,
        allColWidth: number,
        rowHeaders: RowHeader[],
        offsetHeight: number,
        allRowHeight: number,
        showGridLines: boolean,
        zoom: number,
        values: Cell[]} {
            const parsed = JSON.parse(json);
        return {
            activedCell: parsed.activedCell || 'A1',
            selection: parsed.selection || `${parsed.activedCell || 'A1'}:${parsed.activedCell || 'A1'}`,
            printSetting: parsed.printSetting || {},
            colHeaders: parsed.colHeaders || [],
            offsetWidth: parsed.offsetWidth || 0,
            allColWidth: parsed.allColWidth || 0,
            rowHeaders: parsed.rowHeaders || [],
            offsetHeight: parsed.offsetHeight || 0,
            allRowHeight: parsed.allRowHeight || 0,
            showGridLines: parsed.showGridLines || true,
            zoom: parsed.zoom || 1,
            values: (parsed.values || []).map((itemJSON: any) => {
                const item = JSON.parse(itemJSON);
                const cell = new Cell();
                cell.cell = item.cell || '';
                cell.isEdit = item.isEdit || false;
                cell.colspan = item.colspan;
                cell.rowspan = item.rowspan;
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
                const chars = JSON.parse(item.chars || '[]');
                cell.chars = chars.map((charItemJSON: any) => {
                    const charItem = JSON.parse(charItemJSON);
                    const char = new Char();
                    char.char = charItem.char || charItem._char || '';
                    char.fontFamily = charItem.fontFamily;
                    char.fontSize = charItem.fontSize;
                    char.fontWeight = charItem.fontWeight;
                    char.fontStyle = charItem.fontStyle;
                    char.fontColor = charItem.fontColor;
                    char.underline = charItem.underline;
                    char.strikethrough = charItem.strikethrough;
                    return char;
                });
                return cell;
            })
        }
    }
    
    /**
     * 通知工具栏状态更新
     */
    public notifyToolbarStateChange() {
        this.emit(DataEvents.CURSOR_STATE_CHANGED);
    }

}
