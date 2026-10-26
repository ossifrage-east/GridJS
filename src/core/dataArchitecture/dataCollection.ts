import { ColHeader, ColHeaders, RowHeader, RowHeaders, encodeColAxis, encodeRowAxis, decodeColAxis, decodeRowAxis, SparseAxis } from "./header";
import { Cell, CellStyleTable, TextAlign, VerticalAlign } from "./cell";
import { LineText, SelectionRange, FontStyle, BorderStyle, CellBorderStyle, CursorPosition, } from "./cellContent";
import { Char } from "./char";
import { VisibleView } from "./visibleView";
import { CELL_PADDING, DataEvents, DEFAULT_CELL_HEIGHT, DEFAULT_CELL_WIDTH, DEFAULT_SELECTED_BORDER_WIDTH, MouseLocation } from "../constant";
import { EventEmitter } from "../../utils/eventEmitter";
import { IndexDB } from "./indexDB";
import { UndoManager, UndoAction, FullStateSnapshot, CellPatchEntry, UndoAnchor, StatePatchSnapshot, StructuralSnapshot, ContinuousUndoSnapshot, SortPermutationSnapshot } from "../../utils/undoManager";
import { ClipboardManager } from "../../utils/clipboardManager";
import { MessageBox } from "../common/message";
import { MatrixSorter, type SortConfig } from "./matrixSorter";
import { PrinterSetting } from "../../utils/printer";
import { SliceBudget } from "../../utils/scheduler";
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
 * 稀疏扫描的候选格（内部用，不对外暴露）
 *
 * 与 {@link SearchHit} 的差别是多了 `ref`（`values` 里的**活引用**，替换阶段直接改它）
 * 与 `key`（行优先序号，仅用于排序）。分开定义是为了不让「活引用」泄进公开的命中项里：
 * 命中项会被对话框长期持有，带着一个可写的数据引用迟早出事。
 */
interface ScanCandidate {
    /** 单元格引用（`values` 里的活引用） */
    ref: Cell;
    /** 单元格名称 */
    cell: string;
    /** 列号（1-based） */
    col: number;
    /** 行号（1-based） */
    row: number;
    /** 逐字符对齐的文本（与 `Cell.getText()` 结果一致） */
    text: string;
    /** 行优先序号 = `(row - startRow) * width + (col - startCol)` */
    key: number;
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
     * 稀疏单元格索引：单元格名 → Cell。
     *
     * **本索引是「按名称取单元格」的唯一入口**（{@link getCellByName}）。
     * 历史实现用 `values.find(v => v.cell === name)` 线性查找，渲染路径上
     * 「每个单元格 × 每个字符」都要查一次 → O(N²·L)，是大表卡死/崩溃的主因。
     * 数组 `_values` 仅作为**有序视图**（迭代、序列化、长度）保留。
     *
     * 索引为惰性构建 + 失效重建：见 {@link _rebuildCellIndex} 与 {@link getCellByName}。
     */
    private _cellIndex: Map<string, Cell> = new Map();
    /** 构建索引时 `_values` 的长度：用于廉价侦测 push/splice 造成的结构变化 */
    private _indexedLength: number = -1;
    /** 结构世代：任何可能改变「名 → 单元格」映射的操作都会自增（见 {@link invalidateCellIndex}） */
    private _valuesEpoch: number = 0;
    /** 索引当前所处的世代；与 `_valuesEpoch` 不一致说明索引可能已过期 */
    private _indexEpoch: number = -1;
    /** 索引是否因结构操作（重命名/增删单元格）而需要重建；属性变化（着色/格式）不设此标记 */
    private _indexStructurallyInvalid: boolean = false;
    /**
     * 合并单元格锚点缓存（形如 `A1:C3`），随索引一并重建。
     * 原先 {@link getMergedCells} 每次调用都全表扫描一遍，而它又被
     * 「每个单元格」的矩形计算调用，构成第二处 O(N²)。
     */
    private _mergedCells: string[] = [];
    /**
     * 合并区域的**数值矩形**缓存（与 {@link _mergedCells} 同源、同时重建）。
     *
     * 字符串形式的 `A1:C3` 每次使用都要 `split(':')` + 逐字符解析行列号，
     * 而合并查找处在「每个可见单元格 × 每帧」的热路径上（主画布的背景/边框/文本
     * 三遍绘制各查一次），临时字符串与解析开销会被放大成几百万次。
     * 这里在重建时一次性把行列号算好，热路径只做整数比较。
     */
    private _mergeRects: MergeRect[] = [];
    /**
     * 「列号 → 覆盖该列的合并区域」倒排索引（1-based 列号）。
     *
     * 历史实现按「每个单元格 × 遍历全部合并区域」查找所属合并块，
     * 复杂度 O(可见单元格数 × 合并数)；合并单元格数以千计时，
     * 一帧就是上千万次比较 —— 实测 8000 处合并时 `drawAllMergedCellBackground`
     * 单帧 3.8 秒（渲染彻底卡死）。倒排后单次查找只与「该列上的合并块数」相关。
     */
    private _mergeByCol: Map<number, MergeRect[]> = new Map();
    /**
     * 跨列数超过 {@link MERGE_COL_INDEX_SPAN_LIMIT} 的合并区域：不铺进列索引，
     * 查找时单独线性兜底（这类「整行合并」通常只有个位数，代价可忽略）。
     */
    private _wideMergeRects: MergeRect[] = [];
    /**
     * 带筛选标记（`filter !== undefined`）的单元格缓存，随索引一并重建。
     * 供 SheetCanvas 每帧调用 {@link getFilterCells} 取用，
     * 避免原先每帧 `values.filter(...)` 的全表扫描。
     */
    private _filterCells: Cell[] = [];
    /**
     * 「带筛选标记的单元格」缓存对应的世代（与 {@link _valuesEpoch} 比对判失效）
     *
     * 供 {@link invalidateFilterCells} 在**已知有筛选按钮存在**时提前失效缓存，
     * 使 `SheetCanvas.setFilters()` 能 O(1) 判定「本帧要不要再问一次筛选单元格」
     * 而无需先重建整份单元格索引。
     */
    private _filterCellsEpoch: number = -1;
    /**
     * 筛选标记集合已被**显式置脏**，下次取用前必须重扫一次
     *
     * 增删筛选标记既不改 `values.length`，也未必让 `_filterCellsEpoch` 与
     * `_valuesEpoch` 拉开差距（挂/摘 `filter` 后往往紧跟一次 `markContentChanged()`，
     * 两者会同步推进），因此**单靠世代无法判定**「缓存里的按钮列表是否还准」。
     * 曾有实现只置 `_filterCellsEpoch = -1`，而 `_ensureCellIndex` 的
     * 「内容变了但结构没变」分支会顺手把它刷回 `_valuesEpoch`，
     * 于是返回的是**不含/仍含新按钮的旧缓存** —— 表现为
     * 「点筛选后按钮不显示 / 清筛选后按钮不消失」。故引入本标记，
     * 由 {@link invalidateFilterCells} 置位、{@link _rebuildFilterCells} 清除。
     */
    private _filterCellsDirty: boolean = true;
    /**
     * 「列号 → 该列上最后一个**有文本**的行号」缓存（1-based，含）
     *
     * 供筛选按钮的显隐判定（{@link hasContentBelow}）做 **O(1)** 查询。
     *
     * 历史实现是「每个筛选实例各自扫一遍整表」：`Filter.updateFilterVisibility`
     * 挂在每个实例的 VALUES_CHANGED 上，F 个按钮就是 **F × O(N)**。
     * 基准 `node test/_bench-filter-many.js` 实测 4 万格 × 40 个按钮 = **17.6 ms
     * / 次数据变化**（400 万格量级直接上百毫秒），这就是「筛选按钮一多就卡」的根因。
     *
     * **与单元格索引分开、惰性、按世代缓存**：只有筛选按钮会问这个结构，
     * 而 {@link _rebuildCellIndex} 的重建频率高得多，不该为筛选多付一次整表遍历。
     * 同一世代内只构建一次，F 个按钮共用同一份结果。
     */
    private _colContentBottom: Map<number, number> = new Map();
    /** {@link _colContentBottom} 所属的世代（与 {@link _valuesEpoch} 比对判失效） */
    private _colContentEpoch: number = -1;
    /** 构建 {@link _colContentBottom} 时 `_values` 的长度（长度变化即失效） */
    private _colContentLength: number = -1;
    /**
     * 「有文本单元格覆盖到的最大行号」（1-based；0 = 无内容）
     * 供筛选下拉清单用它替代 `matrixValues()` 求统计行范围。
     */
    private _contentMaxRow: number = 0;
    /** 「有文本单元格覆盖到的最大列号」（1-based；0 = 无内容） */
    private _contentMaxCol: number = 0;
    /**
     * 「有文本的行」集合（1-based）
     *
     * 筛选只可能隐藏**有内容**的行（全空行始终显示），因此判定筛选效果时
     * 只需在这批行上做，不必逐行扫到 `rowHeaders.length`（大文件下可达百万行）。
     */
    private _rowsWithText: Set<number> = new Set();
    /**
     * {@link _contentMaxRow} / {@link _contentMaxCol} / {@link _rowsWithText} 所属的世代
     *
     * 与 {@link _colContentBottom} **分开缓存**：后者挂在筛选按钮的显隐判定上、
     * 每次数据变化（含每次按键）都会被问到，必须保持「跳过无文本单元格、不解析名字」的
     * 廉价路径；而本索引只在**筛选交互**（打开下拉 / 执行筛选 / 清除筛选）时构建，
     * 因此可以顺带解析全部有文本单元格的名字并建集合。两者混在一起会让每次按键
     * 都白做一次整表名字解析。
     */
    private _filterIndexEpoch: number = -1;
    /** 构建筛选清单索引时 `_values` 的长度（长度变化即失效） */
    private _filterIndexLength: number = -1;
    /**
     * 「有内容的列」集合（1-based 列号，含合并块覆盖到的列）
     *
     * 供 {@link validataColInFilter} 做 **O(1)** 列有效性判定。
     *
     * 历史实现是「每列调一次 `matrixValues()`」——后者会分配一张
     * `maxRow × maxCol` 的**稠密二维数组**并遍历整表两次。选中 200 列时
     * 等于重建 200 次稠密矩阵：基准 `node test/_bench-filter-add.js`
     * 实测 6 万格（60 列 × 3000 行）下 200 列 = **1596 ms**，这就是
     * 「加载筛选按钮时主页面卡顿、取消（清除分支不走这条路径）却很流畅」的根因。
     *
     * 与 {@link _rowsWithText} 同一个索引、同一次遍历产出、同一处失效。
     */
    private _colsWithContent: Set<number> = new Set();
    /**
     * 上一次提交时的 this.values 深拷贝快照。
     * 作为下一次变更的「变化前」基准：每次原地修改后调用 commitValues()，
     * 便会以本快照为 before、当前 _values 为 after 记录一条撤销操作。
     *
     * **null 表示「无基准」**：大表（见 {@link LARGE_TABLE_UNDO_CELLS}）不做整表
     * 撤销记账，此时既不为基准做深拷贝、也不入栈。读到 null 的消费点必须
     * 「重建基准但不入栈」，否则会把空基准当真基准而错误地吞掉一次撤销。
     */
    private _lastValuesSnapshot: Cell[] | null = null;
    /** 是否正在回写撤销/重做快照（期间禁止再次记录，避免递归入栈） */
    private _isApplyingUndo: boolean = false;
    /** 是否禁止记录撤销（初始从数据库装载时为 true，避免把首次装载误记为一次变更） */
    private _suppressValuesUndo: boolean = false;
    /**
     * 「有未落库改动」标记
     *
     * 由 {@link _markDirty} 置位：凡改变了 {@link toJSON} 快照内容的状态变更事件
     * （单元格内容/样式、列宽行高、网格线、活动单元格与选区、缩放、打印设置）都会触发。
     * 落库时读的是**当下完整快照**，因此本标记只决定「要不要写」，不决定「写什么」。
     */
    private _valuesDirty: boolean = false;
    /**
     * 图片是否有未落库改动
     *
     * 与 {@link _valuesDirty} 分开记：图片是二进制大块，落库要整块写一遍，
     * 不能因为「改了某个单元格」就顺手把几十 MB 图片重写一次。
     * 只有图片集合真正变化（导入文件 / 切换文件 / 显式清空）时才置位。
     */
    private _imagesDirty: boolean = false;
    /**
     * 当前工作表内嵌图片（`ExtractedImage`，含二进制与锚点）
     *
     * 图片与单元格值分库存储（`images` 对象存储，键同为工作表名）：
     * 值走 JSON 字符串、图走结构化克隆，两者互不拖累。
     */
    private _images: ExtractedImage[] = [];
    /** 自动保存防抖定时器（非 null 表示已排期一次落库） */
    private _saveTimer: ReturnType<typeof setTimeout> | null = null;
    /** 空闲落库句柄（requestIdleCallback 的返回值，非 null 表示已排期下一次空闲执行） */
    private _idleHandle: number | null = null;
    /** 自动保存防抖窗口（毫秒）：连续输入/拖拽期间合并为一次落库 */
    private static readonly AUTO_SAVE_DEBOUNCE_MS = 500;
    /** 空闲落库的最迟等待（毫秒）：页面一直不空闲也必须在这个上限内把数据落库 */
    private static readonly AUTO_SAVE_IDLE_TIMEOUT_MS = 1000;
    /**
     * 「大表」阈值（单元格数）
     *
     * 超过它就不再为撤销/自动记录做**整表深拷贝**（见 {@link _undoTrackingEnabled}）。
     * 依据：一次 `cloneValues` 的成本与单元格数成正比，实测约 0.55 ms/千格
     * （40 万格 ≈ 220 ms、内存上还要多留一整份），而撤销栈每次操作要拷 2~3 份。
     * 在十万格量级上还感觉不到，到了几十万格，每一次结构操作都会变成肉眼可见的停顿。
     */
    private static readonly LARGE_TABLE_UNDO_CELLS = 100000;
    /**
     * 大表下「按影响面记账」的单次影响面上限（格子数）
     *
     * 大表关掉整表快照后，可撤销性改由**影响面补丁**提供（见 {@link runWithFullStateUndo}
     * 的 `scope` 参数）。补丁的代价与影响面格数成正比，因此仍需一道上限兜底：
     * 用户 Ctrl+A 全选后在百万格的表上「清除内容」，影响面就是整表，
     * 为它记账等于绕回整表快照（1.9M 格实测 1.7 秒 + 上百 MB）。
     *
     * 2 万格的实测成本约 10 ms / 几 MB，正好卡在「整行整列、一大片手选区域」
     * 都还能记账、而全选这种文档级操作会退化成不可撤销的位置。
     */
    private static readonly LARGE_TABLE_PATCH_CELLS = 20000;
    /**
     * 合并区域进「列倒排索引」的最大跨列数
     *
     * 倒排索引按「合并块覆盖的每一列」各存一份引用，跨列数极大的合并（如整行 A1:XFD1）
     * 会让索引条目数按列数膨胀。超过本阈值的合并块单独放进
     * {@link _wideMergeRects} 线性兜底 —— 这类区块在实践中只有个位数，
     * 线性扫描的代价远低于让索引膨胀上万条。
     *
     * 索引总条目数有硬上界：合并块两两不重叠，故
     * `Σ colspan ≤ Σ (colspan × rowspan) ≤ 单元格总数`，
     * 即倒排索引最多 N 条引用（40 万格量级约 3MB），不会随合并块数二次膨胀。
     */
    private static readonly MERGE_COL_INDEX_SPAN_LIMIT = 64;
    /**
     * 快照序列化 / 解码的片段粒度（单元格数）
     *
     * 4000 格在 40 万格的实测口径下约合 5 ms，既能让分片版及时让出主线程，
     * 又不至于把快照切成过多片段而拖慢拼接（`join`/`+=` 的次数）。
     */
    private static readonly SNAPSHOT_VALUE_CHUNK = 4000;
    /**
     * 全局筛选条件注册表：列号(1-based) → { anchorRow: 筛选按钮所在行(1-based), keep: 保留内容 }。
     * 注册表随全量状态快照捕获/回写，使筛选操作可整体撤销/重做；
     * 筛选组件（assistant/filter）通过 get/set/clear API 读写，不再自持静态状态。
     */
    private _filterConditions: Map<number, { anchorRow: number; keep: Set<string> }> = new Map();
    /** 筛选操作顺序：已筛选列号按首次执行筛选的先后排列（重放筛选条件时使用） */
    private _filterOrder: number[] = [];
    /**
     * 已隐藏行号倒排索引（1-based）
     *
     * `setRowsHidden` 同步维护：隐藏=true 加入集合，=false 移除。
     * 筛选（{@link computeFilterHiddenRows} / Filter.applyFilter）需要遍历
     * 「当前已隐藏的行」做差量更新；历史实现 `for r = anchorRow+1..totalRows`
     * 是 O(总行数) 全表扫描，百万行大表卡顿主因。改用本索引后为 O(已隐藏行数)，
     * 而已隐藏行通常只有几百到几千行。
     */
    private _hiddenRowsSet: Set<number> = new Set();
    /** runWithFullStateUndo 嵌套深度：仅最外层推入撤销记录，嵌套调用的变更折叠进外层快照 */
    private _fullStateUndoDepth: number = 0;
    /**
     * runWithCellPatchUndo 嵌套深度
     *
     * 包裹期间 {@link _recordValuesChange} 不再刷新整表基准（见该方法注释）：
     * 补丁自己管「变化前/后」，中间步骤刷新基准只会白白深拷贝一次整表。
     */
    private _cellPatchUndoDepth: number = 0;
    /**
     * 编辑期间的撤销栈：保存每次按键后的 values 深拷贝快照及对应光标绝对索引。
     * 栈底（index 0）为编辑前的基准快照，供编辑内逐字撤销；
     * 整段编辑在 endEditSession() 时折叠为 undoManager 的一条文档级撤销操作。
     */
    private _editUndoStack: { values: Cell[]; cursor: number }[] = [];
    /** 编辑期间的重做栈，配合 _editUndoStack 使用 */
    private _editRedoStack: { values: Cell[]; cursor: number }[] = [];
    /**
     * 最近一次编辑内快照入栈时间戳（performance.now()）
     *
     * 用于时间窗口合并：连续输入/删除在约 500ms 内的多次提交合并为栈顶一条，
     * 避免「输入 abcdef 要按 6 次 Ctrl+Z 才能撤销」的过细粒度——
     * 这是用户反馈「撤销/重做不精准」的根因。窗口内仅替换栈顶快照（保留最新 cursor），
     * 超过窗口或切换操作类型（输入↔删除）才新压一条。
     */
    private _lastEditCommitTime: number = 0;
    /**
     * 最近一次编辑提交的操作类型（'input' | 'delete' | 'style' | null）
     *
     * 时间窗口合并只在同类型内进行：连续输入合并、连续删除合并，
     * 但「输入→删除」之间不合并（让用户能分别撤销输入与删除）。
     */
    private _lastEditCommitType: 'input' | 'delete' | 'style' | null = null;
    /**
     * 大表编辑期间的「单格版」撤销栈
     *
     * 一次编辑会话只会改动活动单元格，因此大表（整表深拷贝被禁用，见
     * {@link LARGE_TABLE_UNDO_CELLS}）下改为仅对该单元格做逐状态深拷贝，
     * 成本与表大小无关。cell 为 null 表示该状态下活动单元格尚不存在。
     * 栈底为编辑前基准；_editUndoStack 为空即表示当前处于单格记账模式。
     */
    private _editCellUndoStack: { cell: Cell | null; cursor: number }[] = [];
    /** 大表编辑期间的单格重做栈 */
    private _editCellRedoStack: { cell: Cell | null; cursor: number }[] = [];
    /**
     * 本次编辑会话的视图锚点（在 {@link beginEditSession} 时捕获）
     *
     * 必须在编辑**开始**时捕获，不能在 {@link endEditSession} 时才采：
     * 以「点击另一个单元格结束编辑」为例，鼠标处理是先把 `activedCell` 改成新点的
     * 格子、随后才由 IS_EDITTING_CHANGED 触发 endEditSession —— 此刻采到的
     * `activedCell` 已经是新格子，而 `selection` 还停留在被编辑格的选区，二者错位。
     * 用这份错位锚点做撤销，就会出现「选中框跟随到被编辑格、活动格却留在新格子上」
     * （用户反馈：撤销/重做后按任意键编辑的不是被操作的单元格）。
     */
    private _editSessionAnchor: UndoAnchor | null = null;
    /**
     * 编辑内撤销/重做后应恢复到的光标绝对索引。
     * 由 _undoEditStep/_redoEditStep 写入，供 editCanvas 在重测 lines、
     * 定位光标时读取（光标定位必须发生在 lines 刷新之后）。
     */
    private _editRestoreCursor: number = 0;
    /**
     * 当前选中的工作表名称
     *
     * 同时也是**当前打开文件的存档键**（换算见 `utils/fileName` 的 `toArchiveKey`：
     * 只去掉原生 `.grid` 后缀，其它后缀保留完整路径，
     * 如 `归档/销售数据.grid` → `归档/销售数据`、`报表.text` → `报表.text`）。
     *
     * 默认值 `sheet1` 对应**默认空白文件**：文件树在「刷新后一个文件都没有」时会补一个
     * `sheet1.grid`，其存档键正是这里的 `sheet1`，因此初次装载的这张空表就是那个空白
     * 文件的内容，不会另外多出一条无归属的存档。
     */
    private _activedSheetName: string = 'sheet1';
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
    /**
     * 复制/剪切后留在源范围上的范围串（如 'A1:C5'）
     *
     * 仅用于界面层绘制「流动虚线框」（{@link CopyMarquee}），不参与数据/存档；
     * null 表示当前没有待粘贴的源范围。
     */
    private _copyRange: string | null = null;
    /**
     * {@link _copyRange} 所属的工作表名
     *
     * 虚线框只属于复制发生的那张表：切到其它表时 {@link copyRange} 对外返回 null
     * （界面隐藏），切回原表时又恢复显示；内部剪贴板不受影响，切表后仍可粘贴。
     */
    private _copyRangeSheet: string | null = null;
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
     * 从数据库获取当前选中工作表的所有有内容的单元格数据
     *
     * 解码走 {@link DataCollection.fromJSONChunked}（分片让出）：40 万格快照的
     * 解码实测约 1.4 秒整段同步，而这里正是**刷新页面**时的装载入口。
     */
    public async syncValuesFromDB(): Promise<void> {
        await this.db.getValue(this._activedSheetName).then(async (valueRes) => {
            // 装载数据属于初始化，不应被记为一次「变更」，期间禁止记录撤销
            this._suppressValuesUndo = true;
            try {
                if(valueRes) {
                    this._applyLoadedState(await DataCollection.fromSnapshotChunked(valueRes));
                }
                // 图片与值分库：装载顺序无关，但必须在 `_applyLoadedState` 之后
                // （值装载可能会清空图片集合，见 _clearSheetContent）
                await this._loadImages(this._activedSheetName);
            } finally {
                // 同步最新快照基准，解除抑制
                this._refreshBaseline();
                this._suppressValuesUndo = false;
                this._dataLoadedResolve();
            }
        });
    }

    /**
     * 把从数据库读出的快照装载进当前数据集合
     *
     * 供 {@link syncValuesFromDB} 与 {@link switchToSheet} 复用；不负责撤销抑制
     * 与快照基准，调用方需自行包裹（`_suppressValuesUndo` + `_lastValuesSnapshot`）。
     * @param data - {@link DataCollection.fromJSON} 的解析结果
     * @private
     */
    private _applyLoadedState(data: ReturnType<typeof DataCollection.fromJSON>): void {
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
        // 走公开 setter 而非直写 _zoom：缩放滑块、滚动条都依赖 ZOOM_CHANGED 同步，
        // 直接写私有字段会让「切换文件后」的缩放显示与实际渲染不一致
        this.zoom = data.zoom;
        // 打印设置统一入口随文档恢复（未保存过时保持接口默认值）
        this._printSetting = data.printSetting || {};
        this.values = data.values;
    }

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
     * 显式广播一次当前工作表内容（补发被抑制的 VALUES_CHANGED）
     *
     * `runWithFullStateUndo` 包裹期间 values 的自动广播被抑制（避免在撤销栈里产生空操作），
     * 但「打开文件」这类**会话中途的全量装载**必须让主页面重绘一次，否则
     * 画面上仍是上一个文件的内容。装载结束后调用本方法即可：
     * 先刷新 `_lastValuesSnapshot` 基准（使后续自动保存与撤销基准对齐），再广播。
     */
    public notifyValuesChanged(): void {
        this._refreshBaseline();
        this.emit(DataEvents.VALUES_CHANGED, this._values);
    }

    /**
     * 清空当前工作表并解除与文件的关联（主页面回到「没有打开任何文件」的空表）
     *
     * 用于「当前打开的文件被删除且无文件可回退」的场景。与切换一致，清空属于文档级操作：
     * 不记录撤销，并清空撤销/重做历史。
     *
     * 工作表名一并清空（{@link syncValuesToDB} 据此跳过落库）：名字就是存档键，
     * 若停留在刚被删除的文件名上，之后每帧重绘都会把这份空存档写回该键 ——
     * 删除等于没删，库里还会多出一条树中并不存在的记录。
     */
    public clearCurrentSheet(): void {
        // 走 setter 而非直写字段：通知外部（文件树据此清掉「定位 / 高亮」）
        this.activedSheetName = '';
        this._suppressValuesUndo = true;
        try {
            this._clearSheetContent();
        } finally {
            this._refreshBaseline();
            this._suppressValuesUndo = false;
        }
        // 装载期间广播被抑制，此处显式补发一次让主页面重绘为空表
        this.emit(DataEvents.VALUES_CHANGED, this._values);
        this.undoManager.clear();
    }

    /**
     * 标记「有未落库改动」并排入一次防抖自动保存
     *
     * 所有会改变 {@link toJSON} 快照内容的状态变更都应走到这里（见构造函数中订阅的事件
     * 与 {@link printSetting} 的 setter）。仅置脏 + 排期，不做任何序列化 ——
     * 真正的整表序列化（{@link toSnapshotChunked}）只发生在防抖窗口结束时的 {@link syncValuesToDB}。
     */
    private _markDirty(): void {
        this._valuesDirty = true;
        this._scheduleAutoSave();
    }

    /**
     * 排入一次防抖自动保存（已有排期则合并，不再重复排期）
     *
     * 未打开文件（无存档键）时不排期：那只是一张无归属的空表，
     * 写下去会造出「树里没有、库里却有」的幽灵存档（见 {@link clearCurrentSheet}）。
     */
    private _scheduleAutoSave(): void {
        if (!this._activedSheetName) return;
        if (this._saveTimer !== null || this._idleHandle !== null) return;
        this._saveTimer = setTimeout(() => {
            this._saveTimer = null;
            this._runIdleSave();
        }, DataCollection.AUTO_SAVE_DEBOUNCE_MS);
    }

    /**
     * 在浏览器空闲时段执行一次落库
     *
     * 整表序列化是**同步**的：大表上一次可能上百毫秒。若让它在防抖窗口
     * 结束时无条件执行，就可能正好落在「刚打开完文件、用户开始滚动」的当口造成卡顿。
     * 因此这里用 `requestIdleCallback` 把序列化推到空闲时段，并设 1s 上限保证最终一定执行；
     * 环境不支持（Node 测试 / 老浏览器）时立即执行，行为与原先一致。
     * @private
     */
    private _runIdleSave(): void {
        const body = (): void => {
            // 防抖窗口内若已被显式落库（_valuesDirty 被清）则不再重复写
            if (!this._valuesDirty) return;
            // 定时器里失败不应变成未处理的 Promise 拒绝；_writeSnapshot 已交还脏标记，
            // 下一次变更 / 页面卸载固化时会自动重试
            void this._writeSnapshot().catch((error) => {
                console.error('自动保存失败，将在下次变更或页面卸载时重试：', error);
            });
        };
        const ric = (globalThis as unknown as {
            requestIdleCallback?: (cb: () => void, opts?: { timeout: number }) => number;
        }).requestIdleCallback;
        if (typeof ric !== 'function') {
            body();
            return;
        }
        const handle = ric(() => {
            // 期间若已被 syncValuesToDB/_cancelPendingSave 取消，本次排期作废
            if (this._idleHandle !== handle) return;
            this._idleHandle = null;
            body();
        }, { timeout: DataCollection.AUTO_SAVE_IDLE_TIMEOUT_MS });
        this._idleHandle = handle;
    }

    /** 取消挂起的防抖保存 / 空闲落库（显式落库、强制固化前调用，避免随后又被重复写一遍） */
    private _cancelPendingSave(): void {
        if (this._saveTimer !== null) {
            clearTimeout(this._saveTimer);
            this._saveTimer = null;
        }
        if (this._idleHandle !== null) {
            const cic = (globalThis as unknown as {
                cancelIdleCallback?: (handle: number) => void;
            }).cancelIdleCallback;
            if (typeof cic === 'function') cic(this._idleHandle);
            this._idleHandle = null;
        }
    }

    /**
     * 写入当前整表快照并清除脏标记（落库的唯一出口）
     *
     * 先清脏标记再 `await` 写库：写库是异步的，期间若又有新变更会重新置脏并重新排期，
     * 不会因为「这次写库」而丢掉写库期间产生的改动。
     *
     * 序列化走 {@link toSnapshotChunked}（分片让出）而非同步 `toSnapshot()`：
     * 大表快照一次要 500 ms 左右，而它偏偏排在「打开文件」的防抖窗口之后，
     * 若不切片，用户会在文件已经显示出来之后再卡一下。
     */
    private async _writeSnapshot(): Promise<void> {
        const key = this._activedSheetName;
        if (!key) return;
        this._valuesDirty = false;
        // 图片只在**自己脏了**才写：它是二进制大块，值的小改动不该顺带重写一遍图片记录
        const imagesDirty = this._imagesDirty;
        this._imagesDirty = false;
        try {
            await this.db.setValue(key, await this.toSnapshotChunked());
            if (imagesDirty) await this.db.setImages(key, this._images);
        } catch (error) {
            // 写库失败（如配额不足）：交还脏标记，保证后续变更 / 卸载固化 / 下次显式落库还会重试，
            // 不至于把「本次未能持久化的改动」当成已保存而永久丢失
            this._valuesDirty = true;
            if (imagesDirty) this._imagesDirty = true;
            throw error;
        }
    }

    /**
     * 立即把当前工作表的**完整快照**落库（强制），并取消挂起的防抖保存
     *
     * - 未打开文件时直接返回：那只是一张无归属的空表，写下去会造出
     *   「树里没有、库里却有」的幽灵存档（见 {@link clearCurrentSheet}）；
     * - 与自动保存不同，这里**不跳过无改动的写**：它服务于「切换 / 重命名 / 复制 /
     *   打开」等需要「先把内存内容无条件固化再操作」的流程，作为防抖路径之外的安全网
     *   （这些流程本就低频，代价可忽略）；
     * - 原先 `Sheets.draw()` 每帧调用本方法（每帧整表序列化 + 写 IndexedDB），
     *   是打开大文件卡顿/崩溃的成因之一；现已从绘制线程移除，改由防抖自动保存驱动。
     */
    public async syncValuesToDB(): Promise<void> {
        if (!this._activedSheetName) return;
        this._cancelPendingSave();
        await this._writeSnapshot();
    }

    /**
     * 同步（不 await）固化未落库改动，用于页面卸载等无法等待异步完成的场景
     *
     * 仅在「确有未落库改动」时写：快照必须**同步**取到（这里刻意走 {@link toSnapshot}
     * 而非分片的 {@link toSnapshotChunked} —— 页面即将销毁，没有让出的机会；
     * 先把字符串拿在手里，再把写库请求交给 IndexedDB 异步完成，
     * 即便随后页面被销毁，数据也已被捕获）。
     */
    public flushPendingSave(): void {
        this._cancelPendingSave();
        if (!this._activedSheetName || (!this._valuesDirty && !this._imagesDirty)) return;
        const key = this._activedSheetName;
        const imagesDirty = this._imagesDirty;
        this._valuesDirty = false;
        this._imagesDirty = false;
        // 卸载路径无法等待/恢复：失败也只把脏标记交还，不抛出（避免未处理的 Promise 拒绝）
        void this.db.setValue(key, this.toSnapshot()).catch(() => { this._valuesDirty = true; });
        if (imagesDirty) {
            // 图片本身已在内存里是最终对象，IndexedDB 在收到请求瞬间就完成结构化克隆，
            // 因此调用完即可返回，不依赖后续事件循环
            void this.db.setImages(key, this._images).catch(() => { this._imagesDirty = true; });
        }
    }

    /**
     * 切换到指定工作表并装载其内容
     *
     * 调用顺序上先固化「当前工作表」的内容，再切换名称并装载目标表，
     * 因此不会丢改动。目标表尚无存档（如树中刚新建的文件）时按空表处理：
     * 清空单元格内容、保留行列头尺寸。
     *
     * 切换属于文档级操作：不记录撤销，并在切换后清空撤销/重做历史，
     * 避免「在 A 文件里撤销出 B 文件的内容」这类跨文件污染。
     * @param {string} sheetName - 目标工作表名（不含扩展名）
     * @returns {Promise<boolean>} 是否真正发生切换；当前已显示该表时返回 false（内容不变）
     */
    public async switchToSheet(sheetName: string): Promise<boolean> {
        const target = (sheetName ?? '').trim();
        // 已在主页面显示当前文件 → 内容无变化
        if (!target || target === this._activedSheetName) return false;

        // 先落库当前表：切换后 activedSheetName 即指向新表，届时再存就存错键了
        await this.syncValuesToDB();

        const saved = await this.db.getValue(target);
        // 先改名称：ACTIVED_SHEET_NAME_CHANGED 通知外部（文件树「定位」据此同步），
        // 后续自动保存也才会写到新表的键上
        this.activedSheetName = target;

        this._suppressValuesUndo = true;
        try {
            if (saved) {
                // 分片解码：双击树节点切换文件时，大表快照不再造成一至两秒的假死
                // （形态分派见 fromSnapshotChunked：v4 对象 / v3 字符串都能读）
                this._applyLoadedState(await DataCollection.fromSnapshotChunked(saved));
            } else {
                this._clearSheetContent();
            }
            // 图片跟着目标文件走：不清空就会把上一个文件的图片留在新文件上
            await this._loadImages(target);
        } finally {
            this._refreshBaseline();
            this._suppressValuesUndo = false;
        }

        // 装载期间 VALUES_CHANGED 被抑制（与初始化装载一致，避免误记撤销），
        // 但切换发生在会话中途，必须显式广播一次让主页面重绘为切换后的内容
        this.emit(DataEvents.VALUES_CHANGED, this._values);

        this.undoManager.clear();
        return true;
    }

    /**
     * 清空当前工作表内容（保留行列头尺寸），用于打开尚无存档的新文件
     * @private
     */
    private _clearSheetContent(): void {
        this._activedCell = 'A1';
        this._selection = 'A1:A1';
        this._anchorSelection = null;
        // 图片与单元格内容同属「文件内容」：清空内容必须连图片一起清，
        // 否则新文件会顶着上一个文件的图片（且下一次落库会把它们写到新键上）
        this._images = [];
        this._imagesDirty = false;
        // 滚动偏移随内容一起清空：旧文件可能滚动到很深的行，新（空）文件行数组被重置后，
        // 残留的 offsetHeight 会让 getVisibleRows() 探到越界行，行头画布据此 getAt(r)
        // 读出 undefined（Cannot read properties of undefined (reading 'isHidden')）
        this.rowHeaders.offsetHeight = 0;
        this.colHeaders.offsetWidth = 0;
        // values setter 会广播 VALUES_CHANGED → 主页面重绘为空表
        this.values = [];
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
        this._recordValuesChange(oldValue);
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
     * 从数据库装载当前工作表的图片（无记录时按「没有图片」处理）
     *
     * 读回的是外部数据，必须先做形状校验（见 {@link sanitizeImages}）：
     * 损坏的图片记录不允许把装载流程带崩。
     * @param {string} sheetName - 工作表名（存档键）
     * @returns {Promise<void>} 装载完成的 Promise
     * @private
     */
    private async _loadImages(sheetName: string): Promise<void> {
        let raw: unknown[] = [];
        try {
            raw = await this.db.getImages(sheetName);
        } catch {
            // 图片库不可用（老库未升级 / 配额问题）不应影响工作表装载：按无图片处理
            raw = [];
        }
        this._images = DataCollection.sanitizeImages(raw);
        this._imagesDirty = false;
    }

    /**
     * 校验并规范化从数据库读回的图片记录
     *
     * 只做「形状」校验（丢弃不合法项），不做任何解码：
     * - 二进制必须是 `Uint8Array`（结构化克隆保留类型）；
     * - 锚点必须存在且四个字段都是有限数。
     * 其余可选字段（`to`/`ext`/`name`）缺失时按缺省处理 —— 渲染侧本就有回落路径。
     * @param raw - 数据库读回的原始值
     * @returns {ExtractedImage[]} 可安全送入渲染层的图片数组
     */
    private static sanitizeImages(raw: unknown[]): ExtractedImage[] {
        if (!Array.isArray(raw)) return [];
        const result: ExtractedImage[] = [];
        for (const item of raw) {
            if (!item || typeof item !== 'object') continue;
            const candidate = item as Partial<ExtractedImage>;
            if (!(candidate.data instanceof Uint8Array) || candidate.data.length === 0) continue;
            if (typeof candidate.mimeType !== 'string') continue;
            const anchor = candidate.anchor;
            if (!anchor || !Number.isFinite(anchor.col) || !Number.isFinite(anchor.row)
                || !Number.isFinite(anchor.colOffset) || !Number.isFinite(anchor.rowOffset)) continue;
            result.push({
                data: candidate.data,
                mimeType: candidate.mimeType,
                anchor: {
                    col: anchor.col, row: anchor.row,
                    colOffset: anchor.colOffset, rowOffset: anchor.rowOffset,
                },
                to: candidate.to,
                ext: candidate.ext,
                name: candidate.name,
            });
        }
        return result;
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
    private _ensureCellIndex(): void {
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
     * 重扫 `values` 收集带筛选标记的单元格（刷新 {@link _filterCells} 缓存）
     *
     * 只对已存在的单元格做一次线性扫描，不建 Map、不建合并倒排，
     * 因此代价远低于 {@link _rebuildCellIndex}。由 {@link _ensureCellIndex}
     * 在 {@link _filterCellsDirty} 置位时调用。
     * @private
     */
    private _rebuildFilterCells(): void {
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
     * 单元格是否含非空字符（不拼接字符串、不产生中间数组）
     * @param {Cell} cell - 单元格
     * @returns {boolean} 是否存在非空字符
     * @private
     */
    private static hasAnyText(cell: Cell): boolean {
        const chars = cell.chars;
        if (!chars) return false;
        for (let i = 0; i < chars.length; i++) {
            const ch = chars[i].char;
            if (ch != null && ch !== '') return true;
        }
        return false;
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
    private _ensureFilterIndex(): void {
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
            if (!DataCollection.hasAnyText(cell)) continue;
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
    private _refreshBaseline(): void {
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
            this.undoManager.push({
                type: 'values', data: { before: this.cloneValues(before), after: this.cloneValues(after) },
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
        const name = this.activedCell;
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
    public beginEditSession(): void {
        // 在编辑开始时捕获视图锚点：此刻 activedCell 与 selection 都指向被编辑的
        // 那一个单元格，二者一致；推迟到结束时再采会因鼠标已改动 activedCell 而错位
        // （见 _editSessionAnchor 字段注释）。
        this._editSessionAnchor = this._captureAnchor();
        if (this._undoTrackingEnabled()) {
            this._editUndoStack = [{ values: this.cloneValues(this._values), cursor: this.cursorAbsolutePosition }];
        } else {
            // 大表：整表深拷贝关闭，改对活动单元格建立单格基准快照，
            // 编辑内撤销/重做依然可用（见 _editCellUndoStack）。
            const active = this.getCellByName(this.activedCell);
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
            const name = first.cell ? first.cell.cell : this._activedCell;
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
        // 撤销 F10 的输入，锚点会把活动格留在 A1）；选区仍按锚点恢复，
        // 活动格单独落到操作发生处
        const targetCell = this._resolveUndoTargetCell(action, isUndo);
        if (targetCell) {
            this.activedCell = targetCell;
            const { col, row } = this.getCellColAndRow(targetCell);
            if (isFinite(col) && isFinite(row) && col >= 1 && row >= 1) {
                this.scrollCellIntoView(col, row);
            }
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
     * - values 及未知类型：返回 null，保持锚点恢复的活动格。
     * @param {UndoAction} action - 操作对象
     * @param {boolean} isUndo - true=撤销，false=重做
     * @returns {string | null} 目标单元格名，无法推导时为 null
     */
    private _resolveUndoTargetCell(action: UndoAction, isUndo: boolean): string | null {
        const side = isUndo ? 'before' : 'after';
        const data = action.data as Record<string, any>;
        switch (action.type) {
            case 'cellPatch': {
                const before = data?.before as CellPatchEntry[] | undefined;
                const after = data?.after as CellPatchEntry[] | undefined;
                return before?.[0]?.cell ?? after?.[0]?.cell ?? null;
            }
            case 'statePatch': {
                const before = data?.before as StatePatchSnapshot | undefined;
                const after = data?.after as StatePatchSnapshot | undefined;
                return before?.cells?.[0]?.cell ?? after?.cells?.[0]?.cell ?? null;
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
     * 生成快照的 JSON 文本片段（**同步与分片两条路径共用的唯一实现**）
     *
     * 逐段 yields 而非一次 `JSON.stringify(整个对象)`，是为了让
     * {@link toSnapshotChunked} 能在片段之间让出主线程：几十万格的快照序列化
     * 原本是一段**不可中断**的同步任务（40 万格实测约 500 ms，且正好落在
     * 「刚打开完文件」的防抖落库时刻），是打开大文件后仍会卡顿的最后一环。
     *
     * 片段粒度按 {@link DataCollection.SNAPSHOT_VALUE_CHUNK} 划分（4000 格 ≈ 5 ms），
     * 保证分片版既能及时让出，又不会因为片段过多而拖慢拼接。
     *
     * ## 输出格式 v3：样式表
     *
     * 单元格不再逐格内联样式字段，而是写样式表编号 `s`；表项集中在尾部
     * `styles` 数组。样式在真实表格里重复度极高，而装载路径的 `JSON.parse`
     * 是整段同步的 —— 快照小一倍，卡顿就少一半（详见 `CellStyleTable`）。
     * 槽位稳定：即使表为空也输出 `"styles":[]`，读回侧无需特判。
     *
     * {@link toJSON} 只是把这些片段直接拼起来（同步、与历史语义一致）。
     * @returns {Generator<string>} JSON 文本片段
     * @private
     */
    private *_snapshotParts(): Generator<string> {
        yield '{"v":3';
        yield ',"activedCell":' + JSON.stringify(this.activedCell);
        yield ',"selection":' + JSON.stringify(this._selection);
        yield ',"printSetting":' + JSON.stringify(this._printSetting);
        yield ',"colHeaders":' + JSON.stringify(this.colHeaders.colHeaders);
        yield ',"offsetWidth":' + JSON.stringify(this.colHeaders.offsetWidth);
        yield ',"allColWidth":' + JSON.stringify(this.colHeaders.allColWidth);
        yield ',"rowHeaders":' + JSON.stringify(this.rowHeaders.rowHeaders);
        yield ',"offsetHeight":' + JSON.stringify(this.rowHeaders.offsetHeight);
        yield ',"allRowHeight":' + JSON.stringify(this.rowHeaders.allRowHeight);
        yield ',"showGridLines":' + JSON.stringify(this.showGridLines);
        yield ',"zoom":' + JSON.stringify(this.zoom);
        yield ',"values":[';

        // 样式表：随 values 一起逐格填充，写入快照尾部（读回时已整份 JSON.parse，
        // 键序不影响语义，故可以放在 values 之后，省掉一次预扫描）
        const styles = new CellStyleTable();
        let first = true;
        for (const part of this._valueChunkParts(styles)) {
            yield (first ? '' : ',') + part;
            first = false;
        }
        yield ']';
        yield ',"styles":' + JSON.stringify(styles.entries);
        yield '}';
    }

    /**
     * 逐片生成 `values` 的单元格 JSON 片段（**单元格序列化的唯一实现**）
     *
     * 每片产出「不含外层方括号、不含分隔逗号」的片段，形如
     * `{"c":"A1","t":"x"},{"c":"A2","s":3}`。两条落库路径共用它：
     * - v3 文本存档（{@link _snapshotParts}）把片段用逗号拼进 `"values":[…]`；
     * - v4 对象存档（{@link toSnapshotChunked}）把每片各自包成 `[…]`，
     *   于是「每片都是一份可独立 `JSON.parse` 的数组」。
     *
     * 之所以做成生成器：分片版可以在片与片之间 `await` 让出主线程，
     * 而同步版（页面卸载）直接一次跑完。
     * @param {CellStyleTable} styles - 样式表（逐格填充，由调用方负责写出）
     * @returns {Generator<string>} 每片一段 JSON 文本
     * @private
     */
    private *_valueChunkParts(styles: CellStyleTable): Generator<string> {
        const values = this._values;
        for (let i = 0; i < values.length; i += DataCollection.SNAPSHOT_VALUE_CHUNK) {
            const end = Math.min(i + DataCollection.SNAPSHOT_VALUE_CHUNK, values.length);
            const parts: string[] = new Array(end - i);
            for (let k = i; k < end; k++) {
                parts[k - i] = JSON.stringify(values[k].toCompact(styles));
            }
            yield parts.join(',');
        }
    }

    /**
     * 生成 v4 存档（**同步**，页面卸载等无法等待的场景用）
     *
     * 与 {@link toSnapshotChunked} 输出完全相同（共用 {@link _valueChunkParts}）。
     * 后台落库走分片版；本方法只在 `pagehide`/`visibilitychange` 这类
     * 「必须立刻把快照拿到手」的时刻调用。
     * @returns {SheetSnapshotV4} 存档对象（可直接交给 {@link IndexDB.setValue}）
     */
    public toSnapshot(): SheetSnapshotV4 {
        const styles = new CellStyleTable();
        const chunks: string[] = [];
        for (const part of this._valueChunkParts(styles)) chunks.push('[' + part + ']');
        return { ...this._snapshotEnvelope(), chunks, styles: styles.entries };
    }

    /**
     * 生成 v4 存档（**分片让出**，常规落库用）
     *
     * 与 {@link toSnapshot} 输出完全相同，区别是每片之间 `await` 让出一次：
     * 40 万格的单元格序列化原本是约 310 ms 的整段同步任务。
     * @returns {Promise<SheetSnapshotV4>} 存档对象
     */
    public async toSnapshotChunked(): Promise<SheetSnapshotV4> {
        const styles = new CellStyleTable();
        const chunks: string[] = [];
        const slice = new SliceBudget();
        for (const part of this._valueChunkParts(styles)) {
            chunks.push('[' + part + ']');
            if (slice.expired()) await slice.yieldNow();
        }
        return { ...this._snapshotEnvelope(), chunks, styles: styles.entries };
    }

    /**
     * 生成 v4 存档的「元信息 + 稀疏行列头」部分
     * @returns {Omit<SheetSnapshotV4, 'chunks' | 'styles'>} 不含单元格数据的部分
     * @private
     */
    private _snapshotEnvelope(): Omit<SheetSnapshotV4, 'chunks' | 'styles'> {
        return {
            v: 4,
            activedCell: this.activedCell,
            selection: this._selection,
            printSetting: this._printSetting,
            showGridLines: this.showGridLines,
            zoom: this.zoom,
            offsetWidth: this.colHeaders.offsetWidth,
            offsetHeight: this.rowHeaders.offsetHeight,
            cols: encodeColAxis(this.colHeaders.colHeaders),
            rows: encodeRowAxis(this.rowHeaders.rowHeaders),
        };
    }

    /**
     * 将数据集合转换为JSON表示（**紧凑格式 v3**）
     *
     * ## 为什么不再 `values.map(v => v.toJSON())`（性能约定，请勿回退）
     * 历史实现是「每个 Cell 先 `JSON.stringify` 成字符串，再把这些字符串放进
     * 数组整体 `JSON.stringify` 一次」。这是**双层编码**：第二层会把第一层里
     * 每一个引号都转义一次，而 `Cell` 内部的 `chars` 又是「每字符一串 JSON」的
     * 字符串 —— 于是一个 6 字符的单元格能膨胀到约 690 字节。
     * 实测 40 万格 ⇒ 134 MB 临时字符串、约 1.9 秒纯主线程同步耗时，
     * 且紧随打开文件之后（打开即回写）立刻发生，是「大文件打开卡死」的主因之一。
     *
     * 现在改为：`Cell.toCompact()` 产出**普通对象**（省略默认值 + 统一样式的字符
     * 合并为「文本 + 共享字体」），交给序列化一次性写出。
     * 同样 6 字符的单元格只剩约 40~110 字节，单格体积降到约 1/6~1/17。
     *
     * 样式字段进一步抽成 **v3 样式表**（`CellStyleTable`）：单元格只写编号
     * `s`，表项集中在尾部的 `styles` 数组。样式在真实表格里重复度极高
     * （400 行加粗背景同一种样式），而装载路径的 `JSON.parse` 是整段同步的 ——
     * 实测 20 万格带样式：快照 30.0 MB → 7.4 MB，`JSON.parse` 388 ms → 138 ms。
     *
     * 版本标记由 {@link DataCollection.fromJSON} 分派解码路径：
     * `v:3` 读样式表、`v:2` 读内联样式、无 `v`（历史存档）走遗留双层解码，
     * 三种存档**永远都能读回**。
     *
     * 本方法保持**同步**（既有调用方与测试依赖同步语义）。后台落库走
     * {@link toSnapshotText}（分片让出）；页面卸载走本方法（必须同步取到快照）。
     * **注意：v4 起常规落库不再走本方法**（见 {@link toSnapshot}），
     * 它保留为「v3 文本格式」的写入端，供测试与旧格式兼容。
     * @returns {string} 数据集合的JSON表示
     */
    public toJSON(): string {
        let text = '';
        for (const part of this._snapshotParts()) text += part;
        return text;
    }

    /**
     * 分片生成同一份快照文本，期间让出主线程
     *
     * 与 {@link toJSON} 输出**完全相同**（共用 {@link _snapshotParts}），
     * 区别是每片之间 `await` 让出一次，浏览器得以处理输入与渲染。
     * 自动保存与显式落库都走这里：40 万格快照原本是约 500 ms 的整段同步任务，
     * 且恰好紧跟在「打开文件」的防抖窗口之后，是打开大文件后仍会卡一下的元凶。
     * @returns {Promise<string>} 数据集合的JSON表示
     */
    public async toSnapshotText(): Promise<string> {
        const parts: string[] = [];
        const slice = new SliceBudget();
        for (const part of this._snapshotParts()) {
            parts.push(part);
            if (slice.expired()) await slice.yieldNow();
        }
        return parts.join('');
    }

    /**
     * 空白工作表存档（内容为空，视图设置沿用当前会话）
     *
     * 供「新建文件」时写入数据库：让「文件树里有这个文件」与「`values` 库里
     * 有这个文件的存档」始终一一对应（见 app.ts 的内容桥）。
     *
     * 之所以沿用当前视图设置（列宽、行高、缩放、打印设置）而不是全用零值：
     * 打开一个尚无内容的文件时 {@link switchToSheet} 走的是 {@link _clearSheetContent}，
     * 本来就保留当前行列头尺寸；若这里写入空的行列头，重新打开该文件会得到
     * `allColWidth = 0`，滚动条与画布都会异常。
     *
     * 形态是 {@link SheetSnapshotV4}（与常规落库一致）：`chunks` 为空数组
     * 即代表「没有单元格」，{@link isV4Snapshot} + app.ts 的 `isEmptySnapshot`
     * 都据此判空。
     * @returns {SheetSnapshotV4} 空白工作表存档（可直接交给 {@link IndexDB.setValue}）
     */
    public blankSnapshot(): SheetSnapshotV4 {
        const snapshot = this.toSnapshot();
        snapshot.activedCell = 'A1';
        snapshot.selection = 'A1:A1';
        snapshot.chunks = [];
        // 内容清空后样式表随之作废（留着只会让每个空白存档多背一份无用数据）
        snapshot.styles = [];
        return snapshot;
    }

    /**
     * 判断一个存档值是否为 v4 形态
     *
     * `chunks` 必须是数组：缺了它就没有单元格数据可言，
     * 按「不是 v4」处理会走回字符串分支（`typeof raw === 'string'` 不成立）
     * 而落到空表兜底 —— 这正是我们想要的保守行为。
     * @param raw - 存档值（来自 IndexedDB，形态未知）
     * @returns {boolean} 是否为 v4 存档
     */
    public static isV4Snapshot(raw: unknown): raw is SheetSnapshotV4 {
        return !!raw && typeof raw === 'object'
            && (raw as { v?: unknown }).v === 4
            && Array.isArray((raw as { chunks?: unknown }).chunks);
    }

    /**
     * 从**任意形态**的存档值装载（v4 对象 / v3 字符串 / v2 与遗留字符串）
     *
     * 分派规则：
     * - `string` → {@link fromJSONChunked}（v3 样式表 / v2 内联样式 / 无版本标记的遗留双层编码，
     *   三种历史存档**永远都能读回**）；
     * - v4 对象 → 逐片 `JSON.parse` + 解码，片间让出主线程；
     * - 其它（空对象、损坏值）→ 按空表处理，不让损坏存档带崩装载流程。
     *
     * 这是装载路径的**唯一入口**（{@link syncValuesFromDB} 与 {@link switchToSheet} 都走它）。
     * @param raw - 存档值
     * @returns {Promise} 解码后的各部分（与 {@link fromJSON} 结构一致）
     */
    public static async fromSnapshotChunked(raw: unknown): Promise<ReturnType<typeof DataCollection.fromJSON>> {
        if (typeof raw === 'string') return DataCollection.fromJSONChunked(raw);
        if (!DataCollection.isV4Snapshot(raw)) {
            return DataCollection._assembleSnapshot({}, []);
        }

        const chunks = raw.chunks;
        const styles = Array.isArray(raw.styles) ? raw.styles : [];
        const values: Cell[] = [];
        const slice = new SliceBudget();
        for (const chunk of chunks) {
            if (typeof chunk !== 'string' || chunk.length === 0) continue;
            let parsed: unknown;
            try {
                parsed = JSON.parse(chunk);
            } catch {
                // 单片损坏不该拖垮整份存档：跳过该片，其余内容照常装载
                continue;
            }
            if (!Array.isArray(parsed) || parsed.length === 0) continue;
            // 每片是各自解析出来的独立数组：源下标从 0 起，用 destOffset 映射回整表绝对位置
            const base = values.length;
            values.length = base + parsed.length;
            DataCollection._decodeValueRange(parsed, values, 0, parsed.length, styles, base);
            if (slice.expired()) await slice.yieldNow();
        }

        return DataCollection._assembleSnapshot({
            activedCell: raw.activedCell,
            selection: raw.selection,
            printSetting: raw.printSetting,
            colHeaders: decodeColAxis(raw.cols),
            offsetWidth: raw.offsetWidth,
            rowHeaders: decodeRowAxis(raw.rows),
            offsetHeight: raw.offsetHeight,
            showGridLines: raw.showGridLines,
            zoom: raw.zoom,
        }, values);
    }

    /**
     * 从JSON表示创建数据集合
     *
     * 保持**同步**（历史 API，测试与若干调用方依赖同步语义）。
     * 从数据库装载大表请用 {@link fromJSONChunked}：它会分片让出主线程，
     * 避免刷新页面 / 切换文件时卡住一至两秒。
     * @param {string} json - 数据集合的JSON表示
     * @returns 解码后的各部分（values 为 `Cell` 实例数组）
     */
    public static fromJSON(json: string): {
        activedCell: string,
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
        const rawValues: any[] = parsed.values || [];
        const values: Cell[] = new Array(rawValues.length);
        DataCollection._decodeValueRange(rawValues, values, 0, rawValues.length, parsed.styles);
        return DataCollection._assembleSnapshot(parsed, values);
    }

    /**
     * 分片解码快照文本（装载大表用）
     *
     * 与 {@link fromJSON} 解码结果**完全相同**（共用 {@link _decodeValueRange}），
     * 区别是每解码一段就让出一次主线程。40 万格的快照解码实测约 1.4 秒整段同步，
     * 而这段代码正跑在「刷新页面」与「双击树节点切换文件」的路径上 ——
     * 分片后期间窗口可响应，不再出现切文件即假死。
     *
     * 注意：`JSON.parse` 本身仍是**不可分片**的一次调用（大快照约数百毫秒），
     * 这是继续压缩装载耗时的下一步方向（见项目记忆「待办」）。
     * @param {string} json - 数据集合的JSON表示
     * @returns {Promise} 解码后的各部分
     */
    public static async fromJSONChunked(json: string): Promise<{
        activedCell: string,
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
        values: Cell[]}> {
        const parsed = JSON.parse(json);
        const rawValues: any[] = parsed.values || [];
        const values: Cell[] = new Array(rawValues.length);
        const slice = new SliceBudget();
        for (let i = 0; i < rawValues.length; i += DataCollection.SNAPSHOT_VALUE_CHUNK) {
            const end = Math.min(i + DataCollection.SNAPSHOT_VALUE_CHUNK, rawValues.length);
            DataCollection._decodeValueRange(rawValues, values, i, end, parsed.styles);
            if (slice.expired()) await slice.yieldNow();
        }
        return DataCollection._assembleSnapshot(parsed, values);
    }

    /**
     * 解码 `values` 的 `[from, to)` 区间（**同步与分片两条路径共用的唯一实现**）
     *
     * 逐格分派：v3/v2 紧凑对象走 `Cell.fromCompact`，遗留格式（`values[i]` 是字符串）
     * 走双层解码。刻意做成「按区间」而非「整表」，分片路径才能把同一份逻辑
     * 拆成多段执行，不会出现两份实现行为漂移。
     * **源区间与目标位置必须分开**：v3 文本存档里 `rawValues` 就是整张表
     * （源下标 = 目标下标），而 v4 存档的**每一片是各自 `JSON.parse` 出来的
     * 独立数组**（源下标从 0 起），要靠 `destOffset` 映射回整表的绝对位置。
     * 历史上这里只按「源下标 = 目标下标」写，于是 v4 装载时第 2 片起全部
     * 读到越界下标 → 解出空单元格（多片存档丢内容）。见 `test-snapshot-v4.js`。
     * @param {any[]} rawValues - 源数组（整表，或 v4 的**单片**）
     * @param {Cell[]} values - 输出数组（按绝对下标填充）
     * @param {number} from - 源起始下标（含）
     * @param {number} to - 源结束下标（不含）
     * @param {any[]} [styles] - v3 样式表（`values[i].s` 是它的下标）；v2 / 遗留存档为 undefined
     * @param {number} [destOffset=from] - 目标起始下标；默认与 `from` 相同（v3 场景源/目标同下标）
     * @private
     */
    private static _decodeValueRange(
        rawValues: any[],
        values: Cell[],
        from: number,
        to: number,
        styles?: any[],
        destOffset: number = from
    ): void {
        for (let i = from; i < to; i++) {
            const raw = rawValues[i];
            values[destOffset + (i - from)] = (typeof raw === 'string')
                ? DataCollection.decodeLegacyCell(raw)
                : Cell.fromCompact(raw, styles);
        }
    }

    /**
     * 把解码好的 values 与快照其余字段组装成统一结构
     * @param {any} parsed - `JSON.parse` 的原始结果
     * @param {Cell[]} values - 已解码的单元格数组
     * @returns 组装结果
     * @private
     */
    private static _assembleSnapshot(parsed: any, values: Cell[]): {
        activedCell: string,
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
            values
        }
    }

    /**
     * 解码遗留格式（双层编码）的单个单元格
     *
     * 历史存档里 `values[i]` 是「Cell 的 JSON 字符串」，其 `chars` 字段又是
     * 「每字符一串 JSON」的字符串数组 —— 必须解析两层才能拿到字符。
     * 字符还原走 `Char.fromRaw`（稀疏复制）：逐字段无条件赋值会把 7 个样式
     * 字段全部物化成自有属性，大表装载时白白抬高内存。
     * @param {string} itemJSON - 单元格的 JSON 字符串
     * @returns {Cell} 单元格实例
     * @private
     */
    private static decodeLegacyCell(itemJSON: string): Cell {
        const item = JSON.parse(itemJSON);
        const cell = new Cell();
        cell.cell = item.cell || '';
        cell.isEdit = item.isEdit || false;
        cell.colspan = item.colspan;
        cell.rowspan = item.rowspan;
        cell.borderTopWidth = item.borderTopWidth;
        cell.borderBottomWidth = item.borderBottomWidth;
        cell.borderLeftWidth = item.borderLeftWidth;
        cell.borderRightWidth = item.borderRightWidth;
        cell.textAlign = item.textAlign;
        cell.alignItems = item.alignItems;
        cell.backgroundColor = item.backgroundColor;
        cell.numberFormat = item.numberFormat;
        cell.borderColor = item.borderColor;
        cell.filter = item.filter;
        cell.wrap = item.wrap;
        cell.letterSpacing = item.letterSpacing;
        cell.lineSpacing = item.lineSpacing;
        const chars = JSON.parse(item.chars || '[]');
        cell.chars = chars.map((charItem: any) => Char.fromRaw(
            typeof charItem === 'string' ? JSON.parse(charItem) : charItem
        ));
        return cell;
    }
    
    /* ==================== 查询与替换 ====================
     *
     * 「工具栏 → 查询 / 替换」的数据层实现。**全部限定在给定矩形范围内**，
     * UI 层（toolbar/start/searchDialog.ts）只负责取值与展示：
     * - 范围是显式参数，不读也不改当前选区（`gotoSearchHit` 除外），
     *   因此「跳转命中项（会改写选区）」之后继续「查找下一个」不会把范围缩小成一个格子；
     * - **扫描只遍历 `values`（真实存在的单元格）**，不遍历 行 × 列 网格 ——
     *   见 `_collectSearchCells`。整表范围是常态（单格选区会被解释成整表，
     *   见 `expandSingleCellRect`），而 10 万行 × 100 列的表按网格遍历是**亿级**空转；
     * - 替换的撤销补丁按**命中的格子**逐个捕获（`runWithCellsUndo`），不是按矩形逐格捕获：
     *   命中天然稀疏，矩形版在整表范围下会为每一格都克隆一份快照；
     * - 行/列的隐藏状态按 copySelection 的口径：**隐藏行列不参与查找**，
     *   否则筛选状态下会命中用户看不见的格子。
     * ==================== */

    /**
     * 解析一段区域串（`'B3'` 或 `'A1:C9'`，大小写不敏感）
     *
     * 只认这两种形态；任何一处不合法返回 null，由调用方决定如何提示
     * （不要让本方法抛错：范围串直接来自对话框输入框，用户可能手打）。
     * @param {string} range - 区域串
     * @returns {SheetRect | null} 规范化后的矩形（1-based 闭区间，端点已按大小端排序）
     */
    private _parseRect(range: string): SheetRect | null {
        const parts = range.split(':');
        if (parts.length > 2) return null;
        const a = this._parseCellRef(parts[0]);
        if (!a) return null;
        if (parts.length === 1) {
            return { startCol: a.col, startRow: a.row, endCol: a.col, endRow: a.row };
        }
        const b = this._parseCellRef(parts[1]);
        if (!b) return null;
        return {
            startCol: Math.min(a.col, b.col),
            startRow: Math.min(a.row, b.row),
            endCol: Math.max(a.col, b.col),
            endRow: Math.max(a.row, b.row)
        };
    }

    /**
     * 解析单个单元格引用（`'AB12'` → `{ col: 28, row: 12 }`）
     *
     * 手写这里而不是复用 {@link getCellColAndRow}：后者对 `'A0'`、`''`、`'1A'`
     * 这类输入会返回 `row: 0` / `col: 0` / `NaN` 而不报错，作为范围来源会让
     * 「查询第 0 行」这种不存在的结果悄悄进入扫描循环。
     * @param {string} ref - 单元格引用
     * @returns {{ col: number, row: number } | null} 行列号（1-based）；非法返回 null
     */
    private _parseCellRef(ref: string): { col: number; row: number } | null {
        const trimmed = ref.trim();
        const m = /^([A-Za-z]{1,3})([1-9][0-9]{0,6})$/.exec(trimmed);
        if (!m) return null;
        const col = this.getColNumber(m[1].toUpperCase());
        const row = parseInt(m[2], 10);
        if (!isFinite(col) || col < 1 || !isFinite(row) || row < 1) return null;
        return { col, row };
    }

    /**
     * 决定本次查询 / 替换的作用范围
     *
     * 三级回退：
     * ① 显式区域串（对话框「范围」输入框的值优先，让用户在 dialog 里点选或手填的范围说了算）；
     * ② 当前选区（工具栏入口的默认语义 = **只在当前选中区域内查找**）；
     * ③ 无选区时回退到**有内容的区域**（{@link getContentBounds}），而不是行列表头的
     *    逻辑上限（1048575 × 16384）——后者会让一次查询扫过整张虚拟表。
     * @param {string} [range] - 显式区域串
     * @returns {SheetRect | null} 查询范围；连空表都没有时返回 null
     */
    public resolveSearchRect(range?: string): SheetRect | null {
        if (range && range.trim()) {
            const rect = this._parseRect(range);
            if (rect) return rect;
        }
        if (this.selection) {
            const rect = this._parseRect(this.selection);
            if (rect) return rect;
        }
        const { maxRow, maxCol } = this.getContentBounds();
        if (!maxRow || !maxCol) return null;
        return { startCol: 1, startRow: 1, endCol: maxCol, endRow: maxRow };
    }

    /**
     * 把范围矩形格式化成 `'A1:C9'`（单格时为 `'A1:A1'`，与 selection 的既有口径一致）
     * @param {SheetRect} rect - 范围矩形
     * @returns {string} 区域串
     */
    public formatRect(rect: SheetRect): string {
        const start = this.getCellName(rect.startCol, rect.startRow);
        const end = this.getCellName(rect.endCol, rect.endRow);
        return `${start}:${end}`;
    }

    /**
     * 把「单格范围」解释为「整表范围 + 起点锚」
     *
     * 选中单个单元格之后查找，要找的显然不是这一格本身（它的内容就摆在眼前），
     * 而是**整张表**里还有哪些地方出现同样的内容；「查找下一个」从这一格往下走、
     * 到头回绕 —— 与 Excel「选中一格 → Ctrl+F → 查找全部/下一个」的语义一致。
     *
     * 返回 `null` 表示**保持原样**：范围不是单格，或者表里根本没有内容
     * （`getContentBounds()` 为 0，此时连一个候选格都没有，扩不扩张结果都一样）。
     *
     * 注意这是 `DataCollection` 的语义、不是「范围串恰好写成单格」的判定：
     * 由主表格选出的单格走本方法（范围会被**显示**成展开后的整表范围，用户看得见
     * 将要搜多大一片）；用户手工在范围框里敲 `'B2:B2'` 时按字面处理，就是只查那一格。
     * @param {SheetRect} rect - 待检查的范围（1-based 闭区间）
     * @returns {{ rect: SheetRect; anchor: { col: number; row: number } } | null} 整表范围 + 起点锚；不适用时 null
     */
    public expandSingleCellRect(rect: SheetRect): { rect: SheetRect; anchor: { col: number; row: number } } | null {
        if (rect.startCol !== rect.endCol || rect.startRow !== rect.endRow) return null;
        const { maxRow, maxCol } = this.getContentBounds();
        if (!maxRow || !maxCol) return null;
        return {
            rect: { startCol: 1, startRow: 1, endCol: maxCol, endRow: maxRow },
            anchor: { col: rect.startCol, row: rect.startRow }
        };
    }

    /**
     * 构造查询用正则
     *
     * - `useWildcard`：`?` → 单个字符、`*` → 任意长度（含换行，故用 `[\s\S]` 而非 `.`），
     *   其余字符按正则元字符转义；
     * - `matchEntire`：首尾锚定（「单元格内容完全匹配」而非包含即可）；
     * - `matchCase`：不加 `i` 标志。
     * 统一返回**带 g 标志**的正则：命中判定与「把所有匹配位置替换掉」共用同一个 「lastIndex 游标」语义。
     * @param {string} keyword - 关键字
     * @param {SearchMatchOptions} [options] - 匹配选项
     * @returns {RegExp | null} 正则；关键字为空返回 null（空关键字不构成查询）
     */
    private _buildSearchPattern(keyword: string, options?: SearchMatchOptions): RegExp | null {
        if (!keyword) return null;
        const escape = DataCollection._escapeRegExp;
        let source = '';
        if (options?.useWildcard === true) {
            for (let i = 0; i < keyword.length; i++) {
                const ch = keyword[i];
                source += ch === '?' ? '[\\s\\S]'
                    : ch === '*' ? '[\\s\\S]*'
                        : escape(ch);
            }
        } else {
            source = escape(keyword);
        }
        if (options?.matchEntire === true) source = `^(?:${source})$`;
        try {
            return new RegExp(source, options?.matchCase === true ? 'g' : 'gi');
        } catch (e) {
            return null;
        }
    }

    /**
     * 转义正则元字符（`'a?b'` → `'a\\?b'`）
     *
     * **刻意不用 `text.replace(/[...]/g, '\\$&')` 这种字符类写法**：
     * 字符类里要同时含 `]` 与 `\`，写错一个（例如少写 `]`，或把 `]` 放在 `\` 之后）
     * 就会静默改变语义 —— 而这里一旦漏转义 `?`，通配符模式下的 `'a?b'` 会退化成
     * 字面量 "a?b"（`?` 只对前一个字符生效，于是永远查不到东西），
     * 表面上「查询没结果」，排查起来极费时间。
     * 下面用显式字符集 + `includes` 分派：没有字符类、没有一层转义，读一眼就知道对不对。
     * @param {string} text - 原文本
     * @returns {string} 转义后文本
     */
    private static _escapeRegExp(text: string): string {
        const SPECIAL = '.?*+^$()[]{}|\\';
        let out = '';
        for (let i = 0; i < text.length; i++) {
            const ch = text[i];
            out += SPECIAL.indexOf(ch) >= 0 ? '\\' + ch : ch;
        }
        return out;
    }

    /**
     * 取出单元格**逐字符对齐**的文本
     *
     * 与 {@link Cell.getText} 的唯一差别：这里对每个字符都产出一个位（哪怕 `char` 为
     * null/undefined 也记一个下标），保证「文本下标 ↔ chars 下标」一一对应 ——
     * 替换时要按匹配的下标去切 `chars`，两者错位会把样式/字符切错乱。
     * @param {Cell} cell - 单元格
     * @returns {string} 对齐文本（与 getText 的结果一致）
     */
    private _alignedCellText(cell: Cell): string {
        const chars = cell.chars;
        if (!chars || chars.length === 0) return '';
        let text = '';
        for (let i = 0; i < chars.length; i++) {
            const ch = chars[i];
            if (ch && ch.char !== null && ch.char !== undefined) text += ch.char;
        }
        return text;
    }

    /**
     * 稀疏收集矩形内的候选格（**行优先序**）
     *
     * ## 性能约定（勿回退为「按行列号逐格 getCellName + getCellByName」）
     * 历史实现是 `for row { for col { getCellName(...) → getCellByName(...) } }`：
     * 每次迭代都要拼一个字符串、再查一次 Map，**空网格也要一格一格走过**。
     * 整表范围（单格选区展开后的常态）下这是 行数 × 列数 次迭代 ——
     * 10 万行 × 100 列 = 一千万次，而其中 99% 的格子在 `values` 里根本不存在。
     *
     * 这里只遍历 `values`（O(存在的格子数)），顺序由行列号算出；`values` 恰好已是
     * 行优先时（装载 / 顺序编辑的主流情形）**不排序**，否则按行优先键排一次 ——
     * 命中列表必须按「行优先」呈现（查找下一个要自上而下、自左向右走）。
     * 「是否已有序」在扫描过程中顺手判定，不需要额外的探测遍历。
     *
     * 过滤顺序刻意从便宜到贵：非空文本 → 解析行列号 → 范围 → 隐藏 → 拼文本。
     * @param {SheetRect} rect - 范围矩形（1-based 闭区间）
     * @returns {ScanCandidate[]} 候选格（行优先；已排除隐藏行列与无文本的格子）
     * @private
     */
    private _collectSearchCells(rect: SheetRect): ScanCandidate[] {
        const values = this._values;
        const out: ScanCandidate[] = [];
        const { startCol, startRow, endCol, endRow } = rect;
        const width = endCol - startCol + 1;
        let ordered = true;
        let lastKey = -1;
        for (let i = 0; i < values.length; i++) {
            const cell = values[i];
            // 最便宜的判定先行：只有样式 / 空 Cell 的格子直接跳过，连名字都不用解析
            if (!DataCollection.hasAnyText(cell)) continue;
            const { col, row } = this.getCellColAndRow(cell.cell);
            if (col < startCol || col > endCol || row < startRow || row > endRow) continue;
            if (this.isRowHidden(row - 1) || this.isColHidden(col - 1)) continue;
            const key = (row - startRow) * width + (col - startCol);
            if (key < lastKey) ordered = false;
            lastKey = key;
            out.push({
                ref: cell,
                cell: cell.cell,
                col,
                row,
                text: this._alignedCellText(cell),
                key
            });
        }
        if (!ordered) out.sort((a, b) => a.key - b.key);
        return out;
    }

    /**
     * 在指定范围内找出**所有**命中项（行优先：先逐列扫完一行，再进下一行）
     *
     * 每个单元格最多产出一条命中（即便格内有多处匹配）—— 与 Excel 的命中列表口径一致；
     * 「一处 Cell 内多处匹配」只在替换时展开处理（见 {@link replaceInRect}）。
     * @param {SheetRect} rect - 查询范围
     * @param {string} keyword - 关键字
     * @param {SearchMatchOptions} [options] - 匹配选项
     * @param {number} [limit=Infinity] - 最多返回多少条（用于 UI 上限保护）
     * @returns {SearchHit[]} 命中列表（行优先升序）
     */
    public findAllInRect(
        rect: SheetRect,
        keyword: string,
        options?: SearchMatchOptions,
        limit: number = Infinity
    ): SearchHit[] {
        const pattern = this._buildSearchPattern(keyword, options);
        const hits: SearchHit[] = [];
        if (!pattern) return hits;
        const candidates = this._collectSearchCells(rect);
        for (let i = 0; i < candidates.length; i++) {
            const candidate = candidates[i];
            if (!this._matchesPattern(pattern, candidate.text)) continue;
            hits.push({ cell: candidate.cell, col: candidate.col, row: candidate.row, text: candidate.text });
            if (hits.length >= limit) break;
        }
        return hits;
    }

    /**
     * 用正则做一次「是否命中」判定（带 g 标志，必须先归零 lastIndex）
     * @param {RegExp} pattern - {@link _buildSearchPattern} 产出的正则
     * @param {string} text - 待判定文本
     * @returns {boolean} 是否命中
     */
    private _matchesPattern(pattern: RegExp, text: string): boolean {
        pattern.lastIndex = 0;
        return pattern.test(text);
    }

    /**
     * 在范围内找出 `after` 之后的**第一条**命中（找不到则回绕到首条）
     *
     * 与 {@link findAllInRect} 同源（同一次稀疏扫描），只是不建整张命中列表 ——
     * 调用方只想要「下一个」时没必要为整个范围分配数组。
     * 候选格自身已按行优先有序，因此「首个序号大于起点」的命中就是要找的那条。
     * @param {SheetRect} rect - 查询范围
     * @param {string} keyword - 关键字
     * @param {SearchMatchOptions} [options] - 匹配选项
     * @param {{ col: number; row: number }} [after] - 起点（不含）；缺省从头找，即返回首条命中
     * @returns {SearchHit | null} 命中项；无命中返回 null
     */
    public findNextInRect(
        rect: SheetRect,
        keyword: string,
        options?: SearchMatchOptions,
        after?: { col: number; row: number } | null
    ): SearchHit | null {
        const pattern = this._buildSearchPattern(keyword, options);
        if (!pattern) return null;
        const width = rect.endCol - rect.startCol + 1;
        const afterKey = after
            ? (after.row - rect.startRow) * width + (after.col - rect.startCol)
            : -1;
        const candidates = this._collectSearchCells(rect);
        let wrapTarget: SearchHit | null = null;
        for (let i = 0; i < candidates.length; i++) {
            const candidate = candidates[i];
            if (!this._matchesPattern(pattern, candidate.text)) continue;
            const hit = {
                cell: candidate.cell,
                col: candidate.col,
                row: candidate.row,
                text: candidate.text
            };
            // 回绕目标 = 首条命中（行优先序，故第一次赋值即最终值）
            if (wrapTarget === null) wrapTarget = hit;
            if (candidate.key > afterKey) return hit;
        }
        return wrapTarget;
    }

    /**
     * 判定某格**当前**是否命中，命中则返回它的命中项
     *
     * 供「替换」之后局部刷新命中列表用：替换只改动**一个**格子，重扫整个范围
     * 在整表尺度上是几百毫秒的代价，而「替换」恰恰是连点操作。
     * 本方法是单格的 O(1) 检查（就表规模而言）。
     * @param {string} name - 单元格名称（如 `'B3'`）
     * @param {string} keyword - 关键字
     * @param {SearchMatchOptions} [options] - 匹配选项
     * @returns {SearchHit | null} 命中项；不命中 / 格子不存在 / 落在隐藏行列上时返回 null
     */
    public cellMatches(name: string, keyword: string, options?: SearchMatchOptions): SearchHit | null {
        const pattern = this._buildSearchPattern(keyword, options);
        if (!pattern) return null;
        const { col, row } = this.getCellColAndRow(name);
        if (!isFinite(col) || !isFinite(row) || col < 1 || row < 1) return null;
        if (this.isRowHidden(row - 1) || this.isColHidden(col - 1)) return null;
        const cell = this.getCellByName(name);
        if (!cell) return null;
        const text = this._alignedCellText(cell);
        if (text === '' || !this._matchesPattern(pattern, text)) return null;
        return { cell: name, col, row, text };
    }

    /**
     * 在范围内把关键字替换为目标串（一步可撤销）
     *
     * 整体包在 {@link runWithCellsUndo} 内：撤销补丁按**命中的格子**逐个捕获，
     * 代价与命中处数成正比、与范围面积无关 —— 整表范围（单格选区展开后的常态）
     * 里只有 3 处命中时，也只克隆这 3 个格子，因此百万格的大表照样能 Ctrl+Z / Ctrl+Y
     * （全量快照超过 {@link LARGE_TABLE_UNDO_CELLS} 会被关闭）。
     *
     * 分两阶段：**先只读扫描收集待改格子**，再在补丁作用域内统一改写。
     * 边扫边改会让补丁捕获到的「变化前」已经是被改过的内容（撤销失效）；
     * 而把扫描塞进 `action` 里则无从预知哪些格子会变，只能退化成整矩形捕获。
     *
     * 合并区域内部的格子通常已被 merge 删除，天然不会被重复替换；若保留着
     * （带边框/底色的内部格子，见合并铁律）则按普通格子处理。
     * @param {SheetRect} rect - 替换范围
     * @param {string} keyword - 关键字
     * @param {string} replacement - 替换内容（空串 = 删除匹配文本）
     * @param {SearchMatchOptions} [options] - 匹配选项（与查询共用，保证「查得到才替换得掉」）
     * @param {SearchHit[] | null} [targets] - 仅替换这些单元格（「替换」按钮）；缺省替换范围内全部命中（「全部替换」）
     * @param {string} [label='替换'] - 撤销操作标签
     * @returns {number} 实际替换的**匹配处数**（一个格内多处匹配计多处）
     */
    public replaceInRect(
        rect: SheetRect,
        keyword: string,
        replacement: string,
        options?: SearchMatchOptions,
        targets?: SearchHit[] | null,
        label: string = '替换'
    ): number {
        const pattern = this._buildSearchPattern(keyword, options);
        if (!pattern) return 0;
        let nameFilter: Set<string> | null = null;
        if (targets) {
            nameFilter = new Set<string>();
            for (let i = 0; i < targets.length; i++) nameFilter.add(targets[i].cell);
        }
        // 阶段一（只读）：确定「哪些格子会被改动」—— 一次稀疏扫描
        const candidates = this._collectSearchCells(rect);
        const affected: Cell[] = [];
        for (let i = 0; i < candidates.length; i++) {
            const candidate = candidates[i];
            if (nameFilter && !nameFilter.has(candidate.cell)) continue;
            if (!this._matchesPattern(pattern, candidate.text)) continue;
            affected.push(candidate.ref);
        }
        if (affected.length === 0) return 0;
        // 阶段二：补丁作用域内改写。两个阶段之间没有让出，文本快照仍然有效。
        let count = 0;
        this.runWithCellsUndo(label, affected, () => {
            for (let i = 0; i < affected.length; i++) {
                count += this._replaceCharsInCell(affected[i], pattern, replacement);
            }
        });
        return count;
    }

    /**
     * 把单个单元格里**所有**匹配片段替换掉
     *
     * 字符级样式保护：被替换片段原位插入到 a+b+c 的位置；原片段若被替换为空串则该片段直接删除。
     * @param {Cell} cell - 目标单元格（values 里的活引用）
     * @param {RegExp} pattern - 带 g 标志的正则
     * @param {string} replacement - 替换内容
     * @returns {number} 替换的片段数
     */
    private _replaceCharsInCell(cell: Cell, pattern: RegExp, replacement: string): number {
        const chars = cell.chars;
        if (!chars || chars.length === 0) return 0;
        const text = this._alignedCellText(cell);
        if (text === '') return 0;
        // 先收集所有片段位置（不重叠，由正则自身的 g 语义保证），再一次性重建 chars：
        // 边替换边改 chars 会让后续片段的下标全部失效。
        //
        // 【性能约定，勿回退为「先收集位置、再统一重建 chars」的两趟写法】
        // 常见的替换（范围内每格都是 `old → new` 这类单片段、且首字符就命中）走下面这条
        // 快路径：原地改字符、样式留空（保留各处样式差异）。只有**多处片段**才退到
        // 重建路径 —— 「全部替换」动辄上万格，重建会为每格都新分配一个 chars 数组，
        // 而这些格子往往一个字符都不需要换。
        const replacementChars = Array.from(replacement);
        pattern.lastIndex = 0;
        const first = pattern.exec(text);
        if (first === null) return 0;
        if (first.index === 0 && first[0].length === text.length) {
            for (let i = 0; i < chars.length; i++) {
                chars[i].char = i < replacementChars.length ? replacementChars[i] : '';
            }
            // 短替换：多出来的字符（char = ''）参与 isEmpty 判定时会一并释放（'' 不算文字信息）
            // 长替换：样式留空，由渲染层按单元格/默认字体兜底
            return 1;
        }
        const starts: number[] = [first.index];
        const ends: number[] = [first.index + first[0].length];
        if (first[0].length === 0) pattern.lastIndex++;
        let m: RegExpExecArray | null;
        while ((m = pattern.exec(text)) !== null) {
            starts.push(m.index);
            ends.push(m.index + m[0].length);
            if (m[0].length === 0) pattern.lastIndex++; // 零宽匹配：手动前进，否则死循环
        }
        if (starts.length === 1 && ends[0] - starts[0] === text.length) {
            // 全格被替换但首字符不命中（如完全匹配时首字符被通配符吃掉）：同样走原位路径
            for (let i = 0; i < chars.length; i++) {
                chars[i].char = i < replacementChars.length ? replacementChars[i] : '';
            }
            return 1;
        }
        const out: Char[] = [];
        let cursor = 0;
        for (let i = 0; i < starts.length; i++) {
            const start = starts[i];
            const end = ends[i];
            if (start > cursor) {
                for (let k = cursor; k < start; k++) out.push(chars[k]);
            }
            if (replacementChars.length > 0) {
                // 继承被替换片段**首字符**的样式，替换进去的字保留原格式（字号/字色/加粗…）
                const styleSource = chars[start] || chars[cursor] || new Char();
                for (let k = 0; k < replacementChars.length; k++) {
                    const char = styleSource.clone();
                    char.char = replacementChars[k];
                    out.push(char);
                }
            }
            cursor = end;
        }
        for (let k = cursor; k < chars.length; k++) out.push(chars[k]);
        cell.chars = out;
        return starts.length;
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
     * 跳转到指定单元格（活动单元格 + 单格选区 + 滚动入可视区）
     *
     * 供「查找下一个 / 上一个」跳转命中项。注意它**会改写选区**，所以调用方必须
     * 事先把查询范围记住（searchDialog 用一个输入控件持有范围串），
     * 否则下一次查找的范围就会被这次跳转缩小成一个格子。
     * @param {SearchHit} hit - 命中项
     */
    public gotoSearchHit(hit: SearchHit): void {
        this.activedCell = hit.cell;
        this.selection = `${hit.cell}:${hit.cell}`;
        this.scrollCellIntoView(hit.col, hit.row);
    }

    /**
     * 通知工具栏状态更新
     */
    public notifyToolbarStateChange() {
        this.emit(DataEvents.CURSOR_STATE_CHANGED);
    }

}
