/*  # 工作表数据层职责拆分（K3）
    # 从 dataCollection.ts 原样迁出，仅调整类继承关系与成员可见性，逻辑零改动
*/
import { ColHeader, ColHeaders, RowHeader, RowHeaders, encodeColAxis, encodeRowAxis, decodeColAxis, decodeRowAxis, SparseAxis } from "../dataArchitecture/header";
import { Cell, CellStyleTable } from "../dataArchitecture/cell";
import { LineText, SelectionRange, CursorPosition } from "../dataArchitecture/cellContent";
import { Char } from "../dataArchitecture/char";
import { VisibleView } from "../dataArchitecture/visibleView";
import { DataEvents, MouseLocation } from "../constant";
import { EventEmitter } from "../../utils/eventEmitter";
import { IndexDB } from "../dataArchitecture/indexDB";
import { UndoManager, UndoAnchor, FullStateSnapshot, ContinuousUndoSnapshot } from "../../utils/undoManager";
import { ClipboardManager } from "../../utils/clipboardManager";
import { MessageBox } from "../common/message";
import { MatrixSorter } from "../dataArchitecture/matrixSorter";
import { SliceBudget } from "../../utils/scheduler";
import type { ExtractedImage } from "../../utils/parser/sheetParser";
import type { PrinterSetting } from "./sheetPrinter";
import type { SheetSnapshotV4, MergeRect } from "../dataArchitecture/dataCollection";
/**
 * 工作表存档序列化与落库基类
 *
 * 迁出自 dataCollection.ts：v3/v4 快照编解码、分片 JSON、防抖/空闲落库、切表与装载。
 * 所有公开方法名/签名与迁移前一致（继承后对外调用方无感）。
 */
export abstract class SheetStoreBase extends EventEmitter {
    /* ---------- 由派生类（DataCollection）实现的抽象成员 ---------- */
    public abstract get values(): Cell[];
    public abstract set values(value: Cell[]);
    public abstract get zoom(): number;
    public abstract set zoom(value: number);
    public abstract get activedCell(): string;
    public abstract set activedCell(value: string);
    public abstract get activedSheetName(): string;
    public abstract set activedSheetName(value: string);
    public abstract get showGridLines(): boolean;
    public abstract get selection(): string | null;
    public abstract set selection(value: string | null);
    protected abstract _refreshBaseline(): void;
    protected abstract _loadImages(sheetName: string): Promise<void>;

    constructor() {
        super();
        this.db = new IndexDB();
    }
    /** 索引数据库实例 */
    readonly db: IndexDB;
    /** 可见视图实例 */
    visibleView: VisibleView;
    /** 列头集合 */
    colHeaders: ColHeaders;
    /** 行头集合 */
    rowHeaders: RowHeaders;
    /** 当前工作表sheet，所有有内容的JSON字符串（内部后备字段，统一通过 values getter/setter 访问） */
    protected _values: Cell[] = [];
    /**
     * 本表 Python 编辑器脚本（未编写过为 null）
     *
     * 与其它元信息一样随快照落库（见 {@link SheetSnapshotV4} 的 `py-editor:script`），
     * 不单独写 settings。
     */
    protected _pyScript: string | null = null;
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
    protected _cellIndex: Map<string, Cell> = new Map();
    /** 构建索引时 `_values` 的长度：用于廉价侦测 push/splice 造成的结构变化 */
    protected _indexedLength: number = -1;
    /** 结构世代：任何可能改变「名 → 单元格」映射的操作都会自增（见 {@link invalidateCellIndex}） */
    protected _valuesEpoch: number = 0;
    /** 索引当前所处的世代；与 `_valuesEpoch` 不一致说明索引可能已过期 */
    protected _indexEpoch: number = -1;
    /** 索引是否因结构操作（重命名/增删单元格）而需要重建；属性变化（着色/格式）不设此标记 */
    protected _indexStructurallyInvalid: boolean = false;
    /**
     * 合并单元格锚点缓存（形如 `A1:C3`），随索引一并重建。
     * 原先 {@link getMergedCells} 每次调用都全表扫描一遍，而它又被
     * 「每个单元格」的矩形计算调用，构成第二处 O(N²)。
     */
    protected _mergedCells: string[] = [];
    /**
     * 合并区域的**数值矩形**缓存（与 {@link _mergedCells} 同源、同时重建）。
     *
     * 字符串形式的 `A1:C3` 每次使用都要 `split(':')` + 逐字符解析行列号，
     * 而合并查找处在「每个可见单元格 × 每帧」的热路径上（主画布的背景/边框/文本
     * 三遍绘制各查一次），临时字符串与解析开销会被放大成几百万次。
     * 这里在重建时一次性把行列号算好，热路径只做整数比较。
     */
    protected _mergeRects: MergeRect[] = [];
    /**
     * 「列号 → 覆盖该列的合并区域」倒排索引（1-based 列号）。
     *
     * 历史实现按「每个单元格 × 遍历全部合并区域」查找所属合并块，
     * 复杂度 O(可见单元格数 × 合并数)；合并单元格数以千计时，
     * 一帧就是上千万次比较 —— 实测 8000 处合并时 `drawAllMergedCellBackground`
     * 单帧 3.8 秒（渲染彻底卡死）。倒排后单次查找只与「该列上的合并块数」相关。
     */
    protected _mergeByCol: Map<number, MergeRect[]> = new Map();
    /**
     * 跨列数超过 {@link MERGE_COL_INDEX_SPAN_LIMIT} 的合并区域：不铺进列索引，
     * 查找时单独线性兜底（这类「整行合并」通常只有个位数，代价可忽略）。
     */
    protected _wideMergeRects: MergeRect[] = [];
    /**
     * 带筛选标记（`filter !== undefined`）的单元格缓存，随索引一并重建。
     * 供 SheetCanvas 每帧调用 {@link getFilterCells} 取用，
     * 避免原先每帧 `values.filter(...)` 的全表扫描。
     */
    protected _filterCells: Cell[] = [];
    /**
     * 「带筛选标记的单元格」缓存对应的世代（与 {@link _valuesEpoch} 比对判失效）
     *
     * 供 {@link invalidateFilterCells} 在**已知有筛选按钮存在**时提前失效缓存，
     * 使 `SheetCanvas.setFilters()` 能 O(1) 判定「本帧要不要再问一次筛选单元格」
     * 而无需先重建整份单元格索引。
     */
    protected _filterCellsEpoch: number = -1;
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
    protected _filterCellsDirty: boolean = true;
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
    protected _colContentBottom: Map<number, number> = new Map();
    /** {@link _colContentBottom} 所属的世代（与 {@link _valuesEpoch} 比对判失效） */
    protected _colContentEpoch: number = -1;
    /** 构建 {@link _colContentBottom} 时 `_values` 的长度（长度变化即失效） */
    protected _colContentLength: number = -1;
    /**
     * 「有文本单元格覆盖到的最大行号」（1-based；0 = 无内容）
     * 供筛选下拉清单用它替代 `matrixValues()` 求统计行范围。
     */
    protected _contentMaxRow: number = 0;
    /** 「有文本单元格覆盖到的最大列号」（1-based；0 = 无内容） */
    protected _contentMaxCol: number = 0;
    /**
     * 「有文本的行」集合（1-based）
     *
     * 筛选只可能隐藏**有内容**的行（全空行始终显示），因此判定筛选效果时
     * 只需在这批行上做，不必逐行扫到 `rowHeaders.length`（大文件下可达百万行）。
     */
    protected _rowsWithText: Set<number> = new Set();
    /**
     * {@link _contentMaxRow} / {@link _contentMaxCol} / {@link _rowsWithText} 所属的世代
     *
     * 与 {@link _colContentBottom} **分开缓存**：后者挂在筛选按钮的显隐判定上、
     * 每次数据变化（含每次按键）都会被问到，必须保持「跳过无文本单元格、不解析名字」的
     * 廉价路径；而本索引只在**筛选交互**（打开下拉 / 执行筛选 / 清除筛选）时构建，
     * 因此可以顺带解析全部有文本单元格的名字并建集合。两者混在一起会让每次按键
     * 都白做一次整表名字解析。
     */
    protected _filterIndexEpoch: number = -1;
    /** 构建筛选清单索引时 `_values` 的长度（长度变化即失效） */
    protected _filterIndexLength: number = -1;
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
    protected _colsWithContent: Set<number> = new Set();
    /**
     * 上一次提交时的 this.values 深拷贝快照。
     * 作为下一次变更的「变化前」基准：每次原地修改后调用 commitValues()，
     * 便会以本快照为 before、当前 _values 为 after 记录一条撤销操作。
     *
     * **null 表示「无基准」**：大表（见 {@link LARGE_TABLE_UNDO_CELLS}）不做整表
     * 撤销记账，此时既不为基准做深拷贝、也不入栈。读到 null 的消费点必须
     * 「重建基准但不入栈」，否则会把空基准当真基准而错误地吞掉一次撤销。
     */
    protected _lastValuesSnapshot: Cell[] | null = null;
    /** 是否正在回写撤销/重做快照（期间禁止再次记录，避免递归入栈） */
    protected _isApplyingUndo: boolean = false;
    /** 是否禁止记录撤销（初始从数据库装载时为 true，避免把首次装载误记为一次变更） */
    protected _suppressValuesUndo: boolean = false;
    /**
     * 「有未落库改动」标记
     *
     * 由 {@link _markDirty} 置位：凡改变了 {@link toJSON} 快照内容的状态变更事件
     * （单元格内容/样式、列宽行高、网格线、活动单元格与选区、缩放、打印设置）都会触发。
     * 落库时读的是**当下完整快照**，因此本标记只决定「要不要写」，不决定「写什么」。
     */
    protected _valuesDirty: boolean = false;
    /**
     * 图片是否有未落库改动
     *
     * 与 {@link _valuesDirty} 分开记：图片是二进制大块，落库要整块写一遍，
     * 不能因为「改了某个单元格」就顺手把几十 MB 图片重写一次。
     * 只有图片集合真正变化（导入文件 / 切换文件 / 显式清空）时才置位。
     */
    protected _imagesDirty: boolean = false;
    /**
     * 当前工作表内嵌图片（`ExtractedImage`，含二进制与锚点）
     *
     * 图片与单元格值分库存储（`images` 对象存储，键同为工作表名）：
     * 值走 JSON 字符串、图走结构化克隆，两者互不拖累。
     */
    protected _images: ExtractedImage[] = [];
    /** 自动保存防抖定时器（非 null 表示已排期一次落库） */
    protected _saveTimer: ReturnType<typeof setTimeout> | null = null;
    /** 空闲落库句柄（requestIdleCallback 的返回值，非 null 表示已排期下一次空闲执行） */
    protected _idleHandle: number | null = null;
    /** 自动保存防抖窗口（毫秒）：连续输入/拖拽期间合并为一次落库 */
    protected static readonly AUTO_SAVE_DEBOUNCE_MS = 500;
    /** 空闲落库的最迟等待（毫秒）：页面一直不空闲也必须在这个上限内把数据落库 */
    protected static readonly AUTO_SAVE_IDLE_TIMEOUT_MS = 1000;
    /**
     * 「大表」阈值（单元格数）
     *
     * 超过它就不再为撤销/自动记录做**整表深拷贝**（见 {@link _undoTrackingEnabled}）。
     * 依据：一次 `cloneValues` 的成本与单元格数成正比，实测约 0.55 ms/千格
     * （40 万格 ≈ 220 ms、内存上还要多留一整份），而撤销栈每次操作要拷 2~3 份。
     * 在十万格量级上还感觉不到，到了几十万格，每一次结构操作都会变成肉眼可见的停顿。
     */
    protected static readonly LARGE_TABLE_UNDO_CELLS = 100000;
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
    protected static readonly LARGE_TABLE_PATCH_CELLS = 20000;
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
    protected static readonly MERGE_COL_INDEX_SPAN_LIMIT = 64;
    /**
     * 快照序列化 / 解码的片段粒度（单元格数）
     *
     * 4000 格在 40 万格的实测口径下约合 5 ms，既能让分片版及时让出主线程，
     * 又不至于把快照切成过多片段而拖慢拼接（`join`/`+=` 的次数）。
     */
    protected static readonly SNAPSHOT_VALUE_CHUNK = 4000;
    /**
     * 全局筛选条件注册表：列号(1-based) → { anchorRow: 筛选按钮所在行(1-based), keep: 保留内容 }。
     * 注册表随全量状态快照捕获/回写，使筛选操作可整体撤销/重做；
     * 筛选组件（assistant/filter）通过 get/set/clear API 读写，不再自持静态状态。
     */
    protected _filterConditions: Map<number, { anchorRow: number; keep: Set<string> }> = new Map();
    /** 筛选操作顺序：已筛选列号按首次执行筛选的先后排列（重放筛选条件时使用） */
    protected _filterOrder: number[] = [];
    /**
     * 已隐藏行号倒排索引（1-based）
     *
     * `setRowsHidden` 同步维护：隐藏=true 加入集合，=false 移除。
     * 筛选（{@link computeFilterHiddenRows} / Filter.applyFilter）需要遍历
     * 「当前已隐藏的行」做差量更新；历史实现 `for r = anchorRow+1..totalRows`
     * 是 O(总行数) 全表扫描，百万行大表卡顿主因。改用本索引后为 O(已隐藏行数)，
     * 而已隐藏行通常只有几百到几千行。
     */
    protected _hiddenRowsSet: Set<number> = new Set();
    /** runWithFullStateUndo 嵌套深度：仅最外层推入撤销记录，嵌套调用的变更折叠进外层快照 */
    protected _fullStateUndoDepth: number = 0;
    /**
     * runWithCellPatchUndo 嵌套深度
     *
     * 包裹期间 {@link _recordValuesChange} 不再刷新整表基准（见该方法注释）：
     * 补丁自己管「变化前/后」，中间步骤刷新基准只会白白深拷贝一次整表。
     */
    protected _cellPatchUndoDepth: number = 0;
    /**
     * 编辑期间的撤销栈：保存每次按键后的 values 深拷贝快照及对应光标绝对索引。
     * 栈底（index 0）为编辑前的基准快照，供编辑内逐字撤销；
     * 整段编辑在 endEditSession() 时折叠为 undoManager 的一条文档级撤销操作。
     */
    protected _editUndoStack: { values: Cell[]; cursor: number }[] = [];
    /** 编辑期间的重做栈，配合 _editUndoStack 使用 */
    protected _editRedoStack: { values: Cell[]; cursor: number }[] = [];
    /**
     * 最近一次编辑内快照入栈时间戳（performance.now()）
     *
     * 用于时间窗口合并：连续输入/删除在约 500ms 内的多次提交合并为栈顶一条，
     * 避免「输入 abcdef 要按 6 次 Ctrl+Z 才能撤销」的过细粒度——
     * 这是用户反馈「撤销/重做不精准」的根因。窗口内仅替换栈顶快照（保留最新 cursor），
     * 超过窗口或切换操作类型（输入↔删除）才新压一条。
     */
    protected _lastEditCommitTime: number = 0;
    /**
     * 最近一次编辑提交的操作类型（'input' | 'delete' | 'style' | null）
     *
     * 时间窗口合并只在同类型内进行：连续输入合并、连续删除合并，
     * 但「输入→删除」之间不合并（让用户能分别撤销输入与删除）。
     */
    protected _lastEditCommitType: 'input' | 'delete' | 'style' | null = null;
    /**
     * 大表编辑期间的「单格版」撤销栈
     *
     * 一次编辑会话只会改动活动单元格，因此大表（整表深拷贝被禁用，见
     * {@link LARGE_TABLE_UNDO_CELLS}）下改为仅对该单元格做逐状态深拷贝，
     * 成本与表大小无关。cell 为 null 表示该状态下活动单元格尚不存在。
     * 栈底为编辑前基准；_editUndoStack 为空即表示当前处于单格记账模式。
     */
    protected _editCellUndoStack: { cell: Cell | null; cursor: number }[] = [];
    /** 大表编辑期间的单格重做栈 */
    protected _editCellRedoStack: { cell: Cell | null; cursor: number }[] = [];
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
    protected _editSessionAnchor: UndoAnchor | null = null;
    /**
     * 本次编辑会话编辑的是**哪一格**（在 {@link beginEditSession} 时记下）
     *
     * 与 {@link _editSessionAnchor} 同理：结束编辑时 `activedCell` 可能已经被
     * 鼠标移走（点到别的格子才退出编辑），若拿那时的 `activedCell` 去推
     * 「补丁目标格」，空白格的编辑会被记到**另一个格子**上 —— 撤掉的是别人的内容，
     * 用户看到的是「输入后撤销毫无反应」（新格子恰好也空白时两侧等价，干脆不入栈）。
     * 编辑会话只会改这一格，故在开始时记死。
     */
    protected _editCellName: string | null = null;
    /**
     * 编辑内撤销/重做后应恢复到的光标绝对索引。
     * 由 _undoEditStep/_redoEditStep 写入，供 editCanvas 在重测 lines、
     * 定位光标时读取（光标定位必须发生在 lines 刷新之后）。
     */
    protected _editRestoreCursor: number = 0;
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
    protected _activedSheetName: string = 'sheet1';
    /** 是否正在编辑 */
    protected _isEditting: boolean = false;
    /** 是否显示网格线 */
    protected _showGridLines: boolean = true;
    /** 当前选中的单元格 */
    protected _activedCell: string = 'A1';
    /** sheetCanvas选中区域 */
    protected _selection: string;
    /** sheetCanvas锚点选择框 */
    protected _anchorSelection: string | null;
    /** 缩放比例 */
    protected _zoom: number = 1;
    /** 打印设置 */
    protected _printSetting: PrinterSetting = {};
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
    protected _copyRange: string | null = null;
    /**
     * {@link _copyRange} 所属的工作表名
     *
     * 虚线框只属于复制发生的那张表：切到其它表时 {@link copyRange} 对外返回 null
     * （界面隐藏），切回原表时又恢复显示；内部剪贴板不受影响，切表后仍可粘贴。
     */
    protected _copyRangeSheet: string | null = null;
    /** @type {'single' | 'continuous'} 格式刷模式（单次模式或连续模式） */
    protected _brushMode: 'single' | 'continuous' | null = null;
    /** 数据加载完成的 Promise，供外部组件等待数据就绪 */
    public dataLoaded: Promise<void>;
    /** 数据加载完成的 Promise 解析函数 */
    protected _dataLoadedResolve!: () => void;
    /** 消息框实例 */
    public messageBox: MessageBox;
    /** 矩阵排序器实例 */
    protected matrixSorter: MatrixSorter;

    /**
     * 单元格是否含非空字符（不拼接字符串、不产生中间数组）
     * @param {Cell} cell - 单元格
     * @returns {boolean} 是否存在非空字符
     * @private
     */
    protected static hasAnyText(cell: Cell): boolean {
        const chars = cell.chars;
        if (!chars) return false;
        for (let i = 0; i < chars.length; i++) {
            const ch = chars[i].char;
            if (ch != null && ch !== '') return true;
        }
        return false;
    }

    /**
     * 从数据库获取当前选中工作表的所有有内容的单元格数据
     *
     * 解码走 {@link SheetStoreBase.fromJSONChunked}（分片让出）：40 万格快照的
     * 解码实测约 1.4 秒整段同步，而这里正是**刷新页面**时的装载入口。
     */
    public async syncValuesFromDB(): Promise<void> {
        await this.db.getValue(this._activedSheetName).then(async (valueRes) => {
            // 装载数据属于初始化，不应被记为一次「变更」，期间禁止记录撤销
            this._suppressValuesUndo = true;
            try {
                if(valueRes) {
                    this._applyLoadedState(await SheetStoreBase.fromSnapshotChunked(valueRes));
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
     * @param data - {@link SheetStoreBase.fromJSON} 的解析结果
     * @private
     */
    protected _applyLoadedState(data: ReturnType<typeof SheetStoreBase.fromJSON>): void {
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
        this._pyScript = data.pyScript;
        this.values = data.values;
        // 脚本随文档装载：编辑器在初始化之后才创建，靠事件让它显示新表的脚本
        this.emit(DataEvents.SHEET_SCRIPT_CHANGED, this._pyScript);
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
    protected _markDirty(): void {
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
        }, SheetStoreBase.AUTO_SAVE_DEBOUNCE_MS);
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
        }, { timeout: SheetStoreBase.AUTO_SAVE_IDLE_TIMEOUT_MS });
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
                this._applyLoadedState(await SheetStoreBase.fromSnapshotChunked(saved));
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
        // 切换/新建文件时脚本跟着文件走：旧文件的脚本文本不能留在新文件里
        this._pyScript = null;
        // values setter 会广播 VALUES_CHANGED → 主页面重绘为空表
        this.values = [];
        this.emit(DataEvents.SHEET_SCRIPT_CHANGED, null);
    }

    /**
     * 生成快照的 JSON 文本片段（**同步与分片两条路径共用的唯一实现**）
     *
     * 逐段 yields 而非一次 `JSON.stringify(整个对象)`，是为了让
     * {@link toSnapshotChunked} 能在片段之间让出主线程：几十万格的快照序列化
     * 原本是一段**不可中断**的同步任务（40 万格实测约 500 ms，且正好落在
     * 「刚打开完文件」的防抖落库时刻），是打开大文件后仍会卡顿的最后一环。
     *
     * 片段粒度按 {@link SheetStoreBase.SNAPSHOT_VALUE_CHUNK} 划分（4000 格 ≈ 5 ms），
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
        for (let i = 0; i < values.length; i += SheetStoreBase.SNAPSHOT_VALUE_CHUNK) {
            const end = Math.min(i + SheetStoreBase.SNAPSHOT_VALUE_CHUNK, values.length);
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
            // 未编写过脚本时不落字段：空白存档不必多背一个空串
            ...(this._pyScript !== null ? { 'py-editor:script': this._pyScript } : {}),
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
     * 版本标记由 {@link SheetStoreBase.fromJSON} 分派解码路径：
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
        // 空白存档对应「新文件」：不能把当前表的脚本带进去
        delete snapshot['py-editor:script'];
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
    public static async fromSnapshotChunked(raw: unknown): Promise<ReturnType<typeof SheetStoreBase.fromJSON>> {
        if (typeof raw === 'string') return SheetStoreBase.fromJSONChunked(raw);
        if (!SheetStoreBase.isV4Snapshot(raw)) {
            return SheetStoreBase._assembleSnapshot({}, []);
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
            SheetStoreBase._decodeValueRange(parsed, values, 0, parsed.length, styles, base);
            if (slice.expired()) await slice.yieldNow();
        }

        return SheetStoreBase._assembleSnapshot({
            activedCell: raw.activedCell,
            selection: raw.selection,
            printSetting: raw.printSetting,
            colHeaders: decodeColAxis(raw.cols),
            offsetWidth: raw.offsetWidth,
            rowHeaders: decodeRowAxis(raw.rows),
            offsetHeight: raw.offsetHeight,
            showGridLines: raw.showGridLines,
            zoom: raw.zoom,
            'py-editor:script': raw['py-editor:script'],
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
        values: Cell[],
        pyScript: string | null} {
        const parsed = JSON.parse(json);
        const rawValues: any[] = parsed.values || [];
        const values: Cell[] = new Array(rawValues.length);
        SheetStoreBase._decodeValueRange(rawValues, values, 0, rawValues.length, parsed.styles);
        return SheetStoreBase._assembleSnapshot(parsed, values);
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
        values: Cell[],
        pyScript: string | null}> {
        const parsed = JSON.parse(json);
        const rawValues: any[] = parsed.values || [];
        const values: Cell[] = new Array(rawValues.length);
        const slice = new SliceBudget();
        for (let i = 0; i < rawValues.length; i += SheetStoreBase.SNAPSHOT_VALUE_CHUNK) {
            const end = Math.min(i + SheetStoreBase.SNAPSHOT_VALUE_CHUNK, rawValues.length);
            SheetStoreBase._decodeValueRange(rawValues, values, i, end, parsed.styles);
            if (slice.expired()) await slice.yieldNow();
        }
        return SheetStoreBase._assembleSnapshot(parsed, values);
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
                ? SheetStoreBase.decodeLegacyCell(raw)
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
        values: Cell[],
        pyScript: string | null} {
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
            values,
            // 脚本字段名由编辑器侧约定，缺失 / 类型不对一律按「没有脚本」处理
            pyScript: typeof parsed['py-editor:script'] === 'string' ? parsed['py-editor:script'] : null,
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
}
