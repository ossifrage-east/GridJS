/*  # 编辑契约层 · 编辑宿主接口 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目描述: 基于 HTML5 Canvas 的 Web 电子表格组件，无需服务端支持即可在浏览器中创建、读取、写入 Excel 文件。
    # 项目目标: 用「外壳 / 契约 / 实现」三层夹心把「工具栏长什么样」与「它在编辑什么」彻底分开。
    # 本文件职责: 定义「编辑宿主」的全部契约 —— 命令清单（CommandId）、命令载荷（CommandPayload）、
    #             状态的归一化快照（EditStateSnapshot）、宿主接口（EditHost），以及按钮只需认识的
    #             编辑态事件源（EditingStateSource）。**只声明形状，不含任何实现**。
    # 依赖方向: 本文件只允许 import 同目录契约文件；**禁止 import sheet / document / shell 下的任何模块**，
    #           也禁止 import dataArchitecture/cell（其中含表格专属语义）。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

/**
 * 编辑宿主类型标识
 *
 * 用于「同一套工具栏外壳挂到不同载体上」时区分宿主。当前只有表格一种实现，
 * 文档工作区落地后新增 `'doc'` 即可，外壳层无需改动。
 */
export type WorkspaceKind = 'sheet' | 'doc';

/**
 * 水平对齐取值
 *
 * 刻意在此处独立声明、而非从 `dataArchitecture/cell` 引入 `TextAlign`：
 * 契约层不能依赖表格的数据模型。二者是结构等价的可选值集合，赋值时无需转换。
 */
export type TextAlignValue = 'left' | 'center' | 'right' | 'justify' | 'distribute';

/**
 * 垂直对齐取值
 *
 * 同 {@link TextAlignValue}，独立于表格侧的 `VerticalAlign`。
 */
export type VerticalAlignValue = 'top' | 'middle' | 'bottom';

/**
 * 格式刷模式
 *
 * - `null`：未启用
 * - `'single'`：单次（刷一次即退出）
 * - `'continuous'`：连续（可反复刷，直到再次点击退出）
 */
export type BrushMode = 'single' | 'continuous' | null;

/**
 * 全部可派发命令名
 *
 * ## 为什么用一张扁平字符串联合类型，而不是「每个命令一个方法」
 *
 * 工具栏的每个按钮只需要「发一个命令 + 问一个状态」，命令集合是**开集**（后续加按钮就要加命令）。
 * 用联合类型可以做到：新增命令只改这一处，`EditHost.execute` 的 `switch` 会因穷尽性检查而
 * 提醒实现方补齐；而如果改成接口方法，每加一个命令所有宿主实现都要改签名。
 *
 * ## 分组
 *
 * - **字符级**：`toggleBold` … `setTextColor`，作用于选区内的字符（表格：`chars`）
 * - **段落级**：`setAlign` … `toggleWrap`，作用于整段/整格
 * - **剪贴板**：`copy` / `cut` / `paste` 三个变体
 * - **格式刷**：`brushClick` / `brushDoubleClick`（保留原始「单击 = 单次、双击 = 连续」的手势语义）
 * - **历史**：`undo` / `redo`（快速访问区的两个按钮；两种载体各有自己的撤销栈）
 * - **段落样式（文档）**：`setParaStyle` / `toggleList` —— 表格没有「段落」概念，
 *   表格宿主的 {@link EditHost.canExecute} 对它们返回 `false`
 * - **表格专属**：`setBorder` / `toggleMergeCells` / `toggleFilter` / `sortUp` / `sortDown`
 *   —— 文档宿主对这些命令的 {@link EditHost.canExecute} 返回 `false`，按钮自动置灰
 *
 * ## 为什么「查询 / 替换」不在这里
 *
 * 它是**对话框**而不是一次编辑动作：需要拿到具体数据源去全表扫描、还要持有跨次打开的关键字
 * 与命中游标。这与「计算菜单」（`onCalculate`）同类 —— 属于外壳向宿主**注入的能力钩子**，
 * 走构造回调（`onFindReplace`）而不是命令派发。硬塞进命令集会逼宿主去 import 对话框，
 * 变成「实现依赖外壳」的反向依赖。
 */
export type CommandId =
    // ── 字符级 ──
    | 'setFontFamily'
    | 'setFontSize'
    | 'toggleBold'
    | 'toggleItalic'
    | 'toggleUnderline'
    | 'toggleStrike'
    | 'setTextColor'
    | 'setBackground'
    // ── 段落级 ──
    | 'setAlign'
    | 'setVAlign'
    | 'toggleWrap'
    | 'indentStep'
    // ── 剪贴板 ──
    | 'copy'
    | 'cut'
    | 'paste'
    | 'pasteValue'
    | 'pasteStyle'
    // ── 格式刷 ──
    | 'brushClick'
    | 'brushDoubleClick'
    // ── 历史 ──
    | 'undo'
    | 'redo'
    // ── 段落样式（文档专属）──
    | 'setParaStyle'
    | 'toggleList'
    // ── 表格专属 ──
    | 'setBorder'
    | 'toggleMergeCells'
    | 'toggleFilter'
    | 'sortUp'
    | 'sortDown';

/**
 * 段落样式标识（文档专属）
 *
 * 与 Word 的「样式」下拉同构：正文 / 标题 1–6 / 引用 / 代码块。
 * 刻意用**扁平字符串**而不是「块类型 + 级别」二元组：菜单项、按钮文字、
 * 快照字段三处都要用它，二元组会让每一处都得拆一次包。
 */
export type ParaStyleId =
    | 'paragraph'
    | 'h1'
    | 'h2'
    | 'h3'
    | 'h4'
    | 'h5'
    | 'h6'
    | 'quote'
    | 'code';

/** 列表风格（文档专属） */
export type DocListKind = 'bullet' | 'ordered';

/**
 * 命令载荷
 *
 * 统一成一个极小的结构，而不是给每个命令定义专属载荷类型：
 * - `value`：设值类命令的取值（字体名 / 字号像素 / 颜色 / 对齐方式 / 增量方向…）
 * - `toggle`：开关类命令的目标状态（由按钮当前高亮状态取反得到，见各 section 的点击处理）
 *
 * 取值类型的具体含义由命令自身约定，`EditHost` 实现在 `switch` 分支里做窄化。
 * 之所以不用泛型 `CommandPayload<T>`：命令名是运行时字符串，泛型在派发点无法关联，
 * 反而会逼调用方到处写泛型实参。
 */
export interface CommandPayload {
    /** 设值类命令的取值 */
    value?: unknown;
    /** 开关类命令的目标状态 */
    toggle?: boolean;
}

/**
 * 字符级样式的归一化快照
 *
 * 工具栏只认识这些**扁平字段**，不关心宿主内部是 `Cell` + `Char[]`（表格）还是块树（文档）。
 * 字段语义与表格的 `Char` / `FontStyle` 保持一致，使同一套按钮在两种载体上行为相同。
 *
 * 缺省（`undefined`）表示「无法确定」—— 例如表格在编辑态但光标处尚无字符时，
 * 由 UI 层回退到自己的默认值（见各 section 的 `setToolbarsState`）。
 */
export interface CharStyleState {
    /** 字体族 */
    fontFamily?: string;
    /** 字号（表格侧存的是已乘 devicePixelRatio 的像素值） */
    fontSize?: number;
    /** 加粗 */
    bold?: boolean;
    /** 斜体 */
    italic?: boolean;
    /** 下划线 */
    underline?: boolean;
    /** 删除线 */
    strike?: boolean;
    /** 字体颜色 */
    color?: string;
    /** 背景色（表格：单元格底色；文档：字符高亮） */
    background?: string;
}

/**
 * 段落/布局级样式的归一化快照
 */
export interface ParaStyleState {
    /** 水平对齐 */
    align?: TextAlignValue;
    /** 垂直对齐 */
    valign?: VerticalAlignValue;
    /** 自动换行 */
    wrap?: boolean;
    /** 缩进级数/像素（表格侧为像素累计值） */
    indent?: number;
    /** 段落样式（文档专属；表格恒为 `undefined`） */
    paraStyle?: ParaStyleId;
    /** 当前列表风格（文档专属；不在列表里为 `null`，表格恒为 `undefined`） */
    list?: DocListKind | null;
}

/**
 * 宿主状态快照
 *
 * ## 为什么是「一次取整份」而不是「一次问一个属性」
 *
 * 工具栏刷新按钮状态时，字体分区要 8 个值、对齐分区要 3 个值、格式刷要 1 个值。
 * 若按属性逐个查询，一次刷新就是十几次跨模块调用；整份快照只走一次，
 * 且**快照是一次性的纯数据**，UI 层不会因为持有它而产生对宿主的长期引用。
 *
 * ## 为什么 `editing` 与样式分开
 *
 * `editing` 决定「读光标处字符样式」还是「读活动格样式」（表格语义），
 * 这个分支属于宿主内部实现，UI 层只需要最终结果 —— 因此快照里给出的是**已经归一化好的样式**，
 * 而不是让 UI 层自己去判断该读哪里。
 */
export interface EditStateSnapshot {
    /** 是否处于字符级编辑会话（表格：单元格内容编辑中；文档：正文编辑中） */
    editing: boolean;
    /** 字符级样式（编辑态取光标处字符，非编辑态取活动格/选区代表值） */
    char: CharStyleState;
    /** 段落/布局级样式 */
    para: ParaStyleState;
    /** 筛选是否已生效（表格专属，文档宿主恒为 false） */
    filterActive: boolean;
    /** 当前格式刷模式 */
    brushMode: BrushMode;
    /**
     * 撤销栈是否还有可撤销的一步
     *
     * 快速访问区的两个按钮据此置灰。放进快照而不是让 UI 直接问 `undoManager`：
     * 两种载体的撤销栈类型完全不同（表格是 `UndoManager`，文档是块树快照栈），
     * 快照把它们归一化成一个布尔值，UI 就不再认识任何一种栈。
     */
    undoable: boolean;
    /** 重做栈是否还有可重做的一步 */
    redoable: boolean;
}

/**
 * 「编辑态」事件源 —— `BtnBase` / `Select` 只需要的那个极小契约
 *
 * 按钮的 `isDisabledOnEditting` 特性需要订阅「进入/退出编辑态」。历史实现直接吃
 * `DataCollection`，于是一个通用 UI 组件被迫依赖表格的数据集合。
 * 这里退化为「能订阅/退订事件」两个方法，`EventEmitter`（及其子类 `DataCollection`、
 * `CommandBus`）天然满足，按钮侧不再认识任何具体类型。
 */
export interface EditingStateSource {
    /**
     * 注册事件监听
     * @param event 事件名
     * @param handler 回调
     * @returns 事件源自身（便于链式调用）
     */
    on(event: string, handler: (...args: any[]) => void): unknown;
    /**
     * 移除事件监听
     * @param event 事件名
     * @param handler 回调
     * @returns 事件源自身（便于链式调用）
     */
    off(event: string, handler?: (...args: any[]) => void): unknown;
    /**
     * 载体类型（可选）
     *
     * `CommandBus` 有这个 getter；直接传 `DataCollection` 的老调用点没有。
     * 按钮据此决定「编辑态置灰」要不要生效 —— 见
     * {@link BtnBase.edittingGuardOn}：文档没有编辑会话概念，一律豁免。
     */
    readonly kind?: WorkspaceKind;
}

/**
 * 编辑宿主
 *
 * 工具栏分区（UI 外壳）与宿主实现之间**唯一**的接触面。外壳只做三件事：
 * 问「能不能点」（{@link canExecute}）、「当前是什么状态」（{@link snapshot}）、
 * 发「我要做什么」（{@link execute}）。
 *
 * @example
 * ```ts
 * // 外壳侧（不认识宿主是谁）
 * bus.dispatch('toggleBold', { toggle: !boldBtn.isActived });
 * const state = bus.snapshot();
 * boldBtn.setActived(!!state.char.bold);
 * ```
 */
export interface EditHost {
    /** 宿主类型（表格 / 文档） */
    readonly kind: WorkspaceKind;

    /**
     * 宿主的事件源
     *
     * 至少会在「状态可能变化」时发出 {@link EDIT_STATE_CHANGED}，供 {@link CommandBus} 转发。
     * 视图/UI 层只订阅总线，不直接订阅宿主。
     *
     * 【不叫 `events`】宿主实现自身往往就是 `EventEmitter`（表格宿主就是），
     * 而 `EventEmitter` 已经占用了 `events`（内部监听表 Map）这个名字，
     * 同名成员会让实现类无法同时满足两者。叫 `stateSource` 既避开冲突，
     * 也准确表达它只是「一个能订阅的源」，不承诺任何具体事件语义。
     */
    readonly stateSource: EditingStateSource;

    /**
     * 该命令在当前宿主下是否可执行
     *
     * 工具栏据此置灰按钮。表格宿主对文档专属命令、文档宿主对表格专属命令都应返回 `false`，
     * 这是「一套 UI 挂两种载体」时按钮自动收敛的机制。
     * @param cmd 命令名
     * @returns 是否可执行
     */
    canExecute(cmd: CommandId): boolean;

    /**
     * 执行命令
     *
     * 实现方负责命令的全部副作用（改数据、记账撤销、发广播）。命令不可执行时应当静默忽略，
     * 但调用方（{@link CommandBus.dispatch}）已先做了一次 {@link canExecute} 判断，实现方
     * 自行再判一次只是防御。
     * @param cmd 命令名
     * @param payload 命令载荷
     */
    execute(cmd: CommandId, payload?: CommandPayload): void;

    /**
     * 取当前状态的归一化快照
     *
     * 必须是**纯读取**：不得触发落库、广播、撤销记账，也不得新建数据对象。
     * 会被工具栏在每次状态变化时调用，实现方需保证其为 O(1) 级（表格侧靠 `getCellByName` 的名称索引）。
     * @returns 状态快照
     */
    snapshot(): EditStateSnapshot;
}

/**
 * 宿主状态可能变化时发出的事件名
 *
 * `CommandBus` 收到宿主的同名事件后转发给 UI 外壳。用自定义字符串而非 `DataEvents` 里的
 * 表格事件名，是为了让外壳层与表格的事件常量彻底脱钩。
 */
export const EDIT_STATE_CHANGED = 'edit:state:changed';

/**
 * 命令执行结果为空时的共享空快照
 *
 * 供实现方在「尚无数据」阶段返回，避免每次分配。**必须视为只读**：
 * 调用方不得修改其字段（尤其是嵌套的 `char` / `para`）。
 */
export const EMPTY_EDIT_STATE: Readonly<EditStateSnapshot> = Object.freeze({
    editing: false,
    char: Object.freeze({}) as CharStyleState,
    para: Object.freeze({}) as ParaStyleState,
    filterActive: false,
    brushMode: null,
    undoable: false,
    redoable: false,
}) as Readonly<EditStateSnapshot>;
