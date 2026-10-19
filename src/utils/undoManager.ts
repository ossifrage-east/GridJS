/**
 * 撤销/恢复管理器 - 按照 Excel 标准实现
 * 提供操作历史记录、撤销、恢复、操作组合等功能
 *
 * 核心设计：以「快照」为基本单位记录 this.values 的变更。
 * 每次 this.values 发生变化时，DataCollection 会把「变化前」与「变化后」
 * 两份深拷贝快照一并推入栈中，撤销/重做时无需复现操作过程，
 * 只需把对应快照回写即可，逻辑简单且无歧义。
 *
 * 栈状态通知：撤销栈/重做栈发生增减时（push/undo/redo/clear/截断）广播
 * {@link UNDO_STACK_CHANGED} 事件并携带 {@link UndoStackState}，
 * 供工具栏撤销/重做按钮同步可用（禁用）状态。
 *
 * @module UndoManager
 * @author 东方鹗
 * @see
 * # B站: https://space.bilibili.com/194359739
 * # 知乎: https://www.zhihu.com/people/eastossifrage
 * # CSDN: https://blog.csdn.net/os373
 */

import { EventEmitter } from './eventEmitter';

/**
 * 撤销/重做栈状态变化事件名
 * 栈内容发生增减时由 UndoManager 广播，回调参数为当前 UndoStackState
 */
export const UNDO_STACK_CHANGED = 'undo:stack:changed';

/**
 * 操作类型枚举
 * - 'values'：this.values 的前后快照（推荐用法）
 * - 'group' / 'batchOperation'：组合操作，data.actions 为子操作列表
 * - 其它字符串：兼容历史自定义操作
 */
export type UndoActionType = 'values' | 'group' | 'batchOperation' | string;

/**
 * 值快照接口
 * before / after 均为调用方传入的深拷贝，UndoManager 不会再修改它们，
 * 因此同一份快照在栈中与在恢复回调中可安全共享只读引用。
 */
export interface ValuesSnapshot {
    /** 变更前的 this.values 深拷贝 */
    before: any[];
    /** 变更后的 this.values 深拷贝 */
    after: any[];
}

/**
 * 全量状态快照接口
 * 用于「绕过 values setter 的结构性操作」（如右键菜单的删除行列、隐藏行列、
 * 清除内容、设置格式等）：除 values 外还需捕获行/列表头（含隐藏标记、行高列宽）、
 * 网格线开关、活动单元格与选区，才能完整撤销。
 */
export interface FullStateSnapshot {
    /** 单元格数组深拷贝 */
    values: any[];
    /**
     * 内嵌图片数组（**浅拷贝**：只拷数组本身，图片对象按引用共享）
     *
     * 图片是二进制大块，深拷贝一次 undo 快照就要多留一份几十 MB 的内存；
     * 而图片对象在本项目里是「只读替换」语义（永不原地修改），共享引用是安全的。
     * 可选：旧快照没有该字段。
     */
    images?: any[];
    /** 行表头数组深拷贝（top / height / isHidden 等） */
    rowHeaders: any[];
    /** 列表头数组深拷贝（left / width / isHidden 等） */
    colHeaders: any[];
    /** 是否显示网格线 */
    showGridLines: boolean;
    /** 活动单元格名称 */
    activedCell: string;
    /** 选区字符串（如 "A1:C3"） */
    selection: string | null;
    /** 筛选条件列表（可选，旧快照无此字段）：按筛选操作顺序排列，
     *  col=筛选列号(1-based)，anchorRow=筛选按钮所在行(1-based)，keep=保留内容（Set 的数组形式） */
    filters?: Array<{ col: number; anchorRow: number; keep: string[] }>;
}

/**
 * 跨事件连续操作（如鼠标拖拽调整行高/列宽）的「变化前」基准快照
 *
 * 整个拖拽手势折叠为一条撤销记录：手势开始时捕获一次 before，
 * 手势结束（mouseup）时捕获 after 并入栈。
 * - 小表：{@link FullStateSnapshot}（整表前后快照）；
 * - 大表：{@link StatePatchSnapshot}（'headers' 档，只含行/列表头浅拷贝 +
 *   网格线/图片/筛选，不克隆任何格子 —— 拖拽行高列宽只改表头，
 *   一万多行表头的浅拷贝也仅几毫秒，代价与整表大小无关）。
 */
export type ContinuousUndoSnapshot =
    | { kind: 'full'; snapshot: FullStateSnapshot }
    | { kind: 'patch'; snapshot: StatePatchSnapshot };

/**
 * 局部单元格补丁的一项
 *
 * `data` 为 **null 表示该格子在该时刻不存在**：合并会把区域内除左上角外的
 * 格子整体删掉，取消合并又会把它们补回来，补丁必须能表达「存在 ↔ 不存在」
 * 两种方向，否则撤销/重做会在大表上留下空洞或残骸。
 */
export interface CellPatchEntry {
    /** 单元格名称（如 `B7`） */
    cell: string;
    /** 该时刻的单元格克隆；null = 不存在 */
    data: any | null;
}

/**
 * 局部单元格补丁快照接口
 *
 * 用于「只影响矩形区域内少量单元格」的操作（合并/取消合并、区域着色、区域清内容…）：
 * 只记录区域内每个格子操作前后的状态，代价与**区域大小**成正比、与整表大小无关。
 *
 * 之所以需要它：整表快照（{@link ValuesSnapshot} / {@link FullStateSnapshot}）在
 * 大表上会被调用方按阈值整体关掉（一次操作要 2~3 份整表深拷贝，几十万格时
 * 单次就上百毫秒），于是「大表里合并 50 个单元格再 Ctrl+Z」会静默失效。
 * 补丁把成本压到「选区格数」，百万格的表也能正常撤销/重做。
 */
export interface CellPatchSnapshot {
    /** 操作前：区域内每个格子的克隆（不存在为 null） */
    before: CellPatchEntry[];
    /** 操作后：区域内每个格子的克隆（不存在为 null） */
    after: CellPatchEntry[];
}

/**
 * 撤销/重做后应恢复的「视图锚点」
 *
 * 选区与活动单元格**不属于**任何一种数据快照，但必须随快照一起存取：
 * 撤销「在第 5000 行做的格式修改」时，若视口仍停在第 1 行，用户根本看不见
 * 刚被还原的内容，会误判成「撤销没生效」。因此每种快照类型都在 action 上
 * 额外挂一对锚点（before / after），由 DataCollection 统一捕获与回写。
 */
export interface UndoAnchor {
    /** 活动单元格名称（如 `B7`） */
    activedCell: string;
    /** 选区字符串（如 `A1:C3`）；null = 无选区 */
    selection: string | null;
}

/**
 * 稀疏状态快照接口（大表专用）
 *
 * 与 {@link FullStateSnapshot} 的差别只有一处：`values` 从「整表深拷贝」换成
 * **影响面内的格子补丁**（{@link CellPatchEntry} 列表）。行/列表头、网格线、
 * 图片、筛选条件仍然全量捕获 —— 它们要么条数远小于单元格（1.9M 格的表
 * 也就一万多行表头），要么本来就是标量，深/浅拷贝的代价可以忽略。
 *
 * 之所以需要它：百万格的表做一次整表深拷贝要 1.7 秒（实测），整表快照这条路
 * 在大表上被 {@link DataCollection} 按阈值整体关掉，于是「大表里清除一片区域
 * 再 Ctrl+Z」会静默失效。补丁把成本压到「影响面格数」，于是大表也能正常撤销。
 *
 * **前提**：`cells` 必须完整覆盖本次操作的影响面。影响面会溢出选区的操作
 * （插入/删除整行整列会把下方/右侧所有格子重映射）不能用本类型记账，
 * 否则撤销后只还原了选区那一块、其余格子停在改后状态 —— 数据自相矛盾。
 */
export interface StatePatchSnapshot {
    /**
     * 影响面内格子在**该时刻**的状态（行优先，两条快照同序）
     *
     * 单侧设计：before 快照与 after 快照各自只装自己那一半，
     * 由外层 `UndoAction` 的 before/after 字段配对，不再在内部再分一次前后
     * （多一层嵌套只会让「取错侧」这类错误更难一眼看出）。
     * `'headers'` 档（操作不动 values）恒为空数组。
     */
    cells: CellPatchEntry[];
    /** 行表头数组浅拷贝（每项 {...rh}） */
    rowHeaders: any[];
    /** 列表头数组浅拷贝（每项 {...ch}） */
    colHeaders: any[];
    /** 是否显示网格线 */
    showGridLines: boolean;
    /** 内嵌图片数组（浅拷贝，与 FullStateSnapshot 同口径） */
    images?: any[];
    /** 筛选条件列表（可选） */
    filters?: Array<{ col: number; anchorRow: number; keep: string[] }>;
}

/**
 * 结构性操作快照接口（插入/删除整行整列，大表专用）
 *
 * 采用**命令模式**（记录「操作参数 + 被删数据」），而不是前后两份状态快照：
 * - 插入行/列不销毁任何数据，撤销 = 删掉刚插入的那段；重做 = 再插一次；
 * - 删除行/列销毁被删带上的格子与表头，撤销时除逆操作（结构上插回去，
 *   下方/右侧格子自动移回原位）外，还需把这份「被删数据」原样放回。
 *
 * 之所以不用前后快照：插入/删除会把下方/右侧**所有**格子重命名，
 * 前后快照等价于两份整表深拷贝 —— 正是大表下被 {@link DataCollection}
 * 按阈值整体关掉的那条路；而命令记录的代价只与「被删带内的格子数」
 * 成正比，与整表大小无关。
 */
export interface StructuralSnapshot {
    /** 操作类别：行/列 × 插入/删除 */
    op: 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols';
    /** 插入位置（insert）或删除带起始行/列号（delete），1-based */
    at: number;
    /** 插入/删除的数量（delete 时 = 实际删除数，已按表尾收敛） */
    count: number;
    /** delete 专用：被删带上的格子克隆（名字为删除前的名字）；insert 时为空数组 */
    removedCells: CellPatchEntry[];
    /** delete 专用：被删带上的行/列头条目（含自定义行高列宽与隐藏标记）；insert 时为空数组 */
    removedHeaders: any[];
    /**
     * delete 专用：被收缩的「跨越删除带」合并锚点的原始跨距
     *
     * 删除会把锚点在删除带**上方/左侧**、但跨距末端伸进删除带的合并收缩
     * （收缩量 = 与删除带的重叠数，可能小于 count）。撤销时的结构插回只能
     * 按 count 还原「整段跨越」的合并，端点落在删除带内部的合并无法由它
     * 还原 —— 用这里记录的原始值按名字直接赋回（赋值语义，天然幂等）。
     */
    shrunkMerges: Array<{ cell: string; rowspan?: number; colspan?: number }>;
}

/**
 * 排序排列快照接口（大表专用的可撤销排序记录）
 *
 * 排序本质上只是选定行/列范围内的一次**置换**，没有任何格子被创建或销毁。
 * 旧实现无论表多大都克隆两份整表（排序前/后各一份，含全部 chars 深拷贝），
 * 百万格单次排序仅撤销记账就要数秒与上百 MB 内存。改为只记两条索引数组：
 * - `positions[i]`（0-based 绝对行/列号）位置上，排序后放着原来位于
 *   `origins[i]` 的那一行/列；
 * - 撤销 = 交换两个数组的角色再做一次置换，重做 = 按原数组再置换一次。
 *
 * 记账成本 O(参与排序的行/列数)（仅两个数字数组），与整表格子数无关；
 * 执行一次撤销/重做是 O(实际格子数) 的单遍改名，与排序自身同阶。
 *
 * axis 沿用 MatrixSorter 的语义：'column' = 按列排序**行**（置换行），
 * 'row' = 按行排序**列**（置换列）。
 */
export interface SortPermutationSnapshot {
    /** 排序轴：'column' 时置换行，'row' 时置换列 */
    axis: 'column' | 'row';
    /** 目标位置序列（0-based 绝对行/列号，升序且已剔除筛选锚点） */
    positions: number[];
    /** 排序后落在 {@link positions} 各位置上的原始行/列号序列 */
    origins: number[];
}

/**
 * 撤销操作接口
 * 记录单个操作的所有信息
 */
export interface UndoAction {
    /** 操作类型 */
    type: UndoActionType;
    /** 操作数据（'values' 类型时为 ValuesSnapshot） */
    data: ValuesSnapshot | { actions: UndoAction[]; groupId: string } | any;
    /** 操作时间戳 */
    timestamp: number;
    /** 操作组 ID（用于组合操作） */
    groupId?: string;
    /** 可读标签，便于调试与撤销菜单展示 */
    label?: string;
    /**
     * 撤销后应恢复到的视图锚点（操作发生**前**的选区与活动格）
     *
     * 可选：旧快照没有该字段，缺失时 DataCollection 只还原数据、不动选区。
     */
    anchorBefore?: UndoAnchor;
    /**
     * 重做后应恢复到的视图锚点（操作发生**后**的选区与活动格）
     */
    anchorAfter?: UndoAnchor;
}

/**
 * 撤销栈状态接口
 */
export interface UndoStackState {
    /** 撤销栈大小 */
    undoSize: number;
    /** 重做栈大小 */
    redoSize: number;
    /** 是否可撤销 */
    canUndo: boolean;
    /** 是否可重做 */
    canRedo: boolean;
}

/**
 * 恢复回调函数类型
 * @param action 需要恢复的操作
 * @param isUndo true=正在撤销（应回写 before）；false=正在重做（应回写 after）
 */
export type UndoRestoreCallback = (action: UndoAction, isUndo: boolean) => void;

/**
 * UndoManager 类
 * 管理电子表格的撤销/恢复操作，支持快照回写、操作组合和历史记录管理
 *
 * @class UndoManager
 * @example
 * const undoManager = new UndoManager(50);
 *
 * // 记录一次 this.values 变更（before/after 均为深拷贝）
 * undoManager.pushValuesSnapshot(oldValues, newValues, '合并单元格');
 *
 * // 执行撤销：回调收到 isUndo=true，应回写 action.data.before
 * undoManager.undo((action, isUndo) => {
 *     if (action.type === 'values') {
 *         applySnapshot(isUndo ? action.data.before : action.data.after);
 *     }
 * });
 */
export class UndoManager extends EventEmitter {
    /** 撤销栈 */
    private undoStack: UndoAction[] = [];
    /** 重做栈 */
    private redoStack: UndoAction[] = [];
    /** 最大栈大小 */
    private maxStackSize: number = 100;
    /** 是否正在执行撤销 */
    private _isUndoing: boolean = false;
    /** 是否正在执行重做 */
    private _isRedoing: boolean = false;
    /** 当前操作组 ID */
    private currentGroupId: string | null = null;
    /** 当前组的操作列表 */
    private groupActions: UndoAction[] = [];

    /**
     * 创建 UndoManager 实例
     * @param {number} [maxSize=100] - 最大历史记录数（0-100）
     */
    constructor(maxSize: number = 100) {
        super();
        this.maxStackSize = Math.max(0, Math.min(100, maxSize));
    }

    /**
     * 获取最大撤销次数
     * @returns {number} 最大撤销次数
     */
    public get maxUndoLevels(): number {
        return this.maxStackSize;
    }

    /**
     * 设置最大撤销次数（范围 0-100）
     * @param {number} value - 最大撤销次数
     */
    public set maxUndoLevels(value: number) {
        this.maxStackSize = Math.max(0, Math.min(100, value));
        // 如果新的最大值小于当前栈大小，截断旧的记录
        const sizeBefore = this.undoStack.length;
        while (this.undoStack.length > this.maxStackSize) {
            this.undoStack.shift();
        }
        if (this.undoStack.length !== sizeBefore) {
            this.emit(UNDO_STACK_CHANGED, this.getStackState());
        }
    }

    /**
     * 是否正在执行撤销（供外部在恢复回写期间抑制再次记录）
     * @returns {boolean}
     */
    public get isUndoing(): boolean {
        return this._isUndoing;
    }

    /**
     * 是否正在执行重做（供外部在恢复回写期间抑制再次记录）
     * @returns {boolean}
     */
    public get isRedoing(): boolean {
        return this._isRedoing;
    }

    /**
     * 开始一个操作组（用于组合操作）
     * 同一组的操作会被合并为一步撤销。
     * 典型场景：单元格编辑期间，每次按键都会产生一次 values 快照，
     * 用 beginGroup/endGroup 包裹后，整段输入只会折叠成一步撤销。
     * @param {string} [groupId] - 操作组 ID（可选，默认自动生成）
     */
    public beginGroup(groupId?: string): void {
        // 已存在未关闭的组时先安全收尾，避免子操作丢失
        if (this.currentGroupId) {
            this.endGroup();
        }
        this.currentGroupId = groupId || `group-${Date.now()}`;
        this.groupActions = [];
    }

    /**
     * 结束当前操作组
     * 组内全部为 values 快照时，会折叠为单条 { before: 首条.before, after: 末条.after }，
     * 既节省内存又保证撤销/重做一步到位；否则按普通组合操作存储。
     */
    public endGroup(): void {
        if (!this.currentGroupId) {
            this.groupActions = [];
            return;
        }
        const groupId = this.currentGroupId;
        const actions = this.groupActions;
        // 先清空组状态，使后续 push 直接落入 undoStack
        this.currentGroupId = null;
        this.groupActions = [];
        if (actions.length === 0) return;
        if (actions.length === 1) {
            this.push(actions[0]);
            return;
        }
        this.push(this.collapseValueSnapshots(actions, groupId));
    }

    /**
     * 折叠连续的 values 快照
     * 仅当全部子操作均为 'values' 类型时才折叠为首尾单条快照，
     * 否则按 'group' 组合操作原样保留，确保非 values 子操作不被吞并。
     * @param {UndoAction[]} actions - 组内子操作
     * @param {string} groupId - 组 ID
     * @returns {UndoAction} 折叠或封装后的操作
     */
    private collapseValueSnapshots(actions: UndoAction[], groupId: string): UndoAction {
        const allValues = actions.every(a =>
            a.type === 'values' && a.data && Array.isArray((a.data as ValuesSnapshot).before) && Array.isArray((a.data as ValuesSnapshot).after)
        );
        if (allValues) {
            const first = actions[0].data as ValuesSnapshot;
            const last = actions[actions.length - 1].data as ValuesSnapshot;
            return {
                type: 'values',
                data: { before: first.before, after: last.after },
                timestamp: Date.now(),
                groupId,
                label: actions[0].label
            };
        }
        return {
            type: 'group',
            data: { actions: [...actions], groupId },
            timestamp: Date.now(),
            groupId
        };
    }

    /**
     * 添加操作到撤销栈
     * @param {UndoAction} action - 操作对象
     */
    public push(action: UndoAction): void {
        // 如果正在撤销或重做，不记录新操作
        if (this._isUndoing || this._isRedoing) return;

        // 如果在操作组中，先添加到组里
        if (this.currentGroupId) {
            this.groupActions.push(action);
            return;
        }

        this.undoStack.push(action);

        // 如果达到最大容量，移除最旧的记录
        if (this.undoStack.length > this.maxStackSize) {
            this.undoStack.shift();
        }

        // 清空重做栈（因为新操作会清空重做历史）
        this.redoStack = [];

        // 栈内容已变化，通知监听方（如工具栏按钮）同步可用状态
        this.emit(UNDO_STACK_CHANGED, this.getStackState());
    }

    /**
     * 快捷方法：推送一次 this.values 的前后快照
     * @param {any[]} before - 变更前的深拷贝快照
     * @param {any[]} after - 变更后的深拷贝快照
     * @param {string} [label] - 可读标签
     */
    public pushValuesSnapshot(before: any[], after: any[], label?: string): void {
        this.push({ type: 'values', data: { before, after }, timestamp: Date.now(), label });
    }

    /**
     * 执行撤销
     * @param {UndoRestoreCallback} restoreCallback - 恢复回调函数
     * @returns {boolean} 是否撤销成功
     */
    public undo(restoreCallback: UndoRestoreCallback): boolean {
        if (!this.canUndo()) return false;

        this._isUndoing = true;
        try {
            const action = this.undoStack.pop();
            if (action) {
                this.redoStack.push(action);
                this.applyAction(action, restoreCallback, true);
                return true;
            }
        } finally {
            this._isUndoing = false;
            // 栈内容已变化（undoStack → redoStack），回写完成后通知监听方同步可用状态
            this.emit(UNDO_STACK_CHANGED, this.getStackState());
        }
        return false;
    }

    /**
     * 执行重做
     * @param {UndoRestoreCallback} restoreCallback - 恢复回调函数
     * @returns {boolean} 是否重做成功
     */
    public redo(restoreCallback: UndoRestoreCallback): boolean {
        if (!this.canRedo()) return false;

        this._isRedoing = true;
        try {
            const action = this.redoStack.pop();
            if (action) {
                this.undoStack.push(action);
                this.applyAction(action, restoreCallback, false);
                return true;
            }
        } finally {
            this._isRedoing = false;
            // 栈内容已变化（redoStack → undoStack），回写完成后通知监听方同步可用状态
            this.emit(UNDO_STACK_CHANGED, this.getStackState());
        }
        return false;
    }

    /**
     * 应用单个操作（处理组合操作的展开）
     * @param {UndoAction} action - 操作对象
     * @param {UndoRestoreCallback} restoreCallback - 恢复回调函数
     * @param {boolean} isUndo - true=撤销（回写 before），false=重做（回写 after）
     */
    private applyAction(action: UndoAction, restoreCallback: UndoRestoreCallback, isUndo: boolean): void {
        if ((action.type === 'group' || action.type === 'batchOperation') && action.data && Array.isArray((action.data as { actions: UndoAction[] }).actions)) {
            const subActions: UndoAction[] = (action.data as { actions: UndoAction[] }).actions;
            // 撤销时反向回写，重做时正向回写；方向对所有子操作一致
            const order = isUndo ? [...subActions].reverse() : subActions;
            for (const subAction of order) {
                restoreCallback(subAction, isUndo);
            }
        } else {
            restoreCallback(action, isUndo);
        }
    }

    /**
     * 检查是否可以撤销
     * @returns {boolean} 是否可以撤销
     */
    public canUndo(): boolean {
        return this.undoStack.length > 0;
    }

    /**
     * 检查是否可以重做
     * @returns {boolean} 是否可以重做
     */
    public canRedo(): boolean {
        return this.redoStack.length > 0;
    }

    /**
     * 查看撤销栈顶操作（不移除），便于 UI 展示下一步将撤销的内容
     * @returns {UndoAction | null} 栈顶操作，栈空时返回 null
     */
    public peek(): UndoAction | null {
        return this.undoStack.length > 0 ? this.undoStack[this.undoStack.length - 1] : null;
    }

    /**
     * 清空所有记录
     */
    public clear(): void {
        this.undoStack = [];
        this.redoStack = [];
        this.groupActions = [];
        this.currentGroupId = null;
        this.emit(UNDO_STACK_CHANGED, this.getStackState());
    }

    /**
     * 获取撤销栈大小
     * @returns {number} 撤销栈大小
     */
    public getUndoStackSize(): number {
        return this.undoStack.length;
    }

    /**
     * 获取重做栈大小
     * @returns {number} 重做栈大小
     */
    public getRedoStackSize(): number {
        return this.redoStack.length;
    }

    /**
     * 获取当前栈状态
     * @returns {UndoStackState} 栈状态对象
     */
    public getStackState(): UndoStackState {
        return {
            undoSize: this.undoStack.length,
            redoSize: this.redoStack.length,
            canUndo: this.canUndo(),
            canRedo: this.canRedo()
        };
    }

    /**
     * 获取撤销历史记录（用于调试）
     * @returns {UndoAction[]} 撤销历史记录副本
     */
    public getUndoHistory(): UndoAction[] {
        return [...this.undoStack];
    }

    /**
     * 获取重做历史记录（用于调试）
     * @returns {UndoAction[]} 重做历史记录副本
     */
    public getRedoHistory(): UndoAction[] {
        return [...this.redoStack];
    }
}
