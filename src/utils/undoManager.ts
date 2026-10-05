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
