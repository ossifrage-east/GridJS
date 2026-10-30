/*  # 文档域 · 文档存储 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目目标: 把「文档块树」的装载、脏标记、防抖落库、撤销/重做收敛到一个类里，
    #           使视图层只需要「读 blocks / 交回 blocks」，不关心存储。
    # 本文件职责: 持有当前文档存档（DocSnapshot），与 `docs` 仓（IndexedDB）同步，
    #             并维护**独立的**撤销栈（不复用表格的 `DataCollection` 撤销栈，
    #             否则 Ctrl+Z 会把表格的改动撤到文档上）。
    # 依赖方向: dataArchitecture/indexDB（存储）+ utils/undoManager（撤销原语）+ 同目录模型/转换器。
    #           **不 import DOM**——视图通过事件订阅本类，而不是反过来。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { EventEmitter } from '../../utils/eventEmitter';
import { UndoManager } from '../../utils/undoManager';
import { IndexDB } from '../dataArchitecture/indexDB';
import { DOC_SAVE_DEBOUNCE, DOC_SNAPSHOT_VERSION, DOC_TYPING_BURST } from '../constant';
import {
    DocBlock,
    DocMeta,
    DocSnapshot,
    blockText,
    cloneBlocks,
    createBlankSnapshot,
    isEmptySnapshot,
    normalizeBlocks,
    restoreBlocks,
} from './blockTree';

/**
 * 取块树的纯文本（用于「连续输入批次是否产生净变化」的比较）
 * @param blocks - 块树
 * @returns 以换行连接的纯文本
 */
function blocksToText(blocks: DocBlock[]): string {
    return blocks.map((block) => blockText(block)).join('\n');
}

/** 文档内容发生变化（视图据此重绘） */
export const DOC_CHANGED = 'doc:changed';

/** 视图被整体替换（撤销 / 重做 / 切换文档 / 导入），视图需要重置光标 */
export const DOC_RESET = 'doc:reset';

/* 落库防抖与「连续输入合并」的时长统一见 `core/constant.ts`（DOC_SAVE_DEBOUNCE / DOC_TYPING_BURST） */

/**
 * 更新来源
 *
 * - `coalesce`：连续输入（打字），多次改动合并成一条撤销记录
 * - `single`：一次性动作（粘贴大段、改对齐、切换列表），立刻记一条撤销
 * - `none`：不记撤销（撤销/重做自身的回写、装载）
 */
export type DocUpdateMode = 'coalesce' | 'single' | 'none';

/**
 * 文档存储
 *
 * @example
 * ```ts
 * const store = new DocStore();
 * await store.load('说明.text', '说明');
 * store.on(DOC_CHANGED, () => view.render(store.blocks));
 * view.onEdit = (blocks) => store.update(blocks, '输入');
 * ```
 */
export class DocStore extends EventEmitter {
    /** 当前存档（真源） */
    private snapshotData: DocSnapshot = createBlankSnapshot();

    /** 当前文档存档键（`toArchiveKey(路径)`）；空串表示没有打开任何文档 */
    private key: string = '';

    /** 是否有未落库的改动 */
    private dirty: boolean = false;

    /** 防抖落库定时器 */
    private saveTimer: number | null = null;

    /** 文档专用撤销管理器（与表格的实例完全隔离） */
    private readonly undoManager: UndoManager = new UndoManager(100);

    /** 存储（`docs` 仓） */
    private readonly db: IndexDB = new IndexDB();

    /** 连续输入期间「第一次改动之前」的块树快照 */
    private burstBefore: DocBlock[] | null = null;

    /** 连续输入的收尾定时器 */
    private burstTimer: number | null = null;

    /** 最近一次落库失败的原因（供上层提示，不影响继续编辑） */
    private lastError: string = '';

    // ─────────────────────────── 只读访问 ───────────────────────────

    /** 当前文档存档键 */
    public get docKey(): string {
        return this.key;
    }

    /** 是否已打开文档 */
    public get hasDoc(): boolean {
        return this.key.length > 0;
    }

    /** 文档元信息 */
    public get meta(): DocMeta {
        return this.snapshotData.meta;
    }

    /** 块树（**只读**：改动请走 {@link update}） */
    public get blocks(): DocBlock[] {
        return this.snapshotData.blocks;
    }

    /** 最近一次落库错误（空串表示无错） */
    public get error(): string {
        return this.lastError;
    }

    /** 是否可撤销 */
    public get canUndo(): boolean {
        this.flushBurst();
        return this.undoManager.canUndo();
    }

    /** 是否可重做 */
    public get canRedo(): boolean {
        return this.undoManager.canRedo();
    }

    // ─────────────────────────── 装载与切换 ───────────────────────────

    /**
     * 装载一份文档
     *
     * 存档不存在时按**空文档**处理并立刻补一份空白存档：与表格侧「树里有文件、
     * 库里就得有存档」的约定一致，避免「新建的文档刷新后消失」。
     * @param docKey - 文档存档键
     * @param title - 文档标题（通常为该文件去掉后缀的名字）
     */
    public async load(docKey: string, title: string = ''): Promise<void> {
        await this.flush();
        this.key = docKey;
        this.resetUndo();

        let snapshot: DocSnapshot | null = null;
        try {
            const raw = await this.db.getDoc(docKey);
            snapshot = this.adopt(raw);
        } catch {
            snapshot = null;
        }

        if (!snapshot) {
            snapshot = createBlankSnapshot(title || this.titleFromKey(docKey));
            // 落一份空白存档：让「库里有存档」与「树里有文件」一一对应
            this.snapshotData = snapshot;
            this.dirty = true;
            this.scheduleSave();
        } else {
            this.snapshotData = snapshot;
            if (!this.snapshotData.meta.title) this.snapshotData.meta.title = title || this.titleFromKey(docKey);
        }
        this.emit(DOC_RESET);
        this.emit(DOC_CHANGED);
    }

    /**
     * 切换到另一份文档
     *
     * 与 {@link load} 的唯一差别是「同名不重载」：双击同一个文件不该重置光标与撤销栈。
     * @param docKey - 文档存档键
     * @param title - 文档标题
     * @returns 是否发生了切换
     */
    public async switchTo(docKey: string, title: string = ''): Promise<boolean> {
        if (docKey === this.key && docKey) return false;
        await this.load(docKey, title);
        return true;
    }

    /**
     * 关闭当前文档（清空内存态，但不删存档）
     */
    public async close(): Promise<void> {
        await this.flush();
        this.key = '';
        this.snapshotData = createBlankSnapshot();
        this.resetUndo();
        this.emit(DOC_RESET);
    }

    /**
     * 把从存储读回的原始对象洗成可信的存档
     *
     * 三件事：版本检查（不认的版本当空文档，而不是崩）、`Char` 原型重建、
     * 块树规范化。**原型重建不能省**：结构化克隆丢了 `Char.prototype`，
     * 之后任何 `char.clone()` 都会抛「is not a function」。
     * @param raw - `IndexDB.getDoc` 的返回值
     * @returns 可信存档；无法识别时为 null
     */
    private adopt(raw: unknown): DocSnapshot | null {
        if (!raw || typeof raw !== 'object') return null;
        const candidate = raw as DocSnapshot;
        if (typeof candidate.v !== 'number' || candidate.v > DOC_SNAPSHOT_VERSION) return null;
        if (!Array.isArray(candidate.blocks)) return null;

        const rawViewState = candidate.meta?.viewState;
        const viewState: DocMeta['viewState'] = rawViewState && typeof rawViewState === 'object'
            ? {
                zoom: typeof rawViewState.zoom === 'number' ? rawViewState.zoom : undefined,
                scrollLeft: typeof rawViewState.scrollLeft === 'number' ? rawViewState.scrollLeft : undefined,
                scrollTop: typeof rawViewState.scrollTop === 'number' ? rawViewState.scrollTop : undefined,
            }
            : undefined;
        const meta: DocMeta = {
            title: typeof candidate.meta?.title === 'string' ? candidate.meta.title : '',
            view: candidate.meta?.view === 'markdown' ? 'markdown' : 'rich',
            pageSetup: candidate.meta?.pageSetup ?? { orientation: 'portrait', paperSize: 'A4', marginPreset: 'normal', scaleMode: 'none' },
            ...(viewState ? { viewState } : {}),
        };
        return {
            v: DOC_SNAPSHOT_VERSION,
            meta,
            blocks: normalizeBlocks(restoreBlocks(candidate.blocks)),
        };
    }

    /**
     * 从存档键推断标题（去掉目录与后缀）
     * @param docKey - 存档键
     * @returns 标题
     */
    private titleFromKey(docKey: string): string {
        const slash = docKey.lastIndexOf('/');
        const base = slash >= 0 ? docKey.slice(slash + 1) : docKey;
        const dot = base.lastIndexOf('.');
        return dot > 0 ? base.slice(0, dot) : base;
    }

    // ─────────────────────────── 变更入口 ───────────────────────────

    /**
     * 提交新的块树
     *
     * 这是视图写回模型的**唯一**入口。它负责三件事：记撤销、标脏、广播。
     * @param blocks - 新块树（视图从 DOM 同步出来的结果）
     * @param label - 撤销记录的可读标签
     * @param mode - 更新来源（决定撤销如何记账）
     */
    public update(blocks: DocBlock[], label: string = '编辑', mode: DocUpdateMode = 'coalesce'): void {
        const next = normalizeBlocks(blocks);
        if (mode === 'single') {
            const before = cloneBlocks(this.snapshotData.blocks);
            this.pushUndo(before, cloneBlocks(next), label);
        } else if (mode === 'coalesce') {
            if (!this.burstBefore) this.burstBefore = cloneBlocks(this.snapshotData.blocks);
            this.resetBurstTimer();
        }
        this.snapshotData.blocks = next;
        this.markDirty();
        this.emit(DOC_CHANGED);
    }

    /**
     * 只改元信息（视图模式 / 页面设置 / 标题）
     * @param patch - 局部元信息
     */
    public updateMeta(patch: Partial<DocMeta>): void {
        this.snapshotData.meta = { ...this.snapshotData.meta, ...patch };
        this.markDirty();
        this.emit(DOC_CHANGED);
    }

    /**
     * 整体替换块树（导入 / 撤销回写用，不产生新的撤销记录）
     * @param blocks - 新块树
     */
    public replaceBlocks(blocks: DocBlock[]): void {
        this.flushBurst();
        this.snapshotData.blocks = normalizeBlocks(blocks);
        this.markDirty();
        this.emit(DOC_RESET);
        this.emit(DOC_CHANGED);
    }

    // ─────────────────────────── 撤销 / 重做 ───────────────────────────

    /**
     * 撤销
     * @returns 是否执行了撤销
     */
    public undo(): boolean {
        this.flushBurst();
        const done = this.undoManager.undo((action, isUndo) => {
            const target = isUndo ? action.data.before : action.data.after;
            this.snapshotData.blocks = cloneBlocks(target as DocBlock[]);
            this.markDirty();
        });
        if (done) {
            this.emit(DOC_RESET);
            this.emit(DOC_CHANGED);
        }
        return done;
    }

    /**
     * 重做
     * @returns 是否执行了重做
     */
    public redo(): boolean {
        this.flushBurst();
        const done = this.undoManager.redo((action, isUndo) => {
            const target = isUndo ? action.data.before : action.data.after;
            this.snapshotData.blocks = cloneBlocks(target as DocBlock[]);
            this.markDirty();
        });
        if (done) {
            this.emit(DOC_RESET);
            this.emit(DOC_CHANGED);
        }
        return done;
    }

    /**
     * 清空撤销栈（切换文档时调用，避免撤到另一份文档的内容）
     */
    private resetUndo(): void {
        this.flushBurst();
        this.burstBefore = null;
        this.undoManager.clear();
    }

    /**
     * 推一条撤销记录
     * @param before - 改动前块树（已深拷贝）
     * @param after - 改动后块树（已深拷贝）
     * @param label - 标签
     */
    private pushUndo(before: DocBlock[], after: DocBlock[], label: string): void {
        this.undoManager.push({
            type: 'docBlocks',
            data: { before, after },
            timestamp: Date.now(),
            label,
        });
    }

    /**
     * 结束当前的「连续输入」批次，把它合成一条撤销记录
     */
    private flushBurst(): void {
        if (this.burstTimer !== null) {
            clearTimeout(this.burstTimer);
            this.burstTimer = null;
        }
        if (!this.burstBefore) return;
        const before = this.burstBefore;
        this.burstBefore = null;
        // 批内没有实际变化（例如只移动了光标）就不记撤销。
        // 【为什么比文本而不是比引用】`before` 是深拷贝，引用永远不等；
        // 比「块数 + 纯文本」能识别出「输入后又删掉」这类净变化为零的批次。
        if (before.length === this.snapshotData.blocks.length
            && blocksToText(before) === blocksToText(this.snapshotData.blocks)) {
            return;
        }
        this.pushUndo(before, cloneBlocks(this.snapshotData.blocks), '输入');
    }

    /**
     * 重置连续输入的收尾定时器
     */
    private resetBurstTimer(): void {
        if (this.burstTimer !== null) clearTimeout(this.burstTimer);
        this.burstTimer = window.setTimeout(() => {
            this.burstTimer = null;
            this.flushBurst();
        }, DOC_TYPING_BURST);
    }

    // ─────────────────────────── 持久化 ───────────────────────────

    /**
     * 标记为脏并安排一次防抖落库
     */
    private markDirty(): void {
        this.dirty = true;
        this.scheduleSave();
    }

    /**
     * 安排防抖落库
     */
    private scheduleSave(): void {
        if (this.saveTimer !== null) return;
        this.saveTimer = window.setTimeout(() => {
            this.saveTimer = null;
            void this.flush();
        }, DOC_SAVE_DEBOUNCE);
    }

    /**
     * 立刻落库（切换文档 / 页面关闭 / 显式保存时调用）
     */
    public async flush(): Promise<void> {
        if (this.saveTimer !== null) {
            clearTimeout(this.saveTimer);
            this.saveTimer = null;
        }
        if (!this.dirty || !this.key) return;
        this.dirty = false;
        try {
            await this.db.setDoc(this.key, this.snapshotData);
            this.lastError = '';
        } catch (error) {
            // 落库失败不算致命：保留脏标记，下一次改动会重试
            this.dirty = true;
            this.lastError = error instanceof Error ? error.message : String(error);
        }
    }

    /**
     * 读取某份文档的存档（不改变当前内存态），供复制 / 移动 / 对账使用
     * @param docKey - 文档存档键
     * @returns 存档；不存在时为 null
     */
    public async readSnapshot(docKey: string): Promise<DocSnapshot | null> {
        try {
            return this.adopt(await this.db.getDoc(docKey));
        } catch {
            return null;
        }
    }

    /**
     * 直接写一份存档（复制 / 移动文件内容时使用）
     * @param docKey - 目标存档键
     * @param snapshot - 存档
     */
    public async writeSnapshot(docKey: string, snapshot: DocSnapshot): Promise<void> {
        await this.db.setDoc(docKey, snapshot);
    }

    /**
     * 删除一份存档
     * @param docKey - 存档键
     */
    public async deleteSnapshot(docKey: string): Promise<void> {
        await this.db.deleteDoc(docKey);
    }

    /**
     * 列出全部文档存档键
     * @returns 存档键数组
     */
    public async listDocKeys(): Promise<string[]> {
        try {
            return await this.db.getAllDocKeys();
        } catch {
            return [];
        }
    }

    /**
     * 补建空白存档（文件树上新建文件时调用；已有内容时不动）
     * @param docKey - 存档键
     * @param title - 标题
     */
    public async ensureSnapshot(docKey: string, title: string = ''): Promise<void> {
        const existing = await this.readSnapshot(docKey);
        if (existing) return;
        await this.db.setDoc(docKey, createBlankSnapshot(title || this.titleFromKey(docKey)));
    }

    /**
     * 清理孤儿存档：库里有、树里没有 —— **只删空文档**，非空的一律留着
     * @param owned - 树里仍存在的存档键集合（小写）
     */
    public async pruneOrphans(owned: Set<string>): Promise<void> {
        const keys = await this.listDocKeys();
        for (const key of keys) {
            if (!key || owned.has(key.toLowerCase())) continue;
            const snapshot = await this.readSnapshot(key);
            if (!isEmptySnapshot(snapshot)) continue;
            await this.deleteSnapshot(key);
        }
    }
}
