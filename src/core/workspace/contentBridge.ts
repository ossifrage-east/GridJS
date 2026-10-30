/*  # 工作区职责拆分（K4）
    # 从 app.ts 原样迁出，仅调整类继承关系与成员可见性，逻辑零改动
*/
import { DataCollection } from '../dataArchitecture/dataCollection';
import type { FileMovePlan } from '../common/fileTree/fileTree';
import { createBlankSnapshot } from '../document/blockTree';
import { stripExtension } from '../../utils/fileName';
import { ACTIVED_FILE_KEY } from './fileSession';
import { OpenRouterBase } from './openRouter';

/**
 * 内容桥基类：新增 / 复制 / 改名 / 移动 / 删除文件时同步数据库里的内容
 *
 * 参数一律是**路径**（含扩展名、相对根节点），存档键由 `toSheetKey` 换算。
 * 表格走 `values` / `images` 仓，文档走 `docs` 仓，按载体分派。
 */
export abstract class ContentBridgeBase extends OpenRouterBase {
    /**
     * 补建文件存档（新增文件时由内容桥调用）
     *
     * 让「树里有这个文件」与「库里有这个文件的存档」始终一一对应：库里还没有
     * 该文件的存档时写入一份空白快照。
     *
     * **只补不覆盖**：这条路径也会在「登记已打开的文件」时被调用，那时内容刚刚由
     * 打开器写库，若直接覆写会把刚打开的文件清空。
     * @param paths - 新增文件的路径集合
     */
    protected async ensureFileContents(paths: string[]): Promise<void> {
        try {
            await this.data.dataLoaded;
            for (const path of paths) {
                const key = this.toSheetKey(path);
                if (!key) continue;
                // 文档文件补的是 docs 仓的空白存档，不是表格的空白快照：
                // 往 values 库写一份「空表」会让这个键同时被两个库认领，
                // 删除时容易只清一边而留下孤儿
                if (this.isDocPath(path)) {
                    await this.docWorkspace?.docStore.ensureSnapshot(key, this.docTitleOf(path));
                    continue;
                }
                const existing = await this.settingsDb.getValue(key);
                if (existing && existing !== '{}') continue;   // 已有存档：保持原样
                await this.settingsDb.setValue(key, this.data.blankSnapshot());
            }
        } catch {
            // 存档不可用不影响结构操作：此时新文件按空表处理
        }
    }

    /**
     * 复制文件内容（粘贴 / 复制文件夹时由内容桥调用）
     *
     * 源即当前打开的文件时，先把内存中的最新内容落库再复制，
     * 否则会把「打开后尚未触发自动保存」的旧版本复制过去。
     * 源为空表时同时清掉目标键上的旧存档，避免粘贴出来的文件显示上一个同名文件的内容。
     *
     * 存档键是**路径**：因此「不同目录下的同名文件」是完全独立的两个存档
     * （历史实现只按文件名当键，会把它们混在一起）。
     * @param plans - 搬迁计划（源路径 → 目标路径，文件夹已递归展开为文件）
     */
    protected async copyFileContents(plans: FileMovePlan[]): Promise<void> {
        try {
            await this.data.dataLoaded;
            const sources = plans
                .filter((plan) => plan.from && plan.to && this.toSheetKey(plan.from) !== this.toSheetKey(plan.to))
                .map((plan) => this.toSheetKey(plan.from));
            if (sources.includes(this.data.activedSheetName)) {
                await this.data.syncValuesToDB();
            }

            for (const plan of plans) {
                const from = this.toSheetKey(plan.from);
                const to = this.toSheetKey(plan.to);
                if (!from || !to || from === to) continue;

                // 文档：复制块树（源为空文档时写一份空白存档，与表格侧「空表也占位」同口径）
                if (this.isDocPath(plan.from) || this.isDocPath(plan.to)) {
                    const store = this.docWorkspace?.docStore;
                    if (!store) continue;
                    const snapshot = await store.readSnapshot(from);
                    await store.writeSnapshot(to, snapshot ?? createBlankSnapshot(this.docTitleOf(plan.to)));
                    continue;
                }

                const raw = await this.settingsDb.getValue(from);
                // 源为空表时也写入空白存档（而不是删掉目标键）：目标文件同样是
                // 「树里有文件、库里就得有存档」，同时顺带清掉目标键上可能存在的旧内容
                await this.settingsDb.setValue(to, raw && raw !== '{}' ? raw : this.data.blankSnapshot());
                // 图片与单元格值分库存储，必须同步复制：否则粘贴出来的文件
                // 要么没有图，要么顶着目标键上残留的旧图
                const images = await this.settingsDb.getImages(from);
                if (images.length > 0) await this.settingsDb.setImages(to, images);
                else await this.settingsDb.deleteImages(to);
            }
        } catch {
            // 存档不可用不影响结构操作：此时粘贴出来的文件按空表处理
        }
    }

    /**
     * 搬迁文件内容（重命名 / 剪切粘贴 / 拖拽移动时由内容桥调用）
     *
     * 内容存档的键就是**路径**，因此只要路径变了（改名、剪切到别的目录、拖拽换目录、
     * 祖先文件夹改名）就必须把存档搬到新键下，否则目标任务显示为空表。
     * 被改名的正是当前打开的文件时，同步更新 `activedSheetName`（后续自动保存要写到新键）
     * 与「当前文件」记忆。
     * @param plans - 搬迁计划（源路径 → 目标路径，文件夹已递归展开为文件）
     */
    protected async moveFileContents(plans: FileMovePlan[]): Promise<void> {
        try {
            await this.data.dataLoaded;
            for (const plan of plans) {
                const from = this.toSheetKey(plan.from);
                const to = this.toSheetKey(plan.to);
                if (!from || !to || from === to) {
                    // 存档键没变：无需搬迁。
                    // 注意这与历史实现的判断不同 —— 路径键下「剪切到别的目录」确实
                    // 改变存档键（`a/x.xlsx` → `b/x.xlsx`），必须搬迁，
                    // 只有「同一目录内名字也没变」才是真正的空操作。
                    continue;
                }

                // 文档：搬运块树（先落库再读，避免搬走的是防抖窗口里的旧版本）
                if (this.isDocPath(plan.from) || this.isDocPath(plan.to)) {
                    const store = this.docWorkspace?.docStore;
                    if (!store) continue;
                    if (from === store.docKey) await store.flush();
                    const snapshot = await store.readSnapshot(from);
                    await store.writeSnapshot(to, snapshot ?? createBlankSnapshot(this.docTitleOf(plan.to)));
                    await store.deleteSnapshot(from);
                    // 改名/移动的正是当前文档：后续自动保存要写到新键
                    if (from === store.docKey) {
                        this.docWorkspace.render();
                    }
                    this.rememberRenamedFile(plan.from, plan.to);
                    continue;
                }

                // 源即当前工作表：先把内存内容落库，否则搬走的是旧版本
                if (from === this.data.activedSheetName) await this.data.syncValuesToDB();

                const raw = await this.settingsDb.getValue(from);
                // 空表也写一份空白存档：改名后「树里有文件、库里就得有存档」，不留空档
                await this.settingsDb.setValue(to, raw && raw !== '{}' ? raw : this.data.blankSnapshot());
                await this.settingsDb.deleteValue(from);
                // 图片记录同样随路径搬迁（键就是路径）
                const images = await this.settingsDb.getImages(from);
                if (images.length > 0) await this.settingsDb.setImages(to, images);
                else await this.settingsDb.deleteImages(to);
                await this.settingsDb.deleteImages(from);

                // 被改动的是当前打开的文件：后续自动保存必须写到新键，
                // 否则编辑内容会被写回旧键（等于没改动的那个文件被悄悄改掉）
                if (from === this.data.activedSheetName) this.data.activedSheetName = to;
                this.rememberRenamedFile(plan.from, plan.to);
            }
        } catch {
            // 存档不可用不影响结构操作
        }
    }

    /**
     * 删除文件内容（删除树节点时由内容桥调用，文件夹已递归展开为文件）
     *
     * 被删的正是当前打开的文件时，先回退到**上一个打开过的文件**并高亮它
     * （见 {@link ContentBridgeBase.pickFallbackFile}）；无文件可回退则清空主页面。
     * 注意顺序：切换/清空会先固化当前内容，因此存档删除必须放在其后，
     * 否则刚写回的内容会留在库里，看起来「删不掉」。
     * @param paths - 被删除文件的路径集合
     */
    protected async removeFileContents(paths: string[]): Promise<void> {
        try {
            await this.data.dataLoaded;
            const keys = paths.map((path) => this.toSheetKey(path)).filter(Boolean);
            if (keys.length === 0) return;

            const removingActived = keys.includes(this.data.activedSheetName);
            let fallback: string | null = null;
            if (removingActived) {
                fallback = this.pickFallbackFile(keys);
                if (fallback) await this.openPath(fallback, false);
                else {
                    // 一个文件都不剩：先把文档工作区关掉，再清空表格
                    await this.docWorkspace?.close();
                    this.workspaceRouter?.showSheet();
                    this.data.clearCurrentSheet();
                }
            }

            // 落库/切换完成后才真正清理存档
            for (const key of keys) {
                await this.settingsDb.deleteValue(key);
                await this.settingsDb.deleteImages(key);
                // 文档仓同键：路径键与 values 库同源，删文件时两库一起清，
                // 否则删掉 .text 后 docs 里会留一份孤儿存档（刷新后可能被对账复活）
                await this.docWorkspace?.docStore.deleteSnapshot(key);
            }

            if (removingActived) {
                this.fileHistory = this.fileHistory.filter((item) => !keys.includes(this.toSheetKey(item)));
                if (fallback) {
                    // 回退目标即当前文件：记忆与树中高亮同步到它
                    this.persistActivedFile(fallback);
                    this.fileTree.showActivedFile(fallback);
                    this.data.messageBox.success(`已删除当前文件，已回到 ${fallback}`);
                } else {
                    // 一个文件都不剩：清掉高亮与「当前文件」记忆。
                    // 此后刷新时树里没有文件 → 按需求补一个空白 sheet1.grid，而不是复活已删的文件。
                    this.fileTree.setActivedFileName('');
                    this.settingsDb.deleteSetting(ACTIVED_FILE_KEY).catch(() => {
                        // 记忆清理失败不影响交互
                    });
                }
            }
        } catch {
            // 存档不可用不影响结构操作
        }
    }

    /**
     * 从打开历史中挑选删除后的回退目标
     *
     * 两层：
     * 1. 打开历史中最近一个**仍在树中**、且不是本次被删除的文件 —— 即「上一个打开过的文件」；
     * 2. 历史上都没有（从未打开过 / 都已被删）时，退回树中剩下的第一个文件。
     *
     * 第二层不是可有可无的兜底：主页面对应的存档键就是当前文件名，若停留在
     * 已删除的文件名上，重绘的自动保存会立刻把这份空存档写回该键（删除等于没删），
     * 刷新后还会按该名字重新登记出一个文件。因此只要树里还有文件，就切到其中一个。
     * @param removing - 本次被删除的存档键集合
     * @returns {string | null} 回退目标的**文件路径**（用于树中高亮与记忆）；树里一个文件都不剩时为 null
     */
    private pickFallbackFile(removing: string[]): string | null {
        // 打开历史里存的就是路径，可直接用于 hasFile 与高亮
        for (const path of this.fileHistory) {
            const key = this.toSheetKey(path);
            if (!key || removing.includes(key)) continue;
            if (this.fileTree?.hasFile(path)) return path;
        }
        for (const path of this.fileTree?.listFiles() ?? []) {
            const key = this.toSheetKey(path);
            if (key && !removing.includes(key)) return path;
        }
        return null;
    }

    /**
     * 读取某文件在数据库中的存档大小（字节），供「另存为」对话框的「文件大小」列使用
     *
     * 存档由两部分组成（与 `values` / `images` 两个库一一对应）：
     * 单元格值（JSON 文本，按 UTF-8 字节数计）与图片（累加二进制长度）。
     * 没有存档时返回 0（视为空文件）；读取异常返回 null，由调用方显示「—」。
     * @param path - 文件路径（含扩展名，相对根节点）
     * @returns {Promise<number | null>} 字节数；读取失败为 null
     */
    protected async getFileSize(path: string): Promise<number | null> {
        const key = this.toSheetKey(path);
        if (!key) return null;
        // 文档：块树的结构化克隆体量，直接按 JSON 文本长度估算（与表格侧同口径）
        if (this.isDocPath(path)) {
            try {
                const snapshot = await this.docWorkspace?.docStore.readSnapshot(key);
                if (!snapshot) return 0;
                return new TextEncoder().encode(JSON.stringify(snapshot)).byteLength;
            } catch {
                return null;
            }
        }
        try {
            const raw = await this.settingsDb.getValue(key);
            let bytes = 0;
            if (typeof raw === 'string') {
                // `'{}'` 是 getValue 对「无记录」的返回值，不应计作文件内容
                bytes = raw === '{}' ? 0 : new TextEncoder().encode(raw).byteLength;
            } else if (raw !== null && raw !== undefined) {
                bytes = new TextEncoder().encode(JSON.stringify(raw)).byteLength;
            }
            const images = await this.settingsDb.getImages(key);
            for (const image of images) {
                const data = (image as { data?: unknown } | null)?.data;
                if (data instanceof Uint8Array) bytes += data.byteLength;
            }
            return bytes;
        } catch {
            return null;
        }
    }

    /**
     * 对账：让数据库里的存档与文件树里的文件一一对应
     *
     * 内容桥已在每次新增 / 改名 / 复制 / 删除的当下同步过数据库，这里只兜底清理
     * **历史遗留**（例如上一版程序留下的、树里已不存在的空存档），因此：
     * - 补建：树里有文件、库里没有存档 → 写入一份空白快照；
     * - 清理：库里有存档、树里已无该文件 → 删除，但**只删空存档**：
     *   非空却无归属的存档宁可留着，也不冒误删内容的风险。
     *
     * 图片记录（独立的 `images` 对象存储）一并按「无归属即删」清理，
     * 但**不补建**：没有图片的文件本来就不该有图片记录（补一条空记录只会白占空间）。
     *
     * 归属判定用**路径键**（{@link FileSessionBase.toSheetKey}）：`归档/销售数据.grid` 与
     * `销售数据.grid` 是两个不同的文件，因此也是两个不同的存档。
     *
     * 先做一次**存量键迁移**：`toArchiveKey` 之前「一律去扩展名」，现在改为
     * 「只去原生 `.grid`、其它后缀保留」。因此历史上从本机打开的 `.text` / `.demo` 等
     * 文件，其存档键从 `报表` 变成 `报表.text`——这里把旧键的存档搬到新键，
     * 再删旧键，避免升级后这些文件看起来「内容空了」。
     */
    protected async reconcileValueStore(): Promise<void> {
        try {
            // ① 存量键迁移（必须在「补建」之前：否则会先给新键写入空白存档，把迁移挤掉）
            for (const path of this.fileTree.listFiles()) {
                const key = this.toSheetKey(path);
                const legacy = stripExtension(path).trim();
                if (!key || !legacy || key === legacy) continue;   // 原生 .grid：键未变
                const moved = await this.settingsDb.getValue(legacy);
                if (!moved || moved === '{}') continue;            // 旧键无内容，无需迁移
                const current = await this.settingsDb.getValue(key);
                if (current && current !== '{}') continue;         // 新键已有内容：不覆盖
                await this.settingsDb.setValue(key, moved);
                await this.settingsDb.deleteValue(legacy);
            }

            const owned = new Set<string>();
            for (const path of this.fileTree.listFiles()) {
                const key = this.toSheetKey(path);
                if (key) owned.add(key.toLowerCase());
            }
            const records = await this.settingsDb.getAllValues();
            const existing = new Set(records.map((record) => String(record.sheetName ?? '').toLowerCase()));

            // 补建：保证每个文件都有存档
            for (const path of this.fileTree.listFiles()) {
                const key = this.toSheetKey(path);
                if (key && !existing.has(key.toLowerCase())) {
                    await this.settingsDb.setValue(key, this.data.blankSnapshot());
                }
            }

            // 清理：删掉不再属于任何文件的空存档
            for (const record of records) {
                const key = String(record.sheetName ?? '');
                if (!key || owned.has(key.toLowerCase())) continue;
                if (!this.isEmptySnapshot(record.value)) continue;
                await this.settingsDb.deleteValue(key);
            }

            // 图片记录：无归属的直接删（与「空存档才删」不同 —— 图片是可再生内容，
            // 留着既占内存又永远不会被读，没有「宁可留着也别误删」的权衡）。
            // 放在最后：读键失败时不应影响上面已经完成的值对账。
            const imageKeys = await this.settingsDb.getAllImageKeys();
            for (const key of imageKeys) {
                if (!key || owned.has(key.toLowerCase())) continue;
                await this.settingsDb.deleteImages(key);
            }

            // 文档仓同口径对账：**只删空文档**，非空的一律留着（宁可留孤儿，也不误删内容）
            await this.docWorkspace?.docStore.pruneOrphans(owned);
        } catch {
            // 存档不可用（IndexDB 被禁用等）时跳过对账，不影响使用
        }
    }

    /**
     * 判断存档是否为「空表」（没有单元格内容）
     *
     * 两态兼容：
     * - v4 对象存档 → 没有分片（或分片全空）即空表；
     * - v3 及更早的 JSON 字符串 → 解析后 `values` 为空数组即空表。
     *
     * 解析失败一律按「非空」处理：无法确认内容的一律不删，避免误删数据。
     * @param raw - 存档原始值（v4 对象 / JSON 字符串 / 未知形态）
     * @returns {boolean} 是否为空表
     */
    private isEmptySnapshot(raw: unknown): boolean {
        if (DataCollection.isV4Snapshot(raw)) {
            return raw.chunks.every((chunk) => typeof chunk !== 'string' || chunk.length === 0 || chunk === '[]');
        }
        if (typeof raw !== 'string') return false;
        try {
            const parsed = JSON.parse(raw);
            return Array.isArray(parsed?.values) ? parsed.values.length === 0 : true;
        } catch {
            return false;
        }
    }
}
