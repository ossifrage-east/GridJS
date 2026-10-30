/*  # 工作区职责拆分（K4）
    # 从 app.ts 原样迁出，仅调整类继承关系与成员可见性，逻辑零改动
*/
import { extractExtension, stripExtension } from '../../utils/fileName';
import type { DocSnapshot } from '../document/blockTree';
import { FileSessionBase } from './fileSession';

/**
 * 判定一个文件是否由**文档工作区**承载
 *
 * 应用只有三种原生后缀（见 `utils/fileName` 的 `NATIVE_EXTENSIONS`）：
 * `.grid` 表格、`.text` 文档、`.demo` 演示文稿。
 * `.md` 也按文档打开 —— 它本来就是 Markdown 源码，落在文档工作区的源码视图里正合适。
 * @param path - 文件路径（含扩展名）
 * @returns 是否走文档工作区
 */
export function isDocPath(path: string): boolean {
    const extension = extractExtension(path);
    return extension === 'text' || extension === 'md';
}

/**
 * 文件打开分派基类：按后缀把文件路由到表格 / 文档工作区
 *
 * 只用 `WorkspaceRouter` 与 `docWorkspace.docStore` 这两个既有入口，不认识具体实现。
 */
export abstract class OpenRouterBase extends FileSessionBase {
    /** 补建文件存档（由 `contentBridge.ts` 实现，打开空白文件时会用到） */
    protected abstract ensureFileContents(paths: string[]): Promise<void>;

    /**
     * 判定文件是否由文档工作区承载（模块级函数的类内入口）
     * @param path - 文件路径（含扩展名）
     * @returns 是否走文档工作区
     */
    protected isDocPath(path: string): boolean {
        return isDocPath(path);
    }

    /**
     * 取路径中的文件显示名（末段）
     * @param path - 文件路径
     * @returns {string} 显示名；空路径返回空串
     */
    protected baseNameOf(path: string): string {
        const text = (path ?? '').trim();
        const slash = text.lastIndexOf('/');
        return slash >= 0 ? text.slice(slash + 1) : text;
    }

    /**
     * 取文档标题（末段去后缀）
     * @param path - 文件路径
     * @returns 标题
     */
    protected docTitleOf(path: string): string {
        return stripExtension(this.baseNameOf(path));
    }

    /**
     * 双击树节点：把该文件内容装载到主页面（按后缀路由到对应工作区）
     *
     * `isDocPath` 为真走文档工作区，否则走表格工作区。两条路都保证：
     * 内容装载 → 切换工作区 → 树中高亮 → 记忆「当前文件」→ 记入打开历史。
     * @param path - 文件路径（含扩展名，相对根节点）
     */
    protected async activateFile(path: string): Promise<void> {
        await this.openPath(path, true);
    }

    /**
     * 按后缀打开文件（唯一的文件打开入口）
     * @param path - 文件路径（含扩展名）
     * @param announce - 是否弹「已打开」提示（刷新恢复时不弹）
     */
    protected async openPath(path: string, announce: boolean): Promise<void> {
        if (!path) return;
        if (isDocPath(path)) {
            await this.openDocFile(path, announce);
            return;
        }
        // `.demo` 没有工作区：明确降级为表格模式并提示，不静默
        if (extractExtension(path) === 'demo') {
            if (announce) this.data.messageBox.info('演示文稿暂不支持编辑，已按表格打开');
        }
        await this.openSheetFile(path, announce);
    }

    /**
     * 在表格工作区打开文件
     * @param path - 文件路径
     * @param announce - 是否弹提示
     */
    protected async openSheetFile(path: string, announce: boolean): Promise<void> {
        const sheetName = this.toSheetKey(path);
        if (!sheetName) return;

        // 离开文档前先落库：文档的自动保存是防抖的，切走时窗口内可能还有改动
        await this.docWorkspace?.flush();
        this.workspaceRouter?.showSheet();

        const switched = await this.data.switchToSheet(sheetName);
        if (!switched) {
            // 已在主页面显示当前文件 → 内容无变化（但工作区可能刚从文档切回来）
            this.fileTree.showActivedFile(path);
            return;
        }

        this.fileTree.showActivedFile(path);
        this.persistActivedFile(path);
        this.rememberOpenedFile(path);
        if (announce) this.data.messageBox.success(`已打开 ${this.baseNameOf(path)}`);
    }

    /**
     * 在文档工作区打开文件
     *
     * 与表格侧的差别是「同文件不重载」：`DocStore.switchTo` 在键相同时返回 false，
     * 此时只切工作区，不重建 DOM —— 重复双击同一个文档不会丢光标与撤销栈。
     * @param path - 文件路径
     * @param announce - 是否弹提示
     */
    protected async openDocFile(path: string, announce: boolean): Promise<void> {
        const key = this.toSheetKey(path);
        if (!key) return;
        this.workspaceRouter?.showDoc();
        await this.docWorkspace.open(key, this.docTitleOf(path));
        this.fileTree.showActivedFile(path);
        this.persistActivedFile(path);
        this.rememberOpenedFile(path);
        if (announce) this.data.messageBox.success(`已打开 ${this.baseNameOf(path)}`);
    }

    /**
     * 装载「刚从本机打开」的文档
     *
     * 先落库再装载：`docs` 仓是文档的唯一来源，这条顺序让「导入」与「刷新后恢复」
     * 走完全相同的读取路径，避免导入的内容只活在内存里、刷新即失。
     * @param fileName - 原生文件名（如 `说明.text`）
     * @param snapshot - 解析出的存档
     */
    protected async openParsedDocument(fileName: string, snapshot: DocSnapshot): Promise<void> {
        const key = this.toSheetKey(fileName);
        if (!key) return;
        const store = this.docWorkspace.docStore;
        await store.writeSnapshot(key, snapshot);
        this.workspaceRouter.showDoc();
        await this.docWorkspace.open(key, this.docTitleOf(fileName));
        this.docWorkspace.render();
    }

    /**
     * 打开一个空白 Excel 文件，并把它登记到文件树中
     *
     * 用于「刷新后树里一个文件都没有」：默认工作表（`sheet1`）本身就是这个文件的
     * 空表，因此同名时不切换、主页面保持空白；被占用时（如同名文件夹）按命名规则
     * 取 `sheet2.grid` 并切换过去。
     */
    protected async openBlankFile(): Promise<void> {
        const displayName = this.fileTree.nextFileName();
        const sheetName = this.toSheetKey(displayName);
        if (sheetName && sheetName !== this.data.activedSheetName) {
            await this.data.switchToSheet(sheetName);
        }
        this.fileTree.showActivedFile(displayName);
        this.persistActivedFile(displayName);
        this.rememberOpenedFile(displayName);
        // 新增即落库：库里补一份空白存档，使「树里有文件」与「库里有存档」一一对应
        await this.ensureFileContents([displayName]);
        await this.reconcileValueStore();
    }
}
