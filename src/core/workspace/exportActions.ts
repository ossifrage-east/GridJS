/*  # 工作区职责拆分（K4）
    # 从 app.ts 原样迁出，仅调整类继承关系与成员可见性，逻辑零改动
*/
import { stripExtension } from '../../utils/fileName';
import { createXlsxTask, type XlsxTask } from '../../utils/parser/xlsxWriter';
import type { DocSnapshot } from '../document/blockTree';
import { createBlankSnapshot } from '../document/blockTree';
import { snapshotToDocx } from '../document/converter/docxCodec';
import { FileTree, FileNode, FileMovePlan } from '../common/fileTree/fileTree';
import { FileOpener } from '../common/fileTree/fileOpener';
import { Navigator } from '../common/navigator';
import { DataEvents, NAV_PANEL } from '../constant';
import { ContentBridgeBase } from './contentBridge';

/**
 * 文件域装配基类：把文件树接到「打开分派 / 内容桥 / 导出 / 当前文件记忆」上
 *
 * 位于 workspace 继承链顶端（`App → ExportActions → ContentBridge → OpenRouter → FileSession`），
 * 因此下层能力在这里全部是既成事实，无需再声明抽象成员。
 * 导出动作本身也在本层：它们是文件树的两个选项（`convertXlsx` / `convertDocx`）。
 */
export abstract class ExportActionsBase extends ContentBridgeBase {
    /**
     * 初始化「文件」页签：树形文件页面
     *
     * 顶部为搜索栏（同行右侧为展开 / 收缩 / 定位按钮），下方为树形文件列表；
     * 新建 / 打开 / 剪切粘贴等操作统一通过树节点右键菜单完成。
     *
     * 树结构本身也持久化：建树前先读取上次的树（{@link ExportActionsBase.loadFileTree}），
     * 之后任何结构变化由 {@link ExportActionsBase.persistFileTree} 整体覆写，
     * 因此刷新页面后树里仍是刷新前的那些文件与文件夹。
     * @param navigator - 右侧导航器实例（内容项挂到「文件」页签下）
     */
    protected async initFileTree(navigator: Navigator): Promise<void> {
        // 文件打开器：负责「打开」按钮与主页面拖放，解析成功后把文件登记到文件树
        this.fileOpener = new FileOpener({
            data: this.data,
            // 同名冲突去重：不同原始后缀换算后可能得到同一原生名
            // （`test_sample.xls` 与 `test_sample.xlsx` 都 → `test_sample.grid`），
            // 由文件树在根目录内判冲突并给出加序号后的名称（`test_sample1.grid`），
            // 保证每个打开的文件都有独立的树节点与存档，不会互相覆盖
            resolveName: (nativeName: string) => this.fileTree.resolveOpenedFileName(nativeName),
            // 从本机打开的文件只有文件名（浏览器不暴露所在目录），因此它是
            // 根目录下的文件 —— 与「路径」语义天然一致（根目录下路径 ≡ 文件名）
            onOpened: (fileName: string) => {
                // 打开的文件即「当前文件」：登记到树中并高亮，刷新后据此恢复
                this.fileTree.showActivedFile(fileName);
                this.persistActivedFile(fileName);
                this.rememberOpenedFile(fileName);
                this.data.messageBox.success(`已打开 ${fileName}`);
            },
            onError: (message: string) => this.data.messageBox.error(message),
            // 文档类文件（docx / doc / md / txt）落到文档工作区：解析结果是块树，
            // 先写进 docs 仓再装载，使「刚打开的文件」与「刷新后恢复的文件」走同一条路
            onDocument: (fileName: string, snapshot: DocSnapshot) => this.openParsedDocument(fileName, snapshot),
        });
        document.body.appendChild(this.fileOpener.getFileInput());

        // 刷新还原：先取出上次的树结构再建树（读不到则用内置默认树「工作区」）
        const savedNodes = await this.loadFileTree();

        this.fileTree = new FileTree({
            menu: this.menu,
            nodes: savedNodes ?? undefined,
            onOpen: () => this.fileOpener.pickAndOpen(),
            // 双击树节点：切换到该文件（已在主页面显示时内容不变，见 activateFile）
            onActivate: (path: string) => this.activateFile(path),
            // 树结构/名称/路径变化即落库：下次刷新据此还原「刷新前的文件」
            onChange: (nodes: FileNode[]) => this.persistFileTree(nodes),
            // 内容桥：新增 / 复制 / 改名 / 移动 / 删除文件时同步数据库中该文件的内容。
            // 参数一律是**路径**（含扩展名、相对根节点）：存档键由 toSheetKey 换算
            // （只去原生 `.grid` 后缀，其它后缀保留，见 toArchiveKey），
            // 因此不同目录的同名文件、同目录的同名不同后缀文件，各自拥有独立存档。
            content: {
                create: (paths: string[]) => this.ensureFileContents(paths),
                copy: (plans: FileMovePlan[]) => this.copyFileContents(plans),
                move: (plans: FileMovePlan[]) => this.moveFileContents(plans),
                remove: (paths: string[]) => this.removeFileContents(paths),
            },
            // 「另存为」对话框的「文件大小」列：文件内容存在 IndexedDB，按路径换算存档键后读取
            fileSize: (path: string) => this.getFileSize(path),
            // 「另存为xlsx」的真实转换：读存档 → 生成 xlsx Blob（下载 / 打包下载直接写盘）
            convertXlsx: (path: string, name: string) => this.convertGridToXlsx(path, name),
            // 「另存为docx」的真实转换：读 docs 仓的块树 → 编成 docx Blob
            convertDocx: (path: string) => this.convertTextToDocx(path),
        });
        navigator.navContentItemMount(NAV_PANEL[0], this.fileTree.getElements().container);   // NAV_PANEL[0] 即「文件」
        // 「定位」的目标是「主页面正在显示的那个文件」。
        // `activedSheetName` 是**存档键**（`.grid` 去后缀、其它后缀保留），
        // 树的定位匹配（{@link FileTree.findByName}）既能按路径精确命中带后缀的键，
        // 也支持「无后缀查询 → 路径去扩展名」的兜底，因此两种形态都能对上；
        // 从本机打开的单个文件没有目录，键与路径名相同，同样能定位到。
        this.fileTree.setActivedFileName(this.data.activedSheetName);
        this.data.on(DataEvents.ACTIVED_SHEET_NAME_CHANGED, (_, name: string) => {
            this.fileTree?.setActivedFileName(name);
        });

        // 刷新恢复：等主页面的初始装载完成后再恢复「上次打开的文件」
        this.restoreActivedFile();

        // 主页面拖放：拖入 xlsx/xls/docx/doc/pptx/ppt 即自动打开
        this.fileOpener.bindDropZone(this.sheets.getElements().sheets);
    }

    /**
     * 创建 `.grid` 文件 → xlsx 的转换作业（「另存为xlsx」对话框的真实转换）
     *
     * 存档键与文件大小列同源（见 {@link ExportActionsBase.toSheetKey}）：当前主页面正在编辑的文件
     * 直接取内存快照（含尚未触发自动保存的改动），其余文件读 IndexedDB 里的存档。
     * 返回的作业由对话框按时间片驱动（单飞队列），因此这里只负责取数据，不做转换。
     * @param path - 文件路径（含扩展名，相对根节点）
     * @param name - 文件名（用于工作表名）
     * @returns {Promise<XlsxTask>} 可暂停 / 可续传的转换作业
     */
    protected async convertGridToXlsx(path: string, name: string): Promise<XlsxTask> {
        const key = this.toSheetKey(path);
        if (!key) throw new Error('无效的文件路径');
        const raw = key === this.data.activedSheetName
            ? this.data.toSnapshot()
            : await this.settingsDb.getValue(key);
        return createXlsxTask(raw, stripExtension(name));
    }

    /**
     * 创建 `.text` 文档 → docx 的转换作业（「另存为docx」对话框的真实转换）
     *
     * 与 {@link ExportActionsBase.convertGridToXlsx} 同一套路：树只给路径，装配层按路径取内容。
     * 差别是文档的「内容」不在 `values` 仓而在 `docs` 仓，且**当前打开的文档**
     * 可能还停在防抖窗口里（`DocStore` 500ms 才落库），因此必须先 `flush()`
     * 再读，否则读到的是上一次的内容。
     *
     * docx 编码本身是一次同步动作，这里把它包成与 xlsx 同形的作业：
     * 首次 `step()` 才真正编码，对话框照常按时间片驱动、显示「已完成」并允许下载。
     * @param path - 文件路径（含扩展名，相对根节点）
     * @returns {Promise<XlsxTask>} 可暂停 / 可续传的转换作业
     */
    protected async convertTextToDocx(path: string): Promise<XlsxTask> {
        const key = this.toSheetKey(path);
        if (!key) throw new Error('无效的文件路径');
        const store = this.docWorkspace?.docStore;
        if (!store) throw new Error('文档功能尚未就绪');
        await store.flush();

        const snapshot = await store.readSnapshot(key);
        // 没有存档 = 这个文件从来没被编辑过（空文档）：仍导出一份空白 Word，
        // 比「点了没反应」更符合「导出这个文件」的预期
        const source = snapshot ?? createBlankSnapshot(this.docTitleOf(path));
        let blob: Blob | null = null;
        return {
            get progress(): number {
                return blob ? 100 : 0;
            },
            get blob(): Blob | null {
                return blob;
            },
            step: (): boolean => {
                if (!blob) blob = snapshotToDocx(source);
                return true;
            },
        };
    }
}
