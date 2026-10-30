/*  # 工作区职责拆分（K4）
    # 从 app.ts 原样迁出，仅调整类继承关系与成员可见性，逻辑零改动
*/
import { DataCollection } from '../dataArchitecture/dataCollection';
import { IndexDB } from '../dataArchitecture/indexDB';
import { Split } from '../common/split';
import { Menu } from '../common/menu';
import { Navigator } from '../common/navigator';
import { Slider } from '../common/slider';
import { FileTree, FileNode } from '../common/fileTree/fileTree';
import { FileOpener } from '../common/fileTree/fileOpener';
import { PYEditor } from '../assistant/pyEditor';
import { Sheets } from '../sheets/sheets';
import { DocWorkspace } from '../document/docWorkspace';
import { WorkspaceRouter } from './workspaceRouter';
import { toArchiveKey } from '../../utils/fileName';

/**
 * 当前打开文件在 IndexDB 中的设置键
 *
 * 值 = 文件**路径**（含扩展名，如 `归档/销售数据.grid`）：刷新后据此决定主页面显示哪个文件。
 * 存路径而不是文件名：不同目录下的同名文件是两个独立文件，只存名字无法区分该恢复哪一个。
 * 它只是「上次看的是哪一个」的记忆，**文件是否存在以文件树为准**
 * （见 {@link FileSessionBase.restoreActivedFile}），因此记忆失效（文件被删 / 被改名）时
 * 会自动退回树里的第一个文件，而不是凭记忆造出一个文件。
 */
export const ACTIVED_FILE_KEY = 'actived-file';

/**
 * 文件树结构在 IndexDB 中的设置键
 *
 * 值 = `FileNode[]`（含每个文件夹的展开状态）：树发生新增 / 删除 / 重命名 /
 * 粘贴时整体覆写一次，刷新后据此重建，使树里保持刷新前的文件与目录。
 */
export const FILE_TREE_KEY = 'file-tree';

/**
 * 「打开历史」的保留长度
 *
 * 删除当前打开的文件时要回退到**上一个打开过的文件**，因此需要记住最近打开顺序。
 * 只保留最近的若干项即可：历史项还可能因文件被删除而失效，需要时向后继续找。
 */
export const FILE_HISTORY_LIMIT = 20;

/**
 * 文件会话基类：共用状态 + 文件树持久化与活动文件记忆
 *
 * 位于 workspace 继承链最底层（`App → ExportActions → ContentBridge → OpenRouter → FileSession`），
 * 各职责组共用的字段统一声明在这里，派生类直接以 `this.xxx` 取用。
 * 需要派生类实现的三处能力用 `abstract` 声明，避免基类反向依赖具体职责组。
 */
export abstract class FileSessionBase {
    /* ---------- 由派生类（App）装配的共用状态 ---------- */
    /** 表格数据层（剪贴板 / 单元格模型 / 存档落库） */
    protected data: DataCollection = new DataCollection();
    /** 界面偏好的持久化存储（导航内容区宽度、当前打开的文件等设置项共用） */
    protected readonly settingsDb: IndexDB = new IndexDB();
    /** 主分栏（左：菜单 + 工具栏 + 工作区；右：导航内容区） */
    protected containerSplit!: Split;
    /** 右侧工具线（导航面板与表格缩放滑块的宿主） */
    protected rightToolsSplit!: Split;
    protected menu!: Menu;
    protected navigator!: Navigator;
    /** 右侧工具条底部的缩放滑块（与 Ctrl+滚轮 缩放双向绑定） */
    protected zoomSlider!: Slider;
    /** 表格工作区 */
    protected sheets!: Sheets;
    protected fileTree!: FileTree;
    protected fileOpener!: FileOpener;
    protected pyEditor!: PYEditor;
    /**
     * 文档工作区（块树 + DOM 视图 + Markdown 源码视图）
     *
     * 与表格工作区并列存在、靠显隐切换：两者都在 DOM 里，切换时不重建，
     * 因此「切过去再切回来」不会丢掉未落库的编辑状态与撤销栈。
     */
    protected docWorkspace!: DocWorkspace;
    /** 工作区路由：显隐 + 命令总线换绑，切换副作用全部收在里面 */
    protected workspaceRouter!: WorkspaceRouter;
    /**
     * 打开历史（文件**路径**，越靠前越近）
     *
     * 删除「当前打开的文件」时，据此回退到上一个仍在树中的文件。
     */
    protected fileHistory: string[] = [];
    /** 当前显示的导航项文本：拖拽保存宽度时作为键 */
    protected currentNavItemText = '';

    /* ---------- 由更上层（打开分派 / 内容桥）实现的抽象成员 ---------- */
    /** 按后缀打开文件（唯一的文件打开入口，见 `openRouter.ts`） */
    protected abstract openPath(path: string, announce: boolean): Promise<void>;
    /** 打开一个空白 Excel 文件（见 `openRouter.ts`） */
    protected abstract openBlankFile(): Promise<void>;
    /** 让数据库里的存档与文件树里的文件一一对应（见 `contentBridge.ts`） */
    protected abstract reconcileValueStore(): Promise<void>;

    /**
     * 把文件**路径**换算为内容存档键（IndexedDB `values` / `images` 库的键）
     *
     * 规则见 {@link toArchiveKey}：只去掉应用原生表格后缀（`.grid`），其它后缀保留。
     * - `归档/销售数据.grid` → `归档/销售数据`
     * - `报表.grid` → `报表`、`报表.text` → `报表.text`（同名不同后缀互不干扰）
     *
     * **必须只由路径决定**：早期版本按「树里当前是否有同基名冲突」来决定是否保留后缀，
     * 但删除同名的兄弟文件后冲突消失、键立刻漂移（`报表.grid` → `报表`），
     * 致使「打开时写入的键」与「删除/回退时读的键」不一致 —— 删不掉、回退到空白。
     * 存档键是路径的纯函数，才能保证任何时刻都指向同一份存档。
     *
     * 存量兼容：`.grid` 的键与历史 `.xlsx` 完全一致（`sheet1.grid` / `sheet1.xlsx` → `sheet1`），无需迁移；
     * 历史遗留的其它后缀存档（键为去扩展名）由 {@link FileSessionBase.reconcileValueStore} 一次性迁移。
     * @param path - 文件路径（含扩展名，相对根节点）
     * @returns {string} 存档键；名称为空时返回空串
     */
    protected toSheetKey(path: string): string {
        return toArchiveKey(path);
    }

    /**
     * 读取并规整「上次的文件树结构」（页面刷新后还原）
     *
     * 存档为 {@link FileNode} 数组；IndexDB 不可用、存档缺失或结构损坏时返回 null，
     * 由调用方退回内置默认树（仅「工作区」）。
     * @returns {Promise<FileNode[] | null>} 树数据；无有效存档时为 null
     */
    protected async loadFileTree(): Promise<FileNode[] | null> {
        try {
            return this.sanitizeFileNodes(await this.settingsDb.getSetting(FILE_TREE_KEY));
        } catch {
            return null;
        }
    }

    /**
     * 校验并规整文件树存档，避免损坏的数据把建树流程带崩
     *
     * 可接受的节点形状：
     * `{ name: string, path?: string, children?: FileNode[], open?: boolean,
     *    createdAt?: number, lastModifiedAt?: number, createdBy?: string,
     *    lastModifiedBy?: string, editable?: boolean }`。
     * 匿名节点与非法项一律丢弃；`children` 非数组时视为文件。
     *
     * 字段类型不符时**丢弃该字段**而不是丢弃节点（缺时间戳 / 缺创建人只影响「属性」
     * 对话框的一行，不该让整个文件消失）；`path` 缺失时保持缺省，由 {@link FileTree}
     * 建树阶段按父链重算，因此旧版存档（只有 `name/children/open`）依然可用。
     * @param input - 从 IndexDB 取出的原始值
     * @returns {FileNode[] | null} 规整后的树数据；无可保留节点时为 null
     */
    private sanitizeFileNodes(input: unknown): FileNode[] | null {
        if (!Array.isArray(input)) return null;

        const walk = (list: unknown[]): FileNode[] => {
            const nodes: FileNode[] = [];
            for (const raw of list) {
                if (!raw || typeof raw !== 'object') continue;
                const source = raw as Record<string, unknown>;
                const text = typeof source.name === 'string' ? source.name.trim() : '';
                if (!text) continue;
                const node: FileNode = { name: text };
                // 路径：非空字符串才接受（空串等价于「缺失」，同样由建树阶段重算）
                if (typeof source.path === 'string' && source.path.trim()) node.path = source.path.trim();
                if (typeof source.open === 'boolean') node.open = source.open;
                if (typeof source.createdAt === 'number' && Number.isFinite(source.createdAt)) {
                    node.createdAt = source.createdAt;
                }
                if (typeof source.lastModifiedAt === 'number' && Number.isFinite(source.lastModifiedAt)) {
                    node.lastModifiedAt = source.lastModifiedAt;
                }
                if (typeof source.createdBy === 'string' && source.createdBy) node.createdBy = source.createdBy;
                if (typeof source.lastModifiedBy === 'string' && source.lastModifiedBy) {
                    node.lastModifiedBy = source.lastModifiedBy;
                }
                if (typeof source.editable === 'boolean') node.editable = source.editable;
                if (Array.isArray(source.children)) node.children = walk(source.children);
                nodes.push(node);
            }
            return nodes;
        };

        const nodes = walk(input);
        return nodes.length > 0 ? nodes : null;
    }

    /**
     * 持久化文件树结构（树结构或名称变化时整体覆写）
     * @param nodes - 当前整棵树（{@link FileTree.getNodes} 导出的纯数据）
     */
    protected persistFileTree(nodes: FileNode[]): void {
        this.settingsDb.setSetting(FILE_TREE_KEY, nodes, 'file-tree').catch(() => {
            // 持久化失败不影响交互
        });
    }

    /**
     * 恢复「上次打开的文件」（页面刷新后）
     *
     * **文件树是文件的唯一来源**，记忆只是「上次看的是哪一个」：
     * - 树里一个文件都没有（首次使用，或上次把文件删到一个不剩）→ 打开一个空白
     *   Excel 并把 `sheet1.grid` 登记进树（见 {@link FileSessionBase.openBlankFile}）；
     * - 记忆里的文件仍在树中 → 恢复它：装载其内容、登记并高亮；
     * - 记忆缺失，或所指文件已不在树中（被删 / 被改名）→ 退回树里的第一个文件，
     *   而不是凭记忆造出一个树里没有的文件。
     *
     * 主页面的初始装载固定读取默认工作表（`sheet1`），因此必须等
     * {@link DataCollection.dataLoaded} 之后再切换，否则两次装载会互相覆盖。
     */
    protected async restoreActivedFile(): Promise<void> {
        try {
            await this.data.dataLoaded;

            // 树里一个文件都没有：给用户一个可直接编辑的空白文件（需求：补 sheet1.grid）
            const files = this.fileTree.listFiles();
            if (files.length === 0) {
                await this.openBlankFile();
                return;
            }

            const saved = await this.settingsDb.getSetting(ACTIVED_FILE_KEY);
            const remembered = typeof saved === 'string' ? saved.trim() : '';
            // files 是**路径**集合；记忆里存的也是路径，两者可直接比较
            const displayName = remembered && this.fileTree.hasFile(remembered) ? remembered : files[0];

            // 先对账（含存量键迁移与补建）再装载：否则历史遗留键的文件会被读成空表
            await this.reconcileValueStore();

            // 上次打开的是别的文件：按后缀路由到对应工作区装载（无存档时按空文件处理）。
            // 刷新恢复不弹「已打开」提示，否则每次刷新都会闪一条消息。
            await this.openPath(displayName, false);
            this.fileTree.showActivedFile(displayName);
            this.persistActivedFile(displayName);
            this.rememberOpenedFile(displayName);
        } catch {
            // 恢复失败（IndexDB 不可用等）不影响后续使用：此时树中不含当前文件
        }
    }

    /**
     * 记忆当前打开的文件路径（刷新后据此决定主页面显示哪个文件）
     *
     * 存**路径**而不是文件名：不同目录的同名文件是两个文件，
     * 只存名字无法在刷新后区分该恢复哪一个。
     * 路径为空时不写入：文件是否存在以树为准。
     * @param path - 文件路径（含扩展名，相对根节点）
     */
    protected persistActivedFile(path: string): void {
        const name = (path ?? '').trim();
        if (!name) return;
        this.settingsDb.setSetting(ACTIVED_FILE_KEY, name, 'actived-file').catch(() => {
            // 持久化失败不影响交互
        });
    }

    /**
     * 记入「打开历史」（最近打开的排在最前，按存档键去重）
     * @param path - 文件路径（含扩展名，相对根节点）
     */
    protected rememberOpenedFile(path: string): void {
        const name = (path ?? '').trim();
        if (!name) return;
        const key = this.toSheetKey(name);
        this.fileHistory = [name, ...this.fileHistory.filter((item) => this.toSheetKey(item) !== key)];
        if (this.fileHistory.length > FILE_HISTORY_LIMIT) {
            this.fileHistory.length = FILE_HISTORY_LIMIT;
        }
    }

    /**
     * 同步「打开历史 / 当前文件记忆」中的文件路径（改名 / 移动后调用）
     * @param from - 旧路径
     * @param to - 新路径
     */
    protected rememberRenamedFile(from: string, to: string): void {
        const key = this.toSheetKey(from);
        if (!key) return;
        this.fileHistory = this.fileHistory.map((item) => (this.toSheetKey(item) === key ? to : item));
        this.settingsDb.getSetting(ACTIVED_FILE_KEY).then((saved: unknown) => {
            if (typeof saved === 'string' && this.toSheetKey(saved) === key) this.persistActivedFile(to);
        }).catch(() => {
            // 读不到记忆时忽略：下一次打开文件会重新写入
        });
    }
}
