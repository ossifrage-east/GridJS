/*  # 表格接口 TS 文件
    # 项目名称: 2023年1月1日起正式更名为 silkspaces
    # 项目地址: 
    # 项目描述: 基于 HTML5 的 Excel 表格组件，无需服务器端支持，即可在浏览器中创建、读取、写入 Excel 文件。
    # 项目目标: 提供一个简单、易用、功能完善的 Excel 表格组件，满足用户在浏览器中处理 Excel 文件的需求。
    # 项目状态: 开发中
    # 项目维护者: 东方鹗
    # 项目许可证: MIT 许可证
    # 项目版本: 0.0.1
    # 项目开始日期: 2023年1月1日
    # 吐槽: 中国的教育是在吃人，学生的学习进度是被控制的，学生的学习能力是被限制的，学生的学习兴趣是被抑制的。
    # 一个简单的课堂纪律，每天都得由学生填报，小到椅子没有摆放整齐，试想一下，一个人天天这样被控制，是什么心态，
    # 而且老师还会通过即时通讯工具，发给每一个学生家长。无视孩子的青春叛逆期，徒增学生与家长的矛盾，这是中国教育的一个严重问题。
    # 而老师的心中只是家教不好。
    # 或许是错怪了老师，当学生的学习进度被控制时，老师的压力也会增加。教材是一部分专家编辑的，考试卷又是另一部分专家写的，课外辅导教程又是另一部分人，
    # 于是学生的学习成了一部分“专家”的炫耀技巧的场所，这些人其心可诛。
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { TOOLS_CLASS_NAMES, DEFAULT_CELL_HEIGHT, ROW_HEADER_PADDING,
    NAV_PANEL,
    NAV_CONTENT_WIDTH,
    DEFAULT_FONT_FAMILY,
    DEFAULT_FONT_SIZE,
    FONT_FAMILY_LIST,
    DataEvents
} from "./constant";
import { Split } from "./common/split";
import { Navigator } from "./common/navigator";
import { FileTree, FileNode, FileMovePlan } from "./common/fileTree";
import { FileOpener } from "./common/fileOpener";
import { Sheets } from "./sheets/sheets";
import { DataCollection } from "./dataArchitecture/dataCollection";
import { IndexDB } from "./dataArchitecture/indexDB";
import { Scroller } from "./common/scroller";
import { Slider } from "./common/slider";
import { Toolbar } from "./toolbar/toolbar";
import { Menu } from "./common/menu";
import { QuickAccess } from "./toolbar/quickAccess";
import { stripExtension } from "../utils/fileName";
import favicon from "../assets/images/favicon.ico";

/** 导航内容区宽度在 IndexDB 中的键前缀（完整键 = 前缀 + 导航项文本） */
const NAV_WIDTH_KEY_PREFIX = 'nav-content-width:';

/**
 * 当前打开文件在 IndexDB 中的设置键
 *
 * 值 = 文件**路径**（含扩展名，如 `归档/销售数据.xlsx`）：刷新后据此决定主页面显示哪个文件。
 * 存路径而不是文件名：不同目录下的同名文件是两个独立文件，只存名字无法区分该恢复哪一个。
 * 它只是「上次看的是哪一个」的记忆，**文件是否存在以文件树为准**
 * （见 {@link App.restoreActivedFile}），因此记忆失效（文件被删 / 被改名）时
 * 会自动退回树里的第一个文件，而不是凭记忆造出一个文件。
 */
const ACTIVED_FILE_KEY = 'actived-file';

/**
 * 文件树结构在 IndexDB 中的设置键
 *
 * 值 = `FileNode[]`（含每个文件夹的展开状态）：树发生新增 / 删除 / 重命名 /
 * 粘贴时整体覆写一次，刷新后据此重建，使树里保持刷新前的文件与目录。
 */
const FILE_TREE_KEY = 'file-tree';

/**
 * 「打开历史」的保留长度
 *
 * 删除当前打开的文件时要回退到**上一个打开过的文件**，因此需要记住最近打开顺序。
 * 只保留最近的若干项即可：历史项还可能因文件被删除而失效，需要时向后继续找。
 */
const FILE_HISTORY_LIMIT = 20;

/** 缩放滑块取值下限（百分比）：与 Ctrl+滚轮 缩放的下限同源 */
const ZOOM_SLIDER_MIN = Math.round(Scroller.ZOOM_MIN * 100);
/** 缩放滑块取值上限（百分比）：与 Ctrl+滚轮 缩放的上限同源 */
const ZOOM_SLIDER_MAX = Math.round(Scroller.ZOOM_MAX * 100);


export class App {
    private worktop!: HTMLDivElement;
    private data: DataCollection = new DataCollection();
    private containerSplit!: Split;
    private sheets!: Sheets;
    private bottomToolsSplit!: Split;
    private rightToolsSplit!: Split;
    private hScroller!: Scroller;
    private vScroller!: Scroller;
    private toolbar!: Toolbar;
    private menu!: Menu;
    private navigator!: Navigator;
    private fileTree!: FileTree;
    private fileOpener!: FileOpener;

    /**
     * 打开历史（文件**路径**，越靠前越近）
     *
     * 删除「当前打开的文件」时，据此回退到上一个仍在树中的文件。
     */
    private fileHistory: string[] = [];
    /** 右侧工具条底部的缩放滑块（与 Ctrl+滚轮 缩放双向绑定） */
    private zoomSlider!: Slider;
    /** 界面偏好的持久化存储（导航内容区宽度、当前打开的文件等设置项共用） */
    private readonly settingsDb: IndexDB = new IndexDB();
    /** 当前显示的导航项文本：拖拽保存宽度时作为键 */
    private currentNavItemText = '';

    /**
     * 初始化工具栏
     */
    constructor(id: string) {
        const link = document.createElement('link');
        link.rel = 'icon';
        link.type = 'image/x-icon';
        link.href = favicon;
        document.head.appendChild(link);
        this.initUI(id);
    }
    /**
     * 初始化应用
     */
    private initUI(id: string) {
        const worktop = document.querySelector(`${id}`);
        if (!worktop) {
            throw new Error(`Element with id ${id} not found`);
        }
        this.worktop = worktop as HTMLDivElement;  // 工作区域
        this.initContainerSplit();
        this.initMenu();
        this.initToolbar();
        this.initSheets();
        this.initbottomToolsSplit();
        this.initRightToolsSplit();
        this.initScroller();
        this.setupDataListeners();
        this.initNavigator(); // 初始化导航器，必须放到最后
    }

    /**
     * 初始化容器分割线
     */
    private initContainerSplit() {
        this.containerSplit = new Split({
            parentElement: this.worktop,
            id: TOOLS_CLASS_NAMES.CONTAINER,
            storageKey: TOOLS_CLASS_NAMES.CONTAINER,
            nextSize: 360,   // 左右分栏不默认平分：右侧面板（导航内容区）默认宽度 360px
            // 仅在真实拖拽结束后保存：容器尺寸变化 / 程序化设置宽度不会污染已保存值
            onDragEnd: (nextSize: number) => this.saveNavContentWidth(nextSize),
        });
        this.containerSplit.handleVisibility();   // 元素是否可见监控
    }
    
    /**
     * 初始化菜单
     */
    private initMenu() {
        this.menu = new Menu({
            parentElement: this.containerSplit.getElements().splitPrev,
        });
    }

    /**
     * 初始化工具栏
     */
    private initToolbar() {
        this.toolbar = new Toolbar({
            parentElement: this.containerSplit.getElements().splitPrev,
            data: this.data,
            menu: this.menu,
            onTabContentShow: (width: number, height: number) => {
                const rect = this.sheets.getElements().sheets.parentElement.getBoundingClientRect();
                // this.sheets.getElements().sheets.style.width = `${rect.width - width}px`;
                this.sheets.getElements().sheets.style.height = `${rect.height - height}px`;
            }
        });
        const quickAccess = new QuickAccess({  // 快速访问分区
            parentElement: this.toolbar.getElements().tabs,
            data: this.data,
            menu: this.menu,
            title: '快速访问'
        });
    }

    /**
     * 初始化导航器
     */
    private async initNavigator() {
        const navigator = new Navigator({
            horizontal: false,
            // 导航项内容区变为可见时，恢复其宽度（已保存值优先，否则用初始化值）
            onItemShow: (navItemText: string) => this.restoreNavContentWidth(navItemText),
        });
        this.navigator = navigator;
        const { panel, content } = navigator.getElements();
        const { splitNext } = this.containerSplit.getElements();
        splitNext.appendChild(content);
        this.rightToolsSplit.getElements().splitPrev.appendChild(panel);
        // navItemMount 内部读取 IndexDB（异步）：必须等待页签与内容项创建完成，
        // 否则 navContentItemMount 会因找不到目标内容项而挂载失败
        try {
            await navigator.navItemMount(...NAV_PANEL);
        } catch {
            // IndexDB 不可用时仅跳过状态恢复，页签元素已创建，继续挂载各页签内容
        }
        await this.initFileTree(navigator);
        this.initZoomSlider();   // 缩放滑块须在导航面板之后挂载：DOM 顺序即视觉顺序（面板在上、滑块在下）
    }

    /**
     * 读取并规整「上次的文件树结构」（页面刷新后还原）
     *
     * 存档为 {@link FileNode} 数组；IndexDB 不可用、存档缺失或结构损坏时返回 null，
     * 由调用方退回内置默认树（仅「工作区」）。
     * @returns {Promise<FileNode[] | null>} 树数据；无有效存档时为 null
     */
    private async loadFileTree(): Promise<FileNode[] | null> {
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
    private persistFileTree(nodes: FileNode[]): void {
        this.settingsDb.setSetting(FILE_TREE_KEY, nodes, 'file-tree').catch(() => {
            // 持久化失败不影响交互
        });
    }

    /**
     * 恢复指定导航项的内容区宽度
     *
     * 优先级：IndexedDB 中已保存的宽度（人工拖拽产生） > `NAV_CONTENT_WIDTH`
     * 中的初始化宽度 > 容器宽度平分。
     * @param navItemText - 导航项文本
     */
    private restoreNavContentWidth(navItemText: string): void {
        const split = this.containerSplit;
        if (!split) return;
        // 记录当前导航项，后续拖拽保存宽度以此为键
        this.currentNavItemText = navItemText;
        const containerWidth = split.getElements().container.clientWidth;
        if (!containerWidth) return;

        // 读取已保存宽度（异步，不阻塞切换：先用初始化值兜底，读取到后再校正）
        this.settingsDb.getSetting(NAV_WIDTH_KEY_PREFIX + navItemText).then((saved: unknown) => {
            if (typeof saved === 'number' && Number.isFinite(saved) && saved > 0) {
                split.setNextSize(saved);
            }
        }).catch(() => {
            // IndexDB 不可用时忽略，退回初始化值
        });

        const initWidth = NAV_CONTENT_WIDTH[navItemText];
        split.setNextSize(initWidth ?? Math.round(containerWidth / 2));
    }

    /**
     * 持久化当前导航项的内容区宽度（人工拖拽后调用）
     * @param width - 内容区宽度（像素）
     */
    private saveNavContentWidth(width: number): void {
        const navItemText = this.currentNavItemText;
        if (!navItemText || !Number.isFinite(width) || width <= 0) return;
        this.settingsDb.setSetting(
            NAV_WIDTH_KEY_PREFIX + navItemText,
            Math.round(width),
            'nav-content-width'
        ).catch(() => {
            // 持久化失败不影响交互
        });
    }

    /**
     * 初始化「文件」页签：树形文件页面
     *
     * 顶部为搜索栏（同行右侧为展开 / 收缩 / 定位按钮），下方为树形文件列表；
     * 新建 / 打开 / 剪切粘贴等操作统一通过树节点右键菜单完成。
     *
     * 树结构本身也持久化：建树前先读取上次的树（{@link loadFileTree}），
     * 之后任何结构变化由 {@link persistFileTree} 整体覆写，
     * 因此刷新页面后树里仍是刷新前的那些文件与文件夹。
     */
    private async initFileTree(navigator: Navigator): Promise<void> {
        // 文件打开器：负责「打开」按钮与主页面拖放，解析成功后把文件登记到文件树
        this.fileOpener = new FileOpener({
            data: this.data,
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
            // 参数一律是**路径**（含扩展名、相对根节点）：存档键由 toSheetKey 去扩展名后得到，
            // 因此不同目录下的同名文件各自拥有独立存档。
            content: {
                create: (paths: string[]) => this.ensureFileContents(paths),
                copy: (plans: FileMovePlan[]) => this.copyFileContents(plans),
                move: (plans: FileMovePlan[]) => this.moveFileContents(plans),
                remove: (paths: string[]) => this.removeFileContents(paths),
            },
        });
        navigator.navContentItemMount(NAV_PANEL[0], this.fileTree.getElements().container);   // NAV_PANEL[0] 即「文件」
        // 「定位」的目标是「主页面正在显示的那个文件」。
        // `activedSheetName` 是**存档键**（去扩展名的路径），树的定位匹配
        // （{@link FileTree.findByName}）末档支持「路径去扩展名」，因此可直接对上；
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
     * 恢复「上次打开的文件」（页面刷新后）
     *
     * **文件树是文件的唯一来源**，记忆只是「上次看的是哪一个」：
     * - 树里一个文件都没有（首次使用，或上次把文件删到一个不剩）→ 打开一个空白
     *   Excel 并把 `sheet1.xlsx` 登记进树（见 {@link openBlankFile}）；
     * - 记忆里的文件仍在树中 → 恢复它：装载其内容、登记并高亮；
     * - 记忆缺失，或所指文件已不在树中（被删 / 被改名）→ 退回树里的第一个文件，
     *   而不是凭记忆造出一个树里没有的文件。
     *
     * 主页面的初始装载固定读取默认工作表（`sheet1`），因此必须等
     * {@link DataCollection.dataLoaded} 之后再切换，否则两次装载会互相覆盖。
     */
    private async restoreActivedFile(): Promise<void> {
        try {
            await this.data.dataLoaded;

            // 树里一个文件都没有：给用户一个可直接编辑的空白文件（需求：补 sheet1.xlsx）
            const files = this.fileTree.listFiles();
            if (files.length === 0) {
                await this.openBlankFile();
                return;
            }

            const saved = await this.settingsDb.getSetting(ACTIVED_FILE_KEY);
            const remembered = typeof saved === 'string' ? saved.trim() : '';
            // files 是**路径**集合；记忆里存的也是路径，两者可直接比较
            const displayName = remembered && this.fileTree.hasFile(remembered) ? remembered : files[0];

            const sheetName = this.toSheetKey(displayName);
            if (sheetName && sheetName !== this.data.activedSheetName) {
                // 上次打开的是别的文件：装载其内容（无存档时按空表处理）
                await this.data.switchToSheet(sheetName);
            }
            this.fileTree.showActivedFile(displayName);
            this.persistActivedFile(displayName);
            this.rememberOpenedFile(displayName);
            await this.reconcileValueStore();
        } catch {
            // 恢复失败（IndexDB 不可用等）不影响后续使用：此时树中不含当前文件
        }
    }

    /**
     * 打开一个空白 Excel 文件，并把它登记到文件树中
     *
     * 用于「刷新后树里一个文件都没有」：默认工作表（`sheet1`）本身就是这个文件的
     * 空表，因此同名时不切换、主页面保持空白；被占用时（如同名文件夹）按命名规则
     * 取 `sheet2.xlsx` 并切换过去。
     */
    private async openBlankFile(): Promise<void> {
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

    /**
     * 双击树节点：把该文件内容装载到主页面
     *
     * 主页面已显示该文件时 {@link DataCollection.switchToSheet} 返回 false，
     * 此时直接返回：内容保持原样（不重绘、不重载），登记 / 高亮 / 记忆也都已就位，
     * 因此「重复双击同一个文件」不产生任何副作用。
     * @param path - 文件路径（含扩展名，相对根节点）
     */
    private async activateFile(path: string): Promise<void> {
        const sheetName = this.toSheetKey(path);
        if (!sheetName) return;

        const switched = await this.data.switchToSheet(sheetName);
        if (!switched) return;   // 已在主页面显示当前文件 → 内容无变化

        this.fileTree.showActivedFile(path);
        this.persistActivedFile(path);
        this.rememberOpenedFile(path);
        this.data.messageBox.success(`已打开 ${this.baseNameOf(path)}`);
    }

    /**
     * 取路径中的文件显示名（末段）
     * @param path - 文件路径
     * @returns {string} 显示名；空路径返回空串
     */
    private baseNameOf(path: string): string {
        const text = (path ?? '').trim();
        const slash = text.lastIndexOf('/');
        return slash >= 0 ? text.slice(slash + 1) : text;
    }

    /**
     * 记忆当前打开的文件路径（刷新后据此决定主页面显示哪个文件）
     *
     * 存**路径**而不是文件名：不同目录的同名文件是两个文件，
     * 只存名字无法在刷新后区分该恢复哪一个。
     * 路径为空时不写入：文件是否存在以树为准。
     * @param path - 文件路径（含扩展名，相对根节点）
     */
    private persistActivedFile(path: string): void {
        const name = (path ?? '').trim();
        if (!name) return;
        this.settingsDb.setSetting(ACTIVED_FILE_KEY, name, 'actived-file').catch(() => {
            // 持久化失败不影响交互
        });
    }

    /**
     * 把文件**路径**换算为内容存档键（IndexedDB `values` / `images` 库的键）
     *
     * 存档键 = 「去扩展名的完整路径」，如 `归档/销售数据.xlsx` → `归档/销售数据`。
     *
     * **为什么要带目录**：不同目录下的同名文件是两个独立文件，各自应有独立存档。
     * 历史实现只用去扩展名的文件名作键（扁平、无目录），导致不同目录的同名文件
     * 共用同一份存档、互相覆盖。
     *
     * 兼容：传入纯文件名（无 `/`）时结果与历史实现完全一致（`sheet1.xlsx` → `sheet1`），
     * 因此「默认空白文件」「从本机打开的文件」等路径均不受影响。
     * @param path - 文件路径（含扩展名，相对根节点）
     * @returns {string} 存档键；名称为空时返回空串
     */
    private toSheetKey(path: string): string {
        return stripExtension(path ?? '').trim();
    }

    /**
     * 记入「打开历史」（最近打开的排在最前，按存档键去重）
     * @param path - 文件路径（含扩展名，相对根节点）
     */
    private rememberOpenedFile(path: string): void {
        const name = (path ?? '').trim();
        if (!name) return;
        const key = this.toSheetKey(name);
        this.fileHistory = [name, ...this.fileHistory.filter((item) => this.toSheetKey(item) !== key)];
        if (this.fileHistory.length > FILE_HISTORY_LIMIT) {
            this.fileHistory.length = FILE_HISTORY_LIMIT;
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
    private async copyFileContents(plans: FileMovePlan[]): Promise<void> {
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
    private async moveFileContents(plans: FileMovePlan[]): Promise<void> {
        try {
            await this.data.dataLoaded;
            for (const plan of plans) {
                const from = this.toSheetKey(plan.from);
                const to = this.toSheetKey(plan.to);
                if (!from || !to || from === to) {
                    // 存档键（去扩展名的路径）没变：无需搬迁。
                    // 注意这与历史实现的判断不同 —— 路径键下「剪切到别的目录」确实
                    // 改变存档键（`a/x.xlsx` → `b/x.xlsx`），必须搬迁，
                    // 只有「同一目录内名字也没变」才是真正的空操作。
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
     * 同步「打开历史 / 当前文件记忆」中的文件路径（改名 / 移动后调用）
     * @param from - 旧路径
     * @param to - 新路径
     */
    private rememberRenamedFile(from: string, to: string): void {
        const key = this.toSheetKey(from);
        if (!key) return;
        this.fileHistory = this.fileHistory.map((item) => (this.toSheetKey(item) === key ? to : item));
        this.settingsDb.getSetting(ACTIVED_FILE_KEY).then((saved: unknown) => {
            if (typeof saved === 'string' && this.toSheetKey(saved) === key) this.persistActivedFile(to);
        }).catch(() => {
            // 读不到记忆时忽略：下一次打开文件会重新写入
        });
    }

    /**
     * 删除文件内容（删除树节点时由内容桥调用，文件夹已递归展开为文件）
     *
     * 被删的正是当前打开的文件时，先回退到**上一个打开过的文件**并高亮它
     * （见 {@link pickFallbackFile}）；无文件可回退则清空主页面。
     * 注意顺序：切换/清空会先固化当前内容，因此存档删除必须放在其后，
     * 否则刚写回的内容会留在库里，看起来「删不掉」。
     * @param paths - 被删除文件的路径集合
     */
    private async removeFileContents(paths: string[]): Promise<void> {
        try {
            await this.data.dataLoaded;
            const keys = paths.map((path) => this.toSheetKey(path)).filter(Boolean);
            if (keys.length === 0) return;

            const removingActived = keys.includes(this.data.activedSheetName);
            let fallback: string | null = null;
            if (removingActived) {
                fallback = this.pickFallbackFile(keys);
                if (fallback) await this.data.switchToSheet(this.toSheetKey(fallback));
                else this.data.clearCurrentSheet();
            }

            // 落库/切换完成后才真正清理存档
            for (const key of keys) {
                await this.settingsDb.deleteValue(key);
                await this.settingsDb.deleteImages(key);
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
                    // 此后刷新时树里没有文件 → 按需求补一个空白 sheet1.xlsx，而不是复活已删的文件。
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
     * 补建文件存档（新增文件时由内容桥调用）
     *
     * 让「树里有这个文件」与「库里有这个文件的存档」始终一一对应：库里还没有
     * 该文件的存档时写入一份空白快照。
     *
     * **只补不覆盖**：这条路径也会在「登记已打开的文件」时被调用，那时内容刚刚由
     * 打开器写库，若直接覆写会把刚打开的文件清空。
     * @param paths - 新增文件的路径集合
     */
    private async ensureFileContents(paths: string[]): Promise<void> {
        try {
            await this.data.dataLoaded;
            for (const path of paths) {
                const key = this.toSheetKey(path);
                if (!key) continue;
                const existing = await this.settingsDb.getValue(key);
                if (existing && existing !== '{}') continue;   // 已有存档：保持原样
                await this.settingsDb.setValue(key, this.data.blankSnapshot());
            }
        } catch {
            // 存档不可用不影响结构操作：此时新文件按空表处理
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
     * 归属判定用**路径键**（{@link toSheetKey}）：`归档/销售数据.xlsx` 与
     * `销售数据.xlsx` 是两个不同的文件，因此也是两个不同的存档。
     */
    private async reconcileValueStore(): Promise<void> {
        try {
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

    /**
     * 初始化底部工具线
     */
    private initbottomToolsSplit() {
        this.bottomToolsSplit = new Split({
            parentElement: this.sheets.getElements().sheets, 
            id: TOOLS_CLASS_NAMES.BOTTOM_TOOLS, 
            storageKey: TOOLS_CLASS_NAMES.BOTTOM_TOOLS
        });
    }

    /**
     * 初始化右侧工具线
     */
    private initRightToolsSplit() {
        this.rightToolsSplit = new Split({
            parentElement: this.sheets.getElements().sheets, 
            id: TOOLS_CLASS_NAMES.RIGHT_TOOLS,
            horizontal:false,
            storageKey: TOOLS_CLASS_NAMES.RIGHT_TOOLS
        });
    }

    /**
     * 初始化缩放滑块（挂载于 rightTools 的 split-pre 内部、贴底显示）
     *
     * 采用垂直细线样式（{@link Slider} 的 vertical + thin），外观由 slider.scss 的主题色板控制；
     * 垂直布局下数值输入框自然落在轨道下方，可手工填写百分比缩放值。
     * 取值与 `data.zoom` 双向绑定：
     * - 拖动滑块 / 填写数值框 → 写回 `data.zoom`（百分比 ÷ 100）；
     * - Ctrl+滚轮 缩放 → 触发 ZOOM_CHANGED → 回写滑块与数值框。
     * 取值范围直接取自 {@link Scroller} 的缩放上下限，避免两处常量漂移。
     */
    private initZoomSlider() {
        this.zoomSlider = new Slider({
            parentElement: this.rightToolsSplit.getElements().splitPrev,
            vertical: true,      // 垂直方向：最小值在下端，数值框在轨道下方
            thin: true,          // 细线样式：2px 轨道 + 小圆点手柄
            min: ZOOM_SLIDER_MIN,
            max: ZOOM_SLIDER_MAX,
            step: 1,
            value: this.zoomPercent(),
            showInput: true,     // 数值框：手工填写缩放百分比（无单位文本，纯数字）
            inputWidth: '',      // 置空：宽度交给样式表（26px 窄条需占满整行）
            accentColor: '',     // 置空走样式表主题色，避免注入内联的浏览器默认蓝
            onChange: (value: number) => {
                this.data.zoom = value / 100;
            },
        });

        // Ctrl+滚轮 或其它入口改变缩放时，同步滑块位置
        this.data.on(DataEvents.ZOOM_CHANGED, () => {
            this.zoomSlider.setValue(this.zoomPercent());
        });
    }

    /**
     * 当前缩放比例对应的百分比整数值（滑块取值单位与 data.zoom 不同，此处统一换算）
     * @returns {number} 百分比整数，如 1.21 → 121
     */
    private zoomPercent(): number {
        return Math.round(this.data.zoom * 100);
    }

    /**
     * 初始化工作表
     */
    private initSheets() {
        const { splitPrev } = this.containerSplit.getElements();
        this.sheets = new Sheets({
            parentElement: splitPrev,
            data: this.data,
            menu: this.menu,
            onResize: (left: number, top: number, width: number, height: number) => {
                const bottomToolsSplitRect  =  this.bottomToolsSplit.getElements().container.getBoundingClientRect();
                const rightToolsSplitRect = this.rightToolsSplit.getElements().container.getBoundingClientRect();
                this.data.visibleView.sheetWidth = Math.round(width - rightToolsSplitRect.width - this.data.visibleView.rowHeaderWidth);
                this.data.visibleView.sheetHeight = Math.round(height - bottomToolsSplitRect.height - this.data.visibleView.colHeaderHeight);
                this.sheets.updateCanvasSize();
            }
        });
    }

    private initScroller() {
        // 滚轮监听统一挂到工作表容器：滚动条实例只负责各自方向的分量，
        // 使 sheetCanvas/editCanvas/textarea 激活时滚轮也能滚动主页面滚动条
        const wheelElement = this.sheets.getElements().sheets;
        this.hScroller = new Scroller({
            parentElement: this.bottomToolsSplit.getElements().splitNext,
            wheelElement,
            direction: 'horizontal',
            data: this.data,
        });
        this.vScroller = new Scroller({
            parentElement: this.rightToolsSplit.getElements().splitNext,
            wheelElement,
            direction: 'vertical',
            data: this.data,
        });
    }

    private setupDataListeners() {
            this.data.on(DataEvents.ACTIVED_CELL_CHANGED, (_, cell) => {
                this.data.isEditting = false;
                // this.sheets.draw(); // 与SELECT_CHANGED事件冲突
            });
            this.data.colHeaders.on(DataEvents.OFFSETX_CHANGED, () => {
                this.sheets.draw();
            });
            this.data.on(DataEvents.ALL_COL_WIDTH_CHANGED, () => {
                this.sheets.draw();
            });
            this.data.rowHeaders.on(DataEvents.OFFSETY_CHANGED, () => {
                this.sheets.draw();
            });
            this.data.on(DataEvents.ALL_ROW_HEIGHT_CHANGED, () => {
                this.sheets.draw();
            });
            this.data.on(DataEvents.SELECTION_CHANGED, () => {
                this.sheets.draw();
            });
            this.data.on(DataEvents.ANCHOR_SELECTION_CHANGED, () => {
                this.sheets.draw();
            });
            this.data.on(DataEvents.IS_EDITTING_CHANGED, (_, isEditting) => {
                if(!isEditting) this.sheets.draw();
            });
            this.data.on(DataEvents.ZOOM_CHANGED, () => {
                this.sheets.draw();
            });
            this.data.on(DataEvents.VALUES_CHANGED, (newValues) => {
                this.sheets.draw();
            });
            this.data.on(DataEvents.COLUMNS_CHANGED, () => {
                this.sheets.draw();
            });
            this.data.on(DataEvents.ROWS_CHANGED, () => {
                this.sheets.draw();
            });
        // this.data.on(DataEvents.FONT_FAMILY_CHANGED, () => {
        //     this.sheets.draw();
        // });
        // this.data.on(DataEvents.FONT_SIZE_CHANGED, () => {
        //     this.sheets.draw();
        // });
        // this.data.on(DataEvents.FONT_WEIGHT_CHANGED, () => {
        //     this.sheets.draw();
        // });
        // this.data.on(DataEvents.FONT_STYLE_CHANGED, () => {
        //     this.sheets.draw();
        // });
        // this.data.on(DataEvents.FONT_COLOR_CHANGED, () => {
        //     this.sheets.draw();
        // });
        // this.data.on(DataEvents.UNDERLINE_CHANGED, () => {
        //     this.sheets.draw();
        // });
        // this.data.on(DataEvents.STRIKETHROUGH_CHANGED, () => {
        //     this.sheets.draw();
        // });
        document.addEventListener('contextmenu', function (e) {
            e.preventDefault(); // 阻止默认右键菜单
        });

        // 自动保存已从「每帧落库」改为「变更后防抖落库」，因此关闭/隐藏页面时必须补一次
        // 立即固化，否则防抖窗口内的最后一次改动会随页面一起丢失。
        // pagehide 覆盖关闭/刷新/前进后退；visibilitychange 覆盖切到后台（移动端可能不触发 pagehide）。
        window.addEventListener('pagehide', () => this.data.flushPendingSave());
        document.addEventListener('visibilitychange', () => {
            if (document.visibilityState === 'hidden') this.data.flushPendingSave();
        });
    }
}