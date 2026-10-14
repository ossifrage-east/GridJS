/*  # 通用文件树组件 TS 文件（「文件」导航页）
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { createDiv, createSpan, createInput, createButton, createElement } from "../../utils/dom";
import { stripExtension, extractExtension } from "../../utils/fileName";
import { FILE_TREE_CLASS_NAMES as TREE, MENU_CLASS_NAMES } from "../constant";
import { Menu, MenuContent } from "./menu";
import { Dialog } from "./dialog";
import '../style/fileTree.scss';

/**
 * 文件节点数据接口
 * @interface FileNode
 */
export interface FileNode {
    /** 节点名称（显示文本，同时用于搜索匹配） */
    name: string;
    /**
     * 节点在树中的完整路径（以 `/` 分隔，含自身名称，不含根节点「工作区」）
     *
     * **这是文件内容的存档键来源**（IndexedDB `values` / `images` 库的 `sheetName`）：
     * 「不同目录下的同名文件」因此拥有各自独立的存档，不再互相覆盖。
     *
     * 由树在导出时按父链推导并写入（见 {@link FileTree.getNodes}），
     * 因此外部无需手工维护；旧存档没有该字段时按「名称即路径」兼容。
     */
    path?: string;
    /** 子节点：存在即视为文件夹（空数组表示空文件夹） */
    children?: FileNode[];
    /**
     * 文件夹展开状态（仅文件夹会带此字段，仅供持久化还原使用）
     *
     * 缺省时按默认规则处理：根层展开、更深层收起（见 {@link createNode}）。
     */
    open?: boolean;
    /**
     * 节点创建时间（Unix 毫秒）
     *
     * 新建节点时写入，随 `onChange` 落库，供「属性」对话框显示「最初建立时间」。
     * 旧存档没有该字段时由 {@link FileTree.ensureMetadata} 补一个（以当前时间为准）。
     */
    createdAt?: number;
    /**
     * 最后修改人
     *
     * 本项目是纯前端单机应用，没有登录态与服务端，因此记录的是**本地会话标识**
     * （见 {@link LOCAL_USER}）——结构化地留住「何人所改」，将来接入账号体系时
     * 只需替换这一处的取值来源。
     */
    lastModifiedBy?: string;
    /** 最后修改时间（Unix 毫秒） */
    lastModifiedAt?: number;
    /** 创建人（同 {@link FileNode.lastModifiedBy}，本机会话标识） */
    createdBy?: string;
    /**
     * 是否可编辑
     *
     * 单机应用里所有文件都可编辑，统一为 `true`；保留该字段是为了让「属性」对话框
     * 的结构与真实协作场景一致（将来接入只读/权限后无需改动调用方）。
     */
    editable?: boolean;
}

/**
 * 「新建文件」下拉菜单支持的扩展名
 * 与 FileOpener.SUPPORTED_EXTENSIONS 保持一致（此处为展示用的菜单项文案）
 */
const NEW_FILE_MENU_ITEMS = [
    { todo: 'XLSX', text: '新建XLSX' },
    // { todo: 'XLS', text: '新建XLS' },
    { todo: 'DOCX', text: '新建DOCX' },
    // { todo: 'DOC', text: '新建DOC' },
    { todo: 'PPTX', text: '新建PPTX' },
    // { todo: 'PPT', text: '新建PPT' },
];

/**
 * 新建文件的初始名称
 *
 * 基础名自带序号（`sheet1.xlsx` / `doc1.docx` / `ppt1.pptx`），
 * 当前目录内重名时**递增其中的数字**（`sheet1.xlsx` → `sheet2.xlsx` …）。
 */
const NEW_FILE_BASE_NAMES: Record<string, string> = {
    XLSX: 'sheet1.xlsx',
    DOCX: 'doc1.docx',
    PPTX: 'ppt1.pptx',
};

/**
 * 文件内容搬迁计划项
 *
 * 树只认识**路径**，而文件内容存在外部（IndexedDB 的 `values` 库，键为
 * 「去根目录的完整路径」，如 `归档/销售数据.xlsx` → `归档/销售数据`）。
 * 因此「文件（或其祖先文件夹）的路径变了 / 被复制了 / 被删除了」这类副作用一律以
 * **路径**告知外部，由外部换算成存档键后执行。
 *
 * 之所以用路径而不是文件名：不同目录下的同名文件是两个独立的文件，
 * 各自拥有独立存档；只传文件名会让它们互相覆盖（历史实现的已知缺陷）。
 * @interface FileMovePlan
 */
export interface FileMovePlan {
    /** 源文件路径（含扩展名，相对根节点） */
    from: string;
    /** 目标文件路径（含扩展名，相对根节点） */
    to: string;
}

/**
 * 内容桥：把「结构变化」翻译成「文件内容的同步动作」
 *
 * 四个动作都由树在**语义发生变化的时刻**调用一次，调用方负责真正读写数据；
 * 参数一律是**路径**（相对根节点，含扩展名），不是文件名：
 * - {@link FileContentBridge.create}：新建文件（库中尚无存档时补一份空白存档）；
 * - {@link FileContentBridge.copy}：复制 / 粘贴（复制模式），含文件夹递归；
 * - {@link FileContentBridge.move}：路径变化（重命名、剪切粘贴、拖拽移动、
 *   以及**祖先文件夹改名**导致其下所有文件路径一起变化）；
 * - {@link FileContentBridge.remove}：删除（含文件夹递归）。
 * @interface FileContentBridge
 */
export interface FileContentBridge {
    /**
     * 补建文件存档：新建出来的文件在数据库中还没有存档时写入一份空白存档
     *
     * 只补不覆盖：这条路径也会在「登记已打开的文件」时被调用，此时内容刚由打开器写入。
     * @param paths - 新增文件的路径集合
     */
    create?: (paths: string[]) => void;
    /** 复制文件内容：把 `from` 的存档复制到 `to` */
    copy?: (plans: FileMovePlan[]) => void;
    /** 搬迁文件内容：把 `from` 的存档搬到 `to`（源键不再保留） */
    move?: (plans: FileMovePlan[]) => void;
    /**
     * 删除文件内容：清理这些路径对应文件在数据库中的存档
     * @param paths - 被删除文件的路径集合
     */
    remove?: (paths: string[]) => void;
}

/**
 * 文件树构造选项接口
 * @interface FileTreeOptions
 */
export interface FileTreeOptions {
    /** 父元素（缺省时仅创建 DOM，不挂载） */
    parentElement?: HTMLElement | null;
    /** 树数据，缺省时使用内置示例数据 {@link DEFAULT_FILE_NODES} */
    nodes?: FileNode[];
    /** 菜单实例：「新建文件」下拉菜单复用全局菜单 */
    menu?: Menu | null;
    /** 点击「打开」按钮时触发（由外部弹出文件选择对话框） */
    onOpen?: () => void;
    /**
     * 双击文件节点时触发（参数为**文件路径**，含扩展名）
     *
     * 由外部装载该文件内容到主页面；若主页面已显示该文件，外部应保持内容不变。
     */
    onActivate?: (path: string) => void;
    /**
     * 树结构、名称、路径或展开状态发生变化时触发
     * （新增 / 删除 / 重命名 / 粘贴 / 拖拽移动 / 展开收起）
     *
     * 用于外部持久化：把回调参数存起来，页面刷新时再通过
     * {@link FileTreeOptions.nodes} 注入还原，即可让树回到刷新前的样子。
     * 参数为当前整棵树的深拷贝（纯数据，可直接交给结构化克隆存储），
     * 每个节点都带 `path`（即内容存档键的来源）。
     */
    onChange?: (nodes: FileNode[]) => void;
    /**
     * 内容桥：复制 / 改名 / 移动 / 删除文件时同步数据库中的文件内容
     *
     * 不传则树只做结构操作（粘贴出来的文件是空表、删除不清理存档）。
     */
    content?: FileContentBridge;
    /**
     * 「属性」对话框单击确定后的回调（可选）
     *
     * 供外部在对话框关闭时做额外处理（如埋点、刷新视图）。不传则仅关闭对话框。
     */
    onPropertyClose?: (path: string) => void;
}

/**
 * 视图节点：数据节点 + DOM 引用 + 交互状态
 * @interface TreeNodeView
 */
interface TreeNodeView {
    /** 节点名称（已转小写，用于搜索匹配） */
    key: string;
    /** 是否文件夹 */
    isFolder: boolean;
    /** 父节点视图（根节点为 null） */
    parent: TreeNodeView | null;
    /** 节点容器：搜索未命中时直接隐藏该元素（行与子树一并隐藏） */
    el: HTMLDivElement;
    /** 节点行：选中、定位闪烁、拖拽的作用元素 */
    row: HTMLDivElement;
    /** 子节点容器（文件夹才有） */
    childrenBox: HTMLDivElement | null;
    /** 类型图标（文件夹 / 文件） */
    icon: HTMLElement;
    /** 节点名称文本元素（重命名时被输入框临时替换，故单独持有引用） */
    label: HTMLSpanElement;
    /** 子节点视图集合 */
    children: TreeNodeView[];
    /** 展开状态：用户操作记忆，搜索期间不被覆盖，清空搜索后据此恢复 */
    isOpen: boolean;
    /** 创建时间（Unix 毫秒） */
    createdAt: number;
    /** 最后修改时间（Unix 毫秒） */
    lastModifiedAt: number;
    /** 创建人 */
    createdBy: string;
    /** 最后修改人 */
    lastModifiedBy: string;
    /** 是否可编辑 */
    editable: boolean;
}

/**
 * 本机会话标识（「创建人 / 最后修改人」的取值）
 *
 * 纯前端单机应用没有登录态，因此以「本机 + 当日会话」作为一个稳定的可读标识：
 * 同一天同一台机器上操作产生同一标识，既不假装成真实用户名，也让「何人」这一栏
 * 有可读、可区分的内容。将来接入账号体系时只需替换本函数。
 * @returns {string} 本机会话标识，如 `本机用户 · 10-14`
 */
function localUser(): string {
    const now = new Date();
    const pad = (n: number): string => String(n).padStart(2, '0');
    return `本机用户 · ${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
}

/** {@link localUser} 的求值结果（同一会话内保持一致） */
const LOCAL_USER = localUser();

/**
 * 默认文件树：仅保留顶层「工作区」
 *
 * 真实数据可通过 {@link FileTreeOptions.nodes} 注入；
 * 新建 / 打开 / 粘贴得到的节点会在运行时追加到工作区之下。
 */
const DEFAULT_FILE_NODES: FileNode[] = [
    {
        name: '工作区',
        children: [],
    },
];

/**
 * 文件树组件
 *
 * 用于右侧「文件」导航页：顶部一行是搜索栏（其右侧紧邻展开、收缩、定位三个图标按钮），
 * 下方是可折叠的树形文件列表。
 *
 * 交互约定：
 * - 搜索：按名称模糊匹配（不区分大小写）；命中节点及其全部祖先保持可见，
 *   命中子孙的文件夹自动展开；清空搜索（含按 Esc）后恢复搜索前的展开状态；
 * - 展开 / 收缩：一键展开或收起全部文件夹（仅影响文件夹节点，不影响搜索过滤结果）；
 * - 定位：定位**当前主页面显示的文件**（由 {@link setActivedFileName} 同步
 *   `data.activedSheetName`）在树中的位置，展开全部祖先并闪烁高亮；
 *   找不到时依次降级为：搜索首个命中项 → 当前选中项 → 第一个根节点。
 * - 右键菜单：新建文件（hover 展开二级菜单：新建XLSX / 新建DOCX / 新建PPTX）/
 *   新建文件夹 / 打开 / 分割线 / 剪切 / 复制 / 粘贴 / 分割线 / 重命名 / 删除。
 *   右键文件夹时「新建」作用于该文件夹，右键文件时作用于其所在文件夹。
 * - 命名规则：新建时**仅当当前目录内同名冲突才递增序号** ——
 *   文件夹无冲突用基础名（`新建文件夹`），冲突时追加序号（`新建文件夹1`…）；
 *   文件基础名自带序号（`sheet1.xlsx` / `doc1.docx` / `ppt1.pptx`），
 *   冲突时递增其中的数字（`sheet1.xlsx` → `sheet2.xlsx` …）。
 *   粘贴 / 复制重名时用同一套规则（{@link uniqueName}）：名称里带数字就递增它
 *   （`doc1.docx` → `doc2.docx`），否则追加序号（`报表.xlsx` → `报表1.xlsx`）。
 * - 双击文件节点：由 {@link FileTreeOptions.onActivate} 交给外部装载内容到主页面；
 *   主页面已显示该文件时不产生任何变化（由外部保证）。
 * - 当前文件高亮：主页面正在显示的文件由 {@link setActivedFileName} /
 *   {@link showActivedFile} 同步，行上加 {@link TREE}.ACTIVED 常驻高亮
 *   （与「选中」`is-selected`、「定位闪烁」`is-located` 三态相互独立）；
 *   文件被改名时高亮跟随新名，被删除时高亮随之消失。
 * - 内容同步：新增 / 复制 / 改名 / 移动 / 删除文件会通过 {@link FileTreeOptions.content}
 *   通知外部同步数据库中该文件的内容（树只认识**路径**，不接触存档数据）；
 *   新增只「补建」不覆盖，因此刷新恢复、打开文件登记等路径同样安全。
 * - 拖拽移动：把文件 / 文件夹拖到另一个文件夹即可改变其路径（含全部子孙），
 *   路径变化由 {@link FileTreeOptions.content} 的 `move` 通知外部搬迁存档；
 *   不能拖到自身或自身的子孙里（防环），拖到原父目录视为无操作。
 * - 属性对话框：右键菜单「属性」打开一个对话框，展示名称与后缀、位置、创建与
 *   修改时间及人员、是否可编辑。
 * - 持久化：新增 / 删除 / 重命名 / 粘贴，以及文件夹展开收起，都会通过
 *   {@link FileTreeOptions.onChange} 通知外部落库；外部在下次启动时用
 *   {@link FileTreeOptions.nodes} 注入还原，即「刷新后树里仍是刷新前的文件，
 *   原来折叠的文件夹也依然是折叠的」（{@link FileTree.getNodes} 可取当前整棵树）。
 * - 删除保护：根节点（「工作区」）不可删除。
 *
 * @class FileTree
 * @example
 * const fileTree = new FileTree({ menu, onOpen: () => opener.pickAndOpen() });
 * navigator.navContentItemMount('文件', fileTree.getElements().container);
 * // 打开文件成功后登记节点，并同步主页面文件名供「定位」使用
 * fileTree.showActivedFile('销售数据.xlsx');
 */
export class FileTree {
    /** 页面容器（挂载到导航内容项） */
    private readonly container: HTMLDivElement;
    /** 顶部工具栏（搜索栏 + 操作按钮） */
    private readonly toolbar: HTMLDivElement;
    /** 搜索输入框 */
    private readonly searchInput: HTMLInputElement;
    /** 树主体（滚动容器） */
    private readonly body: HTMLDivElement;
    /** 无匹配结果提示 */
    private readonly empty: HTMLDivElement;
    /** 配置选项 */
    private readonly options: FileTreeOptions;
    /** 根节点视图集合 */
    private rootViews: TreeNodeView[] = [];
    /** 当前选中节点 */
    private selected: TreeNodeView | null = null;
    /** 当前搜索关键字（小写，空串表示未搜索） */
    private keyword = '';
    /** 定位闪烁定时器 */
    private flashTimer: number | null = null;
    /** 当前主页面显示的文件名（用于「定位」在树中找到对应节点） */
    private activedFileName = '';
    /** 剪贴板暂存：剪切 / 复制的节点数据（粘贴时深拷贝使用） */
    private clipboard: FileNode[] | null = null;
    /** 剪贴板模式：true=剪切（粘贴后删除源），false=复制 */
    private clipboardIsCut = false;
    /** 剪切来源节点：粘贴完成后据此删除原节点（复制模式为 null） */
    private clipboardSources: TreeNodeView[] = [];
    /** 正在重命名的节点视图（同一时刻仅允许一个） */
    private renaming: { view: TreeNodeView; input: HTMLInputElement } | null = null;
    /** 正在被拖拽的节点（拖拽期间非 null，供放置校验与落点目标判定） */
    private dragging: TreeNodeView | null = null;
    /** 「属性」对话框实例（复用同一实例，避免反复创建 DOM） */
    private propertyDialog: Dialog | null = null;

    /**
     * 构造函数
     * @param {FileTreeOptions} [options={}] - 组件配置选项
     */
    constructor(options: FileTreeOptions = {}) {
        this.options = options;
        this.container = createDiv({ className: TREE.CONTAINER });
        this.body = createDiv({ className: TREE.BODY });
        this.empty = createDiv({ className: TREE.EMPTY, textContent: '未找到匹配的文件' });
        this.empty.hidden = true;
        this.searchInput = createInput({
            className: TREE.SEARCH_INPUT,
            type: 'text',
            attributes: { placeholder: '搜索文件', spellcheck: 'false', autocomplete: 'off' },
        });

        this.body.append(this.empty);
        this.toolbar = this.createToolbar();
        this.container.append(this.toolbar, this.body);

        this.rootViews = this.createNodes(options.nodes ?? DEFAULT_FILE_NODES, this.body, 0, null);
        // 旧存档（只存 name/children/open）恢复出来的节点补齐创建/修改等元数据，
        // 使「属性」对话框不出现空白项
        this.rootViews.forEach((view) => this.ensureMetadata(view));

        this.bindEvents();
        this.updateEmptyState();

        if (options.parentElement) {
            options.parentElement.appendChild(this.container);
        }
    }

    /**
     * 在指定父节点下新建一个文件夹
     * @private
     * @param {TreeNodeView | null} parent - 父节点（缺省为根节点「工作区」）
     */
    private createFolderNode(parent: TreeNodeView | null = null): void {
        this.addNode(this.availableName('新建文件夹', parent), { isFolder: true, select: true, parent });
    }

    /**
     * 在指定父节点下新建一个空白文件节点
     * @private
     * @param {string} todo - 菜单动作标识（XLSX / DOCX / PPTX）
     * @param {TreeNodeView | null} parent - 父节点（缺省为根节点「工作区」）
     */
    private createFileNode(todo: string, parent: TreeNodeView | null = null): void {
        this.addNode(this.availableFileName(todo, parent), { select: true, parent });
    }

    /**
     * 生成当前目录内可用的文件名称
     *
     * 规则：基础名自带序号（{@link NEW_FILE_BASE_NAMES}，如 `sheet1.xlsx`），
     * 当前目录已存在同名时**递增其中的数字**：`sheet1.xlsx` → `sheet2.xlsx` → …
     * 序号自基础名的数字起步，因此删除 `sheet2.xlsx` 后再新建会重新得到 `sheet2.xlsx`，
     * 与文件夹命名（{@link availableName}）「以目录实际内容为准」的风格一致。
     * @private
     * @param {string} todo - 菜单动作标识（XLSX / DOCX / PPTX）
     * @param {TreeNodeView | null} parent - 父节点（缺省为根节点）
     * @returns {string} 当前目录内不冲突的文件名
     */
    private availableFileName(todo: string, parent: TreeNodeView | null = null): string {
        const target = parent?.childrenBox ? parent : this.rootViews[0];
        const taken = new Set(target?.children.map((child) => child.key) ?? []);
        const base = NEW_FILE_BASE_NAMES[todo] ?? `${todo.toLowerCase()}1.${todo.toLowerCase()}`;

        // 拆出「前缀 + 数字 + 扩展名」，重名时递增中间的数字
        const matched = /^(.*?)(\d+)(\.[^.]+)?$/.exec(base);
        const prefix = matched ? matched[1] : base;
        const suffix = matched ? (matched[3] ?? '') : '';
        const start = matched ? Number(matched[2]) : 1;

        for (let seq = start; seq < start + 10000; seq++) {
            const candidate = `${prefix}${seq}${suffix}`;
            if (!taken.has(candidate.toLowerCase())) return candidate;
        }
        return base;
    }

    /**
     * 生成当前目录内可用的名称（文件夹命名规则）
     *
     * 规则：**仅当当前目录已存在同名节点时才追加序号** —— 目录内无冲突则原样返回，
     * 如「新建文件夹」→「新建文件夹1」→「新建文件夹2」…
     *
     * 与「删除后序号继续递增」不同，这里以目录实际内容为准，
     * 因此删除「新建文件夹1」后再次新建会重新得到「新建文件夹1」，
     * 始终保证「名称不重复」且「序号从冲突处开始」。
     * 文件命名见 {@link availableFileName}（递增基础名中的数字）。
     * @private
     * @param {string} base - 基础名称（如 '新建文件夹'）
     * @param {TreeNodeView | null} parent - 父节点（缺省为根节点）
     * @returns {string} 当前目录内不冲突的名称
     */
    private availableName(base: string, parent: TreeNodeView | null = null): string {
        const target = parent?.childrenBox ? parent : this.rootViews[0];
        const taken = new Set(target?.children.map((child) => child.key) ?? []);

        // 无冲突：直接用基础名（不追加任何序号）
        if (!taken.has(base.toLowerCase())) return base;

        // 有冲突：从 1 开始递增，跳过已占用者
        for (let seq = 1; seq < 10000; seq++) {
            const candidate = `${base}${seq}`;
            if (!taken.has(candidate.toLowerCase())) return candidate;
        }
        return base;
    }

    /**
     * 创建顶部工具栏
     * 同一行内：左侧搜索栏，右侧展开 / 收缩 / 定位按钮（按钮只显示图标）
     * @private
     * @returns {HTMLDivElement} 工具栏元素
     */
    private createToolbar(): HTMLDivElement {
        const toolbar = createDiv({ className: TREE.TOOLBAR });

        const search = createDiv({ className: TREE.SEARCH });
        search.append(createElement('i', { className: 'icon-search' }), this.searchInput);

        const actions = createDiv({ className: TREE.ACTIONS });
        actions.append(
            this.createIconButton('展开', 'icon-expand-all', () => this.expandAll()),
            this.createIconButton('收缩', 'icon-collapse-all', () => this.collapseAll()),
            this.createIconButton('定位', 'icon-locate', () => this.locate()),
        );

        toolbar.append(search, actions);
        return toolbar;
    }

    /**
     * 创建图标按钮
     * 按钮内容只有图标，无文字；标题通过 title 属性提供悬停提示
     * @private
     * @param {string} title - 按钮标题
     * @param {string} icon - 图标类名
     * @param {() => void} onClick - 点击回调
     * @returns {HTMLButtonElement} 按钮元素
     */
    private createIconButton(title: string, icon: string, onClick: () => void): HTMLButtonElement {
        const button = createButton({
            className: TREE.BUTTON,
            attributes: { type: 'button', title },
        });
        button.appendChild(createElement('i', { className: icon }));
        button.addEventListener('click', () => onClick());
        return button;
    }

    /**
     * 绑定工具栏事件
     * @private
     */
    private bindEvents(): void {
        this.searchInput.addEventListener('input', () => this.applyKeyword());
        this.searchInput.addEventListener('keydown', (e: KeyboardEvent) => {
            // Esc：清空搜索并恢复完整树
            if (e.key === 'Escape') {
                this.searchInput.value = '';
                this.applyKeyword();
            }
        });

        // 内容区任意位置右键均可打开菜单：
        // 节点行上的右键已在 createNode 中 stopPropagation 拦截并自行弹菜单，
        // 能冒泡到这里的多为空白区域（树主体空隙、工具栏、无匹配提示等）。
        // 操作目标取当前选中节点；无选中时回落到根节点（「工作区」），
        // 使「新建 / 粘贴」作用于工作区、其余项随选中状态自然可用或置灰语义一致。
        this.container.addEventListener('contextmenu', (e: MouseEvent) => {
            e.preventDefault();
            const view = this.selected ?? this.rootViews[0];
            if (!view) return;
            this.openContextMenu(e, view);
        });
    }

    /**
     * 递归创建节点视图
     * @private
     * @param {FileNode[]} nodes - 数据节点集合
     * @param {HTMLElement} parentElement - 追加到的父元素
     * @param {number} depth - 层级（0 为根节点，用于缩进）
     * @param {TreeNodeView | null} parent - 父节点视图
     * @returns {TreeNodeView[]} 节点视图集合
     */
    private createNodes(
        nodes: FileNode[],
        parentElement: HTMLElement,
        depth: number,
        parent: TreeNodeView | null
    ): TreeNodeView[] {
        return nodes.map((node) => this.createNode(node, parentElement, depth, parent));
    }

    /**
     * 创建单个节点（含子树）
     * @private
     * @param {FileNode} data - 节点数据
     * @param {HTMLElement} parentElement - 追加到的父元素
     * @param {number} depth - 层级
     * @param {TreeNodeView | null} parent - 父节点视图
     * @returns {TreeNodeView} 节点视图
     */
    private createNode(
        data: FileNode,
        parentElement: HTMLElement,
        depth: number,
        parent: TreeNodeView | null
    ): TreeNodeView {
        const isFolder = Array.isArray(data.children);
        const el = createDiv({ className: TREE.NODE });
        const row = createDiv({
            className: TREE.ROW,
            style: { paddingLeft: `${depth * 14 + 4}px` },
        });

        // 展开箭头：文件夹显示；文件使用等宽占位，保证同层图标对齐
        const toggle = createDiv({ className: TREE.TOGGLE });
        if (isFolder) {
            toggle.appendChild(createElement('i', { className: 'icon-chevron-down' }));
        }
        row.appendChild(toggle);

        // 类型图标：文件夹 / 文件
        const iconWrap = createDiv({ className: TREE.ICON });
        const icon = createElement('i', { className: isFolder ? 'icon-folder' : 'icon-file' });
        iconWrap.appendChild(icon);
        row.appendChild(iconWrap);

        // 节点名称
        const label = createSpan({ className: TREE.LABEL, textContent: data.name });
        row.appendChild(label);
        el.appendChild(row);

        const view: TreeNodeView = {
            key: data.name.toLowerCase(),
            isFolder,
            parent,
            el,
            row,
            childrenBox: null,
            icon,
            label,
            children: [],
            isOpen: true,
            // 元数据：存档里带就用存档的，缺失（旧存档）则以当前时间补齐
            createdAt: typeof data.createdAt === 'number' && Number.isFinite(data.createdAt)
                ? data.createdAt : Date.now(),
            lastModifiedAt: typeof data.lastModifiedAt === 'number' && Number.isFinite(data.lastModifiedAt)
                ? data.lastModifiedAt : Date.now(),
            createdBy: typeof data.createdBy === 'string' && data.createdBy ? data.createdBy : LOCAL_USER,
            lastModifiedBy: typeof data.lastModifiedBy === 'string' && data.lastModifiedBy
                ? data.lastModifiedBy : LOCAL_USER,
            editable: data.editable !== false,
        };

        if (isFolder) {
            const childrenBox = createDiv({ className: TREE.CHILDREN });
            view.childrenBox = childrenBox;
            // 子节点通过 createNode 内部登记到 view.children，此处无需再接收返回值
            this.createNodes(data.children as FileNode[], childrenBox, depth + 1, view);
            el.appendChild(childrenBox);
            // 默认展开第一层（根节点），更深的层级收起；
            // 持久化还原时以存档里的展开状态为准（data.open）
            this.setOpen(view, data.open ?? depth < 1);
        }

        // 登记到父节点的子节点集合：新增节点（打开文件 / 新建）同样依赖此登记，
        // 否则搜索过滤、展开收起与名称去重都会漏掉该节点
        if (parent) parent.children.push(view);

        row.addEventListener('click', () => this.handleRowClick(view));
        // 双击文件节点：由外部把该文件内容装载到主页面（传**路径**，即存档键的来源）
        // （文件夹不处理：展开 / 收起已由单击负责，双击会连点两次反而相互抵消）
        row.addEventListener('dblclick', () => {
            if (view.isFolder) return;
            this.options.onActivate?.(this.pathWithSelf(view));
        });
        // 右键：在鼠标位置弹出上下文菜单（选中被右键的节点作为操作目标）
        row.addEventListener('contextmenu', (e) => {
            e.preventDefault();
            e.stopPropagation();
            this.openContextMenu(e, view);
        });
        this.bindDragEvents(view);
        parentElement.appendChild(el);
        return view;
    }

    // ════════════════════════════════════════════════════════════
    //  拖拽移动（改变路径）
    // ════════════════════════════════════════════════════════════

    /**
     * 绑定节点的拖拽事件（HTML5 Drag & Drop）
     *
     * 语义与「剪切 → 粘贴」等价：把节点（含其子树）移动到目标文件夹下，
     * 由 {@link moveNodeTo} 统一完成校验、搬迁、内容同步与落库。
     *
     * 拖拽越界保护在**拖起时就判定**（`dragstart` 里检测自身/子孙行）
     * 而不是等到 `dragover`：被拖拽的元素在拖拽期间通常不派发 dragover/drop，
     * 只有靠 `dragstart` 阶段标记「禁止」才能可靠拦下「拖到自己子树里」。
     * @private
     * @param {TreeNodeView} view - 节点
     */
    private bindDragEvents(view: TreeNodeView): void {
        view.row.draggable = true;

        view.row.addEventListener('dragstart', (e: DragEvent) => {
            // 重命名输入框内的拖拽是文本选择，不应触发节点搬移
            if (this.renaming?.view === view) {
                e.preventDefault();
                return;
            }
            this.dragging = view;
            // 拖拽期间禁用「自身与全部子孙」作为放置目标，避免拖成环
            this.rootViews.forEach((root) => this.walk(root, (node) => {
                if (node === view || this.isDescendantView(view, node)) {
                    node.row.draggable = false;
                    node.row.classList.add(TREE.DRAG_DISABLED);
                }
            }));
            if (e.dataTransfer) {
                e.dataTransfer.effectAllowed = 'move';
                // 部分浏览器（Firefox）要求必须写入数据才会开始拖拽
                e.dataTransfer.setData('text/plain', view.label.textContent ?? '');
            }
        });

        view.row.addEventListener('dragend', () => {
            this.clearDragState();
        });

        // 可放置：仅文件夹（文件节点不接受放置，与资源管理器一致）
        view.row.addEventListener('dragover', (e: DragEvent) => {
            if (!view.isFolder || !this.canDropOn(view)) return;
            e.preventDefault();
            e.stopPropagation();
            if (e.dataTransfer) e.dataTransfer.dropEffect = 'move';
            view.row.classList.add(TREE.DROP_TARGET);
        });

        view.row.addEventListener('dragleave', () => {
            view.row.classList.remove(TREE.DROP_TARGET);
        });

        view.row.addEventListener('drop', (e: DragEvent) => {
            view.row.classList.remove(TREE.DROP_TARGET);
            if (!view.isFolder || !this.canDropOn(view)) return;
            e.preventDefault();
            e.stopPropagation();
            const source = this.dragging;
            if (!source) return;
            // 不能先 clearDragState：它会清空 `this.dragging`，
            // 而 moveNodeTo 内部还要用 canDropOn 做一次校验（二次校验会因此恒为 false）
            this.moveNodeTo(source, view);
            this.clearDragState();
        });
    }

    /**
     * 判断当前拖拽的节点能否放置到目标文件夹下（O(1) 校验）
     * @private
     * @param {TreeNodeView} target - 目标文件夹
     * @returns {boolean} 是否允许放置
     */
    private canDropOn(target: TreeNodeView): boolean {
        const source = this.dragging;
        if (!source || source === target) return false;
        // 不能拖进自己的子孙（会形成环），也不能拖到当前父节点（无变化）
        if (this.isDescendantView(source, target)) return false;
        if (source.parent === target) return false;
        return true;
    }

    /**
     * 清理拖拽过程中的临时状态（游标、禁止标记、放置高亮）
     * @private
     */
    private clearDragState(): void {
        this.dragging = null;
        this.rootViews.forEach((root) => this.walk(root, (node) => {
            node.row.draggable = true;
            node.row.classList.remove(TREE.DRAG_DISABLED);
            node.row.classList.remove(TREE.DROP_TARGET);
        }));
    }

    /**
     * 把节点（含子树）移动到目标文件夹下 —— 即改变其路径
     *
     * 顺序很重要，且与「剪切 → 粘贴」的既有做法一致：
     * 1. 先确定新名称（目标目录内重名去重，规则见 {@link uniqueName}）；
     * 2. 在**移动之前**按「旧路径 → 新路径」收集搬迁计划并交给内容桥
     *    （移动之后父链已变，再也算不出旧路径）；
     * 3. 再从 DOM 与父链上摘除、挂到目标文件夹并修正缩进；
     * 4. 最后落库。
     *
     * 文件夹整体移动时，其中每个文件的路径都变了 → 计划里逐个列出。
     * @private
     * @param {TreeNodeView} source - 被拖拽的节点
     * @param {TreeNodeView} target - 目标文件夹
     */
    private moveNodeTo(source: TreeNodeView, target: TreeNodeView): void {
        if (!this.canDropOn(target) || !target.childrenBox) return;

        const oldPaths = this.collectFilePaths(source);
        const name = this.uniqueName(target, source.label.textContent ?? '');

        // 先提交重命名（名称可能因重名去重而变），再算新路径
        const renamed = name !== source.label.textContent;
        if (renamed) {
            source.label.textContent = name;
            source.key = name.toLowerCase();
        }

        // 摘除 → 挂载：parent 与 children 数组同步维护
        const siblings = source.parent?.children;
        if (siblings) {
            const index = siblings.indexOf(source);
            if (index >= 0) siblings.splice(index, 1);
        }
        source.parent = target;
        target.children.push(source);
        target.childrenBox.appendChild(source.el);
        source.el.style.paddingLeft = '';

        // 修正整棵子树的缩进（深度变了）
        this.walk(source, (node) => {
            node.row.style.paddingLeft = `${this.getDepth(node) * 14 + 4}px`;
        });

        // 内容搬迁：按「旧路径 → 新路径」一一对应（顺序与 collectFilePaths 一致）
        const newPaths = this.collectFilePaths(source);
        const plans: FileMovePlan[] = [];
        for (let i = 0; i < oldPaths.length && i < newPaths.length; i++) {
            if (oldPaths[i] !== newPaths[i]) plans.push({ from: oldPaths[i], to: newPaths[i] });
        }
        if (plans.length > 0) this.options.content?.move?.(plans);

        // 被移动的是当前文件、或其祖先时，高亮与「当前文件名」需要跟着走
        this.retargetActived(source, oldPaths, newPaths);

        this.touchNode(source);
        this.setOpen(target, true);
        this.select(source);
        this.notifyTreeChanged();
    }

    /**
     * 移动后修正「当前主页面显示的文件」的记忆
     *
     * 被移动的节点本身可能就是当前文件，也可能是它的祖先文件夹（此时文件路径整体改变）。
     * @private
     * @param {TreeNodeView} moved - 被移动的节点
     * @param {string[]} oldPaths - 移动前的全部文件路径（先序）
     * @param {string[]} newPaths - 移动后的全部文件路径（先序，与 oldPaths 一一对应）
     */
    private retargetActived(moved: TreeNodeView, oldPaths: string[], newPaths: string[]): void {
        if (!this.activedFileName) return;
        const current = this.findByName(this.activedFileName);
        if (!current) return;
        // 当前文件不在被移动的子树里：路径未变
        if (current !== moved && !this.isDescendantView(moved, current)) return;

        for (let i = 0; i < oldPaths.length; i++) {
            if (stripExtension(oldPaths[i]) === stripExtension(this.activedFileName)) {
                this.activedFileName = newPaths[i] ?? this.activedFileName;
                return;
            }
        }
    }

    /**
     * 行点击：选中节点；文件夹同时切换展开 / 收起
     * @private
     * @param {TreeNodeView} view - 被点击的节点
     */
    private handleRowClick(view: TreeNodeView): void {
        this.select(view);
        if (view.isFolder) {
            this.setOpen(view, !view.isOpen);
            // 展开 / 收起决定「刷新前能看到哪些内容」，因此一并落库：
            // 否则刚建好并放满文件的文件夹刷新后是折叠态，里面的文件看起来「消失」了。
            // 一次点击一次写入（记录很小），且仅此处与 expandAll / collapseAll / locate 会写。
            this.notifyTreeChanged();
        }
    }

    // ════════════════════════════════════════════════════════════
    //  右键菜单
    // ════════════════════════════════════════════════════════════

    /**
     * 打开节点右键菜单
     *
     * 菜单结构（自上而下）：新建文件（hover 展开二级菜单）/ 新建文件夹 / 打开 / 分割线 /
     * 剪切 / 复制 / 粘贴 / **删除** / 分割线 / 重命名 / **属性**。
     *
     * 「删除」放在「粘贴」之下、分割线之上（危险操作紧邻同组的文件操作，
     * 与「重命名 / 属性」这类节点自身操作分隔开）。
     * 右键文件节点时「新建」类操作作用于其所在文件夹；右键文件夹时作用于该文件夹。
     * @private
     * @param {MouseEvent} e - 右键事件
     * @param {TreeNodeView} view - 被右键的节点
     */
    private openContextMenu(e: MouseEvent, view: TreeNodeView): void {
        const { menu } = this.options;
        if (!menu) return;

        // 右键即选中，使菜单操作对象与高亮一致
        this.select(view);
        // 新建目标：文件夹自身；文件则用其父节点（根节点的父为 null，回落到根节点）
        const parent = view.isFolder ? view : view.parent;

        const content = new MenuContent({
            items: [
                // 新建文件：hover 展开多项二级菜单
                () => this.createNewFileSubmenuItem(view),
                { todo: 'newFolder', icon: 'icon-new-folder', text: '新建文件夹' },
                { todo: 'open', icon: 'icon-open', text: '打开' },
                'separator',
                { todo: 'cut', icon: 'icon-scissors', text: '剪切' },
                { todo: 'copy', icon: 'icon-copy', text: '复制' },
                { todo: 'paste', icon: 'icon-paste', text: '粘贴' },
                { todo: 'remove', icon: 'icon-bin', text: '删除' },
                'separator',
                { todo: 'rename', icon: 'icon-text-height', text: '重命名' },
                { todo: 'property', icon: 'icon-eye', text: '属性' },
            ],
            onClick: (todo: string) => {
                menu.closeMenu();
                this.handleContextAction(todo, view, parent);
            },
        });

        menu.openContextMenu(e.clientX, e.clientY, content.getElements().container);
    }

    /**
     * 构建「新建文件」右键菜单项（含 hover 展开的二级菜单）
     *
     * 二级菜单沿用主表格右键菜单的既有做法：独立容器 append 到 body
     * （主菜单 max-height + overflow 会被裁剪），定位到父项右侧并重叠 1px
     * 消除 hover 间隙，mouseleave 延迟 120ms 关闭以免鼠标穿越间隙时误关。
     * @private
     * @param {TreeNodeView} view - 被右键的节点
     * @returns {HTMLElement} 菜单项元素
     */
    private createNewFileSubmenuItem(view: TreeNodeView): HTMLElement {
        const parent = view.isFolder ? view : view.parent;
        const wrap = createDiv({
            className: 'menu-item menu-item-has-submenu',
            style: { position: 'relative' },
        });
        const iconEl = createElement('i', { className: 'icon-new-file' });
        const textSpan = createSpan({ className: 'text', textContent: '新建文件' });
        const arrowSpan = createSpan({ textContent: '▶' });
        arrowSpan.style.cssText = 'margin-left:auto; font-size:10px; color:#999;';
        wrap.append(iconEl, textSpan, arrowSpan);

        let submenuEl: HTMLElement | null = null;
        let closeTimer: ReturnType<typeof setTimeout> | null = null;
        const cancelPendingClose = (): void => {
            if (closeTimer !== null) {
                clearTimeout(closeTimer);
                closeTimer = null;
            }
        };
        const closeSubmenu = (): void => {
            cancelPendingClose();
            if (submenuEl?.isConnected) submenuEl.remove();
            submenuEl = null;
        };

        wrap.addEventListener('mouseenter', () => {
            cancelPendingClose();
            if (submenuEl) return;   // 已展开则跳过
            const subContent = new MenuContent({
                items: NEW_FILE_MENU_ITEMS,
                onClick: (todo: string) => {
                    closeSubmenu();
                    this.options.menu?.closeMenu();
                    this.createFileNode(todo, parent);
                },
            });
            submenuEl = subContent.getElements().container;
            submenuEl.classList.add(MENU_CLASS_NAMES.CONTAINER);
            // 内联覆盖 .menu 的 opacity:0 / visibility:hidden / transition，
            // 否则子菜单需等淡入动画，感知为延迟显示
            submenuEl.style.cssText = 'position:absolute; display:flex; opacity:1; visibility:visible; transition:none; min-width:160px; box-shadow:0 4px 8px rgba(0,0,0,0.15); z-index:1002;';
            document.body.appendChild(submenuEl);
            // 定位到父项右边缘（重叠 1px 消除 hover gap），空间不足时翻到左侧
            const wrapRect = wrap.getBoundingClientRect();
            const subRect = submenuEl.getBoundingClientRect();
            let left = wrapRect.right - 1;
            let top = wrapRect.top;
            if (left + subRect.width > window.innerWidth) {
                left = wrapRect.left - subRect.width + 1;
            }
            if (top + subRect.height > window.innerHeight) {
                top = window.innerHeight - subRect.height - 4;
            }
            if (top < 0) top = 4;
            submenuEl.style.left = `${left}px`;
            submenuEl.style.top = `${top}px`;
            submenuEl.addEventListener('mouseenter', cancelPendingClose);
        });

        // 主菜单关闭时兜底清理，防止子菜单残留
        this.options.menu?.on('close', closeSubmenu);
        wrap.addEventListener('mouseleave', () => {
            cancelPendingClose();
            closeTimer = setTimeout(closeSubmenu, 120);
        });
        return wrap;
    }

    // ════════════════════════════════════════════════════════════
    //  属性对话框
    // ════════════════════════════════════════════════════════════

    /**
     * 显示节点「属性」对话框
     *
     * 内容为**结构化占位**（见 {@link buildPropertyContent}）：数据全部取自节点自身
     * （名称 / 路径 / 创建与修改的时间与人 / 可编辑性），不依赖任何服务端。
     * 「最后修改」「创建人」等字段所依据的数据在真实协作场景中来自账号体系与后端，
     * 本项目暂以本机会话标识占位（见 {@link LOCAL_USER}），字段结构已按最终形态预留。
     * @private
     * @param {TreeNodeView} view - 目标节点
     */
    private showProperty(view: TreeNodeView): void {
        const path = this.pathWithSelf(view);
        if (!this.propertyDialog) {
            this.propertyDialog = new Dialog({
                title: '属性',
                buttons: [{ text: '确定', variant: 'primary', autoFocus: true }],
                size: 'md',
            });
        }
        this.propertyDialog.setTitle(`${view.isFolder ? '文件夹' : '文件'}属性`);
        this.propertyDialog.setContent(this.buildPropertyContent(view, path));
        this.propertyDialog.show();
        this.options.onPropertyClose?.(path);
    }

    /**
     * 构建属性对话框的正文（定义列表：字段名 + 字段值）
     *
     * 五项内容与需求一一对应：
     * a. **名称及后缀** —— 文件显示 `名称` + `后缀` 两行；文件夹无后缀则不显示该行；
     * b. **最初建立时间**；
     * c. **何人何时做最后修改**（人 + 时间两行）；
     * d. **是否可以编辑**；
     * e. **创建人**。
     * @private
     * @param {TreeNodeView} view - 目标节点
     * @param {string} path - 节点完整路径
     * @returns {HTMLElement} 正文容器
     */
    private buildPropertyContent(view: TreeNodeView, path: string): HTMLElement {
        const name = view.label.textContent ?? '';
        const ext = extractExtension(name);
        const rows: Array<[string, string]> = [];

        // a. 名称及后缀（文件夹无后缀 → 不显示「后缀」这一行）
        rows.push(['名称', name]);
        if (!view.isFolder && ext) rows.push(['后缀', `.${ext}`]);
        rows.push(['位置', path]);

        // b. 最初建立时间
        rows.push(['最初建立时间', this.formatTime(view.createdAt)]);

        // c. 何人何时做最后修改
        rows.push(['最后修改人', view.lastModifiedBy || '—']);
        rows.push(['最后修改时间', this.formatTime(view.lastModifiedAt)]);

        // d. 是否可以编辑
        rows.push(['是否可以编辑', view.editable ? '是' : '否']);

        // e. 创建人
        rows.push(['创建人', view.createdBy || '—']);

        const table = createDiv({ className: TREE.PROPERTY });
        rows.forEach(([label, value]) => {
            const row = createDiv({ className: TREE.PROPERTY_ROW });
            row.appendChild(createDiv({ className: TREE.PROPERTY_LABEL, textContent: label }));
            row.appendChild(createDiv({ className: TREE.PROPERTY_VALUE, textContent: value }));
            table.appendChild(row);
        });
        return table;
    }

    /**
     * 格式化时间戳为可读文本（`YYYY-MM-DD HH:mm:ss`）
     *
     * 手工拼接而不用 `toLocaleString`：后者的输出随运行环境语言/时区变化，
     * 既不利于测试断言，也让同一份数据在不同机器上显示不一致。
     * @private
     * @param {number} ts - Unix 毫秒时间戳
     * @returns {string} 格式化文本；非法值返回 `—`
     */
    private formatTime(ts: number): string {
        if (!Number.isFinite(ts) || ts <= 0) return '—';
        const d = new Date(ts);
        const pad = (n: number): string => String(n).padStart(2, '0');
        return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} `
            + `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
    }

    /**
     * 处理右键菜单动作
     * @private
     * @param {string} todo - 动作标识
     * @param {TreeNodeView} view - 操作目标节点
     * @param {TreeNodeView | null} parent - 新建 / 粘贴的目标父节点
     */
    private handleContextAction(
        todo: string,
        view: TreeNodeView,
        parent: TreeNodeView | null
    ): void {
        switch (todo) {
            case 'newFolder':
                this.createFolderNode(parent);
                break;
            case 'open':
                // 「打开」由外部接管（弹出文件选择对话框并装载）
                this.options.onOpen?.();
                break;
            case 'cut':
                this.clipboard = this.serializeNodes([view]);
                this.clipboardIsCut = true;
                this.clipboardSources = [view];
                break;
            case 'copy':
                this.clipboard = this.serializeNodes([view]);
                this.clipboardIsCut = false;
                // 复制同样记录来源：用于阻止把文件夹粘贴进它自己的子孙
                this.clipboardSources = [view];
                break;
            case 'paste':
                this.pasteNodes(parent);
                break;
            case 'rename':
                this.startRename(view);
                break;
            case 'remove':
                this.removeNodes([view]);
                break;
            case 'property':
                this.showProperty(view);
                break;
            default:
                break;
        }
    }

    // ════════════════════════════════════════════════════════════
    //  剪切 / 复制 / 粘贴
    // ════════════════════════════════════════════════════════════

    /**
     * 将视图序列化为纯数据（供剪贴板暂存 / 持久化导出，切断与 DOM 的引用）
     * @private
     * @param {TreeNodeView[]} views - 节点集合
     * @param {boolean} [withMeta=false] - 是否附带路径与元数据（持久化用；剪贴板不需要）
     * @returns {FileNode[]} 深拷贝后的数据
     */
    private serializeNodes(views: TreeNodeView[], withMeta = false): FileNode[] {
        return views.map((view) => {
            const name = view.label.textContent ?? '';
            const node: FileNode = { name };
            if (view.isFolder) {
                node.children = view.children.map((child) => this.serializeNodes([child], withMeta)[0]);
            }
            if (withMeta) {
                // 路径按父链推导（文件夹自带 `/` 前缀、文件不带）
                node.path = this.pathWithSelf(view);
                node.createdAt = view.createdAt;
                node.lastModifiedAt = view.lastModifiedAt;
                node.createdBy = view.createdBy;
                node.lastModifiedBy = view.lastModifiedBy;
                node.editable = view.editable;
                if (view.isFolder) node.open = view.isOpen;
            }
            return node;
        });
    }

    /**
     * 导出当前整棵树（深拷贝，可直接交给结构化克隆持久化）
     *
     * 与剪贴板用的 {@link serializeNodes} 同源，额外带上：
     * - **路径**（`path`，内容存档键的来源，见 {@link FileNode.path}）；
     * - 创建 / 最后修改的时间与人（供「属性」对话框显示）；
     * - 文件夹展开状态（`open`），使刷新还原后「原来折叠的文件夹依然是折叠的」。
     * @returns {FileNode[]} 树数据
     */
    public getNodes(): FileNode[] {
        return this.serializeNodes(this.rootViews, true);
    }

    /**
     * 通知外部树已变化（供持久化）
     *
     * 在**结构、名称或文件夹展开状态**发生变化时调用：新增 / 删除 / 重命名 /
     * 粘贴，以及展开收起（单点、全展开 / 全收起、定位时的祖先展开）。
     * 纯选中不触发；搜索只做临时渲染（{@link renderOpen}）不改状态，同样不触发。
     * @private
     */
    private notifyTreeChanged(): void {
        this.options.onChange?.(this.getNodes());
    }

    /**
     * 粘贴剪贴板中的节点到指定父节点下
     * @private
     * @param {TreeNodeView | null} parent - 目标父节点
     */
    private pasteNodes(parent: TreeNodeView | null): void {
        if (!this.clipboard || this.clipboard.length === 0) return;
        const target = parent?.childrenBox ? parent : this.rootViews[0];
        if (!target?.childrenBox) return;

        // 粘贴目标若位于剪贴板来源节点自身或其子孙下，会形成循环引用，直接忽略
        // （注意方向：是「来源的子树里是否含目标」，反之会把合法的「粘到原父级」也拦掉）
        if (this.clipboardSources.some((source) => source === target || this.isDescendantView(source, target))) {
            return;
        }

        const plans: FileMovePlan[] = [];
        // 目标目录路径：粘贴时用「目标节点作为目录」的路径作为其中文件的所属目录。
        // 根节点「工作区」是容器而非目录，其路径为 `''`（否则会得到「工作区/x.xlsx」）
        const targetDir = this.containerPathOf(target);
        this.clipboard.forEach((node, index) => {
            // 粘贴保留原名称：仅当目标目录已存在同名节点时才追加序号去重
            const name = this.uniqueName(target, node.name);
            // 源目录取自剪贴板来源视图（此时源节点仍在树上，父链完整）；
            // 复制模式下来源不消失，其路径就是源目录
            const sourceView = this.clipboardSources[index];
            const sourceDir = sourceView ? this.pathOf(sourceView) : '';
            // 连同内容一起搬：复制 → 复制存档；剪切 → 存档改名（名称未变则无需搬）
            this.collectFilePlans(node, sourceDir, targetDir, plans, name);
            this.createFromData({ ...node, name }, target);
        });

        // 内容先于结构删除同步：剪切模式下随后要移除源节点，
        // 若先删源就再也找不到内容的来源键了
        if (plans.length > 0) {
            if (this.clipboardIsCut) this.options.content?.move?.(plans);
            else this.options.content?.copy?.(plans);
        }

        if (this.clipboardIsCut) {
            // 剪切语义：粘贴完成后移除源节点（源已在树上，先粘贴再删除避免丢失数据）；
            // 内容已由内容桥搬到新名字下，因此不再清理存档
            this.removeNodes(this.clipboardSources, false);
        }
        // 无论剪切与否，粘贴后剪贴板即失效（与 Windows 资源管理器一致）
        this.clipboard = null;
        this.clipboardSources = [];
        this.clipboardIsCut = false;
        this.notifyTreeChanged();
    }

    /**
     * 收集「源文件路径 → 目标文件路径」的内容搬迁计划
     *
     * 文件夹只影响其中文件的**路径**，不影响文件名，因此递归展开为其中的文件；
     * 文件夹自身不产生计划项。
     * @private
     * @param {FileNode} source - 源节点数据（来自剪贴板的深拷贝）
     * @param {string} sourceDir - 源节点在其原树中的所属目录路径
     * @param {string} targetDir - 该节点粘贴后的所属目录路径
     * @param {FileMovePlan[]} plans - 计划收集容器
     * @param {string} [targetName] - 顶层节点粘贴后的名称（去重改名时与源名不同）
     */
    private collectFilePlans(
        source: FileNode,
        sourceDir: string,
        targetDir: string,
        plans: FileMovePlan[],
        targetName?: string
    ): void {
        const sourcePath = sourceDir ? `${sourceDir}/${source.name}` : source.name;
        // 顶层节点可能因目标目录同名而被去重改名（`台账.xlsx` → `台账1.xlsx`），
        // 此时目标路径的第一段必须用改名后的名字，否则内容桥会去搬一个不存在的新键
        const name = targetName ?? source.name;
        const targetPath = targetDir ? `${targetDir}/${name}` : name;
        if (Array.isArray(source.children)) {
            // 只有顶层节点会改名：其子孙沿用自身名称，父级路径已由 targetPath 带上
            source.children.forEach((child) =>
                this.collectFilePlans(child, sourcePath, targetPath, plans));
            return;
        }
        plans.push({ from: sourcePath, to: targetPath });
    }

    /**
     * 收集节点（含子孙）中所有文件的**完整路径**（先序）
     *
     * **这是全仓唯一的「搬迁计划」来源**：显示名不含目录，用它构造
     * `{from, to}` 会让不同目录的同名文件互相混淆（历史实现即如此）。
     * 同一次收集的先后顺序固定（先序、深度优先），因此「移动前后各收集一次」
     * 的两份数组可按下标一一对应，从而构造出准确的内容搬迁计划。
     * @private
     * @param {TreeNodeView} view - 起始节点
     * @param {string[]} out - 路径收集容器
     */
    private collectFilePaths(view: TreeNodeView, out: string[] = []): string[] {
        if (view.isFolder) {
            view.children.forEach((child) => this.collectFilePaths(child, out));
            return out;
        }
        const dir = this.pathOf(view);
        const name = view.label.textContent ?? '';
        out.push(dir ? `${dir}/${name}` : name);
        return out;
    }

    /**
     * 计算节点的所属目录路径（不含自身名称）
     *
     * 根节点「工作区」不计入路径 —— 它就是整个树的容器，
     * 因此根下的文件路径即其文件名（如 `sheet1.xlsx`），与「默认空白文件」一致。
     * @private
     * @param {TreeNodeView | null} view - 节点（可为 null，返回空串）
     * @returns {string} 目录路径；根目录下为空串
     */
    private pathOf(view: TreeNodeView | null): string {
        const segments: string[] = [];
        // 从**父节点**起上溯：pathOf 返回的是「所在目录」，不含节点自身名称
        for (let node = view?.parent ?? null; node !== null; node = node.parent) {
            // 根节点是「工作区」，其名称不属于路径
            if (node.parent === null) break;
            segments.unshift(node.label.textContent ?? '');
        }
        return segments.filter(Boolean).join('/');
    }

    /**
     * 计算节点的完整路径（含自身名称）
     * @private
     * @param {TreeNodeView} view - 节点
     * @returns {string} 完整路径
     */
    private pathWithSelf(view: TreeNodeView): string {
        const dir = this.pathOf(view);
        const name = view.label.textContent ?? '';
        return dir ? `${dir}/${name}` : name;
    }

    /**
     * 把节点当作「目录」看待时的路径（用于「粘贴到该目录下」）
     *
     * 与 {@link pathWithSelf} 的唯一区别是**根节点「工作区」返回空串** ——
     * 它是整棵树的容器，不是真实目录，其名称不该出现在任何路径里
     * （否则粘贴到根目录会产出 `工作区/台账.xlsx` 这样的路径）。
     * @private
     * @param {TreeNodeView} view - 目录节点
     * @returns {string} 目录路径；根节点返回空串
     */
    private containerPathOf(view: TreeNodeView): string {
        if (view.parent === null) return '';
        return this.pathWithSelf(view);
    }

    /**
     * 更新节点的「最后修改」元数据（名称/位置变化时调用）
     * @private
     * @param {TreeNodeView} view - 节点
     */
    private touchNode(view: TreeNodeView): void {
        view.lastModifiedAt = Date.now();
        view.lastModifiedBy = LOCAL_USER;
    }

    /**
     * 递归补齐节点元数据（旧存档缺失这些字段时）
     *
     * 旧存档只存 `name / children / open`，因此恢复出来的节点没有创建时间等信息。
     * 这里按「存档里有什么就用什么，缺什么补什么」统一补齐：
     * 缺失时间以**当前时间**为准（无法凭空知道历史时刻，宁可给一个自洽的值，
     * 也不留空让「属性」对话框出现空白项）。
     * @private
     * @param {TreeNodeView} view - 节点
     */
    private ensureMetadata(view: TreeNodeView): void {
        const now = Date.now();
        if (!Number.isFinite(view.createdAt)) view.createdAt = now;
        if (!Number.isFinite(view.lastModifiedAt)) view.lastModifiedAt = view.createdAt;
        if (!view.createdBy) view.createdBy = LOCAL_USER;
        if (!view.lastModifiedBy) view.lastModifiedBy = view.createdBy;
        view.children.forEach((child) => this.ensureMetadata(child));
    }

    /**
     * 判断候选节点是否为给定视图的子孙
     * @private
     * @param {TreeNodeView} view - 候选祖先
     * @param {TreeNodeView} target - 待查找节点
     * @returns {boolean} 是否为子孙
     */
    private isDescendantView(view: TreeNodeView, target: TreeNodeView): boolean {
        return view.children.some((child) => child === target || this.isDescendantView(child, target));
    }

    /**
     * 依据数据递归创建节点（名称已由调用方确定）
     * @private
     * @param {FileNode} data - 节点数据
     * @param {TreeNodeView} parent - 父节点
     * @returns {TreeNodeView} 新建的节点视图
     */
    private createFromData(data: FileNode, parent: TreeNodeView): TreeNodeView {
        return this.createNode(
            Array.isArray(data.children) ? { name: data.name, children: data.children } : { name: data.name },
            parent.childrenBox as HTMLDivElement,
            this.getDepth(parent) + 1,
            parent
        );
    }

    // ════════════════════════════════════════════════════════════
    //  重命名 / 删除
    // ════════════════════════════════════════════════════════════

    /** 开始重命名：行内隐藏标签并插入输入框
     * @private
     * @param {TreeNodeView} view - 目标节点
     */
    private startRename(view: TreeNodeView): void {
        if (this.renaming) this.commitRename(false);

        const input = createInput({
            className: TREE.RENAME_INPUT,
            type: 'text',
            attributes: { spellcheck: 'false', autocomplete: 'off' },
        });
        input.value = view.label.textContent ?? '';
        // 标签隐藏 + 输入框追加（而非 replaceChild）：标签引用始终留在行内，
        // 提交/取消时只需移除输入框并取消隐藏，不会丢失节点引用
        view.label.hidden = true;
        view.row.appendChild(input);
        this.renaming = { view, input };

        input.addEventListener('blur', () => this.commitRename(true));
        input.addEventListener('keydown', (e) => {
            e.stopPropagation();   // 避免按键冒泡触发搜索等全局行为
            if (e.key === 'Enter') this.commitRename(true);
            if (e.key === 'Escape') this.commitRename(false);
        });
        input.focus();
        input.select();
    }

    /**
     * 提交 / 取消重命名
     * @param {boolean} apply - 是否应用新名称（false 为取消）
     * @private
     */
    private commitRename(apply: boolean): void {
        const pending = this.renaming;
        if (!pending) return;
        this.renaming = null;
        const { view, input } = pending;
        const next = input.value.trim();

        // 输入框移除、标签恢复
        input.remove();
        view.label.hidden = false;

        if (!apply || !next || next === view.label.textContent) return;

        // 同级重名校验：与其它兄弟节点冲突时保持原名
        const siblings = view.parent?.children.filter((child) => child !== view) ?? [];
        if (siblings.some((child) => child.key === next.toLowerCase())) return;

        const oldName = view.label.textContent ?? '';
        // 改名后「当前文件」的高亮与定位要跟着走：先判断被改名的就是主页面当前显示的文件
        // （要在改动 label / key 之前判断，否则查找条件已经变了）
        const wasActived = !!this.activedFileName && this.findByName(this.activedFileName) === view;
        // 路径搬迁计划必须在改名（及其后续）**之前**采集：这时父链与旧名都还在
        const oldPaths = this.collectFilePaths(view);

        view.label.textContent = next;
        view.key = next.toLowerCase();
        const newPaths = this.collectFilePaths(view);

        if (wasActived) {
            // 当前文件可能就是被改名的文件本身，也可能是它（作为文件夹）的某个子孙
            for (let i = 0; i < oldPaths.length; i++) {
                if (stripExtension(oldPaths[i]) === stripExtension(this.activedFileName)) {
                    this.activedFileName = newPaths[i];
                    break;
                }
            }
            if (wasActived && !view.isFolder) this.activedFileName = next;
        }

        // 内容键 = 路径，因此改名必须把存档搬到新路径下，否则改名后的文件是空表。
        // 文件夹改名会改变其下**所有**文件的路径，故一次提交全部计划（递归展开）。
        const plans: FileMovePlan[] = [];
        for (let i = 0; i < oldPaths.length; i++) {
            if (oldPaths[i] !== newPaths[i]) plans.push({ from: oldPaths[i], to: newPaths[i] });
        }
        if (plans.length > 0) this.options.content?.move?.(plans);

        this.touchNode(view);
        this.refreshActivedHighlight();
        this.notifyTreeChanged();
    }

    /**
     * 删除节点（含子孙），并清理选中、重命名状态与对应的文件内容
     *
     * 幂等：对同一节点重复调用安全（已从树上摘除的节点直接跳过），
     * 便于剪切粘贴等场景无需额外判断节点是否仍在树上
     * @private
     * @param {TreeNodeView[]} views - 待删除节点集合
     * @param {boolean} [purgeContent=true] - 是否清理这些文件在数据库中的内容；
     *   剪切粘贴移除源节点时为 false（内容已随粘贴搬走，不能再删）
     */
    private removeNodes(views: TreeNodeView[], purgeContent = true): void {
        let removed = false;
        const paths: string[] = [];
        views.forEach((view) => {
            // 根节点（parent 为 null）受保护，不可删除：删除后树将失去唯一入口
            if (this.isRootView(view)) return;

            // 已从父节点 children 摘除说明此前已删除，跳过避免误删同位置的新节点
            const siblings = view.parent?.children;
            const index = siblings ? siblings.indexOf(view) : -1;
            if (siblings && index < 0) return;

            // 路径必须在摘除**之前**收集：摘除后父链断开，路径就无法再推导
            this.collectFilePaths(view, paths);
            if (siblings) siblings.splice(index, 1);
            if (this.selected === view) this.selected = null;
            if (this.renaming?.view === view) this.renaming = null;
            view.el.remove();
            removed = true;
        });
        if (removed) {
            // 删除的是主页面当前显示的文件时，其节点已不存在 → 高亮随之取消
            this.refreshActivedHighlight();
            if (purgeContent && paths.length > 0) this.options.content?.remove?.(paths);
            this.updateEmptyState();
            this.notifyTreeChanged();
        }
    }

    /**
     * 判断节点是否为根节点
     * @private
     * @param {TreeNodeView} view - 节点
     * @returns {boolean} 是否根节点
     */
    private isRootView(view: TreeNodeView): boolean {
        return this.rootViews.includes(view);
    }

    /**
     * 选中节点
     * @private
     * @param {TreeNodeView | null} view - 目标节点
     */
    private select(view: TreeNodeView | null): void {
        if (this.selected === view) return;
        if (this.selected) {
            this.selected.row.classList.remove(TREE.SELECTED);
        }
        this.selected = view;
        if (view) {
            view.row.classList.add(TREE.SELECTED);
        }
    }

    /**
     * 设置展开状态并渲染（同时更新用户记忆状态）
     * @private
     * @param {TreeNodeView} view - 节点
     * @param {boolean} isOpen - 是否展开
     */
    private setOpen(view: TreeNodeView, isOpen: boolean): void {
        view.isOpen = isOpen;
        this.renderOpen(view, isOpen);
    }

    /**
     * 仅渲染展开状态（不修改用户记忆状态，供搜索期间使用）
     * @private
     * @param {TreeNodeView} view - 节点
     * @param {boolean} isOpen - 是否展开
     */
    private renderOpen(view: TreeNodeView, isOpen: boolean): void {
        if (!view.isFolder || !view.childrenBox) return;
        view.childrenBox.hidden = !isOpen;
        view.row.classList.toggle(TREE.OPEN, isOpen);
        view.icon.className = isOpen ? 'icon-folder-open' : 'icon-folder';
    }

    /**
     * 先序遍历子树
     * @private
     * @param {TreeNodeView} view - 起始节点
     * @param {(node: TreeNodeView) => void} visit - 访问回调
     */
    private walk(view: TreeNodeView, visit: (node: TreeNodeView) => void): void {
        visit(view);
        view.children.forEach((child) => this.walk(child, visit));
    }

    /**
     * 展开全部文件夹
     */
    public expandAll(): void {
        let changed = false;
        this.rootViews.forEach((view) => this.walk(view, (node) => {
            if (node.isFolder && !node.isOpen) {
                this.setOpen(node, true);
                changed = true;
            }
        }));
        // 整批展开只落库一次，避免逐节点写入
        if (changed) this.notifyTreeChanged();
    }

    /**
     * 收起全部文件夹
     */
    public collapseAll(): void {
        let changed = false;
        this.rootViews.forEach((view) => this.walk(view, (node) => {
            if (node.isFolder && node.isOpen) {
                this.setOpen(node, false);
                changed = true;
            }
        }));
        if (changed) this.notifyTreeChanged();
    }

    /**
     * 设置搜索关键字（供外部驱动搜索，与输入框输入等价）
     * @param {string} keyword - 关键字
     */
    public search(keyword: string): void {
        if (this.searchInput.value !== keyword) {
            this.searchInput.value = keyword;
        }
        this.applyKeyword();
    }

    /**
     * 应用关键字过滤
     * @private
     */
    private applyKeyword(): void {
        this.keyword = this.searchInput.value.trim().toLowerCase();

        if (!this.keyword) {
            this.rootViews.forEach((view) => this.resetFilter(view));
        } else {
            this.rootViews.forEach((view) => this.filterNode(view));
        }

        this.updateEmptyState();
    }

    /**
     * 恢复节点：全部可见，并按用户记忆的展开状态渲染
     * @private
     * @param {TreeNodeView} view - 节点
     */
    private resetFilter(view: TreeNodeView): void {
        view.el.hidden = false;
        this.renderOpen(view, view.isOpen);
        view.children.forEach((child) => this.resetFilter(child));
    }

    /**
     * 递归过滤节点：祖先因子孙命中而保留可见
     * @private
     * @param {TreeNodeView} view - 节点
     * @returns {boolean} 本节点（含子孙）是否命中
     */
    private filterNode(view: TreeNodeView): boolean {
        const selfMatched = view.key.includes(this.keyword);
        let childMatched = false;
        view.children.forEach((child) => {
            if (this.filterNode(child)) childMatched = true;
        });

        const matched = selfMatched || childMatched;
        view.el.hidden = !matched;
        // 子孙命中时自动展开以便直接看到结果；仅自身命中时保持用户此前的展开状态
        this.renderOpen(view, childMatched ? true : view.isOpen);

        return matched;
    }

    /**
     * 更新无匹配提示的显隐
     * @private
     */
    private updateEmptyState(): void {
        const hasVisible = this.rootViews.some((view) => !view.el.hidden);
        this.empty.hidden = hasVisible;
    }

/**
     * 定位
     *
     * 目标为**当前主页面显示的文件**（由 {@link setActivedFileName} 同步，
     * 通常取自 `data.activedSheetName`）在树中的位置：展开其全部祖先、
     * 滚动到可视区域并闪烁高亮。
     *
     * 降级顺序：主页面文件 → 搜索首个命中项 → 当前选中项 → 第一个根节点。
     */
    public locate(): void {
        const target = this.getLocateTarget();
        if (!target) return;

        let opened = false;
        for (let parent = target.parent; parent !== null; parent = parent.parent) {
            if (parent.isFolder && !parent.isOpen) {
                this.setOpen(parent, true);
                opened = true;
            }
        }
        // 定位会展开祖先以便看到目标文件：一并落库，刷新后目标仍可见
        if (opened) this.notifyTreeChanged();

        this.select(target);
        target.row.scrollIntoView({ block: 'nearest' });
        this.flash(target);
    }

    /**
     * 设置当前主页面显示的文件名（「定位」与高亮据此在树中查找）
     *
     * 可传显示名（`销售数据.xlsx`）或工作表名（`销售数据`），
     * 匹配时会容忍「有无扩展名」的差异（见 {@link findByName}）。
     * @param {string} name - 文件名或工作表名
     */
    public setActivedFileName(name: string): void {
        this.activedFileName = name ?? '';
        this.refreshActivedHighlight();
    }

    /**
     * 获取当前主页面显示的文件名
     * @returns {string} 文件名（可能是显示名或工作表名，取决于外部最后一次同步的值）
     */
    public getActivedFileName(): string {
        return this.activedFileName;
    }

    /**
     * 判断树中是否存在该文件节点（文件夹不算）
     *
     * 供外部在「删除当前文件后回退」等场景确认回退目标是否还留在树里。
     * 参数可为**路径**（`归档/销售数据.xlsx`）、显示名（`销售数据.xlsx`）
     * 或工作表名（`销售数据`）——匹配见 {@link findByName}。
     * @param {string} name - 文件路径 / 显示名 / 工作表名
     * @returns {boolean} 是否存在对应的文件节点
     */
    public hasFile(name: string): boolean {
        const found = this.findByName(name);
        return !!found && !found.isFolder;
    }

    /**
     * 列出树中所有文件的**完整路径**（先序：从上到下、从浅到深）
     *
     * 返回路径而不是文件名：路径才是内容存档键的来源，且不同目录的同名文件
     * 是两个独立文件（返回文件名会让它们无法区分）。
     * 供外部在「打开历史都用不上」时挑一个仍在树中的文件作为回退目标。
     * @returns {string[]} 文件路径集合（不含文件夹）
     */
    public listFiles(): string[] {
        const out: string[] = [];
        this.rootViews.forEach((view) => this.collectFilePaths(view, out));
        return out;
    }

    /**
     * 计算根目录「工作区」下下一个可用的新文件名
     *
     * 与右键「新建文件」共用同一套命名规则（{@link availableFileName}），
     * 因此默认得到 `sheet1.xlsx`（同名被占用时递增序号）。
     * 供外部在「刷新后树里一个文件都没有」时补一个默认空白文件用。
     * @param {string} [todo='XLSX'] - 文件类型标识（XLSX / DOCX / PPTX）
     * @returns {string} 可用的文件显示名（含扩展名）
     */
    public nextFileName(todo: string = 'XLSX'): string {
        return this.availableFileName(todo, null);
    }

    /**
     * 刷新「当前主页面显示的文件」的常驻高亮
     *
     * 高亮与「选中」（{@link TREE}.SELECTED，正在操作的节点）、
     * 「定位闪烁」（{@link TREE}.LOCATED，一次性动效）相互独立：
     * 选中是操作目标，高亮是主页面正在显示的文件，两者可能不是同一个节点。
     * 找不到对应节点（例如文件刚被删除）时不残留任何高亮。
     * @private
     */
    private refreshActivedHighlight(): void {
        const found = this.activedFileName ? this.findByName(this.activedFileName) : null;
        const target = found && !found.isFolder ? found : null;
        this.rootViews.forEach((view) => this.walk(view, (node) => {
            node.row.classList.toggle(TREE.ACTIVED, node === target);
        }));
    }

    /**
     * 登记「当前主页面打开的文件」并在树中显示
     *
     * 供打开文件 / 页面刷新恢复时调用：树中已有同名节点则直接复用（不重复登记），
     * 否则追加到根节点「工作区」下；同时同步「定位」的目标名称并高亮该节点。
     * @param {string} displayName - 文件显示名（含扩展名，如 `sheet1.xlsx`）
     * @returns {TreeNodeView | null} 对应的节点视图
     */
    public showActivedFile(displayName: string): TreeNodeView | null {
        const name = (displayName ?? '').trim();
        if (!name) return null;
        this.activedFileName = name;

        const existing = this.findByName(name);
        const view = existing && !existing.isFolder ? existing : this.addNode(name);
        this.refreshActivedHighlight();
        return view;
    }

    /**
     * 按路径 / 名称在整棵树中查找节点
     *
     * 匹配规则（不区分大小写，按优先级）：
     * 1. **完整路径一致**（`归档/销售数据.xlsx`）—— 最精确，可区分同名文件；
     * 2. 名称完全一致（`销售数据.xlsx`）；
     * 3. 「去掉扩展名后」一致 —— 树中节点是 `sheet1.xlsx`，而主页面记录的是
     *    工作表的键（IndexDB 的存储键即路径去扩展名，见 `FileNode.path`），
     *    两者必须能互相定位。
     *
     * 同名时文件优先于文件夹。路径命中优先于名称命中：否则两个不同目录下的
     * 同名文件会定位到第一个（错的）那个。
     * @private
     * @param {string} name - 文件路径 / 文件名 / 工作表名
     * @returns {TreeNodeView | null} 命中的节点
     */
    private findByName(name: string): TreeNodeView | null {
        const key = name.trim().toLowerCase();
        if (!key) return null;
        const bare = stripExtension(key);
        const byPath: TreeNodeView[] = [];
        const byName: TreeNodeView[] = [];
        const fallback: TreeNodeView[] = [];
        this.rootViews.forEach((root) => this.walk(root, (view) => {
            const path = this.pathWithSelf(view).toLowerCase();
            if (path === key) byPath.push(view);
            if (view.key === key || stripExtension(view.key) === bare) byName.push(view);
            if (stripExtension(path) === bare) fallback.push(view);
        }));
        // 1) 路径精确命中（去掉扩展名也允许）→ 2) 名称命中（文件优先）→ 3) 路径去扩展名
        const exactPath = byPath.find((v) => !v.isFolder) ?? byPath[0];
        if (exactPath) return exactPath;
        return byName.find((v) => !v.isFolder) ?? byName[0]
            ?? fallback.find((v) => !v.isFolder) ?? fallback[0] ?? null;
    }

    /**
     * 获取定位目标节点
     * @private
     * @returns {TreeNodeView | null} 目标节点
     */
    private getLocateTarget(): TreeNodeView | null {
        // 首选：当前主页面显示的文件
        if (this.activedFileName) {
            const current = this.findByName(this.activedFileName);
            if (current) return current;
        }
        // 搜索状态下定位首个命中项
        if (this.keyword) {
            const matched = this.findFirstMatch(this.rootViews);
            if (matched) return matched;
        }
        return this.selected ?? this.rootViews[0] ?? null;
    }

    /**
     * 查找首个命中的可见节点（先序）
     * @private
     * @param {TreeNodeView[]} views - 节点集合
     * @returns {TreeNodeView | null} 命中节点
     */
    private findFirstMatch(views: TreeNodeView[]): TreeNodeView | null {
        for (const view of views) {
            if (!view.el.hidden && view.key.includes(this.keyword)) return view;
            const inChildren = this.findFirstMatch(view.children);
            if (inChildren) return inChildren;
        }
        return null;
    }

    /**
     * 定位反馈：闪烁高亮
     * @private
     * @param {TreeNodeView} view - 目标节点
     */
    private flash(view: TreeNodeView): void {
        if (this.flashTimer !== null) {
            window.clearTimeout(this.flashTimer);
        }
        view.row.classList.remove(TREE.LOCATED);
        // 读取布局属性强制重排，保证连续定位时动画可以重新播放
        void view.row.offsetWidth;
        view.row.classList.add(TREE.LOCATED);
        this.flashTimer = window.setTimeout(() => {
            view.row.classList.remove(TREE.LOCATED);
            this.flashTimer = null;
        }, 900);
    }

    /**
     * 新增节点（默认挂在根节点「工作区」下）
     *
     * 用于「打开文件」后把文件名称登记到树中；同层同名节点会被自动去重为 `名称 (2)`
     * @param {string} name - 节点名称
     * @param {{ isFolder?: boolean, parent?: TreeNodeView | null, select?: boolean }} [options={}] - 附加选项
     * @returns {TreeNodeView | null} 新建的节点视图；无可用父节点时为 null
     */
    public addNode(
        name: string,
        options: { isFolder?: boolean; parent?: TreeNodeView | null; select?: boolean } = {}
    ): TreeNodeView | null {
        const { isFolder = false, parent, select = false } = options;
        // 父节点缺省取根节点；父节点非文件夹（无子容器）时同样回退到根节点
        const target = parent?.childrenBox ? parent : this.rootViews[0];
        if (!target?.childrenBox) return null;

        const uniqueName = this.uniqueName(target, name);
        const view = this.createNode(
            isFolder ? { name: uniqueName, children: [] } : { name: uniqueName },
            target.childrenBox,
            this.getDepth(target) + 1,
            target
        );
        // 新建的文件夹默认展开（空文件夹），避免创建后还需再点一次才能看到结构
        if (view.isFolder) this.setOpen(view, true);

        // 新增文件即同步数据库：库中还没有该文件的存档时补一份空白存档，
        // 使「树里看到的文件」与「库里存的内容」始终一一对应（只补不覆盖）。
        // 传**路径**而不是文件名：不同目录的同名文件是各自独立的存档
        if (!view.isFolder) this.options.content?.create?.([this.pathWithSelf(view)]);

        // 展开祖先并选中，保证新增节点可见
        for (let p = view.parent; p !== null; p = p.parent) {
            if (p.isFolder) this.setOpen(p, true);
        }
        this.select(view);
        if (select) this.flash(view);
        this.updateEmptyState();
        this.notifyTreeChanged();
        return view;
    }

    /**
     * 在指定父节点下生成不重复的名称（粘贴 / 登记文件时用于兜底去重）
     *
     * 重名规则与「新建文件」{@link availableFileName} 保持一致：
     * - 文件名中带数字（`sheet1.xlsx` / `doc1.docx`）→ **递增其中的数字**：
     *   复制 `doc1.docx` 得到 `doc2.docx`（而不是 `doc11.docx`）；
     * - 文件名中无数字（`报表.xlsx`）→ 在扩展名前追加序号：`报表1.xlsx`；
     * - 文件夹名（`新建文件夹` / `归档`）→ 追加序号：`归档1`。
     * @private
     * @param {TreeNodeView} parent - 父节点
     * @param {string} name - 期望名称
     * @returns {string} 去重后的名称
     */
    private uniqueName(parent: TreeNodeView, name: string): string {
        const taken = new Set(parent.children.map((child) => child.key));
        if (!taken.has(name.toLowerCase())) return name;

        const dot = name.lastIndexOf('.');
        const base = dot > 0 ? name.slice(0, dot) : name;
        const suffix = dot > 0 ? name.slice(dot) : '';
        // 带扩展名的文件名：若名称里已有数字则从该数字起递增，否则追加序号
        const matched = dot > 0 ? /^(.*?)(\d+)$/.exec(base) : null;
        const prefix = matched ? matched[1] : base;
        const start = matched ? Number(matched[2]) : 1;
        for (let seq = start; seq < start + 10000; seq++) {
            const candidate = `${prefix}${seq}${suffix}`;
            if (!taken.has(candidate.toLowerCase())) return candidate;
        }
        return name;
    }

    /**
     * 计算节点深度（根节点为 0）
     * @private
     * @param {TreeNodeView} view - 节点
     * @returns {number} 深度
     */
    private getDepth(view: TreeNodeView): number {
        let depth = 0;
        for (let p = view.parent; p !== null; p = p.parent) depth++;
        return depth;
    }

    /**
     * 获取页面元素集合
     * @returns {{ container: HTMLDivElement, toolbar: HTMLDivElement, body: HTMLDivElement }} 元素集合
     */
    public getElements(): { container: HTMLDivElement; toolbar: HTMLDivElement; body: HTMLDivElement } {
        return { container: this.container, toolbar: this.toolbar, body: this.body };
    }

    /**
     * 销毁组件，清理定时器与 DOM
     */
    public destroy(): void {
        if (this.flashTimer !== null) {
            window.clearTimeout(this.flashTimer);
            this.flashTimer = null;
        }
        this.selected = null;
        this.rootViews = [];
        this.container.remove();
    }
}
