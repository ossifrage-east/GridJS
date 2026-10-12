/*  # 通用文件树组件 TS 文件（「文件」导航页）
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { createDiv, createSpan, createInput, createButton, createElement } from "../../utils/dom";
import { stripExtension } from "../../utils/fileName";
import { FILE_TREE_CLASS_NAMES as TREE, MENU_CLASS_NAMES } from "../constant";
import { Menu, MenuContent } from "./menu";
import '../style/fileTree.scss';

/**
 * 文件节点数据接口
 * @interface FileNode
 */
export interface FileNode {
    /** 节点名称（显示文本，同时用于搜索匹配） */
    name: string;
    /** 子节点：存在即视为文件夹（空数组表示空文件夹） */
    children?: FileNode[];
    /**
     * 文件夹展开状态（仅文件夹会带此字段，仅供持久化还原使用）
     *
     * 缺省时按默认规则处理：根层展开、更深层收起（见 {@link createNode}）。
     */
    open?: boolean;
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
 * 树只认识**显示名**，而文件内容存在外部（IndexedDB 的 `values` 库，键为去扩展名的工作表名）。
 * 因此「某个文件的名字变了 / 被复制了 / 被删除了」这类副作用一律以显示名告知外部，
 * 由外部换算成存档键后执行。
 * @interface FileMovePlan
 */
export interface FileMovePlan {
    /** 源文件显示名（含扩展名） */
    from: string;
    /** 目标文件显示名（含扩展名） */
    to: string;
}

/**
 * 内容桥：把「结构变化」翻译成「文件内容的同步动作」
 *
 * 四个动作都由树在**语义发生变化的时刻**调用一次，调用方负责真正读写数据：
 * - {@link FileContentBridge.create}：新建文件（库中尚无存档时补一份空白存档）；
 * - {@link FileContentBridge.copy}：复制 / 粘贴（复制模式），含文件夹递归；
 * - {@link FileContentBridge.move}：改名（重命名、剪切粘贴导致的改名）；
 * - {@link FileContentBridge.remove}：删除（含文件夹递归）。
 * @interface FileContentBridge
 */
export interface FileContentBridge {
    /**
     * 补建文件存档：新建出来的文件在数据库中还没有存档时写入一份空白存档
     *
     * 只补不覆盖：这条路径也会在「登记已打开的文件」时被调用，此时内容刚由打开器写入。
     */
    create?: (names: string[]) => void;
    /** 复制文件内容：把 `from` 的存档复制到 `to` */
    copy?: (plans: FileMovePlan[]) => void;
    /** 搬迁文件内容：把 `from` 的存档搬到 `to`（源键不再保留） */
    move?: (plans: FileMovePlan[]) => void;
    /** 删除文件内容：清理这些显示名对应文件在数据库中的存档 */
    remove?: (names: string[]) => void;
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
     * 双击文件节点时触发（参数为节点显示名，含扩展名）
     *
     * 由外部装载该文件内容到主页面；若主页面已显示该文件，外部应保持内容不变。
     */
    onActivate?: (fileName: string) => void;
    /**
     * 树结构、名称或展开状态发生变化时触发（新增 / 删除 / 重命名 / 粘贴 / 展开收起）
     *
     * 用于外部持久化：把回调参数存起来，页面刷新时再通过
     * {@link FileTreeOptions.nodes} 注入还原，即可让树回到刷新前的样子。
     * 参数为当前整棵树的深拷贝（纯数据，可直接交给结构化克隆存储）。
     */
    onChange?: (nodes: FileNode[]) => void;
    /**
     * 内容桥：复制 / 改名 / 删除文件时同步数据库中的文件内容
     *
     * 不传则树只做结构操作（粘贴出来的文件是空表、删除不清理存档）。
     */
    content?: FileContentBridge;
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
    /** 节点行：选中、定位闪烁的作用元素 */
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
}

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
 * - 内容同步：新增 / 复制 / 改名 / 删除文件会通过 {@link FileTreeOptions.content}
 *   通知外部同步数据库中该文件的内容（树只认识显示名，不接触存档数据）；
 *   新增只「补建」不覆盖，因此刷新恢复、打开文件登记等路径同样安全。
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
        // 双击文件节点：由外部把该文件内容装载到主页面
        // （文件夹不处理：展开 / 收起已由单击负责，双击会连点两次反而相互抵消）
        row.addEventListener('dblclick', () => {
            if (view.isFolder) return;
            this.options.onActivate?.(view.label.textContent ?? '');
        });
        // 右键：在鼠标位置弹出上下文菜单（选中被右键的节点作为操作目标）
        row.addEventListener('contextmenu', (e) => {
            e.preventDefault();
            e.stopPropagation();
            this.openContextMenu(e, view);
        });
        parentElement.appendChild(el);
        return view;
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
     * 菜单结构：新建文件（hover 展开六项二级菜单）/ 新建文件夹 / 分割线 /
     * 剪切 / 复制 / 粘贴 / 分割线 / 重命名 / 删除。
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
                // 新建文件：hover 展开六项二级菜单
                () => this.createNewFileSubmenuItem(view),
                { todo: 'newFolder', icon: 'icon-new-folder', text: '新建文件夹' },
                { todo: 'open', icon: 'icon-open', text: '打开' },
                'separator',
                { todo: 'cut', icon: 'icon-scissors', text: '剪切' },
                { todo: 'copy', icon: 'icon-copy', text: '复制' },
                { todo: 'paste', icon: 'icon-paste', text: '粘贴' },
                'separator',
                { todo: 'rename', icon: 'icon-text-height', text: '重命名' },
                { todo: 'remove', icon: 'icon-bin', text: '删除' },
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
     * @param {boolean} [withOpen=false] - 是否附带文件夹展开状态（持久化用；剪贴板不需要）
     * @returns {FileNode[]} 深拷贝后的数据
     */
    private serializeNodes(views: TreeNodeView[], withOpen = false): FileNode[] {
        return views.map((view) => ({
            name: view.label.textContent ?? '',
            ...(view.isFolder
                ? { children: view.children.map((child) => this.serializeNodes([child], withOpen)[0]) }
                : {}),
            ...(withOpen && view.isFolder ? { open: view.isOpen } : {}),
        }));
    }

    /**
     * 导出当前整棵树（深拷贝，可直接交给结构化克隆持久化）
     *
     * 与剪贴板用的 {@link serializeNodes} 同源，额外带上文件夹展开状态，
     * 使刷新还原后「原来折叠的文件夹依然是折叠的」。
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
        this.clipboard.forEach((node) => {
            // 粘贴保留原名称：仅当目标目录已存在同名节点时才追加序号去重
            const name = this.uniqueName(target, node.name);
            // 连同内容一起搬：复制 → 复制存档；剪切 → 存档改名（名称未变则无需搬）
            this.collectFilePlans(node, name, plans);
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
     * 收集「源文件 → 目标文件」的内容搬迁计划
     *
     * 文件夹只影响其中文件的**路径**，不影响文件名，因此递归展开为其中的文件；
     * 文件夹自身不产生计划项。
     * @private
     * @param {FileNode} source - 源节点数据（来自剪贴板的深拷贝）
     * @param {string} targetName - 该节点粘贴后的名称
     * @param {FileMovePlan[]} plans - 计划收集容器
     */
    private collectFilePlans(source: FileNode, targetName: string, plans: FileMovePlan[]): void {
        if (Array.isArray(source.children)) {
            source.children.forEach((child) => this.collectFilePlans(child, child.name, plans));
            return;
        }
        plans.push({ from: source.name, to: targetName });
    }

    /**
     * 收集节点（含子孙）中所有文件的显示名
     * @private
     * @param {TreeNodeView} view - 起始节点
     * @param {string[]} out - 名称收集容器
     */
    private collectFileNames(view: TreeNodeView, out: string[]): void {
        if (view.isFolder) {
            view.children.forEach((child) => this.collectFileNames(child, out));
            return;
        }
        out.push(view.label.textContent ?? '');
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

        view.label.textContent = next;
        view.key = next.toLowerCase();
        if (wasActived) this.activedFileName = next;

        // 文件改名后内容也要跟着搬：否则存档仍留在旧键上，改名后的文件看起来是空的
        // （文件夹改名只改变其中文件的路径，内容键是去扩展名的文件显示名，无需搬迁）
        if (!view.isFolder && oldName && oldName !== next) {
            this.options.content?.move?.([{ from: oldName, to: next }]);
        }
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
        const names: string[] = [];
        views.forEach((view) => {
            // 根节点（parent 为 null）受保护，不可删除：删除后树将失去唯一入口
            if (this.isRootView(view)) return;

            // 已从父节点 children 摘除说明此前已删除，跳过避免误删同位置的新节点
            const siblings = view.parent?.children;
            const index = siblings ? siblings.indexOf(view) : -1;
            if (siblings && index < 0) return;

            if (siblings) siblings.splice(index, 1);
            this.collectFileNames(view, names);
            if (this.selected === view) this.selected = null;
            if (this.renaming?.view === view) this.renaming = null;
            view.el.remove();
            removed = true;
        });
        if (removed) {
            // 删除的是主页面当前显示的文件时，其节点已不存在 → 高亮随之取消
            this.refreshActivedHighlight();
            if (purgeContent && names.length > 0) this.options.content?.remove?.(names);
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
     * @param {string} name - 文件名或工作表名
     * @returns {boolean} 是否存在对应的文件节点
     */
    public hasFile(name: string): boolean {
        const found = this.findByName(name);
        return !!found && !found.isFolder;
    }

    /**
     * 列出树中所有文件的显示名（先序：从上到下、从浅到深）
     *
     * 供外部在「打开历史都用不上」时挑一个仍在树中的文件作为回退目标。
     * @returns {string[]} 文件显示名集合（不含文件夹）
     */
    public listFiles(): string[] {
        const out: string[] = [];
        this.rootViews.forEach((view) => this.collectFileNames(view, out));
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
     * 按名称在整棵树中查找节点
     *
     * 匹配规则（不区分大小写）：
     * - 名称完全一致；
     * - 或「去掉扩展名后」一致 —— 树中节点是 `sheet1.xlsx`，而主页面记录的是
     *   工作表名 `sheet1`（IndexDB 的存储键即工作表名），两者必须能互相定位。
     * 同名时文件优先于文件夹。
     * @private
     * @param {string} name - 文件名或工作表名
     * @returns {TreeNodeView | null} 命中的节点
     */
    private findByName(name: string): TreeNodeView | null {
        const key = name.trim().toLowerCase();
        if (!key) return null;
        const bare = stripExtension(key);
        const fallback: TreeNodeView[] = [];
        const walk = (views: TreeNodeView[]): TreeNodeView | null => {
            for (const view of views) {
                if (view.key === key || stripExtension(view.key) === bare) {
                    // 文件优先：同名文件夹可能被当前文件之外的内容占用
                    if (!view.isFolder) return view;
                    fallback.push(view);
                }
                const found = walk(view.children);
                if (found) return found;
            }
            return null;
        };
        return walk(this.rootViews) ?? fallback[0] ?? null;
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
        // 使「树里看到的文件」与「库里存的内容」始终一一对应（只补不覆盖）
        if (!view.isFolder) this.options.content?.create?.([view.label.textContent ?? uniqueName]);

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
