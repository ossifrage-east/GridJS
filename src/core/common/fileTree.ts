/*  # 通用文件树组件 TS 文件（「文件」导航页）
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { createDiv, createSpan, createInput, createButton, createElement } from "../../utils/dom";
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
}

/**
 * 「新建文件」下拉菜单支持的扩展名
 * 与 FileOpener.SUPPORTED_EXTENSIONS 保持一致（此处为展示用的菜单项文案）
 */
const NEW_FILE_MENU_ITEMS = [
    { todo: 'XLSX', text: '新建XLSX' },
    { todo: 'XLS', text: '新建XLS' },
    { todo: 'DOCX', text: '新建DOCX' },
    { todo: 'DOC', text: '新建DOC' },
    { todo: 'PPTX', text: '新建PPTX' },
    { todo: 'PPT', text: '新建PPT' },
];

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
 * - 右键菜单：新建文件（hover 展开六项二级菜单）/ 新建文件夹 / 打开 /
 *   分割线 / 剪切 / 复制 / 粘贴 / 分割线 / 重命名 / 删除。
 *   右键文件夹时「新建」作用于该文件夹，右键文件时作用于其所在文件夹。
 * - 命名规则：新建时**仅当当前目录内同名冲突才追加序号** ——
 *   无冲突直接用基础名（`新建文件夹` / `新建文件.xlsx`），
 *   冲突时递增（`新建文件夹1`、`新建文件夹2` …）。
 *   粘贴保留原文件名/夹名，仅在目标目录冲突时追加序号去重。
 * - 删除保护：根节点（「工作区」）不可删除。
 *
 * @class FileTree
 * @example
 * const fileTree = new FileTree({ menu, onOpen: () => opener.pickAndOpen() });
 * navigator.navContentItemMount('文件', fileTree.getElements().container);
 * // 打开文件成功后登记节点，并同步主页面文件名供「定位」使用
 * fileTree.addNode('销售数据.xlsx');
 * fileTree.setActivedFileName('销售数据');
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
     * @param {string} extension - 文件扩展名（大写，如 'XLSX'）
     * @param {TreeNodeView | null} parent - 父节点（缺省为根节点「工作区」）
     */
    private createFileNode(extension: string, parent: TreeNodeView | null = null): void {
        this.addNode(
            this.availableName('新建文件', parent, extension.toLowerCase()),
            { select: true, parent }
        );
    }

    /**
     * 生成当前目录内可用的名称
     *
     * 规则：**仅当当前目录已存在同名节点时才追加序号** —— 目录内无冲突则原样返回，
     * 如「新建文件夹」→「新建文件夹1」→「新建文件夹2」…
     *
     * 与「删除后序号继续递增」不同，这里以目录实际内容为准，
     * 因此删除「新建文件夹1」后再次新建会重新得到「新建文件夹1」，
     * 始终保证「名称不重复」且「序号从冲突处开始」。
     * @private
     * @param {string} base - 基础名称（如 '新建文件'）
     * @param {TreeNodeView | null} parent - 父节点（缺省为根节点）
     * @param {string} [extension] - 文件扩展名（不含点）
     * @returns {string} 当前目录内不冲突的名称
     */
    private availableName(base: string, parent: TreeNodeView | null = null, extension?: string): string {
        const target = parent?.childrenBox ? parent : this.rootViews[0];
        const suffix = extension ? `.${extension}` : '';
        const taken = new Set(target?.children.map((child) => child.key) ?? []);

        // 无冲突：直接用基础名（不追加任何序号）
        const plain = `${base}${suffix}`;
        if (!taken.has(plain.toLowerCase())) return plain;

        // 有冲突：从 1 开始递增，跳过已占用者
        for (let seq = 1; seq < 10000; seq++) {
            const candidate = `${base}${seq}${suffix}`;
            if (!taken.has(candidate.toLowerCase())) return candidate;
        }
        return plain;
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
            // 默认展开第一层（根节点），更深的层级收起
            this.setOpen(view, depth < 1);
        }

        // 登记到父节点的子节点集合：新增节点（打开文件 / 新建）同样依赖此登记，
        // 否则搜索过滤、展开收起与名称去重都会漏掉该节点
        if (parent) parent.children.push(view);

        row.addEventListener('click', () => this.handleRowClick(view));
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
     * 将视图序列化为纯数据（供剪贴板暂存，切断与 DOM 的引用）
     * @private
     * @param {TreeNodeView[]} views - 节点集合
     * @returns {FileNode[]} 深拷贝后的数据
     */
    private serializeNodes(views: TreeNodeView[]): FileNode[] {
        return views.map((view) => ({
            name: view.label.textContent ?? '',
            ...(view.isFolder ? { children: view.children.map((child) => this.serializeNodes([child])[0]) } : {}),
        }));
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

        this.clipboard.forEach((node) => {
            // 粘贴保留原名称：仅当目标目录已存在同名节点时才追加序号去重
            this.createFromData({ ...node, name: this.uniqueName(target, node.name) }, target);
        });

        if (this.clipboardIsCut) {
            // 剪切语义：粘贴完成后移除源节点（源已在树上，先粘贴再删除避免丢失数据）
            this.removeNodes(this.clipboardSources);
        }
        // 无论剪切与否，粘贴后剪贴板即失效（与 Windows 资源管理器一致）
        this.clipboard = null;
        this.clipboardSources = [];
        this.clipboardIsCut = false;
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

        view.label.textContent = next;
        view.key = next.toLowerCase();
    }

    /**
     * 删除节点（含子孙），并清理选中与重命名状态
     *
     * 幂等：对同一节点重复调用安全（已从树上摘除的节点直接跳过），
     * 便于剪切粘贴等场景无需额外判断节点是否仍在树上
     * @private
     * @param {TreeNodeView[]} views - 待删除节点集合
     */
    private removeNodes(views: TreeNodeView[]): void {
        let removed = false;
        views.forEach((view) => {
            // 根节点（parent 为 null）受保护，不可删除：删除后树将失去唯一入口
            if (this.isRootView(view)) return;

            // 已从父节点 children 摘除说明此前已删除，跳过避免误删同位置的新节点
            const siblings = view.parent?.children;
            const index = siblings ? siblings.indexOf(view) : -1;
            if (siblings && index < 0) return;

            if (siblings) siblings.splice(index, 1);
            if (this.selected === view) this.selected = null;
            if (this.renaming?.view === view) this.renaming = null;
            view.el.remove();
            removed = true;
        });
        if (removed) this.updateEmptyState();
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
        this.rootViews.forEach((view) => this.walk(view, (node) => {
            if (node.isFolder) this.setOpen(node, true);
        }));
    }

    /**
     * 收起全部文件夹
     */
    public collapseAll(): void {
        this.rootViews.forEach((view) => this.walk(view, (node) => {
            if (node.isFolder) this.setOpen(node, false);
        }));
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

        for (let parent = target.parent; parent !== null; parent = parent.parent) {
            if (parent.isFolder) this.setOpen(parent, true);
        }

        this.select(target);
        target.row.scrollIntoView({ block: 'nearest' });
        this.flash(target);
    }

    /**
     * 设置当前主页面显示的文件名（「定位」据此在树中查找）
     * @param {string} name - 文件名（不含扩展名，如 data.activedSheetName）
     */
    public setActivedFileName(name: string): void {
        this.activedFileName = name ?? '';
    }

    /**
     * 按文件名在整棵树中查找节点（不区分大小写，优先匹配文件而非文件夹）
     * @private
     * @param {string} name - 文件名
     * @returns {TreeNodeView | null} 命中的节点
     */
    private findByName(name: string): TreeNodeView | null {
        const key = name.trim().toLowerCase();
        if (!key) return null;
        const fallback: TreeNodeView[] = [];
        const walk = (views: TreeNodeView[]): TreeNodeView | null => {
            for (const view of views) {
                if (view.key === key) {
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

        // 展开祖先并选中，保证新增节点可见
        for (let p = view.parent; p !== null; p = p.parent) {
            if (p.isFolder) this.setOpen(p, true);
        }
        this.select(view);
        if (select) this.flash(view);
        this.updateEmptyState();
        return view;
    }

    /**
     * 在指定父节点下生成不重复的名称
     *
     * 名称冲突时在扩展名前追加递增序号（`报表` → `报表1` → `报表2` …），
     * 与 {@link availableName} 的序号风格保持一致；粘贴时用于保留原名的兜底去重。
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
        for (let seq = 1; seq < 10000; seq++) {
            const candidate = `${base}${seq}${suffix}`;
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
