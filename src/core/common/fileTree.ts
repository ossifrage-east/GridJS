/*  # 通用文件树组件 TS 文件（「文件」导航页）
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { createDiv, createSpan, createInput, createButton, createElement } from "../../utils/dom";
import { FILE_TREE_CLASS_NAMES as TREE } from "../constant";
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
 * 文件树构造选项接口
 * @interface FileTreeOptions
 */
export interface FileTreeOptions {
    /** 父元素（缺省时仅创建 DOM，不挂载） */
    parentElement?: HTMLElement | null;
    /** 树数据，缺省时使用内置示例数据 {@link DEFAULT_FILE_NODES} */
    nodes?: FileNode[];
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
    /** 子节点视图集合 */
    children: TreeNodeView[];
    /** 展开状态：用户操作记忆，搜索期间不被覆盖，清空搜索后据此恢复 */
    isOpen: boolean;
}

/** 内置示例文件树：真实数据可通过 {@link FileTreeOptions.nodes} 注入 */
const DEFAULT_FILE_NODES: FileNode[] = [
    {
        name: '工作区',
        children: [
            {
                name: '表格',
                children: [
                    { name: '销售数据.xlsx' },
                    { name: '库存清单.xlsx' },
                    { name: '财务预算.xlsx' },
                ],
            },
            {
                name: '报表',
                children: [
                    { name: '月度汇总.xlsx' },
                    { name: '季度分析.xlsx' },
                ],
            },
            { name: '说明.md' },
        ],
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
 * - 定位：展开目标节点的全部祖先并滚动到可视区域，同时闪烁高亮。
 *   搜索状态下定位首个命中项，否则定位当前选中项，无选中项时定位第一个根节点。
 *
 * @class FileTree
 * @example
 * const fileTree = new FileTree();
 * navigator.navContentItemMount('文件', fileTree.getElements().container);
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
    /** 根节点视图集合 */
    private rootViews: TreeNodeView[] = [];
    /** 当前选中节点 */
    private selected: TreeNodeView | null = null;
    /** 当前搜索关键字（小写，空串表示未搜索） */
    private keyword = '';
    /** 定位闪烁定时器 */
    private flashTimer: number | null = null;

    /**
     * 构造函数
     * @param {FileTreeOptions} [options={}] - 组件配置选项
     */
    constructor(options: FileTreeOptions = {}) {
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
        row.appendChild(createSpan({ className: TREE.LABEL, textContent: data.name }));
        el.appendChild(row);

        const view: TreeNodeView = {
            key: data.name.toLowerCase(),
            isFolder,
            parent,
            el,
            row,
            childrenBox: null,
            icon,
            children: [],
            isOpen: true,
        };

        if (isFolder) {
            const childrenBox = createDiv({ className: TREE.CHILDREN });
            view.childrenBox = childrenBox;
            view.children = this.createNodes(data.children as FileNode[], childrenBox, depth + 1, view);
            el.appendChild(childrenBox);
            // 默认展开第一层（根节点），更深的层级收起
            this.setOpen(view, depth < 1);
        }

        row.addEventListener('click', () => this.handleRowClick(view));
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
     * 展开目标节点的全部祖先，滚动到可视区域并闪烁高亮。
     * 搜索状态下定位首个命中项；否则定位当前选中项，无选中项时定位第一个根节点。
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
     * 获取定位目标节点
     * @private
     * @returns {TreeNodeView | null} 目标节点
     */
    private getLocateTarget(): TreeNodeView | null {
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
