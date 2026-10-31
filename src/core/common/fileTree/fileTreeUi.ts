/*  # 通用文件树组件 TS 文件（「文件」导航页）
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

/*  # 文件树交互实现基类：字段与私有方法（由 fileTree.ts 的 FileTree 继承，保持原有行为） */

import { createDiv, createSpan, createInput, createButton, createElement } from "../../../utils/dom";
import { extractExtension, toArchiveKey } from "../../../utils/fileName";
import type { XlsxTask } from "../../../utils/parser/xlsxWriter";
import { FILE_TREE_CLASS_NAMES as TREE } from "../../constant";
import { Menu, MenuContent } from "../menu";
import { Dialog } from "../dialog";
import type { FileContentBridge, FileMovePlan } from "./fileContentBridge";
import type { FileNode, TreeNodeView } from "./fileTreeModel";
import {
    DEFAULT_FILE_NODES,
    LOCAL_USER,
    availableFileName as modelAvailableFileName,
    availableName as modelAvailableName,
    collectFilePaths as modelCollectFilePaths,
    collectFilePlans as modelCollectFilePlans,
    containerPathOf as modelContainerPathOf,
    findByName as modelFindByName,
    findFirstMatch as modelFindFirstMatch,
    getDepth as modelGetDepth,
    isDescendantView as modelIsDescendantView,
    pathOf as modelPathOf,
    pathWithSelf as modelPathWithSelf,
    serializeNodes as modelSerializeNodes,
    uniqueName as modelUniqueName,
    walk as modelWalk,
} from "./fileTreeModel";
import { createNewFileSubmenuItem, createSaveAsSubmenuItem } from "./saveAsDialog";
import { showPropertyDialog } from "./propertyDialog";
import { FileTransferPanel, type TransferKind } from "./fileTransferPanel";

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
    /**
     * 读取文件内容的存档大小（字节），用于「另存为」对话框的「文件大小」列（可选）
     *
     * 文件内容存在外部（IndexedDB），树只认识路径，因此由外部按路径读取；
     * 不传或读取失败时该列显示「—」。
     */
    fileSize?: (path: string) => Promise<number | null>;
    /**
     * 为 `.grid` 文件创建 xlsx 转换作业（「另存为xlsx」对话框的真实转换，可选）
     *
     * 存档在 IndexedDB 中，树只认识路径，因此由外部读取内容并创建作业；
     * 作业由对话框按时间片推进（单飞队列：一次只跑一个），暂停即停止推进、
     * 继续则从断点接着跑。不传时转换会报错并回到「待开始」。
     * @param path - 文件路径（含扩展名）
     * @param name - 文件名（用于工作表名）
     */
    convertXlsx?: (path: string, name: string) => Promise<XlsxTask>;
    /**
     * 为 `.text` 文档创建 docx 转换作业（「另存为docx」对话框的真实转换，可选）
     *
     * 文档的块树存在 IndexedDB 的 `docs` 仓里，而「怎么把块树编成 docx」
     * 属于文档域的知识 —— 树只认识路径，因此由外部按路径读取并编码。
     * 作业由对话框按时间片推进（与 xlsx 同一套单飞队列）；不传时转换会报错并回到「待开始」。
     * @param path - 文件路径（含扩展名）
     * @param name - 文件名
     */
    convertDocx?: (path: string, name: string) => Promise<XlsxTask>;
}

export class FileTreeUi {
    /** 页面容器（挂载到导航内容项） */
    protected container: HTMLDivElement;
    /** 顶部工具栏（搜索栏 + 操作按钮） */
    protected toolbar: HTMLDivElement;
    /** 搜索输入框 */
    protected searchInput: HTMLInputElement;
    /** 树主体（滚动容器） */
    protected body: HTMLDivElement;
    /** 无匹配结果提示 */
    protected empty: HTMLDivElement;
    /** 配置选项 */
    protected options: FileTreeOptions;
    /** 根节点视图集合 */
    protected rootViews: TreeNodeView[] = [];
    /** 当前选中节点（Shift 连选时为主操作对象，即最后点击的那个） */
    protected selected: TreeNodeView | null = null;
    /**
     * Shift 连选的锚点（上一次**单击**选中的节点）
     *
     * 连选期间保持不变，因此连续 Shift 点击可反复调整选区另一端（与桌面文件管理器一致）。
     */
    protected selectionAnchor: TreeNodeView | null = null;
    /** 当前所有高亮行：单选 1 个，Shift 连选为区间内的全部可见行 */
    protected highlighted = new Set<TreeNodeView>();
    /** 当前搜索关键字（小写，空串表示未搜索） */
    protected keyword = '';
    /** 定位闪烁定时器 */
    protected flashTimer: number | null = null;
    /** 当前主页面显示的文件名（用于「定位」在树中找到对应节点） */
    protected activedFileName = '';
    /** 剪贴板暂存：剪切 / 复制的节点数据（粘贴时深拷贝使用） */
    protected clipboard: FileNode[] | null = null;
    /** 剪贴板模式：true=剪切（粘贴后删除源），false=复制 */
    protected clipboardIsCut = false;
    /** 剪切来源节点：粘贴完成后据此删除原节点（复制模式为 null） */
    protected clipboardSources: TreeNodeView[] = [];
    /** 正在重命名的节点视图（同一时刻仅允许一个） */
    protected renaming: { view: TreeNodeView; input: HTMLInputElement } | null = null;
    /** 正在被拖拽的节点（拖拽期间非 null，用于判定「当前是否处于拖拽中」） */
    protected dragging: TreeNodeView | null = null;
    /**
     * 本次拖拽要移动的节点集合（多选时即整组选区）
     *
     * 只保存选区里**最靠上**的节点：文件夹与其子孙同时被选中时，子孙会随文件夹
     * 整体搬走，若再单独移动一次就会被从文件夹里摘出来。顺序按树中的可见顺序。
     */
    protected dragSources: TreeNodeView[] = [];
    /** 「属性」对话框实例（复用同一实例，避免反复创建 DOM） */
    protected propertyDialog: Dialog | null = null;
    /** 「另存为」批量任务面板（xlsx / docx 共用；懒创建，内部自带对话框与单飞队列） */
    protected transferPanel: FileTransferPanel | null = null;

    /**
     * 在指定父节点下新建一个文件夹
     * @private
     * @param {TreeNodeView | null} parent - 父节点（缺省为根节点「工作区」）
     */
    protected createFolderNode(parent: TreeNodeView | null = null): void {
        this.addNode(this.availableName('新建文件夹', parent), { isFolder: true, select: true, parent });
    }

    /**
     * 在指定父节点下新建一个空白文件节点
     * @private
     * @param {string} todo - 菜单动作标识（XLSX / DOCX / PPTX）
     * @param {TreeNodeView | null} parent - 父节点（缺省为根节点「工作区」）
     */
    protected createFileNode(todo: string, parent: TreeNodeView | null = null): void {
        this.addNode(this.availableFileName(todo, parent), { select: true, parent });
    }



    /**
     * 创建顶部工具栏
     * 同一行内：左侧搜索栏，右侧展开 / 收缩 / 定位按钮（按钮只显示图标）
     * @private
     * @returns {HTMLDivElement} 工具栏元素
     */
    protected createToolbar(): HTMLDivElement {
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
    protected createIconButton(title: string, icon: string, onClick: () => void): HTMLButtonElement {
        const button = createButton({
            className: TREE.BUTTON,
            attributes: { type: 'button', title },
        });
        button.appendChild(createElement('i', { className: icon }));
        button.addEventListener('click', () => onClick());
        return button;
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
    protected createNodes(
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
    protected createNode(
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

        row.addEventListener('click', (e: MouseEvent) => this.handleRowClick(e, view));
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
     * 由 {@link moveNodesTo} 统一完成校验、搬迁、内容同步与落库。
     * 若拖起的行在当前选区内（Shift 连选出的多行），则整组选区一起移动。
     *
     * 拖拽越界保护在**拖起时就判定**（`dragstart` 里检测自身/子孙行）
     * 而不是等到 `dragover`：被拖拽的元素在拖拽期间通常不派发 dragover/drop，
     * 只有靠 `dragstart` 阶段标记「禁止」才能可靠拦下「拖到自己子树里」。
     * @private
     * @param {TreeNodeView} view - 节点
     */
    protected bindDragEvents(view: TreeNodeView): void {
        view.row.draggable = true;

        view.row.addEventListener('dragstart', (e: DragEvent) => {
            // 重命名输入框内的拖拽是文本选择，不应触发节点搬移
            if (this.renaming?.view === view) {
                e.preventDefault();
                return;
            }
            // 拖起的行不在选区里时先选中它：mousedown 后直接拖动不会派发 click，
            // 选区不会自动更新（与桌面文件管理器一致）
            if (!this.highlighted.has(view)) this.select(view);
            this.dragging = view;
            this.dragSources = this.collectDragSources(view);
            // 拖拽期间禁用「被拖节点自身及其全部子孙」作为放置目标，避免拖成环
            this.rootViews.forEach((root) => this.walk(root, (node) => {
                if (this.isInDragSources(node)) {
                    node.row.draggable = false;
                    node.row.classList.add(TREE.DRAG_DISABLED);
                }
            }));
            // 整组降透明度：拖拽期间能看出「一次搬走的是这几个」
            this.dragSources.forEach((source) => source.row.classList.add(TREE.DRAGGING));
            if (e.dataTransfer) {
                e.dataTransfer.effectAllowed = 'move';
                // 部分浏览器（Firefox）要求必须写入数据才会开始拖拽
                const count = this.dragSources.length;
                e.dataTransfer.setData('text/plain', count > 1
                    ? `${count} 个项目`
                    : view.label.textContent ?? '');
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
            if (!this.dragging) return;
            // 不能先 clearDragState：它会清空 `this.dragSources`，
            // 而 moveNodesTo 正需要这份整组来源清单
            this.moveNodesTo(this.dragSources, view);
            this.clearDragState();
        });
    }

    /**
     * 判断当前拖拽的节点能否放置到目标文件夹下（O(1) 校验）
     *
     * 整组移动的判定：
     * - 目标不能是任一被拖节点自身或其子孙（会形成环）；
     * - 至少要有一个节点不在目标目录下（全部已在目标目录 = 拖了没有变化）。
     * @private
     * @param {TreeNodeView} target - 目标文件夹
     * @returns {boolean} 是否允许放置
     */
    protected canDropOn(target: TreeNodeView): boolean {
        if (!target.isFolder || this.dragSources.length === 0) return false;
        if (this.dragSources.some((source) =>
            source === target || this.isDescendantView(source, target))) {
            return false;
        }
        return this.dragSources.some((source) => source.parent !== target);
    }

    /**
     * 判断节点是否属于「本次拖拽的节点或其子孙」
     * @private
     * @param {TreeNodeView} node - 待判定节点
     * @returns {boolean} 是否在拖拽子树内
     */
    protected isInDragSources(node: TreeNodeView): boolean {
        return this.dragSources.some((source) =>
            source === node || this.isDescendantView(source, node));
    }

    /**
     * 清理拖拽过程中的临时状态（游标、禁止标记、放置高亮、拖拽视觉）
     * @private
     */
    protected clearDragState(): void {
        this.dragging = null;
        this.dragSources = [];
        this.rootViews.forEach((root) => this.walk(root, (node) => {
            node.row.draggable = true;
            node.row.classList.remove(TREE.DRAG_DISABLED);
            node.row.classList.remove(TREE.DROP_TARGET);
            node.row.classList.remove(TREE.DRAGGING);
        }));
    }

    /**
     * 把一组节点（含各自子树）移动到目标文件夹下 —— 即改变它们的路径
     *
     * 多个节点按传入顺序（树中的可见顺序）依次挂载，因此目标目录内的先后次序
     * 与拖拽前一致。每个节点都按「剪切 → 粘贴」的既有顺序处理：
     * 1. 先确定新名称（目标目录内重名去重，规则见 {@link uniqueName}）；
     * 2. 在**移动之前**按「旧路径 → 新路径」收集搬迁计划
     *    （移动之后父链已变，再也算不出旧路径）；
     * 3. 再从 DOM 与父链上摘除、挂到目标文件夹并修正缩进。
     *
     * 已在目标目录下的节点（路径没有变化）原地不动，直接跳过；
     * 文件夹整体移动时，其中每个文件的路径都变了 → 计划里逐个列出。
     * 全部节点的搬迁计划汇总后一次性交给内容桥，最后统一落库一次。
     * @private
     * @param {TreeNodeView[]} sources - 被拖拽的节点（同一次拖拽的整组）
     * @param {TreeNodeView} target - 目标文件夹
     */
    protected moveNodesTo(sources: TreeNodeView[], target: TreeNodeView): void {
        if (!target.childrenBox || sources.length === 0) return;

        const movable = sources.filter((source) =>
            source !== target
            && source.parent !== target
            && !this.isDescendantView(source, target));
        if (movable.length === 0) return;

        const plans: FileMovePlan[] = [];
        movable.forEach((source) => {
            const oldPaths = this.collectFilePaths(source);
            const oldName = source.label.textContent ?? '';
            const name = this.uniqueName(target, oldName);

            // 先提交重命名（名称可能因重名去重而变），再算新路径
            if (name !== oldName) {
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
            for (let i = 0; i < oldPaths.length && i < newPaths.length; i++) {
                if (oldPaths[i] !== newPaths[i]) plans.push({ from: oldPaths[i], to: newPaths[i] });
            }

            // 被移动的是当前文件、或其祖先时，高亮与「当前文件名」需要跟着走
            this.retargetActived(source, oldPaths, newPaths);
            this.touchNode(source);
        });

        if (plans.length > 0) this.options.content?.move?.(plans);

        this.setOpen(target, true);
        // 移动后整组保持选中（与桌面文件管理器一致），可继续拖动或删除
        this.selectMany(movable);
        this.notifyTreeChanged();
    }

    /**
     * 计算本次拖拽要移动的节点集合（多选时整组一起移动）
     *
     * - 拖起的行未被选中时，`dragstart` 已把选区重置为它，这里只返回它自己；
     * - 拖起的行在选区里时，返回选区内的**全部可见行**（搜索过滤 / 折叠的内部行不参与）；
     * - 选区里文件夹与其子孙同时被选中时只保留文件夹：子孙会随文件夹整体搬走，
     *   再单独移动一次就会被从文件夹里摘出来；
     * - 根节点「工作区」不可移动，一律排除。
     *
     * 顺序按树中的可见顺序（{@link visibleRows}），移动后目标目录内的次序即此顺序。
     * @private
     * @param {TreeNodeView} view - 被拖起的节点
     * @returns {TreeNodeView[]} 要移动的节点集合
     */
    protected collectDragSources(view: TreeNodeView): TreeNodeView[] {
        const candidates = this.visibleRows()
            .filter((node) => this.highlighted.has(node) && !this.isRootView(node));
        const topLevel = candidates.filter((node) => !candidates.some(
            (other) => other !== node && other.isFolder && this.isDescendantView(other, node)));
        return topLevel.length > 0 ? topLevel : [view];
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
    protected retargetActived(moved: TreeNodeView, oldPaths: string[], newPaths: string[]): void {
        if (!this.activedFileName) return;
        const current = this.findByName(this.activedFileName);
        if (!current) return;
        // 当前文件不在被移动的子树里：路径未变
        if (current !== moved && !this.isDescendantView(moved, current)) return;

        for (let i = 0; i < oldPaths.length; i++) {
            // 用**存档键**比对（而不是去扩展名）：`报表.grid` 与 `报表.text` 去扩展名后都是
            // `报表`，会把另一个同名文件误判成当前文件。activedFileName 可能是键也可能是路径，
            // 而 toArchiveKey 是幂等的「路径 → 键」映射，两边过一遍即可精确对齐。
            if (toArchiveKey(oldPaths[i]) === toArchiveKey(this.activedFileName)) {
                this.activedFileName = newPaths[i] ?? this.activedFileName;
                return;
            }
        }
    }

    /**
     * 行点击：选中节点；文件夹同时切换展开 / 收起
     *
     * Shift + 左键为**连选**：从上次单击的节点（锚点）到本次点击的节点，
     * 区间内所有**可见**行一并高亮（用于表示多个文件被选中），
     * 且不改变展开状态 —— Shift 点文件夹只是把它并入选区，不应顺带收起。
     * @private
     * @param {MouseEvent} e - 点击事件（取 shiftKey）
     * @param {TreeNodeView} view - 被点击的节点
     */
    protected handleRowClick(e: MouseEvent, view: TreeNodeView): void {
        if (e.shiftKey && this.selectionAnchor && this.highlighted.size > 0) {
            this.selectRange(this.selectionAnchor, view);
            return;
        }
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
     * 菜单结构（自上而下）：新建文件（hover 展开二级菜单）/ 新建文件夹 / 打开 /
     * 另存为（hover 展开二级菜单） / 分割线 / 剪切 / 复制 / 粘贴 / **删除** /
     * 分割线 / 重命名 / **属性**。
     *
     * 「删除」放在「粘贴」之下、分割线之上（危险操作紧邻同组的文件操作，
     * 与「重命名 / 属性」这类节点自身操作分隔开）。
     * 右键文件节点时「新建」类操作作用于其所在文件夹；右键文件夹时作用于该文件夹。
     * @private
     * @param {MouseEvent} e - 右键事件
     * @param {TreeNodeView} view - 被右键的节点
     */
    protected openContextMenu(e: MouseEvent, view: TreeNodeView): void {
        const { menu } = this.options;
        if (!menu) return;

        // 右键即选中；但右键**已在选区中**的行时保留多选，
        // 使「另存为」等整组操作能作用于当前选中的全部文件
        if (!this.highlighted.has(view)) this.select(view);
        // 新建目标：文件夹自身；文件则用其父节点（根节点的父为 null，回落到根节点）
        const parent = view.isFolder ? view : view.parent;

        const content = new MenuContent({
            items: [
                // 新建文件：hover 展开多项二级菜单
                () => this.createNewFileSubmenuItem(view),
                { todo: 'newFolder', icon: 'icon-new-folder', text: '新建文件夹' },
                { todo: 'open', icon: 'icon-open', text: '打开' },
                // 另存为：hover 展开「另存为xlsx / docx / pptx / pdf」
                () => this.createSaveAsSubmenuItem(),
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



    // ════════════════════════════════════════════════════════════
    //  另存为（导出清单对话框）
    // ════════════════════════════════════════════════════════════

    /**
     * 收集选中的文件（按树中可见顺序，序号即返回顺序）
     *
     * 只认「已选中且不是文件夹」的行，再按原生后缀过滤；
     * xlsx 取 `.grid`，docx 取 `.text`。
     * @private
     * @param extension - 原生后缀（`grid` / `text`）
     * @returns {Array<{ path: string; name: string }>} 选中文件
     */
    protected collectSelectedFiles(extension: string): Array<{ path: string; name: string }> {
        return this.visibleRows()
            .filter((view) => this.highlighted.has(view) && !view.isFolder)
            .map((view) => ({ path: this.pathWithSelf(view), name: view.label.textContent ?? '' }))
            .filter((file) => extractExtension(file.name) === extension);
    }

    /**
     * 显示「另存为」对话框（xlsx / docx 同一套清单）
     *
     * 清单渲染、单飞队列与时间片驱动都在 {@link FileTransferPanel} 内，
     * 树只负责「收集选中的源文件」并把转换 / 读大小能力透传过去。
     * @private
     * @param extension - 源文件原生后缀（`grid` / `text`）
     * @param kind - 目标格式（`xlsx` / `docx`）
     */
    protected showSaveAsDialog(extension: string, kind: TransferKind): void {
        const files = this.collectSelectedFiles(extension);
        if (!this.transferPanel) {
            this.transferPanel = new FileTransferPanel({
                convertXlsx: this.options.convertXlsx,
                convertDocx: this.options.convertDocx,
                fileSize: this.options.fileSize,
            });
        }
        this.transferPanel.open(files, kind);
    }

    /**
     * 显示「另存为xlsx」对话框（列出选中的 `.grid` 文件）
     * @private
     */
    protected showSaveAsXlsxDialog(): void {
        this.showSaveAsDialog('grid', 'xlsx');
    }

    /**
     * 显示「另存为docx」对话框（列出选中的 `.text` 文件）
     * @private
     */
    protected showSaveAsDocxDialog(): void {
        this.showSaveAsDialog('text', 'docx');
    }

    // ════════════════════════════════════════════════════════════
    //  属性对话框
    // ════════════════════════════════════════════════════════════


    /**
     * 显示节点「属性」对话框
     *
     * 正文构建与对话框复用都在 {@link showPropertyDialog} 内；
     * 树只负责取路径，并在打开后回调外部。
     * @private
     * @param {TreeNodeView} view - 目标节点
     */
    protected showProperty(view: TreeNodeView): void {
        const path = this.pathWithSelf(view);
        this.propertyDialog = showPropertyDialog(this.propertyDialog, view, path);
        this.options.onPropertyClose?.(path);
    }

    /**
     * 处理右键菜单动作
     * @private
     * @param {string} todo - 动作标识
     * @param {TreeNodeView} view - 操作目标节点
     * @param {TreeNodeView | null} parent - 新建 / 粘贴的目标父节点
     */
    protected handleContextAction(
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
                this.clipboard = modelSerializeNodes([view]);
                this.clipboardIsCut = true;
                this.clipboardSources = [view];
                break;
            case 'copy':
                this.clipboard = modelSerializeNodes([view]);
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
     * 导出当前整棵树（深拷贝，可直接交给结构化克隆持久化）
     *
     * 与剪贴板用的 {@link serializeNodes} 同源，额外带上：
     * - **路径**（`path`，内容存档键的来源，见 {@link FileNode.path}）；
     * - 创建 / 最后修改的时间与人（供「属性」对话框显示）；
     * - 文件夹展开状态（`open`），使刷新还原后「原来折叠的文件夹依然是折叠的」。
     * @returns {FileNode[]} 树数据
     */
    /**
     * 导出当前整棵树（深拷贝，可直接交给结构化克隆持久化）
     *
     * 与剪贴板用的 {@link serializeNodes} 同源，额外带上：
     * - **路径**（`path`，内容存档键的来源，见 {@link FileNode.path}）；
     * - 创建 / 最后修改的时间与人（供「属性」对话框显示）；
     * - 文件夹展开状态（`open`），使刷新还原后「原来折叠的文件夹依然是折叠的」。
     * @returns {FileNode[]} 树数据
     */
    protected getNodes(): FileNode[] {
        return modelSerializeNodes(this.rootViews, true);
    }
    /**
     * 通知外部树已变化（供持久化）
     *
     * 在**结构、名称或文件夹展开状态**发生变化时调用：新增 / 删除 / 重命名 /
     * 粘贴，以及展开收起（单点、全展开 / 全收起、定位时的祖先展开）。
     * 纯选中不触发；搜索只做临时渲染（{@link renderOpen}）不改状态，同样不触发。
     * @private
     */
    protected notifyTreeChanged(): void {
        this.options.onChange?.(this.getNodes());
    }

    /**
     * 粘贴剪贴板中的节点到指定父节点下
     * @private
     * @param {TreeNodeView | null} parent - 目标父节点
     */
    protected pasteNodes(parent: TreeNodeView | null): void {
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
        // 根节点「工作区」是容器而非目录，其路径为 `''`（否则会得到「工作区/x.grid」）
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
     * 更新节点的「最后修改」元数据（名称/位置变化时调用）
     * @private
     * @param {TreeNodeView} view - 节点
     */
    protected touchNode(view: TreeNodeView): void {
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
    protected ensureMetadata(view: TreeNodeView): void {
        const now = Date.now();
        if (!Number.isFinite(view.createdAt)) view.createdAt = now;
        if (!Number.isFinite(view.lastModifiedAt)) view.lastModifiedAt = view.createdAt;
        if (!view.createdBy) view.createdBy = LOCAL_USER;
        if (!view.lastModifiedBy) view.lastModifiedBy = view.createdBy;
        view.children.forEach((child) => this.ensureMetadata(child));
    }


    /**
     * 依据数据递归创建节点（名称已由调用方确定）
     * @private
     * @param {FileNode} data - 节点数据
     * @param {TreeNodeView} parent - 父节点
     * @returns {TreeNodeView} 新建的节点视图
     */
    protected createFromData(data: FileNode, parent: TreeNodeView): TreeNodeView {
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
    protected startRename(view: TreeNodeView): void {
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
    protected commitRename(apply: boolean): void {
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
                // 同上：按存档键精确比对，避免 `报表.grid` / `报表.text` 互相误判
                if (toArchiveKey(oldPaths[i]) === toArchiveKey(this.activedFileName)) {
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
    protected removeNodes(views: TreeNodeView[], purgeContent = true): void {
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
            // 连选锚点 / 高亮集合也要同步摘除：节点已销毁，留着会在下次 Shift 点击时
            // 用 indexOf 找不到锚点（正确退化），高亮集合里则是永久泄漏的死引用
            if (this.selectionAnchor === view) this.selectionAnchor = null;
            this.highlighted.delete(view);
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
    protected isRootView(view: TreeNodeView): boolean {
        return this.rootViews.includes(view);
    }

    /**
     * 选中节点（单选：清掉此前的范围选区）
     * @private
     * @param {TreeNodeView | null} view - 目标节点
     */
    protected select(view: TreeNodeView | null): void {
        if (this.selected === view && this.highlighted.size === (view ? 1 : 0)) return;
        this.clearHighlight();
        this.selected = view;
        this.selectionAnchor = view;
        if (view) {
            view.row.classList.add(TREE.SELECTED);
            this.highlighted.add(view);
        }
    }

    /**
     * Shift 连选：把锚点到目标节点之间**全部可见行**标记为选中
     *
     * 只取可见行（搜索过滤掉的行、折叠文件夹内部的行都不参与），
     * 因为「背景色变化」表达的正是「用户当前看得见的这些文件被选中」。
     * 锚点保持不变，主操作对象（{@link selected}）更新为本次点击的节点，
     * 与桌面文件管理器的连选语义一致。
     * @private
     * @param {TreeNodeView} from - 锚点节点（上次单击处）
     * @param {TreeNodeView} to - 本次点击的节点
     */
    protected selectRange(from: TreeNodeView, to: TreeNodeView): void {
        const rows = this.visibleRows();
        const startIndex = rows.indexOf(from);
        const endIndex = rows.indexOf(to);
        // 锚点已不可见（被过滤 / 折叠 / 删除）时退化为单选，避免选中一片看不见的节点
        if (startIndex < 0 || endIndex < 0) {
            this.select(to);
            return;
        }
        this.clearHighlight();
        const first = Math.min(startIndex, endIndex);
        const last = Math.max(startIndex, endIndex);
        for (let i = first; i <= last; i++) {
            rows[i].row.classList.add(TREE.SELECTED);
            this.highlighted.add(rows[i]);
        }
        // 锚点保持不动；主操作对象改为被点节点
        this.selected = to;
    }

    /**
     * 把一组节点设为选中（多选：整组保持高亮）
     *
     * 主操作对象（{@link selected}）与连选锚点都取该组**最后一个**节点，
     * 之后 Shift 点击可从这里继续扩展选区。
     * @private
     * @param {TreeNodeView[]} views - 要选中的节点（按可见顺序）
     */
    protected selectMany(views: TreeNodeView[]): void {
        this.clearHighlight();
        const last = views.length > 0 ? views[views.length - 1] : null;
        views.forEach((view) => {
            view.row.classList.add(TREE.SELECTED);
            this.highlighted.add(view);
        });
        this.selected = last;
        this.selectionAnchor = last;
    }

    /**
     * 清空所有行的高亮（不动锚点）
     * @private
     */
    protected clearHighlight(): void {
        this.highlighted.forEach((view) => view.row.classList.remove(TREE.SELECTED));
        this.highlighted.clear();
        this.selected = null;
    }

    /**
     * 按先序收集当前**可见**的行
     *
     * 可见性判定用 `offsetParent`：搜索过滤（`hidden` 属性）与折叠文件夹
     * （`childrenBox.hidden`）都会让行脱离布局，无需为两条路径各写一遍判断。
     * @private
     * @returns {TreeNodeView[]} 可见行（文档顺序）
     */
    protected visibleRows(): TreeNodeView[] {
        const rows: TreeNodeView[] = [];
        const walk = (view: TreeNodeView): void => {
            if (view.el.hidden || view.row.offsetParent === null) return;
            rows.push(view);
            view.children.forEach(walk);
        };
        this.rootViews.forEach(walk);
        return rows;
    }

    /**
     * 设置展开状态并渲染（同时更新用户记忆状态）
     * @private
     * @param {TreeNodeView} view - 节点
     * @param {boolean} isOpen - 是否展开
     */
    protected setOpen(view: TreeNodeView, isOpen: boolean): void {
        view.isOpen = isOpen;
        this.renderOpen(view, isOpen);
    }

    /**
     * 仅渲染展开状态（不修改用户记忆状态，供搜索期间使用）
     * @private
     * @param {TreeNodeView} view - 节点
     * @param {boolean} isOpen - 是否展开
     */
    protected renderOpen(view: TreeNodeView, isOpen: boolean): void {
        if (!view.isFolder || !view.childrenBox) return;
        view.childrenBox.hidden = !isOpen;
        view.row.classList.toggle(TREE.OPEN, isOpen);
        view.icon.className = isOpen ? 'icon-folder-open' : 'icon-folder';
    }
    /**
     * 应用关键字过滤
     * @private
     */
    protected applyKeyword(): void {
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
    protected resetFilter(view: TreeNodeView): void {
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
    protected filterNode(view: TreeNodeView): boolean {
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
    protected updateEmptyState(): void {
        const hasVisible = this.rootViews.some((view) => !view.el.hidden);
        this.empty.hidden = hasVisible;
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
    protected refreshActivedHighlight(): void {
        const found = this.activedFileName ? this.findByName(this.activedFileName) : null;
        const target = found && !found.isFolder ? found : null;
        this.rootViews.forEach((view) => this.walk(view, (node) => {
            node.row.classList.toggle(TREE.ACTIVED, node === target);
        }));
    }
    /**
     * 获取定位目标节点
     * @private
     * @returns {TreeNodeView | null} 目标节点
     */
    protected getLocateTarget(): TreeNodeView | null {
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
     * 定位反馈：闪烁高亮
     * @private
     * @param {TreeNodeView} view - 目标节点
     */
    protected flash(view: TreeNodeView): void {
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
    // ── 数据模型 / 菜单构建委托（实现见 fileTree/fileTreeModel.ts 与 fileTree/saveAsDialog.ts） ──

    /** 生成当前目录内可用的文件名称（递增基础名中的数字） */
    protected availableFileName(todo: string, parent: TreeNodeView | null = null): string {
        return modelAvailableFileName(todo, parent, this.rootViews[0] ?? null);
    }

    /** 生成当前目录内可用的名称（文件夹命名规则） */
    protected availableName(base: string, parent: TreeNodeView | null = null): string {
        return modelAvailableName(base, parent, this.rootViews[0] ?? null);
    }


    /** 收集「源文件路径 → 目标文件路径」的内容搬迁计划 */
    protected collectFilePlans(
        source: FileNode,
        sourceDir: string,
        targetDir: string,
        plans: FileMovePlan[],
        targetName?: string
    ): void {
        modelCollectFilePlans(source, sourceDir, targetDir, plans, targetName);
    }

    /** 收集节点（含子孙）中所有文件的完整路径（先序） */
    protected collectFilePaths(view: TreeNodeView, out: string[] = []): string[] {
        return modelCollectFilePaths(view, out);
    }

    /** 节点的所属目录路径（不含自身名称） */
    protected pathOf(view: TreeNodeView | null): string {
        return modelPathOf(view);
    }

    /** 节点的完整路径（含自身名称） */
    protected pathWithSelf(view: TreeNodeView): string {
        return modelPathWithSelf(view);
    }

    /** 节点作为目录看待时的路径（根节点为空串） */
    protected containerPathOf(view: TreeNodeView): string {
        return modelContainerPathOf(view);
    }

    /** 判断候选节点是否为给定视图的子孙 */
    protected isDescendantView(view: TreeNodeView, target: TreeNodeView): boolean {
        return modelIsDescendantView(view, target);
    }

    /** 先序遍历子树 */
    protected walk(view: TreeNodeView, visit: (node: TreeNodeView) => void): void {
        modelWalk(view, visit);
    }

    /** 按路径 / 名称在整棵树中查找节点 */
    protected findByName(name: string): TreeNodeView | null {
        return modelFindByName(this.rootViews, name);
    }

    /** 计算节点深度（根节点为 0） */
    protected getDepth(view: TreeNodeView): number {
        return modelGetDepth(view);
    }

    /** 查找首个命中的可见节点（先序） */
    protected findFirstMatch(views: TreeNodeView[]): TreeNodeView | null {
        return modelFindFirstMatch(views, this.keyword);
    }

    /** 在指定父节点下生成不重复的名称 */
    protected uniqueName(parent: TreeNodeView, name: string): string {
        return modelUniqueName(parent, name);
    }

    /** 构建「新建文件」右键菜单项（含 hover 展开的二级菜单） */
    protected createNewFileSubmenuItem(view: TreeNodeView): HTMLElement {
        return createNewFileSubmenuItem(view, this.options.menu, (todo) => {
            this.createFileNode(todo, view.isFolder ? view : view.parent);
        });
    }

    /** 构建「另存为」右键菜单项（含 hover 展开的二级菜单） */
    protected createSaveAsSubmenuItem(): HTMLElement {
        return createSaveAsSubmenuItem(
            this.options.menu,
            (todo) => {
                if (todo === 'saveAsXlsx') this.showSaveAsXlsxDialog();
                if (todo === 'saveAsDocx') this.showSaveAsDocxDialog();
            }
        );
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
}