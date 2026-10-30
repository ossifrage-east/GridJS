/*  # 通用文件树组件 TS 文件（「文件」导航页）
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { createDiv, createInput } from "../../../utils/dom";
import { FILE_TREE_CLASS_NAMES as TREE } from "../../constant";
import type { FileNode, TreeNodeView } from "./fileTreeModel";
import {
    DEFAULT_FILE_NODES,
    availableFileName as modelAvailableFileName,
    collectFilePaths as modelCollectFilePaths,
    findByName as modelFindByName,
    getDepth as modelGetDepth,
    pathWithSelf as modelPathWithSelf,
    serializeNodes as modelSerializeNodes,
    uniqueName as modelUniqueName,
} from "./fileTreeModel";
import { FileTreeUi } from "./fileTreeUi";
import type { FileTreeOptions } from "./fileTreeUi";
import '../../style/fileTree.scss';

export type { FileNode, TreeNodeView } from "./fileTreeModel";
export type { FileContentBridge, FileMovePlan } from "./fileContentBridge";
export type { FileTreeOptions } from "./fileTreeUi";


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
 *   文件基础名自带序号（`sheet1.grid` / `doc1.text` / `ppt1.demo`），
 *   冲突时递增其中的数字（`sheet1.grid` → `sheet2.grid` …）。
 *   粘贴 / 复制重名时用同一套规则（{@link uniqueName}）：名称里带数字就递增它
 *   （`doc1.text` → `doc2.text`），否则追加序号（`报表.grid` → `报表1.grid`）。
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
 *   **多选（Shift 连选）后拖动选区中的任意一行，整组选区一起移动**，落点目录内
 *   的先后次序与拖拽前在树中的顺序一致；不能拖到自身或自身的子孙里（防环），
 *   拖到原父目录视为无操作。
 * - 属性对话框：右键菜单「属性」打开一个对话框，展示名称与后缀、位置、创建与
 *   修改时间及人员、是否可编辑。
 * - 另存为：右键菜单「打开」下方为「另存为」（hover 展开二级菜单：另存为xlsx /
 *   另存为docx / 另存为pptx / 另存为pdf，后三项暂未实现、置灰不可点）。
 *   「另存为xlsx」弹出一个表格对话框，列出当前选中的 `.grid` 文件：
 *   多选框 + 序号（表头多选框全选）/ 文件名 / 文件大小 / 进度 /
 *   操作（开始、暂停、删除，仅图标）/ 下载（表头为「打包下载」按钮）。
 *   转换由 {@link FileTreeOptions.convertXlsx} 建出的作业真实生成 xlsx，按单飞队列
 *   （一次只跑一个，其余显示「排队中」）与时间片驱动；暂停会真正停下并保留断点，
 *   再次「开始」从断点续跑。未转换完成的行其「下载」与表头「打包下载」均不可用。
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
export class FileTree extends FileTreeUi {

    /**
     * 构造函数
     * @param {FileTreeOptions} [options={}] - 组件配置选项
     */
    constructor(options: FileTreeOptions = {}) {
        super();
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

    /** 瑙嗗浘搴忓垪鍖栦负绾暟鎹紙鎸佷箙鍖栧鍑虹敤锛?*/
    private serializeNodes(views: TreeNodeView[], withMeta = false): FileNode[] {
        return modelSerializeNodes(views, withMeta);
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
     * 因此默认得到 `sheet1.grid`（同名被占用时递增序号）。
     * 供外部在「刷新后树里一个文件都没有」时补一个默认空白文件用。
     * @param {string} [todo='XLSX'] - 文件类型标识（XLSX / DOCX / PPTX）
     * @returns {string} 可用的文件显示名（含扩展名）
     */
    public nextFileName(todo: string = 'XLSX'): string {
        return this.availableFileName(todo, null);
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
     * 为「从本机打开的文件」解析一个在根目录内不冲突的文件名
     *
     * 打开时**不同原始后缀可能换算成同一原生名**：`test_sample.xls` 与
     * `test_sample.xlsx` 都换算为 `test_sample.grid`（同理 `doc/docx → .text`、
     * `ppt/pptx → .demo`）。若不去重，第二个文件会被 {@link showActivedFile} 当成
     * 「树里已有同名节点」而复用，同时写入同一个存档键 —— 表现为「新文件不进树、
     * 内容覆盖掉先打开的那个」。
     *
     * 这里按与新建 / 粘贴一致的规则（{@link uniqueName}）在冲突时加序号：
     * 名称内无数字则追加序号（`test_sample.grid` → `test_sample1.grid`），
     * 名称内已有数字则递增该数字（`报告2.grid` → `报告3.grid`）。
     *
     * ⚠️ 调用方必须把返回名**同时**用于装载（存档键）与登记（树节点），二者一致
     * 才能保证「一节点一存档」。
     * @param {string} name - 换算后的原生文件名（如 `test_sample.grid`）
     * @returns {string} 根目录内不冲突的文件名
     */
    public resolveOpenedFileName(name: string): string {
        const target = this.rootViews[0];
        if (!target?.childrenBox) return name;
        return this.uniqueName(target, name);
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
        this.highlighted.clear();
        this.selectionAnchor = null;
        this.selected = null;
        this.rootViews = [];
        this.container.remove();
    }
}