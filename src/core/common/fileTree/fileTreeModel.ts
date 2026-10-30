/*  # 文件树数据模型（纯结构 / 路径 / 命名规则，零 DOM 操作）
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { stripExtension, extractExtension } from "../../../utils/fileName";
import type { FileMovePlan } from "./fileContentBridge";

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
     */
    path?: string;
    /** 子节点：存在即视为文件夹（空数组表示空文件夹） */
    children?: FileNode[];
    /** 文件夹展开状态（仅文件夹会带此字段，仅供持久化还原使用） */
    open?: boolean;
    /** 节点创建时间（Unix 毫秒） */
    createdAt?: number;
    /** 最后修改人（本机会话标识，见 {@link LOCAL_USER}） */
    lastModifiedBy?: string;
    /** 最后修改时间（Unix 毫秒） */
    lastModifiedAt?: number;
    /** 创建人（本机会话标识） */
    createdBy?: string;
    /** 是否可编辑（单机应用恒为 true，结构为将来接入权限预留） */
    editable?: boolean;
}

/**
 * 视图节点：数据节点 + DOM 引用 + 交互状态
 * @interface TreeNodeView
 */
export interface TreeNodeView {
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
 * 「新建文件」下拉菜单支持的扩展名
 * 与 FileOpener.SUPPORTED_EXTENSIONS 保持一致（此处为展示用的菜单项文案）
 */
export const NEW_FILE_MENU_ITEMS = [
    { todo: 'XLSX', text: '新建表格' },
    { todo: 'DOCX', text: '新建文档' },
    { todo: 'PPTX', text: '新建演示文稿' },
];

/**
 * 新建文件的初始名称
 *
 * 应用的原生后缀只有三种（见 `utils/fileName` 的 `NATIVE_EXTENSIONS`）：
 * `grid`（表格）/ `text`（文档）/ `demo`（演示文稿）。
 * 基础名自带序号（`sheet1.grid` / `doc1.text` / `ppt1.demo`），
 * 当前目录内重名时**递增其中的数字**（`sheet1.grid` → `sheet2.grid` …）。
 */
export const NEW_FILE_BASE_NAMES: Record<string, string> = {
    XLSX: 'sheet1.grid',
    DOCX: 'doc1.text',
    PPTX: 'ppt1.demo',
};

/**
 * 默认文件树：仅保留顶层「工作区」
 *
 * 真实数据可通过 `FileTreeOptions.nodes` 注入；
 * 新建 / 打开 / 粘贴得到的节点会在运行时追加到工作区之下。
 */
export const DEFAULT_FILE_NODES: FileNode[] = [
    {
        name: '工作区',
        children: [],
    },
];

/**
 * 本机会话标识（「创建人 / 最后修改人」的取值）
 *
 * 纯前端单机应用没有登录态，因此以「本机 + 当日会话」作为一个稳定的可读标识。
 * 将来接入账号体系时只需替换本函数。
 * @returns {string} 本机会话标识，如 `本机用户 · 10-14`
 */
export function localUser(): string {
    const now = new Date();
    const pad = (n: number): string => String(n).padStart(2, '0');
    return `本机用户 · ${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
}

/** {@link localUser} 的求值结果（同一会话内保持一致） */
export const LOCAL_USER = localUser();

/**
 * 把字节数格式化为易读的文件大小（「另存为」对话框的「文件大小」列）
 *
 * 小于 1KB 显示为整数 B；更大的单位保留一位小数（`1.5 KB` / `2.3 MB`）。
 * 非法值（负数 / 非有限数）返回「—」，与「读取失败」的展示保持一致。
 * @param {number} bytes - 字节数
 * @returns {string} 易读的文件大小，如 `860 B` / `12.4 KB`
 */
export function formatFileSize(bytes: number): string {
    if (!Number.isFinite(bytes) || bytes < 0) return '—';
    if (bytes < 1024) return `${Math.round(bytes)} B`;
    const units = ['KB', 'MB', 'GB', 'TB'];
    let value = bytes / 1024;
    let unit = 0;
    while (value >= 1024 && unit < units.length - 1) {
        value /= 1024;
        unit++;
    }
    return `${value.toFixed(1)} ${units[unit]}`;
}

/**
 * 格式化时间戳为可读文本（`YYYY-MM-DD HH:mm:ss`）
 *
 * 手工拼接而不用 `toLocaleString`：后者的输出随运行环境语言/时区变化，
 * 既不利于测试断言，也让同一份数据在不同机器上显示不一致。
 * @param {number} ts - Unix 毫秒时间戳
 * @returns {string} 格式化文本；非法值返回 `—`
 */
export function formatTime(ts: number): string {
    if (!Number.isFinite(ts) || ts <= 0) return '—';
    const d = new Date(ts);
    const pad = (n: number): string => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} `
        + `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

/**
 * 计算节点的所属目录路径（不含自身名称）
 *
 * 根节点「工作区」不计入路径 —— 它就是整个树的容器，
 * 因此根下的文件路径即其文件名（如 `sheet1.grid`）。
 * @param {TreeNodeView | null} view - 节点（可为 null，返回空串）
 * @returns {string} 目录路径；根目录下为空串
 */
export function pathOf(view: TreeNodeView | null): string {
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
 * @param {TreeNodeView} view - 节点
 * @returns {string} 完整路径
 */
export function pathWithSelf(view: TreeNodeView): string {
    const dir = pathOf(view);
    const name = view.label.textContent ?? '';
    return dir ? `${dir}/${name}` : name;
}

/**
 * 把节点当作「目录」看待时的路径（用于「粘贴到该目录下」）
 *
 * 与 {@link pathWithSelf} 的唯一区别是**根节点「工作区」返回空串** ——
 * 它是整棵树的容器，不是真实目录，其名称不该出现在任何路径里。
 * @param {TreeNodeView} view - 目录节点
 * @returns {string} 目录路径；根节点返回空串
 */
export function containerPathOf(view: TreeNodeView): string {
    if (view.parent === null) return '';
    return pathWithSelf(view);
}

/**
 * 收集节点（含子孙）中所有文件的**完整路径**（先序）
 *
 * **这是全仓唯一的「搬迁计划」来源**：显示名不含目录，用它构造
 * `{from, to}` 会让不同目录的同名文件互相混淆。
 * 同一次收集的先后顺序固定（先序、深度优先），因此「移动前后各收集一次」
 * 的两份数组可按下标一一对应，从而构造出准确的内容搬迁计划。
 * @param {TreeNodeView} view - 起始节点
 * @param {string[]} out - 路径收集容器
 */
export function collectFilePaths(view: TreeNodeView, out: string[] = []): string[] {
    if (view.isFolder) {
        view.children.forEach((child) => collectFilePaths(child, out));
        return out;
    }
    const dir = pathOf(view);
    const name = view.label.textContent ?? '';
    out.push(dir ? `${dir}/${name}` : name);
    return out;
}

/**
 * 收集「源文件路径 → 目标文件路径」的内容搬迁计划
 *
 * 文件夹只影响其中文件的**路径**，不影响文件名，因此递归展开为其中的文件；
 * 文件夹自身不产生计划项。
 * @param {FileNode} source - 源节点数据（来自剪贴板的深拷贝）
 * @param {string} sourceDir - 源节点在其原树中的所属目录路径
 * @param {string} targetDir - 该节点粘贴后的所属目录路径
 * @param {FileMovePlan[]} plans - 计划收集容器
 * @param {string} [targetName] - 顶层节点粘贴后的名称（去重改名时与源名不同）
 */
export function collectFilePlans(
    source: FileNode,
    sourceDir: string,
    targetDir: string,
    plans: FileMovePlan[],
    targetName?: string
): void {
    const sourcePath = sourceDir ? `${sourceDir}/${source.name}` : source.name;
    // 顶层节点可能因目标目录同名而被去重改名（`台账.grid` → `台账1.grid`），
    // 此时目标路径的第一段必须用改名后的名字，否则内容桥会去搬一个不存在的新键
    const name = targetName ?? source.name;
    const targetPath = targetDir ? `${targetDir}/${name}` : name;
    if (Array.isArray(source.children)) {
        // 只有顶层节点会改名：其子孙沿用自身名称，父级路径已由 targetPath 带上
        source.children.forEach((child) =>
            collectFilePlans(child, sourcePath, targetPath, plans));
        return;
    }
    plans.push({ from: sourcePath, to: targetPath });
}

/** 文件名拆解：前缀 + 数字 + 扩展名（用于「重名时递增数字」） */
interface NameParts {
    /** 名称前缀（不含递增数字与扩展名） */
    prefix: string;
    /** 起始序号 */
    start: number;
    /** 扩展名（含点，可能为空） */
    suffix: string;
}

/**
 * 把「带数字的文件名」拆成前缀 / 起始序号 / 扩展名
 * @param {string} name - 名称（如 `sheet1.grid`）
 * @returns {NameParts} 拆解结果；无数字时 start=1
 */
export function splitNumberedName(name: string): NameParts {
    const matched = /^(.*?)(\d+)(\.[^.]+)?$/.exec(name);
    return {
        prefix: matched ? matched[1] : name,
        start: matched ? Number(matched[2]) : 1,
        suffix: matched ? (matched[3] ?? '') : '',
    };
}

/**
 * 生成当前目录内可用的文件名称
 *
 * 规则：基础名自带序号（{@link NEW_FILE_BASE_NAMES}，如 `sheet1.grid`），
 * 当前目录已存在同名时**递增其中的数字**：`sheet1.grid` → `sheet2.grid` → …
 * @param {string} todo - 菜单动作标识（XLSX / DOCX / PPTX）
 * @param {TreeNodeView | null} parent - 父节点（缺省为根节点）
 * @param {TreeNodeView | null} root - 根节点
 * @returns {string} 当前目录内不冲突的文件名
 */
export function availableFileName(
    todo: string,
    parent: TreeNodeView | null,
    root: TreeNodeView | null
): string {
    const target = parent?.childrenBox ? parent : root;
    const taken = new Set(target?.children.map((child) => child.key) ?? []);
    const base = NEW_FILE_BASE_NAMES[todo] ?? `${todo.toLowerCase()}1.${todo.toLowerCase()}`;

    const { prefix, start, suffix } = splitNumberedName(base);
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
 * @param {string} base - 基础名称（如 '新建文件夹'）
 * @param {TreeNodeView | null} parent - 父节点（缺省为根节点）
 * @param {TreeNodeView | null} root - 根节点
 * @returns {string} 当前目录内不冲突的名称
 */
export function availableName(
    base: string,
    parent: TreeNodeView | null,
    root: TreeNodeView | null
): string {
    const target = parent?.childrenBox ? parent : root;
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
 * 在指定父节点下生成不重复的名称（粘贴 / 登记文件时用于兜底去重）
 *
 * 重名规则与「新建文件」{@link availableFileName} 保持一致：
 * - 文件名中带数字（`sheet1.grid` / `doc1.text`）→ **递增其中的数字**；
 * - 文件名中无数字（`报表.grid`）→ 在扩展名前追加序号：`报表1.grid`；
 * - 文件夹名（`新建文件夹` / `归档`）→ 追加序号：`归档1`。
 * @param {TreeNodeView} parent - 父节点
 * @param {string} name - 期望名称
 * @returns {string} 去重后的名称
 */
export function uniqueName(parent: TreeNodeView, name: string): string {
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
 * 判断候选节点是否为给定视图的子孙
 * @param {TreeNodeView} view - 候选祖先
 * @param {TreeNodeView} target - 待查找节点
 * @returns {boolean} 是否为子孙
 */
export function isDescendantView(view: TreeNodeView, target: TreeNodeView): boolean {
    return view.children.some((child) => child === target || isDescendantView(child, target));
}

/**
 * 先序遍历子树
 * @param {TreeNodeView} view - 起始节点
 * @param {(node: TreeNodeView) => void} visit - 访问回调
 */
export function walk(view: TreeNodeView, visit: (node: TreeNodeView) => void): void {
    visit(view);
    view.children.forEach((child) => walk(child, visit));
}

/**
 * 将视图序列化为纯数据（供剪贴板暂存 / 持久化导出，切断与 DOM 的引用）
 * @param {TreeNodeView[]} views - 节点集合
 * @param {boolean} [withMeta=false] - 是否附带路径与元数据（持久化用；剪贴板不需要）
 * @returns {FileNode[]} 深拷贝后的数据
 */
export function serializeNodes(views: TreeNodeView[], withMeta = false): FileNode[] {
    return views.map((view) => {
        const name = view.label.textContent ?? '';
        const node: FileNode = { name };
        if (view.isFolder) {
            node.children = view.children.map((child) => serializeNodes([child], withMeta)[0]);
        }
        if (withMeta) {
            // 路径按父链推导（文件夹自带 `/` 前缀、文件不带）
            node.path = pathWithSelf(view);
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
 * 计算节点深度（根节点为 0）
 * @param {TreeNodeView} view - 节点
 * @returns {number} 深度
 */
export function getDepth(view: TreeNodeView): number {
    let depth = 0;
    for (let p = view.parent; p !== null; p = p.parent) depth++;
    return depth;
}

/**
 * 按路径 / 名称在整棵树中查找节点
 *
 * 匹配规则（不区分大小写，按优先级）：
 * 1. **完整路径一致**（`归档/销售数据.grid`）—— 最精确，可区分同名文件；
 * 2. 名称完全一致（`销售数据.grid`）；
 * 3. 「去掉扩展名后」一致 —— 仅当查询名本身**没有扩展名**时启用
 *    （让工作表名 `销售数据` 能定位到显示名 `销售数据.grid`）。
 *
 * 同名时文件优先于文件夹。路径命中优先于名称命中。
 * @param {TreeNodeView[]} roots - 根节点集合
 * @param {string} name - 文件路径 / 文件名 / 工作表名
 * @returns {TreeNodeView | null} 命中的节点
 */
export function findByName(roots: TreeNodeView[], name: string): TreeNodeView | null {
    const key = name.trim().toLowerCase();
    if (!key) return null;
    const bare = stripExtension(key);
    // 查询名自带扩展名时，禁止回落到「去扩展名」档：`报表.text` 不该命中 `报表.grid`
    const queryHasExt = extractExtension(key) !== '';
    const byPath: TreeNodeView[] = [];
    const byName: TreeNodeView[] = [];
    const fallback: TreeNodeView[] = [];
    roots.forEach((root) => walk(root, (view) => {
        const path = pathWithSelf(view).toLowerCase();
        if (path === key) byPath.push(view);
        if (view.key === key) byName.push(view);
        if (!queryHasExt && stripExtension(path) === bare) fallback.push(view);
    }));
    // 1) 路径精确命中 → 2) 名称精确命中（文件优先）→ 3) 无后缀查询的路径去扩展名兜底
    const exactPath = byPath.find((v) => !v.isFolder) ?? byPath[0];
    if (exactPath) return exactPath;
    return byName.find((v) => !v.isFolder) ?? byName[0]
        ?? fallback.find((v) => !v.isFolder) ?? fallback[0] ?? null;
}

/**
 * 查找首个命中的可见节点（先序）
 * @param {TreeNodeView[]} views - 节点集合
 * @param {string} keyword - 关键字（小写）
 * @returns {TreeNodeView | null} 命中节点
 */
export function findFirstMatch(views: TreeNodeView[], keyword: string): TreeNodeView | null {
    for (const view of views) {
        if (!view.el.hidden && view.key.includes(keyword)) return view;
        const inChildren = findFirstMatch(view.children, keyword);
        if (inChildren) return inChildren;
    }
    return null;
}
