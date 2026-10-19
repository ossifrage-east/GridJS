/*  # 文件名工具 TS 文件
    # 统一「显示文件名」与「工作表名」之间的换算规则
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

/**
 * 提取文件扩展名（小写，不含点）
 *
 * 无扩展名时返回空字符串；以点开头的隐藏文件名（如 `.gitignore`）视为无扩展名。
 * @param name - 文件名，如 `sheet1.XLSX`
 * @returns {string} 扩展名，如 `xlsx`
 * @example
 * extractExtension('sheet1.XLSX');  // 'xlsx'
 * extractExtension('归档');          // ''
 */
export function extractExtension(name: string): string {
    const dot = (name ?? '').lastIndexOf('.');
    return dot > 0 ? name.slice(dot + 1).toLowerCase() : '';
}

/**
 * 去掉文件扩展名，得到基础名（工作表名）
 *
 * 与 {@link extractExtension} 使用同一套判定：无扩展名时原样返回，
 * 因此「显示文件名」与「工作表名」的换算在全局保持一致。
 * @param name - 文件名，如 `销售数据.xlsx`
 * @returns {string} 基础名，如 `销售数据`
 * @example
 * stripExtension('销售数据.xlsx');  // '销售数据'
 * stripExtension('Sheet1');        // 'Sheet1'
 */
export function stripExtension(name: string): string {
    const dot = (name ?? '').lastIndexOf('.');
    return dot > 0 ? name.slice(0, dot) : (name ?? '');
}

/** 应用的三种原生文件后缀：`grid`（表格）/ `text`（文档）/ `demo`（演示文稿） */
export const NATIVE_EXTENSIONS = ['grid', 'text', 'demo'] as const;

/**
 * 外部真实格式 → 应用原生后缀
 *
 * 应用只有三种后缀，各自对应一类内容载体。打开外部文件时按此表把真实扩展名
 * 换算为原生后缀，使同类内容在文件树里呈现统一命名：
 * - `xls` / `xlsx` → `grid`
 * - `doc` / `docx` → `text`
 * - `ppt` / `pptx` → `demo`
 */
export const EXTENSION_TO_NATIVE: Record<string, typeof NATIVE_EXTENSIONS[number]> = {
    xls: 'grid', xlsx: 'grid',
    doc: 'text', docx: 'text',
    ppt: 'demo', pptx: 'demo',
};

/**
 * 把外部文件名换算为应用内的原生文件名
 *
 * 例：`报表.xlsx` → `报表.grid`、`说明.docx` → `说明.text`、`路演.pptx` → `路演.demo`。
 * 不在 {@link EXTENSION_TO_NATIVE} 内的后缀原样返回（不臆造后缀）。
 * @param name - 外部文件名（如 `报表.xlsx`）
 * @returns {string} 原生文件名（如 `报表.grid`）；名称为空时原样返回
 * @example
 * toNativeFileName('报表.xlsx');  // '报表.grid'
 * toNativeFileName('报表.grid');  // '报表.grid'（已是原生后缀，无映射则原样）
 */
export function toNativeFileName(name: string): string {
    const text = (name ?? '').trim();
    if (!text) return text;
    const native = EXTENSION_TO_NATIVE[extractExtension(text)];
    return native ? `${stripExtension(text)}.${native}` : text;
}

/**
 * 应用的**原生表格后缀**（`grid`）
 *
 * 「新建表格」产出的即此格式（见 fileTree 的 `NEW_FILE_BASE_NAMES`）；
 * 历史上的 `*.xlsx` 在打开时会换算为 `*.grid`，二者承担同一角色
 * （见 {@link toArchiveKey} 的存量兼容说明）。
 */
export const NATIVE_GRID_EXTENSION = 'grid';

/**
 * 把文件**路径**换算为内容存档键（IndexedDB `values` / `images` 库的键）
 *
 * 规则：**只去掉应用原生表格后缀（{@link NATIVE_GRID_EXTENSION}，即 `.grid`）；
 * 其它后缀一律保留**。
 * - `sheet1.grid`  → `sheet1`        （与历史 `.xlsx` 存档键完全一致，存量零迁移）
 * - `归档/销售数据.grid` → `归档/销售数据`（带目录：不同目录的同名文件各自独立存档）
 * - `报表.text`    → `报表.text`      （保留后缀）
 * - `报表`         → `报表`           （无后缀，原样）
 *
 * **为什么只去 `.grid`**：`报表.grid` 与 `报表.text` 去扩展名后同为 `报表`，
 * 若都去扩展名就会共用一份存档、互相覆盖。只让「原生表格格式」沿用历史的去扩展名键，
 * 其它后缀保留后缀，即可在**不迁移任何存量存档**的前提下让二者互不干扰。
 * 历史存量 `.xlsx` 的键（`sheet1.xlsx` → `sheet1`）在换算为 `.grid` 后仍是 `sheet1`，
 * 因此老存档无需迁移即可继续命中。
 *
 * **为什么必须只由路径决定（不能看「树里有没有同名冲突」）**：
 * 删除同名的兄弟文件后，冲突就消失了 —— 若键依赖树状态，该文件的键会立刻改变
 * （`报表.grid` → `报表`），于是「打开时用的键」与「删除/回退时算出的键」不一致，
 * 表现为删不掉、回退到空白文件、内容丢失。因此存档键必须是**路径的纯函数**。
 * @param path - 文件路径（含扩展名，相对根节点）
 * @returns {string} 存档键；名称为空时返回空串
 */
export function toArchiveKey(path: string): string {
    const text = (path ?? '').trim();
    if (!text) return '';
    return extractExtension(text) === NATIVE_GRID_EXTENSION ? stripExtension(text) : text;
}
