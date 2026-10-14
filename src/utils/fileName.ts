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

/**
 * 应用的**原生表格后缀**
 *
 * 「新建文件」产出的即此格式（见 fileTree 的 `NEW_FILE_BASE_NAMES`），
 * 因此历史上库里所有存档都对应一个 `*.xlsx` 文件。
 */
export const NATIVE_SHEET_EXTENSION = 'xlsx';

/**
 * 把文件**路径**换算为内容存档键（IndexedDB `values` / `images` 库的键）
 *
 * 规则：**只去掉应用原生表格后缀（{@link NATIVE_SHEET_EXTENSION}，即 `.xlsx`）；
 * 其它后缀一律保留**。
 * - `sheet1.xlsx`  → `sheet1`        （与历史存档键完全一致，存量零迁移）
 * - `归档/销售数据.xlsx` → `归档/销售数据`（带目录：不同目录的同名文件各自独立存档）
 * - `报表.docx`    → `报表.docx`      （保留后缀）
 * - `报表`         → `报表`           （无后缀，原样）
 *
 * **为什么只去 `.xlsx`**：`报表.xlsx` 与 `报表.docx` 去扩展名后同为 `报表`，
 * 若都去扩展名就会共用一份存档、互相覆盖。只让「原生格式」沿用历史的去扩展名键，
 * 其它后缀保留后缀，即可在**不迁移任何存量存档**的前提下让二者互不干扰。
 *
 * **为什么必须只由路径决定（不能看「树里有没有同名冲突」）**：
 * 删除同名的兄弟文件后，冲突就消失了 —— 若键依赖树状态，该文件的键会立刻改变
 * （`报表.xlsx` → `报表`），于是「打开时用的键」与「删除/回退时算出的键」不一致，
 * 表现为删不掉、回退到空白文件、内容丢失。因此存档键必须是**路径的纯函数**。
 * @param path - 文件路径（含扩展名，相对根节点）
 * @returns {string} 存档键；名称为空时返回空串
 */
export function toArchiveKey(path: string): string {
    const text = (path ?? '').trim();
    if (!text) return '';
    return extractExtension(text) === NATIVE_SHEET_EXTENSION ? stripExtension(text) : text;
}
