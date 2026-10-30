/*  # 工作表数据层职责拆分（K3）
    # 从 dataCollection.ts 原样迁出，仅调整类继承关系与成员可见性，逻辑零改动
*/
import { Cell } from "../dataArchitecture/cell";
import { Char } from "../dataArchitecture/char";
import type { SheetRect, SearchHit, SearchMatchOptions } from "../dataArchitecture/dataCollection";
import { SheetFilterBase } from "./sheetFilter";
import { SheetStoreBase } from "./sheetStore";

/**
 * 稀疏扫描的候选格（内部用，不对外暴露）
 *
 * 与 {@link SearchHit} 的差别是多了 `ref`（`values` 里的**活引用**，替换阶段直接改它）
 * 与 `key`（行优先序号，仅用于排序）。分开定义是为了不让「活引用」泄进公开的命中项里：
 * 命中项会被对话框长期持有，带着一个可写的数据引用迟早出事。
 */
interface ScanCandidate {
/** 单元格引用（`values` 里的活引用） */
ref: Cell;
/** 单元格名称 */
cell: string;
/** 列号（1-based） */
col: number;
/** 行号（1-based） */
row: number;
/** 逐字符对齐的文本（与 `Cell.getText()` 结果一致） */
text: string;
/** 行优先序号 = `(row - startRow) * width + (col - startCol)` */
key: number;
}

/**
 * 查询 / 替换基类
 *
 * 全部限定在给定矩形范围内；扫描只遍历 values（真实存在的单元格）。
 */
export abstract class SheetSearchBase extends SheetFilterBase {
    /* ---------- 仍定义在 DataCollection 中、供查询/替换调用的成员 ---------- */
    public abstract scrollCellIntoView(col: number, row: number): boolean;
    public abstract runWithCellsUndo(label: string, cells: Cell[], action: () => void): void;
    /* ==================== 查询与替换 ====================
     *
     * 「工具栏 → 查询 / 替换」的数据层实现。**全部限定在给定矩形范围内**，
     * UI 层（toolbar/start/searchDialog.ts）只负责取值与展示：
     * - 范围是显式参数，不读也不改当前选区（`gotoSearchHit` 除外），
     *   因此「跳转命中项（会改写选区）」之后继续「查找下一个」不会把范围缩小成一个格子；
     * - **扫描只遍历 `values`（真实存在的单元格）**，不遍历 行 × 列 网格 ——
     *   见 `_collectSearchCells`。整表范围是常态（单格选区会被解释成整表，
     *   见 `expandSingleCellRect`），而 10 万行 × 100 列的表按网格遍历是**亿级**空转；
     * - 替换的撤销补丁按**命中的格子**逐个捕获（`runWithCellsUndo`），不是按矩形逐格捕获：
     *   命中天然稀疏，矩形版在整表范围下会为每一格都克隆一份快照；
     * - 行/列的隐藏状态按 copySelection 的口径：**隐藏行列不参与查找**，
     *   否则筛选状态下会命中用户看不见的格子。
     * ==================== */

    /**
     * 解析一段区域串（`'B3'` 或 `'A1:C9'`，大小写不敏感）
     *
     * 只认这两种形态；任何一处不合法返回 null，由调用方决定如何提示
     * （不要让本方法抛错：范围串直接来自对话框输入框，用户可能手打）。
     * @param {string} range - 区域串
     * @returns {SheetRect | null} 规范化后的矩形（1-based 闭区间，端点已按大小端排序）
     */
    protected _parseRect(range: string): SheetRect | null {
        const parts = range.split(':');
        if (parts.length > 2) return null;
        const a = this._parseCellRef(parts[0]);
        if (!a) return null;
        if (parts.length === 1) {
            return { startCol: a.col, startRow: a.row, endCol: a.col, endRow: a.row };
        }
        const b = this._parseCellRef(parts[1]);
        if (!b) return null;
        return {
            startCol: Math.min(a.col, b.col),
            startRow: Math.min(a.row, b.row),
            endCol: Math.max(a.col, b.col),
            endRow: Math.max(a.row, b.row)
        };
    }

    /**
     * 解析单个单元格引用（`'AB12'` → `{ col: 28, row: 12 }`）
     *
     * 手写这里而不是复用 {@link getCellColAndRow}：后者对 `'A0'`、`''`、`'1A'`
     * 这类输入会返回 `row: 0` / `col: 0` / `NaN` 而不报错，作为范围来源会让
     * 「查询第 0 行」这种不存在的结果悄悄进入扫描循环。
     * @param {string} ref - 单元格引用
     * @returns {{ col: number, row: number } | null} 行列号（1-based）；非法返回 null
     */
    private _parseCellRef(ref: string): { col: number; row: number } | null {
        const trimmed = ref.trim();
        const m = /^([A-Za-z]{1,3})([1-9][0-9]{0,6})$/.exec(trimmed);
        if (!m) return null;
        const col = this.getColNumber(m[1].toUpperCase());
        const row = parseInt(m[2], 10);
        if (!isFinite(col) || col < 1 || !isFinite(row) || row < 1) return null;
        return { col, row };
    }

    /**
     * 决定本次查询 / 替换的作用范围
     *
     * 三级回退：
     * ① 显式区域串（对话框「范围」输入框的值优先，让用户在 dialog 里点选或手填的范围说了算）；
     * ② 当前选区（工具栏入口的默认语义 = **只在当前选中区域内查找**）；
     * ③ 无选区时回退到**有内容的区域**（{@link getContentBounds}），而不是行列表头的
     *    逻辑上限（1048575 × 16384）——后者会让一次查询扫过整张虚拟表。
     * @param {string} [range] - 显式区域串
     * @returns {SheetRect | null} 查询范围；连空表都没有时返回 null
     */
    public resolveSearchRect(range?: string): SheetRect | null {
        if (range && range.trim()) {
            const rect = this._parseRect(range);
            if (rect) return rect;
        }
        if (this.selection) {
            const rect = this._parseRect(this.selection);
            if (rect) return rect;
        }
        const { maxRow, maxCol } = this.getContentBounds();
        if (!maxRow || !maxCol) return null;
        return { startCol: 1, startRow: 1, endCol: maxCol, endRow: maxRow };
    }

    /**
     * 把范围矩形格式化成 `'A1:C9'`（单格时为 `'A1:A1'`，与 selection 的既有口径一致）
     * @param {SheetRect} rect - 范围矩形
     * @returns {string} 区域串
     */
    public formatRect(rect: SheetRect): string {
        const start = this.getCellName(rect.startCol, rect.startRow);
        const end = this.getCellName(rect.endCol, rect.endRow);
        return `${start}:${end}`;
    }

    /**
     * 把「单格范围」解释为「整表范围 + 起点锚」
     *
     * 选中单个单元格之后查找，要找的显然不是这一格本身（它的内容就摆在眼前），
     * 而是**整张表**里还有哪些地方出现同样的内容；「查找下一个」从这一格往下走、
     * 到头回绕 —— 与 Excel「选中一格 → Ctrl+F → 查找全部/下一个」的语义一致。
     *
     * 返回 `null` 表示**保持原样**：范围不是单格，或者表里根本没有内容
     * （`getContentBounds()` 为 0，此时连一个候选格都没有，扩不扩张结果都一样）。
     *
     * 注意这是 `DataCollection` 的语义、不是「范围串恰好写成单格」的判定：
     * 由主表格选出的单格走本方法（范围会被**显示**成展开后的整表范围，用户看得见
     * 将要搜多大一片）；用户手工在范围框里敲 `'B2:B2'` 时按字面处理，就是只查那一格。
     * @param {SheetRect} rect - 待检查的范围（1-based 闭区间）
     * @returns {{ rect: SheetRect; anchor: { col: number; row: number } } | null} 整表范围 + 起点锚；不适用时 null
     */
    public expandSingleCellRect(rect: SheetRect): { rect: SheetRect; anchor: { col: number; row: number } } | null {
        if (rect.startCol !== rect.endCol || rect.startRow !== rect.endRow) return null;
        const { maxRow, maxCol } = this.getContentBounds();
        if (!maxRow || !maxCol) return null;
        return {
            rect: { startCol: 1, startRow: 1, endCol: maxCol, endRow: maxRow },
            anchor: { col: rect.startCol, row: rect.startRow }
        };
    }

    /**
     * 构造查询用正则
     *
     * - `useWildcard`：`?` → 单个字符、`*` → 任意长度（含换行，故用 `[\s\S]` 而非 `.`），
     *   其余字符按正则元字符转义；
     * - `matchEntire`：首尾锚定（「单元格内容完全匹配」而非包含即可）；
     * - `matchCase`：不加 `i` 标志。
     * 统一返回**带 g 标志**的正则：命中判定与「把所有匹配位置替换掉」共用同一个 「lastIndex 游标」语义。
     * @param {string} keyword - 关键字
     * @param {SearchMatchOptions} [options] - 匹配选项
     * @returns {RegExp | null} 正则；关键字为空返回 null（空关键字不构成查询）
     */
    private _buildSearchPattern(keyword: string, options?: SearchMatchOptions): RegExp | null {
        if (!keyword) return null;
        const escape = SheetSearchBase._escapeRegExp;
        let source = '';
        if (options?.useWildcard === true) {
            for (let i = 0; i < keyword.length; i++) {
                const ch = keyword[i];
                source += ch === '?' ? '[\\s\\S]'
                    : ch === '*' ? '[\\s\\S]*'
                        : escape(ch);
            }
        } else {
            source = escape(keyword);
        }
        if (options?.matchEntire === true) source = `^(?:${source})$`;
        try {
            return new RegExp(source, options?.matchCase === true ? 'g' : 'gi');
        } catch (e) {
            return null;
        }
    }

    /**
     * 转义正则元字符（`'a?b'` → `'a\\?b'`）
     *
     * **刻意不用 `text.replace(/[...]/g, '\\$&')` 这种字符类写法**：
     * 字符类里要同时含 `]` 与 `\`，写错一个（例如少写 `]`，或把 `]` 放在 `\` 之后）
     * 就会静默改变语义 —— 而这里一旦漏转义 `?`，通配符模式下的 `'a?b'` 会退化成
     * 字面量 "a?b"（`?` 只对前一个字符生效，于是永远查不到东西），
     * 表面上「查询没结果」，排查起来极费时间。
     * 下面用显式字符集 + `includes` 分派：没有字符类、没有一层转义，读一眼就知道对不对。
     * @param {string} text - 原文本
     * @returns {string} 转义后文本
     */
    private static _escapeRegExp(text: string): string {
        const SPECIAL = '.?*+^$()[]{}|\\';
        let out = '';
        for (let i = 0; i < text.length; i++) {
            const ch = text[i];
            out += SPECIAL.indexOf(ch) >= 0 ? '\\' + ch : ch;
        }
        return out;
    }

    /**
     * 取出单元格**逐字符对齐**的文本
     *
     * 与 {@link Cell.getText} 的唯一差别：这里对每个字符都产出一个位（哪怕 `char` 为
     * null/undefined 也记一个下标），保证「文本下标 ↔ chars 下标」一一对应 ——
     * 替换时要按匹配的下标去切 `chars`，两者错位会把样式/字符切错乱。
     * @param {Cell} cell - 单元格
     * @returns {string} 对齐文本（与 getText 的结果一致）
     */
    private _alignedCellText(cell: Cell): string {
        const chars = cell.chars;
        if (!chars || chars.length === 0) return '';
        let text = '';
        for (let i = 0; i < chars.length; i++) {
            const ch = chars[i];
            if (ch && ch.char !== null && ch.char !== undefined) text += ch.char;
        }
        return text;
    }

    /**
     * 稀疏收集矩形内的候选格（**行优先序**）
     *
     * ## 性能约定（勿回退为「按行列号逐格 getCellName + getCellByName」）
     * 历史实现是 `for row { for col { getCellName(...) → getCellByName(...) } }`：
     * 每次迭代都要拼一个字符串、再查一次 Map，**空网格也要一格一格走过**。
     * 整表范围（单格选区展开后的常态）下这是 行数 × 列数 次迭代 ——
     * 10 万行 × 100 列 = 一千万次，而其中 99% 的格子在 `values` 里根本不存在。
     *
     * 这里只遍历 `values`（O(存在的格子数)），顺序由行列号算出；`values` 恰好已是
     * 行优先时（装载 / 顺序编辑的主流情形）**不排序**，否则按行优先键排一次 ——
     * 命中列表必须按「行优先」呈现（查找下一个要自上而下、自左向右走）。
     * 「是否已有序」在扫描过程中顺手判定，不需要额外的探测遍历。
     *
     * 过滤顺序刻意从便宜到贵：非空文本 → 解析行列号 → 范围 → 隐藏 → 拼文本。
     * @param {SheetRect} rect - 范围矩形（1-based 闭区间）
     * @returns {ScanCandidate[]} 候选格（行优先；已排除隐藏行列与无文本的格子）
     * @private
     */
    private _collectSearchCells(rect: SheetRect): ScanCandidate[] {
        const values = this._values;
        const out: ScanCandidate[] = [];
        const { startCol, startRow, endCol, endRow } = rect;
        const width = endCol - startCol + 1;
        let ordered = true;
        let lastKey = -1;
        for (let i = 0; i < values.length; i++) {
            const cell = values[i];
            // 最便宜的判定先行：只有样式 / 空 Cell 的格子直接跳过，连名字都不用解析
            if (!SheetStoreBase.hasAnyText(cell)) continue;
            const { col, row } = this.getCellColAndRow(cell.cell);
            if (col < startCol || col > endCol || row < startRow || row > endRow) continue;
            if (this.isRowHidden(row - 1) || this.isColHidden(col - 1)) continue;
            const key = (row - startRow) * width + (col - startCol);
            if (key < lastKey) ordered = false;
            lastKey = key;
            out.push({
                ref: cell,
                cell: cell.cell,
                col,
                row,
                text: this._alignedCellText(cell),
                key
            });
        }
        if (!ordered) out.sort((a, b) => a.key - b.key);
        return out;
    }

    /**
     * 在指定范围内找出**所有**命中项（行优先：先逐列扫完一行，再进下一行）
     *
     * 每个单元格最多产出一条命中（即便格内有多处匹配）—— 与 Excel 的命中列表口径一致；
     * 「一处 Cell 内多处匹配」只在替换时展开处理（见 {@link replaceInRect}）。
     * @param {SheetRect} rect - 查询范围
     * @param {string} keyword - 关键字
     * @param {SearchMatchOptions} [options] - 匹配选项
     * @param {number} [limit=Infinity] - 最多返回多少条（用于 UI 上限保护）
     * @returns {SearchHit[]} 命中列表（行优先升序）
     */
    public findAllInRect(
        rect: SheetRect,
        keyword: string,
        options?: SearchMatchOptions,
        limit: number = Infinity
    ): SearchHit[] {
        const pattern = this._buildSearchPattern(keyword, options);
        const hits: SearchHit[] = [];
        if (!pattern) return hits;
        const candidates = this._collectSearchCells(rect);
        for (let i = 0; i < candidates.length; i++) {
            const candidate = candidates[i];
            if (!this._matchesPattern(pattern, candidate.text)) continue;
            hits.push({ cell: candidate.cell, col: candidate.col, row: candidate.row, text: candidate.text });
            if (hits.length >= limit) break;
        }
        return hits;
    }

    /**
     * 用正则做一次「是否命中」判定（带 g 标志，必须先归零 lastIndex）
     * @param {RegExp} pattern - {@link _buildSearchPattern} 产出的正则
     * @param {string} text - 待判定文本
     * @returns {boolean} 是否命中
     */
    private _matchesPattern(pattern: RegExp, text: string): boolean {
        pattern.lastIndex = 0;
        return pattern.test(text);
    }

    /**
     * 在范围内找出 `after` 之后的**第一条**命中（找不到则回绕到首条）
     *
     * 与 {@link findAllInRect} 同源（同一次稀疏扫描），只是不建整张命中列表 ——
     * 调用方只想要「下一个」时没必要为整个范围分配数组。
     * 候选格自身已按行优先有序，因此「首个序号大于起点」的命中就是要找的那条。
     * @param {SheetRect} rect - 查询范围
     * @param {string} keyword - 关键字
     * @param {SearchMatchOptions} [options] - 匹配选项
     * @param {{ col: number; row: number }} [after] - 起点（不含）；缺省从头找，即返回首条命中
     * @returns {SearchHit | null} 命中项；无命中返回 null
     */
    public findNextInRect(
        rect: SheetRect,
        keyword: string,
        options?: SearchMatchOptions,
        after?: { col: number; row: number } | null
    ): SearchHit | null {
        const pattern = this._buildSearchPattern(keyword, options);
        if (!pattern) return null;
        const width = rect.endCol - rect.startCol + 1;
        const afterKey = after
            ? (after.row - rect.startRow) * width + (after.col - rect.startCol)
            : -1;
        const candidates = this._collectSearchCells(rect);
        let wrapTarget: SearchHit | null = null;
        for (let i = 0; i < candidates.length; i++) {
            const candidate = candidates[i];
            if (!this._matchesPattern(pattern, candidate.text)) continue;
            const hit = {
                cell: candidate.cell,
                col: candidate.col,
                row: candidate.row,
                text: candidate.text
            };
            // 回绕目标 = 首条命中（行优先序，故第一次赋值即最终值）
            if (wrapTarget === null) wrapTarget = hit;
            if (candidate.key > afterKey) return hit;
        }
        return wrapTarget;
    }

    /**
     * 判定某格**当前**是否命中，命中则返回它的命中项
     *
     * 供「替换」之后局部刷新命中列表用：替换只改动**一个**格子，重扫整个范围
     * 在整表尺度上是几百毫秒的代价，而「替换」恰恰是连点操作。
     * 本方法是单格的 O(1) 检查（就表规模而言）。
     * @param {string} name - 单元格名称（如 `'B3'`）
     * @param {string} keyword - 关键字
     * @param {SearchMatchOptions} [options] - 匹配选项
     * @returns {SearchHit | null} 命中项；不命中 / 格子不存在 / 落在隐藏行列上时返回 null
     */
    public cellMatches(name: string, keyword: string, options?: SearchMatchOptions): SearchHit | null {
        const pattern = this._buildSearchPattern(keyword, options);
        if (!pattern) return null;
        const { col, row } = this.getCellColAndRow(name);
        if (!isFinite(col) || !isFinite(row) || col < 1 || row < 1) return null;
        if (this.isRowHidden(row - 1) || this.isColHidden(col - 1)) return null;
        const cell = this.getCellByName(name);
        if (!cell) return null;
        const text = this._alignedCellText(cell);
        if (text === '' || !this._matchesPattern(pattern, text)) return null;
        return { cell: name, col, row, text };
    }

    /**
     * 在范围内把关键字替换为目标串（一步可撤销）
     *
     * 整体包在 {@link runWithCellsUndo} 内：撤销补丁按**命中的格子**逐个捕获，
     * 代价与命中处数成正比、与范围面积无关 —— 整表范围（单格选区展开后的常态）
     * 里只有 3 处命中时，也只克隆这 3 个格子，因此百万格的大表照样能 Ctrl+Z / Ctrl+Y
     * （全量快照超过 {@link LARGE_TABLE_UNDO_CELLS} 会被关闭）。
     *
     * 分两阶段：**先只读扫描收集待改格子**，再在补丁作用域内统一改写。
     * 边扫边改会让补丁捕获到的「变化前」已经是被改过的内容（撤销失效）；
     * 而把扫描塞进 `action` 里则无从预知哪些格子会变，只能退化成整矩形捕获。
     *
     * 合并区域内部的格子通常已被 merge 删除，天然不会被重复替换；若保留着
     * （带边框/底色的内部格子，见合并铁律）则按普通格子处理。
     * @param {SheetRect} rect - 替换范围
     * @param {string} keyword - 关键字
     * @param {string} replacement - 替换内容（空串 = 删除匹配文本）
     * @param {SearchMatchOptions} [options] - 匹配选项（与查询共用，保证「查得到才替换得掉」）
     * @param {SearchHit[] | null} [targets] - 仅替换这些单元格（「替换」按钮）；缺省替换范围内全部命中（「全部替换」）
     * @param {string} [label='替换'] - 撤销操作标签
     * @returns {number} 实际替换的**匹配处数**（一个格内多处匹配计多处）
     */
    public replaceInRect(
        rect: SheetRect,
        keyword: string,
        replacement: string,
        options?: SearchMatchOptions,
        targets?: SearchHit[] | null,
        label: string = '替换'
    ): number {
        const pattern = this._buildSearchPattern(keyword, options);
        if (!pattern) return 0;
        let nameFilter: Set<string> | null = null;
        if (targets) {
            nameFilter = new Set<string>();
            for (let i = 0; i < targets.length; i++) nameFilter.add(targets[i].cell);
        }
        // 阶段一（只读）：确定「哪些格子会被改动」—— 一次稀疏扫描
        const candidates = this._collectSearchCells(rect);
        const affected: Cell[] = [];
        for (let i = 0; i < candidates.length; i++) {
            const candidate = candidates[i];
            if (nameFilter && !nameFilter.has(candidate.cell)) continue;
            if (!this._matchesPattern(pattern, candidate.text)) continue;
            affected.push(candidate.ref);
        }
        if (affected.length === 0) return 0;
        // 阶段二：补丁作用域内改写。两个阶段之间没有让出，文本快照仍然有效。
        let count = 0;
        this.runWithCellsUndo(label, affected, () => {
            for (let i = 0; i < affected.length; i++) {
                count += this._replaceCharsInCell(affected[i], pattern, replacement);
            }
        });
        return count;
    }

    /**
     * 把单个单元格里**所有**匹配片段替换掉
     *
     * 字符级样式保护：被替换片段原位插入到 a+b+c 的位置；原片段若被替换为空串则该片段直接删除。
     * @param {Cell} cell - 目标单元格（values 里的活引用）
     * @param {RegExp} pattern - 带 g 标志的正则
     * @param {string} replacement - 替换内容
     * @returns {number} 替换的片段数
     */
    private _replaceCharsInCell(cell: Cell, pattern: RegExp, replacement: string): number {
        const chars = cell.chars;
        if (!chars || chars.length === 0) return 0;
        const text = this._alignedCellText(cell);
        if (text === '') return 0;
        // 先收集所有片段位置（不重叠，由正则自身的 g 语义保证），再一次性重建 chars：
        // 边替换边改 chars 会让后续片段的下标全部失效。
        //
        // 【性能约定，勿回退为「先收集位置、再统一重建 chars」的两趟写法】
        // 常见的替换（范围内每格都是 `old → new` 这类单片段、且首字符就命中）走下面这条
        // 快路径：原地改字符、样式留空（保留各处样式差异）。只有**多处片段**才退到
        // 重建路径 —— 「全部替换」动辄上万格，重建会为每格都新分配一个 chars 数组，
        // 而这些格子往往一个字符都不需要换。
        const replacementChars = Array.from(replacement);
        pattern.lastIndex = 0;
        const first = pattern.exec(text);
        if (first === null) return 0;
        if (first.index === 0 && first[0].length === text.length) {
            for (let i = 0; i < chars.length; i++) {
                chars[i].char = i < replacementChars.length ? replacementChars[i] : '';
            }
            // 短替换：多出来的字符（char = ''）参与 isEmpty 判定时会一并释放（'' 不算文字信息）
            // 长替换：样式留空，由渲染层按单元格/默认字体兜底
            return 1;
        }
        const starts: number[] = [first.index];
        const ends: number[] = [first.index + first[0].length];
        if (first[0].length === 0) pattern.lastIndex++;
        let m: RegExpExecArray | null;
        while ((m = pattern.exec(text)) !== null) {
            starts.push(m.index);
            ends.push(m.index + m[0].length);
            if (m[0].length === 0) pattern.lastIndex++; // 零宽匹配：手动前进，否则死循环
        }
        if (starts.length === 1 && ends[0] - starts[0] === text.length) {
            // 全格被替换但首字符不命中（如完全匹配时首字符被通配符吃掉）：同样走原位路径
            for (let i = 0; i < chars.length; i++) {
                chars[i].char = i < replacementChars.length ? replacementChars[i] : '';
            }
            return 1;
        }
        const out: Char[] = [];
        let cursor = 0;
        for (let i = 0; i < starts.length; i++) {
            const start = starts[i];
            const end = ends[i];
            if (start > cursor) {
                for (let k = cursor; k < start; k++) out.push(chars[k]);
            }
            if (replacementChars.length > 0) {
                // 继承被替换片段**首字符**的样式，替换进去的字保留原格式（字号/字色/加粗…）
                const styleSource = chars[start] || chars[cursor] || new Char();
                for (let k = 0; k < replacementChars.length; k++) {
                    const char = styleSource.clone();
                    char.char = replacementChars[k];
                    out.push(char);
                }
            }
            cursor = end;
        }
        for (let k = cursor; k < chars.length; k++) out.push(chars[k]);
        cell.chars = out;
        return starts.length;
    }

    /**
     * 跳转到指定单元格（活动单元格 + 单格选区 + 滚动入可视区）
     *
     * 供「查找下一个 / 上一个」跳转命中项。注意它**会改写选区**，所以调用方必须
     * 事先把查询范围记住（searchDialog 用一个输入控件持有范围串），
     * 否则下一次查找的范围就会被这次跳转缩小成一个格子。
     * @param {SearchHit} hit - 命中项
     */
    public gotoSearchHit(hit: SearchHit): void {
        this.activedCell = hit.cell;
        this.selection = `${hit.cell}:${hit.cell}`;
        this.scrollCellIntoView(hit.col, hit.row);
    }
}
