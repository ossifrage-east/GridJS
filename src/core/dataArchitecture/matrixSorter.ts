/*  # 矩阵排序器组件 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { Cell } from './cell';
import { DataCollection } from './dataCollection';

/**
 * 排序配置
 */
interface SortConfig {
    /** 排序轴：'column' = 按某列的值排序行（列排序），'row' = 按某行的值排序列（行排序） */
    axis: 'column' | 'row';
    /** 排序依据的列/行索引（0-based） */
    keyIndex: number;
    /** true=升序，false=降序 */
    ascending: boolean;
    /** 连续范围起始索引（0-based，包含）。与 endIndex 配合使用表示连续范围 */
    startIndex?: number;
    /** 连续范围结束索引（0-based，包含） */
    endIndex?: number;
    /** 不连续索引数组（0-based）。提供时优先于 startIndex/endIndex，表示仅对这些位置排序 */
    indices?: number[];
}

/**
 * 单元格内容的数据类型分类，枚举值即混合排序的优先级（升序时从小到大）
 */
enum CellDataType {
    /** 空值（空单元格或纯空白文本） */
    Empty = 0,
    /** 时间（日期、日期时间、纯时间） */
    Time = 1,
    /** 数字（整数、小数、科学计数法） */
    Number = 2,
    /** 英文字符 */
    English = 3,
    /** 中文字符 */
    Chinese = 4,
    /** 其它外语字符 */
    Foreign = 5,
}

/** 日期(+可选时间)正则：2024-01-01、2024/1/1、2024年1月1日，可附带 12:30(:45) */
const DATE_TIME_PATTERN = /^(\d{4})[-/年.](\d{1,2})[-/月.](\d{1,2})日?(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;

/** 纯时间正则：12:30、12:30:45 */
const TIME_ONLY_PATTERN = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/;

/** 数字正则：整数、小数、正负号、科学计数法 */
const NUMBER_PATTERN = /^[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?$/;

/** 中文字符正则：CJK 统一表意文字及其扩展 A、兼容表意文字 */
const CHINESE_PATTERN = /[\u4e00-\u9fff\u3400-\u4dbf\uf900-\ufaff]/;

/**
 * 判定单元格文本的数据类型分类
 *
 * 判定顺序：空值 → 数字 → 时间 → 按首字符判英文/中文/其它外语。
 * 数字优先于时间（如 20240101 是数字而非日期）；混合内容以首字符类型为准。
 * @param {string} text - 单元格文本（内部会 trim）
 * @returns {CellDataType} 类型分类
 */
function classifyCellText(text: string): CellDataType {
    const t = text.trim();
    if (t === '') return CellDataType.Empty;
    if (NUMBER_PATTERN.test(t)) return CellDataType.Number;
    if (DATE_TIME_PATTERN.test(t) || TIME_ONLY_PATTERN.test(t)) return CellDataType.Time;

    const first = t[0];
    if (/[A-Za-z]/.test(first)) return CellDataType.English;
    if (CHINESE_PATTERN.test(first)) return CellDataType.Chinese;
    return CellDataType.Foreign;
}

/**
 * 解析时间文本为可比较的数值
 *
 * - 纯时间（hh:mm(:ss)）→ 当日秒数，便于纯时间之间比较；
 * - 日期(+可选时间) → Date.UTC 时间戳（按 UTC 换算，不受运行环境时区影响）；
 * - 月/日超出有效范围或格式不符 → null。
 * @param {string} text - 时间文本
 * @returns {number | null} 可比较的数值；无法解析时返回 null
 */
function parseTime(text: string): number | null {
    const timeOnly = TIME_ONLY_PATTERN.exec(text);
    if (timeOnly) {
        return Number(timeOnly[1]) * 3600 + Number(timeOnly[2]) * 60 + Number(timeOnly[3] ?? 0);
    }

    const dt = DATE_TIME_PATTERN.exec(text);
    if (!dt) return null;
    const month = Number(dt[2]);
    const day = Number(dt[3]);
    if (month < 1 || month > 12 || day < 1 || day > 31) return null;
    return Date.UTC(Number(dt[1]), month - 1, day, Number(dt[4] ?? 0), Number(dt[5] ?? 0), Number(dt[6] ?? 0));
}

/**
 * 比较两个时间文本
 *
 * 均可解析时按时间先后比较；任一无法解析时退化为字符串比较。
 * @param {string} a - 时间文本 a
 * @param {string} b - 时间文本 b
 * @returns {number} 比较结果
 */
function compareTime(a: string, b: string): number {
    const timeA = parseTime(a);
    const timeB = parseTime(b);
    if (timeA !== null && timeB !== null) return timeA - timeB;
    return a.localeCompare(b);
}

/**
 * 同类型文本的比较规则
 * @param {CellDataType} type - 数据类型分类
 * @param {string} a - 已 trim 的文本 a
 * @param {string} b - 已 trim 的文本 b
 * @returns {number} 比较结果
 */
function compareSameType(type: CellDataType, a: string, b: string): number {
    switch (type) {
        case CellDataType.Time:
            return compareTime(a, b);
        case CellDataType.Number:
            return Number(a) - Number(b);
        case CellDataType.English:
            return a.localeCompare(b, 'en');
        case CellDataType.Chinese:
            return a.localeCompare(b, 'zh-Hans-CN');
        default:
            return a.localeCompare(b);
    }
}

/**
 * 混合类型文本的升序比较
 *
 * 先按 {@link CellDataType} 分级，类型不同直接按类型优先级比较；
 * 类型相同再按 {@link compareSameType} 的类内规则比较。
 * 升序类型顺序：空值 → 时间 → 数字 → 英文 → 中文 → 其它外语。
 * 需要降序时由调用方对返回值取反。
 * @param {string} a - 文本 a
 * @param {string} b - 文本 b
 * @returns {number} 比较结果（负数 a 在前，正数 b 在前）
 */
function compareMixedText(a: string, b: string): number {
    const ta = classifyCellText(a);
    const tb = classifyCellText(b);
    if (ta !== tb) return ta - tb;
    return compareSameType(ta, a.trim(), b.trim());
}

/**
 * 二维数组排序类（支持 String 类型数据）
 *
 * 统一使用 {@link sort} 方法进行排序，支持三种模式：
 * - 全量模式：不提供 indices/startIndex/endIndex，对全部行/列排序
 * - 连续范围模式：提供 startIndex + endIndex，对 [startIndex, endIndex] 范围排序
 * - 不连续模式：提供 indices 数组，对指定索引位置排序
 *
 * 每种模式均支持列排序（按列值排序行）和行排序（按行值排序列）。
 *
 * 比较规则支持时间、数字、英文字符、中文字符、其它外语字符混合排序：
 * 升序按 时间 → 数字 → 英文 → 中文 → 其它外语 的类型顺序，
 * 降序为完全反转（类型顺序与类内顺序均反转）。
 */
class MatrixSorter {
    /** 数据集合 */
    private data: DataCollection;
    /** 原始数据 */
    private originalValues: (Cell | null)[][] = [];
    /** 排序后的数据 */
    private values: (Cell | null)[][] = [];
    constructor(data: DataCollection) {
        this.data = data;
    }

    /**
     * 初始化数据
     * @param {(Cell | null)[][]} values - 二维数组
     * @returns {MatrixSorter} 返回 this，支持链式调用
     */
    public init(values: (Cell | null)[][]): MatrixSorter {
        this.originalValues = values.map(row => [...row]);
        this.values = values.map(row => [...row]);
        return this;
    }

    /**
     * 重置数据到原始状态
     * @returns {MatrixSorter} 返回 this
     */
    public reset(): MatrixSorter {
        this.values = this.originalValues.map(row => [...row]);
        return this;
    }

    /**
     * 获取当前数据
     * @returns {(Cell | null)[][]} 排序后的二维数组
     */
    public getData(): (Cell | null)[][] {
        return this.values;
    }

    /**
     * 获取数据副本
     * @returns {(Cell | null)[][]} 数据的深拷贝
     */
    public getDataCopy(): (Cell | null)[][] {
        return this.values.map(row => [...row]);
    }

    // ==================== 统一排序方法 ====================

    /**
     * 统一排序入口。
     *
     * 根据 config.axis 决定排序方向，根据 config 中的范围参数决定排序模式：
     * - 提供 indices → 不连续模式（仅对 indices 指定的位置排序）
     * - 提供 startIndex + endIndex → 连续范围模式
     * - 均不提供 → 全量模式（对所有位置排序）
     *
     * 排序前会自动排除筛选按钮（cell.filter 已设置的单元格）所在的行/列，
     * 保证筛选锚点位置不受排序影响：列排序跳过筛选锚点行，行排序跳过筛选锚点列。
     * 若排除后无可排序的行/列（所选范围全部为筛选锚点），中止排序并弹出警告提示。
     *
     * @example
     * // 列排序：按第 2 列对所有行升序
     * sorter.sort({ axis: 'column', keyIndex: 1, ascending: true });
     * // 列排序：按第 2 列对第 0-4 行升序
     * sorter.sort({ axis: 'column', keyIndex: 1, ascending: true, startIndex: 0, endIndex: 4 });
     * // 列排序：按第 2 列对第 0,2,4 行降序
     * sorter.sort({ axis: 'column', keyIndex: 1, ascending: false, indices: [0, 2, 4] });
     * // 行排序：按第 1 行对所有列升序
     * sorter.sort({ axis: 'row', keyIndex: 0, ascending: true });
     * @param {SortConfig} config - 排序配置
     * @returns {MatrixSorter} 返回 this，支持链式调用
     */
    public sort(config: SortConfig): MatrixSorter {
        const { axis, keyIndex, ascending, startIndex, endIndex, indices } = config;

        // 解析目标索引：不连续 > 连续范围 > 全量
        let targetIndices: number[];
        if (indices && indices.length > 0) {
            targetIndices = [...new Set(indices)].sort((a, b) => a - b);
        } else if (startIndex !== undefined && endIndex !== undefined) {
            targetIndices = [];
            for (let i = startIndex; i <= endIndex; i++) {
                targetIndices.push(i);
            }
        } else {
            const total = axis === 'column' ? this.values.length : (this.values[0]?.length ?? 0);
            targetIndices = Array.from({ length: total }, (_, i) => i);
        }

        // 验证
        if (targetIndices.length === 0) return this;
        const validIndices = this._validateIndices(axis, keyIndex, targetIndices);
        // validIndices 为 undefined（数据为空或 keyIndex 越界）或钳制后为空数组时不执行排序
        if (validIndices === undefined || validIndices.length === 0) return this;

        // 排除筛选按钮所在行/列：排序不得移动筛选锚点单元格，否则筛选按钮位置错乱
        const skipSet = this._collectFilterAnchorIndices(axis);
        const sortIndices = skipSet.size > 0 ? validIndices.filter(i => !skipSet.has(i)) : validIndices;
        if (sortIndices.length === 0) {
            // 所选范围全部为筛选锚点行/列：无可用排序位置，明确提示而非静默失败
            const label = ascending ? '升序' : '降序';
            const unit = axis === 'column' ? '行' : '列';
            this.data.messageBox.warning(`执行${label}排列操作时，所选范围均为筛选按钮所在${unit}，无法执行排序`);
            return this;
        }

        // 检查合并单元格一致性：内容单元格的合并格式与主单元格不一致时中止排序并提示
        if (!this._checkMergeConsistency(axis, keyIndex, sortIndices)) {
            const label = ascending ? '升序' : '降序';
            this.data.messageBox.warning(`执行${label}排列操作时，请确保内容所处的单元格的合并格式与主单元格一致`);
            return this;
        }

        // 排序
        if (axis === 'column') {
            this._sortByColumn(keyIndex, ascending, sortIndices);
        } else {
            this._sortByRow(keyIndex, ascending, sortIndices);
        }

        return this;
    }

    // ==================== 核心排序逻辑 ====================

    /**
     * 列排序：按指定列的值对 targetIndices 对应的行进行排序。
     *
     * 提取目标行 → 按列值排序 → 回填到原位。
     * @param {number} colIndex - 排序依据的列索引
     * @param {boolean} ascending - 升序/降序
     * @param {number[]} rowIndices - 要排序的行索引
     * @private
     */
    private _sortByColumn(colIndex: number, ascending: boolean, rowIndices: number[]): void {
        // 提取要排序的行
        const rowsToSort = rowIndices.map(i => this.values[i]);

        // 按列值排序
        rowsToSort.sort((a, b) => {
            return this._compareCells(a[colIndex], b[colIndex], ascending);
        });

        // 回填到原位
        rowIndices.forEach((rowIdx, i) => {
            this.values[rowIdx] = rowsToSort[i];
        });
    }

    /**
     * 行排序：按指定行的值对 targetIndices 对应的列进行排序。
     *
     * 提取关键行的列值 → 排序列索引 → 将相同的列排列方式应用到所有行。
     * @param {number} rowIndex - 排序依据的行索引
     * @param {boolean} ascending - 升序/降序
     * @param {number[]} colIndices - 要排序的列索引
     * @private
     */
    private _sortByRow(rowIndex: number, ascending: boolean, colIndices: number[]): void {
        const keyRow = this.values[rowIndex];

        // 按关键行的值排序列索引，得到排列后的列顺序
        const sortedColIndices = [...colIndices].sort((a, b) => {
            return this._compareCells(keyRow[a], keyRow[b], ascending);
        });

        // 将相同的列排列方式应用到所有行
        for (let r = 0; r < this.values.length; r++) {
            // 保存各目标列在该行的原始值
            const originalSlice = colIndices.map(c => this.values[r][c]);
            // sortedColIndices[j] 是排序后第 j 位的原始列索引，
            // 需要将其映射回 originalSlice 中的位置取出值
            const reordered = sortedColIndices.map(origIdx => {
                const posInOriginal = colIndices.indexOf(origIdx);
                return originalSlice[posInOriginal];
            });
            // 回填：第 j 位放排序后的第 j 个值
            colIndices.forEach((origCol, j) => {
                this.values[r][origCol] = reordered[j];
            });
        }
    }

    // ==================== 辅助方法 ====================

    /**
     * 提取单元格的字符串值
     * @param {Cell | null | undefined} cell - 单元格
     * @returns {string} 字符串值（空单元格返回空字符串）
     * @private
     */
    private _getCellStringValue(cell: Cell | null | undefined): string {
        if (!cell || !cell.chars) return '';
        return cell.chars.map(c => c.char).join('');
    }

    /**
     * 比较两个单元格的值（混合类型分级排序）
     *
     * 委托模块级 {@link compareMixedText} 完成升序比较，降序取反。
     * 升序类型顺序：空值 → 时间 → 数字 → 英文 → 中文 → 其它外语；
     * 降序为完全反转（类型顺序与类内顺序均反转）。
     * @param {Cell | null | undefined} a - 单元格 a
     * @param {Cell | null | undefined} b - 单元格 b
     * @param {boolean} ascending - 升序/降序
     * @returns {number} 比较结果
     * @private
     */
    private _compareCells(a: Cell | null | undefined, b: Cell | null | undefined, ascending: boolean): number {
        const strA = this._getCellStringValue(a);
        const strB = this._getCellStringValue(b);
        const result = compareMixedText(strA, strB);
        return ascending ? result : -result;
    }

    /**
     * 验证并修正排序索引。
     *
     * 对 targetIndices 中超出 [0, max-1] 范围的索引进行过滤（钳制到有效范围内），
     * 保证最大行/列不超过矩阵的实际最大行列数，避免越界访问。
     * keyIndex 若越界则报错提示（排序依据列/行无法钳位，否则会导致按错误列排序）。
     *
     * @param {'column' | 'row'} axis - 排序轴
     * @param {number} keyIndex - 排序依据的列/行索引（0-based）
     * @param {number[]} targetIndices - 目标索引数组（0-based）
     * @returns {number[] | undefined} 过滤后的有效索引数组；数据为空或 keyIndex 越界时返回 undefined
     * @private
     */
    private _validateIndices(axis: 'column' | 'row', keyIndex: number, targetIndices: number[]): number[] | undefined {
        if (this.values.length === 0) {
            this.data.messageBox.error('数据为空');
            return;
        }
        const rowCount = this.values.length;
        const colCount = this.values[0]?.length ?? 0;
        const maxAxis = axis === 'column' ? rowCount : colCount;
        const keyMax = axis === 'column' ? colCount : rowCount;

        // keyIndex 越界无法钳位（否则按错误的列/行排序），直接报错
        if (keyIndex < 0 || keyIndex >= keyMax) {
            const label = axis === 'column' ? `列索引 ${this.data.getColName(keyIndex + 1)}` : `行索引 ${keyIndex + 1}`;
            this.data.messageBox.error(`${label} 无效`);
            return;
        }

        // 钳制 targetIndices：过滤掉 < 0 或 >= maxAxis 的索引，去重并保持升序
        const validIndices = [...new Set(
            targetIndices.filter(i => i >= 0 && i < maxAxis)
        )].sort((a, b) => a - b);

        return validIndices;
    }

    /**
     * 收集设置了筛选按钮（cell.filter !== undefined）的单元格所在的行/列索引
     *
     * 列排序时返回这些单元格所在行（0-based），行排序时返回所在列（0-based）。
     * 排序时跳过这些索引，保证筛选按钮位置不被排序移动。
     *
     * 注意：getCellColAndRow 返回 1-based 行列号，需减 1 转为矩阵的 0-based 索引。
     * @param {'column' | 'row'} axis - 排序轴
     * @returns {Set<number>} 需跳过的索引集合（0-based）
     * @private
     */
    private _collectFilterAnchorIndices(axis: 'column' | 'row'): Set<number> {
        const result = new Set<number>();
        for (const cell of this.data.values) {
            if (cell.filter === undefined) continue;
            const { col, row } = this.data.getCellColAndRow(cell.cell);
            result.add(axis === 'column' ? row - 1 : col - 1);
        }
        return result;
    }

    /**
     * 检查参与排序的行/列中，内容单元格的合并格式是否与主单元格一致。
     *
     * 主单元格定义：
     * - 列排序：待排序行在排序依据列（keyIndex）上的单元格；
     * - 行排序：待排序列在排序依据行（keyIndex）上的单元格。
     *
     * 合并格式为 (colspan, rowspan) 二元组（未设置视为 1）。主单元格为空时视为 (1,1)。
     * 任一非空内容单元格与主单元格的合并格式不一致即返回 false，由调用方中止排序并提示。
     * @param {'column' | 'row'} axis - 排序轴
     * @param {number} keyIndex - 排序依据的列/行索引（0-based）
     * @param {number[]} sortIndices - 参与排序的行/列索引（0-based）
     * @returns {boolean} true=格式一致（可排序），false=存在不一致
     * @private
     */
    private _checkMergeConsistency(axis: 'column' | 'row', keyIndex: number, sortIndices: number[]): boolean {
        const getSpan = (cell: Cell | null | undefined): { colspan: number; rowspan: number } => ({
            colspan: cell?.colspan ?? 1,
            rowspan: cell?.rowspan ?? 1,
        });

        for (const i of sortIndices) {
            if (axis === 'column') {
                // 列排序：逐行检查，主单元格为该行在排序依据列上的单元格
                const row = this.values[i];
                if (!row) continue;
                const mainSpan = getSpan(row[keyIndex]);
                for (let c = 0; c < row.length; c++) {
                    if (c === keyIndex) continue;
                    const cell = row[c];
                    if (!cell) continue;
                    const span = getSpan(cell);
                    if (span.colspan !== mainSpan.colspan || span.rowspan !== mainSpan.rowspan) {
                        return false;
                    }
                }
            } else {
                // 行排序：逐列检查，主单元格为该列在排序依据行上的单元格
                const mainSpan = getSpan(this.values[keyIndex]?.[i]);
                for (let r = 0; r < this.values.length; r++) {
                    if (r === keyIndex) continue;
                    const cell = this.values[r]?.[i];
                    if (!cell) continue;
                    const span = getSpan(cell);
                    if (span.colspan !== mainSpan.colspan || span.rowspan !== mainSpan.rowspan) {
                        return false;
                    }
                }
            }
        }
        return true;
    }
}

export { MatrixSorter, type SortConfig, CellDataType, compareMixedText, classifyCellText };
