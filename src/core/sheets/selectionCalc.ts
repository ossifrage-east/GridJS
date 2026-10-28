/*  # 选区即时计算模块 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { DataCollection } from "../dataArchitecture/dataCollection";
import { Cell } from "../dataArchitecture/cell";
import { Char } from "../dataArchitecture/char";
import { DataEvents } from "../constant";
import { debounce } from "../../utils/debounce";
import { createDiv, createSpan } from "../../utils/dom";
import { formatDateTimeValue, formatNumberValue, isDateTimeCode, parseDateValue } from "../../utils/numberFormat";

/**
 * 支持参与即时计算的单元格类型
 *
 * 由单元格的数字格式码归类：数值（含常规数字）、货币（含会计专用）、
 * 科学记数、日期、时间。百分比、文本等不在支持范围内。
 */
export type SelectionCalcType = 'number' | 'currency' | 'scientific' | 'date' | 'time';

/** 单元格类型判定结果：支持类型之外统一归为 other（文本 / 百分比等） */
type CellValueType = SelectionCalcType | 'other';

/**
 * 一条即时计算展示项（如「求和 3.3」）
 *
 * 底部栏按「项名 + 值」两部分渲染：项名只作标识，值可点击复制。
 */
export interface SelectionCalcItem {
    /** 展示名称（求和 / 平均 / 计数 / 最大 / 最小） */
    label: string;
    /** 已格式化好的值文本 */
    value: string;
}

/**
 * 选区即时计算结果
 *
 * 为 null 时不显示任何内容（没有选区）。
 */
export interface SelectionCalcResult {
    /** 选区内有内容的单元格个数（空白格不计） */
    count: number;
    /** 是否全部为同一支持类型（为 false 时只显示计数） */
    comparable: boolean;
    /** 同类时的单元格类型 */
    type: SelectionCalcType | null;
    /** 展示项：同类且可计算时为五项，否则仅「计数」一项 */
    items: SelectionCalcItem[];
}

/** 选区即时计算初始化选项 */
export interface SelectionCalcOptions {
    /** 数据实例 */
    data: DataCollection;
    /** 底部栏挂载容器（bottomTools 的 split-pre） */
    parentElement: HTMLElement;
}

/**
 * 选区即时计算
 *
 * 监听选区 / 内容变化，对当前选区做一次矩形扫描并聚合出求和、平均、计数、
 * 最大、最小，把即时值渲染到底部工具条的 split-pre 内部（DOM，见 {@link bar}）。
 *
 * ## 显示前提
 * 只要存在选区就显示（与活动格是否有文字无关）。选区没有任何有内容的单元格时
 * 显示「计数 0」。
 *
 * ## 点击行为
 * - 点击即时值本身：只复制值（不含「求和 / 平均…」字样），并同步内部剪贴板，
 *   使该值可以直接用 Ctrl+V 粘贴进表格；
 * - 点击工具栏「计算」菜单的某一项：把该项结果写入**空白的**活动单元格
 *   （chars 为空时才写），单元格内即是纯结果，不带项名。
 *
 * ## 类型一致性
 * 选中区域内的有内容单元格必须全部属于同一种受支持类型（数值 / 货币 /
 * 会计专用 / 科学记数 / 日期 / 时间），才显示五项结果；一旦混合或含不支持
 * 类型（如百分比、文本），只显示计数。空白单元格不参与判定也不计数。
 *
 * ## 性能
 * 每次只扫描选区矩形（不遍历整张表），并对选区 / 内容变化做 80ms 防抖，
 * 避免拖选、Ctrl+A、连续输入时反复做全选区计算。
 *
 * @class SelectionCalc
 * @example
 * const calc = new SelectionCalc({
 *     data,
 *     parentElement: bottomToolsSplit.getElements().splitPrev,
 * });
 */
export class SelectionCalc {
    /** 支持参与计算的类型集合 */
    private static readonly SUPPORTED: ReadonlySet<CellValueType> = new Set<CellValueType>([
        'number', 'currency', 'scientific', 'date', 'time',
    ]);
    /** 纯数字文本（与数字格式化的解析口径保持一致） */
    private static readonly NUMERIC_TEXT = /^[-+]?\d+(\.\d+)?([eE][-+]?\d+)?$/;
    /** 货币 / 会计专用格式码：货币符号，或会计专用的 _( 与 * 占位符 */
    private static readonly CURRENCY_CODE = /[¥$€£￥]|_\(|\*\s/;
    /** 科学记数格式码（如 0.00E+00） */
    private static readonly SCIENTIFIC_CODE = /E[+-]?0+/i;

    /** 数据实例 */
    private readonly data: DataCollection;
    /** 底部栏 DOM 容器 */
    private readonly bar: HTMLDivElement;
    /** 防抖后的计算入口 */
    private readonly schedule: () => void;
    /** 最近一次计算结果 */
    private result: SelectionCalcResult | null = null;
    /** 结果签名：与上次相同时跳过 DOM 更新 */
    private signature = '';

    /**
     * 构造函数
     * @param {SelectionCalcOptions} options - 初始化选项
     */
    constructor(options: SelectionCalcOptions) {
        this.data = options.data;
        this.bar = createDiv({ className: 'selection-calc' });
        options.parentElement.appendChild(this.bar);
        this.schedule = debounce(() => this.compute(), 80);

        // 选区 / 活动格变化、内容变化、撤销重做、结构变化都会影响结果
        this.data.on(DataEvents.SELECTION_CHANGED, this.schedule);
        this.data.on(DataEvents.ACTIVED_CELL_CHANGED, this.schedule);
        this.data.on(DataEvents.VALUES_CHANGED, this.schedule);
        this.data.on(DataEvents.SNAPSHOT_RESTORED, this.schedule);
        this.data.on(DataEvents.ROWS_CHANGED, this.schedule);
        this.data.on(DataEvents.COLUMNS_CHANGED, this.schedule);
    }

    /**
     * 立即排期一次重算（初始化时调用，避免首个事件到来前底部栏为空）
     */
    public refresh(): void {
        this.schedule();
    }

    /**
     * 把某项统计结果写入活动单元格（工具栏「计算」菜单的点选结果）
     *
     * 前提是活动单元格没有文字内容（chars 为空）：此时把该项的结果值（不带
     * 「求和 / 平均…」字样）作为单元格内容写入，走标准写回路径（可撤销、
     * 会持久化）。活动格已有文字时什么都不做，不覆盖原有内容。
     * @param {string} label - 项名（求和 / 平均 / 计数 / 最大 / 最小）
     */
    public applyToActivedCell(label: string): void {
        // 立即重算：防抖可能还没跑（例如刚改完选区就点菜单）
        this.compute();
        const result = this.result;
        const activeName = this.data.activedCell;
        if (!result || !activeName) return;

        // 合并区内的非锚点位置会重定向到锚点，空判定要按真正的落点做
        const targetName = this.data.resolveWriteTargetName(activeName);
        const target = this.data.getCellByName(targetName);
        if (target?.chars && target.chars.length > 0) return;

        // 类型混合时只有「计数」一项，点其它项无可写内容
        const item = result.items.find(entry => entry.label === label);
        if (!item) return;
        this.data.applyScriptWrites('计算', { [targetName]: { value: item.value } });
    }

    /**
     * 扫描选区并聚合结果
     *
     * 只要有选区就给出结果（活动格是否有文字不影响底部即时值）；选区按行列
     * 遍历（列名只求一次），空白格跳过；全为空白时显示「计数 0」。
     * @private
     */
    private compute(): void {
        const { data } = this;
        const range = data.selection;
        if (!range) {
            this.update(null);
            return;
        }
        const [startName, endName] = range.split(':');
        const start = data.getCellColAndRow(startName || endName);
        const end = data.getCellColAndRow(endName || startName);

        const startCol = Math.min(start.col, end.col);
        const endCol = Math.max(start.col, end.col);
        const startRow = Math.min(start.row, end.row);
        const endRow = Math.max(start.row, end.row);

        let count = 0;
        let numericCount = 0;
        let sum = 0;
        let max = -Infinity;
        let min = Infinity;
        let type: CellValueType | null = null;
        let mixed = false;
        let numberFormat: string | undefined;

        for (let col = startCol; col <= endCol; col++) {
            const colName = data.getColName(col);
            for (let row = startRow; row <= endRow; row++) {
                const cell = data.getCellByName(`${colName}${row}`);
                if (!cell) continue;
                const text = cell.getText();
                if (!text) continue;

                const cellType = SelectionCalc.detectType(text, cell.numberFormat);
                if (type === null) {
                    type = cellType;
                    numberFormat = cell.numberFormat;
                } else if (type !== cellType) {
                    mixed = true;
                }
                count++;

                if (!SelectionCalc.SUPPORTED.has(cellType)) continue;
                const value = SelectionCalc.toNumber(cellType, text);
                if (value === null) continue;
                numericCount++;
                sum += value;
                if (value > max) max = value;
                if (value < min) min = value;
            }
        }

        // 选区里没有任何有内容单元格：仍然显示，只是计数为 0
        if (count === 0 || type === null) {
            this.update({
                count: 0,
                comparable: false,
                type: null,
                items: [{ label: '计数', value: '0' }],
            });
            return;
        }

        // 只有全部同类且属于受支持类型（非 other），且至少解析出一个数值时才计算五项
        const resultType: SelectionCalcType | null =
            (!mixed && numericCount > 0 && type !== 'other') ? type : null;
        const items: SelectionCalcItem[] = [];
        if (resultType !== null) {
            const format = (value: number): string =>
                SelectionCalc.formatValue(value, resultType, numberFormat);
            items.push({ label: '求和', value: format(sum) });
            items.push({ label: '平均', value: format(sum / numericCount) });
            items.push({ label: '计数', value: String(count) });
            items.push({ label: '最大', value: format(max) });
            items.push({ label: '最小', value: format(min) });
        } else {
            items.push({ label: '计数', value: String(count) });
        }

        this.update({
            count,
            comparable: resultType !== null,
            type: resultType,
            items,
        });
    }

    /**
     * 写入结果并同步底部栏（结果未变化时跳过，避免多余 DOM 更新）
     * @param {SelectionCalcResult | null} result - 新结果
     * @private
     */
    private update(result: SelectionCalcResult | null): void {
        const signature = SelectionCalc.signatureOf(result);
        if (signature === this.signature) return;
        this.signature = signature;
        this.result = result;

        this.bar.textContent = '';
        if (result) {
            for (const item of result.items) {
                // 项名只是标识（不可点、不参与复制）；值单独一段：悬停才变色，点击只复制值本身
                this.bar.appendChild(createSpan({
                    className: 'selection-calc__item',
                    children: [
                        createSpan({
                            className: 'selection-calc__label',
                            textContent: item.label,
                        }),
                        createSpan({
                            className: 'selection-calc__value',
                            textContent: item.value,
                            attributes: { title: `点击复制：${item.value}` },
                            events: { click: () => this.copyValue(item.value) },
                        }),
                    ],
                }));
            }
        }
    }

    /**
     * 点击即时值：复制值本身，并让表格内部粘贴可用
     *
     * 系统剪贴板供外部程序使用；内部剪贴板同时更新成「1×1 单元格」，
     * 这样在表格里按 Ctrl+V（或菜单粘贴）就能把这个值落到当前活动格。
     * @param {string} text - 要复制的值（不含项名）
     * @private
     */
    private copyValue(text: string): void {
        const name = this.data.activedCell;
        if (name) {
            const cell = new Cell();
            cell.cell = name;
            cell.chars = Array.from(text, (ch: string) => {
                const _char = new Char();
                _char.char = ch;
                return _char;
            });
            this.data.clipboardManager.setInternalClipboard({
                type: 'cell',
                cells: [cell],
                sourceRange: `${name}:${name}`,
                visibleColOffsets: [0],
                visibleRowOffsets: [0],
            });
        }
        SelectionCalc.copyText(text);
    }

    /**
     * 复制文本到剪贴板
     *
     * 优先用异步剪贴板 API（需要安全上下文与焦点）；不可用时退回临时 textarea +
     * `execCommand('copy')` 的旧路径，保证在非 https 的本地环境下同样可用。
     * @param {string} text - 要复制的文本
     * @private
     */
    private static copyText(text: string): void {
        if (navigator.clipboard && window.isSecureContext) {
            navigator.clipboard.writeText(text).catch(() => SelectionCalc.copyTextFallback(text));
            return;
        }
        SelectionCalc.copyTextFallback(text);
    }

    /**
     * 剪贴板写入的兜底实现（临时 textarea + execCommand）
     * @param {string} text - 要复制的文本
     * @private
     */
    private static copyTextFallback(text: string): void {
        // 临时 textarea 会抢走焦点，复制完成后交还给原来的元素（表格画布 / 编辑区）
        const previous = document.activeElement as HTMLElement | null;
        const area = document.createElement('textarea');
        area.value = text;
        area.setAttribute('readonly', '');
        area.style.position = 'fixed';
        area.style.top = '-1000px';
        area.style.opacity = '0';
        document.body.appendChild(area);
        area.select();
        try {
            document.execCommand('copy');
        } catch {
            // 复制失败时静默：不打断用户当前操作
        }
        document.body.removeChild(area);
        previous?.focus?.();
    }

    /**
     * 结果签名：只包含展示项，用于判断「需要更新界面吗」
     * @param {SelectionCalcResult | null} result - 结果
     * @returns {string} 签名
     * @private
     */
    private static signatureOf(result: SelectionCalcResult | null): string {
        if (!result) return '';
        return result.items.map(item => `${item.label}=${item.value}`).join('\u0001');
    }

    /**
     * 按数字格式码与文本内容判定单元格类型
     *
     * 有格式码时以格式码为准（日期/时间 → 进一步区分；百分比不在支持列表）；
     * 常规 / 无格式时按文本是否为纯数字判定为「数值」，否则归为 other。
     * @param {string} text - 单元格文本
     * @param {string} [format] - 数字格式码
     * @returns {CellValueType} 判定结果
     * @private
     */
    private static detectType(text: string, format?: string): CellValueType {
        if (format && format !== 'General' && format !== '@') {
            if (isDateTimeCode(format)) {
                return (/[hs]/i.test(format) || /AM\/PM/i.test(format)) ? 'time' : 'date';
            }
            if (format.includes('%')) return 'other';
            if (SelectionCalc.SCIENTIFIC_CODE.test(format)) return 'scientific';
            if (SelectionCalc.CURRENCY_CODE.test(format)) return 'currency';
            if (/[0#]/.test(format)) return 'number';
            return 'other';
        }
        return SelectionCalc.NUMERIC_TEXT.test(text.trim()) ? 'number' : 'other';
    }

    /**
     * 把单元格文本转换成参与聚合的数值
     *
     * 日期 / 时间按其底层 Excel 序列号参与运算（与 Excel 一致）；
     * 数值 / 货币 / 科学记数按纯数字解析（容错千分位与货币符号）。
     * @param {CellValueType} type - 单元格类型
     * @param {string} text - 单元格文本
     * @returns {number | null} 数值；无法解析时为 null
     * @private
     */
    private static toNumber(type: CellValueType, text: string): number | null {
        if (type === 'date' || type === 'time') {
            const date = parseDateValue(text);
            if (!date) return null;
            const ms = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate(),
                date.getHours(), date.getMinutes(), date.getSeconds(), date.getMilliseconds());
            return (ms - Date.UTC(1899, 11, 30)) / 86400000;
        }
        const normalized = text.trim().replace(/[,\uFF0C\s]/g, '').replace(/[¥$€£￥]/g, '');
        if (!SelectionCalc.NUMERIC_TEXT.test(normalized)) return null;
        const value = parseFloat(normalized);
        return Number.isFinite(value) ? value : null;
    }

    /**
     * 按单元格类型与格式码格式化计算结果
     *
     * 日期 / 时间结果按同类型单元格的格式码回显（如 2026/10/10）；
     * 其余类型有格式码时走数字格式化（保留货币符号 / 小数位 / 千分位），
     * 常规格式则输出去掉浮点误差的普通数字。
     * @param {number} value - 结果数值
     * @param {string} type - 单元格类型
     * @param {string} [format] - 数字格式码
     * @returns {string} 展示文本
     * @private
     */
    private static formatValue(value: number, type: string, format?: string): string {
        if (type === 'date' || type === 'time') {
            const date = parseDateValue(String(value));
            if (date) {
                return formatDateTimeValue(date, format || (type === 'time' ? 'h:mm:ss' : 'yyyy/m/d'));
            }
            return SelectionCalc.formatPlain(value);
        }
        if (format && format !== 'General' && format !== '@') {
            return formatNumberValue(value, format);
        }
        return SelectionCalc.formatPlain(value);
    }

    /**
     * 常规格式数字输出：收敛浮点误差（如 0.1 + 0.2 → 3.3 而不是 3.3000000000000003）
     * @param {number} value - 数值
     * @returns {string} 文本
     * @private
     */
    private static formatPlain(value: number): string {
        if (!Number.isFinite(value)) return String(value);
        const rounded = Number(value.toFixed(10));
        return String(Number.isFinite(rounded) ? rounded : value);
    }
}
