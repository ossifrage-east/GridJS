import { Char } from "./char";
import { DataEvents } from "../constant";
import { EventEmitter } from "../../utils/eventEmitter";

/**
 * 文本对齐方式类型
 */
export type TextAlign = 'left' | 'center' | 'right' | 'justify';

/**
 * 垂直对齐方式类型
 */
export type VerticalAlign = 'top' | 'middle' | 'bottom';


/**
 * Cell 类
 * 表示电子表格的单元格，包含样式和值属性
 * 提供属性变更监听和事件通知机制
 * 
 * @class Cell
 * @extends Char
 * @example
 * const cell = new Cell();
 * cell.rowspan = 2;
 * cell.borderTopWidth = 1;
 * cell.borderBottomWidth = 1;
 * cell.borderLeftWidth = 1;
 * cell.borderRightWidth = 1;
 * cell.textAlign = 'center';
 * cell.alignItems = 'middle';
 * cell.backgroundColor = '#f0f0f0';
 * cell.filter = 32;
 * cell.wrap = true;
 * cell.letterSpacing = 1;
 * cell.lineSpacing = 1;
 * cell.colspan = 2;
 * cell.chars = [new Char()];
 * cell.on('cell:rowspan', ({ property, oldValue, newValue }) => console.log(`${property} changed`));
 */
export class Cell extends EventEmitter {
    /** 单元格名称 */
    private _cell!: string;
    /** 是否可以编辑 */ 
    private _isEdit: boolean = true;
    /** 跨列数 */
    private _colspan?: number;
    /** 跨行数 */
    private _rowspan?: number;
    // /** 字体族 */
    // private _fontFamily?: string;
    // /** 字体大小 */
    // private _fontSize?: number;
    // /** 字体粗细 */
    // private _fontWeight?: boolean;
    // /** 字体样式 */
    // private _fontStyle?: boolean;
    // /** 字体颜色 */
    // private _fontColor?: string;
    // /** 下划线 */
    // private _underline?: boolean;
    // /** 删除线 */
    // private _strikethrough?: boolean;
    /** 上边框宽度 */
    private _borderTopWidth?: number;
    /** 下边框宽度 */
    private _borderBottomWidth?: number;
    /** 左边框宽度 */
    private _borderLeftWidth?: number;
    /** 右边框宽度 */
    private _borderRightWidth?: number;
    /** 文本对齐方式 */
    private _textAlign?: TextAlign;
    /** 垂直对齐方式 */
    private _alignItems?: VerticalAlign;
    /** 背景颜色 */
    private _backgroundColor?: string;
    /** 数字格式码（如 'General'/'0.00'/'#,##0.00'/'0.00%'/'¥#,##0.00'/'0.00E+00'/'@'） */
    private _numberFormat?: string;
    /** 边框颜色，默认 '#000000' */
    private _borderColor?: string;
    /** 单元格值 */
    private _chars?: Char[] = new Array<Char>();
    /** 过滤 */
    private _filter?: number | undefined;
    /** 换行 */
    private _wrap?: boolean;
    /** 字母间距 */
    private _letterSpacing?: number;
    /** 行高 */
    private _lineSpacing?: number;

    /**
     * 构造函数
     */
    constructor() {
        super();
    }

    /**
     * 获取单元格名称
     */
    get cell(): string {
        return this._cell;
    }

    /**
     * 设置单元格名称
     * @param {string} value - 单元格名称
     */
    set cell(value: string) {
        this.updateProperty(DataEvents.CELL_CHANGED, '_cell', value);
    }
    
    /**
     * 获取是否正在编辑
     * @returns {boolean} 是否正在编辑
     */
    get isEdit(): boolean {
        return this._isEdit;
    }

    /**
     * 设置是否正在编辑
     * @param {boolean} value - 是否正在编辑
     */
    set isEdit(value: boolean) {
        this.updateProperty(DataEvents.IS_EDIT_CHANGED, '_isEdit', value);
    }
    
    /**
     * 获取跨列数
     * @returns {number | undefined} 跨列数
     */
    get colspan(): number | undefined {
        return this._colspan;
    }
    
    /**
     * 设置跨列数
     * @param {number} colspan - 跨列数
     */
    set colspan(colspan: number) {
        this.updateProperty(DataEvents.COLSPAN_CHANGED, '_colspan', colspan);
    }
    
    /**
     * 获取跨行数
     * @returns {number | undefined} 跨行数
     */
    get rowspan(): number | undefined {
        return this._rowspan;
    }
    
    /**
     * 设置跨行数
     * @param {number} rowspan - 跨行数
     */
    set rowspan(rowspan: number) {
        this.updateProperty(DataEvents.ROWSPAN_CHANGED, '_rowspan', rowspan);
    }

    /**
     * 获取字体族
     * @returns {string | undefined} 字体族
     */
    get fontFamily(): string | undefined {
        // 如果单元格有字符,则返回第一个字符的字体族
        if (this.chars && this.chars.length > 0) {
            return this.chars[0].fontFamily;
        }
        return undefined;
    }

    /**
     * 设置字体族
     * @param {string | undefined} value - 字体族
     */
    set fontFamily(value: string | undefined) {
        // 如果单元格有字符,则将字体族设置为所有字符的字体族
        if (this.chars && this.chars.length > 0) {
            this.chars.forEach(char => char.fontFamily = value);
        }
    }

    /**
     * 获取字体大小
     * @returns {number | undefined} 字体大小
     */
    get fontSize(): number | undefined {
        // 如果单元格有字符,则返回第一个字符的字体大小
        if (this.chars && this.chars.length > 0) {
            return this.chars[0].fontSize;
        }
        return undefined;
    }

    /**
     * 设置字体大小
     * @param {number | undefined} value - 字体大小
     */
    set fontSize(value: number | undefined) {
        // 如果单元格有字符,则将字体大小设置为所有字符的字体大小
        if (this.chars && this.chars.length > 0) {
            this.chars.forEach(char => char.fontSize = value);
        }
    }

    /**
     * 获取字体粗细
     * @returns {boolean | undefined} 是否加粗
     */
    get fontWeight(): boolean | undefined {
        // 如果单元格有字符,则返回第一个字符的字体粗细
        if (this.chars && this.chars.length > 0) {
            return this.chars[0].fontWeight;
        }
        return undefined;
    }

    /**
     * 设置字体粗细
     * @param {boolean | undefined} value - 是否加粗
     */
    set fontWeight(value: boolean | undefined) {
        // 如果单元格有字符,则将字体粗细设置为所有字符的字体粗细
        if (this.chars && this.chars.length > 0) {
            this.chars.forEach(char => char.fontWeight = value);
        }
    }

    /**
     * 获取字体样式
     * @returns {boolean | undefined} 是否斜体
     */
    get fontStyle(): boolean | undefined {
        // 如果单元格有字符,则返回第一个字符的字体样式
        if (this.chars && this.chars.length > 0) {
            return this.chars[0].fontStyle;
        }
        return undefined;
    }

    /**
     * 设置字体样式
     * @param {boolean | undefined} value - 是否斜体
     */
    set fontStyle(value: boolean | undefined) {
        // 如果单元格有字符,则将字体样式设置为所有字符的字体样式
        if (this.chars && this.chars.length > 0) {
            this.chars.forEach(char => char.fontStyle = value);
        }
    }

    /**
     * 获取字体颜色
     * @returns {string | undefined} 字体颜色
     */
    get fontColor(): string | undefined {
        // 如果单元格有字符,则返回第一个字符的字体颜色
        if (this.chars && this.chars.length > 0) {
            return this.chars[0].fontColor;
        }
        return undefined;
    }

    /**
     * 设置字体颜色
     * @param {string | undefined} value - 字体颜色
     */
    set fontColor(value: string | undefined) {
        // 如果单元格有字符,则将字体颜色设置为所有字符的字体颜色
        if (this.chars && this.chars.length > 0) {
            this.chars.forEach(char => char.fontColor = value);
        }
    }
    
    /**
     * 获取下划线状态
     * @returns {boolean | undefined} 是否有下划线
     */
    get underline(): boolean | undefined {
        // 如果单元格有字符,则返回第一个字符的下划线状态
        if (this.chars && this.chars.length > 0) {
            return this.chars[0].underline;
        }
        return undefined;
    }

    /**
     * 设置下划线状态
     * @param {boolean | undefined} value - 是否有下划线
     */
    set underline(value: boolean | undefined) {
        // 如果单元格有字符,则将下划线状态设置为所有字符的下划线状态
        if (this.chars && this.chars.length > 0) {
            this.chars.forEach(char => char.underline = value);
        }
    }
    
    /**
     * 获取删除线状态
     * @returns {boolean | undefined} 是否有删除线
     */
    get strikethrough(): boolean | undefined {
        // 如果单元格有字符,则返回第一个字符的删除线状态
        if (this.chars && this.chars.length > 0) {
            return this.chars[0].strikethrough;
        }
        return undefined;
    }

    /**
     * 设置删除线状态
     * @param {boolean | undefined} value - 是否有删除线
     */
    set strikethrough(value: boolean | undefined) {
        // 如果单元格有字符,则将删除线状态设置为所有字符的删除线状态
        if (this.chars && this.chars.length > 0) {
            this.chars.forEach(char => char.strikethrough = value);
        }
    }

    /**
     * 获取上边框宽度
     * @returns {number | undefined} 上边框宽度
     */
    get borderTopWidth(): number | undefined {
        return this._borderTopWidth;
    }
    /**
     * 设置上边框宽度
     * @param {number} borderTopWidth - 上边框宽度
     */
    set borderTopWidth(borderTopWidth: number) {
        this.updateProperty(DataEvents.BORDER_TOP_WIDTH_CHANGED, '_borderTopWidth', borderTopWidth);
    }
    
    /**
     * 获取下边框宽度
     * @returns {number | undefined} 下边框宽度
     */
    get borderBottomWidth(): number | undefined {
        return this._borderBottomWidth;
    }
    
    /**
     * 设置下边框宽度
     * @param {number} borderBottomWidth - 下边框宽度
     */
    set borderBottomWidth(borderBottomWidth: number) {
        this.updateProperty(DataEvents.BORDER_BOTTOM_WIDTH_CHANGED, '_borderBottomWidth', borderBottomWidth);
    }
    
    /**
     * 获取左边框宽度
     * @returns {number | undefined} 左边框宽度
     */
    get borderLeftWidth(): number | undefined {
        return this._borderLeftWidth;
    }
    
    /**
     * 设置左边框宽度
     * @param {number} borderLeftWidth - 左边框宽度
     */
    set borderLeftWidth(borderLeftWidth: number) {
        this.updateProperty(DataEvents.BORDER_LEFT_WIDTH_CHANGED, '_borderLeftWidth', borderLeftWidth);
    }
    
    /**
     * 获取右边框宽度
     * @returns {number | undefined} 右边框宽度
     */
    get borderRightWidth(): number | undefined {
        return this._borderRightWidth;
    }
    
    /**
     * 设置右边框宽度
     * @param {number} borderRightWidth - 右边框宽度
     */
    set borderRightWidth(borderRightWidth: number) {
        this.updateProperty(DataEvents.BORDER_RIGHT_WIDTH_CHANGED, '_borderRightWidth', borderRightWidth);
    }
    
    /**
     * 获取文本对齐方式
     * @returns {TextAlign | undefined} 文本对齐方式
     */
    get textAlign(): TextAlign | undefined {
        return this._textAlign;
    }
    
    /**
     * 设置文本对齐方式
     * @param {TextAlign} textAlign - 文本对齐方式
     */
    set textAlign(textAlign: TextAlign) {
        this.updateProperty(DataEvents.TEXT_ALIGN_CHANGED, '_textAlign', textAlign);
    }
    
    /**
     * 获取垂直对齐方式
     * @returns {VerticalAlign | undefined} 垂直对齐方式
     */
    get alignItems(): VerticalAlign | undefined {
        return this._alignItems;
    }
    
    /**
     * 设置垂直对齐方式
     * @param {VerticalAlign} alignItems - 垂直对齐方式
     */
    set alignItems(alignItems: VerticalAlign) {
        this.updateProperty(DataEvents.ALIGN_ITEMS_CHANGED, '_alignItems', alignItems);
    }
    
    /**
     * 获取背景颜色
     * @returns {string | undefined} 背景颜色
     */
    get backgroundColor(): string | undefined {
        return this._backgroundColor;
    }
    
    /**
     * 设置背景颜色
     * @param {string} backgroundColor - 背景颜色
     */
    set backgroundColor(backgroundColor: string) {
        this.updateProperty(DataEvents.BACKGROUND_COLOR_CHANGED, '_backgroundColor', backgroundColor);
    }

    /**
     * 获取数字格式码
     * @returns {string | undefined} 数字格式码
     */
    get numberFormat(): string | undefined {
        return this._numberFormat;
    }

    /**
     * 设置数字格式码
     * @param {string} numberFormat - 数字格式码
     */
    set numberFormat(numberFormat: string) {
        this.updateProperty(DataEvents.NUMBER_FORMAT_CHANGED, '_numberFormat', numberFormat);
    }

    /**
     * 获取边框颜色
     * @returns {string | undefined} 边框颜色
     */
    get borderColor(): string | undefined {
        return this._borderColor;
    }

    /**
     * 设置边框颜色
     * @param {string} borderColor - 边框颜色
     */
    set borderColor(borderColor: string) {
        this.updateProperty(DataEvents.BORDER_COLOR_CHANGED, '_borderColor', borderColor);
    }
    
    /**
     * 获取单元格值内容
     * @returns {Char[] | undefined} 单元格值数组
     */
    get chars(): Char[] | undefined {
        return this._chars;
    }    
    
    /**
     * 设置单元格值内容
     * @param {Char[]} chars - 单元格值数组
     */
    set chars(chars: Char[]) {
        this.updateProperty(DataEvents.VALUE_CHANGED, '_chars', chars);
    }
    /**
     * 获取是否启用筛选
     * @returns {number | undefined} 是否启用筛选
     */
    get filter(): number | undefined {
        return this._filter;
    }
    
    /**
     * 设置是否启用筛选
     * @param {number} filter - 是否启用筛选
     */
    set filter(filter: number) {
        this.updateProperty(DataEvents.FILTER_CHANGED, '_filter', filter);
    }
    
    /**
     * 获取是否启用换行
     * @returns {boolean | undefined} 是否启用换行
     */
    get wrap(): boolean | undefined {
        return this._wrap;
    }
    
    /**
     * 设置是否启用换行
     * @param {boolean} wrap - 是否启用换行
     */
    set wrap(wrap: boolean) {
        this.updateProperty(DataEvents.WRAP_CHANGED, '_wrap', wrap);
    }
    
    /**
     * 获取字间距
     * @returns {number | undefined} 字间距
     */
    get letterSpacing(): number | undefined {
        return this._letterSpacing;
    }
    
    /**
     * 设置字间距
     * @param {number} letterSpacing - 字间距
     */
    set letterSpacing(letterSpacing: number) {
        this.updateProperty(DataEvents.LETTER_SPACING_CHANGED, '_letterSpacing', letterSpacing);
    }
    
    /**
     * 获取行间距
     * @returns {number | undefined} 行间距
     */
    get lineSpacing(): number | undefined {
        return this._lineSpacing;
    }
    
    /**
     * 设置行间距
     * @param {number} lineSpacing - 行间距
     */
    set lineSpacing(lineSpacing: number) {
        this.updateProperty(DataEvents.LINE_SPACING_CHANGED, '_lineSpacing', lineSpacing);
    }

    /**
     * 添加单元格值，不带样式
     * @param {string} value - 单元格值字符串
     */
    public addValue(value: string) {
        const chars = Array.from(value); // 将字符串转换为字符数组
        const _chars: Char[] = [];
        for (const char of chars) {
            const _char = new Char();
            _char.char = char;
            _chars.push(_char);
        }
        this.chars = _chars;
    }
    
    /**
     * 清空单元格值
     */
    public clearValue() {
        this.chars = [];
    }

    /**
     * 重置单元格的所有内容与格式（保留 cell 名称和功能性字段 colspan/rowspan/filter）
     *
     * 与 clearValue（只清 chars）不同，reset 会同时清除：
     * - 字符级样式（每个 Char 的 fontFamily/fontSize/fontWeight/fontStyle/fontColor/underline/strikethrough）
     * - 单元格级样式（textAlign/alignItems/wrap/letterSpacing/lineSpacing）
     * - 边框（四边 borderWidth + 边框颜色 borderColor）
     * - 背景色（backgroundColor）
     *
     * 功能性字段（cell/colspan/rowspan/filter/isEdit）保留，因为它们承载合并结构、筛选标记等非样式语义。
     * 数字格式码（numberFormat）同样保留——它属于格式配置而非即用样式，reset 只清内容与即用样式。
     */
    public reset(): void {
        this._chars = [];
        this._textAlign = undefined;
        this._alignItems = undefined;
        this._wrap = undefined;
        this._letterSpacing = undefined;
        this._lineSpacing = undefined;
        this._borderTopWidth = undefined;
        this._borderBottomWidth = undefined;
        this._borderLeftWidth = undefined;
        this._borderRightWidth = undefined;
        this._borderColor = undefined;
        this._backgroundColor = undefined;
    }

    /**
     * 判断是否为「空单元格」：不包含任何样式与内容信息
     * 「样式与内容信息」的定义（任一存在即非空）：
     * - 文字信息：chars 中存在非空字符；
     * - 文字的样式·字符级：字符携带字体族/字号/加粗/斜体/字色/下划线/删除线；
     * - 文字的样式·单元格级：文本对齐、垂直对齐、换行、字间距、行间距；
     * - 背景颜色：backgroundColor；
     * - 边框：四边边框宽度。
     * 另：合并结构（colspan/rowspan）与筛选标记（filter）不属于样式与内容信息，
     * 但为功能性信息（合并锚点被合并拆分、筛选、打印逻辑消费），同样不可随空单元格移除。
     * @returns {boolean} 是否为空单元格
     */
    public isEmpty(): boolean {
        // 功能性信息：合并锚点承载合并结构，filter 承载筛选按钮标记
        if (this._colspan !== undefined || this._rowspan !== undefined || this._filter !== undefined) {
            return false;
        }
        // 文字信息 + 文字的样式（字符级）：逐字符检查（字体样式携带于 Char 上，无字符即无字符级样式载体）
        if (this._chars && this._chars.length > 0) {
            for (const char of this._chars) {
                if (char.char !== '' && char.char != null) return false; // 文字信息
                if (char.fontFamily !== undefined || char.fontSize !== undefined
                    || char.fontWeight !== undefined || char.fontStyle !== undefined
                    || char.fontColor !== undefined || char.underline !== undefined
                    || char.strikethrough !== undefined) {
                    return false; // 文字的样式（字符级）
                }
            }
        }
        // 文字的样式（单元格级）：对齐、换行、间距
        // if (this._textAlign !== undefined || this._alignItems !== undefined
        //     || this._wrap !== undefined || this._letterSpacing !== undefined
        //     || this._lineSpacing !== undefined) {
        //     return false;
        // } 
        // 背景颜色
        if (this._backgroundColor !== undefined) {
            return false;
        }
        // 边框
        if (this._borderTopWidth !== undefined || this._borderBottomWidth !== undefined
            || this._borderLeftWidth !== undefined || this._borderRightWidth !== undefined) {
            return false;
        }
        return true;
    }
    
    /**
     * 删除单元格值中的字符
     * @param {number} position - 字符索引
     * @param {number} count - 删除字符数量,默认删除一个字符
     */
    public deleteChar(position: number, count: number = 1) {
        const chars = this.chars;
        chars.splice(position, count);
        this.chars = chars;
    }

    /**
     * 插入单元格值中的字符
     * @param {number} position - 字符索引
     * @param {string} chars - 字符字符串
     */
    public insertChar(position: number, chars: string) {
        const _chars = Array.from(chars); // 将字符串转换为字符数组
        for (const _char of _chars) {
            const char= new Char();
            char.char = _char;
            this.chars.splice(position, 0, char);
            position++;
        }
    }

    /**
     * 深拷贝当前 Cell 对象
     * 直接通过内部字段赋值（绕过 setter），避免触发事件与不必要的回写，
     * 字符数组同样逐个调用 Char.clone() 进行深拷贝，确保快照与原数据完全解耦。
     * @returns {Cell} 深拷贝后的新 Cell 对象
     */
    public clone(): Cell {
        const cell = new Cell();
        cell._cell = this._cell;
        cell._isEdit = this._isEdit;
        cell._colspan = this._colspan;
        cell._rowspan = this._rowspan;
        cell._borderTopWidth = this._borderTopWidth;
        cell._borderBottomWidth = this._borderBottomWidth;
        cell._borderLeftWidth = this._borderLeftWidth;
        cell._borderRightWidth = this._borderRightWidth;
        cell._textAlign = this._textAlign;
        cell._alignItems = this._alignItems;
        cell._backgroundColor = this._backgroundColor;
        cell._numberFormat = this._numberFormat;
        cell._borderColor = this._borderColor;
        cell._filter = this._filter;
        cell._wrap = this._wrap;
        cell._letterSpacing = this._letterSpacing;
        cell._lineSpacing = this._lineSpacing;
        cell._chars = this._chars ? this._chars.map((c: Char) => c.clone()) : new Array<Char>();
        return cell;
    }

    /**
     * 获取单元格值中的字符字符串数组JSON字符串
     * @returns {string} 单元格值中的字符字符串数组JSON字符串
     */
    public toCharsJSON(): string {
        const charsJSON: string[] = [];
        for (const char of this._chars) {
            charsJSON.push(char.toJSON());
        }
        return JSON.stringify(charsJSON);
    }
    
    /**
     * 从字符字符串数组JSON字符串创建单元格值
     * @param {string} charsJSON - 字符字符串数组JSON字符串
     * @returns {Char[]} 单元格值中的字符字符串数组
     */
    public fromCharsJSON(charsJSON: string): Char[] {
        return Char.fromJSON(charsJSON);
    }

    /**
     * 转换为普通JSON字符串
     * @returns {string} 包含所有属性的JSON字符串
     */
    public toJSON(): string {
        return JSON.stringify({
            cell: this._cell,
            isEdit: this._isEdit,
            colspan: this._colspan,
            rowspan: this._rowspan,
            borderTopWidth: this._borderTopWidth,
            borderBottomWidth: this._borderBottomWidth,
            borderLeftWidth: this._borderLeftWidth,
            borderRightWidth: this._borderRightWidth,
            textAlign: this._textAlign,
            alignItems: this._alignItems,
            backgroundColor: this._backgroundColor,
            numberFormat: this._numberFormat,
            borderColor: this._borderColor,
            chars: this.toCharsJSON(),
            filter: this._filter,
            wrap: this._wrap,
            letterSpacing: this._letterSpacing,
            lineSpacing: this._lineSpacing
        });
    }
    
    /**
     * 从普通JSON创建单元格实例
     * @param {string} json - 包含所有属性的JSON字符串
     * @returns {Cell} 单元格实例
     */
    public static fromJSON(json: string): Cell {
        const parsed = JSON.parse(json);
        return parsed.map((item: any) => {
            const cell = new Cell();
            cell._cell = item.cell;
            cell._isEdit = item.isEdit;
            cell._colspan = item.colspan;
            cell._rowspan = item.rowspan;
            cell._borderTopWidth = item.borderTopWidth;
            cell._borderBottomWidth = item.borderBottomWidth;
            cell._borderLeftWidth = item.borderLeftWidth;
            cell._borderRightWidth = item.borderRightWidth;
            cell._textAlign = item.textAlign;
            cell._alignItems = item.alignItems;
            cell._backgroundColor = item.backgroundColor;
            cell._numberFormat = item.numberFormat;
            cell._borderColor = item.borderColor;
            cell._filter = item.filter;
            cell._wrap = item.wrap;
            cell._letterSpacing = item.letterSpacing;
            cell._lineSpacing = item.lineSpacing;
            cell.fromCharsJSON(item.chars || []);
            return cell;
        });
    }

    
}