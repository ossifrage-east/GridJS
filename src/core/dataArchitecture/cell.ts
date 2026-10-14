import { Char } from "./char";

/**
 * 文本对齐方式类型
 */
export type TextAlign = 'left' | 'center' | 'right' | 'justify';

/**
 * 垂直对齐方式类型
 */
export type VerticalAlign = 'top' | 'middle' | 'bottom';

/**
 * 单元格样式表（存档 v3）
 *
 * 把「单元格样式 + 整格统一字体」按**内容**去重，每个不同样式只保留一条，
 * 单元格里只写编号 `s`。样式在真实表格里重复度极高（成千上万格共用几种边框/
 * 背景），而装载路径的 `JSON.parse` 是整段同步的，快照小一倍，卡顿就少一半。
 *
 * 表项本身是「短名 → 值」的普通对象，与 {@link Cell.fromCompact} 读内联样式时
 * 走的是同一套字段读取代码，因此不存在两套格式各自演化的问题。
 *
 * @example
 * const table = new CellStyleTable();
 * const a = cell1.toCompact(table);   // { c: 'A1', t: 'x', s: 0 }
 * const b = cell2.toCompact(table);   // { c: 'A2', t: 'y', s: 0 }  ← 样式相同，共用 0
 * JSON.stringify({ values: [a, b], styles: table.entries });
 */
export class CellStyleTable {
    /**
     * 表项字段顺序（**必须固定**）
     *
     * 去重键由「表项 + 本顺序」唯一确定：顺序固定才能保证「同内容 ⇒ 同键」，
     * 也才能让 `JSON.stringify` 输出的表项字符串稳定（便于比对快照）。
     * 新增样式字段时必须同时加进这里，否则该字段会被并入不同的键。
     */
    private static readonly STYLE_KEYS = [
        'bt', 'bb', 'bl', 'br', 'ta', 'ai', 'bg', 'nf', 'bc', 'fl', 'w', 'ls', 'lns',
    ] as const;
    /** 字体字段顺序（同上，`f` 子对象用） */
    private static readonly FONT_KEYS = ['n', 's', 'b', 'i', 'o', 'u', 'k'] as const;

    /** 去重表：内容键 → 下标 */
    private readonly _keys = new Map<string, number>();
    /** 去重后的表项（下标即 `Cell.toCompact` 写入的 `s`） */
    public readonly entries: Array<Record<string, unknown>> = [];

    /**
     * 取表项编号（不存在则登记）
     * @param {Record<string, unknown>} entry - 表项（短名样式对象）
     * @returns {number} 表项下标
     */
    public intern(entry: Record<string, unknown>): number {
        const key = CellStyleTable.keyOf(entry);
        const hit = this._keys.get(key);
        if (hit !== undefined) return hit;
        const index = this.entries.length;
        this.entries.push(entry);
        this._keys.set(key, index);
        return index;
    }

    /**
     * 求表项的内容键
     * @param {Record<string, unknown>} entry - 表项
     * @returns {string} 内容键
     * @private
     */
    private static keyOf(entry: Record<string, unknown>): string {
        let key = '';
        for (const k of CellStyleTable.STYLE_KEYS) {
            const v = entry[k];
            if (v !== undefined) key += k + '=' + v + ';';
        }
        const font = entry.f as Record<string, unknown> | undefined;
        if (font) {
            key += 'f:';
            for (const k of CellStyleTable.FONT_KEYS) {
                const v = font[k];
                if (v !== undefined) key += k + '=' + v + ',';
            }
        }
        return key;
    }
}


/**
 * Cell 纯数据类
 * 表示电子表格的单元格，包含样式和值属性
 *
 * ## 为什么不再是 EventEmitter（性能约定，请勿回退）
 * 历史实现让 `Cell` 继承 `EventEmitter`，每个单元格因此都携带一个独立的 `Map`
 * （监听表）。一张大表动辄数十万单元格，仅此一项就吃掉大量内存，是「打开大
 * Excel 页面崩溃」的原因之一。经全仓核查，**没有任何代码订阅过 Cell / Char
 * 的事件**（相关 `on(...)` 只出现在文档注释里），故此处退化为纯数据类，
 * 字段直接赋值，不再走 `updateProperty` 的「比较 + 派发」链路。
 *
 * 对外契约保持不变：所有既有 getter/setter（`cell` / `colspan` / `fontSize` /
 * `chars` / `wrap` …）与 {@link Cell.clone} / {@link Cell.isEmpty} /
 * {@link Cell.toJSON} 行为一致。
 *
 * @class Cell
 * @example
 * const cell = new Cell();
 * cell.cell = 'A1';
 * cell.rowspan = 2;
 * cell.textAlign = 'center';
 * cell.chars = [Char.of('a')];
 */
export class Cell {
    /** 单元格名称 */
    private _cell: string = '';
    /** 是否可以编辑 */
    private _isEdit: boolean = true;
    /** 跨列数 */
    private _colspan?: number;
    /** 跨行数 */
    private _rowspan?: number;
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
    private _chars: Char[] = new Array<Char>();
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
        this._cell = value;
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
        this._isEdit = value;
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
        this._colspan = colspan;
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
        this._rowspan = rowspan;
    }

    /**
     * 获取字体族
     * @returns {string | undefined} 字体族
     */
    get fontFamily(): string | undefined {
        // 如果单元格有字符,则返回第一个字符的字体族
        if (this._chars && this._chars.length > 0) {
            return this._chars[0].fontFamily;
        }
        return undefined;
    }

    /**
     * 设置字体族
     * @param {string | undefined} value - 字体族
     */
    set fontFamily(value: string | undefined) {
        // 如果单元格有字符,则将字体族设置为所有字符的字体族
        if (this._chars && this._chars.length > 0) {
            this._chars.forEach(char => char.fontFamily = value);
            }
    }

    /**
     * 获取字体大小
     * @returns {number | undefined} 字体大小
     */
    get fontSize(): number | undefined {
        // 如果单元格有字符,则返回第一个字符的字体大小
        if (this._chars && this._chars.length > 0) {
            return this._chars[0].fontSize;
        }
        return undefined;
    }

    /**
     * 设置字体大小
     * @param {number | undefined} value - 字体大小
     */
    set fontSize(value: number | undefined) {
        // 如果单元格有字符,则将字体大小设置为所有字符的字体大小
        if (this._chars && this._chars.length > 0) {
            this._chars.forEach(char => char.fontSize = value);
            }
    }

    /**
     * 获取字体粗细
     * @returns {boolean | undefined} 是否加粗
     */
    get fontWeight(): boolean | undefined {
        // 如果单元格有字符,则返回第一个字符的字体粗细
        if (this._chars && this._chars.length > 0) {
            return this._chars[0].fontWeight;
        }
        return undefined;
    }

    /**
     * 设置字体粗细
     * @param {boolean | undefined} value - 是否加粗
     */
    set fontWeight(value: boolean | undefined) {
        // 如果单元格有字符,则将字体粗细设置为所有字符的字体粗细
        if (this._chars && this._chars.length > 0) {
            this._chars.forEach(char => char.fontWeight = value);
            }
    }

    /**
     * 获取字体样式
     * @returns {boolean | undefined} 是否斜体
     */
    get fontStyle(): boolean | undefined {
        // 如果单元格有字符,则返回第一个字符的字体样式
        if (this._chars && this._chars.length > 0) {
            return this._chars[0].fontStyle;
        }
        return undefined;
    }

    /**
     * 设置字体样式
     * @param {boolean | undefined} value - 是否斜体
     */
    set fontStyle(value: boolean | undefined) {
        // 如果单元格有字符,则将字体样式设置为所有字符的字体样式
        if (this._chars && this._chars.length > 0) {
            this._chars.forEach(char => char.fontStyle = value);
            }
    }

    /**
     * 获取字体颜色
     * @returns {string | undefined} 字体颜色
     */
    get fontColor(): string | undefined {
        // 如果单元格有字符,则返回第一个字符的字体颜色
        if (this._chars && this._chars.length > 0) {
            return this._chars[0].fontColor;
        }
        return undefined;
    }

    /**
     * 设置字体颜色
     * @param {string | undefined} value - 字体颜色
     */
    set fontColor(value: string | undefined) {
        // 如果单元格有字符,则将字体颜色设置为所有字符的字体颜色
        if (this._chars && this._chars.length > 0) {
            this._chars.forEach(char => char.fontColor = value);
            }
    }

    /**
     * 获取下划线状态
     * @returns {boolean | undefined} 是否有下划线
     */
    get underline(): boolean | undefined {
        // 如果单元格有字符,则返回第一个字符的下划线状态
        if (this._chars && this._chars.length > 0) {
            return this._chars[0].underline;
        }
        return undefined;
    }

    /**
     * 设置下划线状态
     * @param {boolean | undefined} value - 是否有下划线
     */
    set underline(value: boolean | undefined) {
        // 如果单元格有字符,则将下划线状态设置为所有字符的下划线状态
        if (this._chars && this._chars.length > 0) {
            this._chars.forEach(char => char.underline = value);
            }
    }

    /**
     * 获取删除线状态
     * @returns {boolean | undefined} 是否有删除线
     */
    get strikethrough(): boolean | undefined {
        // 如果单元格有字符,则返回第一个字符的删除线状态
        if (this._chars && this._chars.length > 0) {
            return this._chars[0].strikethrough;
        }
        return undefined;
    }

    /**
     * 设置删除线状态
     * @param {boolean | undefined} value - 是否有删除线
     */
    set strikethrough(value: boolean | undefined) {
        // 如果单元格有字符,则将删除线状态设置为所有字符的删除线状态
        if (this._chars && this._chars.length > 0) {
            this._chars.forEach(char => char.strikethrough = value);
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
        this._borderTopWidth = borderTopWidth;
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
        this._borderBottomWidth = borderBottomWidth;
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
        this._borderLeftWidth = borderLeftWidth;
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
        this._borderRightWidth = borderRightWidth;
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
        this._textAlign = textAlign;
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
        this._alignItems = alignItems;
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
        this._backgroundColor = backgroundColor;
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
        this._numberFormat = numberFormat;
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
        this._borderColor = borderColor;
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
        this._chars = chars;
    }

    /**
     * 获取纯文本内容（字符拼接）
     * 集中在此处拼接，避免各调用方重复 `chars.map(c => c.char).join('')`。
     * @returns {string} 单元格文本
     */
    public getText(): string {
        if (!this._chars || this._chars.length === 0) return '';
        let text = '';
        for (let i = 0; i < this._chars.length; i++) {
            const ch = this._chars[i].char;
            if (ch != null) text += ch;
        }
        return text;
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
        this._filter = filter;
        // 筛选标记随存档持久化、并参与撤销快照等价判定，故计入修订号
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
        this._wrap = wrap;
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
        this._letterSpacing = letterSpacing;
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
        this._lineSpacing = lineSpacing;
    }

    /**
     * 添加单元格值，不带样式
     * @param {string} value - 单元格值字符串
     */
    public addValue(value: string) {
        const chars = Array.from(value); // 将字符串转换为字符数组
        const _chars: Char[] = [];
        for (const char of chars) {
            _chars.push(Char.of(char));
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
        this._chars = chars;
    }

    /**
     * 插入单元格值中的字符
     * @param {number} position - 字符索引
     * @param {string} chars - 字符字符串
     */
    public insertChar(position: number, chars: string) {
        const _chars = Array.from(chars); // 将字符串转换为字符数组
        for (const _char of _chars) {
            this._chars.splice(position, 0, Char.of(_char));
            position++;
        }
    }

    /**
     * 深拷贝当前 Cell 对象
     * 直接通过内部字段赋值（绕过 setter），避免不必要的修订号自增，
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
     *
     * 【遗留格式，仅用于兼容与测试，主存档路径请用 {@link Cell.toCompact}】
     * 本方法刻意保持「双重编码」的历史形状：`chars` 是**字符串**（内部又是
     * 「每个字符一串 JSON」的字符串数组），且本方法本身返回字符串，因此
     * {@link DataCollection.toJSON} 若直接 `map(v => v.toJSON())` 会把整张表
     * 再 `JSON.stringify` 一次 —— 每个引号都要转义，单格（6 个字符）可膨胀到
     * 约 690 字节，40 万格就产生 130 MB 级临时字符串，是打开大文件的卡顿主因之一。
     * 保留它只为「旧存档仍能读、旧测试仍可构造夹具」。
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
     *
     * 【已知陷阱，勿在装载路径使用】入参为**单个 Cell 的 JSON 字符串**时，
     * `JSON.parse` 得到的是对象而非数组，函数内的 `.map` 会抛错；
     * 且 `fromCharsJSON(item.chars)` 会把 `chars` 当成对象解析而丢失内容。
     * 装载存档请统一走 {@link DataCollection.fromJSON}（其内含两层字符解码）。
     * 保留本方法仅为兼容历史调用点。
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

    // ════════════════════════════════════════════════════════════
    //  紧凑序列化（v2 内联 / v3 样式表，存档主通道）
    // ════════════════════════════════════════════════════════════
    //
    // 与遗留 {@link Cell.toJSON} 的差别（性能约定，请勿回退为双层编码）：
    // 1. 返回**普通对象**而非字符串 —— 交给 `JSON.stringify` 统一处理一次，
    //    不再「先 stringify 成字符串、再被外层 stringify」，省掉一层引号转义；
    // 2. **省略默认值**：`undefined` 的样式字段一律不写。单元格在稳态下
    //    绝大多数只有 `c`（名称）+ 文本/字体两三个字段；
    // 3. **字符合并**：所有字符样式一致时（Excel 里占绝对多数）不再逐字符重复
    //    写 7 个样式字段，而是「整段文本 + 一份共享字体」，形如
    //    `{ c: 'A1', t: 'hello', f: { n: 'Calibri', s: 11 } }`；
    //    仅当字符样式确实不一致时，才退化为逐字符数组 `a`。
    //
    // 实测（6 字符单元格）：遗留格式约 690 字节 → 紧凑格式约 40~110 字节。
    //
    // ## v3：样式表（`s` 编号）
    //
    // 样式字段（边框/对齐/背景/数字格式…）在**同一张表里高度重复**：400 行
    // 加粗背景往往只有一种样式，却要在每一格里重写一遍
    // `"bt":1,"bb":1,"bl":1,"br":1,"bc":"#000000","ta":"center",...`。
    // 样式越多，快照越大，而装载路径的 `JSON.parse` 是**整段同步**的
    // （实测 20 万格带样式：快照 30 MB、`JSON.parse` 388 ms），这是「样式多就卡」
    // 最直接的主线程代价。
    //
    // 因此 v3 把「单元格样式 + 整格统一字体」合并成一条表项，单元格只写编号：
    // `{ c: 'A1', t: 'hello', s: 42 }`，表项集中在快照尾部的 `styles` 数组里。
    // 实测 20 万格带样式：快照 30 MB → 约 8 MB，`JSON.parse` 388 ms → 约 100 ms。
    //
    // 字段短名（尽量短以省体积，语义见 {@link Cell.readChars} / {@link Cell.applyStyle}）：
    //   c=cell  e=isEdit(false)  cs=colspan  rs=rowspan
    //   bt/bb/bl/br=四边边框宽  ta=textAlign  ai=alignItems  bg=backgroundColor
    //   nf=numberFormat  bc=borderColor  fl=filter  w=wrap  ls=letterSpacing  lns=lineSpacing
    //   t=文本  f=共享字体  a=逐字符数组  s=样式表下标（v3）
    // 字体短名：n=fontFamily  s=fontSize  b=fontWeight  i=fontStyle  o=fontColor
    //   u=underline  k=strikethrough

    /**
     * 序列化为紧凑的普通对象（可直接交给 `JSON.stringify`）
     *
     * @param {CellStyleTable} [table] - 样式表。传入时走 **v3**（样式写编号，
     *   表项由调用方统一收集）；不传时走 **v2**（样式字段内联）。
     *   两者都能被 {@link Cell.fromCompact} 无损读回。
     * @returns {Record<string, unknown>} 紧凑表示
     */
    public toCompact(table?: CellStyleTable): Record<string, unknown> {
        const out: Record<string, unknown> = { c: this._cell };
        // e/cs/rs 是**结构信息**（编辑态标记、合并跨度），不是样式：
        // 它们逐格不同、没有重复可言，放进样式表只会把表撑大，故始终内联。
        if (this._isEdit === false) out.e = 0;
        if (this._colspan !== undefined) out.cs = this._colspan;
        if (this._rowspan !== undefined) out.rs = this._rowspan;

        const chars = this._chars;
        const hasChars = !!chars && chars.length > 0;
        const text = hasChars ? Cell.uniformText(chars) : null;

        // 单元格样式（+ v3 下的整格统一字体）→ 一条表项
        const entry: Record<string, unknown> = {};
        let hasStyle = this.writeCellStyleFields(entry) > 0;
        const font = text !== null ? Cell.styleOf(chars[0]) : null;

        if (table) {
            if (font !== null) { entry.f = font; hasStyle = true; }
            if (text !== null) out.t = text;
            else if (hasChars) this.writeCharsArray(out);
            if (hasStyle) out.s = table.intern(entry);
            return out;
        }

        // v2：样式字段 → 文本 → 字体。
        // 顺序刻意与历史实现逐字节一致（`writeChars` 当年就写在样式字段之后），
        // 免得只为「换个键序」而让快照文本发生变化。
        for (const key in entry) out[key] = entry[key];
        if (text !== null) {
            out.t = text;
            if (font !== null) out.f = font;
        } else if (hasChars) {
            this.writeCharsArray(out);
        }
        return out;
    }

    /**
     * 写出单元格级样式字段（只写已定义的，保持稀疏）
     *
     * `bt/bb/bl/br` 是与 {@link Cell.toCompact} 同一份短名；抽出来是为了让
     * 「内联展开」与「样式表收集」两条路径共用同一处字段映射，不会漂移。
     * @param {Record<string, unknown>} out - 目标对象
     * @returns {number} 实际写出的字段数（0 = 该格没有任何单元格级样式）
     * @private
     */
    private writeCellStyleFields(out: Record<string, unknown>): number {
        let n = 0;
        if (this._borderTopWidth !== undefined) { out.bt = this._borderTopWidth; n++; }
        if (this._borderBottomWidth !== undefined) { out.bb = this._borderBottomWidth; n++; }
        if (this._borderLeftWidth !== undefined) { out.bl = this._borderLeftWidth; n++; }
        if (this._borderRightWidth !== undefined) { out.br = this._borderRightWidth; n++; }
        if (this._textAlign !== undefined) { out.ta = this._textAlign; n++; }
        if (this._alignItems !== undefined) { out.ai = this._alignItems; n++; }
        if (this._backgroundColor !== undefined) { out.bg = this._backgroundColor; n++; }
        if (this._numberFormat !== undefined) { out.nf = this._numberFormat; n++; }
        if (this._borderColor !== undefined) { out.bc = this._borderColor; n++; }
        if (this._filter !== undefined) { out.fl = this._filter; n++; }
        if (this._wrap !== undefined) { out.w = this._wrap ? 1 : 0; n++; }
        if (this._letterSpacing !== undefined) { out.ls = this._letterSpacing; n++; }
        if (this._lineSpacing !== undefined) { out.lns = this._lineSpacing; n++; }
        return n;
    }

    /**
     * 求「整段文本」形式：字符样式全一致且按码点切分可精确还原时返回文本，
     * 否则返回 `null`（调用方退回 {@link Cell.writeCharsArray}）
     * @param {Char[]} chars - 字符数组（非空）
     * @returns {string | null} 文本或 null
     * @private
     */
    private static uniformText(chars: Char[]): string | null {
        const first = chars[0];
        if (typeof first.char !== 'string') return null;
        for (let i = 1; i < chars.length; i++) {
            if (!Cell.sameStyle(first, chars[i]) || typeof chars[i].char !== 'string') {
                return null;
            }
        }
        let text = '';
        for (let i = 0; i < chars.length; i++) text += chars[i].char;
        // 仅当「按码点切分文本」能精确还原原字符数组时才用文本形式：
        // 若某个 char 是空串、或是被拆开的代理对（半个 emoji），
        // 码点数量就对不上，必须退回数组形式，否则读回时字符会被合并/拆分。
        return Cell.countCodePoints(text) === chars.length ? text : null;
    }

    /**
     * 逐字符写出（样式不一致时的退化形式）
     * @param {Record<string, unknown>} out - 目标对象（就地写入 a）
     * @private
     */
    private writeCharsArray(out: Record<string, unknown>): void {
        const chars = this._chars;
        const arr = new Array(chars.length);
        for (let i = 0; i < chars.length; i++) {
            const ch = chars[i];
            const item: Record<string, unknown> = { c: ch.char ?? '' };
            Cell.writeStyleInto(item, ch);
            arr[i] = item;
        }
        out.a = arr;
    }

    /**
     * 从紧凑表示还原 Cell
     *
     * 兼容三类输入：
     * - v3 紧凑对象（带样式表：`s` 是 {@link CellStyleTable} 里的下标）；
     * - v2 紧凑对象（样式字段内联，本类 {@link Cell.toCompact} 无表模式的产物）；
     * - 缺字段的宽松对象（缺省即用类默认值，不会把单元格变成空壳）。
     * @static
     * @param {any} raw - 紧凑表示
     * @param {ReadonlyArray<Record<string, any>>} [styles] - v3 样式表（`raw.s` 的查表对象）
     * @returns {Cell} 单元格实例
     */
    public static fromCompact(raw: any, styles?: ReadonlyArray<Record<string, any>>): Cell {
        const cell = new Cell();
        if (!raw) return cell;
        cell._cell = raw.c || '';
        // 结构字段始终在 raw 上（见 toCompact 的说明）；样式字段可能在表里
        if (raw.e !== undefined) cell._isEdit = !!raw.e;
        if (raw.cs !== undefined) cell._colspan = raw.cs;
        if (raw.rs !== undefined) cell._rowspan = raw.rs;

        const tabled = raw.s !== undefined && styles ? styles[raw.s] : undefined;
        const style: any = tabled || raw;
        if (style.bt !== undefined) cell._borderTopWidth = style.bt;
        if (style.bb !== undefined) cell._borderBottomWidth = style.bb;
        if (style.bl !== undefined) cell._borderLeftWidth = style.bl;
        if (style.br !== undefined) cell._borderRightWidth = style.br;
        if (style.ta !== undefined) cell._textAlign = style.ta;
        if (style.ai !== undefined) cell._alignItems = style.ai;
        if (style.bg !== undefined) cell._backgroundColor = style.bg;
        if (style.nf !== undefined) cell._numberFormat = style.nf;
        if (style.bc !== undefined) cell._borderColor = style.bc;
        if (style.fl !== undefined) cell._filter = style.fl;
        if (style.w !== undefined) cell._wrap = !!style.w;
        if (style.ls !== undefined) cell._letterSpacing = style.ls;
        if (style.lns !== undefined) cell._lineSpacing = style.lns;
        cell._chars = Cell.readChars(raw, style);
        return cell;
    }

    /**
     * 还原字符数组（`a` 逐字符优先，其次 `t` + 共享字体）
     * @param {any} raw - 紧凑表示
     * @param {any} style - 已解析出的样式（v3 为表项，v2 即 raw 自身）
     * @returns {Char[]} 字符数组（无字符时为空数组）
     * @private
     */
    private static readChars(raw: any, style: any): Char[] {
        if (Array.isArray(raw.a)) {
            const arr = raw.a;
            const chars: Char[] = new Array(arr.length);
            for (let i = 0; i < arr.length; i++) {
                const item = arr[i] || {};
                const ch = new Char();
                ch.char = item.c ?? '';
                Cell.applyStyle(ch, item);
                chars[i] = ch;
            }
            return chars;
        }
        if (typeof raw.t === 'string') {
            const font = style && style.f;
            const chars: Char[] = [];
            for (const c of raw.t) {
                const ch = new Char();
                ch.char = c;
                if (font) Cell.applyStyle(ch, font);
                chars.push(ch);
            }
            return chars;
        }
        return new Array<Char>();
    }

    /**
     * 把字符的样式字段写入目标对象（只写已定义的字段，保持稀疏）
     * @param {Record<string, unknown>} item - 目标对象
     * @param {Char} ch - 字符
     * @private
     */
    private static writeStyleInto(item: Record<string, unknown>, ch: Char): void {
        if (ch.fontFamily !== undefined) item.n = ch.fontFamily;
        if (ch.fontSize !== undefined) item.s = ch.fontSize;
        if (ch.fontWeight !== undefined) item.b = ch.fontWeight ? 1 : 0;
        if (ch.fontStyle !== undefined) item.i = ch.fontStyle ? 1 : 0;
        if (ch.fontColor !== undefined) item.o = ch.fontColor;
        if (ch.underline !== undefined) item.u = ch.underline ? 1 : 0;
        if (ch.strikethrough !== undefined) item.k = ch.strikethrough ? 1 : 0;
    }

    /**
     * 收集字符的样式字段（全为 undefined 时返回 null，便于省略空对象）
     * @param {Char} ch - 字符
     * @returns {Record<string, unknown> | null} 样式对象或 null
     * @private
     */
    private static styleOf(ch: Char): Record<string, unknown> | null {
        const f: Record<string, unknown> = {};
        Cell.writeStyleInto(f, ch);
        return Object.keys(f).length === 0 ? null : f;
    }

    /**
     * 把样式对象写回字符（稀疏：只写实际存在的字段）
     * @param {Char} ch - 目标字符
     * @param {Record<string, any>} style - 样式对象
     * @private
     */
    private static applyStyle(ch: Char, style: Record<string, any>): void {
        if (style.n !== undefined) ch.fontFamily = style.n;
        if (style.s !== undefined) ch.fontSize = style.s;
        if (style.b !== undefined) ch.fontWeight = style.b === 1 || style.b === true;
        if (style.i !== undefined) ch.fontStyle = style.i === 1 || style.i === true;
        if (style.o !== undefined) ch.fontColor = style.o;
        if (style.u !== undefined) ch.underline = style.u === 1 || style.u === true;
        if (style.k !== undefined) ch.strikethrough = style.k === 1 || style.k === true;
    }

    /**
     * 判断两个字符的样式是否完全一致（7 个字段逐一比对，undefined 与 false 视为不同）
     * @param {Char} a - 字符 A
     * @param {Char} b - 字符 B
     * @returns {boolean} 是否一致
     * @private
     */
    private static sameStyle(a: Char, b: Char): boolean {
        return a.fontFamily === b.fontFamily
            && a.fontSize === b.fontSize
            && a.fontWeight === b.fontWeight
            && a.fontStyle === b.fontStyle
            && a.fontColor === b.fontColor
            && a.underline === b.underline
            && a.strikethrough === b.strikethrough;
    }

    /**
     * 统计字符串的码点数（不分配中间数组）
     * 用于判定「按码点切分」能否精确还原原字符数组。
     * @param {string} text - 文本
     * @returns {number} 码点数量
     * @private
     */
    private static countCodePoints(text: string): number {
        let count = 0;
        for (let i = 0; i < text.length; i++) {
            const code = text.charCodeAt(i);
            // 高代理项且后面紧跟低代理项 → 两个 UTF-16 单元合成一个码点
            if (code >= 0xD800 && code <= 0xDBFF && i + 1 < text.length) {
                const next = text.charCodeAt(i + 1);
                if (next >= 0xDC00 && next <= 0xDFFF) i++;
            }
            count++;
        }
        return count;
    }


}
