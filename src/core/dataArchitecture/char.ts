import { DataEvents } from "../constant";
import { EventEmitter } from "../../utils/eventEmitter";

/**
 * Char 类
 * 表示电子表格单元格的单个字符，包含样式属性
 * 提供属性变更监听和事件通知机制
 *
 * @class Char
 * @extends EventEmitter
 * @example
 * const char = new Char();
 * char.char = 'H';
 * char.on(DataEvents.VALUE_CHAR_CHANGED, ({ property, oldValue, newValue }) => console.log(`${property} changed`));
 */
export class Char extends EventEmitter {
    private _char: any = '';
    private _fontFamily?: string;
    private _fontSize?: number;
    private _fontWeight?: boolean;
    private _fontStyle?: boolean;
    private _fontColor?: string;
    private _underline?: boolean;
    private _strikethrough?: boolean;

    constructor() {
        super();
    }

    /**
     * 获取字符内容
     * @returns {any} 字符内容
     */
    get char(): any {
        return this._char;
    }

    /**
     * 设置字符内容
     * @param {any} value - 字符内容
     */
    set char(value: any) {
        this.updateProperty(DataEvents.CHAR_CHANGED, '_char', value);
    }

    /**
     * 获取字体族
     * @returns {string | undefined} 字体族
     */
    get fontFamily(): string | undefined {
        return this._fontFamily;
    }

    /**
     * 设置字体族
     * @param {string | undefined} value - 字体族
     */
    set fontFamily(value: string | undefined) {
        this.updateProperty(DataEvents.CHAR_FONT_FAMILY_CHANGED, '_fontFamily', value);
    }

    /**
     * 获取字体大小
     * @returns {number | undefined} 字体大小
     */
    get fontSize(): number | undefined {
        return this._fontSize;
    }

    /**
     * 设置字体大小
     * @param {number | undefined} value - 字体大小
     */
    set fontSize(value: number | undefined) {
        this.updateProperty(DataEvents.CHAR_FONT_SIZE_CHANGED, '_fontSize', value);
    }

    /**
     * 获取字体粗细
     * @returns {boolean | undefined} 是否加粗
     */
    get fontWeight(): boolean | undefined {
        return this._fontWeight;
    }

    /**
     * 设置字体粗细
     * @param {boolean | undefined} value - 是否加粗
     */
    set fontWeight(value: boolean | undefined) {
        this.updateProperty(DataEvents.CHAR_FONT_WEIGHT_CHANGED, '_fontWeight', value);
    }

    /**
     * 获取字体样式
     * @returns {boolean | undefined} 是否斜体
     */
    get fontStyle(): boolean | undefined {
        return this._fontStyle;
    }

    /**
     * 设置字体样式
     * @param {boolean | undefined} value - 是否斜体
     */
    set fontStyle(value: boolean | undefined) {
        this.updateProperty(DataEvents.CHAR_FONT_STYLE_CHANGED, '_fontStyle', value);
    }

    /**
     * 获取字体颜色
     * @returns {string | undefined} 字体颜色
     */
    get fontColor(): string | undefined {
        return this._fontColor;
    }

    /**
     * 设置字体颜色
     * @param {string | undefined} value - 字体颜色
     */
    set fontColor(value: string | undefined) {
        this.updateProperty(DataEvents.CHAR_FONT_COLOR_CHANGED, '_fontColor', value);
    }

    /**
     * 获取下划线状态
     * @returns {boolean | undefined} 是否有下划线
     */
    get underline(): boolean | undefined {
        return this._underline;
    }

    /**
     * 设置下划线状态
     * @param {boolean | undefined} value - 是否有下划线
     */
    set underline(value: boolean | undefined) {
        this.updateProperty(DataEvents.CHAR_UNDERLINE_CHANGED, '_underline', value);
    }

    /**
     * 获取删除线状态
     * @returns {boolean | undefined} 是否有删除线
     */
    get strikethrough(): boolean | undefined {
        return this._strikethrough;
    }

    /**
     * 设置删除线状态
     * @param {boolean | undefined} value - 是否有删除线
     */
    set strikethrough(value: boolean | undefined) {
        this.updateProperty(DataEvents.CHAR_STRIKETHROUGH_CHANGED, '_strikethrough', value);
    }
    
    /**
     * 克隆当前 Char 对象
     * @returns {Char} 克隆后的新 Char 对象
     */
    public clone(): Char {
        const cloned = new Char();
        cloned._char = this._char;
        cloned._fontFamily = this._fontFamily;
        cloned._fontSize = this._fontSize;
        cloned._fontWeight = this._fontWeight;
        cloned._fontStyle = this._fontStyle;
        cloned._fontColor = this._fontColor;
        cloned._underline = this._underline;
        cloned._strikethrough = this._strikethrough;
        return cloned;
    }

    /**
     * 转换为普通对象
     * @returns {string} 包含所有属性的JSON字符串
     */
    public toJSON(): string {
        return JSON.stringify({
            char: this._char,
            fontFamily: this._fontFamily,
            fontSize: this._fontSize,
            fontWeight: this._fontWeight,
            fontStyle: this._fontStyle,
            fontColor: this._fontColor,
            underline: this._underline,
            strikethrough: this._strikethrough
        });
    }

    /**
     * 从JSON对象创建 Char 实例
     * @static
     * @param {string} json - 包含属性的JSON字符串
     * @returns {Char[]} 新创建的 Char 实例数组
     */
    public static fromJSON(json: string): Char[] {
        const parsed = JSON.parse(json);
        return parsed.map((item: any) => {
            const char = new Char();
            char._char = item.char || item._char || '';
            char._fontFamily = item.fontFamily;
            char._fontSize = item.fontSize;
            char._fontWeight = item.fontWeight;
            char._fontStyle = item.fontStyle;
            char._fontColor = item.fontColor;
            char._underline = item.underline;
            char._strikethrough = item.strikethrough;
            return char;
        });
    }
}
