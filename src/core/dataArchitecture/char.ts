/**
 * Char 纯数据类
 *
 * 表示电子表格单元格中的单个字符及其字符级样式。
 *
 * ## 为什么不再是 EventEmitter（性能约定，请勿回退）
 * 历史实现让 `Char` 继承 `EventEmitter`，于是**每一个字符**都在内存中携带一个
 * 独立的 `Map` 实例。一张 10 万格、平均 10 字的表格会产生约 100 万个 Map，
 * 打开大文件时直接导致渲染进程内存耗尽（页面崩溃）。
 * 经全仓核查，**没有任何代码订阅过 Cell / Char 的事件**（仅存在于文档注释中），
 * 因此这里退化为纯数据类：
 * - 无可变共享状态、无监听表，单实例内存占用仅为原先的零头；
 * - 样式字段按需赋值（未赋值的字段不成为自有属性），稀疏天然成立。
 *
 * 对外契约保持不变：`char` / `fontFamily` / `fontSize` / `fontWeight` /
 * `fontStyle` / `fontColor` / `underline` / `strikethrough`，以及
 * {@link Char.clone} / {@link Char.toJSON} / {@link Char.fromJSON}。
 *
 * @class Char
 * @example
 * const char = new Char();
 * char.char = 'H';
 * char.fontSize = 12;
 */
export class Char {
    /** 字符内容 */
    public char: any = '';
    /** 字体族 */
    public fontFamily?: string;
    /** 字体大小 */
    public fontSize?: number;
    /** 字体粗细（加粗） */
    public fontWeight?: boolean;
    /** 字体样式（斜体） */
    public fontStyle?: boolean;
    /** 字体颜色 */
    public fontColor?: string;
    /** 下划线 */
    public underline?: boolean;
    /** 删除线 */
    public strikethrough?: boolean;

    /**
     * 以字符内容构造一个 Char（热路径工厂，等价于 new Char() 后赋 char）
     * @param {any} char - 字符内容
     * @returns {Char} Char 实例
     */
    public static of(char: any): Char {
        const item = new Char();
        item.char = char;
        return item;
    }

    /**
     * 克隆当前 Char 对象
     * @returns {Char} 克隆后的新 Char 对象
     */
    public clone(): Char {
        const cloned = new Char();
        cloned.char = this.char;
        // 仅复制「已存在」的样式字段，保持实例的稀疏性（避免无意义地物化全部字段）
        if (this.fontFamily !== undefined) cloned.fontFamily = this.fontFamily;
        if (this.fontSize !== undefined) cloned.fontSize = this.fontSize;
        if (this.fontWeight !== undefined) cloned.fontWeight = this.fontWeight;
        if (this.fontStyle !== undefined) cloned.fontStyle = this.fontStyle;
        if (this.fontColor !== undefined) cloned.fontColor = this.fontColor;
        if (this.underline !== undefined) cloned.underline = this.underline;
        if (this.strikethrough !== undefined) cloned.strikethrough = this.strikethrough;
        return cloned;
    }

    /**
     * 转换为普通对象
     * @returns {string} 包含所有属性的JSON字符串
     */
    public toJSON(): string {
        return JSON.stringify({
            char: this.char,
            fontFamily: this.fontFamily,
            fontSize: this.fontSize,
            fontWeight: this.fontWeight,
            fontStyle: this.fontStyle,
            fontColor: this.fontColor,
            underline: this.underline,
            strikethrough: this.strikethrough
        });
    }

    /**
     * 从普通对象创建 Char 实例（稀疏复制：只复制确实存在的字段）
     *
     * 装载存档时的主通道。刻意不写成「逐字段赋值」——那样会把 7 个样式字段
     * 全部物化为自有属性，使稀疏性丧失、装载大表时内存显著抬高。
     * @static
     * @param {any} item - 形如 `{ char, fontFamily?, ... }` 的普通对象
     * @returns {Char} Char 实例
     */
    public static fromRaw(item: any): Char {
        const char = new Char();
        if (!item) return char;
        char.char = item.char ?? item._char ?? '';
        if (item.fontFamily !== undefined) char.fontFamily = item.fontFamily;
        if (item.fontSize !== undefined) char.fontSize = item.fontSize;
        if (item.fontWeight !== undefined) char.fontWeight = item.fontWeight;
        if (item.fontStyle !== undefined) char.fontStyle = item.fontStyle;
        if (item.fontColor !== undefined) char.fontColor = item.fontColor;
        if (item.underline !== undefined) char.underline = item.underline;
        if (item.strikethrough !== undefined) char.strikethrough = item.strikethrough;
        return char;
    }

    /**
     * 从JSON对象创建 Char 实例
     * @static
     * @param {string} json - 包含属性的JSON字符串
     * @returns {Char[]} 新创建的 Char 实例数组
     */
    public static fromJSON(json: string): Char[] {
        const parsed = JSON.parse(json);
        return parsed.map((item: any) => Char.fromRaw(item));
    }
}
