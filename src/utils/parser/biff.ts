/**
 * 原生 BIFF8 记录解析模块（.xls 二进制格式，无第三方依赖）
 *
 * Excel 97-2003 的 .xls 文件在 OLE2 复合文档的 `Workbook` 流中，存放的是
 * 一串连续的 **BIFF 记录**（Binary Interchange File Format）。每条记录的结构：
 *
 * ```
 * ┌────────────┬────────────┬──────────────────┐
 * │ 记录类型 u16 │ 长度 u16   │ 记录数据（length 字节） │
 * └────────────┴────────────┴──────────────────┘
 * ```
 *
 * 工作表在流中的位置由 **BOUNDSHEET** 记录给出（记录内含子流起始偏移），
 * 因此解析需**两轮扫描**：
 *
 * 1. **第一轮**：扫描全流收集全局记录——SST（共享字符串表）、FONT、
 *    FORMAT（数字格式）、XF（样式索引）、PALETTE（调色板）、BOUNDSHEET（工作表位置）
 * 2. **第二轮**：按 BOUNDSHEET 给出的偏移分别扫描每个工作表子流，读取单元格数据
 *
 * **SST 的 CONTINUE 机制**（本模块最棘手的部分）：
 * 共享字符串表可能超过 65535 字节而被拆成 SST + 多个 CONTINUE 记录。
 * 关键在于：一个字符串**被切断为两半时，后半部分的开头会重新带一个 grbit 标志字节**
 * （0 = UTF-16LE，1 = 压缩 8 位），前半部分末尾的 grbit 不再生效。
 *
 * @module biff
 * @author 东方鹗
 */

/** BIFF 记录类型 */
const REC = {
    /** 工作簿全局子流开始 */
    BOF: 0x0809,
    /** 子流结束 */
    EOF: 0x000A,
    /** 工作表位置与名称 */
    BOUNDSHEET: 0x0085,
    /** 共享字符串表 */
    SST: 0x00FC,
    /** SST 的续接记录 */
    CONTINUE: 0x003C,
    /** 双精度数值 */
    NUMBER: 0x0203,
    /** 整数（压缩存储） */
    RK: 0x027E,
    /** 多列整数（压缩存储） */
    MULRK: 0x00BD,
    /** 字符串（引用 SST 索引） */
    LABELSST: 0x00FD,
    /** 字符串（直接内联） */
    LABEL: 0x0204,
    /** 富文本字符串 */
    RSTRING: 0x00D6,
    /** 空单元格（但有样式） */
    BLANK: 0x0201,
    /** 多列空单元格 */
    MULBLANK: 0x00BE,
    /** 布尔值或错误值 */
    BOOLERR: 0x0205,
    /** 公式 */
    FORMULA: 0x0006,
    /** 公式的字符串结果（紧随 FORMULA 之后） */
    STRING: 0x0207,
    /** 数组公式 */
    ARRAY: 0x0221,
    /** 样式索引记录 */
    XF: 0x00E0,
    /** 字体 */
    FONT: 0x0031,
    /** 数字格式 */
    FORMAT: 0x041E,
    /** 内置数字格式 */
    FORMAT2: 0x001E,
    /** 调色板 */
    PALETTE: 0x0092,
    /** 合并单元格区域 */
    MERGEDCELLS: 0x00E5,
    /** 列宽 */
    COLINFO: 0x007D,
    /** 行属性（含行高） */
    ROW: 0x0208,
    /** 数据区域范围 */
    DIMENSIONS: 0x0200,
    /** 默认行高 */
    DEFAULTROWHEIGHT: 0x0225,
    /** 默认列宽 */
    DEFCOLWIDTH: 0x0055,
    /** 绘制对象（图片容器） */
    drawing: 0x00EC,
    /** Escher 容器 */
    msofbtSpContainer: 0xF004,
    /** Escher BSE（存储的 blip 引用） */
    msofbtBSE: 0xF007,
    /** Escher BLIP 图片记录 */
    msofbtBlip_EMF: 0xF01A,
    msofbtBlip_WMF: 0xF01B,
    msofbtBlip_PICT: 0xF01C,
    msofbtBlip_JPEG: 0xF01D,
    msofbtBlip_PNG: 0xF01E,
    msofbtBlip_DIB: 0xF01F
};

/** BIFF 版本号：8 表示 Excel 97+ */
const BIFF8_VERSION = 0x0600;
/** BIFF5 版本号（Excel 5.0） */
const BIFF5_VERSION = 0x0500;

/**
 * BIFF 字体信息
 */
export interface BiffFont {
    name: string;
    height: number;
    bold: boolean;
    italic: boolean;
    underline: boolean;
    strikethrough: boolean;
    /** 颜色索引（指向 PALETTE） */
    colorIndex: number;
    /** 是否为上标/下标 */
    superscript: boolean;
    subscript: boolean;
}

/**
 * BIFF 样式索引（XF）信息
 *
 * 字段口径严格对齐 BIFF8 的 20 字节 XF 记录（见 {@link BiffParser.parseXf}）：
 * 边框线型是 **4 位**宽（不是 3 位），四边颜色**各自独立**存放。
 */
export interface BiffXf {
    /** 字体索引 */
    fontIndex: number;
    /** 数字格式索引 */
    formatIndex: number;
    /** 水平对齐（0=常规 1=左 2=居中 3=右 4=填充 5=两端 6=跨列居中 7=分散） */
    horizontalAlignment: number;
    /** 垂直对齐（0=顶 1=居中 2=底 3=两端 4=分散） */
    verticalAlignment: number;
    /** 自动换行 */
    wrapText: boolean;
    /** 文本旋转角度（0-180，255 表示竖排） */
    rotation: number;
    /** 缩进 */
    indent: number;
    /** 是否缩小字体以适应单元格 */
    shrinkToFit: boolean;
    /** 是否为父 XF（style XF）——由 offset+4 的 bit2 判定 */
    isStyle: boolean;
    /** 父样式 XF 索引（offset+4 的 bit15-4；style XF 恒为 0xFFF） */
    parentXfIndex: number;
    /** 边框线型（上/下/左/右），索引进 `BIFF_BORDER_STYLES`；0 表示无边框 */
    borderTop: number;
    borderBottom: number;
    borderLeft: number;
    borderRight: number;
    /** 边框颜色索引（**四边各自独立**，不再合并成一个值） */
    borderColorTop: number;
    borderColorBottom: number;
    borderColorLeft: number;
    borderColorRight: number;
    /** 填充图案索引（进 `BIFF_FILL_PATTERNS`） */
    fillPattern: number;
    /**
     * 填充**前景**色索引（即 Excel UI 里的「图案颜色」，solid 填充时它就是底色）
     *
     * ⚠️ BIFF 的命名与直觉相反：solid 填充只有 `fillColorIndex` 有意义，
     * 而 Excel 的「背景色」下拉写的就是它。
     */
    fillColorIndex: number;
    /** 填充**背景**色索引（图案的底纹色，`patternType` 非 solid 时才有视觉意义） */
    fillBackgroundIndex: number;
}

/**
 * BIFF 单元格信息
 */
export interface BiffCell {
    /** 行号（0-based） */
    row: number;
    /** 列号（0-based） */
    col: number;
    /** 样式索引（指向 XF 表） */
    styleIndex: number;
    /** 值类型 */
    type: 'string' | 'number' | 'boolean' | 'error' | 'blank';
    /** 数值（type 为 number 时有效） */
    numberValue?: number;
    /** 字符串（type 为 string 时有效） */
    stringValue?: string;
    /** 布尔值 */
    booleanValue?: boolean;
    /** 错误码字符串 */
    errorValue?: string;
}

/**
 * BIFF 工作表信息
 */
export interface BiffSheet {
    /** 工作表名称 */
    name: string;
    /** 在 Workbook 流中的起始偏移 */
    position: number;
    /** 是否隐藏 */
    hidden: boolean;
    /** 是否为图表工作表 */
    isChart: boolean;
}

/**
 * BIFF 图片信息
 */
export interface BiffImage {
    /** 图片二进制数据 */
    data: Uint8Array;
    /** MIME 类型 */
    mimeType: string;
    /** 锚点（0-based 行列），若无法确定则为 undefined */
    anchor?: { col: number, row: number, colOffset: number, rowOffset: number };
}

/**
 * BIFF 解析结果
 */
export interface BiffWorkbook {
    sheets: BiffSheet[];
    /**
     * 字体表，**下标严格等于 BIFF8 的 FontIndex**（MS-XLS 2.5.129）
     *
     * 与「第几条 FONT 记录」并不相同：FontIndex < 4 时是 0 起算的记录序号，
     * FontIndex > 4 时是 **1 起算** 的记录序号，且 4 永不出现。
     * 因此下标 4 恒为空洞（`undefined`）。详见 {@link BiffParser.parseFont}。
     */
    fonts: Array<BiffFont | undefined>;
    formats: Map<number, string>;
    xfs: BiffXf[];
    palette: string[];
    /** 共享字符串表 */
    sharedStrings: string[];
    /** Excel 内置的错误值表 */
    errorValues: string[];
}

/**
 * BIFF8 内置调色板（文件里没有 PALETTE 记录时的兜底）
 *
 * ⚠️ **索引有固定语义，不能从 0 开始平铺**（历史实现就是这么错的：64 个颜色
 * 从下标 0 挨着放，于是标准 56 色整体前移了 58 位，所有颜色索引都取错色）：
 *
 * | 索引 | 含义 |
 * |---|---|
 * | 0-7 | 保留（未定义）；0 在代码中已被 `paletteColor()` 当作「无颜色」 |
 * | 8-63 | 系统色：8=黑 9=白 10=红 11=绿 12=蓝 13=黄 14=品红 15=青 … |
 * | 64 | 系统前景色（黑） |
 * | 65 | 系统背景色（白） |
 * | 66-121 | 标准 56 色（与 8-63 是**同一组颜色**，Excel 的默认 palette 这么排） |
 *
 * 实测夹具 `test_sample.xls` 无 PALETTE 记录，却大量使用索引 13（黄）、23（灰）、
 * 53（橙）等，所以这张表必须准确，否则底色/字色/边框色全错。
 */
export const DEFAULT_PALETTE: string[] = (() => {
    const palette: string[] = [];
    // 0-7：保留位（占位，保持下标与索引对齐）
    for (let i = 0; i < 8; i++) palette.push('#000000');
    // 8-63：系统色（56 项）
    const systemColors = [
        '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
        '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#C0C0C0', '#808080',
        '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF',
        '#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000', '#008080', '#0000FF',
        '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF', '#FF99CC', '#CC99FF', '#FFCC99',
        '#3366FF', '#33CCCC', '#99CC00', '#FFCC00', '#FF9900', '#FF6600', '#666699', '#969696',
        '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333'
    ];
    for (const c of systemColors) palette.push(c);
    // 64：系统前景色（黑）；65：系统背景色（白）
    palette.push('#000000');
    palette.push('#FFFFFF');
    // 66-121：标准 56 色（与 8-63 同组）
    for (const c of systemColors) palette.push(c);
    return palette;
})();

/** BIFF 错误码表（ERR7 记录的索引） */
const ERROR_VALUES: string[] = [
    '#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA'
];

/**
 * 按 BIFF8 FontIndex 取字体（MS-XLS 2.5.129）
 *
 * {@link BiffWorkbook.fonts} 的下标已按 FontIndex 语义对齐（下标 4 是空洞），
 * 所以这里就是一次直查；抽成函数是为了让「取值集合 = {0,1,2,3}∪{5,6,…}」
 * 这条规则只有一处实现、便于测试与复用（SST 富文本 run、TXO run 共用同一套索引）。
 *
 * 索引越界或落到空洞（`undefined`）时回退到默认字体（下标 0），
 * 与 Excel 遇到非法字体索引时的行为一致。
 *
 * @param fonts - 字体表（下标 == FontIndex）
 * @param fontIndex - XF / 富文本 run 里的 ifnt
 * @returns 命中的字体；表为空时返回 undefined
 */
export function resolveFont(
    fonts: Array<BiffFont | undefined>,
    fontIndex: number
): BiffFont | undefined {
    // 规范保证 ifnt 永不等于 4；这里不做特判，靠「下标 4 是空洞」自然回退。
    return fonts[fontIndex] ?? fonts[0];
}

/**
 * BIFF8 解析器
 *
 * @example
 * const parser = new BiffParser();
 * const wb = parser.parse(workbookStreamBytes);
 * console.log(wb.sheets.map(s => s.name));
 */
export class BiffParser {

    /** Workbook 流数据 */
    private data: Uint8Array;
    /** 小端 DataView */
    private view: DataView;

    /** 解析结果 */
    public workbook: BiffWorkbook;
    /** SST 记录头偏移（供第一轮扫描定位 CONTINUE 链） */
    private sstHeaderPos: number = -1;
    /**
     * SST 主记录的**数据**起始偏移（= 记录头 + 4）
     *
     * 段索引 0 对应的数据不在独立数组里，而是 `this.data` 的
     * `[sstDataStart, sstDataEnd)` 区间；{@link readSstString} 需要它
     * 才能把「段内偏移」映射回绝对偏移。
     */
    private sstDataStart: number = 0;
    /**
     * 是否解析内嵌图片（Escher BLIP）
     *
     * 默认 false：`xl`/`xls` 的图片读取在本项目里没有消费方（见
     * `SheetParserOptions.readImages`），而 .xls 的图片字节是逐张
     * **复制**出来的（{@link parseBlip} 里的 `new Uint8Array(...)`）。
     * 一个含上千张图的 .xls 会白复制上百 MB 字节并常驻内存 —— 打开即卡死。
     * 只有调用方显式要求时才置 true。
     */
    private readImages: boolean = false;

    /**
     * @param {Uint8Array} data - Workbook 流的完整字节数据
     * @param {object} [options] - 解析选项
     * @param {boolean} [options.readImages=false] - 是否解析内嵌图片（默认 false，见 {@link readImages}）
     */
    constructor(data: Uint8Array, options: { readImages?: boolean } = {}) {
        this.data = data;
        this.view = new DataView(data.buffer, data.byteOffset, data.byteLength);
        this.readImages = options.readImages === true;
        this.workbook = {
            sheets: [],
            // 字体表**按下标 == FontIndex 语义**填充，见 parseFont：
            // 前 4 条 FONT 记录落到下标 0-3；下标 4 是规范禁止引用的空洞，
            // 第 5 条记录落到下标 5。
            fonts: [],
            formats: new Map<number, string>(),
            xfs: [],
            palette: DEFAULT_PALETTE.slice(),
            sharedStrings: [],
            errorValues: ERROR_VALUES.slice()
        };
    }

    /**
     * 解析 Workbook 流
     *
     * @returns {BiffWorkbook} 全局解析结果（工作表位置、样式表、共享字符串等）
     * @throws {Error} 流格式非法
     */
    public parse(): BiffWorkbook {
        this.readGlobalRecords();
        return this.workbook;
    }

    /**
     * 读取一个 BIFF 记录的元信息
     *
     * @param {number} pos - 记录起始偏移
     * @returns {{ type: number, length: number, dataStart: number }} 记录信息
     * @private
     */
    private readRecordHeader(pos: number): { type: number, length: number, dataStart: number } {
        if (pos + 4 > this.data.length) {
            throw new Error('BIFF 记录越界于偏移 ' + pos);
        }
        const type = this.view.getUint16(pos, true);
        const length = this.view.getUint16(pos + 2, true);
        return { type, length, dataStart: pos + 4 };
    }

    /**
     * 第一轮扫描：读取全局记录
     *
     * 只在「工作簿全局子流」范围内扫描（即第一个 BOF 到其对应 EOF），
     * 避免把工作表子流中的同类记录重复计入。
     *
     * @private
     */
    private readGlobalRecords(): void {
        // 每条记录额外保存其在流中的绝对位置 pos，供 SST 定位 CONTINUE 链使用
        const globals: { type: number, length: number, dataStart: number, pos: number }[] = [];

        let pos = 0;
        // 跳过开头的 BOF
        if (pos + 4 <= this.data.length && this.view.getUint16(pos, true) === REC.BOF) {
            const bof = this.readRecordHeader(pos);
            pos = bof.dataStart + bof.length;
        }

        let nesting = 1;  // 子流嵌套深度

        while (pos + 4 <= this.data.length && nesting > 0) {
            const rec = this.readRecordHeader(pos);
            if (rec.dataStart + rec.length > this.data.length) break;

            if (rec.type === REC.EOF) {
                nesting--;
                if (nesting === 0) break;
            } else if (rec.type === REC.BOF) {
                // 工作表子流的开始，跳过该子流直到其 EOF
                const subStart = rec.dataStart + rec.length;
                pos = this.skipSubstream(subStart);
                continue;
            } else {
                globals.push({
                    type: rec.type,
                    length: rec.length,
                    dataStart: rec.dataStart,
                    pos
                });
            }

            pos = rec.dataStart + rec.length;
        }

        // 逐个处理收集到的记录
        // 注意 SST 需在最后处理——它依赖 FONT/FORMAT/XF 等记录已就位，
        // 且 CONTINUE 链的扫描会跳过中间记录
        for (const rec of globals) {
            if (rec.type === REC.SST) {
                this.sstHeaderPos = rec.pos;
                continue;
            }
            switch (rec.type) {
                case REC.BOUNDSHEET:
                    this.parseBoundsheet(rec.dataStart, rec.length);
                    break;
                case REC.FONT:
                    this.parseFont(rec.dataStart, rec.length);
                    break;
                case REC.FORMAT:
                case REC.FORMAT2:
                    this.parseFormat(rec.dataStart, rec.length);
                    break;
                case REC.XF:
                    this.parseXf(rec.dataStart, rec.length);
                    break;
                case REC.PALETTE:
                    this.parsePalette(rec.dataStart, rec.length);
                    break;
                case REC.DEFAULTROWHEIGHT:
                case REC.DEFCOLWIDTH:
                    // 默认行高/列宽，工作表级的 COLINFO/ROW 会覆盖
                    break;
            }
        }

        // SST 最后解析（ CONTINUE 记录已在 globals 收集阶段被排除）
        if (this.sstHeaderPos >= 0) {
            const sstRec = globals.find(r => r.type === REC.SST)!;
            this.parseSst(this.sstHeaderPos, sstRec.length);
        }

        // 规范（MS-XLS 2.5.129）要求前 4 条 FONT 记录必定存在且依次是
        // 「默认字体 / 粗体 / 斜体 / 粗斜体」。少数第三方写入器只写其中一部分，
        // 此时按该语义把缺失的下标补出来，否则 XF 用 ifnt 0-3 会取不到字体。
        this.ensureDefaultFontSlots();
    }

    /**
     * 补齐前 4 个字体槽位（下标 0-3）
     *
     * 只在文件缺记录时补，语义按 MS-XLS 2.5.129 的对照表：
     * `0` 默认字体、`1` 默认字体加粗、`2` 默认字体倾斜、`3` 默认字体加粗倾斜。
     *
     * @private
     */
    private ensureDefaultFontSlots(): void {
        const fonts = this.workbook.fonts;
        const base = fonts[0];
        if (!base) return;
        for (let i = 0; i < 4; i++) {
            if (fonts[i]) continue;
            fonts[i] = {
                ...base,
                bold: i === 1 || i === 3 || base.bold,
                italic: i === 2 || i === 3 || base.italic
            };
        }
    }

    /**
     * 跳过一个子流（从指定偏移扫描到匹配的 EOF）
     * @param start - 子流起始偏移
     * @returns 子流结束后的偏移
     * @private
     */
    private skipSubstream(start: number): number {
        let pos = start;
        let depth = 1;
        while (pos + 4 <= this.data.length && depth > 0) {
            const rec = this.readRecordHeader(pos);
            if (rec.dataStart + rec.length > this.data.length) break;
            if (rec.type === REC.BOF) depth++;
            if (rec.type === REC.EOF) depth--;
            pos = rec.dataStart + rec.length;
        }
        return pos;
    }

    /**
     * 解析 BOUNDSHEET 记录（工作表位置与名称）
     *
     * 结构：
     * - 0：子流在 Workbook 流中的绝对偏移(u32)
     * - 4：可见性与类型标志(u16)，bit0=隐藏，bit2=图表工作表
     * - 6：名称长度(u8)
     * - 7：名称标志(u8)，bit0=16 位字符
     * - 8：名称数据
     *
     * @param offset - 记录数据起始偏移
     * @param length - 记录长度
     * @private
     */
    private parseBoundsheet(offset: number, length: number): void {
        if (length < 8) return;

        const position = this.view.getUint32(offset, true);
        const flags = this.view.getUint16(offset + 4, true);
        const nameLength = this.data[offset + 6];
        const nameFlags = this.data[offset + 7];

        const name = this.readUnicodeString(offset + 8, nameLength, (nameFlags & 0x01) !== 0);

        this.workbook.sheets.push({
            name,
            position,
            hidden: (flags & 0x01) !== 0,
            isChart: (flags & 0x02) !== 0
        });
    }

    /**
     * 解析 FONT 记录
     *
     * BIFF8 结构：
     * -  0：字高（twips，1 磅 = 20 twips）(u16)
     * -  2：标志(u16)：bit0 粗体，bit1 斜体，bit2 下划线，bit3 删除线，
     *   bit4 上标，bit5 下标
     * -  4：颜色索引(u16)
     * -  6：字重(u16)，400 = 常规，700 = 粗体
     * -  8：上下标(u16)
     * - 10：下划线类型(u16)
     * - 12：字体族(u8)
     * - 13：字符集(u8)
     * - 14：名称字符数(u8)
     * - 15：名称标志(u8)，bit0 = 1 表示 UTF-16LE
     * - 16：名称数据
     *
     * @param offset - 记录数据起始偏移
     * @param length - 记录长度
     * @private
     */
    /**
     * 解析 FONT 记录
     *
     * BIFF8 布局（20 字节定长头 + 变长名称）：
     *
     * | 偏移 | 大小 | 内容 |
     * |---|---|---|
     * | 0 | 2 | 字体高度（twips，1 磅 = 20 twips） |
     * | 2 | 2 | 选项位：bit0 粗体(冗余) bit1 斜体 bit2 下划线(冗余) bit3 删除线 bit4 轮廓 bit5 阴影 |
     * | 4 | 2 | 颜色索引（**0x7FFF 表示「自动」**，应映射为默认前景色而非某个调色板色） |
     * | 6 | 2 | 字重 100-1000（400=常规 700=粗体） |
     * | 8 | 2 | escapement：0=无 1=上标 2=下标 |
     * | 10 | 1 | 下划线类型：0=无 1=单 2=双 0x21=单会计 0x22=双会计 |
     * | 11 | 1 | 字体族 |
     * | 12 | 1 | 字符集 |
     * | 13 | 1 | 未使用 |
     * | 14 | 1 | 名称字符数（8 位） |
     * | 15 | 1 | 名称选项：bit0=1 表示 UTF-16LE，否则压缩单字节 |
     * | 16 | var | 名称字符数据 |
     *
     * ⚠️ 历史实现的三处错误：
     * ① 下划线取自 `flags & 0x0004`（options 里的冗余位）——规范明确说「读取字体时
     *    不使用该位」，应取偏移 10 的 underline type；
     * ② 上下标取自 `flags` 的 bit4/bit5（那是轮廓/阴影）——应取偏移 8 的 escapement；
     * ③ 未处理颜色索引 0x7FFF（自动色），会去查一个不存在的调色板项。
     *
     * **字体索引（本方法负责维护数组下标）** —— MS-XLS 2.5.129 FontIndex：
     *
     * > If this value is less than 4, then it specifies a **zero-based** index of a Font
     * > record …; if this value is greater than 4, then it specifies a **one-based** index
     * > …; MUST NOT equal 4.
     *
     * 即取值集合是 `{0,1,2,3} ∪ {5,6,7,…}` —— **下标 4 被永久跳过**，
     * 第 5 条 FONT 记录对应 ifnt 5（不是 4）。所以填充字体表时要在下标 4 留一个空洞，
     * 数组下标才与 FontIndex 一一对应（`wb.fonts[ifnt]` 直接可用）。
     *
     * 历史实现把整个表按 `ifnt - 4` 对齐（前 4 格占位、记录从下标 4 开始），
     * 于是**每个字体都错位 3 格** —— 症状正是「字号/加粗/颜色与 Excel 对不上」：
     * 例如夹具里 XF 用 ifnt=7（应为第 6 条记录：Arial 36pt）会被取成第 3 条
     * （宋体 11pt）。现按规范对齐。
     *
     * @param offset - 记录数据起始偏移
     * @param length - 记录长度
     * @private
     */
    private parseFont(offset: number, length: number): void {
        if (length < 16) return;

        // 前 4 条记录占下标 0-3；第 5 条记录起 ifnt 是「1 起算」的，
        // 因此在准备写入第 5 条（此时数组长度恰为 4）时把下标 4 挖空。
        if (this.workbook.fonts.length === 4) {
            this.workbook.fonts.length = 5;
        }

        const heightTwips = this.view.getUint16(offset, true);
        const flags = this.view.getUint16(offset + 2, true);
        const colorIndex = this.view.getUint16(offset + 4, true);
        const boldWeight = this.view.getUint16(offset + 6, true);
        const escapement = this.view.getUint16(offset + 8, true);
        // 下划线类型在偏移 10（单字节），不是 options 的 bit2
        const underlineType = this.data[offset + 10];

        // 名称：无名称时保持默认（旧版 Excel 默认字体）
        let name = 'Calibri';
        if (length > 16) {
            const nameLength = this.data[offset + 14];
            const nameFlags = this.data[offset + 15];
            if (nameLength > 0 && 16 + nameLength <= length) {
                name = this.readUnicodeString(offset + 16, nameLength, (nameFlags & 0x01) !== 0);
            }
        }

        this.workbook.fonts.push({
            name,
            // twips → 磅（1 磅 = 20 twips）
            height: Math.round((heightTwips / 20) * 10) / 10,
            // 粗体判定：字重为主（>= 600 视为粗体），options 的 bit0 作为兜底
            bold: (boldWeight >= 600 && boldWeight !== 0) || (flags & 0x0001) !== 0,
            italic: (flags & 0x0002) !== 0,
            // 下划线取自偏移 10 的 underline type（0 为无，其余均为有）
            underline: underlineType !== 0,
            strikethrough: (flags & 0x0008) !== 0,
            superscript: escapement === 1,
            subscript: escapement === 2,
            // 0x7FFF = 自动色：交给下游按「无显式颜色」处理
            colorIndex: colorIndex === 0x7FFF ? 0 : colorIndex
        });
    }

    /**
     * 解析 FORMAT 记录（自定义数字格式）
     *
     * BIFF8 结构：
     * - 0：格式索引(u16)
     * - 2：格式码字符数(u16)  ← 注意是 2 字节，与 LABELSST 的 SST 字符串不同
     * - 4：格式码标志(u8)，bit0 = 1 表示 UTF-16LE
     * - 5：格式码字符串数据
     *
     * @param offset - 记录数据起始偏移
     * @param length - 记录长度
     * @private
     */
    private parseFormat(offset: number, length: number): void {
        if (length < 6) return;

        const formatIndex = this.view.getUint16(offset, true);
        const strLength = this.view.getUint16(offset + 2, true);
        const strFlags = this.data[offset + 4];

        const code = this.readUnicodeString(offset + 5, strLength, (strFlags & 0x01) !== 0);
        if (code) {
            this.workbook.formats.set(formatIndex, code);
        }
    }

    /**
     * 解析 XF（扩展格式）记录
     *
     * BIFF8 的 XF 记录**固定 20 字节**，权威布局如下（改前务必对照，历史实现
     * 把边框/填充字段的偏移几乎全读错了，症状是「边框、底色、文字颜色、字号
     * 解析不准确」）：
     *
     * | 偏移 | 大小 | 内容 |
     * |---|---|---|
     * | 0 | 2 | FONT 索引 |
     * | 2 | 2 | FORMAT 索引 |
     * | 4 | 2 | bit2-0 保护/类型（bit2=1 表示 style XF）；bit15-4 父样式 XF 索引 |
     * | 6 | 1 | bit2-0 水平对齐；bit3 右边界换行；bit6-4 **垂直对齐** |
     * | 7 | 1 | 旋转角度 |
     * | 8 | 1 | bit3-0 缩进；bit4 缩小填充；bit7-6 文本方向 |
     * | 9 | 1 | XF_USED_ATTRIB（bit0 数字格式 bit1 字体 bit2 对齐 bit3 边框 bit4 底色 bit5 保护） |
     * | 10 | 4 | **边框线型 + 左右边框色**（见下） |
     * | 14 | 4 | **上/下边框色 + 对角线 + 填充图案**（见下） |
     * | 18 | 2 | **填充前景色 + 填充背景色**（见下） |
     *
     * 偏移 10 的 u32（小端）位分配：
     * - bit3-0 左边框线型、bit7-4 右边框线型、bit11-8 上边框线型、bit15-12 下边框线型
     * - bit22-16 左边框色索引、**bit29-23 右边框色索引**
     * - bit30/31 两条对角线标志
     *
     * 偏移 14 的 u32（小端）位分配：
     * - bit6-0 上边框色、bit13-7 下边框色、bit20-14 对角线色
     * - bit24-21 对角线线型、**bit31-26 填充图案**
     *
     * 偏移 18 的 u16（小端）位分配：
     * - bit6-0 填充**前景**色（solid 时即底色）、bit13-7 填充**背景**色
     *
     * @param offset - 记录数据起始偏移
     * @param length - 记录长度
     * @private
     */
    private parseXf(offset: number, length: number): void {
        if (length < 20) return;

        const fontIndex = this.view.getUint16(offset, true);
        const formatIndex = this.view.getUint16(offset + 2, true);
        const typeProt = this.view.getUint16(offset + 4, true);
        const alignByte = this.data[offset + 6];
        const rotationByte = this.data[offset + 7];
        const indentByte = this.data[offset + 8];
        // XF_USED_ATTRIB（偏移 9）：bit0=fAtrNum bit1=fAtrFnt bit2=fAtrAlc
        // bit3=fAtrBdr bit4=fAtrPat …，语义是「该属性是否有效」。
        // **刻意不用它做门控**：两份夹具实测都从不置起 bit0/bit1 ——
        // test/_ref_xlwt.xls 只出现 0xF4 / 0xF8，test/test_sample.xls 出现
        // 0x00/0x94/0x9C/0xB4/0xD4/0xF4/0xF8/0xFC，**全部** bit0=bit1=0。
        // 按它过滤会把字体和数字格式整片丢掉，而这两项在 XF 里本就是实值。
        // 单元格 XF（fStyle=0）的各字段直接采信即可。
        const usedByte = this.data[offset + 9];

        // ── 边框线型与「左/右」边框色：全部位于偏移 10 的 u32 内 ──
        const borderMain = this.view.getUint32(offset + 10, true);
        const lineLeft = borderMain & 0x0F;              // bit3-0
        const lineRight = (borderMain >>> 4) & 0x0F;     // bit7-4
        const lineTop = (borderMain >>> 8) & 0x0F;       // bit11-8
        const lineBottom = (borderMain >>> 12) & 0x0F;   // bit15-12
        const colorLeft = (borderMain >>> 16) & 0x7F;    // bit22-16
        const colorRight = (borderMain >>> 23) & 0x7F;   // bit29-23

        // ── 上/下边框色 + 填充图案：位于偏移 14 的 u32 内 ──
        const borderExtra = this.view.getUint32(offset + 14, true);
        const colorTop = borderExtra & 0x7F;             // bit6-0
        const colorBottom = (borderExtra >>> 7) & 0x7F;  // bit13-7
        const fillPattern = (borderExtra >>> 26) & 0x3F; // bit31-26

        // ── 填充前景/背景色：位于偏移 18 的 u16 内 ──
        const fillColors = this.view.getUint16(offset + 18, true);
        const fillColorIndex = fillColors & 0x7F;         // bit6-0（solid 时为底色）
        const fillBackgroundIndex = (fillColors >>> 7) & 0x7F; // bit13-7

        // ⚠️ isStyle 只由 offset+4 的 bit2（0x0004）决定。
        // 历史写法 `(flags & 0xFFF0) === 0x0000 || ...` 是恒真/恒假的无效表达式。
        const isStyle = (typeProt & 0x0004) !== 0;
        const parentXfIndex = (typeProt >>> 4) & 0x0FFF;

        this.workbook.xfs.push({
            fontIndex,
            formatIndex,
            horizontalAlignment: alignByte & 0x07,
            // ⚠️ 垂直对齐是 bit6-4（**3 位**），历史实现写成 bit3-4 只取了 1 位
            verticalAlignment: (alignByte >>> 4) & 0x07,
            // 对齐字节 @6 的位分配（MS-XLS 2.4.353 / OpenOffice BIFF8 §5.5）：
            //   bit0-2 = alc（水平）  bit3 = fWrap（自动换行）  bit4-6 = alcV（垂直）
            // ⚠️ alcV 是 **3 位**（历史实现写成 bit3-4 只取了 1 位，1 与 4 都读成 0）。
            // fWrap 由 test/_ref_xlwt.xls 的 `al.wrap = 1` 实证：该 XF 读出 true。
            wrapText: (alignByte & 0x08) !== 0,
            rotation: rotationByte,
            indent: indentByte & 0x0F,
            shrinkToFit: (indentByte & 0x10) !== 0,
            isStyle,
            parentXfIndex,
            borderTop: lineTop,
            borderBottom: lineBottom,
            borderLeft: lineLeft,
            borderRight: lineRight,
            borderColorTop: colorTop,
            borderColorBottom: colorBottom,
            borderColorLeft: colorLeft,
            borderColorRight: colorRight,
            fillPattern,
            fillColorIndex,
            fillBackgroundIndex
        });
    }

    /**
     * 解析 PALETTE 记录（自定义调色板）
     *
     * - 0：颜色个数(u16)
     * - 2：起始索引(u16)
     * - 4：颜色数据，每色 4 字节（RGB 3 字节 + 标志 1 字节）
     *
     * @param offset - 记录数据起始偏移
     * @param length - 记录长度
     * @private
     */
    private parsePalette(offset: number, length: number): void {
        if (length < 4) return;

        const count = this.view.getUint16(offset, true);
        const startIndex = this.view.getUint16(offset + 2, true);

        for (let i = 0; i < count; i++) {
            const pos = offset + 4 + i * 4;
            if (pos + 3 > this.data.length) break;
            // ⚠️ BIFF8 调色板每项 4 字节，顺序是 **R, G, B, 0x00**（红在前）。
            // 别按「BGR」读——那会把红蓝对调（历史注释曾这么写，属误导）。
            const r = this.data[pos];
            const g = this.data[pos + 1];
            const b = this.data[pos + 2];
            const hex = rgbToHex(r, g, b);
            const targetIndex = startIndex + i;
            // 扩展调色板数组以容纳索引
            while (this.workbook.palette.length <= targetIndex) {
                this.workbook.palette.push('#000000');
            }
            this.workbook.palette[targetIndex] = hex;
        }
    }

    /**
     * 解析 SST（共享字符串表）及其 CONTINUE 记录
     *
     * 这是 xls 解析中最复杂的部分。SST 记录的结构：
     * - 0：字符串总数(u32)
     * - 4：唯一字符串数(u32)
     * - 8：字符串数据序列
     *
     * 每个字符串：长度(u16) + 标志(u8，bit0=1 表示 UTF-16LE) + 字符数据
     * 若含富文本运行(rich text)则字符串头后还有 2 字节运行数，
     * 若含亚洲 phonetic 数据则再加 4 字节长度。
     *
     * **跨 CONTINUE 边界的处理**：
     * 当一个字符串正好在 CONTINUE 边界处被切断时，新记录开头会插入
     * 一个新的 grbit 字节，且字符串长度不再重复读取——
     * 续接部分直接延续前面的字符数据。
     *
     * @param pos - SST 记录头偏移（用于定位 CONTINUE 链）
     * @param length - SST 记录长度
     * @private
     */
    private parseSst(pos: number, length: number): void {
        const start = pos + 4;   // 跳过记录头
        // ⚠️ end 必须由 start 推出：length 是**数据区**长度（不含 4 字节记录头），
        // 写成 pos + length 会少 4 字节，把 SST 尾部截断。
        const end = start + length;

        if (start + 8 > this.data.length) return;
        const uniqueCount = this.view.getUint32(start + 4, true);

        // 段索引 0 = 本记录的数据区间 [start, end)，供 readSstString 定位
        this.sstDataStart = start;

        // 收集 SST 之后的所有 CONTINUE 记录
        const continuations: Uint8Array[] = [];
        let scanPos = end;
        while (scanPos + 4 <= this.data.length) {
            const rec = this.readRecordHeader(scanPos);
            if (rec.type !== REC.CONTINUE) break;
            if (rec.dataStart + rec.length > this.data.length) break;
            continuations.push(this.data.subarray(rec.dataStart, rec.dataStart + rec.length));
            scanPos = rec.dataStart + rec.length;
        }

        // 逐个读取字符串，段间可跨越 CONTINUE 边界
        // 注意起始位置是 start + 8：数据区前 8 字节是 cstTotal + cstUnique 两个 u32
        let segIndex = 0;
        let segPos = start + 8;
        const strings: string[] = [];

        for (let i = 0; i < uniqueCount; i++) {
            const result = this.readSstString(continuations, segIndex, segPos, end);
            if (!result) break;
            strings.push(result.text);
            segIndex = result.nextSegIndex;
            segPos = result.nextSegPos;
        }

        this.workbook.sharedStrings = strings;
    }

    /**
     * 从 SST 数据流中读取一个字符串（支持跨 CONTINUE 边界）
     *
     * ⚠️ **这一段的坑最多，改动前务必读懂下面两条规则**（历史上两条全踩，
     * 症状是整表中文全部变成 `\u0000\u0007...` 之类的乱码）：
     *
     * 1. **不能按「逐字节」读，必须按「段内快照」读。**
     *    长度(u16)/grbit(u8)/字符数据都只能落在**同一个段**内，不允许被拆散
     *    （BIFF 规范明文：字符串的 cch 与 grbit 不会被切到两个记录里）。
     *    若写成「逐字节拼接」，段边界处会把 CONTINUE 记录头误当成数据，
     *    于是 cch/grbit 整体错位 —— 这就是乱码的直接成因。
     *
     * 2. **只有「字符数据」允许跨段，且跨段后必须重读一个新的 grbit。**
     *    续段开头的那个字节不再是长度，而是新的 grbit（可能从 UTF-16
     *    切换成压缩 8 位，或反之）。旧实现完全漏了这一步。
     *
     * 判定「是否需要跨段」的正确做法：只有当**当前段剩余字节数**
     * 不足以容纳剩余字符所需字节数时，才跨到下一段。
     *
     * @param continuations - CONTINUE 记录的数据块数组
     * @param segIndex - 当前所在段索引（0 表示 SST 主记录，其余为 CONTINUE）
     * @param segPos - 段内偏移
     * @param end - SST 主记录的数据结束偏移
     * @returns 字符串内容与新的位置；解析失败返回 null
     * @private
     */
    private readSstString(
        continuations: Uint8Array[],
        segIndex: number,
        segPos: number,
        end: number
    ): { text: string, nextSegIndex: number, nextSegPos: number } | null {

        // ⚠️ curPos 是「**段内**偏移」：段 0 的段内偏移 = 绝对偏移 − sstDataStart。
        // 调用方传入的 segPos 对段 0 是绝对偏移，这里统一折算，避免两套口径混用。
        let curSeg = segIndex;
        let curPos = segIndex === 0 ? segPos - this.sstDataStart : segPos;

        /** idx 0 = SST 主记录，其数据是 this.data 的 [sstDataStart, end) 区间 */
        const segLength = (idx: number): number =>
            idx === 0 ? end - this.sstDataStart
                : (continuations[idx - 1] ? continuations[idx - 1].length : 0);

        /** 从当前段读一个字节（不跨段）；段耗尽返回 -1 */
        const peekSegByte = (): number => {
            if (curPos >= segLength(curSeg)) return -1;
            if (curSeg === 0) return this.data[this.sstDataStart + curPos];
            const seg = continuations[curSeg - 1];
            return seg ? seg[curPos] : -1;
        };

        /** 读一个 u16（必须在同一段内，不跨段） */
        const readSegUint16 = (): number => {
            const b0 = peekSegByte();
            if (b0 === -1) return -1;
            const len = segLength(curSeg);
            if (curPos + 1 >= len) return -1;
            curPos++;
            const b1 = peekSegByte();
            curPos++;
            return b0 | (b1 << 8);
        };

        // —— 字符串长度(u16) —— 必须完整落在同一段内
        const charCount = readSegUint16();
        if (charCount < 0) return null;

        // —— 标志(u8) —— 也必须落在同一段内
        //
        // grbit 位定义（BIFF8 SST）：
        // - bit0 (0x01) fHighByte：1 = UTF-16LE，0 = 压缩单字节
        // - bit2 (0x04) fExtSt  ：1 = 字符串头含 cbExtRst（phonetic 扩展）
        // - bit3 (0x08) fRichSt ：1 = 字符串头含 cRun（富文本运行数）
        //
        // ⚠️ cRun / cbExtRst **只在对应标志置位时才存在**，不是「UTF-16 就有」。
        // 旧实现无条件读取这两个字段，把字符数据的头两个汉字当成了 cRun，
        // 于是字符串整体错位、整表中文全变乱码。
        let flagByte = peekSegByte();
        if (flagByte === -1) return null;
        curPos++;
        let isUnicode = (flagByte & 0x01) !== 0;
        const hasExt = (flagByte & 0x04) !== 0;
        const hasRich = (flagByte & 0x08) !== 0;

        // —— 富文本运行数 cRun(u16)（仅 fRichSt 置位时存在）——
        if (hasRich) {
            const cRun = readSegUint16();
            if (cRun > 0) {
                // 跳过 cRun 个 4 字节的格式运行（每项 4 字节：ichFont + ifnt）
                curPos += cRun * 4;
            }
        }

        // —— phonetic 扩展长度 cbExtRst(i32)（仅 fExtSt 置位时存在）——
        // 必须用**局部变量**记录：readSstString 是被 parseSst 循环调用的，
        // 写成实例字段会让「上一个字符串的 cbExtRst」泄漏到下一个字符串上。
        let phoneticSkip = 0;
        if (hasExt) {
            const cbExtRstLow = readSegUint16();
            const cbExtRstHigh = readSegUint16();
            if (cbExtRstLow >= 0 && cbExtRstHigh >= 0) {
                // phonetic 字节体在**字符数据之后**，先记下长度，读完字符再跳
                phoneticSkip = cbExtRstLow | (cbExtRstHigh << 16);
            }
        }

        // —— 字符数据（唯一允许跨段的部分）——
        let remaining = charCount;
        let result = '';

        while (remaining > 0) {
            const bytesPerChar = isUnicode ? 2 : 1;
            const avail = segLength(curSeg) - curPos;
            const need = remaining * bytesPerChar;

            // 当前段能放下全部剩余字符 → 直接在本段解完
            if (avail >= need) {
                if (isUnicode) {
                    for (let i = 0; i < remaining; i++) {
                        const b0 = peekSegByte(); curPos++;
                        const b1 = peekSegByte(); curPos++;
                        result += String.fromCharCode(b0 | (b1 << 8));
                    }
                } else {
                    const buf: number[] = [];
                    for (let i = 0; i < remaining; i++) { buf.push(peekSegByte()); curPos++; }
                    result += decodeSingleByte(buf);
                }
                break;
            }

            // 当前段装不下 → 本段能解多少解多少，然后跨段
            const fitChars = Math.floor(avail / bytesPerChar);
            if (fitChars > 0) {
                if (isUnicode) {
                    for (let i = 0; i < fitChars; i++) {
                        const b0 = peekSegByte(); curPos++;
                        const b1 = peekSegByte(); curPos++;
                        result += String.fromCharCode(b0 | (b1 << 8));
                    }
                } else {
                    const buf: number[] = [];
                    for (let i = 0; i < fitChars; i++) { buf.push(peekSegByte()); curPos++; }
                    result += decodeSingleByte(buf);
                }
                remaining -= fitChars;
            }

            // 切到下一个 CONTINUE
            curSeg++;
            curPos = 0;
            if (curSeg - 1 >= continuations.length) return null;
            // 跨段后：段首是一个**新的 grbit**（旧实现漏掉这一步 → 乱码根源）
            const newFlag = peekSegByte();
            if (newFlag === -1) return null;
            curPos++;
            isUnicode = (newFlag & 0x01) !== 0;
        }

        // —— 跳过 phonetic 数据（字节体位于字符数据之后，可能跨段）——
        let skip = phoneticSkip;
        while (skip > 0) {
            const avail = segLength(curSeg) - curPos;
            if (avail <= 0) {
                curSeg++;
                curPos = 0;
                if (curSeg - 1 >= continuations.length) break;
                continue;
            }
            const step = Math.min(avail, skip);
            curPos += step;
            skip -= step;
        }

        // 返回值统一口径：段 0 用**绝对偏移**（与调用方一致），其余段用段内偏移
        return {
            text: result,
            nextSegIndex: curSeg,
            nextSegPos: curSeg === 0 ? this.sstDataStart + curPos : curPos
        };
    }

    /**
     * 读取 BIFF 的 Unicode 字符串（长度已知的，非 SST 场景）
     *
     * @param offset - 字符串数据偏移
     * @param charCount - 字符数
     * @param isUnicode - 是否 UTF-16LE
     * @returns 解码后的字符串
     * @private
     */
    private readUnicodeString(offset: number, charCount: number, isUnicode: boolean): string {
        if (charCount <= 0 || offset + charCount > this.data.length) return '';

        if (isUnicode) {
            let result = '';
            const end = offset + charCount * 2;
            for (let i = offset; i + 1 < end; i += 2) {
                result += String.fromCharCode(this.view.getUint16(i, true));
            }
            return result;
        }

        const bytes: number[] = [];
        for (let i = offset; i < offset + charCount; i++) bytes.push(this.data[i]);
        return decodeSingleByte(bytes);
    }

    // ════════════════════════════════════════════════════════════
    //  工作表子流解析
    // ════════════════════════════════════════════════════════════

    /**
     * 解析单个工作表子流
     *
     * 从 BOUNDSHEET 记录的 position 开始扫描，读取单元格记录、合并区域、
     * 列宽、行高与数据范围。
     *
     * @param sheet - 工作表位置信息
     * @returns 提取的单元格、合并区域、行列尺寸与数据范围
     * @example
     * const data = parser.parseSheet(wb.sheets[0]);
     * console.log(data.cells.length, data.range);
     */
    public parseSheet(sheet: BiffSheet): {
        cells: BiffCell[];
        merges: Array<{ s: { r: number, c: number }, e: { r: number, c: number } }>;
        cols: Array<{ index: number, width?: number }>;
        rows: Array<{ index: number, height?: number }>;
        images: BiffImage[];
        range: { minRow: number, maxRow: number, minCol: number, maxCol: number };
    } {

        const cells: BiffCell[] = [];
        const merges: Array<{ s: { r: number, c: number }, e: { r: number, c: number } }> = [];
        const cols: Array<{ index: number, width?: number }> = [];
        const rows: Array<{ index: number, height?: number }> = [];
        const images: BiffImage[] = [];

        // 数据范围用可变对象跟踪（避免闭包内局部变量无法被辅助方法修改）
        const range = { minRow: Infinity, maxRow: -1, minCol: Infinity, maxCol: -1 };

        /** 将某个坐标纳入数据范围 */
        const expand = (row: number, col: number): void => {
            if (row < range.minRow) range.minRow = row;
            if (row > range.maxRow) range.maxRow = row;
            if (col < range.minCol) range.minCol = col;
            if (col > range.maxCol) range.maxCol = col;
        };

        let pos = sheet.position;
        // 上一个 FORMULA 记录若为字符串结果，等待紧随的 STRING 记录回填
        let pendingFormulaCell: BiffCell | null = null;

        while (pos + 4 <= this.data.length) {
            const rec = this.readRecordHeader(pos);
            if (rec.dataStart + rec.length > this.data.length) break;

            const { type, length, dataStart } = rec;

            switch (type) {
                case REC.EOF:
                    return this.buildResult(cells, merges, cols, rows, images, range);

                case REC.DIMENSIONS: {
                    // BIFF8 DIMENSIONS：首行(u32)、末行+1(u32)、首列(u16)、末列+1(u16)、保留(u16)
                    if (length >= 14) {
                        const firstRow = this.view.getUint32(dataStart, true);
                        const lastRow = this.view.getUint32(dataStart + 4, true);
                        const firstCol = this.view.getUint16(dataStart + 8, true);
                        const lastCol = this.view.getUint16(dataStart + 10, true);
                        // 更新范围（数据行号为 0-based，末行为开区间）
                        // DIMENSIONS 的末行/末列为开区间，需减 1
                        if (lastRow > 0) expand(firstRow, 0);
                        if (lastCol > 0) {
                            range.minCol = Math.min(range.minCol, firstCol);
                            range.maxCol = Math.max(range.maxCol, lastCol - 1);
                        }
                    }
                    break;
                }

                case REC.NUMBER:
                case REC.RK: {
                    if (length >= 14) {
                        const row = this.view.getUint16(dataStart, true);
                        const col = this.view.getUint16(dataStart + 2, true);
                        const styleIndex = this.view.getUint16(dataStart + 4, true);
                        let value: number;
                        if (type === REC.NUMBER) {
                            value = this.view.getFloat64(dataStart + 6, true);
                        } else {
                            value = this.decodeRK(this.view.getUint32(dataStart + 6, true));
                        }
                        cells.push({ row, col, styleIndex, type: 'number', numberValue: value });
                        expand(row, col);
                    }
                    break;
                }

                case REC.MULRK: {
                    // 首行(u16) 首列(u16) [样式(u16) 值(u32)]* 末列(u16)
                    if (length >= 6) {
                        const row = this.view.getUint16(dataStart, true);
                        const firstCol = this.view.getUint16(dataStart + 2, true);
                        const entryCount = Math.floor((length - 6) / 6);
                        for (let i = 0; i < entryCount; i++) {
                            const base = dataStart + 4 + i * 6;
                            const styleIndex = this.view.getUint16(base, true);
                            const value = this.decodeRK(this.view.getUint32(base + 2, true));
                            const col = firstCol + i;
                            cells.push({ row, col, styleIndex, type: 'number', numberValue: value });
                            expand(row, col);
                        }
                    }
                    break;
                }

                case REC.LABELSST: {
                    if (length >= 10) {
                        const row = this.view.getUint16(dataStart, true);
                        const col = this.view.getUint16(dataStart + 2, true);
                        const styleIndex = this.view.getUint16(dataStart + 4, true);
                        const sstIndex = this.view.getUint32(dataStart + 6, true);
                        const text = this.workbook.sharedStrings[sstIndex] ?? '';
                        cells.push({ row, col, styleIndex, type: 'string', stringValue: text });
                        expand(row, col);
                    }
                    break;
                }

                case REC.LABEL:
                case REC.RSTRING: {
                    if (length >= 8) {
                        const row = this.view.getUint16(dataStart, true);
                        const col = this.view.getUint16(dataStart + 2, true);
                        const styleIndex = this.view.getUint16(dataStart + 4, true);
                        // 字符串：长度(u16) + 标志(u8) + 数据
                        const strLength = this.view.getUint16(dataStart + 6, true);
                        const strFlags = this.data[dataStart + 8];
                        const text = this.readUnicodeString(
                            dataStart + 9, strLength, (strFlags & 0x01) !== 0
                        );
                        cells.push({ row, col, styleIndex, type: 'string', stringValue: text });
                        expand(row, col);
                    }
                    break;
                }

                case REC.BLANK: {
                    // 空单元格：仅有样式，无值（不计入数据范围）
                    if (length >= 6) {
                        const row = this.view.getUint16(dataStart, true);
                        const col = this.view.getUint16(dataStart + 2, true);
                        const styleIndex = this.view.getUint16(dataStart + 4, true);
                        cells.push({ row, col, styleIndex, type: 'blank' });
                    }
                    break;
                }

                case REC.MULBLANK: {
                    // 结构：行(u16) 首列(u16) 样式(u16)* 末列(u16)
                    // ⚠️ 尾部还有 2 字节 lastCol，样式项数 = (length - 6) / 2，
                    // 写成 length / 2 或忘记扣 lastCol 都会多造一个空单元格。
                    if (length >= 6) {
                        const row = this.view.getUint16(dataStart, true);
                        const firstCol = this.view.getUint16(dataStart + 2, true);
                        const entryCount = Math.floor((length - 6) / 2);
                        for (let i = 0; i < entryCount; i++) {
                            const styleIndex = this.view.getUint16(dataStart + 4 + i * 2, true);
                            cells.push({ row, col: firstCol + i, styleIndex, type: 'blank' });
                        }
                    }
                    break;
                }

                case REC.BOOLERR: {
                    if (length >= 8) {
                        const row = this.view.getUint16(dataStart, true);
                        const col = this.view.getUint16(dataStart + 2, true);
                        const styleIndex = this.view.getUint16(dataStart + 4, true);
                        const value = this.data[dataStart + 6];
                        const isError = this.data[dataStart + 7] !== 0;
                        cells.push({
                            row, col, styleIndex,
                            type: isError ? 'error' : 'boolean',
                            errorValue: isError ? (this.workbook.errorValues[value] ?? '#ERR') : undefined,
                            booleanValue: !isError ? value !== 0 : undefined
                        });
                        expand(row, col);
                    }
                    break;
                }

                case REC.FORMULA: {
                    if (length >= 20) {
                        const row = this.view.getUint16(dataStart, true);
                        const col = this.view.getUint16(dataStart + 2, true);
                        const styleIndex = this.view.getUint16(dataStart + 4, true);
                        // 结果为 8 字节：若第 6-7 字节为 0xFFFF 则为特殊值
                        const resultHigh = this.view.getUint16(dataStart + 12, true);
                        const resultLow = this.view.getUint16(dataStart + 14, true);

                        if (resultHigh === 0xFFFF) {
                            // 特殊结果
                            const kind = this.data[dataStart + 6];
                            const value = this.data[dataStart + 8];
                            if (kind === 0) {
                                // 字符串结果：等待随后的 STRING 记录
                                const cell: BiffCell = {
                                    row, col, styleIndex, type: 'string', stringValue: ''
                                };
                                cells.push(cell);
                                pendingFormulaCell = cell;
                            } else if (kind === 1) {
                                cells.push({
                                    row, col, styleIndex, type: 'boolean', booleanValue: value !== 0
                                });
                            } else if (kind === 2) {
                                cells.push({
                                    row, col, styleIndex, type: 'error',
                                    errorValue: this.workbook.errorValues[value] ?? '#ERR'
                                });
                            } else {
                                // 空字符串
                                cells.push({ row, col, styleIndex, type: 'string', stringValue: '' });
                            }
                        } else {
                            // 普通 double 结果（8 字节）
                            const value = this.view.getFloat64(dataStart + 6, true);
                            cells.push({ row, col, styleIndex, type: 'number', numberValue: value });
                        }
                        expand(row, col);
                    }
                    break;
                }

                case REC.STRING: {
                    // FORMULA 的字符串结果，回填到 pending 单元格
                    if (pendingFormulaCell && length >= 4) {
                        const strLength = this.view.getUint16(dataStart, true);
                        const strFlags = this.data[dataStart + 2];
                        pendingFormulaCell.stringValue = this.readUnicodeString(
                            dataStart + 3, strLength, (strFlags & 0x01) !== 0
                        );
                        pendingFormulaCell = null;
                    }
                    break;
                }

                case REC.MERGEDCELLS: {
                    if (length >= 2) {
                        const count = this.view.getUint16(dataStart, true);
                        for (let i = 0; i < count; i++) {
                            const base = dataStart + 2 + i * 8;
                            if (base + 8 > this.data.length) break;
                            // MS-XLS [2.4.168] Pcmg 结构：
                            //   offset 0: rwFirst（首行）
                            //   offset 2: rwLast（末行）
                            //   offset 4: colFirst（首列）
                            //   offset 6: colLast（末列）
                            // ⚠️ 历史实现把 rwLast 和 colFirst 的偏移写反了
                            //   （rLast 读 offset+4、cFirst 读 offset+2），
                            //   导致合并区域的行列互换，渲染位置完全错乱。
                            let rFirst = this.view.getUint16(base, true);
                            let rLast = this.view.getUint16(base + 2, true);
                            let cFirst = this.view.getUint16(base + 4, true);
                            let cLast = this.view.getUint16(base + 6, true);

                            // ⚠️ 容错：厂商工具写出的文件里偶见 rFirst > rLast 这类
                            // 逆序（甚至超出 DIMENSIONS）的区域。若不规范化，
                            // 渲染端按「左上/右下」消费会得到负宽高的矩形，
                            // 表现为莫名其妙的空白遮挡或绘制越界。
                            if (rFirst > rLast) { const t = rFirst; rFirst = rLast; rLast = t; }
                            if (cFirst > cLast) { const t = cFirst; cFirst = cLast; cLast = t; }

                            // 零面积（单格「合并」）无意义，直接丢弃
                            if (rFirst === rLast && cFirst === cLast) continue;

                            merges.push({
                                s: { r: rFirst, c: cFirst },
                                e: { r: rLast, c: cLast }
                            });
                        }
                    }
                    break;
                }

                case REC.COLINFO: {
                    if (length >= 12) {
                        const firstCol = this.view.getUint16(dataStart, true);
                        const lastCol = this.view.getUint16(dataStart + 2, true);
                        // 宽度单位为 1/256 字符宽
                        const widthRaw = this.view.getUint16(dataStart + 4, true);
                        const width = widthRaw / 256;
                        for (let c = firstCol; c <= lastCol; c++) {
                            cols.push({ index: c, width });
                        }
                        range.maxCol = Math.max(range.maxCol, lastCol);
                    }
                    break;
                }

                case REC.ROW: {
                    if (length >= 16) {
                        const row = this.view.getUint16(dataStart, true);
                        const flags = this.view.getUint16(dataStart + 12, true);
                        // bit 15 = 行高由用户自定义，bit 14 = 默认行高标志
                        const hasCustomHeight = (flags & 0x8000) !== 0;
                        if (hasCustomHeight) {
                            // 行高单位为 twips（1/20 磅）
                            const heightTwips = this.view.getUint16(dataStart + 6, true);
                            rows.push({ index: row, height: Math.round(heightTwips / 20) });
                        }
                        range.maxRow = Math.max(range.maxRow, row);
                    }
                    break;
                }

                case REC.drawing: {
                    // Escher 绘制对象，提取内嵌图片。
                    // 未请求图片时**整段跳过**：BLIP 解析会把每张图的字节复制一份
                    // （见 parseBlip），含上千张图的 .xls 会凭空白拷上百 MB。
                    if (!this.readImages) break;
                    const sheetImages = this.parseDrawing(dataStart, length);
                    for (const img of sheetImages) {
                        images.push(img);
                    }
                    break;
                }
            }

            pos = dataStart + length;
        }

        return this.buildResult(cells, merges, cols, rows, images, range);
    }

    /**
     * 解析 BIFF 结果对象（统一处理空范围）
     *
     * @private
     */
    private buildResult(
        cells: BiffCell[],
        merges: Array<{ s: { r: number, c: number }, e: { r: number, c: number } }>,
        cols: Array<{ index: number, width?: number }>,
        rows: Array<{ index: number, height?: number }>,
        images: BiffImage[],
        range: { minRow: number, maxRow: number, minCol: number, maxCol: number }
    ) {
        return {
            cells, merges, cols, rows, images,
            range: {
                minRow: range.minRow === Infinity ? 0 : range.minRow,
                maxRow: range.maxRow < 0 ? 0 : range.maxRow,
                minCol: range.minCol === Infinity ? 0 : range.minCol,
                maxCol: range.maxCol < 0 ? 0 : range.maxCol
            }
        };
    }

    /**
     * 解码 RK 压缩数值
     *
     * RK 是 Excel 的整数/浮点压缩格式（32 位，小端序读入）：
     * - bit0：为 1 表示该值是 30 位**有符号整数**（值 = rk >> 2）
     * - bit0：为 0 表示该值是 IEEE double 的**高 32 位**（低 32 位补 0）
     * - bit1：为 1 表示结果还需 **除以 100**（用于存储两位小数）
     *
     * ⚠️ **字节序陷阱（曾导致 .xls 数值整表变成 1e-315 量级的垃圾值）**：
     * bit0=0 时，RK 的 30 位不是「double 的低位」，而是**高位**
     * （RK 用 32 位换取 double 的低 4 字节，因为低字节基本全 0）。
     * 因此必须写进 double 的**高 4 字节**（偏移 4），偏移 0 应保持 0：
     *
     * ```ts
     * dv.setUint32(0, 0, true);              // 低位补 0
     * dv.setUint32(4, rk & 0xFFFFFFFC, true); // ★ 高位 = RK 去掉低 2 位标志
     * ```
     *
     * 写成 `setUint32(0, rk & 0xFFFFFFFC)` 会得到非规格化（subnormal）数，
     * 典型症状：`0x403C0000`（应为 28）被解成 `5.32441693e-315`。
     *
     * @param {number} rk - RK 原始值（已按小端序读入的 u32）
     * @returns {number} 解码后的数值
     * @private
     */
    private decodeRK(rk: number): number {
        let value: number;

        if (rk & 0x01) {
            // 30 位有符号整数：整体左移 2 再算术右移 2，保留符号位
            value = (rk << 2) >> 2;
        } else {
            // RK 的 30 位是 double 的高 4 字节（低 4 字节补 0）
            const buf = new ArrayBuffer(8);
            const dv = new DataView(buf);
            dv.setUint32(0, 0, true);
            dv.setUint32(4, rk & 0xFFFFFFFC, true);
            value = dv.getFloat64(0, true);
        }

        // bit1 置位 → 除以 100（两位小数压缩）
        return (rk & 0x02) ? value / 100 : value;
    }

    /**
     * 解析 Escher 绘制对象中的图片
     *
     * .xls 的图片以 Escher BLIP 记录形式嵌入。记录结构（递归）：
     * ```
     * ┌──────────────┬──────────────┬────────────────┐
     * │ ver/instance │ 记录类型 u16 │ 长度 u32        │
     * └──────────────┴──────────────┴────────────────┘
     * ```
     * 容器记录（SpContainer/BSE）内嵌套具体的 BLIP 记录，
     * BLIP 记录头部含 16 字节 UID + 1 字节标记 + 1 字节 tag，
     * 之后才是真正的图片二进制数据。
     *
     * @param offset - 绘制对象记录的数据起始偏移
     * @param length - 记录长度
     * @returns 提取的图片列表
     * @private
     */
    private parseDrawing(offset: number, length: number): BiffImage[] {
        const images: BiffImage[] = [];
        // Escher 记录头为 8 字节（ver/instance u16 + type u16 + len u32）
        this.walkEscher(offset, offset + length, images);
        return images;
    }

    /**
     * 递归遍历 Escher 记录树
     *
     * @param start - 记录树起始偏移
     * @param end - 记录树结束偏移
     * @param images - 收集图片的数组
     * @private
     */
    private walkEscher(start: number, end: number, images: BiffImage[]): void {
        let pos = start;

        while (pos + 8 <= end) {
            // 实例号与版本(u16)：低 4 位是版本，高 12 位是实例
            const instanceAndVer = this.view.getUint16(pos, true);
            const recType = this.view.getUint16(pos + 2, true);
            // 长度可能是 u32，也可能用低 28 位表示（BIFF 限制 2^28）
            const recLen = this.view.getUint32(pos + 4, true) & 0x0FFFFFFF;
            const dataStart = pos + 8;
            const dataEnd = Math.min(dataStart + recLen, end);

            if (dataEnd <= dataStart && recLen > 0) break;

            // 判断是否为容器记录（需要递归）
            const isContainer = (instanceAndVer & 0x0F00) !== 0 ||
                recType === REC.msofbtSpContainer ||
                recType === REC.msofbtBSE;

            switch (recType) {
                case REC.msofbtBlip_PNG:
                case REC.msofbtBlip_JPEG:
                case REC.msofbtBlip_DIB:
                case REC.msofbtBlip_PICT:
                case REC.msofbtBlip_EMF:
                case REC.msofbtBlip_WMF: {
                    const img = this.parseBlip(recType, dataStart, dataEnd);
                    if (img) images.push(img);
                    break;
                }
                default: {
                    if (isContainer) {
                        // 容器：递归处理内部记录
                        this.walkEscher(dataStart, dataEnd, images);
                    }
                    break;
                }
            }

            pos = dataEnd > pos ? dataEnd : pos + 8;
        }
    }

    /**
     * 解析单个 BLIP 图片记录
     *
     * BLIP 记录结构：
     * - 0..15：UID（16 字节）
     * - 16：标记字节（0xFF 表示后面有 tag 字节）
     * - 17：tag 字节
     * - 18..：图片数据（部分格式前面还有 1-17 字节的额外头部）
     *
     * @param recType - BLIP 记录类型
     * @param start - 数据起始偏移
     * @param end - 数据结束偏移
     * @returns 解析出的图片；类型不支持时返回 null
     * @private
     */
    private parseBlip(recType: number, start: number, end: number): BiffImage | null {
        // 跳过 16 字节 UID
        let pos = start + 16;
        if (pos + 2 > end) return null;

        // 标记字节：0xFF 表示后面跟 1 字节 tag；也可能是直接给出两字节 tag
        const marker = this.data[pos];
        if (marker === 0xFF) {
            pos += 2;   // 跳过 0xFF 与 tag 字节
        } else {
            pos += 2;   // 两字节 tag
        }

        let mimeType: string;
        let headerSkip = 0;

        switch (recType) {
            case REC.msofbtBlip_PNG:
                mimeType = 'image/png';
                break;
            case REC.msofbtBlip_JPEG:
                mimeType = 'image/jpeg';
                // JPEG 额外头部：1 字节 tag + 1 字节 UID(可省略) + 4 字节起始偏移
                headerSkip = 1 + (marker === 0xFF ? 0 : 16) + 4;
                break;
            case REC.msofbtBlip_DIB:
                mimeType = 'image/bmp';
                headerSkip = marker === 0xFF ? 1 : 16 + 4;
                break;
            case REC.msofbtBlip_PICT:
                mimeType = 'image/pict';
                headerSkip = 0;
                break;
            case REC.msofbtBlip_EMF:
                mimeType = 'image/emf';
                headerSkip = 0;
                break;
            case REC.msofbtBlip_WMF:
                mimeType = 'image/wmf';
                headerSkip = 0;
                break;
            default:
                return null;
        }

        pos += headerSkip;
        if (pos >= end) return null;

        // 用 subarray 创建视图（不复制数据），再复制一份保证数据独立
        const imgData = new Uint8Array(this.data.subarray(pos, end));

        return { data: imgData, mimeType };
    }
}

/**
 * RGB 三字节转十六进制颜色
 * @private
 */
function rgbToHex(r: number, g: number, b: number): string {
    const toHex = (v: number): string => v.toString(16).padStart(2, '0');
    return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

/**
 * 单字节字符集解码
 *
 * BIFF 的压缩字符串使用工作簿的 CODEPAGE 指定的单字节编码。
 * 常见情形：
 * - 0x8001（1200）：UTF-16LE（此情况应走 Unicode 分支）
 * - 936（GBK）：简体中文
 * - 1252：西欧（Excel 默认）
 * - 950：Big5（繁体）
 *
 * 浏览器原生 TextDecoder 可解码其中大部分编码，无需自行实现。
 *
 * @param bytes - 单字节字符数据
 * @returns 解码后的字符串
 * @private
 */
function decodeSingleByte(bytes: number[]): string {
    if (bytes.length === 0) return '';

    // 全部为 ASCII 时直接拼接（最常见且最快）
    let isAscii = true;
    for (let i = 0; i < bytes.length; i++) {
        if (bytes[i] > 0x7F) { isAscii = false; break; }
    }
    if (isAscii) {
        let s = '';
        for (let i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
        return s;
    }

    // 含非 ASCII：交给 TextDecoder 处理
    const buffer = new Uint8Array(bytes);
    if (typeof TextDecoder !== 'undefined') {
        try {
            // 先尝试 UTF-8（现代 xls 多数用 UTF-16，非 ASCII 单字节多为本地代码页）
            return new TextDecoder('utf-8').decode(buffer);
        } catch (e) {
            // 回退到 Latin-1（不会抛异常，保证总能返回）
            let s = '';
            for (let i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
            return s;
        }
    }

    // 无 TextDecoder 环境的兜底
    let s = '';
    for (let i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
    return s;
}
