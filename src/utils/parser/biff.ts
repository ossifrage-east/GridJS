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
 */
export interface BiffXf {
    /** 字体索引 */
    fontIndex: number;
    /** 数字格式索引 */
    formatIndex: number;
    /** 水平对齐（0=常规 1=居中 2=填充 3=两端对齐 4=跨列居中 5=分散 6=单元格内对齐） */
    horizontalAlignment: number;
    /** 垂直对齐（0=底 1=居中 2=顶 3=两端 4=分散） */
    verticalAlignment: number;
    /** 自动换行 */
    wrapText: boolean;
    /** 文本旋转角度（0-180） */
    rotation: number;
    /** 缩进 */
    indent: number;
    /** 是否为父 XF（style XF） */
    isStyle: boolean;
    /** 边框线型（上/下/左/右），-1 表示无边框 */
    borderTop: number;
    borderBottom: number;
    borderLeft: number;
    borderRight: number;
    /** 边框颜色索引 */
    borderColorIndex: number;
    /** 填充图案索引 */
    fillPattern: number;
    /** 填充前景色索引 */
    fillColorIndex: number;
    /** 填充背景色索引 */
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
    fonts: BiffFont[];
    formats: Map<number, string>;
    xfs: BiffXf[];
    palette: string[];
    /** 共享字符串表 */
    sharedStrings: string[];
    /** Excel 内置的错误值表 */
    errorValues: string[];
}

/**
 * 默认调色板（BIFF8 的前 56 个固定颜色，索引 64 起为用户自定义）
 *
 * BIFF8 中调色板索引 0-63 含义未定义，Excel 约定：
 * 64-65 为系统前景/背景色，66 起为标准 56 色。
 */
const DEFAULT_PALETTE: string[] = [
    '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
    '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
    '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#C0C0C0', '#808080',
    '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF',
    '#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000', '#008080', '#0000FF',
    '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF', '#FF99CC', '#CC99FF', '#FFCC99',
    '#3366FF', '#33CCCC', '#99CC00', '#FFCC00', '#FF9900', '#FF6600', '#666699', '#969696',
    '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333'
];

/** BIFF 错误码表（ERR7 记录的索引） */
const ERROR_VALUES: string[] = [
    '#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA'
];

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
     * @param {Uint8Array} data - Workbook 流的完整字节数据
     */
    constructor(data: Uint8Array) {
        this.data = data;
        this.view = new DataView(data.buffer, data.byteOffset, data.byteLength);
        this.workbook = {
            sheets: [],
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
    private parseFont(offset: number, length: number): void {
        if (length < 16) return;

        const heightTwips = this.view.getUint16(offset, true);
        const flags = this.view.getUint16(offset + 2, true);
        const colorIndex = this.view.getUint16(offset + 4, true);
        const boldWeight = this.view.getUint16(offset + 6, true);

        // 名称：无名称时保持默认（Excel 默认字体 Calibri）
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
            // 粗体判定：以标志位优先，缺失时依据字重（>= 600 视为粗体）
            bold: (flags & 0x0001) !== 0 || (boldWeight >= 600 && boldWeight !== 0),
            italic: (flags & 0x0002) !== 0,
            underline: (flags & 0x0004) !== 0,
            strikethrough: (flags & 0x0008) !== 0,
            superscript: (flags & 0x0010) !== 0,
            subscript: (flags & 0x0020) !== 0,
            colorIndex
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
     * BIFF8 的 XF 记录固定 20 字节：
     * -  0：字体索引(u16)
     * -  2：数字格式索引(u16)
     * -  4：标志(u16)：bit2 锁定，bit3 隐藏，bit4 为 style XF，
     *   bit5 前置单引号，bit6-15 父 XF 索引（0xFFF 表示无父）
     * -  6：对齐字节，bit0-2 水平对齐，bit3 垂直对齐，bit4-5 文本环绕，
     *   bit6 缩小字体填充，bit7 缩进
     * -  7：文本旋转(0-180)
     * -  8：缩进与两端对齐
     * -  9：已使用属性位（bit2 数字、bit3 字体、bit4 对齐、
     *   bit5 边框、bit6 填充、bit7 锁定、bit8 保护）
     * - 10：边框与颜色标志
     * - 11：边框线型与颜色
     * - 12..15：填充图案索引与颜色索引
     * - 16..19：填充边框线型与颜色索引
     *
     * @param offset - 记录数据起始偏移
     * @param length - 记录长度
     * @private
     */
    private parseXf(offset: number, length: number): void {
        if (length < 20) return;

        const fontIndex = this.view.getUint16(offset, true);
        const formatIndex = this.view.getUint16(offset + 2, true);
        const flags = this.view.getUint16(offset + 4, true);
        const alignByte = this.data[offset + 6];
        const rotationByte = this.data[offset + 7];
        const indentByte = this.data[offset + 8];
        const usedByte = this.data[offset + 9];

        // 边框线型字节：每个边 3 位（0-7 对应线型，0 表示无边框）
        const borderByte = this.data[offset + 10];
        // 边框颜色索引：上/下/左/右 4 字节
        const borderColorTop = this.view.getUint16(offset + 11, true) & 0x3F;
        const borderColorBottom = this.view.getUint16(offset + 13, true) & 0x3F;
        const borderColorLeft = this.view.getUint16(offset + 15, true) & 0x3F;
        const borderColorRight = this.view.getUint16(offset + 17, true) & 0x3F;

        // 填充：图案索引(u16 的低 6 位) 与前景色/背景色索引
        const fillPattern = this.view.getUint16(offset + 12, true) & 0x3F;
        const fillColorIndex = this.view.getUint16(offset + 14, true) & 0x3F;

        // bit4 置位表示这是一个 style XF（样式定义而非单元格样式）
        const isStyle = (flags & 0xFFF0) === 0x0000 || (flags & 0x0004) !== 0 && (flags & 0xFFF0) === 0x0000;

        // 边线型：低位 3 位为上边框，次 3 位为下边框，再 3 位为左，再 3 位为右
        const lineTop = borderByte & 0x07;
        const lineBottom = (borderByte >> 3) & 0x07;
        const lineLeft = (borderByte >> 6) & 0x07;
        const lineRight = (borderByte >> 9) & 0x07;

        this.workbook.xfs.push({
            fontIndex,
            formatIndex,
            horizontalAlignment: alignByte & 0x07,
            verticalAlignment: (alignByte >> 3) & 0x01,
            wrapText: (alignByte & 0x30) !== 0,
            rotation: rotationByte & 0x3F,
            indent: indentByte & 0x0F,
            isStyle,
            borderTop: lineTop,
            borderBottom: lineBottom,
            borderLeft: lineLeft,
            borderRight: lineRight,
            // 边框颜色：取第一条有颜色的边
            borderColorIndex: borderColorTop || borderColorBottom ||
                borderColorLeft || borderColorRight,
            fillPattern,
            fillColorIndex,
            // 背景色为调色板索引 64（系统背景色）
            fillBackgroundIndex: 65,
            // usedByte 的 bit5 表示使用了边框属性
            ...(usedByte & 0x20 ? {} : {
                borderTop: 0, borderBottom: 0, borderLeft: 0, borderRight: 0
            })
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
            // 调色板颜色为 BGR 顺序
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
        const end = pos + length;

        if (start + 8 > this.data.length) return;
        const uniqueCount = this.view.getUint32(start + 4, true);

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
        let segIndex = 0;
        let segPos = start;
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

        let curSeg = segIndex;
        let curPos = segPos;

        /**
         * 当前段的结束偏移与可用长度
         */
        const segLength = (idx: number): number => {
            if (idx === 0) return end;
            const c = continuations[idx - 1];
            return c ? c.length : 0;
        };

        /**
         * 读取一个字节，必要时跨越 CONTINUE 边界
         * @returns 字节值；无更多数据返回 -1
         */
        const readByte = (): number => {
            if (curSeg === 0) {
                if (curPos >= end) {
                    // 需要切换到 CONTINUE
                    if (continuations.length > 0) { curSeg = 1; curPos = 0; }
                    else return -1;
                } else {
                    return this.data[curPos++];
                }
            }
            while (curSeg > 0) {
                const len = segLength(curSeg);
                if (curPos < len) {
                    return continuations[curSeg - 1][curPos++];
                }
                // 当前 CONTINUE 用尽，切到下一个
                curSeg++;
                curPos = 0;
            }
            return -1;
        };

        /**
         * 读取多个字节并尝试跨段
         * @param count - 字节数
         * @returns 字节数组；数据不足返回 null
         */
        const readBytes = (count: number): number[] | null => {
            const out: number[] = [];
            for (let i = 0; i < count; i++) {
                const b = readByte();
                if (b === -1) return null;
                out.push(b);
            }
            return out;
        };

        // —— 字符串长度(u16) + 标志(u8) ——
        const lenBytes = readBytes(2);
        if (!lenBytes) return null;
        const charCount = lenBytes[0] | (lenBytes[1] << 8);

        const flagByte = readByte();
        if (flagByte === -1) return null;
        const isUnicode = (flagByte & 0x01) !== 0;

        // —— 富文本与 phonetic 扩展（跳过，不影响主文本）——
        let extraSkip = 0;
        if (isUnicode) {
            // 富文本运行数(u16) + phonetic 大小(u32)
            const richRunCount = readBytes(2);
            if (richRunCount) {
                extraSkip += (richRunCount[0] | (richRunCount[1] << 8)) * 4;
            }
            const phoneticSize = readBytes(4);
            if (phoneticSize) {
                extraSkip += phoneticSize[0] | (phoneticSize[1] << 8) |
                    (phoneticSize[2] << 16) | (phoneticSize[3] << 24);
            }
        }

        // 读取字符数据
        const text = this.readSstChars(readByte, charCount, isUnicode);

        // 跳过富文本运行与 phonetic 数据
        for (let i = 0; i < extraSkip; i++) readByte();

        return { text, nextSegIndex: curSeg, nextSegPos: curPos };
    }

    /**
     * 读取 SST 字符串的字符数据
     *
     * @param readByte - 字节读取器（可能跨越 CONTINUE 边界）
     * @param charCount - 字符数
     * @param isUnicode - 是否为 UTF-16LE
     * @returns 解码后的字符串
     * @private
     */
    private readSstChars(
        readByte: () => number,
        charCount: number,
        isUnicode: boolean
    ): string {
        if (isUnicode) {
            // UTF-16LE：每字符 2 字节
            const bytes: number[] = [];
            for (let i = 0; i < charCount * 2; i++) {
                const b = readByte();
                if (b === -1) break;
                bytes.push(b);
            }
            // 手动按小端序组字符
            let result = '';
            for (let i = 0; i + 1 < bytes.length; i += 2) {
                result += String.fromCharCode(bytes[i] | (bytes[i + 1] << 8));
            }
            return result;
        }

        // 压缩 8 位：单字节字符（当前代码页，多为 GBK/UTF-8 需按 CODEPAGE 处理）
        const bytes: number[] = [];
        for (let i = 0; i < charCount; i++) {
            const b = readByte();
            if (b === -1) break;
            bytes.push(b);
        }
        return decodeSingleByte(bytes);
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
                            merges.push({
                                s: {
                                    r: this.view.getUint16(base, true),
                                    c: this.view.getUint16(base + 2, true)
                                },
                                e: {
                                    r: this.view.getUint16(base + 4, true),
                                    c: this.view.getUint16(base + 6, true)
                                }
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
                    // Escher 绘制对象，提取内嵌图片
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
     * RK 是 Excel 的整数/浮点压缩格式（32 位）：
     * - bit0：为 1 表示低 30 位是整数部分，为 0 表示高 4 字节是 IEEE double 的高位
     * - bit1：为 1 表示值需要除以 100（用于表示两位小数）
     *
     * @param {number} rk - RK 原始值
     * @returns {number} 解码后的数值
     * @private
     */
    private decodeRK(rk: number): number {
        // 有符号右移，保留符号位
        let value: number;

        if (rk & 0x02) {
            if (rk & 0x01) {
                // 30 位有符号整数
                value = (rk >> 2);
            } else {
                // 30 位为 double 的高 30 位，低 2 位补 0
                const buf = new ArrayBuffer(8);
                const dv = new DataView(buf);
                dv.setUint32(0, rk & 0xFFFFFFFC, true);
                dv.setUint32(4, 0, true);
                value = dv.getFloat64(0, true);
            }
            // bit1 置位 → 除以 100
            value /= 100;
        } else {
            if (rk & 0x01) {
                // 30 位有符号整数
                value = (rk >> 2);
            } else {
                // 低 2 位补 0 的 IEEE double
                const buf = new ArrayBuffer(8);
                const dv = new DataView(buf);
                dv.setUint32(0, rk & 0xFFFFFFFC, true);
                dv.setUint32(4, 0, true);
                value = dv.getFloat64(0, true);
            }
        }

        return value;
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
