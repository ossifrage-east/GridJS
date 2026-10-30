/*
 * XLSX 写出器
 *
 * 把 `.grid` 文件在 IndexedDB 中保存的**表快照**（v3 文本 / v4 分片对象）转换成
 * 一份可被 Excel / WPS 打开的 `.xlsx`。写出用「存储（不压缩）」方式的 ZIP，
 * 因此不依赖任何第三方库（项目里原本只有解析用的 ZipParser，没有写出能力）。
 *
 * 覆盖范围：单元格文本 / 数字、字体、背景色、四边边框、水平与垂直对齐、
 * 自动换行、缩进、数字格式、合并单元格、自定义列宽与行高、隐藏行列、网格线开关。
 * 图片（IndexedDB 的 `images` 库）暂不写入 —— 写 drawing 部件属于另一条独立链路。
 *
 * 对外提供的是**可暂停的作业**（见 {@link createXlsxTask}）：转换被拆成多个小单元，
 * 由调用方按时间片推进，「暂停」就是不推进，「继续」从断点接着跑。
 */

import { DEFAULT_CELL_WIDTH, DEFAULT_CELL_HEIGHT } from '../../core/constant';

/** ZIP 条目：文件名 + 原始字节 */
export interface ZipEntryInput {
    /** 条目名（ZIP 内路径，如 `xl/workbook.xml`） */
    name: string;
    /** 文件内容 */
    data: Uint8Array;
}

/** 单元格样式（从快照的样式表项归一化而来） */
interface CellStyle {
    fontFamily?: string;
    fontSize?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strike?: boolean;
    fontColor?: string;
    bg?: string;
    ta?: string;
    ai?: string;
    wrap?: boolean;
    indent?: number;
    numFmt?: string;
    bt?: number;
    bb?: number;
    bl?: number;
    br?: number;
    borderColor?: string;
}

/** 归一化后的单元格 */
interface ExportCell {
    col: number;
    row: number;
    text: string;
    colspan: number;
    rowspan: number;
    style: CellStyle;
}

/** 轴（列 / 行）的尺寸信息 */
interface Axis {
    count: number;
    defaultSize: number;
    sizeFlat: number[];
    hidden: number[];
}

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MAX_COL = 16384;
const MAX_ROW = 1048576;

/** XML 文本转义（文本节点与属性值共用；属性里的引号一并转义） */
function escapeXml(value: string): string {
    return value
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;')
        // XML 1.0 不允许的控制字符直接丢弃
        .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
}

/** 单元格内的换行在 `<t>` 里写成字符实体，避免被解析器折叠 */
function escapeCellText(value: string): string {
    return escapeXml(value).replace(/\r\n|\r|\n/g, '&#10;');
}

/** 列号（1-based）→ 列名（如 1 → A、27 → AA） */
function colName(col: number): string {
    let n = col;
    let name = '';
    while (n > 0) {
        const rem = (n - 1) % 26;
        name = String.fromCharCode(65 + rem) + name;
        n = Math.floor((n - 1) / 26);
    }
    return name || 'A';
}

/** 列名 → 列号（1-based），非法名称返回 0 */
function colIndex(name: string): number {
    let n = 0;
    for (const ch of name) {
        const code = ch.charCodeAt(0);
        if (code < 65 || code > 90) return 0;
        n = n * 26 + (code - 64);
    }
    return n;
}

/** 解析单元格名称（`A1` / `$B$3`） */
function parseRef(ref: string): { col: number; row: number } | null {
    const match = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(ref ?? '');
    if (!match) return null;
    const col = colIndex(match[1].toUpperCase());
    const row = parseInt(match[2], 10);
    if (col <= 0 || row <= 0) return null;
    return { col, row };
}

/** 数值颜色 → ARGB（无法识别时返回 null，调用方直接省略颜色） */
function toArgb(color: string | undefined): string | null {
    if (!color) return null;
    const raw = color.trim().toLowerCase();
    if (!raw) return null;
    const named: Record<string, string> = {
        black: 'FF000000', white: 'FFFFFFFF', red: 'FFFF0000', green: 'FF008000',
        blue: 'FF0000FF', yellow: 'FFFFFF00', gray: 'FF808080', grey: 'FF808080',
        orange: 'FFFFA500', purple: 'FF800080',
    };
    if (named[raw]) return named[raw];
    if (/^#[0-9a-f]{3}$/.test(raw)) {
        return 'FF' + raw.slice(1).split('').map((c) => c + c).join('').toUpperCase();
    }
    if (/^#[0-9a-f]{6}$/.test(raw)) return 'FF' + raw.slice(1).toUpperCase();
    if (/^#[0-9a-f]{8}$/.test(raw)) return raw.slice(1).toUpperCase();
    return null;
}

/** 背景色是否值得写出一条填充（白色与缺省等价） */
function isVisibleFill(color: string | undefined): boolean {
    const argb = toArgb(color);
    return argb !== null && argb !== 'FFFFFFFF';
}

/** 边框宽度 → OOXML 边框样式 */
function borderStyle(width: number | undefined): string | null {
    if (!width || width <= 0) return null;
    if (width <= 1) return 'thin';
    if (width <= 2) return 'medium';
    return 'thick';
}

/** 纯数字文本 → 数值；带前导零、空串、非有限数一律返回 null（按文本写） */
function numericValue(text: string): number | null {
    if (!text || text.trim() !== text) return null;
    if (!/^-?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(text)) return null;
    const unsigned = text.replace(/^-/, '');
    const integerPart = unsigned.split(/[.eE]/)[0];
    if (integerPart.length > 1 && integerPart.startsWith('0')) return null;
    const value = Number(text);
    return Number.isFinite(value) ? value : null;
}

/** 从快照的样式表项 / 内联字段提取单元格样式与整格字体 */
function readStyle(entry: Record<string, unknown> | undefined, raw: Record<string, unknown>): CellStyle {
    const style: CellStyle = {};
    const src = entry ?? raw;
    const font = (src.f ?? raw.f ?? null) as Record<string, unknown> | null;
    if (font) {
        if (typeof font.n === 'string') style.fontFamily = font.n;
        if (typeof font.s === 'number') style.fontSize = font.s;
        if (font.b === 1 || font.b === true) style.bold = true;
        if (font.i === 1 || font.i === true) style.italic = true;
        if (font.u === 1 || font.u === true) style.underline = true;
        if (font.k === 1 || font.k === true) style.strike = true;
        if (typeof font.o === 'string') style.fontColor = font.o;
    }
    if (typeof src.bg === 'string') style.bg = src.bg;
    if (typeof src.ta === 'string') style.ta = src.ta;
    if (typeof src.ai === 'string') style.ai = src.ai;
    if (src.w === 1 || src.w === true) style.wrap = true;
    if (typeof src.in === 'number') style.indent = src.in;
    if (typeof src.nf === 'string' && src.nf !== 'General' && src.nf !== '@') style.numFmt = src.nf;
    if (typeof src.bt === 'number') style.bt = src.bt;
    if (typeof src.bb === 'number') style.bb = src.bb;
    if (typeof src.bl === 'number') style.bl = src.bl;
    if (typeof src.br === 'number') style.br = src.br;
    if (typeof src.bc === 'string') style.borderColor = src.bc;
    return style;
}

/** 单元格文本：`t` 优先，否则逐字符拼接 `a` */
function cellText(raw: Record<string, unknown>): string {
    if (typeof raw.t === 'string') return raw.t;
    if (Array.isArray(raw.a)) {
        let text = '';
        for (const item of raw.a) {
            const ch = item as Record<string, unknown> | null;
            if (ch && typeof ch.c === 'string') text += ch.c;
        }
        return text;
    }
    return '';
}

/** 逐字符数组里第一枚字符的字体（没有 `t` 时用它当整格字体） */
function firstCharFont(raw: Record<string, unknown>): Record<string, unknown> | null {
    if (!Array.isArray(raw.a) || raw.a.length === 0) return null;
    const first = raw.a[0] as Record<string, unknown> | null;
    return first ?? null;
}

/** 快照里的稀疏轴（v4 `cols` / `rows`）归一化 */
function normalizeAxis(raw: unknown, defaultSize: number, limit: number): Axis {
    const axis = (raw ?? {}) as Partial<Axis>;
    const count = Number.isFinite(axis.count)
        ? Math.max(0, Math.min(Math.floor(axis.count as number), limit))
        : 0;
    return {
        count,
        defaultSize: Number.isFinite(axis.defaultSize) ? (axis.defaultSize as number) : defaultSize,
        sizeFlat: Array.isArray(axis.sizeFlat) ? axis.sizeFlat.filter((v) => Number.isFinite(v)) : [],
        hidden: Array.isArray(axis.hidden) ? axis.hidden.filter((v) => Number.isFinite(v)) : [],
    };
}

/** v3 的列头 / 行头数组 → 稀疏轴 */
function axisFromHeaders(raw: unknown, defaultSize: number, sizeKey: 'width' | 'height', limit: number): Axis {
    if (!Array.isArray(raw)) return { count: 0, defaultSize, sizeFlat: [], hidden: [] };
    const count = Math.min(raw.length, limit);
    const sizeFlat: number[] = [];
    const hidden: number[] = [];
    for (let i = 0; i < count; i++) {
        const item = raw[i] as Record<string, unknown> | null;
        if (!item) continue;
        const size = item[sizeKey];
        if (typeof size === 'number' && size !== defaultSize) sizeFlat.push(i, size);
        if (item.isHidden === true) hidden.push(i);
    }
    return { count, defaultSize, sizeFlat, hidden };
}

/** 一次时间片处理的单元格数（把单次同步工作压在亚毫秒级，暂停才能及时生效） */
const CELLS_PER_UNIT = 512;

/** 一次时间片处理的行数 */
const ROWS_PER_UNIT = 64;

/** 进度分段（0–100）：解析存档 → 逐格归一化 → 按行聚合 → 逐行生成 XML → 收尾 */
const PROGRESS_PARSE = 10;
const PROGRESS_CELLS = 45;
const PROGRESS_GROUP = 48;
const PROGRESS_ROWS = 90;

/** 样式登记表：把「字体 / 填充 / 边框 / 数字格式 / 组合样式」统一去重后产出 styles.xml */
class StyleRegistry {
    private readonly fonts: string[] = [];
    private readonly fontKeys = new Map<string, number>();
    private readonly fills: string[] = [];
    private readonly fillKeys = new Map<string, number>();
    private readonly borders: string[] = [];
    private readonly borderKeys = new Map<string, number>();
    private readonly numFmts: Array<{ id: number; code: string }> = [];
    private readonly numFmtKeys = new Map<string, number>();
    private readonly xfs: string[] = [];
    private readonly xfKeys = new Map<string, number>();

    public constructor() {
        // 0 号字体 / 填充 / 边框 / xf 是 Excel 必需的基础项
        this.defaultFont();
        this.fills.push('<fill><patternFill patternType="none"/></fill>');
        this.fills.push('<fill><patternFill patternType="gray125"/></fill>');
        this.borders.push('<border><left/><right/><top/><bottom/><diagonal/></border>');
        this.xfs.push('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>');
    }

    /** 返回单元格样式对应的 cellXfs 下标（无任何样式时为 0） */
    public xfId(style: CellStyle): number {
        const hasFont = !!(style.fontFamily || style.fontSize || style.bold || style.italic
            || style.underline || style.strike || style.fontColor);
        const hasFill = isVisibleFill(style.bg);
        const hasBorder = !!(style.bt || style.bb || style.bl || style.br);
        const hasAlign = !!(style.ta || style.ai || style.wrap || style.indent);
        const hasNumFmt = !!style.numFmt;
        if (!hasFont && !hasFill && !hasBorder && !hasAlign && !hasNumFmt) return 0;

        const fontId = hasFont ? this.fontId(style) : 0;
        const fillId = hasFill ? this.fillId(style.bg as string) : 0;
        const borderId = hasBorder ? this.borderId(style) : 0;
        const numFmtId = hasNumFmt ? this.numFmtId(style.numFmt as string) : 0;

        const horizontal = this.horizontal(style.ta);
        const vertical = this.vertical(style.ai);
        const wrap = style.wrap === true;
        const indent = typeof style.indent === 'number' && style.indent > 0 ? style.indent : 0;
        const align = horizontal || vertical || wrap || indent > 0;

        const key = [numFmtId, fontId, fillId, borderId, horizontal, vertical, wrap, indent].join('|');
        const hit = this.xfKeys.get(key);
        if (hit !== undefined) return hit;

        const attrs = [
            `numFmtId="${numFmtId}"`, `fontId="${fontId}"`, `fillId="${fillId}"`,
            `borderId="${borderId}"`, 'xfId="0"',
            numFmtId > 0 ? 'applyNumberFormat="1"' : '',
            fontId > 0 ? 'applyFont="1"' : '',
            fillId > 0 ? 'applyFill="1"' : '',
            borderId > 0 ? 'applyBorder="1"' : '',
            align ? 'applyAlignment="1"' : '',
        ].filter(Boolean).join(' ');

        let inner = '';
        if (align) {
            const parts = [
                horizontal ? `horizontal="${horizontal}"` : '',
                vertical ? `vertical="${vertical}"` : '',
                wrap ? 'wrapText="1"' : '',
                indent > 0 ? `indent="${indent}"` : '',
            ].filter(Boolean).join(' ');
            inner = `<alignment ${parts}/>`;
        }
        const xml = inner ? `<xf ${attrs}>${inner}</xf>` : `<xf ${attrs}/>`;
        const id = this.xfs.length;
        this.xfs.push(xml);
        this.xfKeys.set(key, id);
        return id;
    }

    /** 产出 styles.xml */
    public toXml(): string {
        const numFmts = this.numFmts.length
            ? `<numFmts count="${this.numFmts.length}">${this.numFmts
                .map((item) => `<numFmt numFmtId="${item.id}" formatCode="${escapeXml(item.code)}"/>`)
                .join('')}</numFmts>`
            : '';
        return XML_HEAD
            + `<styleSheet xmlns="${NS_MAIN}">`
            + numFmts
            + `<fonts count="${this.fonts.length}">${this.fonts.join('')}</fonts>`
            + `<fills count="${this.fills.length}">${this.fills.join('')}</fills>`
            + `<borders count="${this.borders.length}">${this.borders.join('')}</borders>`
            + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            + `<cellXfs count="${this.xfs.length}">${this.xfs.join('')}</cellXfs>`
            + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            + '<dxfs count="0"/>'
            + '</styleSheet>';
    }

    /** 默认字体（0 号） */
    private defaultFont(): void {
        this.fonts.push('<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>');
        this.fontKeys.set('default', 0);
    }

    /** 字体：家族 / 字号 / 粗斜体 / 下划线 / 删除线 / 颜色 */
    private fontId(style: CellStyle): number {
        const family = style.fontFamily || 'Calibri';
        const size = typeof style.fontSize === 'number' ? style.fontSize : 11;
        const color = toArgb(style.fontColor);
        const key = [family, size, style.bold === true, style.italic === true,
            style.underline === true, style.strike === true, color ?? ''].join('|');
        const hit = this.fontKeys.get(key);
        if (hit !== undefined) return hit;
        const parts = [
            style.bold ? '<b/>' : '',
            style.italic ? '<i/>' : '',
            style.strike ? '<strike/>' : '',
            style.underline ? '<u/>' : '',
            `<sz val="${size}"/>`,
            color ? `<color rgb="${color}"/>` : '',
            `<name val="${escapeXml(family)}"/>`,
            '<family val="2"/>',
        ].join('');
        const id = this.fonts.length;
        this.fonts.push(`<font>${parts}</font>`);
        this.fontKeys.set(key, id);
        return id;
    }

    /** 纯色填充 */
    private fillId(color: string): number {
        const argb = toArgb(color) ?? 'FFFFFFFF';
        const key = argb;
        const hit = this.fillKeys.get(key);
        if (hit !== undefined) return hit;
        const id = this.fills.length;
        this.fills.push(`<fill><patternFill patternType="solid"><fgColor rgb="${argb}"/><bgColor indexed="64"/></patternFill></fill>`);
        this.fillKeys.set(key, id);
        return id;
    }

    /** 四边边框（颜色统一取 `bc`，缺省黑色） */
    private borderId(style: CellStyle): number {
        const color = toArgb(style.borderColor) ?? 'FF000000';
        const sides = {
            left: borderStyle(style.bl),
            right: borderStyle(style.br),
            top: borderStyle(style.bt),
            bottom: borderStyle(style.bb),
        };
        const key = `${sides.left}|${sides.right}|${sides.top}|${sides.bottom}|${color}`;
        const hit = this.borderKeys.get(key);
        if (hit !== undefined) return hit;
        const sideXml = (name: string, value: string | null): string =>
            value ? `<${name} style="${value}"><color rgb="${color}"/></${name}>` : `<${name}/>`;
        const id = this.borders.length;
        this.borders.push('<border>'
            + sideXml('left', sides.left)
            + sideXml('right', sides.right)
            + sideXml('top', sides.top)
            + sideXml('bottom', sides.bottom)
            + '<diagonal/>'
            + '</border>');
        this.borderKeys.set(key, id);
        return id;
    }

    /** 自定义数字格式（从 164 起编号） */
    private numFmtId(code: string): number {
        const hit = this.numFmtKeys.get(code);
        if (hit !== undefined) return hit;
        const id = 164 + this.numFmts.length;
        this.numFmts.push({ id, code });
        this.numFmtKeys.set(code, id);
        return id;
    }

    /** 水平对齐 */
    private horizontal(value: string | undefined): string {
        switch (value) {
            case 'left': return 'left';
            case 'center': return 'center';
            case 'right': return 'right';
            case 'justify': return 'justify';
            default: return '';
        }
    }

    /** 垂直对齐 */
    private vertical(value: string | undefined): string {
        switch (value) {
            case 'top': return 'top';
            case 'middle': return 'center';
            case 'bottom': return 'bottom';
            default: return '';
        }
    }
}

/** 表名清洗：去掉 Excel 不允许的字符，最长 31 字符 */
function sheetName(name: string): string {
    const cleaned = (name || '').replace(/[\\/*?:[\]]/g, '_').trim().slice(0, 31);
    return cleaned || 'Sheet1';
}

/** 列宽（像素）→ Excel 字符宽（按 7px ≈ 1 字符） */
function widthChars(px: number): number {
    return Math.round((px / 7) * 100) / 100;
}

/** 行高（像素）→ 磅（96dpi → 72dpi） */
function heightPoints(px: number): number {
    return Math.round(px * 0.75 * 100) / 100;
}

/** 构建 workbook.xml */
function buildWorkbookXml(name: string): string {
    return XML_HEAD
        + `<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">`
        + '<bookViews><workbookView/></bookViews>'
        + `<sheets><sheet name="${escapeXml(sheetName(name))}" sheetId="1" r:id="rId1"/></sheets>`
        + '</workbook>';
}

/** workbook 的关系：工作表 + 样式表 */
const WORKBOOK_RELS = XML_HEAD
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + `<Relationship Id="rId1" Type="${NS_REL}/worksheet" Target="worksheets/sheet1.xml"/>`
    + `<Relationship Id="rId2" Type="${NS_REL}/styles" Target="styles.xml"/>`
    + '</Relationships>';

const PACKAGE_RELS = XML_HEAD
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + `<Relationship Id="rId1" Type="${NS_REL}/officeDocument" Target="xl/workbook.xml"/>`
    + '</Relationships>';

const CONTENT_TYPES = XML_HEAD
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    + '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    + '</Types>';

/* ══════════════════════════════════════════════════════════════════
 *  ZIP 写出（存储方式，不压缩）
 * ══════════════════════════════════════════════════════════════════ */

let crcTable: Uint32Array | null = null;

/** 构建 CRC32 查找表 */
function getCrcTable(): Uint32Array {
    if (crcTable) return crcTable;
    const table = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
        let c = i;
        for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
        table[i] = c >>> 0;
    }
    crcTable = table;
    return table;
}

/** 计算字节块的 CRC32 */
function crc32(data: Uint8Array): number {
    const table = getCrcTable();
    let crc = 0xFFFFFFFF;
    for (let i = 0; i < data.length; i++) {
        crc = table[(crc ^ data[i]) & 0xFF] ^ (crc >>> 8);
    }
    return (crc ^ 0xFFFFFFFF) >>> 0;
}

/** 当前时间 → ZIP 的 DOS 时间 / 日期 */
function dosDateTime(now: Date): { time: number; date: number } {
    const year = Math.max(1980, now.getFullYear());
    return {
        time: (now.getHours() << 11) | (now.getMinutes() << 5) | Math.floor(now.getSeconds() / 2),
        date: ((year - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate(),
    };
}

/** 写入一个 4 字节小端整数 */
function writeU32(view: Uint8Array, offset: number, value: number): void {
    view[offset] = value & 0xFF;
    view[offset + 1] = (value >>> 8) & 0xFF;
    view[offset + 2] = (value >>> 16) & 0xFF;
    view[offset + 3] = (value >>> 24) & 0xFF;
}

/** 写入一个 2 字节小端整数 */
function writeU16(view: Uint8Array, offset: number, value: number): void {
    view[offset] = value & 0xFF;
    view[offset + 1] = (value >>> 8) & 0xFF;
}

/**
 * 打包成 ZIP（不压缩，UTF-8 文件名）
 * @param {ZipEntryInput[]} entries - 条目列表
 * @returns {Blob} ZIP 文件
 */
export function createZipBlob(entries: ZipEntryInput[]): Blob {
    const encoder = new TextEncoder();
    const { time, date } = dosDateTime(new Date());
    const localParts: Uint8Array[] = [];
    const centralParts: Uint8Array[] = [];
    let offset = 0;

    for (const entry of entries) {
        const nameBytes = encoder.encode(entry.name);
        const data = entry.data;
        const crc = crc32(data);

        const local = new Uint8Array(30 + nameBytes.length);
        writeU32(local, 0, 0x04034b50);
        writeU16(local, 4, 20);               // version needed
        writeU16(local, 6, 0x0800);           // UTF-8 文件名
        writeU16(local, 8, 0);                // method: store
        writeU16(local, 10, time);
        writeU16(local, 12, date);
        writeU32(local, 14, crc);
        writeU32(local, 18, data.length);
        writeU32(local, 22, data.length);
        writeU16(local, 26, nameBytes.length);
        writeU16(local, 28, 0);
        local.set(nameBytes, 30);

        localParts.push(local, data);

        const central = new Uint8Array(46 + nameBytes.length);
        writeU32(central, 0, 0x02014b50);
        writeU16(central, 4, 20);             // version made by
        writeU16(central, 6, 20);             // version needed
        writeU16(central, 8, 0x0800);
        writeU16(central, 10, 0);
        writeU16(central, 12, time);
        writeU16(central, 14, date);
        writeU32(central, 16, crc);
        writeU32(central, 20, data.length);
        writeU32(central, 24, data.length);
        writeU16(central, 28, nameBytes.length);
        writeU16(central, 30, 0);
        writeU16(central, 32, 0);
        writeU16(central, 34, 0);
        writeU16(central, 36, 0);
        writeU32(central, 38, 0);             // external attrs
        writeU32(central, 42, offset);        // local header offset
        central.set(nameBytes, 46);
        centralParts.push(central);

        offset += local.length + data.length;
    }

    const centralSize = centralParts.reduce((sum, part) => sum + part.length, 0);
    const eocd = new Uint8Array(22);
    writeU32(eocd, 0, 0x06054b50);
    writeU16(eocd, 4, 0);
    writeU16(eocd, 6, 0);
    writeU16(eocd, 8, entries.length);
    writeU16(eocd, 10, entries.length);
    writeU32(eocd, 12, centralSize);
    writeU32(eocd, 16, offset);
    writeU16(eocd, 20, 0);

    return new Blob([...localParts, ...centralParts, eocd] as BlobPart[], { type: 'application/zip' });
}

/**
 * 可暂停的 xlsx 转换作业
 *
 * 调用方按时间片反复调用 {@link XlsxTask.step}，返回 false 表示「本片额度用完，
 * 但还没转换完」，此时停止调用即可把转换冻在断点；再次调用 step 会从断点继续
 * （不会重新解析已解析的分片，也不会重复生成已产出的行）。
 */
export interface XlsxTask {
    /** 转换进度（0–100，单调递增；未开始转换时为 0） */
    readonly progress: number;
    /** 转换产物；尚未完成时为 null */
    readonly blob: Blob | null;
    /**
     * 同步推进转换，直到完成或越过 `deadline`（`performance.now()` 时间轴）
     * @param {number} deadline - 本片允许消耗到的时间点
     * @returns {boolean} 是否已完成（完成时 {@link XlsxTask.blob} 可用）
     */
    step(deadline: number): boolean;
}

/** 转换阶段 */
type XlsxPhase = 'parse' | 'cells' | 'group' | 'rows' | 'finalize' | 'done';

/**
 * {@link XlsxTask} 的实现：把原本一次跑完的同步转换拆成可复用的单元
 *
 * 单元粒度：一个分片的 `JSON.parse` / 一批 {@link CELLS_PER_UNIT} 个单元格 /
 * 一批 {@link ROWS_PER_UNIT} 行 / 收尾打包。状态全部保存在实例字段上，
 * 因此暂停 = 不再调用 {@link step}，继续 = 接着调用。
 *
 * 不可中断的同步段只有两处：单个分片的 `JSON.parse`，以及最后的
 * 「编码 + CRC + ZIP」（通常几十毫秒），后者被安排在收尾阶段。
 */
class XlsxJob implements XlsxTask {
    /** 进度（0–100） */
    public progress = 0;
    /** 完成后的产物 */
    public blob: Blob | null = null;

    /** 当前阶段 */
    private phase: XlsxPhase = 'parse';
    private readonly name: string;
    private readonly raw: unknown;

    /** 解析阶段：快照头部只解析一次，之后逐片解析 `chunks` */
    private snapshot: Record<string, unknown> = {};
    private styles: Record<string, unknown>[] = [];
    private rawCells: Record<string, unknown>[] = [];
    private chunks: unknown[] = [];
    private chunkIndex = 0;
    private headerParsed = false;
    private cols: Axis = { count: 0, defaultSize: DEFAULT_CELL_WIDTH, sizeFlat: [], hidden: [] };
    private rows: Axis = { count: 0, defaultSize: DEFAULT_CELL_HEIGHT, sizeFlat: [], hidden: [] };
    private showGridLines = true;

    /** 归一化阶段 */
    private cells: ExportCell[] = [];
    private cellsCursor = 0;

    /** 聚合阶段 */
    private groupCursor = 0;
    private byRow = new Map<number, ExportCell[]>();
    private mergeRefs: string[] = [];
    private maxCol = 1;
    private maxRow = 1;
    private rowHeights = new Map<number, number>();
    private rowNumbers: number[] = [];
    private rowCursor = 0;
    private rowsXml: string[] = [];
    private colSizes = new Map<number, number>();
    private colNumbers: number[] = [];

    private readonly registry = new StyleRegistry();

    public constructor(raw: unknown, name: string) {
        this.raw = raw;
        this.name = name;
    }

    public step(deadline: number): boolean {
        for (;;) {
            this.advance();
            if (this.phase === 'done') {
                this.progress = 100;
                return true;
            }
            if (performance.now() >= deadline) return false;
        }
    }

    /** 推进一个单元（分片 / 一批单元格 / 一批行 / 收尾整段） */
    private advance(): void {
        switch (this.phase) {
            case 'parse': this.advanceParse(); break;
            case 'cells': this.advanceCells(); break;
            case 'group': this.advanceGroup(); break;
            case 'rows': this.advanceRows(); break;
            case 'finalize': this.advanceFinalize(); break;
            default: break;
        }
    }

    /**
     * 解析阶段：先解析存档本体（v3 文本是一次整段 `JSON.parse`），再逐片解析 `chunks`
     *
     * 分片是独立 JSON 数组，天然可以「一片一个单元」地切分，因此进度按分片数推进。
     */
    private advanceParse(): void {
        if (!this.headerParsed) {
            this.headerParsed = true;
            if (typeof this.raw === 'string') {
                if (this.raw && this.raw !== '{}') {
                    try {
                        const parsed = JSON.parse(this.raw);
                        if (parsed && typeof parsed === 'object') this.snapshot = parsed as Record<string, unknown>;
                    } catch {
                        this.snapshot = {};
                    }
                }
            } else if (this.raw && typeof this.raw === 'object') {
                this.snapshot = this.raw as Record<string, unknown>;
            }
            this.styles = Array.isArray(this.snapshot.styles)
                ? this.snapshot.styles as Record<string, unknown>[]
                : [];
            if (Array.isArray(this.snapshot.values)) {
                this.rawCells = this.snapshot.values as Record<string, unknown>[];
            } else if (Array.isArray(this.snapshot.chunks)) {
                this.chunks = this.snapshot.chunks;
            }
        }

        if (this.chunkIndex < this.chunks.length) {
            const chunk = this.chunks[this.chunkIndex];
            if (typeof chunk === 'string') {
                const text = chunk.trim();
                if (text && text !== '[]') {
                    try {
                        const parsed = JSON.parse(text);
                        if (Array.isArray(parsed)) {
                            for (const item of parsed) {
                                if (item && typeof item === 'object') this.rawCells.push(item as Record<string, unknown>);
                            }
                        }
                    } catch {
                        // 损坏的分片：跳过，不让一个坏片毁掉整份导出
                    }
                }
            }
            this.chunkIndex++;
            this.progress = PROGRESS_PARSE * this.chunkIndex / this.chunks.length;
            return;
        }

        this.cols = this.snapshot.cols !== undefined
            ? normalizeAxis(this.snapshot.cols, DEFAULT_CELL_WIDTH, MAX_COL)
            : axisFromHeaders(this.snapshot.colHeaders, DEFAULT_CELL_WIDTH, 'width', MAX_COL);
        this.rows = this.snapshot.rows !== undefined
            ? normalizeAxis(this.snapshot.rows, DEFAULT_CELL_HEIGHT, MAX_ROW)
            : axisFromHeaders(this.snapshot.rowHeaders, DEFAULT_CELL_HEIGHT, 'height', MAX_ROW);
        this.showGridLines = this.snapshot.showGridLines !== false;
        this.progress = PROGRESS_PARSE;
        this.phase = 'cells';
    }

    /** 归一化阶段：紧凑单元格 → {@link ExportCell}（每片处理 {@link CELLS_PER_UNIT} 个） */
    private advanceCells(): void {
        const total = this.rawCells.length;
        if (total === 0) {
            this.progress = PROGRESS_CELLS;
            this.phase = 'group';
            return;
        }
        const end = Math.min(total, this.cellsCursor + CELLS_PER_UNIT);
        for (; this.cellsCursor < end; this.cellsCursor++) {
            const rawCell = this.rawCells[this.cellsCursor];
            if (!rawCell || typeof rawCell !== 'object') continue;
            const ref = parseRef(String(rawCell.c ?? ''));
            if (!ref) continue;
            const index = rawCell.s;
            const entry = typeof index === 'number' && this.styles[index] ? this.styles[index] : undefined;
            const font = (entry?.f ?? rawCell.f ?? firstCharFont(rawCell)) as Record<string, unknown> | null;
            const style = readStyle(entry, font ? { ...rawCell, f: font } : rawCell);
            const text = cellText(rawCell);
            const colspan = typeof rawCell.cs === 'number' && rawCell.cs > 1 ? rawCell.cs : 1;
            const rowspan = typeof rawCell.rs === 'number' && rawCell.rs > 1 ? rawCell.rs : 1;
            this.cells.push({ col: ref.col, row: ref.row, text, colspan, rowspan, style });
        }
        this.progress = PROGRESS_PARSE + (PROGRESS_CELLS - PROGRESS_PARSE) * this.cellsCursor / total;
        if (this.cellsCursor >= total) this.phase = 'group';
    }

    /**
     * 聚合阶段：按行分组、收集合并区、统计最大行列
     *
     * 处理完最后一批单元格后，顺带把行高 / 列宽 / 行号 / 列号集合算好 ——
     * 这些集合只依赖轴信息，避免在逐行生成时反复过滤。
     */
    private advanceGroup(): void {
        const total = this.cells.length;
        const end = Math.min(total, this.groupCursor + CELLS_PER_UNIT);
        for (; this.groupCursor < end; this.groupCursor++) {
            const cell = this.cells[this.groupCursor];
            if (cell.col > MAX_COL || cell.row > MAX_ROW) continue;
            this.maxCol = Math.max(this.maxCol, cell.col);
            this.maxRow = Math.max(this.maxRow, cell.row);
            const list = this.byRow.get(cell.row);
            if (list) list.push(cell);
            else this.byRow.set(cell.row, [cell]);
            if (cell.colspan > 1 || cell.rowspan > 1) {
                this.mergeRefs.push(`${colName(cell.col)}${cell.row}:${colName(cell.col + cell.colspan - 1)}${cell.row + cell.rowspan - 1}`);
            }
        }
        this.progress = PROGRESS_CELLS + (PROGRESS_GROUP - PROGRESS_CELLS) * this.groupCursor / Math.max(1, total);
        if (this.groupCursor < total) return;

        // 自定义行高：让「有内容行」以外的空行也能保留用户设置的高度
        for (let i = 0; i + 1 < this.rows.sizeFlat.length; i += 2) {
            const index = this.rows.sizeFlat[i];
            const size = this.rows.sizeFlat[i + 1];
            if (index >= 0 && index < MAX_ROW) this.rowHeights.set(index + 1, size);
        }
        for (const row of this.rowHeights.keys()) this.maxRow = Math.max(this.maxRow, row);

        // 列：只为「非默认宽度」与隐藏列写 <col>
        for (let i = 0; i + 1 < this.cols.sizeFlat.length; i += 2) {
            const index = this.cols.sizeFlat[i];
            const size = this.cols.sizeFlat[i + 1];
            if (index >= 0 && index < MAX_COL) this.colSizes.set(index + 1, size);
        }

        this.rowNumbers = Array.from(new Set([
            ...this.byRow.keys(),
            ...this.rowHeights.keys(),
            ...this.rows.hidden.map((index) => index + 1),
        ])).filter((row) => row >= 1 && row <= MAX_ROW).sort((a, b) => a - b);
        this.colNumbers = Array.from(new Set([
            ...this.colSizes.keys(),
            ...this.cols.hidden.map((index) => index + 1),
        ])).filter((col) => col >= 1 && col <= MAX_COL).sort((a, b) => a - b);
        this.progress = PROGRESS_GROUP;
        this.phase = 'rows';
    }

    /** 逐行阶段：每片生成 {@link ROWS_PER_UNIT} 行 XML */
    private advanceRows(): void {
        const total = this.rowNumbers.length;
        if (this.rowCursor >= total) {
            this.progress = PROGRESS_ROWS;
            this.phase = 'finalize';
            return;
        }
        const end = Math.min(total, this.rowCursor + ROWS_PER_UNIT);
        for (; this.rowCursor < end; this.rowCursor++) {
            const rowNumber = this.rowNumbers[this.rowCursor];
            const list = (this.byRow.get(rowNumber) ?? []).sort((a, b) => a.col - b.col);
            const height = this.rowHeights.get(rowNumber);
            const hidden = this.rows.hidden.includes(rowNumber - 1);
            const attrs = [`r="${rowNumber}"`]
                .concat(height !== undefined ? [`ht="${heightPoints(height)}"`, 'customHeight="1"'] : [])
                .concat(hidden ? ['hidden="1"'] : [])
                .join(' ');
            const xmlCells: string[] = [];
            for (const cell of list) {
                const ref = `${colName(cell.col)}${rowNumber}`;
                const styleId = this.registry.xfId(cell.style);
                const sAttr = styleId > 0 ? ` s="${styleId}"` : '';
                const number = numericValue(cell.text);
                if (number !== null) {
                    xmlCells.push(`<c r="${ref}"${sAttr}><v>${number}</v></c>`);
                } else if (cell.text) {
                    xmlCells.push(`<c r="${ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${escapeCellText(cell.text)}</t></is></c>`);
                } else if (styleId > 0) {
                    xmlCells.push(`<c r="${ref}"${sAttr}/>`);
                }
            }
            this.rowsXml.push(`<row ${attrs}>${xmlCells.join('')}</row>`);
        }
        this.progress = PROGRESS_GROUP + (PROGRESS_ROWS - PROGRESS_GROUP) * this.rowCursor / Math.max(1, total);
        if (this.rowCursor >= total) {
            this.progress = PROGRESS_ROWS;
            this.phase = 'finalize';
        }
    }

    /**
     * 收尾阶段：拼 sheet1.xml、样式表，编码并打包成 ZIP
     *
     * 这一段是整段同步的（`TextEncoder` + CRC32 + Blob 组装无法中断），
     * 通常几十毫秒，因此放在进度 90 → 100 的尾部一次性完成。
     */
    private advanceFinalize(): void {
        const colsXml: string[] = [];
        for (const col of this.colNumbers) {
            const size = this.colSizes.get(col);
            const hidden = this.cols.hidden.includes(col - 1);
            const attrs = [`min="${col}"`, `max="${col}"`]
                .concat(size !== undefined ? [`width="${widthChars(size)}"`, 'customWidth="1"'] : [])
                .concat(hidden ? ['hidden="1"'] : [])
                .join(' ');
            colsXml.push(`<col ${attrs}/>`);
        }

        const sheetView = this.showGridLines ? '' : '<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>';
        const cols = colsXml.length ? `<cols>${colsXml.join('')}</cols>` : '';
        const merges = this.mergeRefs.length
            ? `<mergeCells count="${this.mergeRefs.length}">${this.mergeRefs.map((ref) => `<mergeCell ref="${ref}"/>`).join('')}</mergeCells>`
            : '';
        const sheetXml = XML_HEAD
            + `<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">`
            + `<dimension ref="A1:${colName(this.maxCol)}${this.maxRow}"/>`
            + sheetView
            + `<sheetFormatPr defaultColWidth="${widthChars(this.cols.defaultSize)}" defaultRowHeight="${heightPoints(this.rows.defaultSize)}"/>`
            + cols
            + `<sheetData>${this.rowsXml.join('')}</sheetData>`
            + merges
            + '</worksheet>';

        const encoder = new TextEncoder();
        const entries: ZipEntryInput[] = [
            { name: '[Content_Types].xml', data: encoder.encode(CONTENT_TYPES) },
            { name: '_rels/.rels', data: encoder.encode(PACKAGE_RELS) },
            { name: 'xl/workbook.xml', data: encoder.encode(buildWorkbookXml(this.name)) },
            { name: 'xl/_rels/workbook.xml.rels', data: encoder.encode(WORKBOOK_RELS) },
            { name: 'xl/styles.xml', data: encoder.encode(this.registry.toXml()) },
            { name: 'xl/worksheets/sheet1.xml', data: encoder.encode(sheetXml) },
        ];
        this.blob = new Blob([createZipBlob(entries)], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        this.progress = 100;
        this.phase = 'done';
    }
}

/**
 * 创建一个可暂停 / 可续传的 xlsx 转换作业
 * @param {unknown} raw - IndexedDB 里的存档（v4 对象 / v3 文本 / `'{}'`）
 * @param {string} name - 文件名（用于工作表名，不含扩展名）
 * @returns {XlsxTask} 由调用方按时间片驱动的作业
 */
export function createXlsxTask(raw: unknown, name: string): XlsxTask {
    return new XlsxJob(raw, name);
}
