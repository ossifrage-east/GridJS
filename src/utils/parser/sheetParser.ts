/**
 * 通用 SheetParser 组件
 * 提供解析 xls, xlsx 表格的功能，支持读取单元格数据、样式、图片等
 *
 * **全部使用原生浏览器 API 实现，不依赖任何第三方库：**
 *
 * ```
 *                    ┌──────────────┐
 *                    │  SheetParser │
 *                    └──────┬───────┘
 *              ┌────────────┴────────────┐
 *           .xlsx                      .xls
 *              │                         │
 *    ┌─────────▼─────────┐     ┌─────────▼─────────┐
 *    │ inflate + zip.ts  │     │  cfb.ts           │  OLE2 复合文档
 *    │ 原生 ZIP 解包      │     │  (FAT/DIFAT/目录树) │
 *    └─────────┬─────────┘     └─────────┬─────────┘
 *              │                         │
 *    ┌─────────▼─────────┐     ┌─────────▼─────────┐
 *    │ DOMParser 解析 XML │     │  biff.ts          │  BIFF8 记录流
 *    └─────────┬─────────┘     └─────────┬─────────┘
 *              └────────────┬────────────┘
 *                    ┌────────▼─────────┐
 *                    │  映射为项目格式   │  Cell / Char / Header
 *                    └──────────────────┘
 * ```
 *
 * 三步解析架构：
 * 1. `decode(file)`    — 解压/解码：xlsx 解包得到内部 XML；xls 定位 OLE2 中的 Workbook 流
 * 2. `extract(decoded)` — 读取：从解压产物中提取单元格数据、样式、图片等结构化信息
 * 3. `generate(extracted)` — 生成：映射为项目数据格式 Cell[]/ColHeader[]/RowHeader[]
 *
 * @module SheetParser
 * @author 东方鹗
 * @see
 * # B站: https://space.bilibili.com/194359739
 * # 知乎: https://www.zhihu.com/people/eastossifrage
 * # CSDN: https://blog.csdn.net/os373
 */

import { Cell } from '../../core/dataArchitecture/cell';
import { Char } from '../../core/dataArchitecture/char';
import type { ColHeader, RowHeader } from '../../core/dataArchitecture/header';
import { DEFAULT_CELL_WIDTH, DEFAULT_CELL_HEIGHT } from '../../core/constant';
import { ZipParser, ZipArchive } from './zip';
import { CompoundFile } from './cfb';
import { BiffParser, BiffXf, BiffFont, BiffWorkbook } from './biff';

// ── 中间数据结构 ──

/** 第一步输出：解压/解码后的工作簿 */
export interface DecodedWorkbook {
    /** 文件格式 */
    format: 'xlsx' | 'xls';
    /** xlsx 解压后的文件映射（路径 → 内容），xls 为 null */
    files: Map<string, Uint8Array> | null;
    /** xls：OLE2 复合文档中的 Workbook 流；xlsx 为 null */
    rawData: Uint8Array | null;
}

/** 单元格样式信息（从 Excel 样式表提取） */
export interface CellStyleInfo {
    font?: {
        name?: string;
        size?: number;
        bold?: boolean;
        italic?: boolean;
        color?: string;     // #RRGGBB
        underline?: boolean;
        strikethrough?: boolean;
    };
    fill?: {
        patternType?: string;
        fgColor?: string;   // #RRGGBB
    };
    alignment?: {
        horizontal?: 'left' | 'center' | 'right' | 'justify';
        vertical?: 'top' | 'middle' | 'bottom';
        wrapText?: boolean;
        indent?: number;
    };
    border?: {
        top?: { style: string; color: string };
        bottom?: { style: string; color: string };
        left?: { style: string; color: string };
        right?: { style: string; color: string };
    };
}

/** 第二步输出：提取的单个单元格信息 */
export interface ExtractedCell {
    /** 行号（0-based，源自 Excel） */
    row: number;
    /** 列号（0-based，源自 Excel） */
    col: number;
    /** 单元格值 */
    value: string | number | boolean | null;
    /** 值类型 */
    type: 'string' | 'number' | 'boolean' | 'date' | 'formula' | 'error' | 'blank';
    /** 样式信息 */
    style?: CellStyleInfo;
    /** 数字格式码 */
    numberFormat?: string;
}

/** 提取的图片信息 */
export interface ExtractedImage {
    /** 图片二进制数据 */
    data: Uint8Array;
    /** MIME 类型 */
    mimeType: string;
    /** 锚点位置（0-based 行列） */
    anchor: {
        col: number;
        row: number;
        colOffset: number;  // EMU 单位
        rowOffset: number;
    };
    /** 图片名称 */
    name?: string;
}

/** 第二步输出：提取的单个工作表 */
export interface ExtractedSheet {
    /** 工作表名称 */
    name: string;
    /** 所有单元格 */
    cells: ExtractedCell[];
    /** 合并区域 */
    merges: Array<{ s: { r: number; c: number }; e: { r: number; c: number } }>;
    /** 列宽信息（0-based 索引，像素） */
    cols: Array<{ index: number; width?: number; customWidth?: boolean }>;
    /** 行高信息（0-based 索引，像素） */
    rows: Array<{ index: number; height?: number; customHeight?: boolean }>;
    /** 图片 */
    images: ExtractedImage[];
    /** 数据范围 */
    range: { minRow: number; maxRow: number; minCol: number; maxCol: number };
}

/** 第三步输出：与项目数据格式一致的结果 */
export interface SheetParserResult {
    /** 单元格数组（项目 Cell 格式） */
    values: Cell[];
    /** 列头数组 */
    colHeaders: ColHeader[];
    /** 行头数组 */
    rowHeaders: RowHeader[];
    /** 图片列表（项目暂无图片基础设施，预留） */
    images: ExtractedImage[];
}

/** Excel 样式表（xlsx 的 styles.xml 解析中间结构） */
interface ExcelStyleSheet {
    numFmts: Map<number, string>;
    fonts: Array<{
        name?: string; size?: number; bold?: boolean; italic?: boolean;
        color?: string; underline?: boolean; strikethrough?: boolean;
    }>;
    fills: Array<{ patternType?: string; fgColor?: string }>;
    borders: Array<{
        left?: { style?: string; color?: string };
        right?: { style?: string; color?: string };
        top?: { style?: string; color?: string };
        bottom?: { style?: string; color?: string };
    }>;
    cellXfs: Array<{
        numFmtId: number; fontId: number; fillId: number; borderId: number;
        alignment?: {
            horizontal?: string; vertical?: string;
            wrapText?: boolean; indent?: number;
        };
        applyFont?: boolean; applyFill?: boolean;
        applyBorder?: boolean; applyAlignment?: boolean;
    }>;
}

/** SheetParser 工作簿解析选项 */
export interface SheetParserOptions {
    /** 读取第几个工作表（默认 0） */
    sheetIndex?: number;
    /** 是否读取样式（默认 true） */
    readStyles?: boolean;
    /** 是否读取图片（默认 true） */
    readImages?: boolean;
}

/**
 * 边框线型宽度表（xls 的 XF 记录中，边框以 3 位数值表示线型）
 *
 * 值 0 表示无边框，1-7 依次为不同线型。
 */
const BIFF_BORDER_STYLES = ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair'];

/** xls 填充图案索引 → xlsx patternType 名称 */
const BIFF_FILL_PATTERNS = [
    'none', 'solid', 'mediumGray', 'darkGray', 'lightGray', 'darkHorizontal',
    'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis',
    'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis',
    'gray125', 'gray0625'
];

export class SheetParser {

    // ════════════════════════════════════════════════════════════
    //  第一步：解压/解码
    // ════════════════════════════════════════════════════════════

    /**
     * 第一步：解压/解码 Excel 文件
     *
     * - **xlsx**：本质为 ZIP 压缩包，用自研 `ZipParser` 解包（压缩项经
     *   `inflateRaw` 解压，优先走浏览器原生 `DecompressionStream`）
     * - **xls**：OLE2 复合文档，用自研 `CompoundFile` 定位并读取
     *   `Workbook`（Excel 97+）或 `Book`（Excel 5/95）流
     *
     * @param {File} file - 用户选择的 .xlsx 或 .xls 文件
     * @returns {Promise<DecodedWorkbook>} 解压/解码后的工作簿
     * @throws {Error} 文件格式不支持或内容损坏
     * @example
     * const decoded = await parser.decode(file);
     */
    public async decode(file: File): Promise<DecodedWorkbook> {
        const extension = file.name.toLowerCase().split('.').pop() || '';
        const buffer = new Uint8Array(await file.arrayBuffer());

        if (extension === 'xlsx') {
            // —— xlsx：原生 ZIP 解包 ——
            const zip = new ZipParser(buffer);
            const archive: ZipArchive = zip.parseCentralDirectory();

            const files = new Map<string, Uint8Array>();
            for (const path of archive.entries.keys()) {
                // 目录条目（以 / 结尾）无需读取
                if (path.endsWith('/')) continue;
                const bytes = await zip.readEntry(archive, path);
                if (bytes) files.set(path, bytes);
            }
            return { format: 'xlsx', files, rawData: null };
        }

        if (extension === 'xls') {
            // —— xls：OLE2 复合文档中提取 Workbook 流 ——
            const cfb = new CompoundFile(buffer);
            // Excel 97+ 命名为 Workbook，Excel 5/95 命名为 Book
            const stream = cfb.readStream('Workbook') || cfb.readStream('Book');
            if (!stream) {
                throw new Error('xls 文件中未找到 Workbook 数据流，文件可能已损坏');
            }
            return { format: 'xls', files: null, rawData: stream };
        }

        throw new Error(`不支持的文件格式: .${extension}，仅支持 .xlsx 和 .xls`);
    }

    // ════════════════════════════════════════════════════════════
    //  第二步：读取单元格数据、样式、图片
    // ════════════════════════════════════════════════════════════

    /**
     * 第二步：从解压产物中读取单元格数据、样式、图片
     *
     * - **xlsx**：用浏览器原生 `DOMParser` 解析 XML，完整提取字体/填充/
     *   对齐/边框/数字格式/合并/图片
     * - **xls**：用自研 `BiffParser` 解析 BIFF 记录流，提取单元格记录、
     *   XF 样式表、FONT 字体表、合并区域、行列尺寸与 Escher BLIP 图片
     *
     * @param {DecodedWorkbook} decoded - 第一步的解压产物
     * @param {SheetParserOptions} [options] - 解析选项
     * @returns {Promise<ExtractedSheet[]>} 提取的工作表数组
     * @example
     * const sheets = await parser.extract(decoded, { sheetIndex: 0 });
     */
    public async extract(
        decoded: DecodedWorkbook,
        options: SheetParserOptions = {}
    ): Promise<ExtractedSheet[]> {
        const { readStyles = true, readImages = true } = options;

        if (decoded.format === 'xlsx' && decoded.files) {
            return this.extractXlsx(decoded.files, readStyles, readImages);
        }

        if (decoded.format === 'xls' && decoded.rawData) {
            return this.extractXls(decoded.rawData, readStyles, readImages);
        }

        throw new Error('解压产物无效');
    }

    /**
     * 解析 xlsx 的 XML 文件，完整提取单元格数据、样式、图片
     *
     * @param files - 解压后的文件映射
     * @param readStyles - 是否读取样式
     * @param readImages - 是否读取图片
     * @returns 工作表数组
     * @private
     */
    private async extractXlsx(
        files: Map<string, Uint8Array>,
        readStyles: boolean,
        readImages: boolean
    ): Promise<ExtractedSheet[]> {
        // —— 解析 workbook.xml 获取工作表名称与 r:id ——
        const workbookXml = this.readXml(files, 'xl/workbook.xml');
        const sheetInfos: Array<{ name: string; rId: string }> = [];
        if (workbookXml) {
            const sheets = workbookXml.getElementsByTagName('sheet');
            for (let i = 0; i < sheets.length; i++) {
                const sheet = sheets[i];
                sheetInfos.push({
                    name: sheet.getAttribute('name') || '',
                    rId: sheet.getAttribute('r:id') ||
                        sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') ||
                        ''
                });
            }
        }

        // —— 解析 workbook.xml.rels 获取 r:id → sheetN.xml 路径 ——
        const relsXml = this.readXml(files, 'xl/_rels/workbook.xml.rels');
        const sheetPaths = new Map<string, string>();
        const drawingRelPaths = new Map<string, string>();
        if (relsXml) {
            const rels = relsXml.getElementsByTagName('Relationship');
            for (let i = 0; i < rels.length; i++) {
                const rel = rels[i];
                const id = rel.getAttribute('Id') || '';
                const target = rel.getAttribute('Target') || '';
                if (target.includes('worksheets/')) {
                    sheetPaths.set(id, target.startsWith('/') ? target.slice(1) : `xl/${target}`);
                } else if (target.includes('drawing')) {
                    drawingRelPaths.set(id, target.startsWith('/') ? target.slice(1) : `xl/${target}`);
                }
            }
        }

        // —— 解析样式表 styles.xml ——
        let styleSheet: ExcelStyleSheet | null = null;
        if (readStyles) {
            const stylesXml = this.readXml(files, 'xl/styles.xml');
            if (stylesXml) {
                styleSheet = this.parseStyleSheet(stylesXml);
            }
        }

        // —— 解析共享字符串表 sharedStrings.xml ——
        const sharedStrings: string[] = [];
        const ssXml = this.readXml(files, 'xl/sharedStrings.xml');
        if (ssXml) {
            const sis = ssXml.getElementsByTagName('si');
            for (let i = 0; i < sis.length; i++) {
                sharedStrings.push(this.extractInlineText(sis[i]));
            }
        }

        // —— 遍历每个工作表 ——
        const result: ExtractedSheet[] = [];
        for (const info of sheetInfos) {
            const sheetPath = sheetPaths.get(info.rId) || 'xl/worksheets/sheet1.xml';
            const sheetXml = this.readXml(files, sheetPath);
            if (!sheetXml) continue;

            const sheet = this.parseWorksheet(
                sheetXml, sharedStrings, styleSheet, info.name
            );

            // —— 提取图片：由工作表 rels 中的 drawing 关系精确定位 ——
            if (readImages) {
                sheet.images = this.extractSheetImages(files, sheetPath, drawingRelPaths);
            }

            result.push(sheet);
        }

        return result;
    }

    /**
     * 解析 xls 的 BIFF 记录流（自研实现，无第三方库）
     *
     * 流程：BIFF 全局记录（BOUNDSHEET/FONT/FORMAT/XF/SST）
     *      → 按 BOUNDSHEET 偏移逐表扫描子流（单元格/MERGEDCELLS/COLINFO/ROW）
     *
     * @param rawData - Workbook 流的字节数据
     * @param readStyles - 是否读取样式
     * @param readImages - 是否读取图片
     * @returns 工作表数组
     * @private
     */
    private extractXls(
        rawData: Uint8Array,
        readStyles: boolean,
        readImages: boolean
    ): ExtractedSheet[] {
        const biff = new BiffParser(rawData);
        const wb: BiffWorkbook = biff.parse();

        return wb.sheets.map(sheet => {
            // 图表工作表无单元格数据，返回空表
            if (sheet.isChart) {
                return {
                    name: sheet.name, cells: [], merges: [],
                    cols: [], rows: [], images: [],
                    range: { minRow: 0, maxRow: 0, minCol: 0, maxCol: 0 }
                };
            }

            const raw = biff.parseSheet(sheet);
            const cells: ExtractedCell[] = [];

            for (const cell of raw.cells) {
                const style = readStyles
                    ? this.biffStyleToCellStyle(wb, cell.styleIndex)
                    : undefined;
                const numberFormat = readStyles
                    ? this.biffNumberFormat(wb, cell.styleIndex)
                    : undefined;

                // 空单元格：无值且无任何样式时可跳过
                if (cell.type === 'blank' && !style && !numberFormat) continue;

                cells.push({
                    row: cell.row,
                    col: cell.col,
                    value: cell.type === 'number' ? (cell.numberValue ?? null)
                        : cell.type === 'string' ? (cell.stringValue ?? '')
                            : cell.type === 'boolean' ? (cell.booleanValue ?? false)
                                : cell.type === 'error' ? (cell.errorValue ?? null)
                                    : null,
                    type: cell.type === 'blank' ? 'blank'
                        : cell.type === 'error' ? 'error' : cell.type,
                    style,
                    numberFormat
                });
            }

            // 列宽：BIFF 以「字符宽」为单位，换算为像素（公式同 Excel）
            const cols = raw.cols.map(col => ({
                index: col.index,
                width: col.width !== undefined ? Math.round(col.width * 7 + 5) : undefined,
                customWidth: col.width !== undefined
            }));

            // 行高：BIFF 的 twips 已由 BiffParser 换算为磅，此处转为像素（磅 × 96/72）
            const rows = raw.rows.map(row => ({
                index: row.index,
                height: row.height !== undefined ? Math.round(row.height * 96 / 72) : undefined,
                customHeight: row.height !== undefined
            }));

            // 图片：Escher BLIP → ExtractedImage（xls 的锚点信息不可靠，置为 0）
            const images = readImages
                ? raw.images.map(img => ({
                    data: img.data,
                    mimeType: img.mimeType,
                    anchor: img.anchor || { col: 0, row: 0, colOffset: 0, rowOffset: 0 }
                }))
                : [];

            return {
                name: sheet.name,
                cells,
                merges: raw.merges,
                cols,
                rows,
                images,
                range: raw.range
            };
        });
    }

    /**
     * 将 xls 的 XF 样式索引转换为项目的 CellStyleInfo
     *
     * @param wb - BIFF 工作簿（提供 XF 表、FONT 表、调色板）
     * @param styleIndex - 单元格记录的 XF 索引
     * @returns 样式信息；索引无效时返回 undefined
     * @private
     */
    private biffStyleToCellStyle(wb: BiffWorkbook, styleIndex: number): CellStyleInfo | undefined {
        // XF 索引可能越界（部分文件写入器不完整），需做防御
        const xf: BiffXf | undefined = wb.xfs[styleIndex];
        if (!xf) return undefined;

        const style: CellStyleInfo = {};

        // —— 字体 ——
        const font: BiffFont | undefined = wb.fonts[xf.fontIndex];
        if (font) {
            style.font = {
                name: font.name,
                size: font.height || undefined,
                bold: font.bold || undefined,
                italic: font.italic || undefined,
                // 调色板索引 → RGB
                color: this.paletteColor(wb.palette, font.colorIndex),
                underline: font.underline || undefined,
                strikethrough: font.strikethrough || undefined
            };
        }

        // —— 填充 ——
        // BIFF 中 patternType 0（none）与 17/18（灰度底纹）视为无背景色
        const fillPatternName = BIFF_FILL_PATTERNS[xf.fillPattern];
        if (xf.fillPattern !== 0 && fillPatternName &&
            fillPatternName !== 'gray125' && fillPatternName !== 'gray0625') {
            const fg = this.paletteColor(wb.palette, xf.fillColorIndex);
            const bg = this.paletteColor(wb.palette, xf.fillBackgroundIndex);
            style.fill = {
                patternType: fillPatternName,
                fgColor: fg || bg
            };
        }

        // —— 对齐 ——
        // BIFF 水平对齐：0 常规 1 居中 2 填充 3 两端 4 跨列居中 5 分散 6 单元格内对齐
        // BIFF 垂直对齐：0 底对齐 1 居中
        const hAlignMap: Record<number, NonNullable<CellStyleInfo['alignment']>['horizontal']> = {
            1: 'center', 3: 'justify', 5: 'justify', 7: 'center'
        };
        const vAlignMap: Record<number, NonNullable<CellStyleInfo['alignment']>['vertical']> = {
            0: 'bottom', 1: 'middle'
        };
        const horizontal = xf.horizontalAlignment === 2 ? 'left' : hAlignMap[xf.horizontalAlignment];
        const vertical = vAlignMap[xf.verticalAlignment];
        if (horizontal || vertical || xf.wrapText || xf.indent) {
            style.alignment = {
                horizontal: horizontal || 'left',
                vertical: vertical || 'bottom',
                wrapText: xf.wrapText || undefined,
                indent: xf.indent || undefined
            };
        }

        // —— 边框 ——
        const borderColor = this.paletteColor(wb.palette, xf.borderColorIndex) || '#000000';
        const border: NonNullable<CellStyleInfo['border']> = {};
        const topStyle = BIFF_BORDER_STYLES[xf.borderTop];
        const bottomStyle = BIFF_BORDER_STYLES[xf.borderBottom];
        const leftStyle = BIFF_BORDER_STYLES[xf.borderLeft];
        const rightStyle = BIFF_BORDER_STYLES[xf.borderRight];
        if (topStyle && topStyle !== 'none') border.top = { style: topStyle, color: borderColor };
        if (bottomStyle && bottomStyle !== 'none') border.bottom = { style: bottomStyle, color: borderColor };
        if (leftStyle && leftStyle !== 'none') border.left = { style: leftStyle, color: borderColor };
        if (rightStyle && rightStyle !== 'none') border.right = { style: rightStyle, color: borderColor };
        if (Object.keys(border).length > 0) style.border = border;

        return Object.keys(style).length > 0 ? style : undefined;
    }

    /**
     * 读取 xls 单元格样式对应的数字格式码
     *
     * @param wb - BIFF 工作簿
     * @param styleIndex - XF 索引
     * @returns 格式码；为内置 General 格式时返回 undefined
     * @private
     */
    private biffNumberFormat(wb: BiffWorkbook, styleIndex: number): string | undefined {
        const xf = wb.xfs[styleIndex];
        if (!xf) return undefined;
        const formatIndex = xf.formatIndex;
        // 0 是内置的 General 格式，无需保留
        if (formatIndex === 0) return undefined;
        return wb.formats.get(formatIndex);
    }

    /**
     * 调色板索引转 RGB 颜色
     *
     * @param palette - 调色板
     * @param index - 颜色索引
     * @returns `#RRGGBB` 颜色；索引无效时返回 undefined
     * @private
     */
    private paletteColor(palette: string[], index: number): string | undefined {
        if (index <= 0) return undefined;
        const color = palette[index];
        return color || undefined;
    }

    // ════════════════════════════════════════════════════════════
    //  第三步：生成与项目数据格式一致的数据
    // ════════════════════════════════════════════════════════════

    /**
     * 第三步：将提取的数据映射为项目 Cell[]/ColHeader[]/RowHeader[] 格式
     *
     * 映射规则：
     * - Excel 0-based 行列 → 项目 1-based（col=0→1=A, row=0→1）
     * - Excel 单元格样式 → Cell 属性（backgroundColor/borderWidth/textAlign/alignItems/numberFormat）
     * - Excel 字体样式 → Char 属性（fontFamily/fontSize/fontWeight/fontStyle/fontColor/underline/strikethrough）
     * - Excel 合并区域 → Cell.colspan/rowspan（左上角为锚点）
     * - Excel 列宽 → ColHeader.width（像素）
     * - Excel 行高 → RowHeader.height（像素）
     *
     * @param {ExtractedSheet[]} sheets - 第二步提取的工作表
     * @param {SheetParserOptions} [options] - 选项
     * @returns {SheetParserResult} 与项目格式一致的数据
     * @example
     * const result = parser.generate(sheets, { sheetIndex: 0 });
     */
    public generate(
        sheets: ExtractedSheet[],
        options: SheetParserOptions = {}
    ): SheetParserResult {
        const { sheetIndex = 0 } = options;
        const sheet = sheets[sheetIndex] || sheets[0];
        if (!sheet) {
            return { values: [], colHeaders: [], rowHeaders: [], images: [] };
        }

        // —— 构建合并映射：左上角单元格 → { rowspan, colspan } ——
        const mergeMap = new Map<string, { rowspan: number; colspan: number }>();
        for (const merge of sheet.merges) {
            const rowspan = merge.e.r - merge.s.r + 1;
            const colspan = merge.e.c - merge.s.c + 1;
            if (rowspan > 1 || colspan > 1) {
                mergeMap.set(`${merge.s.r},${merge.s.c}`, { rowspan, colspan });
            }
        }

        // —— 转换单元格 ——
        const values: Cell[] = [];
        for (const extCell of sheet.cells) {
            const col1 = extCell.col + 1;  // 0-based → 1-based
            const row1 = extCell.row + 1;
            const cellName = this.colToName(col1) + row1;

            const cell = new Cell();
            cell.cell = cellName;

            // 合并信息
            const merge = mergeMap.get(`${extCell.row},${extCell.col}`);
            if (merge) {
                if (merge.colspan > 1) cell.colspan = merge.colspan;
                if (merge.rowspan > 1) cell.rowspan = merge.rowspan;
            }

            // 值 → Char[]（每个字符一个 Char，共享同一字体样式）
            // 布尔值转为 Excel 风格的 "TRUE"/"FALSE" 文本显示
            let text: string;
            if (typeof extCell.value === 'boolean') {
                text = extCell.value ? 'TRUE' : 'FALSE';
            } else if (extCell.value === null || extCell.value === undefined) {
                text = '';
            } else {
                text = String(extCell.value);
            }

            if (text) {
                const chars: Char[] = [];
                for (const ch of Array.from(text)) {
                    const char = new Char();
                    char.char = ch;
                    // 字体样式映射
                    const f = extCell.style?.font;
                    if (f) {
                        if (f.name) char.fontFamily = f.name;
                        if (f.size) char.fontSize = f.size;
                        if (f.bold) char.fontWeight = true;
                        if (f.italic) char.fontStyle = true;
                        if (f.color) char.fontColor = f.color;
                        if (f.underline) char.underline = true;
                        if (f.strikethrough) char.strikethrough = true;
                    }
                    chars.push(char);
                }
                cell.chars = chars;
            }

            // 单元格级样式映射
            const s = extCell.style;
            if (s) {
                // 背景色
                if (s.fill?.fgColor) {
                    cell.backgroundColor = s.fill.fgColor;
                }
                // 对齐
                if (s.alignment?.horizontal) {
                    cell.textAlign = s.alignment.horizontal;
                }
                if (s.alignment?.vertical) {
                    cell.alignItems = s.alignment.vertical;
                }
                if (s.alignment?.wrapText) {
                    cell.wrap = true;
                }
                // 边框：Excel border style → 项目 border width
                if (s.border) {
                    if (s.border.top) cell.borderTopWidth = this.borderStyleToWidth(s.border.top.style);
                    if (s.border.bottom) cell.borderBottomWidth = this.borderStyleToWidth(s.border.bottom.style);
                    if (s.border.left) cell.borderLeftWidth = this.borderStyleToWidth(s.border.left.style);
                    if (s.border.right) cell.borderRightWidth = this.borderStyleToWidth(s.border.right.style);
                    // 边框颜色（取第一个有颜色的边框）
                    const borderColor = s.border.top?.color || s.border.bottom?.color
                        || s.border.left?.color || s.border.right?.color;
                    if (borderColor) cell.borderColor = borderColor;
                }
            }

            // 数字格式
            if (extCell.numberFormat && extCell.numberFormat !== 'General') {
                cell.numberFormat = extCell.numberFormat;
            }

            values.push(cell);
        }

        // —— 列头：生成从 1 到 maxCol 的 ColHeader[] ——
        const maxCol = sheet.range.maxCol + 1;  // 1-based
        const colWidthMap = new Map<number, number>();
        for (const col of sheet.cols) {
            if (col.width) colWidthMap.set(col.index + 1, col.width);
        }
        const colHeaders: ColHeader[] = [];
        let colLeft = 0;
        for (let c = 1; c <= maxCol; c++) {
            const width = colWidthMap.get(c) ?? DEFAULT_CELL_WIDTH;
            colHeaders.push({ left: colLeft, width });
            colLeft += width;
        }

        // —— 行头：生成从 1 到 maxRow 的 RowHeader[] ——
        const maxRow = sheet.range.maxRow + 1;
        const rowHeightMap = new Map<number, number>();
        for (const row of sheet.rows) {
            if (row.height) rowHeightMap.set(row.index + 1, row.height);
        }
        const rowHeaders: RowHeader[] = [];
        let rowTop = 0;
        for (let r = 1; r <= maxRow; r++) {
            const height = rowHeightMap.get(r) ?? DEFAULT_CELL_HEIGHT;
            rowHeaders.push({ top: rowTop, height });
            rowTop += height;
        }

        return { values, colHeaders, rowHeaders, images: sheet.images };
    }

    /**
     * 一站式解析：decode → extract → generate
     *
     * @param {File} file - Excel 文件
     * @param {SheetParserOptions} [options] - 解析选项
     * @returns {Promise<SheetParserResult>} 与项目格式一致的数据
     * @example
     * const result = await parser.parse(file);
     */
    public async parse(file: File, options?: SheetParserOptions): Promise<SheetParserResult> {
        const decoded = await this.decode(file);
        const sheets = await this.extract(decoded, options);
        return this.generate(sheets, options);
    }

    // ════════════════════════════════════════════════════════════
    //  XML 解析辅助方法（xlsx 路径）
    // ════════════════════════════════════════════════════════════

    /**
     * 从解压文件中读取并解析 XML
     *
     * @param files - 解压后的文件映射
     * @param path - XML 文件路径
     * @returns Document 或 null
     * @private
     */
    private readXml(files: Map<string, Uint8Array>, path: string): Document | null {
        const data = files.get(path);
        if (!data) return null;

        let text: string;
        if (typeof TextDecoder !== 'undefined') {
            text = new TextDecoder('utf-8').decode(data);
        } else {
            // 老旧环境兜底：手动 UTF-8 解码
            let s = '';
            for (let i = 0; i < data.length; i++) s += String.fromCharCode(data[i]);
            text = decodeURIComponent(escape(s));
        }

        const parser = new DOMParser();
        const doc = parser.parseFromString(text, 'application/xml');

        // 解析错误时返回 null，避免后续遍历空文档
        if (doc.getElementsByTagName('parsererror').length > 0) {
            return null;
        }
        return doc;
    }

    /**
     * 提取元素的全部文本内容（合并 `<t>` 与富文本 `<r><t>` 结构）
     *
     * @param element - 元素节点
     * @returns 文本内容
     * @private
     */
    private extractInlineText(element: Element): string {
        let text = '';
        const nodes = element.getElementsByTagName('t');
        for (let i = 0; i < nodes.length; i++) {
            text += nodes[i].textContent || '';
        }
        return text;
    }

    /**
     * 解析 styles.xml 样式表
     *
     * @param xml - styles.xml 的 Document
     * @returns 样式表结构
     * @private
     */
    private parseStyleSheet(xml: Document): ExcelStyleSheet {
        const result: ExcelStyleSheet = {
            numFmts: new Map(), fonts: [], fills: [], borders: [], cellXfs: []
        };

        // — 数字格式 —
        const numFmtNodes = xml.getElementsByTagName('numFmt');
        for (let i = 0; i < numFmtNodes.length; i++) {
            const id = parseInt(numFmtNodes[i].getAttribute('numFmtId') || '0', 10);
            const code = numFmtNodes[i].getAttribute('formatCode') || '';
            result.numFmts.set(id, code);
        }

        // — 字体 —
        const fontNodes = xml.getElementsByTagName('font');
        for (let i = 0; i < fontNodes.length; i++) {
            const fontEl = fontNodes[i];
            // 粗体/斜体可能是 <b/> 或 <b val="1"/>
            const boldEl = this.firstChild(fontEl, 'b');
            const italicEl = this.firstChild(fontEl, 'i');
            result.fonts.push({
                name: this.childAttr(fontEl, 'name', 'val') || undefined,
                size: parseFloat(this.childAttr(fontEl, 'sz', 'val') || '0') || undefined,
                bold: boldEl ? this.isOn(boldEl.getAttribute('val')) : false,
                italic: italicEl ? this.isOn(italicEl.getAttribute('val')) : false,
                color: this.childRgb(fontEl, 'color'),
                underline: !!this.firstChild(fontEl, 'u'),
                strikethrough: !!this.firstChild(fontEl, 'strike')
            });
        }

        // — 填充 —
        const fillNodes = xml.getElementsByTagName('fill');
        for (let i = 0; i < fillNodes.length; i++) {
            const pf = this.firstChild(fillNodes[i], 'patternFill');
            result.fills.push({
                patternType: pf?.getAttribute('patternType') || undefined,
                fgColor: pf ? this.childRgb(pf, 'fgColor') : undefined
            });
        }

        // — 边框 —
        const borderNodes = xml.getElementsByTagName('border');
        for (let i = 0; i < borderNodes.length; i++) {
            const borderEl = borderNodes[i];
            result.borders.push({
                left: this.parseBorderSide(borderEl, 'left'),
                right: this.parseBorderSide(borderEl, 'right'),
                top: this.parseBorderSide(borderEl, 'top'),
                bottom: this.parseBorderSide(borderEl, 'bottom')
            });
        }

        // — 单元格样式交叉引用 (cellXfs) —
        const xfsNode = xml.getElementsByTagName('cellXfs');
        if (xfsNode.length > 0) {
            const xfNodes = xfsNode[0].getElementsByTagName('xf');
            for (let i = 0; i < xfNodes.length; i++) {
                const xf = xfNodes[i];
                const align = this.firstChild(xf, 'alignment');
                result.cellXfs.push({
                    numFmtId: parseInt(xf.getAttribute('numFmtId') || '0', 10),
                    fontId: parseInt(xf.getAttribute('fontId') || '0', 10),
                    fillId: parseInt(xf.getAttribute('fillId') || '0', 10),
                    borderId: parseInt(xf.getAttribute('borderId') || '0', 10),
                    applyFont: this.isOn(xf.getAttribute('applyFont')),
                    applyFill: this.isOn(xf.getAttribute('applyFill')),
                    applyBorder: this.isOn(xf.getAttribute('applyBorder')),
                    applyAlignment: this.isOn(xf.getAttribute('applyAlignment')),
                    alignment: align ? {
                        horizontal: align.getAttribute('horizontal') || undefined,
                        vertical: align.getAttribute('vertical') || undefined,
                        wrapText: this.isOn(align.getAttribute('wrapText')),
                        indent: align.getAttribute('indent')
                            ? parseInt(align.getAttribute('indent')!, 10) : undefined
                    } : undefined
                });
            }
        }

        return result;
    }

    /**
     * 解析边框的某一边
     *
     * @param parent - border 元素
     * @param tag - 边的标签名（left/right/top/bottom）
     * @returns 边信息；不存在时 undefined
     * @private
     */
    private parseBorderSide(
        parent: Element, tag: string
    ): { style?: string; color?: string } | undefined {
        const side = this.firstChild(parent, tag);
        if (!side) return undefined;
        return {
            style: side.getAttribute('style') || undefined,
            color: this.childRgb(side, 'color')
        };
    }

    /**
     * 获取元素的第一个直接子元素
     *
     * 注意 `getElementsByTagName` 会返回**所有后代**，
     * 而样式表存在嵌套结构（如 border 内含 color），故需校验直接子级关系。
     *
     * @param parent - 父元素
     * @param tag - 子元素标签名
     * @returns 子元素；不存在时 null
     * @private
     */
    private firstChild(parent: Element, tag: string): Element | null {
        const children = parent.children;
        for (let i = 0; i < children.length; i++) {
            // 兼容带命名空间的标签（用 localName 比较）
            const localName = children[i].localName || children[i].nodeName;
            if (localName === tag || children[i].nodeName === tag) return children[i];
        }
        return null;
    }

    /**
     * 获取子元素的属性值
     *
     * @param parent - 父元素
     * @param tag - 子元素标签名
     * @param attr - 属性名
     * @returns 属性值；不存在时 null
     * @private
     */
    private childAttr(parent: Element, tag: string, attr: string): string | null {
        const child = this.firstChild(parent, tag);
        return child ? child.getAttribute(attr) : null;
    }

    /**
     * 获取子元素的 rgb 属性并转为 `#RRGGBB`
     *
     * @param parent - 父元素
     * @param tag - 子元素标签名
     * @returns 颜色；不存在或为自动色/主题色/索引色时 undefined
     * @private
     */
    private childRgb(parent: Element, tag: string): string | undefined {
        const child = this.firstChild(parent, tag);
        if (!child) return undefined;
        const rgb = child.getAttribute('rgb');
        // indexed / theme / auto 颜色不做解析
        if (!rgb || rgb.length < 6) return undefined;
        return this.argbToHex(rgb);
    }

    /**
     * 解析布尔开关属性
     *
     * xlsx 中布尔值可能写作 `val="1"`、`val="true"` 或省略 val（此时默认 true）。
     *
     * @param value - 属性值字符串或 null
     * @returns 是否为开启状态
     * @private
     */
    private isOn(value: string | null | undefined): boolean {
        if (value === null || value === undefined) return false;
        return value === '1' || value === 'true' || value === 'on';
    }

    /**
     * 解析单个工作表 XML（sheetN.xml）
     *
     * @param xml - 工作表 Document
     * @param sharedStrings - 共享字符串表
     * @param styleSheet - 样式表（可选）
     * @param name - 工作表名称
     * @returns 提取的工作表数据
     * @private
     */
    private parseWorksheet(
        xml: Document,
        sharedStrings: string[],
        styleSheet: ExcelStyleSheet | null,
        name: string
    ): ExtractedSheet {
        const cells: ExtractedCell[] = [];
        let minRow = Infinity, maxRow = 0, minCol = Infinity, maxCol = 0;

        // — 解析单元格 —
        const rowNodes = xml.getElementsByTagName('row');
        for (let r = 0; r < rowNodes.length; r++) {
            const rowEl = rowNodes[r];
            const rowIdx = parseInt(rowEl.getAttribute('r') || '0', 10) - 1;  // 转 0-based
            if (rowIdx < minRow) minRow = rowIdx;
            if (rowIdx > maxRow) maxRow = rowIdx;

            const cellNodes = rowEl.getElementsByTagName('c');
            for (let c = 0; c < cellNodes.length; c++) {
                const cEl = cellNodes[c];
                const ref = cEl.getAttribute('r') || '';

                // r 属性缺失时按顺序推断；存在时以 ref 为准
                const pos = this.refToColRow(ref);
                const row = ref ? pos.row : rowIdx;
                const col = ref ? pos.col : c;

                if (col < minCol) minCol = col;
                if (col > maxCol) maxCol = col;

                // — 读取值 —
                const type = cEl.getAttribute('t') || 'n';
                const vEl = this.firstChild(cEl, 'v');
                const isEl = this.firstChild(cEl, 'is');  // inline string
                let value: string | number | boolean | null = null;
                let cellType: ExtractedCell['type'] = 'string';

                if (type === 's' && vEl) {
                    // 共享字符串
                    const idx = parseInt(vEl.textContent || '0', 10);
                    value = sharedStrings[idx] ?? '';
                    cellType = 'string';
                } else if (type === 'inlineStr' && isEl) {
                    // 内联字符串
                    value = this.extractInlineText(isEl);
                    cellType = 'string';
                } else if (type === 'str' && vEl) {
                    // 公式的字符串结果
                    value = vEl.textContent || '';
                    cellType = 'string';
                } else if (type === 'b' && vEl) {
                    value = vEl.textContent === '1';
                    cellType = 'boolean';
                } else if (type === 'e' && vEl) {
                    value = vEl.textContent || '';
                    cellType = 'error';
                } else if (type === 'd' && vEl) {
                    // ISO 8601 日期
                    value = vEl.textContent || '';
                    cellType = 'date';
                } else if (vEl) {
                    const num = parseFloat(vEl.textContent || '0');
                    value = isNaN(num) ? (vEl.textContent || '') : num;
                    cellType = 'number';
                }

                // — 样式：通过 s 属性（样式索引）查 cellXfs —
                let style: CellStyleInfo | undefined;
                let numberFormat: string | undefined;
                const sIdx = cEl.getAttribute('s');
                if (sIdx !== null && styleSheet) {
                    const parsed = this.styleFromXf(styleSheet, parseInt(sIdx, 10));
                    style = parsed.style;
                    numberFormat = parsed.numberFormat;
                }

                cells.push({ row, col, value, type: cellType, style, numberFormat });
            }
        }

        // — 合并区域 —
        const merges: Array<{ s: { r: number; c: number }; e: { r: number; c: number } }> = [];
        const mergeNodes = xml.getElementsByTagName('mergeCell');
        for (let i = 0; i < mergeNodes.length; i++) {
            const ref = mergeNodes[i].getAttribute('ref') || '';
            const parts = ref.split(':');
            if (parts.length === 2) {
                const s = this.refToColRow(parts[0]);
                const e = this.refToColRow(parts[1]);
                merges.push({ s: { r: s.row, c: s.col }, e: { r: e.row, c: e.col } });
            }
        }

        // — 列宽 —
        // xlsx 的 width 以「最大数字宽度字符数」为单位，换算为像素：≈ width * 7 + 5
        const cols: Array<{ index: number; width?: number; customWidth?: boolean }> = [];
        const colsNodes = xml.getElementsByTagName('cols');
        if (colsNodes.length > 0) {
            const colChildren = colsNodes[0].children;
            for (let i = 0; i < colChildren.length; i++) {
                const colEl = colChildren[i];
                if ((colEl.localName || colEl.nodeName) !== 'col') continue;

                const min = parseInt(colEl.getAttribute('min') || '0', 10);
                const max = parseInt(colEl.getAttribute('max') || '0', 10);
                const width = parseFloat(colEl.getAttribute('width') || '0') || undefined;
                const customWidth = colEl.getAttribute('customWidth') === '1';

                for (let c = min; c <= max; c++) {
                    cols.push({
                        index: c - 1,  // 0-based
                        width: width ? Math.round(width * 7 + 5) : undefined,
                        customWidth
                    });
                }
            }
        }

        // — 行高 —
        // xlsx 的 ht 单位为磅，换算为像素：像素 = 磅 × 96 / 72
        const rowsData: Array<{ index: number; height?: number; customHeight?: boolean }> = [];
        for (let i = 0; i < rowNodes.length; i++) {
            const r = parseInt(rowNodes[i].getAttribute('r') || '0', 10) - 1;
            const ht = parseFloat(rowNodes[i].getAttribute('ht') || '0') || undefined;
            const customHeight = rowNodes[i].getAttribute('customHeight') === '1';
            rowsData.push({
                index: r,
                height: ht ? Math.round(ht * 96 / 72) : undefined,
                customHeight
            });
        }

        if (minRow === Infinity) minRow = 0;
        if (minCol === Infinity) minCol = 0;

        return {
            name, cells, merges, cols,
            rows: rowsData, images: [],
            range: { minRow, maxRow, minCol, maxCol }
        };
    }

    /**
     * 依据 cellXfs 索引解析单元格样式
     *
     * @param styleSheet - 样式表
     * @param sIndex - cellXfs 索引
     * @returns 样式与数字格式码
     * @private
     */
    private styleFromXf(
        styleSheet: ExcelStyleSheet,
        sIndex: number
    ): { style?: CellStyleInfo, numberFormat?: string } {
        const xf = styleSheet.cellXfs[sIndex];
        if (!xf) return {};

        // 数字格式：自定义表优先，其次内置表
        const numberFormat = styleSheet.numFmts.get(xf.numFmtId) ||
            this.builtinNumFmt(xf.numFmtId);

        const style: CellStyleInfo = {};

        // — 字体 —
        const font = styleSheet.fonts[xf.fontId];
        if (font && xf.applyFont !== false) {
            style.font = {
                name: font.name,
                size: font.size,
                bold: font.bold || undefined,
                italic: font.italic || undefined,
                color: font.color,
                underline: font.underline || undefined,
                strikethrough: font.strikethrough || undefined
            };
        }

        // — 填充（跳过默认填充：none 与 gray125）—
        const fill = styleSheet.fills[xf.fillId];
        if (fill && fill.patternType && fill.patternType !== 'none' &&
            fill.patternType !== 'gray125' && xf.applyFill !== false) {
            style.fill = {
                patternType: fill.patternType,
                fgColor: fill.fgColor
            };
        }

        // — 对齐 —
        if (xf.alignment && xf.applyAlignment !== false) {
            const horizontal = xf.alignment.horizontal as
                'left' | 'center' | 'right' | 'justify' | undefined;
            style.alignment = {
                horizontal,
                // xlsx 的 "center" → 项目的 "middle"
                vertical: xf.alignment.vertical === 'center'
                    ? 'middle' as const
                    : xf.alignment.vertical as 'top' | 'bottom' | undefined,
                wrapText: xf.alignment.wrapText || undefined,
                indent: xf.alignment.indent
            };
        }

        // — 边框 —
        const border = styleSheet.borders[xf.borderId];
        if (border && xf.applyBorder !== false) {
            style.border = {
                top: border.top?.style ? { style: border.top.style, color: border.top.color || '#000000' } : undefined,
                bottom: border.bottom?.style ? { style: border.bottom.style, color: border.bottom.color || '#000000' } : undefined,
                left: border.left?.style ? { style: border.left.style, color: border.left.color || '#000000' } : undefined,
                right: border.right?.style ? { style: border.right.style, color: border.right.color || '#000000' } : undefined,
            };
        }

        return { style: Object.keys(style).length > 0 ? style : undefined, numberFormat };
    }

    /**
     * 提取工作表中的图片
     *
     * 流程：工作表 rels 中定位 drawing 关系 → 读取 drawingN.xml
     *      → 遍历 pic 元素 → blip 的 r:embed 经 drawing rels 映射到 media 文件
     *
     * @param files - 解压后的文件映射
     * @param sheetPath - 工作表 XML 路径
     * @param drawingRelPaths - workbook rels 中的 drawing 路径映射（备用定位手段）
     * @returns 图片信息数组
     * @private
     */
    private extractSheetImages(
        files: Map<string, Uint8Array>,
        sheetPath: string,
        drawingRelPaths: Map<string, string>
    ): ExtractedImage[] {
        // —— 读取工作表自身的 rels，定位 drawing ——
        const sheetRelsPath = sheetPath
            .replace('worksheets/', 'worksheets/_rels/')
            .replace('.xml', '.xml.rels');
        const sheetRels = this.readXml(files, sheetRelsPath);
        if (!sheetRels) return [];

        let drawingPath = '';
        const rels = sheetRels.getElementsByTagName('Relationship');
        for (let i = 0; i < rels.length; i++) {
            const id = rels[i].getAttribute('Id') || '';
            const target = rels[i].getAttribute('Target') || '';
            if (target.includes('drawing')) {
                const resolved = target.startsWith('/') ? target.slice(1) : `xl/${target}`;
                // 优先使用 workbook rels 中的映射（更可靠）
                drawingPath = drawingRelPaths.get(id) || resolved;
                break;
            }
        }

        if (!drawingPath) return [];
        const drawingXml = this.readXml(files, drawingPath);
        if (!drawingXml) return [];

        // —— 读取 drawing 的 rels：rId → media 路径 ——
        const drawingRelsPath = drawingPath.replace('drawings/', 'drawings/_rels/')
            .replace('.xml', '.xml.rels');
        const drawingRels = this.readXml(files, drawingRelsPath);
        if (!drawingRels) return [];

        const mediaMap = new Map<string, string>();
        const blipRels = drawingRels.getElementsByTagName('Relationship');
        for (let i = 0; i < blipRels.length; i++) {
            const id = blipRels[i].getAttribute('Id') || '';
            const target = blipRels[i].getAttribute('Target') || '';
            if (target.includes('media/')) {
                mediaMap.set(id, target.startsWith('/') ? target.slice(1) : `xl/${target}`);
            }
        }

        if (mediaMap.size === 0) return [];

        // —— 遍历 pic 元素提取图片 ——
        const images: ExtractedImage[] = [];
        const pics = drawingXml.getElementsByTagName('pic');
        for (let i = 0; i < pics.length; i++) {
            const picEl = pics[i];

            // blip 的 r:embed 指向图片数据
            const blips = picEl.getElementsByTagName('blip');
            if (blips.length === 0) continue;
            const embed = blips[0].getAttribute('r:embed') ||
                blips[0].getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed') || '';
            const mediaPath = mediaMap.get(embed);
            if (!mediaPath) continue;

            const imgData = files.get(mediaPath);
            if (!imgData) continue;

            // 向上查找所属锚点元素以读取坐标
            let anchorEl: Element | null = picEl;
            const rootEl = drawingXml as unknown as Element;
            while (anchorEl && anchorEl !== rootEl) {
                const localName = anchorEl.localName || anchorEl.nodeName;
                if (localName === 'twoCellAnchor' || localName === 'oneCellAnchor') break;
                anchorEl = anchorEl.parentNode as Element | null;
            }

            // 解析锚点坐标（EMU 单位）
            let col = 0, row = 0, colOff = 0, rowOff = 0;
            if (anchorEl) {
                const fromEl = this.firstChild(anchorEl, 'from');
                if (fromEl) {
                    const readEmu = (tag: string): number => {
                        const nodes = fromEl.getElementsByTagName(tag);
                        if (nodes.length === 0) return 0;
                        return parseInt(nodes[0].textContent || '0', 10) || 0;
                    };
                    col = readEmu('col');
                    row = readEmu('row');
                    colOff = readEmu('colOff');
                    rowOff = readEmu('rowOff');
                }
            }

            // 从文件名推断 MIME
            const ext = mediaPath.split('.').pop()?.toLowerCase() || '';
            const mimeMap: Record<string, string> = {
                png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg',
                gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp',
                tif: 'image/tiff', tiff: 'image/tiff', emf: 'image/emf', wmf: 'image/wmf'
            };

            // 图片名称
            const nvPr = picEl.getElementsByTagName('cNvPr');
            const name = nvPr.length > 0 ? (nvPr[0].getAttribute('name') || undefined) : undefined;

            images.push({
                data: imgData,
                mimeType: mimeMap[ext] || 'image/png',
                anchor: { col, row, colOffset: colOff, rowOffset: rowOff },
                name
            });
        }

        return images;
    }

    // ════════════════════════════════════════════════════════════
    //  工具方法
    // ════════════════════════════════════════════════════════════

    /**
     * ARGB 颜色字符串转 #RRGGBB
     *
     * @param argb - 如 "FF0000FF" 或 "FF0000"
     * @returns 如 "#0000FF"
     * @private
     */
    private argbToHex(argb: string): string {
        const s = argb.replace('#', '');
        if (s.length === 8) return '#' + s.slice(2);  // AARRGGBB → RRGGBB
        if (s.length === 6) return '#' + s;            // RRGGBB
        return '#000000';
    }

    /**
     * Excel 边框样式 → 项目边框宽度
     *
     * @param style - Excel 边框样式名（thin/medium/thick 等）
     * @returns 边框宽度数字；无边框时 undefined
     * @private
     */
    private borderStyleToWidth(style?: string): number | undefined {
        if (!style || style === 'none') return undefined;
        const thinStyles = ['thin', 'dotted', 'dashed', 'dashDot', 'dashDotDot', 'hair', 'slantDashDot'];
        const mediumStyles = ['medium', 'mediumDashDot', 'mediumDashDotDot', 'mediumDashed'];
        if (thinStyles.includes(style)) return 1;
        if (mediumStyles.includes(style)) return 2;
        if (style === 'thick') return 3;
        return 1;  // 默认细线
    }

    /**
     * 列号（1-based）→ 列名（A, B, ..., Z, AA, AB...）
     *
     * @param col - 1-based 列号
     * @returns 列名
     * @private
     */
    private colToName(col: number): string {
        let result = '';
        let n = col;
        while (n > 0) {
            const rem = (n - 1) % 26;
            result = String.fromCharCode(65 + rem) + result;
            n = Math.floor((n - 1) / 26);
        }
        return result;
    }

    /**
     * 单元格引用（如 "A1"）→ { col, row }（0-based）
     *
     * @param ref - 单元格引用
     * @returns { col: number; row: number }
     * @private
     */
    private refToColRow(ref: string): { col: number; row: number } {
        const match = ref.match(/^([A-Z]+)(\d+)$/);
        if (!match) return { col: 0, row: 0 };
        const colName = match[1];
        const row = parseInt(match[2], 10) - 1;  // 0-based
        let col = 0;
        for (let i = 0; i < colName.length; i++) {
            col = col * 26 + (colName.charCodeAt(i) - 64);
        }
        return { col: col - 1, row };  // 0-based
    }

    /**
     * 内置数字格式 ID → 格式码
     *
     * 见 ECMA-376 附录，此处收录常用项。
     *
     * @param id - 内置格式 ID
     * @returns 格式码或 undefined
     * @private
     */
    private builtinNumFmt(id: number): string | undefined {
        const builtins: Record<number, string> = {
            0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
            9: '0%', 10: '0.00%', 11: '0.00E+00', 14: 'm/d/yyyy',
            15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
            18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss',
            22: 'm/d/yyyy h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)',
            39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
            45: 'mm:ss', 46: '[h]:mm:ss',
            47: 'mmss.0', 48: '##0.0', 49: '@'
        };
        return builtins[id];
    }
}
