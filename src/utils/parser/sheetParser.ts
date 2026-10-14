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
 *    │ xmlScanner 流式扫描│     │  biff.ts          │  BIFF8 记录流
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
import {
    ArraySheetSource,
    PackedSheetSource,
    SheetSource,
    PackedSheet,
    ExtractedSheetLike,
} from './cellSource';
import { ZipParser, ZipArchive } from './zip';
import { CompoundFile } from './cfb';
import { BiffParser, BiffXf, BiffFont, BiffWorkbook } from './biff';
import { scanXml, attrValue, XmlAttrs } from './xmlScanner';
import { yieldToUI, SliceBudget } from '../../utils/scheduler';

// ── 中间数据结构 ──

/** 第一步输出：解压/解码后的工作簿 */
export interface DecodedWorkbook {
    /** 文件格式 */
    format: 'xlsx' | 'xls';
    /** xlsx 解压后的文件映射（路径 → 内容），xls 为 null */
    files: Map<string, Uint8Array> | null;
    /** xls：OLE2 复合文档中的 Workbook 流；xlsx 为 null */
    rawData: Uint8Array | null;
    /**
     * 本次解压是否**跳过了媒体条目**（`xl/media/**`）
     *
     * 为 true 时，`extract(decoded, { readImages: true })` 也拿不到任何图片 ——
     * media 字节根本没被 inflate 进来。这是刻意的：见 {@link SheetParser.decode}。
     */
    mediaSkipped?: boolean;
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
    /**
     * 样式内部编号（xlsx 即 cellXfs 索引，xls 即 XF 索引）
     *
     * **同一编号的 `style` 引用必定是同一个对象**：提取阶段按编号做了内联缓存
     * （见 `SheetParser.parseWorksheetBytes`）。下游（列式打包）因此可以拿这个
     * 数字直接做样式去重，不必对每格执行 `JSON.stringify(style)`。
     *
     * 未走内联缓存的来源（测试手工造的 `ExtractedSheet`、旧调用方）此字段缺失，
     * 打包侧会退回按内容比较。
     */
    styleId?: number;
    /** 数字格式码 */
    numberFormat?: string;
}

/** 图片锚点（0-based 行列 + EMU 单位偏移） */
export interface ImageAnchor {
    /** 列号（0-based） */
    col: number;
    /** 行号（0-based） */
    row: number;
    /** 列内偏移（EMU） */
    colOffset: number;
    /** 行内偏移（EMU） */
    rowOffset: number;
}

/**
 * 提取的图片信息
 *
 * ## 三种锚点写法与渲染所需信息（性能/正确性约定）
 *
 * xlsx 的 drawing XML 用三种锚点描述一张图，**它们携带的信息不同**：
 * - `twoCellAnchor`：`<from>` + `<to>` 两个锚点，尺寸由「两处的列宽/行高之和」决定
 *   （源文件**不写** `<ext>`），因此必须把 `<to>` 一并提取出来，渲染侧才能算对宽高；
 * - `oneCellAnchor`：只有 `<from>`，尺寸来自 `<ext cx cy>`（EMU），必须提取 `<ext>`；
 * - `absoluteAnchor`：坐标以 EMU 从工作表原点算起、不落在单元格上，**不提取**
 *   （渲染侧无法用列宽/行高表达，硬塞到 A1 只会画错位置）。
 *
 * 历史实现只提取了 `from`（左上角），于是得到「位置对、尺寸为 0」的图片，
 * 渲染不出任何东西 —— 这就是「图能提取但没有绘制通道」的一半原因。
 */
export interface ExtractedImage {
    /** 图片二进制数据 */
    data: Uint8Array;
    /** MIME 类型 */
    mimeType: string;
    /** 左上角锚点（`twoCellAnchor`/`oneCellAnchor` 的 `<from>`） */
    anchor: ImageAnchor;
    /**
     * 图片尺寸（EMU）
     *
     * 仅 `oneCellAnchor` 有（来自 `<ext cx cy>`）；`twoCellAnchor` 由
     * {@link to} 与列宽/行高换算，故这里缺省。`.xls`（Escher）暂无可靠锚点信息。
     */
    ext?: { cx: number; cy: number };
    /**
     * 右下角锚点（`twoCellAnchor` 的 `<to>`）
     *
     * 渲染侧据此用「列宽/行高 + 锚点内 EMU 偏移」算出图片矩形，
     * 比 `ext` 更准（不受 Excel 写入的近似尺寸影响）。
     */
    to?: ImageAnchor;
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
    /** 图片列表（需 `readImages: true`；默认不读，故通常为空数组） */
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
    /**
     * 是否读取图片（**默认 false**）
     *
     * 默认关闭的理由是「不需要就别付代价」：媒体是不可压缩的大块二进制，
     * 实测 37.5 MB 媒体要多花 657 ms、常驻 +150 MB。渲染管道已就绪
     * （`DataCollection.images` + `sheets/sheetImageLayer.ts`），
     * 打开文件的路径会显式传 `readImages: true`；
     * 只取单元格值时（批处理、测试、预览）保持默认即可。
     *
     * 注意：**必须在 `decode` 阶段就声明**（`decode(file, { readImages: true })`）——
     * media 条目在解压阶段就被跳过了，后续 `extract` 无从补救
     * （`DecodedWorkbook.mediaSkipped` 会如实标记这一点）。
     */
    readImages?: boolean;
}

/**
 * 边框线型宽度表（xls 的 XF 记录中，边框以 3 位数值表示线型）
 *
 * 值 0 表示无边框，1-7 依次为不同线型。
 */
/**
 * xls 边框线型索引 → xlsx `borderStyle` 名称
 *
 * BIFF8 的边框线型是 4 位（0–15），共 14 种有效取值（MS-XLS 2.4.268 / BORDER 枚举）。
 * 历史实现只列了 8 项（0–7），索引 8–13（BORDER_HAIR..BORDER_DASHDOTDOT）会落空
 * 而导致这些边的边框被整条丢弃。索引 14/15 未定义，刻意留空。
 */
const BIFF_BORDER_STYLES = [
    'none',          // 0  BORDER_NONE
    'thin',          // 1  BORDER_THIN
    'medium',        // 2  BORDER_MEDIUM
    'dashed',        // 3  BORDER_DASHED
    'dotted',        // 4  BORDER_DOTTED
    'thick',         // 5  BORDER_THICK
    'double',        // 6  BORDER_DOUBLE
    'hair',          // 7  BORDER_HAIR
    'mediumDashed',  // 8  BORDER_MEDIUMDASHED
    'dashDot',       // 9  BORDER_DASHDOT
    'mediumDashDot', // 10 BORDER_MEDIUMDASHDOT
    'dashDotDot',    // 11 BORDER_DASHDOTDOT
    'mediumDashDotDot', // 12 BORDER_MEDIUMDASHDOTDOT
    'slantDashDot'   // 13 BORDER_SLANTDASHDOT
];

/** xls 填充图案索引 → xlsx patternType 名称 */
const BIFF_FILL_PATTERNS = [
    'none', 'solid', 'mediumGray', 'darkGray', 'lightGray', 'darkHorizontal',
    'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis',
    'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis',
    'gray125', 'gray0625'
];

export class SheetParser {

    /**
     * 合并映射数值键的列跨度：`key = row * CELL_KEY_STRIDE + col`
     *
     * 取 Excel 单表的最大列数（16384）。列号恒小于该值，因此 (row,col) →
     * 数值键一一对应，可安全替代 `${row},${col}` 字符串键。
     */
    private static readonly CELL_KEY_STRIDE = 16384;

    /**
     * 列名缓存：下标 = 1-based 列号，值为 "A"/"B"/..."AA"
     *
     * `colToName` 在生成单元格名时被逐格调用，而列号的取值空间很小
     * （≤ 16384）。缓存后每列只计算一次，百万格规模下省掉百万次字符串构造。
     * 下标 0 占位空串，使下标与列号直接对齐。
     */
    private readonly _colNameCache: string[] = [''];

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
     * ## 媒体条目默认不解压（性能约定，请勿回退）
     *
     * `xl/media/**` 是嵌入的图片二进制（照片/截图/扫描件）。它们**不可压缩**，
     * 逐字节 inflate 的代价与体积成正比，且解压后整份常驻内存。实测：
     *
     * | 媒体总量 | inflate 耗时 | 内存增量 |
     * |---|---|---|
     * | 37.5 MB | 657 ms | 150 MB |
     * | 94 MB | 838 ms | 199 MB |
     *
     * 而项目当前**没有任何图片渲染基建**（`SheetParserResult.images` 无人消费，
     * `fileOpener` 直接丢弃），所以这份代价 100% 是浪费 —— 既拖慢打开，
     * 又把几十上百 MB 的字节留在内存里造成 GC 抖动。
     *
     * 因此：只有调用方显式传 `readImages: true` 时才读取 media 条目
     * （见 {@link SheetParserOptions.readImages}）。
     *
     * @param {File} file - 用户选择的 .xlsx 或 .xls 文件
     * @param {SheetParserOptions} [options] - 解析选项（仅 `readImages` 在此步生效）
     * @returns {Promise<DecodedWorkbook>} 解压/解码后的工作簿
     * @throws {Error} 文件格式不支持或内容损坏
     * @example
     * const decoded = await parser.decode(file);
     */
    public async decode(file: File, options: SheetParserOptions = {}): Promise<DecodedWorkbook> {
        const extension = file.name.toLowerCase().split('.').pop() || '';
        const buffer = new Uint8Array(await file.arrayBuffer());

        if (extension === 'xlsx') {
            // —— xlsx：原生 ZIP 解包 ——
            const zip = new ZipParser(buffer);
            const archive: ZipArchive = zip.parseCentralDirectory();

            const readImages = options.readImages === true;
            const files = new Map<string, Uint8Array>();
            let mediaSkipped = false;
            for (const path of archive.entries.keys()) {
                // 目录条目（以 / 结尾）无需读取
                if (path.endsWith('/')) continue;
                // 媒体条目：除非明确要图，否则连 inflate 都不做（见方法头说明）
                if (!readImages && SheetParser.isMediaEntry(path)) {
                    mediaSkipped = true;
                    continue;
                }
                const bytes = await zip.readEntry(archive, path);
                if (bytes) files.set(path, bytes);
            }
            return { format: 'xlsx', files, rawData: null, mediaSkipped };
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

    /**
     * 取 ZIP 路径的所在目录（不含尾斜杠）
     *
     * `xl/worksheets/sheet1.xml` → `xl/worksheets`；无斜杠时返回空串。
     * @param path - ZIP 内路径
     * @returns 目录路径
     * @private
     */
    private static dirOf(path: string): string {
        const i = path.lastIndexOf('/');
        return i < 0 ? '' : path.slice(0, i);
    }

    /**
     * 把 OOXML 关系文件里的 `Target` 解析为 ZIP 内路径
     *
     * OPC 规范中 `Target` 相对**关系文件所属部件的目录**，Excel 一贯写
     * `../media/image1.png`。历史实现写死了 `xl/${target}`，于是
     * `xl/../media/image1.png` 查不到任何条目 —— 图片提取**永远返回 0 张**
     * （因为没人消费 `images`，这个 bug 一直没被发现）。
     *
     * 同时处理三种写法：
     * - 包内绝对路径 `/xl/media/a.png`；
     * - 相对路径 `../media/a.png` / `./a.xml`；
     * - 外部链接（`http:` / `file:` 等）→ 返回空串（不适用，调用方跳过）。
     * @param baseDir - 关系文件所属部件的目录（如 `xl/drawings`）
     * @param target - 关系文件里的 Target 原文
     * @returns ZIP 内路径；不适用时为空串
     * @private
     */
    private static resolveZipTarget(baseDir: string, target: string): string {
        const t = target.replace(/\\/g, '/');
        if (t.startsWith('/')) return t.slice(1);
        // 带协议的外部目标（http:、file: 等）不是包内部件
        if (/^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(t)) return '';
        const parts = baseDir ? baseDir.split('/').filter(Boolean) : [];
        for (const seg of t.split('/')) {
            if (seg === '' || seg === '.') continue;
            if (seg === '..') parts.pop();
            else parts.push(seg);
        }
        return parts.join('/');
    }

    /**
     * 判断 ZIP 条目是否为「图片/媒体」二进制（可安全跳过解压）
     *
     * 覆盖 Excel 的媒体目录约定：`xl/media/**`（也有少数工具写 `media/**`，
     * 或把嵌入式 OLE 对象放在 `xl/embeddings/**`）。只按扩展名白名单放行
     * 会漏掉自定义文件名，因此按「目录 + 常见媒体扩展名」双条件判断。
     * @param path - ZIP 内路径
     * @returns 是否属于媒体条目
     * @private
     */
    private static isMediaEntry(path: string): boolean {
        if (!path.includes('media/')) return false;
        const ext = path.slice(path.lastIndexOf('.') + 1).toLowerCase();
        switch (ext) {
            case 'png': case 'jpg': case 'jpeg': case 'gif': case 'bmp': case 'webp':
            case 'tif': case 'tiff': case 'emf': case 'wmf': case 'svg':
            case 'mp4': case 'mov': case 'avi': case 'wmv': case 'm4a': case 'mp3': case 'wav':
                return true;
            default:
                return false;
        }
    }

    // ════════════════════════════════════════════════════════════
    //  第二步：读取单元格数据、样式、图片
    // ════════════════════════════════════════════════════════════

    /**
     * 第二步：从解压产物中读取单元格数据、样式、图片
     *
     * - **xlsx**：用自研流式扫描器 {@link scanXml} 逐字节解析 XML，完整提取
     *   字体/填充/对齐/边框/数字格式/合并/图片（不构建 DOM）
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
        // readImages 默认 false：见 SheetParserOptions.readImages（读图是纯浪费）
        const { readStyles = true, readImages = false } = options;

        if (decoded.format === 'xlsx' && decoded.files) {
            return this.extractXlsx(decoded.files, readStyles, readImages);
        }

        if (decoded.format === 'xls' && decoded.rawData) {
            return this.extractXls(decoded.rawData, readStyles, readImages);
        }

        throw new Error('解压产物无效');
    }

    /**
     * 解析 xlsx 的工作表，完整提取单元格数据、样式、图片
     *
     * **全部走流式扫描器（{@link scanXml}），不构建 DOM**（性能约定，请勿回退）：
     * `DOMParser` 会为整份 sheet XML 建立驻留 DOM，是大文件内存峰值的首要来源，
     * 且该 API 在 Web Worker 中不存在；扫描器单趟逐字节推进、内存 O(1)，
     * 使这条路径可以直接运行在主线程或 Worker 中。
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
        // —— workbook.xml：工作表名称与 r:id ——
        const sheetInfos: Array<{ name: string; rId: string }> = [];
        const workbookBytes = files.get('xl/workbook.xml');
        if (workbookBytes) {
            scanXml(workbookBytes, {
                onStart: (name, attrs) => {
                    if (name !== 'sheet') return;
                    sheetInfos.push({
                        name: attrs['name'] || '',
                        rId: attrValue(attrs, 'r:id') || attrValue(attrs, 'id') || ''
                    });
                }
            });
        }

        // —— workbook.xml.rels：r:id → sheetN.xml 路径 ——
        // 该 rels 文件位于 `xl/_rels/`，故其 Target 相对 `xl/` 解析
        const sheetPaths = new Map<string, string>();
        const drawingRelPaths = new Map<string, string>();
        const relsBytes = files.get('xl/_rels/workbook.xml.rels');
        if (relsBytes) {
            scanXml(relsBytes, {
                onStart: (name, attrs) => {
                    if (name !== 'Relationship') return;
                    const id = attrs['Id'] || '';
                    const target = attrs['Target'] || '';
                    const resolved = SheetParser.resolveZipTarget('xl', target);
                    if (!resolved) return;
                    if (target.includes('worksheets/')) {
                        sheetPaths.set(id, resolved);
                    } else if (target.includes('drawing')) {
                        drawingRelPaths.set(id, resolved);
                    }
                }
            });
        }

        // —— styles.xml：样式表 ——
        let styleSheet: ExcelStyleSheet | null = null;
        if (readStyles) {
            const stylesBytes = files.get('xl/styles.xml');
            if (stylesBytes) {
                styleSheet = this.parseStyleSheetBytes(stylesBytes);
            }
        }

        // —— sharedStrings.xml：共享字符串表 ——
        // <si> 内所有后代 <t> 的文本按文档顺序拼接（与 DOM 版 extractInlineText 语义一致）
        const sharedStrings: string[] = [];
        const ssBytes = files.get('xl/sharedStrings.xml');
        if (ssBytes) {
            let inSi = false;
            let inT = false;
            let buffer = '';
            scanXml(ssBytes, {
                onStart: (name) => {
                    if (name === 'si') { inSi = true; buffer = ''; }
                    else if (inSi && name === 't') inT = true;
                },
                onEnd: (name) => {
                    if (name === 'si') { sharedStrings.push(buffer); inSi = false; }
                    else if (name === 't') inT = false;
                },
                onText: (text) => { if (inSi && inT) buffer += text; }
            });
        }

        // —— 遍历每个工作表 ——
        const result: ExtractedSheet[] = [];
        for (const info of sheetInfos) {
            const sheetPath = sheetPaths.get(info.rId) || 'xl/worksheets/sheet1.xml';
            const sheetBytes = files.get(sheetPath);
            if (!sheetBytes) continue;

            const sheet = this.parseWorksheetBytes(
                sheetBytes, sharedStrings, styleSheet, info.name
            );

            // —— 提取图片：由工作表 rels 中的 drawing 关系精确定位 ——
            if (readImages) {
                sheet.images = this.extractSheetImagesBytes(files, sheetPath, drawingRelPaths);
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
        // readImages 必须传给 BiffParser：不读图时连 Escher 绘制记录都不解析
        // （否则每张图的字节都会被复制一份，见 BiffParser.readImages）
        const biff = new BiffParser(rawData, { readImages });
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

            // 样式内联缓存：XF 索引 → 已解析样式（见 parseWorksheetBytes 里的同款说明）
            // xls 的样式是「同一 XF 索引必然同内容」，逐格重解析纯属浪费。
            const styleCache = new Map<number, {
                style: CellStyleInfo | undefined;
                numberFormat: string | undefined;
            }>();

            for (const cell of raw.cells) {
                let style: CellStyleInfo | undefined;
                let styleId: number | undefined;
                let numberFormat: string | undefined;
                if (readStyles) {
                    const idx = cell.styleIndex;
                    let cached = styleCache.get(idx);
                    if (cached === undefined) {
                        cached = {
                            style: this.biffStyleToCellStyle(wb, idx),
                            numberFormat: this.biffNumberFormat(wb, idx),
                        };
                        styleCache.set(idx, cached);
                    }
                    style = cached.style;
                    numberFormat = cached.numberFormat;
                    styleId = idx;
                }

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
                    styleId,
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
            // ⚠️ solid（实心，pattern 1）时底色取自**前景色**（fillColorIndex）；
            // 其余图案类填充则由前景色画网点/条纹、背景色作衬底，故取背景色更接近观感。
            // 历史实现只写 `fg || bg`，在 `fg` 为 0（自动/未设）时会错误地用背景色当底色。
            const fg = this.paletteColor(wb.palette, xf.fillColorIndex);
            const bg = this.paletteColor(wb.palette, xf.fillBackgroundIndex);
            style.fill = {
                patternType: fillPatternName,
                fgColor: xf.fillPattern === 1 ? (fg || bg) : (bg || fg)
            };
        }

        // —— 对齐 ——
        // BIFF8 水平对齐：0 常规 1 左 2 居中 3 右 4 填充 5 两端 6 跨列居中 7 分散
        //   （⚠️ 2 = 居中，不是左；历史把 2 当 left 且 7 当 center，均错）
        // BIFF8 垂直对齐：0 顶 1 居中 2 底 3 两端 4 分散
        //   （⚠️ 历史只映射了 0/1，且把 0 当 bottom —— 实际 0 是 top）
        //
        // 项目的 `TextAlign`/`VerticalAlign` 只有 4/3 个取值（无 distributed），
        // 故 7/4 这类 BIFF 取值就近归一到 'justify'，避免写入渲染端不认识的字符串。
        const hAlignMap: Record<number, NonNullable<CellStyleInfo['alignment']>['horizontal']> = {
            0: 'left', 1: 'left', 2: 'center', 3: 'right',
            4: 'justify', 5: 'justify', 6: 'center', 7: 'justify'
        };
        const vAlignMap: Record<number, NonNullable<CellStyleInfo['alignment']>['vertical']> = {
            0: 'top', 1: 'middle', 2: 'bottom', 3: 'bottom', 4: 'middle'
        };
        const horizontal = hAlignMap[xf.horizontalAlignment];
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
        // 四边**各自取色**：BIFF8 把左/右色编码在 u32@10 的高位，
        // 上/下色编码在 u32@14 的低位（见 biff.ts 的 parseXf），实现层面已拆为
        // 4 个独立字段，不能再共用一个颜色。
        const border: NonNullable<CellStyleInfo['border']> = {};
        const makeEdge = (lineStyle: number, colorIndex: number) => {
            const name = BIFF_BORDER_STYLES[lineStyle];
            if (!name || name === 'none') return undefined;
            return { style: name, color: this.paletteColor(wb.palette, colorIndex) || '#000000' };
        };
        const bTop = makeEdge(xf.borderTop, xf.borderColorTop);
        const bBottom = makeEdge(xf.borderBottom, xf.borderColorBottom);
        const bLeft = makeEdge(xf.borderLeft, xf.borderColorLeft);
        const bRight = makeEdge(xf.borderRight, xf.borderColorRight);
        if (bTop) border.top = bTop;
        if (bBottom) border.bottom = bBottom;
        if (bLeft) border.left = bLeft;
        if (bRight) border.right = bRight;
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
        // 索引 0 = 「自动/未设」，1–7 为 BIFF8 保留值（不得用于单元格样式），
        // 三者都当作「无颜色」返回 undefined。
        if (index <= 7) return undefined;
        const color = palette[index];
        return color || undefined;
    }

    // ════════════════════════════════════════════════════════════
    //  第三步：生成与项目数据格式一致的数据
    // ════════════════════════════════════════════════════════════

    /**
     * 第三步：将提取的数据映射为项目 Cell[]/ColHeader[]/RowHeader[] 格式（数组式输入）
     *
     * 输入是 `extract()` 原样产出的 `ExtractedCell[]`，用于**主线程回退路径**
     * （Worker 不可用）与 `xls` 直连路径。
     *
     * Worker 可用时走的是 {@link generatePacked}：数据已被 {@link packSheets}
     * 打成列式结构，跨线程克隆成本低一个数量级（详见 `cellSource.ts` 的说明）。
     * 两条路径共用 {@link generateFrom}，映射规则只有一份。
     * @param {ExtractedSheet[]} sheets - 第二步提取的工作表
     * @param {SheetParserOptions} [options] - 选项
     * @returns {Promise<SheetParserResult>} 与项目格式一致的数据
     * @example
     * const result = await parser.generate(sheets, { sheetIndex: 0 });
     */
    public async generate(
        sheets: ExtractedSheet[],
        options: SheetParserOptions = {}
    ): Promise<SheetParserResult> {
        const sheet = sheets[options.sheetIndex ?? 0] || sheets[0];
        return this.generateFrom(sheet ? new ArraySheetSource(sheet) : null);
    }

    /**
     * 第三步：将**列式打包**的数据映射为项目 Cell[]/ColHeader[]/RowHeader[] 格式
     *
     * 与 {@link generate} 的唯一区别是输入形态：这里是 Worker 回传的
     * {@link PackedSheet}（typed array + 样式表 + 文本池）。
     * 之所以要有这条路径：把 6 万格「普通对象 + 嵌套样式对象」原样结构化克隆回
     * 主线程要 489 ms，而列式打包（typed array 走 transfer）只要 48 ms。
     * @param {PackedSheet[]} packed - Worker 打包回传的工作表
     * @param {SheetParserOptions} [options] - 选项
     * @returns {Promise<SheetParserResult>} 与项目格式一致的数据
     * @example
     * const result = await parser.generatePacked(packed, { sheetIndex: 0 });
     */
    public async generatePacked(
        packed: PackedSheet[],
        options: SheetParserOptions = {}
    ): Promise<SheetParserResult> {
        const sheet = packed[options.sheetIndex ?? 0] || packed[0];
        return this.generateFrom(sheet ? new PackedSheetSource(sheet) : null);
    }

    /**
     * 第三步的实现：把数据源映射为项目格式（**两种输入共用的唯一映射路径**）
     *
     * 映射规则：
     * - Excel 0-based 行列 → 项目 1-based（col=0→1=A, row=0→1）
     * - Excel 单元格样式 → Cell 属性（backgroundColor/borderWidth/textAlign/alignItems/numberFormat）
     * - Excel 字体样式 → Char 属性（fontFamily/fontSize/fontWeight/fontStyle/fontColor/underline/strikethrough）
     * - Excel 合并区域 → Cell.colspan/rowspan（左上角为锚点）
     * - Excel 列宽 → ColHeader.width（像素）
     * - Excel 行高 → RowHeader.height（像素）
     *
     * ## 为什么是 async（性能约定，请勿改回同步）
     * 本步在主线程执行，是打开大文件时**最后一段长任务**：几十万格要逐个构造
     * `Cell`/`Char` 实例。若一口气跑完，期间页面完全无法响应（表现为「打开即卡死」）。
     * 因此改为**分片执行**：每片最多占用主线程 {@link SliceBudget} 的预算（默认 12ms），
     * 片间 `await yieldToUI()` 让出，浏览器得以处理输入与渲染。
     * 总耗时基本不变（让出开销很小），但期间窗口可响应、可关闭、可取消。
     * @param {SheetSource | null} source - 数据源（数组式或列式打包式）；null 表示空
     * @returns {Promise<SheetParserResult>} 与项目格式一致的数据
     * @private
     */
    private async generateFrom(source: SheetSource | null): Promise<SheetParserResult> {
        if (!source) {
            return { values: [], colHeaders: [], rowHeaders: [], images: [] };
        }

        // —— 构建合并映射：左上角单元格 → { rowspan, colspan } ——
        // 用数值键 `row * COL_STRIDE + col` 取代模板字符串键：
        // 逐格拼接 `${row},${col}` 会在百万格规模下产生百万次字符串分配，是纯粹的浪费。
        const mergeMap = new Map<number, { rowspan: number; colspan: number }>();
        for (const merge of source.merges) {
            const rowspan = merge.e.r - merge.s.r + 1;
            const colspan = merge.e.c - merge.s.c + 1;
            if (rowspan > 1 || colspan > 1) {
                mergeMap.set(merge.s.r * SheetParser.CELL_KEY_STRIDE + merge.s.c, { rowspan, colspan });
            }
        }

        // —— 转换单元格 ——
        // 字体模板缓存：同一份「字体签名」只构造一次模板 Char，
        // 之后对每个字符做稀疏克隆（只复制确实存在的样式字段）。
        // 原实现对**每个字符**都执行 7 次条件赋值，百万格 × 平均十字即千万级写入。
        const fontTemplateCache = new Map<string, Char>();
        const count = source.length;
        const values: Cell[] = new Array(count);
        // 分片让出：超过预算就让出主线程一次，
        // 使「打开大文件」期间页面仍可响应（滚动/关闭/取消）。
        const slice = new SliceBudget();
        for (let i = 0; i < count; i++) {
            values[i] = this.buildCell(source, i, mergeMap, fontTemplateCache);
            if (slice.shouldYield(i)) await slice.yieldNow();
        }

        // —— 列头：生成从 1 到 maxCol 的 ColHeader[] ——
        const maxCol = source.range.maxCol + 1;  // 1-based
        const colWidthMap = new Map<number, number>();
        for (const col of source.cols) {
            if (col.width) colWidthMap.set(col.index + 1, col.width);
        }
        const colHeaders: ColHeader[] = [];
        let colLeft = 0;
        for (let c = 1; c <= maxCol; c++) {
            const width = colWidthMap.get(c) ?? DEFAULT_CELL_WIDTH;
            colHeaders.push({ left: colLeft, width });
            colLeft += width;
            if (slice.shouldYield(c)) await slice.yieldNow();
        }

        // —— 行头：生成从 1 到 maxRow 的 RowHeader[] ——
        const maxRow = source.range.maxRow + 1;
        const rowHeightMap = new Map<number, number>();
        for (const row of source.rows) {
            if (row.height) rowHeightMap.set(row.index + 1, row.height);
        }
        const rowHeaders: RowHeader[] = [];
        let rowTop = 0;
        for (let r = 1; r <= maxRow; r++) {
            const height = rowHeightMap.get(r) ?? DEFAULT_CELL_HEIGHT;
            rowHeaders.push({ top: rowTop, height });
            rowTop += height;
            if (slice.shouldYield(r)) await slice.yieldNow();
        }

        return { values, colHeaders, rowHeaders, images: source.images };
    }

    /**
     * 把数据源中的第 i 个单元格映射为项目格式的 `Cell`
     *
     * 从 {@link generateFrom} 的主循环中抽出，使主循环只剩「调用 + 让出」两件事，
     * 便于分片执行时保持可读性。**只读数据源、不持有其引用**：分片让出后
     * 数据源仍可能被调用方替换。
     * @param {SheetSource} source - 数据源
     * @param {number} index - 单元格下标
     * @param {Map<number, { rowspan: number; colspan: number }>} mergeMap - 合并映射（数值键）
     * @param {Map<string, Char>} fontTemplateCache - 字体模板缓存（同一字体只构造一次）
     * @returns {Cell} 项目格式的单元格
     * @private
     */
    private buildCell(
        source: SheetSource,
        index: number,
        mergeMap: Map<number, { rowspan: number; colspan: number }>,
        fontTemplateCache: Map<string, Char>
    ): Cell {
        const col0 = source.colAt(index);
        const row0 = source.rowAt(index);
        const cell = new Cell();
        cell.cell = this.colToName(col0 + 1) + (row0 + 1);

        // 合并信息
        const merge = mergeMap.get(row0 * SheetParser.CELL_KEY_STRIDE + col0);
        if (merge) {
            if (merge.colspan > 1) cell.colspan = merge.colspan;
            if (merge.rowspan > 1) cell.rowspan = merge.rowspan;
        }

        // 值 → Char[]（每个字符一个 Char，共享同一字体样式）
        // 布尔值/空值的显示文本由数据源统一处理（见 cellSource.cellDisplayText）
        const text = source.textAt(index);
        const style = source.styleAt(index);
        if (text) {
            // 字体样式映射：统一到一份「模板」上，避免逐字符重复判定
            const f = style?.font;
            let template: Char | null = null;
            if (f) {
                const key = `${f.name ?? ''}\u0000${f.size ?? ''}\u0000${f.bold ? 1 : 0}${f.italic ? 1 : 0}\u0000${f.color ?? ''}\u0000${f.underline ? 1 : 0}${f.strikethrough ? 1 : 0}`;
                template = fontTemplateCache.get(key) ?? null;
                if (template === null) {
                    template = new Char();
                    if (f.name) template.fontFamily = f.name;
                    if (f.size) template.fontSize = f.size;
                    if (f.bold) template.fontWeight = true;
                    if (f.italic) template.fontStyle = true;
                    if (f.color) template.fontColor = f.color;
                    if (f.underline) template.underline = true;
                    if (f.strikethrough) template.strikethrough = true;
                    fontTemplateCache.set(key, template);
                }
            }

            const chars: Char[] = [];
            // 直接迭代码点（与 Array.from 语义一致，但省掉每格一个中间数组）
            for (const ch of text) {
                if (template === null) {
                    chars.push(Char.of(ch));
                } else {
                    const char = Char.fromRaw(template);
                    char.char = ch;
                    chars.push(char);
                }
            }
            cell.chars = chars;
        }

        // 单元格级样式映射
        if (style) {
            // 背景色
            if (style.fill?.fgColor) {
                cell.backgroundColor = style.fill.fgColor;
            }
            // 对齐
            if (style.alignment?.horizontal) {
                cell.textAlign = style.alignment.horizontal;
            }
            if (style.alignment?.vertical) {
                cell.alignItems = style.alignment.vertical;
            }
            if (style.alignment?.wrapText) {
                cell.wrap = true;
            }
            // 边框：Excel border style → 项目 border width
            if (style.border) {
                if (style.border.top) cell.borderTopWidth = this.borderStyleToWidth(style.border.top.style);
                if (style.border.bottom) cell.borderBottomWidth = this.borderStyleToWidth(style.border.bottom.style);
                if (style.border.left) cell.borderLeftWidth = this.borderStyleToWidth(style.border.left.style);
                if (style.border.right) cell.borderRightWidth = this.borderStyleToWidth(style.border.right.style);
                // 边框颜色（取第一个有颜色的边框）
                const borderColor = style.border.top?.color || style.border.bottom?.color
                    || style.border.left?.color || style.border.right?.color;
                if (borderColor) cell.borderColor = borderColor;
            }
        }

        // 数字格式
        const numberFormat = source.formatAt(index);
        if (numberFormat && numberFormat !== 'General') {
            cell.numberFormat = numberFormat;
        }

        return cell;
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
        // readImages 必须传给 decode：媒体条目在解压阶段就被跳过了
        const decoded = await this.decode(file, options);
        const sheets = await this.extract(decoded, options);
        return await this.generate(sheets, options);
    }

    // ════════════════════════════════════════════════════════════
    //  XML 解析辅助方法（xlsx 路径，全部基于流式扫描器）
    // ════════════════════════════════════════════════════════════

    /**
     * 解析 styles.xml 样式表（流式扫描）
     *
     * 与旧 DOM 版语义对齐的取舍（保持行为一致，含既有癖性）：
     * - font / fill / border 元素**不区分容器**（与 `getElementsByTagName`
     *   会连 dxfs 中同名元素一起收集的行为一致）；
     * - cellXfs 中的 xf 仅在 `cellXfs` 容器内生效（cellStyleXfs 中的不算）；
     * - `<b/>`（无 val）按 `isOn(undefined) === false` 处理，与旧实现一致。
     *
     * @param bytes - styles.xml 的字节
     * @returns 样式表结构
     * @private
     */
    private parseStyleSheetBytes(bytes: Uint8Array): ExcelStyleSheet {
        const result: ExcelStyleSheet = {
            numFmts: new Map(), fonts: [], fills: [], borders: [], cellXfs: []
        };

        // —— 当前累积对象 ——
        let font: NonNullable<ExcelStyleSheet['fonts'][number]> | null = null;
        let fill: NonNullable<ExcelStyleSheet['fills'][number]> | null = null;
        let inPatternFill = false;
        let border: NonNullable<ExcelStyleSheet['borders'][number]> | null = null;
        let borderSide: { style?: string; color?: string } | null = null;
        let borderSideName: 'left' | 'right' | 'top' | 'bottom' = 'left';
        let xf: NonNullable<ExcelStyleSheet['cellXfs'][number]> | null = null;
        let inCellXfs = false;

        scanXml(bytes, {
            onStart: (name, attrs) => {
                switch (name) {
                    case 'numFmt':
                        result.numFmts.set(
                            parseInt(attrs['numFmtId'] || '0', 10),
                            attrs['formatCode'] || ''
                        );
                        break;
                    case 'font':
                        font = {
                            name: undefined, size: undefined, bold: false, italic: false,
                            color: undefined, underline: false, strikethrough: false
                        };
                        break;
                    case 'b':
                        if (font) font.bold = this.isOn(attrValue(attrs, 'val'));
                        break;
                    case 'i':
                        if (font) font.italic = this.isOn(attrValue(attrs, 'val'));
                        break;
                    case 'u':
                        if (font) font.underline = true;
                        break;
                    case 'strike':
                        if (font) font.strikethrough = true;
                        break;
                    case 'name':
                        if (font) font.name = attrValue(attrs, 'val') || undefined;
                        break;
                    case 'sz':
                        if (font) font.size = parseFloat(attrValue(attrs, 'val') || '0') || undefined;
                        break;
                    case 'color':
                        // 字体色 / 边框色：<color rgb="..."/>（按当前上下文分派）
                        if (font) font.color = this.rgbFromAttrs(attrs);
                        else if (borderSide) borderSide.color = this.rgbFromAttrs(attrs);
                        break;
                    case 'fill':
                        fill = { patternType: undefined, fgColor: undefined };
                        break;
                    case 'patternFill':
                        if (fill) {
                            inPatternFill = true;
                            fill.patternType = attrs['patternType'] || undefined;
                        }
                        break;
                    case 'fgColor':
                        if (inPatternFill && fill) fill.fgColor = this.rgbFromAttrs(attrs);
                        break;
                    case 'border':
                        border = { left: undefined, right: undefined, top: undefined, bottom: undefined };
                        break;
                    case 'left':
                    case 'right':
                    case 'top':
                    case 'bottom':
                        if (border) {
                            borderSide = { style: attrs['style'] || undefined, color: undefined };
                            borderSideName = name;
                        }
                        break;
                    case 'cellXfs':
                        inCellXfs = true;
                        break;
                    case 'xf':
                        if (inCellXfs) {
                            xf = {
                                numFmtId: parseInt(attrs['numFmtId'] || '0', 10),
                                fontId: parseInt(attrs['fontId'] || '0', 10),
                                fillId: parseInt(attrs['fillId'] || '0', 10),
                                borderId: parseInt(attrs['borderId'] || '0', 10),
                                applyFont: this.isOn(attrs['applyFont']),
                                applyFill: this.isOn(attrs['applyFill']),
                                applyBorder: this.isOn(attrs['applyBorder']),
                                applyAlignment: this.isOn(attrs['applyAlignment']),
                                alignment: undefined
                            };
                        }
                        break;
                    case 'alignment':
                        if (xf) {
                            xf.alignment = {
                                horizontal: attrs['horizontal'] || undefined,
                                vertical: attrs['vertical'] || undefined,
                                wrapText: this.isOn(attrs['wrapText']),
                                indent: attrs['indent'] ? parseInt(attrs['indent'], 10) : undefined
                            };
                        }
                        break;
                }
            },
            onEnd: (name) => {
                switch (name) {
                    case 'font':
                        if (font) { result.fonts.push(font); font = null; }
                        break;
                    case 'fill':
                        if (fill) { result.fills.push(fill); fill = null; }
                        inPatternFill = false;
                        break;
                    case 'patternFill':
                        inPatternFill = false;
                        break;
                    case 'border':
                        if (border) { result.borders.push(border); border = null; }
                        break;
                    case 'left':
                    case 'right':
                    case 'top':
                    case 'bottom':
                        if (borderSide && border) {
                            border[borderSideName] = borderSide;
                            borderSide = null;
                        }
                        break;
                    case 'cellXfs':
                        inCellXfs = false;
                        break;
                    case 'xf':
                        if (xf) { result.cellXfs.push(xf); xf = null; }
                        break;
                }
            }
        });

        return result;
    }

    /**
     * 从属性表中取 `rgb` 并转为 `#RRGGBB`
     *
     * 与旧 DOM 版 `childRgb` 语义一致：indexed / theme / auto / 短值一律返回 undefined。
     *
     * @param attrs - 属性表
     * @returns 颜色；不可解析时 undefined
     * @private
     */
    private rgbFromAttrs(attrs: XmlAttrs): string | undefined {
        const rgb = attrs['rgb'];
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
     * 解析单个工作表 XML（sheetN.xml，流式扫描）
     *
     * 单趟扫描，内存占用 O(1)（不随表格大小增长）。取值分支与旧 DOM 版
     * 逐条对齐（含 r 属性缺失时的推断规则、空 `<v>` 的取值等癖性）。
     *
     * @param bytes - 工作表 XML 字节
     * @param sharedStrings - 共享字符串表
     * @param styleSheet - 样式表（可选）
     * @param name - 工作表名称
     * @returns 提取的工作表数据
     * @private
     */
    private parseWorksheetBytes(
        bytes: Uint8Array,
        sharedStrings: string[],
        styleSheet: ExcelStyleSheet | null,
        name: string
    ): ExtractedSheet {
        const cells: ExtractedCell[] = [];
        const merges: Array<{ s: { r: number; c: number }; e: { r: number; c: number } }> = [];
        const cols: Array<{ index: number; width?: number; customWidth?: boolean }> = [];
        const rowsData: Array<{ index: number; height?: number; customHeight?: boolean }> = [];
        let minRow = Infinity, maxRow = 0, minCol = Infinity, maxCol = 0;

        /**
         * 样式内联缓存：cellXfs 索引 → 已解析出的样式
         *
         * **按索引缓存，而不是逐格新建**（性能约定，勿回退）。历史实现每遇到一个
         * `<c s="N">` 就调一次 `styleFromXf`，而它每次都会新建
         * `style` + `font` + `fill` + `alignment` + `border`（最多 6 个对象）。
         * 实测 20 万格带样式（去重后仅 5000 种）时：提取阶段约 978 ms，
         * 且内存里躺着 20 万个内容重复的样式对象 —— 这是「样式多就卡」的主因之一。
         * 缓存后同一样式只解析一次，20 万格只剩 5000 个对象。
         *
         * 缓存里存的是 `{ style, numberFormat }`，其中 `numberFormat` 在
         * `resolve` 时会被单独取出（它是原始字符串，不需要克隆）。
         */
        const styleCache = new Map<number, { style?: CellStyleInfo; numberFormat?: string }>();

        // —— 扫描状态 ——
        let inRow = false;
        let rowIdx = -1;            // 当前行 0-based 行号（r 属性 - 1）
        let cellCounter = 0;        // 当前行内单元格序号（r 属性缺失时据此推断列号）
        let inCell = false;
        let cellRef = '';
        let cellTypeAttr = 'n';
        let cellStyleAttr: string | undefined;
        let hasV = false;           // 存在 <v> 直接子元素
        let vText = '';
        let hasIs = false;          // 存在 <is> 直接子元素
        let isText = '';
        let inV = false;
        let inIs = false;
        let inT = false;            // <is> 内的 <t>（内联字符串的文本载体）
        let inCols = false;
        let seenCols = false;       // 与 DOM 版「只读第一个 <cols>」对齐

        scanXml(bytes, {
            onStart: (el, attrs) => {
                switch (el) {
                    case 'row': {
                        rowIdx = parseInt(attrs['r'] || '0', 10) - 1;  // 转 0-based
                        if (rowIdx < minRow) minRow = rowIdx;
                        if (rowIdx > maxRow) maxRow = rowIdx;
                        // 行高：ht 单位为磅，换算为像素 = 磅 × 96 / 72
                        const ht = parseFloat(attrs['ht'] || '0') || undefined;
                        rowsData.push({
                            index: rowIdx,
                            height: ht ? Math.round(ht * 96 / 72) : undefined,
                            customHeight: attrs['customHeight'] === '1'
                        });
                        inRow = true;
                        cellCounter = 0;
                        break;
                    }
                    case 'c':
                        if (!inRow) break;
                        cellRef = attrs['r'] || '';
                        cellTypeAttr = attrs['t'] || 'n';
                        cellStyleAttr = attrs['s'];
                        hasV = false;
                        vText = '';
                        hasIs = false;
                        isText = '';
                        inCell = true;
                        cellCounter++;
                        break;
                    case 'v':
                        if (inCell) { inV = true; hasV = true; }
                        break;
                    case 'is':
                        if (inCell) { inIs = true; hasIs = true; }
                        break;
                    case 't':
                        // <is> 内任意层级的 <t> 都计入（与 DOM 版
                        // extractInlineText 取全部后代 <t> 的语义一致）
                        if (inIs) inT = true;
                        break;
                    case 'mergeCell': {
                        const ref = attrs['ref'] || '';
                        const parts = ref.split(':');
                        if (parts.length === 2) {
                            const s = this.refToColRow(parts[0]);
                            const e = this.refToColRow(parts[1]);
                            merges.push({ s: { r: s.row, c: s.col }, e: { r: e.row, c: e.col } });
                        }
                        break;
                    }
                    case 'cols':
                        if (!seenCols) { inCols = true; seenCols = true; }
                        break;
                    case 'col': {
                        if (!inCols) break;
                        // 列宽以「最大数字宽度字符数」为单位：像素 ≈ width * 7 + 5
                        const min = parseInt(attrs['min'] || '0', 10);
                        const max = parseInt(attrs['max'] || '0', 10);
                        const width = parseFloat(attrs['width'] || '0') || undefined;
                        const customWidth = attrs['customWidth'] === '1';
                        for (let c = min; c <= max; c++) {
                            cols.push({
                                index: c - 1,  // 0-based
                                width: width ? Math.round(width * 7 + 5) : undefined,
                                customWidth
                            });
                        }
                        break;
                    }
                }
            },
            onEnd: (el) => {
                switch (el) {
                    case 'row':
                        inRow = false;
                        break;
                    case 'v':
                        inV = false;
                        break;
                    case 'is':
                        inIs = false;
                        break;
                    case 't':
                        inT = false;
                        break;
                    case 'cols':
                        inCols = false;
                        break;
                    case 'c': {
                        if (!inCell) break;
                        inCell = false;

                        // — 定位：r 属性缺失时按行内序号推断；存在时以 ref 为准 —
                        const pos = this.refToColRow(cellRef);
                        const row = cellRef ? pos.row : rowIdx;
                        const col = cellRef ? pos.col : cellCounter - 1;
                        if (col < minCol) minCol = col;
                        if (col > maxCol) maxCol = col;

                        // — 读取值（分支顺序与旧 DOM 版一致）—
                        let value: string | number | boolean | null = null;
                        let cellType: ExtractedCell['type'] = 'string';

                        if (cellTypeAttr === 's' && hasV) {
                            // 共享字符串
                            value = sharedStrings[parseInt(vText || '0', 10)] ?? '';
                            cellType = 'string';
                        } else if (cellTypeAttr === 'inlineStr' && hasIs) {
                            // 内联字符串
                            value = isText;
                            cellType = 'string';
                        } else if (cellTypeAttr === 'str' && hasV) {
                            // 公式的字符串结果
                            value = vText || '';
                            cellType = 'string';
                        } else if (cellTypeAttr === 'b' && hasV) {
                            value = vText === '1';
                            cellType = 'boolean';
                        } else if (cellTypeAttr === 'e' && hasV) {
                            value = vText || '';
                            cellType = 'error';
                        } else if (cellTypeAttr === 'd' && hasV) {
                            // ISO 8601 日期
                            value = vText || '';
                            cellType = 'date';
                        } else if (hasV) {
                            const num = parseFloat(vText || '0');
                            value = isNaN(num) ? (vText || '') : num;
                            cellType = 'number';
                        }

                        // — 样式：通过 s 属性（样式索引）查 cellXfs —
                        // 按索引走内联缓存：同一样式只解析/新建一次对象（见 styleCache）
                        let style: CellStyleInfo | undefined;
                        let styleId: number | undefined;
                        let numberFormat: string | undefined;
                        if (cellStyleAttr !== undefined && styleSheet) {
                            const idx = parseInt(cellStyleAttr, 10);
                            let cached = styleCache.get(idx);
                            if (cached === undefined) {
                                cached = this.styleFromXf(styleSheet, idx);
                                styleCache.set(idx, cached);
                            }
                            style = cached.style;
                            styleId = idx;
                            numberFormat = cached.numberFormat;
                        }

                        cells.push({ row, col, value, type: cellType, style, styleId, numberFormat });
                        break;
                    }
                }
            },
            onText: (text) => {
                if (inV) vText += text;
                else if (inIs && inT) isText += text;
            }
        });

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
    private extractSheetImagesBytes(
        files: Map<string, Uint8Array>,
        sheetPath: string,
        drawingRelPaths: Map<string, string>
    ): ExtractedImage[] {
        // —— 读取工作表自身的 rels，定位 drawing ——
        const sheetRelsPath = sheetPath
            .replace('worksheets/', 'worksheets/_rels/')
            .replace('.xml', '.xml.rels');
        const sheetRelsBytes = files.get(sheetRelsPath);
        if (!sheetRelsBytes) return [];

        // 关系目标（Target）是**相对于所在目录**的，Excel 一贯写 `../drawings/...`
        const sheetDir = SheetParser.dirOf(sheetPath);

        let drawingPath = '';
        scanXml(sheetRelsBytes, {
            onStart: (name, attrs) => {
                if (name !== 'Relationship' || drawingPath) return;
                const id = attrs['Id'] || '';
                const target = attrs['Target'] || '';
                if (target.includes('drawing')) {
                    // 优先使用 workbook rels 中的映射（更可靠）
                    drawingPath = drawingRelPaths.get(id) ||
                        SheetParser.resolveZipTarget(sheetDir, target);
                }
            }
        });

        if (!drawingPath) return [];
        const drawingBytes = files.get(drawingPath);
        if (!drawingBytes) return [];

        // —— 读取 drawing 的 rels：rId → media 路径 ——
        const drawingRelsPath = drawingPath.replace('drawings/', 'drawings/_rels/')
            .replace('.xml', '.xml.rels');
        const drawingRelsBytes = files.get(drawingRelsPath);
        if (!drawingRelsBytes) return [];

        const drawingDir = SheetParser.dirOf(drawingPath);

        const mediaMap = new Map<string, string>();
        scanXml(drawingRelsBytes, {
            onStart: (name, attrs) => {
                if (name !== 'Relationship') return;
                const id = attrs['Id'] || '';
                const target = attrs['Target'] || '';
                if (target.includes('media/')) {
                    const resolved = SheetParser.resolveZipTarget(drawingDir, target);
                    if (resolved) mediaMap.set(id, resolved);
                }
            }
        });

        if (mediaMap.size === 0) return [];

        // —— 遍历 pic 元素提取图片（锚点坐标 + 尺寸 + blip 引用 + 名称）——
        // DOM 版靠「向上回溯父链」找 pic 所属锚点；流式扫描改为
        // 「进入锚点时重置全部坐标累积，pic 结束时取当前锚点」——两者等价，
        // 且锚点之外的 pic 同样回退为全零坐标（与 DOM 版行为一致）。
        //
        // 【三种锚点，勿回退为「只读 from」】
        // - `twoCellAnchor`：`<from>` + `<to>` 各自四个字段 → 右下角由 `to` 决定；
        // - `oneCellAnchor`：`<from>` + 锚点直属 `<ext cx cy>`（EMU）→ 尺寸由 ext 决定；
        // - `absoluteAnchor`：坐标以 EMU 从工作表原点算起，不落在单元格上，
        //   沿用历史行为（全零锚点）而不做像素换算。
        // 只取 `<from>` 会让所有图片退化成「左上角已知、尺寸未知」，
        // 渲染端只能瞎猜尺寸（历史上正是如此，故一直画不出来）。
        const images: ExtractedImage[] = [];
        /** 当前锚点种类（空串 = 不在任何锚点内） */
        let anchorKind: '' | 'twoCell' | 'oneCell' | 'absolute' = '';
        /** 是否真的见到过 `<to>` / `<ext>`（缺失时不得写出半截坐标） */
        let hasTo = false;
        let hasExt = false;
        let inFrom = false;
        let inTo = false;
        let fromField = '';         // 正在接收文本的 from 子元素（col/colOff/row/rowOff）
        let toField = '';           // 同上，to 子元素
        let fromCol = 0, fromRow = 0, fromColOff = 0, fromRowOff = 0;
        let toCol = 0, toRow = 0, toColOff = 0, toRowOff = 0;
        let extCx = 0, extCy = 0;
        let inPic = false;
        let picEmbed = '';
        let picName: string | undefined;

        scanXml(drawingBytes, {
            onStart: (name, attrs) => {
                switch (name) {
                    case 'twoCellAnchor':
                    case 'oneCellAnchor':
                    case 'absoluteAnchor':
                        anchorKind = name === 'twoCellAnchor' ? 'twoCell'
                            : name === 'oneCellAnchor' ? 'oneCell' : 'absolute';
                        hasTo = false;
                        hasExt = false;
                        inFrom = false; inTo = false;
                        fromField = ''; toField = '';
                        fromCol = 0; fromRow = 0; fromColOff = 0; fromRowOff = 0;
                        toCol = 0; toRow = 0; toColOff = 0; toRowOff = 0;
                        extCx = 0; extCy = 0;
                        break;
                    case 'from':
                        if (anchorKind) inFrom = true;
                        break;
                    case 'to':
                        if (anchorKind) { inTo = true; hasTo = true; }
                        break;
                    case 'ext':
                        // oneCellAnchor 的尺寸：`<xdr:ext cx cy/>`（EMU）。
                        // 必须是锚点的直属子元素：`<pic>` 内的 `<a:xfrm>` 也有 `ext`，
                        // 若不加 `!inPic` 判定会把「图片内部原始尺寸」误当成锚点尺寸。
                        if (anchorKind === 'oneCell' && !inPic) {
                            extCx = parseInt(attrValue(attrs, 'cx') || '0', 10) || 0;
                            extCy = parseInt(attrValue(attrs, 'cy') || '0', 10) || 0;
                            hasExt = true;
                        }
                        break;
                    case 'col':
                    case 'colOff':
                    case 'row':
                    case 'rowOff':
                        if (inFrom) fromField = name;
                        else if (inTo) toField = name;
                        break;
                    case 'pic':
                        inPic = true;
                        picEmbed = '';
                        picName = undefined;
                        break;
                    case 'blip':
                        if (inPic) picEmbed = attrValue(attrs, 'r:embed') || '';
                        break;
                    case 'cNvPr':
                        if (inPic) picName = attrs['name'] || undefined;
                        break;
                }
            },
            onEnd: (name) => {
                switch (name) {
                    case 'twoCellAnchor':
                    case 'oneCellAnchor':
                    case 'absoluteAnchor':
                        anchorKind = '';
                        inFrom = false; inTo = false;
                        fromField = ''; toField = '';
                        break;
                    case 'from':
                        inFrom = false;
                        fromField = '';
                        break;
                    case 'to':
                        inTo = false;
                        toField = '';
                        break;
                    case 'col':
                    case 'colOff':
                    case 'row':
                    case 'rowOff':
                        fromField = '';
                        toField = '';
                        break;
                    case 'pic': {
                        inPic = false;
                        // 只提取「落在单元格上」的图片：absoluteAnchor 以 EMU 从工作表
                        // 原点定位，锚点之外的 pic 属畸形结构，两者都没有列/行可依附，
                        // 硬按全零锚点画到 A1 只会是错误的位置（历史实现正是如此，因无人
                        // 消费 images 而未被发现）。渲染通道建好后必须在这里挡住。
                        if (anchorKind !== 'twoCell' && anchorKind !== 'oneCell') break;
                        const mediaPath = mediaMap.get(picEmbed);
                        if (!mediaPath) break;
                        const imgData = files.get(mediaPath);
                        if (!imgData) break;

                        // 从文件名推断 MIME
                        const ext = mediaPath.split('.').pop()?.toLowerCase() || '';
                        const mimeMap: Record<string, string> = {
                            png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg',
                            gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp',
                            tif: 'image/tiff', tiff: 'image/tiff', emf: 'image/emf', wmf: 'image/wmf'
                        };

                        const image: ExtractedImage = {
                            data: imgData,
                            mimeType: mimeMap[ext] || 'image/png',
                            anchor: {
                                col: fromCol, row: fromRow,
                                colOffset: fromColOff, rowOffset: fromRowOff
                            },
                            name: picName
                        };
                        // 尺寸信息按锚点种类分流：twoCell 才有 `to`，oneCell 才有 `ext`。
                        // 缺 `<to>` 的 twoCellAnchor（Excel 极少见，但属合法写法）保留
                        // 全零的 to 会让渲染端算出零/负宽高，故只在真读到时才写出。
                        if (anchorKind === 'twoCell' && hasTo) {
                            image.to = {
                                col: toCol, row: toRow,
                                colOffset: toColOff, rowOffset: toRowOff
                            };
                        } else if (anchorKind === 'oneCell' && hasExt) {
                            image.ext = { cx: extCx, cy: extCy };
                        }

                        images.push(image);
                        break;
                    }
                }
            },
            onText: (text) => {
                // 锚点坐标（EMU 单位）
                const value = parseInt(text, 10) || 0;
                if (fromField) {
                    if (fromField === 'col') fromCol = value;
                    else if (fromField === 'colOff') fromColOff = value;
                    else if (fromField === 'row') fromRow = value;
                    else if (fromField === 'rowOff') fromRowOff = value;
                } else if (toField) {
                    if (toField === 'col') toCol = value;
                    else if (toField === 'colOff') toColOff = value;
                    else if (toField === 'row') toRow = value;
                    else if (toField === 'rowOff') toRowOff = value;
                }
            }
        });

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
        const cached = this._colNameCache[col];
        if (cached !== undefined) return cached;
        // 顺序补齐到 col：每列只真正计算一次（摊还 O(1)）
        for (let n = this._colNameCache.length; n <= col; n++) {
            this._colNameCache[n] = SheetParser.computeColName(n);
        }
        return this._colNameCache[col];
    }

    /**
     * 计算列名（A, B, ..., Z, AA, AB...）——无缓存的纯函数版本
     *
     * @param col - 1-based 列号
     * @returns 列名
     * @private
     */
    private static computeColName(col: number): string {
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
