/*  # 文档域 · docx 编解码 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目目标: 让「块树 ⇄ .docx」双向互通：打开 Word 文档得到块树，块树能另存为 Word 文档。
    # 本文件职责: 两条单向转换
    #             - {@link docxToSnapshot}：解压 → 扫描 `word/document.xml` → 块树
    #             - {@link snapshotToDocx}：块树 → OOXML 部件 → ZIP（复用通用 ZIP 写入器）
    # 依赖方向: xmlScanner（流式 XML 扫描，**禁 DOMParser**）、zip（解压）、
    #           xlsxWriter 的 createZipBlob（通用 ZIP 打包）、同目录 blockTree。
    # 关键约定:
    #   - 段落级语义（标题 / 列表 / 对齐）走 `w:pPr`，字符级样式走 `w:rPr`，与块树一一对应；
    #   - 列表用 `w:numPr` + 自带 `word/numbering.xml`，不把项目符号写进正文文本 ——
    #     否则「docx → 块树 → docx」会把符号变成正文，来回一次多一层圆点。
    #   - 字号换算：OOXML 的 `w:sz` 是**半磅**，块树存 **CSS 像素**。
    #     12pt = 24 半磅 = 16px，即 `px = 半磅 / 2 * 4 / 3`、`半磅 = px * 1.5`。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { Char } from '../../dataArchitecture/char';
import { scanXml } from '../../../utils/parser/xmlScanner';
import { ZipParser } from '../../../utils/parser/zip';
import { createZipBlob, type ZipEntryInput } from '../../../utils/parser/xlsxWriter';
import {
    CharStylePatch,
    DocBlock,
    DocSnapshot,
    DocSpan,
    blockText,
    charsFromText,
    createBlankSnapshot,
    createBlock,
    normalizeBlocks,
    writeCharStyle,
} from '../blockTree';

/** 主文档部件在 ZIP 内的路径 */
const DOCUMENT_PART = 'word/document.xml';

/** 命名空间（写入侧用） */
const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** XML 声明 */
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/**
 * 解码 UTF-8 字节
 * @param bytes - 字节
 * @returns 文本
 */
function decodeUtf8(bytes: Uint8Array): string {
    return new TextDecoder('utf-8').decode(bytes);
}

/**
 * 解析 .docx 为文档存档
 *
 * 只读 `word/document.xml`：样式表（`styles.xml`）里的「样式继承」不做展开 ——
 * 展开需要建立整张样式表并解析 `basedOn` 链，收益是「标题的字号更准」，
 * 代价是每次打开都要多解压 + 扫描一个部件。段落语义（是不是标题、几级）已经
 * 能从 `w:pStyle` 拿到，够用。
 * @param buffer - .docx 文件字节
 * @param title - 文档标题（通常取文件名）
 * @returns 文档存档
 * @throws {Error} 不是 ZIP 包、或缺少 `word/document.xml`
 */
export async function docxToSnapshot(buffer: Uint8Array, title: string): Promise<DocSnapshot> {
    const zip = new ZipParser(buffer);
    const archive = zip.parseCentralDirectory();
    const bytes = await zip.readEntry(archive, DOCUMENT_PART);
    if (!bytes) throw new Error('不是有效的 .docx（缺少 word/document.xml）');

    const blocks = parseDocumentXml(decodeUtf8(bytes));
    const snapshot = createBlankSnapshot(title);
    snapshot.blocks = normalizeBlocks(blocks);
    return snapshot;
}

/**
 * 扫描 `word/document.xml` 并组装块树
 * @param xml - 部件文本
 * @returns 块树
 */
function parseDocumentXml(xml: string): DocBlock[] {
    const root: DocBlock[] = [];
    /** 段落目标栈：根数组，或某个单元格块的 children */
    const containers: DocBlock[][] = [root];
    /** 当前段落 */
    let paragraph: DocBlock | null = null;
    /** 当前段落的段落属性（在 `w:pPr` 期间累积） */
    let paraStyle = '';
    let paraAlign: 'left' | 'center' | 'right' | 'justify' | undefined;
    let paraListLevel: number | null = null;
    let paraIndent = 0;
    /** 当前 run 的字符样式 */
    let run: CharStylePatch = {};
    /** 当前 run 的高亮 / 上下标（落 DocSpan） */
    let runHighlight: string | undefined;
    let runVertAlign: 'superscript' | 'subscript' | undefined;
    /** 当前表格 */
    let table: DocBlock | null = null;
    /** 当前表格行 */
    let row: DocBlock[] | null = null;
    /** 是否位于 `w:rPr` 内（只有 rPr 里的 b/i/u 才是字符样式） */
    let inRunProps = false;
    /** 是否位于 `w:t` 内（只有 t 的文本才是正文） */
    let inText = false;
    /** 文本段内的暂存（`w:t` 的文本可能分多次回调） */
    let textBuffer = '';

    const currentContainer = (): DocBlock[] => containers[containers.length - 1];

    /** 把文本追加到当前段落 */
    const appendText = (text: string): void => {
        if (!paragraph || !text) return;
        const start = paragraph.chars.length;
        const created = charsFromText(text);
        for (let i = 0; i < created.length; i++) {
            writeCharStyle(created[i], run);
            paragraph.chars.push(created[i]);
        }
        const end = paragraph.chars.length;
        if (end > start && (runHighlight || runVertAlign)) {
            const span: DocSpan = { start, end };
            if (runHighlight) span.highlight = runHighlight;
            if (runVertAlign) span.vertAlign = runVertAlign;
            if (!paragraph.spans) paragraph.spans = [];
            paragraph.spans.push(span);
        }
    };

    /** 结束当前段落：按累积的段落属性定型 */
    const finishParagraph = (): void => {
        if (!paragraph) return;
        const style = paraStyle.toLowerCase();
        const headingLevel = /^heading\s*(\d)$/.exec(style) || /^标题\s*(\d)$/.exec(style);
        if (headingLevel) {
            paragraph.type = 'heading';
            paragraph.level = Math.max(1, Math.min(6, parseInt(headingLevel[1], 10)));
        } else if (style === 'title') {
            paragraph.type = 'heading';
            paragraph.level = 1;
        } else if (paraListLevel !== null) {
            paragraph.type = 'listItem';
            paragraph.level = Math.max(0, Math.min(4, paraListLevel));
            paragraph.listStyle = 'bullet';
        }
        if (paraAlign) paragraph.align = paraAlign;
        if (paraIndent > 0) paragraph.indent = paraIndent;
        paragraph = null;
        paraStyle = '';
        paraAlign = undefined;
        paraListLevel = null;
        paraIndent = 0;
    };

    scanXml(new TextEncoder().encode(xml), {
        onStart(name, attrs) {
            switch (name) {
                case 'tbl': {
                    finishParagraph();
                    const block = createBlock({ type: 'table' });
                    block.rows = [];
                    currentContainer().push(block);
                    table = block;
                    break;
                }
                case 'tr': {
                    row = [];
                    if (table) (table.rows as DocBlock[][]).push(row);
                    break;
                }
                case 'tc': {
                    const cell = createBlock({ type: 'paragraph' });
                    if (!cell.children) cell.children = [];
                    if (!row) row = [];
                    row.push(cell);
                    containers.push(cell.children as DocBlock[]);
                    break;
                }
                case 'p': {
                    const block = createBlock({ type: 'paragraph' });
                    currentContainer().push(block);
                    paragraph = block;
                    break;
                }
                case 'pStyle':
                    paraStyle = attrs['w:val'] ?? '';
                    break;
                case 'jc': {
                    const value = (attrs['w:val'] ?? '').toLowerCase();
                    if (value === 'center' || value === 'right' || value === 'both' || value === 'left') {
                        paraAlign = value === 'both' ? 'justify' : (value as 'left' | 'center' | 'right');
                    }
                    break;
                }
                case 'numPr':
                    paraListLevel = 0;
                    break;
                case 'ilvl':
                    paraListLevel = parseInt(attrs['w:val'] ?? '0', 10) || 0;
                    break;
                case 'ind': {
                    const left = parseInt(attrs['w:left'] ?? attrs['w:start'] ?? '0', 10) || 0;
                    if (left > 0) paraIndent = Math.round(left / 420);
                    break;
                }
                case 'r':
                    run = {};
                    runHighlight = undefined;
                    runVertAlign = undefined;
                    break;
                case 'rPr':
                    inRunProps = true;
                    break;
                case 'b':
                    if (inRunProps) run.bold = (attrs['w:val'] ?? '1') !== '0';
                    break;
                case 'i':
                    if (inRunProps) run.italic = (attrs['w:val'] ?? '1') !== '0';
                    break;
                case 'u': {
                    if (inRunProps) {
                        const value = (attrs['w:val'] ?? 'single').toLowerCase();
                        run.underline = value !== 'none';
                    }
                    break;
                }
                case 'strike':
                    if (inRunProps) run.strike = (attrs['w:val'] ?? '1') !== '0';
                    break;
                case 'color': {
                    if (inRunProps) {
                        const value = attrs['w:val'] ?? '';
                        if (value && value.toLowerCase() !== 'auto') run.color = `#${value.toUpperCase()}`;
                    }
                    break;
                }
                case 'sz': {
                    if (inRunProps) {
                        const half = parseInt(attrs['w:val'] ?? '0', 10);
                        if (half > 0) run.fontSize = Math.round((half / 2) * 4 / 3);
                    }
                    break;
                }
                case 'rFonts': {
                    if (inRunProps) {
                        const family = attrs['w:ascii'] ?? attrs['w:hAnsi'];
                        if (family) run.fontFamily = family;
                    }
                    break;
                }
                case 'highlight': {
                    const value = (attrs['w:val'] ?? '').toLowerCase();
                    if (value && value !== 'none') runHighlight = HIGHLIGHT_COLORS[value] ?? '#FFFF00';
                    break;
                }
                case 'shd': {
                    const fill = attrs['w:fill'];
                    if (fill && fill.toLowerCase() !== 'auto') runHighlight = `#${fill.toUpperCase()}`;
                    break;
                }
                case 'vertAlign': {
                    const value = (attrs['w:val'] ?? '').toLowerCase();
                    if (value === 'superscript') runVertAlign = 'superscript';
                    else if (value === 'subscript') runVertAlign = 'subscript';
                    break;
                }
                case 't':
                    inText = true;
                    textBuffer = '';
                    break;
                case 'tab':
                    appendText('\t');
                    break;
                case 'br':
                    appendText('\n');
                    break;
                default:
                    break;
            }
        },
        onText(text) {
            if (inText) textBuffer += text;
        },
        onEnd(name) {
            switch (name) {
                case 't':
                    inText = false;
                    appendText(textBuffer);
                    textBuffer = '';
                    break;
                case 'rPr':
                    inRunProps = false;
                    break;
                case 'p':
                    finishParagraph();
                    break;
                case 'tc':
                    if (containers.length > 1) containers.pop();
                    break;
                case 'tr':
                    row = null;
                    break;
                case 'tbl':
                    table = null;
                    break;
                default:
                    break;
            }
        },
    });

    return root;
}

/** OOXML 高亮命名色 → CSS 颜色 */
const HIGHLIGHT_COLORS: Record<string, string> = {
    yellow: '#FFFF00',
    green: '#00FF00',
    cyan: '#00FFFF',
    magenta: '#FF00FF',
    blue: '#0000FF',
    red: '#FF0000',
    darkblue: '#000080',
    darkcyan: '#008080',
    darkgreen: '#008000',
    darkmagenta: '#800080',
    darkred: '#800000',
    darkyellow: '#808000',
    darkgray: '#808080',
    lightgray: '#C0C0C0',
    black: '#000000',
};

/**
 * 块树 → .docx
 * @param snapshot - 文档存档
 * @returns docx 文件（ZIP 包）
 */
export function snapshotToDocx(snapshot: DocSnapshot): Blob {
    const links: string[] = [];
    const body = (snapshot.blocks ?? []).map((block) => renderBlock(block, links)).join('');
    const document = `${XML_HEAD}<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>${body}`
        + `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>`
        + `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>`
        + `</w:body></w:document>`;

    const entries: ZipEntryInput[] = [
        { name: '[Content_Types].xml', data: encode(contentTypesXml()) },
        { name: '_rels/.rels', data: encode(packageRelsXml()) },
        { name: 'word/document.xml', data: encode(document) },
        { name: 'word/styles.xml', data: encode(stylesXml()) },
        { name: 'word/numbering.xml', data: encode(numberingXml()) },
        { name: 'word/_rels/document.xml.rels', data: encode(documentRelsXml(links)) },
    ];
    // 不压缩（store）也能被 Word / WPS 正常打开；好处是不依赖 CompressionStream，
    // 且导出的包体积对文档量级（10²–10⁴ 块）完全可接受
    return createZipBlob(entries);
}

/**
 * 字符串 → UTF-8 字节
 * @param text - 文本
 * @returns 字节
 */
function encode(text: string): Uint8Array {
    return new TextEncoder().encode(text);
}

/**
 * 渲染单个块
 * @param block - 块
 * @param links - 超链接收集器（用于生成关系表）
 * @returns OOXML 片段
 */
function renderBlock(block: DocBlock, links: string[]): string {
    switch (block.type) {
        case 'table':
            return renderTable(block, links);
        case 'hr':
            return `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:color="auto"/></w:pBdr></w:pPr></w:p>`;
        case 'code':
            return renderParagraph(block, links, { mono: true });
        case 'image':
            // 图片二进制在 images 仓，docx 侧本版不内嵌（会显著放大文件且需要额外的关系与绘图部件）
            return `<w:p><w:r><w:t xml:space="preserve">[图片]</w:t></w:r></w:p>`;
        default:
            return renderParagraph(block, links);
    }
}

/**
 * 渲染表格块
 * @param block - 表格块
 * @param links - 超链接收集器
 * @returns OOXML 片段
 */
function renderTable(block: DocBlock, links: string[]): string {
    const rows = block.rows ?? [];
    const columns = rows.reduce((max, row) => Math.max(max, row.length), 0);
    const grid = new Array(columns).fill('<w:gridCol w:w="2400"/>').join('');
    const body = rows.map((row) => {
        const cells = row.map((cell) => {
            const inner = (cell.children && cell.children.length > 0 ? cell.children : [cell])
                .map((child) => `<w:p><w:r><w:t xml:space="preserve">${escapeXml(blockText(child))}</w:t></w:r></w:p>`)
                .join('');
            return `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>${inner || '<w:p/>'}</w:tc>`;
        }).join('');
        return `<w:tr>${cells}</w:tr>`;
    }).join('');
    return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>`
        + `<w:tblGrid>${grid}</w:tblGrid>${body}</w:tbl>`;
}

/**
 * 渲染段落（含字符级样式分段）
 * @param block - 块
 * @param links - 超链接收集器
 * @param options - 附加选项（等宽）
 * @returns OOXML 片段
 */
function renderParagraph(
    block: DocBlock,
    links: string[],
    options: { mono?: boolean } = {}
): string {
    let props = '';
    if (block.type === 'heading') {
        props += `<w:pStyle w:val="Heading${Math.max(1, Math.min(6, block.level ?? 1))}"/>`;
    } else if (block.type === 'listItem') {
        const ordered = block.listStyle === 'ordered';
        props += `<w:pStyle w:val="${ordered ? 'ListNumber' : 'ListBullet'}"/>`;
        props += `<w:numPr><w:ilvl w:val="${Math.max(0, Math.min(4, block.level ?? 0))}"/>`
            + `<w:numId w:val="${ordered ? 2 : 1}"/></w:numPr>`;
    } else if (block.type === 'quote') {
        props += `<w:pStyle w:val="Quote"/>`;
    } else if (block.type === 'code') {
        props += `<w:pStyle w:val="Code"/>`;
    }
    if (block.align) {
        const map: Record<string, string> = { left: 'left', center: 'center', right: 'right', justify: 'both' };
        props += `<w:jc w:val="${map[block.align]}"/>`;
    }
    if (block.indent) props += `<w:ind w:left="${block.indent * 420}"/>`;
    const paraProps = props ? `<w:pPr>${props}</w:pPr>` : '';

    return `<w:p>${paraProps}${renderInline(block, links, options.mono === true)}</w:p>`;
}

/**
 * 渲染行内内容（按「同样式 + 同扩展区间」切成 run）
 * @param block - 块
 * @param links - 超链接收集器
 * @param mono - 是否强制等宽（代码块）
 * @returns OOXML 片段
 */
function renderInline(block: DocBlock, links: string[], mono: boolean): string {
    const chars = block.chars ?? [];
    if (chars.length === 0) return '';
    let out = '';
    let index = 0;
    while (index < chars.length) {
        const span = spanOf(block, index);
        const key = [
            chars[index].fontFamily ?? '',
            chars[index].fontSize ?? '',
            chars[index].fontWeight ? 1 : 0,
            chars[index].fontStyle ? 1 : 0,
            chars[index].underline ? 1 : 0,
            chars[index].strikethrough ? 1 : 0,
            chars[index].fontColor ?? '',
            span?.highlight ?? '',
            span?.link ?? '',
            span?.vertAlign ?? '',
        ].join('|');
        let end = index + 1;
        while (end < chars.length) {
            const next = spanOf(block, end);
            const nextKey = [
                chars[end].fontFamily ?? '',
                chars[end].fontSize ?? '',
                chars[end].fontWeight ? 1 : 0,
                chars[end].fontStyle ? 1 : 0,
                chars[end].underline ? 1 : 0,
                chars[end].strikethrough ? 1 : 0,
                chars[end].fontColor ?? '',
                next?.highlight ?? '',
                next?.link ?? '',
                next?.vertAlign ?? '',
            ].join('|');
            if (nextKey !== key) break;
            end += 1;
        }
        let text = '';
        for (let i = index; i < end; i++) text += String(chars[i].char ?? '');
        out += renderRun(text, chars[index], span, mono, links);
        index = end;
    }
    return out;
}

/**
 * 取某位置命中的扩展区间
 * @param block - 块
 * @param index - 码点下标
 * @returns 区间或 null
 */
function spanOf(block: DocBlock, index: number): DocSpan | null {
    if (!block.spans) return null;
    for (let i = 0; i < block.spans.length; i++) {
        const span = block.spans[i];
        if (index >= span.start && index < span.end) return span;
    }
    return null;
}

/**
 * 渲染单个 run
 * @param text - 文本
 * @param char - 样式来源字符
 * @param span - 扩展区间
 * @param mono - 是否强制等宽
 * @returns OOXML 片段
 */
function renderRun(
    text: string,
    char: Char,
    span: DocSpan | null,
    mono: boolean,
    links: string[]
): string {
    if (!text) return '';
    let props = '';
    const family = mono ? 'Consolas' : char.fontFamily;
    if (family) props += `<w:rFonts w:ascii="${escapeXml(family)}" w:hAnsi="${escapeXml(family)}"/>`;
    if (char.fontWeight) props += '<w:b/>';
    if (char.fontStyle) props += '<w:i/>';
    if (char.underline) props += '<w:u w:val="single"/>';
    if (char.strikethrough) props += '<w:strike/>';
    if (char.fontColor) props += `<w:color w:val="${toHex(char.fontColor)}"/>`;
    if (char.fontSize) props += `<w:sz w:val="${Math.max(2, Math.round(char.fontSize * 1.5))}"/>`;
    if (span?.highlight) props += `<w:shd w:val="clear" w:color="auto" w:fill="${toHex(span.highlight)}"/>`;
    if (span?.vertAlign) props += `<w:vertAlign w:val="${span.vertAlign}"/>`;
    const rPr = props ? `<w:rPr>${props}</w:rPr>` : '';
    const body = `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
    if (span?.link) {
        const id = `rIdLink${links.length + 1}`;
        links.push(span.link);
        return `<w:hyperlink r:id="${id}">${body}</w:hyperlink>`;
    }
    return body;
}

/**
 * CSS 颜色 → OOXML 的 6 位十六进制（不带 `#`）
 * @param color - CSS 颜色
 * @returns 6 位十六进制
 */
function toHex(color: string): string {
    const text = (color ?? '').trim();
    const short = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/i.exec(text);
    if (short) return `${short[1]}${short[1]}${short[2]}${short[2]}${short[3]}${short[3]}`.toUpperCase();
    if (/^#[0-9a-f]{6}$/i.test(text)) return text.slice(1).toUpperCase();
    const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i.exec(text);
    if (rgb) {
        const hex = (n: string): string => {
            const value = Math.max(0, Math.min(255, parseInt(n, 10)));
            return (value < 16 ? '0' : '') + value.toString(16).toUpperCase();
        };
        return hex(rgb[1]) + hex(rgb[2]) + hex(rgb[3]);
    }
    return '000000';
}

/**
 * XML 文本转义
 * @param text - 原始文本
 * @returns 转义后的文本
 */
function escapeXml(text: string): string {
    return (text ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

/**
 * `[Content_Types].xml`
 * @returns XML
 */
function contentTypesXml(): string {
    return `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
        + `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`
        + `<Default Extension="xml" ContentType="application/xml"/>`
        + `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>`
        + `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>`
        + `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>`
        + `</Types>`;
}

/**
 * 包级关系 `_rels/.rels`
 * @returns XML
 */
function packageRelsXml(): string {
    return `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
        + `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>`
        + `</Relationships>`;
}

/**
 * 主文档关系表（样式 / 编号 / 超链接）
 * @param links - 超链接地址（顺序对应 rIdLink1…）
 * @returns XML
 */
function documentRelsXml(links: string[]): string {
    const base = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    let rels = `<Relationship Id="rIdStyles" Type="${base}/styles" Target="styles.xml"/>`
        + `<Relationship Id="rIdNumbering" Type="${base}/numbering" Target="numbering.xml"/>`;
    for (let i = 0; i < links.length; i++) {
        rels += `<Relationship Id="rIdLink${i + 1}" Type="${base}/hyperlink"`
            + ` Target="${escapeXml(links[i])}" TargetMode="External"/>`;
    }
    return `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
}

/**
 * 样式表（标题 1–6 / 引用 / 代码 / 列表 / 正文）
 *
 * 自带一份最小样式表，而不是依赖 Word 的默认模板：`w:pStyle` 引用一个不存在的
 * 样式时，Word 会显示为正文（**静默降级**），标题层级就全丢了。
 * @returns XML
 */
function stylesXml(): string {
    const heading = (level: number): string => {
        const size = Math.max(20, 44 - level * 4);   // 半磅：H1=40pt? 这里取 20–40 的递减
        return `<w:style w:type="paragraph" w:styleId="Heading${level}">`
            + `<w:name w:val="heading ${level}"/><w:basedOn w:val="Normal"/><w:qFormat/>`
            + `<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="${level - 1}"/></w:pPr>`
            + `<w:rPr><w:b/><w:sz w:val="${size}"/></w:rPr></w:style>`;
    };
    return `${XML_HEAD}<w:styles xmlns:w="${NS_W}">`
        + `<w:docDefaults><w:rPrDefault><w:rPr>`
        + `<w:rFonts w:ascii="HarmonyOS Sans SC Regular" w:hAnsi="HarmonyOS Sans SC Regular" w:eastAsia="宋体"/>`
        + `<w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>`
        + `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>`
        + heading(1) + heading(2) + heading(3) + heading(4) + heading(5) + heading(6)
        + `<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/>`
        + `<w:pPr><w:ind w:left="420"/></w:pPr><w:rPr><w:i/><w:color w:val="555555"/></w:rPr></w:style>`
        + `<w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/>`
        + `<w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F4F4F4"/></w:pPr>`
        + `<w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/></w:rPr></w:style>`
        + `<w:style w:type="paragraph" w:styleId="ListBullet"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/></w:style>`
        + `<w:style w:type="paragraph" w:styleId="ListNumber"><w:name w:val="List Number"/><w:basedOn w:val="Normal"/></w:style>`
        + `<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>`
        + `<w:tblPr><w:tblBorders>`
        + `<w:top w:val="single" w:sz="4" w:color="auto"/><w:left w:val="single" w:sz="4" w:color="auto"/>`
        + `<w:bottom w:val="single" w:sz="4" w:color="auto"/><w:right w:val="single" w:sz="4" w:color="auto"/>`
        + `<w:insideH w:val="single" w:sz="4" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:color="auto"/>`
        + `</w:tblBorders></w:tblPr></w:style>`
        + `</w:styles>`;
}

/**
 * 编号定义（列表 1 = 圆点，列表 2 = 数字）
 * @returns XML
 */
function numberingXml(): string {
    const bulletLevels = (): string => {
        let out = '';
        for (let level = 0; level < 5; level++) {
            out += `<w:lvl w:ilvl="${level}"><w:start w:val="1"/><w:numFmt w:val="bullet"/>`
                + `<w:lvlText w:val="${level % 2 === 0 ? '\u2022' : '\u25E6'}"/>`
                + `<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${420 * (level + 1)}" w:hanging="420"/></w:pPr></w:lvl>`;
        }
        return out;
    };
    const numberLevels = (): string => {
        let out = '';
        for (let level = 0; level < 5; level++) {
            out += `<w:lvl w:ilvl="${level}"><w:start w:val="1"/><w:numFmt w:val="decimal"/>`
                + `<w:lvlText w:val="%${level + 1}."/>`
                + `<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${420 * (level + 1)}" w:hanging="420"/></w:pPr></w:lvl>`;
        }
        return out;
    };
    return `${XML_HEAD}<w:numbering xmlns:w="${NS_W}">`
        + `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${bulletLevels()}</w:abstractNum>`
        + `<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${numberLevels()}</w:abstractNum>`
        + `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`
        + `<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>`
        + `</w:numbering>`;
}
