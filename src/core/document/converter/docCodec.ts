/*  # 文档域 · doc（Word 97-2003 二进制）编解码 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目目标: 把 `.doc` 二进制文档还原成块树，使打开后的字号 / 字体 / 颜色 /
    #           高亮 / 段落缩进 / 对齐与 Word 中所见一致，而不是「一锅纯文本」。
    # 本文件职责: 单一入口 {@link docToSnapshot}：
    #             OLE2 复合文档 → FIB → 片段表（PlcPcd）解码正文 → 字体表（SttbfFfn）
    #             → 字符属性（CHPX）/ 段落属性（PAPX）→ 块树。
    # 依赖方向: 只依赖 cfb（OLE2 读取）与 blockTree / Char / 常量；**不碰 DOM**，
    #           因此可以整体搬进 Worker，也能在测试桩里直接跑。
    # 关键约定（与 MS-DOC 规范对应）:
    #   - 片段表：cp→fc 映射；fCompressed 片段每字符 1 字节，否则 UTF-16LE 2 字节；
    #   - 正文只取 `[0, ccpText)`：其后的页脚 / 域指令（`PAGE \* MERGEFORMAT`）不属于正文；
    #   - 域结构 `0x13 指令 0x14 结果 0x15`：指令段整体丢弃，结果段保留；
    #   - CHPX / PAPX 的 FKP 页存放在 `WordDocument` 流中（页号 × 512），
    #     字体表与片段表存放在 `0Table` / `1Table` 流中；
    #   - PAPX 的 `cb == 0` 是「双字节长度」形态（见 {@link readPapxProps}），
    #     按单字节长度读会把 istd 读成 2560 / 3328 这类垃圾值。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { DEFAULT_DOC_FONT_SIZE } from '../../constant';
import { CompoundFile } from '../../../utils/parser/cfb';
import { Char } from '../../dataArchitecture/char';
import {
    CharStylePatch,
    DocBlock,
    DocSnapshot,
    DocSpan,
    blockText,
    createBlankSnapshot,
    createBlock,
    normalizeBlocks,
    prefixBlockText,
    writeCharStyle,
} from '../blockTree';

/** Word 97 的 FIB 里用到的字段偏移（单位：字节，均从 `WordDocument` 流首字节起算） */
const FIB = {
    /** flags：bit 9（0x0200）为 1 时表流是 `1Table`，否则 `0Table` */
    flags: 0x0A,
    /** 正文（不含页眉页脚 / 脚注）的字符数 */
    ccpText: 0x4C,
    /** 片段表（CLX）在表流中的偏移与长度 */
    fcClx: 0x01A2,
    lcbClx: 0x01A6,
    /** 字符属性 PLC（PlcfBteChpx） */
    fcPlcfBteChpx: 0xFA,
    lcbPlcfBteChpx: 0xFE,
    /** 段落属性 PLC（PlcfBtePapx） */
    fcPlcfBtePapx: 0x102,
    lcbPlcfBtePapx: 0x106,
    /** 字体表（SttbfFfn） */
    fcSttbfFfn: 0x112,
    lcbSttbfFfn: 0x116,
} as const;

/** FKP（Formatted disk page）每页 512 字节 */
const FKP_SIZE = 512;

/** 字符调色板（Word 的 Ico 索引） */
const ICO_COLORS: Record<number, string> = {
    1: '#000000',
    2: '#0000FF',
    3: '#00FFFF',
    4: '#00FF00',
    5: '#FF00FF',
    6: '#FF0000',
    7: '#FFFF00',
    8: '#FFFFFF',
    9: '#000080',
    10: '#008080',
    11: '#008000',
    12: '#800080',
    13: '#800000',
    14: '#808000',
    15: '#808080',
    16: '#C0C0C0',
};

/** Word 97 正文里的控制字符 */
const CHAR_TAB = 0x09;
const CHAR_LINE_BREAK = 0x0B;
const CHAR_PAGE_BREAK = 0x0C;
const CHAR_CELL_MARK = 0x07;
const CHAR_PARAGRAPH_MARK = 0x0D;
const CHAR_FIELD_BEGIN = 0x13;
const CHAR_FIELD_SEPARATOR = 0x14;
const CHAR_FIELD_END = 0x15;

/**
 * 单个可见字符（带它在 `WordDocument` 流中的字节偏移）
 *
 * CHPX 是按 FC 区间登记的：没有 FC 就没法给字符找回字号 / 颜色，
 * 因此不能先把正文拼成字符串再处理。
 */
interface DocChar {
    /** 字符（一个码点） */
    text: string;
    /** 起始 FC（UTF-16 片段下是 2 字节对齐的偏移） */
    fc: number;
}

/** 一个 FKP 里的 CHPX 运行段（FC 区间 + 属性字节） */
interface ChpxRun {
    fcStart: number;
    fcEnd: number;
    /** grpprl（未解析的 sprm 序列；空数组表示该段无字符属性） */
    grpprl: number[];
}

/** 一个 FKP 里的 PAPX 运行段（FC 区间 + 段落属性） */
interface PapxRun {
    fcStart: number;
    fcEnd: number;
    /** istd（样式表索引）；本文档未展开样式表，只用于兜底判断 */
    istd: number;
    /** grpprl（sprm 序列） */
    grpprl: number[];
}

/**
 * 把 `.doc` 文件解析为文档存档
 *
 * @param buffer - `.doc` 文件字节（OLE2 复合文档）
 * @param title - 文档标题（通常取文件名）
 * @returns 文档存档
 * @throws {Error} 不是有效的 OLE2 文档 / 缺少 `WordDocument` 流
 *
 * @example
 * ```ts
 * const snapshot = docToSnapshot(new Uint8Array(await file.arrayBuffer()), '通知');
 * ```
 */
export function docToSnapshot(buffer: Uint8Array, title: string): DocSnapshot {
    const cfb = new CompoundFile(buffer);
    const stream = cfb.readStream('WordDocument');
    if (!stream) throw new Error('不是有效的 .doc（缺少 WordDocument 数据流）');

    const flags = readUint16(stream, FIB.flags);
    const table = cfb.readStream((flags & 0x0200) !== 0 ? '1Table' : '0Table');
    if (!table) throw new Error('不是有效的 .doc（缺少表流 0Table / 1Table）');

    const fonts = readFonts(table, readUint32(stream, FIB.fcSttbfFfn), readUint32(stream, FIB.lcbSttbfFfn));
    const chpxRuns = readChpxRuns(stream, table, readUint32(stream, FIB.fcPlcfBteChpx), readUint32(stream, FIB.lcbPlcfBteChpx));
    const papxRuns = readPapxRuns(stream, table, readUint32(stream, FIB.fcPlcfBtePapx), readUint32(stream, FIB.lcbPlcfBtePapx));

    const ccpText = readUint32(stream, FIB.ccpText);
    const chars = readDocChars(
        stream,
        table,
        readUint32(stream, FIB.fcClx),
        readUint32(stream, FIB.lcbClx),
        ccpText
    );

    const blocks = buildBlocks(chars, fonts, chpxRuns, papxRuns);
    const snapshot = createBlankSnapshot(title);
    snapshot.blocks = normalizeBlocks(blocks);
    return snapshot;
}

/**
 * 解码正文的可见字符（含 FC 映射）
 *
 * 只按片段表逐片段解码：整段按单一编码读会让中文变成乱码。
 * 解码完成后过滤域指令与对象占位符 —— 否则文末会冒出 `PAGE \* MERGEFORMAT`。
 * @param stream - `WordDocument` 流
 * @param table - 表流
 * @param fcClx - 片段表偏移
 * @param lcbClx - 片段表长度
 * @param ccpText - 正文字符数（不含页脚 / 脚注）
 * @returns 可见字符序列（尚未拆段）
 */
function readDocChars(
    stream: Uint8Array,
    table: Uint8Array,
    fcClx: number,
    lcbClx: number,
    ccpText: number
): DocChar[] {
    const pieces = readPieces(table, fcClx, lcbClx);
    const limit = ccpText > 0 ? ccpText : pieces.reduce((max, piece) => Math.max(max, piece.cpEnd), 0);
    const out: DocChar[] = [];

    /** 域嵌套层数 */
    let fieldDepth = 0;
    /** 是否处于域指令段（0x13 与 0x14 之间） */
    let inInstruction = false;

    for (let p = 0; p < pieces.length; p++) {
        const piece = pieces[p];
        if (piece.cpStart >= limit) break;
        const end = Math.min(piece.cpEnd, limit);
        for (let cp = piece.cpStart; cp < end; cp++) {
            const delta = cp - piece.cpStart;
            let code: number;
            let fc: number;
            if (piece.compressed) {
                fc = piece.fc + delta;
                code = stream[fc];
            } else {
                fc = piece.fc + delta * 2;
                code = readUint16(stream, fc);
            }
            if (code === CHAR_FIELD_BEGIN) {
                fieldDepth += 1;
                inInstruction = true;
                continue;
            }
            if (code === CHAR_FIELD_SEPARATOR) {
                if (fieldDepth > 0) inInstruction = false;
                continue;
            }
            if (code === CHAR_FIELD_END) {
                if (fieldDepth > 0) fieldDepth -= 1;
                inInstruction = false;
                continue;
            }
            if (fieldDepth > 0 && inInstruction) continue;
            // 图片 / 嵌入对象的占位符（0x01 / 0x08 等）没有可编辑文本
            if (code > 0 && code < 0x20 && code !== CHAR_TAB
                && code !== CHAR_LINE_BREAK && code !== CHAR_PAGE_BREAK
                && code !== CHAR_CELL_MARK && code !== CHAR_PARAGRAPH_MARK) {
                continue;
            }
            if (code === 0) continue;
            out.push({ text: String.fromCharCode(code), fc });
        }
    }
    return out;
}

/** 片段表中的一段：CP 区间 + 在 `WordDocument` 流中的起始 FC */
interface DocPiece {
    cpStart: number;
    cpEnd: number;
    fc: number;
    compressed: boolean;
}

/**
 * 解析片段表（CLX / PlcPcd）
 *
 * CLX 里可能先出现若干 Prc（`clxt = 1`），跳过它们才到 Pcdt（`clxt = 2`）。
 * PlcPcd 结构：`n + 1` 个 CP（每个 4 字节）+ `n` 个 PCD（每个 8 字节）；
 * PCD 的 `fcRaw` 第 30 位（0x40000000）为 fCompressed，压缩片段的真实偏移是
 * `(fcRaw & 0x3FFFFFFF) / 2`（规范里 fc 以 2 字节为单位存储）。
 * @param table - 表流
 * @param fcClx - CLX 偏移
 * @param lcbClx - CLX 长度
 * @returns 片段列表
 */
function readPieces(table: Uint8Array, fcClx: number, lcbClx: number): DocPiece[] {
    const result: DocPiece[] = [];
    if (fcClx <= 0 || lcbClx <= 0 || fcClx + lcbClx > table.length) return result;

    const end = fcClx + lcbClx;
    let cursor = fcClx;
    while (cursor + 3 <= end && table[cursor] === 0x01) {
        const cb = readUint16(table, cursor + 1);
        cursor += 3 + cb;
    }
    if (cursor + 5 > end || table[cursor] !== 0x02) return result;

    const lcbPcdt = readUint32(table, cursor + 1);
    if (lcbPcdt < 12) return result;
    const plc = cursor + 5;
    if (plc + lcbPcdt > table.length) return result;
    const count = Math.floor((lcbPcdt - 4) / 12);
    if (count <= 0) return result;

    const cpBase = plc;
    const pcdBase = plc + (count + 1) * 4;
    for (let i = 0; i < count; i++) {
        const cpStart = readUint32(table, cpBase + i * 4);
        const cpEnd = readUint32(table, cpBase + (i + 1) * 4);
        if (cpEnd <= cpStart) continue;
        const fcRaw = readUint32(table, pcdBase + i * 8 + 2);
        const compressed = (fcRaw & 0x40000000) !== 0;
        const fc = compressed ? Math.floor((fcRaw & 0x3FFFFFFF) / 2) : (fcRaw & 0x3FFFFFFF);
        result.push({ cpStart, cpEnd, fc, compressed });
    }
    return result;
}

/**
 * 解析字体表（SttbfFfn）
 *
 * 每条 FFN 的长度由首字节 `cbFfnM1` 给出（实际长度 = cbFfnM1 + 1），
 * 固定头部 40 字节，字体名在 +40 处、以 UTF-16LE 的 0x0000 结尾。
 * @param table - 表流
 * @param fc - 字体表偏移
 * @param lcb - 字体表长度
 * @returns 字体名数组（下标即 CHPX 里的 ftc）
 */
function readFonts(table: Uint8Array, fc: number, lcb: number): string[] {
    const fonts: string[] = [];
    if (fc <= 0 || lcb <= 0 || fc + 4 > table.length) return fonts;

    const count = readUint16(table, fc);
    const cbExtra = readUint16(table, fc + 2);
    let cursor = fc + 4;
    const end = Math.min(fc + lcb, table.length);
    for (let i = 0; i < count && cursor < end; i++) {
        const size = table[cursor] + 1;
        let name = '';
        for (let k = 40; k + 1 < size && cursor + k + 1 < end; k += 2) {
            const code = readUint16(table, cursor + k);
            if (code === 0) break;
            name += String.fromCharCode(code);
        }
        fonts.push(name);
        cursor += size + cbExtra;
    }
    return fonts;
}

/**
 * 读取 CHPX 运行段（字符属性）
 *
 * CHPX FKP 页结构：`crun + 1` 个 FC（4 字节）→ `crun` 个字节（每个是该 CHPX
 * 在页内的 2 字节单位偏移，0 表示无属性）→ 页尾 511 处是 `crun`。
 * CHPX 本体 = 1 字节 `cb` + `cb` 字节 grpprl。
 * @param stream - `WordDocument` 流
 * @param table - 表流
 * @param fcPlc - PlcfBteChpx 偏移
 * @param lcbPlc - 长度
 * @returns 按 FC 升序排列的运行段
 */
function readChpxRuns(stream: Uint8Array, table: Uint8Array, fcPlc: number, lcbPlc: number): ChpxRun[] {
    const runs: ChpxRun[] = [];
    const pages = readFkpPages(table, fcPlc, lcbPlc);
    for (let p = 0; p < pages.length; p++) {
        const page = pages[p] * FKP_SIZE;
        if (page + FKP_SIZE > stream.length) continue;
        const crun = stream[page + 511];
        if (crun === 0 || crun > 255) continue;
        const rgbBase = page + (crun + 1) * 4;
        for (let i = 0; i < crun; i++) {
            const fcStart = readUint32(stream, page + i * 4);
            const fcEnd = readUint32(stream, page + (i + 1) * 4);
            const wordOffset = stream[rgbBase + i];
            let grpprl: number[] = [];
            if (wordOffset > 0) {
                const off = page + wordOffset * 2;
                if (off < page + FKP_SIZE) {
                    const cb = stream[off];
                    const chunk = stream.subarray(off + 1, off + 1 + cb);
                    grpprl = Array.from(chunk);
                }
            }
            runs.push({ fcStart, fcEnd, grpprl });
        }
    }
    runs.sort((a, b) => a.fcStart - b.fcStart);
    return runs;
}

/**
 * 读取 PAPX 运行段（段落属性）
 *
 * PAPX FKP 页结构：`crun + 1` 个 FC → `crun` 个 13 字节 BX（首字节是 PAPX 的
 * 2 字节单位偏移）→ 页尾 511 处是 `crun`。
 *
 * PAPX 本体的长度有两种形态（见 MS-DOC `PapxInFkp`）：
 * - `cb != 0`：`grpprlInPapx` 长度 = `2 * cb - 1`，从 `cb` 之后开始；
 * - `cb == 0`：下一字节是 `cb'`，其后再跟 `2 * cb'` 字节。
 * 两种形态的**前 2 字节都是 istd**，istd 之后才是 grpprl。
 * @param stream - `WordDocument` 流
 * @param table - 表流
 * @param fcPlc - PlcfBtePapx 偏移
 * @param lcbPlc - 长度
 * @returns 按 FC 升序排列的运行段
 */
function readPapxRuns(stream: Uint8Array, table: Uint8Array, fcPlc: number, lcbPlc: number): PapxRun[] {
    const runs: PapxRun[] = [];
    const pages = readFkpPages(table, fcPlc, lcbPlc);
    for (let p = 0; p < pages.length; p++) {
        const page = pages[p] * FKP_SIZE;
        if (page + FKP_SIZE > stream.length) continue;
        const crun = stream[page + 511];
        if (crun === 0 || crun > 255) continue;
        const bxBase = page + (crun + 1) * 4;
        for (let i = 0; i < crun; i++) {
            const fcStart = readUint32(stream, page + i * 4);
            const fcEnd = readUint32(stream, page + (i + 1) * 4);
            const bx = bxBase + i * 13;
            const props = readPapxProps(stream, page, stream[bx]);
            runs.push({ fcStart, fcEnd, istd: props.istd, grpprl: props.grpprl });
        }
    }
    runs.sort((a, b) => a.fcStart - b.fcStart);
    return runs;
}

/**
 * 解析单个 PAPX（见 {@link readPapxRuns} 的长度规则）
 * @param stream - `WordDocument` 流
 * @param page - FKP 页起始偏移
 * @param offsetInPage - PAPX 在页内的 2 字节单位偏移（0 = 无 PAPX）
 * @returns istd 与 grpprl
 */
function readPapxProps(stream: Uint8Array, page: number, offsetInPage: number): { istd: number; grpprl: number[] } {
    if (!offsetInPage) return { istd: 0, grpprl: [] };
    const off = page + offsetInPage * 2;
    const pageEnd = page + FKP_SIZE;
    if (off + 1 >= pageEnd) return { istd: 0, grpprl: [] };

    const cb = stream[off];
    if (cb !== 0) {
        const size = 2 * cb - 1;
        const istd = readUint16(stream, off + 1);
        return { istd, grpprl: Array.from(stream.subarray(off + 3, Math.min(off + 1 + size, pageEnd))) };
    }
    const cb2 = stream[off + 1];
    const istd = readUint16(stream, off + 2);
    return { istd, grpprl: Array.from(stream.subarray(off + 4, Math.min(off + 2 + 2 * cb2, pageEnd))) };
}

/**
 * 从 PlcfBte（Chpx / Papx 共用的 PLC 形态）里读 FKP 页号
 *
 * `fcPlc + lcbPlc` 的布局：`n + 1` 个 FC（4 字节）→ `n` 个页号 PN（4 字节），
 * 因此 `n = (lcb - 4) / 8`。
 * @param table - 表流
 * @param fcPlc - PLC 偏移
 * @param lcbPlc - PLC 长度
 * @returns 页号数组
 */
function readFkpPages(table: Uint8Array, fcPlc: number, lcbPlc: number): number[] {
    const pages: number[] = [];
    if (fcPlc <= 0 || lcbPlc < 8 || fcPlc + lcbPlc > table.length) return pages;
    const count = Math.floor((lcbPlc - 4) / 8);
    const pnBase = fcPlc + (count + 1) * 4;
    for (let i = 0; i < count; i++) pages.push(readUint32(table, pnBase + i * 4));
    return pages;
}

/**
 * 把可见字符序列按段落标记拆块，并套上字符 / 段落属性
 * @param chars - 可见字符（含 FC）
 * @param fonts - 字体表
 * @param chpxRuns - CHPX 段
 * @param papxRuns - PAPX 段
 * @returns 块树
 */
function buildBlocks(chars: DocChar[], fonts: string[], chpxRuns: ChpxRun[], papxRuns: PapxRun[]): DocBlock[] {
    const blocks: DocBlock[] = [];
    let current: DocChar[] = [];
    let markFc = -1;

    const flush = (): void => {
        const block = createBlock({ type: 'paragraph' });
        let numbered = false;
        if (current.length > 0) {
            const para = buildParagraph(current, markFc, fonts, chpxRuns, papxRuns);
            block.chars = para.chars;
            if (para.spans.length > 0) block.spans = para.spans;
            if (para.align) block.align = para.align;
            if (para.indent) block.indent = para.indent;
            if (para.firstLineIndent) block.firstLineIndent = para.firstLineIndent;
            numbered = para.numbered;
        } else if (markFc >= 0) {
            // 空段落：Word 的空行在导出 / 往返后必须仍是空行，不能吞掉
            const para = buildParagraph([], markFc, fonts, chpxRuns, papxRuns);
            if (para.align) block.align = para.align;
            numbered = para.numbered;
        }
        // 编号段落：文件里常查不到编号定义（见 applyManualNumbering）
        if (numbered) applyManualNumbering(blocks, block);
        blocks.push(block);
        current = [];
        markFc = -1;
    };

    for (let i = 0; i < chars.length; i++) {
        const item = chars[i];
        const code = item.text.charCodeAt(0);
        if (code === CHAR_PARAGRAPH_MARK || code === CHAR_CELL_MARK || code === CHAR_PAGE_BREAK) {
            markFc = item.fc;
            flush();
            continue;
        }
        if (code === CHAR_LINE_BREAK) {
            current.push({ text: '\n', fc: item.fc });
            continue;
        }
        if (code === CHAR_TAB) {
            current.push({ text: '\t', fc: item.fc });
            continue;
        }
        // 代理对：UTF-16LE 正文里高位与低位分别占一个 FC，合并成一个 Char 才能
        // 在 DOM 中按「一个字符」显示（否则 emoji 会被劈成两个方块）
        if (code >= 0xd800 && code <= 0xdbff && i + 1 < chars.length) {
            const next = chars[i + 1];
            const low = next.text.charCodeAt(0);
            if (low >= 0xdc00 && low <= 0xdfff) {
                current.push({ text: item.text + next.text, fc: item.fc });
                i += 1;
                continue;
            }
        }
        current.push(item);
    }
    // 正文末尾没有段落标记的残段也要保留
    if (current.length > 0) flush();
    return blocks;
}

/**
 * 编号段落（`sprmPIlfo`）在文件里查不到编号定义时，按前文手工编号续号
 *
 * WPS / 转换器生成的 .doc 常只留下 `sprmPIlfo`，`ListTables` / `LFO` 要么缺失、
 * 要么被内嵌对象覆盖，编号文字（「三、」）就完全取不到。这类文档的标题通常是
 * 手工写的「一、」「二、」，所以用最近的上一段手工编号推出下一号 ——
 * 比起直接把编号丢掉，这更接近 Word 里看到的样子。
 * @param previous - 已产出的块（正文顺序）
 * @param block - 当前段落
 */
function applyManualNumbering(previous: DocBlock[], block: DocBlock): void {
    const own = blockText(block);
    // 正文里已经写了编号（用户手打）时不重复添加
    if (/^(?:[一二三四五六七八九十]+|\d+)、/.test(own)) return;
    for (let i = previous.length - 1; i >= 0; i--) {
        const text = blockText(previous[i]);
        const chinese = /^([一二三四五六七八九十]+)、/.exec(text);
        if (chinese) {
            const next = nextChineseNumber(chinese[1]);
            if (next) prefixBlockText(block, `${next}、`);
            return;
        }
        const arabic = /^(\d+)、/.exec(text);
        if (arabic) {
            prefixBlockText(block, `${parseInt(arabic[1], 10) + 1}、`);
            return;
        }
    }
}

/**
 * 中文数字文本 → 下一个中文数字（只处理 1–98）
 * @param text - 如「二」
 * @returns 如「三」；解析失败返回 null
 */
function nextChineseNumber(text: string): string | null {
    const value = parseChineseNumber(text);
    if (value <= 0 || value >= 99) return null;
    return formatChineseNumber(value + 1);
}

/**
 * 中文数字文本 → 整数（只处理 1–99）
 * @param text - 中文数字
 * @returns 整数值；解析失败返回 0
 */
function parseChineseNumber(text: string): number {
    const digits = '零一二三四五六七八九';
    const ten = text.indexOf('十');
    if (ten < 0) return text.length === 1 ? digits.indexOf(text) : 0;
    const high = ten === 0 ? 1 : digits.indexOf(text.charAt(0));
    const low = ten === text.length - 1 ? 0 : digits.indexOf(text.charAt(ten + 1));
    if (high < 0 || low < 0) return 0;
    return high * 10 + low;
}

/**
 * 整数 → 中文数字（1–99）
 * @param value - 整数
 * @returns 中文数字
 */
function formatChineseNumber(value: number): string {
    const digits = '零一二三四五六七八九';
    if (value < 10) return digits.charAt(value);
    const tens = Math.floor(value / 10);
    const ones = value % 10;
    return (tens === 1 ? '' : digits.charAt(tens)) + '十' + (ones ? digits.charAt(ones) : '');
}

/**
 * 构造一个段落：字符样式 + 段落级区间 / 缩进 / 对齐
 * @param items - 段落内字符（含 FC）
 * @param markFc - 段落标记的 FC（PAPX 以它定位；空段落也靠它取对齐）
 * @param fonts - 字体表
 * @param chpxRuns - CHPX 段
 * @param papxRuns - PAPX 段
 * @returns 段落数据
 */
function buildParagraph(
    items: DocChar[],
    markFc: number,
    fonts: string[],
    chpxRuns: ChpxRun[],
    papxRuns: PapxRun[]
): { chars: Char[]; spans: DocSpan[]; align?: DocBlock['align']; indent?: number; firstLineIndent?: number; numbered: boolean } {
    const chars: Char[] = [];
    const spans: DocSpan[] = [];
    const highlights: (string | undefined)[] = [];
    const verts: (DocSpan['vertAlign'] | undefined)[] = [];

    for (let i = 0; i < items.length; i++) {
        const item = items[i];
        const char = Char.of(item.text);
        const charStyle: CharStylePatch = {};
        let highlight: string | undefined;
        let vert: DocSpan['vertAlign'] | undefined;

        const chpx = findRun(chpxRuns, item.fc);
        if (chpx && chpx.grpprl.length > 0) {
            const props = parseSprms(chpx.grpprl);
            applyCharProps(charStyle, props, fonts);
            highlight = highlightOf(props);
            vert = vertAlignOf(props);
        }
        writeCharStyle(char, charStyle);
        chars.push(char);
        highlights.push(highlight);
        verts.push(vert);
    }

    coalesceSpans(spans, highlights, 'highlight');
    coalesceSpans(spans, verts, 'vertAlign');

    const paragraph = findRun(papxRuns, markFc >= 0 ? markFc : 0);
    let align: DocBlock['align'];
    let indent: number | undefined;
    let firstLineIndent: number | undefined;
    let numbered = false;

    if (paragraph) {
        const props = parseSprms(paragraph.grpprl);
        const jc = lastOperand(props, 0x2403);
        if (jc !== undefined) {
            if (jc === 1) align = 'center';
            else if (jc === 2) align = 'right';
            else if (jc === 3) align = 'justify';
            else align = 'left';
        }
        const leftTwips = lastOperand(props, 0x840F) ?? lastOperand(props, 0x845E);
        if (leftTwips !== undefined && leftTwips > 0) indent = Math.round(leftTwips / 420) || 1;

        const firstTwips = lastOperand(props, 0x8411) ?? lastOperand(props, 0x8460);
        if (firstTwips !== undefined && firstTwips > 0) {
            // 首行缩进在模型里是「字符数」：1 字符 = 段落字号(pt) × 20 twips
            const markStyle = markFc >= 0 ? findRun(chpxRuns, markFc) : undefined;
            const markProps = markStyle ? parseSprms(markStyle.grpprl) : new Map<number, number[]>();
            const halfPoints = lastOperand(markProps, 0x4A43);
            const fontPt = halfPoints && halfPoints > 0 ? halfPoints / 2 : 12;
            firstLineIndent = snapIndent(firstTwips / (20 * fontPt));
        }

        // sprmPIlfo：段落的编号定义下标（0 = 不编号）
        const ilfo = lastOperand(props, 0x460B);
        numbered = ilfo !== undefined && ilfo > 0;

        // 段落底纹（Shd80）：整段铺底色。白色的底纹在白纸上不可见，但对
        // 「拷进表格 / 换主题色」这类场景仍是必要信息，因此照样登记。
        const shd = lastOperandBytes(props, 0x442D);
        if (shd && items.length > 0) {
            const color = shd80Color(shd);
            if (color) spans.push({ start: 0, end: chars.length, highlight: color });
        }
    }
    return { chars, spans, align, indent, firstLineIndent, numbered };
}

/**
 * 把逐字符的属性压成连续区间（同值相邻字符合成一个 span，避免 span 爆炸）
 * @param spans - 输出区间
 * @param values - 逐字符属性
 * @param key - 属性名
 */
function coalesceSpans(
    spans: DocSpan[],
    values: (string | DocSpan['vertAlign'] | undefined)[],
    key: 'highlight' | 'vertAlign'
): void {
    let start = -1;
    let value: string | DocSpan['vertAlign'] | undefined;
    const flush = (end: number): void => {
        if (start >= 0 && value !== undefined && end > start) {
            const span: DocSpan = { start, end };
            if (key === 'highlight') span.highlight = value as string;
            else span.vertAlign = value as DocSpan['vertAlign'];
            spans.push(span);
        }
        start = -1;
        value = undefined;
    };
    for (let i = 0; i < values.length; i++) {
        const current = values[i];
        if (start >= 0 && current === value) continue;
        flush(i);
        if (current !== undefined) {
            start = i;
            value = current;
        }
    }
    flush(values.length);
}

/**
 * 把 CHPX 的 sprm 集合写进字符样式补丁
 * @param style - 输出样式
 * @param props - 已解析的 sprm（code → 操作数）
 * @param fonts - 字体表
 */
function applyCharProps(style: CharStylePatch, props: Map<number, number[]>, fonts: string[]): void {
    const halfPoints = lastOperand(props, 0x4A43);
    if (halfPoints && halfPoints > 0) {
        style.fontSize = Math.round((halfPoints / 2) * 4 / 3);
    }
    const ftc = lastOperand(props, 0x4A5E) ?? lastOperand(props, 0x4A51)
        ?? lastOperand(props, 0x4A50) ?? lastOperand(props, 0x4A4F);
    if (ftc !== undefined && ftc >= 0 && ftc < fonts.length && fonts[ftc]) {
        style.fontFamily = fonts[ftc];
    }
    if ((lastOperand(props, 0x0835) ?? 0) !== 0) style.bold = true;
    if ((lastOperand(props, 0x0836) ?? 0) !== 0) style.italic = true;
    if ((lastOperand(props, 0x0837) ?? 0) !== 0) style.strike = true;
    const underline = lastOperand(props, 0x2A3E);
    if (underline !== undefined && underline !== 0) style.underline = true;

    // 颜色：0x6870 是 COLORREF（字节序 R,G,B,0），0x2A42 是调色板索引。
    // 两者并存时以 COLORREF 为准 —— 它才是用户在取色器里选中的精确颜色。
    const colorRef = lastOperandBytes(props, 0x6870);
    if (colorRef && colorRef.length >= 3) {
        const rgb = `#${hex(colorRef[0])}${hex(colorRef[1])}${hex(colorRef[2])}`;
        // Word 的「自动」颜色（COLORREF `00 00 00 FF`）就是纯黑：它等于默认正文色，
        // 写进样式反而会把默认的深灰覆盖成死黑。与 .docx 的表现对齐：纯黑不落样式。
        if (rgb !== '#000000') style.color = rgb;
    } else {
        const ico = lastOperand(props, 0x2A42);
        if (ico !== undefined && ico > 1) style.color = ICO_COLORS[ico] ?? undefined;
    }
}

/**
 * 把 twips 折算出的「首行缩进字符数」吸附到 0.5 字符的刻度
 *
 * Word 里显示「2 字符」的段落在二进制里常是 643 twips（而非整算的 640），
 * 直接保留两位小数会得到 2.01，与 .docx 的 2 产生可见的细微差异。
 * 与 0.5 的整数倍相差不超过 0.05 字符时按刻度取值，其余保留两位小数。
 * @param value - 原始字符数
 * @returns 吸附后的字符数
 */
function snapIndent(value: number): number {
    const snapped = Math.round(value * 2) / 2;
    if (Math.abs(value - snapped) <= 0.05) return snapped;
    return Math.round(value * 100) / 100;
}

/**
 * 取字符高亮底色（`0x2A0C` 调色板 / `0x4866` Shd80，后者优先）
 * @param props - sprm 集合
 * @returns CSS 颜色；无高亮时为 undefined
 */
function highlightOf(props: Map<number, number[]>): string | undefined {
    const shd = lastOperandBytes(props, 0x4866);
    if (shd) {
        const color = shd80Color(shd);
        if (color) return color;
    }
    const ico = lastOperand(props, 0x2A0C);
    if (ico !== undefined && ico > 0 && ico <= 16) return ICO_COLORS[ico] ?? undefined;
    return undefined;
}

/**
 * 取上下标（`0x2A48` Iss：1 上标、2 下标）
 * @param props - sprm 集合
 * @returns 上下标；无设置时为 undefined
 */
function vertAlignOf(props: Map<number, number[]>): DocSpan['vertAlign'] | undefined {
    const iss = lastOperand(props, 0x2A48);
    if (iss === 1) return 'superscript';
    if (iss === 2) return 'subscript';
    return undefined;
}

/**
 * 把 Shd80 位域解成 CSS 颜色
 *
 * Shd80 是 2 字节位域：`icoFore`（bit 0–4）、`icoBack`（bit 5–9）、`ipat`（bit 10–15）。
 * `0xFFFF` 表示「无底纹」；`icoBack = 0` 表示自动（同样视为无底纹）。
 * @param bytes - 2 字节 Shd80
 * @returns 背景色；无底纹时为 undefined
 */
function shd80Color(bytes: number[]): string | undefined {
    if (bytes.length < 2) return undefined;
    const value = bytes[0] | (bytes[1] << 8);
    if (value === 0xFFFF) return undefined;
    const icoBack = (value >> 5) & 0x1F;
    if (icoBack === 0) return undefined;
    return ICO_COLORS[icoBack] ?? undefined;
}

/**
 * 解析 grpprl（sprm 序列）为「code → 操作数字节」
 *
 * 操作数长度由 spra 决定：`0/1 → 1`、`2/4/5 → 2`、`3 → 4`、`7 → 3`、
 * `6 → 首字节长度 + 1`（首字节代表后续字节数）。
 * @param grpprl - 原始字节
 * @returns sprm 集合（同一个 code 出现多次时保留最后一次，符合「后写覆盖」）
 */
function parseSprms(grpprl: number[]): Map<number, number[]> {
    const props = new Map<number, number[]>();
    let i = 0;
    while (i + 2 <= grpprl.length) {
        const code = grpprl[i] | (grpprl[i + 1] << 8);
        const spra = (code >> 13) & 7;
        let length: number;
        if (spra === 0 || spra === 1) length = 1;
        else if (spra === 2 || spra === 4 || spra === 5) length = 2;
        else if (spra === 3) length = 4;
        else if (spra === 7) length = 3;
        else length = 1 + (grpprl[i + 2] ?? 0);
        const operand = grpprl.slice(i + 2, i + 2 + length);
        props.set(code, operand);
        i += 2 + length;
    }
    return props;
}

/**
 * 取某个 FC 命中的运行段（FC 区间为闭开区间）
 * @param runs - 运行段（FC 升序）
 * @param fc - 字符 FC
 * @returns 命中的运行段；越界时为 null
 */
function findRun<T extends { fcStart: number; fcEnd: number }>(runs: T[], fc: number): T | null {
    for (let i = 0; i < runs.length; i++) {
        if (fc >= runs[i].fcStart && fc < runs[i].fcEnd) return runs[i];
        if (runs[i].fcStart > fc) break;
    }
    return null;
}

/**
 * 取某个 sprm 的第一个操作数（2 字节小端）
 * @param props - sprm 集合
 * @param code - sprm 码
 * @returns 数值；不存在时 undefined
 */
function lastOperand(props: Map<number, number[]>, code: number): number | undefined {
    const bytes = props.get(code);
    if (!bytes || bytes.length === 0) return undefined;
    return bytes[0] | ((bytes[1] ?? 0) << 8);
}

/**
 * 取某个 sprm 的原始操作数
 * @param props - sprm 集合
 * @param code - sprm 码
 * @returns 操作数字节；不存在时 undefined
 */
function lastOperandBytes(props: Map<number, number[]>, code: number): number[] | undefined {
    return props.get(code);
}

/**
 * 读小端 16 位整数（越界时返回 0）
 * @param bytes - 字节数据
 * @param offset - 偏移
 * @returns 数值
 */
function readUint16(bytes: Uint8Array, offset: number): number {
    if (offset < 0 || offset + 2 > bytes.length) return 0;
    return bytes[offset] | (bytes[offset + 1] << 8);
}

/**
 * 读小端 32 位整数（越界时返回 0）
 * @param bytes - 字节数据
 * @param offset - 偏移
 * @returns 数值
 */
function readUint32(bytes: Uint8Array, offset: number): number {
    if (offset < 0 || offset + 4 > bytes.length) return 0;
    return (bytes[offset] | (bytes[offset + 1] << 8)
        | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24)) >>> 0;
}

/**
 * 数字 → 两位十六进制
 * @param value - 0–255
 * @returns 两位大写十六进制
 */
function hex(value: number): string {
    const n = Math.max(0, Math.min(255, value));
    return (n < 16 ? '0' : '') + n.toString(16).toUpperCase();
}
