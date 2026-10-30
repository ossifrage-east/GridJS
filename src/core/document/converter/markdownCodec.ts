/*  # 文档域 · Markdown 编解码 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目目标: 让块树与 Markdown 互转，且**往返尽量无损**（样式保留率 ≥ 95%）。
    # 本文件职责: 两个纯函数 —— 块树投影为 Markdown、Markdown 还原为块树；
    #             以及支撑它们的行内解析器与转义器。
    # 依赖方向: 只 import 同目录 blockTree 与 dataArchitecture/char（纯数据）。
    #           **不碰 DOM、不碰存储**。
    # 扩展语法（标准 Markdown 表达不了、但块树需要记住的样式）:
    #   `==高亮==`        字符底色        → DocSpan.highlight
    #   `[文字]{#RRGGBB}` 字体颜色        → Char.fontColor
    #   `<u>文字</u>`     下划线          → Char.underline（HTML 是 Markdown 的合法子集）
    #   `<sup>/<sub>`     上标 / 下标     → DocSpan.vertAlign
    # 这些写法在外部阅读器里的表现：`==` 与 `{#}` 退化为字面文本，`<u>/<sup>` 正常渲染。
    # 因此「导出给外部工具」时应选用 {@link blocksToMarkdown} 的 standardOnly 选项。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { Char } from '../../dataArchitecture/char';
import {
    CharStylePatch,
    DocBlock,
    DocSpan,
    blockText,
    charsFromText,
    createBlock,
    normalizeBlocks,
    spanAt,
    writeCharStyle,
} from '../blockTree';

/** 等宽字体族名（代码段识别用） */
const MONO_FAMILY = 'monospace';

/** 默认正文字号（导出 HTML 式样式时用不到，这里仅作占位说明） */

/**
 * 块树 → Markdown
 * @param blocks - 块树
 * @param standardOnly - 为 true 时**丢弃**标准 Markdown 表达不了的样式（高亮 / 颜色），
 *                       使输出可被任意外部工具正确阅读
 * @returns Markdown 文本
 */
export function blocksToMarkdown(blocks: DocBlock[], standardOnly: boolean = false): string {
    const parts: string[] = [];
    for (let i = 0; i < blocks.length; i++) {
        const block = blocks[i];
        if (!block) continue;
        const text = inlineToMarkdown(block, standardOnly);
        switch (block.type) {
            case 'heading': {
                const level = Math.max(1, Math.min(6, block.level ?? 1));
                parts.push(`${repeat('#', level)} ${text}`);
                break;
            }
            case 'listItem': {
                const depth = (block.level ?? 0) + (block.indent ?? 0);
                const marker = block.listStyle === 'ordered' ? '1. ' : '- ';
                parts.push(repeat('  ', Math.max(0, depth)) + marker + text);
                break;
            }
            case 'quote': {
                parts.push(text.split('\n').map((line) => `> ${line}`).join('\n'));
                break;
            }
            case 'code': {
                const body = Array.isArray(block.children) && block.children.length > 0
                    ? block.children.map((child) => blockText(child)).join('\n')
                    : blockText(block);
                // 围栏长度按内容自适应：正文里出现三个反引号时用四个，避免提前收尾
                const fence = body.indexOf('```') >= 0 ? '````' : '```';
                parts.push(`${fence}\n${body}\n${fence}`);
                break;
            }
            case 'hr':
                parts.push('---');
                break;
            case 'table': {
                const rows = block.rows ?? [];
                const lines: string[] = [];
                for (let r = 0; r < rows.length; r++) {
                    const cells = (rows[r] ?? []).map((cell) =>
                        inlineToMarkdown(cell, standardOnly).replace(/\|/g, '\\|')
                    );
                    lines.push(`| ${cells.join(' | ')} |`);
                    if (r === 0) lines.push(`| ${cells.map(() => '---').join(' | ')} |`);
                }
                parts.push(lines.join('\n'));
                break;
            }
            case 'image': {
                const ref = block.imageRef;
                parts.push(ref ? `![${text || '图片'}](${ref.assetId})` : '');
                break;
            }
            default:
                parts.push(text);
        }
    }
    // 块之间用空行分隔：Markdown 里「一个空行 = 一个新段落」，与块树一一对应
    return parts.join('\n\n') + '\n';
}

/**
 * 重复字符串（避免依赖 String.prototype.repeat 在 ES5 target 下的 polyfill）
 * @param piece - 片段
 * @param count - 次数
 * @returns 拼接结果
 */
function repeat(piece: string, count: number): string {
    let out = '';
    for (let i = 0; i < count; i++) out += piece;
    return out;
}

/**
 * 把一个块的行内内容渲染成 Markdown 片段
 *
 * 按「样式相同且扩展区间相同」把字符切成连续段，每段独立加包裹符号。
 * 这样 `**粗**体**再粗**` 不会被合并成一个更长的粗体段（那会改变原文的语义边界）。
 * @param block - 块
 * @param standardOnly - 是否只输出标准 Markdown
 * @returns Markdown 片段
 */
function inlineToMarkdown(block: DocBlock, standardOnly: boolean): string {
    const chars = block.chars ?? [];
    let out = '';
    let index = 0;
    while (index < chars.length) {
        const head = chars[index];
        const key = styleKey(head, spanAt(block, index));
        let end = index + 1;
        while (end < chars.length && styleKey(chars[end], spanAt(block, end)) === key) end += 1;

        let text = '';
        for (let i = index; i < end; i++) text += String(chars[i].char ?? '');
        out += wrapSegment(text, head, spanAt(block, index), standardOnly);
        index = end;
    }
    return out;
}

/**
 * 计算一个字符的样式签名（用于把相邻同样式字符合并成一段）
 * @param char - 字符
 * @param span - 该位置命中的扩展区间
 * @returns 签名字符串
 */
function styleKey(char: Char, span: DocSpan | null): string {
    return [
        char.fontFamily ?? '',
        char.fontSize ?? '',
        char.fontWeight ? 1 : 0,
        char.fontStyle ? 1 : 0,
        char.underline ? 1 : 0,
        char.strikethrough ? 1 : 0,
        char.fontColor ?? '',
        span?.highlight ?? '',
        span?.link ?? '',
        span?.vertAlign ?? '',
    ].join('|');
}

/**
 * 判断是否等宽字体（等宽 → 代码段）
 * @param char - 字符
 * @returns 是否等宽
 */
function isMono(char: Char): boolean {
    const family = (char.fontFamily ?? '').toLowerCase();
    return family.indexOf('mono') >= 0 || family.indexOf('courier') >= 0 || family.indexOf('consol') >= 0;
}

/**
 * 给一段同样式的文本加上 Markdown 包裹
 *
 * 包裹顺序（由内到外）：代码 → 粗斜 → 删除线 → 下划线 → 上下标 → 高亮 → 链接/颜色。
 * 顺序不能随意调换：`~~**t**~~` 与 `**~~t~~**` 在多数解析器里等价，但前者与
 * 我们自己的解析器实现一致，往返才稳定。
 * @param text - 段文本
 * @param char - 该段首个字符（取样式）
 * @param span - 该段的扩展区间
 * @param standardOnly - 是否只输出标准 Markdown
 * @returns 包裹后的片段
 */
function wrapSegment(
    text: string,
    char: Char,
    span: DocSpan | null,
    standardOnly: boolean
): string {
    if (!text) return '';
    let body = isMono(char) ? '`' + text + '`' : escapeMarkdown(text);

    const bold = !!char.fontWeight;
    const italic = !!char.fontStyle;
    if (bold && italic) body = `***${body}***`;
    else if (bold) body = `**${body}**`;
    else if (italic) body = `*${body}*`;

    if (char.strikethrough) body = `~~${body}~~`;
    if (char.underline) body = `<u>${body}</u>`;
    if (span?.vertAlign === 'superscript') body = `<sup>${body}</sup>`;
    else if (span?.vertAlign === 'subscript') body = `<sub>${body}</sub>`;
    if (!standardOnly && span?.highlight) body = `==${body}==`;

    if (span?.link) body = `[${body}](${span.link})`;
    if (!standardOnly && char.fontColor) body = `[${body}]{#${normalizeHex(char.fontColor)}}`;
    return body;
}

/**
 * 把 CSS 颜色归一化成 6 位十六进制（Markdown 扩展语法只认这种形态）
 * @param color - 任意 CSS 颜色（`#f00` / `#FF0000` / `rgb(255,0,0)`）
 * @returns 6 位十六进制；无法识别时原样返回
 */
function normalizeHex(color: string): string {
    const text = (color ?? '').trim();
    const short = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/i.exec(text);
    if (short) return `${short[1]}${short[1]}${short[2]}${short[2]}${short[3]}${short[3]}`.toUpperCase();
    if (/^#[0-9a-f]{6}$/i.test(text)) return text.slice(1).toUpperCase();
    const rgb = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i.exec(text);
    if (rgb) {
        const hex = (n: string) => {
            const value = Math.max(0, Math.min(255, parseInt(n, 10)));
            return (value < 16 ? '0' : '') + value.toString(16).toUpperCase();
        };
        return hex(rgb[1]) + hex(rgb[2]) + hex(rgb[3]);
    }
    return text.replace(/^#/, '').toUpperCase();
}

/**
 * 转义 Markdown 里有语法含义的字符
 *
 * 刻意**不转义** `=` 与 `~`：它们在正文里极常见（`a=b`、`约 3~5 天`），
 * 而我们的解析器只在成对出现（`==` / `~~`）时才当语法，单个字符不会误判。
 * @param text - 原始文本
 * @returns 转义后的文本
 */
function escapeMarkdown(text: string): string {
    return text.replace(/([\\*_`\[\]<>|])/g, '\\$1');
}

/**
 * Markdown → 块树
 * @param markdown - Markdown 文本
 * @returns 块树（至少一个块）
 */
export function markdownToBlocks(markdown: string): DocBlock[] {
    const lines = (markdown ?? '').replace(/\r\n?/g, '\n').split('\n');
    const blocks: DocBlock[] = [];
    let i = 0;

    while (i < lines.length) {
        const line = lines[i];

        // ① 代码块围栏（允许 ``` 与 ````` 等任意长度 ≥3 的围栏）
        const fence = /^(\s*)(`{3,}|~{3,})(.*)$/.exec(line);
        if (fence) {
            const marker = fence[2];
            const body: string[] = [];
            i += 1;
            while (i < lines.length && lines[i].trim().indexOf(marker) !== 0) {
                body.push(lines[i]);
                i += 1;
            }
            i += 1;   // 跳过收尾围栏
            const block = createBlock({ type: 'code', text: body.join('\n') });
            block.children = body.map((item) => createBlock({ type: 'paragraph', text: item }));
            blocks.push(block);
            continue;
        }

        // ② 空行 = **块之间的分隔符**，不生成空块
        //
        // 标准 Markdown 用「一个空行」表达「这里另起一段」，并没有「空段落」这个概念。
        // 若把空行也建成空块，`块 → Markdown → 块` 会让块数翻倍（每个分隔空行都变成
        // 一个空段落），往返两次文档就满是空行。代价是用户刻意插入的空段落会被吞掉 ——
        // 这是格式本身的表达力上限，不是实现缺陷。
        if (line.trim() === '') {
            i += 1;
            continue;
        }

        // ③ 分隔线
        if (/^\s*(-{3,}|\*{3,}|_{3,})\s*$/.test(line)) {
            blocks.push(createBlock({ type: 'hr' }));
            i += 1;
            continue;
        }

        // ④ 标题
        const heading = /^(#{1,6})\s+(.*)$/.exec(line);
        if (heading) {
            const block = createBlock({ type: 'heading', level: heading[1].length });
            parseInlineInto(block, heading[2]);
            blocks.push(block);
            i += 1;
            continue;
        }

        // ⑤ 引用（连续 `>` 行合并成一个引用块，行内用 \n 连接）
        if (/^\s*>\s?/.test(line)) {
            const body: string[] = [];
            while (i < lines.length && /^\s*>\s?/.test(lines[i])) {
                body.push(lines[i].replace(/^\s*>\s?/, ''));
                i += 1;
            }
            const block = createBlock({ type: 'quote' });
            parseInlineInto(block, body.join('\n'));
            blocks.push(block);
            continue;
        }

        // ⑥ 列表项（支持 2 空格一级的缩进）
        const list = /^(\s*)([-*+]|\d{1,9}[.)])\s+(.*)$/.exec(line);
        if (list) {
            const depth = Math.floor(list[1].replace(/\t/g, '  ').length / 2);
            const ordered = /\d/.test(list[2]);
            const block = createBlock({
                type: 'listItem',
                level: Math.min(4, depth),
                listStyle: ordered ? 'ordered' : 'bullet',
            });
            parseInlineInto(block, list[3]);
            blocks.push(block);
            i += 1;
            continue;
        }

        // ⑦ 表格（连续以 `|` 开头的行，第二行若是分隔行则跳过）
        if (/^\s*\|.*\|\s*$/.test(line)) {
            const rows: DocBlock[][] = [];
            let sawSeparator = false;
            while (i < lines.length && /^\s*\|.*\|\s*$/.test(lines[i])) {
                const raw = lines[i].trim().replace(/^\|/, '').replace(/\|$/, '');
                const cells = splitTableRow(raw);
                if (!sawSeparator && cells.every((cell) => /^:?-{2,}:?$/.test(cell.trim()))) {
                    sawSeparator = true;
                    i += 1;
                    continue;
                }
                rows.push(cells.map((cell) => {
                    const paragraph = createBlock({ type: 'paragraph' });
                    parseInlineInto(paragraph, cell.trim());
                    return paragraph;
                }));
                i += 1;
            }
            const table = createBlock({ type: 'table' });
            table.rows = rows;
            blocks.push(table);
            continue;
        }

        // ⑧ 普通段落
        const paragraph = createBlock({ type: 'paragraph' });
        parseInlineInto(paragraph, line);
        blocks.push(paragraph);
        i += 1;
    }

    if (blocks.length === 0) blocks.push(createBlock({ type: 'paragraph' }));
    return normalizeBlocks(blocks);
}

/**
 * 拆分表格行（照顾 `\|` 转义）
 * @param raw - 去掉首尾竖线后的行内容
 * @returns 单元格文本数组
 */
function splitTableRow(raw: string): string[] {
    const cells: string[] = [];
    let buffer = '';
    for (let i = 0; i < raw.length; i++) {
        const ch = raw[i];
        if (ch === '\\' && raw[i + 1] === '|') {
            buffer += '|';
            i += 1;
            continue;
        }
        if (ch === '|') {
            cells.push(buffer);
            buffer = '';
            continue;
        }
        buffer += ch;
    }
    cells.push(buffer);
    return cells;
}

/** 行内解析上下文 */
interface InlineContext {
    /** 字符样式补丁 */
    style: CharStylePatch;
    /** 链接地址（来自 `[text](url)`） */
    link?: string;
    /** 高亮底色（来自 `==text==`） */
    highlight?: string;
    /** 上下标（来自 `<sup>/<sub>`） */
    vertAlign?: 'superscript' | 'subscript';
}

/**
 * 把一段行内 Markdown 解析进块里
 * @param block - 目标块（`chars` / `spans` 会被追加）
 * @param text - 行内文本
 */
function parseInlineInto(block: DocBlock, text: string): void {
    const spans: DocSpan[] = block.spans ?? [];
    parseInline(text, { style: {} }, block.chars, spans);
    if (spans.length > 0) block.spans = spans;
}

/**
 * 行内 Markdown 解析器（递归下降）
 *
 * 只认识 8 种行内语法，全部成对出现才生效；不匹配时按字面文本处理。
 * 递归处理是为了让 `**粗 *斜* 粗**` 这类嵌套正确落到「粗 + 斜」的字符样式上。
 * @param text - 待解析文本
 * @param ctx - 当前上下文（样式 / 链接 / 高亮 / 上下标）
 * @param chars - 输出字符数组
 * @param spans - 输出区间数组
 */
function parseInline(text: string, ctx: InlineContext, chars: Char[], spans: DocSpan[]): void {
    let i = 0;
    let plain = '';
    const flush = (): void => {
        if (!plain) return;
        appendText(plain, ctx, chars, spans);
        plain = '';
    };

    while (i < text.length) {
        const ch = text[i];

        // 反斜杠转义
        if (ch === '\\' && i + 1 < text.length) {
            const next = text[i + 1];
            if ('\\*_`[]<>|~#-=!'.indexOf(next) >= 0) {
                plain += next;
                i += 2;
                continue;
            }
        }

        // 行内代码：`code`
        if (ch === '`') {
            const close = text.indexOf('`', i + 1);
            if (close > i) {
                flush();
                appendText(text.slice(i + 1, close), {
                    style: { ...ctx.style, fontFamily: MONO_FAMILY },
                    link: ctx.link,
                    highlight: ctx.highlight,
                    vertAlign: ctx.vertAlign,
                }, chars, spans);
                i = close + 1;
                continue;
            }
        }

        // 粗斜体：*** / ** / *
        if (ch === '*' || ch === '_') {
            const triple = text.substr(i, 3);
            if (triple === '***' || triple === '___') {
                const close = text.indexOf(triple, i + 3);
                if (close > i + 3) {
                    flush();
                    parseInline(text.slice(i + 3, close), {
                        ...ctx,
                        style: { ...ctx.style, bold: true, italic: true },
                    }, chars, spans);
                    i = close + 3;
                    continue;
                }
            }
            const double = text.substr(i, 2);
            if (double === '**' || double === '__') {
                const close = text.indexOf(double, i + 2);
                if (close > i + 2) {
                    flush();
                    parseInline(text.slice(i + 2, close), {
                        ...ctx,
                        style: { ...ctx.style, bold: true },
                    }, chars, spans);
                    i = close + 2;
                    continue;
                }
            }
            const single = text[i];
            const close = text.indexOf(single, i + 1);
            if (close > i + 1) {
                flush();
                parseInline(text.slice(i + 1, close), {
                    ...ctx,
                    style: { ...ctx.style, italic: true },
                }, chars, spans);
                i = close + 1;
                continue;
            }
        }

        // 删除线：~~text~~
        if (ch === '~' && text.substr(i, 2) === '~~') {
            const close = text.indexOf('~~', i + 2);
            if (close > i + 2) {
                flush();
                parseInline(text.slice(i + 2, close), {
                    ...ctx,
                    style: { ...ctx.style, strike: true },
                }, chars, spans);
                i = close + 2;
                continue;
            }
        }

        // 高亮：==text==（扩展语法）
        if (ch === '=' && text.substr(i, 2) === '==') {
            const close = text.indexOf('==', i + 2);
            if (close > i + 2) {
                flush();
                parseInline(text.slice(i + 2, close), { ...ctx, highlight: '#FFFF00' }, chars, spans);
                i = close + 2;
                continue;
            }
        }

        // 链接 / 颜色：[text](url) 或 [text]{#RRGGBB}
        if (ch === '[') {
            const parsed = parseBracket(text, i);
            if (parsed) {
                flush();
                const inner = { ...ctx };
                if (parsed.url) inner.link = parsed.url;
                if (parsed.color) inner.style = { ...ctx.style, color: parsed.color };
                parseInline(parsed.label, inner, chars, spans);
                i = parsed.end;
                continue;
            }
        }

        // 下划线 / 上下标：<u> / <sup> / <sub>
        if (ch === '<') {
            const tag = /^<(u|sup|sub)>/i.exec(text.slice(i));
            if (tag) {
                const name = tag[1].toLowerCase();
                const closeTag = `</${name}>`;
                const close = text.toLowerCase().indexOf(closeTag, i + tag[0].length);
                if (close > i) {
                    flush();
                    const inner: InlineContext = { ...ctx };
                    if (name === 'u') inner.style = { ...ctx.style, underline: true };
                    else inner.vertAlign = name === 'sup' ? 'superscript' : 'subscript';
                    parseInline(text.slice(i + tag[0].length, close), inner, chars, spans);
                    i = close + closeTag.length;
                    continue;
                }
            }
        }

        plain += ch;
        i += 1;
    }
    flush();
}

/**
 * 解析 `[文字](链接)` / `[文字]{#颜色}` 形式
 * @param text - 全文
 * @param start - `[` 的下标
 * @returns 解析结果；不是该形式时为 null
 */
function parseBracket(
    text: string,
    start: number
): { label: string; url?: string; color?: string; end: number } | null {
    let depth = 0;
    let close = -1;
    for (let i = start; i < text.length; i++) {
        const ch = text[i];
        if (ch === '\\') {
            i += 1;
            continue;
        }
        if (ch === '[') depth += 1;
        else if (ch === ']') {
            depth -= 1;
            if (depth === 0) {
                close = i;
                break;
            }
        }
    }
    if (close < 0) return null;
    const label = text.slice(start + 1, close);
    const after = text.slice(close + 1);

    if (after.charAt(0) === '(') {
        const end = after.indexOf(')');
        if (end < 0) return null;
        const url = after.slice(1, end).trim();
        const tail = after.slice(end + 1);
        const colorMatch = /^\{#([0-9a-fA-F]{3,8})\}/.exec(tail);
        if (colorMatch) {
            return {
                label,
                url,
                color: colorMatch[1].length === 3
                    ? `#${colorMatch[1][0]}${colorMatch[1][0]}${colorMatch[1][1]}${colorMatch[1][1]}${colorMatch[1][2]}${colorMatch[1][2]}`
                    : `#${colorMatch[1]}`,
                end: close + 1 + end + 1 + colorMatch[0].length,
            };
        }
        return { label, url, end: close + 1 + end + 1 };
    }

    const colorMatch = /^\{#([0-9a-fA-F]{3,8})\}/.exec(after);
    if (colorMatch) {
        return {
            label,
            color: colorMatch[1].length === 3
                ? `#${colorMatch[1][0]}${colorMatch[1][0]}${colorMatch[1][1]}${colorMatch[1][1]}${colorMatch[1][2]}${colorMatch[1][2]}`
                : `#${colorMatch[1]}`,
            end: close + 1 + colorMatch[0].length,
        };
    }
    return null;
}

/**
 * 把纯文本按当前上下文追加为 `Char[]`，并在需要时登记扩展区间
 * @param text - 文本
 * @param ctx - 上下文
 * @param chars - 输出字符数组
 * @param spans - 输出区间数组
 */
function appendText(text: string, ctx: InlineContext, chars: Char[], spans: DocSpan[]): void {
    if (!text) return;
    const start = chars.length;
    const created = charsFromText(text);
    for (let i = 0; i < created.length; i++) {
        writeCharStyle(created[i], ctx.style);
        chars.push(created[i]);
    }
    const end = chars.length;
    if (end <= start) return;
    if (ctx.highlight || ctx.link || ctx.vertAlign) {
        const span: DocSpan = { start, end };
        if (ctx.highlight) span.highlight = ctx.highlight;
        if (ctx.link) span.link = ctx.link;
        if (ctx.vertAlign) span.vertAlign = ctx.vertAlign;
        spans.push(span);
    }
}
