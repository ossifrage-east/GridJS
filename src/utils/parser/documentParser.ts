/*  # 文档类文件纯文本提取组件 TS 文件（DOCX / DOC / PPTX / PPT）
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { ZipParser } from "./zip";
import { CompoundFile } from "./cfb";

/** 幻灯片分隔行的前缀：用于在表格中辨识页与页边界，标题取值时据此跳过 */
const SLIDE_SEPARATOR_PREFIX = '—— 第 ';

/**
 * 文档解析结果
 * @interface DocumentParseResult
 */
export interface DocumentParseResult {
    /** 按段落/换行拆分的纯文本行 */
    lines: string[];
    /** 文档标题（取首个非空段落，缺省为空串） */
    title: string;
}

/**
 * 通用文档组件
 *
 * 从 Office 文档中提取纯文本，供「打开文档」时铺排到表格 A 列：
 * - **docx**：ZIP 包，取 `word/document.xml`，按 `<w:p>` 段落取 `<w:t>` 文本
 *   （表格单元格 `<w:tc>` 内段落天然按行排列，制表符以 `\t` 保留列分隔）；
 * - **pptx**：ZIP 包，按 `ppt/slides/slideN.xml` 顺序遍历每张幻灯片，
 *   取 `<a:p>` 段落中的 `<a:t>` 文本，幻灯片之间以空行分隔；
 * - **doc**：OLE2 复合文档，从 `WordDocument` 流按可见字符扫描提取文本；
 * - **ppt**：OLE2 复合文档，从 `PowerPoint Document` 流扫描文本记录提取文本。
 *
 * doc / ppt 为二进制格式，无法像 docx / pptx 那样精确还原结构，
 * 此处采用「可打印字符扫描」尽力提取可读文本，段落划分按换行符还原。
 *
 * @module DocumentParser
 * @example
 * const parser = new DocumentParser();
 * const result = await parser.parse(file);
 * console.log(result.title, result.lines.length);
 */
export class DocumentParser {

    /**
     * 一站式解析：按扩展名分派到对应的提取策略
     *
     * @param {File} file - 文档文件（.docx / .doc / .pptx / .ppt）
     * @returns {Promise<DocumentParseResult>} 提取出的纯文本
     * @throws {Error} 文件格式不支持或内容损坏
     */
    public async parse(file: File): Promise<DocumentParseResult> {
        const extension = file.name.toLowerCase().split('.').pop() || '';
        const buffer = new Uint8Array(await file.arrayBuffer());

        switch (extension) {
            case 'docx':
                return await this.parseDocx(buffer);
            case 'pptx':
                return await this.parsePptx(buffer);
            case 'doc':
                return this.parseDoc(buffer);
            case 'ppt':
                return this.parsePpt(buffer);
            default:
                throw new Error(`不支持的文档格式: .${extension}，仅支持 .docx、.doc、.pptx 和 .ppt`);
        }
    }

    // ════════════════════════════════════════════════════════════
    //  OOXML 路径（docx / pptx）：ZIP + XML
    // ════════════════════════════════════════════════════════════

    /**
     * 解析 docx：读取 word/document.xml，按段落提取文本
     *
     * @param {Uint8Array} buffer - 文件字节数据
     * @returns {Promise<DocumentParseResult>} 提取结果
     * @private
     */
    private async parseDocx(buffer: Uint8Array): Promise<DocumentParseResult> {
        const xml = await this.readZipEntry(buffer, 'word/document.xml');
        if (!xml) {
            throw new Error('docx 文件中未找到 word/document.xml，文件可能已损坏');
        }

        const doc = this.parseXml(xml);
        if (!doc) {
            throw new Error('docx 正文 XML 解析失败，文件可能已损坏');
        }

        // 段落级遍历：<w:p> 内所有 <w:t> 文本按顺序拼接，<w:tab/> 还原为制表符
        const lines: string[] = [];
        const paragraphs = doc.getElementsByTagName('w:p');
        for (let i = 0; i < paragraphs.length; i++) {
            lines.push(this.extractParagraphText(paragraphs[i]));
        }

        return this.buildResult(lines);
    }

    /**
     * 解析 pptx：按幻灯片顺序遍历，提取每页的文本框内容
     *
     * @param {Uint8Array} buffer - 文件字节数据
     * @returns {Promise<DocumentParseResult>} 提取结果
     * @private
     */
    private async parsePptx(buffer: Uint8Array): Promise<DocumentParseResult> {
        const zip = new ZipParser(buffer);
        const archive = zip.parseCentralDirectory();

        // 幻灯片路径形如 ppt/slides/slide1.xml，按数字后缀升序（保证与放映顺序一致）
        const slidePaths = Array.from(archive.entries.keys())
            .filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))
            .sort((a, b) => this.slideIndex(a) - this.slideIndex(b));

        if (slidePaths.length === 0) {
            throw new Error('pptx 文件中未找到任何幻灯片，文件可能已损坏');
        }

        const lines: string[] = [];
        for (const path of slidePaths) {
            const bytes = await zip.readEntry(archive, path);
            if (!bytes) continue;
            const doc = this.parseXml(this.decodeUtf8(bytes));
            if (!doc) continue;

            // 幻灯片分隔标题，便于在表格中辨识页与页的边界
            lines.push(`${SLIDE_SEPARATOR_PREFIX}${this.slideIndex(path)} 页 ——`);
            const paragraphs = doc.getElementsByTagName('a:p');
            for (let i = 0; i < paragraphs.length; i++) {
                const text = this.extractDrawingText(paragraphs[i]);
                if (text) lines.push(text);
            }
            lines.push('');
        }

        return this.buildResult(lines);
    }

    // ════════════════════════════════════════════════════════════
    //  二进制路径（doc / ppt）：OLE2 + 可打印字符扫描
    // ════════════════════════════════════════════════════════════

    /**
     * 解析 doc：从 OLE2 的 WordDocument 流扫描文本
     *
     * @param {Uint8Array} buffer - 文件字节数据
     * @returns {DocumentParseResult} 提取结果
     * @private
     */
    private parseDoc(buffer: Uint8Array): DocumentParseResult {
        const cfb = new CompoundFile(buffer);
        const stream = cfb.readStream('WordDocument');
        if (!stream) {
            throw new Error('doc 文件中未找到 WordDocument 数据流，文件可能已损坏');
        }
        // Word 正文以 CP1252（类 Latin-1）单字节存储为主，整体按 Latin-1 解码后扫描可打印字符
        return this.buildResult(this.scanPrintableText(this.decodeLatin1(stream)));
    }

    /**
     * 解析 ppt：从 OLE2 的 PowerPoint Document 流扫描文本
     *
     * @param {Uint8Array} buffer - 文件字节数据
     * @returns {DocumentParseResult} 提取结果
     * @private
     */
    private parsePpt(buffer: Uint8Array): DocumentParseResult {
        const cfb = new CompoundFile(buffer);
        const stream = cfb.readStream('PowerPoint Document');
        if (!stream) {
            throw new Error('ppt 文件中未找到 PowerPoint Document 数据流，文件可能已损坏');
        }
        // PowerPoint 97-2003 的文本以 UTF-16LE 为主，按 UTF-16LE 解码后扫描可打印字符
        return this.buildResult(this.scanPrintableText(this.decodeUtf16Le(stream)));
    }

    // ════════════════════════════════════════════════════════════
    //  内部辅助
    // ════════════════════════════════════════════════════════════

    /**
     * 从 ZIP 包中读取指定条目并解码为字符串
     * @param {Uint8Array} buffer - 文件字节数据
     * @param {string} path - 条目路径
     * @returns {Promise<string | null>} 条目文本；条目不存在时为 null
     * @private
     */
    private async readZipEntry(buffer: Uint8Array, path: string): Promise<string | null> {
        const zip = new ZipParser(buffer);
        const archive = zip.parseCentralDirectory();
        const bytes = await zip.readEntry(archive, path);
        return bytes ? this.decodeUtf8(bytes) : null;
    }

    /**
     * 解析 XML 字符串为 DOM（解析失败返回 null）
     * @param {string} text - XML 文本
     * @returns {Document | null} 解析结果
     * @private
     */
    private parseXml(text: string): Document | null {
        const doc = new DOMParser().parseFromString(text, 'application/xml');
        return doc.getElementsByTagName('parsererror').length > 0 ? null : doc;
    }

    /**
     * 提取 Word 段落文本：拼接 <w:t>，并将 <w:tab/> 还原为制表符
     *
     * 遍历全部后代节点（getElementsByTagName('*') 按文档顺序返回），
     * 从而保证 <w:t> 与 <w:tab/> 的相对位置被原样保留，
     * 表格中的制表符才能正确充当列分隔
     * @param {Element} paragraph - <w:p> 元素
     * @returns {string} 段落文本
     * @private
     */
    private extractParagraphText(paragraph: Element): string {
        let text = '';
        const descendants = paragraph.getElementsByTagName('*');
        for (let i = 0; i < descendants.length; i++) {
            const tag = descendants[i].tagName;
            if (tag === 'w:t') {
                text += descendants[i].textContent ?? '';
            } else if (tag === 'w:tab') {
                text += '\t';
            }
        }
        return text.replace(/\s+$/, '');
    }

    /**
     * 提取 DrawingML 段落文本：拼接 <a:t>
     * @param {Element} paragraph - <a:p> 元素
     * @returns {string} 段落文本
     * @private
     */
    private extractDrawingText(paragraph: Element): string {
        let text = '';
        const nodes = paragraph.getElementsByTagName('a:t');
        for (let i = 0; i < nodes.length; i++) {
            text += nodes[i].textContent ?? '';
        }
        return text.trim();
    }

    /**
     * 从幻灯片路径中取出序号（slide12.xml → 12）
     * @param {string} path - 幻灯片条目路径
     * @returns {number} 幻灯片序号
     * @private
     */
    private slideIndex(path: string): number {
        const matched = /slide(\d+)\.xml$/.exec(path);
        return matched ? Number(matched[1]) : 0;
    }

    /**
     * UTF-8 解码
     * @param {Uint8Array} bytes - 字节数据
     * @returns {string} 解码文本
     * @private
     */
    private decodeUtf8(bytes: Uint8Array): string {
        if (typeof TextDecoder !== 'undefined') {
            return new TextDecoder('utf-8').decode(bytes);
        }
        let text = '';
        for (let i = 0; i < bytes.length; i++) text += String.fromCharCode(bytes[i]);
        return decodeURIComponent(escape(text));
    }

    /**
     * Latin-1（CP1252 近似）解码：二进制文档流按字节直映字符
     * @param {Uint8Array} bytes - 字节数据
     * @returns {string} 解码文本
     * @private
     */
    private decodeLatin1(bytes: Uint8Array): string {
        let text = '';
        // 分块拼接，避免超长流触发字符串参数上限
        const CHUNK = 8192;
        for (let i = 0; i < bytes.length; i += CHUNK) {
            text += String.fromCharCode.apply(null, Array.from(bytes.subarray(i, i + CHUNK)));
        }
        return text;
    }

    /**
     * UTF-16LE 解码：忽略奇数长度尾部的不完整码元
     * @param {Uint8Array} bytes - 字节数据
     * @returns {string} 解码文本
     * @private
     */
    private decodeUtf16Le(bytes: Uint8Array): string {
        const usable = bytes.length - (bytes.length % 2);
        const units = new Uint16Array(usable / 2);
        for (let i = 0; i < units.length; i++) {
            units[i] = bytes[i * 2] | (bytes[i * 2 + 1] << 8);
        }
        let text = '';
        const CHUNK = 8192;
        for (let i = 0; i < units.length; i += CHUNK) {
            text += String.fromCharCode.apply(null, Array.from(units.subarray(i, i + CHUNK)));
        }
        return text;
    }

    /**
     * 扫描可打印文本：连续的可打印字符聚为一段，遇到换行/控制字符断开
     *
     * 二进制文档流夹杂大量结构字段，直接整体解码会得到乱码；
     * 此处仅保留可打印字符（ASCII 可见字符、常见标点与中日韩文字），
     * 并把长度不足 2 的碎片丢弃以过滤结构噪声。
     * @param {string} text - 全文
     * @returns {string[]} 文本行
     * @private
     */
    private scanPrintableText(text: string): string[] {
        const lines: string[] = [];
        let current = '';

        for (const char of text) {
            const code = char.codePointAt(0) ?? 0;
            const isBreak = code === 0x0D || code === 0x0A || code === 0x00;
            const isPrintable = this.isPrintableChar(code);

            if (isPrintable) {
                current += char;
                continue;
            }
            // 换行符：仅在已有内容时收段；控制字符则仅在内容足够长时收段
            if (current.length >= (isBreak ? 1 : 2)) {
                lines.push(current.trim());
            }
            current = '';
        }
        if (current.length >= 2) lines.push(current.trim());

        return lines;
    }

    /**
     * 判断码点是否为可打印字符
     * @param {number} code - 字符码点
     * @returns {boolean} 是否可打印
     * @private
     */
    private isPrintableChar(code: number): boolean {
        // ASCII 可见字符（空格 ~）
        if (code >= 0x20 && code <= 0x7E) return true;
        // Latin-1 补充区可打印字符
        if (code >= 0xA0 && code <= 0xFF) return true;
        // 常用标点与符号
        if (code >= 0x2000 && code <= 0x206F) return true;
        // CJK 统一表意文字（含扩展 A）、CJK 标点、全角字符
        if (code >= 0x3000 && code <= 0x9FFF) return true;
        // CJK 扩展 B（ surrogate pair 已在 for...of 中按完整码点迭代）
        if (code >= 0x20000 && code <= 0x2A6DF) return true;
        // 常用符号、货币、箭头
        if (code >= 0x20A0 && code <= 0x27BF) return true;
        // 全角字母数字与符号
        if (code >= 0xFF01 && code <= 0xFF5E) return true;
        return false;
    }

    /**
     * 组装解析结果：过滤空行噪声，取首个非空且非分隔行的段落作为标题
     * @param {string[]} lines - 原始文本行
     * @returns {DocumentParseResult} 解析结果
     * @private
     */
    private buildResult(lines: string[]): DocumentParseResult {
        const cleaned = lines.filter((line) => line.trim().length > 0);
        // 幻灯片分隔行是本组件插入的结构标记，不应被当作文档标题
        const title = cleaned.find((line) => !line.startsWith(SLIDE_SEPARATOR_PREFIX))?.trim() ?? '';
        return { lines: cleaned, title };
    }
}