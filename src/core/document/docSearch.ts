/*  # 文档域 · 查找 / 替换引擎 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目目标: 在**文档正文的 DOM**上做查找与替换，会话与结构都像 Word 的「查找和替换」。
    # 本文件职责: 把 `.doc-blocks` 里的每个块摊成「虚拟文本 + 字符偏移 → DOM 点」的映射，
    #             再据此算出命中、选中命中、就地替换。**不做对话框**（那是 docSearchDialog 的事），
    #             也**不碰块树模型**（替换完由 DocView.sync 回写）。
    # 为什么在 DOM 上做而不是在块树模型上做:
    #   1. 输入期间模型是滞后的 —— 用户刚打的字还在 DOM 里、尚未同步进模型，
    #      直接查模型会「刚打的字查不到」；查 DOM 则始终是最新画面。
    #   2. 命中要给用户**看得见**的选区，Range 只能落在 DOM 上；反过来由「块 + 偏移」
    #      算 DOM 点，比由模型算 DOM 点简单得多（模型还要处理 spans / 表格 / 代码块拆行）。
    #   3. 替换后调用 `DocView.sync` 回写，撤销、落库、字符统计都走既有链路，
    #      与「样式命令直接改 DOM 再同步」是同一种做法。
    # 依赖方向: utils/dom + core/constant；**不 import 表格数据层、不 import toolbar**。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { DOC_PAGE_GAP_ATTR } from '../constant';

/** 查找匹配选项（与表格侧同名，语义对齐 Word：`matchEntire` = 全字匹配） */
export interface DocSearchOptions {
    /** 区分大小写（默认 false） */
    matchCase?: boolean;
    /** 全字匹配（默认 false；只影响拉丁字母/数字/下划线构成的词，与 Word 一致） */
    matchEntire?: boolean;
    /** 使用通配符：`?` 单个字符、`*` 任意长度（默认 false） */
    useWildcard?: boolean;
}

/** 文档内位置（块序号 + 块内虚拟字符偏移） */
export interface DocPosition {
    /** 块序号（`0` = 文档第一块） */
    blockIndex: number;
    /** 块内虚拟文本偏移 */
    offset: number;
}

/** 一处命中 */
export interface DocSearchHit {
    /** 块序号（`0` = 文档第一块） */
    blockIndex: number;
    /** 命中所在块（`data-id` 的 `.doc-block`） */
    blockEl: HTMLElement;
    /** 命中起始偏移（块内虚拟文本） */
    start: number;
    /** 命中结束偏移（不含） */
    end: number;
    /** 命中文本 */
    text: string;
}

/**
 * 命中上限
 *
 * 对话框只需要「第几处 / 共几处」和前后导航，不需要把所有命中常驻内存。
 * 一篇超长文档里搜一个常见字可能命中数万次，列表化会拖慢每次查找；
 * 截断时对话框会如实说明（与表格侧的 `SEARCH_HIT_LIMIT` 同一口径）。
 */
export const DOC_SEARCH_HIT_LIMIT = 5000;

/** DOM 点（`Range.setStart/setEnd` 的参数） */
interface DomPoint {
    node: Node;
    offset: number;
}

/** 一个块的虚拟文本与偏移映射 */
interface BlockText {
    blockIndex: number;
    blockEl: HTMLElement;
    text: string;
    /** `points[i]` = 虚拟偏移 `i` 对应的 DOM 点（长度 `text.length + 1`） */
    points: DomPoint[];
    /** 文本节点 → 它在块内的起始偏移（由「DOM 点 → 偏移」快速换算） */
    nodeStarts: Map<Node, number>;
}

/** 一条原始匹配（块内偏移） */
interface RawMatch {
    start: number;
    end: number;
    text: string;
}

/**
 * 收集文档里所有可查找的块
 *
 * 跳过：分页留白（不是内容）、分隔线 / 图片（没有可编辑文本，替换会改坏它们的标签）。
 * 表格块照常收集 —— 表格单元格里的文字也应当能被查到，替换后由 `readTable` 读回。
 * @param blocksEl - `.doc-blocks` 容器
 * @returns 块文本映射列表（按文档顺序）
 */
function collectBlocks(blocksEl: HTMLElement): BlockText[] {
    const blocks: BlockText[] = [];
    const children = blocksEl.children;
    for (let i = 0; i < children.length; i++) {
        const el = children[i] as HTMLElement;
        if (!el || !el.getAttribute) continue;
        if (el.hasAttribute(DOC_PAGE_GAP_ATTR)) continue;
        if (!el.getAttribute('data-id')) continue;
        const type = el.getAttribute('data-type');
        if (type === 'image' || type === 'hr') continue;
        blocks.push(collectBlock(el, blocks.length));
    }
    return blocks;
}

/**
 * 把一个块摊成虚拟文本
 *
 * 文本节点逐字符登记 DOM 点；`<br>`（代码块换行）记作 `\n`，其边界点落在
 * 父元素的子节点序号上 —— 这样跨行的命中也能落到合法的 Range 位置上。
 * @param blockEl - 块元素
 * @param blockIndex - 块序号
 * @returns 块文本映射
 */
function collectBlock(blockEl: HTMLElement, blockIndex: number): BlockText {
    const parts: string[] = [];
    const points: DomPoint[] = [];
    const nodeStarts = new Map<Node, number>();
    let index = 0;

    const walk = (node: Node): void => {
        if (node.nodeType === 3) {
            const textNode = node as Text;
            const value = textNode.nodeValue ?? '';
            if (!value) return;
            nodeStarts.set(node, index);
            for (let i = 0; i < value.length; i++) {
                parts.push(value.charAt(i));
                points[index] = { node, offset: i };
                index += 1;
            }
            // 节点末尾也是合法落点；后续节点的首字符会把它覆盖成等价位置
            points[index] = { node, offset: value.length };
            return;
        }
        if (node.nodeType !== 1) return;
        const element = node as HTMLElement;

        /** 在某个元素之前插一个虚拟换行（边界点落在「父元素 + 子节点序号」上） */
        const appendBreak = (boundary: HTMLElement): void => {
            const parent = boundary.parentNode;
            if (!parent) return;
            const childIndex = Array.prototype.indexOf.call(parent.childNodes, boundary);
            parts.push('\n');
            points[index] = { node: parent, offset: childIndex };
            index += 1;
            points[index] = { node: parent, offset: childIndex + 1 };
        };

        if (element.tagName === 'BR') {
            appendBreak(element);
            return;
        }
        // 表格单元格之间补一个换行：虚拟文本是一串拼接字符，若不隔开，
        // 相邻两格的内容会被当成连在一起（"he" + "llo" 会误命中 "hello"），
        // 替换还可能跨格删掉单元格结构。普通关键字输入不含换行，因此天然隔断。
        if ((element.tagName === 'TD' || element.tagName === 'TH') && index > 0) {
            appendBreak(element);
        }
        // 单元格内的多段（`.doc-cell-line`）同理：跨段命中会让 Range 删掉段容器，
        // `readTable` 只认 `.doc-cell-line`，删掉的那段文字会从模型里消失。
        if (element.classList && element.classList.contains('doc-cell-line') && index > 0) {
            appendBreak(element);
        }
        const children = element.childNodes;
        for (let i = 0; i < children.length; i++) walk(children[i]);
    };

    walk(blockEl);
    const text = parts.join('');
    if (!points[0]) points[0] = { node: blockEl, offset: 0 };
    if (!points[text.length]) points[text.length] = { node: blockEl, offset: blockEl.childNodes.length };
    return { blockIndex, blockEl, text, points, nodeStarts };
}

/** 转义正则元字符（非通配符模式下关键字按字面处理） */
function escapeRegExp(text: string): string {
    return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/** 按匹配选项构造全局正则 */
function buildPattern(keyword: string, options: DocSearchOptions): RegExp {
    let source: string;
    if (options.useWildcard === true) {
        const pieces: string[] = [];
        for (let i = 0; i < keyword.length; i++) {
            const ch = keyword.charAt(i);
            // 与 Word 一致：`?` / `*` 不跨段落与换行（换行在虚拟文本里就是 `\n`）
            if (ch === '?') pieces.push('[^\\n]');
            else if (ch === '*') pieces.push('[^\\n]*');
            else pieces.push(escapeRegExp(ch));
        }
        source = pieces.join('');
    } else {
        source = escapeRegExp(keyword);
    }
    return new RegExp(source, options.matchCase === true ? 'g' : 'gi');
}

/** 是否是「词字符」（全字匹配以它判定词边界；CJK 不算词字符，与 Word 的行为一致） */
function isWordChar(ch: string): boolean {
    return ch !== '' && /[0-9A-Za-z_]/.test(ch);
}

/** 全字匹配的词边界校验 */
function acceptsBoundary(text: string, start: number, end: number, options: DocSearchOptions): boolean {
    if (options.matchEntire !== true) return true;
    const before = start > 0 ? text.charAt(start - 1) : '';
    const after = end < text.length ? text.charAt(end) : '';
    return !isWordChar(before) && !isWordChar(after);
}

/** 在一个块的虚拟文本里找出全部匹配 */
function matchInText(text: string, keyword: string, options: DocSearchOptions): RawMatch[] {
    if (!keyword) return [];
    const pattern = buildPattern(keyword, options);
    const matches: RawMatch[] = [];
    let result: RegExpExecArray | null;
    while ((result = pattern.exec(text)) !== null) {
        const matched = result[0];
        if (matched.length === 0) {
            // 通配符 `*` 可以匹配空串：不跳过就会让 lastIndex 原地踏步、死循环
            pattern.lastIndex = result.index + 1;
            if (pattern.lastIndex > text.length) break;
            continue;
        }
        if (acceptsBoundary(text, result.index, result.index + matched.length, options)) {
            matches.push({ start: result.index, end: result.index + matched.length, text: matched });
        }
    }
    return matches;
}

/**
 * 在文档里找出全部命中（按文档顺序）
 * @param blocksEl - `.doc-blocks` 容器
 * @param keyword - 关键字（空串返回空数组）
 * @param options - 匹配选项
 * @param limit - 命中上限
 * @returns 命中列表
 */
export function findDocHits(
    blocksEl: HTMLElement,
    keyword: string,
    options: DocSearchOptions = {},
    limit: number = DOC_SEARCH_HIT_LIMIT,
): DocSearchHit[] {
    const hits: DocSearchHit[] = [];
    if (!keyword) return hits;
    const blocks = collectBlocks(blocksEl);
    for (let b = 0; b < blocks.length; b++) {
        const block = blocks[b];
        const matches = matchInText(block.text, keyword, options);
        for (let i = 0; i < matches.length; i++) {
            const match = matches[i];
            hits.push({
                blockIndex: block.blockIndex,
                blockEl: block.blockEl,
                start: match.start,
                end: match.end,
                text: match.text,
            });
            if (hits.length >= limit) return hits;
        }
    }
    return hits;
}

/**
 * 把 DOM 点换算成「块序号 + 块内偏移」
 * @param blocks - 块映射
 * @param node - DOM 节点
 * @param offset - 节点内偏移
 * @returns 文档内位置；不在任何块内时为 null
 */
function locatePoint(blocks: BlockText[], node: Node, offset: number): DocPosition | null {
    for (let b = 0; b < blocks.length; b++) {
        const block = blocks[b];
        if (!block.blockEl.contains(node)) continue;
        if (node.nodeType === 3) {
            const base = block.nodeStarts.get(node);
            if (base !== undefined) {
                return { blockIndex: block.blockIndex, offset: Math.min(base + offset, block.text.length) };
            }
        }
        // 元素节点（块本身 / `<br>` 的父元素）：在边界点表里找同一个点
        for (let i = 0; i <= block.text.length; i++) {
            const point = block.points[i];
            if (point && point.node === node && point.offset === offset) {
                return { blockIndex: block.blockIndex, offset: i };
            }
        }
        return null;
    }
    return null;
}

/**
 * 取文档视图内的当前选区
 *
 * 对话框打开后焦点在输入框里，浏览器选区不一定还落在正文上；调用方需要
 * 在「聚焦输入框之前」取一次，或者接受「没有选区就从文档开头查」。
 * @param blocksEl - `.doc-blocks` 容器
 * @returns 选中的文本与文档内起止位置；选区不在正文内时为 null
 */
export function getDocSelection(blocksEl: HTMLElement): { text: string; start: DocPosition; end: DocPosition } | null {
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) return null;
    const range = selection.getRangeAt(0);
    if (!blocksEl.contains(range.startContainer) || !blocksEl.contains(range.endContainer)) return null;
    const blocks = collectBlocks(blocksEl);
    const start = locatePoint(blocks, range.startContainer, range.startOffset);
    const end = locatePoint(blocks, range.endContainer, range.endOffset);
    if (!start || !end) return null;
    return { text: range.toString(), start, end };
}

/**
 * 由块内偏移构造 Range
 * @param block - 块映射
 * @param start - 起始偏移
 * @param end - 结束偏移
 * @returns Range；边界点缺失时为 null
 */
function rangeOf(block: BlockText, start: number, end: number): Range | null {
    const startPoint = block.points[start];
    const endPoint = block.points[end];
    if (!startPoint || !endPoint) return null;
    const range = document.createRange();
    try {
        range.setStart(startPoint.node, startPoint.offset);
        range.setEnd(endPoint.node, endPoint.offset);
    } catch {
        return null;
    }
    return range;
}

/**
 * 在文档里选中一处命中并滚动到它
 * @param hit - 命中
 * @param blocksEl - `.doc-blocks` 容器
 */
export function selectDocHit(hit: DocSearchHit, blocksEl: HTMLElement): void {
    const blocks = collectBlocks(blocksEl);
    let block: BlockText | null = null;
    for (let i = 0; i < blocks.length; i++) {
        if (blocks[i].blockEl === hit.blockEl) {
            block = blocks[i];
            break;
        }
    }
    if (!block) return;
    const range = rangeOf(block, hit.start, hit.end);
    if (!range) return;
    const selection = window.getSelection();
    if (!selection) return;
    selection.removeAllRanges();
    selection.addRange(range);
    // 命中所在块滚到视野中间：Word 找到之后也是把命中行带到眼前
    block.blockEl.scrollIntoView({ block: 'center', inline: 'nearest' });
}

/**
 * 就地替换一个块内区间
 *
 * 命中落在单个文本节点内时直接改 `nodeValue`：保留该节点的父元素样式
 * （加粗 / 链接 / 高亮的区间都不会因为替换而漂移）；跨节点（罕见，
 * 例如命中横跨两个不同样式的 run）才走 Range 删除 + 插入。
 * @param block - 块映射
 * @param start - 起始偏移
 * @param end - 结束偏移
 * @param replacement - 替换文本
 * @returns 是否替换成功
 */
function replaceRange(block: BlockText, start: number, end: number, replacement: string): boolean {
    const startPoint = block.points[start];
    const endPoint = block.points[end];
    if (!startPoint || !endPoint) return false;
    if (startPoint.node.nodeType === 3 && startPoint.node === endPoint.node) {
        const textNode = startPoint.node as Text;
        const value = textNode.nodeValue ?? '';
        textNode.nodeValue = value.slice(0, startPoint.offset) + replacement + value.slice(endPoint.offset);
        return true;
    }
    const range = rangeOf(block, start, end);
    if (!range) return false;
    range.deleteContents();
    if (replacement) range.insertNode(document.createTextNode(replacement));
    return true;
}

/**
 * 替换一处命中
 * @param hit - 命中（须由当前 DOM 算出）
 * @param replacement - 替换文本
 * @param blocksEl - `.doc-blocks` 容器
 * @returns 是否替换成功
 */
export function replaceDocHit(hit: DocSearchHit, replacement: string, blocksEl: HTMLElement): boolean {
    const blocks = collectBlocks(blocksEl);
    for (let i = 0; i < blocks.length; i++) {
        if (blocks[i].blockEl === hit.blockEl) return replaceRange(blocks[i], hit.start, hit.end, replacement);
    }
    return false;
}

/**
 * 替换文档内全部命中
 *
 * 逐块「从后往前」替换：后面的改动不会移动前面命中的偏移，因此一次收集的
 * 映射表足以完成整轮的替换（不必每替换一处就重扫全文）。
 * @param blocksEl - `.doc-blocks` 容器
 * @param keyword - 关键字
 * @param replacement - 替换文本
 * @param options - 匹配选项
 * @returns 实际替换处数
 */
export function replaceAllDocHits(
    blocksEl: HTMLElement,
    keyword: string,
    replacement: string,
    options: DocSearchOptions,
): number {
    if (!keyword) return 0;
    const blocks = collectBlocks(blocksEl);
    let count = 0;
    for (let b = 0; b < blocks.length; b++) {
        const block = blocks[b];
        const matches = matchInText(block.text, keyword, options);
        for (let i = matches.length - 1; i >= 0; i--) {
            if (replaceRange(block, matches[i].start, matches[i].end, replacement)) count += 1;
        }
    }
    return count;
}
