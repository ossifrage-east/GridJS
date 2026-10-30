/*  # 文档域 · 块树模型 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目描述: 基于 HTML5 的办公套件；表格走 Canvas，文档走 DOM。
    # 项目目标: 用一棵「块树」作为文档的唯一真源，使 docx / 纯文本 / Markdown 三种格式
    #           都只是同一棵树的不同投影，任何方向转换都不产生信息分叉。
    # 本文件职责: 定义文档的真源数据结构（DocBlock / DocSpan / DocSnapshot），
    #             并提供**只操作块树、不碰 DOM、不碰存储**的纯函数：
    #             取文本、改字符样式、查字符样式、克隆、原型重建、规范化。
    # 依赖方向: 只 import dataArchitecture/char（纯数据类）与 core/constant（叶子常量）。
    #           **禁止 import DOM / 存储 / 工具栏**——本文件必须能在 Worker 与测试桩里跑。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { Char } from '../dataArchitecture/char';
import { DOC_SNAPSHOT_VERSION } from '../constant';

/**
 * 块类型
 *
 * 只保留「排版语义」上有区别的类型：具体的字号/颜色一律落在 {@link DocBlock.chars}
 * 的字符样式里，而不是再拆出 `h1` `h2` 这类「带样式的类型」——否则同一个「大号加粗」
 * 既可能是 `heading level=1` 也可能是 `paragraph + fontSize`，往返转换必然分叉。
 */
export type BlockType =
    | 'paragraph'
    | 'heading'
    | 'listItem'
    | 'quote'
    | 'code'
    | 'table'
    | 'image'
    | 'hr';

/**
 * 块级水平对齐（与契约层 `TextAlignValue` 结构等价）
 *
 * `distribute` 是 Word 的「分散对齐」：最后一行也拉满行宽。
 * CSS 里没有单一的 `text-align` 值能做到，必须 `justify` + `text-align-last: justify` 两条一起写，
 * 因此读写两端都要特殊处理（见 `docView` 的 `applyBlockAlign` / `readBlockAlign`）。
 */
export type DocAlign = 'left' | 'center' | 'right' | 'justify' | 'distribute';

/** 列表风格 */
export type DocListStyle = 'bullet' | 'ordered' | 'none';

/**
 * 区间型扩展属性
 *
 * ## 为什么不把这些字段塞进 `Char`
 *
 * `Char` 的 `toJSON` 形状被表格存档 v4 与快照比对（`_sameCellSnapshot`）依赖，
 * 给它加字段等于让存量快照全部重算。因此文档专属的「链接 / 高亮 / 上下标」
 * 一律以**区间**形式挂在块上：它们是稀疏的（一份文档里高亮通常只有几段），
 * 区间比「逐字符挂字段」更省内存，也更贴近 Markdown / docx 的表达方式。
 *
 * `start` / `end` 是**码点下标**（与 {@link DocBlock.chars} 的数组下标一致，闭开区间）。
 */
export interface DocSpan {
    /** 起始码点下标（含） */
    start: number;
    /** 结束码点下标（不含） */
    end: number;
    /** 超链接地址 */
    link?: string;
    /** 字符高亮底色（CSS 颜色） */
    highlight?: string;
    /** 上标 / 下标 */
    vertAlign?: 'superscript' | 'subscript';
}

/**
 * 字符样式补丁（工具栏命令的载荷形状）
 *
 * 刻意不叫 `FontStyle`：那个名字已被 `cellContent.ts` 占用（14 个字段，含对齐、合并）。
 * 这里只保留**字符级** 8 项，与 `Char` 的 7 个样式字段 + 高亮一一对应。
 * `undefined` 表示「这条命令不管这个字段」，实现方据此做稀疏赋值。
 */
export interface CharStylePatch {
    /** 字体族 */
    fontFamily?: string;
    /** 字号（CSS 像素） */
    fontSize?: number;
    /** 加粗 */
    bold?: boolean;
    /** 斜体 */
    italic?: boolean;
    /** 下划线 */
    underline?: boolean;
    /** 删除线 */
    strike?: boolean;
    /** 字体颜色 */
    color?: string;
    /**
     * 字符高亮底色
     *
     * 落在 {@link DocSpan.highlight} 而非 `Char`，见 {@link DocSpan} 的说明。
     */
    background?: string;
}

/**
 * 文档块
 */
export interface DocBlock {
    /** 稳定标识：DOM 节点通过 `data-block-id` 与它对应 */
    id: string;
    /** 块类型 */
    type: BlockType;
    /**
     * 行内内容
     *
     * **直接复用表格的 `Char`**：每个码点一个实例，样式字段按需赋值（稀疏）。
     * 不另造 `Run` 类——`Char` 已满足全部需求，且它的 `fromRaw` 稀疏复制正好
     * 用来在 IndexedDB 读回时重建原型。
     */
    chars: Char[];
    /** 区间型扩展属性（链接 / 高亮 / 上下标） */
    spans?: DocSpan[];
    /** 块级水平对齐 */
    align?: DocAlign;
    /** `heading` 的级别 1–6；`listItem` 的缩进层级 0–4 */
    level?: number;
    /** `listItem` 的列表风格 */
    listStyle?: DocListStyle;
    /** `listItem` / `quote` 的额外缩进级数 */
    indent?: number;
    /** 容器块的子块（本版用于表格单元格内的段落） */
    children?: DocBlock[];
    /** `table` 的行列：每行若干单元格块 */
    rows?: DocBlock[][];
    /** `image` 的图片引用（图片二进制另存 `images` 仓，此处只留指针与尺寸） */
    imageRef?: { assetId: string; width: number; height: number };
}

/**
 * 页面设置（文档级）
 *
 * 只保留排版必需的字段，取值语义与 `utils/printer` 的 `PrinterSetting` 对齐，
 * 但**不 import 它**——那会把整份打印器（含样式副作用）拖进模型层。
 * 打印时由 `printDoc` 负责把本结构映射成 `PrinterSetting`。
 */
export interface DocPageSetup {
    /** 纸张方向 */
    orientation?: 'portrait' | 'landscape';
    /** 纸张大小 */
    paperSize?: 'A4' | 'A5' | 'A3' | 'Letter' | 'Legal';
    /** 页边距预设 */
    marginPreset?: 'normal' | 'wide' | 'narrow';
    /** 缩放模式 */
    scaleMode?: 'none' | 'fit-page' | 'custom';
    /** 自定义缩放比例 */
    scaleRatio?: number;
}

/**
 * 文档视图状态（每份文档各自记住的「看的姿势」）
 *
 * 与表格一致：缩放与滚动位置属于**视图状态**，随文件持久化，重新打开时恢复。
 * 三项都可缺省 —— 老存档没有这段时按「适应页宽 + 从页首开始」处理。
 */
export interface DocViewState {
    /** 视图缩放（0.5–4） */
    zoom?: number;
    /** 横向滚动位置（px） */
    scrollLeft?: number;
    /** 纵向滚动位置（px） */
    scrollTop?: number;
}

/** 文档元信息 */
export interface DocMeta {
    /** 标题（首个标题块或文件名，供另存为 / 打印使用） */
    title: string;
    /** 当前视图：富文本 / Markdown 源码 */
    view: 'rich' | 'markdown';
    /** 页面设置 */
    pageSetup: DocPageSetup;
    /** 视图状态（缩放 / 滚动位置） */
    viewState?: DocViewState;
}

/**
 * 文档存档（`docs` 仓的记录体）
 *
 * 与表格存档 v4 不同，这里**不分片**：文档规模是 10²–10⁴ 块，
 * 结构化克隆直存即可；分片是为百万级单元格准备的，用在这里只增复杂度。
 */
export interface DocSnapshot {
    /** 存档版本；结构不兼容时递增，装载侧据此决定迁移或丢弃 */
    v: number;
    /** 元信息 */
    meta: DocMeta;
    /** 块树（真源） */
    blocks: DocBlock[];
}

/**
 * 默认字体族 / 字号 / 缩进步长 / 纸张与页边距 / 存档版本
 *
 * 定义统一在 `core/constant.ts` 的「文档（Word）实现常量」段（本文件只在这里使用版本号）。
 * 渲染、命令、打印、导出都从同一处取值，避免「点了有缩进、重渲染后跳一格」这类分裂。
 */

/**
 * 块 id 发号器
 *
 * 用「时间戳 + 自增」而不是随机串：同一毫秒内连续建块也不会撞号，
 * 且生成的 id 单调递增，调试时能直接看出块的创建顺序。
 */
let blockSeq = 0;

/**
 * 生成一个块 id
 * @returns 形如 `b3f2a1-17` 的稳定标识
 */
export function nextBlockId(): string {
    blockSeq += 1;
    return `b${Date.now().toString(36)}-${blockSeq.toString(36)}`;
}

/**
 * 把纯文本拆成 `Char[]`（每个码点一个实例）
 *
 * 用 `Array.from` 而不是 `split('')`：后者会把 emoji / 生僻字（代理对）劈成两半，
 * 使「一个字符」变成两个 `Char`，光标与区间样式都会错位。
 * @param text - 纯文本
 * @param style - 可选的整体样式补丁（新建块时可带上当前输入样式）
 * @returns Char 数组
 */
export function charsFromText(text: string, style?: CharStylePatch): Char[] {
    const chars: Char[] = [];
    const source = text ?? '';
    for (let i = 0; i < source.length; i++) {
        // 逐码点推进：代理对的第二个码元不单独成 Char
        const code = source.charCodeAt(i);
        let piece = source.charAt(i);
        if (code >= 0xd800 && code <= 0xdbff && i + 1 < source.length) {
            piece += source.charAt(i + 1);
            i += 1;
        }
        const char = Char.of(piece);
        if (style) writeCharStyle(char, style);
        chars.push(char);
    }
    return chars;
}

/**
 * 把一个样式补丁写进单个 `Char`（稀疏：只写给出的字段）
 * @param char - 目标字符
 * @param style - 样式补丁
 */
export function writeCharStyle(char: Char, style: CharStylePatch): void {
    if (style.fontFamily !== undefined) char.fontFamily = style.fontFamily;
    if (style.fontSize !== undefined) char.fontSize = style.fontSize;
    if (style.bold !== undefined) char.fontWeight = style.bold;
    if (style.italic !== undefined) char.fontStyle = style.italic;
    if (style.underline !== undefined) char.underline = style.underline;
    if (style.strike !== undefined) char.strikethrough = style.strike;
    if (style.color !== undefined) char.fontColor = style.color;
    // 高亮不在 Char 上，见 CharStylePatch.background
}

/**
 * 创建一个新块
 * @param options - 块类型与初始文本
 * @returns 新块
 */
export function createBlock(options: {
    type?: BlockType;
    text?: string;
    level?: number;
    listStyle?: DocListStyle;
    align?: DocAlign;
} = {}): DocBlock {
    const block: DocBlock = {
        id: nextBlockId(),
        type: options.type ?? 'paragraph',
        chars: charsFromText(options.text ?? ''),
    };
    if (options.level !== undefined) block.level = options.level;
    if (options.listStyle !== undefined) block.listStyle = options.listStyle;
    if (options.align !== undefined) block.align = options.align;
    return block;
}

/**
 * 取块内纯文本
 *
 * 与 `Cell.getText` 同构：只拼 `char`，不做任何排版加工。
 * @param block - 块
 * @returns 纯文本
 */
export function blockText(block: DocBlock): string {
    if (!block || !Array.isArray(block.chars)) return '';
    let text = '';
    for (let i = 0; i < block.chars.length; i++) {
        const piece = block.chars[i]?.char;
        if (piece != null) text += String(piece);
    }
    return text;
}

/**
 * 取块内码点长度（等价于 `block.chars.length`，但容忍脏数据）
 * @param block - 块
 * @returns 码点数量
 */
export function blockLength(block: DocBlock): number {
    return block && Array.isArray(block.chars) ? block.chars.length : 0;
}

/**
 * 用纯文本替换块内容（保留块级属性，重建字符数组）
 *
 * 供 Markdown / 纯文本导入与「设置段落文本」类命令使用。字符样式会被丢弃 ——
 * 调用方要保留样式时应当直接改 `chars`，而不是走这里。
 * @param block - 目标块
 * @param text - 新文本
 */
export function setBlockText(block: DocBlock, text: string): void {
    block.chars = charsFromText(text);
}

/**
 * 对字符区间应用样式（与 `Cell.setSelectedCharsFont` 同构）
 *
 * 区间按**码点下标**、闭开区间。越界一律收敛到有效范围（不抛错）：
 * 调用方的区间来自 DOM 选区，边界漂一格是常态，抛错只会让工具栏按钮时灵时不灵。
 * @param chars - 目标字符数组
 * @param start - 起始下标（含）
 * @param end - 结束下标（不含）
 * @param style - 样式补丁
 */
export function applyCharStyle(
    chars: Char[],
    start: number,
    end: number,
    style: CharStylePatch
): void {
    if (!Array.isArray(chars) || chars.length === 0) return;
    const from = Math.max(0, Math.min(start, chars.length));
    const to = Math.max(from, Math.min(end, chars.length));
    for (let i = from; i < to; i++) {
        writeCharStyle(chars[i], style);
    }
}

/**
 * 查询字符区间的代表样式（取区间内**首个**字符，与表格 `queryState` 口径一致）
 *
 * 为什么不做「多数字符一致才算选中」：工具栏按钮的高亮是「提示当前输入会得到什么」，
 * 首个字符就是答案；做一致性统计反而会让「选中一段混排文本」时所有按钮全灭。
 * @param chars - 字符数组
 * @param start - 起始下标（含）
 * @param end - 结束下标（不含）
 * @returns 代表样式；区间为空时返回空对象
 */
export function queryCharStyle(chars: Char[], start: number, end: number): CharStylePatch {
    const style: CharStylePatch = {};
    if (!Array.isArray(chars) || chars.length === 0) return style;
    const from = Math.max(0, Math.min(start, chars.length - 1));
    const to = Math.max(from + 1, Math.min(end, chars.length));
    if (to <= from) return style;
    const head = chars[from];
    if (!head) return style;
    if (head.fontFamily !== undefined) style.fontFamily = head.fontFamily;
    if (head.fontSize !== undefined) style.fontSize = head.fontSize;
    if (head.fontWeight !== undefined) style.bold = head.fontWeight;
    if (head.fontStyle !== undefined) style.italic = head.fontStyle;
    if (head.underline !== undefined) style.underline = head.underline;
    if (head.strikethrough !== undefined) style.strike = head.strikethrough;
    if (head.fontColor !== undefined) style.color = head.fontColor;
    return style;
}

/**
 * 取某个码点位置上的区间属性（链接 / 高亮 / 上下标）
 * @param block - 块
 * @param index - 码点下标
 * @returns 命中的区间；未命中时为 null
 */
export function spanAt(block: DocBlock, index: number): DocSpan | null {
    if (!block || !Array.isArray(block.spans)) return null;
    for (let i = 0; i < block.spans.length; i++) {
        const span = block.spans[i];
        if (index >= span.start && index < span.end) return span;
    }
    return null;
}

/**
 * 在块上登记一段区间属性（同类型区间先清除再写入，避免叠加出多层链接）
 * @param block - 块
 * @param span - 新区间
 */
export function addSpan(block: DocBlock, span: DocSpan): void {
    if (!Array.isArray(block.spans)) block.spans = [];
    const keep = block.spans.filter((item) => !overlaps(item, span));
    keep.push(span);
    keep.sort((a, b) => a.start - b.start);
    block.spans = keep;
}

/**
 * 清除块上某个字符区间的全部属性
 * @param block - 块
 * @param start - 起始下标
 * @param end - 结束下标
 */
export function clearSpans(block: DocBlock, start: number, end: number): void {
    if (!Array.isArray(block.spans)) return;
    const probe: DocSpan = { start, end };
    block.spans = block.spans.filter((item) => !overlaps(item, probe));
}

/**
 * 判断两段区间是否相交
 * @param a - 区间一
 * @param b - 区间二
 * @returns 是否相交
 */
function overlaps(a: DocSpan, b: DocSpan): boolean {
    return a.start < b.end && b.start < a.end;
}

/**
 * 深拷贝块树
 *
 * 用 `Char.clone` 而不是结构化克隆：后者在 ES5 target 下不可用，
 * 且 `clone` 保持了稀疏语义（不给未赋值字段物化属性）。
 * @param blocks - 源块树
 * @returns 拷贝
 */
export function cloneBlocks(blocks: DocBlock[]): DocBlock[] {
    return (blocks ?? []).map((block) => cloneBlock(block));
}

/**
 * 深拷贝单个块（含子块与行列）
 * @param block - 源块
 * @returns 拷贝
 */
export function cloneBlock(block: DocBlock): DocBlock {
    const copy: DocBlock = {
        id: block.id,
        type: block.type,
        chars: (block.chars ?? []).map((char) => char.clone()),
    };
    if (block.spans) copy.spans = block.spans.map((span) => ({ ...span }));
    if (block.align !== undefined) copy.align = block.align;
    if (block.level !== undefined) copy.level = block.level;
    if (block.listStyle !== undefined) copy.listStyle = block.listStyle;
    if (block.indent !== undefined) copy.indent = block.indent;
    if (block.children) copy.children = block.children.map((child) => cloneBlock(child));
    if (block.rows) copy.rows = block.rows.map((row) => row.map((cell) => cloneBlock(cell)));
    if (block.imageRef) copy.imageRef = { ...block.imageRef };
    return copy;
}

/**
 * 重建块树中的 `Char` 原型
 *
 * **IndexedDB 的结构化克隆会丢原型**：`Char` 实例存进去、读回来是普通对象，
 * 之后调 `clone()` / `toJSON()` 会直接抛 `is not a function`。
 * 这与「解析器在主线程生成 Cell/Char 实例」是同一个坑，文档仓同理。
 *
 * 另外兜底一种更隐蔽的形态：`Char.toJSON()` 返回的是**字符串**（表格域遗留 API，
 * 受「Char 零改动」约束不能修），所以只要有代码对存档做过一次 `JSON.stringify`，
 * `chars` 就会变成 `string[]`。结构化克隆不走 `toJSON`，正常链路碰不到，
 * 但一旦发生就是「样式静默全丢」，因此这里顺手解析回来。
 * @param blocks - 从存储读回的块树
 * @returns 原型已重建的块树（原地修改并返回）
 */
export function restoreBlocks(blocks: DocBlock[]): DocBlock[] {
    if (!Array.isArray(blocks)) return [];
    for (let i = 0; i < blocks.length; i++) {
        const block = blocks[i];
        if (!block) continue;
        if (Array.isArray(block.chars)) {
            // 用「重建数组」而不是 map：字符串降级时一个条目可能展开成多个字符，
            // 否则 `blockLength` / 光标偏移会把 8 个字错算成 1 个。
            const rebuilt: Char[] = [];
            for (let c = 0; c < block.chars.length; c++) {
                const char = block.chars[c];
                if (char instanceof Char) {
                    rebuilt.push(char);
                } else if (typeof char === 'string') {
                    let parsed: any = null;
                    try {
                        parsed = JSON.parse(char);
                    } catch {
                        parsed = null;
                    }
                    if (parsed && typeof parsed === 'object') {
                        rebuilt.push(Char.fromRaw(parsed));
                    } else {
                        const plain = charsFromText(char);
                        for (let p = 0; p < plain.length; p++) rebuilt.push(plain[p]);
                    }
                } else {
                    rebuilt.push(Char.fromRaw(char));
                }
            }
            block.chars = rebuilt;
        }
        if (Array.isArray(block.children)) restoreBlocks(block.children);
        if (Array.isArray(block.rows)) {
            for (let r = 0; r < block.rows.length; r++) {
                const row = block.rows[r];
                if (Array.isArray(row)) restoreBlocks(row);
            }
        }
    }
    return blocks;
}

/**
 * 规范化块树：补齐 id、纠正类型、保证至少有一个块
 *
 * 装载任意来源（存档、docx、Markdown）后都过一遍，使下游可以无条件相信
 * 「一定有块、每块都有 id 与 chars 数组」。
 * @param blocks - 原始块树
 * @returns 规范化后的块树（原地修改）
 */
export function normalizeBlocks(blocks: DocBlock[]): DocBlock[] {
    const list = Array.isArray(blocks) ? blocks : [];
    const seen = new Set<string>();
    const walk = (items: DocBlock[]): void => {
        for (let i = 0; i < items.length; i++) {
            const block = items[i];
            if (!block) continue;
            if (!block.id || seen.has(block.id)) block.id = nextBlockId();
            seen.add(block.id);
            if (!block.type) block.type = 'paragraph';
            if (!Array.isArray(block.chars)) block.chars = [];
            if (Array.isArray(block.children)) walk(block.children);
            if (Array.isArray(block.rows)) {
                for (let r = 0; r < block.rows.length; r++) {
                    const row = block.rows[r];
                    if (Array.isArray(row)) walk(row);
                }
            }
        }
    };
    walk(list);
    if (list.length === 0) list.push(createBlock());
    return list;
}

/**
 * 创建空白文档存档
 * @param title - 文档标题
 * @returns 存档
 */
export function createBlankSnapshot(title: string = ''): DocSnapshot {
    return {
        v: DOC_SNAPSHOT_VERSION,
        meta: {
            title,
            view: 'rich',
            pageSetup: { orientation: 'portrait', paperSize: 'A4', marginPreset: 'normal', scaleMode: 'none' },
        },
        blocks: [createBlock()],
    };
}

/**
 * 判断存档是否为空文档（没有可见文本）
 *
 * 供 `docs` 仓的孤儿清理使用：**只删空存档**，非空的一律留着（与表格侧同口径）。
 * @param snapshot - 存档
 * @returns 是否为空
 */
export function isEmptySnapshot(snapshot: DocSnapshot | null | undefined): boolean {
    if (!snapshot || !Array.isArray(snapshot.blocks)) return true;
    for (let i = 0; i < snapshot.blocks.length; i++) {
        const block = snapshot.blocks[i];
        if (!block) continue;
        if (block.type === 'hr' || block.type === 'image' || block.type === 'table') return false;
        if (blockText(block).trim().length > 0) return false;
        if (block.children && !isEmptyBlocks(block.children)) return false;
    }
    return true;
}

/**
 * 判断一组块是否全为空（供 {@link isEmptySnapshot} 递归）
 * @param blocks - 块集合
 * @returns 是否全空
 */
function isEmptyBlocks(blocks: DocBlock[]): boolean {
    for (let i = 0; i < blocks.length; i++) {
        const block = blocks[i];
        if (!block) continue;
        if (blockText(block).trim().length > 0) return false;
        if (block.children && !isEmptyBlocks(block.children)) return false;
    }
    return true;
}

/**
 * 取整棵树的首个非空文本（用于推断标题）
 * @param blocks - 块树
 * @param limit - 截断长度
 * @returns 标题文本；全空时返回空串
 */
export function firstTextOf(blocks: DocBlock[], limit: number = 60): string {
    for (let i = 0; i < blocks.length; i++) {
        const text = blockText(blocks[i]).trim();
        if (text) return text.slice(0, limit);
    }
    return '';
}

/**
 * 断言块树里每个 `Char` 都是 `Char` 实例（测试与装载后自检用）
 *
 * 结构克隆丢原型是**静默故障**：不报错，直到某处调用 `char.clone()` 才炸。
 * 因此装载路径显式调用一次，把问题暴露在装载处而不是渲染处。
 * @param blocks - 块树
 * @returns 是否全部为 Char 实例
 */
export function isCharTreeIntact(blocks: DocBlock[]): boolean {
    for (let i = 0; i < blocks.length; i++) {
        const block = blocks[i];
        if (!block) continue;
        if (Array.isArray(block.chars)) {
            for (let c = 0; c < block.chars.length; c++) {
                if (!(block.chars[c] instanceof Char)) return false;
            }
        }
        if (Array.isArray(block.children) && !isCharTreeIntact(block.children)) return false;
        if (Array.isArray(block.rows)) {
            for (let r = 0; r < block.rows.length; r++) {
                const row = block.rows[r];
                if (Array.isArray(row) && !isCharTreeIntact(row)) return false;
            }
        }
    }
    return true;
}
