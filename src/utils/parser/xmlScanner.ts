/**
 * 流式 XML 扫描器（SAX 风格，无 DOM 依赖）
 *
 * ## 为什么需要它（性能约定，请勿回退为 DOMParser）
 *
 * 打开大数据量 xlsx 时，`DOMParser.parseFromString` 会把整份 sheet XML
 * （动辄数百 MB 文本）构建成**整棵 DOM 驻留内存**——DOM 节点内存是原始
 * 文本的数倍，是「打开大文件即崩溃」的首要内存峰值来源；且逐格
 * `getAttribute` / `firstChild` 是慢速 DOM 访问路径。
 *
 * 本扫描器**单趟逐字节**推进：只在遇到元素时回调，内存占用 O(1)
 * （不随文档大小增长），也不再依赖 `DOMParser`——因此**可以直接运行在
 * Web Worker 中**（`DOMParser` 是 Window 专有 API，Worker 里不存在），
 * 这也是把解析搬出主线程的前提。
 *
 * ## 语义约定（与 DOM 版对齐的取舍）
 *
 * - 元素名回调一律给 **localName**（去掉 `xdr:` 等前缀）；属性键保留
 *   **书写的完整名**（如 `r:id`），取属性可用 {@link attrValue}（先精确名，
 *   再按 localName 兜底，等价于 DOM 的 `getAttribute` + `getAttributeNS`）。
 * - 文本以「连续文本段」为单位回调（CDATA 段单独回调，不做实体解码），
 *   普通文本段已做实体解码（`&amp;` / `&#65;` / `&#x41;` 等）。
 * - 自闭合元素（`<c r="A1"/>`）先 `onStart` 后 `onEnd`。
 * - 注释 / 处理指令 / DOCTYPE 一律跳过。
 * - **不校验良构性**：损坏的 XML 产出部分结果而非报错（对大文件更鲁棒）。
 *
 * @module xmlScanner
 */

/** 属性表：键为 XML 中书写的完整属性名（可能带 `r:` 等前缀） */
export type XmlAttrs = Record<string, string>;

/**
 * 扫描回调
 *
 * `onStart` 的 `attrs` 参数仅在回调期间有效（扫描器复用/重建该对象），
 * 需要保留的值请在回调内取出。
 */
export interface XmlScanHandler {
    /** 元素开始（含自闭合）；name 为 localName */
    onStart?(name: string, attrs: XmlAttrs): void;
    /** 元素结束（含自闭合） */
    onEnd?(name: string): void;
    /** 文本段（已解码）；CDATA 段同样走这里（cData 为 true，内容不解码实体） */
    onText?(text: string, cData?: boolean): void;
}

/** 空属性表（无属性的元素复用，避免每次分配） */
const EMPTY_ATTRS: XmlAttrs = Object.freeze({}) as XmlAttrs;

/** 命名实体表 */
const NAMED_ENTITIES: Record<string, string> = {
    amp: '&', lt: '<', gt: '>', quot: '"', apos: "'"
};

/**
 * 取属性值：先按完整名精确匹配，再按 localName 兜底
 *
 * 等价于 DOM 的 `el.getAttribute('r:id') || el.getAttributeNS(ns, 'id')`。
 *
 * @param attrs - 属性表
 * @param name - 完整属性名（如 `r:id`）或 localName（如 `id`）
 * @returns 属性值；不存在时 undefined
 */
export function attrValue(attrs: XmlAttrs | undefined, name: string): string | undefined {
    if (!attrs) return undefined;
    if (attrs[name] !== undefined) return attrs[name];
    // localName 兜底：匹配任何前缀下的同名属性（r:embed / d:embed → embed）
    const suffix = ':' + name;
    for (const key in attrs) {
        if (key.endsWith(suffix)) return attrs[key];
    }
    return undefined;
}

/**
 * 解码字符串中的 XML 实体（`&amp;` `&#65;` `&#x41;` …）
 *
 * 未知的实体原样保留。仅单趟解码：`&amp;lt;` → `&lt;`（与 DOM 行为一致）。
 *
 * @param s - 原始字符串
 * @returns 解码后的字符串
 * @private
 */
function decodeEntities(s: string): string {
    if (s.indexOf('&') === -1) return s;
    return s.replace(/&([^;&\s]{1,10});/g, (whole, body: string) => {
        if (body.charCodeAt(0) === 0x23 /* # */) {
            const isHex = body.length > 1 && (body[1] === 'x' || body[1] === 'X');
            const code = parseInt(body.slice(isHex ? 2 : 1), isHex ? 16 : 10);
            if (Number.isFinite(code) && code > 0 && code <= 0x10FFFF) {
                try { return String.fromCodePoint(code); } catch { return whole; }
            }
            return whole;
        }
        const named = NAMED_ENTITIES[body];
        return named !== undefined ? named : whole;
    });
}

/**
 * 在字节数组中查找多字节序列（从 from 开始）
 *
 * @param bytes - 字节数组
 * @param seq - 目标序列（如 `-->` 的字节）
 * @param from - 起始偏移
 * @returns 首个匹配的偏移；找不到返回 -1
 * @private
 */
function indexOfSeq(bytes: Uint8Array, seq: number[], from: number): number {
    const first = seq[0];
    let i = bytes.indexOf(first, from);
    const lastStart = bytes.length - seq.length;
    outer: while (i !== -1) {
        if (i > lastStart) return -1;
        for (let k = 1; k < seq.length; k++) {
            if (bytes[i + k] !== seq[k]) {
                i = bytes.indexOf(first, i + 1);
                continue outer;
            }
        }
        return i;
    }
    return -1;
}

/** `-->` 的字节序列 */
const COMMENT_END = [0x2D, 0x2D, 0x3E];
/** `]]>` 的字节序列 */
const CDATA_END = [0x5D, 0x5D, 0x3E];
/** `?>` 的字节序列 */
const PI_END = [0x3F, 0x3E];
/** `<![CDATA[` 的字节序列 */
const CDATA_START = [0x21, 0x5B, 0x43, 0x44, 0x41, 0x54, 0x41, 0x5B];

/**
 * 流式扫描一份 UTF-8 XML 字节
 *
 * 单趟逐字节推进：定位 `<` → 分派到注释/CDATA/PI/结束标签/开始标签；
 * `<` 之间的内容作为文本段回调。元素名与属性名按 ASCII 快速截取。
 *
 * @param bytes - XML 字节（UTF-8）
 * @param handler - 回调集合
 * @example
 * scanXml(bytes, {
 *     onStart: (name, attrs) => { if (name === 'c') ref = attrs['r']; },
 *     onText: (t) => { buffer += t; }
 * });
 */
export function scanXml(bytes: Uint8Array, handler: XmlScanHandler): void {
    const { onStart, onEnd, onText } = handler;
    const decoder = new TextDecoder('utf-8');
    const n = bytes.length;
    let i = 0;

    /** 解码并回调一段普通文本（含实体解码） */
    const emitText = (start: number, end: number): void => {
        if (end <= start || !onText) return;
        onText(decodeEntities(decoder.decode(bytes.subarray(start, end))), false);
    };

    while (i < n) {
        const lt = bytes.indexOf(0x3C /* < */, i);
        if (lt === -1) {
            emitText(i, n);
            return;
        }
        emitText(i, lt);

        // —— 标记分派 ——
        const next = lt + 1;
        if (next >= n) return;

        // 注释 <!-- ... -->
        if (bytes[next] === 0x21 /* ! */) {
            if (next + 3 < n && bytes[next + 1] === 0x2D && bytes[next + 2] === 0x2D) {
                const end = indexOfSeq(bytes, COMMENT_END, next + 3);
                if (end === -1) return;
                i = end + 3;
                continue;
            }
            // CDATA <![CDATA[ ... ]]>
            let isCData = true;
            for (let k = 0; k < CDATA_START.length; k++) {
                if (bytes[next + k] !== CDATA_START[k]) { isCData = false; break; }
            }
            if (isCData) {
                const contentStart = next + CDATA_START.length;
                const end = indexOfSeq(bytes, CDATA_END, contentStart);
                if (end === -1) return;
                if (onText && end > contentStart) {
                    // CDATA 内容不做实体解码
                    onText(decoder.decode(bytes.subarray(contentStart, end)), true);
                }
                i = end + 3;
                continue;
            }
            // DOCTYPE / 其他声明：跳到「括号深度为 0 的 '>'」
            // （DOCTYPE 内部子集 <!ENTITY x "y"> 里可能出现 '>'，不能只找第一个）
            let depth = 0;
            let q = next;
            while (q < n) {
                const b = bytes[q];
                if (b === 0x5B /* [ */) depth++;
                else if (b === 0x5D /* ] */) depth = Math.max(0, depth - 1);
                else if (b === 0x3E /* > */ && depth === 0) break;
                q++;
            }
            if (q >= n) return;
            i = q + 1;
            continue;
        }

        // 处理指令 <?xml ... ?>
        if (bytes[next] === 0x3F /* ? */) {
            const end = indexOfSeq(bytes, PI_END, next + 1);
            if (end === -1) return;
            i = end + 2;
            continue;
        }

        // 结束标签 </name>
        if (bytes[next] === 0x2F /* / */) {
            const end = bytes.indexOf(0x3E /* > */, next);
            if (end === -1) return;
            if (onEnd) onEnd(localNameOf(decoder.decode(bytes.subarray(next + 1, end)).trim()));
            i = end + 1;
            continue;
        }

        // 开始标签 <name attr="v" ...> 或 <name .../>
        i = parseStartTag(bytes, lt, decoder, onStart, onEnd);
        if (i === -1) return;
    }
}

/**
 * 解析一个开始标签，返回扫描应继续的偏移（-1 表示终止）
 *
 * @param bytes - 全部字节
 * @param lt - '<' 的偏移
 * @param decoder - 共享解码器
 * @param onStart - 开始回调
 * @param onEnd - 结束回调（自闭合时紧跟调用）
 * @returns '>' 之后的偏移；标签非法时 -1
 * @private
 */
function parseStartTag(
    bytes: Uint8Array,
    lt: number,
    decoder: TextDecoder,
    onStart: XmlScanHandler['onStart'],
    onEnd: XmlScanHandler['onEnd']
): number {
    const n = bytes.length;
    let p = lt + 1;

    // —— 元素名：到空白 / '/' / '>' 为止 ——
    const nameStart = p;
    while (p < n) {
        const b = bytes[p];
        if (b === 0x20 || b === 0x09 || b === 0x0A || b === 0x0D || b === 0x2F || b === 0x3E) break;
        p++;
    }
    if (p === nameStart) {
        // '<' 后直接结束：非法标签，跳过
        const end = bytes.indexOf(0x3E, p);
        return end === -1 ? -1 : end + 1;
    }
    const rawName = decoder.decode(bytes.subarray(nameStart, p));

    // —— 属性 ——
    let attrs: XmlAttrs | undefined;
    let selfClosing = false;

    while (p < n) {
        // 跳过空白
        let b = bytes[p];
        while (b === 0x20 || b === 0x09 || b === 0x0A || b === 0x0D) {
            p++;
            if (p >= n) return -1;
            b = bytes[p];
        }
        if (b === 0x3E /* > */) { p++; break; }
        if (b === 0x2F /* / */) {
            selfClosing = true;
            p++;
            // 期望紧随 '>'
            if (p < n && bytes[p] === 0x3E) { p++; break; }
            continue;
        }

        // —— 属性名 ——
        const attrStart = p;
        while (p < n) {
            const c = bytes[p];
            if (c === 0x3D /* = */ || c === 0x20 || c === 0x09 || c === 0x0A || c === 0x0D
                || c === 0x2F || c === 0x3E) break;
            p++;
        }
        if (p === attrStart) { p++; continue; } // 畸形字符，跳过
        const attrName = decoder.decode(bytes.subarray(attrStart, p));

        // —— 属性值（= "..." 或 '...'；容忍无引号值）——
        let value = '';
        let q = p;
        // 跳过 '=' 与空白
        while (q < n && (bytes[q] === 0x20 || bytes[q] === 0x09 || bytes[q] === 0x0A || bytes[q] === 0x0D)) q++;
        if (q < n && bytes[q] === 0x3D /* = */) {
            q++;
            while (q < n && (bytes[q] === 0x20 || bytes[q] === 0x09 || bytes[q] === 0x0A || bytes[q] === 0x0D)) q++;
            if (q < n && (bytes[q] === 0x22 || bytes[q] === 0x27)) {
                const quote = bytes[q];
                const valueStart = q + 1;
                const valueEnd = bytes.indexOf(quote, valueStart);
                if (valueEnd === -1) return -1;
                value = decodeEntities(decoder.decode(bytes.subarray(valueStart, valueEnd)));
                p = valueEnd + 1;
            } else {
                // 无引号值：到空白或 '>' 为止
                const valueStart = q;
                while (q < n) {
                    const c = bytes[q];
                    if (c === 0x20 || c === 0x09 || c === 0x0A || c === 0x0D || c === 0x3E) break;
                    q++;
                }
                value = decodeEntities(decoder.decode(bytes.subarray(valueStart, q)));
                p = q;
            }
        } else {
            // 布尔型属性（无值）：值置空串
            p = q;
        }

        if (attrs === undefined) attrs = {};
        attrs[attrName] = value;
    }

    const local = localNameOf(rawName);
    if (onStart) onStart(local, attrs !== undefined ? attrs : EMPTY_ATTRS);
    if (selfClosing && onEnd) onEnd(local);
    return p;
}

/**
 * 完整元素名 → localName（去掉 `xdr:` 等前缀）
 *
 * @param rawName - 如 `xdr:pic`
 * @returns 如 `pic`
 * @private
 */
function localNameOf(rawName: string): string {
    const colon = rawName.indexOf(':');
    return colon === -1 ? rawName : rawName.slice(colon + 1);
}
