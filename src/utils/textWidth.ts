/*  # 字符宽度与缩进步长 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

/**
 * 字符宽度（单字节 / 双字节）与「按字符计」的缩进步长
 *
 * ## 为什么需要区分单双字节
 *
 * 缩进是按**字符个数**表达的（Excel 的「增加缩进」也是一级 = 一个字符宽），
 * 而一个字符占多宽取决于它属于哪一类：
 * - 双字节（CJK 汉字、假名、韩文、全角标点/字母）：字宽 = 1em，与字号相等；
 * - 单字节（ASCII 字母数字、半角标点）：字宽 ≈ 0.5em。
 *
 * 若一律按 2em 缩进，中文单元格正好是 2 个字宽，而纯数字/英文单元格会缩进
 * 4 个字宽 —— 同一级缩进在不同内容的格子上看起来完全不一样。
 *
 * ## 为什么是「表驱动」而不是 measureText
 *
 * 这里只做**分类 + 比例**，不建 canvas 上下文：
 * 1. 该模块被工具栏直接引用，而工具栏不应为了量一个字宽去建离屏画布；
 * 2. 比例（1em / 0.5em）是字体无关的通用约定，结果稳定、可测试；
 * 3. 逐帧绘制路径用的是真实的 `measureText`（见 `Canvas.calculateCharGeometry`），
 *    两处的目的不同：那边要对齐像素，这边只决定「一次缩进多少」。
 *
 * @module textWidth
 */

/**
 * 单字节字符的宽度（以 em 为单位）
 *
 * 拉丁字母在常见无衬线字体里约为 0.5em（有的字体 0.52~0.55），取 0.5 是
 * 与「一个半角字符 = 半个全角字符」这一排版常识一致的值。
 */
export const HALF_WIDTH_EM = 0.5;

/**
 * 一次缩进相当于几个字符宽
 *
 * 2 = 「2em」这个说法在汉字单元格上的等价物；单字节字符则按 0.5em/字折算，
 * 保证两类内容的缩进视觉上都是「两个字」。
 */
export const INDENT_CHARS = 2;


/** 双字节字符的宽度（以 em 为单位）：正好一个字号 */
export const FULL_WIDTH_EM = 1;

/**
 * 码点是否属于「双字节 / 全角」字符
 *
 * 覆盖 Unicode 里的 W（Wide）与 F（Fullwidth）两类（East Asian Width），
 * 取常用区段即可：CJK 标点、假名、韩文、CJK 统一表意文字（含扩展 A/B）、
 * 全角字母数字与符号、以及常见的 emoji。
 * @param {number} code - 码点（`codePointAt(0)` 的结果，surrogate pair 已合并）
 * @returns {boolean} 双字节为 true
 */
export function isWideCodePoint(code: number): boolean {
    if (!Number.isFinite(code) || code < 0) return false;
    if (code < 0x1100) return false;                    // 拉丁 / 希腊 / 西里尔等，都是半角
    return (
        (code >= 0x1100 && code <= 0x115F)              // 韩文字母（Jamo）
        || (code >= 0x2E80 && code <= 0x303E)           // CJK 部首、康熙部首、CJK 标点
        || (code >= 0x3041 && code <= 0x33FF)           // 平假名 / 片假名 / 注音 / CJK 兼容
        || (code >= 0x3400 && code <= 0x4DBF)           // CJK 扩展 A
        || (code >= 0x4E00 && code <= 0x9FFF)           // CJK 统一表意文字
        || (code >= 0xA000 && code <= 0xA4CF)           // 彝文
        || (code >= 0xA960 && code <= 0xA97F)           // 韩文字母扩展 A
        || (code >= 0xAC00 && code <= 0xD7A3)           // 韩文音节
        || (code >= 0xF900 && code <= 0xFAFF)           // CJK 兼容表意文字
        || (code >= 0xFE10 && code <= 0xFE19)           // 竖排标点
        || (code >= 0xFE30 && code <= 0xFE6F)           // CJK 兼容形式、小写变体
        || (code >= 0xFF00 && code <= 0xFF60)           // 全角 ASCII / 全角标点
        || (code >= 0xFFE0 && code <= 0xFFE6)           // 全角符号（￥ ＄ 等）
        || (code >= 0x1F300 && code <= 0x1F64F)         // emoji（表情）
        || (code >= 0x1F900 && code <= 0x1F9FF)         // emoji（补充）
        || (code >= 0x20000 && code <= 0x3FFFD)         // CJK 扩展 B 及以后
    );
}

/**
 * 单个字符是否双字节（空串 / undefined 视为单字节）
 * @param {string} [char] - 一个字符（取首个码点判断）
 * @returns {boolean} 双字节为 true
 */
export function isWideChar(char: string | undefined): boolean {
    if (!char) return false;
    const code = char.codePointAt(0);
    return code === undefined ? false : isWideCodePoint(code);
}

/**
 * 单个字符的宽度（以 em 为单位）
 * @param {string} [char] - 一个字符
 * @returns {number} 双字节 1em、单字节 0.5em
 */
export function charWidthEm(char: string | undefined): number {
    return isWideChar(char) ? FULL_WIDTH_EM : HALF_WIDTH_EM;
}

/**
 * 求「按字符计」的缩进步长（未乘缩放，单位 px）
 *
 * 一次缩进 = {@link INDENT_CHARS} 个字符宽，字符宽按该字符是单字节还是双字节取
 * `0.5em` / `1em`：
 * - 汉字单元格、16px 字号 → 2 × 16 = 32px（2 个字宽）
 * - 数字/英文单元格、16px 字号 → 2 × 8 = 16px（同样是 2 个字宽，视觉一致）
 *
 * **em 必须以调用方传进来的字号为准**（调用方应当每次都取当前单元格的实时字号），
 * 本函数不做任何缓存 —— 改过字号之后立刻点击也能拿到新的步长。
 * @param {number} fontSize - 基准字号（px），即 1em
 * @param {string} [sampleChar] - 代表字符（一般取单元格首字符），决定单/双字节
 * @returns {number} 缩进像素值（未乘 zoom）
 */
export function indentStepPx(fontSize: number, sampleChar: string | undefined): number {
    const em = Number.isFinite(fontSize) && fontSize > 0 ? fontSize : 16;
    return em * charWidthEm(sampleChar) * INDENT_CHARS;
}

