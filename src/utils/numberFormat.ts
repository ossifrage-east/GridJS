/**
 * 数字格式化工具模块
 *
 * 提供 Excel 风格的数字格式码解析与格式化能力。
 * 支持常规、数值（小数位/千分位）、货币、百分比、科学记数、文本等格式码。
 * 支持日期/时间格式码（yyyy/m/d/h:mm/ss/AM-PM 等），含 Excel 序列号与文本日期解析。
 *
 * @module numberFormat
 */

/**
 * 按格式码将数值格式化为显示文本
 *
 * 支持的格式码模式：
 * - 'General'：常规，原样输出
 * - '@'：文本，原样输出
 * - '0' / '0.00'：固定位数小数
 * - '#,##0' / '#,##0.00'：千分位 + 小数
 * - '0%' / '0.00%'：百分比（值 × 100 后追加 %）
 * - '¥#,##0.00' / '$#,##0.00' 等：货币符号前缀 + 数值格式
 * - '0.00E+00'：科学记数法
 * - 多段码 'positive;negative;zero'：按值的正/负/零选择对应段
 * - 括号负数样式 '#,##0.00;(#,##0.00)'：负数用括号包裹
 * - 颜色标签 '[Red]'：仅解析时剥离，颜色不在文本中体现（需由调用方在绘制层处理）
 *
 * @param value - 要格式化的数值
 * @param code  - 格式码字符串
 * @returns 格式化后的显示文本
 */
export function formatNumberValue(value: number, code: string): string {
    // 多段格式码：按 正;负;零 切分，选择对应段
    const sections = code.split(';');
    const isNegative = value < 0;
    const absValue = Math.abs(value);

    let section: string;
    if (value > 0) {
        section = sections[0];
    } else if (value < 0) {
        section = sections[1] ?? sections[0];
    } else {
        section = sections[2] ?? sections[0];
    }

    // 去除颜色标签（如 [Red]），颜色仅在 UI 预览中体现，不在文本中输出
    section = section.replace(/\[Red\]/gi, '');

    // 文本格式 / 常规格式：原样输出
    const trimmed = section.trim();
    if (trimmed === '@' || trimmed === 'General' || trimmed === '') {
        return String(value);
    }

    // 括号负数样式：负数用括号包裹
    const useParentheses = /[()]/.test(section);
    section = section.replace(/[()]/g, '');

    // 百分比：值 × 100，稍后追加 %
    const isPercent = section.includes('%');
    let numValue = absValue;
    if (isPercent) {
        numValue *= 100;
        section = section.replace(/%/g, '');
    }

    // 科学记数法
    const sciMatch = section.match(/E\+?(0+)/i);
    if (sciMatch) {
        const expDigits = sciMatch[1].length;
        const decimalPlaces = getDecimalPlaces(section);
        const expStr = numValue.toExponential(decimalPlaces);
        const parts = expStr.split(/[eE]/);
        const mantissa = parts[0];
        let exponent = parseInt(parts[1] ?? '0', 10);
        const expSign = exponent >= 0 ? '+' : '-';
        exponent = Math.abs(exponent);
        const expPadded = String(exponent).padStart(expDigits, '0');
        let result = `${mantissa}E${expSign}${expPadded}`;
        if (isPercent) result += '%';
        if (isNegative) {
            result = useParentheses ? `(${result})` : `-${result}`;
        }
        return result;
    }

    // 小数位数
    const decimalPlaces = getDecimalPlaces(section);

    // 千分位分隔符
    const useThousands = section.includes(',');

    // 货币符号前缀/后缀：提取格式字符（0#,.）以外的字符
    const prefixMatch = section.match(/^[^0#,.]+/);
    const suffixMatch = section.match(/[^0#,.]+$/);
    const prefix = prefixMatch ? prefixMatch[0] : '';
    const suffix = suffixMatch ? suffixMatch[0] : '';

    // 格式化数值
    let formatted = numValue.toFixed(decimalPlaces);
    if (useThousands) {
        formatted = addThousandsSeparator(formatted);
    }

    let result = `${prefix}${formatted}${suffix}`;
    if (isPercent) result += '%';

    // 应用负数样式
    if (isNegative) {
        if (useParentheses) {
            result = `(${result})`;
        } else {
            result = `-${result}`;
        }
    }

    return result;
}

/**
 * 获取单元格的显示文本（应用数字格式化）
 *
 * 当格式码为 'General'、'@' 或不存在时返回原始文本；
 * 当原始文本无法解析为数字时返回原始文本（非数字不受格式码影响）；
 * 否则按格式码格式化数值。
 *
 * @param rawText      - 单元格原始文本（chars 拼接结果）
 * @param numberFormat - 数字格式码（undefined / 'General' 表示无格式）
 * @returns 显示文本
 */
export function getCellDisplayText(rawText: string, numberFormat?: string): string {
    if (!numberFormat || numberFormat === 'General' || numberFormat === '@') {
        return rawText;
    }
    // 日期/时间格式码：解析为 Date 后按码渲染
    if (isDateTimeCode(numberFormat)) {
        const date = parseDateValue(rawText);
        if (!date) return rawText;
        return formatDateTimeValue(date, numberFormat);
    }
    const trimmed = rawText.trim();
    if (!/^-?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(trimmed)) {
        return rawText;
    }
    const value = parseFloat(trimmed);
    if (!Number.isFinite(value)) {
        return rawText;
    }
    return formatNumberValue(value, numberFormat);
}

// ── 日期/时间格式化 ──

/** 中文星期映射（短）：ddd → 周一..周日 */
const WEEKDAY_SHORT_ZH = ['周日', '周一', '周二', '周三', '周四', '周五', '周六'];
/** 中文星期映射（长）：dddd → 星期一..星期日 */
const WEEKDAY_LONG_ZH = ['星期日', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六'];
/** 中文月份映射（短）：mmm → 1月..12月 */
const MONTH_SHORT_ZH = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
/** 中文月份映射（长）：mmmm → January..December */
const MONTH_LONG_ZH = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

/**
 * 去除格式码中引号内的字面量（"..." → 空），保留结构以便检测 token
 * @param code - 格式码
 * @returns 去除引号字面量后的码
 * @private
 */
function stripQuotedLiterals(code: string): string {
    return code.replace(/"[^"]*"/g, '');
}

/**
 * 判断格式码是否为日期/时间格式码
 *
 * 数字格式码（0/#/,/. /%/E 等）不含 y/d/h/s/m 字母（引号内除外），
 * 因此去颜色标签和引号字面量后若含这些字母即为日期/时间码。
 *
 * @param code - 格式码
 * @returns 是否为日期/时间码
 */
export function isDateTimeCode(code: string): boolean {
    // 去除颜色标签 [Red] 等
    let s = code.replace(/\[[^\]]*\]/g, '');
    // 去除引号内字面量
    s = stripQuotedLiterals(s);
    // 检查是否含日期/时间 token 字符
    return /[ydhsm]/i.test(s);
}

/**
 * 将单元格原始文本解析为 Date
 *
 * 解析顺序：
 * 1. 纯数字 → Excel 序列号（基准 1899-12-30，value=1 对应 1899-12-31）
 * 2. new Date(rawText)（支持 ISO '2026-10-10'、'2026/10/10 14:30:45' 等）
 * 3. 中文格式 '2026年10月10日' / '2026年10月10日 14:30'（正则提取 y/m/d/h/m/s）
 * 4. 均失败 → 返回 null
 *
 * @param rawText - 单元格原始文本
 * @returns 解析后的 Date，或 null（解析失败）
 * @private
 */
function parseDateValue(rawText: string): Date | null {
    const trimmed = rawText.trim();

    // 纯数字 → Excel 序列号
    if (/^-?\d+(\.\d+)?$/.test(trimmed)) {
        const value = parseFloat(trimmed);
        if (!Number.isFinite(value)) return null;
        // 基准 1899-12-30，先按 UTC 算出绝对时刻
        const ms = Date.UTC(1899, 11, 30) + value * 86400000;
        const utc = new Date(ms);
        if (isNaN(utc.getTime())) return null;
        // 以 UTC 分量构造本地 Date，使本地 getHours() 等返回序列号代表的原始时分秒（避免时区漂移）
        return new Date(
            utc.getUTCFullYear(), utc.getUTCMonth(), utc.getUTCDate(),
            utc.getUTCHours(), utc.getUTCMinutes(), utc.getUTCSeconds()
        );
    }

    // 尝试 new Date(rawText)（支持 ISO、斜杠日期等）
    let date = new Date(trimmed);
    if (!isNaN(date.getTime())) return date;

    // 尝试中文格式：2026年10月10日 / 2026年10月10日 14:30
    const cnMatch = trimmed.match(/(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日(?:\s+(\d{1,2})\s*[:：]\s*(\d{1,2})(?:\s*[:：]\s*(\d{1,2}))?)?/);
    if (cnMatch) {
        const y = cnMatch[1];
        const mo = cnMatch[2];
        const d = cnMatch[3];
        const h = cnMatch[4];
        const mi = cnMatch[5];
        const s = cnMatch[6];
        if (y && mo && d) {
            date = new Date(
                parseInt(y, 10), parseInt(mo, 10) - 1, parseInt(d, 10),
                h ? parseInt(h, 10) : 0, mi ? parseInt(mi, 10) : 0, s ? parseInt(s, 10) : 0
            );
            if (!isNaN(date.getTime())) return date;
        }
    }

    return null;
}

/**
 * 从指定位置向前查找最近的非分隔符、非引号字面量的 token 字符
 * @param section - 格式码段（含引号字面量）
 * @param pos     - 起始位置（不含）
 * @returns 最近的 token 字符（小写），无则空串
 * @private
 */
function findPrevTokenChar(section: string, pos: number): string {
    let j = pos - 1;
    while (j >= 0) {
        if (section[j] === '"') {
            // 跳过引号字面量（向后找匹配的开引号）
            j = section.lastIndexOf('"', j - 1);
            if (j === -1) return '';
            j--;
            continue;
        }
        if (/[:\-/\s,]/.test(section[j])) { j--; continue; }
        return section[j].toLowerCase();
    }
    return '';
}

/**
 * 从指定位置向后查找最近的非分隔符、非引号字面量的 token 字符
 * @param section - 格式码段（含引号字面量）
 * @param pos     - 起始位置（含）
 * @returns 最近的 token 字符（小写），无则空串
 * @private
 */
function findNextTokenChar(section: string, pos: number): string {
    let j = pos;
    while (j < section.length) {
        if (section[j] === '"') {
            // 跳过引号字面量（向前找匹配的闭引号）
            const end = section.indexOf('"', j + 1);
            if (end === -1) return '';
            j = end + 1;
            continue;
        }
        if (/[:\-/\s,]/.test(section[j])) { j++; continue; }
        return section[j].toLowerCase();
    }
    return '';
}

/**
 * 判断 m/mm 在当前位置应解释为"分"还是"月"
 *
 * 规则：前方最近的 token 字符是 h（小时）或后方最近的 token 字符是 s（秒）→ 分；否则 → 月。
 * 参照 Excel 的 m/mm 歧义消解，支持 yyyy-mm-dd h:mm:ss 等混合码中日期段 mm=月、时间段 mm=分。
 *
 * @param section  - 格式码段
 * @param pos      - m/mm 起始位置
 * @param tokenLen - token 长度（1 或 2）
 * @returns true=分，false=月
 * @private
 */
function isMinuteContext(section: string, pos: number, tokenLen: number): boolean {
    const prevChar = findPrevTokenChar(section, pos);
    const nextChar = findNextTokenChar(section, pos + tokenLen);
    return prevChar === 'h' || nextChar === 's';
}

/**
 * 按日期/时间格式码将 Date 格式化为显示文本
 *
 * 支持的 token（大小写不敏感）：
 * - 年：yyyy（4 位）、yy（2 位）
 * - 月：m/mm（1/2 位）、mmm（英文缩写）、mmmm（英文全称）
 * - 日：d/dd（1/2 位）
 * - 星期：ddd（周一）、dddd（星期一）
 * - 时：h/hh（1/2 位，含 AM/PM 时为 12 小时制）
 * - 分：m/mm（时间上下文中）
 * - 秒：s/ss
 * - 时段：AM/PM、上午/下午
 * - 字面量：双引号内内容原样输出
 *
 * @param date - 日期对象
 * @param code - 日期/时间格式码
 * @returns 格式化后的显示文本
 */
export function formatDateTimeValue(date: Date, code: string): string {
    // 取第一段（日期/时间码通常只有一段）
    let section = code.split(';')[0];
    // 去除颜色标签
    section = section.replace(/\[[^\]]*\]/g, '');

    // 判断是否含 AM/PM（影响小时为 12 小时制）
    const stripped = stripQuotedLiterals(section);
    const hasAMPM = /AM\/PM/i.test(stripped);
    const hasCNAMPM = /上午\/下午/.test(stripped);
    const use12Hour = hasAMPM || hasCNAMPM;

    let result = '';
    let i = 0;
    while (i < section.length) {
        // 引号内字面量原样输出
        if (section[i] === '"') {
            const end = section.indexOf('"', i + 1);
            if (end === -1) { result += section.slice(i + 1); break; }
            result += section.slice(i + 1, end);
            i = end + 1;
            continue;
        }

        const rest = section.slice(i);

        // AM/PM（5 字符）
        if (rest.match(/^AM\/PM/i)) {
            result += date.getHours() < 12 ? 'AM' : 'PM';
            i += 5; continue;
        }
        // 上午/下午
        if (rest.startsWith('上午/下午')) {
            result += date.getHours() < 12 ? '上午' : '下午';
            i += 5; continue;
        }
        // yyyy
        if (rest.match(/^yyyy/i)) {
            result += String(date.getFullYear());
            i += 4; continue;
        }
        // mmmm
        if (rest.match(/^mmmm/i)) {
            result += MONTH_LONG_ZH[date.getMonth()];
            i += 4; continue;
        }
        // dddd
        if (rest.match(/^dddd/i)) {
            result += WEEKDAY_LONG_ZH[date.getDay()];
            i += 4; continue;
        }
        // mmm
        if (rest.match(/^mmm/i)) {
            result += MONTH_SHORT_ZH[date.getMonth()];
            i += 3; continue;
        }
        // ddd
        if (rest.match(/^ddd/i)) {
            result += WEEKDAY_SHORT_ZH[date.getDay()];
            i += 3; continue;
        }
        // yy
        if (rest.match(/^yy/i)) {
            result += String(date.getFullYear()).slice(-2).padStart(2, '0');
            i += 2; continue;
        }
        // mm
        if (rest.match(/^mm/i)) {
            if (isMinuteContext(section, i, 2)) {
                result += String(date.getMinutes()).padStart(2, '0');
            } else {
                result += String(date.getMonth() + 1).padStart(2, '0');
            }
            i += 2; continue;
        }
        // dd
        if (rest.match(/^dd/i)) {
            result += String(date.getDate()).padStart(2, '0');
            i += 2; continue;
        }
        // hh
        if (rest.match(/^hh/i)) {
            const h = use12Hour ? (date.getHours() % 12 || 12) : date.getHours();
            result += String(h).padStart(2, '0');
            i += 2; continue;
        }
        // ss
        if (rest.match(/^ss/i)) {
            result += String(date.getSeconds()).padStart(2, '0');
            i += 2; continue;
        }
        // m（单字符，mm/mmm/mmmm 已在上方匹配）
        if (rest.match(/^m/i)) {
            if (isMinuteContext(section, i, 1)) {
                result += String(date.getMinutes());
            } else {
                result += String(date.getMonth() + 1);
            }
            i += 1; continue;
        }
        // d（单字符）
        if (rest.match(/^d/i)) {
            result += String(date.getDate());
            i += 1; continue;
        }
        // h（单字符）
        if (rest.match(/^h/i)) {
            const h = use12Hour ? (date.getHours() % 12 || 12) : date.getHours();
            result += String(h);
            i += 1; continue;
        }
        // s（单字符）
        if (rest.match(/^s/i)) {
            result += String(date.getSeconds());
            i += 1; continue;
        }

        // 其他字符原样输出
        result += section[i];
        i++;
    }
    return result;
}

/**
 * 从格式段中解析小数位数（'.' 后 '0' 的个数）
 * @param section - 格式码段
 * @returns 小数位数（无小数则返回 0）
 * @private
 */
function getDecimalPlaces(section: string): number {
    const match = section.match(/\.(0+)/);
    return match ? match[1].length : 0;
}

/**
 * 为数值字符串的整数部分添加千分位分隔符
 * @param numStr - 数值字符串（如 "1234567.89"）
 * @returns 添加千分位后的字符串（如 "1,234,567.89"）
 * @private
 */
function addThousandsSeparator(numStr: string): string {
    const parts = numStr.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return parts.join('.');
}
