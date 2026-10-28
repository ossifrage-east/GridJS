import { Cell } from '../core/dataArchitecture/cell';
import { Char } from '../core/dataArchitecture/char';


/**
 * 解析单元格名（如 'AB12'）为行列号
 * @param {string | undefined} name - 单元格名
 * @returns {{col: number, row: number} | null} 列/行均从 1 起；格式非法返回 null
 */
function parseCellName(name: string | undefined): { col: number, row: number } | null {
    if (!name) return null;
    const match = /^([A-Za-z]+)(\d+)$/.exec(name);
    if (!match) return null;
    let col = 0;
    for (const ch of match[1].toUpperCase()) {
        col = col * 26 + (ch.charCodeAt(0) - 64);
    }
    return { col, row: parseInt(match[2], 10) };
}

/**
 * TSV 字段转义
 * 字段内含制表符、换行或双引号时按 CSV 规则加引号并双写内部引号（Excel/WPS 可识别）
 * @param {string} text - 原始文本
 * @returns {string} 可直接拼进 TSV 的字段
 */
function escapeTSVField(text: string): string {
    if (!/[\t\r\n"]/.test(text)) return text;
    return `"${text.replace(/"/g, '""')}"`;
}

/**
 * 剪贴板数据类型枚举
 */
export type ClipboardDataType = 'cell' | 'char';

/**
 * 剪贴板数据接口
 * 存储复制或剪切的单元格/字符数据
 */
export interface ClipboardData {
    /** 数据类型：单元格或字符 */
    type: ClipboardDataType;
    /** 单元格数组（复制单元格时使用） */
    cells?: Cell[];
    /** 源单元格范围（复制单元格时使用） */
    sourceRange?: string;
    /** 源范围内可见列的相对偏移序列（升序，复制时记录，用于粘贴紧凑排列；旧数据缺省回退为全量偏移） */
    visibleColOffsets?: number[];
    /** 源范围内可见行的相对偏移序列（升序，复制时记录，用于粘贴紧凑排列；旧数据缺省回退为全量偏移） */
    visibleRowOffsets?: number[];
    /** 字符数据（复制字符时使用） */
    chars?: Char[];
}

/**
 * ClipboardManager 类
 * 管理电子表格的剪贴板操作，包括复制、剪切、粘贴等功能
 * 提供系统剪贴板和内部剪贴板的读写接口
 *
 * @class ClipboardManager
 * @example
 * const clipboard = new ClipboardManager();
 * clipboard.copyToSystemClipboard('Hello World');
 */
export class ClipboardManager {
    /** 内部剪贴板数据 */
    private clipboard: ClipboardData | null = null;

    constructor() {
    }

    /**
     * 获取当前剪贴板数据
     * @returns {ClipboardData | null} 剪贴板数据
     */
    public getClipboard(): ClipboardData | null {
        return this.clipboard;
    }

    /**
     * 设置剪贴板数据
     * @param {ClipboardData} data - 剪贴板数据
     */
    public setClipboard(data: ClipboardData): void {
        this.clipboard = data;
        this.copyToSystemClipboard(this.clipboardDataToText(data));
    }

    /**
     * 只更新内部剪贴板，不写系统剪贴板
     *
     * 供「点击即时值复制」这类已经自行写入系统剪贴板的场景使用：内部剪贴板一并
     * 更新后，表格自己的 Ctrl+V / 菜单粘贴才能落到格子里。
     * @param {ClipboardData} data - 剪贴板数据
     */
    public setInternalClipboard(data: ClipboardData): void {
        this.clipboard = data;
    }

    /**
     * 检查剪贴板是否有数据
     * @returns {boolean} 是否有数据
     */
    public hasClipboard(): boolean {
        return this.clipboard !== null;
    }

    /**
     * 清空剪贴板
     */
    public clearClipboard(): void {
        this.clipboard = null;
    }

    /**
     * 复制数据到系统剪贴板
     * @async
     * @param {string} text - 要复制的文本
     * @returns {Promise<boolean>} 是否复制成功
     */
    public async copyToSystemClipboard(text: string): Promise<boolean> {
        try {
            if (navigator.clipboard && navigator.clipboard.writeText) {
                await navigator.clipboard.writeText(text);
                return true;
            }
            console.warn('System clipboard is not available');
            return false;
        } catch (error) {
            console.error('Failed to copy to system clipboard:', error);
            return false;
        }
    }

    /**
     * 从系统剪贴板读取文本
     * @async
     * @returns {Promise<string>} 剪贴板文本内容
     */
    public async readFromSystemClipboard(): Promise<string> {
        try {
            if (navigator.clipboard && navigator.clipboard.readText) {
                return await navigator.clipboard.readText();
            }
            console.warn('System clipboard is not available');
            return '';
        } catch (error) {
            console.error('Failed to read from system clipboard:', error);
            return '';
        }
    }

    /**
     * 执行复制操作
     * @param {ClipboardData} data - 要复制的数据
     */
    public copy(data: ClipboardData): void {
        this.setClipboard(data);
    }

    /**
     * 执行剪切操作
     * @param {ClipboardData} data - 要剪切的数据
     */
    public cut(data: ClipboardData): void {
        this.setClipboard(data);
    }

    /**
     * 执行粘贴操作
     * @param {string} [targetCell] - 目标单元格（可选）
     * @returns {ClipboardData | null} 粘贴的数据
     */
    public paste(targetCell?: string): ClipboardData | null {
        if (!this.hasClipboard()) {
            return null;
        }

        return this.getClipboard();
    }

    /**
     * 将剪贴板数据转换为文本
     *
     * 单元格数据输出为 **TSV**（制表符分列、CRLF 换行）：
     * - 粘贴到 Excel / WPS 时按行列自动对齐到对应单元格；
     * - 粘贴到纯文本编辑器时每个单元格内容以 tab 分隔。
     *
     * @param {ClipboardData} data - 剪贴板数据
     * @returns {string} 转换后的文本
     */
    public clipboardDataToText(data: ClipboardData): string {
        if (data.type === 'cell') {
            return this.cellsToTSV(data.cells || []);
        } else if (data.type === 'char') {
            return data.chars?.map(char => char.char).join('') || '';
        }
        return '';
    }

    /**
     * 将单元格数组序列化为 TSV 文本
     *
     * 通过单元格名还原真实行列坐标，并以「最小~最大行列」为矩形补空字段：
     * 源范围内没有数据的空白单元格会输出为空字段，因此粘到 Excel 不会错列。
     * 行尾连续的空白字段会被裁掉，避免每行拖一串多余制表符。
     * @param {Cell[]} cells - 单元格数组（顺序不限）
     * @returns {string} TSV 文本（无有效单元格时为空串）
     */
    private cellsToTSV(cells: Cell[]): string {
        const textByPos = new Map<string, string>();
        let minCol = Infinity, maxCol = 0, minRow = Infinity, maxRow = 0;
        for (const cell of cells) {
            const pos = parseCellName(cell.cell);
            if (!pos) continue;
            minCol = Math.min(minCol, pos.col);
            maxCol = Math.max(maxCol, pos.col);
            minRow = Math.min(minRow, pos.row);
            maxRow = Math.max(maxRow, pos.row);
            textByPos.set(`${pos.col},${pos.row}`, cell.chars.map(char => char.char).join(''));
        }
        if (textByPos.size === 0) return '';

        const lines: string[] = [];
        for (let row = minRow; row <= maxRow; row++) {
            const fields: string[] = [];
            for (let col = minCol; col <= maxCol; col++) {
                fields.push(escapeTSVField(textByPos.get(`${col},${row}`) ?? ''));
            }
            while (fields.length > 0 && fields[fields.length - 1] === '') fields.pop();
            lines.push(fields.join('\t'));
        }
        return lines.join('\r\n');
    }

    // public textToClipboardData(text: string): ClipboardData {
    //     if (text.includes('\n')) {
    //         return {
    //             type: 'cell',
    //             cells: text.split('\n').map(line => new Cell(line)),
    //         };
    //     } else {
    //         return {
    //             type: 'char',
    //             chars: text.split('').map(char => new Char(char)),
    //         };
    //     }
    // }
}
