import { Cell } from '../core/dataArchitecture/cell';
import { Char } from '../core/dataArchitecture/char';


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
     * @param {ClipboardData} data - 剪贴板数据
     * @returns {string} 转换后的文本
     */
    public clipboardDataToText(data: ClipboardData): string {
        if (data.type === 'cell') {
            let text = ''; 
            let lastCell = data.cells?.[0]?.cell;
            data.cells?.forEach(cell => {
                text += cell.chars.map(char => char.char).join('');
                text += cell.cell?.replace(/[^\d]/g, '') === lastCell.replace(/[^\d]/g, '') ? '\n' : ' ';
                lastCell = cell.cell;
            });
            return text;
        } else if (data.type === 'char') {
            return data.chars?.map(char => char.char).join('');
        }
        return '';
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
