/*  # 文件打开器 TS 文件
    # 统一处理「打开文件」的入口：文件选择对话框、拖拽落点、解析装载与文件树登记
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { createDiv, createInput } from "../../utils/dom";
import { Cell } from "../dataArchitecture/cell";
import { Char } from "../dataArchitecture/char";
import { ColHeader, RowHeader } from "../dataArchitecture/header";
import { DataCollection } from "../dataArchitecture/dataCollection";
import { DEFAULT_CELL_HEIGHT, DEFAULT_CELL_WIDTH } from "../constant";
import { stripExtension } from "../../utils/fileName";
import { parseWorkbook } from "../../utils/parser/parserClient";
import type { ExtractedImage } from "../../utils/parser/sheetParser";
import { DocumentParser } from "../../utils/parser/documentParser";
import '../style/fileOpener.scss';

/**
 * 支持打开的文件扩展名
 */
export const SUPPORTED_EXTENSIONS = ['xlsx', 'xls', 'docx', 'doc', 'pptx', 'ppt'] as const;

/** 支持的文件扩展名联合类型 */
export type SupportedExtension = typeof SUPPORTED_EXTENSIONS[number];

/**
 * 解析结果：与项目数据格式一致的待装载数据
 * @interface OpenedDocument
 */
interface OpenedDocument {
    /** 单元格数组（项目 Cell 格式） */
    values: Cell[];
    /** 列头尺寸（无尺寸信息时为空数组，由 applyHeaders 补足） */
    colHeaders: ColHeader[];
    /** 行头尺寸 */
    rowHeaders: RowHeader[];
    /** 内嵌图片（docx/pptx 等文档类为空数组） */
    images: ExtractedImage[];
}

/**
 * 文件打开器配置选项接口
 * @interface FileOpenerOptions
 */
export interface FileOpenerOptions {
    /** 数据集合：解析结果写入此处 */
    data: DataCollection;
    /** 文件打开成功回调（参数为文件名），用于在文件树中登记节点 */
    onOpened?: (fileName: string) => void;
    /** 文件打开失败回调（参数为提示信息） */
    onError?: (message: string) => void;
}

/**
 * 文件打开器
 *
 * 提供三条打开路径，解析与装载逻辑共用：
 * 1. {@link pickAndOpen}：弹出系统文件选择对话框；
 * 2. {@link bindDropZone}：绑定拖放区域，拖入受支持的文件即自动打开；
 * 3. {@link openFile}：直接以 File 对象打开（供上述两者复用）。
 *
 * 装载规则：
 * - **xlsx / xls**：经 {@link SheetParser} 解析为单元格、列表头、行头，
 *   完整还原值、样式、合并区域与行列尺寸；
 * - **docx / doc / pptx / ppt**：经 {@link DocumentParser} 提取纯文本，
 *   按行写入 A 列（首个非空段落作为标题同步为工作表名）。
 *
 * 装载整体作为一次全量状态撤销操作记录（可用 Ctrl+Z 撤销打开动作），
 * 并落库持久化，保证刷新后仍是刚打开的文件内容。
 *
 * @class FileOpener
 * @example
 * const opener = new FileOpener({
 *     data,
 *     onOpened: (name) => fileTree.showActivedFile(name),
 *     onError: (msg) => data.messageBox.error(msg),
 * });
 * document.body.appendChild(opener.getFileInput());
 * opener.bindDropZone(document.querySelector('#worktop'));
 */
export class FileOpener {
    /** 数据集合 */
    private readonly data: DataCollection;
    /** 配置选项 */
    private readonly options: FileOpenerOptions;
    /** 拖放提示层 */
    private readonly dropHint: HTMLDivElement;
    /** 隐藏的文件选择输入框 */
    private readonly fileInput: HTMLInputElement;
    /** 拖放区域绑定的元素集合（destroy 时解绑） */
    private dropTargets: HTMLElement[] = [];
    /** 正在拖拽的进入计数（子元素冒泡会让 leave/enter 不配对，用计数配平） */
    private dragDepth = 0;

    /**
     * 构造函数
     * @param {FileOpenerOptions} options - 配置选项
     */
    constructor(options: FileOpenerOptions) {
        this.options = options;
        this.data = options.data;

        this.dropHint = createDiv({
            className: 'file-drop-hint',
            innerHTML: '<span>释放以打开文件</span>',
        });
        this.dropHint.hidden = true;

        this.fileInput = createInput({
            className: 'file-opener-input',
            type: 'file',
            attributes: {
                accept: SUPPORTED_EXTENSIONS.map((ext) => `.${ext}`).join(','),
                multiple: 'multiple',
            },
        });
        this.fileInput.hidden = true;
        this.fileInput.addEventListener('change', () => {
            const files = Array.from(this.fileInput.files ?? []);
            // 重置以便重复选择同一文件时仍能触发 change
            this.fileInput.value = '';
            files.forEach((file) => this.openFile(file));
        });
    }

    /**
     * 获取文件选择对话框输入框元素（需由外部挂载到 document 后才会响应点击）
     * @returns {HTMLInputElement} 文件选择输入框
     */
    public getFileInput(): HTMLInputElement {
        return this.fileInput;
    }

    /**
     * 弹出系统文件选择对话框
     */
    public pickAndOpen(): void {
        this.fileInput.click();
    }

    /**
     * 绑定拖放区域：拖入受支持的文件时自动打开
     *
     * 提示层挂载到第一个拖放目标内，随拖拽进入/离开显隐
     * @param {...HTMLElement} targets - 拖放目标元素（通常为主页面容器）
     */
    public bindDropZone(...targets: HTMLElement[]): void {
        this.dropTargets = targets;
        if (targets.length === 0) return;

        targets.forEach((target) => {
            target.addEventListener('dragenter', this.onDragEnter);
            target.addEventListener('dragover', this.onDragOver);
            target.addEventListener('dragleave', this.onDragLeave);
            target.addEventListener('drop', this.onDrop);
        });

        targets[0].appendChild(this.dropHint);
    }

    /**
     * 判断文件是否为受支持的类型
     * @param {File} file - 文件
     * @returns {boolean} 是否受支持
     */
    public static isSupported(file: File): boolean {
        const extension = file.name.toLowerCase().split('.').pop() ?? '';
        return (SUPPORTED_EXTENSIONS as readonly string[]).includes(extension);
    }

    /**
     * 打开单个文件（解析 → 装载 → 通知外部登记文件树）
     * @param {File} file - 待打开的文件
     */
    public openFile(file: File): void {
        if (!FileOpener.isSupported(file)) {
            this.options.onError?.(`不支持的文件格式：${file.name}`);
            return;
        }
        const extension = (file.name.toLowerCase().split('.').pop() ?? '') as SupportedExtension;

        // 解析 → 装载 → 通知外部登记文件树。
        //
        // 【为什么这里不再 `syncValuesToDB()`（性能约定，请勿加回）】
        // 原先装载完立刻整表序列化 + 写库，对大文件来说等于「刚读完一遍，马上再
        // 序列化一遍」：40 万格实测约 1.9 秒纯同步耗时、134 MB 临时字符串，
        // 叠加在打开动作上直接造成「打开即卡死」。
        // 现在改由装载结束时的 VALUES_CHANGED 触发脏标记 → 防抖 + 空闲时段落库
        // （见 DataCollection._scheduleAutoSave / _runIdleSave），持久化语义不变：
        // 导入进来的内容仍会被写回该文件名对应的存档键，只是不再阻塞打开动作。
        // 而「切换 / 改名 / 复制」这类必须先把旧表无条件固化再操作的流程，
        // 仍由 syncValuesToDB() 强制落库兜底。
        this.parse(file, extension)
            .then((document) => this.applyDocument(file.name, document))
            .then(() => this.options.onOpened?.(file.name))
            .catch((error: unknown) => {
                const reason = error instanceof Error ? error.message : String(error);
                this.options.onError?.(`打开「${file.name}」失败：${reason}`);
            });
    }

    /**
     * 按扩展名分派解析策略
     * @param {File} file - 待打开的文件
     * @param {SupportedExtension} extension - 文件扩展名
     * @returns {Promise<OpenedDocument>} 与项目格式一致的待装载数据
     * @private
     */
    private async parse(file: File, extension: SupportedExtension): Promise<OpenedDocument> {
        if (extension === 'xlsx' || extension === 'xls') {
            // Worker 优先：把解压 + XML 扫描移出主线程；不可用时内部自动回退主线程。
            //
            // `readImages: true` —— 位图渲染通道已经建好（见 `sheets/sheetImageLayer.ts`），
            // 图片不再是「读出来没人用」的浪费。**必须在 decode 阶段声明**：
            // media 条目在解压时就按目录跳过了，之后无从补救。
            const result = await parseWorkbook(file, { readImages: true });
            return {
                values: result.values,
                colHeaders: result.colHeaders,
                rowHeaders: result.rowHeaders,
                images: result.images,
            };
        }

        // 文档类：提取纯文本，逐行写入 A 列
        const result = await new DocumentParser().parse(file);
        const values = result.lines.map((line, index) => {
            const cell = new Cell();
            cell.cell = `A${index + 1}`;
            cell.chars = Array.from(line).map((char) => {
                const item = new Char();
                item.char = char;
                return item;
            });
            return cell;
        });
        return { values, colHeaders: [], rowHeaders: [], images: [] };
    }

    /**
     * 将解析结果装载进数据集合
     *
     * 整段装载包成一次全量状态撤销操作（Ctrl+Z 可撤销打开动作）；
     * 工作表名同步为去掉扩展名的文件名，作为 IndexedDB 的存储键。
     *
     * 从本机打开的文件**只有文件名**（浏览器不暴露所在目录），因此它就相当于
     * 文件树根目录下的文件 —— 存档键 = 去扩展名的路径，此时路径 ≡ 文件名，
     * 与文件树给出的路径天然一致（不同目录的同名文件才会带上目录前缀）。
     * @param {string} fileName - 文件名
     * @param {OpenedDocument} document - 待装载数据
     * @private
     */
    private applyDocument(fileName: string, document: OpenedDocument): void {
        const sheetName = stripExtension(fileName);
        this.data.runWithFullStateUndo(`打开 ${fileName}`, () => {
            this.data.activedSheetName = sheetName;
            this.applyHeaders(document.colHeaders, document.rowHeaders);
            // 图片先于 values 赋值：values setter 会广播 VALUES_CHANGED，
            // 画布在该事件里重绘 —— 先赋值图片才能让「首帧」就带上图
            this.data.images = document.images;
            this.data.values = document.values;
        });
        // runWithFullStateUndo 抑制了包裹期间的所有广播（避免在撤销栈里产生空操作），
        // 但「打开文件」是会话中途的装载：必须显式补发一次，
        // 否则主页面仍显示上一个文件的内容（画布不会重绘）
        this.data.notifyValuesChanged();
    }

    /**
     * 应用行列头尺寸
     *
     * 文档类无行列尺寸信息，按默认尺寸补足到可填充可视区的数量；
     * 工作簿沿用解析出的尺寸，数量不足可视区时同样按默认尺寸补足
     * @param {ColHeader[]} colHeaders - 解析出的列头（缺省为空数组）
     * @param {RowHeader[]} rowHeaders - 解析出的行头
     * @private
     */
    private applyHeaders(colHeaders: ColHeader[], rowHeaders: RowHeader[]): void {
        const { visibleView } = this.data;
        // 可视区至少需要的行列数（留白策略与 Sheets.ensureDefaultHeaders 保持一致）
        const minCols = Math.ceil(visibleView.sheetWidth / DEFAULT_CELL_WIDTH / this.data.zoom) + 2;
        const minRows = Math.ceil(visibleView.sheetHeight / DEFAULT_CELL_HEIGHT / this.data.zoom) + 2;

        // 解析器产出的表头已经是 {left,width}/{top,height} 形状，且其数组是一次性产物；
        // 仅在「数量不足、需要补足可视区」时才复制一份再扩容，避免对百万行表头做全量克隆。
        const cols: ColHeader[] = (colHeaders.length >= minCols)
            ? colHeaders
            : colHeaders.map((header) => ({ left: header.left, width: header.width }));
        const rows: RowHeader[] = (rowHeaders.length >= minRows)
            ? rowHeaders
            : rowHeaders.map((header) => ({ top: header.top, height: header.height }));

        let colLeft = cols.length > 0 ? cols[cols.length - 1].left + cols[cols.length - 1].width : 0;
        for (let c = cols.length; c < minCols; c++) {
            cols.push({ left: colLeft, width: DEFAULT_CELL_WIDTH });
            colLeft += DEFAULT_CELL_WIDTH;
        }
        let rowTop = rows.length > 0 ? rows[rows.length - 1].top + rows[rows.length - 1].height : 0;
        for (let r = rows.length; r < minRows; r++) {
            rows.push({ top: rowTop, height: DEFAULT_CELL_HEIGHT });
            rowTop += DEFAULT_CELL_HEIGHT;
        }

        // 直接赋值后显式重算总尺寸：否则 allColWidth/allRowHeight 保持旧值，
        // 滚动条 thumb 会停留在初始的「占满 track」状态导致无法拖动
        this.data.colHeaders.colHeaders = cols;
        this.data.colHeaders.recalcAllColWidth();
        this.data.colHeaders.offsetWidth = 0;
        this.data.rowHeaders.rowHeaders = rows;
        this.data.rowHeaders.recalcAllRowHeight();
        this.data.rowHeaders.offsetHeight = 0;
    }

    // —— 拖放事件处理 ——

    /**
     * 判断是否为文件拖拽（而非页面内元素拖拽）
     * @param {DragEvent} e - 拖拽事件
     * @returns {boolean} 是否文件拖拽
     * @private
     */
    private isFileDrag(e: DragEvent): boolean {
        return Array.from(e.dataTransfer?.types ?? []).includes('Files');
    }

    /**
     * 拖拽进入：显示提示层
     * @param {DragEvent} e - 拖拽事件
     * @private
     */
    private onDragEnter = (e: DragEvent): void => {
        if (!this.isFileDrag(e)) return;
        e.preventDefault();
        this.dragDepth++;
        this.dropHint.hidden = false;
    };

    /**
     * 拖拽悬停：阻止默认行为，否则 drop 不会触发
     * @param {DragEvent} e - 拖拽事件
     * @private
     */
    private onDragOver = (e: DragEvent): void => {
        if (!this.isFileDrag(e)) return;
        e.preventDefault();
        if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy';
    };

    /**
     * 拖拽离开：离开计数归零时隐藏提示层
     * @param {DragEvent} e - 拖拽事件
     * @private
     */
    private onDragLeave = (e: DragEvent): void => {
        if (!this.isFileDrag(e)) return;
        this.dragDepth = Math.max(0, this.dragDepth - 1);
        if (this.dragDepth === 0) this.dropHint.hidden = true;
    };

    /**
     * 放下：打开拖入的受支持文件
     * @param {DragEvent} e - 拖拽事件
     * @private
     */
    private onDrop = (e: DragEvent): void => {
        if (!this.isFileDrag(e)) return;
        e.preventDefault();
        this.dragDepth = 0;
        this.dropHint.hidden = true;

        const files = Array.from(e.dataTransfer?.files ?? []);
        const unsupported = files.filter((file) => !FileOpener.isSupported(file));
        const supported = files.filter((file) => FileOpener.isSupported(file));

        if (unsupported.length > 0) {
            this.options.onError?.(
                `不支持的文件格式：${unsupported.map((file) => file.name).join('、')}`
            );
        }
        supported.forEach((file) => this.openFile(file));
    };

    /**
     * 销毁组件：解绑拖放事件并移除 DOM
     */
    public destroy(): void {
        this.dropTargets.forEach((target) => {
            target.removeEventListener('dragenter', this.onDragEnter);
            target.removeEventListener('dragover', this.onDragOver);
            target.removeEventListener('dragleave', this.onDragLeave);
            target.removeEventListener('drop', this.onDrop);
        });
        this.dropTargets = [];
        this.dropHint.remove();
        this.fileInput.remove();
    }
}