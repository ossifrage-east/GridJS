/*  # 文件批量任务面板（「另存为」导出清单：xlsx / docx 共用）
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { createDiv, createSpan, createInput, createButton, createElement } from "../../../utils/dom";
import { stripExtension } from "../../../utils/fileName";
import { createZipBlob, type XlsxTask } from "../../../utils/parser/xlsxWriter";
import { FILE_TREE_CLASS_NAMES as TREE } from "../../constant";
import { Dialog } from "../dialog";
import { MessageBox } from "../message";
import { formatFileSize } from "./fileTreeModel";

/**
 * 导出目标格式（决定对话框标题、产物后缀与可选的转换器）
 *
 * 两类文件走同一套清单 / 队列 / 下载流程，只有「源文件后缀 → 目标后缀」不同。
 */
export type TransferKind = 'xlsx' | 'docx';

/** 各目标格式对应的源文件后缀（用于「没有可导出文件」的提示文案） */
const SOURCE_EXTENSION: Record<TransferKind, string> = { xlsx: 'grid', docx: 'text' };

/**
 * 批量任务面板的外部依赖
 *
 * 面板只认识「路径 → 作业 / 大小」的抽象能力，真正的读档与转换由外部注入
 * （内容存在 IndexedDB，面板不接触存档）。
 * @interface FileTransferPanelOptions
 */
export interface FileTransferPanelOptions {
    /**
     * 为 `.grid` 文件创建 xlsx 转换作业
     * @param path - 文件路径（含扩展名）
     * @param name - 文件名（用于工作表名）
     */
    convertXlsx?: (path: string, name: string) => Promise<XlsxTask>;
    /**
     * 为 `.text` 文档创建 docx 转换作业
     * @param path - 文件路径（含扩展名）
     * @param name - 文件名
     */
    convertDocx?: (path: string, name: string) => Promise<XlsxTask>;
    /**
     * 读取文件内容的存档大小（字节），用于「文件大小」列
     * @param path - 文件路径（含扩展名）
     */
    fileSize?: (path: string) => Promise<number | null>;
}

/**
 * 清单行的转换状态
 *
 * `queued` 是单飞队列里的排队态：批量开始后，除第一个之外的行都在排队等待，
 * 只有当前转换被暂停 / 完成 / 删除，队列才会轮到下一行。
 */
type RowState = 'idle' | 'queued' | 'converting' | 'paused' | 'done';

/**
 * 清单的一行
 *
 * 行内元素引用随行创建一并持有，后续状态变化（进度、按钮可用性、下载可用性）
 * 只改这些引用，不重建 DOM。真实转换由 {@link FileTransferPanelOptions.convertXlsx}
 * 建出的 {@link XlsxTask} 承担，进度取自作业的真实值。
 * @interface Row
 */
interface Row {
    /** 文件路径（相对根节点，含扩展名；打包下载时据此还原目录层级，避免同名冲突） */
    path: string;
    /** 文件名（显示用） */
    name: string;
    /** 转换状态 */
    state: RowState;
    /** 转换进度（0–100，取自作业的真实进度） */
    progress: number;
    /**
     * 转换作业：首次出队时创建，暂停 / 排队期间一直保留，因此再次「开始」
     * 会从断点继续，而不是从头重跑。产物 Blob 也在作业里（`task.blob`）。
     */
    task: XlsxTask | null;
    /**
     * 驱动序号：每次出队自增
     *
     * 「暂停后立刻又按开始」会让同一个行先后存在两个驱动循环（旧的停在
     * `await` 上，新的已经出队），据此让旧循环在下一个检查点自行退出。
     */
    runToken: number;
    /** 行元素（删除时整体移除） */
    rowEl: HTMLTableRowElement;
    /** 行首多选框 */
    check: HTMLInputElement;
    /** 序号文本（删除行后需要重新编号） */
    indexEl: HTMLSpanElement;
    /** 进度单元格 */
    progressCell: HTMLTableCellElement;
    /** 开始按钮 */
    startBtn: HTMLButtonElement;
    /** 暂停按钮 */
    pauseBtn: HTMLButtonElement;
    /** 下载单元格 */
    downloadCell: HTMLTableCellElement;
}

/** 单次驱动转换的时间片（毫秒）：到点就交还主线程，保证暂停能及时生效、界面不卡 */
const SAVE_AS_SLICE_MS = 8;

/** 交还主线程用的共享消息通道（懒创建） */
let yieldPort: MessageChannel | null = null;
/** 同一时刻等待让出完成的回调 */
let yieldResolvers: Array<() => void> = [];

/**
 * 让出主线程一次（给浏览器渲染与用户点击的机会）
 *
 * 优先用 `scheduler.yield`（语义最准）；否则用 `MessageChannel`，
 * 避免 `setTimeout(0)` 被钳到 4ms —— 后者会让每个时间片都白等一截。
 * @returns {Promise<void>} 下一轮任务开始时兑现
 */
function yieldToHost(): Promise<void> {
    const scheduler = (globalThis as { scheduler?: { yield?: () => Promise<void> } }).scheduler;
    if (scheduler && typeof scheduler.yield === 'function') return scheduler.yield();
    if (!yieldPort) {
        yieldPort = new MessageChannel();
        yieldPort.port1.onmessage = () => {
            const batch = yieldResolvers;
            yieldResolvers = [];
            batch.forEach((resolve) => resolve());
        };
    }
    return new Promise<void>((resolve) => {
        yieldResolvers.push(resolve);
        yieldPort?.port2.postMessage(0);
    });
}

/**
 * 触发浏览器下载
 * @param {Blob} blob - 文件内容
 * @param {string} filename - 保存的文件名
 */
function triggerDownload(blob: Blob, filename: string): void {
    const url = URL.createObjectURL(blob);
    const anchor = createElement('a', {
        attributes: { href: url, download: filename },
    }) as HTMLAnchorElement;
    anchor.style.display = 'none';
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    // 给浏览器留出读取时间后再释放，避免大文件下载中断
    setTimeout(() => URL.revokeObjectURL(url), 5000);
}

/**
 * 取出错误对象的可读文本（转换失败提示用）
 * @param {unknown} error - 捕获到的异常
 * @returns {string} 可读的错误信息
 */
function errorMessage(error: unknown): string {
    if (error instanceof Error && error.message) return error.message;
    if (typeof error === 'string' && error) return error;
    return '未知错误';
}

/**
 * 批量任务面板（「另存为」导出清单对话框：xlsx / docx 共用）
 *
 * 正文为六列表格（转换在后台按时间片推进）：
 * 1. 多选框 + 序号（表头多选框全选）
 * 2. 文件名
 * 3. 文件大小（由 {@link FileTransferPanelOptions.fileSize} 异步补全，未提供时「—」）
 * 4. 进度（0–100%；排队中显示「排队中」，暂停时带「已暂停」后缀）
 * 5. 操作（每行开始 / 暂停 / 删除，仅图标；表头同样是这三枚按钮）
 * 6. 下载（表头为「打包下载」按钮；文件转换完成前整列不可用）
 *
 * 转换按**单飞队列**推进（一次只跑一个，其余显示「排队中」）；
 * 暂停会真正停下并保留断点，再次「开始」从断点续跑。
 */
export class FileTransferPanel {
    /** 外部依赖（转换 / 读大小） */
    private readonly options: FileTransferPanelOptions;
    /** 本次打开的目标格式（决定标题、产物后缀与选用的转换器） */
    private kind: TransferKind = 'xlsx';
    /** 对话框实例（复用同一实例，避免反复创建 DOM） */
    private dialog: Dialog | null = null;
    /** 清单行集合 */
    private rows: Row[] = [];
    /** 单飞队列：等待转换的行（FIFO，一次只跑一个） */
    private queue: Row[] = [];
    /** 当前正在转换的行（单飞：同一时刻至多一个） */
    private running: Row | null = null;
    /**
     * 批量暂停正在进行：暂停当前行时先不急着起下一行，
     * 否则「暂停第 1 行 → 起第 2 行 → 又立刻暂停第 2 行…」会连锁空转
     */
    private batchPausing = false;
    /** 表头的全选多选框 */
    private selectAll: HTMLInputElement | null = null;
    /** 表头「打包下载」按钮 */
    private downloadAll: HTMLButtonElement | null = null;
    /** 表头批量操作按钮：开始 / 暂停 / 删除勾选的行 */
    private headerStart: HTMLButtonElement | null = null;
    private headerPause: HTMLButtonElement | null = null;
    private headerRemove: HTMLButtonElement | null = null;

    /**
     * 构造函数
     * @param {FileTransferPanelOptions} [options={}] - 外部依赖
     */
    constructor(options: FileTransferPanelOptions = {}) {
        this.options = options;
    }

    /**
     * 打开导出清单（每次打开都从「待开始」重新来）
     * @param {Array<{ path: string; name: string }>} files - 待导出的文件（按树中可见顺序）
     * @param {TransferKind} [kind='xlsx'] - 目标格式（xlsx / docx）
     */
    public open(files: Array<{ path: string; name: string }>, kind: TransferKind = 'xlsx'): void {
        // 上一次的转换作业不跨次保留
        this.stop();
        this.kind = kind;
        this.rows = [];
        this.selectAll = null;
        this.downloadAll = null;
        this.headerStart = null;
        this.headerPause = null;
        this.headerRemove = null;
        if (!this.dialog) {
            this.dialog = new Dialog({
                title: `另存为${kind}`,
                size: 'lg',
                scrollable: true,
                buttons: [{ text: '关闭', variant: 'secondary' }],
            });
            // 关闭后停止转换并释放作业，避免对话框隐藏后仍在后台空转
            this.dialog.on('close', () => this.stop());
        }
        // 复用同一实例时标题随本次的目标格式切换
        this.dialog.setTitle(`另存为${kind}`);

        const body = createDiv();
        /** 文件大小单元格：异步补全时按行更新 */
        const sizeCells: Array<{ path: string; cell: HTMLTableCellElement }> = [];

        if (files.length === 0) {
            body.appendChild(createDiv({
                className: TREE.SAVE_AS_EMPTY,
                textContent: this.emptyHint(),
            }));
        } else {
            const table = createElement('table', { className: TREE.SAVE_AS_TABLE });
            table.appendChild(this.createTableHead());
            const tbody = createElement('tbody');
            files.forEach((file, index) => {
                tbody.appendChild(this.createTableRow(file, index + 1, sizeCells).rowEl);
            });
            table.appendChild(tbody);
            body.appendChild(table);
        }

        this.dialog.setContent(body);
        this.dialog.show();
        this.updateHeader();
        void this.fillFileSizes(sizeCells);
    }

    /**
     * 「清单为空」时的提示文案（按本次的目标格式指出应选中的源文件后缀）
     * @private
     * @returns {string} 提示文案
     */
    private emptyHint(): string {
        return `没有可导出的文件：请先在文件树中选中 .${SOURCE_EXTENSION[this.kind]} 文件`;
    }

    /**
     * 停止全部转换（对话框关闭或重新打开时调用）
     *
     * 清空排队，让当前行在下一个时间片停下（驱动循环看到「当前运行」已清空就退出），
     * 并丢弃所有作业 —— 关闭对话框后不再占用内存。
     */
    public stop(): void {
        this.queue = [];
        const running = this.running;
        this.running = null;
        if (running && running.state === 'converting') {
            running.state = 'paused';
            this.updateRow(running);
        }
        this.rows.forEach((row) => {
            row.task = null;
        });
    }

    /**
     * 销毁面板（释放对话框 DOM 与作业）
     */
    public destroy(): void {
        this.stop();
        this.rows = [];
        this.dialog?.destroy();
        this.dialog = null;
    }

    /**
     * 构建清单的表头
     * @private
     * @returns {HTMLTableSectionElement} thead 元素
     */
    private createTableHead(): HTMLTableSectionElement {
        const thead = createElement('thead') as HTMLTableSectionElement;
        const headRow = createElement('tr');

        // 第一列：全选多选框 + 序号
        const indexHead = createElement('th');
        const selectAll = createInput({
            type: 'checkbox',
            className: TREE.SAVE_AS_SELECT,
            attributes: { title: '全选' },
        });
        selectAll.addEventListener('change', () => {
            this.rows.forEach((row) => {
                row.check.checked = selectAll.checked;
            });
            this.updateHeader();
        });
        this.selectAll = selectAll;
        indexHead.append(selectAll, createSpan({ textContent: '序号' }));
        headRow.appendChild(indexHead);

        ['文件名', '文件大小', '进度'].forEach((text) => {
            headRow.appendChild(createElement('th', { textContent: text }));
        });
        // 操作列表头：同样三枚图标按钮，对勾选的行做批量操作
        const actionHead = createElement('th');
        const headerStart = this.createActionButton('icon-play3', '开始转换勾选的文件');
        const headerPause = this.createActionButton('icon-pause2', '暂停勾选的文件');
        const headerRemove = this.createActionButton('icon-bin', '从清单中删除勾选的文件');
        headerStart.addEventListener('click', () => this.startChecked());
        headerPause.addEventListener('click', () => this.pauseChecked());
        headerRemove.addEventListener('click', () => this.removeChecked());
        actionHead.append(headerStart, headerPause, headerRemove);
        headRow.appendChild(actionHead);
        this.headerStart = headerStart;
        this.headerPause = headerPause;
        this.headerRemove = headerRemove;

        // 下载列：表头是「打包下载」按钮，没有可下载文件时不可用
        const downloadHead = createElement('th');
        const downloadAll = createButton({
            className: TREE.SAVE_AS_DOWNLOAD_ALL,
            textContent: '打包下载',
            attributes: { type: 'button', title: '打包下载勾选的已转换文件' },
        });
        downloadAll.disabled = true;
        downloadAll.addEventListener('click', () => this.downloadChecked());
        this.downloadAll = downloadAll;
        downloadHead.appendChild(downloadAll);
        headRow.appendChild(downloadHead);

        thead.appendChild(headRow);
        return thead;
    }

    /**
     * 构建清单的一行（同时登记到 {@link rows}）
     * @private
     * @param file - 文件路径与名称
     * @param index - 序号（1-based）
     * @param sizeCells - 文件大小待补全列表（由调用方统一异步填充）
     * @returns {Row} 行状态与 DOM 引用
     */
    private createTableRow(
        file: { path: string; name: string },
        index: number,
        sizeCells: Array<{ path: string; cell: HTMLTableCellElement }>
    ): Row {
        const rowEl = createElement('tr') as HTMLTableRowElement;

        // 多选框 + 序号
        const indexCell = createElement('td') as HTMLTableCellElement;
        const check = createInput({
            type: 'checkbox',
            className: TREE.SAVE_AS_SELECT,
            attributes: { title: `选择 ${file.name}` },
        });
        check.addEventListener('change', () => this.updateHeader());
        const indexEl = createSpan({ className: TREE.SAVE_AS_INDEX, textContent: String(index) });
        indexCell.append(check, indexEl);
        rowEl.appendChild(indexCell);

        rowEl.appendChild(createElement('td', {
            textContent: file.name,
            attributes: { title: file.path },
        }));
        const sizeCell = createElement('td', { textContent: '—' }) as HTMLTableCellElement;
        rowEl.appendChild(sizeCell);
        sizeCells.push({ path: file.path, cell: sizeCell });

        const progressCell = createElement('td', { textContent: '0%' }) as HTMLTableCellElement;
        rowEl.appendChild(progressCell);

        // 操作列：开始 / 暂停 / 删除（仅图标，不写文字）
        const actionCell = createElement('td') as HTMLTableCellElement;
        const startBtn = this.createActionButton('icon-play3', `开始转换 ${file.name}`);
        const pauseBtn = this.createActionButton('icon-pause2', `暂停转换 ${file.name}`);
        const removeBtn = this.createActionButton('icon-bin', `从清单中删除 ${file.name}`);
        actionCell.append(startBtn, pauseBtn, removeBtn);
        rowEl.appendChild(actionCell);

        // 下载列：转换完成前不可用
        const downloadCell = createElement('td', { textContent: '下载' }) as HTMLTableCellElement;
        downloadCell.classList.add(TREE.SAVE_AS_DOWNLOAD, TREE.SAVE_AS_DISABLED);
        rowEl.appendChild(downloadCell);

        const row: Row = {
            path: file.path,
            name: file.name,
            state: 'idle',
            progress: 0,
            task: null,
            runToken: 0,
            rowEl,
            check,
            indexEl,
            progressCell,
            startBtn,
            pauseBtn,
            downloadCell,
        };
        startBtn.addEventListener('click', () => this.startRow(row));
        pauseBtn.addEventListener('click', () => this.pauseRow(row));
        removeBtn.addEventListener('click', () => this.removeRow(row));
        downloadCell.addEventListener('click', () => this.downloadRow(row));

        this.rows.push(row);
        this.updateRow(row);
        return row;
    }

    /**
     * 创建操作列的一枚纯图标按钮
     * @private
     * @param iconClass - 图标类名（见 icon.css）
     * @param title - 悬停提示（同时也是无障碍名称）
     * @returns {HTMLButtonElement} 按钮元素
     */
    private createActionButton(iconClass: string, title: string): HTMLButtonElement {
        const button = createButton({
            className: TREE.SAVE_AS_ACTION,
            attributes: { type: 'button', title },
        });
        button.appendChild(createElement('i', { className: iconClass }));
        return button;
    }

    /**
     * 开始（或继续）一行文件的转换：只入队，真正启动由 {@link pumpQueue} 负责
     * @private
     * @param {Row} row - 目标行
     */
    private startRow(row: Row): void {
        if (row.state === 'converting' || row.state === 'queued' || row.state === 'done') return;
        row.state = 'queued';
        this.queue.push(row);
        this.updateRow(row);
        this.updateHeader();
        void this.pumpQueue();
    }

    /**
     * 出队并运行队首（单飞：已有行在跑时直接返回）
     * @private
     */
    private async pumpQueue(): Promise<void> {
        if (this.running || this.queue.length === 0) return;
        const row = this.queue.shift();
        if (!row) return;
        this.running = row;
        row.state = 'converting';
        this.updateRow(row);
        this.updateHeader();
        await this.runRow(row);
    }

    /**
     * 用时间片驱动一行的转换，直到完成 / 被暂停 / 被删除
     *
     * 每片最多跑 {@link SAVE_AS_SLICE_MS} 毫秒就交还主线程：界面保持可响应，
     * 暂停也在片与片之间生效。停止推进时作业状态原样保留，下次「开始」续跑。
     * @private
     * @param {Row} row - 目标行
     */
    private async runRow(row: Row): Promise<void> {
        const token = ++row.runToken;
        let task = row.task;
        try {
            if (!task) {
                const convert = this.kind === 'docx' ? this.options.convertDocx : this.options.convertXlsx;
                if (!convert) throw new Error(`未配置 ${this.kind} 转换器`);
                const created = await convert(row.path, row.name);
                // 等待读档期间可能被「暂停 → 又开始」：交由新的驱动循环接管
                if (row.runToken !== token) return;
                row.task = created;
                task = created;
            }
            while (this.running === row
                && this.rows.includes(row)
                && row.state === 'converting'
                && row.runToken === token) {
                if (task.step(performance.now() + SAVE_AS_SLICE_MS)) {
                    row.progress = 100;
                    row.state = 'done';
                    this.updateRow(row);
                    break;
                }
                const value = Math.min(99, Math.round(task.progress));
                if (value > row.progress) {
                    row.progress = value;
                    this.updateRow(row);
                }
                await yieldToHost();
            }
        } catch (error) {
            // 行可能在等待读档期间被删除：此时不再改已失效的 DOM
            if (row.runToken === token && this.rows.includes(row)) {
                row.task = null;
                row.state = 'idle';
                row.progress = 0;
                this.updateRow(row);
                new MessageBox().show(`「${row.name}」转换失败：${errorMessage(error)}`, { type: 'error' });
            }
        } finally {
            // 只有仍占着「当前运行」的驱动才交棒（暂停 / 删除 / 新驱动接管都已让位）
            if (this.running === row && row.runToken === token) {
                this.running = null;
                this.updateHeader();
                void this.pumpQueue();
            }
        }
    }

    /**
     * 暂停一行文件的转换
     *
     * 排队中的行直接从队列摘下；正在跑的行在下一个时间片边界停下来，
     * 作业状态原样留在 `row.task` 上 —— 再次「开始」从断点继续。
     * @private
     * @param {Row} row - 目标行
     */
    private pauseRow(row: Row): void {
        if (row.state === 'queued') {
            this.queue = this.queue.filter((item) => item !== row);
            // 从未出队（还没有作业）时回到「待开始」，否则保持可续传的「已暂停」
            row.state = row.task ? 'paused' : 'idle';
            this.updateRow(row);
            this.updateHeader();
            return;
        }
        if (row.state !== 'converting') return;
        row.state = 'paused';
        this.updateRow(row);
        if (this.running === row) {
            this.running = null;
            this.updateHeader();
            if (!this.batchPausing) void this.pumpQueue();
        }
    }

    /**
     * 从清单中删除一行（同时把它从队列 / 当前运行位置摘除并重新编号）
     * @private
     * @param {Row} row - 目标行
     */
    private removeRow(row: Row): void {
        this.queue = this.queue.filter((item) => item !== row);
        // 正在跑的行：释放「当前运行」占位，驱动循环会在下一个时间片自行退出
        if (this.running === row) this.running = null;
        row.task = null;
        const tbody = row.rowEl.parentElement;
        row.rowEl.remove();
        this.rows = this.rows.filter((item) => item !== row);
        this.rows.forEach((item, index) => {
            item.indexEl.textContent = String(index + 1);
        });
        this.updateHeader();
        // 全部删空后换成与「无选中文件」一致的提示，避免只剩一个空表头
        if (this.rows.length === 0) {
            tbody?.parentElement?.remove();
            this.dialog?.getBody().appendChild(createDiv({
                className: TREE.SAVE_AS_EMPTY,
                textContent: this.emptyHint(),
            }));
        } else {
            void this.pumpQueue();
        }
    }

    /**
     * 批量开始：对勾选的行逐一启动转换（已完成的会被 {@link startRow} 跳过）
     * @private
     */
    private startChecked(): void {
        this.rows.filter((row) => row.check.checked).forEach((row) => this.startRow(row));
    }

    /**
     * 批量暂停：对勾选且正在转换 / 排队中的行暂停（其余状态忽略）
     *
     * 先暂停完所有目标行，再统一起下一行，避免暂停过程中反复启动新行。
     * @private
     */
    private pauseChecked(): void {
        const targets = this.rows.filter((row) => row.check.checked);
        this.batchPausing = true;
        try {
            targets.forEach((row) => this.pauseRow(row));
        } finally {
            this.batchPausing = false;
        }
        void this.pumpQueue();
    }

    /**
     * 批量删除：把勾选的行从清单中移除（行内删除逻辑与单行一致）
     * @private
     */
    private removeChecked(): void {
        this.rows.filter((row) => row.check.checked).forEach((row) => this.removeRow(row));
    }

    /**
     * 按行状态刷新进度文本、按钮可用性与下载可用性
     * @private
     * @param {Row} row - 目标行
     */
    private updateRow(row: Row): void {
        if (row.state === 'done') row.progressCell.textContent = '已完成';
        else if (row.state === 'queued') row.progressCell.textContent = '排队中';
        else row.progressCell.textContent = `${row.progress}%${row.state === 'paused' ? '（已暂停）' : ''}`;
        row.startBtn.disabled = row.state === 'converting' || row.state === 'queued' || row.state === 'done';
        row.pauseBtn.disabled = !(row.state === 'converting' || row.state === 'queued');
        const done = row.state === 'done';
        row.rowEl.classList.toggle(TREE.SAVE_AS_DONE, done);
        row.downloadCell.classList.toggle(TREE.SAVE_AS_DISABLED, !done);
        row.downloadCell.title = done ? `下载 ${row.name}` : '转换完成后可下载';
    }

    /**
     * 刷新表头：全选框的三态，以及「打包下载」的可用性
     *
     * 下载列整列（含表头）都遵循「转换完成才可用」：表头按钮要求
     * 「勾选的这些文件全部已完成」；对应行未完成时该行的「下载」同样置灰。
     * @private
     */
    private updateHeader(): void {
        const rows = this.rows;
        const checked = rows.filter((row) => row.check.checked);
        if (this.selectAll) {
            this.selectAll.checked = rows.length > 0 && checked.length === rows.length;
            this.selectAll.indeterminate = checked.length > 0 && checked.length < rows.length;
            this.selectAll.disabled = rows.length === 0;
        }
        if (this.downloadAll) {
            this.downloadAll.disabled =
                checked.length === 0 || checked.some((row) => row.state !== 'done');
        }
        // 表头的批量操作同样以「有勾选行」为前提
        const noChecked = checked.length === 0;
        [this.headerStart, this.headerPause, this.headerRemove].forEach((button) => {
            if (button) button.disabled = noChecked;
        });
    }

    /**
     * 下载单行「转换后」的产物（转换完成且拿到 Blob 才生效）
     * @private
     * @param {Row} row - 目标行
     */
    private downloadRow(row: Row): void {
        const blob = row.task?.blob;
        if (row.state !== 'done' || !blob) return;
        triggerDownload(blob, `${stripExtension(row.name)}.${this.kind}`);
    }

    /**
     * 打包下载：把勾选且已完成的产物打成 zip
     *
     * ZIP 内的条目名**保留文件在树中的目录层级**（`归档/销售数据.grid` →
     * `归档/销售数据.xlsx`），因此不同文件夹下的同名文件互不冲突。
     * 同一目录下万一仍有重名，再按 `名称(2).xlsx` 递增去重；
     * 所有文件都以「存储」方式写入，不重新压缩产物自身。
     * @private
     */
    private async downloadChecked(): Promise<void> {
        const targets = this.rows.filter((row) => row.check.checked && row.state === 'done' && row.task?.blob);
        if (targets.length === 0) return;
        try {
            const used = new Set<string>();
            const entries: Array<{ name: string; data: Uint8Array }> = [];
            for (const row of targets) {
                const blob = row.task?.blob;
                if (!blob) continue;
                // 相对路径（含扩展名）→ zip 条目名：去掉源后缀、补目标后缀，分隔符统一为 `/`。
                // 只在最后一个 `/` 之后去扩展名，避免目录名里的点被误当作扩展名分隔符。
                const normalized = (row.path || row.name).replace(/\\/g, '/')
                    .split('/')
                    .filter((segment) => segment !== '')
                    .join('/');
                const slash = normalized.lastIndexOf('/');
                const dir = slash >= 0 ? normalized.slice(0, slash + 1) : '';
                const base = stripExtension(slash >= 0 ? normalized.slice(slash + 1) : normalized);
                let entryName = `${dir}${base}.${this.kind}`;
                let index = 2;
                while (used.has(entryName.toLowerCase())) {
                    entryName = `${dir}${base}(${index++}).${this.kind}`;
                }
                used.add(entryName.toLowerCase());
                entries.push({ name: entryName, data: new Uint8Array(await blob.arrayBuffer()) });
            }
            triggerDownload(createZipBlob(entries), `${this.kind}打包下载.zip`);
        } catch (error) {
            new MessageBox().show(`打包下载失败：${errorMessage(error)}`, { type: 'error' });
        }
    }

    /**
     * 异步补全清单中的「文件大小」列
     *
     * 逐行读取并写回，读取失败或未提供 {@link FileTransferPanelOptions.fileSize} 时保持「—」；
     * 对话框被重新打开后旧单元格已脱离文档，随即停止更新。
     * @private
     * @param {Array<{ path: string; cell: HTMLTableCellElement }>} rows - 待补全的行
     */
    private async fillFileSizes(rows: Array<{ path: string; cell: HTMLTableCellElement }>): Promise<void> {
        const getSize = this.options.fileSize;
        if (!getSize) return;
        for (const row of rows) {
            let bytes: number | null = null;
            try {
                bytes = await getSize(row.path);
            } catch {
                bytes = null;
            }
            if (!row.cell.isConnected) return;
            row.cell.textContent = bytes === null ? '—' : formatFileSize(bytes);
        }
    }
}
