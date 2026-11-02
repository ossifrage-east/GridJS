/*  # 文档域 · Word 风格查找 / 替换对话框 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目目标: 把 Word「查找和替换」对话框的形态搬到文档编辑区：
    #           查找 / 替换两个页签、带冒号的字段标签、「更多 ≫」折叠的匹配选项、
    #           底部「查找下一处 / 替换 / 全部替换 / 关闭」按钮与状态行。
    # 本文件职责: 只做 UI 与交互编排 —— 取值、展示、按钮显隐、键盘回车；
    #             真正的命中扫描 / 选中 / 替换全在 `docSearch.ts`。
    # 与表格侧的关系: 两者外观一致（同样的页签条、标签宽度、更多折叠、状态行），
    #                 但数据载体不同：这里扫文档 DOM，表格侧扫单元格字符。
    # 依赖方向: 同目录 docSearch + common/Dialog + utils/dom；不 import 表格数据层。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { createButton, createDiv, createInput, createSpan } from '../../utils/dom';
import { Dialog, DialogButtonOptions } from '../common/dialog';
import {
    DOC_SEARCH_HIT_LIMIT,
    DocPosition,
    DocSearchHit,
    DocSearchOptions,
    findDocHits,
    getDocSelection,
    replaceAllDocHits,
    replaceDocHit,
    selectDocHit,
} from './docSearch';

/** 查找 / 替换的目标页签 */
export type DocSearchMode = 'find' | 'replace';

/** 「查找」页签文本 */
const TAB_FIND = '查找';
/** 「替换」页签文本 */
const TAB_REPLACE = '替换';
/** 标签列宽（像素；与表格侧保持同一视觉节奏） */
const LABEL_WIDTH = 82;
/** 对话框附加样式表 id（只注入一次） */
const STYLE_ELEMENT_ID = 'gs-doc-search-dialog-styles';

/**
 * 文档查找宿主
 *
 * 对话框只认这三样东西：块容器、有没有文档、把 DOM 结果同步回模型。
 * `DocWorkspace` 结构化满足它（见 `docWorkspace.ts` 的同名成员），
 * 因此装配层可以直接把工作区实例递进来，不必再包一层适配器。
 */
export interface DocSearchHost {
    /** 当前可查找的块容器（Markdown 源码视图下为 null） */
    readonly docBlocks: HTMLDivElement | null;
    /** 是否已打开文档 */
    readonly hasDoc: boolean;
    /** 把 DOM 上的替换结果同步回模型 */
    syncDocFromDom(label?: string): void;
}

export interface DocSearchDialogOptions {
    /** 查找宿主（文档工作区） */
    host: DocSearchHost;
}

/**
 * Word 风格的文档查找 / 替换对话框
 *
 * 同一实例长期持有：隐藏不销毁，页签、关键字、匹配选项与查找位置全部保留，
 * 再次打开即可继续（与表格侧 `SearchDialog` 的约定一致）。
 */
export class DocSearchDialog {
    /** 查找宿主 */
    private readonly host: DocSearchHost;
    /** 对话框实例 */
    private dialog: Dialog | null = null;
    /** 内部结构是否已构建 */
    private built: boolean = false;
    /** 当前页签 */
    private activeTab: DocSearchMode = 'find';

    /** 页签按钮 */
    private findTab!: HTMLButtonElement;
    private replaceTab!: HTMLButtonElement;
    /** 「查找内容」输入框 */
    private findInput!: HTMLInputElement;
    /** 「替换为」输入框 */
    private replaceInput!: HTMLInputElement;
    /** 「替换为」整行（查找页签下收起） */
    private replaceRow!: HTMLDivElement;
    /** 匹配选项（更多里） */
    private caseInput!: HTMLInputElement;
    private entireInput!: HTMLInputElement;
    private wildcardInput!: HTMLInputElement;
    /** 「更多 / 更少」折叠按钮与选项容器 */
    private moreButton!: HTMLButtonElement;
    private optionsWrap!: HTMLDivElement;
    /** 状态行 */
    private statusElement!: HTMLDivElement;

    /** 当前选中的命中（「替换」据此判断该替换还是该先找下一处） */
    private currentHit: DocSearchHit | null = null;
    /** 上一处命中的结束位置（「查找下一处」从这里往后走） */
    private lastPosition: DocPosition | null = null;

    constructor(options: DocSearchDialogOptions) {
        this.host = options.host;
    }

    /** 宿主（供共享实例判断是否需要重建） */
    public getHost(): DocSearchHost {
        return this.host;
    }

    /** 对话框当前是否打开 */
    public isVisible(): boolean {
        return this.dialog !== null && this.dialog.isVisible();
    }

    /** 已打开时只切页签并聚焦，不重置查找位置 */
    public refocus(mode?: DocSearchMode): void {
        if (!this.isVisible()) return;
        if (mode) this.switchTab(mode);
        this.focusFind();
    }

    /**
     * 打开对话框
     * @param mode - 目标页签（省略时沿用上次）
     */
    public open(mode?: DocSearchMode): void {
        this.ensureBuilt();
        const dialog = this.dialog;
        if (!dialog) return;

        // 复用实例时清除上次拖拽位移，避免第二次打开停在旧位置
        dialog.resetPosition();
        dialog.show();
        // Word 的查找对话框是**非模态**的：对话框开着也能继续在正文里点选 / 修改。
        // 遮罩改成点击穿透，按钮与输入框本身仍照常可点。
        dialog.setRangePickMode(true);

        // 与 Word 一致：正文里有选中文字时，用它预填「查找内容」，
        // 并把「从哪儿继续」的游标放到选区末尾 —— 下一次查找从选中处之后开始。
        const blocks = this.host.docBlocks;
        if (blocks) {
            const selected = getDocSelection(blocks);
            if (selected && selected.text) {
                this.findInput.value = selected.text;
                this.resetTracking();
                this.lastPosition = { blockIndex: selected.end.blockIndex, offset: selected.end.offset };
            }
        }

        this.switchTab(mode ?? this.activeTab);
        this.focusFind();
        this.setDefaultStatus();
    }

    /** 隐藏（不销毁，保留输入与位置） */
    public hide(): void {
        this.dialog?.hide();
    }

    /** 销毁（仅在宿主更换 / 页面卸载时） */
    public destroy(): void {
        const dialog = this.dialog;
        this.dialog = null;
        this.built = false;
        this.currentHit = null;
        this.lastPosition = null;
        if (dialog) dialog.destroy();
    }

    // ─────────────────────────── 构建 ───────────────────────────

    /** 构建内部结构（幂等） */
    private ensureBuilt(): void {
        if (this.built) return;
        this.built = true;
        this.injectStyles();

        const content = this.buildContent();
        const buttons: DialogButtonOptions[] = [
            { text: '查找下一处', variant: 'primary', close: false, onClick: () => this.findNext() },
            { text: '查找上一处', variant: 'secondary', close: false, onClick: () => this.findPrev() },
            { text: '替换', variant: 'warning', close: false, onClick: () => this.replaceCurrent() },
            { text: '全部替换', variant: 'danger', close: false, onClick: () => this.replaceAll() },
            { text: '关闭', variant: 'light' },
        ];
        this.dialog = new Dialog({
            title: '查找和替换',
            content,
            size: 'md',
            centered: true,
            buttons,
        });
    }

    /** 构建正文（页签条 + 字段 + 更多选项 + 状态行） */
    private buildContent(): HTMLElement {
        this.findTab = createButton({
            className: 'gs-doc-search-tab active',
            textContent: TAB_FIND,
            attributes: { type: 'button' },
        });
        this.replaceTab = createButton({
            className: 'gs-doc-search-tab',
            textContent: TAB_REPLACE,
            attributes: { type: 'button' },
        });
        this.findTab.addEventListener('click', () => this.switchTab('find'));
        this.replaceTab.addEventListener('click', () => this.switchTab('replace'));

        const tabBar = createDiv({
            className: 'gs-doc-search-tabs',
            children: [this.findTab, this.replaceTab],
        });

        const inputStyle = {
            flex: '1',
            padding: '0.25rem 0.375rem',
            fontSize: '0.875rem',
            lineHeight: '1.5',
            border: '1px solid #ced4da',
            borderRadius: '0.375rem',
            outline: 'none',
        };

        this.findInput = createInput({ type: 'text', style: inputStyle });
        this.findInput.addEventListener('input', () => this.resetTracking());
        this.findInput.addEventListener('keydown', (event) => {
            if (event.key !== 'Enter') return;
            event.preventDefault();
            if (event.shiftKey) this.findPrev();
            else this.findNext();
        });

        this.replaceInput = createInput({ type: 'text', style: inputStyle });
        this.replaceInput.addEventListener('keydown', (event) => {
            if (event.key !== 'Enter') return;
            event.preventDefault();
            this.replaceCurrent();
        });

        const findRow = this.buildRow('查找内容：', this.findInput);
        this.replaceRow = this.buildRow('替换为：', this.replaceInput);
        this.replaceRow.style.display = 'none';

        this.moreButton = createButton({
            className: 'gs-doc-search-more',
            textContent: '更多 ≫',
            attributes: { type: 'button' },
        });
        this.optionsWrap = this.buildOptions();
        this.optionsWrap.style.display = 'none';
        this.moreButton.addEventListener('click', () => this.toggleOptions());

        this.statusElement = createDiv({
            style: {
                minHeight: '1.25rem',
                marginTop: '0.25rem',
                fontSize: '0.8125rem',
                color: '#6c757d',
                userSelect: 'text',
            },
        });

        return createDiv({
            className: 'gs-doc-search',
            children: [tabBar, findRow, this.replaceRow, this.moreButton, this.optionsWrap, this.statusElement],
            style: { display: 'flex', flexDirection: 'column' },
        });
    }

    /** 构建「标签 + 控件」一行 */
    private buildRow(labelText: string, control: HTMLElement): HTMLDivElement {
        const label = createSpan({
            textContent: labelText,
            style: { flex: `0 0 ${LABEL_WIDTH}px`, color: '#495057' },
        });
        return createDiv({
            children: [label, control],
            style: { display: 'flex', alignItems: 'center', gap: '0.5rem', marginBottom: '0.5rem' },
        });
    }

    /** 构建匹配选项（默认折叠，点「更多」才展开） */
    private buildOptions(): HTMLDivElement {
        const makeCheckbox = (text: string, assign: (input: HTMLInputElement) => void, title?: string): HTMLElement => {
            const input = createInput({ type: 'checkbox' });
            input.addEventListener('change', () => this.resetTracking());
            assign(input);
            const wrapper = createDiv({
                children: [input, createSpan({ textContent: text })],
                style: { display: 'inline-flex', alignItems: 'center', gap: '0.25rem' },
            });
            if (title) wrapper.setAttribute('title', title);
            return wrapper;
        };
        const caseWrap = makeCheckbox('区分大小写', (input) => { this.caseInput = input; });
        const entireWrap = makeCheckbox('全字匹配', (input) => { this.entireInput = input; });
        const wildcardWrap = makeCheckbox('使用通配符', (input) => { this.wildcardInput = input; },
            '? 匹配单个字符，* 匹配任意长度');
        return createDiv({
            children: [caseWrap, entireWrap, wildcardWrap],
            style: {
                display: 'flex',
                flexWrap: 'wrap',
                gap: '0.75rem',
                marginLeft: `${LABEL_WIDTH + 8}px`,
                marginBottom: '0.5rem',
            },
        });
    }

    /** 仿 Word 的「更多 ≫ / 更少 ≪」折叠 */
    private toggleOptions(): void {
        const show = this.optionsWrap.style.display === 'none';
        this.optionsWrap.style.display = show ? 'flex' : 'none';
        this.moreButton.textContent = show ? '更少 ≪' : '更多 ≫';
    }

    /** 注入对话框样式（只注入一次） */
    private injectStyles(): void {
        if (document.getElementById(STYLE_ELEMENT_ID)) return;
        const styleSheet = document.createElement('style');
        styleSheet.id = STYLE_ELEMENT_ID;
        styleSheet.textContent = `
            /* 页签条：与表格侧查找对话框同一形态（下划线页签） */
            .gs-doc-search .gs-doc-search-tabs {
                display: flex;
                gap: 4px;
                padding: 0 0.25rem;
                margin-bottom: 0.5rem;
                border-bottom: 1px solid #dee2e6;
            }
            .gs-doc-search .gs-doc-search-tab {
                border: none;
                border-bottom: 2px solid transparent;
                background: none;
                padding: 0.25rem 0.75rem;
                font-size: 0.875rem;
                color: #495057;
                cursor: pointer;
            }
            .gs-doc-search .gs-doc-search-tab.active {
                color: #0d6efd;
                border-bottom-color: #0d6efd;
            }
            .gs-doc-search .gs-doc-search-tab:hover {
                border-bottom-color: #0d6efd;
            }
            /* 「更多 / 更少」：Word 里是一个靠左的文字链，不是按钮 */
            .gs-doc-search .gs-doc-search-more {
                align-self: flex-start;
                border: none;
                background: none;
                color: #0d6efd;
                padding: 0;
                margin-bottom: 0.5rem;
                font-size: 0.8125rem;
                cursor: pointer;
            }
            .gs-doc-search .gs-doc-search-more:hover {
                text-decoration: underline;
            }
        `;
        document.head.appendChild(styleSheet);
    }

    // ─────────────────────────── 页签 / 状态 ───────────────────────────

    /** 切换页签：显隐「替换为」与替换按钮 */
    private switchTab(mode: DocSearchMode): void {
        this.activeTab = mode;
        this.findTab.classList.toggle('active', mode === 'find');
        this.replaceTab.classList.toggle('active', mode === 'replace');
        this.replaceRow.style.display = mode === 'replace' ? 'flex' : 'none';
        const dialog = this.dialog;
        if (!dialog) return;
        for (const button of dialog.getFooterButtons()) {
            const text = button.textContent ?? '';
            if (text === '替换' || text === '全部替换') button.hidden = mode === 'find';
        }
    }

    /** 聚焦「查找内容」并全选已输入内容 */
    private focusFind(): void {
        this.findInput.focus();
        this.findInput.select();
    }

    /** 打开时的默认提示（没有文档 / 源码视图都要说清楚） */
    private setDefaultStatus(): void {
        if (!this.host.hasDoc) {
            this.setStatus('请先双击左侧的 .text 文件打开文档');
            return;
        }
        if (!this.host.docBlocks) {
            this.setStatus('Markdown 源码视图下暂不支持查找替换，请切回页面视图');
            return;
        }
        this.setStatus('输入查找内容后按 Enter 查找下一处');
    }

    /** 写状态行 */
    private setStatus(text: string): void {
        this.statusElement.textContent = text;
    }

    // ─────────────────────────── 查找 / 替换 ───────────────────────────

    /** 当前匹配选项 */
    private currentOptions(): DocSearchOptions {
        return {
            matchCase: this.caseInput.checked === true,
            matchEntire: this.entireInput.checked === true,
            useWildcard: this.wildcardInput.checked === true,
        };
    }

    /** 关键字 / 选项变化后清空「从哪儿继续」的游标 */
    private resetTracking(): void {
        this.currentHit = null;
        this.lastPosition = null;
    }

    /** 取可查找的块容器；不可用时写提示并返回 null */
    private requireBlocks(): HTMLDivElement | null {
        const blocks = this.host.docBlocks;
        if (!blocks) {
            this.setDefaultStatus();
            return null;
        }
        return blocks;
    }

    /** 取关键字；空串时写提示并返回 null */
    private requireKeyword(): string | null {
        const keyword = this.findInput.value;
        if (!keyword) {
            this.setStatus('请输入查找内容');
            return null;
        }
        return keyword;
    }

    /** 查找下一处（到末尾后回到第一处） */
    private findNext(): void {
        const blocks = this.requireBlocks();
        if (!blocks) return;
        const keyword = this.requireKeyword();
        if (!keyword) return;

        const hits = findDocHits(blocks, keyword, this.currentOptions(), DOC_SEARCH_HIT_LIMIT);
        if (hits.length === 0) {
            this.resetTracking();
            this.setStatus('未找到匹配内容');
            return;
        }

        let index = 0;
        let wrapped = false;
        if (this.lastPosition) {
            const found = this.indexOfHitAfter(hits, this.lastPosition);
            if (found >= 0) index = found;
            else wrapped = true;
        }
        const hit = hits[index];
        this.currentHit = hit;
        this.lastPosition = { blockIndex: hit.blockIndex, offset: hit.end };
        selectDocHit(hit, blocks);
        this.setStatus(this.progressStatus(index, hits.length, wrapped));
    }

    /** 查找上一处（到第一处后回到最后一处） */
    private findPrev(): void {
        const blocks = this.requireBlocks();
        if (!blocks) return;
        const keyword = this.requireKeyword();
        if (!keyword) return;

        const hits = findDocHits(blocks, keyword, this.currentOptions(), DOC_SEARCH_HIT_LIMIT);
        if (hits.length === 0) {
            this.resetTracking();
            this.setStatus('未找到匹配内容');
            return;
        }

        let index = hits.length - 1;
        let wrapped = false;
        const currentIndex = this.currentHit ? this.indexOfCurrent(hits, this.currentHit) : -1;
        if (currentIndex >= 0) {
            if (currentIndex > 0) index = currentIndex - 1;
            else wrapped = true;
        }
        const hit = hits[index];
        this.currentHit = hit;
        this.lastPosition = { blockIndex: hit.blockIndex, offset: hit.end };
        selectDocHit(hit, blocks);
        this.setStatus(this.progressStatus(index, hits.length, wrapped));
    }

    /**
     * 替换
     *
     * Word 的语义：当前选中处若就是命中项，替换它并自动找下一处；
     * 若没有选中命中，则先「查找下一处」（只选中、不替换）。
     */
    private replaceCurrent(): void {
        const blocks = this.requireBlocks();
        if (!blocks) return;
        const keyword = this.requireKeyword();
        if (!keyword) return;

        const hits = findDocHits(blocks, keyword, this.currentOptions(), DOC_SEARCH_HIT_LIMIT);
        if (hits.length === 0) {
            this.resetTracking();
            this.setStatus('未找到匹配内容');
            return;
        }

        const current = this.currentHit;
        const exactIndex = current ? this.indexOfCurrent(hits, current) : -1;
        if (exactIndex < 0) {
            // 没有选中命中：先定位到下一处，与 Word 的「替换」按钮首次点击一致
            this.findNext();
            return;
        }

        const hit = hits[exactIndex];
        const replacement = this.replaceInput.value;
        if (!replaceDocHit(hit, replacement, blocks)) {
            this.setStatus('替换失败，请重试');
            return;
        }
        this.host.syncDocFromDom('替换');
        this.currentHit = null;
        this.lastPosition = { blockIndex: hit.blockIndex, offset: hit.start + replacement.length };
        this.findNext();
    }

    /** 全部替换 */
    private replaceAll(): void {
        const blocks = this.requireBlocks();
        if (!blocks) return;
        const keyword = this.requireKeyword();
        if (!keyword) return;

        const count = replaceAllDocHits(blocks, keyword, this.replaceInput.value, this.currentOptions());
        if (count > 0) this.host.syncDocFromDom('全部替换');
        this.resetTracking();
        this.setStatus(count > 0 ? `已替换 ${count} 处` : '未找到匹配内容');
    }

    // ─────────────────────────── 辅助 ───────────────────────────

    /** 命中列表里是否有一处与当前命中相同（块元素 + 起点；防止视图重绘后误配旧命中） */
    private indexOfCurrent(hits: DocSearchHit[], current: DocSearchHit): number {
        for (let i = 0; i < hits.length; i++) {
            if (hits[i].blockEl === current.blockEl &&
                hits[i].blockIndex === current.blockIndex &&
                hits[i].start === current.start) return i;
        }
        return -1;
    }

    /**
     * 找到「从给定位置之后」的第一处命中
     * @returns 命中下标；都在它之前时为 -1
     */
    private indexOfHitAfter(hits: DocSearchHit[], position: DocPosition): number {
        for (let i = 0; i < hits.length; i++) {
            const hit = hits[i];
            if (hit.blockIndex > position.blockIndex) return i;
            if (hit.blockIndex === position.blockIndex && hit.start >= position.offset) return i;
        }
        return -1;
    }

    /** 命中进度提示（命中上限截断时如实说明，不谎报总数） */
    private progressStatus(index: number, total: number, wrapped: boolean): string {
        const prefix = wrapped ? '已到文档末尾，回到开头：' : '';
        if (total >= DOC_SEARCH_HIT_LIMIT) {
            return `${prefix}第 ${index + 1} 处（命中超过 ${DOC_SEARCH_HIT_LIMIT} 处，仅列出前 ${DOC_SEARCH_HIT_LIMIT} 处）`;
        }
        return `${prefix}第 ${index + 1} / ${total} 处`;
    }
}

// ─────────────────────────── 全局共享实例 ───────────────────────────

/**
 * 全局唯一的文档查找 / 替换对话框
 *
 * 与表格侧的共享实例分开：两个对话框各自持有不同宿主（文档工作区 / 表格数据），
 * 互斥显隐由 {@link Dialog} 的静态登记表负责。
 */
let sharedDocSearchDialog: DocSearchDialog | null = null;

/**
 * 打开文档查找 / 替换对话框
 * @param host - 查找宿主（`DocWorkspace`）
 * @param mode - 目标页签
 */
export function openDocFindReplace(host: DocSearchHost, mode?: DocSearchMode): void {
    if (sharedDocSearchDialog && sharedDocSearchDialog.getHost() !== host) {
        sharedDocSearchDialog.destroy();
        sharedDocSearchDialog = null;
    }
    if (!sharedDocSearchDialog) sharedDocSearchDialog = new DocSearchDialog({ host });
    if (sharedDocSearchDialog.isVisible()) {
        sharedDocSearchDialog.refocus(mode);
        return;
    }
    sharedDocSearchDialog.open(mode);
}

/** 隐藏文档查找 / 替换对话框（保留状态） */
export function hideDocFindReplace(): void {
    sharedDocSearchDialog?.hide();
}

/** 当前共享实例（供测试与调试） */
export function getSharedDocSearchDialog(): DocSearchDialog | null {
    return sharedDocSearchDialog;
}
