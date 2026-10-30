/*  # 文档域 · Markdown 源码视图 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目目标: 让同一份文档既能所见即所得地编辑，也能直接改 Markdown 源码。
    # 本文件职责: 一个「行号 + 高亮层 + 输入层」的叠层编辑器：
    #             textarea 负责输入（保留原生撤销 / 输入法 / 选择行为），
    #             pre 负责着色（透明文字压在 textarea 上方，只显示颜色不拦截事件）。
    # 依赖方向: markdownCodec（互转）+ docStore（真源）。
    # 实现范式: 与 `assistant/pyEditor.ts` 的 Python 编辑器同构（同一套叠层技巧），
    #           只是 tokenizer 换成 Markdown 的。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { createDiv } from '../../utils/dom';
import { blocksToMarkdown, markdownToBlocks } from './converter/markdownCodec';
import { DocStore } from './docStore';
import '../style/document.scss';

/** 源码改动到模型落库的防抖（毫秒） */
const APPLY_DEBOUNCE = 260;

/**
 * Markdown 源码视图
 */
export class MarkdownView {
    /** 文档存储 */
    private readonly store: DocStore;

    /** 外层容器 */
    private readonly container: HTMLDivElement;

    /** 行号栏 */
    private readonly gutter: HTMLDivElement;

    /** 高亮层 */
    private readonly highlight: HTMLPreElement;

    /** 输入层 */
    private readonly input: HTMLTextAreaElement;

    /** 应用防抖定时器 */
    private applyTimer: number | null = null;

    /** 正在程序化写入（此时忽略 input，避免回环） */
    private applying: boolean = false;

    /**
     * 构造函数
     * @param options.store - 文档存储
     * @param options.parentElement - 挂载点
     */
    constructor(options: { store: DocStore; parentElement: HTMLElement }) {
        this.store = options.store;

        this.container = createDiv({ className: 'doc-markdown' });
        this.gutter = createDiv({ className: 'doc-markdown-gutter' });
        const stage = createDiv({ className: 'doc-markdown-stage' });
        this.highlight = document.createElement('pre');
        this.highlight.className = 'doc-markdown-highlight';
        this.input = document.createElement('textarea');
        this.input.className = 'doc-markdown-input';
        this.input.setAttribute('spellcheck', 'false');
        this.input.setAttribute('wrap', 'off');
        stage.appendChild(this.highlight);
        stage.appendChild(this.input);
        this.container.appendChild(this.gutter);
        this.container.appendChild(stage);
        options.parentElement.appendChild(this.container);

        this.bindEvents();
    }

    /** 取容器元素 */
    public getElements(): { container: HTMLDivElement } {
        return { container: this.container };
    }

    /** 显示 / 隐藏 */
    public setVisible(visible: boolean): void {
        this.container.hidden = !visible;
        if (visible) this.refresh();
    }

    /**
     * 从块树重建源码文本（切到本视图时调用）
     */
    public refresh(): void {
        const markdown = blocksToMarkdown(this.store.blocks);
        this.applying = true;
        this.input.value = markdown;
        this.applying = false;
        this.paint(markdown);
    }

    /**
     * 焦点放到源码编辑器
     */
    public focus(): void {
        this.input.focus();
    }

    /**
     * 绑定事件
     */
    private bindEvents(): void {
        this.input.addEventListener('input', () => {
            if (this.applying) return;
            const text = this.input.value;
            this.paint(text);
            this.scheduleApply(text);
        });

        // 滚动同步：高亮层与行号栏跟着输入层走，否则一滚动就错位
        this.input.addEventListener('scroll', () => {
            this.highlight.scrollTop = this.input.scrollTop;
            this.highlight.scrollLeft = this.input.scrollLeft;
            this.gutter.scrollTop = this.input.scrollTop;
        });

        // Tab 插入两个空格（默认行为是切焦点，源码编辑里很碍事）
        this.input.addEventListener('keydown', (event) => {
            if (event.key !== 'Tab') return;
            event.preventDefault();
            const start = this.input.selectionStart;
            const end = this.input.selectionEnd;
            const value = this.input.value;
            this.input.value = `${value.slice(0, start)}  ${value.slice(end)}`;
            this.input.selectionStart = start + 2;
            this.input.selectionEnd = start + 2;
            this.paint(this.input.value);
            this.scheduleApply(this.input.value);
        });
    }

    /**
     * 安排一次「源码 → 块树」的应用
     * @param text - 源码
     */
    private scheduleApply(text: string): void {
        if (this.applyTimer !== null) clearTimeout(this.applyTimer);
        this.applyTimer = window.setTimeout(() => {
            this.applyTimer = null;
            if (this.applying) return;
            const blocks = markdownToBlocks(text);
            this.applying = true;
            try {
                this.store.update(blocks, '编辑 Markdown', 'coalesce');
            } finally {
                this.applying = false;
            }
        }, APPLY_DEBOUNCE);
    }

    /**
     * 重新着色并更新行号
     * @param text - 源码
     */
    private paint(text: string): void {
        this.highlight.innerHTML = highlightMarkdown(text);
        const count = text.split('\n').length;
        let numbers = '';
        for (let i = 1; i <= count; i++) numbers += (i < 10 ? ` ${i}` : `${i}`) + '\n';
        this.gutter.textContent = numbers;
    }
}

/**
 * Markdown 语法着色
 *
 * **必须保证「去掉标签后与原文逐字相同」**：高亮层压在输入层下方，两者靠
 * 字符位置对齐；少一个字符就会整行错位（表现为文字重影）。
 * 因此这里只在原文上追加 `<span>`，绝不删除或替换任何字符。
 * @param text - 源码
 * @returns HTML
 */
export function highlightMarkdown(text: string): string {
    const lines = (text ?? '').split('\n');
    const out: string[] = [];
    for (let i = 0; i < lines.length; i++) {
        out.push(highlightLine(lines[i]));
    }
    return out.join('\n') + '\n';
}

/**
 * 着色单行
 * @param line - 行文本
 * @returns HTML
 */
function highlightLine(line: string): string {
    if (!line) return '';

    // 标题：`### 文本`
    const heading = /^(#{1,6})(\s+)(.*)$/.exec(line);
    if (heading) {
        return `<span class="md-marker">${escapeHtml(heading[1])}</span>`
            + escapeHtml(heading[2])
            + `<span class="md-heading">${highlightInline(heading[3])}</span>`;
    }

    // 引用：`> 文本`
    const quote = /^(\s*>+\s?)(.*)$/.exec(line);
    if (quote) {
        return `<span class="md-marker">${escapeHtml(quote[1])}</span>`
            + `<span class="md-quote">${highlightInline(quote[2])}</span>`;
    }

    // 列表：`- 文本` / `1. 文本`
    const list = /^(\s*)([-*+]|\d{1,9}[.)])(\s+)(.*)$/.exec(line);
    if (list) {
        return escapeHtml(list[1])
            + `<span class="md-marker">${escapeHtml(list[2])}</span>`
            + escapeHtml(list[3])
            + highlightInline(list[4]);
    }

    // 围栏
    if (/^\s*(`{3,}|~{3,})/.test(line)) {
        return `<span class="md-marker">${escapeHtml(line)}</span>`;
    }

    return highlightInline(line);
}

/**
 * 行内着色（粗体 / 斜体 / 行内代码 / 链接）
 * @param text - 行内文本
 * @returns HTML
 */
function highlightInline(text: string): string {
    if (!text) return '';
    let out = '';
    let index = 0;
    const pattern = /(`[^`]*`)|(\*\*[^*]+\*\*)|(\*[^*]+\*)|(\[[^\]]*\]\([^)]*\))/g;
    let match = pattern.exec(text);
    while (match) {
        out += escapeHtml(text.slice(index, match.index));
        const piece = match[0];
        if (piece.charAt(0) === '`') {
            out += `<span class="md-code">${escapeHtml(piece)}</span>`;
        } else if (piece.indexOf('**') === 0) {
            out += `<span class="md-strong">${escapeHtml(piece)}</span>`;
        } else if (piece.charAt(0) === '*') {
            out += `<span class="md-em">${escapeHtml(piece)}</span>`;
        } else {
            out += `<span class="md-link">${escapeHtml(piece)}</span>`;
        }
        index = match.index + piece.length;
        match = pattern.exec(text);
    }
    out += escapeHtml(text.slice(index));
    return out;
}

/**
 * HTML 转义（只转 `& < >`，保持其它字符原样）
 * @param text - 原文
 * @returns 转义结果
 */
function escapeHtml(text: string): string {
    return (text ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
