/*  # 文档域 · 文档工作区 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目目标: 把「文档视图 / Markdown 源码视图 / 文档存储」装配成一个
    #           与表格工作区并列、可整体显隐的工作区。
    # 本文件职责: 装配与生命周期 —— 建 DOM、挂两个视图、处理快捷键、
    #             负责打开 / 关闭 / 导出。**不实现编辑逻辑**（那是 DocView 的事），
    #             也**不实现存储**（那是 DocStore 的事），**也不再自建工具条**
    #             （段落样式 / 列表 / 缩进 / 撤销 / Markdown 已全部并入顶部工具栏，
    #              见 core/toolbar/profile.ts）。
    # 依赖方向: 同目录 docView / markdownView / docStore / docEditHost / 转换器
    #           + common（Split / Slider / 菜单实例类型）
    #           + utils/dom、utils/debounce。
    #           不 import 表格数据层，也不 import toolbar ——
    #           工作区不认识外壳，外壳通过命令总线找它。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { createDiv, createSpan } from '../../utils/dom';
import { debounce } from '../../utils/debounce';
import { CommandBus } from '../contract/commandBus';
import { Slider } from '../common/slider';
import { Split } from '../common/split';
import { DOC_CHANGED, DOC_RESET } from './docStore';
import { DOC_PAPER_SIZES, DocPageSetup, blockText } from './blockTree';
import { snapshotToDocx } from './converter/docxCodec';
import { blocksToMarkdown } from './converter/markdownCodec';
import { DocEditHost } from './docEditHost';
import { DocPrinter } from './docPrinter';
import { DocScrollbar } from './docScroller';
import { DocStore } from './docStore';
import { DocView } from './docView';
import { MarkdownView } from './markdownView';

/**
 * 缩放滑块的下限 / 上限（百分比）
 *
 * 与 {@link DocView.setZoom} 的 0.5–4 区间同源：写死两处就会出现
 * 「滑块能拖到 400% 但视图只吃到 200%」这种各说各话的情形。
 */
const ZOOM_PERCENT_MIN = 50;
const ZOOM_PERCENT_MAX = 400;

/** Ctrl+滚轮每格改多少缩放（与表格滚动条的手感对齐：一格 10%） */
const ZOOM_WHEEL_STEP = 0.1;

/**
 * 文档工作区
 */
export class DocWorkspace {
    /** 存储（真源） */
    private readonly store: DocStore;

    /** 视图 */
    private readonly view: DocView;

    /** Markdown 源码视图 */
    private readonly markdownView: MarkdownView;

    /** 编辑宿主 */
    private readonly host: DocEditHost;

    /** 外层容器 */
    private readonly container: HTMLDivElement;

    /** 正文区（放两边视图；位置与表格的画布区一致：右侧 / 下侧各让出 26px 工具线） */
    private readonly body: HTMLDivElement;

    /**
     * 右侧工具线（`Split` 容器，与表格的 `#rightTools` 同构）
     *
     * 纵向分割：前侧（上半）留空，后侧（下半）挂纵向滚动条。
     * 用**类名**而不是 id —— 两个工作区同时存在于 DOM 里（一个是 hidden），
     * 用 id 会撞成重复 id。
     */
    private readonly rightTools: HTMLDivElement;

    /** 右侧工具线的分割组件（纵向） */
    private readonly rightToolsSplit: Split;

    /** 底部工具线（`Split` 容器，与表格的 `#bottomTools` 同构：左半状态区、右半横向滚动条） */
    private readonly bottomTools: HTMLDivElement;

    /** 底部工具线的分割组件（横向） */
    private readonly bottomToolsSplit: Split;

    /** 横向滚动条 */
    private hScroller!: DocScrollbar;

    /** 纵向滚动条 */
    private vScroller!: DocScrollbar;

    /** 无文档时的占位 */
    private readonly empty: HTMLDivElement;

    /** 缩放滑块（垂直细线样式，挂在右侧工具线上半，与表格右侧的缩放滑块同形） */
    private zoomSlider!: Slider;

    /**
     * 导航面板（右侧工具线上半，由外壳注入）
     *
     * 与表格共用同一个 `Navigator` 实例：两个工作区不会同时显示，切换时外壳
     * 把这块面板在两个右侧工具线之间搬动，导航内容区（`containerSplit.splitNext`）
     * 始终不动 —— 这样既不会出现两个同 id 的面板，也不会出现两份同键的选中状态。
     */
    private navPanel: HTMLElement | null = null;

    /** 打印预览（与表格共用外壳；按需创建，复用同一个实例） */
    private printer: DocPrinter | null = null;

    /** 命令总线（视图状态变化时借它广播给工具栏） */
    private readonly bus: CommandBus;

    /** 当前视图模式 */
    private mode: 'rich' | 'markdown' = 'rich';

    /** 是否已装配快捷键 */
    private shortcutsBound: boolean = false;

    /** 底部即时计算栏里显示字符数的值元素 */
    private charCountValue!: HTMLSpanElement;

    /** 视图状态（缩放 / 滚动位置）落库的防抖定时器 */
    private viewStateTimer: number | null = null;

    /**
     * 构造函数
     * @param options.parentElement - 挂载点（主页面 splitPrev）
     * @param options.bus - 命令总线
     */
    constructor(options: { parentElement: HTMLElement; bus: CommandBus }) {
        this.bus = options.bus;
        this.store = new DocStore();

        this.container = createDiv({ className: 'doc-workspace' });
        this.container.hidden = true;

        // 布局严格照搬表格工作区（`.sheet-workspace` + `#rightTools` / `#bottomTools`）：
        // 正文与两条工具线都是**工作区内的绝对定位子元素**，工具线贴在右边 / 底边，
        // 滚动条挂在工具线里（而不是正文内部），才不会随内容一起滚走、也不会被正文挤出可视区
        this.body = createDiv({ className: 'doc-body' });
        this.container.appendChild(this.body);

        // 工具线用 Split 生成（与表格的 #bottomTools / #rightTools 同构）：
        // 前侧放状态内容，后侧（split-next）挂滚动条，中间是可按比例拖动的分割条。
        // Split 构造时会量容器尺寸，量不到会在 rAF 里反复重试，因此先让容器
        // 「参与布局但不绘制」（visibility:hidden）—— 同一个任务里随后就收起，不会闪一下
        // （容器必须先挂进 DOM，否则量出来的永远是 0）
        options.parentElement.appendChild(this.container);
        this.container.hidden = false;
        this.container.style.visibility = 'hidden';
        this.container.style.pointerEvents = 'none';

        this.bottomToolsSplit = new Split({
            parentElement: this.container,
            id: 'doc-bottomTools',
            storageKey: 'doc-bottomTools',
        });
        this.bottomTools = this.bottomToolsSplit.getElements().container;
        this.bottomTools.classList.add('bottomTools');

        this.rightToolsSplit = new Split({
            parentElement: this.container,
            id: 'doc-rightTools',
            horizontal: false,
            storageKey: 'doc-rightTools',
        });
        this.rightTools = this.rightToolsSplit.getElements().container;
        this.rightTools.classList.add('rightTools');

        this.container.style.pointerEvents = '';
        this.container.style.visibility = '';
        this.container.hidden = true;

        this.view = new DocView({ store: this.store, parentElement: this.body });
        // 标尺挂在**正文**（而不是滚动区）上：滚动区里的元素会随纸一起滚走，
        // 滚一屏之后刻度就与纸边错开了。标尺与滚动区同为 body 的直接子元素，
        // 由 DocView 监听 scroll 反向平移刻度条（见 DocView.mountRulers / syncRulers）。
        this.view.mountRulers(this.body);
        this.markdownView = new MarkdownView({ store: this.store, parentElement: this.body });
        this.markdownView.getElements().container.hidden = true;
        this.host = new DocEditHost(this.view, this.store);

        this.empty = createDiv({ className: 'doc-empty', innerHTML: '<span>双击左侧 .text 文件开始编辑</span>' });
        this.body.appendChild(this.empty);

        // 滚动条：父元素分别是底部 / 右侧工具线的**后侧面板**（split-next），
        // 与表格 hScroller / vScroller 的挂载点完全相同。
        // 它们只认「一个能滚的容器」，读 scrollLeft/Top、写回去，不碰纸张与块树
        const scroller = this.view.getElements().scroller;
        this.hScroller = new DocScrollbar({
            parentElement: this.bottomToolsSplit.getElements().splitNext,
            target: scroller,
            direction: 'horizontal',
        });
        this.vScroller = new DocScrollbar({
            parentElement: this.rightToolsSplit.getElements().splitNext,
            target: scroller,
            direction: 'vertical',
        });
        // 滚动位置也要落库（与表格把偏移量写进存档同一目的）：滚动停下后记一次
        scroller.addEventListener('scroll', () => this.scheduleViewState(), { passive: true });

        // 底部工具线的**前侧面板**：与表格同形的即时计算栏。
        // 默认统计正文全部字符；有选区时只统计选区内的字符
        // （文案 / 外观沿用表格的 .selection-calc 样式）
        this.charCountValue = createSpan({ className: 'selection-calc__value', textContent: '0' });
        this.bottomToolsSplit.getElements().splitPrev.appendChild(createDiv({
            className: 'selection-calc',
            children: [createSpan({
                className: 'selection-calc__item',
                children: [
                    createSpan({ className: 'selection-calc__label', textContent: '字符' }),
                    this.charCountValue,
                ],
            })],
        }));
        // 内容变化 / 撤销重做 / 选区变化都会影响计数；做 80ms 防抖，避免连续输入时每键都全文拼接
        const scheduleCharCount = debounce(() => this.refreshCharCount(), 80);
        this.store.on(DOC_CHANGED, scheduleCharCount);
        this.store.on(DOC_RESET, scheduleCharCount);
        document.addEventListener('selectionchange', scheduleCharCount);
        this.refreshCharCount();

        // 右侧工具线的**前侧面板**：上半是导航面板（由外壳在切换工作区时注入，
        // 见 attachNavPanel），下半是垂直细线缩放滑块（同表格的缩放控件）
        this.initZoomSlider(this.rightToolsSplit.getElements().splitPrev);

        this.bindShortcuts();
        this.refresh();
    }

    /** 编辑宿主（供路由层绑定到命令总线） */
    public get editHost(): DocEditHost {
        return this.host;
    }

    /** 文档存储（供装配层做内容桥） */
    public get docStore(): DocStore {
        return this.store;
    }

    /**
     * 刷新底部即时计算栏里的字符数
     *
     * 默认统计正文全部字符；文档内有选区时只统计选区内的字符（均为「不计空白」
     * 口径）。没有文档时显示 0。
     */
    private refreshCharCount(): void {
        const selected = this.selectedText();
        const source = selected ?? this.store.blocks.map((block) => blockText(block)).join('\n');
        this.charCountValue.textContent = String(source.replace(/\s/g, '').length);
    }

    /**
     * 取文档视图内的选中文本
     *
     * 只认落在滚动区内部（纸张 / 正文）的选区：工具栏、导航等处的文字选中
     * 不应把计数切走。没有选区或选区为空时返回 null，调用方退回全文计数。
     * @returns 选中的纯文本；无有效选区时为 null
     */
    private selectedText(): string | null {
        const selection = window.getSelection();
        if (!selection || selection.rangeCount === 0 || selection.isCollapsed) return null;
        const { anchorNode, focusNode } = selection;
        if (!anchorNode || !focusNode) return null;
        const scroller = this.view.getElements().scroller;
        if (!scroller.contains(anchorNode) || !scroller.contains(focusNode)) return null;
        const text = selection.toString();
        return text ? text : null;
    }

    /**
     * 排期一次视图状态落库（缩放 / 滚动位置）
     *
     * 滚动与拖缩放滑块都会高频触发，因此做 300ms 防抖：停下之后才写一次，
     * 避免每帧都改文档元信息（那会连带触发整份文档的落库）。
     */
    private scheduleViewState(): void {
        if (this.viewStateTimer !== null) window.clearTimeout(this.viewStateTimer);
        this.viewStateTimer = window.setTimeout(() => {
            this.viewStateTimer = null;
            this.captureViewState();
        }, 300);
    }

    /**
     * 立即把待落库的视图状态写进文档元信息（取消未触发的防抖）
     *
     * 切换文档 / 隐藏 / 落库前调用：防抖 300ms 期间用户可能已经切走，
     * 不补这一次就会丢掉最后一次滚动或缩放。
     */
    private flushViewState(): void {
        if (this.viewStateTimer !== null) {
            window.clearTimeout(this.viewStateTimer);
            this.viewStateTimer = null;
        }
        this.captureViewState();
    }

    /**
     * 把当前缩放与滚动位置写进文档元信息
     *
     * 与表格把 `zoom` / 偏移量存进存档同一目的：重新打开这份文档时恢复「上次
     * 看到哪里」。没有打开文档时什么都不做。
     */
    private captureViewState(): void {
        if (!this.store.hasDoc) return;
        const scroller = this.view.getElements().scroller;
        this.store.updateMeta({
            viewState: {
                zoom: this.view.zoom,
                scrollLeft: Math.max(0, Math.round(scroller.scrollLeft)),
                scrollTop: Math.max(0, Math.round(scroller.scrollTop)),
            },
        });
    }

    /** 取容器元素 */
    public getElements(): { container: HTMLDivElement } {
        return { container: this.container };
    }

    /**
     * 挂载导航面板（由外壳注入；与表格共用同一个导航实例）
     *
     * 面板要放在缩放滑块**之前**：DOM 顺序即视觉顺序（上半导航、下半滑块）。
     * @param panel - 导航面板元素（`.nav-panel`）
     */
    public attachNavPanel(panel: HTMLElement): void {
        this.navPanel = panel;
        const prev = this.rightToolsSplit.getElements().splitPrev;
        const slider = this.zoomSlider.getElements().container;
        if (slider.parentElement === prev) prev.insertBefore(panel, slider);
        else prev.appendChild(panel);
    }

    /**
     * 摘除导航面板（切回表格时把同一个面板还给表格的右侧工具线）
     *
     * 只在本工作区持有它时才移除，避免外壳的重复切换把面板从表格里误摘走。
     */
    public detachNavPanel(): void {
        const prev = this.rightToolsSplit.getElements().splitPrev;
        if (this.navPanel && this.navPanel.parentElement === prev) prev.removeChild(this.navPanel);
    }

    /**
     * 设置工作区高度
     *
     * 与 `.sheet-workspace` 的做法完全相同：工作区在工具栏之下，若直接用 `height: 100%`
     * 会多占工具栏那一截、把底部工具线顶出可视区，因此由外壳按
     * 「分隔容器高度 − 工具栏高度」写成**内联像素值**。
     * @param height - 可用高度（px）
     */
    public setHostHeight(height: number): void {
        this.container.style.height = `${Math.max(0, Math.round(height))}px`;
    }

    /**
     * 打开一份文档
     * @param docKey - 存档键
     * @param title - 标题
     */
    public async open(docKey: string, title: string): Promise<void> {
        // 切文档前先把上一份的缩放 / 滚动位置写回去（防抖可能还没触发）
        this.flushViewState();
        await this.store.load(docKey, title);
        this.show();
        this.applyMode(this.store.meta.view === 'markdown' ? 'markdown' : 'rich');
        this.refresh();

        // 恢复「上次看到哪里」：有存过的缩放就用它，没有（老存档 / 新文件）
        // 才退回「适应页宽」这个默认
        const viewState = this.store.meta.viewState ?? {};
        const scroller = this.view.getElements().scroller;
        if (typeof viewState.zoom === 'number') this.setZoom(viewState.zoom);
        else this.applyScaleMode('fit-page');

        // 滚动位置要等缩放后的排版落定再写回：同一帧写会被旧的内容尺寸钳住
        requestAnimationFrame(() => {
            if (!this.store.hasDoc) return;
            scroller.scrollLeft = Math.max(0, viewState.scrollLeft ?? 0);
            scroller.scrollTop = Math.max(0, viewState.scrollTop ?? 0);
            this.updateScrollbars();
        });
        if (this.mode === 'rich') this.view.focus();
    }

    /**
     * 关闭工作区（不删存档）
     */
    public async close(): Promise<void> {
        this.flushViewState();
        await this.store.close();
        this.container.hidden = true;
    }

    /** 显示 */
    public show(): void {
        this.container.hidden = false;
        this.empty.hidden = this.store.hasDoc;
        // 工具线在隐藏期间量不到尺寸，读回的分割位置会挂起：变可见时立刻落地
        this.bottomToolsSplit.refreshLayout();
        this.rightToolsSplit.refreshLayout();
        // 隐藏期间量不到尺寸（offsetWidth = 0），槽位与标尺会停在 0：
        // 变可见的这一刻必须补算一次，否则纸是一条线、标尺不出现
        this.view.refreshFrame();
        this.updateScrollbars();
        this.refresh();
    }

    /** 隐藏 */
    public hide(): void {
        // 隐藏之后量不到滚动位置：先记一次再收起
        this.flushViewState();
        this.container.hidden = true;
    }

    /**
     * 落库（页面隐藏 / 切换文件 / 关闭前调用）
     */
    public async flush(): Promise<void> {
        this.flushViewState();
        await this.store.flush();
    }

    /**
     * 把当前文档导出为 .docx 并下载
     * @param fileName - 下载文件名（含或不含后缀）
     */
    public exportDocx(fileName?: string): void {
        if (!this.store.hasDoc) return;
        const blob = snapshotToDocx({
            v: 1,
            meta: this.store.meta,
            blocks: this.store.blocks,
        });
        const name = (fileName ?? this.store.meta.title ?? 'document').replace(/\.(docx|text|md)$/i, '');
        downloadBlob(blob, `${name}.docx`);
    }

    /**
     * 导出为 Markdown（标准语法，样式降级到可被外部工具识别）
     * @returns Markdown 文本
     */
    public toMarkdown(standardOnly: boolean = false): string {
        return blocksToMarkdown(this.store.blocks, standardOnly);
    }

    /**
     * 通知工具栏「状态可能已变化」
     *
     * 文档工作区**不再自己画工具条**：段落样式 / 列表 / 缩进 / 撤销 / Markdown 全在顶部工具栏里，
     * 它们的高亮由各自的分区订阅总线后按快照回读。因此这里只需要把「状态变了」这件事转出去，
     * 保留这个方法名是为了让调用点读起来仍是「刷新外观」。
     */
    public refresh(): void {
        this.host.notifyStateChanged();
    }

    /**
     * 当前是否处于 Markdown 源码视图
     * @returns 源码视图为 true
     */
    public get isMarkdown(): boolean {
        return this.mode === 'markdown';
    }

    /**
     * 当前页面设置
     *
     * 页面设置是**文档级**元数据（存在 `DocMeta.pageSetup` 里，跟着存档走），
     * 工具栏的「页面」页签只读它的字符串形态 —— 那层转换在装配层完成，
     * 免得工具栏认识文档模型。
     * @returns 页面设置
     */
    public get pageSetup(): DocPageSetup {
        return this.store.meta?.pageSetup ?? {};
    }

    /**
     * 改一项页面设置并立刻重排纸张
     *
     * 不进命令集、不进撤销栈：它改的是「纸」而不是「内容」，
     * 放进撤销栈会出现「撤销一次编辑，纸也跟着变回 A4」的怪事。
     * @param key - 设置项
     * @param value - 新值
     */
    public setPageSetup(key: string, value: string | number): void {
        const current = this.pageSetup;
        const patch: DocPageSetup = { ...current };
        if (key === 'orientation') patch.orientation = value === 'landscape' ? 'landscape' : 'portrait';
        else if (key === 'paperSize') patch.paperSize = value as DocPageSetup['paperSize'];
        else if (key === 'marginPreset') patch.marginPreset = value as DocPageSetup['marginPreset'];
        else if (key === 'scaleMode') {
            patch.scaleMode = value === 'fit-page' || value === 'custom' ? value : 'none';
        } else if (key === 'scaleRatio') {
            // 自定义百分比：与视图缩放同一区间收敛（0.5–4），越界会在预览里缩成一团
            const ratio = Number(value);
            patch.scaleRatio = Number.isFinite(ratio) ? Math.min(4, Math.max(0.5, ratio)) : 1;
        } else return;
        this.store.updateMeta({ pageSetup: patch });
        this.view.applyPageSetup(patch);
        if (key === 'scaleMode' || key === 'scaleRatio') this.applyScaleMode(patch.scaleMode);
    }

    /**
     * 应用缩放模式
     *
     * 「适应页宽」是**视图缩放**，不是文档数据：按容器宽度反算缩放比，
     * 收敛到 0.5–4 的区间内（与 `setZoom` 同一区间）。
     * @param mode - 缩放模式
     */
    private applyScaleMode(mode: DocPageSetup['scaleMode']): void {
        if (mode === 'custom') {
            this.setZoom(this.pageSetup.scaleRatio ?? 1);
            return;
        }
        if (mode !== 'fit-page') {
            this.setZoom(1);
            return;
        }
        const host = this.view.getElements().scroller;
        const available = host.clientWidth || 0;
        const paper = DOC_PAPER_SIZES[this.pageSetup.paperSize ?? 'A4'] ?? DOC_PAPER_SIZES.A4;
        const width = this.pageSetup.orientation === 'landscape' ? paper.height : paper.width;
        if (available <= 0 || width <= 0) return;
        this.setZoom(Math.max(0.5, Math.min(4, available / width)));
    }

    /**
     * 打印当前文档
     *
     * @param pages - 1-based 页码数组；省略或 null = 全部页（打印预览里选了范围时传入）
     */
    public print(pages: number[] | null = null): void {
        this.view.print(pages);
    }

    /**
     * 打印预览
     *
     * 与**表格同一套外壳**（`utils/printPreviewShell`：全屏遮罩 + 左侧单页预览 +
     * 页脚翻页与缩放 + 右侧打印设置），只是内容由「表格 `<table>`」换成「文档文本页」。
     * 走 {@link DocPrinter} 而不是 {@link Dialog}：表格预览长什么样，文档就该长什么样，
     * 另做一套对话框会让两边的外观与手势各自漂移。
     */
    public printPreview(): void {
        if (!this.store.hasDoc) return;
        // 预览前先把 DOM 里的最新编辑同步进模型，否则最后几个字符不在纸上
        if (this.mode === 'rich') this.view.sync('打印预览', 'coalesce', false);
        if (!this.printer) {
            this.printer = new DocPrinter({
                getPageSetup: () => this.pageSetup,
                setPageSetup: (key, value) => this.setPageSetup(key, value),
                renderPages: (host) => this.view.renderPreview(host, false),
                print: (pages) => this.print(pages),
            });
        }
        this.printer.printPreview();
    }

    // ─────────────────────────── 缩放 ───────────────────────────

    /**
     * 初始化缩放滑块（挂在右侧工具线上半，控件与表格右侧的缩放滑块完全相同）
     *
     * 采用垂直细线样式（{@link Slider} 的 vertical + thin），数值框落在轨道下方。
     * 取值与 {@link DocView} 的缩放双向绑定：
     * - 拖滑块 / 填数值框 → {@link setZoom}（百分比 ÷ 100）；
     * - Ctrl+滚轮 / Ctrl+=、Ctrl+- → 也走 {@link setZoom}，由它回写滑块。
     * @param parentElement - 挂载容器（rightTools 的 split-pre）
     */
    private initZoomSlider(parentElement: HTMLElement): void {
        this.zoomSlider = new Slider({
            parentElement,
            vertical: true,      // 垂直方向：最小值在下端，数值框在轨道下方
            thin: true,          // 细线样式：2px 轨道 + 小圆点手柄
            min: ZOOM_PERCENT_MIN,
            max: ZOOM_PERCENT_MAX,
            step: 1,             // 与表格右侧的缩放滑块一致：每格 1%
            value: this.zoomPercent(),
            showInput: true,     // 数值框：手工填写缩放百分比（无单位文本，纯数字）
            inputWidth: '',      // 置空：宽度交给样式表（26px 窄条需占满整行）
            accentColor: '',     // 置空走样式表主题色，避免注入内联的浏览器默认蓝
            onChange: (value: number) => this.setZoom(value / 100),
        });
        // Ctrl + 滚轮：与表格侧同一个手势（表格在 Scroller 里处理）
        const scroller = this.view.getElements().scroller;
        scroller.addEventListener('wheel', (event: WheelEvent) => {
            if (!event.ctrlKey && !event.metaKey) return;
            // 必须拦下：不拦会变成浏览器自身的页面缩放（整个应用一起放大）
            event.preventDefault();
            const step = event.deltaY < 0 ? ZOOM_WHEEL_STEP : -ZOOM_WHEEL_STEP;
            this.setZoom(this.view.zoom + step);
        }, { passive: false });
    }

    /** 当前缩放的百分比整数（滑块的取值单位） */
    private zoomPercent(): number {
        return Math.round(this.view.zoom * 100);
    }

    /**
     * 设置缩放
     *
     * 只改视图（{@link DocView.setZoom}），不写模型：缩放是「看的比例」，不是内容属性。
     * @param value - 缩放比例（0.5–4）
     */
    public setZoom(value: number): void {
        this.view.setZoom(value);
        // 缩放属于「看的姿势」，随文档持久化
        this.scheduleViewState();
        // 回写滑块（外部入口改的缩放也要让滑块跟手）
        const percent = this.zoomPercent();
        if (this.zoomSlider && this.zoomSlider.getValue() !== percent) this.zoomSlider.setValue(percent);
        // 缩放改的是内容总长：滑块比例必须跟着重算（不滚动就不会有 scroll 事件）
        this.updateScrollbars();
    }

    /**
     * 让两条滚动条按当前内容重算滑块
     *
     * 缩放、换纸、分页都会改变内容总长，而它们**不改变滚动位置**，
     * 因此不会触发 `scroll` 事件 —— 不主动补一次，滑块会停在旧比例上
     * （表现为「放大后拖到底却看不到最后一页」）。
     */
    private updateScrollbars(): void {
        this.hScroller.update();
        this.vScroller.update();
    }

    /**
     * 切换富文本 / Markdown 源码视图
     */
    public toggleMode(): void {
        this.applyMode(this.mode === 'rich' ? 'markdown' : 'rich');
    }

    /**
     * 应用视图模式
     * @param mode - 视图模式
     */
    private applyMode(mode: 'rich' | 'markdown'): void {
        this.mode = mode;
        const richVisible = mode === 'rich';
        this.view.getElements().scroller.hidden = !richVisible;
        // 源码视图没有「纸」，标尺跟着收起（否则会留下两条对不上任何东西的刻度）
        this.view.setRulersVisible(richVisible);
        // 源码视图由 textarea 自己滚（它有自己的原生滚动条），两条自定义滚动条
        // 与右侧工具线一起收起，免得留下两条拖不动的空轨道
        this.rightTools.hidden = !richVisible;
        // 右侧工具线收起后，正文要同步让回那 26px（否则会留一条空白竖缝）
        this.container.classList.toggle('doc-source-mode', !richVisible);
        this.hScroller.getElements().container.hidden = !richVisible;
        this.vScroller.getElements().container.hidden = !richVisible;
        this.markdownView.setVisible(!richVisible);
        if (!richVisible) {
            // 切到源码前先把 DOM 里的最新编辑同步进模型，否则最后几个字符会丢
            this.view.sync('切换视图', 'coalesce', false);
            this.markdownView.refresh();
            this.markdownView.focus();
        } else {
            this.view.render();
            this.view.focus();
        }
        this.store.updateMeta({ view: mode });
        this.refresh();
    }

    // ─────────────────────────── 快捷键 ───────────────────────────

    /**
     * 绑定快捷键
     *
     * 【为什么挂在容器上而不是 document】工作区是并存的：文档打开时表格仍在 DOM 里，
     * 若挂 document 会给表格也接手 Ctrl+B（表格自己有编辑画布处理这些键）。
     */
    private bindShortcuts(): void {
        if (this.shortcutsBound) return;
        this.shortcutsBound = true;
        this.container.addEventListener('keydown', (event) => {
            const meta = event.ctrlKey || event.metaKey;
            if (!meta) return;
            const key = event.key.toLowerCase();
            if (key === 'z' && !event.shiftKey) {
                event.preventDefault();
                this.host.undo();
                return;
            }
            if (key === 'y' || (key === 'z' && event.shiftKey)) {
                event.preventDefault();
                this.host.redo();
                return;
            }
            // 缩放：Ctrl+= / Ctrl+- / Ctrl+0（与 Excel 同一组快捷键）。
            // 放在 `mode !== 'rich'` 判断**之前**：源码视图同样需要放大看。
            if (key === '=' || key === '+') {
                event.preventDefault();
                this.setZoom(this.view.zoom + ZOOM_WHEEL_STEP);
                return;
            }
            if (key === '-' || key === '_') {
                event.preventDefault();
                this.setZoom(this.view.zoom - ZOOM_WHEEL_STEP);
                return;
            }
            if (key === '0') {
                event.preventDefault();
                this.setZoom(1);
                return;
            }
            if (this.mode !== 'rich') return;
            const map: Record<string, 'toggleBold' | 'toggleItalic' | 'toggleUnderline'> = {
                b: 'toggleBold',
                i: 'toggleItalic',
                u: 'toggleUnderline',
            };
            const cmd = map[key];
            if (!cmd) return;
            event.preventDefault();
            const style = this.view.currentCharStyle();
            const active = cmd === 'toggleBold' ? !!style.bold : cmd === 'toggleItalic' ? !!style.italic : !!style.underline;
            this.bus.dispatch(cmd, { toggle: !active });
        });
        // 焦点离开正文时也刷新一次（按钮高亮跟随光标位置）
        this.container.addEventListener('keyup', () => this.refresh());
        this.container.addEventListener('mouseup', () => this.refresh());
    }

    /**
     * 当前文档的纯文本（供查找 / 统计）
     * @returns 纯文本
     */
    public plainText(): string {
        return this.store.blocks.map((block) => blockText(block)).join('\n');
    }

    /**
     * 从块树重新渲染（内容被外部替换后调用）
     */
    public render(): void {
        this.view.render();
        if (this.mode === 'markdown') this.markdownView.refresh();
        this.empty.hidden = this.store.hasDoc;
    }
}

/**
 * 触发浏览器下载
 * @param blob - 文件内容
 * @param fileName - 文件名
 */
export function downloadBlob(blob: Blob, fileName: string): void {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    // 立刻 revoke 会让部分浏览器来不及读取：放到下一轮事件循环
    window.setTimeout(() => URL.revokeObjectURL(url), 0);
}
