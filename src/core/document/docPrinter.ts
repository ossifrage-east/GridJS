/*  # 文档域 · 文档打印预览 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目目标: 给文档一套**与表格同形**的打印预览：全屏遮罩 + 标题栏 + 左侧预览区
    #           （一次一页 + 页脚翻页 + 缩放滑块）+ 右侧打印设置（纸张大小 / 方向 /
    #           页边距 / 缩放 / 页面范围 + 打印 / 关闭）。
    # 本文件职责: 只做「预览」这件事 —— 外壳直接复用 utils/printPreviewShell（与表格
    #           逐字同源），内容换成文档的**文本页**（由 DocView.renderPreview 生成）。
    #           分页、纸张尺寸、页边距一律问调用方（工作区），自己不存一份。
    # 为什么单独一个文件: 表格的 Printer 与 DataCollection 深度耦合（分页计划按单元格
    #           网格算），文档是「一条流 + 纸高」两种完全不同的分页口径；与其在一个
    #           类里塞两套分支，不如共用外壳、各算各的页。
    # 依赖方向: utils/printPreviewShell（纯字符串）+ common/slider + 同目录 blockTree。
    #           **不 import utils/printer**（那是表格的），也不 import docView
    #           （页面怎么画由调用方以回调给出）。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194739739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { Slider } from '../common/slider';
import { buildPrintPreviewShellHtml, PREVIEW_ZOOM_MAX, PREVIEW_ZOOM_MIN } from '../../utils/printPreviewShell';
import { DOC_PAPER_SIZES, DocPageSetup } from './blockTree';

/** 预览区四周留白（外壳 .gs-print-scroll 的 padding 是 1.5rem = 24px） */
const SCROLL_PADDING = 48;

/** 预览文档时来自工作区的四个回调 */
export interface DocPrinterOptions {
    /** 取当前页面设置（纸张 / 方向 / 页边距 / 缩放） */
    getPageSetup: () => DocPageSetup;
    /** 改一项页面设置（工作区内部会重排纸张与分页） */
    setPageSetup: (key: keyof DocPageSetup, value: string | number | undefined) => void;
    /** 把当前文档的全部页渲染进容器，返回页数（页元素带 `.doc-preview-page` 类名） */
    renderPages: (host: HTMLElement) => number;
    /** 打印；`pages` 为 null 表示全部页，否则是 1-based 页码数组 */
    print: (pages: number[] | null) => void;
}

/**
 * 文档打印预览
 *
 * 与表格预览的关系：外壳同源（{@link buildPrintPreviewShellHtml}），差别只有两点 ——
 * 1. 内容是**文本页**而不是 `<table>`；
 * 2. 设置项里去掉了文档没有语义的那些（按列 / 按行适配、页边距自定义编辑器）。
 */
export class DocPrinter {
    /** 工作区回调 */
    private readonly options: DocPrinterOptions;

    /** 遮罩层（未打开时为 null） */
    private overlay: HTMLDivElement | null = null;

    /** 页面容器 */
    private pagesHost: HTMLElement | null = null;

    /** 当前页已渲染的页元素 */
    private pageEls: HTMLElement[] = [];

    /** 当前页索引（0-based） */
    private currentPage = 0;

    /** 总页数 */
    private totalPages = 1;

    /** 页码指示器 */
    private indicator: HTMLElement | null = null;

    /** 上一页 / 下一页按钮 */
    private prevBtn: HTMLButtonElement | null = null;
    private nextBtn: HTMLButtonElement | null = null;

    /** 预览缩放滑块 */
    private zoomSlider: Slider | null = null;

    /** 手动缩放（百分比）；null = 按可用区域自动适配 */
    private zoomPercent: number | null = null;

    /** 打印机名称（浏览器沙箱无法枚举系统打印机，仅回显） */
    private printerName = '默认打印机';

    /** 份数（仅回显：份数最终在浏览器打印对话框里确认） */
    private copies = 1;

    /** 打印顺序（同上，仅回显） */
    private collate: 'page' | 'copies' = 'page';

    /** 页面范围输入（如「1-3,5」） */
    private pageRange = '';

    /** 键盘 / resize 的监听引用（关闭时精确退订） */
    private boundKeyDown: ((e: KeyboardEvent) => void) | null = null;
    private boundResize: (() => void) | null = null;

    /**
     * 构造函数
     * @param options - 见 {@link DocPrinterOptions}
     */
    constructor(options: DocPrinterOptions) {
        this.options = options;
    }

    /**
     * 打开打印预览
     */
    public printPreview(): void {
        if (this.overlay) this.closePreview();

        const overlay = document.createElement('div');
        overlay.className = 'gs-print-overlay';
        overlay.innerHTML = buildPrintPreviewShellHtml({
            title: '打印预览',
            // 正文样式：预览页是主页面 DOM 的克隆，类名样式必须一起搬进来，
            // 否则预览里只剩一堆没有排版的裸标签（见 {@link collectDocStyles}）
            extraCss: collectDocStyles(),
        });
        document.body.appendChild(overlay);
        this.overlay = overlay;

        this.pagesHost = overlay.querySelector<HTMLElement>('.gs-print-pages');
        this.indicator = overlay.querySelector<HTMLElement>('.gs-print-indicator');
        this.prevBtn = overlay.querySelector<HTMLButtonElement>('[data-gs-print-nav="prev"]');
        this.nextBtn = overlay.querySelector<HTMLButtonElement>('[data-gs-print-nav="next"]');

        this.bindPanel(overlay);
        this.bindZoom(overlay);

        // 事件委托：打印 / 关闭 / 翻页（与表格同一个写法）
        overlay.addEventListener('click', (e) => {
            const target = (e.target as HTMLElement).closest('[data-gs-print-action], [data-gs-print-nav]') as HTMLElement | null;
            if (!target) return;
            if (target.dataset.gsPrintAction === 'print') {
                this.options.print(this.selectedPages());
                return;
            }
            if (target.dataset.gsPrintAction === 'close') {
                this.closePreview();
                return;
            }
            if (target.dataset.gsPrintNav === 'prev') this.goToPage(this.currentPage - 1);
            if (target.dataset.gsPrintNav === 'next') this.goToPage(this.currentPage + 1);
        });

        this.boundKeyDown = (e: KeyboardEvent) => {
            if (e.key === 'Escape') this.closePreview();
        };
        document.addEventListener('keydown', this.boundKeyDown);
        this.boundResize = () => this.applyZoom();
        window.addEventListener('resize', this.boundResize);

        this.rerender();
    }

    /**
     * 关闭预览并清理监听
     */
    public closePreview(): void {
        if (this.boundKeyDown) document.removeEventListener('keydown', this.boundKeyDown);
        if (this.boundResize) window.removeEventListener('resize', this.boundResize);
        this.boundKeyDown = null;
        this.boundResize = null;
        if (this.overlay && this.overlay.parentElement) this.overlay.parentElement.removeChild(this.overlay);
        this.overlay = null;
        this.pagesHost = null;
        this.pageEls = [];
        this.zoomSlider = null;
    }

    // ─────────────────────────── 右侧设置面板 ───────────────────────────

    /**
     * 绑定设置面板
     *
     * 与表格面板的唯一差别是「去掉文档没有的项」：按列 / 按行适配（文档没有列的概念）、
     * 自定义页边距编辑器（文档的页边距是三档预设，见 `DOC_MARGIN_PRESETS`）。
     * @param overlay - 遮罩层
     * @private
     */
    private bindPanel(overlay: HTMLDivElement): void {
        const setup = this.options.getPageSetup();

        // 纸张大小：换成文档支持的那一组（表格是 A4/A3/B5/Letter）
        const paperSelect = overlay.querySelector<HTMLSelectElement>('[data-gs-print-role="paper-size"]');
        if (paperSelect) {
            paperSelect.innerHTML = Object.keys(DOC_PAPER_SIZES)
                .map((size) => `<option value="${size}">${size}</option>`)
                .join('');
            paperSelect.value = setup.paperSize ?? 'A4';
            paperSelect.addEventListener('change', () => {
                this.options.setPageSetup('paperSize', paperSelect.value);
                this.zoomPercent = null;   // 换纸 → 自动适配失效，回到自动
                this.rerender();
            });
        }

        const orientationSelect = overlay.querySelector<HTMLSelectElement>('[data-gs-print-role="orientation"]');
        if (orientationSelect) {
            orientationSelect.value = setup.orientation ?? 'portrait';
            orientationSelect.addEventListener('change', () => {
                this.options.setPageSetup('orientation', orientationSelect.value);
                this.zoomPercent = null;
                this.rerender();
            });
        }

        // 页边距：只剩三档预设，配套的「调整页边距」按钮一并移除
        const marginSelect = overlay.querySelector<HTMLSelectElement>('[data-gs-print-role="margin"]');
        if (marginSelect) {
            const custom = marginSelect.querySelector('option[value="custom"]');
            if (custom) custom.remove();
            marginSelect.value = setup.marginPreset ?? 'normal';
            marginSelect.addEventListener('change', () => {
                this.options.setPageSetup('marginPreset', marginSelect.value);
                this.zoomPercent = null;
                this.rerender();
            });
        }
        overlay.querySelector<HTMLElement>('[data-gs-print-action="margin-adjust"]')?.remove();
        overlay.querySelector<HTMLElement>('[data-gs-print-role="margin-editor"]')?.remove();

        // 缩放：去掉「所有列一页 / 所有行一页」（文档没有列），「整个工作表」改名
        const scaleSelect = overlay.querySelector<HTMLSelectElement>('[data-gs-print-role="scale"]');
        const scaleInput = overlay.querySelector<HTMLInputElement>('[data-gs-print-role="scale-input"]');
        if (scaleSelect) {
            scaleSelect.querySelector('option[value="fit-cols"]')?.remove();
            scaleSelect.querySelector('option[value="fit-rows"]')?.remove();
            const fitSheet = scaleSelect.querySelector('option[value="fit-sheet"]');
            if (fitSheet) {
                fitSheet.setAttribute('value', 'fit-page');
                fitSheet.textContent = '整个文档一页';
            }
            scaleSelect.value = setup.scaleMode ?? 'none';
            const syncScaleInput = () => {
                if (!scaleInput) return;
                scaleInput.value = String(Math.round((setup.scaleRatio ?? 1) * 100));
                scaleInput.disabled = scaleSelect.value !== 'custom';
            };
            syncScaleInput();
            scaleSelect.addEventListener('change', () => {
                const mode = scaleSelect.value as DocPageSetup['scaleMode'];
                this.options.setPageSetup('scaleMode', mode);
                syncScaleInput();
                if (mode === 'custom') scaleInput?.select();
                this.rerender();
            });
            scaleInput?.addEventListener('change', () => {
                const n = Number(scaleInput.value);
                if (!Number.isFinite(n) || n <= 0) return;
                const ratio = Math.min(4, Math.max(0.1, n / 100));
                this.options.setPageSetup('scaleRatio', ratio);
                this.options.setPageSetup('scaleMode', 'custom');
                scaleInput.value = String(Math.round(ratio * 100));
                this.rerender();
            });
        }

        // 打印机 / 份数 / 顺序 / 打印方式：与表格同样的「仅回显」口径
        const printerSelect = overlay.querySelector<HTMLSelectElement>('[data-gs-print-role="printer"]');
        if (printerSelect) {
            printerSelect.innerHTML = '<option value="默认打印机">默认打印机</option>';
            printerSelect.value = this.printerName;
            printerSelect.addEventListener('change', () => {
                this.printerName = printerSelect.value;
            });
        }
        const copiesInput = overlay.querySelector<HTMLInputElement>('[data-gs-print-role="copies"]');
        if (copiesInput) {
            copiesInput.value = String(this.copies);
            copiesInput.addEventListener('change', () => {
                const n = Number(copiesInput.value);
                this.copies = Number.isFinite(n) ? Math.min(999, Math.max(1, Math.floor(n))) : 1;
                copiesInput.value = String(this.copies);
            });
            overlay.querySelectorAll<HTMLButtonElement>('[data-gs-print-step]').forEach((btn) => {
                const delta = btn.dataset.gsPrintStep === 'inc' ? 1 : -1;
                btn.addEventListener('click', () => {
                    this.copies = Math.min(999, Math.max(1, this.copies + delta));
                    copiesInput.value = String(this.copies);
                });
            });
        }
        const collateSelect = overlay.querySelector<HTMLSelectElement>('[data-gs-print-role="collate"]');
        if (collateSelect) {
            collateSelect.value = this.collate;
            collateSelect.addEventListener('change', () => {
                this.collate = collateSelect.value as 'page' | 'copies';
            });
        }
        // 双面打印依赖「页」的奇偶排版，文档暂不提供，直接移除该行
        const duplexSelect = overlay.querySelector<HTMLSelectElement>('[data-gs-print-role="duplex"]');
        if (duplexSelect) duplexSelect.closest('.gs-print-field')?.remove();

        // 页面范围
        const rangeSelect = overlay.querySelector<HTMLSelectElement>('[data-gs-print-role="page-range"]');
        const rangeRow = overlay.querySelector<HTMLElement>('[data-gs-print-role="page-range-row"]');
        const rangeInput = overlay.querySelector<HTMLInputElement>('[data-gs-print-role="page-range-input"]');
        if (rangeSelect) {
            rangeSelect.addEventListener('change', () => {
                const custom = rangeSelect.value === 'custom';
                if (rangeRow) rangeRow.style.display = custom ? '' : 'none';
                if (custom) rangeInput?.focus();
            });
        }
        if (rangeInput) {
            rangeInput.addEventListener('change', () => {
                this.pageRange = rangeInput.value.trim();
            });
        }
    }

    /**
     * 绑定页脚的预览缩放滑块
     *
     * 拖动中不回写轨道（`SliderChangeSource === 'track'`）：程序回写 value 会打断
     * 原生拖拽手势，表现是「滑块拖到一半自己跳回去」。
     * @param overlay - 遮罩层
     * @private
     */
    private bindZoom(overlay: HTMLDivElement): void {
        const zoomHost = overlay.querySelector<HTMLElement>('[data-gs-print-role="zoom"]');
        if (!zoomHost) return;
        this.zoomSlider = new Slider({
            parentElement: zoomHost,
            min: PREVIEW_ZOOM_MIN,
            max: PREVIEW_ZOOM_MAX,
            step: 1,
            unit: '%',
            sliderWidth: '120px',
            accentColor: '#0d6efd',
            onChange: (value: number, source: string) => {
                this.zoomPercent = Math.min(PREVIEW_ZOOM_MAX, Math.max(PREVIEW_ZOOM_MIN, value));
                this.applyZoom(source !== 'track');
            },
        });
    }

    // ─────────────────────────── 渲染 / 翻页 / 缩放 ───────────────────────────

    /**
     * 重新渲染全部页并回到当前页
     *
     * 纸张 / 方向 / 页边距 / 缩放任一变化都走这里：这些设置会改变**分页结果**，
     * 只改缩放不动内容是不对的。
     * @private
     */
    private rerender(): void {
        if (!this.pagesHost) return;
        const pages = this.options.renderPages(this.pagesHost);
        this.totalPages = Math.max(1, pages);
        this.pageEls = Array.prototype.slice.call(
            this.pagesHost.querySelectorAll('.doc-preview-page')
        ) as HTMLElement[];
        this.currentPage = Math.min(this.currentPage, Math.max(0, this.totalPages - 1));
        this.showPage();
    }

    /**
     * 只显示当前页（其余隐藏，不销毁：翻页来回切不必重排）
     * @private
     */
    private showPage(): void {
        for (let i = 0; i < this.pageEls.length; i++) {
            const el = this.pageEls[i];
            // 不显示的那几页用 display:none 而不是移除：移除后翻回去要重渲染，
            // 而「页数」是由渲染函数回答的，反复重建容易与页码指示器对不上
            el.style.display = i === this.currentPage ? '' : 'none';
            el.style.margin = i === this.currentPage ? '0 auto' : 'none';
        }
        if (this.indicator) this.indicator.textContent = `第${this.currentPage + 1}页/共${this.totalPages}页`;
        if (this.prevBtn) this.prevBtn.disabled = this.currentPage <= 0;
        if (this.nextBtn) this.nextBtn.disabled = this.currentPage >= this.totalPages - 1;
        this.applyZoom();
    }

    /**
     * 跳页（越界自动收敛）
     * @param page - 目标页索引（0-based）
     * @private
     */
    private goToPage(page: number): void {
        this.currentPage = Math.max(0, Math.min(page, Math.max(0, this.totalPages - 1)));
        this.showPage();
    }

    /**
     * 应用预览缩放
     *
     * 与表格同一口径：用 CSS `zoom`（而不是 `transform: scale`）—— 后者不改变布局盒，
     * 页面会在滚动区里溢出/留白，`zoom` 则会把页面连同它的占位一起缩小。
     * @param syncControls - 是否把当前比例回写滑块
     * @private
     */
    private applyZoom(syncControls: boolean = true): void {
        const zoom = this.zoomPercent !== null ? this.zoomPercent / 100 : this.calcFitZoom();
        const el = this.pageEls[this.currentPage];
        if (el) {
            if (Math.abs(zoom - 1) > 0.001) el.style.setProperty('zoom', String(zoom));
            else el.style.removeProperty('zoom');
        }
        if (!syncControls) return;
        this.zoomSlider?.setValue(Math.round(zoom * 100));
    }

    /**
     * 自动适配：把整页塞进预览区的可用高度（宽度同时受限时取小者）
     * @returns 缩放比例
     * @private
     */
    private calcFitZoom(): number {
        const scroll = this.overlay?.querySelector<HTMLElement>('.gs-print-scroll');
        const el = this.pageEls[this.currentPage];
        if (!scroll || !el) return 1;
        const availH = scroll.clientHeight - SCROLL_PADDING;
        const availW = scroll.clientWidth - SCROLL_PADDING;
        const width = el.offsetWidth || 0;
        const height = el.offsetHeight || 0;
        if (availH <= 0 || availW <= 0 || width <= 0 || height <= 0) return 1;
        return Math.min(4, Math.max(0.1, Math.min(availH / height, availW / width)));
    }

    // ─────────────────────────── 打印范围 ───────────────────────────

    /**
     * 解析页面范围（如「1-3,5」）
     *
     * 与表格同一口径：非法片段直接丢掉而不是整体判错 —— 手输一个错字符就把整次打印
     * 变成「全部页」，比悄悄少打几页更难发现。
     * @returns 1-based 页码数组；未指定范围返回 null（打印全部）
     * @private
     */
    private selectedPages(): number[] | null {
        const text = this.pageRange.trim();
        if (!text) return null;
        const out: number[] = [];
        for (const part of text.split(',')) {
            const seg = part.trim();
            if (!seg) continue;
            const range = /^(\d+)\s*-\s*(\d+)$/.exec(seg);
            if (range) {
                const from = Math.max(1, Number(range[1]));
                const to = Math.min(this.totalPages, Number(range[2]));
                for (let i = from; i <= to; i++) out.push(i);
                continue;
            }
            const single = Number(seg);
            if (Number.isFinite(single) && single >= 1 && single <= this.totalPages) out.push(Math.floor(single));
        }
        return out.length > 0 ? out : null;
    }
}

/**
 * 收集主样式表里与文档正文有关的规则
 *
 * 预览页是主页面 DOM 的克隆（见 `DocView.renderPreview`），标签靠**类名**排版；
 * 把 `.doc-*` 规则一起搬进预览的 `<style>`，预览才与屏幕所见一致。
 * 只收正文相关的类：布局类（滚动区 / 标尺 / 缩放条 …）跟进来只会让预览页变形。
 * @returns CSS 文本
 */
function collectDocStyles(): string {
    const skip = /(doc-scroller|doc-ruler|doc-workspace|doc-zoom|doc-empty|doc-stage|doc-markdown|doc-preview-stage)/;
    const out: string[] = [];
    for (const sheet of Array.from(document.styleSheets)) {
        try {
            const rules = Array.from((sheet as CSSStyleSheet).cssRules);
            for (const rule of rules) {
                const selector = (rule as CSSStyleRule).selectorText;
                if (!selector || !/\.doc-/.test(selector) || skip.test(selector)) continue;
                out.push(rule.cssText);
            }
        } catch {
            // 跨域样式表不允许访问 cssRules，跳过即可
        }
    }
    return out.join('\n');
}
