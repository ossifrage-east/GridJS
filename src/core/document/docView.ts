/*  # 文档域 · DOM 视图与编辑器 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目目标: 文档正文的所见即所得编辑。**渲染走 DOM，不走 Canvas** ——
    #           表格用 Canvas 是因为「十万个同构单元格」；文档要的是行内混排、任意字符区间选区、
    #           输入法组字与命中测试，这四项恰好都是 DOM 的原生强项、Canvas 的弱项。
    # 本文件职责: 把块树渲染成 DOM（{@link DocView.render}），并把 DOM 上的编辑结果
    #             **同步回块树**（{@link DocView.sync}）；同时提供工具栏需要的
    #             「当前选区样式查询」与「对选区应用样式」两组能力。
    # 编辑模型（重要，勿随手改）:
    #   - **输入期间以 DOM 为准**：打字 / 回车 / 删除交给浏览器，`input` 事件后同步回模型；
    #   - **程序化改动以模型为准**：撤销、导入、切换文档时整体重渲染；
    #   - 样式命令（粗体 / 字号 / 对齐…）直接操作 DOM 选区再同步，避免"重渲染 + 复位光标"
    #     这套在 IME 下极不稳定的做法。
    #   完全受控（每次模型变更都重建 DOM）在中文输入法下会出现「组字被重置、候选窗错位」，
    #   这是本实现刻意避开的地方。
    # 依赖方向: 同目录 blockTree / docStore + utils/dom；**不 import 表格数据层**。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { createDiv } from '../../utils/dom';
import { Char } from '../dataArchitecture/char';
import {
    DEFAULT_DOC_FONT_FAMILY,
    DEFAULT_DOC_FONT_SIZE,
    DOC_INDENT_STEP_PX,
    DOC_MARGIN_PRESETS,
    DOC_MAX_PAGINATE_BLOCKS,
    DOC_PAGE_GAP_ATTR,
    DOC_PAPER_SIZES,
    DOC_ZOOM_PERCENT_MAX,
    DOC_ZOOM_PERCENT_MIN,
} from '../constant';
import {
    CharStylePatch,
    DocAlign,
    DocBlock,
    DocListStyle,
    DocPageSetup,
    DocSpan,
    blockText,
    charsFromText,
    createBlock,
    nextBlockId,
    writeCharStyle,
} from './blockTree';
import { DOC_CHANGED, DOC_RESET, DocStore } from './docStore';
import { buildRuler } from './docRuler';
import '../style/document.scss';

/** 工具栏需要的段落状态 */
export interface DocParaState {
    /** 块类型 */
    type: DocBlock['type'];
    /** 标题级别 / 列表层级 */
    level: number;
    /** 列表风格 */
    listStyle: DocListStyle;
    /** 水平对齐 */
    align: DocAlign;
    /** 额外缩进级数（`data-indent`，与列表层级 `level` 是两回事） */
    indent: number;
}

/** 选区快照（用于重渲染后恢复光标） */
export interface DocCaret {
    /** 块 id */
    blockId: string;
    /** 块内码点偏移 */
    offset: number;
}

/* 分页标记属性名与逐块排页上限统一见 `core/constant.ts`
 *（DOC_PAGE_GAP_ATTR / DOC_MAX_PAGINATE_BLOCKS）：
 * 超过上限就退化成「只算页数、不插留白」（见 {@link DocView.paginate}）——
 * 分页是显示层的东西，不该把输入延迟拖下水。 */

/** 纸张几何信息（分页 / 打印预览共用，避免两处各算一遍纸高） */
export interface DocPageGeometry {
    /** 纸张宽（CSS 像素，未缩放） */
    width: number;
    /** 纸张高（CSS 像素，未缩放） */
    height: number;
    /** 页边距（四边同值） */
    margin: number;
    /** 一页正文区高度 = 纸高 − 上下页边距 */
    contentHeight: number;
    /** 总页数（至少 1） */
    pages: number;
}

/**
 * 文档视图
 *
 * @example
 * ```ts
 * const view = new DocView({ store, parentElement });
 * view.render();
 * ```
 */
export class DocView {
    /** 文档存储（真源） */
    private readonly store: DocStore;

    /** 滚动容器 */
    private readonly scroller: HTMLDivElement;

    /**
     * 横向标尺（页宽方向，带厘米数字）
     *
     * 标尺**不挂在滚动区里**：滚一屏纸，标尺若跟着滚走就失去了「参照物」的意义。
     * 它挂在工作区正文（`.doc-body`，见 {@link DocView.mountRulers}）的上边与左边，
     * 靠 `scroll` 事件反向平移内部的刻度条 {@link DocView.rulerHInner} 来跟随内容。
     */
    private readonly rulerH: HTMLDivElement;

    /** 纵向标尺（页高方向，只画刻度） */
    private readonly rulerV: HTMLDivElement;

    /** 横向标尺的刻度条（宽度 = 纸 × 缩放，随滚动反向平移） */
    private readonly rulerHInner: HTMLDivElement;

    /** 纵向标尺的刻度条（高度 = 纸高 × 缩放，随滚动反向平移） */
    private readonly rulerVInner: HTMLDivElement;

    /** 纸张槽位（尺寸 = 纸 × 缩放；纸在其中以 `top left` 为原点缩放） */
    private readonly slot: HTMLDivElement;

    /** 纸张 */
    private readonly page: HTMLDivElement;

    /** 块容器（contenteditable 宿主） */
    private readonly blocksEl: HTMLDivElement;

    /** 同步节流句柄（`input` 高频触发，用 rAF 合并） */
    private syncFrame: number | null = null;

    /** 正在程序化重渲染（此时忽略 `input`） */
    private rendering: boolean = false;

    /** 折叠光标上待生效的样式（下次输入时套用） */
    private pendingStyle: CharStylePatch | null = null;

    /** 缩放比例（仅视觉缩放，不改模型字号） */
    private zoomValue: number = 1;

    /** 最近一次应用的页面设置（`setZoom` 重算纸张时要复用） */
    private pageSetupCache: DocPageSetup | null = null;

    /** 总页数（由 {@link DocView.paginate} 算出，供打印预览与状态栏显示） */
    private pageCountValue: number = 1;

    /** 已安排但还没跑的一次分页（同一帧内多次编辑只分页一次） */
    private paginateFrame: number | null = null;

    /**
     * 构造函数
     * @param options.store - 文档存储
     * @param options.parentElement - 挂载点
     */
    constructor(options: { store: DocStore; parentElement: HTMLElement }) {
        this.store = options.store;

        this.scroller = createDiv({ className: 'doc-scroller' });
        this.rulerH = createDiv({ className: 'doc-ruler doc-ruler-h' });
        this.rulerV = createDiv({ className: 'doc-ruler doc-ruler-v' });
        this.rulerHInner = createDiv({ className: 'doc-ruler-inner' });
        this.rulerVInner = createDiv({ className: 'doc-ruler-inner' });
        // 刻度条是尺子的**内容**：尺子本身定长（贴在正文上边 / 左边），
        // 刻度条才按纸长铺开并由滚动反向平移 —— 换个容器就能直接复用
        this.rulerH.appendChild(this.rulerHInner);
        this.rulerV.appendChild(this.rulerVInner);
        this.slot = createDiv({ className: 'doc-page-slot' });
        this.page = createDiv({ className: 'doc-page' });
        this.blocksEl = createDiv({ className: 'doc-blocks' });
        this.blocksEl.setAttribute('contenteditable', 'true');
        this.blocksEl.setAttribute('spellcheck', 'false');
        this.blocksEl.setAttribute('role', 'textbox');
        this.blocksEl.setAttribute('aria-multiline', 'true');

        this.page.appendChild(this.blocksEl);
        this.slot.appendChild(this.page);
        this.scroller.appendChild(this.slot);
        options.parentElement.appendChild(this.scroller);

        // 滚动时反向平移刻度条：纸往左走多少，刻度就跟着往左挪多少，零点始终贴在纸边
        this.scroller.addEventListener('scroll', () => this.syncRulers());

        this.bindEvents();
        this.store.on(DOC_CHANGED, () => this.onStoreChanged());
        this.store.on(DOC_RESET, () => {
            this.render();
            // 换文档 = 换一份页面设置：纸张 / 页边距跟着新文档的元数据走
            this.applyPageSetup(this.store.meta?.pageSetup ?? {});
        });
        this.applyPageSetup(this.store.meta?.pageSetup ?? {});
    }

    /**
     * 取文档区的各个元素
     *
     * `page` 是编辑宿主判断「选区在不在正文里」的依据（见 `DocEditHost`），
     * `scroller` 用于工作区显隐与「Ctrl+滚轮缩放」的监听目标。
     */
    public getElements(): {
        scroller: HTMLDivElement;
        page: HTMLDivElement;
        slot: HTMLDivElement;
        rulerH: HTMLDivElement;
        rulerV: HTMLDivElement;
    } {
        return { scroller: this.scroller, page: this.page, slot: this.slot, rulerH: this.rulerH, rulerV: this.rulerV };
    }

    /**
     * 把两把标尺挂到指定容器（工作区正文 `.doc-body`）
     *
     * 标尺与滚动区是**兄弟**而不是父子：挂在滚动区里会随内容一起滚走，滚一屏之后
     * 刻度就与纸边错开了。挂到正文上再由 {@link DocView.syncRulers} 反向平移刻度条，
     * 才是「尺子不动、刻度跟着纸走」的正确形态（与 Word 一致）。
     * @param host - 挂载点（绝对定位的容器：标尺钉在它的上边与左边）
     */
    public mountRulers(host: HTMLElement): void {
        host.appendChild(this.rulerH);
        host.appendChild(this.rulerV);
    }

    /**
     * 显隐两把标尺
     *
     * Markdown 源码视图里没有「纸」，标尺也就没有刻度可参照，留着只会让人以为
     * 页面没渲染完。
     * @param visible - 是否显示
     */
    public setRulersVisible(visible: boolean): void {
        this.rulerH.hidden = !visible;
        this.rulerV.hidden = !visible;
    }

    /**
     * 按滚动位置反向平移刻度条
     *
     * 只改 `transform`：不触发重排，滚动时也不会把刻度条自身的宽度算进去。
     */
    private syncRulers(): void {
        this.rulerHInner.style.transform = `translateX(${-this.scroller.scrollLeft}px)`;
        this.rulerVInner.style.transform = `translateY(${-this.scroller.scrollTop}px)`;
    }

    /** 当前缩放比例 */
    public get zoom(): number {
        return this.zoomValue;
    }

    /**
     * 设置缩放（0.5–4）
     *
     * 缩放只改**视觉尺寸**（`transform: scale`），不动模型里的字号 ——
     * 否则「看不清放大一点」会变成一次不可逆的内容修改。
     */
    public setZoom(value: number): void {
        const next = Math.max(DOC_ZOOM_PERCENT_MIN / 100, Math.min(DOC_ZOOM_PERCENT_MAX / 100, value));
        if (next === this.zoomValue) return;
        this.zoomValue = next;
        // 缩放改的是同一张纸的视觉大小：纸张 / 页边距设置必须跟着重算，
        // 否则「换 A3 之后调缩放」会把纸张尺寸覆盖回默认。
        if (this.pageSetupCache) this.applyPageSetup(this.pageSetupCache);
        else this.applyFrame();
    }

    /**
     * 应用页面设置（纸张方向 / 大小 / 页边距）
     *
     * 只改「纸」本身（宽度与内边距），不动块树 —— 页面设置是**文档级**元数据，
     * 存在 `DocMeta.pageSetup` 里，跟着存档走。
     * @param setup - 页面设置
     */
    public applyPageSetup(setup: DocPageSetup): void {
        this.pageSetupCache = { ...setup };
        const { width, height, margin } = this.pageGeometry();
        this.page.style.width = `${width}px`;
        this.page.style.minHeight = `${height}px`;
        this.page.style.padding = `${margin}px`;
        // 纸变了 → 分页位置与标尺都要跟着重算（旧的分页留白是按上一张纸算的）
        this.applyFrame();
        this.schedulePaginate();
    }

    /**
     * 把「纸的缩放 + 槽位尺寸 + 两把标尺」一次性刷成当前状态
     *
     * 三件事必须同帧做完，否则会出现「纸缩了、标尺没缩」这类半新半旧的画面：
     * - 纸：`transform: scale`，原点 `top left`（与槽位左上角对齐）；
     * - 槽位：尺寸 = 纸的**布局尺寸** × 缩放（`transform` 不占布局，没有槽位纸会溢出容器）；
     * - 标尺：长度取纸的布局尺寸（横向纸宽、纵向纸高 × 页数），与槽位同宽高。
     * @private
     */
    private applyFrame(): void {
        const z = this.zoomValue;
        // 缩放原点必须是左上角：用 `top center` 时纸会相对槽位水平漂移（差半个缩放差），
        // 标尺与纸边就对不上了
        this.page.style.transformOrigin = 'top left';
        this.page.style.transform = `scale(${z})`;
        const width = this.page.offsetWidth;
        const height = this.page.offsetHeight;
        if (width > 0) this.slot.style.width = `${width * z}px`;
        if (height > 0) this.slot.style.height = `${height * z}px`;
        buildRuler(this.rulerHInner, width, z, true);
        // 纵向量的是「整叠纸」：分页后纸会被撑到 N 页高，标尺要跟着长
        buildRuler(this.rulerVInner, height, z, false);
        // 缩放 / 换纸之后滚动位置可能已经越界，刻度条要按新的 scrollLeft/Top 归位
        this.syncRulers();
    }

    /**
     * 重算「槽位尺寸 + 标尺 + 分页」
     *
     * 供「容器刚从隐藏变可见」时调用：隐藏元素的 `offsetWidth` 是 0，
     * 那一帧算出来的槽位与标尺全是 0，之后没有任何事件会再触发一次 ——
     * 表现为「打开文档后纸张缩成一条、标尺不见」。显隐切换只能由工作区主动补一次。
     */
    public refreshFrame(): void {
        this.applyFrame();
        this.paginate();
    }

    /**
     * 取当前纸张的几何信息（分页 / 打印预览共用的唯一口径）
     * @returns 纸张宽高、页边距、正文区高度
     */
    public pageGeometry(): Omit<DocPageGeometry, 'pages'> {
        const setup = this.pageSetupCache ?? {};
        const paper = DOC_PAPER_SIZES[setup.paperSize ?? 'A4'] ?? DOC_PAPER_SIZES.A4;
        const landscape = setup.orientation === 'landscape';
        const width = landscape ? paper.height : paper.width;
        const height = landscape ? paper.width : paper.height;
        const margin = DOC_MARGIN_PRESETS[setup.marginPreset ?? 'normal'] ?? DOC_MARGIN_PRESETS.normal;
        return { width, height, margin, contentHeight: Math.max(120, height - margin * 2) };
    }

    /** 当前总页数（至少 1） */
    public get pageCount(): number {
        return this.pageCountValue;
    }

    /**
     * 打印当前文档
     *
     * 走隐藏 iframe 而不是 `window.print()`：后者会连应用外壳（工具栏、文件树）一起印出来。
     * iframe 里**按页**复制纸张内容（每页一个与纸张等大的壳，靠 `break-after: page`
     * 一页一张纸），再把主文档的样式表原样抄进去，这样打印结果与屏幕所见一致。
     *
     * 打印前先 `paginate()`：分页留白决定了「哪一段在第几页」，不排一次页，
     * 打印出来的分页位置会与屏幕上看到的对不上。
     */
    public print(pages: number[] | null = null): void {
        this.paginate();
        const iframe = document.createElement('iframe');
        iframe.setAttribute('aria-hidden', 'true');
        iframe.style.position = 'fixed';
        iframe.style.right = '0';
        iframe.style.bottom = '0';
        iframe.style.width = '0';
        iframe.style.height = '0';
        iframe.style.border = '0';
        document.body.appendChild(iframe);
        const win = iframe.contentWindow;
        const doc = iframe.contentDocument;
        if (!win || !doc) {
            iframe.remove();
            return;
        }
        doc.open();
        doc.write('<!DOCTYPE html><html><head><meta charset="utf-8"><title>打印</title></head><body></body></html>');
        doc.close();
        // 抄样式表：跨域表读 cssRules 会抛，逐张 try/catch 跳过
        const sheets = document.styleSheets;
        for (let i = 0; i < sheets.length; i++) {
            try {
                const rules = sheets[i].cssRules;
                let css = '';
                for (let j = 0; j < rules.length; j++) css += rules[j].cssText + '\n';
                if (!css) continue;
                const style = doc.createElement('style');
                style.appendChild(doc.createTextNode(css));
                doc.head.appendChild(style);
            } catch {
                continue;
            }
        }
        const { width, height } = this.pageGeometry();
        const style = doc.createElement('style');
        // @page 用**毫米**：打印引擎按物理纸张分页，写像素会被 DPI 换算搅乱。
        // margin 置零：页边距已经由每页壳的 padding 画出来了，再让引擎加一次会双倍。
        style.appendChild(doc.createTextNode(
            `@page { size: ${pxToMm(width)}mm ${pxToMm(height)}mm; margin: 0; }\n`
            + `.doc-preview-page { break-after: page; page-break-after: always; overflow: hidden; background: #fff; }\n`
            + `.doc-preview-clip { overflow: hidden; }\n`
        ));
        doc.head.appendChild(style);
        this.renderPreview(doc.body, true);
        // 页码范围（打印预览里选的）：整页已经渲染好了，按范围删掉不要的那几页 ——
        // 比「只渲染范围内的页」简单，也不会因为页号换算差一位而少打一页
        if (pages && pages.length > 0) {
            const shells = doc.body.querySelectorAll('.doc-preview-page');
            const keep = new Set<number>(pages);
            for (let i = 0; i < shells.length; i++) {
                if (keep.has(i + 1)) continue;
                const shell = shells[i];
                if (shell.parentElement) shell.parentElement.removeChild(shell);
            }
        }
        win.focus();
        win.print();
        window.setTimeout(() => iframe.remove(), 1000);
    }

    /**
     * 按当前分页把文档渲染成「一页一壳」的静态副本
     *
     * 屏幕上的正文是**一条连续流**（分页只是在流里插留白），打印/预览要的是
     * 「一张纸一段」。做法是把整条流克隆 N 份，每份向上平移 `i × 纸高` 再裁掉多余部分 ——
     * 不拆 DOM、不动画布，内容与屏幕严格一致。
     * @param host - 目标容器（会被清空）
     * @param forPrint - 打印用（去阴影与边框）还是预览用（保留纸感外观）
     * @returns 页数
     */
    public renderPreview(host: HTMLElement, forPrint: boolean = false): number {
        this.paginate();
        const { width, height, margin } = this.pageGeometry();
        const pages = this.pageCountValue;
        while (host.firstChild) host.removeChild(host.firstChild);
        const contentWidth = Math.max(80, width - margin * 2);
        for (let i = 0; i < pages; i++) {
            const shell = createDiv({ className: 'doc-preview-page' });
            shell.style.width = `${width}px`;
            shell.style.height = `${height}px`;
            shell.style.padding = `${margin}px`;
            if (forPrint) shell.style.boxShadow = 'none';
            const clip = createDiv({ className: 'doc-preview-clip' });
            clip.style.height = `${Math.max(80, height - margin * 2)}px`;

            const inner = this.page.cloneNode(true) as HTMLElement;
            // 打印按真实纸张尺寸：去掉屏幕缩放、去掉阴影，宽度收成正文区宽
            inner.style.transform = 'none';
            inner.style.margin = '0';
            inner.style.padding = '0';
            inner.style.width = `${contentWidth}px`;
            inner.style.minHeight = '0';
            inner.style.boxShadow = 'none';
            // 第 i 页的正文起点在版面坐标的 i×纸高处（与 {@link paginate} 同一口径）
            inner.style.marginTop = `${-(i * height)}px`;
            // 副本不是给人编辑的：留着 contenteditable 会在预览里弹出光标与拼写提示
            const blocks = inner.querySelector('.doc-blocks');
            if (blocks) blocks.setAttribute('contenteditable', 'false');
            clip.appendChild(inner);
            shell.appendChild(clip);
            host.appendChild(shell);
        }
        return pages;
    }

    /**
     * 整体重渲染（模型 → DOM）
     *
     * 会尽量保留光标位置：先记下当前 `(块 id, 偏移)`，重建后再按 id 找回落点。
     * 撤销 / 重做 / 导入都走这里。
     */
    public render(): void {
        const caret = this.captureCaret();
        this.rendering = true;
        try {
            this.blocksEl.textContent = '';
            const blocks = this.store.blocks;
            for (let i = 0; i < blocks.length; i++) {
                this.blocksEl.appendChild(this.renderBlock(blocks[i]));
            }
            if (blocks.length === 0) {
                this.blocksEl.appendChild(this.renderBlock(createBlock()));
            }
        } finally {
            this.rendering = false;
        }
        if (caret) this.restoreCaret(caret);
        // 重渲染把分页留白也清掉了：内容一变就得重排一次页
        this.schedulePaginate();
    }

    // ─────────────────────────── 分页 ───────────────────────────

    /**
     * 安排一次分页重排（同一帧内多次编辑只排一次）
     *
     * 分页要逐个块读 `getBoundingClientRect`，一次 O(块数) 的强制重排；
     * 打字时每一帧都会同步一次模型，跟着排页会把输入延迟压到不可用的程度。
     * 按帧合并后，观感不变（人眼看不出一帧的差），开销降到 1/N。
     */
    private schedulePaginate(): void {
        if (this.paginateFrame !== null) return;
        const run = (): void => {
            this.paginateFrame = null;
            this.paginate();
        };
        if (typeof requestAnimationFrame === 'function') this.paginateFrame = requestAnimationFrame(run);
        else this.paginateFrame = window.setTimeout(run, 0) as unknown as number;
    }

    /**
     * 按当前纸张把内容排进若干页
     *
     * ## 做法：留白块（gap）而不是拆 DOM
     *
     * 真正的分页要「把第 2 页的开头挪到下一页的顶部」。把块搬进第二个容器会破坏
     * 「一个 contenteditable 承载全部正文」的前提（光标跨越容器时浏览器行为不稳）；
     * 因此这里**只插入留白**：在每个跨页块前面插一个 `data-page-gap` 的空元素，
     * 高度正好把它顶到下一页的正文起点。块本身不动，模型与光标都不受影响。
     *
     * ## 为什么量的是缩放后的矩形再除回去
     *
     * `getBoundingClientRect` 返回的是**已缩放**的坐标，除以 zoom 才回到版面坐标；
     * 否则「放大看 → 重排页」会得到与 100% 时完全不同的分页结果。
     *
     * ## 顺序
     *
     * 先摘掉旧留白再量（旧留白会把后续块往下推，量出来的是「已经排过页」的位置），
     * 量完再统一插：插入本身会改变后续块的坐标，用「插入前的位置 + 累计位移」推算，
     * 避免边量边改的反馈环。
     * @returns 总页数
     */
    public paginate(): number {
        const { height, margin } = this.pageGeometry();
        const z = this.zoomValue || 1;
        const children = Array.prototype.slice.call(this.blocksEl.children) as HTMLElement[];
        // ① 摘掉上一轮的分页留白
        for (let i = children.length - 1; i >= 0; i--) {
            if (children[i].hasAttribute(DOC_PAGE_GAP_ATTR)) children[i].remove();
        }
        const blocks = Array.prototype.slice.call(this.blocksEl.children) as HTMLElement[];
        if (blocks.length === 0) {
            this.pageCountValue = 1;
            this.applyFrame();
            return 1;
        }
        const top0 = this.blocksEl.getBoundingClientRect().top;
        // 【超大文档退化处理】逐块量一次是一个 O(块数) 的强制重排，几千块时打字会明显发涩。
        // 这时只回答「一共几页」（量最后一块的底部即可），不做逐块插留白 ——
        // 屏幕上少几条分页线不影响编辑，页数与打印仍然正确。
        if (blocks.length > DOC_MAX_PAGINATE_BLOCKS) {
            const last = blocks[blocks.length - 1].getBoundingClientRect();
            const total = (last.bottom - top0) / z;
            const pages = Math.max(1, Math.floor((total + margin) / height) + 1);
            this.pageCountValue = pages;
            this.page.style.minHeight = `${height * pages}px`;
            this.applyFrame();
            return pages;
        }
        const tops: number[] = [];
        const bottoms: number[] = [];
        for (let i = 0; i < blocks.length; i++) {
            const rect = blocks[i].getBoundingClientRect();
            tops.push((rect.top - top0) / z);
            bottoms.push((rect.bottom - top0) / z);
        }

        // ② 决定每个块落在第几页（跨页的块整体下推，与 Word 的「段中不分页」同口径）
        //    版面坐标 → 页号：整叠纸是「一页接一页」排的，第 n 页的纸顶在 n×纸高处，
        //    它的正文起点再往下 margin；因此**不是**除以正文区高度（那会让第二页
        //    的正文顶到上一页页脚的位置，页与页之间少掉两个页边距）。
        const pageOf = (y: number): number => Math.floor((y + margin) / height);
        let page = 0;
        let shift = 0;
        for (let i = 0; i < blocks.length; i++) {
            const start = pageOf(tops[i]);
            const end = pageOf(Math.max(bottoms[i] - 1, tops[i]));
            let target = Math.max(page, start);
            if (end > target) target = end;
            if (target <= page) continue;
            // 需要把该块顶到 target 页的正文起点：留白 = 目标位置 − 当前（含已插留白）位置
            const gap = target * height - (tops[i] + shift);
            if (gap > 0) {
                const spacer = createDiv({ className: 'doc-page-gap' });
                spacer.setAttribute(DOC_PAGE_GAP_ATTR, String(target));
                spacer.setAttribute('contenteditable', 'false');
                spacer.style.height = `${gap}px`;
                // 分页线画在**上一页纸的底边**上（留白的顶端再往上 margin 处），
                // 交给样式表用这个变量定位，省得把「纸的边距」再算一遍
                spacer.style.setProperty('--doc-break-offset', `${margin}px`);
                this.blocksEl.insertBefore(spacer, blocks[i]);
                shift += gap;
            }
            page = target;
        }
        const total = bottoms.length > 0 ? bottoms[bottoms.length - 1] + shift : 0;
        const pages = Math.max(1, pageOf(Math.max(total - 1, 0)) + 1);
        this.pageCountValue = pages;
        this.page.style.minHeight = `${Math.max(height, height * pages)}px`;
        this.applyFrame();
        return pages;
    }

    /**
     * 让正文获得焦点
     */
    public focus(): void {
        if (document.activeElement === this.blocksEl) return;
        this.blocksEl.focus();
        const first = this.blocksEl.firstElementChild as HTMLElement | null;
        if (first) this.placeCaretAtStart(first);
    }

    /**
     * 切换「格式刷已拿起」的视觉反馈
     *
     * 只加一个类名，光标形状交给样式表：刷子生效时鼠标要变成刷子，
     * 否则用户无法区分「点了没生效」和「正在等下一次选择」。
     * @param active - 格式刷是否在生效中
     */
    public setBrushing(active: boolean): void {
        this.blocksEl.classList.toggle('doc-brushing', !!active);
    }

    // ─────────────────────────── 渲染 ───────────────────────────

    /**
     * 把块级对齐写到元素上
     *
     * 「分散对齐」不是单一 CSS 值：`text-align: justify` 对最后一行无效，
     * 必须再叠加 `text-align-last: justify`。切回其它对齐时要**显式清掉**后者，
     * 否则残留的 `text-align-last` 会让「两端对齐」看起来像「分散对齐」。
     * @param el - 块元素
     * @param align - 对齐方式
     */
    private applyBlockAlign(el: HTMLElement, align: DocAlign): void {
        if (align === 'distribute') {
            el.style.textAlign = 'justify';
            el.style.textAlignLast = 'justify';
            return;
        }
        el.style.textAlign = align;
        el.style.textAlignLast = '';
    }

    /**
     * 从元素上读回块级对齐
     * @param el - 块元素
     * @returns 对齐方式；无显式设置时为 undefined
     */
    private readBlockAlign(el: HTMLElement): DocAlign | undefined {
        const align = (el.style.textAlign || '') as DocAlign | '';
        if (align !== 'left' && align !== 'center' && align !== 'right' && align !== 'justify') return undefined;
        if (align === 'justify' && el.style.textAlignLast === 'justify') return 'distribute';
        return align;
    }

    /**
     * 渲染单个块
     * @param block - 块
     * @returns 块元素
     */
    private renderBlock(block: DocBlock): HTMLElement {
        if (block.type === 'table') return this.renderTable(block);

        const el = createDiv({ className: 'doc-block' });
        el.setAttribute('data-id', block.id);
        el.setAttribute('data-type', block.type);
        if (block.level !== undefined) el.setAttribute('data-level', String(block.level));
        if (block.listStyle) el.setAttribute('data-list', block.listStyle);
        if (block.indent) {
            el.setAttribute('data-indent', String(block.indent));
            // 缩进是**块级**属性：必须在这里还原。只写 `data-indent` 不写样式的话，
            // 缩进只在当次点击里有效（那时是命令直接改的 DOM），
            // 一旦重渲染（撤销、切视图、重新打开）就整段弹回左边。
            el.style.marginLeft = `${block.indent * DOC_INDENT_STEP_PX}px`;
        }
        if (block.align) this.applyBlockAlign(el, block.align);
        // 块内边距承担列表缩进：用 CSS 变量把「层级」传给样式表，
        // 而不是在 DOM 里插入项目符号文本（插文本会污染模型）
        if (block.type === 'listItem') {
            el.style.paddingLeft = `${(block.level ?? 0) * 24 + 24}px`;
        }
        if (block.type === 'quote') el.classList.add('doc-quote');

        switch (block.type) {
            case 'heading':
                // 标题字号跟着级别走，但**不写进 Char**：级别是块级语义，
                // 写进字符样式会让「改级别」变成「逐字符改字号」，也污染 Markdown 往返
                el.classList.add('doc-heading');
                el.setAttribute('data-level', String(Math.max(1, Math.min(6, block.level ?? 1))));
                break;
            case 'code':
                el.classList.add('doc-code');
                break;
            case 'hr': {
                el.classList.add('doc-hr');
                el.setAttribute('contenteditable', 'false');
                const line = createDiv({ className: 'doc-hr-line' });
                el.appendChild(line);
                return el;
            }
            case 'image': {
                el.classList.add('doc-image');
                el.textContent = block.imageRef ? `[图片 ${block.imageRef.assetId}]` : '[图片]';
                return el;
            }
            default:
                break;
        }

        if (block.type === 'code' && block.children && block.children.length > 0) {
            for (let i = 0; i < block.children.length; i++) {
                if (i > 0) el.appendChild(document.createElement('br'));
                el.appendChild(this.renderInline(block.children[i], true));
            }
            return el;
        }
        el.appendChild(this.renderInline(block, block.type === 'code'));
        return el;
    }

    /**
     * 渲染表格块
     * @param block - 表格块
     * @returns 表格元素
     */
    private renderTable(block: DocBlock): HTMLElement {
        const wrapper = createDiv({ className: 'doc-block doc-table' });
        wrapper.setAttribute('data-id', block.id);
        wrapper.setAttribute('data-type', 'table');
        const table = document.createElement('table');
        const body = document.createElement('tbody');
        const rows = block.rows ?? [];
        for (let r = 0; r < rows.length; r++) {
            const tr = document.createElement('tr');
            const cells = rows[r] ?? [];
            for (let c = 0; c < cells.length; c++) {
                const td = document.createElement('td');
                const inner = cells[c].children && cells[c].children.length > 0
                    ? cells[c].children
                    : [cells[c]];
                for (let i = 0; i < inner.length; i++) {
                    const p = createDiv({ className: 'doc-cell-line' });
                    p.appendChild(this.renderInline(inner[i], false));
                    td.appendChild(p);
                }
                tr.appendChild(td);
            }
            body.appendChild(tr);
        }
        table.appendChild(body);
        wrapper.appendChild(table);
        return wrapper;
    }

    /**
     * 把块的行内内容渲染成 DOM（按样式切段）
     * @param block - 块
     * @param mono - 是否强制等宽（代码块）
     * @returns 片段
     */
    private renderInline(block: DocBlock, mono: boolean): DocumentFragment {
        const fragment = document.createDocumentFragment();
        const chars = block.chars ?? [];
        let index = 0;
        while (index < chars.length) {
            const span = this.spanAt(block, index);
            const key = runKey(chars[index], span, mono);
            let end = index + 1;
            while (end < chars.length && runKey(chars[end], this.spanAt(block, end), mono) === key) {
                end += 1;
            }
            let text = '';
            for (let i = index; i < end; i++) text += String(chars[i].char ?? '');

            const holder = document.createElement('span');
            this.applyRunStyle(holder, chars[index], span, mono);
            if (text.indexOf('\n') >= 0) {
                const lines = text.split('\n');
                for (let i = 0; i < lines.length; i++) {
                    if (i > 0) holder.appendChild(document.createElement('br'));
                    if (lines[i]) holder.appendChild(document.createTextNode(lines[i]));
                }
            } else {
                holder.appendChild(document.createTextNode(text));
            }
            fragment.appendChild(holder);
            index = end;
        }
        // 保证每个块都有可落脚的内容节点：空块也要能被点中并输入
        if (chars.length === 0) fragment.appendChild(document.createTextNode(''));
        return fragment;
    }

    /**
     * 把字符样式写到 span 上
     * @param el - 目标元素
     * @param char - 样式来源
     * @param span - 扩展区间
     * @param mono - 是否等宽
     */
    private applyRunStyle(el: HTMLElement, char: Char, span: DocSpan | null, mono: boolean): void {
        if (mono) {
            el.style.fontFamily = 'Consolas, "Courier New", monospace';
        } else if (char.fontFamily) {
            el.style.fontFamily = `"${char.fontFamily}", ${DEFAULT_DOC_FONT_FAMILY}`;
        }
        if (char.fontSize) el.style.fontSize = `${char.fontSize}px`;
        if (char.fontWeight) el.style.fontWeight = 'bold';
        if (char.fontStyle) el.style.fontStyle = 'italic';
        const decorations: string[] = [];
        if (char.underline) decorations.push('underline');
        if (char.strikethrough) decorations.push('line-through');
        if (decorations.length > 0) el.style.textDecoration = decorations.join(' ');
        if (char.fontColor) el.style.color = char.fontColor;
        if (span?.highlight) el.style.backgroundColor = span.highlight;
        // 上下标只改基线位置，**不改字号**：改字号会被同步回模型，
        // 使「导出 → 再导入」一次比一次小
        if (span?.vertAlign) el.style.verticalAlign = span.vertAlign === 'superscript' ? 'super' : 'sub';
        if (span?.link) {
            // 链接只上色不加下划线：下划线是装饰，同步时会被读成「字符带下划线」，
            // 于是每往返一次 Markdown 就多一层 <u>
            el.setAttribute('data-link', span.link);
            el.style.color = char.fontColor ?? '#1a73e8';
            el.style.cursor = 'pointer';
        }
    }

    /**
     * 取某位置命中的扩展区间
     * @param block - 块
     * @param index - 码点下标
     * @returns 区间或 null
     */
    private spanAt(block: DocBlock, index: number): DocSpan | null {
        if (!block.spans) return null;
        for (let i = 0; i < block.spans.length; i++) {
            const span = block.spans[i];
            if (index >= span.start && index < span.end) return span;
        }
        return null;
    }

    // ─────────────────────────── 编辑同步 ───────────────────────────

    /**
     * 绑定编辑事件
     */
    private bindEvents(): void {
        // 输入 → 合并到一帧后同步（避免每个字符都走一次 getComputedStyle 全树遍历）
        this.blocksEl.addEventListener('input', () => {
            if (this.rendering) return;
            this.scheduleSync();
        });

        // 粘贴一律按纯文本处理：带 HTML 的粘贴会把外部样式与标签带进模型，
        // 而我们的模型只认 chars + spans，多出来的结构会在下一次同步时被静默丢弃
        this.blocksEl.addEventListener('paste', (event) => {
            const data = event.clipboardData;
            if (!data) return;
            event.preventDefault();
            const text = data.getData('text/plain');
            this.insertText(text);
        });

        // 折叠光标上的待生效样式：焦点/选区一变就作废（否则会「莫名加粗下一段」）
        document.addEventListener('selectionchange', () => {
            const selection = window.getSelection();
            if (!selection || selection.rangeCount === 0 || selection.isCollapsed === false) {
                this.pendingStyle = null;
                return;
            }
            if (!this.blocksEl.contains(selection.anchorNode)) {
                this.pendingStyle = null;
                return;
            }
        });

        // 折叠光标 + 有待生效样式时，自己插入文本（保证拿到刚设的样式）
        this.blocksEl.addEventListener('beforeinput', (event) => {
            const input = event as InputEvent;
            if (!this.pendingStyle || input.inputType !== 'insertText' || input.isComposing) return;
            if (input.data == null) return;
            event.preventDefault();
            this.insertText(input.data, this.pendingStyle);
        });

        // 【单击 = 取消选区 + 把光标放到点击处】
        // 只有 `blocksEl` 是 contenteditable：点在纸张页边距、纸张空白处时浏览器**不会动选区**，
        // 于是「Ctrl+A 全选之后点一下页边」，蓝色高亮一直挂着，看着像点不动。
        // 判定「单击」而不是「拖选」：按下与抬起之间位移超过 3px 就认为用户在拖选，保留选区。
        let downAt: { x: number; y: number } | null = null;
        this.scroller.addEventListener('mousedown', (event: MouseEvent) => {
            downAt = event.button === 0 ? { x: event.clientX, y: event.clientY } : null;
        });
        this.scroller.addEventListener('mouseup', (event: MouseEvent) => {
            const down = downAt;
            downAt = null;
            if (!down || event.button !== 0) return;
            const dragged = Math.abs(event.clientX - down.x) > 3 || Math.abs(event.clientY - down.y) > 3;
            if (dragged) return;
            const selection = window.getSelection();
            if (selection && selection.isCollapsed) return;   // 浏览器已经把光标放好了
            this.placeCaretAtPoint(event.clientX, event.clientY);
        });

        // Ctrl/⌘ + B/I/U 与工具栏走同一条命令通道（由 DocEditHost 接管键盘，见 docWorkspace）
    }

    /**
     * 清掉当前选区并把光标放到指定屏幕坐标处
     *
     * 坐标可能落在纸张页边距、块与块之间的缝隙上 —— 那里没有文本节点，
     * 浏览器不会自己放光标，需要我们用 `caretRangeFromPoint` 找最近的落点。
     * @param x - 视口横坐标
     * @param y - 视口纵坐标
     */
    private placeCaretAtPoint(x: number, y: number): void {
        const selection = window.getSelection();
        if (!selection) return;
        const range = this.caretRangeAt(x, y);
        selection.removeAllRanges();
        if (range) {
            range.collapse(true);
            selection.addRange(range);
        }
        if (document.activeElement !== this.blocksEl) this.blocksEl.focus({ preventScroll: true });
    }

    /**
     * 取某个屏幕坐标对应的插入点
     * @param x - 视口横坐标
     * @param y - 视口纵坐标
     * @returns 插入点；找不到时为 null（此时只清选区，不设光标）
     */
    private caretRangeAt(x: number, y: number): Range | null {
        // 两个 API 都不是标准成员：`caretRangeFromPoint`（WebKit/Blink 全系）与
        // `caretPositionFromPoint`（Firefox），因此按能力探测，不用类型断言之外的方式绕。
        const doc = document as Document & {
            caretRangeFromPoint?: (x: number, y: number) => Range | null;
            caretPositionFromPoint?: (x: number, y: number) => { offsetNode: Node; offset: number } | null;
        };
        if (typeof doc.caretRangeFromPoint === 'function') {
            const range = doc.caretRangeFromPoint(x, y);
            if (range && this.blocksEl.contains(range.startContainer)) return range;
        }
        if (typeof doc.caretPositionFromPoint === 'function') {
            const position = doc.caretPositionFromPoint(x, y);
            if (position && this.blocksEl.contains(position.offsetNode)) {
                try {
                    const range = document.createRange();
                    range.setStart(position.offsetNode, position.offset);
                    return range;
                } catch {
                    return null;
                }
            }
        }
        // 兜底：按纵坐标找最近的块 —— 点在某块下半部分就落它的末尾，否则落开头
        const children = this.blocksEl.children;
        if (children.length === 0) return null;
        let target = children[0] as HTMLElement;
        let toEnd = false;
        for (let i = 0; i < children.length; i++) {
            const el = children[i] as HTMLElement;
            const rect = el.getBoundingClientRect();
            if (rect.top > y) break;
            target = el;
            toEnd = y > rect.bottom;
        }
        const range = document.createRange();
        range.selectNodeContents(target);
        range.collapse(toEnd);
        return range;
    }

    /**
     * 安排一次同步（同一帧内多次 input 只同步一次）
     */
    private scheduleSync(): void {
        if (this.syncFrame !== null) return;
        this.syncFrame = window.requestAnimationFrame(() => {
            this.syncFrame = null;
            this.sync('输入', 'coalesce', false);
        });
    }

    /**
     * 把 DOM 上的编辑结果同步回块树
     *
     * **只在 DOM 真正含有块时才同步**：`render()` 过程中会短暂清空容器，
     * 若此时触发同步会把整篇文档读成空文档（表现为「撤销一次内容全没了」）。
     * @param label - 撤销标签
     * @param mode - 记账方式
     * @param reset - 同步后是否重渲染（样式命令改的是 DOM，不需要重渲染）
     */
    public sync(label: string = '编辑', mode: 'coalesce' | 'single' | 'none' = 'coalesce', reset: boolean = false): void {
        if (this.rendering) return;
        const blocks: DocBlock[] = [];
        const children = this.blocksEl.children;
        for (let i = 0; i < children.length; i++) {
            const el = children[i] as HTMLElement;
            if (!el || !el.getAttribute) continue;
            // 分页留白不是内容：读进模型会变成一段空段落（表现为「每存一次多一行」）
            if (el.hasAttribute(DOC_PAGE_GAP_ATTR)) continue;
            const block = this.readBlock(el);
            if (block) blocks.push(block);
        }
        if (blocks.length === 0) return;
        this.store.update(blocks, label, mode);
        if (reset) this.render();
        // 内容高度可能变了（打字换行、插入图片）：重排一次页
        this.schedulePaginate();
    }

    /**
     * 从 DOM 元素读出一个块
     * @param el - 块元素
     * @returns 块
     */
    private readBlock(el: HTMLElement): DocBlock | null {
        const type = (el.getAttribute('data-type') ?? 'paragraph') as DocBlock['type'];
        const id = el.getAttribute('data-id') ?? nextBlockId();
        if (type === 'table') return this.readTable(el, id);
        if (type === 'hr') return { id, type: 'hr', chars: [] };

        const block: DocBlock = { id, type: type === 'listItem' || type === 'heading' || type === 'quote' || type === 'code' ? type : 'paragraph', chars: [] };
        if (type === 'heading') block.level = Math.max(1, Math.min(6, parseInt(el.getAttribute('data-level') ?? '1', 10) || 1));
        if (type === 'listItem') {
            block.level = Math.max(0, Math.min(4, parseInt(el.getAttribute('data-level') ?? '0', 10) || 0));
            block.listStyle = el.getAttribute('data-list') === 'ordered' ? 'ordered' : 'bullet';
        }
        if (type === 'heading' || type === 'listItem' || type === 'quote') {
            const level = parseInt(el.getAttribute('data-level') ?? '', 10);
            if (type === 'listItem' && Number.isFinite(level)) block.level = Math.max(0, Math.min(4, level));
        }
        const indent = parseInt(el.getAttribute('data-indent') ?? '', 10);
        if (Number.isFinite(indent) && indent > 0) block.indent = indent;

        const align = this.readBlockAlign(el);
        if (align) block.align = align;

        // 代码块：按 <br> 拆回子块，保留多行
        const baseline = readComputedStyle(el);
        if (block.type === 'code') {
            block.chars = [];
            this.readInline(el, block.chars, block, baseline);
            block.children = splitByNewline(block.chars).map((chars) => ({ id: nextBlockId(), type: 'paragraph', chars }));
            return block;
        }
        this.readInline(el, block.chars, block, baseline);
        return block;
    }

    /**
     * 读取表格块
     * @param el - 表格容器
     * @param id - 块 id
     * @returns 表格块
     */
    private readTable(el: HTMLElement, id: string): DocBlock {
        const block: DocBlock = { id, type: 'table', chars: [], rows: [] };
        const table = el.querySelector('table');
        if (!table) return block;
        const trs = table.querySelectorAll('tr');
        const rows: DocBlock[][] = [];
        for (let r = 0; r < trs.length; r++) {
            const cells = trs[r].querySelectorAll('td,th');
            const row: DocBlock[] = [];
            for (let c = 0; c < cells.length; c++) {
                const cell: DocBlock = { id: nextBlockId(), type: 'paragraph', chars: [], children: [] };
                const lines = cells[c].querySelectorAll('.doc-cell-line');
                const target = lines.length > 0 ? Array.prototype.slice.call(lines) : [cells[c]];
                for (let i = 0; i < target.length; i++) {
                    const child: DocBlock = { id: nextBlockId(), type: 'paragraph', chars: [] };
                    const lineEl = target[i] as HTMLElement;
                    this.readInline(lineEl, child.chars, child, readComputedStyle(lineEl));
                    (cell.children as DocBlock[]).push(child);
                }
                row.push(cell);
            }
            rows.push(row);
        }
        block.rows = rows;
        return block;
    }

    /**
     * 读取元素内的行内内容为 `Char[]`（并可登记扩展区间）
     * @param root - 根元素
     * @param out - 输出字符数组
     * @param block - 输出块（用于登记 spans）
     */
    private readInline(
        root: HTMLElement,
        out: Char[],
        block: DocBlock,
        baseline?: CharStylePatch
    ): void {
        const base = baseline ?? { fontFamily: DEFAULT_DOC_FONT_FAMILY, fontSize: DEFAULT_DOC_FONT_SIZE };
        const walk = (node: Node): void => {
            if (node.nodeType === 3) {
                const text = (node as Text).nodeValue ?? '';
                if (!text) return;
                const parent = (node.parentElement ?? root) as HTMLElement;
                // 【只记「相对块基线的增量」】块的 CSS 会带来加粗（标题）、等宽（代码）、
                // 默认字体字号（正文）—— 这些是**呈现**，不是**数据**。若照抄计算样式，
                // 标题的每个字符都会被存成 bold，导出 Markdown 就变成 `# **标题**`，
                // 往返一次多一层加粗。基线一减，模型里只剩用户真正设过的样式。
                const style = diffStyle(readComputedStyle(parent), base);
                const start = out.length;
                const created = charsFromText(text);
                for (let i = 0; i < created.length; i++) {
                    writeCharStyle(created[i], style);
                    out.push(created[i]);
                }
                const end = out.length;
                if (end > start) this.recordSpan(parent, block, start, end);
                return;
            }
            if (node.nodeType === 1) {
                const el = node as HTMLElement;
                if (el.tagName === 'BR') {
                    const char = Char.of('\n');
                    writeCharStyle(char, diffStyle(readComputedStyle(el.parentElement ?? root), base));
                    out.push(char);
                    return;
                }
                const childNodes = node.childNodes;
                for (let i = 0; i < childNodes.length; i++) walk(childNodes[i]);
            }
        };
        const childNodes = root.childNodes;
        for (let i = 0; i < childNodes.length; i++) walk(childNodes[i]);
        // 浏览器在空行里会插一个占位 <br>：它不该被当成换行
        if (out.length === 1 && out[0].char === '\n') out.length = 0;
    }

    /**
     * 从 DOM 元素上读出扩展区间（链接 / 高亮 / 上下标）并登记
     * @param el - 元素
     * @param block - 目标块
     * @param start - 起始下标
     * @param end - 结束下标
     */
    private recordSpan(el: HTMLElement, block: DocBlock, start: number, end: number): void {
        const link = el.getAttribute('data-link') ?? undefined;
        const background = el.style.backgroundColor || undefined;
        const vertical = el.style.verticalAlign;
        const vertAlign = vertical === 'super' ? 'superscript' : vertical === 'sub' ? 'subscript' : undefined;
        const highlight = background && !/rgba\(0, 0, 0, 0\)|transparent/i.test(background) ? rgbToHex(background) : undefined;
        if (!link && !highlight && !vertAlign) return;
        if (!block.spans) block.spans = [];
        const span: DocSpan = { start, end };
        if (link) span.link = link;
        if (highlight) span.highlight = highlight;
        if (vertAlign) span.vertAlign = vertAlign;
        block.spans.push(span);
    }

    // ─────────────────────────── 选区状态 ───────────────────────────

    /**
     * 取当前选区的字符样式（供工具栏高亮按钮）
     * @returns 样式补丁（未命中时为默认值）
     */
    public currentCharStyle(): CharStylePatch {
        const selection = window.getSelection();
        const node = selection && selection.rangeCount > 0 ? selection.getRangeAt(0).startContainer : null;
        let el: HTMLElement | null = null;
        if (node && this.blocksEl.contains(node)) {
            el = (node.nodeType === 1 ? node : node.parentElement) as HTMLElement | null;
        }
        if (!el) {
            // 没有焦点：用首块首字符做代表（与表格「读活动格」口径一致）
            const first = this.store.blocks[0];
            return first ? styleOfChar(first.chars[0]) : {};
        }
        // 【选区内按「并集」读】只看起点会让混排选区永远读成「没下划线」：
        // 起点字符没有下划线 → 按钮显示未激活 → 点击又是「加上下划线」，
        // 于是「选中既有下划线又有普通文字时，永远去不掉线条」。
        // 布尔类样式取并集（任一字符有即视为有），再点一次就能整段去掉；
        // 字体 / 字号 / 颜色取起点（没有「并集」语义，取第一个读到的更可预期）。
        const base = readComputedStyle(el);
        if (!selection || selection.isCollapsed) return base;
        return this.unionStyleOverRange(selection.getRangeAt(0), base);
    }

    /**
     * 在选区范围内取布尔样式的并集
     * @param range - 选区
     * @param base - 起点样式（字体 / 字号 / 颜色沿用它）
     * @returns 并集后的样式
     */
    private unionStyleOverRange(range: Range, base: CharStylePatch): CharStylePatch {
        const merged: CharStylePatch = { ...base };
        const nodes = collectTextNodes(this.blocksEl, range);
        for (let i = 0; i < nodes.length; i++) {
            const parent = nodes[i].parentElement;
            if (!parent || !this.blocksEl.contains(parent)) continue;
            const style = readComputedStyle(parent);
            if (style.bold) merged.bold = true;
            if (style.italic) merged.italic = true;
            if (style.underline) merged.underline = true;
            if (style.strike) merged.strike = true;
        }
        return merged;
    }

    /**
     * 取当前选区的段落状态
     * @returns 段落状态
     */
    public currentParaState(): DocParaState {
        const selection = window.getSelection();
        const node = selection && selection.rangeCount > 0 ? selection.getRangeAt(0).startContainer : null;
        const host = node ? this.closestBlock(node) : null;
        if (!host) {
            return { type: 'paragraph', level: 0, listStyle: 'none', align: 'left', indent: 0 };
        }
        const type = (host.getAttribute('data-type') ?? 'paragraph') as DocBlock['type'];
        return {
            type,
            level: parseInt(host.getAttribute('data-level') ?? '0', 10) || 0,
            listStyle: (host.getAttribute('data-list') as DocListStyle) ?? 'none',
            indent: parseInt(host.getAttribute('data-indent') ?? '0', 10) || 0,
            align: this.readBlockAlign(host) ?? 'left',
        };
    }

    /**
     * 找光标所在块
     * @param node - 选区节点
     * @returns 块元素或 null
     */
    private closestBlock(node: Node): HTMLElement | null {
        let current: Node | null = node;
        while (current && current !== this.blocksEl) {
            if (current.nodeType === 1) {
                const el = current as HTMLElement;
                if (el.getAttribute && el.getAttribute('data-type')) return el;
            }
            current = current.parentNode;
        }
        return null;
    }

    // ─────────────────────────── 样式命令 ───────────────────────────

    /**
     * 对当前选区应用字符样式
     *
     * 折叠光标时记成「待生效样式」，下次输入时套用（Word 的行为）；
     * 有选区时直接包裹/改写选区内的文本节点。
     * @param patch - 样式补丁
     * @returns 是否生效
     */
    public applyCharStyle(patch: CharStylePatch): boolean {
        const selection = window.getSelection();
        if (!selection || selection.rangeCount === 0) return false;
        const range = selection.getRangeAt(0);
        if (!this.blocksEl.contains(range.commonAncestorContainer)) return false;

        if (selection.isCollapsed) {
            this.pendingStyle = { ...(this.pendingStyle ?? {}), ...patch };
            return true;
        }
        this.pendingStyle = null;
        this.styleRange(range, patch);
        this.sync('字体样式', 'single', false);
        return true;
    }

    /**
     * 给选区内的文本节点套样式（拆分文本节点后逐段包裹）
     * @param range - 选区
     * @param patch - 样式补丁
     */
    private styleRange(range: Range, patch: CharStylePatch): void {
        const nodes = collectTextNodes(this.blocksEl, range);
        const created: HTMLElement[] = [];
        // 「去掉下划线 / 删除线」要先处理：线条来自**祖先**时，包一层
        // `text-decoration: none` 是压不住的（见 {@link clearAncestorDecorations}）
        const off = { underline: patch.underline === false, strike: patch.strike === false };
        const needSplit = off.underline || off.strike;
        for (let i = 0; i < nodes.length; i++) {
            const node = nodes[i];
            let start = 0;
            let end = node.nodeValue ? node.nodeValue.length : 0;
            if (node === range.startContainer) start = range.startOffset;
            if (node === range.endContainer) end = range.endOffset;
            if (end <= start) continue;

            let target: Text = node;
            if (end < (node.nodeValue ?? '').length) target.splitText(end);
            if (start > 0) target = target.splitText(start);
            if (!target.nodeValue) continue;

            // 【先切断祖先的线条，再包新样式】否则新 span 的 `text-decoration: none`
            // 会被祖先的 underline 盖掉，表现为「下划线点不掉」
            if (needSplit) this.clearAncestorDecorations(target, off);

            const parent = (target.parentElement ?? this.blocksEl) as HTMLElement;
            const wrapper = document.createElement('span');
            applyStyleElement(wrapper, parent, patch);
            const parentNode = target.parentNode;
            if (!parentNode) continue;
            parentNode.insertBefore(wrapper, target);
            wrapper.appendChild(target);
            created.push(wrapper);
        }
        if (created.length > 0) {
            selectElements(created[0], created[created.length - 1]);
        }
    }

    /**
     * 去掉「来自祖先」的下划线 / 删除线
     *
     * ## 为什么包一层 span 不管用
     *
     * `text-decoration` 是**向下传播**的：祖先画的那条线会一路穿过后代，
     * CSS 里没有任何「后代声明 `none` 就能擦掉祖先线条」的办法 ——
     * 于是「选中已带下划线的文字再点一次下划线」看起来毫无反应
     * （加粗 / 倾斜因为是真正的继承属性，覆盖一次就生效，所以只有线条会出这个 bug）。
     *
     * ## 做法：沿着祖先链把「带线条的那一层」就地劈成三段
     *
     * ```
     *   <span style="text-decoration:underline"> 前半 〔选中文字〕 后半 </span>
     *        ↓
     *   <span …underline>前半</span><span …去掉线条>〔选中文字〕</span><span …underline>后半</span>
     * ```
     *
     * 左右两份是**浅克隆**（属性与内联样式原样带走，线条保留），中间那份只摘掉被关掉的线条。
     * 这样才能做到「只去掉选中的那一段」——把整层的线条直接摘掉会连选区外的文字一起改。
     *
     * 只走到块元素为止：再往上是块容器，劈开它等于把一段拆成两个块，模型会多出一行。
     * @param node - 选中的文本节点
     * @param off - 要关掉哪几种线条
     */
    private clearAncestorDecorations(node: Node, off: { underline: boolean; strike: boolean }): void {
        const host = this.closestBlock(node) ?? this.blocksEl;
        let current: HTMLElement | null = node.parentElement;
        while (current && current !== host && current !== this.blocksEl) {
            const parent: HTMLElement | null = current.parentElement;
            if (!parent) break;
            const dec = readDecorations(current);
            const hit = (off.underline && dec.underline) || (off.strike && dec.strike);
            if (!hit) {
                current = parent;
                continue;
            }
            const children = Array.prototype.slice.call(current.childNodes) as Node[];
            const index = children.indexOf(node);
            if (index >= 0) {
                const before = children.slice(0, index);
                const after = children.slice(index + 1);
                if (before.length > 0) {
                    const left = current.cloneNode(false) as HTMLElement;
                    for (let i = 0; i < before.length; i++) left.appendChild(before[i]);
                    parent.insertBefore(left, current);
                }
                if (after.length > 0) {
                    const right = current.cloneNode(false) as HTMLElement;
                    for (let i = 0; i < after.length; i++) right.appendChild(after[i]);
                    parent.insertBefore(right, current.nextSibling);
                }
            }
            // 中间那份（current 自己）留下 node，并摘掉被关掉的线条
            writeDecorations(current, {
                underline: dec.underline && !off.underline,
                strike: dec.strike && !off.strike,
            });
            current = parent;
        }
    }

    /**
     * 设置段落对齐
     * @param align - 对齐方式
     */
    public setAlign(align: DocAlign): void {
        const hosts = this.selectedBlocks();
        if (hosts.length === 0) return;
        for (let i = 0; i < hosts.length; i++) this.applyBlockAlign(hosts[i], align);
        this.sync('对齐方式', 'single', false);
    }

    /**
     * 切换列表（圆点 / 编号）；已是同种列表时转回普通段落
     * @param style - 目标列表风格
     */
    public toggleList(style: 'bullet' | 'ordered'): void {
        const hosts = this.selectedBlocks();
        if (hosts.length === 0) return;
        const allSame = hosts.every((host) => host.getAttribute('data-list') === style);
        for (let i = 0; i < hosts.length; i++) {
            const host = hosts[i];
            if (allSame) {
                host.removeAttribute('data-list');
                host.setAttribute('data-type', 'paragraph');
                host.style.paddingLeft = '';
            } else {
                host.setAttribute('data-type', 'listItem');
                host.setAttribute('data-list', style);
                const level = parseInt(host.getAttribute('data-level') ?? '0', 10) || 0;
                host.style.paddingLeft = `${level * 24 + 24}px`;
            }
        }
        this.sync('列表', 'single', false);
    }

    /**
     * 设置块类型（段落 / 标题 n / 引用 / 代码）
     * @param type - 块类型
     * @param level - 标题级别
     */
    public setBlockType(type: DocBlock['type'], level: number = 1): void {
        const hosts = this.selectedBlocks();
        if (hosts.length === 0) return;
        for (let i = 0; i < hosts.length; i++) {
            const host = hosts[i];
            host.setAttribute('data-type', type);
            if (type === 'heading') host.setAttribute('data-level', String(level));
            else host.removeAttribute('data-level');
            if (type === 'listItem') {
                host.setAttribute('data-list', 'bullet');
                host.style.paddingLeft = '24px';
            }
        }
        this.sync('段落样式', 'single', false);
    }

    /**
     * 调整缩进
     * @param delta - +1 / -1
     */
    public indentBy(delta: number): void {
        const hosts = this.selectedBlocks();
        if (hosts.length === 0) return;
        for (let i = 0; i < hosts.length; i++) {
            const host = hosts[i];
            const current = parseInt(host.getAttribute('data-indent') ?? '0', 10) || 0;
            const next = Math.max(0, Math.min(8, current + delta));
            host.setAttribute('data-indent', String(next));
            host.style.marginLeft = next > 0 ? `${next * DOC_INDENT_STEP_PX}px` : '';
        }
        this.sync('缩进', 'single', false);
    }

    /**
     * 在当前选区插入纯文本
     * @param text - 文本
     * @param style - 可选的字符样式（折叠光标的待生效样式）
     */
    public insertText(text: string, style?: CharStylePatch): void {
        if (!text) return;
        const selection = window.getSelection();
        if (!selection || selection.rangeCount === 0) return;
        const range = selection.getRangeAt(0);
        if (!this.blocksEl.contains(range.commonAncestorContainer)) return;

        if (!style) {
            // 无待生效样式：交给浏览器（保留继承样式与撤销栈），再同步模型
            range.deleteContents();
            range.insertNode(document.createTextNode(text));
            range.collapse(false);
            selection.removeAllRanges();
            selection.addRange(range);
            this.sync('输入', 'coalesce', false);
            return;
        }

        range.deleteContents();
        const wrapper = document.createElement('span');
        applyStyleElement(wrapper, (range.startContainer.nodeType === 1
            ? range.startContainer as HTMLElement
            : range.startContainer.parentElement ?? this.blocksEl), style);
        // 多行文本拆成「文本 + <br>」
        const lines = text.replace(/\r\n?/g, '\n').split('\n');
        for (let i = 0; i < lines.length; i++) {
            if (i > 0) wrapper.appendChild(document.createElement('br'));
            if (lines[i]) wrapper.appendChild(document.createTextNode(lines[i]));
        }
        range.insertNode(wrapper);
        const after = document.createRange();
        after.setStartAfter(wrapper);
        after.collapse(true);
        selection.removeAllRanges();
        selection.addRange(after);
        this.sync('输入', 'coalesce', false);
    }

    /**
     * 取选区覆盖的块元素（无选区时取光标所在块）
     * @returns 块元素数组
     */
    private selectedBlocks(): HTMLElement[] {
        const selection = window.getSelection();
        if (!selection || selection.rangeCount === 0) return [];
        const range = selection.getRangeAt(0);
        const result: HTMLElement[] = [];
        const children = this.blocksEl.children;
        for (let i = 0; i < children.length; i++) {
            const el = children[i] as HTMLElement;
            if (el.hasAttribute(DOC_PAGE_GAP_ATTR)) continue;
            if (selection.isCollapsed) {
                if (this.closestBlock(range.startContainer) === el) result.push(el);
            } else if (range.intersectsNode(el)) {
                result.push(el);
            }
        }
        return result;
    }

    // ─────────────────────────── 光标保存 / 恢复 ───────────────────────────

    /**
     * 记录当前光标位置（块 id + 块内偏移）
     * @returns 光标位置；无光标时为 null
     */
    public captureCaret(): DocCaret | null {
        const selection = window.getSelection();
        if (!selection || selection.rangeCount === 0) return null;
        const range = selection.getRangeAt(0);
        if (!this.blocksEl.contains(range.startContainer)) return null;
        const host = this.closestBlock(range.startContainer);
        if (!host) return null;
        const blockId = host.getAttribute('data-id');
        if (!blockId) return null;
        return { blockId, offset: offsetInBlock(host, range.startContainer, range.startOffset) };
    }

    /**
     * 恢复光标位置
     * @param caret - 光标位置
     */
    public restoreCaret(caret: DocCaret): void {
        const host = this.findBlockElement(caret.blockId);
        if (!host) return;
        const point = locateInBlock(host, caret.offset);
        if (!point) return;
        const range = document.createRange();
        range.setStart(point.node, point.offset);
        range.collapse(true);
        const selection = window.getSelection();
        if (!selection) return;
        selection.removeAllRanges();
        selection.addRange(range);
    }

    /**
     * 按块 id 找元素
     * @param id - 块 id
     * @returns 元素或 null
     */
    private findBlockElement(id: string): HTMLElement | null {
        const children = this.blocksEl.children;
        for (let i = 0; i < children.length; i++) {
            const el = children[i] as HTMLElement;
            if (el.getAttribute('data-id') === id) return el;
        }
        return null;
    }

    /**
     * 把光标放到块首
     * @param el - 块元素
     */
    private placeCaretAtStart(el: HTMLElement): void {
        const range = document.createRange();
        range.setStart(el, 0);
        range.collapse(true);
        const selection = window.getSelection();
        if (!selection) return;
        selection.removeAllRanges();
        selection.addRange(range);
    }

    /**
     * 存储变化时的响应
     *
     * `sync` 触发的 DOC_CHANGED 不需要重渲染（DOM 就是刚同步的来源），
     * 但撤销 / 导入走的是 DOC_RESET → `render()`，所以这里只在「非本视图发起」时补一次渲染判断。
     * 由于无法区分来源，这里**什么都不做**：`render()` 由 DOC_RESET 独占触发。
     */
    private onStoreChanged(): void {
        // 占位：视图不跟随 DOC_CHANGED 重渲染，避免打字过程中重建 DOM 打断 IME。
        // 需要重渲染的场景（撤销 / 导入 / 切换文档）都会发 DOC_RESET。
    }
}

/**
 * 读一个元素上的线型（内联样式优先，其次计算样式）
 *
 * 之所以优先读内联样式：计算样式会把**祖先传播的线条**也算进来
 * （`getComputedStyle(span).textDecorationLine` 可能是祖先画的），
 * 照它去改就会把「祖先的线」当成「这一层的线」劈开，白切一刀。
 * @param el - 元素
 * @returns 是否有下划线 / 删除线
 */
function readDecorations(el: HTMLElement): { underline: boolean; strike: boolean } {
    const inline = el.style ? `${el.style.textDecoration || ''}` : '';
    const value = inline || `${window.getComputedStyle(el).textDecorationLine || ''}`;
    return {
        underline: value.indexOf('underline') >= 0,
        strike: value.indexOf('line-through') >= 0,
    };
}

/**
 * 把一个元素的线型改写成指定的组合（两种都没有时写 `none`）
 * @param el - 元素
 * @param dec - 目标线型
 */
function writeDecorations(el: HTMLElement, dec: { underline: boolean; strike: boolean }): void {
    const parts: string[] = [];
    if (dec.underline) parts.push('underline');
    if (dec.strike) parts.push('line-through');
    el.style.textDecoration = parts.length > 0 ? parts.join(' ') : 'none';
}

/**
 * 计算 run 的样式签名
 * @param char - 字符
 * @param span - 扩展区间
 * @param mono - 是否等宽
 * @returns 签名
 */
function runKey(char: Char, span: DocSpan | null, mono: boolean): string {
    return [
        mono ? 1 : 0,
        char.fontFamily ?? '',
        char.fontSize ?? '',
        char.fontWeight ? 1 : 0,
        char.fontStyle ? 1 : 0,
        char.underline ? 1 : 0,
        char.strikethrough ? 1 : 0,
        char.fontColor ?? '',
        span?.highlight ?? '',
        span?.link ?? '',
        span?.vertAlign ?? '',
    ].join('|');
}

/**
 * 读取元素的计算样式并归一化成本项目的字符样式
 * @param el - 元素
 * @returns 样式补丁
 */
function readComputedStyle(el: HTMLElement): CharStylePatch {
    const style: CharStylePatch = {};
    const computed = window.getComputedStyle(el);
    const family = (computed.fontFamily || '').split(',')[0].replace(/["']/g, '').trim();
    if (family) style.fontFamily = family;
    const size = parseFloat(computed.fontSize);
    if (Number.isFinite(size) && size > 0) style.fontSize = Math.round(size);
    const weight = computed.fontWeight === 'bold' ? 700 : parseInt(computed.fontWeight || '400', 10);
    if (Number.isFinite(weight) && weight >= 600) style.bold = true;
    if (computed.fontStyle === 'italic' || computed.fontStyle === 'oblique') style.italic = true;
    const decoration = `${computed.textDecorationLine || computed.textDecoration || ''}`;
    if (decoration.indexOf('underline') >= 0) style.underline = true;
    if (decoration.indexOf('line-through') >= 0) style.strike = true;
    if (computed.color) style.color = rgbToHex(computed.color);
    // 字符高亮（油漆桶）：`transparent` / `rgba(0,0,0,0)` 表示「没有底色」，
    // 必须过滤掉 —— 否则每个字符都会被记成 `background: transparent`，
    // 而「有底色」与「透明底」在模型里就再也无法区分（去掉高亮的功能随之失效）。
    const background = computed.backgroundColor || '';
    if (background && !/rgba\(0,\s*0,\s*0,\s*0\)|transparent/i.test(background)) {
        style.background = rgbToHex(background);
    }
    return style;
}

/**
 * 取「相对基线的样式增量」
 *
 * 只保留与块基线不同的字段：块的 CSS（标题加粗、代码等宽、正文字体字号）
 * 属于呈现，不该被写进字符数据。基线为 undefined 的字段用默认值兜底 ——
 * 否则「标题里的正文颜色」会因为基线没记录 color 而被当成差异存下来。
 * @param style - 实际计算样式
 * @param base - 块基线样式
 * @returns 增量样式
 */
function diffStyle(style: CharStylePatch, base: CharStylePatch): CharStylePatch {
    const delta: CharStylePatch = {};
    if (style.fontFamily && style.fontFamily !== base.fontFamily) delta.fontFamily = style.fontFamily;
    if (style.fontSize && style.fontSize !== base.fontSize) delta.fontSize = style.fontSize;
    if (!!style.bold !== !!base.bold) delta.bold = !!style.bold;
    if (!!style.italic !== !!base.italic) delta.italic = !!style.italic;
    if (!!style.underline !== !!base.underline) delta.underline = !!style.underline;
    if (!!style.strike !== !!base.strike) delta.strike = !!style.strike;
    if (style.color && style.color !== base.color) delta.color = style.color;
    // 【高亮不在这里】字符底色属于**区间属性**（`DocSpan.highlight`），由 `recordSpan` 从
    // 元素的内联样式单独登记；写进 `Char` 会让「去掉高亮」无法表达（Char 没有「无底色」位）。
    return delta;
}

/**
 * 从字符反推样式（用于「没有焦点」时的代表值）
 * @param char - 字符
 * @returns 样式补丁
 */
function styleOfChar(char: Char | undefined): CharStylePatch {
    const style: CharStylePatch = {};
    if (!char) return style;
    if (char.fontFamily) style.fontFamily = char.fontFamily;
    if (char.fontSize) style.fontSize = char.fontSize;
    if (char.fontWeight) style.bold = true;
    if (char.fontStyle) style.italic = true;
    if (char.underline) style.underline = true;
    if (char.strikethrough) style.strike = true;
    if (char.fontColor) style.color = char.fontColor;
    return style;
}

/**
 * 把「继承样式 + 补丁」写成一个自洽的行内样式（不依赖父元素）
 * @param wrapper - 目标 span
 * @param source - 样式来源元素
 * @param patch - 样式补丁
 */
function applyStyleElement(wrapper: HTMLElement, source: HTMLElement, patch: CharStylePatch): void {
    const base = readComputedStyle(source);
    const merged: CharStylePatch = { ...base, ...patch };
    if (merged.fontFamily) wrapper.style.fontFamily = `"${merged.fontFamily}", ${DEFAULT_DOC_FONT_FAMILY}`;
    if (merged.fontSize) wrapper.style.fontSize = `${merged.fontSize}px`;
    if (merged.bold !== undefined) wrapper.style.fontWeight = merged.bold ? 'bold' : 'normal';
    if (merged.italic !== undefined) wrapper.style.fontStyle = merged.italic ? 'italic' : 'normal';
    if (merged.underline !== undefined || merged.strike !== undefined) {
        const parts: string[] = [];
        const underline = merged.underline !== undefined ? merged.underline : base.underline;
        const strike = merged.strike !== undefined ? merged.strike : base.strike;
        if (underline) parts.push('underline');
        if (strike) parts.push('line-through');
        wrapper.style.textDecoration = parts.length > 0 ? parts.join(' ') : 'none';
    }
    if (merged.color) wrapper.style.color = merged.color;
    if (merged.background) wrapper.style.backgroundColor = merged.background;
}

/**
 * 选中两个元素之间的内容
 * @param first - 起始元素
 * @param last - 结束元素
 */
function selectElements(first: HTMLElement, last: HTMLElement): void {
    const selection = window.getSelection();
    if (!selection) return;
    const range = document.createRange();
    range.setStart(first, 0);
    range.setEnd(last, last.childNodes.length);
    selection.removeAllRanges();
    selection.addRange(range);
}

/**
 * 收集与选区相交的文本节点（文档顺序）
 * @param root - 根
 * @param range - 选区
 * @returns 文本节点数组
 */
function collectTextNodes(root: HTMLElement, range: Range): Text[] {
    const result: Text[] = [];
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null);
    let node = walker.nextNode();
    while (node) {
        const text = node as Text;
        if ((text.nodeValue ?? '').length > 0) {
            try {
                if (range.intersectsNode(text)) result.push(text);
            } catch {
                // 已脱离文档的节点会抛错：忽略即可
            }
        }
        node = walker.nextNode();
    }
    return result;
}

/**
 * 计算某个位置在块内的码点偏移
 * @param host - 块元素
 * @param container - 选区容器
 * @param offset - 容器内偏移
 * @returns 码点偏移
 */
function offsetInBlock(host: HTMLElement, container: Node, offset: number): number {
    let count = 0;
    let found = false;
    const walk = (node: Node): void => {
        if (found) return;
        if (node === container && node.nodeType === 1) {
            // 容器是元素：只数它前面 offset 个子节点里的字符
            for (let i = 0; i < offset && i < node.childNodes.length; i++) count += textLength(node.childNodes[i]);
            found = true;
            return;
        }
        if (node.nodeType === 3) {
            if (node === container) {
                count += offset;
                found = true;
                return;
            }
            count += (node.nodeValue ?? '').length;
            return;
        }
        if (node.nodeType === 1 && (node as HTMLElement).tagName === 'BR') {
            count += 1;
        }
        const children = node.childNodes;
        for (let i = 0; i < children.length; i++) walk(children[i]);
    };
    walk(host);
    return count;
}

/**
 * 取节点内文本长度（`<br>` 记 1）
 * @param node - 节点
 * @returns 长度
 */
function textLength(node: Node): number {
    if (node.nodeType === 3) return (node.nodeValue ?? '').length;
    if (node.nodeType === 1 && (node as HTMLElement).tagName === 'BR') return 1;
    let count = 0;
    const children = node.childNodes;
    for (let i = 0; i < children.length; i++) count += textLength(children[i]);
    return count;
}

/**
 * 按码点偏移定位 DOM 位置
 * @param host - 块元素
 * @param offset - 码点偏移
 * @returns 定位点；越界时落到末尾
 */
function locateInBlock(host: HTMLElement, offset: number): { node: Node; offset: number } | null {
    let remaining = offset;
    // 【为什么用盒子而不是 let 变量】赋值发生在闭包里，TypeScript 的控制流分析
    // 不会跟踪「回调内赋值」，会在 `if (!result)` 处把类型窄化成 null 而报错
    const box: { point: { node: Node; offset: number } | null } = { point: null };
    const walk = (node: Node): boolean => {
        if (node.nodeType === 3) {
            const length = (node.nodeValue ?? '').length;
            if (remaining <= length) {
                box.point = { node, offset: remaining };
                return true;
            }
            remaining -= length;
            return false;
        }
        if (node.nodeType === 1 && (node as HTMLElement).tagName === 'BR') {
            if (remaining === 0) {
                box.point = { node: node.parentNode ?? host, offset: indexOfNode(node) };
                return true;
            }
            remaining -= 1;
            return false;
        }
        const children = node.childNodes;
        for (let i = 0; i < children.length; i++) {
            if (walk(children[i])) return true;
        }
        return false;
    };
    walk(host);
    if (!box.point && host.firstChild) {
        const last = host.lastChild as Node;
        box.point = {
            node: last.nodeType === 3 ? last : host,
            offset: last.nodeType === 3 ? (last.nodeValue ?? '').length : host.childNodes.length,
        };
    }
    return box.point;
}

/**
 * 取节点在父元素中的下标
 * @param node - 节点
 * @returns 下标
 */
function indexOfNode(node: Node): number {
    const parent = node.parentNode;
    if (!parent) return 0;
    const children = parent.childNodes;
    for (let i = 0; i < children.length; i++) {
        if (children[i] === node) return i;
    }
    return 0;
}

/**
 * 按换行符把字符数组切成多行
 * @param chars - 字符数组
 * @returns 每行的字符数组
 */
function splitByNewline(chars: Char[]): Char[][] {
    const lines: Char[][] = [];
    let current: Char[] = [];
    for (let i = 0; i < chars.length; i++) {
        if (chars[i].char === '\n') {
            lines.push(current);
            current = [];
            continue;
        }
        current.push(chars[i]);
    }
    lines.push(current);
    return lines;
}

/**
 * CSS 像素 → 毫米（96dpi；打印样式 `@page size` 要的是物理单位）
 * @param px - 像素
 * @returns 毫米（保留两位小数，够打印引擎用）
 */
export function pxToMm(px: number): number {
    return Math.round((px * 25.4 / 96) * 100) / 100;
}

/**
 * `rgb()/rgba()` → `#RRGGBB`
 * @param color - CSS 颜色
 * @returns 十六进制颜色
 */
export function rgbToHex(color: string): string {
    const text = (color ?? '').trim();
    if (/^#[0-9a-f]{3,8}$/i.test(text)) return text.toUpperCase();
    const match = /^rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i.exec(text);
    if (!match) return text;
    const hex = (value: string): string => {
        const n = Math.max(0, Math.min(255, parseInt(value, 10)));
        return (n < 16 ? '0' : '') + n.toString(16).toUpperCase();
    };
    return `#${hex(match[1])}${hex(match[2])}${hex(match[3])}`;
}
