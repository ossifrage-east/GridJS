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
    DOC_PAGE_STACK_GAP,
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

    /**
     * 纸张背景层（一页一张白卡，卡间露出灰缝）
     *
     * 正文是**一条连续流**（分页只在流里插留白），所以纸本身是一整张长纸；
     * 屏幕上要像 Word 那样「一页一张纸」，就在纸的下面铺一排与正文同宽的纸卡，
     * 纸卡彼此错开 {@link DOC_PAGE_STACK_GAP} 像素，页与页之间于是露出灰底。
     */
    private readonly pagesLayer: HTMLDivElement;

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
     * 输入法组字中（`compositionstart` → `compositionend`）
     *
     * 组字期间**绝不能改正文 DOM**：分页要先摘掉旧留白，再对父节点 `normalize()`
     * 把拆开的文本节点合回去 —— 正在组字的那个文本节点会被并进去，
     * 浏览器随即结束组字，表现就是「拼音字母直接落进正文」。
     * 分页留白又是插在页边界上的，所以问题只在页末 / 页首附近出现。
     */
    private composing: boolean = false;

    /**
     * 组字文本所在的文本节点（`compositionstart` 时切出来，见 {@link DocView.isolateComposingText}）
     *
     * 把拼音和它前面的正文**分到两个文本节点**，页末组字溢到页脚时才能在
     * 「拼音节点之前」插块级留白，把组字整体顶到下一页正文起点 ——
     * 插元素不会搬动拼音节点，输入法不会被打断（搬节点 / 切节点才会）。
     */
    private composingNode: Text | null = null;

    /**
     * 组字期间被摘掉的**旧行级留白**（摘的时候先记下来，见 {@link DocView.removePageGaps}）
     *
     * 排页的开头总是先摘掉旧留白再量。组字期间若正巧碰到「必须拆开组字文本节点」
     * 的行首，这一轮就拆不动 —— 此时若让旧留白一直缺席，页边界就消失，
     * 后续内容整体上移、压到分页位置，整篇排版都会乱掉。所以这里先把旧留白
     * 记下来，拆不动时原样放回，保住页边界；组字结束后再正常拆行。
     */
    private composingGapBackup: {
        /** 留白所属的块元素 */
        block: HTMLElement;
        /** 留白原来的父节点 */
        parent: Node;
        /** 留白原来的下一个兄弟（放回时的锚点） */
        anchor: Node | null;
        /** 留白元素本身 */
        gap: HTMLElement;
        /** 留白高度（CSS 像素） */
        height: number;
    }[] = [];

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
        this.pagesLayer = createDiv({ className: 'doc-pages' });
        this.page = createDiv({ className: 'doc-page' });
        this.blocksEl = createDiv({ className: 'doc-blocks' });
        this.blocksEl.setAttribute('contenteditable', 'true');
        this.blocksEl.setAttribute('spellcheck', 'false');
        this.blocksEl.setAttribute('role', 'textbox');
        this.blocksEl.setAttribute('aria-multiline', 'true');

        this.page.appendChild(this.blocksEl);
        // 背景层在前、纸在后：纸的底色是透明的，白卡从纸的下面透出来
        this.slot.appendChild(this.pagesLayer);
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
            // 撤销 / 重做：store 已按「旧树 → 新树」的文本差算好「改动末尾」，
            // 这里在重渲染之后把光标送过去（不这样做会停在改动前的位置）
            const caret = this.store.historyCaret;
            if (!caret) return;
            this.store.clearHistoryCaret();
            // 焦点可能还留在工具栏的撤销按钮上：先让正文拿回焦点，光标才看得见
            this.blocksEl.focus({ preventScroll: true });
            this.restoreCaret(caret);
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
        /** 块容器（`contenteditable` 宿主；文档查找 / 替换在它上面做 DOM 扫描） */
        blocks: HTMLDivElement;
    } {
        return {
            scroller: this.scroller,
            page: this.page,
            slot: this.slot,
            rulerH: this.rulerH,
            rulerV: this.rulerV,
            blocks: this.blocksEl,
        };
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
        // 纸卡的步长取**单页**纸高（page 的高度是「N 页叠起来」的整叠高度）
        if (width > 0) this.syncPageCards(width, this.pageGeometry().height, z);
        buildRuler(this.rulerHInner, width, z, true);
        // 纵向量的是「整叠纸」：分页后纸会被撑到 N 页高，标尺要跟着长
        buildRuler(this.rulerVInner, height, z, false);
        // 缩放 / 换纸之后滚动位置可能已经越界，刻度条要按新的 scrollLeft/Top 归位
        this.syncRulers();
    }

    /**
     * 重排「一页一张纸」的背景纸卡
     *
     * 正文是一条连续流（分页只在流里插留白），因此纸本身是**一张整纸**；
     * 这里在纸下面铺一排白卡，每张卡比纸高少一截 {@link DOC_PAGE_STACK_GAP}，
     * 卡与卡之间就露出工作区的灰底（页与页看起来是分开的两张纸，与 Word 一致）。
     *
     * 纸卡不参与布局、不进模型，纯装饰：
     * - 位置 / 尺寸全部按**缩放后的像素**现算（纸是 `transform: scale`，纸卡不在纸里，吃不到这个缩放）；
     * - 页数不变时只改已有纸卡的尺寸，避免每帧重建 N 个节点。
     * @param width - 纸张宽（布局尺寸，未缩放）
     * @param paperHeight - 单页纸高（布局尺寸，未缩放）
     * @param z - 当前缩放比例
     * @private
     */
    private syncPageCards(width: number, paperHeight: number, z: number): void {
        const pages = Math.max(1, this.pageCountValue);
        const layer = this.pagesLayer;
        while (layer.childElementCount > pages) layer.removeChild(layer.lastElementChild as HTMLElement);
        while (layer.childElementCount < pages) layer.appendChild(createDiv({ className: 'doc-page-card' }));
        const cardWidth = `${width * z}px`;
        const cardHeight = `${Math.max(1, (paperHeight - DOC_PAGE_STACK_GAP) * z)}px`;
        const step = paperHeight * z;
        for (let i = 0; i < pages; i++) {
            const card = layer.children[i] as HTMLElement;
            card.style.width = cardWidth;
            card.style.height = cardHeight;
            card.style.top = `${i * step}px`;
        }
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
        // 重建 DOM 已经把组字打断（浏览器这时未必再发 compositionend）：
        // 标记归零，否则后面的排页会被永久挡住；组字节点引用也要跟着失效
        this.composing = false;
        this.composingNode = null;
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
        // 排页会 splitText / normalize，可能把光标所在的文本节点切走：
        // 先按（块 id + 偏移）记下光标，排完再放回去，否则光标会跳到上一页末尾。
        // 组字期间选区归输入法所有，这里不能碰（改法见 removePageGaps / 行级拆分）。
        const selection = window.getSelection();
        const collapsed = !!selection && selection.rangeCount > 0 && selection.isCollapsed;
        const caret = !this.composing && collapsed ? this.captureCaret() : null;
        const { height, margin, contentHeight } = this.pageGeometry();
        const z = this.zoomValue || 1;
        // ① 摘掉上一轮的分页留白（含块内「行级分页」留下的），并把被拆开的
        //    文本节点合回去 —— 否则每排一次页，块里的文本节点就多切一刀
        this.removePageGaps();
        const blocks = Array.prototype.slice.call(this.blocksEl.children) as HTMLElement[];
        if (blocks.length === 0) {
            this.pageCountValue = 1;
            this.applyFrame();
            if (caret) this.restoreCaret(caret);
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
            if (caret) this.restoreCaret(caret);
            return pages;
        }
        const tops: number[] = [];
        const bottoms: number[] = [];
        for (let i = 0; i < blocks.length; i++) {
            const rect = blocks[i].getBoundingClientRect();
            tops.push((rect.top - top0) / z);
            bottoms.push((rect.bottom - top0) / z);
        }

        // ② 逐块排页
        //
        // 版面坐标相对正文起点，第 n 页正文区是 `[n × 纸高, n × 纸高 + 正文区高]`；
        // `pageOf` 把「正文区内的纵坐标」映射成页号，正文区之外（上下页边距）不参与。
        //
        // 跨页块**优先按行拆分**：把「刚好放不下的那一行」之前的文本留在本页，
        // 之后的内容顶到下一页的正文起点。整段下推会让上一页底部空出一大块
        //（正文区只有 931px，一段 200px 的段落被推走就空 200px），这正是要修掉的现象。
        // 行高本身就是不可分割的排版单位，按行拆不会把一行劈成两半。
        const pageOf = (y: number): number => Math.floor((y + margin) / height);
        /** 第 page 页正文区的下边界（版面坐标） */
        const limitOf = (page: number): number => page * height + contentHeight;
        let page = 0;
        let shift = 0;

        for (let i = 0; i < blocks.length; i++) {
            const top = tops[i] + shift;
            let target = pageOf(top);
            // 落在页脚区（正文区之下）的块不能留在这一页：整块顶到下一页正文起点
            if (top > limitOf(target)) target += 1;
            if (target > page) {
                const gap = target * height - top;
                if (gap > 0) {
                    this.insertGap(blocks[i], gap, target);
                    shift += gap;
                }
                page = target;
            }
            const bottomNow = bottoms[i] + shift;
            if (bottomNow <= limitOf(page) + 1) continue;

            // 跨页：先尝试行级拆分
            const split = this.splitBlockAcrossPages(blocks[i], top0, z, page, height, contentHeight, shift);
            if (split) {
                shift += split.added;
                page = split.lastPage;
                continue;
            }
            // 拆不动（表格 / 空块 / 单行比正文区还高）：退回整块下推
            const endPage = Math.max(page + 1, pageOf(bottomNow));
            const gap = endPage * height - (tops[i] + shift);
            if (gap > 0) {
                this.insertGap(blocks[i], gap, endPage);
                shift += gap;
            }
            page = endPage;
        }

        const total = bottoms.length > 0 ? bottoms[bottoms.length - 1] + shift : 0;
        const pages = Math.max(1, pageOf(Math.max(total - 1, 0)) + 1);
        this.pageCountValue = pages;
        this.page.style.minHeight = `${Math.max(height, height * pages)}px`;
        this.applyFrame();
        if (caret) this.restoreCaret(caret);
        return pages;
    }

    /**
     * 摘掉上一轮插进正文的分页留白，并合并被拆开的文本节点
     *
     * 留白分两类：块与块之间的（直接子元素）与**块内跨页**的（夹在行内 span 之间）。
     * 两类都是 `contenteditable=false` 的空元素，不承载文本；删掉它们之后
     * `normalize()` 会把 `splitText` 切开的相邻文本节点合回去，模型读取不受影响。
     */
    private removePageGaps(): void {
        const gaps = this.blocksEl.querySelectorAll(`[${DOC_PAGE_GAP_ATTR}]`);
        // 每轮重新记：上一轮记下但没用上的（已经正常重排）不需要留到这一轮
        this.composingGapBackup = [];
        for (let i = gaps.length - 1; i >= 0; i--) {
            const gap = gaps[i] as HTMLElement;
            const parent = gap.parentNode;
            if (!parent) continue;
            // 组字期间先把**行级**留白的位置记下来：拆不动组字节点时要用它们把页边界放回去
            if (this.composing && gap.getAttribute(DOC_PAGE_GAP_ATTR) === 'line') {
                // 留白的父节点是 run 段（span）时，块是它的父元素；直接挂在块下时块就是父节点
                const parentEl = parent.nodeType === 1 ? (parent as HTMLElement) : null;
                const block = parentEl ? (parentEl.hasAttribute('data-id') ? parentEl : parentEl.parentElement) : null;
                if (block) {
                    this.composingGapBackup.push({
                        block,
                        parent,
                        anchor: gap.nextSibling,
                        gap,
                        height: parseFloat(gap.style.height) || 0,
                    });
                }
            }
            parent.removeChild(gap);
            // 组字期间不能 normalize：它会把正在组字的文本节点并回去，
            // 浏览器随即结束组字（拼音字母直接落进正文）。组字结束后再合并。
            if (!this.composing && parent.nodeType === 1) (parent as HTMLElement).normalize();
        }
    }

    /**
     * 把组字那个块上一轮的页边界放回去
     *
     * 只放回「这一块」的行级留白：其余块的留白会由本轮排页重新插好，
     * 这时若把全部旧留白都塞回去，就会插两遍、把内容顶到更下面去。
     * @param el - 当前正在拆的块
     * @returns 放回的总高度与被放回的留白条数；这一块本来就没有行级留白时为 0 条
     */
    private restoreComposingGaps(el: HTMLElement): { added: number; count: number } {
        let added = 0;
        let count = 0;
        for (let i = this.composingGapBackup.length - 1; i >= 0; i--) {
            const item = this.composingGapBackup[i];
            if (item.block !== el) continue;
            this.composingGapBackup.splice(i, 1);
            // 父节点已经被本次排页改掉 / 锚点已不在它下面：放不回去了，放弃
            if (!item.parent.isConnected || (item.anchor && item.anchor.parentNode !== item.parent)) continue;
            item.parent.insertBefore(item.gap, item.anchor);
            added += item.height;
            count += 1;
        }
        return { added, count };
    }

    /**
     * 在块前插入一个分页留白
     * @param before - 目标块
     * @param gap - 留白高度（CSS 像素，未缩放）
     * @param page - 目标页号（仅用于调试标记）
     */
    private insertGap(before: HTMLElement, gap: number, page: number): void {
        const spacer = createDiv({ className: 'doc-page-gap' });
        spacer.setAttribute(DOC_PAGE_GAP_ATTR, String(page));
        spacer.setAttribute('contenteditable', 'false');
        spacer.style.height = `${Math.max(0, Math.round(gap))}px`;
        this.blocksEl.insertBefore(spacer, before);
    }

    /**
     * 把一个跨页的块按行拆分（在「放不下的那一行」之前插入块级留白）
     *
     * 先逐字符量出每一行的行首，再把留白插在行首之前 —— 行内不能拆，
     * 否则会在页面底部留下半行字。插留白只改 DOM 的视觉布局，
     * 块 id / 块树 / 同步逻辑都不受影响（留白里没有文本）。
     * @param el - 块元素
     * @param top0 - 正文起点（视口坐标）
     * @param z - 缩放比
     * @param startPage - 块开始所在页
     * @param height - 纸高（未缩放）
     * @param contentHeight - 正文区高
     * @param shift - 前面已插留白造成的累计位移
     * @returns 新增留白总高与最后所在页；无法拆分时 null
     */
    private splitBlockAcrossPages(
        el: HTMLElement,
        top0: number,
        z: number,
        startPage: number,
        height: number,
        contentHeight: number,
        shift: number
    ): { added: number; lastPage: number } | null {
        // 表格 / 分割线 / 图片没有「行」的语义，拆开就不再是一个整体
        const type = el.getAttribute('data-type') ?? '';
        if (type === 'table' || type === 'hr' || type === 'image') return null;

        const rects = this.measureCharRects(el, top0, z);
        if (!rects || rects.length < 2) return null;

        let page = startPage;
        let extra = 0;
        let lineTop = rects[0].top;
        const splits: { node: Text; offset: number; gap: number; compose?: boolean }[] = [];

        for (let i = 0; i < rects.length; i++) {
            const item = rects[i];
            // 行首判定：纵坐标比当前行顶低一截才是新行。阈值取 8px（版面坐标）：
            // 上下标只偏移 2–4px，不会被误判成换行；正常行距是 20px 以上。
            const isLineStart = i === 0 || item.top > lineTop + 8;
            if (isLineStart) lineTop = item.top;
            else continue;

            const actualBottom = item.bottom + shift + extra;
            if (actualBottom <= page * height + contentHeight + 1) continue;
            // 第一行都放不下：不是「拆行」能解决的，交给调用方整块下推
            // （整块下推只插块级留白、不碰任何文本节点，组字期间也是安全的）
            if (i === 0) return null;
            // 组字期间遇到「会动到组字文本节点」的行首：不能走常规拆行
            // （`insertLineGap` 会把组字节点搬进新 span，输入法会立刻结束）。
            if (this.composing && this.hitsComposingNode(item.node, item.offset)) {
                // 组字所在行溢到页脚：在行首插**块级**留白，断开行而**不搬动
                // 组字文本节点**，拼音就会整体落到下一页正文起点，不再占页脚。
                // 这比「把组字节点搬进新 span」温和：搬节点会立刻结束输入法。
                const gapHere = (page + 1) * height - (item.top + shift + extra);
                // 与普通行级留白一起**延后倒序插入**：先插靠后的行首，
                // 前面节点的下标才不会被这一次 splitText 改掉
                if (gapHere > 0 && this.composingNode !== null) {
                    splits.push({ node: item.node, offset: item.offset, gap: gapHere, compose: true });
                    extra += gapHere;
                    page += 1;
                    continue;
                }
                // 连行首都切不动（光标不在文本节点里）：退回旧做法，保住页边界
                const restored = this.restoreComposingGaps(el);
                // 放回了几条留白，这一块就还是停在原来的页数上；`added` 要计入后续块的位移
                return { added: restored.added, lastPage: Math.min(page, startPage + restored.count) };
            }
            // 单行比整页正文区还高，继续拆只会无限循环：放弃
            if (page > startPage + 200) return null;

            const gap = (page + 1) * height - (item.top + shift + extra);
            if (gap > 0) splits.push({ node: item.node, offset: item.offset, gap });
            extra += Math.max(0, gap);
            page += 1;
        }

        if (splits.length === 0) return null;
        // 倒序插入：先处理靠后的行首，前面文本节点的下标才不受影响
        for (let s = splits.length - 1; s >= 0; s--) {
            const split = splits[s];
            if (!split.node.parentNode) continue;
            if (split.compose) {
                // 组字行首：只插块级留白，绝不搬动组字文本节点
                this.insertComposingPageBreak(split.node, split.offset, split.gap);
                continue;
            }
            this.insertLineGap(el, split.node, split.offset, split.gap);
        }
        return { added: extra, lastPage: page };
    }

    /**
     * 逐字符量出版面坐标（只用于行级分页）
     *
     * 用 Range 逐字符取矩形：比「按行取矩形」多花一点时间，但能把行首精确定位到
     * 文本节点 + 偏移，之后才能安全地 `splitText` / 拆 span。
     * 超长块直接放弃（返回 null）——分页是显示层，不值得为一段几千字的文本卡住输入。
     * @param el - 块元素
     * @param top0 - 正文起点（视口坐标）
     * @param z - 缩放比
     * @returns 字符矩形；无文本或超出上限时 null
     */
    private measureCharRects(
        el: HTMLElement,
        top0: number,
        z: number
    ): { node: Text; offset: number; top: number; bottom: number }[] | null {
        const MAX_CHARS = 2400;
        const result: { node: Text; offset: number; top: number; bottom: number }[] = [];
        const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, null);
        const range = document.createRange();
        let node = walker.nextNode() as Text | null;
        while (node) {
            const text = node.nodeValue ?? '';
            for (let i = 0; i < text.length; i++) {
                if (result.length > MAX_CHARS) return null;
                range.setStart(node, i);
                range.setEnd(node, i + 1);
                const rect = range.getBoundingClientRect();
                if (rect.height <= 0) continue;
                result.push({
                    node,
                    offset: i,
                    top: (rect.top - top0) / z,
                    bottom: (rect.bottom - top0) / z,
                });
            }
            node = walker.nextNode() as Text | null;
        }
        return result.length > 0 ? result : null;
    }

    /**
     * 这个行首的拆分会不会动到正在组字的文本节点
     *
     * 会动到的两种情况：
     * 1. 行首就在组字节点里 —— `splitText` 直接把节点切开；
     * 2. 组字节点与行首在同一个父节点、且排在它后面 —— `insertLineGap` 会把
     *    行首之后的兄弟节点搬进新 span，组字节点跟着被搬走。
     * 两种都会让浏览器立刻结束组字（拼音字母落进正文），所以只能等组字结束再拆。
     * @param node - 行首所在文本节点
     * @param offset - 行首在节点内的偏移
     * @returns 是否命中组字节点
     */
    private hitsComposingNode(node: Text, offset: number): boolean {
        const selection = window.getSelection();
        if (!selection || selection.rangeCount === 0) return false;
        const caret = selection.getRangeAt(0).startContainer;
        if (caret.nodeType !== 3) return false;
        if (caret === node) {
            // 同一个文本节点：只有「拆分点落在组字区里」才会打断输入法。
            // 组字那段拼音永远在光标左侧，所以拆分点在光标**之后**时，
            // 拼音仍留在原节点里，切出来的只是它右边还没组字的文字 —— 安全。
            // 这一条很关键：普通段落整段只有一个文本节点，若一律拦下，
            // 只要段落跨页、在页内任意位置打字都会命中守卫，整篇内容就会上移。
            return offset < selection.getRangeAt(0).startOffset;
        }
        if (caret.parentNode !== node.parentNode) return false;
        const probe = document.createRange();
        probe.setStart(node, offset);
        probe.collapse(true);
        return probe.comparePoint(caret, 0) >= 1;
    }

    /**
     * 把「即将组字的拼音」切到一个独立文本节点
     *
     * `compositionstart` 时拼音还没写进正文，此刻在光标处 `splitText` 是安全的；
     * 切完之后拼音从新节点的下标 0 开始，页末换行时行首正好落在节点开头，
     * 分页才能在它前面插留白（见 {@link DocView.insertComposingPageBreak}）——
     * 组字期间再切节点会立刻结束输入法。
     * @returns 承接组字的文本节点；光标不在文本节点里时 null（退回旧的保守处理）
     */
    private isolateComposingText(): Text | null {
        const selection = window.getSelection();
        if (!selection || selection.rangeCount === 0) return null;
        const range = selection.getRangeAt(0);
        if (!range.collapsed) return null;
        const node = range.startContainer;
        if (node.nodeType !== 3) return null;
        const text = node as Text;
        const offset = range.startOffset;
        if (offset === 0) return text;
        // 光标在节点中间 / 末尾：切出一个从光标开始的节点（末尾是空节点），
        // 拼音落进去之后与前面的正文就不再共享文本节点
        const host = text.splitText(offset);
        const next = document.createRange();
        next.setStart(host, 0);
        next.collapse(true);
        selection.removeAllRanges();
        selection.addRange(next);
        return host;
    }

    /**
     * 组字行溢出到页脚时，在行首插一个**块级**页边界
     *
     * 关键：只往 DOM 里**插**一个 `contenteditable=false` 的留白元素，既不切
     * 组字文本节点、也不搬动它。块级元素插在行内文本之间会把行断开，断点之后
     * 的拼音与后续正文整体落到留白之后（下一页正文起点），输入法不受影响。
     * @param node - 行首所在文本节点
     * @param offset - 行首在节点内的偏移
     * @param gap - 留白高度（把这一行顶到下一页正文起点所需）
     * @returns 是否插入成功
     */
    private insertComposingPageBreak(node: Text, offset: number, gap: number): boolean {
        if (gap <= 0) return false;
        const composing = this.composingNode;
        // 没能把拼音单独切出来时不敢动：切组字节点会直接结束输入法
        if (!composing) return false;
        const parent = node.parentNode;
        if (!parent) return false;
        // 行首在组字节点之前：按行首切开（切到的不是组字节点，安全）
        const pivot = node !== composing && offset > 0 ? node.splitText(offset) : node;
        const spacer = createDiv({ className: 'doc-page-gap' });
        spacer.setAttribute(DOC_PAGE_GAP_ATTR, 'compose');
        spacer.setAttribute('contenteditable', 'false');
        // 同一节点可能跨多条页边界：本轮已经插过，就把高度累加
        const previous = pivot.previousSibling as HTMLElement | null;
        if (previous && previous.getAttribute && previous.getAttribute(DOC_PAGE_GAP_ATTR) === 'compose') {
            const height = parseFloat(previous.style.height) || 0;
            previous.style.height = `${Math.round((height + gap) * 100) / 100}px`;
            return true;
        }
        spacer.style.height = `${Math.max(1, Math.round(gap))}px`;
        const container = pivot.parentNode;
        if (!container) return false;
        container.insertBefore(spacer, pivot);
        return true;
    }

    /**
     * 在某个文本位置插入「行级分页留白」
     *
     * 行内 span 是块的直接子元素，所以拆法是：把「该位置之后」的节点搬进一个
     * 同样式的空 span，再在两者之间插留白。这样块的直接子元素顺序仍是
     * 「文字、留白、文字」，`readInline` 遍历时留白没有文本，模型不受影响。
     * @param el - 块元素
     * @param node - 行首所在文本节点
     * @param offset - 行首在文本节点内的偏移
     * @param gap - 留白高度
     */
    private insertLineGap(el: HTMLElement, node: Text, offset: number, gap: number): void {
        const spacer = createDiv({ className: 'doc-page-gap' });
        spacer.setAttribute(DOC_PAGE_GAP_ATTR, 'line');
        spacer.setAttribute('contenteditable', 'false');
        spacer.style.height = `${Math.max(1, Math.round(gap))}px`;

        // 行首在文本节点中间：先切开文本
        const pivot = offset > 0 ? node.splitText(offset) : node;
        const parent = pivot.parentNode;
        if (!parent) return;

        // 文本直接挂在块下（空段落里刚输入的字）：留白插在文本之前
        if (parent === el) {
            el.insertBefore(spacer, pivot);
            return;
        }
        // 常规形态：文本在 run span 里，span 是块的直接子元素
        if (parent.nodeType === 1 && parent.parentNode === el) {
            const holder = parent as HTMLElement;
            const clone = holder.cloneNode(false) as HTMLElement;
            let cursor: Node | null = pivot;
            while (cursor) {
                const next = cursor.nextSibling;
                clone.appendChild(cursor);
                cursor = next;
            }
            el.insertBefore(spacer, holder.nextSibling);
            el.insertBefore(clone, spacer.nextSibling);
            return;
        }
        // 更深的嵌套（当前渲染不会出现）：保留切好的文本，不插留白，
        // 下一次 `normalize()` 会把相邻文本节点合回去，内容不丢
    }

    /**
     * 光标右侧紧邻的分页留白
     *
     * 只在「右侧第一个可编辑内容之前正好是留白」时返回它；其余情况返回 null，
     * 让浏览器的原生删除行为照常工作（同段落内的删字、段尾合并下一段都靠它）。
     * @param range - 折叠光标
     * @returns 留白元素；右侧没有留白时为 null
     */
    private gapAfterCaret(range: Range): HTMLElement | null {
        const node = range.startContainer;
        if (node.nodeType === 3 /* Node.TEXT_NODE */) {
            // 光标不在文本末尾：右侧字符在同一个节点里，原生删除可用
            if (range.startOffset < (node.nodeValue ?? '').length) return null;
        } else if (node.nodeType === 1 /* Node.ELEMENT_NODE */) {
            let sibling = (node as HTMLElement).childNodes[range.startOffset] ?? null;
            while (sibling) {
                if (isPageGap(sibling)) return sibling as HTMLElement;
                if (hasEditableText(sibling)) return null;
                sibling = sibling.nextSibling;
            }
        }
        // 沿祖先向上找「右侧兄弟」：留白夹在当前节点的外层容器之后时也能命中
        let cursor: Node | null = node;
        while (cursor && cursor !== this.blocksEl) {
            let sibling = cursor.nextSibling;
            while (sibling) {
                if (isPageGap(sibling)) return sibling as HTMLElement;
                if (hasEditableText(sibling)) return null;
                sibling = sibling.nextSibling;
            }
            cursor = cursor.parentNode;
        }
        return null;
    }

    /**
     * 光标左侧紧邻的分页留白（{@link gapAfterCaret} 的镜像）
     * @param range - 折叠光标
     * @returns 留白元素；左侧没有留白时为 null
     */
    private gapBeforeCaret(range: Range): HTMLElement | null {
        const node = range.startContainer;
        if (node.nodeType === 3 /* Node.TEXT_NODE */) {
            if (range.startOffset > 0) return null;
        } else if (node.nodeType === 1 /* Node.ELEMENT_NODE */) {
            let sibling = (node as HTMLElement).childNodes[range.startOffset - 1] ?? null;
            while (sibling) {
                if (isPageGap(sibling)) return sibling as HTMLElement;
                if (hasEditableText(sibling)) return null;
                sibling = sibling.previousSibling;
            }
        }
        let cursor: Node | null = node;
        while (cursor && cursor !== this.blocksEl) {
            let sibling = cursor.previousSibling;
            while (sibling) {
                if (isPageGap(sibling)) return sibling as HTMLElement;
                if (hasEditableText(sibling)) return null;
                sibling = sibling.previousSibling;
            }
            cursor = cursor.parentNode;
        }
        return null;
    }

    /**
     * 越过一个分页留白删掉相邻页的一个字符
     *
     * 删除后光标**留在原处**：向前删时原处就是「上一页页末」，向后删时就是
     * 「下一页页首」，正是删完一个字之后该在的位置。程序化改 DOM 不会触发
     * `input`，因此这里手动排一次同步（模型、撤销栈、分页都靠它收尾）。
     * @param gap - 分页留白
     * @param forward - true 删除留白之后的第一个字符（Delete），false 删除之前的最后一个字符（Backspace）
     */
    private deleteAcrossPageGap(gap: HTMLElement, forward: boolean): void {
        const text = forward
            ? firstTextAfter(gap, this.blocksEl)
            : lastTextBefore(gap, this.blocksEl);
        if (!text) return;
        const value = text.nodeValue ?? '';
        if (!value) return;
        if (forward) {
            text.deleteData(0, firstCodePointLength(value));
        } else {
            const size = lastCodePointLength(value);
            text.deleteData(value.length - size, size);
        }
        this.scheduleSync();
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
        if (block.firstLineIndent) {
            el.setAttribute('data-first-line', String(block.firstLineIndent));
            // 首行缩进存的是**字符数**，渲染时按块内首字符字号折成像素：
            // 1 字符 = 该字号下的 1em（16px 正文即 16px，16pt 正文即 21.3px）。
            // 与 `data-indent` 同理，样式必须在这里还原，否则重渲染后首行缩进会丢。
            el.style.textIndent = `${block.firstLineIndent * firstCharPx(block)}px`;
        }
        // 空格宽度：块级默认值兜底「直接落在块元素下的文本」（空段落里刚打的字），
        // 每个 run 段另有自己的 word-spacing（见 renderInline），会把这里的值盖掉
        this.applySpaceWidth(el, block);
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
            // 只加「一个字符宽 − 空格自然宽」的差额（word-spacing）：改的是空格自己的步进宽度，
            // 换行机会、光标、输入法都不受影响。把空格换成撑宽的盒子会吃掉断行点，不能这么做。
            if (text.indexOf(' ') >= 0) holder.style.wordSpacing = spaceWordSpacingOfText(chars.slice(index, end), mono);
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
     * 给块元素写「块级空格宽度」
     *
     * 渲染出的每个 run 段都有自己的 `word-spacing`，块级值只兜底一种情况：
     * 空段落里刚打的字直接挂在块元素下（没有 run 段可继承），此时靠块级值让空格立刻是「一格」宽。
     * @param el - 块元素
     * @param block - 块
     */
    private applySpaceWidth(el: HTMLElement, block: DocBlock): void {
        if (block.type === 'table') return;
        const value = spaceWordSpacingOfText(block.chars ?? [], block.type === 'code');
        // 只在值真的变了才写：同步是一帧一次、且要遍历全部块，
        // 每块都写一次内联样式会让「每敲一个字」都触发一轮全篇样式重算
        if (el.style.wordSpacing !== value) el.style.wordSpacing = value;
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
        this.blocksEl.addEventListener('input', (event) => {
            if (this.rendering) return;
            // 组字中途的 input 是临时的拼音串：不写进模型，但**仍要排页** ——
            // 否则拼音占位把后面的文字挤到页脚空白里，要等组字结束才跳回下一页。
            if (this.composing || (event as InputEvent).isComposing) {
                this.schedulePaginate();
                return;
            }
            this.scheduleSync();
        });

        // 输入法组字：组字期间同步不做、排页照做（排页内部会绕开组字文本节点）。
        this.blocksEl.addEventListener('compositionstart', () => {
            this.composing = true;
            // 组字还没往正文写拼音：此刻把拼音单独切到一个文本节点最安全，
            // 之后页末组字溢到页脚时才有「干净的分界点」插页边界
            this.composingNode = this.isolateComposingText();
        });
        this.blocksEl.addEventListener('compositionend', () => {
            this.composing = false;
            this.composingNode = null;
            // 此刻正文里已经是最终汉字：同步回模型，再由 sync 末尾安排一次排页
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

        // 【跨页删除】页末 / 页首的删除键要能越过分页留白
        //
        // 分页留白是 `contenteditable=false` 的块级元素，浏览器把它当成一个
        // 「不可编辑的整体」：光标停在它前面按 Delete、或停在它后面按 Backspace，
        // 默认行为删不掉相邻页上的第一个 / 最后一个字符（有时还会把留白本身删掉）。
        // 这里只在「紧邻位置就是留白」时接管，其余情况一律放行给浏览器 —— 
        // 段落合并、词删除（Ctrl+Delete）、输入法组字都仍由原生行为处理。
        this.blocksEl.addEventListener('keydown', (event) => {
            if (event.key !== 'Delete' && event.key !== 'Backspace') return;
            if (event.ctrlKey || event.metaKey || event.altKey || event.isComposing) return;
            const selection = window.getSelection();
            if (!selection || selection.rangeCount === 0 || !selection.isCollapsed) return;
            const range = selection.getRangeAt(0);
            if (!this.blocksEl.contains(range.startContainer)) return;

            const forward = event.key === 'Delete';
            const gap = forward ? this.gapAfterCaret(range) : this.gapBeforeCaret(range);
            if (!gap) return;
            event.preventDefault();
            this.deleteAcrossPageGap(gap, forward);
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
            if (block) {
                blocks.push(block);
                // 文字变了，空格的「一格宽」也要跟着这一块的字号 / 语种重算
                // （只改 word-spacing 这种纯呈现属性，不动节点，光标与输入法不受影响）
                this.applySpaceWidth(el, block);
            }
        }
        if (blocks.length === 0) return;
        this.store.update(blocks, label, mode);
        if (reset) this.render();
        // 内容高度可能变了（打字换行、插入图片）：重排一次页。
        // **必须在本帧内同步排完**：若再排到下一帧，浏览器会先画一帧
        // 「新字符还留在页脚空白里」的中间状态，页末逐字输入时就表现为闪一下。
        // （重渲染那条路自己已经重新安排过排页，这里不必重复。）
        if (!reset) this.paginate();
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
        // 首行缩进（字符数）：属性是唯一来源，样式里的 px 只是它的投影
        const firstLine = parseFloat(el.getAttribute('data-first-line') ?? '');
        if (Number.isFinite(firstLine) && firstLine > 0) block.firstLineIndent = firstLine;

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
            // 【多行文本按行拆成多个块】不能整段塞进一个块：分页是「整块不跨页」，
            // 一个装着 N 行的块会被整块推到下一页，当前页剩下的空白就白搭了。
            // 拆成「一行一个块」后，分页可以在行间断开 —— 与 Enter 换行的模型一致。
            const lines = text.replace(/\r\n?/g, '\n').split('\n');
            const caret = lines.length > 1 ? this.splitIntoBlocks(range, lines) : null;
            if (caret) {
                this.sync('粘贴', 'single', false);
                this.restoreCaret(caret);
                return;
            }
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
     * 把多行文本按行拆成多个块，插在当前选区处（粘贴多行时用）
     *
     * 拆块的语义与 Word 一致：光标**前**的文字留在原块并与第一行拼接，
     * 光标**后**的文字搬到最后一个新块（跟在最后一行之后），光标停在粘贴内容末尾。
     *
     * 表格 / 代码块 / 图片 / 分割线是「整块语义」的宿主（拆开就不再是一个表格或一段代码），
     * 不在这里拆，返回 null 交给调用方按原样整段插入。
     * @param range - 当前选区（可能非折叠，先删掉选区内容再插入）
     * @param lines - 已按行拆好的文本
     * @returns 粘贴后光标应落的位置；未拆块时为 null
     * @private
     */
    private splitIntoBlocks(range: Range, lines: string[]): DocCaret | null {
        const host = this.closestBlock(range.startContainer);
        if (!host) return null;
        const type = host.getAttribute('data-type') ?? 'paragraph';
        if (type === 'table' || type === 'code' || type === 'image' || type === 'hr') return null;

        range.deleteContents();
        // 光标之后的节点整段摘出来，最后接到末块末尾（前缀…粘贴内容…后缀）
        const tail = document.createRange();
        tail.selectNodeContents(host);
        tail.setStart(range.startContainer, range.startOffset);
        const after = tail.extractContents();

        if (lines[0]) range.insertNode(document.createTextNode(lines[0]));

        // 一行一个块：克隆宿主的块级属性（类型 / 级别 / 列表 / 缩进 / 对齐）但换新 id
        let anchor: HTMLElement = host;
        let caretBlockId = '';
        for (let i = 1; i < lines.length; i++) {
            const el = host.cloneNode(false) as HTMLElement;
            el.setAttribute('data-id', nextBlockId());
            const last = i === lines.length - 1;
            if (lines[i]) el.appendChild(document.createTextNode(lines[i]));
            // 空行给个占位 <br>（读模型时会认成空块）；末行有后缀时不加，
            // 否则后缀会被顶到空行的下一行去
            else if (!last || after.childNodes.length === 0) el.appendChild(document.createElement('br'));
            anchor.parentNode?.insertBefore(el, anchor.nextSibling);
            anchor = el;
            caretBlockId = el.getAttribute('data-id') ?? '';
        }
        if (after.childNodes.length > 0) anchor.appendChild(after);
        if (!caretBlockId) return null;

        // 光标落在粘贴内容末尾（后缀之前）
        return { blockId: caretBlockId, offset: Array.from(lines[lines.length - 1]).length };
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

/* ———————————————————— 分页留白两侧的删除 ——————————————————— */

/**
 * 是否分页留白元素
 * @param node - 节点
 * @returns 是否留白
 */
function isPageGap(node: Node): boolean {
    return node.nodeType === 1 && (node as Element).hasAttribute(DOC_PAGE_GAP_ATTR);
}

/**
 * 子树里是否还有可编辑的文字
 *
 * 空 span / 空段落不算：光标可以直接越过它们继续找下一段真正的内容。
 * @param node - 节点
 * @returns 是否含文字
 */
function hasEditableText(node: Node): boolean {
    if (node.nodeType === 3 /* Node.TEXT_NODE */) return (node.nodeValue ?? '').length > 0;
    if (node.nodeType !== 1 /* Node.ELEMENT_NODE */) return false;
    if (isPageGap(node)) return false;
    const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT, null);
    return walker.nextNode() !== null;
}

/**
 * 取留白之后的第一段文字（按文档顺序，不跨出正文）
 * @param gap - 分页留白
 * @param host - 正文容器
 * @returns 文本节点；没有时 null
 */
function firstTextAfter(gap: Node, host: HTMLElement): Text | null {
    const walker = document.createTreeWalker(host, NodeFilter.SHOW_TEXT, null);
    let node = walker.nextNode() as Text | null;
    while (node) {
        // 参数在前、node 在后 → 位掩码含 FOLLOWING（4）
        if (gap.compareDocumentPosition(node) & 4 && (node.nodeValue ?? '').length > 0) return node;
        node = walker.nextNode() as Text | null;
    }
    return null;
}

/**
 * 取留白之前的最后一段文字（{@link firstTextAfter} 的镜像）
 * @param gap - 分页留白
 * @param host - 正文容器
 * @returns 文本节点；没有时 null
 */
function lastTextBefore(gap: Node, host: HTMLElement): Text | null {
    const walker = document.createTreeWalker(host, NodeFilter.SHOW_TEXT, null);
    let found: Text | null = null;
    let node = walker.nextNode() as Text | null;
    while (node) {
        // 参数在前、node 在后 → 位掩码含 PRECEDING（2）
        if ((gap.compareDocumentPosition(node) & 2) && (node.nodeValue ?? '').length > 0) found = node;
        node = walker.nextNode() as Text | null;
    }
    return found;
}

/**
 * 首个码点占几个 UTF-16 单元（代理对算一个字符）
 * @param text - 文本
 * @returns 1 或 2
 */
function firstCodePointLength(text: string): number {
    const first = text.charCodeAt(0);
    return first >= 0xd800 && first <= 0xdbff && text.length > 1 ? 2 : 1;
}

/**
 * 末个码点占几个 UTF-16 单元（代理对算一个字符）
 * @param text - 文本
 * @returns 1 或 2
 */
function lastCodePointLength(text: string): number {
    if (text.length < 2) return 1;
    const last = text.charCodeAt(text.length - 1);
    return last >= 0xdc00 && last <= 0xdfff ? 2 : 1;
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
 * 块内「一个字符」的像素宽（首行缩进按它折算）
 *
 * 首行缩进在模型里存的是字符数，落到屏幕必须乘上该段落的字号：
 * 12pt（16px）正文的一格是 16px，16pt（21.3px）正文的一格是 21.3px。
 * @param block - 块
 * @returns 像素宽
 */
function firstCharPx(block: DocBlock): number {
    const chars = block.chars ?? [];
    for (let i = 0; i < chars.length; i++) {
        const size = chars[i]?.fontSize;
        if (size) return size;
        // 首个实义字符没有显式字号 → 用正文默认字号；开头的空格不参与判断
        if (String(chars[i]?.char ?? '') !== ' ') break;
    }
    return DEFAULT_DOC_FONT_SIZE;
}

/* ———————————————————— 空格宽度（半角） ———————————————————— */

/** 空格自然宽度测量缓存（key = 字体描述|字号），避免每个文字段都 new 一个 canvas */
const spaceAdvanceCache = new Map<string, number>();

/** 复用的 2D 上下文；取不到（受限环境）时按 1/4 字符宽估算 */
let spaceMeasureCtx: CanvasRenderingContext2D | null = null;

/**
 * 量一个空格在某字体 / 字号下的自然宽度（px）
 * @param font - CSS font-family 列表
 * @param fontSize - 字号（px）
 * @returns 自然宽度
 */
function spaceAdvance(font: string, fontSize: number): number {
    const key = `${font}|${fontSize}`;
    const cached = spaceAdvanceCache.get(key);
    if (cached !== undefined) return cached;
    let width = fontSize * 0.25;
    try {
        if (!spaceMeasureCtx) spaceMeasureCtx = document.createElement('canvas').getContext('2d');
        if (spaceMeasureCtx) {
            spaceMeasureCtx.font = `${fontSize}px ${font}`;
            const measured = spaceMeasureCtx.measureText(' ').width;
            if (Number.isFinite(measured) && measured > 0) width = measured;
        }
    } catch {
        // 量不到就沿用「1/4 字符宽」的估算值
    }
    spaceAdvanceCache.set(key, width);
    return width;
}

/**
 * 一段文字里「空格的附加宽度」（即 `word-spacing` 的值）
 *
 * 目标宽度是 **0.5em**：Word 里的 ASCII 空格是半角，中文段里「4 个空格」正好等于
 * 「2 个汉字宽」（官方文档常用 4 个空格做两字缩进）。按 1em 撑宽会让这 4 个空格
 * 变成 4 字宽，导入 Word 文档时首段就比原文多缩一倍。
 * 写差额而不是绝对宽度：空格仍是普通空格字符，断行机会、光标、输入法全交给浏览器，
 * 不会像「撑宽的盒子」那样吃掉断行点。
 * @param chars - 段内字符
 * @param mono - 是否等宽（代码块）
 * @returns CSS 长度
 */
function spaceWordSpacingOfText(chars: Char[], mono: boolean): string {
    let head: Char | undefined = chars[0];
    for (let i = 0; i < chars.length; i++) {
        const piece = String(chars[i]?.char ?? '');
        if (piece === ' ' || piece === '\t' || piece === '\n' || piece === '') continue;
        head = chars[i];
        break;
    }
    const size = head && head.fontSize ? head.fontSize : DEFAULT_DOC_FONT_SIZE;
    const font = mono
        ? 'Consolas, "Courier New", monospace'
        : head && head.fontFamily
            ? `"${head.fontFamily}", ${DEFAULT_DOC_FONT_FAMILY}`
            : DEFAULT_DOC_FONT_FAMILY;
    const target = size / 2;
    const extra = Math.max(0, target - spaceAdvance(font, size));
    return `${Math.round(extra * 100) / 100}px`;
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
