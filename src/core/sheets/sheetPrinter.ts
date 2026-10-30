/*  # 打印工具模块 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import type { Cell } from "../dataArchitecture/cell";
import type { DataCollection } from "../dataArchitecture/dataCollection";
import { Slider, SliderChangeSource } from "../common/slider";
import { getRegisteredFontFaceCss } from "../../utils/fontsLoader";
// 预览外壳（遮罩 / 标题 / 左右分栏 / 页脚翻页与缩放）与文档预览共用同一份
import { buildPrintPreviewShellHtml } from "../../utils/printPreviewShell";
// 默认字体与主页面 canvas 保持一致（单一数据源），保证预览/打印与主页面的字体观感相同
import { DEFAULT_FONT_FAMILY } from "../constant";

/**
 * 打印组件配置选项接口
 */
export interface PrinterOptions {
    /** 数据集合 */
    data: DataCollection;
    /** 打印文档标题（同时用于浏览器打印任务的默认名称），默认「打印预览」 */
    title?: string;
    /** 纸张方向：portrait=纵向（默认），landscape=横向 */
    orientation?: PaperOrientation;
    /** 纸张大小：默认 A4 */
    paperSize?: PaperSize;
    /**
     * 自定义字体的 @font-face CSS 文本（可选，预留接口）。
     * 会与主文档自动收集的 @font-face 规则、registerFontCss() 追加的文本、
     * 全局字体注册表（fontsLoader.registerFontFaceCss）合并去重后，
     * 同时注入打印预览与打印文档，保证自定义字体在预览与打印中一致生效。
     */
    fontFaceCss?: string;
}

export interface PrinterSetting {
    /** 打印文档标题（同时用于浏览器打印任务的默认名称），默认「打印预览」 */
    title?: string;
    /** 纸张方向：portrait=纵向（默认），landscape=横向 */
    orientation?: PaperOrientation;
    /** 纸张大小：默认 A4 */
    paperSize?: PaperSize;
    /** 页边距预设：normal=常规（默认），wide=宽，narrow=窄 */
    marginPreset?: MarginPreset;
    /** 打印缩放模式：none=无缩放（默认），fit-sheet=整个工作表打印在一页，
     * fit-cols=所有列打印在一页，fit-rows=所有行打印在一页，custom=自定义缩放比例 */
    scaleMode?: ScaleMode;
    /** 自定义缩放比例（custom 模式下有效），默认 1 */
    scaleRatio?: number;
    /** 打印方式：simplex=单面打印（默认），duplex-long=双面打印-长边翻页，duplex-short=双面打印-短边翻页 */
    duplex?: DuplexMode;
    /** 页面范围（1-based 页序，如 '1-3,5'；空/未设置=全部页面），仅对实际打印输出生效 */
    pageRange?: string;
    /** 自定义页边距（mm，marginPreset='custom' 时生效） */
    customMargins?: PageMargins;
    /** 打印区域（可选），默认打印整个工作表（0-based，全网格坐标，按内容顺序排列），例如：A1:B10 */
    printArea?: string;
}

/**
 * 纸张方向类型
 */
export type PaperOrientation = 'portrait' | 'landscape';

/**
 * 纸张大小类型
 */
export type PaperSize = 'A4' | 'A3' | 'B5' | 'Letter';

/**
 * 页边距预设类型：normal=常规（默认），wide=宽，narrow=窄，custom=自定义（customMargins）
 */
export type MarginPreset = 'normal' | 'wide' | 'narrow' | 'custom';

/**
 * 打印方式类型：simplex=单面打印（默认），duplex-long=双面打印-长边翻页，duplex-short=双面打印-短边翻页
 */
export type DuplexMode = 'simplex' | 'duplex-long' | 'duplex-short';

/**
 * 打印缩放模式类型：none=无缩放，fit-sheet=整个工作表打印在一页，
 * fit-cols=所有列打印在一页，fit-rows=所有行打印在一页，custom=自定义缩放比例
 */
export type ScaleMode = 'none' | 'fit-sheet' | 'fit-cols' | 'fit-rows' | 'custom';

/**
 * 页边距（mm）：上下左右独立设置
 */
export interface PageMargins {
    top: number;
    bottom: number;
    left: number;
    right: number;
}

/**
 * 页边距预设表（mm，对齐 Excel 标准）：常规=上下19.05/左右17.78，宽=四周25.4，窄=四周6.35
 */
const MARGIN_PRESETS: Record<MarginPreset, PageMargins> = {
    normal: { top: 19.05, bottom: 19.05, left: 17.78, right: 17.78 },
    wide: { top: 25.4, bottom: 25.4, left: 25.4, right: 25.4 },
    narrow: { top: 6.35, bottom: 6.35, left: 6.35, right: 6.35 },
    // custom 的实际值来自 printSetting.customMargins，此处仅作缺省兜底（缺失/非法时回落常规值）
    custom: { top: 19.05, bottom: 19.05, left: 17.78, right: 17.78 },
};

/**
 * 纸张尺寸表：width/height 为纵向时的宽高（mm），page 为打印 @page 规则的 CSS 尺寸关键字
 */
const PAPER_SIZES: Record<PaperSize, { width: number; height: number; page: string }> = {
    A4: { width: 210, height: 297, page: 'A4' },
    A3: { width: 297, height: 420, page: 'A3' },
    B5: { width: 176, height: 250, page: 'B5' },
    Letter: { width: 215.9, height: 279.4, page: 'letter' },
};

/** CSS px 与 mm 的换算系数（96dpi 标准） */
const MM_TO_PX = 96 / 25.4;

/**
 * 垂直分页的单页行几何信息
 */
interface PageGeometry {
    /** 本页包含的可见行索引（0-based，全网格坐标，按内容顺序排列） */
    rows: number[];
    /** 本页各行的高度（px，实际尺寸，与 rows 一一对应） */
    heights: number[];
}

/**
 * 水平分页的单页列组信息
 */
interface ColPage {
    /** 本页包含的可见列索引（0-based，全网格坐标，按内容顺序排列） */
    cols: number[];
}

/**
 * 分页计划中的合并块：数值矩形 + 可见首末行列。
 * 可见首末行列用于判断块体在本页是否被页边界截断（截断边不绘制边框）。
 */
interface PrintBlock {
    /** 锚点单元格名 */
    anchor: string;
    /** 锚点行列（0-based） */
    r0: number;
    c0: number;
    /** 原始跨度 */
    rs: number;
    cs: number;
    /** 块内可见首末行/列（0-based，全网格坐标） */
    firstRow: number;
    lastRow: number;
    firstCol: number;
    lastCol: number;
}

/**
 * 打印分页计划：只描述页面几何与合并块索引，不包含任何单元格内容与 HTML。
 *
 * 预览打开时先算这份计划（复杂度只与行列数、合并块数相关），随后按需渲染当前页，
 * DOM 成本从「整张表所有页」降到「当前一页」。
 */
interface PrintPlan {
    /** 纸张宽高（mm，已按方向换算） */
    paperWmm: number;
    paperHmm: number;
    /** 本次分页使用的缩放系数 */
    renderScale: number;
    /** 参与打印的可见列/行（0-based） */
    visCols: number[];
    visRows: number[];
    /** 可见行列的全局布局（缩放后尺寸） */
    globalX: Map<number, { x: number; w: number }>;
    globalY: Map<number, { y: number; h: number }>;
    /** 水平/垂直分页结果 */
    colPages: ColPage[];
    rowPages: PageGeometry[];
    /** 合并块索引：锚点名 → 数值矩形与可见范围 */
    blocks: Map<string, PrintBlock>;
    /** 总页数（列组数 × 行组数） */
    totalPages: number;
    /** 空白文档（无任何行列）时固定输出一张空白页 */
    blank: boolean;
}

/**
 * 打印组件
 * 将 DataCollection 中的表格数据按纸张分页生成为 HTML 表格页面，提供两个功能模块：
 * - 打印预览（printPreview）：左右结构 —— 左侧为按纸张逐页生成的预览页面（单页显示），
 *   右侧为打印设置区域（打印机、份数、顺序、打印方式、纸张大小、纸张方向、页边距、
 *   缩放、页面范围、打印、关闭，分组间以分割线分隔），底部为「< 第N页/共M页 >」分页导航
 *   与「滑块+输入框」预览缩放控件；重新打开时沿用用户上次设置；
 * - 打印（print）：将各页表格 HTML 写入隐藏 iframe 并调用浏览器打印引擎。
 *
 * 说明：
 * - 页面以 table（table-layout: fixed + colgroup 列宽 + 固定行高）矢量渲染，
 *   文字与边框由浏览器打印引擎按打印机分辨率输出，精度优于 canvas 位图光栅化；
 * - 内容不缩放，按实际尺寸与纸张可打印区域比较：宽度过大时水平分页（列不跨页拆分），
 *   高度过大时垂直分页（行不跨页拆分），总页数 = 列组数 × 行组数；
 * - 隐藏的行/列不参与打印；跨越隐藏行/列的合并块按可见范围收缩合并跨度；
 * - 单元格边框严格按数据中的实际值输出：仅显式设置且大于 0 的边显示边框；
 * - 自定义字体：预览与打印文档自动注入主文档收集到的 @font-face 规则，
 *   并预留 registerFontCss() / 全局 registerFontFaceCss() 接口，
 *   后期用户自定义加入的字体无需改动打印组件即可应用到预览与打印；
 * - 打印设置（标题/页边距/纸张方向/纸张大小/打印区域/打印缩放）：统一读写
 *   data.printSetting（工具栏 set 系列方法与预览面板共用的唯一入口，随文档序列化持久化），
 *   预览打开时即时重绘生效。
 *
 * @class Printer
 * @example
 * const printer = new Printer({ data });
 * printer.printPreview();  // 打印预览
 * printer.print();         // 直接打印
 */
export class Printer {
    /** 数据集合 */
    private data: DataCollection;

    // —— 打印设置统一读自 data.printSetting（工具栏与打印预览共用的唯一入口，随文档序列化持久化）——
    /** 打印文档标题（默认「打印预览」） */
    private get title(): string {
        return this.data.printSetting.title ?? '打印预览';
    }
    /** 纸张方向（默认纵向）；历史遗留的非法值一并回落默认 */
    private get orientation(): PaperOrientation {
        const value = this.data.printSetting.orientation;
        return value === 'landscape' ? value : 'portrait';
    }
    /** 纸张大小（默认 A4）；历史遗留的非法值一并回落默认 */
    private get paperSize(): PaperSize {
        const value = this.data.printSetting.paperSize;
        return value && value in PAPER_SIZES ? value : 'A4';
    }
    /** 页边距预设（默认常规）；历史遗留的非法值一并回落默认 */
    private get marginPreset(): MarginPreset {
        const value = this.data.printSetting.marginPreset;
        return value && value in MARGIN_PRESETS ? value : 'normal';
    }
    /** 打印方式（默认单面打印）；历史遗留的非法值一并回落默认 */
    private get duplex(): DuplexMode {
        const value = this.data.printSetting.duplex;
        return value === 'duplex-long' || value === 'duplex-short' ? value : 'simplex';
    }
    /**
     * 当前页边距（mm）：custom 预设读 printSetting.customMargins（缺失/非法回落常规值），
     * 其余预设读 MARGIN_PRESETS
     * @private
     */
    private get margins(): PageMargins {
        if (this.marginPreset === 'custom') {
            const cm = this.data.printSetting.customMargins;
            if (cm
                && Number.isFinite(cm.top) && cm.top >= 0
                && Number.isFinite(cm.bottom) && cm.bottom >= 0
                && Number.isFinite(cm.left) && cm.left >= 0
                && Number.isFinite(cm.right) && cm.right >= 0) {
                return cm;
            }
            return MARGIN_PRESETS.normal;
        }
        return MARGIN_PRESETS[this.marginPreset];
    }
    /** 打印缩放模式（默认无缩放）；历史遗留的非法值一并回落默认 */
    private get scaleMode(): ScaleMode {
        const value = this.data.printSetting.scaleMode;
        return value === 'fit-sheet' || value === 'fit-cols' || value === 'fit-rows' || value === 'custom' ? value : 'none';
    }
    /** 自定义缩放比例（百分比表示，100=1x） */
    private get customScale(): number {
        const ratio = this.data.printSetting.scaleRatio;
        return (ratio && ratio > 0 ? ratio : 1) * 100;
    }
    /** 打印预览遮罩层（null 表示未打开） */
    private overlay: HTMLDivElement | null = null;
    /** 预览层事件处理器（移除监听用） */
    private boundPreviewKeyDown: (e: KeyboardEvent) => void = () => {};
    /** 预览窗口 resize 处理器（移除监听用）：重算预览页等比缩放 */
    private boundPreviewResize: () => void = () => {};
    /** 预览当前页索引（0-based） */
    private currentPage = 0;
    /** 预览总页数 */
    private totalPages = 0;
    /** 当前分页计划（打开预览/设置变更时重建；只含几何与索引，不含页面 HTML） */
    private _plan: PrintPlan | null = null;
    /** 当前已渲染的预览页元素（预览为单页模式，只保留一页 DOM） */
    private _pageDiv: HTMLElement | null = null;
    /** 设置变更合并刷新定时器（同一轮同步设置变更只重建一次） */
    private _renderTimer = 0;
    /** 渲染任务令牌：关闭预览或再次变更时使旧任务失效 */
    private _renderToken = 0;
    /** 底部页码指示器元素 */
    private pageIndicator: HTMLElement | null = null;
    /** 上一页 / 下一页按钮 */
    private prevBtn: HTMLButtonElement | null = null;
    private nextBtn: HTMLButtonElement | null = null;
    /** 目标打印机名称（浏览器沙箱无法枚举/指定系统打印机，最终在打印对话框中确认） */
    private printerName = '';
    /** 打印份数（默认 1）：Web API 无法预设浏览器对话框份数，此处仅用于面板回显与持久化 */
    private copies = 1;
    /** 打印顺序：page=逐页打印（1,1,2,2,3,3），copies=逐份打印（1,2,3,1,2,3）；同上，仅面板回显 */
    private collate: 'page' | 'copies' = 'page';
    /** 追加注册的自定义字体 @font-face CSS（来自 options.fontFaceCss 与 registerFontCss 调用） */
    private _fontFaceCss: string[] = [];
    /** 汇总后的 @font-face CSS 缓存（样式表/注册表未变化时直接复用，避免每次打开预览全量扫描） */
    private _fontFaceCssCache: string | null = null;
    /** 缓存对应的签名（样式表数量 + 各来源条目数） */
    private _fontFaceCssSignature = '';
    /** 打印任务令牌：新的打印/关闭预览使旧的异步写入任务停止 */
    private _printToken = 0;
    /** 当前打印 iframe（新打印开始时清理旧 iframe） */
    private _printIframe: HTMLIFrameElement | null = null;
    /** 打印准备进度提示元素 */
    private _printProgress: HTMLElement | null = null;
    /** 本次渲染使用的缩放系数（_buildPlan 解析后供 td 字号/行距/字距缩放读取） */
    private renderScale = 1;
    /** 预览容器与设置面板控件引用（打印设置变更时同步回显并重绘） */
    private _pagesContainer: HTMLElement | null = null;
    private _paperSizeSelect: HTMLSelectElement | null = null;
    private _orientationSelect: HTMLSelectElement | null = null;
    private _duplexSelect: HTMLSelectElement | null = null;
    private _marginSelect: HTMLSelectElement | null = null;
    private _scaleSelect: HTMLSelectElement | null = null;
    private _scaleInput: HTMLInputElement | null = null;
    private _pageRangeSelect: HTMLSelectElement | null = null;
    private _pageRangeInput: HTMLInputElement | null = null;
    private _pageRangeRow: HTMLElement | null = null;
    /** 页脚预览缩放滑块组件（common/slider 共用组件：轨道 + 百分比输入框 + 单位） */
    private _zoomSlider: Slider | null = null;
    /** 预览缩放手动覆盖值（百分比，10-400；null=自动按可用高度适配） */
    private _previewZoomPercent: number | null = null;

    /**
     * 创建打印组件实例
     * @param {PrinterOptions} options - 配置选项
     */
    constructor(options: PrinterOptions) {
        this.data = options.data;
        // 构造参数作为打印设置初始值写入统一入口 data.printSetting（未提供时用接口默认值）
        if (options.title) this.data.printSetting.title = options.title;
        if (options.orientation) this.data.printSetting.orientation = options.orientation;
        if (options.paperSize) this.data.printSetting.paperSize = options.paperSize;
        if (options.fontFaceCss?.trim()) this._fontFaceCss.push(options.fontFaceCss.trim());
    }

    /**
     * 追加注册自定义字体的 @font-face CSS（预留接口，可多次调用，重复文本自动去重）。
     * 注册后于每次打开打印预览 / 生成打印文档时，与主文档自动收集的 @font-face 规则、
     * 全局字体注册表（fontsLoader.registerFontFaceCss）合并注入两个文档，
     * 后期用户自定义加入的字体无需改动打印组件即可自动应用到预览与打印。
     * @param {string} css - 完整的 @font-face 规则文本，可一次传入多条规则
     * @example
     * printer.registerFontCss(`@font-face { font-family: 'MyFont'; src: url('/fonts/myfont.woff2') format('woff2'); }`);
     */
    public registerFontCss(css: string): void {
        const trimmed = css?.trim();
        if (trimmed && !this._fontFaceCss.includes(trimmed)) {
            this._fontFaceCss.push(trimmed);
            this._fontFaceCssCache = null;
        }
    }

    /**
     * 汇总待注入的 @font-face CSS（预览与打印文档共用），来源与优先级从低到高：
     * 1. 主文档 document.styleSheets 中自动收集的全部 @font-face 规则
     *    （覆盖 index.scss 等打包产物中的字体声明，后期以 CSS 声明的新增字体零配置生效）；
     * 2. 本实例追加注册的 CSS（options.fontFaceCss / registerFontCss）；
     * 3. 全局字体注册表的 CSS（fontsLoader.registerFontFaceCss，供任意模块登记动态加载的字体）。
     * 合并去重；同名字体家族后者覆盖前者。
     * @returns {string} 合并后的 CSS 文本（无任何字体时为空串）
     * @private
     */
    private _getFontFaceCss(): string {
        const registered = getRegisteredFontFaceCss();
        // 签名使用各来源的条目数量：样式表/字体注册数量不变时复用缓存，
        // 避免每次打开预览、每次打印都全量遍历 document.styleSheets 的规则
        const signature = `${document.styleSheets.length}|${this._fontFaceCss.length}|${registered.length}`;
        if (this._fontFaceCssCache !== null && signature === this._fontFaceCssSignature) {
            return this._fontFaceCssCache;
        }
        const collected: string[] = [];
        for (const sheet of Array.from(document.styleSheets)) {
            try {
                for (const rule of Array.from((sheet as CSSStyleSheet).cssRules)) {
                    if (rule instanceof CSSFontFaceRule) collected.push(rule.cssText.trim());
                }
            } catch {
                // 跨域样式表不允许访问 cssRules，跳过即可
            }
        }
        this._fontFaceCssSignature = signature;
        this._fontFaceCssCache = [...new Set([...collected, ...this._fontFaceCss, ...registered])].join('\n');
        return this._fontFaceCssCache;
    }

    /**
     * 页边距 CSS padding 简写（mm，上 右 下 左），预览页面与打印文档共用
     * @returns {string} padding 值
     * @private
     */
    private _marginCss(): string {
        const m = this.margins;
        return `${m.top}mm ${m.right}mm ${m.bottom}mm ${m.left}mm`;
    }

    /**
     * 预览打开时同步设置面板控件回显并重绘当前页；未打开时不做任何事。
     * 重绘通过 {@link _scheduleRender} 合并到下一轮事件循环，连续设置变更只重建一次。
     * @private
     */
    private _refreshPreview(): void {
        if (!this.overlay) return;
        this._syncPanelControls();
        this._scheduleRender(0);
    }

    /**
     * 合并刷新：延迟重建分页计划并渲染指定页。
     *
     * 每次调用递增令牌，旧任务（包括拖动设置时的连续变更）自动失效；
     * 同时给遮罩层留出一帧绘制时间，避免「点击预览 → 主线程被占满 → 白屏」。
     * @param {number} page - 目标页索引（0-based）
     * @private
     */
    private _scheduleRender(page: number): void {
        const token = ++this._renderToken;
        if (this._renderTimer) window.clearTimeout(this._renderTimer);
        this._renderTimer = window.setTimeout(() => {
            this._renderTimer = 0;
            if (token !== this._renderToken || !this.overlay) return;
            this._plan = this._buildPlan();
            this.totalPages = this._plan.totalPages;
            this._goToPage(page);
        }, 0);
    }

    /**
     * 在预览区显示加载占位（分页计划与首屏页面准备期间）
     * @private
     */
    private _showPreviewLoading(): void {
        const host = this._pagesContainer;
        if (!host) return;
        const loading = document.createElement('div');
        loading.className = 'gs-print-loading';
        loading.textContent = '正在生成打印预览…';
        host.replaceChildren(loading);
        this._pageDiv = null;
    }

    /**
     * 同步设置面板各控件回显（预览打开期间工具栏或程序化设置变更时保持一致）；
     * 输入框聚焦编辑中时跳过回显，避免覆盖用户输入
     * @private
     */
    private _syncPanelControls(): void {
        if (this._paperSizeSelect) this._paperSizeSelect.value = this.paperSize;
        if (this._orientationSelect) this._orientationSelect.value = this.orientation;
        if (this._duplexSelect) this._duplexSelect.value = this.duplex;
        if (this._marginSelect) this._marginSelect.value = this.marginPreset;
        if (this._scaleSelect) this._scaleSelect.value = this.scaleMode;
        if (this._scaleInput) {
            this._scaleInput.value = String(Math.round(this.customScale));
            this._scaleInput.disabled = this.scaleMode !== 'custom';
        }
        const range = (this.data.printSetting.pageRange ?? '').trim();
        if (this._pageRangeSelect) this._pageRangeSelect.value = range ? 'custom' : 'all';
        if (this._pageRangeRow) this._pageRangeRow.style.display = range ? '' : 'none';
        if (this._pageRangeInput && document.activeElement !== this._pageRangeInput) this._pageRangeInput.value = range;
    }

    /** @returns {MarginPreset} 当前页边距预设 */
    public getMarginPreset(): MarginPreset { return this.marginPreset; }
    /** @returns {PaperOrientation} 当前纸张方向 */
    public getOrientation(): PaperOrientation { return this.orientation; }
    /** @returns {PaperSize} 当前纸张大小 */
    public getPaperSize(): PaperSize { return this.paperSize; }
    /** @returns {ScaleMode} 当前打印缩放模式 */
    public getScaleMode(): ScaleMode { return this.scaleMode; }
    /** @returns {number} 自定义缩放比例（百分比） */
    public getCustomScale(): number { return this.customScale; }
    /** @returns {boolean} 是否已设置打印区域 */
    public isPrintAreaSet(): boolean { return this.printAreaRange !== null; }

    /**
     * 解析打印区域（统一读自 data.printSetting.printArea，如 'A1:F20' 或 'B3'）
     * 为 1-based 含端点行列范围；未设置或无法解析时返回 null
     * @private
     */
    private get printAreaRange(): { startCol: number; startRow: number; endCol: number; endRow: number } | null {
        const range = this.data.printSetting.printArea;
        if (!range) return null;
        const [start, end] = range.split(':');
        if (!start) return null;
        const s = this.data.getCellColAndRow(start);
        const e = this.data.getCellColAndRow(end ?? start);
        return {
            startCol: Math.min(s.col, e.col),
            startRow: Math.min(s.row, e.row),
            endCol: Math.max(s.col, e.col),
            endRow: Math.max(s.row, e.row),
        };
    }

    /**
     * @returns {string} 打印区域状态文案（工具栏菜单回显用）：已设置返回当前范围，未设置返回提示
     */
    public getPrintAreaLabel(): string {
        const range = this.data.printSetting.printArea;
        return range ? `当前打印区域：${range}` : '未设置打印区域（打印全部内容）';
    }

    /**
     * 设置页边距预设（写入 data.printSetting），预览打开时即时重绘；
     * 页边距影响纸张可打印区域，同步清除预览缩放手动覆盖值（恢复自动适配）
     * @param {MarginPreset} preset - 页边距预设
     */
    public setMarginPreset(preset: MarginPreset): void {
        this.data.printSetting.marginPreset = preset;
        this._previewZoomPercent = null;
        this._refreshPreview();
    }

    /**
     * 设置纸张方向（写入 data.printSetting），预览打开时即时重绘；
     * 纸张几何变化，同步清除预览缩放手动覆盖值（恢复自动适配）
     * @param {PaperOrientation} orientation - 纸张方向
     */
    public setOrientation(orientation: PaperOrientation): void {
        this.data.printSetting.orientation = orientation;
        this._previewZoomPercent = null;
        this._refreshPreview();
    }

    /**
     * 设置纸张大小（写入 data.printSetting），预览打开时即时重绘；
     * 纸张几何变化，同步清除预览缩放手动覆盖值（恢复自动适配）
     * @param {PaperSize} size - 纸张大小
     */
    public setPaperSize(size: PaperSize): void {
        this.data.printSetting.paperSize = size;
        this._previewZoomPercent = null;
        this._refreshPreview();
    }

    /**
     * 设置打印缩放模式（写入 data.printSetting）；custom 模式可同时指定缩放比例
     * @param {ScaleMode} mode - 缩放模式
     * @param {number} [customPercent] - 自定义缩放比例（百分比，10-400，越界收敛；内部以 ratio 存储，1=100%）
     */
    public setScaleMode(mode: ScaleMode, customPercent?: number): void {
        this.data.printSetting.scaleMode = mode;
        if (mode === 'custom' && customPercent !== undefined && Number.isFinite(customPercent)) {
            this.data.printSetting.scaleRatio = Math.min(4, Math.max(0.1, Math.round(customPercent) / 100));
        }
        this._refreshPreview();
    }

    /**
     * 设置打印方式（写入 data.printSetting）。
     * 双面翻页方向由打印机驱动执行，Web 打印 API 无法预设浏览器打印对话框，
     * 此处仅记录用户偏好（与打印机/份数/顺序同一处理方式）。
     * @param {DuplexMode} mode - 打印方式
     */
    public setDuplex(mode: DuplexMode): void {
        this.data.printSetting.duplex = mode;
    }

    /**
     * 设置打印区域：以当前选区（无选区时为活动单元格）的矩形范围为打印区域，
     * 仅区域内的行列参与分页打印
     */
    public setPrintArea(): void {
        const range = this.data.selection ?? this.data.activedCell;
        if (!range) return;
        this.data.printSetting.printArea = range;
        this._refreshPreview();
    }

    /**
     * 取消打印区域，恢复打印全部内容
     */
    public clearPrintArea(): void {
        this.data.printSetting.printArea = undefined;
        this._refreshPreview();
    }

    /**
     * 解析打印缩放系数：none=1；custom=自定义百分比（10%-400）；
     * fit-cols/fit-rows/fit-sheet=可打印区域与内容实际尺寸之比（只缩小不放大，下限 0.1）
     * @param {number} totalW - 内容总宽（px，实际尺寸）
     * @param {number} totalH - 内容总高（px，实际尺寸）
     * @param {number} printableW - 可打印宽度（px）
     * @param {number} printableH - 可打印高度（px）
     * @returns {number} 缩放系数
     * @private
     */
    private _resolveScale(totalW: number, totalH: number, printableW: number, printableH: number): number {
        const fit = (total: number, printable: number): number =>
            total > 0 && printable > 0 ? Math.max(0.1, Math.min(1, printable / total)) : 1;
        switch (this.scaleMode) {
            case 'custom': return Math.min(4, Math.max(0.1, this.customScale / 100));
            case 'fit-cols': return fit(totalW, printableW);
            case 'fit-rows': return fit(totalH, printableH);
            case 'fit-sheet': return Math.min(fit(totalW, printableW), fit(totalH, printableH));
            default: return 1;
        }
    }

    /**
     * 打印预览：上方仅显示粗体「打印预览」标题；主体 .gs-print-body（横向 flex）
     * 内含左列 .gs-print-main（纵向：按纸张逐页生成的预览页面，一次显示一页 +
     * 底部页脚「< 第N页/共M页 >」与右侧预览缩放控件（滑块 + 百分比输入框，
     * 手动值优先，纸张/方向/页边距变更后恢复自动适配），页脚仅覆盖预览区宽度）
     * 与右侧打印设置区域（独占全高直达页面底部，行内 label 与控件水平对齐，
     * 分组间以分割线分隔：打印机、份数、顺序、打印方式、纸张大小、纸张方向、
     * 页边距（预设下拉 +「调整页边距」自定义编辑器）、缩放（模式下拉 +
     * 自定义百分比输入框）、页面范围（全部/页码范围），设置变更即时重绘预览）。
     * 重新打开预览时自动沿用用户上次设置。
     */
    public printPreview(): void {
        // 重复调用时先关闭已有预览，避免遮罩层叠加
        if (this.overlay) this.closePreview();

        const overlay = document.createElement('div');
        overlay.className = 'gs-print-overlay';
        overlay.innerHTML = buildPrintPreviewShellHtml({
            extraCss: this._tableStyles(),
            fontFaceCss: this._getFontFaceCss(),
        });
        document.body.appendChild(overlay);
        // 先登记遮罩层：后续渲染被延迟到下一轮事件循环，期间关闭预览/重开需要能正确清理
        this.overlay = overlay;

        // 打印设置：回显当前纸张大小/方向（新开预览与用户上次设置一致），变更即时重绘预览页；
        // 控件引用存入实例字段，供工具栏设置变更时（_refreshPreview）同步回显
        const pagesContainer = overlay.querySelector<HTMLElement>('.gs-print-pages');
        const paperSizeSelect = overlay.querySelector<HTMLSelectElement>('[data-gs-print-role="paper-size"]');
        const orientationSelect = overlay.querySelector<HTMLSelectElement>('[data-gs-print-role="orientation"]');
        this._pagesContainer = pagesContainer;
        this._paperSizeSelect = paperSizeSelect;
        this._orientationSelect = orientationSelect;
        if (paperSizeSelect) {
            paperSizeSelect.value = this.paperSize;
            paperSizeSelect.addEventListener('change', () => {
                this.setPaperSize(paperSizeSelect.value as PaperSize);
            });
        }
        if (orientationSelect) {
            orientationSelect.value = this.orientation;
            orientationSelect.addEventListener('change', () => {
                this.setOrientation(orientationSelect.value as PaperOrientation);
            });
        }
        // 打印机选择：填充选项并记录用户偏好（Web 沙箱无法枚举/指定系统打印机，
        // 实际目标打印机在打印对话框中确认）
        const printerSelect = overlay.querySelector<HTMLSelectElement>('[data-gs-print-role="printer"]');
        if (printerSelect) {
            this._fillPrinterOptions(printerSelect);
            printerSelect.addEventListener('change', () => {
                this.printerName = printerSelect.value;
            });
        }
        // 份数：输入校验 + 步进按钮（鼠标按下立即执行并长按连发，抬起/移出即停止）
        const copiesInput = overlay.querySelector<HTMLInputElement>('[data-gs-print-role="copies"]');
        if (copiesInput) {
            copiesInput.value = String(this.copies);
            const applyCopies = (n: number) => {
                this.copies = Number.isFinite(n) ? Math.min(999, Math.max(1, Math.floor(n))) : 1;
                copiesInput.value = String(this.copies);
            };
            copiesInput.addEventListener('change', () => applyCopies(Number(copiesInput.value)));

            let holdTimer = 0;
            let repeatTimer = 0;
            const stopHold = () => {
                window.clearTimeout(holdTimer);
                window.clearInterval(repeatTimer);
            };
            overlay.querySelectorAll<HTMLButtonElement>('[data-gs-print-step]').forEach(btn => {
                const delta = btn.dataset.gsPrintStep === 'inc' ? 1 : -1;
                btn.addEventListener('mousedown', (e) => {
                    e.preventDefault();  // 阻止聚焦拖选，保证按住期间状态稳定
                    applyCopies(this.copies + delta);
                    stopHold();
                    holdTimer = window.setTimeout(() => {
                        repeatTimer = window.setInterval(() => applyCopies(this.copies + delta), 80);
                    }, 320);
                });
                // 抬起或移出按钮即停止连发
                btn.addEventListener('mouseup', stopHold);
                btn.addEventListener('mouseleave', stopHold);
            });
        }
        // 顺序：逐页打印 / 逐份打印
        const collateSelect = overlay.querySelector<HTMLSelectElement>('[data-gs-print-role="collate"]');
        if (collateSelect) {
            collateSelect.value = this.collate;
            collateSelect.addEventListener('change', () => {
                this.collate = collateSelect.value as 'page' | 'copies';
            });
        }
        // 打印方式：单面 / 双面（长边翻页、短边翻页），偏好写入 data.printSetting
        const duplexSelect = overlay.querySelector<HTMLSelectElement>('[data-gs-print-role="duplex"]');
        this._duplexSelect = duplexSelect;
        if (duplexSelect) {
            duplexSelect.value = this.duplex;
            duplexSelect.addEventListener('change', () => {
                this.setDuplex(duplexSelect.value as DuplexMode);
            });
        }
        // 页边距：预设下拉 + 「调整页边距」自定义编辑器（上下左右 mm，应用后 marginPreset='custom'）
        const marginEditor = overlay.querySelector<HTMLElement>('[data-gs-print-role="margin-editor"]');
        const marginInputs = {
            top: overlay.querySelector<HTMLInputElement>('[data-gs-print-role="margin-top"]'),
            bottom: overlay.querySelector<HTMLInputElement>('[data-gs-print-role="margin-bottom"]'),
            left: overlay.querySelector<HTMLInputElement>('[data-gs-print-role="margin-left"]'),
            right: overlay.querySelector<HTMLInputElement>('[data-gs-print-role="margin-right"]'),
        };
        // 打开编辑器：四向输入框填入当前实际页边距
        const openMarginEditor = () => {
            const m = this.margins;
            if (marginInputs.top) marginInputs.top.value = String(m.top);
            if (marginInputs.bottom) marginInputs.bottom.value = String(m.bottom);
            if (marginInputs.left) marginInputs.left.value = String(m.left);
            if (marginInputs.right) marginInputs.right.value = String(m.right);
            if (marginEditor) marginEditor.style.display = '';
            marginInputs.top?.focus();
        };
        const hideMarginEditor = () => { if (marginEditor) marginEditor.style.display = 'none'; };
        const marginSelect = overlay.querySelector<HTMLSelectElement>('[data-gs-print-role="margin"]');
        this._marginSelect = marginSelect;
        if (marginSelect) {
            marginSelect.value = this.marginPreset;
            marginSelect.addEventListener('change', () => {
                const preset = marginSelect.value as MarginPreset;
                if (preset === 'custom') {
                    openMarginEditor();  // 直接选「自定义」：打开编辑器，应用前保持原页边距
                } else {
                    hideMarginEditor();
                    this.setMarginPreset(preset);
                }
            });
        }
        // 「调整页边距」按钮：展开/收起编辑器（收起时下拉回显当前预设）
        overlay.querySelector<HTMLButtonElement>('[data-gs-print-action="margin-adjust"]')
            ?.addEventListener('click', () => {
                if (marginEditor && marginEditor.style.display !== 'none') {
                    hideMarginEditor();
                    if (marginSelect) marginSelect.value = this.marginPreset;
                } else {
                    openMarginEditor();
                }
            });
        // 应用：读取四向输入（mm，非法按 0、保留 1 位小数），写入 customMargins 并置预设为 custom
        overlay.querySelector<HTMLButtonElement>('[data-gs-print-action="margin-apply"]')
            ?.addEventListener('click', () => {
                const num = (input: HTMLInputElement | null): number => {
                    const n = Number(input?.value);
                    return Number.isFinite(n) ? Math.max(0, Math.round(n * 10) / 10) : 0;
                };
                this.data.printSetting.customMargins = {
                    top: num(marginInputs.top),
                    bottom: num(marginInputs.bottom),
                    left: num(marginInputs.left),
                    right: num(marginInputs.right),
                };
                this.data.printSetting.marginPreset = 'custom';
                this._previewZoomPercent = null;  // 页边距变化影响自动适配，恢复适配缩放
                this._refreshPreview();           // 内部同步 marginSelect 显示「自定义」并重绘
                hideMarginEditor();
            });
        // 取消：仅收起编辑器并回显当前预设（不写入任何设置）
        overlay.querySelector<HTMLButtonElement>('[data-gs-print-action="margin-cancel"]')
            ?.addEventListener('click', () => {
                hideMarginEditor();
                if (marginSelect) marginSelect.value = this.marginPreset;
            });
        // 缩放：模式下拉 + 自定义百分比输入框（仅 custom 模式可编辑，10-400）
        const scaleSelect = overlay.querySelector<HTMLSelectElement>('[data-gs-print-role="scale"]');
        const scaleInput = overlay.querySelector<HTMLInputElement>('[data-gs-print-role="scale-input"]');
        this._scaleSelect = scaleSelect;
        this._scaleInput = scaleInput;
        const syncScaleControls = () => {
            if (scaleSelect) scaleSelect.value = this.scaleMode;
            if (scaleInput) {
                scaleInput.value = String(Math.round(this.customScale));
                scaleInput.disabled = this.scaleMode !== 'custom';
            }
        };
        syncScaleControls();
        if (scaleSelect) {
            scaleSelect.addEventListener('change', () => {
                this.setScaleMode(scaleSelect.value as ScaleMode);
                syncScaleControls();
                if (scaleSelect.value === 'custom') scaleInput?.select();
            });
        }
        if (scaleInput) {
            scaleInput.addEventListener('change', () => {
                const n = Number(scaleInput.value);
                if (Number.isFinite(n) && n > 0) this.setScaleMode('custom', n);
                syncScaleControls();  // 回显收敛后的值（越界收敛到 10-400）
            });
        }
        // 页面范围：全部 / 页码范围（如 1-3,5），仅对实际打印输出生效（预览始终显示全部页）
        const rangeSelect = overlay.querySelector<HTMLSelectElement>('[data-gs-print-role="page-range"]');
        const rangeInput = overlay.querySelector<HTMLInputElement>('[data-gs-print-role="page-range-input"]');
        const rangeRow = overlay.querySelector<HTMLElement>('[data-gs-print-role="page-range-row"]');
        this._pageRangeSelect = rangeSelect;
        this._pageRangeInput = rangeInput;
        this._pageRangeRow = rangeRow;
        const syncRangeControls = () => {
            const range = (this.data.printSetting.pageRange ?? '').trim();
            if (rangeSelect) rangeSelect.value = range ? 'custom' : 'all';
            if (rangeRow) rangeRow.style.display = range ? '' : 'none';
            if (rangeInput && document.activeElement !== rangeInput) rangeInput.value = range;
        };
        syncRangeControls();
        if (rangeSelect) {
            rangeSelect.addEventListener('change', () => {
                if (rangeSelect.value === 'custom') {
                    if (rangeRow) rangeRow.style.display = '';
                    rangeInput?.focus();
                } else {
                    this.data.printSetting.pageRange = undefined;
                    if (rangeRow) rangeRow.style.display = 'none';
                }
            });
        }
        if (rangeInput) {
            rangeInput.addEventListener('change', () => {
                this.data.printSetting.pageRange = rangeInput.value.trim() || undefined;
            });
        }
        // 底部分页导航：页码指示器与上一页/下一页按钮
        this.pageIndicator = overlay.querySelector<HTMLElement>('.gs-print-indicator');
        this.prevBtn = overlay.querySelector<HTMLButtonElement>('[data-gs-print-nav="prev"]');
        this.nextBtn = overlay.querySelector<HTMLButtonElement>('[data-gs-print-nav="next"]');
        // 预览缩放：页脚共用滑块组件（common/slider：轨道 + 百分比输入框，10-400），
        // 覆盖自动适配（null=自动按可用高度适配）；拖动中不回写轨道，避免打断原生手势
        const zoomHost = overlay.querySelector<HTMLElement>('[data-gs-print-role="zoom"]');
        this._zoomSlider = zoomHost ? new Slider({
            parentElement: zoomHost,
            min: 10,
            max: 400,
            step: 1,
            unit: '%',
            sliderWidth: '120px',
            accentColor: '#0d6efd',
            onChange: (value: number, source: SliderChangeSource) => {
                this._previewZoomPercent = value;
                // 轨道拖动中传 false：程序回写轨道 value 会打断原生拖拽手势（滑块卡住不跟手），
                // 此时仅同步数值输入框；数值框提交则回写完整控件状态
                this._applyPreviewZoom(source !== 'track');
            },
        }) : null;
        // 首屏：先挂加载占位，下一轮事件循环再计算分页计划并只渲染当前页。
        // 全部页面 HTML 与 DOM 不再一次性生成，大表打开预览不再造成长时间主线程阻塞。
        this._showPreviewLoading();
        this._scheduleRender(0);

        // 事件委托：设置区打印/关闭按钮 + 底部分页导航按钮
        overlay.addEventListener('click', (e) => {
            const target = (e.target as HTMLElement).closest('[data-gs-print-action], [data-gs-print-nav]') as HTMLElement | null;
            if (!target) return;
            if (target.dataset.gsPrintAction === 'print') this.print();
            if (target.dataset.gsPrintAction === 'close') this.closePreview();
            if (target.dataset.gsPrintNav === 'prev') this._goToPage(this.currentPage - 1);
            if (target.dataset.gsPrintNav === 'next') this._goToPage(this.currentPage + 1);
        });

        // Esc 键：打印准备期间先终止生成（此时打印对话框尚未出现），否则关闭预览
        this.boundPreviewKeyDown = (e: KeyboardEvent) => {
            if (e.key !== 'Escape') return;
            if (this._printProgress) {
                this._cancelPrint();
                return;
            }
            this.closePreview();
        };
        document.addEventListener('keydown', this.boundPreviewKeyDown);

        // 窗口尺寸变化：仅重算预览页等比缩放（高度始终填满可用区域），不重建内容、不跳页
        this.boundPreviewResize = () => this._applyPreviewZoom();
        window.addEventListener('resize', this.boundPreviewResize);
    }

    /**
     * 关闭打印预览遮罩层并清理事件监听
     */
    public closePreview(): void {
        // 打印准备尚未完成时关闭预览：一并终止分批生成并收起进度提示
        if (this._printProgress) this._cancelPrint();
        // 使延迟渲染任务失效，避免关闭后仍在后台重建
        this._renderToken++;
        if (this._renderTimer) {
            window.clearTimeout(this._renderTimer);
            this._renderTimer = 0;
        }
        this.overlay?.remove();
        this.overlay = null;
        this._pagesContainer = null;
        this._paperSizeSelect = null;
        this._orientationSelect = null;
        this._duplexSelect = null;
        this._marginSelect = null;
        this._scaleSelect = null;
        this._scaleInput = null;
        this._pageRangeSelect = null;
        this._pageRangeInput = null;
        this._pageRangeRow = null;
        this._zoomSlider?.destroy();
        this._zoomSlider = null;
        this._previewZoomPercent = null;
        this._plan = null;
        this._pageDiv = null;
        this.pageIndicator = null;
        this.prevBtn = null;
        this.nextBtn = null;
        this.currentPage = 0;
        this.totalPages = 0;
        document.removeEventListener('keydown', this.boundPreviewKeyDown);
        window.removeEventListener('resize', this.boundPreviewResize);
    }

    /**
     * 打印：按当前纸张设置生成各页表格 HTML，写入隐藏 iframe 并调用浏览器打印引擎，
     * 不干扰当前页面，打印结束（afterprint）后自动移除 iframe。
     *
     * 与预览一致复用分页计划；只生成「页面范围」选中的页，并按时间片分批写入 iframe，
     * 避免一次性生成全量 HTML 造成长时间主线程阻塞。准备期间显示进度提示。
     *
     * 页面数量约定：**一张输出页对应一张预览页**（一个 .gs-print-sheet = 一张纸），
     * 因此浏览器打印对话框的页数与预览页数一致；不在文档内按份数复制页面。
     * 份数、打印方式（双面）与目标打印机无法通过 Web API 预设，需在对话框中选择。
     */
    public print(): void {
        const plan = this._buildPlan();
        const selected = this._selectedPageIndices(plan.totalPages);
        if (selected.length === 0) return;
        const sheetCount = selected.length;

        // 新的打印任务：使上一次未完成的异步写入失效并清理其 iframe/进度提示
        const token = ++this._printToken;
        this._clearPrintProgress();
        this._printIframe?.remove();

        // 借助 iframe 隔离打印文档，避免当前页面样式与结构干扰
        const iframe = document.createElement('iframe');
        iframe.style.position = 'fixed';
        iframe.style.right = '0';
        iframe.style.bottom = '0';
        iframe.style.width = '0';
        iframe.style.height = '0';
        iframe.style.border = '0';
        document.body.appendChild(iframe);
        this._printIframe = iframe;

        const win = iframe.contentWindow;
        if (!win) {
            iframe.remove();
            this._printIframe = null;
            return;
        }

        // 打印结束后清理 iframe（afterprint 兼容性兜底：120s 后强制移除）。
        // 用户在打印对话框中取消同样触发 afterprint：立即终止仍在进行的分批生成并收起进度提示。
        win.addEventListener('afterprint', () => {
            if (this._printIframe === iframe) {
                this._cancelPrint();
                return;
            }
            iframe.remove();
        });
        const removeIframeLater = (): void => {
            // 仍在准备（进度提示未消失）说明大表分批生成尚未结束，顺延检查，避免误删正在写入的 iframe
            if (this._printProgress) {
                window.setTimeout(removeIframeLater, 30000);
                return;
            }
            if (this._printIframe === iframe) {
                iframe.remove();
                this._printIframe = null;
            }
        };
        window.setTimeout(removeIframeLater, 120000);

        const doc = win.document;
        doc.open();
        doc.write(this._buildPrintHead());
        doc.close();

        this._showPrintProgress(sheetCount);

        // 空白文档：预览显示一张空白纸，打印也要输出且只输出一张空白页。
        // 零高度容器可能被判为无内容而不产生页面，故给到一个「可打印区高度」，
        // 该值不超过 .gs-print-sheet 的 max-height，不会溢出成两页。
        const blankSheetHmm = Math.max(1, plan.paperHmm - this.margins.top - this.margins.bottom - 0.5);

        let cursor = 0;
        const writeBatch = (): void => {
            if (token !== this._printToken) {
                iframe.remove();
                return;
            }
            // 每批最多占用约 12ms，随后让出主线程，保证进度提示与界面可响应
            const start = Date.now();
            const chunks: string[] = [];
            while (cursor < sheetCount && Date.now() - start < 12) {
                const pageIndex = selected[cursor];
                chunks.push(plan.blank
                    ? `<div class="gs-print-sheet" style="height:${blankSheetHmm}mm"></div>`
                    : `<div class="gs-print-sheet">${this._renderPageHtml(plan, pageIndex)}</div>`);
                cursor++;
            }
            if (chunks.length > 0) {
                doc.body.insertAdjacentHTML('beforeend', chunks.join(''));
            }
            this._setPrintProgress(cursor, sheetCount);
            if (cursor < sheetCount) {
                window.setTimeout(writeBatch, 0);
                return;
            }
            // 等待文档解析与图片加载完成后再唤起打印对话框
            window.setTimeout(() => {
                if (token !== this._printToken) return;
                this._clearPrintProgress();
                win.focus();
                win.print();
            }, 50);
        };
        writeBatch();
    }

    /**
     * 解析「页面范围」设置，返回参与打印的页索引（0-based，升序）。
     * 未设置/无法解析时返回全部页；页码范围只影响打印输出，不影响预览页数。
     * @param {number} total - 总页数
     * @returns {number[]} 参与打印的页索引（0-based）
     * @private
     */
    private _selectedPageIndices(total: number): number[] {
        if (total <= 0) return [];
        const hits = this._parsePageRange(total);
        if (!hits) {
            const all: number[] = [];
            for (let i = 0; i < total; i++) all.push(i);
            return all;
        }
        return Array.from(hits).sort((a, b) => a - b).map(p => p - 1);
    }

    /**
     * 显示打印准备进度提示（固定居中，打印对话框出现前移除）。
     * 大表分批生成耗时较长，提示条内置「取消」按钮，随时可终止生成。
     * @param {number} total - 总纸张数
     * @private
     */
    private _showPrintProgress(total: number): void {
        this._clearPrintProgress();
        const el = document.createElement('div');
        el.style.cssText = 'position:fixed;left:50%;top:50%;transform:translate(-50%,-50%);z-index:10000;'
            + 'display:flex;align-items:center;gap:1rem;'
            + 'background:rgba(33,37,41,.92);color:#fff;padding:.75rem 1.25rem;border-radius:.5rem;'
            + 'font:14px system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;box-shadow:0 8px 24px rgba(0,0,0,.25);';
        const text = document.createElement('span');
        text.className = 'gs-print-progress-text';
        text.textContent = `正在准备打印… 0/${total}`;
        const cancel = document.createElement('button');
        cancel.type = 'button';
        cancel.className = 'gs-print-progress-cancel';
        cancel.textContent = '取消';
        cancel.style.cssText = 'cursor:pointer;border:1px solid rgba(255,255,255,.45);border-radius:.3rem;'
            + 'background:transparent;color:#fff;font:inherit;padding:.15rem .6rem;';
        cancel.addEventListener('click', () => this._cancelPrint());
        el.append(text, cancel);
        document.body.appendChild(el);
        this._printProgress = el;
    }

    /**
     * 更新打印进度提示
     * @param {number} done - 已完成纸张数
     * @param {number} total - 总纸张数
     * @private
     */
    private _setPrintProgress(done: number, total: number): void {
        const text = this._printProgress?.querySelector<HTMLElement>('.gs-print-progress-text');
        if (text) text.textContent = `正在准备打印… ${done}/${total}`;
    }

    /**
     * 取消打印准备：使进行中的分批写入任务失效、移除隐藏 iframe 并立即收起进度提示。
     *
     * 触发来源：进度提示条上的「取消」按钮、准备期间按 Esc、关闭预览、
     * 以及打印对话框被取消（afterprint）。
     * @private
     */
    private _cancelPrint(): void {
        this._printToken++;
        this._printIframe?.remove();
        this._printIframe = null;
        this._clearPrintProgress();
    }

    /**
     * 移除打印进度提示
     * @private
     */
    private _clearPrintProgress(): void {
        this._printProgress?.remove();
        this._printProgress = null;
    }

    /**
     * 将当前页渲染进预览容器。
     *
     * 预览为单页显示模式：页面几何已在分页计划中算好，这里只生成「当前页」的 HTML，
     * 并替换容器内已有页面；翻页/设置变更时按需重绘，绝不再一次性创建全部页 DOM。
     * 页面与实际纸张等比例：宽高按当前纸张大小/方向内联指定（mm），整页以 zoom 等比缩放。
     * @private
     */
    private _renderCurrentPage(): void {
        const container = this._pagesContainer;
        this._pageDiv = null;
        if (!container || !this._plan) return;
        if (this._plan.totalPages <= 0) {
            container.replaceChildren();
            this._applyPreviewZoom();
            return;
        }
        const pageDiv = document.createElement('div');
        pageDiv.className = 'gs-print-page';
        pageDiv.style.width = `${this._plan.paperWmm}mm`;
        pageDiv.style.height = `${this._plan.paperHmm}mm`;
        pageDiv.style.padding = this._marginCss();
        pageDiv.innerHTML = this._renderPageHtml(this._plan, this.currentPage);
        container.replaceChildren(pageDiv);
        this._pageDiv = pageDiv;
        this._applyPreviewZoom();
    }

    /**
     * 计算预览页等比缩放系数：白色纸张区域高度填满滚动区（.gs-print-scroll）的
     * 可用高度（clientHeight 减去上下 padding），宽度与内容随同一系数等比缩放。
     * 布局未就绪（clientHeight 为 0）或纸张自然高度非法时返回 1（不缩放）。
     * @returns {number} 缩放系数（可为 >1 的放大值）
     * @private
     */
    private _calcPreviewZoom(): number {
        const scroll = this._pagesContainer?.parentElement;
        if (!scroll) return 1;
        const scrollStyle = getComputedStyle(scroll);
        const availH = scroll.clientHeight
            - (parseFloat(scrollStyle.paddingTop) + parseFloat(scrollStyle.paddingBottom));
        const size = PAPER_SIZES[this.paperSize];
        const naturalH = (this.orientation === 'landscape' ? size.width : size.height) * MM_TO_PX;
        return availH > 0 && naturalH > 0 ? availH / naturalH : 1;
    }

    /**
     * 将等比缩放系数应用到当前预览页（zoom 整体缩放：宽度、页边距与内容
     * 随同一系数缩放，矢量渲染不失真）。
     * 缩放来源：滑块/输入框的手动覆盖值（10-400%）优先；否则自动按可用高度适配。
     * 只作用于当前页元素，避免为不存在的页面写样式。
     * @param {boolean} [syncControls=true] - 是否同步页脚滑块与输入框回显。
     * 滑块自身拖动产生的 input 事件必须传 false：拖动进行中程序赋值滑块 value
     * 会中断浏览器原生拖拽手势（滑块卡住不跟手），此时仅同步数字输入框即可。
     * 同步时输入框若正聚焦编辑中也跳过，避免覆盖输入。
     * @private
     */
    private _applyPreviewZoom(syncControls: boolean = true): void {
        const zoom = this._previewZoomPercent !== null
            ? Math.min(4, Math.max(0.1, this._previewZoomPercent / 100))
            : this._calcPreviewZoom();
        const div = this._pageDiv;
        if (div) {
            if (zoom !== 1) {
                div.style.setProperty('zoom', String(zoom));
            } else {
                div.style.removeProperty('zoom');
            }
        }
        if (!syncControls) return;
        const percent = String(Math.round(zoom * 100));
        this._zoomSlider?.setValue(Number(percent));
    }

    /**
     * 跳转到指定预览页（越界自动收敛），按需渲染目标页并同步底部导航状态
     * @param {number} page - 目标页索引（0-based）
     * @private
     */
    private _goToPage(page: number): void {
        this.currentPage = this.totalPages > 0 ? Math.max(0, Math.min(page, this.totalPages - 1)) : 0;
        this._renderCurrentPage();
        this._updatePageIndicator();
    }

    /**
     * 更新底部页码指示器（第N页/共M页）与上一页/下一页按钮的可用状态
     * （首页禁用 <，末页禁用 >）
     * @private
     */
    private _updatePageIndicator(): void {
        if (this.pageIndicator) {
            this.pageIndicator.textContent = `第${this.currentPage + 1}页/共${this.totalPages}页`;
        }
        if (this.prevBtn) this.prevBtn.disabled = this.currentPage <= 0;
        if (this.nextBtn) this.nextBtn.disabled = this.currentPage >= this.totalPages - 1;
    }

    /**
     * 核心：按当前纸张设置计算分页计划（只含几何与合并块索引，不生成页面 HTML）。
     *
     * 内容按打印缩放系数缩放后与纸张可打印区域比较：宽度超出时水平分页（列不跨页拆分），
     * 高度超出时垂直分页（行不跨页拆分），总页数 = 列组数 × 行组数，打印顺序「先下后右」；
     * 设置打印区域时仅区域内行列参与分页。
     *
     * 相比历史实现：
     * - 不再调用 matrixValues() 生成 rows×cols 稠密矩阵（避免偏远单元格导致的内存爆炸）；
     * - 合并块直接复用 DataCollection 的合并区倒排索引，全表只扫描一次；
     * - 页面 HTML 延迟到渲染时按页生成。
     * @returns {PrintPlan} 分页计划
     * @private
     */
    private _buildPlan(): PrintPlan {
        const colHeaders = this.data.colHeaders.colHeaders;
        const rowHeaders = this.data.rowHeaders.rowHeaders;
        const size = PAPER_SIZES[this.paperSize];
        const landscape = this.orientation === 'landscape';
        const paperWmm = landscape ? size.height : size.width;
        const paperHmm = landscape ? size.width : size.height;
        const emptyPlan = (blank: boolean): PrintPlan => ({
            paperWmm,
            paperHmm,
            renderScale: 1,
            visCols: [],
            visRows: [],
            globalX: new Map(),
            globalY: new Map(),
            colPages: [],
            rowPages: [],
            blocks: new Map(),
            totalPages: blank ? 1 : 0,
            blank,
        });

        // 1) 内容范围：一次遍历数据即可（不分配 rows×cols 矩阵）
        const bounds = this._contentBounds();
        const rows = bounds.rows;
        const cols = bounds.cols;
        if (rows === 0 || cols === 0) {
            // 完全空白的新文档也生成一张空白页：
            // 预览始终显示纸张白区、打印输出一张空页（对齐 Excel 空表行为）
            return emptyPlan(true);
        }

        // 2) 参与打印的行列：隐藏行/列不参与打印；设置打印区域后仅保留区域内的行列
        let colMin = 0;
        let colMax = cols - 1;
        let rowMin = 0;
        let rowMax = rows - 1;
        const area = this.printAreaRange;
        if (area) {
            colMin = Math.max(0, area.startCol - 1);
            colMax = Math.min(cols - 1, area.endCol - 1);
            rowMin = Math.max(0, area.startRow - 1);
            rowMax = Math.min(rows - 1, area.endRow - 1);
        }
        const visCols: number[] = [];
        for (let c = colMin; c <= colMax; c++) if (!this._isHidden(colHeaders, c)) visCols.push(c);
        const visRows: number[] = [];
        for (let r = rowMin; r <= rowMax; r++) if (!this._isHidden(rowHeaders, r)) visRows.push(r);
        if (visCols.length === 0 || visRows.length === 0) return emptyPlan(false);

        // 3) 纸张几何：可打印区域 = 纸张尺寸 - 上下左右页边距（当前预设）
        const paperW = paperWmm * MM_TO_PX;
        const paperH = paperHmm * MM_TO_PX;
        const m = this.margins;
        const printableW = paperW - (m.left + m.right) * MM_TO_PX;
        const printableH = paperH - (m.top + m.bottom) * MM_TO_PX;

        // 4) 缩放系数：按缩放模式解析（fit 系列只缩小不放大），布局与字体统一按其缩放
        let totalW = 0;
        for (const c of visCols) totalW += colHeaders[c]?.width ?? 100;
        let totalH = 0;
        for (const r of visRows) totalH += rowHeaders[r]?.height ?? 24;
        const scale = this._resolveScale(totalW, totalH, printableW, printableH);

        // 5) 可见行列的全局布局（缩放后尺寸），供分页与合并块跨页定位使用
        const globalX = new Map<number, { x: number; w: number }>();
        let gx = 0;
        for (const c of visCols) {
            const w = (colHeaders[c]?.width ?? 100) * scale;
            globalX.set(c, { x: gx, w });
            gx += w;
        }
        const globalY = new Map<number, { y: number; h: number }>();
        let gy = 0;
        for (const r of visRows) {
            const h = (rowHeaders[r]?.height ?? 24) * scale;
            globalY.set(r, { y: gy, h });
            gy += h;
        }

        // 6) 水平分页：列宽累加超出可打印宽度即换页（列不跨页拆分）
        const colPages: ColPage[] = [];
        let curCols: number[] = [];
        let curColsW = 0;
        for (const c of visCols) {
            const { w } = globalX.get(c)!;
            if (curCols.length > 0 && curColsW + w > printableW + 0.5) {
                colPages.push({ cols: curCols });
                curCols = [];
                curColsW = 0;
            }
            curCols.push(c);
            curColsW += w;
        }
        if (curCols.length > 0) colPages.push({ cols: curCols });

        // 7) 垂直分页：行高累加超出可打印高度即换页（行不跨页拆分）
        const rowPages: PageGeometry[] = [];
        let curRows: number[] = [];
        let curHeights: number[] = [];
        let curH = 0;
        for (const r of visRows) {
            const { h } = globalY.get(r)!;
            if (curRows.length > 0 && curH + h > printableH + 0.5) {
                rowPages.push({ rows: curRows, heights: curHeights });
                curRows = [];
                curHeights = [];
                curH = 0;
            }
            curRows.push(r);
            curHeights.push(h);
            curH += h;
        }
        if (curRows.length > 0) rowPages.push({ rows: curRows, heights: curHeights });

        // 8) 合并块索引：复用 DataCollection 已建好的合并区列表，全表只扫描一次，
        //    记录块内可见首末行列（跨页截断判断用）。渲染每页时再按页窗口取相交子集。
        const blocks = new Map<string, PrintBlock>();
        for (const rect of this.data.getMergeRects()) {
            const r0 = rect.startRow - 1;
            const c0 = rect.startCol - 1;
            const rs = rect.endRow - rect.startRow + 1;
            const cs = rect.endCol - rect.startCol + 1;
            let firstRow = -1;
            let lastRow = -1;
            let firstCol = -1;
            let lastCol = -1;
            for (let dr = 0; dr < rs; dr++) {
                if (globalY.has(r0 + dr)) {
                    if (firstRow < 0) firstRow = r0 + dr;
                    lastRow = r0 + dr;
                }
            }
            for (let dc = 0; dc < cs; dc++) {
                if (globalX.has(c0 + dc)) {
                    if (firstCol < 0) firstCol = c0 + dc;
                    lastCol = c0 + dc;
                }
            }
            if (firstRow < 0 || firstCol < 0) continue;  // 块内无可见行/列
            blocks.set(rect.anchor, { anchor: rect.anchor, r0, c0, rs, cs, firstRow, lastRow, firstCol, lastCol });
        }

        return {
            paperWmm,
            paperHmm,
            renderScale: scale,
            visCols,
            visRows,
            globalX,
            globalY,
            colPages,
            rowPages,
            blocks,
            totalPages: colPages.length * rowPages.length,
            blank: false,
        };
    }

    /**
     * 取数据中「最远单元格」的行列号（1-based，作为行列总数使用）。
     * 用一次遍历替代 matrixValues() 的稠密 rows×cols 分配。
     * 合并区右下角即使没有数据也计入范围（否则打印时合并块会被截断成单格）。
     * @returns {{rows: number, cols: number}} 行数/列数
     * @private
     */
    private _contentBounds(): { rows: number; cols: number } {
        let rows = 0;
        let cols = 0;
        for (const cell of this.data.values) {
            const pos = this.data.getCellColAndRow(cell.cell);
            if (pos.row > rows) rows = pos.row;
            if (pos.col > cols) cols = pos.col;
        }
        // 合并单元格：内容只存在锚点格子里，合并区右下角的格子可能**根本没有数据**。
        // 若只按「有值的格子」求边界，合并区的跨度会被截断在锚点所在的行列，
        // 打印时该块只剩 1 行 1 列、文字被裁掉（下方或右侧另有内容时才会被撑开）。
        // 因此把每个合并区的右下角也计入内容范围。
        for (const rect of this.data.getMergeRects()) {
            if (rect.endRow > rows) rows = rect.endRow;
            if (rect.endCol > cols) cols = rect.endCol;
        }
        return { rows, cols };
    }

    /**
     * 渲染指定页的表格 HTML（按需调用：预览只渲染当前页，打印按时间片逐页生成）。
     * 页面顺序「先下后右」：同一列组内的行页从上到下，再进入下一列组。
     * @param {PrintPlan} plan - 分页计划
     * @param {number} pageIndex - 页索引（0-based）
     * @returns {string} 本页表格 HTML
     * @private
     */
    private _renderPageHtml(plan: PrintPlan, pageIndex: number): string {
        if (plan.blank || plan.totalPages <= 0) return '';
        const rowPageCount = plan.rowPages.length;
        const colIdx = Math.floor(pageIndex / rowPageCount);
        const rowIdx = pageIndex % rowPageCount;
        if (colIdx < 0 || colIdx >= plan.colPages.length || rowIdx < 0 || rowIdx >= rowPageCount) return '';
        this.renderScale = plan.renderScale;
        return this._buildPageTable(plan, plan.rowPages[rowIdx], plan.colPages[colIdx]);
    }

    /**
     * 生成本页表格 HTML：table-layout: fixed + colgroup 显式列宽 + 固定行高，
     * 文字与边框交由浏览器渲染/打印引擎矢量输出（精度优于 canvas 位图光栅化）。
     *
     * 合并块只取「与本页窗口相交」的子集（DataCollection 的列倒排索引 O(列数 + 块数)），
     * 而非像历史实现那样每页遍历全表所有合并块；页内发出位置与跳过位置用一维数组记录，
     * 避免为每个格子创建 `"r:c"` 字符串键。
     * @param {PrintPlan} plan - 分页计划
     * @param {PageGeometry} rowPage - 本页行几何
     * @param {ColPage} colPage - 本页列组
     * @returns {string} 本页表格 HTML
     * @private
     */
    private _buildPageTable(plan: PrintPlan, rowPage: PageGeometry, colPage: ColPage): string {
        /** 合并块在本页的发出规格：收缩后的跨度、本页可见高度与四边边框显隐 */
        interface EmitSpec { cell: Cell; colspan: number; rowspan: number; height: number; showTop: boolean; showBottom: boolean; showLeft: boolean; showRight: boolean; }

        const pageRows = rowPage.rows;
        const pageCols = colPage.cols;
        const rowCount = pageRows.length;
        const colCount = pageCols.length;

        // 页内位置：一维下标 = 行位置 * 列数 + 列位置
        const emits: Array<EmitSpec | undefined> = new Array(rowCount * colCount);
        const skip = new Uint8Array(rowCount * colCount);

        if (plan.blocks.size > 0 && rowCount > 0 && colCount > 0) {
            // 1-based 页窗口，取与本页相交的合并块（只返回相关子集）
            const rects = this.data.getVisibleMergeRects(
                pageCols[0] + 1,
                pageCols[colCount - 1] + 1,
                pageRows[0] + 1,
                pageRows[rowCount - 1] + 1
            );
            for (const rect of rects) {
                const b = plan.blocks.get(rect.anchor);
                if (!b) continue;
                // 块与本页的可见交集（页内局部下标）：页行列数很小，
                // 直接按页窗口过滤，避免为跨越多页的高合并块重复遍历整段跨度
                const visColsLocal: number[] = [];
                for (let ci = 0; ci < colCount; ci++) {
                    const c = pageCols[ci];
                    if (c >= b.firstCol && c <= b.lastCol) visColsLocal.push(ci);
                }
                if (visColsLocal.length === 0) continue;
                const visRowsLocal: number[] = [];
                for (let ri = 0; ri < rowCount; ri++) {
                    const r = pageRows[ri];
                    if (r >= b.firstRow && r <= b.lastRow) visRowsLocal.push(ri);
                }
                if (visRowsLocal.length === 0) continue;
                let height = 0;
                for (const ri of visRowsLocal) height += rowPage.heights[ri];
                const anchorCell = this.data.getCellByName(b.anchor);
                if (!anchorCell) continue;  // 数据变化后锚点已不存在
                const emitKey = visRowsLocal[0] * colCount + visColsLocal[0];
                emits[emitKey] = {
                    cell: anchorCell,
                    colspan: visColsLocal.length,
                    rowspan: visRowsLocal.length,
                    height,
                    showTop: pageRows[visRowsLocal[0]] === b.firstRow,
                    showBottom: pageRows[visRowsLocal[visRowsLocal.length - 1]] === b.lastRow,
                    showLeft: pageCols[visColsLocal[0]] === b.firstCol,
                    showRight: pageCols[visColsLocal[visColsLocal.length - 1]] === b.lastCol,
                };
                for (const ri of visRowsLocal) {
                    const base = ri * colCount;
                    for (const ci of visColsLocal) skip[base + ci] = 1;
                }
            }
        }

        // 表格宽度 = 本页列宽之和：列组可能不满页宽，不能被固定布局拉伸补齐
        const colWidths: number[] = new Array(colCount);
        let tableW = 0;
        for (let ci = 0; ci < colCount; ci++) {
            const w = plan.globalX.get(pageCols[ci])!.w;
            colWidths[ci] = w;
            tableW += w;
        }

        const parts: string[] = [`<table class="gs-print-table" style="width:${tableW}px"><colgroup>`];
        for (let ci = 0; ci < colCount; ci++) parts.push(`<col style="width:${colWidths[ci]}px">`);
        parts.push('</colgroup><tbody>');
        for (let ri = 0; ri < rowCount; ri++) {
            const r = pageRows[ri];
            const rowHeight = rowPage.heights[ri];
            parts.push(`<tr style="height:${rowHeight}px">`);
            const base = ri * colCount;
            for (let ci = 0; ci < colCount; ci++) {
                const emit = emits[base + ci];
                if (emit) {
                    parts.push(this._buildCellTd(emit.cell, emit.colspan, emit.rowspan, emit.height,
                        emit.showTop, emit.showBottom, emit.showLeft, emit.showRight));
                    continue;
                }
                if (skip[base + ci]) continue;
                const cell = this.data.getCellByName(this.data.getCellName(pageCols[ci] + 1, r + 1));
                parts.push(this._buildCellTd(cell, 1, 1, rowHeight, true, true, true, true));
            }
            parts.push('</tr>');
        }
        parts.push('</tbody></table>');
        return parts.join('');
    }

    /**
     * 生成单个 td 的 HTML（含合并跨度、边框显隐与单元格样式）。
     * 内容包一层固定高度、flex 垂直对齐、溢出裁剪的容器：
     * 行高不被内容撑破（分页几何保持精确），文字超出时裁剪（对应画布方案的裁剪行为）。
     * @param {Cell | undefined} cell - 单元格数据（可为 undefined，输出空白格）
     * @param {number} colspan - 本页可见的水平合并跨度
     * @param {number} rowspan - 本页可见的垂直合并跨度
     * @param {number} height - 单元格总高度（px，含上下 padding 与上下边框）
     * @param {boolean} showTop - 是否绘制上边框（块体被页边界截断的边不画）
     * @param {boolean} showBottom - 是否绘制下边框
     * @param {boolean} showLeft - 是否绘制左边框
     * @param {boolean} showRight - 是否绘制右边框
     * @returns {string} td HTML
     * @private
     */
    private _buildCellTd(
        cell: Cell | undefined,
        colspan: number,
        rowspan: number,
        height: number,
        showTop: boolean,
        showBottom: boolean,
        showLeft: boolean,
        showRight: boolean
    ): string {
        const attrs: string[] = [];
        if (colspan > 1) attrs.push(`colspan="${colspan}"`);
        if (rowspan > 1) attrs.push(`rowspan="${rowspan}"`);
        const attrText = attrs.length > 0 ? ` ${attrs.join(' ')}` : '';
        if (!cell) {
            return `<td${attrText} style="height:${height}px"></td>`;
        }
        const style: string[] = [`height:${height}px`];
        if (cell.backgroundColor) style.push(`background-color:${this._escapeHtml(cell.backgroundColor)}`);
        // 边框：严格按数据实际值输出——仅显式设置且大于 0 的边显示
        const bw = (w?: number): number => (w !== undefined && w > 0 ? w : 0);
        const borderTop = showTop ? bw(cell.borderTopWidth) : 0;
        const borderBottom = showBottom ? bw(cell.borderBottomWidth) : 0;
        if (borderTop > 0) style.push(`border-top:${borderTop}px solid #000`);
        if (borderBottom > 0) style.push(`border-bottom:${borderBottom}px solid #000`);
        if (showLeft && bw(cell.borderLeftWidth) > 0) style.push(`border-left:${bw(cell.borderLeftWidth)}px solid #000`);
        if (showRight && bw(cell.borderRightWidth) > 0) style.push(`border-right:${bw(cell.borderRightWidth)}px solid #000`);
        // 字体与文字样式（字号/行距/字距按打印缩放系数缩放，对应整体缩放打印）
        const fontSize = (cell.fontSize ?? 14) * this.renderScale;
        style.push(`font-size:${fontSize}px`);
        style.push(`line-height:${fontSize * 1.25 + (cell.lineSpacing ?? 0) * this.renderScale}px`);
        style.push(`font-family:${this._escapeHtml(cell.fontFamily ?? DEFAULT_FONT_FAMILY)}`);
        if (cell.fontWeight) style.push('font-weight:bold');
        if (cell.fontStyle) style.push('font-style:italic');
        style.push(`color:${cell.fontColor ? this._escapeHtml(cell.fontColor) : '#000'}`);
        if (cell.letterSpacing) style.push(`letter-spacing:${cell.letterSpacing * this.renderScale}px`);
        // 下划线 / 删除线：直接标注在内层 span 上（flex 项会阻断从 td 的装饰传播）
        const decorations: string[] = [];
        if (cell.underline) decorations.push('underline');
        if (cell.strikethrough) decorations.push('line-through');
        const decorationStyle = decorations.length > 0 ? ` style="text-decoration:${decorations.join(' ')}"` : '';
        // 换行：wrap 时保留换行符并按字符折行，否则去除换行符按单行显示
        style.push(cell.wrap ? 'white-space:pre-wrap;word-break:break-all' : 'white-space:pre');
        // 水平对齐：center（默认）/ left（含 justify）/ right（随继承链作用于内层文本）
        const hAlign = cell.textAlign ?? 'center';
        style.push(`text-align:${hAlign === 'right' ? 'right' : hAlign === 'left' || hAlign === 'justify' ? 'left' : 'center'}`);
        const raw = cell.chars?.map(ch => ch.char).join('') ?? '';
        const text = this._escapeHtml(cell.wrap ? raw : raw.replace(/\n/g, ''));
        // 内层固定高度容器：垂直对齐（middle 默认）+ 裁剪，高度扣除 padding 与已绘制的上下边框
        const PAD = 2;
        const innerH = Math.max(0, height - 2 * PAD - borderTop - borderBottom);
        const vAlign = cell.alignItems ?? 'middle';
        const justify = vAlign === 'top' ? 'flex-start' : vAlign === 'bottom' ? 'flex-end' : 'center';
        const content = `<div style="height:${innerH}px;overflow:hidden;display:flex;flex-direction:column;justify-content:${justify}"><span${decorationStyle}>${text}</span></div>`;
        return `<td${attrText} style="${style.join(';')}">${content}</td>`;
    }

    /**
     * 表格基础样式（预览与打印文档共用）：固定布局 + 合并边框模型 + 单元格内边距。
     * border-collapse: collapse 与主页面 canvas 的描边语义一致——相邻单元格共享边缘只画一条线
     * （宽度取大者），避免 separate 模型下相邻两边框并排叠加、内部线条比外围粗一倍的问题。
     * @returns {string} CSS 规则文本
     * @private
     */
    private _tableStyles(): string {
        return [
            '.gs-print-table { border-collapse: collapse; table-layout: fixed; }',
            '.gs-print-table td { box-sizing: border-box; padding: 2px; text-align: center; vertical-align: middle; overflow: hidden; }'
        ].join('\n');
    }

    /**
     * HTML 转义：& < > "（单元格内容、颜色值、字体名等拼接进 HTML/属性前调用）
     * @param {string} text - 原始文本
     * @returns {string} 转义后的文本
     * @private
     */
    private _escapeHtml(text: string): string {
        return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    /**
     * 解析页面范围设置（data.printSetting.pageRange，如 '1-3,5'，1-based 页序，
     * 兼容中英文逗号/分号/空白分隔），返回命中的页序号集合。
     * 未设置/全部无法解析时返回 null（表示打印全部页面），越界页序自动忽略
     * @param {number} total - 总页数
     * @returns {Set<number> | null} 命中的页序号集合；null=打印全部
     * @private
     */
    private _parsePageRange(total: number): Set<number> | null {
        const raw = (this.data.printSetting.pageRange ?? '').trim();
        if (!raw) return null;
        const hits = new Set<number>();
        for (const token of raw.split(/[,，;；\s]+/)) {
            const m = token.match(/^(\d+)(?:\s*-\s*(\d+))?$/);
            if (!m) continue;
            const a = Number(m[1]);
            const b = m[2] !== undefined ? Number(m[2]) : a;
            const lo = Math.max(1, Math.min(a, b));
            const hi = Math.min(total, Math.max(a, b));
            for (let p = lo; p <= hi; p++) hits.add(p);
        }
        return hits.size > 0 ? hits : null;
    }

    /**
     * 填充打印机下拉选项并回显用户上次选择
     * @param {HTMLSelectElement} select - 打印机下拉框
     * @private
     */
    private _fillPrinterOptions(select: HTMLSelectElement): void {
        const names = this._getPrinterNames();
        select.innerHTML = '';
        for (const name of names) {
            const opt = document.createElement('option');
            opt.value = name;
            opt.textContent = name;
            select.appendChild(opt);
        }
        if (this.printerName && names.includes(this.printerName)) {
            select.value = this.printerName;
        } else {
            this.printerName = names[0] ?? '';
        }
    }

    /**
     * 获取打印机名称列表。
     * 浏览器沙箱不提供枚举系统打印机的 API，故固定提供「系统默认打印机」，
     * 实际目标打印机在打印对话框中选择；接入桌面壳（如 Electron）后可在此扩展真实列表。
     * @returns {string[]} 打印机名称列表
     * @private
     */
    private _getPrinterNames(): string[] {
        return ['系统默认打印机'];
    }

    /**
     * 生成打印文档骨架（head + 空 body）。
     *
     * 页面不再一次性拼进文档：print() 按时间片逐页生成并追加到 body，
     * 大表打印时主线程不再被一次性字符串生成与 HTML 解析占满。
     *
     * 打印设置同步到浏览器打印对话框的方式：
     * - 纸张大小与方向：@page size 直接采用当前设置，浏览器打印预览随之变化；
     * - 页边距：@page margin 直接采用当前设置（mm），打印输出与预览的版面一致；
     *   页面内容不再用 padding 模拟页边距，避免与实际页边距叠加；
     * - 页数：每个 .gs-print-sheet 恰好对应一张纸（内容高度已按可打印区域分页，
     *   不设固定页高，避免 mm→px 取整溢出多出空白页），对话框页数 = 预览页数；
     * - 页眉页脚：由浏览器对话框的「页眉和页脚」选项控制，Web 端无法预设；
     * - 打印机 / 份数 / 打印方式（双面）：Web API 无法预设，需在对话框中选择。
     * @returns {string} 可直接写入 iframe 的 HTML 文档骨架
     * @private
     */
    private _buildPrintHead(): string {
        const size = PAPER_SIZES[this.paperSize];
        const m = this.margins;
        const paperHmm = this.orientation === 'portrait' ? size.height : size.width;
        // 可打印高度（mm），再留 0.5mm 余量：分页容差为 0.5px，理论上单页内容可能
        // 恰好贴边，留出余量后由 max-height 兜底，杜绝溢出再拆出一张空白页
        const sheetMaxHmm = Math.max(1, paperHmm - m.top - m.bottom - 0.5);
        return `<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>${this.title}</title>
    <style>
        /* 纸张与页边距直接采用当前打印设置，浏览器打印预览随之一致 */
        @page { size: ${size.page} ${this.orientation}; margin: ${this._marginCss()}; }
        html, body { margin: 0; padding: 0; }
        /* 逐页容器：内容已按「纸张 - 页边距」分页且行高固定，不会超出一页；
           max-height 兜底并裁剪，保证「一个容器 = 一张纸」，
           不设固定页高，避免 mm→px 取整溢出产生多余空白页 */
        .gs-print-sheet { box-sizing: border-box; max-height: ${sheetMaxHmm}mm; overflow: hidden; page-break-after: always; break-after: page; break-inside: avoid; }
        .gs-print-sheet:last-child { page-break-after: auto; break-after: auto; }
        /* 打印时强制保留背景色与文字颜色（浏览器默认不打印背景） */
        .gs-print-sheet .gs-print-table { print-color-adjust: exact; -webkit-print-color-adjust: exact; }
        ${this._tableStyles()}
        /* 自定义字体：iframe 为独立文档，主文档字体不共享，需重新注入 @font-face */
        ${this._getFontFaceCss()}
    </style>
</head>
<body></body>
</html>`;
    }

    /**
     * 判断指定索引的行/列是否隐藏
     * @param {Array<{isHidden?: boolean}>} headers - 行/列头部数据
     * @param {number} index - 索引（0-based）；超出头部范围视为可见
     * @returns {boolean} 是否隐藏
     * @private
     */
    private _isHidden(headers: Array<{ isHidden?: boolean }>, index: number): boolean {
        return headers[index]?.isHidden === true;
    }
}

/** 共享打印实例：工具栏打印设置分区与快速访问栏共用，保证打印设置互通 */
let sharedPrinter: Printer | null = null;

/**
 * 获取共享 Printer 实例（首次调用时创建，此后返回同一实例）。
 * 「页面」工具栏分区的打印设置按钮与快速访问栏的打印/打印预览按钮
 * 通过本方法取得同一实例，打印设置变更互通。
 * @param {DataCollection} data - 数据集合
 * @returns {Printer} 共享实例
 * @example
 * const printer = getSharedPrinter(data);
 * printer.setOrientation('landscape');
 */
export function getSharedPrinter(data: DataCollection): Printer {
    if (!sharedPrinter) sharedPrinter = new Printer({ data });
    return sharedPrinter;
}
