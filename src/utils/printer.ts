/*  # 打印工具模块 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import type { Cell } from "../core/dataArchitecture/cell";
import type { DataCollection } from "../core/dataArchitecture/dataCollection";
import { getRegisteredFontFaceCss } from "./fontsLoader";
// 默认字体与主页面 canvas 保持一致（单一数据源），保证预览/打印与主页面的字体观感相同
import { DEFAULT_FONT_FAMILY } from "../core/constant";

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
    /** 预览各页容器元素（与页序一一对应，非当前页隐藏） */
    private pageDivs: HTMLElement[] = [];
    /** 底部页码指示器元素 */
    private pageIndicator: HTMLElement | null = null;
    /** 上一页 / 下一页按钮 */
    private prevBtn: HTMLButtonElement | null = null;
    private nextBtn: HTMLButtonElement | null = null;
    /** 目标打印机名称（浏览器沙箱无法枚举/指定系统打印机，最终在打印对话框中确认） */
    private printerName = '';
    /** 打印份数（默认 1） */
    private copies = 1;
    /** 打印顺序：page=逐页打印（1,1,2,2,3,3），copies=逐份打印（1,2,3,1,2,3） */
    private collate: 'page' | 'copies' = 'page';
    /** 追加注册的自定义字体 @font-face CSS（来自 options.fontFaceCss 与 registerFontCss 调用） */
    private _fontFaceCss: string[] = [];
    /** 本次分页使用的缩放系数（_buildPageHtmls 解析后供 td 字号/行距/字距缩放读取） */
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
    private _zoomSlider: HTMLInputElement | null = null;
    private _zoomInput: HTMLInputElement | null = null;
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
        return [...new Set([...collected, ...this._fontFaceCss, ...getRegisteredFontFaceCss()])].join('\n');
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
     * 预览打开时同步设置面板控件回显并重绘预览页；未打开时不做任何事
     * @private
     */
    private _refreshPreview(): void {
        if (!this.overlay) return;
        this._syncPanelControls();
        this._renderPagesInto(this._pagesContainer);
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
        overlay.innerHTML = `
            <style>
                /* 样式遵循 Bootstrap 5 设计规范：字体堆栈、颜色变量、圆角、阴影、表单焦点环、按钮体系 */
                .gs-print-overlay { position: fixed; inset: 0; z-index: 9999; background: rgba(0, 0, 0, 0.5); display: flex; flex-direction: column; font-family: system-ui, -apple-system, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, 'Noto Sans', 'Liberation Sans', sans-serif; font-size: 1rem; color: #212529; }
                .gs-print-toolbar { flex: none; display: flex; align-items: center; padding: 0.75rem 1rem; background: #f8f9fa; border-bottom: 1px solid #dee2e6; }
                .gs-print-title { font-size: 1rem; font-weight: bold; color: #212529; }
                .gs-print-body { flex: 1; display: flex; overflow: hidden; min-height: 0; }
                /* 主体列：预览滚动区 + 页脚（页脚仅覆盖预览区宽度）；右侧设置面板独占全高直达页面底部 */
                .gs-print-main { flex: 1; display: flex; flex-direction: column; min-height: 0; }
                .gs-print-scroll { flex: 1; overflow: auto; padding: 1.5rem; background: #6c757d; }
                /* 预览页面：与实际纸张等比例（宽高与页边距 padding 由 _renderPagesInto 按当前设置内联指定），
                   整页以 zoom 等比缩放——高度填满滚动区可用高度，宽度与内容随同一系数缩放（见 _applyPreviewZoom）；
                   单页显示模式，内容不足一页时白色背景仍是完整纸张 */
                .gs-print-page { box-sizing: border-box; margin: 0 auto; background: #fff; overflow: hidden; border-radius: 0.375rem; box-shadow: 0 0.5rem 1rem rgba(0, 0, 0, 0.15); }
                .gs-print-panel { flex: none; width: 264px; border-left: 1px solid #dee2e6; background: #f8f9fa; padding: 1rem; display: flex; flex-direction: column; gap: 0.75rem; overflow: auto; }
                .gs-print-panel-title { font-size: 1rem; font-weight: bold; color: #212529; }
                /* 设置行：文字提示与操作控件同一水平位置（label 固定宽左对齐，控件右侧伸展） */
                .gs-print-field { display: flex; align-items: center; gap: 0.5rem; font-size: 0.875rem; color: #212529; }
                .gs-print-label { flex: none; width: 4.25em; }
                .gs-print-field > .gs-print-select,
                .gs-print-field > .gs-print-input,
                .gs-print-field > .gs-print-stepper { flex: 1; min-width: 0; }
                .gs-print-field > .gs-print-input { width: auto; }
                /* 分割线：设置分组之间 */
                .gs-print-divider { flex: none; height: 1px; background: #dee2e6; }
                /* 表单控件：form-select / form-control 风格（含 focus 蓝色焦点环） */
                .gs-print-select { padding: 0.375rem 0.75rem; font-size: 0.875rem; border: 1px solid #ced4da; border-radius: 0.375rem; background: #fff; color: #212529; cursor: pointer; outline: none; transition: border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out; }
                .gs-print-select:focus { outline: none; border-color: #86b7fe; box-shadow: 0 0 0 0.25rem rgba(13, 110, 253, 0.25); }
                .gs-print-input { padding: 0.375rem 0.5rem; font-size: 0.875rem; border: 1px solid #ced4da; border-radius: 0.375rem; background: #fff; color: #212529; width: 100%; box-sizing: border-box; outline: none; transition: border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out; }
                .gs-print-input:focus { outline: none; border-color: #86b7fe; box-shadow: 0 0 0 0.25rem rgba(13, 110, 253, 0.25); }
                /* 数字输入框统一隐藏原生步进箭头（步进由自定义按钮/滑块承担） */
                .gs-print-input::-webkit-outer-spin-button,
                .gs-print-input::-webkit-inner-spin-button { -webkit-appearance: none; margin: 0; }
                .gs-print-input[type="number"] { -moz-appearance: textfield; appearance: textfield; }
                .gs-print-stepper { display: flex; align-items: stretch; }
                .gs-print-stepper .gs-print-input { flex: 1; min-width: 0; border-radius: 0.375rem 0 0 0.375rem; -moz-appearance: textfield; appearance: textfield; }
                .gs-print-stepper .gs-print-input::-webkit-outer-spin-button,
                .gs-print-stepper .gs-print-input::-webkit-inner-spin-button { -webkit-appearance: none; margin: 0; }
                .gs-print-stepper .gs-print-input:focus { position: relative; z-index: 1; }
                .gs-print-stepper-btns { flex: none; display: flex; flex-direction: column; border: 1px solid #ced4da; border-left: none; border-radius: 0 0.375rem 0.375rem 0; overflow: hidden; background: #fff; }
                .gs-print-step-btn { flex: 1; width: 22px; padding: 0; font-size: 0.75rem; line-height: 1; border: none; background: #fff; color: #212529; cursor: pointer; transition: background 0.15s ease-in-out; }
                .gs-print-step-btn:hover { background: #e9ecef; }
                .gs-print-step-btn:active { background: #dee2e6; }
                /* 小按钮：行内操作（调整页边距/应用/取消） */
                .gs-print-btn-sm { flex: none; padding: 0.25rem 0.5rem; font-size: 0.8125rem; line-height: 1.25; border: 1px solid #dee2e6; border-radius: 0.375rem; background: #fff; color: #212529; cursor: pointer; white-space: nowrap; transition: color 0.15s ease-in-out, background 0.15s ease-in-out, border-color 0.15s ease-in-out; }
                .gs-print-btn-sm:hover { background: #f8f9fa; }
                .gs-print-btn-sm.gs-print-btn-primary { background: #0d6efd; border-color: #0d6efd; color: #fff; }
                .gs-print-btn-sm.gs-print-btn-primary:hover { background: #0b5ed7; border-color: #0a58ca; color: #fff; }
                /* 自定义页边距编辑器：展开于「页边距」行下方，上下左右四向 mm 输入 */
                .gs-print-margin-editor { display: flex; flex-direction: column; gap: 0.5rem; padding: 0.625rem; background: #fff; border: 1px solid #dee2e6; border-radius: 0.375rem; }
                .gs-print-margin-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 0.5rem; }
                .gs-print-mini-field { display: flex; align-items: center; gap: 0.375rem; font-size: 0.8125rem; color: #212529; }
                .gs-print-mini-field > .gs-print-input { flex: 1; min-width: 0; width: auto; padding: 0.25rem 0.375rem; font-size: 0.8125rem; }
                .gs-print-mini-actions { display: flex; gap: 0.5rem; }
                .gs-print-mini-actions .gs-print-btn-sm { flex: 1; white-space: normal; }
                /* 按钮行：吸附面板底部，但与面板底边留出大于页脚条高度（≈51px）的间距 */
                .gs-print-actions { display: flex; gap: 0.5rem; margin-top: auto; margin-bottom: 3.5rem; }
                /* 按钮：默认 .btn 风格（白底 + #dee2e6 边 + hover #f8f9fa） */
                .gs-print-btn { flex: 1; padding: 0.375rem 0; font-size: 0.875rem; border: 1px solid #dee2e6; border-radius: 0.375rem; background: #fff; color: #212529; cursor: pointer; transition: color 0.15s ease-in-out, background 0.15s ease-in-out, border-color 0.15s ease-in-out; }
                .gs-print-btn:hover { background: #f8f9fa; }
                /* 打印按钮 = .btn-primary（含 hover/active 层级色） */
                .gs-print-btn[data-gs-print-action="print"] { background: #0d6efd; border-color: #0d6efd; color: #fff; }
                .gs-print-btn[data-gs-print-action="print"]:hover { background: #0b5ed7; border-color: #0a58ca; color: #fff; }
                .gs-print-btn[data-gs-print-action="print"]:active { background: #0a58ca; border-color: #0a58ca; color: #fff; }
                .gs-print-footer { flex: none; display: flex; align-items: center; justify-content: center; gap: 0.75rem; padding: 0.75rem 1rem; background: #f8f9fa; border-top: 1px solid #dee2e6; position: relative; }
                .gs-print-indicator { font-size: 0.875rem; color: #212529; min-width: 110px; text-align: center; }
                /* 预览缩放控件：固定于页脚右侧（页码居中不受影响），滑块 + 百分比输入框 */
                .gs-print-footer-zoom { position: absolute; right: 1rem; top: 50%; transform: translateY(-50%); display: flex; align-items: center; gap: 0.5rem; }
                .gs-print-zoom-slider { width: 120px; accent-color: #0d6efd; cursor: pointer; touch-action: none; }
                .gs-print-zoom-input { width: 3.5rem; padding: 0.25rem 0.375rem; font-size: 0.8125rem; border: 1px solid #ced4da; border-radius: 0.375rem; background: #fff; color: #212529; text-align: right; outline: none; box-sizing: border-box; transition: border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out; }
                .gs-print-zoom-input:focus { outline: none; border-color: #86b7fe; box-shadow: 0 0 0 0.25rem rgba(13, 110, 253, 0.25); }
                .gs-print-zoom-input::-webkit-outer-spin-button,
                .gs-print-zoom-input::-webkit-inner-spin-button { -webkit-appearance: none; margin: 0; }
                .gs-print-zoom-input[type="number"] { -moz-appearance: textfield; appearance: textfield; }
                .gs-print-zoom-unit { font-size: 0.8125rem; color: #6c757d; }
                /* 翻页按钮 = .btn-outline-secondary 风格 */
                .gs-print-nav-btn { padding: 0.25rem 0.75rem; font-size: 0.875rem; border: 1px solid #6c757d; border-radius: 0.375rem; background: #fff; color: #6c757d; cursor: pointer; transition: color 0.15s ease-in-out, background 0.15s ease-in-out, border-color 0.15s ease-in-out; }
                .gs-print-nav-btn:hover:not(:disabled) { background: #6c757d; border-color: #6c757d; color: #fff; }
                .gs-print-nav-btn:disabled { color: #adb5bd; border-color: #dee2e6; background: #fff; cursor: not-allowed; }
                /* 预览表格基础样式（与打印文档共用） */
                ${this._tableStyles()}
                /* 自定义字体：注入收集与注册的 @font-face 规则，保证预览与打印字体一致 */
                ${this._getFontFaceCss()}
            </style>
            <div class="gs-print-toolbar">
                <span class="gs-print-title">打印预览</span>
            </div>
            <div class="gs-print-body">
                <div class="gs-print-main">
                    <div class="gs-print-scroll">
                        <div class="gs-print-pages"></div>
                    </div>
                    <div class="gs-print-footer">
                        <button class="gs-print-nav-btn" data-gs-print-nav="prev">&lt;</button>
                        <span class="gs-print-indicator"></span>
                        <button class="gs-print-nav-btn" data-gs-print-nav="next">&gt;</button>
                        <div class="gs-print-footer-zoom">
                            <input type="range" class="gs-print-zoom-slider" min="10" max="400" step="1" data-gs-print-role="zoom-slider">
                            <input type="number" class="gs-print-zoom-input" min="10" max="400" step="1" data-gs-print-role="zoom-input">
                            <span class="gs-print-zoom-unit">%</span>
                        </div>
                    </div>
                </div>
                <div class="gs-print-panel">
                    <div class="gs-print-panel-title">打印设置</div>
                    <label class="gs-print-field"><span class="gs-print-label">打印机</span>
                        <select class="gs-print-select" data-gs-print-role="printer"></select>
                    </label>
                    <label class="gs-print-field"><span class="gs-print-label">份数</span>
                        <span class="gs-print-stepper">
                            <input class="gs-print-input" type="number" min="1" max="999" step="1" value="1" data-gs-print-role="copies">
                            <span class="gs-print-stepper-btns">
                                <button type="button" class="gs-print-step-btn" data-gs-print-step="inc">+</button>
                                <button type="button" class="gs-print-step-btn" data-gs-print-step="dec">&minus;</button>
                            </span>
                        </span>
                    </label>
                    <label class="gs-print-field"><span class="gs-print-label">顺序</span>
                        <select class="gs-print-select" data-gs-print-role="collate">
                            <option value="page">逐页打印</option>
                            <option value="copies">逐份打印</option>
                        </select>
                    </label>
                    <label class="gs-print-field"><span class="gs-print-label">打印方式</span>
                        <select class="gs-print-select" data-gs-print-role="duplex">
                            <option value="simplex">单面打印</option>
                            <option value="duplex-long">双面打印-长边翻页</option>
                            <option value="duplex-short">双面打印-短边翻页</option>
                        </select>
                    </label>
                    <div class="gs-print-divider"></div>
                    <label class="gs-print-field"><span class="gs-print-label">纸张大小</span>
                        <select class="gs-print-select" data-gs-print-role="paper-size">
                            <option value="A4">A4</option>
                            <option value="A3">A3</option>
                            <option value="B5">B5</option>
                            <option value="Letter">Letter</option>
                        </select>
                    </label>
                    <label class="gs-print-field"><span class="gs-print-label">纸张方向</span>
                        <select class="gs-print-select" data-gs-print-role="orientation">
                            <option value="portrait">纵向</option>
                            <option value="landscape">横向</option>
                        </select>
                    </label>
                    <div class="gs-print-field"><span class="gs-print-label">页边距</span>
                        <select class="gs-print-select" data-gs-print-role="margin">
                            <option value="normal">常规</option>
                            <option value="wide">宽</option>
                            <option value="narrow">窄</option>
                            <option value="custom">自定义</option>
                        </select>
                        <button type="button" class="gs-print-btn-sm" data-gs-print-action="margin-adjust">调整页边距</button>
                    </div>
                    <div class="gs-print-margin-editor" data-gs-print-role="margin-editor" style="display:none">
                        <div class="gs-print-margin-grid">
                            <label class="gs-print-mini-field"><span>上</span><input class="gs-print-input" type="number" min="0" max="100" step="0.5" data-gs-print-role="margin-top"></label>
                            <label class="gs-print-mini-field"><span>下</span><input class="gs-print-input" type="number" min="0" max="100" step="0.5" data-gs-print-role="margin-bottom"></label>
                            <label class="gs-print-mini-field"><span>左</span><input class="gs-print-input" type="number" min="0" max="100" step="0.5" data-gs-print-role="margin-left"></label>
                            <label class="gs-print-mini-field"><span>右</span><input class="gs-print-input" type="number" min="0" max="100" step="0.5" data-gs-print-role="margin-right"></label>
                        </div>
                        <div class="gs-print-mini-actions">
                            <button type="button" class="gs-print-btn-sm gs-print-btn-primary" data-gs-print-action="margin-apply">应用</button>
                            <button type="button" class="gs-print-btn-sm" data-gs-print-action="margin-cancel">取消</button>
                        </div>
                    </div>
                    <div class="gs-print-divider"></div>
                    <div class="gs-print-field"><span class="gs-print-label">缩放</span>
                        <select class="gs-print-select" data-gs-print-role="scale">
                            <option value="none">无缩放</option>
                            <option value="fit-sheet">整个工作表</option>
                            <option value="fit-cols">所有列一页</option>
                            <option value="fit-rows">所有行一页</option>
                            <option value="custom">自定义</option>
                        </select>
                        <input class="gs-print-input" type="number" min="10" max="400" step="5" data-gs-print-role="scale-input" disabled>
                    </div>
                    <label class="gs-print-field"><span class="gs-print-label">页面范围</span>
                        <select class="gs-print-select" data-gs-print-role="page-range">
                            <option value="all">全部</option>
                            <option value="custom">页码范围</option>
                        </select>
                    </label>
                    <div class="gs-print-field" data-gs-print-role="page-range-row" style="display:none">
                        <span class="gs-print-label"></span>
                        <input class="gs-print-input" type="text" placeholder="如：1-3,5" data-gs-print-role="page-range-input">
                    </div>
                    <div class="gs-print-actions">
                        <button class="gs-print-btn" data-gs-print-action="print">打印</button>
                        <button class="gs-print-btn" data-gs-print-action="close">关闭</button>
                    </div>
                </div>
                </div>
            </div>
        `;
        document.body.appendChild(overlay);

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
        // 预览缩放：页脚滑块 + 百分比输入框（10-400），覆盖自动适配（null=自动按可用高度适配）
        const zoomSlider = overlay.querySelector<HTMLInputElement>('[data-gs-print-role="zoom-slider"]');
        const zoomInput = overlay.querySelector<HTMLInputElement>('[data-gs-print-role="zoom-input"]');
        this._zoomSlider = zoomSlider;
        this._zoomInput = zoomInput;
        if (zoomSlider) {
            // 拖动跟手保障：range 原生拖拽手势在此环境下会失效（拖动不跟手）。
            // 统一改用指针事件自实现拖动：按下后按指针位置换算滑块值（按 step 吸附）
            // 并派发 input 事件，复用下方监听完成缩放应用与双向同步；
            // 键盘方向键仍走 range 原生行为
            const applyPointer = (clientX: number) => {
                const rect = zoomSlider.getBoundingClientRect();
                if (rect.width <= 0) return;
                const min = Number(zoomSlider.min) || 10;
                const max = Number(zoomSlider.max) || 400;
                const step = Number(zoomSlider.step) || 1;
                const ratio = Math.min(1, Math.max(0, (clientX - rect.left) / rect.width));
                const snapped = Math.round((min + ratio * (max - min)) / step) * step;
                zoomSlider.value = String(Math.min(max, Math.max(min, snapped)));
                zoomSlider.dispatchEvent(new Event('input'));
            };
            zoomSlider.addEventListener('pointerdown', (e) => {
                if (e.pointerType === 'mouse' && e.button !== 0) return;  // 仅响应鼠标左键
                e.preventDefault();  // 接管拖动，避免原生手势叠加；焦点被阻止后手动恢复
                zoomSlider.focus();
                // 指针捕获：拖出滑块/窗口范围仍持续跟手；不支持时退化为窗口级监听
                try { zoomSlider.setPointerCapture(e.pointerId); } catch { /* ignore */ }
                applyPointer(e.clientX);  // 按下即跳转到点击位置（与原生轨道点击行为一致）
                const onMove = (ev: PointerEvent) => applyPointer(ev.clientX);
                const onUp = () => {
                    window.removeEventListener('pointermove', onMove);
                    window.removeEventListener('pointerup', onUp);
                    window.removeEventListener('pointercancel', onUp);
                };
                window.addEventListener('pointermove', onMove);
                window.addEventListener('pointerup', onUp);
                window.addEventListener('pointercancel', onUp);
            });
            // 阻断 range 原生拖拽手势，由上方指针拖动统一接管（鼠标专有兜底）
            zoomSlider.addEventListener('mousedown', (e) => e.preventDefault());
            zoomSlider.addEventListener('input', () => {
                const n = Number(zoomSlider.value);
                if (!Number.isFinite(n)) return;
                this._previewZoomPercent = Math.min(400, Math.max(10, n));
                // 传 false：拖动期间不回写滑块 value，仅同步右侧数字输入框
                this._applyPreviewZoom(false);
                if (this._zoomInput && document.activeElement !== this._zoomInput) {
                    this._zoomInput.value = String(this._previewZoomPercent);
                }
            });
        }
        if (zoomInput) {
            zoomInput.addEventListener('change', () => {
                const n = Number(zoomInput.value);
                if (Number.isFinite(n) && n > 0) {
                    this._previewZoomPercent = Math.min(400, Math.max(10, Math.round(n)));
                }
                this._applyPreviewZoom();
            });
        }
        this._renderPagesInto(pagesContainer);

        // 事件委托：设置区打印/关闭按钮 + 底部分页导航按钮
        overlay.addEventListener('click', (e) => {
            const target = (e.target as HTMLElement).closest('[data-gs-print-action], [data-gs-print-nav]') as HTMLElement | null;
            if (!target) return;
            if (target.dataset.gsPrintAction === 'print') this.print();
            if (target.dataset.gsPrintAction === 'close') this.closePreview();
            if (target.dataset.gsPrintNav === 'prev') this._goToPage(this.currentPage - 1);
            if (target.dataset.gsPrintNav === 'next') this._goToPage(this.currentPage + 1);
        });

        // Esc 键关闭预览
        this.boundPreviewKeyDown = (e: KeyboardEvent) => {
            if (e.key === 'Escape') this.closePreview();
        };
        document.addEventListener('keydown', this.boundPreviewKeyDown);

        // 窗口尺寸变化：仅重算预览页等比缩放（高度始终填满可用区域），不重建内容、不跳页
        this.boundPreviewResize = () => this._applyPreviewZoom();
        window.addEventListener('resize', this.boundPreviewResize);

        this.overlay = overlay;
    }

    /**
     * 关闭打印预览遮罩层并清理事件监听
     */
    public closePreview(): void {
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
        this._zoomSlider = null;
        this._zoomInput = null;
        this._previewZoomPercent = null;
        this.pageDivs = [];
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
     * 不干扰当前页面，打印结束（afterprint）后自动移除 iframe
     */
    public print(): void {
        // 借助 iframe 隔离打印文档，避免当前页面样式与结构干扰
        const iframe = document.createElement('iframe');
        iframe.style.position = 'fixed';
        iframe.style.right = '0';
        iframe.style.bottom = '0';
        iframe.style.width = '0';
        iframe.style.height = '0';
        iframe.style.border = '0';
        document.body.appendChild(iframe);

        const win = iframe.contentWindow;
        if (!win) {
            iframe.remove();
            return;
        }

        // 打印结束后清理 iframe（afterprint 兼容性兜底：120s 后强制移除）
        win.addEventListener('afterprint', () => iframe.remove());
        window.setTimeout(() => iframe.remove(), 120000);

        const doc = win.document;
        doc.open();
        doc.write(this._buildPrintDocument(this._expandCopies(this._filterPageRange(this._buildPageHtmls()))));
        doc.close();
        // 等待文档解析与图片加载完成后再唤起打印对话框
        window.setTimeout(() => {
            win.focus();
            win.print();
        }, 50);
    }

    /**
     * 将按当前设置生成的各页表格渲染进预览容器
     * 预览为单页显示模式：全部页面预先生成，仅当前页可见，
     * 通过底部导航切换；纸张设置变更重绘后回到第 1 页。
     * 预览页面与实际纸张等比例：宽高按当前纸张大小/方向内联指定（mm），
     * 整页以 zoom 等比缩放——高度填满滚动区可用高度，宽度与内容随同一系数缩放。
     * @param {HTMLElement | null} container - 页面容器（.gs-print-pages）
     * @private
     */
    private _renderPagesInto(container: HTMLElement | null): void {
        if (!container) return;
        container.innerHTML = '';
        this.pageDivs = [];
        // 纸张在当前方向下的宽高（mm），与打印文档 .gs-print-sheet 的尺寸规则一致
        const size = PAPER_SIZES[this.paperSize];
        const landscape = this.orientation === 'landscape';
        const paperW = `${landscape ? size.height : size.width}mm`;
        const paperH = `${landscape ? size.width : size.height}mm`;
        for (const pageHtml of this._buildPageHtmls()) {
            const pageDiv = document.createElement('div');
            pageDiv.className = 'gs-print-page';
            pageDiv.style.width = paperW;
            pageDiv.style.height = paperH;
            pageDiv.style.padding = this._marginCss();
            pageDiv.innerHTML = pageHtml;
            container.appendChild(pageDiv);
            this.pageDivs.push(pageDiv);
        }
        this._applyPreviewZoom();
        this.totalPages = this.pageDivs.length;
        this._goToPage(0);
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
     * 将等比缩放系数应用到全部预览页（zoom 整体缩放：宽度、页边距与内容
     * 随同一系数缩放，矢量渲染不失真）。
     * 缩放来源：滑块/输入框的手动覆盖值（10-400%）优先；否则自动按可用高度适配。
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
        for (const div of this.pageDivs) {
            if (zoom !== 1) {
                div.style.setProperty('zoom', String(zoom));
            } else {
                div.style.removeProperty('zoom');
            }
        }
        if (!syncControls) return;
        const percent = String(Math.round(zoom * 100));
        if (this._zoomSlider) this._zoomSlider.value = percent;
        if (this._zoomInput && document.activeElement !== this._zoomInput) this._zoomInput.value = percent;
    }

    /**
     * 跳转到指定预览页（越界自动收敛），仅显示当前页并同步底部导航状态
     * @param {number} page - 目标页索引（0-based）
     * @private
     */
    private _goToPage(page: number): void {
        this.currentPage = this.totalPages > 0 ? Math.max(0, Math.min(page, this.totalPages - 1)) : 0;
        this.pageDivs.forEach((div, i) => {
            div.style.display = i === this.currentPage ? '' : 'none';
        });
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
     * 核心：按当前纸张设置计算分页并逐页生成表格 HTML。
     * 内容按打印缩放系数（无缩放/适配一页/自定义比例）缩放后与纸张可打印区域比较：
     * 宽度超出可打印宽度时水平分页（列不跨页拆分），高度超出时垂直分页（行不跨页拆分），
     * 总页数 = 列组数 × 行组数，打印顺序「先下后右」；设置打印区域时仅区域内行列参与分页。
     * 流程：筛选打印区域内的可见行列 → 解析缩放系数 → 建立缩放后全局布局 →
     * 水平/垂直贪心分页 → 逐页生成表格。
     * @returns {string[]} 各页表格 HTML
     * @private
     */
    private _buildPageHtmls(): string[] {
        const matrix = this.data.matrixValues();
        const rows = matrix.length;
        const cols = rows > 0 ? matrix[0].length : 0;
        const colHeaders = this.data.colHeaders.colHeaders;
        const rowHeaders = this.data.rowHeaders.rowHeaders;

        // 1) 参与打印的行列：隐藏行/列不参与打印；设置打印区域后仅保留区域内的行列
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
        if (visCols.length === 0 || visRows.length === 0) {
            // 完全空白的新文档（矩阵无行列）也生成一张空白页：
            // 预览始终显示纸张白区、打印输出一张空页（对齐 Excel 空表行为）
            if (rows === 0 || cols === 0) return [''];
            return [];
        }

        // 2) 纸张几何：可打印区域 = 纸张尺寸 - 上下左右页边距（当前预设）
        const size = PAPER_SIZES[this.paperSize];
        const landscape = this.orientation === 'landscape';
        const paperW = (landscape ? size.height : size.width) * MM_TO_PX;
        const paperH = (landscape ? size.width : size.height) * MM_TO_PX;
        const m = this.margins;
        const printableW = paperW - (m.left + m.right) * MM_TO_PX;
        const printableH = paperH - (m.top + m.bottom) * MM_TO_PX;

        // 3) 缩放系数：按缩放模式解析（fit 系列只缩小不放大），布局与字体统一按其缩放
        let totalW = 0;
        for (const c of visCols) totalW += colHeaders[c]?.width ?? 100;
        let totalH = 0;
        for (const r of visRows) totalH += rowHeaders[r]?.height ?? 24;
        const scale = this._resolveScale(totalW, totalH, printableW, printableH);
        this.renderScale = scale;

        // 4) 可见行列的全局布局（缩放后尺寸），供分页与合并块跨页定位使用
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

        // 5) 水平分页：列宽累加超出可打印宽度即换页（列不跨页拆分）
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
        if (curCols.length > 0) {
            colPages.push({ cols: curCols });
        }

        // 6) 垂直分页：行高累加超出可打印高度即换页（行不跨页拆分）
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
        if (curRows.length > 0) {
            rowPages.push({ rows: curRows, heights: curHeights });
        }

        // 7) 合并块扫描（全表一次）：记录锚点、跨度与块内可见首末行列（跨页截断判断用）
        const blocks: Array<{
            cell: Cell; r0: number; c0: number; rs: number; cs: number;
            firstRow: number; lastRow: number; firstCol: number; lastCol: number;
        }> = [];
        for (const row of matrix) {
            for (const cell of row) {
                if (!cell) continue;
                const rs = cell.rowspan ?? 1;
                const cs = cell.colspan ?? 1;
                if (rs <= 1 && cs <= 1) continue;
                const { row: r1, col: c1 } = this.data.getCellColAndRow(cell.cell);
                const r0 = r1 - 1;
                const c0 = c1 - 1;
                // 块内可见首末行列（globalX/globalY 仅含可见行列）
                let firstRow = -1;
                let lastRow = -1;
                let firstCol = -1;
                let lastCol = -1;
                for (let dr = 0; dr < rs; dr++) if (globalY.has(r0 + dr)) { if (firstRow < 0) firstRow = r0 + dr; lastRow = r0 + dr; }
                for (let dc = 0; dc < cs; dc++) if (globalX.has(c0 + dc)) { if (firstCol < 0) firstCol = c0 + dc; lastCol = c0 + dc; }
                if (firstRow < 0 || firstCol < 0) continue;  // 块内无可见行/列
                blocks.push({ cell, r0, c0, rs, cs, firstRow, lastRow, firstCol, lastCol });
            }
        }

        // 8) 逐页生成：打印顺序「先下后右」——同一列组内的行页从上到下，再进入下一列组
        const pages: string[] = [];
        for (const colPage of colPages) {
            for (const rowPage of rowPages) {
                pages.push(this._buildPageTable(rowPage, colPage, blocks, globalX, matrix));
            }
        }
        return pages;
    }

    /**
     * 生成本页表格 HTML：table-layout: fixed + colgroup 显式列宽 + 固定行高，
     * 文字与边框交由浏览器渲染/打印引擎矢量输出（精度优于 canvas 位图光栅化）。
     * 合并块与本页窗口求交集：块体仅在本页发出一个 td——位于块内首个可见位置，
     * colspan/rowspan 收缩为本页可见范围；被页边界截断的块体外缘不绘制边框。
     * @param {PageGeometry} rowPage - 本页行几何
     * @param {ColPage} colPage - 本页列组
     * @param {Array<{cell: Cell; r0: number; c0: number; rs: number; cs: number; firstRow: number; lastRow: number; firstCol: number; lastCol: number}>} blocks - 全表合并块列表
     * @param {Map<number, {x: number; w: number}>} globalX - 可见列全局布局（列宽来源）
     * @param {(Cell | undefined)[][]} matrix - 数据矩阵
     * @returns {string} 本页表格 HTML
     * @private
     */
    private _buildPageTable(
        rowPage: PageGeometry,
        colPage: ColPage,
        blocks: Array<{ cell: Cell; r0: number; c0: number; rs: number; cs: number; firstRow: number; lastRow: number; firstCol: number; lastCol: number; }>,
        globalX: Map<number, { x: number; w: number }>,
        matrix: (Cell | undefined)[][]
    ): string {
        const pageRowSet = new Set(rowPage.rows);
        const pageColSet = new Set(colPage.cols);
        const rowHeight = new Map<number, number>();
        rowPage.rows.forEach((r, i) => rowHeight.set(r, rowPage.heights[i]));
        /** 合并块在本页的发出规格：收缩后的跨度、本页可见高度与四边边框显隐 */
        interface EmitSpec { cell: Cell; colspan: number; rowspan: number; height: number; showTop: boolean; showBottom: boolean; showLeft: boolean; showRight: boolean; }
        const emits = new Map<string, EmitSpec>();
        const skip = new Set<string>();
        for (const b of blocks) {
            const visCols: number[] = [];
            for (let dc = 0; dc < b.cs; dc++) if (pageColSet.has(b.c0 + dc)) visCols.push(b.c0 + dc);
            const visRows: number[] = [];
            for (let dr = 0; dr < b.rs; dr++) if (pageRowSet.has(b.r0 + dr)) visRows.push(b.r0 + dr);
            if (visCols.length === 0 || visRows.length === 0) continue;  // 与本页无交集
            let height = 0;
            for (const r of visRows) height += rowHeight.get(r) ?? 0;
            emits.set(`${visRows[0]}:${visCols[0]}`, {
                cell: b.cell,
                colspan: visCols.length,
                rowspan: visRows.length,
                height,
                showTop: visRows[0] === b.firstRow,
                showBottom: visRows[visRows.length - 1] === b.lastRow,
                showLeft: visCols[0] === b.firstCol,
                showRight: visCols[visCols.length - 1] === b.lastCol,
            });
            for (const r of visRows) {
                for (const c of visCols) skip.add(`${r}:${c}`);
            }
        }
        // 表格宽度 = 本页列宽之和：列组可能不满页宽，不能被固定布局拉伸补齐
        let tableW = 0;
        for (const c of colPage.cols) tableW += globalX.get(c)!.w;
        let html = `<table class="gs-print-table" style="width:${tableW}px"><colgroup>`;
        for (const c of colPage.cols) html += `<col style="width:${globalX.get(c)!.w}px">`;
        html += '</colgroup><tbody>';
        for (let i = 0; i < rowPage.rows.length; i++) {
            const r = rowPage.rows[i];
            html += `<tr style="height:${rowPage.heights[i]}px">`;
            for (const c of colPage.cols) {
                const key = `${r}:${c}`;
                const emit = emits.get(key);
                if (emit) {
                    html += this._buildCellTd(emit.cell, emit.colspan, emit.rowspan, emit.height, emit.showTop, emit.showBottom, emit.showLeft, emit.showRight);
                    continue;
                }
                if (skip.has(key)) continue;
                html += this._buildCellTd(matrix[r]?.[c], 1, 1, rowPage.heights[i], true, true, true, true);
            }
            html += '</tr>';
        }
        html += '</tbody></table>';
        return html;
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
     * 按页面范围设置过滤各页表格 HTML（仅对实际打印输出生效，预览始终显示全部页）；
     * 页序与 _buildPageHtmls 的输出顺序一致（「先下后右」）
     * @param {string[]} pages - 各页表格 HTML
     * @returns {string[]} 过滤后的页面序列
     * @private
     */
    private _filterPageRange(pages: string[]): string[] {
        const hits = this._parsePageRange(pages.length);
        if (!hits) return pages;
        return pages.filter((_, i) => hits.has(i + 1));
    }

    /**
     * 按份数与打印顺序展开页面序列
     * 逐份打印（copies）：1,2,3 → 1,2,3,1,2,3；逐页打印（page）：1,2,3 → 1,1,2,2,3,3
     * @param {string[]} pages - 各页表格 HTML
     * @returns {string[]} 展开后的页面序列
     * @private
     */
    private _expandCopies(pages: string[]): string[] {
        const copies = Math.max(1, Math.floor(this.copies));
        if (copies <= 1 || pages.length === 0) return pages;
        const seq: string[] = [];
        if (this.collate === 'copies') {
            for (let i = 0; i < copies; i++) seq.push(...pages);
        } else {
            for (const p of pages) {
                for (let i = 0; i < copies; i++) seq.push(p);
            }
        }
        return seq;
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
     * 生成完整打印文档：各页表格 HTML 按页分页输出。
     * 自定义打印设置同步到浏览器打印对话框的方式：
     * - 纸张大小与方向：通过 @page size 规则同步，Chrome/Edge 打印预览随其变化；
     * - 页眉页脚：浏览器页眉页脚渲染在页边距区域内，@page 边距设为 0 后不再显示；
     *   页边距改由逐页包裹层（.gs-print-sheet）的 padding 实现，版面效果与原来一致；
     * - 打印机 / 份数 / 顺序：Web API 无法预设浏览器打印对话框，需在对话框中确认。
     * @param {string[]} pages - 各页表格 HTML
     * @returns {string} 可直接写入 iframe 的 HTML 文档字符串
     * @private
     */
    private _buildPrintDocument(pages: string[]): string {
        const size = PAPER_SIZES[this.paperSize];
        // 纸张在当前方向下的实际高度（mm）：@page 边距为 0，逐页容器必须与纸张等高，
        // 否则页高超出可打印区域会导致出现空白页
        const paperH = this.orientation === 'portrait' ? size.height : size.width;
        const sheets = pages
            .map(pageHtml => `<div class="gs-print-sheet">${pageHtml}</div>`)
            .join('');
        return `<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>${this.title}</title>
    <style>
        /* 页边距必须为 0：浏览器页眉页脚渲染在页边距内，边距为 0 时即不再显示 */
        @page { size: ${size.page} ${this.orientation}; margin: 0; }
        html, body { margin: 0; padding: 0; }
        /* 逐页容器：与纸张等高，页边距内置于 padding，超出部分裁剪防溢出产生空白页 */
        .gs-print-sheet { box-sizing: border-box; width: 100%; height: ${paperH}mm; padding: ${this._marginCss()}; overflow: hidden; page-break-after: always; break-after: page; }
        .gs-print-sheet:last-child { page-break-after: auto; break-after: auto; }
        /* 打印时强制保留背景色与文字颜色（浏览器默认不打印背景） */
        .gs-print-sheet .gs-print-table { print-color-adjust: exact; -webkit-print-color-adjust: exact; }
        ${this._tableStyles()}
        /* 自定义字体：iframe 为独立文档，主文档字体不共享，需重新注入 @font-face */
        ${this._getFontFaceCss()}
    </style>
</head>
<body>${sheets}</body>
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
