import { SHEETS_NAME, DEFAULT_CELL_WIDTH, DEFAULT_CELL_HEIGHT, DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE } from "../constant";
import { createDiv } from '../../utils/dom';
import { EventEmitter } from '../../utils/eventEmitter';
import { debounce } from '../../utils/debounce';
import { DataCollection } from "../dataArchitecture/dataCollection";
import { RowHeaderCanvas } from "./rowHeaderCanvas";
import { ColHeaderCanvas } from "./colHeaderCanvas";
import { SheetCanvas } from "./sheetCanvas";
import { AllSelect } from "./allSelect";
import { EditCanvas } from "./editCanvas";
import { CopyMarquee } from "./copyMarquee";
import { Menu } from "../common/menu";
import '../style/sheets.scss';

/**
 * Sheets 组件选项接口
 * @interface SheetsOptions
 */
export interface SheetsOptions {
    /** 父容器元素 */
    parentElement: HTMLElement;
    /** 数据实例 */
    data: DataCollection;
    /** 菜单实例 */
    menu?: Menu;
    /** 尺寸变化回调函数 */
    onResize?: (top: number, left: number, width: number, height: number) => void;
    /** 是否启用防抖，默认 true */
    useDebounce?: boolean;
    /** 防抖延迟时间（毫秒），默认 16ms */
    debounceDelay?: number;
}

/**
 * Sheets 组件类
 * 作为 canvas 的父容器，负责监听容器尺寸变化并触发事件
 * 
 * @class Sheets
 * @extends EventEmitter
 * @example
 * const sheets = new Sheets({
 *     parentElement: document.getElementById('container'),
 *     onResize: (top, left, width, height) => {
 *         console.log('Sheets resized:', width, height);
 *     }
 * });
 * 
 * // 或者通过事件监听
 * sheets.on('resize', (top, left, width, height) => {
 *     console.log('Resized:', width, height);
 * });
 */
export class Sheets extends EventEmitter {
    /** Sheets 容器元素 */
    private sheets: HTMLDivElement;
    /** 数据实例 */
    private data: DataCollection;
    /** 父容器元素 */
    private parentElement: HTMLElement;
    /** 菜单实例 */
    private menu?: Menu;
    /** 全选选择框元素 */
    private allSelect: AllSelect;
    /** 列头选择框元素 */
    private colHeaderCanvas: ColHeaderCanvas;
    /** 行头选择框元素 */
    private rowHeaderCanvas: RowHeaderCanvas;
    /** 工作表画布元素 */
    private sheetCanvas: SheetCanvas;
    /** 编辑画布元素 */
    private editCanvas: EditCanvas;
    /** 复制/剪切后的流动虚线框（DOM 层，位置由每帧 draw 同步） */
    private copyMarquee: CopyMarquee | null = null;
    /**
     * 表格区域内是否正在拖拽（mousedown 起于画布、尚未 mouseup）
     *
     * 用于区分「表格拖拽」与「在交互控件上操作」：前者需持续接管 mousemove
     * （即使指针途经输入框），后者必须放行默认行为。
     */
    private isGridDragging = false;

    /** ResizeObserver 实例 */
    private resizeObserver: ResizeObserver | null;
    /** 防抖后的 resize 处理函数 */
    private debouncedHandleResize: ((entries: ResizeObserverEntry[]) => void) | null;
    /** CSS 自定义字体是否已加载完成 */
    private fontsLoaded: boolean = false;
    /** IndexDB 数据是否已加载完成，未完成前不触发 resize 事件 */
    private dataLoaded: boolean = false;
    /** 加载动画动画帧 ID */
    private loadingAnimationId: number | null = null;
    /** 加载动画 Promise（守卫变量，防止并发调用 startLoadingAnimation） */
    private loadingPromise: Promise<void> | null = null;
    /** 合帧周期内的 draw() Promise（非 null 表示已排入 rAF，同帧多次调用共享同一绘制） */
    private _drawPromise: Promise<void> | null = null;
    /** draw() 排入的 rAF 回调 ID（用于销毁/重建时取消挂起重绘） */
    private _drawRafId: number | null = null;

    /**
     * 构造函数
     * @param {SheetsOptions} options - 初始化选项
     */
    constructor(options: SheetsOptions) {
        super();
        this.data = options.data;
        this.parentElement = options.parentElement;
        this.menu = options.menu;
        this.resizeObserver = null;
        this.debouncedHandleResize = null;
        
        if (options.onResize) {
            this.on('resize', options.onResize);
        }
        
        this.initSheets(options.useDebounce ?? true, options.debounceDelay ?? 16);
    }

    /**
     * 异步初始化 Sheets
     * @param {boolean} useDebounce - 是否启用防抖
     * @param {number} debounceDelay - 防抖延迟时间
     */
    private async initSheets(useDebounce: boolean, debounceDelay: number): Promise<void> {
        this.createSheets().then(async () => {
            await this.data.syncValuesFromDB();
            this.initVisibleView();
            this.ensureDefaultHeaders();
            this.dataLoaded = true;
            this.createCanvas();
            this.createAllSelect();
            this.setupResizeObserver(useDebounce, debounceDelay);
            this.setupEventListeners();
        });
    }

    /**
     * 创建 Sheets 容器元素
     */
    private async createSheets(): Promise<void> {
        this.sheets = createDiv({
            className: SHEETS_NAME.SHEETS,
            style: {
                position: 'relative',
                // top: '0',
                // left: '0',
                width: '100%',
                height: '100%',
                overflow: 'hidden',
            }
        });
        this.mountToParent(this.parentElement);
        await new Promise(resolve => requestAnimationFrame(resolve));
    }

    /**
     * 将 Sheets 挂载到父容器
     * @param {HTMLElement} parentElement - 父容器元素
     */
    private mountToParent(parentElement: HTMLElement): void {
        if (!(parentElement instanceof HTMLElement)) {
            throw new TypeError('parentElement must be a valid HTMLElement');
        }
        parentElement.appendChild(this.sheets);
    }
    
    /**
     * 创建全选选择框元素
     */
    private createAllSelect(): void {
        this.allSelect = new AllSelect(
            {
                parentElement: this.sheets,
                data: this.data
            }
        );        
        
    }    

    /**
     * 创建行头、列头和内容区元素
     */
    /**
     * 创建工作表画布元素
     */
    private createCanvas(): void {
        this.rowHeaderCanvas = new RowHeaderCanvas({parentElement: this.sheets, data: this.data, menu: this.menu});    
        this.colHeaderCanvas = new ColHeaderCanvas({parentElement: this.sheets, data: this.data, menu: this.menu});
        this.sheetCanvas = new SheetCanvas({parentElement: this.sheets, data: this.data, menu: this.menu});
        this.editCanvas = new EditCanvas({parentElement: this.sheets, data: this.data, menu: this.menu});
        // 流动虚线框挂在最上层 DOM（不画在 canvas 上）：CSS 动画跑在合成线程，
        // 不必为「虚线跑动」每帧全量重绘表格
        this.copyMarquee = new CopyMarquee({parentElement: this.sheets, data: this.data});
    }

    /**
     * 更新所有画布大小
     */
    public updateCanvasSize(): void {
        this.rowHeaderCanvas.updateCanvasSize();
        this.colHeaderCanvas.updateCanvasSize();
        this.sheetCanvas.updateCanvasSize();
        // 全选区（左上角）尺寸只在构造时按当时的行头宽度设过一次，行头宽度变化后需同步
        this.allSelect?.updateSize();
        this.draw();
    }

    /**
     * 等待 CSS 自定义字体加载完成
     * @param {number} timeout - 超时时间（毫秒），默认 3000ms
     * @returns {Promise<void>}
     */
    private async waitForFontsLoaded(timeout: number = 3000): Promise<void> {
        if (this.fontsLoaded) return;
        if (typeof document.fonts === 'undefined') { // 浏览器不支持 document.fonts API
            this.fontsLoaded = true;
            return;
        }
        try {
            // 使用 document.fonts.load() 显式触发字体加载
            // document.fonts.ready 仅等待"已经开始加载的字体"完成，不会主动触发加载；
            // canvas 的 fillText 使用的字体若未被 DOM 元素引用，document.fonts.ready 会立即 resolve，
            // 导致首次绘制时字体尚未就绪，使用 fallback 字体渲染。
            for (const cell of this.data.values) {
                for (const char of cell.chars) {
                    if (char.fontFamily) {
                        document.fonts.load(`${char.fontSize || DEFAULT_FONT_SIZE}px "${char.fontFamily}"`);
                    }
                }
            }
            await Promise.race([
                document.fonts.ready,
                new Promise((_, reject) => setTimeout(() => reject(new Error('Font load timeout')), timeout))
            ]);
        } catch (error) {
            console.warn('Font loading timeout or error:', error);
        }
        this.fontsLoaded = true;
    }

    /**
     * 启动加载动画
     * 若动画已在进行中，返回已有的 Promise 避免并发启动多个动画循环
     * @returns {Promise<void>}
     */
    private startLoadingAnimation(): Promise<void> {
        if (this.loadingPromise) return this.loadingPromise;

        this.loadingPromise = new Promise<void>((resolve) => {
            let rotation = 0;
            const animate = () => {
                rotation += 0.05;
                this.sheetCanvas.drawLoadingAnimation(rotation);
                this.loadingAnimationId = requestAnimationFrame(animate);
            };
            animate();

            this.waitForFontsLoaded().then(() => {
                if (this.loadingAnimationId !== null) {
                    cancelAnimationFrame(this.loadingAnimationId);
                    this.loadingAnimationId = null;
                }
                this.loadingPromise = null;
                resolve();
            });
        });
        return this.loadingPromise;
    }

    /**
     * 绘制所有画布内容（rAF 合帧）
     *
     * 同一帧/同一宏任务内多次调用 draw() 只真正绘制一次——多次请求合并到单个
     * requestAnimationFrame 回调中执行。事件链（VALUES_CHANGED / SELECTION_CHANGED /
     * ALL_*_CHANGED 等）与右键菜单钩子 refreshAfterContextMenu 常在一次动作中并发触发
     * draw，合帧后避免重复全量重绘（主画布全量重绘成本高）。
     * 首次调用仍会等待 CSS 自定义字体加载完成后再绘制。
     *
     * 实现要点：
     * - _drawPromise 非 null 即代表本帧已排入绘制，新调用直接复用该 Promise，不再排 rAF；
     * - rAF 回调内执行真实绘制，finally 中清空 _drawPromise（释放合帧锁）并 resolve，
     *   保证异常路径也不会卡死后续 draw()；
     * - canvas 未创建时（DB 加载阶段 OFFSETX/Y_CHANGED 同步回调）走守卫 return，
     *   首次正式绘制由 setupResizeObserver 的 ResizeObserver 首次回调兜底；
     * - 绘制前先按「最后一行行号」的文本宽度同步行头宽度（行号位数变多时行头需变宽），
     *   宽度有变才触发一次重排——内容区可用宽度依赖行头宽度，必须先重排再画。
     * @returns {Promise<void>} 共享的绘制 Promise（合帧周期内所有调用方拿到同一个）
     */
    public draw(): Promise<void> {
        // 已有挂起的合帧绘制：直接复用，不再排入新 rAF
        if (this._drawPromise) return this._drawPromise;
        this._drawPromise = new Promise<void>((resolve) => {
            this._drawRafId = requestAnimationFrame(async () => {
                this._drawRafId = null;
                try {
                    if (!this.sheetCanvas) return;
                    if (!this.fontsLoaded) await this.startLoadingAnimation();
                    // 行号位数变化（滚动到更深的行 / 打开新文件 / 增删行）时，行头宽度要跟随
                    // 「最后一行行号」的文本宽度。宽度变了就**先重排再绘制**：内容区可用宽度
                    // = 容器宽 − 右侧工具条 − 行头宽，重排把新宽度分发给各画布，
                    // 于是本帧就用新宽度绘制，不会出现一帧的错位/裁切
                    if (this.rowHeaderCanvas.syncRowHeaderWidth()) this.relayout();
                    this.rowHeaderCanvas.draw();
                    this.colHeaderCanvas.draw();
                    this.sheetCanvas.draw();
                    this.editCanvas.draw();
                    // 虚线框位置跟随滚动/缩放/选区变化：每帧按最新的 copyRange 重算
                    this.copyMarquee?.sync();
                    // 注意：这里**不再**调用 data.syncValuesToDB()。
                    // 每帧整表 toJSON() 序列化 + 写 IndexedDB 是打开大文件卡顿/崩溃的成因之一；
                    // 落库已改由 DataCollection 的「变更事件 → 置脏 → 防抖保存」驱动
                    // （见 DataCollection._markDirty / _scheduleAutoSave）。
                } finally {
                    // 释放合帧锁：本帧绘制结束（或被守卫跳过）后，后续 draw() 可排入新帧
                    this._drawPromise = null;
                    resolve();
                }
            });
        });
        return this._drawPromise;
    }

    /**
     * 设置 ResizeObserver 监听
     * @param {boolean} useDebounce - 是否启用防抖
     * @param {number} debounceDelay - 防抖延迟时间（毫秒）
     */
    private setupResizeObserver(useDebounce: boolean, debounceDelay: number): void {
        if (!this.sheets || !this.parentElement) {
            console.error('Sheets or parent element is not initialized');
            return;
        }

        if (this.resizeObserver) {
            this.resizeObserver.disconnect();
            this.resizeObserver = null;
        }

        const observerCallback: ResizeObserverCallback = (entries: ResizeObserverEntry[]) => {
            if (useDebounce && this.debouncedHandleResize) {
                this.debouncedHandleResize(entries);
            } else {
                this.handleResize(entries);
            }
        };

        if (useDebounce) {
            this.debouncedHandleResize = debounce((entries: ResizeObserverEntry[]) => {
                this.handleResize(entries);
            }, debounceDelay);
        }

        this.resizeObserver = new ResizeObserver(observerCallback);
        
        this.resizeObserver.observe(this.sheets, {
            box: 'content-box'
        });
    }

    /**
     * 初始化可见视图
     */
    public initVisibleView(): void {        
        if (this.data.visibleView.sheetWidth === 0 || this.data.visibleView.sheetHeight === 0) { // 初始化可见视图, 初次设置宽高为sheets容器尺寸
            this.data.visibleView.sheetWidth = window.innerWidth;
            this.data.visibleView.sheetHeight = window.innerHeight;
        }
    }

    /**
     * 确保表头数据已初始化
     * 从 IndexDB 加载后，若有数据则保留；若无数据则生成默认行列头
     */
    private ensureDefaultHeaders(): void {
        const cols = Math.ceil(this.data.visibleView.sheetWidth / DEFAULT_CELL_WIDTH / this.data.zoom) + 2;
        const rows = Math.ceil(this.data.visibleView.sheetHeight / DEFAULT_CELL_HEIGHT / this.data.zoom) + 2;
        if (this.data.colHeaders.length === 0) {
            this.data.colHeaders.addCol(0, cols);
        }
        if (this.data.rowHeaders.length === 0) {
            this.data.rowHeaders.addRow(0, rows);
        }
    }

    /**
     * 处理尺寸变化
     * @param {ResizeObserverEntry[]} entries - ResizeObserver 条目数组
     */
    private handleResize(entries: ResizeObserverEntry[]): void {
        if (!this.dataLoaded) return; // 数据未加载完成前不触发 resize 事件
        if (!entries || entries.length === 0) {
            return;
        }
        
        for (const entry of entries) {
            if (entry.target === this.sheets) {
                this.relayout();
                break;
            }
        }
    }

    /**
     * 按当前容器尺寸重排工作表（可见区宽高 + 各画布尺寸）
     *
     * 与 ResizeObserver 回调走同一条路径：对外广播 `resize` 事件，由宿主
     * （`app.onResize`）扣掉右侧工具条与行头宽度、算出内容区可用宽高再调
     * {@link updateCanvasSize}。除容器尺寸变化外，**行头宽度变化**（见
     * {@link RowHeaderCanvas.syncRowHeaderWidth}）也走这里——行头变宽会挤压内容区，
     * 必须让宿主按新的行头宽度重算，否则内容区会被行头盖住一截。
     * @private
     */
    private relayout(): void {
        if (!this.dataLoaded) return; // 数据未加载完成前不触发 resize 事件
        const rect = this.sheets.getBoundingClientRect();
        if (rect.width === 0 || rect.height === 0) return;  // DOM 没有加载完成，直接返回
        this.emitResize(rect.left, rect.top, rect.width, rect.height);
    }

    /**
     * 触发 resize 事件
     * @param {number} top - 顶部位置
     * @param {number} left - 左侧位置
     * @param {number} width - 宽度
     * @param {number} height - 高度
     */
    private emitResize(top: number, left: number, width: number, height: number): void {
        this.emit('resize', top, left, width, height);
    }

    /**
     * 判断鼠标事件是否源自交互控件（输入框 / 按钮 / 可编辑元素等）
     *
     * Sheets 的鼠标处理只服务于表格自身的画布，且会 `preventDefault()` 掉浏览器
     * 默认行为。而 mousedown 的默认行为之一就是**把焦点交给被点击的元素**——
     * 若事件来自交互控件（如右侧工具条里的缩放数值框），此处必须放行，
     * 否则点击输入框无法聚焦、无法键入。
     * @param e 鼠标事件对象
     * @returns 事件源自交互控件返回 true
     */
    private isFromInteractiveControl(e: MouseEvent): boolean {
        const target = e.target as HTMLElement | null;
        if (!target) return false;
        if (target.isContentEditable) return true;
        return ['INPUT', 'SELECT', 'TEXTAREA', 'BUTTON', 'OPTION', 'LABEL'].includes(target.tagName);
    }

    /**
     * 处理鼠标按下事件
     * @param e 鼠标事件对象
     */
    private handleMouseDown(e: MouseEvent): void {
        // 交互控件（右栏的缩放数值框等）的 mousedown 一律放行：
        // 否则 preventDefault 会吃掉「点击即聚焦」的默认行为
        if (this.isFromInteractiveControl(e)) return;

        this.isGridDragging = true;
        e.preventDefault();
        // textArea 焦点统一由 editCanvas 的文档级 mousedown 焦点管理处理：
        // 仅点击 sheetCanvas（主页面）时聚焦，点击其它位置（含行列表头）时失焦
        this.allSelect.handleMouseDown(e);
        this.colHeaderCanvas.handleMouseDown(e);
        this.rowHeaderCanvas.handleMouseDown(e);
        this.sheetCanvas.handleMouseDown(e);
    }

    /**
     * 处理鼠标移动事件
     * @param e 鼠标事件对象
     */
    private handleMouseMove(e: MouseEvent): void {
        // 拖拽已开始但漏收了 mouseup（典型场景：在本窗口外抬起鼠标后再移回页面）：
        // 此时左键已不再按下（e.buttons 位 0），必须按抬起收尾，否则 isGridDragging
        // 仍为 true，选区会继续跟随鼠标（事件源在窗口外时 document 收不到 mouseup）。
        if (this.isGridDragging && (e.buttons & 1) === 0) {
            this.handleMouseUp(e);
            return;
        }
        // 无表格拖拽进行中且事件源自交互控件时放行：不阻止其默认行为，
        // 使输入框内按住拖动选择文本仍可用；拖拽进行中则保持原有接管逻辑
        if (!this.isGridDragging && this.isFromInteractiveControl(e)) return;

        e.preventDefault();
        this.colHeaderCanvas.handleMouseMove(e);
        this.rowHeaderCanvas.handleMouseMove(e);
        this.sheetCanvas.handleMouseMove(e);
    }

    /**
     * 处理鼠标释放事件
     * @param e 鼠标事件对象
     */
    private handleMouseUp(e: MouseEvent): void {
        e.preventDefault();
        this.isGridDragging = false;
        // 先处理锚点选择并重置共享拖拽状态（isMouseDown/startLocation 等）
        this.sheetCanvas.handleMouseUp(e);
        // 再收尾行高/列宽拖拽：取消待执行的 rAF、同步应用最终值、提交连续撤销。
        // 收尾逻辑不依赖 isMouseDown（此时可能已被 sheetCanvas 重置）。
        this.rowHeaderCanvas.handleMouseUp(e);
        this.colHeaderCanvas.handleMouseUp(e);
    }

    /**
     * 处理鼠标双击事件
     * @param e 鼠标事件对象
     */
    private handleDoubleClick(e: MouseEvent): void {
        e.preventDefault();
        this.sheetCanvas.handleDoubleClick(e);
        this.colHeaderCanvas.handleDoubleClick(e);
        this.rowHeaderCanvas.handleDoubleClick(e);
    }

    private setupEventListeners(): void {
        this.sheets.addEventListener('mousedown', this.handleMouseDown.bind(this));
        document.addEventListener('mousemove', this.handleMouseMove.bind(this));
        document.addEventListener('mouseup', this.handleMouseUp.bind(this));
        this.sheets.addEventListener('dblclick', this.handleDoubleClick.bind(this));
    }

    /**
     * 获取 Sheets 容器元素
     * @returns {{ sheetCanvas: SheetCanvas,
     * rowHeader: RowHeaderCanvas,
     * sheets: HTMLDivElement,
     * allSelect: AllSelect, }} Sheets 元素对象
     */
    public getElements(): { 
        sheetCanvas: SheetCanvas;
        editCanvas: EditCanvas;
        inColHeader: ColHeaderCanvas;
        rowHeader: RowHeaderCanvas;
        sheets: HTMLDivElement;
        allSelect: AllSelect;
    } {
        return {
            sheetCanvas: this.sheetCanvas,
            editCanvas: this.editCanvas,
            inColHeader: this.colHeaderCanvas,
            rowHeader: this.rowHeaderCanvas,
            sheets: this.sheets,
            allSelect: this.allSelect
        };
    }

    /**
     * 销毁 Sheets 组件
     * 清理 ResizeObserver 和事件监听
     */
    public destroy(): void {
        if (this.resizeObserver) {
            this.resizeObserver.disconnect();
            this.resizeObserver = null;
        }
        
        this.removeAllListeners();

        this.copyMarquee?.destroy();
        this.copyMarquee = null;
        
        if (this.sheets.parentElement) {
            this.sheets.parentElement.removeChild(this.sheets);
        }
    }
}