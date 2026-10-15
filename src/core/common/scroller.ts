/* 
 * 表格的 scroller 面板组件
 * 支持水平和垂直两种滚动方式
 * 功能：可拖拽调整面板大小、记忆布局状态、响应式适应容器变化
 * author: 东方鹗
 * B站: https://space.bilibili.com/194359739
 * 知乎: https://www.zhihu.com/people/eastossifrage
 * CSDN: https://blog.csdn.net/os373
 */

import { createDiv } from "../../utils/dom";
import { EventEmitter } from '../../utils/eventEmitter';
import { SCROLLER_CLASS_NAMES, DataEvents, DEFAULT_CELL_WIDTH, DEFAULT_CELL_HEIGHT } from "../constant"; 
import { DataCollection } from "../dataArchitecture/dataCollection";
import '../style/scroller.scss';


/**
 * 滚动条元素接口
 */
export interface ScrollbarElements {
    container: HTMLDivElement; 
    track: HTMLDivElement;
    back: HTMLDivElement;
    thumb: HTMLDivElement;
    forward: HTMLDivElement;
}

/**
 * 滚动条配置选项接口
 */
interface ScrollbarOptions {
    parentElement?: HTMLElement | null;
    /** 滚轮事件监听目标（如包含 sheetCanvas/editCanvas/textarea 的工作表容器）；缺省回退 parentElement */
    wheelElement?: HTMLElement | null;
    direction?: 'horizontal' | 'vertical';
    data?: DataCollection;
    thumbMinSize?: number;
    debounceTime?: number;
    onChange?: () => void;
    onResize?: () => void;
}

/**
 * 滚动条组件类
 * @class Scroller
 * @extends EventEmitter
 */
export class Scroller extends EventEmitter {
    public options: ScrollbarOptions;
    private elements: ScrollbarElements;
    private rafId: number | null = null;
    private isDragging: boolean = false;
    private resizeObserver: ResizeObserver;
    private lastTouchIdentifier: number | null = null;
    private debounceTimer: number | null = null;
    private startDragPosition: number = 0;
    private startScrollPosition: number = 0;
    private stepIntervalId: number | null = null;
    private stepHoldTimer: number | null = null;
    private readonly STEP_DELAY = 300;
    private readonly STEP_INTERVAL = 50;
    private stepDirection: 'back' | 'forward' | null = null;
    /** 滚轮累积待应用的像素偏移量（合并高频 wheel 事件用） */
    private wheelTargetOffset: number = 0;
    /** 滚轮合并 rAF ID */
    private wheelRafId: number | null = null;
    /** 滚轮事件实际监听的元素（wheelElement ?? parentElement），destroy 时移除监听用 */
    private wheelListenerTarget: HTMLElement | null = null;
    /** 缩放步进因子（Ctrl+滚轮 每次 ±10%） */
    private static readonly ZOOM_FACTOR = 1.1;
    /** 缩放下限（对外公开：外部缩放控件如右侧工具栏滑块需与之保持一致） */
    public static readonly ZOOM_MIN = 0.25;
    /** 缩放上限（对外公开：同上） */
    public static readonly ZOOM_MAX = 4;

    // —— document 级具名处理器：注册/移除使用同一引用，保证 destroy 能真正解除监听 ——
    private readonly onDocumentMouseMove = (e: MouseEvent): void => this.handleDrag(e);
    /** mouseup 同时结束拖拽与步进（合并原先两个重复监听） */
    private readonly onDocumentMouseUp = (e: MouseEvent): void => {
        this.endDrag(e);
        this.stopStep();
    };
    private readonly onDocumentTouchMove = (e: TouchEvent): void => this.handleTouchMove(e);
    /** touchend 同时结束触摸与步进（合并原先两个重复监听） */
    private readonly onDocumentTouchEnd = (e: TouchEvent): void => {
        this.endTouch(e);
        this.stopStep();
    };
    private readonly onDocumentTouchCancel = (e: TouchEvent): void => this.endTouch(e);
    private readonly onWheelEvent = (e: WheelEvent): void => this.handleWheel(e);

    // —— 数据事件具名处理器：destroy 时可精确 off ——
    private readonly onOffsetXChanged = (): void => this.debouncedUpdate();
    private readonly onOffsetYChanged = (): void => this.debouncedUpdate();
    private readonly onAllColWidthChanged = (): void => this.debouncedUpdate();
    private readonly onAllRowHeightChanged = (): void => this.debouncedUpdate();
    private readonly onZoomChanged = (): void => this.debouncedUpdate();

    /**
     * 构造函数
     * @param {ScrollbarOptions} options - 配置选项
     */
    constructor(options: ScrollbarOptions = {}) {
        super();
        this.options = {
            direction: 'horizontal',
            thumbMinSize: 20,
            debounceTime: 100,
            ...options
        };
        this.init();

        if (options.parentElement) {
            options.parentElement.appendChild(this.elements.container);
        }
        if (options.onChange) {
            this.on('change', options.onChange);
        }
        if (options.onResize) {
            this.on('resize', options.onResize);
        }
        this.setupEventListeners();
        this.setupResizeObservers();
        this.setupDataListeners();
    }

    /**
     * 初始化滚动条元素
     */
    private init(): void {
        this.elements = {
            container: createDiv({ className: `${SCROLLER_CLASS_NAMES.CONTAINER}${this.options.direction}` }),
            track: createDiv({ className: `${SCROLLER_CLASS_NAMES.TRACK}` }),
            back: createDiv({ className: `${SCROLLER_CLASS_NAMES.BACK}` }),
            thumb: createDiv({ className: `${SCROLLER_CLASS_NAMES.THUMB}` }),
            forward: createDiv({ className: `${SCROLLER_CLASS_NAMES.FORWARD}` }),
        };

        const { container, track, back, thumb, forward } = this.elements;
        track.appendChild(thumb);
        container.append(track, back, forward);
    }

    /**
     * 设置事件监听器
     */
    private setupEventListeners(): void {
        const { back, thumb, track, forward } = this.elements;

        thumb.addEventListener('mousedown', (e) => this.startDrag(e));
        track.addEventListener('mousedown', (e) => this.handleTrackClick(e));

        thumb.addEventListener('touchstart', (e) => this.handleTouchStart(e), { passive: false });
        track.addEventListener('touchstart', (e) => this.handleTrackTouch(e), { passive: false });

        document.addEventListener('mousemove', this.onDocumentMouseMove);
        document.addEventListener('mouseup', this.onDocumentMouseUp);

        document.addEventListener('touchmove', this.onDocumentTouchMove, { passive: false });
        document.addEventListener('touchend', this.onDocumentTouchEnd);
        document.addEventListener('touchcancel', this.onDocumentTouchCancel);

        back.addEventListener('mousedown', (e) => this.startStep(e, 'back'));
        forward.addEventListener('mousedown', (e) => this.startStep(e, 'forward'));
        back.addEventListener('touchstart', (e) => this.startStep(e, 'back'));
        forward.addEventListener('touchstart', (e) => this.startStep(e, 'forward'));

        // 滚轮监听绑定到工作表容器（含 sheetCanvas/editCanvas/textarea），
        // 滚轮在画布或编辑态 textarea 上均可滚动主页面；未提供时回退 parentElement
        const wheelTarget = this.options.wheelElement ?? this.options.parentElement;
        if (wheelTarget) {
            this.wheelListenerTarget = wheelTarget;
            wheelTarget.addEventListener('wheel', this.onWheelEvent, { passive: false });
        }
    }

    /**
     * 设置尺寸变化监听
     */
    private setupResizeObservers(): void {
        const { container } = this.elements;

        this.resizeObserver = new ResizeObserver(() => {
            this.debouncedUpdate();
        });
        this.resizeObserver.observe(container);
    }

    /**
     * 设置数据变化监听
     */
    private setupDataListeners(): void {
        if (this.options.data) {
            this.options.data.colHeaders.on(DataEvents.OFFSETX_CHANGED, this.onOffsetXChanged);
            this.options.data.rowHeaders.on(DataEvents.OFFSETY_CHANGED, this.onOffsetYChanged);
            // 总宽/总高变化时也要更新滑块尺寸与位置：
            // 否则 addCol/addRow 后 allColWidth/allRowHeight 已变但滑块仍停留在
            // 初始的「占满 track」状态，导致 maxDelta=0 拖不动
            this.options.data.colHeaders.on(DataEvents.ALL_COL_WIDTH_CHANGED, this.onAllColWidthChanged);
            this.options.data.rowHeaders.on(DataEvents.ALL_ROW_HEIGHT_CHANGED, this.onAllRowHeightChanged);
            // 缩放变化影响滑块长度与位置（totalWidth * zoom）
            this.options.data.on(DataEvents.ZOOM_CHANGED, this.onZoomChanged);
        }
    }

    /**
     * 防抖更新方法
     */
    public debouncedUpdate(): void {
        if (this.debounceTimer) {
            clearTimeout(this.debounceTimer);
        }
        this.debounceTimer = window.setTimeout(() => {
            this.updateScrollbars();
            this.debounceTimer = null;
        }, this.options.debounceTime);
    }

    /**
     * 更新滚动条位置和尺寸
     */
    public updateScrollbars(): void {
        if (!this.options.data) return;
        const { track, thumb } = this.elements;
        const data = this.options.data;
        if (this.options.direction === 'horizontal') {
            const trackWidth = track.clientWidth;
            const totalWidth = data.colHeaders.allColWidth * data.zoom;
            const visibleWidth = data.visibleView.sheetWidth;
            if (totalWidth <= visibleWidth) {
                thumb.style.width = `${trackWidth}px`;
                thumb.style.left = '0px';
                return;
            }

            const ratio = visibleWidth / totalWidth;
            const thumbWidth = Math.max(this.options.thumbMinSize, trackWidth * ratio);
            thumb.style.width = `${thumbWidth}px`;

            const scrollPosition = (data.colHeaders.offsetWidth * data.zoom / totalWidth) * trackWidth;
            thumb.style.left = `${Math.max(0, Math.min(scrollPosition, trackWidth - thumbWidth))}px`;
        } else {
            const trackHeight = track.clientHeight;
            const totalHeight = data.rowHeaders.allRowHeight * data.zoom;
            const visibleHeight = data.visibleView.sheetHeight;

            if (totalHeight <= visibleHeight) {
                thumb.style.height = `${trackHeight}px`;
                thumb.style.top = '0px';
                return;
            }

            const ratio = visibleHeight / totalHeight;
            const thumbHeight = Math.max(this.options.thumbMinSize, trackHeight * ratio);
            thumb.style.height = `${thumbHeight}px`;

            const scrollPosition = (data.rowHeaders.offsetHeight * data.zoom / totalHeight) * trackHeight;
            thumb.style.top = `${Math.max(0, Math.min(scrollPosition, trackHeight - thumbHeight))}px`;
        }
    }

    /**
     * 更新滚动偏移量
     * @param {number} delta - 偏移量变化
     */
    private updateOffset(delta: number = 0): void {
        if (!this.options.data) return;

        const { track, thumb } = this.elements;
        const data = this.options.data;

        if (this.options.direction === 'horizontal') {
            const startScrollPosition = this.isDragging ? this.startScrollPosition : parseFloat(thumb.style.left) || 0;
            const maxDelta = track.clientWidth - thumb.clientWidth - startScrollPosition;
            const clampedDelta = Math.min(Math.max(delta, -startScrollPosition), maxDelta);

            const totalWidth = data.colHeaders.allColWidth * data.zoom;
            // thumb 在可用滑动空间（track - thumb）中的比例 × 内容可滚动范围（total - viewport）
            // 确保 thumb 拖到最右时 offset 正好到达 maxOffset（旧公式除以 track.clientWidth
            // 会导致比例 < 100%，无法滚到最右/最下）
            const scrollableTrack = track.clientWidth - thumb.clientWidth;
            const scrollableContent = totalWidth - data.visibleView.sheetWidth;
            const thumbPos = startScrollPosition + clampedDelta;
            const newOffset = scrollableTrack > 0 ? (thumbPos / scrollableTrack) * scrollableContent : 0;
            data.colHeaders.offsetWidth = Math.max(0, Math.min(newOffset, scrollableContent) / data.zoom);
        } else {
            const startScrollPosition = this.isDragging ? this.startScrollPosition : parseFloat(thumb.style.top) || 0;
            const maxDelta = track.clientHeight - thumb.clientHeight - startScrollPosition;
            const clampedDelta = Math.min(Math.max(delta, -startScrollPosition), maxDelta);

            const totalHeight = data.rowHeaders.allRowHeight * data.zoom;
            // 同水平方向：按可用滑动空间比例映射，确保 thumb 到底时到达 maxOffset
            const scrollableTrack = track.clientHeight - thumb.clientHeight;
            const scrollableContent = totalHeight - data.visibleView.sheetHeight;
            const thumbPos = startScrollPosition + clampedDelta;
            const newOffset = scrollableTrack > 0 ? (thumbPos / scrollableTrack) * scrollableContent : 0;
            data.rowHeaders.offsetHeight = Math.max(0, Math.min(newOffset, scrollableContent) / data.zoom);
        }
    }

    /**
     * 开始拖拽
     * @param {MouseEvent} e - 鼠标事件
     */
    private startDrag(e: MouseEvent): void {
        if (e.button !== 0 || e.target !== this.elements.thumb) return;
        e.preventDefault();

        const { thumb } = this.elements;
        this.isDragging = true;
        this.startScrollPosition = this.options.direction === 'horizontal'
            ? parseFloat(thumb.style.left) || 0
            : parseFloat(thumb.style.top) || 0;
        this.startDragPosition = this.options.direction === 'horizontal' ? e.clientX : e.clientY;
        document.body.style.userSelect = 'none';
        document.body.style.cursor = 'grabbing';
    }

    /**
     * 处理拖拽
     * @param {MouseEvent} e - 鼠标事件
     */
    private handleDrag(e: MouseEvent): void {
        if (!this.isDragging) return;
        e.preventDefault();

        if (this.rafId) {
            cancelAnimationFrame(this.rafId);
        }
        this.rafId = requestAnimationFrame(() => {
            const delta = this.options.direction === 'horizontal'
                ? e.clientX - this.startDragPosition
                : e.clientY - this.startDragPosition;
            this.updateOffset(delta);
            this.updateScrollbars();
            this.rafId = null;
        });
    }

    /**
     * 处理轨道点击
     * @param {MouseEvent} e - 鼠标事件
     */
    private handleTrackClick(e: MouseEvent): void {
        if (e.button !== 0 || e.target !== this.elements.track) return;
        e.preventDefault();

        const { thumb, track } = this.elements;
        const rect = thumb.getBoundingClientRect();
        const trackRect = track.getBoundingClientRect();

        const clickPos = this.options.direction === 'horizontal' ? e.clientX : e.clientY;
        const thumbPos = this.options.direction === 'horizontal' ? rect.left : rect.top;
        const trackPos = this.options.direction === 'horizontal' ? trackRect.left : trackRect.top;
        const thumbSize = this.options.direction === 'horizontal' ? rect.width : rect.height;

        let delta: number;
        if (clickPos < thumbPos) {
            delta = -thumbSize;
        } else if (clickPos > thumbPos + thumbSize) {
            delta = thumbSize;
        } else {
            return;
        }

        this.updateOffset(delta);
        this.updateScrollbars();
    }

    /**
     * 处理触摸开始
     * @param {TouchEvent} e - 触摸事件
     */
    private handleTouchStart(e: TouchEvent): void {
        if (this.lastTouchIdentifier !== null) return;
        e.preventDefault();

        const touch = e.touches[0];
        this.lastTouchIdentifier = touch.identifier;
        this.isDragging = true;
        this.startScrollPosition = this.options.direction === 'horizontal'
            ? parseFloat(this.elements.thumb.style.left) || 0
            : parseFloat(this.elements.thumb.style.top) || 0;
        this.startDragPosition = this.options.direction === 'horizontal' ? touch.clientX : touch.clientY;
    }

    /**
     * 处理触摸移动
     * @param {TouchEvent} e - 触摸事件
     */
    private handleTouchMove(e: TouchEvent): void {
        if (!this.isDragging || this.lastTouchIdentifier === null) return;
        e.preventDefault();

        const touch = Array.from(e.touches).find(t => t.identifier === this.lastTouchIdentifier);
        if (!touch) return;

        if (this.rafId) {
            cancelAnimationFrame(this.rafId);
        }
        this.rafId = requestAnimationFrame(() => {
            const delta = this.options.direction === 'horizontal'
                ? touch.clientX - this.startDragPosition
                : touch.clientY - this.startDragPosition;
            this.updateOffset(delta);
            this.updateScrollbars();
            this.rafId = null;
        });
    }

    /**
     * 处理轨道触摸
     * @param {TouchEvent} e - 触摸事件
     */
    private handleTrackTouch(e: TouchEvent): void {
        if (e.target !== this.elements.track || this.lastTouchIdentifier !== null) return;
        e.preventDefault();

        const touch = e.touches[0];
        this.lastTouchIdentifier = touch.identifier;

        const { thumb, track } = this.elements;
        const rect = thumb.getBoundingClientRect();
        const trackRect = track.getBoundingClientRect();

        const touchPos = this.options.direction === 'horizontal' ? touch.clientX : touch.clientY;
        const thumbPos = this.options.direction === 'horizontal' ? rect.left : rect.top;
        const trackPos = this.options.direction === 'horizontal' ? trackRect.left : trackRect.top;
        const thumbSize = this.options.direction === 'horizontal' ? rect.width : rect.height;

        let delta: number;
        if (touchPos < thumbPos) {
            delta = -thumbSize;
        } else if (touchPos > thumbPos + thumbSize) {
            delta = thumbSize;
        } else {
            return;
        }

        this.updateOffset(delta);
        this.updateScrollbars();
    }

    /**
     * 开始步进
     * @param {MouseEvent | TouchEvent} e - 鼠标或触摸事件
     * @param {'back' | 'forward'} direction - 步进方向
     */
    private startStep(e: MouseEvent | TouchEvent, direction: 'back' | 'forward'): void {
        if (e instanceof MouseEvent && e.button !== 0) return;
        e.preventDefault();

        this.stepDirection = direction;
        this.executeStep();

        this.stepHoldTimer = window.setTimeout(() => {
            this.stepIntervalId = window.setInterval(() => {
                this.executeStep();
            }, this.STEP_INTERVAL);
        }, this.STEP_DELAY);
    }

    /**
     * 执行步进
     */
    private executeStep(): void {
        if (!this.stepDirection || !this.options.data) return;

        const data = this.options.data;
        if (this.options.direction === 'horizontal') {
            const { startCol } = data.colHeaders.getVisibleStartCol();
            if (this.stepDirection === 'back') {
                const stepSize = data.colHeaders.getAt(startCol - 1).left < data.colHeaders.offsetWidth
                 ? data.colHeaders.offsetWidth - data.colHeaders.getAt(startCol - 1).left 
                 : (startCol - 2 < 0 ? 0 : data.colHeaders.getAt(startCol - 2).width);
                data.colHeaders.offsetWidth = Math.max(0, data.colHeaders.offsetWidth - stepSize);
            } else {
                const stepSize = data.colHeaders.getAt(startCol - 1).left + data.colHeaders.getAt(startCol - 1).width > data.colHeaders.offsetWidth
                 ? data.colHeaders.getAt(startCol - 1).left + data.colHeaders.getAt(startCol - 1).width - data.colHeaders.offsetWidth
                 : data.colHeaders.getAt(startCol).width;
                if (data.colHeaders.offsetWidth + stepSize > data.colHeaders.allColWidth - data.visibleView.sheetWidth / data.zoom) {
                    const cols = Math.ceil(stepSize / DEFAULT_CELL_WIDTH);
                    data.colHeaders.addCol(data.colHeaders.length, data.colHeaders.length + cols);
                }
                const maxOffset = data.colHeaders.allColWidth - data.visibleView.sheetWidth / data.zoom;
                data.colHeaders.offsetWidth = Math.min(maxOffset, data.colHeaders.offsetWidth + stepSize);
            }
        } else {
            const { startRow } = data.rowHeaders.getVisibleStartRow();
            if (this.stepDirection === 'back') {
                const stepSize = data.rowHeaders.getAt(startRow - 1).top < data.rowHeaders.offsetHeight
                 ? data.rowHeaders.offsetHeight - data.rowHeaders.getAt(startRow - 1).top 
                 : (startRow - 2 < 0 ? 0 : data.rowHeaders.getAt(startRow - 2).height);
                data.rowHeaders.offsetHeight = Math.max(0, data.rowHeaders.offsetHeight - stepSize);
            } else {
                const stepSize = data.rowHeaders.getAt(startRow - 1).top + data.rowHeaders.getAt(startRow - 1).height > data.rowHeaders.offsetHeight
                 ? data.rowHeaders.getAt(startRow - 1).top + data.rowHeaders.getAt(startRow - 1).height - data.rowHeaders.offsetHeight
                 : data.rowHeaders.getAt(startRow).height;
                if (data.rowHeaders.offsetHeight + stepSize > data.rowHeaders.allRowHeight - data.visibleView.sheetHeight / data.zoom) {
                    const rows = Math.ceil(stepSize / DEFAULT_CELL_HEIGHT);
                    data.rowHeaders.addRow(data.rowHeaders.length, data.rowHeaders.length + rows);
                }
                const maxOffset = data.rowHeaders.allRowHeight - data.visibleView.sheetHeight / data.zoom;
                data.rowHeaders.offsetHeight = Math.min(maxOffset, data.rowHeaders.offsetHeight + stepSize);
            }
        }
        this.updateScrollbars();
    }

    /**
     * 停止步进
     */
    private stopStep(): void {
        if (this.stepHoldTimer) {
            clearTimeout(this.stepHoldTimer);
            this.stepHoldTimer = null;
        }
        if (this.stepIntervalId) {
            clearInterval(this.stepIntervalId);
            this.stepIntervalId = null;
        }
        this.stepDirection = null;
    }

    /**
     * 结束拖拽
     * @param {MouseEvent | TouchEvent} e - 鼠标或触摸事件
     */
    private endDrag(e: MouseEvent | TouchEvent): void {
        if (!this.isDragging) return;

        this.isDragging = false;
        this.startScrollPosition = 0;
        this.lastTouchIdentifier = null;

        document.body.style.cursor = '';
        document.body.style.userSelect = '';

        if (this.rafId) {
            cancelAnimationFrame(this.rafId);
            this.rafId = null;
        }
    }

    /**
     * 结束触摸
     * @param {TouchEvent} e - 触摸事件
     */
    private endTouch(e: TouchEvent): void {
        this.endDrag(e);
    }
    
    /**
     * 处理鼠标滚轮事件，实现像素级连续滚动。
     * 对 wheel 的高频事件做 rAF 合并，每帧只应用一次累积偏移，保证丝滑。
     * 同时处理 deltaMode 归一化（行/页 → 像素），兼容不同设备。
     * 垂直方向不自动扩行，滚动到最后一行即停止；
     * 水平方向到右侧边界时自动扩列，允许继续滚动。
     * @param e 鼠标滚轮事件对象
     */
    private handleWheel(e: WheelEvent) {
        if (!this.options.data) return;
        const data = this.options.data;

        // Ctrl + 滚轮：缩放主页面（由垂直滚动条实例统一处理，水平实例仅拦截默认行为）
        if (e.ctrlKey) {
            e.preventDefault(); // 阻止浏览器页面缩放
            if (this.options.direction !== 'vertical') return;
            const factor = e.deltaY < 0 ? Scroller.ZOOM_FACTOR : 1 / Scroller.ZOOM_FACTOR;
            const zoom = Math.min(Scroller.ZOOM_MAX, Math.max(Scroller.ZOOM_MIN, data.zoom * factor));
            data.zoom = Math.round(zoom * 100) / 100; // 保留两位小数，避免浮点累积漂移
            return;
        }

        // 根据 deltaMode 归一化到像素单位，兼容不同设备
        // deltaMode: 0=像素, 1=行, 2=页
        const rowHeight = DEFAULT_CELL_HEIGHT * data.zoom;
        const colWidth = DEFAULT_CELL_WIDTH * data.zoom;

        let deltaX = e.deltaX;
        let deltaY = e.deltaY;
        if (e.deltaMode === 1) {
            deltaY *= rowHeight;
            deltaX *= colWidth;
        } else if (e.deltaMode === 2) {
            deltaY *= data.visibleView.sheetHeight;
            deltaX *= data.visibleView.sheetWidth;
        }

        // 垂直滚动条处理 deltaY，水平滚动条处理 deltaX
        const delta = this.options.direction === 'vertical' ? deltaY : deltaX;
        if (delta === 0) return;

        e.preventDefault();

        // 累积偏移量，合并到 rAF 每帧只应用一次
        this.wheelTargetOffset += delta;
        if (this.wheelRafId !== null) return;
        this.wheelRafId = requestAnimationFrame(() => {
            this.applyWheelDelta();
            this.wheelRafId = null;
        });
    }

    /**
     * 应用累积的滚轮偏移量（由 rAF 调用，每帧一次）。
     * 支持垂直扩行 / 水平扩列，自动处理边界。
     */
    private applyWheelDelta(): void {
        if (!this.options.data) return;
        const data = this.options.data;

        if (this.options.direction === 'vertical') {
            const rows = data.rowHeaders;
            const visibleHeight = data.visibleView.sheetHeight / data.zoom;
            // 不自动扩行：滚轮滚动以下方已有内容的最后一行为下界
            const maxOffset = Math.max(0, rows.allRowHeight - visibleHeight);

            const newOffset = rows.offsetHeight + this.wheelTargetOffset;
            this.wheelTargetOffset = 0;

            rows.offsetHeight = Math.max(0, Math.min(newOffset, maxOffset));
        } else {
            const cols = data.colHeaders;
            const visibleWidth = data.visibleView.sheetWidth / data.zoom;
            const totalWidth = cols.allColWidth;
            let maxOffset = totalWidth - visibleWidth;
            if (maxOffset < 0) maxOffset = 0;

            let newOffset = cols.offsetWidth + this.wheelTargetOffset;
            this.wheelTargetOffset = 0;

            // 向右滚超出上界时，自动扩列以允许继续滚动
            if (newOffset > maxOffset) {
                const overflow = newOffset - maxOffset;
                const colsToAdd = Math.ceil(overflow / DEFAULT_CELL_WIDTH) + 2;
                cols.addCol(cols.length, cols.length + colsToAdd);
                maxOffset = cols.allColWidth - visibleWidth;
            }

            cols.offsetWidth = Math.max(0, Math.min(newOffset, maxOffset));
        }
    }

    /**
     * 获取容器元素
     * @returns {HTMLDivElement} 容器元素
     */
    public getContainer(): HTMLDivElement {
        return this.elements.container;
    }

    /**
     * 获取所有元素
     * @returns {ScrollbarElements} 元素对象
     */
    public getElements(): ScrollbarElements {
        return this.elements;
    }

    /**
     * 销毁滚动条组件
     *
     * 清理顺序：动画帧/定时器 → 尺寸监听 → document 级监听 → 滚轮监听 →
     * 数据事件监听 → DOM 移除 → 事件总线清空。
     * document 级与数据级监听均使用与注册时相同的具名引用，确保真正解除，
     * 避免旧实现传入新箭头函数导致 removeEventListener 失效（内存泄漏）。
     */
    public destroy(): void {
        // 1. 停止进行中的动画帧与定时器
        if (this.rafId) {
            cancelAnimationFrame(this.rafId);
            this.rafId = null;
        }
        if (this.wheelRafId !== null) {
            cancelAnimationFrame(this.wheelRafId);
            this.wheelRafId = null;
        }
        if (this.debounceTimer) {
            clearTimeout(this.debounceTimer);
            this.debounceTimer = null;
        }
        this.stopStep();

        // 2. 尺寸监听
        this.resizeObserver.disconnect();

        // 3. document 级监听（与注册时同一引用）
        document.removeEventListener('mousemove', this.onDocumentMouseMove);
        document.removeEventListener('mouseup', this.onDocumentMouseUp);
        document.removeEventListener('touchmove', this.onDocumentTouchMove);
        document.removeEventListener('touchend', this.onDocumentTouchEnd);
        document.removeEventListener('touchcancel', this.onDocumentTouchCancel);

        // 4. 滚轮监听
        if (this.wheelListenerTarget) {
            this.wheelListenerTarget.removeEventListener('wheel', this.onWheelEvent);
            this.wheelListenerTarget = null;
        }

        // 5. 数据事件监听
        if (this.options.data) {
            this.options.data.colHeaders.off(DataEvents.OFFSETX_CHANGED, this.onOffsetXChanged);
            this.options.data.rowHeaders.off(DataEvents.OFFSETY_CHANGED, this.onOffsetYChanged);
            this.options.data.colHeaders.off(DataEvents.ALL_COL_WIDTH_CHANGED, this.onAllColWidthChanged);
            this.options.data.rowHeaders.off(DataEvents.ALL_ROW_HEIGHT_CHANGED, this.onAllRowHeightChanged);
            this.options.data.off(DataEvents.ZOOM_CHANGED, this.onZoomChanged);
        }

        // 6. 移除 DOM
        if (this.elements.container.parentElement) {
            this.elements.container.parentElement.removeChild(this.elements.container);
        }

        // 7. 事件总线清空
        this.removeAllListeners();
    }
}