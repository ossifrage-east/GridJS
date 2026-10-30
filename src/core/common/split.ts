/**
 * Split 面板分割组件
 * 支持水平和垂直两种分割方式，可拖拽调整面板大小，支持记忆布局状态
 * 使用 IndexDB 持久化存储位置数据
 * 
 * @class Split
 * @example
 * // 创建水平分割面板
 * const split = new Split({
 *     parentElement: document.getElementById('app'),
 *     horizontal: true,
 *     minSize: { prev: 200, next: 300 }
 * });
 * const { splitPrev, splitNext } = split.getElements();
 * splitPrev.appendChild(leftPanel);
 * splitNext.appendChild(rightPanel);
 */

import '../style/split.scss';
import { createDiv } from '../../utils/dom';
import { debounce } from '../../utils/debounce';
import { IndexDB } from '../dataArchitecture/indexDB';

/**
 * 最小尺寸配置接口
 * @interface MinSize
 */
interface MinSize {
    /** 前侧面板最小尺寸（像素） */
    prev: number;
    /** 后侧面板最小尺寸（像素） */
    next: number;
}

/**
 * 组件配置选项接口
 * @interface SplitOptions
 */
interface SplitOptions {
    /** 父元素，默认值为 document.body */
    parentElement?: HTMLElement;
    /** 主容器的样式ID */
    id?: string;
    /** 是否水平布局，默认值为 true */
    horizontal?: boolean;
    /** 最小面板尺寸，支持数字（统一设置）或对象（分别设置） */
    minSize?: number | MinSize;
    /** 最大面板尺寸，默认值为 Infinity */
    maxSize?: number;
    /**
     * 后侧（右侧 / 下侧）面板的默认尺寸（像素）
     * 设置后，无历史记录时按该尺寸分配后侧面板，前侧面板自动占据剩余空间；
     * 未设置（默认 0）时前后两侧平分容器
     */
    nextSize?: number;
    /** 存储键名，用于记忆面板位置 */
    storageKey?: string;
    /** 是否记住面板大小，默认值为 true */
    rememberSize?: boolean;
    /**
     * 拖拽结束回调（参数为后侧面板当前尺寸，像素）
     *
     * 仅在用户**实际拖动分割条并松开**后触发，用于把导航内容区宽度按导航项持久化。
     * 与 {@link SplitOptions.onResize} 的区别：后者在容器尺寸变化 / 程序化布局时也会触发。
     */
    onDragEnd?: (nextSize: number) => void;
}

/**
 * DOM元素类型定义
 * @typedef {Object} SplitElements
 */
type SplitElements = {
    /** 主容器元素 */
    container: HTMLDivElement;
    /** 前侧面板元素 */
    splitPrev: HTMLDivElement;
    /** 分割条元素 */
    split: HTMLDivElement;
    /** 后侧面板元素 */
    splitNext: HTMLDivElement;
};

/**
 * 判断元素是否可见的辅助函数
 * @param {HTMLElement} element - 要检查的元素
 * @returns {boolean} 是否可见
 */
function isElementVisible(element: HTMLElement): boolean {
    return element.hidden === false;
}

/**
 * 获取容器尺寸的辅助函数
 * @param {HTMLElement} container - 容器元素
 * @param {boolean} horizontal - 是否水平布局
 * @returns {number} 容器尺寸（宽度或高度）
 */
function getContainerSize(container: HTMLElement, horizontal: boolean): number {
    return horizontal ? container.clientWidth : container.clientHeight;
}

export class Split {
    /** 组件配置选项 */
    private readonly options: Required<SplitOptions>;
    /** DOM元素集合 */
    private readonly elements: SplitElements;
    /** 是否正在拖拽 */
    private isDragging = false;
    /** 拖拽起始位置 */
    private startPosition = 0;
    /** 当前分割位置 */
    private currentPosition = 0;
    /** 容器大小变化观察器 */
    private resizeObserver!: ResizeObserver;
    /** DOM变化观察器 */
    private mutationObserver?: MutationObserver;
    /** IndexDB 实例 */
    private readonly db: IndexDB;
    /** 布局是否已初始化完成（首次同步布局已应用） */
    private layoutInitialized = false;
    /** 布局加载完成的 Promise（从 IndexDB 读取位置完成） */
    public positionLoaded: Promise<void>;
    private _positionLoadedResolve!: () => void;
    /** 是否已尝试过从 IndexDB 恢复位置（未恢复前不把默认位置写回库，避免覆盖存档） */
    private positionRestored: boolean = false;
    /**
     * 待应用的分割位置
     *
     * 容器隐藏（尺寸为 0）时量不到尺寸，读回的存档位置先挂在这里；
     * 容器重新可见后由 {@link refreshLayout}（或 ResizeObserver）应用。
     */
    private pendingPosition: number | null = null;

    /**
     * 默认配置选项
     * @static
     * @readonly
     */
    private static readonly defaultOptions: Required<SplitOptions> = {
        parentElement: document.body,
        id: 'split-container',
        horizontal: true,
        minSize: { prev: 300, next: 300 },
        maxSize: Infinity,
        nextSize: 0,
        storageKey: 'split-position',
        rememberSize: true,
        onDragEnd: undefined
    };

    /**
     * 构造函数
     * @param {SplitOptions} [options={}] - 组件配置选项
     */
    constructor(options: SplitOptions = {}) {
        this.options = this.mergeOptions(options);
        this.elements = this.createElements();
        this.db = new IndexDB();
        this.positionLoaded = new Promise(resolve => { this._positionLoadedResolve = resolve; });
        this.setupEventListeners();
        this.setupResizeObserver();
        this.restorePosition();
    }

    /**
     * 合并用户配置和默认配置
     * @private
     * @param {SplitOptions} options - 用户配置
     * @returns {Required<SplitOptions>} 合并后的配置
     */
    private mergeOptions(options: SplitOptions): Required<SplitOptions> {
        const processedMinSize = typeof options.minSize === 'number'
            ? { prev: options.minSize, next: options.minSize }
            : options.minSize || Split.defaultOptions.minSize;

        return { ...Split.defaultOptions, ...options, minSize: processedMinSize };
    }

    /**
     * 创建DOM元素结构
     * @private
     * @returns {SplitElements} DOM元素集合
     */
    private createElements(): SplitElements {
        const { id, horizontal, parentElement } = this.options;
        const container = createDiv({ id: `${id}`, className: horizontal ? 'horizontal' : 'vertical' });
        const splitPrev = createDiv({ className: 'split-pre' });
        const split = createDiv({ className: 'split' });
        const splitNext = createDiv({ className: 'split-next' });

        container.append(splitPrev, split, splitNext);
        parentElement.append(container);

        return { container, splitPrev, split, splitNext };
    }

    /**
     * 恢复保存的面板位置
     * 容器尺寸就绪后先同步应用中间布局（保证 ResizeObserver/MutationObserver 拿到有效值），
     * 然后异步从 IndexDB 读取保存位置并覆盖（若无则保持中间位置）
     * @private
     */
    private restorePosition(): void {
        const { container } = this.elements;

        if (!getContainerSize(container, this.options.horizontal)) {
            requestAnimationFrame(() => this.restorePosition());
            return;
        }

        // 1. 同步应用默认布局，确保 currentPosition 与面板尺寸立即可用
        this.updateLayout(this.calculateDefaultPosition());
        this.layoutInitialized = true;

        // 2. 异步加载 IndexDB 中的保存位置并覆盖
        this.loadPositionFromDB().then(savedPosition => {
            if (savedPosition !== null) {
                this.updateLayout(savedPosition);
            }
        }).catch(() => {
            // 加载失败则保持默认位置不变
        }).finally(() => {
            // 读完（无论有没有存档）之后才允许把布局位置写回库：
            // 读回是异步的，构造时先写默认位置会把旧存档覆盖掉
            this.positionRestored = true;
            this._positionLoadedResolve();
        });
    }

    /**
     * 从 IndexDB 加载位置数据
     * @private
     * @returns {Promise<number | null>} 保存的位置，无效返回 null
     */
    private async loadPositionFromDB(): Promise<number | null> {
        const { storageKey, rememberSize, minSize } = this.options;

        if (!rememberSize) return null;

        try {
            const savedPosition = await this.db.getSetting(storageKey);
            if (savedPosition === null) return null;

            const parsedPosition = typeof savedPosition === 'number' ? savedPosition : parseInt(String(savedPosition), 10);
            const containerSize = getContainerSize(this.elements.container, this.options.horizontal);

            if (isNaN(parsedPosition) ||
                parsedPosition < (minSize as MinSize).prev) {
                return null;
            }
            // 容器隐藏（尺寸为 0）时算不出上界：先收下，等重新可见后再用真实尺寸收敛
            if (containerSize > 0 && parsedPosition > containerSize - (minSize as MinSize).next) {
                return null;
            }

            return parsedPosition;
        } catch {
            return null;
        }
    }

    /**
     * 计算默认位置
     * - 未设置 nextSize（0）时：前后两侧平分容器（中间位置）
     * - 设置了 nextSize 时：后侧（右侧 / 下侧）面板取该默认尺寸，前侧占据剩余空间
     * 结果均受最小 / 最大尺寸约束
     * @private
     * @returns {number} 默认位置（满足最小尺寸限制）
     */
    private calculateDefaultPosition(): number {
        const { minSize, nextSize } = this.options;
        const containerSize = getContainerSize(this.elements.container, this.options.horizontal);

        const preferredPosition = nextSize > 0 ? containerSize - nextSize : containerSize / 2;

        return Math.max(
            (minSize as MinSize).prev,
            Math.min(containerSize - (minSize as MinSize).next, preferredPosition)
        );
    }

    /**
     * 更新面板布局
     * 根据位置约束更新前后面板尺寸，并保存位置（如果允许）
     * @private
     * @param {number} position - 分割位置（像素）
     */
    private updateLayout(position: number): void {
        const { horizontal, minSize, maxSize, rememberSize, storageKey } = this.options;
        const { container, splitPrev, split, splitNext } = this.elements;

        const containerSize = getContainerSize(container, horizontal);
        // 容器隐藏（尺寸为 0）时不做布局：异步读回保存位置可能恰好发生在隐藏期间，
        // 此时按 0 计算会把分割位置压成最小值，切回来就恢复不了。
        // 先把目标位置挂起，容器重新可见时由 refreshLayout 应用
        if (!containerSize) {
            this.pendingPosition = position;
            return;
        }
        this.pendingPosition = null;
        const constrainedPosition = this.constrainPosition(position, containerSize);

        this.applyLayoutStyles(constrainedPosition, containerSize);
        this.currentPosition = constrainedPosition;

        if (rememberSize && this.positionRestored) {
            this.db.setSetting(storageKey, constrainedPosition, 'split-position').catch(() => {});
        }
    }

    /**
     * 约束位置在有效范围内
     * @private
     * @param {number} position - 原始位置
     * @param {number} containerSize - 容器尺寸
     * @returns {number} 约束后的位置
     */
    private constrainPosition(position: number, containerSize: number): number {
        const { minSize, maxSize } = this.options;
        return Math.min(maxSize, Math.max((minSize as MinSize).prev, Math.min(containerSize - (minSize as MinSize).next, position)));
    }

    /**
     * 应用布局样式
     * @private
     * @param {number} position - 分割位置
     * @param {number} containerSize - 容器尺寸
     */
    private applyLayoutStyles(position: number, containerSize: number): void {
        const { horizontal } = this.options;
        const { splitPrev, split, splitNext } = this.elements;

        if (horizontal) {
            const nextWidth = containerSize - position;
            splitPrev.style.width = `${position}px`;
            split.style.left = `${position - 2}px`;
            splitNext.style.marginLeft = `${position}px`;
            splitNext.style.width = `${nextWidth}px`;
        } else {
            const nextHeight = containerSize - position;
            splitPrev.style.height = `${position}px`;
            split.style.top = `${position}px`;
            splitNext.style.height = `${nextHeight}px`;
        }
    }

    /**
     * 获取后侧面板（导航内容区）的当前尺寸
     *
     * 由 {@link currentPosition} 与容器尺寸推算，而非读取 splitNext.clientWidth：
     * 后者依赖实际布局（容器尚未完成布局时为 0，且每次读取都会触发强制重排）
     * @returns {number} 后侧面板尺寸（像素）
     */
    public getNextSize(): number {
        const { horizontal } = this.options;
        const containerSize = getContainerSize(this.elements.container, horizontal);
        return Math.max(0, containerSize - this.currentPosition);
    }

    /**
     * 设置后侧面板（导航内容区）的目标尺寸
     *
     * 与 {@link setPosition}（以分割位置为参数）不同，此处以「后侧面板宽度」为参数，
     * 便于按导航项的初始化宽度 / 已保存宽度直接设置。
     * 实际生效尺寸会被最小 / 最大尺寸约束与容器尺寸收敛。
     * @param {number} size - 期望的后侧面板尺寸（像素）
     */
    public setNextSize(size: number): void {
        const { container } = this.elements;
        const { horizontal } = this.options;
        const containerSize = getContainerSize(container, horizontal);
        if (!containerSize) return;
        // 分割位置 = 容器尺寸 - 后侧尺寸（两种方向同构）
        this.updateLayout(containerSize - size);
    }

    /**
     * 设置事件监听器
     * @private
     */
    private setupEventListeners(): void {
        const { split } = this.elements;

        split.addEventListener('mousedown', this.handleMouseDown);
        document.addEventListener('mousemove', this.handleMouseMove, { passive: true });
        document.addEventListener('mouseup', this.handleMouseUp);
    }

    /**
     * 处理鼠标按下事件
     * @private
     * @param {MouseEvent} e - 鼠标事件
     */
    private handleMouseDown = (e: MouseEvent): void => {
        if (e.button !== 0 || e.target !== this.elements.split) return;
        e.preventDefault();

        this.isDragging = true;
        this.startPosition = this.options.horizontal ? e.clientX : e.clientY;

        document.body.style.userSelect = 'none';
        document.body.style.cursor = this.options.horizontal ? 'col-resize' : 'row-resize';
    };

    /**
     * 处理鼠标移动事件
     * @private
     * @param {MouseEvent} e - 鼠标事件
     */
    private handleMouseMove = (e: MouseEvent): void => {
        if (!this.isDragging) return;

        const current = this.options.horizontal ? e.clientX : e.clientY;
        const delta = current - this.startPosition;

        this.updateLayout(this.currentPosition + delta);
        this.startPosition = current;
    };

    /**
     * 处理鼠标释放事件
     * @private
     */
    private handleMouseUp = (): void => {
        if (!this.isDragging) return;

        this.isDragging = false;
        document.body.style.userSelect = '';
        document.body.style.cursor = '';
        // 真实拖拽结束：回调当前后侧尺寸，供调用方持久化「用户拖出来的宽度」
        this.options.onDragEnd?.(this.getNextSize());
    };

    /**
     * 设置容器大小变化观察器
     * @private
     */
    private setupResizeObserver(): void {
        const debounceResize = debounce(() => {
            if (!this.isDragging) {
                this.handleContainerResize();
            }
        }, 100);

        this.resizeObserver = new ResizeObserver(debounceResize);
        this.resizeObserver.observe(this.elements.container);
    }

    /**
     * 处理容器大小变化
     * 保持面板比例，同时满足最小尺寸限制
     * @private
     */
    private handleContainerResize(): void {
        // 布局未初始化前跳过，避免 NaN/0 计算错误
        if (!this.layoutInitialized) return;

        const { horizontal, minSize } = this.options;
        const { container, splitNext } = this.elements;

        const containerSize = getContainerSize(container, horizontal);
        // 容器被隐藏（display:none，尺寸为 0）时不要重排：否则会把分割位置算成
        // 最小尺寸、把后侧面板压成 0 宽，切回来时又因「旧总长为 0」判定为无效而永远恢复不了
        if (!containerSize) return;

        // 隐藏期间读回 / 设置过的位置先落地：容器重新可见时用它，而不是按比例重算
        if (this.pendingPosition !== null) {
            const pending = this.pendingPosition;
            this.pendingPosition = null;
            this.updateLayout(pending);
            return;
        }

        const oldNextSize = parseInt(horizontal
            ? splitNext.style.width || '0'
            : splitNext.style.height || '0', 10);
        const oldContainerSize = this.currentPosition + oldNextSize;

        if (oldContainerSize === containerSize) return;
        if (oldContainerSize <= 0) return;

        const newPosition = this.constrainPosition(
            (this.currentPosition / oldContainerSize) * containerSize,
            containerSize
        );

        this.updateLayout(newPosition);
    }

    /**
     * 重新应用布局
     *
     * 容器从隐藏（尺寸为 0）重新可见后调用：把隐藏期间挂起的 {@link pendingPosition}
     * 立刻落地，不必等 ResizeObserver 的防抖回调。
     */
    public refreshLayout(): void {
        this.handleContainerResize();
    }

    /**
     * 处理子元素可见性变化
     * 根据子元素的可见状态自动调整面板布局
     */
    public handleVisibility(): void {
        if (this.mutationObserver) {
            this.mutationObserver.disconnect();
        }

        const { horizontal } = this.options;
        const { splitPrev, split, splitNext } = this.elements;

        this.mutationObserver = new MutationObserver(() => {
            const prevVisible = Array.from(splitPrev.children).some((element) => (element as HTMLElement).hidden === false);
            const nextVisible = Array.from(splitNext.children).some((element) => (element as HTMLElement).hidden === false);

            this.adjustLayoutByVisibility(prevVisible, nextVisible);
        });

        this.mutationObserver.observe(splitNext, {
            attributes: true,
            childList: true,
            subtree: true,
            characterData: true,
        });
    }

    /**
     * 根据子元素可见性调整布局
     * @private
     * @param {boolean} prevVisible - 前侧面板是否有可见子元素
     * @param {boolean} nextVisible - 后侧面板是否有可见子元素
     */
    private adjustLayoutByVisibility(prevVisible: boolean, nextVisible: boolean): void {
        const { horizontal } = this.options;
        const { split, splitPrev, splitNext } = this.elements;

        if (prevVisible && nextVisible) {
            split.hidden = false;
            const dimension = horizontal ? 'width' : 'height';
            splitPrev.style[dimension] = `${this.layoutInitialized ? this.currentPosition : this.calculateDefaultPosition()}px`;
        } else {
            split.hidden = true;
            if (!nextVisible) {
                const dimension = horizontal ? 'width' : 'height';
                splitPrev.style[dimension] = '100%';
            }
            if (!prevVisible) {
                const dimension = horizontal ? 'width' : 'height';
                splitNext.style[dimension] = '100%';
            }
        }
    }

    /**
     * 设置分割位置
     * @param {number} position - 新的分割位置（像素）
     */
    public setPosition(position: number): void {
        this.updateLayout(position);
    }

    /**
     * 获取DOM元素集合
     * @returns {SplitElements} 包含 container、splitPrev、split、splitNext 的对象副本
     */
    public getElements(): SplitElements {
        return { ...this.elements };
    }

    /**
     * 销毁组件，清理资源
     * 移除事件监听、断开观察器、删除DOM元素
     */
    public destroy(): void {
        const { split } = this.elements;
        split.removeEventListener('mousedown', this.handleMouseDown);
        document.removeEventListener('mousemove', this.handleMouseMove);
        document.removeEventListener('mouseup', this.handleMouseUp);

        this.mutationObserver?.disconnect();
        this.resizeObserver.disconnect();

        this.elements.container.remove();
    }
}
