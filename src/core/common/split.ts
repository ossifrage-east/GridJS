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
        rememberSize: true
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
                parsedPosition < (minSize as MinSize).prev ||
                parsedPosition > containerSize - (minSize as MinSize).next) {
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
        const constrainedPosition = this.constrainPosition(position, containerSize);

        this.applyLayoutStyles(constrainedPosition, containerSize);
        this.currentPosition = constrainedPosition;

        if (rememberSize) {
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
