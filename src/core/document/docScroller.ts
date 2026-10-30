/*  # 文档域 · 滚动条 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目目标: 让文档正文区拥有与表格**同一副长相、同一套手势**的横 / 纵滚动条
    #           （两端箭头 + 可拖拽滑块 + 点轨道跳转 + 按住连续滚动）。
    # 本文件职责: 只做「一个滚动容器 ↔ 一条自定义滚动条」的双向绑定。
    #             **不认识块树、不认识纸张、不认识分页** —— 它读的只有
    #             scrollWidth/scrollHeight 与 scrollLeft/scrollTop 四个数。
    # 为什么另写一个而不是复用 Scroller: common/scroller 的滚动条是**表格专用**的：
    #             它的滑块位置由 `data.colHeaders.offsetWidth * data.zoom` 反算，
    #             拖动写回的也是表格的偏移量 —— 那套模型里根本没有「滚动容器」。
    #             文档是一条 DOM 流，天生就有 scrollLeft/Top，套表格那套反而要伪造
    #             一个 DataCollection。这里只复用它的**样式类名**（SCROLLER_CLASS_NAMES
    #             + style/scroller.scss），因此两边外观与手感一致而互不依赖。
    # 依赖方向: utils/dom + core/constant（类名常量）+ style/scroller.scss。
    #           **不 import dataArchitecture**，也不 import common/scroller。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { createDiv } from '../../utils/dom';
import { SCROLLER_CLASS_NAMES } from '../constant';
import '../style/scroller.scss';

/** 滚动条配置 */
export interface DocScrollbarOptions {
    /** 挂载点（表格里是 rightTools / bottomTools 的 split-next；文档里同类容器） */
    parentElement: HTMLElement;
    /** 被驱动的滚动容器（文档的 `.doc-scroller`） */
    target: HTMLElement;
    /** 方向 */
    direction: 'horizontal' | 'vertical';
}

/** 单端点箭头点一下滚多少像素 */
const STEP_PX = 60;
/** 按住箭头后，连续滚动的起始延时与间隔（与表格滚动条同一手感） */
const STEP_DELAY = 300;
const STEP_INTERVAL = 50;
/** 滑块最小尺寸（内容超长时不至于缩成一点，拖不动） */
const THUMB_MIN_SIZE = 30;

/**
 * 文档滚动条
 *
 * 与 {@link Scroller} 的区别只在**数据来源**：那条读 DataCollection 的偏移量，
 * 这条读真实 DOM 容器的 scrollLeft / scrollTop。因此它不需要任何「滚动事件 → 重绘」
 * 的中间层 —— 容器自己会滚，滚动条只是把位置画回来、把拖拽翻译过去。
 */
export class DocScrollbar {
    /** 配置 */
    private readonly options: DocScrollbarOptions;

    /** 外层容器（类名与表格滚动条同源） */
    private readonly container: HTMLDivElement;

    /** 轨道 */
    private readonly track: HTMLDivElement;

    /** 起点箭头 */
    private readonly back: HTMLDivElement;

    /** 滑块 */
    private readonly thumb: HTMLDivElement;

    /** 终点箭头 */
    private readonly forward: HTMLDivElement;

    /** 是否正在拖拽滑块 */
    private dragging: boolean = false;

    /** 拖拽起始时的指针坐标 */
    private startPointer: number = 0;

    /** 拖拽起始时滑块距轨道起点的距离 */
    private startThumbPos: number = 0;

    /** 端点箭头的连发定时器 */
    private stepTimer: number | null = null;

    /** 端点箭头的首次延时定时器 */
    private stepHoldTimer: number | null = null;

    /** 尺寸变化观察器（容器尺寸 / 内容长度变化都要重算滑块） */
    private observer: ResizeObserver | null = null;

    /** 观察的目标（换 target 时要先 unobserve 旧的） */
    private observed: HTMLElement | null = null;

    // —— 具名处理器：注册与移除必须是同一个引用，否则 destroy 解不掉 ——
    private readonly onDocMouseMove = (e: MouseEvent): void => this.handleDrag(e);
    private readonly onDocMouseUp = (): void => this.stopStep();
    private readonly onTargetScroll = (): void => this.update();

    /**
     * 构造函数
     * @param options - 见 {@link DocScrollbarOptions}
     */
    constructor(options: DocScrollbarOptions) {
        this.options = options;
        this.container = createDiv({ className: `${SCROLLER_CLASS_NAMES.CONTAINER}${options.direction}` });
        this.track = createDiv({ className: SCROLLER_CLASS_NAMES.TRACK });
        this.back = createDiv({ className: SCROLLER_CLASS_NAMES.BACK });
        this.thumb = createDiv({ className: SCROLLER_CLASS_NAMES.THUMB });
        this.forward = createDiv({ className: SCROLLER_CLASS_NAMES.FORWARD });
        this.track.appendChild(this.thumb);
        this.container.append(this.track, this.back, this.forward);
        options.parentElement.appendChild(this.container);

        this.bindEvents();
        this.observe();
        this.update();
    }

    /** 取外层容器（供工作区显隐） */
    public getElements(): { container: HTMLDivElement } {
        return { container: this.container };
    }

    /**
     * 换一个被驱动的滚动容器
     *
     * 视图切换时调用：监听必须跟着换，否则旧容器尺寸变化仍会触发一次无意义的重算，
     * 而新容器的滚动 / 尺寸变化没人监听（滑块停在初始值）。
     * @param target - 新的滚动容器
     */
    public setTarget(target: HTMLElement): void {
        if (this.options.target === target) return;
        this.options.target.removeEventListener('scroll', this.onTargetScroll);
        this.options.target = target;
        target.addEventListener('scroll', this.onTargetScroll);
        this.observe();
        this.update();
    }

    /**
     * 重算滑块尺寸与位置
     *
     * 内容没超出视口时滑块**占满轨道**（与表格滚动条同一处理），
     * 而不是藏起来 —— 藏起来会让轨道看起来像坏了。
     */
    public update(): void {
        const horizontal = this.options.direction === 'horizontal';
        const target = this.options.target;
        const trackSize = horizontal ? this.track.clientWidth : this.track.clientHeight;
        const total = horizontal ? target.scrollWidth : target.scrollHeight;
        const visible = horizontal ? target.clientWidth : target.clientHeight;
        if (trackSize <= 0) return;

        const scrollableContent = total - visible;
        // 内容没溢出（或容器正隐藏，量到 0）：滑块铺满轨道、回到起点
        if (scrollableContent <= 0 || visible <= 0) {
            if (horizontal) {
                this.thumb.style.width = `${trackSize}px`;
                this.thumb.style.left = '0px';
            } else {
                this.thumb.style.height = `${trackSize}px`;
                this.thumb.style.top = '0px';
            }
            return;
        }

        const ratio = visible / total;
        const size = Math.max(THUMB_MIN_SIZE, trackSize * ratio);
        const scrollableTrack = Math.max(0, trackSize - size);
        const offset = horizontal ? target.scrollLeft : target.scrollTop;
        const pos = Math.max(0, Math.min(scrollableTrack, (offset / scrollableContent) * scrollableTrack));
        if (horizontal) {
            this.thumb.style.width = `${size}px`;
            this.thumb.style.left = `${pos}px`;
        } else {
            this.thumb.style.height = `${size}px`;
            this.thumb.style.top = `${pos}px`;
        }
    }

    /** 绑定事件 */
    private bindEvents(): void {
        this.thumb.addEventListener('mousedown', (e) => this.startDrag(e));
        this.track.addEventListener('mousedown', (e) => this.jumpTo(e));
        this.back.addEventListener('mousedown', (e) => this.startStep(e, -1));
        this.forward.addEventListener('mousedown', (e) => this.startStep(e, 1));
        // 端点箭头按住连发：松开 / 移出都要停，否则松手后仍在滚
        document.addEventListener('mouseup', this.onDocMouseUp);
        document.addEventListener('mousemove', this.onDocMouseMove);
        // 容器自己滚动时（滚轮、键盘、程序化滚动）把滑块画回去
        this.options.target.addEventListener('scroll', this.onTargetScroll);
    }

    /** 观察滚动容器与轨道的尺寸变化 */
    private observe(): void {
        if (this.observed) {
            this.observer?.unobserve(this.observed);
            this.observed = null;
        }
        if (typeof ResizeObserver === 'undefined') return;
        if (!this.observer) {
            this.observer = new ResizeObserver(() => this.update());
        }
        this.observed = this.options.target;
        this.observer.observe(this.options.target);
        this.observer.observe(this.track);
    }

    /** 开始拖拽滑块 */
    private startDrag(e: MouseEvent): void {
        if (e.button !== 0) return;
        e.preventDefault();
        e.stopPropagation();
        this.dragging = true;
        this.startPointer = this.options.direction === 'horizontal' ? e.clientX : e.clientY;
        this.startThumbPos = this.options.direction === 'horizontal'
            ? parseFloat(this.thumb.style.left) || 0
            : parseFloat(this.thumb.style.top) || 0;
        document.body.style.userSelect = 'none';
    }

    /** 拖拽中 */
    private handleDrag(e: MouseEvent): void {
        if (!this.dragging) return;
        e.preventDefault();
        const horizontal = this.options.direction === 'horizontal';
        const delta = (horizontal ? e.clientX : e.clientY) - this.startPointer;
        const trackSize = horizontal ? this.track.clientWidth : this.track.clientHeight;
        const thumbSize = horizontal ? this.thumb.clientWidth : this.thumb.clientHeight;
        const scrollableTrack = Math.max(0, trackSize - thumbSize);
        const target = this.options.target;
        const scrollableContent = horizontal
            ? target.scrollWidth - target.clientWidth
            : target.scrollHeight - target.clientHeight;
        if (scrollableTrack <= 0 || scrollableContent <= 0) return;
        const pos = Math.max(0, Math.min(scrollableTrack, this.startThumbPos + delta));
        const offset = (pos / scrollableTrack) * scrollableContent;
        if (horizontal) target.scrollLeft = offset;
        else target.scrollTop = offset;
    }

    /** 松开滑块 */
    private endDrag(): void {
        if (!this.dragging) return;
        this.dragging = false;
        document.body.style.userSelect = '';
    }

    /**
     * 点轨道跳转
     *
     * 点在滑块上不算（那是拖拽）；点在轨道空白处按「滑块中心落在点击处」翻页，
     * 与表格滚动条同一口径。
     */
    private jumpTo(e: MouseEvent): void {
        if (e.button !== 0 || e.target === this.thumb) return;
        e.preventDefault();
        const horizontal = this.options.direction === 'horizontal';
        const trackRect = this.track.getBoundingClientRect();
        const click = horizontal ? e.clientX - trackRect.left : e.clientY - trackRect.top;
        const trackSize = horizontal ? this.track.clientWidth : this.track.clientHeight;
        const thumbSize = horizontal ? this.thumb.clientWidth : this.thumb.clientHeight;
        const scrollableTrack = Math.max(0, trackSize - thumbSize);
        const target = this.options.target;
        const scrollableContent = horizontal
            ? target.scrollWidth - target.clientWidth
            : target.scrollHeight - target.clientHeight;
        if (scrollableTrack <= 0 || scrollableContent <= 0) return;
        const pos = Math.max(0, Math.min(scrollableTrack, click - thumbSize / 2));
        const offset = (pos / scrollableTrack) * scrollableContent;
        if (horizontal) target.scrollLeft = offset;
        else target.scrollTop = offset;
    }

    /** 按住端点箭头：先延时，再连发 */
    private startStep(e: MouseEvent, direction: number): void {
        if (e.button !== 0) return;
        e.preventDefault();
        e.stopPropagation();
        this.stepBy(direction);
        this.stopStep();
        this.stepHoldTimer = window.setTimeout(() => {
            this.stepTimer = window.setInterval(() => this.stepBy(direction), STEP_INTERVAL);
        }, STEP_DELAY);
    }

    /** 结束端点箭头的连发 */
    private stopStep(): void {
        this.endDrag();
        if (this.stepHoldTimer !== null) {
            window.clearTimeout(this.stepHoldTimer);
            this.stepHoldTimer = null;
        }
        if (this.stepTimer !== null) {
            window.clearInterval(this.stepTimer);
            this.stepTimer = null;
        }
    }

    /** 端点箭头滚一格 */
    private stepBy(direction: number): void {
        const target = this.options.target;
        if (this.options.direction === 'horizontal') target.scrollLeft += direction * STEP_PX;
        else target.scrollTop += direction * STEP_PX;
    }

    /**
     * 销毁
     *
     * 工作区是常驻 DOM 的（切到表格只是 hidden），正常流程用不到它；
     * 留着是为了让「观察器 / 全局监听」有明确的收口，避免打包工具发现无法回收时误报。
     */
    public destroy(): void {
        this.stopStep();
        document.removeEventListener('mousemove', this.onDocMouseMove);
        document.removeEventListener('mouseup', this.onDocMouseUp);
        this.options.target.removeEventListener('scroll', this.onTargetScroll);
        this.observer?.disconnect();
        this.observer = null;
        this.observed = null;
        this.container.remove();
    }
}
