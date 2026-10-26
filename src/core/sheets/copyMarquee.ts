import { DataCollection } from '../dataArchitecture/dataCollection';
import '../style/copyMarquee.scss';

/** 虚线框的四条边（顺时针：上 → 右 → 下 → 左） */
const MARQUEE_EDGES = ['top', 'right', 'bottom', 'left'] as const;

/**
 * 流动虚线框组件选项
 */
export interface CopyMarqueeOptions {
    /** 挂载容器（Sheets 的定位容器，需 position: relative） */
    parentElement: HTMLElement;
    /** 数据实例 */
    data: DataCollection;
}

/**
 * 复制/剪切后的流动虚线框（Excel 的 marching ants）
 *
 * ## 为什么是 DOM 而不是画在 canvas 上
 *
 * 虚线要「跑动」，就必须持续重绘。若画在 {@link SheetCanvas} 里，每帧都要重绘
 * 整个表格（网格线、底色、边框、全部文字），大表上 60fps 是灾难。改用四个绝对
 * 定位的 1px 边缘 div + CSS 位移动画：动画跑在合成线程上，与表格重绘完全解耦，
 * 表格只在滚动/缩放/选区变化时正常重绘一帧。
 *
 * 四条边按顺时针方向流动：上边向右、右边向下、下边向左、左边向上，
 * 视觉上就是一圈绕行（Excel 同款）。
 *
 * ## 位置同步
 *
 * 自身不监听滚动事件，而是由 {@link Sheets.draw} 每帧调用 {@link sync}：
 * Sheets.draw 已覆盖滚动、缩放、尺寸变化等全部场景，且已经合帧，
 * 这里只需按最新的 {@link DataCollection.copyRange} 重算矩形。
 *
 * @class CopyMarquee
 * @example
 * const marquee = new CopyMarquee({ parentElement: sheetsDiv, data });
 * marquee.sync();   // 按 data.copyRange 显示/隐藏并定位
 */
export class CopyMarquee {
    /** 数据实例 */
    private readonly data: DataCollection;
    /** 虚线框容器（四条边的父元素） */
    private readonly container: HTMLDivElement;
    /** 当前是否可见 */
    private visible: boolean = false;
    /** 上一次应用的位置，避免每帧重复写样式触发排版 */
    private readonly lastRect = { left: NaN, top: NaN, width: NaN, height: NaN };

    /**
     * 构造函数
     * @param {CopyMarqueeOptions} options - 初始化选项
     */
    constructor(options: CopyMarqueeOptions) {
        this.data = options.data;
        this.container = document.createElement('div');
        this.container.className = 'copy-marquee';
        for (const edge of MARQUEE_EDGES) {
            const line = document.createElement('div');
            line.className = `copy-marquee-edge copy-marquee-${edge}`;
            this.container.appendChild(line);
        }
        this.container.style.display = 'none';
        options.parentElement.appendChild(this.container);
    }

    /**
     * 按当前复制源范围同步虚线框的位置与显隐
     *
     * 范围为空、完全落在隐藏行列上（宽高为 0）时隐藏。
     * 矩形可能超出视口：容器由 Sheets 的 `overflow: hidden` 裁掉，
     * 超出部分不会盖住行头/列头（那两个画布的 z-index 更高）。
     */
    public sync(): void {
        const range = this.data.copyRange;
        if (!range) {
            this.hide();
            return;
        }
        const { left, top, width, height } = this.data.getRangeRect(range);
        if (width <= 0 || height <= 0) {
            this.hide();
            return;
        }
        const zoom = this.data.zoom;
        const rect = {
            left: this.data.visibleView.rowHeaderWidth + left * zoom,
            top: this.data.visibleView.colHeaderHeight + top * zoom,
            width: width * zoom,
            height: height * zoom
        };
        const style = this.container.style;
        if (rect.left !== this.lastRect.left) style.left = `${rect.left}px`;
        if (rect.top !== this.lastRect.top) style.top = `${rect.top}px`;
        if (rect.width !== this.lastRect.width) style.width = `${rect.width}px`;
        if (rect.height !== this.lastRect.height) style.height = `${rect.height}px`;
        this.lastRect.left = rect.left;
        this.lastRect.top = rect.top;
        this.lastRect.width = rect.width;
        this.lastRect.height = rect.height;
        if (!this.visible) {
            style.display = '';
            this.visible = true;
        }
    }

    /**
     * 隐藏虚线框
     * @private
     */
    private hide(): void {
        if (!this.visible) return;
        this.container.style.display = 'none';
        this.visible = false;
    }

    /**
     * 销毁组件（移除 DOM，CSS 动画随元素一起停止）
     */
    public destroy(): void {
        this.container.remove();
    }
}
