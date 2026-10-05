/*  # web 表格辅助组件-隐藏组件 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { createDiv } from "../../utils/dom";
import { ASSISTANT } from "../constant";
import { DataCollection } from "../dataArchitecture/dataCollection";
import "../style/hidden.scss";

/**
 * 隐藏组件初始化选项
 * @property {HTMLElement} parentElement - 隐藏组件的挂载父元素，隐藏根节点会被追加到该元素内
 * @property {DataCollection} data - 表格数据集合
 * @property {'col' | 'row'} axis - 指示器方向：'row'=上下双三角（隐藏行），'col'=左右双三角（隐藏列）
 */
export interface HiddenOptions {
    parentElement: HTMLElement;
    data: DataCollection;
    axis: 'col' | 'row';
}

/**
 * 隐藏组件类
 * @class
 * @description 隐藏行列的悬停指示按钮（双三角图标）。
 * 默认隐藏，由行头/列头画布在鼠标悬停到「分割线且邻接隐藏行列」的位置时调用
 * {@link show} 显示、离开时调用 {@link hide} 隐藏。每个画布复用同一个实例。
 *
 * @example
     * const hidden = new Hidden({ parentElement, data, axis: 'row' });
     * hidden.show(100, 200, { start: 3, end: 5 });  // 在父容器坐标 (100, 200) 处显示，点击可取消隐藏第 3~5 行
     * hidden.hide();          // 隐藏
     * hidden.destroy();       // 销毁实例并移除 DOM
 */
export class Hidden {
    private parentElement: HTMLElement;
    private hiddenContainer: HTMLElement;
    private data: DataCollection;
    private axis: string;
    /** 按钮当前对应的隐藏行列块（1-based，含端点）；null 表示按钮未显示 */
    private hiddenRange: { start: number; end: number } | null = null;
    constructor(options: HiddenOptions) {
        const { parentElement, data, axis } = options;
        this.parentElement = parentElement;
        this.hiddenContainer = createDiv({
            className: ASSISTANT.HIDDEN_CONTAINER_CLASS_NAME + ` ${axis}`,
        });
        this.hiddenContainer.style.display = 'none'; // 默认隐藏，悬停命中分割线时显示
        this.parentElement.appendChild(this.hiddenContainer);
        this.data = data;
        this.axis = axis;
        this.setupEventListeners();
    }

    /**
     * 在指定位置显示隐藏指示按钮
     * @method
     * @description 按钮采用 absolute 定位，left/top 为相对挂载父容器的 CSS 像素偏移
     * @param {number} left - 相对父容器的 X 偏移
     * @param {number} top - 相对父容器的 Y 偏移
     * @param {{ start: number, end: number }} range - 按钮对应的隐藏行列块（1-based，含端点），点击按钮时取消隐藏该块
     */
    public show(left: number, top: number, range: { start: number; end: number }): void {
        this.hiddenContainer.style.left = `${left}px`;
        this.hiddenContainer.style.top = `${top}px`;
        this.hiddenRange = range;
        // 必须用 flex：行内样式优先级高于样式表，若设为 block 会覆盖
        // .hidden-container 的 display: flex，导致双三角（块级伪元素）垂直堆叠错行
        this.hiddenContainer.style.display = 'flex';
    }

    /**
     * 隐藏指示按钮
     * @method
     * @description 同步清空按钮对应的隐藏行列块记录，防止残留过期范围
     */
    public hide(): void {
        this.hiddenContainer.style.display = 'none';
        this.hiddenRange = null;
    }

    /**
     * 判断指定节点是否位于指示按钮内部
     * @method
     * @description 供画布 mouseleave 事件区分「鼠标移入按钮」与「真正移出画布」：
     * 移入按钮时不隐藏（按钮已可交互，由其自身 mouseleave 负责隐藏）。
     * @param {Node | null} target - 待判断的节点（通常取 mouseleave 事件的 relatedTarget）
     * @returns {boolean} 节点位于按钮内返回 true，否则返回 false
     */
    public contains(target: Node | null): boolean {
        return !!target && this.hiddenContainer.contains(target);
    }

    /**
     * 销毁组件实例，从父元素中移除 DOM 节点
     * @method
     */
    public destroy(): void {
        this.parentElement.removeChild(this.hiddenContainer);
    }

    /**
     * 设置按钮的交互事件监听
     * @method
     * @description 点击按钮：取消隐藏该按钮位置对应的隐藏行列块。按钮可正常接收鼠标
     * 事件，点击被按钮拦截后画布不会收到 mousedown/click，因此不会触发行列头的选中
     * 逻辑，选中的单元格保持不变。取消隐藏使用全量状态快照，支持撤销/重做。
     * 鼠标离开按钮时隐藏：若移向画布外，画布不会再产生 mousemove/mouseleave 事件，
     * 需要按钮自行负责隐藏。
     */
    private setupEventListeners(): void {
        this.hiddenContainer.addEventListener('click', () => {
            if (!this.hiddenRange) return;
            const range = this.hiddenRange;
            this.data.runWithFullStateUndo(this.axis === 'row' ? '取消隐藏行' : '取消隐藏列', () => {
                if (this.axis === 'row') {
                    this.data.setRowsHidden(range.start, range.end, false);
                } else {
                    this.data.setColsHidden(range.start, range.end, false);
                }
            });
            this.hide();
        });
        this.hiddenContainer.addEventListener('mouseleave', () => this.hide());
    }
}
