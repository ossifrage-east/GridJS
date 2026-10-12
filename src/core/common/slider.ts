/*  # 通用滑块组件 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { createDiv, createInput, createSpan } from "../../utils/dom";
import '../style/slider.scss';

/**
 * 滑块配置选项接口
 * @interface SliderOptions
 */
export interface SliderOptions {
    /** 父元素（缺省时仅创建 DOM，不挂载） */
    parentElement?: HTMLElement | null;
    /** 最小值 */
    min?: number;
    /** 最大值 */
    max?: number;
    /** 步长 */
    step?: number;
    /** 初始值 */
    value?: number;
    /** 数值输入框右侧的单位文本，如 '%'；为空字符串则不渲染 */
    unit?: string;
    /** 是否显示数值输入框（与滑块双向同步），默认 true */
    showInput?: boolean;
    /**
     * 是否垂直方向，默认 false（水平）
     *
     * 垂直时轨道纵向排列、**最小值在下端**，取值换算改为读指针 Y 坐标；
     * 未显式指定 {@link sliderWidth} 时不写死内联长度，由样式表撑满容器高度
     */
    vertical?: boolean;
    /**
     * 是否使用细线样式，默认 false（浏览器原生外观）
     *
     * 细线样式下轨道为 2px 细线、手柄为小圆点，配色取自主题色板
     * （见 slider.scss），适合窄条工具区里的紧凑控件
     */
    thin?: boolean;
    /** 滑块轨道长度（CSS 长度）：水平方向写入 width，垂直方向写入 height */
    sliderWidth?: string;
    /** 数值输入框宽度（CSS 长度） */
    inputWidth?: string;
    /** 滑块激活色（accent-color） */
    accentColor?: string;
    /** 值变化回调（拖动过程与输入框提交均触发） */
    onChange?: (value: number, source: SliderChangeSource) => void;
}

/**
 * 值变化来源
 * - 'track'：滑动轨道（拖动 / 键盘方向键）导致的变化
 * - 'input'：数值输入框提交导致的变化
 */
export type SliderChangeSource = 'track' | 'input';

/**
 * 滑块组件
 *
 * 由「滑动轨道 + 可选数值输入框 + 可选单位文本」组成的三联动控件：
 * - 轨道：按下即跳转到点击位置，按住拖动连续跟手（自实现指针手势，
 *   规避部分环境下 range 原生拖拽不跟手的问题），键盘方向键仍走原生行为；
 * - 数值输入框：提交时收敛到 [min, max] 并按 step 吸附后回写轨道；
 * - 单位文本：仅作展示，不参与取值。
 *
 * 双向同步规则：
 * - 轨道 → 数值框：拖动过程中同步；数值框聚焦**且存在未提交的编辑**时跳过，
 *   避免覆盖输入（仅聚焦、内容未变时照常同步）；
 * - 数值框 → 轨道：提交时（change 失焦 / 回车）同步；
 * - 外部 {@link setValue}：默认两者都同步，可通过参数只同步数值框
 *   （拖动期间由控件内部完成，避免程序赋值打断原生拖拽手势）。
 *
 * 方向与外观：
 * - `vertical: true` 时轨道纵向排列、最小值在下端，拖动换算读指针 Y 坐标；
 * - `thin: true` 时使用细线外观（2px 轨道 + 小圆点手柄，主题色），
 *   适合右侧工具栏这类窄条区域里的紧凑滑块。
 *
 * @class Slider
 * @example
 * const slider = new Slider({
 *     parentElement,
 *     min: 10,
 *     max: 400,
 *     unit: '%',
 *     onChange: (value) => console.log(value),
 * });
 * slider.setValue(120);
 */
export class Slider {
    /** 组件根容器 */
    private readonly container: HTMLDivElement;
    /** 滑动轨道（input[type=range]） */
    private readonly track: HTMLInputElement;
    /** 数值输入框（showInput 为 false 时为 null） */
    private readonly numberInput: HTMLInputElement | null;
    /** 单位文本元素（unit 为空时为 null） */
    private readonly unitElement: HTMLSpanElement | null;
    /** 配置选项 */
    private readonly options: SliderOptions;
    /**
     * 最近一次由组件写入数值框的值
     *
     * 用于区分「框内存在未提交的编辑」与「仅聚焦但内容未变」：
     * 前者同步时跳过（避免覆盖用户输入），后者照常同步
     * （例如提交后输入框仍保持聚焦时，Ctrl+滚轮 缩放仍能回写显示）。
     */
    private lastInputSynced = '';

    /**
     * 构造函数
     * @param {SliderOptions} [options={}] - 组件配置选项
     */
    constructor(options: SliderOptions = {}) {
        this.options = {
            min: 0,
            max: 100,
            step: 1,
            showInput: true,
            // 垂直方向的长度默认交由样式表撑满容器（100%），故不注入内联长度
            sliderWidth: options.vertical ? undefined : '120px',
            inputWidth: '3.5rem',
            accentColor: '#0d6efd',
            ...options
        };

        this.container = createDiv({ className: 'slider' });
        if (this.options.vertical) this.container.classList.add('slider-vertical');
        if (this.options.thin) this.container.classList.add('slider-thin');

        const { min, max, step } = this.options;
        this.track = createInput({
            className: 'slider-track',
            type: 'range',
            attributes: {
                min: String(min),
                max: String(max),
                step: String(step),
            },
        });
        if (this.options.sliderWidth) {
            // 轨道长度写在主轴方向：水平看宽度、垂直看高度
            if (this.options.vertical) {
                this.track.style.height = this.options.sliderWidth;
            } else {
                this.track.style.width = this.options.sliderWidth;
            }
        }
        if (this.options.accentColor) {
            this.track.style.accentColor = this.options.accentColor;
        }

        this.numberInput = this.options.showInput
            ? createInput({
                className: 'slider-input',
                type: 'number',
                attributes: {
                    min: String(min),
                    max: String(max),
                    step: String(step),
                },
            })
            : null;
        if (this.numberInput && this.options.inputWidth) {
            this.numberInput.style.width = this.options.inputWidth;
        }

        this.unitElement = this.options.unit
            ? createSpan({ className: 'slider-unit', textContent: this.options.unit })
            : null;

        this.container.append(this.track);
        if (this.numberInput) this.container.appendChild(this.numberInput);
        if (this.unitElement) this.container.appendChild(this.unitElement);

        this.bindEvents();
        this.setValue(this.options.value ?? min);

        if (options.parentElement) {
            options.parentElement.appendChild(this.container);
        }
    }

    /**
     * 绑定轨道与数值输入框事件
     * @private
     */
    private bindEvents(): void {
        const { track, numberInput } = this;
        const { vertical } = this.options;

        // —— 轨道：自实现指针拖动，保证跟手 ——
        // 水平：左端 min、右端 max；垂直：下端 min、上端 max（与 slider.scss 的纵向书写模式一致）
        const applyPointer = (clientX: number, clientY: number): void => {
            const rect = track.getBoundingClientRect();
            const min = Number(track.min) || 0;
            const max = Number(track.max) || 100;
            const step = Number(track.step) || 1;
            const length = vertical ? rect.height : rect.width;
            if (length <= 0) return;
            const ratio = vertical
                ? (rect.bottom - clientY) / length
                : (clientX - rect.left) / length;
            const clamped = Math.min(1, Math.max(0, ratio));
            const snapped = Math.round((min + clamped * (max - min)) / step) * step;
            track.value = String(Math.min(max, Math.max(min, snapped)));
            track.dispatchEvent(new Event('input'));
        };

        track.addEventListener('pointerdown', (e) => {
            if (e.pointerType === 'mouse' && e.button !== 0) return;  // 仅响应鼠标左键
            e.preventDefault();  // 接管拖动，避免原生手势叠加；焦点被阻止后手动恢复
            track.focus();
            // 指针捕获：拖出滑块/窗口范围仍持续跟手；不支持时退化为窗口级监听
            try { track.setPointerCapture(e.pointerId); } catch { /* ignore */ }
            applyPointer(e.clientX, e.clientY);  // 按下即跳转到点击位置（与原生轨道点击行为一致）
            const onMove = (ev: PointerEvent) => applyPointer(ev.clientX, ev.clientY);
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
        track.addEventListener('mousedown', (e) => e.preventDefault());

        // —— 轨道值变化：同步数值框并回调 ——
        track.addEventListener('input', () => {
            const n = Number(track.value);
            if (!Number.isFinite(n)) return;
            // 仅同步数值框：拖动期间不回写轨道 value，否则会打断原生拖拽手势
            this.syncNumberInput(n);
            this.options.onChange?.(n, 'track');
        });

        // —— 数值框提交：收敛并吸附后回写轨道 ——
        if (numberInput) {
            // 显式聚焦：宿主页面可能吞掉 mousedown 的默认行为（如表格组件的
            // 全局 preventDefault），导致点击无法聚焦、无法键入。与轨道一致，
            // 由组件自己接管焦点（pointerdown 覆盖鼠标 / 触屏 / 触控笔）。
            numberInput.addEventListener('pointerdown', () => numberInput.focus());

            numberInput.addEventListener('change', () => this.commitNumberInput());
            // 回车提交：手动填数后无需先点别处失焦
            numberInput.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    this.commitNumberInput();
                }
            });
        }
    }

    /**
     * 提交数值输入框内容：解析 → 吸附收敛 → 回写轨道 → 触发回调
     * （输入非法时回退为当前值，不改变状态）
     * @private
     */
    private commitNumberInput(): void {
        const { numberInput, track } = this;
        if (!numberInput) return;

        const n = Number(numberInput.value);
        const snapped = Number.isFinite(n) ? this.snap(n) : this.getValue();
        numberInput.value = String(snapped);
        this.lastInputSynced = String(snapped);   // 本次写入视为已同步，不再是「未提交的编辑」
        track.value = String(snapped);
        this.options.onChange?.(snapped, 'input');
    }

    /**
     * 按 step 吸附并收敛到 [min, max]
     * @private
     * @param {number} value - 原始值
     * @returns {number} 吸附后的值
     */
    private snap(value: number): number {
        const min = Number(this.track.min) || 0;
        const max = Number(this.track.max) || 100;
        const step = Number(this.track.step) || 1;
        const clamped = Math.min(max, Math.max(min, value));
        return Math.round(clamped / step) * step;
    }

    /**
     * 同步数值输入框显示
     *
     * 仅当「输入框处于聚焦状态 **且框内存在未提交的编辑**」时跳过，
     * 以免覆盖用户正在输入的内容；其它情况（含提交后仍聚焦、
     * 仅聚焦未改动）一律照常同步，保证外部变化（如 Ctrl+滚轮 缩放）
     * 能实时反映到数值框。
     * @private
     * @param {number} value - 目标值
     */
    private syncNumberInput(value: number): void {
        const { numberInput } = this;
        if (!numberInput) return;

        const target = String(value);
        const isEditing = document.activeElement === numberInput && numberInput.value !== this.lastInputSynced;
        if (isEditing) return;

        numberInput.value = target;
        this.lastInputSynced = target;
    }

    /**
     * 获取当前值（以轨道为准）
     * @returns {number} 当前值
     */
    public getValue(): number {
        const n = Number(this.track.value);
        return Number.isFinite(n) ? n : Number(this.options.min);
    }

    /**
     * 设置当前值
     * @param {number} value - 目标值（自动按 step 吸附并收敛到 [min, max]）
     * @param {boolean} [syncTrack=true] - 是否同步写回轨道。
     * 拖动过程中调用时应传 false：程序赋值轨道 value 会打断原生拖拽手势（滑块卡住不跟手），
     * 此时仅同步数值输入框即可
     */
    public setValue(value: number, syncTrack: boolean = true): void {
        const snapped = this.snap(value);
        if (syncTrack) {
            this.track.value = String(snapped);
        }
        this.syncNumberInput(snapped);
    }

    /**
     * 获取组件元素集合
     * @returns {{ container: HTMLDivElement, track: HTMLInputElement, numberInput: HTMLInputElement | null, unit: HTMLSpanElement | null }} 元素集合
     */
    public getElements(): {
        container: HTMLDivElement;
        track: HTMLInputElement;
        numberInput: HTMLInputElement | null;
        unit: HTMLSpanElement | null;
    } {
        return {
            container: this.container,
            track: this.track,
            numberInput: this.numberInput,
            unit: this.unitElement,
        };
    }

    /**
     * 销毁组件，移除 DOM
     */
    public destroy(): void {
        this.container.remove();
    }
}