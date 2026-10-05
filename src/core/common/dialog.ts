/**
 * 通用 Dialog 组件
 * 按 Bootstrap 5 Modal 风格实现的对话框控件：
 * 遮罩层 + 弹窗容器（header/body/footer 三段式结构），
 * 支持 ESC 关闭、点击遮罩关闭、尺寸、垂直居中、内容滚动与底部按钮组。
 *
 * @module Dialog
 * @author 东方鹗
 * @see
 * # B站: https://space.bilibili.com/194359739
 * # 知乎: https://www.zhihu.com/people/eastossifrage
 * # CSDN: https://blog.csdn.net/os373
 *
 * @example
 * ```typescript
 * const dialog = new Dialog({
 *     title: '确认操作',
 *     content: '确定要删除选中内容吗？',
 *     buttons: [
 *         { text: '取消', variant: 'secondary' },
 *         { text: '确定', variant: 'primary', onClick: () => console.log('confirmed') }
 *     ]
 * });
 * dialog.show();
 * ```
 */

import { createDiv, createButton } from "../../utils/dom";
import { EventEmitter } from "../../utils/eventEmitter";

/** 底部按钮的视觉风格，对应 Bootstrap 的 btn-* 系列 */
export type DialogButtonVariant =
    | 'primary' | 'secondary' | 'success' | 'danger'
    | 'warning' | 'info' | 'light' | 'dark' | 'link';

/** 对话框尺寸，对应 Bootstrap 的 modal-sm / 默认 / modal-lg / modal-xl */
export type DialogSize = 'sm' | 'md' | 'lg' | 'xl';

/** 底部按钮配置 */
export interface DialogButtonOptions {
    /** 按钮文字 */
    text: string;
    /** 按钮视觉风格，默认 'primary' */
    variant?: DialogButtonVariant;
    /** 点击回调 */
    onClick?: (dialog: Dialog) => void;
    /** 点击后是否关闭对话框（对应 data-bs-dismiss="modal"），默认 true */
    close?: boolean;
    /** 是否在对话框打开时自动聚焦该按钮，默认 false */
    autoFocus?: boolean;
}

export interface DialogOptions {
    /** 父元素，默认 document.body */
    parentElement?: HTMLElement | null;
    /** 标题文字 */
    title?: string;
    /** 正文内容：纯文本或已构建好的 DOM 元素 */
    content?: string | HTMLElement;
    /** 底部按钮组，提供时才渲染 footer 区域 */
    buttons?: DialogButtonOptions[];
    /** 是否显示右上角关闭按钮（×），默认 true */
    showClose?: boolean;
    /** 是否允许点击遮罩层关闭，默认 false（点击对话框外不自动关闭，需显式传 true 开启） */
    closeOnBackdrop?: boolean;
    /** 是否允许按 ESC 关闭，默认 true */
    closeOnEsc?: boolean;
    /** 对话框尺寸，默认 'md'（500px） */
    size?: DialogSize;
    /** 是否垂直居中显示（对应 modal-dialog-centered），默认 true */
    centered?: boolean;
    /** 是否允许按住标题栏拖拽移动对话框位置，默认 true */
    draggable?: boolean;
    /** 正文超出时是否内部滚动（对应 modal-dialog-scrollable），默认 false */
    scrollable?: boolean;
    /** 显隐过渡动画时长（毫秒），默认 300 */
    duration?: number;
    /** 对话框关闭后的回调 */
    onClose?: () => void;
}

export class Dialog extends EventEmitter {
    private options: DialogOptions;
    /** 遮罩层根元素（整层绝对定位铺满视口） */
    private dialogElement: HTMLElement;
    /** modal-dialog 元素（拖拽位移作用于它） */
    private dialogBoxElement: HTMLElement;
    /** 标题栏元素（拖拽手柄） */
    private headerElement: HTMLElement;
    /** 拖拽累计偏移（相对初始居中位置），hide/show 间保留 */
    private dialogOffset: { x: number; y: number } = { x: 0, y: 0 };
    /** modal-content 容器（header/body/footer 的父级） */
    private contentWrapperElement: HTMLElement;
    /** 标题元素 */
    private titleElement: HTMLElement;
    /** 正文区域（modal-body） */
    private contentElement: HTMLElement;
    /** 右上角关闭按钮（×） */
    private closeElement: HTMLElement | null;
    /** 底部按钮区域（modal-footer），未提供 buttons 时为 null */
    private footerElement: HTMLElement | null;
    /** 底部按钮实例列表 */
    private footerButtons: HTMLButtonElement[] = [];
    /** 显隐过渡动画时长（毫秒） */
    private duration: number;
    /** 是否显示右上角关闭按钮 */
    private showClose: boolean;
    /** 对话框关闭后的回调 */
    private onClose: () => void | null;
    /** 当前是否处于打开状态 */
    private isOpen: boolean = false;
    /** 关闭动画的定时器句柄，防止快速开关竞态 */
    private hideTimer: ReturnType<typeof setTimeout> | null = null;
    /** ESC 键处理器（show 时挂载、hide/destroy 时移除） */
    private readonly handleEscKeyDown = (e: KeyboardEvent): void => {
        if (e.key === 'Escape' && this.options.closeOnEsc !== false) {
            this.hide();
        }
    };

    constructor(options: DialogOptions = {}) {
        super();
        this.options = options;
        this.duration = options.duration ?? 300;
        this.showClose = options.showClose ?? true;
        this.onClose = options.onClose ?? null;

        // —— 遮罩层根元素：默认隐藏，show() 时切换为 flex ——
        this.dialogElement = createDiv({
            className: 'gs-modal',
            style: {
                position: 'fixed',
                top: '0',
                left: '0',
                width: '100%',
                height: '100%',
                zIndex: '3000',
                display: 'none',
                justifyContent: 'center',
            }
        });
        if (options.centered !== false) {
            this.dialogElement.style.alignItems = 'center';
        }

        // —— modal-dialog：控制尺寸与进出场位移 ——
        const dialog = this.dialogBoxElement = createDiv({
            className: `gs-modal-dialog gs-modal-${options.size ?? 'md'}`,
            style: { width: '100%' }
        });
        if (options.scrollable) {
            dialog.classList.add('gs-modal-scrollable');
        }

        // —— modal-content ——
        this.contentWrapperElement = createDiv({ className: 'gs-modal-content' });

        // —— modal-header：标题 + 关闭按钮 ——
        const header = this.headerElement = createDiv({ className: 'gs-modal-header' });
        this.titleElement = createDiv({ className: 'gs-modal-title', textContent: options.title ?? '' });
        header.appendChild(this.titleElement);

        this.closeElement = null;
        if (this.showClose) {
            this.closeElement = createButton({
                className: 'gs-btn-close',
                textContent: '×',
                attributes: { 'aria-label': 'Close' }
            });
            this.closeElement.addEventListener('click', () => this.hide());
            header.appendChild(this.closeElement);
        }

        // —— modal-body ——
        this.contentElement = createDiv({ className: 'gs-modal-body' });
        if (typeof options.content === 'string') {
            this.contentElement.textContent = options.content;
        } else if (options.content) {
            this.contentElement.appendChild(options.content);
        }

        this.contentWrapperElement.appendChild(header);
        this.contentWrapperElement.appendChild(this.contentElement);

        // 标题栏拖拽移动对话框（默认开启，draggable: false 可关闭）
        if (options.draggable !== false) {
            this.setupDialogDrag();
        }

        // —— modal-footer：仅在提供按钮时渲染 ——
        this.footerElement = null;
        if (options.buttons?.length) {
            this.footerElement = createDiv({ className: 'gs-modal-footer' });
            for (const btnOptions of options.buttons) {
                this.footerButtons.push(this.createFooterButton(btnOptions));
            }
            this.contentWrapperElement.appendChild(this.footerElement);
        }

        dialog.appendChild(this.contentWrapperElement);
        this.dialogElement.appendChild(dialog);

        // 点击遮罩层（而非弹窗内部）关闭：默认关闭该行为（点击他处不消失），需显式传 closeOnBackdrop: true 开启
        this.dialogElement.addEventListener('click', (e) => {
            if (e.target === this.dialogElement && this.options.closeOnBackdrop === true) {
                this.hide();
            }
        });

        // 样式只注入一次，避免每个实例都向 head 追加重复的 style 标签
        if (!document.getElementById('gs-modal-styles')) {
            const styleSheet = document.createElement('style');
            styleSheet.id = 'gs-modal-styles';
            styleSheet.textContent = this.buildStyles();
            document.head.appendChild(styleSheet);
        }
    }

    /**
     * 标题栏拖拽移动对话框
     *
     * 按住 modal-header（× 关闭按钮除外）拖动，位移以内联 transform 叠加在居中布局上；
     * 拖拽期间添加 gs-dragging 类关闭过渡动画，使移动实时跟手。
     * 偏移量在 hide/show 间保留：复用同一实例再次打开时保持在用户拖放的位置。
     * @private
     */
    private setupDialogDrag(): void {
        const dialog = this.dialogBoxElement;
        const header = this.headerElement;
        header.style.cursor = 'move';
        header.addEventListener('pointerdown', (e) => {
            if (e.button !== 0) return;
            // × 关闭按钮不作为拖拽手柄，保证点击关闭行为不受影响
            if ((e.target as HTMLElement | null)?.closest('.gs-btn-close')) return;
            e.preventDefault();
            const startX = e.clientX;
            const startY = e.clientY;
            const baseX = this.dialogOffset.x;
            const baseY = this.dialogOffset.y;
            dialog.classList.add('gs-dragging');
            const onMove = (ev: PointerEvent) => {
                this.dialogOffset.x = baseX + (ev.clientX - startX);
                this.dialogOffset.y = baseY + (ev.clientY - startY);
                dialog.style.transform = `translate(${this.dialogOffset.x}px, ${this.dialogOffset.y}px)`;
            };
            const onUp = () => {
                dialog.classList.remove('gs-dragging');
                window.removeEventListener('pointermove', onMove);
                window.removeEventListener('pointerup', onUp);
                window.removeEventListener('pointercancel', onUp);
            };
            window.addEventListener('pointermove', onMove);
            window.addEventListener('pointerup', onUp);
            window.addEventListener('pointercancel', onUp);
        });
    }

    /**
     * 创建底部按钮
     * @param {DialogButtonOptions} btnOptions - 按钮配置
     * @returns {HTMLButtonElement} 创建的按钮元素
     * @private
     */
    private createFooterButton(btnOptions: DialogButtonOptions): HTMLButtonElement {
        const variant = btnOptions.variant ?? 'primary';
        const button = createButton({
            className: `gs-btn gs-btn-${variant}`,
            textContent: btnOptions.text
        });
        if (btnOptions.autoFocus) {
            button.classList.add('gs-btn-focus');
        }
        button.addEventListener('click', () => {
            btnOptions.onClick?.(this);
            // 未显式声明 close 时默认点击后关闭（对应 data-bs-dismiss="modal"）
            if (btnOptions.close !== false) {
                this.hide();
            }
        });
        this.footerElement?.appendChild(button);
        return button;
    }

    /**
     * 生成组件样式表（Bootstrap 5 Modal 视觉规格）
     * @returns {string} CSS 文本
     * @private
     */
    private buildStyles(): string {
        const duration = this.duration;
        return `
            .gs-modal {
                opacity: 0;
                transition: opacity ${duration}ms ease;
            }
            .gs-modal.gs-show { opacity: 1; }
            .gs-modal-dialog {
                max-width: 500px;
                margin: 1.75rem;
                transform: translateY(-50px);
                transition: transform ${duration}ms ease-out;
                pointer-events: auto;
            }
            .gs-modal.gs-show .gs-modal-dialog { transform: none; }
            /* 拖拽中关闭位移过渡，使标题栏拖动实时跟手 */
            .gs-modal-dialog.gs-dragging { transition: none; }
            /* 尺寸：对应 Bootstrap modal-sm / 默认 / modal-lg / modal-xl */
            .gs-modal-dialog.gs-modal-sm { max-width: 300px; }
            .gs-modal-dialog.gs-modal-lg { max-width: 800px; }
            .gs-modal-dialog.gs-modal-xl { max-width: 1140px; }
            .gs-modal-content {
                display: flex;
                flex-direction: column;
                background: #fff;
                color: #212529;
                border: 1px solid rgba(0, 0, 0, 0.175);
                border-radius: 0.5rem;
                box-shadow: 0 0.5rem 1rem rgba(0, 0, 0, 0.15);
                outline: 0;
            }
            /* 滚动模式：正文超出时内部滚动（对应 modal-dialog-scrollable） */
            .gs-modal-dialog.gs-modal-scrollable { max-height: calc(100% - 3.5rem); }
            .gs-modal-dialog.gs-modal-scrollable .gs-modal-content { max-height: 100%; overflow: hidden; }
            .gs-modal-dialog.gs-modal-scrollable .gs-modal-body { overflow-y: auto; }
            .gs-modal-header {
                display: flex;
                align-items: center;
                justify-content: space-between;
                flex-shrink: 0;
                padding: 0.5rem 0.75rem;
                border-bottom: 1px solid #dee2e6;
            }
            .gs-modal-title {
                margin: 0;
                font-size: 1rem;
                line-height: 1.5;
                color: inherit;
            }
            .gs-btn-close {
                background: transparent;
                border: none;
                font-size: 1.25rem;
                line-height: 1;
                color: #000;
                opacity: 0.5;
                cursor: pointer;
                padding: 0.125rem 0.375rem;
                border-radius: 0.375rem;
                transition: opacity 0.15s ease-in-out;
            }
            .gs-btn-close:hover { opacity: 1; }
            .gs-btn-close:focus { outline: 0; opacity: 1; box-shadow: 0 0 0 0.25rem rgba(13, 110, 253, 0.25); }
            /* 隐藏原生数字步进按钮：对话框内统一使用自定义步进按钮（支持按住连发） */
            .gs-modal input[type='number'] { appearance: textfield; -moz-appearance: textfield; }
            .gs-modal input[type='number']::-webkit-inner-spin-button,
            .gs-modal input[type='number']::-webkit-outer-spin-button { -webkit-appearance: none; margin: 0; }
            .gs-modal-body {
                position: relative;
                flex: 1 1 auto;
                padding: 0.75rem;
                font-size: 0.875rem;
                line-height: 1.5;
                color: #212529;
                word-break: break-word;
            }
            .gs-modal-footer {
                display: flex;
                flex-wrap: wrap;
                justify-content: flex-end;
                gap: 0.25rem;
                flex-shrink: 0;
                padding: 0.5rem 0.75rem;
                border-top: 1px solid #dee2e6;
            }
            /* 按钮：对应 Bootstrap .btn 与 btn-* 视觉规格 */
            .gs-btn {
                display: inline-block;
                padding: 0.25rem 0.5rem;
                font-size: 0.875rem;
                line-height: 1.5;
                text-align: center;
                border: 1px solid transparent;
                border-radius: 0.375rem;
                cursor: pointer;
                transition: color 0.15s ease-in-out, background-color 0.15s ease-in-out,
                    border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out;
            }
            .gs-btn:focus { outline: 0; box-shadow: 0 0 0 0.25rem rgba(13, 110, 253, 0.25); }
            .gs-btn-primary { color: #fff; background-color: #0d6efd; border-color: #0d6efd; }
            .gs-btn-primary:hover { color: #fff; background-color: #0b5ed7; border-color: #0a58ca; }
            .gs-btn-primary:focus { box-shadow: 0 0 0 0.25rem rgba(49, 132, 253, 0.5); }
            .gs-btn-secondary { color: #fff; background-color: #6c757d; border-color: #6c757d; }
            .gs-btn-secondary:hover { color: #fff; background-color: #5c636a; border-color: #565e64; }
            .gs-btn-secondary:focus { box-shadow: 0 0 0 0.25rem rgba(130, 138, 145, 0.5); }
            .gs-btn-success { color: #fff; background-color: #198754; border-color: #198754; }
            .gs-btn-success:hover { color: #fff; background-color: #157347; border-color: #146c43; }
            .gs-btn-success:focus { box-shadow: 0 0 0 0.25rem rgba(60, 153, 110, 0.5); }
            .gs-btn-danger { color: #fff; background-color: #dc3545; border-color: #dc3545; }
            .gs-btn-danger:hover { color: #fff; background-color: #bb2d3b; border-color: #b02a37; }
            .gs-btn-danger:focus { box-shadow: 0 0 0 0.25rem rgba(225, 83, 97, 0.5); }
            .gs-btn-warning { color: #000; background-color: #ffc107; border-color: #ffc107; }
            .gs-btn-warning:hover { color: #000; background-color: #ffca2c; border-color: #ffc720; }
            .gs-btn-warning:focus { box-shadow: 0 0 0 0.25rem rgba(217, 164, 6, 0.5); }
            .gs-btn-info { color: #000; background-color: #0dcaf0; border-color: #0dcaf0; }
            .gs-btn-info:hover { color: #000; background-color: #31d2f2; border-color: #25cff2; }
            .gs-btn-info:focus { box-shadow: 0 0 0 0.25rem rgba(11, 172, 204, 0.5); }
            .gs-btn-light { color: #000; background-color: #f8f9fa; border-color: #f8f9fa; }
            .gs-btn-light:hover { color: #000; background-color: #d3d4d5; border-color: #c6c7c8; }
            .gs-btn-light:focus { box-shadow: 0 0 0 0.25rem rgba(211, 212, 213, 0.5); }
            .gs-btn-dark { color: #fff; background-color: #212529; border-color: #212529; }
            .gs-btn-dark:hover { color: #fff; background-color: #424649; border-color: #373b3e; }
            .gs-btn-dark:focus { box-shadow: 0 0 0 0.25rem rgba(66, 70, 73, 0.5); }
            .gs-btn-link {
                color: #0d6efd;
                background-color: transparent;
                border-color: transparent;
                text-decoration: underline;
            }
            .gs-btn-link:hover { color: #0a58ca; }
            .gs-btn-link:focus { box-shadow: none; }
        `;
    }

    /**
     * 显示对话框（挂载到父元素并播放进场动画）
     * @public
     */
    public show(): void {
        if (this.isOpen) return;
        // 若上次关闭动画未结束，取消其隐藏回调，避免动画中途被反向打断后误触发 onClose
        if (this.hideTimer !== null) {
            clearTimeout(this.hideTimer);
            this.hideTimer = null;
        }
        const parent = this.options.parentElement ?? document.body;
        if (!this.dialogElement.parentElement) {
            parent.appendChild(this.dialogElement);
        }
        this.isOpen = true;
        this.dialogElement.style.display = 'flex';
        // 强制回流，使 display 变化先生效，随后添加 gs-show 才能触发过渡动画
        this.dialogElement.getBoundingClientRect();
        this.dialogElement.classList.add('gs-show');
        if (this.options.closeOnEsc !== false) {
            document.addEventListener('keydown', this.handleEscKeyDown);
        }
        // 聚焦声明的 autoFocus 按钮，否则聚焦关闭按钮，保证键盘操作起点在弹窗内
        const focusTarget =
            this.footerButtons.find((btn) => btn.classList.contains('gs-btn-focus')) ?? this.closeElement;
        focusTarget?.focus();
        this.emit('show');
    }

    /**
     * 隐藏对话框（播放出场动画后隐藏，触发 onClose 回调与 close 事件）
     * @public
     */
    public hide(): void {
        if (!this.isOpen) return;
        this.isOpen = false;
        document.removeEventListener('keydown', this.handleEscKeyDown);
        this.dialogElement.classList.remove('gs-show');
        this.hideTimer = setTimeout(() => {
            this.hideTimer = null;
            this.dialogElement.style.display = 'none';
            this.onClose?.();
            this.emit('close');
        }, this.duration);
    }

    /**
     * 隐藏对话框（hide 的别名，语义化方法）
     * @public
     */
    public close(): void {
        this.hide();
    }

    /**
     * 对话框是否处于打开状态
     * @returns {boolean} 打开返回 true
     * @public
     */
    public isVisible(): boolean {
        return this.isOpen;
    }

    /**
     * 销毁对话框：移除事件监听并从 DOM 摘除根元素
     * @public
     */
    public destroy(): void {
        document.removeEventListener('keydown', this.handleEscKeyDown);
        if (this.hideTimer !== null) {
            clearTimeout(this.hideTimer);
            this.hideTimer = null;
        }
        this.isOpen = false;
        this.dialogElement.remove();
    }

    /**
     * 设置标题
     * @param {string} title - 标题文字
     * @public
     */
    public setTitle(title: string): void {
        this.titleElement.textContent = title;
    }

    /**
     * 设置正文内容
     * @param {string | HTMLElement} content - 纯文本或 DOM 元素
     * @public
     */
    public setContent(content: string | HTMLElement): void {
        this.contentElement.innerHTML = '';
        if (typeof content === 'string') {
            this.contentElement.textContent = content;
        } else {
            this.contentElement.appendChild(content);
        }
    }

    /**
     * 获取正文容器，便于调用方自行构建复杂内容
     * @returns {HTMLElement} modal-body 元素
     * @public
     */
    public getBody(): HTMLElement {
        return this.contentElement;
    }

    /**
     * 获取根元素（遮罩层），便于调用方追加自定义样式
     * @returns {HTMLElement} 遮罩层根元素
     * @public
     */
    public getElement(): HTMLElement {
        return this.dialogElement;
    }
}
