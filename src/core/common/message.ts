/*  # 消息提示框 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { createDiv } from "../../utils/dom";

export type MessageType = 'success' | 'error' | 'warning' | 'info';

export interface MessageOptions {
    type?: MessageType;
    duration?: number;
    showClose?: boolean;
    onClose?: () => void;
}

export class MessageBox {
    private container: HTMLDivElement | null = null;
    private messageBox: HTMLDivElement | null = null;
    private confirmMask: HTMLDivElement | null = null;
    private timer: ReturnType<typeof setTimeout> | null = null;

    constructor() {
        this.initContainer();
    }

    private initContainer(): void {
        const existingContainer = document.getElementById('message-container');
        if (existingContainer) {
            this.container = existingContainer as HTMLDivElement;
            return;
        }

        this.container = createDiv({
            id: 'message-container',
            className: 'message-container'
        });
        
        this.container.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            z-index: 9999;
            display: flex;
            flex-direction: column;
            gap: 12px;
        `;

        document.body.appendChild(this.container);
    }

    public show(message: string, options: MessageOptions = {}): void {
        const { 
            type = 'info', 
            duration = 3000, 
            showClose = true,
            onClose 
        } = options;

        this.clearTimer();

        // 先安全清理旧提示框（可能处于隐藏状态）
        // 使用 Element.remove() 代替 container.removeChild()，元素不在父节点时不会抛错
        this.disposeMessageBox();

        const messageBox = createDiv({
            className: `message-box message-${type}`
        });

        const icon = createDiv({ className: 'message-icon' });
        const content = createDiv({ className: 'message-content' });
        const closeBtn = createDiv({ className: 'message-close' });

        content.textContent = message;
        closeBtn.innerHTML = '×';

        messageBox.appendChild(icon);
        messageBox.appendChild(content);
        
        if (showClose) {
            messageBox.appendChild(closeBtn);
        }

        messageBox.style.cssText = `
            display: flex;
            align-items: center;
            padding: 14px 16px;
            background: #fff;
            border-radius: 8px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
            min-width: 280px;
            max-width: 400px;
            animation: slideIn 0.3s ease-out;
        `;

        icon.style.cssText = `
            width: 20px;
            height: 20px;
            margin-right: 12px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 14px;
            font-weight: bold;
            flex-shrink: 0;
        `;

        content.style.cssText = `
            flex: 1;
            font-size: 14px;
            color: #333;
            line-height: 1.5;
        `;

        closeBtn.style.cssText = `
            width: 20px;
            height: 20px;
            margin-left: 12px;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            font-size: 18px;
            color: #999;
            border-radius: 50%;
            transition: all 0.2s;
        `;

        closeBtn.addEventListener('mouseenter', () => {
            closeBtn.style.cssText = `
                width: 20px;
                height: 20px;
                margin-left: 12px;
                display: flex;
                align-items: center;
                justify-content: center;
                cursor: pointer;
                font-size: 18px;
                color: #666;
                border-radius: 50%;
                transition: all 0.2s;
                background: #f5f5f5;
            `;
        });

        closeBtn.addEventListener('mouseleave', () => {
            closeBtn.style.cssText = `
                width: 20px;
                height: 20px;
                margin-left: 12px;
                display: flex;
                align-items: center;
                justify-content: center;
                cursor: pointer;
                font-size: 18px;
                color: #999;
                border-radius: 50%;
                transition: all 0.2s;
            `;
        });

        this.setIconStyle(icon, type);

        let closed = false;
        const close = () => {
            // 防止重复关闭（点击关闭按钮与定时器到期同时触发会导致 removeChild 报错）
            if (closed) return;
            closed = true;
            this.clearTimer();
            messageBox.style.animation = 'slideOut 0.2s ease-in';
            setTimeout(() => {
                // 隐藏而非移除，避免 removeChild 错误
                messageBox.style.display = 'none';
                onClose?.();
            }, 200);
        };

        closeBtn.addEventListener('click', close);

        if (duration > 0) {
            this.timer = setTimeout(close, duration);
        }

        this.messageBox = messageBox;
        this.container?.appendChild(messageBox);

        // 样式只注入一次，避免每次 show 都向 head 追加重复的 style 标签
        if (!document.getElementById('message-box-styles')) {
            const styleSheet = document.createElement('style');
            styleSheet.id = 'message-box-styles';
            styleSheet.textContent = `
                @keyframes slideIn {
                    from {
                        opacity: 0;
                        transform: translateX(100%);
                    }
                    to {
                        opacity: 1;
                        transform: translateX(0);
                    }
                }
                @keyframes slideOut {
                    from {
                        opacity: 1;
                        transform: translateX(0);
                    }
                    to {
                        opacity: 0;
                        transform: translateX(100%);
                    }
                }
            `;
            document.head.appendChild(styleSheet);
        }
    }

    /**
     * 安全清理当前提示框。
     * 使用 Element.remove() 代替 container.removeChild()：
     * Element.remove() 不依赖父节点，元素已不在 DOM 中时不会抛出
     * "The node to be removed is not a child of this node" 错误。
     * @private
     */
    private disposeMessageBox(): void {
        this.clearTimer();
        if (this.messageBox) {
            this.messageBox.remove();
            this.messageBox = null;
        }
    }

    private setIconStyle(icon: HTMLDivElement, type: MessageType): void {
        switch (type) {
            case 'success':
                icon.style.background = '#52c41a';
                icon.style.color = '#fff';
                icon.textContent = '✓';
                break;
            case 'error':
                icon.style.background = '#ff4d4f';
                icon.style.color = '#fff';
                icon.textContent = '✕';
                break;
            case 'warning':
                icon.style.background = '#faad14';
                icon.style.color = '#fff';
                icon.textContent = '!';
                break;
            case 'info':
            default:
                icon.style.background = '#1890ff';
                icon.style.color = '#fff';
                icon.textContent = 'i';
                break;
        }
    }

    private clearTimer(): void {
        if (this.timer) {
            clearTimeout(this.timer);
            this.timer = null;
        }
    }

    public close(): void {
        this.clearTimer();
        const box = this.messageBox;
        // 已隐藏（已关闭）时直接返回，防止重复触发动画
        if (!box || box.style.display === 'none') return;
        box.style.animation = 'slideOut 0.2s ease-in';
        setTimeout(() => {
            // 隐藏而非移除，避免 removeChild 错误
            box.style.display = 'none';
        }, 200);
    }

    public success(message: string, options?: Omit<MessageOptions, 'type'>): void {
        this.show(message, { ...options, type: 'success' });
    }

    public error(message: string, options?: Omit<MessageOptions, 'type'>): void {
        this.show(message, { ...options, type: 'error' });
    }

    public warning(message: string, options?: Omit<MessageOptions, 'type'>): void {
        this.show(message, { ...options, type: 'warning' });
    }

    public info(message: string, options?: Omit<MessageOptions, 'type'>): void {
        this.show(message, { ...options, type: 'info' });
    }

    /**
     * 显示确认对话框
     * @param message 提示消息
     * @param onConfirm 确认回调
     * @param onCancel 取消回调
     */
    public confirm(message: string, onConfirm?: () => void, onCancel?: () => void): void {
        // 先安全清理上一次确认框（可能处于隐藏状态），避免 DOM 累积
        if (this.confirmMask) {
            this.confirmMask.remove();
            this.confirmMask = null;
        }

        const confirmBox = createDiv({ className: 'confirm-box' });
        const mask = createDiv({ className: 'confirm-mask' });
        const content = createDiv({ className: 'confirm-content' });
        const messageEl = createDiv({ className: 'confirm-message' });
        const btnGroup = createDiv({ className: 'confirm-btn-group' });
        const confirmBtn = createDiv({ className: 'confirm-btn confirm-ok' });
        const cancelBtn = createDiv({ className: 'confirm-btn confirm-cancel' });

        messageEl.textContent = message;
        confirmBtn.textContent = '确认';
        cancelBtn.textContent = '取消';

        mask.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(0, 0, 0, 0.5);
            display: flex;
            align-items: center;
            justify-content: center;
            z-index: 10000;
            animation: fadeIn 0.2s ease-out;
        `;

        confirmBox.style.cssText = `
            background: #fff;
            border-radius: 8px;
            padding: 24px;
            min-width: 320px;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.15);
            animation: scaleIn 0.2s ease-out;
        `;

        content.style.cssText = `
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 20px;
        `;

        messageEl.style.cssText = `
            font-size: 14px;
            color: #333;
            line-height: 1.5;
            text-align: center;
        `;

        btnGroup.style.cssText = `
            display: flex;
            gap: 12px;
            margin-top: 8px;
        `;

        confirmBtn.style.cssText = `
            padding: 8px 24px;
            background: #1890ff;
            color: #fff;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
            transition: all 0.2s;
        `;

        cancelBtn.style.cssText = `
            padding: 8px 24px;
            background: #f5f5f5;
            color: #666;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
            transition: all 0.2s;
        `;

        confirmBtn.addEventListener('mouseenter', () => {
            confirmBtn.style.background = '#40a9ff';
        });

        confirmBtn.addEventListener('mouseleave', () => {
            confirmBtn.style.background = '#1890ff';
        });

        cancelBtn.addEventListener('mouseenter', () => {
            cancelBtn.style.background = '#e8e8e8';
        });

        cancelBtn.addEventListener('mouseleave', () => {
            cancelBtn.style.background = '#f5f5f5';
        });

        let closed = false;
        const close = (confirmed: boolean) => {
            // 防止重复关闭（快速多次点击按钮会导致 removeChild 报错）
            if (closed) return;
            closed = true;
            confirmBox.style.animation = 'scaleOut 0.2s ease-in';
            mask.style.animation = 'fadeOut 0.2s ease-in';
            setTimeout(() => {
                // 隐藏而非移除，避免 removeChild 错误
                mask.style.display = 'none';
                if (confirmed) {
                    onConfirm?.();
                } else {
                    onCancel?.();
                }
            }, 200);
        };

        confirmBtn.addEventListener('click', () => close(true));
        cancelBtn.addEventListener('click', () => close(false));

        content.appendChild(messageEl);
        btnGroup.appendChild(confirmBtn);
        btnGroup.appendChild(cancelBtn);
        content.appendChild(btnGroup);
        confirmBox.appendChild(content);
        mask.appendChild(confirmBox);
        document.body.appendChild(mask);
        this.confirmMask = mask;

        // 样式只注入一次，避免每次 confirm 都向 head 追加重复的 style 标签
        if (!document.getElementById('confirm-box-styles')) {
            const styleSheet = document.createElement('style');
            styleSheet.id = 'confirm-box-styles';
            styleSheet.textContent = `
                @keyframes fadeIn {
                    from { opacity: 0; }
                    to { opacity: 1; }
                }
                @keyframes fadeOut {
                    from { opacity: 1; }
                    to { opacity: 0; }
                }
                @keyframes scaleIn {
                    from { transform: scale(0.9); opacity: 0; }
                    to { transform: scale(1); opacity: 1; }
                }
                @keyframes scaleOut {
                    from { transform: scale(1); opacity: 1; }
                    to { transform: scale(0.9); opacity: 0; }
                }
            `;
            document.head.appendChild(styleSheet);
        }
    }
}