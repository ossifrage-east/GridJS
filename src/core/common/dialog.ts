/**
 * 通用 Dialog 组件
 * 提供对话框容器和对话框内容的管理功能
 *
 * @module Dialog
 * @author 东方鹗
 * @see
 * # B站: https://space.bilibili.com/194359739
 * # 知乎: https://www.zhihu.com/people/eastossifrage
 * # CSDN: https://blog.csdn.net/os373
 */

import { createDiv } from "../../utils/dom";
import { EventEmitter } from "../../utils/eventEmitter";

export interface DialogOptions {
     /** 父元素 */
    parentElement?: HTMLElement | null;
}

export class Dialog extends EventEmitter {
    private options: DialogOptions;
    private dialogElement: HTMLElement;
    private contentElement: HTMLElement;
    private closeElement: HTMLElement | null;
    private duration: number;
    private showClose: boolean;
    private onClose: () => void | null;

    constructor(options: DialogOptions) {
        super();
        this.options = options;
        this.dialogElement = createDiv(
            {
                style: {
                    position: 'absolute',
                    top: '50%',
                    left: '50%',
                    transform: 'translate(-50%, -50%)',
                    zIndex: '99',
                    backgroundColor: 'white',
                    padding: '20px',
                    borderRadius: '10px',
                    boxShadow: '0 0 10px rgba(0, 0, 0, 0.5)',
                }
            }
        );
        options.parentElement?.appendChild(this.dialogElement);
    }
}
