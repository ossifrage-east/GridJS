/*  # web 表格顶部快速访问工具栏 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { createDiv } from "../../utils/dom";
import { UNDO_STACK_CHANGED } from "../../utils/undoManager";
import { TOOLBAR_CLASS_NAME, DataEvents } from "../constant";
import { BtnBase } from "../common/button";
import { Menu } from "../common/menu";
import { DataCollection } from "../dataArchitecture/dataCollection";
import { getSharedPrinter } from "../../utils/printer";
import favicon from "../../assets/images/favicon.png";
import '../style/brand.scss'


export interface QuickAccessOptions {
    parentElement: HTMLElement;
    data: DataCollection;
    menu: Menu;
    title?: string;
}


export class QuickAccess {
    private quickAccess: HTMLDivElement;
    private data: DataCollection;
    private menu: Menu;
    constructor(options: QuickAccessOptions) {
        this.quickAccess = createDiv({
            className: TOOLBAR_CLASS_NAME.QUICK_ACCESS,
        });        
        if (options.parentElement) {
            options.parentElement.append(this.quickAccess);
        }
        this.data = options.data;  
        this.menu = options.menu;
        this.init();
    }

    public init() {        
        /*------------------ 快速访问分区 ------------------*/

        const brand = createDiv({className: TOOLBAR_CLASS_NAME.BRAND});
        const brandImg = new Image();
        brandImg.src = favicon;  // 添加图片地址
        brand.append(brandImg);
        this.quickAccess.append(brand);

        const backBtn = new BtnBase(  // 撤销按钮
            {
                parentElement: this.quickAccess, 
                data: this.data,
                isDisabledOnEditting: false,
                menu: this.menu,
                toDo: 'back',
                icon: 'icon-undo',
                text: '撤销 Ctrl + Z',
                ishow: true,
                layout: 'vertical',
                onClick: (todo) => {
                    this.data.undo();
                }
            }
        );

        const forwardBtn = new BtnBase(  // 恢复按钮
            {
                parentElement: this.quickAccess,
                data: this.data,
                isDisabledOnEditting: false,
                menu: this.menu,
                toDo: 'forward',
                icon: 'icon-redo',
                text: '恢复 Ctrl + Y',
                ishow: true,
                layout: 'vertical',
                onClick: (todo) => {
                    this.data.redo();
                }
            }
        );

        // 撤销/重做按钮可用状态监控：撤销栈/重做栈为空时按钮置灰不可点
        const updateUndoRedoState = () => {
            backBtn.setDisabled(!this.data.undoManager.canUndo());
            forwardBtn.setDisabled(!this.data.undoManager.canRedo());
        };
        updateUndoRedoState();
        this.data.undoManager.on(UNDO_STACK_CHANGED, updateUndoRedoState);
        // 编辑期间按钮服务于编辑内逐字撤销，状态保持不变；编辑结束后按文档栈重新同步
        this.data.on(DataEvents.IS_EDITTING_CHANGED, (_, isEditting: boolean) => {
            if (!isEditting) updateUndoRedoState();
        });

        // 打印组件：「页面」分区的打印设置按钮与打印/打印预览共用同一共享实例，设置互通
        const printer = getSharedPrinter(this.data);

        const printBtn = new BtnBase(  // 打印按钮
            {
                parentElement: this.quickAccess,
                data: this.data,
                menu: this.menu,
                toDo: 'print',
                icon: 'icon-printer',
                text: '打印',
                ishow: true,
                layout: 'vertical',
                onClick: (todo) => {
                    printer.print();
                }
            }
        );

        const printViewBtn = new BtnBase(  // 打印预览按钮
            {
                parentElement: this.quickAccess,
                data: this.data,
                menu: this.menu,
                toDo: 'printView',
                icon: 'icon-printer-view',
                text: '打印预览',
                ishow: true,
                layout: 'vertical',
                onClick: (todo) => {
                    printer.printPreview();
                }
            },
        );
    }

    public getElements() {
        return {
            container: this.quickAccess,
        }
    }
}