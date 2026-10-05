/*  # web 表格顶部工具栏对齐工具集合 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/


import { createDiv } from "../../../utils/dom";
import { DataEvents, TOOLBAR_CLASS_NAME } from "../../constant";
import { DataCollection } from "../../dataArchitecture/dataCollection";
import { BtnBase } from "../../common/button";
import { Menu } from "../../common/menu";
import { Cell, TextAlign, VerticalAlign } from "../../dataArchitecture/cell";

/**
 * SectionAlign 构造选项接口
 * @interface
 * @property {HTMLElement} [parentElement] - 父元素容器
 * @property {DataCollection} data - 数据集合实例
 * @property {Menu} menu - 菜单实例
 * @property {string} [title] - 工具栏标题
 */
export interface sectionAlignOptions {
    parentElement?: HTMLElement;
    data: DataCollection;
    menu: Menu;
    title?: string;
}

/**
 * 工具栏对齐区域类
 * 提供单元格垂直对齐、水平对齐、缩进、换行和合并等功能按钮
 * @class
 */
export class SectionAlign {
    /** @type {HTMLDivElement} 对齐区域容器元素 */
    private sectionAlign: HTMLDivElement;
    /** @type {DataCollection} 数据集合实例 */
    private data: DataCollection;
    /** @type {Menu} 菜单实例 */
    private menu: Menu;
    /** @type {HTMLDivElement} 垂直对齐按钮容器 */
    private top: HTMLDivElement;
    /** @type {HTMLDivElement} 水平对齐按钮容器 */
    private bottom: HTMLDivElement;
    /** @type {HTMLDivElement} 换行和合并按钮容器 */
    private right: HTMLDivElement;
    /** @type {Record<string, BtnBase | undefined>} 工具栏按钮集合 */
    private toolbars: Record<string, BtnBase | undefined> = {};

    /**
     * 创建 SectionAlign 实例
     * @param {sectionAlignOptions} options - 构造选项
     */
    constructor(options: sectionAlignOptions) {
        this.sectionAlign = createDiv({
            className: TOOLBAR_CLASS_NAME.SECTION,
            
        });

        if (options.parentElement) {
            options.parentElement.append(this.sectionAlign);
        }
        this.data = options.data;
        this.menu = options.menu;
        
        const div = createDiv({
            style: { 
                display: 'flex',
                flexDirection: 'row', 
                gap: '4px' 
            }
        });

        const _div = createDiv({
            style: { 
                display: 'flex',
                flexDirection: 'column', 
                gap: '4px' 
            }
        });

        this.right = createDiv({
            style: { 
                display: 'flex',
                flexDirection: 'row', 
                gap: '4px' 
            }
        });

        
        const line = createDiv({
            className: TOOLBAR_CLASS_NAME.SECTION_LINE, 
        });
        this.sectionAlign.append(line, div);
        div.append(_div, this.right);

        
        this.top = createDiv({
            style: {
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                textAlign: 'center'
            }
        });
        this.bottom = createDiv({
            style: {
                display:  'flex',
                alignItems : 'center',
                justifyContent: 'center',
                textAlign: 'center', 
                gap: '4px' 
            }
        });

        _div.append(this.top, this.bottom);

        this.init();
    }    

    /**
     * 初始化工具栏按钮
     * 创建垂直对齐、水平对齐、缩进、换行和合并等功能按钮
     * @public
     */
    public init() {        
        /*------------------ 对齐方式分区 ------------------*/        
        const topBtn = new BtnBase({
            parentElement: this.top,
            data: this.data,            
            menu: this.menu,
            toDo: 'top',
            icon: 'icon-align-top font15',
            text: '顶部对齐',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                this.data.setSelectedCellsFont({ alignItems: 'top' });
                topBtn.setActived(true);
                middleBtn.setActived(false);
                bottomBtn.setActived(false);
            }
        });

        const middleBtn = new BtnBase({
            parentElement: this.top,
            data: this.data,
            menu: this.menu,
            toDo: 'middle',
            icon: 'icon-align-middle font15',
            text: '垂直居中对齐',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                this.data.setSelectedCellsFont({ alignItems: 'middle' });
                topBtn.setActived(false);
                middleBtn.setActived(true);
                bottomBtn.setActived(false);
            }
        });

        const bottomBtn = new BtnBase({
            parentElement: this.top,
            data: this.data,
            menu: this.menu,
            toDo: 'bottom',
            icon: 'icon-align-bottom font15',
            text: '底部对齐',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                this.data.setSelectedCellsFont({ alignItems: 'bottom' });
                topBtn.setActived(false);
                middleBtn.setActived(false);
                bottomBtn.setActived(true);
            }
        });

        const indentIncreaseBtn = new BtnBase({
            parentElement: this.top,
            data: this.data,
            menu: this.menu,
            toDo: 'indentIncrease',
            icon: 'icon-text-indent-left font15',
            text: '缩进增加',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                // for (const cell of this.sheetCanvas.data.getSelectedCells()) {
                //     const cellIndex = this.sheetCanvas.data.cells.findIndex(v => v.cell === cell);            
                //     if (cellIndex >= 0) {
                //         this.sheetCanvas.data.cells[cellIndex].indent += 4;
                //     } else {
                //         this.sheetCanvas.data.cells.push({
                //             cell: cell,
                //             indent: 4
                //         });
                //     }
                // }
            }
        });

        const indentDecreaseBtn = new BtnBase({
            parentElement: this.top,
            data: this.data,
            menu: this.menu,
            toDo: 'indentDecrease',
            icon: 'icon-text-indent-right font15',
            text: '缩进减少',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                // for (const cell of this.sheetCanvas.data.getSelectedCells()) {
                //     const cellIndex = this.sheetCanvas.data.cells.findIndex(v => v.cell === cell);            
                //     if (cellIndex >= 0) {
                //         this.sheetCanvas.data.cells[cellIndex].indent -= 4;
                //     } else {
                //         this.sheetCanvas.data.cells.push({
                //             cell: cell,
                //             indent: -4
                //         });
                //     }
                // }
            }
        }); 
        
        const leftBtn = new BtnBase({
            parentElement: this.bottom,
            data: this.data,
            menu: this.menu,
            toDo: 'left',
            icon: 'icon-text-left font15',
            text: '左对齐',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                this.data.setSelectedCellsFont({ textAlign: 'left' });
                leftBtn.setActived(true);
                centerBtn.setActived(false);
                rightBtn.setActived(false);
                justifyBtn.setActived(false);
            }
        });
        
        const centerBtn = new BtnBase({
            parentElement: this.bottom,
            data: this.data,
            menu: this.menu,
            toDo: 'center',
            icon: 'icon-text-center font15',
            text: '水平居中对齐',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                this.data.setSelectedCellsFont({ textAlign: 'center' });
                leftBtn.setActived(false);
                centerBtn.setActived(true);
                rightBtn.setActived(false);
                justifyBtn.setActived(false);
            }
        });

        const rightBtn = new BtnBase({
            parentElement: this.bottom,
            data: this.data,
            menu: this.menu,
            toDo: 'right',
            icon: 'icon-text-right font15',
            text: '右对齐',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                this.data.setSelectedCellsFont({ textAlign: 'right' });
                leftBtn.setActived(false);
                centerBtn.setActived(false);
                rightBtn.setActived(true);
                justifyBtn.setActived(false);
            }
        });

        const justifyBtn = new BtnBase({
            parentElement: this.bottom,
            data: this.data,
            menu: this.menu,
            toDo: 'justify',
            icon: 'icon-justify font15',
            title: '两端对齐',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                this.data.setSelectedCellsFont({ textAlign: 'justify' });
                leftBtn.setActived(false);
                centerBtn.setActived(false);
                rightBtn.setActived(false);
                justifyBtn.setActived(true);
            }
        });

        const lineBreakBtn = new BtnBase({
            parentElement: this.right,
            data: this.data,
            menu: this.menu,
            toDo: 'lineBreak',
            icon: 'icon-line-break font20',
            title: '换行',
            ishow: true,
            tshow: false,
            layout: 'horizontal',
            onClick: (todo) => {
                const bool = !lineBreakBtn.isActived;
                lineBreakBtn.setActived(bool);
                this.data.setSelectedCellsFont({ wrap: bool });
            }
        });

        const mergeBtn = new BtnBase({
            parentElement: this.right,
            data: this.data,
            menu: this.menu,
            toDo: 'merge',
            icon: 'icon-merge font20',
            title: '合并单元格',
            ishow: true,
            tshow: true,
            layout: 'horizontal',
            separate: true,
            ashow: true,
            menuContent: {
                items: [
                    { todo: 'merge', icon: 'icon-merge font16', text: '合并后居中'},
                    'separator',
                    { todo: 'merge-row', icon: 'icon-merge font16', text: '按行合并'},
                    { todo: 'merge-col', icon: 'icon-merge font16', text: '按列合并'}
                ]
            },
            onClick: (todo) => {
                if (todo.includes('MenuOpen')) {
                    this.menu.openMenu(mergeBtn.getElements().container);
                    this.menu.addMenuElement(mergeBtn.getElements().menuContainer);
                } else {
                    this.handleMergeCells(todo);
                }
            }
        });

        this.toolbars = { topBtn, middleBtn, bottomBtn, leftBtn, centerBtn, rightBtn, justifyBtn, lineBreakBtn, mergeBtn };
        this.setupDataListeners();
    }
    
    /**
     * 处理合并单元格操作
     * @param {string} todo - 操作类型标识
     * @private
     */
    private handleMergeCells(todo: string) {
        switch (todo) {
            case 'merge':
                this.data.toggleMergeCells();
                break;
            case 'merge-row':
                this.data.toggleMergeCells('row');
                break;
            case 'merge-col':
                this.data.toggleMergeCells('col');
                break;
        }
    }

    /**
     * 设置工具栏状态
     * @private
     */
    private setToolbarsState() {
        const cell = this.data.values.find(v => v.cell === this.data.activedCell);
        if (!this.data.isEditting) {
            this.toolbars.topBtn.setActived(cell?.alignItems === 'top');
            this.toolbars.middleBtn.setActived(cell?.alignItems === 'middle' || cell?.alignItems === undefined);
            this.toolbars.bottomBtn.setActived(cell?.alignItems === 'bottom');
            this.toolbars.leftBtn.setActived(cell?.textAlign === 'left');
            this.toolbars.centerBtn.setActived(cell?.textAlign === 'center' || cell?.textAlign === undefined);
            this.toolbars.rightBtn.setActived(cell?.textAlign === 'right');
            this.toolbars.justifyBtn.setActived(cell?.textAlign === 'justify');
            this.toolbars.lineBreakBtn.setActived(cell?.wrap);
        }
        
    }

    /**
     * 设置数据事件监听器
     * @private
     */
    private setupDataListeners() {
        this.data.syncValuesFromDB().then(() => { 
            this.setToolbarsState();
        });
        this.data.on(DataEvents.ACTIVED_CELL_CHANGED, () => this.setToolbarsState());
        // 撤销/重做回写后同步按钮状态（活动单元格未变化时无 ACTIVED_CELL_CHANGED 事件）
        this.data.on(DataEvents.SNAPSHOT_RESTORED, () => this.setToolbarsState());
        this.data.on(DataEvents.IS_EDITTING_CHANGED, (_, isEditting) => {
            if(!isEditting) this.setToolbarsState();
        });
        this.data.on(DataEvents.CURSOR_STATE_CHANGED, () => this.setToolbarsState());
    }

    /**
     * 获取所有工具栏按钮
     * @returns {Record<string, BtnBase | undefined>} 工具栏按钮集合
     * @public
     */
    public getToolbars(): Record<string, BtnBase | undefined> {
        return this.toolbars;
    }
    
    /**
     * 获取工具栏元素
     * @returns {{sectionAlign: HTMLDivElement, top: HTMLDivElement, bottom: HTMLDivElement, right: HTMLDivElement}} 工具栏元素对象
     * @public
     */
    public getElements(): {sectionAlign: HTMLDivElement, top: HTMLDivElement, bottom: HTMLDivElement, right: HTMLDivElement} {
        return {sectionAlign: this.sectionAlign, top: this.top, bottom: this.bottom, right: this.right};
    }
}