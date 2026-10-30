/*  # web 表格顶部工具栏对齐工具集合 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目描述: 基于 HTML5 Canvas 的 Web 电子表格组件，无需服务端支持即可在浏览器中创建、读取、写入 Excel 文件。
    # 项目目标: 工具栏的「对齐方式」分区 —— 垂直/水平对齐、缩进、换行、合并单元格。
    # 本文件职责: 只负责**外观与交互**；样式与数据改动一律经 {@link CommandBus} 派发。
    #             「一次缩进多宽」曾以 getIndentStep 的形式住在这里（依赖字号、单双字节、缩放），
    #             那是载体相关知识，已随解耦搬到 SheetEditHost —— 本文件现在只说「+1 / -1」。
    # 依赖方向: contract/commandBus + common/*，**不再 import dataArchitecture**。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/


import { createDiv } from "../../../utils/dom";
import { TOOLBAR_CLASS_NAME } from "../../constant";
import { BtnBase } from "../../common/button";
import { Menu } from "../../common/menu";
import { CommandBus } from "../../contract/commandBus";
import { EDIT_STATE_CHANGED } from "../../contract/editHost";

/**
 * SectionAlign 构造选项接口
 * @interface
 * @property {HTMLElement} [parentElement] - 父元素容器
 * @property {CommandBus} bus - 命令总线（编辑宿主的状态与命令入口）
 * @property {Menu} menu - 菜单实例
 * @property {string} [title] - 工具栏标题
 */
export interface sectionAlignOptions {
    parentElement?: HTMLElement;
    bus: CommandBus;
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
    /** @type {CommandBus} 命令总线（外壳与宿主之间唯一的接触面） */
    private bus: CommandBus;
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
        this.bus = options.bus;
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
            stateSource: this.bus,            
            menu: this.menu,
            toDo: 'top',
            icon: 'icon-align-top font15',
            text: '顶部对齐',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                this.bus.dispatch('setVAlign', { value: 'top' });
                topBtn.setActived(true);
                middleBtn.setActived(false);
                bottomBtn.setActived(false);
            }
        });

        const middleBtn = new BtnBase({
            parentElement: this.top,
            stateSource: this.bus,
            menu: this.menu,
            toDo: 'middle',
            icon: 'icon-align-middle font15',
            text: '垂直居中对齐',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                this.bus.dispatch('setVAlign', { value: 'middle' });
                topBtn.setActived(false);
                middleBtn.setActived(true);
                bottomBtn.setActived(false);
            }
        });

        const bottomBtn = new BtnBase({
            parentElement: this.top,
            stateSource: this.bus,
            menu: this.menu,
            toDo: 'bottom',
            icon: 'icon-align-bottom font15',
            text: '底部对齐',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                this.bus.dispatch('setVAlign', { value: 'bottom' });
                topBtn.setActived(false);
                middleBtn.setActived(false);
                bottomBtn.setActived(true);
            }
        });

        const indentIncreaseBtn = new BtnBase({
            parentElement: this.top,
            stateSource: this.bus,
            menu: this.menu,
            toDo: 'indentIncrease',
            icon: 'icon-text-indent-left font15',
            text: '缩进增加',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                // 步长由宿主按「当前字号 + 单双字节 + 缩放」现算，外壳只说方向
                this.bus.dispatch('indentStep', { value: 1 });
            }
        });

        const indentDecreaseBtn = new BtnBase({
            parentElement: this.top,
            stateSource: this.bus,
            menu: this.menu,
            toDo: 'indentDecrease',
            icon: 'icon-text-indent-right font15',
            text: '缩进减少',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                this.bus.dispatch('indentStep', { value: -1 });
            }
        }); 
        
        const leftBtn = new BtnBase({
            parentElement: this.bottom,
            stateSource: this.bus,
            menu: this.menu,
            toDo: 'left',
            icon: 'icon-text-left font15',
            text: '左对齐',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                this.bus.dispatch('setAlign', { value: 'left' });
                leftBtn.setActived(true);
                centerBtn.setActived(false);
                rightBtn.setActived(false);
                justifyBtn.setActived(false);
            }
        });
        
        const centerBtn = new BtnBase({
            parentElement: this.bottom,
            stateSource: this.bus,
            menu: this.menu,
            toDo: 'center',
            icon: 'icon-text-center font15',
            text: '水平居中对齐',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                this.bus.dispatch('setAlign', { value: 'center' });
                leftBtn.setActived(false);
                centerBtn.setActived(true);
                rightBtn.setActived(false);
                justifyBtn.setActived(false);
            }
        });

        const rightBtn = new BtnBase({
            parentElement: this.bottom,
            stateSource: this.bus,
            menu: this.menu,
            toDo: 'right',
            icon: 'icon-text-right font15',
            text: '右对齐',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                this.bus.dispatch('setAlign', { value: 'right' });
                leftBtn.setActived(false);
                centerBtn.setActived(false);
                rightBtn.setActived(true);
                justifyBtn.setActived(false);
            }
        });

        const justifyBtn = new BtnBase({
            parentElement: this.bottom,
            stateSource: this.bus,
            menu: this.menu,
            toDo: 'justify',
            icon: 'icon-justify font15',
            title: '两端对齐',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                this.bus.dispatch('setAlign', { value: 'justify' });
                leftBtn.setActived(false);
                centerBtn.setActived(false);
                rightBtn.setActived(false);
                justifyBtn.setActived(true);
            }
        });

        const lineBreakBtn = new BtnBase({
            parentElement: this.right,
            stateSource: this.bus,
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
                this.bus.dispatch('toggleWrap', { toggle: bool });
            }
        });

        const mergeBtn = new BtnBase({
            parentElement: this.right,
            stateSource: this.bus,
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
        this.setupStateListeners();
    }
    
    /**
     * 处理合并单元格操作
     * @param {string} todo - 操作类型标识（'merge' | 'merge-row' | 'merge-col'）
     * @private
     */
    private handleMergeCells(todo: string) {
        switch (todo) {
            case 'merge':
                this.bus.dispatch('toggleMergeCells');
                break;
            case 'merge-row':
                this.bus.dispatch('toggleMergeCells', { value: 'row' });
                break;
            case 'merge-col':
                this.bus.dispatch('toggleMergeCells', { value: 'col' });
                break;
        }
    }

    /**
     * 按当前宿主状态刷新按钮高亮
     *
     * 编辑会话中不刷新：段落级样式在字符编辑时不会变，刷了反而会在退出编辑的瞬间闪一下。
     * @private
     */
    private setToolbarsState() {
        const state = this.bus.snapshot();
        if (state.editing) return;
        const para = state.para;
        this.toolbars.topBtn.setActived(para.valign === 'top');
        this.toolbars.middleBtn.setActived(para.valign === 'middle' || para.valign === undefined);
        this.toolbars.bottomBtn.setActived(para.valign === 'bottom');
        this.toolbars.leftBtn.setActived(para.align === 'left');
        this.toolbars.centerBtn.setActived(para.align === 'center' || para.align === undefined);
        this.toolbars.rightBtn.setActived(para.align === 'right');
        this.toolbars.justifyBtn.setActived(para.align === 'justify');
        this.toolbars.lineBreakBtn.setActived(!!para.wrap);
    }

    /**
     * 订阅「状态可能已变化」
     * @private
     */
    private setupStateListeners() {
        this.bus.on(EDIT_STATE_CHANGED, () => this.setToolbarsState());
        this.setToolbarsState();
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
