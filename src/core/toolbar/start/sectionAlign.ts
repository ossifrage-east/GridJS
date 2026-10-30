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
import { EDIT_STATE_CHANGED, WorkspaceKind } from "../../contract/editHost";

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
    /**
     * @type {HTMLDivElement} 上下两行的容器（上行：垂直对齐 + 缩进；下行：水平对齐 / 缩进）
     *
     * 提升为字段而不是构造期的局部变量：`setKind` 要把缩进那一行在两个父节点之间搬，
     * 局部变量出了构造函数就取不到了。
     */
    private column: HTMLDivElement;
    /** @type {HTMLDivElement} 换行和合并按钮容器 */
    private right: HTMLDivElement;
    /**
     * @type {HTMLDivElement} 缩进按钮所在的一行
     *
     * 用户要的是「对齐一组、缩进一组，上下两行」。两个缩进按钮原先住在「段落样式」分区里，
     * 与列表按钮挤在同一行，找起来要横向扫一遍；搬进对齐分区并独占一行后，
     * 一眼就能看出它们与「段落整体位置」是同一类动作。
     *
     * 这一行**在两类工作区里挂的位置不同**（见 {@link setKind}）：
     * 表格里它并入垂直对齐那一行（表格的缩进是单元格内容缩进，与垂直对齐同级），
     * 文档里它单独成行，落在四个对齐按钮下面。
     */
    private indentRow: HTMLDivElement;
    /** @type {Record<string, BtnBase | undefined>} 工具栏按钮集合 */
    private toolbars: Record<string, BtnBase | undefined> = {};
    /** 当前工作区类型（决定「编辑态要不要跳过刷新」与置灰策略，见 setToolbarsState / setKind） */
    private kind: WorkspaceKind = 'sheet';

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

        this.column = createDiv({
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

        this.indentRow = createDiv({
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
        div.append(this.column, this.right);

        
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

        this.column.append(this.top, this.bottom);
        // 默认挂在垂直对齐那一行里（表格口径）；文档下由 setKind 挪到对齐按钮下面单独成行
        this.top.append(this.indentRow);

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
            parentElement: this.indentRow,
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
            parentElement: this.indentRow,
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
        
        /** 段落对齐是单选：点中谁就只亮谁（分散对齐也在这组里） */
        const markAlign = (active: BtnBase): void => {
            const bars = this.toolbars;
            bars.leftBtn?.setActived(bars.leftBtn === active);
            bars.centerBtn?.setActived(bars.centerBtn === active);
            bars.rightBtn?.setActived(bars.rightBtn === active);
            bars.justifyBtn?.setActived(bars.justifyBtn === active);
        };

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
                markAlign(leftBtn);
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
                markAlign(centerBtn);
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
                markAlign(rightBtn);
            }
        });

        // 【第四个图标就是「分散对齐」】图标字体里 `icon-justify`（U+E902）画的是
        // **四行全部撑满版心**——最后一行也拉满，那正是分散对齐的字形；
        // 而 Word 的「两端对齐」图标最后一行是短的。既然字形与命名不一致，
        // 以用户看到的字形为准：这个按钮派发 `distribute`，不再另加一个文字按钮。
        const justifyBtn = new BtnBase({
            parentElement: this.bottom,
            stateSource: this.bus,
            menu: this.menu,
            toDo: 'justify',
            icon: 'icon-justify font15',
            text: '分散对齐',
            title: '分散对齐',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                this.bus.dispatch('setAlign', { value: 'distribute' });
                markAlign(justifyBtn);
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

        this.toolbars = {
            topBtn, middleBtn, bottomBtn, leftBtn, centerBtn, rightBtn, justifyBtn, lineBreakBtn, mergeBtn,
            indentIncreaseBtn, indentDecreaseBtn,
        };
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
     * 按工作区类型收敛分区内容
     *
     * 文档没有「垂直对齐 / 合并单元格」，整行（top + right）直接隐藏；
     * 横向对齐四件在文档里还要解除「编辑态置灰」——文档没有编辑会话，
     * `editing` 常为 true，沿用表格策略会让这四个按钮永远点不动。
     * @param kind - 工作区类型
     */
    public setKind(kind: WorkspaceKind): void {
        this.kind = kind;
        const doc = kind === 'doc';
        this.top.hidden = doc;
        this.right.hidden = doc;
        // 【缩进按钮整行搬家】文档里「垂直对齐 / 换行 / 合并」那一行是隐藏的，
        // 若不动这一行，缩进按钮会跟着 `top` 一起消失 —— 但文档恰恰是**最需要缩进**的载体。
        // append 会自动把它从原父节点摘下，不需要先 remove（重复套用也幂等）。
        if (doc) this.column.append(this.indentRow);
        else this.top.append(this.indentRow);
        const bars = this.toolbars;
        const alignButtons = [bars.leftBtn, bars.centerBtn, bars.rightBtn, bars.justifyBtn];
        for (let i = 0; i < alignButtons.length; i++) {
            alignButtons[i]?.setDisabledOnEditting(!doc);
        }
        // 换到文档后补齐一次状态：表格下 `editing` 为真时是不刷新的（见 setToolbarsState），
        // 切过来那一刻可能正处在编辑态，不补刷会停在上一个载体的高亮
        this.setToolbarsState();
    }

    /**
     * 按当前宿主状态刷新按钮高亮
     *
     * 编辑会话中不刷新：段落级样式在字符编辑时不会变，刷了反而会在退出编辑的瞬间闪一下。
     * @private
     */
    private setToolbarsState() {
        const state = this.bus.snapshot();
        // 表格：编辑会话中段落级样式不会变，跳过刷新（避免退出编辑的瞬间闪一下）。
        // 文档：光标在不同段落间移动时**段落样式确实会变**，不能跳过 ——
        // 否则对齐按钮会一直停在上一个块的高亮上。
        // 【只有表格能在编辑态跳过刷新】文档没有编辑会话：正文一获得焦点 `editing` 就恒为
        // true，若这里无差别 return，缩进 / 对齐按钮会**永远停在进入编辑区那一刻**的状态
        // （表现为「鼠标一进编辑区，减少缩进就再也亮不起来」）。
        if (state.editing && this.kind === 'sheet') return;
        const para = state.para;
        this.toolbars.topBtn.setActived(para.valign === 'top');
        this.toolbars.middleBtn.setActived(para.valign === 'middle' || para.valign === undefined);
        this.toolbars.bottomBtn.setActived(para.valign === 'bottom');
        this.toolbars.leftBtn.setActived(para.align === 'left');
        this.toolbars.centerBtn.setActived(para.align === 'center' || para.align === undefined);
        this.toolbars.rightBtn.setActived(para.align === 'right');
        // 第四个按钮派发的是 distribute（字形与命名的取舍见 init 里的注释）
        this.toolbars.justifyBtn.setActived(para.align === 'distribute');
        this.toolbars.lineBreakBtn.setActived(!!para.wrap);
        // 「减少缩进」到头后置灰：缩进级数已经是 0 时再点也不会有变化，
        // 灰掉比「点了没反应」更容易理解（表格侧 indent 是像素累计值，同样适用）
        const decrease = this.toolbars.indentDecreaseBtn;
        if (decrease) decrease.setDisabled(!((para.indent ?? 0) > 0));
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
    public getElements(): {sectionAlign: HTMLDivElement, top: HTMLDivElement, bottom: HTMLDivElement, right: HTMLDivElement, indentRow: HTMLDivElement} {
        return {sectionAlign: this.sectionAlign, top: this.top, bottom: this.bottom, right: this.right, indentRow: this.indentRow};
    }
}
