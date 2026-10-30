/*  # web 表格顶部工具栏字体工具集合 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目描述: 基于 HTML5 Canvas 的 Web 电子表格组件，无需服务端支持即可在浏览器中创建、读取、写入 Excel 文件。
    # 项目目标: 工具栏的「字体样式」分区 —— 字体、字号、粗体、斜体、下划线、删除线、边框、油漆桶、文字色。
    # 本文件职责: 只负责**外观与交互**：建按钮、开菜单、按状态高亮；
    #             真正「改什么」一律通过 {@link CommandBus} 派发，改完再按快照回读状态。
    # 依赖方向: contract/commandBus + common/*，**不再 import dataArchitecture**。
    #           改造前本文件有 7 处 `this.data.isEditting ? setSelectedCharsFont : setSelectedCellsFont`
    #           的同构分支，现已全部收敛到宿主的 applyCharStyle 一处。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { createDiv } from "../../../utils/dom";
import { FONT_FAMILY_LIST, TOOLBAR_CLASS_NAME, DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE } from "../../constant";
import { BtnBase } from "../../common/button";
import { ColorPicker } from "../../common/colorPicker";
import { Menu } from "../../common/menu";
import { Select } from "../../common/select";
import { CommandBus } from "../../contract/commandBus";
import { EDIT_STATE_CHANGED } from "../../contract/editHost";


/**
 * SectionFont 构造选项接口
 * @interface
 * @property {HTMLElement} [parentElement] - 父元素容器
 * @property {CommandBus} bus - 命令总线（编辑宿主的状态与命令入口）
 * @property {Menu} menu - 菜单实例
 * @property {string} [title] - 工具栏标题
 */
export interface sectionFontOptions {
    parentElement?: HTMLElement;
    bus: CommandBus;
    menu: Menu;
    title?: string;
}

/**
 * 工具栏字体区域类
 * 提供字体选择、字号设置、粗体、斜体、下划线、删除线、边框、油漆桶和文本颜色等功能按钮
 * @class
 */
export class SectionFont {
    /** @type {HTMLDivElement} 字体样式区域容器元素 */
    private sectionFont: HTMLDivElement;
    /** @type {HTMLDivElement} 字体和字号选择容器 */
    private top: HTMLDivElement;
    /** @type {HTMLDivElement} 字体样式按钮容器 */
    private bottom: HTMLDivElement;
    /** @type {CommandBus} 命令总线（外壳与宿主之间唯一的接触面） */
    private bus: CommandBus;
    /** @type {Menu} 菜单实例 */
    private menu: Menu;
    /** @type {Record<string, (Select | BtnBase) | undefined>} 工具栏按钮集合 */
    private toolbars: Record<string, (Select | BtnBase) | undefined> = {};

    /**
     * 创建 SectionFont 实例
     * @param {sectionFontOptions} options - 构造选项
     */
    constructor(options: sectionFontOptions) {
        this.sectionFont = createDiv({
            className: TOOLBAR_CLASS_NAME.SECTION,
            
        });

        if (options.parentElement) {
            options.parentElement.append(this.sectionFont);
        }
        this.bus  = options.bus;
        this.menu = options.menu;
        
        const __div = createDiv({
            style: { 
                display: 'flex',
                flexDirection: 'column', 
                gap: '4px' 
            }
        });

        
        const line = createDiv({
            className: TOOLBAR_CLASS_NAME.SECTION_LINE,
        });
        this.sectionFont.append(line, __div);
        

        
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

        __div.append(this.top, this.bottom);
        this.init();
    }

    /**
     * 初始化工具栏按钮
     * 创建字体选择、字号设置、粗体、斜体、下划线、删除线、边框、油漆桶和文本颜色等功能按钮
     * @public
     */
    public init() {
        /*------------------ 字体样式分区 ------------------*/                
        // detectSystemFonts().then((availableFonts: string[]) => {
        //     const items = availableFonts.map((font) => {
        //         const fontInfo = fontsZhData.find((item) => item.en === font);
        //         return {
        //             todo: fontInfo ? fontInfo.en : font,
        //             text: fontInfo?.cn || font,
        //         };
        //     });
            
        //     items.unshift({ todo: 'HarmonyOS Sans SC Thin', text: '鸿蒙黑体细' });  // 将自定义的字体添加到字体列表的最前面
        //     items.unshift({ todo: 'HarmonyOS Sans SC Light', text: '鸿蒙黑体轻' });  // 将自定义的字体添加到字体列表的最前面
        //     items.unshift({ todo: 'HarmonyOS Sans SC Black', text: '鸿蒙黑体黑' });  // 将自定义的字体添加到字体列表的最前面
        //     items.unshift({ todo: 'HarmonyOS Sans SC Regular', text: '鸿蒙黑体' });  // 将自定义的字体添加到字体列表的最前面
        //     console.log(items);
        // });
        
        const fontItems = FONT_FAMILY_LIST;
        const fontSelect = new Select({
            parentElement: this.top,
            inputID: 'fontSelect',
            toDo: 'fontSelect',
            title: '字体',
            menuContent: {
                items: fontItems,
            },
            onClick: (todo) => {
                this.menu.closeMenu();
                if (todo.includes('MenuOpen')) {
                    this.menu.openMenu(fontSelect.getElements().container);
                    this.menu.addMenuElement(fontSelect.getElements().menuContainer);
                    const spans = fontSelect.getElements().menuContainer.querySelectorAll('.text');
                    Array.from(spans).forEach((span) => {  // 遍历下拉菜单中的每个字体项，并设置样式
                        (span as HTMLElement).style.fontFamily = fontItems.find(item => item.text === (span as HTMLElement).innerText)?.todo || '';
                    })
                }
                if (!todo.includes('Menu')) {  // 点击菜单时，只要不是下拉菜单打开事项，则改变显示内容
                    this.bus.dispatch('setFontFamily', { value: todo });
                    fontSelect.setText(todo);
                }
            }
        });

        const fontSizeSelect = new Select({
            parentElement: this.top,
            containerID: 'fontSize',
            inputID: 'fontSizeSelect',
            toDo: 'fontSizeSelect',
            title: '字号',
            menuContent: {
                items: [
                    { todo: '6', text: '6' },
                    { todo: '7', text: '7' },
                    { todo: '8', text: '8' },
                    { todo: '9', text: '9' },
                    { todo: '10', text: '10' },
                    { todo: '12', text: '12' },
                    { todo: '14', text: '14' },
                    { todo: '16', text: '16' },
                    { todo: '18', text: '18' },
                    { todo: '20', text: '20' },
                    { todo: '24', text: '24' },
                    { todo: '28', text: '28' },
                    { todo: '32', text: '32' },
                    { todo: '36', text: '36' },
                    { todo: '40', text: '40' },
                ],
            },
            onClick: (todo) => {
                this.menu.closeMenu();
                if (todo.includes('MenuOpen')) {
                    this.menu.openMenu(fontSizeSelect.getElements().container);
                    this.menu.addMenuElement(fontSizeSelect.getElements().menuContainer);
                }
                if (!todo.includes('Menu')) {  // 点击菜单时，只要不是下拉菜单打开事项，则改变显示内容
                    // 菜单里是磅值，宿主存的是像素（乘 devicePixelRatio）；换算留在外壳，
                    // 因为它是「菜单项怎么写」的问题，不是「编辑怎么改」的问题
                    const fontSize = Math.floor(parseInt(todo) * window.devicePixelRatio);
                    this.bus.dispatch('setFontSize', { value: fontSize });
                    fontSizeSelect.setText(todo);
                }
            }
        });

        const boldBtn = new BtnBase({
            parentElement: this.bottom,
            stateSource: this.bus,
            isDisabledOnEditting: false,
            menu: this.menu,
            toDo: 'bold',
            icon: 'icon-bold font15',
            text: '粗体',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                // 【先算再派发】宿主执行完会同步广播状态变化，setToolbarsState 会按真实值
                // 把按钮刷一遍；若在派发之后再读 isActived，读到的是「已刷新后的值」，
                // 再取反就把按钮翻回去了（表现为点一次亮、下一次点不灭）
                const bool = !boldBtn.isActived;
                boldBtn.setActived(bool);
                this.bus.dispatch('toggleBold', { toggle: bool });
            }
        });
        
        const italicBtn = new BtnBase({
            parentElement: this.bottom,
            stateSource: this.bus,
            isDisabledOnEditting: false,
            menu: this.menu,
            toDo: 'italic',
            icon: 'icon-italic font15',
            text: '斜体',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                const bool = !italicBtn.isActived;
                italicBtn.setActived(bool);
                this.bus.dispatch('toggleItalic', { toggle: bool });
            }
        });
        
        const underlineBtn = new BtnBase({
            parentElement: this.bottom,
            stateSource: this.bus,
            isDisabledOnEditting: false,
            menu: this.menu,
            toDo: 'underline',
            icon: 'icon-underline font15',
            text: '下划线',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                const bool = !underlineBtn.isActived;
                underlineBtn.setActived(bool);
                this.bus.dispatch('toggleUnderline', { toggle: bool });
            }
        });

        const strikethroughBtn = new BtnBase({
            parentElement: this.bottom,
            stateSource: this.bus,
            isDisabledOnEditting: false,
            menu: this.menu,
            toDo: 'strikethrough',
            icon: 'icon-strikethrough font15',
            text: '删除线',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                const bool = !strikethroughBtn.isActived;
                strikethroughBtn.setActived(bool);
                this.bus.dispatch('toggleStrike', { toggle: bool });
            }
        });

        const borderBtn = new BtnBase({
            parentElement: this.bottom,
            stateSource: this.bus,
            menu: this.menu,
            toDo: 'border',
            icon: 'icon-border-none font15',
            title: '边框',
            ishow: true,
            tshow: true,
            layout: 'horizontal',
            separate: true,
            ashow: true,
            menuContent: {
                items: [
                    { todo: 'border-none', icon: 'icon-border-none', text: '无边框' },
                    { todo: 'border-all', icon: 'icon-border-all', text: '所有边框' }, 
                    { todo: 'border-outer', icon: 'icon-border-outer', text: '外侧边框' }, 
                    { todo: 'border-wide-outer', icon: 'icon-border-wide-outer', text: '粗匣边框' },
                    'separator',
                    { todo: 'border-bottom', icon: 'icon-border-bottom', text: '下边框' },
                    { todo: 'border-top', icon: 'icon-border-top', text: '上边框' },
                    { todo: 'border-left', icon: 'icon-border-left', text: '左边框' },
                    { todo: 'border-right', icon: 'icon-border-right', text: '右边框' },
                    'separator'
                ],
            },
            onClick: (todo) => {
                if (todo.includes('MenuOpen')) {
                    this.menu.openMenu(borderBtn.getElements().container);
                    this.menu.addMenuElement(borderBtn.getElements().menuContainer);
                }
                if (!todo.includes('Menu')) {  // 点击菜单时，只要不是下拉菜单打开事项，则改变图标
                    let clickBtn = false, clickMenuItem = false;
                    if (todo !== 'border') {
                        borderBtn.setIcon(todo);
                        clickMenuItem = true;
                    } else {
                        clickBtn = true;
                    }

                    // 八种线型映射成同一个命令：菜单点选优先，其次看按钮当前图标
                    const icon = borderBtn.getIcon();
                    const chosen = clickMenuItem ? todo : (clickBtn ? icon : '');
                    if (chosen === 'border-none' || (chosen || '').includes('icon-border-none')) {
                        this.bus.dispatch('setBorder', { value: 'border-none' });
                    } else if (chosen === 'border-all' || (chosen || '').includes('icon-border-all')) {
                        this.bus.dispatch('setBorder', { value: 'border-all' });
                    } else if (chosen === 'border-outer' || (chosen || '').includes('icon-border-outer')) {
                        this.bus.dispatch('setBorder', { value: 'border-outer' });
                    } else if (chosen === 'border-wide-outer' || (chosen || '').includes('icon-border-wide-outer')) {
                        this.bus.dispatch('setBorder', { value: 'border-wide-outer' });
                    } else if (chosen === 'border-bottom' || (chosen || '').includes('icon-border-bottom')) {
                        this.bus.dispatch('setBorder', { value: 'border-bottom' });
                    } else if (chosen === 'border-top' || (chosen || '').includes('icon-border-top')) {
                        this.bus.dispatch('setBorder', { value: 'border-top' });
                    } else if (chosen === 'border-left' || (chosen || '').includes('icon-border-left')) {
                        this.bus.dispatch('setBorder', { value: 'border-left' });
                    } else if (chosen === 'border-right' || (chosen || '').includes('icon-border-right')) {
                        this.bus.dispatch('setBorder', { value: 'border-right' });
                    }
                }
            }
        });

        const paintBucketBtn = new BtnBase({
            parentElement: this.bottom,
            stateSource: this.bus,
            menu: this.menu,
            toDo: 'paintBucket',
            icon: 'icon-paint-bucket font15',
            title: '油漆桶',
            ishow: true,
            tshow: true,
            layout: 'horizontal',
            separate: true,
            ashow: true,
            isColor: true,
            menuContent: {
                items:  () => {
                        const colorPicker = new ColorPicker({
                            initialColor: 'none',
                            parentElement: document.body,
                            onChange: (color) => {
                                this.menu.closeMenu();
                                paintBucketBtn.setSpanColor(color);
                                this.bus.dispatch('setBackground', { value: color !== 'none' ? color : 'white' });
                            }
                        });
                        return colorPicker.getElements().pickerContainer;
                    }
            },
            onClick: (todo) => {
                this.menu.closeMenu();
                if (todo.includes('MenuOpen')) {
                    this.menu.openMenu(paintBucketBtn.getElements().container);
                    this.menu.addMenuElement(paintBucketBtn.getElements().menuContainer);
                }
                
                if (!todo.includes('Menu')) {
                    const color = paintBucketBtn.getSpanColor();
                    if (color) {
                        this.bus.dispatch('setBackground', { value: color !== 'none' ? color : 'white' });
                    }
                }
            }
        });
        
        const textColorBtn = new BtnBase({
            parentElement: this.bottom,
            stateSource: this.bus,
            isDisabledOnEditting: false,
            menu: this.menu,
            toDo: 'textColor',
            title: '文本颜色',
            icon: 'icon-text-color font15',
            ishow: true,
            tshow: true,
            layout: 'horizontal',
            separate: true,
            ashow: true,
            isColor: true,
            menuContent: {
                items:  () => {
                        const colorPicker = new ColorPicker({
                            initialColor: 'black',
                            parentElement: document.body,
                            onChange: (color) => {
                                if (color) {
                                    this.menu.closeMenu();
                                    textColorBtn.setSpanColor(color);
                                    this.bus.dispatch('setTextColor', { value: color !== 'none' ? color : 'black' });
                                }
                            }
                        });
                        return colorPicker.getElements().pickerContainer;
                    }
            },
            onClick: (todo) => {
                this.menu.closeMenu();
                if (todo.includes('MenuOpen')) {
                    this.menu.openMenu(textColorBtn.getElements().container);
                    this.menu.addMenuElement(textColorBtn.getElements().menuContainer);
                }
                if (!todo.includes('Menu')) {
                    const color = textColorBtn.getSpanColor();
                    if (color) {
                        this.bus.dispatch('setTextColor', { value: color !== 'none' ? color : 'black' });
                    }
                }
            }
        });
        this.toolbars = {fontSelect, fontSizeSelect, boldBtn, italicBtn, underlineBtn, strikethroughBtn, paintBucketBtn, textColorBtn};
        this.setupStateListeners();
    }

    /**
     * 按当前宿主状态刷新按钮外观（单一数据来源：{@link CommandBus.snapshot}）
     *
     * 编辑态与非编辑态的差异**已由宿主在快照里归一化**：编辑态取光标处字符、
     * 非编辑态取活动格代表值，外壳不判断该读哪里。唯一保留的分支是油漆桶——
     * 背景色是单元格属性，编辑会话中无处可挂（{@link SheetEditHost} 的 setBackground 同样只走单元格级）。
     * @private
     */
    private setToolbarsState() {
        const state = this.bus.snapshot();
        const char = state.char;
        (this.toolbars.fontSelect as Select).setText(char.fontFamily || DEFAULT_FONT_FAMILY);
        (this.toolbars.fontSizeSelect as Select).setText((char.fontSize || DEFAULT_FONT_SIZE).toString());
        (this.toolbars.boldBtn as BtnBase).setActived(!!char.bold);
        (this.toolbars.italicBtn as BtnBase).setActived(!!char.italic);
        (this.toolbars.underlineBtn as BtnBase).setActived(!!char.underline);
        (this.toolbars.strikethroughBtn as BtnBase).setActived(!!char.strike);
        (this.toolbars.textColorBtn as BtnBase).setSpanColor(char.color || 'black');
        if (!state.editing) {
            (this.toolbars.paintBucketBtn as BtnBase).setSpanColor(char.background || 'white');
        }
    }
    
    /**
     * 订阅「状态可能已变化」
     *
     * 改造前这里订阅了 4 个表格事件（ACTIVED_CELL_CHANGED / SNAPSHOT_RESTORED /
     * IS_EDITTING_CHANGED / CURSOR_STATE_CHANGED）并各自调一次 `syncValuesFromDB()`。
     * 现在统一为总线的一个事件，且**不再由 UI 触发数据装载**——
     * 那是数据层的职责（`sheets.ts` 已经在做），UI 顺手再读一次库纯属浪费。
     * @private
     */
    private setupStateListeners() {
        this.bus.on(EDIT_STATE_CHANGED, () => this.setToolbarsState());
        // 首帧先按当前（可能还是空的）状态刷一次，避免按钮停留在默认外观
        this.setToolbarsState();
    }

    /**
     * 获取工具栏元素
     * @returns {{sectionFont: HTMLDivElement, top: HTMLDivElement, bottom: HTMLDivElement}} 工具栏元素对象
     * @public
     */
    public getElements() {
        return {sectionFont: this.sectionFont, top: this.top, bottom: this.bottom};
    }

    /**
     * 获取所有工具栏按钮
     * @returns {Record<string, (Select | BtnBase) | undefined>} 工具栏按钮集合
     * @public
     */
    public getToolbars() {
        return this.toolbars;
    }
    
}
