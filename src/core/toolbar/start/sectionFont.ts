/*  # web 表格顶部工具栏字体工具集合 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { createDiv } from "../../../utils/dom";
import { FONT_FAMILY_LIST, TOOLBAR_CLASS_NAME, DataEvents, DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE } from "../../constant";
import { BtnBase } from "../../common/button";
import { ColorPicker } from "../../common/colorPicker";
import { Menu } from "../../common/menu";
import { Select } from "../../common/select";
import { DataCollection } from "../../dataArchitecture/dataCollection";
import { CellBorderStyle } from "../../dataArchitecture/cellContent";


/**
 * SectionFont 构造选项接口
 * @interface
 * @property {HTMLElement} [parentElement] - 父元素容器
 * @property {DataCollection} data - 数据集合实例
 * @property {Menu} menu - 菜单实例
 * @property {string} [title] - 工具栏标题
 */
export interface sectionFontOptions {
    parentElement?: HTMLElement;
    data: DataCollection; 
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
    /** @type {DataCollection} 数据集合实例 */
    private data: DataCollection;
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
        this.data  = options.data;
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
                    if (this.data.isEditting) {
                        this.data.setSelectedCharsFont({ fontFamily: todo }); 
                    } else {
                        this.data.setSelectedCellsFont({ fontFamily: todo });
                    }
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
                    const fontSize = Math.floor(parseInt(todo) * window.devicePixelRatio);
                    if (this.data.isEditting) {
                        this.data.setSelectedCharsFont({ fontSize: fontSize }); 
                    } else {
                        this.data.setSelectedCellsFont({ fontSize: fontSize });
                    }                    
                    fontSizeSelect.setText(todo);
                }
            }
        });

        const boldBtn = new BtnBase({
            parentElement: this.bottom,
            data: this.data,
            isDisabledOnEditting: false,
            menu: this.menu,
            toDo: 'bold',
            icon: 'icon-bold font15',
            text: '粗体',
            ishow: true,
            tshow: false,
            layout: 'vertical',
            onClick: (todo) => {
                const bool = !boldBtn.isActived;
                boldBtn.setActived(bool);
                if (this.data.isEditting) {
                    this.data.setSelectedCharsFont({ fontWeight: bool });
                } else {
                    this.data.setSelectedCellsFont({ fontWeight: bool });
                }
            }
        });
        
        const italicBtn = new BtnBase({
            parentElement: this.bottom,
            data: this.data,
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
                if (this.data.isEditting) {
                    this.data.setSelectedCharsFont({ fontStyle: bool });
                } else {
                    this.data.setSelectedCellsFont({ fontStyle: bool });
                }
            }
        });
        
        const underlineBtn = new BtnBase({
            parentElement: this.bottom,
            data: this.data,
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
                if (this.data.isEditting) {
                    this.data.setSelectedCharsFont({ underline: bool });
                } else {
                    this.data.setSelectedCellsFont({ underline: bool });
                }
            }
        });

        const strikethroughBtn = new BtnBase({
            parentElement: this.bottom,
            data: this.data,
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
                if (this.data.isEditting) {
                    this.data.setSelectedCharsFont({ strikethrough: bool });
                } else {
                    this.data.setSelectedCellsFont({ strikethrough: bool });
                }
            }
        });

        const borderBtn = new BtnBase({
            parentElement: this.bottom,
            data: this.data,
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

                    const icon = borderBtn.getIcon();
                    if ((clickMenuItem && todo === 'border-none') || (clickBtn && icon.includes('icon-border-none'))) {
                        this.data.setSelectedCellsBorder('border-none' as CellBorderStyle);   
                    } else if ((clickMenuItem && todo === 'border-all') || (clickBtn && icon.includes('icon-border-all'))) {
                        this.data.setSelectedCellsBorder('border-all' as CellBorderStyle);
                    } else if ((clickMenuItem && todo === 'border-outer') || (clickBtn && icon.includes('icon-border-outer'))) {
                        this.data.setSelectedCellsBorder('border-outer' as CellBorderStyle);
                    } else if ((clickMenuItem && todo === 'border-wide-outer') || (clickBtn && icon.includes('icon-border-wide-outer'))) {
                        this.data.setSelectedCellsBorder('border-wide-outer' as CellBorderStyle);
                    } else if ((clickMenuItem && todo === 'border-bottom') || (clickBtn && icon.includes('icon-border-bottom'))) {
                        this.data.setSelectedCellsBorder('border-bottom' as CellBorderStyle);
                    } else if ((clickMenuItem && todo === 'border-top') || (clickBtn && icon.includes('icon-border-top'))) {
                        this.data.setSelectedCellsBorder('border-top' as CellBorderStyle);
                    } else if ((clickMenuItem && todo === 'border-left') || (clickBtn && icon.includes('icon-border-left'))) {
                        this.data.setSelectedCellsBorder('border-left' as CellBorderStyle);
                    } else if ((clickMenuItem && todo === 'border-right') || (clickBtn && icon.includes('icon-border-right'))) {
                        this.data.setSelectedCellsBorder('border-right' as CellBorderStyle);
                    }
                }
            }
        });

        const paintBucketBtn = new BtnBase({
            parentElement: this.bottom,
            data: this.data,
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
                                this.data.setSelectedCellsFont({ backgroundColor: color !== 'none' ? color : 'white' });
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
                        this.data.setSelectedCellsFont({ backgroundColor: color !== 'none' ? color : 'white' });
                    }
                }
            }
        });
        
        const textColorBtn = new BtnBase({
            parentElement: this.bottom,
            data: this.data,
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
                                    if (this.data.isEditting) {
                                        this.data.setSelectedCharsFont({ fontColor: color !== 'none' ? color : 'black' }); 
                                    } else {
                                        this.data.setSelectedCellsFont({ fontColor: color !== 'none' ? color : 'black' });
                                    }
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
                        if (this.data.isEditting) {
                            this.data.setSelectedCharsFont({ fontColor: color !== 'none' ? color : 'black' }); 
                        } else {
                            this.data.setSelectedCellsFont({ fontColor: color !== 'none' ? color : 'black' });
                        }
                    }
                }
            }
        });
        this.toolbars = {fontSelect, fontSizeSelect, boldBtn, italicBtn, underlineBtn, strikethroughBtn, paintBucketBtn, textColorBtn};
        this.setupDataListeners();
    }

    /**
     * 根据当前激活单元格的数据设置工具栏按钮状态
     * @private
     */
    private setToolbarsState() {
        const cell = this.data.values.find(v => v.cell === this.data.activedCell);
        if (this.data.isEditting) {
            const cursorPosition = this.data.cursorAbsolutePosition - 1;
            const char = cell?.chars?.[cursorPosition];
            (this.toolbars.fontSelect as Select).setText(char?.fontFamily || DEFAULT_FONT_FAMILY);
            (this.toolbars.fontSizeSelect as Select).setText((char?.fontSize || DEFAULT_FONT_SIZE).toString());
            (this.toolbars.boldBtn as BtnBase).setActived(char?.fontWeight);
            (this.toolbars.italicBtn as BtnBase).setActived(char?.fontStyle);
            (this.toolbars.underlineBtn as BtnBase).setActived(char?.underline);
            (this.toolbars.strikethroughBtn as BtnBase).setActived(char?.strikethrough);
            (this.toolbars.textColorBtn as BtnBase).setSpanColor(char?.fontColor || 'black');
            
        } else {
            (this.toolbars.fontSelect as Select).setText(cell?.fontFamily || DEFAULT_FONT_FAMILY);
            (this.toolbars.fontSizeSelect as Select).setText((cell?.fontSize || DEFAULT_FONT_SIZE).toString());
            (this.toolbars.boldBtn as BtnBase).setActived(cell?.fontWeight);
            (this.toolbars.italicBtn as BtnBase).setActived(cell?.fontStyle);
            (this.toolbars.underlineBtn as BtnBase).setActived(cell?.underline);
            (this.toolbars.strikethroughBtn as BtnBase).setActived(cell?.strikethrough);
            (this.toolbars.paintBucketBtn as BtnBase).setSpanColor(cell?.backgroundColor || 'white');
            (this.toolbars.textColorBtn as BtnBase).setSpanColor(cell?.fontColor || 'black');
        }
    }
    
    /**
     * 设置数据事件监听器
     * 监听激活单元格变更和编辑状态变更事件，自动更新工具栏按钮状态
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