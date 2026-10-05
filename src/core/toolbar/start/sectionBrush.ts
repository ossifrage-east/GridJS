/*  # web 表格顶部工具栏格式刷集合 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/


import { createDiv } from "../../../utils/dom";
import { TOOLBAR_CLASS_NAME, DataEvents } from "../../constant";
import { DataCollection } from "../../dataArchitecture/dataCollection";
import { Menu } from "../../common/menu";
import { BtnBase } from "../../common/button";

/**
 * SectionBrush 构造选项接口
 * @interface
 * @property {HTMLElement} [parentElement] - 父元素容器
 * @property {DataCollection} data - 数据集合实例
 * @property {Menu} menu - 菜单实例
 * @property {string} [title] - 工具栏标题
 */
export interface sectionBrushOptions {
    parentElement?: HTMLElement;
    data: DataCollection;
    menu: Menu;
    title?: string;
}

/**
 * 工具栏格式刷区域类
 * 提供格式刷、复制、剪切和粘贴等剪贴板操作功能按钮
 * @class
 */
export class SectionBrush {
    /** @type {HTMLDivElement} 格式刷区域容器元素 */
    private sectionBrush: HTMLDivElement;
    /** @type {DataCollection} 数据集合实例 */
    private data: DataCollection;
    /** @type {Menu} 菜单实例 */
    private menu: Menu;

    /**
     * 创建 SectionBrush 实例
     * @param {sectionBrushOptions} options - 构造选项
     */
    constructor(options: sectionBrushOptions) {
        this.sectionBrush = createDiv({
            className: TOOLBAR_CLASS_NAME.SECTION,
        }); 

        if (options.parentElement) {
            options.parentElement.append(this.sectionBrush);
        }
        this.data = options.data;
        this.menu = options.menu;
        this.init();
    }

    /**
     * 初始化工具栏按钮
     * 创建格式刷、复制、剪切和粘贴等剪贴板操作功能按钮
     * @public
     */
    public init() {        
        /*------------------ 格式刷分区 ------------------*/        
        const brushBtn = new BtnBase({  // 格式刷按钮
            parentElement: this.sectionBrush,
            data: this.data,
            menu: this.menu,
            toDo: 'brush',
            icon: 'icon-brush',
            text: '格式刷',
            ishow: true,
            tshow: true,
            layout: 'vertical',
            onClick: (todo: string) => {
                if (this.data.brushMode === 'continuous') {
                    this.data.brushMode = null;
                    brushBtn.setActived(false);
                } else {
                    this.data.brushMode = 'single';
                    this.data.copySelection();
                }
            },
            onDoubleClick: (todo: string) => {
                if (this.data.brushMode === 'continuous') {
                    this.data.brushMode = null;
                    brushBtn.setActived(false);
                } else {
                    this.data.brushMode = 'continuous';
                    // brushBtn.setActived(true);
                }
                this.data.copySelection();
            }
        });
        
        const copyBtn = new BtnBase({  // 复制按钮
            parentElement: this.sectionBrush,
            data: this.data,
            isDisabledOnEditting: false,
            menu: this.menu,
            toDo: 'copy',
            icon: 'icon-copy',
            text: '复制',
            ishow: true,
            tshow: true,
            layout: 'vertical',
            onClick: (todo: string) => {
                if (this.data.selectionOnEditting && this.data.isEditting) {
                    // 编辑状态下有字符选择，进行字符级复制
                    this.data.copyChars();
                } else {
                    // 单元格级复制
                    this.data.copySelection();
                }
            }
        });
        
        const cutBtn = new BtnBase({  // 剪切按钮
            parentElement: this.sectionBrush,
            data: this.data,
            isDisabledOnEditting: false,
            menu: this.menu,
            toDo: 'cut',
            icon: 'icon-scissors',
            text: '剪切',
            ishow: true,
            tshow: true,
            layout: 'vertical',
            onClick: (todo: string) => {
                if (this.data.selectionOnEditting && this.data.isEditting) {
                    // cutChars 返回目标光标绝对位置（被删除区间起点）。
                    // VALUES_CHANGED 会同步触发重绘并刷新 data.lines，因此此处写入
                    // cursorAbsolutePosition 时 lines 已是最新布局。
                    const caretIndex = this.data.cutChars();
                    this.data.cursorAbsolutePosition = caretIndex;
                } else {
                    // 单元格级剪切
                    this.data.cutSelection();
                }
            }
        });

        const pasteBtn = new BtnBase({  // 粘贴按钮
            parentElement: this.sectionBrush,
            data: this.data,
            isDisabledOnEditting: false,
            menu: this.menu,
            toDo: 'paste',
            icon: 'icon-paste',
            text: '粘贴',
            ishow: true,
            tshow: true,
            layout: 'vertical',
            separate: true,
            ashow: true,
            menuContent: {
                items: [
                    { todo: 'paste', icon: 'icon-paste', text: '粘贴'},
                    'separator',
                    { todo: 'paste-value', icon: 'icon-paste', text: '值'},
                    { todo: 'paste-style', icon: 'icon-paste', text: '格式'}
                ]
            },
            onClick: (todo: string) => {
                if (todo.includes('MenuOpen')) {
                    this.menu.openMenu(pasteBtn.getElements().container);
                    this.menu.addMenuElement(pasteBtn.getElements().menuContainer);
                } else {
                    this.handlePaste(todo);
                }
            }
        });

        this.data.on(DataEvents.BRUSH_MODE_CHANGED, (mode) => {
            brushBtn.setActived(mode === 'continuous');
        });        
    }

    /**
     * 处理粘贴操作
     * 根据操作类型执行不同的粘贴逻辑（全部粘贴、仅粘贴值、仅粘贴格式）
     * @param {string} todo - 操作类型标识（'paste' | 'paste-value' | 'paste-style'）
     * @private
     */
    private handlePaste(todo: string) {
        switch (todo) {
            case 'paste':
                if (this.data.isEditting) {
                    // 编辑状态下进行字符级粘贴
                    const clipboard = this.data.clipboardManager.getClipboard();
                    if (clipboard && clipboard.type === 'char' && clipboard.chars) {
                        // getSelectedCharsStartIndex(): 有选区时返回选区起点，无选区时返回光标处，
                        // 即真正的插入起点。pasteChars 返回插入完成后光标应处的绝对位置。
                        // VALUES_CHANGED 会同步触发重绘并刷新 data.lines，因此此处写入
                        // cursorAbsolutePosition 时 lines 已是最新布局。
                        const insertIndex = this.data.getSelectedCharsStartIndex();
                        const caretIndex = this.data.pasteChars(insertIndex, clipboard.chars);
                        this.data.cursorAbsolutePosition = caretIndex;
                    }
                } else {
                    // 非编辑状态下进行单元格级粘贴
                    this.data.pasteSelection();
                }
                break;
            case 'paste-value':
                this.data.pastePlainText();
                break;
            case 'paste-style':
                this.data.pasteFormatOnly();
                break;
        }
    }

    /**
     * 获取工具栏元素
     * @returns {{sectionBrush: HTMLDivElement}} 工具栏元素对象
     * @public
     */
    public getElements() {
        return {
            sectionBrush: this.sectionBrush,
        };
    }
}