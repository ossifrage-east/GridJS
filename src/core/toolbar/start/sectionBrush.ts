/*  # web 表格顶部工具栏格式刷集合 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目描述: 基于 HTML5 Canvas 的 Web 电子表格组件，无需服务端支持即可在浏览器中创建、读取、写入 Excel 文件。
    # 项目目标: 工具栏的「格式刷」分区 —— 格式刷、复制、剪切、粘贴（含只粘值 / 只粘格式）。
    # 本文件职责: 只负责**外观与交互**；剪贴板与撤销记账全部经 {@link CommandBus} 派发。
    #             改造前这里直接调了 runWithFullStateUndo / emit(VALUES_CHANGED) ——
    #             「撤销怎么记账」是宿主的事，外壳不该知道，现已搬进 SheetEditHost。
    # 依赖方向: contract/commandBus + common/*，**不再 import dataArchitecture**。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { createDiv } from "../../../utils/dom";
import { TOOLBAR_CLASS_NAME } from "../../constant";
import { Menu, MenuItemSpec } from "../../common/menu";
import { BtnBase } from "../../common/button";
import { CommandBus } from "../../contract/commandBus";
import { EDIT_STATE_CHANGED, WorkspaceKind } from "../../contract/editHost";

/**
 * 表格工作区的粘贴菜单
 *
 * 表格的「只粘值 / 只粘格式」是单元格语义，用角标图标（T / 格式刷）区分，
 * 与文字标签相比更省横向空间。
 */
const SHEET_PASTE_ITEMS: MenuItemSpec[] = [
    { todo: 'paste', icon: 'icon-paste', text: '粘贴' },
    'separator',
    { todo: 'paste-value', icon: 'icon-paste', badge: 'T' as const, text: '值' },
    { todo: 'paste-style', icon: 'icon-paste', badge: 'brush' as const, text: '格式' },
];

/**
 * 文档工作区的粘贴菜单
 *
 * 文档里没有「值」这个概念，照 Word 的措辞写「粘贴文字 / 粘贴格式」。
 * 两条命令与表格复用同一组 `CommandId`（`pasteValue` / `pasteStyle`）：
 * 「只保留文本」和「只保留格式」在两种载体里是同一件事，没必要再造两个动词。
 */
const DOC_PASTE_ITEMS: MenuItemSpec[] = [
    { todo: 'paste', icon: 'icon-paste', text: '粘贴' },
    'separator',
    { todo: 'paste-value', icon: 'icon-paste', text: '粘贴文字' },
    { todo: 'paste-style', icon: 'icon-paste', text: '粘贴格式' },
];

/**
 * SectionBrush 构造选项接口
 * @interface
 * @property {HTMLElement} [parentElement] - 父元素容器
 * @property {CommandBus} bus - 命令总线（编辑宿主的状态与命令入口）
 * @property {Menu} menu - 菜单实例
 * @property {string} [title] - 工具栏标题
 */
export interface sectionBrushOptions {
    parentElement?: HTMLElement;
    bus: CommandBus;
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
    /** @type {CommandBus} 命令总线（外壳与宿主之间唯一的接触面） */
    private bus: CommandBus;
    /** @type {Menu} 菜单实例 */
    private menu: Menu;
    /** 格式刷按钮（状态刷新时需要单独回写高亮） */
    private brushBtn!: BtnBase;
    /** 粘贴按钮（文档工作区要换一套菜单项，见 setKind） */
    private pasteBtn!: BtnBase;

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
        this.bus = options.bus;
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
            stateSource: this.bus,
            menu: this.menu,
            toDo: 'brush',
            icon: 'icon-brush',
            text: '格式刷',
            ishow: true,
            tshow: true,
            layout: 'vertical',
            // 单击 / 双击是两个不同命令：单击 = 单次刷，双击 = 连续刷。
            // 高亮不再由这里手写 setActived，而是统一按快照的 brushMode 推导
            onClick: (todo: string) => {
                this.bus.dispatch('brushClick');
            },
            onDoubleClick: (todo: string) => {
                this.bus.dispatch('brushDoubleClick');
            }
        });
        
        const copyBtn = new BtnBase({  // 复制按钮
            parentElement: this.sectionBrush,
            stateSource: this.bus,
            isDisabledOnEditting: false,
            menu: this.menu,
            toDo: 'copy',
            icon: 'icon-copy',
            text: '复制',
            ishow: true,
            tshow: true,
            layout: 'vertical',
            // 「编辑态有字符选区 → 字符级复制，否则单元格级」这个判断在宿主内
            onClick: (todo: string) => {
                this.bus.dispatch('copy');
            }
        });
        
        const cutBtn = new BtnBase({  // 剪切按钮
            parentElement: this.sectionBrush,
            stateSource: this.bus,
            isDisabledOnEditting: false,
            menu: this.menu,
            toDo: 'cut',
            icon: 'icon-scissors',
            text: '剪切',
            ishow: true,
            tshow: true,
            layout: 'vertical',
            onClick: (todo: string) => {
                this.bus.dispatch('cut');
            }
        });

        const pasteBtn = new BtnBase({  // 粘贴按钮
            parentElement: this.sectionBrush,
            stateSource: this.bus,
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
            menuContent: { items: SHEET_PASTE_ITEMS },
            onClick: (todo: string) => {
                if (todo.includes('MenuOpen')) {
                    this.menu.openMenu(pasteBtn.getElements().container);
                    this.menu.addMenuElement(pasteBtn.getElements().menuContainer);
                } else {
                    this.handlePaste(todo);
                }
            }
        });

        this.brushBtn = brushBtn;
        this.pasteBtn = pasteBtn;
        this.setupStateListeners();
    }

    /**
     * 按工作区类型收敛分区内容
     *
     * 文档里这个分区是「剪贴板 + 格式刷」：
     * - 格式刷**保留**：文档宿主实现了 `brushClick` / `brushDoubleClick`
     *   （抄字符样式 → 刷到下一次选区），与表格同一套手势（单击单次、双击连续）；
     * - 粘贴**保留**下拉菜单：Word 的粘贴同样有「只粘文字 / 只粘格式」两项，
     *   收成普通按钮反而让人以为功能缺失（早期版本就是这么做的，被当成 bug 报回来）。
     * @param kind - 工作区类型
     */
    public setKind(kind: WorkspaceKind): void {
        const doc = kind === 'doc';
        // 粘贴按钮在两类工作区里都保留下拉菜单（形状一致才不会让人以为功能缺失），
        // 只是菜单项按载体换一套措辞：表格是「值 / 格式」，文档是「文字 / 格式」。
        this.pasteBtn.setMenuEnabled(true);
        this.pasteBtn.setMenuItems(doc ? DOC_PASTE_ITEMS : SHEET_PASTE_ITEMS);
    }

    /**
     * 处理粘贴操作
     * 根据操作类型派发不同命令（全部粘贴、仅粘贴值、仅粘贴格式）
     * @param {string} todo - 操作类型标识（'paste' | 'paste-value' | 'paste-style'）
     * @private
     */
    private handlePaste(todo: string) {
        switch (todo) {
            case 'paste':
                this.bus.dispatch('paste');
                break;
            case 'paste-value':
                this.bus.dispatch('pasteValue');
                break;
            case 'paste-style':
                this.bus.dispatch('pasteStyle');
                break;
        }
    }

    /**
     * 按当前宿主状态刷新按钮高亮
     *
     * 只有「连续模式」才让格式刷按钮保持高亮：单次模式刷一次就自动退出，
     * 亮着反而误导（改造前 BRUSH_MODE_CHANGED 的订阅回调就是这个口径）。
     * @private
     */
    private setToolbarsState() {
        this.brushBtn.setActived(this.bus.snapshot().brushMode === 'continuous');
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
