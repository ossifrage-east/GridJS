/*  # 工具栏·段落分区 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目目标: 把「段落样式 / 项目符号 / 编号 / 缩进」做成**与载体无关**的分区 ——
    #           同一份按钮既能挂在文档上，将来也能挂到任何「有段落概念」的载体上。
    # 本文件职责: 只负责外观与交互：建一个段落样式下拉、两个列表按钮；
    #             点下去一律 {@link CommandBus.dispatch}，高亮一律按 {@link CommandBus.snapshot} 回读。
    #             缩进的两个按钮**不在这里**：用户要它们与对齐按钮成组、并独占一行，
    #             已搬进「对齐方式」分区（见 sectionAlign 的 indentRow）。
    # 依赖方向: contract/commandBus + common/*。**不 import document/**、**不 import dataArchitecture**：
    #           分区不认识「块树」也不认识「单元格」，只认识 4 个命令名。
    # 为什么在工具栏里而不是文档工作区里: 需求要求「正文、项目符号等全部纳入原有 toolbar」，
    #           文档工作区因此不再自带功能区（原 `.doc-ribbon` 已删除）。
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
import { Select } from "../../common/select";
import { CommandBus } from "../../contract/commandBus";
import { DocListKind, EDIT_STATE_CHANGED, ParaStyleId } from "../../contract/editHost";

/**
 * 段落样式菜单项
 *
 * `todo` 直接就是 {@link ParaStyleId}，因此「菜单点选值 → 命令载荷」不需要任何翻译表，
 * 少一层映射就少一处会写错的地方。
 */
const PARA_STYLE_ITEMS: Array<{ todo: ParaStyleId; text: string }> = [
    { todo: 'paragraph', text: '正文' },
    { todo: 'h1', text: '标题 1' },
    { todo: 'h2', text: '标题 2' },
    { todo: 'h3', text: '标题 3' },
    { todo: 'h4', text: '标题 4' },
    { todo: 'h5', text: '标题 5' },
    { todo: 'h6', text: '标题 6' },
    { todo: 'quote', text: '引用' },
    { todo: 'code', text: '代码块' },
];

/**
 * SectionParagraph 构造选项接口
 * @interface
 * @property {HTMLElement} [parentElement] - 父元素容器
 * @property {CommandBus} bus - 命令总线
 * @property {Menu} menu - 菜单实例
 */
export interface SectionParagraphOptions {
    parentElement?: HTMLElement;
    bus: CommandBus;
    menu: Menu;
    title?: string;
}

/**
 * 工具栏段落分区
 *
 * 四个命令：`setParaStyle`（下拉）、`toggleList`（项目符号 / 编号）、
 * `indentStep`（±1 级，与「对齐方式」分区共用同一个命令）。
 * @class
 */
export class SectionParagraph {
    /** 分区容器 */
    private section: HTMLDivElement;
    /** 命令总线 */
    private readonly bus: CommandBus;
    /** 菜单实例 */
    private readonly menu: Menu;
    /** 段落样式下拉 */
    private styleSelect!: Select;
    /** 列表按钮 */
    private bulletBtn!: BtnBase;
    private orderedBtn!: BtnBase;

    /**
     * 创建实例
     * @param options - 构造选项
     */
    constructor(options: SectionParagraphOptions) {
        this.section = createDiv({ className: TOOLBAR_CLASS_NAME.SECTION });
        if (options.parentElement) options.parentElement.append(this.section);
        this.bus = options.bus;
        this.menu = options.menu;
        this.init();
    }

    /**
     * 建按钮
     * @private
     */
    private init(): void {
        const line = createDiv({ className: TOOLBAR_CLASS_NAME.SECTION_LINE });
        const column = createDiv({
            style: { display: 'flex', flexDirection: 'column', gap: '4px' },
        });
        this.section.append(line, column);

        const top = createDiv({ style: { display: 'flex', alignItems: 'center', justifyContent: 'center' } });
        const bottom = createDiv({
            style: { display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '4px' },
        });
        column.append(top, bottom);

        this.styleSelect = new Select({
            parentElement: top,
            toDo: 'paraStyle',
            title: '段落样式',
            menuContent: { items: PARA_STYLE_ITEMS },
            onClick: (todo) => {
                this.menu.closeMenu();
                if (todo.includes('MenuOpen')) {
                    this.menu.openMenu(this.styleSelect.getElements().container);
                    this.menu.addMenuElement(this.styleSelect.getElements().menuContainer);
                    return;
                }
                if (todo.includes('Menu')) return;
                this.bus.dispatch('setParaStyle', { value: todo as ParaStyleId });
            },
        });
        // 首帧先显示「正文」：文档还没打开时下拉不该是空白
        this.styleSelect.setText('paragraph');

        this.bulletBtn = this.createListButton(bottom, 'bullet', '项目符号', '无序列表（项目符号）');
        this.orderedBtn = this.createListButton(bottom, 'ordered', '编号', '有序列表（编号）');

        this.setupStateListeners();
    }

    /**
     * 建一个列表开关按钮
     *
     * 用文字标签而不是图标：图标字体里没有「项目符号 / 编号」两个字型，
     * 就近借用别的图标（如 `icon-text-height`）比没有图标更让人困惑。
     * @param parent - 父容器
     * @param kind - 列表风格
     * @param text - 按钮文字
     * @param title - 悬浮提示
     * @returns 按钮实例
     * @private
     */
    private createListButton(
        parent: HTMLElement,
        kind: DocListKind,
        text: string,
        title: string,
    ): BtnBase {
        return new BtnBase({
            parentElement: parent,
            stateSource: this.bus,
            isDisabledOnEditting: false,
            menu: this.menu,
            toDo: kind,
            text,
            title,
            ishow: false,
            tshow: true,
            layout: 'horizontal',
            onClick: () => this.bus.dispatch('toggleList', { value: kind }),
        });
    }

    /**
     * 订阅状态变化，按快照刷新段落样式下拉与列表高亮
     * @private
     */
    private setupStateListeners(): void {
        this.bus.on(EDIT_STATE_CHANGED, () => this.setToolbarsState());
        this.setToolbarsState();
    }

    /**
     * 按快照刷新外观
     *
     * `para.paraStyle` / `para.list` 由宿主给出**归一化后的结论**，
     * 分区不需要知道「块类型是 heading 还是 listItem」。
     * @private
     */
    private setToolbarsState(): void {
        const { para } = this.bus.snapshot();
        if (para.paraStyle) this.styleSelect.setText(para.paraStyle);
        this.bulletBtn.setActived(para.list === 'bullet');
        this.orderedBtn.setActived(para.list === 'ordered');
    }

    /**
     * 取分区元素
     * @returns 分区容器
     */
    public getElements(): { sectionParagraph: HTMLDivElement } {
        return { sectionParagraph: this.section };
    }
}
