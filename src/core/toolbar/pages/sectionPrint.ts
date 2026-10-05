/*  # web 表格顶部工具栏打印设置集合 TS 文件（「页面」标签页）
    # 提供：打印预览、页边距、纸张方向、纸张大小、打印区域、打印缩放按钮。
    # 除打印预览外均为「上图标下文字 + 下拉菜单」按钮，
    # 设置通过共享 Printer 实例（getSharedPrinter）写入，与快速访问栏互通。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { createDiv } from "../../../utils/dom";
import { TOOLBAR_CLASS_NAME, MENU_CLASS_NAMES } from "../../constant";
import { DataCollection } from "../../dataArchitecture/dataCollection";
import { Menu } from "../../common/menu";
import { BtnBase } from "../../common/button";
import { Printer, getSharedPrinter, MarginPreset, PaperOrientation, PaperSize, ScaleMode } from "../../../utils/printer";

/**
 * sectionPrintOptions 接口
 */
export interface sectionPrintOptions {
    parentElement?: HTMLElement;  // 父元素，工具栏容器
    data: DataCollection;  // 数据集合
    menu: Menu;  // 菜单实例
    title?: string;
}

/**
 * 工具栏打印设置区域类（「页面」标签页）
 *
 * 按钮清单：
 * 1. 打印预览：直接打开打印预览；
 * 2. 页边距：常规 / 宽（默认）/ 窄 三种预设；
 * 3. 纸张方向：纵向 / 横向；
 * 4. 纸张大小：A4 / A3 / B5 / Letter；
 * 5. 打印区域：设置打印区域（以当前选区为准）/ 取消打印区域；
 * 6. 打印缩放：无缩放 / 将整个工作表打印在一页 / 将所有列打印在一页 /
 *    将所有行打印在一页 / 自定义（输入框输入 10-400 缩放比例）。
 *
 * @example
 * ```ts
 * const sectionPrint = new SectionPrint({ parentElement, data, menu });
 * const { sectionPrint: el } = sectionPrint.getElements();
 * toolbar.append(el);
 * ```
 */
export class SectionPrint {
    private sectionPrint: HTMLDivElement;  // 打印设置区域容器
    private data: DataCollection;  // 数据集合
    private menu: Menu;  // 菜单实例
    /** 共享打印实例：与快速访问栏的打印/打印预览按钮共用，打印设置互通 */
    private printer: Printer;

    /**
     * 构造函数：创建打印设置区域容器并初始化按钮
     * @param {sectionPrintOptions} options - 可配置项
     */
    constructor(options: sectionPrintOptions) {
        this.sectionPrint = createDiv({
            className: TOOLBAR_CLASS_NAME.SECTION,
        });
        if (options.parentElement) {
            options.parentElement.append(this.sectionPrint);
        }
        this.data = options.data;
        this.menu = options.menu;
        this.printer = getSharedPrinter(options.data);
        this.init();
    }

    /**
     * 初始化打印设置区域：依次创建打印预览、页边距、纸张方向、纸张大小、
     * 打印区域、打印缩放按钮（上图标下文字，除打印预览外均带下拉菜单）
     */
    public init() {
        /*------------------ 打印预览 ------------------*/
        new BtnBase({
            parentElement: this.sectionPrint,
            data: this.data,
            menu: this.menu,
            toDo: 'printPreview',
            icon: 'icon-printer-view',
            text: '打印预览',
            ishow: true,
            tshow: true,
            layout: 'vertical',
            onClick: () => this.printer.printPreview(),
        });

        /*------------------ 页边距：常规 / 宽（默认）/ 窄 ------------------*/
        const marginTodos = ['normal', 'wide', 'narrow'];
        const marginBtn = new BtnBase({
            parentElement: this.sectionPrint,
            data: this.data,
            menu: this.menu,
            toDo: 'margin',
            icon: 'icon-page-margins',
            text: '页边距',
            ishow: true,
            tshow: true,
            ashow: true,
            layout: 'vertical',
            menuContent: {
                items: [
                    { todo: 'normal', text: '常规' },
                    { todo: 'wide', text: '宽' },
                    { todo: 'narrow', text: '窄' },
                ],
            },
            onClick: (todo: string) => {
                if (todo === 'marginMenuOpen') {
                    this.markActiveItem(marginBtn, marginTodos, this.printer.getMarginPreset());
                    this.openDropdown(marginBtn);
                    return;
                }
                if (todo === 'marginMenuClose') return;
                this.printer.setMarginPreset(todo as MarginPreset);
            },
        });

        /*------------------ 纸张方向：纵向 / 横向 ------------------*/
        const orientationTodos = ['portrait', 'landscape'];
        const orientationBtn = new BtnBase({
            parentElement: this.sectionPrint,
            data: this.data,
            menu: this.menu,
            toDo: 'orientation',
            icon: 'icon-orientation',
            text: '纸张方向',
            ishow: true,
            tshow: true,
            ashow: true,
            layout: 'vertical',
            menuContent: {
                items: [
                    { todo: 'portrait', text: '纵向' },
                    { todo: 'landscape', text: '横向' },
                ],
            },
            onClick: (todo: string) => {
                if (todo === 'orientationMenuOpen') {
                    this.markActiveItem(orientationBtn, orientationTodos, this.printer.getOrientation());
                    this.openDropdown(orientationBtn);
                    return;
                }
                if (todo === 'orientationMenuClose') return;
                this.printer.setOrientation(todo as PaperOrientation);
            },
        });

        /*------------------ 纸张大小：A4 / A3 / B5 / Letter ------------------*/
        const paperTodos = ['A4', 'A3', 'B5', 'Letter'];
        const paperBtn = new BtnBase({
            parentElement: this.sectionPrint,
            data: this.data,
            menu: this.menu,
            toDo: 'paper',
            icon: 'icon-paper-size',
            text: '纸张大小',
            ishow: true,
            tshow: true,
            ashow: true,
            layout: 'vertical',
            menuContent: {
                items: paperTodos.map(todo => ({ todo, text: todo })),
            },
            onClick: (todo: string) => {
                if (todo === 'paperMenuOpen') {
                    this.markActiveItem(paperBtn, paperTodos, this.printer.getPaperSize());
                    this.openDropdown(paperBtn);
                    return;
                }
                if (todo === 'paperMenuClose') return;
                this.printer.setPaperSize(todo as PaperSize);
            },
        });

        /*------------------ 打印区域：状态回显 + 设置 / 取消 ------------------*/
        // 自定义菜单内容：状态行（回显当前打印区域，统一读自 data.printSetting）+ 分隔线 + 设置/取消项
        const areaMenu = document.createElement('div');
        const areaStatus = document.createElement('div');
        areaStatus.className = MENU_CLASS_NAMES.ITEM;
        const areaStatusText = document.createElement('span');
        areaStatusText.className = 'text';
        areaStatusText.style.opacity = '0.65';
        areaStatus.appendChild(areaStatusText);
        areaMenu.appendChild(areaStatus);
        const areaLine = document.createElement('div');
        areaLine.className = MENU_CLASS_NAMES.LINE;
        areaMenu.appendChild(areaLine);
        const areaActions = [
            { todo: 'set', text: '设置打印区域' },
            { todo: 'clear', text: '取消打印区域' },
        ];
        for (const item of areaActions) {
            const menuItem = document.createElement('div');
            menuItem.className = MENU_CLASS_NAMES.ITEM;
            const span = document.createElement('span');
            span.className = 'text';
            span.textContent = item.text;
            menuItem.appendChild(span);
            menuItem.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                if (item.todo === 'set') this.printer.setPrintArea();
                if (item.todo === 'clear') this.printer.clearPrintArea();
                areaStatusText.textContent = this.printer.getPrintAreaLabel();
                this.menu.closeMenu();
            });
            areaMenu.appendChild(menuItem);
        }
        const areaBtn = new BtnBase({
            parentElement: this.sectionPrint,
            data: this.data,
            menu: this.menu,
            toDo: 'area',
            icon: 'icon-print-area',
            text: '打印区域',
            ishow: true,
            tshow: true,
            ashow: true,
            layout: 'vertical',
            menuContent: {
                items: () => areaMenu,
            },
            onClick: (todo: string) => {
                if (todo === 'areaMenuOpen') {
                    areaStatusText.textContent = this.printer.getPrintAreaLabel();
                    this.openDropdown(areaBtn);
                }
            },
        });

        /*------------------ 打印缩放：预设 + 自定义（输入框） ------------------*/
        const scaleTodos = ['none', 'fit-sheet', 'fit-cols', 'fit-rows', 'custom'];
        const scaleItems = [
            { todo: 'none', text: '无缩放' },
            { todo: 'fit-sheet', text: '将整个工作表打印在一页' },
            { todo: 'fit-cols', text: '将所有列打印在一页' },
            { todo: 'fit-rows', text: '将所有行打印在一页' },
        ];
        // 自定义菜单内容（MenuContent 支持 () => HTMLElement 项）：预设项 + 分隔线 + 自定义行
        const scaleMenu = document.createElement('div');
        for (const item of scaleItems) {
            const menuItem = document.createElement('div');
            menuItem.className = MENU_CLASS_NAMES.ITEM;
            const span = document.createElement('span');
            span.className = 'text';
            span.textContent = item.text;
            menuItem.appendChild(span);
            menuItem.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                this.printer.setScaleMode(item.todo as ScaleMode);
                this.menu.closeMenu();
            });
            scaleMenu.appendChild(menuItem);
        }
        const line = document.createElement('div');
        line.className = MENU_CLASS_NAMES.LINE;
        scaleMenu.appendChild(line);
        // 自定义缩放行：比例输入框（10-400）+ 确定
        const customRow = document.createElement('div');
        customRow.className = MENU_CLASS_NAMES.ITEM;
        customRow.style.display = 'flex';
        customRow.style.alignItems = 'center';
        customRow.style.gap = '6px';
        const customLabel = document.createElement('span');
        customLabel.className = 'text';
        customLabel.textContent = '自定义缩放';
        const customInput = document.createElement('input');
        customInput.type = 'number';
        customInput.min = '10';
        customInput.max = '400';
        customInput.step = '5';
        customInput.value = String(this.printer.getCustomScale());
        customInput.style.width = '56px';
        customInput.style.padding = '2px 4px';
        customInput.style.border = '1px solid #ced4da';
        customInput.style.borderRadius = '4px';
        customInput.style.outline = 'none';
        const percentSpan = document.createElement('span');
        percentSpan.textContent = '%';
        const confirmBtn = document.createElement('button');
        confirmBtn.type = 'button';
        confirmBtn.textContent = '确定';
        confirmBtn.style.padding = '2px 10px';
        confirmBtn.style.border = '1px solid #dee2e6';
        confirmBtn.style.borderRadius = '4px';
        confirmBtn.style.background = '#fff';
        confirmBtn.style.cursor = 'pointer';
        const applyCustomScale = () => {
            const v = Number(customInput.value);
            if (!Number.isFinite(v)) return;
            this.printer.setScaleMode('custom', v);
            this.menu.closeMenu();
        };
        confirmBtn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            applyCustomScale();
        });
        customInput.addEventListener('keydown', (e) => {
            e.stopPropagation();  // 避免按键冒泡触发表格快捷键
            if (e.key === 'Enter') applyCustomScale();
        });
        customRow.append(customLabel, customInput, percentSpan, confirmBtn);
        scaleMenu.appendChild(customRow);

        const scaleBtn = new BtnBase({
            parentElement: this.sectionPrint,
            data: this.data,
            menu: this.menu,
            toDo: 'scale',
            icon: 'icon-print-zoom',
            text: '打印缩放',
            ishow: true,
            tshow: true,
            ashow: true,
            layout: 'vertical',
            menuContent: {
                items: () => scaleMenu,
            },
            onClick: (todo: string) => {
                if (todo === 'scaleMenuOpen') {
                    this.markActiveItem(scaleBtn, scaleTodos, this.printer.getScaleMode());
                    customInput.value = String(this.printer.getCustomScale());
                    this.openDropdown(scaleBtn);
                }
            },
        });

        /*------------------ 菜单内容统一标记 ------------------*/
        // 为 5 个下拉按钮的菜单内容打 .print-menu 标记：
        // 每个菜单项文字前预留固定「对号位」留白（样式见 menu.scss），选中时 ✓ 显示在留白处
        for (const btn of [marginBtn, orientationBtn, paperBtn, areaBtn, scaleBtn]) {
            btn.getElements().menuContainer.classList.add('print-menu');
        }
    }

    /**
     * 打开按钮的下拉菜单（菜单内容已在 menuContent 中构建完成）
     * @param {BtnBase} btn - 触发菜单的按钮
     * @private
     */
    private openDropdown(btn: BtnBase): void {
        this.menu.openMenu(btn.getElements().container);
        this.menu.addMenuElement(btn.getElements().menuContainer);
    }

    /**
     * 回显下拉菜单当前选中项：为选中项加 .checked 类显示 ✓（✓ 渲染在文字前
     * 预留的对号位中，见 menu.scss 的 .print-menu 规则），文字本身不改动
     * @param {BtnBase} btn - 菜单所属按钮
     * @param {string[]} todos - 与菜单项顺序一致的 todo 列表
     * @param {string} active - 当前选中的 todo
     * @private
     */
    private markActiveItem(btn: BtnBase, todos: string[], active: string): void {
        const items = btn.getElements().menuContainer?.querySelectorAll<HTMLElement>('.menu-item');
        if (!items) return;
        items.forEach((item, i) => {
            const todo = todos[i];
            if (todo === undefined) return;
            item.classList.toggle('checked', todo === active);
        });
    }

    /**
     * 获取打印设置区域元素
     * @returns 打印设置区域容器
     */
    public getElements() {
        return {
            sectionPrint: this.sectionPrint,
        };
    }
}
