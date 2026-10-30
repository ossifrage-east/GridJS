/*  # 工具栏·文档页面设置分区 TS 文件（「页面」标签页）
    # 项目名称: silkspaces（藕丝空间）
    # 项目目标: 给文档工作区补上「页面」根分类 —— 与表格的「页面」同样是
    #           纸张方向 / 纸张大小 / 页边距 / 缩放 + 打印，但落到文档自己的页面设置上。
    # 本文件职责: 只负责**外观与交互**：点选后把「哪个键、改成什么」交给装配层注入的回调。
    #             **不 import document/**：页面设置是文档模型的一部分，而工具栏必须
    #             对「文档」一无所知（与 Markdown 视图开关同一取舍，见 sectionDocView）。
    # 依赖方向: contract/commandBus（只借它订阅编辑态）+ common/* + constant。
    # author: 东方鹗
    # B站: https://space.bilibili.com/video/...
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { createDiv } from "../../../utils/dom";
import { TOOLBAR_CLASS_NAME } from "../../constant";
import { BtnBase } from "../../common/button";
import { Menu } from "../../common/menu";
import { CommandBus } from "../../contract/commandBus";

/** 页面设置的四个键（与 `DocPageSetup` 的字段名一一对应，但这里是**纯字符串**，工具栏不认识模型） */
export type DocPageKey = 'orientation' | 'paperSize' | 'marginPreset' | 'scaleMode';

/** 分区回读用的页面设置快照（全部是字符串，便于本文件不依赖文档模型） */
export interface DocPageSetupView {
    /** 纸张方向 */
    orientation: string;
    /** 纸张大小 */
    paperSize: string;
    /** 页边距预设 */
    marginPreset: string;
    /** 缩放模式 */
    scaleMode: string;
}

/**
 * SectionDocPage 构造选项
 * @interface
 */
export interface SectionDocPageOptions {
    parentElement?: HTMLElement;
    bus: CommandBus;
    menu: Menu;
    /** 回读当前页面设置（打开下拉时标记选中项） */
    getSetup: () => DocPageSetupView;
    /** 改某一项（装配层负责写回文档元数据并重排纸张） */
    onChange: (key: DocPageKey, value: string) => void;
    /** 打印当前文档 */
    onPrint: () => void;
}

/**
 * 工具栏「页面设置」分区（文档工作区的「页面」根分类）
 *
 * 与表格「打印设置」分区的关系：两者外观一致（同一套下拉菜单写法），
 * 差异只在**写入目标** —— 表格写 `Printer`，文档写 `DocMeta.pageSetup`。
 * 因此这里不复用 `SectionPrint`：那个类把 `Printer` 写进了构造签名，
 * 而文档没有（也不该有）表格打印器。
 * @class
 */
export class SectionDocPage {
    /** 分区容器 */
    private section: HTMLDivElement;
    /** 命令总线（只用来订阅状态刷新） */
    private readonly bus: CommandBus;
    /** 菜单实例 */
    private readonly menu: Menu;
    /** 回读回调 */
    private readonly getSetup: () => DocPageSetupView;
    /** 变更回调 */
    private readonly onChange: (key: DocPageKey, value: string) => void;
    /** 打印回调 */
    private readonly onPrint: () => void;

    /**
     * 创建实例
     * @param options - 构造选项
     */
    constructor(options: SectionDocPageOptions) {
        this.section = createDiv({ className: TOOLBAR_CLASS_NAME.SECTION });
        if (options.parentElement) options.parentElement.append(this.section);
        this.bus = options.bus;
        this.menu = options.menu;
        this.getSetup = options.getSetup;
        this.onChange = options.onChange;
        this.onPrint = options.onPrint;
        this.init();
    }

    /**
     * 建按钮：打印 / 页边距 / 纸张方向 / 纸张大小 / 缩放
     * @private
     */
    private init(): void {
        new BtnBase({
            parentElement: this.section,
            stateSource: this.bus,
            isDisabledOnEditting: false,
            menu: this.menu,
            toDo: 'docPrint',
            icon: 'icon-printer',
            text: '打印',
            title: '打印当前文档',
            ishow: true,
            tshow: true,
            layout: 'vertical',
            onClick: () => this.onPrint(),
        });

        this.createDropdown({
            toDo: 'docMargin',
            icon: 'icon-page-margins',
            text: '页边距',
            key: 'marginPreset',
            items: [
                { todo: 'normal', text: '常规' },
                { todo: 'wide', text: '宽' },
                { todo: 'narrow', text: '窄' },
            ],
        });

        this.createDropdown({
            toDo: 'docOrientation',
            icon: 'icon-orientation',
            text: '纸张方向',
            key: 'orientation',
            items: [
                { todo: 'portrait', text: '纵向' },
                { todo: 'landscape', text: '横向' },
            ],
        });

        this.createDropdown({
            toDo: 'docPaperSize',
            icon: 'icon-paper-size',
            text: '纸张大小',
            key: 'paperSize',
            items: [
                { todo: 'A4', text: 'A4' },
                { todo: 'A5', text: 'A5' },
                { todo: 'A3', text: 'A3' },
                { todo: 'Letter', text: 'Letter' },
                { todo: 'Legal', text: 'Legal' },
            ],
        });

        this.createDropdown({
            toDo: 'docScaleMode',
            icon: 'icon-print-zoom',
            text: '缩放',
            key: 'scaleMode',
            items: [
                { todo: 'none', text: '无缩放' },
                { todo: 'fit-page', text: '适应页宽' },
            ],
        });
    }

    /**
     * 建一个「图标 + 文字 + 下拉菜单」的设置按钮
     *
     * 打开菜单前先回读一次当前值并标记选中项：下拉不标选中，
     * 用户就得靠记忆判断「现在到底是 A4 还是 A3」。
     * @param options - 按钮与菜单的配置
     * @private
     */
    private createDropdown(options: {
        toDo: string;
        icon: string;
        text: string;
        key: DocPageKey;
        items: { todo: string; text: string }[];
    }): void {
        const { toDo, icon, text, key, items } = options;
        const current = (): string => {
            const setup = this.getSetup();
            const value = setup[key];
            return typeof value === 'string' ? value : '';
        };
        const todos: string[] = [];
        for (let i = 0; i < items.length; i++) todos.push(items[i].todo);
        const btn = new BtnBase({
            parentElement: this.section,
            stateSource: this.bus,
            isDisabledOnEditting: false,
            menu: this.menu,
            toDo,
            icon,
            text,
            title: text,
            ishow: true,
            tshow: true,
            ashow: true,
            layout: 'vertical',
            menuContent: { items },
            onClick: (todo: string) => {
                if (todo === `${toDo}MenuOpen`) {
                    // 打开前回读一次并标记选中项：下拉不标选中，
                    // 用户就只能凭记忆猜「现在到底是 A4 还是 A3」
                    this.markActiveItem(btn, todos, current());
                    this.menu.openMenu(btn.getElements().container);
                    this.menu.addMenuElement(btn.getElements().menuContainer);
                    return;
                }
                if (todo.includes('Menu')) return;
                this.onChange(key, todo);
            },
        });
    }

    /**
     * 在下拉菜单里标记当前生效的那一项
     * @param btn - 按钮实例
     * @param todos - 菜单项标识（顺序与菜单一致）
     * @param active - 当前生效值
     * @private
     */
    private markActiveItem(btn: BtnBase, todos: string[], active: string): void {
        const nodes = btn.getElements().menuContainer?.querySelectorAll('.menu-item');
        if (!nodes) return;
        for (let i = 0; i < nodes.length && i < todos.length; i++) {
            const node = nodes[i] as HTMLElement;
            if (node.classList) node.classList.toggle('checked', todos[i] === active);
        }
    }

    /**
     * 取分区元素
     * @returns 分区容器
     */
    public getElements(): { sectionDocPage: HTMLDivElement } {
        return { sectionDocPage: this.section };
    }
}
