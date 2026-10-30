/*  # 工具栏·文档视图分区 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目目标: 给「富文本 / Markdown 源码」两种视图之间提供切换入口，
    #           放在「视图」根分类下 —— 与 Word 的「视图」页签同一位置语义。
    # 本文件职责: 一个开关按钮 + 高亮回读。**不切换任何东西**，
    #             真正的切换动作由装配层注入（`onToggle`），因为它属于「工作区怎么显示」
    #             而不是「一次编辑动作」，塞进命令集会让宿主凭空多出一个视图职责。
    # 依赖方向: common/* + contract（仅取事件名）。不 import document/**。
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
 * SectionDocView 构造选项接口
 * @interface
 */
export interface SectionDocViewOptions {
    parentElement?: HTMLElement;
    bus: CommandBus;
    menu: Menu;
    /** 切换视图（由装配层注入，实际动作在文档工作区） */
    onToggle: () => void;
    /** 当前是否处于源码视图（用于回读高亮） */
    isActive: () => boolean;
}

/**
 * 工具栏「文档视图」分区
 * @class
 */
export class SectionDocView {
    /** 分区容器 */
    private section: HTMLDivElement;
    /** 命令总线（只用来订阅状态刷新） */
    private readonly bus: CommandBus;
    /** 菜单实例 */
    private readonly menu: Menu;
    /** 切换回调 */
    private readonly onToggle: () => void;
    /** 当前是否源码视图 */
    private readonly isActive: () => boolean;
    /** 按钮 */
    private toggleBtn!: BtnBase;

    /**
     * 创建实例
     * @param options - 构造选项
     */
    constructor(options: SectionDocViewOptions) {
        this.section = createDiv({ className: TOOLBAR_CLASS_NAME.SECTION });
        if (options.parentElement) options.parentElement.append(this.section);
        this.bus = options.bus;
        this.menu = options.menu;
        this.onToggle = options.onToggle;
        this.isActive = options.isActive;
        this.init();
    }

    /**
     * 建按钮
     * @private
     */
    private init(): void {
        const line = createDiv({ className: TOOLBAR_CLASS_NAME.SECTION_LINE });
        const row = createDiv({
            style: { display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '4px', padding: '4px' },
        });
        this.section.append(line, row);

        this.toggleBtn = new BtnBase({
            parentElement: row,
            stateSource: this.bus,
            isDisabledOnEditting: false,
            menu: this.menu,
            toDo: 'markdownView',
            icon: 'icon-file font15',
            text: 'Markdown 源码',
            title: '在富文本与 Markdown 源码之间切换',
            ishow: true,
            tshow: true,
            layout: 'horizontal',
            onClick: () => {
                this.onToggle();
                // 切换完立刻回读一次：视图模式不是宿主状态，不会自己走总线广播
                this.refresh();
            },
        });

        this.bus.on(EDIT_STATE_CHANGED, () => this.refresh());
        this.refresh();
    }

    /**
     * 回读高亮
     * @private
     */
    private refresh(): void {
        this.toggleBtn.setActived(!!this.isActive());
    }

    /**
     * 取分区元素
     * @returns 分区容器
     */
    public getElements(): { sectionDocView: HTMLDivElement } {
        return { sectionDocView: this.section };
    }
}
