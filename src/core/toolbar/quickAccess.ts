/*  # web 表格顶部快速访问工具栏 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目目标: 撤销 / 重做 / 打印三组「跨载体」的快捷入口，固定在工具栏最左侧。
    # 本文件职责: 只建按钮、发命令、按快照置灰。**不认识任何数据层**。
    # 依赖方向: contract/commandBus + utils/printer（打印实例由装配层注入）。
    #           改造前本文件直接持有 `DataCollection` 并调 `data.undo()` / `data.undoManager`，
    #           于是「通用快捷栏」被迫认识表格数据集合；现在撤销/重做走命令总线，
    #           可用状态读快照的 `undoable` / `redoable`，两种载体共用同一份 UI。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { createDiv } from "../../utils/dom";
import { TOOLBAR_CLASS_NAME } from "../constant";
import { BtnBase } from "../common/button";
import { Menu } from "../common/menu";
import { CommandBus } from "../contract/commandBus";
import { EDIT_STATE_CHANGED } from "../contract/editHost";
import { Printer } from "../sheets/sheetPrinter";
import favicon from "../../assets/images/favicon.png";
import '../style/brand.scss'


export interface QuickAccessOptions {
    parentElement: HTMLElement;
    /** 命令总线：撤销 / 重做走「发命令」，可用状态走「读快照」 */
    bus: CommandBus;
    menu: Menu;
    /** 共享打印实例（装配层注入，与「页面」页签的打印设置共用） */
    printer: Printer;
    /**
     * 打印当前**文档**（文档工作区用）
     *
     * 与 `printer` 的关系：共享打印器只认表格（它的分页是按单元格网格算的），
     * 文档有自己的一套（按纸张分页 + iframe 打印），因此走注入回调而不是实例。
     */
    onPrintDoc?: () => void;
    /** 打印预览当前文档（文档工作区用） */
    onPrintPreviewDoc?: () => void;
    title?: string;
}


export class QuickAccess {
    private quickAccess: HTMLDivElement;
    /** 撤销/重做/打印/打印预览的容器：窄屏时整组收进浮层 */
    private panel: HTMLDivElement;
    private bus: CommandBus;
    private menu: Menu;
    private printer: Printer;
    /** 打印文档的回调（文档工作区注入） */
    private readonly onPrintDoc?: () => void;
    /** 打印预览文档的回调（文档工作区注入） */
    private readonly onPrintPreviewDoc?: () => void;
    /** 撤销 / 重做按钮（按快照置灰） */
    private backBtn!: BtnBase;
    private forwardBtn!: BtnBase;
    /**
     * 打印 / 打印预览按钮
     *
     * 两个载体**都显示**：Word 与 Excel 的快速访问栏里都有这两个按钮，
     * 按工作区换的是背后的实现（表格走共享打印器，文档走注入回调），不是显隐。
     */
    private printBtn!: BtnBase;
    private printViewBtn!: BtnBase;
    constructor(options: QuickAccessOptions) {
        this.quickAccess = createDiv({
            className: TOOLBAR_CLASS_NAME.QUICK_ACCESS,
        });        
        if (options.parentElement) {
            options.parentElement.append(this.quickAccess);
        }
        this.bus = options.bus;
        this.menu = options.menu;
        this.printer = options.printer;
        this.onPrintDoc = options.onPrintDoc;
        this.onPrintPreviewDoc = options.onPrintPreviewDoc;
        this.init();

        // 点空白处收起浮层（根折叠时按钮组是浮层，不点它就会一直挂着）
        document.addEventListener('pointerdown', (e: PointerEvent) => {
            if (!this.quickAccess.classList.contains(TOOLBAR_CLASS_NAME.QUICK_OPEN)) return;
            const target = e.target as Node | null;
            if (target && this.quickAccess.contains(target)) return;
            this.quickAccess.classList.remove(TOOLBAR_CLASS_NAME.QUICK_OPEN);
        });
    }

    public init() {        
        /*------------------ 快速访问分区 ------------------*/

        const brand = createDiv({className: TOOLBAR_CLASS_NAME.BRAND});
        const brandImg = new Image();
        brandImg.src = favicon;  // 添加图片地址
        brand.append(brandImg);
        brand.title = '快速访问';
        this.quickAccess.append(brand);

        // 窄屏（工具栏根折叠）时只露品牌图标，点它才展开这一组按钮
        brand.addEventListener('click', () => {
            this.quickAccess.classList.toggle(TOOLBAR_CLASS_NAME.QUICK_OPEN);
        });

        // 按钮统一放进 panel：宽屏时它只是普通的一行 flex 项，
        // 根折叠时才变成品牌下方的浮层（样式见 toolbar.scss）
        this.panel = createDiv({ className: TOOLBAR_CLASS_NAME.QUICK_PANEL });
        this.quickAccess.append(this.panel);

        this.backBtn = new BtnBase(  // 撤销按钮
            {
                parentElement: this.panel, 
                stateSource: this.bus,
                isDisabledOnEditting: false,
                menu: this.menu,
                toDo: 'back',
                icon: 'icon-undo',
                text: '撤销 Ctrl + Z',
                ishow: true,
                layout: 'vertical',
                onClick: () => this.bus.dispatch('undo'),
            }
        );

        this.forwardBtn = new BtnBase(  // 恢复按钮
            {
                parentElement: this.panel,
                stateSource: this.bus,
                isDisabledOnEditting: false,
                menu: this.menu,
                toDo: 'forward',
                icon: 'icon-redo',
                text: '恢复 Ctrl + Y',
                ishow: true,
                layout: 'vertical',
                onClick: () => this.bus.dispatch('redo'),
            }
        );

        this.printBtn = new BtnBase(  // 打印按钮
            {
                parentElement: this.panel,
                stateSource: this.bus,
                menu: this.menu,
                toDo: 'print',
                icon: 'icon-printer',
                text: '打印',
                ishow: true,
                layout: 'vertical',
                onClick: () => {
                    if (this.bus.kind === 'doc') this.onPrintDoc?.();
                    else this.printer.print();
                },
            }
        );

        this.printViewBtn = new BtnBase(  // 打印预览按钮
            {
                parentElement: this.panel,
                stateSource: this.bus,
                menu: this.menu,
                toDo: 'printView',
                icon: 'icon-printer-view',
                text: '打印预览',
                ishow: true,
                layout: 'vertical',
                onClick: () => {
                    if (this.bus.kind === 'doc') this.onPrintPreviewDoc?.();
                    else this.printer.printPreview();
                },
            },
        );

        this.setToolbarsState();
        this.bus.on(EDIT_STATE_CHANGED, () => this.setToolbarsState());
    }

    /**
     * 按快照刷新按钮可用性
     *
     * 撤销 / 重做是否可点由宿主回答（两种载体的撤销栈类型不同，快照里已归一化为布尔值）。
     * 打印两个按钮**不再按工作区隐藏**：两类工作区都能打印，只是走不同的实现
     * （表格走共享打印器，文档走注入回调，见 init 里的两个 onClick）。
     * @private
     */
    private setToolbarsState(): void {
        const state = this.bus.snapshot();
        this.backBtn.setDisabled(!state.undoable);
        this.forwardBtn.setDisabled(!state.redoable);
    }

    public getElements() {
        return {
            container: this.quickAccess,
            panel: this.panel,
        }
    }
}
