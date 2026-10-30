/*  # 工作区路由 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目目标: 主页面同一时刻只显示一个工作区（表格 / 文档），二者共用同一套工具栏。
    # 本文件职责: 工作区切换的唯一入口 —— 显隐两个容器，并把命令总线重新绑定到
    #             对应工作区的编辑宿主。**切换的全部副作用都集中在这一个类里**，
    #             使 `App` 不必知道「切到文档时还要做哪几件事」。
    # 依赖方向: contract（EditHost / CommandBus）+ document/docWorkspace。
    #           反向不成立：文档工作区不认识路由，也不认识表格。
    # 为什么需要它: `App` 原本有四处假定「主页面永远是表格」（初始化 / 激活文件 /
    #             数据监听 / 打开器的装载目标）。把切换收进一个类后，新增加的工作区
    #             只需在这里登记，而不必再改 `App` 的那四处。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { CommandBus } from '../contract/commandBus';
import { EditHost, WorkspaceKind } from '../contract/editHost';
import { DocWorkspace } from '../document/docWorkspace';

/** 路由配置 */
export interface WorkspaceRouterOptions {
    /** 命令总线（切换工作区时重新 bind） */
    bus: CommandBus;
    /** 表格宿主 */
    sheetHost: EditHost;
    /** 表格工作区容器（`.sheets` 元素：行头 / 列头 / 画布 / 滚动条都在里面） */
    sheetContainer: HTMLElement;
    /** 文档工作区 */
    doc: DocWorkspace;
    /**
     * 切回表格时的补算回调
     *
     * 表格容器在文档显示期间是 `display:none`，量不到尺寸；切回来必须由
     * `Sheets` 自己按真实尺寸重排一次（见 `Sheets.refreshLayout`）。
     * 用回调而不是直接持有 `Sheets`：本文件在文档域，**不 import 表格**。
     */
    onShowSheet?: () => void;
    /**
     * 切到文档工作区时的补算回调
     *
     * 导航面板与表格共用同一个实例：切到文档时由外壳把它搬进文档的右侧工具线
     * （见 `DocWorkspace.attachNavPanel`）。用回调而不是直接持有导航器：
     * 导航器晚于路由创建（见 `App.initNavigator`）。
     */
    onShowDoc?: () => void;
}

/**
 * 工作区路由
 */
export class WorkspaceRouter {
    /** 命令总线 */
    private readonly bus: CommandBus;

    /** 表格宿主 */
    private readonly sheetHost: EditHost;

    /** 表格容器 */
    private readonly sheetContainer: HTMLElement;

    /** 文档工作区 */
    private readonly doc: DocWorkspace;

    /** 当前工作区 */
    private kind: WorkspaceKind = 'sheet';

    /** 切回表格时的补算回调 */
    private readonly onShowSheet?: () => void;

    /** 切到文档时的补算回调 */
    private readonly onShowDoc?: () => void;

    /**
     * 构造函数
     * @param options - 路由配置
     */
    constructor(options: WorkspaceRouterOptions) {
        this.bus = options.bus;
        this.sheetHost = options.sheetHost;
        this.sheetContainer = options.sheetContainer;
        this.doc = options.doc;
        this.onShowSheet = options.onShowSheet;
        this.onShowDoc = options.onShowDoc;
    }

    /** 当前工作区类型 */
    public get active(): WorkspaceKind {
        return this.kind;
    }

    /**
     * 切到表格工作区
     */
    public showSheet(): void {
        this.doc.hide();
        this.sheetContainer.style.display = '';
        // 隐藏期间量不到尺寸（offsetWidth = 0），切回来必须主动补一次重排：
        // 等 ResizeObserver 自己触发不可靠（尺寸没变时它不会回调）
        this.onShowSheet?.();
        if (this.kind === 'sheet') return;
        this.kind = 'sheet';
        this.bus.bind(this.sheetHost);
    }

    /**
     * 切到文档工作区
     *
     * 顺序有讲究：先显隐容器，再 bind 宿主 —— `bind` 会立即广播一次状态变化，
     * 工具栏据此把按钮刷成文档的真实状态；容器还没显示就刷新，量出来的是 0 尺寸。
     */
    public showDoc(): void {
        this.sheetContainer.style.display = 'none';
        this.doc.show();
        // 导航面板搬到文档的右侧工具线上（与表格共用同一个导航实例）
        this.onShowDoc?.();
        if (this.kind === 'doc') return;
        this.kind = 'doc';
        this.bus.bind(this.doc.editHost);
    }

    /**
     * 初始化：表格工作区就地绑定（不触发多余的显隐变化）
     */
    public mountSheet(): void {
        this.kind = 'sheet';
        this.bus.bind(this.sheetHost);
    }
}
