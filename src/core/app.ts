/*  # 表格接口 TS 文件
    # 项目名称: 2023年1月1日起正式更名为 silkspaces
    # 项目地址: 
    # 项目描述: 基于 HTML5 的 Excel 表格组件，无需服务器端支持，即可在浏览器中创建、读取、写入 Excel 文件。
    # 项目目标: 提供一个简单、易用、功能完善的 Excel 表格组件，满足用户在浏览器中处理 Excel 文件的需求。
    # 项目状态: 开发中
    # 项目维护者: 东方鹗
    # 项目许可证: MIT 许可证
    # 项目版本: 0.0.1
    # 项目开始日期: 2023年1月1日
    # 吐槽: 中国的教育是在吃人，学生的学习进度是被控制的，学生的学习能力是被限制的，学生的学习兴趣是被抑制的。
    # 一个简单的课堂纪律，每天都得由学生填报，小到椅子没有摆放整齐，试想一下，一个人天天这样被控制，是什么心态，
    # 而且老师还会通过即时通讯工具，发给每一个学生家长。无视孩子的青春叛逆期，徒增学生与家长的矛盾，这是中国教育的一个严重问题。
    # 而老师的心中只是家教不好。
    # 或许是错怪了老师，当学生的学习进度被控制时，老师的压力也会增加。教材是一部分专家编辑的，考试卷又是另一部分专家写的，课外辅导教程又是另一部分人，
    # 于是学生的学习成了一部分“专家”的炫耀技巧的场所，这些人其心可诛。
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

/*  # 工作区职责拆分（K4）
    # 本文件只剩「布局装配 + 导航细节」；文件会话 / 打开分派 / 内容桥 / 导出已下沉到 core/workspace/*
    # 继承链：App → ExportActionsBase → ContentBridgeBase → OpenRouterBase → FileSessionBase
*/

import { TOOLS_CLASS_NAMES, DEFAULT_CELL_HEIGHT, ROW_HEADER_PADDING,
    NAV_PANEL,
    NAV_CONTENT_WIDTH,
    DEFAULT_FONT_FAMILY,
    DEFAULT_FONT_SIZE,
    FONT_FAMILY_LIST,
    DataEvents
} from "./constant";
import { Split } from "./common/split";
import { Navigator } from "./common/navigator";
import { PYEditor } from "./assistant/pyEditor";
import { Sheets } from "./sheets/sheets";
import { SelectionCalc } from "./sheets/selectionCalc";
import { Scroller } from "./common/scroller";
import { Slider } from "./common/slider";
import { Toolbar } from "./toolbar/toolbar";
import { Menu } from "./common/menu";
import { QuickAccess } from "./toolbar/quickAccess";
import { getSharedPrinter } from "./sheets/sheetPrinter";
import { CommandBus } from "./contract/commandBus";
import { SheetEditHost } from "./sheets/sheetEditHost";
import { openFindReplace } from "./toolbar/start/searchDialog";
import { DocWorkspace } from "./document/docWorkspace";
import { WorkspaceRouter } from "./workspace/workspaceRouter";
import { ExportActionsBase } from "./workspace/exportActions";
import favicon from "../assets/images/favicon.ico";

/** 导航内容区宽度在 IndexDB 中的键前缀（完整键 = 前缀 + 导航项文本） */
const NAV_WIDTH_KEY_PREFIX = 'nav-content-width:';

/** 缩放滑块取值下限（百分比）：与 Ctrl+滚轮 缩放的下限同源 */
const ZOOM_SLIDER_MIN = Math.round(Scroller.ZOOM_MIN * 100);
/** 缩放滑块取值上限（百分比）：与 Ctrl+滚轮 缩放的上限同源 */
const ZOOM_SLIDER_MAX = Math.round(Scroller.ZOOM_MAX * 100);

/**
 * 装配器
 *
 * 职责只剩布局装配（分栏 / 菜单 / 工具栏 / 导航器 / 滚动条 / 工作区高度 / 缩放滑块）
 * 与导航细节（面板搬动 / 宽度记忆）。文件会话、打开分派、内容桥、导出见 `core/workspace/`。
 */
export class App extends ExportActionsBase {
    private worktop!: HTMLDivElement;
    private bottomToolsSplit!: Split;
    private hScroller!: Scroller;
    private vScroller!: Scroller;
    /**
     * 滚动/缩放引起的重绘合帧句柄
     *
     * 横向与纵向滚动在同一帧内可能各发一次 OFFSET 事件，拖拽滚动条时
     * 事件频率还会高于显示器刷新率。用 requestAnimationFrame 合并后，
     * 一帧最多执行一次 sheets.draw()，避免重复全量重绘导致的滚动卡顿。
     */
    private _sheetDrawFrame: number | null = null;
    private toolbar!: Toolbar;
    /**
     * 命令总线：工具栏外壳与当前编辑宿主之间的唯一通道
     *
     * 外壳（工具栏 / 按钮 / 下拉）只认识它，不认识 `DataCollection`。
     * 将来新增文档工作区时，只需 `bind` 一个文档宿主，UI 一行都不用改。
     */
    private commandBus!: CommandBus;
    /**
     * 表格编辑宿主：把「粗体 / 对齐 / 粘贴 / 合并…」翻译成 `DataCollection` 调用
     *
     * 与 {@link commandBus} 一对一；工作区切换时换掉宿主即可（`bind` 会退订旧的）。
     */
    private sheetEditHost!: SheetEditHost;
    /**
     * 选区即时计算模块（底部栏即时值 + 活动格内统计）
     *
     * 工具栏先于本模块创建，因此 {@link initToolbar} 用 `onCalculate` 回调
     * 延迟取用这个实例（点击菜单时才真正访问）。
     */
    private selectionCalc: SelectionCalc | null = null;

    /**
     * 初始化工具栏
     */
    constructor(id: string) {
        super();
        const link = document.createElement('link');
        link.rel = 'icon';
        link.type = 'image/x-icon';
        link.href = favicon;
        document.head.appendChild(link);
        this.initUI(id);
    }
    /**
     * 初始化应用
     */
    private initUI(id: string) {
        const worktop = document.querySelector(`${id}`);
        if (!worktop) {
            throw new Error(`Element with id ${id} not found`);
        }
        this.worktop = worktop as HTMLDivElement;  // 工作区域
        this.initContainerSplit();
        this.initMenu();
        this.initToolbar();
        this.initSheets();
        this.initDocWorkspace();   // 依赖 initSheets（路由要拿表格容器）与 initToolbar（总线）
        this.initbottomToolsSplit();
        this.initSelectionCalc();   // 依赖 bottomToolsSplit 的 split-pre，必须在其之后
        this.initRightToolsSplit();
        this.initScroller();
        this.setupDataListeners();
        this.initNavigator(); // 初始化导航器，必须放到最后
    }

    /**
     * 初始化容器分割线
     */
    private initContainerSplit() {
        this.containerSplit = new Split({
            parentElement: this.worktop,
            id: TOOLS_CLASS_NAMES.CONTAINER,
            storageKey: TOOLS_CLASS_NAMES.CONTAINER,
            nextSize: 360,   // 左右分栏不默认平分：右侧面板（导航内容区）默认宽度 360px
            // 仅在真实拖拽结束后保存：容器尺寸变化 / 程序化设置宽度不会污染已保存值
            onDragEnd: (nextSize: number) => this.saveNavContentWidth(nextSize),
        });
        this.containerSplit.handleVisibility();   // 元素是否可见监控
    }
    
    /**
     * 初始化菜单
     */
    private initMenu() {
        this.menu = new Menu({
            parentElement: this.containerSplit.getElements().splitPrev,
        });
    }

    /**
     * 初始化工具栏
     *
     * ## 装配顺序有讲究
     *
     * 1. 先建宿主、再建总线、**先 bind**：分区在构造时就订阅总线，
     *    绑定晚于构造会漏掉 bind 那一发的初始广播。
     * 2. 数据是从 IndexedDB 异步读出来的，各分区构造时读到的还是空表 ——
     *    等 `dataLoaded` 兑现后补一次广播，让按钮刷新成真实状态
     *    （改造前这件事由每个分区各自调一次 `syncValuesFromDB()` 完成，三处重复读库）。
     */
    private initToolbar() {
        this.sheetEditHost = new SheetEditHost(this.data);
        this.commandBus = new CommandBus();
        this.commandBus.bind(this.sheetEditHost);

        this.toolbar = new Toolbar({
            parentElement: this.containerSplit.getElements().splitPrev,
            bus: this.commandBus,
            menu: this.menu,
            printer: getSharedPrinter(this.data),
            onTabContentShow: (width: number, height: number) => {
                // 表格与文档两个工作区共用同一条尺寸规则（见 applyWorkspaceHeight）
                this.applyWorkspaceHeight(height);
            },
            // 计算菜单点选 → 把该项结果写入空白的活动单元格（实例稍后创建，回调按需取用）
            onCalculate: (label: string) => this.selectionCalc?.applyToActivedCell(label),
            // 查询 / 替换是对话框而非一次编辑动作，需要数据源全表扫描，因此由装配层注入
            onFindReplace: (mode: 'find' | 'replace') => openFindReplace(this.data, mode),
            // 视图切换同理：它是「工作区怎么显示」，不是一次编辑动作。文档工作区在
            // initDocWorkspace 里才创建（晚于工具栏），因此用闭包延迟取用。
            onToggleMarkdown: () => this.docWorkspace?.toggleMode(),
            isMarkdownActive: () => this.docWorkspace?.isMarkdown ?? false,
            // 页面设置同理属于「纸怎么摆」：读回与打印都延迟取文档工作区（它晚于工具栏创建）
            getDocPageSetup: () => {
                const setup = this.docWorkspace?.pageSetup;
                return {
                    orientation: setup?.orientation ?? 'portrait',
                    paperSize: setup?.paperSize ?? 'A4',
                    marginPreset: setup?.marginPreset ?? 'normal',
                    scaleMode: setup?.scaleMode ?? 'none',
                };
            },
            onChangeDocPage: (key, value) => this.docWorkspace?.setPageSetup(key, value),
            onPrintDoc: () => this.docWorkspace?.print(),
        });
        const quickAccess = new QuickAccess({  // 快速访问分区
            parentElement: this.toolbar.getElements().tabs,
            bus: this.commandBus,
            menu: this.menu,
            printer: getSharedPrinter(this.data),
            // 文档工作区晚于快速访问栏创建（见 initDocWorkspace），同样用闭包延迟取用
            onPrintDoc: () => this.docWorkspace?.print(),
            onPrintPreviewDoc: () => this.docWorkspace?.printPreview(),
        });
        // 快速访问区是构造完工具栏之后才挂进页签行的：页签行凭空多了一截宽度，
        // 但内容区尺寸没变（ResizeObserver 不会触发），得主动重排一次，
        // 否则「加上快速访问区之后已经装不下」这种情况要等到下一次窗口变化才被发现。
        this.toolbar.refreshLayout();

        // 数据装载完成 → 通知外壳按真实状态刷新一次（见方法头注释第 2 点）
        void this.data.dataLoaded.then(() => this.commandBus.notifyStateChanged());
    }

    /**
     * 初始化导航器
     */
    private async initNavigator() {
        const navigator = new Navigator({
            horizontal: false,
            // 导航项内容区变为可见时，恢复其宽度（已保存值优先，否则用初始化值）
            onItemShow: (navItemText: string) => this.restoreNavContentWidth(navItemText),
        });
        this.navigator = navigator;
        const { panel, content } = navigator.getElements();
        const { splitNext } = this.containerSplit.getElements();
        splitNext.appendChild(content);
        this.rightToolsSplit.getElements().splitPrev.appendChild(panel);
        // navItemMount 内部读取 IndexDB（异步）：必须等待页签与内容项创建完成，
        // 否则 navContentItemMount 会因找不到目标内容项而挂载失败
        try {
            await navigator.navItemMount(...NAV_PANEL);
        } catch {
            // IndexDB 不可用时仅跳过状态恢复，页签元素已创建，继续挂载各页签内容
        }
        await this.initFileTree(navigator);
        this.initPyEditor(navigator);
        this.initZoomSlider();   // 缩放滑块须在导航面板之后挂载：DOM 顺序即视觉顺序（面板在上、滑块在下）
    }

    /**
     * 把导航面板交给文档工作区（切到文档时调用）
     *
     * 两个工作区不会同时显示，因此**共用同一个 `Navigator` 实例**：面板元素在
     * 两条右侧工具线之间搬动，导航内容区（`containerSplit.splitNext`）不动。
     * 这样不会产生两个同 id 的面板，也不会出现两份互相覆盖的选中状态。
     */
    private moveNavPanelToDoc(): void {
        if (!this.navigator || !this.docWorkspace) return;
        this.docWorkspace.attachNavPanel(this.navigator.getElements().panel);
    }

    /**
     * 把导航面板搬回表格工作区（切回表格时调用）
     *
     * 面板要插在表格缩放滑块**之前**：与初始挂载时的 DOM 顺序保持一致
     * （导航面板在上、缩放滑块在下）。
     */
    private moveNavPanelToSheet(): void {
        if (!this.navigator) return;
        this.docWorkspace?.detachNavPanel();
        const panel = this.navigator.getElements().panel;
        const prev = this.rightToolsSplit.getElements().splitPrev;
        const slider = this.zoomSlider.getElements().container;
        if (slider.parentElement === prev) prev.insertBefore(panel, slider);
        else prev.appendChild(panel);
    }

    /**
     * 恢复指定导航项的内容区宽度
     *
     * 优先级：IndexedDB 中已保存的宽度（人工拖拽产生） > `NAV_CONTENT_WIDTH`
     * 中的初始化宽度 > 容器宽度平分。
     * @param navItemText - 导航项文本
     */
    private restoreNavContentWidth(navItemText: string): void {
        const split = this.containerSplit;
        if (!split) return;
        // 记录当前导航项，后续拖拽保存宽度以此为键
        this.currentNavItemText = navItemText;
        const containerWidth = split.getElements().container.clientWidth;
        if (!containerWidth) return;

        // 读取已保存宽度（异步，不阻塞切换：先用初始化值兜底，读取到后再校正）
        this.settingsDb.getSetting(NAV_WIDTH_KEY_PREFIX + navItemText).then((saved: unknown) => {
            if (typeof saved === 'number' && Number.isFinite(saved) && saved > 0) {
                split.setNextSize(saved);
            }
        }).catch(() => {
            // IndexDB 不可用时忽略，退回初始化值
        });

        const initWidth = NAV_CONTENT_WIDTH[navItemText];
        split.setNextSize(initWidth ?? Math.round(containerWidth / 2));
    }

    /**
     * 持久化当前导航项的内容区宽度（人工拖拽后调用）
     * @param width - 内容区宽度（像素）
     */
    private saveNavContentWidth(width: number): void {
        const navItemText = this.currentNavItemText;
        if (!navItemText || !Number.isFinite(width) || width <= 0) return;
        this.settingsDb.setSetting(
            NAV_WIDTH_KEY_PREFIX + navItemText,
            Math.round(width),
            'nav-content-width'
        ).catch(() => {
            // 持久化失败不影响交互
        });
    }

    /**
     * 初始化「脚本」页签：Python 编辑器
     *
     * 编辑器内容与状态由组件自身持久化（见 {@link PYEditor}），
     * 这里只负责把组件挂进导航内容区。
     * @param navigator - 右侧导航器实例
     */
    private initPyEditor(navigator: Navigator): void {
        this.pyEditor = new PYEditor({ data: this.data });
        navigator.navContentItemMount(NAV_PANEL[2], this.pyEditor.getElements().container);   // NAV_PANEL[2] 即「脚本」
    }

    /**
     * 初始化底部工具线
     */
    private initbottomToolsSplit() {
        this.bottomToolsSplit = new Split({
            parentElement: this.sheets.getElements().sheets, 
            id: TOOLS_CLASS_NAMES.BOTTOM_TOOLS, 
            storageKey: TOOLS_CLASS_NAMES.BOTTOM_TOOLS
        });
    }

    /**
     * 初始化选区即时计算（求和 / 平均 / 计数 / 最大 / 最小）
     *
     * 只要存在选区就显示即时值（底部工具条 split-pre 内部）；
     * 点击工具栏「计算」菜单的某一项时，才把该项结果写入空白的活动单元格。
     * 由 {@link SelectionCalc} 自行监听选区 / 内容变化并防抖重算。
     */
    private initSelectionCalc() {
        const selectionCalc = new SelectionCalc({
            data: this.data,
            parentElement: this.bottomToolsSplit.getElements().splitPrev,
        });
        this.selectionCalc = selectionCalc;   // 供工具栏「计算」菜单点选时使用
        selectionCalc.refresh();   // 首屏选区可能早于监听建立，主动补算一次
    }

    /**
     * 初始化右侧工具线
     */
    private initRightToolsSplit() {
        this.rightToolsSplit = new Split({
            parentElement: this.sheets.getElements().sheets, 
            id: TOOLS_CLASS_NAMES.RIGHT_TOOLS,
            horizontal:false,
            storageKey: TOOLS_CLASS_NAMES.RIGHT_TOOLS
        });
    }

    /**
     * 初始化缩放滑块（挂载于 rightTools 的 split-pre 内部、贴底显示）
     *
     * 采用垂直细线样式（{@link Slider} 的 vertical + thin），外观由 slider.scss 的主题色板控制；
     * 垂直布局下数值输入框自然落在轨道下方，可手工填写百分比缩放值。
     * 取值与 `data.zoom` 双向绑定：
     * - 拖动滑块 / 填写数值框 → 写回 `data.zoom`（百分比 ÷ 100）；
     * - Ctrl+滚轮 缩放 → 触发 ZOOM_CHANGED → 回写滑块与数值框。
     * 取值范围直接取自 {@link Scroller} 的缩放上下限，避免两处常量漂移。
     */
    private initZoomSlider() {
        this.zoomSlider = new Slider({
            parentElement: this.rightToolsSplit.getElements().splitPrev,
            vertical: true,      // 垂直方向：最小值在下端，数值框在轨道下方
            thin: true,          // 细线样式：2px 轨道 + 小圆点手柄
            min: ZOOM_SLIDER_MIN,
            max: ZOOM_SLIDER_MAX,
            step: 1,
            value: this.zoomPercent(),
            showInput: true,     // 数值框：手工填写缩放百分比（无单位文本，纯数字）
            inputWidth: '',      // 置空：宽度交给样式表（26px 窄条需占满整行）
            accentColor: '',     // 置空走样式表主题色，避免注入内联的浏览器默认蓝
            onChange: (value: number) => {
                this.data.zoom = value / 100;
            },
        });

        // Ctrl+滚轮 或其它入口改变缩放时，同步滑块位置
        this.data.on(DataEvents.ZOOM_CHANGED, () => {
            this.zoomSlider.setValue(this.zoomPercent());
        });
    }

    /**
     * 当前缩放比例对应的百分比整数值（滑块取值单位与 data.zoom 不同，此处统一换算）
     * @returns {number} 百分比整数，如 1.21 → 121
     */
    private zoomPercent(): number {
        return Math.round(this.data.zoom * 100);
    }

    /**
     * 初始化工作表
     */
    private initSheets() {
        const { splitPrev } = this.containerSplit.getElements();
        this.sheets = new Sheets({
            parentElement: splitPrev,
            data: this.data,
            menu: this.menu,
            onResize: (left: number, top: number, width: number, height: number) => {
                const bottomToolsSplitRect  =  this.bottomToolsSplit.getElements().container.getBoundingClientRect();
                const rightToolsSplitRect = this.rightToolsSplit.getElements().container.getBoundingClientRect();
                this.data.visibleView.sheetWidth = Math.round(width - rightToolsSplitRect.width - this.data.visibleView.rowHeaderWidth);
                this.data.visibleView.sheetHeight = Math.round(height - bottomToolsSplitRect.height - this.data.visibleView.colHeaderHeight);
                this.sheets.updateCanvasSize();
            }
        });
    }

    /**
     * 同步表格 / 文档两个工作区的高度
     *
     * 两者都挂在 `containerSplit.splitPrev` 里、排在工具栏之下，而工具栏高度随窗口宽度
     * 变化（紧凑模式会变矮）。若沿用 `height: 100%`，工作区会多占工具栏那一截，
     * 底部工具线就被顶出可视区（文档滚动条「位置不对」正是这个原因）。
     * 因此与 `.sheet-workspace` 一致：按「分隔容器高度 − 工具栏高度」写成内联像素值。
     * @param toolbarHeight - 工具栏高度（px，由 `tabContentShow` 事件给出）
     */
    private applyWorkspaceHeight(toolbarHeight: number): void {
        const { splitPrev } = this.containerSplit.getElements();
        const available = Math.max(0, Math.round(splitPrev.getBoundingClientRect().height - toolbarHeight));
        if (this.sheets) this.sheets.getElements().sheets.style.height = `${available}px`;
        this.docWorkspace?.setHostHeight(available);
    }

    /**
     * 初始化文档工作区
     *
     * 挂载点与表格相同（`containerSplit.splitPrev`）：工具栏在两者之上，
     * 切换工作区时工具栏不动，只换它下面显示谁 —— 这正是「一套 UI 挂两种载体」的形态。
     */
    private initDocWorkspace() {
        this.docWorkspace = new DocWorkspace({
            parentElement: this.containerSplit.getElements().splitPrev,
            bus: this.commandBus,
        });
        // 工作区晚于工具栏创建：补一次尺寸（`tabContentShow` 首次触发时它还不存在）
        this.applyWorkspaceHeight(this.toolbar.getElements().container.getBoundingClientRect().height);
        this.workspaceRouter = new WorkspaceRouter({
            bus: this.commandBus,
            sheetHost: this.sheetEditHost,
            sheetContainer: this.sheets.getElements().sheets,
            doc: this.docWorkspace,
            // 切回表格要主动补一次重排：表格容器在文档显示期间是 display:none，
            // 那一帧量到的尺寸是 0，画布会停在切换前的旧尺寸上；
            // 同时把导航面板从文档的右侧工具线搬回表格这边
            onShowSheet: () => {
                this.moveNavPanelToSheet();
                this.sheets.refreshLayout();
            },
            onShowDoc: () => this.moveNavPanelToDoc(),
        });
        this.workspaceRouter.mountSheet();
    }

    private initScroller() {
        // 滚轮监听统一挂到工作表容器：滚动条实例只负责各自方向的分量，
        // 使 sheetCanvas/editCanvas/textarea 激活时滚轮也能滚动主页面滚动条
        const wheelElement = this.sheets.getElements().sheets;
        this.hScroller = new Scroller({
            parentElement: this.bottomToolsSplit.getElements().splitNext,
            wheelElement,
            direction: 'horizontal',
            data: this.data,
        });
        this.vScroller = new Scroller({
            parentElement: this.rightToolsSplit.getElements().splitNext,
            wheelElement,
            direction: 'vertical',
            data: this.data,
        });
    }

    /**
     * 调度一次「合帧」的表格重绘
     *
     * 同一动画帧内的多次调用（横向 + 纵向滚动、滚动条拖拽的高频事件）
     * 只会产生一次 sheets.draw()。数据编辑类事件不走这里——它们要求立即
     * 反馈，仍直接调用 draw()。
     * @private
     */
    private scheduleSheetDraw(): void {
        if (this._sheetDrawFrame !== null) return;
        this._sheetDrawFrame = requestAnimationFrame(() => {
            this._sheetDrawFrame = null;
            this.sheets.draw();
        });
    }

    private setupDataListeners() {
            this.data.on(DataEvents.ACTIVED_CELL_CHANGED, (_, cell) => {
                this.data.isEditting = false;
                // this.sheets.draw(); // 与SELECT_CHANGED事件冲突
            });
            this.data.colHeaders.on(DataEvents.OFFSETX_CHANGED, () => {
                this.scheduleSheetDraw();
            });
            this.data.on(DataEvents.ALL_COL_WIDTH_CHANGED, () => {
                this.sheets.draw();
            });
            this.data.rowHeaders.on(DataEvents.OFFSETY_CHANGED, () => {
                this.scheduleSheetDraw();
            });
            this.data.on(DataEvents.ALL_ROW_HEIGHT_CHANGED, () => {
                this.sheets.draw();
            });
            this.data.on(DataEvents.SELECTION_CHANGED, () => {
                this.sheets.draw();
            });
            this.data.on(DataEvents.ANCHOR_SELECTION_CHANGED, () => {
                this.sheets.draw();
            });
            this.data.on(DataEvents.IS_EDITTING_CHANGED, (_, isEditting) => {
                if(!isEditting) this.sheets.draw();
            });
            this.data.on(DataEvents.ZOOM_CHANGED, () => {
                this.scheduleSheetDraw();
            });
            this.data.on(DataEvents.VALUES_CHANGED, (newValues) => {
                this.sheets.draw();
            });
            // 复制/剪切后源范围的流动虚线框：位置由 Sheets.draw 同步（跟随滚动/缩放）
            this.data.on(DataEvents.COPY_RANGE_CHANGED, () => {
                this.sheets.draw();
            });
            this.data.on(DataEvents.COLUMNS_CHANGED, () => {
                this.sheets.draw();
            });
            this.data.on(DataEvents.ROWS_CHANGED, () => {
                this.sheets.draw();
            });
        document.addEventListener('contextmenu', function (e) {
            e.preventDefault(); // 阻止默认右键菜单
        });

        // 自动保存已从「每帧落库」改为「变更后防抖落库」，因此关闭/隐藏页面时必须补一次
        // 立即固化，否则防抖窗口内的最后一次改动会随页面一起丢失。
        // pagehide 覆盖关闭/刷新/前进后退；visibilitychange 覆盖切到后台（移动端可能不触发 pagehide）。
        window.addEventListener('pagehide', () => {
            this.data.flushPendingSave();
            void this.docWorkspace?.flush();
        });
        document.addEventListener('visibilitychange', () => {
            if (document.visibilityState !== 'hidden') return;
            this.data.flushPendingSave();
            void this.docWorkspace?.flush();
        });
    }
}
