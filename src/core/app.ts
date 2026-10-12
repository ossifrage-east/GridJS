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
import { FileTree } from "./common/fileTree";
import { FileOpener } from "./common/fileOpener";
import { Sheets } from "./sheets/sheets";
import { DataCollection } from "./dataArchitecture/dataCollection";
import { IndexDB } from "./dataArchitecture/indexDB";
import { Scroller } from "./common/scroller";
import { Toolbar } from "./toolbar/toolbar";
import { Menu } from "./common/menu";
import { QuickAccess } from "./toolbar/quickAccess";
import favicon from "../assets/images/favicon.ico";

/** 导航内容区宽度在 IndexDB 中的键前缀（完整键 = 前缀 + 导航项文本） */
const NAV_WIDTH_KEY_PREFIX = 'nav-content-width:';


export class App {
    private worktop!: HTMLDivElement;
    private data: DataCollection = new DataCollection();
    private containerSplit!: Split;
    private sheets!: Sheets;
    private bottomToolsSplit!: Split;
    private rightToolsSplit!: Split;
    private hScroller!: Scroller;
    private vScroller!: Scroller;
    private toolbar!: Toolbar;
    private menu!: Menu;
    private navigator!: Navigator;
    private fileTree!: FileTree;
    private fileOpener!: FileOpener;
    /** 导航内容区宽度的持久化存储（键：NAV_WIDTH_KEY_PREFIX + 导航项文本） */
    private readonly navWidthDb: IndexDB = new IndexDB();
    /** 当前显示的导航项文本：拖拽保存宽度时作为键 */
    private currentNavItemText = '';

    /**
     * 初始化工具栏
     */
    constructor(id: string) {
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
        this.initbottomToolsSplit();
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
     */
    private initToolbar() {
        this.toolbar = new Toolbar({
            parentElement: this.containerSplit.getElements().splitPrev,
            data: this.data,
            menu: this.menu,
            onTabContentShow: (width: number, height: number) => {
                const rect = this.sheets.getElements().sheets.parentElement.getBoundingClientRect();
                // this.sheets.getElements().sheets.style.width = `${rect.width - width}px`;
                this.sheets.getElements().sheets.style.height = `${rect.height - height}px`;
            }
        });
        const quickAccess = new QuickAccess({  // 快速访问分区
            parentElement: this.toolbar.getElements().tabs,
            data: this.data,
            menu: this.menu,
            title: '快速访问'
        });
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
        this.initFileTree(navigator);
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
        this.navWidthDb.getSetting(NAV_WIDTH_KEY_PREFIX + navItemText).then((saved: unknown) => {
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
        this.navWidthDb.setSetting(
            NAV_WIDTH_KEY_PREFIX + navItemText,
            Math.round(width),
            'nav-content-width'
        ).catch(() => {
            // 持久化失败不影响交互
        });
    }

    /**
     * 初始化「文件」页签：树形文件页面
     *
     * 顶部为搜索栏（同行右侧为展开 / 收缩 / 定位按钮），下方为树形文件列表；
     * 新建 / 打开 / 剪切粘贴等操作统一通过树节点右键菜单完成
     */
    private initFileTree(navigator: Navigator) {
        // 文件打开器：负责「打开」按钮与主页面拖放，解析成功后把文件名登记到文件树
        this.fileOpener = new FileOpener({
            data: this.data,
            onOpened: (fileName: string) => {
                this.fileTree.addNode(fileName);
                this.data.messageBox.success(`已打开 ${fileName}`);
            },
            onError: (message: string) => this.data.messageBox.error(message),
        });
        document.body.appendChild(this.fileOpener.getFileInput());

        this.fileTree = new FileTree({
            menu: this.menu,
            onOpen: () => this.fileOpener.pickAndOpen(),
        });
        navigator.navContentItemMount(NAV_PANEL[0], this.fileTree.getElements().container);   // NAV_PANEL[0] 即「文件」
        // 「定位」以主页面当前显示的文件名为目标：随工作表切换同步
        this.fileTree.setActivedFileName(this.data.activedSheetName);
        this.data.on(DataEvents.ACTIVED_SHEET_NAME_CHANGED, (_, name: string) => {
            this.fileTree?.setActivedFileName(name);
        });

        // 主页面拖放：拖入 xlsx/xls/docx/doc/pptx/ppt 即自动打开
        this.fileOpener.bindDropZone(this.sheets.getElements().sheets);
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

    private setupDataListeners() {
            this.data.on(DataEvents.ACTIVED_CELL_CHANGED, (_, cell) => {
                this.data.isEditting = false;
                // this.sheets.draw(); // 与SELECT_CHANGED事件冲突
            });
            this.data.colHeaders.on(DataEvents.OFFSETX_CHANGED, () => {
                this.sheets.draw();
            });
            this.data.on(DataEvents.ALL_COL_WIDTH_CHANGED, () => {
                this.sheets.draw();
            });
            this.data.rowHeaders.on(DataEvents.OFFSETY_CHANGED, () => {
                this.sheets.draw();
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
                this.sheets.draw();
            });
            this.data.on(DataEvents.VALUES_CHANGED, (newValues) => {
                this.sheets.draw();
            });
            this.data.on(DataEvents.COLUMNS_CHANGED, () => {
                this.sheets.draw();
            });
            this.data.on(DataEvents.ROWS_CHANGED, () => {
                this.sheets.draw();
            });
        // this.data.on(DataEvents.FONT_FAMILY_CHANGED, () => {
        //     this.sheets.draw();
        // });
        // this.data.on(DataEvents.FONT_SIZE_CHANGED, () => {
        //     this.sheets.draw();
        // });
        // this.data.on(DataEvents.FONT_WEIGHT_CHANGED, () => {
        //     this.sheets.draw();
        // });
        // this.data.on(DataEvents.FONT_STYLE_CHANGED, () => {
        //     this.sheets.draw();
        // });
        // this.data.on(DataEvents.FONT_COLOR_CHANGED, () => {
        //     this.sheets.draw();
        // });
        // this.data.on(DataEvents.UNDERLINE_CHANGED, () => {
        //     this.sheets.draw();
        // });
        // this.data.on(DataEvents.STRIKETHROUGH_CHANGED, () => {
        //     this.sheets.draw();
        // });
        document.addEventListener('contextmenu', function (e) {
            e.preventDefault(); // 阻止默认右键菜单
        });
    }
}