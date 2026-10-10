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
    DEFAULT_FONT_FAMILY,
    DEFAULT_FONT_SIZE,
    FONT_FAMILY_LIST,
    DataEvents
} from "./constant"; 
import { Split } from "./common/split";
import { Navigator } from "./common/navigator";
import { FileTree } from "./common/fileTree";
import { Sheets } from "./sheets/sheets";
import { DataCollection } from "./dataArchitecture/dataCollection";
import { Scroller } from "./common/scroller";
import { Toolbar } from "./toolbar/toolbar";
import { Menu } from "./common/menu";
import { QuickAccess } from "./toolbar/quickAccess";
import favicon from "../assets/images/favicon.ico";


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
            nextSize: 360   // 左右分栏不默认平分：右侧面板（导航内容区）默认宽度 360px
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
        const navigator = new Navigator(false);
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
     * 初始化「文件」页签：树形文件页面
     * 顶部为搜索栏（同行右侧为展开 / 收缩 / 定位按钮），下方为树形文件列表
     */
    private initFileTree(navigator: Navigator) {
        this.fileTree = new FileTree();
        navigator.navContentItemMount(NAV_PANEL[0], this.fileTree.getElements().container);   // NAV_PANEL[0] 即「文件」
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
        this.hScroller = new Scroller({
            parentElement: this.bottomToolsSplit.getElements().splitNext,
            direction: 'horizontal',
            data: this.data,
        });
        this.vScroller = new Scroller({
            parentElement: this.rightToolsSplit.getElements().splitNext,
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