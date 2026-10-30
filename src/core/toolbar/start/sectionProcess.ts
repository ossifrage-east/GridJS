/*  # web 表格顶部工具栏数据处理工具集合 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目描述: 基于 HTML5 Canvas 的 Web 电子表格组件，无需服务端支持即可在浏览器中创建、读取、写入 Excel 文件。
    # 项目目标: 工具栏的「数据处理」分区 —— 计算、排序、冻结、筛选、查询/替换。
    # 本文件职责: 建按钮、开菜单、按状态高亮；排序/筛选经 {@link CommandBus} 派发。
    #             「计算」与「查询替换」不走命令总线，而是构造期注入的回调钩子：
    #             前者是「把结果写进活动格」的展示偏好，后者是**对话框**（要拿数据源全表扫描、
    #             还要跨次打开保留关键字与命中游标），都不属于「一次编辑动作」。
    # 依赖方向: contract/commandBus + common/*，**不再 import dataArchitecture / searchDialog**。
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
 * 「计算」菜单项的 todo → 即时计算项名
 *
 * 菜单点选后由宿主把对应项名交给选区即时计算模块，使其只在活动格内显示该项。
 */
const CALC_ITEM_LABELS: Record<string, string> = {
    'sum': '求和',
    'average': '平均',
    'count': '计数',
    'max': '最大',
    'min': '最小',
};

/**
 * SectionProcess 构造选项接口
 * @interface
 * @property {HTMLElement} [parentElement] - 父元素容器
 * @property {CommandBus} bus - 命令总线（编辑宿主的状态与命令入口）
 * @property {Menu} menu - 菜单实例
 * @property {string} [title] - 工具栏标题
 * @property {(label: string) => void} [onCalculate] - 点选求和/平均/计数/最大/最小后的回调
 * @property {(mode: 'find' | 'replace') => void} [onFindReplace] - 点选查询/替换后的回调（拉起对话框）
 */
export interface sectionProcessOptions {
    parentElement?: HTMLElement;
    bus: CommandBus;
    menu: Menu;
    title?: string;
    onCalculate?: (label: string) => void;
    onFindReplace?: (mode: 'find' | 'replace') => void;
}

/**
 * 工具栏数据处理区域类
 * 提供计算、排序、冻结、筛选和查询等数据处理功能按钮
 * @class
 */
export class SectionProcess {
    /** @type {HTMLDivElement} 数据处理区域容器元素 */
    private sectionProcess: HTMLDivElement; 
    /** @type {CommandBus} 命令总线（外壳与宿主之间唯一的接触面） */
    private bus: CommandBus;
    /** @type {Menu} 菜单实例 */
    private menu: Menu;
    /** @type {Record<string, BtnBase | undefined>} 工具栏按钮集合 */
    private toolbars: Record<string, BtnBase | undefined> = {};
    /** 点选求和/平均/计数/最大/最小后的回调 */
    private readonly onCalculate?: (label: string) => void;
    /** 点选查询/替换后的回调 */
    private readonly onFindReplace?: (mode: 'find' | 'replace') => void;

    /**
     * 创建 SectionProcess 实例
     * @param {sectionProcessOptions} options - 构造选项
     */
    constructor(options: sectionProcessOptions) {
        this.sectionProcess = createDiv({
            className: TOOLBAR_CLASS_NAME.SECTION,
        }); 
        
        if (options.parentElement) {
            options.parentElement.append(this.sectionProcess);
        }
        this.bus = options.bus;
        this.menu = options.menu;
        this.onCalculate = options.onCalculate;
        this.onFindReplace = options.onFindReplace;
        const line = createDiv({
            className: TOOLBAR_CLASS_NAME.SECTION_LINE, 
        });

        this.sectionProcess.append(line);
        this.init();
    }

    /**
     * 初始化工具栏按钮
     * 创建计算、排序、冻结、筛选和查询等数据处理功能按钮
     * @public
     */
    public init() {
        /*------------------ 数据处理分区 ------------------*/
        const calculateBtn = new BtnBase({
            parentElement: this.sectionProcess,
            stateSource: this.bus,
            menu: this.menu,
            toDo: 'sum',
            icon: 'icon-sum',
            text: '计算',
            ishow: true,
            tshow: true,
            layout: 'vertical',
            separate: true,
            ashow: true,
            menuContent: {
                items: [
                    { todo: 'sum', icon: 'icon-sum', text: '求和'},
                    { todo: 'average', icon: 'icon-avg', text: '平均'},
                    { todo: 'count', icon: 'icon-count', text: '计数'},
                    { todo: 'max', icon: 'icon-max', text: '最大'},
                    { todo: 'min', icon: 'icon-min', text: '最小'}
                ]
            },
            onClick: (todo) => {
                if (todo.includes('MenuOpen')) {
                    this.menu.openMenu(calculateBtn.getElements().container);
                    this.menu.addMenuElement(calculateBtn.getElements().menuContainer);
                } else {
                    // 点选某一项 → 只在活动格内显示该项的统计结果（底部栏仍显示全部即时值）
                    const label = CALC_ITEM_LABELS[todo];
                    if (label) this.onCalculate?.(label);
                }
            }
        });
        
        const sortBtn = new BtnBase({
            parentElement: this.sectionProcess, 
            stateSource: this.bus,
            menu: this.menu,
            toDo: 'sort-up',
            icon: 'icon-sort-down-alt',
            text: '排序',
            ishow: true,
            tshow: true,
            layout: 'vertical',
            separate: true,
            ashow: true,
            menuContent: {
                items: [
                    { todo: 'sort-up', icon: 'icon-sort-amount-asc', text: '升序'},
                    { todo: 'sort-down', icon: 'icon-sort-amount-desc', text: '降序'}
                ]
            },
            onClick: (todo) => {
                if (todo.includes('MenuOpen')) {
                    this.menu.openMenu(sortBtn.getElements().container);
                    this.menu.addMenuElement(sortBtn.getElements().menuContainer);
                } else {
                    if (todo.includes('sort-up')) {
                        this.bus.dispatch('sortUp');
                    } else if (todo.includes('sort-down')) {
                        this.bus.dispatch('sortDown');
                    }
                }
            }
        });
        
        const frozenBtn = new BtnBase({
            parentElement: this.sectionProcess, 
            stateSource: this.bus,
            menu: this.menu,
            toDo: 'frozen',
            icon: 'icon-frozen-grid',
            text: '冻结',
            ishow: true,
            tshow: true,
            layout: 'vertical',
            separate: true,
            ashow: true,
            onClick: (todo) => {
            }
        });

        const filterBtn = new BtnBase({
            parentElement: this.sectionProcess, 
            stateSource: this.bus,
            menu: this.menu,
            toDo: 'filter',
            icon: 'icon-funnel',
            text: '筛选',
            layout: 'vertical',
            ishow: true,
            tshow: true,
            // separate: true,
            // ashow: true,
            onClick: (todo) => {
                this.bus.dispatch('toggleFilter');
            }
        });

        const searchBtn = new BtnBase({
            parentElement: this.sectionProcess, 
            stateSource: this.bus,
            menu: this.menu,
            toDo: 'search',
            icon: 'icon-search',
            text: '查询',
            ishow: true,
            tshow: true,
            layout: 'vertical',
            separate: true,
            ashow: true,
            menuContent: {
                items: [
                    { todo: 'search', icon: 'icon-search', text: '查询'},
                    { todo: 'replace', icon: 'icon-replace', text: '替换'}
                ]
            },
            onClick: (todo) => {
                if (todo.includes('MenuOpen')) {
                    this.menu.openMenu(searchBtn.getElements().container);
                    this.menu.addMenuElement(searchBtn.getElements().menuContainer);
                } else if (todo.includes('replace')) {
                    // 判断顺序：'replace' 在前 —— 'replace'.includes('search') 为 false，
                    // 但反过来写（先判 search）会让 todo='replace' 被 'search' 分支截胡
                    // （真实菜单项的 todo 就是 'replace'，不含 'search' 字样，但保持
                    // 「特殊的分支先判」这一顺序，后续加 'search-in-range' 之类的条目不会踩坑）
                    this.onFindReplace?.('replace');
                } else if (todo.includes('search')) {
                    this.onFindReplace?.('find');
                }
            }        });
        this.toolbars = {
            calculateBtn,
            sortBtn,
            filterBtn,
            searchBtn,
        };
        this.setupStateListeners();
    }

    /**
     * 按当前宿主状态刷新按钮高亮
     *
     * 编辑会话中不刷新筛选按钮：此时筛选条件不可能变，刷了只会白写 DOM。
     * @private
     */
    private setToolbarsState() {
        const state = this.bus.snapshot();
        if (state.editing) return;
        this.toolbars.filterBtn.setActived(state.filterActive);
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
     * @returns {{sectionProcess: HTMLDivElement}} 工具栏元素对象
     * @public
     */
    public getElement(): {sectionProcess: HTMLDivElement} {
        return { sectionProcess: this.sectionProcess };
    }
}
