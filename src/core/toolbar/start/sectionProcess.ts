/*  # web 表格顶部工具栏数据处理工具集合 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/


import { createDiv } from "../../../utils/dom";
import { DataEvents, TOOLBAR_CLASS_NAME } from "../../constant";
import { BtnBase } from "../../common/button";
import { Menu } from "../../common/menu";
import { DataCollection } from "../../dataArchitecture/dataCollection";

/**
 * SectionProcess 构造选项接口
 * @interface
 * @property {HTMLElement} [parentElement] - 父元素容器
 * @property {DataCollection} data - 数据集合实例
 * @property {Menu} menu - 菜单实例
 * @property {string} [title] - 工具栏标题
 */
export interface sectionProcessOptions {
    parentElement?: HTMLElement;
    data: DataCollection;
    menu: Menu;
    title?: string;
}

/**
 * 工具栏数据处理区域类
 * 提供计算、排序、冻结、筛选和查询等数据处理功能按钮
 * @class
 */
export class SectionProcess {
    /** @type {HTMLDivElement} 数据处理区域容器元素 */
    private sectionProcess: HTMLDivElement; 
    /** @type {DataCollection} 数据集合实例 */
    private data: DataCollection;
    /** @type {Menu} 菜单实例 */
    private menu: Menu;
    /** @type {Record<string, BtnBase | undefined>} 工具栏按钮集合 */
    private toolbars: Record<string, BtnBase | undefined> = {};

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
        this.data = options.data;
        this.menu = options.menu;
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
            data: this.data,
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
                    
                }
            }
        });
        
        const sortBtn = new BtnBase({
            parentElement: this.sectionProcess, 
            data: this.data,
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
                        this.data.sortUp();
                    } else if (todo.includes('sort-down')) {
                        this.data.sortDown();
                    }
                }
            }
        });
        
        const frozenBtn = new BtnBase({
            parentElement: this.sectionProcess, 
            data: this.data,
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
            data: this.data,
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
                this.data.setFilter();
            }
        });

        const searchBtn = new BtnBase({
            parentElement: this.sectionProcess, 
            data: this.data,
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
                } else {
                    
                }
            }
        });
        this.toolbars = {
            calculateBtn,
            sortBtn,
            filterBtn,
            searchBtn,
        };
        this.setupEventListeners();
    }
    
    
    /**
     * 设置工具栏状态
     * @private
     */
    private setToolbarsState() {
        if (!this.data.isEditting) {
            this.toolbars.filterBtn.setActived(this.data.validataFilter());
        }
        
    }

    private setupEventListeners() {
        this.data.syncValuesFromDB().then(() => {
            this.setToolbarsState();
        });
        this.data.on(DataEvents.VALUES_CHANGED, () => this.setToolbarsState());
        // 撤销/重做回写后同步筛选按钮状态（保持与其它分区一致的刷新契约）
        this.data.on(DataEvents.SNAPSHOT_RESTORED, () => this.setToolbarsState());
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