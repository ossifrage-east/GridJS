/*  # 全选区域 组件 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import crossEmpty from '../../assets/images/cursor/cross-empty.svg'
import { EventEmitter } from "../../utils/eventEmitter";
import { CELL_PADDING, DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE, GRID_LINE_COLOR, KEYWORDS, MouseLocation, SHEETS_NAME, TEXT_COLOR } from "../constant";
import { DataCollection } from "../dataArchitecture/dataCollection";
import { createDiv } from '../../utils/dom';

/**
 * 全选区域组件 初始化选项接口
 * @interface AllSelectOptions
 */
export interface AllSelectOptions {
    /** 父容器元素 */
    parentElement: HTMLElement;
    /** 数据集合实例 */
    data: DataCollection;
}

export class AllSelect extends EventEmitter {
    /** 父容器元素 */
    private parentElement: HTMLElement;
    /** 数据集合实例 */
    private data: DataCollection;
    /** 全选选择框元素 */
    private allSelect: HTMLDivElement;
    /**
     * 构造函数
     * @constructor
     * @param {AllSelectOptions} options - 初始化选项
     */
    constructor(options: AllSelectOptions) {
        super();        
        this.parentElement = options.parentElement;
        this.data = options.data;
        this.init();
    }

    /**
     * 初始化全选选择框元素
     */
    private init(){
        this.allSelect = createDiv({
            id: SHEETS_NAME.ALL_SELECT,
            style: {
                top: '0px',
                left: '0px',
                width: `${this.data.visibleView.rowHeaderWidth}px`,
                height: `${this.data.visibleView.colHeaderHeight}px`,
                cursor: `url("${crossEmpty}") 8 8, auto`
            }
        });
        this.parentElement.appendChild(this.allSelect);
    }

    /**
     * 处理鼠标按下事件
     * @param {MouseEvent} e - 鼠标事件对象
     */
    public handleMouseDown(e: MouseEvent) {
        if (e.button !== 0 || e.target !== this.allSelect) return;
        this.data.isMouseDown = true;
        this.data.startLocation = MouseLocation.IN_All_SELECT;
        this.data.activedCell = `A1`;
        this.data.selection = `A1:${this.data.getColName(this.data.colHeaders.length)}${this.data.rowHeaders.length}`;
    }

    /**
     * 处理鼠标移动事件
     * @param {MouseEvent} e - 鼠标事件对象
     */
    public handleMouseMove(e: MouseEvent) {
        return;
    }

    /**
     * 处理鼠标释放事件
     * @param {MouseEvent} e - 鼠标事件对象
     */
    public handleMouseUp(e: MouseEvent) {
        if (e.button !== 0 || e.target !== this.allSelect) return;
        this.data.isMouseDown = false;
    }

    /**
     * 获取全选选择框元素
     * @returns {HTMLDivElement} - 全选选择框元素
     */
    public getElements(): { allSelect: HTMLDivElement } {
        return {
            allSelect: this.allSelect
        }
    }
}