/*  # 通用 select 组件 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目描述: 基于 HTML5 Canvas 的 Web 电子表格组件，无需服务端支持即可在浏览器中创建、读取、写入 Excel 文件。
    # 项目目标: 提供一个「输入框 + 下拉箭头 + 菜单」的轻量选择器，供工具栏的字体 / 字号等下拉项复用。
    # 本文件职责: 创建与维护下拉选择器的 DOM 结构，把菜单点选结果通过 `click` 事件抛给调用方。
    #             **它不持有任何数据模型**：改造前曾声明一个 `data?: DataCollection` 字段，
    #             但从头到尾只赋值、从未读取（死依赖），本次一并删除。
    # 依赖方向: utils/dom + utils/eventEmitter + common/menu，不依赖 dataArchitecture。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { createDiv, createInput, createSpan } from "../../utils/dom";
import { EventEmitter } from '../../utils/eventEmitter';
import chevronDown from '../../assets/images/toolbar/chevron-down.svg';
import { MenuContent, MenuContentOptions } from "./menu";

/**
 * 下拉选择器配置选项
 */
export interface SelectOptions {
    /** 父元素 */
    parentElement?: HTMLElement | null;
    /** 提示文字（同时作为容器 title） */
    title?: string;
    /** 容器 id */
    containerID?: string;
    /** 输入框 id */
    inputID?: string;
    /** 操作名（点击箭头时发出 `${toDo}MenuOpen` / `${toDo}MenuClose`） */
    toDo?: string;
    /** 下拉菜单内容，如果存在会显示下拉菜单 */
    menuContent?: MenuContentOptions;
    /** 点击回调 */
    onClick?: (todo: string) => void;
}

/**
 * 下拉选择器
 *
 * 只负责「长得像一个选择器」：输入框显示当前值、箭头负责开合菜单、菜单项点击后把 `todo` 抛给调用方。
 * 选中项的**语义**（是改字体还是改字号）由调用方在 `onClick` 里决定。
 */
export class Select extends EventEmitter {
    private toDo: string;
    private container: HTMLDivElement;
    private title: string;
    private input: HTMLInputElement;
    private arrowElement: HTMLDivElement;
    private arrow: HTMLDivElement;
    private options: SelectOptions;
    private menuContent: MenuContent | null = null;  // 下拉菜单内容，如果存在，会显示下拉菜单
    private menuContainer: HTMLDivElement | null = null;  // 下拉菜单容器
    private isOpen: boolean = false;  // 是否被选中或打开 menu

    /**
     * 创建下拉选择器
     * @param options - 配置选项
     */
    constructor(options: SelectOptions) {
        super();
        this.options = options;       
        this.toDo = options.toDo;
        this.title = options.title;
        
        this.container = createDiv({
            id: `${options.containerID? options.containerID : ''}`,
            className: 'toolbar-select'
        });

        this.input = createInput({
            id: `${options.inputID}`,
            className: 'toolbar-select-input'
        });

        this.arrowElement = createDiv({
            className: 'toolbar-select-icon',
        });

        if (options.menuContent) {
            this.menuContent = new MenuContent(options.menuContent);
        }

        // 添加箭头
        const dataURI = `data:image/svg+xml;base64,${btoa(chevronDown.replace('data:image/svg+xml,', '').replace(/%3c/g, '<').replace(/%3e/g, '>'))}`;
        this.arrow = createDiv({
            dataset: {
                todo: `menu`,
            },
            style: {
                position: 'relative',
                width: '8px',
                height: '8px',
                backgroundImage: `url(${dataURI})`,
                backgroundRepeat: 'no-repeat',
                backgroundPosition: 'center center',
                backgroundSize: '8px 8px',
                transition: 'transform 0.3s ease',
            }
        });
        this.container.append(this.input, this.arrowElement);
        this.arrowElement.appendChild(this.arrow);

        if (options.parentElement) {
            options.parentElement.appendChild(this.container);
        }
        
        if (options.onClick) {
            this.on('click', options.onClick);
        }
        this.container.title = this.title;

        if (options.menuContent) {
            this.setMenuContent();
        }

        this.init();
    }

    private todo(whatToDo: string) {
        this.emit('click', whatToDo); 
    }

    private init() {
        this.arrowElement.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            if (this.arrowElement.contains(e.target as Node)) {
                this.isOpen = this.arrow.style.transform === 'rotate(180deg)' ? true : false;
                this.isOpen = !this.isOpen;
                this.todo(`${this.toDo}${this.isOpen ? 'MenuOpen' : 'MenuClose'}`);
            }
        });
    }

    public setText(todo: string) {  // 点击菜单后，用于设置文字
        const items = this.options.menuContent?.items;
        if (Array.isArray(items)) {
            const found = items.find((item): item is { todo?: string; icon?: string; text?: string } => typeof item !== 'string' && typeof item !== 'function' && item.todo === todo);
            this.input.value = found?.text || todo;
        } else {
            this.input.value = todo;
        }
        this.input.style.fontFamily = todo;
    }

    private setMenuContent() {
        this.menuContent = new MenuContent({
            items: this.options.menuContent.items,  
            onClick: (todo: string) => {
                this.todo(todo);
            }
        });
        this.menuContainer = this.menuContent.getElements().container;
        const items = Array.isArray(this.options.menuContent?.items) ? this.options.menuContent.items : [];
        for (let item of items) {
            if (item === 'separator') {
                continue;
            }
            if (typeof item !== 'string' && typeof item !== 'function') {
                this.input.value = item.text || '';
                break;
            }
        }
    }

    public getElements() {
        return { container: this.container, menuContainer: this.menuContainer };
    }
}