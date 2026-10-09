/*  # 通用 select 组件 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { createDiv, createInput, createSpan } from "../../utils/dom";
import { EventEmitter } from '../../utils/eventEmitter';
import chevronDown from '../../assets/images/toolbar/chevron-down.svg';
import { MenuContent, MenuContentOptions } from "./menu";
import { DataCollection } from "../dataArchitecture/dataCollection";

export interface SelectOptions {
    parentElement?: HTMLElement | null;    
    data?: DataCollection;
    title?: string;
    containerID?: string;
    inputID?: string;
    toDo?: string;
    menuContent?: MenuContentOptions;  // 下拉菜单内容 ，如果存在，会显示下拉菜单
    onClick?: (todo: string) => void;
}


export class Select extends EventEmitter {
    private data: DataCollection;
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
        
        if (options.data) {
            this.data = options.data;
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