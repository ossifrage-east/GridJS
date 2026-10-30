/**
 * 通用 Button 组件
 * 提供按钮的创建、图标、文字、下拉菜单等功能
 *
 * ## 依赖边界
 *
 * 按钮是**通用 UI 外壳**，它唯一需要的外部能力是「订阅编辑态变化」——
 * 用来实现 `isDisabledOnEditting`（编辑时置灰）。历史实现直接持有 `DataCollection`，
 * 于是一个与表格无关的组件被迫依赖整个表格数据层；现在退化为
 * {@link EditingStateSource}（只有 `on` / `off` 两个方法），
 * `DataCollection`、`CommandBus` 都能直接传入。
 *
 * @module Button
 * @author 东方鹗
 * @see
 * # B站: https://space.bilibili.com/194359739
 * # 知乎: https://www.zhihu.com/people/eastossifrage
 * # CSDN: https://blog.csdn.net/os373
 */

import { createDiv, createSpan } from "../../utils/dom";
import { EventEmitter } from '../../utils/eventEmitter';
import chevronDown from '../../assets/images/toolbar/chevron-down.svg';
import { Menu, MenuContent, MenuContentOptions } from "../common/menu";
import { EditingStateSource } from "../contract/editHost";
import { DataEvents } from "../constant";

/**
 * 按钮配置选项接口
 */
export interface BtnOptions {
    /** 父元素 */
    parentElement?: HTMLElement | null;
    /**
     * 编辑态事件源（用来实现「编辑时置灰」）
     *
     * 只需要能订阅 / 退订 `IS_EDITTING_CHANGED`，不关心它是表格还是文档。
     */
    stateSource?: EditingStateSource;
    /** 是否在编辑状态下禁用按钮 */
    isDisabledOnEditting?: boolean;
    /** 菜单实例 */
    menu?: Menu;
    /** 点击时触发的操作名称 */
    toDo?: string;
    /** 图标类名 */
    icon?: string | null;
    /** 按钮文字 */
    text?: string | null;
    /** 按钮标题（如果存在text，这里会显示text，否则会显示title） */
    title?: string | null;
    /** 是否显示图标 */
    ishow?: boolean;
    /** 是否显示文字 */
    tshow?: boolean;
    /** 布局方向 */
    layout?: 'horizontal' | 'vertical';
    /** 是否分离图标和文字 */
    separate?: boolean;
    /** 是否显示箭头 */
    ashow?: boolean;
    /** 下拉菜单内容 */
    menuContent?: MenuContentOptions;
    /** 是否显示颜色选择内容 */
    isColor?: boolean;
    /** 点击回调函数 */
    onClick?: (todo: string) => void;
    /** 双击回调函数 */
    onDoubleClick?: (todo: string) => void;
}

/**
 * 按钮基础类
 * 提供按钮的核心功能，包括图标、文字、下拉菜单、事件处理等
 *
 * @class BtnBase
 * @extends EventEmitter
 */
export class BtnBase extends EventEmitter {
    /** 编辑态事件源（用于「编辑时置灰」；未提供时该特性静默关闭） */
    private stateSource: EditingStateSource | undefined;
    /** 是否在编辑状态下禁用按钮 */
    private isDisabledOnEditting: boolean = true;
    /** 菜单实例 */
    private menu: Menu;
    /** 图标元素 */
    private iconElement: HTMLElement;
    /** 文字元素 */
    private textElement: HTMLDivElement;
    /** 箭头元素 */
    private arrowElement: HTMLDivElement;
    /** 按钮容器 */
    private container: HTMLDivElement;
    /** 配置选项 */
    private options: BtnOptions;
    /** 菜单内容实例 */
    private menuContent?: MenuContent;
    /** 下拉菜单容器 */
    private menuContainer: HTMLDivElement;
    /** 是否打开菜单 */
    private isOpen: boolean = false;
    /** 是否被选中 */
    public isActived: boolean = false;

    /**
     * 构造函数
     * @param {BtnOptions} options - 按钮配置选项
     */
    constructor(options: BtnOptions) {
        super();
        this.options = options;
        this.container = createDiv({
            style: {
                display: 'flex',
                flexDirection: 'row',
                padding: `${this.options.separate ? '' : '2px'}`,
                borderRadius: '4px',
                background: 'transparent',
                userSelect: 'none',
                position: 'relative',
                textTransform: 'none',
                wordWrap: 'normal',
            },
        });

        if (options.parentElement) {
            options.parentElement.appendChild(this.container);
        }

        this.container.title = this.options.text ? this.options.text : this.options.title;

        if (options.onClick) {
            this.on('click', options.onClick);
        }
        if (options.onDoubleClick) {
            this.on('dblclick', options.onDoubleClick);
        }
        if (options.stateSource !== undefined) {
            this.stateSource = options.stateSource;
        }
        if (options.menu !== undefined) {
            this.menu = options.menu;
        }
        if (options.isDisabledOnEditting !== undefined) {
            this.isDisabledOnEditting = options.isDisabledOnEditting;
        }
        this.init();
    }

    /**
     * 触发点击事件
     * @param {string} whatToDo - 操作名称
     */
    private todo(whatToDo: string) {
        this.emit('click', whatToDo);
    }

    /**
     * 初始化按钮
     */
    private init() {
        if (this.options.ishow) {
            this._setIcon();
        }
        if (this.options.tshow) {
            this._setText();
        }
        if (this.options.ashow) {
            this._setArrow();
        }

        if (this.options.tshow && this.options.ashow) {
            this.textElement.appendChild(this.arrowElement);
        } else if (!this.options.tshow && this.options.ishow && this.options.ashow) {
            this.iconElement.appendChild(this.arrowElement);
        }

        if (this.options.layout === 'horizontal') {
            this.container.style.display = 'flex';
            this.container.style.flexDirection = 'row';
        } else {
            this.container.style.display = 'block';
        }

        if (!this.options.separate) {
            this.addEventer(this.container);
        } else {
            this.addEventer(this.iconElement);
            this.addEventer(this.textElement);
        }

        if (this.options.menuContent) {
            this.setMenuContent();
            // 菜单可能被**别人**关掉（全局 document 点击、右键菜单、其它按钮抢开），
            // 此时箭头已被 Menu.closeMenu 复位，但本按钮的 `isOpen` 不会自动跟着变 ——
            // 状态一旦分叉，下次点击按钮就会「翻转成打开」却发出 MenuClose。
            // 订阅 close 事件把 isOpen 拉回真实状态（closeMenu 每次都 emit('close')）。
            this.menu.on('close', () => { this.isOpen = false; });
        }

        this.container.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            // 【勿无条件调 closeMenu()】：它会把**本按钮即将打开**的菜单一起拆掉
            // （清空容器 + 移除 show + emit('close')），而 emit('close') 又会让下面
            // 订阅的回调把 `isOpen` 归零 → 每次都读成「未打开」，于是永远只发 MenuOpen、
            // 第二次点击也关不上菜单。这里只关**别人**的菜单。
            if (this.menu.hasMenu && this.menu.hasMenu !== this.container) {
                this.menu.closeMenu();
            }
            if (!this.options.separate) {
                if (this.options.ashow) {
                    this.isOpen = !this.isOpen;
                    if (!this.isOpen) {
                        // 收起：真正关闭菜单（清内容 + 复位箭头 + 广播 close）
                        this.menu.closeMenu();
                    }
                    // 与 separate 分支保持一致：发出 `${toDo}MenuOpen` / `${toDo}MenuClose`（带前缀），
                    // 旧代码以 includes('MenuOpen') 判断不受影响，严格匹配前缀的新代码才能正确收到
                    this.todo(`${this.options.toDo}${this.isOpen ? 'MenuOpen' : 'MenuClose'}`);
                } else {
                    this.todo(`${this.options.toDo}`);
                }
            } else {
                if (this.iconElement.contains(e.target as Node)) {
                    if (this.options.ashow && this.isOpen) {
                        // 点图标 = 执行动作并顺手收起菜单（不再弹回菜单）
                        this.isOpen = false;
                        this.menu.closeMenu();
                        this.todo(`${this.options.toDo}MenuClose`);
                    }
                    this.todo(`${this.options.toDo}`);
                } else if (this.textElement.contains(e.target as Node)) {
                    if (this.options.ashow) {
                        this.isOpen = !this.isOpen;
                        if (!this.isOpen) this.menu.closeMenu();
                        this.todo(`${this.options.toDo}${this.isOpen ? 'MenuOpen' : 'MenuClose'}`);
                    }
                }
            }
        });
        this.container.addEventListener('dblclick', (e) => {
            e.preventDefault();
            e.stopPropagation();
            this.menu.closeMenu();
            this.todo(`${this.options.toDo}`);
            this.emit('dblclick', `${this.options.toDo}`);
        });
}

    /**
     * 设置图标
     */
    private _setIcon() {
        let iconBorderRadius;
        if (this.options.layout === 'vertical') {
            iconBorderRadius = '4px 4px 0 0';
        } else if (this.options.layout === 'horizontal') {
            iconBorderRadius = '4px 0 0 4px';
        }
        this.iconElement = createDiv({
            style: {
                padding: `${this.options.layout === 'vertical' ? '0' : '2px 4px'}`,
                display: 'flex',
                flexDirection: 'column',
                alignItems: 'center',
                justifyContent: 'center',
                fontSize: '15px',
                borderRadius: `${iconBorderRadius}`,
            },
        });
        const icon = document.createElement('i');
        icon.className = this.options.icon;
        this.iconElement.appendChild(icon);
        if (this.options.isColor) {
            const colorSpan = createSpan({
                style: {
                    width: '15px',
                    height: '6px',
                    borderRadius: '2px',
                    background: this.options.toDo === 'textColor' ? '' : 'linear-gradient(60deg, transparent 48%, #d32f2f 48%, #d32f2f 52%, transparent 52%)',
                    backgroundColor: this.options.toDo === 'textColor' ? '#000000' : 'transparent',
                    border: '1px solid #ced4da',
                },
            });
            this.iconElement.appendChild(colorSpan);
        }
        this.container.appendChild(this.iconElement);
    }

    /**
     * 设置图标类名
     * @param {string} icon - 图标类名
     */
    public setIcon(icon: string) {
        const iconElement = this.iconElement.querySelector('i') as HTMLDivElement | null;
        if (iconElement) {
            iconElement.className = '';
            iconElement.classList.add(`icon-${icon}`, 'font15');
        }
    }

    /**
     * 获取图标类名
     * @returns {string | undefined} 图标类名
     */
    public getIcon(): string | undefined {
        const iconElement = this.iconElement.querySelector('i') as HTMLDivElement | null;
        return iconElement?.className;
    }

    /**
     * 设置颜色指示器的颜色
     * @param {string} color - 颜色值
     */
    public setSpanColor(color: string) {
        const colorSpan = this.iconElement.querySelector('span') as HTMLDivElement | null;
        if (color === 'none') {
            if (this.options.toDo === 'textColor') {
                colorSpan.style.background = '';
                colorSpan.style.backgroundColor = `black`;
            } else {
                colorSpan.style.backgroundColor = `transparent`;
                colorSpan.style.background = 'linear-gradient(60deg, transparent 48%, #d32f2f 48%, #d32f2f 52%, transparent 52%)';
            }
        } else {
            colorSpan.style.background = ``;
            colorSpan.style.backgroundColor = `${color}`;
        }
    }

    /**
     * 获取颜色指示器的颜色
     * @returns {string} 背景颜色值
     */
    public getSpanColor(): string {
        const colorSpan = this.iconElement.querySelector('span') as HTMLDivElement | null;
        return colorSpan.style.backgroundColor;
    }

    /**
     * 设置文字
     */
    private _setText() {
        let textBorderRadius;
        if (this.options.layout === 'vertical') {
            textBorderRadius = '0 0 4px 4px';
        } else if (this.options.layout === 'horizontal') {
            textBorderRadius = '0 4px 4px 0';
        }
        this.textElement = createDiv({
            style: {
                padding: `${this.options.layout === 'vertical' ? '2px' : '4.5px 1px'}`,
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                fontSize: '14px',
                borderRadius: `${textBorderRadius}`,
            },
        });
        if (this.options.text) {
            const span = createSpan({
                style: { padding: '2px' },
                textContent: this.options.text,
            });
            this.textElement.appendChild(span);
        }
        this.container.appendChild(this.textElement);
    }

    /**
     * 设置文字颜色
     * @param {string} color - 颜色值
     */
    public setTextColor(color: string) {
        this.textElement.style.color = `${color}`;
    }

    /**
     * 设置按钮禁用状态
     * @param {boolean} isDisabled - 是否禁用
     */
    public setDisabled(isDisabled: boolean) {
        if (isDisabled) {
            this.container.setAttribute('disabled', 'true');
            this.container.style.pointerEvents = 'none';
            this.container.style.backgroundColor = 'rgba(0, 0, 0, 0.1)';
        } else {
            this.container.removeAttribute('disabled');
            this.container.style.pointerEvents = '';
            this.container.style.backgroundColor = '';
        }
    }

    /**
     * 设置箭头
     */
    private _setArrow() {
        const dataURI = `data:image/svg+xml;base64,${btoa(chevronDown.replace('data:image/svg+xml,', '').replace(/%3c/g, '<').replace(/%3e/g, '>'))}`;
        this.arrowElement = createDiv({
            dataset: {
                todo: `menu`,
            },
            style: {
                position: 'relative',
                margin: '1px',
                width: '8px',
                height: '8px',
                backgroundImage: `url(${dataURI})`,
                backgroundRepeat: 'no-repeat',
                backgroundPosition: 'center center',
                backgroundSize: '8px 8px',
                transition: 'transform 0.3s ease',
            },
        });
        this.container.appendChild(this.arrowElement);
    }

    /**
     * 设置菜单内容
     */
    private setMenuContent() {
        this.menuContent = new MenuContent({
            items: this.options.menuContent.items,
            onClick: (todo: string) => {
                this.todo(todo);
                this.menu.closeMenu();
            },
        });
        this.menuContainer = this.menuContent.getElements().container;
    }

    /**
     * 为元素添加鼠标事件
     * @param {HTMLElement} element - 目标元素
     */
    private addEventer(element: HTMLElement) {
        element.addEventListener('mouseover', (e) => {
            e.stopPropagation();
            element.style.background = 'rgba(0, 0, 0, 0.1)';
            this.container.style.outline = '1px solid rgba(0, 0, 0, 0.1)';
        });
        element.addEventListener('mouseout', (e) => {
            e.stopPropagation();
            if (this.isActived) return;
            element.style.background = 'transparent';
            this.container.style.outline = 'none';
        });
        element.addEventListener('mousedown', (e) => {
            e.stopPropagation();
            element.style.background = 'rgba(0, 0, 0, 0.2)';
        });
        element.addEventListener('mouseup', (e) => {
            e.stopPropagation();
            element.style.background = 'rgba(0, 0, 0, 0.1)';
        });
        // 编辑态变化 → 按需置灰。未提供事件源时静默跳过（不再像改造前那样
        // 因为漏传 `data` 而在构造期抛 TypeError）
        this.stateSource?.on(DataEvents.IS_EDITTING_CHANGED, (_, isEditting: boolean) => {
            if (this.isDisabledOnEditting) {
                this.setDisabled(isEditting);
            }
        });
    }

    /**
     * 设置按钮选中状态
     * @param {boolean} isActived - 是否选中
     */
    public setActived(isActived: boolean) {
        if (isActived) {
            this.container.style.background = 'rgba(0, 0, 0, 0.1)';
            this.container.style.outline = '1px solid rgba(0, 0, 0, 0.1)';
        } else {
            this.container.style.background = 'transparent';
            this.container.style.outline = 'none';
        }
        this.isActived = isActived;
    }

    /**
     * 获取按钮元素集合
     * @returns {{ container: HTMLDivElement, menuContainer: HTMLDivElement }} 元素集合
     */
    public getElements(): { container: HTMLDivElement, menuContainer: HTMLDivElement } {
        return { container: this.container, menuContainer: this.menuContainer };
    }
}