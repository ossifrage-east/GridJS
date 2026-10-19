/**
 * 通用 Menu 组件
 * 提供菜单容器和菜单内容的管理功能
 *
 * @module Menu
 * @author 东方鹗
 * @see
 * # B站: https://space.bilibili.com/194359739
 * # 知乎: https://www.zhihu.com/people/eastossifrage
 * # CSDN: https://blog.csdn.net/os373
 */

import { createDiv } from "../../utils/dom";
import { EventEmitter } from "../../utils/eventEmitter";
import { MENU_CLASS_NAMES } from "../constant";
import '../style/menu.scss';

/**
 * 菜单选项接口
 */
export interface MenuOptions {
    /** 父元素 */
    parentElement?: HTMLElement | null;
}

/**
 * 菜单类
 * 管理菜单容器，提供打开、关闭菜单的功能
 *
 * @class Menu
 * @extends EventEmitter
 */
export class Menu extends EventEmitter {
    /** 菜单容器 */
    private container: HTMLDivElement;
    /** 当前打开的菜单元素 */
    public hasMenu: HTMLDivElement | null = null;

    /**
     * 构造函数
     * @param {MenuOptions} options - 菜单选项
     */
    constructor(options: MenuOptions) {
        super();
        this.container = createDiv({
            className: MENU_CLASS_NAMES.CONTAINER,
            id: MENU_CLASS_NAMES.ID,
        });
        this.container.setAttribute('tabindex', '99');
        if (options.parentElement) {
            options.parentElement.appendChild(this.container);
        }

        this.container.addEventListener('click', (e) => {
            e.stopPropagation();
            return false;
        });

        document.addEventListener('click', (e) => {
            e.stopPropagation();
            this.closeMenu();
        });
    }

    /**
     * 打开菜单
     * @param {HTMLDivElement} element - 触发菜单的元素
     */
    public openMenu(element: HTMLDivElement): void {
        this.hasMenu = element;
        const arrow = element?.querySelector('[data-todo="menu"]') as HTMLDivElement | null;
        if (arrow) {
            arrow.style.transform = 'rotate(180deg)';
        }
        this.container.classList.add('show');
    }

    /**
     * 关闭菜单
     */
    public closeMenu(): void {
        const arrow = this.hasMenu?.querySelector('[data-todo="menu"]') as HTMLDivElement | null;
        if (arrow) {
            arrow.style.transform = 'rotate(0deg)';
            this.hasMenu = null;
        }
        while (this.container.firstChild) {
            this.container.removeChild(this.container.firstChild);
        }
        this.container.classList.remove('show');
        this.emit('close');
    }

    /**
     * 把菜单收敛到视口内，避免靠近屏幕右边缘 / 下边缘时被裁掉
     *
     * 分两步：
     * 1. 高度上限：内容比视口还高时收窄 `max-height` 到视口以内（超出部分靠菜单自身
     *    `overflow-y: auto` 滚动），保证上下两头都能看全；能装下时清除内联值、回到
     *    SCSS 的默认上限，避免上一次的收窄值残留影响后续较小的菜单。
     * 2. 平移：测量菜单在视口中的实际位置，右 / 下溢出时向内平移，菜单比视口还宽 /
     *    高时优先保证左 / 上边缘可见。
     *
     * 为什么用「平移增量」而不是把 left/top 直接写成视口坐标：菜单是 `position: absolute`，
     * 定位基准是定位祖先（如 `.split-pre`）而非视口，若祖先不在视口原点，直接写视口坐标会
     * 整体偏移；只调整相对当前值的增量则与祖先位置无关，任何容器下都成立。
     * @private
     */
    private fitToViewport(): void {
        if (!this.container.classList.contains('show')) {
            return;
        }
        const margin = 4;   // 菜单与屏幕边缘的最小留白
        const vw = window.innerWidth;
        const vh = window.innerHeight;

        // 1. 高度上限：先清除内联值以取回 SCSS 默认上限，再按视口收紧
        this.container.style.maxHeight = '';
        const available = vh - margin * 2;
        if (this.container.offsetHeight > available) {
            this.container.style.maxHeight = `${available}px`;
        }

        // 2. 平移：依据当前视口位置计算需要内移的增量
        const rect = this.container.getBoundingClientRect();
        const left = parseFloat(this.container.style.left) || 0;
        const top = parseFloat(this.container.style.top) || 0;
        let dx = rect.right > vw - margin ? vw - margin - rect.right : 0;
        if (rect.left + dx < margin) dx = margin - rect.left;
        let dy = rect.bottom > vh - margin ? vh - margin - rect.bottom : 0;
        if (rect.top + dy < margin) dy = margin - rect.top;
        if (dx !== 0) this.container.style.left = `${left + dx}px`;
        if (dy !== 0) this.container.style.top = `${top + dy}px`;
    }

    /**
     * 设置菜单位置
     * @private
     */
    private setMenuPosition(): void {
        if (!this.hasMenu) {
            return;
        }
        const { top, left, height } = this.hasMenu.getBoundingClientRect();
        this.container.style.top = `${top + height + 3}px`;
        this.container.style.left = `${left}px`;
        this.fitToViewport();
    }

    /**
     * 添加菜单元素
     * @param {HTMLElement} element - 要添加的元素
     */
    public addMenuElement(element: HTMLElement): void {
        this.container.appendChild(element);
        this.setMenuPosition();
    }

    /**
     * 在指定视口坐标处打开右键菜单
     *
     * 与 openMenu 不同：不依赖触发元素定位（无箭头旋转），
     * 直接以鼠标坐标作为菜单左上角位置；随后经 {@link fitToViewport} 收敛，
     * 保证靠近屏幕右 / 下边缘时菜单仍完整可见（不会溢出被裁）。
     * @param {number} clientX - 视口X坐标（contextmenu 事件的 clientX）
     * @param {number} clientY - 视口Y坐标（contextmenu 事件的 clientY）
     * @param {HTMLElement} element - 菜单内容元素
     */
    public openContextMenu(clientX: number, clientY: number, element: HTMLElement): void {
        this.hasMenu = null;
        // 先清空容器中已打开的菜单内容，避免跨区域（主表格/行头/列头）连续右键时重复追加
        while (this.container.firstChild) {
            this.container.removeChild(this.container.firstChild);
        }
        this.container.classList.add('show');
        this.container.style.top = `${clientY}px`;
        this.container.style.left = `${clientX}px`;
        this.container.appendChild(element);
        this.fitToViewport();
    }
}

/**
 * 菜单内容选项接口
 */
export interface MenuContentOptions {
    /** 菜单项数组或自定义元素生成函数 */
    items: ({ todo?: string, icon?: string | null, badge?: 'T' | 'brush', text?: string } | 'separator' | (() => HTMLElement))[] | (() => HTMLElement);
    /** 点击回调函数 */
    onClick?: (todo: string) => void;
}

/**
 * 创建带右下角角标的复合图标（用于"只粘贴文本/格式"等需要区分粘贴变体的场景）
 *
 * 在 base 图标右下角叠加一个小角标（T 字 / 格式刷），视觉上保留"粘贴"语义同时标明变体。
 * @param {string} base - 基础图标 className（如 'icon-paste'）
 * @param {'T' | 'brush'} badge - 角标类型：'T' = 字母 T；'brush' = icon-brush 小格式刷
 * @returns {HTMLElement} 复合图标元素（relative 包裹，内含 base + overlay）
 */
export function createCompositeIcon(base: string, badge: 'T' | 'brush'): HTMLElement {
    const wrap = document.createElement('span');
    wrap.style.position = 'relative';
    wrap.style.display = 'inline-flex';
    wrap.style.alignItems = 'center';

    const baseEl = document.createElement('i');
    baseEl.className = base;
    wrap.appendChild(baseEl);

    const overlay = document.createElement('span');
    overlay.style.cssText = 'position:absolute;right:-3px;bottom:-2px;background:#fff;border:1px solid #888;border-radius:2px;font-size:8px;line-height:1;padding:0 1px;color:#333;box-shadow:0 0 0 1px #fff;';
    if (badge === 'T') {
        overlay.textContent = 'T';
        overlay.style.fontWeight = 'bold';
        overlay.style.fontFamily = 'Arial, sans-serif';
    } else {
        const brush = document.createElement('i');
        brush.className = 'icon-brush';
        brush.style.fontSize = '9px';
        brush.style.lineHeight = '1';
        overlay.appendChild(brush);
        overlay.style.padding = '0';
        overlay.style.minWidth = '9px';
        overlay.style.textAlign = 'center';
    }
    wrap.appendChild(overlay);
    return wrap;
}

/**
 * 菜单内容类
 * 创建菜单内容，支持文本项、图标项、分隔符和自定义元素
 *
 * @class MenuContent
 * @extends EventEmitter
 */
export class MenuContent extends EventEmitter {
    /** 菜单内容容器 */
    private container: HTMLDivElement;
    /** 配置选项 */
    private options: MenuContentOptions;

    /**
     * 构造函数
     * @param {MenuContentOptions} options - 菜单内容选项
     */
    constructor(options: MenuContentOptions) {
        super();
        this.options = options;
        this.container = createDiv({
            className: MENU_CLASS_NAMES.MENU_CONTENT,
        });
        if (options.onClick) {
            this.on('click', options.onClick);
        }
        this.addMenuItem();
    }

    /**
     * 添加菜单项
     * @private
     */
    private addMenuItem(): void {
        (Array.isArray(this.options.items) ? this.options.items : [this.options.items]).forEach(
            (item: { todo?: string, icon?: string, badge?: 'T' | 'brush', text: string } | 'separator' | (() => HTMLElement)) => {
                const menuItem = document.createElement('div');
                if (item === 'separator') {
                    menuItem.classList.add(MENU_CLASS_NAMES.LINE);
                    this.container.appendChild(menuItem);
                    return;
                }

                menuItem.classList.add(MENU_CLASS_NAMES.ITEM);

                if (typeof item === 'function') {
                    const element = item();
                    this.container.appendChild(element);
                    return;
                }

                const text = document.createElement('span');
                text.classList.add('text');
                text.textContent = `${item.text}`;

                // badge 优先：复合图标（base icon + 右下角角标）
                if (item.badge && item.icon) {
                    menuItem.appendChild(createCompositeIcon(item.icon, item.badge));
                } else if (item.icon) {
                    const iconEl = document.createElement('i');
                    iconEl.className = `${item.icon}`;
                    menuItem.appendChild(iconEl);
                }
                menuItem.appendChild(text);
                this.container.appendChild(menuItem);

                this.handleClick(menuItem, `${item.todo}`);
            });
    }

    /**
     * 处理点击事件
     * @private
     * @param {HTMLElement} element - 目标元素
     * @param {string} todo - 操作名称
     */
    private handleClick(element: HTMLElement, todo: string): void {
        element.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            this.emit('click', todo);
        });
    }

    /**
     * 获取元素集合
     * @returns {{ container: HTMLDivElement }} 元素集合
     */
    public getElements(): { container: HTMLDivElement } {
        return {
            container: this.container,
        };
    }
}