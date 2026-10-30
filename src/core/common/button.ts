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
import { EDIT_STATE_CHANGED, EditingStateSource } from "../contract/editHost";
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
    /**
     * 当前是否正被「编辑态置灰」压着
     *
     * 解禁时只能回滚**自己**压上去的那一次：撤销 / 重做这类按钮的置灰另有原因
     * （`setToolbarsState` 按快照设的），一刀切 `setDisabled(false)` 会把它们也误放开。
     */
    private disabledByEditting: boolean = false;
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
            // 【`separate` 只是呈现策略，不是存在性保证】`textElement` 只在
            // `tshow: true` 时由 `_setText()` 创建；无条件 `addEventer(undefined)`
            // 会在**构造期**抛 TypeError，进而中断工具栏后续所有分区的装配
            // （症状是「页面显示不完全」而不是「某个按钮点了没反应」）。
            const parts: HTMLElement[] = [];
            if (this.iconElement) parts.push(this.iconElement);
            if (this.textElement) parts.push(this.textElement);
            if (parts.length === 0) this.addEventer(this.container);
            else parts.forEach((element) => this.addEventer(element));
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
                // 同 init：缺文字区时 `textElement` 是 undefined，直接 `.contains` 会抛错。
                // 没有文字区就把容器留白也当图标区，否则点在图标边缘会毫无反应。
                const inIcon = this.iconElement
                    ? this.iconElement.contains(e.target as Node)
                    : !this.textElement || !this.textElement.contains(e.target as Node);
                if (inIcon) {
                    if (this.options.ashow && this.isOpen) {
                        // 点图标 = 执行动作并顺手收起菜单（不再弹回菜单）
                        this.isOpen = false;
                        this.menu.closeMenu();
                        this.todo(`${this.options.toDo}MenuClose`);
                    }
                    this.todo(`${this.options.toDo}`);
                } else if (this.textElement && this.textElement.contains(e.target as Node)) {
                    if (this.options.ashow) {
                        this.isOpen = !this.isOpen;
                        if (!this.isOpen) this.menu.closeMenu();
                        this.todo(`${this.options.toDo}${this.isOpen ? 'MenuOpen' : 'MenuClose'}`);
                    } else {
                        // 菜单被 {@link setMenuEnabled} 运行时关闭：文字区不再是死区，直接执行动作
                        this.todo(`${this.options.toDo}`);
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

        this.bindEdittingGuard();
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
     * 运行时关闭 / 恢复下拉菜单
     *
     * 用于「同一个按钮在另一类工作区里没有菜单」的场景。收掉后**文字区域也会执行
     * 动作**（原本 `separate` 按钮的文字区只负责展开菜单，收掉菜单后它就成了死区）。
     * @param enabled - 是否保留菜单
     */
    public setMenuEnabled(enabled: boolean): void {
        if (!this.options.menuContent) return;
        this.options.ashow = enabled;
        if (!enabled) {
            this.isOpen = false;
            this.menu.closeMenu();
        }
        if (this.arrowElement) this.arrowElement.hidden = !enabled;
    }

    /**
     * 运行时切换「编辑态置灰」策略
     *
     * 表格里「编辑会话中不改段落属性」是合理的（编辑的是字符，不是单元格属性）；
     * 但文档**没有编辑会话**这个概念 —— 正文一获得焦点 `editing` 就是 true，
     * 若照搬表格策略，对齐 / 油漆桶这类按钮在文档里会**永远**是灰的。
     * 因此置灰策略必须能按工作区在运行期切换，而不是写死在构造参数里。
     * @param enabled - true 表示编辑态置灰（表格口径）
     */
    public setDisabledOnEditting(enabled: boolean): void {
        this.isDisabledOnEditting = enabled;
        // 关掉策略时要立刻解除：置灰状态不会因为策略变了而自动回溯
        if (!enabled) this.releaseEdittingGuard();
    }

    /**
     * 运行时替换下拉菜单的菜单项
     *
     * 同一个按钮在两类工作区里的语义可以完全不同（粘贴：表格是「值 / 格式」、
     * 文档是「文字 / 格式」），但按钮本身是同一个实例 —— 与其建两个按钮再互相藏，
     * 不如换菜单内容。`MenuContent` 在构造期就把 DOM 建好了，所以只能整体重建。
     * @param items - 新的菜单项
     */
    public setMenuItems(items: MenuContentOptions['items']): void {
        if (!this.options.menuContent || !this.menuContent) return;
        this.options.menuContent.items = items;
        const previous = this.menuContainer;
        this.isOpen = false;
        this.setMenuContent();
        // 老容器可能正挂在某个父节点下（Menu.openMenu 会把它搬进浮层），就地替换
        if (previous && previous.parentElement && this.menuContainer) {
            previous.parentElement.replaceChild(this.menuContainer, previous);
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
    }

    /**
     * 「编辑态置灰」是否生效
     *
     * 两条收敛规则：
     * - 按钮自己声明 `isDisabledOnEditting: false` → 不生效；
     * - **文档载体整体豁免**：正文一获得焦点 `editing` 就是 true，若照搬表格口径，
     *   打印 / 缩进这类按钮在文档里会**永远**是灰的（用户实测：鼠标一进编辑区就全灰）。
     *
     * 之所以放在这里而不是让每个分区逐个调 {@link setDisabledOnEditting}：
     * 逐个关是**补漏**式的，漏一个就是一个 bug；这里是按载体一刀切，新增按钮自动生效。
     */
    private get edittingGuardOn(): boolean {
        return this.isDisabledOnEditting && this.stateSource?.kind !== 'doc';
    }

    /**
     * 订阅置灰所需的两个事件（构造期绑定一次）
     *
     * 早先这段写在 `addEventer` 里，而 `separate` 按钮会对图标区、文字区各调一次 →
     * 同一个按钮订阅两遍，一次编辑态变化 `setDisabled` 两次（目前无副作用，但监听表翻倍）。
     * @private
     */
    private bindEdittingGuard(): void {
        const source = this.stateSource;
        // 未提供事件源时静默跳过（不再像改造前那样因漏传 `data` 在构造期抛 TypeError）
        if (!source) return;
        source.on(DataEvents.IS_EDITTING_CHANGED, (_: unknown, isEditting: boolean) => {
            if (!this.edittingGuardOn) {
                this.releaseEdittingGuard();
                return;
            }
            this.disabledByEditting = isEditting;
            this.setDisabled(isEditting);
        });
        // 载体切换（表格 ⇄ 文档）时策略会从「开」变「关」，而已经压上去的置灰不会自动回溯。
        // 总线的 `bind` 会广播一次 EDIT_STATE_CHANGED，这里借它兜底解禁。
        source.on(EDIT_STATE_CHANGED, () => {
            if (this.edittingGuardOn) return;
            this.releaseEdittingGuard();
        });
    }

    /**
     * 撤销「编辑态置灰」（只撤自己压上去的那一次）
     * @private
     */
    private releaseEdittingGuard(): void {
        if (!this.disabledByEditting) return;
        this.disabledByEditting = false;
        this.setDisabled(false);
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