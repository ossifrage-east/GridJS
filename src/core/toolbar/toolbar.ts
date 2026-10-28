/*  # web 表格顶部工具栏 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { TOOLBAR_CLASS_NAME, TOOLBAR_TABS } from "../constant";
import { createDiv } from '../../utils/dom';
import { EventEmitter } from '../../utils/eventEmitter';
import { DataCollection } from "../dataArchitecture/dataCollection";
import { Menu } from "../common/menu";
import { SectionBrush } from "./start/sectionBrush";
import { SectionFont } from "./start/sectionFont";
import { SectionAlign } from "./start/sectionAlign";
import { SectionProcess } from "./start/sectionProcess";
import { BtnBase } from "../common/button";
import { Select } from "../common/select";
import { needsCompact } from '../../utils/fitWidth';
import '../style/toolbar.scss';
import { SectionPrint } from "./pages/sectionPrint";


export interface ToolbarOptions {
    parentElement: HTMLElement;
    data: DataCollection;
    menu: Menu;
    onTabContentShow?: (width: number, height: number) => void;
    /**
     * 「计算」菜单（求和 / 平均 / 计数 / 最大 / 最小）点选回调
     *
     * 由宿主转发给选区即时计算模块，使其只在活动格内显示所点的那一项。
     * 工具栏构造先于即时计算模块，因此这里用回调而不是实例。
     */
    onCalculate?: (label: string) => void;
}

/** 分区描述：元素 + 菜单里显示的名字（紧凑模式下作为「一级菜单」的标签） */
export interface ToolbarSectionItem {
    el: HTMLElement;
    title: string;
}

/**
 * 快速访问区留给页签的横向空间（CSS 变量名）
 *
 * 页签那一行要保持「视觉居中」，而快速访问区是绝对定位在**最左侧**的（不参与流式排布），
 * 所以必须把它的宽度以左右 padding 的形式让出来，否则窄屏下两者会叠在一起。
 * 值由 {@link Toolbar.updateCompact} 每次量完写回。
 */
const QUICK_ACCESS_SPACE_VAR = '--qa-space';

/**
 * 顶部工具栏
 *
 * ## 三级渐进披露 + 宽度自适应
 *
 * 结构上一直是「根分类标签（开始 / 页面）→ 分区（格式刷 / 字体样式 / …）→ 分区内的按钮」。
 * 宽度不够时**一级一级往里收**，而不是硬挤：
 *
 * | 宽度 | 表现 |
 * |---|---|
 * | 足够宽 | 快速访问区钉在最左，根分类居中，所有分区并排铺开 |
 * | 分区装不下（{@link TOOLBAR_CLASS_NAME.COMPACT}） | 根分类照旧 → 一排分区小标签（一级菜单）→ 只展开点开的那个分区（二级菜单） |
 * | 连页签那一行都装不下（{@link TOOLBAR_CLASS_NAME.ROOT_COMPACT}） | 根部只露出**一个**分类（活动分类，没有活动分类时是第一个），点它展开分类清单；快速访问区同理只露品牌图标，点开才见撤销/重做/打印 |
 *
 * 判定用的是「自然宽度」（与当前模式无关），见 {@link updateCompact}。
 */
export class Toolbar extends EventEmitter {
    private container: HTMLDivElement;
    private tabs: HTMLDivElement;
    private category: HTMLDivElement;
    private content: HTMLDivElement;
    private groups: Map<string, HTMLDivElement> = new Map();
    /** 每个分组的「一级菜单」条（分组 id → nav 元素） */
    private navs: Map<string, HTMLDivElement> = new Map();
    /** 根分类的「分类清单」浮层（根折叠时展开） */
    private tabsNav: HTMLDivElement;
    /** 根分类标签，创建顺序即菜单顺序 */
    private rootTabs: HTMLDivElement[] = [];
    private data: DataCollection;
    private menu: Menu;
    /** 当前是否处于紧凑模式（分区装不下时） */
    private compact: boolean = false;
    /** 当前是否处于根折叠模式（页签那一行都装不下时） */
    private rootCompact: boolean = false;
    /** 分类清单浮层是否展开 */
    private rootOpen: boolean = false;
    /** 「计算」菜单点选回调（转给数据处理分区） */
    private readonly onCalculate?: (label: string) => void;

    constructor(options: ToolbarOptions) {
        super();
        if (options.onTabContentShow) {
            this.on('tabContentShow', options.onTabContentShow);
        }
        this.onCalculate = options.onCalculate;
        this.container = createDiv({className: TOOLBAR_CLASS_NAME.CONTAINER});
        options.parentElement.appendChild(this.container);

        this.tabs = createDiv({className: TOOLBAR_CLASS_NAME.TABS});
        this.category = createDiv({className: TOOLBAR_CLASS_NAME.CATEGORY});
        this.category.appendChild(this.tabs);
        this.content = createDiv({className: TOOLBAR_CLASS_NAME.CONTENT});
        this.container.appendChild(this.category);
        this.container.appendChild(this.content);
        this.data = options.data;
        this.menu = options.menu;
        this.init();

        // 点空白处收起分类清单（点分类标签自己由标签的点击处理负责）
        document.addEventListener('pointerdown', (e: PointerEvent) => {
            if (!this.rootOpen) return;
            const target = e.target as Node | null;
            if (target && this.category.contains(target)) return;
            this.setRootOpen(false);
        });

        this.setupResizeObserver();
    }

    public init() {
        this.setTabs();
        const sectionBrush = new SectionBrush({ // 格式刷分区
            data: this.data,
            menu: this.menu,
            title: '格式刷'
        });
        
        const sectionFont = new SectionFont({  // 字体样式分区
            data: this.data,
            menu: this.menu,
            title: '字体样式'
        });
        
        const sectionProcess = new SectionProcess({  // 数据处理分区
            data: this.data,
            menu: this.menu,
            title: '数据处理',
            onCalculate: this.onCalculate,
        });
        
        const sectionAlign = new SectionAlign({  // 对齐方式分区
            data: this.data,
            menu: this.menu,
            title: '对齐方式'
        });
        
        const sectionPrint = new SectionPrint({  // 打印设置分区
            data: this.data,
            menu: this.menu,
            title: '打印设置'
        });

        this.setTabContent(TOOLBAR_TABS.START, [
            { el: sectionBrush.getElements().sectionBrush, title: '格式刷' },
            { el: sectionFont.getElements().sectionFont, title: '字体样式' },
            { el: sectionAlign.getElements().sectionAlign, title: '对齐方式' },
            { el: sectionProcess.getElement().sectionProcess, title: '数据处理' },
        ]);
        this.setTabContent(TOOLBAR_TABS.PAGE, [
            { el: sectionPrint.getElements().sectionPrint, title: '打印设置' },
        ]);
        // 首屏就可能已经是窄屏：init 末尾按当前宽度定一次模式
        this.updateCompact();
    }

    public setTabs() { // 设置工具栏标签
        const murmur = require("murmurhash-js");
        // 分类清单浮层：根折叠时点「露出来的那个分类」展开它（同级分类的「一级菜单」）
        this.tabsNav = createDiv({ className: TOOLBAR_CLASS_NAME.TABS_NAV });
        this.category.appendChild(this.tabsNav);

        for (let tabText of Object.values(TOOLBAR_TABS)) {
            const hashString = murmur.murmur3(tabText, 'hawk@#');
            const tabId = hashString.toString(16);
            // 【必须用局部变量】不能用 this.tab —— 那是循环共用的字段，循环结束时指向
            // 最后一个页签，闭包里引用它会让所有页签的点击都落到最后一个上
            // （表现为「点哪个都变成最后一个分类」，工具区还可能是空的）。
            const tab = createDiv({className: TOOLBAR_CLASS_NAME.TAB, textContent: tabText});
            tab.dataset.targetId = tabId;
            this.tabs.appendChild(tab);
            this.rootTabs.push(tab);

            // 分类清单里的同名标签
            const chip = createDiv({
                className: TOOLBAR_CLASS_NAME.TABS_CHIP,
                textContent: tabText,
                attributes: { title: tabText },
            });
            chip.dataset.targetId = tabId;
            chip.addEventListener('click', () => this.activateTab(tab));
            this.tabsNav.appendChild(chip);

            const group = createDiv({className: TOOLBAR_CLASS_NAME.GROUP, id: tabId});
            this.content.appendChild(group);
            this.groups.set(tabId, group);

            tab.addEventListener('click', () => {
                // 根折叠时页签只剩一个：点它就是「展开其它分类」的开关
                if (this.rootCompact && tab.classList.contains('active')) {
                    this.setRootOpen(!this.rootOpen);
                    return;
                }
                this.activateTab(tab);
            });

            tab.addEventListener('dblclick', () => {
                // 双击已选中的分类 = 收起整块工具区（老行为）
                if (!tab.classList.contains('active')) return;
                tab.classList.remove('active');
                group.classList.remove('active');
                this.updateCompact();
            });
        }
    }

    /**
     * 切换到某个根分类
     *
     * 分类标签与分组是两套 DOM（靠 `data-target-id` 关联），切换时必须**成对**改：
     * 只改标签会让页签亮着、内容却还是上一个分类的。分类清单里的同名标签也在这里同步高亮。
     * @param {HTMLDivElement} tab - 目标分类标签
     */
    private activateTab(tab: HTMLDivElement): void {
        const tabId = tab.dataset.targetId;
        const group = tabId !== undefined ? this.groups.get(tabId) : undefined;
        if (!group) return;

        this.groups.forEach((item) => item.classList.remove('active'));
        this.rootTabs.forEach((item) => {
            item.classList.remove('active');
            item.classList.remove(TOOLBAR_CLASS_NAME.TAB_ROOT);
        });

        group.classList.add('active');
        tab.classList.add('active');
        tab.classList.add(TOOLBAR_CLASS_NAME.TAB_ROOT);

        Array.prototype.slice.call(this.tabsNav.children).forEach((chip: HTMLElement) => {
            chip.classList.toggle('active', chip.dataset.targetId === tabId);
        });

        this.setRootOpen(false);
        // 换了一批分区，可用宽度没变但内容宽度变了 —— 重新量一次
        this.updateCompact();
    }

    /**
     * 展开 / 收起「分类清单」浮层（根折叠时）
     *
     * 根折叠后页签行只剩一个分类，其它分类只能从这个浮层里选；不折叠时浮层是隐藏的，
     * 所以这里顺手把「非折叠态却还开着」的情况也纠回来（避免留下看不见的展开态）。
     * @param {boolean} open - 是否展开
     */
    private setRootOpen(open: boolean): void {
        const next = open && this.rootCompact;
        this.rootOpen = next;
        this.container.classList.toggle(TOOLBAR_CLASS_NAME.ROOT_OPEN, next);
    }

    /**
     * 让「露出来的那个分类」与实际状态一致
     *
     * 根折叠时 CSS 只显示带 {@link TOOLBAR_CLASS_NAME.TAB_ROOT} 的那个页签：
     * 有活动分类就露活动分类，一个都没选中（首屏就是这样）就露第一个 —— 否则窄屏下
     * 页签行会整片空白，用户连「还有哪些分类」都点不出来。
     * @private
     */
    private syncRootTab(): void {
        const active = this.rootTabs.filter((item) => item.classList.contains('active'))[0];
        const shown = active ?? this.rootTabs[0];
        this.rootTabs.forEach((item) => {
            item.classList.toggle(TOOLBAR_CLASS_NAME.TAB_ROOT, item === shown);
        });
    }

    /**
     * 往某个根分类里填分区，并生成该分组的「一级菜单」条
     *
     * 紧凑模式下只显示 {@link TOOLBAR_CLASS_NAME.SECTION_ACTIVE} 的那个分区，
     * 所以这里给每个分区打上标题与下标，并默认展开第一个。
     * @param {string} tabText - 根分类名（开始 / 页面）
     * @param {ToolbarSectionItem[]} sections - 分区列表（顺序即菜单顺序）
     */
    public setTabContent(tabText: string, sections: ToolbarSectionItem[]) {
        const murmur = require("murmurhash-js");
        const hashString = murmur.murmur3(tabText, 'hawk@#');
        const group = this.groups.get(hashString.toString(16));
        if (!group) return;

        // 只有一个分区时没有「切换」可言，不生成菜单条（如「页面」页签只有一个打印设置）
        const nav = sections.length > 1
            ? createDiv({ className: TOOLBAR_CLASS_NAME.SECTION_NAV })
            : null;
        sections.forEach((section, index) => {
            // 标题同时挂到 data-* 上：方便调试与测试断言，也让 CSS/选择器有据可依
            section.el.dataset.sectionTitle = section.title;
            section.el.dataset.sectionIndex = String(index);
            if (nav) {
                const chip = createDiv({
                    className: TOOLBAR_CLASS_NAME.SECTION_CHIP,
                    textContent: section.title,
                    attributes: { title: section.title },
                });
                chip.addEventListener('click', () => this.selectSection(group, index));
                nav.appendChild(chip);
            }
            group.appendChild(section.el);
        });
        if (nav) {
            group.insertBefore(nav, group.firstChild);
            this.navs.set(group.id, nav);
        }
        // 即使只有一个分区也要打上 section-active：紧凑模式下靠它决定显示谁
        this.selectSection(group, 0);
    }

    /**
     * 展开某个分组里的第 index 个分区（一级菜单的点击处理）
     * @param {HTMLDivElement} group - 分组元素
     * @param {number} index - 分区下标
     */
    private selectSection(group: HTMLDivElement, index: number) {
        const sections = Array.prototype.slice.call(
            group.querySelectorAll('.' + TOOLBAR_CLASS_NAME.SECTION)) as HTMLElement[];
        sections.forEach((el, i) => {
            el.classList.toggle(TOOLBAR_CLASS_NAME.SECTION_ACTIVE, i === index);
        });
        const nav = this.navs.get(group.id);
        if (nav) {
            Array.prototype.slice.call(nav.children).forEach((chip: HTMLElement, i: number) => {
                chip.classList.toggle('active', i === index);
            });
        }
        group.dataset.activeSection = String(index);
    }

    /**
     * 取当前显示的根分类分组
     * @returns {HTMLDivElement | null} 分组元素；没有选中任何分类时为 null
     */
    private activeGroup(): HTMLDivElement | null {
        let found: HTMLDivElement | null = null;
        this.groups.forEach((group) => {
            if (found === null && group.classList.contains('active')) found = group;
        });
        return found;
    }

    /**
     * 量页签行的**自然宽度**（一行摊开、不含快速访问区让出的 padding）
     *
     * 必须套 {@link TOOLBAR_CLASS_NAME.MEASURING}：根折叠时其它页签是 `display: none`，
     * 直接量 `scrollWidth` 只会得到「只剩一个页签」的宽度，判定就会一直认为装得下。
     * @returns {number} 页签自然宽度（px）
     * @private
     */
    private measureTabsNeed(): number {
        this.tabs.classList.add(TOOLBAR_CLASS_NAME.MEASURING);
        const need = this.tabs.scrollWidth;
        this.tabs.classList.remove(TOOLBAR_CLASS_NAME.MEASURING);
        return need;
    }

    /**
     * 量快速访问区的实际宽度
     *
     * 它是绝对定位在**最左侧**的（{@link QUICK_ACCESS_SPACE_VAR} 用它给页签留位），
     * 所以宽度要现量：宽屏时是「品牌 + 四个按钮」，根折叠后只剩品牌图标。
     * @returns {number} 宽度（px）；元素还没创建时 0
     * @private
     */
    private measureQuickAccess(): number {
        const qa = this.tabs.querySelector('.' + TOOLBAR_CLASS_NAME.QUICK_ACCESS) as HTMLElement | null;
        if (qa === null) return 0;
        if (typeof qa.getBoundingClientRect === 'function') {
            const width = qa.getBoundingClientRect().width;
            if (width > 0) return width;
        }
        return qa.offsetWidth || 0;
    }

    /**
     * 按当前可用宽度决定「铺开」还是「紧凑」/「根折叠」
     *
     * ## 为什么要先摘掉折叠类再量
     *
     * 折叠状态下页签与分区都被 `display: none` 收走了，直接量 `scrollWidth` 会得到
     * 「只剩标签」的宽度 —— 那个值永远装得下，模式就再也退不回去了（抖动/卡死）。
     * 这里改为：**先摘掉折叠类**，临时套上 {@link TOOLBAR_CLASS_NAME.MEASURING}
     * （`position: absolute; width: max-content`），量出来的就是与模式无关的**自然宽度**，
     * 判完再把模式设回去。整个过程在同一帧内完成，不会闪。
     *
     * 两级判定互相独立：分区装不下 → {@link TOOLBAR_CLASS_NAME.COMPACT}；
     * 连页签行（含给快速访问区让出的两段空间）都装不下 → {@link TOOLBAR_CLASS_NAME.ROOT_COMPACT}。
     * @private
     */
    private updateCompact(): void {
        const available = this.container.clientWidth;

        // ── ① 根部分类 ──
        // 先按「全部展开」量：右移的紧凑态不会影响页签宽度，但会影响快速访问区宽度
        this.container.classList.remove(TOOLBAR_CLASS_NAME.ROOT_COMPACT);
        const quickWidth = this.measureQuickAccess();
        const tabsNeed = this.measureTabsNeed();
        // 快速访问区钉在左侧，页签要居中就得左右各让出同样宽的空间（对称才视觉居中）
        const rootCompact = needsCompact(tabsNeed + quickWidth * 2, available);
        this.rootCompact = rootCompact;
        this.container.classList.toggle(TOOLBAR_CLASS_NAME.ROOT_COMPACT, rootCompact);
        if (!rootCompact) this.setRootOpen(false);
        this.syncRootTab();
        // 折叠后快速访问区只剩品牌图标，让出的空间随之变小（宽屏态下写的就是全宽）
        this.container.style.setProperty(QUICK_ACCESS_SPACE_VAR, `${this.measureQuickAccess()}px`);

        // ── ② 分区（一级菜单 / 二级菜单）──
        const group = this.activeGroup();
        if (!group) return;

        this.container.classList.remove(TOOLBAR_CLASS_NAME.COMPACT);
        group.classList.add(TOOLBAR_CLASS_NAME.MEASURING);
        const need = group.scrollWidth;
        group.classList.remove(TOOLBAR_CLASS_NAME.MEASURING);

        const compact = needsCompact(need, available);
        this.compact = compact;
        this.container.classList.toggle(TOOLBAR_CLASS_NAME.COMPACT, compact);
    }

    /**
     * 按当前宽度重排一次（外部改了工具栏内部结构后调用）
     *
     * 典型场景：快速访问区是在 `Toolbar` 构造**之后**才挂进页签行的（见 `App#initToolbar`），
     * 那一刻页签行多了一截宽度，但内容区尺寸没变、`ResizeObserver` 不会触发，
     * 不主动重排就会漏掉「加上快速访问区之后其实已经装不下」这种情况。
     */
    public refreshLayout(): void {
        this.updateCompact();
    }

    /**
     * 当前是否处于紧凑模式（分区收成一级菜单）
     * @returns {boolean} 紧凑模式为 true
     */
    public isCompact(): boolean {
        return this.compact;
    }

    /**
     * 当前是否处于根折叠模式（页签行只露一个分类）
     * @returns {boolean} 根折叠为 true
     */
    public isRootCompact(): boolean {
        return this.rootCompact;
    }

    private setupResizeObserver() {
        // 设置交集观察者
        const resize = (entries: ResizeObserverEntry[]) => {
            this.handleResize(entries);
        };
            
        const resizeObserver = new ResizeObserver(resize);
        resizeObserver.observe(this.content);
    }
    
    private handleResize(entries: ResizeObserverEntry[]) {
        // 处理调整大小事件
        try {
            /* requestAnimationFrame 来异步执行处理尺寸变化的逻辑。
             * requestAnimationFrame 会告诉浏览器你希望执行一个动画，
             * 并请求浏览器在下一次重绘之前调用指定的回调函数来更新动画。
             * 这通常用于创建平滑的动画效果，但也可以用于任何需要在下一个渲染周期中执行的任务，以避免阻塞。
             */
            requestAnimationFrame(() => {
                for (const entry of entries) {
                    // 只在值实际变化时更新
                    if (entry.target === this.content) {
                        // 先定模式：紧凑模式会改变容器高度，随后通知外面的高度才是最终值
                        this.updateCompact();
                        const rect = (entry.target as HTMLElement).parentElement.getBoundingClientRect();
                        this.emit('tabContentShow', rect.width, rect.height);
                    }
                }
            });
        } catch (error) {
            console.error('ResizeObserver error:', error);
        }
    }
    
    public getElements() {
        return {
            container: this.container,
            /** 最后一个创建的分类标签（历史遗留；判断「当前分类」请看 rootTabs 里的 active） */
            tab: this.rootTabs.length > 0 ? this.rootTabs[this.rootTabs.length - 1] : null,
            tabs: this.tabs,
            tabsNav: this.tabsNav,
            content: this.content,
            groups: this.groups,
            navs: this.navs,
        }
    }
}
