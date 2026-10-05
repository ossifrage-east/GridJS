import { NAV } from '../constant';
import { IndexDB } from '../dataArchitecture/indexDB';
import '../style/nav.scss';

/**
 * Navigator 导航组件类
 * 提供面板导航功能，支持多项导航项的切换显示
 * 使用 IndexDB 持久化存储当前选中状态
 * 
 * @class Navigator
 * @example
 * const navigator = new Navigator();
 * navigator.navItemMount('文件', '脚本', '设置');
 * navigator.navContentItemMount('文件', filePanelElement);
 * const { panel, content } = navigator.getElements();
 * document.body.appendChild(panel);
 * document.body.appendChild(content);
 */
export class Navigator {
    /** 导航面板容器 */
    private panel: HTMLDivElement;
    /** 内容区域容器 */
    private content: HTMLDivElement;
    /** 导航项文本与内容ID的映射 */
    private contentItemID: Map<string, string>;
    /** IndexDB 实例，用于持久化存储 */
    private db: IndexDB;
    /** 面板唯一标识 */
    private panelId: string;
    private horizontal: boolean;

    /**
     * 构造函数
     * 初始化导航面板和内容区域元素
     */
    constructor(horizontal: boolean = true) {
        this.panel = document.createElement('div');
        this.content = document.createElement('div');
        this.contentItemID = new Map();
        this.db = new IndexDB();
        this.horizontal = horizontal;
        
        this.panel.className = NAV.PANEL;        
        this.panelId = this.generateHashId('NAV_PANEL');
        this.panel.id = this.panelId;
        this.content.className = NAV.CONTENT;
    }

    /**
     * 根据字符串生成唯一哈希ID
     * @param {string} text - 输入字符串
     * @returns {string} 哈希ID（十六进制）
     */
    private generateHashId(text: string): string {
        const murmur = require("murmurhash-js");
        const hashString = murmur.murmur3(text, 'hawk@#');
        return hashString.toString(16);
    }

    /**
     * 异步初始化导航状态
     * 从 IndexDB 读取上次保存的选中状态并恢复
     * @returns {Promise<void>}
     */
    public async init(): Promise<void> {
        const savedId = await this.db.getSetting(this.panelId);
        if (!savedId) {
            this.content.hidden = true;
        }
    }

    /**
     * 挂载导航项到面板
     * @param {...string} navItemTexts - 导航项文本列表
     * @returns {Promise<void>}
     */
    public async navItemMount(...navItemTexts: string[]): Promise<void> {
        const savedId = await this.db.getSetting(this.panelId);
        
        for (const text of navItemTexts) {
            const panelItem = this.createPanelItem(text);
            const contentItem = this.createContentItem(text);
            
            this.panel.appendChild(panelItem);
            this.content.appendChild(contentItem);
            
            await this.restoreItemState(panelItem, contentItem, savedId);
            this.bindItemEvents(panelItem, contentItem);
        }
        await this.init(); // 初始化导航状态
    }

    /**
     * 创建导航面板项元素
     * @param {string} text - 导航项文本
     * @returns {HTMLDivElement} 创建的面板项元素
     */
    private createPanelItem(text: string): HTMLDivElement {
        const panelItem = document.createElement('div');        
        panelItem.classList = NAV.PANEL_ITEM;
        panelItem.classList.add(this.horizontal ? 'horizontal' : 'vertical');
        panelItem.textContent = text;
        panelItem.dataset.targetId = this.generateHashId(text);
        return panelItem;
    }

    /**
     * 创建内容区域项元素
     * @param {string} text - 对应导航项文本
     * @returns {HTMLDivElement} 创建的内容项元素
     */
    private createContentItem(text: string): HTMLDivElement {
        const contentItem = document.createElement('div');
        const itemId = this.generateHashId(text);
        
        contentItem.classList.add(NAV.CONTENT_ITEM);
        contentItem.id = itemId;
        this.contentItemID.set(text, itemId);
        
        return contentItem;
    }

    /**
     * 从 IndexDB 恢复导航项状态
     * @param {HTMLDivElement} panelItem - 面板项元素
     * @param {HTMLDivElement} contentItem - 内容项元素
     * @param {string | null} savedId - 保存的选中ID
     */
    private async restoreItemState(
        panelItem: HTMLDivElement,
        contentItem: HTMLDivElement,
        savedId: string | null
    ): Promise<void> {
        if (!savedId) {
            contentItem.hidden = true;
            panelItem.classList.remove('checked');
        } else if (savedId === panelItem.dataset.targetId) {
            contentItem.hidden = false;
            panelItem.classList.add('checked');
        } else {
            contentItem.hidden = true;
        }
    }

    /**
     * 绑定导航项的点击和双击事件
     * @param {HTMLDivElement} panelItem - 面板项元素
     * @param {HTMLDivElement} contentItem - 内容项元素
     */
    private bindItemEvents(panelItem: HTMLDivElement, contentItem: HTMLDivElement): void {
        panelItem.addEventListener('click', (e: MouseEvent) => {
            this.handleItemClick(e, panelItem, contentItem);
        });
        
        panelItem.addEventListener('dblclick', (e: MouseEvent) => {
            this.handleItemDoubleClick(e, panelItem);
        });
    }

    /**
     * 处理导航项单击事件
     * @param {MouseEvent} e - 鼠标事件
     * @param {HTMLDivElement} panelItem - 面板项元素
     * @param {HTMLDivElement} contentItem - 内容项元素
     */
    private async handleItemClick(
        e: MouseEvent,
        panelItem: HTMLDivElement,
        contentItem: HTMLDivElement
    ): Promise<void> {
        const target = e.target as HTMLElement;
        if (panelItem !== target) return;

        const currentChecked = this.panel.querySelector('.checked') as HTMLElement | null;
        const navContent = document.querySelector(`.${NAV.CONTENT}`) as HTMLElement | null;

        if (!currentChecked) {
            this.showContent(contentItem, panelItem);
            await this.saveCurrentId(contentItem.id);
        } else if (panelItem !== currentChecked) {
            this.switchContent(currentChecked, panelItem, contentItem);
            await this.saveCurrentId(contentItem.id);
            
            if (navContent?.hidden) {
                navContent.hidden = false;
            }
        }
    }

    /**
     * 处理导航项双击事件（关闭面板）
     * @param {MouseEvent} e - 鼠标事件
     * @param {HTMLDivElement} panelItem - 面板项元素
     */
    private async handleItemDoubleClick(e: MouseEvent, panelItem: HTMLDivElement): Promise<void> {
        const target = e.target as HTMLElement;
        const checked = this.panel.querySelector('.checked') as HTMLElement | null;
        
        if (panelItem === target && panelItem === checked) {
            this.hideAllContent(panelItem);
            await this.clearSavedId();
        }
    }

    /**
     * 显示内容区域
     * @param {HTMLDivElement} contentItem - 内容项元素
     * @param {HTMLDivElement} panelItem - 面板项元素
     */
    private showContent(contentItem: HTMLDivElement, panelItem: HTMLDivElement): void {
        this.content.hidden = false;
        contentItem.hidden = false;
        panelItem.classList.add('checked');
    }

    /**
     * 切换内容显示
     * @param {HTMLElement} currentChecked - 当前选中的面板项
     * @param {HTMLDivElement} newPanelItem - 新选中的面板项
     * @param {HTMLDivElement} newContentItem - 新内容项元素
     */
    private switchContent(
        currentChecked: HTMLElement,
        newPanelItem: HTMLDivElement,
        newContentItem: HTMLDivElement
    ): void {
        currentChecked.classList.remove('checked');
        
        const currentContent = document.getElementById(currentChecked.dataset.targetId || '') as HTMLElement | null;
        if (currentContent) {
            currentContent.hidden = true;
        }
        
        newContentItem.hidden = false;
        newPanelItem.classList.add('checked');
    }

    /**
     * 隐藏所有内容
     * @param {HTMLDivElement} panelItem - 面板项元素
     */
    private hideAllContent(panelItem: HTMLDivElement): void {
        panelItem.classList.remove('checked');
        
        const checkedContent = document.getElementById(panelItem.dataset.targetId || '') as HTMLElement | null;
        if (checkedContent) {
            checkedContent.hidden = true;
        }
        
        const navContent = document.querySelector(`.${NAV.CONTENT}`) as HTMLElement | null;
        if (navContent) {
            navContent.hidden = true;
        }
    }

    /**
     * 保存当前选中的内容ID到 IndexDB
     * @param {string} id - 内容项ID
     */
    private async saveCurrentId(id: string): Promise<void> {
        try {
            await this.db.setSetting(this.panelId, id, 'navigator');
        } catch (error) {
            console.error('Failed to save navigator state:', error);
        }
    }

    /**
     * 清除保存的选中状态
     */
    private async clearSavedId(): Promise<void> {
        try {
            await this.db.deleteSetting(this.panelId);
        } catch (error) {
            console.error('Failed to clear navigator state:', error);
        }
    }

    /**
     * 挂载内容到指定导航项
     *
     * 在 this.content 内部按 id 查找目标内容项（使用属性选择器 [id="..."]，
     * 兼容以数字开头的十六进制哈希 id），不依赖 document.getElementById，
     * 因此内容容器尚未连接到 document 时也能正确挂载。
     * @param {string} navItemText - 导航项文本
     * @param {...HTMLElement} HTMLElements - 要挂载的HTML元素
     */
    public navContentItemMount(navItemText: string, ...HTMLElements: HTMLElement[]): void {
        const navContentItemID = this.contentItemID.get(navItemText) || this.generateHashId(navItemText);
        const navContentItem = this.content.querySelector(`[id="${navContentItemID}"]`);
        
        if (navContentItem) {
            navContentItem.append(...HTMLElements);
        } else {
            console.warn(`Nav content item not found for: ${navItemText}`);
        }
    }

    /**
     * 程序化选中指定导航项
     *
     * 将面板与内容区切换到指定导航项并持久化选中状态：
     * - 指定项即当前选中项：不做切换，仅在内容区整体隐藏时恢复显示；
     * - 当前无选中项：调用 {@link showContent} 显示指定项；
     * - 当前选中其它项：调用 {@link switchContent} 切换到指定项。
     *
     * 用于筛选等场景在菜单每次打开时复位到默认页签。
     * @param {string} navItemText - 导航项文本
     */
    public async selectItem(navItemText: string): Promise<void> {
        const itemId = this.contentItemID.get(navItemText);
        if (!itemId) return;

        const panelItem = this.panel.querySelector(`[data-target-id="${itemId}"]`) as HTMLElement | null;
        const contentItem = this.content.querySelector(`[id="${itemId}"]`) as HTMLElement | null;
        if (!panelItem || !contentItem) return;

        const currentChecked = this.panel.querySelector('.checked') as HTMLElement | null;
        if (currentChecked !== panelItem) {
            if (currentChecked) {
                this.switchContent(currentChecked, panelItem as HTMLDivElement, contentItem as HTMLDivElement);
            } else {
                this.showContent(contentItem as HTMLDivElement, panelItem as HTMLDivElement);
            }
        }
        // 内容区整体处于隐藏状态（如首次打开尚无保存状态）时恢复显示
        if (this.content.hidden) {
            this.content.hidden = false;
        }
        await this.saveCurrentId(itemId);
    }

    /**
     * 获取导航面板和内容区域元素
     * @returns {{ panel: HTMLDivElement, content: HTMLDivElement }} 面板和内容元素对象
     */
    public getElements(): { panel: HTMLDivElement; content: HTMLDivElement } {
        return { panel: this.panel, content: this.content };
    }

    /**
     * 获取当前选中的导航项ID
     * @returns {Promise<string | null>} 当前选中的内容项ID
     */
    public async getCurrentId(): Promise<string | null> {
        try {
            return await this.db.getSetting(this.panelId);
        } catch (error) {
            console.error('Failed to get navigator state:', error);
            return null;
        }
    }

    /**
     * 销毁导航组件，清理事件监听和存储
     * @returns {Promise<void>}
     */
    public async destroy(): Promise<void> {
        while (this.panel.firstChild) {
            this.panel.removeChild(this.panel.firstChild);
        }
        while (this.content.firstChild) {
            this.content.removeChild(this.content.firstChild);
        }
        this.contentItemID.clear();
        await this.clearSavedId();
    }
}