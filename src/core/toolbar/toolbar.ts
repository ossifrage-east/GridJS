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
import '../style/toolbar.scss';
import { SectionPrint } from "./pages/sectionPrint";


export interface ToolbarOptions {
    parentElement: HTMLElement;
    data: DataCollection;
    menu: Menu;
    onTabContentShow?: (width: number, height: number) => void;
}


export class Toolbar extends EventEmitter {
    private container: HTMLDivElement;
    private tabs: HTMLDivElement;
    private tab: HTMLDivElement;
    private category: HTMLDivElement;
    private content: HTMLDivElement;
    private groups: Map<string, HTMLDivElement> = new Map();
    private data: DataCollection;
    private menu: Menu;
    // private toolbars: Record<string, (BtnBase | Select) | undefined> = {};

    constructor(options: ToolbarOptions) {
        super();
        if (options.onTabContentShow) {
            this.on('tabContentShow', options.onTabContentShow);
        }
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
            title: '数据处理'
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

        this.setTabContent(TOOLBAR_TABS.START, 
                            sectionBrush.getElements().sectionBrush,  // 格式刷分区
                            sectionFont.getElements().sectionFont,  // 字体样式分区
                            sectionAlign.getElements().sectionAlign,  // 对齐方式分区
                            sectionProcess.getElement().sectionProcess  // 数据处理分区
                        ); 
        this.setTabContent(TOOLBAR_TABS.PAGE, sectionPrint.getElements().sectionPrint);                
        // this.toolbars = {...sectionFont.getToolbars(), ...sectionAlign.getToolbars()};
    }

    public setTabs() { // 设置工具栏标签
        const self = this;
        const murmur = require("murmurhash-js");
        for (let tabText of Object.values(TOOLBAR_TABS)) {
            const hashString = murmur.murmur3(tabText, 'hawk@#');
            this.tab = createDiv({className: TOOLBAR_CLASS_NAME.TAB, textContent: tabText});
            this.tab.dataset.targetId = hashString.toString(16);
            this.tabs.appendChild(this.tab);
            const group = createDiv({className: TOOLBAR_CLASS_NAME.GROUP, id: hashString.toString(16)});
            this.content.appendChild(group);
            this.groups.set(hashString.toString(16), group);
            
            this.tab.addEventListener('click',  function(e: MouseEvent) {
                const tab = e.target as HTMLDivElement;
                let actived:HTMLElement|null = self.tabs.querySelector('.active');
                if (!actived) actived = tab;
                const tabId = tab.dataset.targetId;
                const group = self.groups.get(tabId);
                
                const activedId = actived.dataset.targetId;
                const activedgroup = self.groups.get(activedId);

                if (tab !== actived) {
                    actived.classList.remove('active');
                    activedgroup.classList.remove('active');
                }
                group.classList.add('active');
                tab.classList.add('active');
            });

            this.tab.addEventListener('dblclick',  function(e: MouseEvent) {
                const tab = e.target as HTMLDivElement;
                let actived:HTMLElement|null = self.tabs.querySelector('.active');
                if (!actived) return;
                const tabId = tab.dataset.targetId;
                const group = self.groups.get(tabId);

                if (tab === actived) {
                    tab.classList.remove('active');
                    group.classList.remove('active');
                }
        });
        }
    }

    public setTabContent(tabText: string, ...sections: HTMLElement[]) {
        const murmur = require("murmurhash-js");
        const hashString = murmur.murmur3(tabText, 'hawk@#');
        const group = this.groups.get(hashString.toString(16));
        if (!group) return;

        for (let section of sections) {
            group.append(section);
        }
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
                        const rect = (entry.target as HTMLElement).parentElement.getBoundingClientRect();
                        this.emit('tabContentShow', rect.width, rect.height);
                    }
                }
            });
        } catch (error) {
            console.error('ResizeObserver error:', error);
        }
    }
    
    // public getToolbars() {
    //     return this.toolbars;
    // }
    
    public getElements() {
        return {
            container: this.container,
            tab: this.tab,
            tabs: this.tabs,
            content: this.content,
            groups: this.groups,
        }
    }
}
