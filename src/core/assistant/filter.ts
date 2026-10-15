/*  # web 表格辅助组件-筛选组件 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { createDiv, createInput } from "../../utils/dom";
import { Menu } from "../common/menu";
import { ASSISTANT, DataEvents } from "../constant";
import { DataCollection } from "../dataArchitecture/dataCollection";
import { Navigator } from "../common/navigator";
import chevronDown from '../../assets/images/toolbar/chevron-down.svg';

import '../style/filter.scss';
import { Cell } from "../dataArchitecture/cell";
import { compareMixedText } from "../dataArchitecture/matrixSorter";

/**
 * 筛选组件初始化选项
 * @property {HTMLElement} parentElement - 筛选组件的挂载父元素，筛选根节点会被追加到该元素内
 * @property {DataCollection} data - 表格数据集合，提供选区、单元格内容与颜色等筛选数据源
 * @property {Menu} menu - 右键菜单实例，筛选容器会注册为菜单项，随菜单弹出/收起
 */
export interface FilterOptions {
    parentElement: HTMLElement;
    data: DataCollection;
    menu: Menu;
}

/**
 * 表格筛选组件
 *
 * 以弹出面板形式提供「按内容」和「按颜色」两个导航页签：
 * - 按内容：关键字输入框 + 筛选结果提示
 * - 按颜色：按文字颜色、按单元格背景颜色两组筛选项
 *
 * 面板内容通过 {@link Navigator} 组织导航项，并注册到 {@link Menu} 右键菜单中随菜单显隐。
 *
 * @example
 * const filter = new Filter({
 *     parentElement: document.body,
 *     data: dataCollection,
 *     menu: contextMenu
 * });
 * filter.show();
 */
export class Filter {
    /** 筛选激活状态下按钮的主题绿色背景 */
    private static readonly ACTIVE_FILTER_BG = '#4CAF50';
    /** 筛选组件根节点（class 为 ASSISTANT.FILTER_CLASS_NAME） */
    private filter: HTMLDivElement;
    /** 表格数据集合，筛选数据的来源 */
    private data: DataCollection;
    /** 右键菜单实例，筛选容器作为菜单项挂载 */
    private menu: Menu;
    /** 导航器，负责页签面板与内容区的组织；**懒构建**（首次点击按钮时才创建） */
    private navigator: Navigator;
    /** 筭选组件箭头元素 */
    private arrowElement: HTMLDivElement;
    /** 筛选组件容器，包含导航器面板与内容区；**懒构建** */
    private container: HTMLDivElement;
    /** 「按内容」页签的筛选结果提示区，展示当前列去重后的内容清单；**懒构建** */
    private tips: HTMLDivElement;
    /**
     * 面板构建 Promise（null = 尚未构建）
     *
     * 【性能约定，勿回退为「构造时建面板」】筛选按钮在大表里可能有几十个，
     * 而每个面板都要建立一个 {@link Navigator}（读 IndexedDB + 建页签 DOM）
     * 外加输入框/清单/颜色区等一整套 DOM。历史实现把这些全部放在构造函数里，
     * 于是「打开一个带几十个筛选列的文件」＝首帧集中创建几十套面板。
     * 改为首次点击时才构建，未打开的按钮再也不付这份成本。
     */
    private _panelPromise: Promise<void> | null = null;
    /** 本实例对应的筛选单元格名称，由 setPosition 记录，用于确定内容清单的统计列 */
    private currentCellName: string = '';
    /** 筛选按钮所在列号（1-based），由 setPosition 记录，筛选时按此列内容过滤 */
    private currentCol: number = 0;
    /** 筛选按钮所在行号（1-based），由 setPosition 记录，该行以上的行不参与筛选 */
    private currentRow: number = 0;
    /** 点击「确定」时勾选的内容集合，作为当前列筛选的保留内容 */
    private confirmedContents: Set<string> = new Set();
    /** 绑定后的点击事件处理函数，保证 add/removeEventListener 引用一致 */
    private readonly boundHandleClick: (e: MouseEvent) => void = this.handleClick.bind(this);
    /**
     * 按钮视觉尺寸缓存（随 zoom 变化失效）
     *
     * `getBoundingClientRect()` 会强制浏览器同步重排，而 {@link setPosition}
     * 又被主画布**每帧**按每个筛选按钮调用一次：30 个按钮 = 每帧 30 次强制布局，
     * 是「筛选按钮一多就卡」的直接原因之一。按钮尺寸只由 zoom 决定，
     * 因此这里按 zoom 缓存，稳态下每帧零次测量。
     */
    private _buttonSize: { zoom: number; width: number; height: number } | null = null;
    /** 上一次写入的按钮位置（CSS 像素）：位置未变则完全不碰 DOM */
    private _appliedLeft: number | null = null;
    private _appliedTop: number | null = null;
    /** 是否已按当前行列号同步过一次按钮显隐（滚动重绘不必重复全表扫描） */
    private _visibilitySynced: boolean = false;
    /**
     * 数据变化监听（DataEvents.VALUES_CHANGED）：
     * 本列存在筛选条件且筛选行以下的数据已全部显示（无任何隐藏行，
     * 如实质取消隐藏/全选筛选恢复）时，清除本列筛选条件并恢复按钮默认背景色。
     *
     * 撤销/重做回写期间（data.isApplyingUndo）：状态由全量快照恢复，
     * 仅同步按钮激活外观，跳过条件自动清除，避免恢复过程中的中间态误触发入栈。
     */
    private readonly handleValuesChanged = (): void => {
        // 列内容变化：同步按钮显隐（该列无内容时不显示筛选按钮）
        this.updateFilterVisibility();
        if (!this.currentCol) return;
        // 取消隐藏后行/列位置变化，按钮需重新定位到当前单元格的新坐标；
        // 仅在按钮可见时重算（隐藏态下 setPosition 会因 isHidden 早退，徒增开销）
        if (this.filter.style.display !== 'none' && this.currentCellName) {
            const cell = this.data.getCellByName(this.currentCellName) || { cell: this.currentCellName } as Cell;
            this.setPosition(cell);
        }
        const col = this.currentCol;
        const cond = this.data.getFilterCondition(col);
        // 回写期间：条件随快照恢复，仅按条件存在与否同步按钮外观
        if (this.data.isApplyingUndo) {
            this.setFilterButtonActive(!!cond);
            return;
        }
        if (!cond) return;
        // 判定本列数据是否全部显示：筛选行以下无任何隐藏行
        // （行 isHidden 的变更只经 setRowsHidden 统一入口，末尾必发 VALUES_CHANGED）
        // 【性能约定，勿回退为「for r = anchorRow+1..rowHeaders.length」】
        // 大表 rowHeaders.length 可达百万级，线性扫描每次着色/操作都触发。
        // 改为遍历 _hiddenRowsSet（已隐藏行通常只有几百到几千行）。
        const hiddenSet = this.data.getHiddenRowsSet();
        for (const r of hiddenSet) {
            if (r > cond.anchorRow) return;  // 筛选行以下仍有隐藏行 → 筛选仍生效，保持绿色
        }
        // 全部显示：清除本列筛选条件（作为可撤销步骤，嵌套在外层操作时折叠为其一步），
        // 按钮恢复默认色，内容清单恢复全部内容
        this.data.runWithFullStateUndo('清除筛选', () => {
            this.data.clearFilterCondition(col);
        });
        this.setFilterButtonActive(false);
        this.updateContentTips();
    };

    /**
     * 创建筛选组件实例
     *
     * 构造只做**廉价**的按钮本体：根节点 + 箭头 + 事件订阅。
     * 面板（{@link Navigator} / 容器 / 内容清单）改为**首次点击时**才构建
     * （见 {@link _panelPromise}）—— 大表里筛选按钮可能几十个，构造时
     * 各建一整套面板会让「打开文件」的首帧明显卡顿。
     * @param {FilterOptions} options - 筛选组件初始化选项
     */
    constructor(options: FilterOptions) {
        this.filter = createDiv({
            className: ASSISTANT.FILTER_CLASS_NAME,
        });

        if (options.parentElement) {
            options.parentElement.append(this.filter);
        }

        this.data = options.data;
        this.menu = options.menu;
        this.setArrow();
        this.setupEventListeners();
    }

    /**
     * 确保面板已构建（首次点击时调用；同一实例只构建一次）
     * @returns {Promise<void>} 面板可用
     * @private
     */
    private ensurePanel(): Promise<void> {
        if (this._panelPromise === null) {
            this._panelPromise = this.initPanel();
        }
        return this._panelPromise;
    }

    /**
     * 构建筛选面板内部结构（懒构建入口）
     *
     * 创建筛选容器，实例化 {@link Navigator} 并将其面板与内容区加入容器，
     * 容器注册到右键菜单后挂载「按内容」「按颜色」两个导航页签，
     * 最后通过 {@link mountContent} 填充各页签内容。
     *
     * 注意：navItemMount 内部读取 IndexedDB（异步），必须 await 完成后
     * 再调用 mountContent，否则 contentItemID 映射与内容项元素尚未创建，
     * navContentItemMount 会因找不到目标元素而挂载失败。
     * @private
     * @returns {Promise<void>}
     */
    private async initPanel(): Promise<void> {
        this.container = createDiv({
            className: ASSISTANT.FILTER_CONTAINER_CLASS_NAME,
        });
        this.navigator = new Navigator(true);
        const { panel, content } = this.navigator.getElements();
        this.container.append(panel, content);
        // 等待页签与内容项创建完成，再挂载各页签内容
        await this.navigator.navItemMount('按内容', '按颜色');
        this.mountContent();
    }

    /**
     * 设置箭头
     */
    private setArrow() {
        const dataURI = `data:image/svg+xml;base64,${btoa(chevronDown.replace('data:image/svg+xml,', '').replace(/%3c/g, '<').replace(/%3e/g, '>'))}`;
        this.arrowElement = createDiv({
            dataset: {
                todo: `menu`,
            },
            style: {
                position: 'relative',
                margin: '1px',
                width: '10px',
                height: '10px',
                backgroundImage: `url(${dataURI})`,
                backgroundRepeat: 'no-repeat',
                backgroundPosition: 'center center',
                backgroundSize: '10px 10px',
                transition: 'transform 0.3s ease',
            },
        });
        this.filter.appendChild(this.arrowElement);
    }

    /**
     * 设置筛选组件位置
     *
     * 将筛选组件定位到当前单元格的右下角：
     * 1. 通过 {@link DataCollection.calculateCellRect} 获取单元格的视图矩形
     *    （已处理合并单元格、colspan/rowspan，并已减去滚动偏移）；
     * 2. 右下角屏幕坐标 =（left + width）* zoom + 行列表头宽度、（top + height）* zoom + 列表头高度，
     *    与编辑器 textArea 的定位方式保持一致；
     * 3. 使用 Math.round 取整，避免亚像素渲染模糊。
     *
     * 【性能约定，勿回退】本方法被主画布**每帧**按每个筛选按钮调用一次，
     * 因此做了两处「无变化就不干活」的处理（否则筛选按钮一多，每帧几十次
     * 强制同步布局 + 几十次全表扫描，滚动就会明显掉帧）：
     * - 按钮尺寸按 zoom 缓存（{@link _buttonSize}），不再每帧 `getBoundingClientRect()`；
     * - 目标坐标与上次相同则跳过 style 写入；
     * - 可见性（该列筛选行以下是否有内容）只随**数据变化**改变，滚动/重绘不需要重算，
     *     故仅在按钮绑定的行列号变化时调用 {@link updateFilterVisibility}
     *     （数据变化本身有 VALUES_CHANGED 监听负责）。
     * @param {Cell} cell - 筛选组件要定位到的单元格
     * @returns {void}
     */
    public setPosition(cell: Cell): void {
        // 记录本实例对应的筛选单元格，打开菜单时据此统计该列内容
        this.currentCellName = cell.cell;
        const { col, row } = this.data.getCellColAndRow(cell.cell);
        // 记录筛选按钮所在的行列号（1-based），筛选执行时使用
        const anchorChanged = col !== this.currentCol || row !== this.currentRow;
        this.currentCol = col;
        this.currentRow = row;
        const { left, top, width, height } = this.data.calculateCellRect(col, row);
        const zoom = this.data.zoom;
        // 先应用缩放变换，再读取尺寸：getBoundingClientRect 返回的是变换后的视觉尺寸，
        // 因此「右下角对齐单元格右下角」的算法在任意缩放下自动成立（详见 applyZoomScale）
        if (this._buttonSize === null || this._buttonSize.zoom !== zoom) {
            this.applyZoomScale();
            const rect = this.filter.getBoundingClientRect();
            // 尺寸为 0 说明此刻尚未完成布局（如父容器暂不可见）：不缓存，下次再测
            if (rect.width > 0 || rect.height > 0) {
                this._buttonSize = { zoom, width: rect.width, height: rect.height };
            }
        }
        const buttonSize = this._buttonSize;
        if (buttonSize === null) return;   // 尚无有效尺寸：本次不定位（下一帧会重试）
        // 单元格右下角的屏幕坐标（内容坐标乘 zoom 后叠加固定行/列表头尺寸）
        const _left = Math.round((left + width) * zoom + this.data.visibleView.rowHeaderWidth - buttonSize.width);
        const _top = Math.round((top + height) * zoom + this.data.visibleView.colHeaderHeight - buttonSize.height);
        if (_left !== this._appliedLeft || _top !== this._appliedTop) {
            this.filter.style.left = `${_left}px`;
            this.filter.style.top = `${_top}px`;
            this._appliedLeft = _left;
            this._appliedTop = _top;
        }
        // 定位时同步显隐：该列无内容的单元格不显示筛选按钮
        if (anchorChanged || !this._visibilitySynced) {
            this._visibilitySynced = true;
            this.updateFilterVisibility();
        }
    }

    /**
     * 按主页面当前缩放比例对筛选按钮应用等比缩放
     *
     * 按钮需与单元格内容同步放大/缩小（否则缩放后按钮相对单元格显得过大/过小）。
     * 采用 transform 而非改 width/height：不影响内部箭头布局，且只触发合成层。
     *
     * **变换原点必须是左上角**，这是 {@link setPosition} 定位算法成立的前提：
     * left/top 定位的是「未变换盒子」的左上角，以左上角为原点时，
     * 按钮的视觉右下角 = (left + rect.width, top + rect.height)，
     * 而 rect 取到的正是变换后的尺寸——于是「用视觉尺寸从单元格右下角回退」即可精确对齐，
     * 无需按 zoom 手工折算尺寸（若用默认的 50% 原点，视觉框会向中心收拢，定位将偏移）。
     * @private
     * @returns {void}
     */
    private applyZoomScale(): void {
        const zoom = this.data.zoom;
        this.filter.style.transformOrigin = 'top left';
        // 1:1 时移除变换，避免无谓的合成层与亚像素模糊
        this.filter.style.transform = zoom === 1 ? '' : `scale(${zoom})`;
    }

    /**
     * 更新筛选按钮的显示状态
     *
     * 统计本列筛选行以下是否存在非空内容（合并单元格覆盖位置取锚点文本，
     * colspan 跨列的锚点覆盖到本列也计入），有内容时显示按钮，
     * 全部为空时隐藏按钮并收起已打开的筛选菜单。
     *
     * 【性能约定，勿回退为「遍历整表 values」】
     * 本方法挂在**每个**筛选实例的 VALUES_CHANGED 上，每次数据变化都会走一遍。
     * 早期写法对每个单元格都做 `getCellColAndRow()`（正则解析）与
     * `chars.map(c => c.char).join('')`（每格再分配一个数组和一个字符串），
     * 后来又改成「整表单遍 + 列号先行比较」—— 单次确实变快了，但**实例数没变**：
     * F 个按钮仍是 F 遍整表。基准 `node test/_bench-filter-many.js` 实测
     * 4 万格 × 40 个按钮 = **17.6 ms / 次数据变化**，400 万格量级直接上百毫秒。
     *
     * 现在改为向 {@link DataCollection.hasContentBelow} 问一次 O(1) 的结果：
     * 列内容索引由数据集合按**世代**缓存（同一世代只构建一次），F 个按钮共用，
     * 于是每次数据变化的总成本从 F × O(N) 降到 O(N)（构索引）+ O(F)（查表）。
     * @private
     * @returns {void}
     */
    private updateFilterVisibility(): void {
        if (!this.currentCol) return;
        // 列被隐藏时筛选按钮一并隐藏（隐藏列不可交互）；
        // setColsHidden 末尾必发 VALUES_CHANGED，本方法会被调用，无需额外监听
        if (this.data.colHeaders.getAt(this.currentCol - 1)?.isHidden) {
            this.filter.style.display = 'none';
            if (this.menu.hasMenu === this.filter) this.menu.closeMenu();
            return;
        }
        // 行被隐藏时筛选按钮一并隐藏（隐藏行不可交互）；
        // setRowsHidden 末尾必发 VALUES_CHANGED，本方法同样会被调用
        if (this.data.rowHeaders.getAt(this.currentRow - 1)?.isHidden) {
            this.filter.style.display = 'none';
            if (this.menu.hasMenu === this.filter) this.menu.closeMenu();
            return;
        }
        // O(1)：该列筛选行以下是否还有文本（语义与历史上的整表遍历完全一致）
        if (this.data.hasContentBelow(this.currentCol, this.currentRow)) {
            this.filter.style.display = '';
            return;
        }
        // 该列筛选行以下无任何内容：隐藏筛选按钮（菜单打开时一并收起）
        this.filter.style.display = 'none';
        if (this.menu.hasMenu === this.filter) this.menu.closeMenu();
    }

    /**
     * 移除筛选组件实例
     *
     * 面板可能是**懒构建**且从未打开过：此时 {@link navigator} 尚未创建，
     * 需按下标判空后再销毁，否则会在「删列/删行触发反向清理」时抛错。
     * @returns {void}
     */
    public destroy(): void {
        this.filter.removeEventListener('click', this.boundHandleClick);
        this.data.off(DataEvents.VALUES_CHANGED, this.handleValuesChanged);
        this.filter.remove();
        if (this.navigator) this.navigator.destroy();
        // 从筛选条件注册表中移除本列条件
        this.data.clearFilterCondition(this.currentCol);
    }
    /**
     * 挂载筛选面板各页签的内容
     *
     * - 「按内容」页签：关键字输入框（#filter-input）、筛选结果提示区
     *   与 tips 下方的确定/取消按钮区（确定收集勾选内容并关闭面板，取消直接关闭）
     * - 「按颜色」页签：文字颜色筛选标签与颜色容器、单元格背景颜色筛选标签与颜色容器
     *
     * 所有元素通过 {@link Navigator.navContentItemMount} 追加到对应页签内容项下。
     * @private
     * @returns {void}
     */
    private mountContent(): void {
        // 「按内容」页签：关键字输入框
        const input = createInput({
            className: ASSISTANT.FILTER_INPUT_CLASS_NAME,
            id: `filter-input`,
        });
        // 「按内容」页签：筛选结果提示区
        const tips = createDiv({
            className: ASSISTANT.FILTER_TIPS_CLASS_NAME,
        });
        this.tips = tips;
        // 「按内容」页签：tips 下方的确定/取消按钮区
        const buttons = createDiv({
            className: ASSISTANT.FILTER_BUTTONS_CLASS_NAME,
        });
        const confirmBtn = createDiv({
            className: ASSISTANT.FILTER_CONFIRM_CLASS_NAME,
        });
        confirmBtn.textContent = '确定';
        const cancelBtn = createDiv({
            className: ASSISTANT.FILTER_CANCEL_CLASS_NAME,
        });
        cancelBtn.textContent = '取消';
        // 确认：收集勾选内容并关闭筛选面板
        confirmBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            this.onConfirm();
        });
        // 取消：直接关闭筛选面板
        cancelBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            this.menu.closeMenu();
        });
        buttons.append(confirmBtn, cancelBtn);
        this.navigator.navContentItemMount('按内容', input, tips, buttons);

        // 「按颜色」页签：文字颜色筛选标签与颜色选项容器
        const labelColor = createDiv({
            className: ASSISTANT.FILTER_COLOR_LABEL_CLASS_NAME,
        });
        labelColor.innerText = '按文字颜色筛选：';
        const colorContainer = createDiv({
            className: ASSISTANT.FILTER_COLOR_CONTAINER_CLASS_NAME,
        });
        // 「按颜色」页签：单元格背景颜色筛选标签与颜色选项容器
        const labelBgColor = createDiv({
            className: ASSISTANT.FILTER_COLOR_LABEL_CLASS_NAME,
        });
        labelBgColor.innerText = '按单元格背景颜色筛选：';
        const bgColorContainer = createDiv({
            className: ASSISTANT.FILTER_COLOR_CONTAINER_CLASS_NAME,
        });
        this.navigator.navContentItemMount('按颜色', labelColor, colorContainer, labelBgColor, bgColorContainer);
    }

    /**
     * 设置筛选组件事件监听
     * @private
     * @returns {void}
     */
    private setupEventListeners(): void {
        this.filter.addEventListener('click', this.boundHandleClick);
        // 监听数据变化：本列数据全部显示时（如实质取消隐藏）恢复按钮默认背景色
        this.data.on(DataEvents.VALUES_CHANGED, this.handleValuesChanged);
    }

    /**
     * 筛选组件点击事件处理函数
     *
     * 切换菜单显隐：
     * - 菜单已由本按钮打开（menu.hasMenu 指向本组件根节点）→ 关闭菜单；
     * - 否则 → 先关闭以清理残留状态并复位其它触发元素的箭头，
     *   再打开菜单并重新挂载筛选容器。
     *
     * 面板（Navigator / 容器 / 清单）在**首次点击**时构建（{@link ensurePanel}），
     * 因此本方法是 async：构建完成后再打开菜单。
     *
     * 注意：{@link Menu.closeMenu} 会清空菜单容器的所有子元素，
     * 因此每次打开前必须重新调用 {@link Menu.addMenuElement} 挂载容器；
     * openMenu 须传本组件根节点（箭头是其后代），箭头旋转与菜单位置定位才能生效。
     * @private
     * @param {MouseEvent} e - 点击事件
     * @returns {Promise<void>}
     */
    private async handleClick(e: MouseEvent): Promise<void> {
        e.stopPropagation();
        // 已打开 → 切换为隐藏
        if (this.menu.hasMenu === this.filter) {
            this.menu.closeMenu();
            return;
        }
        // 未打开 → 先清理，再打开并重新挂载内容
        this.menu.closeMenu();
        // 懒构建：首次点击才建 Navigator / 容器 / 清单（见 Filter._panelPromise）
        await this.ensurePanel();
        // 构建期间按钮可能已被隐藏（如所在列变为空）：此时不弹菜单
        if (this.filter.style.display === 'none') return;
        this.menu.openMenu(this.filter);
        this.menu.addMenuElement(this.container);
        // 每次打开时默认显示「按内容」页签，并按当前单元格所在列刷新内容清单
        this.navigator.selectItem('按内容');
        this.updateContentTips();
    }

    /**
     * 刷新「按内容」页签提示区的内容清单
     *
     * 第一行为头部行：全选多选框（勾选/取消全部内容行）+「全选」文字（点击全部勾选）
     * + 去重后的内容条数（共N条）+「反选」（点击内容行勾选状态全部取反）。
     *
     * 统计当前单元格所在列的所有单元格内容（去除重复后），逐行渲染到 tips 中：
     * 仅统计筛选单元格**以下**同列的内容（不含筛选单元格所在行，
     * 与 Excel 筛选行为一致）：
     * - 行首：多选框（默认勾选）
     * - 内容：去重后的单元格文本
     * - 内容后：该内容在列中的重复次数
     * - 行尾：悬停该行时在最右侧显示「仅筛选此项」，
     *   点击后取消其它行勾选、仅保留当前行，实现单项筛选预选
     *
     * 空白单元格（无内容或列上不存在的位置）单独计数，
     * 作为「空白」项显示在内容清单的最前面并标注个数；
     * 不存在空白单元格时不显示该项，头部条数同步计入空白项。
     *
     * 统计列取本实例绑定的筛选单元格（setPosition 记录的 currentCellName），
     * 缺省回退到 data.activedCell，保证菜单内容与触发筛选按钮所在列一致。
     *     
     * @private
     * @returns {void}
     */
    private updateContentTips(): void {
        // 面板是懒构建的：按钮从未展开过时清单区还不存在，无需刷新
        if (!this.tips) return;
        // 优先使用本实例绑定的筛选单元格，缺省回退到当前激活单元格
        const cellName = this.currentCellName || this.data.activedCell;
        if (!cellName) return;
        const _cell = this.data.getCellByName(cellName);
        if (!_cell || _cell.filter === undefined) return;

        this.tips.textContent = '';
        const { col, row: startRow } = this.data.getCellColAndRow(cellName);
        // 统计范围：筛选行以下 → 「已用内容范围」的最后一行。
        // 【性能约定，勿改回 matrixValues()】历史实现仅为问「统计到第几行」就调用
        // matrixValues()，而它会分配一张 maxRow × maxCol 的稠密二维数组
        // （大表下几百万个 undefined 槽位、几十 MB 临时内存 + 每格一次名字解析）。
        // 现在改为问按世代缓存的零分配内容范围。
        const maxRow = this.data.getContentBounds().maxRow;

        // 统计筛选单元格以下同列的内容（不含筛选单元格所在行，与 Excel 筛选行为一致）：
        // 空白单元格单独计数，非空白内容按出现顺序去重统计。
        // 本列已处于筛选状态 → 统计筛选前全部内容（不跳过隐藏行），并按当前条件设置勾选状态；
        // 其它列 → 只统计当前可见（非隐藏）行的内容，体现已生效筛选的叠加语义。
        // 合并单元格覆盖位置取锚点有效文本（getEffectiveTextAt 为 O(1) 查询，
        // 历史实现在这里构建了整表「位置 → 锚点」的字符串键 Map）。
        const activeKeep = this.data.getFilterCondition(col)?.keep ?? null;
        const countMap = new Map<string, number>();
        let blankCount = 0;
        for (let r = startRow + 1; r <= maxRow; r++) {
            if (!activeKeep && this.data.rowHeaders.getAt(r - 1)?.isHidden) continue;  // 非筛选列：已隐藏的行不计入清单
            const text = this.data.getEffectiveTextAt(r, col);
            if (text === '') {
                blankCount++;
            } else {
                countMap.set(text, (countMap.get(text) ?? 0) + 1);
            }
        }
        // 清单总行数（空白作为一项计入）
        const totalCount = countMap.size + (blankCount > 0 ? 1 : 0);

        // 辅助函数：获取所有内容行的多选框（不含头部多选框）
        const getItemCheckboxes = (): HTMLInputElement[] =>
            Array.from(this.tips.querySelectorAll<HTMLInputElement>('.filter-tips-item input[type="checkbox"]'));

        // 头部行：全选多选框 + 「全选」+ 内容条数 + 「反选」
        const headerRow = createDiv({
            className: ASSISTANT.FILTER_TIPS_HEADER_CLASS_NAME,
        });
        const headerCheckbox = document.createElement('input');
        headerCheckbox.type = 'checkbox';
        headerCheckbox.checked = true;
        const selectAllSpan = createDiv({
            className: ASSISTANT.FILTER_TIPS_TEXT_CLASS_NAME,
        });
        selectAllSpan.textContent = '全选';
        const totalSpan = createDiv({
            className: ASSISTANT.FILTER_TIPS_COUNT_CLASS_NAME,
        });
        totalSpan.textContent = `共${totalCount}条`;
        const invertSpan = createDiv({
            className: ASSISTANT.FILTER_TIPS_INVERT_CLASS_NAME,
        });
        invertSpan.textContent = '反选';

        // 同步头部多选框：所有内容行均勾选时才勾选
        const syncHeaderCheckbox = (): void => {
            const boxes = getItemCheckboxes();
            headerCheckbox.checked = boxes.length > 0 && boxes.every((cb) => cb.checked);
        };
        // 头部多选框：勾选/取消全部内容行
        headerCheckbox.addEventListener('change', () => {
            getItemCheckboxes().forEach((cb) => {
                cb.checked = headerCheckbox.checked;
            });
        });
        // 点击「全选」→ 全部内容行勾选
        selectAllSpan.addEventListener('click', (e) => {
            e.stopPropagation();
            getItemCheckboxes().forEach((cb) => {
                cb.checked = true;
            });
            headerCheckbox.checked = true;
        });
        // 点击「反选」→ 内容行勾选状态全部取反
        invertSpan.addEventListener('click', (e) => {
            e.stopPropagation();
            getItemCheckboxes().forEach((cb) => {
                cb.checked = !cb.checked;
            });
            syncHeaderCheckbox();
        });

        headerRow.append(headerCheckbox, selectAllSpan, totalSpan, invertSpan);
        this.tips.append(headerRow);

        /**
         * 构建一条清单行：多选框 + 内容 + 重复数量 + 悬停显示「仅筛选此项」
         * @param {string} text - 行内容文本
         * @param {number} count - 该内容在列中的出现次数
         * @param {boolean} checked - 多选框初始勾选状态（本列筛选中时按条件设置）
         * @returns {HTMLDivElement} 清单行元素
         */
        const createItemRow = (text: string, count: number, checked: boolean = true): HTMLDivElement => {
            const row = createDiv({
                className: ASSISTANT.FILTER_TIPS_ITEM_CLASS_NAME,
            });
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.checked = checked;
            const textSpan = createDiv({
                className: ASSISTANT.FILTER_TIPS_TEXT_CLASS_NAME,
            });
            textSpan.textContent = text;
            const countSpan = createDiv({
                className: ASSISTANT.FILTER_TIPS_COUNT_CLASS_NAME,
            });
            countSpan.textContent = `×${count}`;
            const onlySpan = createDiv({
                className: ASSISTANT.FILTER_TIPS_ONLY_CLASS_NAME,
            });
            onlySpan.textContent = '仅筛选此项';

            // 点击多选框：阻止冒泡到行（避免行处理器再次取反），并同步头部多选框
            checkbox.addEventListener('click', (e) => {
                e.stopPropagation();
            });
            checkbox.addEventListener('change', () => {
                syncHeaderCheckbox();
            });
            // 点击行其它区域切换该行勾选状态
            row.addEventListener('click', () => {
                checkbox.checked = !checkbox.checked;
                syncHeaderCheckbox();
            });
            // 点击「仅筛选此项」→ 仅保留当前行勾选，并立即执行单项筛选后关闭面板
            onlySpan.addEventListener('click', (e) => {
                e.stopPropagation();
                getItemCheckboxes().forEach((cb) => {
                    cb.checked = false;
                });
                checkbox.checked = true;
                syncHeaderCheckbox();
                // 「空白」是显示文案，实际参与匹配的内容为空字符串
                this.applyFilter(new Set([text === '空白' ? '' : text]));
                this.menu.closeMenu();
            });

            row.append(checkbox, textSpan, countSpan, onlySpan);
            return row;
        };

        // 渲染内容清单。本列处于筛选状态时，勾选状态按当前生效条件设置：
        // 条件中包含的内容勾选，被筛掉的内容不勾选（清单本身仍显示筛选前全部内容）
        const isItemChecked = (displayText: string): boolean =>
            activeKeep ? activeKeep.has(displayText === '空白' ? '' : displayText) : true;
        // 清单可能几百上千行：先全部挂到 DocumentFragment 上，最后**一次性**插入。
        // 逐行 append 会让每行都触发一次插入与样式重算，打开下拉时的卡顿相当一部分来自这里。
        const listFragment = document.createDocumentFragment();
        // 空白行排在清单最前面（存在空白单元格时）
        if (blankCount > 0) {
            listFragment.append(createItemRow('空白', blankCount, isItemChecked('空白')));
        }
        // 非空白内容按 时间 → 数字 → 英文 → 中文 → 其它外语 的类型顺序升序排列后渲染
        const sortedEntries = Array.from(countMap.entries()).sort((a, b) => compareMixedText(a[0], b[0]));
        for (const [text, count] of sortedEntries) {
            listFragment.append(createItemRow(text, count, isItemChecked(text)));
        }
        this.tips.append(listFragment);
        // 头部多选框状态与内容行勾选状态同步（本列筛选中且部分内容被筛掉时头部不勾选）
        syncHeaderCheckbox();
    }

    /**
     * 「确定」按钮点击处理
     *
     * 收集 tips 内容清单中所有勾选行的内容文本（头部行除外），
     * 存入 {@link confirmedContents} 作为当前列筛选的保留内容，
     * 随后关闭筛选面板。
     * @private
     * @returns {void}
     */
    private onConfirm(): void {
        const contents = new Set<string>();
        const items = Array.from(this.tips.querySelectorAll<HTMLDivElement>(`.${ASSISTANT.FILTER_TIPS_ITEM_CLASS_NAME}`));
        for (const item of items) {
            const checkbox = item.querySelector<HTMLInputElement>('input[type="checkbox"]');
            const text = item.querySelector<HTMLDivElement>(`.${ASSISTANT.FILTER_TIPS_TEXT_CLASS_NAME}`)?.textContent ?? '';
            if (checkbox?.checked && text !== '') {
                // 「空白」是显示文案，实际参与匹配的内容为空字符串
                contents.add(text === '空白' ? '' : text);
            }
        }
        this.confirmedContents = contents;
        // 按勾选内容执行筛选：只显示匹配内容所在的行，其余行隐藏
        this.applyFilter(contents);
        this.menu.closeMenu();
    }

    /**
     * 执行筛选：将本列条件写入 DataCollection 筛选条件注册表，按操作顺序重放全部
     * 已筛选列的条件，通过 {@link DataCollection.setRowsHidden} 统一入口隐藏/显示行 ——
     * 与右键菜单的实质隐藏/取消隐藏功能完全一致（含尾部补足/收缩、滚动 clamp 等副作用）。
     *
     * 整个操作（条件注册表写入 + 行隐藏状态变更）包裹在
     * {@link DataCollection.runWithFullStateUndo} 中，折叠为一步可撤销/重做操作。
     *
     * 处理规则：
     * - 本列条件经 {@link DataCollection.setFilterCondition} 写入注册表（按操作顺序记录）；
     * - 「应当隐藏哪些行」由 {@link DataCollection.computeFilterHiddenRows} 统一重放得出，
     *   与「清除筛选时恢复隐藏行」共用同一份判定（不会漂移）；
     * - 只重新判定「有内容的行」与「当前已隐藏的行」：其余行无论筛选条件如何都不会变化；
     * - 对比行当前 isHidden 状态，将需要变化的连续行分段调用 setRowsHidden；
     * - 筛选完成后本列按钮背景按筛选效果标为主题绿色。
     *
     * 【性能约定，勿回退为「整表扫到总行数 + 位置→锚点字符串键 Map」】
     * 历史实现先 `buildAnchorMap()` 遍历整张表、为每个被合并覆盖的位置塞一条
     * `${row}:${col}` 字符串键（大表下上百毫秒），再对 `1..总行数` 逐行 × 全部已筛选列
     * 查这个 Map。大文件总行数可达百万行，而有内容的往往只有几千行 ——
     * 这是「点击筛选按钮、执行筛选时卡顿」的主因。
     * @param {Set<string>} keepContents - 本列保留的内容集合（「空白」已转为空字符串）
     * @private
     * @returns {void}
     */
    private applyFilter(keepContents: Set<string>): void {
        if (!this.currentCol) return;
        const col = this.currentCol;
        const anchorRow = this.currentRow;

        // 条件注册表 + 行隐藏状态整体作为一步可撤销操作
        this.data.runWithFullStateUndo(`筛选（${this.data.getColName(col)}列）`, () => {
            // 1) 本列条件写入注册表（按操作顺序记录已筛选列）
            this.data.setFilterCondition(col, anchorRow, keepContents);

            // 2) 重放全部列条件 → 应当因筛选而隐藏的行
            const targetHidden = this.data.computeFilterHiddenRows();

            // 3) 需要重新判定的行 = 本列筛选行以下「有内容的行」∪「当前已隐藏的行」
            //
            // 【性能约定，勿回退为「for r = anchorRow+1..totalRows」】
            // 历史实现对「当前已隐藏的行」用 `for r 到总行数` 全表扫描：百万行大表
            // 上每次点「确定」/「仅筛选此项」都要扫一遍，是筛选卡顿的主因。
            // 现在改为遍历 {@link DataCollection.getHiddenRowsSet} 倒排索引
            // （已隐藏行通常只有几百到几千行），复杂度与总行数无关。
            const candidates = new Set<number>();
            this.data.getRowsWithText().forEach((r) => {
                if (r > anchorRow) candidates.add(r);
            });
            this.data.getHiddenRowsSet().forEach((r) => {
                if (r > anchorRow) candidates.add(r);
            });

            // 4) 目标状态与现状不一致的行：直接改 RowHeader.isHidden 标记
            //    （同步维护 _hiddenRowsSet），最后调一次 finalizeRowVisibility 收尾——
            //    避免分段调用 setRowsHidden 产生多次 recalcRowPositions + 多次广播/重绘。
            const sorted = Array.from(candidates).sort((a, b) => a - b);
            let minChanged = Infinity;
            let maxChanged = -Infinity;
            let hadHide = false;
            let hadUnhide = false;
            for (let i = 0; i < sorted.length; i++) {
                const r = sorted[i];
                const rh = this.data.rowHeaders.getAt(r - 1);
                if (!rh) continue;
                const shouldHide = targetHidden.has(r);
                if (rh.isHidden === shouldHide) continue;   // 目标与现状一致：跳过
                this.data.setRowHiddenFlag(r, shouldHide);
                if (shouldHide) hadHide = true;
                else hadUnhide = true;
                if (r < minChanged) minChanged = r;
                if (r > maxChanged) maxChanged = r;
            }
            // 有任何行需要变化时做一次收尾（recalc + clamp + 补足/收缩 + 广播）
            if (minChanged !== Infinity) {
                this.data.finalizeRowVisibility(
                    minChanged, maxChanged, hadHide, hadUnhide
                );
            }

            // 5) 按钮状态：本列筛选行以下全部行都命中本列条件（全选，无实际筛选效果）时
            // 恢复默认背景色，否则标为主题绿色。
            // 按内容级判定而非行数 —— 避免其它列条件隐藏的行被误判为本列筛选效果。
            //
            // 【性能约定，勿回退为「for r = anchorRow+1..upper」】
            // 历史实现逐行扫到 upper（百万级大表的卡顿源）。现在改为遍历
            // {@link DataCollection.getRowsWithText}（只含有内容的行，通常几千行），
            // 空白行数按「应判定的行数 − 有内容行数」推算，不逐行扫。
            const maxRow = this.data.getContentBounds().maxRow;
            const totalRows = this.data.rowHeaders.length;
            const upper = Math.min(totalRows, Math.max(anchorRow, maxRow));
            // 应被判定的行数 = upper - anchorRow（筛选行以下到内容末尾/行头末尾）
            const contentRows = Math.max(0, upper - anchorRow);
            let matchedRows = 0;
            // 只遍历有内容的行（O(有内容行数)，与总行数无关）
            let countedContent = 0;
            this.data.getRowsWithText().forEach((r) => {
                if (r <= anchorRow || r > upper) return;
                countedContent++;
                const cell = this.data.getEffectiveCellAt(r, col);
                if (cell !== undefined
                    && this.data.getCellColAndRow(cell.cell).row <= anchorRow) {
                    // 跨越本列筛选行的合并块覆盖的行不受条件约束，视为命中
                    matchedRows++;
                    return;
                }
                const text = cell === undefined ? '' : cell.getText();
                if (keepContents.has(text)) matchedRows++;
            });
            // 空白行数 = 应判定的行数 − 有内容的行数
            const blankRows = Math.max(0, contentRows - countedContent);
            if (blankRows > 0 && keepContents.has('')) matchedRows += blankRows;
            // 内容范围之后的行（行头已补足但无内容）一律为空白
            const rest = Math.max(0, totalRows - upper);
            if (rest > 0 && keepContents.has('')) matchedRows += rest;
            this.setFilterButtonActive(matchedRows < contentRows + rest);
        });
    }

    /**
     * 设置筛选按钮的激活状态外观
     * @param {boolean} active - true=主题绿色背景（已筛选），false=恢复默认
     * @private
     * @returns {void}
     */
    private setFilterButtonActive(active: boolean): void {
        this.filter.style.backgroundColor = active ? Filter.ACTIVE_FILTER_BG : '';
    }
}
