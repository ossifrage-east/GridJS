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
    /** 导航器，负责页签面板与内容区的组织 */
    private navigator: Navigator;
    /** 筭选组件箭头元素 */
    private arrowElement: HTMLDivElement;
    /** 筛选组件容器，包含导航器面板与内容区 */
    private container: HTMLDivElement;
    /** 「按内容」页签的筛选结果提示区，展示当前列去重后的内容清单 */
    private tips: HTMLDivElement;
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
        const rowHeaders = this.data.rowHeaders;
        for (let r = cond.anchorRow + 1; r <= rowHeaders.length; r++) {
            if (rowHeaders.getAt(r - 1)?.isHidden) return;  // 仍有隐藏行 → 筛选仍生效，保持绿色
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
     * 构造时即完成 DOM 创建与挂载：生成筛选根节点追加到 options.parentElement，
     * 随后调用 {@link init} 初始化内部结构。
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
        this.init();
    }

    /**
     * 初始化筛选面板内部结构
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
    private async init(): Promise<void> {
        this.container = createDiv({
            className: ASSISTANT.FILTER_CONTAINER_CLASS_NAME,
        });
        this.navigator = new Navigator(true);
        const { panel, content } = this.navigator.getElements();
        this.container.append(panel, content);
        this.setArrow();
        // 等待页签与内容项创建完成，再挂载各页签内容
        await this.navigator.navItemMount('按内容', '按颜色');
        this.mountContent();
        this.setupEventListeners();
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
     * @param {Cell} cell - 筛选组件要定位到的单元格
     * @returns {void}
     */
    public setPosition(cell: Cell): void {
        // 记录本实例对应的筛选单元格，打开菜单时据此统计该列内容
        this.currentCellName = cell.cell;
        const { col, row } = this.data.getCellColAndRow(cell.cell);
        // 记录筛选按钮所在的行列号（1-based），筛选执行时使用
        this.currentCol = col;
        this.currentRow = row;
        const { left, top, width, height } = this.data.calculateCellRect(col, row);
        // 单元格右下角的屏幕坐标（内容坐标乘 zoom 后叠加固定行/列表头尺寸）
        const rect = this.filter.getBoundingClientRect();
        const _left = Math.round((left + width) * this.data.zoom) + this.data.visibleView.rowHeaderWidth - rect.width;
        const _top = Math.round((top + height) * this.data.zoom) + this.data.visibleView.colHeaderHeight - rect.height;
        this.filter.style.left = `${_left}px`;
        this.filter.style.top = `${_top}px`;
        // 定位时同步显隐：该列无内容的单元格不显示筛选按钮
        this.updateFilterVisibility();
    }

    /**
     * 更新筛选按钮的显示状态
     *
     * 统计本列筛选行以下是否存在非空内容（合并单元格覆盖位置取锚点文本，
     * colspan 跨列的锚点覆盖到本列也计入），有内容时显示按钮，
     * 全部为空时隐藏按钮并收起已打开的筛选菜单。
     * 单次遍历 values 判定（O(n)），供 VALUES_CHANGED 监听与 setPosition 定位时调用。
     * @private
     * @returns {void}
     */
    private updateFilterVisibility(): void {
        if (!this.currentCol) return;
        for (const cell of this.data.values) {
            const { col, row } = this.data.getCellColAndRow(cell.cell);
            const colspan = cell.colspan ?? 1;
            const rowspan = cell.rowspan ?? 1;
            // 单元格覆盖区域不落在本列，或全在筛选行及以上 → 不计入
            if (col > this.currentCol || col + colspan - 1 < this.currentCol) continue;
            if (row + rowspan - 1 <= this.currentRow) continue;
            if ((cell.chars?.map(c => c.char).join('') ?? '') !== '') {
                this.filter.style.display = '';
                return;
            }
        }
        // 该列筛选行以下无任何内容：隐藏筛选按钮（菜单打开时一并收起）
        this.filter.style.display = 'none';
        if (this.menu.hasMenu === this.filter) this.menu.closeMenu();
    }

    /**
     * 移除筛选组件实例
     * @returns {void}
     */
    public destroy(): void {
        this.filter.removeEventListener('click', this.boundHandleClick);
        this.data.off(DataEvents.VALUES_CHANGED, this.handleValuesChanged);
        this.filter.remove();
        this.navigator.destroy();
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
     * 注意：{@link Menu.closeMenu} 会清空菜单容器的所有子元素，
     * 因此每次打开前必须重新调用 {@link Menu.addMenuElement} 挂载容器；
     * openMenu 须传本组件根节点（箭头是其后代），箭头旋转与菜单位置定位才能生效。
     * @private
     * @param {MouseEvent} e - 点击事件
     * @returns {void}
     */
    private handleClick(e: MouseEvent): void {
        e.stopPropagation();
        // 已打开 → 切换为隐藏
        if (this.menu.hasMenu === this.filter) {
            this.menu.closeMenu();
            return;
        }
        // 未打开 → 先清理，再打开并重新挂载内容
        this.menu.closeMenu();
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
        // 优先使用本实例绑定的筛选单元格，缺省回退到当前激活单元格
        const cellName = this.currentCellName || this.data.activedCell;
        if (!cellName) return;
        const _cell = this.data.values.find((c) => c.cell === cellName);
        if (!_cell || _cell.filter === undefined) return;

        this.tips.textContent = '';
        const matrix = this.data.matrixValues();
        const { col, row: startRow } = this.data.getCellColAndRow(cellName);

        // 统计筛选单元格以下同列的内容（不含筛选单元格所在行，与 Excel 筛选行为一致）：
        // 空白单元格单独计数，非空白内容按出现顺序去重统计。
        // 本列已处于筛选状态 → 统计筛选前全部内容（不跳过隐藏行），并按当前条件设置勾选状态；
        // 其它列 → 只统计当前可见（非隐藏）行的内容，体现已生效筛选的叠加语义。
        // 合并单元格覆盖位置取锚点有效文本。
        // 注意：getCellColAndRow 返回 1-based，matrixValues 矩阵为 0-based，需减 1 对齐
        const activeKeep = this.data.getFilterCondition(col)?.keep ?? null;
        const anchorMap = this.buildAnchorMap();
        const countMap = new Map<string, number>();
        let blankCount = 0;
        for (let r = startRow; r < matrix.length; r++) {
            if (!activeKeep && this.data.rowHeaders.getAt(r)?.isHidden) continue;  // 非筛选列：已隐藏的行不计入清单
            const text = this.getEffectiveText(anchorMap, r + 1, col);
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
        // 空白行排在清单最前面（存在空白单元格时）
        if (blankCount > 0) {
            this.tips.append(createItemRow('空白', blankCount, isItemChecked('空白')));
        }
        // 非空白内容按 时间 → 数字 → 英文 → 中文 → 其它外语 的类型顺序升序排列后渲染
        const sortedEntries = Array.from(countMap.entries()).sort((a, b) => compareMixedText(a[0], b[0]));
        for (const [text, count] of sortedEntries) {
            this.tips.append(createItemRow(text, count, isItemChecked(text)));
        }
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
     * 构建「位置 → 合并锚点单元格」覆盖映射
     *
     * 遍历所有单元格，将合并锚点（colspan/rowspan > 1）覆盖的每个行列位置
     * 映射到锚点本身。未合并单元格也包含（覆盖自身位置），
     * 使清单统计与筛选匹配都能取到「视觉上属于该位置」的内容。
     * @returns {Map<string, Cell>} key 为 `${row}:${col}`（1-based 行列号），value 为覆盖该位置的单元格
     * @private
     */
    private buildAnchorMap(): Map<string, Cell> {
        const anchorMap = new Map<string, Cell>();
        for (const cell of this.data.values) {
            const { col, row } = this.data.getCellColAndRow(cell.cell);
            const colspan = cell.colspan ?? 1;
            const rowspan = cell.rowspan ?? 1;
            for (let r = row; r < row + rowspan; r++) {
                for (let c = col; c < col + colspan; c++) {
                    anchorMap.set(`${r}:${c}`, cell);
                }
            }
        }
        return anchorMap;
    }

    /**
     * 获取某位置的有效文本（考虑合并单元格覆盖）
     * @param {Map<string, Cell>} anchorMap - {@link buildAnchorMap} 构建的覆盖映射
     * @param {number} row - 行号（1-based）
     * @param {number} col - 列号（1-based）
     * @returns {string} 覆盖该位置的单元格文本，无覆盖时为空字符串
     * @private
     */
    private getEffectiveText(anchorMap: Map<string, Cell>, row: number, col: number): string {
        const cell = anchorMap.get(`${row}:${col}`);
        return cell?.chars?.map(c => c.char).join('') ?? '';
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
     * - 重放全部列条件：行对每个已筛选列的有效内容（合并锚点文本）都须命中该列保留集合；
     *   跨越某列筛选行的合并块覆盖的行不受该列条件约束（始终显示）；
     * - 完全空白的行保持显示（只隐藏有内容但未命中的行）；
     * - 对比行当前 isHidden 状态，将需要变化的连续行分段调用 setRowsHidden；
     * - 筛选完成后本列按钮背景按筛选效果标为主题绿色。
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
            const rowHeaders = this.data.rowHeaders;
            const totalRows = rowHeaders.length;

            // 1) 本列条件写入注册表（按操作顺序记录已筛选列）
            this.data.setFilterCondition(col, anchorRow, keepContents);

            // 2) 重放全部列条件，计算筛选后应显示的行
            const anchorMap = this.buildAnchorMap();
            const visibleRows = new Set<number>();
            for (let r = 1; r <= totalRows; r++) {
                let matched = true;
                for (const c of this.data.getFilterOrder()) {
                    const cond = this.data.getFilterCondition(c);
                    if (!cond || r <= cond.anchorRow) continue;  // 该列筛选行以上不受此条件约束
                    const cell = anchorMap.get(`${r}:${c}`);
                    if (cell) {
                        // 跨越该列筛选行的合并块覆盖的行强制显示（不受该列条件约束）
                        if (this.data.getCellColAndRow(cell.cell).row <= cond.anchorRow) continue;
                    }
                    const text = cell?.chars?.map(ch => ch.char).join('') ?? '';
                    if (!cond.keep.has(text)) {
                        matched = false;
                        break;
                    }
                }
                if (matched) visibleRows.add(r);
            }

            // 3) 有内容的行集合（1-based）：完全空白的行保持显示，只隐藏有内容但未命中的行
            const nonEmptyRows = new Set<number>();
            for (const cell of this.data.values) {
                if ((cell.chars?.map(ch => ch.char).join('') ?? '') === '') continue;
                const { row: r1 } = this.data.getCellColAndRow(cell.cell);
                for (let r = r1; r < r1 + (cell.rowspan ?? 1); r++) nonEmptyRows.add(r);
            }

            // 4) 与实质隐藏功能对齐：对比当前 isHidden 状态，把需要变化的连续行
            //    分段调用 setRowsHidden（统一入口自带 recalc/clamp/补足/收缩/广播）。
            //    遍历到 totalRows+1 作为哨兵，flush 最后一段。
            let segStart = -1;
            let segTarget = false;
            for (let r = anchorRow + 1; r <= totalRows + 1; r++) {
                const rh = r <= totalRows ? rowHeaders.getAt(r - 1) : null;
                const shouldHide = rh ? (!visibleRows.has(r) && nonEmptyRows.has(r)) : false;
                const needsChange = rh ? rh.isHidden !== shouldHide : false;
                if (needsChange) {
                    if (segStart < 0) {
                        segStart = r;
                        segTarget = shouldHide;
                    } else if (shouldHide !== segTarget) {
                        // 目标状态翻转：先 flush 前一段再开新段
                        this.data.setRowsHidden(segStart, r - 1, segTarget);
                        segStart = r;
                        segTarget = shouldHide;
                    }
                } else if (segStart >= 0) {
                    this.data.setRowsHidden(segStart, r - 1, segTarget);
                    segStart = -1;
                }
            }

            // 5) 按钮状态：本列筛选行以下全部行都命中本列条件（全选，无实际筛选效果）时
            // 恢复默认背景色，否则标为主题绿色。
            // 按内容级判定而非行数 —— 避免其它列条件隐藏的行被误判为本列筛选效果。
            let contentRows = 0;
            let matchedRows = 0;
            for (let r = anchorRow + 1; r <= totalRows; r++) {
                contentRows++;
                const cell = anchorMap.get(`${r}:${col}`);
                if (cell) {
                    // 跨越本列筛选行的合并块覆盖的行不受条件约束，视为命中
                    if (this.data.getCellColAndRow(cell.cell).row <= anchorRow) {
                        matchedRows++;
                        continue;
                    }
                }
                const text = cell?.chars?.map(ch => ch.char).join('') ?? '';
                if (keepContents.has(text)) matchedRows++;
            }
            this.setFilterButtonActive(matchedRows < contentRows);
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
