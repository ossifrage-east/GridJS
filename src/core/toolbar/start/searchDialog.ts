/**
 * 查找 / 替换对话框
 *
 * 查找与替换**合并为同一个对话框**，用 {@link Navigator} 的两个页签切换：
 * 「查找」页只有查找相关控件，「替换」页多一行「替换为」以及底部的
 * 「替换 / 全部替换」按钮（按钮按当前页签显隐）。
 *
 * ## 六条设计约定
 *
 * 1. **数据操作全在查找目标（`SearchTarget`）**：范围解析、命中扫描、字符级替换、跳转滚动
 *    都走 `target.resolveSearchRect / findAllInRect / replaceInRect / gotoSearchHit`，
 *    本文件只做「取值 + 展示 + 按钮编排」，不碰单元格字符，也**不 import 表格数据层**。
 *
 * 2. **只在选定区域内进行**：范围是对话框里的一等公民（可直接输入，也可在主表格上拖选），
 *    打开时取当前选区；此后你把活动单元格挪到哪儿都不会改变查询范围 ——
 *    这点是刻意的：`gotoSearchHit` **会改写选区**（为了醒目地跳到命中项），
 *    若把选区当范围，点一次「查找下一个」之后范围就塌缩成一个格子了。
 *
 * 3. **单格选区 = 整表 + 起点**：选区只有一个单元格时，范围被
 *    {@link DataCollection.expandSingleCellRect} 解释为**整表**，该格成为「起点」——
 *    一次查找从它往下走、到头回绕（见 {@link cursorBefore}）。
 *    范围输入框里显示的是**展开后的整表范围**，不是那个单格：否则用户点「全部替换」
 *    会以为只动了一格，实际改了一片。「起点」在用户手改范围时失效（见 rangeInput 监听）。
 *
 * 4. **不关闭对话框也能选区域（无按钮）**：对话框常驻 {@link Dialog.setRangePickMode}
 *    的点击穿透状态，鼠标事件直接落到主表格上，拖选即改范围。
 *    比「点按钮进入拾取态 → 选完再点一次退出」少一步操作，也不用记当前处在哪个态。
 *
 * 5. **一次只显示一个对话框**：由 {@link Dialog} 的静态登记表统一保证（打开别的
 *    对话框会自动隐去本对话框）；本对话框**隐藏时不销毁**，页签、关键字、匹配选项、
 *    命中游标全部保留，再次打开即可继续。
 *
 * 6. **每次打开复位居中**：复用同一实例，若不清除标题栏拖拽留下的位移，
 *    第二次打开就会停在旧位置。{@link Dialog.resetPosition} 负责复位。
 *
 * @module SearchDialog
 */

import { createDiv, createInput, createSpan } from "../../../utils/dom";
import { Dialog, DialogButtonOptions } from "../../common/dialog";
import { Navigator } from "../../common/navigator";
import { DataEvents } from "../../constant";
import { SearchHit, SearchMatchOptions, SearchRect, SearchTarget } from "../../contract/search";

/** 查找 / 替换的目标页签 */
export type SearchDialogMode = 'find' | 'replace';

/** SearchDialog 构造选项 */
export interface SearchDialogOptions {
    /** 查找目标（契约层窄接口；表格侧由 `DataCollection` 结构化满足） */
    target: SearchTarget;
}

/** 「查找」页签文本 */
const TAB_FIND = '查找';
/** 「替换」页签文本 */
const TAB_REPLACE = '替换';
/** 每行控件的标签宽度（像素） */
const LABEL_WIDTH = 68;
/**
 * 命中列表上限
 *
 * 命中列表要长期持有（「查找下一个」沿着它在范围里来回走），因此不能对
 * 「整表范围内搜一个极常见的字」这种情形无限膨胀 —— 几十万条 `SearchHit`
 * 对象既占内存、也会让排序变慢。超出上限时只列前 N 条，并在提示里说明，
 * 而不是假装总数就是 N（用户据此知道该缩小范围）。
 */
const SEARCH_HIT_LIMIT = 20000;
/** 对话框附加样式表 id（只注入一次） */
const STYLE_ELEMENT_ID = 'gs-search-dialog-styles';

/**
 * 查找 / 替换对话框
 *
 * 实例内部只建一次 DOM（含 Navigator），此后 {@link open} / {@link hide} 复用同一实例，
 * 因此隐藏再打开不会丢失页签、关键字、匹配选项与命中位置。
 *
 * @example
 * // 打开并定位到「替换」页签
 * new SearchDialog({ target: data }).open('replace');
 */
export class SearchDialog {
    /** 查找目标（只认识契约层接口，不认识表格数据集合） */
    private target: SearchTarget;
    /** 对话框实例（创建后长期持有，隐藏不销毁） */
    private dialog: Dialog | null = null;
    /** 页签导航组件 */
    private navigator: Navigator | null = null;
    /** 内部结构是否已构建 */
    private built: boolean = false;
    /** 页签是否已挂载完成（navItemMount 是异步的，未完成前不能 selectItem） */
    private navReady: boolean = false;
    /** 已请求但尚未生效的目标页签（等 navItemMount 完成后补切） */
    private pendingTab: SearchDialogMode | null = null;
    /** 当前生效的页签（默认「查找」） */
    private activeTab: SearchDialogMode = 'find';

    /**
     * 两套页签控件
     *
     * 两个页签各持有一套完整控件（与 Excel 的「查找和替换」对话框一致）：
     * 页签切来切去不会丢输入。同一语义的控件用 {@link linkInputs} 双向镜像，
     * 因此在任一页签里改动都会同步到另一页签。
     */
    private rangeInputs: HTMLInputElement[] = [];
    private findInputs: HTMLInputElement[] = [];
    private replaceInputs: HTMLInputElement[] = [];
    private caseInputs: HTMLInputElement[] = [];
    private entireInputs: HTMLInputElement[] = [];
    private wildcardInputs: HTMLInputElement[] = [];
    private statusElements: HTMLElement[] = [];

    /** 当前关键字对应的命中列表（沿用 findAllInRect 的行优先序） */
    private hits: SearchHit[] = [];
    /** 命中游标（-1 = 尚未定位到任何命中） */
    private cursor: number = -1;
    /** 关键字 / 选项 / 范围变化后置位，下次查找前重算命中列表 */
    private dirty: boolean = true;
    /** 命中列表被 {@link SEARCH_HIT_LIMIT} 截断（提示语与「第 N 个」的写法都要区分） */
    private truncated: boolean = false;
    /**
     * 「起点」——单格选区被扩成整表时记住的那一格
     *
     * 用来把第一次「查找下一个 / 上一个」定位到「该格之后 / 之前」，而不是范围首尾
     * （见 {@link cursorBefore}）。用户手改范围输入框时清空：那时范围不再由那个单格推出。
     */
    private anchor: { col: number; row: number } | null = null;
    /**
     * 「跳转改写选区」抑制窗口
     *
     * `gotoSearchHit` 会写 `selection`，而 selection setter 走 `updateProperty`（`!==` 比较）
     * → 值确实变了就同步广播 `SELECTION_CHANGED`。若不加抑制，一次「查找下一个」就会让
     * 选区变化被当成「用户在表格上重新选区」，范围当场塌缩成命中的那一格 ——
     * 正是设计约定 2 要防的事。故跳转前后夹一个窗口，期间忽略选区回调。
     */
    private suppressRangeSync: boolean = false;

    /**
     * 主表格选区变化 → 回填范围（设计约定 3）
     *
     * `gotoSearchHit` 造成的选区变化已被 {@link suppressRangeSync} 排除，
     * 这里只处理「用户真的在主表格上拖选」。
     */
    private readonly handleSelectionChanged = (): void => {
        if (this.suppressRangeSync) return;
        if (!this.isVisible()) return;
        const selection = this.target.selection;
        // 选区被清空（如刚切换文件）：保留既有范围，不拿「有内容区域」去暗中替换它
        if (!selection) return;
        this.applySelectionRange(selection);
    };

    /**
     * @param {SearchDialogOptions} options - 构造选项
     */
    constructor(options: SearchDialogOptions) {
        this.target = options.target;
    }

    /**
     * 对话框当前是否处于打开状态
     * @returns {boolean} 打开返回 true
     */
    public isVisible(): boolean {
        return this.dialog !== null && this.dialog.isVisible();
    }

    /**
     * 绑定的查找目标
     * @returns {SearchTarget} 查找目标
     */
    public getTarget(): SearchTarget {
        return this.target;
    }

    /**
     * 对话框已打开时的聚焦入口：切页签 + 聚焦输入框，**不动范围**
     *
     * 与 {@link open} 的区别就是最后这一点：快捷键连按不该把用户刚在主表格上
     * 拖出来的范围冲掉。
     * @param {SearchDialogMode} [mode] - 目标页签，省略时沿用当前页签
     */
    public refocus(mode?: SearchDialogMode): void {
        if (!this.isVisible()) return;
        if (mode) this.switchTab(mode);
        this.focusFind();
    }

    /**
     * 打开对话框（已打开时只切换页签并聚焦，不重置范围）
     * @param {SearchDialogMode} [mode] - 目标页签，省略时沿用当前页签
     */
    public open(mode?: SearchDialogMode): void {
        this.ensureBuilt();
        const dialog = this.dialog;
        if (!dialog) return;

        // 复位到居中位置：复用实例时不清除拖拽位移就会停在旧位置（设计约定 5）
        dialog.resetPosition();
        dialog.show();
        // 常驻点击穿透：不关对话框就能直接在主表格上拖选区域（设计约定 3）
        dialog.setRangePickMode(true);

        // 重复注册会叠加回调，先 off 再 on 保证只有一个订阅
        this.target.off(DataEvents.SELECTION_CHANGED, this.handleSelectionChanged);
        this.target.on(DataEvents.SELECTION_CHANGED, this.handleSelectionChanged);

        // 每次打开都以「当前选区」为默认范围（设计约定 2 的默认语义；
        // 单格选区会被解释成「整表 + 起点」，见设计约定 3）
        this.applySelectionRange(this.target.selection);

        this.switchTab(mode ?? this.activeTab);
        this.focusFind();
    }

    /**
     * 用主表格的选区刷新范围（打开时、以及用户在主表格上重新选区时）
     *
     * 单格选区交给 {@link DataCollection.expandSingleCellRect} 解释为「整表 + 起点」：
     * - 输入框写**展开后的整表范围**（不是那个单格）—— 「全部替换」的可视范围必须与
     *   实际作用范围一致，否则用户以为只动了一格；
     * - `anchor` 记住那一格，使第一次「查找下一个」落在它之后、而不是范围首行。
     * @param {string | null | undefined} selection - 主表格当前选区（如 `'A1:C9'` / `'B2:B2'`）
     */
    private applySelectionRange(selection: string | null | undefined): void {
        const rect = selection ? this.target.resolveSearchRect(selection) : this.target.resolveSearchRect();
        if (!rect) {
            this.anchor = null;
            this.setStatus('没有可查询的范围（请先在表格里选中区域，或直接填写如 A1:C20）');
            return;
        }
        const expanded = this.target.expandSingleCellRect(rect);
        const effective = expanded ? expanded.rect : rect;
        const text = this.target.formatRect(effective);
        this.anchor = expanded ? expanded.anchor : null;
        this.setRangeValue(text);
        this.dirty = true;
        if (expanded) {
            const anchorName = this.target.getCellName(expanded.anchor.col, expanded.anchor.row);
            this.setStatus(`选中的是单个单元格 ${anchorName}，已扩为整表 ${text} 查找（从 ${anchorName} 往下）`);
        } else {
            this.setStatus(`范围 = ${text}；可直接在主表格上拖选区域`);
        }
    }

    /**
     * 隐藏对话框（不销毁，保留全部输入与命中状态）
     */
    public hide(): void {
        this.target.off(DataEvents.SELECTION_CHANGED, this.handleSelectionChanged);
        this.dialog?.hide();
    }

    /**
     * 销毁对话框与页签导航（仅在数据集合更换 / 页面卸载时使用）
     */
    public destroy(): void {
        this.target.off(DataEvents.SELECTION_CHANGED, this.handleSelectionChanged);
        this.hits = [];
        this.cursor = -1;
        this.truncated = false;
        this.anchor = null;
        this.built = false;
        this.navReady = false;
        this.pendingTab = null;
        const navigator = this.navigator;
        this.navigator = null;
        if (navigator) void navigator.destroy();
        const dialog = this.dialog;
        this.dialog = null;
        if (dialog) dialog.destroy();
        this.rangeInputs = [];
        this.findInputs = [];
        this.replaceInputs = [];
        this.caseInputs = [];
        this.entireInputs = [];
        this.wildcardInputs = [];
        this.statusElements = [];
    }

    // ─────────────────────────── 构建 ───────────────────────────

    /**
     * 构建内部结构（幂等）
     *
     * 分三步：先同步建好两页控件（此刻尚未入 DOM），再建 Navigator 与 Dialog，
     * 最后等 `navItemMount`（内部要读 IndexedDB，是异步的）完成后把控件挂进
     * 对应页签 —— 这一步必须等，否则 `navContentItemMount` 找不到目标元素。
     */
    private ensureBuilt(): void {
        if (this.built) return;
        this.built = true;
        this.injectStyles();

        const findBody = this.buildPage(true);
        const replaceBody = this.buildPage(false);

        // 两页都建好后再做镜像绑定：每个语义组只挂一轮监听
        // （在 buildPage 里按页绑定会让第一页的控件被重复挂上监听）
        linkInputs(this.rangeInputs);
        linkInputs(this.findInputs);
        linkInputs(this.replaceInputs);
        linkCheckboxes(this.caseInputs);
        linkCheckboxes(this.entireInputs);
        linkCheckboxes(this.wildcardInputs);

        const navigator = new Navigator({
            horizontal: true,
            onItemShow: (text: string) => this.applyTabVisibility(text),
        });
        this.navigator = navigator;
        const { panel, content } = navigator.getElements();
        const root = createDiv({
            className: 'gs-search-dialog',
            children: [panel, content],
        });

        const buttons: DialogButtonOptions[] = [
            { text: '查找下一个', variant: 'primary', close: false, onClick: () => this.findNext() },
            { text: '查找上一个', variant: 'secondary', close: false, onClick: () => this.findPrev() },
            { text: '替换', variant: 'warning', close: false, onClick: () => this.replaceCurrent() },
            { text: '全部替换', variant: 'danger', close: false, onClick: () => this.replaceAll() },
            { text: '关闭', variant: 'light' },
        ];

        const dialog = new Dialog({
            title: '查找和替换',
            content: root,
            // 尺寸取 lg：两页控件较宽，md（500px）下「通配符 ?/* 」那组会换行
            size: 'lg',
            // 默认居中显示（需求 3）
            centered: true,
            buttons,
            // 隐藏（含被其它对话框顶掉）时只摘监听，不销毁 —— 保留状态供下次打开复用
            onClose: () => this.onDialogHide(),
        });
        this.dialog = dialog;

        void navigator.navItemMount(TAB_FIND, TAB_REPLACE).then(() => {
            navigator.navContentItemMount(TAB_FIND, findBody);
            navigator.navContentItemMount(TAB_REPLACE, replaceBody);
            this.navReady = true;
            const tab = this.pendingTab ?? this.activeTab;
            this.pendingTab = null;
            this.switchTab(tab);
        });
    }

    /**
     * 构建某一页的正文
     * @param {boolean} isFind - true 建「查找」页，false 建「替换」页
     * @returns {HTMLElement} 该页正文根元素
     */
    private buildPage(isFind: boolean): HTMLElement {
        const inputStyle = {
            flex: '1',
            padding: '0.25rem 0.375rem',
            fontSize: '0.875rem',
            lineHeight: '1.5',
            border: '1px solid #ced4da',
            borderRadius: '0.375rem',
            outline: 'none'
        };

        const rangeInput = createInput({
            type: 'text',
            style: inputStyle,
            attributes: { placeholder: '例如 A1:C20' }
        });
        // 手工改动范围后，命中列表必须重算（先前算好的行优先序已经不属于这个范围），
        // 并且「起点」失效 —— 范围不再由主表格上那个单格推出，锚在哪儿都成了臆断。
        // （顺带：手填的范围按字面处理，写 `B2:B2` 就是只查那一格，不会被扩成整表。）
        rangeInput.addEventListener('input', () => {
            this.anchor = null;
            this.dirty = true;
            this.setStatus('范围已修改，请重新查找');
        });

        const findInput = createInput({ type: 'text', style: inputStyle });
        findInput.addEventListener('input', () => { this.dirty = true; });
        findInput.addEventListener('keydown', (e) => {
            if ((e as KeyboardEvent).key === 'Enter') {
                e.preventDefault();
                this.findNext();
            }
        });

        const replaceInput = createInput({ type: 'text', style: inputStyle });
        replaceInput.addEventListener('keydown', (e) => {
            if ((e as KeyboardEvent).key === 'Enter') {
                e.preventDefault();
                this.replaceCurrent();
            }
        });

        const statusElement = createDiv({
            style: {
                minHeight: '1.25rem',
                fontSize: '0.8125rem',
                color: '#6c757d',
                userSelect: 'text'
            }
        });

        const rows: HTMLElement[] = [];
        rows.push(this.buildRow('范围', rangeInput));
        rows.push(this.buildRow('查找内容', findInput));
        if (!isFind) rows.push(this.buildRow('替换为', replaceInput));
        rows.push(this.buildOptionsRow());
        rows.push(statusElement);

        // 登记控件：同语义的输入框跨页签双向镜像（见 ensureBuilt 的 linkInputs）
        this.rangeInputs.push(rangeInput);
        this.findInputs.push(findInput);
        if (!isFind) this.replaceInputs.push(replaceInput);
        this.statusElements.push(statusElement);

        return createDiv({
            children: rows,
            style: { display: 'flex', flexDirection: 'column' }
        });
    }

    /**
     * 构建「标签 + 控件」的一行
     * @param {string} labelText - 标签文字
     * @param {HTMLElement} control - 控件元素
     * @returns {HTMLElement} 行元素
     */
    private buildRow(labelText: string, control: HTMLElement): HTMLElement {
        const label = createSpan({
            textContent: labelText,
            style: { flex: `0 0 ${LABEL_WIDTH}px`, color: '#495057' }
        });
        return createDiv({
            children: [label, control],
            style: { display: 'flex', alignItems: 'center', gap: '0.5rem', marginBottom: '0.5rem' }
        });
    }

    /**
     * 构建匹配选项行（区分大小写 / 完全匹配 / 通配符）
     * @returns {HTMLElement} 选项行元素
     */
    private buildOptionsRow(): HTMLElement {
        const makeCheckbox = (text: string, sink: HTMLInputElement[]): HTMLElement => {
            const input = createInput({ type: 'checkbox' });
            input.addEventListener('change', () => { this.dirty = true; });
            sink.push(input);
            return createDiv({
                children: [input, createSpan({ textContent: text })],
                style: { display: 'inline-flex', alignItems: 'center', gap: '0.25rem' }
            });
        };
        const wraps = [
            makeCheckbox('区分大小写', this.caseInputs),
            makeCheckbox('完全匹配', this.entireInputs),
            makeCheckbox('通配符 ?/* ', this.wildcardInputs),
        ];
        return createDiv({
            children: wraps,
            style: {
                display: 'flex',
                flexWrap: 'wrap',
                gap: '0.75rem',
                marginLeft: `${LABEL_WIDTH + 8}px`,
                marginBottom: '0.5rem'
            }
        });
    }

    /**
     * 注入对话框附加样式（只注入一次，与 Dialog 自身的样式表隔离）
     *
     * Navigator 的默认样式是「内容区撑满容器高度」（`height: 100%`），放进
     * 自适应高度的对话框正文里会算出 0 高、内容被压扁，因此在这里把
     * 面板改成页签条、内容区改成随内容自适应。
     */
    private injectStyles(): void {
        if (document.getElementById(STYLE_ELEMENT_ID)) return;
        const styleSheet = document.createElement('style');
        styleSheet.id = STYLE_ELEMENT_ID;
        styleSheet.textContent = `
            /* 页签条：贴住正文顶部，横排、带下边框 */
            .gs-search-dialog .nav-panel {
                display: flex;
                flex-direction: row;
                gap: 6px;
                padding: 0 0.25rem;
                margin-bottom: 0.5rem;
                border-bottom: 1px solid #dee2e6;
            }
            /* 内容区：默认的 height:100% 在自适应高度的对话框里会塌成 0，改为随内容 */
            .gs-search-dialog .nav-content {
                display: block;
                width: auto;
                height: auto;
            }
            .gs-search-dialog .nav-content .content-item {
                width: auto;
                height: auto;
                animation: none;
            }
            /* 页签项：加大点击区，选中态用主色下边框（覆盖 nav.scss 的默认绿） */
            .gs-search-dialog .panel-item.horizontal {
                padding: 0.25rem 0.75rem;
                font-size: 0.875rem;
            }
            .gs-search-dialog .panel-item.horizontal.checked {
                border-bottom-color: #0d6efd;
                background-color: transparent;
            }
            .gs-search-dialog .panel-item.horizontal:not(.checked):hover {
                border-bottom-color: #0d6efd;
            }
        `;
        document.head.appendChild(styleSheet);
    }

    // ─────────────────────────── 页签 ───────────────────────────

    /**
     * 切换到指定页签
     * @param {SearchDialogMode} mode - 目标页签
     */
    private switchTab(mode: SearchDialogMode): void {
        this.activeTab = mode;
        this.applyTabVisibility(mode === 'replace' ? TAB_REPLACE : TAB_FIND);
        if (!this.navReady) {
            // 页签 DOM 还没建好，记下来等 navItemMount 完成后补切
            this.pendingTab = mode;
            return;
        }
        void this.navigator?.selectItem(mode === 'replace' ? TAB_REPLACE : TAB_FIND);
    }

    /**
     * 按当前页签显隐「替换 / 全部替换」按钮
     * @param {string} navItemText - 当前显示的导航项文本
     */
    private applyTabVisibility(navItemText: string): void {
        const isReplace = navItemText === TAB_REPLACE;
        this.activeTab = isReplace ? 'replace' : 'find';
        const dialog = this.dialog;
        if (!dialog) return;
        for (const button of dialog.getFooterButtons()) {
            const text = button.textContent ?? '';
            if (text === '替换' || text === '全部替换') {
                button.hidden = !isReplace;
            }
        }
    }

    /**
     * 聚焦当前页签的「查找内容」输入框并选中已输入内容
     */
    private focusFind(): void {
        const input = this.findInputs[this.activeTab === 'replace' ? 1 : 0] ?? this.findInputs[0];
        if (!input) return;
        input.focus();
        input.select();
    }

    // ─────────────────────────── 查询 / 替换 ───────────────────────────

    /**
     * 当前匹配选项（每次读取，避免持有快照）
     * @returns {SearchMatchOptions} 匹配选项
     */
    private currentOptions(): SearchMatchOptions {
        return {
            matchCase: this.caseInputs[0]?.checked === true,
            matchEntire: this.entireInputs[0]?.checked === true,
            useWildcard: this.wildcardInputs[0]?.checked === true
        };
    }

    /**
     * 解析范围输入框；非法时给出提示
     * @returns {SearchRect | null} 合法范围
     */
    private currentRect(): SearchRect | null {
        const rect = this.target.resolveSearchRect(this.rangeInputs[0]?.value ?? '');
        if (!rect) {
            this.setStatus('没有可查询的范围（请先在表格里选中区域，或直接填写如 A1:C20）');
            return null;
        }
        return rect;
    }

    /**
     * 当前关键字
     * @returns {string} 查找内容输入框的值
     */
    private currentKeyword(): string {
        return this.findInputs[0]?.value ?? '';
    }

    /**
     * 重算命中列表（关键字/选项/范围任一变化后调用）
     *
     * 结果被 {@link SEARCH_HIT_LIMIT} 截断时置 {@link truncated}：列表仍可正常导航，
     * 只是条目不全，提示语会说明（不谎报总数）。
     * @param {SearchRect} rect - 查询范围
     * @param {string} keyword - 关键字
     * @param {{ col: number; row: number } | null} [resumeFrom] - 起点（见 {@link cursorBefore}）
     */
    private refreshHits(
        rect: SearchRect,
        keyword: string,
        resumeFrom: { col: number; row: number } | null = null
    ): void {
        this.hits = this.target.findAllInRect(rect, keyword, this.currentOptions(), SEARCH_HIT_LIMIT);
        this.truncated = this.hits.length >= SEARCH_HIT_LIMIT;
        this.cursor = this.cursorBefore(rect, resumeFrom);
        this.dirty = false;
    }

    /**
     * 把「起点位置」换算成命中列表里的游标位置
     *
     * 约定：返回值指向「起点处或起点之前**最后一个**命中」（`-1` = 一个都没有）。
     * 这样两个方向由同一个游标表达，不需要维护两份状态：
     * - 「查找下一个」的 `cursor + 1` 恰好落在**起点之后的第一条**；
     * - 「查找上一个」的 `cursor - 1` 恰好落在**起点之前的最后一条**。
     *
     * 两处调用：
     * - 单格选区扩成整表时传 `anchor`（「从这个单元格往下找」）；
     * - 「替换」之后传刚改过的那一格（它已从列表里消失，于是下一个命中就是原后继）。
     * @param {SearchRect} rect - 查询范围
     * @param {{ col: number; row: number } | null} position - 起点；null 表示从头开始
     * @returns {number} 游标位置（-1 表示从头开始）
     */
    private cursorBefore(rect: SearchRect, position: { col: number; row: number } | null): number {
        if (!position) return -1;
        const width = rect.endCol - rect.startCol + 1;
        const target = searchKey(rect, position.row, position.col, width);
        for (let i = 0; i < this.hits.length; i++) {
            const hit = this.hits[i];
            if (searchKey(rect, hit.row, hit.col, width) > target) return i - 1;
        }
        return this.hits.length - 1;
    }

    /**
     * 命中进度的提示语
     * @param {SearchHit} hit - 当前命中项
     * @returns {string} 如 `第 3 / 12 个：B5`
     */
    private progressStatus(hit: SearchHit): string {
        return this.truncated
            ? `第 ${this.cursor + 1} 个：${hit.cell}（命中超过 ${SEARCH_HIT_LIMIT} 处，仅列出前 ${SEARCH_HIT_LIMIT} 处）`
            : `第 ${this.cursor + 1} / ${this.hits.length} 个：${hit.cell}`;
    }

    /**
     * 查找下一个（到末尾后回绕到首个命中）
     */
    private findNext(): void {
        const rect = this.currentRect();
        if (!rect) return;
        const keyword = this.currentKeyword();
        if (!keyword) {
            this.setStatus('请输入查找内容');
            return;
        }
        if (this.dirty) this.refreshHits(rect, keyword, this.anchor);
        if (this.hits.length === 0) {
            this.setStatus('没有找到匹配的内容');
            return;
        }
        // 同一命中列表内游标只前进一格
        if (this.cursor + 1 >= this.hits.length) {
            this.cursor = 0;
        } else {
            this.cursor += 1;
        }
        const hit = this.hits[this.cursor];
        this.gotoHit(hit);
        this.setStatus(this.progressStatus(hit));
    }

    /**
     * 查找上一个（到首个后回绕到末尾）
     */
    private findPrev(): void {
        const rect = this.currentRect();
        if (!rect) return;
        const keyword = this.currentKeyword();
        if (!keyword) {
            this.setStatus('请输入查找内容');
            return;
        }
        if (this.dirty) this.refreshHits(rect, keyword, this.anchor);
        if (this.hits.length === 0) {
            this.setStatus('没有找到匹配的内容');
            return;
        }
        this.cursor = this.cursor - 1 < 0 ? this.hits.length - 1 : this.cursor - 1;
        const hit = this.hits[this.cursor];
        this.gotoHit(hit);
        this.setStatus(this.progressStatus(hit));
    }

    /**
     * 跳转到命中项
     *
     * 跳转会把选区收敛成单格并广播 `SELECTION_CHANGED`，必须在窗口内屏蔽
     * {@link handleSelectionChanged}，否则范围会被这次跳转改成单格。
     * @param {SearchHit} hit - 命中项
     */
    private gotoHit(hit: SearchHit): void {
        this.suppressRangeSync = true;
        try {
            this.target.gotoSearchHit(hit);
        } finally {
            this.suppressRangeSync = false;
        }
    }

    /**
     * 替换当前命中项，并自动跳到下一个命中项（与 Excel 的「替换」一致）
     */
    private replaceCurrent(): void {
        const rect = this.currentRect();
        if (!rect) return;
        const keyword = this.currentKeyword();
        if (!keyword) {
            this.setStatus('请输入查找内容');
            return;
        }
        if (this.dirty) this.refreshHits(rect, keyword, this.anchor);
        if (this.hits.length === 0) {
            this.setStatus('没有找到匹配的内容');
            return;
        }
        if (this.cursor < 0 || this.cursor >= this.hits.length) {
            this.setStatus('请先点击「查找下一个」定位到某个单元格');
            return;
        }
        const hit = this.hits[this.cursor];
        // 只替换这一格：把命中项作为 targets 传入，范围内其它命中保持原样
        const count = this.target.replaceInRect(
            rect, keyword, this.replaceInputs[0]?.value ?? '', this.currentOptions(), [hit], '替换'
        );
        // 【只就地更新这一条，**不重扫**整个范围】
        // 「替换」是连点操作，重扫在整表尺度上每次点击都要付一次全表扫描的代价
        // （几十万格 = 几百毫秒），而本次只改动了这一个格子 —— 其余格子的文本
        // 一个字都没变，列表里的 `text` 快照仍然有效。
        // 替换后是否仍命中用 cellMatches 现算一次（关键字 'a' → 替换成 'aa' 的情形
        // 仍然命中，那就把它留在列表里，下次「查找下一个」走过它、再下一轮回到它）。
        const rebuilt = this.target.cellMatches(hit.cell, keyword, this.currentOptions());
        if (rebuilt) {
            this.hits[this.cursor] = rebuilt;
        } else {
            this.hits.splice(this.cursor, 1);
            // 摘掉当前项后，原本的后继挪到了同一个下标上；游标退一格，
            // 于是下面 findNext 的 `cursor + 1` 正好落回它（而不是跳过一格）。
            this.cursor -= 1;
        }
        if (this.hits.length === 0) {
            this.setStatus(`已替换 ${count} 处（${hit.cell}），范围内再无匹配`);
            return;
        }
        this.findNext();
    }

    /**
     * 替换范围内的全部命中项
     */
    private replaceAll(): void {
        const rect = this.currentRect();
        if (!rect) return;
        const keyword = this.currentKeyword();
        if (!keyword) {
            this.setStatus('请输入查找内容');
            return;
        }
        // 不传 targets = 替换范围内全部命中（数据层一次稀疏扫描 + 一次稀疏补丁，
        // 不经命中列表，因此不受 SEARCH_HIT_LIMIT 影响）
        const count = this.target.replaceInRect(
            rect, keyword, this.replaceInputs[0]?.value ?? '', this.currentOptions(), null, '全部替换'
        );
        this.dirty = true;
        this.hits = [];
        this.cursor = -1;
        this.truncated = false;
        this.setStatus(count > 0 ? `已替换 ${count} 处` : '没有找到匹配的内容');
    }

    // ─────────────────────────── 辅助 ───────────────────────────

    /**
     * 写入范围（两页签的输入框一起写）
     * @param {string} value - 范围串
     */
    private setRangeValue(value: string): void {
        for (const input of this.rangeInputs) input.value = value;
    }

    /**
     * 设置提示文本（两页签的提示一起写）
     * @param {string} text - 提示文本
     */
    private setStatus(text: string): void {
        for (const element of this.statusElements) element.textContent = text;
    }

    /**
     * 对话框隐藏后的轻量清理：只摘监听，**不销毁**
     *
     * 「不销毁」是刻意的：被其它对话框顶掉、或用户按 ESC 关掉之后再次打开，
     * 页签、关键字、匹配选项与命中位置都还在（设计约定 4）。
     */
    private onDialogHide(): void {
        this.target.off(DataEvents.SELECTION_CHANGED, this.handleSelectionChanged);
    }
}

/**
 * 单元格在查询范围内的**行优先序号**（与 DataCollection 内部的键同构）
 * @param {SearchRect} rect - 查询范围
 * @param {number} row - 行号（1-based）
 * @param {number} col - 列号（1-based）
 * @param {number} width - 范围宽度（列数）
 * @returns {number} 行优先序号
 */
function searchKey(rect: SearchRect, row: number, col: number, width: number): number {
    return (row - rect.startRow) * width + (col - rect.startCol);
}

/**
 * 把一组输入框的 value 双向镜像
 *
 * 两个页签各有一套控件，用户在任一页签输入都应同步到另一页签，
 * 否则「在查找页输入关键字 → 切到替换页」会发现关键字没了。
 * 加守卫（值相同不写）避免 A 写 B、B 又回写 A 的循环。
 * @param {HTMLInputElement[]} inputs - 同一语义的输入框
 */
function linkInputs(inputs: HTMLInputElement[]): void {
    for (const input of inputs) {
        input.addEventListener('input', () => {
            for (const other of inputs) {
                if (other !== input && other.value !== input.value) other.value = input.value;
            }
        });
    }
}

/**
 * 把一组复选框的勾选状态双向镜像（与 {@link linkInputs} 同理）
 * @param {HTMLInputElement[]} inputs - 同一语义的复选框
 */
function linkCheckboxes(inputs: HTMLInputElement[]): void {
    for (const input of inputs) {
        input.addEventListener('change', () => {
            for (const other of inputs) {
                if (other !== input) other.checked = input.checked;
            }
        });
    }
}

// ─────────────────────────── 全局共享实例 ───────────────────────────

/**
 * 全局唯一的查找 / 替换对话框
 *
 * 工具栏菜单与 Ctrl+F / Ctrl+H 快捷键都要打开**同一个**对话框，因此需要共享实例：
 * 各自 `new` 一个会出现两个互不感知的对话框（快捷键打开时菜单那个不隐藏、
 * 也不会被 {@link Dialog} 的互斥显隐覆盖到，因为它同样处于打开状态）。
 * 由 `openFindReplace` / `hideFindReplace` 统一管理；查找目标更换时自动重建。
 */
let sharedSearchDialog: SearchDialog | null = null;

/**
 * 打开查找 / 替换对话框
 *
 * 已打开时**只切换页签并聚焦**（不重置范围）—— 快捷键连按不该把用户
 * 辛苦拖出来的范围冲掉；需要收起请按 ESC 或点「关闭」。
 * @param {SearchTarget} target - 查找目标（表格侧传 `DataCollection`，结构兼容）
 * @param {SearchDialogMode} [mode] - 目标页签，省略时沿用当前页签
 */
export function openFindReplace(target: SearchTarget, mode?: SearchDialogMode): void {
    if (sharedSearchDialog && sharedSearchDialog.getTarget() !== target) {
        sharedSearchDialog.destroy();
        sharedSearchDialog = null;
    }
    if (!sharedSearchDialog) sharedSearchDialog = new SearchDialog({ target });
    if (sharedSearchDialog.isVisible()) {
        sharedSearchDialog.refocus(mode);
        return;
    }
    sharedSearchDialog.open(mode);
}

/**
 * 隐藏查找 / 替换对话框（保留状态）
 */
export function hideFindReplace(): void {
    sharedSearchDialog?.hide();
}

/**
 * 当前共享对话框实例（供测试与调试）
 * @returns {SearchDialog | null} 实例
 */
export function getSharedSearchDialog(): SearchDialog | null {
    return sharedSearchDialog;
}
