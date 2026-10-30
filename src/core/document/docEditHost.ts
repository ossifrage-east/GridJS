/*  # 文档域 · 文档编辑宿主 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目目标: 让**同一套工具栏**既能编辑表格，也能编辑文档 —— 二者只是在
    #           {@link EditHost} 背后换了实现。
    # 本文件职责: 把 23 条编辑命令翻译成 {@link DocView} / {@link DocStore} 的调用；
    #             把文档里的选区状态归一化成 {@link EditStateSnapshot} 供按钮回读。
    #             这是**唯一**允许「命令名 ↔ 文档 API」互相翻译的地方。
    # 依赖方向: contract/editHost（契约）+ 同目录 docView/docStore。
    #           **不 import 表格数据层** —— 文档宿主不该认识 Cell / DataCollection。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { EventEmitter } from '../../utils/eventEmitter';
import { DataEvents } from '../constant';
import {
    BrushMode,
    CharStyleState,
    CommandId,
    CommandPayload,
    DocListKind,
    EDIT_STATE_CHANGED,
    EditHost,
    EditStateSnapshot,
    EditingStateSource,
    ParaStyleId,
    ParaStyleState,
    TextAlignValue,
    WorkspaceKind,
} from '../contract/editHost';
import { DEFAULT_DOC_FONT_SIZE } from '../constant';
import { CharStylePatch, DocBlock } from './blockTree';
import { DocStore } from './docStore';
import { DocView } from './docView';

/**
 * 文档宿主能执行的命令
 *
 * 字符级 + 段落级 + 段落样式 + 剪贴板 + 历史；表格专属（边框 / 合并 / 筛选 / 排序 /
 * 格式刷 / 垂直对齐 / 自动换行）一律不在此列 —— {@link DocEditHost.canExecute} 对未列出的
 * 命令返回 false，工具栏据此把按钮置灰，而不需要为文档再写一套按钮。
 */
const SUPPORTED: CommandId[] = [
    'setFontFamily',
    'setFontSize',
    'toggleBold',
    'toggleItalic',
    'toggleUnderline',
    'toggleStrike',
    'setTextColor',
    'setBackground',
    'setAlign',
    'indentStep',
    'setParaStyle',
    'toggleList',
    'copy',
    'cut',
    'paste',
    'pasteValue',
    'pasteStyle',
    'brushClick',
    'brushDoubleClick',
    'undo',
    'redo',
];

/**
 * 块类型 + 级别 → 段落样式标识
 * @param type - 块类型
 * @param level - 标题级别
 * @returns 段落样式标识
 */
function toParaStyleId(type: DocBlock['type'], level: number): ParaStyleId {
    if (type === 'heading') {
        const clamped = Math.max(1, Math.min(6, level || 1));
        return `h${clamped}` as ParaStyleId;
    }
    if (type === 'quote') return 'quote';
    if (type === 'code') return 'code';
    return 'paragraph';
}

/**
 * 段落样式标识 → 块类型与级别
 * @param id - 段落样式标识
 * @returns 块类型与级别
 */
function fromParaStyleId(id: ParaStyleId): { type: DocBlock['type']; level: number } {
    if (id === 'quote') return { type: 'quote', level: 1 };
    if (id === 'code') return { type: 'code', level: 1 };
    if (id !== 'paragraph' && id.charCodeAt(0) === 104 /* h */) {
        const level = parseInt(id.slice(1), 10);
        return { type: 'heading', level: Number.isFinite(level) ? level : 1 };
    }
    return { type: 'paragraph', level: 1 };
}

/**
 * 设备像素 → CSS 像素
 *
 * 工具栏的字号菜单发的是「磅值 × devicePixelRatio」（见 `sectionFont` 的换算），
 * 而文档模型里存的是 CSS 像素。这里做一次除法还原，使「选 16」真的得到 16px。
 * @param value - 工具栏发来的字号
 * @returns CSS 像素
 */
function toCssFontSize(value: number): number {
    const dpr = window.devicePixelRatio || 1;
    const css = value / dpr;
    if (!Number.isFinite(css) || css <= 0) return DEFAULT_DOC_FONT_SIZE;
    return Math.max(6, Math.min(96, Math.round(css)));
}

/**
 * CSS 像素 → 工具栏字号（{@link toCssFontSize} 的逆运算）
 * @param size - CSS 像素
 * @returns 工具栏显示值
 */
function toToolbarFontSize(size: number): number {
    const dpr = window.devicePixelRatio || 1;
    return Math.round(size * dpr);
}

/**
 * 文档编辑宿主
 */
export class DocEditHost extends EventEmitter implements EditHost {
    /** 宿主类型 */
    public readonly kind: WorkspaceKind = 'doc';

    /** 状态事件源（自身；`EventEmitter` 已占用 `events` 与 `listeners` 两个名字） */
    public get stateSource(): EditingStateSource {
        return this;
    }

    /** 视图 */
    private readonly view: DocView;

    /** 存储 */
    private readonly store: DocStore;

    /**
     * 上一次广播的聚焦状态
     *
     * `IS_EDITTING_CHANGED` 的载荷约定是 `(旧值, 新值)`（见 `EventEmitter.updateProperty`），
     * 按钮的 `isDisabledOnEditting` 特性读的是第二个参数。因此必须自己记住旧值，
     * 并且**只在真正变化时**发——每帧都发会让所有置灰按钮反复重算。
     */
    private lastEditing: boolean = false;

    /** 是否已有一帧待刷新的状态广播（见 {@link onCaretMove} 的按帧合并） */
    private refreshQueued: boolean = false;

    /**
     * 复制时记下的字符样式（供「粘贴格式」取用）
     *
     * 剪贴板 API 只让写文本，样式那一半只能自己留一份。它是**快照**而不是引用：
     * 复制后原段落再改样式，粘出去的仍是复制那一刻的样子（与 Word 一致）。
     */
    private copiedStyle: CharStylePatch | null = null;

    /**
     * 当前格式刷模式（`null` 表示没在刷）
     *
     * 与表格侧同一个 {@link BrushMode} 取值，按钮高亮由 {@link CommandBus.snapshot} 统一回读。
     */
    private brushModeValue: BrushMode = null;

    /**
     * 格式刷抄下的样式
     *
     * 【为什么是快照】Word 的格式刷刷的是「点下刷子那一刻的样子」：之后再改源文字，
     * 刷出去的仍是旧样式。若持有源字符的引用，样式会跟着变，与直觉相反。
     */
    private brushStyle: CharStylePatch | null = null;

    /**
     * 构造函数
     * @param view - 文档视图（负责 DOM 与选区）
     * @param store - 文档存储（负责块树与撤销）
     */
    constructor(view: DocView, store: DocStore) {
        super();
        this.view = view;
        this.store = store;
        this.listenCaret();
    }

    /**
     * 订阅「光标 / 选区变了」
     *
     * 【为什么是宿主的事，不是工作区的事】工具栏的段落样式下拉与列表按钮要跟着光标走，
     * 而「光标在哪」属于编辑语义 —— 由宿主统一转成 {@link EDIT_STATE_CHANGED} 广播，
     * 外壳只订阅总线一次，不必知道鼠标键盘事件发生在哪个 DOM 上。
     *
     * 只挂 `keyup` / `mouseup` / `selectionchange` 三个**粗粒度**事件：
     * - `selectionchange` 覆盖方向键、点击、拖选、程序化移动；
     * - `keyup` 兜住「焦点还在但选区没变」的输入（打字后样式按钮要跟着刷新）；
     * - `mouseup` 兜住点击表格外区域导致选区丢失的情形。
     * 事件选择只影响「刷新有多及时」，不影响正确性 —— 所有读取都走快照。
     * @private
     */
    private listenCaret(): void {
        const page = this.view.getElements().page;
        page.addEventListener('keyup', this.onCaretMove);
        page.addEventListener('mouseup', this.onCaretMove);
        document.addEventListener('selectionchange', this.onCaretMove);
        // 格式刷：拖选结束的那一刻才刷（见 {@link applyBrush} 的说明）
        document.addEventListener('mouseup', this.onBrushApply);
    }

    /**
     * 鼠标抬起 → 格式刷若在生效中，就把样式刷到当前选区
     *
     * 用字段而不是方法引用：{@link listenCaret} 注册后不需要解绑，
     * 但写成字段能避免以后有人顺手加 `off()` 时因为 `this` 绑定不同而漏退订。
     */
    private readonly onBrushApply = (): void => {
        if (!this.store.hasDoc) return;
        if (!this.brushModeValue) return;
        if (this.applyBrush()) this.notifyStateChanged();
    };

    /**
     * 光标可能移动了 → 一帧内只广播一次
     *
     * `selectionchange` 是高频事件（每一次按键、每一次方向键、每一次拖选都会发），
     * 而一次广播会让工具栏四个分区各回读一遍快照并写十几处 DOM 样式。
     * 这里按帧合并：同一帧里来多少次都只刷一次，观感不变、开销降到 1/N。
     *
     * 用箭头函数字段而不是方法：{@link listenCaret} 需要同一个引用来解绑，
     * 方法引用每次取到的虽然是同一个函数对象，但写成字段更不容易在重构中漏掉 `bind`。
     * @private
     */
    private readonly onCaretMove = (): void => {
        // 文档没打开时（工作区已切到表格）不广播，免得表格的状态被无关事件刷一遍
        if (!this.store.hasDoc) return;
        if (this.refreshQueued) return;
        this.refreshQueued = true;
        const flush = (): void => {
            this.refreshQueued = false;
            this.notifyStateChanged();
        };
        if (typeof requestAnimationFrame === 'function') requestAnimationFrame(flush);
        else window.setTimeout(flush, 0);
    };

    /**
     * 该命令在文档宿主下是否可执行
     * @param cmd - 命令名
     * @returns 是否可执行
     */
    public canExecute(cmd: CommandId): boolean {
        return SUPPORTED.indexOf(cmd) >= 0;
    }

    /**
     * 执行命令
     * @param cmd - 命令名
     * @param payload - 载荷
     */
    public execute(cmd: CommandId, payload?: CommandPayload): void {
        const patch = this.toPatch(cmd, payload);
        if (patch) {
            this.view.applyCharStyle(patch);
            this.notifyStateChanged();
            return;
        }
        switch (cmd) {
            case 'setAlign': {
                const align = payload?.value as TextAlignValue | undefined;
                if (align) this.view.setAlign(align);
                break;
            }
            case 'indentStep': {
                const step = typeof payload?.value === 'number' ? payload.value : 1;
                this.view.indentBy(step);
                break;
            }
            case 'setParaStyle': {
                const id = payload?.value as ParaStyleId | undefined;
                if (id) {
                    const { type, level } = fromParaStyleId(id);
                    this.view.setBlockType(type, level);
                }
                break;
            }
            case 'toggleList': {
                const kind = payload?.value as DocListKind | undefined;
                if (kind) this.view.toggleList(kind);
                break;
            }
            case 'undo':
                this.store.undo();
                break;
            case 'redo':
                this.store.redo();
                break;
            case 'copy':
                void this.copySelection(false);
                break;
            case 'cut':
                void this.copySelection(true);
                break;
            case 'paste':
                void this.pasteFromClipboard();
                break;
            case 'pasteValue':
                // 「粘贴文字」：剪贴板在文档里只有纯文本一个通道，与 `paste` 同源；
                // 保留两个分支是为了让菜单项的措辞与 Word 一致（外壳不再自己编命令）。
                void this.pasteFromClipboard();
                break;
            case 'pasteStyle':
                this.pasteCopiedStyle();
                break;
            case 'brushClick':
                // 单击 = 单次刷；已在连续模式时再点一次 = 退出（与表格同一个手势语义）
                this.startBrush(this.brushModeValue === 'continuous' ? null : 'single');
                break;
            case 'brushDoubleClick':
                this.startBrush(this.brushModeValue === 'continuous' ? null : 'continuous');
                break;
            default:
                // 不可执行的命令由 CommandBus 在派发前挡掉，这里再判一次只是防御
                break;
        }
        this.notifyStateChanged();
    }

    /**
     * 把字符级命令翻译成样式补丁
     * @param cmd - 命令名
     * @param payload - 载荷
     * @returns 补丁；非字符级命令返回 null
     */
    private toPatch(cmd: CommandId, payload?: CommandPayload): CharStylePatch | null {
        switch (cmd) {
            case 'setFontFamily':
                return typeof payload?.value === 'string' ? { fontFamily: payload.value } : null;
            case 'setFontSize':
                return typeof payload?.value === 'number' ? { fontSize: toCssFontSize(payload.value) } : null;
            case 'setTextColor':
                return typeof payload?.value === 'string' ? { color: payload.value } : null;
            case 'setBackground': {
                // 取色盘的「无颜色」发的是 `'none'`：对字符高亮来说它是**清掉底色**，
                // 必须翻译成 CSS 的 transparent —— 发成白色会变成「刷一层白底」，
                // 打印时是一块白，去掉高亮的功能也就名存实亡了。
                const color = payload?.value;
                if (typeof color !== 'string' || !color) return null;
                return { background: color === 'none' ? 'transparent' : color };
            }
            case 'toggleBold':
                return { bold: !!payload?.toggle };
            case 'toggleItalic':
                return { italic: !!payload?.toggle };
            case 'toggleUnderline':
                return { underline: !!payload?.toggle };
            case 'toggleStrike':
                return { strike: !!payload?.toggle };
            default:
                return null;
        }
    }

    /**
     * 复制 / 剪切当前选区
     * @param cut - 是否剪切
     */
    private async copySelection(cut: boolean): Promise<void> {
        const text = this.selectedText();
        if (!text) return;
        // 【先记样式再写剪贴板】`navigator.clipboard` 只允许写文本，格式刷那一半
        // 只能自己存：复制时就地把光标处的字符样式快照下来，供「粘贴格式」取用。
        this.copiedStyle = this.view.currentCharStyle();
        try {
            await navigator.clipboard.writeText(text);
        } catch {
            return; // 无剪贴板权限：保持静默，用户仍可用 Ctrl+C
        }
        if (!cut) return;
        const selection = window.getSelection();
        if (selection && selection.rangeCount > 0) {
            selection.getRangeAt(0).deleteContents();
            this.view.sync('剪切', 'single', false);
        }
    }

    /**
     * 从剪贴板粘贴纯文本
     */
    private async pasteFromClipboard(): Promise<void> {
        try {
            const text = await navigator.clipboard.readText();
            if (text) this.view.insertText(text);
        } catch {
            // 无剪贴板权限：保持静默
        }
    }

    /**
     * 「粘贴格式」：把复制时记下的字符样式套到当前选区
     *
     * 折叠光标时交给 {@link DocView.applyCharStyle} 的待生效样式——与工具栏按钮
     * 一个口径：先刷样式再打字，打出来就是那个样式。
     */
    private pasteCopiedStyle(): void {
        if (!this.copiedStyle) return;
        this.view.applyCharStyle(this.copiedStyle);
    }

    // ─────────────────────────── 格式刷 ───────────────────────────

    /**
     * 进入 / 退出格式刷
     *
     * 进入时**立刻抄下当前选区（或光标处）的字符样式**：那是刷子里的「颜料」，
     * 之后再选别的文字不会再抄一次（与 Word 一致）。
     * @param mode - 目标模式；`null` 表示退出
     */
    private startBrush(mode: BrushMode): void {
        if (mode === null) {
            this.brushModeValue = null;
            this.brushStyle = null;
            this.view.setBrushing(false);
            return;
        }
        const style = this.view.currentCharStyle();
        // 折叠光标处没有「样式可刷」的把握（样式是给接下来要打的字用的），
        // 但 Word 在光标状态下也允许拿刷子：抄光标处的代表样式，够用且不会报错。
        this.brushStyle = style;
        this.brushModeValue = mode;
        this.view.setBrushing(true);
    }

    /**
     * 有选区时把刷子里的样式刷上去
     *
     * 触发点用 `mouseup` 而不是 `selectionchange`：后者在拖选过程中会连续触发，
     * 一边拖一边刷会让「还没选完的那半段」也被刷上，且撤销栈被拆成十几步。
     * @returns 是否刷了一次
     */
    private applyBrush(): boolean {
        if (!this.brushModeValue || !this.brushStyle) return false;
        const selection = window.getSelection();
        if (!selection || selection.rangeCount === 0 || selection.isCollapsed) return false;
        const range = selection.getRangeAt(0);
        if (!this.view.getElements().page.contains(range.commonAncestorContainer)) return false;
        // 【复制一份再刷】`applyCharStyle` 会改 DOM 并同步模型，若把 brushStyle 直接递进去，
        // 单次模式下紧随其后的置空会连带清掉刚写进 DOM 的样式引用（补丁对象被复用时尤其危险）。
        const patch: CharStylePatch = { ...this.brushStyle };
        const done = this.view.applyCharStyle(patch);
        if (!done) return false;
        // 单次模式刷完就退出：亮着的按钮会让人以为还能再刷一次
        if (this.brushModeValue === 'single') {
            this.brushModeValue = null;
            this.brushStyle = null;
            this.view.setBrushing(false);
        }
        return true;
    }

    /**
     * 取当前选区文本
     * @returns 文本；无选区时为空串
     */
    private selectedText(): string {
        const selection = window.getSelection();
        if (!selection || selection.rangeCount === 0) return '';
        return selection.getRangeAt(0).toString();
    }

    /**
     * 取文档状态快照
     *
     * 与表格侧的区别：文档**永远没有「活动格」概念**，所以 `char` 一律取光标处
     * （或首字符）的样式，`editing` 只表示「焦点是否在正文里」——
     * 但按钮只看样式，不看这个标志，因此不会有「文档里按钮全灰」的情况。
     * @returns 状态快照
     */
    public snapshot(): EditStateSnapshot {
        const style = this.view.currentCharStyle();
        const para = this.view.currentParaState();
        const char: CharStyleState = {};
        if (style.fontFamily) char.fontFamily = style.fontFamily;
        if (style.fontSize) char.fontSize = toToolbarFontSize(style.fontSize);
        if (style.bold !== undefined) char.bold = style.bold;
        if (style.italic !== undefined) char.italic = style.italic;
        if (style.underline !== undefined) char.underline = style.underline;
        if (style.strike !== undefined) char.strike = style.strike;
        if (style.color) char.color = style.color;
        if (style.background) char.background = style.background;

        const isListItem = para.type === 'listItem';
        const paraState: ParaStyleState = {
            align: para.align,
            // 文档正文天然自动换行，`wrap` 恒为 true；垂直对齐对段落没有意义
            wrap: true,
            // 列表项在 Word 的「样式」框里显示的仍是「正文」，级别由缩进按钮表达
            paraStyle: toParaStyleId(isListItem ? 'paragraph' : para.type, para.level),
            // 缩进级数取自 `data-indent`（块级额外缩进），而不是列表层级：
            // 工具栏的缩进按钮改的是前者，回读后者会让按钮高亮与点击结果对不上
            indent: isListItem ? para.level + para.indent : para.indent,
            list: isListItem ? (para.listStyle === 'ordered' ? 'ordered' : 'bullet') : null,
        };

        return {
            editing: this.isFocused(),
            char,
            para: paraState,
            filterActive: false,
            brushMode: this.brushModeValue,
            undoable: this.store.canUndo,
            redoable: this.store.canRedo,
        };
    }

    /**
     * 焦点是否在正文里
     * @returns 是否聚焦
     */
    private isFocused(): boolean {
        const active = document.activeElement;
        if (!active) return false;
        const blocks = this.view.getElements().page;
        return blocks.contains(active) || active === blocks;
    }

    /**
     * 广播「状态可能已变化」
     *
     * 同时发契约事件（工具栏刷新）与 `IS_EDITTING_CHANGED`（按钮的
     * `isDisabledOnEditting` 特性订阅的是这个原始事件名，见 CommandBus 的转发表）。
     */
    public notifyStateChanged(): void {
        this.emit(EDIT_STATE_CHANGED);
        const editing = this.isFocused();
        if (editing !== this.lastEditing) {
            const previous = this.lastEditing;
            this.lastEditing = editing;
            this.emit(DataEvents.IS_EDITTING_CHANGED, previous, editing);
        }
    }

    /**
     * 撤销（供文档功能区按钮与快捷键调用）
     * @returns 是否执行
     */
    public undo(): boolean {
        const done = this.store.undo();
        this.notifyStateChanged();
        return done;
    }

    /**
     * 重做
     * @returns 是否执行
     */
    public redo(): boolean {
        const done = this.store.redo();
        this.notifyStateChanged();
        return done;
    }
}
