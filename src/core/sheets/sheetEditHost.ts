/*  # 表格编辑宿主 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目描述: 基于 HTML5 Canvas 的 Web 电子表格组件，无需服务端支持即可在浏览器中创建、读取、写入 Excel 文件。
    # 项目目标: 把「表格」包装成工具栏外壳认识的标准编辑宿主（{@link EditHost}），
    #           使工具栏的字体 / 对齐 / 剪贴板 / 格式刷分区与表格数据层彻底脱钩。
    # 本文件职责: 实现 {@link SheetEditHost} ——
    #             canExecute（命令可执行性）、execute（命令 → DataCollection 调用）、
    #             snapshot（DataCollection 状态 → 归一化快照）。
    #             **它是唯一允许「命令名 ↔ 表格 API」互相翻译的地方。**
    # 依赖方向: contract（接口）→ dataArchitecture（数据层）→ 本文件，单向。
    #           本文件可以 import DataCollection，而 contract / toolbar 下的任何文件都不可以。
    # 搬迁原则: 命令分支里的每一行都**原样搬自**改造前的分区 onClick 回调，
    #           连注释一起搬 —— 表格行为必须零变化，这是本次解耦的验收底线。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { DataEvents } from '../constant';
import {
    BrushMode,
    CharStyleState,
    CommandId,
    CommandPayload,
    EDIT_STATE_CHANGED,
    EditHost,
    EditStateSnapshot,
    EditingStateSource,
    ParaStyleState,
    TextAlignValue,
    VerticalAlignValue,
    WorkspaceKind,
} from '../contract/editHost';
import { Cell, TextAlign, VerticalAlign } from '../dataArchitecture/cell';
import { CellBorderStyle, FontStyle } from '../dataArchitecture/cellContent';
import { DataCollection } from '../dataArchitecture/dataCollection';
import { EventEmitter } from '../../utils/eventEmitter';
import { indentStepPx } from '../../utils/textWidth';

/**
 * 需要转成「状态可能已变化」的表格事件
 *
 * 这是改造前四个分区订阅事件的**并集**，不多不少：
 * - `ACTIVED_CELL_CHANGED` / `CURSOR_STATE_CHANGED` / `SNAPSHOT_RESTORED` —— 字体与对齐分区
 * - `VALUES_CHANGED` —— 数据处理分区（筛选按钮高亮）
 * - `BRUSH_MODE_CHANGED` —— 格式刷分区
 * - `IS_EDITTING_CHANGED` —— 全部分区 + 按钮的「编辑时置灰」
 *
 * 刻意**不加**别的表格事件：每多转发一个，四个分区就会跟着多刷一遍，
 * 而快照刷新虽便宜（`getCellByName` 是 O(1)），却会写十几处 DOM 样式。
 */
const STATE_EVENTS: readonly string[] = [
    DataEvents.ACTIVED_CELL_CHANGED,
    DataEvents.SNAPSHOT_RESTORED,
    DataEvents.CURSOR_STATE_CHANGED,
    DataEvents.VALUES_CHANGED,
    DataEvents.BRUSH_MODE_CHANGED,
    DataEvents.IS_EDITTING_CHANGED,
];

/**
 * 表格编辑宿主
 *
 * 把 {@link DataCollection} 的全部「编辑语义」翻译成 {@link CommandId}，供工具栏外壳调用。
 * 外壳侧因此不再出现任何 `data.xxx()`（改造前四个分区共有 51 处直连）。
 *
 * ## 两个职责边界
 *
 * 1. **只翻译，不实现**：所有命令都转发给 `DataCollection` 的既有方法，
 *    本类不做任何数据计算，也不持有任何编辑状态 —— 状态仍然只有 `DataCollection` 一份。
 * 2. **把「编辑态」这句判断收进自己**：改造前每个按钮都写
 *    `if (data.isEditting) setSelectedCharsFont(...) else setSelectedCellsFont(...)`，
 *    同一段分支在字体分区重复了 7 次。现在收敛成 {@link applyCharStyle} 一处，
 *    快照里给出的也是**已经判断完**的样式值。
 *
 * @example
 * ```ts
 * const host = new SheetEditHost(data);
 * const bus = new CommandBus();
 * bus.bind(host);                       // 之后工具栏只认识 bus
 * bus.dispatch('toggleBold', { toggle: true });
 * ```
 */
export class SheetEditHost extends EventEmitter implements EditHost {
    /** @inheritdoc */
    public readonly kind: WorkspaceKind = 'sheet';

    /** 被包装的数据集合（表格侧唯一的状态真源） */
    private readonly data: DataCollection;

    /**
     * 已注册到 `DataCollection` 的监听器
     *
     * 保存引用是为了 {@link dispose} 时精确退订：`EventEmitter.off(event)` 不带 handler
     * 会清掉该事件的全部监听，而 `DataCollection` 上还有画布、滚动条等一大票订阅者。
     */
    private readonly subscriptions: Array<{ event: string; handler: (...args: any[]) => void }> = [];

    /**
     * 筛选是否生效的缓存值
     *
     * 快照会被四个分区在**每个**状态事件里各取一次，而 `validataFilter()` 的快路径一旦不成立
     * （筛选格索引被标脏、`_indexedLength` 落后于 `_values.length`）就会走到 `_ensureCellIndex()`。
     * 改造前只有数据处理分区在 `VALUES_CHANGED` / `SNAPSHOT_RESTORED` 时问一次，
     * 现在若不加缓存，编辑时每一次光标移动都会问四遍，大表上等于把索引重建摊到每个按键上。
     *
     * 因此按「可能改变筛选状态的事件」失效（见 {@link listen}），其余时间直接返回缓存。
     */
    private filterActiveCache: boolean = false;

    /** 缓存是否已失效（需要重算 {@link filterActiveCache}） */
    private filterActiveCacheDirty: boolean = true;

    /**
     * 创建表格编辑宿主
     * @param data - 数据集合实例
     */
    constructor(data: DataCollection) {
        super();
        this.data = data;
        this.listen();
    }

    /**
     * 宿主事件源
     *
     * 直接返回自身：`EventEmitter` 的 `on` / `off` 天然满足 {@link EditingStateSource}，
     * 不必再包一层适配对象（多一层就多一处泄漏监听的可能）。
     * @returns 事件源
     */
    public get stateSource(): EditingStateSource {
        return this;
    }

    /**
     * 该命令在当前宿主下是否可执行
     *
     * 表格宿主实现了命令全集，因此恒为 `true`。
     * 文档宿主将在此对 `setBorder` / `toggleMergeCells` / `toggleFilter` / `sortUp` / `sortDown`
     * 返回 `false`，对应按钮自动置灰 —— 这是「一套 UI 挂两种载体」的收敛机制。
     * @param cmd - 命令名（形参保留以对齐接口，当前未使用）
     * @returns 恒为 true
     */
    public canExecute(cmd: CommandId): boolean {
        return true;
    }

    /**
     * 执行命令
     *
     * 每个分支都是改造前某个分区 `onClick` 回调的**原样搬迁**：
     * 顺序、边界判断、撤销包裹方式、补发的广播，全部保持不变。
     * @param cmd - 命令名
     * @param payload - 命令载荷
     */
    public execute(cmd: CommandId, payload?: CommandPayload): void {
        switch (cmd) {
            /* ── 字符级 ── */
            case 'setFontFamily':
                this.applyCharStyle({ fontFamily: String(payload?.value ?? '') });
                break;
            case 'setFontSize':
                this.applyCharStyle({ fontSize: Number(payload?.value) });
                break;
            case 'toggleBold':
                this.applyCharStyle({ fontWeight: !!payload?.toggle });
                break;
            case 'toggleItalic':
                this.applyCharStyle({ fontStyle: !!payload?.toggle });
                break;
            case 'toggleUnderline':
                this.applyCharStyle({ underline: !!payload?.toggle });
                break;
            case 'toggleStrike':
                this.applyCharStyle({ strikethrough: !!payload?.toggle });
                break;
            case 'setTextColor':
                this.applyCharStyle({ fontColor: String(payload?.value ?? '') });
                break;
            case 'setBackground':
                // 【只走单元格级】背景色是 Cell 的属性，Char 上没有这个字段
                // （见 cellContent.ts 的 FontStyle 与 char.ts 的字段对比）。
                // 改造前油漆桶就只调 setSelectedCellsFont，这里保持原样。
                this.data.setSelectedCellsFont({ backgroundColor: String(payload?.value ?? '') });
                break;

            /* ── 段落级 ── */
            case 'setAlign':
                this.data.setSelectedCellsFont({ textAlign: payload?.value as TextAlign });
                break;
            case 'setVAlign':
                this.data.setSelectedCellsFont({ alignItems: payload?.value as VerticalAlign });
                break;
            case 'toggleWrap':
                this.data.setSelectedCellsFont({ wrap: !!payload?.toggle });
                break;
            case 'indentStep':
                // 载荷是**方向**（+1 / -1），步长由宿主自己算：
                // 「一次缩进多宽」是载体相关知识（表格看字号与缩放，文档看段落级别），
                // 外壳不该知道，因此改造前 SectionAlign#getIndentStep 整体搬到这里。
                this.data.changeSelectedCellsIndent(Number(payload?.value ?? 0) * this.getIndentStep());
                break;

            /* ── 剪贴板 ── */
            case 'copy':
                if (this.data.selectionOnEditting && this.data.isEditting) {
                    // 编辑状态下有字符选择，进行字符级复制
                    this.data.copyChars();
                } else {
                    // 单元格级复制
                    this.data.copySelection();
                }
                break;
            case 'cut':
                if (this.data.selectionOnEditting && this.data.isEditting) {
                    // cutChars 返回目标光标绝对位置（被删除区间起点）。
                    // VALUES_CHANGED 会同步触发重绘并刷新 data.lines，因此此处写入
                    // cursorAbsolutePosition 时 lines 已是最新布局。
                    const caretIndex = this.data.cutChars();
                    this.data.cursorAbsolutePosition = caretIndex;
                } else {
                    // 单元格级剪切：与右键菜单同口径包裹，大表按选区稀疏补丁记账可撤销/重做
                    this.data.runWithFullStateUndo('剪切', () => this.data.cutSelection(), 'selection');
                    // 包裹抑制了 commitValues 的自动广播，这里补一次驱动重绘（draw 按 rAF 合帧）
                    this.data.emit(DataEvents.VALUES_CHANGED, this.data.values);
                }
                break;
            case 'paste':
                if (this.data.isEditting) {
                    // 编辑状态下进行字符级粘贴
                    const clipboard = this.data.clipboardManager.getClipboard();
                    if (clipboard && clipboard.type === 'char' && clipboard.chars) {
                        // getSelectedCharsStartIndex(): 有选区时返回选区起点，无选区时返回光标处，
                        // 即真正的插入起点。pasteChars 返回插入完成后光标应处的绝对位置。
                        const insertIndex = this.data.getSelectedCharsStartIndex();
                        const caretIndex = this.data.pasteChars(insertIndex, clipboard.chars);
                        this.data.cursorAbsolutePosition = caretIndex;
                    }
                } else {
                    // 非编辑态单元格级粘贴：与右键菜单同口径包裹，'paste' 档按粘贴目标矩形
                    // 捕获稀疏补丁，大表同样可撤销/重做（超大块粘贴超上限时放弃记账）
                    this.data.runWithFullStateUndo('粘贴', () => this.data.pasteSelection(), 'paste');
                    this.data.emit(DataEvents.VALUES_CHANGED, this.data.values);
                }
                break;
            case 'pasteValue':
                // 编辑态下不纳入全量/补丁记账（字符级编辑有独立编辑内栈），保持裸调用
                if (this.data.isEditting) {
                    this.data.pastePlainText();
                } else {
                    this.data.runWithFullStateUndo('只粘贴文本', () => this.data.pastePlainText(), 'paste');
                    this.data.emit(DataEvents.VALUES_CHANGED, this.data.values);
                }
                break;
            case 'pasteStyle':
                // 同上：编辑态裸调用，避免污染编辑内撤销栈
                if (this.data.isEditting) {
                    this.data.pasteFormatOnly();
                } else {
                    this.data.runWithFullStateUndo('只粘贴格式', () => this.data.pasteFormatOnly(), 'paste');
                    this.data.emit(DataEvents.VALUES_CHANGED, this.data.values);
                }
                break;

            /* ── 格式刷 ── */
            case 'brushClick':
                // 连续模式再点一次 = 退出；否则进入单次模式并立刻抓取源格式。
                // 按钮高亮由 UI 按快照的 brushMode 自行推导（改造前是 BRUSH_MODE_CHANGED
                // 的订阅回调，现在统一走 EDIT_STATE_CHANGED）。
                if (this.data.brushMode === 'continuous') {
                    this.data.brushMode = null;
                } else {
                    this.data.brushMode = 'single';
                    this.data.copySelection();
                }
                break;
            case 'brushDoubleClick':
                if (this.data.brushMode === 'continuous') {
                    this.data.brushMode = null;
                } else {
                    this.data.brushMode = 'continuous';
                }
                this.data.copySelection();
                break;

            /* ── 表格专属 ── */
            case 'setBorder':
                this.data.setSelectedCellsBorder(payload?.value as CellBorderStyle);
                break;
            case 'toggleMergeCells':
                this.data.toggleMergeCells(payload?.value as 'row' | 'col' | undefined);
                break;
            case 'toggleFilter':
                this.data.setFilter();
                // 筛选开关由快照回读；`setFilter` 不保证发 VALUES_CHANGED（取消筛选可能只改行可见性），
                // 这里显式失效，保证下一次快照读到的是真实状态
                this.filterActiveCacheDirty = true;
                break;
            case 'sortUp':
                this.data.sortUp();
                break;
            case 'sortDown':
                this.data.sortDown();
                break;
            default:
                // 穷尽性检查：新增命令而此处未补分支时，编译期在这里报错
                this.assertNever(cmd);
        }
    }

    /**
     * 取当前状态的归一化快照
     *
     * 纯读取：只调 `getCellByName`（名称索引，O(1)）与几个 getter，不落库、不广播、不新建 Cell。
     * @returns 状态快照
     */
    public snapshot(): EditStateSnapshot {
        const editing = this.data.isEditting;
        const cell = this.data.getCellByName(this.data.activedCell);

        // 筛选状态按事件失效，不在快照里现算（见 filterActiveCache 的注释）
        if (this.filterActiveCacheDirty) {
            this.filterActiveCache = this.data.validataFilter();
            this.filterActiveCacheDirty = false;
        }

        return {
            editing,
            // 编辑态读光标处字符（没字符就是空对象 → UI 回退到默认字体/字号），
            // 非编辑态读活动格的代表值（Cell 的字体属性本身就是「取首字符」）
            char: editing ? this.readCursorChar(cell) : this.readCellChar(cell),
            para: this.readParaStyle(cell),
            filterActive: this.filterActiveCache,
            brushMode: this.data.brushMode as BrushMode,
        };
    }

    /**
     * 退订全部表格事件
     *
     * 宿主被替换（工作区切换）或文件关闭时调用，避免宿主对象被 `DataCollection` 的监听表钉住不放。
     */
    public dispose(): void {
        for (const { event, handler } of this.subscriptions) {
            this.data.off(event, handler);
        }
        this.subscriptions.length = 0;
    }

    /**
     * 订阅表格事件并转成「状态可能已变化」
     *
     * `IS_EDITTING_CHANGED` 额外**原样转发一次**：按钮的 `isDisabledOnEditting` 特性
     * 历史上就是订阅这个事件（回调签名 `(oldValue, newValue)`），总线按同名事件转发给它。
     * @private
     */
    private listen(): void {
        for (const event of STATE_EVENTS) {
            const handler = (...args: any[]): void => {
                // 只有「数据/结构真的变了」才可能改变筛选生效状态；
                // 光标移动、活动格切换都不影响，不必为它们重算
                if (event === DataEvents.VALUES_CHANGED || event === DataEvents.SNAPSHOT_RESTORED) {
                    this.filterActiveCacheDirty = true;
                }
                if (event === DataEvents.IS_EDITTING_CHANGED) {
                    this.emit(event, ...args);
                }
                this.emit(EDIT_STATE_CHANGED);
            };
            this.data.on(event, handler);
            this.subscriptions.push({ event, handler });
        }
    }

    /**
     * 施加字符级样式（改造前在四个分区重复了 12 次的同一段分支）
     *
     * 编辑会话中改**已输入字符**，否则改**选中单元格的全部字符**。
     * @param style - 待施加的字体样式
     * @private
     */
    private applyCharStyle(style: FontStyle): void {
        if (this.data.isEditting) {
            this.data.setSelectedCharsFont(style);
        } else {
            this.data.setSelectedCellsFont(style);
        }
    }

    /**
     * 读编辑会话中光标前一个字符的样式
     *
     * `cursorAbsolutePosition` 是 1-based 的插入位置，减 1 才是「光标左边那个字符」的下标；
     * 不改！减成 0-based 的原因是改造前 `SectionFont#setToolbarsState` 就是这么取的，
     * 换成别的口径会让「刚点进格子时按钮高亮」与原来差一格。
     * @param cell - 活动单元格
     * @returns 字符样式（无字符时为空对象，由 UI 回退默认值）
     * @private
     */
    private readCursorChar(cell: Cell | undefined): CharStyleState {
        const cursorPosition = this.data.cursorAbsolutePosition - 1;
        const char = cell?.chars?.[cursorPosition];
        if (!char) return {};
        return {
            fontFamily: char.fontFamily,
            fontSize: char.fontSize,
            bold: char.fontWeight,
            italic: char.fontStyle,
            underline: char.underline,
            strike: char.strikethrough,
            color: char.fontColor,
        };
    }

    /**
     * 读活动单元格级别的样式
     *
     * `Cell` 上的字体属性本身就是「取首字符」（见 cell.ts 的 `_hasUniformCharStyle` 守卫），
     * 因此非编辑态不需要自己去翻 `chars`。
     * @param cell - 活动单元格
     * @returns 字符样式（无单元格时为空对象）
     * @private
     */
    private readCellChar(cell: Cell | undefined): CharStyleState {
        if (!cell) return {};
        return {
            fontFamily: cell.fontFamily,
            fontSize: cell.fontSize,
            bold: cell.fontWeight,
            italic: cell.fontStyle,
            underline: cell.underline,
            strike: cell.strikethrough,
            color: cell.fontColor,
            background: cell.backgroundColor,
        };
    }

    /**
     * 读段落 / 布局级样式
     *
     * `wrap` 归一化成布尔（`Cell.wrap` 是三态：undefined 表示继承默认的「不换行」），
     * 快照的消费者只需要「亮不亮」这一个布尔。
     * @param cell - 活动单元格
     * @returns 段落样式
     * @private
     */
    private readParaStyle(cell: Cell | undefined): ParaStyleState {
        return {
            align: cell?.textAlign as TextAlignValue | undefined,
            valign: cell?.alignItems as VerticalAlignValue | undefined,
            wrap: !!cell?.wrap,
            indent: cell?.indent,
        };
    }

    /**
     * 获取单次缩进量（2 个字符宽，「2em」）
     *
     * ## 基准是**当前单元格**的字号
     *
     * em 取活动单元格的实时字号（不是根元素字号）：16px 的格子汉字缩进 32px，
     * 20px 则是 40px —— 缩进与文字大小始终同比例。
     * 每次点击都现取（{@link indentStepPx} 不做缓存），改过字号立刻生效。
     * 单元格为空（字号挂在字符上，空格取不到）时回退根元素字号。
     *
     * ## 单字节 / 双字节分别折算
     *
     * 一次缩进 = 2 个字符宽，而字符宽取决于内容：双字节（汉字/全角）1 字 = 1em、
     * 单字节（数字/英文）1 字 = 0.5em。故同样是一次缩进：
     * - 汉字单元格：2 × 1em = 2em（32px @16px）
     * - 数字单元格：2 × 0.5em = 1em（16px @16px）
     * 两者视觉上都是「缩进两个字」，不会出现「中文正好、数字缩了四个字」。
     * @returns {number} 单次缩进像素值（已含缩放）
     * @private
     */
    private getIndentStep(): number {
        // 【必须每次现取】字号是逐字符属性（挂在 chars 上），改完字号立刻点击也要拿到新值，
        // 因此这里不做任何缓存：`Cell.fontSize` 取首字符字号，无字符时是 undefined。
        const cell = this.data.getCellByName(this.data.activedCell);
        let fontSize = cell?.fontSize;
        if (fontSize === undefined || !(fontSize > 0)) {
            // 单元格为空（字号无处可挂）时回退到根元素字号 —— 仍是「相对当前环境」的 1em
            fontSize = parseFloat(getComputedStyle(document.documentElement).fontSize) || 16;
        }
        // 单/双字节字符宽度不同：汉字 1em、数字字母 0.5em，取首字符判断（与 Cell.fontSize 同源）
        const sampleChar = cell?.chars?.[0]?.char;
        return indentStepPx(fontSize, sampleChar) * this.data.zoom;
    }

    /**
     * 穷尽性检查辅助
     *
     * `switch` 覆盖了 {@link CommandId} 的全部取值时，`cmd` 在此处被收窄为 `never`；
     * 一旦新增命令忘了补分支，`cmd` 就是那个新字符串，编译期直接报错。
     * @param value - 理论上不可能到达的值
     * @private
     */
    private assertNever(value: never): void {
        throw new Error(`未实现的编辑命令: ${String(value)}`);
    }
}
