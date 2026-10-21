/*  # web 表格辅助组件-python 编辑器 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { DataEvents, PY_EDITOR_CLASS_NAMES } from '../constant';
import { IndexDB } from '../dataArchitecture/indexDB';
import { DataCollection } from '../dataArchitecture/dataCollection';
import { createDiv, createElement, createSpan } from '../../utils/dom';
import { runPython, stopPython } from '../../utils/python/pythonClient';
import { PY_NAMESPACE_SOURCE, buildSnapshot, applyWrites } from '../../utils/python/sheetNameSpace';
import '../style/pyEditor.scss';

/** Python 关键字（高亮为「关键字」色） */
const PY_KEYWORDS = new Set([
  'False','None','True','and','as','assert','async','await','break','class',
  'continue','def','del','elif','else','except','finally','for','from','global',
  'if','import','in','is','lambda','nonlocal','not','or','pass','raise','return',
  'try','while','with','yield'
]);

/** Python 常见内置名（高亮为「内置」色；`self` 按惯例名一并高亮） */
const PY_BUILTINS = new Set([
  'print','len','range','int','str','float','list','dict','set','tuple','bool',
  'type','isinstance','enumerate','zip','map','filter','sum','min','max','abs',
  'open','input','super','self','Exception','ValueError','TypeError'
]);

/**
 * Python 词法分段的组合正则（同一轮匹配中按优先级从左到右竞争）：
 * - 组1 注释：`#` 到行尾；
 * - 组2 字符串：可选前缀（f/r/b/u 组合）+ 三引号块 / 单双引号行串（含转义）；
 * - 组3 数字：整数 / 小数 / 科学计数 / 复数后缀 j，允许下划线分隔；
 * - 组4 装饰器：`@名称`；
 * - 组5 标识符：再按 {@link PY_KEYWORDS} / {@link PY_BUILTINS} 细分。
 *
 * 字符串先于标识符匹配：引号内的单词不会被误判为关键字。
 * 三引号未闭合时整块不匹配（该段按普通文本处理），属于简单词法器的可接受误差。
 */
const PY_TOKEN_PATTERN = /(#[^\n]*)|([fFrRbBuU]{0,2}(?:'''[\s\S]*?'''|"""[\s\S]*?"""|'(?:\\.|[^'\\\n])*'|"(?:\\.|[^"\\\n])*"))|(\b\d[\d_]*(?:\.[\d_]*)?(?:[eE][+-]?\d+)?[jJ]?\b)|(@[A-Za-z_]\w*)|([A-Za-z_]\w*)/g;

/** token 类型对应的高亮 class（见 pyEditor.scss 配色） */
const PY_TOKEN_CLASS = {
    COMMENT: 'py-tok-comment',
    STRING: 'py-tok-string',
    NUMBER: 'py-tok-number',
    DECORATOR: 'py-tok-decorator',
    KEYWORD: 'py-tok-keyword',
    BUILTIN: 'py-tok-builtin',
} as const;

/**
 * 转义 HTML 特殊字符
 *
 * 高亮片段以 innerHTML 拼接，代码文本必须先转义，
 * 否则代码里的 `<div>` 之类会被浏览器当成真实标签解析。
 * @param text - 原始代码文本
 * @returns 转义后的文本
 */
function escapeHtml(text: string): string {
    return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/**
 * Python 代码编辑器（「脚本」导航页内容组件）
 *
 * ## 实现方案：透明输入层 + 高亮层叠
 *
 * 不引入第三方编辑器库（与项目「原生实现」的风格一致），采用经典的
 * 双层结构：下方 `pre` 渲染语法高亮，上方 `textarea` 背景与文字全透明、
 * 只保留光标与选区，负责全部输入与选择。两层使用**完全一致的排版参数**
 * （字体 / 字号 / 行高 / 内边距 / tab-size），文字逐字符对齐；
 * `textarea` 是唯一滚动容器，高亮层与行号栏的滚动位置由 JS 同步。
 *
 * ## 功能
 *
 * - 语法高亮：关键字 / 内置名 / 字符串 / 注释 / 数字 / 装饰器；
 * - 行号栏：随内容行数更新，与代码区同步滚动；
 * - 缩进辅助：Tab 缩进、Shift+Tab 反缩进（选区跨行时按行批量处理）、
 *   回车自动延续缩进（行尾为 `:` 时自动加一级）；
 * - 状态栏：实时显示光标所在行 / 列；
 * - 运行脚本：工具栏「运行 / 停止」按钮，经 Pyodide Worker 执行当前脚本
 *   （见 {@link PYEditor.handleRun}），stdout / stderr 实时流式追加到页签下方的
 *   输出面板，脚本对工作表的写入会套进撤销栈落回表格；
 * - 持久化：内容变化防抖写入 IndexedDB，**按当前工作表名分键存储**，
 *   切换工作表时先落盘旧表脚本、再载入新表脚本，刷新后自动恢复；
 * - 选区保持：存在选区（如 Ctrl+A 全选）时，点击选中的文字保持选区不变，
 *   点击空白处（行尾右侧 / 文本下方）则按浏览器默认行为取消选区、定位光标。
 *
 * @class PYEditor
 * @example
 * const editor = new PYEditor({ data });
 * navigator.navContentItemMount('脚本', editor.getElements().container);
 */
export class PYEditor {
    /** 主数据集合（当前工作表名与工作表切换事件的来源） */
    private readonly data: DataCollection;
    /** 根容器 */
    private readonly container: HTMLDivElement;
    /** 行号栏 */
    private readonly gutter: HTMLPreElement;
    /** 高亮渲染层（`pre`，不可滚动，滚动位置由输入层同步） */
    private readonly highlight: HTMLPreElement;
    /** 透明输入层（`textarea`，唯一可滚动层） */
    private readonly textarea: HTMLTextAreaElement;
    /** 状态栏（行 / 列指示） */
    private readonly status: HTMLSpanElement;
    /** 工具栏「运行」按钮 */
    private readonly runButton: HTMLButtonElement;
    /** 工具栏「停止」按钮 */
    private readonly stopButton: HTMLButtonElement;
    /** 输出面板（stdout / stderr / 运行结果） */
    private readonly output: HTMLDivElement;
    /**
     * 是否正在运行
     *
     * 兼作串行保护：Pyodide 的 stdout/stderr 绑定是「按当前请求 id」的，
     * 两次运行重叠会让输出串到错误的运行上（见 pythonClient 的并发约定）。
     */
    private isRunning = false;
    /** 界面状态持久化存储 */
    private readonly db: IndexDB;
    /** 当前行号栏渲染的行数（行数未变时跳过重排） */
    private gutterLineCount = 0;
    /** 防抖保存定时器 */
    private saveTimer: ReturnType<typeof setTimeout> | null = null;
    /** 用户是否已产生输入（restore 异步返回时不覆盖用户新输入） */
    private dirty = false;
    /**
     * 文本宽度测量上下文（鼠标坐标换算为文本偏移时用）
     *
     * 输入层是等宽字体，但中文等非 ASCII 字符宽度并非整数格，
     * 逐行前缀测量才能在混排时把「点击像素」精确换算为「文本偏移」。
     * 只在鼠标按下/拖动时使用（低频），字体与输入层保持一致。
     */
    private measureCtx: CanvasRenderingContext2D | null = null;
    /** 有选区时按下鼠标的落点偏移（mouseup 未拖动时据此落光标；null 表示当前没有待定的按下） */
    private downOffset: number | null = null;
    /** 按下后是否发生过拖动（拖则自实现选区延伸，未拖则收拢为点击处光标） */
    private isDragSelecting = false;

    /** 脚本内容在 IndexedDB 设置表中的键前缀（完整键 = 前缀 + 工作表名） */
    private static readonly STORAGE_KEY_PREFIX = 'py-editor:script:';
    /** 内容变化后写入 IndexedDB 的防抖窗口（毫秒） */
    private static readonly SAVE_DEBOUNCE_MS = 400;
    /** 一级缩进（4 空格，PEP 8） */
    private static readonly INDENT = '    ';

    /**
     * 旧版默认脚本模板（历史存档迁移用）
     *
     * 「首行标题随文件名动态生成」之前的默认模板：对所有文件都是同一句固定标题
     * `# SilkSpaces Python 脚本`。迁移时据此**逐字符**比对，只有从未被编辑过的
     * 历史存档才会命中，用户真实编写过的内容不受影响。
     */
    private static readonly LEGACY_DEFAULT_SCRIPT = [
        '# SilkSpaces Python 脚本',
        '# 在此编写脚本，内容会自动保存在本机（IndexedDB）',
        '',
        '',
        'B2 = "Hello, SilkSpaces!"',
        '',
        '',
    ].join('\n');

    /**
     * 旧默认存档迁移的共享 Promise（null = 尚未开始 / 上次迁移失败可重试）
     *
     * 构造与切换工作表各会触发一次 {@link restore}，两处都必须在读取存档前等迁移
     * 完成，用同一个 Promise 保证全库只扫一次、且后发者能等到先发者的结果。
     */
    private static legacyMigration: Promise<void> | null = null;

    /**
     * 生成「该文件尚无存档」时的默认脚本内容
     *
     * 首行注释带上这个文件的名字，使用户一眼看清「这段脚本属于哪个文件」：
     * 每个文件各有自己的脚本，若默认模板对所有文件都是同一句固定标题，
     * 无存档的新文件之间就无法区分。名字取
     * {@link DataCollection.activedSheetName}（它同时是该文件脚本的存储键，
     * 故注释与存档键天然一致）。
     *
     * 注释里同时列出读写当前工作表可用的 API：脚本页签里没有别的文档入口，
     * 把 API 清单放在「新建文件时第一眼看到的地方」是唯一不增加界面的做法。
     * @param sheetName - 文件名（工作表名）；为空（未打开文件）时退化为不带名字的标题
     * @returns 默认脚本文本
     */
    private static defaultScript(sheetName: string): string {
        const title = sheetName ? `${sheetName} Python 脚本` : 'Python 脚本';
        return [
            `# ${title}`,
            '# 在此编写脚本，内容会自动保存在本机（IndexedDB）',
            '#',
            '# 可用 API（读写当前工作表）：',
            '#   sheet_name / row_count / col_count   当前工作表名、行数、列数',
            '#   get_cell("A1")                        读取单元格文本（空格子返回空串）',
            '#   set_cell("A1", "值") / clear_cell("A1")  写入 / 清空单元格，运行后可 Ctrl+Z 撤销',
            '#   get_range("A1:C3")                    读取区域，返回二维列表',
            '#   set_range("A1:C3", [["a1","b1"],["a2","b2"]])  写入区域，以 A1 为左上角',
            '#',
            '# 也可把单元格名直接当变量名（列字母须大写，x1 / b2 这类小写名仍按普通变量），',
            '# 模块级与 def 函数体内都成立：',
            '#   A1 = 1                等价于 set_cell("A1", 1)',
            '#   C10 = "2026-10-20"    等价于 set_cell("C10", "2026-10-20")',
            '#   x = A1                等价于 x = get_cell("A1")',
            '#   def main():',
            '#       B2 = 42           函数体内同样写入 B2 单元格',
            '',
            '',
            'B2 = "Hello, SilkSpaces!"',
            '',
            '',
        ].join('\n');
    }

    constructor(options: { data: DataCollection }) {
        this.data = options.data;
        this.db = new IndexDB();
        this.container = createDiv({ className: PY_EDITOR_CLASS_NAMES.CONTAINER });

        this.status = createSpan({ className: PY_EDITOR_CLASS_NAMES.STATUS });
        // 按钮只放图标（i.icon-play3 / i.icon-stop2，来自 icon.css），文字提示走 title
        this.runButton = createElement('button', {
            className: PY_EDITOR_CLASS_NAMES.RUN,
            attributes: { type: 'button', title: '运行当前脚本' },
        }) as HTMLButtonElement;
        this.runButton.appendChild(createElement('i', { className: 'icon-play3' }));
        this.stopButton = createElement('button', {
            className: PY_EDITOR_CLASS_NAMES.STOP,
            attributes: { type: 'button', disabled: 'disabled', title: '强制终止正在运行的脚本' },
        }) as HTMLButtonElement;
        this.stopButton.appendChild(createElement('i', { className: 'icon-stop2' }));
        const actions = createDiv({
            className: PY_EDITOR_CLASS_NAMES.ACTIONS,
            children: [this.runButton, this.stopButton],
        });
        const toolbar = createDiv({
            className: PY_EDITOR_CLASS_NAMES.TOOLBAR,
            children: [
                createSpan({ className: PY_EDITOR_CLASS_NAMES.TITLE, textContent: 'Python 脚本' }),
                actions,
                this.status,
            ],
        });

        this.gutter = createElement('pre', { className: PY_EDITOR_CLASS_NAMES.GUTTER }) as HTMLPreElement;
        this.highlight = createElement('pre', { className: PY_EDITOR_CLASS_NAMES.HIGHLIGHT }) as HTMLPreElement;
        this.textarea = createElement('textarea', {
            className: PY_EDITOR_CLASS_NAMES.INPUT,
            attributes: {
                spellcheck: 'false',
                autocapitalize: 'off',
                autocomplete: 'off',
                // 关闭软换行：与 pre 的 white-space:pre 对齐方式一致，
                // 长行以水平滚动呈现，行号才能与代码行一一对应
                wrap: 'off',
            },
        }) as HTMLTextAreaElement;

        const codeWrap = createDiv({
            className: PY_EDITOR_CLASS_NAMES.CODE,
            children: [this.highlight, this.textarea],
        });
        const body = createDiv({
            className: PY_EDITOR_CLASS_NAMES.BODY,
            children: [this.gutter, codeWrap],
        });
        // tabindex=-1：容器可获焦点，从而接收键盘事件（Ctrl+C 复制等）
        this.output = createDiv({
            className: PY_EDITOR_CLASS_NAMES.OUTPUT,
            attributes: { tabindex: '-1' },
        });
        this.container.append(toolbar, body, this.output);

        this.bindEvents();
        this.setValue(PYEditor.defaultScript(this.data.activedSheetName));
        // 异步恢复当前工作表上次的脚本；期间若用户已开始输入，以用户内容为准（见 restore）
        void this.restore();
        // 工作表切换：先落盘旧表脚本，再载入新表脚本（事件参数为 (旧名, 新名)，见 updateProperty）
        this.data.on(DataEvents.ACTIVED_SHEET_NAME_CHANGED, (oldName: string, newName: string) => {
            this.handleSheetChanged(oldName, newName);
        });
    }

    /**
     * 获取编辑器当前内容
     * @returns 脚本文本
     */
    public getValue(): string {
        return this.textarea.value;
    }

    /**
     * 设置编辑器内容（高亮、行号、状态栏随之刷新）
     * @param code - 脚本文本
     */
    public setValue(code: string): void {
        this.textarea.value = code;
        this.refresh();
    }

    /**
     * 获取组件根元素（由调用方挂进导航内容区）
     * @returns 根元素对象
     */
    public getElements(): { container: HTMLDivElement } {
        return { container: this.container };
    }

    /**
     * 运行当前脚本（工具栏「运行」）
     *
     * 流程：导出当前工作表快照 → 交给 Pyodide Worker 执行（stdout/stderr 实时
     * 追加到输出面板）→ 把脚本的写回改动套进撤销栈落到表格上。
     *
     * 失败一律在输出面板里呈现（而不是弹窗）：
     * - Worker 基础设施不可用 → {@link PythonWorkerUnavailableError}，文案里带排查提示；
     * - 脚本自身异常 → Python 的完整 traceback；
     * - 用户点了「停止」或运行超时 → 由 {@link runPython} 的拒绝文案给出。
     * @private
     */
    private async handleRun(): Promise<void> {
        if (this.isRunning) return;
        this.setRunningState(true);
        this.clearOutput();
        this.appendOutput('正在运行…（首次运行需先加载 Python 运行环境）', false);

        const snapshot = buildSnapshot(this.data);
        try {
            const { writesJson } = await runPython({
                script: this.getValue(),
                sheetName: snapshot.sheetName,
                snapshotJson: snapshot.json,
                rowCount: snapshot.rowCount,
                colCount: snapshot.colCount,
                namespaceSource: PY_NAMESPACE_SOURCE,
                onStdout: (text: string) => this.appendOutput(text, false),
                onStderr: (text: string) => this.appendOutput(text, true),
            });
            const count = applyWrites(this.data, writesJson, '运行 Python 脚本');
            this.appendOutput(count > 0 ? `运行完成，已写入 ${count} 个单元格` : '运行完成', false);
        } catch (err) {
            this.appendOutput(err instanceof Error ? err.message : String(err), true);
        } finally {
            this.setRunningState(false);
        }
    }

    /**
     * 停止正在运行的脚本（工具栏「停止」）
     *
     * 死循环里的 Pyodide 没有任何抢占点，唯一可靠的打断方式是丢弃整个 Worker
     * （见 `pythonClient.stopPython`）。代价是下次运行需要重新加载运行环境，
     * 这一点由 {@link handleRun} 开头的提示覆盖。
     * 停止后 {@link runPython} 会以「已停止」被拒绝，提示文案由 handleRun 的
     * catch 分支统一负责，这里不重复输出。
     * @private
     */
    private handleStop(): void {
        if (!this.isRunning) return;
        stopPython();
    }

    /**
     * 切换「运行中 / 空闲」的按钮状态
     * @param running - 是否正在运行
     * @private
     */
    private setRunningState(running: boolean): void {
        this.isRunning = running;
        this.runButton.disabled = running;
        this.stopButton.disabled = !running;
    }

    /** 清空输出面板（每次运行开始时调用） */
    private clearOutput(): void {
        this.output.textContent = '';
    }

    /**
     * 向输出面板追加文本（实时流式）
     *
     * Pyodide 的 `batched` 回调按批给出文本（内含换行），故按行拆分后逐行建节点：
     * 一行一个 div，空行也保留（`min-height` 维持行高），空行被折叠会破坏
     * 脚本输出的可读性。末尾因「文本以换行结尾」多出的空段则丢弃，否则每次
     * 追加都会凭空多出一行空白。
     * @param text - 待追加的文本（可能含多行）
     * @param isError - 是否为错误输出（stderr / traceback），用红色区分
     * @private
     */
    private appendOutput(text: string, isError: boolean): void {
        if (!text) return;
        const lines = text.split('\n');
        if (lines[lines.length - 1] === '') lines.pop();
        for (const line of lines) {
            const el = createDiv({
                className: PY_EDITOR_CLASS_NAMES.OUTPUT_LINE,
                textContent: line,
            });
            if (isError) el.classList.add(PY_EDITOR_CLASS_NAMES.OUTPUT_ERROR);
            this.output.appendChild(el);
        }
        // 流式输出必须跟着滚动到底，否则运行中看不到最新一行
        this.output.scrollTop = this.output.scrollHeight;
    }

    /**
     * 绑定输入 / 滚动 / 键盘事件
     * @private
     */
    private bindEvents(): void {
        this.runButton.addEventListener('click', () => { void this.handleRun(); });
        this.stopButton.addEventListener('click', () => this.handleStop());
        // 输出面板的复制 / 剪切 / 全选交还给浏览器：
        // editCanvas 在 document 级（冒泡阶段）监听 keydown，遇非输入类元素会
        // preventDefault 并复制表格选区，会把输出面板的 Ctrl+C 抢走。这里在
        // 容器上截断冒泡（不阻止默认行为，浏览器原生复制照常执行）。
        this.output.addEventListener('keydown', (e: KeyboardEvent) => {
            if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
            if (e.key === 'c' || e.key === 'C' || e.key === 'x' || e.key === 'X'
                || e.key === 'a' || e.key === 'A') {
                e.stopPropagation();
            }
        });
        this.textarea.addEventListener('input', () => {
            this.dirty = true;
            this.refresh();
            this.persistDebounced();
        });
        // textarea 是唯一滚动容器：高亮层与行号栏跟随其滚动位置
        this.textarea.addEventListener('scroll', () => this.syncScroll());
        this.textarea.addEventListener('keydown', (e: KeyboardEvent) => this.handleKeydown(e));
        // 存在选区时：按下不做任何操作（高亮保持），抬起后才在点击处落光标并清除高亮
        // （见 handleMousedown / handleSelectionDragUp）
        this.textarea.addEventListener('mousedown', (e: MouseEvent) => this.handleMousedown(e));
        // 光标移动不触发 input（方向键、点击），单独刷新状态栏
        this.textarea.addEventListener('keyup', () => this.renderStatus());
        this.textarea.addEventListener('click', () => this.renderStatus());
    }

    /**
     * 处理鼠标按下：存在选区时「按下不做任何操作」，落光标/清除高亮推迟到 mouseup
     *
     * textarea 原生行为是「按下即折叠选区」，那样高亮在按下瞬间就消失了，
     * 与「先保持高亮、抬起后再落光标」的交互不符。因此这里在**有选区**时
     * preventDefault 阻止原生折叠，并临时挂上 document 级 move/up 监听：
     * - 松手前指针有移动 → 视为拖拽，自实现选区延伸（原生拖选已被 preventDefault 禁掉）；
     * - 松手时未移动 → 只在点击处落光标，高亮随之取消。
     * 无选区时不干预，直接交回浏览器完成定位。
     * @param e - 鼠标事件
     * @private
     */
    private handleMousedown(e: MouseEvent): void {
        const ta = this.textarea;
        if (e.button !== 0) return;                        // 只处理左键
        if (ta.selectionStart === ta.selectionEnd) return; // 无选区：默认行为
        e.preventDefault();                                // 阻止「按下即折叠」，让高亮保持到 mouseup
        ta.focus();                                        // preventDefault 会一并阻止焦点获取，手动补回
        this.downOffset = this.caretOffsetFromPoint(e);
        this.isDragSelecting = false;
        // 仅在按下期间挂监听：抬起即摘除，不常驻 document
        document.addEventListener('mousemove', this.handleSelectionDragMove);
        document.addEventListener('mouseup', this.handleSelectionDragUp);
    }

    /**
     * 拖拽延伸选区（仅在「有选区的按下期间」挂载）
     *
     * 用按下点作锚点，把选区从锚点延伸到当前指针 —— 替代被 preventDefault
     * 禁掉的原生拖选，保证「按下保持高亮」的同时仍能重新框选。
     * @param e - 鼠标事件
     * @private
     */
    private handleSelectionDragMove = (e: MouseEvent): void => {
        if (this.downOffset === null) return;
        this.isDragSelecting = true;
        const offset = this.caretOffsetFromPoint(e);
        const start = Math.min(this.downOffset, offset);
        const end = Math.max(this.downOffset, offset);
        this.textarea.setSelectionRange(start, end);
        this.renderStatus();
    };

    /**
     * 结束按下：未拖动则把选区收拢成点击处的光标（高亮随之取消）
     *
     * @private
     */
    private handleSelectionDragUp = (): void => {
        document.removeEventListener('mousemove', this.handleSelectionDragMove);
        document.removeEventListener('mouseup', this.handleSelectionDragUp);
        const offset = this.downOffset;
        this.downOffset = null;
        if (!this.isDragSelecting && offset !== null) {
            // 只在点击区域落光标：setSelectionRange(offset, offset) 即收拢选区、取消高亮
            this.textarea.setSelectionRange(offset, offset);
        }
        this.isDragSelecting = false;
        this.renderStatus();
    };

    /**
     * 把鼠标坐标换算为文本偏移（空白处就近取整）
     *
     * 换算口径：行 = 纵坐标 / 行高，列 = 逐字符前缀二分测量（等宽字体下中文等
     * 非等宽字符也能精确命中）。与旧实现不同：这里不再用 null 表示空白，
     * 因为 mouseup 必须**总能**给出一个落点 —— 文本下方取全文末尾、行尾右侧
     * 取该行行尾、空行取该行行首。坐标一律以输入层矩形左上角为原点（document
     * 级 mousemove 的 offsetX 是相对事件目标元素的，不能用）。
     * @param e - 鼠标事件（使用 clientX/clientY，与事件目标无关）
     * @returns 文本偏移（0 ~ 文本长度）
     * @private
     */
    private caretOffsetFromPoint(e: MouseEvent): number {
        const ta = this.textarea;
        const style = getComputedStyle(ta);
        const lineHeight = parseFloat(style.lineHeight);
        if (!Number.isFinite(lineHeight) || lineHeight <= 0) return 0;
        const paddingTop = parseFloat(style.paddingTop) || 0;
        const paddingLeft = parseFloat(style.paddingLeft) || 0;
        const rect = ta.getBoundingClientRect();
        // 文字原点 = (左上角 + padding - 滚动偏移)
        const textX = e.clientX - rect.left - paddingLeft + ta.scrollLeft;
        const textY = e.clientY - rect.top - paddingTop + ta.scrollTop;

        const lines = ta.value.split('\n');
        const row = Math.max(0, Math.min(lines.length - 1, Math.floor(textY / lineHeight)));
        let offset = 0;
        for (let i = 0; i < row; i++) offset += lines[i].length + 1;
        const line = lines[row];
        // 行首左侧、空行、文本上方空白：一律落在该行行首
        if (textX <= 0 || line.length === 0 || textY < 0) return offset;

        if (!this.measureCtx) {
            this.measureCtx = document.createElement('canvas').getContext('2d');
        }
        if (!this.measureCtx) return offset;
        this.measureCtx.font = style.font;
        if (this.measureCtx.measureText(line).width <= textX) {
            return offset + line.length;                       // 行尾右侧空白：落在行尾
        }
        // 二分：把像素横坐标换算为列号（最大的「前缀宽 <= textX」的字符数）
        let lo = 0, hi = line.length;
        while (lo < hi) {
            const mid = (lo + hi + 1) >> 1;
            if (this.measureCtx.measureText(line.slice(0, mid)).width <= textX) lo = mid;
            else hi = mid - 1;
        }
        return offset + lo;
    }

    /**
     * 刷新全部视图（高亮、行号、状态栏、滚动同步）
     * @private
     */
    private refresh(): void {
        this.renderHighlight();
        this.renderGutter();
        this.renderStatus();
        this.syncScroll();
    }

    /**
     * 渲染语法高亮层
     *
     * 末尾追加一个换行：textarea 以换行结尾时 pre 少算一行高度，
     * 两层底部会对不齐。
     * @private
     */
    private renderHighlight(): void {
        this.highlight.innerHTML = this.highlightPython(this.textarea.value) + '\n';
    }

    /**
     * 对 Python 代码做词法分段并输出带 class 的 HTML
     * @param code - 原始代码
     * @returns 高亮 HTML（已转义）
     * @private
     */
    private highlightPython(code: string): string {
        let html = '';
        let lastIndex = 0;
        // 模块级共享全局正则：exec 会推进 lastIndex，每轮使用前必须归零
        PY_TOKEN_PATTERN.lastIndex = 0;
        let match: RegExpExecArray | null;
        while ((match = PY_TOKEN_PATTERN.exec(code)) !== null) {
            html += escapeHtml(code.slice(lastIndex, match.index));
            let className = '';
            if (match[1] !== undefined) className = PY_TOKEN_CLASS.COMMENT;
            else if (match[2] !== undefined) className = PY_TOKEN_CLASS.STRING;
            else if (match[3] !== undefined) className = PY_TOKEN_CLASS.NUMBER;
            else if (match[4] !== undefined) className = PY_TOKEN_CLASS.DECORATOR;
            else if (PY_KEYWORDS.has(match[5])) className = PY_TOKEN_CLASS.KEYWORD;
            else if (PY_BUILTINS.has(match[5])) className = PY_TOKEN_CLASS.BUILTIN;
            const text = escapeHtml(match[0]);
            html += className ? `<span class="${className}">${text}</span>` : text;
            lastIndex = match.index + match[0].length;
        }
        html += escapeHtml(code.slice(lastIndex));
        return html;
    }

    /**
     * 渲染行号栏
     *
     * 行号用单个 `pre` 的文本一次写入，而不是逐行建节点：
     * 几千行的脚本也只产生一次 DOM 写入。
     * @private
     */
    private renderGutter(): void {
        const lineCount = this.textarea.value.split('\n').length;
        if (lineCount === this.gutterLineCount) return;
        this.gutterLineCount = lineCount;
        const numbers = new Array<string>(lineCount);
        for (let i = 0; i < lineCount; i++) numbers[i] = String(i + 1);
        this.gutter.textContent = numbers.join('\n');
    }

    /**
     * 刷新状态栏（光标所在行 / 列，均为 1 起始）
     * @private
     */
    private renderStatus(): void {
        const { selectionStart, value } = this.textarea;
        const before = value.slice(0, selectionStart);
        const line = before.split('\n').length;
        // 无换行时 lastIndexOf 返回 -1，列号恰为 selectionStart + 1
        const col = selectionStart - before.lastIndexOf('\n');
        this.status.textContent = `行 ${line}，列 ${col}`;
    }

    /**
     * 同步高亮层与行号栏的滚动位置到输入层
     * @private
     */
    private syncScroll(): void {
        this.highlight.scrollTop = this.textarea.scrollTop;
        this.highlight.scrollLeft = this.textarea.scrollLeft;
        this.gutter.scrollTop = this.textarea.scrollTop;
    }

    /**
     * 处理编辑按键：Tab / Shift+Tab 缩进、Enter 自动缩进
     * @param e - 键盘事件
     * @private
     */
    private handleKeydown(e: KeyboardEvent): void {
        if (e.key === 'Tab') {
            e.preventDefault();
            this.handleTab(e.shiftKey);
        } else if (e.key === 'Enter') {
            e.preventDefault();
            this.handleEnter();
        }
    }

    /**
     * 处理 Tab / Shift+Tab
     *
     * - 无选区 + Tab：光标处插入一级缩进；
     * - 无选区 + Shift+Tab：去掉光标所在行的行首缩进（最多一级），光标跟随；
     * - 有选区：选区涉及的整批行统一加 / 减一级缩进，操作后选区覆盖整个行块。
     *
     * @param outdent - true 为反缩进（Shift+Tab）
     * @private
     */
    private handleTab(outdent: boolean): void {
        const ta = this.textarea;
        const { selectionStart, selectionEnd, value } = ta;
        const hasSelection = selectionStart !== selectionEnd;

        if (!hasSelection && !outdent) {
            this.insertText(PYEditor.INDENT);
            return;
        }

        // 块范围：从首行行首开始；无选区时扩到行尾，有选区且结束于行首时该行不纳入
        const firstLineStart = value.lastIndexOf('\n', selectionStart - 1) + 1;
        let blockEnd: number;
        if (!hasSelection) {
            const lineEnd = value.indexOf('\n', selectionEnd);
            blockEnd = lineEnd === -1 ? value.length : lineEnd;
        } else {
            blockEnd = value[selectionEnd - 1] === '\n' ? selectionEnd - 1 : selectionEnd;
        }

        const lines = value.slice(firstLineStart, blockEnd).split('\n');
        let removedBeforeCursor = 0; // 无选区反缩进时，光标之前被删掉的空白数
        const nextLines = lines.map((line, index) => {
            if (outdent) {
                // 去掉行首最多一级宽度的空白（4 空格、1 个 tab、或不足一级的若干空格）
                const removed = line.startsWith(PYEditor.INDENT) ? PYEditor.INDENT.length
                    : line.startsWith('\t') ? 1
                    : line.match(/^ {1,3}/)?.[0].length ?? 0;
                if (index === 0 && !hasSelection) {
                    removedBeforeCursor = Math.min(removed, selectionStart - firstLineStart);
                }
                return line.slice(removed);
            }
            // 纯空白行不缩进，避免堆积无意义空格
            return line.trim() === '' ? line : PYEditor.INDENT + line;
        });

        const nextBlock = nextLines.join('\n');
        ta.value = value.slice(0, firstLineStart) + nextBlock + value.slice(blockEnd);
        if (hasSelection) {
            // 选区覆盖整个行块，便于连续按 Tab 逐级调整
            ta.selectionStart = firstLineStart;
            ta.selectionEnd = firstLineStart + nextBlock.length;
        } else {
            ta.selectionStart = ta.selectionEnd = selectionStart - removedBeforeCursor;
        }
        this.refresh();
        this.persistDebounced();
    }

    /**
     * 处理 Enter：换行并自动延续缩进
     *
     * 当前行（光标之前部分）以 `:` 结尾（允许尾部空白或注释）时，
     * 视为复合语句头（if / for / def / try ...），下一行再进一级。
     * @private
     */
    private handleEnter(): void {
        const ta = this.textarea;
        const before = ta.value.slice(0, ta.selectionStart);
        const line = before.slice(before.lastIndexOf('\n') + 1);
        const baseIndent = line.match(/^[ \t]*/)?.[0] ?? '';
        const extra = /:\s*(#.*)?$/.test(line) ? PYEditor.INDENT : '';
        this.insertText('\n' + baseIndent + extra);
    }

    /**
     * 在选区处插入文本（选区非空时先替换）
     *
     * 优先走 `document.execCommand('insertText')`：文本会进入浏览器原生的
     * 撤销栈，编辑器的 Ctrl+Z / Ctrl+Y 仍然可用。该 API 已废弃但现代浏览器
     * 均仍支持；失败时退化为直接赋值（原生撤销栈随之清空，属可接受代价）。
     * execCommand 成功时浏览器会派发 input 事件，此处的刷新与之幂等。
     * @param text - 要插入的文本
     * @private
     */
    private insertText(text: string): void {
        if (!document.execCommand('insertText', false, text)) {
            const ta = this.textarea;
            const start = ta.selectionStart;
            ta.value = ta.value.slice(0, start) + text + ta.value.slice(ta.selectionEnd);
            ta.selectionStart = ta.selectionEnd = start + text.length;
        }
        this.refresh();
        this.persistDebounced();
    }

    /**
     * 当前工作表对应的存储键；未打开文件（无名表）时返回 null，不持久化
     * @param sheetName - 工作表名，默认取当前活动表
     * @returns 完整存储键，或 null
     * @private
     */
    private storageKey(sheetName: string = this.data.activedSheetName): string | null {
        return sheetName ? `${PYEditor.STORAGE_KEY_PREFIX}${sheetName}` : null;
    }

    /**
     * 工作表切换：先把当前脚本立即写入旧表的键（清掉挂起的防抖），再载入新表的脚本
     *
     * 挂起的防抖若不清掉，会在 400ms 后把「新表的初始内容」写进「新表的键」之前，
     * 先把「旧表的内容」错误地写进「新表的键」（防抖触发时 activedSheetName 已是新表）。
     * @param oldName - 切换前的工作表名
     * @param newName - 切换后的工作表名（新表尚无存档时的默认脚本要以它命名首行）
     * @private
     */
    private handleSheetChanged(oldName: string, newName: string): void {
        if (this.saveTimer !== null) { clearTimeout(this.saveTimer); this.saveTimer = null; }
        const oldKey = this.storageKey(oldName);
        if (oldKey) {
            this.db.setSetting(oldKey, this.textarea.value, 'py-editor').catch(() => {});
        }
        this.dirty = false;
        // 新表无存档时展示它自己的默认模板（首行为新表名）；有存档则由 restore 覆盖
        this.setValue(PYEditor.defaultScript(newName));
        void this.restore();
    }

    /**
     * 内容变化后防抖写入 IndexedDB
     * @private
     */
    private persistDebounced(): void {
        const key = this.storageKey();
        if (!key) return;
        if (this.saveTimer !== null) clearTimeout(this.saveTimer);
        this.saveTimer = setTimeout(() => {
            this.saveTimer = null;
            this.db.setSetting(key, this.textarea.value, 'py-editor').catch(() => {
                // 持久化失败不影响编辑（下次变化会再次尝试）
            });
        }, PYEditor.SAVE_DEBOUNCE_MS);
    }

    /**
     * 迁移「内容与旧版默认模板完全一致」的历史脚本存档（全库只扫一次）
     *
     * 需求改为「首行标题取文件名」后，从未编辑过的历史存档首行仍是固定标题
     * `# SilkSpaces Python 脚本`，会被「已有存档优先」的 {@link restore} 原样载入，
     * 使文件名对应的新标题体现不出来。这里把这些存档的首行重写为
     * `# <文件名> Python 脚本`；判定用**逐字符相等**，用户真实编写过的内容不会被误改。
     * @returns 迁移完成的 Promise（共享，见 {@link PYEditor.legacyMigration}）
     * @private
     */
    private migrateLegacyDefaults(): Promise<void> {
        if (!PYEditor.legacyMigration) {
            PYEditor.legacyMigration = this.runLegacyMigration().catch(() => {
                // 迁移失败不阻断编辑；置回 null，下次 restore 时再试
                PYEditor.legacyMigration = null;
            });
        }
        return PYEditor.legacyMigration;
    }

    /**
     * 实际执行历史存档迁移：扫描 settings 中所有 `py-editor:script:*` 键
     * @private
     */
    private async runLegacyMigration(): Promise<void> {
        const all = await this.db.getAllSettings();
        for (const [key, value] of Object.entries(all)) {
            if (!key.startsWith(PYEditor.STORAGE_KEY_PREFIX)) continue;
            if (value !== PYEditor.LEGACY_DEFAULT_SCRIPT) continue;
            const name = key.slice(PYEditor.STORAGE_KEY_PREFIX.length);
            await this.db.setSetting(key, PYEditor.defaultScript(name), 'py-editor');
        }
    }

    /**
     * 从 IndexedDB 恢复上次的脚本内容
     *
     * 恢复是异步的：等待期间用户可能已经开始输入（`dirty` 置位），
     * 此时以用户内容为准，不再用存档覆盖。
     * @private
     */
    private async restore(): Promise<void> {
        const key = this.storageKey();
        if (!key) return;
        try {
            // 先迁移历史默认存档：否则旧模板会盖住文件名对应的新标题
            await this.migrateLegacyDefaults();
            const saved = await this.db.getSetting(key);
            if (this.dirty) return;
            if (typeof saved === 'string' && saved !== '') {
                this.setValue(saved);
            }
        } catch {
            // IndexedDB 不可用时保留默认示例脚本
        }
    }
}
