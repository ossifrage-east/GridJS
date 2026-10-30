/*  # 编辑契约层 · 命令总线 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目描述: 基于 HTML5 Canvas 的 Web 电子表格组件，无需服务端支持即可在浏览器中创建、读取、写入 Excel 文件。
    # 项目目标: 让工具栏分区只认识「命令总线」，不认识任何具体宿主（表格 / 文档）。
    # 本文件职责: 实现 {@link CommandBus} —— 宿主切换（bind）、命令派发（dispatch）、
    #             状态查询（snapshot / canExecute）与状态变化广播（EDIT_STATE_CHANGED）。
    #             它本身**不含任何表格或文档语义**，只是把「谁在编辑」这件事委托给当前宿主。
    # 依赖方向: contract 内部 + core/constant（仅为转发按钮需要的编辑态事件名，constant 是叶子模块）。
    #           **不得 import sheet / document / shell 下的任何模块。**
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { EventEmitter } from '../../utils/eventEmitter';
import { DataEvents } from '../constant';
import {
    CommandId,
    CommandPayload,
    EDIT_STATE_CHANGED,
    EditHost,
    EditStateSnapshot,
    EMPTY_EDIT_STATE,
    WorkspaceKind,
} from './editHost';

/**
 * 需要原样转发给按钮的事件名
 *
 * `BtnBase` 的 `isDisabledOnEditting` 特性历史上直接订阅 `DataCollection` 的
 * `IS_EDITTING_CHANGED`。改成订阅总线后，总线必须**用同一个事件名**转发，
 * 否则按钮的「编辑时置灰」会静默失效（不报错、只是行为没了）。
 */
const FORWARDED_TO_BUTTONS = [DataEvents.IS_EDITTING_CHANGED] as const;

/**
 * 命令总线
 *
 * ## 职责边界
 *
 * 它**只做转发与解耦**，不缓存业务状态、不做命令校验之外的任何决策：
 * - 命令是否可执行由宿主回答（{@link EditHost.canExecute}）
 * - 状态从哪来由宿主决定（{@link EditHost.snapshot}）
 * - 「状态变了」由宿主发事件，总线只负责改个名广播出去
 *
 * ## 为什么它继承 EventEmitter 而不是自己管监听表
 *
 * 项目里 `EventEmitter` 已经是统一的订阅原语（`on` / `off` / `emit` / `once`），
 * 复用它可以顺带让总线**直接充当 `BtnBase` 的 `data` 字段**
 * （`EditingStateSource` 结构匹配），按钮侧一行都不用改。
 *
 * @example
 * ```ts
 * const bus = new CommandBus();
 * bus.bind(sheetHost);                     // 工作区切换时换宿主
 * bus.dispatch('toggleBold', { toggle: true });
 * const { char } = bus.snapshot();
 * console.log(char.bold);
 * ```
 */
export class CommandBus extends EventEmitter {
    /** 当前宿主；未绑定时所有命令静默忽略 */
    private host: EditHost | null = null;

    /**
     * 宿主事件监听器（用于切换宿主时精确退订，避免监听泄漏）
     *
     * 必须保存引用：`EventEmitter.off(event)` 不带 handler 会清掉该事件的**全部**监听，
     * 若用匿名函数订阅、退订时又传不了同一个引用，反复切换宿主会让监听表无限增长
     * （表现为切几次之后一次点击触发多次重绘）。
     */
    private hostListeners: Array<{ event: string; handler: (...args: any[]) => void }> = [];

    /** 当前宿主类型；未绑定时为 `null`（UI 可据此决定是否显示工具栏内容） */
    public get kind(): WorkspaceKind | null {
        return this.host ? this.host.kind : null;
    }

    /**
     * 是否处于字符级编辑态
     *
     * 供 UI 层在点击瞬间判断分支（如格式刷、复制、粘贴要区分字符级与单元格级）。
     * 该值**随快照一起取**，不单独维护一份镜像状态 —— 镜像会有「谁先更新」的时序问题。
     * @returns 编辑态
     */
    public get isEditting(): boolean {
        return this.host ? this.host.snapshot().editing : false;
    }

    /**
     * 切换宿主
     *
     * 工作区切换（表格 ⇄ 文档）时调用。会先解除对旧宿主的全部监听，再订阅新的，
     * 因此反复切换不会累积监听。订阅完成后主动广播一次状态变化，让工具栏立即按新宿主刷新外观。
     * @param host - 新宿主
     */
    public bind(host: EditHost): void {
        this.unbind();
        this.host = host;

        // 状态变化：宿主的通用事件 → 总线事件（改名，外壳层不认识宿主的内部事件名）
        this.subscribe(EDIT_STATE_CHANGED, () => {
            this.emit(EDIT_STATE_CHANGED);
        });
        // 编辑态：按钮需要的原始事件名，原样转发
        for (const event of FORWARDED_TO_BUTTONS) {
            this.subscribe(event, (...args: any[]) => {
                this.emit(event, ...args);
            });
        }

        // 宿主刚挂上：立即广播一次，让各分区把按钮刷成新宿主的真实状态
        this.emit(EDIT_STATE_CHANGED);
    }

    /**
     * 解除当前宿主绑定
     *
     * 只退订自己注册过的监听（按保存的 handler 引用精确退订），不影响其它订阅者。
     */
    public unbind(): void {
        if (this.host) {
            for (const { event, handler } of this.hostListeners) {
                this.host.stateSource.off(event, handler);
            }
        }
        this.hostListeners = [];
        this.host = null;
    }

    /**
     * 派发命令
     *
     * 未绑定宿主、或宿主声明该命令不可执行时静默返回 —— 界面上对应按钮此时本就应当是置灰的，
     * 这里再挡一次是为了防止「快捷键 / 菜单」这类绕过按钮状态的入口误触发。
     * @param cmd - 命令名
     * @param payload - 命令载荷
     */
    public dispatch(cmd: CommandId, payload?: CommandPayload): void {
        if (!this.host || !this.host.canExecute(cmd)) return;
        this.host.execute(cmd, payload);
    }

    /**
     * 该命令当前是否可执行
     * @param cmd - 命令名
     * @returns 是否可执行；未绑定宿主时恒为 false
     */
    public canExecute(cmd: CommandId): boolean {
        return this.host ? this.host.canExecute(cmd) : false;
    }

    /**
     * 取当前状态快照
     *
     * 未绑定宿主时返回共享的空快照（{@link EMPTY_EDIT_STATE}），
     * 使工具栏在没有宿主的情况下也能安全刷新（统一显示为「全部未激活」）。
     * @returns 状态快照
     */
    public snapshot(): EditStateSnapshot {
        return this.host ? this.host.snapshot() : (EMPTY_EDIT_STATE as EditStateSnapshot);
    }

    /**
     * 主动通知「状态可能已变化」
     *
     * 宿主内部发生的变化若不在其订阅的事件范围内（例如工具栏自己改完样式后想让别处同步），
     * 可由宿主或装配层调用本方法手动广播一次。
     */
    public notifyStateChanged(): void {
        this.emit(EDIT_STATE_CHANGED);
    }

    /**
     * 订阅宿主事件并记录 handler 引用（供 {@link unbind} 精确退订）
     * @param event - 事件名
     * @param handler - 回调
     */
    private subscribe(event: string, handler: (...args: any[]) => void): void {
        if (!this.host) return;
        this.host.stateSource.on(event, handler);
        this.hostListeners.push({ event, handler });
    }
}
