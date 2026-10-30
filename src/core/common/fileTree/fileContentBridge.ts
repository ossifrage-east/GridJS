/*  # 文件内容桥契约（文件树 ↔ 文件存档）
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

/**
 * 文件内容搬迁计划项
 *
 * 树只认识**路径**，而文件内容存在外部（IndexedDB 的 `values` 库，键为
 * 「去根目录的完整路径」，如 `归档/销售数据.grid` → `归档/销售数据`）。
 * 因此「文件（或其祖先文件夹）的路径变了 / 被复制了 / 被删除了」这类副作用一律以
 * **路径**告知外部，由外部换算成存档键后执行。
 *
 * 之所以用路径而不是文件名：不同目录下的同名文件是两个独立的文件，
 * 各自拥有独立存档；只传文件名会让它们互相覆盖（历史实现的已知缺陷）。
 * @interface FileMovePlan
 */
export interface FileMovePlan {
    /** 源文件路径（含扩展名，相对根节点） */
    from: string;
    /** 目标文件路径（含扩展名，相对根节点） */
    to: string;
}

/**
 * 内容桥：把「结构变化」翻译成「文件内容的同步动作」
 *
 * 四个动作都由树在**语义发生变化的时刻**调用一次，调用方负责真正读写数据；
 * 参数一律是**路径**（相对根节点，含扩展名），不是文件名：
 * - {@link FileContentBridge.create}：新建文件（库中尚无存档时补一份空白存档）；
 * - {@link FileContentBridge.copy}：复制 / 粘贴（复制模式），含文件夹递归；
 * - {@link FileContentBridge.move}：路径变化（重命名、剪切粘贴、拖拽移动、
 *   以及**祖先文件夹改名**导致其下所有文件路径一起变化）；
 * - {@link FileContentBridge.remove}：删除（含文件夹递归）。
 * @interface FileContentBridge
 */
export interface FileContentBridge {
    /**
     * 补建文件存档：新建出来的文件在数据库中还没有存档时写入一份空白存档
     *
     * 只补不覆盖：这条路径也会在「登记已打开的文件」时被调用，此时内容刚由打开器写入。
     * @param paths - 新增文件的路径集合
     */
    create?: (paths: string[]) => void;
    /** 复制文件内容：把 `from` 的存档复制到 `to` */
    copy?: (plans: FileMovePlan[]) => void;
    /** 搬迁文件内容：把 `from` 的存档搬到 `to`（源键不再保留） */
    move?: (plans: FileMovePlan[]) => void;
    /**
     * 删除文件内容：清理这些路径对应文件在数据库中的存档
     * @param paths - 被删除文件的路径集合
     */
    remove?: (paths: string[]) => void;
}
