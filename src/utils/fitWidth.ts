/*  # 工具条宽度自适应判定 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

/**
 * 「装不下就折叠」的纯计算
 *
 * 之所以单独抽出来：判定逻辑（含边距、非法值兜底）要能脱离 DOM 单测 ——
 * 而真正的宽度来自 `scrollWidth` / `clientWidth`，在测试环境里是假的。
 * @module fitWidth
 */

/** 折叠判定额外预留的边距（px）：贴边也算装不下，避免出现 1px 的横向滚动条 */
export const COMPACT_MARGIN = 8;

/**
 * 内容是否装不下（需要折叠）
 *
 * - `needWidth` ≤ 0 或非有限值 → 视为「量不到」，不折叠（宁可挤，也不要无故切到紧凑模式）
 * - `availableWidth` ≤ 0 → 容器还没布局好，不折叠
 * @param {number} needWidth - 内容自然宽度（`scrollWidth`）
 * @param {number} availableWidth - 容器可用宽度（`clientWidth`）
 * @param {number} [margin] - 额外边距，默认 {@link COMPACT_MARGIN}
 * @returns {boolean} 需要折叠为紧凑模式时 true
 */
export function needsCompact(needWidth: number, availableWidth: number, margin: number = COMPACT_MARGIN): boolean {
    if (!Number.isFinite(needWidth) || !Number.isFinite(availableWidth)) return false;
    if (needWidth <= 0 || availableWidth <= 0) return false;
    return needWidth + Math.max(0, margin) > availableWidth;
}
