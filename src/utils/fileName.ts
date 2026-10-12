/*  # 文件名工具 TS 文件
    # 统一「显示文件名」与「工作表名」之间的换算规则
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

/**
 * 提取文件扩展名（小写，不含点）
 *
 * 无扩展名时返回空字符串；以点开头的隐藏文件名（如 `.gitignore`）视为无扩展名。
 * @param name - 文件名，如 `sheet1.XLSX`
 * @returns {string} 扩展名，如 `xlsx`
 * @example
 * extractExtension('sheet1.XLSX');  // 'xlsx'
 * extractExtension('归档');          // ''
 */
export function extractExtension(name: string): string {
    const dot = (name ?? '').lastIndexOf('.');
    return dot > 0 ? name.slice(dot + 1).toLowerCase() : '';
}

/**
 * 去掉文件扩展名，得到基础名（工作表名）
 *
 * 与 {@link extractExtension} 使用同一套判定：无扩展名时原样返回，
 * 因此「显示文件名」与「工作表名」的换算在全局保持一致。
 * @param name - 文件名，如 `销售数据.xlsx`
 * @returns {string} 基础名，如 `销售数据`
 * @example
 * stripExtension('销售数据.xlsx');  // '销售数据'
 * stripExtension('Sheet1');        // 'Sheet1'
 */
export function stripExtension(name: string): string {
    const dot = (name ?? '').lastIndexOf('.');
    return dot > 0 ? name.slice(0, dot) : (name ?? '');
}
