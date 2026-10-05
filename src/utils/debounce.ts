/*  # 防抖函数 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

// 防抖函数
export const debounce = <T extends (...args: any[]) => void>(func: T, wait: number): T => {
    let timeout: number | null = null;
    return ((...args: any[]) => {
        if (timeout) {
            clearTimeout(timeout);
        }
        timeout = window.setTimeout(() => {
            func(...args);
        }, wait);
    }) as T;
}