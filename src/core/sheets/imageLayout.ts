/*  # 图片锚点布局 TS 文件
    # 把解析器给出的「锚点 + EMU 偏移」换算成画布像素矩形
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import type { ExtractedImage, ImageAnchor } from '../../utils/parser/sheetParser';
import { EMU_PER_PIXEL } from '../constant';

export { EMU_PER_PIXEL };

/** 像素矩形（不含缩放系数，坐标系由调用方约定） */
export interface ImageRect {
    left: number;
    top: number;
    width: number;
    height: number;
}

/**
 * 列/行几何查询接口
 *
 * 只暴露「取某列左边界 / 某列宽 / 某行上边界 / 某行高」四个查询，
 * 便于在 Node 测试里用普通函数替代真实表头对象。
 */
export interface ImageLayoutGrid {
    /** 列号（0-based）→ 该列左边界像素（累计值，隐藏列已按紧凑排列处理） */
    colLeft(col: number): number;
    /** 列号（0-based）→ 该列宽度像素 */
    colWidth(col: number): number;
    /** 行号（0-based）→ 该行上边界像素 */
    rowTop(row: number): number;
    /** 行号（0-based）→ 该行高度像素 */
    rowHeight(row: number): number;
}

/** 视口窗口（**绝对**工作表坐标：left/top 即当前滚动偏移） */
export interface ViewportWindow {
    left: number;
    top: number;
    /** 视口宽度（逻辑像素） */
    width: number;
    /** 视口高度（逻辑像素） */
    height: number;
}

/** 定位结果：图片 + 视图坐标矩形（已减去滚动偏移，尚未乘缩放） */
export interface PlacedImage {
    image: ExtractedImage;
    rect: ImageRect;
}

/** 锚点 → 绝对坐标（像素） */
function anchorToPixel(anchor: ImageAnchor, grid: ImageLayoutGrid): { x: number; y: number } {
    return {
        x: grid.colLeft(anchor.col) + anchor.colOffset / EMU_PER_PIXEL,
        y: grid.rowTop(anchor.row) + anchor.rowOffset / EMU_PER_PIXEL,
    };
}

/**
 * 计算单张图片的**绝对**像素矩形（工作表坐标系，未减滚动偏移、未乘缩放）
 *
 * 尺寸优先级（与 OOXML 三种锚点一一对应）：
 * 1. `to` 存在（twoCellAnchor）→ 右下角由 `to` 决定，尺寸 = 右下 − 左上；
 * 2. 否则 `ext` 存在（oneCellAnchor）→ 尺寸直接取 `ext`（EMU 转像素）；
 * 3. 两者都没有 → 退化为「锚点单元格的尺寸」，保证任何输入都能画出一个可见矩形
 *    （absoluteAnchor 与残缺锚点走这一支）。
 *
 * 退化尺寸（`to` 与 `from` 重合、`ext` 为零）同样回落到单元格尺寸：
 * 零宽/零高的矩形在 `drawImage` 上无意义，在部分浏览器上还会抛异常。
 *
 * @param image - 提取出的图片（含锚点信息）
 * @param grid - 列/行几何查询
 * @returns {ImageRect} 绝对像素矩形
 */
export function computeImageRect(image: ExtractedImage, grid: ImageLayoutGrid): ImageRect {
    const start = anchorToPixel(image.anchor, grid);
    let width = 0;
    let height = 0;

    if (image.to) {
        const end = anchorToPixel(image.to, grid);
        width = end.x - start.x;
        height = end.y - start.y;
    } else if (image.ext) {
        width = image.ext.cx / EMU_PER_PIXEL;
        height = image.ext.cy / EMU_PER_PIXEL;
    }

    if (!(width > 0)) width = grid.colWidth(image.anchor.col);
    if (!(height > 0)) height = grid.rowHeight(image.anchor.row);

    return { left: start.x, top: start.y, width, height };
}

/**
 * 视口裁剪 + 坐标换算：只保留与视口相交的图片，并转成视图坐标
 *
 * 【性能约定，勿回退】返回值按视口过滤，因此渲染代价与「当前屏上的图片数」成正比，
 * 而不是与工作簿里图片总数成正比 —— 与单元格渲染走 `getVisibleCells()` 同一原则。
 *
 * @param images - 当前工作表全部图片
 * @param grid - 列/行几何查询
 * @param viewport - 视口窗口（绝对坐标）
 * @returns {PlacedImage[]} 视图坐标矩形（已减滚动偏移），可直接乘缩放后绘制
 */
export function layoutVisibleImages(
    images: readonly ExtractedImage[],
    grid: ImageLayoutGrid,
    viewport: ViewportWindow
): PlacedImage[] {
    const placed: PlacedImage[] = [];
    if (images.length === 0) return placed;

    const viewRight = viewport.left + viewport.width;
    const viewBottom = viewport.top + viewport.height;

    for (let i = 0; i < images.length; i++) {
        const image = images[i];
        const rect = computeImageRect(image, grid);
        // 先判相交再换算（相交判定比减法便宜，且能挡掉绝大多数屏外图片）
        if (rect.left >= viewRight || rect.top >= viewBottom) continue;
        if (rect.left + rect.width <= viewport.left || rect.top + rect.height <= viewport.top) continue;
        placed.push({
            image,
            rect: {
                left: rect.left - viewport.left,
                top: rect.top - viewport.top,
                width: rect.width,
                height: rect.height,
            },
        });
    }
    return placed;
}
