/*  # 图片绘制层 TS 文件
    # 异步解码 + 位图缓存 + 视口内绘制（主画布的图片通道）
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import type { ExtractedImage } from '../../utils/parser/sheetParser';
import type { PlacedImage } from './imageLayout';

/**
 * 浏览器可直接解码的图片 MIME
 *
 * `emf`/`wmf` 是 Windows 矢量图元文件，浏览器 `createImageBitmap` 一律拒绝；
 * `tiff` 同样不在支持列表内。这些类型不去尝试解码（每次失败都会走一遍异常路径）。
 */
const DECODABLE_MIME = new Set([
    'image/png', 'image/jpeg', 'image/gif', 'image/bmp', 'image/webp', 'image/svg+xml',
]);

/** 解码状态（非位图的两种终态） */
type DecodeState = 'pending' | 'error';

/** 缓存项：位图 / 解码中 / 解码失败 */
type CacheEntry = ImageBitmap | DecodeState;

export interface SheetImageLayerOptions {
    /**
     * 位图解码完成时的回调（请求重绘一帧）
     *
     * 解码是异步的：首帧只发起解码、拿不到位图，这一帧的图片是空的。
     * 解码完成后必须让画布再画一次，否则要等到下一次滚动/选中才「突然出现」。
     */
    requestRedraw: () => void;
    /** 单帧最多发起多少个解码（默认 16；避免一口气解码上百张图抢占主线程） */
    maxDecodePerFrame?: number;
}

/**
 * 工作表图片绘制层
 *
 * 职责边界：只负责「位图解码缓存」与「把定位好的矩形画到 2D 上下文」，
 * 几何换算在 {@link imageLayout} 里（纯函数，可在 Node 里测）。
 *
 * ## 为什么缓存挂在图片对象上
 * 缓存键就是 `ExtractedImage` 实例本身（`WeakMap`）：
 * - 同一张图在每一帧都会被问到，必须只解码一次；
 * - 切换文件后旧图片对象不可达，`WeakMap` 让位图随 GC 一起回收，不需要手工失效；
 * - 不需要给图片编 id，也就不会出现「id 复用导致串图」的问题。
 *
 * @class SheetImageLayer
 */
export class SheetImageLayer {
    /** 位图缓存（键 = 图片对象本身，见类注释） */
    private readonly cache = new WeakMap<ExtractedImage, CacheEntry>();
    /** 配置项 */
    private readonly options: SheetImageLayerOptions;

    constructor(options: SheetImageLayerOptions) {
        this.options = options;
    }

    /**
     * 绘制视口内的图片
     *
     * 【性能约定，勿回退】入参已由 {@link layoutVisibleImages} 按视口裁剪过，
     * 因此这里只与「屏上图片数」相关，与工作簿图片总数无关。
     * 未解码完成 / 解码失败的图片直接跳过（首帧为空是正常的，解码完成后会请求重绘）。
     *
     * @param ctx - 主画布上下文（物理像素坐标系，与单元格绘制一致）
     * @param placed - 视口内的图片与视图坐标矩形
     * @param zoom - 当前缩放比例
     * @param dpr - `window.devicePixelRatio`
     */
    public draw(
        ctx: CanvasRenderingContext2D,
        placed: readonly PlacedImage[],
        zoom: number,
        dpr: number
    ): void {
        if (placed.length === 0) return;
        const budget = this.options.maxDecodePerFrame ?? 16;
        let started = 0;

        ctx.save();
        for (let i = 0; i < placed.length; i++) {
            const { image, rect } = placed[i];
            const entry = this.cache.get(image);
            if (entry === undefined) {
                // 首帧：只发起解码，不绘制
                if (started >= budget) continue;
                started++;
                this.decode(image);
                continue;
            }
            if (typeof entry === 'string') continue;   // pending / error

            const width = Math.max(1, Math.round(rect.width * zoom * dpr));
            const height = Math.max(1, Math.round(rect.height * zoom * dpr));
            ctx.drawImage(
                entry,
                Math.round(rect.left * zoom * dpr),
                Math.round(rect.top * zoom * dpr),
                width,
                height
            );
        }
        ctx.restore();
    }

    /**
     * 发起一次异步解码（同一次解码只会被发起一次）
     * @param image - 待解码图片
     * @private
     */
    private decode(image: ExtractedImage): void {
        this.cache.set(image, 'pending');

        const create = (globalThis as unknown as {
            createImageBitmap?: (source: Blob) => Promise<ImageBitmap>;
        }).createImageBitmap;
        if (typeof create !== 'function' || typeof Blob === 'undefined' || !DECODABLE_MIME.has(image.mimeType)) {
            // 环境不支持（Node 测试 / 老浏览器）或格式不可解码：直接记终态，
            // 不要每次绘制都重试一遍必然失败的路径
            this.cache.set(image, 'error');
            return;
        }

        let promise: Promise<ImageBitmap>;
        try {
            promise = create(new Blob([image.data as unknown as BlobPart], { type: image.mimeType }));
        } catch {
            this.cache.set(image, 'error');
            return;
        }

        promise.then((bitmap) => {
            this.cache.set(image, bitmap);
            this.options.requestRedraw();
        }).catch(() => {
            // 字节损坏 / 解码失败：记终态，不阻塞后续绘制
            this.cache.set(image, 'error');
        });
    }
}
