/*  # 工作表数据层职责拆分（K3）
    # 从 dataCollection.ts 原样迁出，仅调整类继承关系与成员可见性，逻辑零改动
*/
import type { ExtractedImage } from "../../utils/parser/sheetParser";
import { SheetStoreBase } from "./sheetStore";

/**
 * 内嵌图片装载基类（图片与单元格值分库存储在 IndexDB 的 images 对象仓库）
 */
export abstract class SheetImagesBase extends SheetStoreBase {
    /**
     * 从数据库装载当前工作表的图片（无记录时按「没有图片」处理）
     *
     * 读回的是外部数据，必须先做形状校验（见 {@link sanitizeImages}）：
     * 损坏的图片记录不允许把装载流程带崩。
     * @param {string} sheetName - 工作表名（存档键）
     * @returns {Promise<void>} 装载完成的 Promise
     * @private
     */
    protected async _loadImages(sheetName: string): Promise<void> {
        let raw: unknown[] = [];
        try {
            raw = await this.db.getImages(sheetName);
        } catch {
            // 图片库不可用（老库未升级 / 配额问题）不应影响工作表装载：按无图片处理
            raw = [];
        }
        this._images = SheetImagesBase.sanitizeImages(raw);
        this._imagesDirty = false;
    }

    /**
     * 校验并规范化从数据库读回的图片记录
     *
     * 只做「形状」校验（丢弃不合法项），不做任何解码：
     * - 二进制必须是 `Uint8Array`（结构化克隆保留类型）；
     * - 锚点必须存在且四个字段都是有限数。
     * 其余可选字段（`to`/`ext`/`name`）缺失时按缺省处理 —— 渲染侧本就有回落路径。
     * @param raw - 数据库读回的原始值
     * @returns {ExtractedImage[]} 可安全送入渲染层的图片数组
     */
    private static sanitizeImages(raw: unknown[]): ExtractedImage[] {
        if (!Array.isArray(raw)) return [];
        const result: ExtractedImage[] = [];
        for (const item of raw) {
            if (!item || typeof item !== 'object') continue;
            const candidate = item as Partial<ExtractedImage>;
            if (!(candidate.data instanceof Uint8Array) || candidate.data.length === 0) continue;
            if (typeof candidate.mimeType !== 'string') continue;
            const anchor = candidate.anchor;
            if (!anchor || !Number.isFinite(anchor.col) || !Number.isFinite(anchor.row)
                || !Number.isFinite(anchor.colOffset) || !Number.isFinite(anchor.rowOffset)) continue;
            result.push({
                data: candidate.data,
                mimeType: candidate.mimeType,
                anchor: {
                    col: anchor.col, row: anchor.row,
                    colOffset: anchor.colOffset, rowOffset: anchor.rowOffset,
                },
                to: candidate.to,
                ext: candidate.ext,
                name: candidate.name,
            });
        }
        return result;
    }
}
