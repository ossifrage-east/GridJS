/*  # 文档域 · 标尺 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目目标: 在文档编辑区的**上方与左侧**画两把标尺（与 Word 的「标尺」同一形态）。
    # 本文件职责: 只做一件事 —— 按「纸张长度（毫米）+ 缩放比例」把刻度画进给定的元素。
    #             **不认识块树、不认识存储**：它拿到的全部输入就是两个数字。
    # 为什么单独一个文件: 标尺的刻度算法（毫米 → 像素、主次刻度、标签密度）与
    #             「标尺挂在哪个容器里」是两件事；混进 docView 会让那个已经很长的文件
    #             再多出两位数的行，而这里的代码是可以被单元测试直接喂数字验证的。
    # 依赖方向: utils/dom（叶子工具）。**不 import document/** 下任何模块。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { createDiv } from '../../utils/dom';

/**
 * 1 毫米对应的 CSS 像素（96dpi）
 *
 * 与 `blockTree.DOC_PAPER_SIZES` 同一口径（A4 = 794×1123px 即 210×297mm），
 * 纸张尺寸与标尺刻度必须共用这个除数，否则「标尺上的 10cm」和「纸上的 10cm」会错开。
 */
export const MM_PER_PX = 25.4 / 96;

/** 1 毫米 = 多少 CSS 像素 */
export const PX_PER_MM = 96 / 25.4;

/** 次刻度间距（毫米）：5mm 一格，与 Word 标尺的最小格一致 */
const MINOR_STEP_MM = 5;

/** 主刻度间距（毫米）：10mm 一格（每 1cm 一条长线） */
const MAJOR_STEP_MM = 10;

/** 数字标签间距（毫米）：每 2cm 标一个数，密到 1cm 会糊成一片 */
const LABEL_STEP_MM = 20;

/**
 * 画一把标尺
 *
 * 刻度用**绝对定位**而不是 flex 均分：长度不是格距的整数倍时（A4 的 210mm 恰好是，
 * 但 Letter 的 215.9mm 不是），均分会让整把尺子悄悄伸长，与纸边对不齐。
 *
 * **两把尺都既有刻度也有数字**。早期版本「横向只画数字、纵向只画刻度」——
 * 横向的刻度是绝对定位在刻度条底边的，而刻度条当时高度为 0（子元素全绝对定位），
 * 于是刻度被 `overflow: hidden` 裁在尺子上边界之外，看起来就是「只有数字没有刻度」。
 * 修法是让横向刻度条撑满尺子高度（见 document.scss 的 `.doc-ruler-h .doc-ruler-inner`），
 * 本函数只需保证**两个方向都下发数字**。
 * 纵向的数字旋转 90°（竖排数字在 18px 窄条里横着放不下三位数），旋转由样式表完成。
 * @param el - 标尺元素（横向标尺需由调用方定高，纵向定宽）
 * @param lengthPx - 纸张长度（**未缩放**的 CSS 像素；横向为纸宽，纵向为纸高×页数）
 * @param zoom - 缩放比例（标尺与纸一样跟着缩放，才谈得上「对齐」）
 * @param horizontal - 横向标尺为 true
 * @returns 画出的刻度条数（供测试断言）
 */
export function buildRuler(el: HTMLElement, lengthPx: number, zoom: number, horizontal: boolean): number {
    // 先清空：缩放 / 纸张变化都会重画，追加会留下一堆旧刻度
    while (el.firstChild) el.removeChild(el.firstChild);
    const lengthMm = lengthPx * MM_PER_PX;
    const total = lengthPx * zoom;
    // 另一维必须显式清空：切方向复用同一个元素时，旧的定长会残留在内联样式上
    if (horizontal) {
        el.style.width = `${total}px`;
        el.style.height = '';
    } else {
        el.style.height = `${total}px`;
        el.style.width = '';
    }
    if (!isFinite(lengthMm) || lengthMm <= 0) return 0;

    const count = Math.floor(lengthMm / MINOR_STEP_MM);
    for (let i = 0; i <= count; i++) {
        const mm = i * MINOR_STEP_MM;
        const pos = mm * PX_PER_MM * zoom;
        const major = mm % MAJOR_STEP_MM === 0;
        const tick = createDiv({ className: `doc-ruler-tick${major ? ' doc-ruler-major' : ''}` });
        if (horizontal) tick.style.left = `${pos}px`;
        else tick.style.top = `${pos}px`;
        el.appendChild(tick);
        // 两个方向都标数字（标签用厘米：标尺上的数字是「第几厘米」）
        if (mm > 0 && mm % LABEL_STEP_MM === 0) {
            const label = createDiv({
                className: 'doc-ruler-label',
                textContent: String(mm / 10),
            });
            if (horizontal) label.style.left = `${pos}px`;
            else label.style.top = `${pos}px`;
            el.appendChild(label);
        }
    }
    return count + 1;
}
