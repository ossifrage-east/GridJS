/*  # 文档域 · 纯文本编解码 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目目标: 让文档能在「块树 ⇄ 纯文本」之间往返，使 .text 原生文件有可读的落盘形态。
    # 本文件职责: 两个纯函数 —— 块树投影为纯文本、纯文本还原为块树。
    # 依赖方向: 只 import 同目录 blockTree（纯数据层）。
    # 有损说明: 纯文本没有样式概念，往返后**字符样式全部丢失**（这是格式本身的限制，
    #           不是实现缺陷）；块级类型同理。要保样式请走 Markdown 或 docx 通道。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { DocBlock, blockText, createBlock, normalizeBlocks } from '../blockTree';

/**
 * 列表项的纯文本前缀
 *
 * 纯文本通道刻意**不加** `- ` / `1. ` 这类前缀：它是「给别的程序读的正文」，
 * 加了前缀会让粘贴到别处时多出一层符号。列表语义请走 Markdown 通道。
 */

/**
 * 块树 → 纯文本
 *
 * 每个可见块一行，行间用 `\n` 分隔；表格块按「单元格用制表符分隔、行用换行分隔」
 * 展开 —— 与 `clipboardManager.cellsToTSV` 的 TSV 约定一致，使「文档里的表格」
 * 复制出去与「表格选区」复制出去得到同一种文本。
 * @param blocks - 块树
 * @returns 纯文本
 */
export function blocksToText(blocks: DocBlock[]): string {
    const lines: string[] = [];
    collect(blocks, lines);
    return lines.join('\n');
}

/**
 * 递归收集文本行
 * @param blocks - 块集合
 * @param lines - 输出缓冲
 */
function collect(blocks: DocBlock[], lines: string[]): void {
    for (let i = 0; i < blocks.length; i++) {
        const block = blocks[i];
        if (!block) continue;
        switch (block.type) {
            case 'hr':
                lines.push('----');
                break;
            case 'table':
                if (Array.isArray(block.rows)) {
                    for (let r = 0; r < block.rows.length; r++) {
                        const row = block.rows[r] ?? [];
                        lines.push(row.map((cell: DocBlock) => blocksToText(cell.children ?? [cell])).join('\t'));
                    }
                }
                break;
            case 'code':
                if (Array.isArray(block.children)) collect(block.children, lines);
                else lines.push(blockText(block));
                break;
            default:
                lines.push(blockText(block));
        }
    }
}

/**
 * 纯文本 → 块树（每个换行一段）
 *
 * 空行保留为**空段落**而不是被吞掉：Markdown 往返时「两个空行 = 段落间隔」的
 * 语义全靠它，吞掉空行会让「块树 → 文本 → 块树」改变段落间距。
 * @param text - 纯文本
 * @returns 块树（至少一个块）
 */
export function textToBlocks(text: string): DocBlock[] {
    const source = (text ?? '').replace(/\r\n?/g, '\n');
    const blocks: DocBlock[] = source.split('\n').map((line) =>
        createBlock({ type: 'paragraph', text: line })
    );
    return normalizeBlocks(blocks);
}
