/**
 * Excel 解析 Worker 入口（decode + extract + 列式打包）
 *
 * 只做纯计算的三步：
 *   1. `decode`  — ZIP 解包（inflate）/ OLE2 定位 Workbook 流
 *   2. `extract` — 流式扫描 XML，提取单元格/样式/合并/行列尺寸/图片
 *   3. `pack`    — 列式打包（{@link packSheets}）：平行数组 + 样式表去重 + 文本池
 *
 * 是否读图片由调用方经 `options.readImages` 决定（默认 false）：
 * 嵌入图片是不可压缩的大块二进制，不需要时必须**在解压阶段**就跳过
 * （媒体条目在 decode 里按目录过滤，之后无从补救）。
 * 打开文件路径会显式传 `readImages: true`（见 `common/fileOpener.ts`），
 * 因为位图渲染通道已经建好（`sheets/sheetImageLayer.ts`）。
 *
 * 第三步 `generate`（映射为 Cell/Char 类实例）留在主线程：类实例跨线程
 * 传输会丢原型，且 `Cell.getText()` 等方法在接收侧不可用。
 *
 * **为什么回传的是列式打包而不是 `ExtractedSheet[]` 原样（性能约定，勿回退）**：
 * 结构化克隆要逐对象重建，实测 6.3 万格（6.2 万格带样式）的 `ExtractedCell[]`
 * 在主线程反序列化要 489 ms；打成列式后 typed array 走 transfer，
 * 只剩 48 ms，常驻内存也从 58 MB 降到 12 MB。打包本身的 ~91 ms 留在 Worker 里，
 * 主线程完全无感 —— 这是「打开大文件不再卡死」的关键一步。
 *
 * 本文件由 webpack 的独立 worker 配置打包为固定文件名 `parser.worker.js`
 * （见 webpack.config.js）。
 *
 * 传输约定：文件字节以 `ArrayBuffer` 转移进来；打包结果的 typed array 与
 * 图片二进制以 `ArrayBuffer` 转移出去（零拷贝）。
 *
 * @module parser.worker
 */

import { SheetParser } from './sheetParser';
import { packSheets, packTransferables } from './cellSource';
import type { ParserWorkerRequest, ParserWorkerResponse } from './parserClient';

/**
 * Worker 全局作用域的最小类型视图
 *
 * 不引入 `WebWorker` lib（会与项目的 DOM lib 冲突），
 * 仅声明实际用到的两个成员。
 */
const ctx = self as unknown as {
    onmessage: ((event: MessageEvent<ParserWorkerRequest>) => void) | null;
    postMessage: (message: ParserWorkerResponse, transfer?: Transferable[]) => void;
};

/**
 * 把转移进来的字节包装成 `decode` 所需的 File 视图
 *
 * `decode` 只用到 `name`（分派策略）与 `arrayBuffer()`，故无需真实 File。
 */
function fileLikeFrom(name: string, buffer: ArrayBuffer): File {
    return {
        name,
        size: buffer.byteLength,
        arrayBuffer: async () => buffer,
    } as unknown as File;
}

ctx.onmessage = (event: MessageEvent<ParserWorkerRequest>) => {
    const { id, name, buffer, options } = event.data;

    void (async () => {
        try {
            const parser = new SheetParser();
            // options 必须传给 decode：媒体条目（xl/media/**）在解压阶段就被跳过，
            // 只有 readImages === true 才读（见 SheetParser.decode）
            const decoded = await parser.decode(fileLikeFrom(name, buffer), options);
            const sheets = await parser.extract(decoded, options);

            // 列式打包：把「逐格普通对象」换成「平行 typed array + 样式表 + 文本池」，
            // 回主线程的克隆成本下降一个数量级（详见模块头注释）
            const packed = packSheets(sheets);

            // typed array 与图片二进制一并转移，避免结构化克隆时的二次拷贝
            const transfer = packTransferables(packed);

            ctx.postMessage({ id, sheets: packed }, transfer);
        } catch (err) {
            ctx.postMessage({
                id,
                error: err instanceof Error ? err.message : String(err),
            });
        }
    })();
};
