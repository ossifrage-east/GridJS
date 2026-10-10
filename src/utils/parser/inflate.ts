/**
 * 原生 DEFLATE（RFC 1951）解压模块
 *
 * ZIP 容器（xlsx 本质）使用 DEFLATE 压缩其内部 XML。本模块提供两种解压途径：
 *
 * 1. **浏览器原生 `DecompressionStream('deflate-raw')`**（首选）
 *    现代浏览器（Chrome 80+/Edge 80+/Firefox 113+/Safari 16.4+）内置，
 *    由浏览器原生代码实现，速度快、代码量为零。使用条件探测，不可用时自动回退。
 *
 * 2. **纯 TypeScript 实现的 inflate 算法**（回退）
 *    采用经典的 *puff* 参考实现思路：规范霍夫曼（canonical Huffman）解码 +
 *    固定/动态霍夫曼树构建 + LZ77 回溯拷贝。约 200 行代码，无任何第三方依赖，
 *    可在 IE10+ / 老旧 Safari 等缺少原生 API 的环境下工作。
 *
 * @module inflate
 * @author 东方鹗
 */

/**
 * 解压 raw DEFLATE 数据（无 zlib/gzip 包装）
 *
 * 优先走浏览器原生 `DecompressionStream`，不可用时回退纯 JS 实现。
 *
 * @param {Uint8Array} data - 压缩数据
 * @param {number} [expectedSize] - 已知解压后大小（ZIP 本地头/中央目录提供），
 *        用于预分配缓冲、减少扩容。传入后即使原生路径失败也可借助其校验。
 * @returns {Promise<Uint8Array>} 解压后的数据
 * @throws {Error} 数据流损坏或格式非法
 */
export async function inflateRaw(data: Uint8Array, expectedSize?: number): Promise<Uint8Array> {
    // —— 路径 1：浏览器原生 DecompressionStream ——
    if (typeof DecompressionStream !== 'undefined') {
        try {
            return await inflateWithNativeStream(data);
        } catch (e) {
            // 原生路径失败（如部分环境的 deflate-raw 实现有缺陷）→ 落到纯 JS 路径
            console.warn('[inflate] 原生 DecompressionStream 解压失败，回退纯 JS 实现：', e);
        }
    }

    // —— 路径 2：纯 TypeScript 实现 ——
    return inflatePure(data, expectedSize);
}

// ════════════════════════════════════════════════════════════
//  原生 DecompressionStream 路径
// ════════════════════════════════════════════════════════════

/**
 * 使用浏览器原生 DecompressionStream 解压 raw DEFLATE
 *
 * 原理：把压缩数据包成一条不可读流（ReadableStream），交给 `DecompressionStream`
 * 转换，再通过 reader 逐块读出并拼接。
 *
 * @param {Uint8Array} data - 压缩数据
 * @returns {Promise<Uint8Array>} 解压后的数据
 * @private
 */
async function inflateWithNativeStream(data: Uint8Array): Promise<Uint8Array> {
    // 直接构造 ReadableStream 推入整块压缩数据，避免经 Blob 复制一份内存
    const source = new ReadableStream<Uint8Array>({
        start(controller) {
            controller.enqueue(data);
            controller.close();
        }
    });

    // DecompressionStream 在部分TS lib 声明中 writable 泛型为 BufferSource，与 pipeThrough 的
    // ReadableWritablePair<Uint8Array> 签名不兼容；此处按标准 Web Streams 行为断言即可
    const transform = new DecompressionStream('deflate-raw') as unknown as
        ReadableWritablePair<Uint8Array, Uint8Array>;
    const stream = source.pipeThrough(transform);
    const reader = stream.getReader();

    const chunks: Uint8Array[] = [];
    let total = 0;

    for (; ;) {
        const { done, value } = await reader.read();
        if (done) break;
        if (value) {
            chunks.push(value);
            total += value.length;
        }
    }

    // 合并所有分块
    const result = new Uint8Array(total);
    let offset = 0;
    for (const chunk of chunks) {
        result.set(chunk, offset);
        offset += chunk.length;
    }
    return result;
}

// ════════════════════════════════════════════════════════════
//  纯 TypeScript inflate 实现（puff 算法思路）
// ════════════════════════════════════════════════════════════

/** 长度码 257~285 的基值表：码长 → 基础长度 */
const LENGTH_BASE: number[] = [
    3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31,
    35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258
];

/** 长度码 257~285 的额外位���表 */
const LENGTH_EXTRA: number[] = [
    0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2,
    3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0
];

/** 距离码 0~29 的基值表：码长 → 基础距离 */
const DISTANCE_BASE: number[] = [
    1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193,
    257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145,
    8193, 12289, 16385, 24577
];

/** 距离码 0~29 的额外位数表 */
const DISTANCE_EXTRA: number[] = [
    0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6,
    7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13
];

/** 码长字母表顺序（动态块的 code length alphabet 传输顺序） */
const CODE_LENGTH_ORDER: number[] = [
    16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15
];

/** 霍夫曼解码表：count 为各码长的码字数量，symbol 为按码长排序的符号表 */
interface HuffmanTable {
    count: number[];
    symbol: number[];
}

/**
 * 规范化霍夫曼表（canonical Huffman）
 *
 * DEFLATE 的霍夫曼码是规范化的：同一码长的码字按符号值递增排列。
 * 因此只需记录「每个码长有多少个码字」和「按码长排序的符号序列」，
 * 解码时逐位累加即可还原符号，无需构建完整树。
 */
class Huffman {
    /** count[len] = 码长为 len 的码字数量（索引 0 未使用，恒为 0） */
    public count: number[] = new Array(16).fill(0);
    /** symbol = 按码长升序、同码长内按符号值升序排列的符号表 */
    public symbol: number[] = [];

    /**
     * 依据各符号的码长数组构建规范霍夫曼表
     * @param {Uint8Array} lengths - 每个符号的码长（值为 0 表示该符号未使用）
     * @param {number} n - 参与构建的符号数量
     */
    constructor(lengths: Uint8Array, n: number) {
        for (let i = 0; i < n; i++) {
            this.count[lengths[i]]++;
        }
        // 码长为 0 的符号不参与编码，需从计数中剔除
        this.count[0] = 0;

        // 计算每个码长的起始偏移（offs[len] 表示该码长在 symbol 表中的起始位置）
        const offs = new Array(16).fill(0);
        for (let len = 1; len < 15; len++) {
            offs[len + 1] = offs[len] + this.count[len];
        }

        // 按偏移填入符号
        for (let sym = 0; sym < n; sym++) {
            if (lengths[sym] !== 0) {
                this.symbol[offs[lengths[sym]]++] = sym;
            }
        }
    }
}

/**
 * 位流读取器（DEFLATE 的比特序为 LSB 优先）
 *
 * DEFLATE 的比特流不是字节对齐的：位从字节的最低位开始输出，
 * 跨字节时继续向高位推进。因此需要维护一个位缓冲区与已读位数。
 */
class BitStream {
    private data: Uint8Array;
    /** 数据中的字节位置 */
    private pos: number = 0;
    /** 位缓冲区 */
    private bitBuf: number = 0;
    /** 缓冲区中已填充的有效位数 */
    private bitCnt: number = 0;
    /** 自流开始以来已消费的比特总数（用于计算字节对齐位置） */
    private consumed: number = 0;

    constructor(data: Uint8Array) {
        this.data = data;
    }

    /**
     * 读取指定数量的比特（LSB 优先）
     * @param {number} need - 需要的比特数（1~16）
     * @returns {number} 读取到的数值（无符号）
     */
    public bits(need: number): number {
        let val = this.bitBuf;
        // 位缓冲区不足时，从后续字节补足
        while (this.bitCnt < need) {
            if (this.pos >= this.data.length) {
                throw new Error('DEFLATE 数据意外结束');
            }
            val |= this.data[this.pos++] << this.bitCnt;
            this.bitCnt += 8;
        }
        this.bitBuf = val >>> need;
        this.bitCnt -= need;
        this.consumed += need;
        return val & ((1 << need) - 1);
    }

    /**
     * 丢弃当前字节内未使用的位，对齐到下一个字节边界
     *
     * stored 块的 LEN/NLEN 必须按字节对齐存放，因此需先跳到字节边界。
     * 注意必须依据「已消费比特总数」而非缓冲区位数来计算——
     * 缓冲区可能跨越多个字节，位数并非字节对齐的。
     * @private
     */
    private alignToByte(): void {
        const remainder = this.consumed % 8;
        if (remainder !== 0) {
            this.bits(8 - remainder);
        }
    }

    /**
     * 按霍夫曼表解码一个符号
     *
     * 逐位读取并累积码值，与各码长的累计边界比较：
     * 若当前码值落在该码长的起始区间内，则对应符号即为解码结果。
     *
     * @param {Huffman} h - 霍夫曼表
     * @returns {number} 解码出的符号
     */
    public decode(h: Huffman): number {
        let code = 0;
        let first = 0;
        let index = 0;

        for (let len = 1; len <= 15; len++) {
            code |= this.bits(1);
            const count = h.count[len];
            // code - first 为当前码长内的序号，落在 [0, count) 则命中
            if (code - first < count) {
                return h.symbol[index + (code - first)];
            }
            // 未命中则跳过该码长的所有码字，进入下一码长
            index += count;
            first += count;
            first <<= 1;
            code <<= 1;
        }
        throw new Error('DEFLATE 霍夫曼解码失败：非法码字');
    }

    /**
     * 构造 DEFLATE 固定霍夫曼表（见 RFC 1951 §3.2.6）
     * @returns {{ literal: Huffman, distance: Huffman }} 固定字面量表与距离表
     * @private
     */
    private static buildFixedTables(): { literal: Huffman, distance: Huffman } {
        const litLengths = new Uint8Array(288);
        // 0~143: 8 位；144~255: 9 位；256~279: 7 位；280~287: 8 位
        for (let i = 0; i < 144; i++) litLengths[i] = 8;
        for (let i = 144; i < 256; i++) litLengths[i] = 9;
        for (let i = 256; i < 280; i++) litLengths[i] = 7;
        for (let i = 280; i < 288; i++) litLengths[i] = 8;

        // 距离码固定为 5 位，共 32 个（其中 30/31 保留未用）
        const distLengths = new Uint8Array(30).fill(5);

        return {
            literal: new Huffman(litLengths, 288),
            distance: new Huffman(distLengths, 30)
        };
    }

    /**
     * 读取动态霍夫曼表的头部描述（HLIT/HDIST/HCLEN 及码长序列）
     * @returns {{ literal: Huffman, distance: Huffman }} 动态字面量表与距离表
     * @private
     */
    private readDynamicTables(): { literal: Huffman, distance: Huffman } {
        // HLIT：字面量/长度码数量 - 257（5 位，最大 286）
        const hlit = this.bits(5) + 257;
        // HDIST：距离码数量 - 1（5 位，最大 30）
        const hdist = this.bits(5) + 1;
        // HCLEN：码长码数量 - 4（4 位，最大 19）
        const hclen = this.bits(4) + 4;

        // —— 读取码长的码长（共 HCLEN 个，按 CODE_LENGTH_ORDER 顺序）——
        const codeLengths = new Uint8Array(19).fill(0);
        for (let i = 0; i < hclen; i++) {
            codeLengths[CODE_LENGTH_ORDER[i]] = this.bits(3);
        }
        const codeLengthTable = new Huffman(codeLengths, 19);

        // —— 用码长表解码出 (HLIT + HDIST) 个码长 ——
        const lengths = new Uint8Array(hlit + hdist).fill(0);
        let i = 0;
        while (i < hlit + hdist) {
            const sym = this.decode(codeLengthTable);
            if (sym < 16) {
                // 0~15：直接是一个码长
                lengths[i++] = sym;
            } else if (sym === 16) {
                // 16：重复上一个码长，次数 3~6（2 位 + 3）
                if (i === 0) throw new Error('DEFLATE 动态块：无前导码长却出现重复码 16');
                const prev = lengths[i - 1];
                const repeat = 3 + this.bits(2);
                for (let j = 0; j < repeat && i < hlit + hdist; j++) lengths[i++] = prev;
            } else if (sym === 17) {
                // 17：重复 0，次数 3~10（3 位 + 3）
                const repeat = 3 + this.bits(3);
                for (let j = 0; j < repeat && i < hlit + hdist; j++) lengths[i++] = 0;
            } else {
                // 18：重复 0，次数 11~138（7 位 + 11）
                const repeat = 11 + this.bits(7);
                for (let j = 0; j < repeat && i < hlit + hdist; j++) lengths[i++] = 0;
            }
        }

        // 拆分为字面量/长度表与距离表
        return {
            literal: new Huffman(lengths.subarray(0, hlit), hlit),
            distance: new Huffman(lengths.subarray(hlit), hdist)
        };
    }

    /**
     * 纯 TypeScript inflate 主流程
     *
     * 依次处理各压缩块：stored（未压缩）、fixed（固定霍夫曼）、dynamic（动态霍夫曼），
     * 直到遇到 end-of-block 符号 256。
     *
     * @param data - 压缩数据
     * @param expectedSize - 已知的解压后大小（用于预分配）
     * @returns 解压后的数据
     * @private
     */
    private static runInflate(data: Uint8Array, expectedSize?: number): Uint8Array {
        const stream = new BitStream(data);
        // 输出缓冲：已知大小时预分配，避免多次扩容
        let out = new Uint8Array(expectedSize && expectedSize > 0 ? expectedSize : Math.max(1024, data.length * 4));
        let outLen = 0;

        /** 确保输出缓冲至少有 additional 字节余量 */
        const ensureCapacity = (additional: number): void => {
            const needed = outLen + additional;
            if (needed <= out.length) return;
            // 按 1.5 倍扩容，降低扩容次数
            let newSize = Math.max(out.length * 1.5 | 0, needed);
            const bigger = new Uint8Array(newSize);
            bigger.set(out.subarray(0, outLen));
            out = bigger;
        };

        for (; ;) {
            // —— 读取块头：1 位 final 标志 + 2 位块类型 ——
            const isFinal = stream.bits(1);
            const blockType = stream.bits(2);

            if (blockType === 0) {
                // ═══ stored（未压缩块）═══
                // 先对齐到字节边界，再读 LEN(2) 与 NLEN(2)，最后跟原始数据
                stream.alignToByte();
                const len = stream.bits(16);
                const nlen = stream.bits(16);
                if ((len ^ 0xFFFF) !== nlen) {
                    throw new Error('DEFLATE stored 块长度校验失败');
                }
                ensureCapacity(len);
                for (let i = 0; i < len; i++) {
                    out[outLen++] = stream.bits(8);
                }
            } else if (blockType === 1 || blockType === 2) {
                // ═══ fixed / dynamic（霍夫曼压缩块）═══
                const tables = blockType === 1
                    ? BitStream.buildFixedTables()
                    : stream.readDynamicTables();

                for (; ;) {
                    const sym = stream.decode(tables.literal);

                    if (sym < 256) {
                        // 字面量：直接输出
                        ensureCapacity(1);
                        out[outLen++] = sym;
                    } else if (sym === 256) {
                        // end-of-block：当前块结束
                        break;
                    } else {
                        // 长度/距离对：回溯拷贝
                        const lenIdx = sym - 257;
                        if (lenIdx >= LENGTH_BASE.length) {
                            throw new Error('DEFLATE 非法长度码：' + sym);
                        }
                        const length = LENGTH_BASE[lenIdx] + stream.bits(LENGTH_EXTRA[lenIdx]);

                        const distSym = stream.decode(tables.distance);
                        if (distSym >= DISTANCE_BASE.length) {
                            throw new Error('DEFLATE 非法距离码：' + distSym);
                        }
                        const distance = DISTANCE_BASE[distSym] + stream.bits(DISTANCE_EXTRA[distSym]);

                        if (distance > outLen) {
                            throw new Error('DEFLATE 非法距离：超出已解压数据范围');
                        }

                        ensureCapacity(length);
                        // 从 outLen - distance 处逐字节拷贝（可能自我重叠，需逐字节而非块拷贝）
                        let from = outLen - distance;
                        for (let i = 0; i < length; i++) {
                            out[outLen++] = out[from++];
                        }
                    }
                }
            } else {
                // ═══ blockType === 3：保留的无效类型 ═══
                throw new Error('DEFLATE 非法块类型：3');
            }

            if (isFinal) break;
        }

        return out.subarray(0, outLen);
    }
}

/**
 * 纯 TypeScript inflate 入口
 * @param data - 压缩数据
 * @param expectedSize - 已知的解压后大小
 * @returns 解压后的数据
 * @private
 */
function inflatePure(data: Uint8Array, expectedSize?: number): Uint8Array {
    // BitStream.runInflate 为 private static，此处通过同一模块内的包装调用
    return (BitStream as unknown as {
        runInflate(d: Uint8Array, s?: number): Uint8Array
    }).runInflate(data, expectedSize);
}
