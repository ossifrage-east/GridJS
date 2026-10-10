/**
 * 原生 ZIP 解析模块（无第三方依赖）
 *
 * ZIP 是 xlsx 文件的容器格式。解析分两步：
 *
 * 1. **定位中央目录（EOCD）**：从文件尾部反向扫描 `0x06054b50` 签名。
 *    中央目录集中记录了每个条目的文件名、压缩方法、CRC32 与压缩后大小。
 * 2. **读取每个条目**：从中央目录拿到本地文件头偏移，读取本地头得到
 *    文件名与额外字段长度，从而算出数据起始位置，再按压缩方法解压。
 *
 * 兼容处理：
 * - 压缩方法 0（stored，未压缩）与 8（deflate）
 * - Zip64：中央目录/本地头中 32 位字段溢出时的扩展字段（0x0001 头标识）
 * - 数据描述符（bit 3 置位时大小在数据后，本实现以中央目录为准，无需回读）
 *
 * @module zip
 * @author 东方鹗
 */

import { inflateRaw } from './inflate';

/** EOCD（End of Central Directory）签名 */
const EOCD_SIGNATURE = 0x06054b50;
/** Zip64 EOCD 定位器签名 */
const EOCD64_LOCATOR_SIGNATURE = 0x07064b50;
/** Zip64 EOCD 记录签名 */
const EOCD64_SIGNATURE = 0x06064b50;
/** 中央目录文件头签名 */
const CENTRAL_HEADER_SIGNATURE = 0x02014b50;
/** 本地文件头签名 */
const LOCAL_HEADER_SIGNATURE = 0x04034b50;

/** EOCD 固定长度（不含变长注释） */
const EOCD_MIN_SIZE = 22;

/** 压缩方法：未压缩 */
const METHOD_STORED = 0;
/** 压缩方法：DEFLATE */
const METHOD_DEFLATE = 8;

/** 单条目解压后的最大字节数（防止 ZIP 炸弹：约 512MB） */
const MAX_ENTRY_SIZE = 512 * 1024 * 1024;

/**
 * ZIP 条目信息
 */
export interface ZipEntry {
    /** 条目内的文件路径（如 `xl/worksheets/sheet1.xml`） */
    path: string;
    /** 压缩方法（0 = stored，8 = deflate） */
    method: number;
    /** 压缩后的字节大小 */
    compressedSize: number;
    /** 解压后的字节大小 */
    uncompressedSize: number;
    /** 本地文件头在文件中的偏移 */
    localHeaderOffset: number;
    /** CRC32 校验值 */
    crc32: number;
}

/**
 * ZIP 解析结果
 */
export interface ZipArchive {
    /** 条目总数 */
    count: number;
    /** 路径 → 条目信息 */
    entries: Map<string, ZipEntry>;
}

/**
 * ZIP 解析器
 *
 * 以「先读中央目录拿到全部条目，再按需读取条目内容」的方式工作，
 * 避免一次性把整个压缩包解压到内存。
 *
 * @example
 * const zip = new ZipParser(buffer);
 * const archive = zip.parseCentralDirectory();
 * const xmlBytes = await zip.readEntry(archive, 'xl/worksheets/sheet1.xml');
 */
export class ZipParser {

    /** 整个文件的字节数据 */
    private data: Uint8Array;
    /** DataView 视图（小端读取） */
    private view: DataView;

    /**
     * @param {Uint8Array} data - ZIP 文件的完整字节数据
     */
    constructor(data: Uint8Array) {
        this.data = data;
        this.view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    }

    /**
     * 定位 EOCD 记录位置
     *
     * EOCD 位于文件尾部，其后紧跟「文件注释」，长度可达 65535 字节。
     * 因此需从尾部向前在最多 65557 字节内反向扫描签名。
     * @returns EOCD 在文件中的偏移；未找到时返回 -1
     * @private
     */
    private findEOCD(): number {
        // 注释最大 65535 + EOCD 固定 22
        const maxSearch = Math.min(this.data.length, EOCD_MIN_SIZE + 0xFFFF);
        const start = this.data.length - maxSearch;

        for (let i = this.data.length - EOCD_MIN_SIZE; i >= start; i--) {
            if (this.view.getUint32(i, true) === EOCD_SIGNATURE) {
                return i;
            }
        }
        return -1;
    }

    /**
     * 定位中央目录的偏移与条目数
     *
     * 常规 EOCD 的 32 位字段在文件超过 4GB 或条目超过 65535 时会溢出，
     * 此时需通过 Zip64 EOCD 定位器读取真正的 64 位值。
     *
     * @returns 中央目录偏移与条目总数
     * @private
     */
    private locateCentralDirectory(): { offset: number; count: number } {
        const eocd = this.findEOCD();
        if (eocd === -1) {
            throw new Error('不是有效的 ZIP 文件：未找到中央目录结束记录(EOCD)');
        }

        // EOCD 结构：偏移 10 为条目总数(u16)，偏移 12 为中央目录大小(u32)，偏移 16 为中央目录偏移(u32)
        let count = this.view.getUint16(eocd + 10, true);
        let offset = this.view.getUint32(eocd + 16, true);

        // —— Zip64 分支 ——
        // 条目数为 0xFFFF 或中央目录偏移为 0xFFFFFFFF 时，表示真实值存放在 Zip64 记录中
        if (count === 0xFFFF || offset === 0xFFFFFFFF) {
            const zip64 = this.locateZip64EOCD(eocd);
            if (zip64) {
                count = zip64.count;
                offset = zip64.offset;
            }
        }

        return { offset, count };
    }

    /**
     * 读取 Zip64 EOCD 记录中的条目数与中央目录偏移
     *
     * Zip64 EOCD 定位器固定为 20 字节，位于常规 EOCD 之前，
     * 其中含 Zip64 EOCD 记录的绝对偏移。
     *
     * @param eocdOffset - 常规 EOCD 的偏移
     * @returns 条目数与中央目录偏移；未找到返回 null
     * @private
     */
    private locateZip64EOCD(eocdOffset: number): { count: number, offset: number } | null {
        const locatorOffset = eocdOffset - 20;
        if (locatorOffset < 0) return null;
        if (this.view.getUint32(locatorOffset, true) !== EOCD64_LOCATOR_SIGNATURE) {
            return null;
        }

        // 定位器偏移 8 为 Zip64 EOCD 记录的绝对偏移（8 字节）
        const zip64Offset = Number(this.view.getBigUint64(locatorOffset + 8, true));
        if (zip64Offset < 0 || zip64Offset + 56 > this.data.length) return null;
        if (this.view.getUint32(zip64Offset, true) !== EOCD64_SIGNATURE) return null;

        // Zip64 EOCD：偏移 32 为条目总数(u64)，偏移 48 为中央目录偏移(u64)
        return {
            count: Number(this.view.getBigUint64(zip64Offset + 32, true)),
            offset: Number(this.view.getBigUint64(zip64Offset + 48, true))
        };
    }

    /**
     * 解析中央目录，得到全部条目信息
     *
     * 中央目录中每个条目是一段「中央目录文件头 + 文件名 + 额外字段 + 文件注释」，
     * 逐条顺序扫描即可建立完整的条目索引。
     *
     * @returns {ZipArchive} 包含全部条目信息的压缩包
     * @throws {Error} 中央目录格式非法
     * @example
     * const archive = zip.parseCentralDirectory();
     * console.log(archive.count, [...archive.entries.keys()]);
     */
    public parseCentralDirectory(): ZipArchive {
        const { offset, count } = this.locateCentralDirectory();

        if (offset < 0 || offset > this.data.length) {
            throw new Error('中央目录偏移越界，文件可能已损坏');
        }

        const entries = new Map<string, ZipEntry>();
        let pos = offset;

        for (let i = 0; i < count; i++) {
            if (pos + 46 > this.data.length) {
                throw new Error('中央目录条目越界');
            }
            if (this.view.getUint32(pos, true) !== CENTRAL_HEADER_SIGNATURE) {
                throw new Error('中央目录文件头签名不匹配于第 ' + i + ' 项');
            }

            // 中央目录文件头结构（偏移相对 pos）：
            //  4  压缩方法(u16)  6  最后修改时间  8  最后修改日期
            // 10  CRC32(u32)    14  压缩大小(u32) 18  解压大小(u32)
            // 22  文件名长度(u16) 24 额外字段长度(u16) 26 文件注释长度(u16)
            // 28  磁盘起始号(u16) 30 内部属性(u16) 32 外部属性(u32)
            // 36  本地头偏移(u32)
            const method = this.view.getUint16(pos + 10, true);
            const crc32 = this.view.getUint32(pos + 16, true);
            let compressedSize = this.view.getUint32(pos + 20, true);
            let uncompressedSize = this.view.getUint32(pos + 24, true);
            const nameLength = this.view.getUint16(pos + 28, true);
            const extraLength = this.view.getUint16(pos + 30, true);
            const commentLength = this.view.getUint16(pos + 32, true);
            let localHeaderOffset = this.view.getUint32(pos + 42, true);

            // 文件名（UTF-8）
            const nameBytes = this.data.subarray(pos + 46, pos + 46 + nameLength);
            const path = utf8Decode(nameBytes);

            // 额外字段可能包含 Zip64 扩展（头标识 0x0001），内含真实的大小与偏移
            const extraStart = pos + 46 + nameLength;
            const zip64 = this.readZip64Extra(
                extraStart, extraLength, uncompressedSize, compressedSize, localHeaderOffset
            );
            if (zip64) {
                uncompressedSize = zip64.uncompressedSize;
                compressedSize = zip64.compressedSize;
                localHeaderOffset = zip64.localHeaderOffset;
            }

            entries.set(path, {
                path,
                method,
                compressedSize,
                uncompressedSize,
                localHeaderOffset,
                crc32
            });

            // 前进到下一个条目
            pos = extraStart + extraLength + commentLength;
        }

        return { count, entries };
    }

    /**
     * 解析 Zip64 扩展字段（额外字段中头标识为 0x0001 的那一段）
     *
     * 该扩展按固定顺序存放，且**仅包含那些在常规字段中溢出（值为 0xFFFFFFFF）** 的值，
     * 因此需逐个判断是否需要读取。
     *
     * @param offset - 额外字段起始偏移
     * @param length - 额外字段总长度
     * @param uncompressedSize - 常规解压大小（溢出时为 0xFFFFFFFF）
     * @param compressedSize - 常规压缩大小
     * @param localHeaderOffset - 常规本地头偏移
     * @returns 修正后的值；无 Zip64 扩展时返回 null
     * @private
     */
    private readZip64Extra(
        offset: number,
        length: number,
        uncompressedSize: number,
        compressedSize: number,
        localHeaderOffset: number
    ): { uncompressedSize: number, compressedSize: number, localHeaderOffset: number } | null {
        let pos = offset;
        const end = offset + length;

        while (pos + 4 <= end) {
            const headerId = this.view.getUint16(pos, true);
            const dataSize = this.view.getUint16(pos + 2, true);
            const dataStart = pos + 4;

            if (headerId === 0x0001) {
                // Zip64 扩展字段：顺序为 解压大小、压缩大小、本地头偏移、磁盘起始号
                let p = dataStart;
                if (uncompressedSize === 0xFFFFFFFF && p + 8 <= dataStart + dataSize) {
                    uncompressedSize = Number(this.view.getBigUint64(p, true));
                    p += 8;
                }
                if (compressedSize === 0xFFFFFFFF && p + 8 <= dataStart + dataSize) {
                    compressedSize = Number(this.view.getBigUint64(p, true));
                    p += 8;
                }
                if (localHeaderOffset === 0xFFFFFFFF && p + 8 <= dataStart + dataSize) {
                    localHeaderOffset = Number(this.view.getBigUint64(p, true));
                }
                return { uncompressedSize, compressedSize, localHeaderOffset };
            }

            pos = dataStart + dataSize;
        }

        return null;
    }

    /**
     * 读取并解压指定条目
     *
     * 步骤：定位本地文件头 → 读出文件名长度与额外字段长度以确定数据起点 →
     * 取出压缩数据 → 按压缩方法处理。
     *
     * 注意本地文件头中的文件名可能与中央目录不同（同一文件被重命名的情况），
     * 因此数据起点必须以本地文件头为准。
     *
     * @param {ZipArchive} archive - 中央目录解析结果
     * @param {string} path - 条目路径
     * @returns {Promise<Uint8Array>} 条目内容；路径不存在时返回 null
     * @throws {Error} 压缩方法不支持或数据损坏
     * @example
     * const bytes = await zip.readEntry(archive, 'xl/styles.xml');
     */
    public async readEntry(archive: ZipArchive, path: string): Promise<Uint8Array | null> {
        const entry = archive.entries.get(path);
        if (!entry) return null;

        const localPos = entry.localHeaderOffset;
        if (localPos + 30 > this.data.length) {
            throw new Error('本地文件头越界：' + path);
        }
        if (this.view.getUint32(localPos, true) !== LOCAL_HEADER_SIGNATURE) {
            throw new Error('本地文件头签名不匹配：' + path);
        }

        // 本地文件头：偏移 26 为文件名长度(u16)，偏移 28 为额外字段长度(u16)
        const nameLength = this.view.getUint16(localPos + 26, true);
        const extraLength = this.view.getUint16(localPos + 28, true);
        const dataStart = localPos + 30 + nameLength + extraLength;
        const dataEnd = dataStart + entry.compressedSize;

        if (dataEnd > this.data.length) {
            throw new Error('条目数据越界：' + path);
        }

        // 防止 ZIP 炸弹：解压后大小超出上限时直接拒绝
        if (entry.uncompressedSize > MAX_ENTRY_SIZE) {
            throw new Error('条目解压后过大（超过 512MB），疑似 ZIP 炸弹：' + path);
        }

        const compressed = this.data.subarray(dataStart, dataEnd);

        if (entry.method === METHOD_STORED) {
            // 未压缩：直接返回原数据副本
            return new Uint8Array(compressed);
        }
        if (entry.method === METHOD_DEFLATE) {
            return inflateRaw(compressed, entry.uncompressedSize);
        }

        throw new Error('不支持的压缩方法 ' + entry.method + '：' + path);
    }
}

/**
 * UTF-8 字节解码为字符串
 *
 * ZIP 的文件名统一使用 UTF-8（general purpose bit 11 置位）或 CP437。
 * xlsx 文件名均为 ASCII 路径，故按 UTF-8 解码即可覆盖全部实际场景。
 *
 * @param {Uint8Array} bytes - UTF-8 字节
 * @returns {string} 解码后的字符串
 * @private
 */
function utf8Decode(bytes: Uint8Array): string {
    if (typeof TextDecoder !== 'undefined') {
        return new TextDecoder('utf-8').decode(bytes);
    }
    // 老旧环境回退：手动解码
    let result = '';
    for (let i = 0; i < bytes.length;) {
        const b = bytes[i];
        if (b < 0x80) {
            result += String.fromCharCode(b);
            i += 1;
        } else if (b >= 0xC0 && b < 0xE0) {
            result += String.fromCharCode(((b & 0x1F) << 6) | (bytes[i + 1] & 0x3F));
            i += 2;
        } else if (b >= 0xE0 && b < 0xF0) {
            result += String.fromCharCode(
                ((b & 0x0F) << 12) | ((bytes[i + 1] & 0x3F) << 6) | (bytes[i + 2] & 0x3F)
            );
            i += 3;
        } else {
            result += String.fromCharCode(b);
            i += 1;
        }
    }
    return result;
}
