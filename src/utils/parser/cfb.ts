/**
 * 原生 OLE2 / CFB（Compound File Binary）复合文档解析模块
 *
 * .xls 文件并非单一字节流，而是 OLE2 复合文档——内部包含多个「流」（Stream）
 * 与「存储」（Storage）的树形结构，Excel 的数据存放在名为 `Workbook`
 * （Excel 97+）或 `Book`（Excel 5/95）的流中。
 *
 * **CFB 文件结构：**
 *
 * ```
 * ┌─────────────────────────────────┬────────┐
 * │ 文件头 Header（512 字节）         │ 扇区 0 │
 * ├─────────────────────────────────┼────────┤
 * │ 扇区 1..N（每扇区 512 或 4096 字节）        │
 * │   - FAT 扇区：扇区分配表                    │
 * │   - MiniFAT 扇区：小流的分配表                │
 * │   - 目录扇区：条目树                    │
 * │   - 数据扇区：实际流内容                      │
 * ├─────────────────────────────────┼────────┤
 * │ EOCD + DIFAT                    │
 * └─────────────────────────────────┴────────┘
 * ```
 *
 * **关键概念：**
 *
 * - **扇区（Sector）**：文件被切成固定大小的块（512 或 4096 字节），
 *   扇区 N 的起始偏移为 `(N + 1) * sectorSize`（+1 是因为首扇区是文件头）
 * - **FAT（File Allocation Table）**：记录每个扇区的后继扇区号，
 *   构成从扇区到扇区的链表，用于把散布的扇区串成连续逻辑流；
 *   值为 `0xFFFFFFFE`（ENDOFCHAIN）表示链表结束
 * - **DIFAT（Double-Indirect FAT）**：当 FAT 本身超过 109 个扇区时，
 *   需要额外的扇区来存放 FAT 的扇区号，由 DIFAT 扇区链式串联
 * - **MiniFAT / mini 流**：小于 4096 字节的流不占用完整扇区，
 *   而是打包进「mini 扇区」（64 字节）存放于根目录的 mini 流中，
 *   其分配关系由 MiniFAT 记录
 *
 * @module cfb
 * @author 东方鹗
 */

/** CFB 文件头签名字节：D0 CF 11 E0 A1 B1 1A E1 */
const CFB_SIGNATURE_BYTES: number[] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

/** 目录条目类型：未分配 */
const DIR_TYPE_EMPTY = 0;
/** 目录条目类型：存储（目录） */
const DIR_TYPE_STORAGE = 1;
/** 目录条目类型：流 */
const DIR_TYPE_STREAM = 2;
/** 目录条目类型：根目录 */
const DIR_TYPE_ROOT = 5;

/** 扇区链结束标记 */
const ENDOFCHAIN = 0xFFFFFFFE;
/** FAT 中的空闲扇区标记 */
const FREESECT = 0xFFFFFFFF;
/** DIFAT 中的扇区号标记（不再扩展） */
const DIFSECT = 0xFFFFFFFC;
/** DIFAT 中的结束标记 */
const ENDOFDIFAT = 0xFFFFFFFE;

/** 文件头固定长度 */
const HEADER_SIZE = 512;
/** 目录条目固定长度 */
const DIR_ENTRY_SIZE = 128;
/** DIFAT 在文件头中的固定条目数（109 个） */
const HEADER_DIFAT_COUNT = 109;

/** 根目录条目在目录链中的索引（恒为 0） */
const ROOT_ENTRY_INDEX = 0;

/**
 * CFB 目录条目（128 字节）
 */
export interface CfbEntry {
    /** 条目名称（UTF-16LE 编码，去除结尾的 null 字符） */
    name: string;
    /** 条目类型 */
    type: number;
    /** 该条目数据的起始扇区号 */
    startSector: number;
    /** 该条目数据的大小（字节） */
    size: number;
}

/**
 * CFB 复合文档
 */
export class CompoundFile {

    /** 文件全部字节 */
    private data: Uint8Array;
    /** 小端 DataView */
    private view: DataView;
    /** 扇区大小（512 或 4096） */
    private sectorSize: number;
    /** mini 扇区大小（通常为 64） */
    private miniSectorSize: number;
    /** mini 流的截止大小（通常为 4096） */
    private miniStreamCutoff: number;
    /** 完整 FAT：sector → nextSector */
    private fat: number[] = [];
    /** MiniFAT：miniSector → nextMiniSector */
    private miniFat: number[] = [];
    /** 根目录条目（mini 流存放在其中） */
    private rootEntry: CfbEntry | null = null;
    /** mini 流的完整数据（懒加载） */
    private miniStream: Uint8Array | null = null;
    /** 目录条目缓存 */
    private entries: CfbEntry[] = [];

    /**
     * @param {Uint8Array} data - OLE2 复合文档的完整字节数据
     * @throws {Error} 文件签名不匹配（非合法 OLE2 文件）
     */
    constructor(data: Uint8Array) {
        this.data = data;
        this.view = new DataView(data.buffer, data.byteOffset, data.byteLength);
        this.readHeader();
        this.buildFat();
        this.readDirectory();
    }

    /**
     * 读取文件头中的基础参数
     *
     * 复合文档头共 512 字节，关键字段偏移（MS-CFB 规范）：
     * -  0：签名（8 字节：D0 CF 11 E0 A1 B1 1A E1）
     * -  8：CLSID（16 字节）
     * - 24：minorVersion(u16)
     * - 26：majorVersion(u16)，3 = 512 字节扇区，4 = 4096 字节扇区
     * - 28：字节序标记（0xFFFE 表示小端）
     * - 30：扇区大小幂数(u16)，实际大小 = 1 << 该值
     * - 32：mini 扇区大小幂数(u16)，通常为 6 → 64 字节
     * - 34..39：保留
     * - 40：目录扇区个数(u32，仅 v4 有效，v3 中该位为保留)
     * - 44：FAT 扇区个数(u32)
     * - 48：目录链起始扇区(u32)
     * - 52：事务签名(u32)
     * - 56：mini 流截止大小(u32)，通常 4096
     * - 60：MiniFAT 起始扇区(u32)
     * - 64：MiniFAT 扇区个数(u32)
     * - 68：DIFAT 起始扇区(u32)
     * - 72：DIFAT 扇区个数(u32)
     * - 76..511：DIFAT 数组，109 个 u32 条目
     *
     * @private
     */
    private readHeader(): void {
        if (this.data.length < HEADER_SIZE) {
            throw new Error('文件过小，不是有效的 OLE2 复合文档');
        }

        // 校验签名（逐字节比较，避免字节序混淆）
        if (this.data.length < 8) {
            throw new Error('文件过小，不是有效的 OLE2 复合文档');
        }
        for (let i = 0; i < CFB_SIGNATURE_BYTES.length; i++) {
            if (this.data[i] !== CFB_SIGNATURE_BYTES[i]) {
                throw new Error('不是有效的 OLE2 复合文档（签名不匹配），可能不是 .xls 文件');
            }
        }

        // 扇区大小为 2 的幂：v3 恒为 512（shift 9），v4 通常为 4096（shift 12）
        // 直接信任 header 中的 sectorShift，但对非法值做兜底
        const sectorShift = this.view.getUint16(30, true);
        this.sectorSize = (sectorShift === 9 || sectorShift === 12)
            ? (1 << sectorShift)
            : (this.view.getUint16(26, true) === 4 ? 4096 : 512);

        // mini 扇区大小（默认 64 字节）
        const miniSectorShift = this.view.getUint16(32, true);
        this.miniSectorSize = 1 << (miniSectorShift || 6);

        // mini 流截止大小（通常为 4096 字节；小于此值的流存放在 mini 流中）
        this.miniStreamCutoff = this.view.getUint32(56, true) || 4096;
    }

    /**
     * 构建完整 FAT 与 MiniFAT
     *
     * FAT 扇区号列表来自两处：
     * 1. 文件头中的 109 个 DIFAT 条目（覆盖绝大多数文件）
     * 2. DIFAT 扇区链（FAT 超过 109 个扇区时才会出现的大文件）
     *
     * @private
     */
    private buildFat(): void {
        // —— 收集所有 FAT 扇区号 ——
        const fatSectors: number[] = [];

        // 文件头内的 109 个 DIFAT 条目（偏移 76 起）
        for (let i = 0; i < HEADER_DIFAT_COUNT; i++) {
            const sector = this.view.getUint32(76 + i * 4, true);
            if (sector === FREESECT || sector === ENDOFCHAIN) break;
            fatSectors.push(sector);
        }

        // DIFAT 扇区链：每个 DIFAT 扇区含 (sectorSize/4 - 1) 个 FAT 扇区号
        //              + 末尾 4 字节指向下一个 DIFAT 扇区
        const entriesPerDifatSector = this.sectorSize / 4 - 1;
        let difatSector = this.view.getUint32(68, true);
        let guard = 0;

        while (difatSector !== ENDOFCHAIN && difatSector !== FREESECT &&
            difatSector < this.sectorCount() && guard < 100000) {
            const base = this.sectorOffset(difatSector);
            for (let i = 0; i < entriesPerDifatSector; i++) {
                const sector = this.view.getUint32(base + i * 4, true);
                if (sector === FREESECT || sector === ENDOFCHAIN) continue;
                fatSectors.push(sector);
            }
            // 末尾 4 字节是下一个 DIFAT 扇区号
            difatSector = this.view.getUint32(base + entriesPerDifatSector * 4, true);
            guard++;
        }

        // —— 读取 FAT 内容 ——
        for (const sector of fatSectors) {
            const base = this.sectorOffset(sector);
            for (let i = 0; i < this.sectorSize / 4; i++) {
                this.fat.push(this.view.getUint32(base + i * 4, true));
            }
        }

        // —— 读取 MiniFAT（偏移 60 为起始扇区）——
        const miniFatSector = this.view.getUint32(64, true);
        if (miniFatSector !== ENDOFCHAIN && miniFatSector !== FREESECT) {
            const chain = this.readSectorChain(miniFatSector);
            const chainView = new DataView(
                chain.buffer, chain.byteOffset, chain.byteLength
            );
            for (let i = 0; i + 4 <= chain.length; i += 4) {
                this.miniFat.push(chainView.getUint32(i, true));
            }
        }
    }

    /**
     * 文件中的扇区总数
     * @returns 扇区总数
     * @private
     */
    private sectorCount(): number {
        return Math.max(0, Math.floor((this.data.length - HEADER_SIZE) / this.sectorSize));
    }

    /**
     * 计算扇区在文件中的字节偏移
     *
     * 扇区 0 从文件头的下一个字节开始，即偏移 = (sector + 1) * sectorSize
     *
     * @param {number} sector - 扇区号
     * @returns 字节偏移
     * @private
     */
    private sectorOffset(sector: number): number {
        return (sector + 1) * this.sectorSize;
    }

    /**
     * 沿 FAT 链读取一个流的内容
     *
     * 从起始扇区出发，每次通过 FAT 找到后继扇区，直到遇到 ENDOFCHAIN。
     * 结果可能跨越多个不连续扇区，需拷贝拼接。
     *
     * @param {number} startSector - 起始扇区号
     * @param {number} size - 流的字节大小；不传则读取到链结束
     * @returns {Uint8Array} 流内容
     * @private
     */
    private readSectorChain(startSector: number, size?: number): Uint8Array {
        // 已知大小时预分配，未知时按扇区数估算
        let result = new Uint8Array(size !== undefined ? size : this.sectorSize);
        let resultLen = 0;

        let sector = startSector;
        let guard = 0;

        while (sector !== ENDOFCHAIN && sector !== FREESECT &&
            sector !== DIFSECT && sector < this.fat.length && guard < 1000000) {

            const base = this.sectorOffset(sector);
            const available = Math.min(this.sectorSize, this.data.length - base);
            if (available <= 0) break;

            // 按需扩容
            if (resultLen + available > result.length) {
                const bigger = new Uint8Array(Math.max(result.length * 2, resultLen + available));
                bigger.set(result.subarray(0, resultLen));
                result = bigger;
            }

            result.set(this.data.subarray(base, base + available), resultLen);
            resultLen += available;

            // 已知大小且已读满则提前结束
            if (size !== undefined && resultLen >= size) break;

            sector = this.fat[sector];
            guard++;
        }

        return result.subarray(0, size !== undefined ? Math.min(size, resultLen) : resultLen);
    }

    /**
     * 读取根目录的 mini 流（存放所有小流的实际数据）
     * @returns {Uint8Array} mini 流数据；不存在时返回空数组
     * @private
     */
    private getMiniStream(): Uint8Array {
        if (this.miniStream) return this.miniStream;
        if (!this.rootEntry || this.rootEntry.startSector >= this.fat.length) {
            this.miniStream = new Uint8Array(0);
            return this.miniStream;
        }
        this.miniStream = this.readSectorChain(this.rootEntry.startSector, this.rootEntry.size);
        return this.miniStream;
    }

    /**
     * 沿 MiniFAT 链读取一个小流
     * @param {number} startMiniSector - 起始 mini 扇区号
     * @param {number} size - 数据大小
     * @returns {Uint8Array} 流内容
     * @private
     */
    private readMiniChain(startMiniSector: number, size: number): Uint8Array {
        const mini = this.getMiniStream();
        const result = new Uint8Array(size);
        let written = 0;
        let sector = startMiniSector;
        let guard = 0;

        while (sector !== ENDOFCHAIN && sector !== FREESECT &&
            sector < this.miniFat.length && written < size && guard < 1000000) {

            const offset = sector * this.miniSectorSize;
            if (offset + this.miniSectorSize > mini.length) break;

            const chunk = Math.min(this.miniSectorSize, size - written);
            result.set(mini.subarray(offset, offset + chunk), written);
            written += chunk;

            sector = this.miniFat[sector];
            guard++;
        }

        return result.subarray(0, written);
    }

    /**
     * 读取目录链并解析全部条目
     *
     * 目录是一个「红黑树」结构（此实现只需顺序遍历，不依赖树的平衡性），
     * 所有条目连续存放在目录扇区链中，每个 128 字节。
     *
     * @private
     */
    private readDirectory(): void {
        // 目录链起始扇区（偏移 48）
        const dirSector = this.view.getUint32(48, true);
        if (dirSector === ENDOFCHAIN || dirSector === FREESECT) {
            throw new Error('OLE2 复合文档缺少目录链');
        }

        // 读取目录链（条目数未知时读取到链结束）
        const dirData = this.readSectorChain(dirSector);
        const dirView = new DataView(dirData.buffer, dirData.byteOffset, dirData.byteLength);
        const count = Math.floor(dirData.length / DIR_ENTRY_SIZE);

        for (let i = 0; i < count; i++) {
            const base = i * DIR_ENTRY_SIZE;

            // 名称：UTF-16LE，固定 64 字节，实际长度由偏移 64 处的 u16 给出
            const nameLength = dirView.getUint16(base + 64, true);
            let name = '';
            if (nameLength >= 2) {
                // 名称长度含结尾的 UTF-16 null，按字节数 / 2 转为字符数再减 1
                const chars = Math.min(nameLength / 2 - 1, 32);
                for (let c = 0; c < chars; c++) {
                    name += String.fromCharCode(dirView.getUint16(base + c * 2, true));
                }
            }

            const type = dirData[base + 66];
            // 起始扇区（偏移 116，u32）；大小（偏移 120，u64，取低 32 位）
            const startSector = dirView.getUint32(base + 116, true);
            const size = dirView.getUint32(base + 120, true);

            if (type === DIR_TYPE_EMPTY && name === '') continue;

            this.entries.push({ name, type, startSector, size });
        }

        // 根目录条目（index 0）
        if (this.entries.length > ROOT_ENTRY_INDEX) {
            this.rootEntry = this.entries[ROOT_ENTRY_INDEX];
        }
    }

    /**
     * 列出复合文档中的全部流
     *
     * @returns {string[]} 流名称列表
     * @example
     * console.log(cfb.listStreams()); // ['Workbook', 'SummaryInformation', ...]
     */
    public listStreams(): string[] {
        return this.entries
            .filter(e => e.type === DIR_TYPE_STREAM)
            .map(e => e.name);
    }

    /**
     * 判断指定名称的流是否存在
     *
     * 名称比较不区分大小写，因为 CFB 存储的名称可能被转为大写。
     *
     * @param {string} name - 流名称
     * @returns 是否存在
     */
    public hasStream(name: string): boolean {
        return this.entries.some(e => e.type === DIR_TYPE_STREAM &&
            e.name.toLowerCase() === name.toLowerCase());
    }

    /**
     * 读取指定流的内容
     *
     * 按流大小决定存储方式：
     * - 小于 mini 流截止大小（4096）→ 存放在根目录的 mini 流中，走 MiniFAT 链
     * - 大于等于截止大小 → 直接占用完整扇区，走 FAT 链
     *
     * @param {string} name - 流名称（如 `Workbook`）
     * @returns {Uint8Array | null} 流内容；不存在时返回 null
     * @throws {Error} 流索引损坏
     * @example
     * const bytes = cfb.readStream('Workbook');
     */
    public readStream(name: string): Uint8Array | null {
        const entry = this.entries.find(e => e.type === DIR_TYPE_STREAM &&
            e.name.toLowerCase() === name.toLowerCase());

        if (!entry) return null;
        if (entry.size === 0) return new Uint8Array(0);

        if (entry.size < this.miniStreamCutoff) {
            // 小流：走 MiniFAT
            return this.readMiniChain(entry.startSector, entry.size);
        }

        // 大流：走 FAT
        return this.readSectorChain(entry.startSector, entry.size);
    }
}

/** 目录条目类型常量（供外部判断使用） */
export const CfbEntryType = {
    EMPTY: DIR_TYPE_EMPTY,
    STORAGE: DIR_TYPE_STORAGE,
    STREAM: DIR_TYPE_STREAM,
    ROOT: DIR_TYPE_ROOT
};
