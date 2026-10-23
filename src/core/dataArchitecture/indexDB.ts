/**
 * IndexDB 数据库操作类
 * 使用浏览器原生 IndexDB API 实现数据持久化存储
 * 支持 3 种对象存储：settings, values, images
 * 
 * @class IndexDB
 * @example
 * // 初始化
 * const db = new IndexDB();
 * await db.init();
 * 
 * // settings: 键值对存储 (key, value, type)
 * await db.setSetting('theme', 'dark');
 * await db.setSetting('language', 'zh-CN');
 * const theme = await db.getSetting('theme');
 * 
 * // values: 工作表值存储 (sheetName, value)
 * await db.setValue('Sheet1', JSON.stringify(data));
 * const sheetData = await db.getValue('Sheet1');
 * 
 * // images: 工作表图片存储 (sheetName, images)，结构化克隆直存 Uint8Array
 * await db.setImages('Sheet1', images);
 * const images = await db.getImages('Sheet1');
 */

/** 对象存储名称常量 */
const STORE_NAMES = {
    SETTINGS: 'settings',
    VALUES: 'values',
    IMAGES: 'images',
} as const;

/** 对象存储名称类型 */
export type StoreName = typeof STORE_NAMES[keyof typeof STORE_NAMES];

/** 对象存储配置 */
const OBJECT_STORE_CONFIG: Record<StoreName, IDBObjectStoreParameters> = {
    [STORE_NAMES.SETTINGS]: { keyPath: 'key', autoIncrement: false },
    [STORE_NAMES.VALUES]: { keyPath: 'sheetName', autoIncrement: false },
    [STORE_NAMES.IMAGES]: { keyPath: 'sheetName', autoIncrement: false },
};

/** 对象存储索引配置 */
const OBJECT_STORE_INDEXES: Record<StoreName, Array<{ name: string; keyPath: string; unique: boolean }>> = {
    [STORE_NAMES.SETTINGS]: [
        { name: 'type', keyPath: 'type', unique: false },
        { name: 'key', keyPath: 'key', unique: true },
    ],
    [STORE_NAMES.VALUES]: [
        { name: 'sheetName', keyPath: 'sheetName', unique: true },
    ],
    [STORE_NAMES.IMAGES]: [
        { name: 'sheetName', keyPath: 'sheetName', unique: true },
    ],
};

/** Settings 存储记录接口 */
export interface SettingRecord {
    /** 配置键 */
    key: string;
    /** 配置值 */
    value: any;
    /** 配置类型 */
    type?: string;
}

/** Values 存储记录接口 */
export interface ValueRecord {
    /** 工作表名称 */
    sheetName: string;
    /**
     * 工作表存档
     *
     * v4 起是**结构化克隆对象**（`SheetSnapshotV4`，含分片 JSON 文本与稀疏行列头）；
     * v3 及更早是 JSON 字符串。读取侧必须两态兼容（见 `DataCollection.fromSnapshotChunked`）。
     */
    value: unknown;
}

/**
 * Images 存储记录接口
 *
 * 图片是二进制大块，塞进 `values` 的 JSON 字符串里要过一遍 base64（体积 +33% 且
 * 多一次全量字符串拼接），故单独一个对象存储按工作表名归档，走结构化克隆直存。
 */
export interface ImageRecord {
    /** 工作表名称（与 values 库同键） */
    sheetName: string;
    /** 图片数组（`ExtractedImage[]`；以结构化克隆落库，Uint8Array 原样保留） */
    images: unknown[];
}

export class IndexDB {
    /** 数据库实例 */
    private db: IDBDatabase | null = null;
    /** 数据库名称 */
    private readonly DB_NAME: string = 'silkspaces';
    /**
     * 数据库版本号
     *
     * v2 起新增 `images` 对象存储（图片与单元格值分开存，见 {@link ImageRecord}）。
     * `onupgradeneeded` 里的建表逻辑是「缺哪个建哪个」，因此升级不会被旧库结构卡住。
     */
    private readonly DB_VERSION: number = 2;
    /** 是否已初始化 */
    private initialized: boolean = false;

    /**
     * 构造函数
     */
    constructor() {
        this.init();
    }

    /**
     * 异步初始化数据库
     * @returns {Promise<IDBDatabase>} 数据库实例
     */
    public async init(): Promise<IDBDatabase> {
        if (this.initialized && this.db) {
            return this.db;
        }
        
        this.db = await this.createDatabase();
        this.initialized = true;
        return this.db;
    }

    /**
     * 创建 IndexedDB 数据库
     * @private
     * @returns {Promise<IDBDatabase>} 数据库实例
     */
    private createDatabase(): Promise<IDBDatabase> {
        return new Promise((resolve, reject) => {
            const request = indexedDB.open(this.DB_NAME, this.DB_VERSION);

            request.onerror = () => reject(new Error(`Failed to open database: ${request.error}`));
            request.onsuccess = () => resolve(request.result);

            request.onupgradeneeded = (event) => {
                const database = (event.target as IDBOpenDBRequest).result;
                this.createObjectStores(database);
            };
        });
    }

    /**
     * 创建对象存储
     * @private
     * @param {IDBDatabase} database - 数据库实例
     */
    private createObjectStores(database: IDBDatabase): void {
        const storeNames: StoreName[] = [
            STORE_NAMES.SETTINGS,
            STORE_NAMES.VALUES,
            STORE_NAMES.IMAGES,
        ];

        storeNames.forEach(storeName => {
            if (!database.objectStoreNames.contains(storeName)) {
                const config = OBJECT_STORE_CONFIG[storeName];
                const objectStore = database.createObjectStore(storeName, config);
                
                const indexes = OBJECT_STORE_INDEXES[storeName];
                indexes.forEach(index => {
                    objectStore.createIndex(index.name, index.keyPath, { unique: index.unique });
                });
            }
        });
    }

    /**
     * 获取事务对象
     * @private
     * @param {'readonly' | 'readwrite'} mode - 事务模式
     * @param {StoreName} [storeName] - 对象存储名称，不传则包含所有存储
     * @returns {IDBTransaction} 事务对象
     */
    private getTransaction(mode: 'readonly' | 'readwrite', storeName?: StoreName): IDBTransaction {
        if (!this.db) {
            throw new Error('Database not initialized. Call init() first.');
        }
        const stores = storeName ? [storeName] : Object.values(STORE_NAMES);
        return this.db.transaction(stores, mode);
    }

    /**
     * 获取对象存储
     * @private
     * @param {IDBTransaction} transaction - 事务对象
     * @param {StoreName} storeName - 对象存储名称
     * @returns {IDBObjectStore} 对象存储
     */
    private getObjectStore(transaction: IDBTransaction, storeName: StoreName): IDBObjectStore {
        return transaction.objectStore(storeName);
    }

    /**
     * 确保数据库已初始化
     * @private
     */
    private async ensureInitialized(): Promise<void> {
        if (!this.initialized || !this.db) {
            await this.init();
        }
    }

    /**
     * 封装 IDBRequest 为 Promise
     * @private
     * @template T - 返回值类型
     * @param {IDBRequest} request - IDBRequest 对象
     * @returns {Promise<T>} Promise 对象
     */
    private wrapRequest<T>(request: IDBRequest): Promise<T> {
        return new Promise((resolve, reject) => {
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(new Error(`Database operation failed: ${request.error}`));
        });
    }

    // ==================== Settings 相关方法 ====================
    
    /**
     * 设置配置项（如果已存在则更新）
     * @param {string} key - 配置键
     * @param {any} value - 配置值
     * @param {string} [type='default'] - 配置类型
     * @returns {Promise<void>} 操作成功的 Promise
     */
    public async setSetting(key: string, value: any, type: string = 'default'): Promise<void> {
        await this.ensureInitialized();
        const transaction = this.getTransaction('readwrite', STORE_NAMES.SETTINGS);
        const store = this.getObjectStore(transaction, STORE_NAMES.SETTINGS);
        
        return this.wrapRequest<void>(store.put({ key, value, type }));
    }

    /**
     * 获取配置项
     * @param {string} key - 配置键
     * @returns {Promise<any | null>} 配置值或 null
     */
    public async getSetting(key: string): Promise<any | null> {
        await this.ensureInitialized();
        const transaction = this.getTransaction('readonly', STORE_NAMES.SETTINGS);
        const store = this.getObjectStore(transaction, STORE_NAMES.SETTINGS);
        
        const result = await this.wrapRequest<SettingRecord | undefined>(store.get(key));
        return result ? result.value : null;
    }

    /**
     * 删除配置项
     * @param {string} key - 配置键
     * @returns {Promise<void>} 操作成功的 Promise
     */
    public async deleteSetting(key: string): Promise<void> {
        await this.ensureInitialized();
        const transaction = this.getTransaction('readwrite', STORE_NAMES.SETTINGS);
        const store = this.getObjectStore(transaction, STORE_NAMES.SETTINGS);
        
        return this.wrapRequest<void>(store.delete(key));
    }

    /**
     * 获取指定类型的所有配置
     * @param {string} type - 配置类型
     * @returns {Promise<any[]>} 配置值数组
     */
    public async getSettingsByType(type: string): Promise<any[]> {
        await this.ensureInitialized();
        const transaction = this.getTransaction('readonly', STORE_NAMES.SETTINGS);
        const store = this.getObjectStore(transaction, STORE_NAMES.SETTINGS);
        const index = store.index('type');
        
        const results = await this.wrapRequest<SettingRecord[]>(index.getAll(type));
        return results.map(item => item.value);
    }

    /**
     * 获取所有配置项
     * @returns {Promise<Record<string, any>>} 所有配置的键值对
     */
    public async getAllSettings(): Promise<Record<string, any>> {
        await this.ensureInitialized();
        const transaction = this.getTransaction('readonly', STORE_NAMES.SETTINGS);
        const store = this.getObjectStore(transaction, STORE_NAMES.SETTINGS);
        
        const results = await this.wrapRequest<SettingRecord[]>(store.getAll());
        const map: Record<string, any> = {};
        results.forEach(item => {
            map[item.key] = item.value;
        });
        return map;
    }

    // ==================== Values 相关方法 ====================

    /**
     * 设置工作表存档（如果已存在则更新）
     *
     * 走结构化克隆：v4 存档对象（含 `string[]` 分片与 typed array 行列头）直存，
     * **不要再 JSON.stringify**（那是 v3 的老形态，装载时必然要一次不可分片的
     * 整份 `JSON.parse`）。
     * @param {string} sheetName - 工作表名称
     * @param {unknown} value - 工作表存档（`SheetSnapshotV4` 对象，或历史 JSON 字符串）
     * @returns {Promise<void>} 操作成功的 Promise
     */
    public async setValue(sheetName: string, value: unknown): Promise<void> {
        await this.ensureInitialized();
        const transaction = this.getTransaction('readwrite', STORE_NAMES.VALUES);
        const store = this.getObjectStore(transaction, STORE_NAMES.VALUES);
        
        return this.wrapRequest<void>(store.put({ sheetName, value }));
    }

    /**
     * 获取工作表存档
     * @param {string} sheetName - 工作表名称
     * @returns {Promise<unknown>} 存档值（v4 对象 / 历史 JSON 字符串）；无记录时返回 `'{}'`
     */
    public async getValue(sheetName: string): Promise<unknown> {
        await this.ensureInitialized();
        const transaction = this.getTransaction('readonly', STORE_NAMES.VALUES);
        const store = this.getObjectStore(transaction, STORE_NAMES.VALUES);
        
        const result = await this.wrapRequest<ValueRecord | undefined>(store.get(sheetName));
        return result ? result.value : '{}';
    }

    /**
     * 删除工作表值
     * @param {string} sheetName - 工作表名称
     * @returns {Promise<void>} 操作成功的 Promise
     */
    public async deleteValue(sheetName: string): Promise<void> {
        await this.ensureInitialized();
        const transaction = this.getTransaction('readwrite', STORE_NAMES.VALUES);
        const store = this.getObjectStore(transaction, STORE_NAMES.VALUES);
        
        return this.wrapRequest<void>(store.delete(sheetName));
    }

    /**
     * 获取所有工作表值
     * @returns {Promise<ValueRecord[]>} 所有工作表值记录数组
     */
    public async getAllValues(): Promise<ValueRecord[]> {
        await this.ensureInitialized();
        const transaction = this.getTransaction('readonly', STORE_NAMES.VALUES);
        const store = this.getObjectStore(transaction, STORE_NAMES.VALUES);
        
        return this.wrapRequest<ValueRecord[]>(store.getAll());
    }

    /**
     * 只取所有工作表**名称**（键），不反序列化存档内容
     *
     * 【勿用 `getAllValues().map(r => r.sheetName)` 代替】`getAll()` 会把每条存档的
     * `value` 一起读出来 —— 大表的工作区里那是几十万格的快照，纯粹为了几个名字
     * 付一次全量反序列化。`getAllKeys()` 走的是键空间游标，代价与记录大小无关。
     * @returns {Promise<string[]>} 工作表名称（即存档键）列表
     */
    public async getAllSheetNames(): Promise<string[]> {
        await this.ensureInitialized();
        const transaction = this.getTransaction('readonly', STORE_NAMES.VALUES);
        const store = this.getObjectStore(transaction, STORE_NAMES.VALUES);

        return this.wrapRequest<IDBValidKey[]>(store.getAllKeys()).then(keys =>
            keys.map(key => String(key)));
    }

    // ==================== Images 相关方法 ====================

    /**
     * 写入某工作表的图片（已存在则整条替换）
     *
     * 走结构化克隆：`Uint8Array` 原样落库，无需 base64，也不产生中间字符串。
     * @param {string} sheetName - 工作表名称（与 values 库同键）
     * @param {unknown[]} images - 图片数组（`ExtractedImage[]`）
     * @returns {Promise<void>} 操作成功的 Promise
     */
    public async setImages(sheetName: string, images: unknown[]): Promise<void> {
        await this.ensureInitialized();
        const transaction = this.getTransaction('readwrite', STORE_NAMES.IMAGES);
        const store = this.getObjectStore(transaction, STORE_NAMES.IMAGES);

        return this.wrapRequest<void>(store.put({ sheetName, images }));
    }

    /**
     * 读取某工作表的图片
     * @param {string} sheetName - 工作表名称
     * @returns {Promise<unknown[]>} 图片数组；无记录时为**空数组**（调用方无需再判 undefined）
     */
    public async getImages(sheetName: string): Promise<unknown[]> {
        await this.ensureInitialized();
        const transaction = this.getTransaction('readonly', STORE_NAMES.IMAGES);
        const store = this.getObjectStore(transaction, STORE_NAMES.IMAGES);

        const result = await this.wrapRequest<ImageRecord | undefined>(store.get(sheetName));
        return result && Array.isArray(result.images) ? result.images : [];
    }

    /**
     * 删除某工作表的图片
     * @param {string} sheetName - 工作表名称
     * @returns {Promise<void>} 操作成功的 Promise
     */
    public async deleteImages(sheetName: string): Promise<void> {
        await this.ensureInitialized();
        const transaction = this.getTransaction('readwrite', STORE_NAMES.IMAGES);
        const store = this.getObjectStore(transaction, STORE_NAMES.IMAGES);

        return this.wrapRequest<void>(store.delete(sheetName));
    }

    /**
     * 获取所有图片记录
     * @returns {Promise<ImageRecord[]>} 全部图片记录（含二进制，慎用于启动路径）
     */
    public async getAllImages(): Promise<ImageRecord[]> {
        await this.ensureInitialized();
        const transaction = this.getTransaction('readonly', STORE_NAMES.IMAGES);
        const store = this.getObjectStore(transaction, STORE_NAMES.IMAGES);

        return this.wrapRequest<ImageRecord[]>(store.getAll());
    }

    /**
     * 获取所有图片记录的键（工作表名）
     *
     * 只取键不取值：启动对账只需要知道「哪些名字有图片记录」，
     * 用 `getAllImages()` 会把全部图片二进制读进内存（可能上百 MB）。
     * @returns {Promise<string[]>} 工作表名数组
     */
    public async getAllImageKeys(): Promise<string[]> {
        await this.ensureInitialized();
        const transaction = this.getTransaction('readonly', STORE_NAMES.IMAGES);
        const store = this.getObjectStore(transaction, STORE_NAMES.IMAGES);

        const keys = await this.wrapRequest<IDBValidKey[]>(store.getAllKeys());
        return keys.map((key) => String(key));
    }

    // ==================== 通用方法 ====================

    /**
     * 清空指定对象存储的所有数据
     * @param {StoreName} storeName - 对象存储名称
     * @returns {Promise<void>} 操作成功的 Promise
     */
    public async clearStore(storeName: StoreName): Promise<void> {
        await this.ensureInitialized();
        const transaction = this.getTransaction('readwrite', storeName);
        const store = this.getObjectStore(transaction, storeName);
        
        return this.wrapRequest<void>(store.clear());
    }

    /**
     * 清空所有对象存储的数据
     * @returns {Promise<void>} 操作成功的 Promise
     */
    public async clearAll(): Promise<void> {
        const storeNames: StoreName[] = [
            STORE_NAMES.SETTINGS,
            STORE_NAMES.VALUES,
            STORE_NAMES.IMAGES,
        ];

        for (const storeName of storeNames) {
            await this.clearStore(storeName);
        }
    }

    /**
     * 关闭数据库连接
     */
    public close(): void {
        if (this.db) {
            this.db.close();
            this.db = null;
            this.initialized = false;
        }
    }

    /**
     * 删除数据库
     * @returns {Promise<void>} 操作成功的 Promise
     */
    public static async deleteDatabase(): Promise<void> {
        const DB_NAME = 'silkspaces';
        return new Promise((resolve, reject) => {
            const request = indexedDB.deleteDatabase(DB_NAME);
            request.onsuccess = () => resolve();
            request.onerror = () => reject(new Error(`Failed to delete database: ${request.error}`));
        });
    }

    /**
     * 检查浏览器是否支持 IndexedDB
     * @returns {boolean} 是否支持
     */
    public static isSupported(): boolean {
        return typeof indexedDB !== 'undefined';
    }
}
