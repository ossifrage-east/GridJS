/* 
 * 检测系统字体的工具函数
 * author: 东方鹗
 * B站: https://space.bilibili.com/194359739
 * 知乎: https://www.zhihu.com/people/eastossifrage
 * CSDN: https://blog.csdn.net/os373
 */


// 常见字体列表
const commonFonts = [
    // 中文字体
    'SimSun', 'NSimSun', 'FangSong', 'KaiTi', 'SimHei', 'HarmonyOS Sans SC Regular',
    'Microsoft YaHei', 'Microsoft JhengHei', 'PingFang SC',
    'Hiragino Sans GB', 'WenQuanYi Micro Hei', 'Source Han Sans CN',
    'Noto Sans CJK SC', 'Droid Sans Fallback', 'STHeiti', 'STKaiti',
    'STSong', 'STFangsong', 'LiHei Pro', 'LiSong Pro',
    // 英文字体
    'Arial', 'Arial Black', 'Arial Narrow', 'Arial Rounded MT Bold',
    'Avant Garde', 'Baskerville', 'Bodoni MT', 'Book Antiqua',
    'Bookman Old Style', 'Calibri', 'Cambria', 'Candara',
    'Century Gothic', 'Comic Sans MS', 'Consolas', 'Constantia',
    'Corbel', 'Courier New', 'DejaVu Sans', 'DejaVu Serif',
    'Didot', 'Franklin Gothic Medium', 'Futura', 'Garamoñd',
    'Geneva', 'Georgia', 'Gill Sans', 'Goudy Old Style',
    'Helvetica', 'Helvetica Neue', 'Impact', 'Lucida Bright',
    'Lucida Console', 'Lucida Grande', 'Lucida Sans Unicode',
    'Microsoft Sans Serif', 'Monaco', 'Monospace', 'Myriad Pro',
    'Optima', 'Palatino', 'Papyrus', 'Perpetua', 'Rockwell',
    'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS',
    'Verdana', 'Webdings', 'Wingdings', 'Zapfino',
    // 更多字体
    'Bahnschrift', 'Caladea', 'Carlito', 'Cascadia Code', 'Cascadia Mono',
    'Comfortaa', 'Fira Code', 'Fira Mono', 'Fira Sans', 'Inconsolata',
    'Inter', 'JetBrains Mono', 'Lato', 'Liberation Mono', 'Liberation Sans',
    'Liberation Serif', 'Montserrat', 'Noto Sans', 'Noto Serif', 'Open Sans',
    'Oswald', 'Playfair Display', 'Poppins', 'PT Sans', 'PT Serif',
    'Quicksand', 'Roboto', 'Roboto Condensed', 'Roboto Mono', 'Roboto Slab',
    'Rubik', 'Source Code Pro', 'Source Sans Pro', 'Source Serif Pro',
    'Ubuntu', 'Ubuntu Condensed', 'Ubuntu Mono'
];

// 中文字体数据
export const fontsZhData = [
    { en: 'SimSun', cn: '宋体', category: 'windows', usage: '中文印刷体，最常用的中文字体之一' },
    { en: 'NSimSun', cn: '新宋体', category: 'windows', usage: '宋体的变体，笔画更清晰' },
    { en: 'FangSong', cn: '仿宋', category: 'windows', usage: '仿照宋版书写的字体，笔画纤细' },
    { en: 'KaiTi', cn: '楷体', category: 'windows', usage: '模仿楷书书写的字体，常用于正式文档' },
    { en: 'SimHei', cn: '黑体', category: 'windows', usage: '笔画均匀，无衬线，醒目易读' },
    { en: 'Microsoft YaHei', cn: '微软雅黑', category: 'windows', usage: '微软开发的清晰屏幕显示字体' },
    { en: 'Microsoft JhengHei', cn: '微软正黑体', category: 'windows', usage: '微软为繁体中文开发的字体' },
    { en: 'YouYuan', cn: '幼圆', category: 'windows', usage: '圆润的黑体变体，风格柔和' },
    { en: 'LiSu', cn: '隶书', category: 'windows', usage: '模仿隶书书写的字体，古朴典雅' },
    { en: 'STSong', cn: '华文宋体', category: 'mac', usage: 'Mac系统常用宋体' },
    { en: 'STKaiti', cn: '华文楷体', category: 'mac', usage: 'Mac系统常用楷体' },
    { en: 'STHeiti', cn: '华文黑体', category: 'mac', usage: 'Mac系统常用黑体' },
    { en: 'STFangsong', cn: '华文仿宋', category: 'mac', usage: 'Mac系统常用仿宋' },
    { en: 'PingFang SC', cn: '苹方', category: 'mac', usage: '苹果iOS/macOS系统默认中文字体' },
    { en: 'Hiragino Sans GB', cn: '冬青黑体', category: 'mac', usage: 'Mac OS X系统内置的中文黑体' },
    { en: 'WenQuanYi Micro Hei', cn: '文泉驿微米黑', category: 'other', usage: '开源中文字体，Linux系统常用' },
    { en: 'Source Han Sans CN', cn: '思源黑体', category: 'other', usage: 'Adobe与Google合作开发的开源字体' },
    { en: 'Noto Sans CJK SC', cn: '思源黑体(Noto)', category: 'other', usage: 'Google的Noto字体家族中的中文黑体' },
    { en: 'FZShuTi', cn: '方正舒体', category: 'other', usage: '方正字库开发的舒体风格字体' },
    { en: 'FZYaoti', cn: '方正姚体', category: 'other', usage: '方正字库开发的姚体风格字体' },
    { en: 'STXingkai', cn: '华文行楷', category: 'mac', usage: 'Mac系统行楷字体' },
    { en: 'STLiti', cn: '华文隶书', category: 'mac', usage: 'Mac系统隶书字体' },
    { en: 'STHupo', cn: '华文琥珀', category: 'mac', usage: 'Mac系统琥珀体，笔画饱满' },
    { en: 'STCaiyun', cn: '华文彩云', category: 'mac', usage: 'Mac系统彩云体，空心效果' },
    { en: 'STXinwei', cn: '华文新魏', category: 'mac', usage: 'Mac系统新魏体，书法风格' },
    { en: 'LiHei Pro', cn: '丽黑 Pro', category: 'mac', usage: 'Mac系统专业版黑体' },
    { en: 'LiSong Pro', cn: '丽宋 Pro', category: 'mac', usage: 'Mac系统专业版宋体' },
    { en: 'BiauKai', cn: '标楷体', category: 'other', usage: '繁体中文常用楷体' },
    { en: 'MingLiU', cn: '细明体', category: 'windows', usage: 'Windows繁体中文默认字体' },
    { en: 'PMingLiU', cn: '新细明体', category: 'windows', usage: '细明体的更新版本' },
    { en: 'DFKai-SB', cn: '标楷体', category: 'windows', usage: 'Windows系统标楷体' },
    { en: 'SimLi', cn: '仿隶书', category: 'windows', usage: 'Windows系统仿隶书字体' },
    { en: 'SimYou', cn: '幼线体', category: 'windows', usage: 'Windows系统幼线字体' },
    { en: 'HarmonyOS Sans SC Regular', cn: '鸿蒙黑体', category: 'other', usage: '鸿蒙系统默认中文字体' }
];

export async function detectSystemFonts(fontList: string[] | null = null, testString = "abcdefghijklmnopqrstuvwxyz0123456789", timeout = 3000) {  // 探测系统字体
    // 默认字体列表
    const defaultFonts = commonFonts;

    const fontsToCheck = fontList || defaultFonts;
    const availableFonts: string[] = [];

    return new Promise((resolve) => {
        // 创建隐藏的canvas用于字体检测
        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');
        canvas.width = 200;
        canvas.height = 50;
        canvas.style.display = 'none';
        document.body.appendChild(canvas);

        // 获取基础字体的参考尺寸
        const baseFonts = ['sans-serif', 'serif', 'monospace'];
        const baseMeasurements: { [key: string]: { width: number, actualBoundingBoxAscent: number } } = {};
        
        baseFonts.forEach(baseFont => {
            context.font = `72px ${baseFont}`;
            baseMeasurements[baseFont] = {
                width: context.measureText(testString).width,
                // 对于支持OT-Metrics的浏览器，还可以检查其他度量
                actualBoundingBoxAscent: context.measureText(testString).actualBoundingBoxAscent || 0
            };
        });

        let checkedCount = 0;
        const totalFonts = fontsToCheck.length;

        // 设置超时
        const timeoutId = setTimeout(() => {
            document.body.removeChild(canvas);
            resolve(availableFonts);
        }, timeout);

        // 检测每个字体
        fontsToCheck.forEach(font => {
            // 测试字体是否存在
            const isAvailable = testFontAvailability(context, font, testString, baseMeasurements);
            
            if (isAvailable) {
                availableFonts.push(font);
            }
            
            checkedCount++;
            
            // 所有字体检测完成
            if (checkedCount === totalFonts) {
                clearTimeout(timeoutId);
                document.body.removeChild(canvas);
                resolve(availableFonts);
            }
        });
    });
}

/**
 * 测试单个字体的可用性
 */
function testFontAvailability(context: CanvasRenderingContext2D, fontName: string, testString: string, baseMeasurements: { [key: string]: { width: number, actualBoundingBoxAscent: number } }) {
    // 方法1：比较文本宽度
    context.font = `72px "${fontName}", sans-serif`;
    const widthWithFont = context.measureText(testString).width;
    
    // 如果宽度与基础字体不同，说明字体可能可用
    if (Math.abs(widthWithFont - baseMeasurements['sans-serif'].width) > 1) {
        return true;
    }
    
    // 方法2：尝试其他基础字体回退
    context.font = `72px "${fontName}", serif`;
    const widthWithSerifFallback = context.measureText(testString).width;
    
    if (Math.abs(widthWithSerifFallback - baseMeasurements['serif'].width) > 1) {
        return true;
    }
    
    // 方法3：检查字体度量（如果浏览器支持）
    const metrics = context.measureText(testString);
    if (metrics.actualBoundingBoxAscent && 
        Math.abs(metrics.actualBoundingBoxAscent - baseMeasurements['sans-serif'].actualBoundingBoxAscent) > 1) {
        return true;
    }
    
    return false;
}

/**
 * 全局自定义字体 @font-face CSS 注册表（预留接口）。
 * 打印组件（printer.ts）生成预览/打印文档时会读取此处登记的 CSS 并注入文档，
 * 后期用户自定义加入的字体只需在加载时登记一次，即可自动应用到打印预览与打印。
 */
const userFontFaceCssRegistry: string[] = [];

/**
 * 登记自定义字体的 @font-face CSS 文本（追加式，可多次调用，重复文本自动去重）
 * @param {string} css - 完整的 @font-face 规则文本，可一次传入多条规则
 * @example
 * registerFontFaceCss(`@font-face { font-family: 'MyFont'; src: url('/fonts/myfont.woff2') format('woff2'); }`);
 */
export function registerFontFaceCss(css: string): void {
    const trimmed = css?.trim();
    if (trimmed && !userFontFaceCssRegistry.includes(trimmed)) {
        userFontFaceCssRegistry.push(trimmed);
    }
}

/**
 * 读取全部已登记的 @font-face CSS（打印组件注入预览/打印文档时调用）
 * @returns {string[]} 已登记的 CSS 文本数组（副本）
 */
export function getRegisteredFontFaceCss(): string[] {
    return [...userFontFaceCssRegistry];
}


class BatchFontLoader {
    options: { display: string; timeout: number; };
    fonts: Map<any, any>;
    constructor(options = {}) {
        this.options = {
            display: 'swap',
            timeout: 30000, // 30秒超时
            ...options
        };
        this.fonts = new Map();
    }

    /**
     * 注册多个字体
     * @param {Object} fontConfig - 字体配置对象
     * 示例: { 'FontName': { woff2: 'url', woff: 'url' } }
     */
    registerFonts(fontConfig: { [s: string]: unknown; } | ArrayLike<unknown>) {
        Object.entries(fontConfig).forEach(([name, files]) => {
            this.fonts.set(name, {
                files,
                status: 'pending',
                fontFace: null
            });
        });
    }

    /**
     * 批量加载所有已注册字体
     * @returns {Promise<Object>} - 返回加载结果
     */
    async loadAll() {
        const results = {};
        const promises = [];

        // 为每个字体创建加载Promise
        for (const [name, data] of this.fonts) {
        if (data.status === 'loaded') continue;

        const loadPromise = this._loadFont(name, data.files)
            .then(fontFace => {
                (results as { [key: string]: { success: boolean; fontFace: FontFace } })[name] = { success: true, fontFace };
                data.status = 'loaded';
                data.fontFace = fontFace;
            })
            .catch(error => {
                (results as { [key: string]: { success: boolean; error: string } })[name] = { success: false, error: error.message };
                data.status = 'error';
            });

        // 设置超时
        const timeoutPromise = new Promise((_, reject) => {
            setTimeout(() => reject(new Error(`字体 "${name}" 加载超时`)), 
            this.options.timeout);
        });

        promises.push(Promise.race([loadPromise, timeoutPromise]));
        }

        // 等待所有字体加载完成
        await Promise.allSettled(promises);
        return results;
    }

    /**
     * 加载单个字体（内部方法）
     */
    async _loadFont(name: string, files: { [s: string]: unknown; } | ArrayLike<unknown>) {
        const src = Object.entries(files)
        .map(([format, url]) => `url('${url}') format('${format}')`)
        .join(', ');

        const fontFace = new FontFace(name, src, {
            display: this.options.display as FontDisplay,
            style: 'normal',
            weight: '400'
        });

        const loadedFont = await fontFace.load();
        (document.fonts as any).add(loadedFont as FontFace);

        // 登记到全局字体注册表：打印预览/打印文档注入 @font-face 后即可使用本字体
        registerFontFaceCss(`@font-face { font-family: '${name}'; src: ${src}; font-display: ${this.options.display}; font-weight: 400; font-style: normal; }`);

        return loadedFont;
    }

    /**
     * 应用字体到元素
     * @param {string} fontName - 字体名称
     * @param {HTMLElement|string} element - 元素或选择器
     */
    applyFont(fontName: string, element: HTMLElement | string) {  
        const target = typeof element === 'string' 
        ? document.querySelector(element) 
        : element;
        
        if (!target) return false;

        const fontData = this.fonts.get(fontName);
        if (fontData && fontData.status === 'loaded') {
            (target as HTMLElement).style.fontFamily = `'${fontName}', sans-serif`;
        return true;
        }
        return false;
    }
}

/**
 * 等待所有字体加载完成
 * @param {string[]} [fontNames] - 可选的字体名称列表，指定要等待的字体
 * @param {number} [timeout] - 超时时间（毫秒），默认 5000ms
 * @returns {Promise<boolean>} 字体是否加载成功
 */
export async function waitForFontsLoaded(fontNames?: string[], timeout: number = 5000): Promise<boolean> {
    try {
        if (fontNames && fontNames.length > 0) {
            const promises = fontNames.map(fontName => {
                return document.fonts.load(`1em "${fontName}"`);
            });
            await Promise.race([
                Promise.all(promises),
                new Promise((_, reject) => setTimeout(() => reject(new Error('Font load timeout')), timeout))
            ]);
        } else {
            await Promise.race([
                document.fonts.ready,
                new Promise((_, reject) => setTimeout(() => reject(new Error('Font load timeout')), timeout))
            ]);
        }
        return true;
    } catch (error) {
        console.warn('Font loading timeout or error:', error);
        return false;
    }
}
