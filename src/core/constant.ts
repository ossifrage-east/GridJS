/* 
 * 常量模块
 * author: 东方鹗
 * B站: https://space.bilibili.com/194359739
 * 知乎: https://www.zhihu.com/people/eastossifrage
 * CSDN: https://blog.csdn.net/os373
 */

// 常量定义
export const EDIT_CANVAS = 'editCanvas';
export const SHEETS_NAME = {
    SHEETS: 'sheet-workspace',
    ALL_SELECT: 'allSelect',
    COL_HEADER: 'colHeaderCanvas',
    ROW_HEADER: 'rowHeaderCanvas',
    SHEET: 'sheetCanvas'
}
export const SCROLLER_CLASS_NAMES = {   // 定义 scrollbar 类名
    CONTAINER: 'scroll-',
    TRACK: 'scroll-track',
    BACK: 'scroll-back',
    THUMB: 'scroll-thumb',
    FORWARD: 'scroll-forward'
}
export const SCROLLER_THUMB_MIN_SIZE = 20;  // 滚动条滑块的最小尺寸(像素)

export const TOOLBAR_CLASS_NAME = {
    CONTAINER: 'toolbar-container',
    BRAND: 'brand',
    QUICK_ACCESS: 'toolbar-quick-access',
    CATEGORY: 'toolbar-category',
    TABS: 'toolbar-tabs',
    TAB: 'toolbar-tab',
    CONTENT: 'toolbar-content',
    GROUP: 'toolbar-group',
    SECTION: 'toolbar-section',
    SECTION_LINE: 'toolbar-section-line',
    /** 紧凑模式下的「一级菜单」条（各分区的名字做成一排小标签） */
    SECTION_NAV: 'toolbar-section-nav',
    /** 一级菜单里的单个分区标签 */
    SECTION_CHIP: 'toolbar-section-chip',
    /** 紧凑模式下当前展开的那个分区 */
    SECTION_ACTIVE: 'toolbar-section-active',
    /** 容器上的紧凑模式开关（宽度不够时加） */
    COMPACT: 'toolbar-compact',
    /** 量「内容自然宽度」时临时加在分组上（absolute + max-content + 隐藏） */
    MEASURING: 'toolbar-measuring',
    /** 根分类折叠模式（一页签那行放不下时加在容器上） */
    ROOT_COMPACT: 'toolbar-root-compact',
    /** 根分类清单（浮层）展开中 */
    ROOT_OPEN: 'toolbar-root-open',
    /** 根折叠模式下**露出来的那一个**分类（活动分类；没有活动分类时是第一个） */
    TAB_ROOT: 'toolbar-tab-root',
    /** 根折叠后的「分类清单」浮层 */
    TABS_NAV: 'toolbar-tabs-nav',
    /** 分类清单里的单个分类标签 */
    TABS_CHIP: 'toolbar-tab-chip',
    /** 快速访问区的按钮组（窄屏时整组收进浮层） */
    QUICK_PANEL: 'toolbar-quick-access-panel',
    /** 快速访问区的浮层展开中 */
    QUICK_OPEN: 'toolbar-quick-access-open',
}
export const NAV_PANEL: string[] = ['文件', '工具', '脚本']
/**
 * 各导航项「内容区」的初始化宽度（像素）
 *
 * 未在此声明的导航项，初始化宽度为「容器宽度平分」。
 * 用户人工拖拽后的宽度会以 `nav-content-width:<导航项文本>` 为键存入
 * IndexedDB，刷新或点击切换到该导航项时优先读取已保存值。
 */
export const NAV_CONTENT_WIDTH: Record<string, number> = {
    '文件': 360,
    '工具': 360,
    '脚本': 480,
};
export const TOOLBAR_TABS = {
    START:'开始',
    INSERT:'插入',
    PAGE:'页面',
    DATA:'数据',
    VIEW:'视图',
    TOOLS:'工具'
};

export const ASSISTANT = {
    FILTER_CLASS_NAME: 'filter',
    FILTER_CONTAINER_CLASS_NAME: 'filter-container',
    FILTER_COLOR_LABEL_CLASS_NAME: 'filter-color-label',
    FILTER_INPUT_CLASS_NAME: 'filter-input',
    FILTER_TIPS_CLASS_NAME: 'filter-tips',
    FILTER_TIPS_HEADER_CLASS_NAME: 'filter-tips-header',
    FILTER_TIPS_INVERT_CLASS_NAME: 'filter-tips-invert',
    FILTER_TIPS_ITEM_CLASS_NAME: 'filter-tips-item',
    FILTER_TIPS_TEXT_CLASS_NAME: 'filter-tips-text',
    FILTER_TIPS_COUNT_CLASS_NAME: 'filter-tips-count',
    FILTER_TIPS_ONLY_CLASS_NAME: 'filter-tips-only',
    FILTER_BUTTONS_CLASS_NAME: 'filter-buttons',
    FILTER_CONFIRM_CLASS_NAME: 'filter-confirm',
    FILTER_CANCEL_CLASS_NAME: 'filter-cancel',
    FILTER_COLOR_CONTAINER_CLASS_NAME: 'filter-color-container',
    HIDDEN_CLASS_NAME: 'hidden',
    HIDDEN_CONTAINER_CLASS_NAME: 'hidden-container',
}

export const DEFAULT_CELL_WIDTH = 78;
export const DEFAULT_CELL_HEIGHT = 24;
export const MIN_WIDTH =  20;  // 最小列宽
export const MIN_HEIGHT = 15;  // 最小行高
export const ROW_HEADER_PADDING = 12;  // 有偏差，与预期有不符
export const CELL_PADDING = 2;
export const DEFAULT_CELL_BG_COLOR = 'white';  
export const DEFAULT_FONT_SIZE = 16;
export const DEFAULT_FONT_FAMILY = 'HarmonyOS Sans SC Regular';
export const HEADER_BG_COLOR = 'whitesmoke';
export const GRID_LINE_COLOR = '#ccc';
export const TEXT_COLOR = 'black';
export const DEFAULT_CELL_HIGHLIGHT_COLOR = '#ffeb3b';
export const DEFAULT_SELECTED_BG_COLOR = 'rgba(128,128,128, 0.2)';
/**
 * EMU（English Metric Unit）与 CSS 像素的换算比
 *
 * OOXML 用 EMU 表示长度：1 英寸 = 914400 EMU，而 CSS 以 96 dpi 计，1 英寸 = 96 px，
 * 故 914400 / 96 = 9525。xlsx drawing 的 `colOff`/`rowOff`/`ext/@cx` 全用这个单位。
 * 解析侧只如实保存 EMU 原始值，换算发生在渲染侧（见 `sheets/imageLayout.ts`）。
 */
export const EMU_PER_PIXEL = 9525;
export const DEFAULT_SELECTED_BORDER_COLOR = '#4CAF50';
export const DEFAULT_ANCHOR_SELECTED_BORDER_COLOR = '#A9A9A9';
export const DEFAULT_SELECTED_BORDER_WIDTH = 3;
export const TOOLS_CLASS_NAMES = {   // 定义工具栏类名
    CONTAINER: 'split-container',
    TOP_TOOLS: 'topTools',
    BOTTOM_TOOLS: 'bottomTools',
    LEFT_TOOLS: 'leftTools',
    RIGHT_TOOLS: 'rightTools'
}

export const NAV = {
    PANEL: 'nav-panel',
    PANEL_ITEM: 'panel-item',
    CONTENT: 'nav-content',
    CONTENT_ITEM: 'content-item',
}

export const FILE_TREE_CLASS_NAMES = {   // 定义文件树（「文件」导航页）类名
    CONTAINER: 'file-tree',
    TOOLBAR: 'file-tree-toolbar',
    SEARCH: 'file-tree-search',
    SEARCH_INPUT: 'file-tree-search-input',
    ACTIONS: 'file-tree-actions',
    BUTTON: 'file-tree-button',
    ROOT_ACTIONS: 'file-tree-root-actions',
    BODY: 'file-tree-body',
    EMPTY: 'file-tree-empty',
    NODE: 'file-tree-node',
    ROW: 'file-tree-row',
    TOGGLE: 'file-tree-toggle',
    ICON: 'file-tree-icon',
    LABEL: 'file-tree-label',
    RENAME_INPUT: 'file-tree-rename-input',
    CHILDREN: 'file-tree-children',
    PROPERTY: 'file-tree-property',
    PROPERTY_ROW: 'file-tree-property-row',
    PROPERTY_LABEL: 'file-tree-property-label',
    PROPERTY_VALUE: 'file-tree-property-value',
    SAVE_AS_TABLE: 'file-tree-save-as-table',
    SAVE_AS_EMPTY: 'file-tree-save-as-empty',
    SAVE_AS_SELECT: 'file-tree-save-as-select',
    SAVE_AS_INDEX: 'file-tree-save-as-index',
    SAVE_AS_ACTION: 'file-tree-save-as-action',
    SAVE_AS_DOWNLOAD: 'file-tree-save-as-download',
    SAVE_AS_DOWNLOAD_ALL: 'file-tree-save-as-download-all',
    SAVE_AS_DISABLED: 'is-disabled',
    SAVE_AS_DONE: 'is-done',
    OPEN: 'is-open',
    SELECTED: 'is-selected',
    LOCATED: 'is-located',
    ACTIVED: 'is-actived',
    DRAG_DISABLED: 'is-drag-disabled',
    DRAGGING: 'is-dragging',
    DROP_TARGET: 'is-drop-target',
}

export const PY_EDITOR_CLASS_NAMES = {   // 定义 Python 编辑器（「脚本」导航页）类名
    CONTAINER: 'py-editor',
    TOOLBAR: 'py-editor-toolbar',
    TITLE: 'py-editor-title',
    ACTIONS: 'py-editor-actions',
    RUN: 'py-editor-run',
    STOP: 'py-editor-stop',
    STATUS: 'py-editor-status',
    BODY: 'py-editor-body',
    GUTTER: 'py-editor-gutter',
    CODE: 'py-editor-code',
    HIGHLIGHT: 'py-editor-highlight',
    INPUT: 'py-editor-input',
    OUTPUT: 'py-editor-output',
    OUTPUT_NOTICE: 'py-editor-output-notice',
    OUTPUT_LINE: 'py-editor-output-line',
    OUTPUT_ERROR: 'py-editor-output-error',
}

export const MENU_CLASS_NAMES = {
    ID: 'theMenu',
    CONTAINER: 'menu',
    MENU_CONTENT: 'menu-content',
    ITEM: 'menu-item',
    LINE: 'menu-line',
    SUBMENU: 'submenu',
    SUBMENU_ITEM: 'submenu-item',
}

export const KEYWORDS = ['\n', 'undefined'];

export const FONT_FAMILY_LIST = [
    {todo: 'HarmonyOS Sans SC Regular', text: '鸿蒙黑体'},
    {todo: 'HarmonyOS Sans SC Black', text: '鸿蒙黑体黑'},
    {todo: 'HarmonyOS Sans SC Light', text: '鸿蒙黑体轻'},
    {todo: 'SimSun', text: '宋体'},
    {todo: 'NSimSun', text: '新宋体'},
    {todo: 'FangSong', text: '仿宋'},
    {todo: 'KaiTi', text: '楷体'},
    {todo: 'SimHei', text: '黑体'},
    {todo: 'Microsoft YaHei', text: '微软雅黑'},
    {todo: 'Microsoft JhengHei', text: '微软正黑体'},
    {todo: 'STKaiti', text: '华文楷体'},
    {todo: 'STSong', text: '华文宋体'},
    {todo: 'STFangsong', text: '华文仿宋'},
    {todo: 'Arial', text: 'Arial'},
    {todo: 'Arial Black', text: 'Arial Black'},
    {todo: 'Arial Narrow', text: 'Arial Narrow'},
    {todo: 'Bodoni MT', text: 'Bodoni MT'},
    {todo: 'Book Antiqua', text: 'Book Antiqua'},
    {todo: 'Bookman Old Style', text: 'Bookman Old Style'},
    {todo: 'Calibri', text: 'Calibri'},
    {todo: 'Cambria', text: 'Cambria'},
    {todo: 'Candara', text: 'Candara'},
    {todo: 'Century Gothic', text: 'Century Gothic'},
    {todo: 'Comic Sans MS', text: 'Comic Sans MS'},
    {todo: 'Consolas', text: 'Consolas'},
    {todo: 'Constantia', text: 'Constantia'},
    {todo: 'Corbel', text: 'Corbel'},
    {todo: 'Courier New', text: 'Courier New'},
    {todo: 'Franklin Gothic Medium', text: 'Franklin Gothic Medium'},
    {todo: 'Georgia', text: 'Georgia'},
    {todo: 'Goudy Old Style', text: 'Goudy Old Style'},
    {todo: 'Helvetica', text: 'Helvetica'},
    {todo: 'Impact', text: 'Impact'},
    {todo: 'Lucida Bright', text: 'Lucida Bright'},
    {todo: 'Lucida Console', text: 'Lucida Console'},
    {todo: 'Lucida Sans Unicode', text: 'Lucida Sans Unicode'},
    {todo: 'Microsoft Sans Serif', text: 'Microsoft Sans Serif'},
    {todo: 'Papyrus', text: 'Papyrus'},
    {todo: 'Perpetua', text: 'Perpetua'},
    {todo: 'Rockwell', text: 'Rockwell'},
    {todo: 'Segoe UI', text: 'Segoe UI'},
    {todo: 'Tahoma', text: 'Tahoma'},
    {todo: 'Times New Roman', text: 'Times New Roman'},
    {todo: 'Trebuchet MS', text: 'Trebuchet MS'},
    {todo: 'Verdana', text: 'Verdana'},
    {todo: 'Webdings', text: 'Webdings'},
    {todo: 'Wingdings', text: 'Wingdings'},
    {todo: 'Bahnschrift', text: 'Bahnschrift'},
    {todo: 'Cascadia Code', text: 'Cascadia Code'},
    {todo: 'Cascadia Mono', text: 'Cascadia Mono'}
];

/**
 * 单元格值属性类型枚举
 */
enum ValueProperty {
    CHAR_CHANGED = 'char:changed',
    CHAR_FONT_FAMILY_CHANGED = 'char:fontFamily:changed',
    CHAR_FONT_SIZE_CHANGED = 'char:fontSize:changed',    
    CHAR_FONT_WEIGHT_CHANGED = 'char:fontWeight:changed',
    CHAR_FONT_STYLE_CHANGED = 'char:fontStyle:changed',
    CHAR_FONT_COLOR_CHANGED = 'char:fontColor:changed',
    CHAR_UNDERLINE_CHANGED = 'char:underline:changed',
    CHAR_STRIKETHROUGH_CHANGED = 'char:strikethrough:changed'
}

/**
 * 单元格属性类型枚举
 */
enum CellProperty {
    // CELL_STYLE_CHANGED = 'cell:style:changed',
    ACTIVED_CELL_CHANGED = 'actived:changed',
    COLSPAN_CHANGED = 'cell:colspan:changed',
    ROWSPAN_CHANGED = 'cell:rowspan:changed',
    BORDER_TOP_WIDTH_CHANGED = 'cell:topBorderWidth:changed',
    BORDER_BOTTOM_WIDTH_CHANGED = 'cell:bottomBorderWidth:changed',
    BORDER_LEFT_WIDTH_CHANGED = 'cell:leftBorderWidth:changed',
    BORDER_RIGHT_WIDTH_CHANGED = 'cell:rightBorderWidth:changed',
    // FONT_FAMILY_CHANGED = 'cell:fontFamily:changed',
    // FONT_SIZE_CHANGED = 'cell:fontSize:changed',    
    // FONT_WEIGHT_CHANGED = 'cell:fontWeight:changed',
    // FONT_STYLE_CHANGED = 'cell:fontStyle:changed',
    // FONT_COLOR_CHANGED = 'cell:fontColor:changed',
    // UNDERLINE_CHANGED = 'cell:underline:changed',
    // STRIKETHROUGH_CHANGED = 'cell:strikethrough:changed',
    TEXT_ALIGN_CHANGED = 'cell:textAlign:changed',
    ALIGN_ITEMS_CHANGED = 'cell:alignItems:changed',
    BACKGROUND_COLOR_CHANGED = 'cell:backgroundColor:changed',
    NUMBER_FORMAT_CHANGED = 'cell:numberFormat:changed',
    BORDER_COLOR_CHANGED = 'cell:borderColor:changed',
    FILTER_CHANGED = 'cell:filter:changed',
    WRAP_CHANGED = 'cell:wrap:changed',
    LETTER_SPACING_CHANGED = 'cell:letterSpacing:changed',
    LINE_SPACING_CHANGED = 'cell:lineSpacing:changed',
}

/**
 * 工作表属性类型枚举
 */
enum SheetProperty {
    IS_EDITTING_CHANGED = 'sheet:isEditting:changed',
    SHOW_GRID_LINES_CHANGED = 'sheet:showGridLines:changed',
    ACTIVED_SHEET_NAME_CHANGED = 'sheet:activedSheetName:changed',
    VALUES_CHANGED = 'sheet:values:changed',
    SELECTION_CHANGED = 'sheet:selection:changed',
    ANCHOR_SELECTION_CHANGED = 'sheet:anchorSelection:changed',
    /** 复制/剪切的源范围变化（CopyMarquee 据此显示/隐藏流动虚线框） */
    COPY_RANGE_CHANGED = 'sheet:copyRange:changed',
    ZOOM_CHANGED = 'sheet:zoom:changed',
    RESIZE_CHANGED = 'sheet:resize:changed',
    OFFSETX_CHANGED = 'offsetx:changed',
    OFFSETY_CHANGED = 'offsety:changed',
    CURSOR_STATE_CHANGED = 'cursor:state:changed',
    BRUSH_MODE_CHANGED = 'brush:mode:changed',
    ALL_COL_WIDTH_CHANGED = 'allcolwidth:changed',
    ALL_ROW_HEIGHT_CHANGED = 'allrowheight:changed',
    COLUMNS_CHANGED = 'columns:changed',
    ROWS_CHANGED = 'rows:changed',
    CURSOR_POSITION_CHANGED = 'cursor:position:changed',
    /** 撤销/重做快照回写完成：此时数据已完整恢复到目标状态，
     *  供依赖 ACTIVED_CELL_CHANGED 等差量事件刷新外观的分区（工具栏按钮等）
     *  在回写未触发差量事件（如活动单元格未变化）时也能同步状态 */
    SNAPSHOT_RESTORED = 'sheet:snapshot:restored',
    /** 工作表装载后广播其 Python 脚本内容（无脚本时为 null），Python 编辑器据此切换显示 */
    SHEET_SCRIPT_CHANGED = 'sheet:script:changed',
}

enum VisibleViewProperty {
    ROW_HEADER_WIDTH_CHANGED = 'visibleView:rowHeaderWidth:changed',
    COL_HEADER_HEIGHT_CHANGED = 'visibleView:colHeaderHeight:changed',
    SHEET_WIDTH_CHANGED = 'visibleView:sheetWidth:changed',
    SHEET_HEIGHT_CHANGED = 'visibleView:sheetHeight:changed',
}

// 值属性变化回调函数类型
// export type ChangeCallback = (property: string, oldValue: any, newValue: any) => void;


// 事件名称常量, 包含工作表、单元格和值属性变化事件
export const DataEvents = Object.assign({}, SheetProperty, CellProperty, ValueProperty, VisibleViewProperty);
export type DataEvents = SheetProperty | CellProperty | ValueProperty | VisibleViewProperty;

/* ════════════════════════════════════════════════════════════
 *  文档（Word）实现常量
 *
 *  文档工作区（`core/document/*`）里的尺子、归档、渲染、滚动条与转换器共用这些值。
 *  集中在此而不是散落在各文件顶部，是为了让「调一次缩放步长 / 存档版本 / 纸张尺寸」
 *  只改一处：这些值彼此约束（例如存档版本与装载校验、缩放步长与滑块档位），
 *  分散后极易改一处忘另一处。
 * ════════════════════════════════════════════════════════════ */

/** 文档存档版本（结构不兼容时递增，装载侧据此决定迁移或丢弃） */
export const DOC_SNAPSHOT_VERSION = 1;

/** 文档默认字体族（与表格默认值同源，避免两种载体之间视觉割裂） */
export const DEFAULT_DOC_FONT_FAMILY = 'HarmonyOS Sans SC Regular';

/** 文档默认字号（CSS 像素；对应 12pt） */
export const DEFAULT_DOC_FONT_SIZE = 16;

/**
 * 每级缩进的像素步长
 *
 * 渲染（还原 `indent`）与命令（写回 DOM）**必须共用同一个常量**：
 * 两处各写一个 24，改一处忘另一处就会出现「点了有缩进、重渲染后跳一格」。
 */
export const DOC_INDENT_STEP_PX = 24;

/**
 * 纸张尺寸（CSS 像素，96dpi）
 *
 * 只用 A4 / A5 / A3 / Letter / Legal 五档 —— 与 `DocPageSetup.paperSize` 的取值一一对应。
 * 纵向为「宽 × 高」，横向时两者互换。
 */
export const DOC_PAPER_SIZES: Record<'A4' | 'A5' | 'A3' | 'Letter' | 'Legal', { width: number; height: number }> = {
    A4: { width: 794, height: 1123 },
    A5: { width: 559, height: 794 },
    A3: { width: 1123, height: 1587 },
    Letter: { width: 816, height: 1056 },
    Legal: { width: 816, height: 1344 },
};

/**
 * 页边距预设（CSS 像素，四边同值）
 *
 * 常规 96px ≈ 1 英寸（25.4mm），与导出 docx 时 `sectPr` 的页边距同一口径。
 */
export const DOC_MARGIN_PRESETS: Record<'normal' | 'wide' | 'narrow', number> = {
    normal: 96,
    wide: 120,
    narrow: 56,
};

/** 缩放百分比下限（与缩放滑块的档位同一口径） */
export const DOC_ZOOM_PERCENT_MIN = 50;
/** 缩放百分比上限 */
export const DOC_ZOOM_PERCENT_MAX = 400;
/** Ctrl + 滚轮的每格缩放增量（比例，不是百分比） */
export const DOC_ZOOM_WHEEL_STEP = 0.1;

/** 文档落库防抖（毫秒） */
export const DOC_SAVE_DEBOUNCE = 500;
/** 连续键入的「一阵」窗口（毫秒）：窗口内只按一次归档，避免每个键都写库 */
export const DOC_TYPING_BURST = 700;

/** 文档滚动条单次点击的滚动距离（像素） */
export const DOC_SCROLL_STEP_PX = 60;
/** 长按滚动按钮后开始连续滚动的延迟（毫秒） */
export const DOC_SCROLL_STEP_DELAY = 300;
/** 长按期间每次连滚的间隔（毫秒） */
export const DOC_SCROLL_STEP_INTERVAL = 50;
/** 滚动条滑块的最小尺寸（像素） */
export const DOC_SCROLL_THUMB_MIN_SIZE = 30;
/** 打印预览可用区四周的留白（像素） */
export const DOC_PRINT_SCROLL_PADDING = 48;

/** 分页留白块的属性名（同步回块树时据此跳过它） */
export const DOC_PAGE_GAP_ATTR = 'data-page-gap';
/**
 * 逐块排页的块数上限
 *
 * 超过它就退化成「只算页数、不插留白」：分页是显示层的东西，不该把输入延迟拖下水。
 */
export const DOC_MAX_PAGINATE_BLOCKS = 1200;

/** Markdown 源码编辑回写模型的防抖（毫秒） */
export const DOC_MARKDOWN_APPLY_DEBOUNCE = 260;

/** 毫米与 CSS 像素的换算比（96dpi：1 英寸 = 25.4mm = 96px） */
export const DOC_MM_PER_PX = 25.4 / 96;
/** CSS 像素与毫米的换算比（{@link DOC_MM_PER_PX} 的倒数） */
export const DOC_PX_PER_MM = 96 / 25.4;
/** 标尺最小刻度间隔（毫米） */
export const DOC_RULER_MINOR_STEP_MM = 5;
/** 标尺主刻度间隔（毫米，刻度线加长） */
export const DOC_RULER_MAJOR_STEP_MM = 10;
/** 标尺数字标签间隔（毫米） */
export const DOC_RULER_LABEL_STEP_MM = 20;

/**
 * 鼠标位置枚举
 */
export enum MouseLocation {
    IN_All_SELECT = 'inAllSelect',
    IN_COL_HEADER = 'inColHeader',
    IN_COL_HEADER_SPLIT = 'inColHeaderSplit',
    IN_COL_FROZEN_SPLIT = 'inColFrozenSplit',
    IN_ROW_HEADER = 'inRowHeader',
    IN_ROW_HEADER_SPLIT = 'inRowHeaderSplit',
    IN_ROW_FROZEN_SPLIT = 'inRowFrozenSplit',
    IN_CELL = 'inCell',
    IN_ANCHOR = 'inAnchor',
}
// export const SHEET_TODO = {
//     BRUSH: 'brush',
//     COPY: 'copy',
//     CUT: 'cut',
//     PASTE: 'paste',
//     SAVE: 'save',
//     FONT_SELECT: 'font-select',
//     FONT_SIZE_SELECT: 'font-size-select',
//     BOLD: 'bold',
//     ITALIC: 'italic',
//     UNDERLINE: 'underline',
//     STRIKETHROUGH: 'strikethrough',    
//     BORDER: 'border',
//     BORDER_ALL: 'border-all',
//     BORDER_OUTER: 'border-outer',
//     BORDER_WIDE_OUTER: 'border-wide-outer',
//     BORDER_BOTTOM: 'border-bottom',
//     BORDER_TOP: 'border-top',
//     BORDER_LEFT: 'border-left',
//     BORDER_RIGHT: 'border-right',
//     PAINTBUCKET: 'paintbucket',
//     TEXT_COLOR: 'text-color',
// }
