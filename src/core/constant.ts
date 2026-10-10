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
    SHEETS: 'sheets',
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
}
export const NAV_PANEL: string[] = ['文件', '工具', '脚本']
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
    BODY: 'file-tree-body',
    EMPTY: 'file-tree-empty',
    NODE: 'file-tree-node',
    ROW: 'file-tree-row',
    TOGGLE: 'file-tree-toggle',
    ICON: 'file-tree-icon',
    LABEL: 'file-tree-label',
    CHILDREN: 'file-tree-children',
    OPEN: 'is-open',
    SELECTED: 'is-selected',
    LOCATED: 'is-located',
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
