/**
 * 行文本信息接口
 * @interface LineText
 */
export interface LineText {
    /** 文本内容 */
    text: any;
    /** 行宽（包含字间距） */
    lineWidth: number;
    /** 行高（包含行间距） */
    lineHeight: number;
    /** 相对于单元格每行左边界的X坐标 */
    x: number;
    /** 相对于单元格每行顶部边界的Y坐标 */
    y: number;
    /** 相对于编辑框每行左边界的X坐标 */
    editX: number;
    /** 相对于编辑框每行顶部边界的Y坐标 */
    editY: number;
}

/**
 * 单元格内容接口
 * @interface CellContent
 */
export interface CellContent {
    /** 相对于单元格每行左边界的X坐标 */
    editLeft: number;
    /** 相对于单元格每行顶部边界的Y坐标 */
    editTop: number;
    /** 文本宽度 */
    charsWidth: number;
    /** 文本高度 */
    charsHeight: number;
    /** 内容总宽度 */
    contentWidth: number;
    /** 内容总高度 */
    contentHeight: number;
    /** 行文本信息数组 */
    lines: LineText[];
}


/**
 * 编辑状态下光标位置接口
 */
export interface CursorPosition {
    /** 行索引 */
    lineIndex: number;
    /** 字符索引 */
    charIndex: number;
    /** 编辑区域 X 坐标 */
    editX: number;
    /** 编辑区域 Y 坐标 */
    editY: number;
}

/**
 * 编辑状态下选择范围接口
 */
export interface SelectionRange {
    /** 选择起始位置 */
    start: CursorPosition;
    /** 选择结束位置 */
    end: CursorPosition;
}

/**
 * 字体样式接口
 */
export interface FontStyle {
    /** 字体名称 */
    fontFamily?: string;
    /** 字体大小 */
    fontSize?: number;
    /** 字体粗细 */
    fontWeight?: boolean;
    /** 字体样式 */
    fontStyle?: boolean;
    /** 字体颜色 */
    fontColor?: string;
    /** 字体下划线 */
    underline?: boolean;
    /** 字体删除线 */
    strikethrough?: boolean;
    /** 字体对齐方式 */
    textAlign?: string;
    /** 字体垂直对齐方式 */
    alignItems?: string;
    /** 背景颜色 */
    backgroundColor?: string;
    /** 字体间距 */
    letterSpacing?: number;
    /** 行间距 */
    lineSpacing?: number;
    /** 是否换行 */
    wrap?: boolean;
    /** 是否合并单元格 */
    merge?: boolean;
}

/**
 * 边框样式接口
 */
export interface BorderStyle {
    /** 上边框宽度 */
    borderTopWidth?: number;
    /** 左边框宽度 */
    borderLeftWidth?: number;
    /** 下边框宽度 */
    borderBottomWidth?: number;
    /** 右边框宽度 */
    borderRightWidth?: number;
}

/**
 * 单元格边框样式接口
 */
export enum CellBorderStyle {
    BorderNone = "border-none",
    BorderAll = "border-all",
    BorderOuter = "border-outer",
    BorderWideOuter = "border-wide-outer",
    BorderBottom = "border-bottom",
    BorderTop = "border-top",
    BorderLeft = "border-left",
    BorderRight = "border-right",
}