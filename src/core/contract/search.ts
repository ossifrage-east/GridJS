/*  # 编辑契约层 · 查找替换契约 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目目标: 让「查找 / 替换」对话框只认识一个极窄的查询接口，不认识表格的数据集合。
    # 本文件职责: 定义查找所需的矩形（{@link SearchRect}）、匹配选项（{@link SearchMatchOptions}）、
    #             命中项（{@link SearchHit}）与查询目标（{@link SearchTarget}）。**只声明形状**。
    # 依赖方向: 只允许 import 同目录 editHost；**不得 import dataArchitecture / toolbar / sheets**。
    # 为什么是「接口」而不是「命令」:
    #   查找要**把结果带回来**（命中列表、替换处数），而 {@link EditHost.execute} 的返回是 void；
    #   更重要的是「跳转命中」会改写选区、需要订阅选区变化，这些都不是一次命令能表达的。
    #   因此这里退化为一个「查询目标」窄接口：对话框只做 UI，数据侧实现方（表格数据集合）
    #   结构化满足它即可，无需适配器、也无需往命令集里塞会返回值的新命令。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { EditingStateSource } from './editHost';

/**
 * 查询矩形（闭区间，1-based）
 *
 * 与表格侧的 `SheetRect` 结构等价，但刻意在契约层独立声明：对话框不能依赖表格数据模型。
 */
export interface SearchRect {
    /** 起始列号（1-based，含） */
    startCol: number;
    /** 起始行号（1-based，含） */
    startRow: number;
    /** 结束列号（1-based，含） */
    endCol: number;
    /** 结束行号（1-based，含） */
    endRow: number;
}

/** 匹配选项（与表格侧的 `SearchMatchOptions` 结构等价） */
export interface SearchMatchOptions {
    /** 区分大小写（默认 false） */
    matchCase?: boolean;
    /** 单元格内容**完全匹配**才算命中（默认 false，即「包含即命中」） */
    matchEntire?: boolean;
    /** 通配符模式：`?` 匹配单个字符、`*` 匹配任意长度字符串（默认 false） */
    useWildcard?: boolean;
}

/** 一条查询命中 */
export interface SearchHit {
    /** 单元格名称（如 `'B3'`） */
    cell: string;
    /** 列号（1-based） */
    col: number;
    /** 行号（1-based） */
    row: number;
    /** 命中时的单元格全文 */
    text: string;
}

/** 单格范围扩成整表后的结果（`anchor` 是查找起点） */
export interface SearchRectExpansion {
    /** 展开后的整表范围 */
    rect: SearchRect;
    /** 起点（当次查询从它往下走） */
    anchor: { col: number; row: number };
}

/**
 * 查找 / 替换目标
 *
 * 对话框的**唯一**数据入口。它同时充当「编辑态事件源」（`CommandBus` / `DataCollection`
 * 都天然满足 `on` / `off`），因此隐藏时要摘 `SELECTION_CHANGED` 监听也无须认识具体类型。
 *
 * 【为什么把这么多方法摊在接口上】它们本来就是「一次查询会话」的完整生命周期：
 * 解析范围 → 算命中 → 跳转 → 替换 → 局部复查。拆成多个小接口只会让对话框持有
 * 三四个对象，反而更难看清「谁负责什么」。
 */
export interface SearchTarget extends EditingStateSource {
    /** 当前选区（如 `'A1:C9'`；没有选区时为 `null`） */
    readonly selection: string | null;

    /**
     * 解析查询范围
     *
     * 优先级：显式范围串 → 当前选区 → 有内容区域。非法串回退到选区（不静默扫全表）。
     * @param range 范围串（如 `'A1:C20'`），省略时按选区
     * @returns 规范化后的矩形；无可用范围时 `null`
     */
    resolveSearchRect(range?: string): SearchRect | null;

    /** 把矩形格式化成范围串（`A1:C20`） */
    formatRect(rect: SearchRect): string;

    /** 列号 / 行号 → 单元格名（`B3`） */
    getCellName(col: number, row: number): string;

    /**
     * 把「单格范围」解释为「整表范围 + 起点锚」
     *
     * 选区只有一个单元格时，查找按整表进行、从该格往下走；非单格返回 `null`。
     */
    expandSingleCellRect(rect: SearchRect): SearchRectExpansion | null;

    /**
     * 范围内全量命中（行优先序）
     * @param rect 范围
     * @param keyword 关键字（空串返回空数组）
     * @param options 匹配选项
     * @param limit 命中上限（缺省不限）
     */
    findAllInRect(rect: SearchRect, keyword: string, options?: SearchMatchOptions, limit?: number): SearchHit[];

    /**
     * 单个单元格是否仍命中关键字（替换后局部复查，避免整表重扫）
     * @param cellName 单元格名
     */
    cellMatches(cellName: string, keyword: string, options?: SearchMatchOptions): SearchHit | null;

    /**
     * 范围替换
     * @param rect 范围
     * @param keyword 关键字
     * @param replacement 替换文本
     * @param options 匹配选项
     * @param targets 只替换这些命中项；`null` / 省略 = 范围内全部命中
     * @param label 撤销记录的动作名
     * @returns 实际替换的处数
     */
    replaceInRect(
        rect: SearchRect,
        keyword: string,
        replacement: string,
        options?: SearchMatchOptions,
        targets?: SearchHit[] | null,
        label?: string,
    ): number;

    /** 跳转到命中项（收敛选区并滚动入视野） */
    gotoSearchHit(hit: SearchHit): void;
}
