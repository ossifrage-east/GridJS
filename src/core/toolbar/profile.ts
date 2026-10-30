/*  # 工具栏·工作区档案 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目目标: 把「哪类工作区该看到哪些工具」这张表从工具栏的布局逻辑里拆出来。
    # 本文件职责: 纯数据 —— 声明每种 {@link WorkspaceKind} 下「显示哪些根分类、
    #             每个分类挂哪些分区、分区叫什么名字」。**不含任何 DOM 操作**，
    #             也不 import 任何分区实现（只引用它们的类型）。
    # 依赖方向: contract/editHost（只有 WorkspaceKind 一个类型）+ 各分区的类型。
    #           **不 import dataArchitecture / document**。
    # 为什么值得单独一个文件: 「Word 页签里该有哪些按钮」是一条会反复调整的产品规则，
    #             而「怎么量宽度、怎么收起三级菜单」是稳定的布局算法。两者混在一个文件里时，
    #             每次调整按钮都要重新读一遍 500 行布局代码。
    # 参照对象: Office Word 的「开始」页签 —— 剪贴板 / 字体 / 段落 / 样式四组。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

import { WorkspaceKind } from "../contract/editHost";
import { TOOLBAR_TABS } from "../constant";
import { SectionParagraph } from "./start/sectionParagraph";
import { SectionBrush } from "./start/sectionBrush";
import { SectionFont } from "./start/sectionFont";
import { SectionAlign } from "./start/sectionAlign";
import { SectionProcess } from "./start/sectionProcess";
import { SectionPrint } from "./pages/sectionPrint";
import { SectionDocView } from "./pages/sectionDocView";
import { SectionDocPage } from "./pages/sectionDocPage";

/**
 * 一个分区在某个档案里的条目
 *
 * `title` 是「紧凑模式下的一级菜单标签」，因此**允许随档案变化**：
 * 同一个 `SectionBrush` 在表格里叫「格式刷」，在文档里只剩复制 / 剪切 / 粘贴，
 * 那时它应该叫「剪贴板」。
 */
export interface ProfileSection {
    /** 分区根元素 */
    el: HTMLElement;
    /** 分区名（紧凑模式下作为一级菜单标签） */
    title: string;
}

/**
 * 一个根分类页签在某个档案里的条目
 */
export interface ProfileTab {
    /** 根分类名（必须是 {@link TOOLBAR_TABS} 的值） */
    tab: string;
    /** 该分类下挂的分区（顺序即显示顺序） */
    sections: ProfileSection[];
}

/** 一套完整的工作区档案：没出现在这里的根分类会被隐藏 */
export type ToolbarProfile = ProfileTab[];

/** 装配档案所需的全部分区实例 */
export interface ToolbarSections {
    /** 剪贴板 / 格式刷 */
    brush: SectionBrush;
    /** 字体 */
    font: SectionFont;
    /** 对齐 */
    align: SectionAlign;
    /** 数据处理（求和 / 排序 / 筛选 / 查找） */
    process: SectionProcess;
    /** 打印设置 */
    print: SectionPrint;
    /** 段落（正文 / 标题 / 项目符号 / 编号 / 缩进） */
    para: SectionParagraph;
    /** 文档视图（Markdown 源码开关） */
    docView: SectionDocView;
    /** 文档页面设置（方向 / 纸张 / 页边距 / 缩放 / 打印） */
    docPage: SectionDocPage;
}

/**
 * 构建两种工作区的档案
 *
 * ## 表格（参照 Excel 的「开始」）
 * 格式刷 / 字体样式 / 对齐方式 / 数据处理；「页面」里是打印设置。
 *
 * ## 文档（参照 Word 的「开始」+「视图」）
 * - 剪贴板：复用格式刷分区的复制 / 剪切 / 粘贴（格式刷按钮本身在文档里被隐藏）
 * - 字体：复用字体样式分区（边框按钮在文档里被隐藏）
 * - 段落对齐：复用对齐方式分区（垂直对齐 / 自动换行 / 合并单元格在文档里被隐藏）
 * - 段落样式：正文 / 标题 1–6 / 引用 / 代码块 + 项目符号 / 编号 + 缩进
 * - 视图：Markdown 源码开关
 *
 * **「页面」页签挂的是文档自己的页面设置**：页边距 / 纸张 / 缩放写 `DocMeta.pageSetup`，
 * 打印走文档视图的 iframe 打印 —— 表格的「打印设置」整组绑在 `Printer` 上，
 * 在文档里点不动，因此文档档案挂 `SectionDocPage` 而不是 `SectionPrint`。
 * @param sections - 分区实例集合
 * @returns 按工作区类型索引的档案
 */
export function buildProfiles(sections: ToolbarSections): Record<WorkspaceKind, ToolbarProfile> {
    const sheet: ToolbarProfile = [
        {
            tab: TOOLBAR_TABS.START,
            sections: [
                { el: sections.brush.getElements().sectionBrush, title: '格式刷' },
                { el: sections.font.getElements().sectionFont, title: '字体样式' },
                { el: sections.align.getElements().sectionAlign, title: '对齐方式' },
                { el: sections.process.getElement().sectionProcess, title: '数据处理' },
            ],
        },
        {
            tab: TOOLBAR_TABS.PAGE,
            sections: [{ el: sections.print.getElements().sectionPrint, title: '打印设置' }],
        },
    ];

    const doc: ToolbarProfile = [
        {
            tab: TOOLBAR_TABS.START,
            sections: [
                // 同一个分区换了名字：文档里它只剩剪贴板三件套
                { el: sections.brush.getElements().sectionBrush, title: '剪贴板' },
                { el: sections.font.getElements().sectionFont, title: '字体' },
                { el: sections.align.getElements().sectionAlign, title: '段落对齐' },
                { el: sections.para.getElements().sectionParagraph, title: '段落样式' },
            ],
        },
        {
            tab: TOOLBAR_TABS.VIEW,
            sections: [{ el: sections.docView.getElements().sectionDocView, title: '文档视图' }],
        },
        {
            // 与表格「页面」同名的根分类：Word 的「布局」里有同样一组纸张设置，
            // 换到文档这里写的是 DocMeta.pageSetup（而不是表格的 Printer）
            tab: TOOLBAR_TABS.PAGE,
            sections: [{ el: sections.docPage.getElements().sectionDocPage, title: '页面设置' }],
        },
    ];

    return { sheet, doc };
}
