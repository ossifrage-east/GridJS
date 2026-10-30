/*  # 文件树「属性」对话框
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { createDiv } from "../../../utils/dom";
import { extractExtension } from "../../../utils/fileName";
import { FILE_TREE_CLASS_NAMES as TREE } from "../../constant";
import { Dialog } from "../dialog";
import { TreeNodeView, formatTime } from "./fileTreeModel";

/**
 * 构建属性对话框的正文（定义列表：字段名 + 字段值）
 *
 * 五项内容与需求一一对应：
 * a. **名称及后缀** —— 文件显示 `名称` + `后缀` 两行；文件夹无后缀则不显示该行；
 * b. **最初建立时间**；
 * c. **何人何时做最后修改**（人 + 时间两行）；
 * d. **是否可以编辑**；
 * e. **创建人**。
 * @param {TreeNodeView} view - 目标节点
 * @param {string} path - 节点完整路径
 * @returns {HTMLElement} 正文容器
 */
export function buildPropertyContent(view: TreeNodeView, path: string): HTMLElement {
    const name = view.label.textContent ?? '';
    const ext = extractExtension(name);
    const rows: Array<[string, string]> = [];

    // a. 名称及后缀（文件夹无后缀 → 不显示「后缀」这一行）
    rows.push(['名称', name]);
    if (!view.isFolder && ext) rows.push(['后缀', `.${ext}`]);
    rows.push(['位置', path]);

    // b. 最初建立时间
    rows.push(['最初建立时间', formatTime(view.createdAt)]);

    // c. 何人何时做最后修改
    rows.push(['最后修改人', view.lastModifiedBy || '—']);
    rows.push(['最后修改时间', formatTime(view.lastModifiedAt)]);

    // d. 是否可以编辑
    rows.push(['是否可以编辑', view.editable ? '是' : '否']);

    // e. 创建人
    rows.push(['创建人', view.createdBy || '—']);

    const table = createDiv({ className: TREE.PROPERTY });
    rows.forEach(([label, value]) => {
        const row = createDiv({ className: TREE.PROPERTY_ROW });
        row.appendChild(createDiv({ className: TREE.PROPERTY_LABEL, textContent: label }));
        row.appendChild(createDiv({ className: TREE.PROPERTY_VALUE, textContent: value }));
        table.appendChild(row);
    });
    return table;
}

/**
 * 打开（或复用）「属性」对话框
 *
 * 内容为结构化占位：数据全部取自节点自身（名称 / 路径 / 创建与修改的时间与人 /
 * 可编辑性），不依赖任何服务端。
 * @param {Dialog | null} current - 已有的对话框实例（复用，避免反复创建 DOM）
 * @param {TreeNodeView} view - 目标节点
 * @param {string} path - 节点完整路径
 * @returns {Dialog} 对话框实例（调用方保存以便复用）
 */
export function showPropertyDialog(current: Dialog | null, view: TreeNodeView, path: string): Dialog {
    let dialog = current;
    if (!dialog) {
        dialog = new Dialog({
            title: '属性',
            buttons: [{ text: '确定', variant: 'primary', autoFocus: true }],
            size: 'md',
        });
    }
    dialog.setTitle(`${view.isFolder ? '文件夹' : '文件'}属性`);
    dialog.setContent(buildPropertyContent(view, path));
    dialog.show();
    return dialog;
}
