/*  # 文件树菜单项：新建 / 另存为（含 hover 二级菜单）
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { createDiv, createElement, createSpan } from "../../../utils/dom";
import { MENU_CLASS_NAMES } from "../../constant";
import { Menu, MenuContent, MenuItemSpec } from "../menu";
import { NEW_FILE_MENU_ITEMS, TreeNodeView } from "./fileTreeModel";

/**
 * 构建带 hover 二级菜单的菜单项（「新建文件」「另存为」共用）
 *
 * 二级菜单沿用主表格右键菜单的既有做法：独立容器 append 到 body
 * （主菜单 max-height + overflow 会被裁剪），定位到父项右侧并重叠 1px
 * 消除 hover 间隙，mouseleave 延迟 120ms 关闭以免鼠标穿越间隙时误关。
 * @param {Menu | null | undefined} menu - 主菜单实例（用于关闭与兜底清理）
 * @param {string} iconClass - 一级项图标类名
 * @param {string} text - 一级项文字
 * @param {MenuItemSpec[]} items - 二级菜单项
 * @param {(todo: string) => void} onPick - 二级菜单项点击回调（菜单已关闭）
 * @returns {HTMLElement} 菜单项元素
 */
export function createHoverSubmenuItem(
    menu: Menu | null | undefined,
    iconClass: string,
    text: string,
    items: MenuItemSpec[],
    onPick: (todo: string) => void
): HTMLElement {
    const wrap = createDiv({
        className: 'menu-item menu-item-has-submenu',
        style: { position: 'relative' },
    });
    const iconEl = createElement('i', { className: iconClass });
    const textSpan = createSpan({ className: 'text', textContent: text });
    const arrowSpan = createSpan({ textContent: '▶' });
    arrowSpan.style.cssText = 'margin-left:auto; font-size:10px; color:#999;';
    wrap.append(iconEl, textSpan, arrowSpan);

    let submenuEl: HTMLElement | null = null;
    let closeTimer: ReturnType<typeof setTimeout> | null = null;
    const cancelPendingClose = (): void => {
        if (closeTimer !== null) {
            clearTimeout(closeTimer);
            closeTimer = null;
        }
    };
    const onMenuClose = (): void => closeSubmenu();
    const closeSubmenu = (): void => {
        cancelPendingClose();
        // 每次展开都注册一次兜底清理，关闭时同步注销，避免反复右键累积监听
        menu?.off('close', onMenuClose);
        if (submenuEl?.isConnected) submenuEl.remove();
        submenuEl = null;
    };

    wrap.addEventListener('mouseenter', () => {
        cancelPendingClose();
        if (submenuEl) return;   // 已展开则跳过
        const subContent = new MenuContent({
            items,
            onClick: (todo: string) => {
                closeSubmenu();
                menu?.closeMenu();
                onPick(todo);
            },
        });
        submenuEl = subContent.getElements().container;
        submenuEl.classList.add(MENU_CLASS_NAMES.CONTAINER);
        // 内联覆盖 .menu 的 opacity:0 / visibility:hidden / transition，
        // 否则子菜单需等淡入动画，感知为延迟显示
        submenuEl.style.cssText = 'position:absolute; display:flex; opacity:1; visibility:visible; transition:none; min-width:160px; box-shadow:0 4px 8px rgba(0,0,0,0.15); z-index:1002;';
        document.body.appendChild(submenuEl);
        // 定位到父项右边缘（重叠 1px 消除 hover gap），空间不足时翻到左侧
        const wrapRect = wrap.getBoundingClientRect();
        const subRect = submenuEl.getBoundingClientRect();
        let left = wrapRect.right - 1;
        let top = wrapRect.top;
        if (left + subRect.width > window.innerWidth) {
            left = wrapRect.left - subRect.width + 1;
        }
        if (top + subRect.height > window.innerHeight) {
            top = window.innerHeight - subRect.height - 4;
        }
        if (top < 0) top = 4;
        submenuEl.style.left = `${left}px`;
        submenuEl.style.top = `${top}px`;
        submenuEl.addEventListener('mouseenter', cancelPendingClose);
        // 主菜单关闭时兜底清理，防止子菜单残留（关闭后由 closeSubmenu 注销）
        menu?.on('close', onMenuClose);
    });

    wrap.addEventListener('mouseleave', () => {
        // 记录定时器，鼠标在延迟内回到这一项（或移入子菜单）时取消关闭
        cancelPendingClose();
        closeTimer = setTimeout(closeSubmenu, 120);
    });
    return wrap;
}

/**
 * 构建「新建文件」右键菜单项（含 hover 展开的二级菜单）
 * @param {TreeNodeView} view - 被右键的节点
 * @param {Menu | null | undefined} menu - 主菜单实例
 * @param {(todo: string) => void} onPick - 选中文件类型后的回调
 * @returns {HTMLElement} 菜单项元素
 */
export function createNewFileSubmenuItem(
    view: TreeNodeView,
    menu: Menu | null | undefined,
    onPick: (todo: string) => void
): HTMLElement {
    return createHoverSubmenuItem(menu, 'icon-new-file', '新建文件', NEW_FILE_MENU_ITEMS, onPick);
}

/**
 * 构建「另存为」右键菜单项（含 hover 展开的二级菜单）
 *
 * 后三项（pptx / pdf）暂未实现，置灰不可点；
 * 「另存为xlsx」打开导出清单对话框（列出当前选中的 `.grid` 文件）；
 * 「另存为docx」与「导出为 Word」同一实现。
 * @param {boolean} hasText - 选中里是否含 `.text` 文档（没有则 docx 项置灰）
 * @param {Menu | null | undefined} menu - 主菜单实例
 * @param {(todo: string) => void} onPick - 选中项回调
 * @returns {HTMLElement} 菜单项元素
 */
export function createSaveAsSubmenuItem(
    hasText: boolean,
    menu: Menu | null | undefined,
    onPick: (todo: string) => void
): HTMLElement {
    const items: MenuItemSpec[] = [
        { todo: 'saveAsXlsx', text: '另存为xlsx' },
        // 文档 → docx：选中里没有 `.text` 时置灰（与「导出为 Word」同一实现）
        { todo: 'saveAsDocx', text: '另存为docx', disabled: !hasText },
        { todo: 'saveAsPptx', text: '另存为pptx', disabled: true },
        { todo: 'saveAsPdf', text: '另存为pdf', disabled: true },
    ];
    return createHoverSubmenuItem(menu, 'icon-file', '另存为', items, onPick);
}
