/*  # 打印预览外壳 TS 文件
    # 项目名称: silkspaces（藕丝空间）
    # 项目目标: 让「表格打印预览」与「文档打印预览」共用**同一套外壳**（全屏遮罩 +
    #           标题栏 + 左侧预览区 + 右侧设置面板 + 页脚翻页与缩放）。
    # 本文件职责: 只吐 HTML 字符串与两个常量，**不含任何表格 / 文档语义**：
    #             内容样式由调用方以 `extraCss` 注入（表格注入 table 样式，
    #             文档注入正文样式），两者因此长得一模一样而互不知道对方。
    # 为什么抽出来: 外壳原本是 Printer.printPreview 里的一段内联模板。文档预览要
    #             「参照表格的实现」，复制一份等于两处样式各自漂移；抽出来后
    #             改一次外壳，两类预览同时生效。
    # 依赖方向: 无 import（纯字符串工厂）。
    # author: 东方鹗
    # B站: https://space.bilibili.com/194739739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
    # 项目许可证: MIT 许可证
*/

/** 预览缩放的上下限（百分比）：与表格 / 文档两个预览共用，避免两边手感不同 */
export const PREVIEW_ZOOM_MIN = 10;
export const PREVIEW_ZOOM_MAX = 400;

/** 预览外壳的构造参数 */
export interface PrintPreviewShellOptions {
    /** 标题栏文字 */
    title?: string;
    /** 内容样式（表格 = table 样式；文档 = 正文样式），注入外壳的 <style> 末尾 */
    extraCss?: string;
    /** 自定义字体的 @font-face 规则（预览与打印共用） */
    fontFaceCss?: string;
}

/**
 * 生成打印预览外壳的 HTML
 *
 * 结构与类名与表格预览**完全一致**（`.gs-print-overlay` 及其后代），
 * 因此调用方可以照抄表格那套 `data-gs-print-role` 绑定逻辑。
 * @param options - 见 {@link PrintPreviewShellOptions}
 * @returns 可直接赋给 `innerHTML` 的 HTML 文本
 */
export function buildPrintPreviewShellHtml(options: PrintPreviewShellOptions = {}): string {
    const title = options.title ?? '打印预览';
    const extraCss = options.extraCss ?? '';
    const fontFaceCss = options.fontFaceCss ?? '';
    return `
            <style>
                /* 样式遵循 Bootstrap 5 设计规范：字体堆栈、颜色变量、圆角、阴影、表单焦点环、按钮体系 */
                .gs-print-overlay { position: fixed; inset: 0; z-index: 9999; background: rgba(0, 0, 0, 0.5); display: flex; flex-direction: column; font-family: system-ui, -apple-system, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, 'Noto Sans', 'Liberation Sans', sans-serif; font-size: 1rem; color: #212529; }
                .gs-print-toolbar { flex: none; display: flex; align-items: center; padding: 0.75rem 1rem; background: #f8f9fa; border-bottom: 1px solid #dee2e6; }
                .gs-print-title { font-size: 1rem; font-weight: bold; color: #212529; }
                .gs-print-body { flex: 1; display: flex; overflow: hidden; min-height: 0; }
                /* 主体列：预览滚动区 + 页脚（页脚仅覆盖预览区宽度）；右侧设置面板独占全高直达页面底部 */
                .gs-print-main { flex: 1; display: flex; flex-direction: column; min-height: 0; min-width: 0; }
                .gs-print-scroll { flex: 1; overflow: auto; padding: 1.5rem; background: #6c757d; }
                /* 首屏加载占位：分页计划生成期间显示（避免白屏与“假死”观感） */
                .gs-print-loading { padding: 3rem 1rem; text-align: center; color: #fff; font-size: 1rem; }
                /* 预览页面：与实际纸张等比例（宽高与页边距 padding 由 _renderCurrentPage 按当前设置内联指定），
                   整页以 zoom 等比缩放——高度填满滚动区可用高度，宽度与内容随同一系数缩放（见 _applyPreviewZoom）；
                   单页显示模式，内容不足一页时白色背景仍是完整纸张 */
                .gs-print-page { box-sizing: border-box; margin: 0 auto; background: #fff; overflow: hidden; border-radius: 0.375rem; box-shadow: 0 0.5rem 1rem rgba(0, 0, 0, 0.15); }
                .gs-print-panel { flex: none; width: 264px; border-left: 1px solid #dee2e6; background: #f8f9fa; padding: 1rem; display: flex; flex-direction: column; gap: 0.75rem; overflow: auto; }
                .gs-print-panel-title { font-size: 1rem; font-weight: bold; color: #212529; }
                /* 设置行：文字提示与操作控件同一水平位置（label 固定宽左对齐，控件右侧伸展） */
                .gs-print-field { display: flex; align-items: center; gap: 0.5rem; font-size: 0.875rem; color: #212529; }
                .gs-print-label { flex: none; width: 4.25em; }
                .gs-print-field > .gs-print-select,
                .gs-print-field > .gs-print-input,
                .gs-print-field > .gs-print-stepper { flex: 1; min-width: 0; }
                .gs-print-field > .gs-print-input { width: auto; }
                /* 分割线：设置分组之间 */
                .gs-print-divider { flex: none; height: 1px; background: #dee2e6; }
                /* 表单控件：form-select / form-control 风格（含 focus 蓝色焦点环） */
                .gs-print-select { padding: 0.375rem 0.75rem; font-size: 0.875rem; border: 1px solid #ced4da; border-radius: 0.375rem; background: #fff; color: #212529; cursor: pointer; outline: none; transition: border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out; }
                .gs-print-select:focus { outline: none; border-color: #86b7fe; box-shadow: 0 0 0 0.25rem rgba(13, 110, 253, 0.25); }
                .gs-print-input { padding: 0.375rem 0.5rem; font-size: 0.875rem; border: 1px solid #ced4da; border-radius: 0.375rem; background: #fff; color: #212529; width: 100%; box-sizing: border-box; outline: none; transition: border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out; }
                .gs-print-input:focus { outline: none; border-color: #86b7fe; box-shadow: 0 0 0 0.25rem rgba(13, 110, 253, 0.25); }
                /* 数字输入框统一隐藏原生步进箭头（步进由自定义按钮/滑块承担） */
                .gs-print-input::-webkit-outer-spin-button,
                .gs-print-input::-webkit-inner-spin-button { -webkit-appearance: none; margin: 0; }
                .gs-print-input[type="number"] { -moz-appearance: textfield; appearance: textfield; }
                .gs-print-stepper { display: flex; align-items: stretch; }
                .gs-print-stepper .gs-print-input { flex: 1; min-width: 0; border-radius: 0.375rem 0 0 0.375rem; -moz-appearance: textfield; appearance: textfield; }
                .gs-print-stepper .gs-print-input::-webkit-outer-spin-button,
                .gs-print-stepper .gs-print-input::-webkit-inner-spin-button { -webkit-appearance: none; margin: 0; }
                .gs-print-stepper .gs-print-input:focus { position: relative; z-index: 1; }
                .gs-print-stepper-btns { flex: none; display: flex; flex-direction: column; border: 1px solid #ced4da; border-left: none; border-radius: 0 0.375rem 0.375rem 0; overflow: hidden; background: #fff; }
                .gs-print-step-btn { flex: 1; width: 22px; padding: 0; font-size: 0.75rem; line-height: 1; border: none; background: #fff; color: #212529; cursor: pointer; transition: background 0.15s ease-in-out; }
                .gs-print-step-btn:hover { background: #e9ecef; }
                .gs-print-step-btn:active { background: #dee2e6; }
                /* 小按钮：行内操作（调整页边距/应用/取消） */
                .gs-print-btn-sm { flex: none; padding: 0.25rem 0.5rem; font-size: 0.8125rem; line-height: 1.25; border: 1px solid #dee2e6; border-radius: 0.375rem; background: #fff; color: #212529; cursor: pointer; white-space: nowrap; transition: color 0.15s ease-in-out, background 0.15s ease-in-out, border-color 0.15s ease-in-out; }
                .gs-print-btn-sm:hover { background: #f8f9fa; }
                .gs-print-btn-sm.gs-print-btn-primary { background: #0d6efd; border-color: #0d6efd; color: #fff; }
                .gs-print-btn-sm.gs-print-btn-primary:hover { background: #0b5ed7; border-color: #0a58ca; color: #fff; }
                /* 自定义页边距编辑器：展开于「页边距」行下方，上下左右四向 mm 输入 */
                .gs-print-margin-editor { display: flex; flex-direction: column; gap: 0.5rem; padding: 0.625rem; background: #fff; border: 1px solid #dee2e6; border-radius: 0.375rem; }
                .gs-print-margin-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 0.5rem; }
                .gs-print-mini-field { display: flex; align-items: center; gap: 0.375rem; font-size: 0.8125rem; color: #212529; }
                .gs-print-mini-field > .gs-print-input { flex: 1; min-width: 0; width: auto; padding: 0.25rem 0.375rem; font-size: 0.8125rem; }
                .gs-print-mini-actions { display: flex; gap: 0.5rem; }
                .gs-print-mini-actions .gs-print-btn-sm { flex: 1; white-space: normal; }
                /* 按钮行：吸附面板底部，但与面板底边留出大于页脚条高度（≈51px）的间距 */
                .gs-print-actions { display: flex; gap: 0.5rem; margin-top: auto; margin-bottom: 3.5rem; }
                /* 按钮：默认 .btn 风格（白底 + #dee2e6 边 + hover #f8f9fa） */
                .gs-print-btn { flex: 1; padding: 0.375rem 0; font-size: 0.875rem; border: 1px solid #dee2e6; border-radius: 0.375rem; background: #fff; color: #212529; cursor: pointer; transition: color 0.15s ease-in-out, background 0.15s ease-in-out, border-color 0.15s ease-in-out; }
                .gs-print-btn:hover { background: #f8f9fa; }
                /* 打印按钮 = .btn-primary（含 hover/active 层级色） */
                .gs-print-btn[data-gs-print-action="print"] { background: #0d6efd; border-color: #0d6efd; color: #fff; }
                .gs-print-btn[data-gs-print-action="print"]:hover { background: #0b5ed7; border-color: #0a58ca; color: #fff; }
                .gs-print-btn[data-gs-print-action="print"]:active { background: #0a58ca; border-color: #0a58ca; color: #fff; }
                .gs-print-footer { flex: none; display: flex; align-items: center; justify-content: center; gap: 0.75rem; padding: 0.75rem 1rem; background: #f8f9fa; border-top: 1px solid #dee2e6; position: relative; }
                .gs-print-indicator { font-size: 0.875rem; color: #212529; min-width: 110px; text-align: center; }
                /* 预览缩放控件容器：固定于页脚右侧（页码居中不受影响）；内部滑块由 common/slider 提供样式 */
                .gs-print-footer-zoom { position: absolute; right: 1rem; top: 50%; transform: translateY(-50%); display: flex; align-items: center; gap: 0.5rem; }
                /* 翻页按钮 = .btn-outline-secondary 风格 */
                .gs-print-nav-btn { padding: 0.25rem 0.75rem; font-size: 0.875rem; border: 1px solid #6c757d; border-radius: 0.375rem; background: #fff; color: #6c757d; cursor: pointer; transition: color 0.15s ease-in-out, background 0.15s ease-in-out, border-color 0.15s ease-in-out; }
                .gs-print-nav-btn:hover:not(:disabled) { background: #6c757d; border-color: #6c757d; color: #fff; }
                .gs-print-nav-btn:disabled { color: #adb5bd; border-color: #dee2e6; background: #fff; cursor: not-allowed; }
                /* 预览表格基础样式（与打印文档共用） */
                ${extraCss}
                /* 自定义字体：注入收集与注册的 @font-face 规则，保证预览与打印字体一致 */
                ${fontFaceCss}
            </style>
            <div class="gs-print-toolbar">
                <span class="gs-print-title">${title}</span>
            </div>
            <div class="gs-print-body">
                <div class="gs-print-main">
                    <div class="gs-print-scroll">
                        <div class="gs-print-pages"></div>
                    </div>
                    <div class="gs-print-footer">
                        <button class="gs-print-nav-btn" data-gs-print-nav="prev">&lt;</button>
                        <span class="gs-print-indicator"></span>
                        <button class="gs-print-nav-btn" data-gs-print-nav="next">&gt;</button>
                        <div class="gs-print-footer-zoom" data-gs-print-role="zoom"></div>
                    </div>
                </div>
                <div class="gs-print-panel">
                    <div class="gs-print-panel-title">打印设置</div>
                    <label class="gs-print-field"><span class="gs-print-label">打印机</span>
                        <select class="gs-print-select" data-gs-print-role="printer"></select>
                    </label>
                    <label class="gs-print-field"><span class="gs-print-label">份数</span>
                        <span class="gs-print-stepper">
                            <input class="gs-print-input" type="number" min="1" max="999" step="1" value="1" data-gs-print-role="copies">
                            <span class="gs-print-stepper-btns">
                                <button type="button" class="gs-print-step-btn" data-gs-print-step="inc">+</button>
                                <button type="button" class="gs-print-step-btn" data-gs-print-step="dec">&minus;</button>
                            </span>
                        </span>
                    </label>
                    <label class="gs-print-field"><span class="gs-print-label">顺序</span>
                        <select class="gs-print-select" data-gs-print-role="collate">
                            <option value="page">逐页打印</option>
                            <option value="copies">逐份打印</option>
                        </select>
                    </label>
                    <label class="gs-print-field"><span class="gs-print-label">打印方式</span>
                        <select class="gs-print-select" data-gs-print-role="duplex">
                            <option value="simplex">单面打印</option>
                            <option value="duplex-long">双面打印-长边翻页</option>
                            <option value="duplex-short">双面打印-短边翻页</option>
                        </select>
                    </label>
                    <div class="gs-print-divider"></div>
                    <label class="gs-print-field"><span class="gs-print-label">纸张大小</span>
                        <select class="gs-print-select" data-gs-print-role="paper-size">
                            <option value="A4">A4</option>
                            <option value="A3">A3</option>
                            <option value="B5">B5</option>
                            <option value="Letter">Letter</option>
                        </select>
                    </label>
                    <label class="gs-print-field"><span class="gs-print-label">纸张方向</span>
                        <select class="gs-print-select" data-gs-print-role="orientation">
                            <option value="portrait">纵向</option>
                            <option value="landscape">横向</option>
                        </select>
                    </label>
                    <div class="gs-print-field"><span class="gs-print-label">页边距</span>
                        <select class="gs-print-select" data-gs-print-role="margin">
                            <option value="normal">常规</option>
                            <option value="wide">宽</option>
                            <option value="narrow">窄</option>
                            <option value="custom">自定义</option>
                        </select>
                        <button type="button" class="gs-print-btn-sm" data-gs-print-action="margin-adjust">调整页边距</button>
                    </div>
                    <div class="gs-print-margin-editor" data-gs-print-role="margin-editor" style="display:none">
                        <div class="gs-print-margin-grid">
                            <label class="gs-print-mini-field"><span>上</span><input class="gs-print-input" type="number" min="0" max="100" step="0.5" data-gs-print-role="margin-top"></label>
                            <label class="gs-print-mini-field"><span>下</span><input class="gs-print-input" type="number" min="0" max="100" step="0.5" data-gs-print-role="margin-bottom"></label>
                            <label class="gs-print-mini-field"><span>左</span><input class="gs-print-input" type="number" min="0" max="100" step="0.5" data-gs-print-role="margin-left"></label>
                            <label class="gs-print-mini-field"><span>右</span><input class="gs-print-input" type="number" min="0" max="100" step="0.5" data-gs-print-role="margin-right"></label>
                        </div>
                        <div class="gs-print-mini-actions">
                            <button type="button" class="gs-print-btn-sm gs-print-btn-primary" data-gs-print-action="margin-apply">应用</button>
                            <button type="button" class="gs-print-btn-sm" data-gs-print-action="margin-cancel">取消</button>
                        </div>
                    </div>
                    <div class="gs-print-divider"></div>
                    <div class="gs-print-field"><span class="gs-print-label">缩放</span>
                        <select class="gs-print-select" data-gs-print-role="scale">
                            <option value="none">无缩放</option>
                            <option value="fit-sheet">整个工作表</option>
                            <option value="fit-cols">所有列一页</option>
                            <option value="fit-rows">所有行一页</option>
                            <option value="custom">自定义</option>
                        </select>
                        <input class="gs-print-input" type="number" min="10" max="400" step="5" data-gs-print-role="scale-input" disabled>
                    </div>
                    <label class="gs-print-field"><span class="gs-print-label">页面范围</span>
                        <select class="gs-print-select" data-gs-print-role="page-range">
                            <option value="all">全部</option>
                            <option value="custom">页码范围</option>
                        </select>
                    </label>
                    <div class="gs-print-field" data-gs-print-role="page-range-row" style="display:none">
                        <span class="gs-print-label"></span>
                        <input class="gs-print-input" type="text" placeholder="如：1-3,5" data-gs-print-role="page-range-input">
                    </div>
                    <div class="gs-print-actions">
                        <button class="gs-print-btn" data-gs-print-action="print">打印</button>
                        <button class="gs-print-btn" data-gs-print-action="close">关闭</button>
                    </div>
                </div>
                </div>
            </div>
        `;
}
