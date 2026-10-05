/**
 * 表格的文字输入区域组件
 * 提供单元格文本编辑、光标控制、文本选择等功能
 *
 * @module EditCanvas
 * @author 东方鹗
 * @see
 * # B站: https://space.bilibili.com/194359739
 * # 知乎: https://www.zhihu.com/people/eastossifrage
 * # CSDN: https://blog.csdn.net/os373
 */

import { createDiv } from "../../utils/dom";
import { CELL_PADDING, DEFAULT_CELL_BG_COLOR,
    DEFAULT_FONT_FAMILY, DEFAULT_FONT_SIZE,
    TEXT_COLOR, EDIT_CANVAS, SHEETS_NAME,
    DataEvents} from "../constant";
import { Char } from "../dataArchitecture/char";
import { Cell } from "../dataArchitecture/cell";
import { Canvas, CanvasOptions } from "./canvas";
import { CursorPosition, LineText } from "../dataArchitecture/cellContent";



/**
 * EditCanvas 类
 * 管理电子表格的单元格文本编辑功能
 * 提供文本输入、光标控制、文本选择、撤销/重做等编辑功能
 *
 * @class EditCanvas
 * @extends Canvas
 * @example
 * ```typescript
 * const editCanvas = new EditCanvas({
 *     parentElement: container,
 *     data: dataCollection
 * });
 * editCanvas.startEditting();
 * ```
 */
export class EditCanvas extends Canvas {
    /** 光标元素 */
    private cursor: HTMLDivElement;
    /** 隐藏的文本输入框 */
    private textArea: HTMLTextAreaElement;
    /** 是否正在输入法组合输入 */
    public isComposing: boolean = false;
    /** 光标闪烁间隔 ID */
    private cursorBlinkInterval: number | null = null;
    /** 是否正在拖动选择 */
    private isDragging: boolean = false;
    /** 拖动开始时的光标位置 */
    private dragStartPosition: CursorPosition | null = null;
    /** 最后一次输入开始时的光标位置 */
    private lastInputStartPos: number = 0;
    /** 编辑开始时的单元格值（用于撤销记录） */
    private initialCellValue: Char[] = [];
    /** 新输入的字符（用于单元格新增字符） */
    private newChars: string[] = [];

    /**
     * 创建 EditCanvas 实例
     * @param {CanvasOptions} options - 画布配置选项
     */
    constructor(options: CanvasOptions) {
        super(options);
        this.parentElement = options.parentElement;
        this.data = options.data;
        this.init();
    }

    /**
     * 初始化编辑画布
     * 创建 canvas 元素、光标元素和隐藏的文本输入框
     * @private
     */
    private init() {
        this.canvas.setAttribute('tabindex', '0');// 确保 canvas 可聚焦
        this.canvas.id = EDIT_CANVAS;
        this.cursor = createDiv({
            id: "editCursor",
            style: {
                position: "absolute",
                zIndex: "3",
                left: "0px",
                top: "0px",
                display: "none"
            }
        });
        this.textArea = document.createElement("textarea");
        this.textArea.id = "hiddenTextarea";
        this.textArea.tabIndex = -1;
        this.textArea.autofocus = true;
        this.textArea.style.position = "absolute";
        if (this.parentElement) {
            this.parentElement.appendChild(this.canvas);
            this.parentElement.appendChild(this.cursor);
            this.parentElement.appendChild(this.textArea);
        }
        // 刷新页面默认由 textArea 持有焦点：autofocus 属性对 JS 动态插入的元素无效，
        // 这里在挂载后显式聚焦；后续焦点增减统一由 handleDocumentMouseDownForFocus 管理
        this.textArea.focus();
        this.setupContext();
        this.setupDataEventListeners();
        this.setupEventListeners();
        // 注入光标坐标解析器：统一走 calcCursorX + line.editY，
        // 使 data.cursorAbsolutePosition 写入时能得到与文本渲染一致的 editX/editY。
        this.data.cursorCoordResolver = (lineIndex: number, charIndex: number) => ({
            editX: this.calcCursorX(lineIndex, charIndex),
            editY: this.data.lines[lineIndex]?.editY ?? 0
        });
    }

    /**
     * 设置数据变化事件监听器
     * @private
     */
    private setupDataEventListeners(): void {        
        this.data.on(DataEvents.ACTIVATED_CELL_CHANGED, (_, cell) => {
            this.updateCanvasSize();
        });

        this.data.on(DataEvents.SELECTION_CHANGED, (_, selection) => {
            this.updateCanvasSize();
        });
        this.data.on(DataEvents.IS_EDITTING_CHANGED, (_, isEditting) => {
            if(isEditting) {
                // 进入编辑态：开启编辑内撤销栈，捕获编辑前基准快照
                this.data.beginEditSession();
                this.startEditting();
            } else {
                this.finishEditting();
                // 退出编辑态：把整段编辑折叠为 undoManager 的一条文档级撤销操作
                this.data.endEditSession();
            }
        });
        this.data.on(DataEvents.CURSOR_POSITION_CHANGED, (_, cursorPosition) => {
            this.syncToTextAreaCursor();
            this.updateCursor();
        });
    }

    /**
     * 设置 DOM 事件监听器
     * 包括输入事件、组合输入事件、鼠标事件和键盘事件
     * @private
     */
    private setupEventListeners(): void {
        this.textArea.addEventListener('beforeinput', (e) => {
            // 在输入发生之前保存光标位置
            this.lastInputStartPos = this.textArea.selectionStart;
        });
        this.textArea.addEventListener('input', this.handleValueChange.bind(this));
        this.textArea.addEventListener('compositionstart', this.handleCompostionStart.bind(this));
        this.textArea.addEventListener('compositionupdate', this.handleCompostionUpdate.bind(this));
        this.textArea.addEventListener('compositionend', this.handleCompostionEnd.bind(this));
        // 添加鼠标事件
        this.canvas.addEventListener('mousedown', this.handleMouseDown.bind(this));
        document.addEventListener('mousemove', this.handleMouseMove.bind(this));
        document.addEventListener('mouseup', this.handleMouseUp.bind(this));
        this.canvas.addEventListener('dblclick', this.handleDoubleClick.bind(this));
        // 焦点管理：仅点击主页面时 textArea 获取焦点，点击其它位置时失去焦点
        document.addEventListener('mousedown', this.handleDocumentMouseDownForFocus.bind(this));

        // 添加键盘事件
        document.addEventListener('keydown', (e) => this.handleKeyDown(e));
    }

    /**
     * 文档级 mousedown 焦点管理：
     * - 点击主页面（sheetCanvas 主网格，或编辑态覆盖在活动单元格上的编辑覆盖层）→ textArea 获取焦点；
     * - 点击其它任何位置（工具栏/行列表头/滚动条/打印预览等）→ textArea 失去焦点；
     * - 刷新页面的默认焦点由 init() 中显式聚焦保证。
     * @param {MouseEvent} e - 鼠标按下事件
     * @private
     */
    private handleDocumentMouseDownForFocus(e: MouseEvent): void {
        const target = e.target as HTMLElement | null;
        if (target && (target.id === SHEETS_NAME.SHEET || target.id === EDIT_CANVAS)) {
            this.textArea.focus();
        } else {
            this.textArea.blur();
        }
    }

    /**
     * 计算单元格文本尺寸，并设置编辑框位置
     * @public
     */
    public updateCanvasSize() {
        if(!this.data.isEditting) this.cursor.style.display = 'none';
        // 安全检查：如果 rowHeaders 或 colHeaders 未初始化，直接返回
        if (!this.data.rowHeaders.length || !this.data.colHeaders.length) {
            return;
        }
        
        const { editLeft, editTop, contentWidth, contentHeight, lines } = this.measureCellText(this.data.activedCell);
        this.data.lines = lines;
        // 计算输入框位置
        this.textArea.value = this.data.getCellValue(this.data.activedCell);
        this.textArea.style.left = `${Math.round(editLeft + this.data.visibleView.rowHeaderWidth)}px`;
        this.textArea.style.top = `${Math.round(editTop + this.data.visibleView.colHeaderHeight)}px`;
        this.textArea.style.height = `${Math.round(contentHeight - CELL_PADDING * 2)}px`;

        this.canvas.style.left = `${Math.round(editLeft + this.data.visibleView.rowHeaderWidth)}px`;
        this.canvas.style.top = `${Math.round(editTop + this.data.visibleView.colHeaderHeight)}px`;
        this.canvas.width = Math.round((contentWidth) * window.devicePixelRatio);
        this.canvas.height = Math.round((contentHeight) * window.devicePixelRatio);
        this.canvas.style.width = `${Math.round(contentWidth)}px`;
        this.canvas.style.height = `${Math.round(contentHeight)}px`;
        this.ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);
    }

    /**
     * 绘制单元格文本
     * 在编辑画布上绘制单元格的文本内容，包括字符样式（字体、颜色、下划线、删除线等）
     * @param {string} cell - 单元格名称
     * @public
     */
    public drawEditCellsText(cell: string) {
        const i = this.data.values.findIndex((item) => item.cell === cell); // 单元格索引
        if (i === -1) return;

        const { col, row } = this.data.getCellColAndRow(cell);
        // 获取单元格尺寸
        const { left: cellLeft, top: cellTop, width: cellWidth, height: cellHeight } = this.data.calculateCellRect(col, row);
        this.ctx.clearRect(
            Math.round((cellLeft - this.data.colHeaders.offsetWidth) * this.data.zoom ) * window.devicePixelRatio + 1, 
            Math.round((cellTop - this.data.rowHeaders.offsetHeight) * this.data.zoom) * window.devicePixelRatio + 1, 
            Math.round(cellWidth * this.data.zoom * window.devicePixelRatio) - 1, 
            Math.round(cellHeight * this.data.zoom * window.devicePixelRatio) - 1); // 清空主canvas指定区域

        const { contentWidth, contentHeight, lines } = this.measureCellText(cell); 
        
        const backgroundColor = this.data.values[i]?.backgroundColor || 'whitesmoke';
        this.ctx.save();
        this.ctx.fillStyle = backgroundColor;
        this.fillRect(0, 0, contentWidth, contentHeight);
        this.ctx.restore();
        // 复制主canvas的绘图上下文设置
        const letterSpacing = this.data.values[i]?.letterSpacing || 0;
        const lineSpacing = this.data.values[i]?.lineSpacing || 0;
        this.ctx.textAlign = 'left';
        this.ctx.textBaseline = 'bottom';
        // 在editCanvas上绘制文本（使用editX/editY确保与非编辑模式对齐）
        let k = 0; // 特殊字符索引
        for (let line of lines) {
            let currentX = line.editX;
            let currentY = line.editY + line.lineHeight - lineSpacing;
            for (let j=0; j < line.text.length; j++) {
                const char = line.text[j];
                const { charWidth, fontColor, fontWeight, fontStyle, fontSize, fontFamily, underline, strikethrough } = this.calculateCharGeometry(cell, k);
                
                this.ctx.save();
                this.ctx.fillStyle = `${fontColor}`;
                const fontParts = [];
                if (fontStyle) fontParts.push(fontStyle);
                if (fontWeight) fontParts.push(fontWeight);
                fontParts.push(`${Math.round(fontSize * window.devicePixelRatio)}px`);
                fontParts.push(fontFamily);
                this.ctx.font = fontParts.join(' ');
                this.fillText(char, currentX, currentY);
                
                // 绘制下划线
                if (underline) {
                    this.ctx.strokeStyle = fontColor;
                    this.ctx.lineWidth = Math.max(1, Math.round(fontSize / 10 * window.devicePixelRatio));
                    const underlineY = currentY - 1;
                    this.ctx.beginPath();
                    this.moveTo(currentX, underlineY);
                    this.lineTo(currentX + charWidth, underlineY);
                    this.ctx.stroke();
                }
                
                // 绘制删除线
                if (strikethrough) {
                    this.ctx.strokeStyle = fontColor;
                    this.ctx.lineWidth = Math.max(1, Math.round(fontSize / 10 * window.devicePixelRatio));
                    const strikethroughY = currentY - fontSize * 0.48;
                    this.ctx.beginPath();
                    this.moveTo(currentX, strikethroughY);
                    this.lineTo(currentX + charWidth, strikethroughY);
                    this.ctx.stroke();
                }
                
                this.ctx.restore();
                k += 1;
                // letterSpacing 仅加在"当前字符是可见字符，且下一字符也是可见字符"之间
                // 与 calculateSingleLineGeometry 的累加规则保持一致，避免测量宽度与绘制位移不一致
                const hasNextVisible = j + 1 < line.text.length && line.text[j + 1] !== '\n';
                currentX += charWidth + (char !== '\n' && hasNextVisible ? letterSpacing : 0);
            }
        }
    }
    
    /**
     * 处理组合输入开始事件
     * @param {CompositionEvent} e - 组合输入事件对象
     * @public
     */
    public handleCompostionStart(e: CompositionEvent) {
        // 在输入法组合开始前保存光标位置
        this.lastInputStartPos = this.textArea.selectionStart;
        this.isComposing = true;
        if (!this.data.isEditting) {
            this.startEditting(true);
        }
    }

    /**
     * 处理组合输入更新事件
     * @param {CompositionEvent} e - 组合输入事件对象
     * @public
     */
    public handleCompostionUpdate(e: CompositionEvent) {
    }

    /**
     * 处理组合输入结束事件
     * @param {CompositionEvent} e - 组合输入事件对象
     * @public
     */
    public handleCompostionEnd(e: CompositionEvent) {
        this.isComposing = false;
        // 用于保存新输入的中文字符
        this.newChars.push(e.data);
        this.handleValueChange();
    }

    /**
     * 处理鼠标按下事件
     * @param {MouseEvent} e - 鼠标事件对象
     * @public
     */
    public handleMouseDown(e: MouseEvent) {
        if (e.button !== 0) return;
        e.preventDefault();
        e.stopPropagation();
        if (e.target === this.canvas)  {  // 点击editCanvas, 为拖动做准备或选择光标位置
            this.isDragging = true;        
            this.data.selectionOnEditting = null; // 重置选择范围，确保使用新的光标位置
            this.setCursorPosition(e.clientX, e.clientY);    
            this.cursor.style.display = 'block';
            this.dragStartPosition = this.data.cursorPosition;
            this.draw();
        }
    }

    /**
     * 处理鼠标移动事件
     * @param {MouseEvent} e - 鼠标事件对象
     * @public
     */
    public handleMouseMove(e: MouseEvent) {
        if (!this.isDragging) return;
        this.setCursorPosition(e.clientX, e.clientY);
        this.updateSelection();
        this.syncToTextAreaCursor();
        this.draw();
    } 

    /**
     * 处理鼠标释放事件
     * @param {MouseEvent} e - 鼠标事件对象
     * @public
     */
    public handleMouseUp(e: MouseEvent) {
        if (e.button !== 0) return;
        this.isDragging = false;
        this.dragStartPosition = null;
        // this.data.selectionOnEditting = null;
    }

    /**
     * 处理鼠标双击事件，在编辑模式下全选文本
     * @param {MouseEvent} e - 鼠标事件对象
     * @public
     */
    public handleDoubleClick(e: MouseEvent) {
        if (e.button !== 0) return;
        e.preventDefault();
        e.stopPropagation();
        
        // 如果双击的是画布，全选所有文本
        if (e.target === this.canvas && this.data.lines.length > 0) {
            const start = {
                lineIndex: 0,
                charIndex: 0,
                editX: this.data.lines[0].editX,
                editY: this.data.lines[0].editY
            };
            const end = {
                lineIndex: this.data.lines.length - 1, 
                charIndex: this.data.lines[this.data.lines.length - 1].text.length,
                editX: this.data.lines[this.data.lines.length - 1].editX + this.data.lines[this.data.lines.length - 1].lineWidth,
                editY: this.data.lines[this.data.lines.length - 1].editY + this.data.lines[this.data.lines.length - 1].lineHeight
            };
            this.data.selectionOnEditting = { start, end };
            this.draw();
            // 同步到textArea的光标位置
            this.syncToTextAreaCursor();
        }
    }

    /**
     * 计算指定行和字符索引处的光标 X 坐标
     * @param {number} lineIndex - 行索引
     * @param {number} charIndex - 行内字符索引
     * @returns {number} 光标 X 坐标
     * @private
     */
    private calcCursorX(lineIndex: number, charIndex: number): number {
        const line = this.data.lines[lineIndex];
        if (!line) return 0;
        let x = line.editX;
        const baseIndex = this.data.getPreLineCharsLength(lineIndex);
        const cellObj = this.data.values.find(v => v.cell === this.data.activedCell);
        const letterSpacing = cellObj?.letterSpacing || 0;
        for (let i = 0; i < charIndex; i++) {
            const { charWidth } = this.calculateCharGeometry(this.data.activedCell, baseIndex + i);
            // letterSpacing 仅加在"当前字符是可见字符，且下一字符也是可见字符"之间
            // 与 calculateSingleLineGeometry / drawEditCellsText 的累加规则保持一致
            const hasNextVisible = i + 1 < line.text.length && line.text[i + 1] !== '\n';
            x += charWidth + (line.text[i] !== '\n' && hasNextVisible ? letterSpacing : 0);
        }
        return x;
    }

    /**
     * 光标移动后的统一处理
     * 更新光标显示、同步 textarea 光标位置、保存输入起始位置
     * @private
     */
    private afterCursorMove(): void {
        this.updateCursor();
        this.syncToTextAreaCursor();
        this.lastInputStartPos = this.textArea.selectionStart;
    }

    /**
     * 计算某一行的"有效行尾"字符索引（光标可停留的最大 charIndex）。
     * 若该行以换行符结尾（段落行），行尾停留点在换行符之前（text.length - 1），
     * 这样光标左右位移时会忽略换行符；否则停留点为 text.length。
     * @param {number} lineIndex - 行索引
     * @returns {number} 有效行尾的 charIndex
     * @private
     */
    private getLineEffectiveEnd(lineIndex: number): number {
        const line = this.data.lines[lineIndex];
        if (!line) return 0;
        const len = line.text.length;
        return len > 0 && line.text[len - 1] === '\n' ? len - 1 : len;
    }

    /**
     * 刷新行文本数据
     * 重新测量当前活动单元格的文本，更新 data.lines 缓存
     * @returns {LineText[]} 最新的行文本数据数组
     * @private
     */
    private refreshLines(): LineText[] {
        const { lines } = this.measureCellText(this.data.activedCell);
        this.data.lines = lines;
        return lines;
    }

    /**
     * 更新选择范围
     * 根据拖动起始位置和当前光标位置计算选择范围，确保 start 始终在 end 之前
     * @private
     */
    private updateSelection() {
        if (this.dragStartPosition) {
            // 确保start总是在end的左边和上边
            let start = { ...this.dragStartPosition };
            let end = { ...this.data.cursorPosition };
            if (start.lineIndex > end.lineIndex) {
                // 上下颠倒，交换位置
                [start, end] = [end, start];
            } else if (start.lineIndex === end.lineIndex && start.charIndex > end.charIndex) {
                // 左右颠倒，交换位置
                [start, end] = [end, start];
            }
            
            this.data.selectionOnEditting = {
                start,
                end
            };
        }
    }
    
    /**
     * 绘制选择范围
     * 在编辑画布上绘制文本选择区域的高亮背景
     * @private
     */
    private drawSelection() {
        if (!this.data.selectionOnEditting || (this.data.selectionOnEditting.start.lineIndex === this.data.selectionOnEditting.end.lineIndex
            && this.data.selectionOnEditting.start.charIndex === this.data.selectionOnEditting.end.charIndex)) {
            return;
        }
        this.cursor.style.display = 'none';
        const { start, end } = this.data.selectionOnEditting;
        start.editX = this.calcCursorX(start.lineIndex, start.charIndex);
        start.editY = start.editY;
        end.editX = this.calcCursorX(end.lineIndex, end.charIndex);
        end.editY = end.editY;
        // 绘制选择背景
        this.ctx.save();
        this.ctx.fillStyle = 'rgba(0, 120, 215, 0.3)';
        
        if (start.lineIndex === end.lineIndex) {
            // 同一行内的选择
            this.fillRect(
                start.editX,
                (start.editY - CELL_PADDING / 2),
                (end.editX - start.editX + CELL_PADDING / 2),
                this.data.lines[start.lineIndex].lineHeight + CELL_PADDING / 2
            );
        } else {
            // 跨行选择
            // 第一行
            this.fillRect(
                start.editX,
                (start.editY - CELL_PADDING / 2),
                (this.data.lines[start.lineIndex].lineWidth + this.data.lines[start.lineIndex].editX - start.editX + CELL_PADDING / 2),
                this.data.lines[start.lineIndex].lineHeight + CELL_PADDING / 2
            );
            
            // 中间行
            for (let i = start.lineIndex + 1; i < end.lineIndex; i++) {
                this.fillRect(
                    this.data.lines[i].editX,
                    (this.data.lines[i].editY - CELL_PADDING / 2),
                    (this.data.lines[i].lineWidth + CELL_PADDING / 2),
                    this.data.lines[i].lineHeight + CELL_PADDING / 2
                );
            }
            
            // 最后一行
            this.fillRect(
                this.data.lines[end.lineIndex].editX,
                (this.data.lines[end.lineIndex].editY - CELL_PADDING / 2),
                (end.editX - this.data.lines[end.lineIndex].editX + CELL_PADDING / 2),
                this.data.lines[end.lineIndex].lineHeight + CELL_PADDING / 2
            );
        }
        
        this.ctx.restore();
    }

    /**
     * 设置光标位置
     * 根据鼠标点击位置计算并设置光标在文本中的位置
     * @param {number} clientX - 鼠标 X 坐标
     * @param {number} clientY - 鼠标 Y 坐标
     * @public
     */
    public setCursorPosition(clientX: number, clientY: number) {
        const rect = this.canvas.getBoundingClientRect();
        const x = clientX - rect.left;
        const y = clientY - rect.top;
        // 计算光标应该在的位置
        let lineIndex = 0;
        let charIndex = 0;
        let cursorX = 0;
        let cursorY = 0;
        // 设置光标位置时，需要重新测量文本行
        this.data.lines = this.refreshLines();
        if (this.data.lines.length > 0) {
            // 如果字符为空
            if (this.data.lines[0].text.length === 0) {
                const { activedWidth, activedHeight } = this.data.getActivedRect();
                lineIndex = 0;
                charIndex = 0;
                cursorX = CELL_PADDING
                        + activedWidth * this.data.zoom / 2;;
                cursorY = CELL_PADDING
                        + (activedHeight- DEFAULT_FONT_SIZE) * this.data.zoom  / 2; 
            } else {
                // 找到对应的行
                for (let i = 0; i < this.data.lines.length; i++) {
                    const line = this.data.lines[i];
                    if (y <= line.editY + line.lineHeight) {
                        lineIndex = i;
                        break;
                    }
                }
                
                let currentLine = this.data.lines[lineIndex];
                // 如果点击位置在最后一行之后
                if (y > this.data.lines[this.data.lines.length - 1].editY + currentLine.lineHeight) {
                    lineIndex = this.data.lines.length - 1;
                    currentLine = this.data.lines[lineIndex];
                }
                
                // 计算光标Y坐标
                cursorY = currentLine.editY;
                
                // 找到对应的字符位置
                let currentX = currentLine.editX;
                // 该行起始的全局字符索引
                let k = this.data.toAbsoluteIndex({ lineIndex, charIndex: 0 });
                
                const cellObj = this.data.values.find(v => v.cell === this.data.activedCell);
                const letterSpacing = cellObj?.letterSpacing || 0;
                for (let i = 0; i < currentLine.text.length; i++) {
                    const ch = currentLine.text[i];
                    const { charWidth } = this.calculateCharGeometry(this.data.activedCell, k);
                    if (x <= currentX + charWidth / 2) {
                        charIndex = i;
                        break;
                    }
                    // letterSpacing 仅加在"当前字符是可见字符，且下一字符也是可见字符"之间
                    // 与 calculateSingleLineGeometry / drawEditCellsText 的累加规则保持一致
                    const hasNextVisible = i + 1 < currentLine.text.length && currentLine.text[i + 1] !== '\n';
                    currentX += charWidth + (ch !== '\n' && hasNextVisible ? letterSpacing : 0);
                    k++;
                }
                
                // 计算光标X坐标
                cursorX = currentX;
                // 如果点击位置在最后一个字符之后
                if (currentLine.text.length > 0) {
                    const lastCharIndex = this.data.toAbsoluteIndex({ lineIndex, charIndex: currentLine.text.length - 1 });
                    const { charWidth } = this.calculateCharGeometry(this.data.activedCell, lastCharIndex);
                    if (x > currentLine.editX + currentLine.lineWidth - charWidth / 2) {
                        charIndex = currentLine.text.length;
                        cursorX = currentLine.editX + currentLine.lineWidth;
                    }
                } else {
                    charIndex = 0;
                    cursorX = currentLine.editX;
                }
            }            
        }
        
        // 更新光标位置
        this.data.cursorPosition = {
            lineIndex,
            charIndex,
            editX: cursorX,
            editY: cursorY
        };
        
        // 更新光标样式和位置
        this.updateCursor();
        
        // 同步到textArea的光标位置
        this.syncToTextAreaCursor();
        
        // 保存光标位置，用于后续输入时的样式继承
        this.lastInputStartPos = this.textArea.selectionStart;
    }

    /**
     * 设置光标到最后一个字符
     * 将光标移动到单元格文本的最后一个字符位置
     * @public
     */
    public setCursorToLastChar() {
        const lines = this.refreshLines();
        
        // 如果没有文本，保持默认位置
        if (!lines || lines.length === 0) {
            return;
        }
        
        // 获取最后一行
        const lineIndex = lines.length - 1;
        const lastLine = lines[lineIndex];
        const charIndex = lastLine.text.length;
        
        // 更新光标位置
        this.data.cursorPosition = {
            lineIndex,
            charIndex,
            editX: this.calcCursorX(lineIndex, charIndex),
            editY: lastLine.editY
        };
        
        // 更新光标显示
        this.updateCursor();
        this.syncToTextAreaCursor();
        this.lastInputStartPos = this.textArea.selectionStart;
    }
    
    /**
     * 更新光标样式和位置
     * 根据当前光标位置更新光标元素的位置、高度，并启动闪烁动画
     * @public
     */
    public updateCursor() {
        const { lineIndex, charIndex, editX, editY } = this.data.cursorPosition;
        const { editLeft, editTop, lines } = this.measureCellText(this.data.activedCell);
        
        const currentCharIndex = this.data.toAbsoluteIndex({ lineIndex, charIndex });
        const cellValue = this.data.values.find(
            v => v.cell === this.data.activedCell
        )?.chars;
        let cursorHeight: number;
        if (cellValue && cellValue.length > 0) {
            // 优先使用当前字符的字体大小，如果 charIndex 超出范围则使用第一个或最后一个字符
            const charInfo = cellValue[Math.min(Math.max(currentCharIndex - 1, 0), cellValue.length - 1)];
            cursorHeight = charInfo?.fontSize 
                ? Math.round(charInfo.fontSize * this.data.zoom)
                : Math.round(DEFAULT_FONT_SIZE * this.data.zoom);
        } else {
            cursorHeight = Math.round(DEFAULT_FONT_SIZE * this.data.zoom);
        }
        this.cursor.style.left = `${Math.round(editX + editLeft+ this.data.visibleView.rowHeaderWidth) - CELL_PADDING / 2}px`;
        this.cursor.style.top = `${Math.round(editY + editTop + this.data.visibleView.colHeaderHeight
            + lines[lineIndex].lineHeight
            - cursorHeight) - CELL_PADDING}px`;
        this.cursor.style.height = `${cursorHeight}px`;
        
        // 启动光标闪烁
        this.startCursorBlink();
        this.notifyToolbarStateFromCursor();
    }
    
    /**
     * 启动光标闪烁动画
     * @private
     */
    private startCursorBlink() {
        if (this.cursorBlinkInterval) {
            clearInterval(this.cursorBlinkInterval);
        }
        
        let visible = true;
        this.cursorBlinkInterval = window.setInterval(() => {
            visible = !visible;
            this.cursor.style.opacity = visible ? '1' : '0';
        }, 500);
    }
    
    /**
     * 同步文本区域的光标位置
     * 将画布上的光标位置同步到隐藏的文本输入框
     * @private
     */
    private syncToTextAreaCursor() {
        // 计算总的字符索引
        let startIndex = 0, endIndex = 0;
        if (this.data.selectionOnEditting) {  // 有选中区域
            startIndex = this.data.getSelectedCharsStartIndex();
            endIndex = this.data.getSelectedCharsEndIndex();
        } else {
            startIndex = this.data.getSelectedCharsStartIndex();
            endIndex = startIndex;
        }
        this.textArea.focus();
        this.textArea.setSelectionRange(startIndex, endIndex);
    }
    
    /**
     * 从文本区域同步光标位置到画布
     * 将隐藏文本输入框的光标位置同步到画布上显示
     * @private
     */
    private syncFromTextAreaCursor() {
        const selectionStart = this.textArea.selectionStart;
        const selectionEnd = this.textArea.selectionEnd;
        // 绝对索引 → 结构化位置（复用统一转换）
        const { lineIndex, charIndex } = this.data.fromAbsoluteIndex(selectionStart);
        // 计算光标坐标（复用 calcCursorX，确保 letterSpacing 累加规则一致）
        const editX = this.calcCursorX(lineIndex, charIndex);
        const editY = this.data.lines[lineIndex].editY;
        // 更新光标位置
        this.data.cursorPosition = {
            lineIndex,
            charIndex,
            editX,
            editY
        };
        
        // 更新光标显示
        this.updateCursor();
        
        // 处理选择范围
        // if (selectionStart !== selectionEnd) {
        //     this.data.selectionOnEditting = {
        //         start: { ...this.cursorPosition },
        //         end: { ...this.cursorPosition } // 需要进一步计算
        //     };
        //     // 这里需要实现选择范围的计算
        // } else {
        //     this.data.selectionOnEditting = null;
        // }
    }

    /**
     * 开始编辑单元格
     * 初始化编辑画布、光标位置，并保存初始值用于撤销
     * @public
     */
    public startEditting(refresh: boolean = false) {
        const cell = this.data.activedCell;
        if (!cell) return;
        
        // 保存初始值用于撤销记录（在编辑开始时记录整个单元格状态）
        const cellObj = this.data.values.find(v => v.cell === cell);
        if (cellObj && refresh && !this.data.isEditting) {  // 刷新单元格内容, 清空所有字符
            cellObj.chars = [];
        }
        this.data.isEditting = true;

        this.initialCellValue = cellObj?.chars ? cellObj.chars.map((v: Char) => v.clone()) : [];
        
        this.updateCanvasSize();   // 更新编辑画布大小
        this.drawEditCellsText(cell);  // 绘制单元格文本
        
        // 初始化光标位置到最后一个字符的位置
        if (this.data.lines.length > 0 && this.data.lines[0].text.length > 0) {
            this.data.cursorPosition = {
                lineIndex: this.data.lines.length - 1,
                charIndex: this.data.lines[this.data.lines.length - 1].text.length,
                editX: this.data.lines[this.data.lines.length - 1].editX + this.data.lines[this.data.lines.length - 1].lineWidth,
                editY: this.data.lines[this.data.lines.length - 1].editY
            };
        } else {
            // 空单元格的情况
            const { activedWidth, activedHeight } = this.data.getActivedRect();
            this.data.cursorPosition = {
                lineIndex: 0,
                charIndex: 0,
                editX: CELL_PADDING + activedWidth * this.data.zoom / 2,
                editY: CELL_PADDING + (activedHeight - DEFAULT_FONT_SIZE) * this.data.zoom / 2
            };
        }
        this.updateCursor();  // 更新光标位置和高度
        this.showEditor(); // 显示编辑器，即显示编辑框和光标
    }

    /**
     * 根据光标位置获取当前的字符索引并发出工具栏状态更新
     * @private
     */
    private notifyToolbarStateFromCursor() {
        if (!this.data.activedCell) return;
        
        // const { lineIndex, charIndex } = this.cursorPosition;
        // let totalIndex = 0;
        // for (let i = 0; i < lineIndex; i++) {
        //     totalIndex += this.data.lines[i].text.length; // 加上换行符
        // }
        // totalIndex += charIndex;
        // this.data.cursorPosition = totalIndex;
        this.data.notifyToolbarStateChange();
    }

    /**
     * 结束编辑单元格
     * 保存最终编辑结果，记录撤销操作，并隐藏编辑画布
     * @public
     */
    public finishEditting() {
        this.hideEditor();
        this.initialCellValue = [];
        this.data.isEditting = false;
        // 重置光标位置
        this.data.cursorAbsolutePosition = 0;
    }

    /**
     * 显示编辑画布和光标
     * 用于滚动或大小变化后保持编辑状态
     * @public
     */
    public showEditor() {
        this.canvas.style.display = 'block';
        this.cursor.style.display = 'block';
    }

    /**
     * 隐藏编辑画布和光标
     * @public
     */
    public hideEditor() {
        this.canvas.style.display = 'none';
        this.cursor.style.display = 'none';
    }

    /**
     * 处理键盘按下事件
     * 包括快捷键、导航键、编辑键等的处理
     * @param {KeyboardEvent} e - 键盘事件对象
     * @public
     */
    public handleKeyDown(e: KeyboardEvent): void {
        if (!this.data.activedCell) return;
        // 焦点位于其它可编辑元素（如打印预览面板的输入框）时，忽略表格的键盘处理：
        // 避免在这些输入框中打字误触发单元格编辑，或拦截复制/粘贴/撤销等快捷键。
        // 注意：this.textArea 自身不拦截，否则单元格编辑的键盘导航会失效。
        const keyTarget = e.target as HTMLElement | null;
        if (keyTarget && keyTarget !== this.textArea &&
            (keyTarget instanceof HTMLInputElement ||
             keyTarget instanceof HTMLSelectElement ||
             keyTarget instanceof HTMLTextAreaElement ||
             keyTarget.isContentEditable)) {
            return;
        }
        if (e.key && e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
            this.newChars.push(e.key);
            this.startEditting(true);
            return;
        }

        if (e.ctrlKey || e.metaKey) {
            switch (e.key.toLowerCase()) {
                case 'c':
                    e.preventDefault();
                    if (this.data.selectionOnEditting && this.data.isEditting) {
                        this.data.copyChars();
                    } else {
                        this.data.copySelection();
                    }
                    return;
                case 'x':
                    e.preventDefault();
                    if (this.data.selectionOnEditting && this.data.isEditting) {
                        // cutChars 只修改数据并返回目标光标绝对位置；
                        // 必须先重新测量文本(lines)，再写入 cursorAbsolutePosition，
                        // 否则会基于旧 lines 反算出错误的 lineIndex/charIndex。
                        const caretIndex = this.data.cutChars();
                        this.data.selectionOnEditting = null;
                        this.updateCanvasSize();
                        this.drawEditCellsText(this.data.activedCell);
                        this.data.cursorAbsolutePosition = caretIndex;
                    } else {
                        this.data.cutSelection();
                    }
                    return;
                case 'v':
                    e.preventDefault();
                    if (this.data.isEditting) {
                        const clipboard = this.data.clipboardManager.getClipboard();
                        if (clipboard && clipboard.type === 'char' && clipboard.chars) {
                            // getSelectedCharsStartIndex(): 有选区时返回选区起点，无选区时返回光标处，
                            // 即真正的插入起点。pasteChars 返回"插入完成后光标应处的绝对位置"
                            // （插入起点 + 新增字符数）。
                            const insertIndex = this.data.getSelectedCharsStartIndex();
                            const caretIndex = this.data.pasteChars(insertIndex, clipboard.chars);
                            // 必须在重新测量文本(lines)之后再写入 cursorAbsolutePosition，
                            // 否则 fromAbsoluteIndex 会用旧的 lines 把 charIndex 裁剪到过短的值，
                            // 导致光标只落在新增字符串的第一个字符之后。
                            this.data.selectionOnEditting = null;
                            this.updateCanvasSize();
                            this.drawEditCellsText(this.data.activedCell);
                            this.data.cursorAbsolutePosition = caretIndex;
                        }
                    } else {
                        this.data.pasteSelection();
                    }
                    return;
                case 'z':
                    e.preventDefault();
                    // Ctrl+Shift+Z 视为重做；普通 Ctrl+Z 视为撤销
                    if (e.shiftKey ? this.data.redo() : this.data.undo()) {
                        // 编辑期间撤销/重做后，values 已回写，
                        // 需重新同步编辑框文本、重绘并定位光标到内容末尾
                        if (this.data.isEditting) {
                            this.syncAfterEditUndoRedo();
                        }
                    }
                    return;
                case 'y':
                    e.preventDefault();
                    if (this.data.redo() && this.data.isEditting) {
                        this.syncAfterEditUndoRedo();
                    }
                    return;
            }
        }

        // 输入框获得焦点时的按键处理
        if (document.activeElement === this.textArea) {
            const { col, row } = this.data.getCellColAndRow(this.data.activedCell);
            let newActive = '';
            // 编辑状态下的键盘导航
            if (this.data.isEditting) {
                switch(e.key) {
                    case 'ArrowUp':
                        this.moveCursorUp();
                        e.preventDefault();
                        break;
                    case 'ArrowDown':
                        this.moveCursorDown();
                        e.preventDefault();
                        break;
                    case 'ArrowLeft':
                        this.moveCursorLeft();
                        e.preventDefault();
                        break;
                    case 'ArrowRight':
                        this.moveCursorRight();
                        e.preventDefault();
                        break;
                    case 'Home':
                        this.moveCursorToStart();
                        e.preventDefault();
                        break;
                    case 'PageUp':
                        this.moveCursorToLineStart();
                        e.preventDefault();
                        break;
                    case 'PageDown':
                        this.moveCursorToLineEnd();
                        e.preventDefault();
                        break;
                    case 'End':
                        this.moveCursorToEnd();
                        e.preventDefault();
                        break;
                    case 'Enter':
                        // 在编辑状态下，回车键插入换行符
                        e.preventDefault();
                        this.handleEnterKey();
                        break;
                    case 'Delete':
                    case 'Backspace':
                        // 编辑状态下处理删除操作
                        e.preventDefault();
                        this.handleDeleteKey(e.key === 'Delete');
                        break;
                    default:
                        // 普通字符输入，在输入前保存光标位置
                        if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
                            this.lastInputStartPos = this.textArea.selectionStart;
                        }
                        break;
                }
                return;
            } else {
                // 非编辑状态下的键盘处理
                switch(e.key) {
                    case 'ArrowUp':
                        if (row > 1) newActive = `${this.data.getColName(col)}${row-1}`;
                        break;
                    case 'ArrowDown':
                        if (row < this.data.rowHeaders.length) newActive = `${this.data.getColName(col)}${row+1}`;
                        break;
                    case 'ArrowLeft':
                        if (col > 1) newActive = `${this.data.getColName(col-1)}${row}`;
                        break;
                    case 'ArrowRight':
                        if (col < this.data.colHeaders.length) newActive = `${this.data.getColName(col+1)}${row}`;
                        break;
                    case 'Home':
                        newActive = `A${row}`;
                        break;
                    case 'End':
                        newActive = `${this.data.getColName(this.data.colHeaders.length)}${row}`;
                        break;
                    case 'PageUp':
                        const pageUpRow = Math.max(1, row - 10);
                        newActive = `${this.data.getColName(col)}${pageUpRow}`;
                        break;
                    case 'PageDown':
                        const pageDownRow = Math.min(this.data.rowHeaders.length, row + 10);
                        newActive = `${this.data.getColName(col)}${pageDownRow}`;
                        break;
                    case "Escape": // ESC键取消编辑
                        this.finishEditting();
                        break;
                    case "Delete":
                    case "Backspace": // 删除键处理
                        e.preventDefault();
                        if (!this.isComposing) { // 不在输入法组合状态时才处理
                            for (const cell of this.data.getSelectedCells()) {
                                const cellIndex = this.data.values.findIndex(v => v.cell === cell);
                                if (cellIndex >= 0) {
                                    this.data.values[cellIndex].chars = [];
                                }
                            }
                        }
                        this.data.commitValues();
                        break;
                    case 'Enter' : // 处理回车键
                        // 移动到下一行
                        e.preventDefault();    // 阻止默认行为，避免触发其他事件
                        if (row < this.data.rowHeaders.length) {
                            newActive = `${this.data.getColName(col)}${row+1}`;
                        }
                        break;
                    case 'Tab': // Tab键处理
                        e.preventDefault();
                        // 移动到下一列
                        if (col < this.data.colHeaders.length) {
                            newActive = `${this.data.getColName(col+1)}${row}`;
                        }
                        break;
                    default:
                        // 普通字符输入时进入编辑状态，保存光标位置
                        if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
                            this.lastInputStartPos = 0; // 非编辑状态下，光标位置为 0
                            this.startEditting();
                        }
                        break; 
                }
            }
            if (newActive) {
                // 确保单元格可见
                const { col, row } = this.data.getCellColAndRow(newActive);
                this.scrollCellIntoView(col, row);
                this.finishEditting();
                this.data.activedCell = newActive;
                // 只有在导航到新单元格时才更新 selection
                this.data.selection = `${this.data.activedCell}:${this.data.activedCell}`;
            }            
        }
    }
    
    /**
     * 向上移动光标
     * @private
     */
    private moveCursorUp() {
        if (this.data.cursorPosition.lineIndex > 0) {
            const newLineIndex = this.data.cursorPosition.lineIndex - 1;
            const newLine = this.data.lines[newLineIndex];
            const charIndex = Math.min(this.data.cursorPosition.charIndex, this.getLineEffectiveEnd(newLineIndex));
            
            this.data.cursorPosition = {
                lineIndex: newLineIndex,
                charIndex,
                editX: this.calcCursorX(newLineIndex, charIndex),
                editY: newLine.editY
            };
            
            this.afterCursorMove();
        }
    }
    
    /**
     * 向下移动光标
     * @private
     */
    private moveCursorDown() {
        if (this.data.cursorPosition.lineIndex < this.data.lines.length - 1) {
            const newLineIndex = this.data.cursorPosition.lineIndex + 1;
            const newLine = this.data.lines[newLineIndex];
            const charIndex = Math.min(this.data.cursorPosition.charIndex, this.getLineEffectiveEnd(newLineIndex));
            
            this.data.cursorPosition = {
                lineIndex: newLineIndex,
                charIndex,
                editX: this.calcCursorX(newLineIndex, charIndex),
                editY: newLine.editY
            };
            
            this.afterCursorMove();
        }
    }
    
    /**
     * 向左移动光标
     * @private
     */
    private moveCursorLeft() {
        if (this.data.cursorPosition.charIndex > 0) {
            // 同一行内向左移动
            const currentLine = this.data.lines[this.data.cursorPosition.lineIndex];
            const newCharIndex = this.data.cursorPosition.charIndex - 1;
            
            this.data.cursorPosition = {
                ...this.data.cursorPosition,
                charIndex: newCharIndex,
                editX: this.calcCursorX(this.data.cursorPosition.lineIndex, newCharIndex),
                editY: currentLine.editY
            };
        } else if (this.data.cursorPosition.lineIndex > 0) {
            // 移动到上一行的有效行尾（忽略行尾换行符）
            const newLineIndex = this.data.cursorPosition.lineIndex - 1;
            const newLine = this.data.lines[newLineIndex];
            const newCharIndex = this.getLineEffectiveEnd(newLineIndex);
            
            this.data.cursorPosition = {
                lineIndex: newLineIndex,
                charIndex: newCharIndex,
                editX: this.calcCursorX(newLineIndex, newCharIndex),
                editY: newLine.editY
            };
        }
        
        this.afterCursorMove();
    }
    
    /**
     * 向右移动光标
     * @private
     */
    private moveCursorRight() {
        const lineIndex = this.data.cursorPosition.lineIndex;
        const effectiveEnd = this.getLineEffectiveEnd(lineIndex);
        if (this.data.cursorPosition.charIndex < effectiveEnd) {
            // 同一行内向右移动（可达有效行尾，忽略行尾换行符）
            const newCharIndex = this.data.cursorPosition.charIndex + 1;
            
            this.data.cursorPosition = {
                ...this.data.cursorPosition,
                charIndex: newCharIndex,
                editX: this.calcCursorX(lineIndex, newCharIndex),
                editY: this.data.lines[lineIndex].editY
            };            
        } else if (lineIndex < this.data.lines.length - 1) {
            // 已在行尾，移动到下一行行首
            const newLineIndex = lineIndex + 1;
            const newLine = this.data.lines[newLineIndex];
            this.data.cursorPosition = {
                lineIndex: newLineIndex,
                charIndex: 0,
                editX: this.calcCursorX(newLineIndex, 0),
                editY: newLine.editY
            };  
        }
        this.afterCursorMove();
    }


    /**
     * 移动光标到文档开头
     * @private
     */
    private moveCursorToStart() {
        this.data.cursorPosition = {
            lineIndex: 0,
            charIndex: 0,
            editX: this.calcCursorX(0, 0),
            editY: this.data.lines[0].editY
        };
        this.afterCursorMove();
    }
    
    /**
     * 移动光标到文档结尾
     * @private
     */
    private moveCursorToEnd() {
        const lastLineIndex = this.data.lines.length - 1;
        const charIndex = this.getLineEffectiveEnd(lastLineIndex);
        this.data.cursorPosition = {
            lineIndex: lastLineIndex,
            charIndex,
            editX: this.calcCursorX(lastLineIndex, charIndex),
            editY: this.data.lines[lastLineIndex].editY
        };
        this.afterCursorMove();
    }
    
    /**
     * 移动光标到行首
     * @private
     */
    private moveCursorToLineStart() {
        const currentLine = this.data.lines[this.data.cursorPosition.lineIndex];

        this.data.cursorPosition = {
            ...this.data.cursorPosition,
            charIndex: 0,
            editX: currentLine.editX
        };

        this.afterCursorMove();
    }
    
    /**
     * 移动光标到行尾
     * @private
     */
    private moveCursorToLineEnd() {
        const lineIndex = this.data.cursorPosition.lineIndex;
        const charIndex = this.getLineEffectiveEnd(lineIndex);
        
        this.data.cursorPosition = {
            ...this.data.cursorPosition,
            charIndex,
            editX: this.calcCursorX(lineIndex, charIndex)
        };
        
        this.afterCursorMove();
    }
    
    /**
     * 处理回车键
     * 在光标位置插入换行符
     * @private
     */
    private handleEnterKey() {
        // 在光标位置插入换行符。插入与光标定位交由 handleValueChange 统一处理：
        // 它会在重新测量 lines 之后把光标置于换行符之后。
        this.newChars.push('\n');
        this.handleValueChange();
    }
    
    /**
     * 处理删除键
     * 支持 Delete 键和 Backspace 键
     * @param {boolean} isDelete - 是否为 Delete 键（true: Delete, false: Backspace）
     * @private
     */
    private handleDeleteKey(isDelete: boolean) {
        const cellValue = this.data.values.find(v => v.cell === this.data.activedCell)?.chars;
        if (!cellValue) return;
        // 统一以 cursorAbsolutePosition 为准计算删除区间；有选区时用选区端点。
        const hasSelection = !!this.data.selectionOnEditting;
        const selectionStart = hasSelection
            ? this.data.getSelectedCharsStartIndex()
            : this.data.cursorAbsolutePosition;
        const selectionEnd = hasSelection
            ? this.data.getSelectedCharsEndIndex()
            : selectionStart;
        // 删除完成后光标应处的绝对位置
        let caretIndex = selectionStart;
        if (selectionStart !== selectionEnd) {
            // 有选择内容，删除选择内容
            cellValue.splice(selectionStart, selectionEnd - selectionStart);
            caretIndex = selectionStart;
            // 删除所选字符后，清除高亮选区，避免残留高亮背景
            this.data.selectionOnEditting = null;
        } else if (isDelete && selectionStart < cellValue.length) {
            // Delete键，光标不在末尾时，删除光标后一个字符
            cellValue.splice(selectionStart, 1);
            caretIndex = selectionStart;
        } else {
            // 向左删除：包含两种情况
            // 1) Backspace 键；
            // 2) Delete 键但光标已处于单元格字符串最末位置（selectionStart === length），
            //    此时改为像 Backspace 一样从右往左逐个删除。
            caretIndex = this.deleteCharBackward(cellValue, selectionStart);
        }
        // 先重新测量文本(lines)，再按绝对索引统一设置光标并同步 textarea 选区。
        this.updateCanvasSize();
        this.drawEditCellsText(this.data.activedCell);
        this.data.cursorAbsolutePosition = caretIndex;
        this.cursor.style.display = 'block';
        this.data.commitValues();
    }

    /**
     * 向左（Backspace 方向）删除字符。
     *
     * 常规情况删除光标前一个字符；但当光标前一个字符是换行符时
     * （即光标位于行头、上一行行尾存在换行符），需要把该换行符
     * 及其之前的一个字符一起删除。
     *
     * 该判断完全基于绝对索引与字符数组内容，避免依赖行边界处
     * cursorPosition 的歧义（fromAbsoluteIndex 在行边界会归属到上一行行尾）。
     *
     * @param {Char[]} cellValue - 当前单元格的字符数组（会被原地修改）
     * @param {number} caretIndex - 删除前光标的绝对位置
     * @returns {number} 删除完成后光标应处的绝对位置
     * @private
     */
    private deleteCharBackward(cellValue: Char[], caretIndex: number): number {
        if (caretIndex <= 0) return 0;
        // 光标前一个字符是换行符，且其前面还有可删除的字符时，
        // 连同换行符前的一个字符一起删除（共删除 2 个字符）。
        if (cellValue[caretIndex - 1]?.char === '\n' && caretIndex >= 2) {
            cellValue.splice(caretIndex - 2, 2);
            return caretIndex - 2;
        }
        // 常规：删除光标前一个字符（含换行符位于串首、其前无字符的情况）。
        cellValue.splice(caretIndex - 1, 1);
        return caretIndex - 1;
    }

    /**
     * 编辑期间撤销/重做后的编辑框重同步
     * values 已由 DataCollection 回写，这里负责按正确顺序重同步：
     * 1. 清除可能残留的选区，使光标以 cursorAbsolutePosition 为唯一真相源；
     * 2. updateCanvasSize()：按回写后的 chars 重新测量 lines 并同步 textArea.value；
     * 3. drawEditCellsText()：重绘编辑画布文本；
     * 4. 按 editRestoreCursor 定位光标（必须发生在 lines 刷新之后，
     *    cursorAbsolutePosition 会同步 textarea 选区与画布光标）；
     * 5. 最后才广播 VALUES_CHANGED，确保 sheets.draw() 触发的重绘使用的是已定位好的光标，
     *    避免在旧光标下绘制导致光标错乱。
     * @private
     */
    private syncAfterEditUndoRedo(): void {
        const cell = this.data.activedCell;
        this.data.selectionOnEditting = null;
        this.updateCanvasSize();
        this.drawEditCellsText(cell);
        this.data.cursorAbsolutePosition = this.data.editRestoreCursor;
        this.data.emit(DataEvents.VALUES_CHANGED, this.data.values);
    }

    /**
     * 处理值变化
     * 当文本输入框的值发生变化时更新单元格数据
     * 注意：编辑过程中的每次按键变更会记入编辑内撤销栈（_editUndoStack），
     * 整段编辑在 finishEditting 时折叠为 undoManager 的一条文档级撤销操作。
     * @private
     */
    private handleValueChange() {
        if (this.isComposing) return;

        const cell = this.data.activedCell;
        if (!cell) return;

        // 处理添加字符：在当前光标处（或选区起点）插入，返回插入完成后光标应处的绝对位置
        const caretIndex = this.handleAddChars();
        this.updateCanvasSize();
        this.drawEditCellsText(cell);
        // 必须在 updateCanvasSize 重新测量 lines 之后再写入光标位置，
        // 才能把绝对索引正确反算为最新布局下的 lineIndex/charIndex，
        // 并同步 textarea.setSelectionRange，保证光标落在新增字符串之后。
        if (caretIndex !== null) {
            this.data.cursorAbsolutePosition = caretIndex;
        } else {
            // 无新增字符（如删除后重绘），仍按当前绝对位置重新解析并同步。
            this.syncFromTextAreaCursor();
        }
        this.cursor.style.display = 'block';
        // 提交当前单元格的值到数据集合
        this.data.commitValues();
    }

    /**
     * 处理添加字符
     * 在当前光标位置（若有选区则先删除选区内容再从选区起点）插入 newChars。
     * 仅修改字符数据，不在此处根据旧 lines 反算光标，光标定位由调用方在
     * 重新测量 lines 之后统一处理。
     * @returns {number | null} 插入完成后光标应处的绝对位置（插入起点 + 新增字符数）；
     *                          无新增字符时返回 null
     * @private
     */
    private handleAddChars(): number | null {
        //获取当前单元格的旧值
        const oldCellObj = this.data.values.find(v => v.cell === this.data.activedCell);
        if (this.newChars.length === 0) {
            return null;
        }
        const addedText = this.newChars.join('');
        let caretIndex: number;
        if (oldCellObj) { // 插入新增字符
            let insertIdx = 0;
            if (this.data.selectionOnEditting) { // 有选择内容，先删除选择内容
                const startIndex = this.data.toAbsoluteIndex(this.data.selectionOnEditting.start);
                const endIndex = this.data.toAbsoluteIndex(this.data.selectionOnEditting.end);

                oldCellObj.chars.splice(startIndex, endIndex - startIndex);  // 先删除所选择内容
                insertIdx = startIndex;
                this.data.selectionOnEditting = null;
            } else {
                insertIdx = this.data.cursorAbsolutePosition;
            }
            oldCellObj.chars.splice(insertIdx, 0, ...addedText.split('').map(char => {
                const v = new Char();
                v.char = char;
                v.fontFamily = oldCellObj.chars[insertIdx - 1]?.fontFamily;
                v.fontSize = oldCellObj.chars[insertIdx - 1]?.fontSize;
                v.fontWeight = oldCellObj.chars[insertIdx - 1]?.fontWeight;
                v.fontStyle = oldCellObj.chars[insertIdx - 1]?.fontStyle;
                v.fontColor = oldCellObj.chars[insertIdx - 1]?.fontColor;
                v.underline = oldCellObj.chars[insertIdx - 1]?.underline;
                v.strikethrough = oldCellObj.chars[insertIdx - 1]?.strikethrough;
                return v;
            }));
            // 光标落在插入字符之后
            caretIndex = insertIdx + addedText.length;
        } else {
            // 如果单元格不存在，创建新的单元格对象
            const newCell = new Cell();
            newCell.cell = this.data.activedCell;
            newCell.chars = addedText.split('').map(char => {
                const v = new Char();
                v.char = char;
                return v;
            });
            this.data.values.push(newCell);
            // 新建单元格，光标在所有字符之后
            caretIndex = addedText.length;
        }
        // 清空新Chars
        this.newChars = [];
        return caretIndex;
    }

    /**
     * 绘制编辑画布
     * 重新绘制文本和选择范围
     * @public
     */
    public draw(): void {
        if (!this.data.isEditting) return;
        this.updateCanvasSize();
        this.updateCursor();
        this.drawEditCellsText(this.data.activedCell);
        this.drawSelection();
    }
    /**
     * 获取编辑画布的 DOM 元素
     * @returns {Object} 包含编辑画布、上下文和文本输入框的对象
     * @public
     */
    public getElements(): { editCanvas: HTMLCanvasElement, editCtx: CanvasRenderingContext2D, textArea: HTMLTextAreaElement } {
        return {
            editCanvas: this.canvas,
            editCtx: this.ctx,
            textArea: this.textArea,
        }
    }

    /**
     * 销毁编辑画布
     * 移除所有 DOM 元素
     * @public
     */
    public destroy(): void {
        this.canvas.remove();
        this.cursor.remove();
        this.textArea.remove();
    }
}