/*  # DOM操作工具模块 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

/**
 * 创建DOM元素的选项接口
 */
export interface CreateElementOptions {
    className?: string;
    id?: string;
    textContent?: string;
    innerHTML?: string;
    style?: Partial<CSSStyleDeclaration>;
    attributes?: Record<string, string>;
    dataset?: Record<string, string>;
    children?: HTMLElement[];
    contentEditable?: string;
    tabindex?: string;
    events?: Record<string, EventListenerOrEventListenerObject>;
}

/**
 * 创建div元素
 * @param options 创建元素的选项
 * @returns 创建的div元素
 */
export function createDiv(options: CreateElementOptions = {}): HTMLDivElement {
    return createElement('div', options) as HTMLDivElement;
}

/**
 * 创建span元素
 * @param options 创建元素的选项
 * @returns 创建的span元素
 */
export function createSpan(options: CreateElementOptions = {}): HTMLSpanElement {
    return createElement('span', options) as HTMLSpanElement;
}

/**
 * 创建input元素
 * @param options 创建元素的选项
 * @returns 创建的input元素
 */
export function createInput(options: CreateElementOptions & { type?: string } = {}): HTMLInputElement {
    const input = createElement('input', options) as HTMLInputElement;
    if (options.type) {
        input.type = options.type;
    }
    return input;
}

/**
 * 创建button元素
 * @param options 创建元素的选项
 * @returns 创建的button元素
 */
export function createButton(options: CreateElementOptions = {}): HTMLButtonElement {
    return createElement('button', options) as HTMLButtonElement;
}

/**
 * 创建任意HTML元素
 * @param tagName 标签名
 * @param options 创建元素的选项
 * @returns 创建的HTML元素
 */
export function createElement(tagName: string, options: CreateElementOptions = {}): HTMLElement {
    const element = document.createElement(tagName);

    // 设置类名
    if (options.className) {
        element.className = options.className;
    }

    // 设置ID
    if (options.id) {
        element.id = options.id;
    }

    // 设置文本内容
    if (options.textContent) {
        element.textContent = options.textContent;
    }

    // 设置HTML内容
    if (options.innerHTML) {
        element.innerHTML = options.innerHTML;
    }

    // 设置样式
    if (options.style) {
        Object.assign(element.style, options.style);
    }

    // 设置属性
    if (options.attributes) {
        for (const [key, value] of Object.entries(options.attributes)) {
            element.setAttribute(key, value);
        }
    }

    // 设置数据集
    if (options.dataset) {
        for (const [key, value] of Object.entries(options.dataset)) {
            element.dataset[key] = value;
        }
    }

    // 添加子元素
    if (options.children) {
        for (const child of options.children) {
            element.appendChild(child);
        }
    }

    // 添加事件监听器
    if (options.events) {
        for (const [eventName, handler] of Object.entries(options.events)) {
            element.addEventListener(eventName, handler);
        }
    }

    return element;
}

/**
 * 移除元素
 * @param element 要移除的元素
 */
export function removeElement(element: HTMLElement): void {
    if (element && element.parentNode) {
        element.parentNode.removeChild(element);
    }
}

/**
 * 清空元素内容
 * @param element 要清空的元素
 */
export function clearElement(element: HTMLElement): void {
    if (element) {
        element.innerHTML = '';
    }
}

/**
 * 获取元素的绝对位置
 * @param element 元素
 * @returns 元素的绝对位置 {left, top}
 */
export function getElementPosition(element: HTMLElement): { left: number; top: number } {
    const rect = element.getBoundingClientRect();
    const scrollLeft = window.pageXOffset || document.documentElement.scrollLeft;
    const scrollTop = window.pageYOffset || document.documentElement.scrollTop;
    
    return {
        left: rect.left + scrollLeft,
        top: rect.top + scrollTop
    };
}

/**
 * 判断元素是否包含指定类名
 * @param element 元素
 * @param className 类名
 * @returns 是否包含类名
 */
export function hasClass(element: HTMLElement, className: string): boolean {
    return element.classList.contains(className);
}

/**
 * 添加类名
 * @param element 元素
 * @param className 类名
 */
export function addClass(element: HTMLElement, className: string): void {
    element.classList.add(className);
}

/**
 * 移除类名
 * @param element 元素
 * @param className 类名
 */
export function removeClass(element: HTMLElement, className: string): void {
    element.classList.remove(className);
}

/**
 * 切换类名
 * @param element 元素
 * @param className 类名
 */
export function toggleClass(element: HTMLElement, className: string): void {
    element.classList.toggle(className);
}