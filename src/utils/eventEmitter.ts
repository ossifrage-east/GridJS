/*  # 事件发射器工具模块 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { DataEvents } from "../core/constant";

/**
 * 事件处理函数类型
 */
export type EventHandler = (...args: any[]) => void;

/**
 * 事件发射器类
 * 用于实现事件的订阅和发布
 */
export class EventEmitter {
    private events: Map<string, EventHandler[]> = new Map();

    /**
     * 注册事件监听器
     * @param event 事件名称
     * @param handler 事件处理函数
     * @returns 当前实例，用于链式调用
     */
    public on(event: string, handler: EventHandler): this {
        if (!this.events.has(event)) {
            this.events.set(event, []);
        }
        
        const handlers = this.events.get(event)!;
        handlers.push(handler);
        
        return this;
    }

    /**
     * 注册一次性事件监听器，触发后自动移除
     * @param event 事件名称
     * @param handler 事件处理函数
     * @returns 当前实例，用于链式调用
     */
    public once(event: string, handler: EventHandler): this {
        const onceHandler = (...args: any[]) => {
            handler(...args);
            this.off(event, onceHandler);
        };
        
        return this.on(event, onceHandler);
    }

    /**
     * 移除事件监听器
     * @param event 事件名称
     * @param handler 事件处理函数，如果不提供则移除该事件的所有监听器
     * @returns 当前实例，用于链式调用
     */
    public off(event: string, handler?: EventHandler): this {
        if (!this.events.has(event)) {
            return this;
        }
        
        if (!handler) {
            // 移除该事件的所有监听器
            this.events.delete(event);
            return this;
        }
        
        const handlers = this.events.get(event)!;
        const index = handlers.indexOf(handler);
        
        if (index !== -1) {
            handlers.splice(index, 1);
            
            // 如果没有监听器了，删除该事件
            if (handlers.length === 0) {
                this.events.delete(event);
            }
        }
        
        return this;
    }

    /**
     * 触发事件
     * @param event 事件名称
     * @param args 传递给事件处理函数的参数
     * @returns 当前实例，用于链式调用
     */
    public emit(event: string, ...args: any[]): this {
        if (!this.events.has(event)) {
            return this;
        }
        
        const handlers = this.events.get(event)!.slice();
        
        for (const handler of handlers) {
            try {
                handler(...args);
            } catch (error) {
                console.error(`Error in event handler for "${event}":`, error);
            }
        }
        
        return this;
    }

    /**
     * 获取指定事件的所有监听器
     * @param event 事件名称
     * @returns 事件处理函数数组
     */
    public listeners(event: string): EventHandler[] {
        return this.events.has(event) ? [...this.events.get(event)!] : [];
    }

    /**
     * 获取所有已注册的事件名称
     * @returns 事件名称数组
     */
    public eventNames(): string[] {
        return Array.from(this.events.keys());
    }
    
    /**
     * 更新属性并触发变更事件
     * @param {SheetProperty} propName - 属性名称
     * @param {string} internalProp - 内部属性名
     * @param {any} value - 新值
     */
    public updateProperty(propName: DataEvents, internalProp: string, value: any): void {
        const oldValue = (this as any)[internalProp];
        if (oldValue !== value) {
            (this as any)[internalProp] = value;
            this.emit(propName, oldValue, value);
        }
    }

    /**
     * 移除所有事件监听器
     * @returns 当前实例，用于链式调用
     */
    public removeAllListeners(): this {
        this.events.clear();
        return this;
    }
}