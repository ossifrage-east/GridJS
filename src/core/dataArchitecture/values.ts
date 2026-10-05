import { DataEvents } from "../constant";
import { EventEmitter } from "../../utils/eventEmitter";

/**
 * 值对象接口
 */
export interface ValueJSON {
    
}

/**
 * 值类
 */
export class Value extends EventEmitter {
    /** 值对象 */
    private _valueObj: ValueJSON;
    /** 值 */
    private _value: any;

    constructor() {
        super();
    }
}