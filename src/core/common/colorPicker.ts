/*  # 通用颜色选择器组件 TS 文件
    # author: 东方鹗
    # B站: https://space.bilibili.com/194359739
    # 知乎: https://www.zhihu.com/people/eastossifrage
    # CSDN: https://blog.csdn.net/os373
*/

import { createDiv, createSpan } from '../../utils/dom';
import { EventEmitter } from '../../utils/eventEmitter';

export interface ColorPickerOptions {
    parentElement?: HTMLElement;
    initialColor?: string;
    onChange?: (color: string | null) => void;
}

export class ColorPicker extends EventEmitter {
    private parentElement: HTMLElement;
    private pickerContainer: HTMLDivElement;
    private initColor: string | null;
    private currentColor: string | null;
    private advancedPickerOpen: boolean = false;
    private advancedPickerContainer: HTMLDivElement;
    private colorPreviewCurrent: HTMLElement;
    private colorPreviewNew: HTMLElement;
    private colorSpectrum: HTMLElement;
    private spectrumSelector: HTMLElement;
    private tempColor: string | null;
    private rgbaInput!: HTMLInputElement;
    private eyeDropperButton!: HTMLButtonElement;
    private confirmButton!: HTMLButtonElement;
    private cancelButton!: HTMLButtonElement;
    private currentHue: number = 0;
    private currentSaturation: number = 100;
    private currentLightness: number = 50;
    private isDragging: boolean = false;

    private themeColors: string[][] = [
        ['#000000', '#595959', '#7F7F7F', '#A5A5A5', '#BFBFBF', '#BBBBBB', '#CCCCCC', '#DDDDDD', '#E5E5E5', '#EFEFEF'],
        ['#44546A', '#4A5C70', '#5B6B88', '#8496B0', '#A9B3C8', '#A8B9D0', '#BCC7D8', '#CEDAE4', '#D8E1EA', '#DFE6EE'],
        ['#0F6FC6', '#1A7CDD', '#4A95E0', '#76AEE9', '#A3C7F0', '#98C1EF', '#B5D0F4', '#CBDBF7', '#DBE8FB', '#E5F0FD'],
        ['#ED7D31', '#F4A460', '#F7B989', '#F9CDA8', '#FCE0C7', '#FBDBC0', '#FDE0CA', '#FEE9D5', '#FEF2E1', '#FEF7EB'],
        ['#FFC000', '#FFD966', '#FFE699', '#FFF2CC', '#FFF8E6', '#FFF7D9', '#FFF9E6', '#FFFBF0', '#FFFEF7', '#FFFFFB'],
        ['#70AD47', '#8CC168', '#A8D08D', '#C4E0B2', '#DFF0D8', '#CCEBC5', '#D7F0CE', '#E1F5D8', '#EBF9E1', '#F2FCEB'],
        ['#4472C4', '#6C8CD5', '#94AAE0', '#B9C7EB', '#D6E0F5', '#C7D6EE', '#D5E0F2', '#E2EBF7', '#EAF1FB', '#F2F7FD'],
        ['#5B9BD5', '#7CAFDD', '#9DC3E6', '#BDD7EE', '#DBEAF7', '#C9DEF0', '#D9E8F4', '#E5F0F8', '#EFF6FC', '#F6FAFE']
    ];

    private standardColors: string[] = [
        '#FF0000', '#FF9900', '#FFFF00', '#00FF00', '#00FFFF',
        '#0000FF', '#9900FF', '#FF00FF', '#FF0066', '#00FF99'
    ];

    constructor(options: ColorPickerOptions) {
        super();
        this.parentElement = options.parentElement;
        this.initColor = options.initialColor || null;
        this.currentColor = this.initColor;
        this.tempColor = this.currentColor;
        
        this.createColorPicker();
        
        if (options.onChange) {
            this.on('change', options.onChange);
        }
        
        this.setupClickOutsideListener();
    }

    private createColorPicker(): void {
        this.pickerContainer = createDiv({
            className: 'color-picker-container',
            style: {
                width: '220px',
                backgroundColor: '#ffffff',
                border: '1px solid #d1d1d1',
                boxShadow: '0 2px 5px rgba(0, 0, 0, 0.15)',
                zIndex: '1000',
                padding: '8px'
            }
        });
        
        if (this.parentElement) {
            this.parentElement.appendChild(this.pickerContainer);
        }
        
        // 无填充颜色选项
        const noFillOption = createDiv({
            className: 'no-fill-option',
            style: {
                display: 'flex',
                alignItems: 'center',
                padding: '5px',
                cursor: 'pointer',
                marginBottom: '8px',
                borderRadius: '3px',
                transition: 'background-color 0.2s'
            }
        });
        
        noFillOption.addEventListener('mouseover', () => {
            noFillOption.style.backgroundColor = '#f0f0f0';
        });
        
        noFillOption.addEventListener('mouseout', () => {
            noFillOption.style.backgroundColor = 'transparent';
        });
        
        const noFillIcon = createDiv({
            style: {
                width: '16px',
                height: '16px',
                border: '1px solid #d1d1d1',
                marginRight: '8px',
                position: 'relative',
                backgroundColor: this.initColor === 'black' ? 'black' : ''
            }
        });
        
        const redLine = createDiv({
            style: {
                position: 'absolute',
                width: '22px',
                height: '1px',
                backgroundColor: 'red',
                transform: 'rotate(45deg)',
                transformOrigin: 'left top',
                top: '0',
                left: '0'
            }
        });
        
        if (!this.initColor) {
            noFillIcon.appendChild(redLine);
        }
        noFillOption.appendChild(noFillIcon);
        noFillOption.appendChild(createSpan({ textContent: this.initColor === 'black' ? '默认颜色黑色' : '无填充颜色' }));
        
        noFillOption.addEventListener('click', () => {
            this.selectColor('none');
        });
        
        this.pickerContainer.appendChild(noFillOption);
        
        // 主题颜色部分
        this.pickerContainer.appendChild(createSpan({ 
            textContent: '主题颜色', 
            style: { 
                display: 'block', 
                marginBottom: '5px',
                fontSize: '12px',
                fontWeight: 'bold'
            } 
        }));
        
        const themeColorsContainer = createDiv({
            className: 'theme-colors-container',
            style: {
                display: 'grid',
                gridTemplateColumns: 'repeat(10, 1fr)',
                gap: '2px',
                marginBottom: '10px'
            }
        });
        
        this.themeColors.forEach(colorRow => {
            colorRow.forEach(color => {
                const colorCell = createDiv({
                    style: {
                        width: '16px',
                        height: '16px',
                        backgroundColor: color,
                        border: '1px solid #d1d1d1',
                        cursor: 'pointer'
                    }
                });
                
                colorCell.addEventListener('click', () => {
                    this.selectColor(color);
                });
                
                themeColorsContainer.appendChild(colorCell);
            });
        });
        
        this.pickerContainer.appendChild(themeColorsContainer);
        
        // 标准色部分
        this.pickerContainer.appendChild(createSpan({ 
            textContent: '标准色', 
            style: { 
                display: 'block', 
                marginBottom: '5px',
                fontSize: '12px',
                fontWeight: 'bold'
            } 
        }));
        
        const standardColorsContainer = createDiv({
            className: 'standard-colors-container',
            style: {
                display: 'grid',
                gridTemplateColumns: 'repeat(10, 1fr)',
                gap: '2px',
                marginBottom: '10px'
            }
        });
        
        this.standardColors.forEach(color => {
            const colorCell = createDiv({
                style: {
                    width: '16px',
                    height: '16px',
                    backgroundColor: color,
                    border: '1px solid #d1d1d1',
                    cursor: 'pointer'
                }
            });
            
            colorCell.addEventListener('click', () => {
                this.selectColor(color);
            });
            
            standardColorsContainer.appendChild(colorCell);
        });
        
        this.pickerContainer.appendChild(standardColorsContainer);
        
        // 其他颜色选项
        const otherColorOption = createDiv({
            className: 'other-color-option',
            style: {
                display: 'flex',
                alignItems: 'center',
                padding: '5px',
                cursor: 'pointer'
            }
        });
        
        otherColorOption.appendChild(createSpan({ textContent: '其他颜色(M)...' }));
        
        otherColorOption.addEventListener('click', () => {
            this.showAdvancedPicker();
        });
        
        this.pickerContainer.appendChild(otherColorOption);
        
        this.createAdvancedColorPicker();
    }
    
    private createAdvancedColorPicker(): void {
        this.advancedPickerContainer = createDiv({
            className: 'advanced-color-picker',
            style: {
                position: 'absolute',
                width: '280px',
                backgroundColor: '#ffffff',
                border: '1px solid #d1d1d1',
                borderRadius: '4px',
                boxShadow: '0 4px 12px rgba(0, 0, 0, 0.15)',
                zIndex: '1001',
                padding: '15px',
                display: 'none'
            }
        });
        
        document.body.appendChild(this.advancedPickerContainer);
        
        // 颜色预览
        const previewContainer = createDiv({
            style: {
                display: 'flex',
                marginBottom: '12px',
                alignItems: 'center'
            }
        });
        
        this.colorPreviewCurrent = createDiv({
            style: {
                width: '40px',
                height: '40px',
                border: '1px solid #d1d1d1',
                marginRight: '10px',
                backgroundColor: this.currentColor || 'transparent'
            }
        });
        
        this.colorPreviewNew = createDiv({
            style: {
                flex: '1',
                height: '40px',
                border: '1px solid #d1d1d1',
                backgroundColor: this.tempColor || 'transparent'
            }
        });
        
        previewContainer.appendChild(this.colorPreviewCurrent);
        previewContainer.appendChild(this.colorPreviewNew);
        
        this.advancedPickerContainer.appendChild(previewContainer);
        
        // 颜色光谱选择区域
        this.colorSpectrum = createDiv({
            style: {
                width: '100%',
                height: '150px',
                background: 'linear-gradient(to bottom, rgb(255, 0, 0) 0%, rgb(255, 255, 0) 17%, rgb(0, 255, 0) 33%, rgb(0, 255, 255) 50%, rgb(0, 0, 255) 67%, rgb(255, 0, 255) 83%, rgb(255, 0, 0) 100%)',
                position: 'relative',
                marginBottom: '12px',
                cursor: 'crosshair',
                borderRadius: '3px'
            }
        });
        
        this.spectrumSelector = createDiv({
            style: {
                width: '12px',
                height: '12px',
                border: '2px solid white',
                boxShadow: '0 0 3px rgba(0,0,0,0.5)',
                borderRadius: '50%',
                position: 'absolute',
                top: '0',
                left: '0',
                transform: 'translate(-50%, -50%)',
                pointerEvents: 'none',
                boxSizing: 'border-box'
            }
        });
        
        this.colorSpectrum.appendChild(this.spectrumSelector);
        this.advancedPickerContainer.appendChild(this.colorSpectrum);
        
        // RGBA 输入框
        const rgbaLabelRow = createDiv({
            style: {
                marginBottom: '8px'
            }
        });
        rgbaLabelRow.appendChild(createSpan({ 
            textContent: 'RGBA 颜色值:', 
            style: { fontSize: '12px' } 
        }));
        this.advancedPickerContainer.appendChild(rgbaLabelRow);
        
        this.rgbaInput = document.createElement('input');
        this.rgbaInput.type = 'text';
        this.rgbaInput.value = this.currentColor || 'rgba(0, 0, 0, 1)';
        this.rgbaInput.style.width = '100%';
        this.rgbaInput.style.padding = '6px 8px';
        this.rgbaInput.style.boxSizing = 'border-box';
        this.rgbaInput.style.border = '1px solid #d1d1d1';
        this.rgbaInput.style.borderRadius = '3px';
        this.rgbaInput.style.marginBottom = '12px';
        this.rgbaInput.style.fontSize = '12px';
        
        this.rgbaInput.addEventListener('input', () => {
            const value = this.rgbaInput.value;
            this.tempColor = value;
            this.colorPreviewNew.style.backgroundColor = value;
        });
        
        this.advancedPickerContainer.appendChild(this.rgbaInput);
        
        // 拾色器和按钮行
        const buttonRow = createDiv({
            style: {
                display: 'flex',
                gap: '8px',
                marginBottom: '12px'
            }
        });
        
        // 拾色器按钮
        this.eyeDropperButton = document.createElement('button');
        this.eyeDropperButton.textContent = '🎯 拾取颜色';
        this.eyeDropperButton.style.flex = '1';
        this.eyeDropperButton.style.padding = '6px';
        this.eyeDropperButton.style.cursor = 'pointer';
        this.eyeDropperButton.style.backgroundColor = '#f0f0f0';
        this.eyeDropperButton.style.border = '1px solid #d1d1d1';
        this.eyeDropperButton.style.borderRadius = '3px';
        this.eyeDropperButton.style.fontSize = '12px';
        
        if ('EyeDropper' in window) {
            this.eyeDropperButton.addEventListener('click', async () => {
                try {
                    const eyeDropper = new (window as any).EyeDropper();
                    const result = await eyeDropper.open();
                    const colorValue = result.sRGBHex;
                    this.setColorFromHex(colorValue);
                    this.updatePreviewFromColor();
                } catch (e) {
                    // 用户取消拾取
                }
            });
        } else {
            this.eyeDropperButton.disabled = true;
            this.eyeDropperButton.title = '您的浏览器不支持拾色器功能';
            this.eyeDropperButton.style.opacity = '0.5';
        }
        
        buttonRow.appendChild(this.eyeDropperButton);
        this.advancedPickerContainer.appendChild(buttonRow);
        
        // 确认和取消按钮
        const actionButtonRow = createDiv({
            style: {
                display: 'flex',
                justifyContent: 'flex-end'
            }
        });
        
        this.cancelButton = document.createElement('button');
        this.cancelButton.textContent = '取消';
        this.cancelButton.style.padding = '6px 16px';
        this.cancelButton.style.marginRight = '8px';
        this.cancelButton.style.cursor = 'pointer';
        this.cancelButton.style.backgroundColor = '#ffffff';
        this.cancelButton.style.border = '1px solid #d1d1d1';
        this.cancelButton.style.borderRadius = '3px';
        this.cancelButton.style.fontSize = '12px';
        this.cancelButton.addEventListener('click', (e) => {
            e.stopPropagation();
            this.tempColor = this.currentColor;
            this.hideAdvancedPicker();
        });
        
        this.confirmButton = document.createElement('button');
        this.confirmButton.textContent = '确认';
        this.confirmButton.style.padding = '6px 16px';
        this.confirmButton.style.cursor = 'pointer';
        this.confirmButton.style.backgroundColor = '#0078d4';
        this.confirmButton.style.color = '#ffffff';
        this.confirmButton.style.border = '1px solid #0078d4';
        this.confirmButton.style.borderRadius = '3px';
        this.confirmButton.style.fontSize = '12px';
        this.confirmButton.addEventListener('click', (e) => {
            e.stopPropagation();
            this.selectColor(this.tempColor);
            this.hideAdvancedPicker();
        });
        
        actionButtonRow.appendChild(this.cancelButton);
        actionButtonRow.appendChild(this.confirmButton);
        
        this.advancedPickerContainer.appendChild(actionButtonRow);
        
        // 颜色光谱拖动事件
        this.setupSpectrumInteraction();
    }
    
    private setupSpectrumInteraction(): void {
        const updateColorFromSpectrum = (e: MouseEvent) => {
            const rect = this.colorSpectrum.getBoundingClientRect();
            let x = e.clientX - rect.left;
            let y = e.clientY - rect.top;
            
            x = Math.max(0, Math.min(rect.width, x));
            y = Math.max(0, Math.min(rect.height, y));
            
            // 色相从顶部到底部（0-360）
            this.currentHue = Math.round((y / rect.height) * 360);
            
            // 饱和度从左到右（0-100）
            this.currentSaturation = Math.round((x / rect.width) * 100);
            
            // 亮度从左到右
            this.currentLightness = Math.round(50 - (x / rect.width) * 50);
            
            this.updateSpectrumSelector();
            this.updatePreviewFromColor();
        };
        
        this.colorSpectrum.addEventListener('mousedown', (e) => {
            this.isDragging = true;
            updateColorFromSpectrum(e);
            e.preventDefault();
        });
        
        document.addEventListener('mousemove', (e) => {
            if (this.isDragging) {
                updateColorFromSpectrum(e);
            }
        });
        
        document.addEventListener('mouseup', () => {
            this.isDragging = false;
        });
    }
    
    private updateSpectrumSelector(): void {
        const rect = this.colorSpectrum.getBoundingClientRect();
        const selectorWidth = this.spectrumSelector.offsetWidth;
        const selectorHeight = this.spectrumSelector.offsetHeight;
        
        let x = (this.currentSaturation / 100) * rect.width;
        let y = (this.currentHue / 360) * rect.height;
        
        x = Math.max(selectorWidth / 2, Math.min(rect.width - selectorWidth / 2, x));
        y = Math.max(selectorHeight / 2, Math.min(rect.height - selectorHeight / 2, y));
        
        this.spectrumSelector.style.left = `${x}px`;
        this.spectrumSelector.style.top = `${y}px`;
    }
    
    private setColorFromHex(hex: string): void {
        const r = parseInt(hex.substring(1, 3), 16);
        const g = parseInt(hex.substring(3, 5), 16);
        const b = parseInt(hex.substring(5, 7), 16);
        
        const hsl = this.rgbToHsl(r, g, b);
        this.currentHue = Math.round(hsl.h);
        this.currentSaturation = Math.round(hsl.s);
        this.currentLightness = Math.round(hsl.l);
        
        this.updateSpectrumSelector();
        this.updatePreviewFromColor();
    }
    
    private updatePreviewFromColor(): void {
        const rgb = this.hslToRgb(this.currentHue, this.currentSaturation, this.currentLightness);
        this.tempColor = `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, 1)`;
        this.rgbaInput.value = this.tempColor;
        this.colorPreviewNew.style.backgroundColor = this.tempColor;
    }
    
    private hslToRgb(h: number, s: number, l: number): { r: number; g: number; b: number } {
        h /= 360;
        s /= 100;
        l /= 100;
        
        let r, g, b;
        
        if (s === 0) {
            r = g = b = l;
        } else {
            const hue2rgb = (p: number, q: number, t: number) => {
                if (t < 0) t += 1;
                if (t > 1) t -= 1;
                if (t < 1/6) return p + (q - p) * 6 * t;
                if (t < 1/2) return q;
                if (t < 2/3) return p + (q - p) * (2/3 - t) * 6;
                return p;
            };
            
            const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
            const p = 2 * l - q;
            
            r = hue2rgb(p, q, h + 1/3);
            g = hue2rgb(p, q, h);
            b = hue2rgb(p, q, h - 1/3);
        }
        
        return {
            r: Math.round(r * 255),
            g: Math.round(g * 255),
            b: Math.round(b * 255)
        };
    }
    
    private rgbToHsl(r: number, g: number, b: number): { h: number; s: number; l: number } {
        r /= 255;
        g /= 255;
        b /= 255;
        
        const max = Math.max(r, g, b);
        const min = Math.min(r, g, b);
        let h = 0, s = 0;
        const l = (max + min) / 2;
        
        if (max !== min) {
            const d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            
            switch (max) {
                case r: h = (g - b) / d + (g < b ? 6 : 0); break;
                case g: h = (b - r) / d + 2; break;
                case b: h = (r - g) / d + 4; break;
            }
            
            h /= 6;
        }
        
        return {
            h: h * 360,
            s: s * 100,
            l: l * 100
        };
    }
    
    private selectColor(color: string | null): void {
        this.currentColor = color;
        this.tempColor = color;
        this.emit('change', color);
    }
    
    private showAdvancedPicker(): void {
        this.advancedPickerOpen = true;
        this.pickerContainer.style.display = 'none';
        
        if (this.parentElement) {
            const targetRect = this.parentElement.getBoundingClientRect();
            const advancedPickerWidth = 280;
            const advancedPickerHeight = 400;
            
            const left = targetRect.left + (targetRect.width - advancedPickerWidth) / 2;
            const top = targetRect.top + (targetRect.height - advancedPickerHeight) / 2;
            
            this.advancedPickerContainer.style.left = `${left}px`;
            this.advancedPickerContainer.style.top = `${top}px`;
        } else {
            const pickerRect = this.pickerContainer.getBoundingClientRect();
            this.advancedPickerContainer.style.top = `${pickerRect.top}px`;
            this.advancedPickerContainer.style.left = `${pickerRect.left + pickerRect.width + 10}px`;
        }
        
        this.advancedPickerContainer.style.display = 'block';
        
        this.colorPreviewCurrent.style.backgroundColor = this.currentColor || 'transparent';
        this.tempColor = this.currentColor;
        this.colorPreviewNew.style.backgroundColor = this.tempColor || 'transparent';
        
        if (this.currentColor) {
            this.parseColorToRGBA(this.currentColor);
        }
        
        this.updateSpectrumSelector();
    }
    
    private parseColorToRGBA(color: string): void {
        let r = 0, g = 0, b = 0;
        
        const rgbaMatch = color.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
        if (rgbaMatch) {
            r = parseInt(rgbaMatch[1]);
            g = parseInt(rgbaMatch[2]);
            b = parseInt(rgbaMatch[3]);
        } else if (color.startsWith('#') && color.length >= 7) {
            r = parseInt(color.substring(1, 3), 16);
            g = parseInt(color.substring(3, 5), 16);
            b = parseInt(color.substring(5, 7), 16);
        }
        
        const hsl = this.rgbToHsl(r, g, b);
        this.currentHue = Math.round(hsl.h);
        this.currentSaturation = Math.round(hsl.s);
        this.currentLightness = Math.round(hsl.l);
        
        this.rgbaInput.value = `rgba(${r}, ${g}, ${b}, 1)`;
    }
    
    private hideAdvancedPicker(): void {
        this.advancedPickerOpen = false;
        this.advancedPickerContainer.style.display = 'none';
        this.pickerContainer.style.display = 'block';
    }
    
    public show(x: number, y: number): void {
        this.pickerContainer.style.display = 'block';
        this.pickerContainer.style.top = `${y}px`;
        this.pickerContainer.style.left = `${x}px`;
    }
    
    public getColor(): string | null {
        return this.currentColor;
    }
    
    public setColor(color: string | null): void {
        this.currentColor = color;
    }
    
    public getElements() {
        return {
            pickerContainer: this.pickerContainer,
            advancedPickerContainer: this.advancedPickerContainer
        };
    }
    
    public destroy(): void {
        if (this.pickerContainer.parentNode) {
            this.pickerContainer.parentNode.removeChild(this.pickerContainer);
        }
        
        if (this.advancedPickerContainer.parentNode) {
            this.advancedPickerContainer.parentNode.removeChild(this.advancedPickerContainer);
        }
    }
    
    private setupClickOutsideListener(): void {
        document.addEventListener('click', (e) => {
            if (this.advancedPickerOpen && 
                !this.advancedPickerContainer.contains(e.target as Node) && 
                !this.pickerContainer.contains(e.target as Node)) {
                this.hideAdvancedPicker();
            }
        });
    }
}